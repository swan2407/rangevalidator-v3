"""Tolerant C header/source analysis for struct bit-field ranges (V3.1).

What this module does
- Removes comments, joins line splices (backslash-newline), tolerates BOM.
- Preprocessor: collects ``#define NAME <number or simple expression>`` into a macro table
  (used for bit-field widths such as ``: N_BITS``); every other directive is ignored and
  all ``#if`` branches are kept (a conflicting duplicate definition is reported per item).
- Builds a **type dictionary** first: built-in integer types, common fixed-width names
  (uint8_t, uint8, u1, u2, u4, s1, ...), and every scalar ``typedef`` found in the sources,
  resolved through chains (``typedef unsigned char u1; typedef u1 boolean;``).
- Parses struct/union definitions (typedef'd or tagged, nested allowed), bit-fields
  ``type name : width;`` where width is a number, a macro or a simple expression, and
  computes the value range from the resolved base type's signedness.
- Collects global/extern variables of struct types (arrays, pointers, multiple declarators,
  typedef'd pointer types) so ``var.member`` / ``ptr->member`` in a design document can be
  mapped to a struct.
- Diagnostics are per item (one struct member, one declaration). A file is never rejected
  as a whole, and function bodies, enums and prototypes are skipped.
"""
from dataclasses import dataclass, field
import re

ID = re.compile(r"[A-Za-z_][A-Za-z_0-9]*\Z")
KEYWORDS = set("auto break case char const continue default do double else enum extern float for goto if inline int long register restrict return short signed sizeof static struct switch typedef union unsigned void volatile while _Alignas _Alignof _Atomic _Bool _Complex _Generic _Imaginary _Noreturn _Static_assert _Thread_local".split())
QUALIFIERS = {"const", "volatile", "static", "extern", "register", "restrict", "inline", "_Atomic", "__restrict", "__restrict__", "__inline", "__inline__", "near", "far", "__near", "__far", "__packed", "__attribute__"}
_KEYWORD_SPELLINGS = {word.casefold(): word for word in KEYWORDS | QUALIFIERS}


def canonical_keyword(word):
    """Accept case variants of declaration keywords for document comparison."""
    return _KEYWORD_SPELLINGS.get(word.casefold(), word)


def matching_names(names, name):
    """Prefer source spelling, otherwise retain every case-insensitive candidate."""
    if name in names:
        return (name,)
    folded = name.casefold()
    return tuple(item for item in names if item.casefold() == folded)

# Built-in / conventional integer profiles: name -> (signed, bits). Source typedefs override these.
# 32-bit targets (ILP32) are assumed for 'long'; bit-fields are never wider than that in practice.
BASE_PROFILES = {
    "_Bool": (False, 1), "bool": (False, 1),
    "signed char": (True, 8), "unsigned char": (False, 8),
    "short": (True, 16), "signed short": (True, 16), "short int": (True, 16), "signed short int": (True, 16),
    "unsigned short": (False, 16), "unsigned short int": (False, 16),
    "int": (True, 32), "signed": (True, 32), "signed int": (True, 32),
    "unsigned": (False, 32), "unsigned int": (False, 32),
    "long": (True, 32), "signed long": (True, 32), "long int": (True, 32), "signed long int": (True, 32),
    "unsigned long": (False, 32), "unsigned long int": (False, 32),
    "long long": (True, 64), "signed long long": (True, 64), "long long int": (True, 64),
    "unsigned long long": (False, 64), "unsigned long long int": (False, 64),
}
CONVENTIONAL_PROFILES = {
    "uint8_t": (False, 8), "uint16_t": (False, 16), "uint32_t": (False, 32), "uint64_t": (False, 64),
    "int8_t": (True, 8), "int16_t": (True, 16), "int32_t": (True, 32), "int64_t": (True, 64),
    "uint8": (False, 8), "uint16": (False, 16), "uint32": (False, 32), "uint64": (False, 64),
    "sint8": (True, 8), "sint16": (True, 16), "sint32": (True, 32), "sint64": (True, 64),
    "int8": (True, 8), "int16": (True, 16), "int32": (True, 32), "int64": (True, 64),
    "u1": (False, 8), "u2": (False, 16), "u4": (False, 32), "u8": (False, 64),
    "s1": (True, 8), "s2": (True, 16), "s4": (True, 32), "s8": (True, 64),
    "uint": (False, 32), "ulong": (False, 32), "ushort": (False, 16), "uchar": (False, 8),
    "boolean": (False, 8), "BOOL": (True, 32), "BYTE": (False, 8), "WORD": (False, 16), "DWORD": (False, 32),
    "U8": (False, 8), "U16": (False, 16), "U32": (False, 32), "S8": (True, 8), "S16": (True, 16), "S32": (True, 32),
    "UINT8": (False, 8), "UINT16": (False, 16), "UINT32": (False, 32), "INT8": (True, 8), "INT16": (True, 16), "INT32": (True, 32),
}
# Plain 'char' bit-field signedness is implementation-defined: reported, not certified.
IMPLEMENTATION_DEFINED = {"char"}
# Kept for compatibility with older callers/tests.
BUILTINS = set(BASE_PROFILES) | {"char"}
FIXED_INTEGER_PROFILES = {k: CONVENTIONAL_PROFILES[k] for k in ("u1", "u2", "s1", "uint")} | {
    "unsigned char": (False, 8), "unsigned short": (False, 16), "signed char": (True, 8),
    "unsigned": (False, 32), "unsigned int": (False, 32), "int": (True, 32), "signed": (True, 32), "signed int": (True, 32)}


def identifier(s):
    return bool(ID.fullmatch(s)) and s not in KEYWORDS


@dataclass(frozen=True)
class Diagnostic:
    file: str
    line: int
    reason: str
    code: str = "UNSUPPORTED_SYNTAX"
    status: str = "REVIEW"


@dataclass(frozen=True)
class BitField:
    struct_name: str
    struct_tag: str
    member: str
    declared_type: str
    width: int
    file: str
    line: int
    signed: bool | None = None
    minimum: int | None = None
    maximum: int | None = None
    note: str = ""          # why the range could not be certified (empty when certified)


@dataclass
class Struct:
    name: str               # typedef name, or the tag when there is no typedef
    tag: str
    typedef: bool
    file: str
    line: int
    fields: list = field(default_factory=list)          # BitField members (in order)
    ordinary: list = field(default_factory=list)        # non-bit-field member names
    members: dict = field(default_factory=dict)         # member name -> declared type text (all members)
    nested: dict = field(default_factory=dict)          # member name -> Struct (struct/union typed member)
    aliases: list = field(default_factory=list)         # extra typedef names for the same struct
    conflict: str = ""                                  # set when another definition with the same identity differs


@dataclass(frozen=True)
class Variable:
    name: str
    type_index: int             # index into SourceModel.types, or -1 when the type is defined in another file
    pointer: bool
    file: str
    line: int
    array: bool = False
    extent: str = ""
    type_text: str = ""         # declared type as written ('Device_t', 'struct DeviceTag', 'Device_p')


@dataclass
class SourceModel:
    fields: list = field(default_factory=list)
    types: list = field(default_factory=list)
    variables: list = field(default_factory=list)
    reviews: list = field(default_factory=list)
    typedefs: dict = field(default_factory=dict)        # alias -> declared type text (scalar typedefs)
    macros: dict = field(default_factory=dict)          # NAME -> int (numeric #define / enumerators)
    enums: set = field(default_factory=set)             # enum type names (typedef or tag)


@dataclass(frozen=True)
class Match:
    status: str
    reason: str
    field: BitField | None = None
    declaration: Variable | None = None


@dataclass(frozen=True)
class Token:
    text: str
    line: int


# ------------------------------------------------------------------ lexing
_DEFINE = re.compile(r"^[ \t]*#[ \t]*define[ \t]+([A-Za-z_]\w*)(?!\()[ \t]*(.*?)[ \t]*$", re.M)
_DIRECTIVE = re.compile(r"^[ \t]*#[^\n]*", re.M)
_CMT_STR = re.compile(r'//[^\n]*|/\*.*?\*/|"(?:\\.|[^"\\\n])*"|\'(?:\\.|[^\'\\\n])*\'', re.S)


def preprocess(source):
    """-> (code without comments/directives, {macro: int})"""
    text = source.replace("\r\n", "\n").replace("\r", "\n").lstrip("﻿")
    text = text.replace("\\\n", "")                                     # line splices
    text = _CMT_STR.sub(lambda m: m.group(0) if m.group(0)[0] in "\"'" else " " + "\n" * m.group(0).count("\n"), text)
    macros = {}
    for name, body in _DEFINE.findall(text):
        v = eval_int(body, macros)
        if v is not None:
            macros[name] = v
    code = _DIRECTIVE.sub(lambda m: " " * len(m.group(0)), text)        # keep line numbers
    return code, macros


_EXPR_OK = re.compile(r"^[\sA-Za-z_0-9()+\-*/%<>|&^~]*$")


def eval_int(expr, macros, depth=0):
    """Integer value of a constant expression (numbers, macros, + - * / % << >> & | ^ ~ parentheses). None if unknown."""
    s = str(expr).strip()
    if not s or depth > 8 or not _EXPR_OK.match(s) or "//" in s or "**" in s:
        return None
    s = re.sub(r"\b(0[xX][0-9A-Fa-f]+|0[bB][01]+|\d+)[uUlL]+\b", r"\1", s)        # 10U, 0x1FUL
    s = re.sub(r"\b0[bB]([01]+)\b", lambda m: str(int(m.group(1), 2)), s)
    s = re.sub(r"\b0([0-7]+)\b", lambda m: str(int(m.group(1), 8)), s)              # octal (rare)

    def sub(m):
        w = m.group(0)
        if w in macros:
            return f"({macros[w]})"
        return w
    s2 = re.sub(r"\b[A-Za-z_]\w*\b", sub, s)
    if re.search(r"[A-Za-z_]", s2.replace("0x", "").replace("0X", "")):
        return None                                                      # unresolved identifier
    s2 = s2.replace("/", "//")
    try:
        v = eval(s2, {"__builtins__": {}}, {})
    except Exception:
        return None
    return v if isinstance(v, int) and not isinstance(v, bool) else None


def tokenize(source, file, reviews):
    """Tokens with line numbers. Literals collapse to '<literal>'. Never rejects a file."""
    out, i, line = [], 0, 1
    n = len(source)
    while i < n:
        c = source[i]
        if c.isspace():
            line += c == "\n"
            i += 1
        elif c in "\"'":
            quote, start = c, line
            i += 1
            while i < n:
                q = source[i]
                line += q == "\n"
                i += 1
                if q == "\\" and i < n:
                    line += source[i] == "\n"
                    i += 1
                elif q == quote or q == "\n":
                    break
            out.append(Token("<literal>", start))
        elif c.isascii() and (c.isalnum() or c == "_"):
            start = i
            i += 1
            while i < n and source[i].isascii() and (source[i].isalnum() or source[i] == "_"):
                i += 1
            out.append(Token(source[start:i], line))
        elif source.startswith(("->", "<<", ">>", "::"), i):
            out.append(Token(source[i:i + 2], line))
            i += 2
        else:
            out.append(Token(c, line))
            i += 1
    return out


# ------------------------------------------------------------------ parsing
class _Parser:
    collect_bitfields = False

    def __init__(self, tokens, file, model):
        self.t = [Token(canonical_keyword(token.text), token.line) for token in tokens]
        self.file = file
        self.m = model

    def review(self, line, reason, code="UNSUPPORTED_SYNTAX"):
        self.m.reviews.append(Diagnostic(self.file, line, reason, code))

    # -- helpers
    @staticmethod
    def split(tokens, sep=","):
        """Split a token list on sep at depth 0 (parentheses/brackets/braces)."""
        parts, depth, cur = [], 0, []
        for tok in tokens:
            if tok.text in "([{":
                depth += 1
            elif tok.text in ")]}":
                depth -= 1
            if tok.text == sep and depth == 0:
                parts.append(cur)
                cur = []
            else:
                cur.append(tok)
        if cur:
            parts.append(cur)
        return parts

    @staticmethod
    def matching(tokens, start):
        """Index of the brace matching tokens[start] ('{')."""
        depth = 0
        for i in range(start, len(tokens)):
            if tokens[i].text == "{":
                depth += 1
            elif tokens[i].text == "}":
                depth -= 1
                if depth == 0:
                    return i
        return -1

    def declarations(self, on_function=None):
        """Top-level declarations as token lists.
        A declaration ends at ';' (brace depth 0). A '{' at depth 0 that is not part of a
        struct/union/enum body is either a function body (declaration had '(') → skipped, or an
        initializer → skipped over. extern "C" { ... } braces are flattened."""
        t = self.t
        out, i, n = [], 0, len(t)
        while i < n:
            x = t[i].text
            if x in (";", "}"):
                i += 1
                continue
            if x == "extern" and i + 2 < n and t[i + 1].text == "<literal>" and t[i + 2].text == "{":
                close = self.matching(t, i + 2)
                if close < 0:
                    break
                t = t[:i] + t[i + 3:close] + t[close + 1:]
                n = len(t)
                continue
            j, depth, aggregate, paren, function = i, 0, False, False, False
            while j < n:
                x = t[j].text
                if depth == 0:
                    if x in ("struct", "union", "enum"):
                        aggregate = True
                    elif x == "(":
                        paren = True
                    elif x == "{" and (not aggregate or (
                            self.collect_bitfields and paren
                            and not any(tok.text == "=" for tok in t[i:j]))):
                        close = self.matching(t, j)
                        if close < 0:
                            j = n
                            break
                        if paren:                                         # function definition
                            function = True
                            if on_function is not None:
                                on_function(t[i:j], t[j + 1:close])
                            j = close
                            break
                        j = close + 1                                     # initializer block
                        continue
                    elif x == ";":
                        break
                if x == "{":
                    depth += 1
                elif x == "}":
                    depth -= 1
                j += 1
            if not function and j > i:
                out.append(t[i:j])
            i = j + 1
        self.t = t
        return out

    # -- aggregates
    def parse_aggregate(self, tokens, pos, typedef, line):
        """tokens[pos] is 'struct'/'union'. Parse '[tag] { body }' → (Struct, index after '}') or (None, pos)."""
        kind = tokens[pos].text
        p = pos + 1
        while p < len(tokens) and tokens[p].text in ("__attribute__", "__packed") or (p < len(tokens) and tokens[p].text == "(" and tokens[p - 1].text == "__attribute__"):
            if tokens[p].text == "__attribute__":
                p += 1
                if p < len(tokens) and tokens[p].text == "(":
                    depth = 0
                    while p < len(tokens):
                        if tokens[p].text == "(":
                            depth += 1
                        elif tokens[p].text == ")":
                            depth -= 1
                            if depth == 0:
                                p += 1
                                break
                        p += 1
            else:
                p += 1
        tag = ""
        if p < len(tokens) and identifier(tokens[p].text):
            tag, p = tokens[p].text, p + 1
        if p >= len(tokens) or tokens[p].text != "{":
            return None, p, tag                                           # 'struct tag' without body
        close = self.matching(tokens, p)
        if close < 0:
            self.review(line, f"{kind} body not closed")
            return None, len(tokens), tag
        st = Struct(tag or f"<{kind}@{tokens[pos].line}>", tag, typedef, self.file, tokens[pos].line)
        self.parse_members(st, tokens[p + 1:close], kind)
        return st, close + 1, tag

    def parse_members(self, st, body, kind):
        seen = set()
        for decl in self.split(body, ";"):
            if not decl:
                continue
            words = [x.text for x in decl]
            line = decl[0].line
            # nested struct/union member
            agg_pos = next((k for k, w in enumerate(words) if w in ("struct", "union")), None)
            if agg_pos is not None and "{" in words:
                inner, after, _tag = self.parse_aggregate(decl, agg_pos, False, line)
                if inner is None:
                    self.review(line, "nested aggregate not parsed")
                    continue
                self.register_struct(inner)
                declarators = self.split(decl[after:], ",")
                if not any(d for d in declarators):                          # anonymous nested → flatten
                    for f in inner.fields:
                        st.fields.append(f)
                        self.m.fields.append(f)
                        st.members[f.member] = f.declared_type
                    st.ordinary.extend(inner.ordinary)
                    st.members.update(inner.members)
                    st.nested.update(inner.nested)
                    continue
                for d in declarators:
                    names = [x.text for x in d if identifier(x.text)]
                    if names:
                        st.nested[names[-1]] = inner
                        st.ordinary.append(names[-1])
                        st.members[names[-1]] = (inner.name or kind)
                continue
            if "(" in words and not (self.collect_bitfields and ":" in words):  # function pointer member
                names = [w for w in words if identifier(w)]
                if names:
                    st.ordinary.append(names[-1])
                    st.members[names[-1]] = "function pointer"
                continue
            # specifiers end before the first declarator: find first identifier that is followed by ':' ',' '[' or end
            # declarators are separated by ','; each: [*]* name [ '[' N ']' ] [ ':' width ]   or  ':' width (padding)
            parts = self.split(decl, ",")
            if not parts:
                continue
            first = parts[0]
            spec_end = len(first)
            colon_search = range(len(first)) if self.collect_bitfields else range(len(first) - 1, -1, -1)
            for k in colon_search:
                if first[k].text == ":":
                    spec_end = k
                    break
            # specifier words = everything before the declarator name of the first part
            fw = [x.text for x in first[:spec_end]]
            # declarator name = last identifier before ':' / '[' / end
            name_idx = None
            for k in range(spec_end - 1, -1, -1):
                if identifier(first[k].text):
                    name_idx = k
                    break
                if first[k].text in ("]", "[") or first[k].text.isdigit():
                    continue
            if name_idx is None:
                # unnamed bit-field padding: ': 3;'
                continue
            spec_words = [w for w in fw[:name_idx] if w not in QUALIFIERS and w != "*"]
            if not spec_words and ":" in [x.text for x in first]:
                continue                                                   # 'uint : 4;' → unnamed padding (uint is the type)
            declared = " ".join(spec_words)
            declarators = [first[name_idx:]] + parts[1:]
            for d in declarators:
                dw = [x.text for x in d]
                if not dw:
                    continue
                pointer = "*" in dw[:1] or (dw and dw[0] == "*")
                ids = [k for k, w in enumerate(dw) if identifier(w)]
                if ":" in dw:
                    c = dw.index(":")
                    if not ids or ids[0] > c:                              # ': width' padding
                        continue
                    member = dw[ids[0]]
                    width_expr = " ".join(dw[c + 1:])
                    at = d[0].line
                    if member in seen:
                        self.review(at, f"duplicate member '{member}'")
                    seen.add(member)
                    st.members[member] = declared
                    if self.collect_bitfields and ("*" in fw or "[" in fw or "*" in dw[:c] or "[" in dw[:c]):
                        declared_type = "<unsupported bit-field declarator>"
                    else:
                        declared_type = declared
                    f = self.parse_bitfield(st, member, declared_type, width_expr, at)
                    st.fields.append(f)
                    self.m.fields.append(f)
                else:
                    if not ids:
                        continue
                    member = dw[ids[0]]
                    if member in seen:
                        self.review(d[0].line, f"duplicate member '{member}'")
                    seen.add(member)
                    st.ordinary.append(member)
                    st.members[member] = declared + (" *" if pointer else "") + ("[]" if "[" in dw else "")

    def parse_bitfield(self, st, member, declared, width_expr, at):
        width = eval_int(width_expr, self.m.macros)
        if width is None or width <= 0:
            self.review(at, f"bit-field width '{width_expr}' unresolved for {member}", "WIDTH_UNRESOLVED")
            return BitField(st.name, st.tag, member, declared, 0, self.file, at, None, None, None,
                            f"비트폭 '{width_expr}' 을 계산하지 못함 (매크로 미정의)")
        return self.bitfield(st, member, declared, width, at)

    def bitfield(self, st, member, declared, width, line):
        signed, bits, note = self.m_resolve(declared)
        if signed is None:
            return BitField(st.name, st.tag, member, declared, width, self.file, line, None, None, None, note)
        if bits and width > bits:
            self.review(line, f"bit-field width {width} exceeds base type width {bits} ({declared})")
            return BitField(st.name, st.tag, member, declared, width, self.file, line, None, None, None,
                            f"비트폭 {width} 이 기본형 {declared}({bits}비트)보다 큼")
        if signed:
            lo, hi = -(1 << (width - 1)), (1 << (width - 1)) - 1
        else:
            lo, hi = 0, (1 << width) - 1
        return BitField(st.name, st.tag, member, declared, width, self.file, line, signed, lo, hi, "")

    def m_resolve(self, declared):
        return resolve_type(declared, self.m.typedefs, self.m.enums)

    def rename(self, st, name):
        """Give the struct its typedef name (known only after the closing brace) and relabel its bit-fields."""
        from dataclasses import replace as _replace
        st.name = name
        renamed = {id(f): _replace(f, struct_name=name) for f in st.fields}
        if renamed:
            self.m.fields = [renamed.get(id(f), f) for f in self.m.fields]
        st.fields = [renamed[id(f)] for f in st.fields]

    def register_struct(self, st):
        self.m.types.append(st)
        return len(self.m.types) - 1

    # -- top-level
    def run(self):
        decls = self.declarations()
        # pass 1: scalar typedefs and enums (type dictionary), so later bit-fields resolve regardless of order
        for d in decls:
            w = [x.text for x in d]
            if not w:
                continue
            if w[0] == "typedef" and "{" not in w and "(" not in w and not any(x in ("struct", "union") for x in w):
                if w[1:2] == ["enum"]:
                    if len(w) >= 3:
                        self.m.enums.add(w[-1])
                    continue
                if len(w) >= 3:
                    spec = " ".join(x for x in w[1:-1] if x not in QUALIFIERS)
                    alias = w[-1]
                    if identifier(alias) and spec and "*" not in spec and "[" not in spec:
                        if alias in self.m.typedefs and self.m.typedefs[alias] != spec:
                            self.review(d[0].line, f"typedef '{alias}' redefined differently ({self.m.typedefs[alias]} / {spec})", "TYPEDEF_CONFLICT")
                        self.m.typedefs.setdefault(alias, spec)
            elif w[0] == "enum" or (w[0] == "typedef" and w[1:2] == ["enum"]):
                self.parse_enum(d)
        # pass 2: aggregates and variables
        for d in decls:
            w = [x.text for x in d]
            if not w:
                continue
            try:
                self.declaration(d, w)
            except Exception as exc:                                       # never abort a file
                self.review(d[0].line, f"declaration skipped: {type(exc).__name__}")
        # pass 3: variables whose struct type appeared later in this file
        resolved = []
        for v in self.m.variables:
            if v.type_index < 0:
                hit = self.struct_for(v.type_text)
                if hit is not None:
                    v = Variable(v.name, hit[0], v.pointer or hit[1], v.file, v.line, v.array, v.extent, v.type_text)
            resolved.append(v)
        self.m.variables = resolved
        return self.m

    def parse_enum(self, d):
        w = [x.text for x in d]
        if "{" not in w:
            return
        open_, close = w.index("{"), self.matching(d, w.index("{"))
        if close < 0:
            return
        tag = w[open_ - 1] if open_ >= 1 and identifier(w[open_ - 1]) else ""
        if tag:
            self.m.enums.add(tag)
        for name in [x.text for x in d[close + 1:] if identifier(x.text)]:
            self.m.enums.add(name)
        value = -1
        for item in self.split(d[open_ + 1:close], ","):
            iw = [x.text for x in item]
            if not iw:
                continue
            if "=" in iw:
                v = eval_int(" ".join(iw[iw.index("=") + 1:]), self.m.macros)
                value = v if v is not None else value + 1
            else:
                value += 1
            if identifier(iw[0]):
                self.m.macros.setdefault(iw[0], value)

    def declaration(self, d, w):
        td = w[0] == "typedef"
        if td and w[1:2] == ["enum"]:
            return                                                         # handled in pass 1
        agg_pos = next((k for k, x in enumerate(w) if x in ("struct", "union")), None)
        if agg_pos is not None and "{" in w and w.index("{") > agg_pos and "(" not in w[:w.index("{")]:
            st, after, tag = self.parse_aggregate(d, agg_pos, td, d[0].line)
            if st is None:
                return
            tail = self.split(d[after:], ",")
            names = [[x.text for x in part] for part in tail]
            if td:
                plain = [n for n in names if n and "*" not in n and "[" not in n and identifier(n[-1])]
                if plain:
                    self.rename(st, plain[0][-1])
                    st.aliases = [n[-1] for n in plain[1:]]
                elif not tag:
                    self.rename(st, f"<anonymous@{st.line}>")
                idx = self.register_struct(st)
                for n in names:                                            # typedef struct {...} T, *PT;
                    if n and "*" in n and identifier(n[-1]):
                        self.m.typedefs[n[-1]] = "<pointer-to:%d>" % idx
            else:
                if not tag:
                    self.rename(st, f"<anonymous@{st.line}>")
                idx = self.register_struct(st)
                for n in names:                                            # struct X {...} var, *p, arr[3];
                    if n and identifier(n[-1]):
                        self.m.variables.append(Variable(n[-1], idx, "*" in n, self.file, d[0].line, "[" in n, _extent(n)))
            return
        if td:
            # typedef struct tag alias;  /  typedef struct tag *alias_p;  (no body)
            if agg_pos is not None:
                tag = w[agg_pos + 1] if agg_pos + 1 < len(w) and identifier(w[agg_pos + 1]) else ""
                alias = w[-1]
                if tag and identifier(alias):
                    self.m.typedefs[alias] = ("<pointer-to-tag:%s>" if "*" in w else "<tag:%s>") % tag
            return                                                         # scalar typedefs handled in pass 1
        if "(" in w:
            return                                                         # prototype / function pointer
        # variable declaration: specifiers + declarators
        parts = self.split(d, ",")
        first = parts[0]
        fw = [x.text for x in first]
        if "=" in fw:
            fw_head = fw[:fw.index("=")]
        else:
            fw_head = fw
        name_idx = None
        for k in range(len(fw_head) - 1, -1, -1):
            if identifier(fw_head[k]):
                name_idx = k
                break
            if fw_head[k] in ("]", "[") or fw_head[k].isdigit():
                continue
        if name_idx is None:
            return
        spec_words = [x for x in fw[:name_idx] if x not in QUALIFIERS and x != "*"]
        pointer_first = "*" in fw[:name_idx]
        type_text = " ".join(spec_words)
        if not type_text or type_text in BASE_PROFILES or type_text in CONVENTIONAL_PROFILES:
            return                                                         # ordinary integers are not our concern
        if type_text in self.m.typedefs and not self.m.typedefs[type_text].startswith("<"):
            return                                                         # scalar typedef (u1, boolean, ...)
        hit = self.struct_for(type_text)
        idx, via_pointer_alias = hit if hit is not None else (-1, False)
        declarators = [fw[name_idx:]] + [[x.text for x in p] for p in parts[1:]]
        for k, dw in enumerate(declarators):
            if "=" in dw:
                dw = dw[:dw.index("=")]
            ids = [x for x in dw if identifier(x)]
            if not ids:
                continue
            pointer = "*" in dw or via_pointer_alias or (k == 0 and pointer_first)
            self.m.variables.append(Variable(ids[0], idx, pointer, self.file, d[0].line, "[" in dw, _extent(dw), type_text))

    def struct_for(self, type_text):
        """Struct index for a declaration type text ('Device_t', 'struct Tag', pointer typedef) → (index, is_pointer_alias) or None."""
        return struct_index(type_text, self.m.types, self.m.typedefs)


def struct_index(type_text, types, typedefs):
    """Resolve a type spelling to (index into types, pointer_alias) using tags, typedef names, aliases and
    pointer typedefs. None when the spelling is not a known struct in `types`."""
    words = str(type_text).split()
    if not words:
        return None
    if words[0] in ("struct", "union") and len(words) >= 2:
        tag = words[1]
        idx = next((i for i, s in enumerate(types) if s.tag == tag), None)
        return (idx, False) if idx is not None else None
    name, seen, pointer = words[-1], set(), False
    while name and name not in seen:
        seen.add(name)
        idx = next((i for i, s in enumerate(types) if s.name == name or name in s.aliases), None)
        if idx is not None:
            return idx, pointer
        spec = typedefs.get(name)
        if spec is None:
            return None
        if spec.startswith("<pointer-to:"):
            return int(spec[len("<pointer-to:"):-1]), True
        if spec.startswith("<pointer-to-tag:") or spec.startswith("<tag:"):
            tag = spec.split(":", 1)[1][:-1]
            idx = next((i for i, s in enumerate(types) if s.tag == tag), None)
            return (idx, spec.startswith("<pointer")) if idx is not None else None
        name = spec.split()[-1]
    return None


def _extent(words):
    if "[" in words and "]" in words:
        return " ".join(words[words.index("[") + 1:words.index("]")])
    return ""


def resolve_type(declared, typedefs, enums=(), depth=0):
    """Declared type text → (signed, bits, note). (None, 0, note) when the range cannot be certified."""
    words = [w for w in str(declared).split() if w not in QUALIFIERS]
    text = " ".join(words)
    if not text:
        return None, 0, "타입 없음"
    if text in BASE_PROFILES:
        s, b = BASE_PROFILES[text]
        return s, b, ""
    if text in IMPLEMENTATION_DEFINED:
        return None, 8, "plain char 비트필드: 부호가 컴파일러 정의"
    if len(words) == 1:
        name = words[0]
        if name in typedefs and depth < 16:
            spec = typedefs[name]
            if spec.startswith("<"):
                return None, 0, f"'{name}' 은 구조체/포인터 typedef"
            return resolve_type(spec, typedefs, enums, depth + 1)
        if name in enums or (words[0] == "enum"):
            return None, 32, "enum 타입 비트필드: 부호가 컴파일러 정의 (GCC 무부호 / MSVC 부호)"
        if name in CONVENTIONAL_PROFILES:
            s, b = CONVENTIONAL_PROFILES[name]
            return s, b, ""
        return None, 0, f"알 수 없는 타입 '{name}' (typedef 를 찾지 못함)"
    if words[0] == "enum":
        return None, 32, "enum 타입 비트필드: 부호가 컴파일러 정의"
    # e.g. 'unsigned u1'?? treat trailing alias with explicit sign
    sign = None
    rest = []
    for w in words:
        if w == "unsigned":
            sign = False
        elif w == "signed":
            sign = True
        else:
            rest.append(w)
    if rest:
        s, b, note = resolve_type(" ".join(rest), typedefs, enums, depth + 1)
        if b:
            return (sign if sign is not None else s), b, note
    return None, 0, f"알 수 없는 타입 '{text}'"


def analyze(source: str, file: str = "<memory>") -> SourceModel:
    model = SourceModel()
    code, macros = preprocess(source)
    model.macros.update(macros)
    tokens = tokenize(code, file, model.reviews)
    parser = _Parser(tokens, file, model)
    try:
        parser.run()
    except Exception as exc:
        model.reviews.append(Diagnostic(file, 1, f"analysis aborted: {type(exc).__name__}: {exc}", "ANALYSIS_ERROR"))
    return model


def match_expression(model: SourceModel, expression: str) -> Match:
    """Single-file helper kept for the CLI/tests: var.member / ptr->member → bit-field."""
    m = re.fullmatch(r"\s*([A-Za-z_]\w*)\s*(?:\[[^\]]*\]\s*)?(->|\.)\s*([A-Za-z_]\w*)\s*", expression, flags=re.ASCII)
    if not m or not identifier(m[1]) or not identifier(m[3]):
        return Match("REVIEW", "Unsupported member expression; requires obj.field or ptr->field")
    bases = [v for v in model.variables if v.name == m[1]]
    if len({b.type_index for b in bases}) != 1:
        types = [i for i, t in enumerate(model.types) if t.name == m[1] + "_t" or t.name == m[1]]
        if len(types) != 1:
            return Match("REVIEW", "Unresolved/ambiguous base variable")
        st = model.types[types[0]]
        v = None
    else:
        v = bases[0]
        st = model.types[v.type_index]
    fields = [f for f in st.fields if f.member == m[3]]
    if not fields and m[3] in st.ordinary:
        return Match("SKIP", "Ordinary integer member excluded")
    if len(fields) != 1 or fields[0].minimum is None:
        return Match("REVIEW", "No unambiguous supported bit-field range" + (": " + fields[0].note if fields and fields[0].note else ""))
    return Match("TYPE_MATCH", "Struct/member/width/signedness confirmed; document not compared", fields[0], v)
