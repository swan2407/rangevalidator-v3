import unittest
from rangevalidator.analysis import analyze, match_expression

BASIC = 'typedef struct { unsigned int active:1; signed int offset:8; int normal; } A; A *p; A obj;'


class AnalysisTests(unittest.TestCase):
    def test_metadata_comments_multiline(self):
        m = analyze('// struct Fake { int f:1; };\ntypedef struct Tag {\n unsigned int flag:1;\n signed int state:3;\n} A;\nA *p; /* fake */', 'fixture.h')
        self.assertFalse(m.reviews)
        self.assertEqual([(f.struct_name, f.struct_tag, f.member, f.line, f.file) for f in m.fields], [('A', 'Tag', 'flag', 3, 'fixture.h'), ('A', 'Tag', 'state', 4, 'fixture.h')])

    def test_boundaries(self):
        m = analyze('struct E { unsigned int u:32; signed int s:32; signed int one:1; };')
        self.assertFalse(m.reviews)
        self.assertEqual([(x.minimum, x.maximum, x.signed) for x in m.fields], [(0, 4294967295, False), (-2147483648, 2147483647, True), (-1, 0, True)])

    def test_uint(self):
        m = analyze('typedef unsigned int uint; struct E { uint a:12; };')
        self.assertFalse(m.reviews)
        self.assertEqual(m.fields[0].maximum, 4095)

    def test_type_identity(self):
        m = analyze('typedef struct { unsigned int active:1; } A; typedef struct { unsigned int active:3; } B; A *a; B *b;')
        self.assertEqual(match_expression(m, 'a->active').field.maximum, 1)
        self.assertEqual(match_expression(m, 'b->active').field.maximum, 7)
        self.assertEqual(match_expression(m, 'unknown->active').status, 'REVIEW')

    def test_pointer_object(self):
        m = analyze(BASIC)
        for expression in ('p->active', 'obj.active', ' p -> active '):
            with self.subTest(expression=expression):
                self.assertEqual(match_expression(m, expression).status, 'TYPE_MATCH')
        for expression in ('p->x', 'active', '(*p).active', 'p->nested.active', 'f()', 'M(x)', 'return.active'):
            with self.subTest(expression=expression):
                self.assertEqual(match_expression(m, expression).status, 'REVIEW')

    def test_ordinary_member_excluded(self):
        self.assertEqual(match_expression(analyze(BASIC), 'obj.normal').status, 'SKIP')

    def test_extern_tag(self):
        m = analyze('struct Tag { unsigned int a:2; }; extern struct Tag *p;')
        self.assertFalse(m.reviews)
        self.assertEqual(match_expression(m, 'p->a').field.maximum, 3)

    def test_include_guard(self):
        self.assertFalse(analyze('#ifndef A_H\n#define A_H\n' + BASIC + '\n#endif').reviews)

    def test_lexical_literals(self):
        m = analyze('const char *s="struct Fake { unsigned int x:1; };"; struct A { unsigned int x:2; };')
        self.assertEqual(len(m.fields), 1)
        self.assertEqual(m.fields[0].struct_name, 'A')

    def test_width_macro_and_expression(self):
        m = analyze('#define W 3\n#define W2 (W+1)\nstruct A { unsigned int a:W; unsigned int b:W2; unsigned int c:1<<2; }; struct A *p;')
        self.assertEqual([(f.member, f.width, f.maximum) for f in m.fields], [('a', 3, 7), ('b', 4, 15), ('c', 4, 15)])
        self.assertEqual(match_expression(m, 'p->b').status, 'TYPE_MATCH')

    def test_unresolved_width_is_per_member(self):
        m = analyze('struct A { unsigned int bad:WIDTH; unsigned int good:2; }; struct A *p;')
        self.assertEqual([f.member for f in m.fields], ['bad', 'good'])
        self.assertIsNone(m.fields[0].minimum)
        self.assertEqual(match_expression(m, 'p->good').status, 'TYPE_MATCH')
        self.assertEqual(match_expression(m, 'p->bad').status, 'REVIEW')

    def test_typedef_dictionary_chain(self):
        m = analyze('typedef unsigned char u1; typedef u1 boolean; typedef enum {E0, E1} e_t;\n'
                    'struct A { boolean f:1; u1 n:3; uint8_t s:2; char c:2; e_t e:2; }; struct A obj;')
        by = {f.member: f for f in m.fields}
        self.assertEqual((by['f'].maximum, by['n'].maximum, by['s'].maximum), (1, 7, 3))
        self.assertIsNone(by['c'].minimum)          # plain char: implementation defined
        self.assertIsNone(by['e'].minimum)          # enum typed: implementation defined
        self.assertEqual(match_expression(m, 'obj.n').status, 'TYPE_MATCH')

    def test_preprocessor_and_functions_tolerated(self):
        m = analyze('#ifndef X_H\n#define X_H\n#include "a.h"\n#if MODE\nstruct A { unsigned int a:1; };\n#endif\n'
                    'static inline int helper(int v) { return v + 1; }\nextern struct A obj;\n#endif')
        self.assertEqual([f.member for f in m.fields], ['a'])
        self.assertEqual(match_expression(m, 'obj.a').field.maximum, 1)

    def test_nested_and_multiple_declarators(self):
        m = analyze('typedef struct { unsigned int a:1, b:2; struct { unsigned int c:3; } sub; union { unsigned int d:1; int e; }; } A, *A_p; A obj; A_p ap;')
        t = m.types[-1]
        self.assertEqual([f.member for f in t.fields], ['a', 'b', 'd'])
        self.assertIn('sub', t.nested)
        self.assertEqual([(v.name, v.pointer) for v in m.variables], [('obj', False), ('ap', True)])

    def test_exceeding_width_and_duplicates_are_reviews(self):
        m = analyze('struct A { unsigned int x:33; unsigned int y:1; unsigned int y:2; };')
        self.assertIsNone(m.fields[0].minimum)
        self.assertTrue(any('exceeds' in r.reason for r in m.reviews))
        self.assertTrue(any('duplicate' in r.reason for r in m.reviews))

    def test_local_excluded(self):
        m = analyze('void f() { struct A { unsigned int x:1; }; }')
        self.assertFalse(m.fields)


NEGATIVE = [
    'struct A { unsigned int x:0; };',
    'struct A { unsigned int x:18446744073709551616; };',
    'struct A { unsigned int x:33; };', 'struct A { unsigned int x:64; };',
    'struct A { char x:3; }; struct A *p;',
    'typedef struct { unsigned int active:1; unsigned int active:2; } A; A *p;',
    'A *p; typedef struct { unsigned int other:1; } A;',
]


def make_negative(source):
    def test(self):
        m = analyze(source, 'negative.h')
        self.assertEqual(match_expression(m, 'p->active').status, 'REVIEW')
    return test


for i, source in enumerate(NEGATIVE):
    setattr(AnalysisTests, f'test_negative_{i:02}', make_negative(source))
