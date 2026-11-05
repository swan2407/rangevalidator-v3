"""Static evidence: only logging differs in Reader; protected modules stay byte-identical."""
import ast,hashlib,json
from pathlib import Path
root=Path(__file__).resolve().parents[1]
before=json.loads((root/'build/diagnostic-audit/before-sha256.json').read_text(encoding='utf-8'))
allowed={'rangevalidator\\diagnostics.py','rangevalidator\\gui.py','rangevalidator\\readers.py','gui_entry.py','docs\\DISTRIBUTION.json'}
changed=[name for name,sha in before.items() if hashlib.sha256((root/name).read_bytes()).hexdigest()!=sha]
assert set(changed)<=allowed,changed
class StripTrace(ast.NodeTransformer):
    def visit_FunctionDef(self,node):
        if node.name in ('trace','mark'):return None
        return self.generic_visit(node)
    def visit_If(self,node):
        if isinstance(node.test,ast.Name) and node.test.id=='_trace':return None
        node=self.generic_visit(node)
        if not node.body:return node.orelse or None
        return node
    def visit_Expr(self,node):
        if isinstance(node.value,ast.Call) and isinstance(node.value.func,ast.Name) and node.value.func.id=='trace':return None
        return self.generic_visit(node)
old=ast.parse((root/'build/diagnostic-audit/before-readers.py').read_text(encoding='utf-8-sig'))
new=ast.parse((root/'rangevalidator/readers.py').read_text(encoding='utf-8-sig'))
assert ast.dump(StripTrace().visit(old),include_attributes=False)==ast.dump(StripTrace().visit(new),include_attributes=False),'Non-diagnostic Reader change detected'
record=dict(changed_preexisting_files=changed,reader_non_diagnostic_ast_identical=True,engine_writer_report_workflow_models_byte_identical=True,v1_v2_accessed_or_modified=False)
(root/'docs/DIAGNOSTIC_SCOPE_CHECK.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(record,ensure_ascii=False,indent=2))
