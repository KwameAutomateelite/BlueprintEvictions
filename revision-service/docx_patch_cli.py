import sys,json,base64,hashlib,io,zipfile
from lxml import etree
from docx_patch import patch_docx,W
from rent_table import inspect_rent_tables
from standard_form import is_standard, inspect_standard, printed_signer

def handle(p):
 raw=base64.b64decode(p['docx_base64'],validate=True)
 if not raw or len(raw)>10*1024*1024:raise ValueError('Document size requires review')
 digest=hashlib.sha256(raw).hexdigest()
 if p.get('expected_sha256')!=digest:raise ValueError('Saved document hash mismatch')
 z=zipfile.ZipFile(io.BytesIO(raw))
 if sum(i.file_size for i in z.infolist())>40*1024*1024:raise ValueError('Expanded document too large')
 root=etree.fromstring(z.read('word/document.xml'),etree.XMLParser(resolve_entities=False,no_network=True))
 text='\n'.join(''.join(t.text or '' for t in para.iter(W+'t')) for para in root.iter(W+'p'))
 if p['mode']=='inspect':return {'source_sha256':digest,'text':text,'rent_tables':inspect_rent_tables(root),**({'form_fields':inspect_standard(root)['form_fields']} if is_standard(root) else {})}
 if p['mode']!='patch' or not isinstance(p.get('changes'),list) or not 1<=len(p['changes'])<=30:raise ValueError('Explicit changes required')
 out=patch_docx(raw,p['changes'])
 final=etree.fromstring(zipfile.ZipFile(io.BytesIO(out)).read('word/document.xml'))
 return {'source_sha256':digest,'docx_sha256':hashlib.sha256(out).hexdigest(),'docx_base64':base64.b64encode(out).decode(),'applied_changes':len(p['changes']),**({'printed_signer':printed_signer(final)} if is_standard(final) else {})}
if __name__=='__main__':
 try:
  raw=sys.stdin.buffer.read(15*1024*1024+1)
  if len(raw)>15*1024*1024:raise ValueError('Input too large')
  print(json.dumps(handle(json.loads(raw))))
 except Exception:
  print(json.dumps({'state':'held','reason':'document_revision_requires_review'}));sys.exit(2)
