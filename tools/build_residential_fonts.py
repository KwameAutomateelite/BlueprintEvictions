"""Build embedded editable-document font resources from locally installed TNR.

Only the PDF document is retained. Do not add standalone system font files.
This authoring utility requires reportlab/pypdf; production has no new dependency.
"""
from pathlib import Path
import json,struct
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from pypdf import PdfWriter
from pypdf.generic import ArrayObject,DictionaryObject,NameObject,NumberObject,FloatObject,DecodedStreamObject
ROOT=Path(__file__).resolve().parents[1]
w=PdfWriter();page=w.add_blank_page(612,792);fonts=DictionaryObject();widths={}
for key,file in [('TiRo','Times New Roman.ttf'),('TiBo','Times New Roman Bold.ttf')]:
 path=Path('/System/Library/Fonts/Supplemental')/file;raw=path.read_bytes();count=struct.unpack('>H',raw[4:6])[0]
 for i in range(count):
  tag,_,off,_=struct.unpack('>4sIII',raw[12+16*i:28+16*i])
  if tag==b'OS/2':assert struct.unpack('>H',raw[off+8:off+10])[0]==8,'Editable document embedding required'
 font=TTFont(key,str(path));pdfmetrics.registerFont(font);face=font.face;name=NameObject('/'+face.name.decode())
 embedded=DecodedStreamObject();embedded.set_data(raw);embedded[NameObject('/Length1')]=NumberObject(len(raw))
 desc=DictionaryObject({NameObject('/Type'):NameObject('/FontDescriptor'),NameObject('/FontName'):name,NameObject('/Flags'):NumberObject(34),NameObject('/FontBBox'):ArrayObject([FloatObject(x) for x in face.bbox]),NameObject('/ItalicAngle'):FloatObject(face.italicAngle),NameObject('/Ascent'):FloatObject(face.ascent),NameObject('/Descent'):FloatObject(face.descent),NameObject('/CapHeight'):FloatObject(face.capHeight),NameObject('/StemV'):NumberObject(face.stemV),NameObject('/FontFile2'):w._add_object(embedded.flate_encode())})
 chars=[];ws=[]
 for c in range(256):
  try:t=bytes([c]).decode('cp1252')
  except UnicodeDecodeError:t='?'
  chars.append(t);ws.append(round(pdfmetrics.stringWidth(t,key,1000),6))
 cmap=['/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def /CMapName /AMTNR def /CMapType 2 def 1 begincodespacerange <00> <ff> endcodespacerange']
 for start in range(0,256,100):
  rows=[f'<{i:02x}> <{chars[i].encode("utf-16-be").hex()}>' for i in range(start,min(start+100,256))];cmap.extend([str(len(rows))+' beginbfchar',*rows,'endbfchar'])
 cmap.append('endcmap CMapName currentdict /CMap defineresource pop end end');cm=DecodedStreamObject();cm.set_data('\n'.join(cmap).encode())
 f=DictionaryObject({NameObject('/Type'):NameObject('/Font'),NameObject('/Subtype'):NameObject('/TrueType'),NameObject('/BaseFont'):name,NameObject('/Encoding'):NameObject('/WinAnsiEncoding'),NameObject('/FirstChar'):NumberObject(0),NameObject('/LastChar'):NumberObject(255),NameObject('/Widths'):ArrayObject([FloatObject(x) for x in ws]),NameObject('/FontDescriptor'):w._add_object(desc),NameObject('/ToUnicode'):w._add_object(cm.flate_encode())})
 fonts[NameObject('/'+key)]=w._add_object(f);widths[key]=ws
page[NameObject('/Resources')]=DictionaryObject({NameObject('/Font'):fonts});w.write(ROOT/'templates/residential-input-fonts.pdf');(ROOT/'templates/residential-pdf-font-widths.json').write_text(json.dumps(widths,separators=(',',':')))
print('Embedded Times New Roman regular/bold document resources built')
