import csv,uuid,copy,base64
from pathlib import Path
from core import *

def pdf_bytes(label):
    stream=f'BT /F1 18 Tf 40 760 Td ({label}) Tj ET'.encode('ascii')
    objects=[b'<< /Type /Catalog /Pages 2 0 R >>',b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',b'<< /Length '+str(len(stream)).encode()+b' >>\nstream\n'+stream+b'\nendstream']
    out=bytearray(b'%PDF-1.4\n');offsets=[0]
    for i,o in enumerate(objects,1):offsets.append(len(out));out.extend(str(i).encode()+b' 0 obj\n'+o+b'\nendobj\n')
    start=len(out);out.extend(b'xref\n0 6\n0000000000 65535 f \n')
    for offset in offsets[1:]:out.extend(f'{offset:010} 00000 n \n'.encode())
    out.extend(f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{start}\n%%EOF'.encode());return bytes(out)

PNG=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jB1kAAAAASUVORK5CYII=')

def make_sample(root):
    root=Path(root);materials=root/'材料';materials.mkdir(parents=True,exist_ok=True)
    p=Project('10人虚构示例');p.rules=copy.deepcopy(TEMPLATES['活动报名材料'])
    for i in range(10):
        person=Person(uuid.uuid4().hex,f'示例人员{i+1}',f'{i+1:04}');p.people.append(person)
        if i!=9:(materials/f'{person.identifier}_报名.pdf').write_bytes(pdf_bytes(f'Fictional registration {i+1}'))
        if i<7:(materials/f'{person.identifier}_照片.png').write_bytes(PNG)
    (materials/'匿名照片.png').write_bytes(PNG)
    with (root/'示例名单.csv').open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['姓名','学号']);w.writerows((x.name,x.identifier) for x in p.people)
    files,_=scan(materials);merge_scan(p,materials,files);return p
