import zipfile, xml.etree.ElementTree as ET

src = r'D:\ИИ\Git Projetcs\Nushensia\WorkFlow\Паспорта\Исходные протоколы\Протокол_П-673_КА_02.07.2026.docx'
with zipfile.ZipFile(src) as z:
    doc = z.read('word/document.xml')
    root = ET.fromstring(doc)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    paragraphs = root.findall('.//w:p', ns)
    with open(r'D:\temp_border_check.txt', 'w', encoding='utf-8') as out:
        for i, p in enumerate(paragraphs):
            texts = p.findall('.//w:t', ns)
            text = ''.join(t.text or '' for t in texts).strip()
            if not text:
                continue
            pPr = p.find('w:pPr', ns)
            has_pBdr = False
            pBdr_info = ''
            if pPr is not None:
                pBdr = pPr.find('w:pBdr', ns)
                if pBdr is not None:
                    has_pBdr = True
                    for child in pBdr:
                        tag = child.tag.split('}')[-1]
                        pBdr_info += '%s:%s ' % (tag, str(child.attrib))
            if i <= 15:
                out.write('[%d] text="%s" pBdr=%s %s\n' % (i, text[:80], has_pBdr, pBdr_info))
    
    print('Written')
