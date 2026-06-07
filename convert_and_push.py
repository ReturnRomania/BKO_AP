"""
Konvertiert alle .docx Dateien im BKO_AP Repo zu Markdown und pusht sie zurück.
Voraussetzung: pip install python-docx
Ausführen: python convert_and_push.py (im geklonten Repo-Ordner)
"""

import os
import io
from docx import Document

def heading_level(paragraph):
    style = paragraph.style.name
    if style.startswith('Heading'):
        try:
            return int(style.split(' ')[-1])
        except:
            return 1
    return 0

def docx_to_markdown(docx_path):
    doc = Document(docx_path)
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            lines.append('')
            continue
        level = heading_level(para)
        if level > 0:
            lines.append(f"{'#' * level} {text}")
        else:
            style = para.style.name.lower()
            if 'list' in style:
                lines.append(f"- {text}")
            else:
                lines.append(text)
    return '\n'.join(lines)

def convert_all(root_dir='.'):
    converted = []
    for dirpath, dirnames, filenames in os.walk(root_dir):
        # .git Ordner überspringen
        dirnames[:] = [d for d in dirnames if d != '.git']
        for filename in filenames:
            if filename.endswith('.docx'):
                docx_path = os.path.join(dirpath, filename)
                md_path = docx_path.replace('.docx', '.md')
                try:
                    md_content = docx_to_markdown(docx_path)
                    with open(md_path, 'w', encoding='utf-8') as f:
                        f.write(md_content)
                    print(f"✅ Konvertiert: {md_path}")
                    converted.append(md_path)
                except Exception as e:
                    print(f"❌ Fehler bei {docx_path}: {e}")
    return converted

if __name__ == '__main__':
    print("Starte Konvertierung...\n")
    files = convert_all('.')
    print(f"\n{len(files)} Dateien konvertiert.")
    print("\nJetzt pushen mit:")
    print("  git add .")
    print('  git commit -m "Konvertiere docx zu markdown"')
    print("  git push")
