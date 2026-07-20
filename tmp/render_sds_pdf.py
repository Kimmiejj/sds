from pathlib import Path
import pypdfium2 as pdfium

source = Path('แบบทดสอบ S.D.S. - กระดาษคำตอบ.pdf')
out = Path('tmp/pdfs/sds_pages')
out.mkdir(parents=True, exist_ok=True)
document = pdfium.PdfDocument(source)
for index in range(len(document)):
    image = document[index].render(scale=2.0).to_pil()
    image.save(out / f'page-{index + 1}.png')
print(f'Rendered {len(document)} pages to {out}')
