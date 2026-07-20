from PIL import Image

image = Image.open('tmp/pdfs/sds_pages/page-4.png')
width, height = image.size
image.crop((0, 0, width, height)).save('tmp/pdfs/sds_pages/page-4-crop.png')
