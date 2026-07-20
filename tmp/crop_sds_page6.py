from PIL import Image

image = Image.open('tmp/pdfs/sds_pages/page-6.png')
width, height = image.size
image.crop((0, int(height * 0.54), width, height)).save('tmp/pdfs/sds_pages/page-6-crop.png')
