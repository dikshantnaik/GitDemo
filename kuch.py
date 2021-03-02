from PIL import Image, ImageDraw, ImageFont
import textwrap

font_size = 60
text = 'your text'
rgb_color = (255,255,255)

img = Image.open('Cert.PNG')
draw = ImageDraw.Draw(img)

font = ImageFont.truetype('tango.ttf', font_size)

w, h = draw.textsize(text, font)

left = (img.width - w) / 2
top = 22
img.save("TEst.PNG")