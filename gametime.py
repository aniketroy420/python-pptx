from pptx import Presentation
from pptx.util import Inches, Pt,Cm
from pptx.dml.color import RGBColor
# from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

prs = Presentation()

# Choose a slide layout (consider using one with a title if desired)
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

# Define image path
logo_path = "logo.png"
paytm_img = "paytm.png"  # Replace with your actual image path
lookbeyond_img = "lookbeyond.png"
logo_left = 0.5  # Adjust these values to reposition the image
top = 0.5
paytm_left = 8

background = slide.background
fill = background.fill
fill.solid()
fill.fore_color.rgb = RGBColor(37, 150, 190)

# left = top = width = height = Inches(1.0)
# slide.shapes.add_shape(
#     MSO_SHAPE.OVAL,Inches(paytm_left), Inches(top), width, height
# )

# Add the image (without resize parameters for now)
picture = slide.shapes.add_picture(logo_path, Inches(logo_left), Inches(top),height = Cm(1))
paytm = slide.shapes.add_picture(paytm_img,Inches(paytm_left), Inches(top),height = Cm(1))
lookbeyond= slide.shapes.add_picture(lookbeyond_img,left =0, top = Inches(6.7),width = Inches(10))

## Set textbox position (in inches)
left = Inches(0.5)
top = prs.slide_height//2.5
width = 6
height = 1

# Create textbox object
textbox = slide.shapes.add_textbox(left, top,width,height)

# Get text frame (to access text properties)
tf = textbox.text_frame

# Add text
text_content = [
    "CONFIDENTIAL",
    "PAYTM JUNE 2023"
]
for i in range(len(text_content)):
    p = tf.add_paragraph()
    p.font.bold = True
    p.text = text_content[i]
    p.font.name = 'Calibri'
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.LEFT
    if(i == 0):
        p.font.size = Pt(43.5)
    else:
        p.font.size = Pt(17.5)
    

# for text_line in text_content:
#     p = tf.add_paragraph()
#     p.text = text_line
#     p.font.size = Pt(43.5)
#     p.font.name = 'Calibri'
#     p.alignment = PP_ALIGN.LEFT
print(prs.slide_width)
print(prs.slide_height)
# Save the presentation
prs.save("1.pptx")




