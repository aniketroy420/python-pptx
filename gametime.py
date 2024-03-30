from pptx import Presentation
from pptx.util import Inches, Pt,Cm
from pptx.dml.color import RGBColor
# from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

prs = Presentation()


blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)


# this sets backgrund color of the slide
# background = slide.background
# fill = background.fill
# fill.solid()
# fill.fore_color.rgb = RGBColor(37, 150, 190)

#  image path
logo_path = "logo.png"
paytm_img = "paytm.png"  
lookbeyond_img = "lookbeyond.png"
logo_left = 0.5  # Adjust these values to reposition the image
top_img = 0.5
paytm_left = 8


#This sets the background image of the slide
left = top = 0
img_path = "bg.jpg"
pic = slide.shapes.add_picture(img_path, left, top, width=prs.slide_width, height=prs.slide_height)
slide.shapes._spTree.remove(pic._element)
slide.shapes._spTree.insert(2, pic._element)

# Adds the image 
picture = slide.shapes.add_picture(logo_path, Inches(logo_left), Inches(top_img),height = Cm(1))
paytm = slide.shapes.add_picture(paytm_img,Inches(paytm_left), Inches(top_img),height = Cm(1))
lookbeyond= slide.shapes.add_picture(lookbeyond_img,left =0, top = Inches(6.7),width = Inches(10))

## Sets textbox position 
left = Inches(0.5)
top = prs.slide_height//2.5
width = 6
height = 1

# Creates textbox object
textbox = slide.shapes.add_textbox(left, top,width,height)

# Gets text frame (to access text properties)
tf = textbox.text_frame

# Add your text here
text_content = [
    "CONFIDENTIAL",
    "PAYTM JUNE 2023"
]
#Use different properties for the different text available in the above list
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
    
print(prs.slide_width)
print(prs.slide_height)
# Save the presentation
prs.save("1.pptx")
