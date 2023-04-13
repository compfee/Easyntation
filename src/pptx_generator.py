# Import necessary libraries
from pptx import Presentation
from pptx.util import Inches, Pt
import os
# Create PPT generator class
class PPTGenerator:
    # Initialize function, define PPT properties
    def __init__(self, title):
        self.title = title
        self.prs = Presentation()
        self.slides = self.prs.slides
        self.slide_width = self.prs.slide_width
        self.slide_height = self.prs.slide_height

    def get_slide(self, index):
        return self.slides[index]

    def get_slide_layout(self, index):
        return self.prs.slide_layouts[index]

    def add_slide(self, layout):
        return self.slides.add_slide(layout)

    def delete_slide(self, index):
        self.slides._sldIdLst.remove(self.slides[index]._element)

    # Add title (and subtitle)
    def add_title(self, title, index_slide, subtitle=None):
        slide = self.get_slide(index_slide)
        title_box = slide.shapes.title
        title_box.text = title
        subtitle_box = slide.placeholders[1]
        if subtitle != None:
            subtitle_box.text = subtitle

    # Add Subtitle
    def add_subtitle(self, subtitle, index_slide, subtitle_left, subtitle_top):
        slide = self.get_slide(index_slide)
        subtitle = slide.shapes.add_textbox(subtitle_left, subtitle_top, width=Inches(8), height=Inches(1))
        subtitle.text_frame.paragraphs[0].font.size = Pt(24)

    # Add text
    def add_text(self, title, content, index_slide):
        slide = self.get_slide(index_slide)
        title_box = slide.shapes.title
        title_box.text = title
        body_shape = slide.shapes.placeholders[1]
        body_shape.text_frame.paragraphs[0].font.size = Pt(10)
        tf = body_shape.text_frame
        tf.text = content

    # Add image
    def add_image(self, title, image_file, index_slide):
        slide = self.get_slide(index_slide)
        img = slide.shapes.add_picture(image_file, Inches(6.0), Inches(0.0),
                                       width=Inches(4), height=Inches(7.5))

    # Set background

    # Save PPT
    def save_ppt(self):
        ppt_name = '{}.pptx'.format(self.title)
        self.prs.save(ppt_name)

