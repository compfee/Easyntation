# Import necessary libraries
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE
import os
# Create PPT generator class
class PPTGenerator:
    # Initialize function, define PPT properties
    def __init__(self, title, slide_width=Inches(16), slide_height=Inches(9)):
        self.title = title
        self.prs = Presentation()
        self.slides = self.prs.slides
        # self.prs.slide_width = slide_width
        # self.prs.slide_height = slide_height



    def get_slide(self, index):
        return self.slides[index]

    def get_slide_layout(self, index):
        return self.prs.slide_layouts[index]

    def add_slide(self, layout):
        return self.slides.add_slide(layout)

    def delete_slide(self, index):
        self.slides._sldIdLst.remove(self.slides[index]._element)

    # Add title (and subtitle)
    def add_title(self, title_text, index_slide, subtitle=None):
        # title_slide_layout = self.prs.slide_layouts[0]
        # slide = self.prs.slides.add_slide(title_slide_layout)
        # title = slide.shapes.title
        # subtitle_box = slide.placeholders[1]
        #
        # title.text = title_text
        # subtitle.text = subtitle
        #
        slide = self.get_slide(index_slide)
        title_box = slide.shapes.title
        # title_box.margin_top = Inches(5)
        # title_box.margin_left = Inches(5)
        title_box.text = title_text
        subtitle_box = slide.placeholders[1]
        # title_box.margin_bottom = Inches()

        # title_box.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

        # title_box.word_wrap = False
        # title_box.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        if subtitle != None:
            subtitle_box.text = subtitle

    # Add Subtitle
    def add_subtitle(self, subtitle, index_slide, subtitle_left, subtitle_top):
        # slide = self.prs.slides[index_slide]
        # slide.shapes.title.text = subtitle

        slide = self.get_slide(index_slide)
        subtitle = slide.shapes.add_textbox(subtitle_left, subtitle_top, width=Inches(8), height=Inches(1))
        shape = slide.shapes
        title_shape = shape.title
        # title_shape.text_frame.paragraphs[0].font.size = Pt(10)
        title_shape.text_frame.paragraphs[0].font.size = Pt(15)
        # subtitle.text_frame.paragraphs[0].font.size = Pt(16)
        title_shape.text = subtitle

    # Add text
    def add_text(self, title, content, index_slide):
        slide = self.prs.slides[index_slide]

        title_box = slide.shapes.title
        title_box.text = title

        title_para = slide.shapes.title.text_frame.paragraphs[0]
        title_para.font.name = "Calibri"
        title_para.font.size = Pt(30)

        title_para.font.underline = True
        body_shape = slide.shapes.placeholders[1]
        text_place = body_shape.text_frame.paragraphs[0]
        text_place.font.name = "Calibri"
        text_place.font.size = Pt(20)
        text_place.text = content
        text_place.width = Inches(6)
        text_place.height = Inches(3)
        text_place.margin_bottom = Inches(0.08)
        text_place.margin_left = 0
        text_place.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
        text_place.word_wrap = False
        text_place.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT

        # left = top = Inches(2)
        # width = height = Inches(6)
        # text_place = slide.shapes.add_textbox(left, top, width, height)
        # tf = text_place
        # tf.text = content
        # p = tf.add_paragraph()
        # p.font.size = Pt(14)
        # p.font.name = "Calibri"

        # text_place.font.name = "Calibri"
        # text_place.font.size = Pt(14)

    # Add image
    def add_image(self, title, image_file, index_slide):
        slide = self.get_slide(index_slide)
        img = slide.shapes.add_picture(image_file, Inches(10.0), Inches(1.0),
                                       width=Inches(5), height=Inches(7.5))


    # Set background

    # Save PPT
    def save_ppt(self):
        ppt_name = '{}.pptx'.format(self.title)
        self.prs.save(ppt_name)

