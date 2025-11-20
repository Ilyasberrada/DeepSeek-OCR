#!/usr/bin/env python3
"""
Create a 3D presentation about "Le travail et la sagesse - Il faut cultiver notre jardin"
A philosophical exploration of Voltaire's famous quote from Candide
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def add_3d_effect(shape):
    """Add 3D effects to a shape"""
    # Add shadow for depth
    shadow = shape.shadow
    shadow.inherit = False
    shadow.visible = True
    shadow.distance = Pt(5)
    shadow.angle = 45
    shadow.blur_radius = Pt(4)
    shadow.transparency = 0.5

def create_title_slide(prs):
    """Create the title slide with 3D effects"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Background gradient
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 90
    fill.gradient_stops[0].color.rgb = RGBColor(25, 25, 112)  # Midnight blue
    fill.gradient_stops[1].color.rgb = RGBColor(65, 105, 225)  # Royal blue
    
    # Main title with 3D effect
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "Le Travail et La Sagesse"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(54)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 215, 0)  # Gold
    add_3d_effect(title_box)
    
    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = '"Il faut cultiver notre jardin"'
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.alignment = PP_ALIGN.CENTER
    subtitle_para.font.size = Pt(36)
    subtitle_para.font.italic = True
    subtitle_para.font.color.rgb = RGBColor(255, 255, 255)
    add_3d_effect(subtitle_box)
    
    # Author attribution
    author_box = slide.shapes.add_textbox(Inches(2), Inches(5.5), Inches(6), Inches(0.8))
    author_frame = author_box.text_frame
    author_frame.text = "- Voltaire, Candide (1759)"
    author_para = author_frame.paragraphs[0]
    author_para.alignment = PP_ALIGN.CENTER
    author_para.font.size = Pt(24)
    author_para.font.color.rgb = RGBColor(200, 200, 200)

def create_intro_slide(prs):
    """Introduction to Voltaire and Candide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 45
    fill.gradient_stops[0].color.rgb = RGBColor(47, 79, 79)  # Dark slate gray
    fill.gradient_stops[1].color.rgb = RGBColor(70, 130, 180)  # Steel blue
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Contexte Historique et Philosophique"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 215, 0)
    add_3d_effect(title_box)
    
    # Content box with 3D effect
    content_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(1.8), Inches(8), Inches(4.5)
    )
    content_box.fill.solid()
    content_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    content_box.line.color.rgb = RGBColor(255, 215, 0)
    content_box.line.width = Pt(3)
    add_3d_effect(content_box)
    
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.3)
    text_frame.margin_right = Inches(0.3)
    text_frame.margin_top = Inches(0.2)
    
    # Add content
    p = text_frame.paragraphs[0]
    p.text = "• Voltaire (1694-1778)"
    p.font.size = Pt(24)
    p.font.color.rgb = RGBColor(25, 25, 112)
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "• Philosophe des Lumières, critique de l'optimisme naïf"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(47, 79, 79)
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "• Candide (1759) : conte philosophique satirique"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(47, 79, 79)
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "• Critique de la philosophie de Leibniz : 'Tout est pour le mieux dans le meilleur des mondes possibles'"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(47, 79, 79)
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "• Conclusion : face aux malheurs du monde, il faut agir concrètement"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(47, 79, 79)
    p.font.bold = True

def create_work_slide(prs):
    """Slide about 'Le Travail' (Work)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 135
    fill.gradient_stops[0].color.rgb = RGBColor(34, 139, 34)  # Forest green
    fill.gradient_stops[1].color.rgb = RGBColor(144, 238, 144)  # Light green
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Le Travail : L'Action Concrète"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    add_3d_effect(title_box)
    
    # Left content box
    left_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(1.8), Inches(4), Inches(4.5)
    )
    left_box.fill.solid()
    left_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    left_box.line.color.rgb = RGBColor(34, 139, 34)
    left_box.line.width = Pt(3)
    add_3d_effect(left_box)
    
    text_frame = left_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.2)
    text_frame.margin_right = Inches(0.2)
    
    p = text_frame.paragraphs[0]
    p.text = "Signification :"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(34, 139, 34)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Remède contre l'ennui"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Protection contre le vice"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Source de dignité"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Moyen de subsistance"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Activité productive et créatrice"
    p.font.size = Pt(18)
    
    # Right content box
    right_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.5), Inches(1.8), Inches(4), Inches(4.5)
    )
    right_box.fill.solid()
    right_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    right_box.line.color.rgb = RGBColor(34, 139, 34)
    right_box.line.width = Pt(3)
    add_3d_effect(right_box)
    
    text_frame = right_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.2)
    text_frame.margin_right = Inches(0.2)
    
    p = text_frame.paragraphs[0]
    p.text = "Le travail selon Voltaire :"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(34, 139, 34)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "« Le travail éloigne de nous trois grands maux : l'ennui, le vice et le besoin. »"
    p.font.size = Pt(18)
    p.font.italic = True
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "Le travail n'est pas une punition, mais une libération."
    p.font.size = Pt(18)
    p.font.bold = True

def create_wisdom_slide(prs):
    """Slide about 'La Sagesse' (Wisdom)"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 90
    fill.gradient_stops[0].color.rgb = RGBColor(75, 0, 130)  # Indigo
    fill.gradient_stops[1].color.rgb = RGBColor(138, 43, 226)  # Blue violet
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "La Sagesse : L'Acceptation Lucide"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 215, 0)
    add_3d_effect(title_box)
    
    # Main content box
    content_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(1.8), Inches(8), Inches(4.5)
    )
    content_box.fill.solid()
    content_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    content_box.line.color.rgb = RGBColor(255, 215, 0)
    content_box.line.width = Pt(3)
    add_3d_effect(content_box)
    
    text_frame = content_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.3)
    text_frame.margin_right = Inches(0.3)
    text_frame.margin_top = Inches(0.2)
    
    p = text_frame.paragraphs[0]
    p.text = "La sagesse voltairienne :"
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = RGBColor(75, 0, 130)
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "• Renoncer aux spéculations métaphysiques stériles"
    p.font.size = Pt(20)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Accepter les limites de la condition humaine"
    p.font.size = Pt(20)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Privilégier l'action pratique à la contemplation passive"
    p.font.size = Pt(20)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Trouver le bonheur dans la simplicité et le concret"
    p.font.size = Pt(20)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Comprendre que nous ne pouvons pas tout contrôler"
    p.font.size = Pt(20)
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "La sagesse n'est pas dans la connaissance absolue, mais dans l'action mesurée."
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.italic = True
    p.font.color.rgb = RGBColor(138, 43, 226)

def create_garden_slide(prs):
    """Slide about the garden metaphor"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 180
    fill.gradient_stops[0].color.rgb = RGBColor(85, 107, 47)  # Dark olive green
    fill.gradient_stops[1].color.rgb = RGBColor(189, 183, 107)  # Dark khaki
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Cultiver Notre Jardin : La Métaphore"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    add_3d_effect(title_box)
    
    # Top content box
    top_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(1.8), Inches(8), Inches(1.8)
    )
    top_box.fill.solid()
    top_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    top_box.line.color.rgb = RGBColor(85, 107, 47)
    top_box.line.width = Pt(3)
    add_3d_effect(top_box)
    
    text_frame = top_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.3)
    text_frame.margin_right = Inches(0.3)
    
    p = text_frame.paragraphs[0]
    p.text = "Le jardin symbolise :"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(85, 107, 47)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "Notre sphère d'influence personnelle - ce que nous pouvons réellement contrôler et améliorer dans notre vie quotidienne."
    p.font.size = Pt(18)
    p.font.color.rgb = RGBColor(47, 79, 79)
    
    # Bottom content box
    bottom_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(4), Inches(8), Inches(2.3)
    )
    bottom_box.fill.solid()
    bottom_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    bottom_box.line.color.rgb = RGBColor(85, 107, 47)
    bottom_box.line.width = Pt(3)
    add_3d_effect(bottom_box)
    
    text_frame = bottom_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.3)
    text_frame.margin_right = Inches(0.3)
    
    p = text_frame.paragraphs[0]
    p.text = "Cultiver signifie :"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(85, 107, 47)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Prendre soin de soi et de ses proches"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Développer ses talents et compétences"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Construire une vie simple mais authentique"
    p.font.size = Pt(18)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Agir localement plutôt que de se perdre dans l'abstraction"
    p.font.size = Pt(18)

def create_application_slide(prs):
    """Slide about practical applications"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 45
    fill.gradient_stops[0].color.rgb = RGBColor(178, 34, 34)  # Firebrick
    fill.gradient_stops[1].color.rgb = RGBColor(255, 99, 71)  # Tomato
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Applications Pratiques Aujourd'hui"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 255, 255)
    add_3d_effect(title_box)
    
    # Left box
    left_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), Inches(1.8), Inches(4.3), Inches(4.5)
    )
    left_box.fill.solid()
    left_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    left_box.line.color.rgb = RGBColor(178, 34, 34)
    left_box.line.width = Pt(3)
    add_3d_effect(left_box)
    
    text_frame = left_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.2)
    text_frame.margin_right = Inches(0.2)
    
    p = text_frame.paragraphs[0]
    p.text = "Dans la vie moderne :"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(178, 34, 34)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• Se concentrer sur ce qu'on peut changer"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Éviter la paralysie par l'analyse"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Valoriser le travail manuel et intellectuel"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Trouver un équilibre vie-travail"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Cultiver des relations authentiques"
    p.font.size = Pt(17)
    
    # Right box
    right_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(5.2), Inches(1.8), Inches(4.3), Inches(4.5)
    )
    right_box.fill.solid()
    right_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    right_box.line.color.rgb = RGBColor(178, 34, 34)
    right_box.line.width = Pt(3)
    add_3d_effect(right_box)
    
    text_frame = right_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.2)
    text_frame.margin_right = Inches(0.2)
    
    p = text_frame.paragraphs[0]
    p.text = "Leçons essentielles :"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(178, 34, 34)
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "• L'action vaut mieux que la plainte"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Le bonheur est dans le faire, pas dans l'avoir"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• La simplicité volontaire comme choix de vie"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Responsabilité personnelle et engagement"
    p.font.size = Pt(17)
    p.space_after = Pt(8)
    
    p = text_frame.add_paragraph()
    p.text = "• Créer du sens par l'action"
    p.font.size = Pt(17)

def create_conclusion_slide(prs):
    """Create the conclusion slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Background
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = 90
    fill.gradient_stops[0].color.rgb = RGBColor(25, 25, 112)  # Midnight blue
    fill.gradient_stops[1].color.rgb = RGBColor(72, 61, 139)  # Dark slate blue
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Conclusion : Une Philosophie de Vie"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(40)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(255, 215, 0)
    add_3d_effect(title_box)
    
    # Main quote box
    quote_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1.5), Inches(2), Inches(7), Inches(2)
    )
    quote_box.fill.solid()
    quote_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    quote_box.line.color.rgb = RGBColor(255, 215, 0)
    quote_box.line.width = Pt(4)
    add_3d_effect(quote_box)
    
    text_frame = quote_box.text_frame
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    text_frame.margin_left = Inches(0.3)
    text_frame.margin_right = Inches(0.3)
    
    p = text_frame.paragraphs[0]
    p.text = '"Il faut cultiver notre jardin"'
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.italic = True
    p.font.color.rgb = RGBColor(25, 25, 112)
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(15)
    
    p = text_frame.add_paragraph()
    p.text = "Ce n'est pas une invitation au repli, mais un appel à l'action responsable et concrète."
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(72, 61, 139)
    p.alignment = PP_ALIGN.CENTER
    
    # Bottom summary box
    summary_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(1), Inches(4.5), Inches(8), Inches(1.8)
    )
    summary_box.fill.solid()
    summary_box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    summary_box.line.color.rgb = RGBColor(255, 215, 0)
    summary_box.line.width = Pt(3)
    add_3d_effect(summary_box)
    
    text_frame = summary_box.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.3)
    text_frame.margin_right = Inches(0.3)
    
    p = text_frame.paragraphs[0]
    p.text = "Travail + Sagesse = Vie accomplie"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(178, 34, 34)
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(10)
    
    p = text_frame.add_paragraph()
    p.text = "Agir dans notre sphère d'influence • Accepter nos limites • Trouver le bonheur dans l'action"
    p.font.size = Pt(18)
    p.font.color.rgb = RGBColor(47, 79, 79)
    p.alignment = PP_ALIGN.CENTER

def main():
    """Main function to create the presentation"""
    # Create presentation object
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    print("Creating presentation slides...")
    
    # Create all slides
    create_title_slide(prs)
    print("✓ Title slide created")
    
    create_intro_slide(prs)
    print("✓ Introduction slide created")
    
    create_work_slide(prs)
    print("✓ Work slide created")
    
    create_wisdom_slide(prs)
    print("✓ Wisdom slide created")
    
    create_garden_slide(prs)
    print("✓ Garden metaphor slide created")
    
    create_application_slide(prs)
    print("✓ Applications slide created")
    
    create_conclusion_slide(prs)
    print("✓ Conclusion slide created")
    
    # Save presentation
    output_file = "Le_Travail_et_La_Sagesse_3D.pptx"
    prs.save(output_file)
    print(f"\n✅ Presentation saved as: {output_file}")
    print(f"📊 Total slides: {len(prs.slides)}")
    print("\n🎨 Features:")
    print("  • 3D effects with shadows and depth")
    print("  • Gradient backgrounds")
    print("  • Professional typography")
    print("  • Rounded rectangles with borders")
    print("  • Color-coded themes per slide")

if __name__ == "__main__":
    main()
