#################보고서 제작#################보고서 제작#################보고서 제작#################보고서 제작#################

# 보고서 제작 class,def 모음
from turtle import width
from docx import Document
from docx.enum.section import WD_ORIENTATION

#Cm,Inches,Pt 단위를 사용하기 위한 모듈
from docx.shared import Cm,Inches,Pt

# 문자 스타일 변경
from docx.enum.style import WD_STYLE_TYPE
from pyparsing import col

# table border style
from docx.table import _Cell
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls

# para 정렬
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK

# table 정렬
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.section import WD_ORIENTATION


# font color
from docx.shared import RGBColor

# -*- coding: utf-8 -*-

# filename: add_float_picture.py

'''
Implement floating image based on python-docx.
- Text wrapping style: BEHIND TEXT <wp:anchor behindDoc="1">
- Picture position: top-left corner of PAGE `<wp:positionH relativeFrom="page">`.
Create a docx sample (Layout | Positions | More Layout Options) and explore the
source xml (Open as a zip | word | document.xml) to implement other text wrapping
styles and position modes per `CT_Anchor._anchor_xml()`.
'''

from docx.oxml import parse_xml, register_element_cls
from docx.oxml.ns import nsdecls
from docx.oxml.shape import CT_Picture
from docx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne
from matplotlib import style

# 사진 사이즈 및 위치 변경
# refer to docx.oxml.shape.CT_Inline
class CT_Anchor(BaseOxmlElement):
    """
    ``<w:anchor>`` element, container for a floating image.
    """
    extent = OneAndOnlyOne('wp:extent')
    docPr = OneAndOnlyOne('wp:docPr')
    graphic = OneAndOnlyOne('a:graphic')

    @classmethod
    def new(cls, cx, cy, shape_id, pic, pos_x, pos_y):
        """
        Return a new ``<wp:anchor>`` element populated with the values passed
        as parameters.
        """
        anchor = parse_xml(cls._anchor_xml(pos_x, pos_y))
        anchor.extent.cx = cx
        anchor.extent.cy = cy
        anchor.docPr.id = shape_id
        anchor.docPr.name = 'Picture %d' % shape_id
        anchor.graphic.graphicData.uri = (
            'http://schemas.openxmlformats.org/drawingml/2006/picture'
        )
        anchor.graphic.graphicData._insert_pic(pic)
        return anchor

    @classmethod
    def new_pic_anchor(cls, shape_id, rId, filename, cx, cy, pos_x, pos_y):
        """
        Return a new `wp:anchor` element containing the `pic:pic` element
        specified by the argument values.
        """
        pic_id = 0  # Word doesn't seem to use this, but does not omit it
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        anchor = cls.new(cx, cy, shape_id, pic, pos_x, pos_y)
        anchor.graphic.graphicData._insert_pic(pic)
        return anchor

    @classmethod
    def _anchor_xml(cls, pos_x, pos_y):
        return (
                '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" \n'
                '           behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1" \n'
                '           %s>\n'
                '  <wp:simplePos x="0" y="0"/>\n'
                '  <wp:positionH relativeFrom="page">\n'
                '    <wp:posOffset>%d</wp:posOffset>\n'
                '  </wp:positionH>\n'
                '  <wp:positionV relativeFrom="page">\n'
                '    <wp:posOffset>%d</wp:posOffset>\n'
                '  </wp:positionV>\n'
                '  <wp:extent cx="914400" cy="914400"/>\n'
                '  <wp:wrapNone/>\n'
                '  <wp:docPr id="666" name="unnamed"/>\n'
                '  <wp:cNvGraphicFramePr>\n'
                '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
                '  </wp:cNvGraphicFramePr>\n'
                '  <a:graphic>\n'
                '    <a:graphicData uri="URI not set"/>\n'
                '  </a:graphic>\n'
                '</wp:anchor>' % (nsdecls('wp', 'a', 'pic', 'r'), int(pos_x), int(pos_y))
        )


# refer to docx.parts.story.BaseStoryPart.new_pic_inline
def new_pic_anchor(part, image_descriptor, width, height, pos_x, pos_y):
    """Return a newly-created `w:anchor` element.
    The element contains the image specified by *image_descriptor* and is scaled
    based on the values of *width* and *height*.
    """
    rId, image = part.get_or_add_image(image_descriptor)
    cx, cy = image.scaled_dimensions(width, height)
    shape_id, filename = part.next_id, image.filename
    return CT_Anchor.new_pic_anchor(shape_id, rId, filename, cx, cy, pos_x, pos_y)


# refer to docx.text.run.add_picture
def add_float_picture(p, image_path_or_stream, width=None, height=None, pos_x=0, pos_y=0):
    """Add float picture at fixed position `pos_x` and `pos_y` to the top-left point of page.
    """
    run = p.add_run()
    anchor = new_pic_anchor(run.part, image_path_or_stream, width, height, pos_x, pos_y)
    run._r.add_drawing(anchor)


# refer to docx.oxml.__init__.py
register_element_cls('wp:anchor', CT_Anchor)

# 문단 텍스트 입력
def paragraphText(paragraph,text,fontsize,color,alignment,style):
    text = paragraph.add_run(text)
    text.font.size = Pt(fontsize) 
    font = text.font
    font.color.rgb = RGBColor.from_string(color)
    if alignment == 'CENTER':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        pass
    if style == 'bold':
        text.bold = True
    else:
        pass

def lineSpace(doc,inches,space_before,space_after):
    lineSpace = doc.add_paragraph()
    lineSpace.paragraph_format.line_spacing = Inches(inches)
    lineSpace.paragraph_format.space_before = Pt(space_before)
    lineSpace.paragraph_format.space_after = Pt(space_after)

# table 제작
def makeTable(paragraph,row,col,alignment,width,height):

    table = paragraph.add_table(rows=row, cols=col)
    if width == None:
        pass
    else:
        set_col_widths(table,width)

    if height == None:
        pass
    else:
        set_col_height(table,height)    

    if alignment == 'CENTER':
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
    else:
        pass
    return table

# table 가로 크기 변경
def set_col_widths(table,widths):
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

# table 세로 크기 변경
def set_col_height(table,heights):
    for idx, row in enumerate(table.rows):
        row.height = heights[idx]

# table 전체 스타일 변경
def titleBorder(
                table,
                top_val,
                top_color,
                top_sz,
                bottom_val,
                bottom_color,
                bottom_sz,
                left_val,
                left_color,
                left_sz,
                right_val,
                right_color,
                right_sz
                ):

    tbl = table._tbl # get xml element in table
    for cell in tbl.iter_tcs():
        tcPr = cell.tcPr # get tcPr element, in which we can define style of borders
        tcBorders = OxmlElement('w:tcBorders')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), top_val)
        top.set(qn('w:color'), top_color)
        top.set(qn('w:sz'), top_sz) 

        left = OxmlElement('w:left')
        left.set(qn('w:val'), left_val)
        left.set(qn('w:color'), left_color)
        left.set(qn('w:sz'), left_sz)
        
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), bottom_val)
        bottom.set(qn('w:color'), bottom_color)
        bottom.set(qn('w:sz'), bottom_sz)

        right = OxmlElement('w:right')
        right.set(qn('w:val'), right_val)
        right.set(qn('w:color'), right_color)
        right.set(qn('w:sz'), right_sz)

        tcBorders.append(top)
        tcBorders.append(left)
        tcBorders.append(bottom)
        tcBorders.append(right)
        tcPr.append(tcBorders)

# table border 부분 변경
# def set_cell_border(cell: _Cell, **kwargs):
def set_cell_border(table,row,col,**kwargs):
    """
    Set cell`s border
    Usage:
    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )
    """
    cell = table.rows[row].cells[col]
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

# 셀 개별 배경색 변환
def cellBackColor(table,row,cell,color):
    #GET CELLS XML ELEMENT
    cell_xml_element = table.rows[row].cells[cell]._tc
    #RETRIEVE THE TABLE CELL PROPERTIES
    table_cell_properties = cell_xml_element.get_or_add_tcPr()
    #CREATE SHADING OBJECT
    shade_obj = OxmlElement('w:shd')
    #SET THE SHADING OBJECT
    shade_obj.set(qn('w:fill'), color)
    #APPEND THE PROPERTIES TO THE TABLE CELL PROPERTIES
    table_cell_properties.append(shade_obj)

# 셀 텍스트 삽입
def insertTextCell(table,row,col,text,color,fontSize,fontStyle,vertical_alignment,para_alignment,space_before,space_after):
    cell = table.rows[row].cells[col]
    paragraph = cell.paragraphs[0]
    inputText = paragraph.add_run(text)
    font = inputText.font
    font.color.rgb = RGBColor.from_string(color)

    if fontStyle == 'bold':
        inputText.bold = True
    else:
        pass

    if vertical_alignment == 'CENTER':
        cell.vertical_alignment  = WD_ALIGN_VERTICAL.CENTER
    elif vertical_alignment == 'TOP':
        cell.vertical_alignment  = WD_ALIGN_VERTICAL.TOP
    elif vertical_alignment == 'BOTTOM':
        cell.vertical_alignment  = WD_ALIGN_VERTICAL.BOTTOM
    elif vertical_alignment == 'BOTH':
        cell.vertical_alignment  = WD_ALIGN_VERTICAL.BOTH
    else:
        pass

    if para_alignment == 'CENTER':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif para_alignment == 'RIGHT':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif para_alignment == 'LEFT':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    else:
        pass
    
    paragraph.paragraph_format.space_before = space_before
    paragraph.paragraph_format.space_after = space_after

    font.size = Pt(fontSize)

# 셀 병합
def cellMerge(table,stdCellList,cellList):
    stdCell = table.cell(stdCellList[0],stdCellList[1])
    for i in cellList:
        mergeCell = table.cell(i[0],i[1])
        stdCell.merge(mergeCell)

# header 박스 만들기
def makeHeaderBox(No,text):
    headerBoxWidth = [Cm(1.5), Cm(1), Cm(24.5)]
    headerBoxHeight = [Cm(1)]
    headerBox = makeTable(doc,row=1,col=3,alignment='CENTER',width=headerBoxWidth,height=headerBoxHeight)
    set_cell_border(
        headerBox,
        row=0,
        col=2,
        top={"val": "nil"},
        bottom={"val":"single", "sz":"15","color":"#F58D22"},
        start={"val":"nil"},
        end={"val":"nil"}
    )

    insertTextCell(
                    headerBox,
                    row=0,
                    col=0,
                    text=No,
                    color='FFFFFF',
                    fontSize=20,
                    fontStyle='bold',
                    vertical_alignment='CENTER',
                    para_alignment='CENTER',
                    space_after=Pt(0),
                    space_before=Pt(0)
                    )
    insertTextCell(
                    headerBox,
                    row=0,
                    col=2,
                    text=text,
                    color='000000',
                    fontSize=20,
                    fontStyle='bold',
                    vertical_alignment='CENTER',
                    para_alignment='LEFT',
                    space_after=Pt(0),
                    space_before=Pt(0)
                    )

    cellBackColor(headerBox, row=0,cell=0,color="F58D22")

# header 반복 함수
def set_repeat_table_header(row):
    """ set repeat table row on every new page
    """
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), "true")
    trPr.append(tblHeader)
    return row


farmNameReport = 'test'

# 문서 생성
doc = Document()

# 문서 전체 폰트 변경
style = doc.styles['Normal']
style.font.name = '맑은 고딕'
style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

# 전체 페이지 가로세로 설정
current_section = doc.sections[-1]

# current_section.orientation = WD_ORIENTATION.LANDSCAPE
# current_section.page_width = Cm(29.7)
# current_section.page_height = Cm(21.0)

current_section.top_margin = Cm(1)
current_section.bottom_margin = Cm(1)
current_section.left_margin = Cm(1)
current_section.right_margin = Cm(1)



# page1 표지 만들기

sign = doc.add_paragraph()

# 사진의 크기를 Cm 단위로 설정하여 삽입
add_float_picture(sign,'./signPicture/page.png',width=Cm(21.59), height=Cm(27.94),pos_x=Cm(0), pos_y=Cm(0))

lineSpace(doc,inches=0.6,space_before=0,space_after=0)
lineSpace(doc,inches=0.6,space_before=0,space_after=0)
lineSpace(doc,inches=0.6,space_before=0,space_after=0)

paraTitle1 = doc.add_paragraph()
title1 = paragraphText(paraTitle1,text='한우',fontsize = 32,color='000000',alignment='CENTER',style='bold')

paraTitle2 = doc.add_paragraph()
title2 = paragraphText(paraTitle2,text='유전체 보고서',fontsize = 32,color='000000',alignment='CENTER',style='bold')

paraTitle3 = doc.add_paragraph()
title3 = paragraphText(paraTitle3,text=f'{farmNameReport} 농장',fontsize = 18,color='000000',alignment='CENTER',style='bold')

# page1 표지 만들기 완료

doc.save('유전체보고서.docx')