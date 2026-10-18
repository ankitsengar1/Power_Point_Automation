import pandas as pd 
import numpy as np
from pptx.util import Pt,Inches
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
#from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from formatting_functions import format_with_locale
from pptx.oxml.xmlchemy import OxmlElement
from variables import current_month

# def edit_table(slide_28,df,repeat_sales_colors_dict):
#     for shape in slide_28.shapes:
#         if shape.has_table:
#             table = shape.table
#             for col in range(0,len(table.columns)):
#                 for row in range(0,len(table.rows)):
#                     for paragraph in table.cell(row,col).text_frame.paragraphs:
#                         first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
#                         font = first_run.font
#                         font_name = font.name
#                         font_size = font.size
#                         font_bold = font.bold
#                         font_italic = font.italic
#                         font_underline = font.underline
#                         text = paragraph.text
#                         if font.color.type == MSO_COLOR_TYPE.SCHEME:
#                             # Get RGB from theme color
#                             theme_color = font.color.theme_color
#                             font_rgb_flag = 1
#                         elif font.color.type == None:
#                             font_rgb_flag = 0
#                         elif font.color.type == MSO_COLOR_TYPE.PRESET:
#                             color_index = font.color.theme_color
#                             font_rgb_flag = 3
#                         else:
#                             # Already RGB color, use it directly
#                             rgb = font.color.rgb
#                             font_rgb_flag = 2
                
#                         paragraph.clear()                    
#                         new_run = paragraph.add_run()
#                         new_run.font.name = font_name
#                         new_run.font.size = font_size
#                         new_run.font.bold = font_bold
#                         new_run.font.underline = font_underline
#                         new_run.font.italic = font_italic
#                         # print(font_rgb_flag)
#                         if font_rgb_flag == 1:
#                             new_run.font.color.theme_color = theme_color
#                         elif font_rgb_flag == 2:
#                             new_run.font.color.rgb = rgb
#                         elif font_rgb_flag == 3:
#                             new_run.font.color.color_name = color_name = MSO_THEME_COLOR_INDEX(color_index).name

                    
#                         if row > 0 and col>=0: # leaving headers and first column values
#                             # print(df.iloc[row-2,col])
#                             new_run.text = str(df.iloc[row-1,col])

#                             if col==3:
#                                 table.cell(row,col).fill.solid()
#                                 value = float(str(df.iloc[row-1,col])[:-1])
#                                 #print(loyalty_sales_colors_dict[value][0],loyalty_sales_colors_dict[value][1],loyalty_sales_colors_dict[value][2])
#                                 table.cell(row,col).fill.fore_color.rgb = RGBColor(repeat_sales_colors_dict[value][0],repeat_sales_colors_dict[value][1],repeat_sales_colors_dict[value][2])

#                         elif row == 0 and col>=0: # headers
#                             # print(str(list(df.columns.values)[col]))
#                             new_run.text = str(list(df.columns.values)[col])
                        
#     return slide_28


def SubElement(parent, tagname, **kwargs):
        element = OxmlElement(tagname)
        element.attrib.update(kwargs)
        parent.append(element)
        return element

def _set_cell_border(cell,border_color, border_width='5000'):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for lines in ['a:lnL','a:lnR','a:lnT','a:lnB']:
        # if lines=='a:lnB':
        ln = SubElement(tcPr, lines, w=border_width, cap='flat', cmpd='sng', algn='ctr')
        solidFill = SubElement(ln, 'a:solidFill')
        srgbClr = SubElement(solidFill, 'a:srgbClr', val=border_color)
        prstDash = SubElement(ln, 'a:prstDash', val='solid')
        round_ = SubElement(ln, 'a:round')
        headEnd = SubElement(ln, 'a:headEnd', type='none', w='med', len='med')
        tailEnd = SubElement(ln, 'a:tailEnd', type='none', w='med', len='med')

def create_table(slide_28,df,repeat_sales_colors_dict):

    rows, cols = df.shape
    table_width = Inches(12.48)  # Adjust based on desired width  ####
    table_height = Inches(4.66)  # Adjust based on desired height
    #ht = 3.9 -((len(df)*0.32)/2)
    table = slide_28.shapes.add_table(rows + 1, cols, Inches(0.21),Inches(0.54),table_width, table_height)
    # table_.table.

    header_colors = [RGBColor(51, 63, 79), RGBColor(51, 63, 79), RGBColor(51, 63, 79), RGBColor(51, 63, 79)]
    
    # Add column headers
    for col in range(cols):
        _set_cell_border(table.table.cell(0, col), border_color="000000", border_width="12700")
        table.table.cell(0, col).text_frame.text = df.columns[col]
        table.table.cell(0, col).text_frame.paragraphs[0].font.size = Pt(12)
        table.table.cell(0, col).fill.solid()
        table.table.cell(0, col).fill.fore_color.rgb = header_colors[col % len(header_colors)]
        table.table.cell(0, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    table.table.rows[0].height = Inches(0.23)

    # Populate table data
    for row in range(1, rows + 1):
        table.table.rows[row].height = Inches(0.23)
        for col in range(cols):
            # table_1.table.cell(row,col).text_frame.height = Inches(0.3)
            _set_cell_border(table.table.cell(row, col), border_color="000000", border_width="12700")

            if col==0:
                table.table.columns[col].width = Inches(1.6)
                table.table.cell(row, col).text_frame.text = str(df.iloc[row - 1, col])
            else:
                table.table.cell(row, col).text_frame.text = str(df.iloc[row - 1, col])
                table.table.columns[col].width = Inches(2.22)
            
            if col==3:
                table.table.cell(row,col).fill.solid()
                value = int(str(df.iloc[row - 1, col])[:-1])
                table.table.cell(row,col).fill.fore_color.rgb = RGBColor(repeat_sales_colors_dict[value][0],repeat_sales_colors_dict[value][1],repeat_sales_colors_dict[value][2])
            elif col==0:
                table.table.cell(row, col).fill.solid()
                table.table.cell(row, col).fill.fore_color.rgb = RGBColor(51, 63, 79)
                table.table.cell(row, col).text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
            else:
                table.table.cell(row, col).fill.solid()
                table.table.cell(row, col).fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            table.table.cell(row, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            table.table.cell(row, 0).text_frame.paragraphs[0].font.bold = True
            # table
            table.table.cell(row, col).text_frame.paragraphs[0].font.size = Pt(12)


    return slide_28

def create_df(senco_ppt_xl_path,sheet_name):
    df = pd.read_excel(senco_ppt_xl_path,sheet_name = sheet_name)

    df['% Repeaters'] = (df['repeater']/df['loyalty_customer']*100).round(0).astype(int).astype(str) + '%'
    df['% Repeat Sales'] = (df['repater_sales']/df['loyalty_sales']*100).round(0).astype(int).astype(str) + '%'
    df['ATV'] = df['loyalty_sales']/df['loyalty_bills']
    df['% Sales Contribution'] = ((df['loyalty_sales']/df['loyalty_sales'].sum())*100).round(1).astype(str) + '%'
    df = df.sort_values(by='% Repeat Sales', ascending=False)
    df = df[['state','store_count','% Repeaters','% Repeat Sales','ATV','% Sales Contribution']]
    df.rename(columns={'state':'States','store_count':'Store Count'},inplace = True)
    df['ATV'] = df['ATV'].apply(format_with_locale)
    df = df[~((df['% Repeaters'] == '100%') & (df['% Repeat Sales'] == '100%'))]

    return df