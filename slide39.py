import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import PercentFormatter, FuncFormatter
from matplotlib import font_manager
import matplotlib.pyplot as plt
import matplotlib.cm as cm
import numpy as np
from matplotlib.colors import LinearSegmentedColormap
# import datetime as dt
import numpy as np
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from pptx.util import Pt,Inches


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



def create_table(slide,df_1,df_2,df_semi_old_colors_dict,df_new_colors_dict):

    df =df_1
    rows, cols = df_1.shape
    table_width = Inches(8)  # Adjust based on desired width  ####
    table_height = Inches(7)  # Adjust based on desired height
    # Custom Height
    ht = 3.9 -((len(df)*0.32)/2)
    table_1 = slide.shapes.add_table(rows + 1, cols, Inches(0.32),Inches(ht),table_width, table_height)
    # table_.table.

    header_colors = [RGBColor(51, 63, 79), RGBColor(51, 63, 79), RGBColor(51, 63, 79), RGBColor(51, 63, 79)]
    
    # Add column headers
    for col in range(cols):
        _set_cell_border(table_1.table.cell(0, col), border_color="000000", border_width="12700")
        table_1.table.cell(0, col).text_frame.text = df.columns[col]
        table_1.table.cell(0, col).text_frame.paragraphs[0].font.size = Pt(9)
        table_1.table.cell(0, col).fill.solid()
        table_1.table.cell(0, col).fill.fore_color.rgb = header_colors[col % len(header_colors)]
        table_1.table.cell(0, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    table_1.table.rows[0].height = Inches(0.2)

    # Populate table data
    for row in range(1, rows + 1):
        table_1.table.rows[row].height = Inches(0.32)
        for col in range(cols):
            # table_1.table.cell(row,col).text_frame.height = Inches(0.3)
            _set_cell_border(table_1.table.cell(row, col), border_color="000000", border_width="12700")
            
            if col==0:
                table_1.table.columns[col].width = Inches(3.25)
                table_1.table.cell(row, col).text_frame.text = str(df.iloc[row - 1, col])
                table_1.table.cell(row, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            else:
                table_1.table.cell(row, col).text_frame.text = str(df.iloc[row - 1, col])
                table_1.table.columns[col].width = Inches(1)

            table_1.table.cell(row,col).text_frame.margin_bottom = Inches(0.07)
            table_1.table.cell(row,col).text_frame.margin_top = Inches(0.07)

            if len(str(df.iloc[row - 1, col]))>=35:
                table_1.table.cell(row, col).text_frame.paragraphs[0].font.size = Pt(8)
            else:
                table_1.table.cell(row, col).text_frame.paragraphs[0].font.size = Pt(9)
            
            if col==2:
                table_1.table.cell(row,col).fill.solid()
                value = int(str(df.iloc[row - 1, col])[:-1])
                table_1.table.cell(row,col).fill.fore_color.rgb = RGBColor(df_semi_old_colors_dict[value][0],df_semi_old_colors_dict[value][1],df_semi_old_colors_dict[value][2])
            elif col==0:
                table_1.table.cell(row, col).fill.solid()
                table_1.table.cell(row, col).fill.fore_color.rgb = RGBColor(51, 63, 79)
                table_1.table.cell(row, col).text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
            else:
                table_1.table.cell(row, col).fill.solid()
                table_1.table.cell(row, col).fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            if col!=0:
                table_1.table.cell(row, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            table_1.table.cell(row, 0).text_frame.paragraphs[0].font.bold = True
            # table

    rows, cols = df_2.shape
    df = df_2
    table_width = Inches(8)  # Adjust based on desired width
    table_height = Inches(7)  # Adjust based on desired height
    # Custom Height
    table_2 = slide.shapes.add_table(rows + 1, cols, Inches(6.75),Inches(ht),table_width, table_height)
    # table_.table.
    header_colors = [RGBColor(51, 63, 79), RGBColor(51, 63, 79), RGBColor(51, 63, 79), RGBColor(51, 63, 79)]
    
    # Add column headers
    for col in range(cols):
        _set_cell_border(table_2.table.cell(0, col), border_color="000000", border_width="12700")
        table_2.table.cell(0, col).text_frame.text = df.columns[col]
        table_2.table.cell(0, col).text_frame.paragraphs[0].font.size = Pt(9)
        table_2.table.cell(0, col).fill.solid()
        table_2.table.cell(0, col).fill.fore_color.rgb = header_colors[col % len(header_colors)]
        table_2.table.cell(0, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    table_2.table.rows[0].height = Inches(0.2)

    # Populate table data
    for row in range(1, rows + 1):
        table_2.table.rows[row].height = Inches(0.32)
        for col in range(cols):
            # table_2.table.cell(row,col).text_frame.height = Inches(0.3)
            _set_cell_border(table_2.table.cell(row, col), border_color="000000", border_width="12700")
            
            if col==0:
                table_2.table.columns[col].width = Inches(3.25)
                table_2.table.cell(row, col).text_frame.text = str(df.iloc[row - 1, col])
                table_2.table.cell(row, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            else:
                table_2.table.cell(row, col).text_frame.text = str(df.iloc[row - 1, col])
                table_2.table.columns[col].width = Inches(1)

            table_2.table.cell(row,col).text_frame.margin_bottom = Inches(0.07)
            table_2.table.cell(row,col).text_frame.margin_top = Inches(0.07)
            if len(str(df.iloc[row - 1, col]))>=35:            
                table_2.table.cell(row, col).text_frame.paragraphs[0].font.size = Pt(8)
            else:
                table_2.table.cell(row, col).text_frame.paragraphs[0].font.size = Pt(9)

            
            if col==2:
                table_2.table.cell(row,col).fill.solid()
                value = int(str(df.iloc[row - 1, col])[:-1])
                table_2.table.cell(row,col).fill.fore_color.rgb = RGBColor(df_new_colors_dict[value][0],df_new_colors_dict[value][1],df_new_colors_dict[value][2])
            elif col==0:
                table_2.table.cell(row, col).fill.solid()
                table_2.table.cell(row, col).fill.fore_color.rgb = RGBColor(51, 63, 79)
                table_2.table.cell(row, col).text_frame.paragraphs[0].font.color.rgb = RGBColor(255,255,255)
            else:
                table_2.table.cell(row, col).fill.solid()
                table_2.table.cell(row, col).fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            if col!=0:
                table_2.table.cell(row, col).text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            table_2.table.cell(row, 0).text_frame.paragraphs[0].font.bold = True
            # table

    return slide



def cond_formatting(df):

    norm_series = df['% Repeat Sales']
    norm_series = [float(str(x)[:-1]) for x in norm_series]

    # print(sorted_max)

    norm_colors = [
        (248/255, 105/255, 107/255),   # Red
        (255/255, 255/255, 255/255),   # White
        (83/255, 176/255, 108/255)  # Green
        # (176/255, 222/255, 188/255),  # light Green
        # (254/255, 241/255, 167/255),  # Yellow
        
    ]

    norm_cmap = LinearSegmentedColormap.from_list('green_white', norm_colors)

    # Normalize the numbers to a range between 0 and 1
    norm_norm = plt.Normalize(vmin=min(norm_series), vmax=max(norm_series))

    # Create a colormap for each number
    norm_colors_mapping = norm_cmap(norm_norm(norm_series))

    norm_rgb_colors = []
    for color in norm_colors_mapping:
        norm_rgb_colors.append(tuple(int(c * 255) for c in color))

    colors_dict = {}
    for i in range(len(norm_series)):
        colors_dict[int(norm_series[i])]=(norm_rgb_colors[i][0],norm_rgb_colors[i][1],norm_rgb_colors[i][2])
    
    return colors_dict

