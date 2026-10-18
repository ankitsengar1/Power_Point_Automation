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
import variables


def edit_table(slide_43,df_1,df_2):
        # #code to check index position of table
        # table_indices = []
        # for index, shape in enumerate(slide_42.shapes):
        #     if shape.has_table:
        #         table_indices.append(index)
        # print("######################################################")
        # print(table_indices)
        # print("#####################################################")

    for i in [2,3]:
        if slide_43.shapes[i].has_table:
            table = slide_43.shapes[i].table
            for col in range(0,len(table.columns)):
                for row in range(0,len(table.rows)):
                    for paragraph in table.cell(row,col).text_frame.paragraphs:
                        first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                        font = first_run.font
                        font_name = font.name
                        font_size = font.size
                        font_bold = font.bold
                        font_italic = font.italic
                        font_underline = font.underline
                        text = paragraph.text
                        if font.color.type == MSO_COLOR_TYPE.SCHEME:
                            # Get RGB from theme color
                            theme_color = font.color.theme_color
                            font_rgb_flag = 1
                        elif font.color.type == None:
                            font_rgb_flag = 0
                        elif font.color.type == MSO_COLOR_TYPE.PRESET:
                            color_index = font.color.theme_color
                            font_rgb_flag = 3
                        else:
                            # Already RGB color, use it directly
                            rgb = font.color.rgb
                            font_rgb_flag = 2

                        paragraph.clear()                    
                        new_run = paragraph.add_run()
                        new_run.font.name = font_name
                        new_run.font.size = font_size
                        new_run.font.bold = font_bold
                        new_run.font.underline = font_underline
                        new_run.font.italic = font_italic
                        # print(font_rgb_flag)
                        if font_rgb_flag == 1:
                            new_run.font.color.theme_color = theme_color
                        elif font_rgb_flag == 2:
                            new_run.font.color.rgb = rgb
                        elif font_rgb_flag == 3:
                            new_run.font.color.color_name = color_name = MSO_THEME_COLOR_INDEX(color_index).name

                        if i==2:
                            # print(row,col)
                            if row > 0 and col>=0: 
                                new_run.text = str(df_1.iloc[row-1,col])

                            elif row == 0 and col>=0: # headers
                                #print(str(list(df_win.columns.values)[col]))
                                new_run.text = str(list(df_1.columns.values)[col])
                
                        elif i==3:
                            if row >= 0 and col>=0: 
                                new_run.text =str(df_2.iloc[row,col])
                        
    return slide_43


def create_df(senco_ppt_xl_path,sheet_name1,sheet_name2):
   
    # df_1
    df_1 = pd.read_excel(senco_ppt_xl_path,sheet_name1)
    column_name = df_1.columns[1]
    var = column_name.split('(')[-1].strip(')')
    df_1.rename(columns={df_1.columns[0]:"Ageing of Points",df_1.columns[1]:f"Points Accrual (as on {var})"},inplace = True)
    df_1[f"% Points Accrual Distribution (as on {var})"] = (df_1[f"Points Accrual (as on {var})"]/df_1[f"Points Accrual (as on {var})"].sum()*100).round(1).astype(str) + '%'
    df_1[f"Points Accrual (as on {var})"] = df_1[f"Points Accrual (as on {var})"].apply(format_with_locale)
    new_index = {"0-180": "0-180 days", "181-365": "181 days-365 days", "366-730": "366 days-730 days","> 730":">=730 days"}
    df_1.loc[:,"Ageing of Points"] = df_1.loc[:,"Ageing of Points"].replace(new_index)

# #--------------------------------------------------------------------------------------------------------------------------------------------------------------
    #df_2
    df_2 = pd.read_excel(senco_ppt_xl_path,sheet_name2)
    new_index = {"0-30": "Points to be Lapsed in next 30 days", "30-60": "Points to be Lapsed in 30 - 60 days", "60-90": "Points to be Lapsed in 60 - 90 days"}
    df_2.loc[:,df_2.columns[0]] = df_2.loc[:,df_2.columns[0]].replace(new_index)
    df_2[df_2.columns[1]] = df_2[df_2.columns[1]].apply(format_with_locale)

    return df_1,df_2