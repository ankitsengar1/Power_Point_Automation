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


def edit_table(slide_44,df_1,df_2,df_3):
        #code to check index position of table
        # table_indices = []
        # for index, shape in enumerate(slide_43.shapes):
        #     if shape.has_table:
        #         table_indices.append(index)
        # print("######################################################")
        # print(table_indices)
        # print("#####################################################")
    for i in [0,1,2]:
        if slide_44.shapes[i].has_table:
            table = slide_44.shapes[i].table
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

                        if row==0 and col==0:
                            pass
                        else:
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

                            if i==0:
                                # print(row,col)
                                if row > 1 and col>=0: 
                                    new_run.text = str(df_1.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_win.columns.values)[col]))
                                    new_run.text = str(list(df_1.columns.values)[col])
                    
                            elif i==1:
                                if row > 1 and col>=0: 
                                    new_run.text =str(df_2.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_new.columns.values)[col]))
                                    new_run.text = str(list(df_2.columns.values)[col])

                            elif i==2:
                                if row > 1 and col>=0: 
                                    new_run.text =str(df_3.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_lap.columns.values)[col]))
                                    new_run.text = str(list(df_3.columns.values)[col])

                        
    return slide_44

def create_df(senco_ppt_xl_path,sheet_name1):

    df = pd.read_excel(senco_ppt_xl_path,sheet_name1)
    df = df.set_index(df.columns[0]).T
    df = df.reset_index()
    df.rename(columns={'index': 'KPIs'}, inplace=True)
    new_index = {"redemption_sales": "Revenue generated through points redemption","points_earned":"Points Earned","points_burned":"Points Burned"}
    df.loc[:,'KPIs'] = df.loc[:,'KPIs'].replace(new_index)

#---------------------------------------------------------------------------------------------------------------------------------------------------------------
    # df_1
    df_1 = df[[df.columns[0],df.columns[1],df.columns[2]]]
    df_1.rename(columns = {df_1.columns[1]:variables.ytd_previous_year,df_1.columns[2]:variables.ytd_current_year},inplace = True)
    df_1[df_1.columns[1]] = df_1[df_1.columns[1]].apply(format_with_locale)
    df_1[df_1.columns[2]] = df_1[df_1.columns[2]].apply(format_with_locale)

# #--------------------------------------------------------------------------------------------------------------------------------------------------------------
    #df_2
    df_2 = df[[df.columns[0],df.columns[3],df.columns[4]]]
    df_2.rename(columns = {df_2.columns[1]:variables.ytd_previous_year,df_2.columns[2]:variables.ytd_current_year},inplace = True)
    df_2[df_2.columns[1]] = df_2[df_2.columns[1]].apply(format_with_locale)
    df_2[df_2.columns[2]] = df_2[df_2.columns[2]].apply(format_with_locale)

# #---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    #df_3
    df_3 = df[[df.columns[0],df.columns[5],df.columns[6]]]
    df_3.rename(columns = {df_3.columns[1]:variables.ytd_previous_year,df_3.columns[2]:variables.ytd_current_year},inplace = True)
    df_3[df_3.columns[1]] = df_3[df_3.columns[1]].apply(format_with_locale)
    df_3[df_3.columns[2]] = df_3[df_3.columns[2]].apply(format_with_locale)

    return df_1,df_2,df_3