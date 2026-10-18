import pandas as pd 
import numpy as np
from pptx.util import Pt,Inches
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from formatting_functions import format_with_locale
from variables import current_month

def edit_table(slide_41,df):
    for shape in slide_41.shapes:
        if shape.has_table:
            table = shape.table
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

                        if row > 0 and col>=0: # leaving headers values
                            # print(df.iloc[row-2,col])
                            new_run.text = str(df.iloc[row-1,col])

                        elif row == 0 and col>=0: # headers
                            # print(str(list(df.columns.values)[col]))
                            new_run.text = str(list(df.columns.values)[col])
                        
    return slide_41

def create_df(senco_ppt_xl_path,sheet_name):
    df = pd.read_excel(senco_ppt_xl_path,sheet_name = sheet_name)

    df = df[['points_band','total_customers','points_redeemed','points_issued','available_points']]
    df.rename(columns = {'points_band':'Points Band','total_customers':'Total customers','points_redeemed':'Points redeemed','points_issued':'Points issued','available_points':'Available points'},inplace = True)
    sums = df.select_dtypes(include='number').sum()
    grand_total_row = pd.DataFrame([["Grand Total", *sums]], columns=df.columns)
    df = pd.concat([df, grand_total_row], ignore_index=True)

    top_rows = df.iloc[:-1]  # Exclude the last row
    grand_total_row = df.iloc[-1:].apply(format_with_locale)      # Isolate the last row

    top_rows['% Redemption'] = (top_rows['Points redeemed']/top_rows['Points issued']*100).round(0).astype(int).astype(str) + '%'
    top_rows['% Available points'] = (top_rows['Available points']/top_rows['Available points'].sum()*100).round(1).astype(str) + '%'
    top_rows['Total customers'] = top_rows['Total customers'].apply(format_with_locale)
    top_rows['Points redeemed'] = top_rows['Points redeemed'].apply(format_with_locale)
    top_rows['Points issued'] = df['Points issued'].apply(format_with_locale)
    top_rows['Available points'] = top_rows['Available points'].apply(format_with_locale)
    df = pd.concat([top_rows, grand_total_row], ignore_index=True)
    df = df.fillna(' ')
    
    return df