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
from variables import current_month

def edit_table(slide_14,df,loyalty_sales_colors_dict,loyalty_bills_colors_dict):
    for shape in slide_14.shapes:
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
                
                        if (col>0 and row==0) or (col==0 and row>0):
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
 
                        
                            if row > 1 and col>0: # leaving headers and first column values
                                # print(df.iloc[row-2,col])
                                new_run.text = str(df.iloc[row-2,col])

                                if col==3:
                                    table.cell(row,col).fill.solid()
                                    value = float(str(df.iloc[row-2,col])[:-1])
                                    #print(loyalty_sales_colors_dict[value][0],loyalty_sales_colors_dict[value][1],loyalty_sales_colors_dict[value][2])
                                    table.cell(row,col).fill.fore_color.rgb = RGBColor(loyalty_sales_colors_dict[value][0],loyalty_sales_colors_dict[value][1],loyalty_sales_colors_dict[value][2])
                                
                                if col==4:
                                    table.cell(row,col).fill.solid()
                                    value = float(str(df.iloc[row-2,col])[:-1])
                                    #print(loyalty_bills_colors_dict[value][0],loyalty_bills_colors_dict[value][1],loyalty_bills_colors_dict[value][2])
                                    table.cell(row,col).fill.fore_color.rgb = RGBColor(loyalty_bills_colors_dict[value][0],loyalty_bills_colors_dict[value][1],loyalty_bills_colors_dict[value][2]) 

                            elif row == 1 and col>0: # headers
                                # print(str(list(df.columns.values)[col]))
                                new_run.text = str(list(df.columns.values)[col])
                            elif row==0 and col==0:
                                # print(f"Duration : Till {curr_month_curr_year}")
                                new_run.text = f"For the month of {current_month}"

                        
    return slide_14

def create_df(senco_ppt_xl_path,sheet_name):
    df = pd.read_excel(senco_ppt_xl_path,sheet_name = sheet_name)

    df = df[['range','loyaltysales_billband','loyaltybills_billband']]
    df.rename(columns={'range':'Bill Band','loyaltysales_billband':'Loyalty Sales','loyaltybills_billband':'Loyalty Bills'},inplace = True)
    df['% Loyalty Sales'] = (df['Loyalty Sales']/df['Loyalty Sales'].sum()*100).round(1).astype(str) + '%'
    df['% Loyalty Bills'] = (df['Loyalty Bills']/df['Loyalty Bills'].sum()*100).round(1).astype(str) + '%'
    df['Loyalty Sales'] = df['Loyalty Sales'].apply(format_with_locale)
    df['Loyalty Bills'] = df['Loyalty Bills'].apply(format_with_locale)

    return df