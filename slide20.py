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

def edit_table(slide_20,df1,df2,loyalty_sales_colors_dict1,loyalty_bills_colors_dict1,loyalty_sales_colors_dict2,loyalty_bills_colors_dict2):
    for i in [2,3]:
        if slide_20.shapes[i].has_table:
            table = slide_20.shapes[i].table
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
                
                        if (col > 0 and row == 0) or (col >= 0 and row == 1):
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
 
                            if i==2:                                                          

                                new_run.text = str(df2.iloc[row-2,col])

                                if col==3 and row > 1:
                                    table.cell(row,col).fill.solid()
                                    value = float(str(df2.iloc[row-2,col])[:-1])
                                    #print(loyalty_sales_colors_dict[value][0],loyalty_sales_colors_dict[value][1],loyalty_sales_colors_dict[value][2])
                                    table.cell(row,col).fill.fore_color.rgb = RGBColor(loyalty_sales_colors_dict2[value][0],loyalty_sales_colors_dict2[value][1],loyalty_sales_colors_dict2[value][2])
                                
                                if col==4 and row > 1:
                                    table.cell(row,col).fill.solid()
                                    value = float(str(df2.iloc[row-2,col])[:-1])
                                    #print(loyalty_bills_colors_dict[value][0],loyalty_bills_colors_dict[value][1],loyalty_bills_colors_dict[value][2])
                                    table.cell(row,col).fill.fore_color.rgb = RGBColor(loyalty_bills_colors_dict2[value][0],loyalty_bills_colors_dict2[value][1],loyalty_bills_colors_dict2[value][2]) 

                                elif row==0 and col==0:
                                    # print(f"Duration : Till {curr_month_curr_year}")
                                    new_run.text = f"For the month of {variables.current_year_prev_month1}"

                            elif i==3:

                                new_run.text =str(df1.iloc[row-2,col])

                                if col==3 and row > 1:
                                    table.cell(row,col).fill.solid()
                                    value = float(str(df1.iloc[row-2,col])[:-1])
                                    #print(value)
                                    #print(total_sales_colors_dict[value][0],total_sales_colors_dict[value][1],total_sales_colors_dict[value][2])
                                    table.cell(row,col).fill.fore_color.rgb = RGBColor(loyalty_sales_colors_dict1[value][0],loyalty_sales_colors_dict1[value][1],loyalty_sales_colors_dict1[value][2])
                                    
                                if col==4 and row >1:
                                    table.cell(row,col).fill.solid()
                                    value = float(str(df1.iloc[row-2,col])[:-1])
                                    #print(value)
                                    #print(total_bills_colors_dict[value][0],total_bills_colors_dict[value][1],total_bills_colors_dict[value][2])
                                    table.cell(row,col).fill.fore_color.rgb = RGBColor(loyalty_bills_colors_dict1[value][0],loyalty_bills_colors_dict1[value][1],loyalty_bills_colors_dict1[value][2]) 

                                elif row==0 and col==0:
                                    #print(f"Duration : Till {variables.current_year_prev_month1}")
                                    new_run.text = f"For the month of {variables.current_year_prev_month1}"
                        
    return slide_20

def create_df(prev_slide_20):
    for i in [2,3]:
        if prev_slide_20.shapes[i].has_table:
            prev_table = prev_slide_20.shapes[i].table

            table_data = []
            # Loop through the table and append data, skipping the first row and first column
            for row_idx in range(2, len(prev_table.rows)):  # Start from the second row (index 1)
                row_data = []
                for col_idx in range(0, len(prev_table.columns)):  # Start from the second column (index 1)
                    cell_text = prev_table.cell(row_idx, col_idx).text_frame.paragraphs[0].text
                    row_data.append(cell_text)
                table_data.append(row_data)

            # Extract the header from the second row (index 1)
            headers = [
                prev_table.cell(1, col_idx).text_frame.paragraphs[0].text
                for col_idx in range(0, len(prev_table.columns))
            ]
            if i==2:
                df1 = pd.DataFrame(table_data, columns=headers)

            elif i==3:
                df2 = pd.DataFrame(table_data, columns=headers)

    return df1,df2