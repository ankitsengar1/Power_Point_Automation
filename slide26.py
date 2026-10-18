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

def edit_table(slide_26,df_win,df_new,df_lap):
    for i in [0,1,4]:
        if slide_26.shapes[i].has_table:
            table = slide_26.shapes[i].table
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
                                    new_run.text = str(df_win.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_win.columns.values)[col]))
                                    new_run.text = str(list(df_win.columns.values)[col])
                    
                            elif i==4:
                                if row > 1 and col>=0: 
                                    new_run.text =str(df_new.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_new.columns.values)[col]))
                                    new_run.text = str(list(df_new.columns.values)[col])

                            elif i==1:
                                if row > 1 and col>=0: 
                                    new_run.text =str(df_lap.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_lap.columns.values)[col]))
                                    new_run.text = str(list(df_lap.columns.values)[col])

                        
    return slide_26


def fetch_df(prev_slide_26):
    #code to check index position of table
    # table_indices = []
    # for index, shape in enumerate(prev_slide_20.shapes):
    #     if shape.has_table:
    #         table_indices.append(index)
    # print(table_indices)

    for i in [0,1,4]:
        if prev_slide_26.shapes[i].has_table:
            prev_table = prev_slide_26.shapes[i].table

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
            if i==0:
                df1 = pd.DataFrame(table_data, columns=headers)
                df1 = df1.drop(df1.columns[1], axis=1) #droping last column 

            elif i==1:
                df3 = pd.DataFrame(table_data, columns=headers)
                df3 = df3.drop(df3.columns[1], axis=1)#droping last column 

            elif i==4:
                df2 = pd.DataFrame(table_data, columns=headers)
                df2 = df2.drop(df2.columns[1], axis=1) #droping last column 
         
    return df1,df2,df3


def create_df(df,df1,df2,df3,senco_ppt_xl_path,sheet_name1):

    df_Winback_lapsed = pd.read_excel(senco_ppt_xl_path,sheet_name1)
    df_Winback_lapsed.rename(columns={df_Winback_lapsed.columns[0]:'KPIs',df_Winback_lapsed.columns[1]:variables.current_month,df_Winback_lapsed.columns[2]:variables.current_year_six_month,df_Winback_lapsed.columns[3]:variables.previous_year_six_month},inplace = True)
#---------------------------------------------------------------------------------------------------------------------------------------------------------------
    # New customer
    df_new = df[[df.columns[0],df.columns[1]]]
    df_new.rename(columns={df_new.columns[1]:variables.current_month},inplace  = True)
    overall_list = ['New Customers','New Customers Sales','New Customers Bills']
    df_new = df_new[df_new["KPIs"].isin(overall_list)]

    atv_new = df_new[df_new["KPIs"] == "New Customers Sales"].iloc[:, 1:].values / df_new[df_new["KPIs"] == "New Customers Bills"].iloc[:, 1:].values
    atv_new_row = pd.DataFrame([["ATV"] + atv_new.flatten().tolist()], columns=df_new.columns)

    amv_new = df_new[df_new["KPIs"] == "New Customers Sales"].iloc[:, 1:].values / df_new[df_new["KPIs"] == "New Customers"].iloc[:, 1:].values
    amv_new_row = pd.DataFrame([["AMV"] + amv_new.flatten().tolist()], columns=df_new.columns)

    # Append new rows to the DataFrame
    df_new = pd.concat([df_new, atv_new_row, amv_new_row], ignore_index=True)

    # Rename specific rows
    df_new['KPIs'] = df_new['KPIs'].replace({
        'New Customers Sales': 'Sales',
        'New Customers Bills': 'Bills'
    })

    df_new.loc[df_new['KPIs'].isin(['Sales']), variables.current_month] = df_new.loc[df_new['KPIs'].isin(['Sales']), variables.current_month].apply(
    lambda x: f"{round(x / 10000000)} Cr")

    df_new.loc[~df_new['KPIs'].isin(['Sales']), variables.current_month] = df_new.loc[~df_new['KPIs'].isin(['Sales']), variables.current_month].apply(format_with_locale)
    df_new = pd.merge(df2, df_new, on="KPIs")

#--------------------------------------------------------------------------------------------------------------------------------------------------------------
    # Winback customer

    df_win = df_Winback_lapsed[df_Winback_lapsed['KPIs'].isin(['winback_customers', 'winback_sales', 'winback_bills'])][['KPIs', variables.current_month]]

    atv_win = df_win[df_win["KPIs"] == "winback_sales"].iloc[:, 1:].values / df_win[df_win["KPIs"] == "winback_bills"].iloc[:, 1:].values
    atv_win_row = pd.DataFrame([["ATV"] + atv_win.flatten().tolist()], columns=df_win.columns)

    amv_win = df_win[df_win["KPIs"] == "winback_sales"].iloc[:, 1:].values / df_win[df_win["KPIs"] == "winback_customers"].iloc[:, 1:].values
    amv_win_row = pd.DataFrame([["AMV"] + amv_win.flatten().tolist()], columns=df_win.columns)

    # Append new rows to the DataFrame
    df_win = pd.concat([df_win, atv_win_row, amv_win_row], ignore_index=True)

    # Rename specific rows
    df_win['KPIs'] = df_win['KPIs'].replace({
        'winback_customers': 'Customers',
        'winback_sales':'Sales',
        'winback_bills': 'Bills'
    })

    df_win.loc[df_win['KPIs'].isin(['Sales']), variables.current_month] = df_win.loc[df_win['KPIs'].isin(['Sales']), variables.current_month].apply(
    lambda x: f"{round(x / 10000000)} Cr")

    df_win.loc[~df_win['KPIs'].isin(['Sales']), variables.current_month] = df_win.loc[~df_win['KPIs'].isin(['Sales']), variables.current_month].apply(format_with_locale)
    df_win = pd.merge(df1, df_win, on="KPIs")

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    # lapsed customer

    df_lap = df_Winback_lapsed[df_Winback_lapsed['KPIs'].isin(['lapsed_customers', 'lapsed_sales', 'lapsed_bills'])][['KPIs', variables.current_month]]

    atv_lap = df_lap[df_lap["KPIs"] == "lapsed_sales"].iloc[:, 1:].values / df_lap[df_lap["KPIs"] == "lapsed_bills"].iloc[:, 1:].values
    atv_lap_row = pd.DataFrame([["ATV"] + atv_lap.flatten().tolist()], columns=df_lap.columns)

    amv_lap = df_lap[df_lap["KPIs"] == "lapsed_sales"].iloc[:, 1:].values / df_lap[df_lap["KPIs"] == "lapsed_customers"].iloc[:, 1:].values
    amv_lap_row = pd.DataFrame([["AMV"] + amv_lap.flatten().tolist()], columns=df_lap.columns)

    # Append new rows to the DataFrame
    df_lap = pd.concat([df_lap, atv_lap_row, amv_lap_row], ignore_index=True)

    # Rename specific rows
    df_lap['KPIs'] = df_lap['KPIs'].replace({
        'lapsed_customers': 'Customers',
        'lapsed_sales':'Sales',
        'lapsed_bills': 'Bills'
    })

    df_lap.loc[df_lap['KPIs'].isin(['Sales']), variables.current_month] = df_lap.loc[df_lap['KPIs'].isin(['Sales']), variables.current_month].apply(
    lambda x: f"{round(x / 10000000)} Cr")

    df_lap.loc[~df_lap['KPIs'].isin(['Sales']), variables.current_month] = df_lap.loc[~df_lap['KPIs'].isin(['Sales']), variables.current_month].apply(format_with_locale)
    df_lap = pd.merge(df3, df_lap, on="KPIs")

    return df_win,df_new,df_lap
