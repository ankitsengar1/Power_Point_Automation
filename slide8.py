import pandas as pd
import numpy as np 
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from formatting_functions import format_with_locale
import variables

def edit_table(slide_8,df_1,df_2,df_3):
        # #code to check index position of table
        # table_indices = []
        # for index, shape in enumerate(slide_8.shapes):
        #     if shape.has_table:
        #         table_indices.append(index)
        # print("######################################################")
        # print(table_indices)
        # print("#####################################################")

    for i in [0,2,3]:
        if slide_8.shapes[i].has_table:
            table = slide_8.shapes[i].table
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

                        # if row==0:
                        #     pass
                        # else:
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
                            if row > 0 and col>=0: 
                                new_run.text = str(df_1.iloc[row-1,col])

                            elif row == 0 and col>=0: # headers
                                #print(str(list(df_1.columns.values)[col]))
                                new_run.text = str(list(df_1.columns.values)[col])
                
                        elif i==2:
                            if row > 0 and col>=0: 
                                new_run.text =str(df_2.iloc[row-1,col])

                            elif row == 0 and col>=0: # headers
                                #print(str(list(df_2.columns.values)[col]))
                                new_run.text = str(list(df_2.columns.values)[col])
                                

                        elif i==3:
                            if row > 0 and col>=0: 
                                new_run.text =str(df_3.iloc[row-1,col])

                            elif row == 0 and col>=0: # headers
                                #print(str(list(df_3.columns.values)[col]))
                                new_run.text = str(list(df_3.columns.values)[col])                            
    return slide_8


def create_df(df):
    df.rename(columns={df.columns[1]:variables.current_month,df.columns[2]:variables.current_year_prev_month1,
                       df.columns[3]:variables.current_year_prev_month2,df.columns[4]:variables.current_month_prev_year},inplace=True)
    target_columns = [variables.current_month, variables.current_year_prev_month1, variables.current_year_prev_month2, variables.current_month_prev_year]
    
    #---------------------------------------first table---------------------------------------------------------------------

    df_sales = df[df["KPIs"].isin(["Total Sales","Repeater Sales","New Customers Sales"])]

    per_repeat_sales  = df_sales[df_sales["KPIs"] == "Repeater Sales"].iloc[:, 1:].values / df_sales[df_sales["KPIs"] == "Total Sales"].iloc[:, 1:].values*100
    per_repeat_sales_row = pd.DataFrame([["% Repeater Sales"] + per_repeat_sales.flatten().tolist()], columns=df_sales.columns)

    per_new_sales  = df_sales[df_sales["KPIs"] == "New Customers Sales"].iloc[:, 1:].values / df_sales[df_sales["KPIs"] == "Total Sales"].iloc[:, 1:].values*100
    per_new_sales_row = pd.DataFrame([["% New Customers Sales"] + per_new_sales.flatten().tolist()], columns=df_sales.columns)

    df_sales = pd.concat([df_sales, per_repeat_sales_row, per_new_sales_row], ignore_index=True)
    df_sales = df_sales[~df_sales["KPIs"].isin(["Repeater Sales","New Customers Sales"])]

    # Rows and columns to apply the formatting
    target_rows1 = ["Total Sales"]

    for col in target_columns:
        df_sales.loc[df_sales['KPIs'].isin(target_rows1), col] = df_sales.loc[df_sales['KPIs'].isin(target_rows1), col].apply(
        lambda x: f"{round(x / 10000000)} Cr")

    for col in target_columns:
        df_sales.loc[~df_sales['KPIs'].isin(target_rows1), col] = df_sales.loc[~df_sales['KPIs'].isin(target_rows1), col].apply(lambda x: f"{int(round(x))}%")
        
    #-----------------------------------------second table ---------------------------------------------------------------

    df_atv = df[df["KPIs"].isin(["ATV","Repeater Sales","New Customers Sales","Repeater Bills","New Customers Bills"])]

    df_atv['KPIs'] = df_atv['KPIs'].replace({"ATV":"Overall ATV"})
    per_repeat_atv  = df_atv[df_atv["KPIs"] == "Repeater Sales"].iloc[:, 1:].values / df_atv[df_atv["KPIs"] == "Repeater Bills"].iloc[:, 1:].values
    per_repeat_atv_row = pd.DataFrame([["Repeaters ATV"] + per_repeat_atv.flatten().tolist()], columns=df_atv.columns)

    per_new_atv  = df_atv[df_atv["KPIs"] == "New Customers Sales"].iloc[:, 1:].values / df_atv[df_atv["KPIs"] == "New Customers Bills"].iloc[:, 1:].values
    per_new_atv_row = pd.DataFrame([["New Customers ATV"] + per_new_atv.flatten().tolist()], columns=df_atv.columns)

    df_atv = pd.concat([df_atv, per_repeat_atv_row, per_new_atv_row], ignore_index=True)
    df_atv = df_atv[~df_atv["KPIs"].isin(["Repeater Bills","New Customers Bills","Repeater Sales","New Customers Sales"])]

    # # Rows and columns to apply the formatting
    target_rows2 = ["Overall ATV","Repeaters ATV","New Customers ATV"]

    for col in target_columns:
        df_atv.loc[df_atv['KPIs'].isin(target_rows2), col] = df_atv.loc[df_atv['KPIs'].isin(target_rows2), col].apply(format_with_locale)
    
    #-------------------------------------------Third table----------------------------------------------------------------

    df_sales_gdp = df[df["KPIs"].isin(["Total Sales","gold_sales","diamond_sales","platinum_sales"])]

    per_repeat_sales_g  = df_sales_gdp[df_sales_gdp["KPIs"] == "gold_sales"].iloc[:, 1:].values / df_sales_gdp[df_sales_gdp["KPIs"] == "Total Sales"].iloc[:, 1:].values*100
    per_repeat_sales_g_row = pd.DataFrame([["% Gold Sales"] + per_repeat_sales_g.flatten().tolist()], columns=df_sales_gdp.columns)

    per_new_sales_d  = df_sales_gdp[df_sales_gdp["KPIs"] == "diamond_sales"].iloc[:, 1:].values / df_sales_gdp[df_sales_gdp["KPIs"] == "Total Sales"].iloc[:, 1:].values*100
    per_new_sales_d_row = pd.DataFrame([["% Diamond Sales"] + per_new_sales_d.flatten().tolist()], columns=df_sales_gdp.columns)

    per_new_sales_p  = df_sales_gdp[df_sales_gdp["KPIs"] == "platinum_sales"].iloc[:, 1:].values / df_sales_gdp[df_sales_gdp["KPIs"] == "Total Sales"].iloc[:, 1:].values*100
    per_new_sales_p_row = pd.DataFrame([["% Platinum Sales"] + per_new_sales_p.flatten().tolist()], columns=df_sales_gdp.columns)

    df_sales_gdp = pd.concat([df_sales_gdp, per_repeat_sales_g_row, per_new_sales_d_row, per_new_sales_p_row], ignore_index=True)
    df_sales_gdp = df_sales_gdp[~df_sales_gdp["KPIs"].isin(["gold_sales","diamond_sales","platinum_sales"])]

    target_rows3 = ["Total Sales"]

    for col in target_columns:
        df_sales_gdp.loc[df_sales_gdp['KPIs'].isin(target_rows3), col] = df_sales_gdp.loc[df_sales_gdp['KPIs'].isin(target_rows3), col].apply(
        lambda x: f"{round(x / 10000000)} Cr")

    for col in target_columns:
        df_sales_gdp.loc[~df_sales_gdp['KPIs'].isin(target_rows3), col] = df_sales_gdp.loc[~df_sales_gdp['KPIs'].isin(target_rows3), col].apply(lambda x: f"{int(round(x))}%")

    return df_sales,df_atv,df_sales_gdp