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

def edit_table(slide_25,df1,df2):
    for i in [2,3]:
        if slide_25.shapes[i].has_table:
            table = slide_25.shapes[i].table
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
                
                        # if (col>0 and row==0) or (col==0 and row>0):
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

                        if i==2:
                            if row==0 and col==0:
                                new_run.text = f"For the month of {variables.current_month}"

                            # print(row,col)
                            elif row > 1 and col>=0: 
                                new_run.text = str(df1.iloc[row-2,col])

                            elif row == 1 and col>=0: # headers
                                # print(str(list(df.columns.values)[col]))
                                new_run.text = str(list(df1.columns.values)[col])
                
                        elif i==3:

                            if row > 0 and col>=0: 
                                new_run.text =str(df2.iloc[row-1,col])

                            elif row == 0 and col>=0: # headers
                                # print(str(list(df.columns.values)[col]))
                                new_run.text = str(list(df2.columns.values)[col])

                        
    return slide_25


def create_df(senco_ppt_xl_path,sheet_name,df):

    df_winback = pd.read_excel(senco_ppt_xl_path,sheet_name = sheet_name)
    df_winback.rename(columns={df_winback.columns[0]:'KPIs',df_winback.columns[1]:variables.current_month,df_winback.columns[2]:variables.ytd_current_year,df_winback.columns[3]:variables.ytd_previous_year},inplace = True)
    df_winback = df_winback[df_winback['KPIs'].isin(['winback_customers', 'winback_sales', 'winback_bills'])][['KPIs', variables.current_month,variables.ytd_current_year,variables.ytd_previous_year]]

    # creating first dataframe 
    df_win1 = df_winback[['KPIs',variables.current_month]]

    df_overall = df[['KPIs',df.columns[1]]]
    df_overall.rename(columns={df.columns[1]:variables.current_month},inplace  = True)
    df_overall = df_overall[df_overall["KPIs"].isin(['Total Sales','Total Bills'])]

    df1 = pd.concat([df_win1, df_overall], axis=0, ignore_index=True)

    atv = df1[df1["KPIs"] == "winback_sales"].iloc[:, 1:].values / df1[df1["KPIs"] == "winback_bills"].iloc[:, 1:].values
    atv_row = pd.DataFrame([["ATV"] + atv.flatten().tolist()], columns=df1.columns)

    amv = df1[df1["KPIs"] == "winback_sales"].iloc[:, 1:].values / df1[df1["KPIs"] == "winback_customers"].iloc[:, 1:].values
    amv_row = pd.DataFrame([["AMV"] + amv.flatten().tolist()], columns=df1.columns)

    o_atv = df1[df1["KPIs"] == "Total Sales"].iloc[:, 1:].values / df1[df1["KPIs"] == "Total Bills"].iloc[:, 1:].values
    o_atv_row = pd.DataFrame([["Overall ATV"] + o_atv.flatten().tolist()], columns=df1.columns)

    df1 = pd.concat([df1, atv_row, amv_row,o_atv_row], ignore_index=True)

    # Rename specific rows
    df1['KPIs'] = df1['KPIs'].replace({
        'winback_customers': 'Customers',
        'winback_sales':'Sales',
        'winback_bills': 'Bills'
    })

    df1 = df1[~df1["KPIs"].isin(['Total Sales','Total Bills'])]

    df1.loc[df1['KPIs'].isin(['Sales']), variables.current_month] = df1.loc[df1['KPIs'].isin(['Sales']), variables.current_month].apply(
    lambda x: f"{round(x / 10000000)} Cr")

    df1.loc[~df1['KPIs'].isin(['Sales']), variables.current_month] = df1.loc[~df1['KPIs'].isin(['Sales']), variables.current_month].apply(format_with_locale)
    df1.rename(columns={variables.current_month:'Win Back'},inplace = True)


    # creating second dataframe 
    df_win2 = df_winback[['KPIs',variables.ytd_current_year,variables.ytd_previous_year]]

    atv = df_win2[df_win2["KPIs"] == "winback_sales"].iloc[:, 1:].values / df_win2[df_win2["KPIs"] == "winback_bills"].iloc[:, 1:].values
    atv_row = pd.DataFrame([["Winback ATV"] + atv.flatten().tolist()], columns=df_win2.columns)

    df2 = pd.concat([df_win2, atv_row], ignore_index=True)
    df2['KPIs'] = df2['KPIs'].replace({
    'winback_customers': 'Winback Customers',
    'winback_sales':'Winback Sales',
    'winback_bills': 'Winback Bills'
    })

    df2['% Change from previous year'] = (df2[variables.ytd_current_year]-df2[variables.ytd_previous_year])/df2[variables.ytd_previous_year]*100
    df2['% Change from previous year'] = df2['% Change from previous year'].round(0).astype(int).astype(str) + '%'

    target_column = [variables.ytd_current_year,variables.ytd_previous_year]

    for col in target_column:
        df2.loc[df2['KPIs'].isin(['Winback Sales']), col] = df2.loc[df2['KPIs'].isin(['Winback Sales']), col].apply(
        lambda x: f"{round(x / 10000000)} Cr")

    for col in target_column:
        df2.loc[~df2['KPIs'].isin(['Winback Sales']), col] = df2.loc[~df2['KPIs'].isin(['Winback Sales']), col].apply(format_with_locale)

    return df1,df2