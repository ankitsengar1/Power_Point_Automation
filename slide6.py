import pandas as pd
import numpy as np
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
from pptx.dml.color import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from formatting_functions import format_with_locale


def edit_table(slide_6,df_1,df_2):
        ##code to check index position of table
        # table_indices = []
        # for index, shape in enumerate(slide_6.shapes):
        #     if shape.has_table:
        #         table_indices.append(index)
        # print("######################################################")
        # print(table_indices)
        # print("#####################################################")

    for i in [0,1]:
        if slide_6.shapes[i].has_table:
            table = slide_6.shapes[i].table
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

                        if row==0:
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

                            if i==1:
                                # print(row,col)
                                if row > 1 and col>=0: 
                                    new_run.text = str(df_1.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_win.columns.values)[col]))
                                    new_run.text = str(list(df_1.columns.values)[col])
                    
                            elif i==0:
                                if row > 1 and col>=0: 
                                    new_run.text =str(df_2.iloc[row-2,col])

                                elif row == 1 and col>=0: # headers
                                    #print(str(list(df_win.columns.values)[col]))
                                    new_run.text = str(list(df_2.columns.values)[col])
    return slide_6


def create_df(senco_ppt_xl_path,sheet_name):
    df = pd.read_excel(senco_ppt_xl_path,sheet_name = sheet_name)

    #df['blank'] = np.nan # adding extra column
    df = df.drop(df.index[-1]) # droping last row 
    new_columns = [df.columns[0]] + list(range(len(df.columns) - 1)) # Rename columns from index 1 onward to numbers starting from 0
    df.columns = new_columns
    df.iloc[:, 1:] = df.iloc[:, 1:].apply(lambda row: pd.Series([x for x in row if not pd.isna(x)] + [np.nan] * row.isna().sum()), axis=1)
    df.rename(columns={'row labels':'Start Year'}, inplace = True)
    df.insert(0, 'Year', range(len(df)))
    df = df.reset_index(drop=True)
    df = df.fillna(0) 

    # first table 
    df_1 = df.copy()
    df_1 = df_1.applymap(lambda x: int(x) if isinstance(x, float) else x)
    df_1.iloc[:, 2:] = df_1.iloc[:, 2:].applymap(lambda x: '-' if x == 0 else x)
    df_1.iloc[:, 2:] = df_1.iloc[:, 2:].applymap(format_with_locale)

    # second table 
    df_2 = df.copy()
    for index, row in df_2.iterrows():
        base_value = row[0]  # The value from column 0 (index 0)
        
        # Apply the operation for columns starting from index 2 onward
        df_2.iloc[index, 2:] = row[2:] / base_value * 100
        
    df_2.iloc[:, 2:] = df_2.iloc[:, 2:].applymap(lambda x: round(x, 0) if isinstance(x, float) else x)
    df_2 = df_2.applymap(lambda x: int(x) if isinstance(x, float) else x)
    df_2.iloc[:, 2:] = df_2.iloc[:, 2:].applymap(lambda x: str(x) + '%' if isinstance(x, (int, float)) else x)
    df_2.iloc[:, 2:] = df_2.iloc[:, 2:].applymap(lambda x: '-' if x == '0%' else x)
    return df_1,df_2
