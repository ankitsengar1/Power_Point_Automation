import pandas as pd
import numpy as np 
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap,Normalize
import matplotlib.colors as mcolors
import formatting_functions
from formatting_functions import format_with_locale
import variables
from matplotlib.patches import Rectangle

# def edit_table(slide_7,df,previous_month_colors_dict,previous_year_colors_dict):
#     for shape in slide_7.shapes:
#         if shape.has_table:
#             table = shape.table
#             for col in range(0,len(table.columns)):
#                 for row in range(0,len(table.rows)):
#                     for paragraph in table.cell(row,col).text_frame.paragraphs:
#                         first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
#                         font = first_run.font
#                         font_name = font.name
#                         font_size = font.size
#                         font_bold = font.bold
#                         font_italic = font.italic
#                         font_underline = font.underline
#                         text = paragraph.text
#                         if font.color.type == MSO_COLOR_TYPE.SCHEME:
#                             # Get RGB from theme color
#                             theme_color = font.color.theme_color
#                             font_rgb_flag = 1
#                         elif font.color.type == None:
#                             font_rgb_flag = 0
#                         elif font.color.type == MSO_COLOR_TYPE.PRESET:
#                             color_index = font.color.theme_color
#                             font_rgb_flag = 3
#                         else:
#                             # Already RGB color, use it directly
#                             rgb = font.color.rgb
#                             font_rgb_flag = 2
                
#                         paragraph.clear()                    
#                         new_run = paragraph.add_run()
#                         new_run.font.name = font_name
#                         new_run.font.size = font_size
#                         new_run.font.bold = font_bold
#                         new_run.font.underline = font_underline
#                         new_run.font.italic = font_italic
#                         # print(font_rgb_flag)
#                         if font_rgb_flag == 1:
#                             new_run.font.color.theme_color = theme_color
#                         elif font_rgb_flag == 2:
#                             new_run.font.color.rgb = rgb
#                         elif font_rgb_flag == 3:
#                             new_run.font.color.color_name = color_name = MSO_THEME_COLOR_INDEX(color_index).name

                    
#                         if row > 0 and col>=0: # leaving headers and first column values
#                             # print(df.iloc[row-2,col])
#                             new_run.text = str(df.iloc[row-1,col])

#                             if col==5:
#                                 table.cell(row,col).fill.solid()
#                                 value = float(str(df.iloc[row-1,col])[:-1])
#                                 #print(loyalty_sales_colors_dict[value][0],loyalty_sales_colors_dict[value][1],loyalty_sales_colors_dict[value][2])
#                                 table.cell(row,col).fill.fore_color.rgb = RGBColor(previous_month_colors_dict[value][0],previous_month_colors_dict[value][1],previous_month_colors_dict[value][2])

#                             elif col==6:
#                                 table.cell(row,col).fill.solid()
#                                 value = float(str(df.iloc[row-1,col])[:-1])
#                                 #print(loyalty_sales_colors_dict[value][0],loyalty_sales_colors_dict[value][1],loyalty_sales_colors_dict[value][2])
#                                 table.cell(row,col).fill.fore_color.rgb = RGBColor(previous_year_colors_dict[value][0],previous_year_colors_dict[value][1],previous_year_colors_dict[value][2])

#                         elif row == 0 and col>=0: # headers
#                             # print(str(list(df.columns.values)[col]))
#                             new_run.text = str(list(df.columns.values)[col])
                        
#     return slide_7

def create_df(df,slide_7):

    df.rename(columns={df.columns[1]:variables.current_month,df.columns[2]:variables.current_year_prev_month1,df.columns[3]:variables.current_year_prev_month2,df.columns[4]:variables.current_month_prev_year},inplace=True)
    df = df[~df["KPIs"].isin(["gold_sales","diamond_sales","platinum_sales"])]
    df['% Change from previous month'] = (df[variables.current_month]-df[variables.current_year_prev_month1])/df[variables.current_year_prev_month1]*100
    df['% Change from previous year'] = (df[variables.current_month]-df[variables.current_month_prev_year])/df[variables.current_month_prev_year]*100
    df['% Change from previous month'] = df['% Change from previous month'].round(0).astype(int).astype(str) + '%'
    df['% Change from previous year'] = df['% Change from previous year'].round(0).astype(int).astype(str) + '%'

    exceptions = ["Total Sales", "Repeater Sales", "New Customers Sales", 
                "Visits Per Customer", "Bills Per Customer", "ABS"]
    
    target_kpis_inr_cr = ["Total Sales", "Repeater Sales", "New Customers Sales"]
    targetkpis_roundup = ["Visits Per Customer", "Bills Per Customer", "ABS"]

    target_columns = [variables.current_month, variables.current_year_prev_month1, variables.current_year_prev_month2, variables.current_month_prev_year]

    for col in target_columns:
        df.loc[df['KPIs'].isin(targetkpis_roundup), col] = df.loc[df['KPIs'].isin(targetkpis_roundup), col].round(2)

    for col in target_columns:
        df.loc[~df['KPIs'].isin(exceptions), col] = df.loc[~df['KPIs'].isin(exceptions), col].apply(format_with_locale)

    for col in target_columns:
        df.loc[df['KPIs'].isin(target_kpis_inr_cr), col] = df.loc[df['KPIs'].isin(target_kpis_inr_cr), col].apply(
            lambda x: f"INR {round(x / 10000000, 1)} Cr"
        )
###########################################################################################################################################################
# formatting code 

    # Convert '% Change' columns to floats for progress bar calculations
    percent_change_columns = ['% Change from previous month', '% Change from previous year']
    for col in percent_change_columns:
        df[col] = df[col].str.replace('%', '').astype(float)

    # calulating shift and actual value
    shifts = {}
    actual_val = {}
    for col in percent_change_columns:
        positive_max = df[df[col] > 0][col].max() if (df[col] > 0).any() else 0
        negative_max = abs(df[df[col] < 0][col].min()) if (df[col] < 0).any() else 0
        
        if positive_max > 0 and negative_max == 0: 
            shifts[col] = -0.7
            actual_val[col] = positive_max
        elif negative_max > 0 and positive_max == 0: 
            shifts[col] = 0.7
            actual_val[col] = negative_max
        else:  
            shifts[col] = -0.7 if positive_max > negative_max else 0.7
            actual_val[col] = max(positive_max,negative_max)

    # Set up figure and axis
    fig, ax = plt.subplots(figsize=(21, 9))
    ax.axis('off')

    # Table parameters
    n_rows, n_cols = df.shape
    cell_height = 0.4
    cell_widths = [2.9, 1.5, 1.5, 1.5, 1.5, 3.6, 3.6]

    # Colors
    header_color = '#333f4f'
    header_text_color = 'white'
    kpi_color = '#444f60'
    kpi_text_color = 'white'
    row_bg_color = 'white'

    #Helper function to draw progress bar starting from left to right
    # def draw_progress_bar(ax, x, y, width, height, value):
    #     """
    #     Draw a progress bar in a cell.
    #     """
    #     bar_width = (width - 0.1) * abs(value) / 100
    #     bar_color = '#f8696b' if value < 0 else '#7cc28f'  # Red for negative, green for positive
    #     ax.add_patch(Rectangle((x + 0.05, y + 0.05), bar_width, height - 0.1, color=bar_color))


    def draw_progress_bar(ax, x, y, width, height, value, shift,actual_val):
        
        if actual_val<=15:
            center_x = x + width / 2 + shift  # Adjust center based on shift
            bar_width = (2*width - 0.1) * abs(value) / 100  # Scale the width for % value
         
        if (actual_val>15) & (actual_val<=70):
            center_x = x + width / 2 + shift 
            bar_width = (width - 0.1) * abs(value) / 100 

        elif (actual_val>70)&(actual_val<=100):
            center_x = x + width / 2 + shift 
            bar_width = (width/1.5 - 0.1) * abs(value) / 100 

        elif (actual_val>100)&(actual_val<=140):
            center_x = x + width / 2 + shift  
            bar_width = (width/2 - 0.1) * abs(value) / 100 

        elif (actual_val>140):
            center_x = x + width / 2 + shift 
            bar_width = (width/3 - 0.1) * abs(value) / 100 
        
        bar_color = '#f2797b' if value < 0 else '#a3e3b2'  # Red for negative, green for positive
        
        # If value is negative, start drawing from center to the left
        if value < 0:
            ax.add_patch(Rectangle((center_x - bar_width, y + 0.05), bar_width, height - 0.1, color=bar_color))
        # If value is positive, start drawing from center to the right
        elif value > 0:
            ax.add_patch(Rectangle((center_x, y + 0.05), bar_width, height - 0.1, color=bar_color))

    x_start = 0
    y_start = n_rows * cell_height

    # Draw header row
    for col_idx in range(n_cols):
        x = sum(cell_widths[:col_idx])
        y = y_start
        ax.add_patch(Rectangle((x, y), cell_widths[col_idx], cell_height, facecolor=header_color, edgecolor='black'))
        ax.text(x + cell_widths[col_idx] / 2, y + cell_height / 2, df.columns[col_idx],
                ha='center', va='center', color=header_text_color, fontsize=18, weight='bold')

    # Draw body rows
    for row_idx in range(n_rows):
        y = y_start - (row_idx + 1) * cell_height
        # Alternate row background colors
        row_color = row_bg_color if row_idx % 2 == 0 else 'white'
        
        for col_idx in range(n_cols):
            x = sum(cell_widths[:col_idx])
            # Set KPI column background and text styling
            if col_idx == 0:
                cell_color = kpi_color
                text_color = kpi_text_color
                font_weight = 'bold'
            else:
                cell_color = row_color
                text_color = 'black'
                font_weight = 'normal'
            
            # Draw cell rectangle
            ax.add_patch(Rectangle((x, y), cell_widths[col_idx], cell_height, facecolor=cell_color, edgecolor='black'))
            
            # Add text with % sign for the 5th and 6th columns
            value = df.iloc[row_idx, col_idx]
            if col_idx in [5, 6]:  # Columns 5 and 6 need the % sign
                formatted_value = f"{value:.0f}%"  # Add % sign
            else:
                formatted_value = f'{value:.2f}' if isinstance(value, (int, float)) else value

            ax.text(x + cell_widths[col_idx] / 2, y + cell_height / 2, formatted_value,
                    ha='center', va='center', color=text_color, fontsize=18, weight=font_weight)
            
            # Draw progress bar for '% Change' columns
            if col_idx in [5, 6]:  # Progress bar for % Change columns
                col_name = percent_change_columns[col_idx - 5]
                draw_progress_bar(ax, x, y, cell_widths[col_idx], cell_height, value, shifts[col_name],actual_val[col_name])


    # Draw outer border around the entire table
    x_outer = 0
    y_outer = 0
    for col_idx in range(n_cols):
        x = sum(cell_widths[:col_idx])
        ax.add_patch(Rectangle((x, 0), cell_widths[col_idx], y_start + cell_height, edgecolor='black', facecolor='none', linewidth=2))  # Top and bottom border
    for row_idx in range(n_rows + 1):
        y = y_start - row_idx * cell_height
        ax.add_patch(Rectangle((0, y), sum(cell_widths), cell_height, edgecolor='black', facecolor='none', linewidth=2))  # Left and right border

    # Set axis limits
    plt.xlim(0, sum(cell_widths))
    plt.ylim(0, y_start + cell_height)
    plt.tight_layout()

    # Save and show the table
    plt_path = variables.img_path
    plt.savefig(plt_path)

    left = Inches(0.24)
    top = Inches(0.93)
    width = Inches(12.85)
    height = Inches(6.2)
    pic = slide_7.shapes.add_picture(plt_path, left, top, width, height)

    return slide_7

