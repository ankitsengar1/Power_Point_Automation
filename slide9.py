import pandas as pd
import numpy as np
from pptx.util import Pt,Inches
import matplotlib.pyplot as plt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
import warnings
from matplotlib.ticker import PercentFormatter, FuncFormatter
from matplotlib import font_manager
from formatting_functions import format_with_locale
import variables
warnings.simplefilter(action='ignore', category=FutureWarning)

def edit_table(slide_9,df):
    for shape in slide_9.shapes:
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

                        if row > 0 and col>=0: # leaving headers and first column values
                            # print(df.iloc[row-2,col])
                            new_run.text = str(df.iloc[row-1,col])

                        elif row == 0 and col>=0: # headers
                            # print(str(list(df.columns.values)[col]))
                            new_run.text = str(list(df.columns.values)[col])
                        
    return slide_9

def create_df(senco_ppt_xl_path,sheet_name):

    df = pd.read_excel(senco_ppt_xl_path,sheet_name = sheet_name)
    df['mom'] = pd.to_datetime(df['mom']).dt.strftime('%b-%y')
    df.rename(columns ={"mom":"Months"},inplace = True)

    # calculations for table
    df['% Repeat Sales'] = (df['other_month_repeater_sales']/df['loyalty_sales']*100).round(0).astype(int).astype(str) + '%'
    df['% New Customers Sales'] = ((df['current_month_onetimer_sales']+df['current_month_repeater_sales'])/df['loyalty_sales']*100).round(0).astype(int).astype(str) + '%'
    df['Repeaters ATV'] = (df['other_month_repeater_sales']/df['other_month_repeater_bills']).apply(format_with_locale)
    df['New Customers ATV'] = ((df['current_month_onetimer_sales']+df['current_month_repeater_sales'])/(df['current_month_onetimer_bills']+df['current_month_repeater_bills'])).apply(format_with_locale)
    df = df[['Months','% Repeat Sales','% New Customers Sales','New Customers ATV','Repeaters ATV']]

    df_table_i = df.set_index('Months').T
    df_table = df_table_i.reset_index()
    df_table.rename(columns={'index': 'Months'}, inplace=True)

    df_graph = df_table_i.apply(lambda x: x.str.replace('%', '').str.replace(',', '')).apply(pd.to_numeric)
    df_graph = df_graph.T.reset_index()

    return df_table,df_graph

def generate_graph(df_graph,slide9):

    # Create the plot with the specified background color
    fig, ax1 = plt.subplots(figsize=(14, 5))
    fig.patch.set_facecolor("#333F4F")  # Set the figure background color
    ax1.set_facecolor("#333F4F")        # Set the axes background color

    # Plot stacked bar chart for % New Customers Sales and % Repeat Sales on ax1
    ax1.bar(df_graph["Months"], df_graph["% New Customers Sales"], color="#AEAEAE", label="% New Customers Sales", width=0.3)
    ax1.bar(df_graph["Months"], df_graph["% Repeat Sales"], bottom=df_graph["% New Customers Sales"], color="#FFC54B", label="% Repeat Sales", width=0.3)

    # Line plot for New Customers ATV and Repeaters ATV on ax2 (secondary axis)
    ax2 = ax1.twinx()
    ax2.set_facecolor("#333F4F")  # Set the secondary y-axis background color
    ax2.plot(df_graph["Months"], df_graph["New Customers ATV"], color="#4472C4", marker='o', label="New Customers ATV", linewidth=3)
    ax2.plot(df_graph["Months"], df_graph["Repeaters ATV"], color="#ED7D31", marker='o', label="Repeaters ATV", linewidth=3)

    # Swap the y-axis positions
    ax2.yaxis.set_ticks_position('left')
    ax1.yaxis.set_ticks_position('right')

    # Set axis labels and title with custom text color and adjusted title padding
    ax1.set_title("New Customers vs Repeaters", color="#FFFFFF", fontweight='bold', fontsize=16, pad=20)

    # Customize tick labels color
    ax1.tick_params(axis='x', colors="#FFFFFF")
    ax1.tick_params(axis='y', colors="#FFFFFF")
    ax2.tick_params(axis='y', colors="#FFFFFF")

    # Apply bold styling to tick labels
    for label in ax1.get_xticklabels():
        label.set_fontsize(12)
    for label in ax1.get_yticklabels():
        label.set_fontsize(12)
    for label in ax2.get_yticklabels():
        label.set_fontsize(12)

    font_properties = font_manager.FontProperties(size=12)

    # Display legends with custom text color and background
    ax1.legend(loc="lower right", facecolor="#333F4F", ncol=2, framealpha=0.5, labelcolor="#FFFFFF", bbox_to_anchor=(0.5, -0.20), prop=font_properties, frameon=False)
    ax2.legend(loc="lower left", facecolor="#333F4F", ncol=2, framealpha=0.5, labelcolor="#FFFFFF", bbox_to_anchor=(0.5, -0.20), prop=font_properties, frameon=False)

    # Set y-axis limits for better visual match to example
    ax1.set_ylim(0, 120)
    ax1.yaxis.set_major_formatter(PercentFormatter(xmax=100))
    ax2.set_ylim(0, 90000)
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda x, _: f"{int(x):,}"))

    # Rotate x-tick labels for readability
    plt.xticks(rotation=45, color="#FFFFFF")

    plt.grid(True, color='grey', alpha=0.5, axis='y')
    plt.rcParams['axes.spines.left'] = False
    plt.rcParams['axes.spines.right'] = False
    plt.rcParams['axes.spines.top'] = False
    plt.rcParams['axes.spines.bottom'] = True

    # Save the plot to a file with the specified background color
    plt.tight_layout()
    plt_path = variables.img_path
    plt.savefig(plt_path)

    left = Inches(0.53)
    top = Inches(0.65)
    width = Inches(12.04)
    height = Inches(3.9)
    pic = slide9.shapes.add_picture(plt_path, left, top, width, height)

    return slide9