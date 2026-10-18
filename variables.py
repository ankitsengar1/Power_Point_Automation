
import calendar

# img path 
img_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/senco_deck/data/senco_img.png"

#excel path
ppt_xl_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/senco_deck/data/SencoGold Mar'25 MBR 1.xlsx"
deck_path = f"/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/senco_deck/data/MBR_Mar'25_SencoGold.pptx"
previous_deck_path = f"/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/senco_deck/data/MBR_Feb'25_SencoGold.pptx"
template_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/senco_deck/data/MBR_SencoGold_template.pptx"

# current month number for which Deck needs to be created
curr_month_number = 3
fy_year_start_month = 4

# current year
curr_year = 2025
# previous year
prev_year = curr_year-1
prev_2year = curr_year-2

# month numbers
prev_month_number1 = (curr_month_number-1) % 12 or 12
prev_month_number2 = (curr_month_number-2) % 12 or 12
six_months_back_number = (curr_month_number - 5) % 12 or 12

# current month string 
curr_month_name = calendar.month_name[curr_month_number]
current_month = f"{curr_month_name[:3]}'{curr_year % 100}"

# current month previous year string 
curr_month_prev_year_name = calendar.month_name[curr_month_number]
current_month_prev_year = f"{curr_month_name[:3]}'{prev_year % 100}"

# previous i month string 
if curr_month_number == 1:
    curr_year_prev_month1_name = calendar.month_name[prev_month_number1]
    current_year_prev_month1 = f"{curr_year_prev_month1_name[:3]}'{prev_year % 100}"
    
else:
    curr_year_prev_month1_name = calendar.month_name[prev_month_number1]
    current_year_prev_month1 = f"{curr_year_prev_month1_name[:3]}'{curr_year % 100}"

# previous 2 month string
if curr_month_number in (1, 2):
    curr_year_prev_month2_name = calendar.month_name[prev_month_number2]
    current_year_prev_month2 = f"{curr_year_prev_month2_name[:3]}'{prev_year % 100}"
    
else:
    curr_year_prev_month2_name = calendar.month_name[prev_month_number2]
    current_year_prev_month2 = f"{curr_year_prev_month2_name[:3]}'{curr_year % 100}"

# finacial year string
if curr_month_number in (1, 2, 3):
    ytd_curr_year_name = calendar.month_name[fy_year_start_month]
    ytd_current_year = f"{ytd_curr_year_name[:3]}'{prev_year % 100}-{curr_month_name[:3]}'{curr_year % 100}"
    ytd_previous_year = f"{ytd_curr_year_name[:3]}'{prev_2year % 100}-{curr_month_name[:3]}'{prev_year % 100}"

else:
    ytd_curr_year_name = calendar.month_name[fy_year_start_month]
    ytd_current_year = f"{ytd_curr_year_name[:3]}'{curr_year % 100}-{curr_month_name[:3]}'{curr_year % 100}"
    ytd_previous_year = f"{ytd_curr_year_name[:3]}'{prev_year % 100}-{curr_month_name[:3]}'{prev_year % 100}"

# Six month back string 
if curr_month_number in (1,2,3,4,5,6):
    six_months_back_name = calendar.month_name[six_months_back_number]
    current_year_six_month = f"{six_months_back_name[:3]}'{prev_year % 100}-{curr_month_name[:3]}'{curr_year % 100}"
    previous_year_six_month = f"{six_months_back_name[:3]}'{prev_2year % 100}-{curr_month_name[:3]}'{prev_year % 100}"

else:
    six_months_back_name = calendar.month_name[six_months_back_number]
    current_year_six_month = f"{six_months_back_name[:3]}'{curr_year % 100}-{curr_month_name[:3]}'{curr_year % 100}"
    previous_year_six_month = f"{six_months_back_name[:3]}'{prev_year % 100}-{curr_month_name[:3]}'{prev_year % 100}"