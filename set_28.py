from functionDay import *
from functionW import *

def set_28 (yearNum, month2, path, hoilydae_list) : 

    basic_input_file = path + "diaryOrg" +str(month2)+".pptx"
    basic_output_file = path + "b.pptx"
    daily_week_output_file = path + "diary "+ str(yearNum) +".pptx"

    start_slide_index = 691;
    start_link_slid_index = 1803;

    keyword = "주간재정"
    
    page_config = [
        {"type": "yearly", "targetPage": 3, "repeat": 1, "source_slide_index" : 0, "link_page_num" : 150},
        {"type": "daily", "targetPage": 150, "repeat": 365, "source_slide_index" : 97, "link_page_num" : 150, "month_link_page_num" : 15, "week_link_page_num" : 97},
        {"type": "monthly", "targetPage": 656, "repeat": 12, "source_slide_index" : 0, "link_page_num" : 150},
        {"type": "weekly", "targetPage": 97, "repeat": 53, "source_slide_index" : 97, "link_page_num" : 150, "month_link_page_num" : 85},
        {"type": "MoneyWeekly", "targetPage": start_link_slid_index + 1, "repeat": 53, "source_slide_index" :0, "link_page_num" : 150},
    ]      
    daily_week_config = [
        {"type": "모닝로그", "startPage": start_slide_index, "start_slide_index" : start_slide_index},
        {"type": "일기장", "startPage": 1056, "start_slide_index" : 1056},
        {"type": "감정일기", "startPage": 1421, "start_slide_index" : start_slide_index},
        {"type": "일간계획", "startPage": 149, "start_slide_index" : start_slide_index},
    ]   
    lists = {
        "mon_unm_list" : [31,month2,31,30,31,30,31,31,30,31,30,31],
        "hoilydae_list" : hoilydae_list,
        "week_list" : ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일'],
        "Month_Ko_list" : ['해\n오\n름\n달', '시\n샘\n달', '물\n오\n름\n달', '잎\n새\n달', '푸\n른\n달', '누\n리\n달','견\n우\n직\n녀\n달', '타\n오\n름\n달', '열\n매\n달', '하\n늘\n연\n달', '마\n름\n달', '매\n듭\n달']
    }

    extract_D(basic_input_file, basic_output_file, page_config, yearNum, lists)
    extract_W(yearNum, keyword, basic_output_file, daily_week_output_file, daily_week_config, start_slide_index, start_link_slid_index)