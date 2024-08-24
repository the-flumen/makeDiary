from pptx import Presentation
from functionUtily import *
from set_monthly import *
from set_yearly import *
from set_daily import *
from set_weekly_money import *
from set_weekly import *
import sys
sys.setrecursionlimit(10000)             

def extract_D(input_file, output_file, page_config, yearNum, lists):
    print(yearNum)
    mon_unm_list = lists["mon_unm_list"]
    hoilydae_list = lists["hoilydae_list"]
    try:
        prs = Presentation(input_file)# 슬라이드의 도형과 텍스트 추출
        for pageCon in page_config :
            type = pageCon["type"]
                
            if type ==  "yearly" : 
                set_yearly (prs, mon_unm_list, yearNum, pageCon)
            elif type == "monthly" :
                set_monthly (prs, mon_unm_list, yearNum, pageCon, hoilydae_list);
            elif type == "daily" :
                Month_Ko_list = lists["Month_Ko_list"]
                set_daily (prs, mon_unm_list, yearNum, pageCon, Month_Ko_list, hoilydae_list)
            elif type == "weekly" :
                Month_Ko_list = lists["Month_Ko_list"]
                set_weekly (prs, mon_unm_list, yearNum, pageCon, Month_Ko_list, hoilydae_list)
            elif type == "MoneyWeekly" :
                Month_Ko_list = lists["Month_Ko_list"]
                set_weekly_money (prs, mon_unm_list, yearNum, pageCon, hoilydae_list) 

        prs.save(output_file)
        print(f"파일 '{output_file}'이(가) 생성되었습니다.")
              
    except Exception as e:
        print(f"파워포인트 파일을 열거나 데이터를 추출하는 중 오류가 발생했습니다: {e}")
        return None
        
    



