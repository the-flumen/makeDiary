from set_28 import *
from set_29 import *

# 파워포인트 파일 경로 설정
# 입력 파일과 출력 파일 이름
# path = "C:/Users/zhel7/Documents/"


yearNum = 2025;
month2 = 28
path = ""
# 2025
hoilydae_list = [str(yearNum) +'-01-28', str(yearNum) +'-01-29',str(yearNum) +'-01-30'
                 ,str(yearNum) +'-03-03'
                 ,str(yearNum) +'-04-10'
                 ,str(yearNum) +'-05-06'
                 ,str(yearNum) +'-06-06'
                 ,str(yearNum) +'-09-15',str(yearNum) +'-09-16',str(yearNum) +'-09-17',str(yearNum) +'-09-18'
                 ,str(yearNum) +'-10-08',str(yearNum) +'-10-09'
                 ,str(yearNum) +'-12-25']

if (month2 == 29) :     
    set_29(yearNum, month2, path, hoilydae_list);
else : 
    set_28(yearNum, month2, path, hoilydae_list);
