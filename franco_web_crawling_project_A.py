
# Franko Web Crawling Project A


# Package loading

print("패키지 로딩중입니다..\n\n")

from requests import get
from pandas import DataFrame, read_excel
from bs4 import BeautifulSoup as bs
from numpy import array
from sys import exit
from re import findall

# Loading excel files

print("'A하 제조사 KSIC 분류 조사' 자동 웹 크롤링 프로그램입니다.\n\n")

file_name = input("엑셀 파일의 경로를 확장자명을 포함하여 작성해 주세요.\n(예시 : C:/Users/admin/Desktop/Franko/더체크_소상공인_자료.xlsx)\n >>> ")
sheet_name = input("엑셀 시트 이름을 작성해 주세요.\n >>> ")

# If there is no data for crawling...

try:
    df = read_excel(file_name, sheet_name)
    
except:
    print("엑셀 파일 경로와 시트 이름을 다시 확인해 주세요.\n")
    input("엔터를 누르면 프로그램이 종료됩니다.")
    exit()
    

# Pre-work for Excel file output
  
df_columns = ["번호","KSIC","기업명","설립연도","대표자","지역","종업원수","매출액","영업이익","영업이익률","순수익","순수익률","생산 품목","특허수","비고"]
df_rst = DataFrame(columns=df_columns)

com_KSIC_lst = df["KSIC"].values.tolist()
com_name_lst = df["기업명"].values.tolist()
com_dist_lst = df["지역"].values.tolist()
com_sales_lst = df["매출액"].values.tolist()
com_profit_lst = df["영업이익"].values.tolist()
com_profit_rate_lst = df["영업이익률"].values.tolist()
com_pprofit_lst = df["순수익"].values.tolist()
com_pprofit_rate_lst = df["순수익률"].values.tolist()

com_lst = []
com_lst_count = 0

while(1):
    try:
        lst = []
        lst.append(com_name_lst[com_lst_count])
        lst.append(com_dist_lst[com_lst_count])
        sales_str = str(com_sales_lst[com_lst_count])
        lst.append(sales_str)
        com_lst.append(lst)
        com_lst_count += 1
    except IndexError:
        break

# Data crawling

for i in range(len(com_lst)):
    
    com_name = com_name_lst[i]
    search_url = "https://www.atfis.or.kr/business/M002040000/list.do?searchTopIndustry=ALL&businessOpen=001&_businessOpen=on&businessOpen=002&_businessOpen=on&businessOpen=003&_businessOpen=on&businessOpen=004&_businessOpen=on&businessOpen=005&_businessOpen=on&businessOpen=006&_businessOpen=on&sortOrderBy=S_DESC&searchMinYear=2019&searchMaxYear=2019&recordCount=1000&searchBusinessNm={0}&searchProductItem=&searchKeyIndicators=S&searchKeyIndicators=E&searchKeyIndicators=N&x=87&y=20#".format(com_name)

    try:
        res = get(search_url)
    except:
        print("페이지가 일시적으로 다운되었거나 인터넷 연결이 끊어졌습니다. 잠시 후 다시 실행해 주세요. 작업물은 중단된 항목의 이전 항목까지만 저장됩니다.\n")
        break
        
    print("{}번째 항목 처리 중...".format(i+1))

    text = res.text
    soup = bs(text, 'html.parser')
    
    rst_info = soup.select("#contents > div.business_list > div.info > table")[0].text.split()[6:]
    rst_sales = soup.select("#contents > div.business_list > div.price > table")[0].text.split()[6:]

    index_count = 0
    lst_lst = []

    while(1):
        index_count += 1
        if str(index_count + 1) not in rst_info:
            rst_lst = rst_info[rst_info.index(str(index_count)):][-2:]
            lst_lst.append(rst_lst)
            break
        fst_index = rst_info.index(str(index_count))
        snd_index = rst_info.index(str(index_count + 1))
        rst_lst = rst_info[fst_index:snd_index][-2:]
        lst_lst.append(rst_lst)
        
    index_sales_count = 0
    index_info_count = 0

    while(1):
        try:
            lst_lst[index_info_count].append(rst_sales[index_sales_count].replace(",",""))
            index_sales_count += 3
            index_info_count += 1
        except IndexError:
            break

    index_count3 = 0
    
    for lst in lst_lst:
        index_count3 += 1
        if lst == com_lst[i]:
            break

    com_url = "#contents > div.business_list > div.info > table > tbody > tr:nth-child({0}) > td:nth-child(3) > u > a".format(index_count3)
    
    rst_res = get("https://www.atfis.or.kr/business/M002040000/{}".format(soup.select(com_url)[0]["href"]))
    rst_soup = bs(rst_res.text, 'html.parser')
    ceo_name = ""
    est_date = ""
    emplo_num = ""
    prod_items = ""
    pat_all_num = ""
    pat_num = ""
    pat_items = ""
    
    try:
        ceo_name = rst_soup.select("#contents > table:nth-child(8) > tbody > tr:nth-child(3) > td")[0].text.split()[0]
    except IndexError:
        pass
    try:
        est_date = rst_soup.select("#contents > table:nth-child(8) > tbody > tr:nth-child(4) > td:nth-child(2)")[0].text.split()[0][0:4]
    except IndexError:
        pass
    try:
        emplo_num = findall("\d+", rst_soup.select("#contents > table:nth-child(8) > tbody > tr:nth-child(6) > td")[0].text.split()[2])[0]
    except IndexError:
        pass
    try:
        prod_items = " ".join(rst_soup.select("#contents > table:nth-child(8) > tbody > tr:nth-child(10) > td")[0].text.split())
    except IndexError:
        pass
    try:
        pat_num = int(rst_soup.select("#contents > table:nth-child(18) > tbody > tr:nth-child(3) > td:nth-child(6)")[0].text.split()[0])
    except IndexError:
        pass
    try:
        pat_all_num = rst_soup.select("#contents > table:nth-child(18) > tbody > tr:nth-child(3) > td.r.last")[0].text.split()[0]
    except IndexError:
        pass
    try:
        if pat_num != "":
            pat_items = []
            for i0 in range(pat_num):
                pat_item = " ".join(rst_soup.select("#contents > table:nth-child(19) > tbody > tr:nth-child({0}) > td.last".format(str(i0 + 1)))[0].text.split())
                pat_items.append(pat_item)
            pat_items = " | ".join(pat_items)
                
    except IndexError:
        pass
    
# Storing data to Excel file    
    
    add_info_lst = []
    
    add_info_lst.append(i + 1)
    add_info_lst.append(com_KSIC_lst[i])
    add_info_lst.append(com_name_lst[i])
    add_info_lst.append(est_date)
    add_info_lst.append(ceo_name)
    add_info_lst.append(com_dist_lst[i])
    add_info_lst.append(emplo_num)
    add_info_lst.append(com_sales_lst[i])
    add_info_lst.append(com_profit_lst[i])
    add_info_lst.append(com_profit_rate_lst[i])
    add_info_lst.append(com_pprofit_lst[i])
    add_info_lst.append(com_pprofit_rate_lst[i])
    add_info_lst.append(prod_items)
    add_info_lst.append(pat_all_num)
    add_info_lst.append(pat_items)
    
    df_item_rst = DataFrame(data=array([add_info_lst]),columns=["번호","KSIC","기업명","설립연도","대표자","지역","종업원수","매출액","영업이익","영업이익률","순수익","순수익률","생산 품목","특허수","비고"])
    df_rst = df_rst.append(df_item_rst)

rst_file = sheet_name + "_" + 'result.xlsx'

df_rst.to_excel(rst_file)

print("\n결과물이 해당 프로그램이 있는 디렉토리에 '{}'로 저장되었습니다.".format(rst_file))

input("\n엔터를 누르면 프로그램이 종료됩니다.")
exit()

# Created by jihyun jung