import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from datetime import datetime
from openpyxl.writer.excel import save_virtual_workbook
from PIL import Image
image = Image.open(r"D:\\Documents\\Desktop\\k2.jpg")




def get_week_day(day):
    week_day_dict = {
        0: "一",
        1: "二",
        2: "三",
        3: "四",
        4: "五"          
    }
    return week_day_dict[day]
now = datetime.now()
chinese_weekday = get_week_day(now.weekday())

st.title('''
This app is mi-app
''')

st.write('''
***
''')
st.header('上傳區')


uploaded_file = st.file_uploader("請上傳節目受訪者申請表xlsx檔", type = ".xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file,usecols="A:M",header = 1)

    df = df[(df['受訪內容']=="治安交通宣導") | (df['受訪內容']=="預防犯罪宣導") | (df['受訪內容']=="交通宣導")| (df['受訪內容']=="交通安全宣導")| (df['受訪內容']=="治安宣導")]
 
    wb = openpyxl.load_workbook(r"D:\\Documents\\Desktop\\111.xlsx")
    ws =wb['臺東分臺']
    ws.cell(row=1, column=1).value = f'''警察廣播電臺警政宣導節目宣導記錄表
                                                             {now.year - 1911}年{now.month}月{now.day}日（星期{chinese_weekday}）臺東分臺 潘亭羽 製表'''
   
    i = 4
    for ind in df.index:
        ws.cell(row=i, column=1).value = df.loc[ind, "日期"]
        ws.cell(row=i, column=2).value = df.loc[ind, "時間"]
        ws.cell(row=i, column=3).value = "臺東分臺\n（FM94.3）"
        ws.cell(row=i, column=4).value = df.loc[ind, "節目名稱"]
        ws.cell(row=i, column=5).value = df.loc[ind, "主持人"]
        ws.cell(row=i, column=6).value = df.loc[ind, "單位"]
        ws.cell(row=i, column=7).value = df.loc[ind, "職稱"]
        ws.cell(row=i, column=8).value = df.loc[ind, "受訪者"]
        ws.cell(row=i, column=9).value = df.loc[ind, "受訪內容"]
        ws.cell(row=i, column=10).value = "Facebook粉絲專頁發文、官方網站專區宣導、各警分局粉絲專頁連結警廣警政專訪網址分享"
        i = i + 1
    # k = ws["H4"].value
    # st.write(k)

    
        
    data = BytesIO(save_virtual_workbook(wb))
    st.write('''
    ***
    ''')
    st.header('下載區')
    st.download_button("下載檔案",
        data=data,
        mime='xlsx',
        file_name="yoyo_of_file.xlsx")
    st.image(image, caption='阿奇小狗狗')
    # st.download_button(
    #     label="下載檔案",
    #     data=wb,
    #     file_name='已轉換完成的檔案.xlsx',
    #     mime='text/csv',
    # )