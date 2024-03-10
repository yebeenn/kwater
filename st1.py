import streamlit as st
import pandas as pd
from io import BytesIO

def process_excel_file(excel_file_path):
    # 엑셀 파일 읽기
    df = pd.read_excel(excel_file_path, header=None)

    # 행과 열 전환
    df = df.transpose()

    df1 = df[~df[0].astype(str).str.contains('토|일')]
    df1['2'] = df[0].astype(str).str[1]
    df1['3'] = df[0].astype(str).str[3:5]
    df1['4'] = df[0].astype(str).str[6]
    df1['5'] = '정상근무'
    df1[['6', '7']] = df[1].astype(str).str.split('~', expand=True)
    df1['7'] = df1['7'].astype(str).str[:5]
    df1['8'] = '1hr'

    df1 = df1.iloc[4:, 2:]

    target_string = '근거리출장'

    if df[1].astype(str).str.contains(target_string).any():
        df2 = df[df[1].astype(str).str.contains(target_string)]

        df2[0] = '근출'
        df2[1] = df[0].astype(str).str[0:5]
        df2['2'] = df[1].astype(str).str[0:5]
        df2['3'] = df[1].astype(str).str[6:11]

        return df1, df2
    else:
        # Handle the case when the target string is not present
        return df1, None

def main():
    st.header("근로자사용부 만들기:smile:")

    # 엑셀 파일 업로드
    uploaded_file = st.file_uploader("근태현황조회 엑셀파일을 업로드하세용", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # 업로드된 파일을 데이터프레임으로 읽어옴
        df1, df2 = process_excel_file(uploaded_file)

        # 결과 표시
        st.write("시트1:")
        st.write(df1)

        st.write("시트2:")
        if df2 is not None:
            st.write(df2)
        else:
            st.write("시트2가 없습니다.")

        buffer = BytesIO()

        with pd.ExcelWriter(buffer) as writer:
            df1.to_excel(writer, sheet_name='Sheet1', index=False, header=None)
            if df2 is not None:
                df2.to_excel(writer, sheet_name='Sheet2', index=False, header=None)

        buffer.seek(0)

        if st.download_button(
                label="다운로드",
                data=buffer,
                file_name="output.xlsx",
                key="download_button",
                help="다운로드 버튼을 클릭하여 Excel 파일을 다운로드하세요."
        ):
            st.success("다운로드 완료! 파일명: output.xlsx")


if __name__ == "__main__":
    main()
