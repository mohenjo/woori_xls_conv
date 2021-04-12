"""우리은행/우리카드 거래 내역 파일 정리툴

자체 거래 내역 작성에 사용이 용이하도록 우리은행/우리카드의 거래내역 액셀파일을
정리하는 프로그램입니다.
"""

import xlrd
import xlwt
import typing
import win32com.client as win32
import os
import tkinter as tk
import tkinter.filedialog as tkfd


def main():
    mainwindow = tk.Tk()
    _ = Application(mainwindow)
    mainwindow.mainloop()


def get_outputfile_path(inputfile: str):
    path, ext = os.path.splitext(inputfile)
    return path + "_cnv" + ext


def filetype_chk(inputfile: str):
    """액셀 파일 판단
    [1, 2, 3] = [우리은행거래내역, 우리카드거래내역, 다른 내용의 액셀 파일]
    """

    wb = xlrd.open_workbook(inputfile, on_demand=True)
    ws = wb.sheet_by_index(0)

    # 파일의 첫 행을 읽어 액셀 파일 형식 판단
    filetype = 3
    checkval = "".join([str(c) for c in ws.row(0)])
    if "최근거래내역조회" in checkval:  # 우리은행거래내역
        filetype = 1
    elif "승인 상세내역" in checkval:  # 우리카드거래내역
        filetype = 2

    del wb
    return filetype


def process_file(filetype, inputfile: str, outputfile: str):
    """액셀 파일 형태에 따른 프로세스 수행"""
    if filetype == 1:  # 우리은행 거래내역
        process_bank_file(inputfile, outputfile)
    elif filetype == 2:  # 우리카드 거래내역
        process_card_file(inputfile, outputfile)
    else:
        raise Exception("요청하는 형식의 액셀 파일이 아닙니다.")


def process_bank_file(inputfile: str, outputfile: str):
    """우리은행 파일 프로세싱"""

    # region ===== 파일 읽기 =====

    rwb = xlrd.open_workbook(inputfile, on_demand=True)
    rws = rwb.sheet_by_index(0)

    # 헤더 읽기
    header = rws.row_values(3)

    # 거래 내역 읽기
    transactions = []
    for row_idx in range(4, rws.nrows):
        getrow = rws.row_values(row_idx)
        transactions.append(getrow)

    # 거래내역 과거내역 먼저로 정렬
    if transactions[0][1] > transactions[1][1]: # 날짜 비교
        transactions= transactions[::-1]

    del rwb

    # endregion

    # region ===== 변경할 내용 지정 =====

    # 변경을 위한 형식 지정
    # ['No.', '거래일시', '적요', '기재내용', '지급(원)', '입금(원)', '거래후 잔액(원)', '취급점', '메모', '수표·어음·증권금액(원)']

    cnv_type: typing.List[typing.Union, typing.Callable] = [None] * len(header)  # 형 변경을 위한 리스트
    cnv_type[0] = int  # No
    cnv_type[4] = int  # 지급(원)
    cnv_type[5] = int  # 입금(원)
    cnv_type[6] = int  # 거래후 잔액(원)
    cnv_type[9] = int  # 수표·어음·증권금액(원)

    # endregion

    # region ===== 파일 쓰기 =====

    # 새 액셀 파일에 저장
    wwb = xlwt.Workbook(encoding='utf-8')
    wws = wwb.add_sheet("sheet1", cell_overwrite_ok=True)

    # 헤더 쓰기
    for col_idx, title in enumerate(header):
        wws.write(0, col_idx, title)

    # 거래 내역 쓰기
    for row_idx, transaction in enumerate(transactions):
        for col_idx, val in enumerate(transaction):
            if cnv_type[col_idx] is not None:
                val = cnv_type[col_idx](val)
            wws.write(row_idx + 1, col_idx, val)

    wwb.save(outputfile)
    del wwb

    # 액셀 자동 폭
    autofit_excel_file(outputfile)

    # endregion


def process_card_file(inputfile: str, outputfile: str):
    """우리카드 파일 프로세싱"""

    # region ===== 파일 읽기 =====
    rwb = xlrd.open_workbook(inputfile, on_demand=True)
    rws = rwb.sheet_by_index(0)

    # 헤더 읽기
    header = rws.row_values(2)

    # 거래 내역 읽기
    transactions = []
    for row_idx in range(3, rws.nrows):
        getrow = rws.row_values(row_idx)
        if len(getrow[0]) != 0 and getrow != header:
            transactions.append(getrow)

    del rwb

    # endregion

    # region ===== 변경할 내용 지정 =====

    # 빈 컬럼 지우기
    del_col_idx = [1, 5, 7, 10]  # 지울 컬럼
    header = [v for i, v in enumerate(header) if i not in del_col_idx]
    for row_idx in range(len(transactions)):
        transactions[row_idx] = [v for i, v in enumerate(transactions[row_idx]) if i not in del_col_idx]

    # 컬럼 내용 기본 변경
    # 할부개월, 승인금액, 부가세, 취소금액을 배치 변경 가능한 상태로 변경
    for row_idx in range(len(transactions)):
        for col_idx in range(9, len(transactions[row_idx])):
            val = transactions[row_idx][col_idx]
            val = '0' if val == '' else val
            val = "".join(val.split(",")) if "," in val else val
            transactions[row_idx][col_idx] = val

    # ['이용\n일자', '승인번호', '이용카드', '이용가맹점\n(은행)명', '가맹점 주소', '연락처', '업종', '사업자번호',
    # '매출\n구분', '할부\n개월', '승인금액', '부가세', '취소금액']
    cnv_type: typing.List[typing.Union, typing.Callable] = [None] * len(header)  # 형 변경을 위한 리스트
    cnv_type[1] = int  # 승인번호
    cnv_type[2] = int  # 이용카드
    cnv_type[7] = int  # 사업자번호
    cnv_type[9] = int  # 할부개월
    cnv_type[10] = int  # 승인금액
    cnv_type[11] = int  # 부가세
    cnv_type[12] = int  # 취소금액

    # endregion

    # region ===== 파일 쓰기 =====

    # 새 액셀 파일에 저장
    wwb = xlwt.Workbook(encoding='utf-8')
    wws = wwb.add_sheet("sheet1", cell_overwrite_ok=True)

    # 헤더 쓰기
    for col_idx, title in enumerate(header):
        wws.write(0, col_idx, title)

    # 거래내역 과거 내역 먼저로 정렬
    if transactions[0][0] > transactions[1][0]:
        transactions = transactions[::-1]

    # 거래 내역 쓰기
    for row_idx, transaction in enumerate(transactions):
        for col_idx, val in enumerate(transaction):
            if cnv_type[col_idx] is not None:
                val = cnv_type[col_idx](val)
            wws.write(row_idx + 1, col_idx, val)

    wwb.save(outputfile)
    del wwb

    # 액셀 자동 폭
    autofit_excel_file(outputfile)

    # endregion


def autofit_excel_file(filepath: str):
    """액셀 파일 자동 폭 설정 후 재저장"""
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    loadwb = excel.Workbooks.Open(filepath)

    for sh in loadwb.Sheets:
        ws = loadwb.Worksheets(sh.Name)
        ws.Columns.AutoFit()
        ws.Rows.AutoFit()

    loadwb.Save()
    excel.Application.Quit()


class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("우리은행/우리카드 파일 정리툴")
        self.master.geometry("400x130")
        self.master.resizable(False, False)

        self.btn_openfile = tk.Button(self.master, text="Open & Process Woori Bank/Card file",
                                      width=50, height=5,
                                      command=self.load_file)
        self.btn_openfile.place(x=20, y=20)

    def load_file(self):
        self.inputfile = tkfd.askopenfilename(initialdir=os.getenv('USERPROFILE') + r"\\Downloads",
                                              title="액셀 파일을 선택하세요",
                                              filetypes=(("액셀파일", "*.xls *.xlsx"),)
                                              )
        if not self.inputfile:
            return None

        self.outputfile = get_outputfile_path(self.inputfile)
        self.filetype = filetype_chk(self.inputfile)

        if self.filetype == 3:
            print("요청하는 형식의 액셀파일이 아닙니다.")
            return None

        process_file(self.filetype, self.inputfile, self.outputfile)
        print("정리가 완료되었습니다.")
        print(f"저장파일명: {self.outputfile}")
        os.startfile(self.outputfile)


if __name__ == '__main__':
    main()
