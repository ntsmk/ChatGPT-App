import os
import subprocess
from pathlib import Path
import openpyxl
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment


base_dir = Path(__file__).parent
excel_path = base_dir / 'chat_log.xlsx'


def is_output_open_excel() -> bool:
    """
    tell you if the Excel file is open or not
    :return: if it is open or not
    """
    #windows
    if os.name == "nt":
        try:
            with excel_path.open("r+b"):
                return False
        except PermissionError:
            return True
        except FileNotFoundError:
            return False
    #mac
    elif os.name == "posix":
        pass


def load_or_create_workbook() -> tuple[openpyxl.Workbook,bool]:
    """
    read Excel file or create it if it does not exist
    :return: workbook object and flag if the file is created or not
    """
    # checking file existence
    if excel_path.exists():
        # read file and return it
        wb = openpyxl.load_workbook(excel_path)
        return wb, False

    else:
        # create file and return it
        wb = openpyxl.Workbook()
        return wb, True

def create_worksheet (title:str,wb:openpyxl.Workbook,is_new:bool):
    """

    :param title: sheet title
    :param wb: workbook object
    :param is_new: if workbook is created new or not
    :return: worksheet object
    """
    # removing the invalid title
    trimmed_title = trim_invalid_chars(title)
    if is_new:
        # getting active sheet
        ws = wb.active
        ws.title = title

    else:
        # adding sheets
        ws = wb.create_sheet(title=trimmed_title)
        wb.move_sheet(ws,offset=-len(wb.worksheets)+1)
        wb.active = ws
    return ws

def trim_invalid_chars(title:str) -> str:
    """
    remove the invalid character from the sheet
    :param title: sheet title
    :return: sheet title after removing
    """
    invalid_chars = ["/","\\","?","*","[","]"]
    for char in invalid_chars:
        title = title.replace(char,"")
    return title

def header_formatting(ws):
    """
    change the header of worksheet
    :param ws: worksheet to export
    :return:
    """
    # write date on A1
    ws["A1"].value = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    ws["A1"].font = Font(name="Times New Roman")

    # set character on A2 and B2
    role_header_cell = ws["A2"]
    content_header_cell = ws["B2"]

    # set values on cell
    role_header_cell.value = "Role"
    content_header_cell.value = "Description"

    # set font
    header_font_style = Font(name="Times New Roman", bold=True, color="FFFFFF")
    role_header_cell.font = header_font_style
    content_header_cell.font = header_font_style

    # set cell color
    header_color = PatternFill(fill_type="solid", fgColor="156B31")
    role_header_cell.fill = header_color
    content_header_cell.fill = header_color

    # set cell width
    # ws.columns.dimensions["A"].width = 22
    # ws.columns.dimensions["B"].width = 168
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 165

def write_chat_log(ws,chat_log: list[dict]):
    """
    set the format for the chat history
    :param ws: worksheet to export
    :param chat_log: chat history
    """
    row_height_adjustment_standard = 18
    font_style = Font(name="Times New Roman", size=10)
    assistant_color = PatternFill(fill_type="solid",fgColor="d9d9d9")

    for row_number, content in enumerate(chat_log, 3):
        cell_role, cell_content = ws[f"A{row_number}"], ws[f"B{row_number}"]

        # write down on the cell
        cell_role.value = content["role"]
        cell_content.value = content["content"]

        # cell put space
        cell_content.alignment = Alignment(wrapText=True)

        # adjust culum height
        adjusted_row_height = len(content["content"].split("\n"))*row_height_adjustment_standard
        ws.row_dimensions[row_number].height = adjusted_row_height

        # setting format
        cell_role.font = font_style
        cell_content.font = font_style
        if content["role"] == "assistant":
            cell_role.fill = assistant_color
            cell_content.fill = assistant_color

def open_workbook():
    """open excel file"""
    # windows
    if os.name == "nt":
        os.system(f"start {excel_path}")

    # mac
    elif os.name == "posix":
        os.system(f"open{excel_path}")

def output_excel(chat_log:list[dict], chat_summary:str):
    """
    entry point to write the chat history on chat_hitstory.xlsx
    :param chat_log: chat history
    :param chat_summary: chat summary
    :return:
    """

    workbook, is_created = load_or_create_workbook()
    worksheet = create_worksheet(chat_summary,workbook,is_created)
    header_formatting(worksheet)
    write_chat_log(worksheet, chat_log)
    workbook.save(excel_path)
    workbook.close()


log = [
        {"role": "user", "content": "Hello"},
        {"role": "assistant", "content": "AI assistant"},
        {"role": "user", "content": "how are you?"},
        {"role": "assistant", "content": "I am fine\n tha\n nks,\nyou?\n not\n much?\n huh?\n really?\n how you doing\n these days"}
    ]

output_excel(log,"test/\\?*")