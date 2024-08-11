import output_excel
from chatgpt import chat_runner

is_excel_open = output_excel.is_output_open_excel()
if not is_excel_open:
    log, summary = chat_runner()
    output_excel.output_excel(log, summary)
else:
    print("excel file is open can't start the chat")