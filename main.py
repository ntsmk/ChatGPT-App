from output_excel import output_excel, is_output_open_excel, excel_path
from chatgpt import chat_runner

is_excel_open = output_excel.is_output_open_excel()
if not is_excel_open:
    log, summary = chat_runner()
    output_excel.output_excel(log, summary)
else:
    print(f"{excel_path.name} is open can't start the chat")