import output_excel
import os
from dotenv import load_dotenv
from openai import OpenAI


exit_command = "exit()"
default_model = "gpt-3.5-turbo"

# getting API key from env file
load_dotenv()
client = OpenAI(api_key=os.getenv('API_KEY'))


def give_role_to_system() -> str:
    """
    enter the role for AI
    :return: role for AI
    """
    print(f"Let's start conversation with AI. Enter {exit_command} if you wanna end it")

    # enter the role to give AI
    system_role = input("Enter the role for AI if you prefer. Hit enter key if you don't")

    return system_role

def generate_chat_log(gpt_model:str) -> list[dict]:

    # the list to save chat log
    chat_log: list[dict] = []

    # if role entered, add it to the chatlog
    system_role = give_role_to_system()
    if system_role:
        chat_log.append({"role":"system","content":system_role})

    while True:
        prompt = input("\nYou:")
        if prompt == exit_command:
            break

        # adding user to the chat log
        chat_log.append({"role": "user", "content": prompt})

        # getting AI's response
        response = client.chat.completions.create(model=gpt_model, messages=chat_log, stream=True)

        role, content = stream_and_concatenate_response(response)

        chat_log.append({"role": "role", "content": content})
        return chat_log

def stream_and_concatenate_response(response) -> tuple[str, str]:
    """
    shows the AI answer you got through streaming and show it in chunk, add it
    :param response: OpenAI.chat.completions.create()
    :return: AI response and role
    """
    print("\nAI:", end="")
    content_list: list[str] = []
    role = ""
    for chunk in response:
        chunk_delta = chunk.choices[0].delta
        content_chunk = chunk_delta.content if chunk_delta.content is not None else ""
        role_chunk = chunk_delta.role
        if role_chunk:
            role = role_chunk
        content_list.append(content_chunk)
        print(content_chunk, end="")
    else:
        print()
        content = "".join(content_list)

    return role, content

def fetch_gpt_model_list():
    """
    getting list of GPT model
    :return: GPT model list
    """
    # getting all list
    all_model_list = client.models.list()

    # getting only gpt model
    gpt_model_list = []
    for model in all_model_list:
        if "gpt" in model.id:
            gpt_model_list.append(model.id)

    gpt_model_list.sort()

    return gpt_model_list

def choise_model(gpt_model_list:list[str]) -> str:
    """
    let user choose model to use
    :param gpt_model_list:
    :return: model name chosen
    """

    print("enter the model number and hit enter")
    for num,model in enumerate(gpt_model_list):
        print(f"{num}:{model}")

    while True:
        input_number = input(f"if nothing entered, {default_model} will be used. : ")

        # case nothing entered
        # case not numbers entered
        # case not numbers for the model list
        # case correct numbers entered

        # case nothing entered
        if not input_number:
            return default_model

        # case not numbers entered
        if not input_number.isdigit():
            print("enter numbers")

        # case not numbers for the model list
        elif not int(input_number) in range(len(gpt_model_list)):
            print("the number does not exist in the list")

        # case correct numbers entered
        else:
            user_choise_model_name = gpt_model_list[int(input_number)]
            return  user_choise_model_name

def get_initial_prompt(chat_log:list[dict])->str|None:
    """
    getting initial prompt from the chat log
    :param chat_log:
    :return: user's initial propmt
    """

    # getting user's initial prompt
    for log in chat_log:
        if log["role"] == "user":
            initial_prompt = log["content"]
            return initial_prompt

def generate_summary(initial_prompt:str,summary_length:int = 20) -> str:
    """
    summarize the user's initial prompt.
    :param initial_prompt:
    :param summary_length: the maximum length
    :return: summarized prompt
    """

    summary_request = {"role":"system","content":"your job is to summarize the user's request"
                                                 f"please summarize the user's request within {summary_length} words."}
    messages = [summary_request,{"role":"user","content":initial_prompt}]
    response = client.chat.completions.create(model=default_model,messages=messages,max_tokens=summary_length)
    summary = response.choices[0].message.content
    adjustment_summary = summary[:summary_length]
    return adjustment_summary

def chat_runner() -> tuple[list[dict],str]:
    """
    start the chat, summarize chatlog and initial propmt and return them
    if error occurs, the chat will be closed
    :return: chatlog, initial summarized user's initial prompt
    """

    # getting all gpt model
    gpt_models = fetch_gpt_model_list()

    # choosing the model to use in the chat
    choise = choise_model(gpt_models)

    # getting chatlog
    generate_log = generate_chat_log(choise)

    # getting user's initial prompt
    initial_user_prompt = get_initial_prompt(generate_log)

    initial_prompt_summary = ""
    if initial_user_prompt:
        # summarizing the user's initial prompt
        initial_prompt_summary = generate_summary(initial_user_prompt)

    return generate_log,initial_prompt_summary

is_excel_open = output_excel.is_output_open_excel()
if not is_excel_open:
    chat_runner()
else:
    print("could not start the chat because the excel is open")