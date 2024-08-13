import openai

import output_excel
import os
from dotenv import load_dotenv
from openai import OpenAI
import httpx
from colorama import Fore


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

def input_user_prompt()->str:
    """
    enter the user prompt
    :return: user prompt
    """
    user_prompt = ""
    while not user_prompt:
        user_prompt = input(f"{Fore.CYAN}\n you:{Fore.RESET}")
        if not user_prompt:
            print("please enter the prompt")
    return user_prompt

def generate_chat_log(gpt_model:str) -> list[dict]:

    # the list to save chat log
    chat_log: list[dict] = []

    # if role entered, add it to the chatlog
    system_role = give_role_to_system()
    if system_role:
        chat_log.append({"role":"system","content":system_role})

    while True:
        prompt = input_user_prompt()
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
    print(f"{Fore.GREEN}\nAI:{Fore.RESET}", end="")
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

def print_error_message(message:str):
    """
    show the error message
    :param message:
    :return:
    """
    print(f"{Fore.RED}{message}{Fore.RESET}")

def fetch_gpt_model_list()->list[str]|None:
    """
    getting list of GPT model
    :return: GPT model list
    """
    # getting all list
    try:
        all_model_list = client.models.list()
        response = httpx.Response(500,request=httpx.Request("GET","test"))
        #raise openai.RateLimitError(message="test",body="test",response=response)

    except openai.APIError:
        print_error_message("API error occurred")

    else:
        #getting only gpt model
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
            print(f"{Fore.RED}enter numbers{Fore.RESET}")

        # case not numbers for the model list
        elif not int(input_number) in range(len(gpt_model_list)):
            print(f"{Fore.RED}the number does not exist in the list{Fore.RESET}")

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

    # if you did not get gpt model list, exit
    if not gpt_models:
        exit()

    # choosing the model to use in the chat
    choise = choise_model(gpt_models)

    # getting chatlog
    generate_log = generate_chat_log(choise)

    if not generate_log:
        exit()

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

if __name__ == "__main__" :
    chat_runner()