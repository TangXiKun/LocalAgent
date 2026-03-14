from urllib import response
from ollama import Client
import keyboard
from tools import *
from markdown_to_word import convert_markdown_to_word
import streamlit as st


def AI_response():
    global tool_call, content
    in_thinking = False
    content = ""
    tool_call = []
    for chunk in response:
        if keyboard.is_pressed("esc"):
            break

        text = ""
        if chunk.message.thinking:
            if not in_thinking:
                in_thinking = True
                text = "**Thinking...**:\n\n"
            text += chunk.message.thinking
        elif chunk.message.content:
            if in_thinking:
                in_thinking = False
                text = "\n\n**Answer:**\n\n"
            text += chunk.message.content
            content += chunk.message.content

        if chunk.message.tool_calls:
            tool_call.extend(chunk.message.tool_calls)

        yield text


if "messages" not in st.session_state:

    st.session_state.client = Client(host="http://localhost:11434/")

    st.session_state.RECOMMAND_MODELS = [
        "glm-4.7-flash:latest",
        "qwen3-coder:latest",
        "qwen3-coder:30b",
        "qwen3:latest",
        "qwen3:8b",
        "qwen3:4b",
    ]

    st.session_state.model_list = [
        i["model"] for i in st.session_state.client.list()["models"]
    ]

    if st.session_state.model_list:
        st.session_state.MODEL = st.session_state.model_list[0]
        for i in st.session_state.RECOMMAND_MODELS:
            if i in st.session_state.model_list:
                st.session_state.MODEL = i
                break
    else:
        st.error("No avilable ollama models, please download one eg.`ollama pull qwen3:4b`")
        raise Exception("No avilable ollama models")

    st.session_state.MODEL_NAME = "Assistant"
    st.session_state.PROMPT = open("./prompt/prompt.md", "r", encoding="utf-8").read()
    st.session_state.IS_CONTINUE = False

    st.session_state.EXTERNAL_FILES = os.listdir("D:/ExternalFiles/")
    file_list = []
    for i in st.session_state.EXTERNAL_FILES:
        if ".md" in i:
            python_file = i.replace("md", "py")
            if python_file in st.session_state.EXTERNAL_FILES:
                file_list.append(python_file)
    st.session_state.PROMPT = st.session_state.PROMPT.replace(
        "$EXTERNALFILES$", str(file_list)
    )

    st.session_state.KNOWLEDGE = open(
        "D:/ExternalFiles/KNOWLEDGE.txt", "r", encoding="utf-8"
    ).read()
    st.session_state.PROMPT = st.session_state.PROMPT.replace(
        "$KNOWLEDGE$", st.session_state.KNOWLEDGE
    )

    print(st.session_state.PROMPT)

    st.session_state.messages = [{"role": "system", "content": st.session_state.PROMPT}]

    st.session_state.tools = [
        run_powershell,
        wait_user_do,
        read_file,
        create_file,
        run_python_code,
        replace_file_content,
        replace_ppt_content,
        create_ppt_from_txt,
        read_ppt_and_export_txt,
        read_word_and_export_txt,
        read_excel_and_export_txt,
        create_excel_from_2d_list,
        convert_markdown_to_word,
        convert_word_or_txt_to_pdf,
        convert_ppt_to_pdf,
        create_python_tool,
        recognize_image_and_export_markdown,
        add_knowledge,
        # get_screen_image,
    ]

    st.session_state.available_functions = {
        "run_powershell": run_powershell,
        "wait_user_do": wait_user_do,
        "read_file": read_file,
        "create_file": create_file,
        "run_python_code": run_python_code,
        "replace_file_content": replace_file_content,
        "replace_ppt_content": replace_ppt_content,
        "create_ppt_from_txt": create_ppt_from_txt,
        "read_ppt_and_export_txt": read_ppt_and_export_txt,
        "read_word_and_export_txt": read_word_and_export_txt,
        "read_excel_and_export_txt": read_excel_and_export_txt,
        "create_excel_from_2d_list": create_excel_from_2d_list,
        "convert_markdown_to_word": convert_markdown_to_word,
        "convert_word_or_txt_to_pdf": convert_word_or_txt_to_pdf,
        "convert_ppt_to_pdf": convert_ppt_to_pdf,
        "create_python_tool": create_python_tool,
        "recognize_image_and_export_markdown": recognize_image_and_export_markdown,
        "add_knowledge": add_knowledge,
        # "get_screen_image": get_screen_image,
    }

st.set_page_config(
    page_title="LocalAgent",
    layout="wide",
    initial_sidebar_state="expanded",
)

example1 = """| Function Category         | Operation Example          |
| ------------------------- | -------------------------- |
| **File Operations**     | ✅ Office Files           |
| **Code Execution**      | ✅ Data Process        |
| **Image Recognition**   | ✅ Markdown           |
| **Document Conversion** | ✅ Word to PDF            |
| **Chart Generation**    | ✅ Bar Chart              |
| **System Interaction**  | ✅ PowerShell             |
| **Other...**            | ✅ More ...     |"""

example2 = """- `✅ Calculate 160968*(23516-75061)`
- `✅ Create a HTML Snake game`
- `✅ Make a PPT to introduce yourself`
- `✅ Convert all Word documents on the desktop to PDF and save them in one folder`
- `✅ Organize the word list image on the desktop into a Word file on the desktop according to parts of speech`
- `✅ On the desktop, there is 'data.xlsx', which records two sets of data corresponding to each time point. Please plot a line chart and a pie chart.`"""

with st.sidebar:

    st.session_state.MODEL = st.selectbox(
        "Ollama Model",
        options=st.session_state.model_list,
        index=st.session_state.model_list.index(st.session_state.MODEL),
    )
    st.divider()
    with st.expander(label="Function Introduction"):
        st.markdown(example1)
    with st.expander(label="Task Examples"):
        st.markdown(example2)
    st.markdown("> Tip: To stop, click `Stop` and then `...`, `Rerun`")

    if st.session_state.IS_CONTINUE:
        st.button("remove last message", width="stretch", disabled=True)
    else:
        if st.button("remove last message",width="stretch"):
            if st.session_state.messages:
                st.session_state.messages = st.session_state.messages[:-1]
                st.rerun()
    st.markdown("\n \n \nhttps://github.com/TangXiKun/LocalAgent")

st.write("# :material/layers: LocalAgent")
st.markdown("### `developed by 唐希鲲(Xikun Tang)`")
st.markdown("> —— A powerful tool for operating computers and handling work tasks")
st.divider()

for i in st.session_state.messages[1:]:
    if i["role"] == "user":
        with st.chat_message(i["role"], width="stretch", avatar="./images/user.png"):
            st.markdown(i["content"])
    elif i["role"] == "assistant" and i["content"] not in ["","\n"]:
        with st.chat_message(i["role"], width="stretch", avatar="./images/AI.png"):
            # with st.container(border=True):
            st.markdown(i["content"])
    elif i["role"] == "tool":
        with st.chat_message(i["role"], width="stretch", avatar="./images/tool.png"):
            with st.expander(label="工具调用: `%s`"%i["tool_name"],expanded=True):
                st.code(i["content"])

if st.session_state.IS_CONTINUE == False:
    user_inp = st.chat_input("input you task here", width="stretch")
    if user_inp:
        st.session_state.messages.append({"role": "user", "content": user_inp})
        st.session_state.IS_CONTINUE = True
        st.rerun()
else:
    st.session_state.IS_CONTINUE = False
    user_inp = st.chat_input(
        "task is in progress ...", key=0, width="stretch", disabled=True
    )
    with st.chat_message("assistant", width="stretch", avatar="./images/AI.png"):
        # with st.container(border=True):
        response = st.session_state.client.chat(
            model=st.session_state.MODEL,
            messages=st.session_state.messages,
            tools=st.session_state.tools,
            stream=True,
        )
        st.write_stream(AI_response())

        st.session_state.messages.append({"role": "assistant", "content": content})
        print(tool_call)
        for tc in tool_call:
            if tc.function.name in st.session_state.available_functions:
                print(
                        f"Calling {tc.function.name} with arguments {tc.function.arguments}"
                    )
                try:
                    result = st.session_state.available_functions[tc.function.name](
                            **tc.function.arguments
                        )
                    print(f"Result: {result}")
                    st.success("工具调用成功")
                except:
                    st.error("工具调用失败")
                st.session_state.messages.append(
                        {
                            "role": "tool",
                            "tool_name": tc.function.name,
                            "content": str(result),
                        }
                    )
                st.session_state.IS_CONTINUE = True

    st.rerun()
