from langchain_openai import ChatOpenAI
# 初始化 LangChain LLM
llm = ChatOpenAI(
    model="qwen2.5-72b-instruct",
    api_key='sk-4c7eea05ee72441d83716d9f00697769', 
    base_url='https://dashscope.aliyuncs.com/compatible-mode/v1',
    temperature=0.1
)