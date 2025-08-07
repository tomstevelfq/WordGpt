from langchain_openai import ChatOpenAI
# 初始化 LangChain LLM
llm = ChatOpenAI(
    model="qwen2.5-72b-instruct",
    api_key='sk-f5cbec200c9d444ebcd7cb2a0b28b006', 
    base_url='https://dashscope.aliyuncs.com/compatible-mode/v1',
    temperature=0.1
)