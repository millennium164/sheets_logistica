from langchain_core.prompts import ChatPromptTemplate
from langchain_ollama.llms import OllamaLLM
import sys

template = """Question: {question}

Answer: Let's think step by step."""

prompt = ChatPromptTemplate.from_template(template)
model = OllamaLLM(model="phi3")
chain = prompt | model

question = "What is a derivative, in calculus?"

# Usamos o método .stream() para receber os chunks
for chunk in chain.stream({"question": question}):
    # O 'flush=True' garante que o texto apareça no terminal imediatamente
    # O 'end=""' evita que ele pule uma linha a cada palavra
    print(chunk, end="", flush=True)

print("\n") # Apenas para organizar o final da resposta