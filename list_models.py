import google.generativeai as genai
import os

genai.configure(api_key="AIzaSyDTT8hV7mi2cB7sYaZMQF7IeMZmOvn9a3A")

try:
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            print(m.name)
except Exception as e:
    print(e)
