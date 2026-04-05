import google.generativeai as genai
import os

genai.configure(api_key="AIzaSyAA1S8rSahgUIHcUKILdZL_k_uDrCSUm1x")

print("Available models:")
for m in genai.list_models():
  if 'generateContent' in m.supported_generation_methods:
    print(m.name)
