import google.generativeai as genai
import os

genai.configure(api_key="AIzaSyAA1S8rSahgUIMcUKILsZL_k_uDrCSUmlY")

print("Available models:")
for m in genai.list_models():
  if 'generateContent' in m.supported_generation_methods:
    print(m.name)
