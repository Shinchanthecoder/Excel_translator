import pandas as pd
from googletrans import Translator
import time
import math

def translate_text(text):
    if isinstance(text, str) and text.lower() != 'nan':
        try:
            translator = Translator()
            translation = translator.translate(text, src='zh-cn', dest='en')
            return translation.text if translation.text is not None else ''
        except Exception as e:
            print(f"Error translating text: {text}. Error: {e}")
    return ''

def translate_excel(file_path):
    start_time = time.time()

   
    df = pd.read_excel(file_path, header=None)

    total_rows = len(df)
    progress_interval = 10  

    
    df.iloc[0] = df.iloc[0].apply(translate_text)

    
    for col in df.columns:  
        df[col] = df[col].apply(translate_text)

    
    translated_file_path = file_path.replace('.xls', '_translated.xls')
    df.to_excel(translated_file_path, index=False, header=False, engine='openpyxl')

    
    end_time = time.time()
    time_taken = end_time - start_time
    print(f"Translation completed in {time_taken:.2f} seconds. Translated file saved at {translated_file_path}")


file_path = 'C:\\Users\\HP\\Downloads\\Order Export.xls'
translate_excel(file_path)
