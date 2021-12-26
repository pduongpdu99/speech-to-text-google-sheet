import speech_recognition as sr
import pandas as pd 
import lib.exportExcelToGooglesheet as exportExcelToGooglesheet 
import random

def choice_voice(r, source, name):
    print(f"Please say what {name}....")
    audio = r.listen(source)
    # language='vi-VN'
    try:
        query = r.recognize_google(audio, language='vi-VN')
        print(f"{name} index: \n\t" + query)
        return query
    except Exception as e:
        return f"Error: {str(e)}"

def main():
    df = pd.read_excel('test.xlsx')
    r = sr.Recognizer()
    r.energy_threshold = 200
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        row = random.randint(0, df.index.size-1)
        column = random.randint(0, df.columns.size-1)
        
        value = choice_voice(r,source,'value')
        # df[row][column] = value
        df.loc[row,column] = value
        print(f"[{row}][{column}] update to {value}")    
        df.to_excel('test.xlsx', index=False)
    exportExcelToGooglesheet.export_program()
        
if __name__ == "__main__":
    main()