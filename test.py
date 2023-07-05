#This was made by someone who never typed a word of Python before. Input Sanitization? Pfft not here, doubt anyone other than me will ever use this. 
#No I will not clean up the various df variables to make more sense


import os 
import pandas as pd
import string 
from tabulate import tabulate
import sys
import colorama
from colorama import Fore

colorama.init() 

print('\nUpon entering the folder name, a report will automatically be generated for every individual report in the designated folder.\n\
You will need to screen capture the terminal if the raw information is your goal. You can do this by: \n\n 1. Maximixing the screen\n 2. Hitting the windows key + shift + s. It will bring up a screenshot tool.\n\
 3. Either use the cropping tool or save the whole screen.\n 4. It will save the image to the clipboard, open it up and save as a PNG\n 5. Profit\n\
*NOTE* Folder ' + Fore.RED + 'must created on DESKTOP. ' + Fore.RESET +'You can name it whatever you want, but the PATH variable is inflexible. Contact Tyler if something doesn\'t work.')

basePath = r'C:\Users\tspre\Desktop\WORK'  #Double check this if you aren't author

startRow = 0
endRow = 12
maxWidth = 50
totalCount = 0
excludePhrases = ['CORRUGATOR DOWNTIME' , 'TRANSFERSTATION DOWNTIME']
problemCounter = 0


Restart = 'yes'.lower()

while Restart == 'yes': 
    tempTarget= input("Enter the folder where target reports reside. Simply input the name of your folder and hit enter.\n\
No spaces, makes sure capitalization is right.\n\
Enter Name of folder " + Fore.RED + "HERE" + Fore.RESET + " or type " + Fore.RED + "EXIT" + Fore.RESET + " to exit: ")


    if any(specChar in string.punctuation for specChar in tempTarget):
        print(Fore.RED + "\nPlease reread instructions and try again I'm serious about there being no error checking. Almost none at least.\n\n" + Fore.RESET)
        continue

    if tempTarget.lower() == 'exit':
        print('\nEXITING PROGRAM')
        sys.exit() 
        
    targetFolder = '\\' + tempTarget
    totalPath = basePath + targetFolder

    if not os.path.exists(totalPath):
        print(Fore.RED + '\nPath not found, checking spelling and capitalization, TRY AGAIN\n' + Fore.RESET)
        continue

    excelFiles = [file for file in os.listdir(totalPath) if file.endswith('.xlsx')]

    for file in excelFiles:
        filePath = os.path.join(totalPath, file)
        reportHeader = pd.read_excel(filePath, sheet_name='Sheet1', header=0, nrows=0)
        df = pd.read_excel(filePath, sheet_name='Sheet1', usecols=['SHUTTLE NUMBER OR TS', 'DESCRIPTION OF ISSUE '] , header=2)
        temp_df = df.loc[startRow:endRow].copy()
        temp_df['DESCRIPTION OF ISSUE '] = temp_df['DESCRIPTION OF ISSUE '].apply(lambda x: '\n'.join(x[i:i+maxWidth] for i in range(0, len(x), maxWidth)) if isinstance(x, str) and len(x) > maxWidth else x)
        #This was easy to figure out and totally not miserable
        formatted_df = temp_df.dropna(how='any')
        rows2Drop = formatted_df['SHUTTLE NUMBER OR TS'].isin(excludePhrases) | formatted_df['DESCRIPTION OF ISSUE '].isin(excludePhrases)
        final_df = formatted_df[~rows2Drop]
        problemCounter += len(final_df)
        
     
        
        if len(final_df) >= 1:
            print('\n\n', tabulate(reportHeader , headers='keys' , tablefmt = 'simple') , tabulate(final_df, headers='keys', tablefmt='grid' , colalign=['left' , 'left'], showindex=False))
            
        else: continue


    print(Fore.CYAN + '\nTOTAL MALFUNCTIONS: ' + Fore.RED, problemCounter)     
    Restart =input(Fore.RED + 'Restart program? Yes/No + Enter:  ' + Fore.RESET).lower()
    
    
    
