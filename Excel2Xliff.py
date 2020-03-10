#!/usr/bin/python
# -*- coding: utf-8 -*-

#### DEBUGGING ####
# import pdb; pdb.set_trace()

#### SCRIPT VERSION ####
## 0.2
######## LOAD MODULES ########
import subprocess
import sys

try:
    from tkinter import *
    from tkinter.filedialog import *
    from tkinter import font
    from ttkthemes import ThemedStyle
    import pandas as pd
    import xlrd
    import os
    import getpass
    import time
    print("Your Python version is:", sys.version)
    print("\nAll modules are loaded!\n")
except ImportError:
    subprocess.call([sys.executable, "-m", "pip", "install", 'pandas'])
    subprocess.call([sys.executable, "-m", "pip", "install", 'xlrd'])
    subprocess.call([sys.executable, "-m", "pip", "install", 'ttkthemes'])
    print("\nAll modules are being installed...")
finally:
    from tkinter import *
    from tkinter.filedialog import *
    from tkinter import font
    from ttkthemes import ThemedStyle
    import pandas as pd
    import xlrd
    import os
    import getpass
    import time
    print("\nContinue...")
    time.sleep(2)

######## WORKING DIRECTORY ########
username = getpass.getuser()
os.chdir('/Users/' + username + '/Desktop')
time.sleep(2)

################################################################################################################
######## TKINTER ########
window = Tk()

#### WINDOW ####
window.title("Excel2XLIFF App")
window.geometry('650x200')
window.resizable(height=False, width=False)

#### THEME ####
style = ThemedStyle(window)
style.set_theme("aqua")

#### FONT ####
appHighlightFont = font.Font(family='Arial', size=12)
window.option_add("*Font", appHighlightFont)

######## CONTENT ########
lbl = Label(window, text="")
lbl.grid(column=0, row=0, sticky=E)

#### LABEL ####
lbl1 = Label(window, text="Enter Excel file path here", font="Arial 14 bold")
lbl1.grid(column=0, row=0, sticky=W, padx=15, pady=25)

lbl2 = Label(window, text="Enter original XLIFF file path here", font="Arial 14 bold")
lbl2.grid(column=0, row=1, sticky=W, padx=15)

#### INPUT ####
input1 = Entry(window,width=50)
input1.grid(column=1, row=0)

input2 = Entry(window,width=50)
input2.grid(column=1, row=1)

#### RADIO BUTTONS ####
check_language = StringVar()
check_language.set("pt-BR")
rbutton_pt = Radiobutton(window,text="Portuguese",variable=check_language,value="pt-BR", font=('Arial', 14)).grid(column=0,row=4,sticky=E,pady=30)
rbutton_es = Radiobutton(window,text="Spanish",variable=check_language,value="es", font=('Arial', 14)).grid(column=1,row=4,sticky=W,pady=30)

################################################################################################################
#### CONVERTING FUNCTION ####
def translate_XLIFF():

    excel_file = input1.get()
    xliff_file = input2.get()
    language_input = check_language.get()

    ############### LOAD EXCEL FILE ###############
    data = pd.read_excel(excel_file)

    # Make sure all Data types are string
    data["Title"]= data["Title"].astype(str)
    data["Body"]= data["Body"].astype(str)
    data["Body_2"]= data["Body_2"].astype(str)
    data["Body_3"]= data["Body_3"].astype(str)
    data["Link"]= data["Link"].astype(str)

    # Fill NAs with 0
    data.fillna(0, inplace = True)

    # Load en-US XLIFF file
    contents = ""
    file = xliff_file
    with open(file) as f:
        for line in f.readlines():
            contents += line

    # Get original file name
    old_file_name = os.path.basename(file)

    ############### FILE NAME ###############
    # Create new file name
    lenght = len(old_file_name)
    old_file_name_index = contents.find(old_file_name)
    old_file_name_index = old_file_name_index - 10
    new_file_name_FINAL = old_file_name[:old_file_name_index] + language_input

    ############### LANGUAGE ###############
    language = language_input
    index_language = contents.find('target-language="')
    y_language = len('target-language="')
    x_language = index_language + y_language
    output_line = contents[:index_language] + contents[index_language:x_language] + language + contents[x_language:]

    # Save file
    with open(new_file_name_FINAL + '.xliff', 'w') as file:
        file.write(output_line)

    # Load file
    contents = ""
    file = new_file_name_FINAL + '.xliff'
    with open(file) as f:
        for line in f.readlines():
            contents += line

    ############### TITLE CONDITION ###############
    if language_input == 'pt-BR':
        data_title = data['Title'][1]
        title_EN = str(data['Title'][0])
        if data_title != 0 and title_EN in contents:
            # TITLE CONTENT HERE
            title = str(data_title)
            index_title = contents.find(title_EN)
            y_title = len(title_EN)
            x_title = index_title + y_title + len('**</source><target>')
            output_line = contents[:index_title] + contents[index_title:x_title] + '**' + title + '**' + contents[x_title:]
        pass
    elif language_input == 'es':
        data_title = data['Title'][2]
        title_EN = str(data['Title'][0])
        if data_title != 0 and title_EN in contents:
            # TITLE CONTENT HERE
            title = str(data_title)
            index_title = contents.find(title_EN)
            y_title = len(title_EN)
            x_title = index_title + y_title + len('**</source><target>')
            output_line = contents[:index_title] + contents[index_title:x_title] + '**' + title + '**' + contents[x_title:]
        pass
    else:
        print('Something went wrong!')

    ############### SAVE FILE WITH TITLE ###############
    # Save file
    with open(new_file_name_FINAL + '.xliff', 'w') as file:
        file.write(output_line)

    # Load file
    contents = ""
    file = new_file_name_FINAL + '.xliff'
    with open(file) as f:
        for line in f.readlines():
            contents += line

    ############### BODY CONDITION ###############
    if language_input == 'pt-BR':
        data_body = data['Body'][1]
        body_EN = str(data['Body'][0])
        if data_body != 0 and body_EN in contents:
            # BODY CONTENT HERE
            body = str(data_body)
            index_body = contents.find(body_EN)
            y_body = len(body_EN)
            x_body = index_body + y_body + len('</source><target>')
            output_line = contents[:index_body] + contents[index_body:x_body] + data_body + contents[x_body:]
        pass
    elif language_input == 'es':
        data_body = data['Body'][2]
        body_EN = str(data['Body'][0])
        if data_body != 0 and body_EN in contents:
            # BODY CONTENT HERE
            body = str(data_body)
            index_body = contents.find(body_EN)
            y_body = len(body_EN)
            x_body = index_body + y_body + len('</source><target>')
            output_line = contents[:index_body] + contents[index_body:x_body] + data_body + contents[x_body:]
        pass
    else:
        print('Please insert either "pt-BR" or "es" only.')

    ############### SAVE FILE WITH BODY ###############
    # Save file
    with open(new_file_name_FINAL + '.xliff', 'w') as file:
        file.write(output_line)

    # Load file
    contents = ""
    file = new_file_name_FINAL + '.xliff'
    with open(file) as f:
        for line in f.readlines():
            contents += line

    ############### BODY_2 CONDITION ###############
    if language_input == 'pt-BR':
        data_body_2 = data['Body_2'][1]
        body_EN_2 = str(data['Body_2'][0])
        if data_body_2 != 0 and body_EN_2 in contents:
            # BODY_2 CONTENT HERE
            body_2 = str(data_body_2)
            index_body_2 = contents.find(body_EN_2)
            y_body_2 = len(body_EN_2)
            x_body_2 = index_body_2 + y_body_2 + len('</source><target>')
            output_line = contents[:index_body_2] + contents[index_body_2:x_body_2] + body_2 + contents[x_body_2:]
        pass
    elif language_input == 'es':
        data_body_2 = data['Body_2'][2]
        body_EN_2 = str(data['Body_2'][0])
        if data_body_2 != 0 and body_EN_2 in contents:
            # BODY_2 CONTENT HERE
            body_2 = str(data_body_2)
            index_body_2 = contents.find(body_EN_2)
            y_body_2 = len(body_EN_2)
            x_body_2 = index_body_2 + y_body_2 + len('</source><target>')
            output_line = contents[:index_body_2] + contents[index_body_2:x_body_2] + body_2 + contents[x_body_2:]
        pass
    else:
        print('Please insert either "pt-BR" or "es" only.')

    ############### SAVE FILE WITH BODY_2 ###############
    # Save file
    with open(new_file_name_FINAL + '.xliff', 'w') as file:
        file.write(output_line)

    # Load file
    contents = ""
    file = new_file_name_FINAL + '.xliff'
    with open(file) as f:
        for line in f.readlines():
            contents += line

    ############### LINK CONDITION ###############
    if language_input == 'pt-BR':
        data_link = data['Link'][1]
        link_EN = str(data['Link'][0])
        if data_link != 0 and link_EN in contents:
            # LINK CONTENT HERE
            link = str(data_link)
            index_link = contents.find(link_EN)
            y_link = len(link_EN)
            x_link = index_link + y_link + len('</source><target>')
            output_line = contents[:index_link] + contents[index_link:x_link] + link + contents[x_link:]
        pass
    elif language_input == 'es':
        data_link = data['Link'][2]
        link_EN = str(data['Link'][0])
        if data_link != 0 and link_EN in contents:
            # LINK CONTENT HERE
            link = str(data_link)
            index_link = contents.find(link_EN)
            y_link = len(link_EN)
            x_link = index_link + y_link + len('</source><target>')
            output_line = contents[:index_link] + contents[index_link:x_link] + link + contents[x_link:]
        pass
    else:
        print('Please insert either "pt-BR" or "es" only.')

    ############### REPLACE HTML CHARACTERS ###############
    # Replace special characters with HTML Entity Name
    output_line = output_line.replace(" & ", " &amp; ")
    output_line = output_line.replace("›", "&rsaquo;")

    ############### SAVE FILE WITH LINK ###############
    # Save file
    with open(new_file_name_FINAL + '.xliff', 'w') as file:
        file.write(output_line)

    print("\nFile saved on Desktop!")

    del output_line
    del contents

################################################################################################################

#### BUTTON ####
btn = Button(window, text="Convert Files", fg="black", command=translate_XLIFF, width=25, font=('Arial', 14, 'bold'), bd=8, relief=RAISED)
btn.grid(column=1, row=4, sticky=E, pady=30)


#### END ####
window.mainloop()
