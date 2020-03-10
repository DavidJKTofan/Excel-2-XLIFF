# Excel-2-XLIFF
Script that moves content from an Excel file (source) to a XLIFF file (target).

# Instructions
1. The Excel file needs to be in the same structure as the <a href="https://github.com/DavidJKTofan/Excel-2-XLIFF/blob/master/pendo-Translation-Content.xlsx" target="_blank">example</a> provided here (Title, Body, Body_2, Link).
2. The XLIFF file (when downloaded from <a href="https://www.pendo.io/" target="_blank">pendo.io</a>) holds the original source-text in it, which normally is "en-US" (English).
3. Open your Terminal, "cd" to the folder where the .py Script is located and write "python Excel2Xliff.py" or "python3 Excel2Xliff.py" to start the script. This will open a new window. You might have to set permissions to the script so it can access your files and your desktop, where it will save the new XLIFF files with the translations inside.
4. Insert the pathname of the Excel file and the XLIFF file, select the language you wish to translate, and click the button "Convert Files".
5. Your new files will be saved on your Desktop. 
6. Review your new files to see if it worked – hope so.

![Screenshot](https://raw.githubusercontent.com/DavidJKTofan/Excel-2-XLIFF/master/Screenshot.png)

### Tip
On a Mac you can get the pathname of a file with right-click on a file, then maintain the keyboard “option”-button pressed, and click on the “Copy FILENAME as Pathname” option in the menu.

# What I am using 
- macOS Catalina
- Python 3.8.1
- pandas 1.0.1
- tkinter 8.6
- ttkthemes 3.0.0
- xlrd 1.2.0

# Future improvements/ideas
- Overally improve and clean the Python code structure.
- More/Custom languages. Currently only "es" (Spanish) and "pt-BR" (Portuguese) work, but feel free to change them in the code itself.
- Add drag&drop functionality instead of adding the pathname of the files.
- Redesign the window to be more pleasant to the eye and user-friendly.
- ...

# About
Recently started learning about Python – this is a personal side-project which I am working on for a friend.
