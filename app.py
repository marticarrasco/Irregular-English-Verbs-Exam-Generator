from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import random
import os
import shutil 
import openpyxl
import win32com.client
from pywintypes import com_error





window = Tk()

window.title("Irregular Verbs List Exam Generator")
window.geometry("493x577")
window.configure(bg = "#dcdcdc")


canvas = Canvas(
    window,
    bg = "#dcdcdc",
    height = 577,
    width = 493,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = f"assets/background.png")
background = canvas.create_image(
    248.5, 204.0,
    image=background_img)

entry0_img = PhotoImage(file = f"assets/img_textBox0.png")
entry0_bg = canvas.create_image(
    248.5, 404.0,
    image = entry0_img)

#MODELS NUM
entry0 = Entry(
    bd = 0,
    bg = "#89cff0",
    highlightthickness = 0)

entry0.place(
    x = 57.0, y = 383,
    width = 383.0,
    height = 40)




#=================== OUTPUT FOLDER
def open_folder():
    window.outputFolder = filedialog.askdirectory(title="Choose the destination Folder")
    
    global outputPath
    outputPath = window.outputFolder

    canvas.create_text(
        316.0, 303.0,
        text = outputPath,
        fill = "#000000",
        font = ("Rambla-Regular", int(7.0)))

img1 = PhotoImage(file = f"assets/img1.png")
b1 = Button(
    image = img1,
    borderwidth = 0,
    highlightthickness = 0,
    command = open_folder,
    relief = "flat")

b1.place(
    x = 45, y = 278,
    width = 132,
    height = 53)


#======================= INPUT FILES
def open_file():
    window.filepath = filedialog.askopenfilename(initialdir="/", title="Open your Excel")
    global inputPath
    inputPath = window.filepath

    canvas.create_text(
        316.0, 204.0,
        text = inputPath,
        fill = "#000000",
        font = ("Rambla-Regular", int(7.0)))

img2 = PhotoImage(file = f"assets/img2.png")
b2 = Button(
    image = img2,
    borderwidth = 0,
    highlightthickness = 0,
    command = open_file,
    relief = "flat")

b2.place(
    x = 45, y = 179,
    width = 132,
    height = 53)

def generate_exams():
    models_num = int(entry0.get())
    
    print(outputPath + r'/generated_exams')

    if not os.path.exists(outputPath + r'/generated_exams'):
            os.makedirs(outputPath + r'/generated_exams')
    else:
        shutil.rmtree(outputPath + r'/generated_exams')
        os.makedirs(outputPath + r'/generated_exams')


    for num in range(models_num):
        df = pd.read_excel(inputPath) 

        rows, columns = df.shape
        columns = columns - 1


        for row in range(rows):
            column = random.randint(0,columns)

            selected_cell = df.iat[row, column]

            if (column == 0):
                for x in range (1,4):
                    df.iat[row, x] = None
            
            if (column == 1):
                df.iat[row, 0] = None
                for x in range (2,4):
                    df.iat[row, x] = None
            
            if (column == 2):
                for x in range (0,2):
                    df.iat[row, x] = None
                df.iat[row, 3] = None
            
            if (column == 3):
                for x in range (0,3):
                    df.iat[row, x] = None
        print(df)

        outputPath_custom = outputPath + r'/generated_exams' + r'/verbs_exam_model_' + str(num) + '.xlsx'
        df.to_excel(outputPath_custom, sheet_name='Verbs', index = False)

    messagebox.showinfo("Success!", "Exams have been succesfully created!") 
       


#==================== GENERATE EXAMS
img0 = PhotoImage(file = f"assets/img0.png")
b0 = Button(
    image = img0,
    borderwidth = 0,
    highlightthickness = 0,
    command = generate_exams,
    relief = "flat")

b0.place(
    x = 138, y = 459,
    width = 220,
    height = 73)



window.resizable(False, False)
window.mainloop()




