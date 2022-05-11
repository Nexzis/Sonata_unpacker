import zipfile                      # Для распаковки
import tkinter
from tkinter import *
from tkinter import ttk
from _tkinter import TclError
import tkinter.filedialog           # Для диалога
import os, shutil                   # Для операций удаления, создания папок
import comtypes                     # pip install comtypes
from comtypes.client import CreateObject
from comtypes.persist import IPersistFile
from comtypes.shelllink import ShellLink

#pip install pyinstaller Для создания .exe


root = tkinter.Tk()                     #Создание окна
root.geometry('345x395')                #Размер окна
root.title("Sonata Unzipper")           #Заголовок окна
root.resizable(False, False)
f_top = LabelFrame(root, height=100, width=370, text='Target Location')
f_bot = LabelFrame(root, height=100, width=370, text='Actions')

f_bot1 = LabelFrame(f_bot, height=30, width=300)

size = 6
pic_path = '_Pics/'
image_delete_raw = tkinter.PhotoImage(file = pic_path + '5.png' )
image_delete = image_delete_raw.subsample(size,size)
image_update_raw = tkinter.PhotoImage(file = pic_path + '1.png' )
image_update = image_update_raw.subsample(size,size)
image_open_raw   = tkinter.PhotoImage(file = pic_path + '3.png' )
image_open   = image_open_raw.subsample(size,size)
image_PM_raw   = tkinter.PhotoImage(file = pic_path + '2.png' )
image_PM   = image_PM_raw.subsample(size,size)
image_L_raw   = tkinter.PhotoImage(file = pic_path + '4.png' )
image_L   = image_L_raw.subsample(size,size)
image_Ex_raw   = tkinter.PhotoImage(file = pic_path + '6.png' )
image_Ex   = image_Ex_raw.subsample(size,size)
image_SL_raw   = tkinter.PhotoImage(file = pic_path + '7.png' )
image_SL   = image_SL_raw.subsample(size,size)

btn = tkinter.Button(f_top, width = 140, height = 40, anchor='center', 
                      text = ' Choose & Extract' , 
                      image =  image_Ex, compound = LEFT, 
                      command = lambda: work())
btn1 = tkinter.Button(f_bot, width = 140, height = 40, anchor='center',
                      text = ' Start Loader' ,
                      image =  image_L, compound = LEFT,
                      command = lambda: startLoader(combo.get()))
btn2 = tkinter.Button(f_bot,  width = 140, height = 40, anchor='center',
                      text = ' Start PM' ,
                      image =  image_PM, compound = LEFT,
                      command = lambda: startPM(combo.get()))
btn3 = tkinter.Button(f_bot1, width = 40, height = 40, anchor='center',
                      text = ' Open Folder' ,
                      image =  image_open, 
                      command = lambda: openFolder(combo.get()))
btn4 = tkinter.Button(f_bot1, width = 40, height = 40, anchor='center',
                      text = ' Update' ,
                      image =  image_update,
                      command = lambda: dirContent(os.path.abspath(os.path.dirname(__file__))))  
btn5 = tkinter.Button(f_bot1, width = 40, height = 40, anchor='center',
                      text = ' Delete Folder' ,
                      image =  image_delete,
                      command = lambda: delete(combo.get() , inputBox.get() ))
btn6 = tkinter.Button(f_bot, width = 140, height = 40, anchor='center',
                      text = ' Stop Loaders' ,
                      image =  image_SL, compound = LEFT,
                      command = lambda: stopLoaders())

inputBox = tkinter.Entry(f_top, 
                         width = 50)
inputBox.insert(0, os.path.abspath(os.path.dirname(__file__)))             

label1 = tkinter.Label(f_top,
                       text='')
combo = ttk.Combobox(f_bot, 
                         width = 46)

#-----------------------------------
ExtractPath = inputBox.get()

def delete(folder, path):
    file_path = os.path.join(path, folder)
    try:
        shutil.rmtree(file_path)
    except OSError as e:
        label1.configure(text =  'Failed. Loader or PM still worked. Close it and repeat')
        pass
    dirContent(path)

def openFolder(path):
    os.system(f'start {os.path.realpath(os.path.join(inputBox.get(), path))}')

def startPM(path):
    os.chdir(os.path.realpath(os.path.join(inputBox.get(), path)))
    os.startfile(os.path.join(inputBox.get(), path) + r'\ProjectManager.exe')

def startLoader(path):
    os.chdir(os.path.realpath(os.path.join(inputBox.get(), path)))
    os.startfile(os.path.join(inputBox.get(), path) + r'\Loader.lnk')

def stopLoaders():
    os.system("TASKKILL /F /IM " + "Loader.exe")

def createFolderForTiff(folder, path):
    if not os.path.exists(os.path.join(path, folder)):     #if not (os.path.exists(path+ '\' +folder)):
        os.makedirs(os.path.join(path, folder))

def createLnk(path):
    s = CreateObject(ShellLink)
    s.SetPath(path + r'\Loader.exe')
    s.SetArguments('-daemon=10000')
    s.SetWorkingDirectory(path + r'\Loader.exe')
    p = s.QueryInterface(IPersistFile)
    p.Save(path + r'\Loader.lnk', True)

def dirContent(path):
    #content = [f for f in os.listdir(path) if not os.path.isfile(os.path.join(path, f))]
    content_all = []
    content_sonata = []
    for f in os.listdir(path):
        if not os.path.isfile(os.path.join(path, f)):
            content_all.append(f)
    for g in content_all:
        if (os.path.exists(os.path.join(path, g,'Loader.exe')) and os.path.exists(os.path.join(path, g,'ProjectManager.exe'))):
            content_sonata.append(g)
        
    combo.configure(values = content_sonata)
    try:
        combo.current(0)
    except TclError:
        pass
    return content_sonata

def names(tag):
    if tag.find('Sonata-') != -1 and tag.find('_WINDOWS-X86_') != -1:
        tag = tag.split('Sonata-')[1]                 # Отрезаем всё после "Sonata-"
        tag = tag[0:3] + '_' + tag.split('_r')[1]     # Вытаскиваем версию и соединяем с символами после "_r"
        tag = tag.split('.')[0]                       # Убираем расширение
    elif tag.find('Sonata_') != -1 and tag.find('_WINDOWS-X86_') != -1:
        tag = tag.split('_r')[1]     # Вытаскиваем версию и соединяем с символами после "_r"
        tag = tag.split('.')[0]                       # Убираем расширение
    else:
        tag = tag[tag.rfind(r"/") +1 :-4]
    return tag
    #pass   # Заменить на всплывающее окно "Не соната"   

def work():
    label1.configure(text =  'Choose file')  
    zipF = zipfile.ZipFile(tkinter.filedialog.askopenfile().name)
    #label1.configure(text =  str(zipF.filename))        # Отдаёт имя архива с расширением
    #label1.configure(text =  str(zipF.namelist()))      # Отдаёт имя содержимого архива , если указать [0] то первый элемент итд

    '''
    Label1 выдает текст формата
    'C:/Users/admin/Downloads/Sonata-1-3_WINDOWS-X86_20200420_r9897.zip'
    '''
    tag = str(zipF.filename)                      # Получаем строку вида Label 1 

    tag = names(tag) 
    
    ExtractPath = inputBox.get()
    Folder = tag
    
    # label1.configure(text =  'Processing...')  # Пропускает данный вывод информации
    delete(Folder, ExtractPath)                             # Удаление папки    
 
    createFolderForTiff(Folder, ExtractPath)                # Создание папки 
    
    ExtractPathfull = os.path.join(ExtractPath, Folder)     # Изменение пути

    zipF.extractall(ExtractPathfull)                        # Распаковка по этому пути
   
    createLnk(ExtractPathfull)                              # Создания ярлыка
    
    dirContent(ExtractPath)                                 # получение содержимого для всплывающего меню
    
    label1.configure(text =  'Done. Path: ' + str(ExtractPathfull))
    return ExtractPathfull

    
dirContent(inputBox.get())

#-----------------------------------
f_top.pack(side = TOP, padx=5, pady=5, ipadx=5, ipady=5)
f_bot.pack(side = TOP, padx=5, pady=5, ipadx=5, ipady=5)
inputBox.pack(padx = 1, pady = 1)
btn.pack(padx = 1, pady = 1)
label1.pack(padx = 1, pady = 1)
combo.pack(side = TOP, padx = 1, pady = 1)
f_bot1.pack(side = TOP, padx=1, pady=1, ipadx=1, ipady=1)
btn6.pack(side = TOP, padx = 2, pady = 1)
btn1.pack(side = TOP, padx = 2, pady = 1)
btn2.pack(side = TOP, padx = 2, pady = 1)
btn3.pack(side = LEFT, padx = 1, pady = 1 )
btn4.pack(side = LEFT, padx = 1, pady = 1 )
btn5.pack(side = LEFT, padx = 1, pady = 1 )



root.mainloop()

