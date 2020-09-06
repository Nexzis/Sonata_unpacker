import zipfile                      # Для распаковки
import tkinter
from tkinter import *
from tkinter import ttk
import tkinter.filedialog           # Для диалога
import os, shutil                   # Для операций удаления, создания папок
import comtypes                     # pip install comtypes
from comtypes.client import CreateObject
from comtypes.persist import IPersistFile
from comtypes.shelllink import ShellLink


root = tkinter.Tk()                     #Создание окна
root.geometry('400x260')                #Размер окна
root.title("Sonata Unzipper")           #Заголовок окна
f_top = LabelFrame(root, height=100, width=400, text='Target Location')
f_bot = LabelFrame(root, height=100, width=400, text='Actions')
f_top.pack(side = TOP, padx=5, pady=5, ipadx=5, ipady=5)
f_bot.pack(side = TOP, padx=5, pady=5, ipadx=5, ipady=5)
#frame = tkinter.Frame(root)
btn = tkinter.Button(f_top, 
                     text = 'Choose & Extract' , 
                     command = lambda: work())
btn1 = tkinter.Button(f_bot, 
                      text = 'Start Loader' , 
                      command = lambda: startLoader(combo.get()))
btn2 = tkinter.Button(f_bot, 
                      text = 'Start PM' , 
                      command = lambda: startPM(combo.get()))
btn3 = tkinter.Button(f_bot, 
                      text = 'Open Folder' , 
                      command = lambda: openFolder(combo.get()))
inputBox = tkinter.Entry(f_top, 
                         width = 50)
inputBox.insert(0, r'D:\#Work\Sonata')             # Изменять в зависимости от того где хотите видеть папку
#label = tkinter.Label(f_top,
#                      text="Target Location:")
label1 = tkinter.Label(f_top,
                       text='')
combo = ttk.Combobox(f_bot)


ExtractPath = inputBox.get()

def delete(folder, path):
    file_path = os.path.join(path, folder)
    try:
        shutil.rmtree(file_path)
    except OSError as e:
        pass

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

'''
Этот синтаксис работать не хочет.
    s.SetPath(os.path.join(path, r'\Loader.exe'))

    s.SetWorkingDirectory(os.path.join(path, r'\Loader.exe')) 

    p.Save(os.path.join(path, r'\Loader.exe'), True)
'''


def dirContent(path):
    content = [f for f in os.listdir(path) if not os.path.isfile(os.path.join(path, f))]
    
    #combo.configure(values = content, postcommand=createLnk)         # Удалить если не нужна
    combo.configure(values = content)
    combo.current(0)

def names(tag):
    if tag.find('Sonata-') != -1 and tag.find('_WINDOWS-X86_') != -1:
        tag = tag.split('Sonata-')[1]                 # Отрезаем всё после "Sonata-"
        tag = tag[0:3] + '_' + tag.split('_r')[1]     # Вытаскиваем версию и соединяем с символами после "_r"
        tag = tag.split('.')[0]                       # Убираем расширение
    else:
        tag = tag[tag.rfind(r"/") +1 :-4]
    return tag
    #pass   # Заменить на всплывающее окно "Не соната"   
    
def openFolder(path):
   #os.system(f"start D:\#Work\Sonata")
    os.system(f'start {os.path.realpath(os.path.join(inputBox.get(), path))}')

def startPM(path):
    os.startfile(os.path.join(inputBox.get(), path) + r'\ProjectManager.exe')

def startLoader(path):
    os.startfile(os.path.join(inputBox.get(), path) + r'\Loader.lnk')
    
def work():
    label1.configure(text =  'Start work')  
    zipF = zipfile.ZipFile(tkinter.filedialog.askopenfile().name)
    #label1.configure(text =  str(zipF.filename))        # Отдаёт имя архива с расширением
    #label1.configure(text =  str(zipF.namelist()))      # Отдаёт имя содержимого архива , если указать [0] то первый элемент итд

    '''
    Label1 выдает текст формата
    'C:/Users/admin/Downloads/Sonata-1-3_WINDOWS-X86_20200420_r9897.zip'
    '''
    tag = str(zipF.filename)                      # Получаем строку вида Label 1 

    tag = names(tag) 
    
    #ExtractPath = inputBox.get()
    Folder = tag
   
    delete(Folder, ExtractPath)                             # Удаление папки    
 
    createFolderForTiff(Folder, ExtractPath)                # Создание папки 
    
    ExtractPathfull = os.path.join(ExtractPath, Folder)     # Изменение пути

    zipF.extractall(ExtractPathfull)                        # Распаковка по этому пути
   
    createLnk(ExtractPathfull)                              # Создания ярлыка
    
    dirContent(ExtractPath)                                 # получение содержимого для всплывающего меню
    
    label1.configure(text =  ('Done, Path: ', str(ExtractPathfull)))
    return ExtractPathfull

    
dirContent(inputBox.get())


inputBox.pack()
btn.pack()
label1.pack()
#frame.pack()
combo.pack(side = TOP, padx = 1, pady = 1)
btn3.pack(side = TOP, padx = 1, pady = 1 )
btn1.pack(side = TOP, padx = 1, pady = 1)
btn2.pack(side = TOP, padx = 1, pady = 1)


root.mainloop()
