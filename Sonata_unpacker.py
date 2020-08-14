import zipfile                      # Для распаковки
import tkinter
import tkinter.filedialog           # Для диалога
import os, shutil                   # Для операций удаления, создания папок
#import re                          # Регулярные выражения
import comtypes                     # pip install comtypes
from comtypes.client import CreateObject
from comtypes.persist import IPersistFile
from comtypes.shelllink import ShellLink


root = tkinter.Tk()                     #Создание окна
root.geometry('500x100')                #Размер окна
root.title("Sonata Unzipper")           #Заголовок окна
frame = tkinter.Frame(root)
btn = tkinter.Button(frame)
inputBox = tkinter.Entry(frame, width = 50)
inputBox.insert(0, r'D:\#Work\Sonata')             # Изменять в зависимости от того где хотите видеть папку
label = tkinter.Label(frame,text="Target Location:")
label1 = tkinter.Label(frame,text='')

#Mpath = r'D:\#Work\Sonata'


def delete(folder, path):
    #dir_path = folder    # r'D:\#Work\Sonata_1.3\test\Design_NU'
    file_path = os.path.join(path, folder)
    try:
        shutil.rmtree(file_path)
    except OSError as e:
        pass

def createFolderForTiff(folder, path):
    if not os.path.exists(os.path.join(path, folder)):     #if not (os.path.exists(path+ '\' +folder)):
        #os.chdir(path)
        #os.mkdir(folder)
        os.makedirs(os.path.join(path, folder))

# def createLnk(folder, path)
def createLnk(path):
    s = CreateObject(ShellLink)
    # s.SetPath(r'D:\#Work\Sonata_1.3\9897\Loader.exe')
    s.SetPath(path + r'\Loader.exe')
    s.SetArguments('-daemon=10000')
    # s.SetWorkingDirectory(r'D:\#Work\Sonata_1.3\9897\Loader.exe')
    s.SetWorkingDirectory(path + r'\Loader.exe')
    p = s.QueryInterface(IPersistFile)
    # p.Save(r"D:\#Work\Sonata_1.3\9897\Loader.lnk", True)
    p.Save(path + r'\Loader.lnk', True)

def names(tag):
    if tag.find('Sonata-') != -1 and tag.find('_WINDOWS-X86_') != -1:
        tag = tag.split('Sonata-')[1]                 # Отрезаем всё после "Sonata-"
        tag = tag[0:3] + '_' + tag.split('_r')[1]     # Вытаскиваем версию и соединяем с символами после "_r"
        tag = tag.split('.')[0]                       # Убираем расширение
    else:
        tag = tag[tag.rfind(r"/") +1 :-4]
    return tag
    #pass   # Заменить на всплывающее окно "Не соната"   
    
    
    
    
    
def work():
    zipF = zipfile.ZipFile(tkinter.filedialog.askopenfile().name)
    label1.configure(text =  str(zipF.filename))                # Отдаёт имя архива с расширением
    #label1.configure(text =  str(zipF.namelist()))             # Отдаёт имя содержимого архива , если указать [0] то первый элемент итд

    '''
    Label1 выдает текст формата
    'C:/Users/admin/Downloads/Sonata-1-3_WINDOWS-X86_20200420_r9897.zip'
    '''
    tag = str(zipF.filename)                      # Получаем строку вида Label 1 

    tag = names(tag) 
    
    ExtractPath = inputBox.get()
    Folder = tag
    
    delete(Folder, ExtractPath)        # Удаление папки   #delete(os.path.join(ExtractPath, Folder)) 
    
    createFolderForTiff(Folder, ExtractPath)   # Создание папки 
    
    ExtractPath = os.path.join(ExtractPath, Folder) # Изменение пути
    
    zipF.extractall(ExtractPath)    # Распаковка по этому пути
    
    createLnk(ExtractPath)
    
    #os.startfile(ExtractPath + r'\Loader.lnk')



btn['text'] = "Choose & Extract"
btn['command'] = work

label.pack()
inputBox.pack()
btn.pack()
label1.pack()
frame.pack()

root.mainloop()
