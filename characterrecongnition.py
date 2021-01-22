from tkinter import *
from tkinter import filedialog
from PIL import Image, ImageTk
from aip import AipOcr
import os
import tkinter.messagebox


class SortWin(object):
    def __init__(self):
        self.root = Tk()
        self.root.title('文字识别系统')
        self.frame = Frame(self.root, bd=2, relief=SUNKEN)
        self.frame.grid_rowconfigure(0, weight=1)
        self.frame.grid_columnconfigure(0, weight=1)
        self.xscroll = Scrollbar(self.frame, orient=HORIZONTAL)
        self.xscroll.grid(row=1, column=0, sticky=E + W)
        self.yscroll = Scrollbar(self.frame)
        self.yscroll.grid(row=0, column=1, sticky=N + S)
        self.canvas = Canvas(self.frame, bd=0, width=500, height=400)
        self.canvas.grid(row=0, column=0, sticky=N + S + E + W)
        self.xscroll.config(command=self.canvas.xview)
        self.yscroll.config(command=self.canvas.yview)
        self.frame.pack(fill=BOTH, expand=1)
        self.picture_path = ''
        self.card_info = ''
        Label(self.root, text='识别结果:', width=10, height=4).pack(side=LEFT)
        self.card_text=Text(self.root, width=50, height=4)
        self.card_text.pack(side=LEFT)
        self.result = ''

        Button(self.root, text='选择图片', width=10, command=self.printcoords).pack(side=LEFT)
        Button(self.root, text='文字识别', width=10, command=self.identify_card).pack(side=LEFT)
        Button(self.root, text='结果存储', width=10, command=self.save_result).pack(side=LEFT)

    def save_result(self):
        filename = self.picture_path
        save_path = os.path.join(os.getcwd(), "识别结果")
        image_name = filename.split('/')[-1].split('.')[-2]
        file_name = '_'.join([image_name, 'result.txt'])
        result_name = '\\'.join([save_path, file_name])
        self.card_text.insert(END, '\n存储路径：%s' %result_name)
        with open(result_name, 'w') as f:
            f.write(self.card_info)
        tkinter.messagebox.showinfo("存储结果", "成功")
        self.result = result_name


    def printcoords(self):
        File = filedialog.askopenfilename(parent=self.root, initialdir="C:/",title='Choose an image.')
        self.picture_path = File
        filename = ImageTk.PhotoImage(Image.open(File))
        print(filename)
        self.canvas.image = filename  # <--- keep reference of your image
        self.canvas.create_image(0,0,anchor='nw',image=filename)
        self.card_text.delete('1.0', 'end')

    def identify_card(self):
        filename = self.picture_path
        print('*' * 20)
        print(filename)
        print('*' * 20)
        APP_ID = '11573829'
        API_KEY = 'c8EDnGZvYulV44y58U5zhtBx'
        SECRET_KEY = '78XrTt0gNaAxpI8fm1nRZgAPwlAeaZYA'
        aipOcr = AipOcr(APP_ID, API_KEY, SECRET_KEY)
        options = {
            'detect_direction': 'true',
            'language_type': 'CHN_ENG',
        }

        def get_file_content(filePath):
            with open(filePath, 'rb') as fp:
                return fp.read()

        result = aipOcr.basicAccurate(get_file_content(filename), options)
        print(result['words_result'])
        if len(result['words_result']) == 0:
            tkinter.messagebox.showerror("识别错误", "文字识别失败，请谅解，请另换一张。")
        else:
            word_result = list()
            for word in result['words_result']:
                word_result.append(word['words'])
            self.card_info = '\n'.join(word_result)


            print(self.card_info)
            self.card_text.insert(END,self.card_info)
            tkinter.messagebox.showinfo("识别结果",self.card_info)

    def show(self):
        self.root.mainloop()


