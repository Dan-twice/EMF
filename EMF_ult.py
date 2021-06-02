import tkinter as tk
from openpyxl import Workbook
from PIL import Image, ImageTk
import random as rd
import matplotlib.pyplot as plt
import numpy as np

plt.ion()


class DynamicUpdate():
    def on_launch(self):
        # Set up plot
        self.figure, self.ax = plt.subplots()
        self.lines, = self.ax.plot([], [], 'o')
        # Autoscale on unknown axis and known lims on the other
        self.ax.set_autoscaley_on(True)
        # Other stuff
        self.ax.grid()
        self.ax.set_title('Na vs R')
        self.ax.set_xlabel('Зовнішній опір — R, Ом')
        self.ax.set_ylabel('Корисна потужність — Na, Вт')

    def on_running(self, xdata, ydata):
        # Update data (with the new _and_ the old points)
        self.lines.set_xdata(xdata)
        self.lines.set_ydata(ydata)
        # Need both of these in order to rescale
        self.ax.relim()
        self.ax.autoscale_view()
        # We need to draw *and* flush
        self.figure.canvas.draw()
        self.figure.canvas.flush_events()

    def call(self, x, y, value):
        x.append(value)
        y.append(us_power)

        self.on_running(x, y)
        if len(x) >= 21:
            btn1['state'] = 'disable'
            btn2['state'] = 'disable'
            btn3['state'] = 'normal'
            btn3['cursor'] = 'hand2'


class Entry:
    def __init__(self):
        self.enter = tk.Entry(width=7, justify=tk.CENTER, font=('Times', '13'))

    def call(self, rw, cl):
        self.enter.grid(row=rw, column=cl)


class Button:
    def __init__(self):
        self.image2 = Image.open('pictures/r.png')
        self.photo2 = ImageTk.PhotoImage(self.image2)
        self.image2 = canvas.create_image(150, 235, anchor='nw', image=self.photo2)

    def ch_green(self, x, y, graph):
        value = float(ent.get())
        if not 1. <= value <= 20.:
            lb['text'] = 'Впроваджено не вірне число'
            return
        lb['text'] = ''
        btn1['state'] = 'disable'
        btn1['cursor'] = 'arrow'
        btn2['state'] = 'active'
        btn2['cursor'] = 'hand2'
        self.image2 = Image.open('pictures/g.png')
        self.photo2 = ImageTk.PhotoImage(self.image2)
        self.image2 = canvas.create_image(150, 235, anchor='nw', image=self.photo2)

        change_section()

        graph.call(x, y, value)

    def ch_red(self):
        btn1['state'] = 'active'
        btn2['state'] = 'disable'
        btn2['cursor'] = 'arrow'
        btn1['cursor'] = 'hand2'
        self.image2 = Image.open('pictures/r.png')
        self.photo2 = ImageTk.PhotoImage(self.image2)
        self.image2 = canvas.create_image(150, 235, anchor='nw', image=self.photo2)


def create_fr_line():
    ent1 = Entry()
    ent1.enter.insert(0, '№')
    ent1.enter['state'] = 'disable'
    ent1.call(c-1, 0)

    ent2 = Entry()
    ent2.enter.insert(0, 'R, Ом')
    ent2.enter['state'] = 'disable'
    ent2.call(c-1, 1)

    ent3 = Entry()
    ent3.enter.insert(0, 'I, А')
    ent3.enter['state'] = 'disable'
    ent3.call(c-1, 2)

    ent4 = Entry()
    ent4.enter.insert(0, 'U, В')
    ent4.enter['state'] = 'disable'
    ent4.call(c-1, 3)

    ent5 = Entry()
    ent5.enter.insert(0, 'Na, Вт')
    ent5.enter['state'] = 'disable'
    ent5.call(c-1, 4)

    ar_sheet.append(['№', 'R, Ом', 'I, А', 'U, В', 'Na, Вт'])


def change_section():
    global E, r, c, us_power

    ent1 = Entry()
    ent1.enter.insert(0, str(c))
    # ent1.enter['state'] = 'disable'
    ent1.call(c, 0)

    R = round(float(ent.get()), 2)
    ent2 = Entry()
    ent2.enter.insert(0, str(R))
    # ent2.enter['state'] = 'disable'
    ent2.call(c, 1)

    power = round(E / (r + R), 2)
    ent3 = Entry()
    ent3.enter.insert(0, str(power))
    # ent3.enter['state'] = 'disable'
    ent3.call(c, 2)

    tense = round(power * R, 2)
    ent4 = Entry()
    ent4.enter.insert(0, str(tense))
    # ent4.enter['state'] = 'disable'
    ent4.call(c, 3)

    us_power = round(tense * power, 2)
    ent5 = Entry()
    ent5.enter.insert(0, str(us_power))
    # ent5.enter['state'] = 'disable'
    ent5.call(c, 4)

    ar_sheet.append([c, R, power, tense, us_power])

    c += 1


def create_graph():
    resist = np.arange(0.5, 20.1, 0.1)
    power = np.array([E ** 2 * i / (r + i) ** 2 for i in resist])

    plt.figure(figsize=(5, 5), dpi=90)
    plt.subplot(111)
    plt.plot(resist, power)
    plt.title('Na vs R')
    plt.xlabel('Зовнішній опір — R, Ом')
    plt.ylabel('Корисна потужність — Na, Вт')
    plt.grid(True)


def show_theory():
    root2 = tk.Tk()

    lab7 = tk.Label(root2, font=('Times', '13'),
                    text='Програма зображує на графіку залежність \n'
                         'корисної потужності від зовнішнього опір,\n'
                         'в ланцюзі постійного струму, за формулою:\n'
                         '\n'
                         'P = I² * R\n'
                         '\n'
                         'I = {0} / (r + R)\n'
                         '\n'
                         'де, P - корисна потужність,\n'
                         'R - зовнішній опір,\n'
                         'r - внутрішній опір,\n'
                         '{0} — електрорушійна сила (ЕРС).\n'
                         '\n'
                         'Програма запитує R\n'
                         'на рекомендованому діапазоні: 1.0-20.0\n'
                    .format(chr(int('2130', 16))))
    lab7.pack(padx=15, pady=15)

    # root2.geometry('+{}+{}'.format(w1 // 2 - 670, h1 // 2 - 105))
    root2.resizable(False, False)
    root2.mainloop()


def save_sheet():
    wb = Workbook()
    ws = wb.active
    for x in ar_sheet:
        ws.append(x)
    ws.append(['{}={}'.format(chr(int('2130', 16)), E)])
    wb.save("Sheet.xlsx")


def save_txt():
    with open('sh_zero.txt', 'w') as f:
        for x in ar_sheet:
            for y in x:
                y = str(y)
                if y.isdigit() and len(y) < 4:
                    y += y + '0'
                f.write(f'{y}      ')
            f.write('\n')
        f.write(f'ЕРС={E}')
    f.close()


root = tk.Tk()
graph = DynamicUpdate()
graph.on_launch()
ar_sheet = []

counter = 1
j, i = 0, 0
for j in range(21):
    for i in range(5):
        e = 'ent' + str(counter)
        counter += 1
        e = Entry()
        e.call(j, i)

c = 1
E = rd.randint(10, 20)
r = round(rd.uniform(2., 5.), 1)
xarray = [0.5]
yarray = [E ** 2 * xarray[0] / (r + xarray[0]) ** 2]
us_power = 0

create_fr_line()

canvas = tk.Canvas(root, width=600, height=300, bg='#f0f0f0')  # rgb(240, 240, 240)

canvas.create_line(100, 50, 500, 50, width=3)
canvas.create_line(100, 250, 500, 250, width=3)
canvas.create_line(100, 50, 100, 250, width=3)
canvas.create_line(500, 50, 500, 250, width=3)
canvas.create_line(200, 50, 200, 150, width=3)
canvas.create_line(400, 50, 400, 150, width=3)
canvas.create_line(200, 150, 400, 150, width=3)

canvas.create_text(210, 30, text="{}={} В".format(chr(int('2130', 16)), E), font=('Times', '18', 'bold'))
canvas.create_text(310, 215, text="R=", font=('Times', '18', 'bold'))
canvas.create_text(420, 215, text="Ом", font=('Times', '18', 'bold'))
canvas.create_text(167, 215, text="K", font=('Times', '18', 'bold'))

image1 = Image.open('pictures/Volt.jpg')
photo1 = ImageTk.PhotoImage(image1)
image1 = canvas.create_image(258, 125, anchor='nw', image=photo1)
canvas.create_text(300, 150, text="V", font=('Times', '30'))

image2 = Image.open('pictures/A.png')
photo2 = ImageTk.PhotoImage(image2)
image2 = canvas.create_image(475, 105, anchor='nw', image=photo2)
canvas.create_text(500, 148, text="A", font=('Times', '30'))

image3 = Image.open('pictures/Resistor.gif')
photo3 = ImageTk.PhotoImage(image3)
image3 = canvas.create_image(325, 235, anchor='nw', image=photo3)

image4 = Image.open('pictures/r.png')
photo4 = ImageTk.PhotoImage(image4)
image4 = canvas.create_image(150, 235, anchor='nw', image=photo4)

image5 = Image.open('pictures/battary2.png')
photo5 = ImageTk.PhotoImage(image5)
image5 = canvas.create_image(275, 30, anchor='nw', image=photo5)

bt1 = Button()
ent = tk.Entry(root, width=6, justify=tk.CENTER, font=('Times', '16', 'bold'))
canvas.create_window(365, 215, window=ent)
btn1 = tk.Button(root, width=22, height=2, relief='groove', activebackground='#bfc900', bd=6, text='Замкнути',
                 font=('Times', '13'), cursor='hand2', command=lambda: bt1.ch_green(xarray, yarray, graph),
                 state='active')
btn2 = tk.Button(root, width=22, height=2, relief='groove', bd=6, activebackground='#bfc900',
                 text='Розімкнути', font=('Times', '13'), command=lambda: bt1.ch_red(), state='disable')
btn3 = tk.Button(root, width=22, height=2, relief='groove', bd=6, text='Графік', font=('Times', '13'),
                 activebackground='#bfc900', command=lambda: create_graph(), state='disable')
btn4 = tk.Button(root, width=22, height=2, relief='groove', bd=6, text='Теорія', font=('Times', '13'), cursor='hand2',
                 activebackground='#bfc900', command=lambda: show_theory())
lbf = tk.LabelFrame(root, text=' Таблиця в ', font=('Times', '12'), bd=4)
btn5 = tk.Button(lbf, width=15, height=1, relief='groove', bd=6, font=('Times', '11'), cursor='hand2',
                 activebackground='#bfc900', text='Excel', command=lambda: save_sheet())
btn6 = tk.Button(lbf, width=15, height=1, relief='groove', bd=6, font=('Times', '11'), cursor='hand2',
                 activebackground='#bfc900', text='TXT', command=lambda: save_txt())
lb = tk.Label(root, font=('Times', '14'), fg='red')


canvas.grid(row=0, column=5, rowspan=15, columnspan=3)
btn1.grid(row=14, column=5, rowspan=3)
btn2.grid(row=14, column=6, rowspan=3)
btn3.grid(row=17, column=5, rowspan=3)
btn4.grid(row=17, column=6, rowspan=3)
btn5.grid(ipady=5)
btn6.grid(ipady=5)
lbf.grid(ipady=4, row=14, column=7, rowspan=6)
lb.grid(row=20, column=5, columnspan=3)

root.resizable(False, False)
root.mainloop()
