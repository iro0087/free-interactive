#!usr/bin python3 

import matplotlib

matplotlib.use("TkAgg")

import matplotlib.pyplot as plt

from openpyxl import load_workbook

import os

from tkinter import Tk, Label, Button, Entry

import sys

lf = os.listdir(".")

def quit():

    sys.exit()

def fun():

    w = load_workbook("data.xlsx")

    sheet = w.active

    fichier = varfich.get()

    if lf.count(fichier + ".png") > 0:

        fichier = fichier + ".png"

    if lf.count(fichier + ".jpg") > 0:

        fichier = fichier + ".jpg"

    if lf.count(fichier) == 0:

        erf = Label(fen, text="file does not exist", fg="red")

        erf.pack()

    titre = vartitre.get()

    titrey = vary.get()

    titrex = varx.get()

    gg = vargrid.get()

    ygv = varm.get()

    if ygv != "w" and ygv != "m":

        Label(fen, text="Just put 'm' or 'w'", fg="red").pack()

    t = 2

    t2 = 1

    lx = []

    passe = 0

    ly = []

    ly2 = []

    colorl = []

    while sheet.cell(row=t, column=1).value != None:

        t += 1

    numx = t

    t = 1

    while sheet.cell(row=1, column=t2).value != None or t2 == 1:

        if t2 == 1 and t > 1:

            if sheet.cell(row=t, column=t2).value != None:

                lx.append(sheet.cell(row=t, column=t2).value)
                
                lx[-1] = str(lx[-1])

        if t2 > 1:

            if type(sheet.cell(row=t, column=t2).value) != str and sheet.cell(row=t, column=t2).value != None and passe == 0:

                ly.append(sheet.cell(row=t, column=t2).value)

                passe = 1

            else:

                if sheet.cell(row=t, column=t2).value != None and t > 1:

                    colorl.append(sheet.cell(row=t, column=t2).value)
                    
                    t2 += 1

                    t = 1

        if sheet.cell(row=t, column=t2).value == None:

            if sheet.cell(row=t, column=1).value != None and t2 > 1 and passe == 0 and t > 1:

                ly.append("#")

                passe = 1

            else:

                if t2 == 1:

                    t2 += 1

                    t = 1

        passe = 0

        t += 1

    really = len(ly)

    t2 = 0

    lx2 = []

    lx3 = []

    t = 0

    while t2 < len(ly):

        if ly[t2] != "#":

            lx2.append(lx[t])

        if t + 1 == len(lx):

            lx3.append(lx2)

            lx2 = []

            t = -1

        t += 1

        t2 += 1

    t2 = 0

    t = 0

    while t2 < len(ly):

        if ly[t2] == "#":

            ly.pop(t2)

            t2 = -1

        t2 += 1

    maxy = 0

    t = 0

    while t + 1 < len(ly):

        if ly[t] > ly[t - 1]:

            maxy = ly[t] + 1

        t += 1

    t = 0

    if maxy > len(lx):

        r_max = maxy

    else:

        r_max = len(lx)

    img = plt.imread(fichier)

    fig, ax = plt.subplots()

    ax.imshow(img, extent=[-1, r_max, 0, r_max])

    if gg == "y":

        plt.grid()

    pdp = 0

    t2 = 0

    lyref = []

    while t2 < really / len(colorl):

        lyref.append(0)

        t2 += 1

    plt.scatter(lx, lyref, marker="", color=colorl[t])

    t2 = 0

    while t < really / len(lx):

        ly2 = []

        while t2 < len(lx3[t]):

            ly2.append(ly[pdp + t2])

            t2 += 1

        t2 = 0

        phry = sheet.cell(row=1, column=t + 2).value

        while t2 < len(ly2):

            if ygv == "w":

                plt.text(lx3[t][t2], ly2[t2], phry, rotation=90, color=colorl[t])

            else:

                if ygv == "m":

                    plt.text(lx3[t][t2], ly2[t2], "*", color=colorl[t])

            t2 += 1

        t2 = 0

        pdp = pdp + len(lx3[t])

        t += 1

    t = 0

    t2 = 0

    while t < len(lx):

        yval = sheet.cell(row=1, column=t + 1).value

        t += 1

    plt.title(titre)

    plt.ylabel(titrey)

    plt.xlabel(titrex)

    Button(fen, text="RESTART", command=fun2, bg="green").pack()

    plt.show()

def fun2():

    fen = Tk()

    Label(fen, text="Make sure to match y, x values to your raw graph ;)", fg="green").pack()

    Label(fen, text="").pack()

    label = Label(fen, text="Please enter the name of the image file")

    label1 = Label(fen, text="What is the name of the graph? (skippable)")

    label2 = Label(fen, text="Name of y ?(skippable)")

    label3 = Label(fen, text="Name of x ?(skippable)")

    label4 = Label(fen, text="Put the grid? May increase accuracy. (y/n)")

    label5 = Label(fen, text="Do you want to display markers or words? (m/w)")

    label.pack()

    varfich = Entry(fen, width=30)

    vartitre = Entry(fen, width=30)

    vary = Entry(fen, width=30)

    varx = Entry(fen, width=30)

    vargrid = Entry(fen, width=30)

    varm = Entry(fen, width=30)

    varfich.pack()

    label1.pack()

    vartitre.pack()

    label2.pack()

    vary.pack()

    label3.pack()

    varx.pack()

    label4.pack()

    vargrid.pack()

    label5.pack()

    varm.pack()

    Button(fen, text="PROCEED", command=fun, bg="yellow").pack()

    Label(fen, text="").pack()

    Button(fen, text="TERMINATE", command=quit, bg="red").pack()

    fen.mainloop()

fen = Tk()

Label(fen, text="Make sure to match y, x values to your raw graph ;)", fg="green").pack()

Label(fen, text="").pack()

label = Label(fen, text="Please enter the name of the image file")

label1 = Label(fen, text="What is the name of the graph? (skippable)")

label2 = Label(fen, text="Name of y ?(skippable)")

label3 = Label(fen, text="Name of x ?(skippable)")

label4 = Label(fen, text="Put the grid? May increase accuracy. (y/n)")

label5 = Label(fen, text="Do you want to display markers or words? (m/w)")

label.pack()

varfich = Entry(fen, width=30)

vartitre = Entry(fen, width=30)

vary = Entry(fen, width=30)

varx = Entry(fen, width=30)

vargrid = Entry(fen, width=30)

varm = Entry(fen, width=30)

varfich.pack()

label1.pack()

vartitre.pack()

label2.pack()

vary.pack()

label3.pack()

varx.pack()

label4.pack()

vargrid.pack()

label5.pack()

varm.pack()

Button(fen, text="PROCEED", command=fun, bg="yellow").pack()

Label(fen, text="").pack()

Button(fen, text="TERMINATE", command=quit, bg="red").pack()

fen.mainloop()
