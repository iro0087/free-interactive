import matplotlib.pyplot as plt

from openpyxl import load_workbook

from tkinter import *

w = load_workbook("data.xlsx")

sheet = w.active

fichier = str(input("Name of the image file?"))

titre = str(input("What is the name of the graph? (skippable)"))

titrey = str(input("Name of y ?(skippable)"))

titrex = str(input("Name of x ?(skippable)"))

gg = str(input("Put the grid? (y/n)"))

ygv = str(input("Do you want to display markers or words? (markers/words)"))

t = 2

t2 = 1

lx = []

ly = []

ly2 = []

colorl = []

while sheet.cell(row=t, column=1).value != None:

    t += 1

numx = t 

t = 2

while sheet.cell(row=numx, column=t2).value != None or t2 == 1:

    if t2 == 1:

        lx.append(sheet.cell(row=t, column=t2).value)

    if t2 > 1:

        if type(sheet.cell(row=t, column=t2).value) != str and sheet.cell(row=t, column=t2).value != None:

            ly.append(sheet.cell(row=t, column=t2).value)

        else:

            if sheet.cell(row=t, column=t2).value != None:

                colorl.append(sheet.cell(row=t, column=t2).value)

    if sheet.cell(row=t + 1, column=t2).value == None:

        if sheet.cell(row=t, column=1).value != None and t2 > 1:

            ly.append("#")

            t += 1

        else:

            t2 += 1

            t = 1

            if sheet.cell(row=t + 1, column=t2).value == None and sheet.cell(row=1, column=t2).value != None:

                ly.append("#")

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

while t2 < really / len(lx):

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

    if ygv == "markers":

        plt.scatter(lx3[t], ly2, marker="*", color=colorl[t])

    else:

        phry = sheet.cell(row=1, column=t + 2).value

        print(phry)

        while t2 < len(ly2):

            plt.text(lx3[t][t2], ly2[t2], phry, rotation=90)

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

plt.show()
