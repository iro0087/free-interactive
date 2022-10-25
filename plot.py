import matplotlib.pyplot as plt

from openpyxl import load_workbook

w = load_workbook("data.xlsx")

sheet = w.active

fichier = str(input("Name of the image file?"))

titre = str(input("What is the name of the graph? (skippable)"))

titrey = str(input("Name of y ?(skippable)"))

titrex = str(input("Name of x ?(skippable)"))

gg = str(input("Put the grid? (y/n)"))

t = 2

t2 = 1

lx = []

ly = []

ly2 = []

colorl = []

while sheet.cell(row=2, column=t2).value != None:

    if t2 == 1:

        lx.append(sheet.cell(row=t, column=t2).value)

    if t2 > 1:

        if type(sheet.cell(row=t, column=t2).value) != str:

            ly.append(sheet.cell(row=t, column=t2).value)

        else:

            colorl.append(sheet.cell(row=t, column=t2).value)

    if sheet.cell(row=t + 1, column=t2).value == None:

        t2 += 1

        t = 1

    t += 1

t = 0

maxy = 0

print(ly)

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

t2 = 0

pdp = 0

while t < len(ly) / len(lx):

    ly2 = []

    while t2 < len(lx):

        ly2.append(ly[pdp + t2])

        t2 += 1

    t2 = 0

    plt.scatter(lx, ly2, marker="*", color=colorl[t])

    pdp = pdp + len(lx)

    t += 1

plt.title(titre)

plt.ylabel(titrey)

plt.xlabel(titrex)

plt.show()
