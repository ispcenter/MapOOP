'''
MapOOP v0.1
Development started 2020/06/18
Targets:
a) convenient and fast plotting of tkr characteristics	- надо сделать автонастраиваемые оси и выбор файлов не через код
b) possibility of their combination								+
c) aproximation of experimental data							-
d) randomazer															-
e) plotting KPDk levels												-
f) livetime d) by changing data									-
'''

# --------------------------Head------------------------

from matplotlib.ticker import (MultipleLocator, FormatStrFormatter, AutoMinorLocator)
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import load_workbook

class Compressor:
	'''Prototype of a compressor graf
	1) Receive dataset
	2) Get data from it
	3) Minimize equation using data'''

	def __init__(self, dataset, column):
		self.dataset = dataset
		self.column = column
		self.data  = np.array([(row[column].value) for row in self.dataset if row[column].value != None])

		self.a = 0
		self.b = 0
		self.c = 0
		self.d = 0
		self.e = 0

		self.x = 0
		self.y = 0

	def aproximation5(self):
		print(self.data)

class Turbine:
	'''Prototype of a compressor graf
	1) Receive dataset
	2) Get data from it
	3) Minimize equation using data'''

	def __init__(self, dataset, column):
		self.dataset = dataset
		self.column = column
		self.data  = np.array([(row[column].value) for row in self.dataset if row[column].value != None])

		self.a = 0
		self.b = 0
		self.c = 0

		self.x = 0
		self.y = 0

	def aproximation2(self):
		print(self.data)

class Uk2:
	'''Prototype of a vetka
	1) Receive dataset
	2) Activates all data grafs'''

	def __init__(self, ws, label, marker, firstRow, lastRow):

		self.ws = ws
		self.label = label
		self.marker = marker
		self.firstRow = firstRow
		self.lastRow = lastRow
		self.data = self.ws[f'A{self.firstRow}':f'I{self.lastRow}']

		self.Gv  = np.array([(row[1].value) for row in self.data  if row[1].value != None])
		self.Pt  = np.array([(row[4].value) for row in self.data  if row[4].value != None])

		self.Pk = Compressor(self.data, 2)
		self.KPDk = Compressor(self.data, 3)

		self.KPDt = Turbine(self.data, 6)
		self.Gr = Turbine(self.data, 5)
		self.UCo = Turbine(self.data, 7)
		self.mft = Turbine(self.data, 8)

class MapData:
	"""Prototype of a mapData
	1) Receive file and color
	2) Activates all uk2
	3) Activates pompaz
	4) KPDk levels ???"""

	def __init__(self, file, color):
		# color stup
		self.color = color

		# data setup
		self.file = file
		self.wb = load_workbook(filename=file)
		self.ws = self.wb.active

		# all uk2 activation
		self.u200 = Uk2(self.ws, 200, '1', 40, 57)
		self.u250 = Uk2(self.ws, 250, '^', 59, 76)
		self.u300 = Uk2(self.ws, 300, 's', 78, 95)
		self.u350 = Uk2(self.ws, 350, 'D', 97, 114)
		self.u400 = Uk2(self.ws, 400, 'x', 116, 133)
		self.u450 = Uk2(self.ws, 450, 'o', 135, 152)
		self.u500 = Uk2(self.ws, 500, '+', 154, 171)
		self.u550 = Uk2(self.ws, 550, 'v', 173, 190)
		self.u600 = Uk2(self.ws, 600, '2', 192, 209)

		self.uList = [self.u200, self.u250, self.u300, self.u350, self.u400, self.u450, self.u500, self.u550, self.u600]

		# get pompaz coordinates
		self.xpomp = [u.Gv[-1] for u in self.uList if len(u.Gv) != 0]
		self.ypomp = [u.Pk.data[-1] for u in self.uList if len(u.Pk.data) != 0]

# --------------------------Config--------------------------
file_1 = 'Garrett.xlsx'
file_2 = 'Garrett2.xlsx'
file_3 = 'file3'
file_4 = 'file4'

color_1 = 'black'
color_2 = 'red'
color_3 = 'blue'
color_4 = 'green'

markersize = 4

# ----------------------Create the map-----------------------
map1 = MapData(file_1, color_1)
map2 = MapData(file_2, color_2)

# ----------------------- Plot the map-----------------------

# structure of a figure
fig = plt.figure('Зависимость Тк от Gв.пр')
ax1 = plt.subplot2grid((3, 2), (0, 0), rowspan=2)
ax2 = plt.subplot2grid((3, 2), (2, 0))
ax3 = plt.subplot2grid((4, 2), (0, 1))
ax4 = plt.subplot2grid((4, 2), (1, 1))
ax5 = plt.subplot2grid((4, 2), (2, 1))
ax6 = plt.subplot2grid((4, 2), (3, 1))

# axes settings - Pk
ax1.xaxis.set_major_locator(MultipleLocator(0.1))
ax1.xaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax1.xaxis.set_minor_locator(MultipleLocator(0.02))
ax1.set_xticklabels([])

ax1.yaxis.set_major_locator(MultipleLocator(0.4))
ax1.yaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax1.yaxis.set_minor_locator(MultipleLocator(0.2))

# axes settings - KPDk
ax2.xaxis.set_major_locator(MultipleLocator(0.1))
ax2.xaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax2.xaxis.set_minor_locator(MultipleLocator(0.02))

ax2.yaxis.set_major_locator(MultipleLocator(0.1))
ax2.yaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax2.yaxis.set_minor_locator(MultipleLocator(0.02))

# axes settings - mft
ax3.xaxis.set_major_locator(MultipleLocator(0.2))
ax3.xaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax3.xaxis.set_minor_locator(MultipleLocator(0.05))

ax3.yaxis.set_major_locator(MultipleLocator(2))
ax3.yaxis.set_major_formatter(FormatStrFormatter('%.0f'))
ax3.yaxis.set_minor_locator(MultipleLocator(1))

# axes settings - UCo
ax4.xaxis.set_major_locator(MultipleLocator(0.2))
ax4.xaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax4.xaxis.set_minor_locator(MultipleLocator(0.05))

ax4.yaxis.set_major_locator(MultipleLocator(0.05))
ax4.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
ax4.yaxis.set_minor_locator(MultipleLocator(0.01))

# axes settings - Gr
ax5.xaxis.set_major_locator(MultipleLocator(0.2))
ax5.xaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax5.xaxis.set_minor_locator(MultipleLocator(0.05))

ax5.yaxis.set_major_locator(MultipleLocator(0.5))
ax5.yaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax5.yaxis.set_minor_locator(MultipleLocator(0.1))

# axes settings - KPDt
ax6.xaxis.set_major_locator(MultipleLocator(0.2))
ax6.xaxis.set_major_formatter(FormatStrFormatter('%.1f'))
ax6.xaxis.set_minor_locator(MultipleLocator(0.1))

ax6.yaxis.set_major_locator(MultipleLocator(0.1))
ax6.yaxis.set_major_formatter(FormatStrFormatter('%.2f'))
ax6.yaxis.set_minor_locator(MultipleLocator(0.02))

# grid settings
ax1.grid(b=True, which='major', axis='both', color='#000000', linestyle='dotted', linewidth=1)
ax1.grid(b=True, which='minor', axis='both', color='#BDBDBD', linestyle='dotted', linewidth=1)
ax2.grid(b=True, which='major', axis='both', color='#000000', linestyle='dotted', linewidth=1)
ax2.grid(b=True, which='minor', axis='both', color='#BDBDBD', linestyle='dotted', linewidth=1)

ax3.grid(b=True, which='major', axis='both', color='#000000', linestyle='dotted', linewidth=1)
ax3.grid(b=True, which='minor', axis='both', color='#BDBDBD', linestyle='dotted', linewidth=1)
ax4.grid(b=True, which='major', axis='both', color='#000000', linestyle='dotted', linewidth=1)
ax4.grid(b=True, which='minor', axis='both', color='#BDBDBD', linestyle='dotted', linewidth=1)
ax5.grid(b=True, which='major', axis='both', color='#000000', linestyle='dotted', linewidth=1)
ax5.grid(b=True, which='minor', axis='both', color='#BDBDBD', linestyle='dotted', linewidth=1)
ax6.grid(b=True, which='major', axis='both', color='#000000', linestyle='dotted', linewidth=1)
ax6.grid(b=True, which='minor', axis='both', color='#BDBDBD', linestyle='dotted', linewidth=1)

# margins/paddings
fig.subplots_adjust(left=0.04, bottom=0.06, right=0.99, top=0.97, wspace=0.09, hspace=0)

# plots
ax1.set_title('Характеристика компрессора')
ax1.set_ylabel('Пк')
ax1.plot(
	map1.u200.Gv, map1.u200.Pk.data, map1.u200.marker, map1.u200.Gv, map1.u200.Pk.data, '-',
	map1.u250.Gv, map1.u250.Pk.data, map1.u250.marker, map1.u250.Gv, map1.u250.Pk.data, '-',
	map1.u300.Gv, map1.u300.Pk.data, map1.u300.marker, map1.u300.Gv, map1.u300.Pk.data, '-',
	map1.u350.Gv, map1.u350.Pk.data, map1.u350.marker, map1.u350.Gv, map1.u350.Pk.data, '-',
	map1.u400.Gv, map1.u400.Pk.data, map1.u400.marker, map1.u400.Gv, map1.u400.Pk.data, '-',
	map1.u450.Gv, map1.u450.Pk.data, map1.u450.marker, map1.u450.Gv, map1.u450.Pk.data, '-',
	map1.u500.Gv, map1.u500.Pk.data, map1.u500.marker, map1.u500.Gv, map1.u500.Pk.data, '-',
	map1.u550.Gv, map1.u550.Pk.data, map1.u550.marker, map1.u550.Gv, map1.u550.Pk.data, '-',
	map1.u600.Gv, map1.u600.Pk.data, map1.u600.marker, map1.u600.Gv, map1.u600.Pk.data, '-',
	map1.xpomp, map1.ypomp, c=map1.color, markersize = markersize)
ax1.plot(
	map2.u200.Gv, map2.u200.Pk.data, map2.u200.marker, map2.u200.Gv, map2.u200.Pk.data, '-',
	map2.u250.Gv, map2.u250.Pk.data, map2.u250.marker, map2.u250.Gv, map2.u250.Pk.data, '-',
	map2.u300.Gv, map2.u300.Pk.data, map2.u300.marker, map2.u300.Gv, map2.u300.Pk.data, '-',
	map2.u350.Gv, map2.u350.Pk.data, map2.u350.marker, map2.u350.Gv, map2.u350.Pk.data, '-',
	map2.u400.Gv, map2.u400.Pk.data, map2.u400.marker, map2.u400.Gv, map2.u400.Pk.data, '-',
	map2.u450.Gv, map2.u450.Pk.data, map2.u450.marker, map2.u450.Gv, map2.u450.Pk.data, '-',
	map2.u500.Gv, map2.u500.Pk.data, map2.u500.marker, map2.u500.Gv, map2.u500.Pk.data, '-',
	map2.u550.Gv, map2.u550.Pk.data, map2.u550.marker, map2.u550.Gv, map2.u550.Pk.data, '-',
	map2.u600.Gv, map2.u600.Pk.data, map2.u600.marker, map2.u600.Gv, map2.u600.Pk.data, '-',
	map2.xpomp, map2.ypomp, c=map2.color, markersize = markersize)

ax2.set_xlabel('Gв.пр')
ax2.set_ylabel('КПД')
ax2.plot(
	map1.u200.Gv, map1.u200.KPDk.data, map1.u200.marker, map1.u200.Gv, map1.u200.KPDk.data, '-',
	map1.u250.Gv, map1.u250.KPDk.data, map1.u250.marker, map1.u250.Gv, map1.u250.KPDk.data, '-',
	map1.u300.Gv, map1.u300.KPDk.data, map1.u300.marker, map1.u300.Gv, map1.u300.KPDk.data, '-',
	map1.u350.Gv, map1.u350.KPDk.data, map1.u350.marker, map1.u350.Gv, map1.u350.KPDk.data, '-',
	map1.u400.Gv, map1.u400.KPDk.data, map1.u400.marker, map1.u400.Gv, map1.u400.KPDk.data, '-',
	map1.u450.Gv, map1.u450.KPDk.data, map1.u450.marker, map1.u450.Gv, map1.u450.KPDk.data, '-',
	map1.u500.Gv, map1.u500.KPDk.data, map1.u500.marker, map1.u500.Gv, map1.u500.KPDk.data, '-',
	map1.u550.Gv, map1.u550.KPDk.data, map1.u550.marker, map1.u550.Gv, map1.u550.KPDk.data, '-',
	map1.u600.Gv, map1.u600.KPDk.data, map1.u600.marker, map1.u600.Gv, map1.u600.KPDk.data, '-',
	c=map1.color, markersize = markersize)

ax3.set_title('Характеристика турбины')
ax3.set_ylabel('mft')
ax3.plot(
	map1.u200.Pt, map1.u200.mft.data, map1.u200.marker, map1.u200.Pt, map1.u200.mft.data, '-',
	map1.u250.Pt, map1.u250.mft.data, map1.u250.marker, map1.u250.Pt, map1.u250.mft.data, '-',
	map1.u300.Pt, map1.u300.mft.data, map1.u300.marker, map1.u300.Pt, map1.u300.mft.data, '-',
	map1.u350.Pt, map1.u350.mft.data, map1.u350.marker, map1.u350.Pt, map1.u350.mft.data, '-',
	map1.u400.Pt, map1.u400.mft.data, map1.u400.marker, map1.u400.Pt, map1.u400.mft.data, '-',
	map1.u450.Pt, map1.u450.mft.data, map1.u450.marker, map1.u450.Pt, map1.u450.mft.data, '-',
	map1.u500.Pt, map1.u500.mft.data, map1.u500.marker, map1.u500.Pt, map1.u500.mft.data, '-',
	map1.u550.Pt, map1.u550.mft.data, map1.u550.marker, map1.u550.Pt, map1.u550.mft.data, '-',
	map1.u600.Pt, map1.u600.mft.data, map1.u600.marker, map1.u600.Pt, map1.u600.mft.data, '-',
	c=map1.color, markersize = markersize)

ax4.set_ylabel(r'$\frac{U_{т1}}{С_0}$')
ax4.plot(
	map1.u200.Pt, map1.u200.UCo.data, map1.u200.marker, map1.u200.Pt, map1.u200.UCo.data, '-',
	map1.u250.Pt, map1.u250.UCo.data, map1.u250.marker, map1.u250.Pt, map1.u250.UCo.data, '-',
	map1.u300.Pt, map1.u300.UCo.data, map1.u300.marker, map1.u300.Pt, map1.u300.UCo.data, '-',
	map1.u350.Pt, map1.u350.UCo.data, map1.u350.marker, map1.u350.Pt, map1.u350.UCo.data, '-',
	map1.u400.Pt, map1.u400.UCo.data, map1.u400.marker, map1.u400.Pt, map1.u400.UCo.data, '-',
	map1.u450.Pt, map1.u450.UCo.data, map1.u450.marker, map1.u450.Pt, map1.u450.UCo.data, '-',
	map1.u500.Pt, map1.u500.UCo.data, map1.u500.marker, map1.u500.Pt, map1.u500.UCo.data, '-',
	map1.u550.Pt, map1.u550.UCo.data, map1.u550.marker, map1.u550.Pt, map1.u550.UCo.data, '-',
	map1.u600.Pt, map1.u600.UCo.data, map1.u600.marker, map1.u600.Pt, map1.u600.UCo.data, '-',
	c=map1.color, markersize = markersize)

ax5.set_ylabel(r'$\ G_{г.пр.}$')
ax5.plot(
	map1.u200.Pt, map1.u200.Gr.data, map1.u200.marker, map1.u200.Pt, map1.u200.Gr.data, '-',
	map1.u250.Pt, map1.u250.Gr.data, map1.u250.marker, map1.u250.Pt, map1.u250.Gr.data, '-',
	map1.u300.Pt, map1.u300.Gr.data, map1.u300.marker, map1.u300.Pt, map1.u300.Gr.data, '-',
	map1.u350.Pt, map1.u350.Gr.data, map1.u350.marker, map1.u350.Pt, map1.u350.Gr.data, '-',
	map1.u400.Pt, map1.u400.Gr.data, map1.u400.marker, map1.u400.Pt, map1.u400.Gr.data, '-',
	map1.u450.Pt, map1.u450.Gr.data, map1.u450.marker, map1.u450.Pt, map1.u450.Gr.data, '-',
	map1.u500.Pt, map1.u500.Gr.data, map1.u500.marker, map1.u500.Pt, map1.u500.Gr.data, '-',
	map1.u550.Pt, map1.u550.Gr.data, map1.u550.marker, map1.u550.Pt, map1.u550.Gr.data, '-',
	map1.u600.Pt, map1.u600.Gr.data, map1.u600.marker, map1.u600.Pt, map1.u600.Gr.data, '-',
	c=map1.color, markersize = markersize)

ax6.set_xlabel(r'$\pi_t$')
ax6.set_ylabel('KPDt')
ax6.plot(
	map1.u200.Pt, map1.u200.KPDt.data, map1.u200.marker, map1.u200.Pt, map1.u200.KPDt.data, '-',
	map1.u250.Pt, map1.u250.KPDt.data, map1.u250.marker, map1.u250.Pt, map1.u250.KPDt.data, '-',
	map1.u300.Pt, map1.u300.KPDt.data, map1.u300.marker, map1.u300.Pt, map1.u300.KPDt.data, '-',
	map1.u350.Pt, map1.u350.KPDt.data, map1.u350.marker, map1.u350.Pt, map1.u350.KPDt.data, '-',
	map1.u400.Pt, map1.u400.KPDt.data, map1.u400.marker, map1.u400.Pt, map1.u400.KPDt.data, '-',
	map1.u450.Pt, map1.u450.KPDt.data, map1.u450.marker, map1.u450.Pt, map1.u450.KPDt.data, '-',
	map1.u500.Pt, map1.u500.KPDt.data, map1.u500.marker, map1.u500.Pt, map1.u500.KPDt.data, '-',
	map1.u550.Pt, map1.u550.KPDt.data, map1.u550.marker, map1.u550.Pt, map1.u550.KPDt.data, '-',
	map1.u600.Pt, map1.u600.KPDt.data, map1.u600.marker, map1.u600.Pt, map1.u600.KPDt.data, '-',
	c=map1.color, markersize = markersize)

plt.show()
# ---------------------------testRESULT--------------------------
# print(map1.u250.Pk.data)