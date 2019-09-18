
import wx
from os import walk, path, getcwd, listdir, mkdir, system
from shutil import copy
from PIL import Image, ImageEnhance
from PyPDF2 import PdfFileMerger, PdfFileWriter, PdfFileReader


class Ventana (wx.Frame):
	def __init__ (self, parent, title):
		no_resize = wx.DEFAULT_FRAME_STYLE ^(wx.MAXIMIZE_BOX | wx.CLOSE_BOX | wx.RESIZE_BORDER) #& ~ (wx.RESIZE_BORDER)
		wx.Frame.__init__ (self, parent, title=title, size=(554, 600), style=no_resize)
		self.panel = wx.Panel (self, style=wx.RAISED_BORDER)

		# font = wx.SystemSettings_GetFont(wx.SYS_SYSTEM_FONT)
		fuente = wx.Font (14, wx.DEFAULT, wx.NORMAL, wx.NORMAL)

		self.list_ctrl = wx.ListCtrl (self.panel, pos=(10, 80), size=(520, 200), style=wx.LC_REPORT | wx.BORDER_SUNKEN)
		self.list_ctrl.InsertColumn(0, 'File', width=200)

		self.texto = wx.StaticText (self.panel, wx.ID_ANY, "Choose the files to combine", pos=(140, 45), size=(160, -1))
		self.texto.SetFont (wx.Font (12, wx.DEFAULT, wx.NORMAL, wx.NORMAL))

		self.bt_ruta_origen = wx.Button (self.panel, label=". . .", pos=(340, 40), size=(40, 30))
		self.bt_ruta_origen.Bind (wx.EVT_BUTTON, lambda event: self.onOpenFile (event))

		# self.box = wx.StaticBox(self.panel, wx.ID_ANY, "", pos=(20, 290), size=(495, 105))

		# INSERT
		self.radio_insert = wx.RadioButton(self.panel, wx.ID_ANY, label='Insert   (Click on the file, which will be below)', pos=(40, 305))
		self.radio_insert.Bind(wx.EVT_RADIOBUTTON, lambda event: self.active_controls (event, 'insert'))
		# self.check_insert = wx.CheckBox (self.panel, wx.ID_ANY, pos=(40, 305), size=(20, -1))
		# self.texto_insert = wx.StaticText (self.panel, wx.ID_ANY, "Insertar (Clicka en el archivo que quedara debajo)", pos=(60, 305), size=(260, -1))
		self.TC_insert = wx.TextCtrl (self.panel, pos=(40, 330), size=(455, 25), style=wx.TE_READONLY)
		self.TC_insert.Disable()
		self.boton_insert = wx.Button (self.panel, label=". . .", pos=(415, 300), size=(80, 25))
		self.boton_insert.Disable()
		self.boton_insert.Bind (wx.EVT_BUTTON, lambda event: self.onOpenFile2 (event, 'insert'))

		# SUSTITUIR
		self.radio_sustituir = wx.RadioButton(self.panel, wx.ID_ANY, label='Replace   (Click on the file you want to replace)', pos=(40, 370))
		self.radio_sustituir.Bind(wx.EVT_RADIOBUTTON, lambda event: self.active_controls (event, 'sustituir'))
		# self.check_sustituir = wx.CheckBox (self.panel, wx.ID_ANY, pos=(40, 345), size=(20, -1))
		# self.texto_sustituir = wx.StaticText (self.panel, wx.ID_ANY, "Sustituir (Clicka en el archivo que quieres sustituir)", pos=(60, 345), size=(270, -1))
		self.TC_sustituir = wx.TextCtrl (self.panel, pos=(40, 395), size=(455, 25), style=wx.TE_READONLY)
		self.TC_sustituir.Disable()
		self.boton_sustituir = wx.Button (self.panel, label=". . .", pos=(415, 365), size=(80, 25))
		self.boton_sustituir.Disable()
		self.boton_sustituir.Bind (wx.EVT_BUTTON, lambda event: self.onOpenFile2 (event, 'sustituir'))

		# COMBINAR
		# self.check_combine = wx.CheckBox (self.panel, wx.ID_ANY, pos=(40, 450), size=(20, -1))
		# self.check_combine.Bind (wx.EVT_CHECKBOX, self.activate_enter_name)
		# # self.check_combine.Disable()
		# self.texto_combine = wx.StaticText (self.panel, wx.ID_ANY, "Combine PDFs", pos=(60, 450), size=(115, -1))
		# self.texto_combine.Disable()

		# self.pdf_final = wx.TextCtrl (self.panel, pos=(40, 340), size=(455, 30), value="Enter the name of pdf")
		# self.pdf_final.Disable()
		# self.pdf_final.SetFont (fuente)
		# self.pdf_final.SetForegroundColour(wx.Colour(200,200,200))
		# self.pdf_final.Bind (wx.EVT_SET_FOCUS, self.clear)

		self.comenzar = wx.Button (self.panel, label="Start", pos=(50, 500), size=(100, 30))
		self.comenzar.Bind (wx.EVT_BUTTON, self.repartidor)

		self.ca = wx.Button (self.panel, label="Exit", pos=(385, 500), size=(100, 30))
		self.ca.Bind (wx.EVT_BUTTON, self.salir)

		self.rutas = []
		self.carpeta = getcwd() + "\\Temp\\"
		self.flag = 0
		self.ancho = 0
		self.alto = 0

		self.Center(True)
		self.Show(True)

	def active_controls (self, event, origen):
		if origen == 'insert':
			self.TC_insert.Enable()
			self.boton_insert.Enable()
			self.TC_insert.SetBackgroundColour(wx.WHITE)
			self.TC_sustituir.Disable()
			self.boton_sustituir.Disable()

		else:
			self.TC_sustituir.Enable()
			self.boton_sustituir.Enable()
			self.TC_sustituir.SetBackgroundColour(wx.WHITE)
			self.TC_insert.Disable()
			self.boton_insert.Disable()

	def onOpenFile2 (self, event, origen):
		""" Create and show the Open FileDialog """
		wildcard = "Image files (*.png; *.jpg) | *.png; *.jpg |" "All files (*.*)|*.*"
		dlg = wx.FileDialog(self, message="Choose a file", defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			# self.rutas = dlg.GetPaths()
			# mi_rutas = dlg.GetFilenames()

			if origen == 'insert':
				self.TC_insert.SetValue (dlg.GetPaths()[0])

			else:
				self.TC_sustituir.SetValue (dlg.GetPaths()[0])

	def onOpenFile (self, event):
		""" Create and show the Open FileDialog """
		wildcard = "Image files (*.png; *.jpg) | *.png; *.jpg |" "All files (*.*)|*.*"
		dlg = wx.FileDialog(self, message="Choose a file", defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			self.rutas = dlg.GetPaths()
			mi_rutas = dlg.GetFilenames()

			if len(mi_rutas) > 1:
				self.check_combine.Enable()
				self.texto_combine.Enable()

			if len(mi_rutas) == 1:
				# print mi_rutas[0]
				nombre, ext = path.splitext(mi_rutas[0])

				if ext == '.pdf':
					self.disjoin_pdf(self, mi_rutas[0])

				else:
					for ele in reversed(mi_rutas):
						self.list_ctrl.InsertItem (0, ele)

			else:
				for ele in reversed(mi_rutas):
					self.list_ctrl.InsertItem (0, ele)

		dlg.Destroy()

	def convert_file (self, event, archivo):
		if not path.exists(self.carpeta):
			mkdir(self.carpeta)

		nombre = path.basename(archivo)
		im = Image.open(nombre)
		if im.mode == "RGBA":
		    im = im.convert("RGB")

		# fname, _ = path.splitext(nombre)
		newfilename = nombre + ".pdf"

		carpeta_destino = self.carpeta + newfilename

		# print im.size[0]
		# print im.size[1]
		ancho = int(self.ancho * 1.39)
		alto = int(self.alto * 1.39)
		# print ancho, alto
		# 827 1169

		img = im.resize((ancho, alto), Image.ANTIALIAS)
		enfoque = ImageEnhance.Sharpness(img)
		factor = 1 / 4.0
		img2 = enfoque.enhance(3.4)
		# img = im.resize((ancho, alto))
		if not path.exists(carpeta_destino):
			img2.save(carpeta_destino, "PDF", resolution=100.0, quality=95)
			# img.save(carpeta_destino, "PDF", quality=95)

	def convert_files (self, event):
		if not path.exists(self.carpeta):
			mkdir(self.carpeta)

		try:
			for png in self.rutas:
				nombre = path.basename(png)
				im = Image.open(nombre)

				if im.mode == "RGBA":
				    im = im.convert("RGB")

				fname, _ = path.splitext(nombre)
				newfilename = fname + ".pdf"

				carpeta_destino = self.carpeta + newfilename

				if not path.exists(carpeta_destino):
					im.save(carpeta_destino, "PDF", resolution=100.0)

		except Exception as e:
			print(e)

		# seleccionado = self.list_ctrl.GetFirstSelected()
		# if self.radio_sustituir.GetValue():
		# 	pass


		# if self.check_combine.IsChecked():
		# 	self.combine_pdfs (self)

	def combine_pdfs (self, event):
		pdfs = [path.join(self.carpeta, archivo) for archivo in listdir(self.carpeta) if archivo.endswith(".pdf")]
		name_file_output = self.pdf_final.GetValue() + '.pdf'
		fusionador = PdfFileMerger()

		for pdf in pdfs:
			fusionador.append(open(pdf, 'rb'))

		with open(self.carpeta + name_file_output, 'wb') as salida:
		    fusionador.write(salida)

	def obtener_num_paginas (self, event, archivo):
		print archivo
		flag = 0
		with open(archivo, "rb") as f:
		# handler = open("seguro.pdf", "rb")
			PDF = PdfFileReader (f)
			print PDF
			if PDF.isEncrypted:
				try:
					salida = PDF.decrypt('')

				except NotImplementedError:
					command = "C:\\Path\\to\\folder\\qpdf\\bin\\qpdf --password= --decrypt %s %s" % (archivo, self.carpeta + "temp.pdf")
					system(command)
					# with open("\\Temp\\misalida.pdf", "rb"):
					handler2 = open(self.carpeta + "temp.pdf", "rb")
					PDF2 = PdfFileReader (handler2)
					salida = PDF2
					flag = 1

			else:
				copy (archivo, self.carpeta + "temp.pdf")
				salida = PDF

		# ------------------------------------------------- PARA QUE APAREZCA EN EL LISTCTRL
			num_pag = salida.getNumPages()
			# print salida.getPage(0)
			self.ancho = int(salida.getPage(0).mediaBox.getWidth())
			self.alto = int(salida.getPage(0).mediaBox.getHeight())
			# print self.ancho
			# print self.alto
			if flag:
				handler2.close()
			return num_pag

	def disjoin_pdf (self, event, archivo):
		num_pag = self.obtener_num_paginas(self, archivo)

		for ele in reversed(range(1, num_pag+1)):
			nombre = str(ele) + '.pdf'
			self.list_ctrl.InsertItem (0, nombre)

	def repartidor (self, event):
		if self.radio_sustituir.GetValue():
			self.sustituir(self)

		elif self.radio_insert.GetValue():
			self.insertar(self)

		else:
			pass

	def sustituir (self, event):
		contador = self.list_ctrl.GetItemCount()

		with open(self.carpeta + "temp.pdf", "rb") as fr:
			PDF = PdfFileReader (fr)

			PDF_final = PdfFileWriter()

			seleccionado = self.list_ctrl.GetFirstSelected()
			if seleccionado == -1:
				wx.MessageBox('Operation could not be completed.\nYou must select an item', 'Warning', wx.OK | wx.ICON_WARNING)
				return
			# item = self.list_ctrl.GetItem (itemIdx=seleccionado, col=0)

			for idx in range(contador):
				if idx == seleccionado:
					new_ele = self.TC_sustituir.GetValue()
					fname, ext = path.splitext(new_ele)
					if ext == '.png' or ext == '.jpg' or ext == '.bmp':
						self.convert_file (self, new_ele)

						nombre = path.basename(new_ele)
						handler = open(self.carpeta + "%s.pdf" % nombre, "rb")
						PDF_temp = PdfFileReader (handler)
						pag = PDF_temp.getPage(0)
						PDF_final.addPage(pag)

				else:
					pag = PDF.getPage(idx)
					PDF_final.addPage(pag)

			with open(self.carpeta + "output.pdf", "wb") as fw:
				PDF_final.write(fw)

	def insertar (self, event):
		contador = self.list_ctrl.GetItemCount()

		with open(self.carpeta + "temp.pdf", "rb") as fr:
			PDF = PdfFileReader (fr)

			PDF_final = PdfFileWriter()

			seleccionado = self.list_ctrl.GetFirstSelected()
			if seleccionado == -1:
				wx.MessageBox('Operation could not be completed.\nYou must select an item', 'Warning', wx.OK | wx.ICON_WARNING)
				return

			ind = 0
			# item = self.list_ctrl.GetItem (itemIdx=seleccionado, col=0)

			for idx in range(contador + 1):
				if idx == seleccionado:
					new_ele = self.TC_insert.GetValue()
					fname, ext = path.splitext(new_ele)
					if ext == '.png' or ext == '.jpg' or ext == '.bmp':
						self.convert_file (self, new_ele)

						nombre = path.basename(new_ele)
						handler = open(self.carpeta + "%s.pdf" % nombre, "rb")
						PDF_temp = PdfFileReader (handler)
						pag = PDF_temp.getPage(0)
						PDF_final.addPage(pag)

				else:
					pag = PDF.getPage(ind)
					PDF_final.addPage(pag)
					ind += 1

			with open(self.carpeta + "output.pdf", "wb") as fw:
				PDF_final.write(fw)

	def unir_archivos (self, event):
		# -------------------------------------------------
		lista = []
		contador = self.list_ctrl.GetItemCount()

		# for fila in range(contador):
		# 	item = self.list_ctrl.GetItem(itemIdx=fila, col=0)
		# 	lista.append(item.GetText())
		# 	# print item.GetText()
		# print lista

		with open(self.carpeta + "temp.pdf", "rb") as fr:
			PDF = PdfFileReader (fr)

			PDF_final = PdfFileWriter()
			ind = 0

			for idx, ele in enumerate(lista):

				fname, ext = path.splitext(ele)
				if ext == '.png' or ext == '.jpg' or ext == '.bmp':
					print "ES UNA IMAGEN"
					self.convert_file (self, ele)

					handler = open(self.carpeta + "%s.pdf" % ele, "rb")
					PDF_temp = PdfFileReader (handler)
					pag = PDF_temp.getPage(0)
					PDF_final.addPage(pag)

				else:
					pag = PDF.getPage(ind)
					PDF_final.addPage(pag)
					ind += 1

			with open(self.carpeta + "output.pdf", "wb") as fw:
				PDF_final.write(fw)

	def activate_enter_name (self, event):
		if self.flag:
			self.pdf_final.Disable()
			self.flag = 0

		else:
			self.pdf_final.Enable()
			self.flag = 1

	def clear (self, event):
		self.pdf_final.Clear()
		self.pdf_final.SetForegroundColour(wx.BLACK)
		self.pdf_final.Unbind(wx.EVT_SET_FOCUS)
		event.Skip()

	def salir (self, event):
		self.Close(True)


if __name__ == "__main__":
	app = wx.App(False)
	frame = Ventana (None, 'PDFs Converter and Merger')
	app.MainLoop()

