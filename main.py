
import wx
from os import walk, path, getcwd, listdir, mkdir
from PIL import Image
from PyPDF2 import PdfFileMerger


class Ventana (wx.Frame):
	def __init__ (self, parent, title):
		no_resize = wx.DEFAULT_FRAME_STYLE ^(wx.MAXIMIZE_BOX | wx.CLOSE_BOX | wx.RESIZE_BORDER) #& ~ (wx.RESIZE_BORDER)
		wx.Frame.__init__ (self, parent, title=title, size=(554, 500), style=no_resize)
		self.panel = wx.Panel (self, style=wx.RAISED_BORDER)

		# font = wx.SystemSettings_GetFont(wx.SYS_SYSTEM_FONT)
		fuente = wx.Font (14, wx.DEFAULT, wx.NORMAL, wx.NORMAL)

		self.list_ctrl = wx.ListCtrl (self.panel, pos=(10, 80), size=(520, 200), style=wx.LC_REPORT | wx.BORDER_SUNKEN)
		self.list_ctrl.InsertColumn(0, 'File')

		self.texto = wx.StaticText (self.panel, wx.ID_ANY, "Choose the files to combine", pos=(140, 45), size=(160, -1))
		self.texto.SetFont (wx.Font (12, wx.DEFAULT, wx.NORMAL, wx.NORMAL))

		self.bt_ruta_origen = wx.Button (self.panel, label=". . .", pos=(340, 40), size=(40, 30))
		self.bt_ruta_origen.Bind (wx.EVT_BUTTON, lambda event: self.onOpenFile (event))

		self.box = wx.StaticBox(self.panel, wx.ID_ANY, "", pos=(20, 290), size=(495, 105))

		self.check_combine = wx.CheckBox (self.panel, wx.ID_ANY, pos=(40, 320), size=(20, -1))
		self.check_combine.Bind (wx.EVT_CHECKBOX, self.activate_enter_name)
		self.check_combine.Disable()
		self.texto_combine = wx.StaticText (self.panel, wx.ID_ANY, "Combine PDFs", pos=(60, 320), size=(115, -1))
		self.texto_combine.Disable()

		self.pdf_final = wx.TextCtrl (self.panel, pos=(40, 340), size=(455, 30), value="Enter the name of pdf")
		self.pdf_final.Disable()
		self.pdf_final.SetFont (fuente)
		self.pdf_final.SetForegroundColour(wx.Colour(200,200,200))
		self.pdf_final.Bind (wx.EVT_SET_FOCUS, self.clear)

		self.comenzar = wx.Button (self.panel, label="Start", pos=(50, 420), size=(100,30))
		self.comenzar.Bind (wx.EVT_BUTTON, self.convert_files)

		self.ca = wx.Button (self.panel, label="Exit", pos=(385, 420), size=(100, 30))
		self.ca.Bind (wx.EVT_BUTTON, self.salir)

		self.rutas = []
		self.carpeta = getcwd() + "\\Temp\\"
		self.flag = 0

		self.Center(True)
		self.Show(True)


	def onOpenFile(self, event):
		""" Create and show the Open FileDialog """
		wildcard = "Image files (*.png; *.jpg) | *.png; *.jpg |" "All files (*.*)|*.*"
		dlg = wx.FileDialog(self, message="Choose a file", defaultFile="", wildcard=wildcard, style=wx.FD_OPEN | wx.FD_MULTIPLE | wx.FD_CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			self.rutas = dlg.GetPaths()
			paths = dlg.GetFilenames()

			if len(paths) > 1:
				self.check_combine.Enable()
				self.texto_combine.Enable()

			for path in reversed(paths):
				self.list_ctrl.InsertItem (0, path)

		dlg.Destroy()

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

		if self.check_combine.IsChecked():
			self.combine_pdfs (self)

	def combine_pdfs (self, event):
		pdfs = [path.join(self.carpeta, archivo) for archivo in listdir(self.carpeta) if archivo.endswith(".pdf")]
		name_file_output = self.pdf_final.GetValue() + '.pdf'
		fusionador = PdfFileMerger()

		for pdf in pdfs:
			fusionador.append(open(pdf, 'rb'))

		with open(self.carpeta + name_file_output, 'wb') as salida:
		    fusionador.write(salida)

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

