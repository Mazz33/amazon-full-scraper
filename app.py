import sys
import os 
import numpy as np
from PyQt6 import QtCore, QtGui, QtWidgets
from pandas.errors import ParserError
from scraper import Scraper
from Excel import *

class App():
	
	def __init__(this):
		this.threadPool = QtCore.QThreadPool()
		this.createWidgets()
		this.setObjectNames()
		this.setFonts()
		this.addWidgets()
		this.setupUi()
		this.alignAsinViewSection()
		this.setUiText(this.root)
		this.checkOptions()
		this.buddyLabels()
		this.setMaxThreadCount()
		this.setupThreadSlider()
		this.connectEvents()
		this.updateCheckboxValues()
		this.root.closeEvent = this.exitApplication

	
	def showImportedAsinsList(this):
		itemModel = QtGui.QStandardItemModel(this.importedAsinList)
		for asin in this.asinList:
			item = QtGui.QStandardItem(asin)
			itemModel.appendRow(item)
		this.importedAsinList.setModel(itemModel)
		this.importedAsinList.show()
	
	def show(this):
		this.root.show()
	
	def getDir(this, dirname):
		return os.path.dirname(dirname)
	
	def getCurrentDirectory(this):
		return this.getDir(os.path.realpath(__file__))
	
	def createWidgets(this):
		this.root = QtWidgets.QMainWindow() # Parent Window
		this.mainAppGrid = QtWidgets.QWidget(parent=this.root) # Grid that has all other widgets in to display them
		this.gridLayout = QtWidgets.QGridLayout(this.mainAppGrid) # Layout for the mainAppGrid
		this.programTitle = QtWidgets.QLabel(parent=this.mainAppGrid) # The heading at the top of the application
		this.numberOfAsinsImported = QtWidgets.QSpinBox(parent=this.mainAppGrid) # Shows the number of asins that has been imported
		
		# List view for the imported ASINs list
		this.importedAsinList = QtWidgets.QListView(parent=this.mainAppGrid) # After asins are imported they will be displayed here
		
		# Vertical Boxes
		this.asinsView = QtWidgets.QVBoxLayout()
		this.fileButtonsVBox = QtWidgets.QVBoxLayout()
		this.settings = QtWidgets.QVBoxLayout()
		this.scrapingOptions = QtWidgets.QVBoxLayout()
		
		# Horizontal Boxes
		this.asinsViewHBox = QtWidgets.QHBoxLayout()
		this.bottomHBox = QtWidgets.QHBoxLayout()
		this.excelScrapSettings = QtWidgets.QHBoxLayout()
		this.columnSelect = QtWidgets.QHBoxLayout()
		this.threadOptions = QtWidgets.QHBoxLayout()
		this.progressBarHBox = QtWidgets.QHBoxLayout()
		
		# Text Line
		this.outFileSavePath = QtWidgets.QLineEdit(parent=this.mainAppGrid)
		this.columnLetter = QtWidgets.QLineEdit(parent=this.mainAppGrid)
		
		# Slider to pick number of threads
		this.threadsSlider = QtWidgets.QSlider(parent=this.mainAppGrid)
		
		# Display number of threads picked from the slider
		this.numberOfThreads = QtWidgets.QSpinBox(parent=this.mainAppGrid)
		
		# Bars
		this.menubar = QtWidgets.QMenuBar(parent=this.root)
		this.statusbar = QtWidgets.QStatusBar(parent=this.root)
		this.progressBar = QtWidgets.QProgressBar(parent=this.mainAppGrid)
		
		# Checkboxes
		this.titleCheck = QtWidgets.QCheckBox(parent=this.mainAppGrid)
		this.priceCheck = QtWidgets.QCheckBox(parent=this.mainAppGrid)
		this.currencyCheck = QtWidgets.QCheckBox(parent=this.mainAppGrid)
		this.sellerCheck = QtWidgets.QCheckBox(parent=this.mainAppGrid)
		this.imageCheck = QtWidgets.QCheckBox(parent=this.mainAppGrid)
		
		# Labels
		this.numberOfAsinsLabel = QtWidgets.QLabel(parent=this.mainAppGrid)
		this.threadsLabel = QtWidgets.QLabel(parent=this.mainAppGrid)
		this.columnLetterLabel = QtWidgets.QLabel(parent=this.mainAppGrid)
		
		# Buttons
		this.importButton = QtWidgets.QPushButton(parent=this.mainAppGrid)
		this.saveAsButton = QtWidgets.QPushButton(parent=this.mainAppGrid)
		this.startButton = QtWidgets.QPushButton(parent=this.mainAppGrid)
	
	def setObjectNames(this):
		this.root.setObjectName("MainWindow")
		this.mainAppGrid.setObjectName("mainAppGrid")
		this.gridLayout.setObjectName("gridLayout")
		this.programTitle.setObjectName("programTitle")
		this.importedAsinList.setObjectName("importedAsinList")
		this.numberOfAsinsImported.setObjectName("numberOfAsinsImported")
		this.outFileSavePath.setObjectName("outFileSavePath")
		this.columnLetter.setObjectName("columnLetter")
		this.threadsSlider.setObjectName("threadsSlider")
		this.numberOfThreads.setObjectName("numberOfThreads")
		
		# Vertical Boxes
		this.asinsView.setObjectName("asinsView")
		this.fileButtonsVBox.setObjectName("fileButtonsVBox")
		this.settings.setObjectName("settings")
		this.scrapingOptions.setObjectName("scrapingOptions")
		
		# Horizontal Boxes
		this.asinsViewHBox.setObjectName("asinsViewHBox")
		this.bottomHBox.setObjectName("bottomHBox")
		this.excelScrapSettings.setObjectName("excelScrapSettings")
		this.columnSelect.setObjectName("columnSelect")
		this.threadOptions.setObjectName("threadOptions")
		this.progressBarHBox.setObjectName("progressBarHBox")
		
		# Bars
		this.menubar.setObjectName("menubar")
		this.statusbar.setObjectName("statusbar")
		this.progressBar.setObjectName("progressBar")
		
		# Checkboxes
		this.titleCheck.setObjectName("titleCheck")
		this.priceCheck.setObjectName("priceCheck")
		this.currencyCheck.setObjectName("currencyCheck")
		this.sellerCheck.setObjectName("sellerCheck")
		this.imageCheck.setObjectName("imageCheck")
		
		# Labels
		this.numberOfAsinsLabel.setObjectName("numberOfAsinsLabel")
		this.columnLetterLabel.setObjectName("columnLetterLabel")
		this.threadsLabel.setObjectName("threadsLabel")
		
		# Buttons
		this.importButton.setObjectName("importButton")
		this.saveAsButton.setObjectName("saveAsButton")
		this.startButton.setObjectName("startButton")
	
	def addWidgets(this):
		this.gridLayout.addWidget(this.programTitle, 0, 0, 1, 2)
		this.asinsView.addWidget(this.importedAsinList)
		this.asinsViewHBox.addWidget(this.numberOfAsinsLabel)
		this.asinsViewHBox.addWidget(this.numberOfAsinsImported)
		this.bottomHBox.addWidget(this.outFileSavePath)
		this.fileButtonsVBox.addWidget(this.importButton)
		this.fileButtonsVBox.addWidget(this.saveAsButton)
		this.scrapingOptions.addWidget(this.titleCheck)
		this.scrapingOptions.addWidget(this.priceCheck)
		this.scrapingOptions.addWidget(this.currencyCheck)
		this.scrapingOptions.addWidget(this.sellerCheck)
		this.scrapingOptions.addWidget(this.imageCheck)
		this.columnSelect.addWidget(this.columnLetterLabel)
		this.columnSelect.addWidget(this.columnLetter)
		this.threadOptions.addWidget(this.threadsLabel)
		this.threadOptions.addWidget(this.threadsSlider)
		this.threadOptions.addWidget(this.numberOfThreads)
		this.progressBarHBox.addWidget(this.startButton)
		this.progressBarHBox.addWidget(this.progressBar)
	
	def alignAsinViewSection(this):
		this.numberOfAsinsLabel.setMidLineWidth(0)
		this.numberOfAsinsLabel.setAlignment(QtCore.Qt.AlignmentFlag.AlignRight|QtCore.Qt.AlignmentFlag.AlignTrailing|QtCore.Qt.AlignmentFlag.AlignVCenter)
		this.numberOfAsinsImported.setEnabled(True)
		this.numberOfAsinsImported.setInputMethodHints(QtCore.Qt.InputMethodHint.ImhNone)
		this.numberOfAsinsImported.setFrame(False)
		this.numberOfAsinsImported.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeading|QtCore.Qt.AlignmentFlag.AlignLeft|QtCore.Qt.AlignmentFlag.AlignVCenter)
		this.numberOfAsinsImported.setReadOnly(True)
		this.numberOfAsinsImported.setButtonSymbols(QtWidgets.QAbstractSpinBox.ButtonSymbols.NoButtons)
		this.numberOfAsinsImported.setMaximum(50000)
	
	def setupUi(this):
		this.root.resize(580, 420)
		this.programTitle.setTextFormat(QtCore.Qt.TextFormat.AutoText)
		this.programTitle.setScaledContents(False)
		this.programTitle.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
		this.programTitle.setWordWrap(False)
		this.asinsView.addLayout(this.asinsViewHBox)
		this.gridLayout.addLayout(this.asinsView, 1, 0, 1, 1)
		this.saveAsButton.setAutoDefault(False)
		this.saveAsButton.setDefault(False)
		this.saveAsButton.setFlat(False)
		this.startButton.setAutoDefault(False)
		this.startButton.setDefault(False)
		this.startButton.setFlat(False)
		this.bottomHBox.addLayout(this.fileButtonsVBox)
		this.gridLayout.addLayout(this.bottomHBox, 3, 0, 1, 2)
		this.excelScrapSettings.addLayout(this.scrapingOptions)
		this.columnSelect.setSizeConstraint(QtWidgets.QLayout.SizeConstraint.SetDefaultConstraint)
		this.columnLetter.setMaxLength(1)
		this.excelScrapSettings.addLayout(this.columnSelect)
		this.settings.addLayout(this.excelScrapSettings)
		this.settings.addLayout(this.threadOptions)
		this.gridLayout.addLayout(this.settings, 1, 1, 1, 1)
		this.gridLayout.addLayout(this.progressBarHBox, 2, 0, 1, 2)
		this.root.setCentralWidget(this.mainAppGrid)
		this.menubar.setGeometry(QtCore.QRect(0, 0, 580, 21))
		this.root.setMenuBar(this.menubar)
		this.root.setStatusBar(this.statusbar)
		this.outFileSavePath.setClearButtonEnabled(False)
		this.outFileSavePath.setText("{}/data.xlsx".format(this.getCurrentDirectory()))
		
	
	def setFonts(this):
		font = QtGui.QFont()
		font.setFamily("MS Sans Serif")
		font.setPointSize(18)
		font.setBold(True)
		font.setWeight(75)
		this.programTitle.setFont(font)
		font = QtGui.QFont()
		font.setPointSize(12)
		this.importButton.setFont(font)
		this.saveAsButton.setFont(font)
		this.startButton.setFont(font)
	
	def buddyLabels(this):
		this.numberOfAsinsLabel.setBuddy(this.numberOfAsinsImported)
		this.columnLetterLabel.setBuddy(this.columnLetter)
		this.threadsLabel.setBuddy(this.threadsSlider)
	
	def checkOptions(this):
		this.titleCheck.setChecked(True)
		this.priceCheck.setChecked(True)
		this.currencyCheck.setChecked(True)
		this.sellerCheck.setChecked(True)
		this.imageCheck.setChecked(True)
	
	def setMaxThreadCount(this):
		this.numberOfThreads.setMinimum(1)
		maxThreads = this.threadPool.globalInstance().maxThreadCount()
		this.maxAvailableThreadCount = 12 if maxThreads > 12 else maxThreads
		this.numberOfThreads.setMaximum(this.maxAvailableThreadCount)
		this.numberOfThreads.setProperty("value", (this.maxAvailableThreadCount // 2) if (this.maxAvailableThreadCount > 1) else 1)
	
	
	def setupThreadSlider(this):
		this.threadsSlider.setToolTipDuration(-1)
		this.threadsSlider.setMinimum(1)
		this.threadsSlider.setMaximum(this.maxAvailableThreadCount)
		this.threadsSlider.setPageStep(1)
		this.threadsSlider.setProperty("value", this.numberOfThreads.value())
		this.threadsSlider.setOrientation(QtCore.Qt.Orientation.Horizontal)
		this.threadsSlider.setInvertedAppearance(False)
		this.threadsSlider.setInvertedControls(False)
		this.threadsSlider.setTickPosition(QtWidgets.QSlider.TickPosition.NoTicks)
	
	def showError(this):
		msgBox = QtWidgets.QMessageBox()
		msgBox.setWindowTitle("Error!!")
		msgBox.setText("Hello")
		msgBox.exec()
	
	def exitApplication(this, event):
		try:
			this.writer.close()
		except AttributeError:
			pass
	
	def connectEvents(this):
		this.threadsSlider.sliderMoved['int'].connect(this.numberOfThreads.setValue)
		this.numberOfThreads.valueChanged['int'].connect(this.threadsSlider.setValue)
		QtCore.QMetaObject.connectSlotsByName(this.root)
		this.importButton.clicked.connect(this.importFile)
		this.saveAsButton.clicked.connect(this.saveFile)
		this.startButton.clicked.connect(this.startSearch)
		this.columnLetter.textChanged.connect(this.readColumnData)
		this.titleCheck.stateChanged.connect(this.updateCheckboxValues)
		this.priceCheck.stateChanged.connect(this.updateCheckboxValues)
		this.currencyCheck.stateChanged.connect(this.updateCheckboxValues)
		this.sellerCheck.stateChanged.connect(this.updateCheckboxValues)
		this.imageCheck.stateChanged.connect(this.updateCheckboxValues)
	
	def setupProgressBar(this):
		this.progressBar.setProperty("value", 0)
	
	def setUiText(this, MainWindow):
		translate = QtCore.QCoreApplication.translate
		MainWindow.setWindowTitle("Amazon Scraper")
		this.programTitle.setText("Amazon Scraper")
		this.numberOfAsinsLabel.setText(translate("MainWindow", "ASINs Loaded:"))
		this.outFileSavePath.setPlaceholderText(translate("MainWindow", "Data will be saved at this location"))
		this.titleCheck.setText(translate("MainWindow", "Title"))
		this.priceCheck.setText(translate("MainWindow", "Price"))
		this.currencyCheck.setText(translate("MainWindow", "Currency        "))
		this.sellerCheck.setText(translate("MainWindow", "Seller"))
		this.imageCheck.setText(translate("MainWindow", "Image"))
		this.columnLetterLabel.setText(translate("MainWindow", "Column:"))
		this.columnLetter.setText(translate("MainWindow", "B"))
		this.threadsLabel.setText("Threads:")
		this.threadsSlider.setToolTip("Threads increase speed of the search but the more you use the higher the chance you get flagged and banned by Amazon")
		
		# Buttons
		this.importButton.setText(translate("MainWindow", "Import"))
		this.saveAsButton.setText(translate("MainWindow", "Save As"))
		this.startButton.setText(translate("MainWindow", "Start"))
	
	def readColumnData(this):
		try:
			this.reader = ExcelReader(this.inputFileName)
			this.asinList = this.reader.readColumnData(this.columnLetter.text())
			this.asinListLength = len(this.asinList)
			this.updateProgressBarMaximumValue() 
			this.updateNumberOfAsinsImported()
			this.showImportedAsinsList()
		except AttributeError:
			this.unableToReadColumn()
		except TypeError:
			this.emptyColumnError()
		except ParserError:
			this.invalidColumnError()

	
	def importFile(this):
		this.openFileNameDialog()
		try:
			this.writer = ExcelWriter(this.outputFileName, this.checkbox)
		except AttributeError:
			this.writer = ExcelWriter("data.xlsx", this.checkbox)
		this.readColumnData()
	
	def saveFile(this):
		this.saveFileDialog()
		try:
			this.writer.close()
			this.writer = ExcelWriter(this.outputFileName)
		except (AttributeError, NameError):
			this.writer = ExcelWriter(this.outputFileName)
		except Exception as E:
			this.errorOpeningFileForSave()
		
	
	def getStartPoints(this, arr):
		this.startPoints = []
		for i in range(len(arr)):
			if (i==0):
				this.startPoints.append(1)
			else:
				if (len(arr[i]) == len(arr[i-1])):
					this.startPoints.append(this.startPoints[i-1] + len(arr[i]))
				else:
					this.startPoints.append(this.startPoints[i-1] + len(arr[i]) + 1)
		return this.startPoints
	
	def startSearch(this):
		try:
			assert this.columnLetter.text() != ""
			threadCount = this.numberOfThreads.value()
			splitted = np.array_split(this.asinList, threadCount)
			startPoints = this.getStartPoints(splitted)
			workers = []
			for i in range(threadCount):
				worker = Scraper(splitted[i], startPoints[i], this.writer)
				worker.signals.progress.connect(this.updateFinishedItemsValue)
				worker.signals.finished.connect(this.checkProgressFinished)
				workers.append(worker)
			this.disableOptions()
			for w in workers:
				this.threadPool.start(w)
			this.updateFinishedItemsValue() #? For some reason, the last thread does not call the update function. and the last update doesn't occur
		except AssertionError:
			this.emptyColumnLetterError()
		except AttributeError as E:
			if (not hasattr(this, "reader")):
				this.importError()
				return
	
	def checkProgressFinished(this):
		if (this.progressBar.value() == this.progressBar.maximum()):
			this.finishedSearching()
	
	def disableOptions(this):
		this.threadsSlider.setEnabled(False)
		this.numberOfThreads.setEnabled(False)
		this.titleCheck.setEnabled(False)
		this.priceCheck.setEnabled(False)
		this.currencyCheck.setEnabled(False)
		this.sellerCheck.setEnabled(False)
		this.imageCheck.setEnabled(False)
		this.startButton.setEnabled(False)
		this.columnLetter.setEnabled(False)

	
	def openFileNameDialog(this):
		fileName, _ = QtWidgets.QFileDialog.getOpenFileName(this.root, "Open Excel File", "","Excel Files (*.xlsx)",)
		if fileName:
			this.inputFileName = fileName
	
	def saveFileDialog(this):
		fileName, _ = QtWidgets.QFileDialog.getSaveFileName(this.root, "Save as","","Excel Files (*.xlsx)",)
		if fileName:
			this.outFileSavePath.setText(os.path.realpath(fileName))
			this.outputFileName = fileName
	
	def updateFinishedItemsValue(this):
		this.progressBar.setProperty("value", this.progressBar.value() + 1)
	
	def updateProgressBarMaximumValue(this):
		this.progressBar.setMaximum(this.asinListLength)
		
	
	def updateNumberOfAsinsImported(this):
		this.numberOfAsinsImported.setValue(this.asinListLength)
	
	def getInputFileName(this):
		if (this.inputFileName):
			return this.inputFileName
		return None
	
	def getOutputFileName(this):
		if (this.outputFileName):
			return this.outputFileName
		return None
	
	def updateCheckboxValues(this):
		this.checkbox = {
			"title": this.titleCheck.isChecked(),
			"price": this.priceCheck.isChecked(),
			"currency": this.currencyCheck.isChecked(),
			"seller": this.sellerCheck.isChecked(),
			"image": this.imageCheck.isChecked()
		}
		try:
			this.writer.updateCheckbox(this.checkbox)
		except:
			pass

	
	# Pop up Dialogs
	def displayDialogBox(this, title, msg):
		box = QtWidgets.QMessageBox()
		box.setWindowTitle(title)
		box.setText(msg)
		return box.exec()
	
	def finishedSearching(this):
		this.displayDialogBox("Done", "Searching done")
		this.root.close()
	
	def importError(this):
		this.displayDialogBox("Error", "You have to import first")
	
	def emptyColumnLetterError(this):
		this.displayDialogBox("Error", "You must enter a column letter")
	
	def errorOpeningFileForSave(this):
		this.displayDialogBox("Error", "Couldn't open the file for saving")
	
	def errorOpeningFileForImport(this):
		this.displayDialogBox("Error", "Couldn't import file")
	
	def unableToReadColumn(this):
		this.displayDialogBox("Error", "Can't read columns")
	
	def emptyColumnError(this):
		this.displayDialogBox("Error", "Column is empty!")
	
	def invalidColumnError(this):
		this.displayDialogBox("Error", "Column is empty")

if __name__ == "__main__":
	app = QtWidgets.QApplication(sys.argv)
	gui = App()
	gui.show()
	sys.exit(app.exec())
