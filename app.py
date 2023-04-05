import sys
import os 
from PyQt6 import QtCore, QtGui, QtWidgets

from scraper import startSearch
from Excel import *

class App(object):
	def __init__(this):
		this.writer = ExcelWriter()

		this.createWidgets()
		this.setObjectNames()
		this.setFonts()
		this.addWidgets()
		this.setupUi()
		this.setUiText(this.root)
		this.checkOptions()
		this.buddyLabels()
		this.setupThreadSlider()
		this.connectEvents()
		this.root.closeEvent = this.exitApplication


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
		this.imageCheckLink = QtWidgets.QCheckBox(parent=this.mainAppGrid)

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
		this.imageCheckLink.setObjectName("imageCheckLink")

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
		this.scrapingOptions.addWidget(this.imageCheckLink)
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
		this.numberOfThreads.setMinimum(1)
		this.numberOfThreads.setMaximum(12)
		this.numberOfThreads.setProperty("value", 6)
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
		this.imageCheckLink.setChecked(True)

	def setupThreadSlider(this):
		this.threadsSlider.setToolTipDuration(-1)
		this.threadsSlider.setMinimum(1)
		this.threadsSlider.setMaximum(12)
		this.threadsSlider.setPageStep(1)
		this.threadsSlider.setProperty("value", 6)
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
		this.writer.close()

	def connectEvents(this):
		this.threadsSlider.sliderMoved['int'].connect(this.numberOfThreads.setValue)
		this.numberOfThreads.valueChanged['int'].connect(this.threadsSlider.setValue)
		QtCore.QMetaObject.connectSlotsByName(this.root)
		this.importButton.clicked.connect(this.importFile)
		this.saveAsButton.clicked.connect(this.saveFile)
		this.startButton.clicked.connect(this.startSearch)
	
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
		this.currencyCheck.setText(translate("MainWindow", "Currency"))
		this.sellerCheck.setText(translate("MainWindow", "Seller"))
		this.imageCheck.setText(translate("MainWindow", "Image"))
		this.imageCheckLink.setText(translate("MainWindow", "Image Link"))
		this.columnLetterLabel.setText(translate("MainWindow", "Column Letter:"))
		this.columnLetter.setText(translate("MainWindow", "B"))
		this.threadsLabel.setText("Threads:")
		this.threadsSlider.setToolTip("Threads increase speed of the search but the more you use the higher the chance you get flagged and banned by Amazon")
		# Buttons
		this.importButton.setText(translate("MainWindow", "Import"))
		this.saveAsButton.setText(translate("MainWindow", "Save As"))
		this.startButton.setText(translate("MainWindow", "Start"))

	def importFile(this):
		this.openFileNameDialog()
		try:
			this.reader = ExcelReader(this.inputFileName)
		except:
			pass #! Something wrong happened with opening the file. TODO: Display error box

	def saveFile(this):
		this.saveFileDialog()
		try:
			this.writer.close()
			this.writer = ExcelWriter(this.outputFileName)
		except:
			pass #! Can't open a file for writing for some reason. TODO: Display error box
		
	def openFileNameDialog(this):
		fileName, _ = QtWidgets.QFileDialog.getOpenFileName(this.root,"Open Excel File", "","Excel Files (*.xlsx)",)
		if fileName:
			this.inputFileName = fileName

	def saveFileDialog(this):
		fileName, _ = QtWidgets.QFileDialog.getSaveFileName(this.root,"Save as","","Excel Files (*.xlsx)",)
		if fileName:
			this.outFileSavePath.setText(os.path.realpath(fileName))
			this.outputFileName = fileName

	def updateFinishedItemsValue(this):
		this.progressBar.setProperty("value",this.progressBar.value() + 1)

	def updateProgressBarMaximumValue(this):
		this.progressBar.setMaximum(this.asinListLength)
		
	def updateNumberOfAsinsImported(this):
		this.numberOfAsinsImported.setValue(len(this.asinList))

	def startSearch(this):
		try:
			assert this.columnLetter.text() != ""
			this.asinList = this.reader.readColumnData(this.columnLetter.text())
			this.asinListLength = len(this.asinList)
			this.updateProgressBarMaximumValue()
			this.updateNumberOfAsinsImported()
			startSearch(this.asinList, this.numberOfThreads.value(), this.writer, this.updateFinishedItemsValue)
			print(this.progressBar.value())
			this.updateFinishedItemsValue()
		except AssertionError:
			this.emptyColumnLetterError()
		except AttributeError as E:
			if (not hasattr(this, "reader")):
				this.importError()
				return
			raise(E)

	def getInputFileName(this):
		if (this.inputFileName):
			return this.inputFileName
		return None

	def getOutputFileName(this):
		if (this.outputFileName):
			return this.outputFileName
		return None

	# Error Dialogs
	def importError(this):
		box = QtWidgets.QMessageBox()
		box.setWindowTitle("Error!!")
		box.setText("You have to import first!!")
		box.exec()

	def emptyColumnLetterError(this):
		box = QtWidgets.QMessageBox()
		box.setWindowTitle("Error!!")
		box.setText("You must enter a column letter!!")
		box.exec()


if __name__ == "__main__":
	app = QtWidgets.QApplication(sys.argv)
	gui = App()
	gui.show()
	sys.exit(app.exec())
