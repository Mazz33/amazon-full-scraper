import numpy as np
from threading import Thread
from PyQt6 import QtGui, QtWidgets, QtCore
from Item import Item

class ScraperSignals(QtCore.QObject):
	finished = QtCore.pyqtSignal()
	progress = QtCore.pyqtSignal(int)

class Scraper(QtCore.QRunnable):
	def __init__(this, asinList, index, writer):
		super(Scraper, this).__init__()

		this.asinList = asinList
		this.index = index
		this.writer = writer

		this.signals = ScraperSignals()

	def run(this):
		row = this.index
		for asin in this.asinList:
			row += 1
			item = Item(asin)
			this.writer.writeItemData(item, row)
			this.signals.progress.emit(1)
		this.signals.finished.emit()
