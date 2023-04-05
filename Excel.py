import xlsxwriter
from math import isnan
from pandas import read_excel

class ExcelReader:
	def __init__(this, path):
		this.path = path

	def readFullDoc(this):
		sheetData = read_excel(this.path)
		return sheetData.values

	def readColumnData(this, columnLetter):
		columnData = read_excel(this.path, usecols="{}".format(columnLetter))
		formattedArray = []
		for i in columnData.values:
			if (type(i[0]) is not str and isnan(i[0])):
				continue
			formattedArray.append(i[0])
		return formattedArray
			
	

class ExcelWriter:
	def __init__(this, outputFile="data.xlsx", checkBox=False):
		if (checkBox):
			this.checkBox = checkBox
		this.workbook = xlsxwriter.Workbook(outputFile)
		this.worksheet = this.workbook.add_worksheet()
		this.setColumnsDimensions()
		this.writeColumnsTitles()

	def setColumnsDimensions(this):
		this.worksheet.set_column("A:A", 12)
		this.worksheet.set_column("B:B", 50)
		this.worksheet.set_column("C:C", 10)
		this.worksheet.set_column("D:D", 10)
		this.worksheet.set_column("E:E", 10)
		this.worksheet.set_column("F:F", 10)
		this.worksheet.set_column("G:G", 15)
		this.worksheet.set_column("H:H", 10)
		this.worksheet.set_default_row(50)
 
	def writeColumnsTitles(this):
		this.worksheet.write("A1", "ASIN")
		this.worksheet.write("B1", "Title")
		this.worksheet.write("C1", "Currency")
		this.worksheet.write("D1", "Price")
		this.worksheet.write("E1", "Full Price")
		this.worksheet.write("F1", "Discount")
		this.worksheet.write("G1", "Seller")
		this.worksheet.write("H1", "Picture")

	def writeAsin(this, asin, row):
		this.worksheet.write_string("A{}".format(row), asin)
	
	def writeTitle(this, title, row):
		if (this.checkBox and this.checkBox["title"]):
			this.worksheet.write_string("B{}".format(row), title)
		else:
			return
	
	def writeCurrency(this, currency, row):
		if (this.checkBox and this.checkBox["currency"]):
			this.worksheet.write_string("C{}".format(row), currency)
		else:
			return
		
	def writePrice(this, price, row):
		if (this.checkBox and this.checkBox["price"]):
			this.worksheet.write_string("D{}".format(row), price)
		else:
			return

	def writeFullPrice(this, fullPrice, row):
		if (this.checkBox and this.checkBox["price"]):
			this.worksheet.write_string("E{}".format(row), fullPrice)
		else:
			return

	def writeDiscount(this, discount, row):
		if (this.checkBox and this.checkBox["price"]):
			this.worksheet.write_string("F{}".format(row), discount)
		else:
			return

	def writeSeller(this, seller, row):
		if (this.checkBox and this.checkBox["seller"]):
			this.worksheet.write_string("G{}".format(row), seller)
		else:
			return

	def writeImage(this, item, row, write=True):
		if (not write):
			return
		if (this.checkBox and this.checkBox["image"]):
			if (item.imageData is None):
				this.worksheet.write_string("H{}".format(row), "N/A")
				return
			this.worksheet.insert_image("H{}".format(row), item.asin, {
				"object_position": 1,
				"url": item.url,
				"x_scale": item.x_scale,
				"y_scale": item.y_scale,
				"image_data": item.imageData
			})
		else:
			return

	def writeItemData(this, item, row):
		if (item.none):
			this.writeAsin(item.asin, row)
			this.writeTitle("N/A", row)
			this.writeCurrency("N/A", row)
			this.writePrice("N/A", row)
			this.writeFullPrice("N/A", row)
			this.writeDiscount("N/A", row)
			this.writeSeller("N/A", row)
			this.writeImage("N/A", row, write=False)
			return

		this.writeAsin(item.asin, row)
		this.writeTitle(item.title, row)
		this.writeCurrency(item.currency, row)
		this.writePrice(item.price, row)
		this.writeFullPrice(item.fullPrice, row)
		this.writeDiscount(item.discount, row)
		this.writeSeller(item.seller, row)
		this.writeImage(item, row)
		
	def changeOutFile(this, path):
		this.workbook.close() # Close the old workbook.
		this.workbook = xlsxwriter.Workbook(path)
		this.worksheet = this.workbook.add_worksheet()
		this.setColumnsDimensions()
		this.writeColumnsTitles()

	def updateCheckbox(this, checkBox):
		this.checkBox = checkBox

	def close(this):
		this.workbook.close()