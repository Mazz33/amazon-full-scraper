import xlsxwriter
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
			formattedArray.append(i[0])
		return formattedArray

	

class ExcelWriter:
	def __init__(this, outputFile="data.xlsx"):
		this.workbook = xlsxwriter.Workbook(outputFile)
		this.worksheet = this.workbook.add_worksheet()
		this.setColumnsDimensions()
		this.writeColumnsTitles()

	def setColumnsDimensions(this):
		this.worksheet.set_column("A:A", 12)
		this.worksheet.set_column("B:B", 50)
		this.worksheet.set_column("C:C", 12)
		this.worksheet.set_column("D:D", 7)
		this.worksheet.set_column("E:E", 10)

	def writeColumnsTitles(this):
		this.worksheet.write("A1", "ASIN")
		this.worksheet.write("B1", "Title")
		this.worksheet.write("C1", "Price")
		this.worksheet.write("D1", "Currency")
		this.worksheet.write("E1", "Picture")

	def writeItemData(this, item, row):
		if (item.none):
			this.worksheet.write_string("A{}".format(row), item.asin)
			this.worksheet.write_string("B{}".format(row), "N/A")
			this.worksheet.write_string("C{}".format(row), "N/A")
			this.worksheet.write_string("D{}".format(row), "N/A")
			this.worksheet.write_string("E{}".format(row), "N/A")
			return;

		this.worksheet.write_string("A{}".format(row), item.asin)
		this.worksheet.write_string("B{}".format(row), item.title)
		this.worksheet.write_string("C{}".format(row), item.price)
		this.worksheet.write_string("D{}".format(row), item.currency)
		if (item.imageData is None):
			this.worksheet.write_string("E{}".format(row), "N/A")
			return
		this.worksheet.insert_image("E{}".format(row), item.asin, {
		"object_position": 1,
		"url": item.url,
		"x_scale": item.x_scale,
		"y_scale": item.y_scale,
		"image_data": item.imageData
		})

	def changeOutFile(this, path):
		this.workbook.close() # Close the old workbook.
		this.workbook = xlsxwriter.Workbook(path)
		this.worksheet = this.workbook.add_worksheet()
		this.setColumnsDimensions()
		this.writeColumnsTitles()

	def close(this):
		this.workbook.close()