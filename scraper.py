import sys
import xlsxwriter
import numpy as np
from threading import Thread
from Item import Item

NUMBER_OF_THREADS = 6

workbook = xlsxwriter.Workbook("./data.xlsx")
worksheet = workbook.add_worksheet()
worksheet.set_default_row(50)
worksheet.set_column("A:A", 12)
worksheet.set_column("B:B", 50)
worksheet.set_column("C:C", 12)
worksheet.set_column("D:D", 7)
worksheet.set_column("E:E", 10)

worksheet.write("A1", "ASIN")
worksheet.write("B1", "Title")
worksheet.write("C1", "Price")
worksheet.write("D1", "Currency")
worksheet.write("E1", "Picture")

def getAsinList():
	asinList = []
	with open("input.txt") as f:
		content = f.readlines()
	asinList = [x.strip() for x in content]
	return asinList

def writeExcel(item, number):
	if (item.none):
		worksheet.write("A{}".format(number), item.asin)
		worksheet.write("B{}".format(number), "N/A")
		worksheet.write("C{}".format(number), "N/A")
		worksheet.write("D{}".format(number), "N/A")
		worksheet.write("E{}".format(number), "N/A")
		return
	
	worksheet.write_string("A{}".format(number), item.asin)
	worksheet.write_string("B{}".format(number), item.title)
	worksheet.write_string("C{}".format(number), item.price)
	worksheet.write_string("D{}".format(number), item.currency)
	if (item.imageData is None):
		worksheet.write_string("E{}".format(number), "N/A")
		return
	worksheet.insert_image("E{}".format(number), item.asin, {
		"object_position": 1,
		"url": item.url,
		"x_scale": item.x_scale,
		"y_scale": item.y_scale,
		"image_data": item.imageData
	})
	
def amazonSearch(asinList, index, startPoints):
	number = startPoints[index]
	for asin in asinList:
		number += 1
		item = Item(asin)
		writeExcel(item, number)

def main():
	print("Reading ASIN List!")
	asinList = getAsinList()
	arr = np.array_split(asinList, NUMBER_OF_THREADS)
	startPoints = []
	for i in range(len(arr)):
		if (i == 0):
			startPoints.append(1)
		else:
			if (len(arr[i]) == len(arr[i-1])):
				startPoints.append(startPoints[i-1] + len(arr[i]))
			else:
				startPoints.append(startPoints[i-1] + len(arr[i]) + 1)
	threads = []
	for t in range(NUMBER_OF_THREADS):
		t = Thread(target=amazonSearch, args=(arr[t],t,startPoints,))
		threads.append(t)
		t.start()
	for t in threads:
		t.join()
	print("Done!!")

if __name__ == "__main__":
	print("Use Ctrl + C to exit")
	try:
		main()
	except KeyboardInterrupt:
		print("Ctrl + C Detected")
		workbook.close()
		sys.exit("File Written.\nExiting. Bye o/")

	except Exception as E:
		print("Error!! {}".format(E))
		workbook.close()
		sys.exit("File Written.\nExiting. Bye o/")

workbook.close()