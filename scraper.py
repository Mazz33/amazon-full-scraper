import numpy as np
from threading import Thread
from Item import Item

def getAsinList():
	asinList = []
	with open("input.txt") as f:
		content = f.readlines()
	asinList = [x.strip() for x in content]
	return asinList

def amazonSearch(asinList, index, writer, updater):
	number = index
	for asin in asinList:
		number += 1
		item = Item(asin)
		updater()
		writer.writeItemData(item, number)

def getStartPoints(arr):
	startPoints = []
	for i in range(len(arr)):
		if (i==0):
			startPoints.append(1)
		else:
			if (len(arr[i]) == len(arr[i-1])):
				startPoints.append(startPoints[i-1] + len(arr[i]))
			else:
				startPoints.append(startPoints[i-1] + len(arr[i]) + 1)
	return startPoints
				
def createThreads(max):
	threads = []
	for t in range(max):
		t = Thread()
	return threads

def startSearch(asinList, threadCount, writer, updater):
	splitted = np.array_split(asinList, threadCount)
	startPoints = getStartPoints(splitted)
	threads = []
	for t in range(threadCount):
		t = Thread(target=amazonSearch, args=(splitted[t], startPoints[t], writer, updater,))
		threads.append(t)
		t.start()
	for t in threads:
		t.join()
