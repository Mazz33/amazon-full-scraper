import re
import requests
import time
from io import BytesIO
from PIL import Image
from bs4 import BeautifulSoup as bs
from amazoncaptcha import AmazonCaptcha

from headers import chooseRandomHeader as setHeader

class Item:
	def __init__(this, asin) :
		this.asin = asin
		this.none = False
		if (len(asin) < 10):
			this.notFound()
			return
		this.headers = setHeader()
		this.url = "https://www.amazon.ae/dp/{0}".format(asin)
		this.downloadWebpageContent()
		if (this.content == None):
			this.notFound()
			return
		this.configureSoup()
		this.extractItemTitle()
		this.extractItemPrice()
		this.extractFullPrice()
		this.extractDiscount()
		this.extractSeller()
		this.extractImage()
		this.extractWeight()

	def bypassCaptcha(this, content):
		soup = bs(content, "html.parser")
		try:
			captchaLink = soup.find_all("img")[0]["src"]
			captcha = AmazonCaptcha.fromlink(captchaLink)
			solution = captcha.solve()
			amzn = soup.find_all(attrs={"name": "amzn"})[0]["value"]
			amzn_r = soup.find_all(attrs={"name": "amzn-r"})[0]["value"]

			solveUrl = "https://amazon.ae/errors/validateCaptcha?amzn={0}&amzn-r={1}&field-keywords={2}".format(amzn, amzn_r, solution)
			solved = requests.get(solveUrl, headers=this.headers)
			if ("validateCaptcha" in solved.text):
				return this.bypassCaptcha(solved.content)
			return solved.content
		except:
			return None

	def notFound(this):
		this.none = True
		this.soup = None
		this.price = "N/A"
		this.currency = "N/A"
		this.seller = "N/A"
		this.discount = "N/A"
		this.weight = "N/A"
		this.fullPrice = "N/A"
		this.Image = "N/A"

	def downloadWebpageContent(this, content=None):
		if (content != None):
			this.content = content
			return
		try:
			req = requests.get(this.url, headers=this.headers)
			while (req.status_code == 503):
				req = requests.get(this.url.format(this.asin), headers=this.headers)
				time.sleep(0.5)
			if (req.status_code != 200):
				this.content = None
			if ("validateCaptcha" in req.text):
				this.content = this.bypassCaptcha(req.content)
				return
			if (req.status_code == 200):
				this.content = req.content
				return
			this.content = None
		except:
			this.content = None

	def configureSoup(this):
		this.soup = bs(this.content, "html.parser")

	def extractItemTitle(this):
		titleTag = this.soup.find_all(id="productTitle")
		if (len(titleTag) > 0):
			this.title = titleTag[0].text.strip()
			return
		this.title = "N/A"

	def extractItemPrice(this):
		tags = this.soup.find_all("span", "a-offscreen")	
		if (len(tags)> 0):
			fullPrice = tags[0].text
			if (fullPrice.find("AED") < 0):
				this.price = "N/A"
				this.currency = "N/A"
				return
			this.currency = fullPrice[:3]
			this.price = fullPrice[3:]
			return
		this.price = "N/A"
		this.currency = "N/A"

	def extractFullPrice(this):
		try:
			fullPriceSpan = this.soup.find_all("span", ["a-price", "a-text-price"])[-1]
			this.fullPrice = fullPriceSpan.find("span", "a-offscreen").text[3:]
		except IndexError:
			this.fullPrice = this.price

	def extractSeller(this):
		soldByDiv = this.soup.find_all("div", attrs={"tabular-attribute-name": "Sold by"})
		try:
			this.seller = soldByDiv[1].text.strip()
		except IndexError:
			this.seller = "N/A"

	def extractDiscount(this):
		discountSpan = this.soup.find_all("span", "savingsPercentage")
		try:
			discount = discountSpan[0].text
			this.discount = discount
		except IndexError:
			this.discount = "0%"

	def extractWeight(this):
		patternBullets = r";\s\d+(\.\d+)?\sk?g"
		patternDetails = r"\d+(\.\d+)?\s+((Kilograms)|(Grams))" # We are going to use regex for this one. This assumes the only metric unit used is KGs

		try:
			productDetails = this.soup.find("div", id="productDetails_feature_div") # We will first get the table div
			if (productDetails is None):
				raise Exception("Div not found")
			match = re.search(patternDetails, productDetails.text, re.IGNORECASE)
			if (match):
				this.weight = match.group().replace("Kilograms", "kg").replace("Grams", "g")
				return;

			match = re.search(patternBullets, productDetails.text, re.IGNORECASE)
			if (match):
				this.weight = match.group().split("; ")[-1]
				return;
		except Exception:
			pass

		try:
			# It wasn't found, so we will try to search the other field
			productDetails = this.soup.find("div", id="detailBulletsWrapper_feature_div")
			if (productDetails is None):
				raise Exception("Div not found")
			match = re.search(patternBullets, productDetails.text, re.IGNORECASE)
			if (match):
				this.weight = match.group().split("; ")[-1]
				return;
		except Exception:
			pass
		this.weight = "N/A"

	def extractImage(this):
		this.extractImageLink()
		if (this.imageLink is None):
			this.imageData = None
			return
		this.downloadImageData()
		this.calculateImageRatios()
		
	def extractImageLink(this): # for some weird reason amazon pictures can have different ids for multiple products. Have they heard of consistency?
		imgTag = this.soup.find_all(id="landingImage")
		if (len(imgTag) > 0):
			this.imageLink = imgTag[0]["src"]
			return
		imgTag = this.soup.find_all(id="imgBlkFront")
		if (len(imgTag) > 0):
			this.imageLink = imgTag[0]["src"]
			return
		divTag = this.soup.find_all(id="imgTagWrapperId")
		if (len(divTag) > 0):
			divTag = divTag[0]
		if (len(divTag) > 0):
			imgTag = divTag.find_all("img")
			if (len(imgTag) > 0):
				this.imageLink = imgTag[0]["src"]
				return
		this.imageLink = None

	def downloadImageData(this):
		req = requests.get(this.imageLink , headers=this.headers)
		this.imageData = BytesIO(req.content)

	def calculateImageRatios(this):
		cellWidth = 75
		cellHeight = 63

		image = Image.open(this.imageData)
		imageWidth, imageHeight = image.size

		this.x_scale = cellWidth / imageWidth
		this.y_scale = cellHeight / imageHeight
