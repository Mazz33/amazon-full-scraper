# Amazon Full Scraper
A scraper that searches amazon based on a list of ASINs.  
The ASIN is the id used by amazon to identify items.  
It is tested on Linux and windows, but should work on Mac as well.  

## Requirements
Libraries needed: `PyQt6`, `pandas`, `openpyxl` `xlsxwriter`, `Pillow`, `BeautifulSoup4`, `amazoncaptcha`, `numpy`, `requests`

## Input
The application reads an column from an excel file, the column letter needs to be provided in the field.  
To open the excel file you just import it using the button, then specify the column letter to read from.  
After it imports you should see the list and number of ASINs imported on the gui.

## Output
You can choose where to save the file with the save as button. By default it will save the output in the same folder as `data.xlsx`.  

## Running
Run it with console (there is no output for console so it's useless):  
```
python app.py
```

On windows you can create a batch file as a launcher to start it without a gui:  
```batch
@echo off

start pythonw .\app.py
```
On linux:  
```sh
nohup python ./app.py &
```

You can stop it at any time by exiting it, it will save the finished data before exitting.  
After it finishes it will display a box saying it's done, and will exit after accepting.


## Threads
The application is hardcoded to max of 12 threads, obviously you can increase that value in the code, but this would increase the chances of you getting flagged and possibly IP banned, I had no issues with up to 12 threads, you can try more if you need.  

## Other regions
The application was built and tested for `amazon.ae`. Now you can try with another region by just editing the editing `Item.py` at line 18 and change the URL TLD.  
But it might not work always, the application is built using requests and bs4, the data is exctraced by finding the tags by id/class names to exctract data. There is no guarntee that other amazon regions use the same ids/classes. I believe amazon india uses different ids for example.  
If you need to run it for other regions, you will need to edit `item.py` and add thd appropriate tag names to extract the data.  