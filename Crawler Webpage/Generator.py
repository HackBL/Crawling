#!/usr/bin/python
# -*- coding: utf-8 -*-
import sys
import bs4 as bs
from bs4 import BeautifulSoup
import numpy as np
import urllib.request
import xlrd
import xlsxwriter
import os


def writeFile(arr):	# Open and write arr into file
	workbook = xlsxwriter.Workbook('References.xlsx')
	worksheet = workbook.add_worksheet()

	for col, data in enumerate(arr):
		row = 0
		worksheet.write_column(row, col, data)

	workbook.close()


def doesFileExists(filePath): # Check file exists
	return os.path.exists(filePath)


def reshapeArr(arr): # reshape array
	return list(map(list, zip(*arr)))


def combination(oldArr, newArr): # Append both arrays of arr
	oldArr = reshapeArr(oldArr)
	newArr = reshapeArr(newArr)
	finalArray = oldArr
	tempNew = [] # Generate array which has "N/A" for newArr

	for newTag in newArr[0]:	# Add new Tag
		if any(newTag in oldTag for oldTag in oldArr) == False:	
			finalArray[0].append(newTag) # Add new Tag into last position

			for i in range(1,len(finalArray)):	# Assign "N/A"" which doesn't have data on cur tag 
				finalArray[i].append('')

	for oldTag in oldArr[0]: 	# Assign "N/A" to newArr which doesn't have correspond tags with old arr
		if any(oldTag in newTag for newTag in newArr) == False:
			tempNew.append("")

		elif any(oldTag in newTag for newTag in newArr) == True: # Assign data with correspond index
			position = newArr[0].index(oldTag) # Index of oldArr which newArr doesn't have 

			tempNew.append(newArr[1][position])

	finalArray.append(tempNew)

	return finalArray


newArray = [] # Store new attr

# ---- URL starts ----
link = 'https://www.midea.cn/detail/index?itemid='

itemID = sys.argv[1] # Retrieve Data from Command line Argument

idArr = [['Item_ID',itemID]]

url = urllib.request.urlopen(link+itemID).read()

soup = bs.BeautifulSoup(url,'lxml')
# ---- URL ends ----

# ---- Spec tables starts ----
tables = soup.findAll("table", attrs={"class":"spec_table"})

for table in tables:
	if table.findParent("table") is None:
		rows = table.find_all('tr')
		for row in rows:
			cols = row.find_all('td')
			cols = [ele.text.strip() for ele in cols]
			newArray.append([ele for ele in cols if ele]) 
# ---- Spec tables ends ----

# ---- Title starts ----
title = soup.find("div", {"class": "product_right"}).find('h1').text.lstrip().rstrip() # Get title 

titleArr = [['Title',title]]
# ---- Title ends ----

# ---- Currency symbol starts ----
currency_symbol = soup.find("span", {"class": "currency_symbol"}).text
# ---- Currency symbol ends ----

# ---- Price starts ----
price = soup.find("span", {"class": "new_price"}).find('b').text # Get price   
priceArr = [['Price (' + currency_symbol + ')',price]]
# ---- Price ends ----

newArray = idArr + titleArr + priceArr +  newArray # Combine all Attr

# ---- Excel starts ----
if doesFileExists('./References.xlsx'):  # Update File
	loc = ("./References.xlsx")

	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)
	sheet.cell_value(0, 0)

	row = sheet.nrows # Update row
	col = sheet.ncols # Update col

	oldArray = [] # Exist data from excel

	for i in range(col): # Data excel -> array
		pairs = [] # Existed attr in the file

		for j in range(row):
			pairs.append(sheet.cell_value(j,i))

		oldArray.append(pairs)

	finalArray = reshapeArr(combination(oldArray,newArray)) # Combine array

	writeFile(finalArray)

else: # Open excel file
	writeFile(newArray)
# ---- Excel ends ----


	