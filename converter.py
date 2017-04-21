#!/usr/local/bin/python
# -*- coding: utf-8 -*-

from docx import Document
from lxml import etree
from datetime import date
import sys, os
#
# initialize global variables
tableOpen = citOpen = listOpen = lgOpen = speechOpen = False
bodyOpen = backOpen = frontOpen = False
useDiv1 = True
level = 0
isDir = False
curPage = ""
inDocument = False
total_nesting_level = 0
prevHeading = 0

# define functions
def doMetadata(metaTable):
	#load metadata schema
	try:
		f = open('teiHeader.dat', 'rb')
		metaText = f.read()
	except:
		print "\t Error: teiHeader.dat file not in current working directory"
		sys.exit(1)

	# fill metadata schema with data from metadata table in document 

	#BASIC METADATA
	metaText = metaText.replace("{Title of Text}", metaTable.cell(1, 3).text)	
	metaText = metaText.replace("{Cover Page}", metaTable.cell(2, 3).text)	
	metaText = metaText.replace("{Title on Cover}", metaTable.cell(3, 3).text)	
	metaText = metaText.replace("{Title on Spine}", metaTable.cell(4, 3).text)	
	metaText = metaText.replace("{Margin Title}", metaTable.cell(5, 3).text)	
	metaText = metaText.replace("{Author of Text}", metaTable.cell(6, 3).text)	
	metaText = metaText.replace("{Name of Collection}", metaTable.cell(7, 3).text)	
	#PUBLISHING
	metaText = metaText.replace("{Publisher Name}", metaTable.cell(9, 3).text)	
	metaText = metaText.replace("{Publisher Place}", metaTable.cell(10, 3).text)	
	metaText = metaText.replace("{Publisher Date}", metaTable.cell(11, 3).text)	
	#IDs
	metaText = metaText.replace("{ISBN}", metaTable.cell(13, 3).text)	
	metaText = metaText.replace("{Library Call-number}", metaTable.cell(14, 3).text)	
	metaText = metaText.replace("{Other ID number}", metaTable.cell(15, 3).text)	
	metaText = metaText.replace("{Volume Letter}", metaTable.cell(16, 3).text)	
	metaText = metaText.replace("{Volume Number}", metaTable.cell(17, 3).text)	
	metaText = metaText.replace("{Pagination of Text}", metaTable.cell(18, 3).text)	
	metaText = metaText.replace("{Pages Represented in this file}", metaTable.cell(19, 3).text)	
	#CREATION
	metaText = metaText.replace("{Name of Agent Creating Etext}", metaTable.cell(21, 3).text)	
	metaText = metaText.replace("{Date Process Begun}", metaTable.cell(22, 3).text)	
	metaText = metaText.replace("{Date Process Finished}", metaTable.cell(23, 3).text)	
	metaText = metaText.replace("{Place of Process}", metaTable.cell(24, 3).text)	
	metaText = metaText.replace("{Method of Process (OCR, input)}", metaTable.cell(25, 3).text)	
	#PROOFING
	metaText = metaText.replace("{Name of Proofreader}", metaTable.cell(27, 3).text)	
	metaText = metaText.replace("{Date Proof Began}", metaTable.cell(28, 3).text)	
	metaText = metaText.replace("{Date Proof Finished}", metaTable.cell(29, 3).text)	
	metaText = metaText.replace("{Place of Proof}", metaTable.cell(30, 3).text)	
	#MARKUP
	metaText = metaText.replace("{Name of Markup-er}", metaTable.cell(32, 3).text)	
	metaText = metaText.replace("{Date Markup Began}", metaTable.cell(33, 3).text)	
	metaText = metaText.replace("{Date Markup Finished}", metaTable.cell(34, 3).text)	
	metaText = metaText.replace("{Place of Markup}", metaTable.cell(35, 3).text)	
	#CONVERSION
	metaText = metaText.replace("{Name of Converter}", metaTable.cell(37, 3).text)	
	metaText = metaText.replace("{Date Conversion Began}", metaTable.cell(38, 3).text)	
	metaText = metaText.replace("{Date Conversion Finished}", metaTable.cell(39, 3).text)	
	metaText = metaText.replace("{Place of Conversion}", metaTable.cell(40, 3).text)	
	#PROBLEMS
	metaText = metaText.replace("{Problem Cell 1}", metaTable.cell(42, 1).text)	
	metaText = metaText.replace("{Problem Cell 2}", metaTable.cell(42, 3).text)	
	
	#DATE
	metaText = metaText.replace("{Digital Creation Date}", str(date.today()))	

	#if teiHeader.dat ever ends in something other than "<text>", change this
	metaText += "</text></TEI.2>"

	return metaText

def convertItalics(doc):
	return doc

def doHeaders(par, lastElement, root):
	global frontOpen, bodyOpen, backOpen, level, useDiv1, prevHeading
	styName = par.style.name
	closeStyle(styName, lastElement)

	#print styName

	#front
	#if ("Heading1_Front" in styName) or (("Heading 1" in styName) and ("Front" in par.text)) or ("Heading 0 Front" in styName):
	if "Heading 0 Front" == styName:
		lastElement = root.find('text')
		front = etree.SubElement(lastElement, "front")
		front.set("id", "a")
		head = etree.SubElement(front, "head")				
		head.text = par.text
		frontOpen = True
		level = 0
		return front
	#body
	#elif ("Heading1_Body" in styName) or (("Heading 1" in styName) and ("Body" in par.text)) or ("Heading 0 Body" in styName):
	elif "Heading 0 Body" == styName:
		lastElement = root.find('text')
		body = etree.SubElement(lastElement,"body")
		body.set("id", "b")
		head = etree.SubElement(body, "head")
		head.text = par.text
		frontOpen = False
		bodyOpen = True
		level = 0
		return body
	#back sections
	#elif ("Heading1_Back" in styName) or (("Heading 1" in styName) and ("Back" in par.text)) or ("Heading 0 Back" in styName):
	elif "Heading 0 Back" == styName:
		lastElement = root.find('text')
		back = etree.SubElement(lastElement, "back")
		back.set("id", "c")
		head = etree.SubElement(back, "head")
		head.text = par.text
		bodyOpen = False
		backOpen = True
		level = 0
		return back
	
	#do divs within the body
	elif "Heading" in styName:
		#if no front, body, or back divisions, insert default body division 
		if not frontOpen and not bodyOpen and not backOpen:
			lastElement = root.find('text')
			body = etree.SubElement(lastElement, "body")
			body.set("id", "b")
			bodyOpen = True
			lastElement = body
			level = 0
		
		headingNum = styName.split(" ")[1]
		
		if headingNum.isdigit():
			headingNum = int(headingNum) - 1			#take out minus 1 if we decide on Heading 0 -> Heading 1 -> Heading 2 -> etc. (right now this assumes Heading 0 -> Heading 2 -> Heading 3 -> etc.)
			curID = ""
			# push new nested level
			if headingNum > level-1:					#take out minus 1 if ^^
				level += 1
				# Chapter Level Elements (CLE's)
				if headingNum == 1:
					if lastElement.tag == "front":
						curID = "a1"
					elif lastElement.tag == "body":
						curID = "b1"
					elif lastElement.tag == "back":
						curID = "c1"
					#else:
					#	temp = lastElement.get("id").split(".")[0]
					#	curID = temp[0] + str(int(temp[1:])+1)
				# non-CLE's
				#else:
				#	curID = lastElement.get("id") + ".1"
				#	print "curID: " + str(curID)

			# pop out a level or remain on same level
			else:
				# Chapter Level Elements (CLE's)
				if headingNum == 1:
					if lastElement.tag == "front":
						curID = "a1"
					elif lastElement.tag == "body":
						curID = "b1"
					elif lastElement.tag == "back":
						curID = "c1"
					#else: 
					#	temp = lastElement.get("id").split(".")[0]
					#	curID = temp[0] + str(int(temp[1:])+1)
				# non-CLE's
				#else:
				#	curID = lastElement.get("id").split(".")[0][0] + str(int(lastElement.get("id").split(".")[-1]) + 1)

				while headingNum != level:
					level -= 1
					lastElement = lastElement.getparent()
				
				#if headingNum != 1:
				#	ln = len(lastElement.get("id").split(".")[-1])
				#	curID = lastElement.get("id")[:-ln] + str(int(lastElement.get("id").split(".")[-1]) + 1)
				
				lastElement = lastElement.getparent()

			if useDiv1:
				div = etree.SubElement(lastElement,"div"+str(level))
				div.set("id", curID)
			else:
				div = etree.SubElement(lastElement,"div")
				div.set("n",str(level))
				div.set("id", curID)
				#end of level comment
				lastElement.append(etree.Comment("Close of Level: " + str(level)))
	
			head = etree.SubElement(div,"head")
			head.text = par.text
			return div
		else:
			print "\t Error: Heading number of heading (" + styName + ") is NaN"

#fix
def doTable(par, t):
	lst = etree.SubElement(lastElement,"list")
	lst.set("rend","table")
	for r in t.rows:
		row = etree.SubElement(lst, "item")
		for c in r.cells:			
			cell = etree.SubElement(row, "rs")
			iterateRange(par, cell)
	return lastElement


def doNewList(styName, lastElement):
	lst = etree.SubElement(lastElement, "list")
	
	if "List Bullet" in styName:
		lst.set("type","unordered")
		lst.set("rend","bulleted")
	
	elif "List Number" in styName:
		lst.set("type","ordered")
		lst.set("rend","1")
	
	else:
		print "\t Error: " + styName + " is not a supported list type"
	
	return lst

def doNestedList(par, prevSty, lastElement):	
	global total_nesting_level

	styName = par.style.name
	prevListLevel = prevSty.split(" ")[-1]		
	curListLevel = styName.split(" ")[-1]

	try:
		prevListLevel = int(prevListLevel)
	except ValueError:
		prevListLevel = 1
	try:
		curListLevel = int(curListLevel)
	except ValueError:
		curListLevel = 1

	#total_nesting_level = curBulLevel + curNumLevel
	#total_nesting_level + prevListLevel < total_nesting_level + curListLevel



	#push new nested level
	if prevListLevel < curListLevel:
		lst = etree.SubElement(lastElement, "list")

		if "List Bullet" in styName:
			lst.set("type","unordered")
			lst.set("rend","bulleted")

		elif "List Number" in styName:
			lst.set("type","ordered")
			lst.set("rend","1")
	
	#pop out a level
	else:
		lst = etree.SubElement(lastElement.getparent(), "list") 
	
	return lst

def closeStyle(styName, lastElement):
	global citOpen, listOpen, lgOpen, speechOpen

	flag = False;

	if lgOpen and "Verse" not in styName:
		lgOpen = False
		flag = True
	if listOpen and "List" not in styName:
		listOpen = False
		total_nesting_level = 0
		flag = True
	if citOpen and "Citation" not in styName:
		citOpen = False
		flag = True
	if speechOpen and "Speech" not in styName:
		speechOpen = False
		flag = True

	if flag:
		return lastElement.getparent()
	
	return lastElement

def doParaStyles(par, prevSty, lastElement):
	global citOpen, listOpen, lgOpen, speechOpen, total_nesting_level 

	styName = par.style.name
	lastElement = closeStyle(styName, lastElement)

	#print styName

	if  "Paragraph" == styName:
		p = etree.SubElement(lastElement, "p")
		iterateRange(par, p)
		return lastElement
	
	
	## Check if closeStyles() already gets the parent for these "Continued" styles
	## Also check if this works for doubly (or more) nested features.
	## For example, if a citation in a paragraph ends with a list within a list and then there is a "Paragraph Citation Continued," wouldn't this make lastElement the first list instead of the paragraph?
	# if "Paragraph Continued" == styName or "Paragraph Citation Continued" == styName:
	# 	lastElement = lastElement.getparent()
	# 	p = etree.SubElement(lastElement, "p")
	# 	iterateRange(par, p)
	# 	return lastElement
	
	elif "Bibliography" == styName:
		bibl = etree.SubElement(lastElement, "bibl")
		iterateRange(par, bibl) 
		return lastElement
	
	## deprecated?
	# elif "Title Line" in styName:
	# 	titlePage = etree.SubElement(lastElement, "titlePage")
	# 	titlePart = etree.SubElement(titlePage, "titlePart")
	# 	iterateRange(par, titlePart)
	# 	return lastElement

	# textual citations
	elif "Citation" in styName:
		if not citOpen: #and "verse" not in styName:?
			lastElement = etree.SubElement(lastElement, "quote")
			citOpen = True

		#bulletted list in a citation
		if "Citation List Bullet" == styName: #CHANGE to List Bullet Citation
			if "Citation List Bullet" not in prevSty:
				lastElement = etree.SubElement(lastElement, "list")
				lastElement.set("rend", "bullet")
				listOpen = True
			item = etree.SubElement(lastElement, "item") 
			iterateRange(par, item) 
			return lastElement
		
		#numbered list in a citation
		elif "Citation List Number" == styName:	#CHANGE to List Number Citation
			if "Citation List Number" not in prevSty:
				lastElement = etree.SubElement(lastElement, "list")
				lastElement.set("rend", "1")
				listOpen = True
			item = etree.SubElement(lastElement, "item")  
			iterateRange(par, item) 
			return lastElement


		#citOpen at starts of clause takes care of inserting <quote> when any new quote is started...so Paragraph Citation & Paragraph Citation Continued have same behavior
		elif "Z-Depracated Paragraph Citation" == styName: #CHANGE to Paragaph Citation
			p = etree.SubElement(lastElement,"p")
			iterateRange(par, p)
			return lastElement
		elif "Citation Prose 2" == styName: #CHANGE to Paragraph Citation Continued
			p = etree.SubElement(lastElement, "p")
			iterateRange(par, p)
			return lastElement

		elif "Paragraph Citation Nested" == styName:
			quote = etree.SubElement(lastElement,"quote")
			p = etree.SubElement(quote, "p")
			iterateRange(par, quote)
			return quote
			#check on getting parent for closeStyle on this

		# fix all of these up to be more succinct

		elif "Citation Verse 1" == styName: #CHANGE to Verse Citation 1
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par, l) 
			lgOpen = True
			return lg

		elif "Citation Verse 2" == styName: #CHANGE to Verse Citation 2
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement

		elif "Citation Verse Nested 1" == styName: #CHANGE to Verse Citation Nested 1
			lastElement = lastElement.getparent()
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par,l)
			lgOpen = True
			return lg

		elif "Citation Verse Nested 2" == styName: #CHANGE to Verse Citation Nested 2
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement

		#insert Verse Citation Nested 1 & Verse Citation Nested 2
		
		else:
			print "\t Error: " + styName + " is not a supported Citation Style"

	elif "List" in styName:
		# start new list
		if not listOpen:
			lastElement = doNewList(styName, lastElement)
			total_nesting_level = 1
			listOpen = True
		# start new nested list
		elif prevSty != styName:
			lastElement = doNestedList(par, prevSty, lastElement)
		# add item to list
		item = etree.SubElement(lastElement,"item")
		iterateRange(par, item)
		return lastElement

	# speech citations (quotes)
	elif "Speech" in styName:
		if not speechOpen:
			q = etree.SubElement(lastElement,"q")
			if "Inline" in styName:
				q.set("rend","inline")
				iterateRange(par, q)
				return lastElement 						
			elif "Verse" in styName:
				lg = etree.SubElement(q,"lg")
				l = etree.SubElement(lg,"l")
				iterateRange(par, l)
				lgOpen = True
				return lg
			else:
				p = etree.SubElement(q,"p")
				iterateRange(par, p)
				return lastElement 					#or do we return p or q?
			speechOpen = True
		else:
			if "Verse 1" in styName:
				if lgOpen:
					lastElement = lastElement.getparent()
				lg = etree.SubElement(lastElement,"lg")	
				l = etree.SubElement(lg,"l")
				iterateRange(par, l)
				lgOpen = True
				return lg
			elif "Verse 2" in styName:
				l = etree.SubElement(lastElement,"l")		
				iterateRange(par, l)
				return lastElement
			else:
				p = etree.SubElement(lastElement,"p")
				iterateRange(par, p)
				return lastElement 					#or do we return p or q?

	elif "Verse" in styName:
		if "Verse 1" in styName:
			if lgOpen:
				lastElement = lastElement.getparent()
			lg = etree.SubElement(lastElement,"lg")	
			l = etree.SubElement(lg,"l")
			iterateRange(par, l)
			lgOpen = True
			return lg
		elif "Verse 2" in styName:
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement
		else:
			print "\t Error: " + styName + " is not a supported Verse Style"
			return lastElement

	else:
		print "\t Error: " + styName + " is not a supported Paragraph Style"
		p = etree.SubElement(lastElement, "p")
		iterateRange(par, p)
		return lastElement

def iterateRange(par, lastElement):
	styName = par.style.name
	runs = par.runs

	#empty paragraph (blank line)
	if len(runs)==0:
		lastElement.text = " "
		return

	curStyle = styName

	# iterate through runs in current paragraph
	for run in runs:
		charStyle = run.style.name
		# if character style of current run is same as current paragraph style
		if charStyle == styName or charStyle == "Default Paragraph Font":
			try:
				try:
					elem.tail += run.text
				except TypeError:
					elem.tail = run.text
			except UnboundLocalError:
				try:
					lastElement.text += run.text
				except TypeError:
					lastElement.text = run.text
			curStyle = styName

		elif charStyle == "Footnote Reference" and run.text != " ":
			note = etree.SubElement(lastElement,"note")
			iterateNote(run, note, styName)
			#curStyle = charStyle ?
		elif charStyle == "Hyperlink":
			xref = etree.SubElement(lastElement, "xref")	#check if this prints hyperlink
			xref.set("n", run.text)
			#handle displaying link text next to tag
			#curStyle = charStyle?
		else:
			# continue with same character style
			if charStyle == curStyle:
				try:
					try:
						elem.text += run.text
					except TypeError:
						elem.text = run.text
				except UnboundLocalError:
					try:
						lastElement.text += run.text
					except TypeError:
						lastElement.text = run.text
			

			# Page Number,digital
			elif "Page Number" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","digpage")
				elem.set("n",run.text)

			# Line Number,digital
			elif "Line Number" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","digline")
				elem.set("n",run.text)

			# Page Number Number Print Edition,pnp"
			elif "Page Number Print" in charStyle or "PageNumber" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","page")
				curPage = run.text[1:-1]
				elem.set("n","[page " + curPage + "]")

			# Line Number Print,lnp
			elif "Line Number Print" in charStyle or "TibLineNumber"==charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","line")
				#elem.set("n","[line " + curPage + "." + run.text[1:-1] + "]")
				elem.set("n",run.text)


			#Currently, citation text titles are included at end of previous paragraph, to change this uncomment the following
			#elif "Text Title" in charStyle:		
			#	elem = getElement(charStyle, lastElement.getparent())
			#	elem.text = run.text
			#	curStyle = charStyle
			# set new character style

			else:
				elem = getElement(charStyle, lastElement)
				elem.text = run.text
			curStyle = charStyle
			
#implement fully	
def iterateNote(run, lastElement, styName):
	#lastElement.text = run.text

	#FIX/TEST THIS
	styName = par.style.name
	curStyle = styName

	# iterate through characters in the footnote
	for char in run.text:
		charStyle = char.style.name
		# char style is same as paragraph style
		if charStyle == styName or charStyle == "Default Paragraph Font":
			try:
				try:
					elem.tail += char
				except TypeError:
					elem.tail = char
			except UnboundLocalError:
				try:
					lastElement.text += char
				except TypeError:
					lastElement.text = char
			curStyle = styName

		elif charStyle == "Hyperlink":
			xref = etree.SubElement(lastElement, "xref")	#check if this prints hyperlink
			xref.set("n", char)
			#handle displaying link text next to tag
			#curStyle = charStyle?
		else:
			# continue with same character style
			if charStyle == curStyle:
				try:
					elem.text += char
				except TypeError:
					elem.text = char
			
			elif charStyle == "Page Number" or charStyle == "PageNumber":
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","page")
				elem.set("n",char)
			
			# set new character style
			else:
				elem = getElement(charStyle, lastElement)
				elem.text = char

			curStyle = charStyle
			
def getElement(chStyle, lastElement):

	if chStyle == "Annotations":
		elem = etree.SubElement(lastElement,"note")
		elem.set("n","annotation")

	elif chStyle == "Dates":
		elem = etree.SubElement(lastElement,"date")
		
	elif chStyle == "Date Range":
		elem = etree.SubElement(lastElement,"dateRange")
		
	elif chStyle == "Doxographical-Bibliographical Category":
		elem = etree.SubElement(lastElement,"term")
		elem.set("type","doxbibl")
		
	elif chStyle == "Emphasis Strong":
		elem = etree.SubElement(lastElement,"hi")
		elem.set("rend","strong")
	
	elif chStyle == "Emphasis Weak":
		elem = etree.SubElement(lastElement,"hi")
		elem.set("rend","weak")
	
	elif chStyle == "Lang Chinese":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","chi")
	
	elif chStyle == "Lang English":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","eng")
	
	elif chStyle == "Lang French":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","fre")
	
	elif chStyle == "Lang German":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","ger")
	
	elif chStyle == "Lang Japanese":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","jap")
	
	elif chStyle == "Lang Korean":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","kor")
	
	elif chStyle == "Lang Nepali":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","nep")
	
	elif chStyle == "Lang Pali":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","pal")
	
	elif chStyle == "Lang Sanskrit":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","san")
	
	elif chStyle == "Lang Spanish":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","spa")
	
	elif chStyle == "Lang Tibetan":
		elem = etree.SubElement(lastElement,"foreign")
		elem.set("lang","tib")
	
	elif chStyle == "Monuments":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","monument")
	
	elif chStyle == "Name Buddhist Deity" or chStyle == "Name Buddhist  Deity":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("type","bud_deity")
	
	elif chStyle == "Name generic":
		elem = etree.SubElement(lastElement,"name")
	
	elif chStyle == "Name of ethnicity":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("type","ethnic")
	
	elif chStyle == "Name org clan":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("type","clan")
	
	elif chStyle == "Name org lineage":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("type","lineage")
		
	elif chStyle == "Name organization monastery":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("type","monastery")
		
	elif chStyle == "Name organization":
		elem = etree.SubElement(lastElement,"orgName")
		
	elif chStyle == "Name Personal Human":
		elem = etree.SubElement(lastElement,"persName")
		
	elif chStyle == "Name Personal other":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("type","other")
	
	elif chStyle == "Name Place":
		elem = etree.SubElement(lastElement,"placeName")
	
	elif chStyle == "Pages":
		elem = etree.SubElement(lastElement,"num")
		elem.set("type","pagination")

	elif chStyle == "Root text":
		elem = etree.SubElement(lastElement,"seg")
		elem.set("type","roottext")
	
	elif chStyle == "Speaker generic":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("type","speaker")
	
	elif chStyle == "SpeakerBuddhistDeity":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("type","speaker_bud_deity")
	
	elif chStyle == "SpeakerHuman":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("type","speaker_human")
	
	elif chStyle == "SpeakerOther":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("type","speaker_other")
	
	elif chStyle == "Text Title Sanksrit":
		elem = etree.SubElement(lastElement,"title")
		elem.set("lang","san")
		elem.set("level","m")
	
	elif chStyle == "Text Title Tibetan":
		elem = etree.SubElement(lastElement,"title")
		elem.set("lang","tib")
		elem.set("level","m")
	
	elif chStyle == "Text Title":
		elem = etree.SubElement(lastElement,"title")
		elem.set("level","m")
	
	elif chStyle == "TextGroup":
		elem = etree.SubElement(lastElement,"title")
		elem.set("level","s")
		elem.set("type","group")
	
	elif chStyle == "Topical Outline":
		elem = etree.SubElement(lastElement,"seg")
		elem.set("type","outline")
	
	elif chStyle == "Name Author":
		if "Bibliography" in styName:
			elem = etree.SubElement(lastElement,"author")
		else:
			elem = etree.SubElement(lastElement,"persName")
			elem.set("type","author")
	
	elif chStyle == "Code":
		elem = etree.SubElement(lastElement,"seg")
		elem.set("type","code")
	
	elif chStyle == "Reference":
		elem = etree.SubElement(lastElement,"ref")
		elem.set("type","bibl")
	
	elif chStyle == "term":
		elem = etree.SubElement(lastElement,"term")
	
	elif chStyle == "Title Article":
		elem = etree.SubElement(lastElement,"title")
		elem.set("level","a")
	
	elif chStyle == "Title Journal":
		elem = etree.SubElement(lastElement,"title")
		elem.set("level","j")
	
	elif chStyle == "Title Series":
		elem = etree.SubElement(lastElement,"title")
		elem.set("level","s")

	else:
		print "\t Error: " + chStyle + " is not a supported Character Style"
		elem = etree.SubElement(lastElement,"REPLACE")

	return elem

def convertDoc(inputFile):
	global tableOpen, citOpen, listOpen, lgOpen, speechOpen, bodyOpen, backOpen, frontOpen, useDiv1, level, curPage, inDocument, total_nesting_level
	# reset global variables
	tableOpen = citOpen = listOpen = lgOpen = speechOpen = False
	bodyOpen = backOpen = frontOpen = False
	useDiv1 = True
	level = 0
	curPage = ""
	inDocument = False
	total_nesting_level = 0

	# read input file
	document = Document(inputFile)

	# convert endnotes to footnotes?

	# check for unstylized italic usage
	document = convertItalics(document)

	# useDiv1		--> use numbered div (<div#> )	true by default
	# not useDiv1	--> use generic div (<div n=#>)	true for Tibetan style outlines
	#												'#' = nesting level
	for sty in document.styles:
		if "Heading 8" in sty.name:	
			useDiv1 = False
			break

	# process metadata table
	try:
		metaTable = document.tables[0]
	except:
		print "\t Error: metatable not included"
		sys.exit(1)
	metaText = doMetadata(metaTable)

	# create lxml element tree from metadata info
	metaText = metaText.encode('utf-8')
	parser = etree.XMLParser(ns_clean=True, recover=True, encoding='utf-8')
	root = etree.fromstring(metaText, parser=parser)
	#print etree.tostring(root, encoding='UTF-8', xml_declaration=False)

	# iterate through paragraphs
	lastElement = root.find('text')	#add attributes to <text> based on metadata?
	prevSty = ''
	for par in document.paragraphs:
		if "Heading" in par.style.name:
			#inDocument avoids including any paragraphs before the first structural heading
			inDocument = True
			lastElement = doHeaders(par, lastElement, root)
		#elif tableOpen:
			#if a table is open and it ends in this paragraph, set tableOpen to false
		#elif "<w:tbl>" in p.text: 			#p.containsTable:
			#doTable(par, t)
		elif inDocument:
			lastElement = doParaStyles(par, prevSty, lastElement)
		prevSty = par.style.name

	# create XML file
	name = inputFile.split("/")[-1].split(".")[0] + '.xml'
	file = open(name, "wb")
	toString = etree.tostring(root, pretty_print=True)
	file.write(toString);
	file.close()


# MAIN

if len(sys.argv)==0:
	print "\t Error: please include one or more docx files as command line arguments or the name of a folder that contains the files"
	sys.exit(0)

initialPath = os.path.join(os.getcwd(), sys.argv[1])

# user inputs folder with document(s)
if os.path.isdir(initialPath):
	for item in os.listdir(initialPath):
		currentPath = os.path.join(initialPath, item)
		# simply ignore non-docx files
		if item.endswith(".docx") and os.path.isfile(currentPath):
			print "Converting " + item + " to XML..."
			convertDoc(currentPath)
			print "Conversion successful!"
# user inputs docx file(s) directly
else:
	for item in sys.argv[1:]:
		currentPath = os.path.join(os.getcwd(), item)
		if item.endswith(".docx") and os.path.isfile(currentPath):
			print "Converting " + item + " to XML..."
			convertDoc(currentPath)
			print "Conversion successful!"
		else:
			print "\t Error: " + item + " is not a docx file in the current working directory"





