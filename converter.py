#!/usr/local/bin/python
# -*- coding: utf-8 -*-

########## LIBRARIES ##########
from docx import Document
from lxml import etree
from datetime import date
import sys, os

########## GLOBAL VARIABLES ##########
tableOpen = listOpen = lgOpen = citOpen = nestedCitOpen = speechOpen = nestedSpeechOpen = False
bodyOpen = backOpen = frontOpen = False
global_header_level = 0
isDir = False
inDocument = False
global_list_level = 0
noteNumber = 0
idTracker = []
titlePage = None

########## FUNCTIONS ##########

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

	metaText += "</text></TEI.2>"

	return metaText

def closingComments(lastElement):
	global global_header_level
	while global_header_level > 0:
				lastElement = lastElement.getparent()
				lastElement.append(etree.Comment("Close of Level: " + str(global_header_level)))
				global_header_level -= 1

def doTitle(par, lastElement):
	global titlePage
	titlePage = etree.Element("titlePage")
	docTitle = etree.SubElement(titlePage, "docTitle")
	titlePart = etree.SubElement(docTitle, "titlePart")
	titlePart.text = par.text

def doHeaders(par, lastElement, root):
	global frontOpen, bodyOpen, backOpen, global_header_level, idTracker
	styName = par.style.name
	#print styName
	closeStyle(styName, lastElement)

	skippedHeader = False

	# front
	if "Heading 0 Front" == styName:
		# close out previous paragraphs
		while global_header_level > 0:
	 		lastElement = lastElement.getparent()
	 		lastElement.append(etree.Comment("Close of Level: " + str(global_header_level)))
	 		global_header_level -= 1
	 	lastElement = lastElement.getparent()
		lastElement = root.find('text')
		# create Front section
		front = etree.SubElement(lastElement, "front")
		front.set("id", "a")
		idTracker = [0]
		head = etree.SubElement(front, "head")				
		head.text = par.text
		frontOpen = True
		# add title at the top of Front
		if titlePage is not None:
			front.append(titlePage)
		return front

	# body
	elif "Heading 0 Body" == styName:
		# close out previous paragraphs
		while global_header_level > 0:
				lastElement = lastElement.getparent()
				lastElement.append(etree.Comment("Close of Level: " + str(global_header_level)))
				global_header_level -= 1
		lastElement = lastElement.getparent()
		lastElement = root.find('text')
		# create Body section
		body = etree.SubElement(lastElement,"body")
		body.set("id", "b")
		idTracker = [0]
		head = etree.SubElement(body, "head")
		head.text = par.text
		frontOpen = False
		bodyOpen = True
		return body

	# back 
	elif "Heading 0 Back" == styName:
		# close out previous paragraphs
		while global_header_level > 0:
	  		lastElement = lastElement.getparent()
	  		lastElement.append(etree.Comment("Close of Level: " + str(global_header_level)))
	  		global_header_level -= 1
	  	lastElement = lastElement.getparent()
		lastElement = root.find('text')
		# create Back section
		back = etree.SubElement(lastElement, "back")
		back.set("id", "c")
		idTracker = [0]
		head = etree.SubElement(back, "head")
		head.text = par.text
		bodyOpen = False
		backOpen = True
		return back
	
	#do header divs within front/body/back
	elif "Heading" in styName:
		#if no front/body/back, print warning 
		if not frontOpen and not bodyOpen and not backOpen:
			print "\t Warning (IMPROPER HEADER NESTING): all Headings must be inside Front, Body, or Back"
			print "\t\t Header text: " + par.text
			
		headingNum = styName.split(" ")[1]
		
		if headingNum.isdigit():
			headingNum = int(headingNum)
			
			if headingNum-1 > global_header_level:
				print "\t Warning (IMPROPER HEADER NESTING): Jumped from Heading " + str(global_header_level) + " to Heading " + str(headingNum)
				print "\t\t Header text: " + par.text
				skippedHeader = True

			# push new nested level
			if headingNum > global_header_level:
				while headingNum != global_header_level:
					# Calculate <div> id
					curID = ""
					if frontOpen:
						curID = "a"
					elif bodyOpen:
						curID = "b"
					elif backOpen:
						curID = "c"

					if  global_header_level > len(idTracker)-1:
						idTracker.append(1)
					else:
						idTracker[global_header_level] += 1
					
					curID += str(idTracker[0])
					for i in range(1,global_header_level+1):
						curID += "." + str(idTracker[i])

					# push new level
					global_header_level += 1

					# Create <div> on current (newly pushed) level
					lastElement = etree.SubElement(lastElement,"div")
					lastElement.set("n",str(global_header_level))
					lastElement.set("id", curID)
					head = etree.SubElement(lastElement,"head")
					if not skippedHeader:
						head.text = par.text
					skippedHeader = False

				return lastElement

			# pop out a level or remain on same level
			else:
				#pop out as many levels as necessry
				while headingNum != global_header_level:
					lastElement = lastElement.getparent()
					lastElement.append(etree.Comment("Close of Level: " + str(global_header_level)))
					idTracker.pop()
					global_header_level -= 1

				# Calculate <div> id
				curID = ""
				if frontOpen:
					curID = "a"
				elif bodyOpen:
					curID = "b"
				elif backOpen:
					curID = "c"
			
				idTracker[global_header_level-1] += 1
		
				curID += str(idTracker[0])
				for i in range(1,global_header_level):
					curID += "." + str(idTracker[i])

				lastElement = lastElement.getparent()
				lastElement.append(etree.Comment("Close of Level: " + str(global_header_level)))

				#create div on current (or newly popped) level
				lastElement = etree.SubElement(lastElement,"div")
				lastElement.set("n",str(global_header_level))
				lastElement.set("id", curID)
				head = etree.SubElement(lastElement,"head")
				head.text = par.text
				return lastElement

		else:
			print "\t Warning: Heading number of heading (" + styName + ") is NaN"

# def doTable(par, t):
# 	lst = etree.SubElement(lastElement,"list")
# 	lst.set("rend","table")
# 	for r in t.rows:
# 		row = etree.SubElement(lst, "item")
# 		for c in r.cells:			
# 			cell = etree.SubElement(row, "rs")
# 			iterateRange(par, cell)
# 	return lastElement

def doNewList(styName, lastElement, cur_list_level):
	lst = etree.SubElement(lastElement, "list")
	
	if "List Bullet" in styName:
		lst.set("rend","bullet")
		lst.set("n",str(cur_list_level))
	elif "List Number" in styName:
		lst.set("rend","1")
		lst.set("n",str(cur_list_level))
	else:
		lst.set("n",str(cur_list_level))

	return lst

def closeStyle(styName, lastElement):
	global listOpen, lgOpen, citOpen, nestedCitOpen, speechOpen, nestedSpeechOpen, global_list_level

	flag = False

	if listOpen and "List" not in styName:
		listOpen = False
		flag = True
		while global_list_level > 1:
			global_list_level -= 1
			lastElement = lastElement.getparent().getparent()
		global_list_level = 0
		#this is for Citation List Bullet/Number (must pop out twice b/c of <quote> & <list>)
		#check if this breaks or works for cituations where there is a citation paragraph/verse in a regular list
		if citOpen and "Citation" not in styName:
				citOpen = False
				return lastElement.getparent().getparent()
		else:
			lastElement = lastElement.getparent()


	# Close Citation / Nested Citation
	if citOpen and "Citation" not in styName:
		citOpen = False
		if nestedCitOpen and "Nested" not in styName:
			nestedCitOpen = False
			if lgOpen:
				lgOpen = False
				return lastElement.getparent().getparent().getparent()
			else:
				return lastElement.getparent().getparent()
		else:
			if lgOpen:
				lgOpen = False
				return lastElement.getparent().getparent()
			else:
				return lastElement.getparent()
	elif nestedCitOpen and "Nested" not in styName:
		nestedCitOpen = False
		if lgOpen:
			return lastElement.getparent().getparent()
		else:
			return lastElement.getparent()
	
	# Close Speech / Nested Speech
	if speechOpen and "Speech" not in styName:
		speechOpen = False
		if nestedSpeechOpen and "Nested" not in styName:
			nestedSpeechOpen = False
			if lgOpen:
				lgOpen = False
				return lastElement.getparent().getparent().getparent()
			else:
				return lastElement.getparent().getparent()
		else:
			if lgOpen:
				lgOpen = False
				return lastElement.getparent().getparent()
			else:
				return lastElement.getparent()
	elif nestedSpeechOpen and "Nested" not in styName:
		nestedSpeechOpen = False
		if lgOpen:
			return lastElement.getparent().getparent()
		else:
			return lastElement.getparent()
	
	# Close Verse 1 / Verse 2
	if lgOpen and not("Verse 2" in styName or "Nested 1" in styName or "Nested 2" in styName or "1 Nested" in styName or "2 Nested" in styName):
		lgOpen = False
		return lastElement.getparent()
	else:
		return lastElement

def doParaStyles(par, prevSty, lastElement):
	global listOpen, lgOpen, citOpen, nestedCitOpen, speechOpen, nestedSpeechOpen, global_list_level 

	styName = par.style.name
	#print styName
	lastElement = closeStyle(styName, lastElement)

	if  "Paragraph" == styName or "Normal" == styName:
		p = etree.SubElement(lastElement, "p")
		iterateRange(par, p)
		return lastElement

	# fix this based on what Than/David say
	elif "Paragraph Continued" == styName or "ParagraphContinued" == styName:
	 	#lastElement = lastElement.getparent()
	 	p = etree.SubElement(lastElement, "p")
	 	iterateRange(par, p)
	 	return lastElement

	# textual citations
	elif "Citation" in styName:
		if not citOpen:
			lastElement = etree.SubElement(lastElement, "quote")
			citOpen = True

		#bulletted list in a citation
		if "Citation List Bullet" == styName or "List Bullet Citation" == styName:
			if "Citation List Bullet" not in prevSty:
				lastElement = etree.SubElement(lastElement, "list")
				lastElement.set("rend", "bullet")
				listOpen = True
			item = etree.SubElement(lastElement, "item") 
			iterateRange(par, item) 
			return lastElement
		
		#numbered list in a citation
		elif "Citation List Number" == styName or "List Number Citation" == styName:
			if "Citation List Number" not in prevSty:
				lastElement = etree.SubElement(lastElement, "list")
				lastElement.set("rend", "1")
				lastElement.set("n", "1")
				listOpen = True
			item = etree.SubElement(lastElement, "item")  
			iterateRange(par, item) 
			return lastElement


		#citOpen at starts of clause takes care of inserting <quote> when any new quote is started...so Paragraph Citation & Paragraph Citation Continued have same behavior
		elif "Z-Depracated Paragraph Citation" == styName or "Paragraph Citation" == styName:
			p = etree.SubElement(lastElement,"p")
			iterateRange(par, p)
			return lastElement
		elif "Citation Prose 2" == styName or "Paragraph Citation Continued" == styName:
			p = etree.SubElement(lastElement, "p")
			iterateRange(par, p)
			return lastElement

		elif "Paragraph Citation Nested" == styName:
			if "Paragraph Citation" not in prevSty and "Citation Prose 2" not in prevSty and "Z-Depracated Paragraph Citation" not in prevSty:
				print "\t Warning (IMPROPER CITATION NESTING): " + styName + " must be preceded by Paragraph Citation"
			nestedCitOpen = True
			quote = etree.SubElement(lastElement,"quote")
			iterateRange(par, quote)
			return quote

		elif "Citation Verse 1" == styName or "Verse Citation 1" == styName:
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par, l) 
			lgOpen = True
			return lg

		elif "Citation Verse 2" == styName or "Verse Citation 2" == styName:
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement

		elif "Citation Verse Nested 1" == styName or "Verse Citation Nested 1" == styName:
			if "Citation Verse" not in prevSty and "Verse Citation" not in prevSty:
				print "\t Warning (IMPROPER CITATION NESTING): " + styName + " must be preceded by a Verse Citation"
			nestedCitOpen = True
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par,l)
			lgOpen = True
			return lg

		elif "Citation Verse Nested 2" == styName or "Verse Citation Nested 2" == styName:
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement
		
		else:
			print "\t Warning (Paragraph Style): " + styName + " is not a supported Citation Style"

	elif "List" in styName:
		listOpen = True
		try:
			cur_list_level = int(styName.split()[-1])
		except ValueError:
			cur_list_level = 1

		if cur_list_level-1 > global_list_level:
			print "\t Warning (IMPROPER LIST NESTING): Jumped from List " + str(global_list_level) + " to List " + str(cur_list_level)
			print "\t\t List text: " + par.text

		# push new nested list level (or first level)
		if cur_list_level > global_list_level:
			while cur_list_level != global_list_level:
				if cur_list_level > 1:
					lastElement = etree.SubElement(lastElement,"item")
				global_list_level += 1
				lastElement = doNewList(styName, lastElement, cur_list_level)
		# pop out a level or remain on same level
		else:
			while cur_list_level != global_list_level:
				global_list_level -= 1
				lastElement = lastElement.getparent().getparent()
		item = etree.SubElement(lastElement, "item")
		iterateRange(par, item)
		return lastElement


	elif "Speech" in styName:

		if "Speech Inline" == styName:
			q = etree.SubElement(lastElement,"q")
			iterateRange(par, q)
			return lastElement

		if not speechOpen:
			lastElement = etree.SubElement(lastElement, "q")
			speechOpen = True

		if "Speech Prose" == styName or "Speech Paragraph" == styName: 
			p = etree.SubElement(lastElement,"p")
			iterateRange(par, p)
			return lastElement

		elif "Speech Prose Nested" == styName or "Speech Paragraph Nested" == styName:
			if "Speech Prose" not in prevSty and "Speech Paragraph" not in prevSty:
				print "\t Warning (IMPROPER SPEECH NESTING): " + styName +  "must be preceded by Speech Paragraph"
			nestedSpeechOpen = True
			q = etree.SubElement(lastElement,"q")
			iterateRange(par, q)
			return q

		elif "Speech Verse 1" == styName:
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par, l) 
			lgOpen = True
			return lg

		elif "Speech Verse 2" == styName:
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement

		elif "Speech Verse 1 Nested" == styName: 
			if "Speech Verse" not in prevSty:
				print "\t Warning (IMPROPER SPEECH NESTING): " + styName + " must be preceded by a Speech Verse"
			nestedSpeechOpen = True
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par,l)
			lgOpen = True
			return lg

		elif "Speech Verse 2 Nested" == styName: 
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement

		else:
			print "\t Warning (Paragraph Style): " + styName + " is not a supported Speech Style"
	 		return lastElement

	elif "Verse" in styName:
		if "Verse 1" == styName:
			lgOpen = True
			lg = etree.SubElement(lastElement,"lg")
			l = etree.SubElement(lg,"l")
			iterateRange(par, l) 
			return lg

		elif "Verse 2" == styName:
			l = etree.SubElement(lastElement,"l")
			iterateRange(par, l)
			return lastElement

		else:
			print "\t Warning (Paragraph Style): " + styName + " is not a supported Verse Style"
	 		return lastElement



	elif "Section" in styName:
		if "Interstitial" in styName:
	 		div = etree.SubElement(lastElement,"div")
	 		div.set("type","interstitial")
	 		head = etree.SubElement(div,"head")
	 		p = etree.SubElement(div,"p")
	 		p.text = par.text
	 		return lastElement
	 	ms = etree.SubElement(lastElement,"milestone")
	 	if "Chapter Element" in styName:
	 		ms.set("type","cle")
		elif "Top Level" in styName or "Division" in styName: 
	 		ms.set("unit","section")
	 		if "Second" in styName:
	 			ms.set("n","2")
	 		elif "Third" in styName:
	 			ms.set("n","3")
	 		elif "Fourth" in styName:
	 			ms.set("n","4")
	 		else:
	 			ms.set("n","1")
		else:
	 		print "\t Warning (Paragraph Style): " + styName + " is not a supported Section style"
	 	ms.set("rend",par.text)
	 	return lastElement



	else:
		print "\t Warning (Paragraph Style): " + styName + " is not supported"
		p = etree.SubElement(lastElement, "p")
		iterateRange(par, p)
		return lastElement

def iterateRange(par, lastElement):

	#styName is the current paragraph style
	styName = par.style.name
	runs = par.runs

	#empty paragraph (blank line)
	if len(runs)==0:
		lastElement.text = " "
		return

	#prevCharStyle is the most recent character style used. It is initialized to the current paragraph style.
	prevCharStyle = styName

	# iterate through runs in current paragraph
	for run in runs:
		#charStyle is the current character style.
		charStyle = run.style.name

		# place entire run in weak emphasis tag if italics
		#if run.italic:
		#	lastElement = etree.SubElement(lastElement,"hi")
		#	lastElement.set("rend","weak")

		# if character style of current run is same as current paragraph style
		if charStyle == styName or charStyle == "Default Paragraph Font":
			try:
				try:
					elem.tail += run.text
				except TypeError:
					elem.tail = run.text
			except (UnboundLocalError, AttributeError):
				try:
					lastElement.text += run.text
				except TypeError:
					lastElement.text = run.text
			#prevCharStyle = styName
			prevCharStyle = charStyle

		elif charStyle == "Footnote Text" or charStyle == "Endnote Text":
			note = etree.SubElement(lastElement,"note")
			noteNumber += 1
			note.set("n",noteNumber)
			iterateNote(run, note, styName)
			prevCharStyle = charStyle

		elif charStyle == "Hyperlink":
			xref = etree.SubElement(lastElement, "xref")	#check if this prints hyperlink
			xref.set("n", run.text)
			#handle displaying link text next to tag
			prevCharStyle = charStyle

		else:
			# Page Number,digital
			if "Page Number" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","page")
				temp = run.text
				if run.text[0]=="[":
					temp = temp[1:]
				if run.text[-1]=="]":
					temp = temp[:-1]
				elem.set("n",temp)
				elem.set("rend","digital")

			# Line Number,digital
			elif "Line Number" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","line")
				temp = run.text
				if run.text[0]=="[":
					temp = temp[1:]
				if run.text[-1]=="]":
					temp = temp[:-1]
				elem.set("n",temp)
				elem.set("rend","digital")


			elif "Page Number Print" in charStyle or "PageNumber" == charStyle:
				temp = run.text.replace("page","").replace("[","").replace("]","").strip()
				if len(temp)>0:
					elem = etree.SubElement(lastElement,"milestone")
					elem.set("unit","page")
					if "-" in temp:
						temp = temp.split("-")
						elem.set("n", temp[1])
						elem.set("ed", temp[0])
					else:
						elem.set("n", temp)

			# Line Number Print,lnp
			elif "Line Number Print" in charStyle or "TibLineNumber"==charStyle:
				temp = run.text.replace("line","").replace("[","").replace("]","").strip()
				if len(temp)>0:
					elem = etree.SubElement(lastElement,"milestone")
					elem.set("unit","line")
					elem.set("n",temp)

			elif charStyle == "Illegible":
				elem = etree.SubElement(lastElement,"gap")
				elem.set("n",run.text.split("[")[1].split("]")[0])
				elem.set("reason","illegible")

			# continue with same character style within same XML tag
			elif charStyle == prevCharStyle:
				try:
					try:
						elem.text += run.text
					except TypeError:
						elem.text = run.text
				except (UnboundLocalError, AttributeError):
					try:
						lastElement.text += run.text
					except TypeError:
						lastElement.text = run.text

			# new character style or no character style 
			else:
				elem = getElement(charStyle, lastElement)
				if elem == "none type":
					try:
						lastElement.text += run.text
					except TypeError:
						lastElement.text = run.text
				else:
					elem.text = run.text
			prevCharStyle = charStyle

		#pop out of emphasis tag if italics
		#if run.italic:
		#	lastElement = lastElement.getparent()

#implement fully	
def iterateNote(run, lastElement, styName):
	#lastElement.text = run.text
	#FIX/TEST THIS
	prevCharStyle = styName

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
			except (UnboundLocalError, AttributeError):
				try:
					lastElement.text += char
				except TypeError:
					lastElement.text = char
			prevCharStyle = styName

		elif charStyle == "Hyperlink":
			xref = etree.SubElement(lastElement, "xref")	#check if this prints hyperlink
			xref.set("n", char)
			#handle displaying link text next to tag
			prevCharStyle = charStyle
		else:
			# continue with same character style
			if charStyle == prevCharStyle:
				try:
					elem.text += char
				except TypeError:
					elem.text = char
			
			# Page Number,digital
			if "Page Number" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","page")
				temp = char
				if char[0]=="[":
					temp = temp[1:]
				if char[-1]=="]":
					temp = temp[:-1]
				elem.set("n",temp)
				elem.set("rend","digital")

			# Line Number,digital
			elif "Line Number" == charStyle:
				elem = etree.SubElement(lastElement,"milestone")
				elem.set("unit","line")
				temp = char
				if char[0]=="[":
					temp = temp[1:]
				if char[-1]=="]":
					temp = temp[:-1]
				elem.set("n",temp)
				elem.set("rend","digital")

			# Page Number Number Print Edition,pnp"
			elif "Page Number Print" in charStyle or "PageNumber" == charStyle:
				temp = run.text.replace("page","").replace("[","").replace("]","").strip()
				if len(temp)>0:
					elem = etree.SubElement(lastElement,"milestone")
					elem.set("unit","page")
					if "-" in temp:
						temp = temp.split("-")
						elem.set("n", temp[1])
						elem.set("ed", temp[0])
					else:
						elem.set("n", temp)

			# Line Number Print,lnp
			elif "Line Number Print" in charStyle or "TibLineNumber"==charStyle:
				temp = run.text.replace("line","").replace("[","").replace("]","").strip()
				if len(temp)>0:
					elem = etree.SubElement(lastElement,"milestone")
					elem.set("unit","line")
					elem.set("n",temp)


			elif charStyle == "Illegible":
				elem = etree.SubElement(lastElement,"gap")
				elem.set("n",run.text.split("[")[1].split("]")[0])
				elem.set("reason","illegible")

			# set new character style
			else:
				elem = getElement(charStyle, lastElement)
				if elem == "none type":
					try:
						lastElement.text += char
					except TypeError:
						lastElement.text = char
				else:
					elem.text = char
			prevCharStyle = charStyle
			
def getElement(chStyle, lastElement):

	if chStyle == "Added by Editor":
	 	elem = etree.SubElement(lastElement,"add")
	 	elem.set("n","editor")

	elif chStyle == "Annotations":
	 	elem = etree.SubElement(lastElement,"note")
	  	elem.set("n","annotation")

	elif chStyle == "Root Text":
	 	elem = etree.SubElement(lastElement,"seg")
	 	elem.set("type","roottext")

	elif chStyle == "Sa bcad":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("type","sabcad")

	elif chStyle == "Speech Inline":
		elem = etree.SubElement(lastElement,"q")

	elif chStyle == "Title (Own) Tibetan" or chStyle == "Colophon Text Title" or chStyle == "Text Title": 
		elem = etree.SubElement(lastElement,"title")
		elem.set("type","internal")
		elem.set("level","m")
		elem.set("lang","tib")
	
	elif chStyle == "Title (Own) Non-Tibetan Language" or chStyle == "Title (Own) Sanskrit":
	 	elem = etree.SubElement(lastElement,"title")
	 	elem.set("type","internal")
	 	elem.set("level","m")
	 	elem.set("lang","non-tib")

	elif chStyle == "Title in Citing Other Texts":
	 	elem = etree.SubElement(lastElement,"title")
	 	elem.set("type","external")
	 	elem.set("level","m")

	elif chStyle == "Title of Chapter" or chStyle == "Colophon Chapter Title":
		elem = etree.SubElement(lastElement,"title")
	 	elem.set("type","internal")
	 	elem.set("level","a")
	 	elem.set("n","chapter")

	elif chStyle == "Unclear" or chStyle == "z-DeprecatedUnclear":
		elem = etree.SubElement(lastElement,"unclear")

	elif chStyle == "X-Author Generic":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("n","Author")

	elif chStyle == "X-Author Indian":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("n","Author Indian")

	elif chStyle == "X-Author Tibetan":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("n","Author Tibetan")

	elif chStyle == "X-Dates" or chStyle == "Dates":
		elem = etree.SubElement(lastElement,"date")

	elif chStyle == "X-Doxo-Biblio Category" or chStyle == "Doxo-Biblio Category":
		elem = etree.SubElement(lastElement,"term")
		elem.set("n","doxbibl")

	elif chStyle == "X-Emphasis Strong" or chStyle == "Emphasis Strong":
	 	elem = etree.SubElement(lastElement,"hi")
	 	elem.set("rend","strong")
	
	elif chStyle == "X-Emphasis Weak" or chStyle == "Emphasis Weak":
	 	elem = etree.SubElement(lastElement,"hi")
		elem.set("rend","weak")

	elif chStyle == "X-Mantra" or chStyle == "Mantra":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","Mantra")

	elif chStyle == "X-Monuments" or chStyle == "Monuments":
		elem = etree.SubElement(lastElement,"placeName")
	 	elem.set("n","Monuments")

	elif chStyle == "X-Name Buddhist Deity" or chStyle == "Name Buddhist  Deity":
	 	elem = etree.SubElement(lastElement,"persName")
	 	elem.set("n","bud_deity")

	elif chStyle == "X-Name Buddhist Deity Collective":
		elem = etree.SubElement(lastElement,"orgName")
	 	elem.set("n","bud_deity_collective")

	elif chStyle == "X-Name Clan" or chStyle == "Name Clan":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("n","clan")

	elif chStyle == "X-Name Ethnicity" or chStyle == "Name Ethnicity":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("n","ethnicity")

	elif chStyle == "X-Name Festival":
		elem = etree.SubElement(lastElement,"term")
		elem.set("n","festival")

	elif chStyle == "X-Name Generic" or chStyle == "Name Generic":
		elem = etree.SubElement(lastElement,"term")

	elif chStyle == "X-Name Lineage" or chStyle == "Name Lineage":
		elem = etree.SubElement(lastElement,"term")  
		elem.set("n","lineage")

	elif chStyle == "X-Name Monastery" or chStyle == "Name organization monastery":
		elem = etree.SubElement(lastElement,"orgName")
	 	elem.set("n","monastery")

	elif chStyle == "X-Name Organization" or chStyle == "Name Organization":
		elem = etree.SubElement(lastElement,"orgName")
		elem.set("n","organization")

	elif chStyle == "X-Name Personal Human" or chStyle == "Name Personal Human":
		elem = etree.SubElement(lastElement,"persName") 
		elem.set("n","personal_human")

	elif chStyle == "X-Name Personal Other":
		elem = etree.SubElement(lastElement,"persName")
	 	elem.set("n","personal_other")

	elif chStyle == "X-Name Place" or chStyle == "Name Place":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","place")

	elif chStyle == "X-Religious Practice" or chStyle == "Name Ritual" or chStyle == "Name Religious Practice" or chStyle == "Religious Practice":
		elem = etree.SubElement(lastElement,"term") 
		elem.set("n","religious_practice")

	elif chStyle == "X-Speaker Buddhist Deity" or chStyle == "Speaker Buddhist Deity":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("n","speaker_bud_deity")

	elif chStyle == "X-Speaker Unknown":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("n","speaker_unknown")

	elif chStyle == "X-Speaker Human" or chStyle == "Speaker Human":
		elem = etree.SubElement(lastElement,"persName")
	 	elem.set("n","speaker_human")

	elif chStyle == "X-Speaker Other" or chStyle == "Speaker Other":
	 	elem = etree.SubElement(lastElement,"persName")
	 	elem.set("n","speaker_other")

	elif chStyle == "X-Term Chinese" or chStyle == "Lang Chinese":
		elem = etree.SubElement(lastElement,"rs") 
	 	elem.set("lang","chi")

	elif chStyle == "X-Term English" or chStyle == "Lang English":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","eng")

	elif chStyle == "X-Term Mongolian":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","mon")

	elif chStyle == "X-Term Pali" or chStyle == "Lang Pali":
	 	elem = etree.SubElement(lastElement,"rs")
	 	elem.set("lang","pal")

	elif chStyle == "X-Term Sanskrit" or chStyle == "Lang Sanskrit":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","san")

	#guess for technical
	elif chStyle == "X-Term Technical":
	 	elem = etree.SubElement(lastElement,"term")
	 	elem.set("n","technical")

	elif chStyle == "X-Term Tibetan" or chStyle == "Lang Tibetan":
	 	elem = etree.SubElement(lastElement,"term")
	 	elem.set("lang","tib")

	elif chStyle == "X-Text Group" or chStyle == "Text Group":
		elem = etree.SubElement(lastElement,"title")
	 	elem.set("level","s")
	 	elem.set("type","group")

	# DEPRECATED LANGUAGES

	elif chStyle == "Lang French":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","fre")
	
	elif chStyle == "Lang German":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","ger")
	
	elif chStyle == "Lang Japanese":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","jap")
	
	elif chStyle == "Lang Korean":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","kor")
	
	elif chStyle == "Lang Nepali":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","nep")
	
	elif chStyle == "Lang Spanish":
		elem = etree.SubElement(lastElement,"rs")
		elem.set("lang","spa")

	# DEPRECATED
	elif chStyle == "Speaker Generic":
		elem = etree.SubElement(lastElement,"persName")
		elem.set("n","speaker")

	# not in new styles, but are in test doc
	elif chStyle == "Name river" or chStyle == "Name River":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","river")

	elif chStyle == "Name mountain" or chStyle == "Name Mountain":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","mountain")

	elif chStyle == "Name lake" or chStyle == "Name Lake":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","lake")

	elif chStyle == "Name geographical feature" or chStyle == "Name Geographical Feature":
		elem = etree.SubElement(lastElement,"placeName")
		elem.set("n","geographical_feature")

	elif chStyle == "Pages":
		elem = etree.SubElement(lastElement,"num")
		elem.set("type","pagerange")


	elif chStyle == "Document Map":
		#no warning
		return "none type"

	elif (chStyle == "Footnote Reference" or chStyle == "footnote reference" or chStyle == "Endnote Reference" or chStyle == "endnote reference"):
		#no warning
		return "none type"

	else:
	 	print "\t Warning (Character Style): " + chStyle + " is not supported"
		return "none type"

	return elem

def convertDoc(inputFile):
	global tableOpen, listOpen, lgOpen, citOpen, nestedCitOpen, speechOpen, nestedSpeechOpen, bodyOpen, backOpen, frontOpen, global_header_level, inDocument, global_list_level, noteNumber, idTracker, titlePage
	# reset global variables
	tableOpen = listOpen = lgOpen = citOpen = nestedCitOpen = speechOpen = nestedSpeechOpen = False
	bodyOpen = backOpen = frontOpen = False
	global_header_level = 0
	inDocument = False
	global_list_level = 0
	noteNumber = 0
	idTracker = []
	titlePage = None

	# read input file
	document = Document(inputFile)

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
	lastElement = root.find('text')
	prevSty = ''
	for par in document.paragraphs:
		if "Heading" in par.style.name:
			#inDocument avoids including any paragraphs before the first structural heading
			inDocument = True
			lastElement = doHeaders(par, lastElement, root)
		elif inDocument:
			lastElement = doParaStyles(par, prevSty, lastElement)
		elif "Title" == par.style.name:
			doTitle(par, lastElement)
		else:
			print "\t Warning (IMPROPER HEADER NESTING): All paragraphs other than Title must be inside Front, Body, or Back"
			print "\t\t Paragraph text: " + par.text
		prevSty = par.style.name

	closingComments(lastElement)

	# create XML file
	name = inputFile.split("/")[-1].split(".")[0] + '.xml'
	file = open(name, "wb")
	docType = "<!DOCTYPE TEI.2 SYSTEM \"http://www.thlib.org:8080/cocoon/texts/catalogs/xtib3.dtd\">"
	toString = etree.tostring(root, encoding='UTF-8', xml_declaration=True, doctype=docType, pretty_print=True)
	file.write(toString);
	file.close()



########## MAIN ##########

if len(sys.argv)==0:
	print "\t Argument Error: please include one or more docx files as command line arguments or the name of a folder that contains the files"
	sys.exit(0)

initialPath = os.path.join(os.getcwd(), sys.argv[1])

# user inputs folder with document(s)
if os.path.isdir(initialPath):
	for item in os.listdir(initialPath):
		currentPath = os.path.join(initialPath, item)
		if item.endswith(".docx") and os.path.isfile(currentPath):
			print "Converting " + item + " to XML..."
			convertDoc(currentPath)
			print "Conversion successful!"
		# print warning for non-docx files
		else:
			print "\t Warning (IMPROPER ARGUMENT): " + item + " is not a docx file in the current working directory"
# user inputs docx file(s) directly
else:
	for item in sys.argv[1:]:
		currentPath = os.path.join(os.getcwd(), item)
		if item.endswith(".docx") and os.path.isfile(currentPath):
			print "Converting " + item + " to XML..."
			convertDoc(currentPath)
			print "Conversion successful!"
		# print warning for non-docx files
		else:
			print "\t Warning (IMPROPER ARGUMENT): " + item + " is not a docx file in the current working directory"





