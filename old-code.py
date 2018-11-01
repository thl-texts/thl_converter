# From converter.py doMetadata line 88ff:
metaText = metaText.replace("{Title on Spine}", metaTable.cell(4, 3).text)
metaText = metaText.replace("{Margin Title}", metaTable.cell(5, 3).text)
metaText = metaText.replace("{Author of Text}", metaTable.cell(6, 3).text)
metaText = metaText.replace("{Name of Collection}", metaTable.cell(7, 3).text)
# PUBLISHING
metaText = metaText.replace("{Publisher Name}", metaTable.cell(9, 3).text)
metaText = metaText.replace("{Publisher Place}", metaTable.cell(10, 3).text)
metaText = metaText.replace("{Publisher Date}", metaTable.cell(11, 3).text)
# IDs
metaText = metaText.replace("{ISBN}", metaTable.cell(13, 3).text)
metaText = metaText.replace("{Library Call-number}", metaTable.cell(14, 3).text)
metaText = metaText.replace("{Other ID number}", metaTable.cell(15, 3).text)
metaText = metaText.replace("{Volume Letter}", metaTable.cell(16, 3).text)
metaText = metaText.replace("{Volume Number}", metaTable.cell(17, 3).text)
metaText = metaText.replace("{Pagination of Text}", metaTable.cell(18, 3).text)
metaText = metaText.replace("{Pages Represented in this file}", metaTable.cell(19, 3).text)
# CREATION
metaText = metaText.replace("{Name of Agent Creating Etext}", metaTable.cell(21, 3).text)
metaText = metaText.replace("{Date Process Begun}", metaTable.cell(22, 3).text)
metaText = metaText.replace("{Date Process Finished}", metaTable.cell(23, 3).text)
metaText = metaText.replace("{Place of Process}", metaTable.cell(24, 3).text)
metaText = metaText.replace("{Method of Process (OCR, input)}", metaTable.cell(25, 3).text)
# PROOFING
metaText = metaText.replace("{Name of Proofreader}", metaTable.cell(27, 3).text)
metaText = metaText.replace("{Date Proof Began}", metaTable.cell(28, 3).text)
metaText = metaText.replace("{Date Proof Finished}", metaTable.cell(29, 3).text)
metaText = metaText.replace("{Place of Proof}", metaTable.cell(30, 3).text)
# MARKUP
metaText = metaText.replace("{Name of Markup-er}", metaTable.cell(32, 3).text)
metaText = metaText.replace("{Date Markup Began}", metaTable.cell(33, 3).text)
metaText = metaText.replace("{Date Markup Finished}", metaTable.cell(34, 3).text)
metaText = metaText.replace("{Place of Markup}", metaTable.cell(35, 3).text)
# CONVERSION
metaText = metaText.replace("{Name of Converter}", metaTable.cell(37, 3).text)
metaText = metaText.replace("{Date Conversion Began}", metaTable.cell(38, 3).text)
metaText = metaText.replace("{Date Conversion Finished}", metaTable.cell(39, 3).text)
metaText = metaText.replace("{Place of Conversion}", metaTable.cell(40, 3).text)
# PROBLEMS
metaText = metaText.replace("{Problem Cell 1}", metaTable.cell(42, 1).text)
metaText = metaText.replace("{Problem Cell 2}", metaTable.cell(42, 3).text)

# Old getElement() function from Converter.PY around 1150:
def getElementOld(chStyle, lastElement, warn=True):
    global footnoteNum, endnoteNum

    # TODO Use the styleElements.py script to get definition of element based on style

    # TODO Check the list below versus sytleElements.py to make sure all chStyles have been accounted for
    if chStyle == "Added by Editor":
        elem = etree.SubElement(lastElement, "add")
        elem.set("n", "editor")

    elif chStyle == "Annotations":
        elem = etree.SubElement(lastElement, "note")
        elem.set("n", "annotation")

    elif chStyle == "Root Text":
        elem = etree.SubElement(lastElement, "seg")
        elem.set("type", "roottext")

    elif chStyle == "Sa bcad":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("type", "sabcad")

    elif chStyle == "Speech Inline":
        elem = etree.SubElement(lastElement, "q")

    elif chStyle == "Title (Own) Tibetan" or chStyle == "Colophon Text Title" or chStyle == "Text Title":
        elem = etree.SubElement(lastElement, "title")
        elem.set("type", "internal")
        elem.set("level", "m")
        elem.set("lang", "tib")

    elif chStyle == "Title (Own) Non-Tibetan Language" or chStyle == "Title (Own) Sanskrit":
        elem = etree.SubElement(lastElement, "title")
        elem.set("type", "internal")
        elem.set("level", "m")
        elem.set("lang", "non-tib")

    elif chStyle == "Title in Citing Other Texts":
        elem = etree.SubElement(lastElement, "title")
        elem.set("type", "external")
        elem.set("level", "m")

    elif chStyle == "Title of Chapter" or chStyle == "Colophon Chapter Title":
        elem = etree.SubElement(lastElement, "title")
        elem.set("type", "internal")
        elem.set("level", "a")
        elem.set("n", "chapter")

    elif chStyle == "Unclear" or chStyle == "z-DeprecatedUnclear":
        elem = etree.SubElement(lastElement, "unclear")

    elif chStyle == "X-Author Generic":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "Author")

    elif chStyle == "X-Author Indian":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "Author Indian")

    elif chStyle == "X-Author Tibetan":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "Author Tibetan")

    elif chStyle == "X-Dates" or chStyle == "Dates":
        elem = etree.SubElement(lastElement, "date")

    elif chStyle == "X-Doxo-Biblio Category" or chStyle == "Doxo-Biblio Category":
        elem = etree.SubElement(lastElement, "term")
        elem.set("n", "doxbibl")

    elif chStyle == "X-Emphasis Strong" or chStyle == "Emphasis Strong":
        elem = etree.SubElement(lastElement, "hi")
        elem.set("rend", "strong")

    elif chStyle == "X-Emphasis Weak" or chStyle == "Emphasis Weak":
        elem = etree.SubElement(lastElement, "hi")
        elem.set("rend", "weak")

    elif chStyle == "X-Mantra" or chStyle == "Mantra":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "Mantra")

    elif chStyle == "X-Monuments" or chStyle == "Monuments":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "Monuments")

    elif chStyle == "X-Name Buddhist Deity" or chStyle == "Name Buddhist  Deity":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "bud_deity")

    elif chStyle == "X-Name Buddhist Deity Collective":
        elem = etree.SubElement(lastElement, "orgName")
        elem.set("n", "bud_deity_collective")

    elif chStyle == "X-Name Clan" or chStyle == "Name Clan":
        elem = etree.SubElement(lastElement, "orgName")
        elem.set("n", "clan")

    elif chStyle == "X-Name Ethnicity" or chStyle == "Name Ethnicity":
        elem = etree.SubElement(lastElement, "orgName")
        elem.set("n", "ethnicity")

    elif chStyle == "X-Name Festival":
        elem = etree.SubElement(lastElement, "term")
        elem.set("n", "festival")

    elif chStyle == "X-Name Generic" or chStyle == "Name Generic":
        elem = etree.SubElement(lastElement, "term")

    elif chStyle == "X-Name Lineage" or chStyle == "Name Lineage":
        elem = etree.SubElement(lastElement, "term")
        elem.set("n", "lineage")

    elif chStyle == "X-Name Monastery" or chStyle == "Name organization monastery":
        elem = etree.SubElement(lastElement, "orgName")
        elem.set("n", "monastery")

    elif chStyle == "X-Name Organization" or chStyle == "Name Organization":
        elem = etree.SubElement(lastElement, "orgName")
        elem.set("n", "organization")

    elif chStyle == "X-Name Personal Human" or chStyle == "Name Personal Human":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("type", "human")

    elif chStyle == "X-Name Personal Other":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("type", "other")

    elif chStyle == "X-Name Place" or chStyle == "Name Place":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "place")

    elif chStyle == "X-Religious Practice" or chStyle.lower == "name ritual" or chStyle == "Name Religious Practice" or chStyle == "Religious Practice":
        elem = etree.SubElement(lastElement, "term")
        elem.set("n", "religious_practice")

    elif chStyle == "X-Speaker Buddhist Deity" or chStyle == "Speaker Buddhist Deity":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "speaker_bud_deity")

    elif chStyle == "X-Speaker Unknown":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "speaker_unknown")

    elif chStyle == "X-Speaker Human" or chStyle == "Speaker Human":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "speaker_human")

    elif chStyle == "X-Speaker Other" or chStyle == "Speaker Other":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "speaker_other")

    elif chStyle == "X-Term Chinese" or chStyle == "Lang Chinese":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "chi")

    elif chStyle == "X-Term English" or chStyle == "Lang English":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "eng")

    elif chStyle == "X-Term Mongolian":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "mon")

    elif chStyle == "X-Term Pali" or chStyle == "Lang Pali":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "pal")

    elif chStyle == "X-Term Sanskrit" or chStyle == "Lang Sanskrit":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "san")

    # guess for technical
    elif chStyle == "X-Term Technical":
        elem = etree.SubElement(lastElement, "term")
        elem.set("n", "technical")

    elif chStyle == "X-Term Tibetan" or chStyle == "Lang Tibetan":
        elem = etree.SubElement(lastElement, "term")
        elem.set("lang", "tib")

    elif chStyle == "X-Text Group" or chStyle == "Text Group":
        elem = etree.SubElement(lastElement, "title")
        elem.set("level", "s")
        elem.set("type", "group")

    # DEPRECATED LANGUAGES

    elif chStyle == "Lang French":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "fre")

    elif chStyle == "Lang German":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "ger")

    elif chStyle == "Lang Japanese":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "jap")

    elif chStyle == "Lang Korean":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "kor")

    elif chStyle == "Lang Nepali":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "nep")

    elif chStyle == "Lang Spanish":
        elem = etree.SubElement(lastElement, "rs")
        elem.set("lang", "spa")

    # DEPRECATED
    elif chStyle == "Speaker Generic":
        elem = etree.SubElement(lastElement, "persName")
        elem.set("n", "speaker")

    # not in new styles, but are in test doc
    elif chStyle == "Name river" or chStyle == "Name River":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "river")

    elif chStyle == "Name mountain" or chStyle == "Name Mountain":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "mountain")

    elif chStyle == "Name lake" or chStyle == "Name Lake":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "lake")

    elif chStyle == "Name geographical feature" or chStyle == "Name Geographical Feature":
        elem = etree.SubElement(lastElement, "placeName")
        elem.set("n", "geographical_feature")

    elif chStyle == "Pages":
        elem = etree.SubElement(lastElement, "num")
        elem.set("type", "pagerange")

    elif chStyle == "Document Map":
        # no warning
        return "none type"

    # Detect Footnote or Endnote Reference number and place the markedup note at that point in the text
    elif "Footnote" in chStyle or "footnote" in chStyle:
        footnoteNum += 1
        elem = etree.SubElement(lastElement, "note")
        elem.set("n", str(footnoteNum))

    elif "Endnote" in chStyle or "endnote" in chStyle:
        endnoteNum += 1
        elem = etree.SubElement(lastElement, "note")
        elem.set("n", str(endnoteNum))

    else:
        if warn is True:
            print "\t Warning (Character Style): " + chStyle + " is not supported"
        return "none type"

    return elem
