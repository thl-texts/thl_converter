#!/usr/local/bin/python
# -*- coding: utf-8 -*-

########## Python Library to Convert Style Names to Elements ##########

from lxml import etree
import json

##### Global Element Dictionary ######

# "keydict" is a dictionary of keys string matched with arrays of word style names
# The keys are the same keys in the element dictionary
# The values are an array of word styles names that should use that key to look up their element stats
# The function "createStyleKeyDict" creates a flat dictionary with keys of Word Style names matched with the "keys"
# used to look up the xml element specs (tagname and attributes)
keydict = {
    "abbr": ["Abbreviation"],
    "add-by-ed": ["Added by Editor"],
    "annotations": ["Annotations"],
    "auth-chi": ["author Chinese"],
    "auth-eng": ["author English"],
    "auth-gen": ["X-Author Generic"],
    "auth-ind": ["X-Author Indian"],
    "auth-san": ["author Sanskrit"],
    "auth-tib": ["X-Author Tibetan", "author Tibetan"],
    "author": ["author"],
    "date-range": ["Date", "Date Range", "Dates", "X-Dates"],
    "dates": ["X-Dates"],
    "dox-cat": ["Doxographical-Bibliographical Category", "X-Doxo-Biblio Category"],
    "emph-strong": ["Emphasis Strong", "X-Emphasis Strong", "Strong"],
    "emph-weak": ["Emphasis Weak", "Subtle Emphasis1", "X-Emphasis Weak"],
    "endn-char": ["Endnote Characters"],
    "endn-ref": ["endnote reference"],
    "endn-text": ["endnote text"],
    "endn-text-char": ["Endnote Text Char"],
    "epit-budd-deit": ["Epithet Buddhist Deity"],
    "epithet": ["Epithet"],
    "followedhyperlink": ["FollowedHyperlink"],
    "foot-bibl": ["Footnote Bibliography"],
    "foot-char": ["Footnote Characters"],
    "foot-ref": ["footnote reference"],
    "foot-text": ["footnote text"],
    "hyperlink": ["Hyperlink"],
    "illegible": ["Illegible"],
    "lang-chi": ["Lang Chinese"],
    "lang-eng": ["Lang English"],
    "lang-fre": ["Lang French"],
    "lang-ger": ["Lang German"],
    "lang-jap": ["Lang Japanese"],
    "lang-kor": ["Lang Korean"],
    "lang-mon": ["Lang Mongolian"],
    "lang-nep": ["Lang Nepali"],
    "lang-pali": ["Lang Pali"],
    "lang-sans": ["Lang Sanskrit"],
    "lang-span": ["Lang Spanish"],
    "lang-tib": ["Lang Tibetan"],
    "line-num": ["line number", "LineNumber"],
    "line-num-print": ["Line Number Print"],
    "line-num-tib": ["Line Number Tib"],
    "mantra": ["X-Mantra"],
    "monuments": ["Monuments", "X-Monuments"],
    "name-budd-deity": ["X-Name Buddhist  Deity", "Name Buddhist  Deity"],
    "name-budd-deity-coll": ["X-Name Buddhist Deity Collective", "Name Buddhist Deity Collective"],
    "name-clan": ["X-Name Clan"],
    "name-ethnic": ["X-Name Ethnicity", "Name of ethnicity"],
    "name-fest": ["Name festival", "X-Name Festival"],
    "name-gen": ["X-Name Generic", "Name generic"],
    "name-line": ["X-Name Lineage"],
    "name-monastery": ["X-Name Monastery"],
    "name-org": ["X-Name Organization", "Name organization"],
    "name-org-clan": ["Name org clan"],
    "name-org-line": ["Name org lineage"],
    "name-org-monastery": ["Name organization monastery"],
    "name-pers-human": ["X-Name Personal Human", "Name Personal Human"],
    "name-pers-other": ["Name Personal other", "X-Name Personal Other"],
    "name-place": ["X-Name Place", "Name Place"],
    "name-ritual": ["Name ritual"],
    "page-num": ["PageNumber", "page number"],
    "page-num-print-ed": ["Page Number Print Edition"],
    "pages": ["Pages"],
    "placeholder-text": ["Placeholder Text1"],
    "plain-text": ["Plain Text"],
    "pub-place": ["publication place"],
    "publisher": ["publisher"],
    "rel-pract": ["Religious practice", "X-Religious Practice"],
    "root-text": ["Root Text", "Root text"],
    "sa-bcad": ["Sa bcad"],
    "speak-budd-deity": ["X-Speaker Buddhist Deity", "SpeakerBuddhistDeity"],
    "speak-budd-deity-coll": ["Speaker Buddhist Deity Collective"],
    "speak-epit-budd-deity": ["Speaker Epithet Buddhist Deity"],
    "speak-human": ["X-Speaker Human", "SpeakerHuman"],
    "speak-other": ["SpeakerOther", "X-Speaker Other"],
    "speak-unknown": ["X-Speaker Unknown"],
    "speal-gene": ["Speaker generic"],
    "speech-inline": ["Speech Inline"],
    "term-chi": ["X-Term Chinese", "term Chinese"],
    "term-eng": ["term English", "X-Term English"],
    "term-fre": ["term French"],
    "term-ger": ["term German"],
    "term-jap": ["term Japanese"],
    "term-kor": ["term Korean"],
    "term-mon": ["term Mongolian", "X-Term Mongolian"],
    "term-nep": ["term Nepali"],
    "term-pali": ["term Pali", "X-Term Pali"],
    "term-sans": ["X-Term Sanskrit", "term Sanskrit"],
    "term-span": ["term Spanish"],
    "term-tech": ["X-Term Technical"],
    "term-tib": ["X-Term Tibetan", "term Tibetan"],
    "text-group": ["TextGroup", "X-Text Group"],
    "text-title": ["Text Title"],
    "text-title-chap-colon": ["Text Title in Chapter Colophon"],
    "text-title-colon": ["Text Title in Colophon"],
    "text-title-san": ["Text Title Sanksrit"],
    "text-title-tib": ["Text Title Tibetan"],
    "tib-line-number": ["TibLineNumber"],
    "title": ["Title"],
    "title-chap": ["Title of Chapter", "Colophon Chapter Title"],
    "title-cite-other": ["Title in Citing Other Texts"],
    "title-own-non-tib": ["Title (Own) Non-Tibetan Language"],
    "title-own-tib": ["Title (Own) Tibetan", "Colophon Text Titlle"],
    "unclear": ["Unclear"]
}

# Defining global dictionary for Word style to element keys to be populated by createStyleKeyDict() function
styledict = {}

# Elements is a keyed dictionary of information for defining XML elements
elements = {
    "abbr" : {
        "tag" : "abbr",
        "attributes" : {"expand" : "%TXT%"},
    },
    "add-by-ed" : {
        "tag" : "add",
        "attributes" : {"n" : "editor"},
    },
    "annotations" : {
        "tag" : "",
        "attributes" : {"type" : "annotation"},
    },
    "auth-chi" : {
        "tag" : "persName",
        "attributes" : { "type" : "author", "n" : "chinese"},
    },
    "auth-eng" : {
        "tag" : "persName",
        "attributes" : { "type" : "author", "n" : "engish"},
    },
    "auth-gen" : {
        "tag" : "persName",
        "attributes" : { "type" : "author", "n" : "general"},
    },
    "auth-ind" : {
        "tag" : "persName",
        "attributes" : { "type" : "author", "n" : "indian"},
    },
    "auth-san" : {
        "tag" : "persName",
        "attributes" : { "type" : "author", "n" : "sanskrit"},
    },
    "auth-tib" : {
        "tag" : "persName",
        "attributes" : { "type" : "author", "n" : "tibetan"},
    },
    "author" : {
        "tag" : "persName",
        "attributes" : { "type" : "author" },
    },
    "date-range" : {  # need to get converter to recognize split as what to split on and markup accordingly.
        "tag" : "dateRange",
        "attributes" : { "from": "%0%", "to": "%1%"},
        "split" : "-",
        "childels": "date",
    },
    "dox-cat" : {
        "tag" : "",
        "attributes" : {},
    },
    "emph-strong" : {
        "tag" : "",
        "attributes" : {},
    },
    "emph-weak" : {
        "tag" : "",
        "attributes" : {},
    },
    "endn-char" : {
        "tag" : "",
        "attributes" : {},
    },
    "endn-ref" : {
        "tag" : "",
        "attributes" : {},
    },
    "endn-text" : {
        "tag" : "",
        "attributes" : {},
    },
    "endn-text-char" : {
        "tag" : "",
        "attributes" : {},
    },
    "epit-budd-deit" : {
        "tag" : "",
        "attributes" : {},
    },
    "epithet" : {
        "tag" : "",
        "attributes" : {},
    },
    "followedhyperlink" : {
        "tag" : "",
        "attributes" : {},
    },
    "foot-bibl" : {
        "tag" : "",
        "attributes" : {},
    },
    "foot-char" : {
        "tag" : "",
        "attributes" : {},
    },
    "foot-ref" : {
        "tag" : "",
        "attributes" : {},
    },
    "foot-text" : {
        "tag" : "",
        "attributes" : {},
    },
    "hyperlink" : {
        "tag" : "",
        "attributes" : {},
    },
    "illegible" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-chi" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-eng" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-fre" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-ger" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-jap" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-kor" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-mon" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-nep" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-pali" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-sans" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-span" : {
        "tag" : "",
        "attributes" : {},
    },
    "lang-tib" : {
        "tag" : "",
        "attributes" : {},
    },
    "line-num" : {
        "tag" : "",
        "attributes" : {},
    },
    "line-num-print" : {
        "tag" : "",
        "attributes" : {},
    },
    "line-num-tib" : {
        "tag" : "",
        "attributes" : {},
    },
    "mantra" : {
        "tag" : "",
        "attributes" : {},
    },
    "monuments" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-budd-deity" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-budd-deity-coll" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-clan" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-ethnic" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-fest" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-gen" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-line" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-monastery" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-org" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-org-clan" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-org-line" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-org-monastery" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-pers-human" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-pers-other" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-place" : {
        "tag" : "",
        "attributes" : {},
    },
    "name-ritual" : {
        "tag" : "",
        "attributes" : {},
    },
    "page-num" : {
        "tag" : "",
        "attributes" : {},
    },
    "page-num-print-ed" : {
        "tag" : "",
        "attributes" : {},
    },
    "pages" : {
        "tag" : "",
        "attributes" : {},
    },
    "placeholder-text" : {
        "tag" : "",
        "attributes" : {},
    },
    "plain-text" : {
        "tag" : "",
        "attributes" : {},
    },
    "pub-place" : {
        "tag" : "",
        "attributes" : {},
    },
    "publisher" : {
        "tag" : "",
        "attributes" : {},
    },
    "rel-pract" : {
        "tag" : "",
        "attributes" : {},
    },
    "root-text" : {
        "tag" : "",
        "attributes" : {},
    },
    "sa-bcad" : {
        "tag" : "",
        "attributes" : {},
    },
    "speak-budd-deity" : {
        "tag" : "",
        "attributes" : {},
    },
    "speak-budd-deity-coll" : {
        "tag" : "",
        "attributes" : {},
    },
    "speak-epit-budd-deity" : {
        "tag" : "",
        "attributes" : {},
    },
    "speak-human" : {
        "tag" : "",
        "attributes" : {},
    },
    "speak-other" : {
        "tag" : "",
        "attributes" : {},
    },
    "speak-unknown" : {
        "tag" : "",
        "attributes" : {},
    },
    "speal-gene" : {
        "tag" : "",
        "attributes" : {},
    },
    "speech-inline" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-chi" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-eng" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-fre" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-ger" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-jap" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-kor" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-mon" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-nep" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-pali" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-sans" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-span" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-tech" : {
        "tag" : "",
        "attributes" : {},
    },
    "term-tib" : {
        "tag" : "",
        "attributes" : {},
    },
    "text-group" : {
        "tag" : "",
        "attributes" : {},
    },
    "text-title" : {
        "tag" : "",
        "attributes" : {},
    },
    "text-title-chap-colon" : {
        "tag" : "",
        "attributes" : {},
    },
    "text-title-colon" : {
        "tag" : "",
        "attributes" : {},
    },
    "text-title-san" : {
        "tag" : "",
        "attributes" : {},
    },
    "text-title-tib" : {
        "tag" : "",
        "attributes" : {},
    },
    "tib-line-number" : {
        "tag" : "",
        "attributes" : {},
    },
    "title" : {
        "tag" : "",
        "attributes" : {},
    },
    "title-chap" : {
        "tag" : "",
        "attributes" : {},
    },
    "title-cite-other" : {
        "tag" : "",
        "attributes" : {},
    },
    "title-own-non-tib" : {
        "tag" : "",
        "attributes" : {},
    },
    "title-own-tib" : {
        "tag" : "",
        "attributes" : {},
    },
    "unclear" : {}
}

def createStyleKeyDict(tolower=False):
    """
    Creates a dictionary keyed on Word Style name that returns the key for the univeral element array and stores in a global
    Returns the global if it's already populated. The dictionary returned is keyed on Word style name and returns the
    universal element key to use in the Element dictionary. This way more than one Word Style can have the same markup.
    The initial key dict has as its key the key to the element dictionary and as its values arrays of Word Style names.

    :param tolower: whether or not to lowercase the Word Style names used for keys in this dictionary
    :return: styledict: The flat one-to-one dictionary of Word Style Names (capitalized or all lower) and Element dictionary keys.
                        This can then be used to look up the Element definition for any Word Style Names
    """
    global keydict, styledict

    if len(styledict) == 0:
        for k in keydict.keys():
            styles = keydict[k]
            for stnm in styles:
                skey = stnm
                if tolower:
                    skey = skey.lower()
                styledict[skey] = k

    return styledict


def main():
    skl = createStyleKeyDict()

    print json.dumps(skl, indent=4)

if __name__ == "__main__":
    main()