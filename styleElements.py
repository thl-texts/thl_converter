#!/usr/local/bin/python
# -*- coding: utf-8 -*-

########## Python Library to Convert Style Names to Elements ##########

from lxml import etree

##### Global Element Dictionary ######

# Elkeys matches style name to dictionary keys. This allows multiple style names to use the same element markup
keydict = {
    "Added by Editor" : "add-by-ed",
    "Annotations" : "annotations",
    "Colophon Chapter Title" : "title-chap",
    "Colophon Text Titlle" : "title-own-tib",
    "Date" : "date",
    "Date Range" : "date-range",
    "Dates" : "date-range",
    "Doxographical-Bibliographical Category" : "dox-cat",
    "Emphasis Strong" : "emph-strong",
    "Emphasis Weak" : "emph-weak",
    "Endnote Characters" : "endn-char",
    "Endnote Text Char" : "endn-text-char",
    "Epithet" : "epithet",
    "Epithet Buddhist Deity" : "epit-budd-deit",
    "FollowedHyperlink" : "followedhyperlink",
    "Footnote Bibliography" : "foot-bibl",
    "Footnote Characters" : "foot-char",
    "Hyperlink" : "hyperlink",
    "Illegible" : "illegible",
    "Lang Chinese" : "lang-chi",
    "Lang English" : "lang-eng",
    "Lang French" : "lang-fre",
    "Lang German" : "lang-ger",
    "Lang Japanese" : "lang-jap",
    "Lang Korean" : "lang-kor",
    "Lang Mongolian" : "lang-mon",
    "Lang Nepali" : "lang-nep",
    "Lang Pali" : "lang-pali",
    "Lang Sanskrit" : "lang-sans",
    "Lang Spanish" : "lang-span",
    "Lang Tibetan" : "lang-tib",
    "Line Number Print" : "line-num-print",
    "Line Number Tib" : "line-num-tib",
    "LineNumber" : "line-num",
    "Monuments" : "monuments",
    "Name Buddhist  Deity" : "name-budd-deity",
    "Name Buddhist Deity Collective" : "name-budd-deity-coll",
    "Name Personal Human" : "name-pers-human",
    "Name Personal other" : "name-pers-other",
    "Name Place" : "name-place",
    "Name festival" : "name-fest",
    "Name generic" : "name-gen",
    "Name of ethnicity" : "name-ethnic",
    "Name org clan" : "name-org-clan",
    "Name org lineage" : "name-org-line",
    "Name organization" : "name-org",
    "Name organization monastery" : "name-org-monastery",
    "Name ritual" : "name-ritual",
    "Page Number Print Edition" : "page-num-print-ed",
    "PageNumber" : "page-num",
    "Pages" : "pages",
    "Placeholder Text1" : "placeholder-text",
    "Plain Text" : "plain-text",
    "Religious practice" : "rel-pract",
    "Root Text" : "root-text",
    "Root text" : "root-text",
    "Sa bcad" : "sa-bcad",
    "Speaker Buddhist Deity Collective" : "speak-budd-deity-coll",
    "Speaker Epithet Buddhist Deity" : "speak-epit-budd-deity",
    "Speaker generic" : "speal-gene",
    "SpeakerBuddhistDeity" : "speak-budd-deity",
    "SpeakerHuman" : "speak-human",
    "SpeakerOther" : "speak-other",
    "Speech Inline" : "speech-inline",
    "Strong" : "emph-strong",
    "Subtle Emphasis1" : "emph-weak",
    "Text Title" : "text-title",
    "Text Title Sanksrit" : "text-title-san",
    "Text Title Tibetan" : "text-title-tib",
    "Text Title in Chapter Colophon" : "text-title-chap-colon",
    "Text Title in Colophon" : "text-title-colon",
    "TextGroup" : "text-group",
    "TibLineNumber" : "tib-line-number",
    "Title" : "title",
    "Title (Own) Non-Tibetan Language" : "title-own-non-tib",
    "Title (Own) Tibetan" : "title-own-tib",
    "Title in Citing Other Texts" : "title-cite-other",
    "Title of Chapter" : "title-chap",
    "Unclear" : "unclear",
    "X-Author Generic" : "auth-gen",
    "X-Author Indian" : "auth-ind",
    "X-Author Tibetan" : "auth-tib",
    "X-Dates" : "dates",
    "X-Doxo-Biblio Category" : "dox-cat",
    "X-Emphasis Strong" : "emph-strong",
    "X-Emphasis Weak" : "emph-weak",
    "X-Mantra" : "mantra",
    "X-Monuments" : "monuments",
    "X-Name Buddhist  Deity" : "name-budd-deity",
    "X-Name Buddhist Deity Collective" : "name-budd-deity-coll",
    "X-Name Clan" : "name-clan",
    "X-Name Ethnicity" : "name-ethnic",
    "X-Name Festival" : "name-fest",
    "X-Name Generic" : "name-gen",
    "X-Name Lineage" : "name-line",
    "X-Name Monastery" : "name-monastery",
    "X-Name Organization" : "name-org",
    "X-Name Personal Human" : "name-pers-human",
    "X-Name Personal Other" : "name-pers-other",
    "X-Name Place" : "name-place",
    "X-Religious Practice" : "rel-pract",
    "X-Speaker Buddhist Deity" : "speak-budd-deity",
    "X-Speaker Human" : "speak-human",
    "X-Speaker Other" : "speak-other",
    "X-Speaker Unknown" : "speak-unknown",
    "X-Term Chinese" : "term-chi",
    "X-Term English" : "term-eng",
    "X-Term Mongolian" : "term-mon",
    "X-Term Pali" : "term-pali",
    "X-Term Sanskrit" : "term-sans",
    "X-Term Technical" : "term-tech",
    "X-Term Tibetan" : "term-tib",
    "X-Text Group" : "text-group",
    "author" : "author",
    "author Chinese" : "auth-chi",
    "author English" : "auth-eng",
    "author Sanskrit" : "auth-san",
    "author Tibetan" : "auth-tib",
    "endnote reference" : "endn-ref",
    "endnote text" : "endn-text",
    "footnote reference" : "foot-ref",
    "footnote text" : "foot-text",
    "line number" : "line-num",
    "page number" : "page-num",
    "publication place" : "pub-place",
    "publisher" : "publisher",
    "term Chinese" : "term-chi",
    "term English" : "term-eng",
    "term French" : "term-fre",
    "term German" : "term-ger",
    "term Japanese" : "term-jap",
    "term Korean" : "term-kor",
    "term Mongolian" : "term-mon",
    "term Nepali" : "term-nep",
    "term Pali" : "term-pali",
    "term Sanskrit" : "term-sans",
    "term Spanish" : "term-span",
    "term Tibetan" : "term-tib"
}

# Elements is a keyed dictionary of information for defining XML elements
elements = {
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
    "date" : {
        "tag" : "date",
        "attributes" : {},
    },
    "date-range" : {
        "tag" : "dateRange",
        "attributes" : {},
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


def main():
    vals = keydict.values()
    vals.sort()
    valset = set(vals)
    valset = list(valset)
    valset.sort()
    for v in valset:
        print "    \"{0}\" ".format(v) + ': {
        "tag" : "",
        "attributes" : {},
    },'


if __name__ == "__main__":
    main()