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