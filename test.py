from word_processor import WordProcessor, DocxEditor


WordProcessor = WordProcessor()
DocxEditor = DocxEditor()

index = WordProcessor.getDocx('template/index.docx')
urgent = WordProcessor.getDocx('template/urgent.docx')
dest_doc = WordProcessor.getDocx('')

table = index.tables[1]
DocxEditor.add_index_table_row(table, ["\u2022", "  ANNEXURE P-1", " "])

WordProcessor.addToDocx(index, dest_doc)
WordProcessor.addPageBreak(dest_doc)
WordProcessor.addToDocx(urgent, dest_doc)


variables = {"{HIGHCOURT}": "Delhi HighCourt",
             "{JURIDICTION}": "Delhi Juridiction",
             "{PETITIONNUMBER}": "123456789",
             "{PETITIONERNAME}": "Alwin WK",
             "{RESPONDENTNAME}": "Some Name",
             "{PETTITLE}": "Delhi Petition",
             "{PETITIONERADDRESS2}": "123 Main St."}

DocxEditor.replace_placeholders(dest_doc, variables)
WordProcessor.saveTolocal(dest_doc)
