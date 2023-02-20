from word_processor import WordProcessor, DocxEditor

WordProcessor = WordProcessor()
DocxEditor = DocxEditor()


class DocCreator:
    def __init__(self, isUrgent, indexList, placeholderList, file_name):

        self.dest_doc = WordProcessor.getDocx('')
        self.index = WordProcessor.getDocx('template/index.docx')
        self.urgent = WordProcessor.getDocx('template/urgent.docx')
        self.notice_of_motion = WordProcessor.getDocx(
            'template/notice_of_motion.docx')

        self.memo_of_parties = WordProcessor.getDocx(
            'template/memo_of_parties.docx')

        self.pet_synopsis = WordProcessor.getDocx(
            'template/pet_synopsis.docx')

        self.pet_title = WordProcessor.getDocx(
            'template/pet_title.docx')

        self.vakalatnama = WordProcessor.getDocx(
            'template/vakalatnama.docx')

        table_index = self.index.tables[1]
        for item in indexList:
            DocxEditor.add_index_table_row(table_index, ["\u2022", item, " "])

        WordProcessor.addToDocx(self.index, self.dest_doc)

        if isUrgent:
            WordProcessor.addPageBreak(self.dest_doc)  # page break
            WordProcessor.addToDocx(self.urgent, self.dest_doc)

        WordProcessor.addPageBreak(self.dest_doc)  # page break

        WordProcessor.addToDocx(self.notice_of_motion, self.dest_doc)

        WordProcessor.addPageBreak(self.dest_doc)  # page break

        WordProcessor.addToDocx(self.memo_of_parties, self.dest_doc)

        WordProcessor.addPageBreak(self.dest_doc)  # page break

        WordProcessor.addToDocx(self.pet_synopsis, self.dest_doc)

        WordProcessor.addPageBreak(self.dest_doc)  # page break

        WordProcessor.addToDocx(self.pet_title, self.dest_doc)

        WordProcessor.addPageBreak(self.dest_doc)  # page break

        WordProcessor.addToDocx(self.vakalatnama, self.dest_doc)

        DocxEditor.replace_placeholders(self.dest_doc, placeholderList)

        WordProcessor.saveTolocal(self.dest_doc, file_name)
