import docx
import openpyxl

class WordTemplate:
    """Open template file and initializating all table's cells content
    for next activity"""

    def __init__(self, filename):
        self.filename = filename
        self.doc = docx.Document(self.filename)
        self.offer_num = self.doc.tables[1].rows[0].cells[0]
        self.customer_name = self.doc.tables[1].rows[0].cells[2]
        self.offer_head = self.doc.tables[2].rows[0]
        self.pos_head = self.offer_head.cells[0].text
        self.name_head = self.offer_head.cells[1].text
        self.qnt_head = self.offer_head.cells[2].text
        self.deltime_head = self.offer_head.cells[3].text
        self.price_head = self.offer_head.cells[4].text
        self.total_head = self.offer_head.cells[5].text

    def create_head(self, offer_num: str, customer_name: str):
        """
        Filling head of table
        :param offer_num: string
        :param customer_name: string
        """
        self.offer_num.text = offer_num
        self.customer_name.text = customer_name


    def create_prices(self, price= "(без НДС), евро",total='(без НДС), евро'):
        """
        Filling head of table
        :param price: str
        :param total: str
        """
        self.price_head.text = 'Цена '+ price
        self.total_head = 'Сумма ' + total


    def save(self, newfile=None):
        if newfile is None:
            self.doc.save(self.filename)
        else:
            self.doc.save(newfile)

    def generate_rows(self,obj):
        pass


class ExcelParse:
    """
    Class for parsing excel File, each row - parameters for the filling word table
    """
    pass


def main():
    newdoc = WordTemplate('testoff.docx')
    newdoc.create_head('N05jkl05', 'ТАhjkИФ')
    newdoc.save()
    """
    doc = open_doc('testoff.docx')
    offer_num = doc.tables[1].rows[0].cells[0].text
    customer_name = doc.tables[1].rows[0].cells[2].text
    offer_head = doc.tables[2].rows[0]
    pos_head = offer_head.cells[0].text
    name_head = offer_head.cells[1].text
    qnt_head = offer_head.cells[2].text
    deltime_head = offer_head.cells[3].text
    price_head = offer_head.cells[4].text
    total_head = offer_head.cells[5].text
    doc.tables[2].add_row()

    doc.save('testoff.docx')
"""


if __name__ == '__main__':
    main()
