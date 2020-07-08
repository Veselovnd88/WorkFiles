import docx
import openpyxl
import datetime


class WordTemplate:
    """Open template file and initializating all table's cells content
    for next activity"""

    def __init__(self, filename):
        """Конструктор, сразу инициализирует переменные"""
        self.filename = filename  # название открываемого файла
        self.doc = docx.Document(self.filename)  # открывает докх документ
        self.offer_num = self.doc.tables[1].rows[0].cells[0]  # место где находится номер ТКП
        self.customer_name = self.doc.tables[1].rows[0].cells[2]  #  место, где находится имя клиента
        self.main_table = self.doc.tables[2]  # основная таблица
        self.rows = 0  #переменная - номер строки, по умолчанию ноль, как только создается строка - +1
        self.offer_head = self.doc.tables[2].rows[0]
        self.pos_head = self.offer_head.cells[0].text
        self.name_head = self.offer_head.cells[1].text
        self.qnt_head = self.offer_head.cells[2].text
        self.deltime_head = self.offer_head.cells[3].text
        self.price_head = self.offer_head.cells[4].text
        self.total_head = self.offer_head.cells[5].text

    def create_main(self):
        """
        Заполнение констант ТКП
        :param offer_num: string
        :param customer_name: string
        """
        head_dict = ExcelParse().header()
        number = str(head_dict['Дата'])[:10] #TODO разобрать дату и преобразовать в номер

        self.offer_num.text = '№ '+head_dict['Имя ТКП']+' '+str(head_dict['Дата'])[:10] # TODO номер+дата
        self.customer_name.text = head_dict['Заказчик']
        # TODO - отцентровать по середине ячейки и посмотреть шрифты покрасивей
    def create_prices(self, price="(без НДС), евро", total='(без НДС), евро'):
        """
        Filling head of table
        :param price: str
        :param total: str
        """
        self.price_head.text = 'Цена ' + price
        self.total_head = 'Сумма ' + total

    def save(self, newfile=None):
        if newfile is None:
            self.doc.save(self.filename)
        else:
            self.doc.save(newfile)

    def generate_rows(self, params: list):
        self.main_table.add_rows()
        self.rows += 1
        new_row = self.main_table.rows[self.rows]
        pos = new_row.cells[0]
        name = new_row.cells[1]
        deltime = new_row.cells[3]
        price = new_row.cells[4]
        total = new_row.cells[5]



class ExcelParse:
    """
    Class for parsing excel File, each row - parameters for the filling word table
    """

    def __init__(self):
        self.resultdict = {}
        resultlist = []
        sheet = openpyxl.load_workbook('data.xlsx')
        self.worksheet = sheet["Actual"]
        print(self.worksheet.max_column)
        self.basic_head = {}


    def header(self):
        """Считывание основных данных для заполнения ТКП,
        читаем колонки
        1-Дата; 2-Заказчик; 3-Имя ТКп; 12-Сумма; 13-Тип цены
        14 - тип суммы
        15- условия оплаты
        16 - условия доставки
        17-Документация
        18 - Куратор
        """
        self.basic_cols = [1, 2, 3, 13, 14, 15, 16, 17, 18]
        for i in self.basic_cols:
            self.basic_head[i]=self.worksheet.cell(row=1, column = i).value
            self.resultdict[self.basic_head[i]] = self.worksheet.cell(row=2, column=i).value
        print(self.resultdict)

        return self.resultdict
    # for k in range(2, worksheet.max_row):
    #
    #     for i in basic_cols:
    #         print(i)


        # resultlist.append(resultdict)



def main():
    newdoc = WordTemplate('testoff.docx')
    newdoc.create_main()
    newdoc.save()
    # newex = ExcelParse()
    # newex.header()

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
