import docx




def create_doc(name:str):
    doc = docx.Document()
    doc.save(name)



def open_doc(name:str):
    doc = docx.Document(name)
    return doc


def insert_logo(name:str):
    doc = open_doc('demo.docx')
    doc.add_picture(name)
    doc.save('demo.docx')




def main():
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



if __name__ == '__main__':
    main()