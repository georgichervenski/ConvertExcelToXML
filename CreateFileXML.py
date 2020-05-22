from lxml import etree as ET

top = ET.Element('all')
header = ET.SubElement(top, 'header')
reportCode = ET.SubElement(header, 'reportCode')
reportCode.text = 'RPS'
unitCode = ET.SubElement(header, 'unitCode')
unitCode.text = '4001'
unitName = ET.SubElement(header, 'unitName')
unitName.text = 'Айкарт АД'
periodFrom = ET.SubElement(header, 'periodFrom')
periodFrom.text = '01.01.2019'
periodTo = ET.SubElement(header, 'periodTo')
periodTo.text = '01.01.2020'


def add_tr_td(addDict):
    loopCodes = dict(addDict)
    for key, value in loopCodes.items():
        if key == 'Таблица 3.1':
            table = ET.SubElement(top, 'table')
            table.set('name', 'T_3.1')
            tr = ET.SubElement(table, 'tr')
            for nested_key, nested_value in dict(value).items():
                td = ET.SubElement(tr, 'td')
                td.set('code', nested_key)
                td.text = str(nested_value)
        elif key == 'Таблица 4.1':
            table = ET.SubElement(top, 'table')
            table.set('name', 'T_4.1')
            tr = ET.SubElement(table, 'tr')
            for nested_key, nested_value in dict(value).items():
                td = ET.SubElement(tr, 'td')
                td.set('code', nested_key)
                td.text = str(nested_value)
        elif key == 'Таблица 5.1':
            table = ET.SubElement(top, 'table')
            table.set('name', 'T_5.1')
            tr = ET.SubElement(table, 'tr')
            for nested_key, nested_value in dict(value).items():
                td = ET.SubElement(tr, 'td')
                td.set('code', nested_key)
                td.text = str(nested_value)
        else:
            table = ET.SubElement(top, 'table')
            for nested_key, nested_value in dict(value).items():
                tr = ET.SubElement(table, 'tr')
                td = ET.SubElement(tr, 'td')
                td.set('code', nested_key)
                td.text = str(nested_value)


contactPerson = ET.SubElement(top, 'contactPerson')
name = ET.SubElement(contactPerson, 'name')
name.text = 'Savka Georgieva'
phone = ET.SubElement(contactPerson, 'phone')
phone.text = '052706384'
email = ET.SubElement(contactPerson, 'email')
email.text = 'savka.georgieva@icard.com'
tree = ET.ElementTree(top)

def createXML(filename):
    setFilename = filename
    tree.write(setFilename, pretty_print=True, xml_declaration=True, encoding='Windows-1251')

"""
table = ET.SubElement(top, 'table')
def add_tr_td(addDict):
    loopCodes = dict(addDict)
    for key, value in loopCodes.items():
        for nested_key, nested_value in dict(value).items():
            tr = ET.SubElement(table, 'tr')
            td = ET.SubElement(tr, 'td')
            td.set('code', nested_key)
            td.text = str(nested_value)
"""