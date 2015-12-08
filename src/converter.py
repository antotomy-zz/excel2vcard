from openpyxl import load_workbook
import vobject


def get_vcard(contact):
    card = vobject.vCard()
    card.add('fn')
    card.fn.value = contact['fullname']
    card.add('n')
    card.n.value = vobject.vcard.Name(
        contact['last_name'], contact['first_name'])
    card.add('tel')
    card.tel.type_param = 'cell'
    card.tel.value = contact['contact_no']
    return card.serialize()


def convert_xlsx_to_vcard(xlsx_file, column_map, start_row):
    wb = load_workbook(xlsx_file, read_only=True)
    sheet_name = wb.get_sheet_names()[0]
    ws = wb[sheet_name]
    row_no = 1
    vcf_file = ''

    for row in ws.rows:
        if row_no < start_row:
            row_no += 1
            continue

        first_name = row[column_map['first_name_column_no']].value
        last_name = row[column_map['last_name_column_no']].value
        contact_no = str(int(row[column_map['contact_no_column_no']].value))

        contact = {
            'fullname': '%s %s' % (first_name, last_name),
            'first_name': first_name,
            'last_name': last_name,
            'contact_no': contact_no
        }
        vcf_file += get_vcard(contact)

    return vcf_file


