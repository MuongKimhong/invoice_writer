import xlsxwriter
import uuid


def invoice_writer(name, address, phone, date, items):
    random_number = str(uuid.uuid4())[0:4]

    workbook = xlsxwriter.Workbook('invoice_{}.xlsx'.format(name))
    worksheet = workbook.add_worksheet()

    info_format = workbook.add_format({
        'align': 'left',
        'valign': 'left',
        'font_size': 16,
        'font_color': '#003366'
    })
    c_format = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'font_size': 25,
        'font_color': '#003366'
    })
    
    content_merge_format = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'font_size': 14,
        'border_color': '#000000'
    })
    text_format = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'font_size': 14,
        'border_color': '#003366'
    })

    worksheet.merge_range('B2:H2', 'SalaCyber CO.,LTD', c_format)

    worksheet.merge_range('B3:H3', 'Address: ', info_format)
    worksheet.merge_range('B4:H4', 'Phone: 0865 211 536 / 077 222 033', info_format)
    worksheet.merge_range('B5:H5', 'Email: academy@salacyber.com', info_format)
    worksheet.merge_range('B6:H6', 'TIN: K007-902102016', info_format)

    worksheet.merge_range('B8:H8', 'Invoice', workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'font_size': 25,
        'font_color': '#FFFFFF',
        'bg_color': '#003366'
    }))

    worksheet.merge_range('B10:C10', 'Customer:', workbook.add_format({
        'align': 'left',
        'font_size': 14,
        'font_color': '#003366',
    }))
    # Customer name value
    worksheet.merge_range('D10:E10', name, workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))

    worksheet.merge_range('B11:C11', 'Address:', workbook.add_format({
        'align': 'left',
        'font_size': 14,
        'font_color': '#003366',
    }))
    # address value
    worksheet.merge_range('D11:E11', address, workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))

    worksheet.merge_range('B12:C12', 'Phone:', workbook.add_format({
        'align': 'left',
        'font_size': 14,
        'font_color': '#003366',
    }))
    # Customer phone value
    worksheet.merge_range('D12:E12', phone, workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))

    worksheet.merge_range('F10:G10', 'Issue No:', workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))
    worksheet.merge_range('F11:G11', 'Issued Date:', workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))
    # Issue No. value
    worksheet.write('H10', random_number, workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))
    # Issue date value
    worksheet.write('H11', date, workbook.add_format({
        'align': 'center',
        'font_size': 14,
        'font_color': '#003366',
    }))

    worksheet.merge_range('C14:D14', 'Description', content_merge_format)
    worksheet.merge_range('F14:G14', 'Unit Price', content_merge_format)
    worksheet.write('B14', 'No.', text_format)
    worksheet.write('E14', 'Qty', text_format)
    worksheet.write('H14', 'Amount', text_format)

    total_price = 0

    # input values from user go here
    for (i, item) in enumerate(items):
        total_price = total_price + item['amount']

        worksheet.write(f'B{i + 15}', i + 1, text_format)
        worksheet.merge_range(f'C{i+15}:D{i+15}', item['description'], content_merge_format)
        worksheet.merge_range(f'F{i+15}:G{i+15}', item['unit_price'], content_merge_format)
        worksheet.write(f'E{i+15}', item['quantity'], text_format)
        worksheet.write(f'H{i+15}', item['amount'], text_format)
    
    worksheet.merge_range(f'B{len(items) + 15}:G{len(items) + 15}', 'Total in USD', workbook.add_format({
        'align': 'right',
        'font_size': 14,
        'border_color': '#003366'
    }))
    worksheet.write(f'H{len(items) + 15}', total_price, text_format)

    workbook.close()


def main():
    total_amount = 0
    items = []

    name = input('Customer name: ')
    address = input('Customer address: ')
    phone = input('Customer phone: ')
    date = input('Issued Date (d/m/y): ')

    while True:    
        description = input('Description: ')
        quantity = input('Qty (number only): ')
        unit_price = input('Unit price (number only): ')
        amount = int(quantity) * int(unit_price)

        item = {
            'description': description, 
            'quantity': quantity,
            'unit_price': unit_price,
            'amount': amount,
        }
        items.append(item)

        print("Information is added")
        print("-----")
        continue_or_not = input('Do you want to add new item (y/n): ')

        if continue_or_not != 'n' and continue_or_not != 'y':
            print("Wrong answer!")
            break

        elif continue_or_not == 'n':     
            break            
    invoice_writer(name, address, phone, date, items)


if __name__ == '__main__':
    main()
