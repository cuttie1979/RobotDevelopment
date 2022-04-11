'''
#################################################################################################
# Program: Robot to extract fiduciary rates from an Excel and save them to another excel file   #
# Author: Popovics Laszlo                                                                       #
#################################################################################################
'''

'''
Open excel file

read BNP Paribas rates
read Deutsche Bank Luxemburg rates
read LGT rates
read BIL rates
read Societe General rates
read ANZ rates
read CIC rates
read QNB rates
save the results to a new excel file
'''

from RPA.Excel.Files import Files

def read_institution_rates(name,from_x, from_y, to_x, to_y, worksheet, column_names, title_row):
    result = {}
    print('Read data for: ' + name)
    for i in range(from_y,to_y + 1):
        currency = worksheet[i][column_names[from_x-1]]
        rates = {}
        for j in range(from_x, to_x + 1):
            value_label = str(worksheet[title_row][column_names[j]]).replace(' ','')
            value = worksheet[i][column_names[j]]
            rates[value_label] = value
        result[currency] = rates
    return result


lib = Files()
# Open excel file
lib.open_workbook('/Users/cuttie/Deltec/RobotDevelopment/Fiduciary Indicative rates.xls')

try:
    worksheet = lib.read_worksheet('SQ Fiduciary rates')
    column_names = []
    for key in worksheet[0]:
        column_names.append(key)
    # print(worksheet[10][column_names[2]]) - an example how we can show the content of row 11 and column 3

    # build result row
    results_header = ['Institution','Currency','1W', '2W','1M','2M','3M','6M','9M','12M','CALL']
    results = {}
    # read data for institutions
    results['BNP Paribas'] =                read_institution_rates('BNP Paribas',2 ,9 ,9 ,21 , worksheet, column_names, 7)
    results['Deutsche Bank Luxemburg'] =    read_institution_rates('Deutsche Bank Luxemburg', 13,9,18,13, worksheet, column_names, 7)
    results['LGT'] =                        read_institution_rates('LGT',24,9,32,18,worksheet, column_names, 7)
    results['BIL'] =                        read_institution_rates('BIL',2,29,9,32,worksheet, column_names, 28)
    results['Societe General rates'] =      read_institution_rates('Societe General rates',13,29,20,33, worksheet, column_names,28)
    results['ANZ'] =                        read_institution_rates('ANZ',24,29,32,34,worksheet, column_names, 28)
    results['CIC'] =                        read_institution_rates('CIC',5,41,13,45,worksheet, column_names, 40)
    results['QNB'] =                        read_institution_rates('QNB',20,41,27,45,worksheet, column_names, 40)
    # flattern results
    # print(results)
    lib.create_workbook('/Users/cuttie/Deltec/RobotDevelopment/Fiduciary_Indicative_rates_processed.xlsx')
    # create header row
    k = 1
    for h_value in results_header:
        lib.set_worksheet_value(row=0 + 1 , column=k, value=h_value)
        k = k + 1
    # fill data rows
    row = 2
    for instituion in results:
        inst_data = results[instituion]
        for currency in inst_data:
            currency_data = inst_data[currency]
            column = 0
            for cols in range(0,len(results_header)):
                if cols == 0:
                    lib.set_worksheet_value(row, cols + 1, instituion)
                elif cols == 1:
                    lib.set_worksheet_value(row, cols + 1, currency)
                elif results_header[cols] in currency_data.keys():
                    if '-' == currency_data[results_header[cols]]:
                        lib.set_worksheet_value(row, cols + 1, 'n/a')
                    else:    
                        lib.set_worksheet_value(row, cols + 1, currency_data[results_header[cols]])
                else:
                    lib.set_worksheet_value(row, cols + 1, 'n/a')
            row = row + 1
    lib.save_workbook()
finally:
    lib.close_workbook()
