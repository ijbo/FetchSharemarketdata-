from googlefinance import getQuotes
import xlsxwriter

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demo2.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:B', 30)
worksheet.set_column('C:H', 10)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Add a format. Light red fill with dark red text.
format1 = workbook.add_format({'bg_color': '#FFC7CE','font_color': '#9C0006'})

# Add a format. Green fill with dark green text.
format2 = workbook.add_format({'bg_color': '#C6EFCE','font_color': '#006100'})

# Add column header
worksheet.write(0, 0, 'Date',bold)
worksheet.write(0, 1, 'Stock Name',bold)
worksheet.write(0, 2, 'Current Price',bold)
worksheet.write(0, 3, 'Purchase Price',bold)
worksheet.write(0, 4, 'Price Change',bold)
worksheet.write(0, 5, 'Value at Market Price',bold)
worksheet.write(0, 6, 'Number of Stocks',bold)
worksheet.write(0, 7, 'Value at Cost',bold)
worksheet.write(0, 8, 'Percentage Allocation',bold)
worksheet.write(0, 9, 'Earning',bold)

with open('/Users/jyotibo/Desktop/andrewng/TF/Prog/stock_details.txt') as fp:
    lines = fp.read().split("\n")
    myList=""
    total_at_cost=0
    for i,vlines in enumerate(lines) :
      #  print(i,lines[i])
        for j in range(4) :
            clines=lines[i].split(",")
            if (j == 1 ):
                #print(clines[j-1],clines[j])
                myList +=clines[j]+","
                #clines[j + 1]) : Quantity : (clines[j+2]) : purchase price #total : Sum invested
                total_at_cost += round(float(clines[j + 1]) * float(clines[j+2]), 2)

#print(myList)
#print (json.dumps(getQuotes('AAPL'),indent=2))
#print (json.dumps(getQuotes([myList]), indent=36))

stock_data=getQuotes([myList])
#print (stock_data)
clines=""
k=0
total_at_market_value=0

for k in range (i+1):
    for j in range (4):
        clines = lines[k].split(",")
        if j == 3:
                #print (clines[j-3],";",clines[j-1],";",clines[j],";",stock_data[k]['LastTradePrice'],";",stock_data[k]['LastTradeDateTimeLong'])
                #total += round(float(clines[j - 1]) * float(clines[j]), 2)
                #print(total)
                #print(stock_data[k]['LastTradeDateTimeLong'],";",
                #clines[j-3],";",
                #stock_data[k]['LastTradePrice'],";",
                #round(float(stock_data[k]['LastTradePrice'])*float(clines[j-1]),2),";",
                #clines[j-1],";",
                #clines[j],";",
                #round(float(clines[j-1])*float(clines[j]),2),";",
                #round(100 * (float(clines[j-1])*float(clines[j])/float(total_at_cost)),2)
                #)

                #Write some simple text.
                worksheet.write(k+1, 0, stock_data[k]['LastTradeDateTimeLong'])
                worksheet.write(k+1, 1, clines[j - 3])
                worksheet.write(k+1, 2, float(stock_data[k]['LastTradePrice']))
                worksheet.write(k+1, 3, float(clines[j]))
                worksheet.write(k+1, 4, float(stock_data[k]['LastTradePrice'])-float(clines[j]))
                worksheet.write(k+1, 5, round(float(stock_data[k]['LastTradePrice'])*float(clines[j-1]),2))
                worksheet.write(k+1, 6, float(clines[j-1]))
                worksheet.write(k+1, 7, round(float(clines[j-1])*float(clines[j]),2))
                worksheet.write(k+1, 8, round(100 * (float(clines[j-1])*float(clines[j])/float(total_at_cost)),2))
                worksheet.write(k+1, 9, round(float(stock_data[k]['LastTradePrice'])*float(clines[j-1]),2)
                                -round(float(clines[j-1])*float(clines[j]),2))

                # Write a conditional format over a range.
                worksheet.conditional_format(k+1,2,k+1,2, {'type': 'cell',
                                                        'criteria': '>',
                                                        'value': float(clines[j]),
                                                        'format': format2})
                worksheet.conditional_format(k+1,2,k+1,2, {'type': 'cell',
                                                        'criteria': '<=',
                                                        'value': float(clines[j]),
                                                        'format': format1})
                worksheet.conditional_format(k+1,4, k+1,4, {'type': 'cell',
                                                                  'criteria': '<',
                                                                  'value': 0,
                                                                  'format': format1})
                worksheet.conditional_format(k+1,4,k+1,4, {'type': 'cell',
                                                        'criteria': '>=',
                                                        'value': 0,
                                                        'format': format2})
                worksheet.conditional_format(k+1,9, k+1,9, {'type': 'cell',
                                                                  'criteria': '<',
                                                                  'value': 0,
                                                                  'format': format1})
                worksheet.conditional_format(k+1,9,k+1,9, {'type': 'cell',
                                                        'criteria': '>=',
                                                        'value': 0,
                                                        'format': format2})
                total_at_market_value += round(float(clines[j-1]) * float(stock_data[k]['LastTradePrice']), 2)



#Text with formatting.
#worksheet.write('A2', 'World', bold)
worksheet.write(k+2, 0, 'Total Value',bold)
worksheet.write(k+2, 7, total_at_cost)
worksheet.write(k+2, 5, total_at_market_value)
worksheet.write(k+2, 9, total_at_market_value-total_at_cost)
worksheet.conditional_format(k+2,9,k+2,9,{'type': 'cell',
                                                  'criteria': '<',
                                                  'value': 0,
                                                  'format': format1})
worksheet.conditional_format(k+2,9,k+2,9,{'type': 'cell',
                                                  'criteria': '>=',
                                                  'value': 0,
                                                  'format': format2})


workbook.close()