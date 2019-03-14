# -*- coding: utf-8 -*-
"""
Created on Thu Dec 20 20:53:40 2018

@author: Philip Young

ctfs/2018-08-26-MasterCard-GasAdvantage.pdf
"""

from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTChar, LTRect, LTCurve, LTImage
import logging, csv, glob
from datetime import datetime
from decimal import Decimal, getcontext

# Stores the current year, and uses the statement period
# to convert "Month Day" strings into datetime objects
class datemaker(object):
    year = 1900
    split = False

    # initialize the class - change year from 1900 to year specified in period
    def init(period):
        try:
            i = period.find(',') + 1
            i = period.find(',', i) + 2
            datemaker.year = int(period[i:i+4])
            print(period[:i+4])
            if period[0] == 'D':
                datemaker.split = True
        except:
            print("datemaker.init(period) failed! period: " + period)
            pass
    
    # convert "Month Day" string to "Month Day Year" datetime object
    def makedate(date):
        year = datemaker.year
        if date[0] == 'D' and datemaker.split:
            year -= 1
        year = " " + str(year)
        return datetime.strptime(date + year, "%b %d %Y")

# starting balance to work back from
class balance(object):
    bal = Decimal('0.00')
    
    def newbalance(newbal):
        balance.bal = Decimal(newbal)
        
    def getbal():
        return balance.bal
    
# Object to store 
class transaction(object):
    def __init__(self, data):
        # credit transaction
        if len(data) == 3:
            self.date = datemaker.makedate(data[0])
            self.locn = data[1]
            self.type = "credit"
            amt = Decimal(data[2].replace(",", ""))
            if amt > Decimal('0'):
                self.fout = amt
                self.fin = Decimal('0.0')
            else:
                self.fout = Decimal('0.0')
                self.fin = Decimal.copy_abs(amt)
            #print("CREATING: {}, AMT: {}, FIN: {}, FOUT: {}".format(data[1], data[2], self.fin, self.fout))
        # debit transaction
        elif len(data) == 4:
            self.date = datetime.strptime(data[0], "%m/%d/%Y")
            self.locn = data[1]
            self.type = "debit"
            if data[2] != '':
                data[2] = Decimal(data[2])
            else:
                data[2] = Decimal('0.0')
            
            self.fout = data[2]
            if data[3] != '':
                data[3] = Decimal(data[3])
            else:
                data[3] = Decimal('0.0')
            self.fin = data[3]
            
    def __lt__(self, other):
        if self.date == other.date:
            return self.fin > other.fin
        else:
            return self.date < other.date
     
    def getList(self):
        return [self.date.strftime("%Y-%m-%d"), self.locn, self.type, "{0:.02f}".format(self.fin), "{0:.02f}".format(self.fout)]
     
    def myprint(self):
        print(self.getList())

# Takes in a string, scraped from the pdf, and produces a list of transaction objects
def get_transactions(strin):
    months = set(['Jan ', 'Feb ', 'Mar ', 'Apr ', 'May ', 'Jun ', 'Jul ', 'Aug ', 'Sep ', 'Oct ', 'Nov ', 'Dec '])
    lotransactions = []
    i = 0
    
    # set the year for this pdf
    try:
        period = strin.index('For the period') + 16
        datemaker.init(strin[period:period + 40])
    except:
        pass
    
    # set the balance for this month
    try:
        balstart = strin.index('Your New Balance') + 17
        i = balstart
        while strin[i].isdigit() or strin[i] == ',' or strin[i] == '.' or strin[i] == '-':
            i += 1
        balance.newbalance(strin[balstart:i])
        print("Balance: {}".format(balance.getbal()))
    except:
        pass
    
    # set index i to the start of the transactions section
    try:
        i = strin.index('($)') + 3
    except:
        return []
         
    # transactions found, begin parsing
    while (i < len(strin)):
        if (strin[i:i+4] in months):
            # date transaction was happened
            date = strin[i:i+6]
            i += 12
            # location of transaction
            start = i
            while (i < len(strin)):
                # found dollar amount "[0-9].[0-9]"
                if (strin[i] == '.') and (strin[i-1].isdigit()) and (strin[i+1].isdigit()):
                    locend = i-1                 # locend is the end of the location substring
                    while (strin[locend].isdigit() or strin[locend] == ','):
                        locend -= 1             # decriment locend until a non-digit is found
                    if (strin[locend] != '-'):  # check for negative sign on dollar amount
                        locend += 1
                    i += 2
                    break
                else:
                    i += 1
            # location substring
            locn = strin[start:locend]
            # amount substring
            amt = strin[locend:i+1]
            lotransactions.append(transaction([date, locn, amt]))
        elif (strin[i:i+24] == "in the Purchases section"):
            break
        i += 1
    
    return lotransactions

# Scrapes chars from the pdf and produces a long string
def scrape_chars(layout):
    retval = []
    mystring = ''
    for lt_obj in layout:
        #print(lt_obj.__class__.__name__)
        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
            mystring += lt_obj.get_text()   # build string from textbox or textline
        elif isinstance(lt_obj, LTChar):
            mystring += lt_obj.get_text()   # build string from individual chars
        elif isinstance(lt_obj, LTFigure):
            retval = scrape_chars(lt_obj)   # Recursive
        elif isinstance(lt_obj, LTRect) or isinstance(lt_obj, LTImage) or isinstance(lt_obj, LTCurve):
            continue
        else:
            # This should never print
            print("I'M NOT A GOOD THINGY IDK WHAT I AM!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            print(lt_obj.__class__.__name__)
    # only analyse layout elements with name I1 or I2 (they hold the transactions)
    if (len(mystring) > 0):
        retval = get_transactions(mystring)
    # returns a list of transaction classes (empty list if no transactions found)
    return retval

def main():
    # supress pdfminer warnings
    logging.propagate = False 
    logging.getLogger().setLevel(logging.ERROR)
    
    lot = []    # list of transactions
    
    # parse ctfs transactions
    for file in glob.glob("ctfs/*.pdf"):
        #file = "ctfs/2017-12-26-MasterCard-GasAdvantage.pdf"
        fp = open(file, 'rb')
        parser = PDFParser(fp)
        doc = PDFDocument()
        parser.set_document(doc)
        doc.set_parser(parser)
        doc.initialize('')
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        laparams.char_margin = 1.0
        laparams.word_margin = 1.0
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        
        getcontext().prec = 28
        
        accumulator = Decimal('0.00')
        
        # for pages in pdf
        for page in doc.get_pages():
            interpreter.process_page(page)
            layout = device.get_result()
            testing = scrape_chars(layout)
            lot += testing
            
            for item in testing:
                accumulator += item.fout
    
        print("Monthly Spend: {0:.2f}\n".format(accumulator))
    
    
    # parse Simplii transactions
    banktrans = csv.reader(open('simplii/SIMPLII.csv'))
    itertrans = iter(banktrans)     # convert to iterable object
    next(itertrans)                 # skip first item (column headers)
    for row in itertrans:
        lot.append(transaction(row))
    
    lot.sort(reverse = True)
    
    
    # tally balances backwards
    credbal = balance.getbal()
    debitbal = Decimal('0.00')
    total = debitbal - credbal
    
    # open csv writer to send plot data out to excel
    csv_file = open('output.csv', mode='w', newline='')
    csv_out = csv.writer(csv_file)
    
    # give csv column titles
    csv_out.writerow(['Date', 'Location', 'Account', 'Funds-In', 'Funds-Out', 'Balance'])
    for l in lot:
        csv_out.writerow(l.getList() + ["{0:.02f}".format(credbal)])
        if l.type == "credit":
            credbal -= l.fout
            credbal += l.fin
        elif l.type == "debit":
            debitbal += l.fout
            debitbal -= l.fin
        total = debitbal - credbal
    
    # close csv writer
    csv_file.close()
    """
    # Data to plot
    labels = 'Python', 'C++', 'Ruby', 'Java'
    sizes = [215, 130, 245, 210]
    colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
    explode = (0.1, 0, 0, 0)  # explode 1st slice
     
    # Plot
    plt.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
     
    plt.axis('equal')
    plt.show()
    """
if __name__ == '__main__': main()