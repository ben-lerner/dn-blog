from ystockquote import get_historical_prices
from datetime import date
import numpy as np
from matplotlib.mlab import PCA

active_sheet('PCA')

# dates for data - ystockquote uses 'yyyymmdd' format
def two_str(x):
    # adds 0 in front of months and dates under 10
    x = str(x)
    if len(x) == 1:
        return '0' + x
    return x

days = Cell('B1').value
end_date = datetime.date.today()
start_date = end_date - datetime.timedelta(days)
start_date = str(start_date.year) + two_str(start_date.month) + two_str(start_date.day)
end_date = str(end_date.year) + two_str(end_date.month) + two_str(end_date.day)

# get list of tickers
sector = Cell(3,2).value
row = 5 # where tickers are written
tickers = []
if sector == 'Custom':
    while not Cell(row, 1).is_empty():
        tickers.append(Cell(row, 1).value)
        row += 1
else:
    CellRange((row,1),(row + 499,1)).clear()
    active_sheet('S&P500')
    for i in xrange(500):
        if sector == 'All' or Cell(2+i, 3).value == sector:
            tickers.append(Cell(2+i, 1).value)
    active_sheet('PCA')
    CellRange((row, 1), (row + len(tickers) - 1, 1)).value = tickers

L = len(tickers)

# clear old data, get cutoff
cutoff = Cell(3,5).value
while not Cell(4, 4).is_empty():
    del_col(4)

# set up sheet
Cell(3,4).value = 'Cutoff:'
Cell(3,4).font.bold = True
Cell(3,5).value = str(100*cutoff) + '%'
if tickers:
    CellRange((5,4), (4 + L, 4)).value = tickers
    Cell(4,4).value = 'Components:'
    Cell(4,4).font.bold = True
    
# get data
which_price = {'Open': 1, 'High': 2, 'Low': 3, 'Close': 4, 'Adjusted Close': 6,
               'Volume': 5}
price_num = which_price[Cell('B2').value]

data = []
percent_data = []
for ticker in tickers:
    prices = get_historical_prices(ticker, start_date, end_date)
    prices.reverse() # time series order
    prices.pop() # remove header
    try:
        prices = map((lambda x: float(x[price_num])), prices) # filter
    except IndexError:
        raise NameError('%s is not a valid ticker' %ticker)
    data.append(prices)
    price_percent = []
    for i in xrange(len(prices) - 1): # change to % for PCA
        price_percent.append((prices[i+1] - prices[i])/prices[i])
    percent_data.append(price_percent)

# PCA
x = np.array(data).T # one column per ticker
pca = PCA(x)
components = pca.Wt
for i, comp in enumerate(components):
    if pca.fracs[i] < cutoff:
        break
    Cell(4, 5 + i).value = str(100*round(pca.fracs[i], 3)) + '%' # proportion of variance
    CellRange((5, 5+i),(4 + L, 5 + i)).value = map((lambda x: round(x, 3)), comp)

autofit()
