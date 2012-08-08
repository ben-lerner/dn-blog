from ystockquote import get_historical_prices
from datetime import date
import scipy
from scipy.optimize import fmin
from math import sqrt

active_sheet('MVO')
# dates for historical data - ystockquote uses 'yyyymmdd' format
def two_str(x):
    # adds 0 in front of months and dates under 10 for yahoo
    x = str(x)
    if len(x) == 1:
        return '0' + x
    return x

today = datetime.date.today()
start_date = str(today.year - 10) + two_str(today.month) + two_str(today.day)
end_date = str(today.year) + two_str(today.month) + two_str(today.day)

# get list of tickers
row = 2
tickers = []
exp_return = []
r = Cell('C2').value
while not Cell(row, 1).is_empty():
    tickers.append(Cell(row, 1).value)
    exp_return.append(Cell(row, 2).value)
    row += 1

# set up sheet
clear_sheet()
CellRange('A1:E1, G1').value = ['Ticker', 'exp. return', 'risk-free rate', 
                                'sigma', 'optimal weights', 'Correlation Matrix']

L = len(tickers)
CellRange((2,1), (1 + L, 1)).value = tickers
# write expected return and risk free rate as percents
CellRange((2,2), (1 + L, 2)).value = map((lambda x: str(x * 100) + '%'),
                                         exp_return)
Cell('C2').value = str(r * 100) + '%' 
# correlation matrix labels
CellRange((2,7), (1 + L, 7)).value = tickers
CellRange((1,8), (1, 7 + L)).value = tickers

# get data - uses closing price. change f for a different price
data = []
which_price = {'Open': 1, 'High': 2, 'Low': 3, 'Close': 4, 'Adjusted Close': 6}

for ticker in tickers:
    prices = get_historical_prices(ticker, start_date, end_date)
    prices.reverse() # time series order
    prices.pop() # remove header
    f = lambda x: float(x[which_price['Close']])
    prices = map((lambda x: f(x)), prices) # just closing prices
    data.append(prices)

# standardize data size in case some stock history doesn't go back far enough
num_prices = min(map(len, data))
for i, ticker in enumerate(data):
    data[i] = ticker[-num_prices:] # last num_prices data points

# standard deviation
std_dev = []
for i, ticker in enumerate(data):
    dev = float(scipy.std(ticker))/scipy.mean(ticker)
    std_dev.append(dev)
    Cell(2 + i, 4).value = str(round(dev, 4) * 100) + '%'
                
# correlation matrix - for covariance, change scipy.corrcoef to scipy.cov
x = scipy.array(data)
cor = scipy.corrcoef(x)
cor_print = cor.flatten()
cor_print = map((lambda x: round(x,3)), cor_print)
CellRange((2,8),(1 + L, 7 + L)).value = cor_print

# mvo - optimizes Sharpe ratio
def portfolio_variance(a):
    '''Returns the variance of the portfolio with weights a.'''
    var = 0.0
    # to speed up sum covariance over i < j and variance over i
    for i in xrange(L):
        for j in xrange(L):
            var += a[i]*a[j]*std_dev[i]*std_dev[j]*cor[i, j]
    if var <= 0: # floating point errors for very low variance
        return 10**(-6)
    return var

def sharpe_ratio(weights):
    '''Returns the sharpe ratio of the portfolio with weights.'''
    var = portfolio_variance(weights)
    returns = scipy.array(exp_return)
    return (scipy.dot(weights, returns) - r)/sqrt(var)

def sharpe_optimizer(weights):
    # for optimization - computes last weight and returns negative of sharpe
    # ratio (for minimizer to work)
    weights = scipy.append(weights, 1 - sum(weights)) # sum of weights = 1
    return - sharpe_ratio(weights)

guess = scipy.ones(L - 1, dtype=float) * 1.0 / L # L - 1 weights, last asset weight added in optimizer
optimizer = fmin(sharpe_optimizer, guess)
# optimizer docs: http://docs.scipy.org/doc/scipy-0.7.x/reference/generated/scipy.optimize.fmin.html
optimal_weights = scipy.append(optimizer, 1 - sum(optimizer))
optimal_val = sharpe_ratio(optimal_weights)
CellRange((2,5),(1 + L, 5)).value = optimal_weights
Cell(L + 3, 2).value = 'Sharpe ratio:'
Cell(L + 3, 3).value = round(optimal_val, 3)

autofit()
