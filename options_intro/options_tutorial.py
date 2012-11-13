from __future__ import division
from scipy.stats import norm
from math import exp, log, sqrt

# global variables
active_sheet("Options")
cost = Cell("B1").value
sigma = Cell("B2").value
r, q = CellRange("D1:D2").value
prices = range(101) # set of stock prices to value portfolio over. should have length 100 for chart to work

N = norm.cdf

### Black Scholes valuation

def d_one(K, F, t):
    return (log(F/K) + (sigma**2/2)*t)/(sigma*sqrt(t))

def call_val(K, S, t):
    if t == 0: 
        return max(S-K, 0)
    if S == 0:
        return 0
    F = S*exp((r-q)*t)
    d1 = d_one(K, F, t)
    d2 = d1 - sigma*sqrt(t)
    return exp(-r*t)*(F*N(d1) - K*N(d2))

def put_val(K, S, t):
    if t == 0:
        return max(K-S, 0)
    if S == 0:
        return K
    F = S*exp((r-q)*t)
    d1 = d_one(K, F, t)
    d2 = d1 - sigma*sqrt(t)
    return exp(-r*t)*(K*N(-d2) - F*N(-d1))

### general eval code

def opt_val(opt_type, K, S, t):
    if opt_type == "underlying":
        return S
    assert t >= 0
    if opt_type == "call":
        return call_val(K, S, t)
    elif opt_type == "put":
        return put_val(K, S, t)
    else:
        raise TypeError("Invalid option type: %s" %opt_type)

active_sheet("Options")

opt_types = Cell("A5").vertical
opt_types = [x.lower() for x in opt_types] # set opt_types to lower case
num_opts = len(opt_types)
num_calls, num_puts = opt_types.count('call'), opt_types.count('put')

opt_strikes = CellRange((5,2),(4 + num_opts, 2)).value

opt_expiry = CellRange((5,3),(4 + num_opts, 3)).value
opt_expiry = [x if x else 0 for x in opt_expiry] # replace None with 0

opt_qty = Cell("D5").vertical

values = []
def opt_price(price, i):
    # get price of ith option in portfolio at price
    return opt_qty[i] * opt_val(opt_types[i], opt_strikes[i], price, opt_expiry[i]) 

for i in xrange(num_opts):
    # get portfolio value at each price from 0 to 100
    portfolio_val = [opt_price(price, i) for price in prices]
    values.append(portfolio_val)

position = zip(*values) #transpose of values
position = [sum(x) - cost for x in position] # sum position at each price less cost

active_sheet("Data")
clear_sheet()
Cell("A1").value = "Price"
Cell("A2").vertical = prices
Cell("B1").value = "Position"
Cell("B2").vertical = position

call_count, put_count = 1,1
for i, a in enumerate(values):
    # set column title
    opt_title = opt_types[i]
    if opt_title == 'underlying':
        title = 'Underlying'
    if opt_title == 'call':
        if num_calls == 1:
            title = 'Call'
        else:
            title = 'Call %d' %call_count
            call_count += 1
    if opt_title == 'put':
        if num_puts == 1:
            title = 'Put'
        else:
            title = 'Put %d' %put_count
            put_count += 1

    Cell(1, i+3).value = title

    # write data
    Cell(2, i+3).vertical = a

autofit("Data")
