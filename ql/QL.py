from QuantLib import *
from ystockquote import get_price
import datetime

def stock_price(ticker):
    return get_price(ticker)

def make_date(day):
    """ returns a QL Date object from a datetime"""
    return Date(day.day, day.month, day.year)

def opt_price(settlement, maturity, underlying, strike, dividendYield,
              riskFreeRate, volatility):
    # dates
    calendar=TARGET()
    todaysDate = make_date(datetime.datetime.today())
    settlementDate = make_date(settlement)
    maturity = make_date(maturity)
    Settings.instance().evaluationDate = todaysDate
    dayCounter = Actual365Fixed()

    #option parameters
    if Cell('D1').value == 'Call':
        opt_type = Option.Call
    else:
        opt_type = Option.Put

    # basic option
    payoff = PlainVanillaPayoff(opt_type, strike)
    exercise = EuropeanExercise(maturity)
    europeanOption = VanillaOption(payoff, exercise)

    # handle setups
    underlyingH = QuoteHandle(SimpleQuote(underlying))
    flatTermStructure = YieldTermStructureHandle(FlatForward(settlementDate, riskFreeRate, dayCounter))
    dividendYield = YieldTermStructureHandle(FlatForward(settlementDate, dividendYield, Actual365Fixed()))
    flatVolTS = BlackVolTermStructureHandle(BlackConstantVol(settlementDate,
                                                             calendar, volatility,
                                                             dayCounter))

    # done
    bsmProcess = BlackScholesMertonProcess(underlyingH, dividendYield, flatTermStructure, flatVolTS)

    # method: analytic
    europeanOption.setPricingEngine(AnalyticEuropeanEngine(bsmProcess))
    value = europeanOption.NPV()
    return value

### main
active_sheet('main')
Cell('B2').value = get_price(Cell('B1').value)
settlement = Cell('D3').date
maturity = Cell('D4').date
S, K, vol, d, r = CellRange('B2:B5,D5').value
Cell('B7').value = opt_price(settlement, maturity, S, K, d, r, vol)
