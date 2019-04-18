# PRESENT VALUE
def present_value(future_value, discount_rate, periods):
    '''
    Summary: Given a future value cash flow, estimate the present value of that cashflow.
    
    PARA future_value: The future value cash flow.
    PARA type: float
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    RETR type: float
    
    FORMULA: future_value/(1 + discount_rate) ** periods     
    '''

    return future_value / ( 1 + discount_rate) ** periods

# EXAMPLE 
fut_val = 1000.0
rate = 0.2
per = 10

present_value(fv, rate, per)

# FUTURE VALUE
def future_value(present_value, discount_rate, periods):
    '''
    Summary: Given a present value cash flow, estimate the future value of that cashflow.
    
    PARA present_value: The present value cash flow.
    PARA type: float
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    RETR type: float
    
    FORMULA: present_value * (1 + discount_rate) ** periods    
    '''

    return present_value * ( 1 + discount_rate) ** periods

# EXAMPLE 
pres_val = 1000.0
rate = 0.2
per = 10

future_value(pv, rate, per)

# NET PRESENT VALUE
def net_present_value(discount_rate, cashflows):
    '''
    Summary: Given a series of cash flows, calculate the net present value of those cash flows.
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA cashflows: A series of cashflows.
    PARA type: list
    
    RETR type: float
    
    FORMULA: sum(cashflow / (1 + discount_rate) ** periods)     
    '''    
    
    # initalize result
    total_value = 0.0
    
    # loop through cashflows and calculate discounted value.
    for index, cashflow in enumerate(cashflows):
        total_value += cashflow / (1 + discount_rate)**index
        
    return total_value 

# EXAMPLE
rate = 0.2
cashflows = [100, 100, 100, 100, 100]

net_present_value(rate, cashflows)

# present value perpetuity
def present_value_perpetuity(cashflow, discount_rate):
    '''
    Summary: Given a cash flow, calculate the present value in perpetuity.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    RETR type: float
    
    FORMULA: cashflow / discount rate    
    '''       
    
    return cashflow / discount_rate

# EXAMPLE
cf = 1000.0
rate = 0.01
present_value_perpetuity(cf, rate)

# present value perpetuity due
def present_value_perpetuity_due(cashflow, discount_rate):
    '''
    Summary: Given a cash flow, calculate the present value in perpetuity due.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    RETR type: float
    
    FORMULA: cashflow / discount_rate * (1 + discount_rate)
    '''       
    
    return cashflow / discount_rate * (1 + discount_rate)

# EXAMPLE
cf = 1000.0
rate = 0.01
present_value_perpetuity_due(cf, rate)

# present value annuity
def present_value_annuity(cashflow, discount_rate, periods):
    '''
    Summary: Given a cash flow, calculate the present value of an annuity.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    RETR type: float
    
    FORMULA: cashflow / discount_rate * (1 - 1 / ( 1 + discount_rate )** periods)
    '''       
    return cashflow / discount_rate * (1 - 1 / ( 1 + discount_rate )** periods)

# EXAMPLE
cf = 1000.0
rate = 0.01
per = 10
present_value_annuity(cf, rate, per)

# present value annuity due
def present_value_annuity_due(cashflow, discount_rate, periods):
    '''
    Summary: Given a cash flow, calculate the present value of an annuity due.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    RETR type: float
    
    FORMULA: cashflow / discount_rate * (1 - 1 / ( 1 + discount_rate )** periods) * (1 + discount_rate)
    '''   
    
    return cashflow / discount_rate * (1 - 1 / ( 1 + discount_rate )** periods) * (1 + discount_rate)

# EXAMPLE
cf = 1000.0
rate = 0.01
per = 10
present_value_annuity_due(cf, rate, per)

# present value growing annuity
def present_value_growing_annuity(cashflow, discount_rate, periods, growth_rate):
    '''
    Summary: Given a cash flow, calculate the present value of a growing annuity. It is assumed that the discount rate > growth
             rate, otherwise a negative value will be returned.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    PARA growth_rate: The growth rate
    PARA type: float
    
    RETR type: float
    
    FORMULA: cashflow / discount_rate * (1 - (1 + growth_rate) ** periods / (1 + discount_rate) ** periods)
    '''   
    
    return cashflow / discount_rate * (1 - (1 + growth_rate) ** periods / (1 + discount_rate) ** periods)

# EXAMPLE
cf = 1000.0
rate = 0.05
per = 10
gr = 0.03
                                       
present_value_growing_annuity(cf, rate, per, gr)

# future value annuity
def future_value_annuity(cashflow, discount_rate, periods):
    '''
    Summary: Given a cash flow, calculate the future value of an annuity.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    RETR type: float
    
    FORMULA: cashflow / discount_rate * (( 1 + discount_rate )** periods - 1)
    '''  
    return cashflow / discount_rate * (( 1 + discount_rate )** periods - 1)

# EXAMPLE
cf = 1000.0
rate = 0.01
per = 10
future_value_annuity(cf, rate, per)

# future value annuity due
def future_value_annuity_due(cashflow, discount_rate, periods):
    '''
    Summary: Given a cash flow, calculate the future value of an annuity due.
     
    PARA cashflow: A single cashflow.
    PARA type: floar
    
    PARA discount_rate: The discount rate
    PARA type: float
    
    PARA periods: The number of periods.
    PARA type: int
    
    RETR type: float
    
    FORMULA: cashflow / discount_rate * (( 1 + discount_rate )** periods - 1) * (1 + discount_rate)
    '''   
    
    return cashflow / discount_rate * (( 1 + discount_rate )** periods - 1) * (1 + discount_rate)

# EXAMPLE
cf = 1000.0
rate = 0.01
per = 10
future_value_annuity_due(cf, rate, per)

# effective annual rate
def eff_annual_rate(apr, frequency):
    '''
    Summary: Given an APR (Annual Perecentage Rate) calculate the EAR (Effective Annual Rate)
     
    PARA apr: The Annual Percentage Rate.
    PARA type: float
    
    PARA frequency: The compounding frequency.
    PARA type: int
    '''
    
    return (1 + apr / frequency) ** frequency - 1

# EXAMPLE
annual_rate = .03
comp_freq = 10
eff_annual_rate(annual_rate, comp_freq)  

def stock_valuation_n_period(discount_rate ,LT_growth_rate, dividends):
    """
    Summary: Given an array of dividends, estimate the stock price using an n-period model.
    
    PARA discount_rate: The discount rate used to calculate the NPV & PV calcs.
    PARA type: float
    
    PARA LT_growth_rate: The long-term growth rate used to calculate the last dividend in perpituity.
    PARA type: float
    
    PARA dividends: a list of dividends where the last dividend is the one that is earned in perpituity.
    PARA type: float
    
    """
    
    # get the dividend array & the last dividend
    div_array = dividends[:-1]
    div_last = dividends[-1]
    
    # define the number of periods.
    num_pers = len(dividends) - 1

    # calculate the present value of the dividends.
    pres_val = net_present_value(discount_rate, div_array) * (1 + discount_rate)

    # calculate late the terminal value, which is a dividend in perpituity.
    last_div_n = div_last / (discount_rate - LT_growth_rate)    

    #calulate the the total value which is the pv of the dividends cashflow and the present value of the dividend in perpituity.
    total_val = pres_val + present_value(last_div_n, discount_rate, num_pers)
    
    return total_val
    

#EXAMPLE
disc_rate = 0.12
grow_rate = 0.03
dividends = [1.8, 2.07, 2.277, 2.48193, 2.68, 2.7877]

stock_valuation_n_period(disc_rate, grow_rate, dividends)

def divident_discount_model(discount_rate, dividends, growth_rate, stock_price, periods):
    """
    Summary: Given an array of dividends, estimate the stock price using an n-period model.
    
    PARA discount_rate: The discount rate used to calculate the NPV & PV calcs.
    PARA type: float

    PARA dividends: The period 0 dividend.
    PARA type: float
    
    PARA growth_rate: The growth rate that will be applied to the dividend every period.
    PARA type: float
    
    PARA stock_price: The stock price at period n
    PARA type: float
    
    PARA periods: The number of periods that will calculated out to.
    PARA type: int    
    """
    
    # initalize our total cashflows
    total_cashflows = 0
    
    # loop the number of periods and calculate the cash flows.
    for period in range(1, periods + 1):
        
        # define the growth and discount factor
        growth_factor = (1 + growth_rate)
        discount_factor = (1 + discount_rate) ** period
        
        # calculate the cashflow
        cashflow = (dividends * growth_factor) ** period / discount_factor
        total_cashflows = total_cashflows + cashflow
    
    # calculate the terminal value, or the stock price at period n.
    terminal_val =  stock_price / (1 + discount_rate) ** periods 

    return total_cashflows + terminal_val
    

#EXAMPLE
disc_rate = 0.132
grow_rate = 0.05
dividends = 1.0
stock_price = 14.12
periods = 5

divident_discount_model(disc_rate, dividends, grow_rate, stock_price, periods)

def gordon_growth_model(dividend, dividend_growth_rate, required_rate_of_return):
    """
    Summary: Calculate the value of a stock using a gordon growth model.
    
    PARA dividend: The dividend earned over the life of the stock.
    PARA type: float

    PARA dividend_growth_rate: The growth rate in the value of the dividend every period.
    PARA type: float
    
    PARA required_rate_of_return: The required rate of return for the investor.
    PARA type: float
   
    """  
    
    dividend_period_one = dividend * (1 + dividend_growth_rate)
    
    return dividend_period_one / (required_rate_of_return - dividend_growth_rate)

div = 2.00
gro = 0.05
rrr = 0.12

gordon_growth_model(div, gro, rrr)

def multistage_growth_model(dividend, discount_rate, growth_rate, constant_growth_rate, periods):
    """
    Summary: Calculate the value of a stock using a multistage growth model.
    
    PARA dividend: The dividend earned over the life of the stock.
    PARA type: float

    PARA discount_rate: The discount rate used to calculate the NPV & PV calcs.
    PARA type: float
    
    PARA growth_rate: The growth rate during the multistage period.
    PARA type: float
    
    PARA constant_growth: The growth rate in perpituity.
    PARA type: float
    
    PARA periods: The number of periods to be calculated.
    PARA type: int
   
    """  
    total_value= 0
  
    
    for period in range(1, periods + 1):
        
        # if it's the last period calculate the terminal value
        if period == periods:
            
            # calculate the terminal dividend.
            terminal_dividend = (dividend * (1 + growth_rate) ** period)
            
            # calculate the terminal value and then discount it.
            terminal_value = terminal_dividend / (discount_rate - constant_growth_rate)
            terminal_value_disc = terminal_value / (1 + discount_rate) ** (period -1)
            
            # return the total value of the stock
            total_value += terminal_value_disc
        
        # otherwise calculate the cashflow for that period
        else:
            cashflow = (dividend * (1 + growth_rate) ** period) / (1 + discount_rate) ** period
            total_value += cashflow
            
    return total_value
           
# EXAMPLE  
div = 1.00
gro = 0.20
cos = 0.05
dis = 0.10
per = 4

multistage_growth_model(div, dis, gro, cos, per)

def preferred_stock_valuation(dividend, required_rate_of_return):
    """
    Summary: Given a preferred dividend stock, calculate the value of that stock.
    
    PARA dividend: The dividend for each period, earned over an infinite period.
    PARA type: float

    PARA required_rate_of_return: The required rate of return desired for an investment.
    PARA type: float
    
    """  
    
    return dividend / required_rate_of_return

# EXAMPLE
annual_div = 5.00
rrr = .08

preferred_stock_valuation(annual_div, rrr)

def sharpe_ratio(returns, risk_free_rate):
    """
    Summary: Given an array of returns and a risk-free rate, calculate the Sharpe Ratio.
    
    PARA returns: A list of returns, for example daily stock returns.
    PARA type: list
    
    PARA risk_free_rate: The risk-free rate, usually a t-bill.
    PARA type: float
    
    """
    
    import numpy as np
    
    returns = np.array(returns)    
    
    # calculate the avg return and the std return.
    avg_return = returns.mean()
    std_return = returns.std()
    
    return (avg_return - risk_free_rate) / std_return
 
# EXAMPLE
risk_free_rate = .02
returns = [1, 3, 4, 5, 6]
sharpe_ratio(returns, risk_free_rate)

def cost_of_preferred_stock(preferred_dividends, market_price_of_preferred):
    """
    Summary: Calculate the cost of preferred stock that is used in a WACC calculation.
    
    PARA preferred_dividends: The amount of a preferred dividend paid in that period.
    PARA type: float
    
    PARA market_price_of preferred: The price of a share of preferred stock during the period.
    PARA type: float
    
    """    
    
    return preferred_dividends / market_price_of_preferred


# EXAMPLE
dividend = 8
price = 100
cost_of_preferred_stock( dividend, price)

def cost_of_debt(interest_rate, tax_rate):
    """
    Summary: Calculate the cost of debt that is used in a WACC calculation.
    
    PARA interest_rate: The interest rate charged on the debt.
    PARA type: float
    
    PARA tax_rate: The company's marginal federal plus state tax rate..
    PARA type: float
    
    """       
    
    return interest_rate * (1 - tax_rate)

# EXAMPLE
interest_rate = .08
tax_rate = .4
cost_of_debt(interest_rate, tax_rate)

def cost_of_equity_capm(risk_free_rate, market_return, beta):
    """
    Summary: Calculate the cost of equity for WACC using the CAPM method.
    
    PARA risk_free_rate: The risk free rate for the market, usually a t-note.
    PARA type: float
    
    PARA market_return: The required rate of return for the company.
    PARA type: float

    PARA beta: The company's estimated stock beta.
    PARA type: float
    
    """ 
    
    return risk_free_rate + (beta * (market_return - risk_free_rate))

# EXAMPLE
beta = 1.1
rfr = .06
mkt = .11
cost_of_equity_capm(rfr, mkt, beta)

def cost_of_equity_ddm(stock_price, next_year_dividend, growth_rate):
    """
    Summary: Calculate the cost of equity for WACC using the DMM method.
    
    PARA stock_price: The company's current price of a share.
    PARA type: float
    
    PARA next_year_dividend: The expected dividend to be paid next year.
    PARA type: float

    PARA growth_rate: Firm's expected constant growth rate.
    PARA type: float
    
    """    
    return (next_year_dividend / stock_price) + growth_rate

# EXAMPLE
p = 21
d = 1
g = .072
cost_of_equity_ddm(p, d, g)

def cost_of_equity_bond(bond_yield, risk_premium):
    """
    Summary: Calculate the cost of equity for WACC using the Bond yield plus risk premium method.
    
    PARA bond_yield: The company's interest rate on long-term debt.
    PARA type: float
    
    PARA risk_premium: The company's risk premium usually 3% to 5%.
    PARA type: float
    
    """   
    return bond_yield + risk_premium

# EXAMPLE
y = .08
p = .05
cost_of_equity_bond(y, p)

def capital_weights(preferred_stock, total_debt, common_stock):
    """
    Summary: Given a firm's capital structure, calculate the weights of each group.
    
    PARA total_capital: The company's total capital.
    PARA type: float
    
    PARA preferred_stock: The comapny's preferred stock outstanding.
    PARA type: float
    
    PARA common_stock: The comapny's common stock outstanding.
    PARA type: float
    
    PARA total_debt: The comapny's total debt.
    PARA type: float
    
    RTYP weights_dict: A dictionary of all the weights. 
    RTYP weights_dict: dictionary
    
    """      
    # initalize the dictionary
    weights_dict = {}
    
    # calculate the total capital
    total_capital = preferred_stock + common_stock + total_debt
    
    # calculate each weight and store it in the dictionary
    weights_dict['preferred_stock'] = preferred_stock / total_capital
    weights_dict['common_stock'] = common_stock / total_capital
    weights_dict['total_debt'] = total_debt / total_capital
    
    return weights_dict

debt = 8000000
preferred_stock = 2000000
common_stock = 10000000
capital_weights(preferred_stock, debt, common_stock)

def weighted_average_cost_of_capital(cost_of_common, cost_of_debt, cost_of_preferred, weights_dict):
    """
    Summary: Calculate a firm's wACC.
    
    PARA cost_of_common: The firm's cost of common equity.
    PARA type: float
    
    PARA cost_of_debt: The firm's cost of debt.
    PARA type: float
    
    PARA cost_of_preferred: The firm's cost of preferred equity.
    PARA type: float
    
    PARA weights_dict: The capital weights for each capital structure.
    PARA type: dictionary
    
    """    
    
    weight_debt = weights_dict['total_debt']
    weight_common = weights_dict['common_stock']
    weight_preferred = weights_dict['preferred_stock']
    
    return (weight_debt * cost_of_debt) + (weight_common * cost_of_common) + (weight_preferred * cost_of_preferred)

# Cost of Equity
y = .08
p = .05
ke = cost_of_equity_bond(y, p)

# Cost of Debt
interest_rate = .08
tax_rate = .4
kd = cost_of_debt(interest_rate, tax_rate)

# Cost of Preferred
dividend = 8
price = 100
kp = cost_of_preferred_stock( dividend, price)

# Capital Weights
debt = 8000000
preferred_stock = 2000000
common_stock = 10000000
weights = capital_weights(preferred_stock, debt, common_stock)

weighted_average_cost_of_capital(ke, kd, kp, weights)

def asset_beta(tax_rate, equity_beta, debt_to_equity):
    """
    Summary: Calculate the asset beta for a publicly traded firm.
    
    PARA tax_rate: A comparable publicly traded company's marginal tax-rate.
    PARA type: float
    
    PARA equity_beta: A comparable publicly traded company's equity beta.
    PARA type: float
    
    PARA debt_to_equity: A comparable publicly traded company's debt-to-equity ratio.
    PARA type: float
        
    """     
    
    return equity_beta * 1 / (1 + ((1 - tax_rate) * debt_to_equity))

#EXAMPLE
tax_rate = .3
equity_beta = .9
dte = 1.5   
                          
asset_beta(tax_rate, equity_beta, dte)                

def project_beta(tax_rate, asset_beta, debt_to_equity):
    """
    Summary: Calculate the project beta for subject firm.
    
    PARA tax_rate: Company's marginal tax-rate.
    PARA type: float
    
    PARA asset_beta: A comparable publicly traded company's asset beta.
    PARA type: float
    
    PARA debt_to_equity: Company's debt-to-equity ratio.
    PARA type: float
        
    """     
    
    return asset_beta * (1 + ((1 - tax_rate)*debt_to_equity))

#Example
a = asset_beta(.3, .9, 1.5)   
t = .4
d = 2
project_beta(t, a,d )

def profitability_index(net_present_value, initial_cash_outlay):
    
    return 1 + (net_present_value / initial_cash_outlay)

def degree_of_operating_leverage(quantity, variable_cost, price, fixed_cost):
    """
    Summary: Calculate the degree of operating leverage.
    
    PARA quantity: Quantity of units sold.
    PARA type: int
    
    PARA variable_cost: The variable cost per unit.
    PARA type: float
    
    PARA fixed_cost: Total fixed costs
    PARA type: float
    
    PARA price: The price per unit.
    PARA type: float
        
    """        
    
    return (quantity * (price - variable_cost)) /  (quantity * (price - variable_cost) - fixed_cost)

price = 4.00
variable = 3.00
quantity = 100000
fixed = 40000

degree_of_operating_leverage(quantity, variable, price, fixed) 

def degree_of_financial_leverage(ebit, interest):
    """
    Summary: Calculate the degree of operating leverage.
    
    PARA ebit: Earnings before interst & taxes.
    PARA type: int
    
    PARA interest: Interest expense per period.
    PARA type: float
        
    """  
    
    return ebit / (ebit - interest)
    
ebit = 60000
interest = 18000

degree_of_financial_leverage(ebit, interest)

def degree_of_total_leverage(financial_leverage, operating_leverage):
    """
    Summary: Calculate the degree of total leverage.
    
    PARA financial_leverage: Firms degree of financial leverage.
    PARA type: float
    
    PARA operating_leverage: Firms degree of operating leverage.
    PARA type: float
        
    """  
    
    return financial_leverage * operating_leverage

fin_lev = degree_of_financial_leverage(60000, 18000)
opr_lev = degree_of_operating_leverage(100000, 3.0, 4.0, 40000) 
degree_of_total_leverage(fin_lev, opr_lev)

