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
