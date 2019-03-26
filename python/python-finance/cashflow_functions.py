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
