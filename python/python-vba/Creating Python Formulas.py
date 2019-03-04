import pythoncom
import win32com.client
import numpy as np

class HelloWorld:

    _reg_clsid_ = pythoncom.CreateGuid()
    _reg_desc_ = "Python Test COM Server"
    _reg_progid_ = "Python.TestServer"
    _public_methods_ = ['pythonSum', 'pythonMultiply', 'helloWorld', 'ReturnType','multiplyArray']

    def pythonSum(self,x,y):
        return x + y

    def pythonMultiply(self,a,b):
        return a*b

    def helloWorld(self,a):
        return "Hi there {}".format(a)

    def ReturnType(self,a):
        myType = str(type(a))
        return myType

    def multiplyArray(self, rangeOne, rangeTwo):        
        
        # create the instances of the ranges
        rng1 = win32com.client.gencache.EnsureDispatch(rangeOne)
        rng2 = win32com.client.gencache.EnsureDispatch(rangeTwo)

        # Get the values from the ranges
        rng1Val = np.array(list(rng1.Value))
        rng2Val = np.array(list(rng2.Value))

        # multiply the values
        x = rng1Val * rng2Val
  
        return x.sum()
  
if __name__=='__main__':
    print ("Registering COM server...")
    import win32com.server.register
    win32com.server.register.UseCommandLine(HelloWorld)
