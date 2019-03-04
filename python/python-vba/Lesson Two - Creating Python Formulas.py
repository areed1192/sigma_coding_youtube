# import our libraries
import pythoncom
import numpy as np
import win32com.client

class PythonObjectLibrary:

    # This will create a GUID to register it with Windows, it is unique.
    _reg_clsid_ = pythoncom.CreateGuid()

    # Register the object as an EXE file, the alternative is an DLL file (INPROC_SERVER)
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER

    # the program ID, this is the name of the object library that users will use to create the object.
    _reg_progid_ = "Python.ObjectLibrary"

    # this is a description of our object library.
    _reg_desc_ = "This is our Python object library."

    # a list of strings that indicate the public methods for the object. If they aren't listed they are conisdered private.
    _public_methods_ = ['pythonSum', 'pythonMultiply','addArray']

    # multiply two cell values.
    def pythonMultiply(self, a, b):
        return a * b

    # add two cell values
    def pythonSum(self, x, y):
        return x + y

    # add a range of cell values
    def addArray(self, myRange):

        # create an instance of the range object that is passed through
        rng1 = win32com.client.Dispatch(myRange)

        # Get the values from the range
        rng1val = np.array(list(rng1.Value))

        return rng1val.sum()

if __name__ == '__main__':
    import win32com.server.register
    win32com.server.register.UseCommandLine(PythonObjectLibrary)
