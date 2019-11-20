# import the CLR library
import clr

# add a reference to the System.Collections Assembly
clr.AddReference('System.Collections')

# generics are usually used with collections, so let's import the 'Dictionary' Generic
from System.Collections.Generic import Dictionary

# let's also import the different data types we need.
from System import *

'''
    EXAMPLE ONE:
    ------------
    For the Dictionary Generic, we can specify what Type we want the Key and Value to be.
    This helps maintain a little more control over the process. In the first example,
    I'm creating a Dictionary Generic that requires the Key to be of Type String and 
    the Value to be of Type String.
'''
generic_dict_1 = Dictionary[String, String]()


'''
    EXAMPLE TWO:
    ------------
    In the second example, I'm creating a Dictionary Generic that requires the Key to 
    be of Type String and the Value to be of Type Int32.
'''
generic_dict_2 = Dictionary[String, Int32]()


'''
    EXAMPLE THREE:
    --------------
    In the third example, I'm creating a Dictionary Generic that requires the Key to 
    be of Type String and the Value to be of Type Type(Object).
'''
generic_dict_3 = Dictionary[String, Type]()


# let's add a value to our Dictionary Generic 2
generic_dict_1.Add("myKey1","myValue1")
generic_dict_1.Add("myKey2","myValue2")

# finally, let's grab a value.
print('For "myKey1", the value is {}'.format(generic_dict_1["myKey1"]))

# let's add a value to our Dictionary Generic 2
generic_dict_2.Add("myKey1", 100)
generic_dict_2.Add("myKey2", 200)

# finally, let's grab a value.
print('For "myKey1", the value is {}'.format(generic_dict_2["myKey1"]))

# here is the neat thing about generics. Even though I'm passing through a string it'll be converted to int.
generic_dict_2.Add("myKey3", "200")

print('For "myKey3", the value is {}'.format(generic_dict_2["myKey3"]))
print('For "myKey3", the type of the value is {}'.format(type(generic_dict_2["myKey3"])))


# import the Environment Type
from System import Environment

# the Envrionment Type has a property called `MachineName`
my_machine_name = Environment.MachineName

# print out my Machine Name
print('My Machine is Called: {}'.format(my_machine_name))


'''
    The Environment Type also has an attribute called `ExitCode` which is just the ExitCode of a process.
''' 


# let's change it to 1 to indicate the process has not completed successfully.
Environment.ExitCode = 1

# let's change it back to 0, to indicate the process has completed successfully.
Environment.ExitCode = 0

# import the Array Type
from System import Array

'''
    Single-Dimension Array.
'''

# declare a single-dimension array, that will only contain numbers, and contains ten elements that are all 0.
array_1 = Array.CreateInstance(int,10)

# print your newly created array.
print("Here is my Array: {}".format(list(array_1)))

# reassign value
array_1[0] = 100

# grab value from element 1, since we are 0-based.
print("Here is the value of my first element: {}".format(array_1[0]))

# here is an alternative way to declaring a single 
array_2 = Array[int](range(10))

# print your newly created array.
print("Here is my Array: {}".format(list(array_2)))


'''
    Multi-Dimension Array.

'''

# declare a multi dimension arrary, that will contain only numbers, and be 2 by 3.
multi_array_1 = Array.CreateInstance(int, 2, 3)

# assign a value to row 0, column 1
multi_array_1[0, 1] = 100

# assign a value to row 0, column 1
multi_array_1[1, 1] = 100

# assign a value to row 0, column 2
multi_array_1[0, 2] = 300

# print your newly created array.
print("Here is my multi-dimensional Array: {}".format(list(multi_array_1)))

# here is an alternative method.
multi_array_2 = Array[Array[int]]( ( (1, 2), (3, 4) ) )

# print your newly created array.
print("Here is my multi-dimensional Array: {}".format([list(array) for array in list(multi_array_2)]))

# grab the current domain, this belongs under the 'System' namespace.
domain = System.AppDomain.CurrentDomain

# get all the assemblies in the domain, this returns a collection that we can loop through.
for assembly in domain.GetAssemblies():
    
    # grab the assembly name, using the GetName() method.
    assembly_name = assembly.GetName()
    
    print('-'*10)
    print(assembly_name)

# key method
print("Using the key method, I got the value: {}".format(generic_dict_2['myKey1']))

# index method
print("Using the index method, I got the value: {}".format(array_1[0]))

# index method, multi.
print("Using the index method multi, I got the value: {}".format(multi_array_1[0,2]))