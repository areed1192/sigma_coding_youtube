
import clr

# add a reference for both the System Assembly, and System.Reflection Assembly.
clr.AddReference("System")
clr.AddReference("System.Reflection")

# import them both.
import System.Reflection
import System

# define the file path to the Excel Interop Assembly.
file_path = r"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.Excel.dll"

# using the System.Reflection Assembly we can load it into our environment.
excel_interop_lib = System.Reflection.Assembly.LoadFile(file_path)

# let's print some info about our library like it's name, it's location, and the Type of object it is.
print('\n')
print('MAIN INTEROP ASSEMBLY')
print('-'*20)
print("This Assembly is Called: {}".format(excel_interop_lib.FullName))
print("The Location of this Assembly is: {}".format(excel_interop_lib.Location))
print("The Library is of Type: {}".format(excel_interop_lib.GetType()))
print("\n")

# some assemblies reference other assemplies, lets see all the assemblies referenced in our Excel Interop Assembly. 
for assembly in excel_interop_lib.GetReferencedAssemblies():

    print("REFERENCED ASSEMBLIES:")
    print("-"*20)
    print("\tThis Assembly is Called: {}".format(assembly.FullName))
    print("\tThe Library is of Type: {}".format(assembly.GetType()))
    print("\tThe Library Version is: {}".format(assembly.Version))
    print("\tThe Library Code Base is: {}".format(assembly.CodeBase))
    print("\n")

# initalizes some lists to store the different types
interfaces = []
classes = []
enumerations = []
structures = []
unknown = []

# an assembly can contain a collection of Type Objects. These objects contain information about the object for example it's methods
# let's used the `GetTypes()` method to grab all the Type Objects in the Excel Interop Assembly
for excel_type in excel_interop_lib.GetTypes():

    # type comes in many different forms, for example we have interfaces, Classes and many more.
    # depending on the type we encounter we can extract additional information about it.

    # check if it is an Interface
    if excel_type.IsInterface == True:
        interfaces.append(excel_type)

    # check if it is a Class
    elif excel_type.IsClass == True:
        classes.append(excel_type)
    
    # check if it is a Enumeration
    elif excel_type.IsEnum == True:
        enumerations.append(excel_type)

    # check if it is a Structure
    elif excel_type.IsValueType and not excel_type.IsEnum:
        structures.append(excel_type)

    # if we missed something store it.
    else:
        unknown.append(excel_type)

# display a summary
print('\n')
print('There are {} Intrefaces.'.format(len(interfaces)))
print('There are {} Classes.'.format(len(classes)))
print('There are {} Enumerations.'.format(len(enumerations)))
print('There are {} Structures.'.format(len(structures)))
print('\n')


print("CLASSES:")
print("-"*20)

# let's start with classes and see what we can get from them.
for _class in classes:

    # if _class.IsCOMObject == True:

    print('{} CLASS'.format(_class.Name))
    print('-'*20)
    print('\tIs the class Visible: {}'.format(_class.IsVisible))
    print('\tIs the class Sealed: {}'.format(_class.IsSealed))
    print('\tIs the class a COM Object: {}'.format(_class.IsCOMObject))
    print('\tIs the class an Abstract Class: {}'.format(_class.IsAbstract))
    print('\tThe class is in Module: {}'.format(_class.Module))
    print('\tThe class is in Namespace: {}'.format(_class.Namespace))
    print('\tThe class has a Full Name of: {}'.format(_class.FullName))
    print('\tThe class has an Underlying System Type of: {}'.format(_class.UnderlyingSystemType))
    print('\tThe class has {} Members.'.format(len(_class.GetMembers())))
    print('\tThe class has {} Methods.'.format(len(_class.GetMethods())))
    print('\tThe class has {} Properties.'.format(len(_class.GetProperties())))
    print('\tThe class has {} Events.'.format(len(_class.GetEvents())))
    print('\tThe class has {} Interfaces.'.format(len(_class.GetInterfaces())))
    print('\tThe class has {} Fields.'.format(len(_class.GetFields())))
    print('\tThe class has {} Constructors.'.format(len(_class.GetConstructors())))
    print('\tThe class has {} Nested Types.'.format(len(_class.GetNestedTypes())))
    print('\n')


print("INTERFACES:")
print("-"*20)

for _interface in interfaces:

    print('{} INTERFACE'.format(_interface.Name))
    print('-'*20)
    print('\tIs the interface Visible: {}'.format(_interface.IsVisible))
    print('\tIs the interface Sealed: {}'.format(_interface.IsSealed))
    print('\tIs the interface a COM Object: {}'.format(_interface.IsCOMObject))
    print('\tIs the interface an Abstract interface: {}'.format(_interface.IsAbstract))
    print('\tThe interface is in Module: {}'.format(_interface.Module))
    print('\tThe interface is in Namespace: {}'.format(_interface.Namespace))
    print('\tThe interface has a Full Name of: {}'.format(_interface.FullName))
    print('\tThe interface has an Underlying System Type of: {}'.format(_interface.UnderlyingSystemType))
    print('\tThe interface has {} Members.'.format(len(_interface.GetMembers())))
    print('\tThe interface has {} Methods.'.format(len(_interface.GetMethods())))
    print('\tThe interface has {} Properties.'.format(len(_interface.GetProperties())))
    print('\tThe interface has {} Events.'.format(len(_interface.GetEvents())))
    print('\tThe interface has {} Interfaces.'.format(len(_interface.GetInterfaces())))
    print('\tThe interface has {} Fields.'.format(len(_interface.GetFields())))
    print('\tThe interface has {} Constructors.'.format(len(_interface.GetConstructors())))
    print('\tThe interface has {} Nested Types.'.format(len(_interface.GetNestedTypes())))
    print('\n')

print("ENUMERATION:")
print("-"*20)

for _enum in enumerations:

    print('{} ENUMERATION'.format(_enum.Name))
    print('-'*20)
    print('\tThe enum is in Module: {}'.format(_enum.Module))
    print('\tThe enum is in Namespace: {}'.format(_enum.Namespace))
    print('\tThe enum has a Full Name of: {}'.format(_enum.FullName))
    print('\tThe enum has an Underlying System Type of: {}'.format(_enum.UnderlyingSystemType))
    print('\tThe Number of enum Names: {}'.format(len(list(_enum.GetEnumNames()))))
    print('\tThe Number of enum Values: {}'.format(len(list(_enum.GetEnumValues()))))
    print('\tThe enum Values: {}'.format(list(_enum.GetEnumNames())))
    print('\tThe enum Names: {}'.format(list(_enum.GetEnumValues())))
    print('\n')


# Loop through each interface.
for _interface in interfaces:

    # Display for the User.
    print('{} INTERFACE'.format(_interface.Name))
    print('-'*20)

    # Grab all their members.
    interface_members = _interface.GetMembers()

    # Then loop through each of the members.
    for interface_member in interface_members:
        
        # Let's see if the member is a method.
        if interface_member.MemberType == System.Reflection.MemberTypes.Method:

            # Only show methods that can be called by the user. (IsSpecialName is False)
            if interface_member.IsSpecialName == False:

                
                # Grab the name of the method.
                method_name = interface_member.Name

                # Grab the return type of the method.
                method_return_type = interface_member.ReturnType

                # Grab method parameters
                method_parameters = interface_member.GetParameters()


                # Display for the User.
                print('\tMETHOD:')
                print('\t' + "-"*20)
                print("\tMethod: Name={}, Type=System.Reflection.MemberTypes.Method".format(method_name))
                print("\tMethod: Name={}, TypeEnumeration={}".format(method_name, interface_member.MemberType))
                print("\tMethod: Name={}, ReturnType={}".format(method_name, method_return_type))
                print('\n')


                # Determine if there are any paramaters to parse
                if method_parameters is not None:

                    
                    # If there is then loop through them.
                    for parameter in method_parameters:


                        # Grab the parameter name.
                        parameter_name = parameter.Name

                        # Grab the parameter type.
                        parameter_type = parameter.ParameterType

                        # Grab the IsOptional flag.
                        is_parameter_optional = parameter.IsOptional

                        # Grab the parameter position.
                        parameter_position = parameter.Position

                        # Grab the DefaultValue flag.
                        parameter_has_default_value = parameter.HasDefaultValue


                        print('\t\tPARAMETER:')
                        print('\t\t' + "-"*20)
                        print("\t\tParameter: Name={}, Type={}".format(parameter_name, parameter_type))
                        print("\t\tParameter: Name={}, IsOptional={}".format(parameter_name, is_parameter_optional))
                        print("\t\tParameter: Name={}, Position={}".format(parameter_name, parameter_position))
                        print("\t\tParameter: Name={}, HasDefaultValue={}".format(parameter_name, parameter_has_default_value))


                        # If it does...
                        if parameter_has_default_value == True:
                            
                            # Grab the Default Value
                            parameter_default_value = parameter.DefaultValue

                            # Display for the User.
                            print("\t\tParameter: Name={}, DefaultValue={}".format(parameter_name, parameter_default_value))
                        
                        print('\n')

        # Let's see if the member is a property.
        elif interface_member.MemberType == System.Reflection.MemberTypes.Property:

            # Grab the name.
            property_name = interface_member.Name

            # Does the property have a special name?
            property_special_name = interface_member.IsSpecialName

            # Can we read the property.
            can_read_flag = interface_member.CanRead

            # Can we write the property.
            can_write_flag = interface_member.CanWrite

            # What System.Type is the Property?
            property_type = interface_member.PropertyType.Name

            if interface_member.SetMethod:

                # What is the name of the Set_Method() for this property?
                set_method = interface_member.SetMethod.Name
            
            else:

                set_method = 'No Set Method'

            
            if interface_member.GetMethod:

                # What is the name of the Get_Method() for this property?
                get_method = interface_member.GetMethod.Name

            else:

                get_method = 'No Get Method'

            # Display for the User.
            print('\tPROPERTY:')
            print('\t' + "-"*20)
            print("\tProperty: Name={}, Type=System.Reflection.MemberTypes.Property".format(property_name))
            print("\tProperty: Name={}, TypeEnumeration={}".format(property_name, interface_member.MemberType))
            print("\tProperty: Name={}, IsSpecialName={}".format(property_name, property_special_name))
            print("\tProperty: Name={}, PropertyType={}".format(property_name, property_type))
            print("\tProperty: Name={}, CanWrite={}".format(property_name, can_write_flag))
            print("\tProperty: Name={}, CanRead={}".format(property_name, can_read_flag))
            print("\tProperty: Name={}, SetMethod={}".format(property_name, set_method))
            print("\tProperty: Name={}, GetMethod={}".format(property_name, get_method))
            print('\n')

        # Let's see if the member is a Field.
        elif interface_member.MemberType == System.Reflection.MemberTypes.Field:

            # Grab the name.
            field_name = interface_member.Name

            # Is the field private?
            is_private_flag = interface_member.IsPrivate

            # Is the field public?
            is_public_flag = interface_member.IsPublic

            # What class declares this member?
            declaring_type = interface_member.DeclaringType

            # Get the Field Type
            field_type = interface_member.FieldType

            # Display for the User.
            print('\tFIELD:')
            print('\t' + "-"*20)
            print("\tField: Name={}, Type=System.Reflection.MemberTypes.Field".format(field_name))
            print("\tField: Name={}, TypeEnumeration={}".format(field_name, interface_member.MemberType))
            print("\tField: Name={}, IsPrivate={}".format(field_name, is_private_flag))
            print("\tField: Name={}, IsPublic={}".format(field_name, is_public_flag))
            print("\tField: Name={}, DeclaringType={}".format(field_name, declaring_type))
            print("\tField: Name={}, FieldType={}".format(field_name, field_type))
            print('\n')     
        
        # Let's see if the member is a Event.
        if interface_member.MemberType == System.Reflection.MemberTypes.Event:

            # Grab the name
            event_name = interface_member.Name

            # Grab the Add Method
            if interface_member.AddMethod:
                event_add_method = interface_member.AddMethod.Name
            else:
                event_add_method = 'No Add Method'

            # Grab the Remove Method
            if interface_member.RemoveMethod:
                event_remove_method = interface_member.RemoveMethod.Name
            else:
                event_remove_method = 'No Remove Method'

            # Grab the Raise Method
            if interface_member.RaiseMethod:
                event_raise_method = interface_member.RaiseMethod.Name
            else:
                event_raise_method = 'No Raise Method'

            # What Class declares this event
            event_declaring_type = interface_member.DeclaringType

            # What is the Event Handler Type
            event_handler_type = interface_member.EventHandlerType


            # Display for the User.
            print('\tEVENT:')
            print('\t' + "-"*20)
            print("\tEvent: Name={}, Type=System.Reflection.MemberTypes.Event".format(event_name))
            print("\tEvent: Name={}, TypeEnumeration={}".format(event_name, interface_member.MemberType))
            print("\tEvent: Name={}, AddMethod={}".format(event_name, event_add_method))
            print("\tEvent: Name={}, RemoveMethod={}".format(event_name, event_remove_method))
            print("\tEvent: Name={}, RaiseMethod={}".format(event_name, event_raise_method))
            print("\tEvent: Name={}, DeclaringType={}".format(event_name, event_declaring_type))
            print("\tEvent: Name={}, EventHandlerType={}".format(event_name, event_handler_type))
            print('\n')     

print('\n')
