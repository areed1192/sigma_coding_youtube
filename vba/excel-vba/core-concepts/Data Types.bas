Attribute VB_Name = "DataTypes"
Sub DataTypes()
    
   'Name: Variant
   'Allocation: 16 to 22 bytes
   'Range: N/A
    
   'Declare Variables with Variant Data Type
    Dim EmptyValue As Variant
    Dim ErrorValue As Variant
    Dim NothingValue As Variant
    Dim NullValue As Variant
    
   'This means that no value has been assigned to the variable.
    EmptyValue = Empty
    
   'Error help us for identifying errors in our procedure.
    ErrorValue = Error
    
   'Nothing is the uninitialized state of an object variable.
    Set NothingValue = Nothing
    
   'This indicates that the variable is absent of data.
    NullValue = Null
    
End Sub

Sub ByteDataType()

    'Name: Byte
    'Allocation: 1 Bytes
    'Range: 0 to 255
    
    Dim ByteValue As Byte
        ByteValue = 100
        
End Sub

Sub BooleanDataType()

    'Name: Boolean
    'Allocation: 2 Bytes
    'Range: True or False
    
    Dim BooleanValue As Boolean
        BooleanValue = True
        BooleanValue = False
        
End Sub


Sub CurrencyDataType()

    'Name: Currency
    'Allocation: 8 Bytes
    'Range: -922,337,203,685,477.5808 and 922,337,203,685,477.5807
    
    Dim CurrencyValue As Currency
        CurrencyValue = 200000000
        
        'Using Character Type Method
        CurrencyValue2@ = 200000000
        
End Sub

Sub DateDataType()
    
    'Name: Date
    'Allocation: 8 Bytes
    'Range Dates: Dates between January 1, 100 and December 31, 9999
    'Range Times: Times between midnight (00:00:00) and 23:59:59
    
    Dim DateValue As Date
        DateValueSerial = 43435 '<<< This is 12/1/2018
        DateValueLiteral = #12/1/2018#
        
        TimeValueSerial = 0.54  '<<< This is 1:00:00PM
        TimeValueLiteral = #1:00:00 PM#
        
        DateTimeValueSerial = 43435.54 '<<< This is 12/1/2018 1:00:00 PM
        DateTimeValueLiteral = #12/1/2018 1:00:00 PM#
        
        DateValueSerialNeg = -1 '<<< This is 12/30/1899
            
End Sub

Sub DecimalDataType()

    'Name: Decimal
    'Allocation: 14 Byte
    'Range: +/-79,228,162,514,264,337,593,543,950,335 with no decial point / +/-7.9228162514264337593543950335 with 28 places to the right of the decimal
    
    Dim DeciVal As Variant
        DeciVal = 1.15
        DeciVal = CDec(DeciVal) '<<< HAVE TO CONVERT TO DECIMAL

End Sub


Sub DoubleDataType()

    'Name: Double
    'Allocation: 8 Byte
    'Range: -1.79769313486231E308 to -4.94065645841247E-324 for negative numbers / 4.94065645841247E-324 to 1.79769313486232E308 for positive numbers
    
    Dim DoubleVal As Double
        DoubleVal = 5.5
        
        'Using Character Type Method
        DoubleVal# = -5.5

End Sub

Sub IntegerDataType()

    'Name: Integer
    'Allocation: 2 Byte
    'Range: -32,768 to 32,767
    
    Dim IntVal As Integer
        IntVal = 100
        
        'Using Character Type Method
        IntVal% = -100

End Sub

Sub LongDataType()

    'Name: Long
    'Allocation: 4 Byte
    'Range: -2,147,483,648 and 2,147,483,647
    
    Dim LongVal As Long
        LongVal = 2000000
        
        'Using Character Type Method
        LongVal& = -20000000

End Sub
        
Sub LongLongDataType()

    'Name: LongLong <<< ONLY VALID ON 64-BIT PLATFORMS
    'Allocation: 8 Byte
    'Range: -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807
    
    Dim LngLngVal As LongLong
        LngLngVal = -10000000000#
        LngLngVal = 10000000000#

End Sub
        
Sub LongPtrDataType()

    'Name: LongPtr <<< CHANGES DEPENDING ON THE PLATFORM
    'Allocation: 8 Byte
    'Range: -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807
    
    Dim LngPtrVal As LongPtr
        LngPtrVal = 1000000 '<<< 32 Bit Version Platform it becomes Long Data Type
        LngPtrVal = 1000000 '<<< 64 Bit Version Platform it becomes LongLong Data Type

End Sub
        

Sub ObjectDataType()

    'Name: Object
    'Allocation: 4 Byte
    'Range: N/A
    
    Dim Rng As Object
        Set Rng = Range("A1:A4")

End Sub
        
        
Sub SingleDataType()

    'Name: Single
    'Allocation: 4 Byte
    'Range: -3.402823E38 to -1.401298E-45 / 1.401298E-45 to 3.402823E38
    
    Dim SingVal As Single
        SingVal = -1.4
        
        'Using Character Type Method
        SingVal! = 1.4

End Sub
        
        
Sub StringDataType()

    'Name: String Variable / String Fixed
    'Allocation: 10 Byte + length of string / Length of String
    'Range: 0 to approximately 2 billion characters / 1 to approximately 65,400
    
    Dim StrVal As String
        StrVal = "Hello"
        
        'Using Character Type Method
        StrVal$ = "10"

End Sub


Sub UserDefinedDataType()

    'Name: User-Defined Types
    'Allocation: Varies
    'Range: Varies
    
    Type aType
       Field1 As String
       Field2 As Integer
       Field3 As Boolean
    End Type
    
    '("Hello", 10, True)

End Sub

