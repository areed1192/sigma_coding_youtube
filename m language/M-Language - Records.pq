

/*
    What is a Record?
    -----------------
    A Record is a set of fields. A field is a name/value pair where the name is a text value that is unique within the field’s 
    record. The syntax for record values allows the names to be written without quotes, a form also referred to as identifiers. 
    An identifier can take the following two forms:

        1. identifier_name such as OrderID.
        2. #"identifier name" such as #"Today's data is: ".

*/


let

    // Lets define a new record, with 5 fields.
    myRecord = [Name = "Alex", Age = 26, Gender = "Male", CollegeDegree = true, Salary = 1.00],

    // We can add a new field with Record.AddField function.
    myRecordAddField = Record.AddField(myRecord, "Hometown","San Diego"),

    // We can remove a field with the Record.RemoveFields function.
    myRecordRemoveField = Record.RemoveFields(myRecordAddField, "Gender"), // {"Gender", "Salary"} this will remove two fields Gender and Salary.

    // We can rename a field with the Record.RenameFields function.
    myRecordRenamed = Record.RenameFields(myRecordRemoveField, {"Name", "FirstName"}), // {{"Salary", "SalaryCurr"}, {"Name", "FirstName"}} this will rename multiple fields.

    // We can reorder fields with the Record.ReorderFields function.
    myRecordReordered = Record.ReorderFields(myRecordRenamed, {"FirstName", "Age", "CollegeDegree", "Hometown", "Salary"}),

    // We can combine two separate records with the Record.Combine function.
    myRecordCombined = Record.Combine({myRecordReordered, [Birthday = "11/11/1992", FavoriteColor = "Blue"]}), // Notice that both records are wrapped in a list {}
    
    // We can transform the type of data each field has with the Record.TransformFields function.
    myRecordTransformed = Record.TransformFields(myRecordCombined, {{"FirstName", Text.From}, 
                                                                    {"Age", Number.From}, 
                                                                    {"CollegeDegree", Logical.From},
                                                                    {"Hometown", Text.From},
                                                                    {"Salary", Decimal.From}, 
                                                                    {"Birthday", Date.FromText},
                                                                    {"FavoriteColor", Text.From}}),

    // We can grab all the field values with the Record.FieldValues function.
    myRecordValues = Record.FieldValues(myRecordTransformed),

    // We can grab all the field names with the Record.FieldNames function.
    myRecordFields = Record.FieldNames(myRecordTransformed),

    // We can select a fields in the following ways:
    myName = myRecordTransformed[FirstName],
    myAge = Record.Field(myRecordTransformed, "Age"),
    myAgeAndName = Record.SelectFields(myRecordTransformed, {"Age", "FirstName"}),

    // We can check whether a field exists using the Record.HasFields function.
    hasName = Record.HasFields(myRecordTransformed, "FirstName"), // true
    hasNameAndFavoriteBook = Record.HasFields(myRecordTransformed, {"FirstName", "FavoriteBook"}), // {true, false}

    // We can convert a record to a list with Record.ToList function.
    myRecordList = Record.ToList(myRecordTransformed),

    // We can convert a record to a table with the Record.ToTable function.
    myRecordTable = Record.ToTable(myRecordTransformed),

    // We can try and access a field and if doesn't exist set what happens using the Record.FieldOrDefault function.
    myNull = Record.FieldOrDefault(myRecordTransformed, "FavoriteBook"), // returns null as default.
    myBook = Record.FieldOrDefault(myRecordTransformed, "FavoriteBook", "Star Wars"), // returns Star Wars as default.

    /*

        Enumeration         Name                    Description
        -----------         ----                    -----------
            0               MissingField.Error      An optional parameter in record and table functions indicating that missing fields should result in an error. (This is the default parameter value.)
            1               MissingField.Ignore     An optional parameter in record and table functions indicating that missing fields should be ignored.
            2               MissingField.UseNull    An optional parameter in record and table functions indicating that missing fields should be included as null values.


        In use it would look like the following:

        Expression:
        -----------
        myAgeAndName = Record.SelectFields(myRecordTransformed, {"Age", "FavoriteBook"}, 2) 

        Output:
        ----------- 
        {26, null} as 2 represents MissingField.UseNull

    */

    // Missing Fields Option Change Example
    myAgeAndNameBehaviorChange = Record.SelectFields(myRecordTransformed, {"Age", "FavoriteBook"}, 2) 

in 
    
    myAgeAndNameBehaviorChange