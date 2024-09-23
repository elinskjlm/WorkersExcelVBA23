Attribute VB_Name = "PublicFunctions"
Option Explicit

' Why use it as public, when it could easily just be written in each form? ->
' A. Maybe it's a mistake. B. Now it's used for both changing the value when BeforeUpdate >>>
' >>> and checking the txtBox when btnAdd is clicked. C. It's probably a bad call.
' This is a function that attempts to parse a ComboBox object input as a date in the format "dd/mm/yyyy"
' THIS FUNCTION HAS A DOUBLE ROLL: IT BOTH MODIFIES THE OBJECT ITSELF, AND RETURNS A BOOLEAN TO INDICATE SUCCESS.
' This Is In Order To Make It One Address For Both Modifing And Checking The Value In the Object.
Public Function tryParseDate(ByVal dateTxt As MSForms.TextBox) As Boolean
    
    ' Set up error handling to catch any exceptions that might occur during execution
    On Error GoTo ErrHandler
    ' Declare a variable to hold the processed date string
    Dim txtDt As String
    ' Get the text value of the ComboBox object
    txtDt = dateTxt.value
    ' Replace any periods in the input string with slashes to ensure it's in the correct format
    txtDt = Replace(txtDt, ".", "/")
    ' Attempt to convert the processed date string to a date value in the correct format
    txtDt = Format(dateValue(txtDt), "dd/mm/yyyy")
    ' Set the text value of the ComboBox object to the processed date string
    dateTxt.value = txtDt
    ' If we've made it this far without any errors, return True to indicate success
    tryParseDate = True
    Exit Function

ErrHandler:
    ' If an error occurs during execution (e.g. if the input string cannot be parsed as a date), return False to indicate failure
    tryParseDate = False

End Function

Public Function AddActiveNamesToComboBox(cbo As ComboBox) As ComboBox
' It is KNOWN THAT AN ERROR occurs in this function.
' It may have something to do with arrList being an array or not.
' Therefore, the function consistenly returns Nothing,
' But - it does it after compliting it's task,
' So for the current time (21 century)- it lefted as is.

    On Error GoTo ErrHandler
   
    Dim ws As Worksheet
    Dim dvList As String
    Dim arrList As Variant
    
    'Set the worksheet where the data validation list is located
    Set ws = ThisWorkbook.Sheets("גיליון טכני")
    
    'Get the data validation list formula from cell A1
    dvList = ws.Range("hlpCellDrpDwnNames").Validation.Formula1
    
    'Remove the # sign from the start of the formula
    dvList = Right(dvList, Len(dvList) - 1)
    
    'Convert the formula into an array of values
    arrList = ws.Evaluate(dvList)
    
    ' Clear the existing items in the combo box
    cbo.Clear
    
    'Add the values to the ComboBox
    cbo.List = arrList
    
    ' Set the first item as the default selected item
    ' cbo.ListIndex = 0
    
    Set ws = Nothing
    
    ' Return the updated ComboBox object
    Set AddActiveNamesToComboBox = cbo
    
    Exit Function
    
ErrHandler:
    Set ws = Nothing
    ' Return Nothing if an error occurs
    Set AddActiveNamesToComboBox = Nothing
    
End Function

' A function to check an Israeli ID number.
' Source: https://www.excelmaster.co.il/2018/08/06/check_id_excel/
Public Function checkID(ByVal ID As String)
    Dim ID_9 As String
    ID_9 = WorksheetFunction.Rept(0, 9 - Len(ID)) & ID
    If (CInt(Mid(ID_9, 1, 1)) + CInt(Mid("0246813579", Mid(ID_9, 2, 1) + 1, 1)) + CInt(Mid(ID_9, 3, 1) + CInt(Mid("0246813579", Mid(ID_9, 4, 1) + 1, 1))) _
    + CInt(Mid(ID_9, 5, 1)) + CInt(Mid("0246813579", Mid(ID_9, 6, 1) + 1, 1)) + CInt(Mid(ID_9, 7, 1) + CInt(Mid("0246813579", Mid(ID_9, 8, 1) + 1, 1))) _
    + CInt(Mid(ID_9, 9, 1))) Mod 10 = 0 Then
        checkID = True
    Else
        checkID = False
    End If
End Function


Public Function RemoveNonDigits(inputString As String) As String
    Dim regexPattern As String
    regexPattern = "[^0-9]" ' Matches any character that is not a digit

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = regexPattern
    regex.Global = True
    
    
    
    RemoveNonDigits = regex.Replace(inputString, "")
    
    Set regex = Nothing
End Function


Function ValidatePhoneNumber(ByVal phoneNumber As String) As Boolean
    Dim regexPattern As String
    Dim regexObj As Object
    
    ' Define the regular expression pattern
    regexPattern = "^\+?(972|0)(\-)?0?(([23489]{1}\d{7})|([71,72,73,74,75,76,77]{2}\d{7})|[5]{1}\d{8})$"
    
    ' Create a regular expression object
    Set regexObj = CreateObject("VBScript.RegExp")
    
    With regexObj
        .Pattern = regexPattern
        .IgnoreCase = True
        .Global = True
    End With
    
    ' Check if the phone number matches the pattern
    If regexObj.Test(phoneNumber) Then
        ValidatePhoneNumber = True
    Else
        ValidatePhoneNumber = False
    End If
    
    Set regexObj = Nothing
    
End Function

Function ValidateEmailAddress(ByVal emailAddress As String) As Boolean
    Dim regexPattern As String
    Dim regexObj As Object
    
    ' Define the regular expression pattern for email validation
    regexPattern = "^[\w\.-]+@[\w\.-]+\.\w+$" ' very flexible
    
    ' Create a regular expression object
    Set regexObj = CreateObject("VBScript.RegExp")
    
    With regexObj
        .Pattern = regexPattern
        .IgnoreCase = True
        .Global = True
    End With
    
    ' Check if the email address matches the pattern
    If regexObj.Test(emailAddress) Then
        ValidateEmailAddress = True
    Else
        ValidateEmailAddress = False
    End If
    
    Set regexObj = Nothing
    
End Function

