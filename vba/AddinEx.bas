Attribute VB_Name = "AddinEx"
' Copyright (c) 2021 David Briant. All rights reserved.

Option Explicit

Private Declare PtrSafe Function c_add Lib "/Users/david/Library/Group Containers/UBF8T346G9.Office/libaddin.dylib" Alias "addDD_D" (ByVal a As Double, ByVal b As Double) As Double

Private Declare PtrSafe Function c_strlen Lib "/Users/david/Library/Group Containers/UBF8T346G9.Office/libaddin.dylib" Alias "strlenC_L" (ByVal pStr As LongLong) As Long
Private Declare PtrSafe Function c_strlen2 Lib "/Users/david/Library/Group Containers/UBF8T346G9.Office/libaddin.dylib" Alias "strlenC_L" (ByVal x As String) As Long

' returning a char* seems to be unreliable so return the pointer and subsequently use it as a LongLong
Private Declare PtrSafe Function c_concat Lib "/Users/david/Library/Group Containers/UBF8T346G9.Office/libaddin.dylib" Alias "concatCC_C" (ByVal a As String, ByVal b As String) As LongLong
Private Declare PtrSafe Sub c_strcpy Lib "/Users/david/Library/Group Containers/UBF8T346G9.Office/libaddin.dylib" Alias "strcpyCC" (ByVal pSrc As LongLong, ByVal dest As String)
Private Declare PtrSafe Sub c_free Lib "/Users/david/Library/Group Containers/UBF8T346G9.Office/libaddin.dylib" Alias "freeP" (ByVal pSrc As LongLong)


Function add(a As Double, b As Double) As Double
    On Error GoTo errorHandler
    add = c_add(a, b)
Exit Function
errorHandler:
    Debug.Print Err.Description
End Function


Function strlen(x As String) As Long
    strlen = c_strlen2(x)
End Function


Function concat(a As String, b As String) As String
    Dim pResult As LongLong, buf As String
    On Error GoTo errorHandler
    pResult = c_concat(a, b)                    ' pResult is a pointer to d owned memory
    concat = Space$(c_strlen(pResult) + 1)    ' create buffer of Excel owned memory
    c_strcpy pResult, concat                     ' copy d owned string into Excel owned
    c_free pResult                                ' free d owned memory
Exit Function
errorHandler:
    Debug.Print Err.Description
End Function

