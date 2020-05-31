Attribute VB_Name = "M206InternetGW"
Option Explicit

Public Function SetCFn_Resources_InternetGateway() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Resources_InternetGateway = ""
    

    Do While Sheets("CreateIGW").Cells(Row, 3).Value <> ""

        ResetIndent 0
        
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & GetIndentP & Sheets("CreateIGW").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & GetIndentP & Sheets("CreateIGW").Cells(InformationRow, 4).Value & ": " & Sheets("CreateIGW").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & GetIndent & "Properties:" & vbCrLf
 
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & GetIndentP & "Tags: " & vbCrLf
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & GetIndent & "- Key: " & ConvertTagName(Sheets("CreateIGW").Cells(InformationRow, 5).Value) & vbCrLf
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & GetIndent & "  Value: " & Sheets("CreateIGW").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_InternetGateway = SetCFn_Resources_InternetGateway & SetToolInformation

        Row = Row + 1
        
    Loop
    
End Function

Public Function SetCFn_Output_InternetGateway() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Output_InternetGateway = ""
    

    Do While Sheets("CreateIGW").Cells(Row, 3).Value <> ""

        ResetIndent 0
        
        SetCFn_Output_InternetGateway = SetCFn_Output_InternetGateway & GetIndentP & "Export" & Sheets("CreateIGW").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Output_InternetGateway = SetCFn_Output_InternetGateway & GetIndentP & "Value: !Ref " & Sheets("CreateIGW").Cells(Row, 3).Value & vbCrLf
        SetCFn_Output_InternetGateway = SetCFn_Output_InternetGateway & GetIndent & "Export:" & vbCrLf
        SetCFn_Output_InternetGateway = SetCFn_Output_InternetGateway & GetIndentP & "Name: " & Sheets("CreateIGW").Cells(Row, 5).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function
