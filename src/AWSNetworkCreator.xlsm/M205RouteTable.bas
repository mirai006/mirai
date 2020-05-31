Attribute VB_Name = "M205RouteTable"
Option Explicit

Public Function SetCFn_Resources_RouteTable() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Resources_RouteTable = ""
    

    Do While Sheets("CreateRT").Cells(Row, 6).Value <> ""

        ResetIndent 0
        
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndentP & Sheets("CreateRT").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndentP & Sheets("CreateRT").Cells(InformationRow, 4).Value & ": " & Sheets("CreateRT").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndentP & Sheets("CreateRT").Cells(InformationRow, 5).Value & ": " & Sheets("CreateRT").Cells(Row, 5).Value & vbCrLf
        
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndent & "Tags: " & vbCrLf
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndent & "- Key: " & ConvertTagName(Sheets("CreateRT").Cells(InformationRow, 6).Value) & vbCrLf
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & GetIndent & "  Value: " & Sheets("CreateRT").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_RouteTable = SetCFn_Resources_RouteTable & SetToolInformation

        Row = Row + 1
        
    Loop
    
End Function

Public Function SetCFn_Output_RouteTable() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Output_RouteTable = ""
    

    Do While Sheets("CreateRT").Cells(Row, 6).Value <> ""

        ResetIndent 0
        
        SetCFn_Output_RouteTable = SetCFn_Output_RouteTable & GetIndentP & "Export" & Sheets("CreateRT").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Output_RouteTable = SetCFn_Output_RouteTable & GetIndentP & "Value: !Ref " & Sheets("CreateRT").Cells(Row, 3).Value & vbCrLf
        SetCFn_Output_RouteTable = SetCFn_Output_RouteTable & GetIndent & "Export:" & vbCrLf
        SetCFn_Output_RouteTable = SetCFn_Output_RouteTable & GetIndentP & "Name: " & Sheets("CreateRT").Cells(Row, 6).Value & vbCrLf
        
        Row = Row + 1
        
    Loop
    
End Function




