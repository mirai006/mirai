Attribute VB_Name = "M203NetworkACL"
Option Explicit

Public Function SetCFn_Resources_NetwrokACL() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Resources_NetwrokACL = ""

    Do While Sheets("CreateACL").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndentP & Sheets("CreateACL").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndentP & Sheets("CreateACL").Cells(InformationRow, 4).Value & ": " & Sheets("CreateACL").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndentP & Sheets("CreateACL").Cells(InformationRow, 5).Value & ": " & Sheets("CreateACL").Cells(Row, 5).Value & vbCrLf
 
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndent & "Tags: " & vbCrLf
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndent & "- Key: " & ConvertTagName(Sheets("CreateACL").Cells(InformationRow, 6).Value) & vbCrLf
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & GetIndent & "  Value: " & Sheets("CreateACL").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_NetwrokACL = SetCFn_Resources_NetwrokACL & SetToolInformation

        Row = Row + 1
        
    Loop
    
End Function

Public Function SetCFn_Output_NetwrokACL() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Output_NetwrokACL = ""

    Do While Sheets("CreateACL").Cells(Row, 3).Value <> ""

        ResetIndent 0
        
        SetCFn_Output_NetwrokACL = SetCFn_Output_NetwrokACL & GetIndentP & "Export" & Sheets("CreateACL").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Output_NetwrokACL = SetCFn_Output_NetwrokACL & GetIndentP & "Value: !Ref " & Sheets("CreateACL").Cells(Row, 3).Value & vbCrLf
        SetCFn_Output_NetwrokACL = SetCFn_Output_NetwrokACL & GetIndent & "Export:" & vbCrLf
        SetCFn_Output_NetwrokACL = SetCFn_Output_NetwrokACL & GetIndentP & "Name: " & Sheets("CreateACL").Cells(Row, 6).Value & vbCrLf
        
        Row = Row + 1
        
    Loop
    
End Function




