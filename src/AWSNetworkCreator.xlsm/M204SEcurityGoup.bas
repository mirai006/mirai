Attribute VB_Name = "M204SEcurityGoup"
Option Explicit

Public Function SetCFn_Resources_SecurityGroup() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 6
    InformationRow = 4
    SetCFn_Resources_SecurityGroup = ""
    

    Do While Sheets("CreateSG").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndentP & Sheets("CreateSG").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndentP & Sheets("CreateSG").Cells(InformationRow, 4).Value & ": " & Sheets("CreateSG").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndentP & Sheets("CreateSG").Cells(InformationRow, 6).Value & ": " & Sheets("CreateSG").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & Sheets("CreateSG").Cells(InformationRow, 7).Value & ":" & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & "- " & Sheets("CreateSG").Cells(InformationRow + 1, 7).Value & ": " & Sheets("CreateSG").Cells(Row, 7).Value & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & "  " & Sheets("CreateSG").Cells(InformationRow + 1, 8).Value & ": " & Sheets("CreateSG").Cells(Row, 8).Value & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & Sheets("CreateSG").Cells(InformationRow, 9).Value & ": " & Sheets("CreateSG").Cells(Row, 9).Value & vbCrLf
        
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & "Tags: " & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & "- Key: " & ConvertTagName(Sheets("CreateSG").Cells(InformationRow, 10).Value) & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & GetIndent & "  Value: " & Sheets("CreateSG").Cells(Row, 10).Value & vbCrLf
        SetCFn_Resources_SecurityGroup = SetCFn_Resources_SecurityGroup & SetToolInformation

        Row = Row + 1
        
    Loop
    
End Function

Public Function SetCFn_Output_SecurityGroup() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 6
    InformationRow = 4
    SetCFn_Output_SecurityGroup = ""
    

    Do While Sheets("CreateSG").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Output_SecurityGroup = SetCFn_Output_SecurityGroup & GetIndentP & "Export" & Sheets("CreateSG").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Output_SecurityGroup = SetCFn_Output_SecurityGroup & GetIndentP & "Value: !Ref " & Sheets("CreateSG").Cells(Row, 3).Value & vbCrLf
        SetCFn_Output_SecurityGroup = SetCFn_Output_SecurityGroup & GetIndent & "Export:" & vbCrLf
        SetCFn_Output_SecurityGroup = SetCFn_Output_SecurityGroup & GetIndentP & "Name: " & Sheets("CreateSG").Cells(Row, 10).Value & vbCrLf
        
        Row = Row + 1
        
    Loop
    
End Function



