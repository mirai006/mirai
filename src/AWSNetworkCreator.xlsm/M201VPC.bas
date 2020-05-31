Attribute VB_Name = "M201VPC"
Option Explicit

Public Function SetCFn_Resources_VPC() As String

    Dim Row As Integer
    Dim InformationRow As Integer
    
    Row = 5
    InformationRow = 4
    
    SetCFn_Resources_VPC = ""
    
    Do While Sheets("CreateVPC").Cells(Row, 3).Value <> ""
        
        ResetIndent 0
     
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndentP & Sheets("CreateVPC").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndentP & Sheets("CreateVPC").Cells(InformationRow, 4).Value & ": " & Sheets("CreateVPC").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndentP & Sheets("CreateVPC").Cells(InformationRow, 5).Value & ": " & Sheets("CreateVPC").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & Sheets("CreateVPC").Cells(InformationRow, 6).Value & ": " & Sheets("CreateVPC").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & Sheets("CreateVPC").Cells(InformationRow, 7).Value & ": " & Sheets("CreateVPC").Cells(Row, 7).Value & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & Sheets("CreateVPC").Cells(InformationRow, 8).Value & ": " & Sheets("CreateVPC").Cells(Row, 8).Value & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & "Tags: " & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & "- Key: " & ConvertTagName(Sheets("CreateVPC").Cells(InformationRow, 9).Value) & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & GetIndent & "  Value: " & Sheets("CreateVPC").Cells(Row, 9).Value & vbCrLf
        SetCFn_Resources_VPC = SetCFn_Resources_VPC & SetToolInformation
        
        Row = Row + 1
        
    Loop
    
End Function

Public Function SetCFn_Output_VPC() As String

    Dim Row As Integer
    Dim InformationRow As Integer
    
    Row = 5
    InformationRow = 4
    
    SetCFn_Output_VPC = ""
    
    Do While Sheets("CreateVPC").Cells(Row, 3).Value <> ""
        
        ResetIndent 0
     
        SetCFn_Output_VPC = SetCFn_Output_VPC & GetIndentP & "Export" & Sheets("CreateVPC").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Output_VPC = SetCFn_Output_VPC & GetIndentP & "Value: !Ref " & Sheets("CreateVPC").Cells(Row, 3).Value & vbCrLf
        SetCFn_Output_VPC = SetCFn_Output_VPC & GetIndent & "Export:" & vbCrLf
        SetCFn_Output_VPC = SetCFn_Output_VPC & GetIndentP & "Name: " & Sheets("CreateVPC").Cells(Row, 9).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function






