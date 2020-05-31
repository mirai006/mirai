Attribute VB_Name = "M202Subnet"
Option Explicit

Public Function SetCFn_Resources_Subnet() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Resources_Subnet = ""

    Do While Sheets("CreateSubnet").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndentP & Sheets("CreateSubnet").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndentP & Sheets("CreateSubnet").Cells(InformationRow, 4).Value & ": " & Sheets("CreateSubnet").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndentP & Sheets("CreateSubnet").Cells(InformationRow, 5).Value & ": " & Sheets("CreateSubnet").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndent & Sheets("CreateSubnet").Cells(InformationRow, 6).Value & ": " & Sheets("CreateSubnet").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndent & Sheets("CreateSubnet").Cells(InformationRow, 7).Value & ": " & Sheets("CreateSubnet").Cells(Row, 7).Value & vbCrLf
 
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndent & "Tags: " & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndent & "- Key: " & ConvertTagName(Sheets("CreateSubnet").Cells(InformationRow, 8).Value) & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & GetIndent & "  Value: " & Sheets("CreateSubnet").Cells(Row, 8).Value & vbCrLf
        SetCFn_Resources_Subnet = SetCFn_Resources_Subnet & SetToolInformation

        Row = Row + 1
        
    Loop
    
End Function

Public Function SetCFn_Output_Subnet() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Output_Subnet = ""

    Do While Sheets("CreateSubnet").Cells(Row, 3).Value <> ""

        ResetIndent 0
         
        SetCFn_Output_Subnet = SetCFn_Output_Subnet & GetIndentP & "Export" & Sheets("CreateSubnet").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Output_Subnet = SetCFn_Output_Subnet & GetIndentP & "Value: !Ref " & Sheets("CreateSubnet").Cells(Row, 3).Value & vbCrLf
        SetCFn_Output_Subnet = SetCFn_Output_Subnet & GetIndent & "Export:" & vbCrLf
        SetCFn_Output_Subnet = SetCFn_Output_Subnet & GetIndentP & "Name: " & Sheets("CreateSubnet").Cells(Row, 8).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function



