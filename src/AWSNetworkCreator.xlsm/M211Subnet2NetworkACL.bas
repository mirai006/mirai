Attribute VB_Name = "M211Subnet2NetworkACL"
Option Explicit

Public Function SetCFn_Resources_Subnet2NetworkACL() As String

    Dim Row As Integer
    Dim InformationRow As Integer
    
    Row = 5
    InformationRow = 4
    SetCFn_Resources_Subnet2NetworkACL = ""

    Do While Sheets("CreateSubnet").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_Subnet2NetworkACL = SetCFn_Resources_Subnet2NetworkACL & GetIndentP & Sheets("CreateSubnet").Cells(Row, 9).Value & ":" & vbCrLf
        SetCFn_Resources_Subnet2NetworkACL = SetCFn_Resources_Subnet2NetworkACL & GetIndentP & Sheets("CreateSubnet").Cells(InformationRow, 10).Value & ": " & Sheets("CreateSubnet").Cells(Row, 10).Value & vbCrLf
        SetCFn_Resources_Subnet2NetworkACL = SetCFn_Resources_Subnet2NetworkACL & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_Subnet2NetworkACL = SetCFn_Resources_Subnet2NetworkACL & GetIndentP & "SubnetId: !Ref " & Sheets("CreateSubnet").Cells(Row, 3).Value & vbCrLf
        SetCFn_Resources_Subnet2NetworkACL = SetCFn_Resources_Subnet2NetworkACL & GetIndent & "NetworkAclId: !Ref " & Sheets("CreateSubnet").Cells(Row, 11).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function



