Attribute VB_Name = "M212Subnet2RouteTable"
Option Explicit

Public Function SetCFn_Resources_Subnet2RouteTable() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Resources_Subnet2RouteTable = ""

    Do While Sheets("CreateSubnet").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_Subnet2RouteTable = SetCFn_Resources_Subnet2RouteTable & GetIndentP & Sheets("CreateSubnet").Cells(Row, 12).Value & ":" & vbCrLf
        SetCFn_Resources_Subnet2RouteTable = SetCFn_Resources_Subnet2RouteTable & GetIndentP & Sheets("CreateSubnet").Cells(InformationRow, 13).Value & ": " & Sheets("CreateSubnet").Cells(Row, 13).Value & vbCrLf
        SetCFn_Resources_Subnet2RouteTable = SetCFn_Resources_Subnet2RouteTable & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_Subnet2RouteTable = SetCFn_Resources_Subnet2RouteTable & GetIndentP & "SubnetId: !Ref " & Sheets("CreateSubnet").Cells(Row, 3).Value & vbCrLf
        SetCFn_Resources_Subnet2RouteTable = SetCFn_Resources_Subnet2RouteTable & GetIndent & "RouteTableId: !Ref " & Sheets("CreateSubnet").Cells(Row, 14).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function



