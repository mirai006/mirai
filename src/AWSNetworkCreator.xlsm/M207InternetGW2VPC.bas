Attribute VB_Name = "M207InternetGW2VPC"
Option Explicit

Public Function SetCFn_Resources_IGW2VPC() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    SetCFn_Resources_IGW2VPC = ""
    

    Do While Sheets("CreateIGW").Cells(Row, 6).Value <> ""

        ResetIndent 0
        SetCFn_Resources_IGW2VPC = SetCFn_Resources_IGW2VPC & GetIndentP & Sheets("CreateIGW").Cells(Row, 6).Value & ":" & vbCrLf
        SetCFn_Resources_IGW2VPC = SetCFn_Resources_IGW2VPC & GetIndentP & Sheets("CreateIGW").Cells(InformationRow, 7).Value & ": " & Sheets("CreateIGW").Cells(Row, 7).Value & vbCrLf
        SetCFn_Resources_IGW2VPC = SetCFn_Resources_IGW2VPC & GetIndent & "Properties:" & vbCrLf
        
        SetCFn_Resources_IGW2VPC = SetCFn_Resources_IGW2VPC & GetIndentP & "VpcId: !Ref " & Sheets("CreateIGW").Cells(Row, 8).Value & vbCrLf
        SetCFn_Resources_IGW2VPC = SetCFn_Resources_IGW2VPC & GetIndent & "InternetGatewayId: !Ref " & Sheets("CreateIGW").Cells(Row, 3).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function



