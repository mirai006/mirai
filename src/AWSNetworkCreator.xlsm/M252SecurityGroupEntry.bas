Attribute VB_Name = "M252SecurityGroupEntry"
Option Explicit

Public Function SetCFn_Resources_SecurityGroupEntry() As String

    SetCFn_Resources_SecurityGroupEntry = ""
    SetCFn_Resources_SecurityGroupEntry = SetCFn_Resources_SecurityGroupEntry + SetCFn_Resources_SecurityGroupEgress
    SetCFn_Resources_SecurityGroupEntry = SetCFn_Resources_SecurityGroupEntry + SetCFn_Resources_SecurityGroupIngress
    
End Function

Private Function SetCFn_Resources_SecurityGroupEgress() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    
    SetCFn_Resources_SecurityGroupEgress = ""

    Do While Sheets("CreateSGEgressRule").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndentP & Sheets("CreateSGEgressRule").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndentP & Sheets("CreateSGEgressRule").Cells(InformationRow, 5).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndentP & Sheets("CreateSGEgressRule").Cells(InformationRow, 4).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & Sheets("CreateSGEgressRule").Cells(InformationRow, 6).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & Sheets("CreateSGEgressRule").Cells(InformationRow, 7).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 7).Value & vbCrLf
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & Sheets("CreateSGEgressRule").Cells(InformationRow, 8).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 8).Value & vbCrLf
        
        If Sheets("CreateSGEgressRule").Cells(Row, 10).Value <> "" Then
            SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & Sheets("CreateSGEgressRule").Cells(InformationRow, 10).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 10).Value & vbCrLf
        Else
            SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & Sheets("CreateSGEgressRule").Cells(InformationRow, 9).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 9).Value & vbCrLf
        End If
        
        SetCFn_Resources_SecurityGroupEgress = SetCFn_Resources_SecurityGroupEgress & GetIndent & Sheets("CreateSGEgressRule").Cells(InformationRow, 11).Value & ": " & Sheets("CreateSGEgressRule").Cells(Row, 11).Value & vbCrLf
        
        Row = Row + 1
        
    Loop
    
End Function

Private Function SetCFn_Resources_SecurityGroupIngress() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    
    SetCFn_Resources_SecurityGroupIngress = ""

    Do While Sheets("CreateSGIngressRule").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndentP & Sheets("CreateSGIngressRule").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndentP & Sheets("CreateSGIngressRule").Cells(InformationRow, 5).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndentP & Sheets("CreateSGIngressRule").Cells(InformationRow, 4).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & Sheets("CreateSGIngressRule").Cells(InformationRow, 6).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & Sheets("CreateSGIngressRule").Cells(InformationRow, 7).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 7).Value & vbCrLf
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & Sheets("CreateSGIngressRule").Cells(InformationRow, 8).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 8).Value & vbCrLf
        
        If Sheets("CreateSGIngressRule").Cells(Row, 10).Value <> "" Then
            SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & Sheets("CreateSGIngressRule").Cells(InformationRow, 10).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 10).Value & vbCrLf
        Else
            SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & Sheets("CreateSGIngressRule").Cells(InformationRow, 9).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 9).Value & vbCrLf
        End If
        
        SetCFn_Resources_SecurityGroupIngress = SetCFn_Resources_SecurityGroupIngress & GetIndent & Sheets("CreateSGIngressRule").Cells(InformationRow, 11).Value & ": " & Sheets("CreateSGIngressRule").Cells(Row, 11).Value & vbCrLf

        Row = Row + 1
        
    Loop
    
End Function


