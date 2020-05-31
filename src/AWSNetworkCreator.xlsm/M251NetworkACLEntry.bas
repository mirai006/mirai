Attribute VB_Name = "M251NetworkACLEntry"
Option Explicit

Public Function SetCFn_Resources_NetworkACLEntry() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    
    SetCFn_Resources_NetworkACLEntry = ""

    Do While Sheets("CreateACLRoule").Cells(Row, 3).Value <> ""

        ResetIndent 0
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndentP & Sheets("CreateACLRoule").Cells(Row, 3).Value & ":" & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndentP & Sheets("CreateACLRoule").Cells(InformationRow, 5).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndentP & Sheets("CreateACLRoule").Cells(InformationRow, 4).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & Sheets("CreateACLRoule").Cells(InformationRow, 6).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 6).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & Sheets("CreateACLRoule").Cells(InformationRow, 7).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 7).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & Sheets("CreateACLRoule").Cells(InformationRow, 8).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 8).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & Sheets("CreateACLRoule").Cells(InformationRow, 9).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 9).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & Sheets("CreateACLRoule").Cells(InformationRow, 10).Value & ": " & Sheets("CreateACLRoule").Cells(Row, 10).Value & vbCrLf

        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "Icmp:" & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "  Code: " & Sheets("CreateACLRoule").Cells(Row, 11).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "  Type: " & Sheets("CreateACLRoule").Cells(Row, 12).Value & vbCrLf

        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "PortRange:" & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "  From: " & Sheets("CreateACLRoule").Cells(Row, 13).Value & vbCrLf
        SetCFn_Resources_NetworkACLEntry = SetCFn_Resources_NetworkACLEntry & GetIndent & "  To: " & Sheets("CreateACLRoule").Cells(Row, 14).Value & vbCrLf


        Row = Row + 1
        
    Loop
    
End Function



