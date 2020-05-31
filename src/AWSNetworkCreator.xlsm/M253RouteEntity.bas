Attribute VB_Name = "M253RouteEntity"
Option Explicit

Public Function SetCFn_Resources_RouteEntry() As String

    Dim Row As Integer
    Dim InformationRow As Integer
 
    Row = 5
    InformationRow = 4
    
    SetCFn_Resources_RouteEntry = ""

    Do While Sheets("CreateRoute").Cells(Row, 3).Value <> ""

        ResetIndent 0

        SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndentP & Sheets("CreateRoute").Cells(Row, 3).Value & ":" & vbCrLf

        SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndentP & Sheets("CreateRoute").Cells(InformationRow, 4).Value & ": " & Sheets("CreateRoute").Cells(Row, 4).Value & vbCrLf
        SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndent & "Properties:" & vbCrLf
        SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndentP & Sheets("CreateRoute").Cells(InformationRow, 5).Value & ": " & Sheets("CreateRoute").Cells(Row, 5).Value & vbCrLf
        SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndent & Sheets("CreateRoute").Cells(InformationRow, 6).Value & ": " & Sheets("CreateRoute").Cells(Row, 6).Value & vbCrLf
        
        If Sheets("CreateRoute").Cells(Row, 7).Value <> "" Then
            SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndent & Sheets("CreateRoute").Cells(InformationRow, 7).Value & ": " & Sheets("CreateRoute").Cells(Row, 7).Value & vbCrLf
        End If
        
        If Sheets("CreateRoute").Cells(Row, 8).Value <> "" Then
            SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndent & Sheets("CreateRoute").Cells(InformationRow, 8).Value & ": " & Sheets("CreateRoute").Cells(Row, 8).Value & vbCrLf
        End If
        
        If Sheets("CreateRoute").Cells(Row, 9).Value <> "" Then
            SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndent & Sheets("CreateRoute").Cells(InformationRow, 9).Value & ": " & Sheets("CreateRoute").Cells(Row, 9).Value & vbCrLf
        End If
        
        If Sheets("CreateRoute").Cells(Row, 10).Value <> "" Then
            SetCFn_Resources_RouteEntry = SetCFn_Resources_RouteEntry & GetIndent & Sheets("CreateRoute").Cells(InformationRow, 10).Value & ": " & Sheets("CreateRoute").Cells(Row, 10).Value & vbCrLf
        End If
        
        Row = Row + 1
        
    Loop
    
End Function


