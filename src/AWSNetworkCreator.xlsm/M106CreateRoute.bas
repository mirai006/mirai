Attribute VB_Name = "M106CreateRoute"
Option Explicit

Public Sub CreateRouteInforamtion()

    Dim CheckRow As Integer
    Dim WriteRow As Integer

    CheckRow = 5
    WriteRow = 5
    
    Sheets("CreateRoute").Range("c5:j44").ClearContents  'check


    Do While Sheets("RouteTable").Cells(CheckRow, 4).Value <> ""
    
        If Sheets("RouteTable").Cells(CheckRow, 5).Value <> "" Then

            If Sheets("RouteTable").Cells(CheckRow, 6).Value <> "" Then
        
                Sheets("CreateRoute").Cells(WriteRow, 3).Value = ConvertResourceName(Sheets("RouteTable").Cells(CheckRow, 5).Value & Sheets("RouteTable").Cells(CheckRow, 7).Value & Sheets("RouteTable").Cells(CheckRow, 8).Value & Sheets("RouteTable").Cells(CheckRow, 9).Value & Sheets("RouteTable").Cells(CheckRow, 10).Value)
                Sheets("CreateRoute").Cells(WriteRow, 4).Value = "AWS::EC2::Route"
                Sheets("CreateRoute").Cells(WriteRow, 5).Value = ConvertImportValueResourceName(Sheets("RouteTable").Cells(CheckRow, 5).Value)
                Sheets("CreateRoute").Cells(WriteRow, 6).Value = Sheets("RouteTable").Cells(CheckRow, 6).Value
                
                If Sheets("RouteTable").Cells(CheckRow, 7).Value <> "" Then
                    Sheets("CreateRoute").Cells(WriteRow, 7).Value = ConvertImportValueResourceName(Sheets("RouteTable").Cells(CheckRow, 7).Value)
                End If
                
                If Sheets("RouteTable").Cells(CheckRow, 8).Value <> "" Then
                    Sheets("CreateRoute").Cells(WriteRow, 8).Value = ConvertImportValueResourceName(Sheets("RouteTable").Cells(CheckRow, 8).Value)
                End If
                
                If Sheets("RouteTable").Cells(CheckRow, 9).Value <> "" Then
                    Sheets("CreateRoute").Cells(WriteRow, 9).Value = ConvertImportValueResourceName(Sheets("RouteTable").Cells(CheckRow, 9).Value)
                End If
                
                If Sheets("RouteTable").Cells(CheckRow, 10).Value <> "" Then
                    Sheets("CreateRoute").Cells(WriteRow, 10).Value = ConvertImportValueResourceName(Sheets("RouteTable").Cells(CheckRow, 10).Value)
                End If
                
                WriteRow = WriteRow + 1
 
            End If
            
        End If

        CheckRow = CheckRow + 1
    Loop

End Sub

