Attribute VB_Name = "M105RouteTable"
Option Explicit

Public Sub CreateRouteTableInforamtion()

    Dim CheckRow As Integer
    Dim WriteRow As Integer
    Dim CheckName As String
    Dim NameSame As Boolean
    Dim i As Integer

    CheckRow = 5
    WriteRow = 5
    
    Sheets("CreateRT").Range("c5:j24").ClearContents  'check


    Do While Sheets("RouteTable").Cells(CheckRow, 4).Value <> ""
    
        If Sheets("RouteTable").Cells(CheckRow, 5).Value <> "" Then
            
            CheckName = Sheets("RouteTable").Cells(CheckRow, 5)
            NameSame = vbFalse
            
            For i = 5 To CheckRow - 1
            
                If Sheets("RouteTable").Cells(i, 5) = CheckName Then
                    NameSame = vbTrue
                End If
            
            Next

            If Not NameSame Then
        
                Sheets("CreateRT").Cells(WriteRow, 3).Value = ConvertResourceName(Sheets("RouteTable").Cells(CheckRow, 5).Value)
                Sheets("CreateRT").Cells(WriteRow, 4).Value = "AWS::EC2::RouteTable"
                Sheets("CreateRT").Cells(WriteRow, 5).Value = ConvertRefResourceName(GetVPCName)
                Sheets("CreateRT").Cells(WriteRow, 6).Value = Sheets("RouteTable").Cells(CheckRow, 5).Value
                
                WriteRow = WriteRow + 1
 
            End If
            
        End If

        CheckRow = CheckRow + 1
    Loop

End Sub
