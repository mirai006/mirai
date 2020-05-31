Attribute VB_Name = "M104SecurityGroup"
Option Explicit

Public Sub CreateSecurityGroupInforamtion()

    Dim CheckRow As Integer
    Dim WiriteRow As Integer

    CheckRow = 5
    WiriteRow = 6
    
    Sheets("CreateSG").Range("c6:j35").ClearContents


    Do While Sheets("SurperSubnet").Cells(CheckRow, 4).Value <> ""
    
        If Sheets("SurperSubnet").Cells(CheckRow, 12).Value <> "" Then
        
            Sheets("CreateSG").Cells(WiriteRow, 3).Value = ConvertResourceName(Sheets("SurperSubnet").Cells(CheckRow, 12).Value)
            Sheets("CreateSG").Cells(WiriteRow, 4).Value = "AWS::EC2::SecurityGroup"
            Sheets("CreateSG").Cells(WiriteRow, 5).Value = Sheets("SurperSubnet").Cells(CheckRow, 12).Value
            Sheets("CreateSG").Cells(WiriteRow, 6).Value = "Security Group for " & Sheets("SurperSubnet").Cells(CheckRow, 4).Value
            Sheets("CreateSG").Cells(WiriteRow, 7).Value = "127.0.0.1/32"
            Sheets("CreateSG").Cells(WiriteRow, 8).Value = "-1"
            Sheets("CreateSG").Cells(WiriteRow, 9).Value = ConvertRefResourceName(GetVPCName)
            Sheets("CreateSG").Cells(WiriteRow, 10).Value = Sheets("SurperSubnet").Cells(CheckRow, 12).Value
        
            WiriteRow = WiriteRow + 1

        End If

        CheckRow = CheckRow + 1
    Loop

End Sub

