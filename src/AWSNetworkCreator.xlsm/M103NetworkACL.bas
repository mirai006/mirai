Attribute VB_Name = "M103NetworkACL"
Option Explicit

Public Sub CreateNetworkACLInformation()

    Dim CheckRow As Integer
    Dim WiriteRow As Integer

    CheckRow = 5
    WiriteRow = 5
    
    Sheets("CreateACL").Range("c5:j34").ClearContents


    Do While Sheets("SurperSubnet").Cells(CheckRow, 4).Value <> ""
    
        If Sheets("SurperSubnet").Cells(CheckRow, 13).Value <> "" Then
        
            Sheets("CreateACL").Cells(WiriteRow, 3).Value = ConvertResourceName(Sheets("SurperSubnet").Cells(CheckRow, 13).Value)
            Sheets("CreateACL").Cells(WiriteRow, 4).Value = "AWS::EC2::NetworkAcl"
            Sheets("CreateACL").Cells(WiriteRow, 5).Value = ConvertRefResourceName(GetVPCName)
            Sheets("CreateACL").Cells(WiriteRow, 6).Value = Sheets("SurperSubnet").Cells(CheckRow, 13).Value


            WiriteRow = WiriteRow + 1

        End If

        CheckRow = CheckRow + 1
    Loop

End Sub
