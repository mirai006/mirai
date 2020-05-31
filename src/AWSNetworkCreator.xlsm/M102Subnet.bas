Attribute VB_Name = "M102Subnet"
Option Explicit

Public Sub CreateSunetInforamtion()

    Dim AZStartRow As Integer
    Dim CheckRow As Integer
    Dim WiriteRow As Integer
    
    AZStartRow = 20

    CheckRow = 6
    WiriteRow = 5

    Dim AZRow As Integer
    
    Dim SuperSubnetName As String
    Dim SuperSubnetCIDR As String
    Dim SettingCIDR As String
    Dim Subnetmask As Integer
    Dim SubnetPreFix As String
    Dim Length As Integer
    
    Sheets("CreateSubnet").Range("c5:n604").ClearContents
    
    SubnetPreFix = Sheets("ToolSetting").Cells(25, 4).Value
    
    Do While Sheets("SurperSubnet").Cells(CheckRow, 4).Value <> ""
          
        If Sheets("SurperSubnet").Cells(CheckRow, 5).Value = "O" Then
          
        SuperSubnetName = Sheets("SurperSubnet").Cells(CheckRow, 4).Value
        SuperSubnetCIDR = Sheets("SurperSubnet").Cells(CheckRow, 7).Value
        Subnetmask = Sheets("SurperSubnet").Cells(CheckRow, 10).Value
    
        AZRow = 0
        
        Do While Sheets("VPC").Cells(AZStartRow + AZRow, 3).Value <> ""
    
            If Sheets("VPC").Cells(AZStartRow + AZRow, 4).Value = "O" Then
        
                SettingCIDR = NextCIDR(SuperSubnetCIDR, Subnetmask, AZRow)
                Sheets("CreateSubnet").Cells(WiriteRow, 8).Value = GetProjectName() & "-" & SubnetPreFix & "-" & SuperSubnetName & "-" & Get3_Dot4Octet(SettingCIDR)
                
                Sheets("CreateSubnet").Cells(WiriteRow, 3).Value = ConvertResourceName(Sheets("CreateSubnet").Cells(WiriteRow, 8).Value)
                Sheets("CreateSubnet").Cells(WiriteRow, 4).Value = "AWS::EC2::Subnet"
                Sheets("CreateSubnet").Cells(WiriteRow, 5).Value = ConvertRefResourceName(GetVPCName)
                Sheets("CreateSubnet").Cells(WiriteRow, 6).Value = SettingCIDR & "/" & Subnetmask
                Sheets("CreateSubnet").Cells(WiriteRow, 7).Value = Sheets("VPC").Cells(AZStartRow + AZRow, 3).Value
                
                Sheets("CreateSubnet").Cells(WiriteRow, 9).Value = GetProjectName() & Sheets("ToolSetting").Cells(27, 4).Value & SuperSubnetName & Get3_4Octet(SettingCIDR)
                Sheets("CreateSubnet").Cells(WiriteRow, 10).Value = "AWS::EC2::SubnetNetworkAclAssociation"
                Sheets("CreateSubnet").Cells(WiriteRow, 11).Value = ConvertResourceName(GetACL(SuperSubnetName))
                
                Sheets("CreateSubnet").Cells(WiriteRow, 12).Value = GetProjectName() & Sheets("ToolSetting").Cells(28, 4).Value & SuperSubnetName & Get3_4Octet(SettingCIDR)
                Sheets("CreateSubnet").Cells(WiriteRow, 13).Value = "AWS::EC2::SubnetRouteTableAssociation"
                Sheets("CreateSubnet").Cells(WiriteRow, 14).Value = GetProjectName() & Sheets("ToolSetting").Cells(28, 4).Value & Sheets("SurperSubnet").Cells(CheckRow, 8).Value
                
                WiriteRow = WiriteRow + 1
                
            End If
    
            AZRow = AZRow + 1
    
        Loop
        
        End If
    
        CheckRow = CheckRow + 1
        
    Loop

End Sub

Private Function NextCIDR(CIDR As String, Subnetmask As Integer, Hop As Integer) As String

    Dim CIDR1 As Long
    Dim CIDR2 As Long
    Dim CIDR3 As Long
    Dim CIDR4 As Long

    Dim AnalysisCIDR As String
    
    AnalysisCIDR = CIDR
    

    CIDR1 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    AnalysisCIDR = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR2 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    AnalysisCIDR = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR3 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    CIDR4 = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR4 = CIDR4 + 2 ^ (32 - Subnetmask) * Hop
    
    CIDR3 = CIDR3 + CIDR4 \ 256
    CIDR4 = CIDR4 Mod 256

    CIDR2 = CIDR2 + CIDR3 \ 256
    CIDR3 = CIDR3 Mod 256
    
    CIDR1 = CIDR1 + CIDR2 \ 256
    CIDR2 = CIDR2 Mod 256
    
    NextCIDR = CIDR1 & "." & CIDR2 & "." & CIDR3 & "." & CIDR4

End Function

Private Function Get3_4Octet(CIDR As String) As String
    
    Dim CIDR1 As Long
    Dim CIDR2 As Long
    Dim CIDR3 As Long
    Dim CIDR4 As Long

    Dim AnalysisCIDR As String
    
    AnalysisCIDR = CIDR
    

    CIDR1 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    AnalysisCIDR = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR2 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    AnalysisCIDR = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR3 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    CIDR4 = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    Get3_4Octet = Format(CIDR3, "000") & Format(CIDR4, "000")
 

End Function

Private Function Get3_Dot4Octet(CIDR As String) As String
    
    Dim CIDR1 As Long
    Dim CIDR2 As Long
    Dim CIDR3 As Long
    Dim CIDR4 As Long

    Dim AnalysisCIDR As String
    
    AnalysisCIDR = CIDR
    

    CIDR1 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    AnalysisCIDR = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR2 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    AnalysisCIDR = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    CIDR3 = Val(Left(AnalysisCIDR, InStr(AnalysisCIDR, ".") - 1))
    CIDR4 = Mid(AnalysisCIDR, InStr(AnalysisCIDR, ".") + 1)
    
    Get3_Dot4Octet = Format(CIDR3, "000") & "-" & Format(CIDR4, "000")
 

End Function

