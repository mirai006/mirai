Attribute VB_Name = "M151ConvertACL"

Option Explicit

Private RuleNumber As Integer
Private S_SS As String     ' Source SuperSubnet Name
Private S_ACL As String    ' Source Network ACL Name
Private S_SGP As String    ' Source Security Group Name
Private S_CIDR As String   ' Source CIDR
Private S_PrtC As String   ' Source Protocol Code
Private S_PrtN As String   ' Source Protocol Number
Private S_PotF As String   ' Source PortRange From
Private S_PotT As String   ' Source PortRange To
Private S_ICMPC As String  ' Source ICMP Code
Private S_ICMPT As String  ' Source ICMP Type

Private D_SS As String     ' Distination SuperSubnet Name
Private D_ACL As String    ' Distination Network ACL Name
Private D_SGP As String    ' Distination Security Group Name
Private D_CIDR As String   ' Distination CIDR
Private D_PrtC As String   ' Distination Protocol Code
Private D_PrtN As String   ' Distination Protocol Number
Private D_PotF As String   ' Distination PortRange From
Private D_PotT As String   ' Distination PortRange To
Private D_ICMPC As String  ' Distination ICMP Code
Private D_ICMPT As String  ' Distination ICMP Type

Private ACLRow As Integer
Private WriteRow As Integer

Public Sub ConvertACL()
 
    ACLRow = 5
    WriteRow = 5
    
    Sheets("ConvertACL").Range("C5:W1006").ClearContents

    Do While Sheets("ACL").Cells(ACLRow, 3).Value <> ""

        RuleNumber = Sheets("ACL").Cells(ACLRow, 3).Value * 100 + 1
        
        SearchSourceSubnet

        ACLRow = ACLRow + 1

    Loop

End Sub

Private Sub SearchSourceSubnet()

    Dim Row As Integer
    Dim VPCRow As Integer
    
    Row = 5
    
    Do While Sheets("SurperSubnet").Cells(Row, 3).Value <> ""

        If Sheets("ACL").Cells(ACLRow, 4).Value = Sheets("SurperSubnet").Cells(Row, 9).Value Then
        
            S_SS = Sheets("SurperSubnet").Cells(Row, 4).Value
            S_ACL = Sheets("SurperSubnet").Cells(Row, 13).Value
            S_SGP = Sheets("SurperSubnet").Cells(Row, 12).Value
            S_CIDR = Sheets("SurperSubnet").Cells(Row, 11).Value
            
            SearchSoucePort
            
            If Sheets("SurperSubnet").Cells(Row, 9).Value = "VPC" Then

                VPCRow = 6
                
                Do While Sheets("SurperSubnet").Cells(VPCRow, 3).Value <> ""

                    If Sheets("SurperSubnet").Cells(VPCRow, 5).Value = "O" Then
                    
                        S_SS = Sheets("SurperSubnet").Cells(VPCRow, 4).Value
                        S_ACL = Sheets("SurperSubnet").Cells(VPCRow, 13).Value
                        S_SGP = Sheets("SurperSubnet").Cells(VPCRow, 12).Value
                        S_CIDR = Sheets("SurperSubnet").Cells(VPCRow, 11).Value
                        D_ACL = ""
                                    
                        SearchSoucePort
                    
                    End If
                
                    VPCRow = VPCRow + 1
                    
                Loop
        
            End If

        End If

        Row = Row + 1

    Loop

End Sub

Private Sub SearchSoucePort()

    Dim Row As Integer
    
    Row = 6
    
    Do While Sheets("OSPortNumber").Cells(Row, 4).Value <> ""

        If Sheets("ACL").Cells(ACLRow, 5).Value = Sheets("OSPortNumber").Cells(Row, 4).Value Then
        
            ' S_PrtC, S_PrtN Destnation
            S_PotF = Sheets("OSPortNumber").Cells(Row, 5).Value
            S_PotT = Sheets("OSPortNumber").Cells(Row, 6).Value
            
            ' S_ICMPC, S_ICMPT, Destnation
            
            SerchDestinaionSubnet
            
        End If
        Row = Row + 1

    Loop

End Sub

Private Sub SerchDestinaionSubnet()

    Dim Row As Integer
    Dim VPCRow As Integer
    
    Row = 5
    
    Do While Sheets("SurperSubnet").Cells(Row, 3).Value <> ""

        If Sheets("ACL").Cells(ACLRow, 6).Value = Sheets("SurperSubnet").Cells(Row, 9).Value Then
                
            D_SS = Sheets("SurperSubnet").Cells(Row, 4).Value
            D_ACL = Sheets("SurperSubnet").Cells(Row, 13).Value
            D_SGP = Sheets("SurperSubnet").Cells(Row, 12).Value
            D_CIDR = Sheets("SurperSubnet").Cells(Row, 11).Value
        
            SerchDestinaionServicePort
            
            If Sheets("SurperSubnet").Cells(Row, 9).Value = "VPC" Then

                VPCRow = 6
                
                    Do While Sheets("SurperSubnet").Cells(VPCRow, 3).Value <> ""

                 If Sheets("SurperSubnet").Cells(VPCRow, 5).Value = "O" Then
                        
                            D_SS = Sheets("SurperSubnet").Cells(VPCRow, 4).Value
                            D_ACL = Sheets("SurperSubnet").Cells(VPCRow, 13).Value
                            D_SGP = Sheets("SurperSubnet").Cells(VPCRow, 12).Value
                            D_CIDR = Sheets("SurperSubnet").Cells(VPCRow, 11).Value
                            S_ACL = ""
                                        
                            SerchDestinaionServicePort
                        
                        End If
                    
                    VPCRow = VPCRow + 1
                    
                Loop
        
            End If
            
        End If

        Row = Row + 1

    Loop

End Sub

Private Sub SerchDestinaionServicePort()

    Dim Row As Integer
    
    Row = 6
    
    Do While Sheets("ServicePort").Cells(Row, 4).Value <> ""

        If Sheets("ACL").Cells(ACLRow, 7).Value = Sheets("ServicePort").Cells(Row, 4).Value Then
            
            S_PrtC = Sheets("ServicePort").Cells(Row, 5).Value
            D_PrtC = Sheets("ServicePort").Cells(Row, 5).Value
            S_PrtN = Sheets("ServicePort").Cells(Row, 6).Value
            D_PrtN = Sheets("ServicePort").Cells(Row, 6).Value
            D_PotF = Sheets("ServicePort").Cells(Row, 7).Value
            D_PotT = Sheets("ServicePort").Cells(Row, 8).Value
            
            If Sheets("ServicePort").Cells(Row, 9).Value = "" Then
                S_ICMPC = -1
                D_ICMPC = -1
            Else
                S_ICMPC = Sheets("ServicePort").Cells(Row, 9).Value
                D_ICMPC = Sheets("ServicePort").Cells(Row, 9).Value
            End If
            
            If Sheets("ServicePort").Cells(Row, 10).Value = "" Then
                S_ICMPT = -1
                D_ICMPT = -1
            Else
                S_ICMPT = Sheets("ServicePort").Cells(Row, 10).Value
                D_ICMPT = Sheets("ServicePort").Cells(Row, 10).Value
            End If
            
            WriteInformation
            
        End If
        Row = Row + 1

    Loop

End Sub

Private Sub WriteInformation()

    Sheets("ConvertACL").Cells(WriteRow, 3).Value = RuleNumber

    Sheets("ConvertACL").Cells(WriteRow, 4).Value = S_SS
    Sheets("ConvertACL").Cells(WriteRow, 5).Value = S_ACL
    Sheets("ConvertACL").Cells(WriteRow, 6).Value = S_SGP
    Sheets("ConvertACL").Cells(WriteRow, 7).Value = S_CIDR
    Sheets("ConvertACL").Cells(WriteRow, 8).Value = S_PrtC
    Sheets("ConvertACL").Cells(WriteRow, 9).Value = S_PrtN
    Sheets("ConvertACL").Cells(WriteRow, 10).Value = S_PotF
    Sheets("ConvertACL").Cells(WriteRow, 11).Value = S_PotT
    Sheets("ConvertACL").Cells(WriteRow, 12).Value = S_ICMPC
    Sheets("ConvertACL").Cells(WriteRow, 13).Value = S_ICMPT
    
    Sheets("ConvertACL").Cells(WriteRow, 14).Value = D_SS
    Sheets("ConvertACL").Cells(WriteRow, 15).Value = D_ACL
    Sheets("ConvertACL").Cells(WriteRow, 16).Value = D_SGP
    Sheets("ConvertACL").Cells(WriteRow, 17).Value = D_CIDR
    Sheets("ConvertACL").Cells(WriteRow, 18).Value = D_PrtC
    Sheets("ConvertACL").Cells(WriteRow, 19).Value = D_PrtN
    Sheets("ConvertACL").Cells(WriteRow, 20).Value = D_PotF
    Sheets("ConvertACL").Cells(WriteRow, 21).Value = D_PotT
    Sheets("ConvertACL").Cells(WriteRow, 22).Value = D_ICMPC
    Sheets("ConvertACL").Cells(WriteRow, 23).Value = D_ICMPT
    
    RuleNumber = RuleNumber + 1
    WriteRow = WriteRow + 1

End Sub
