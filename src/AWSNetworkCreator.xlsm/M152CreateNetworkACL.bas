Attribute VB_Name = "M152CreateNetworkACL"
Option Explicit

Private CreateACLRuleRow As Integer

Public Sub CreateACLRule()

    Sheets("CreateACLRoule").Range("C5:n4004").ClearContents
    
    CreateACLRuleRow = 5

    CreateACLEgressInformations
    CreateACLIngressInformations

End Sub

Private Sub CreateACLEgressInformations()

    Dim ConvertACLRow As Integer
    
    ConvertACLRow = 5

    Do While Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value <> ""
    
        If Sheets("ConvertACL").Cells(ConvertACLRow, 5).Value <> "" Then
                
            ' Outgoing Information
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 3).Value = ConvertResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 5).Value) & Format(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value, "00000") & "E"
        
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 4).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 5).Value)
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 5).Value = "AWS::EC2::NetworkAclEntry"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 6).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 7).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 19).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 8).Value = "allow"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 9).Value = "'true"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 10).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 17).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 11).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 22).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 12).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 23).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 13).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 20).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 14).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 21).Value
            
            CreateACLRuleRow = CreateACLRuleRow + 1
            
            ' Outgoing Return Information
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 3).Value = ConvertResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 5).Value) & Format(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value, "00000") & "I"
        
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 4).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 5).Value)
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 5).Value = "AWS::EC2::NetworkAclEntry"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 6).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 7).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 9).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 8).Value = "allow"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 9).Value = "'false"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 10).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 17).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 11).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 12).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 12).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 13).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 13).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 10).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 14).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 11).Value
            
            CreateACLRuleRow = CreateACLRuleRow + 1
        
        End If
    
    ConvertACLRow = ConvertACLRow + 1
    
    Loop
    
End Sub

Private Sub CreateACLIngressInformations()

    Dim ConvertACLRow As Integer
    
    ConvertACLRow = 5

    Do While Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value <> ""
    
        If Sheets("ConvertACL").Cells(ConvertACLRow, 15).Value <> "" Then
           
           ' Inbound Information
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 3).Value = ConvertResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 15).Value) & Format(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value, "00000") & "I"
        
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 4).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 15).Value)
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 5).Value = "AWS::EC2::NetworkAclEntry"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 6).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 7).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 9).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 8).Value = "allow"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 9).Value = "'false"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 10).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 7).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 11).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 22).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 12).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 23).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 13).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 20).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 14).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 21).Value
            
            CreateACLRuleRow = CreateACLRuleRow + 1
            
            ' Inbound Information
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 3).Value = ConvertResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 15).Value) & Format(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value, "00000") & "E"
        
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 4).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 15).Value)
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 5).Value = "AWS::EC2::NetworkAclEntry"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 6).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 7).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 19).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 8).Value = "allow"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 9).Value = "'true"
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 10).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 7).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 11).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 12).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 12).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 13).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 13).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 10).Value
            Sheets("CreateACLRoule").Cells(CreateACLRuleRow, 14).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 11).Value
            
            CreateACLRuleRow = CreateACLRuleRow + 1
    
        End If
    
    
    ConvertACLRow = ConvertACLRow + 1
    
    Loop
    
End Sub
