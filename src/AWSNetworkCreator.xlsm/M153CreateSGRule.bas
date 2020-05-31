Attribute VB_Name = "M153CreateSGRule"
Option Explicit

Private CreateSGRoleRow As Integer
Public Sub CreateSecurityGroupRule()

    Sheets("CreateSGEgressRule").Range("C5:K1004").ClearContents
    Sheets("CreateSGIngressRule").Range("C5:K1004").ClearContents

    CreateSecurityGroupEgressRule
    CreateSecurityGroupIngressRule

End Sub

Private Sub CreateSecurityGroupEgressRule()

    Dim ConvertACLRow As Integer
    Dim CreateSGRoleRow As Integer
       
    ConvertACLRow = 5
    CreateSGRoleRow = 5

    Do While Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value <> ""
    
        If Sheets("ConvertACL").Cells(ConvertACLRow, 5).Value <> "" Then
                
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 3).Value = ConvertResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 6).Value) & Format(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value, "00000")
        
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 4).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 6).Value)
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 5).Value = "AWS::EC2::SecurityGroupEgress"
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 6).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 18).Value
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 7).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 20).Value
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 8).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 21).Value
            
            If Sheets("ConvertACL").Cells(ConvertACLRow, 16).Value = "" Then
                Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 9).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 17).Value
            Else
                Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 10).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 16).Value)
            End If
            
            Sheets("CreateSGEgressRule").Cells(CreateSGRoleRow, 11).Value = """" & Right("000" & Int(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value / 100), 3) & " : Rule Number /100 on ACL Sheet of the " & Sheets("ToolSetting").Cells(7, 4).Value & Sheets("ToolSetting").Cells(8, 4).Value & ".xlsm"""
            
            CreateSGRoleRow = CreateSGRoleRow + 1

        End If
    
    ConvertACLRow = ConvertACLRow + 1
    
    Loop
    
End Sub

Private Sub CreateSecurityGroupIngressRule()

    Dim ConvertACLRow As Integer
    Dim CreateSGRoleRow As Integer
       
    ConvertACLRow = 5
    CreateSGRoleRow = 5

    Do While Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value <> ""
    
        If Sheets("ConvertACL").Cells(ConvertACLRow, 15).Value <> "" Then
                
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 3).Value = ConvertResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 16).Value) & Format(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value, "00000")
        
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 4).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 16).Value)
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 5).Value = "AWS::EC2::SecurityGroupIngress"
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 6).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 18).Value
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 7).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 20).Value
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 8).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 21).Value
            
            If Sheets("ConvertACL").Cells(ConvertACLRow, 6).Value = "" Then
                Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 9).Value = Sheets("ConvertACL").Cells(ConvertACLRow, 7).Value
            Else
                Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 10).Value = ConvertImportValueResourceName(Sheets("ConvertACL").Cells(ConvertACLRow, 6).Value)
            End If
            
            Sheets("CreateSGIngressRule").Cells(CreateSGRoleRow, 11).Value = """" & Right("000" & Int(Sheets("ConvertACL").Cells(ConvertACLRow, 3).Value / 100), 3) & " : Rule Number /100 on ACL Sheet of the " & Sheets("ToolSetting").Cells(7, 4).Value & Sheets("ToolSetting").Cells(8, 4).Value & ".xlsm"""
            
            CreateSGRoleRow = CreateSGRoleRow + 1

        End If
    
    ConvertACLRow = ConvertACLRow + 1
    
    Loop
    
End Sub


