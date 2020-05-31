Attribute VB_Name = "M001Main"
Option Explicit

Public Sub main()

    MsgBox "test"

    Dim CloudFormationConfigs As String
    Dim DateTime As String
    
    DateTime = Format(Date, "yymmdd") & "_" & Format(Now, "hhnnss")
    Sheets("ToolSetting").Cells(8, 4).Value = DateTime
     
  ' ChangeSheets
  
    CreateSunetInforamtion
    CreateNetworkACLInformation
    CreateSecurityGroupInforamtion
    CreateRouteTableInforamtion
    CreateRouteInforamtion
    ConvertACL
    CreateACLRule
    CreateSecurityGroupRule
     
  ' 1st Group
    CloudFormationConfigs = SetCloudFormationVersion
    CloudFormationConfigs = CloudFormationConfigs & SetCloudFormationResources
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_VPC
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_Subnet
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_NetwrokACL
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_SecurityGroup
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_Subnet2NetworkACL
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_InternetGateway
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_IGW2VPC
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_RouteTable
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_Subnet2RouteTable
    
    CloudFormationConfigs = CloudFormationConfigs & SetCloudFormationOutputs
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Output_VPC
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Output_Subnet
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Output_NetwrokACL
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Output_SecurityGroup
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Output_InternetGateway
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Output_RouteTable
    
    CreateCloudFomationYamlFile CloudFormationConfigs, "-01NWK"

  ' 2nd Group
    CloudFormationConfigs = SetCloudFormationVersion
    CloudFormationConfigs = CloudFormationConfigs & SetCloudFormationResources
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_RouteEntry
    
    CreateCloudFomationYamlFile CloudFormationConfigs, "-02RTT"
  ' 3rd Group
    CloudFormationConfigs = SetCloudFormationVersion
    CloudFormationConfigs = CloudFormationConfigs & SetCloudFormationResources
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_NetworkACLEntry
    
    CreateCloudFomationYamlFile CloudFormationConfigs, "-03ACL"
  
  ' 4th Group
    CloudFormationConfigs = SetCloudFormationVersion
    CloudFormationConfigs = CloudFormationConfigs & SetCloudFormationResources
    CloudFormationConfigs = CloudFormationConfigs & SetCFn_Resources_SecurityGroupEntry
    
    CreateCloudFomationYamlFile CloudFormationConfigs, "-04SGP"
    
  ' Backup Excel File
    BackupThisExcelBook
    
End Sub

Private Sub CreateCloudFomationYamlFile(strConfig As String, strSuffix As String)

    Dim FileNo As Integer
    Dim CFYamlFileName As String
 
    CFYamlFileName = ActiveWorkbook.Path & "\99CFnTemplate\" & Sheets("ToolSetting").Cells(7, 4).Value & Sheets("ToolSetting").Cells(8, 4).Value & strSuffix & ".yaml"
    
    FileNo = FreeFile

    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(ActiveWorkbook.Path & "\99CFnTemplate\") Then .CreateFolder ActiveWorkbook.Path & "\99CFnTemplate\"
    End With

    Open CFYamlFileName For Output As #FileNo
    Print #FileNo, strConfig
    Close #FileNo

End Sub

Private Sub BackupThisExcelBook()

    Dim BackupFileName As String
    
    BackupFileName = ActiveWorkbook.Path & "\Backup\" & Sheets("ToolSetting").Cells(7, 4).Value & Sheets("ToolSetting").Cells(8, 4).Value & ".xlsm"
    
    With CreateObject("Scripting.FileSystemObject")
        If Not .FolderExists(ActiveWorkbook.Path & "\Backup\") Then .CreateFolder ActiveWorkbook.Path & "\Backup\"
    End With

    ThisWorkbook.Save
    ThisWorkbook.SaveAs Filename:=BackupFileName
    ThisWorkbook.Close

End Sub
