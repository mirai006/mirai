Attribute VB_Name = "m000"
Option Explicit

Public Function ConvertResourceName(Name As String) As String

    ConvertResourceName = Replace(Name, "-", "")
    ConvertResourceName = Replace(ConvertResourceName, "(", "")
    ConvertResourceName = Replace(ConvertResourceName, ")", "")

End Function

Public Function ConvertRefResourceName(Name As String) As String

    ConvertRefResourceName = "!Ref " & ConvertResourceName(Name)

End Function

Public Function ConvertImportValueResourceName(Name As String) As String

    ConvertImportValueResourceName = "!ImportValue " & Name

End Function

Public Function GetVPCName() As String

    GetVPCName = Sheets("VPC").Cells(5, 5).Value
    
End Function

Public Function GetProjectName() As String

    GetProjectName = Sheets("VPC").Cells(5, 4).Value
    
End Function

Public Function GetACL(SurperSubnet As String) As String

    GetACL = GetProjectName & "-" & Sheets("ToolSetting").Cells(27, 4).Value & "(" & SurperSubnet & ")"
 
End Function

