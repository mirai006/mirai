Attribute VB_Name = "M200CFnConstant"
Option Explicit

Public Function SetCloudFormationVersion() As String
    
    ResetIndent 0
    
    SetCloudFormationVersion = "AWSTemplateFormatVersion: '2010-09-09'" & vbCrLf

End Function

Public Function SetCloudFormationResources() As String
    
    ResetIndent 0
    
    SetCloudFormationResources = vbCrLf
    SetCloudFormationResources = SetCloudFormationResources & "Resources: " & vbCrLf
    SetCloudFormationResources = SetCloudFormationResources & vbCrLf

End Function

Public Function SetCloudFormationOutputs() As String
    
    ResetIndent 0
    
    SetCloudFormationOutputs = vbCrLf
    SetCloudFormationOutputs = SetCloudFormationOutputs & "Outputs: " & vbCrLf
    SetCloudFormationOutputs = SetCloudFormationOutputs & vbCrLf

End Function

