Attribute VB_Name = "M200Common"
Option Explicit

Private Indent As Integer
Private IndentString As String


Public Sub ResetIndent(SetIndent As Integer)

    Indent = SetIndent
    IndentString = "  "

End Sub

Public Sub IncreaseIndent()

    Indent = Indent + 1

End Sub


Public Sub ReduceIndent()

    Indent = Indent - 1

End Sub

Public Function GetIndent() As String

    Dim i As Integer
    
    GetIndent = ""
    
    For i = 1 To Indent

        GetIndent = GetIndent & IndentString

    Next
    
End Function

Public Function GetIndentP() As String

    IncreaseIndent
    GetIndentP = GetIndent

End Function

Public Function GetIndenM() As String

    ReduceIndent
    GetIndentP = GetIndent

End Function

Public Function SetBoolean(strBoolean As String) As String

    SetBoolean = "'" & LCase(strBoolean) & "'"
    
End Function

Public Function ConvertTagName(TagName As String) As String

    ConvertTagName = TagName
    ConvertTagName = Replace(ConvertTagName, "Key:", "")
    ConvertTagName = Replace(ConvertTagName, " ", "")
    ConvertTagName = Replace(ConvertTagName, "Tag", "")
    ConvertTagName = Replace(ConvertTagName, ".", "")

End Function

Public Function SetToolInformation() As String

    SetToolInformation = ""
    
    SetToolInformation = SetToolInformation & GetIndent & "- Key: ToolVersion" & vbCrLf
    SetToolInformation = SetToolInformation & GetIndent & "  Value: " & Sheets("ToolSetting").Cells(5, 4).Value & vbCrLf
    
    SetToolInformation = SetToolInformation & GetIndent & "- Key: ToolCopyright" & vbCrLf
    SetToolInformation = SetToolInformation & GetIndent & "  Value: " & Sheets("ToolSetting").Cells(6, 4).Value & vbCrLf
    
    SetToolInformation = SetToolInformation & GetIndent & "- Key: SettingInformation(FileName)" & vbCrLf
    SetToolInformation = SetToolInformation & GetIndent & "  Value: " & Sheets("ToolSetting").Cells(7, 4).Value & Sheets("ToolSetting").Cells(8, 4).Value & ".xlsm" & vbCrLf

End Function
