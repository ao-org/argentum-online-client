Attribute VB_Name = "ModSettings"
Option Explicit

Const CustomSettingsFile = "\..\Recursos\OUTPUT\Configuracion.ini"
Const DefaultSettingsFile = "\..\Recursos\OUTPUT\DefaultSettings.ini"

Public Function InitializeSettings() As Boolean
    
    If Not FileExist(App.path & DefaultSettingsFile, vbArchive) Then
        InitializeSettings = False
        Exit Function
    End If
    If Not FileExist(App.path & CustomSettingsFile, vbArchive) Then
        Call FileSystem.FileCopy(App.path & DefaultSettingsFile, App.path & CustomSettingsFile)
    End If
    InitializeSettings = True
End Function

Public Function GetSetting(ByVal Section As String, ByVal Name As String) As String
    Dim currentValue As String
    currentValue = GetVar(App.path & CustomSettingsFile, Section, Name)
    If currentValue = "" Then
        currentValue = GetVar(App.path & DefaultSettingsFile, Section, Name)
    End If
    GetSetting = currentValue
End Function

Public Sub SaveSetting(ByVal Section As String, ByVal Name As String, ByVal Value As String)
    Call WriteVar(App.path & CustomSettingsFile, Section, Name, Value)
End Sub


