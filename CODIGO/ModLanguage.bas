Attribute VB_Name = "ModLanguage"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Const LOCALE_USER_DEFAULT = &H400

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const LOCALE_SCOUNTRY = &H6
Private Const LOCALE_SLANGUAGE = &H2
Private Const LOCALE_SENGLANGUAGE = &H1001
Private Const MAX_BUF As Long = 260

Public Enum e_language
    Spanish = 1          ' Español (LATAM)
    English = 2
    Portuguese = 3
    French = 4
    Italian = 5
'    Spanish_Spain = 6   ' Español (España)
End Enum


Public language As e_language
Public JsonLanguage As Object
     
Private Function GetCountryName() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SCOUNTRY, sBuf, MAX_BUF)
    GetCountryName = Left$(sBuf, lRet - 1)
End Function

Private Function GetLocaleEngLanguage() As String
    Dim lID As Long, sBuf As String, lRet As Long
    lID = GetUserDefaultLCID
    sBuf = String$(MAX_BUF, Chr(0))
    lRet = GetLocaleInfo(lID, LOCALE_SENGLANGUAGE, sBuf, MAX_BUF)
    GetLocaleEngLanguage = Left$(sBuf, lRet - 1)
End Function

Public Function FileToString(strFileName As String) As String
    Dim IFile As Variant
    IFile = FreeFile
    Open strFileName For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

Public Function DirLanguage() As String
On Error GoTo DirLanguage_Err
    DirLanguage = App.path & "\Languages\"
    Exit Function
DirLanguage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_General.DirLanguage", Erl)
    Resume Next
End Function

Public Sub SetLanguageApplication()
On Error GoTo ErrorHandler
    
    Dim Localization As String
    Dim LangFilePath As String
    Dim LangFileContent As String
    ' Retrieve localization setting
    Localization = GetSetting("OPCIONES", "Language")

    ' If no localization is set, determine the default language based on system locale
    If Len(Localization) = 0 Then

        Select Case GetLocaleEngLanguage
            Case "Spanish"
                language = e_language.Spanish
                
            Case "English"
                language = e_language.English

            Case "Portuguese"
                language = e_language.Portuguese
                
            Case "French"
                language = e_language.French
                
            Case "Italian"
                language = e_language.Italian
'            Case "Spanish_Spain"
'                language = e_language.Spanish_Spain
            Case Else

                ' Default to English if system locale is unsupported
                language = e_language.English
        End Select



        ' Save the determined language as the default localization setting
        SaveSetting "OPCIONES", "Language", language
    Else
        ' Use the stored localization setting
        language = Localization
    End If

    ' Build the file path for the language JSON file
    LangFilePath = DirLanguage() & language & ".json"

    ' Validate the existence of the language file
    If dir(LangFilePath) = "" Then
        Err.Raise vbObjectError + 1, "SetLanguageApplication", "Language file not found: " & LangFilePath
    End If

    ' Load and parse the language JSON file
    LangFileContent = FileToString(LangFilePath)
    Set JsonLanguage = JSON.parse(LangFileContent)

    Exit Sub

ErrorHandler:
    MsgBox "Error in SetLanguageApplication: " & Err.Description, vbCritical, "Error"
    ' Optional: Fallback to a default language in case of an error
    language = e_language.English
End Sub

