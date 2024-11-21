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
    Spanish = 1
    English = 2
End Enum

Public language As e_language
     
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

Public Sub SetLanguageApplication()
    Dim Localization As String
    Localization = GetSetting("OPCIONES", "Localization")
    If Len(Localization) = 0 Then
        Select Case GetLocaleEngLanguage
                Case "English"
                    language = e_language.English
                Case "Spanish"
                    language = e_language.Spanish
        End Select
        Call SaveSetting("OPCIONES", "Localization", language)
    Else
         language = Localization
    End If
End Sub
