Attribute VB_Name = "ModLenguaje"
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

'Constantes para el Api GetLocaleInfo
'************************************
Const LOCALE_USER_DEFAULT = &H400
'Const LOCALE_SENGLANGUAGE = &H1001
  
'Declaracion de la funcion Api GetLocaleInfo
Private Declare Function GetLocaleInfo _
                Lib "kernel32" _
                Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                        ByVal LCType As Long, _
                                        ByVal lpLCData As String, _
                                        ByVal cchData As Long) As Long

Private Declare Function GetThreadLocale Lib "kernel32" () As Long

Public JsonLanguage As Object

Public Enum e_language
    Spanish = 1
    English = 2
End Enum

Public language As e_language

Public Function FileToString(strFileName As String) As String
    '###################################################################################
    ' Convierte un archivo entero a una cadena de texto para almacenarla en una variable
    '###################################################################################
    Dim IFile As Variant
    
    IFile = FreeFile
    Open strFileName For Input As #IFile
        FileToString = StrConv(InputB(LOF(IFile), IFile), vbUnicode)
    Close #IFile
End Function

Public Function ObtainOperativeSystemLanguage(ByVal lInfo As Long) As String
    '*******************************************
    ' Funcion que obtiene el idioma del sistema
    '*******************************************

    Dim Buffer As String, Ret As String

    Buffer = String$(256, 0)
            
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    
    'Si Ret devuelve 0 es porque fallo la llamada al Api
    If Ret > 0 Then
        ObtainOperativeSystemLanguage = Left$(Buffer, Ret - 1)
    Else
        ObtainOperativeSystemLanguage = "No se pudo obtener el idioma del sistema."

    End If
    
End Function

Public Sub SetLanguageApplication()
    '************************************************************************************.
    ' Carga el JSON con las traducciones en un objeto para su uso a lo largo del proyecto
    '************************************************************************************

    Dim LangFile As String
    Dim Localization As String
    Localization = GetSetting("OPCIONES", "Localization")
    
    ' Si no se especifica el idioma en el archivo de configuracion, se le pregunta si quiere usar castellano
    ' y escribimos el archivo de configuracion con el idioma seleccionado
    If LenB(Localization) = 0 Then
        
        Select Case UCase(ObtainOperativeSystemLanguage(1))
            'English US
            Case "0409"
                language = e_language.English
            'Otros english
            Case "0809"
                language = e_language.English
            Case "0C09"
                language = e_language.English
            Case "1009"
                language = e_language.English
            Case "1409"
                language = e_language.English
            Case "1809"
                language = e_language.English
            Case "1c09"
                language = e_language.English
            Case "2009"
                language = e_language.English
            Case "2409"
                language = e_language.English
            Case "2809"
                language = e_language.English
            Case "2C09"
                language = e_language.English
            Case "3009"
                language = e_language.English
            Case "3409"
                language = e_language.English
            Case "4009"
                language = e_language.English
            Case "4409"
                language = e_language.English
            Case "040A"
                language = e_language.Spanish
            Case "080A"
                language = e_language.Spanish
            Case "0C0A"
                language = e_language.Spanish
            Case "100A"
                language = e_language.Spanish
            Case "140A"
                language = e_language.Spanish
            Case "180A"
                language = e_language.Spanish
            Case "1C0A"
                language = e_language.Spanish
            Case "200A"
                language = e_language.Spanish
            Case "240A"
                language = e_language.Spanish
            Case "280A"
                language = e_language.Spanish
            Case "2C0A"
                language = e_language.Spanish
            Case "300A"
                language = e_language.Spanish
            Case "380A"
                language = e_language.Spanish
            Case "3C0A"
                language = e_language.Spanish
            Case "400A"
                language = e_language.Spanish
            Case "440A"
                language = e_language.Spanish
            Case "480A"
                language = e_language.Spanish
            Case "4C0A"
                language = e_language.Spanish
            Case "500A"
                language = e_language.Spanish
            Case "540A"
                language = e_language.Spanish
            Case Else
                language = e_language.English
        End Select
        
        Call SaveSetting("OPCIONES", "Localization", language)
    Else
        language = Localization
    End If
    
End Sub
