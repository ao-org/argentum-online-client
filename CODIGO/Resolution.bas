Attribute VB_Name = "Resolution"
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

Private Const CCDEVICENAME          As Long = 32
Private Const CCFORMNAME            As Long = 32

Private Const DM_BITSPERPEL         As Long = &H40000
Private Const DM_PELSWIDTH          As Long = &H80000
Private Const DM_PELSHEIGHT         As Long = &H100000
Private Const DM_DISPLAYFREQUENCY   As Long = &H400000

Private Const CDS_TEST              As Long = &H4
Private Const ENUM_CURRENT_SETTINGS As Long = -1

Private Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Private oldResHeight As Long
Private oldResWidth  As Long
Private oldDepth     As Integer
Private oldFrequency As Long

Private bResChange As Boolean

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long

'TODO : Change this to not depend on any external public variable using args instead!

Public Sub SetResolution()
    
    On Error GoTo SetResolution_Err
    

    '***************************************************
    'Autor: Unknown
    'Last Modification: 03/29/08
    'Changes the display resolution if needed.
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    ' 03/29/2008: Maraxus - Retrieves current settings storing display depth and frequency for proper restoration.
    '***************************************************
    Dim lRes              As Long
    Dim MidevM            As typDevMODE
    Dim CambiarResolucion As Boolean
    
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
    
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
    
    If NoRes And Not PantallaCompleta Then
        CambiarResolucion = (oldResWidth <= 1024 Or oldResHeight <= 768)
    Else
        CambiarResolucion = (oldResWidth <> 1024 Or oldResHeight <> 768)
    End If
    
    If CambiarResolucion Then
        
        With MidevM
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency
            
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 1024
            .dmPelsHeight = 768
            .dmBitsPerPel = 32

        End With
        
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
        
        If frmMain.Visible Then frmMain.Top = 0: frmMain.Left = 0
        If frmOpciones.Visible Then frmOpciones.Top = (Screen.Height - frmOpciones.Height) \ 2: frmOpciones.Left = (Screen.Width - frmOpciones.Width) \ 2
        
        bResChange = True
    Else
        bResChange = False

    End If

    
    Exit Sub

SetResolution_Err:
    Call RegistrarError(Err.Number, Err.Description, "Resolution.SetResolution", Erl)
    Resume Next
    
End Sub

Public Sub ResetResolution()
    
    On Error GoTo ResetResolution_Err
    

    '***************************************************
    'Autor: Unknown
    'Last Modification: 03/29/08
    'Changes the display resolution if needed.
    'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
    ' 03/29/2008: Maraxus - Properly restores display depth and frequency.
    '***************************************************
    Dim typDevM As typDevMODE
    Dim lRes    As Long
    
    If bResChange Then

        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
        
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency

        End With
        
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
        
        If frmMain.Visible Then frmMain.Top = (Screen.Height - frmMain.Height) \ 2: frmMain.Left = (Screen.Width - frmMain.Width) \ 2
        If frmOpciones.Visible Then frmOpciones.Top = (Screen.Height - frmOpciones.Height) \ 2: frmOpciones.Left = (Screen.Width - frmOpciones.Width) \ 2

    End If

    
    Exit Sub

ResetResolution_Err:
    Call RegistrarError(Err.Number, Err.Description, "Resolution.ResetResolution", Erl)
    Resume Next
    
End Sub
