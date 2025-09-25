VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDebug 
   Caption         =   "DebugTools"
   ClientHeight    =   12750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   12750
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox TraceBox 
      Height          =   8655
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   15266
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDebug.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' Then send the EM_SCROLLCARET message to scroll the caret into view.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_SCROLLCARET As Long = &HB7

Private Sub Form_Load()
    Me.TraceBox.text = vbNullString
End Sub

Public Sub add_text_tracebox(ByVal s As String)
    With Me.TraceBox
        Me.TraceBox.text = Me.TraceBox.text & s & vbCrLf
        Me.TraceBox.SelStart = Len(Me.TraceBox.text)
        Me.TraceBox.SelLength = 0
        SendMessage Me.TraceBox.hWnd, EM_SCROLLCARET, 0, 0
    End With
End Sub
