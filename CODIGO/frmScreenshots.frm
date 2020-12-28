VERSION 5.00
Begin VB.Form frmScreenshots 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Captura de pantalla de X"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Close 
      Caption         =   "Salir"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   7695
   End
   Begin VB.PictureBox ScreenShot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   120
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   1
      Top             =   120
      Width           =   7680
   End
   Begin VB.PictureBox Capture 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmScreenshots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PicData() As Byte
Private Length As Long

Public Sub AddData(Data As String)

    If Right$(Data, 4) = "ERROR" Then
        MsgBox "Ocurrió un error al mostrar la captura de pantalla.", vbOKOnly, "Captura de pantalla"
        Length = 0
        Exit Sub
    End If

    Call CopyMemory(PicData(Length), ByVal StrPtr(StrConv(Data, vbFromUnicode)), Len(Data))
    
    Length = Length + Len(Data)

End Sub

Public Sub ShowScreenShot(Name As String)

    ReDim Preserve PicData(Length - 1) As Byte

    Dim Path As String
    Path = SaveScreenshotFromBytes(Name, PicData)
    
    ReDim PicData(2097152) As Byte

    Capture.Picture = LoadPicture(Path)

    ScreenShot.PaintPicture Capture.Picture, _
        0, 0, ScreenShot.ScaleWidth, ScreenShot.ScaleHeight, _
        0, 0, 1024, 768, _
        vbSrcCopy

    Length = 0

    Me.Caption = "Captura de pantalla de " & Name

    Me.Show

End Sub

Private Sub Close_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    ReDim PicData(2097152) As Byte
    
End Sub
