VERSION 5.00
Begin VB.Form Frmcarga 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Frmcarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Picture = LoadInterface("VentanaCargando.bmp")
    MakeFormTransparent Me, vbBlack
End Sub
