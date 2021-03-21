VERSION 5.00
Begin VB.Form Frmcarga 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cargando..."
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
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadInterface(Language + "\VentanaCargando.bmp")
    MakeFormTransparent Me, vbBlack

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "Frmcarga.Form_Load", Erl)
    Resume Next
    
End Sub
