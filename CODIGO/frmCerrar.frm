VERSION 5.00
Begin VB.Form frmCerrar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Opcion 
      Height          =   420
      Index           =   2
      Left            =   630
      Top             =   1750
      Width           =   1980
   End
   Begin VB.Image Opcion 
      Height          =   420
      Index           =   1
      Left            =   640
      Top             =   1180
      Width           =   1980
   End
   Begin VB.Image Opcion 
      Height          =   420
      Index           =   0
      Left            =   640
      Top             =   610
      Width           =   1980
   End
End
Attribute VB_Name = "frmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
 
Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

Private RealizoCambios As String

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hWnd, 220)
    Me.Picture = LoadInterface("desconectar.bmp")

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Opcion(0).Tag = "0"
    Opcion(0).Picture = Nothing
    Opcion(1).Tag = "0"
    Opcion(1).Picture = Nothing
    Opcion(2).Tag = "0"
    Opcion(2).Picture = Nothing

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Opcion_Click(Index As Integer)
    
    On Error GoTo Opcion_Click_Err
    

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal
            Call WriteQuit
            Unload Me

        Case 1  'Cerrar juego
            Call CloseClient

        Case 2 'Cancelar
            Unload Me

    End Select

    
    Exit Sub

Opcion_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Opcion_Click", Erl)
    Resume Next
    
End Sub

Private Sub Opcion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Opcion_MouseDown_Err
    

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal
            Opcion(Index).Picture = LoadInterface("boton-menu-principal-ES-off.bmp")
            Opcion(Index).Tag = "1"

        Case 1  'Cerrar juego
            Opcion(Index).Picture = LoadInterface("boton-salir-ES-off.bmp")
            Opcion(Index).Tag = "1"

        Case 2 'Cancelar
            Opcion(Index).Picture = LoadInterface("boton-cancelar-ES-off.bmp")
            Opcion(Index).Tag = "1"

    End Select

    
    Exit Sub

Opcion_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Opcion_MouseDown", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Opcion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Opcion_MouseMove_Err
    

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadInterface("boton-menu-principal-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

        Case 1  'Cerrar juego

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadInterface("boton-salir-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

        Case 2 'Cancelar

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadInterface("boton-cancelar-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

    End Select

    
    Exit Sub

Opcion_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Opcion_MouseMove", Erl)
    Resume Next
    
End Sub
