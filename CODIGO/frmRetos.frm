VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retos"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   Icon            =   "frmRetos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtPPT 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Text            =   "1"
      Top             =   105
      Width           =   615
   End
   Begin VB.TextBox Apuesta 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Text            =   "20000"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Retar 
      Caption         =   "Retar"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblPPT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jugadores por Equipo: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Apuesta"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Error 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Equipo 2"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Equipo 1"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Apuesta_Change()
    Dim Sel As Integer
    Sel = Apuesta.SelStart

    Dim NewVal As Long
    NewVal = CLng(Abs(Val(Apuesta.Text)))
    
    If NewVal > 100000000 Then
        NewVal = 10000000
    End If

    Apuesta.Text = NewVal
    Apuesta.SelStart = Sel
End Sub

Private Sub Form_Load()
    Jugador(0) = UserName
End Sub

Private Sub Jugador_Change(Index As Integer)
    Error.Caption = vbNullString
End Sub

Private Sub Retar_Click()

    If Not Validar Then Exit Sub
    
    Dim Players As String, i As Integer
    
    ' No incluímos el jugador que crea el reto
    For i = 1 To txtPPT.Text * 2 - 1
        Players = Players & Jugador(i).Text & ";"
    Next
        
    Players = Left$(Players, Len(Players) - 1)
    
    Call WriteDuel(Players, CLng(Apuesta.Text))
        
    Unload Me

End Sub

Private Function Validar() As Boolean
    Dim ErrorStr As String

    Dim i        As Integer

    For i = 0 To txtPPT.Text * 2 - 1

        If LenB(Jugador(i).Text) = 0 Then
            Error.Caption = "Complete todos los jugadores."
            Exit Function
            
        ElseIf Not ValidarNombre(Jugador(i).Text, ErrorStr) Then
            Error.Caption = "Nombre inválido: """ & Jugador(i).Text & """"
            Exit Function

        End If

    Next
    
    Dim j As Integer

    For i = 0 To txtPPT.Text * 2 - 2
        For j = i + 1 To txtPPT.Text * 2 - 1

            If Jugador(i).Text = Jugador(j).Text Then
                Error.Caption = "¡No puede haber jugadores repetidos!"
                Exit Function

            End If

        Next
    Next
    
    Error.Caption = vbNullString
    Validar = True

End Function

Private Sub txtPPT_Change()
    
    On Error GoTo ErrorHandler:

    If Not IsNumeric(txtPPT.Text) Or txtPPT.Text <= 0 Then
        txtPPT.Text = 1
    End If
        
    If txtPPT.Text > 5 Then txtPPT.Text = 5
        
    Dim CantidadJugadores As Byte
        CantidadJugadores = txtPPT.Text * 2 - 1
        
    Dim i As Byte
    For i = 2 To max(CantidadJugadores, Jugador.UBound)
            
        If i < CantidadJugadores + 1 Then
            
            ' Creamos un nuevo TextBox
            Call Load(Jugador(i))
            
            ' Le asignamos sus propiedades
            With Jugador(i)
    
                .Visible = True
                .Enabled = True
                .Text = vbNullString
                .Width = Jugador(0).Width
                .Height = Jugador(0).Height
                .BackColor = vbWhite
                    
                If (i Mod 2) = 1 Then
                
                    ' Derecha
                    .Left = 1920
                    .Top = Jugador(i - 2).Top + 360
                    
                Else
                
                    ' Izquierda
                    .Left = 240
                    .Top = Jugador(i - 1).Top + 360

                End If
                    
            End With
               
        Else
                  
            ' Destruimos el TextBox
            Call Unload(Jugador(i))
               
        End If

    Next
        
    ' Reordenamos los elementos que estan por debajo del nombre de los jugadores
    Label3.Top = Jugador(Jugador.UBound).Top + 700
    Apuesta.Top = Jugador(Jugador.UBound).Top + 700
    Retar.Top = Jugador(Jugador.UBound).Top + 1180
    Error.Top = Jugador(Jugador.UBound).Top + 1900
    Me.Height = Jugador(Jugador.UBound).Top + 2700
    
ErrorHandler:
    
    ' Si el TextBox ya existe, nos saltamos el Load()
    If Err.number = 360 Then Resume Next
    
End Sub
