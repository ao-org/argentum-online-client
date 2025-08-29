VERSION 5.00
Begin VB.Form frmConnecting 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Connecting"
   ClientHeight    =   2745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleMode       =   0  'User
   ScaleWidth      =   3151.636
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timeout 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   0
      Top             =   0
   End
   Begin VB.Label ConnectionLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Conectando al servidor"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image cmdCancel 
      Height          =   375
      Left            =   600
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmConnecting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCancelButton As clsGraphicalButton
Private RetryCount As Integer
Private TimerProgress As Integer

Private Sub cmdCancel_Click()
    On Error Goto cmdCancel_Click_Err
    Call Disconnect
    Unload Me
    Exit Sub
cmdCancel_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmConnecting.cmdCancel_Click", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hwnd, 240)
    RetryCount = 0
    Me.Picture = LoadInterface("Marco.bmp", False)
    TimerProgress = 0
    Timeout.enabled = True
    Call loadButtons
    Call UpdateConnectionText
    Exit Sub
    
Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCerrar.Form_Load", Erl)
    Resume Next
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmConnecting.Form_Load", Erl)
End Sub

Private Sub loadButtons()
    On Error Goto loadButtons_Err
        
    Set cCancelButton = New clsGraphicalButton
    Call cCancelButton.Initialize(cmdCancel, "boton-cancelar-default.bmp", _
                                                "boton-cancelar-over.bmp", _
                                                "boton-cancelar-off.bmp", Me)
    Exit Sub
loadButtons_Err:
    Call TraceError(Err.Number, Err.Description, "frmConnecting.loadButtons", Erl)
End Sub

Private Sub Timeout_Timer()
    On Error Goto Timeout_Timer_Err
#If DIRECT_PLAY = 0 Then
    Call UpdateConnectionText
    TimerProgress = TimerProgress + 1
    If TimerProgress Mod 20 = 19 Then
        Call RetryWithAnotherIp
    End If
#End If
    Exit Sub
Timeout_Timer_Err:
    Call TraceError(Err.Number, Err.Description, "frmConnecting.Timeout_Timer", Erl)
End Sub

Private Sub UpdateConnectionText()
    On Error Goto UpdateConnectionText_Err
    Dim DisplayText As String
    Dim i As Integer
    DisplayText = "Conectando al servidor"
    For i = 0 To TimerProgress Mod 4
        DisplayText = DisplayText & "."
    Next i
    ConnectionLabel.Caption = DisplayText
    Exit Sub
UpdateConnectionText_Err:
    Call TraceError(Err.Number, Err.Description, "frmConnecting.UpdateConnectionText", Erl)
End Sub
