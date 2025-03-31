VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "DebugTools"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17760
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   17760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame DPLAYParams 
      Caption         =   "DPLAY"
      Height          =   4335
      Left            =   10320
      TabIndex        =   1
      Top             =   360
      Width           =   7095
      Begin VB.CommandButton CmdUpdateDPLAYCaps 
         Caption         =   "UpdateCaps"
         Height          =   495
         Left            =   600
         TabIndex        =   12
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtSystemBufferSize 
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Text            =   "TxtBuffersPerThread"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtNumThreads 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Text            =   "TxtBuffersPerThread"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtDefaultEnumTimeout 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Text            =   "TxtBuffersPerThread"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtDefaultEnumRetryInterval 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Text            =   "TxtBuffersPerThread"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtBuffersPerThread 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Text            =   "TxtBuffersPerThread"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "SystemBufferSizeAs"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "NumThreads"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label label10 
         Caption         =   "DefaultEnumTimeout"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DefaultEnumRetryInterval"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "BuffersPerThread"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox TraceBox 
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmDebug.frx":0000
      Top             =   240
      Width           =   9735
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdUpdateDPLAYCaps_Click()
#If DIRECT_PLAY = 1 Then
    Dim scaps As DPN_SP_CAPS
    
    scaps = dpc.GetSPCaps(DP8SP_TCPIP)
    
     With scaps
        .lBuffersPerThread = CLng(Me.txtBuffersPerThread)
        .lDefaultEnumRetryInterval = CLng(Me.txtDefaultEnumRetryInterval)
        .lDefaultEnumTimeout = CLng(Me.txtDefaultEnumTimeout)
        .lNumThreads = CLng(Me.txtNumThreads)
        .lSystemBufferSize = CLng(Me.txtSystemBufferSize)
        .lFlags = 0
    End With
    dpc.SetSPCaps DP8SP_TCPIP, scaps
#End If
End Sub

Private Sub Form_Load()
Me.TraceBox.Text = vbNullString

End Sub

Public Sub add_text_tracebox(ByVal s As String)
    Debug.Print s
    Me.TraceBox.Text = Me.TraceBox.Text & s & vbCrLf
End Sub
