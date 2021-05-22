VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const p_SEGUNDOS = 10

Private Const p_CANTIDADDIALOGOS As Byte = 7

Private Type t_GuildDlg

    Texto As String

End Type

Private Segundos As Byte
Private p_Dialogos() As t_GuildDlg
Private offset       As Single
Private p_Activo     As Boolean
Private usedFont     As StdFont
Private Alpha        As Single
Private Const DIALOG_HEIGHT As Integer = 10


''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
    
    On Error GoTo RemoveDialogs_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/07
    'Removes all dialogs
    '***************************************************
    Dim i As Long
    
    For i = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(i).Texto = vbNullString
    Next i

    
    Exit Sub

RemoveDialogs_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.RemoveDialogs", Erl)
    Resume Next
    
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Retrieves the number of dialogs to be displayed on screen
    '***************************************************
    
    On Error GoTo CantidadDialogos_Err
    
    CantidadDialogos = p_CANTIDADDIALOGOS

    
    Exit Property

CantidadDialogos_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.CantidadDialogos", Erl)
    Resume Next
    
End Property

Public Property Let Activo(ByVal v As Boolean)
    
    On Error GoTo Activo_Err
    
    p_Activo = v

    
    Exit Property

Activo_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Activo", Erl)
    Resume Next
    
End Property

Public Property Get Activo() As Boolean
    
    On Error GoTo Activo_Err
    
    Activo = p_Activo

    
    Exit Property

Activo_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Activo", Erl)
    Resume Next
    
End Property

Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Err
    
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True

    
    Exit Sub

Class_Initialize_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Class_Initialize", Erl)
    Resume Next
    
End Sub

Public Sub Draw()
    
    On Error GoTo Draw_Err

    Dim i As Long
    Dim y As Integer
    

    If Not p_Activo Then Exit Sub
    
    Dim Color(3) As RGBA
    Call RGBAList(Color, 222, 194, 112)
    
    Dim dY As Single
    Dim alphaOffset As Integer
    alphaOffset = DIALOG_HEIGHT * (p_CANTIDADDIALOGOS - 3)
    
    For i = 1 To p_CANTIDADDIALOGOS
            
        If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
            dY = (DIALOG_HEIGHT * i - offset - alphaOffset) / alphaOffset
            y = frmMain.renderer.ScaleHeight - 6 - DIALOG_HEIGHT * i + offset
            If dY < 0 Then
                Call Engine_Text_Render_No_Ladder(p_Dialogos(i).Texto, 5, y, Color, 1, False, 0, Alpha)
            Else
                Call Engine_Text_Render_No_Ladder(p_Dialogos(i).Texto, 5, y, Color, 1, False, 0, Alpha * (1 - dY))
            End If
        End If

    Next i
    
    If Segundos <= 0 Then
        If Alpha > 0 Then
            Alpha = Alpha - timerTicksPerFrame * 20
            If Alpha < 0 Then Alpha = 0
        Else
            RemoveDialogs
        End If
    End If
    
    If offset > 0 Then
        offset = offset - timerTicksPerFrame * 2
        If offset <= 0 Then
            offset = 0
        End If
    End If
    
    
    Exit Sub

Draw_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Draw", Erl)
    Resume Next
    
End Sub

Public Sub PassTimer()
    
    On Error GoTo PassTimer_Err
    

    If Not p_Activo Then Exit Sub

    If Segundos > 0 Then
        Segundos = Segundos - 1
    End If
        
    Exit Sub

PassTimer_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.PassTimer", Erl)
    Resume Next
    
End Sub

''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String)
    
    On Error GoTo PushBackText_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08
    'Splits dialogs into lines fitting properly the render area and inserts them
    '***************************************************
    Dim str As String

    Dim tmp As Integer
    
    str = s
    
     'If it's too long to fit, split it
    Do While Engine_Text_Width(str) > frmMain.renderer.ScaleWidth - 20
        tmp = InStrRev(str, " ")
    
        If tmp = 0 Then Exit Do
    
        str = Left$(str, tmp - 1)
    Loop
    
    'Show message and continue with the rest
    Call RealPushBackText(str)
    
    If LenB(str) <> LenB(s) Then
        Call PushBackText(Right$(s, Len(s) - tmp))

    End If

    
    Exit Sub

PushBackText_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.PushBackText", Erl)
    Resume Next
    
End Sub

Private Sub RealPushBackText(ByVal s As String)
    
    On Error GoTo RealPushBackText_Err

    Dim i     As Integer

    For i = (p_CANTIDADDIALOGOS - 1) To 1 Step -1
        p_Dialogos(i + 1) = p_Dialogos(i)
    Next i
    
    p_Dialogos(1).Texto = s
    Segundos = p_SEGUNDOS
    offset = DIALOG_HEIGHT
    Alpha = 255
    Exit Sub

RealPushBackText_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.RealPushBackText", Erl)
    Resume Next
    
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Retrieves the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo font_Err
    
    font = usedFont

    
    Exit Property

font_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.font", Erl)
    Resume Next
    
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let font(ByRef nValue As StdFont)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Sets the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo font_Err
    
    Set usedFont = nValue

    
    Exit Property

font_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.font", Erl)
    Resume Next
    
End Property

