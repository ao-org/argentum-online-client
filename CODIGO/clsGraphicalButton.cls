VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGraphicalButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'*****************************************************
'Graphical Button
'Designed and implemented by ZaMa
'zaxtor@gmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
Option Explicit

Private WithEvents ButtonHolder As Image
Attribute ButtonHolder.VB_VarHelpID = -1

Private RollOverButtonImage As Picture  'img holder
Private NormalButtonImage As Picture    'img holder
Private ClickButtonImage As Picture  'img holder

Private Form_Movement As clsFormMovementManager

Private OverButton As Boolean    'Indicates wether the mouse is over the button or not

Public Sub Initialize(ByRef imgButton As Image, ByVal imgNormalPath As String, ByVal imgRollOverPath As String, ByVal imgClickPath As String, Optional ByRef Form As Form)
'***************************************************
'Author: ZaMa
'Last Modify Date: 22/12/09
'Sets the reference to the image container
'***************************************************
    Set ButtonHolder = imgButton
    Set Form_Movement = New clsFormMovementManager
    Call Form_Movement.Initialize(Form)
    Set RollOverButtonImage = LoadInterface(imgRollOverPath)
    Set NormalButtonImage = LoadInterface(imgNormalPath)
    Set ClickButtonImage = LoadInterface(imgClickPath)

    With ButtonHolder
            .Picture = NormalButtonImage
    End With

End Sub


Private Sub ButtonHolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse down event of the image button.
'Toggle to Click image button.
'***************************************************

    ButtonHolder.Picture = ClickButtonImage
    Call Sound.Sound_Play(SND_CLICK)
End Sub

Private Sub ButtonHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 22/12/09
'Implements the mouse move event of the image button.
'Toggle to pressed image button.
'***************************************************

    If Not OverButton Then

        'This helps when two buttons are close
        Form_Movement.LastButtonPressed.ToggleToNormal

        ButtonHolder.Picture = RollOverButtonImage
        OverButton = True

        Set Form_Movement.LastButtonPressed = Me
        Call Sound.Sound_Play(SND_CLICK_OVER)
    End If
End Sub

Public Sub ToggleToNormal()
'***************************************************
'Author: ZaMa
'Last Modify Date: 22/12/09
'Toggle to normal image button.
'***************************************************
    If Not OverButton Then Exit Sub

    ButtonHolder.Picture = NormalButtonImage
    OverButton = False
End Sub

Private Sub ButtonHolder_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse up event of the image button.
'Toggle to RollOver image button.
'***************************************************
    ButtonHolder.Picture = RollOverButtonImage
End Sub


