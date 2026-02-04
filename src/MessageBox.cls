VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''=============================================================================
''' VBA MessageBox for Excel
''' ----------------------------------------------
''' https://github.com/cristianbuse/VBA-MessageBox
''' ----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2022 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

'@PredeclaredId
Option Explicit

Private Const SIDE_VALUE As Single = 18
Private Const ICO_SIZE As Single = 24
Private Const BTN_HEIGHT As Single = 19.5
Private Const BTN_MIN_WIDTH As Single = 54
Private Const SCROLL_WIDTH As Single = 12.75
Private Const MIN_WIDTH As Single = 108 - SIDE_VALUE * 2
Private Const MAX_WIDTH As Single = 600
Private Const MAX_HEIGHT As Single = 288
Private Const MAX_BUTTONS As Long = 5

Public Enum MsgBoxIcon
    icoNone = 0
    icoCritical
    icoExclamation
    icoInformation
    icoQuestion
End Enum

Private WithEvents m_form As MessageForm
Attribute m_form.VB_VarHelpID = -1
Private m_allowClose As Boolean
Private m_result As String

'Controls
Private m_icon As MSForms.Label
Private m_prompt As MSForms.TextBox
Private m_buttonsFrame As MSForms.Frame
Private WithEvents m_escButton As MSForms.CommandButton
Attribute m_escButton.VB_VarHelpID = -1
Private WithEvents m_button1 As MSForms.CommandButton
Attribute m_button1.VB_VarHelpID = -1
Private WithEvents m_button2 As MSForms.CommandButton
Attribute m_button2.VB_VarHelpID = -1
Private WithEvents m_button3 As MSForms.CommandButton
Attribute m_button3.VB_VarHelpID = -1
Private WithEvents m_button4 As MSForms.CommandButton
Attribute m_button4.VB_VarHelpID = -1
Private WithEvents m_button5 As MSForms.CommandButton
Attribute m_button5.VB_VarHelpID = -1
Private m_buttons(1 To MAX_BUTTONS) As MSForms.CommandButton
Private m_visibleButtons As Long

'*******************************************************************************
'Builds the necessary controls and alignment at runtime
'*******************************************************************************
Private Sub BuildForm()
    Const progIDText As String = "Forms.TextBox.1"
    Const progIDLabel As String = "Forms.Label.1"
    Const progIDFrame As String = "Forms.Frame.1"
    Const progIDButton As String = "Forms.CommandButton.1"
    '
    With m_form
        .Font.Name = "Tahoma"
        .Font.Size = 8.25
        .BackColor = vbWhite
        .Width = MIN_WIDTH + SIDE_VALUE * 2
        .Height = BTN_HEIGHT + SIDE_VALUE * 3
        .Width = .Width * 2 - .InsideWidth
        .Height = .Height * 2 - .InsideHeight
    End With
    '
    Set m_escButton = m_form.Controls.Add(progIDButton)
    With CCtrl(m_escButton)
        .Cancel = True 'Allows the form to be closed by pressing the Esc key
        .Default = False
        .Move 0, 0, 0, 0
    End With
    '
    Set m_icon = m_form.Controls.Add(progIDLabel)
    CCtrl(m_icon).Move SIDE_VALUE, SIDE_VALUE, 0, 0
    m_icon.PicturePosition = fmPicturePositionCenter
    m_icon.BackStyle = fmBackStyleTransparent
    '
    Set m_prompt = m_form.Controls.Add(progIDText)
    CCtrl(m_prompt).Move SIDE_VALUE, SIDE_VALUE, MIN_WIDTH, 0
    With m_prompt
        .WordWrap = True
        .MultiLine = True
        .AutoSize = False
        .TextAlign = fmTextAlignLeft
        .SpecialEffect = fmSpecialEffectFlat
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
    End With
    '
    Set m_buttonsFrame = m_form.Controls.Add(progIDFrame)
    CCtrl(m_buttonsFrame).Move 0, SIDE_VALUE * 2 _
                             , m_form.InsideWidth, BTN_HEIGHT + SIDE_VALUE
    m_buttonsFrame.SpecialEffect = fmSpecialEffectFlat
    m_buttonsFrame.BackColor = vbButtonFace
    '
    Dim i As Long
    For i = 1 To MAX_BUTTONS
        Set m_buttons(i) = m_buttonsFrame.Controls.Add(progIDButton)
        CCtrl(m_buttons(i)).Move 0, SIDE_VALUE / 2, 0, BTN_HEIGHT
    Next i
    'Create the required sinks
    Set m_button1 = m_buttons(1)
    Set m_button2 = m_buttons(2)
    Set m_button3 = m_buttons(3)
    Set m_button4 = m_buttons(4)
    Set m_button5 = m_buttons(5)
End Sub
Private Function CCtrl(ByVal c As MSForms.Control) As MSForms.Control
    Set CCtrl = c
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Form/Control events
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub m_form_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then 'User pressed X button
        Cancel = True
        OnCancel
    End If
End Sub
Private Sub m_escButton_Click()
    OnCancel
End Sub
Private Sub OnCancel()
    If Not m_allowClose Then Exit Sub
    m_form.Hide
End Sub
Private Sub m_button1_Click()
    OnClick m_button1
End Sub
Private Sub m_button2_Click()
    OnClick m_button2
End Sub
Private Sub m_button3_Click()
    OnClick m_button3
End Sub
Private Sub m_button4_Click()
    OnClick m_button4
End Sub
Private Sub m_button5_Click()
    OnClick m_button5
End Sub
Private Sub OnClick(ByVal btn As MSForms.CommandButton)
    m_result = btn.Caption
    m_form.Hide
End Sub

'===============================================================================
'Sets the icon shown
'===============================================================================
Private Property Let Icon(ByVal ico As MsgBoxIcon)
    Select Case ico
        Case icoCritical:    Set m_icon.Picture = m_form.icoCritical.Picture
        Case icoExclamation: Set m_icon.Picture = m_form.icoExclamation.Picture
        Case icoInformation: Set m_icon.Picture = m_form.icoInformation.Picture
        Case icoQuestion:    Set m_icon.Picture = m_form.icoQuestion.Picture
        Case Else:           Set m_icon.Picture = Nothing
    End Select
    With CCtrl(m_icon)
        If m_icon.Picture.Handle = 0 Then .Width = 0 Else .Width = ICO_SIZE
        .Height = .Width
    End With
End Property

'===============================================================================
'Sets the button names
'===============================================================================
Private Property Let Buttons(ByRef btnNames As Variant)
    Dim collButtonNames As New Collection
    Dim i As Long, j As Long
    Dim v As Variant
    '
    On Error Resume Next 'Ignore duplicates
    For Each v In btnNames
        If LenB(v) > 0 Then
            collButtonNames.Add v, v
            If StrComp(v, "Cancel", vbTextCompare) = 0 Then m_allowClose = True
        End If
    Next v
    On Error GoTo 0
    If collButtonNames.Count = 0 Then collButtonNames.Add "OK"
    If collButtonNames.Count = 1 Then m_allowClose = True
    '
    i = LBound(m_buttons)
    For Each v In collButtonNames
        m_buttons(i).Caption = v
        CCtrl(m_buttons(i)).Visible = True
        i = i + 1
    Next v
    For j = i To UBound(m_buttons)
        CCtrl(m_buttons(j)).Visible = False
    Next j
    m_visibleButtons = collButtonNames.Count
End Property

'*******************************************************************************
'Utility for positioning
'*******************************************************************************
Private Sub CenterOnApplication()
    Const appMinimizedPos As Long = 10000
    Dim leftPosition As Single
    Dim topPosition As Single
    '
    With Application
        leftPosition = .Left + (.Width - m_form.Width) / 2
        If leftPosition < .Left Then leftPosition = .Left
        '
        topPosition = .Top + (.Height - m_form.Height) / 2
        If topPosition < .Top Then topPosition = .Top
    End With
    If Abs(leftPosition) > appMinimizedPos Then Exit Sub
    If Abs(topPosition) > appMinimizedPos Then Exit Sub
    With m_form
        .StartUpPosition = 0
        .Left = leftPosition
        .Top = topPosition
    End With
End Sub

'*******************************************************************************
'Update the size and position of controls and form itself
'*******************************************************************************
Private Sub Refresh()
    Dim deltaWidth As Single
    Dim deltaHeight As Single
    '
    With CCtrl(m_prompt)
        deltaWidth = .Width + .Left - CCtrl(m_icon).Left - CCtrl(m_icon).Width
        .Left = .Left + .Width - deltaWidth
        .Width = MAX_WIDTH
        .Height = MAX_HEIGHT
        m_prompt.AutoSize = True
        m_prompt.AutoSize = False
        .Width = Min(Max(.Width, MIN_WIDTH), MAX_WIDTH)
        If .Height > MAX_HEIGHT Then
            .Height = MAX_HEIGHT
            .Width = Min(.Width + SCROLL_WIDTH, MAX_WIDTH)
            m_prompt.ScrollBars = fmScrollBarsVertical
        Else
            m_prompt.ScrollBars = fmScrollBarsNone
        End If
        .SetFocus
        m_prompt.SelStart = 0
        deltaWidth = .Width - deltaWidth
        .Top = SIDE_VALUE
        If .Height < CCtrl(m_icon).Height Then
            .Top = .Top + (ICO_SIZE - .Height) / 2
        End If
    End With
    '
    deltaHeight = Max(CCtrl(m_icon).Height, CCtrl(m_prompt).Height) _
                + SIDE_VALUE * 2 - CCtrl(m_buttonsFrame).Top
    With m_form
        .Width = .Width + deltaWidth
        .Height = .Height + deltaHeight
    End With
    With CCtrl(m_buttonsFrame)
        .Width = .Width + deltaWidth
        .Top = .Top + deltaHeight
    End With
    RefreshButtons
    CenterOnApplication
End Sub
Private Function Max(ByVal s1 As Single, ByVal s2 As Single) As Single
    If s1 > s2 Then Max = s1 Else Max = s2
End Function
Private Function Min(ByVal s1 As Single, ByVal s2 As Single) As Single
    If s1 < s2 Then Min = s1 Else Min = s2
End Function
Private Sub RefreshButtons()
    Dim i As Long
    Dim neededWidth As Single:  neededWidth = GetNeededWidth(shrink:=False)
    Dim btnGap As Single:       btnGap = SIDE_VALUE * (m_visibleButtons + 1) / 2
    Dim currentWidth As Single: currentWidth = m_form.InsideWidth - btnGap
    Dim maxWidth As Single:     maxWidth = MAX_WIDTH + CCtrl(m_icon).Width _
                                         + SIDE_VALUE * 2 - btnGap
    Dim posLeft As Single
    
    '
    Do While neededWidth > maxWidth
        neededWidth = GetNeededWidth(shrink:=True)
    Loop
    If neededWidth > currentWidth Then
        Dim deltaWidth As Single: deltaWidth = neededWidth - currentWidth
        CCtrl(m_prompt).Width = CCtrl(m_prompt).Width + deltaWidth
        CCtrl(m_buttonsFrame).Width = CCtrl(m_buttonsFrame).Width + deltaWidth
        m_form.Width = m_form.Width + deltaWidth
    End If
    '
    posLeft = m_form.InsideWidth
    For i = m_visibleButtons To 1 Step -1
        With CCtrl(m_buttons(i))
            posLeft = posLeft - .Width - SIDE_VALUE / 2
            .Left = posLeft
        End With
    Next i
End Sub
Private Function GetNeededWidth(ByVal shrink As Boolean) As Single
    Dim neededWidth As Single
    Dim i As Long
    Dim oldHeight As Single
    Dim newHeight As Single
    Dim btn As MSForms.CommandButton
    '
    For i = 1 To m_visibleButtons
        Set btn = m_buttons(i)
        btn.WordWrap = shrink
        With CCtrl(btn)
            If shrink Then
                oldHeight = .Height
                .Width = Max(.Width * 0.75, BTN_MIN_WIDTH)
            Else
                .Width = MAX_WIDTH
            End If
            btn.AutoSize = True
            btn.AutoSize = False
            .Width = Max(.Width, BTN_MIN_WIDTH)
            neededWidth = neededWidth + .Width
            If shrink Then
                .Height = oldHeight + BTN_HEIGHT / 2
            Else
                .Height = BTN_HEIGHT
            End If
        End With
    Next i
    '
    newHeight = CCtrl(m_buttons(1)).Height + SIDE_VALUE
    With CCtrl(m_buttonsFrame)
        If newHeight <> .Height Then
            oldHeight = .Height
            .Height = newHeight
             m_form.Height = m_form.Height + newHeight - oldHeight
        End If
    End With
    GetNeededWidth = neededWidth
End Function

'*******************************************************************************
'Displays the form in modal mode and returns the text of the pressed button
'*******************************************************************************
'@DefaultMember
Public Function Show(ByVal promptText As String _
                   , Optional ByVal titleText As String _
                   , Optional ByVal ico As MsgBoxIcon = icoNone _
                   , Optional ByVal button1 As String _
                   , Optional ByVal button2 As String _
                   , Optional ByVal button3 As String _
                   , Optional ByVal button4 As String _
                   , Optional ByVal button5 As String _
                   , Optional ByVal defaultButton As Long = 1 _
                   , Optional ByVal allowClose As Boolean = False) As String
Attribute Show.VB_UserMemId = 0
'Attribute Show.VB_UserMemId = 0
    If LenB(titleText) = 0 Then titleText = Application.Name
    If Not IsFormConnected(m_form) Then
        Set m_form = New MessageForm
        BuildForm
    End If
    '
    m_form.Caption = titleText
    Icon = ico
    m_allowClose = allowClose
    Buttons = Array(button1, button2, button3, button4, button5)
    If defaultButton < 1 Then defaultButton = 1
    If defaultButton > m_visibleButtons Then defaultButton = 1
    '
    m_prompt.Value = promptText
    Refresh
    '
    With CCtrl(m_buttons(defaultButton))
        .Default = True
        .SetFocus
    End With
    '
    On Error Resume Next
    CallByName Application, "EnableCancelKey", VbLet, 0
    On Error GoTo 0
    '
    m_result = vbNullString
#If MOUSE_SCROLL_ENABLED Then
    'https://github.com/cristianbuse/VBA-UserForm-MouseScroll
    If m_prompt.ScrollBars <> fmScrollBarsNone Then
        EnableMouseScroll m_form
    End If
#End If
    m_form.Show vbModal
#If MOUSE_SCROLL_ENABLED Then
    If m_prompt.ScrollBars <> fmScrollBarsNone Then
        DisableMouseScroll m_form
    End If
#End If
    Show = m_result
End Function
Private Function IsFormConnected(ByVal obj As Object) As Boolean
    If Not obj Is Nothing Then
        IsFormConnected = TypeName(obj) <> "UserForm"
    End If
End Function
