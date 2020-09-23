VERSION 5.00
Begin VB.UserControl xeroLED 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   ForeColor       =   &H0000FF00&
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   13
   ToolboxBitmap   =   "xeroLED.ctx":0000
End
Attribute VB_Name = "xeroLED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''
' Made by Robbie Lude aka Absolute Xero '
'                                       '
'            e @ robbie1620@hotmail.com '
'''''''''''''''''''''''''''''''''''''''''

'Default Property Values:
Const m_def_Value = 0
'Property Variables:
Dim m_Value As Long


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
Attribute Value.VB_Description = "The binary code for lighting up sections of the display."
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    PropertyChanged "Value"
    Draw
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Screen.TwipsPerPixelX * 13
    UserControl.Height = Screen.TwipsPerPixelY * 25
End Sub

Private Sub UserControl_Show()
    Draw
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub

Private Sub Draw()
    Dim B(8) As Long
    For i = 0 To 8
        B(i) = RGB(0, ((m_Value And (2 ^ i)) \ (2 ^ i)) * 191 + 64, 0)
    Next
    Cls
    DrawHBar 2, 0, B(0)
    DrawHBar 2, 11, B(1)
    DrawHBar 2, 22, B(2)
    DrawVBar 0, 3, B(3)
    DrawVBar 10, 3, B(4)
    DrawVBar 0, 14, B(5)
    DrawVBar 10, 14, B(6)
    DrawDot 5, 5, B(7)
    DrawDot 5, 16, B(8)
    Picture = Image
End Sub

Private Sub DrawHBar(X, Y, Color)
    Line (X + 1, Y)-(X + 8, Y), Color
    Line (X, Y + 1)-(X + 9, Y + 1), Color
    Line (X + 1, Y + 2)-(X + 8, Y + 2), Color
End Sub

Private Sub DrawVBar(X, Y, Color)
    Line (X, Y + 1)-(X, Y + 7), Color
    Line (X + 1, Y)-(X + 1, Y + 8), Color
    Line (X + 2, Y + 1)-(X + 2, Y + 7), Color
End Sub

Private Sub DrawDot(X, Y, Color)
    Line (X, Y + 1)-(X, Y + 3), Color
    Line (X + 1, Y)-(X + 1, Y + 4), Color
    Line (X + 2, Y + 1)-(X + 2, Y + 3), Color
End Sub

Sub SetDigit(n As Byte)
    Select Case n
        Case 0
            Value = 125
        Case 1
            Value = 80
        Case 2
            Value = 55
        Case 3
            Value = 87
        Case 4
            Value = 90
        Case 5
            Value = 79
        Case 6
            Value = 111
        Case 7
            Value = 81
        Case 8
            Value = 127
        Case 9
            Value = 95
        Case 10
            Value = 384
        Case 11
            Value = 256
        Case 12
            Value = 0
        Case Else
            Randomize Timer
            Value = Rnd * 511
    End Select
End Sub
