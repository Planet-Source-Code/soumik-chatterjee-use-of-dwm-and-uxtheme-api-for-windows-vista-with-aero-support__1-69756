VERSION 5.00
Begin VB.UserControl DWM_Label 
   BackColor       =   &H00000000&
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   2  'Use Paint
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   Begin VB.Timer Tm_Blend 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   120
   End
End
Attribute VB_Name = "DWM_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Events and User-defined Types and variables ###################################
'#####################################################################

Event Click()
Event DblClick()
Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
Event MouseLeave()
Event MouseEnter()

Private FA_DWM_Label_Prop_Enabled As Boolean
Private FA_DWM_Label_Prop_Caption As String
Private FA_DWM_Label_Prop_ForeColor As OLE_COLOR
Private FA_DWM_Label_Prop_BackColor As OLE_COLOR
Private FA_DWM_Label_Prop_GlowColor As OLE_COLOR
Private FA_DWM_Label_Prop_HoverColor As OLE_COLOR
Private FA_DWM_Label_Prop_UseHover As Boolean
Private FA_DWM_Label_Prop_GlowSize As Single
Private FA_DWM_Label_Prop_AutoSize As Boolean
Private FA_DWM_Label_Prop_UseBlend As Boolean
Private FA_DWM_Label_Prop_FadeInStep As Single
Private FA_DWM_Label_Prop_FadeOutStep As Single

Private FA_DWM_Label_ThemeTextObj As FA_Type_DWM_ThemeText
Private FA_DWM_Label_IsReady As Boolean
Private FA_DWM_Label_IsMouseIn As Boolean
Private FA_DWM_Label_BlendDone As Long
Private FA_DWM_Label_IsSubClas As Boolean

Private Sub Tm_Blend_Timer()

DoEvents

FA_DWM_Label_BlendDone = FA_DWM_Label_BlendDone + IIf(FA_DWM_Label_IsMouseIn, FA_DWM_Label_Prop_FadeInStep, FA_DWM_Label_Prop_FadeOutStep)

If (FA_DWM_Label_BlendDone >= 255) Then
    Tm_Blend.Enabled = False
    FA_DWM_Label_BlendDone = 255
End If

FA_DWM_Label_RebuildUI

End Sub

Private Sub Usercontrol_Click()
RaiseEvent Click
End Sub

Private Sub Usercontrol_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()

FA_DWM_Label_IsSubClas = False
FA_DWM_Label_IsReady = False
Enabled = True
Caption = "Theme Label"
ForeColor = vbBlack
BackColor = vbBlack
GlowColor = vbWhite
HoverColor = &H764521
UseHover = False
FadeInStep = 25
FadeOutStep = 10
GlowSize = 10
Font = UserControl.Font
AutoSize = True

FA_DWM_Label_IsReady = True
FA_DWM_Label_BlendDone = 255

FA_DWM_Label_RebuildUI

End Sub

Private Sub Usercontrol_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub Usercontrol_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
FA_DWM_Label_Handler_WM_MOUSEHOVER
RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub Usercontrol_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_Paint()
FA_DWM_Label_RebuildUI
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
FA_DWM_Label_Prop_Enabled = PropBag.ReadProperty("Enabled", True)
FA_DWM_Label_Prop_Caption = PropBag.ReadProperty("Caption", "Theme Label")
FA_DWM_Label_Prop_ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
FA_DWM_Label_Prop_BackColor = PropBag.ReadProperty("BackColor ", vbBlack)
FA_DWM_Label_Prop_GlowColor = PropBag.ReadProperty("GlowColor", vbWhite)
FA_DWM_Label_Prop_HoverColor = PropBag.ReadProperty("HoverColor", &H764521)
FA_DWM_Label_Prop_UseHover = PropBag.ReadProperty("UseHover", False)
FA_DWM_Label_Prop_GlowSize = PropBag.ReadProperty("GlowSize", 10)
FA_DWM_Label_Prop_FadeInStep = PropBag.ReadProperty("FadeInStep", 25)
FA_DWM_Label_Prop_FadeOutStep = PropBag.ReadProperty("FadeOutStep", 10)
FA_DWM_Label_Prop_UseBlend = PropBag.ReadProperty("UseBlend", True)
Set UserControl.Font = PropBag.ReadProperty("Font", "Tahoma")
FA_DWM_Label_Prop_AutoSize = PropBag.ReadProperty("AutoSize", True)
End Sub

Private Sub UserControl_Resize()

FA_DWM_Label_RebuildUI

End Sub

Private Sub UserControl_Show()

FA_DWM_Label_IsReady = False
Enabled = FA_DWM_Label_Prop_Enabled
Caption = FA_DWM_Label_Prop_Caption
ForeColor = FA_DWM_Label_Prop_ForeColor
BackColor = FA_DWM_Label_Prop_BackColor
GlowColor = FA_DWM_Label_Prop_GlowColor
HoverColor = FA_DWM_Label_Prop_HoverColor
UseHover = FA_DWM_Label_Prop_UseHover
GlowSize = FA_DWM_Label_Prop_GlowSize
UseBlend = FA_DWM_Label_Prop_UseBlend
FadeInStep = FA_DWM_Label_Prop_FadeInStep
FadeOutStep = FA_DWM_Label_Prop_FadeOutStep
Font = UserControl.Font
FA_DWM_Label_IsReady = True

FA_DWM_Label_RebuildUI

End Sub

Private Sub UserControl_Terminate()
FA_DWM_Label_FreeDC_Src
FA_DWM_Label_FreeDC_Dest
FA_DWM_Label_SubClas_End
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Enabled", FA_DWM_Label_Prop_Enabled, True)
Call PropBag.WriteProperty("Caption", FA_DWM_Label_Prop_Caption, "Theme Label")
Call PropBag.WriteProperty("ForeColor", FA_DWM_Label_Prop_ForeColor, vbBlack)
Call PropBag.WriteProperty("BackColor", FA_DWM_Label_Prop_BackColor, vbBlack)
Call PropBag.WriteProperty("GlowColor", FA_DWM_Label_Prop_GlowColor, vbWhite)
Call PropBag.WriteProperty("HoverColor", FA_DWM_Label_Prop_HoverColor, &H764521)
Call PropBag.WriteProperty("UseHover", FA_DWM_Label_Prop_UseHover, False)
Call PropBag.WriteProperty("GlowSize", FA_DWM_Label_Prop_GlowSize, 10)
Call PropBag.WriteProperty("UseBlend", FA_DWM_Label_Prop_UseBlend, True)
Call PropBag.WriteProperty("FadeInStep", FA_DWM_Label_Prop_FadeInStep, 25)
Call PropBag.WriteProperty("FadeOutStep", FA_DWM_Label_Prop_FadeOutStep, 10)
Call PropBag.WriteProperty("Font", UserControl.Font, "Tahoma")
Call PropBag.WriteProperty("AutoSize", FA_DWM_Label_Prop_AutoSize, True)
End Sub

Private Sub FA_DWM_Label_RebuildUI()

If Not FA_DWM_Label_IsReady Then Exit Sub

Tm_Blend.Enabled = False
If Not FA_DWM_Label_Prop_UseBlend Then FA_DWM_Label_BlendDone = 255

FA_DWM_Label_ThemeTextObj.Caption = Caption
Set FA_DWM_Label_ThemeTextObj.Font = UserControl.Font
FA_DWM_Label_ThemeTextObj.GlowSize = GlowSize
FA_DWM_Label_ThemeTextObj.hWnd = UserControl.hWnd
FA_DWM_Label_ThemeTextObj.IsCustomDC = False
FA_DWM_Label_ThemeTextObj.Width = UserControl.TextWidth(Caption) + (FA_DWM_Label_ThemeTextObj.GlowSize * 2)
FA_DWM_Label_ThemeTextObj.Height = UserControl.TextHeight(Caption) + (FA_DWM_Label_ThemeTextObj.GlowSize * 2)
If AutoSize Then
        UserControl.Width = FA_DWM_Label_ThemeTextObj.Width * Screen.TwipsPerPixelX
        UserControl.Height = FA_DWM_Label_ThemeTextObj.Height * Screen.TwipsPerPixelY
End If
FA_DWM_Label_ThemeTextObj.Left = 0
FA_DWM_Label_ThemeTextObj.Top = (UserControl.ScaleHeight / 2) - (FA_DWM_Label_ThemeTextObj.Height / 2)

FA_DWM_Label_FreeDC_Dest
FA_DWM_Label_FreeDC_Src

Dim ARGBStruct_ForeColor  As FA_Type_ARGB
Dim ARGBStruct_HoverColor  As FA_Type_ARGB
Dim R As Single
Dim G As Single
Dim B As Single

GetARGBVal ForeColor, ARGBStruct_ForeColor
GetARGBVal HoverColor, ARGBStruct_HoverColor

R = ((IIf(FA_DWM_Label_IsMouseIn, 255 - FA_DWM_Label_BlendDone, FA_DWM_Label_BlendDone) / 255) * (ARGBStruct_ForeColor.Red - ARGBStruct_HoverColor.Red)) + ARGBStruct_HoverColor.Red
G = ((IIf(FA_DWM_Label_IsMouseIn, 255 - FA_DWM_Label_BlendDone, FA_DWM_Label_BlendDone) / 255) * (ARGBStruct_ForeColor.Green - ARGBStruct_HoverColor.Green)) + ARGBStruct_HoverColor.Green
B = ((IIf(FA_DWM_Label_IsMouseIn, 255 - FA_DWM_Label_BlendDone, FA_DWM_Label_BlendDone) / 255) * (ARGBStruct_ForeColor.Blue - ARGBStruct_HoverColor.Blue)) + ARGBStruct_HoverColor.Blue

FA_DWM_Label_ThemeTextObj.ForeColor = RGB(R, G, B)
FA_ThemeText_Draw FA_DWM_Label_ThemeTextObj

If FA_DWM_Label_BlendDone < 255 Then Tm_Blend.Enabled = True

End Sub

Public Sub Refresh()
If Not FA_DWM_Label_IsReady Then Exit Sub
FA_ThemeText_Refresh FA_DWM_Label_ThemeTextObj
End Sub

Public Property Let Enabled(ByVal Value As Boolean)
UserControl.Enabled = Value
FA_DWM_Label_Prop_Enabled = Value
PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
Enabled = FA_DWM_Label_Prop_Enabled
End Property

Public Property Let UseHover(ByVal Value As Boolean)
If Not Value Then UseBlend = False
FA_DWM_Label_Prop_UseHover = Value
PropertyChanged "UseHover"
FA_DWM_Label_RebuildUI
End Property

Public Property Get UseHover() As Boolean
UseHover = FA_DWM_Label_Prop_UseHover
End Property

Public Property Let UseBlend(ByVal Value As Boolean)
If Value Then UseHover = True
FA_DWM_Label_Prop_UseBlend = Value
PropertyChanged "UseBlend"
End Property

Public Property Get UseBlend() As Boolean
UseBlend = FA_DWM_Label_Prop_UseBlend
End Property

Public Property Let FadeInStep(ByVal Value As Integer)
FA_DWM_Label_Prop_FadeInStep = Value
PropertyChanged "FadeInStep"
End Property

Public Property Get FadeInStep() As Integer
FadeInStep = FA_DWM_Label_Prop_FadeInStep
End Property

Public Property Let FadeOutStep(ByVal Value As Integer)
FA_DWM_Label_Prop_FadeOutStep = Value
PropertyChanged "FadeOutStep"
End Property

Public Property Get FadeOutStep() As Integer
FadeOutStep = FA_DWM_Label_Prop_FadeOutStep
End Property

Public Property Let AutoSize(ByVal Value As Boolean)
If FA_DWM_Label_Prop_AutoSize = Value Then Exit Property
FA_DWM_Label_Prop_AutoSize = Value
PropertyChanged "AutoSize"
FA_DWM_Label_RebuildUI
End Property

Public Property Get AutoSize() As Boolean
AutoSize = FA_DWM_Label_Prop_AutoSize
End Property

Public Property Let Caption(ByVal Value As String)
If FA_DWM_Label_Prop_Caption = Value Then Exit Property
FA_DWM_Label_Prop_Caption = Value
PropertyChanged "Caption"
FA_DWM_Label_RebuildUI
End Property

Public Property Get Caption() As String
Caption = FA_DWM_Label_Prop_Caption
End Property

Public Property Let Font(Value As StdFont)
If UserControl.Font = Value Then Exit Property
Set UserControl.Font = Value
PropertyChanged "Font"
FA_DWM_Label_RebuildUI
End Property

Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
If FA_DWM_Label_Prop_ForeColor = Value Then Exit Property
FA_DWM_Label_Prop_ForeColor = Value
PropertyChanged "ForeColor"
FA_DWM_Label_RebuildUI
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = FA_DWM_Label_Prop_ForeColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
If FA_DWM_Label_Prop_BackColor = Value Then Exit Property
FA_DWM_Label_Prop_BackColor = Value
UserControl.BackColor = Value
PropertyChanged "BackColor"
Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = FA_DWM_Label_Prop_BackColor
End Property

Public Property Let GlowColor(ByVal Value As OLE_COLOR)
If FA_DWM_Label_Prop_GlowColor = Value Then Exit Property
FA_DWM_Label_Prop_GlowColor = Value
PropertyChanged "GlowColor"
FA_DWM_Label_RebuildUI
End Property

Public Property Get GlowColor() As OLE_COLOR
GlowColor = FA_DWM_Label_Prop_GlowColor
End Property

Public Property Let HoverColor(ByVal Value As OLE_COLOR)
If FA_DWM_Label_Prop_HoverColor = Value Then Exit Property
FA_DWM_Label_Prop_HoverColor = Value
PropertyChanged "HoverColor"
FA_DWM_Label_RebuildUI
End Property

Public Property Get HoverColor() As OLE_COLOR
HoverColor = FA_DWM_Label_Prop_HoverColor
End Property

Public Property Let GlowSize(ByVal Value As Integer)
If FA_DWM_Label_Prop_GlowSize = Value Then Exit Property
FA_DWM_Label_Prop_GlowSize = Value
PropertyChanged "GlowSize"
FA_DWM_Label_RebuildUI
End Property

Public Property Get GlowSize() As Integer
GlowSize = FA_DWM_Label_Prop_GlowSize
End Property

Private Function FA_DWM_Label_SubClas_Start()

If FA_DWM_Label_IsSubClas Then Exit Function
SetProp UserControl.hWnd, "FA_ExWndProcPtr", GetWindowLong(UserControl.hWnd, GWL_WNDPROC)
SetWindowLong UserControl.hWnd, GWL_WNDPROC, AddressOf FA_SubClas_WndProc
SetWindowLong UserControl.hWnd, GWL_USERDATA, ObjPtr(Me)
FA_DWM_Label_IsSubClas = True

End Function

Private Function FA_DWM_Label_SubClas_End()

If Not FA_DWM_Label_IsSubClas Then Exit Function
SetWindowLong UserControl.hWnd, GWL_WNDPROC, GetProp(UserControl.hWnd, "FA_ExWndProcPtr")
RemoveProp UserControl.hWnd, "FA_ExWndProcPtr"
FA_DWM_Label_IsSubClas = False

End Function

Private Function FA_DWM_Label_Handler_WM_MOUSELEAVE()

If Not FA_DWM_Label_IsMouseIn Then Exit Function
FA_DWM_Label_IsMouseIn = False
FA_DWM_Label_BlendDone = 255 - FA_DWM_Label_BlendDone
FA_DWM_Label_RebuildUI
FA_DWM_Label_SubClas_End
RaiseEvent MouseLeave

End Function

Private Function FA_DWM_Label_Handler_WM_MOUSEHOVER()

If FA_DWM_Label_IsMouseIn Then Exit Function
If Not Enabled Then Exit Function
FA_DWM_Label_IsMouseIn = True
FA_DWM_Label_TrackMouse_Start
FA_DWM_Label_BlendDone = 255 - FA_DWM_Label_BlendDone
FA_DWM_Label_RebuildUI
FA_DWM_Label_SubClas_Start
RaiseEvent MouseEnter

End Function

Public Function FA_Handler_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal WParam As Long, ByVal LParam As Long) As Long

If uMsg = WM_MOUSELEAVE Then
        FA_DWM_Label_Handler_WM_MOUSELEAVE
        FA_Handler_WndProc = 1
Else
        FA_Handler_WndProc = 1
End If

End Function

Private Function FA_DWM_Label_TrackMouse_Start()

Dim ET As TRACKMOUSEEVENT
ET.hwndTrack = UserControl.hWnd
ET.dwFlags = TrackMouseEventFlags.TME_LEAVE
ET.cbSize = Len(ET)
TRACKMOUSEEVENT ET

End Function

Private Function FA_DWM_Label_FreeDC_Src()
CloseThemeData FA_DWM_Label_ThemeTextObj.hTheme
SelectObject FA_DWM_Label_ThemeTextObj.hDC_Src, FA_DWM_Label_ThemeTextObj.BMP_Src_Old
SelectObject FA_DWM_Label_ThemeTextObj.hDC_Src, FA_DWM_Label_ThemeTextObj.hFont_Old
DeleteObject FA_DWM_Label_ThemeTextObj.BMP_Src
DeleteObject FA_DWM_Label_ThemeTextObj.hFont
ReleaseDC FA_DWM_Label_ThemeTextObj.hDC_Src, -1
DeleteDC FA_DWM_Label_ThemeTextObj.hDC_Src
End Function

Private Function FA_DWM_Label_FreeDC_Dest()
SelectObject FA_DWM_Label_ThemeTextObj.hDC_Dest, FA_DWM_Label_ThemeTextObj.BMP_Dest_Old
DeleteObject FA_DWM_Label_ThemeTextObj.BMP_Dest
ReleaseDC FA_DWM_Label_ThemeTextObj.hDC_Dest, -1
DeleteDC FA_DWM_Label_ThemeTextObj.hDC_Dest
End Function

