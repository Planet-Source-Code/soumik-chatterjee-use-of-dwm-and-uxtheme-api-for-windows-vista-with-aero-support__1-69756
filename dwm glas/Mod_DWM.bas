Attribute VB_Name = "Mod_DWM"
Option Explicit

Public Function FA_DWM_Init_GlasFrame(ByVal hWnd As Long, Pos_Top As Integer, Pos_Bottom As Integer, Pos_Left As Integer, Pos_Right As Integer)

Dim Margins As TRect
Margins.M_Buttom = Pos_Bottom
Margins.M_Left = Pos_Left
Margins.M_Right = Pos_Right
Margins.M_Top = Pos_Top
DwmExtendFrameIntoClientArea hWnd, Margins

End Function

Public Function FA_DWM_Init_BlurBehind(ByVal hWnd As Long)

Dim BlurFlag As DWM_BlurBehind

BlurFlag.dwFlags = 1
BlurFlag.fEnable = True
BlurFlag.RGNBlur = vbNull
BlurFlag.tMAX = False
DwmEnableBlurBehindWindow hWnd, BlurFlag

End Function

Public Function FA_DWM_CompositionColor(TheForm As Form, R As Integer, G As Integer, B As Integer)

TheForm.BackColor = RGB(R, G, B)

End Function
