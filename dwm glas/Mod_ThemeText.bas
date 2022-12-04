Attribute VB_Name = "Mod_ThemeText"
Public Type FA_Type_DWM_ThemeText
        Caption As String
        ForeColor As Long
        BackColor As Long
        GlowSize As Integer
        GlowColor As Long
        Font As StdFont
        
        Top As Integer
        Left As Integer
        Width As Integer
        Height As Integer
        
        hWnd As Long
        hFont As Long
        hFont_Old As Long
        hTheme As Long
        hDC_Dest As Long
        hDC_Src As Long
        BMP_Src As Long
        BMP_Src_Old As Long
        BMP_Dest As Long
        BMP_Dest_Old As Long
        IsCustomDC As Boolean
End Type

Option Explicit

Public Function FA_ThemeText_Draw(Obj As FA_Type_DWM_ThemeText)

Dim AreaRect As RECT
Dim DTT_Opts As DTTOPTS
Dim DIB As BITMAPINFO

Obj.hTheme = OpenThemeData(Obj.hWnd, StrPtr("Window"))
If (Not Obj.IsCustomDC) Then Obj.hDC_Dest = GetDC(Obj.hWnd)
Obj.hDC_Src = CreateCompatibleDC(Obj.hDC_Dest)
   
With AreaRect
        .Left = Obj.GlowSize
        .Top = 0
        .Right = Obj.Width
        .Bottom = Obj.Height
End With
   
With DIB.bmiHeader
        .biSize = Len(DIB)
        .biWidth = Obj.Width
        .biHeight = -Obj.Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
End With

If (SaveDC(Obj.hDC_Src) <> 0) And (SaveDC(Obj.hDC_Dest) <> 0) Then
        Obj.BMP_Src = CreateDIBSection(Obj.hDC_Src, DIB, 0, 0, 0, 0)
        Obj.BMP_Dest = CreateDIBSection(Obj.hDC_Dest, DIB, 0, 0, 0, 0)
        If (Obj.BMP_Src <> 0) And (Obj.BMP_Dest <> 0) Then
                Obj.BMP_Src_Old = SelectObject(Obj.hDC_Src, Obj.BMP_Src)
                Obj.BMP_Dest_Old = SelectObject(Obj.hDC_Dest, Obj.BMP_Dest)
                Obj.hFont = FA_ThemeText_hFont_Get(Obj.Font, Obj.hDC_Src)
                Obj.hFont_Old = SelectObject(Obj.hDC_Src, Obj.hFont)
                With DTT_Opts
                        .crText = Obj.ForeColor
                        .dwSize = Len(DTT_Opts)
                        .dwFlags = DTT_COMPOSITED Or DTT_GLOWSIZE Or DTT_TEXTCOLOR
                        .iGlowSize = Obj.GlowSize
                End With
                DrawThemeTextEx Obj.hTheme, Obj.hDC_Src, 0, 0, StrPtr(Obj.Caption), -1, DT_TEXTFORMAT, AreaRect, DTT_Opts
                BitBlt Obj.hDC_Dest, Obj.Left, Obj.Top, Obj.Width, Obj.Height, Obj.hDC_Src, 0, 0, vbSrcCopy
        End If
End If

End Function

Private Function FA_ThemeText_hFont_Get(ByRef TheFont As StdFont, ByVal hDC As Long) As Long

Dim TheLF As LOGFONT
FA_ThemeText_OLEFontToLogFont TheFont, hDC, TheLF
FA_ThemeText_hFont_Get = CreateFontIndirect(TheLF)

End Function

Private Sub FA_ThemeText_OLEFontToLogFont(ByRef ThisFont As StdFont, ByVal hDC As Long, ByRef TheLF As LOGFONT)

Dim sFont As String
Dim iChar As Integer
Dim ByteArray() As Byte

With TheLF
     
     sFont = ThisFont.Name
     ByteArray = StrConv(sFont, vbFromUnicode)
     
     For iChar = 1 To Len(sFont)
        .lfFaceName(iChar - 1) = ByteArray(iChar - 1)
     Next iChar
     
     .lfHeight = -MulDiv((ThisFont.size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
     .lfItalic = ThisFont.Italic
     
     If (ThisFont.Bold) Then
       .lfWeight = FW_BOLD
     Else
       .lfWeight = FW_NORMAL
     End If
     
     .lfUnderline = ThisFont.Underline
     .lfStrikeOut = ThisFont.Strikethrough
     .lfCharSet = ThisFont.Charset

End With

End Sub

Public Sub FA_ThemeText_Refresh(Obj As FA_Type_DWM_ThemeText)

BitBlt Obj.hDC_Dest, Obj.Left, Obj.Top, Obj.Width, Obj.Height, Obj.hDC_Src, 0, 0, vbSrcCopy

End Sub
