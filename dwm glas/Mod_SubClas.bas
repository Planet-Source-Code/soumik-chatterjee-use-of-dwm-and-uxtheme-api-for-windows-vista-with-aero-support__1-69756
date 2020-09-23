Attribute VB_Name = "Mod_SubClass"
Option Explicit

Public Function FA_SubClas_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal WParam As Long, ByVal LParam As Long) As Long

On Error Resume Next

Dim FA_OBJ As Object
Dim FA_OBJ_PTR As Long
   
FA_OBJ_PTR = GetWindowLong(hWnd, GWL_USERDATA)
RtlMoveMemory FA_OBJ, FA_OBJ_PTR, 4

If (FA_OBJ.FA_Handler_WndProc(hWnd, uMsg, WParam, LParam) = 0) Then
        FA_SubClas_WndProc = 0
        RtlMoveMemory FA_OBJ, 0&, 4
        Set FA_OBJ = Nothing
        Exit Function
End If

FA_SubClas_WndProc = CallWindowProc(GetProp(hWnd, "FA_ExWndProcPtr"), hWnd, uMsg, WParam, LParam)

RtlMoveMemory FA_OBJ, 0&, 4
Set FA_OBJ = Nothing

End Function


