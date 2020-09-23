VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   467
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   2  'CenterScreen
   Begin DWMEffects.DWM_Label DWM_Label2 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   6480
      Width           =   5715
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This can be used to make a  list/playlist activex control."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      FadeInStep      =   10
      FadeOutStep     =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label1 
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6120
      Width           =   5310
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "Move the cursor over the labels to see hover effect."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      FadeInStep      =   10
      FadeOutStep     =   25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   10
      Left            =   360
      TabIndex        =   10
      Top             =   4920
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   11
      Left            =   360
      TabIndex        =   11
      Top             =   5280
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   4200
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   13
      Left            =   360
      TabIndex        =   13
      Top             =   4560
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
   Begin DWMEffects.DWM_Label DWM_Label 
      Height          =   375
      Index           =   14
      Left            =   360
      TabIndex        =   14
      Top             =   3840
      Width           =   3300
      _ExtentX        =   2752
      _ExtentY        =   873
      Caption         =   "This text is drawn using vista aero theme."
      HoverColor      =   16777215
      UseHover        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   0   'False
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'sum parts still do not work as they r incomplete  - like the png loader as a background picture for the label...
' am workin on it.

Private IsGlasFrame As Boolean

Public Function FA_Handler_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal WParam As Long, ByVal LParam As Long) As Long

' If you uncomment the if-else below, then the window frame will be removed and u can draw in non-client area
' but as  side-effect u cannot resize the form and also there will be no form caption and closebutton wont work by itself
' to correct this , add defwndproc function to the subclasing function . for more info search msdn.com or ask me !

'If uMsg = WM_NCCALCSIZE Then
'        FA_Handler_WndProc = 0
'        Exit Function
'End If

If uMsg = WM_ACTIVATE Then
        If (Not IsGlasFrame) Then
                FA_DWM_Init_GlasFrame hWnd, -1, -1, -1, -1
                FA_Handler_WndProc = 0
                IsGlasFrame = True
                Exit Function
        End If
End If

FA_Handler_WndProc = 1

End Function

Private Function FA_Form_SubClas_Start()

SetProp Me.hWnd, "FA_ExWndProcPtr", GetWindowLong(Me.hWnd, GWL_WNDPROC)
SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FA_SubClas_WndProc
SetWindowLong Me.hWnd, GWL_USERDATA, ObjPtr(Me)

End Function

Private Function FA_Form_SubClas_End()

SetWindowLong Me.hWnd, GWL_WNDPROC, GetProp(Me.hWnd, "FA_ExWndProcPtr")
RemoveProp Me.hWnd, "FA_ExWndProcPtr"

End Function

Private Sub Form_Initialize()
IsGlasFrame = False
FA_Form_SubClas_Start
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FA_Form_SubClas_End
End Sub
