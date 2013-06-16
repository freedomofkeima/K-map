VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "K-Map 4 Variables"
   ClientHeight    =   6000
   ClientLeft      =   17805
   ClientTop       =   2700
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "K-map.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   11400
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   120
      Top             =   120
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   8160
      TabIndex        =   39
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Generate!"
      Height          =   495
      Left            =   6000
      TabIndex        =   37
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok!"
      Height          =   375
      Left            =   6720
      TabIndex        =   36
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok!"
      Height          =   375
      Left            =   6720
      TabIndex        =   35
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6000
      TabIndex        =   34
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   6000
      TabIndex        =   32
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   3840
      TabIndex        =   15
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   3840
      TabIndex        =   14
      Text            =   "00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3840
      TabIndex        =   13
      Text            =   "00"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3840
      TabIndex        =   12
      Text            =   "00"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   4920
      TabIndex        =   11
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   4920
      TabIndex        =   10
      Text            =   "00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   4920
      TabIndex        =   9
      Text            =   "00"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4920
      TabIndex        =   8
      Text            =   "00"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2760
      TabIndex        =   7
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2760
      TabIndex        =   6
      Text            =   "00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Text            =   "00"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   4
      Text            =   "00"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Text            =   "00"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Text            =   "00"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Text            =   "00"
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Text            =   "00"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   5640
      Width           =   11415
      Begin VB.CommandButton Command4 
         Caption         =   "&End"
         Height          =   375
         Left            =   10320
         TabIndex        =   43
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Constructed by freedomofkeima 2012"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   40
      Top             =   5520
      Width           =   11415
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Function :"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   38
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "For Don't Care :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   33
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "For True :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   31
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Table Conditions:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "2 = don't care"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "1 = true"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "K- map 4 variables (SOP Representation) "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   25
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   24
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   720
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   1440
      Y1              =   720
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "CD"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "AB"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "0 / null = false"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0FF&
      FillColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   0
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Height          =   5655
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j, k, l As Integer
Dim minimum, counter As Integer
Dim TempKarakterMin As String
Dim idxtwo, Penyimpan(15) As Integer
Dim b As Boolean
Dim IsTrue(80) As Boolean
Dim Masukkan(15) As Integer
Dim Converter(15) As String
Dim g_nTransparency As Integer
'Dont Edit below
Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
      ByVal x As Long, ByVal y As Long, _
      ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Private Declare Function GetWindowRect Lib "user32" ( _
      ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'--

Public Function Karakter(ByVal Source As String, ByVal Indeks As Integer) As String 'untuk mengakses char
    Karakter = Mid$(Source$, Indeks, 1)
End Function

Public Function StripOut(ByVal Source As String, ByVal Discard As String) As String
 ' Fungsi untuk menghilangkan setiap char dalam "Discard" dari "Source"
    StripOut = Source
    For i = 1 To Len(Discard)
       StripOut = Replace(StripOut, Mid$(Discard, i, 1), "")
    Next i
End Function

Public Function ConvertTriary(ByVal Source As Integer) As String
Dim TempInt As Integer
Dim TempS As String
    TempS = ""
    TempInt = Source
    Do While (TempInt <> 0)
        TempS = Right(Str(TempInt Mod 3), 1) + TempS
        TempInt = TempInt \ 3
    Loop
    Do While (Len(TempS) <> 4)
        TempS = "0" + TempS
    Loop
    ConvertTriary = TempS
End Function

Public Function ConvertBinary(ByVal Source As Integer, ByVal IndeksPass As Integer) As String
Dim TempInt As Integer
Dim TempS As String
    TempS = ""
    TempInt = Source
    Do While (TempInt <> 0)
        TempS = Right(Str(TempInt Mod 2), 1) + TempS
        TempInt = TempInt \ 2
    Loop
    Do While (Len(TempS) <> IndeksPass)
        TempS = "0" + TempS
    Loop
    ConvertBinary = TempS
End Function

Public Function ConvertDecimal(ByVal Source As String) As Integer
Dim TempInt, x  As Integer
    TempInt = 0
    x = 27
    For i = 1 To 4
        TempInt = Val(Karakter(Source, i)) * x + TempInt
        x = x \ 3
    Next i
    ConvertDecimal = TempInt
End Function

Sub TrueNilai()
Dim TempAngka As Integer
    TempAngka = Val(Text2)
    If (TempAngka >= 0) And (TempAngka <= 15) Then
        Masukkan(TempAngka) = 1
        Text1(TempAngka) = 1
    End If
    Text2 = ""
End Sub

Sub CareNilai()
Dim TempAngka As Integer
    TempAngka = Val(Text3)
    If (TempAngka >= 0) And (TempAngka <= 15) Then
        Masukkan(TempAngka) = 2
        Text1(TempAngka) = 2
    End If
        Text3 = ""
End Sub

Sub CekJawaban()
Dim IsPossible, IsUjung As Boolean
Dim TempKarakter As String
Dim TempKarakterTwo As String
Dim TempKarakterThree As String
Dim idx, bypass As Integer
Dim OwnMap(15), IsNeededCheck As Boolean
    IsNeededCheck = False
    For i = 0 To 15
        Converter(i) = ConvertBinary(i, 4)
        If Masukkan(i) = 1 Then IsNeededCheck = True
    Next i
    If IsNeededCheck Then
        For i = 0 To 80 'I.S. Semua IsTrue terdefinisi true
        IsPossible = True 'Inisialisasi awal
       'Keluarkan char ke - j
              For j = 0 To 15
               IsUjung = True
                    For k = 0 To 3
                      TempKarakter = Karakter(ConvertTriary(i), k + 1)
                      TempKarakterTwo = Karakter(ConvertBinary(j, 4), k + 1)
                      If TempKarakter = 0 And TempKarakterTwo = 1 Then IsUjung = False
                      If TempKarakter = 1 And TempKarakterTwo = 0 Then IsUjung = False
                    Next k
                If IsUjung = True And Masukkan(j) = 0 Then IsPossible = False
               Next j
          If IsPossible Then IsTrue(i) = True
         Next i  'Lakukan pengecekan, matikan yang bukan prime implicant
      For i = 0 To 80
          For j = 0 To 80
             If (i <> j) And IsTrue(i) And IsTrue(j) Then
             IsUjung = True
                For k = 0 To 3
                TempKarakter = Karakter(ConvertTriary(i), k + 1)
                TempKarakterTwo = Karakter(ConvertTriary(j), k + 1)
                If (TempKarakter <> TempKarakterTwo) And (TempKarakter <> "2") Then IsUjung = False
                Next k
              If IsUjung Then IsTrue(j) = False
             End If
        Next j
       Next i 'The next checking step, we'll use bruteforce approach
    idxtwo = 0
    For i = 0 To 80
     If IsTrue(i) Then
         idxtwo = idxtwo + 1
         Penyimpan(idxtwo) = i
     End If
    Next i
    If idxtwo > 1 Then
       bypass = 1
         For i = 1 To idxtwo
          bypass = bypass * 2 'Note that maximum = 256
         Next i
       minimum = -1
       TempKarakterMin = ""
         For i = 1 To (bypass - 1) 'Mencari kemungkinan logika terkecil yang mungkin
         counter = 0
         TempKarakter = ConvertBinary(i, idxtwo) 'Setiap logika merupakan bagian dari map
            For j = 0 To 15
               OwnMap(j) = False 'Inisialisasi
           Next j
          For j = 1 To idxtwo   'Create logics own map
           If Karakter(TempKarakter, j) = "1" Then
              counter = counter + 1     'Tentukan logika yang disimpan logika ke - Penyimpan(j) tersebut
            For k = 0 To 15     'Time to check for each logic in own map
                IsUjung = True
                For l = 0 To 3 'Check each character respectively
                    TempKarakterTwo = Karakter(ConvertTriary(Penyimpan(j)), l + 1)
                    TempKarakterThree = Karakter(ConvertBinary(k, 4), l + 1)
                    If TempKarakterTwo = 0 And TempKarakterThree = 1 Then IsUjung = False
                    If TempKarakterTwo = 1 And TempKarakterThree = 0 Then IsUjung = False
                Next l
                If IsUjung Then OwnMap(k) = True
            Next k
         End If
        Next j
         IsPossible = True
         For j = 0 To 15
          If Masukkan(j) = 1 And OwnMap(j) = False Then IsPossible = False 'mencocokkan own map dengan masukkan user, jika sampai ujung, then break, print (don't care diabaikan)
         Next j
        If IsPossible Then
         If minimum = -1 Or (counter < minimum) Then
              TempKarakterMin = TempKarakter
              minimum = counter
         End If
        End If
     Next i
        For j = 1 To idxtwo
             If Karakter(TempKarakterMin, j) = "0" Then IsTrue(Penyimpan(j)) = False
        Next j
      End If
    End If
    idx = 0
    For i = 0 To 79
     If IsTrue(i) Then
          idx = idx + 1
          TempKarakter = ConvertTriary(i)
          TempKarakterTwo = ""
          If Karakter(TempKarakter, 1) = "1" Then TempKarakterTwo = TempKarakterTwo + "A "
          If Karakter(TempKarakter, 1) = "0" Then TempKarakterTwo = TempKarakterTwo + "A! "
          If Karakter(TempKarakter, 2) = "1" Then TempKarakterTwo = TempKarakterTwo + "B "
          If Karakter(TempKarakter, 2) = "0" Then TempKarakterTwo = TempKarakterTwo + "B! "
          If Karakter(TempKarakter, 3) = "1" Then TempKarakterTwo = TempKarakterTwo + "C "
          If Karakter(TempKarakter, 3) = "0" Then TempKarakterTwo = TempKarakterTwo + "C! "
          If Karakter(TempKarakter, 4) = "1" Then TempKarakterTwo = TempKarakterTwo + "D "
          If Karakter(TempKarakter, 4) = "0" Then TempKarakterTwo = TempKarakterTwo + "D! "
          List1.AddItem TempKarakterTwo 'Test Cetak
     End If
    Next i
    If IsTrue(80) Then
     idx = idx + 1
     List1.AddItem "1"
    End If
    If idx = 0 Then List1.AddItem "0"
End Sub

Private Sub Command3_Click()
List1.Clear
If Command3.Caption = "Generate!" Then
    Text2.Enabled = False
    Text3.Enabled = False
    For i = 0 To 15
        Masukkan(i) = Val(Text1(i))
        If (Text1(i) <> "1") And (Text1(i) <> "2") Then
            Text1(i) = "0"
        End If
       Text1(i).Enabled = False
       Command3.Caption = "Re-generate!"
    Next i
    CekJawaban
Else
    Text2.Enabled = True
    Text3.Enabled = True
    For i = 0 To 15
        Text1(i).Enabled = True
        Text1(i) = ""
    Next i
    For i = 0 To 80
        IsTrue(i) = False 'inisialisasi salah, untuk setiap masukkan bernilai 0, buat yg 1 salah dsb.
    Next i
    Command3.Caption = "Generate!"
End If
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame1.Visible = False
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame1.Visible = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
' enter-press -13
If KeyAscii = 13 Then
    TrueNilai
End If
End Sub

Private Sub Command1_Click()
    TrueNilai
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(LCase(Chr(KeyAscii)))
' enter-press -13
If KeyAscii = 13 Then
    CareNilai
End If
End Sub

Private Sub Command2_Click()
CareNilai
End Sub

Private Sub Form_Load()
'Inisialisasi
For i = 0 To 15
    Masukkan(i) = 0
    Text1(i).Text = ""
    Penyimpan(i) = -1
Next i
For i = 0 To 80
    IsTrue(i) = False 'inisialisasi salah, untuk setiap masukkan bernilai 0, buat yg 1 salah dsb.
Next i
Frame1.Visible = False
End Sub

Private Sub Timer1_Timer()
If Form1.Left > 3555 Then
    Form1.Left = Form1.Left - 1000
End If
On Error GoTo ErrorRtn
    g_nTransparency = g_nTransparency + 3
    If g_nTransparency > 255 Then
        g_nTransparency = g_nTransparency - 3
        Timer1.Interval = 0
    End If
    SetTranslucent Me.hwnd, g_nTransparency
    Exit Sub
ErrorRtn:
    MsgBox Err.Description & " Source : " & Err.Source
End Sub

'Created by Iskandar Setiadi - freedomofkeima 2012

