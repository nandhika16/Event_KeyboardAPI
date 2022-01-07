VERSION 5.00
Begin VB.Form PantunPerkenalan 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   Picture         =   "PantunPerkenalan.frx":0000
   ScaleHeight     =   3930
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton mulai 
      BackColor       =   &H80000000&
      Caption         =   "Mulai"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   4800
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   4800
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line2 
      X1              =   4800
      X2              =   4800
      Y1              =   360
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   360
      Y2              =   1800
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Event Keyboard using MS Paint"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PANTUN PERKENALAN"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "PantunPerkenalan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_0 = 48
Const VK_1 = 49
Const VK_2 = 50
Const VK_3 = 51
Const VK_4 = 52
Const VK_5 = 53
Const VK_6 = 54
Const VK_7 = 55
Const VK_8 = 56
Const VK_9 = 57
Const VK_A = 65
Const VK_B = 66
Const VK_C = 67
Const VK_D = 68
Const VK_E = 69
Const VK_F = 70
Const VK_G = 71
Const VK_H = 72
Const VK_I = 73
Const VK_J = 74
Const VK_K = 75
Const VK_L = 76
Const VK_M = 77
Const VK_N = 78
Const VK_O = 79
Const VK_P = 80
Const VK_Q = 81
Const VK_R = 82
Const VK_S = 83
Const VK_T = 84
Const VK_U = 85
Const VK_V = 86
Const VK_W = 87
Const VK_X = 88
Const VK_Y = 89
Const VK_Z = 90
Const VK_UP = 38
Const VK_RIGHT = 39
Const VK_TAB = 9
Const VK_RETURN = 13
Const VK_OEM_TANYA = 63
Const VK_SPACE = 32
Const VK_LWIN = 91
Const VK_LSHIFT = 160
Const VK_LCONTROL = 162
Const VK_OEM_COMMA = 188
Const VK_OEM_MINUS = 189
Const VK_OEM_PERIOD = 190
Const VK_MENU = 18
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub press(Key As String)
    keybd_event Key, 0, 0, 0
    keybd_event Key, 0, KEYEVENTF_KEYUP, 0
End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub mulai_Click()
    'Membuka mspaint
    '1. Windows + R
    keybd_event VK_LWIN, 0, 0, 0
    press (VK_R)
    keybd_event VK_LWIN, 0, KEYEVENTF_KEYUP, 0
    Sleep (500)
    
    '2. Mengetik 'mspaint' & Menekan 'OK'
    press (VK_M)
    Sleep (100)
    press (VK_S)
    Sleep (100)
    press (VK_P)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_T)
    Sleep (100)
    press (VK_RETURN)
    Sleep (500)
    
    '3. Text
    keybd_event VK_MENU, 0, 0, 0
    press (VK_H)
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    Sleep (500)
    press (VK_T)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    press (VK_RETURN)
    Sleep (500)
    
    'Mengetikkan Sebuah Pantun Pada "mspaint"
    ' Baris 1
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_B)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_G)
    Sleep (100)
    press (VK_U)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_P)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_G)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_L)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_G)
    Sleep (100)
    press (VK_S)
    Sleep (100)
    press (VK_U)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_G)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_O)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_H)
    Sleep (100)
    press (VK_R)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_G)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_RETURN)
    Sleep (100)
    
    'Baris 2
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_L)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_R)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_OEM_MINUS)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_R)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_K)
    Sleep (100)
    press (VK_E)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_G)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_K)
    Sleep (100)
    press (VK_O)
    Sleep (100)
    press (VK_T)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_RETURN)
    Sleep (100)
    
    'Baris 3
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_K)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_U)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_B)
    Sleep (100)
    press (VK_O)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_E)
    Sleep (100)
    press (VK_H)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_A)
    Sleep (100)
    press (VK_K)
    Sleep (100)
    press (VK_U)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_B)
    Sleep (100)
    press (VK_E)
    Sleep (100)
    press (VK_R)
    Sleep (100)
    press (VK_T)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_Y)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_RETURN)
    Sleep (100)
    
    'Baris 4
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_N)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_O)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_C)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_T)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_K)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_S)
    Sleep (100)
    press (VK_I)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_P)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_N)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_M)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_Y)
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_RETURN)
    Sleep (100)
    
    
    'Bold text dan warna
    keybd_event VK_LCONTROL, 0, 0, 0
    press (VK_A)
    keybd_event VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep (200)
    
        'Bold Text
        press (VK_MENU)
        Sleep (200)
        press (VK_T)
        Sleep (200)
        press (VK_F)
        Sleep (100)
        press (VK_B)
        Sleep (200)
        
        'Warna Text
        press (VK_MENU)
        Sleep (200)
        press (VK_T)
        Sleep (200)
        press (VK_E)
        Sleep (100)
        press (VK_C)
        Sleep (200)
        press (VK_UP)
        Sleep (100)
        press (VK_UP)
        Sleep (100)
        press (VK_RIGHT)
        Sleep (200)
        press (VK_SPACE)
        Sleep (200)
        press (VK_RETURN)
        Sleep (200)
        
        
        
    'Save Mspaint
    '1. Ctrl + S
    keybd_event VK_LCONTROL, 0, 0, 0
    press (VK_S)
    keybd_event VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep (2000)
    
    '2. Penamaan File
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_P)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_A)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_T)
    Sleep (100)
    press (VK_U)
    Sleep (100)
    press (VK_N)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_P)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_E)
    Sleep (100)
    press (VK_R)
    Sleep (100)
    press (VK_K)
    Sleep (100)
    press (VK_E)
    Sleep (500)
    press (VK_N)
    Sleep (500)
    press (VK_A)
    Sleep (500)
    press (VK_L)
    Sleep (500)
    press (VK_A)
    Sleep (500)
    press (VK_N)
    Sleep (500)
    
    '3. Membuat Folder untuk Menyimpan File
    keybd_event VK_LCONTROL, 0, 0, 0
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_N)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep (2000)
    
    '4. Menamakan Folder 'API Kelompok 04'
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_A)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_P)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_I)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    keybd_event VK_LSHIFT, 0, 0, 0
    press (VK_K)
    keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0
    Sleep (100)
    press (VK_E)
    Sleep (100)
    press (VK_L)
    Sleep (100)
    press (VK_O)
    Sleep (100)
    press (VK_M)
    Sleep (100)
    press (VK_P)
    Sleep (100)
    press (VK_O)
    Sleep (100)
    press (VK_K)
    Sleep (100)
    press (VK_SPACE)
    Sleep (100)
    
    press (VK_0)
    Sleep (100)
    press (VK_4)
    Sleep (500)
    
    '5. Membuka Folder
    press (VK_RETURN)
    Sleep (500)
    press (VK_RETURN)
    Sleep (500)
    
    '6. Mencari Tombol Save
    press (VK_TAB)
    Sleep (500)
    press (VK_TAB)
    Sleep (500)
    press (VK_TAB)
    Sleep (500)
    press (VK_RETURN)
    Sleep (500)
     
    
End Sub

