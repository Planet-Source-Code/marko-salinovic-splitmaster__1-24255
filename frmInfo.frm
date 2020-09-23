VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About..."
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUredu 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ListBox lstInfo 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label lblTekst 
      Caption         =   "######"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label lblNaslov 
      Alignment       =   2  'Center
      Caption         =   "VBSoft SplitMaster"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmInfo.frx":0000
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUredu_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim ram As MEMORYSTATUS
GlobalMemoryStatus ram

Dim autor As String
Dim drzava As String
Dim mail As String
Dim korisnik As String
Dim duzina As Long
Dim rezultat As Long
Dim lKeyHandle As Long
Dim lTemp As Long, sData As String, sDataLen As Long

korisnik = Space(255)
duzina = 255
rezultat = GetUserName(korisnik, duzina)
c$ = Trim(korisnik)

b$ = RGGetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
d$ = RGGetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductName")
e$ = RGGetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "VersionNumber")

a$ = "Autor: Marko Šalinoviæ &" + vbNewLine + "E-mail: mate.salinovic1@st.tel.hr" + vbNewLine + "Country: Croatia (Hrvatska)"
lblTekst.caption = a$

'user name and comp info
lstInfo.AddItem "User name - " & c$
lstInfo.AddItem "RAM - " & Format(ram.dwTotalPhys, "@@@@@@@@@@@") / 1024 & " kb"
lstInfo.AddItem "Registered owner - " + b$
lstInfo.AddItem "OS - " & d$ + " build " & e$
End Sub
