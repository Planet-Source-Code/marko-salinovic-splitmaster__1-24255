VERSION 5.00
Begin VB.Form frmGlavna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "####"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   FillColor       =   &H00800000&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGlavna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freOpcije 
      Caption         =   "Options"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   5895
      Begin VB.ComboBox cboVelièina 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   165
         Width           =   1575
      End
      Begin VB.PictureBox picBoja 
         Height          =   255
         Left            =   4440
         ScaleHeight     =   195
         ScaleWidth      =   435
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdBoja 
         Caption         =   "..."
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   210
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "File size"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "ProgressBar fill color"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin splitmaster.progressbar bar 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   5535
      _extentx        =   9763
      _extenty        =   450
   End
   Begin splitmaster.open dlg 
      Left            =   4920
      Top             =   1920
      _extentx        =   503
      _extenty        =   450
   End
   Begin VB.CommandButton cmdKraj 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfo 
      Caption         =   "&About"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdSpoji 
      Caption         =   "&Join"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "&Split!!!"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1860
      Width           =   1095
   End
   Begin VB.CommandButton cmdDatoteka 
      Caption         =   "..."
      Height          =   285
      Left            =   5700
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtDatoteka 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   15
      TabIndex        =   13
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Label lblDatoteka 
      Caption         =   "Chose the file you want to split:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmGlavna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kopiraj_datoteka As String


Private Sub cmdBoja_Click()
dlg.ShowColor
picBoja.BackColor = dlg.Color
bar.FillColor = picBoja.BackColor
End Sub

Private Sub cmdDatoteka_Click()
Dim i As Integer

dlg.CancelError = True
dlg.Filter = "Zip file|*.zip|Programs|*.exe;*.bat;*.com|All files|*.*|"
dlg.DialogTitle = "Open dialog"
dlg.ShowOpen

For i = 1 To dlg.cFileName.Count
    txtDatoteka.Text = dlg.cFileName(i)
Next
txtDatoteka.SetFocus
End Sub

Private Sub cmdDirektorij_Click()
frmDir.Show vbModal
End Sub

Private Sub cmdInfo_Click()
'about dialog
frmInfo.Show vbModal
End Sub

Private Sub cmdSplit_Click()
'split the file in the current directory
If SplitFile(txtDatoteka.Text, cboVelièina.ItemData(cboVelièina.ListIndex)) Then
    MsgBox "File was split", vbInformation, "Finished"
Else
    MsgBox "Error splitting file", vbCritical, "Error"
End If
End Sub

Private Sub cmdSpoji_Click()
dlg.CancelError = True
dlg.DialogTitle = "Select file..."
dlg.Filter = "First split file(*.000)|*.000"
dlg.ShowOpen

'join files
If Join(dlg.cFileName(1)) Then
    MsgBox "Finished joining spited files", vbInformation, "Finished :-)"
Else
    MsgBox "Error joining files", vbCritical, "Error :-("
End If
End Sub

Private Sub Form_Load()
'app title
Me.caption = "VBSoft SplitMaster v" & App.Major & "." & App.Minor
Call AddSize(cboVelièina)
cboVelièina.ListIndex = cboVelièina.ListCount - 1
picBoja.BackColor = bar.FillColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub txtDatoteka_GotFocus()
lblDatoteka.FontBold = True
txtDatoteka.BackColor = &HE0E0E0
End Sub

Private Sub txtDatoteka_LostFocus()
lblDatoteka.FontBold = False
txtDatoteka.BackColor = vbWhite
End Sub

Public Sub AddSize(cbo As ComboBox)
Dim i As Integer
For i = 32 To 1400 Step 32
cbo.AddItem i & " kb"
cbo.ItemData(cbo.NewIndex) = i
Next i
End Sub

Function Join(datoteka As String) As Boolean
On Error GoTo greska
Dim zaglavlje As String * 16, buffer As String, datoteka_ext As String, broj_datoteka As Integer, trenutni_broj As Integer, brojac As Integer, t_datoteka As String, nova_datoteka As String
    
Open datoteka For Binary As #1
zaglavlje = Input(Len(zaglavlje), #1)
'is this file split file
If Mid$(zaglavlje, 1, 7) <> "Splited" Then
    MsgBox "This is not a split file", vbInformation, "Error"
    Join = False
    Exit Function
Else
    'header values
    broj_datoteka = Val(Mid$(zaglavlje, 11, 3))
    datoteka_ext = Mid$(zaglavlje, 14, 3)
    If trenutni_broj <> 0 Then
        MsgBox "This is not the first file in the sequence!!! AAAGGHH!"
        Join = False
        Exit Function
    End If
End If
Close #1
nova_datoteka = Left$(datoteka, Len(datoteka) - 3) & datoteka_ext
Open nova_datoteka For Binary As #2
'join files
For brojac = 0 To broj_datoteka - 1
    t_datoteka = Left$(datoteka, Len(datoteka) - 3) & Format$(brojac, "000")
    lblStatus.caption = "Joining file... " & t_datoteka
    lblStatus.Refresh
    Open t_datoteka For Binary As #1
    zaglavlje = Input(Len(zaglavlje), #1)
    If Mid$(zaglavlje, 1, 7) <> "Splited" Then
        MsgBox "This is not a split file", vbInformation, "Error"
        Join = False
        Exit Function
    End If
    trenutni_broj = Val(Mid$(zaglavlje, 8, 3))
    If trenutni_broj <> brojac Then
        MsgBox "The file " & t_datoteka & " is out of sequence", vbInformation, "Error"
        Join = False
        Close #2
        Close #1
        Exit Function
    End If
    While Not EOF(1)
        buffer = Input(10240, #1)
        Put #2, , buffer
    Wend
    Close #1
Next brojac
lblStatus.caption = temp$
Close #2
Join = True
Exit Function

greska:
    Join = False
    MsgBox Err.Description, 16, "Error #" & Err.Number
    Exit Function
End Function

Function SplitFile(datoteka As String, n_velicina As Long) As Boolean
On Error GoTo greska
Dim velicina_datoteka As Long, brojac_datoteka As Integer, broj_datoteka As Integer, velicina_t As Long, buffer As String, p_buffer As String, kraj As Long, s_velicina As Long, zaglavlje As String * 16, brojac As Integer, nova_datoteka As String
Dim dpocetak As Date, dkraj As Date

dpocetak = Now

Open datoteka For Binary As #1
velicina_datoteka = LOF(1)
s_velicina = n_velicina * 1024
If velicina_datoteka <= s_velicina Then
    Close #1
    SplitFile = False
    MsgBox "File is smaler than selected split size", vbInformation, "Error"
    Exit Function
End If
'Check if file isn't alread split
zaglavlje = Input(16, #1)
Close #1
If Mid$(zaglavlje, 1, 7) = "Splited" Then
    MsgBox "This file is alread split", vbInformation, "Error"
    SplitFile = False
    Exit Function
End If
Open datoteka For Binary As #1
velicina_datoteka = LOF(1)
s_velicina = n_velicina * 1024
'header of the split file
broj_datoteka = 0
broj_datoteka = (velicina_datoteka \ s_velicina) + 1
zaglavlje = "Splited" & Format$(brojac, "000") & Format$(broj_datoteka, "000") & Right$(datoteka, 3)
nova_datoteka = Left$(datoteka, Len(datoteka) - 3) & Format$(brojac, "000")
Open nova_datoteka For Binary As #2
'Write the header
Put #2, , zaglavlje
velicina_t = Len(zaglavlje)
While Not EOF(1)
    bar.Max = broj_datoteka
    bar.Value = bar.Value + brojac
    lblStatus.caption = "Spliting file... " & brojac & " (" & Int(velicina_t / 1024) & " kb)"
    lblStatus.Refresh
    buffer = Input(10240, #1)
    velicina_t = velicina_t + Len(buffer)
    If velicina_t > s_velicina Then
        kraj = Len(buffer) - (velicina_t - s_velicina) + Len(zaglavlje)
        Put #2, , Mid$(buffer, 1, kraj)
        Close #2
        'Make new file
        brojac = brojac + 1
        zaglavlje = "Splited" & Format$(brojac, "000") & Format$(broj_datoteka, "000") & Right$(datoteka, 3)
        nova_datoteka = Left$(datoteka, Len(datoteka) - 3) & Format$(brojac, "000")
        Open nova_datoteka For Binary As #2
        'Write the header
        Put #2, , zaglavlje
        Put #2, , Mid$(buffer, kraj + 1)
        velicina_t = Len(zaglavlje) + (Len(buffer) - kraj)
    Else
        Put #2, , buffer
    End If
Wend
dkraj = Now
lblStatus.caption = "Start time: " & dpocetak & "  End: " & dkraj
Close #2
Close #1
SplitFile = True
Exit Function

greska:
    SplitFile = False
    MsgBox Err.Description, 16, "Error #" & Err.Number
    Exit Function
End Function


