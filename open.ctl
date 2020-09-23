VERSION 5.00
Begin VB.UserControl open 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   285
   InvisibleAtRuntime=   -1  'True
   Picture         =   "open.ctx":0000
   ScaleHeight     =   255
   ScaleWidth      =   285
End
Attribute VB_Name = "open"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Const def_CancelError = 0
Const def_Filename = ""
Const def_DialogTitle = "Open file"
Const def_InitialDir = ""
Const def_Filter = ""
Const def_FilterIndex = 1
Const def_MultiSelect = 0
Const def_color = vbBlue

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CF_SCREENFONTS = &H1
Private Const CF_EFFECTS = &H100&
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_LIMITSIZE = &H2000&
Private Const DEFAULT_CHARSET = 1
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_QUALITY = 0
Private Const FW_BOLD = 700
Private Const FF_ROMAN = 16
Private Const FW_NORMAL = 400
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const OUT_DEFAULT_PRECIS = 0
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&
Private Const MAX_PATH = 260

Dim m_CancelError As Boolean
Dim m_Filename As String
Dim m_DialogTitle As String
Dim m_InitialDir As String
Dim m_Filter As String
Dim m_FilterIndex As Integer
Dim m_MultiSelect As Boolean
Dim m_color As OLE_COLOR

Public cFileName As New Collection
Public cFileTitle As New Collection

' << Cancel error >>
Public Property Get CancelError() As Boolean
CancelError = m_CancelError
End Property

Public Property Let CancelError(ByVal new_cancel_error As Boolean)
m_CancelError = new_cancel_error
PropertyChanged "CancelError"
End Property

' << Color >>
Public Property Get Color() As OLE_COLOR
Color = m_color
End Property

Public Property Let Color(hColor As OLE_COLOR)
m_color = hColor
PropertyChanged "Color"
End Property

' << Multi select >>
Public Property Get MultiSelect() As Boolean
MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal new_multi_select As Boolean)
m_MultiSelect = new_multi_select
PropertyChanged "Multiselect"
End Property

' << Default filename >>
Public Property Get DefaultFilename() As String
DefaultFilename = m_Filename
End Property

Public Property Let DefaultFilename(ByVal new_file_name As String)
m_Filename = new_file_name
PropertyChanged "DefaultFilename"
End Property

' << Dialog title >>
Public Property Get DialogTitle() As String
DialogTitle = m_DialogTitle
End Property

Public Property Let DialogTitle(ByVal new_dialog_title As String)
m_DialogTitle = new_dialog_title
PropertyChanged "DialogTitle"
End Property

' << Initial directory >>
Public Property Get InitDir() As String
InitDir = m_InitialDir
End Property

Public Property Let InitDir(ByVal new_init_dir As String)
m_InitialDir = new_init_dir
PropertyChanged "InitDir"
End Property

' << Filter >>
Public Property Get Filter() As String
Filter = m_Filter
End Property

Public Property Let Filter(ByVal New_Filter As String)
m_Filter = New_Filter
PropertyChanged "Filter"
End Property

' << Filter index >>
Public Property Get FilterIndex() As Integer
FilterIndex = m_FilterIndex
End Property

Public Property Let FilterIndex(ByVal new_filter_index As Integer)
m_FilterIndex = new_filter_index
PropertyChanged "FilterIndex"
End Property

Private Sub UserControl_Initialize()
'resize control
UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
'initilaze propertys
m_CancelError = def_CancelError
m_Filename = def_Filename
m_DialogTitle = def_DialogTitle
m_InitialDir = def_InitialDir
m_Filter = def_Filter
m_FilterIndex = def_FilterIndex
m_MultiSelect = def_MultiSelect
m_color = def_color
End Sub

Private Sub UserControl_ReadProperties(bag As PropertyBag)
m_CancelError = bag.ReadProperty("CancelError", def_CancelError)
m_Filename = bag.ReadProperty("DefaultFilename", def_Filename)
m_DialogTitle = bag.ReadProperty("DialogTitle", def_DialogTitle)
m_InitialDir = bag.ReadProperty("InitDir", def_InitialDir)
m_Filter = bag.ReadProperty("Filter", def_Filter)
m_FilterIndex = bag.ReadProperty("FilterIndex", def_FilterIndex)
m_MultiSelect = bag.ReadProperty("Multiselect", def_MultiSelect)
m_color = bag.ReadProperty("Color", def_color)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 255
UserControl.Width = 285
End Sub

Private Sub UserControl_WriteProperties(bag As PropertyBag)
Call bag.WriteProperty("CancelError", m_CancelError, def_CancelError)
Call bag.WriteProperty("DefaultFilename", m_Filename, def_Filename)
Call bag.WriteProperty("DialogTitle", m_DialogTitle, def_DialogTitle)
Call bag.WriteProperty("InitDir", m_InitialDir, def_InitialDir)
Call bag.WriteProperty("Filter", m_Filter, def_Filter)
Call bag.WriteProperty("FilterIndex", m_FilterIndex, def_FilterIndex)
Call bag.WriteProperty("MultiSelect", m_MultiSelect, def_MultiSelect)
Call bag.WriteProperty("Color", m_color, def_color)
End Sub

'shows open dialog with out CommonDlg.oxc
Public Function ShowOpen()
Dim OPENFILENAME As OPENFILENAME
Dim Ret As Long

With OPENFILENAME
If MultiSelect Then
    .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    .lpstrFile = DefaultFilename & Space(9999 - Len(DefaultFilename)) & vbNullChar
    .lpstrTitle = Space(9999) & vbNullChar
Else
    .flags = OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    .lpstrFile = DefaultFilename & String(MAX_PATH - Len(DefaultFilename), 0) & vbNullChar
    .lpstrFileTitle = String(MAX_PATH, 0) & vbNullChar
End If

.hwndOwner = UserControl.ContainerHwnd
.lpstrFilter = SetFilter(Filter) & vbNullChar
.lpstrInitialDir = InitDir & vbNullChar
.lpstrTitle = DialogTitle & vbNullChar
.lStructSize = Len(OPENFILENAME)
.nFilterIndex = FilterIndex
.nMaxFile = Len(.lpstrFile)
.nMaxFileTitle = Len(.lpstrFileTitle)
End With

'call open dialog
Ret = GetOpenFileName(OPENFILENAME)
If Ret <> 0 Then
    Rncffapmf OPENFILENAME.lpstrFile
Else
    If CancelError Then
        MsgBox "Cancel was selected", vbCritical, "Cancel"
    End If
End If
End Function

' << Replace "|" with Null char >>
Private Function SetFilter(sFilter As String) As String
Dim sLen As Long
Dim Pos As Long

sLen = Len(sFilter) 'filter lenght
Pos = InStr(1, sFilter, "|")

' Loop while Pos > 0
While Pos > 0
    sFilter = Left(sFilter, Pos - 1) & vbNullChar & Mid(sFilter, Pos + 1, sLen - Pos)
        
    Pos = InStr(Pos + 1, sFilter, "|")
Wend
SetFilter = sFilter
End Function

'Rncffapmf - remove null chars from filename and parse multi filename
Private Function Rncffapmf(sFileName As String)
Dim i As Long, sFiles() As String, Pos As Integer, sFile As String, sFileTitle As String

Set cFileName = New Collection
Set cFileTitle = New Collection

'last two null chars
Pos = InStr(sFileName, vbNullChar & vbNullChar)
sFile = Left(sFileName, Pos - 1)

If InStr(1, sFile, vbNullChar) <> 0 Then
'multi
    sFile = Left(sFileName, Pos) & vbNullChar
    sFile = Left(sFileName, InStr(1, sFileName, Chr(0)) - 1)
    sFiles = Split(sFile, Chr(0))
    
    'add all filenames to collection
    For i = LBound(sFiles) To UBound(sFiles) - 2
        If Right(sPath, 1) = "\" Then
            cFileName.Add sPath & sFiles(i)
        Else
            cFileName.Add sPath & "\" & sFiles(i)
        End If
        
        'file Title
        cFileTitle.Add sFiles(i)
        If i = 1 Then cFileName.Remove 1
        cFileTitle.Remove 1
    Next
Else
    'single
    'add file name to collection
    cFileName.Add sFile
    'file title
    cFileTitle.Add Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
End If
End Function

'shows color dialog with out CommonDlg.oxc
Public Function ShowColor()
Dim c_color As ChooseColor
Dim Ret As Long
Dim Col(0 To 16) As Long
Dim i As Integer

'fill colors with white
For i = 0 To 15
    Col(i) = vbWhite
Next

c_color.hwndOwner = UserControl.ContainerHwnd
c_color.lStructSize = Len(c_color)
c_color.lpCustColors = VarPtr(Col(0))
c_color.rgbResult = 0

Ret = ChooseColor(c_color)
If Ret <> 0 Then
    ShowColor = c_color.rgbResult
    m_color = c_color.rgbResult
Else
    If CancelError Then
        MsgBox "Cancel was selected", vbCritical, "Cancel"
    End If
End If
End Function

Public Function About()
MsgBox "Autor: Marko Šalinoviæ", vbInformation, "About Open.ocx"
End Function

