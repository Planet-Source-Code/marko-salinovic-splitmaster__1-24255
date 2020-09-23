Attribute VB_Name = "modAPI"
Public Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

Public Declare Sub GlobalMemoryStatus Lib "KERNEL32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const CREATE_ALWAYS = 2
Public Const OPEN_ALWAYS = 4
Public Const INVALID_HANDLE_VALUE = -1

Public Declare Function ReadFile Lib "KERNEL32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Public Declare Function WriteFile Lib "KERNEL32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function FlushFileBuffers Lib "KERNEL32" (ByVal hFile As Long) As Long

Private Declare Sub GetLocalTime Lib "KERNEL32" (lpSystemTime As SystemTime)

Type SystemTime
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Property Get getTime() As Double
Dim currentTime As SystemTime
    Call GetLocalTime(currentTime)
    getTime = (currentTime.wHour * CLng(3600000) + currentTime.wMinute * CLng(60000) + currentTime.wSecond * CLng(1000) + currentTime.wMilliseconds) / CLng(1000)
    Exit Property
End Property

Public Function DetermineDirectory(inputString As String) As String
Dim pos As Integer
    pos = InStrRev(inputString, "\", , vbTextCompare)
    DetermineDirectory = Mid(inputString, 1, pos)
End Function

Public Function DetermineFilename(inputString As String) As String
Dim pos As Integer
    pos = InStrRev(inputString, "\", , vbTextCompare)
    DetermineFilename = Mid(inputString, pos + 1, Len(inputString) - pos)
End Function

Public Function DetermineDrive(inputString As String) As String
Dim pos As Integer
    If inputString = "" Then Exit Function
    pos = InStr(1, inputString, ":\", vbTextCompare)
    DetermineDrive = Mid(inputString, 1, pos - 1)
End Function

