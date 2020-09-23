Attribute VB_Name = "modReg"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_SZ = 1

Private Const ERROR_SUCCESS = 0&
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Function RGGetKeyValue(hKey As Long, SubKey As String, ValueName As String, Optional Default As String = "")
Dim lngRet As Long
Dim lngResult As Long
Dim sData As String
    
'Open key
lngRet = RegOpenKeyEx(hKey, SubKey, 0, KEY_ALL_ACCESS, lngResult)
If lngRet = ERROR_SUCCESS Then 'If key exist
    sData = String(128, vbNullChar) 'Fill buffer with null chars
    'Get key value
    lngRet = RegQueryValueEx(lngResult, ValueName, 0, REG_SZ, ByVal sData, Len(sData))
    'If valuename doesnt exist set default value
    If Not lngRet = ERROR_SUCCESS Then RGGetKeyValue = Default: Exit Function
        RGGetKeyValue = Left(sData, InStr(1, sData, vbNullChar) - 1)
        RegCloseKey lngResult 'Close key
    Else
        RGGetKeyValue = Default
End If
End Function

