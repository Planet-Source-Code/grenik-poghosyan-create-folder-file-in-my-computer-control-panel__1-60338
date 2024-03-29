Attribute VB_Name = "mdlRegistry"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_BINARY = 3                     ' Free form binary
Public Sub SaveKey(hKey As Long, strPath As String)
    Dim Keyhand&
    r = RegCreateKey(hKey, strPath, Keyhand&)
    r = RegCloseKey(Keyhand&)
End Sub
Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim Keyhand As Long, datatype As Long, lResult As Long
    Dim strBuf As String, lDataBufSize As Long, intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, Keyhand)
    lResult = RegQueryValueEx(Keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(Keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function
Public Sub SaveString(hKey As Long, strPath As String, strValue As String, strdata As String)
    Dim Keyhand As Long, r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    r = RegSetValueEx(Keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(Keyhand)
End Sub
Function GetDWord(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long, lValueType As Long, lBuf As Long
    Dim lDataBufSize As Long, r As Long, Keyhand As Long
    r = RegOpenKey(hKey, strPath, Keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If
    r = RegCloseKey(Keyhand)
End Function
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim Keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    lResult = RegSetValueEx(Keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(Keyhand)
End Function
Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)
End Function
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim Keyhand As Long
    r = RegOpenKey(hKey, strPath, Keyhand)
    r = RegDeleteValue(Keyhand, strValue)
    r = RegCloseKey(Keyhand)
End Function
Public Sub EnumKey(ByVal hKey As Long, ByVal strPath As String, ByRef cResult As Collection)
    Dim Cnt As Long, sName As String, Keyhand As Long
    RegOpenKey hKey, strPath, Keyhand
    Do
        sName = String(255, vbNullChar)
        If RegEnumKeyEx(Keyhand, Cnt, sName, 255, 0, vbNullString, 0, ByVal 0&) <> 0 Then Exit Do
        cResult.Add StripTerminator(sName)
        Cnt = Cnt + 1
    Loop
    RegCloseKey Keyhand
End Sub
Public Sub EnumValue(ByVal hKey As Long, ByVal strPath As String, ByRef cResult As Collection)
    Dim Cnt As Long, sName As String, Keyhand As Long
    RegOpenKey hKey, strPath, Keyhand
    Do
        sName = String(255, vbNullChar)
        If RegEnumValue(Keyhand, Cnt, sName, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&) <> 0 Then Exit Do
        cResult.Add StripTerminator(sName)
        Cnt = Cnt + 1
    Loop
    RegCloseKey Keyhand
End Sub
Public Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, vbNullChar)
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function
Public Function GetBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, bArray() As Byte) As Boolean
'How to use this function:
'Dim bArray() As Byte
'If GetBinary(KEY, PATH, VALUE, bArray()) = True Then
'   MsgBox StrConv(bArray, vbUnicode)
'End If
    Dim lResult As Long, lValueType As Long, lBuf As Long
    Dim lDataBufSize As Long, r As Long, Keyhand As Long
    r = RegOpenKey(hKey, strPath, Keyhand)
    ' Get length/data type
    lDataBufSize = 0
    ReDim bArray(1 To 1) As Byte
    lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, bArray(1), lDataBufSize)
    If lResult > 0 And lValueType = REG_BINARY Then
        ReDim bArray(1 To lDataBufSize) As Byte
        lResult = RegQueryValueEx(Keyhand, strValueName, 0&, lValueType, bArray(1), lDataBufSize)
        If lResult = ERROR_SUCCESS Then GetBinary = True
    End If
    r = RegCloseKey(Keyhand)
End Function
Public Function SaveBinary(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, bStart As Byte, bLen As Long) As Boolean
'How to use this function:
'Dim bArray(1 To 3) As Byte
'SaveBinary Key, Path, Value, bArray(1), 3
    Dim lResult As Long
    Dim Keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, Keyhand)
    lResult = RegSetValueEx(Keyhand, strValueName, 0&, REG_BINARY, bStart, bLen)
    If lResult = ERROR_SUCCESS Then SaveBinary = True
    r = RegCloseKey(Keyhand)
End Function
