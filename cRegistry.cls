VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_MULTI_SZ = 7
Private Const ERROR_MORE_DATA = 234
Private Const REG_OPENED_EXISTING_KEY = &H2

Public Enum HKEYS
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Private Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))

Private Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                           ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))


Public Function IsWinNT() As Boolean
' Returns True is executed under Windows NT,
' False if executed under Windows 95/98

  IsWinNT = (GetVersion() And &H80000000) = 0

End Function


' Read a Registry value
'
' Use KeyName = "" for the default value
' If the value isn't there, it returns the DefaultValue
' argument, or Empty if the argument has been omitted
'
' Supports DWORD, REG_SZ, REG_EXPAND_SZ, REG_BINARY and REG_MULTI_SZ
' REG_MULTI_SZ values are returned as a null-delimited stream of strings
' (VB6 users can use Split to convert to an array of string)

Public Function GetRegistryValue(ByVal hKey As HKEYS, ByVal KeyName As String, ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
Dim handle As Long
Dim resLong As Long
Dim resString As String
Dim resBinary() As Byte
Dim length As Long
Dim retVal As Long
Dim valueType As Long

' Prepare the default result
  GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)

' Open the key, exit if not found.
  If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
    Exit Function
  End If

' prepare a 1K receiving resBinary
  length = 1024
  ReDim resBinary(0 To length - 1) As Byte

' read the registry key
  retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)

' if resBinary was too small, try again
  If retVal = ERROR_MORE_DATA Then
  ' enlarge the resBinary, and read the value again
    ReDim resBinary(0 To length - 1) As Byte
    retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
  End If

' return a value corresponding to the value type
  Select Case valueType
    Case REG_DWORD
      CopyMemory resLong, resBinary(0), 4
      GetRegistryValue = resLong
    Case REG_SZ, REG_EXPAND_SZ
    ' copy everything but the trailing null char
      resString = Space$(length - 1)
      CopyMemory ByVal resString, resBinary(0), length - 1
       GetRegistryValue = resString
    Case REG_BINARY
    ' resize the result resBinary
      If length <> UBound(resBinary) + 1 Then
        ReDim Preserve resBinary(0 To length - 1) As Byte
      End If
      GetRegistryValue = resBinary()
    Case REG_MULTI_SZ
    ' copy everything but the 2 trailing null chars
      resString = Space$(length - 2)
      CopyMemory ByVal resString, resBinary(0), length - 2
          GetRegistryValue = resString
    Case Else
      RegCloseKey handle
    '  Err.Raise 1001, , "Unsupported value type"
  End Select

' close the registry key
  RegCloseKey handle
  
End Function

' Write or Create a Registry value
' returns True if successful
'
' Use KeyName = "" for the default value
'
' Value can be an integer value (REG_DWORD), a string (REG_SZ)
' or an array of binary (REG_BINARY). Raises an error otherwise.

Public Function SetRegistryValue(ByVal hKey As HKEYS, ByVal KeyName As String, ByVal ValueName As String, Value As Variant) As Boolean
Dim handle As Long
Dim lngValue As Long
Dim strValue As String
Dim binValue() As Byte
Dim length As Long
Dim retVal As Long

' Open the key, exit if not found
  If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then
    Exit Function
  End If

' three cases, according to the data type in Value
  Select Case VarType(Value)
    Case vbInteger, vbLong
      lngValue = Value
      retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
    Case vbString
      strValue = Value
      retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
      Len(strValue))
    Case vbBoolean
      lngValue = Abs(Value)
      retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
    Case vbArray + vbByte
      binValue = Value
      length = UBound(binValue) - LBound(binValue) + 1
      retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, binValue(LBound(binValue)), length)

    Case Else
      RegCloseKey handle
    '  Err.Raise 1001, , "Unsupported value type"
  End Select

' Close the key and signal success
  RegCloseKey handle
    
' signal success if the value was written correctly
  SetRegistryValue = (retVal = 0)
  
End Function

' Enumerate values under a given registry key returns a collection,
' where each element of the collection is a 2-element array of Variants:
' element(0) is the value name, element(1) is the value's value

Public Function EnumRegistryValues(ByVal hKey As HKEYS, ByVal KeyName As String) As Collection
Dim handle As Long
Dim index As Long
Dim valueType As Long
Dim name As String
Dim nameLen As Long
Dim resLong As Long
Dim resString As String
Dim dataLen As Long
Dim valueInfo(0 To 1) As Variant
Dim retVal As Long

' initialize the result
  Set EnumRegistryValues = New Collection

' Open the key, exit if not found.
  If Len(KeyName) Then
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
    ' in all cases, subsequent functions use hKey
      hKey = handle
  End If

  Do
  ' this is the max length for a key name
    nameLen = 260
    name = Space$(nameLen)
  ' prepare the receiving buffer for the value
    dataLen = 4096
    ReDim resBinary(0 To dataLen - 1) As Byte

  ' read the value's name and data exit the loop if not found
    retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)

  ' enlarge the buffer if you need more space
    If retVal = ERROR_MORE_DATA Then
      ReDim resBinary(0 To dataLen - 1) As Byte
      retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, _
      valueType, resBinary(0), dataLen)
    End If
      
  ' exit the loop if any other error (typically, no more values)
    If retVal Then Exit Do

  ' retrieve the value's name
    valueInfo(0) = Left$(name, nameLen)

  ' return a value corresponding to the value type
    Select Case valueType
      Case REG_DWORD
        CopyMemory resLong, resBinary(0), 4
        valueInfo(1) = resLong
      Case REG_SZ, REG_EXPAND_SZ
      ' copy everything but the trailing null char
        resString = Space$(dataLen - 1)
        CopyMemory ByVal resString, resBinary(0), dataLen - 1
        valueInfo(1) = resString
      Case REG_BINARY
      ' shrink the buffer if necessary
        If dataLen < UBound(resBinary) + 1 Then
          ReDim Preserve resBinary(0 To dataLen - 1) As Byte
        End If
        valueInfo(1) = resBinary()
      Case REG_MULTI_SZ
      ' copy everything but the 2 trailing null chars
        resString = Space$(dataLen - 2)
        CopyMemory ByVal resString, resBinary(0), dataLen - 2
        valueInfo(1) = resString
      Case Else
      ' Unsupported value type - do nothing
    End Select

  ' add the array to the result collection the element's key is the value's name
    ' EnumRegistryValues.Add valueInfo(1), valueInfo(0)
      EnumRegistryValues.Add valueInfo, valueInfo(0)
    index = index + 1
  Loop

' Close the key, if it was actually opened
  If handle Then RegCloseKey handle

End Function

' Enumerate registry keys under a given key returns a collection of strings
Public Function EnumRegistryKeys(ByVal hKey As Long, ByVal KeyName As String) As Collection
Dim handle As Long
Dim length As Long
Dim index As Long
Dim subkeyName As String
    
' initialize the result collection
  Set EnumRegistryKeys = New Collection
    
' Open the key, exit if not found
  If Len(KeyName) Then
  
  If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
  ' in all case the subsequent functions use hKey
    hKey = handle
  End If
    
  Do
  ' this is the max length for a key name
    length = 260
    subkeyName = Space$(length)
  ' get the N-th key, exit the loop if not found
    If RegEnumKey(hKey, index, subkeyName, length) Then Exit Do
        
  ' add to the result collection
    subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
    EnumRegistryKeys.Add subkeyName, subkeyName
  
  ' prepare to query for next key
    index = index + 1
  Loop
   
' Close the key, if it was actually opened
  If handle Then RegCloseKey handle
        
End Function

'Enumerate values under a given registry key.
'  Returns a collection, where each element of the collection is a 3-element array
'  of Variants: element(0) is the value name, element(1) is the value's value,
'  element(2) is the type of data type

Function EnumRegistryValuesEx(ByVal hKey As Long, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 2) As Variant
    Dim retVal As Long
    
    ' initialize the result
    Set EnumRegistryValuesEx = New Collection
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
                valueInfo(2) = vbLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
                valueInfo(2) = vbString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
                valueInfo(2) = vbArray + vbByte
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
                valueInfo(2) = vbString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValuesEx.Add valueInfo, valueInfo(0)
        
        index = index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function



'Save the specified registry's key and (optionally) its subkeys to a REG file
' that can be loaded later
' - hKey is the root key
' - sKeyName is the key to save to the file
' - sRegFile is the target file where the text will be saved
' - bIncludeSubKeys specifies whether the routine will save also the subkeys
' - bAppendToFile specifies wheter the generated text will be appended to an
' existent file
'Example:
'  SaveRegToFile HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0",
'  "C:\vb6.reg"

'NOTE: this routine requires EnumRegistryKeys and EnumRegistryValuesEx

Sub SaveRegToFile(ByVal hKey As Long, ByVal sKeyName As String, _
    ByVal sRegFile As String, Optional ByVal bIncludeSubKeys As Boolean = True, _
    Optional ByVal bAppendToFile As Boolean = False)
    
    Dim handle As Integer
    Dim sFirstKeyPart As String
    Dim col As New Collection
    Dim regItem As Variant
    Dim sText As String
    Dim sQuote As String
    Dim sTemp As String
    Dim sHex As String
    Dim i As Long
    Dim vValue As Variant
    Dim iPointer As MousePointerConstants
    Dim sValueName As String
    
    sQuote = Chr$(34)
    
    On Error Resume Next
     
    'conver the hKey value to the descriptive string
    Select Case hKey
        Case HKEY_CLASSES_ROOT: sFirstKeyPart = "HKEY_CLASSES_ROOT\"
        Case HKEY_CURRENT_CONFIG: sFirstKeyPart = "HKEY_CURRENT_CONFIG\"
        Case HKEY_CURRENT_USER: sFirstKeyPart = "HKEY_CURRENT_USER\"
        Case HKEY_LOCAL_MACHINE: sFirstKeyPart = "HKEY_LOCAL_MACHINE\"
        Case HKEY_USERS: sFirstKeyPart = "HKEY_USERS\"
    End Select
    
    'this can be a long operation
    iPointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    
    'if the text won't be appended, add the "REGEDIT4" header
    If bAppendToFile = False Then
        sText = "REGEDIT4" & vbCrLf & vbCrLf
    Else
        'add the same header if the text will be appended to an
        'existent file that does not contain the header.
        ' This works only if the file exists but is empty.
        handle = FreeFile
        Open sRegFile For Binary As #handle
        ' read the string and close the file
        sTemp = Space$(LOF(handle))
        Get #handle, , sTemp
        Close #handle
        'if not found, add it
        If InStr(1, sTemp, "REGEDIT4") = 0 Then
            sText = "REGEDIT4" & vbCrLf & vbCrLf
        End If
    End If
    
    'save the key name with the format [keyname]
    sText = sText & "[" & sFirstKeyPart & sKeyName & "]" & vbCrLf
    
    'get the collection with all the values under this key
    Set col = EnumRegistryValuesEx(hKey, sKeyName)
    For Each regItem In col
        vValue = regItem(1)
        Select Case regItem(2)
            Case vbString
                'if the value is a string, check if it's a path by looking if
                ' the 3 characters
                'are in the form X:\. If so, replace a single "\" with "\\"
                If Left$(vValue, 3) Like "[A-Z,a-z]:\" Then vValue = Replace _
                    (vValue, "\", "\\")
                'quote it
                sTemp = sQuote & vValue & sQuote
            Case vbLong
                'if it's a long, save it with the format dword:num
                sTemp = "dword:" & CLng(vValue)
            Case vbArray + vbByte
                'if it's an array of bytes, save it with the format hex:num1,
                ' num2,num3,...
                sTemp = "hex:"
                For i = 0 To UBound(vValue)
                    sHex = Hex$(vValue(i))
                    'convert from long to hex
                    If Len(sHex) < 2 Then sHex = "0" & sHex
                    sTemp = sTemp & sHex & ","
                Next
                'remove the last comma
                sTemp = Left$(sTemp, Len(sTemp) - 1)
            Case Else
                sTemp = ""
        End Select
        'get the value name: if the string is empty, take @,
        '  else take that name and quote it
        sValueName = IIf(Len(regItem(0)) > 0, sQuote & regItem(0) & sQuote, "@")
        'save this line to the temporary text that will be saved
        sText = sText & sValueName & "=" & sTemp & vbCrLf
    Next
    sText = sText & vbCrLf
    
    handle = FreeFile
    'open the target file with Append or Output mode,
    '  according to the bAppendToFile parameter
    If bAppendToFile Then
        Open sRegFile For Append As #handle
    Else
        Open sRegFile For Output As #handle
    End If
    'save the text
    Print #handle, sText;
    Close #handle
    
    'call recursively this routine to save all the subkeys,
    '  if the bIncludeSubKeys param is true
    If bIncludeSubKeys Then
        Set col = EnumRegistryKeys(hKey, sKeyName)
        For Each regItem In col
            'note: the text will be added to the file just created for the
            'values in the root key
            SaveRegToFile hKey, sKeyName & "\" & regItem, sRegFile, True, True
        Next
    End If
    
    Screen.MousePointer = iPointer
    
End Sub

' Load the specified REG file in the registry
Public Sub ApplyRegFile(ByVal sRegFile As String)
  On Error Resume Next
' // first of all, check if the file exists
  If Not (GetAttr(sRegFile) And vbDirectory) = 0 Then Exit Sub
  ' load the reg file
  ' quote the file name: this is necessary if the file name is something like "token1 token2.reg"
    Shell "regedit /s " & Chr$(34) & sRegFile & Chr$(34)
    
End Sub






' ===[ DELETE ] ===

' Delete a registry key
' Under Windows NT it doesn't work if the key contains subkeys

Public Sub DeleteRegistryKey(ByVal hKey As HKEYS, ByVal KeyName As String)
    RegDeleteKey hKey, KeyName
End Sub

' Delete a registry value
' Return True if successful, False if the value hasn't been found

Public Function DeleteRegistryValue(ByVal hKey As HKEYS, ByVal KeyName As String, ByVal ValueName As String) As Boolean
Dim handle As Long

' Open the key, exit if not found
  If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then Exit Function

' Delete the value (returns 0 if success)
  DeleteRegistryValue = (RegDeleteValue(handle, ValueName) = 0)

' Close the handle
  RegCloseKey handle
End Function


' Return True if a Registry key exists

Public Function CheckRegistryKey(ByVal hKey As HKEYS, ByVal KeyName As String) As Boolean
Dim handle As Long
' Try to open the key
  CheckRegistryKey = (RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0)

' Close it before exiting
  RegCloseKey handle
 
End Function

' Create a registry key, then close it
Function CreateRegistryKey(ByVal hKey As HKEYS, ByVal KeyName As String) As Boolean
' Pre  :
' Post : Returns True if the key already existed, False if it was created.
Dim handle As Long, disposition As Long
    
  If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, 0, handle, disposition) Then
    Err.Raise 1001, , "Unable to create the registry key"
  Else
  ' Return True if the key already existed.
    CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
  ' Close the key.
    RegCloseKey handle
  End If
End Function


