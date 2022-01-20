Attribute VB_Name = "modIni"
Option Explicit

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpSectionNames As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
 
Public Function LeerIni(ByVal strfullpath As String, ByVal strSection As String, ByVal strkey As String, Optional ByVal strDefault As String = "") As String
   Dim strBuffer As String
   Let strBuffer$ = String$(750, Chr$(0&))
   Let LeerIni$ = Left$(strBuffer$, GetPrivateProfileString(strSection$, ByVal LCase$(strkey$), strDefault, strBuffer, Len(strBuffer), strfullpath$))
End Function
 
Public Sub EscribirIni(ByVal strfullpath As String, ByVal strSection As String, ByVal strkey As String, ByVal strkeyvalue As String)
   Call WritePrivateProfileString(strSection$, UCase$(strkey$), strkeyvalue$, strfullpath$)
End Sub
