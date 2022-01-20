Attribute VB_Name = "modFunciones"
Option Explicit

' Internet, gracias por tanto y perdon por tan poco
' Acá hay APIs de Windows compiladas por muchas personas,
' Usadas para suplir las muchas carencias nativas de VB6

Public Declare Function apiIsZoomed Lib "user32" Alias "IsZoomed" (ByVal hwnd As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

' OnTop:  Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
' Normal: Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

Sub Main()
    fFromFrmMain = False
    Config.CfgPath = App.Path & "\Cfg.ini"
    
    LeerConfig
    
' Si todo sale bien, mostrar frmSplash
    frmSplash.Show
End Sub

Public Function FillTags() As Boolean
' Esta funcion crea campos Tags= vacios en el archivo de Tags para todos los archivos relevados
On Error GoTo oError

Dim sFName As String
Dim sExt() As String
Dim sExtNoDot As String

Dim TagOriginal As String: TagOriginal = Config.TagPath
Dim TagTemporal As String: TagTemporal = Config.TagPath & ".tmp"
Dim TagBuffer As String

Name TagOriginal As TagTemporal

sExtNoDot = Replace(Config.Extension, ".", vbNullString, , , vbTextCompare)

sFName = Dir$(Config.WikiPath)

' Mientras haya archivitos
   While Len(sFName)
   sExt = Split(sFName, ".")
        
        If sExt(UBound(sExt)) = sExtNoDot Then
            SetTags sFName, "Articulo"
        End If
            
            TagBuffer = GetTags(sFName, TagTemporal)
        
        If TagBuffer <> "" Then
            SetTags sFName, TagBuffer
        Else
            SetTags sFName, "Articulo"
        End If
        
        sFName = Dir$()
    Wend
    
    Kill TagTemporal
    
    FillTags = True
    
Exit Function
    
oError:
   FillTags = False
End Function

Public Function ReadFile(Ruta As String) As String
    Dim ff As Integer
    Dim sBuffer As String

' Si no existe
    If Not Exists(Ruta) Then
        ReadFile = "----- ----- -----" & vbNewLine & "La ruta «" & Ruta & "» no existe o no pudo ser encontrada." & vbNewLine & "----- ----- -----"
    End If
    
' Leer línea por línea (Lenght Of File)
    ff = FreeFile
    
    Open Ruta For Input Access Read As #ff
        sBuffer = Input$(LOF(ff), ff)
    Close #ff
    
' Si tiene info mostrar, sino, mostrar mensaje
    If Trim(sBuffer) = "" Then
        ReadFile = EMPTY_FILE
    Else
        ReadFile = sBuffer
    End If
End Function

Public Sub WriteFile(Ruta As String, Contenido As String)
    Dim ff As Integer
    ff = FreeFile
    
    Open Ruta For Output As #ff
        Print #ff, Contenido
    Close #ff
End Sub

Public Function LeerConfig() As Boolean
' Cargar configuracion
    
    If Not Exists(Config.CfgPath) Then
        MsgBox "El archivo de configuración «" & Config.CfgPath & "» no existe o está dañado. Obtenga una nueva copia de la aplicación.", vbExclamation, "Error"
        End
    End If
    
    Config.WikiPath = FormatearRuta(Replace(LeerIni(Config.CfgPath, "MAIN", "RutaWiki"), "&f", App.Path))
    Config.TagPath = Replace(LeerIni(Config.CfgPath, "MAIN", "RutaTags"), "&f", App.Path)
    Config.FontName = LeerIni(Config.CfgPath, "MAIN", "FontName", "Calibri")
    Config.FontSize = Val(LeerIni(Config.CfgPath, "MAIN", "FontSize", 12))
    Config.AlwaysOnTop = IIf(LeerIni(Config.CfgPath, "MAIN", "AlwaysOnTop", "0") = "1", True, False)
    SearchBy = Val(LeerIni(Config.CfgPath, "MAIN", "SearchBy", "0"))

    If Not Exists(Config.WikiPath) Then GoTo fError
    If Not Config.TagPath = "" Then If Not Exists(Config.TagPath) Then GoTo fError
    If Left$(Config.Extension, 2) = "." Then Config.Extension = "." & Config.Extension
    
    LeerConfig = True
    
    Exit Function
    
fError:
    MsgBox "El programa no puede iniciar. Compruebe que las rutas especificadas en el archivo de configuración sean válidas.", vbExclamation
            
    LeerConfig = False
            
    End
    
End Function

Public Function ContarArchivos(Ruta As String, Extension As String) As Long
Dim sFName As String
Dim iCount As Integer
Dim sExt() As String
Dim sExtNoDot As String

    If Not Exists(Ruta) Then
        MsgBox "La ruta <" & Ruta & "> no existe o no pudo ser encontrada.", vbExclamation, "Error"
        Exit Function
    End If

sExtNoDot = Replace(Extension, ".", vbNullString, , , vbTextCompare)

sFName = Dir$(Ruta)

   While Len(sFName)
   sExt = Split(sFName, ".")
        
        If sExt(UBound(sExt)) = sExtNoDot Then iCount = iCount + 1
        
      sFName = Dir$()
   Wend
   
   ContarArchivos = iCount
   
End Function

Public Function FormatearRuta(Ruta As String) As String
' Añade \ si no lo tiene
    If Right$(Ruta, 1) <> "\" Then
        FormatearRuta = Ruta & "\"
    Else
        FormatearRuta = Ruta
    End If
End Function

Public Function Exists(sFileName As String) As Boolean
Dim intReturn As Integer

' Usa GetAttr, si devuelve error, no existe
    On Error GoTo FileExists_Error
    intReturn = GetAttr(sFileName)
    Exists = True
    
Exit Function

FileExists_Error:
    Exists = False
End Function

Public Function HasTags(sFileName As String) As Boolean
    If LeerIni(Config.TagPath, sFileName, "Tags") = "" Then
        HasTags = False
    Else
        HasTags = True
    End If
End Function

Public Function GetTags(sFileName As String, sFrom As String) As String
' Exacto
    GetTags = LeerIni(sFrom, sFileName, "Tags")
End Function

Public Sub SetTags(sFileName As String, Etiquetas As String)
' Exacto
    EscribirIni Config.TagPath, sFileName, "Tags", Etiquetas
End Sub

Public Function FileNameValid(sFileName As String) As String
    Const csInvalidChars As String = ":\/?*<>|"""

    Dim lThisChar As Long
    FileNameValid = sFileName
    For lThisChar = 1 To Len(csInvalidChars)
        FileNameValid = Replace$(FileNameValid, Mid(csInvalidChars, lThisChar, 1), "")
    Next
End Function
