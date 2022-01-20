Attribute VB_Name = "modDeclarations"
Option Explicit

' No, no voy a explicar cada Variable

    Type udtConfig
        WikiPath            As String
        CfgPath             As String
        TagPath             As String
        EditPath            As String
        
        Extension           As String
        
        FontName            As String
        FontSize            As Integer
        AlwaysOnTop         As Boolean
    End Type

' Variables
Public Config               As udtConfig

Public SearchBy             As Integer
Public LastChecked          As Byte
Public aItemPath()          As String

' Flags
Public fMouseDown           As Boolean
Public fTextOnlyMode        As Boolean
Public fConfigReady         As Boolean
Public fIsSearchFilterOn    As Boolean
Public fFromFrmMain         As Boolean

' Constantes
Public Const byContent = 0
Public Const byTags = 1
Public Const byFilename = 2
Public Const searchAll = 3

Public Const EMPTY_FILE = "----- ----- -----" & vbNewLine & "El archivo no contiene información." & vbNewLine & "----- ----- -----"
