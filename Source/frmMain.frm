VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Wiki"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReturn 
      Height          =   4500
      Left            =   120
      TabIndex        =   11
      Top             =   -15
      Visible         =   0   'False
      Width           =   495
      Begin VB.PictureBox picRevertView 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         Picture         =   "frmMain.frx":17D2A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   5880
         Width           =   240
      End
   End
   Begin VB.Frame fraContent 
      Height          =   5295
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdOk 
         Caption         =   "Listo"
         Height          =   405
         Left            =   120
         TabIndex        =   13
         Top             =   225
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtContent 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtTags 
         BackColor       =   &H8000000F&
         Height          =   390
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "txtTags"
         Top             =   225
         Width           =   2775
      End
      Begin VB.Label lblEditorMode 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Modo de Edición"
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   4140
         TabIndex        =   14
         Top             =   285
         Width           =   1620
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox picTextOnly 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3360
         Picture         =   "frmMain.frx":180B4
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   4920
         Width           =   240
      End
      Begin VB.PictureBox picTagEdit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         Picture         =   "frmMain.frx":1843E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         Top             =   4920
         Width           =   240
      End
      Begin VB.PictureBox picSearchOptions 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3360
         Picture         =   "frmMain.frx":187C8
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         Top             =   315
         Width           =   240
      End
      Begin VB.PictureBox picSearchFilter 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2640
         Picture         =   "frmMain.frx":18B52
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   4920
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picRefresh 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3000
         Picture         =   "frmMain.frx":18EDC
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   4920
         Width           =   240
      End
      Begin VB.TextBox txtSearch 
         Height          =   390
         Left            =   105
         TabIndex        =   0
         Text            =   "txtSearch"
         Top             =   240
         Width           =   2790
      End
      Begin VB.ListBox lstResults 
         Height          =   4140
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":19266
         Left            =   120
         List            =   "frmMain.frx":19268
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
      Begin VB.Image picSearch 
         Height          =   240
         Left            =   3015
         Picture         =   "frmMain.frx":1926A
         Top             =   330
         Width           =   240
      End
   End
   Begin VB.Shape shpResize 
      BorderColor     =   &H80000010&
      Height          =   6135
      Left            =   4080
      Top             =   120
      Width           =   15
   End
   Begin VB.Menu mArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mNew 
         Caption         =   "&Nuevo archivo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mEdit 
         Caption         =   "&Editar archivo"
         Shortcut        =   ^E
      End
      Begin VB.Menu sepNoseCuanto 
         Caption         =   "-"
      End
      Begin VB.Menu mSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mCambiarFuente 
         Caption         =   "&Cambiar tamaño de fuente"
      End
      Begin VB.Menu mAlternarPanelBusqueda 
         Caption         =   "&Alternar el panel de búsqueda"
         Shortcut        =   ^T
      End
      Begin VB.Menu mSiempreVisible 
         Caption         =   "&Siempre visible"
         Shortcut        =   ^S
      End
      Begin VB.Menu mSeparador 
         Caption         =   "-"
      End
      Begin VB.Menu mFuncExperimentales 
         Caption         =   "&Funciones experimentales"
         Begin VB.Menu mAgregarTagsMasivamente 
            Caption         =   "&Añadir entrada &Tags para todos los documentos"
         End
      End
   End
   Begin VB.Menu mPregunta 
      Caption         =   "?"
      Begin VB.Menu mAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu mSearchOptions 
      Caption         =   "OpcionBusqueda"
      Visible         =   0   'False
      Begin VB.Menu mSearch 
         Caption         =   "Búsqueda por contenido"
         Index           =   0
      End
      Begin VB.Menu mSearch 
         Caption         =   "Búsqueda por etiquetas"
         Index           =   1
      End
      Begin VB.Menu mSearch 
         Caption         =   "Búsqueda por nombre de archivo"
         Index           =   2
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mSearchAll 
         Caption         =   "Buscar según todos los criterios"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub cmdOk_Click()
Select Case MsgBox("¿Desea guardar los cambios?", vbQuestion + vbYesNoCancel, "Guardar cambios")
    Case vbYes
        WriteFile Config.EditPath, txtContent.Text
        EditorMode (False)
    Case vbNo
        Call lstResults_Click
        EditorMode (False)
    Case vbCancel
        Exit Sub
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call Terminar
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Activa el flag fMouseDown si aprieta el click izquierdo
With shpResize
    If X >= .Left And X <= .Left + .Width And _
        Y >= .Top And Y <= .Top + .Height Then
          If Button = vbLeftButton Then
            fMouseDown = True
          End If
    End If
End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Desactiva el flag fMouseDown si suelta el click
    fMouseDown = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Drag and Resize (?
On Error Resume Next
    If fMouseDown Then
    ' Evitemos que quiten los frames del formulario, que esté comprendido entre 10% y 90%
        If X > Me.ScaleWidth * 0.9 Or X < Me.ScaleWidth * 0.1 Then Exit Sub
        
        shpResize.Left = X
        fraSearch.Width = X - 240
        fraContent.Left = shpResize.Left + 120
        fraContent.Width = Me.ScaleWidth - fraContent.Left - 120
        
        ResizeGUI
    End If
    
With shpResize
    If X >= .Left And X <= .Left + .Width And _
        Y >= .Top And Y <= .Top + .Height Then
          Screen.MousePointer = vbSizeWE
    Else
          Screen.MousePointer = vbDefault
    End If
End With
End Sub

Private Sub fraContent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Por si se buguea y queda en modo Resize
    Screen.MousePointer = vbDefault
End Sub

Private Sub fraSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Por si se buguea y queda en modo Resize
    Screen.MousePointer = vbDefault
End Sub

Private Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
' Si aprieta Enter
    If KeyCode = vbKeyReturn Then
        Call lstResults_DblClick
    End If
End Sub

Private Sub lstResults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Por si se buguea y queda en modo Resize
    Screen.MousePointer = vbDefault
End Sub

Private Sub mAgregarTagsMasivamente_Click()
' Si, cumple lo que promete
    If MsgBox("¡Esta es una funcion experimental!" & vbNewLine & vbNewLine & _
               "Si presiona Sí, se le asignará la etiqueta 'Articulo' a cada artículo de la Wiki que no tenga etiquetas asignadas. Esta acción no modificará las etiquetas existentes." & vbNewLine & vbNewLine & _
                "Si bien no se eliminarán las entradas de archivos existentes, sí se eliminarán las etiquetas de artículos que ya no existen en la Wiki." & vbNewLine & vbNewLine & _
                "¿Está seguro de que desea continuar?", vbExclamation + vbYesNo, "Advertencia") = vbYes Then
                
                Screen.MousePointer = vbHourglass
                
                If FillTags = True Then
                    Call DeseleccionarArchivo
                    If MsgBox("La tarea finalizó correctamente." & vbNewLine & vbNewLine & "¿Desea abrir el archivo de etiquetas para agregar etiquetas manualmente?", vbInformation + vbYesNo, "¡Éxito!") = vbYes Then
                        Shell "Notepad " & Config.TagPath, vbNormalFocus
                    End If
                Else
                    MsgBox "Ocurrió un error. Compruebe que el archivo de etiquetas existe, que no es de sólo lectura y que tiene acceso a él.", vbExclamation, "¡Oh, ocurrió un error!)"
                End If
                
                Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mAlternarPanelBusqueda_Click()
' Llama al evento correspondiente en base a la vista actual
    If fTextOnlyMode Then
        Call picRevertView_Click
    Else
        Call picTextOnly_Click
    End If
End Sub

Private Sub Form_Initialize()
' Llama a la GUI del sistema
    InitCommonControls
    
' Si la ventana se cerro maximizada
        Me.Width = Val(LeerIni(Config.CfgPath, "MAIN", "frmMainWidth", "12225"))
        Me.Height = Val(LeerIni(Config.CfgPath, "MAIN", "frmMainheight", "7860"))
        
        If Val(LeerIni(Config.CfgPath, "MAIN", "frmMainMaximized", 0)) = 0 Then
            Me.WindowState = vbNormal
        Else
            Me.WindowState = vbMaximized
        End If
        
    fIsSearchFilterOn = False

End Sub

Private Sub Form_Load()
Dim sFuenteGUI As String
Dim sFuenteTam As Integer

' NADIE EMPIEZA HASTA QUE ESTEMOS LISTOS
    Do Until LeerConfig = True
        DoEvents
    Loop
    
    Me.Icon = frmSplash.Icon
    Me.Caption = "FGWiki | " & "Versión " & App.Major & "." & App.Minor & " r" & App.Revision & " (Beta)"
    mEdit.Enabled = False

' Fuente para el resto de los controles
    sFuenteGUI = LeerIni(Config.CfgPath, "MAIN", "GUIFontName", "Calibri")
    sFuenteTam = Val(LeerIni(Config.CfgPath, "MAIN", "GUIFontSize", "12"))

    Me.FontName = sFuenteGUI
    Me.FontSize = sFuenteTam
    
    txtSearch.FontName = sFuenteGUI
    txtSearch.FontSize = sFuenteTam
    
    lstResults.FontName = sFuenteGUI
    lstResults.FontSize = sFuenteTam
    
    txtTags.FontName = sFuenteGUI
    txtTags.FontSize = sFuenteTam
    
' ResizeGUI frmMain, poner los captions del menu de redimensionado, establecer busqueda, etc
    fraContent.Left = Val(LeerIni(Config.CfgPath, "MAIN", "fraContentLeft", "6120"))
    
    Check_mSearch (SearchBy)
        
    txtSearch.Text = vbNullString
    txtTags.Text = vbNullString
    
    ' I'm not ashamed of this
    Dim i As Integer
    
    With txtContent
        .Text = vbNullString
        .FontName = Config.FontName
        .FontSize = Config.FontSize
    End With
    
    mSiempreVisible.Checked = Config.AlwaysOnTop
    ComprobarVisibilidad
    
    Call FillListBox(lstResults)
    Call LoadTooltips
    Call ResizeGUI
    Call EditorMode(False)
End Sub

Sub FillListBox(lb As ListBox)
Dim i As Integer: i = -1
Dim sFName As String

' Dir$, gracias por tanto y perdón por tan poco
sFName = Dir$(Config.WikiPath & "*", vbArchive)

    While Len(sFName)
    i = i + 1
    ' Añade los archivos a la lista y guarda la ruta completa en aItemPath para que los indices coincidan
        lb.AddItem sFName
        lb.ItemData(lb.NewIndex) = i
                
        ReDim Preserve aItemPath(0 To lb.ListCount)
        aItemPath(i) = Config.WikiPath & sFName
        
    ' Sigue buscando
        sFName = Dir$()
   Wend
End Sub

Sub EditorMode(Estado As Boolean, Optional RutaCompleta As String)
    Config.EditPath = RutaCompleta

    txtContent.Locked = Not Estado
    txtTags.Visible = Not Estado
    cmdOk.Visible = Estado
    
    fraSearch.Enabled = IIf(Estado, False, True)
    fraReturn.Enabled = IIf(Estado, False, True)
    txtContent.BackColor = IIf(Estado, vbWindowBackground, vbButtonFace)
    
    If Estado Then
        If Not fTextOnlyMode Then
            Call picTextOnly_Click
        End If
        If txtContent.Text = EMPTY_FILE Then txtContent.Text = ""
        txtContent.SetFocus
    Else
        Call picRevertView_Click
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("¿Está seguro de que desea salir?", vbYesNo + vbQuestion, "Salir") = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    Call Terminar
End Sub

Private Sub lstResults_Click()
On Error GoTo oErr
Dim sFilePath As String: sFilePath = aItemPath(lstResults.ListIndex)
Dim sFileName As String: sFileName = Right$(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "\", , vbTextCompare))

' Habilitar menú edición
    mEdit.Enabled = lstResults.ListIndex <> -1

' Devuelve la ruta con el nombre de archivo (que obtiene del listbox) y la extensión
    With lstResults
        txtContent.Text = ReadFile(sFilePath)
        
        If HasTags(sFileName) Then
            txtTags.Text = GetTags(sFileName, Config.TagPath)
        Else
            txtTags.Text = "« Este archivo no contiene etiquetas. »"
        End If
    End With
    
    Exit Sub
    
oErr:

txtTags.Text = "<Error " & Err.Number & ">"

txtContent.Text = "----- ----- -----" & vbNewLine & "El archivo " & Chr(34) & sFileName & Chr(34) & " no se pudo cargar." & vbNewLine & _
                                "----- ----- -----" & vbNewLine & vbNewLine & _
                                "Error " & Err.Number & ":" & vbNewLine & _
                                Err.Description
End Sub

Private Sub lstResults_DblClick()
    If lstResults.ListIndex <> -1 Then
        If MsgBox("¿Desea editar el archivo «" & lstResults.List(lstResults.ListIndex) & "»?", vbQuestion + vbYesNo, "Editar archivo") = vbYes Then
            Call EditorMode(True, aItemPath(lstResults.ListIndex))
        End If
    End If
End Sub

Private Sub Form_Resize()
' En realidad, llama
On Error Resume Next
    
    If Me.Width < 1500 Or Me.Height < 2500 Then
        Exit Sub
    End If
        
    Call ResizeGUI
End Sub

Sub ResizeGUI()
On Error Resume Next

Dim fsHeight As Integer     ' Frame Search Height
Dim fsWidth As Integer      ' Frame Search Width
Dim fcHeight As Integer     ' Frame Content Height
Dim fcWidth As Integer      ' Frame Content Width

' Frame Busqueda
    fraSearch.Height = Me.ScaleHeight - 90
    fraSearch.Width = fraContent.Left - 360
    fsHeight = fraSearch.Height
    fsWidth = fraSearch.Width
    
    txtSearch.Width = fsWidth - 945
    lstResults.Width = fsWidth - 240
    lstResults.Height = fsHeight - 1185
    
    picSearchFilter.Top = fsHeight - 375
    picRefresh.Top = fsHeight - 375
    picTextOnly.Top = fsHeight - 375
    picTagEdit.Top = fsHeight - 375
    
    picSearchFilter.Left = fsWidth - 1095
    picRefresh.Left = fsWidth - 735
    picTextOnly.Left = fsWidth - 375
    picSearchOptions.Left = fsWidth - 375
    picSearch.Left = fsWidth - 735
    
' Frame Contenido
    fraContent.Width = Me.ScaleWidth - fraContent.Left - 90
    fraContent.Height = Me.ScaleHeight - 90
    fcWidth = fraContent.Width
    fcHeight = fraContent.Height
    
    txtTags.Width = txtTags.Width - 120
    txtTags.Width = fcWidth - 240

    txtContent.Width = fcWidth - 240
    txtContent.Height = fcHeight - 840
    
    If fTextOnlyMode Then txtContent.Visible = True
    
    lblEditorMode.Left = fraContent.Width - lblEditorMode.Width - 240
    
' Otros
    If fraSearch.Visible Then shpResize.Left = fraSearch.Left + fsWidth + 120
    shpResize.Height = Me.ScaleHeight - 240
End Sub

Private Sub mAcercaDe_Click()
    fFromFrmMain = True
    frmSplash.Show vbModal, Me
End Sub

Private Sub mCambiarFuente_Click()
Dim sFont As String
Dim fName As String
Dim fSize As String
Dim aBuffer() As String

fName = LeerIni(Config.CfgPath, "MAIN", "FontName", "Courier New")
fSize = LeerIni(Config.CfgPath, "MAIN", "FontSize", "10")

VolverAEmpezarQueAunNoTerminaElJuego:
    sFont = InputBox("Ingrese el nombre de la fuente seguido del tamaño en puntos. Separalos con una coma, por ejemplo:" & vbNewLine & vbNewLine & _
                        "Calibri, 12", "Seleccionar fuente", fName & ", " & fSize)
                        
    If Not sFont = "" Then
        sFont = RTrim(sFont)
        sFont = LTrim(sFont)
        aBuffer = Split(sFont, ",")
    
        If UBound(aBuffer) > 1 Then
        ' Validacion berreta
            MsgBox "Parece que está mal escrito." & vbNewLine & vbNewLine & "Revisa los campos o presiona cancelar.", vbExclamation
            GoTo VolverAEmpezarQueAunNoTerminaElJuego
        Else
        ' Asigna la fuente y el tamaño sin usar el CommonDialog
            txtContent.Font = aBuffer(0)
            txtContent.FontSize = Val(aBuffer(1))
            
            EscribirIni Config.CfgPath, "MAIN", "FontName", aBuffer(0)
            EscribirIni Config.CfgPath, "MAIN", "FontSize", aBuffer(1)
        End If
    Else
        Exit Sub
    End If
End Sub

Private Sub mEdit_Click()
' Doble clic
    Call lstResults_DblClick
End Sub

Private Sub mNew_Click()
    Dim ff As Integer
    Dim Nombre As String
    ff = FreeFile
    
Ingresar:
    Nombre = InputBox("Ingrese el nombre completo del nuevo archivo con su extensión, por ejemplo:" & vbNewLine & "Articulo.txt", _
                    "Editar etiquetas", "Articulo.txt")

' Si StrPtr equivale a 0& es porque apreto el boton Cancelar
    If (StrPtr(Nombre) = 0&) Then
        Exit Sub
    ElseIf Len(Trim(Nombre)) < 1 Then
        GoTo Ingresar
    ElseIf Exists(Config.WikiPath & Nombre) Then
        MsgBox "Ya existe un archivo con ese nombre. Elija otro nombre de archivo.", vbExclamation, "Error al crear el archivo"
        GoTo Ingresar
    Else
        Nombre = FileNameValid(Nombre)
        
        Open Config.WikiPath & Nombre For Output As #ff
        Close #ff
        
        Call picRefresh_Click
        Call SeleccionarArchivo(Nombre)
        Call EditorMode(True, Config.WikiPath & Nombre)
    End If
End Sub

Private Sub mSalir_Click()
' Adios
    Terminar
End Sub

Private Sub mSearchAll_Click()
    Check_mSearch (searchAll)
    SearchBy = searchAll
End Sub

Private Sub mSiempreVisible_Click()
With mSiempreVisible
    .Checked = Not .Checked
    ComprobarVisibilidad
End With
End Sub

Sub ComprobarVisibilidad()
' Llama al API de Siempre visible y guarda la configuracion
    If mSiempreVisible.Checked Then
        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        EscribirIni Config.CfgPath, "MAIN", "AlwaysOnTop", "1"
    Else
        Call SetWindowPos(Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        EscribirIni Config.CfgPath, "MAIN", "AlwaysOnTop", "0"
    End If
End Sub

Private Sub picRefresh_Click()
' Limpia los controles y vuelve a cargar los archivos
    lstResults.Clear
    txtTags.Text = ""
    txtContent.Text = ""
    
    FillListBox lstResults
End Sub

Private Sub picRevertView_Click()
' Evitar el flickering
    lstResults.Visible = True
    
' Reordenar
    fraContent.Left = shpResize.Left + 120
    
    fraReturn.Visible = False
    fraSearch.Visible = True
    shpResize.Visible = True
    
    ResizeGUI
    fTextOnlyMode = False
End Sub

Private Sub picSearch_Click()
    If txtSearch.Text = "" Then
        TurnSearchFilterOff
            Exit Sub
    Else
        sSearch txtSearch.Text, lstResults, SearchBy
    End If
End Sub

Sub TurnSearchFilterOff()
' Quita el filtro de busqueda. Duh!
    If Not fIsSearchFilterOn Then Exit Sub

        lstResults.Clear
        TurnSearchFilterOn (False)
        Call FillListBox(lstResults)
        txtSearch.Text = ""
End Sub

Private Sub picSearchFilter_Click()
' Nada
    If MsgBox("¿Desea quitar el filtro de búsqueda?", vbQuestion + vbYesNo, "Quitar filtro de búsqueda") = vbYes Then
        TurnSearchFilterOff
    End If
End Sub

Private Sub picSearchOptions_Click()
    frmMain.PopupMenu mSearchOptions
End Sub

Private Sub mSearch_Click(Index As Integer)
' Tildar y actualizar el tooltip de la búsqueda
    Check_mSearch (Index)
    SearchBy = Index
    
    picSearchOptions.ToolTipText = "(Actual: " & mSearch(Index).Caption & ") Haz clic aquí para cambiar el tipo de búsqueda"
End Sub

Sub Check_mSearch(Index As Integer)
' Marca el tick en el menu correspondiente

    mSearch(byContent).Checked = IIf(Index = byContent, True, False)
    mSearch(byTags).Checked = IIf(Index = byTags, True, False)
    mSearch(byFilename).Checked = IIf(Index = byFilename, True, False)
    mSearchAll.Checked = IIf(Index = searchAll, True, False)
End Sub

Private Sub picSearchOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Por si se buguea y queda en modo Resize
    Screen.MousePointer = vbDefault
End Sub

Private Sub picTagEdit_Click()
    If lstResults.ListIndex < 0 Then Exit Sub
    
    Dim NuevosTags As String
    Dim TagsActuales As String
    Dim sFileName As String

' Preguntar por nuevas etiquetas, por default trae las etiquetas actuales
    sFileName = lstResults.List(lstResults.ListIndex) & Config.Extension
    TagsActuales = GetTags(sFileName, Config.TagPath)
    
Hola:
    NuevosTags = InputBox("Añada o modifique las etiquetas de este archivo y recuerde separarlas por espacios. No es necesario utilizar caracteres adicionales.", _
                    "Editar etiquetas", TagsActuales)
              
    If NuevosTags = TagsActuales Then Exit Sub

' StrPtr es la posta, si equivale a 0& es porque apreto el boton Cancelar
    If (StrPtr(NuevosTags) = 0&) Then
        Exit Sub
    ElseIf NuevosTags = "" Then
        If MsgBox("Esta acción borrará todas las etiquetas del artículo." & vbNewLine & vbNewLine & _
                    "¿Desea continuar?", vbQuestion + vbYesNo, "Advertencia") = vbNo Then
                        GoTo Hola
        End If
    End If
    
    SetTags sFileName, NuevosTags
        
        Call lstResults_Click
        
End Sub

Private Sub picTextOnly_Click()
' Evitar el flickering
    lstResults.Visible = False
    txtContent.Visible = False
    
' Reordenar
    fraContent.Left = fraReturn.Width + fraReturn.Left + 120
    fraContent.Width = Me.ScaleWidth - fraContent.Left - 120
    
    fraSearch.Visible = False
    fraReturn.Visible = True
    shpResize.Visible = False
    fTextOnlyMode = True
    
    fraReturn.Left = 120
    fraReturn.Top = 0
    fraReturn.Height = Me.ScaleHeight - 120
    
    picRevertView.Top = fraReturn.Height - 375
    
    ResizeGUI
End Sub

Private Sub txtContent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Por si se buguea y queda en modo Resize
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0    ' Silenciar beep
        Call picSearch_Click
    End If
End Sub

Sub sSearch(Que As String, lb As ListBox, Optional sSearchBy As Integer = byContent)
Dim sFName As String
Dim FoundSomething As Boolean
Dim i As Integer

' Limpiar controles
    lb.Clear
    txtTags.Text = vbNullString
    txtContent.Text = vbNullString
    FoundSomething = False
    i = -1

sFName = Dir$(Config.WikiPath & "*", vbArchive)

Do While Len(sFName) > 0
    Dim sSearch As String
    Dim sName As String
    
    Select Case sSearchBy
        Case byTags
            sSearch = GetTags(sFName, Config.TagPath)
            sName = InStr(1, sSearch, Que, vbTextCompare)
        
        Case byContent
            sSearch = ReadFile(Config.WikiPath & sFName)
            sName = InStr(1, sSearch, Que, vbTextCompare)
        
        Case byFilename
            sSearch = sFName
            sName = InStr(1, sSearch, Que, vbTextCompare)
        
        Case searchAll
            sSearch = GetTags(sFName, Config.TagPath)
            sName = InStr(1, sSearch, Que, vbTextCompare)
            If sName < 1 Then
                sSearch = ReadFile(Config.WikiPath & sFName)
                sName = InStr(1, sSearch, Que, vbTextCompare)

                If sName < 1 Then
                    sSearch = sFName
                    sName = InStr(1, sSearch, Que, vbTextCompare)
                End If
            End If
    End Select
    
        If sName > 0 Then
        i = i + 1
            FoundSomething = True
            lb.AddItem sFName 'Left$(sFName, InStrRev(sFName, ".") - 1)
            lb.ItemData(lb.NewIndex) = i
                
            ReDim Preserve aItemPath(0 To lb.ListCount)
            aItemPath(i) = Config.WikiPath & sFName
        End If

    sFName = Dir$()
Loop

If FoundSomething Then
    TurnSearchFilterOn (True)
Else
    TurnSearchFilterOn (False)
    Call FillListBox(lb)
    lstResults.ListIndex = -1

    Select Case sSearchBy
        Case byTags
            txtContent.Text = "----- ----- -----" & vbNewLine & "No se encontraron coincidencias de etiquetas para «" & Que & "»." & vbNewLine & "----- ----- -----"
    
        Case byContent
            txtContent.Text = "----- ----- -----" & vbNewLine & "No se encontraron coincidencias de contenido para «" & Que & "»." & vbNewLine & "----- ----- -----"
        
        Case byFilename
            txtContent.Text = "----- ----- -----" & vbNewLine & "No se encontraron coincidencias de nombre de archivo para «" & Que & "»." & vbNewLine & "----- ----- -----"
        
        Case searchAll
            txtContent.Text = "----- ----- -----" & vbNewLine & "No se encontraron coincidencias de ningún criterio para «" & Que & "»." & vbNewLine & "----- ----- -----"
    End Select
End If
    
End Sub

Sub TurnSearchFilterOn(Estado As Boolean)
    fIsSearchFilterOn = Estado
    lstResults.ForeColor = IIf(Estado, vbHighlight, vbWindowText)

With picSearchFilter
    .Visible = Estado
    .Enabled = Estado
End With
End Sub

Sub LoadTooltips()
' Cargar tooltips
    picTextOnly.ToolTipText = "Esconde el panel de búsqueda y lista de archivos (Ctrl+T)"
    picRevertView.ToolTipText = "Muestra el panel de búsqueda y lista de archivos (Ctrl+T)"
    picRefresh.ToolTipText = "Actualiza la lista de archivos"
    picSearchFilter.ToolTipText = "Desactiva el filtro de búsqueda"
    picTagEdit.ToolTipText = "Añade o modifica las etiquetas para el archivo seleccionado"
    picSearch.ToolTipText = "Buscar"
    
    If SearchBy = searchAll Then
        picSearchOptions.ToolTipText = "(Actual: " & mSearchAll.Caption & ") Haz clic aquí para cambiar el tipo de búsqueda"
    Else
        picSearchOptions.ToolTipText = "(Actual: " & mSearch(SearchBy).Caption & ") Haz clic aquí para cambiar el tipo de búsqueda"
    End If
End Sub

Sub Terminar()
' Guarda las configuraciones especificas ya que no se cambiaran
    EscribirIni Config.CfgPath, "MAIN", "LastBarChecked", LastChecked
    EscribirIni Config.CfgPath, "MAIN", "frmMainMaximized", apiIsZoomed(Me.hwnd)
    EscribirIni Config.CfgPath, "MAIN", "SearchBy", SearchBy
    
    If apiIsZoomed(Me.hwnd) = 0 Then
    ' Solo si no esta maximizada (apiIsZoomed = 0) guardar las dimensiones del Form
        EscribirIni Config.CfgPath, "MAIN", "frmMainWidth", Me.Width
        EscribirIni Config.CfgPath, "MAIN", "frmMainHeight", Me.Height
    End If
    
    If Not fTextOnlyMode Then
    ' Si esta en modo solo texto no tiene sentido guardar la coordenada del Frame Contenido
        EscribirIni Config.CfgPath, "MAIN", "fraContentLeft", fraContent.Left
    End If
    
    Unload frmMain
    Unload frmSplash
End Sub

Private Sub txtTags_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Por si se buguea y queda en modo Resize
    Screen.MousePointer = vbDefault
End Sub

Sub DeseleccionarArchivo()
    lstResults.ListIndex = -1
    txtContent.Text = ""
    txtTags.Text = ""
End Sub

Sub SeleccionarArchivo(Archivo As String)
Dim i As Integer
    For i = 0 To lstResults.ListCount - 1
        If lstResults.List(i) = Archivo Then
            lstResults.ListIndex = i
            Call lstResults_Click
            Exit For
        End If
    Next
End Sub
