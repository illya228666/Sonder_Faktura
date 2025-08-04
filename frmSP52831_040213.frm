VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Object = "{2213E283-16BC-101D-AFD4-040224009C08}#8.0#0"; "CM32L8O.OCX"
Begin VB.Form frmSP52831 
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   1035
   ClientTop       =   2580
   ClientWidth     =   11535
   Icon            =   "frmSP52831.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   11535
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      Height          =   60
      Index           =   1
      Left            =   9930
      ScaleHeight     =   60
      ScaleWidth      =   4800
      TabIndex        =   5
      Top             =   9090
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      Height          =   60
      Index           =   0
      Left            =   9900
      ScaleHeight     =   60
      ScaleWidth      =   4800
      TabIndex        =   4
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   9990
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmSP52831.frx":030A
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   4785
      Index           =   0
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9795
      _Version        =   393216
      _ExtentX        =   17277
      _ExtentY        =   8440
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      ColHeaderDisplay=   1
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      GridColor       =   -2147483633
      RetainSelBlock  =   0   'False
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmSP52831.frx":0310
      UnitType        =   2
      ScrollBarTrack  =   3
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2595
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   9795
      _Version        =   393216
      _ExtentX        =   17277
      _ExtentY        =   4577
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      ColHeaderDisplay=   1
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      GridColor       =   -2147483633
      RetainSelBlock  =   0   'False
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmSP52831.frx":4816
      UnitType        =   2
      ScrollBarTrack  =   3
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2595
      Index           =   2
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7380
      Width           =   9795
      _Version        =   393216
      _ExtentX        =   17277
      _ExtentY        =   4577
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      ColHeaderDisplay=   1
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      GridColor       =   -2147483633
      RetainSelBlock  =   0   'False
      ScrollBarShowMax=   0   'False
      SpreadDesigner  =   "frmSP52831.frx":8D1C
      UnitType        =   2
      ScrollBarTrack  =   3
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Unten ausrichten
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   10020
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   635
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd1 
         Caption         =   "Löschen"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Speichern"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1350
         TabIndex        =   11
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Designer"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   2670
         TabIndex        =   10
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Vorschau"
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3990
         TabIndex        =   9
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Drucken"
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   5310
         TabIndex        =   8
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Schließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   6630
         TabIndex        =   7
         Top             =   30
         Width           =   1260
      End
   End
   Begin ListLabel.ListLabel LL1 
      Left            =   10980
      Top             =   30
      _Version        =   65537
      _ExtentX        =   714
      _ExtentY        =   661
      _StockProps     =   64
      Language        =   0
      DialogMode      =   9
      DialogFrame     =   0
      Dialog3DText    =   1
      DialogButtons   =   1
      NewExpressions  =   1
      TableColoring   =   0
      TabStops        =   0
      EnablePageCallback=   -1  'True
      EnableProjectCallback=   -1  'True
      EnableObjectCallback=   -1  'True
      EnableHelpCallback=   -1  'True
      OnlyOneTable    =   0   'False
      MultipleTableLines=   -1  'True
      SortVariables   =   -1  'True
      HelpAvailable   =   -1  'True
      SupportPageBreak=   0   'False
      ShowPredefVars  =   -1  'True
      UseHostprinter  =   0   'False
      EMFResolution   =   0
      AddVarsToFields =   0   'False
      ConvertCRLF     =   0   'False
      WizFileNew      =   0   'False
      VarsCaseSensitive=   -1  'True
      RealTime        =   0   'False
      SpaceOptimization=   -1  'True
      CompressStorage =   0   'False
      NoParameterCheck=   0   'False
      NoNoTableCheck  =   0   'False
      PreviewZoomPerc =   100
      PreviewRectLeft =   0
      PreviewRectTop  =   0
      PreviewRectWidth=   0
      PreviewRectHeight=   0
      Metric          =   1
      TabRepresentationCode=   247
      RetRepresentationCode=   182
      StorageSystem   =   1
      AutoMultipage   =   -1  'True
      UseBarcodeSizes =   0   'False
      MaxRTFVersion   =   256
      DelayTableHeader=   0   'False
      OfnDialogExplorer=   -1  'True
      CreateInfo      =   -1  'True
      XlatVarNames    =   -1  'True
      PhantomSpaceRepresentationCode=   2
      LockNextCharRepresentationCode=   3
      ExprSepRepresentationCode=   164
      TextQuoteRepresentationCode=   1
      InterCharSpacing=   0   'False
      IncludeFontDescent=   0   'False
      AllowMenuManager=   -1  'True
      UseChartFields  =   0   'False
      Dummy1          =   -1  'True
      Dummy2          =   -1  'True
      Dummy3          =   -1  'True
      Dummy4          =   -1  'True
      Dummy5          =   -1  'True
   End
   Begin VB.Image imgSplitter 
      Height          =   60
      Index           =   1
      Left            =   9930
      MousePointer    =   7  'Größenänderung N S
      Top             =   9210
      Width           =   4785
   End
   Begin VB.Image imgSplitter 
      Height          =   60
      Index           =   0
      Left            =   9900
      MousePointer    =   7  'Größenänderung N S
      Top             =   660
      Width           =   4785
   End
End
Attribute VB_Name = "frmSP52831"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Public frmParent As Form
  Private gFPArray() As Variant
  Private gFPArray1() As Variant
  Private gbZeileInCopy As Boolean
  
  Private objDAusw As SPDAOAusw.clsDAOAuswahl
  Private objDAuswDef As SPDAOAusw.clsDAOAuswahl
  Private objHlp As SpHlp.clsHlp
  
  Private gsngKopfHeight As Single
  Private gsngPostenHeight As Single
  Private gsngFussHeight As Single
  Private gsngFussTop As Single
  Private gbAfterLoad As Boolean
  
  Private COL_MENGE_STR As String 'Spalten-Kennung = Chr(64 + COL_MENGE)
  Private COL_EINHEIT_STR As String
  Private COL_EPREIS_STR As String
  Private COL_RABATT_STR As String
  Private COL_GPREIS_STR As String
  Private COL_UST_STR As String
  Private COL_GPREISDUMMY_O_UST_STR As String
  Private COL_GPREISDUMMY_M_UST_STR As String

  Const COL_ZEILENART = 1
  Const COL_ARTSCHL = 2
  Const COL_ARTIKEL = 3
  Const COL_MENGE = 4
  Const COL_EINHEIT = 5
  Const COL_EPREIS = 6
  Const COL_RABATT = 7
  Const COL_GPREIS = 8
  
  Const COL_UST = 9
  Const COL_KOSTSCHL = 10
  Const COL_SACHSCHL = 11
  Const COL_KOSTKTO = 12
  Const COL_SACHKTO = 13

  Const COL_GPREISDUMMY = 14
  Const COL_GPREISDUMMY_O_UST = 15
  Const COL_GPREISDUMMY_M_UST = 16
  Const COL_SUMMEN = 17
  Const COL_ERSTDAT = 18
  Const COL_ERSTVON = 19
  Const COL_AENDDAT = 20
  Const COL_AENDVON = 21
  
  Const COL_LASTEDIT = 13
  
  Const COL_FS_STEUERPFLICHT = 2
  Const COL_FS_UST = 3
  Const COL_FS_STEUERFREI = 4
  Const COL_FS_WRG = 5
  Const COL_FS_GESAMT = 6
  


Private Sub cmd1_Click(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Select Case Index
  Case 2 'Speichern
    If LastRow > 0 Then
      If GlngBelegID = 0 Then
        frmParent.Speichern
      Else
        frmParent.Speichern GlngBelegID
      End If
    Else
      MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gespeichert werden.", vbInformation
      Schalter False
    End If
  Case 3 'Designer
    If LastRow > 0 Then
      If GlngBelegID = 0 Then
        If frmParent.Speichern(GlngArbeitsplatz, True) Then
          LLDesigner Me, LL1, GlngArbeitsplatz, 1, True
        End If
      Else
        If GintDruck = 1 Then
          LLDesigner Me, LL1, GlngBelegID, 1
        Else
          If frmParent.Speichern(GlngBelegID) Then
            LLDesigner Me, LL1, GlngBelegID, 1
          End If
        End If
      End If
    Else
      MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gedruckt werden.", vbInformation
      Schalter False
    End If
  Case 4 'Vorschau
    If LastRow > 0 Then
      If GlngBelegID = 0 Then
        If frmParent.Speichern(GlngArbeitsplatz, True) Then
          LLPrintListe Me, LL1, GlngArbeitsplatz, 2, True
        End If
      Else
        If GintDruck = 1 Then
          LLPrintListe Me, LL1, GlngBelegID, 2
        Else
          If frmParent.Speichern(GlngBelegID) Then
            LLPrintListe Me, LL1, GlngBelegID, 2
          End If
        End If
      End If
    Else
      MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gedruckt werden.", vbInformation
      Schalter False
    End If
  Case 5 'Drucken
    If LastRow > 0 Then
      If GlngBelegID = 0 Then
        'Bevor der Beleg gedruckt werden kann, muss er gespeichert und die BelegID vergeben werden.
        If frmParent.Speichern(, , True) Then
          LLPrintListe Me, LL1, GlngBelegID, 1
        End If
      Else
        If GintDruck = 1 Then
          LLPrintListe Me, LL1, GlngBelegID, 1
        Else
          If frmParent.Speichern(GlngBelegID, , True) Then
            LLPrintListe Me, LL1, GlngBelegID, 1
          End If
        End If
      End If
    Else
      MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gedruckt werden.", vbInformation
      Schalter False
    End If
  Case 6 'Schließen
    Unload Me
  End Select
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "cmd1_Click")
        '***Ende
End Sub

Private Sub Form_Load()
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim i As Integer
  Dim h As Integer
  
  LL1.LlSetOption LL_OPTION_CONVERTCRLF, True  'Verhindern der doppelten Zeilenumbrüche.
  LL1.LlSetPrinterDefaultsDir ArbeitsplatzPfad 'Pfad für Drucker-Einstellungen setzen.
  
  gbAfterLoad = False
  
  Me.left = 500
  Me.top = 500
  Me.Caption = frmParent.Frame1(2).Caption
  
  COL_MENGE_STR = Chr(64 + COL_MENGE)
  COL_EINHEIT_STR = Chr(64 + COL_EINHEIT)
  COL_EPREIS_STR = Chr(64 + COL_EPREIS)
  COL_RABATT_STR = Chr(64 + COL_RABATT)
  COL_GPREIS_STR = Chr(64 + COL_GPREIS)
  COL_UST_STR = Chr(64 + COL_UST)
  COL_GPREISDUMMY_O_UST_STR = Chr(64 + COL_GPREISDUMMY_O_UST)
  COL_GPREISDUMMY_M_UST_STR = Chr(64 + COL_GPREISDUMMY_M_UST)
  
  KopfSetUp
  KopfFuellen
  For i = 1 To fpSpread1(0).MaxRows
    h = h + fpSpread1(0).RowHeight(i)
  Next i
  gsngKopfHeight = h + 300
  fpSpread1(0).Height = gsngKopfHeight
  
  fpSpread1(1).top = gsngKopfHeight
  gsngPostenHeight = 4300
  fpSpread1(1).Height = gsngPostenHeight
  PostenSetUp
  h = 0
  For i = 1 To fpSpread1(1).MaxCols
    h = h + fpSpread1(1).ColWidth(i)
  Next i
  Me.Width = h + 580
  
  FussSetUp
  FussFuellen
  h = 0
  For i = 1 To fpSpread1(2).MaxRows
    h = h + fpSpread1(2).RowHeight(i)
  Next i
  gsngFussHeight = h + 400
  fpSpread1(2).Height = gsngFussHeight
  fpSpread1(2).top = fpSpread1(1).top + fpSpread1(1).Height
  
  gsngFussTop = fpSpread1(2).top
  gsngFussHeight = fpSpread1(2).Height
  
  'SizeControls fpSpread1(0).Height, 0
  'SizeControls fpSpread1(2).top, 1
  
  'Me.Height = gsngFussTop + gsngFussHeight + SSPanel1(0).Height + 450
  
  SizeControls GetSetting("SP50000", "SP52800", "SP52831Split0", fpSpread1(0).Height), 0
  SizeControls GetSetting("SP50000", "SP52800", "SP52831Split1", fpSpread1(2).top), 1
  
  Me.Width = GetSetting("SP50000", "SP52800", "SP52831Width", "10935")
  Me.Height = GetSetting("SP50000", "SP52800", "SP52831Height", gsngFussTop + gsngFussHeight + SSPanel1(0).Height + 450)
  Me.left = GetSetting("SP50000", "SP52800", "SP52831Left", "6850")
  Me.top = GetSetting("SP50000", "SP52800", "SP52831Top", "870")
  
  
  
'  imgSplitter(0).Width = h
  imgSplitter(0).left = 0
'  imgSplitter(0).top = gsngKopfHeight
'  imgSplitter(1).Width = h
  imgSplitter(1).left = 0
'  imgSplitter(1).top = gsngKopfHeight + gsngPostenHeight
'
'  picSplitter(0).Width = h
  picSplitter(0).left = 0
'  picSplitter(0).top = gsngKopfHeight
'  picSplitter(1).Width = h
  picSplitter(1).left = 0
'  picSplitter(1).top = gsngKopfHeight + gsngPostenHeight
  
  txt1(0).Width = TEXT_BREITE

  Set objDAusw = New SPDAOAusw.clsDAOAuswahl
  objDAusw.DatabaseName = GsHauptPfad & "dat\" & CStr(CInt(GsAnwenderNr)) & "\SP50000.dat"
  objDAusw.FilterBar = True
  
  Set objDAuswDef = New SPDAOAusw.clsDAOAuswahl
  objDAuswDef.DatabaseName = GsHauptPfad & "exe\SP50000.def"
  gbAfterLoad = True
  
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "Form_Load")
        '***Ende
End Sub

Sub SizeControls(Y As Single, SpliterIndex As Integer)
    On Error Resume Next
    Dim Dif As Integer

    imgSplitter(SpliterIndex).top = Y
    
    If SpliterIndex = 0 Then
      fpSpread1(SpliterIndex).Height = Y
      
      Dif = fpSpread1(SpliterIndex + 1).top - Y
      fpSpread1(SpliterIndex + 1).top = Y + picSplitter(SpliterIndex).Height
      fpSpread1(SpliterIndex + 1).Height = fpSpread1(SpliterIndex + 1).Height + Dif - picSplitter(SpliterIndex).Height
    Else
      fpSpread1(SpliterIndex).Height = Y - fpSpread1(SpliterIndex).top
      
      Dif = fpSpread1(SpliterIndex + 1).top - Y
      fpSpread1(SpliterIndex + 1).top = Y + picSplitter(SpliterIndex).Height
      fpSpread1(SpliterIndex + 1).Height = fpSpread1(SpliterIndex + 1).Height + Dif - picSplitter(SpliterIndex).Height
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  
  If Me.WindowState = 0 Then
    SaveSetting "SP50000", "SP52800", "SP52831Left", Me.left
    SaveSetting "SP50000", "SP52800", "SP52831Top", Me.top
    SaveSetting "SP50000", "SP52800", "SP52831Width", Me.Width
    SaveSetting "SP50000", "SP52800", "SP52831Height", Me.Height
    
    SaveSetting "SP50000", "SP52800", "SP52831Split0", imgSplitter(0).top
    SaveSetting "SP50000", "SP52800", "SP52831Split1", imgSplitter(1).top
  End If
  
  If LastRow > 0 Then
    If cmd1(2).Enabled Then
      If MsgBox("Sie haben Daten erfasst und nicht gespeichert. Sollen sie jetzt gespeichert werden?", vbYesNo + vbQuestion) = vbYes Then
        If GlngBelegID = 0 Then
          frmParent.Speichern
        Else
          frmParent.Speichern GlngBelegID
        End If
      End If
    End If
  End If
  GlngBelegID = 0

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "Form_QueryUnload")
        '***Ende
End Sub

Private Sub Form_Resize()

On Error GoTo Fehler


  If gbAfterLoad Then
    fpSpread1(0).Width = Me.ScaleWidth
    fpSpread1(1).Width = Me.ScaleWidth
    fpSpread1(2).Width = Me.ScaleWidth
  
    picSplitter(0).Width = Me.ScaleWidth
    picSplitter(1).Width = Me.ScaleWidth
  
    imgSplitter(0).Width = Me.ScaleWidth
    imgSplitter(1).Width = Me.ScaleWidth
    gsngFussTop = Me.ScaleHeight - gsngFussHeight - SSPanel1(0).Height

    SizeControls gsngFussTop, 1
    fpSpread1(2).Height = Me.ScaleHeight - gsngFussTop - SSPanel1(0).Height - imgSplitter(1).Height
    If Me.Height <= 7000 Then
      Me.Height = 7000
    End If
  End If
  
        Exit Sub
Fehler:
  
  If Err.Number = 384 Then
    'Das Formular kann nicht verschoben oder in der Größe geändert werden, während es minimiert oder maximiert ist
    Resume Next
  Else
    Call FehlerErklärung("frmSP52831", "Form_Resize")
  End If
End Sub


Sub KopfSetUp()
    'Lock the entire control
    fpSpread1(0).Row = -1
    fpSpread1(0).Col = -1
    fpSpread1(0).Lock = True
    'Zellen sofort in EditMode Versetzen. (Markieren der Zellen ist in dem Modus nicht möglich.)
    fpSpread1(0).EditModePermanent = True
    
    fpSpread1(0).TypeMaxEditLen = 110
    '
    fpSpread1(0).MaxCols = 8
    fpSpread1(0).MaxRows = 19
    'Change font size
    fpSpread1(0).Row = -1
    fpSpread1(0).Col = -1
    fpSpread1(0).FontSize = 9
    'Turn off grid lines
    fpSpread1(0).GridShowHoriz = False
    fpSpread1(0).GridShowVert = False
    'Allow cell contents to overflow into adjacent cells
    fpSpread1(0).AllowCellOverflow = True
    
    'Allow the tab key to operate
    'fpSpread1(0).ProcessTab = True
    
    'Changes its size to the number of specified columns and rows
    'fpSpread1(0).AutoSize = True
    
    'Turn off headers
    fpSpread1(0).ColHeadersShow = False
    fpSpread1(0).RowHeadersShow = False
    'Highlight entire cell contents when clicked in
    'fpSpread1(0).EditModeReplace = True
    'Set up col widths
    fpSpread1(0).ColWidth(1) = 1120 '
    fpSpread1(0).ColWidth(2) = 1400 '
    fpSpread1(0).ColWidth(3) = 1400 '
    fpSpread1(0).ColWidth(4) = 1400 '
    fpSpread1(0).ColWidth(5) = 1400 '
    fpSpread1(0).ColWidth(6) = 1400 '
    fpSpread1(0).ColWidth(7) = 1400 '
    fpSpread1(0).ColWidth(8) = 950 '
    
    'Set row heights
    fpSpread1(0).RowHeight(6) = 600
    fpSpread1(0).RowHeight(7) = 300
    fpSpread1(0).RowHeight(11) = 400
    fpSpread1(0).RowHeight(14) = 400
    fpSpread1(0).RowHeight(15) = 400
    fpSpread1(0).RowHeight(16) = 300
    fpSpread1(0).RowHeight(17) = 300
    
    'Change font size
    fpSpread1(0).Row = 7
    fpSpread1(0).Col = 2
    fpSpread1(0).FontSize = 8
    
    fpSpread1(0).Row = 15
    fpSpread1(0).Col = 5
    fpSpread1(0).FontSize = 11
    fpSpread1(0).FontBold = True
    
    'Nicht druckbaren Spalten farblich kennzeichnen
    fpSpread1(0).Col = 1
    fpSpread1(0).Row = 1
    fpSpread1(0).Col2 = 1
    fpSpread1(0).Row2 = fpSpread1(0).MaxRows
    fpSpread1(0).BlockMode = True
    fpSpread1(0).BackColor = vbButtonFace
    fpSpread1(0).BlockMode = False
    
End Sub

Sub FussSetUp()
    'Lock the entire control
    fpSpread1(2).Row = -1
    fpSpread1(2).Col = -1
    fpSpread1(2).Lock = True
    fpSpread1(2).TypeHAlign = TypeHAlignRight
    
    'Zellen sofort in EditMode Versetzen. (Markieren der Zellen ist in dem Modus nicht möglich.)
    fpSpread1(2).EditModePermanent = True
    
    fpSpread1(2).TypeMaxEditLen = 110
    '
    fpSpread1(2).MaxCols = 6
    fpSpread1(2).MaxRows = 6
    'Change font size
    fpSpread1(2).Row = -1
    fpSpread1(2).Col = -1
    fpSpread1(2).FontSize = 9
    'Turn off grid lines
    fpSpread1(2).GridShowHoriz = False
    fpSpread1(2).GridShowVert = False
    
    'Allow cell contents to overflow into adjacent cells
    fpSpread1(2).AllowCellOverflow = True
    
    'Allow the tab key to operate
    'fpSpread1(2).ProcessTab = True
    
    'Changes its size to the number of specified columns and rows
    'fpSpread1(2).AutoSize = True
    
    'Turn off headers
    fpSpread1(2).ColHeadersShow = False
    fpSpread1(2).RowHeadersShow = False
    
    'Highlight entire cell contents when clicked in
    'fpSpread1(2).EditModeReplace = True
    'Set up col widths
    fpSpread1(2).ColWidth(1) = 1120
    fpSpread1(2).ColWidth(COL_FS_STEUERPFLICHT) = 2180
    fpSpread1(2).ColWidth(COL_FS_UST) = 2180
    fpSpread1(2).ColWidth(COL_FS_STEUERFREI) = 2180
    fpSpread1(2).ColWidth(COL_FS_WRG) = 600 'Währung
    fpSpread1(2).ColWidth(COL_FS_GESAMT) = 2180
    
    fpSpread1(2).Col = COL_FS_STEUERPFLICHT
    fpSpread1(2).Row = 2
    fpSpread1(2).Col2 = COL_FS_STEUERFREI
    fpSpread1(2).Row2 = 3
    fpSpread1(2).BlockMode = True
    fpSpread1(2).CellType = CellTypeNumber  'Integer
    fpSpread1(2).TypeNumberDecPlaces = 2
    fpSpread1(2).BlockMode = False
    
    fpSpread1(2).Col = COL_FS_GESAMT
    fpSpread1(2).Row = 2
    fpSpread1(2).Col2 = COL_FS_GESAMT
    fpSpread1(2).Row2 = 3
    fpSpread1(2).BlockMode = True
    fpSpread1(2).CellType = CellTypeNumber  'Integer
    fpSpread1(2).TypeNumberDecPlaces = 2
    fpSpread1(2).BlockMode = False
    
    'UST
    fpSpread1(2).Col = COL_FS_UST
    fpSpread1(2).Row = 2 'Summe
    fpSpread1(2).Formula = "B2*" & SQLZahl(frmParent.txt1(20)) & "/100"
    'fpSpread1(2).Formula = "B2*16.00/100"
    
    'Rechnungssumme
    fpSpread1(2).Col = COL_FS_GESAMT
    fpSpread1(2).Row = 2 'Summe
    fpSpread1(2).Formula = "B2+C2+D2"
    
    If Trim(frmParent.Check1(1)) = 1 And Trim(UCase(frmParent.txt1(17))) <> Trim(UCase(frmParent.txt1(18))) Then
    'If Trim(frmParent.txt1(18)) <> "" Then
      '2 Währung wird ausgewiesen
      
      fpSpread1(2).Col = COL_FS_STEUERPFLICHT
      fpSpread1(2).Row = 3 'Summe
      fpSpread1(2).Formula = "B2*" & SQLZahl(frmParent.txt1(19)) & ""
      'fpSpread1(2).Formula = "B2*0.6076"
      
      fpSpread1(2).Col = COL_FS_UST
      fpSpread1(2).Row = 3 'Summe
      fpSpread1(2).Formula = "C2*" & SQLZahl(frmParent.txt1(19)) & ""
      'fpSpread1(2).Formula = "C2*0.6076"
      
      fpSpread1(2).Col = COL_FS_STEUERFREI
      fpSpread1(2).Row = 3 'Summe
      fpSpread1(2).Formula = "D2*" & SQLZahl(frmParent.txt1(19)) & ""
      'fpSpread1(2).Formula = "D2*0.6076"
      
      'Rechnungssumme
      fpSpread1(2).Col = COL_FS_GESAMT
      fpSpread1(2).Row = 3 'Summe
      fpSpread1(2).Formula = "B3+C3+D3"
      fpSpread1(2).RowHeight(3) = 300
    Else
      fpSpread1(2).RowHeight(3) = 0
      fpSpread1(2).RowHeight(2) = 300
    End If
    fpSpread1(2).RowHeight(4) = 300
    fpSpread1(2).RowHeight(5) = 220
    fpSpread1(2).RowHeight(6) = 220
    
    
    fpSpread1(2).Col = 1
    fpSpread1(2).Row = 1
    fpSpread1(2).Col2 = fpSpread1(0).MaxCols
    fpSpread1(2).Row2 = 1
    fpSpread1(2).BlockMode = True
    fpSpread1(2).FontBold = True
    fpSpread1(2).BlockMode = False
    
    'Unterste Zeilen links ausrichten
    fpSpread1(2).Col = COL_FS_STEUERPFLICHT
    fpSpread1(2).Row = 4
    fpSpread1(2).Col2 = COL_FS_GESAMT
    fpSpread1(2).Row2 = 6
    fpSpread1(2).BlockMode = True
    'fpSpread1(2).TypeEditMultiLine = True
    fpSpread1(2).CellType = CellTypeStaticText
    fpSpread1(2).TypeHAlign = TypeHAlignLeft
    fpSpread1(2).BlockMode = False
    
    'Nicht druckbaren Spalten farblich kennzeichnen
    fpSpread1(2).Col = 1
    fpSpread1(2).Row = 1
    fpSpread1(2).Col2 = 1
    fpSpread1(2).Row2 = fpSpread1(0).MaxRows
    fpSpread1(2).BlockMode = True
    fpSpread1(2).BackColor = vbButtonFace
    fpSpread1(2).BlockMode = False
    
    'Rahmen hinzufügen
    fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 1, fpSpread1(2).MaxCols, fpSpread1(2).MaxRows, 16, vbBlack, CellBorderStyleSolid 'Ganzer Bereich
    fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 4, fpSpread1(2).MaxCols, fpSpread1(2).MaxRows - 1, 4, vbBlack, CellBorderStyleSolid 'Oben
    fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 1, COL_FS_STEUERFREI, 3, 2, vbBlack, CellBorderStyleSolid 'Rechts

End Sub


Public Sub KopfFuellen()
        '***Beginn
On Error GoTo Fehler
        '***Ende
    fpSpread1(0).SetText 2, 1, GmandantRS!Name1
    fpSpread1(0).SetText 2, 2, GmandantRS!Name2
    fpSpread1(0).SetText 2, 3, GmandantRS!Straße
    fpSpread1(0).SetText 2, 4, GmandantRS!Lkz & "-" & GmandantRS!plz & " " & GmandantRS!ort & " " & GmandantRS!Ortsteil
    fpSpread1(0).SetText 2, 5, "Tel. " & GmandantRS!Telefon
    fpSpread1(0).SetText 2, 6, "Fax " & GmandantRS!Fax
    fpSpread1(0).SetText 2, 7, GmandantRS!Name1 & ", " & GmandantRS!Name2 & ", " & GmandantRS!Name2 & ", " & GmandantRS!Lkz & "-" & GmandantRS!plz & " " & GmandantRS!ort & " " & GmandantRS!Ortsteil
    fpSpread1(0).SetText 2, 8, frmParent.txt1(1).Text
    fpSpread1(0).SetText 2, 9, frmParent.txt1(2).Text
    
    If Trim(frmParent.txt1(4).Text) <> "" Then
      'Postfach
      fpSpread1(0).SetText 2, 11, "Postfach " & frmParent.txt1(4).Text
      fpSpread1(0).SetText 2, 12, frmParent.txt1(8).Text & "-" & frmParent.txt1(5).Text & " " & frmParent.txt1(6).Text
    Else
      fpSpread1(0).SetText 2, 11, frmParent.txt1(7).Text
      fpSpread1(0).SetText 2, 12, frmParent.txt1(8).Text & "-" & frmParent.txt1(9).Text & " " & frmParent.txt1(10).Text & " " & frmParent.txt1(11).Text
    End If
    
    If Trim(frmParent.txt1(3).Text) <> "" Then
      fpSpread1(0).SetText 2, 10, frmParent.txt1(3).Text
    Else
      fpSpread1(0).RowHeight(10) = 0
    End If
    
    fpSpread1(0).SetText 7, 13, "UID-Nr.: " & GmandantRS!UID
    fpSpread1(0).SetText 7, 14, "USt-Nr.: " & GmandantRS!SteuerNr
    fpSpread1(0).SetText 5, 15, frmParent.Frame1(2).Caption '"Rechnung"
    
    fpSpread1(0).Col = 1
    fpSpread1(0).Row = 16
    fpSpread1(0).Col2 = fpSpread1(0).MaxCols
    fpSpread1(0).Row2 = 16
    fpSpread1(0).BlockMode = True
    fpSpread1(0).FontBold = True
    fpSpread1(0).BlockMode = False
    
    fpSpread1(0).SetText 2, 16, "Kunden-Nr."
    fpSpread1(0).SetText 3, 16, "Beleg-Nr."
    fpSpread1(0).SetText 4, 16, "Beleg-Datum"
    fpSpread1(0).SetText 5, 16, "Bearbeiter"
    
    fpSpread1(0).SetText 2, 17, frmParent.txt1(12).Text
    fpSpread1(0).SetText 3, 17, frmParent.txt1(15).Text
    fpSpread1(0).SetText 4, 17, frmParent.txt1(16).Text
    fpSpread1(0).SetText 5, 17, GsUser
    
    fpSpread1(0).SetText 2, 18, "Bei Zahlungen bitte unbedingt angeben!"
    
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "KopfFuellen")
        '***Ende
End Sub
Public Sub FussFuellen()
        '***Beginn
On Error GoTo Fehler
        '***Ende

  fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 1, "Steuer-Pflichtig"
  fpSpread1(2).SetText COL_FS_UST, 1, "Umsatz-Steuer " & frmParent.txt1(20) & "%"
  fpSpread1(2).SetText COL_FS_STEUERFREI, 1, "Steuer-Frei*"
  fpSpread1(2).SetText COL_FS_GESAMT, 1, "Gesamt"
  
  fpSpread1(2).SetText COL_FS_WRG, 2, frmParent.txt1(17) 'Währung 1
  
  fpSpread1(2).SetText COL_FS_WRG, 3, frmParent.txt1(18) 'Währung 2
    
  fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 4, "UID-Nr.: " & frmParent.txt1(14)
  fpSpread1(2).SetText COL_FS_UST, 4, "USt-Nr.: " & frmParent.txt1(13)
  fpSpread1(2).SetText COL_FS_STEUERFREI, 4, "(Laut Angaben des Beleg-Empfängers)"
  'fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 5, ZahlungsZiel(frmParent.txt1(16), 0, frmParent.txt1(17), frmParent.txt1(21), frmParent.txt1(22), frmParent.txt1(23))
  
  FussZahlungsZiel
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "FussFuellen")
        '***Ende
End Sub


Public Sub PostenSetUp()
        '***Beginn
On Error GoTo Fehler
        '***Ende
    Dim i As Integer
    Dim GPreisFormula As String
    'Lock the entire control
    fpSpread1(1).Row = -1
    fpSpread1(1).Col = -1
    fpSpread1(1).Lock = True
    fpSpread1(1).SetActionKey 0, False, False, 0 'Löschaktion der F2-Taste ausschalten
    
    'Zellen sofort in EditMode Versetzen. (Markieren der Zellen ist in dem Modus nicht möglich.)
    fpSpread1(1).EditModePermanent = True

    'fpSpread1(1).ShadowColor = vbWhite 'Headers-Farbe
    fpSpread1(1).TypeMaxEditLen = 60
    fpSpread1(1).EditEnterAction = EditEnterActionNext

    fpSpread1(1).MaxCols = 21
    'fpSpread1(1).MaxRows = 25
    'Change font size
    fpSpread1(1).Row = -1
    fpSpread1(1).Col = -1
    fpSpread1(1).FontSize = 9
    'Turn off grid lines
    fpSpread1(1).GridShowHoriz = False
    fpSpread1(1).GridShowVert = False
    'Allow cell contents to overflow into adjacent cells
    fpSpread1(1).AllowCellOverflow = True
    'Turn on edit mode
    'fpSpread1(1).EditModePermanent = True
    'Allow the tab key to operate
    fpSpread1(1).ProcessTab = True
    'Changes its size to the number of specified columns and rows
    'fpSpread1(1).AutoSize = True
    'Turn off headers
    'fpSpread1(1).ColHeadersShow = False
    fpSpread1(1).RowHeadersShow = False
    'Highlight entire cell contents when clicked in
    'fpSpread1(1).EditModeReplace = True
    'Set up col widths


    fpSpread1(1).ColWidth(COL_ZEILENART) = 400
    fpSpread1(1).ColWidth(COL_ARTSCHL) = 700

    fpSpread1(1).ColWidth(COL_ARTIKEL) = TEXT_BREITE
    fpSpread1(1).ColWidth(COL_MENGE) = 900
    fpSpread1(1).ColWidth(COL_EINHEIT) = 700
    fpSpread1(1).ColWidth(COL_EPREIS) = 900
    fpSpread1(1).ColWidth(COL_RABATT) = 700
    fpSpread1(1).ColWidth(COL_GPREIS) = 1000

    fpSpread1(1).ColWidth(COL_UST) = 400
    fpSpread1(1).ColWidth(COL_KOSTSCHL) = 700
    fpSpread1(1).ColWidth(COL_SACHSCHL) = 700
    fpSpread1(1).ColWidth(COL_KOSTKTO) = 700
    fpSpread1(1).ColWidth(COL_SACHKTO) = 700

    fpSpread1(1).ColWidth(COL_GPREISDUMMY) = 700
    fpSpread1(1).ColWidth(COL_GPREISDUMMY_O_UST) = 700
    fpSpread1(1).ColWidth(COL_GPREISDUMMY_M_UST) = 700
    fpSpread1(1).ColWidth(COL_SUMMEN) = 700


    fpSpread1(1).SetText COL_ZEILENART, 0, "Art"
    fpSpread1(1).SetText COL_ARTSCHL, 0, "Art." & vbCrLf & "-Schl."

    fpSpread1(1).SetText COL_ARTIKEL, 0, "Artikel"
    fpSpread1(1).SetText COL_MENGE, 0, "Menge"
    fpSpread1(1).SetText COL_EINHEIT, 0, "Einheit"
    fpSpread1(1).SetText COL_EPREIS, 0, "E-Preis"
    fpSpread1(1).SetText COL_RABATT, 0, "Rabatt"
    fpSpread1(1).SetText COL_GPREIS, 0, "G-Preis"

    fpSpread1(1).SetText COL_UST, 0, "USt."
    fpSpread1(1).SetText COL_KOSTSCHL, 0, "Kost." & vbCrLf & "-Schl."
    fpSpread1(1).SetText COL_SACHSCHL, 0, "Sach." & vbCrLf & "-Schl."
    fpSpread1(1).SetText COL_KOSTKTO, 0, "Kost." & vbCrLf & "-Kto."
    fpSpread1(1).SetText COL_SACHKTO, 0, "Sach." & vbCrLf & "-Kto."

    fpSpread1(1).SetText COL_GPREISDUMMY, 0, "G-Pr"
    fpSpread1(1).SetText COL_GPREISDUMMY_O_UST, 0, "G-Pr o. Ust"
    fpSpread1(1).SetText COL_GPREISDUMMY_M_UST, 0, "G-Pr m. Ust"
    fpSpread1(1).SetText COL_SUMMEN, 0, "Summ"
    fpSpread1(1).SetText COL_ERSTDAT, 0, "ErstDat"
    fpSpread1(1).SetText COL_ERSTVON, 0, "ErstVon"
    fpSpread1(1).SetText COL_AENDDAT, 0, "AendDat"
    fpSpread1(1).SetText COL_AENDVON, 0, "AendVon"
  

    fpSpread1(1).RowHeight(0) = 500
    fpSpread1(1).Col = 1
    fpSpread1(1).Row = 0
    fpSpread1(1).Col2 = fpSpread1(1).MaxCols
    fpSpread1(1).Row2 = 0
    fpSpread1(1).BlockMode = True
    fpSpread1(1).FontBold = True
    fpSpread1(1).BlockMode = False

    'Only show buttons when row is active
    fpSpread1(1).ButtonDrawMode = 1  ' Current Cell


    For i = 1 To fpSpread1(1).MaxCols
      Select Case i
      Case COL_ZEILENART
        fpSpread1(1).Col = i
        fpSpread1(1).Row = 1
        fpSpread1(1).Col2 = i
        fpSpread1(1).Row2 = fpSpread1(1).MaxRows
        fpSpread1(1).BlockMode = True
        fpSpread1(1).Lock = False
        'fpSpread1(1).CellType = CellTypeButton
        'fpSpread1(1).TypeButtonAlign = TypeButtonAlignRight
        fpSpread1(1).CellType = CellTypeComboBox
        'fpSpread1(1).TypeComboBoxList = "" & vbTab & "A" & vbTab & "T" & vbTab & "Z"
        fpSpread1(1).TypeComboBoxEditable = True
        fpSpread1(1).BlockMode = False
      Case COL_MENGE, COL_EPREIS, COL_RABATT, COL_GPREIS, COL_GPREISDUMMY, COL_GPREISDUMMY_O_UST, COL_GPREISDUMMY_M_UST, COL_SUMMEN 'Numerischen Felder
        fpSpread1(1).Col = i
        fpSpread1(1).Row = 0
        fpSpread1(1).Col2 = i
        fpSpread1(1).Row2 = fpSpread1(1).MaxRows
        fpSpread1(1).BlockMode = True
        fpSpread1(1).TypeHAlign = 1 'right
        fpSpread1(1).BlockMode = False

        fpSpread1(1).Col = i
        fpSpread1(1).Row = 1
        fpSpread1(1).Col2 = i
        fpSpread1(1).Row2 = fpSpread1(1).MaxRows
        fpSpread1(1).BlockMode = True
        fpSpread1(1).CellType = CellTypeNumber  'Integer
        fpSpread1(1).TypeNumberDecPlaces = 2
        Select Case i
        Case COL_GPREIS, COL_GPREISDUMMY  'Gesamtsumme, Dummy
          ''Ohne % in Spalte Einheit
          'fpSpread1(1).Formula = "(D#*F#)-((D#*F#)*(G#/100))"
          'fpSpread1(1).Formula = "(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)*(" & COL_RABATT_STR & "#/100))"
          
          'fpSpread1(1).Formula = "IF(E#=""%"",(D#*F#/100)-((D#*F#/100)*(G#/100)),(D#*F#)-((D#*F#)*(G#/100)))"
          GPreisFormula = "IF(" & COL_EINHEIT_STR & "#=""%"",(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#/100)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#/100)*(" & COL_RABATT_STR & "#/100)),(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)*(" & COL_RABATT_STR & "#/100)))"
          fpSpread1(1).Formula = GPreisFormula
        Case COL_GPREISDUMMY_O_UST
          'fpSpread1(1).Formula = "IF(I#=""0"",(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)*(" & COL_RABATT_STR & "#/100)),""0"")"
          'fpSpread1(1).Formula = "IF(" & COL_UST_STR & "#=""0"",(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)*(" & COL_RABATT_STR & "#/100)),""0"")"
          
          fpSpread1(1).Formula = "IF(" & COL_UST_STR & "#=""0""," & GPreisFormula & ",""0"")"
        Case COL_GPREISDUMMY_M_UST
          'fpSpread1(1).Formula = "IF(" & COL_UST_STR & "#=""1"",(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)*(" & COL_RABATT_STR & "#/100)),""0"")"
          
          fpSpread1(1).Formula = "IF(" & COL_UST_STR & "#=""1""," & GPreisFormula & ",""0"")"
          
        End Select
        fpSpread1(1).BlockMode = False
      Case COL_UST
        fpSpread1(1).Col = i
        fpSpread1(1).Row = 0
        fpSpread1(1).Col2 = i
        fpSpread1(1).Row2 = fpSpread1(1).MaxRows
        fpSpread1(1).BlockMode = True
        fpSpread1(1).TypeHAlign = TypeHAlignCenter
        fpSpread1(1).TypeCheckCenter = True
        fpSpread1(1).BlockMode = False
      End Select
    Next i

    'Rechnungssumme
    fpSpread1(1).Col = COL_SUMMEN
    fpSpread1(1).Row = 1 'Steuerpflichtige Summe
    fpSpread1(1).Formula = "SUM(" & COL_GPREISDUMMY_M_UST_STR & "1:" & COL_GPREISDUMMY_M_UST_STR & CStr(fpSpread1(1).MaxRows) & ")"
    
    fpSpread1(1).Row = 2 'Steuerfreie Summe
    fpSpread1(1).Formula = "SUM(" & COL_GPREISDUMMY_O_UST_STR & "1:" & COL_GPREISDUMMY_O_UST_STR & CStr(fpSpread1(1).MaxRows) & ")"

    'Die für Berechnungen benutzten Spalten unsichtbar machen
    fpSpread1(1).Col = COL_GPREISDUMMY
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = fpSpread1(1).MaxCols
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).ColHidden = True
    fpSpread1(1).BlockMode = False
         
    'Nicht druckbaren Spalten farblich kennzeichnen
    fpSpread1(1).Col = COL_ZEILENART
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_ARTSCHL
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).BackColor = vbButtonFace
    fpSpread1(1).BlockMode = False
         
    fpSpread1(1).Col = COL_KOSTSCHL
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_SACHKTO
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).BackColor = vbButtonFace
    fpSpread1(1).BlockMode = False
    
    'Focus auf die erste Zelle setzen
    fpSpread1(1).SetActiveCell 1, 1
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "PostenSetUp")
        '***Ende
End Sub

Public Sub ZeilenTyp(ByVal Row As Long)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim Knz As Variant
  Dim i As Long
  Dim SeitenUmbruch As String
  
  Schalter True
  
  fpSpread1(1).Row = Row
  
  If fpSpread1(1).GetText(COL_ZEILENART, Row, Knz) = False Then
    Knz = ""
  End If
      
  fpSpread1(1).ClearRange 2, Row, fpSpread1(1).MaxCols, Row, True
  fpSpread1(1).Col = COL_UST
  fpSpread1(1).CellType = CellTypeStaticText
  fpSpread1(1).Col = COL_ARTSCHL
  fpSpread1(1).CellType = CellTypeStaticText
  Select Case Knz
  Case ""
    For i = 2 To COL_LASTEDIT
      fpSpread1(1).Col = i
      fpSpread1(1).Lock = True
    Next i
  Case "A" 'Artikel
    For i = 2 To COL_LASTEDIT
      If i <> COL_GPREIS Then
        fpSpread1(1).Col = i
        fpSpread1(1).Lock = False
      End If
    Next i
    
    fpSpread1(1).Col = COL_ARTSCHL
    fpSpread1(1).CellType = CellTypeComboBox
    
    'fpSpread1(1).SetText COL_ARTIKEL, Row, "Artikel" & Row 'Test
    fpSpread1(1).Col = COL_ARTIKEL
    fpSpread1(1).TypeMaxEditLen = 60
    
    fpSpread1(1).Col = COL_UST
    fpSpread1(1).CellType = CellTypeCheckBox
  Case "T" 'Text
    For i = 2 To COL_LASTEDIT
      fpSpread1(1).Col = i
      If i < COL_MENGE Then
        fpSpread1(1).Lock = False
      Else
        fpSpread1(1).Lock = True
      End If
    Next i
    fpSpread1(1).Col = COL_ARTSCHL
    fpSpread1(1).CellType = CellTypeComboBox
    
    'fpSpread1(1).SetText COL_ARTIKEL, Row, "Freier Text" & Row 'Test
    fpSpread1(1).Col = COL_ARTIKEL
    fpSpread1(1).TypeMaxEditLen = 90
  Case "Z" 'Zwischensumme
    For i = 2 To COL_LASTEDIT
      fpSpread1(1).Col = i
      If i < COL_MENGE Then
        fpSpread1(1).Lock = False
      Else
        fpSpread1(1).Lock = True
      End If
    Next i
    fpSpread1(1).Col = COL_ARTSCHL
    fpSpread1(1).Lock = True
    fpSpread1(1).SetText COL_ARTIKEL, Row, "Zwischensumme"
    fpSpread1(1).SetText COL_GPREIS, Row, ZwischenSumme(Row)
  Case "S" 'Seitenumbruch
    For i = 2 To COL_LASTEDIT
      fpSpread1(1).Col = i
      fpSpread1(1).Lock = True
    Next i
    'fpSpread1(1).Col = COL_ARTSCHL
    'fpSpread1(1).CellType = CellTypeComboBox
    
    fpSpread1(1).Col = COL_ARTIKEL
    fpSpread1(1).TypeMaxEditLen = 200
    SeitenUmbruch = "- - - - - - - - - - - - - - - - - - - - " '40 Zeichen
    SeitenUmbruch = SeitenUmbruch + SeitenUmbruch + SeitenUmbruch + SeitenUmbruch + SeitenUmbruch
    fpSpread1(1).SetText COL_ARTIKEL, Row, SeitenUmbruch
  End Select
  
  'Erstellung protokollieren
  fpSpread1(1).SetText COL_ERSTDAT, Row, CStr(Now)
  fpSpread1(1).SetText COL_ERSTVON, Row, GsUser
  fpSpread1(1).SetText COL_AENDDAT, Row, CStr(Now)
  fpSpread1(1).SetText COL_AENDVON, Row, GsUser
  
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "ZeilenTyp")
        '***Ende
End Sub
Public Function IstZeilenTyp(ByVal Row As Long, ZeilenTyp As String) As Boolean
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim Knz As Variant
  Dim AltCol As Long
  
  AltCol = fpSpread1(1).Col
  
  fpSpread1(1).Row = Row
  fpSpread1(1).Col = COL_UST
  
  Select Case ZeilenTyp
  Case "A" 'Artikel
    If fpSpread1(1).CellType = CellTypeCheckBox Then
      IstZeilenTyp = True
    End If
  Case "T" 'Text
    If fpSpread1(1).CellType <> CellTypeCheckBox Then
      fpSpread1(1).Col = COL_ARTSCHL
      If fpSpread1(1).CellType = CellTypeComboBox Then
        IstZeilenTyp = True
      End If
    End If
  Case "Z" 'Zwischensumme
    If fpSpread1(1).CellType <> CellTypeCheckBox Then
      fpSpread1(1).Col = COL_ARTSCHL
      If fpSpread1(1).CellType <> CellTypeComboBox Then
        IstZeilenTyp = True
      End If
    End If
  Case "S" 'Seitenumbruch
    If fpSpread1(1).CellType <> CellTypeCheckBox Then
      fpSpread1(1).Col = COL_ARTSCHL
      If fpSpread1(1).CellType <> CellTypeComboBox Then
        IstZeilenTyp = True
      End If
    End If
  End Select
  
  fpSpread1(1).Col = AltCol
  
        '***Beginn
        Exit Function
Fehler:
        Call FehlerErklärung("frmSP52831", "IstZeilenTyp")
        '***Ende
End Function

Public Function ZwischenSumme(ByVal Row As Long) As Double
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim i As Integer
  Dim Knz As Variant
  
  fpSpread1(1).Row = Row
  
  For i = fpSpread1(1).Row - 1 To 1 Step -1
    If fpSpread1(1).GetText(COL_ZEILENART, i, Knz) Then
      If Knz = "A" Then
        If fpSpread1(1).GetText(COL_GPREIS, i, Knz) Then
          If IsNumeric(Knz) Then
            ZwischenSumme = ZwischenSumme + Knz
          End If
        End If
      End If
    End If
    If fpSpread1(1).GetText(COL_ZEILENART, i, Knz) Then
      If Knz = "Z" Then Exit For
    End If
  Next i
    
        '***Beginn
        Exit Function
Fehler:
        Call FehlerErklärung("frmSP52831", "ZwischenSumme")
        '***Ende
End Function

Public Sub ZwischenSummeRefresh(ByVal Row As Long, Optional Alle As Boolean)
  'Alle=True -> Es werden alle Zwischensummen aktuallisiert.
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim i As Integer
  Dim Knz As Variant
  
  For i = Row To LastRow
    If fpSpread1(1).GetText(COL_ZEILENART, i, Knz) Then
      If Knz = "Z" Then
        fpSpread1(1).SetText COL_GPREIS, i, ZwischenSumme(i)
        If Alle = False Then Exit For
      End If
    End If
  Next i
    
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "ZwischenSummeRefresh")
        '***Ende
End Sub

Private Sub fpSpread1_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim Knz As Variant
  Dim Ret As Double
  Static i As Integer
 
  If Row = fpSpread1(1).ActiveRow Then
    'i = i + 1
    'Debug.Print i & " col: " & Col & " row: " & Row
    Select Case Index
    Case 1 'Posten
      Select Case Col
      Case COL_EPREIS
        fpSpread1(1).GetFloat COL_EPREIS, Row, Ret
        fpSpread1(1).Row = Row
        fpSpread1(1).Col = COL_EPREIS
        If Ret = 0 Then
            fpSpread1(1).ForeColor = vbRed
        Else
            fpSpread1(1).ForeColor = vbBlack
        End If
      Case COL_GPREIS
        ZwischenSummeRefresh Row
      Case COL_ZEILENART
        ZeilenTyp Row
      Case COL_SUMMEN
        'Rechnungsfuß aktualisieren
        Select Case Row
        Case 1 'Steuerpflichtige Summe
          If fpSpread1(1).GetText(COL_SUMMEN, Row, Knz) Then
            fpSpread1(2).SetText 2, 2, Knz
          Else
            fpSpread1(2).SetText 2, 2, "0"
          End If
        Case 2 'Steuerfreie Summe
          If fpSpread1(1).GetText(COL_SUMMEN, Row, Knz) Then
            fpSpread1(2).SetText 4, 2, Knz
          Else
            fpSpread1(2).SetText 4, 2, "0"
          End If
        End Select
      Case COL_EINHEIT
        If fpSpread1(1).GetText(COL_EINHEIT, Row, Knz) Then
          If Trim(Knz) = "%" Then
            'Prozentualer Auf-, Abschlag
            fpSpread1(1).SetText COL_EPREIS, Row, LetzterBetrag(Row)
          End If
        End If
  
      Case Else
        If gbAfterLoad Then
          fpSpread1(1).SetText COL_AENDDAT, Row, CStr(Now)
          fpSpread1(1).SetText COL_AENDVON, Row, GsUser
        End If
      End Select
    Case 2 'Fuß
      Select Case Col
      Case COL_FS_GESAMT
        Select Case Row
        Case 2 'Rechnungssmme
          FussZahlungsZiel
        End Select
      End Select
      
    End Select
  End If

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "fpSpread1_Change")
        '***Ende
End Sub

Private Sub fpSpread1_ComboDropDown(Index As Integer, ByVal Col As Long, ByVal Row As Long)
        '***Beginn
'On Error GoTo Fehler
        '***Ende
  
  Auswahl Col, Row

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "fpSpread1_ComboDropDown")
        '***Ende
End Sub

Private Sub fpSpread1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        '***Beginn
'On Error GoTo Fehler
        '***Ende
  Dim NeuRow As Long
  Dim Knz As Variant
  
  If Index = 1 Then
    Select Case KeyCode
    Case vbKeyF2
      Select Case fpSpread1(1).ActiveCol
      Case COL_ZEILENART, COL_ARTSCHL
        Auswahl fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow
      End Select
    Case vbKeyReturn
      If Shift = 1 Then
        'Focus auf die erste Zelle setzen
        If fpSpread1(1).ActiveRow = fpSpread1(1).MaxRows Then
          NeuRow = fpSpread1(1).ActiveRow
        Else
          NeuRow = fpSpread1(1).ActiveRow + 1
        End If
        fpSpread1(1).SetActiveCell COL_ZEILENART, NeuRow
      Else
        Select Case fpSpread1(1).ActiveCol
        Case COL_ARTSCHL
          If fpSpread1(1).GetText(COL_ARTSCHL, fpSpread1(1).ActiveRow, Knz) Then
            If Trim(Knz) <> "" Then
              objDAusw.GetIfOnesHit = True
              Auswahl fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow
              objDAusw.GetIfOnesHit = False
            End If
          End If
        End Select
      End If
    End Select
    
  End If
  
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "fpSpread1_KeyDown")
        '***Ende
End Sub

Public Sub Auswahl(ByVal Col As Long, ByVal Row As Long)
  
  Dim i As Integer
  Dim ColLeft As Integer
  Dim RowBottom As Integer
  Dim Merk As String
  Dim SQL As String
  Dim rc As rect
  Dim Knz As Variant
  Dim Knz1 As Variant
  Dim TextRows As Long
  Dim Ret As Double
  Dim LR As Long
  
 '***Beginn
'On Error GoTo Fehler
 '***Ende
  
  Call GetWindowRect(fpSpread1(1).hwnd, rc)
  ColLeft = rc.left * Screen.TwipsPerPixelX - 70
  RowBottom = rc.top * Screen.TwipsPerPixelY - 340
  
  For i = 0 To Col - 1
    ColLeft = ColLeft + fpSpread1(1).ColWidth(i) + 15 '15 Trenlinienbreite.
  Next i
  RowBottom = RowBottom + fpSpread1(1).RowHeight(0) + (Row * (fpSpread1(1).RowHeight(Row) + 15)) '15 Trenlinienhöhe.
  
  objDAusw.BorderStyle = 4
  If fpSpread1(1).GetText(Col, 0, Knz) Then objDAusw.Caption = StrClean(Knz, 10, 13)
  objDAusw.top = RowBottom
  objDAusw.left = ColLeft
  objDAusw.ColParameter 0, ColWidth, fpSpread1(1).ColWidth(Col)
  
  Select Case Col
  Case COL_ZEILENART
    objDAuswDef.FilterBar = True
    'Datensatz positionieren.
    If fpSpread1(1).GetText(Col, Row, Knz) And Trim(Knz) <> "" Then
      objDAuswDef.FindFirst ("Knz like '" & Knz & "%'")
      Merk = Knz
    End If
    objDAuswDef.RSOpen "SELECT Knz, KnzBez1 FROM [Auswahl] WHERE TabName = '2800_Folge' AND FeldName = 'SatzTyp' ORDER BY Knz"
    If objDAuswDef.Abbruch = False Then
      fpSpread1(1).SetText COL_ZEILENART, Row, objDAuswDef.FieldText(0)
      If Merk <> objDAuswDef.FieldText(0) Then
        ZeilenTyp Row
        fpSpread1(1).SetActiveCell Col + 1, Row
      Else
        If IstZeilenTyp(Row, Merk) = False Then
          'Der ursprunglich erfasste Zeilentyp wurde noch nicht bestätigt.
          'Der Zeilentyp muss jetzt gesetzt werden.
          ZeilenTyp Row
          fpSpread1(1).SetActiveCell Col + 1, Row
        End If
      End If
    Else
      If IstZeilenTyp(Row, Merk) = False Then
        'Der ursprunglich erfasste Zeilentyp wurde noch nicht bestätigt.
        'Der Zeilentyp muss jetzt gesetzt werden.
        ZeilenTyp Row
      End If
      fpSpread1(1).SetActiveCell Col, Row
    End If
  Case COL_ARTSCHL
    objDAusw.FilterBar = True
    objDAusw.MaxWidth = 12000
    If fpSpread1(1).GetText(COL_ZEILENART, Row, Knz) And Trim(Knz) <> "" Then
      Select Case UCase(Trim(Knz))
      Case "A"
        'SQL = "SELECT Schl, Bez, Menge, Einheit, Preis, Rabatt, Steuer, KostSchl, FiBuSchl, KostKonto, FibuKonto, TextSchl FROM [2800_Artikel] ORDER BY Schl"
        SQL = "SELECT Schl, Bez, Menge, Einheit, Preis, Wrg, Rabatt, Steuer, KostSchl, FiBuSchl, KostKonto, FibuKonto, TextSchl FROM [2800_Artikel] ORDER BY Schl"
        objDAusw.ColParameter 2, ColNumberFormat, "#0.00"
        objDAusw.ColParameter 3, ColWidth, 600
        objDAusw.ColParameter 4, ColNumberFormat, "#0.00"
        objDAusw.ColParameter 5, ColWidth, 600
        objDAusw.ColParameter 6, ColNumberFormat, "#0.00"
        
      
        objDAusw.ColParameter 7, ColVisible, 0
        objDAusw.ColParameter 8, ColVisible, 0
        objDAusw.ColParameter 9, ColVisible, 0
        objDAusw.ColParameter 10, ColVisible, 0
        objDAusw.ColParameter 11, ColVisible, 0
        objDAusw.ColParameter 12, ColVisible, 0
      Case "T"
        SQL = "SELECT Schl, Bez, Inhalt FROM [2800_Texte] ORDER BY Schl"
      End Select
    End If
    
    'Datensatz positionieren.
    If fpSpread1(1).GetText(Col, Row, Knz1) And Trim(Knz1) <> "" Then
      objDAusw.FindFirst ("Schl like '" & Knz1 & "%'")
    End If
    
    objDAusw.RSOpen SQL
    If objDAusw.Abbruch = False Then
      If gbZeileInCopy = False Then
        LR = LastRow - Row
        If LR > 0 Then ZeileBearbeiten 2, Row + 1, LR
      End If
      
      fpSpread1(1).SetText COL_ARTSCHL, Row, objDAusw.FieldText(0)
      Select Case UCase(Trim(Knz))
      Case "A"
        fpSpread1(1).SetText COL_ARTIKEL, Row, objDAusw.FieldText(1)
        fpSpread1(1).SetText COL_MENGE, Row, objDAusw.FieldText(2)
        fpSpread1(1).SetText COL_EINHEIT, Row, objDAusw.FieldText(3)
        If Trim(objDAusw.FieldText(3)) = "%" Then
          fpSpread1(1).SetText COL_EPREIS, Row, LetzterBetrag(Row)
        Else
          If Trim(UCase(objDAusw.FieldText(5))) = Trim(UCase(frmParent.txt1(17))) Then
            'Rechnungswährung und Artikelstammwährung sind gleich. Der Betrag wird übernommen.
            fpSpread1(1).SetText COL_EPREIS, Row, objDAusw.FieldText(4)
          Else
            fpSpread1(1).SetText COL_EPREIS, Row, 0
          End If
        End If
        'Farbe der Zelle abhängig vom Betrag setzen.
        fpSpread1(1).GetFloat COL_EPREIS, Row, Ret
        fpSpread1(1).Row = Row
        fpSpread1(1).Col = COL_EPREIS
        If Ret = 0 Then
            fpSpread1(1).ForeColor = vbRed
        Else
            fpSpread1(1).ForeColor = vbBlack
        End If
        
        fpSpread1(1).SetText COL_RABATT, Row, objDAusw.FieldText(6)
        If frmParent.Combo1(0).ListIndex = 0 Or frmParent.Combo1(0).ListIndex = 2 Then
          'Steuerfreier Kunde
          fpSpread1(1).SetText COL_UST, Row, "0"
        Else
          fpSpread1(1).SetText COL_UST, Row, objDAusw.FieldText(7)
        End If
        fpSpread1(1).SetText COL_KOSTSCHL, Row, objDAusw.FieldText(8)
        fpSpread1(1).SetText COL_SACHSCHL, Row, objDAusw.FieldText(9)
        fpSpread1(1).SetText COL_KOSTKTO, Row, objDAusw.FieldText(10)
        fpSpread1(1).SetText COL_SACHKTO, Row, objDAusw.FieldText(11)
        If Trim(objDAusw.FieldText(12)) <> "" Then
          'Verweis auf ein Text in 2800_Texte
          fpSpread1(1).SetText COL_ZEILENART, Row + 1, "T"
          ZeilenTyp Row + 1
          fpSpread1(1).SetText COL_ARTSCHL, Row + 1, objDAusw.FieldText(12)
          objDAusw.GetIfOnesHit = True
          Auswahl COL_ARTSCHL, Row + 1
          objDAusw.GetIfOnesHit = False
        End If
        fpSpread1(1).SetActiveCell Col + 1, Row
      Case "T"
        
        txt1(0) = objDAusw.FieldText(2) 'Dummytextfeld
        Merk = GetRealText(txt1(0), TextRows)
        
        For i = Row + 1 To Row + TextRows - 1
          fpSpread1(1).SetText COL_ZEILENART, i, "T"
          ZeilenTyp i
        Next i
        fpSpread1(1).Row = Row
        fpSpread1(1).Clip = Merk
        fpSpread1(1).SetActiveCell COL_ZEILENART, i
        
      End Select
      
      If LR > 0 Then
        ZeileBearbeiten 3, LastRow + 1, 0
        'Da die Tabelle neuaufgebaut ist müssen alle ZwischenSummen aktualisiert werden.
        ZwischenSummeRefresh 1, True
      End If
    
    Else
      fpSpread1(1).SetActiveCell Col, Row
    End If
  End Select

       
 '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "Auswahl")
 '***Ende

End Sub
Public Sub ZeileEifuegen(Row As Long, Anzahl As Long)
  'Row - Position, an welcher die Zeile(n) eingefügt werden soll(en).
  'Anzahl - Anzahl der Zeilen, die eingefügt werden sollen.
  Dim array1 As Long
  Dim array2 As Long
  Dim i As Integer
  Dim LR As Integer
  
  LR = LastRow
  If LR > Row Then
    'Größe des zu verschiebenden Bereichs ermitteln.
    array1 = LR - Row
    array2 = fpSpread1(1).MaxCols
    
    ReDim fparray(array1, array2) As Variant
    ReDim fparrayArt(array1, 1) As Variant

    'Speichern des zu verschiebenden Bereichs (Get data: ColLeft, RowTop)
    fpSpread1(1).GetArray 1, Row, fparray
    fpSpread1(1).GetArray 1, Row, fparrayArt
    
    'Löschen des in fparray gespeicherten Bereichs.
    fpSpread1(1).ClearRange 1, Row, fpSpread1(1).MaxCols, LR, True
   
    'Erste Grid-Spalte füllen.
    fpSpread1(1).SetArray 1, Row + Anzahl, fparrayArt
    
    LR = LastRow
    For i = Row To LR
      ZeilenTyp i
    Next i
    
    'Grid füllen
    fpSpread1(1).SetArray 1, Row + Anzahl, fparray
    
    'Eingefügte Zeile aktivieren
    fpSpread1(1).SetActiveCell 1, Row
  End If
    
End Sub

Public Sub ZeileBearbeiten(Aktion As Integer, Row As Long, Anzahl As Long)
  'Aktion - 1=Kopieren
  '         2=Ausschneiden
  '         3=Einfügen
  'Row - Position, ab welcher die Zeile(n) bearbeitet werden soll(en).
  'Anzahl - Anzahl der Zeilen, die bearbeitet werden sollen. (Für 3=Einfügen ist der Parameter =0)
  Dim array1 As Long
  Dim array2 As Long
  Dim i As Integer
  Dim LR As Integer
  
  If Aktion < 3 Then
    If Anzahl > 0 Then
      'Größe des zu kopierenden Bereichs ermitteln.
      array1 = Anzahl - 1
      array2 = fpSpread1(1).MaxCols
      
      ReDim gFPArray(array1, array2) As Variant
      ReDim gFPArray1(array1, 1) As Variant
  
      'Speichern des zu kopierenden Bereichs (Get data: ColLeft, RowTop)
      fpSpread1(1).GetArray 1, Row, gFPArray
      fpSpread1(1).GetArray 1, Row, gFPArray1
      
      If Aktion = 2 Then
        'Löschen des in fparray gespeicherten Bereichs.
        fpSpread1(1).ClearRange 1, Row, fpSpread1(1).MaxCols, Row + Anzahl - 1, True
        For i = Row To Row + Anzahl - 1
          ZeilenTyp i
        Next i
      End If
      gbZeileInCopy = True
    End If
  Else
    'If UBound(gFPArray, 1) > -1 Then
    If gbZeileInCopy Then
      'Erste Spalte des Grids füllen.
      fpSpread1(1).SetArray 1, Row, gFPArray1
      
      For i = Row To Row + UBound(gFPArray, 1)
        ZeilenTyp i
      Next i
      
      'Grid füllen
      fpSpread1(1).SetArray 1, Row, gFPArray
      
      'Eingefügte Zeile aktivieren
      fpSpread1(1).SetActiveCell 1, Row
      gbZeileInCopy = False
    End If
  End If
  'Err: 9 Index außer gültigen Bereichs (Wenn gFPArray nicht gefüllt ist.)
End Sub

Public Sub FolgeZeigen(BelegID As Long)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim rs As DAO.Recordset
  Dim SQL As String
  Dim i As Integer
  
  fpSpread1(1).ClearRange 1, 1, fpSpread1(1).MaxCols, fpSpread1(1).MaxRows, True
  fpSpread1(1).Col = 2
  fpSpread1(1).Row = 1
  fpSpread1(1).Col2 = fpSpread1(1).MaxCols
  fpSpread1(1).Row2 = fpSpread1(1).MaxRows
  fpSpread1(1).BlockMode = True
  fpSpread1(1).Lock = True
  fpSpread1(1).Col = COL_UST
  fpSpread1(1).Col2 = COL_UST
  fpSpread1(1).CellType = CellTypeStaticText
  fpSpread1(1).Col = COL_ARTSCHL
  fpSpread1(1).Col2 = COL_ARTSCHL
  fpSpread1(1).CellType = CellTypeStaticText
  fpSpread1(1).BlockMode = False
  
  fpSpread1(1).Col = 1
  fpSpread1(1).Row = 1
  
  
  SQL = "SELECT * FROM [2800_Folge] WHERE BelegID = " & BelegID & " ORDER BY Nr"
  Set rs = GDB.OpenRecordset(SQL, dbOpenDynaset)
  If rs.RecordCount > 0 Then
    Do Until rs.EOF
      i = i + 1
      fpSpread1(1).SetText COL_ZEILENART, i, rs!SatzTyp
      ZeilenTyp i
      Select Case UCase(rs!SatzTyp)
      Case "A"
        fpSpread1(1).SetText COL_ARTSCHL, i, rs!Schl
        fpSpread1(1).SetText COL_ARTIKEL, i, rs!Bez
        fpSpread1(1).SetText COL_MENGE, i, rs!Menge
        fpSpread1(1).SetText COL_EINHEIT, i, rs!Einheit
        fpSpread1(1).SetText COL_EPREIS, i, rs!Epreis
        fpSpread1(1).SetText COL_RABATT, i, rs!Rabatt
        fpSpread1(1).SetText COL_UST, i, rs!Steuer
        fpSpread1(1).SetText COL_KOSTSCHL, i, rs!KostSchl
        fpSpread1(1).SetText COL_SACHSCHL, i, rs!FiBuSchl
        fpSpread1(1).SetText COL_KOSTKTO, i, rs!KostKonto
        fpSpread1(1).SetText COL_SACHKTO, i, rs!FiBuKonto
      Case "S"
      Case "T"
        fpSpread1(1).SetText COL_ARTSCHL, i, rs!Schl
        fpSpread1(1).SetText COL_ARTIKEL, i, rs!Bez
      Case "Z"
        fpSpread1(1).SetText COL_ARTIKEL, i, rs!Bez
      End Select
      
      fpSpread1(1).SetText COL_ERSTDAT, i, rs!ErstDat
      fpSpread1(1).SetText COL_ERSTVON, i, rs!ErstVon
      fpSpread1(1).SetText COL_AENDDAT, i, rs!AendDat
      fpSpread1(1).SetText COL_AENDVON, i, rs!AendVon
      
      rs.MoveNext
    Loop
  End If
  rs.Close
  Set rs = Nothing
  cmd1(2).Enabled = False
  
  fpSpread1(1).SetActiveCell 1, i + 1
    
  If GintDruck = 1 Then
    'Gedruckter Beleg darf nicht verändert werden.
    fpSpread1(1).Col = 1
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = fpSpread1(1).MaxCols
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).Lock = True
    fpSpread1(1).BlockMode = False
  End If
    
    
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "FolgeZeigen")
        '***Ende
End Sub



Private Sub fpSpread1_KeyPress(Index As Integer, KeyAscii As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim Knz As Variant
  
    If Index = 1 Then
      If fpSpread1(Index).ActiveCol = COL_ZEILENART Then
        
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
        Select Case KeyAscii
        Case 65, 83, 84, 90 'Zugelassen sind A, S, T, Z
        Case 8 'Backspace für löschen
          'If fpSpread1(Index).GetText(COL_ZEILENART, fpSpread1(Index).ActiveRow, Knz) Then
          '  If Len(Knz) >= 1 Then KeyAscii = 0
          'End If
        Case Else
          KeyAscii = 0
        End Select
        
'        If KeyAscii <> 65 And KeyAscii <> 83 And KeyAscii <> 84 And KeyAscii <> 90 And KeyAscii <> 8 Then
'          'Zugelassen sind A, S, T, Z, (Back für löschen)
'          KeyAscii = 0
'        End If
'
'        If KeyAscii <> 8 Then
'          If fpSpread1(Index).GetText(COL_ZEILENART, fpSpread1(Index).ActiveRow, Knz) Then
'            If Len(Knz) >= 1 Then KeyAscii = 0
'          End If
'        End If
        
      End If
    End If
    

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "fpSpread1_KeyPress")
        '***Ende
End Sub


Public Sub FussZahlungsZiel()
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim Knz As Variant
  
  fpSpread1(2).Col = COL_FS_STEUERPFLICHT
  fpSpread1(2).Row = 5
  fpSpread1(2).Col2 = COL_FS_GESAMT
  fpSpread1(2).Row2 = 6
  
  If fpSpread1(2).GetText(COL_FS_GESAMT, 2, Knz) = False Then
    Knz = 0
  End If
  
  fpSpread1(2).Clip = ZahlungsZiel(frmParent.txt1(16), Knz, frmParent.txt1(17), frmParent.txt1(21), frmParent.txt1(22), frmParent.txt1(23))

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "FussZahlungsZiel")
        '***Ende
End Sub

Private Sub imgSplitter_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  With imgSplitter(Index)
    picSplitter(Index).Move .left, .top \ 2, .Width, .Height '- 20
  End With
  picSplitter(Index).Visible = True

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "imgSplitter_MouseDown")
        '***Ende
End Sub


Private Sub imgSplitter_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        '***Beginn
On Error GoTo Fehler
        '***Ende

  Dim Pos As Single
  Dim SplitMin As Single
  Dim SplitMax As Single
  
  If Index = 0 Then
    SplitMin = 1200
    SplitMax = gsngKopfHeight
  Else
    SplitMin = gsngFussTop - imgSplitter(Index).Height
    SplitMax = Me.Height - 1500
  End If
  
  Pos = Y + imgSplitter(Index).top
  
  If Pos < SplitMin Then
    picSplitter(Index).top = SplitMin
  ElseIf Pos > SplitMax Then
    picSplitter(Index).top = SplitMax
  Else
    picSplitter(Index).top = Pos
  End If


        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "imgSplitter_MouseMove")
        '***Ende
End Sub


Private Sub imgSplitter_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  SizeControls picSplitter(Index).top, Index
  picSplitter(Index).Visible = False

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "imgSplitter_MouseUp")
        '***Ende
End Sub



Public Sub FolgeSpeichern(BelegID As Long, Optional tmp As Boolean)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim i As Integer
  Dim Row As Integer
  Dim Knz As Variant
  Dim rsFolge As DAO.Recordset
  Dim TmpZusatz As String
  
  If tmp Then
    TmpZusatz = "Tmp"
  Else
    cmd1(2).Enabled = False
  End If
  'Erstmal alle Sätze löschen.
  GDB.Execute "DELETE FROM [2800_Folge" & TmpZusatz & "] WHERE [BelegID] = " & BelegID, dbFailOnError
  Set rsFolge = GDB.OpenRecordset("SELECT * FROM [2800_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID & " ORDER BY Nr", dbOpenDynaset)
  

  For Row = 1 To LastRow
    rsFolge.AddNew
    
    rsFolge!BelegID = BelegID
    rsFolge!nr = Row
    If fpSpread1(1).GetText(COL_ZEILENART, Row, Knz) Then rsFolge!SatzTyp = Knz
    If Knz = "S" Then
      If fpSpread1(1).GetText(COL_ARTIKEL, Row, Knz) Then rsFolge!Bez = left(Knz, 1)
    Else
      If fpSpread1(1).GetText(COL_ARTIKEL, Row, Knz) Then rsFolge!Bez = Knz
    End If
    If fpSpread1(1).GetText(COL_ARTSCHL, Row, Knz) Then rsFolge!Schl = Knz
    If fpSpread1(1).GetText(COL_MENGE, Row, Knz) Then rsFolge!Menge = Knz
    If fpSpread1(1).GetText(COL_EINHEIT, Row, Knz) Then rsFolge!Einheit = Knz
    If fpSpread1(1).GetText(COL_EPREIS, Row, Knz) Then rsFolge!Epreis = Knz
    If fpSpread1(1).GetText(COL_RABATT, Row, Knz) Then rsFolge!Rabatt = Knz
    If fpSpread1(1).GetText(COL_UST, Row, Knz) Then rsFolge!Steuer = Knz
    If fpSpread1(1).GetText(COL_KOSTSCHL, Row, Knz) Then rsFolge!KostSchl = Knz
    If fpSpread1(1).GetText(COL_SACHSCHL, Row, Knz) Then rsFolge!FiBuSchl = Knz
    If fpSpread1(1).GetText(COL_KOSTKTO, Row, Knz) Then rsFolge!KostKonto = Knz
    If fpSpread1(1).GetText(COL_SACHKTO, Row, Knz) Then rsFolge!FiBuKonto = Knz
    If fpSpread1(1).GetText(COL_ERSTDAT, Row, Knz) Then rsFolge!ErstDat = Knz
    If fpSpread1(1).GetText(COL_ERSTVON, Row, Knz) Then rsFolge!ErstVon = Knz
    If fpSpread1(1).GetText(COL_AENDDAT, Row, Knz) Then rsFolge!AendDat = Knz
    If fpSpread1(1).GetText(COL_AENDVON, Row, Knz) Then rsFolge!AendVon = Knz
    
    rsFolge.Update
  Next Row

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "FolgeSpeichern")
        '***Ende
End Sub


Public Function LetzterBetrag(ByVal Row As Long) As Double
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Dim i As Integer
  Dim Knz As Variant
  
  fpSpread1(1).Row = Row
  
  For i = fpSpread1(1).Row - 1 To 1 Step -1
    If fpSpread1(1).GetText(COL_GPREIS, i, Knz) Then
      If IsNumeric(Knz) Then
        LetzterBetrag = Knz
        Exit For
      End If
    End If
  Next i

        '***Beginn
        Exit Function
Fehler:
        Call FehlerErklärung("frmSP52831", "LetzterBetrag")
        '***Ende
End Function

Public Function LastRow() As Long
        '***Beginn
On Error GoTo Fehler
        '***Ende
  'Anzahl der gefüllten Zeilen ermitteln.
  Dim Row As Integer
  Dim Knz As Variant
  
  For Row = 1 To fpSpread1(1).MaxRows
    If fpSpread1(1).GetText(COL_ZEILENART, Row, Knz) Then
      If Trim(Knz) <> "" Then LastRow = Row
    End If
  Next Row

        '***Beginn
        Exit Function
Fehler:
        Call FehlerErklärung("frmSP52831", "LastRow")
        '***Ende
End Function

Public Sub Schalter(Enabled As Boolean)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  cmd1(2).Enabled = Enabled
  cmd1(3).Enabled = Enabled
  cmd1(4).Enabled = Enabled
  cmd1(5).Enabled = Enabled

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52831", "Schalter")
        '***Ende
End Sub

