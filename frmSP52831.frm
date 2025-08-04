VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{2213E283-16BC-101D-AFD4-040224009C1D}#29.0#0"; "CMLL29O.OCX"
Begin VB.Form frmSP52831 
   Caption         =   "Form1"
   ClientHeight    =   10365
   ClientLeft      =   5175
   ClientTop       =   2850
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSP52831.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin ListLabel.ListLabel LL1 
      Left            =   11040
      Top             =   1440
      _Version        =   65537
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   64
      Language        =   -1
      DialogMode      =   14
      DialogFrame     =   0
      Dialog3DText    =   0
      DialogButtons   =   0
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
      Dummy8          =   -1  'True
      ShowPredefVars  =   -1  'True
      UseHostprinter  =   0   'False
      EMFResolution   =   0
      AddVarsToFields =   0   'False
      ConvertCRLF     =   -1  'True
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
      MaxRTFVersion   =   1025
      DelayTableHeader=   0   'False
      OfnDialogExplorer=   -1  'True
      CreateInfo      =   -1  'True
      XlatVarNames    =   -1  'True
      PhantomSpaceRepresentationCode=   8203
      LockNextCharRepresentationCode=   8288
      ExprSepRepresentationCode=   164
      TextQuoteRepresentationCode=   1
      InterCharSpacing=   0   'False
      IncludeFontDescent=   -1  'True
      Dummy6          =   -1  'True
      UseChartFields  =   0   'False
      Dummy7          =   -1  'True
      ProjectPassword =   ""
      LicensingInfo   =   ""
      IncrementalPreview=   -1  'True
      Printerless     =   0   'False
   End
   Begin VB.TextBox TextDummy 
      Enabled         =   0   'False
      Height          =   285
      Left            =   17040
      TabIndex        =   15
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl 
      Left            =   10560
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Timer Timer1 
      Left            =   11010
      Top             =   810
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   1
      Left            =   9930
      ScaleHeight     =   60
      ScaleWidth      =   4800
      TabIndex        =   13
      Top             =   9090
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Index           =   0
      Left            =   9900
      ScaleHeight     =   60
      ScaleWidth      =   4800
      TabIndex        =   12
      Top             =   510
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   9990
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmSP52831.frx":0442
      Top             =   90
      Visible         =   0   'False
      Width           =   915
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   4785
      Index           =   0
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   9795
      _Version        =   524288
      _ExtentX        =   17277
      _ExtentY        =   8440
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ArrowsExitEditMode=   -1  'True
      ColHeaderDisplay=   1
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      SpreadDesigner  =   "frmSP52831.frx":0448
      UnitType        =   2
      ScrollBarTrack  =   3
      CellNoteIndicator=   1
      AppearanceStyle =   0
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2595
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   9795
      _Version        =   524288
      _ExtentX        =   17277
      _ExtentY        =   4577
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColHeaderDisplay=   1
      EditEnterAction =   5
      EditModePermanent=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   -2147483633
      GridColor       =   -2147483633
      NoBorder        =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBarMaxAlign=   0   'False
      ScrollBarShowMax=   0   'False
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmSP52831.frx":08E9
      UnitType        =   2
      UserResize      =   2
      ScrollBarTrack  =   3
      AppearanceStyle =   0
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2595
      Index           =   2
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7380
      Width           =   9795
      _Version        =   524288
      _ExtentX        =   17277
      _ExtentY        =   4577
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      ColHeaderDisplay=   1
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
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
      SpreadDesigner  =   "frmSP52831.frx":0DC9
      UnitType        =   2
      ScrollBarTrack  =   3
      AppearanceStyle =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Unten ausrichten
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   10005
      Width           =   15150
      _Version        =   65536
      _ExtentX        =   26723
      _ExtentY        =   635
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd1 
         Caption         =   "&Pos.Übern."
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   11955
         TabIndex        =   18
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Leeren"
         Height          =   300
         Index           =   10
         Left            =   5640
         TabIndex        =   17
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Appearance      =   0  '2D
         Caption         =   "?"
         CausesValidation=   0   'False
         Height          =   300
         HelpContextID   =   401
         Index           =   9
         Left            =   60
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   30
         Width           =   255
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Zeile &entfernen[F4]"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   2
         Left            =   2130
         TabIndex        =   6
         Top             =   30
         Width           =   1500
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Zeile einfügen[F3]"
         Height          =   300
         Index           =   3
         Left            =   585
         TabIndex        =   5
         Top             =   30
         Width           =   1450
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&torno"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4200
         TabIndex        =   7
         Top             =   30
         Width           =   1380
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&peichern"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   10110
         TabIndex        =   1
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Designer"
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   17700
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "< &Zuruck"
         Height          =   300
         Index           =   5
         Left            =   7485
         TabIndex        =   3
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Ausgabe"
         Height          =   300
         Index           =   4
         Left            =   8805
         TabIndex        =   4
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&chließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   8
         Left            =   13800
         TabIndex        =   2
         Top             =   30
         Width           =   1260
      End
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
   Begin VB.Menu mnu_Dat 
      Caption         =   "Datei"
      Begin VB.Menu mnu_close 
         Caption         =   "&Schließen"
      End
   End
   Begin VB.Menu mnu_Bearb 
      Caption         =   "&Bearbeiten"
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "S&peichern"
         Enabled         =   0   'False
         Index           =   0
         Begin VB.Menu mnuSpeichern 
            Caption         =   "Speichern"
            Index           =   0
         End
         Begin VB.Menu mnuSpeichern 
            Caption         =   "Als neu speichern"
            Index           =   1
         End
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "S&tomo"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "Zeile &entfernen"
         Index           =   2
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "&Zeile einfügen"
         Index           =   3
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "&Drucken"
         Enabled         =   0   'False
         Index           =   4
         Begin VB.Menu mnu_Drucken_U 
            Caption         =   "Vorschau*"
            Index           =   0
         End
         Begin VB.Menu mnu_Drucken_U 
            Caption         =   "Ablage*"
            Index           =   1
         End
         Begin VB.Menu mnu_Drucken_U 
            Caption         =   "Druck*"
            Index           =   2
         End
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "&Vorschau"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "Übernehmen"
         Index           =   6
         Begin VB.Menu mnu_Such 
            Caption         =   "Angebot"
            Index           =   0
            Begin VB.Menu mnuUbernehmenA 
               Caption         =   "Ungedruckt"
               Index           =   0
            End
            Begin VB.Menu mnuUbernehmenA 
               Caption         =   "Gedruckt"
               Index           =   1
            End
            Begin VB.Menu mnuUbernehmenA 
               Caption         =   "Vorlage"
               Index           =   2
            End
         End
         Begin VB.Menu mnu_Such 
            Caption         =   "Auftragsbestätigung"
            Index           =   1
            Begin VB.Menu mnuUbernehmenB 
               Caption         =   "Ungedruckt"
               Index           =   0
            End
            Begin VB.Menu mnuUbernehmenB 
               Caption         =   "Gedruckt"
               Index           =   1
            End
            Begin VB.Menu mnuUbernehmenB 
               Caption         =   "Vorlage"
               Index           =   2
            End
         End
         Begin VB.Menu mnu_Such 
            Caption         =   "Rechnung"
            Index           =   2
            Begin VB.Menu mnuUbernehmenR 
               Caption         =   "Ungedruckt"
               Index           =   0
            End
            Begin VB.Menu mnuUbernehmenR 
               Caption         =   "Gedruckt"
               Index           =   1
            End
            Begin VB.Menu mnuUbernehmenR 
               Caption         =   "Vorlage"
               Index           =   2
            End
         End
         Begin VB.Menu mnu_Such 
            Caption         =   "Gutschrift"
            Index           =   3
            Begin VB.Menu mnuUbernehmenG 
               Caption         =   "Ungedruckt"
               Index           =   0
            End
            Begin VB.Menu mnuUbernehmenG 
               Caption         =   "Gedruckt"
               Index           =   1
            End
            Begin VB.Menu mnuUbernehmenG 
               Caption         =   "Vorlage"
               Index           =   2
            End
         End
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "&Leeren"
         Index           =   10
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnu_Bearb_U 
         Caption         =   "Designer"
         Enabled         =   0   'False
         Index           =   12
      End
   End
   Begin VB.Menu mnuAnsicht 
      Caption         =   "&Ansicht"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuAnsicht_ResFak 
         Caption         =   "Fenstergröße &1 (Standard)"
         Index           =   0
      End
      Begin VB.Menu mnuAnsicht_ResFak 
         Caption         =   "Fenstergröße &2"
         Index           =   2
      End
      Begin VB.Menu mnuAnsicht_ResFak 
         Caption         =   "Fenstergröße &3"
         Index           =   4
      End
      Begin VB.Menu mnuAnsicht_ResFak 
         Caption         =   "Fenstergröße &4"
         Index           =   6
      End
      Begin VB.Menu mnuAnsicht_ResFak 
         Caption         =   "Fenstergröße &5"
         Index           =   8
      End
      Begin VB.Menu mnugrenze_ansicht 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAnsicht_Prop 
         Caption         =   "&Proportional"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAnsicht_Alle 
         Caption         =   "&Alle Unterfenster"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnugrenze_ansicht2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAnsicht_ResetPosition 
         Caption         =   "&Fensterposition Reset"
      End
   End
End
Attribute VB_Name = "frmSP52831"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public frmParent                  As frmSP52830
  
Private gFPArray()                As Variant

Private gFPArray1()               As Variant

Private gbZeileInCopy             As Boolean

Private gvrnMerker                As Variant
  
Private objSQLAusw                As SPSQLAuswahl.clsSQLAuswahl

Private objSQLAuswDef             As SPSQLAuswahl.clsSQLAuswahl

Private objHlp                    As SpHlp.clsHlp

Private objPRM                    As clsPRM
  
Private gsngKopfHeight            As Single

Private gsngPostenHeight          As Single

Private gsngFussHeight            As Single

Private gsngFussTop               As Single

Private gbAfterLoad               As Boolean
  
Private COL_MENGE_STR             As String

Private COL_EINHEIT_STR           As String

Private COL_EPREIS_STR            As String

Private COL_RABATT_STR            As String

Private COL_UST_STR               As String

Private COL_GPREISDUMMY_O_UST_STR As String

Private COL_GPREISDUMMY_M_UST_STR As String

Const COL_ZEILENART = 1

Const COL_ARTSCHL = 2

Const COL_ARTIKEL = 3

Const COL_LS_DATUM = 4

Const COL_LS_NUMMER = 5
  
Const COL_MENGE = 6

Const COL_EINHEIT = 7

Const COL_EPREIS = 8

Const COL_RABATT = 9

Const COL_GPREIS = 10
  
Const COL_UST = 11

Const COL_DURCHLAUFEND = 12

Const COL_KOSTSCHL = 13

Const COL_SACHSCHL = 14

Const COL_KOSTKTO = 15

Const COL_SACHKTO = 16

Const COL_GPREISDUMMY = 17

Const COL_GPREISDUMMY_O_UST = 18

Const COL_GPREISDUMMY_M_UST = 19

Const COL_SUMMEN = 20

Const COL_ERSTDAT = 21

Const COL_ERSTVON = 22

Const COL_AENDDAT = 23

Const COL_AENDVON = 24
  
Const COL_LASTEDIT = 16

Const COL_LAST = 24

Const COL_FS_STEUERPFLICHT = 2

Const COL_FS_UST = 3

Const COL_FS_STEUERFREI = 4

Const COL_FS_WRG = 5

Const COL_FS_GESAMT = 6

Const C_STR_FORMELTEXT_SCHL = "FORMEL"

Enum MaskeLeerenModus
    
    alleDaten = 0
    nurBelegDaten = 1
    ohneBelegDaten = 2
    
End Enum

Dim cReSize        As FormResize

Private m_BelegNeu As Boolean

Private strAutomatischText

Private blnChangeEventFired As Boolean
 
Private dictButtonsState    As New Dictionary

Private dictFormelRows      As New Dictionary

Private dictLSRows          As New Dictionary

Dim ShiftGedrueckt          As Boolean

Dim boolEnterGedruckt       As Boolean

Dim SelectAllText           As Boolean

Dim intOldSteuerTyp         As Integer

Public lngTEMPChangeCount   As Long

Private Sub cmd1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

100     If ShiftGedrueckt = True Then ShiftGedrueckt = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
100     If Shift > 0 Then

105         ShiftGedrueckt = True

        End If

End Sub

Private Sub fpSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)

        On Error GoTo Fehler

        Dim knz As Variant
        
100     Select Case Index
        
            Case 1
            
105             If fpSpread1(1).GetText(COL_ZEILENART, Row, knz) Then
            
110                 If Trim$(knz) = "T" And Col >= COL_LS_DATUM And Col <= COL_RABATT Then SetActiveCellExt 1, COL_ARTIKEL, Row, True
            
                End If
        
        End Select

        Exit Sub

Fehler:
115     Me.MousePointer = vbDefault
120     Call FehlerErklärung("frmSP52831", "fpSpread1_Click()")

End Sub

Private Sub fpSpread1_ColWidthChange(Index As Integer, _
                                     ByVal Col1 As Long, _
                                     ByVal Col2 As Long)

        On Error GoTo Fehler

100     Select Case Index

            Case 1
    
105             txt1(0).width = TEXT_BREITE + fpSpread1(1).colWidth(COL_LS_DATUM) + fpSpread1(1).colWidth(COL_LS_NUMMER) + fpSpread1(1).colWidth(COL_MENGE) + fpSpread1(1).colWidth(COL_EINHEIT) + fpSpread1(1).colWidth(COL_EPREIS) + fpSpread1(1).colWidth(COL_RABATT)
  
        End Select

        Exit Sub

Fehler:
110     Me.MousePointer = vbDefault
115     Call FehlerErklärung("frmSP52831", "fpSpread1_ColWidthChange()")

End Sub

Private Sub fpSpread1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

        On Error GoTo Fehler
        
        Dim NeuRow As Long
  
100     If Index = 1 Then

105         Call SetCellBackColor(1, fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow, True)

110         Select Case KeyCode

                Case vbKeyDown
115                 fpSpread1(1).Row = fpSpread1(1).ActiveRow
120                 fpSpread1(1).Col = fpSpread1(1).ActiveCol

125                 If fpSpread1(1).CellType = CellTypeComboBox Then
130                     If fpSpread1(1).ActiveRow = fpSpread1(1).MaxRows Then
135                         NeuRow = fpSpread1(1).ActiveRow
                        Else
140                         NeuRow = fpSpread1(1).ActiveRow + 1
                        End If

145                     SetActiveCellExt 1, fpSpread1(1).Col, NeuRow, True

                    End If

150             Case vbKeyUp

155                 fpSpread1(1).Row = fpSpread1(1).ActiveRow
160                 fpSpread1(1).Col = fpSpread1(1).ActiveCol

165                 If fpSpread1(1).CellType = CellTypeComboBox Then
170                     If fpSpread1(1).ActiveRow = 1 Then
175                         NeuRow = fpSpread1(1).ActiveRow
                        Else
180                         NeuRow = fpSpread1(1).ActiveRow - 1
                        End If

185                     SetActiveCellExt 1, fpSpread1(1).Col, NeuRow, True

                    End If

            End Select

        End If

        Exit Sub

Fehler:
190     Call FehlerErklärung("frmSP57702", "fpSpread1_KeyUp")

End Sub

Private Sub fpSpread1_EditMode(Index As Integer, _
                               ByVal Col As Long, _
                               ByVal Row As Long, _
                               ByVal Mode As Integer, _
                               ByVal ChangeMade As Boolean)

        On Error GoTo Fehler
    
100     If Mode = 1 Then
        
105         Call SetCellBackColor(Index, Col, Row, False)
        
            Dim cellText As Variant

110         If fpSpread1(Index).GetText(Col, Row, cellText) Then
            
115             If Len(cellText) > 0 Then

120                 Timer1.Enabled = True

125                 Timer1.Interval = 2

130                 SelectAllText = True

                End If
                
            End If

        End If

        Exit Sub

Fehler:
    
135     Call FehlerErklärung("frmSP52831", "fpSpread1_EditMode()")

End Sub

Public Property Get BelegNeu() As Boolean

100     BelegNeu = m_BelegNeu

End Property

Public Property Let BelegNeu(ByVal bBelegNeu As Boolean)

100     m_BelegNeu = bBelegNeu

End Property

Private Sub cmd1_Click(Index As Integer)

        On Error GoTo Fehler

        Dim Row    As Long

        Dim Anzahl As Long
                
100     If Index = 4 Then

105         SaveToArchiveResult.ErrorNumber = 0

110         If ReadyForPrint(GsAnwenderNr, GsHauptPfad) = False Then

                Exit Sub
            
            End If
            
        End If
        
115     Call RefreshFoot
        
120     Select Case Index

            Case 3
                
125             fpSpread1(1).ReDraw = False
                
130             Call SetCellBackColor(1, fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow, True)
                
135             Row = fpSpread1(1).ActiveRow

140             Anzahl = LastRow - Row + 1

145             If Anzahl > 0 Then

150                 ZeileBearbeiten 2, Row, Anzahl
155                 ZeileBearbeiten 3, Row + 1, 0

160                 Call SetCellBackColor(1, fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow, True)
                    
165                 fpSpread1(1).SetActiveCell 1, Row
                    
                End If

170             fpSpread1(1).ReDraw = True

175         Case 2

180             fpSpread1(1).ReDraw = False
                
185             Call SetCellBackColor(1, fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow, True)

190             Row = fpSpread1(1).ActiveRow

195             Anzahl = LastRow - Row + 1

200             If Anzahl > 0 Then

205                 ZeileBearbeiten 2, Row + 1, Anzahl
210                 ZeileBearbeiten 3, Row, 0

215                 ZwischenSummeRefresh Row
                    
220                 LSArtikelGefunden
                    
225                 Call SetCellBackColor(1, fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow, False)
                    
230                 fpSpread1(1).SetActiveCell 1, Row
                    
                End If
                
235             fpSpread1(1).ReDraw = True

240         Case 1

245             Storno

250         Case 0

255             Me.PopupMenu mnu_Bearb_U(0), 0, cmd1(Index).left, cmd1(Index).top + cmd1(Index).height + SSPanel1(0).top

260         Case 12

265             LLDesigner Me, LL1, CLng(frmParent.glngBelegID), 1, False

270         Case 5
                
275             intOldSteuerTyp = intSteuerTyp

280             frmParent.mnuOpt1(4).Enabled = False
                
285             Me.Hide

290             frmParent.Show 0

295         Case 4
                
300             Me.PopupMenu mnu_Bearb_U(4), , cmd1(Index).left, cmd1(Index).top + cmd1(Index).height + SSPanel1(0).top
                
305         Case 8

310             Unload Me
                
315         Case 9
                
320             If ShiftGedrueckt Then

325                 objHlp.HlpShow HlpWrite, "allgemein"
                 
                Else
                
330                 objHlp.HlpShow HlpRead, "allgemein"
                    
                End If
                
335             ShiftGedrueckt = False

340         Case 10

345             If MsgBox(GetMessage(523), vbYesNo + vbQuestion, strMeldungCap) = vbYes Then
                    
350                 fpSpread1(1).ReDraw = False
                    
355                 Call SetCellBackColor(1, fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow, True)
                
360                 cmd1(10).SetFocus
                
365                 Screen.MousePointer = vbHourglass
                
370                 Call MaskeLeeren(ohneBelegDaten)
                
375                 Screen.MousePointer = vbDefault
                    
380                 fpSpread1(1).ReDraw = True
                    
                End If

385         Case 6
       
390             Me.PopupMenu mnu_Bearb_U(6), , cmd1(Index).left, cmd1(Index).top + cmd1(Index).height + SSPanel1(0).top

        End Select

        Exit Sub

Fehler:

395     If printJobInProgress Then Call EnableButtonsOverPrintJob(True)
        
400     printJobInProgress = False

405     Call FehlerErklärung("frmSP52831", "cmd1_Click")

End Sub

Public Sub Storno()

        On Error GoTo Fehler

        Dim sql As String

        Dim rs  As ADODB.Recordset
  
100     If MsgBox("Soll der Beleg wirklich storniert werden?", vbYesNo + vbQuestion, strMeldungCap) = vbYes Then

105         Set rs = New ADODB.Recordset
        
110         If frmParent.gintZwAblage = 1 Then

115             sql = "SELECT BelegNr,Storno,StornoDatum,AendDat,AendVon FROM [2800_Haupt] WHERE [BelegID] = " & frmParent.glngBelegIDVorlage

            Else
            
120             sql = "SELECT BelegNr,Storno,StornoDatum,AendDat,AendVon FROM [2800_Haupt] WHERE [BelegID] = " & frmParent.glngBelegID

            End If

125         OPEN_gConn

130         rs.Open sql, gConn, adOpenKeyset, adLockOptimistic
    
135         If rs.RecordCount > 0 Then

140             rs!Storno = "1"
145             rs!StornoDatum = Now
150             rs!AendDat = Now
155             rs!AendVon = GsUser
160             rs.Update
            
165             rs.MoveLast
170             Protokoll iAppend, "Storno. BelegID: " & frmParent.glngBelegID & " BelegNr: " & rs!BelegNr & " -> " & Now
            End If

175         rs.Close
    
180         stornoDone = True
185         cmd1(0).Enabled = False
    
190         Unload Me
        End If

        Exit Sub
  
Fehler:

195     Call FehlerErklärung("frmSP52831", "Storno()")

End Sub

Private Sub Form_Load()

    On Error GoTo Fehler
        
    SetUst

    Dim i As Integer

    Dim h As Long

    Set cReSize = New FormResize
    cReSize.setSectionBezeichnung = "SP52830"
    cReSize.setKeyBezeichnung = "SP52830"
    cReSize.setIstUnterFenster = True

    Set objPRM = New clsPRM
    Set objPRM.gForm = Me
    objPRM.PRM_Alle

    fpSpread1(0).NoBeep = True
    fpSpread1(1).NoBeep = True
    fpSpread1(2).NoBeep = True

    fpSpread1(0).GrayAreaBackColor = RGB(235, 229, 217)
    fpSpread1(1).GrayAreaBackColor = RGB(235, 229, 217)
    fpSpread1(2).GrayAreaBackColor = RGB(235, 229, 217)

    fpSpread1(1).ShadowColor = RGB(235, 229, 217)

    LL1.LlSetOptionString LL_OPTIONSTR_LICENSINGINFO, "4yi/HQ"
    LL1.LlSetOption LL_OPTION_INCREMENTAL_PREVIEW, False
    LL1.LlSetOption LL_OPTION_RIBBON_DEFAULT_ENABLEDSTATE, 0
    LL1.LlSetOption LL_OPTION_INCLUDEFONTDESCENT, False
    
    LL1.LlSetOption LL_OPTION_CONVERTCRLF, True
    LL1.LlSetPrinterDefaultsDir ArbeitsplatzPfad
    LL1.LlSetOption LL_OPTION_NOPARAMETERCHECK, 1
    LL1.LlPreviewSetTempPath ArbeitsplatzPfad

    gbAfterLoad = False
  
    Me.left = 50
    Me.top = 500
    
    Me.caption = frmParent.caption

    COL_MENGE_STR = Chr(64 + COL_MENGE)
    COL_EINHEIT_STR = Chr(64 + COL_EINHEIT)
    COL_EPREIS_STR = Chr(64 + COL_EPREIS)
    COL_RABATT_STR = Chr(64 + COL_RABATT)
    COL_UST_STR = Chr(64 + COL_UST)
    COL_GPREISDUMMY_O_UST_STR = Chr(64 + COL_GPREISDUMMY_O_UST)
    COL_GPREISDUMMY_M_UST_STR = Chr(64 + COL_GPREISDUMMY_M_UST)

    fpSpread1(1).colWidth(COL_LS_DATUM) = 0
    fpSpread1(1).colWidth(COL_LS_NUMMER) = 0
        
    KopfSetUp

    KopfFuellen

    For i = 1 To fpSpread1(0).MaxRows
        h = h + fpSpread1(0).rowHeight(i)
    Next i

    gsngKopfHeight = h + 300
    fpSpread1(0).height = gsngKopfHeight

    fpSpread1(1).ReDraw = False
    fpSpread1(1).top = gsngKopfHeight

    gsngPostenHeight = 4300
    fpSpread1(1).height = gsngPostenHeight

    PostenSetUp

    h = 0

    For i = 1 To fpSpread1(1).MaxCols
    
        h = h + fpSpread1(1).colWidth(i)
        
    Next i

    Me.width = h + 580

    If GesamtIstBrutto Then
    
        FussSetUpGesamtIstBrutto
        FussFuellenGesamtIstBrutto
        
    Else
    
        FussSetUp
        FussFuellen
        
    End If

    fpSpread1(1).ReDraw = True
        
    h = 0

    For i = 1 To fpSpread1(2).MaxRows
    
        h = h + fpSpread1(2).rowHeight(i)
        
    Next i

    gsngFussHeight = h + 400
    fpSpread1(2).height = gsngFussHeight
    fpSpread1(2).top = fpSpread1(1).top + fpSpread1(1).height

    gsngFussTop = fpSpread1(2).top
    gsngFussHeight = fpSpread1(2).height

    SizeControls GetSetting("SP50000", "SP52800", "SP52831Split0", fpSpread1(0).height), 0
    SizeControls GetSetting("SP50000", "SP52800", "SP52831Split1", fpSpread1(2).top), 1

    SetXPSize Me

    Me.width = GetSetting("SP50000", "SP52800", "SP52831Width", "10935")
    Me.height = GetSetting("SP50000", "SP52800", "SP52831Height", gsngFussTop + gsngFussHeight + SSPanel1(0).height + 450)

    fpSpread1(0).TopRow = GetSetting("SP50000", "SP52800", "SP52831TopRow0", "0")

    imgSplitter(0).left = 0
    imgSplitter(1).left = 0
    picSplitter(0).left = 0
    picSplitter(1).left = 0

    txt1(0).width = TEXT_BREITE + fpSpread1(1).colWidth(COL_LS_DATUM) + fpSpread1(1).colWidth(COL_LS_NUMMER) + fpSpread1(1).colWidth(COL_MENGE) + fpSpread1(1).colWidth(COL_EINHEIT) + fpSpread1(1).colWidth(COL_EPREIS) + fpSpread1(1).colWidth(COL_RABATT)

    Set objSQLAusw = New SPSQLAuswahl.clsSQLAuswahl
    objSQLAusw.FilterBar = True

    Set objSQLAuswDef = New SPSQLAuswahl.clsSQLAuswahl
    gbAfterLoad = True

    Set objHlp = New SpHlp.clsHlp
    objHlp.DatabaseName = GsHauptPfadLokal & "hlp\SP50000.hlp"
    objHlp.table = Me.name
    objHlp.caption = Me.name & " - Feldhilfe"

    If gbAfterLoad Then

        fpSpread1(0).width = Me.ScaleWidth
        fpSpread1(1).width = Me.ScaleWidth
        fpSpread1(2).width = Me.ScaleWidth

        picSplitter(0).width = Me.ScaleWidth
        picSplitter(1).width = Me.ScaleWidth

        imgSplitter(0).width = Me.ScaleWidth
        imgSplitter(1).width = Me.ScaleWidth
        gsngFussTop = Me.ScaleHeight - gsngFussHeight - SSPanel1(0).height

        SizeControls gsngFussTop, 1
        fpSpread1(2).height = Me.ScaleHeight - gsngFussTop - SSPanel1(0).height - imgSplitter(1).height

        If Me.height <= 7000 Then

            Me.height = 7000

        End If
            
    End If
    
    h = 0

    For i = 1 To fpSpread1(1).MaxCols
    
        h = h + fpSpread1(2).colWidth(i)
        
    Next i

    Me.width = h + 2600
    
    cReSize.Form = Me
    
    cReSize.IgnoreTrueDBGridInfo = True

    Call readWindowPos(Me, "SP52800", "SP52831Left", "SP52831Top")

    Protokoll iAppend, ">ERFASSUNG FENSTER OEFFNEN, ZEIT: -> " & Now

    Set frmParentDE = Me

    Exit Sub

Fehler:

    If Err.number = 384 Then

        Resume Next

    Else
        
        Call FehlerErklärung("frmSP52831", "Form_Load()")

    End If

End Sub

Sub SizeControls(Y As Single, SpliterIndex As Integer)

        On Error Resume Next

        Dim Dif As Integer

100     imgSplitter(SpliterIndex).top = Y
    
105     If SpliterIndex = 0 Then
110         fpSpread1(SpliterIndex).height = Y

115         Dif = fpSpread1(SpliterIndex + 1).top - Y
120         fpSpread1(SpliterIndex + 1).top = Y + picSplitter(SpliterIndex).height
125         fpSpread1(SpliterIndex + 1).height = fpSpread1(SpliterIndex + 1).height + Dif - picSplitter(SpliterIndex).height
        Else
130         fpSpread1(SpliterIndex).height = Y - fpSpread1(SpliterIndex).top

135         Dif = fpSpread1(SpliterIndex + 1).top - Y
140         fpSpread1(SpliterIndex + 1).top = Y + picSplitter(SpliterIndex).height
145         fpSpread1(SpliterIndex + 1).height = fpSpread1(SpliterIndex + 1).height + Dif - picSplitter(SpliterIndex).height
        End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        Dim dialogResult As Long

        On Error GoTo Fehler
        
100     If printJobInProgress Then
           
105         Cancel = True
           
110         Protokoll iAppend, ">>>>> ERFASSUNG FENSTER SCHLIESSEN WAEHREND DRUCK/ARC JOB(UnloadMode = " & CStr(UnloadMode) & "). BelegID: " & frmParent.glngBelegID & ", ZEIT: -> " & Now

            Exit Sub

        Else
        
115         Protokoll iAppend, ">ERFASSUNG FENSTER SCHLIESSEN(UnloadMode = " & CStr(UnloadMode) & "). BelegID: " & frmParent.glngBelegID & ", ZEIT: -> " & Now

        End If
        
120     Call writeWindowPos(Me, "SP52800", "SP52831Left", "SP52831Top")
        
125     If WindowPosition(Me) Then

130         SaveSetting "SP50000", "SP52800", "SP52831Width", Me.width / cReSize.CurrScaleFactorWidth
135         SaveSetting "SP50000", "SP52800", "SP52831Height", Me.height / cReSize.CurrScaleFactorHeight
140         SaveSetting "SP50000", "SP52800", "SP52831Split0", imgSplitter(0).top
145         SaveSetting "SP50000", "SP52800", "SP52831Split1", imgSplitter(1).top
150         SaveSetting "SP50000", "SP52800", "SP52831TopRow0", fpSpread1(0).TopRow

        End If
        
155     If LastRow > 0 Then

160         If fpSpread1(1).ActiveCol = COL_EPREIS Then

165             If PlausiEPreis = False Then

170                 Cancel = True

                    Exit Sub

                End If
                
            End If

175         If cmd1(0).Enabled Then

180             dialogResult = MsgBox(GetMessage(18), vbYesNoCancel + vbQuestion + vbDefaultButton3, strMeldungCap)

185             Select Case dialogResult

                    Case vbYes

190                     If Plausi Then

195                         If frmParent.glngBelegID = 0 Then
200                             frmParent.Speichern
                            Else
205                             frmParent.Speichern frmParent.glngBelegID
                            End If

                        Else
                        
210                         Cancel = True

                            Exit Sub
                            
                        End If

215                 Case vbNo

220                 Case vbCancel

225                     Cancel = True

                        Exit Sub
    
                End Select

            End If

        End If
              
230     If printDone Or stornoDone Then

235         frmParent.glngBelegID = 0
240         frmParent.glngBelegIDVorlage = 0
245         frmParent.sta1.Panels(3).text = ""
            
250         frmParent.MaskeLeeren (False)

255         frmParent.txt1(0) = ""

260         If frmParent.txt1(15).Visible Then
                
265             frmParent.txt1(15) = ""
270             frmParent.BelegNr = ""
275             frmParent.txt1(16) = GdtDatum
280             frmParent.belegDatum = GdtDatum

            End If

285         If printDone = True Then printDone = False
290         If stornoDone = True Then stornoDone = False

        Else
            
295         frmParent.ValutaDatum = ""
            
        End If
        
300     Me.Hide
        
305     frmParent.mnuOpt1(4).Enabled = True

310     If Not frmParent.Visible Then frmParent.Show 0

        Exit Sub

Fehler:
315     Call FehlerErklärung("frmSP52831", "Form_QueryUnload")

End Sub

Private Sub Form_Resize()
    
    On Error GoTo Fehler
    
    If gbAfterLoad Then
        
        fpSpread1(0).width = Me.ScaleWidth
        
        fpSpread1(1).width = Me.ScaleWidth
    
        fpSpread1(2).width = Me.ScaleWidth

        picSplitter(0).width = Me.ScaleWidth
        picSplitter(1).width = Me.ScaleWidth

        imgSplitter(0).width = Me.ScaleWidth
        imgSplitter(1).width = Me.ScaleWidth
        gsngFussTop = Me.ScaleHeight - gsngFussHeight - SSPanel1(0).height

        SizeControls gsngFussTop, 1
        fpSpread1(2).height = Me.ScaleHeight - gsngFussTop - SSPanel1(0).height - imgSplitter(1).height

        If Me.height <= 7000 Then
            Me.height = 7000
        End If
    End If

    Exit Sub

Fehler:

    If Err.number = 384 Then
        
        Resume Next

    Else
        Call FehlerErklärung("frmSP52831", "Form_Resize")
    End If
    
End Sub

Sub KopfSetUp()

        On Error GoTo Fehler
    
100     fpSpread1(0).Row = -1
105     fpSpread1(0).Col = -1
110     fpSpread1(0).Lock = True
    
115     fpSpread1(0).EditModePermanent = True
    
120     fpSpread1(0).TypeMaxEditLen = 110
    
125     fpSpread1(0).MaxCols = 8
130     fpSpread1(0).MaxRows = 19
    
135     fpSpread1(0).Row = -1
140     fpSpread1(0).Col = -1
145     fpSpread1(0).fontSize = 8
    
150     fpSpread1(0).GridShowHoriz = False
155     fpSpread1(0).GridShowVert = False
    
160     fpSpread1(0).AllowCellOverflow = True
    
165     fpSpread1(0).ColHeadersShow = False
170     fpSpread1(0).RowHeadersShow = False
    
175     fpSpread1(0).colWidth(1) = 2000
180     fpSpread1(0).colWidth(2) = 1500
185     fpSpread1(0).colWidth(3) = 1500
190     fpSpread1(0).colWidth(4) = 1500
195     fpSpread1(0).colWidth(5) = 2400
200     fpSpread1(0).colWidth(6) = 1500
205     fpSpread1(0).colWidth(7) = 1500
210     fpSpread1(0).colWidth(8) = 1250
    
215     fpSpread1(0).rowHeight(6) = 600
220     fpSpread1(0).rowHeight(7) = 300
    
225     fpSpread1(0).rowHeight(14) = 400
230     fpSpread1(0).rowHeight(15) = 400
235     fpSpread1(0).rowHeight(16) = 300
240     fpSpread1(0).rowHeight(17) = 300
245     fpSpread1(0).rowHeight(18) = 300
    
250     fpSpread1(0).Row = 7
255     fpSpread1(0).Col = 2
260     fpSpread1(0).fontSize = 8

265     fpSpread1(0).Row = 15
270     fpSpread1(0).Col = 5
275     fpSpread1(0).fontSize = 11
    
280     fpSpread1(0).FontBold = True
    
285     fpSpread1(0).Col = 1
290     fpSpread1(0).Row = 1
295     fpSpread1(0).Col2 = 1
300     fpSpread1(0).Row2 = fpSpread1(0).MaxRows
305     fpSpread1(0).BlockMode = True
310     fpSpread1(0).BackColor = RGB(235, 229, 217)
315     fpSpread1(0).BlockMode = False

        Exit Sub
        
Fehler:
        
320     Me.MousePointer = vbDefault
325     Call FehlerErklärung("frmSP52831", "KopfSetUp()")

End Sub

Sub FussSetUp()

    On Error GoTo Fehler
        
    Call SetUst
    
    fpSpread1(2).Row = -1
    fpSpread1(2).Col = -1
    fpSpread1(2).Lock = True
    fpSpread1(2).TypeHAlign = TypeHAlignRight
    
    fpSpread1(2).EditModePermanent = True

    fpSpread1(2).TypeMaxEditLen = 110
        
    fpSpread1(2).MaxCols = 6
    fpSpread1(2).MaxRows = 6
    
    If GintBelegArt = 2 Or GintBelegArt = 3 Then
        fpSpread1(2).MaxRows = 5
        fpSpread1(2).Row = 5
        fpSpread1(2).RowHidden = True
    End If
    
    fpSpread1(2).Row = -1
    fpSpread1(2).Col = -1
    fpSpread1(2).fontSize = 8
    
    fpSpread1(2).GridShowHoriz = False
    fpSpread1(2).GridShowVert = False
    
    fpSpread1(2).AllowCellOverflow = True
    
    fpSpread1(2).ColHeadersShow = False
    fpSpread1(2).RowHeadersShow = False

    fpSpread1(2).colWidth(1) = 2085
    fpSpread1(2).colWidth(COL_FS_STEUERPFLICHT) = 2265
    fpSpread1(2).colWidth(COL_FS_UST) = 2255
    fpSpread1(2).colWidth(COL_FS_STEUERFREI) = 2465
    fpSpread1(2).colWidth(COL_FS_WRG) = 965
    fpSpread1(2).colWidth(COL_FS_GESAMT) = 3115

    fpSpread1(2).Col = COL_FS_STEUERPFLICHT
    fpSpread1(2).Row = 2
    fpSpread1(2).Col2 = COL_FS_STEUERFREI
    fpSpread1(2).Row2 = 3
    fpSpread1(2).BlockMode = True

    fpSpread1(2).CellType = CellTypeNumber
    fpSpread1(2).TypeNumberShowSep = True
    fpSpread1(2).TypeNumberDecPlaces = 2
    fpSpread1(2).BlockMode = False

    fpSpread1(2).Col = COL_FS_GESAMT
    fpSpread1(2).Row = 2
    fpSpread1(2).Col2 = COL_FS_GESAMT
    fpSpread1(2).Row2 = 3
    fpSpread1(2).BlockMode = True
    fpSpread1(2).CellType = CellTypeNumber
    fpSpread1(2).TypeNumberShowSep = True
    fpSpread1(2).TypeNumberDecPlaces = 2
    fpSpread1(2).BlockMode = False
    
    fpSpread1(2).Col = COL_FS_UST
    fpSpread1(2).Row = 2

    fpSpread1(2).Formula = "ROUND(B2*" & SQLZahl(dblUstSatz) & "/100" & ",2)"
    
    fpSpread1(2).Col = COL_FS_GESAMT
    fpSpread1(2).Row = 2
    
    fpSpread1(2).Formula = "B2+C2+D2"
    
    fpSpread1(2).Col = COL_FS_STEUERFREI
    fpSpread1(2).Row = 4
    fpSpread1(2).fontSize = 7
    
    fpSpread1(2).Col = COL_FS_GESAMT
    fpSpread1(2).Row = 4
    fpSpread1(2).fontSize = 7
    
    If Trim(frmParent.Check1(1)) = 1 And Trim(UCase(frmParent.txt1(17))) <> Trim(UCase(frmParent.txt1(18))) Then

        fpSpread1(2).Col = COL_FS_STEUERPFLICHT
        fpSpread1(2).Row = 3
        fpSpread1(2).Formula = "ROUND(B2*" & SQLZahl(frmParent.txt1(19)) & ",2)"

        fpSpread1(2).Col = COL_FS_UST
        fpSpread1(2).Row = 3
        fpSpread1(2).Formula = "ROUND(C2*" & SQLZahl(frmParent.txt1(19)) & ",2)"

        fpSpread1(2).Col = COL_FS_STEUERFREI
        fpSpread1(2).Row = 3
        fpSpread1(2).Formula = "ROUND(D2*" & SQLZahl(frmParent.txt1(19)) & ",2)"
        
        fpSpread1(2).Col = COL_FS_GESAMT
        fpSpread1(2).Row = 3
        
        fpSpread1(2).Formula = "B3+C3+D3"
        fpSpread1(2).rowHeight(3) = 300

    Else

        fpSpread1(2).rowHeight(3) = 0
        fpSpread1(2).rowHeight(2) = 300

    End If

    fpSpread1(2).rowHeight(4) = 300
    fpSpread1(2).rowHeight(5) = 220
    fpSpread1(2).rowHeight(6) = 220

    fpSpread1(2).Col = 1
    fpSpread1(2).Row = 1
    fpSpread1(2).Col2 = fpSpread1(0).MaxCols
    fpSpread1(2).Row2 = 1
    fpSpread1(2).BlockMode = True
    fpSpread1(2).FontBold = True
    fpSpread1(2).BlockMode = False
    
    fpSpread1(2).Col = COL_FS_STEUERPFLICHT
    fpSpread1(2).Row = 4
    
    fpSpread1(2).Col2 = COL_FS_STEUERFREI
    fpSpread1(2).Row2 = 6
    fpSpread1(2).BlockMode = True
    
    fpSpread1(2).CellType = CellTypeStaticText
    fpSpread1(2).TypeHAlign = TypeHAlignLeft
    fpSpread1(2).BlockMode = False
    
    fpSpread1(2).Col = COL_FS_STEUERFREI
    fpSpread1(2).Row = 4

    fpSpread1(2).Col2 = COL_FS_GESAMT
    fpSpread1(2).Row2 = 5
    fpSpread1(2).Col2 = COL_FS_GESAMT
    fpSpread1(2).TypeHAlign = TypeHAlignRight
    fpSpread1(2).Row2 = 6
    fpSpread1(2).BlockMode = True
    fpSpread1(2).CellType = CellTypeStaticText
    fpSpread1(2).TypeHAlign = TypeHAlignRight
    fpSpread1(2).BlockMode = False
    
    fpSpread1(2).Col = 1
    fpSpread1(2).Row = 1
    fpSpread1(2).Col2 = 1
    fpSpread1(2).Row2 = fpSpread1(0).MaxRows
    fpSpread1(2).BlockMode = True
    fpSpread1(2).BackColor = RGB(235, 229, 217)
    fpSpread1(2).BlockMode = False
    
    fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 1, fpSpread1(2).MaxCols, fpSpread1(2).MaxRows, 16, vbBlack, CellBorderStyleSolid
    fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 4, fpSpread1(2).MaxCols, fpSpread1(2).MaxRows - 1, 4, vbBlack, CellBorderStyleSolid
    fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 1, COL_FS_STEUERFREI, 3, 2, vbBlack, CellBorderStyleSolid

    Exit Sub

Fehler:

    Me.MousePointer = vbDefault
    Call FehlerErklärung("frmSP52831", "FussSetUp()")

End Sub

Sub FussSetUpGesamtIstBrutto()

        On Error GoTo Fehler
        
100     SetUst
    
105     fpSpread1(2).Row = -1
110     fpSpread1(2).Col = -1
115     fpSpread1(2).Lock = True
120     fpSpread1(2).TypeHAlign = TypeHAlignRight
    
125     fpSpread1(2).EditModePermanent = True
    
130     fpSpread1(2).TypeMaxEditLen = 110
    
135     fpSpread1(2).MaxCols = 6
140     fpSpread1(2).MaxRows = 6
    
145     fpSpread1(2).Row = -1
150     fpSpread1(2).Col = -1
155     fpSpread1(2).fontSize = 8
    
160     fpSpread1(2).GridShowHoriz = False
165     fpSpread1(2).GridShowVert = False
    
170     fpSpread1(2).AllowCellOverflow = True
    
175     fpSpread1(2).ColHeadersShow = False
180     fpSpread1(2).RowHeadersShow = False
    
185     fpSpread1(2).colWidth(1) = 1140
190     fpSpread1(2).colWidth(COL_FS_STEUERPFLICHT) = 2220
195     fpSpread1(2).colWidth(COL_FS_UST) = 2220
200     fpSpread1(2).colWidth(COL_FS_STEUERFREI) = 2420
205     fpSpread1(2).colWidth(COL_FS_WRG) = 920
210     fpSpread1(2).colWidth(COL_FS_GESAMT) = 3070
    
215     fpSpread1(2).Col = COL_FS_STEUERPFLICHT
220     fpSpread1(2).Row = 2

225     fpSpread1(2).Formula = "ROUND(((F2-D2) / ((100 + " & SQLZahl(dblUstSatz) & ") / 100) ),2)"
    
230     fpSpread1(2).Col2 = COL_FS_STEUERFREI
235     fpSpread1(2).Row2 = 3
240     fpSpread1(2).BlockMode = True
    
245     fpSpread1(2).CellType = CellTypeNumber
250     fpSpread1(2).TypeNumberShowSep = True
255     fpSpread1(2).TypeNumberDecPlaces = 2
260     fpSpread1(2).BlockMode = False
    
265     fpSpread1(2).Col = COL_FS_STEUERPFLICHT
270     fpSpread1(2).Row = 2
275     fpSpread1(2).Col2 = COL_FS_STEUERFREI
280     fpSpread1(2).Row2 = 3
285     fpSpread1(2).BlockMode = True
290     fpSpread1(2).CellType = CellTypeNumber
295     fpSpread1(2).TypeNumberShowSep = True
300     fpSpread1(2).TypeNumberDecPlaces = 2
305     fpSpread1(2).BlockMode = False
    
310     fpSpread1(2).Col = COL_FS_GESAMT
315     fpSpread1(2).Row = 2
320     fpSpread1(2).Col2 = COL_FS_GESAMT
325     fpSpread1(2).Row2 = 3
330     fpSpread1(2).BlockMode = True
335     fpSpread1(2).CellType = CellTypeNumber
340     fpSpread1(2).TypeNumberShowSep = True
345     fpSpread1(2).TypeNumberDecPlaces = 2
350     fpSpread1(2).BlockMode = False
    
355     fpSpread1(2).Col = COL_FS_UST
360     fpSpread1(2).Row = 2
365     fpSpread1(2).Formula = "ROUND((F2-D2-B2) ,2)"
    
370     fpSpread1(2).Col = COL_FS_GESAMT
375     fpSpread1(2).Row = 2
    
380     fpSpread1(2).Col = COL_FS_STEUERFREI
385     fpSpread1(2).Row = 4
390     fpSpread1(2).fontSize = 7
    
395     fpSpread1(2).Col = COL_FS_GESAMT
400     fpSpread1(2).Row = 4
405     fpSpread1(2).fontSize = 7
    
410     If Trim(frmParent.Check1(1)) = 1 And Trim(UCase(frmParent.txt1(17))) <> Trim(UCase(frmParent.txt1(18))) Then

415         fpSpread1(2).Col = COL_FS_STEUERPFLICHT
420         fpSpread1(2).Row = 3
425         fpSpread1(2).Formula = "ROUND(((F2-D2) / ((100 + " & SQLZahl(dblUstSatz) & ") / 100) ),2)"

430         fpSpread1(2).Col = COL_FS_UST
435         fpSpread1(2).Row = 3
440         fpSpread1(2).Formula = "ROUND((F2-D2-B2) ,2)"

445         fpSpread1(2).Col = COL_FS_STEUERFREI
450         fpSpread1(2).Row = 3
455         fpSpread1(2).Formula = "ROUND(D2*" & SQLZahl(frmParent.txt1(19)) & ",2)"
        
460         fpSpread1(2).Col = COL_FS_GESAMT
465         fpSpread1(2).Row = 3
470         fpSpread1(2).Formula = "ROUND(F2*" & SQLZahl(frmParent.txt1(19)) & ",2)"

475         fpSpread1(2).rowHeight(3) = 300

        Else
        
480         fpSpread1(2).rowHeight(3) = 0
485         fpSpread1(2).rowHeight(2) = 300

        End If

490     fpSpread1(2).rowHeight(4) = 300
495     fpSpread1(2).rowHeight(5) = 220
500     fpSpread1(2).rowHeight(6) = 220
    
505     fpSpread1(2).Col = 1
510     fpSpread1(2).Row = 1
515     fpSpread1(2).Col2 = fpSpread1(0).MaxCols
520     fpSpread1(2).Row2 = 1
525     fpSpread1(2).BlockMode = True
530     fpSpread1(2).FontBold = True
535     fpSpread1(2).BlockMode = False
    
540     fpSpread1(2).Col = COL_FS_STEUERPFLICHT
545     fpSpread1(2).Row = 4
    
550     fpSpread1(2).Col2 = COL_FS_STEUERFREI
555     fpSpread1(2).Row2 = 6
560     fpSpread1(2).BlockMode = True
    
565     fpSpread1(2).CellType = CellTypeStaticText
570     fpSpread1(2).TypeHAlign = TypeHAlignLeft
575     fpSpread1(2).BlockMode = False
    
580     fpSpread1(2).Col = COL_FS_STEUERFREI
585     fpSpread1(2).Row = 4
    
590     fpSpread1(2).Col2 = COL_FS_GESAMT
595     fpSpread1(2).Row2 = 5
600     fpSpread1(2).Col2 = COL_FS_GESAMT
605     fpSpread1(2).TypeHAlign = TypeHAlignRight
610     fpSpread1(2).Row2 = 6
615     fpSpread1(2).BlockMode = True
620     fpSpread1(2).CellType = CellTypeStaticText
625     fpSpread1(2).TypeHAlign = TypeHAlignRight
630     fpSpread1(2).BlockMode = False
    
635     fpSpread1(2).Col = 1
640     fpSpread1(2).Row = 1
645     fpSpread1(2).Col2 = 1
650     fpSpread1(2).Row2 = fpSpread1(0).MaxRows
655     fpSpread1(2).BlockMode = True
660     fpSpread1(2).BackColor = RGB(235, 229, 217)
665     fpSpread1(2).BlockMode = False
    
670     fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 1, fpSpread1(2).MaxCols, fpSpread1(2).MaxRows, 16, vbBlack, CellBorderStyleSolid
675     fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 4, fpSpread1(2).MaxCols, fpSpread1(2).MaxRows - 1, 4, vbBlack, CellBorderStyleSolid
680     fpSpread1(2).SetCellBorder COL_FS_STEUERPFLICHT, 1, COL_FS_STEUERFREI, 3, 2, vbBlack, CellBorderStyleSolid

        Exit Sub

Fehler:

685     Me.MousePointer = vbDefault
690     Call FehlerErklärung("frmSP52831", "FussSetUpGesamtIstBrutto()")

End Sub

Public Sub KopfFuellen()

        On Error GoTo Fehler

        Dim rs  As ADODB.Recordset

        Dim knz As Variant
        
100     strAutomatischText = GetMessage(439)
        
105     fpSpread1(0).SetText 2, 1, GmandantRS!Name1
110     fpSpread1(0).SetText 2, 2, GmandantRS!Name2
115     fpSpread1(0).SetText 2, 3, GmandantRS!Straße
120     fpSpread1(0).SetText 2, 4, GmandantRS!Lkz & "-" & GmandantRS!Plz & " " & GmandantRS!Ort & " " & GmandantRS!ORTSTEIL
125     fpSpread1(0).SetText 2, 5, "Tel. " & GmandantRS!Telefon
130     fpSpread1(0).SetText 2, 6, "Fax. " & GmandantRS!Fax
135     fpSpread1(0).SetText 2, 7, GmandantRS!Name1 & ", " & GmandantRS!Name2 & ", " & GmandantRS!Lkz & "-" & GmandantRS!Plz & " " & GmandantRS!Ort & " " & GmandantRS!ORTSTEIL
140     fpSpread1(0).SetText 2, 8, frmParent.txt1(1).text
    
145     If Trim(frmParent.txt1(2).text) <> "" Then
150         fpSpread1(0).SetText 2, 9, frmParent.txt1(2).text
        Else
155         fpSpread1(0).rowHeight(9) = 0
        End If
    
160     If Trim(frmParent.txt1(3).text) <> "" Then
165         fpSpread1(0).SetText 2, 10, frmParent.txt1(3).text
        Else
170         fpSpread1(0).rowHeight(10) = 0
        End If
    
175     If Trim(frmParent.txt1(4).text) <> "" Then
180         fpSpread1(0).SetText 2, 11, "Postfach " & frmParent.txt1(4).text
185         fpSpread1(0).SetText 2, 12, frmParent.txt1(5).text & " " & frmParent.txt1(6).text
        Else
190         fpSpread1(0).SetText 2, 11, frmParent.txt1(7).text
195         fpSpread1(0).SetText 2, 12, frmParent.txt1(9).text & " " & frmParent.txt1(10).text & " " & frmParent.txt1(11).text
        End If
    
200     If Trim(frmParent.txt1(8).text) <> "" Then
205         Set rs = New ADODB.Recordset

210         rs.Open "SELECT Land FROM [1100_Land] WHERE Druck =(-1) AND Lkz = '" & Trim(frmParent.txt1(8).text) & "'", gConn, adOpenStatic, adLockReadOnly

215         If rs.RecordCount > 0 Then
            
220             If fpSpread1(0).GetText(2, 12, knz) Then fpSpread1(0).SetText 2, 12, UCase(knz)
225             fpSpread1(0).SetText 2, 13, UCase(rs!Land)
            End If
        End If
    
230     fpSpread1(0).SetText 7, 13, "UID-Nr.: " & GmandantRS!Uid
235     fpSpread1(0).SetText 7, 14, "USt-Nr.: " & GmandantRS!SteuerNr
240     fpSpread1(0).SetText 5, 15, frmParent.Frame1(2).caption
    
245     fpSpread1(0).Col = 1
250     fpSpread1(0).Row = 16
255     fpSpread1(0).Col2 = fpSpread1(0).MaxCols
260     fpSpread1(0).Row2 = 16
265     fpSpread1(0).BlockMode = True
270     fpSpread1(0).FontBold = True
275     fpSpread1(0).BlockMode = False
280     fpSpread1(0).SetText 2, 16, "Beleg-Nr."
285     fpSpread1(0).SetText 3, 16, "Kunden-Nr."
290     fpSpread1(0).SetText 4, 16, "Beleg-Datum"
295     fpSpread1(0).SetText 5, 16, "Bearbeiter"
300     fpSpread1(0).SetText 3, 17, frmParent.txt1(12).text

305     If Trim(frmParent.BelegNr) = "0" Or Trim(frmParent.BelegNr) = "" Then
310         fpSpread1(0).SetText 2, 17, strAutomatischText
315         fpSpread1(0).Col = 2
320         fpSpread1(0).Row = 17
        Else
325         fpSpread1(0).SetText 2, 17, frmParent.BelegNr
        End If

330     If Trim(frmParent.belegDatum) = "" Then
335         fpSpread1(0).SetText 4, 17, strAutomatischText
340         fpSpread1(0).Col = 4
345         fpSpread1(0).Row = 17
        Else
350         fpSpread1(0).SetText 4, 17, frmParent.belegDatum
        End If

355     fpSpread1(0).SetText 5, 17, GsUser
360     fpSpread1(0).SetText 2, 18, "Bei Zahlungen bitte unbedingt angeben!"

        Exit Sub

Fehler:
365     Call FehlerErklärung("frmSP52831", "KopfFuellen()")

End Sub

Public Sub FussFuellen()

        On Error GoTo Fehler
    
        Dim rsFuä   As New ADODB.Recordset

        Dim strTExt As String

100     fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 1, "Steuer-Pflichtig"

105     fpSpread1(2).SetText COL_FS_UST, 1, "Umsatz-Steuer " & frmParent.txt1(20).text & "%"
        
110     fpSpread1(2).SetText COL_FS_STEUERFREI, 1, "Steuer-Frei*"
115     fpSpread1(2).SetText COL_FS_GESAMT, 1, "Gesamt"
    
120     fpSpread1(2).SetText COL_FS_WRG, 2, frmParent.lbl2(17)
125     fpSpread1(2).SetText COL_FS_WRG, 3, frmParent.lbl2(18)
    
130     fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 4, "UID-Nr.: " & frmParent.txt1(14)
135     fpSpread1(2).SetText COL_FS_UST, 4, "USt-Nr.: " & frmParent.txt1(13)
    
140     Set rsFuä = New ADODB.Recordset
    
145     Select Case GintBelegArt
        
            Case 0, 2, 3
            
150             rsFuä.Open "SELECT * FROM [1100_Texte] WHERE TextArt = 'Rng' AND Lkz = '" & 2 + 2 * intSteuerTyp & "'", gConn, adOpenStatic, adLockReadOnly
            
155         Case 1
            
160             rsFuä.Open "SELECT * FROM [1100_Texte] WHERE TextArt = 'Rng' AND Lkz = '" & 3 + 2 * intSteuerTyp & "'", gConn, adOpenStatic, adLockReadOnly
        
        End Select

165     If rsFuä.RecordCount >= 0 Then
170         strTExt = rsFuä!text
        End If

175     rsFuä.Close

180     fpSpread1(2).SetText COL_FS_GESAMT, 4, strTExt
    
185     If gstrSteuerText <> "" Then
190         fpSpread1(2).SetText COL_FS_STEUERFREI, 4, gstrSteuerText
        End If

195     FussZahlungsZiel

        Exit Sub

Fehler:
200     Call FehlerErklärung("frmSP52831", "FussFuellen()")

End Sub

Public Sub FussFuellenGesamtIstBrutto()

        On Error GoTo Fehler
    
        Dim rsFuä   As New ADODB.Recordset

        Dim strTExt As String

100     fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 1, "Steuer-Pflichtig"

105     fpSpread1(2).SetText COL_FS_UST, 1, "Umsatz-Steuer " & frmParent.txt1(20).text & "%"

110     fpSpread1(2).SetText COL_FS_STEUERFREI, 1, "Steuer-Frei*"
115     fpSpread1(2).SetText COL_FS_GESAMT, 1, "Gesamt"
    
120     fpSpread1(2).SetText COL_FS_WRG, 2, frmParent.lbl2(17)
125     fpSpread1(2).SetText COL_FS_WRG, 3, frmParent.lbl2(18)
    
130     fpSpread1(2).SetText COL_FS_STEUERPFLICHT, 4, "UID-Nr.: " & frmParent.txt1(14)
135     fpSpread1(2).SetText COL_FS_UST, 4, "USt-Nr.: " & frmParent.txt1(13)
    
140     OPEN_gConn
    
145     Select Case GintBelegArt
        
            Case 0, 2, 3
            
150             rsFuä.Open "SELECT * FROM [1100_Texte] WHERE TextArt = 'Rng' AND Lkz = '" & 2 + 2 * intSteuerTyp & "'", gConn, adOpenStatic, adLockReadOnly
            
155         Case 1
            
160             rsFuä.Open "SELECT * FROM [1100_Texte] WHERE TextArt = 'Rng' AND Lkz = '" & 3 + 2 * intSteuerTyp & "'", gConn, adOpenStatic, adLockReadOnly
        
        End Select

165     If rsFuä.RecordCount >= 0 Then
170         strTExt = rsFuä!text
        End If

175     rsFuä.Close
180     Set rsFuä = Nothing
  
185     fpSpread1(2).SetText COL_FS_GESAMT, 4, strTExt
    
190     If gstrSteuerText <> "" Then
195         fpSpread1(2).SetText COL_FS_STEUERFREI, 4, gstrSteuerText
        End If
  
200     FussZahlungsZielGesamtIstBrutto

        Exit Sub

Fehler:
205     Call FehlerErklärung("frmSP52831", "FussFuellenGesamtIstBrutto()")

End Sub

Public Sub PostenSetUp()

    On Error GoTo Fehler

    Dim i              As Integer

    Dim GPreisFormula  As String

    Dim GPreisFormula1 As String
    
    fpSpread1(1).Row = -1
    fpSpread1(1).Col = -1
    fpSpread1(1).Lock = True

    fpSpread1(1).SetActionKey 0, False, False, 0
    fpSpread1(1).SetActionKey 1, False, False, 0
    fpSpread1(1).SetActionKey 2, False, False, 0
    
    fpSpread1(1).EditModePermanent = True

    fpSpread1(1).TypeMaxEditLen = 500
    fpSpread1(1).EditEnterAction = EditEnterActionNext

    fpSpread1(1).MaxCols = COL_LAST
    
    fpSpread1(1).Row = -1
    fpSpread1(1).Col = -1
    
    fpSpread1(1).fontSize = 8
    
    fpSpread1(1).GridShowHoriz = False
    fpSpread1(1).GridShowVert = False
    
    fpSpread1(1).AllowCellOverflow = True
    
    fpSpread1(1).RowHeadersShow = False

    fpSpread1(1).colWidth(COL_ZEILENART) = 600
    fpSpread1(1).colWidth(COL_ARTSCHL) = 1400

    fpSpread1(1).colWidth(COL_ARTIKEL) = TEXT_BREITE
    
    fpSpread1(1).colWidth(COL_MENGE) = 1050
    fpSpread1(1).colWidth(COL_EINHEIT) = 850
    fpSpread1(1).colWidth(COL_EPREIS) = 1050
    fpSpread1(1).colWidth(COL_RABATT) = 850
    fpSpread1(1).colWidth(COL_GPREIS) = 1050

    fpSpread1(1).colWidth(COL_UST) = 650
    fpSpread1(1).colWidth(COL_DURCHLAUFEND) = 800
    fpSpread1(1).colWidth(COL_KOSTSCHL) = 700
    fpSpread1(1).colWidth(COL_SACHSCHL) = 700
    fpSpread1(1).colWidth(COL_KOSTKTO) = 700
    fpSpread1(1).colWidth(COL_SACHKTO) = 700

    fpSpread1(1).colWidth(COL_GPREISDUMMY) = 700
    fpSpread1(1).colWidth(COL_GPREISDUMMY_O_UST) = 700
    fpSpread1(1).colWidth(COL_GPREISDUMMY_M_UST) = 700
    fpSpread1(1).colWidth(COL_SUMMEN) = 700

    fpSpread1(1).SetText COL_ZEILENART, 0, "Art"
    fpSpread1(1).SetText COL_ARTSCHL, 0, "Art." & vbCrLf & "-Schl."
    fpSpread1(1).SetText COL_ARTIKEL, 0, "Artikel"
        
    fpSpread1(1).SetText COL_LS_DATUM, 0, "LS" & vbCrLf & "Datum"
    fpSpread1(1).SetText COL_LS_NUMMER, 0, "LS" & vbCrLf & "Nummer"
    
    fpSpread1(1).SetText COL_MENGE, 0, "Menge"
    fpSpread1(1).SetText COL_EINHEIT, 0, "Einheit"
    fpSpread1(1).SetText COL_EPREIS, 0, "E-Preis"
    fpSpread1(1).SetText COL_RABATT, 0, "Rabatt"
    fpSpread1(1).SetText COL_GPREIS, 0, "G-Preis"

    fpSpread1(1).SetText COL_UST, 0, "USt."
    fpSpread1(1).SetText COL_DURCHLAUFEND, 0, "Durch" & vbCrLf & "-laufend"
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
    
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 0"
    fpSpread1(1).SetText COL_ZEILENART, 0, objPRM.caption("Art")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 1"
    fpSpread1(1).SetText COL_ARTSCHL, 0, objPRM.caption("Schl.")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 2"
    fpSpread1(1).SetText COL_ARTIKEL, 0, objPRM.caption("Artikel")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 3"
    fpSpread1(1).SetText COL_MENGE, 0, objPRM.caption("Menge")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 4"
    fpSpread1(1).SetText COL_EINHEIT, 0, objPRM.caption("Einheit")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 5"
    fpSpread1(1).SetText COL_EPREIS, 0, objPRM.caption("E-Preis")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 6"
    fpSpread1(1).SetText COL_RABATT, 0, objPRM.caption("Rabatt")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 7"
    fpSpread1(1).SetText COL_GPREIS, 0, objPRM.caption("G-Preis")

    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 8"
    fpSpread1(1).SetText COL_UST, 0, objPRM.caption("USt.")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 9"
    fpSpread1(1).SetText COL_DURCHLAUFEND, 0, objPRM.caption("Durch" & vbCrLf & "-laufend")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 10"
    fpSpread1(1).SetText COL_KOSTSCHL, 0, objPRM.caption("Kost." & vbCrLf & "-Schl.")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 11"
    fpSpread1(1).SetText COL_SACHSCHL, 0, objPRM.caption("Sach." & vbCrLf & "-Schl.")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 12"
    fpSpread1(1).SetText COL_KOSTKTO, 0, objPRM.caption("Kost." & vbCrLf & "-Kto.")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 13"
    fpSpread1(1).SetText COL_SACHKTO, 0, objPRM.caption("Sach." & vbCrLf & "-Kto.")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 14"
    fpSpread1(1).SetText COL_LS_DATUM, 0, objPRM.caption("Liefer" & vbCrLf & "-Datum")
    objPRM.FindFirstString = "name = 'ArtikelKopfzeile' AND index = 15"
    fpSpread1(1).SetText COL_LS_NUMMER, 0, objPRM.caption("Liefer" & vbCrLf & "-Referenz")

    fpSpread1(1).rowHeight(0) = 500
    fpSpread1(1).Col = 1
    fpSpread1(1).Row = 0
    fpSpread1(1).Col2 = fpSpread1(1).MaxCols
    fpSpread1(1).Row2 = 0
    fpSpread1(1).BlockMode = True
    fpSpread1(1).FontBold = True
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).ButtonDrawMode = 1

    For i = 1 To fpSpread1(1).MaxCols

        Select Case i

            Case COL_ZEILENART, COL_KOSTSCHL, COL_SACHSCHL, COL_EINHEIT
                
                fpSpread1(1).Col = i
                fpSpread1(1).Row = 1
                fpSpread1(1).Col2 = i
                fpSpread1(1).Row2 = fpSpread1(1).MaxRows
                fpSpread1(1).BlockMode = True
                fpSpread1(1).Lock = False
                    
                fpSpread1(1).CellType = CellTypeComboBox
                
                fpSpread1(1).TypeComboBoxEditable = True
                    
                fpSpread1(1).BlockMode = False

            Case COL_MENGE, COL_EPREIS, COL_RABATT, COL_GPREIS, COL_GPREISDUMMY, COL_GPREISDUMMY_O_UST, COL_GPREISDUMMY_M_UST, COL_SUMMEN

                fpSpread1(1).Col = i
                fpSpread1(1).Row = 0
                fpSpread1(1).Col2 = i
                fpSpread1(1).Row2 = fpSpread1(1).MaxRows
                fpSpread1(1).BlockMode = True
                fpSpread1(1).TypeHAlign = 1
                fpSpread1(1).BlockMode = False

                fpSpread1(1).Col = i
                fpSpread1(1).Row = 1
                fpSpread1(1).Col2 = i
                fpSpread1(1).Row2 = fpSpread1(1).MaxRows
                fpSpread1(1).BlockMode = True
                fpSpread1(1).CellType = CellTypeNumber
                fpSpread1(1).TypeNumberShowSep = True
                fpSpread1(1).TypeNumberDecPlaces = 2

                Select Case i

                    Case COL_GPREISDUMMY
                        
                        GPreisFormula = "ROUND(IF(" & COL_EINHEIT_STR & "#=""%"",(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#/100)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#/100)*(" & COL_RABATT_STR & "#/100)),(" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)-((" & COL_MENGE_STR & "#*" & COL_EPREIS_STR & "#)*(" & COL_RABATT_STR & "#/100))) * 1000000, 0) / 1000000"

                        fpSpread1(1).Formula = GPreisFormula

                End Select

                fpSpread1(1).BlockMode = False

            Case COL_UST, COL_DURCHLAUFEND
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
    
    fpSpread1(1).Col = COL_RABATT
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_RABATT
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).TypeNumberMax = 100#
    fpSpread1(1).TypeNumberMin = 0#
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).Col = COL_MENGE
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_MENGE
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    
    fpSpread1(1).BlockMode = False
    
    GPreisFormula1 = "ROUND(" & Chr(64 + COL_GPREISDUMMY) & "#*100, 0)/100"
    
    fpSpread1(1).Col = COL_GPREIS
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_GPREIS
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).Formula = GPreisFormula1
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).Col = COL_GPREISDUMMY_O_UST
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_GPREISDUMMY_O_UST
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).Formula = "IF(" & COL_UST_STR & "#=""0""," & GPreisFormula1 & ",""0"")"
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).Col = COL_GPREISDUMMY_M_UST
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_GPREISDUMMY_M_UST
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).Formula = "IF(" & COL_UST_STR & "#=""1""," & GPreisFormula1 & ",""0"")"
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).Col = COL_SUMMEN
    fpSpread1(1).Row = 1
    fpSpread1(1).Formula = "SUM(" & COL_GPREISDUMMY_M_UST_STR & "1:" & COL_GPREISDUMMY_M_UST_STR & CStr(fpSpread1(1).MaxRows) & ")"
    
    fpSpread1(1).Row = 2
    fpSpread1(1).Formula = "SUM(" & COL_GPREISDUMMY_O_UST_STR & "1:" & COL_GPREISDUMMY_O_UST_STR & CStr(fpSpread1(1).MaxRows) & ")"
    
    fpSpread1(1).Col = COL_GPREISDUMMY
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = fpSpread1(1).MaxCols
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).ColHidden = True
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).Col = COL_ZEILENART
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_ARTSCHL
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).BackColor = RGB(235, 229, 217)
    fpSpread1(1).BlockMode = False
   
    fpSpread1(1).Col = COL_DURCHLAUFEND
    fpSpread1(1).Row = 1
    fpSpread1(1).Col2 = COL_SACHKTO
    fpSpread1(1).Row2 = fpSpread1(1).MaxRows
    fpSpread1(1).BlockMode = True
    fpSpread1(1).BackColor = RGB(235, 229, 217)
    fpSpread1(1).BlockMode = False
    
    fpSpread1(1).SetActiveCell 1, 1

    Exit Sub

Fehler:
    Call FehlerErklärung("frmSP52831", "PostenSetUp()")

End Sub

Public Sub ZeilenTyp(ByVal Row As Long)

        On Error GoTo Fehler

        Dim knz           As Variant

        Dim i             As Long

        Dim SeitenUmbruch As String
  
100     Schalter True
  
105     fpSpread1(1).Row = Row
  
110     If fpSpread1(1).GetText(COL_ZEILENART, Row, knz) = False Then
115         knz = ""
        End If

120     fpSpread1(1).ClearRange 2, Row, fpSpread1(1).MaxCols, Row, True
125     fpSpread1(1).Col = COL_UST
130     fpSpread1(1).CellType = CellTypeStaticText
135     fpSpread1(1).Col = COL_ARTSCHL
140     fpSpread1(1).CellType = CellTypeStaticText
145     fpSpread1(1).Col = COL_DURCHLAUFEND
150     fpSpread1(1).CellType = CellTypeStaticText

155     Select Case knz

            Case ""

160             For i = 2 To COL_LASTEDIT

165                 fpSpread1(1).Col = i
170                 fpSpread1(1).Lock = True

175             Next i

180         Case "A", "P"

185             For i = 2 To COL_LASTEDIT

190                 fpSpread1(1).Col = i

195                 Select Case i

                        Case COL_ARTSCHL
                        
200                         KontrollierenZellgruppierung Row, COL_ARTSCHL, False

205                         If knz = "A" Then

210                             fpSpread1(1).CellType = CellTypeComboBox
215                             fpSpread1(1).Lock = False
                            
220                             fpSpread1(1).Col = COL_EPREIS
225                             fpSpread1(1).TypeNumberDecPlaces = postCommaPreis

                            End If

230                     Case COL_ARTIKEL

235                         fpSpread1(1).TypeMaxEditLen = 500
240                         fpSpread1(1).Lock = False

245                     Case COL_UST

250                         fpSpread1(1).CellType = CellTypeCheckBox
                        
255                         Select Case intSteuerTyp

                                Case 0
260                                 fpSpread1(1).SetText COL_UST, Row, 0
265                                 fpSpread1(1).Lock = False

270                             Case 1
275                                 fpSpread1(1).SetText COL_UST, Row, 1
280                                 fpSpread1(1).Lock = False

285                             Case 2
290                                 fpSpread1(1).SetText COL_UST, Row, 0
295                                 fpSpread1(1).Lock = True

                            End Select

300                     Case COL_DURCHLAUFEND

305                         fpSpread1(1).CellType = CellTypeCheckBox
310                         fpSpread1(1).Lock = False

315                     Case COL_GPREIS, COL_LS_DATUM, COL_LS_NUMMER

320                         fpSpread1(1).Lock = True
                            
325                     Case Else

330                         fpSpread1(1).Lock = False

                    End Select

335             Next i

340             fpSpread1(1).SetText COL_MENGE, Row, 1

345             If knz = "A" Then

350                 fpSpread1(1).SetText COL_EPREIS, Row, 0
355                 fpSpread1(1).SetText COL_ARTIKEL, Row, ""

                Else
                
360                 fpSpread1(1).SetText COL_EPREIS, Row, frmParent.txt1(24)
365                 fpSpread1(1).SetText COL_ARTIKEL, Row, frmParent.lbl1(24)
370                 fpSpread1(1).SetText COL_EINHEIT, Row, C_STR_COL_EINHEIT_STUECK

                End If

375             fpSpread1(1).SetText COL_MENGE, Row, 1
                
380         Case "L"

385             For i = 2 To COL_LASTEDIT

390                 fpSpread1(1).Col = i

395                 Select Case i

                        Case COL_ARTSCHL
                        
400                         KontrollierenZellgruppierung Row, COL_ARTSCHL, False
                        
405                         fpSpread1(1).CellType = CellTypeComboBox
410                         fpSpread1(1).Lock = False
415                         fpSpread1(1).Col = COL_EPREIS
420                         fpSpread1(1).TypeNumberDecPlaces = postCommaPreis

425                     Case COL_ARTIKEL

430                         fpSpread1(1).TypeMaxEditLen = 500
435                         fpSpread1(1).Lock = False

440                     Case COL_LS_DATUM

445                         fpSpread1(1).TypeMaxEditLen = 10
450                         fpSpread1(1).Lock = False

455                     Case COL_LS_NUMMER

460                         fpSpread1(1).TypeMaxEditLen = 30
465                         fpSpread1(1).Lock = False

470                     Case COL_UST

475                         fpSpread1(1).CellType = CellTypeCheckBox
                        
480                         Select Case intSteuerTyp

                                Case 0
                                
485                                 fpSpread1(1).SetText COL_UST, Row, 0
490                                 fpSpread1(1).Lock = False

495                             Case 1

500                                 fpSpread1(1).SetText COL_UST, Row, 1
505                                 fpSpread1(1).Lock = False

510                             Case 2

515                                 fpSpread1(1).SetText COL_UST, Row, 0
520                                 fpSpread1(1).Lock = True

                            End Select

525                     Case COL_DURCHLAUFEND

530                         fpSpread1(1).CellType = CellTypeCheckBox
535                         fpSpread1(1).Lock = False

540                     Case COL_GPREIS
                            
545                     Case Else

550                         fpSpread1(1).Lock = False

                    End Select

555             Next i

560             If fpSpread1(1).colWidth(COL_LS_NUMMER) = 0 Then

565                 fpSpread1(1).colWidth(COL_ARTIKEL) = TEXT_BREITE * 0.6
570                 fpSpread1(1).colWidth(COL_LS_DATUM) = TEXT_BREITE * 0.2
575                 fpSpread1(1).colWidth(COL_LS_NUMMER) = TEXT_BREITE * 0.2

                End If
            
580             If gEnmKudnenERechnungType = eERechnungType.ZUGFeRD And IsEBelegDoc And frmParent.blnBelegNeu And gbZeileInCopy = False Then
                    
                    Dim lngFindRow     As Long

                    Dim blnShowMessage As Boolean
                    
585                 lngFindRow = fpSpread1(1).SearchCol(COL_ZEILENART, 0, -1, "L", SearchFlagsNone)
                    
590                 If lngFindRow >= 0 Then
                       
595                     If dictLSRows.Count = 0 Then
                           
600                         dictLSRows.Add Row, Row
                            
                        Else
                        
605                         If dictLSRows.Exists(Row) = False Then
                            
610                             dictLSRows.Add Row, Row
615                             blnShowMessage = True
                                
                            End If
                           
                        End If
                        
                    End If
                    
620                 If blnShowMessage Then
                          
625                     MsgBox GetMessage(2391), vbOKOnly + vbExclamation, strMeldungCap

                    End If
                
                End If
                
630         Case "T"

635             KontrollierenZellgruppierung Row, COL_ARTSCHL, True

640             For i = 2 To COL_LASTEDIT
645                 fpSpread1(1).Col = i
                
650                 If i < COL_LS_DATUM Then
655                     fpSpread1(1).Lock = False
                    Else
660                     fpSpread1(1).Lock = True
                    End If

665             Next i

670             fpSpread1(1).Col = COL_ARTSCHL
675             fpSpread1(1).CellType = CellTypeComboBox
            
680             fpSpread1(1).Col = COL_ARTIKEL
685             fpSpread1(1).TypeMaxEditLen = 500

690         Case "Z"

695             For i = 2 To COL_LASTEDIT
700                 fpSpread1(1).Col = i
                
705                 If i < COL_LS_DATUM Then
710                     fpSpread1(1).Lock = False
                    Else
715                     fpSpread1(1).Lock = True
                    End If

720             Next i

725             fpSpread1(1).Col = COL_ARTSCHL
730             fpSpread1(1).Lock = True
735             fpSpread1(1).SetText COL_ARTIKEL, Row, "Zwischensumme"
740             fpSpread1(1).SetText COL_GPREIS, Row, ZwischenSumme(Row)

745         Case "S"

750             For i = 2 To COL_LASTEDIT
755                 fpSpread1(1).Col = i
760                 fpSpread1(1).Lock = True
765             Next i
    
770             fpSpread1(1).Col = COL_ARTIKEL
775             fpSpread1(1).TypeMaxEditLen = 500
780             SeitenUmbruch = "- - - - - - - - - - - - - - - - - - - - "
785             SeitenUmbruch = SeitenUmbruch + SeitenUmbruch + SeitenUmbruch + SeitenUmbruch + SeitenUmbruch
790             fpSpread1(1).SetText COL_ARTIKEL, Row, SeitenUmbruch

        End Select
    
795     fpSpread1(1).SetText COL_ERSTDAT, Row, CStr(Now)
800     fpSpread1(1).SetText COL_ERSTVON, Row, GsUser
805     fpSpread1(1).SetText COL_AENDDAT, Row, CStr(Now)
810     fpSpread1(1).SetText COL_AENDVON, Row, GsUser
    
815     If fpSpread1(1).colWidth(COL_LS_NUMMER) > 0 Then LSArtikelGefunden

        Exit Sub

Fehler:
820     Call FehlerErklärung("frmSP52831", "ZeilenTyp()")

End Sub

Public Function IstZeilenTyp(ByVal Row As Long, ZeilenTyp As String) As Boolean
    
        On Error GoTo Fehler
    
        Dim AltCol As Long
  
100     AltCol = fpSpread1(1).Col
  
105     fpSpread1(1).Row = Row
110     fpSpread1(1).Col = COL_UST
  
115     Select Case ZeilenTyp

            Case "A", "P", "L"

120             If fpSpread1(1).CellType = CellTypeCheckBox Then
125                 IstZeilenTyp = True
                End If

130         Case "T"

135             If fpSpread1(1).CellType <> CellTypeCheckBox Then
140                 fpSpread1(1).Col = COL_ARTSCHL

145                 If fpSpread1(1).CellType = CellTypeComboBox Then
150                     IstZeilenTyp = True
                    End If
                End If

155         Case "Z"

160             If fpSpread1(1).CellType <> CellTypeCheckBox Then
165                 fpSpread1(1).Col = COL_ARTSCHL

170                 If fpSpread1(1).CellType <> CellTypeComboBox Then
175                     IstZeilenTyp = True
                    End If
                End If

180         Case "S"

185             If fpSpread1(1).CellType <> CellTypeCheckBox Then
190                 fpSpread1(1).Col = COL_ARTSCHL

195                 If fpSpread1(1).CellType <> CellTypeComboBox Then
200                     IstZeilenTyp = True
                    End If
                End If

        End Select
  
205     fpSpread1(1).Col = AltCol

        Exit Function

Fehler:
210     Call FehlerErklärung("frmSP52831", "IstZeilenTyp")

End Function

Public Function ZwischenSumme(ByVal Row As Long) As Double
    
        On Error GoTo Fehler
    
        Dim i   As Integer

        Dim knz As Variant
        
100     fpSpread1(1).ReDraw = False
105     fpSpread1(1).Row = Row
  
110     For i = fpSpread1(1).Row - 1 To 1 Step -1

115         If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then

120             If knz = "A" Or knz = "P" Or knz = "L" Then

125                 If fpSpread1(1).GetText(COL_GPREIS, i, knz) Then

130                     If IsNumeric(knz) Then

135                         ZwischenSumme = ZwischenSumme + Runden(CDbl(knz), 2)
                            
                        End If
                        
                    End If
                    
                End If
                
            End If

140         If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then

145             If knz = "Z" Then Exit For

            End If

150     Next i

155     fpSpread1(1).ReDraw = True

        Exit Function

Fehler:
160     Call FehlerErklärung("frmSP52831", "ZwischenSumme")

End Function

Public Sub ZwischenSummeRefresh(ByVal Row As Long, Optional Alle As Boolean)
    
        On Error GoTo Fehler
    
        Dim i   As Integer

        Dim knz As Variant
        
100     fpSpread1(1).ReDraw = False
        
105     For i = Row To LastRow

110         If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then

115             If knz = "Z" Then

120                 fpSpread1(1).SetText COL_GPREIS, i, ZwischenSumme(i)

125                 If Alle = False Then Exit For

                End If
                
            End If

130     Next i

135     fpSpread1(1).ReDraw = True

        Exit Sub

Fehler:
140     Call FehlerErklärung("frmSP52831", "ZwischenSummeRefresh")

End Sub

Public Function LSArtikelGefunden() As Boolean

        On Error GoTo Fehler

        Dim i   As Integer
                
        Dim knz As Variant
  
100     For i = 1 To LastRow

105         If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then

110             If knz = "L" Then

115                 LSArtikelGefunden = True

                    Exit For

                End If
                
            End If

120     Next i
  
125     If LSArtikelGefunden = False Then

130         fpSpread1(1).colWidth(COL_ARTIKEL) = TEXT_BREITE
135         fpSpread1(1).colWidth(COL_LS_DATUM) = 0
140         fpSpread1(1).colWidth(COL_LS_NUMMER) = 0

        End If

        Exit Function

Fehler:
145     Call FehlerErklärung("frmSP52831", "LSArtikelGefunden()")

End Function

Public Function IstPorto() As Boolean
    
        On Error GoTo Fehler
    
        Dim i   As Integer

        Dim knz As Variant
  
100     If IsNumeric(frmParent.txt1(24)) Then
105         If frmParent.txt1(24) > 0 Then

110             For i = 1 To LastRow

115                 If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then
120                     If knz = "P" Then
125                         IstPorto = True

                            Exit For

                        End If
                    End If

130             Next i

            Else
135             IstPorto = True
            End If

        Else
140         IstPorto = True
        End If

        Exit Function

Fehler:
145     Call FehlerErklärung("frmSP52831", "IstPorto")

End Function

Public Function IstArtieklUndMengeOK() As Boolean
    
        On Error GoTo Fehler
        
        Dim i   As Integer

        Dim knz As Variant
    
100     If IstArtieklVorhanden Then
  
105         For i = 1 To LastRow
            
110             If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then

115                 If knz = "A" Or knz = "P" Then

120                     If fpSpread1(1).GetText(COL_ARTIKEL, i, knz) Then

125                         If Trim(knz) = "" Then

130                             fpSpread1(1).ClearSelection

135                             Call SetActiveCellExt(1, COL_ARTIKEL, CLng(i), True)
                              
140                             MsgBox GetMessage(2194), vbExclamation, strMeldungCap

                                Exit Function

                            End If

                        Else
                        
145                         fpSpread1(1).ClearSelection

150                         Call SetActiveCellExt(1, COL_ARTIKEL, CLng(i), True)
                           
155                         MsgBox GetMessage(2194), vbExclamation, strMeldungCap

                            Exit Function

                        End If
                    
                    End If
                End If

160         Next i

        Else
        
165         MsgBox GetMessage(2193), vbExclamation, strMeldungCap
170         IstArtieklUndMengeOK = False

            Exit Function

        End If
        
175     IstArtieklUndMengeOK = True

        Exit Function

Fehler:
180     Call FehlerErklärung("frmSP52831", "IstArtieklUndMengeOK")

End Function

Public Function EinheitsPrufungMitFokus() As Boolean
        
        Dim BelegID As Long
        
        Dim connSQL As ADODB.Connection
    
100     BelegID = GetMAXID_FROM_TABLE("2800_HauptTmp", "BelegID")
105     BelegID = BelegID + 1
        
110     Call FolgeSpeichern(BelegID, True)
        
115     If Not CheckENCodeBeiVerpackung(1, CStr(BelegID), True, "2800_FolgeTmp") Then
                                
120         If lngFirstFalseRow <> 0 Then

125             fpSpread1(1).ClearSelection

130             Call SetActiveCellExt(1, COL_EINHEIT, lngFirstFalseRow, True)

135             fpSpread1(1).EditMode = True

140             lngFirstFalseRow = 0

145             EinheitsPrufungMitFokus = False

            End If

        Else

150         EinheitsPrufungMitFokus = True

        End If
        
155     Set connSQL = New ADODB.Connection

160     connSQL.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
165     connSQL.Open
170     connSQL.Execute "DELETE FROM [2800_FolgeTmp] where BelegId = " & CStr(BelegID)

        Exit Function

Fehler:
175     Call FehlerErklärung("frmSP52831", "EinheitsPrufungMitFokus")

End Function

Private Sub Form_Unload(Cancel As Integer)

        On Error GoTo Fehler
    
100     objDruckOptionen.CurrentValutaDatum = ""
105     objDruckOptionen.CurrentBelegDatum = ""
110     objDruckOptionen.CurrentBelegNr = ""
    
115     Set cReSize = Nothing

120     frmParent.gboolHatKind = False

        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52831", "Form_Unload()")

End Sub

Private Sub fpSpread1_Change(Index As Integer, ByVal Col As Long, ByVal Row As Long)

        On Error GoTo Fehler

        Dim knz       As Variant

        Dim Knz2      As Variant

        Dim KnzGesamt As Variant

        Dim ret       As Double
         
100     Select Case Index

            Case 1
            
105             blnChangeEventFired = True
                
110             Schalter True

115             If (Row = 1 Or Row = 2) And Col = COL_SUMMEN Then
                    
120                 If GesamtIstBrutto Then

125                     Select Case Row

                            Case 1

130                             If fpSpread1(1).GetText(COL_SUMMEN, Row, knz) Then
                                Else
135                                 knz = 0
                                End If

140                             If fpSpread1(1).GetText(COL_SUMMEN, 2, Knz2) Then
                                Else
145                                 Knz2 = 0
                                End If
  
150                             KnzGesamt = knz + Knz2

155                             fpSpread1(2).SetText COL_FS_GESAMT, 2, KnzGesamt

160                         Case 2

165                             If fpSpread1(1).GetText(COL_SUMMEN, Row, Knz2) Then
170                                 fpSpread1(2).SetText 4, 2, Knz2
                                Else
175                                 fpSpread1(2).SetText 4, 2, 0
180                                 Knz2 = 0
                                End If

185                             If fpSpread1(1).GetText(COL_SUMMEN, 1, knz) Then
                                Else
190                                 knz = 0
                                End If
  
195                             KnzGesamt = knz + Knz2

200                             fpSpread1(2).SetText COL_FS_GESAMT, 2, KnzGesamt

                        End Select

                    Else

205                     Select Case Row

                            Case 1

210                             If fpSpread1(1).GetText(COL_SUMMEN, Row, knz) Then
215                                 fpSpread1(2).SetText 2, 2, knz
                                Else
220                                 fpSpread1(2).SetText 2, 2, "0"
                                End If

225                         Case 2

230                             If fpSpread1(1).GetText(COL_SUMMEN, Row, knz) Then
235                                 fpSpread1(2).SetText 4, 2, knz
                                Else
240                                 fpSpread1(2).SetText 4, 2, "0"
                                End If

                        End Select

                    End If
                    
                End If

245             If Row = fpSpread1(1).ActiveRow Then

250                 Select Case Col

                        Case COL_MENGE
                                                             
255                         If dictFormelRows.Exists(Row) Then
                                                             
                                Dim LR                   As Long
                                
                                Dim lngLastFormelTextRow As Long
                                
                                Dim varFormelTextKnz     As Variant
                                
260                             lngLastFormelTextRow = DEFAULT_VALUE(dictFormelRows.Item(Row), 0, True)

265                             LR = lngLastFormelTextRow

270                             fpSpread1(1).GetText COL_ARTSCHL, Row + 1, varFormelTextKnz
                            
275                             If LR > 0 Then ZeileBearbeiten 4, Row + 1, LR

280                             fpSpread1(1).SetText COL_ZEILENART, Row + 1, "T"

285                             ZeilenTyp Row + 1

290                             fpSpread1(1).SetText COL_ARTSCHL, Row + 1, varFormelTextKnz
                            
295                             objSQLAusw.GetIfOnesHit = True
                            
300                             Auswahl COL_ARTSCHL, Row + 1
    
305                             objSQLAusw.GetIfOnesHit = False
                            
                            End If

310                     Case COL_EPREIS
                        
315                         fpSpread1(1).GetFloat COL_EPREIS, Row, ret
320                         fpSpread1(1).Row = Row
325                         fpSpread1(1).Col = COL_EPREIS

330                         If ret = 0 Then

335                             fpSpread1(1).ForeColor = vbRed

                            Else
                            
340                             fpSpread1(1).ForeColor = vbBlack

                            End If

345                     Case COL_GPREIS
                            
350                         fpSpread1(1).EventEnabled(6) = False
                            
355                         ZwischenSummeRefresh Row

360                         fpSpread1(1).EventEnabled(6) = True

365                     Case COL_ZEILENART

370                         ZeilenTyp Row

375                     Case COL_EINHEIT

380                         If fpSpread1(1).GetText(COL_EINHEIT, Row, knz) Then

385                             If Trim(knz) = "%" Then

390                                 ret = LetzterBetrag(Row)

395                                 If ret > 0 Then
                                        
400                                     fpSpread1(1).SetText COL_EPREIS, Row, ret

                                    End If
                                    
                                End If
                                
                            End If

                    End Select

                End If
                
405             If Col = COL_EINHEIT Then

410                 modERechnung.boolVPSchlusselanderung = True

                End If

415         Case 2

420             Select Case Col

                    Case COL_FS_UST

425                     Select Case Row

                            Case 2
                                    
430                             fpSpread1(2).EventEnabled(6) = False
                                    
435                             If frmParent.txt1(20).text <> 0 Then

440                                 fpSpread1(2).SetText COL_FS_UST, 1, "Umsatz-Steuer " & frmParent.txt1(20).text & "%"

                                Else

445                                 fpSpread1(2).GetText 3, 2, vValue

450                                 If vValue <> 0 Then

455                                     fpSpread1(2).SetText COL_FS_UST, 1, "Umsatz-Steuer " & Format(dblUstSatz, "0.00") & "%"

                                    Else

460                                     fpSpread1(2).SetText COL_FS_UST, 1, "Umsatz-Steuer 0,00 %"

                                    End If

                                End If
                                    
465                             fpSpread1(2).EventEnabled(6) = True
                                
                        End Select

470                 Case COL_FS_WRG

475                     fpSpread1(2).SetText COL_FS_WRG, 2, frmParent.lbl2(17)

480                     fpSpread1(2).SetText COL_FS_WRG, 3, frmParent.lbl2(18)
    
                End Select

        End Select

        Exit Sub

Fehler:
485     Call FehlerErklärung("frmSP52831", "fpSpread1_Change()")

End Sub

Private Sub fpSpread1_ComboDropDown(Index As Integer, _
                                    ByVal Col As Long, _
                                    ByVal Row As Long)

        On Error GoTo Fehler
  
100     Auswahl Col, Row

        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52831", "fpSpread1_ComboDropDown()")

End Sub

Private Sub fpSpread1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        On Error GoTo Fehler
  
        Dim NeuRow As Long

        Dim knz    As Variant
    
100     If Shift > 0 Then

105         ShiftGedrueckt = True

        End If
  
110     If Index = 1 Then

115         Select Case KeyCode

                Case vbKeyF1

120                 If Shift = 1 Then
                    
125                     objHlp.HlpShow HlpWrite, "fpSpread1 - 000" & Index & fpSpread1(Index).ActiveCol
                    Else
130                     objHlp.HlpShow HlpRead, "fpSpread1 - 000" & Index & fpSpread1(Index).ActiveCol
                    End If

135             Case vbKeyF2

140                 Select Case fpSpread1(1).ActiveCol
                    
                        Case COL_ZEILENART, COL_ARTSCHL, COL_KOSTSCHL, COL_SACHSCHL, COL_EINHEIT
                    
145                         Auswahl fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow
                    
                    End Select

150             Case vbKeyReturn

155                 If Shift = 1 Then

160                     If fpSpread1(1).ActiveCol = COL_EPREIS Then

165                         If PlausiEPreis = False Then

                                Exit Sub

                            End If
                            
                        End If
                    
170                     If fpSpread1(1).ActiveRow = fpSpread1(1).MaxRows Then

175                         NeuRow = fpSpread1(1).ActiveRow

                        Else
                        
180                         NeuRow = fpSpread1(1).ActiveRow + 1

                        End If
                    
185                     Call SetActiveCellExt(1, COL_ZEILENART, NeuRow, True)

                    Else

190                     Select Case fpSpread1(1).ActiveCol

                            Case COL_ARTSCHL

195                             If fpSpread1(1).GetText(COL_ARTSCHL, fpSpread1(1).ActiveRow, knz) Then

200                                 If Trim(knz) <> "" Then

205                                     If gvrnMerker <> knz Then
                                        
210                                         objSQLAusw.GetIfOnesHit = True

215                                         boolEnterGedruckt = True

220                                         Auswahl fpSpread1(1).ActiveCol, fpSpread1(1).ActiveRow

225                                         objSQLAusw.GetIfOnesHit = False

                                        End If
                                        
                                    End If
                                    
                                End If

230                         Case COL_LS_DATUM

235                             If fpSpread1(1).GetText(COL_LS_DATUM, fpSpread1(1).ActiveRow, knz) Then

240                                 If Trim(knz) <> "" Then

245                                     fpSpread1(1).SetText COL_LS_DATUM, fpSpread1(1).ActiveRow, ZahlToDatum(knz)

                                    End If
                                    
                                End If
                                
250                         Case COL_EINHEIT
            
255                             If fpSpread1(1).GetText(COL_EINHEIT, fpSpread1(1).ActiveRow, knz) Then
            
260                                 If Not CheckENCodeBeiVerpackung(2, CStr(knz), False) Then

265                                     Call SetCellBackColor(1, COL_EINHEIT, fpSpread1(1).ActiveRow, True)

270                                     Call Auswahl(COL_EINHEIT, fpSpread1(1).ActiveRow)

275                                     Call SetActiveCellExt(1, fpSpread1(1).ActiveCol - 1, fpSpread1(1).ActiveRow, True)
            
                                    End If
            
                                End If
           
                        End Select

                    End If

280             Case vbKeyEscape
                    
285                 If Shift = 1 Then
                    
290                     Call SetActiveCellExt(1, COL_ZEILENART, fpSpread1(1).ActiveRow, True, True)
                        
                    Else

295                     If fpSpread1(1).ActiveCol <> COL_ZEILENART Then
                            
300                         Select Case fpSpread1(1).ActiveCol
                                
                                Case COL_MENGE
                                    
305                                 Call fpSpread1(1).GetText(1, fpSpread1(1).ActiveRow, knz)
                                    
310                                 If (fpSpread1(1).colWidth(COL_LS_DATUM) = 0 And fpSpread1(1).colWidth(COL_LS_NUMMER) = 0) Or knz <> "L" Then

315                                     Call SetActiveCellExt(1, fpSpread1(1).ActiveCol - 3, fpSpread1(1).ActiveRow, True, True)

                                    Else
                                    
320                                     Call SetActiveCellExt(1, fpSpread1(1).ActiveCol - 1, fpSpread1(1).ActiveRow, True, True)

                                    End If
                                    
325                             Case COL_UST
                                    
330                                 Call SetActiveCellExt(1, fpSpread1(1).ActiveCol - 2, fpSpread1(1).ActiveRow, True, True)
                                    
335                             Case Else
                                
340                                 Call SetActiveCellExt(1, fpSpread1(1).ActiveCol - 1, fpSpread1(1).ActiveRow, True, True)
                                
                            End Select

                        Else
                            
345                         If fpSpread1(1).ActiveRow <> 1 Then
                            
350                             Call SetActiveCellExt(1, COL_ZEILENART, fpSpread1(1).ActiveRow - 1, True, True)
                                
                            End If
                            
                        End If
                        
                    End If
                    
355             Case vbKeyDelete

360                 Select Case fpSpread1(Index).ActiveCol

                        Case COL_KOSTSCHL, COL_SACHSCHL
                        
365                         fpSpread1(Index).SetText fpSpread1(Index).ActiveCol, fpSpread1(Index).ActiveRow, ""

                    End Select
                    
370             Case vbKeyF5
                    
375                 Call cmd1_Click(3)
                    
380             Case vbKeyF6
                    
385                 Call cmd1_Click(2)
                    
            End Select
    
        End If

        Exit Sub

Fehler:
390     Call FehlerErklärung("frmSP52831", "fpSpread1_KeyDown()")

End Sub

Public Sub Auswahl(ByVal Col As Long, ByVal Row As Long)

        On Error GoTo Fehler

        Dim i          As Long

        Dim ColLeft    As Long

        Dim RowBottom  As Long

        Dim Merk       As String

        Dim rc         As rect

        Dim knz        As Variant

        Dim Knz1       As Variant

        Dim TextRows   As Long

        Dim ret        As Double
        
        Dim vMenge     As Variant
        
        Dim vEinheit   As Variant
        
        Dim LR         As Long

        Dim Vorher     As String

        Dim ActiveCell As Long
        
100     Call GetWindowRect(fpSpread1(1).hwnd, rc)
105     ColLeft = rc.left * Screen.TwipsPerPixelX
110     RowBottom = rc.top * Screen.TwipsPerPixelY

115     Call SetCellBackColor(1, Col, Row, True)
  
120     For i = fpSpread1(1).LeftCol To Col - 1
125         ColLeft = ColLeft + fpSpread1(1).colWidth(i) + 15
130     Next i

135     RowBottom = RowBottom + fpSpread1(1).rowHeight(0) + ((Row - fpSpread1(1).TopRow + 1) * (fpSpread1(1).rowHeight(Row) + 15))
  
140     objSQLAusw.top = RowBottom
145     objSQLAusw.left = ColLeft
150     objSQLAuswDef.left = ColLeft
155     objSQLAuswDef.top = RowBottom
  
160     If fpSpread1(1).GetText(Col, 0, knz) Then

165         objSQLAusw.caption = StrClean(knz, 10, 13)
170         objSQLAuswDef.caption = StrClean(knz, 10, 13)

        End If
  
175     If fpSpread1(1).GetText(Col, Row, knz) And Trim(knz) <> "" Then Vorher = knz
  
180     Select Case Col

            Case COL_ZEILENART

185             objSQLAuswDef.FilterBar = True
            
190             If fpSpread1(1).GetText(Col, Row, knz) And Trim(knz) <> "" Then

195                 objSQLAuswDef.Find = "Knz like '" & knz & "*'"

200                 Merk = knz

                End If

205             objPRM.FindFirstString = "name = 'ArtF2' AND Index = 1"
210             objSQLAuswDef.ColParameter 1, ColCaption, objPRM.caption("KnzBez1")
215             objPRM.FindFirstString = ""

220             objSQLAuswDef.RSOpen GetACCESSConnectionString(DEF_CONNECTION), "SELECT Knz, KnzBez1 FROM [Auswahl] WHERE TabName = '2800_Folge' AND FeldName = 'SatzTyp' ORDER BY Knz"

225             If objSQLAuswDef.Abbruch = False Then

230                 fpSpread1(1).SetText COL_ZEILENART, Row, objSQLAuswDef.FieldText(0)

235                 If Merk <> objSQLAuswDef.FieldText(0) Then

240                     ZeilenTyp Row

245                     Call SetActiveCellExt(1, Col + 1, Row, True)

                    Else

250                     If IstZeilenTyp(Row, Merk) = False Then
                        
255                         ZeilenTyp Row

260                         Call SetActiveCellExt(1, Col + 1, Row, True)

                        End If

                    End If

                Else

265                 If IstZeilenTyp(Row, Merk) = False Then
                    
270                     ZeilenTyp Row
                    End If

275                 Call SetActiveCellExt(1, Col, Row, False)

                End If

280             fpSpread1(1).Col = COL_EPREIS
285             fpSpread1(1).Row = Row
290             fpSpread1(1).TypeNumberDecPlaces = postCommaPreis

295         Case COL_ARTSCHL

300             Call fpSpread1(1).GetText(COL_ZEILENART, Row, knz)
            
305             If fpSpread1(1).GetText(Col, Row, Knz1) And Trim(Knz1) <> "" Then

310                 TextDummy.text = Knz1

                End If

                Dim oF2MCodeSchl As ResultF2_TextSchl

315             oF2MCodeSchl = GetF2_TextSchl(CStr(knz), E_DATATYPE.Sonderfaktura_Rechnung, "Schl", TextDummy, 0, Me, cReSize, objPRM, objSQLAusw.GetIfOnesHit, ColLeft, RowBottom)

320             TextDummy.text = ""

325             If oF2MCodeSchl.Canceled = False Then

330                 fpSpread1(1).SetText COL_ARTSCHL, Row, Trim(oF2MCodeSchl.Schl)

335                 Select Case UCase(Trim(knz))

                        Case "A", "L"
                            
340                         fpSpread1(1).SetText COL_ARTIKEL, Row, Trim(oF2MCodeSchl.bez)
345                         fpSpread1(1).SetText COL_MENGE, Row, Trim(oF2MCodeSchl.Menge)
350                         fpSpread1(1).SetText COL_EINHEIT, Row, Trim(oF2MCodeSchl.Einheit)

355                         If Trim(oF2MCodeSchl.Einheit) = "%" Then

360                             fpSpread1(1).SetText COL_EPREIS, Row, LetzterBetrag(Row)

                            Else
                            
365                             If Trim(UCase(oF2MCodeSchl.Wrg)) = Trim(UCase(frmParent.lbl2(17))) Then
                                
370                                 fpSpread1(1).Col = COL_EPREIS
375                                 fpSpread1(1).Row = Row
380                                 fpSpread1(1).TypeNumberDecPlaces = postCommaPreis

385                                 fpSpread1(1).SetText COL_EPREIS, Row, oF2MCodeSchl.Preis

                                Else

390                                 fpSpread1(1).SetText COL_EPREIS, Row, 0

                                End If

                            End If
                        
395                         fpSpread1(1).GetFloat COL_EPREIS, Row, ret
400                         fpSpread1(1).Row = Row
405                         fpSpread1(1).Col = COL_EPREIS

410                         If ret = 0 Then
415                             fpSpread1(1).ForeColor = vbRed
                            Else
420                             fpSpread1(1).ForeColor = vbBlack
                            End If

425                         fpSpread1(1).SetText COL_RABATT, Row, oF2MCodeSchl.Rabatt
                        
430                         Select Case intSteuerTyp

                                Case 0, 2
435                                 fpSpread1(1).SetText COL_UST, Row, "0"

440                             Case 1
445                                 fpSpread1(1).SetText COL_UST, Row, "1"

                            End Select

450                         fpSpread1(1).SetText COL_DURCHLAUFEND, Row, oF2MCodeSchl.Durchlaufend
455                         fpSpread1(1).SetText COL_KOSTSCHL, Row, Trim(oF2MCodeSchl.KostSchl)
460                         fpSpread1(1).SetText COL_SACHSCHL, Row, Trim(oF2MCodeSchl.FiBuSchl)
465                         fpSpread1(1).SetText COL_KOSTKTO, Row, oF2MCodeSchl.KostKonto
470                         fpSpread1(1).SetText COL_SACHKTO, Row, oF2MCodeSchl.FibuKonto

475                         If Trim(oF2MCodeSchl.TextSchl) <> "" Then

480                             LR = LastRow - Row

485                             If LR > 0 Then ZeileBearbeiten 2, Row + 1, LR
                            
490                             fpSpread1(1).SetText COL_ZEILENART, Row + 1, "T"

495                             ZeilenTyp Row + 1

500                             fpSpread1(1).SetText COL_ARTSCHL, Row + 1, Trim(oF2MCodeSchl.TextSchl)

505                             objSQLAusw.GetIfOnesHit = True

510                             Auswahl COL_ARTSCHL, Row + 1

515                             objSQLAusw.GetIfOnesHit = False

                            Else
                            
520                             ZeileBearbeiten 5, 0, 0

                            End If

525                         If Trim(oF2MCodeSchl.bez) = "" Then

530                             ActiveCell = COL_ARTIKEL
                                
                            Else
                            
535                             If fpSpread1(1).colWidth(COL_LS_DATUM) = 0 And fpSpread1(1).colWidth(COL_LS_NUMMER) = 0 Then
                            
540                                 ActiveCell = COL_MENGE
                            
                                Else
                            
545                                 ActiveCell = COL_LS_DATUM
                            
                                End If
    
                            End If
                            
550                         If boolEnterGedruckt Then

555                             boolEnterGedruckt = False

560                             ActiveCell = ActiveCell - 1

565                             If fpSpread1(1).colWidth(COL_LS_DATUM) = 0 And fpSpread1(1).colWidth(COL_LS_NUMMER) = 0 And ActiveCell = COL_LS_NUMMER Then ActiveCell = ActiveCell - 2

                            Else

                            End If

570                         Call SetActiveCellExt(1, ActiveCell, Row, True)

575                         Timer1.Enabled = True

580                         Timer1.Interval = 2

585                         SelectAllText = True

590                     Case "T"

                            Dim blnIsFormelText As Boolean

595                         txt1(0) = Trim(oF2MCodeSchl.Inhalt)
                        
600                         If Row > 1 Then

605                             fpSpread1(1).GetText COL_ZEILENART, Row - 1, knz

610                             If UCase(Trim(knz)) = "A" Then

615                                 fpSpread1(1).GetText COL_MENGE, Row - 1, vMenge
620                                 fpSpread1(1).GetText COL_EINHEIT, Row - 1, vEinheit

625                                 txt1(0).text = GetFormelText(txt1(0).text, vEinheit, vMenge, Row - 1, blnIsFormelText)

                                End If

                            End If

630                         Merk = GetRealText(txt1(0), TextRows)

635                         If blnIsFormelText Then Call UpdateFormelTextRows(Row - 1, TextRows)

640                         If TextRows = 1 Then
                            
645                             fpSpread1(1).SetText COL_ARTIKEL, Row, txt1(0)

650                             Call SetActiveCellExt(1, COL_ARTIKEL, Row, True)

                            Else

655                             If gbZeileInCopy = False Then
                                
660                                 LR = LastRow - Row
                                
665                                 If dictFormelRows.Exists(Row - 1) Then

                                        Dim lngLastFormelTextRow As Long

670                                     lngLastFormelTextRow = DEFAULT_VALUE(dictFormelRows.Item(Row - 1), 0, True)

675                                     If lngLastFormelTextRow < LR Then

680                                         LR = LR - lngLastFormelTextRow - 1

                                        End If

                                    End If

685                                 If LR > 0 And dictFormelRows.Exists(Row - 1) = False Then

690                                     ZeileBearbeiten 2, Row + 1, LR

                                    End If

                                End If

695                             For i = Row + 1 To Row + TextRows - 1

700                                 fpSpread1(1).SetText COL_ZEILENART, i, "T"

705                                 SetActiveCellExt 1, COL_ZEILENART, i, False

710                                 ZeilenTyp i

715                             Next i

720                             fpSpread1(1).Row = Row
725                             fpSpread1(1).Row2 = i
730                             fpSpread1(1).Col = COL_ARTIKEL
735                             fpSpread1(1).Col2 = COL_ARTIKEL
740                             fpSpread1(1).Clip = Merk

745                             If dictFormelRows.Exists(Row - 1) Then

750                                 Call SetActiveCellExt(1, COL_MENGE, Row - 1, True)

                                Else

755                                 Call SetActiveCellExt(1, COL_ZEILENART, i, True)

                                End If

760                             fpSpread1(1).Col = COL_ZEILENART

765                             If LR > 0 Or gbZeileInCopy Then

770                                 If dictFormelRows.Exists(Row - 1) Then

775                                     ZeileBearbeiten 3, LR + 1, 0

                                    Else

780                                     ZeileBearbeiten 3, LastRow + 1, 0

                                    End If

785                                 ZwischenSummeRefresh 1, True

                                End If

                            End If

                    End Select

                Else

790                 Call SetActiveCellExt(1, Col, Row, False)

                End If
                
795         Case COL_KOSTSCHL, COL_SACHSCHL
            
800             If fpSpread1(1).GetText(Col, Row, Knz1) And Trim(Knz1) <> "" Then

805                 objSQLAusw.Find = "Schl LIKE '" & Knz1 & "'"

                End If

810             If Col = COL_KOSTSCHL Then
815                 objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl,Bezeichnung,Erlä¶se,Kosten FROM [1100_FiBuKostenStellen]"
                Else
820                 objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl,Bezeichnung,Erlä¶se,Kosten FROM [1100_FiBuSachkonten]"
                End If

825             If objSQLAusw.Abbruch = False Then

830                 fpSpread1(1).SetText Col, Row, objSQLAusw.FieldText(0)

835                 Call SetActiveCellExt(1, Col + 1, Row, True)

                Else

840                 Call SetActiveCellExt(1, Col, Row, False)

                End If
            
845         Case COL_EINHEIT

850             If fpSpread1(1).GetText(Col, Row, knz) And Trim(knz) <> "" Then

855                 objSQLAuswDef.Find = "Knz like '" & knz & "*'"

860                 Merk = knz

                End If

                Dim oF2Verpackung As ResultF2_Verpackung

865             TextDummy.text = knz

870             oF2Verpackung = GetF2_Verpackung("Schl", TextDummy, 0, Me, cReSize, objPRM, False, ColLeft, RowBottom)

875             If oF2Verpackung.Canceled = False Then

880                 fpSpread1(1).SetText Col, Row, Trim(oF2Verpackung.Schl)

885                 modERechnung.boolVPSchlusselanderung = True

890                 If Not CheckENCodeBeiVerpackung(0, oF2Verpackung.Schl) Then

895                     fpSpread1(1).SetText Col, Row, ""

900                     Call SetActiveCellExt(1, Col, Row, False)

                    Else

905                     Call SetActiveCellExt(1, Col + 1, Row, True)

                    End If

                Else

910                 Call SetActiveCellExt(1, Col, Row, False)

                End If

        End Select

915     If fpSpread1(1).GetText(Col, Row, knz) And Trim(knz) <> Vorher Then Schalter True

        Exit Sub

Fehler:
920     Call FehlerErklärung("frmSP52831", "Auswahl()")

End Sub

Public Sub ZeileEifuegen(Row As Long, Anzahl As Long)
    
        On Error GoTo Fehler
    
        Dim array1 As Long

        Dim array2 As Long

        Dim i      As Integer

        Dim LR     As Integer
  
100     LR = LastRow

105     If LR > Row Then
        
110         array1 = LR - Row
115         array2 = fpSpread1(1).MaxCols
    
120         ReDim fparray(array1, array2) As Variant
125         ReDim fparrayArt(array1, 1) As Variant
        
130         fpSpread1(1).GetArray 1, Row, fparray
135         fpSpread1(1).GetArray 1, Row, fparrayArt
        
140         fpSpread1(1).ClearRange 1, Row, fpSpread1(1).MaxCols, LR, True
        
145         fpSpread1(1).SetArray 1, Row + Anzahl, fparrayArt
    
150         LR = LastRow

155         For i = Row To LR
160             ZeilenTyp i
165         Next i
        
170         fpSpread1(1).SetArray 1, Row + Anzahl, fparray
        
175         fpSpread1(1).SetActiveCell 1, Row
        End If

        Exit Sub

Fehler:
180     Call FehlerErklärung("frmSP52831", "ZeileEifuegen")

End Sub

Public Sub ZeileBearbeiten(Aktion As Integer, Row As Long, Anzahl As Long)

        On Error GoTo Fehler
        
        Dim array1 As Long

        Dim array2 As Long

        Dim i      As Integer
  
100     Select Case Aktion

            Case 1
            
105             array1 = Anzahl - 1
110             array2 = fpSpread1(1).MaxCols
    
115             ReDim gFPArray(array1, array2) As Variant
120             ReDim gFPArray1(array1, 1) As Variant
            
125             fpSpread1(1).GetArray 1, Row, gFPArray
130             fpSpread1(1).GetArray 1, Row, gFPArray1
    
135             gbZeileInCopy = True

140         Case 2

145             ZeileBearbeiten 1, Row, Anzahl
150             ZeileBearbeiten 4, Row, Anzahl

155         Case 3

160             If gbZeileInCopy Then
                
165                 fpSpread1(1).SetArray 1, Row, gFPArray1

170                 For i = Row To Row + UBound(gFPArray, 1)

175                     ZeilenTyp i

180                 Next i
                
185                 fpSpread1(1).SetArray 1, Row, gFPArray
                
190                 fpSpread1(1).SetActiveCell 1, Row

195                 gbZeileInCopy = False

                End If

200         Case 4
            
205             fpSpread1(1).ClearRange 1, Row, fpSpread1(1).MaxCols, Row + Anzahl - 1, True

210             For i = Row To Row + Anzahl - 1

215                 ZeilenTyp i

220             Next i
                
225             If Not dictFormelRows Is Nothing Then
                
230                 If dictFormelRows.Exists(Row) Then dictFormelRows.Remove (Row)
                    
                End If
                
235             If Not dictLSRows Is Nothing Then
                    
240                 If dictLSRows.Exists(Row) Then dictLSRows.Remove (Row)
                    
                End If
                
245         Case 5

250             ReDim gFPArray(0, 0) As Variant

255             ReDim gFPArray1(0, 0) As Variant

260             gbZeileInCopy = False

        End Select

        Exit Sub

Fehler:
265     Call FehlerErklärung("frmSP52831", "ZeileBearbeiten")

End Sub

Public Sub FolgeZeigen(BelegID As Long, _
                       Optional blnPPIgnore As Boolean, _
                       Optional Erweitern As Boolean = False)

        On Error GoTo Fehler

        Dim rs As ADODB.Recordset

        Dim i  As Long
  
100     Set rs = New ADODB.Recordset
  
105     fpSpread1(1).ClearRange 1, IIf(Erweitern, LastRow + 1, 1), fpSpread1(1).MaxCols, fpSpread1(1).MaxRows, True
110     fpSpread1(1).Col = 2
115     fpSpread1(1).Row = IIf(Erweitern, LastRow + 1, 1)
120     fpSpread1(1).Col2 = fpSpread1(1).MaxCols
125     fpSpread1(1).Row2 = fpSpread1(1).MaxRows
130     fpSpread1(1).BlockMode = True
135     fpSpread1(1).Lock = True
140     fpSpread1(1).Col = COL_UST
145     fpSpread1(1).Col2 = COL_UST
150     fpSpread1(1).CellType = CellTypeStaticText
155     fpSpread1(1).Col = COL_ARTSCHL
160     fpSpread1(1).Col2 = COL_ARTSCHL
165     fpSpread1(1).CellType = CellTypeStaticText
170     fpSpread1(1).Col = COL_DURCHLAUFEND
175     fpSpread1(1).Col2 = COL_DURCHLAUFEND
180     fpSpread1(1).CellType = CellTypeStaticText
185     fpSpread1(1).BlockMode = False
        
190     fpSpread1(1).Col = 1
195     fpSpread1(1).Row = 1

200     rs.Open "SELECT * FROM [2800_Folge] WHERE BelegID = " & BelegID & " ORDER BY Nr", gConn, adOpenStatic, adLockReadOnly

205     If rs.RecordCount > 0 Then

210         GstrAuftraggeber = ""

215         i = IIf(Erweitern, LastRow, 0)

220         Do Until rs.EOF

225             i = i + 1

230             fpSpread1(1).SetText COL_ZEILENART, i, rs!SatzTyp

235             ZeilenTyp i

240             Select Case UCase(rs!SatzTyp)

                    Case "A", "P", "L"
                    
245                     fpSpread1(1).SetText COL_ARTSCHL, i, rs!Schl
250                     fpSpread1(1).SetText COL_ARTIKEL, i, rs!bez
255                     fpSpread1(1).SetText COL_MENGE, i, rs!Menge
260                     fpSpread1(1).SetText COL_EINHEIT, i, rs!Einheit

265                     fpSpread1(1).Col = COL_EPREIS
270                     fpSpread1(1).Row = i
275                     fpSpread1(1).TypeNumberDecPlaces = postCommaPreis

280                     fpSpread1(1).SetText COL_EPREIS, i, rs!EPreis
285                     fpSpread1(1).SetText COL_RABATT, i, rs!Rabatt
290                     fpSpread1(1).SetText COL_UST, i, rs!Steuer
295                     fpSpread1(1).SetText COL_DURCHLAUFEND, i, rs!Durchlaufend
300                     fpSpread1(1).SetText COL_KOSTSCHL, i, rs!KostSchl
305                     fpSpread1(1).SetText COL_SACHSCHL, i, rs!FiBuSchl
310                     fpSpread1(1).SetText COL_KOSTKTO, i, rs!KostKonto
315                     fpSpread1(1).SetText COL_SACHKTO, i, rs!FibuKonto
                        
320                     Select Case UCase(rs!SatzTyp)
                            
                            Case "L"
                            
325                             fpSpread1(1).SetText COL_LS_DATUM, i, rs!LsDatum
330                             fpSpread1(1).SetText COL_LS_NUMMER, i, rs!LsNr

                        End Select
                        
335                 Case "S"

340                 Case "T"

345                     fpSpread1(1).SetText COL_ARTSCHL, i, rs!Schl
350                     fpSpread1(1).SetText COL_ARTIKEL, i, rs!bez
                        
355                     If InStr(1, UCase(rs!Schl), C_STR_FORMELTEXT_SCHL) > 0 Then
                        
360                         Call UpdateFormelTextRows(i - 1, 0)
                            
                        End If
                        
365                 Case "Z"

370                     fpSpread1(1).SetText COL_ARTIKEL, i, rs!bez

                End Select

375             fpSpread1(1).SetText COL_ERSTDAT, i, rs!ErstDat
380             fpSpread1(1).SetText COL_ERSTVON, i, rs!ErstVon
385             fpSpread1(1).SetText COL_AENDDAT, i, rs!AendDat
390             fpSpread1(1).SetText COL_AENDVON, i, rs!AendVon

395             rs.MoveNext

            Loop

        End If

400     rs.Close
405     Set rs = Nothing
  
410     If m_BelegNeu And (GintBelegArt = 0 Or GintBelegArt = 1 Or GintBelegArt = 2) And Not blnPPIgnore Then
        
415         If IsNumeric(frmParent.txt1(24)) Then

420             If CCur(frmParent.txt1(24)) > 0 Then
                
425                 fpSpread1(1).SetText COL_ZEILENART, i + 1, "P"
430                 fpSpread1(1).SetText COL_ARTSCHL, i + 1, ""
435                 fpSpread1(1).SetText COL_ARTIKEL, i + 1, CStr(frmParent.lbl1(24).caption)
440                 fpSpread1(1).SetText COL_MENGE, i + 1, "1"
445                 fpSpread1(1).SetText COL_EINHEIT, i + 1, C_STR_COL_EINHEIT_STUECK
450                 fpSpread1(1).SetText COL_EPREIS, i + 1, Format(frmParent.txt1(24), "0.00")

455                 ZeilenTyp i + 1

                End If
                
            End If
            
        End If

460     cmd1(1).Enabled = True
465     mnu_Bearb_U(1).Enabled = True
    
470     If frmParent.gintDruck = 1 Then

        Else
        
475         If GstrAuftraggeber <> "" Then
        
480             i = i + 1
485             fpSpread1(1).SetText COL_ZEILENART, i, "T"
490             fpSpread1(1).SetText COL_ARTIKEL, i, "Gemä¤äŸ Auftrag: " & GstrAuftraggeber
        
            End If
        
495         GstrAuftraggeber = ""
        
500         fpSpread1(1).SetActiveCell COL_ZEILENART, i + 1

        End If
  
505     If BelegNeu Then

510         cmd1(1).Enabled = False
515         mnu_Bearb_U(1).Enabled = False

        End If
  
520     If frmParent.Check1(0) = 0 And frmParent.gintZwAblage = 1 Then
            
525         frmParent.glngBelegID = 0
530         cmd1(0).Enabled = True
535         mnu_Bearb_U(0).Enabled = True

        Else
        
540         cmd1(0).Enabled = False
545         mnu_Bearb_U(0).Enabled = False

        End If
        
550     SteuernKontrol

        Exit Sub

Fehler:

555     Call FehlerErklärung("frmSP52831", "FolgeZeigen()")

End Sub

Private Sub fpSpread1_KeyPress(Index As Integer, KeyAscii As Integer)

        On Error GoTo Fehler

        Dim knz As Variant
  
100     If Index = 1 Then

105         Select Case fpSpread1(Index).ActiveCol

                Case COL_ZEILENART
                
110                 KeyAscii = Asc(UCase(Chr(KeyAscii)))
  
115                 Select Case KeyAscii

                        Case 65, 76, 80, 83, 84, 90

120                         If fpSpread1(Index).GetText(COL_ZEILENART, fpSpread1(Index).ActiveRow, knz) Then

125                             If Len(knz) >= 1 And fpSpread1(Index).selLength < Len(knz) Then

130                                 KeyAscii = 0

                                End If
                                
                            End If

135                     Case 8

140                     Case Else

145                         KeyAscii = 0

                    End Select

150             Case COL_ARTSCHL
                
155                 If KeyAscii <> 8 Then
160                     If fpSpread1(Index).GetText(COL_ARTSCHL, fpSpread1(Index).ActiveRow, knz) Then
165                         If Len(knz) >= 12 Then
170                             KeyAscii = 0
                            End If
                        End If
                    End If
                    
175             Case COL_ARTIKEL
                    
180                 If KeyAscii <> 8 Then
                
185                     If fpSpread1(Index).GetText(COL_ARTIKEL, fpSpread1(Index).ActiveRow, knz) Then

190                         Me.Font.name = fpSpread1(1).Font.name
195                         Me.Font.Size = fpSpread1(1).Font.Size
                        
200                         If Me.TextWidth(knz) >= txt1(0).width Then
205                             KeyAscii = 0
                            End If
                        
                        End If
                    
                    End If

210             Case COL_EINHEIT
                
215                 If KeyAscii <> 8 Then
220                     If fpSpread1(Index).GetText(COL_EINHEIT, fpSpread1(Index).ActiveRow, knz) Then
225                         If Len(knz) >= 20 Then
230                             KeyAscii = 0
                            End If
                        End If
                    End If

235             Case COL_KOSTSCHL, COL_SACHSCHL
                
240                 Select Case KeyAscii

                        Case 8
245                         fpSpread1(Index).SetText fpSpread1(Index).ActiveCol, fpSpread1(Index).ActiveRow, ""

250                     Case Else
255                         KeyAscii = 0
                    End Select
                    
            End Select

        End If

        Exit Sub

Fehler:

260     Call FehlerErklärung("frmSP52831", "fpSpread1_KeyPress()")

End Sub

Public Sub FussZahlungsZiel()

        On Error GoTo Fehler

        Dim knz As Variant
        
100     fpSpread1(2).Col = COL_FS_STEUERPFLICHT
105     fpSpread1(2).Row = 5

110     fpSpread1(2).Col2 = COL_FS_STEUERFREI
115     fpSpread1(2).Row2 = 6

120     If fpSpread1(2).GetText(COL_FS_GESAMT, 2, knz) = False Then
125         knz = 0
        End If
    
130     fpSpread1(2).Clip = ZahlungsZiel(frmParent.belegDatum, knz, frmParent.lbl2(17), frmParent.txt1(21), frmParent.txt1(22), frmParent.txt1(23), frmParent.ValutaDatum)

135     fpSpread1(2).Col = COL_FS_GESAMT
140     fpSpread1(2).Row = 5
145     fpSpread1(2).Col2 = COL_FS_GESAMT
150     fpSpread1(2).Row2 = 6
    
155     fpSpread1(2).Clip = ZahlungsZielNetto(frmParent.belegDatum, knz, frmParent.lbl2(17), frmParent.txt1(21), frmParent.txt1(22), frmParent.txt1(23), frmParent.ValutaDatum)

        Exit Sub

Fehler:
160     Call FehlerErklärung("frmSP52831", "FussZahlungsZiel()")

End Sub

Public Sub FussZahlungsZielGesamtIstBrutto()

        On Error GoTo Fehler

        Dim knz As Variant

100     If fpSpread1(2).GetText(COL_FS_GESAMT, 2, knz) = False Then
105         knz = 0
        End If

110     fpSpread1(2).Col = COL_FS_STEUERPFLICHT
115     fpSpread1(2).Row = 5
120     fpSpread1(2).Col2 = COL_FS_STEUERFREI
125     fpSpread1(2).Row2 = 6
    
130     fpSpread1(2).Clip = ZahlungsZiel(frmParent.belegDatum, knz, frmParent.lbl2(17), frmParent.txt1(21), frmParent.txt1(22), frmParent.txt1(23), frmParent.ValutaDatum)

135     fpSpread1(2).Col = COL_FS_GESAMT
140     fpSpread1(2).Row = 5
145     fpSpread1(2).Col2 = COL_FS_GESAMT
150     fpSpread1(2).Row2 = 6
    
155     fpSpread1(2).Clip = ZahlungsZielNetto(frmParent.belegDatum, knz, frmParent.lbl2(17), frmParent.txt1(21), frmParent.txt1(22), frmParent.txt1(23), frmParent.ValutaDatum)

        Exit Sub

Fehler:
160     Call FehlerErklärung("frmSP52831", "FussZahlungsZielGesamtIstBrutto()")

End Sub

Private Sub fpSpread1_LeaveCell(Index As Integer, _
                                ByVal Col As Long, _
                                ByVal Row As Long, _
                                ByVal newCol As Long, _
                                ByVal NewRow As Long, _
                                Cancel As Boolean)

        On Error GoTo Fehler

        Dim knz As Variant
        
100     If Index = 1 Then
        
105         Call SetCellBackColor(Index, Col, Row, True)

110         If Col = COL_EPREIS Then

115             If PlausiEPreis = False Then

120                 Cancel = True

                    Exit Sub

                End If
                
            End If

125         Select Case newCol

                Case COL_ARTSCHL

130                 If fpSpread1(1).GetText(COL_ARTSCHL, NewRow, knz) Then

135                     gvrnMerker = knz
                    
                    Else
                        
140                     gvrnMerker = ""
                        
                    End If

            End Select

        End If
    
145     If blnChangeEventFired Then
        
150         Call RefreshFoot
155         blnChangeEventFired = False
        End If
        
160     Call SetCellBackColor(1, newCol, NewRow, False)

        Exit Sub

Fehler:
165     Call FehlerErklärung("frmSP52831", "fpSpread1_LeaveCell()")

End Sub

Private Sub fpSpread1_Validate(Index As Integer, Cancel As Boolean)

        On Error GoTo Fehler

100     If PlausiEPreis = False Then Cancel = True

105     Call RefreshFoot

        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52831", "fpSpread1_Validate()")

End Sub

Private Sub RefreshFoot()

        On Error GoTo Fehler

100     If GesamtIstBrutto Then
105         FussZahlungsZielGesamtIstBrutto
        Else
110         FussZahlungsZiel
        End If

        Exit Sub

Fehler:
115     Call FehlerErklärung("frmSP52831", "RefreshFoot()")

End Sub

Private Sub imgSplitter_MouseDown(Index As Integer, _
                                  Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
        On Error GoTo Fehler
    
100     With imgSplitter(Index)
105         picSplitter(Index).Move .left, .top \ 2, .width, .height
        End With

110     picSplitter(Index).Visible = True

        Exit Sub

Fehler:
115     Call FehlerErklärung("frmSP52831", "imgSplitter_MouseDown")

End Sub

Private Sub imgSplitter_MouseMove(Index As Integer, _
                                  Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
        On Error GoTo Fehler

        Dim pos      As Single

        Dim SplitMin As Single

        Dim SplitMax As Single
  
100     If Index = 0 Then
105         SplitMin = 1200
110         SplitMax = gsngKopfHeight
        Else
115         SplitMin = gsngFussTop - imgSplitter(Index).height
120         SplitMax = Me.height - 1500
        End If
  
125     pos = Y + imgSplitter(Index).top
  
130     If pos < SplitMin Then
135         picSplitter(Index).top = SplitMin
140     ElseIf pos > SplitMax Then
145         picSplitter(Index).top = SplitMax
        Else
150         picSplitter(Index).top = pos
        End If

        Exit Sub

Fehler:
155     Call FehlerErklärung("frmSP52831", "imgSplitter_MouseMove")

End Sub

Private Sub imgSplitter_MouseUp(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
        On Error GoTo Fehler
    
100     SizeControls picSplitter(Index).top, Index
105     picSplitter(Index).Visible = False

        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52831", "imgSplitter_MouseUp")

End Sub

Public Sub FolgeSpeichern(BelegID As Long, _
                          Optional tmp As Boolean, _
                          Optional GetSteuerPflichtig As Double, _
                          Optional GetSteuerFrei As Double)
    
        On Error GoTo Fehler
    
        Dim Row       As Integer

        Dim knz       As Variant

        Dim rsFolge   As ADODB.Recordset

        Dim TmpZusatz As String
  
100     Set rsFolge = New ADODB.Recordset
  
105     If tmp Then
110         TmpZusatz = "Tmp"
        Else
115         cmd1(0).Enabled = False
120         mnu_Bearb_U(0).Enabled = False
        End If
  
125     rsFolge.Open "SELECT * FROM [2800_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID & " ORDER BY Nr", gConn, adOpenKeyset, adLockOptimistic

130     Do Until rsFolge.EOF
135         rsFolge.Delete
140         rsFolge.MoveNext
        Loop

145     For Row = 1 To LastRow
150         rsFolge.AddNew
    
155         rsFolge!BelegID = BelegID
160         rsFolge!nr = Row

165         If fpSpread1(1).GetText(COL_ZEILENART, Row, knz) Then rsFolge!SatzTyp = knz
170         If knz = "S" Then
175             If fpSpread1(1).GetText(COL_ARTIKEL, Row, knz) Then rsFolge!bez = left(knz, 1)
            Else

180             If fpSpread1(1).GetText(COL_ARTIKEL, Row, knz) Then rsFolge!bez = knz
            End If

185         If fpSpread1(1).GetText(COL_ARTSCHL, Row, knz) Then rsFolge!Schl = knz
190         If fpSpread1(1).GetText(COL_MENGE, Row, knz) Then rsFolge!Menge = knz
195         If fpSpread1(1).GetText(COL_EINHEIT, Row, knz) Then rsFolge!Einheit = Trim(knz)
200         If fpSpread1(1).GetText(COL_EPREIS, Row, knz) Then
            
205             rsFolge!EPreis = knz
            End If

210         If fpSpread1(1).GetText(COL_RABATT, Row, knz) Then rsFolge!Rabatt = knz
215         If fpSpread1(1).GetText(COL_UST, Row, knz) Then rsFolge!Steuer = knz
220         If fpSpread1(1).GetText(COL_DURCHLAUFEND, Row, knz) Then rsFolge!Durchlaufend = knz
225         If fpSpread1(1).GetText(COL_KOSTSCHL, Row, knz) Then rsFolge!KostSchl = knz
230         If fpSpread1(1).GetText(COL_SACHSCHL, Row, knz) Then rsFolge!FiBuSchl = knz
235         If fpSpread1(1).GetText(COL_KOSTKTO, Row, knz) Then rsFolge!KostKonto = knz
240         If fpSpread1(1).GetText(COL_SACHKTO, Row, knz) Then rsFolge!FibuKonto = knz
245         If fpSpread1(1).GetText(COL_ERSTDAT, Row, knz) Then rsFolge!ErstDat = CStr(Now)
250         If fpSpread1(1).GetText(COL_ERSTVON, Row, knz) Then rsFolge!ErstVon = GsUser
255         If fpSpread1(1).GetText(COL_AENDDAT, Row, knz) Then rsFolge!AendDat = knz
260         If fpSpread1(1).GetText(COL_AENDVON, Row, knz) Then rsFolge!AendVon = knz
        
265         If fpSpread1(1).GetText(COL_LS_DATUM, Row, knz) Then rsFolge!LsDatum = knz
270         If fpSpread1(1).GetText(COL_LS_NUMMER, Row, knz) Then rsFolge!LsNr = knz
275         rsFolge.Update
        
280         rsFolge.MoveLast
            
285         If Trim(rsFolge!Einheit) = "%" Then
            
290             If GesamtIstBrutto Then

295                 If rsFolge!Steuer = 1 Then
300                     GetSteuerPflichtig = GetSteuerPflichtig + RundenMitVz((rsFolge!Menge * rsFolge!EPreis / 100 - rsFolge!Menge * rsFolge!EPreis / 100 * rsFolge!Rabatt / 100), 2)
                    Else
305                     GetSteuerFrei = GetSteuerFrei + RundenMitVz((rsFolge!Menge * rsFolge!EPreis / 100 - rsFolge!Menge * rsFolge!EPreis / 100 * rsFolge!Rabatt / 100), 2)
                    End If

                Else

310                 If rsFolge!Steuer = 1 Then
315                     GetSteuerPflichtig = GetSteuerPflichtig + RundenMitVz((rsFolge!Menge * rsFolge!EPreis / 100 - rsFolge!Menge * rsFolge!EPreis / 100 * rsFolge!Rabatt / 100), 2)
                    Else
320                     GetSteuerFrei = GetSteuerFrei + RundenMitVz((rsFolge!Menge * rsFolge!EPreis / 100 - rsFolge!Menge * rsFolge!EPreis / 100 * rsFolge!Rabatt / 100), 2)
                    End If

                End If

            Else

325             If GesamtIstBrutto Then

330                 If rsFolge!Steuer = 1 Then
335                     GetSteuerPflichtig = GetSteuerPflichtig + RundenMitVz((rsFolge!Menge * rsFolge!EPreis - rsFolge!Menge * rsFolge!EPreis * rsFolge!Rabatt / 100), 2)
                    Else
340                     GetSteuerFrei = GetSteuerFrei + RundenMitVz((rsFolge!Menge * rsFolge!EPreis - rsFolge!Menge * rsFolge!EPreis * rsFolge!Rabatt / 100), 2)
                    End If

                Else

345                 If rsFolge!Steuer = 1 Then
350                     GetSteuerPflichtig = GetSteuerPflichtig + RundenMitVz((rsFolge!Menge * rsFolge!EPreis - rsFolge!Menge * rsFolge!EPreis * rsFolge!Rabatt / 100), 2)
                    Else
355                     GetSteuerFrei = GetSteuerFrei + RundenMitVz((rsFolge!Menge * rsFolge!EPreis - rsFolge!Menge * rsFolge!EPreis * rsFolge!Rabatt / 100), 2)
                    End If
                    
                End If
                
            End If
        
360     Next Row
        
365     If Trim$(strAutomatischText) = "" Then strAutomatischText = "Automatisch"
        
370     If tmp Then

375         fpSpread1(0).SetText 2, 17, strAutomatischText
380         fpSpread1(0).Col = 2
385         fpSpread1(0).Row = 17
390         objDruckOptionen.CurrentBelegNr = "0"
        
395         fpSpread1(0).SetText 4, 17, strAutomatischText
400         fpSpread1(0).Col = 4
405         fpSpread1(0).Row = 17

        Else

410         If objDruckOptionen.CurrentBelegNr = "" Or objDruckOptionen.CurrentBelegNr = "0" Then
            
415             fpSpread1(0).SetText 2, 17, strAutomatischText

            Else

420             If objDruckOptionen.CurrentBelegNr <> frmParent.BelegNr And Trim(frmParent.BelegNr) <> "" Then
425                 objDruckOptionen.CurrentBelegNr = frmParent.BelegNr
                End If

430             If frmParent.Check1(0).value = 0 Then fpSpread1(0).SetText 2, 17, objDruckOptionen.CurrentBelegNr

            End If

435         If objDruckOptionen.CurrentBelegDatum = "" Then
            
440             fpSpread1(0).SetText 4, 17, strAutomatischText

            Else

445             If objDruckOptionen.CurrentBelegDatum <> frmParent.belegDatum And Trim(frmParent.belegDatum) <> "" Then
450                 objDruckOptionen.CurrentBelegDatum = frmParent.belegDatum
                End If

455             fpSpread1(0).SetText 4, 17, objDruckOptionen.CurrentBelegDatum

            End If
            
        End If

        Exit Sub

Fehler:
460     Call FehlerErklärung("frmSP52831", "FolgeSpeichern()")

End Sub

Public Function GetEndBetrag() As Double

        Dim knz As Variant
  
100     If fpSpread1(2).GetText(COL_FS_GESAMT, 2, knz) Then
105         GetEndBetrag = CDbl(knz)
        End If

End Function

Public Function LetzterBetrag(ByVal Row As Long) As Double
    
        On Error GoTo Fehler
    
        Dim i   As Integer

        Dim knz As Variant
  
100     fpSpread1(1).Row = Row
  
105     For i = fpSpread1(1).Row - 1 To 1 Step -1

110         If fpSpread1(1).GetText(COL_GPREIS, i, knz) Then
115             If IsNumeric(knz) Then
120                 LetzterBetrag = knz

                    Exit For

                End If
            End If

125     Next i

        Exit Function

Fehler:
130     Call FehlerErklärung("frmSP52831", "LetzterBetrag")

End Function

Public Function LastRow() As Long
    
        On Error GoTo Fehler
    
        Dim Row As Integer

        Dim knz As Variant

100     For Row = 1 To fpSpread1(1).MaxRows

105         If fpSpread1(1).GetText(COL_ZEILENART, Row, knz) Then
110             If Trim(knz) <> "" Then LastRow = Row
            End If

115     Next Row

        Exit Function

Fehler:
120     Call FehlerErklärung("frmSP52831", "LastRow")

End Function

Public Sub Schalter(Enabled As Boolean)

        On Error GoTo Fehler

100     cmd1(0).Enabled = Enabled
105     mnu_Bearb_U(0).Enabled = Enabled
        
110     If GbDesigner Then
115         cmd1(12).Enabled = Enabled
120         mnu_Bearb_U(12).Enabled = Enabled
        End If
  
125     mnu_Drucken_U(0).Enabled = Enabled

130     cmd1(5).Enabled = Enabled
135     mnu_Bearb_U(5).Enabled = Enabled

140     mnu_Drucken_U(2).Enabled = Enabled
145     cmd1(4).Enabled = Enabled
150     mnu_Bearb_U(4).Enabled = Enabled

155     cmd1(10).Enabled = Enabled
160     mnu_Bearb_U(10).Enabled = Enabled

165     cmd1(6).Enabled = Enabled
170     mnu_Bearb_U(6).Enabled = Enabled

        Exit Sub

Fehler:
175     Call FehlerErklärung("frmSP52831", "Schalter")

End Sub

Public Function Plausi() As Boolean
    
        On Error GoTo Fehler
    
        Dim ret As Double
    
100     fpSpread1(2).GetFloat COL_FS_GESAMT, 2, ret
    
105     If 1 < 1 Then

110         MsgBox "Gesamtbetrag = (" & Format(ret, "###,###,##0.00") & "). Der Beleg kann nicht Gespeichert werden.", vbExclamation, strMeldungCap

        Else

115         If IstArtieklUndMengeOK Then

120             If IsEinheitOK Then

125                 If IstPorto Then

130                     Plausi = True

                    Else

135                     If MsgBox("Sie haben keinen Porto-Betrag erfasst. Soll der Beleg trotzdem gespeichert werden?", vbYesNo + vbExclamation, strMeldungCap) = vbYes Then

140                         Plausi = True

                        End If
                        
                    End If
                    
                End If
                
            End If
            
        End If

        Exit Function

Fehler:
145     Call FehlerErklärung("frmSP52831", "Plausi")

End Function

Public Sub EingabeSperren()

100     fpSpread1(1).Col = 1
105     fpSpread1(1).Row = 1
110     fpSpread1(1).Col2 = fpSpread1(1).MaxCols
115     fpSpread1(1).Row2 = fpSpread1(1).MaxRows
120     fpSpread1(1).BlockMode = True
125     fpSpread1(1).Lock = True
130     fpSpread1(1).BlockMode = False

135     cmd1(1).Enabled = False
140     mnu_Bearb_U(1).Enabled = False

145     cmd1(3).Enabled = False
150     mnu_Bearb_U(3).Enabled = False

155     cmd1(2).Enabled = False
160     mnu_Bearb_U(2).Enabled = False

165     cmd1(6).Enabled = False
170     mnu_Bearb_U(6).Enabled = False

175     cmd1(10).Enabled = False
180     mnu_Bearb_U(10).Enabled = False

End Sub

Public Function PlausiEPreis() As Boolean

        Dim i   As Integer

        Dim knz As Variant

        Dim ret As Double

        Dim Row As Long
   
        On Error GoTo Fehler
  
100     PlausiEPreis = True

        Exit Function

105     Row = fpSpread1(1).ActiveRow
  
110     fpSpread1(1).GetFloat COL_EPREIS, Row, ret

115     If ret < 0 Then
        
120         For i = COL_KOSTSCHL To COL_SACHKTO

125             If fpSpread1(1).GetText(i, Row, knz) Then

130                 Select Case i

                        Case COL_KOSTSCHL, COL_SACHSCHL

135                         If Trim(knz) <> "" Or Trim(knz) <> "" Then
140                             PlausiEPreis = False

                                Exit For

                            End If

145                     Case COL_KOSTKTO, COL_SACHKTO

150                         If Trim(knz) <> "" And Trim(knz) <> "0" Then
155                             PlausiEPreis = False

                                Exit For

                            End If

                    End Select

                End If

160         Next i
    
165         If PlausiEPreis = False Then
            
170             Timer1.Interval = 1
            Else
175             fpSpread1(1).Col = COL_KOSTSCHL
180             fpSpread1(1).Col2 = COL_SACHKTO
185             fpSpread1(1).Row = Row
190             fpSpread1(1).Row2 = Row
195             fpSpread1(1).BlockMode = True
200             fpSpread1(1).Lock = True
205             fpSpread1(1).BlockMode = False
            End If

        Else
210         fpSpread1(1).Col = COL_KOSTSCHL
215         fpSpread1(1).Col2 = COL_SACHKTO
220         fpSpread1(1).Row = Row
225         fpSpread1(1).Row2 = Row
230         fpSpread1(1).BlockMode = True
235         fpSpread1(1).Lock = False
240         fpSpread1(1).BlockMode = False
        End If

        Exit Function

Fehler:
245     Call FehlerErklärung("SP52831", "PlausiEPreis")

End Function

Private Sub mnu_Bearb_U_Click(Index As Integer)

        Dim objPlausi As clsPlausi

        On Error GoTo Fehler
        
100     If objPlausi Is Nothing Then

105         Set objPlausi = New clsPlausi
110         objPlausi.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)

        End If
        
115     If Index = 6 Or Index = 4 Then

            Exit Sub
  
120         If Index <> 0 Then

125             Call cmd1_Click(Index)

            Else
        
130             mnuSpeichern(0).Enabled = True
            
135             If BelegInfo.Angenommen Then

140                 If BelegInfo.Art <> GintBelegArt Or BelegInfo.Druck = 1 Or (BelegInfo.Druck = 2 And frmParent.Check1(0).value <> 1) Or (BelegInfo.Druck <> 2 And frmParent.Check1(0).value = 1) Then

145                     mnuSpeichern(0).Enabled = False

                    End If
            
                Else
            
150                 mnuSpeichern(0).Enabled = False
            
                End If
                
            End If
            
        End If

        Exit Sub

Fehler:
155     Call FehlerErklärung("frmSP52831", "mnu_Bearb_U_Click")

End Sub

Private Sub mnu_close_Click()

        On Error GoTo Fehler
  
100     Call cmd1_Click(8)

        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52831", "mnu_close_Click")

End Sub

Private Sub BelegSuchen(BelegTyp As E_DATATYPE, Druck As Integer)

        On Error GoTo Fehler

        Dim oF2BelegSuchen As ResultF2_BelegSuchen

100     Set objPRM.gForm = frmParent

105     oF2BelegSuchen = GetF2_BelegSuchen(BelegTyp, Druck, fpSpread1(0), 0, Me, frmParent.cReSize, objPRM)

110     Set objPRM.gForm = Me

115     If oF2BelegSuchen.Canceled = False Then

120         Call FolgeZeigen(oF2BelegSuchen.BelegID, True, MsgBox(GetMessage(2402), vbYesNo + vbQuestion, strMeldungCap) = vbYes)

        End If

        Exit Sub

Fehler:

125     Me.MousePointer = vbDefault
130     Call FehlerErklärung("frmSP52831", "BelegSuchen()")

End Sub

Private Sub mnu_Drucken_U_Click(Index As Integer)
    
        On Error GoTo Fehler
        
100     Select Case Index
        
            Case 0
            
105             Call Vorschau
            
110         Case 1

115             If MsgBox(GetMessage(2385), vbYesNo + vbExclamation, strMeldungCap) = vbYes Then Call Druck(True)
            
120         Case 2

125             Call Druck
        
        End Select

        Exit Sub

Fehler:
130     Me.MousePointer = vbDefault
135     Call FehlerErklärung("frmSP52831", "mnu_Drucken_U_Click()")

End Sub

Private Sub mnuSpeichern_Click(Index As Integer)

        On Error GoTo Fehler
        
        Dim boolNew As Boolean
        
100     Select Case Index
        
            Case 0
            
105             boolNew = False
            
110         Case 1
            
115             boolNew = True
        
        End Select
        
120     If Not boolNew Then
        
125         Call msgText(1, 2373, 0, 0, 0)

130         If MsgBox(GsMsgText(0), vbYesNo + vbQuestion, strMeldungCap) = vbYes Then Call Speichern(boolNew)

        Else

135         Call Speichern(boolNew)

        End If

        Exit Sub

Fehler:
140     Me.MousePointer = vbDefault
145     Call FehlerErklärung("frmSP52831", "mnuSpeichern_Click()")

End Sub

Private Sub mnuUbernehmenA_Click(Index As Integer)

        On Error GoTo Fehler

100     BelegSuchen Sonderfaktura_Angebot, Index

        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52831", "mnuUbernehmenA_Click()")

End Sub

Private Sub mnuUbernehmenB_Click(Index As Integer)
        
        On Error GoTo Fehler

100     BelegSuchen Sonderfaktura_Auftragsbestetigung, Index

        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52831", "mnuUbernehmenB_Click()")

End Sub

Private Sub mnuUbernehmenR_Click(Index As Integer)

        On Error GoTo Fehler

100     BelegSuchen Sonderfaktura_Rechnung, Index

        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52831", "mnuUbernehmenR_Click()")

End Sub

Private Sub mnuUbernehmenG_Click(Index As Integer)

        On Error GoTo Fehler

100     BelegSuchen Sonderfaktura_Gutschrift, Index

        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52831", "mnuUbernehmenG_Click()")

End Sub

Private Sub SSPanel1_Click(Index As Integer)

100     Debug.Print "Testklick fuer RunTimeDebug!"

End Sub

Private Sub Timer1_Timer()

    If Timer1.Interval = 1 Then
        
        MsgBox "Der Minusbetrag in der Spalte Preis ist nicht zugelassen, wenn Spalten: Kostenstellen-, Sachkonten-Schlä¼ssel, Kostenstellen-, Sach-Konto gefä¼llt sind. ", vbExclamation, strMeldungCap
        
    End If
    
    If SelectAllText Then
    
        LockWindowUpdate Me.hwnd
    
        Timer1.Enabled = False

        On Error Resume Next

        fpSpread1(1).SelStart = 0
        fpSpread1(1).selLength = Len(fpSpread1(1).text)

        Timer1.Interval = 0
        SelectAllText = False
        
        LockWindowUpdate (0&)
        
        Call SetCellBackColor(1, fpSpread1(1).Col, fpSpread1(1).Row, False)
        
    End If
    
End Sub

Private Sub mnuAnsicht_ResFak_Click(Index As Integer)
    
100     Select Case Index

            Case 0
105             Call cReSize.ResizeAboutPercent(0#, 0)

110         Case 2
115             Call cReSize.ResizeAboutPercent(20#, 2)

120         Case 4
125             Call cReSize.ResizeAboutPercent(40#, 4)

130         Case 6
135             Call cReSize.ResizeAboutPercent(60#, 6)

140         Case 8
145             Call cReSize.ResizeAboutPercent(80#, 8)
        End Select

End Sub

Private Sub mnuAnsicht_ResetPosition_Click()
    
100     Call ResetWindowPos(Me.hwnd, "SP51000")
    
105     cReSize.RemoveRegistryKeys

End Sub

Private Sub mnuAnsicht_Alle_Click()
    
100     mnuAnsicht_Alle.Checked = Not mnuAnsicht_Alle.Checked
105     cReSize.ResizeAllForms = mnuAnsicht_Alle.Checked

End Sub

Private Sub mnuAnsicht_Prop_Click()
    
100     mnuAnsicht_Prop.Checked = Not mnuAnsicht_Prop.Checked
105     cReSize.ScalingProportional = mnuAnsicht_Prop.Checked
    
110     cReSize.resize

End Sub

Public Sub MaskeLeeren(leerenModus As MaskeLeerenModus, Optional blnPPIgnore As Boolean)
    
        On Error GoTo Fehler
        
100     frmParent.gintDruck = 0
        
105     Select Case leerenModus
        
            Case MaskeLeerenModus.nurBelegDaten
            
110             frmParent.BelegNr = "0"
115             frmParent.belegDatum = ""
120             frmParent.ValutaDatum = ""
125             frmParent.blnBelegNeu = True
            
130             frmParent.glngBelegID = 0
135             frmParent.glngBelegIDTmp = 0
140             frmParent.glngBelegIDVorlage = 0
            
145             frmParent.txt1(15).text = ""
150             frmParent.txt1(16).text = ""
            
155             If Not objDruckOptionen Is Nothing Then

160                 objDruckOptionen.clearVars
165                 objDruckOptionen.EnableBelegNr = True
170                 objDruckOptionen.EnableBelegDatum = True
175                 objDruckOptionen.EnableValutaDatum = True

                End If
            
180             BelegNeu = True
                                        
185             frmParent.sta1.Panels(3).text = ""

190             Call KopfFuellen

195             If cmd1(4).Enabled Then cmd1(4).SetFocus

                Exit Sub
            
200         Case MaskeLeerenModus.alleDaten

205             Me.MousePointer = vbHourglass

210             Call MaskeLeeren(nurBelegDaten)
215             Call MaskeLeeren(ohneBelegDaten, True)
220             Call frmParent.MaskeLeeren(False)
                
225             Me.MousePointer = vbDefault

                Exit Sub
                
230         Case MaskeLeerenModus.ohneBelegDaten
                
        End Select
    
235     dictLSRows.RemoveAll
        
240     fpSpread1(1).ReDraw = False
            
245     Call PostenSetUp
        
250     Call FolgeZeigen(-999, blnPPIgnore)
                                        
255     Call FussFuellen

260     fpSpread1(1).ReDraw = True
    
265     printDone = False
    
270     If fpSpread1(1).Enabled Then
275         If Me.Visible Then fpSpread1(1).SetFocus
280         fpSpread1(1).SetActiveCell 1, 1
        End If

        Exit Sub
    
Fehler:
285     Me.MousePointer = vbDefault
    
290     Call FehlerErklärung("frmSP52831", "MaskeLeeren()")

End Sub

Public Function SetCellBackColor(Index As Integer, _
                                 Col As Long, _
                                 Row As Long, _
                                 resetColor As Boolean) As String
        
        On Error GoTo Fehler
    
100     Select Case Index

            Case 0
                 
105         Case 1

110             fpSpread1(Index).Row = fpSpread1(Index).ActiveRow
115             fpSpread1(Index).Col = fpSpread1(Index).ActiveCol
                
120             If resetColor = False Then
                   
125                 Select Case CStr(Col)
                 
                        Case COL_ARTIKEL, COL_MENGE, COL_EINHEIT, COL_EPREIS, COL_RABATT, COL_GPREIS, COL_UST, COL_LS_DATUM, COL_LS_NUMMER
                    
130                         fpSpread1(Index).Col = Col
135                         fpSpread1(Index).Row = Row
140                         fpSpread1(Index).BackColor = &HC0E0FF
                        
145                         fpSpread1(Index).CellTag = "Colored"

150                     Case COL_ZEILENART, COL_ARTSCHL, COL_DURCHLAUFEND, COL_KOSTSCHL, COL_SACHSCHL, COL_KOSTKTO, COL_SACHKTO

155                         fpSpread1(Index).Col = Col
160                         fpSpread1(Index).Row = Row
165                         fpSpread1(Index).BackColor = &HC0E0FF

170                         fpSpread1(Index).CellTag = "ColoredDeep"

                    End Select
                
                Else
                    
175                 Select Case fpSpread1(Index).CellTag
                
                        Case "Colored"
180                         fpSpread1(Index).BackColor = vbWhite
185                         fpSpread1(Index).CellTag = ""

190                     Case "ColoredDeep"
195                         fpSpread1(Index).BackColor = RGB(235, 229, 217)
200                         fpSpread1(Index).CellTag = ""

                    End Select
                    
                End If
                
205         Case 2
                
        End Select

        Exit Function
    
Fehler:
210     Call FehlerErklärung("frmSP52831", "SetCellBackColor()")

End Function

Private Sub SetActiveCellExt(Index As Integer, _
                             Col As Long, _
                             Row As Long, _
                             blnSetActiveColor As Boolean, _
                             Optional blnRuck As Boolean = False)

        On Error GoTo Fehler
        
        Dim i As Long
    
100     fpSpread1(Index).SetActiveCell Col, Row
    
105     If blnSetActiveColor Then
            
110         If blnRuck Then
            
115             For i = fpSpread1(Index).ActiveCol To fpSpread1(Index).MaxCols Step -1
               
120                 If fpSpread1(Index).IsVisible(fpSpread1(Index).ActiveCol, fpSpread1(Index).ActiveRow, False) = True Then
               
125                     Call SetCellBackColor(Index, fpSpread1(Index).ActiveCol, fpSpread1(Index).ActiveRow, False)
                  
                        Exit For

                    End If
               
130                 fpSpread1(Index).SetActiveCell fpSpread1(Index).ActiveCol - 1, fpSpread1(Index).ActiveRow
               
                Next
               
            Else
            
135             For i = fpSpread1(Index).ActiveCol To fpSpread1(Index).MaxCols
               
140                 If fpSpread1(Index).IsVisible(fpSpread1(Index).ActiveCol, fpSpread1(Index).ActiveRow, False) = True Then
               
145                     Call SetCellBackColor(Index, fpSpread1(Index).ActiveCol, fpSpread1(Index).ActiveRow, False)
                  
                        Exit For

                    End If
               
150                 fpSpread1(Index).SetActiveCell fpSpread1(Index).ActiveCol + 1, fpSpread1(Index).ActiveRow
               
                Next

            End If

        End If

        Exit Sub
    
Fehler:
155     Call FehlerErklärung("frmSP52831", "SetActiveCellExt()")

End Sub

Public Function IstArtieklVorhanden() As Boolean
        
        On Error GoTo Fehler

        Dim i      As Integer

        Dim knz    As Variant
        
        Dim result As Boolean
        
100     result = False
  
105     For i = 1 To LastRow
            
110         If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then
115             If knz = "A" Or knz = "P" Or knz = "L" Then
120                 result = True
125                 IstArtieklVorhanden = result

                    Exit Function

                End If
            End If

130     Next i
  
135     IstArtieklVorhanden = result

        Exit Function

Fehler:
140     Call FehlerErklärung("frmSP52831", "IstPorto")

End Function

Private Sub EnableButtonsOverPrintJob(blnEnabled As Boolean)
    
        On Error GoTo Fehler
        
        Dim vKey As Variant
        
100     If blnEnabled = False Then
            
105         dictButtonsState.RemoveAll
        
110         dictButtonsState.Add 1, cmd1(1).Enabled
115         dictButtonsState.Add 2, cmd1(2).Enabled
120         dictButtonsState.Add 3, cmd1(3).Enabled
125         dictButtonsState.Add 4, cmd1(4).Enabled
130         dictButtonsState.Add 5, cmd1(5).Enabled
135         dictButtonsState.Add 6, cmd1(6).Enabled
140         dictButtonsState.Add 10, cmd1(10).Enabled
145         dictButtonsState.Add 12, cmd1(12).Enabled

150         cmd1(1).Enabled = blnEnabled
155         mnu_Bearb_U(1).Enabled = blnEnabled

160         cmd1(3).Enabled = blnEnabled
165         mnu_Bearb_U(3).Enabled = blnEnabled

170         cmd1(2).Enabled = blnEnabled
175         mnu_Bearb_U(2).Enabled = blnEnabled

180         If GbDesigner Then

185             cmd1(12).Enabled = blnEnabled
190             mnu_Bearb_U(12).Enabled = blnEnabled

            End If
  
195         cmd1(5).Enabled = blnEnabled
200         mnu_Bearb_U(5).Enabled = blnEnabled

205         cmd1(4).Enabled = blnEnabled
210         mnu_Bearb_U(4).Enabled = blnEnabled

215         cmd1(10).Enabled = blnEnabled
220         mnu_Bearb_U(10).Enabled = blnEnabled

225         cmd1(6).Enabled = blnEnabled
230         mnu_Bearb_U(6).Enabled = blnEnabled
            
        Else
            
235         If Not dictButtonsState Is Nothing Then
            
240             For Each vKey In dictButtonsState.Keys
                
245                 cmd1(CInt(vKey)).Enabled = dictButtonsState.Item(vKey)
250                 mnu_Bearb_U(CInt(vKey)).Enabled = dictButtonsState.Item(vKey)
                
                Next
            
            End If
            
        End If

        Exit Sub
    
Fehler:
    
255     Me.MousePointer = vbDefault
260     Call FehlerErklärung("frmSP52831", "EnableButtonsOverPrintJob()")

End Sub

Public Function GetFormelText(strTExt As String, _
                              vEinheit As Variant, _
                              vMenge As Variant, _
                              lngRow As Long, _
                              ByRef blnIsFormelText As Boolean) As String
        
        On Error GoTo Fehler
        
        Dim i              As Integer
        
        Dim lngPos1        As Long
        
        Dim lngPos2        As Long
        
        Dim lngPos3        As Long
        
        Dim lngPos4        As Long
        
        Dim lngPosStart    As Long
        
        Dim strTemp        As String
        
        Dim strFormel      As String
        
        Dim strFormelOrig  As String
                
        Dim strResult      As String
        
        Dim strFx          As String
        
        Dim intFormelCount As Integer
        
        Dim blnBreak       As Boolean
        
        Dim intNotBreak    As Integer
        
        Dim dictFxBetrag   As New Dictionary
        
        Dim objKeys
        
        Dim objItems
    
100     strResult = strTExt
    
105     lngPos1 = InStr(1, strTExt, "%F1%")
        
110     If lngPos1 > 0 Then
            
115         blnIsFormelText = True
            
120         lngPosStart = lngPos1
            
125         strTExt = Replace(strTExt, "%" + CStr(vEinheit) + "%", CStr(vMenge))
            
130         Do While blnBreak = False
                
135             intNotBreak = intNotBreak + 1
                
140             If intNotBreak >= 100 Then blnBreak = True
                
145             intFormelCount = intFormelCount + 1
                
150             strFx = "%F" + CStr(intFormelCount) + "%"
                
155             lngPos1 = InStr(lngPosStart, strTExt, strFx)

160             If lngPos1 > 0 Then

165                 lngPos2 = InStr(lngPos1, strTExt, "{")
170                 lngPos3 = InStr(lngPos1, strTExt, "}")
175                 lngPos4 = InStr(lngPosStart, strTExt, "=")
                
180                 If lngPos1 < lngPos2 And lngPos2 < lngPos3 Then
                    
185                     strFormelOrig = Mid$(strTExt, lngPos2, lngPos3 - lngPos2 + 1)
190                     strTemp = Mid$(strTExt, lngPos2 + 1, lngPos3 - lngPos2 - 1)
                    
195                     strFormel = Replace$(strTemp, ",", ".")
200                     strFormel = Replace$(strFormel, vbCrLf, " ")
205                     strFormel = Replace$(strFormel, vbCr, " ")
210                     strFormel = Replace$(strFormel, vbNewLine, " ")
                    
215                     If dictFxBetrag.Count > 0 Then
                        
220                         objKeys = dictFxBetrag.Keys
225                         objItems = dictFxBetrag.Items
                    
230                         For i = 0 To dictFxBetrag.Count - 1
                        
235                             strFormel = Replace$(strFormel, objKeys(i), objItems(i))
                        
                            Next
                    
                        End If
                        
                        On Error GoTo Message
                    
240                     strFormel = ScriptControl.Eval(strFormel)
                    
245                     strFormel = Format(strFormel, ZahlFormat(postCommaPreis))
250                     strFormel = RundenMitVz(CDbl(strFormel), postCommaPreis)
                    
255                     If dictFxBetrag.Exists(strFx) = False Then dictFxBetrag.Add strFx, Replace$(strFormel, ",", ".")
                    
260                     strFormel = Format(strFormel, ZahlFormat(postCommaPreis))
                    
265                     strTExt = Replace$(strTExt, strFormelOrig, "", , 1)
                    
270                     If lngPos4 > 0 And lngPos4 < lngPos2 Then
                    
275                         strTExt = Replace$(strTExt, "=", "", , 1)
                        
                        End If
                    
280                     strTExt = Replace$(strTExt, strFx, strFormel, , 1)
                    
                    End If
                    
                Else
                
285                 blnBreak = True
                    
                End If

            Loop
        
290         If blnIsFormelText Then Call UpdateFormelTextRows(lngRow, 0)
            
        End If
        
295     strResult = strTExt
        
300     GetFormelText = strResult

        Exit Function

Message:

305     GetFormelText = strResult
        
310     MsgBox GetMessage(2338), vbOKOnly + vbCritical, strMeldungCap

        Exit Function
        
Fehler:

315     Call FehlerErklärung("frmSP52831", "GetFormelText()")

End Function

Private Sub UpdateFormelTextRows(lngArtRowIndex As Long, lngFormelTextRowsCount As Long)
    
        On Error GoTo Fehler
        
100     If dictFormelRows.Exists(lngArtRowIndex) = False Then

105         dictFormelRows.Add lngArtRowIndex, lngFormelTextRowsCount
                
        Else
            
110         dictFormelRows.Item(lngArtRowIndex) = lngFormelTextRowsCount
            
        End If

        Exit Sub
    
Fehler:
    
115     Me.MousePointer = vbDefault
120     Call FehlerErklärung("frmSP52831", "UpdateFormelTextRows()")

End Sub

Private Sub SetUst()

100     If frmParent.txt1(20).text = 0 Then

105         dblUstSatz = GetWaehrung(frmParent.txt1(17).text, False).MwSt

        Else
        
110         dblUstSatz = frmParent.txt1(20).text

        End If

End Sub

Public Function GetfpSpread1Value() As Integer

100     GetfpSpread1Value = fpSpread1(2).GetText(3, 2, vValue)

End Function

Public Function IsEinheitOK() As Boolean
    
        On Error GoTo Fehler
    
        Dim i   As Integer

        Dim knz As Variant
        
100     If gEnmKudnenERechnungType = eERechnungType.None Then
        
105         IsEinheitOK = True

            Exit Function
            
        End If
        
110     For i = 1 To LastRow

115         If fpSpread1(1).GetText(COL_ZEILENART, i, knz) Then

120             If knz = "A" Or knz = "P" Or knz = "L" Then

125                 If fpSpread1(1).GetText(COL_EINHEIT, i, knz) Then

130                     If Trim(knz) = "" Then

135                         MsgBox GetMessage(2194), vbExclamation, strMeldungCap
140                         Call SetActiveCellExt(1, COL_EINHEIT, CLng(i), True)

                            Exit Function

                        End If

                    Else
                    
145                     MsgBox GetMessage(2194), vbExclamation, strMeldungCap

150                     DoEvents

155                     Call SetActiveCellExt(1, COL_EINHEIT, CLng(i), True)

                        Exit Function

                    End If
                    
                End If
                
            End If

160     Next i
        
165     IsEinheitOK = True

        Exit Function

Fehler:
170     Call FehlerErklärung("frmSP56431", "IstArtieklOK")

End Function

Public Sub KontrollierenZellgruppierung(Row As Long, _
                                        startCol As Long, _
                                        boolZustand As Boolean)

        On Error GoTo Fehler

        Dim i As Integer

100     fpSpread1(1).Row = Row

105     If Not boolZustand Then

110         fpSpread1(1).RemoveCellSpan COL_ARTIKEL, Row

115         fpSpread1(1).Col = COL_EINHEIT

120         fpSpread1(1).CellType = CellTypeComboBox
125         fpSpread1(1).Lock = False

        Else

130         For i = 4 To 9

135             fpSpread1(1).Col = i
140             fpSpread1(1).Lock = False

145         Next i

150         fpSpread1(1).Col = COL_EINHEIT

155         fpSpread1(1).CellType = CellTypeStaticText

160         fpSpread1(1).AddCellSpan COL_ARTIKEL, Row, 7, 1

        End If
        
165     fpSpread1(1).Col = startCol

        Exit Sub

Fehler:
170     Me.MousePointer = vbDefault
175     Call FehlerErklärung("frmSP52831", "KontrollierenZellgruppierung()")

End Sub

Public Sub Speichern(boolNew As Boolean)

        On Error GoTo Fehler

100     If LastRow > 0 Then

105         If Plausi Then

110             If EinheitsPrufungMitFokus Then
                   
115                 If Trim(frmParent.BelegNr) <> "" And Trim(frmParent.BelegNr) <> "0" Then

120                     If Trim(frmParent.belegDatum) = "" Then frmParent.belegDatum = CStr(Date)

                    End If

125                 If frmParent.glngBelegID = 0 Or boolNew Then

130                     frmParent.Speichern

                    Else
                        
135                     frmParent.Speichern frmParent.glngBelegID

                    End If
                
140                 Call FillBelegInfo(True, frmParent.glngBelegID, frmParent.BelegNr, GintBelegArt, 0)
                        
145                 If cmd1(8).Enabled Then cmd1(8).SetFocus
150                 If cmd1(1).Enabled = False Then cmd1(1).Enabled = True
155                 If mnu_Bearb_U(1).Enabled = False Then mnu_Bearb_U(1).Enabled = True

160                 BelegNeu = False
                
165                 If frmParent.Check1(0).value = 1 Then
                    
170                     MsgBox GetMessage(2188), vbExclamation, strMeldungCap

                    Else
                    
175                     MsgBox GetMessage(2189), vbExclamation, strMeldungCap

                    End If

180                 Call MaskeLeeren(alleDaten)

185                 Unload Me

                End If
            
            End If
                
        Else
                
190         MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gespeichert werden.", vbInformation, strMeldungCap

195         If fpSpread1(1).Enabled Then fpSpread1(1).SetFocus
                    
        End If

        Exit Sub

Fehler:

200     Me.MousePointer = vbDefault
205     Call FehlerErklärung("frmSP52831", "Speichern()")

End Sub

Public Sub Druck(Optional bAblageOhneDruck As Boolean)
                
        Dim lngPrintRet As Long
        
        On Error GoTo Fehler
        
100     lngPrintRet = 999
    
105     If FileExists(ArbeitsplatzPfad & "\SP52800.LL") Then
                    
110         Call DeleteFile(ArbeitsplatzPfad & "\SP52800.LL")
                    
        End If
    
115     Select Case GintBelegArt
                
            Case 0
                    
120             objDruckOptionen.FormularNr = 35
                    
125         Case 1
                    
130             objDruckOptionen.FormularNr = 36
                    
135         Case 2
                    
140             objDruckOptionen.FormularNr = 120
                    
145         Case 3
                    
150             objDruckOptionen.FormularNr = 121
                
        End Select
                
155     objDruckOptionen.CurrentSteuerValue = GetRNGGUTSteuerTextLkz(GintBelegArt, intSteuerTyp)
                
160     If LL18CheckBildFile(CStr(objDruckOptionen.FormularNr)) = False Then

            Exit Sub
                    
        End If

165     If frmParent.belegDatum <> "" Then
170         objDruckOptionen.CurrentBelegDatum = frmParent.belegDatum
                    
        End If

175     If frmParent.ValutaDatum <> "" Then
180         objDruckOptionen.CurrentValutaDatum = frmParent.ValutaDatum
                    
        End If

185     If frmParent.BelegNr <> "" Then

190         If val(frmParent.BelegNr) > 0 Then
195             objDruckOptionen.EnableBelegNr = False
200             objDruckOptionen.CurrentBelegNr = frmParent.BelegNr
            Else
205             objDruckOptionen.EnableBelegNr = True
210             objDruckOptionen.CurrentBelegNr = ""
                        
            End If
                 
        Else

215         If frmParent.Check1(0) = 1 Then
220             objDruckOptionen.EnableBelegNr = False
            Else
225             objDruckOptionen.EnableBelegNr = True
230             objDruckOptionen.CurrentBelegNr = ""
                        
            End If
                    
        End If
  
235     objDruckOptionen.resizeFactor = frmParent.getResizeFactor
 
240     Call objDruckOptionen.ShowMe(frmParent.lblDruckOption.caption, Me)

245     If objDruckOptionen.Canceled Then

250         LL1.LlPrintEnd (0)

            Exit Sub
                    
        End If
                                
255     frmParent.belegDatum = objDruckOptionen.CurrentBelegDatum
                      
260     frmParent.BelegNr = objDruckOptionen.CurrentBelegNr

265     frmParent.ValutaDatum = objDruckOptionen.CurrentValutaDatum

270     Call FussZahlungsZiel
                
275     If objDruckOptionen.CurrentSteuertext = "Automatisch" Then
280         gstrSteuerText = ""
                 
        Else
285         gstrSteuerText = objDruckOptionen.CurrentSteuertext
                    
        End If

290     If LastRow > 0 Then
                    
295         printJobInProgress = True
                    
300         Call EnableButtonsOverPrintJob(False)
                    
305         If Plausi Then

310             If EinheitsPrufungMitFokus Then
  
315                 If frmParent.glngBelegID = 0 Then

320                     If frmParent.Speichern(, , True) Then

325                         If objDruckOptionen.CurrentBelegNr <> frmParent.BelegNr And val(Trim(frmParent.BelegNr)) <> 0 Then

330                             objDruckOptionen.CurrentBelegNr = frmParent.BelegNr

                            End If
                        
335                         If bAblageOhneDruck Then
                                                            
340                             lngPrintRet = LLPrintListe(Me, LL1, frmParent.glngBelegID, 4)
                            Else
                                
345                             lngPrintRet = LLPrintListe(Me, LL1, frmParent.glngBelegID, 1)
                            End If
                                
350                         Select Case lngPrintRet
                                
                                Case 0
                                        
355                                 fpSpread1(0).SetText 2, 17, SP52800B.gLngBelegNr
                                                                 
360                                 LLPrintListe Me, LL1, frmParent.glngBelegID, 3, False, True
                                        
365                                 If SP52800B.blnMaskeLeeren Then

370                                     Call MaskeLeeren(alleDaten)

                                    End If
                                        
375                             Case LL_ERR_USER_ABORTED
            
380                                 Call MaskeLeeren(nurBelegDaten)
                                        
385                             Case Else
                                    
390                                 frmParent.gintDruck = 0
                                        
                            End Select
                                
                        End If

                    Else
                            
                        Dim lngTempBelegId As Long
                            
395                     lngTempBelegId = CLng(frmParent.glngBelegID)
                                                        
400                     If BelegInfo.Art <> GintBelegArt Or BelegInfo.Druck = 1 Then lngTempBelegId = 0
                                                        
405                     If frmParent.Speichern(lngTempBelegId, , True) Then

410                         If objDruckOptionen.CurrentBelegNr <> frmParent.BelegNr And val(Trim(frmParent.BelegNr)) <> 0 Then

415                             objDruckOptionen.CurrentBelegNr = frmParent.BelegNr

                            End If
                        
420                         If bAblageOhneDruck Then
                                                            
425                             lngPrintRet = LLPrintListe(Me, LL1, frmParent.glngBelegID, 4)

                            Else
                                
430                             lngPrintRet = LLPrintListe(Me, LL1, frmParent.glngBelegID, 1)

                            End If
                                
435                         Select Case lngPrintRet
                                
                                Case 0
                                
440                                 fpSpread1(0).SetText 2, 17, SP52800B.gLngBelegNr

445                                 LLPrintListe Me, LL1, frmParent.glngBelegID, 3, False, True
                                        
450                                 If SP52800B.blnMaskeLeeren Then

455                                     Call MaskeLeeren(alleDaten)

                                    End If
                                            
460                             Case LL_ERR_USER_ABORTED
            
465                                 Call MaskeLeeren(nurBelegDaten)
                                        
470                             Case Else
                                    
475                                 frmParent.gintDruck = 0
                                        
                            End Select
                                        
                        Else
                                
480                         frmParent.gintDruck = 0
                                    
                        End If
                                    
                    End If
                            
                End If
                        
            End If
                    
485         printJobInProgress = False
                    
490         Call EnableButtonsOverPrintJob(True)
                    
495         If lngPrintRet = 0 Then

500             If frmParent.Check1(0).value = 1 Then

505                 Me.SetFocus
510                 MsgBox GetMessage(2190), vbExclamation + vbApplicationModal, strMeldungCap

                Else
                
515                 Me.SetFocus
520                 MsgBox GetMessage(2191), vbInformation + vbApplicationModal, strMeldungCap

                End If

525             Call MaskeLeeren(alleDaten)

530             Unload Me
            
            End If
                 
        Else
                
535         MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gedruckt werden.", vbInformation, strMeldungCap

540         If fpSpread1(1).Enabled Then fpSpread1(1).SetFocus
                    
        End If

        Exit Sub
        
Fehler:
        
545     Me.MousePointer = vbDefault
550     Call FehlerErklärung("frmSP52831", "Druck()")

End Sub

Public Sub Vorschau()

        On Error GoTo Fehler
                
100     If FileExists(ArbeitsplatzPfad & "\SP52800.LL") Then
                    
105         Call DeleteFile(ArbeitsplatzPfad & "\SP52800.LL")
                    
        End If
    
110     Select Case GintBelegArt

            Case 0
                    
115             objDruckOptionen.FormularNr = 35
                    
120         Case 1
                    
125             objDruckOptionen.FormularNr = 36
                    
130         Case 2
                    
135             objDruckOptionen.FormularNr = 120

140         Case 3
                    
145             objDruckOptionen.FormularNr = 121
                
        End Select
                
150     objDruckOptionen.CurrentSteuerValue = GetRNGGUTSteuerTextLkz(GintBelegArt, intSteuerTyp)
                                                                   
155     If LL18CheckBildFile(CStr(objDruckOptionen.FormularNr)) = False Then

            Exit Sub
                    
        End If

160     If frmParent.belegDatum <> "" Then
165         objDruckOptionen.CurrentBelegDatum = frmParent.belegDatum
                 
        Else
170         objDruckOptionen.CurrentBelegDatum = Format(Date, "dd.mm.yyyy")
                    
        End If

175     If frmParent.ValutaDatum <> "" Then
180         objDruckOptionen.CurrentValutaDatum = frmParent.ValutaDatum
                 
        Else
185         objDruckOptionen.CurrentValutaDatum = ""
                    
        End If

190     If frmParent.BelegNr <> "" Then
195         objDruckOptionen.CurrentBelegNr = frmParent.BelegNr
                 
        Else
200         objDruckOptionen.CurrentBelegNr = "0"
                    
        End If
                
205     objDruckOptionen.EnableBelegNr = False

210     objDruckOptionen.resizeFactor = frmParent.getResizeFactor

215     Call objDruckOptionen.ShowMe(frmParent.lblDruckOption.caption, Me)

220     If objDruckOptionen.Canceled Then

225         LL1.LlPrintEnd (0)

            Exit Sub
                    
        End If

230     frmParent.ValutaDatum = objDruckOptionen.CurrentValutaDatum

235     If GesamtIstBrutto Then
240         Call FussZahlungsZielGesamtIstBrutto
                 
        Else
245         Call FussZahlungsZiel
                    
        End If

250     frmParent.belegDatum = objDruckOptionen.CurrentBelegDatum

255     If objDruckOptionen.CurrentSteuertext = "Automatisch" Then
260         gstrSteuerText = ""
                 
        Else
265         gstrSteuerText = objDruckOptionen.CurrentSteuertext
                    
        End If

270     If LastRow > 0 Then
            
275         If Plausi Then

280             If EinheitsPrufungMitFokus Then
                    
285                 If frmParent.glngBelegID = 0 Then
                        
290                     If frmParent.Speichern(frmParent.glngBelegIDTmp, True) Then
                            
295                         LLPrintListe Me, LL1, frmParent.glngBelegIDTmp, 2, True

                        End If

                    Else
                            
300                     If frmParent.Speichern(frmParent.glngBelegID, True) Then
                                                                
305                         LLPrintListe Me, LL1, frmParent.glngBelegID, 2, True

                        End If
                            
                    End If
                    
                End If
                
            End If
                 
        Else
                
310         MsgBox "Es sind keine Posten erfasst. Der Beleg kann nicht gedruckt werden.", vbInformation, strMeldungCap

315         If fpSpread1(1).Enabled Then fpSpread1(1).SetFocus
                    
        End If

        Exit Sub

Fehler:
320     Me.MousePointer = vbDefault
325     Call FehlerErklärung("frmSP52831", "Vorschau()")

End Sub

Public Sub TabellenAktualisieren(Optional lngNewBelegId As Long, _
                                 Optional Zeigen As Boolean = True)
    
        On Error GoTo Fehler
        
100     Call KopfFuellen

105     If GesamtIstBrutto Then
110         FussSetUpGesamtIstBrutto
115         FussFuellenGesamtIstBrutto
        Else
120         FussSetUp
125         FussFuellen
        End If
        
130     If lngNewBelegId <> 0 Then FolgeZeigen (lngNewBelegId)

135     If Zeigen Then Me.Show 0

140     SteuernKontrol True

        Exit Sub

Fehler:
145     Me.MousePointer = vbDefault
150     Call FehlerErklärung("frmSP52831", "TabellenAktualisieren()")

End Sub

Public Sub SteuernKontrol(Optional Meldung As Boolean)
    
        On Error GoTo Fehler
        
        Dim i As Integer
                
100     If Meldung Then
        
105         If intSteuerTyp <> intOldSteuerTyp Then

110             Select Case intOldSteuerTyp

                    Case 0
    
115                     Select Case intSteuerTyp
    
                            Case 1
                            
120                             Call MsgBox(GetMessage(2379), vbOKOnly + vbExclamation, strMeldungCap)
        
125                         Case 2

130                             Call MsgBox(GetMessage(2380), vbOKOnly + vbExclamation, strMeldungCap)
    
                        End Select
    
135                 Case 1
    
140                     Select Case intSteuerTyp
    
                            Case 0
                            
145                             Call MsgBox(GetMessage(2381), vbOKOnly + vbExclamation, strMeldungCap)
        
150                         Case 2

155                             Call MsgBox(GetMessage(2382), vbOKOnly + vbExclamation, strMeldungCap)
    
                        End Select
    
160                 Case 2
    
165                     Select Case intSteuerTyp
    
                            Case 0
                            
170                             Call MsgBox(GetMessage(2383), vbOKOnly + vbExclamation, strMeldungCap)
        
175                         Case 1

180                             Call MsgBox(GetMessage(2384), vbOKOnly + vbExclamation, strMeldungCap)
    
                        End Select

                End Select
                
            End If
        
        End If
        
185     For i = 1 To LastRow
 
190         If IstZeilenTyp(i, "A") Or IstZeilenTyp(i, "P") Or IstZeilenTyp(i, "L") Then

195             fpSpread1(1).Col = COL_UST

200             Select Case intSteuerTyp

                    Case 0

205                     If Meldung Then fpSpread1(1).SetText COL_UST, i, 0
210                     fpSpread1(1).Lock = False

215                 Case 1

220                     If Meldung Then fpSpread1(1).SetText COL_UST, i, 1
225                     fpSpread1(1).Lock = False

230                 Case 2

235                     fpSpread1(1).SetText COL_UST, i, 0
240                     fpSpread1(1).Lock = True

                End Select

            End If

245     Next i

        Exit Sub

Fehler:
250     Me.MousePointer = vbDefault
255     Call FehlerErklärung("frmSP52831", "SteuernKontrol()")

End Sub

Public Sub InputProcessing(strInput As String)

        On Error GoTo Fehler

100     objERechnung.XMLResult = strInput

        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52831", "InputProcessing()")

End Sub

