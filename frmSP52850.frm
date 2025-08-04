VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{2213E283-16BC-101D-AFD4-040224009C1D}#29.0#0"; "CMLL29O.OCX"
Begin VB.Form frmSP52850 
   Caption         =   "Sammeldruck"
   ClientHeight    =   5715
   ClientLeft      =   18705
   ClientTop       =   6045
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSP52850.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10470
   Begin ListLabel.ListLabel LL1 
      Left            =   7890
      Top             =   1470
      _Version        =   65537
      _ExtentX        =   820
      _ExtentY        =   820
      _StockProps     =   64
      Language        =   0
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
      PreviewRectLeft =   -1
      PreviewRectTop  =   -1
      PreviewRectWidth=   1
      PreviewRectHeight=   1
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
   Begin VB.Frame Frame1 
      Caption         =   "Auswahl"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Index           =   1
      Left            =   30
      TabIndex        =   12
      Top             =   960
      Width           =   10425
      Begin TrueOleDBGrid70.TDBGrid TDBG1 
         Height          =   3555
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   6271
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Microsoft Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Microsoft Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         Appearance      =   0
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   -2147483638
         ScrollTrack     =   -1  'True
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=Microsoft Sans Serif "
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Microsoft Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Microsoft Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000007&,.bold=0"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Microsoft Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   5010
      Width           =   10500
      _Version        =   65536
      _ExtentX        =   18521
      _ExtentY        =   635
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmd1 
         Caption         =   "&Alle"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Keine"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   4
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&peichern"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1706
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Ö&ffnen"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   3016
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Drucken"
         Height          =   300
         Index           =   5
         Left            =   7230
         TabIndex        =   8
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&chließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   9060
         TabIndex        =   7
         Top             =   30
         Width           =   1260
      End
   End
   Begin MSComctlLib.StatusBar sta1 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   5370
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1376
            MinWidth        =   1376
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1376
            MinWidth        =   1376
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15134
            MinWidth        =   12541
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSP52850.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSP52850.frx":0964
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Abgrenzung"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   60
      Width           =   10425
      Begin VB.CommandButton cmdAuswahl 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   7650
         Picture         =   "frmSP52850.frx":0E86
         Style           =   1  'Grafisch
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3350
         Picture         =   "frmSP52850.frx":0F68
         Style           =   1  'Grafisch
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   0
         Left            =   2280
         TabIndex        =   0
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   1
         Left            =   6570
         TabIndex        =   1
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Ffunclbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   1
         Left            =   10000
         TabIndex        =   20
         ToolTipText     =   "F2 Funktion"
         Top             =   15
         Width           =   255
      End
      Begin VB.Label Ffunclbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   135
         Index           =   0
         Left            =   9785
         TabIndex        =   21
         ToolTipText     =   "F1 Funktion"
         Top             =   15
         Width           =   255
      End
      Begin VB.Shape FfuncShp 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Undurchsichtig
         Height          =   165
         Index           =   0
         Left            =   9750
         Top             =   0
         Width           =   195
      End
      Begin VB.Shape FfuncShp 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Undurchsichtig
         Height          =   165
         Index           =   2
         Left            =   9980
         Top             =   0
         Width           =   195
      End
      Begin VB.Label lblDruckOptionen 
         Caption         =   "SPDruckOptionen Caption"
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblDummy 
         Caption         =   "Rechnung"
         Height          =   210
         Index           =   0
         Left            =   8160
         TabIndex        =   18
         Top             =   210
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblDummy 
         Caption         =   "Gutschrift"
         Height          =   210
         Index           =   1
         Left            =   8160
         TabIndex        =   17
         Top             =   420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbl1 
         Caption         =   "Von Erfassungs-Datum"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lbl1 
         Caption         =   "Bis Erfassungs-Datum"
         Height          =   210
         Index           =   1
         Left            =   4320
         TabIndex        =   13
         Top             =   300
         Width           =   1815
      End
   End
   Begin VB.Menu mnuDat 
      Caption         =   "Datei"
      Index           =   0
      Begin VB.Menu mnuDat1 
         Caption         =   "Schließen"
         Index           =   6
      End
   End
   Begin VB.Menu mnuBearb 
      Caption         =   "Bearbeiten"
      Index           =   0
      Begin VB.Menu mnuBearb1 
         Caption         =   "Vorschau"
         Index           =   0
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "&Ausgabe"
         Index           =   1
         Begin VB.Menu mnuAusgabe 
            Caption         =   "Vorschau"
            Index           =   0
         End
         Begin VB.Menu mnuAusgabe 
            Caption         =   "Ablage"
            Index           =   1
         End
         Begin VB.Menu mnuAusgabe 
            Caption         =   "Druck"
            Index           =   2
         End
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Alle auswählen"
         Index           =   2
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Keine Auswählen"
         Index           =   3
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Designer"
         Enabled         =   0   'False
         Index           =   5
         Begin VB.Menu mnuDesign 
            Caption         =   "alle Mandanten"
            Index           =   0
         End
         Begin VB.Menu mnuDesign 
            Caption         =   "aktueller Mandant"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuAnsicht 
      Caption         =   "&Ansicht"
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
   Begin VB.Menu mnuOpt 
      Caption         =   "Optionen"
      Index           =   0
      Begin VB.Menu mnuOpt1 
         Caption         =   "Druckerauswahldialog"
         Checked         =   -1  'True
         Index           =   0
      End
   End
   Begin VB.Menu mnuInfoMain 
      Caption         =   "*?"
      Begin VB.Menu mnuInfo 
         Caption         =   "*Programmbeschreibung"
         Index           =   0
      End
      Begin VB.Menu mnuUpdateInfo 
         Caption         =   "*Update-Info"
         Index           =   1
      End
   End
   Begin VB.Menu mnuDummy 
      Caption         =   "Dummy"
      Index           =   0
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSP52850"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public gintPrivBelegArt  As Integer

Private objPRM           As clsPRM

Private objTDBG          As clsTDBG7ole

Private objHlp           As SpHlp.clsHlp

'Private Datumausw        As New SPDatumAusw.clsDatumAusw
  
Private WithEvents gRS   As ADODB.Recordset
Attribute gRS.VB_VarHelpID = -1

Private Col              As TrueOleDBGrid70.Column

Private cols             As TrueOleDBGrid70.Columns

Private gIntSortColIndex As Integer

Private gstrSortOrder    As String

'LL

Private glRet            As Long
  
Private Const VK_LEFT = &H25  'Linke Pfeiltaste

Private Const VK_UP = &H26    'Obere Pfeiltaste

Private Const VK_RIGHT = &H27 'Rechte Pfeiltaste

Private Const VK_DOWN = &H28  'Untere Pfeiltaste
  
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

Private allowFocusLostEvent As Boolean

Private showDruckOptionen   As Boolean
  
'####### Subclassing ########################
'DeW, Mai 2011
'Variablen notwendig fuer Verwendung der SSubTmr Klasse,
'um eine schoenes Vergroesserung von Fenster und Inhalt und
'Begrenzung der Fenstergroesse zu ermoeglichen!
Implements ISubclass

Private emrConsume   As EMsgResponse
'
'DeW, notwendige WM_... Nachrichten fuer das
'Subclassing wurden als Public in SP50000B.bas
'definiert
'############################################

'####### Formular Resizing ##################
'
Dim cReSize          As FormResize 'HW 03.02.2011

'<Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
Private shiftPressed As Boolean
'</Added by: GW at: 11.02.2019, Ver.: 6.5.109 >

Dim strEBelegNone           As String                                          'E-Beleg-Spalte Bezeichnungen

Dim strEBelegZUGFeRDE       As String

Dim strEBelegZUGFeRDS       As String

'
'############################################

'DH, 18.10.2011, Workaround fuer das Problem, dass die Pfeiltasten immer auf
'den naechsten TabStop springen anstatt auf den Button daneben.
'Methode reagiert jetzt auch selber auf Buttons mit Enable = False. Einzig das Array buttons()
'und der Name der Control muss selbst angepasst werden.
Private Sub cmd1_LostFocus(Index As Integer)

        '4 - 5 - 0 - 1 - 6
        Dim buttons(5) As Integer

        Dim pos        As Integer
  
        'Button-Reihenfolge anhand der Indizes von Links nach Rechts
        '100     buttons(0) = 4                                                'IL 29.11.2024 , Ver.: 6.7.101 : Die Vorschaufunktion wurde in die Dropdown-Liste verschoben
100     buttons(1) = 5
105     buttons(2) = 0
110     buttons(3) = 1
115     buttons(4) = 6
  
        'Indexposition im Array herausfinden - pos = Position des Buttons der den Fokus verliert in buttons()
120     For pos = 0 To UBound(buttons) - 1

125         If buttons(pos) = Index Then Exit For
        Next
    
130     If allowFocusLostEvent Then
135         allowFocusLostEvent = False
    
            'Pfeiltaste Links
140         If GetAsyncKeyState(VK_LEFT) Then

                Do

145                 If pos - 1 < 0 Then                    'Ist man ganz links im Array angekommen...
150                     pos = UBound(buttons) - 1            '...dann ist der naechste Index der Wert ganz rechts im Array
                    Else
155                     pos = pos - 1                        'Ansonsten ist der Naechste Index, ein Index weiter links
                    End If

160             Loop While cmd1(buttons(pos)).Enabled = False

165             cmd1(buttons(pos)).SetFocus
170             DoEvents
            End If
    
            'Pfeiltaste Rechts
175         If GetAsyncKeyState(VK_RIGHT) Then

                Do

180                 If pos + 1 > UBound(buttons) - 1 Then  'Ist man ganz rechts im Array angekommen...
185                     pos = 0                              '...dann ist der naechste Index der Wert ganz links im Array
                    Else
190                     pos = pos + 1                        'Ansonsten ist der Naechste Index, ein Index weiter links
                    End If

195             Loop While cmd1(buttons(pos)).Enabled = False

200             cmd1(buttons(pos)).SetFocus
205             DoEvents
            End If
    
210         allowFocusLostEvent = True

        End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

        '<Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
100     If Shift <> 1 Then
105         shiftPressed = False
        End If

        '</Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
End Sub

Private Sub Form_Load()

        On Error GoTo Fehler
        
100     shiftPressed = False                                                    'Added by: GW at: 11.02.2019, Ver.: 6.5.109

105     If GintBelegArt = 0 Then

110         SaveSetting "SP50000", App.EXEName, "SP62850_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

        Else
        
115         SaveSetting "SP50000", App.EXEName, "SP62860_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

        End If

        '########### SkinFramework ##############################
        'HW 04.03.2011 - Is ein Windows Skin Tool von Codejock Software
        'SkinFramework1.LoadSkin GsHauptPfad & "\exe\Spedifix.cjstyles", "NormalSilver.ini"
        'SkinFramework1.ApplyWindow Me.hwnd
        '########################################################
    
        '########## Subclassing: Messages festlegen #############
        ' DeW, ZyG Mai 2011
120     AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO 'DeW
125     AttachMessage Me, Me.hwnd, WM_EXITSIZEMOVE 'DeW
        '
        '########################################################
    
        '####### Subclassing: Groessenbegrenzung Formular #######
        ' TODO in Arbeit, MagicNumbers
        'dieser Aufruf kann je nach Programm-Modul woanders in der
        'Form_Load Methode stehen!
        'Zuerst muss der "alte" Code die Zuweisung von Breit und
        'Hoehe korrekt vorgenommen haben!
130     SetMinMaxInfo Me.hwnd, Me.height, (Me.height * 2), Me.width, (Me.width * 2)
        '
        '########################################################

        'DH, 08.05.2013, 6.1.121, Umstellung auf List & Label 18
135     LL1.LlSetOptionString LL_OPTIONSTR_LICENSINGINFO, "4yi/HQ"              'HW 29.10.2014
140     LL1.LlSetOption LL_OPTION_INCREMENTAL_PREVIEW, False
145     LL1.LlSetOption LL_OPTION_RIBBON_DEFAULT_ENABLEDSTATE, 0
150     LL1.LlSetOption LL_OPTION_INCLUDEFONTDESCENT, False

155     LL1.LlSetOption LL_OPTION_CONVERTCRLF, True                             'Verhindern der doppelten Zeilenumbrüche.
160     LL1.LlSetPrinterDefaultsDir ArbeitsplatzPfad                            'Pfad für Drucker-Einstellungen setzen.
165     LL1.LlSetOption LL_OPTION_NOPARAMETERCHECK, 1                           'Nach den Tests LL-Parameterüberprüfung ausschalten (Geschwindigkeitvorteil)
170     LL1.LlPreviewSetTempPath ArbeitsplatzPfad                               'Pfad für Vorschaudateien setzen.
  
175     gintPrivBelegArt = GintBelegArt
  
180     Set objPRM = New clsPRM
185     Set objPRM.gForm = Me
190     objPRM.PRM_Alle
        
195     Call FillPRMValues                                                      'DF 16.12.2024 , Ver.: 6.7.103
        
200     strMeldungCap = mnuDummy(0).caption

205     Me.caption = Me.caption & "-" & lblDummy(gintPrivBelegArt)

210     Me.width = 10000                                                        'DH, 16.09.2011, in Folge der Button-Neuausrichtung verbreitert
215     Me.height = 5000 'FORM_HEIGHT
  
220     SetXPSize Me

225     Call setSkinnerBackColor(sta1)
  
230     sta1.Panels(1).text = "SP62850"

235     sta1.Panels(2).text = DisplayVerInfo(GsHauptPfadLokal & "exe\" & Gc_strExeFile)

240     Set objHlp = New SpHlp.clsHlp
245     objHlp.DatabaseName = GsHauptPfadLokal & "hlp\SP50000.hlp"
250     objHlp.table = Me.name
255     objHlp.caption = Me.name & " - Feldhilfe"

260     Set objTDBG = New clsTDBG7ole
265     Set objTDBG.TDBG = TDBG1
270     Set objTDBG.PrmDataBase = GDBprm
275     Set objERechnung = New clsERechnung                                     'DF 23.08.2024 , Ver.: 6.7.101 : Evtl. erst instanzieren wenn der Kunde gezogen wird.

280     Call AusDBLesen
  
285     objTDBG.SichtbareZeilen = 14
290     objTDBG.SatzAnzahl = gRS.RecordCount
295     objTDBG.GridParameter ("SELECT * FROM PRM52850 WHERE name = 'TDBG1' ORDER BY index")
300     TDBG1.width = 10250
305     TDBG1.height = 3600
310     TDBG1.HoldFields

315     TDBG1.MarqueeStyle = dbgHighlightCell
320     TDBG1.HighlightRowStyle.BackColor = vbWindowBackground                  '&H80000005
325     TDBG1.HighlightRowStyle.ForeColor = vbWindowText                        'Farbe des Textes in Fenstern
330     TDBG1.AllowColSelect = False
335     TDBG1.ExtendRightColumn = False                                         'Die äußere rechte Spalte wird ans rechte Gridende nicht erweitert.
340     TDBG1.FilterBar = True

345     Set cols = TDBG1.Columns
350     cols("ERechnungArtGK").NumberFormat = "FormatText Event"                'DF 16.12.2024 , Ver.: 6.7.103
        
355     Set objDruckOptionen = New SPDruckOptionen.clsDruckOptionen             'DH, 08.10.2013, 6.2.100, DruckOptionen Fenster instanzieren und konfigurieren

        '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
360     Select Case GintBelegArt

            Case 0
    
365             objDruckOptionen.FormularNr = 35
    
370         Case 1
    
375             objDruckOptionen.FormularNr = 36
    
380         Case 2
    
385             objDruckOptionen.FormularNr = 120
    
390         Case 3
    
395             objDruckOptionen.FormularNr = 121

        End Select

        '345     If GintBelegArt = 0 Then
        '350         objDruckOptionen.FormularNr = 35
        '        Else
        '355         objDruckOptionen.FormularNr = 36
        '        End If
        '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
        
400     objDruckOptionen.EnableValutaDatum = True
405     mnuOpt1(0).Checked = GetSetting("SP50000", "SP52800", "SP52850DruckerDialog", "-1")
  
410     BearbeiterDrucken = CBool(GetSetting("SP50000", "SP52800", "SP52830Bearbeiter", "-1"))             'HW 29.04.2016
     
415     blnFolgeseitenKurzDrucken = GetSetting("SP50000", "SP52800", "SP52830FolgeseitenKurzDrucken", "0") 'Added by: GW at: 24.04.2019, Ver.: 6.5.111

420     If gintPrivBelegArt = 0 Then
425         SaveSetting "SP50000", "SP52800", "SP52850", Me.caption
        Else
430         SaveSetting "SP50000", "SP52800", "SP52860", Me.caption
        End If

435     allowFocusLostEvent = True

        '###### Formular Resizing: Parameter setzen#############
        ' DeW, ZyG, Mai 2011
        'Section- oder KeyBezeichnung sind in vielen Faellen in
        'altem Code hart eincodiert worden, manchmal wird auch
        'eine Variable verwendet...
440     Set cReSize = New FormResize
445     cReSize.setSectionBezeichnung = "SP52850"
450     cReSize.setKeyBezeichnung = "SP52850"
455     cReSize.setIstUnterFenster = False
        '
        '########################################################

        'Speichere keine Informationen (Spaltenbreiten usw.) fuer die
        'Tabellen im Form, wenn z.B. nur eine einzelne Tabelle
        'vorhanden ist, die jeweils mit neuen Daten gefuellt
        'und an eine andere Position verschoben wird (z.B. SP51000
        'Mandantenstamm
460     cReSize.IgnoreTrueDBGridInfo = True

        '######## Formular Resizing: Formular zuweisen ##########
        ' DeW, Mai 2011
        'Zuweisung von Form erst nach Groessensetzung s.o. Me.Width = ...
        'aber auf jeden Fall nach SetMinMaxInfo ... fuer die
        'Groessenbegrenzung
465     cReSize.Form = Me
470     cReSize.resize

475     Call readWindowPos(Me, "SP52800", "SP52850" & gintPrivBelegArt & "Left", "SP52850" & gintPrivBelegArt & "Top")

480     If GsTitel <> "" Then
485         GlSP51000hwnd = FindWindow(vbNullString, GsTitel)
490         SetWindowLong Me.hwnd, GWL_HWNDPARENT, GlSP51000hwnd
        End If
        
495     Set frmParentDE = Me

        Exit Sub

Fehler:
500     Call FehlerErklärung("frmSP52850", "Form_Load()")
End Sub

'############# Subclassing Methoden  ####################
'
Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
    ' This Property Let is not really needed!
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
        ' This will tell you which message you are responding to:
        ' Tell the subclasser what to do for this message (here we do all processing):
100     ISubClass_MsgResponse = emrConsume
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

100     If iMsg = WM_EXITSIZEMOVE Then
105         cReSize.resize
        End If

        'emrConsume = emrPostProcess

End Function

'#########################################################
  
'<Removed by: DFiebach at: 02.09.2024, Ver.: 6.7.101 >
'# Evtl. nicht mehr verwendet. ###KEW###
'Public Function LLPrintSammel() As Long
'
'        On Error GoTo Fehler
'
'        Dim MsgBoxText    As String
'
'        Dim SperrText     As String
'
'        Dim NrKreisText   As String
'
'        Dim rsH           As ADODB.Recordset
'
'        Dim trans         As Boolean
'
'        Dim RechnNr       As Long
'
'        Dim DruckerDialog As String
'
'        Dim SteuerPfl     As Double
'
'        Dim SteuerFr      As Double
'
'100     If gRS.RecordCount > 0 Then
'
'105         gRS.MoveFirst
'
'110         Screen.MousePointer = 11
'115         OPEN_gConn
'
'120         Do Until gRS.EOF
'
'125             trans = False
'
'130             If gRS!status = -1 Then
'
'135                 rsH.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = " & gRS!BelegID, gConn, adOpenKeyset, adLockOptimistic
'
'140                 If rsH.RecordCount > 0 Then
'
'                        'Überprüfen, ob in der Zwischenzeit von einer anderen Arbeitsstation der Beleg gedruckt wurde.
'145                     If rsH!Druck = 0 Then
'
'150                         If rsH!BelegNr = 0 Then
'
'                                'Neue Rechnungsnummer
'                                Do
'155                                 RechnNr = NummernKreisSQL(NummernKreisWaehlenSQL(gintPrivBelegArt + 8))  'Prüfen, ob allgemainer Nummernkreis für Rechnungen benutzt werden soll.
'
'                                    'Wenn die ermittelte RechnNr bereits existiert, neue RechnNr ermitteln.
'160                                 If RechnNr = -1 Then Exit Do
'165                             Loop Until IstBelegNrFrei(RechnNr + 1, gRS!BelegID, gintPrivBelegArt)
'
'                                'RechnNr = NummernKreis(NummernKreisWaehlen(gintPrivBelegArt + 8), False, False) 'Prüfen, ob allgemainer Nummernkreis für Rechnungen benutzt werden soll.
'170                             If RechnNr = -1 Then
'
'                                    'Fehler beim Ziehen der Rechnungsnummer.
'175                                 If NrKreisText = "" Then
'180                                     NrKreisText = "Folgende Belege wurden nicht gedruckt da der Nummernkreis von einem anderen Benutzer gesperrt war:"
'185                                     NrKreisText = NrKreisText & vbCrLf & "Erf.-Nr: - " & rsH!BelegID & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
'                                    Else
'190                                     NrKreisText = NrKreisText & vbCrLf & "Erf.-Nr: - " & rsH!BelegID & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
'                                    End If
'
'195                                 rsH.Close
'200                                 Protokoll iAppend, vbCrLf & "Sammeldruck -> Nummernkreis gesperrt. Rollback. BelegID " & gRS!BelegID
'205                                 GoTo NextRecord
'
'
'                                'Rechnungsnummer muss genau so ermittelt werden, wie das im Programm 571 ist.
'                                'HW 02.05.2011 Ver.: 6.1.102 - hier darf die BelegNr nicht hochgezählt werden, weil in 571 die ogik geändert wurde!
'                                '270                 RechnNr = RechnNr + 1
'
'210                             Protokoll iAppend, "Speichern beim Samelldruck. Automatisch ermittelte Beleg-Nummer: " & RechnNr & " BelegID: " & rsH!BelegID
'                            End If
'
'                            'Beträge für das Speditionsbuch
'215                         SteuerPfl = 0
'220                         SteuerFr = 0
'225                         EndBetraege "2800_Folge", gRS!BelegID, SteuerPfl, SteuerFr
'
'230                         gConn.BeginTrans
'235                         trans = True
'
'240                         If RechnNr > 0 Then rsH!BelegNr = RechnNr 'Automatisch ermittelte Beleg-Nummer
'245                         rsH!Druck = 1
'250                         rsH!AendDat = Now
'255                         rsH!AendVon = GsUser
'
'260                         If IsNull(rsH!BelegDatum) Then
'265                             rsH!BelegDatum = GdtDatum
'                            End If
'
'270                         rsH.Update
'
'                            '440               rsH.bookmark = rsH.LastModified
'275                         rsH.MoveLast
'280                         SpeditionsBuch rsH, SteuerPfl, SteuerFr
'
'285                         gConn.CommitTrans
'290                         trans = False
'
'295                         Protokoll iAppend, vbCrLf & "Einteldruck aus Sammeldruck -> BelegNr: " & rsH!BelegNr & " BelegID: " & rsH!BelegID
'
'                            '***** Einzeldruck starten
'300                         If LLPrintListe(Me, LL1, rsH!BelegID, 1, , , True) = 0 Then
'305                             If DruckerDialog = "" Then
'                                    'Sorgen dafür, dass der Druckerauswahldialog nur beim ersten Durchlauf der Schleife gezeigt wird (Falls eingestellt).
'310                                 DruckerDialog = GetSetting("SP50000", "SP52800", "SP52850DruckerDialog", "-1")
'315                                 SaveSetting "SP50000", "SP52800", "SP52850DruckerDialog", "0"
'                                End If
'
'                                'Beleg archivieren
'320                             LLPrintListe Me, LL1, rsH!BelegID, 2, False, True
'
'325                             If GblnExternesArchiv = False Then 'GblnExternesArchiv wird in SP52800B.Archivieren gesetzt.
'330                                 Screen.MousePointer = 0
'
'335                                 If DruckerDialog <> "" Then SaveSetting "SP50000", "SP52800", "SP52850DruckerDialog", DruckerDialog 'Einstellug in den Ursprunszustand setzen.
'
'                                    Exit Function
'
'                                End If
'
'                            Else
'340                             Screen.MousePointer = 0
'
'345                             If DruckerDialog <> "" Then SaveSetting "SP50000", "SP52800", "SP52850DruckerDialog", DruckerDialog 'Einstellug in den Ursprunszustand setzen.
'
'                                Exit Function
'
'                            End If
'
'                        Else
'
'350                         If MsgBoxText = "" Then
'355                             MsgBoxText = "Folgende Belege wurden von anderen Benutzern gedruckt:"
'360                             MsgBoxText = MsgBoxText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
'                            Else
'365                             MsgBoxText = MsgBoxText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
'                            End If
'                        End If
'                    End If
'                End If
'
'NextRecord:
'370             gRS.MoveNext
'            Loop
'
'375         Screen.MousePointer = 0
'
'380         If MsgBoxText <> "" Or SperrText <> "" Or NrKreisText <> "" Then
'385             If MsgBoxText <> "" Then
'390                 If SperrText <> "" Then
'395                     MsgBoxText = MsgBoxText & vbCrLf & vbCrLf & SperrText
'                    End If
'
'                Else
'
'400                 If SperrText <> "" Then
'405                     MsgBoxText = SperrText
'                    End If
'                End If
'
'410             If MsgBoxText <> "" Then
'415                 If NrKreisText <> "" Then
'420                     MsgBoxText = MsgBoxText & vbCrLf & vbCrLf & NrKreisText
'                    End If
'
'                Else
'
'425                 If NrKreisText <> "" Then
'430                     MsgBoxText = NrKreisText
'                    End If
'                End If
'
'435             MsgBox MsgBoxText, vbInformation
'            End If
'        End If
'
'440     AusDBLesen
'
'445     If DruckerDialog <> "" Then SaveSetting "SP50000", "SP52800", "SP52850DruckerDialog", DruckerDialog 'Einstellug in den Ursprunszustand setzen.
'
'        Exit Function
'
'Fehler:
'
'450     If trans Then
'455         gConn.Rollback
'460         trans = False
'        End If
'
'465     LLPrintSammel = Err.number
'
'470     If IsUpdateError(Err.number) Then
'475         If SperrText = "" Then
'480             SperrText = "Folgende Belege wurden nicht gedruckt da sie von anderen Benutzern bearbeitet waren:"
'485             SperrText = SperrText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
'            Else
'490             SperrText = SperrText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
'            End If
'
'495         If Not rsH Is Nothing Then rsH.Close
'500         LLPrintSammel = 0
'505         GoTo NextRecord
'        Else
'
'510         If DruckerDialog <> "" Then SaveSetting "SP50000", "SP52800", "SP52850DruckerDialog", DruckerDialog 'Einstellug in den Ursprunszustand setzen.
'515         gRS.MoveFirst
'520         Screen.MousePointer = 0
'
'525         If glRet <> 0 Then
'530             Call FehlerErklärung("frmSP52850", "LLPrintSammel LLFehler: " & glRet)
'            Else
'535             Call FehlerErklärung("frmSP52850", "LLPrintSammel")
'            End If
'        End If
'
'End Function
'</Removed by: DFiebach at: 02.09.2024, Ver.: 6.7.101 >

Public Function LLPrint(Mode As Integer, Optional Save As Boolean) As Long
        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       LlPrint
        ' Description:       ACHTUNG, HIER WIRD DAS GLEICHE FORMULAR ANGESTEUERT WIE IN: frmSP52831 -> SP52800B -> LLPrintListe
        ' Created by :
        ' Date-Time  :
        '
        ' Parameters :       Mode (Integer) : 1 = Druck, 2 = Vorschau
        '                    Save (Boolean) : True = LL-Datei wird gesichert. Die Option wird genutzt um die Belege zu archivieren. (Im 2-ten Durchlauf nachdem die Belege gedruckt wurden.)
        '--------------------------------------------------------------------------------
  
        Dim Formular             As String

        Dim i                    As Integer

        Dim Msg                  As Boolean

        Dim MsgBoxText           As String

        Dim SperrText            As String

        Dim NrKreisText          As String

        Dim DruckIni             As Boolean                                       'Dient zur Einmaliger Initialisierung des Druckers und Druckerauswahl pro Sammel-Druckvorgang

        Dim rs                   As ADODB.Recordset

        Dim rsH                  As ADODB.Recordset

        Dim RS1                  As ADODB.Recordset

        Dim trans                As Boolean                                       'DF 07.02.2019 , Ver.: 6.5.109 : wird evtl. nicht mehr verwendet, weil seit 2013 von HW die Logik auskommentiert ist.

        Dim ZwSumme              As Double

        Dim sql                  As String
  
        Dim SteuerPfl            As Double

        Dim SteuerFr             As Double

        Dim Ust                  As Double

        Dim Betrag               As Double

        Dim SteuerPflWrg         As Double

        Dim SteuerFrWrg          As Double

        Dim UStWrg               As Double

        Dim BetragWrg            As Double

        Dim Kurs                 As Double
  
        Dim BelegArt             As Integer

        Dim belegDatum           As Variant

        Dim BelegID              As Long

        Dim Waehrung             As String

        Dim Skonto               As Single

        Dim SkontoTage           As Integer

        Dim nettoTage            As Integer

        Dim MwSt                 As Single

        Dim RechnNr              As Long

        Dim idCollection         As Collection

        Dim currentDocType       As E_DATATYPE
  
        Dim Seite                As Long

        Dim lngBelegID           As Long                                          'HW 28.12.2015

        Static BelegNr           As Long

        Static GedruckteBelege   As String                                        'Wird gefüllt mit tatsächlich gedruckten Belegen um nur die Belege zu archivieren.
        
        Dim rsFuß                As New ADODB.Recordset                           'HW 05.11.2010  Ver.: 6.1.101 - Hier muss die Text-Logig aus den Rechnungen hin !

        Dim cnFuß                As ADODB.Connection

        Dim strTExt              As String

        Dim rec1100Texte         As ADODB.Recordset                               'HW 09.07.2012 Ver.: 6.1.115
  
        Dim HauptRS              As ADODB.Recordset                               'HW 02.09.2014

        Dim connSF               As ADODB.Connection
        
        Dim intNrKreis           As Integer                                       'DF 29.01.2019 , Ver.: 6.5.109 : Nummer des NrKreises, woher die BelegNr gezogen wurde.

        Dim strLogBelegArt       As String                                        'DF 16.01.2019 , Ver.: 6.5.109 : DruckArt -String für LOG-Datei
        
        Dim blnBelegNrFrei       As Boolean                                       'DF 22.01.2019 , Ver.: 6.5.109 : Zeigt, ob ein BelegNr bereits verwendent wurden (RAB)

        Dim blnBelegNrFortL      As Boolean                                       'DF 22.01.2019 , Ver.: 6.5.109 : Zeigt, ob ein BelegNr in einer fortlaufender Reihenfolge sich befindet

        Dim lngArt               As Integer                                       'DF 22.01.2019 , Ver.: 6.5.109 :
        
        Dim intSofa              As Integer                                       'DF 24.01.2019 , Ver.: 6.5.109 : Zeigt, ob BelegNr in RAB aus eigenen oder aus Standardnummernkreis überprüft werden soll.
        
        Dim barcodeDaten         As BarcodeData                                   'DF 07.02.2019 , Ver.: 6.5.109 : Barcode (wurde bis deiser Verison bei SAMMELDRUCK nicht gedruckt)
        
        Dim strFormularNr        As String                                        'DF 13.11.2019 , Ver.: 6.5.113 : Nummer des Druckformulares
        
        Dim strStCodeH           As String                                        'DF 29.07.2024 , Ver.: 6.7.101 : St.Code des Hauptsatzes (E-Rechnung)
                
        Dim intSteuerTextLkz     As Integer                                       'DF 29.07.2024 , Ver.: 6.7.101 : Lkz des SteuerTextes für die ganze Rechnung , wird anhand des Steuer-Schlüssel der Rechnung ermittelt.
        
        Dim strZahlungsText      As String                                        'DF 03.09.2024 , Ver.: 6.7.101 : ZahlungsKonditionen Text
        
        Dim strZahlungsTextNetto As String                                        'DF 03.09.2024 , Ver.: 6.7.101 : ZahlungsKonditionen Text (Netto)
        
        On Error GoTo Fehler

100     Set rs = New ADODB.Recordset
105     Set rsH = New ADODB.Recordset
110     Set RS1 = New ADODB.Recordset
115     Set rsFuß = New ADODB.Recordset
120     Set rec1100Texte = New ADODB.Recordset
        
        '<Added by: DFiebach at: 16.01.2019, Ver.: 6.5.109 >
125     If gintPrivBelegArt = 0 Then

130         strLogBelegArt = "Rechnungsdruck"
135         lngArt = 1   'Ausg.-Rechn.
140         strFormularNr = "35"

        Else
        
145         strLogBelegArt = "Gutschriftsdruck"
150         lngArt = 2   'Ausg.-Gutschr.
155         strFormularNr = "36"

        End If
        
160     gblnBelegNrChecked = False                                              'globalen Zeiger auf bereits durchgeführte ÜBerprüfung auf "fortlaufende" BelegNr zurücksetzen.
        
        '</Added by: DFiebach at: 16.01.2019, Ver.: 6.5.109 >
        
165     If Not Save Then
            ' Wenn KEINE Archivierung
170         GedruckteBelege = ""
            
175         If LL18CheckBildFile((strFormularNr)) = False Then
                
180             LLPrint = -1
                
                Exit Function
                
            End If
            
        End If
        
185     If gRS.RecordCount > 0 Then

190         Set HauptRS = gRS.Clone(adLockReadOnly)
195         Set HauptRS.ActiveConnection = Nothing
200         HauptRS.filter = gRS.filter

205         HauptRS.MoveFirst
            
            'HW 17.10.2013
            '#######################################
            
            'CSBmk <DRUCK-OPTIONEN>
210         RechnNr = CLng(val(objDruckOptionen.CurrentBelegNr))
            
            'DF 07.02.2019 , Ver.: 6.5.109 : Bei Sammelfaktura darf BelegNr nicht freigeschaltet werden.
215         objDruckOptionen.EnableBelegNr = False
220         objDruckOptionen.resizeFactor = cReSize.CurrFactor
            
225         If showDruckOptionen Then objDruckOptionen.CurrentBelegNr = "0"
230         If showDruckOptionen Then objDruckOptionen.ShowMe Me.lblDruckOptionen.caption, Me            'DruckOptionen Dialog aufrufen
    
235         showDruckOptionen = False                                          'DruckOptionen Fenster nur einmal anzeigen

240         If objDruckOptionen.Canceled Then                                  'Wird das DruckOptionen Fenster geschlossen, Druckroutine beenden

245             LLPrint = -2
            
                Exit Function

            End If

250         Screen.MousePointer = 11
            
            'CSBmk <ALLE VORHANDENEN BELEGE (HAUPTSCHLEIFE)>
255         Do Until HauptRS.EOF
                
260             trans = False

                '<Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
                
265             intNrKreis = 0
                
                '230             blnBelegNrFrei = False
                '235             blnBelegNrFortL = False
                '</Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
                
                'CSBmk <TEMP-VARIABLEN ZURÜCKSETZEN>
270             gEnmKudnenERechnungType = eERechnungType.None                  'DF 03.09.2024 , Ver.: 6.7.101
275             intSteuerTextLkz = 0
280             strStCodeH = ""
285             strZahlungsTextNetto = ""
290             strZahlungsText = ""

295             If Not objERechnung Is Nothing Then objERechnung.Clear
                
                'CSBmk <IST BELEG MARKIERT ?>
300             If HauptRS!status = -1 Then

305                 RechnNr = 0

310                 Msg = False

315                 BelegID = HauptRS!BelegID

320                 If Save Then

                        'If InStr(1, GedruckteBelege, BelegID & ",") = 0 Then GoTo NextRecord

                    End If

325                 DoEvents

330                 If Not rsH Is Nothing Then

                        On Error Resume Next

335                     rsH.Close

                        On Error GoTo Fehler

                    End If
                    
                    'CSBmk <HAUPT-RECORDSET VARIABLEN>
340                 rsH.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = " & BelegID, gConn, adOpenKeyset, adLockOptimistic

345                 If rsH.RecordCount > 0 Then

                        '<Added by: IL at: 9.2.2024-09:15:17 on machine: T017>

                        'Überprüfung auf das Vorhandensein von Verpackungen und der entsprechenden EN-Code
                        
                        'CSBmk <ÜBERPRÜFUNG DER EINHEITEN>
350                     If Not CheckENCodeBeiVerpackung(1, CStr(BelegID), False) And Mode = 1 Then

355                         Call msgText(1, 2357, 0, 0, 0)

360                         GsMsgText(0) = Replace(GsMsgText(0), "%1", BelegID)

365                         Call MsgBox(GsMsgText(0), vbExclamation, strMeldungCap)
                        
                        End If

                        '</Added by: IL at: 9.2.2024-09:15:17 on machine: T017>

                        'Beträge
370                     SteuerPfl = 0

375                     SteuerFr = 0
                        
                        'CSBmk <ENDBETRÄGE PRO HAUPT SATZ AUS DESSEN FOLGESÄTZEN>
380                     EndBetraege "2800_Folge", BelegID, SteuerPfl, SteuerFr
                        
                        'CSBmk <BELEGDATUM PRO HAUPTSATZ>
                        
                        '<Removed by: DFiebach at: 07.02.2019, Ver.: 6.5.109 >
                        ' # BelegDatum darf nur beim gedruckten Beleg gespeichert werden
                        '                        'DH, 21.04.2014, Das im DruckOptionen Fenster eingegebene Datum uebernehmen
                        '305                     rsH.Fields("BelegDatum").Value = objDruckOptionen.CurrentBelegDatum
                        '310                     rsH.Update
                        '</Removed by: DFiebach at: 07.02.2019, Ver.: 6.5.109 >

385                     llCurrentFormNr = CInt(rsH.Fields("Art").value) + 35        'Formularnummer fuer diesen Druck merken
                        
390                     If Mode < 2 Then
                            
                            'wenn DRUCK
                            
                            'Überprüfen, ob in der Zwischenzeit von einer anderen Arbeitsstation der Beleg gedruckt wurde.
                            
395                         If rsH!Druck = 0 Then

400                             If rsH!BelegNr = 0 Then
                                    
                                    'CSBmk <BELEG-NR VERGABE>
                                    
                                    '<Added by: DFiebach at: 07.02.2019, Ver.: 6.5.109 >
                                    ' #  BelegNr-Vergabe erst nach DRUCKERAUSWAHL-DIALOG
405                                 RechnNr = 0
                                    '</Added by: DFiebach at: 07.02.2019, Ver.: 6.5.109 >
                  
                                Else 'HW 17.10.2013 Wenn die BelegNr nicht Null ist
                                    
                                    'CSBmk <BELEG-NR BEREITS EXISTIERT>
410                                 objDruckOptionen.CurrentBelegNr = rsH("Belegnr").value
415                                 RechnNr = rsH("Belegnr").value

                                End If
                            
                            Else
                                
                                'Beleg bereits gedruckt
                                
420                             Msg = True

425                             If MsgBoxText = "" Then
430                                 MsgBoxText = "Folgende Belege wurden von anderen Benutzern gedruckt:"
435                                 MsgBoxText = MsgBoxText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
                                
                                Else
440                                 MsgBoxText = MsgBoxText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
                                
                                End If
                            
                            End If
                        
                        Else
                            
                            'Wenn Archivierung
                            
                            'CSBmk <BELEG-NR FÜR ARCHIVIERUNG-2>
445                         BelegNr = rsH!BelegNr  'HW 17.10.2013 eingebaut um jede Rechnung zu Archivieren! Damit immer im Richtigen Modus die Belegnr zugewiesen wird!
                        
                        End If
                        
450                     If Msg = False Then
                                      
                            'CSBmk <DEFINIERE VARIABLEN UND FELDER FÜR HAUPT-SATZ>
                            
455                         If Mode = 1 Then GedruckteBelege = GedruckteBelege & BelegID & ","

460                         LL1.LlDefineVariableStart                           'Variablenpuffer löschen.
465                         LL1.LlDefineFieldStart                              'Variablenpuffer löschen.

470                         Call LL18GestaltungFormular(LL1, rsH!Art + 35, rsH("MCode").value, MandantArr(1))
475                         Call LLDefineVariablen(LL1, rsH, "Kd_")
480                         Call LLDefineFelder(LL1, rsH, "Kd_")
485                         Call LLDefineTexte(LL1)                             'DF 24.10.2024 , Ver.: 6.7.101 : ZusatzTexte usw.
                            
                            'CSBmk <ZUSATZ-TEXT AUF RNG-GUT>
490                         Call DefineZusatztext(rsH, LL1)                     'MW 13.11.08 Ver.: 5.4.119 Zusatztext

495                         LL1.LlDefineVariableExt "Kd_VonDatum", "" & rsH!vonDatum, LL_TEXT
500                         LL1.LlDefineVariableExt "Kd_BisDatum", "" & rsH!bisDatum, LL_TEXT
505                         LL1.LlDefineVariableExt "ERechnungArt", 0, LL_NUMERIC               'DF 04.11.2024 , Ver.: 6.7.101

510                         BelegArt = rsH!Art
515                         belegDatum = rsH!belegDatum
520                         Waehrung = rsH!Wrg1
525                         Skonto = rsH!ZSkto
530                         SkontoTage = rsH!ZSktoTage
535                         nettoTage = rsH!ZTage
540                         MwSt = rsH!MwSt
545                         Kurs = rsH!Kurs

                            'CSBmk <OPT:BEARBEITER DRUCKEN>
                            
550                         If BearbeiterDrucken Then
555                             LL1.LlDefineFieldExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
560                             LL1.LlDefineVariableExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
                            
                            Else
565                             LL1.LlDefineFieldExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
570                             LL1.LlDefineVariableExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
                            
                            End If
                            
                            'CSBmk <OPT:ADRESSE AUF FOLGESEITEN DRUCKEN>
                            '<Added by: GW at: 24.04.2019, Ver.: 6.5.111 >
                            
575                         If blnFolgeseitenKurzDrucken Then
580                             LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN
585                             LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN
                            
                            Else
590                             LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN
595                             LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN
                            
                            End If

                            '</Added by: GW at: 24.04.2019, Ver.: 6.5.111 >
            
                            '<Added by: DFiebach at: 07.02.2019, Ver.: 6.5.109 >
                            ' # Behandlung des BelgDatums
                            
600                         If objDruckOptionen.CurrentBelegDatum <> "" Then
605                             LL1.LlDefineFieldExt "Kd_BelegDatum", objDruckOptionen.CurrentBelegDatum, LL_DATE_LOCALIZED
610                             belegDatum = objDruckOptionen.CurrentBelegDatum
                            
                            End If
                           
615                         If Not IsDate(belegDatum) Then belegDatum = GdtDatum
                            
                            '</Added by: DFiebach at: 07.02.2019, Ver.: 6.5.109 >
                            
                            'HW 17.10.2013
                            '#################################
620                         LL1.LlDefineFieldExt "ProbeDruckText", ZusatzText(4, "55710"), LL_TEXT
                            
625                         If Mode = 2 And Save = False Then 'HW 17.10.2013 Wenn ProbeDruck Dann auf 1 ansonsten auf 0

630                             LL1.LlDefineFieldExt "ProbeDruck", 1, LL_NUMERIC 'HW 17.10.2013
                            
                            Else
                            
635                             LL1.LlDefineFieldExt "ProbeDruck", 0, LL_NUMERIC 'HW 17.10.2013
                            
                            End If

                            '#################################

                            'HW 09.07.2012 Ver.: 6.1.114
                            'HW 30.07.2012 Ver.: 6.1.115 Steuertexte nachgepflegt! weil vorher Fehler passierte!
                            
                            'CSBmk <STEUER-TEXTE>
640                         objERechnung.colSteuerTexte.Clear
                            
645                         rec1100Texte.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Sort <= 7", gConn, adOpenStatic, adLockReadOnly

650                         If rec1100Texte.RecordCount > 0 Then
                             
655                             Do While Not rec1100Texte.EOF

660                                 LL1.LlDefineFieldExt "Steuertext" & rec1100Texte!Sort, "" & rec1100Texte!text, LL_TEXT

665                                 If Not objERechnung.colSteuerTexte.ContainsKey(CStr(rec1100Texte!Sort)) Then
                        
670                                     Call objERechnung.colSteuerTexte.Add("" & rec1100Texte!text, CStr(rec1100Texte!Sort)) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung

                                    End If

675                                 rec1100Texte.MoveNext
                  
                                Loop
                            
                            Else
                            
680                             LL1.LlDefineFieldExt "Steuertext", "", LL_TEXT
                            
                            End If

685                         rec1100Texte.Close

690                         LL1.LlDefineFieldExt "Steuertext", "" & gstrSteuerText, LL_TEXT 'HW 05.07.2012  Ver.: 6.1.129
695                         LL1.LlDefineFieldExt "SteuerSchl", rsH!Ust, LL_NUMERIC
                            
700                         objERechnung.SteuerText = GetSteuerText(CInt(rsH!Ust), SteuerFr, gstrSteuerText, objERechnung.colSteuerTexte.GetItem("2"), objERechnung.colSteuerTexte.GetItem("4"), objERechnung.colSteuerTexte.GetItem("6")) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung
                
                            '#################################
                            
                            '<Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
                            'St.Code des Hauptsatzes ahnahd des gewählten St.Schl
                
                            'CSBmk <STEUER-CODE HAUPT>
705                         gEnmKudnenERechnungType = modERechnung.GetKundenERechnungType(rsH!MCode) 'DF 23.07.2024 , Ver.: 6.7.101

710                         If IsEBelegDoc Then LL1.LlDefineVariableExt "ERechnungArt", CInt(gEnmKudnenERechnungType), LL_NUMERIC
                            
715                         intSteuerTextLkz = GetRNGGUTSteuerTextLkz(GintBelegArt, CInt(rsH!Ust))
                
720                         strStCodeH = GetStCodeFromSteuerText(CStr(intSteuerTextLkz), "Rng")

                            '</Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
                            
                            'CSBmk <DEFINIERE VARIABLEN UND FELDER FÜR FOLGE-SÄTZE DES HAUPTSATZES>
                            'CSBmk <FOLGE-RECORDSET VARIABLEN>
                            
725                         rs.Open "SELECT * FROM [2800_Folge] WHERE BelegID = " & BelegID & " ORDER BY Nr", gConn, adOpenStatic, adLockReadOnly     'IL adLockReadOnly   adLockOptimistic

730                         If rs.RecordCount > 0 Then

735                             Call LLDefineVariablen(LL1, rs, "Re_")

740                             ZwSumme = 0
745                             Seite = 1

750                             LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
755                             LL1.LlDefineFieldExt "Re_ZwSumme", 0, LL_NUMERIC
760                             LL1.LlDefineFieldExt "ZahlungsZiel", "", LL_TEXT
765                             LL1.LlDefineFieldExt "LetzteSeite", 0, LL_NUMERIC
770                             LL1.LlDefineFieldExt "Re_EPreisDezStellen", 2, LL_NUMERIC 'MW 11.01.07 Ver.: 5.3.106
                                  
                                'CSBmk <BETRÄGE>
775                             Ust = 0
780                             Betrag = 0
785                             SteuerPflWrg = 0
790                             SteuerFrWrg = 0
795                             UStWrg = 0
800                             BetragWrg = 0
  
                                '<Removed by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >
                                '680                             Ust = Runden((SteuerPfl * MwSt / 100), 2)
                                '</Removed by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >

                                '<Added by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >
805                             Ust = SteuerBetrag(CCur(SteuerPfl), Format(CCur(MwSt), "#.00"))
                                '</Added by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >

810                             Betrag = SteuerPfl + Ust + SteuerFr

                                'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
815                             SteuerPflWrg = RundenMitVz(SteuerPfl * Kurs, 2)
820                             SteuerFrWrg = RundenMitVz(SteuerFr * Kurs, 2)
825                             UStWrg = Runden(Ust * Kurs, 2)
830                             BetragWrg = SteuerPflWrg + UStWrg + SteuerFrWrg
  
835                             LL1.LlDefineFieldExt "Re_SummeSteuerPfl", SteuerPfl, LL_NUMERIC
840                             LL1.LlDefineFieldExt "Re_SummeSteuerFr", SteuerFr, LL_NUMERIC
845                             LL1.LlDefineFieldExt "Re_USt", Ust, LL_NUMERIC
850                             LL1.LlDefineFieldExt "Re_Betrag", Betrag, LL_NUMERIC
                                                                
                                'CSBmk <ZAHLUNGS-KONDITIONEN NETTO>
855                             strZahlungsTextNetto = "" & ZahlungsZielNetto(belegDatum, Betrag, Waehrung, Skonto, SkontoTage, nettoTage, objDruckOptionen.CurrentValutaDatum)
860                             LL1.LlDefineFieldExt "ZahlungsZielNetto", strZahlungsTextNetto, LL_TEXT 'DH, 16.02.2015, 6.4.103, ValutaDatum aus den DruckOptionen uebernehmen
865                             objERechnung.ZHinweisNetto = strZahlungsTextNetto
                                
                                'HW 08.08.2011 Ver.: 6.1.101 - Hier muss die Text-Logik aus den Rechnungen hin !
                                '#############################################################
870                             Set cnFuß = New ADODB.Connection
875                             cnFuß.ConnectionString = GetACCESSConnectionString(LOG_CONNECTION)
880                             cnFuß.Open
  
885                             strTExt = ""

890                             If rsH!Ust = 2 Then

895                                 rsFuß.Open "SELECT * FROM [ZusatzTexte_55710] WHERE Nr = 29", cnFuß, adOpenStatic, adLockReadOnly

900                                 If rsFuß.RecordCount >= 0 Then

905                                     strTExt = rsFuß.Fields("DE").value
                                    
                                    End If
                                    
910                                 If rsFuß.state = adStateOpen Then rsFuß.Close 'IL 29.11.2024 , Ver.: 6.7.101 : Muss Mann den Recordset schließen, damit beim nächsten Öffnen keine Fehlermeldung angezeigt wird
                                    
915                                 rsFuß.Open "SELECT * FROM [ZusatzTexte_55710] WHERE Nr = 109", cnFuß, adOpenStatic, adLockReadOnly

920                                 If rsFuß.RecordCount >= 0 Then

925                                     strTExt = strTExt + " " + rsFuß.Fields("DE").value
                                    
                                    End If
                                
                                Else

930                                 If rsH!Ust = 0 Then

935                                     rsFuß.Open "SELECT * FROM [ZusatzTexte_55710] WHERE Nr = 108", cnFuß, adOpenStatic, adLockReadOnly

940                                     If rsFuß.RecordCount >= 0 Then

945                                         strTExt = strTExt + " " + rsFuß.Fields("DE").value
                                        
                                        End If

                                    Else
950                                     rsFuß.Open "SELECT * FROM [ZusatzTexte_55710] WHERE Nr = 110", cnFuß, adOpenStatic, adLockReadOnly

955                                     If rsFuß.RecordCount >= 0 Then

960                                         strTExt = strTExt + " " + rsFuß.Fields("DE").value
                                        
                                        End If
                                    
                                    End If
                                
                                End If

965                             If rsFuß.state = adStateOpen Then rsFuß.Close
970                             cnFuß.Close

975                             LL1.LlDefineFieldExt "Re_SteuerText", strTExt, LL_TEXT
                                '#############################################################
  
980                             LL1.LlDefineFieldExt "Re_SummeSteuerPflWrg", SteuerPflWrg, LL_NUMERIC
985                             LL1.LlDefineFieldExt "Re_SummeSteuerFrWrg", SteuerFrWrg, LL_NUMERIC
990                             LL1.LlDefineFieldExt "Re_UStWrg", UStWrg, LL_NUMERIC
995                             LL1.LlDefineFieldExt "Re_BetragWrg", BetragWrg, LL_NUMERIC
  
                                'Soll Spalte Rabatt sichtbar sein?
1000                            sql = "SELECT Max([Rabatt]) AS MaxRabatt "
1005                            sql = sql & "FROM [2800_Folge] WHERE BelegID = " & BelegID
1010                            RS1.Open sql, gConn, adOpenStatic, adLockReadOnly
1015                            LL1.LlDefineFieldExt "RabattVisible", RS1!MaxRabatt, LL_NUMERIC
1020                            RS1.Close
  
                                'MW 26.04.05
                                'Handelt es sich um ein Beleg mit LieferscheinArtikel
1025                            sql = "SELECT TOP 1 SatzTyp "
1030                            sql = sql & "FROM [2800_Folge] WHERE SatzTyp = 'L' AND BelegID = " & BelegID
1035                            RS1.Open sql, gConn, adOpenStatic, adLockReadOnly
1040                            LL1.LlDefineFieldExt "LSArtikel", RS1.RecordCount, LL_NUMERIC
1045                            RS1.Close

1050                            LL1.LlDefineFieldExt "KostenstellenDruck", Abs(GbKostenstellenPflicht), LL_NUMERIC 'MW 28.12.07
  
                                'Damit rs.RecordCount abgefragt werden kann.
1055                            rs.MoveLast
1060                            rs.MoveFirst
                               
                            Else
1065                            Msg = True
                               
                            End If
                           
                        End If
                       
                    Else
1070                    Msg = True
                       
                    End If
                    
1075                If Msg = False Then

1080                    If DruckIni = False Then
                            
                            'CSBmk <DRUCKER-DEFINITION + OPTIONEN + EINSTELLUNGEN>
                            
                            'Der Drucker wird nur beim ersten Durchlauf initialisiert.
1085                        Formular = FormularPfad("SP52800.lst")

1090                        LL1.LlSetDebug (LL_DEBUG_CMBTLL)

1095                        glRet = LL1.LlPreviewDeleteFiles(Formular, "")

1100                        If Mode < 2 Then

                                '<Added by: IL at: 03.12.2024, Ver.: 6.7.101 >
1105                            If Save Then

                                    'Ablage

1110                                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, Me.hwnd, "printing list")

                                Else
                                    '</Added by: IL at: 03.12.2024, Ver.: 6.7.101 >
                               
                                    'Druck

1115                                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, Me.hwnd, "printing list")

                                End If
                               
                            Else
                            
                                'Vorschau

1120                            If Save Then

1125                                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, Me.hwnd, "Archivierung")
                                   
                                Else
                                
1130                                glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, Me.hwnd, "printing list to preview")
                                   
                                End If
                               
                            End If

1135                        If glRet < 0 Then

1140                            GoTo Fehler
                               
                            End If

                            'Eigentlich muesste im Sammeldruck ja genauso wie im Einzeldruck das Druckoptionen-Fenster gezeigt werden
                            'LL18PositionierungFormular LL1, objDruckOptionen.FormularNr   'DH, 02.10.2013, 6.2.100, Ausrichtung der Objekte nachgepflegt
1145                        LL18PositionierungFormular LL1, 35   'DH, 02.10.2013, 6.2.100, Ausrichtung der Objekte nachgepflegt - FormularNr erstmal fest eingebaut, da das DruckOptionen Fenster hier noch nicht genutzt wird

                            'DH, 10.02.2014, 6.2.102, Im Sammeldruck die Kopiensteuerung ermoeglichen
                            '##########
                       
1150                        If Mode = 1 Then  'Nur beim richtigen Druck und nicht bei der Wiederholung

1155                            Call LL18SetCopies(LL1, 1, 35, "SP52850")
                               
                            End If

                            '##########
                            
                            'CSBmk <DRUCKERAUSWAHL-DIALOG>
              
1160                        If CBool(GetSetting("SP50000", "SP52800", "SP52850DruckerDialog", "-1")) = True Then 'Druckdialog

1165                            If Not Save Then

1170                                glRet = LL1.LlPrintOptionsDialog(Me.hwnd, "Drucker")

1175                                If glRet = LL_ERR_USER_ABORTED Then

1180                                    Screen.MousePointer = 0
                                        
1185                                    LLPrint = LL_ERR_USER_ABORTED
                                        
1190                                    LL1.LlPrintEnd 0                        'DH, 08.10.2013, 6.2.100, Wenn ueber den Druckdialog abgebrochen wurde, muss der Druckjob beendet werden

1195                                    If Mode = 1 Then

1200                                        rsH!belegDatum = Null               'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
1205                                        rsH!Druck = 0
1210                                        rsH.Update
                                           
                                        End If
                                        
1215                                    rsH.Close
                                        
                                        Exit Function
                                       
                                    End If

1220                                If Mode < 2 Then SaveSetting "SP50000", "SP52800", "SP52850_PRNOPT_COPIES", LL1.LlPrintGetOption(LL_PRNOPT_COPIES)
                                   
                                End If
                               
                            End If

                            'Nach Combit ist es unbedingt notwendig, die von LlPrintSetOption gesetzte Kopienanzahl
                            'durch den Aufruf von LL_PRNOPT_COPIES_SUPPORTED zu bestätigen.
1225                        glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)

1230                        Screen.MousePointer = 11

1235                        DruckIni = True
                           
                        End If
                        
                        '##### HIER BELEG NR ZIEHEN + BARCODE #####
                        
1240                    If Mode = 1 Then
                        
1245                        RechnNr = GetBelegNr(GintBelegNrKreisNr, BelegID, BelegArt, intNrKreis, True, programmNr) 'DF 14.11.2024 , Ver.: 6.7.101 : gintPrivBelegArt + 8 -> GintBelegNrKreisNr

1250                        If RechnNr > 0 Then

1255                            rsH.Fields("BelegNr").value = RechnNr
1260                            rsH.Fields("BelegNrKreis").value = intNrKreis   'DF 29.01.2019 , Ver.: 6.5.109 : Nummer des BelegNr NrKreises speichern

1265                            If objDruckOptionen.CurrentValutaDatum <> "" Then rsH.Fields("ValutaDatum").value = objDruckOptionen.CurrentValutaDatum
1270                            rsH.Fields("BelegDatum").value = belegDatum
1275                            rsH!Druck = 1
1280                            rsH!AendDat = Now
1285                            rsH!AendVon = GsUser

1290                            If IsNull(rsH!belegDatum) Then

1295                                rsH!belegDatum = GdtDatum
                                   
                                End If

1300                            rsH.Update

1305                            rsH.MoveLast
                                
                                'CSBmk <SPEDITIONSBUCH-ÜBERGABE>
1310                            SpeditionsBuch rsH, SteuerPfl, SteuerFr

                                'CSBmk <BELEG-NR FÜR ARCHIVIERUNG-1>
1315                            BelegNr = rsH!BelegNr ' Für die Archivierung

1320                            Protokoll iAppend, vbCrLf & "Sammeldruck -> BelegNr: " & rsH!BelegNr & " BelegID: " & BelegID
                               
                            Else
                            
1325                            rsH!Druck = 0
1330                            rsH!belegDatum = Null                            'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
1335                            rsH.Update
1340                            rsH.Close
                                
1345                            Protokoll iAppend, "##### -> ABBRUCH DURCH BENUTZER, BELEG-NR NICHT FORTLAUFEND, SF-" & strLogBelegArt & ", BelegNr = " & BelegNr & ", BelegID =" & BelegID & ""
                        
1350                            LL1.LlPrintEnd 0

1355                            LLPrint = LL_ERR_USER_ABORTED
                        
                                Exit Function
                               
                            End If
                            
1360                        Call LLDefineVariablen(LL1, rsH, "Kd_")
1365                        Call LLDefineFelder(LL1, rsH, "Kd_")
                           
                        End If
                        
1370                    LL1.LlDefineFieldExt "Kd_BelegDatum", belegDatum, LL_DATE_LOCALIZED

                        'CSBmk <BARCODE DEFINIEREN>
1375                    barcodeDaten.Seperator = ";"
1380                    barcodeDaten.BelegNr = rsH!BelegNr
1385                    barcodeDaten.belegDatum = IIf(IsNull(rsH!belegDatum), "", rsH!belegDatum)

1390                    barcodeDaten.Name1 = "" & rsH.Fields("Name1").value
1395                    barcodeDaten.Name2 = "" & rsH.Fields("Name2").value
1400                    barcodeDaten.Adresse = "" & rsH.Fields("Straße").value
1405                    barcodeDaten.Lkz = "" & rsH.Fields("Lkz").value
1410                    barcodeDaten.Plz = "" & rsH.Fields("Plz").value
1415                    barcodeDaten.Ort = "" & rsH.Fields("Ort").value
1420                    barcodeDaten.ORTSTEIL = rsH.Fields("Ortsteil").value

1425                    Call LL18DefineBarcode(LL1, barcodeDaten, 35, rsH.Fields("MCode").value) 'Barcode im Formular definieren
                                                
                        'CSBmk <DRUCK-VORGANG START PRO HAUPT-SATZ>
                        
                        'Variablen drucken
1430                    glRet = LL1.LLPrint

1435                    Sleep (1)
  
                        'Solange das Ende der Posten-Tabelle nicht erreicht ist...
1440                    While Not rs.EOF

1445                        DoEvents

                            'Prozentbalken setzen
1450                        glRet = LL1.LlPrintSetBoxText("Drucken", (100# * HauptRS.AbsolutePosition / HauptRS.RecordCount))

                            'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz

                            'Datensatzfelder der Liste bekanntmachen.
                 
1455                        If Trim(rs!Einheit) = "%" Then

                                'ORIG
                                'ZwSumme = ZwSumme + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)

                                'DF 03.03.2025 , Ver.: 6.7.106: NEU -> rs!Menge / 100 führte zum falschen Ergebnis, -> rs!EPreis / 100 analog zum fpSread-Formel.
1460                            ZwSumme = ZwSumme + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)
                               
                            Else
                             
1465                            ZwSumme = ZwSumme + RundenMitVz((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                               
                            End If

1470                        LL1.LlDefineFieldExt "Re_EPreisDezStellen", GetMin2DezimalStellen(rs!EPreis), LL_NUMERIC 'MW 11.01.07 Ver.: 5.3.106
1475                        LL1.LlDefineFieldExt "Re_ZwSumme", ZwSumme, LL_NUMERIC
            
1480                        If rs.RecordCount = rs.AbsolutePosition Then          'DH, 07.09.2017, 6.5.100,If rs.RecordCount = rs.AbsolutePosition + 1 Then. +1 funktioniert mit ADODB scheinbar nicht mehr

1485                            LL1.LlDefineFieldExt "LetzteSeite", 1, LL_NUMERIC
                                
1490                            If gintPrivBelegArt = 0 Then 'Nur bei Rechnungen
                                    
                                    'CSBmk <ZAHLUNGS-KONDITIONEN>
1495                                strZahlungsText = ZahlungsZiel(belegDatum, Betrag, Waehrung, Skonto, SkontoTage, nettoTage, objDruckOptionen.CurrentValutaDatum)

1500                                LL1.LlDefineFieldExt "ZahlungsZiel", strZahlungsText, LL_TEXT  'DH, 16.02.2015, 6.4.103, Valuta aus den DruckOptionen uebernehmen
1505                                objERechnung.ZHinweisBrutto = strZahlungsText
                                   
                                End If
                               
                            End If

1510                        Call LLDefineFelder(LL1, rs, "Re_")

                            'Seitenumbruch
1515                        If rs!SatzTyp = "S" Then

1520                            Seite = Seite + 1
1525                            LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC

1530                            glRet = LL1.LLPrint
                               
                            End If

                            'Felder drucken und wenn Seitenumbruch erfolgt ist,
                            'Variablen und Felder erneut drucken
                            'HW 01.11.2013 Select Case eingeführt da hier auch auf User_Aborted abgefragt werden muss! Durckdialog abbrechen!
                            '#############################################################
1535                        glRet = LL1.LlPrintFields
    
1540                        Select Case glRet
                     
                                Case LL_WRN_REPEAT_DATA
                                
1545                                Seite = Seite + 1
1550                                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1555                                glRet = LL1.LLPrint
         
1560                                While LL1.LlPrintFields = LL_WRN_REPEAT_DATA

1565                                    Seite = Seite + 1
1570                                    LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1575                                    glRet = LL1.LLPrint
                      
                                    Wend

1580                            Case LL_ERR_USER_ABORTED

                                    'ABBRUCH
1585                                HauptRS.MoveFirst
                          
1590                                Do While Not HauptRS.EOF

1595                                    rsH.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = " & HauptRS("BelegID").value, gConn, adOpenKeyset, adLockOptimistic
                                     
1600                                    If rsH.RecordCount > 0 Then
1605                                        rsH!belegDatum = Null                                     'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
1610                                        rsH!Druck = 0
1615                                        rsH.Update
                                           
                                        End If

1620                                    rsH.Close
1625                                    HauptRS.MoveNext
               
                                    Loop

1630                                LL1.LlPrintEnd 0 'HW 22.10.2013
1635                                Screen.MousePointer = 0

1640                                If HauptRS.RecordCount > 0 Then HauptRS.MoveFirst
1645                                LLPrint = LL_ERR_USER_ABORTED

                                    Exit Function
                                   
                            End Select

                            '#############################################################
                            
1650                        If rs!SatzTyp = "Z" Then ZwSumme = 0
1655                        rs.MoveNext
                            
                        Wend
                        
                        Do
1660                        glRet = LL1.LlPrintFieldsEnd()

1665                        If glRet = LL_WRN_REPEAT_DATA Then

1670                            Seite = Seite + 1
1675                            LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1680                            LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC
1685                            LL1.LlDefineFieldExt "LetzteSeite", 1, LL_NUMERIC
                                'Neue Seite
1690                            LL1.LLPrint
                               
                            End If
              
1695                    Loop Until glRet <> LL_WRN_REPEAT_DATA
                        
                        'CSBmk <STEUER-CODE SPEICHERN HAUPT UND FOLGE>
            
                        '<Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >
                        '# Texte im Hauptsatz speichern. Damit in der PDF und ERechnung das gleiche steht.
1700                    If Mode = 1 Then                                                                 'Nur beim Druck.

1705                        rsH!ERechnungArt = modERechnung.GetERechnungTypeValueForDB(gEnmKudnenERechnungType)    'DF 23.07.2024 , Ver.: 6.7.101
1710                        rsH!StCode = strStCodeH
            
1715                        rsH.Update

1720                        Call SetStCode(E_DATATYPE.Sonderfaktura_Rechnung, 1, rsH!BelegID, intSteuerTextLkz, rsH!Ust, False, GintBelegArt) ' An der Stelle wird zw. SF-RNG und -GUT nicht unterschieden, da beide in der gelichen Tabelle gespeichert werden.

                        End If

                        '</Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >
                        
                    End If

                    'HW 02.09.2014 Testweise mal reingenommen!
1725                LL1.LlPrintResetProjectState
               
1730                If Save Then                                                'HW 17.10.2013 eingebaut um jede Rechnung zu Archivieren!

                        'CSBmk <ARCHIVIERUNG PRO BELEG>
1735                    GoSub Archivieren
                       
                    End If

                End If

                'CSBmk <NÄCHSTER BELEG>
NextRecord:

1740            If rs.state = adStateOpen Then rs.Close

1745            HauptRS.MoveNext

            Loop
            
            '######################                                             'HW 17.10.2013, Wenn Archiviert wurde und er fertig mit dem Durchlauf ist Beenden!
   
1750        If Save Then

1755            GoSub Beenden
               
            End If

            '######################

            'CSBmk <BELEG-ARCHIVIEREN LOGIK>
Archivieren:
            'HW 17.10.2013 eingebaut um jede Rechnung zu Archivieren!

1760        If DruckIni Then

1765            Screen.MousePointer = 0

                'Tabellen-Ausdruck beenden
              
                Do
1770                glRet = LL1.LlPrintFieldsEnd()

1775                If glRet = LL_WRN_REPEAT_DATA Then

                        'Neue Seite
1780                    LL1.LLPrint
1785                    glRet = LL1.LlPrintFieldsEnd()
          
                    End If
      
1790            Loop Until glRet <> LL_WRN_REPEAT_DATA
                
                'CSBmk <DRUCK BEENDEN>
1795            glRet = LL1.LlPrintEnd(0)
     
                'Beim Preview-Druck Preview anzeigen und dann Preview-Datei (.LL) löschen
                
1800            If Mode = 2 Then
                    
                    'VORSCHAU-MODE
                    
1805                If Save Then
                        
                        'ARCHIVIERUNG
                        
1810                    lngBelegID = BelegID                                    'HW 28.12.2015 Ver.: 6.4.116 BelegID überprüfen und notfalls neu ziehen

1815                    Call CheckBelegID(lngBelegID, "[5700_Haupt]")

                        '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
                        Dim strTableName As String

1820                    Select Case GintBelegArt

                            Case 0
    
1825                            strTableName = "[2800_BelegArchiv_Rng]"
    
1830                        Case 1
    
1835                            strTableName = "[2800_BelegArchiv_Gut]"
    
1840                        Case 2
    
1845                            strTableName = "[2800_BelegArchiv_Ang]"
    
1850                        Case 3
    
1855                            strTableName = "[2800_BelegArchiv_Auf]"

                        End Select

1860                    Call CheckBelegID(lngBelegID, strTableName)
                        
                        'Call CheckBelegID(lngBelegID, IIf(GintBelegArt = 0, "[2800_BelegArchiv_Rng]", "[2800_BelegArchiv_Gut]")) 'DH, 18.01.2016, 6.4.117, Ebenfalls pruefen ob die BelegID noch nicht im BelegArchiv existiert
                        '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
                        
1865                    If lngBelegID <> BelegID Then                           'DH, 19.01.2016, 6.4.118, Wenn sich die BelegID geaendert hat, muss ein Update auf die Sonderfaktura Tabelle erfolgen

1870                        Set connSF = New ADODB.Connection
  
1875                        connSF.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
1880                        connSF.Open
1885                        connSF.Execute "UPDATE [2800_Folge] SET BelegID = '" & lngBelegID & "' WHERE BelegID = '" & BelegID & "'"
1890                        connSF.Close

1895                        rsH.Fields("BelegID").value = lngBelegID            'Der Haupt-Datensatz kann direkt ueber das Recordset aktualisiert werden
1900                        rsH.Update
                           
                        End If

                        'CSBmk <PDF-ARCHIVIERUNG>
                        
1905                    If BelegArt = 0 Then

1910                        Call BelegAnAusgangsbuch(lngBelegID, E_DATATYPE.Sonderfaktura_Rechnung)                         'DH, 28.10.2015, 6.4.110, Vor dem Archivieren, Beleg an Ausgangsbuch uebergeben

1915                        Call BelegAnOpUebergeben(lngBelegID)                'GW 12.3.2021, Ver.: 6.6.110
                            
1920                        Call ArchivierenPDF(LL1, "SFR", BelegNr, rsH, rs)   'HW 10.06.2013
                           
                        Else

1925                        Call BelegAnAusgangsbuch(lngBelegID, E_DATATYPE.Sonderfaktura_Gutschrift)

1930                        Call BelegAnOpUebergeben(lngBelegID)                'GW 12.3.2021, Ver.: 6.6.110

1935                        Call ArchivierenPDF(LL1, "SFG", BelegNr, rsH, rs)   'HW 10.06.2013
                           
                        End If
                    
1940                    If BelegArt = 0 Then
1945                        currentDocType = E_DATATYPE.Sonderfaktura_Rechnung
                           
                        Else
1950                        currentDocType = E_DATATYPE.Sonderfaktura_Gutschrift
                           
                        End If
                        
                        'CSBmk <EMAIL-VERSAND>
                        '<Modified by: GW at 21.02.2020, Ver.: GOBD >
              
1955                    If emailActivated(rsH.Fields("MCode").value, CInt(currentDocType)) Then           'DH, 21.12.2015, 6.4.114, Wenn der eMail-Versand aktiviert ist (Mandanten-/Kundenstamm)
                      
1960                        If objEmailSending Is Nothing Then Set objEmailSending = New clsEmailSending
                   
1965                        If idCollection Is Nothing Then Set idCollection = New Collection
1970                        idCollection.Add lngBelegID                                                   'Die aktuelle BelegID der Collection hinzufuegen
                           
                        End If

                        '</Modified by: GW at 21.02.2020, Ver.: GOBD >
                       
                    Else
                    
                        'CSBmk <VORSCHAU ANZEIGEN>
1975                    glRet = LL1.LlPreviewDisplay(ArbeitsplatzPfad & "\SP52800.LL", "", Me.hwnd)
                       
                    End If

                End If
     
1980            If Save Then                                                    'HW 17.10.2013

                    'CSBmk <TEMP DATEI LÖSCHEN>
1985                LL1.LlPreviewDeleteFiles ArbeitsplatzPfad & "\SP52800.LL", ""

1990                LL1.LlPrintEnd (0)

1995                DruckIni = False

2000                GoSub NextRecord
                    
                End If
                
            End If
            
        End If
  
2005 Beenden:                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                'HW 17.10.2013 eingebaut um jede Rechnung zu Archivieren!
 
2010    If Not idCollection Is Nothing Then             'DH, 19.01.2016, 6.4.118, Nach dem der Druck beendet ist, eMail versenden

2015        If idCollection.Count > 0 Then              'EMail-Versand nur starten, wenn auch BelegIDs zum Senden vorhanden sind

                '<Modified by: GW at 21.02.2020, Ver.: GOBD_EMAIL >

2020            Call objEmailSending.StartEmailSending(cReSize.CurrScaleFactorHeight, cReSize.CurrScaleFactorWidth, voll_automatik, currentDocType, idCollection)

                '</Modified by: GW at 21.02.2020, Ver.: GOBD_EMAIL >
                
            End If
            
        End If

        'HW 17.10.2013
        '##################
        On Error Resume Next

2025    rs.Close
        
2030    If Err.number <> 0 Then Err.Clear

2035    Set rs = Nothing
2040    rsH.Close
        
2045    If Err.number <> 0 Then Err.Clear
        '##################
  
2050    Screen.MousePointer = 0
        
2055    If MsgBoxText <> "" Or SperrText <> "" Or NrKreisText <> "" Then

2060        If MsgBoxText <> "" Then
        
2065            If SperrText <> "" Then

2070                MsgBoxText = MsgBoxText & vbCrLf & vbCrLf & SperrText
                    
                End If
               
            Else
           
2075            If SperrText <> "" Then

2080                MsgBoxText = SperrText
                    
                End If
                
            End If
    
2085        If MsgBoxText <> "" Then

2090            If NrKreisText <> "" Then

2095                MsgBoxText = MsgBoxText & vbCrLf & vbCrLf & NrKreisText
                    
                End If

            Else

2100            If NrKreisText <> "" Then

2105                MsgBoxText = NrKreisText
                    
                End If
                
            End If
    
2110        MsgBox MsgBoxText, vbInformation
            
        End If

        Exit Function

Fehler:
2115    LLPrint = Err.number

2120    If IsUpdateError(Err.number) Then

2125        If SperrText = "" Then

2130            SperrText = "Folgende Belege wurden nicht gedruckt da sie von anderen Benutzern bearbeitet waren:"
2135            SperrText = SperrText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
                
            Else
2140            SperrText = SperrText & vbCrLf & rsH!BelegNr & " - " & rsH!Name1 & " Benutzer: " & rsH!AendVon & " am " & rsH!AendDat
                
            End If

2145        If Not rsH Is Nothing Then rsH.Close

2150        LLPrint = 0
2155        GoTo NextRecord
     
        Else
        
2160        HauptRS.MoveFirst
2165        Screen.MousePointer = 0

2170        If glRet <> 0 Then

2175            Call FehlerErklärung("frmSP52850", "LLPrint LLFehler: " & glRet)
                
            Else
             
2180            Call FehlerErklärung("frmSP52850", "LLPrint")
                
            End If
            
        End If

End Function

Public Function AusDBLesen() As Boolean

        On Error GoTo Fehler
   
        Dim rs        As New ADODB.Recordset

        Dim sql       As String

        Dim Datei     As String

        Dim exist     As Boolean

        Dim PanelText As String

        Dim Felder    As String
        
100     Felder = objPRM.GetUseFields("TDBG1")

105     If left(Felder, 7) = "Status," Then

            'Das Feld Status wird erst später eingefügt.
110         Felder = right(Felder, Len(Felder) - 7)

        End If
        
115     Felder = Replace$(Felder, ",ERechnungArtGK", "")                        'DF 17.12.2024 , Ver.: 6.7.103 : Die Temporäre Spalte wird erst unten mit JOIN aus Kunden Grudnkonditionen geholt.
  
        'In ErstDat ist Datum und Uhrzeit gespeichert. Die Uhrzeit muss ausgefiltert werden
        'um die richtige Filterung zu gewährleisten.
        'Felder = Replace(Felder, "ErstDat", "CDate(Format([ErstDat],'dd/mm/yy')) AS ErstDat")
120     Felder = Replace(Felder, "ErstDat", "CONVERT(nvarchar(10), ErstDat, 104) AS ErstDat")

125     Felder = Replace(Felder, "DatumDateTime", "ErstDat AS DatumDateTime")                     'IL 03.12.2024 , Ver.: 6.7.102 : Fügen eine weitere Spalte zur Verwendung beim Sortieren nach Datum hinzu

130     Felder = Replace(Felder, "BelegID", "CONVERT(nvarchar(10), BelegID, 104) AS BelegID")     'IL 03.12.2024 , Ver.: 6.7.102 : notwendig für eine bequeme Filterung
  
135     sql = "SELECT " & Felder & " FROM [2800_Haupt] WHERE Storno = '0' AND Druck = 0 AND ZwAblage = 0 AND Art = " & gintPrivBelegArt
        
140     sql = "SELECT HAUPT.*, GRUND.ERechnungArt AS ERechnungArtGK FROM (" & sql & ") AS HAUPT LEFT JOIN [1200_GrundKonditionen] AS GRUND ON HAUPT.MCODE = GRUND.MCODE "   'DF 17.12.2024 , Ver.: 6.7.103 : aktuelles E-Beleg Parameter aus Kundenstamm für jeden Beleg holen.
        
145     rs.Open sql, gConn, adOpenStatic, adLockReadOnly

150     If Not gRS Is Nothing Then

155         gRS.Close
160         Set gRS = Nothing

        End If
        
165     Set gRS = New ADODB.Recordset
170     Set gRS = RsOhneVerbindung(rs, "Status")                                'Das Feld Status wurde an den RS angehängt.
    
175     rs.Close

180     gRS.Sort = "Name1"
    
185     Set TDBG1.DataSource = gRS
        
        Exit Function

Fehler:
190     Call FehlerErklärung("frmSP52850", "AusDBLesen()")

End Function

Private Sub cmd1_Click(Index As Integer)

        Dim iFileNr      As Integer

        Dim StartVersuch As Integer

        Dim Zeichen      As String

        Dim BelegArt     As String
  
        On Error GoTo Fehler
        
        '<Removed by: IL at: 29.11.2024, Ver.: 6.7.101 >
        '# Die Vorschaufunktion und Druckfunktion wurden in die Dropdown-Liste verschoben

        '        '<Added by: GW at: 17.03.2020, Ver.: GOBD >
        '        'Wenn auf GOBD umgestellt wurde, muss zuerst geprüft werden, ob
        '        'Verbindung zu GOBD-Datenbank besteht, und erst dann mit Fakturierung
        '        'fortfahren.
        '100     If Index = 4 Or Index = 5 Then
        '
        '105         modGOBD.IsGOBDArchiveActive = GOBDActive(GsHauptPfad) 'GW GOBD
        '110         modGOBD.mandantenNr = GsAnwenderNr
        '115         Call modGOBD.GetGOBDPath(GsHauptPfad)
        '
        '120         If modGOBD.IsGOBDArchiveActive Then
        '125             If modGOBD.ConnectionExists = False Then
        '130                 Call Logbuch("Keine Verbindung zu GOBD-Datenbank")
        '135                 MsgBox GetMessage(2150), vbExclamation, strMeldungCap
        '
        '                    Exit Sub
        '
        '                End If
        '            End If
        '        End If
        '
        '        '</Added by: GW at: 17.03.2020, Ver.: GOBD >
        '</Removed by: IL at: 29.11.2024, Ver.: 6.7.101 >
        
100     showDruckOptionen = True 'HW 17.10.2013
  
105     Select Case Index

            Case 0 'alle
            
115             SetStatus True

120         Case 1 'keine

125             SetStatus False

                '<Removed by: IL at: 29.11.2024, Ver.: 6.7.101 >
                '# Funktionen wurden in eine separate Funktionen verschoben

                '155         Case 4 'Vorschau
                '
                '160             TDBG1.Update
                '165             LlPrint 2

                '170         Case 5 'Drucken

                '175             TDBG1.Update
                '
                '                'Sicherstellen, dass Sammeldruck nur von einem Arbeitsplatz ausgeführt werden kann.
                '                'Wenn die Datei bereits geöffnet ist, wird die Errorbehandlung ausgelöst.
                '180             iFileNr = FreeFile
                '185             Open GsHauptPfad & "dat\" & CStr(CInt(GsAnwenderNr)) & "\sp52800lock.dat" For Output Lock Read Write As #iFileNr
                '190             sta1.Panels(3).text = ""
                '
                '195             If GbArchiv Then
                '
                '                Else
                '
                '200                 Screen.MousePointer = 11
                '
                '                    'DH, 08.10.2013, Schoen, dass hier soviele Kommentare gemacht wurden.
                '                    '                Was hat es mit dem Archivieren auf sich und warum muessen dazu die Daten neu geladen werden
                '                    '                wodurch die Haekchen wiederum alle neu gesetzt werden
                '205                 If LlPrint(1) = 0 Then
                '
                '                        'Archivieren. Hier werden Belege zusammen archiviert.
                '210                     sta1.Panels(3).text = "" & GetZusatzText("ZusatzTexte", 1401) 'HW 18.10.2013
                '
                '215                     LlPrint 2, True
                '
                '220                     sta1.Panels(3).text = "" 'HW 18.10.2013
                '
                '225                     AusDBLesen
                '
                '230                     gRS.filter = getFilter()
                '
                '235                 ElseIf -99 Then
                '
                '240                     Debug.Print "Abbruch"
                '
                '                    End If
                '
                '245                 gblnBelegNrChecked = False                                  'DF 07.02.2019 , Ver.: 6.5.109 : Zeiger bereits durchgeführte Überprüfung zurücksetzen
                '
                '250                 Screen.MousePointer = 0
                '
                '                End If
                '
                '255             Close #iFileNr

                '</Removed by: IL at: 29.11.2024, Ver.: 6.7.101 >
                
130         Case 5

                '<Added by: IL at: 29.11.2024, Ver.: 6.7.101 >
                '# Öffnen ein Menü mit der Auswahl zwischen Drucken und Vorschau (der ursprüngliche Code von „Drucken“ und „Vorschau“ wurde in separate Funktionen verschoben)
                
135             Me.PopupMenu mnuBearb1(1), , cmd1(Index).left, cmd1(Index).top + cmd1(Index).height + SSPanel1(0).top
                '</Added by: IL at: 29.11.2024, Ver.: 6.7.101 >

140         Case 6 'Schließen

145             Unload Me

        End Select
  
        Exit Sub

Fehler:

150     Select Case Err.number

            Case 55, 70 'Datei bereits geöffnet, Zugrif verweigert, Datei konnte nicht gesperrt werden.
155             StartVersuch = StartVersuch + 1

160             If StartVersuch < 6 Then

165                 Select Case StartVersuch

                        Case 1
170                         Zeichen = "|"

175                     Case 2
180                         Zeichen = "/"

185                     Case 3
190                         Zeichen = "--"

195                     Case 4
200                         Zeichen = "\"

205                     Case 5
210                         Zeichen = "--"
                    End Select

215                 sta1.Panels(3).text = "Warten auf Druckfreigabe " & Zeichen
220                 Sleep 1

225                 Resume 0

                Else

230                 If MsgBox("Auf einem anderen Arbeitsplatz wird der Sammeldruck ausgeführt.", vbInformation + vbRetryCancel) = vbRetry Then
235                     StartVersuch = 0
240                     Sleep 1

245                     Resume 0

                    Else
250                     sta1.Panels(3).text = ""
255                     Screen.MousePointer = 0
                    End If
                End If

260         Case Else
265             Screen.MousePointer = 0
270             Call FehlerErklärung("frmSP52850", "cmd1_Click")
        End Select
  
End Sub

Private Sub cmdAuswahl_Click(Index As Integer)

        On Error GoTo Fehler

        Dim dt  As Date

        Dim erg As Boolean         'GW

100     Select Case Index

            Case 0, 1
                'DH, 30.05.2017,  6.4.126, Zur Datumsauswahl wird jetzt die SPKalender.dll verwendet
                '30        Datumausw.BuddyHWnd = txt1(Index).hWnd
105             g_objCal.BuddyHWnd = txt1(Index).hwnd
                '40        If IsDate(txt1(Index)) Then
                '50          dt = CDate(txt1(Index))
                '60        Else
                '70          dt = Date
                '80        End If

                '90        If (Datumausw.Show(dt)) Then
110             If g_objCal.GetData(dt) Then
                    '100         txt1(Index) = Datumausw.Value
115                 txt1(Index).text = dt
120                 gRS.filter = getFilter()
                    'Call objPRM.SprungNeu("Vorwärts", False, txt1(Index).TabIndex, True)

                    'GW_05.03.2018 Ver.6.5.106 Datumsvalidierung ------------------------------------

125                 erg = DatumUnterschiedRek(2, txt1(Index), txt1(0), 90, "d")

130                 If erg = True Then
135                     txt1(Index) = ""

140                     If txt1(Index).Enabled Then txt1(Index).SetFocus
                    Else

145                     If objPRM.EingabeFehler(txt1(Index)) = False Then
150                         Call objPRM.SprungNeu("Vorwärts", False, txt1(Index).TabIndex, True)
                        End If
                    End If

                    'GW-----------------------------------------------------------------------------
                Else
155                 txt1(Index).SetFocus
                End If

        End Select

        Exit Sub

Fehler:
160     Call FehlerErklärung("frmSP52850", "cmdAuswahl_Click()")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        Dim AltDown

        Const vbAltMask = 4
                
        '<Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
100     If Shift = 1 Then
105         shiftPressed = True
        End If

        '</Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
  
        'Beim clicken auf das Programm-Symbol im Register Aktive Programme (SP51.ProgrammAktivieren)
        'wird SendKeys "%{F12}", True - Befehl ausgeführt.
        'Aktuelle form wird zu aktiven Form in normaler Größe.
110     AltDown = (Shift And vbAltMask) > 0

115     If KeyCode = vbKeyF12 Then
120         If AltDown Then
125             Call Main
            End If
        End If
  
130     If Shift = 2 Then
135         GStrg = True
        End If

        '***Beginn
        Exit Sub

Fehler:
140     Call FehlerErklärung("frmSP52850", "Form_KeyDown")
        '***Ende

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        '<Added by: IL at: 03.12.2024, Ver.: 6.7.102 >
100     If printJobInProgress Then

105         Cancel = True

            Exit Sub

        End If

        '</Added by: IL at: 03.12.2024, Ver.: 6.7.102 >

        '####### Subclassing: Messages austragen #############
        'DeW, Mai 2011
110     DetachMessage Me, Me.hwnd, WM_EXITSIZEMOVE
115     DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
        '
        '#####################################################

        '####### Subclassing: Groessenbegrenzung loeschen #######
        'DeW, Mai 2011
120     RemoveMinMaxInfo Me.hwnd
        '
        '########################################################

        'HW 10.07.2014
125     Call writeWindowPos(Me, "SP52800", "SP52850" & gintPrivBelegArt & "Left", "SP52850" & gintPrivBelegArt & "Top")
        '40      If WindowPosition(Me) Then
        '50        SaveSetting "SP50000", "SP52800", "SP52850" & gintPrivBelegArt & "Left", Me.left
        '60        SaveSetting "SP50000", "SP52800", "SP52850" & gintPrivBelegArt & "Top", Me.Top
        '70      End If

130     Me.Visible = False
135     Me.Hide
140     DoEvents
145     Sleep (0.5)

        'Unterrutine in SP50000B.bas
150     If gintPrivBelegArt = 0 Then
155         Call ProgrammAus("285")
160         Protokoll iAppend, vbCrLf & "Programm beendet: 285 -> " & Now & vbCrLf & "-----"
        Else
165         Call ProgrammAus("286")
170         Protokoll iAppend, vbCrLf & "Programm beendet: 286 -> " & Now & vbCrLf & "-----"
        End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

        '***Beginn
        On Error GoTo Fehler
  
100     If GintBelegArt = 0 Then
105         SaveSetting "SP50000", App.EXEName, "SP62850_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
        Else
110         SaveSetting "SP50000", App.EXEName, "SP62860_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
        End If

        '***Ende
        '########## Formular Resizing: stoppen###################
        '
        'DeW, folgendes terminiert die Klasse, und loest dort
        'das _Terminate Ereigniss aus -> Speicherung der eingestellten
        'Vergroesserungswerte und Spaltenbreiten aus den TrueDBGrid
        'Info-Daten in der Registry
        '
115     Set cReSize = Nothing
        '
        '########################################################

        'HW 26.07.2013
        '########################################################
120     Set objPRM = Nothing
125     Set objTDBG = Nothing
130     Set objHlp = Nothing
135     Set objERechnung = Nothing

140     gRS.Close

145     If Err.number <> 0 Then Err.Clear
150     Set gRS = Nothing
        '########################################################

155     DisposeObjects Me                                                       'HW 26.07.2013

        Exit Sub

Fehler:
160     Call FehlerErklärung("frmSP52850", "Form_Unload()")
End Sub

Public Function RsOhneVerbindung(QuellRs As ADODB.Recordset, _
                                 Optional StatusFeld As String, _
                                 Optional OhneInhalt As Boolean) As ADODB.Recordset

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        'Die Funktion erzeugt ein verbindungsloses Recordset mit der Struktur und Inhalt des QuellRs. (Wenn OhneInhalt = False)
        'StatusFeld wird eingefügt als numerisches Feld mit dem Standardwert = 0
        Dim i        As Integer

        Dim Bookmark As Variant

        Dim text     As String
        
100     Set RsOhneVerbindung = New ADODB.Recordset

105     For i = 0 To QuellRs.Fields.Count - 1

110         If QuellRs(i).name = "BelegNr" Then
115             RsOhneVerbindung.Fields.Append QuellRs.Fields(i).name, adVarWChar, 20, adFldIsNullable
            Else
120             RsOhneVerbindung.Fields.Append QuellRs.Fields(i).name, QuellRs.Fields(i).Type, QuellRs.Fields(i).DefinedSize, adFldIsNullable
            End If

125     Next i
        
130     If StatusFeld <> "" Then
135         RsOhneVerbindung.Fields.Append StatusFeld, adSmallInt
        End If
        
140     RsOhneVerbindung.Open
        
145     If OhneInhalt Then Exit Function
        
150     If QuellRs.RecordCount > 0 Then
155         Bookmark = QuellRs.Bookmark
160         QuellRs.MoveFirst
        End If
        
165     Do Until QuellRs.EOF
170         RsOhneVerbindung.AddNew

175         For i = 0 To QuellRs.Fields.Count - 1

180             If QuellRs(i).Type = adVarWChar Then
185                 RsOhneVerbindung.Fields(i).value = "" & QuellRs.Fields(i).value
                Else
190                 RsOhneVerbindung.Fields(i).value = QuellRs.Fields(i).value
                End If

195             If QuellRs(i).name = "BelegNr" Then
200                 If RsOhneVerbindung.Fields(i).value = 0 Then RsOhneVerbindung.Fields(i).value = "neu"
                End If

205         Next i

210         If StatusFeld <> "" Then
215             RsOhneVerbindung.Fields(i).value = -1
            End If
          
220         RsOhneVerbindung.Update
225         QuellRs.MoveNext
        Loop
        
230     If QuellRs.RecordCount > 0 Then
235         RsOhneVerbindung.MoveFirst
240         QuellRs.Bookmark = Bookmark
        End If
        
        '***Beginn
        Exit Function

Fehler:
245     Call FehlerErklärung("SP52800B", "RsOhneVerbindung")
        '***Ende
End Function

Private Sub gRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, _
                             ByVal pError As ADODB.error, _
                             adStatus As ADODB.EventStatusEnum, _
                             ByVal pRecordset As ADODB.Recordset)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     If adReason = adRsnRequery Then
105         If gRS.RecordCount = 0 Then
110             Schalter False
            Else
115             Schalter True
            End If
        End If

        '***Beginn
        Exit Sub

Fehler:
120     Call FehlerErklärung("frmSP52850", "gRs_MoveComplete")
        '***Ende

End Sub

Public Sub Schalter(Enabled As Boolean)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     cmd1(0).Enabled = Enabled
105     cmd1(1).Enabled = Enabled
110     mnuBearb1(2).Enabled = Enabled
115     mnuBearb1(3).Enabled = Enabled
        '120     cmd1(4).Enabled = Enabled                                              'IL 29.11.2024 , Ver.: 6.7.101 : Die Vorschaufunktion wurde in die Dropdown-Liste verschoben
120     cmd1(5).Enabled = Enabled
        '130     mnuBearb1(0).Enabled = Enabled                                         'IL 29.11.2024 , Ver.: 6.7.101 : Die Vorschaufunktion wurde in die Dropdown-Liste verschoben
125     mnuBearb1(1).Enabled = Enabled

130     If GbDesigner Then mnuBearb1(5).Enabled = Enabled

        '***Beginn
        Exit Sub

Fehler:
135     Call FehlerErklärung("frmSP52831", "Schalter")
        '***Ende
End Sub

Private Sub mnuAusgabe_Click(Index As Integer)

        On Error GoTo Fehler
        
100     Call Ausgabe(Index)
        
        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52850", "mnuAusgabe_Click()")
End Sub

Private Sub mnuBearb1_Click(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     Select Case Index

                '<Removed by: IL at: 29.11.2024, Ver.: 6.7.101 >
                '# Funktionen wurden in eine separate Dropdown-Liste verschoben
                '            Case 0 'Vorschau
                '105             Call cmd1_Click(4)

                '110         Case 1 'Drucken
                '115             Call cmd1_Click(5)
                '</Removed by: IL at: 29.11.2024, Ver.: 6.7.101 >

            Case 2 'Alle
105             Call cmd1_Click(0)

110         Case 3 'Keine
115             Call cmd1_Click(1)
                '  Case 0, 1
                '    cmd1_Click (Index)
        End Select

        '***Beginn
        Exit Sub

Fehler:
120     Call FehlerErklärung("frmSP52850", "mnuBearb1_Click")
        '***Ende
End Sub

Private Sub mnuDat1_Click(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     Select Case Index

            Case 6
105             cmd1_Click (Index)
                '  Case 4, 5, 6
                '    cmd1_Click (Index)
        End Select

        '***Beginn
        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52850", "mnuDat1_Click")
        '***Ende
End Sub

Private Sub mnuDesign_Click(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     LLDesigner Me, LL1, gRS!BelegID, Index

        '***Beginn
        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52850", "mnuDesign_Click")
        '***Ende
End Sub

Private Sub mnuInfo_Click(Index As Integer)
   
        '<Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
        On Error GoTo Fehler
    
100     If shiftPressed Then
105         objHlp.HlpShow HlpWrite, "InfoSammeldruck"
        Else
110         objHlp.HlpShow HlpRead, "InfoSammeldruck"
        End If

115     shiftPressed = False
120     objPRM.FindFirstString = "name = '" & Me.ActiveControl.name & "' AND index = " & Me.ActiveControl.Index
   
        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52850", "mnuInfo_Click()")
        '</Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
    
End Sub

Private Sub mnuOpt1_Click(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
  
100     mnuOpt1(Index).Checked = Not mnuOpt1(Index).Checked

105     Select Case Index

            Case 0 'Druckerauswahl
110             SaveSetting "SP50000", "SP52800", "SP52850DruckerDialog", CInt(mnuOpt1(Index).Checked)
        End Select
  
        '***Beginn
        Exit Sub

Fehler:
115     Call FehlerErklärung("frmSP52850", "mnuOpt1_Click")
        '***Ende
End Sub

Private Sub mnuUpdateInfo_Click(Index As Integer)
  
        '<Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
        On Error GoTo Fehler
   
100     If shiftPressed = True Then

105         objHlp.HlpShow HlpWrite, "UpdateAenderungSammeldruck285"
110         shiftPressed = False
        Else
115         objHlp.HlpShow HlpRead, "UpdateAenderungSammeldruck285"
        End If

120     objHlp.UpdateAnzeigen = False
   
        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52850", "mnuUpdateInfo_Click()")
        '</Added by: GW at: 11.02.2019, Ver.: 6.5.109 >
   
End Sub

Private Sub TDBG1_BeforeColEdit(ByVal colIndex As Integer, _
                                ByVal KeyAscii As Integer, _
                                Cancel As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
  
100     If colIndex > 0 Then Cancel = True

        '***Beginn
        Exit Sub
Fehler:
105     Call FehlerErklärung("frmSP52850", "TDBG1_BeforeColEdit")
        '***Ende
End Sub

Private Sub TDBG1_FilterChange()

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        Dim c As Integer

        'Sperren der Bildschirmausgabe während die Form verendert wird.
        'LockWindowUpdate (Me.hwnd)
100     c = TDBG1.Col
105     TDBG1.HoldFields
110     gRS.filter = getFilter()
115     TDBG1.Col = c
120     TDBG1.EditActive = True
        'Entsperren der Bildschirmausgabe während die Form verendert wird.
125     LockWindowUpdate (0&)

        '***Beginn
        Exit Sub

Fehler:
        'Entsperren der Bildschirmausgabe während die Form verendert wird.
130     LockWindowUpdate (0&)
135     Call FehlerErklärung("frmSP52850", "TDBG1_FilterChange")
        '***Ende

End Sub

Private Function getFilter() As String

        On Error GoTo Fehler

        'Creates the SQL statement in adodc1.recordset.filter
        'and only filters text currently. It must be modified to filter other data types.
  
        Dim tmp      As String

        Dim N        As Integer

        Dim operator As String
  
100     For Each Col In cols

105         If Trim(Col.FilterText) <> "" Then
110             If InStr(1, Col.FilterText, Chr(34)) = 0 And InStr(1, Col.FilterText, Chr(39)) = 0 Then
                    ' " und ' müssen ausgeschlossen werden. (Um SQL-Fehler zu vermeiden.)
115                 N = N + 1

120                 If N > 1 Then
125                     operator = " AND "
                    End If
        
130                 Select Case gRS.Fields(Col.DataField).Type

                        Case adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
135                         tmp = tmp & operator & "[" & Col.DataField & "] LIKE '" & Col.FilterText & "*'"

140                     Case adDate, adDBDate, adDBTime

145                         If InStr(1, Col.FilterText, ">=") = 1 Then
150                             If IsDate(Mid(Col.FilterText, 3)) Then
155                                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 3)
                                End If

160                         ElseIf InStr(1, Col.FilterText, "=>") = 1 Then

165                             If IsDate(Mid(Col.FilterText, 3)) Then
170                                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 3)
                                End If

175                         ElseIf InStr(1, Col.FilterText, "=<") = 1 Then

180                             If IsDate(Mid(Col.FilterText, 3)) Then
185                                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 3)
                                End If

190                         ElseIf InStr(1, Col.FilterText, "<=") = 1 Then

195                             If IsDate(Mid(Col.FilterText, 3)) Then
200                                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 3)
                                End If

205                         ElseIf InStr(1, Col.FilterText, "<>") = 1 Then

210                             If IsDate(Mid(Col.FilterText, 3)) Then
215                                 tmp = tmp & operator & "[" & Col.DataField & "] <> " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 3)
                                End If

220                         ElseIf InStr(1, Col.FilterText, ">") = 1 Then

225                             If IsDate(Mid(Col.FilterText, 2)) Then
230                                 tmp = tmp & operator & "[" & Col.DataField & "] > " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 2)
                                End If

235                         ElseIf InStr(1, Col.FilterText, "<") = 1 Then

240                             If IsDate(Mid(Col.FilterText, 2)) Then
245                                 tmp = tmp & operator & "[" & Col.DataField & "] < " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 2)
                                End If

250                         ElseIf InStr(1, Col.FilterText, "=") = 1 Then

255                             If IsDate(Mid(Col.FilterText, 2)) Then
260                                 tmp = tmp & operator & "[" & Col.DataField & "] = " & Mid(Format(Col.FilterText, "dd.mm.yyyy"), 2)
                                End If

                            Else

265                             If IsDate(Col.FilterText) Then
270                                 tmp = tmp & operator & "[" & Col.DataField & "] = " & Format(Col.FilterText, "dd.mm.yyyy")
                                End If
                            End If
        
275                     Case Else 'Numerischen Werte

280                         If InStr(1, Col.FilterText, ">=") = 1 Then
285                             If IsNumeric(Mid(Col.FilterText, 3)) Then
290                                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & SQLZahl(Mid(Col.FilterText, 3))
                                End If

295                         ElseIf InStr(1, Col.FilterText, "=>") = 1 Then

300                             If IsNumeric(Mid(Col.FilterText, 3)) Then
305                                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & SQLZahl(Mid(Col.FilterText, 3))
                                End If

310                         ElseIf InStr(1, Col.FilterText, "=<") = 1 Then

315                             If IsNumeric(Mid(Col.FilterText, 3)) Then
320                                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & SQLZahl(Mid(Col.FilterText, 3))
                                End If

325                         ElseIf InStr(1, Col.FilterText, "<=") = 1 Then

330                             If IsNumeric(Mid(Col.FilterText, 3)) Then
335                                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & SQLZahl(Mid(Col.FilterText, 3))
                                End If

340                         ElseIf InStr(1, Col.FilterText, "<>") = 1 Then

345                             If IsNumeric(Mid(Col.FilterText, 3)) Then
350                                 tmp = tmp & operator & "[" & Col.DataField & "] <> " & SQLZahl(Mid(Col.FilterText, 3))
                                End If

355                         ElseIf InStr(1, Col.FilterText, ">") = 1 Then

360                             If IsNumeric(Mid(Col.FilterText, 2)) Then
365                                 tmp = tmp & operator & "[" & Col.DataField & "] > " & SQLZahl(Mid(Col.FilterText, 2))
                                End If

370                         ElseIf InStr(1, Col.FilterText, "<") = 1 Then

375                             If IsNumeric(Mid(Col.FilterText, 2)) Then
380                                 tmp = tmp & operator & "[" & Col.DataField & "] < " & SQLZahl(Mid(Col.FilterText, 2))
                                End If

385                         ElseIf InStr(1, Col.FilterText, "=") = 1 Then

390                             If IsNumeric(Mid(Col.FilterText, 2)) Then
395                                 tmp = tmp & operator & "[" & Col.DataField & "] = " & SQLZahl(Mid(Col.FilterText, 2))
                                End If

                            Else

400                             If IsNumeric(Col.FilterText) Then
405                                 tmp = tmp & operator & "[" & Col.DataField & "] = " & SQLZahl(Col.FilterText)
                                End If
                            End If

                    End Select

                End If
            End If

410     Next Col
  
415     If Trim(tmp) <> "" Then
  
420         If IsDate(txt1(0)) Then
425             If operator = "" Then
430                 operator = " AND "
                End If

435             tmp = tmp & operator & "[ErstDat] >= " & Format(txt1(0), "dd.mm.yyyy")
            End If

440         If IsDate(txt1(1)) Then
445             If operator = "" Then
450                 operator = " AND "
                End If

455             tmp = tmp & operator & "[ErstDat] <= " & Format(txt1(1), "dd.mm.yyyy")
            End If
    
        Else
  
460         If IsDate(txt1(0)) Then
465             tmp = "[ErstDat] >= " & Format(txt1(0), "dd.mm.yyyy")
            End If

470         If Trim(tmp) <> "" Then
475             If IsDate(txt1(1)) Then
480                 tmp = tmp & " AND [ErstDat] <= " & Format(txt1(1), "dd.mm.yyyy")
                End If

            Else

485             If IsDate(txt1(1)) Then
490                 tmp = "[ErstDat] <= " & Format(txt1(1), "dd.mm.yyyy")
                End If
            End If
    
        End If
  
495     getFilter = tmp
  
        Exit Function

Fehler:
500     Call FehlerErklärung("frmSP52850", "getFilter()")
End Function

Private Sub TDBG1_FormatText(ByVal colIndex As Integer, _
                             value As Variant, _
                             Bookmark As Variant)

        On Error GoTo Fehler
    
100     Select Case TDBG1.Columns(colIndex).DataField                           'DF 16.12.2024 , Ver.: 6.7.103
            
            Case "ERechnungArtGK"
            
105             Select Case value
                
                    Case 0
                    
110                     value = strEBelegNone
                    
115                 Case 1
                        
120                     value = strEBelegZUGFeRDE
                        
125                 Case 2
                        
130                     value = strEBelegZUGFeRDS
                        
                End Select
            
        End Select
    
        Exit Sub

Fehler:

135     Me.MousePointer = vbDefault
140     Call FehlerErklärung("frmSP52850", "TDBG1_FormatText()")

End Sub

Private Sub TDBG1_GotFocus()

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     TDBG1.HighlightRowStyle.BackColor = &HC0E0FF  'Farbe der aktiven Zeile (orange).
105     TDBG1.HighlightRowStyle.ForeColor = &H80000002  'Aktive titelleiste vbBlack

        '***Beginn
        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52850", "TDBG1_GotFocus")
        '***Ende
End Sub

Private Sub TDBG1_HeadClick(ByVal colIndex As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        Dim sortCol   As String
        
        Dim strSpalte As String
        
100     strSpalte = TDBG1.Columns(colIndex).DataField
        
105     Select Case UCase(strSpalte)

            Case "ERSTDAT"
            
110             strSpalte = "DatumDateTime"

        End Select
            
115     If gIntSortColIndex = colIndex Then

            '<Modified by: IL at 03.12.2024, Ver.: 6.7.101 >
            '# TDBG1.Columns(gIntSortColIndex).DataField ----> strSpalte

120         Select Case gstrSortOrder

                Case ""
125                 gstrSortOrder = "ASC"
130                 gRS.Sort = strSpalte & " " & gstrSortOrder
135                 TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicture = ImageList1.ListImages(1).Picture
140                 TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPRight

145             Case "ASC"
150                 gstrSortOrder = "DESC"
155                 gRS.Sort = strSpalte & " " & gstrSortOrder
160                 TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicture = ImageList1.ListImages(2).Picture
165                 TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPRight

170             Case "DESC"
175                 TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPTextOnly
                    'gIntSortColIndex = 0
180                 gstrSortOrder = ""
185                 gRS.Sort = ""
            End Select

            '</Modified by: IL at 03.12.2024, Ver.: 6.7.101 >

        Else
190         TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPTextOnly
195         gIntSortColIndex = colIndex
200         gstrSortOrder = "ASC"
205         gRS.Sort = TDBG1.Columns(gIntSortColIndex).DataField & " " & gstrSortOrder
210         TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicture = ImageList1.ListImages(1).Picture
215         TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPRight
        End If
              
        '***Beginn
        Exit Sub

Fehler:

220     If Err.number = -2147217824 Then
            'Die Reihenfolge kann nicht angewendet werden.
225         TDBG1.Columns(gIntSortColIndex).HeadingStyle.ForegroundPicturePosition = dbgFPTextOnly

230         Resume Next

        Else
235         Call FehlerErklärung("frmSP52850", "TDBG1_HeadClick")
        End If

        '***Ende

End Sub

Private Sub TDBG1_KeyDown(KeyCode As Integer, Shift As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     Select Case KeyCode

            Case vbKeyF1

105             If Shift = 1 Then
                    'Die UMSCHALT-TASTE ist gedrückt.
                    'Hilfetexte können erfast oder bearbeitet werden.
110                 objHlp.HlpShow HlpWrite, "TDBG10000"
                Else
115                 objHlp.HlpShow HlpRead, "TDBG10000"
                End If

        End Select

        '***Beginn
        Exit Sub

Fehler:
120     Call FehlerErklärung("frmSP52850", "TDBG1_KeyDown")
        '***Ende
End Sub

Private Sub TDBG1_LostFocus()

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     TDBG1.HighlightRowStyle.BackColor = vbWindowBackground '&H80000005
105     TDBG1.HighlightRowStyle.ForeColor = vbWindowText  'Farbe des Textes in Fenstern

        '***Beginn
        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52850", "TDBG1_LostFocus")
        '***Ende
End Sub

Private Sub txt1_GotFocus(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     txt1(Index).SelStart = 0
105     txt1(Index).selLength = Len(txt1(Index))
110     txt1(Index).ForeColor = vbWindowText
115     txt1(Index).BackColor = &HC0E0FF  'hellorange
120     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
        'txt1(Index).MaxLength = objPRM.EingabeLaenge(txt1(Index))
    
        '***Beginn
        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52850", "txt1_GotFocus")
        '***Ende

End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        '***Beginn
        Dim bCancel As Boolean

        On Error GoTo Fehler

        '***Ende

100     Select Case KeyCode

            Case vbKeyF1

105             If Shift = 1 Then
                    'Die UMSCHALT-TASTE ist gedrückt.
                    'Hilfetexte können erfast oder bearbeitet werden.
110                 objHlp.HlpShow HlpWrite, objPRM.HlpID
                Else
115                 objHlp.HlpShow HlpRead, objPRM.HlpID
                End If

120         Case vbKeyF2

125             Select Case Index

                    Case 0, 1
130                     Call cmdAuswahl_Click(Index)
                End Select

135         Case vbKeyReturn, vbKeyDown
140             objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
                'Weil objPRM.SprungNeu Validate-Ereignis nicht auslöst,
                'muss die Umwandlung und Prüfung an der Stelle stattfinden.
145             txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , False, False) ' GW Jahreshinweis auf false

                'GW_07.03.2018 Ver.6.5.106 Datumsvalidierung----------------------------
150             If Index = 0 Or Index = 1 Then
155                 txt1_Validate Index, bCancel

160                 If bCancel = True Then
165                     KeyCode = 0

                        Exit Sub

                    Else
170                     Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                    End If

                Else
175                 txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , True, True)
                End If

                'GW---------------------------------------------------------------------

180             If objPRM.EingabeFehler(txt1(Index)) = False Then
185                 gRS.filter = getFilter()
190                 Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                End If

195         Case vbKeyEscape, vbKeyUp
200             objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
205             txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , False, True)

210             If objPRM.EingabeFehler(txt1(Index)) = False Then
215                 gRS.filter = getFilter()
220                 Call objPRM.SprungNeu("Rückwerts", Shift, txt1(Index).TabIndex, True)
                End If

        End Select

        '***Beginn
        Exit Sub

Fehler:
225     Call FehlerErklärung("frmSP52850", "txt1_KeyDown")
        '***Ende

End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     Select Case KeyAscii

            Case vbKeyReturn, vbKeyEscape
105             KeyAscii = 0

        End Select
        
110     KeyAscii = objPRM.EingabePrüfung(KeyAscii)
        
115     If txt1(Index).MaxLength = 0 Then

            'Prüfen, ob die neue Zeichenfolge, die zugelassene Länge übersteigt. (Numerischen Felder. Sonst gilt die MaxLength-Eigenschaft von clsPRM)
            Dim NeuText As String

120         If txt1(Index).selLength = 0 Then NeuText = Mid(txt1(Index), 1, txt1(Index).SelStart) & Chr(KeyAscii) & Mid(txt1(Index), txt1(Index).SelStart + 1)

125         If KeyAscii <> 0 And KeyAscii <> 8 Then

130             If Len(NeuText) > objPRM.EingabeLaenge(NeuText) Then
                    'Die neue Zeichenfolge würde zu lang. Eingabe unterdrücken.
135                 Beep
140                 KeyAscii = 0
                End If
                
            End If
            
        End If

        '***Beginn
        Exit Sub

Fehler:
145     Call FehlerErklärung("frmSP52850", "txt1_KeyPress")
        '***Ende

End Sub

Private Sub txt1_LostFocus(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     txt1(Index).ForeColor = vbActiveTitleBar
105     txt1(Index).BackColor = vbWindowBackground  'Fensterhintergrund(weiß)
        
        '***Beginn
        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52850", "txt1_LostFocus")
        '***Ende

End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        'GW_07.03.2018 Ver.6.5.106 Datumsvalidierung-------------------------------------------------------
100     If Index = 0 Or Index = 1 Then
105         txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

            Dim erg As Boolean

110         erg = DatumUnterschiedRek(2, txt1(Index), txt1(0), 90, "d")

115         If erg = True Then
120             txt1(Index) = ""

125             If txt1(Index).Enabled Then txt1(Index).SetFocus
130             Cancel = True
            Else
135             Cancel = False
            End If
        End If

        'GW_07.03.2018-------------------------------------------------------------------------------------

140     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
145     txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , False, False) 'Jahreshinweis auf false gesetzt

150     If objPRM.EingabeFehler(txt1(Index)) Then
155         Cancel = True
        Else
160         gRS.filter = getFilter()
        End If

        '***Beginn
        Exit Sub

Fehler:
165     Call FehlerErklärung("frmSP52850", "txt1_Validate")
        '***Ende

End Sub

Public Sub SetStatus(Wert As Boolean)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        Dim Bookmark As Variant
  
100     If gRS.RecordCount > 0 Then
105         If gRS.BOF Or gRS.EOF Then  'DH, 02.10.2013, 6.2.100
110             gRS.MoveFirst
            End If

115         Bookmark = gRS.Bookmark
120         gRS.MoveFirst

125         Do Until gRS.EOF
130             gRS!status = Wert
135             gRS.Update
140             gRS.MoveNext
            Loop

145         gRS.Bookmark = Bookmark
        End If
  
        '***Beginn
        Exit Sub

Fehler:
150     Call FehlerErklärung("frmSP52850", "SetStatus")
        '***Ende
End Sub

Private Sub mnuAnsicht_ResFak_Click(Index As Integer)

        'DeW 17.03.2011, neu eingefuegt, wegen neuen Ansicht Menues,
        'Schrittweisen Vergroesserung des Formulars
100     Select Case Index

            Case 0  'auf Originale Groesse setzen
105             Call cReSize.ResizeAboutPercent(0#, 0)

110         Case 2  'auf 120 Prozent
115             Call cReSize.ResizeAboutPercent(20#, 2)

120         Case 4  'auf 140 Prozent
125             Call cReSize.ResizeAboutPercent(40#, 4)

130         Case 6  'auf 160 Prozent
135             Call cReSize.ResizeAboutPercent(60#, 6)

140         Case 8  'auf 180 Prozent
145             Call cReSize.ResizeAboutPercent(80#, 8)
        End Select

End Sub

Private Sub mnuAnsicht_ResetPosition_Click()
        'DeW 25.03.2011 neu eingebaut
        '
        '25.03.2011, neu Wunsch, nur das aktuelle Fenster
        'resetten... Daher in jedem Programmfenster einen Menueintrag
        'Unterroutine ResetWindowPos() in SP50000B.bas definiert
        '
100     Call ResetWindowPos(Me.hwnd, "SP51000")
        'DeW, neu, 16.05.2011, Loeschen aller Resize Eintraege,
        'falls Werte einmal falsch gespeichert werden
105     cReSize.RemoveRegistryKeys

End Sub

Private Sub mnuAnsicht_Alle_Click()
        'DeW 17.03.2011, neu eingefuegt, neuer Eintrag im Ansicht Menue
        'fuer eine Aenderung der Groesse in allen Unterformularen
100     mnuAnsicht_Alle.Checked = Not mnuAnsicht_Alle.Checked
105     cReSize.ResizeAllForms = mnuAnsicht_Alle.Checked
End Sub

Private Sub mnuAnsicht_Prop_Click()
        'DeW 17.03.2011, neu eingefuegt, neuer Eintrag im Ansicht Menue
        'fuer eine Proportionale Vergroesserung der Fenster
100     mnuAnsicht_Prop.Checked = Not mnuAnsicht_Prop.Checked
105     cReSize.ScalingProportional = mnuAnsicht_Prop.Checked
        'Trigger Resize Event, to update View
110     cReSize.resize

End Sub

Private Sub Vorschau()

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       Vorschau
        ' Description:       Startet Vorschau
        ' Created by :       IL
        ' Date-Time  :       29.11.2024-14:21:53
        '
        ' Parameters :
        '--------------------------------------------------------------------------------
        
        On Error GoTo Fehler
        
100     TDBG1.Update
105     LLPrint 2
        
        Exit Sub

Fehler:

110     Me.MousePointer = vbDefault
115     Call FehlerErklärung("frmSP52850", "Vorschau()")

End Sub

Private Sub Druck(Optional bAblage As Boolean)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       Druck
        ' Description:       Der Einfachheit halber wurde die Funktion von CMD_Click verschoben. beginnt mit dem Drucken oder Ablage der ausgewählten Dokumente
        ' Created by :       IL
        ' Date-Time  :       29.11.2024-14:24:26
        '
        ' Parameters :       Optional bAblage - Wenn true, werden die ausgewählten Dokumente nur archiviert und nicht gedruckt
        '--------------------------------------------------------------------------------
        
        On Error GoTo Fehler
        
        Dim iFileNr       As Integer

        Dim StartVersuch  As Integer

        Dim Zeichen       As String

        Dim BelegArt      As String
        
        Dim strFilterExt  As String
        
        Dim recCheckBeleg As ADODB.Recordset                                    'DF 17.02.2025 , Ver.: 6.7.102 : Kopie-Recordset zur Überprüfung der Belege.
        
100     TDBG1.Update

        'Sicherstellen, dass Sammeldruck nur von einem Arbeitsplatz ausgeführt werden kann.
        'Wenn die Datei bereits geöffnet ist, wird die Errorbehandlung ausgelöst.
105     iFileNr = FreeFile
110     Open GsHauptPfad & "dat\" & CStr(CInt(GsAnwenderNr)) & "\sp52800lock.dat" For Output Lock Read Write As #iFileNr
115     sta1.Panels(3).text = ""
    
120     If GbArchiv Then

        Else
                
125         Screen.MousePointer = 11
            
            '<Added by: DFiebach at: 12.12.2024, Ver.: 6.7.106 >
            
            'CSBmk <ÜBERPRÜFUNG DER E-BELEG MANDANTEN PFLICHTFELDER>
            
130         Set recCheckBeleg = gRS.Clone(adLockReadOnly)
135         Set recCheckBeleg.ActiveConnection = Nothing

140         strFilterExt = IIf(gRS.filter <> adFilterNone, gRS.filter & " AND " & "ERechnungArtGK <> '0' AND Status = -1", "ERechnungArtGK <> '0' AND Status = -1")

145         If Trim$(strFilterExt) <> "" Then recCheckBeleg.filter = strFilterExt

150         If recCheckBeleg.EOF = False Then

155             recCheckBeleg.MoveFirst

160             If modERechnung.IsEBelegDoc Then

165                 If modMandant.CheckEBelegMandantenFelder(True) = False Then  'Or Not CheckTlb

170                     TDBG1.Refresh

175                     Screen.MousePointer = 0

180                     Close #iFileNr

                        Exit Sub

                    End If

                End If

            End If
            
185         Set recCheckBeleg = Nothing
            
            '</Added by: DFiebach at: 12.12.2024, Ver.: 6.7.106 >
            
            'DH, 08.10.2013, Schoen, dass hier soviele Kommentare gemacht wurden.
            '                Was hat es mit dem Archivieren auf sich und warum muessen dazu die Daten neu geladen werden
            '                wodurch die Haekchen wiederum alle neu gesetzt werden
190         If LLPrint(1, bAblage) = 0 Then

                'Archivieren. Hier werden Belege zusammen archiviert.
195             sta1.Panels(3).text = "" & GetZusatzText("ZusatzTexte", 1401)   'HW 18.10.2013

200             LLPrint 2, True

205             sta1.Panels(3).text = ""                                        'HW 18.10.2013

210             AusDBLesen

215             gRS.filter = getFilter()

220         ElseIf -99 Then
            
225             Debug.Print "Abbruch"
            
            End If
                    
230         gblnBelegNrChecked = False                                          'DF 07.02.2019 , Ver.: 6.5.109 : Zeiger bereits durchgeführte Überprüfung zurücksetzen
                    
235         Screen.MousePointer = 0

        End If

240     Close #iFileNr
        
        Exit Sub

Fehler:

245     Me.MousePointer = vbDefault
250     Call FehlerErklärung("frmSP52850", "Druck()")

End Sub

Private Sub Ausgabe(Index As Integer)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       Ausgabe
        ' Description:       Startet Druck, Vorschau oder Ablage.
        ' Created by :       IL
        ' Date-Time  :       29.11.2024-14:30:14
        '
        ' Parameters :       Index (Integer)
        '                        0 - Vorschau
        '                        1 - Ablage
        '                        2 - Druck
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler
        
100     printJobInProgress = True

105     Call EnableButtonsOverPrintJob(Not printJobInProgress)

        '<Added by: GW at: 17.03.2020, Ver.: GOBD >
        'Wenn auf GOBD umgestellt wurde, muss zuerst geprüft werden, ob
        'Verbindung zu GOBD-Datenbank besteht, und erst dann mit Fakturierung
        'fortfahren.
        
110     Select Case Index
        
            Case 1, 2

115             modGOBD.IsGOBDArchiveActive = GOBDActive(GsHauptPfad) 'GW GOBD
120             modGOBD.mandantenNr = GsAnwenderNr
125             Call modGOBD.GetGOBDPath(GsHauptPfad)

130             If modGOBD.IsGOBDArchiveActive Then
135                 If modGOBD.ConnectionExists = False Then
140                     Call Logbuch("Keine Verbindung zu GOBD-Datenbank")
145                     MsgBox GetMessage(2150), vbExclamation, strMeldungCap

                        Exit Sub

                    End If
                End If
            
        End Select

        '</Added by: GW at: 17.03.2020, Ver.: GOBD >
        
150     showDruckOptionen = True 'HW 17.10.2013

155     Select Case Index

            Case 0
    
160             Call Vorschau
    
165         Case 1

170             If MsgBox(GetMessage(2385), vbYesNo + vbExclamation, strMeldungCap) = vbYes Then Call Druck(True)
    
175         Case 2

180             Call Druck

        End Select
        
185     printJobInProgress = False

190     Call EnableButtonsOverPrintJob(Not printJobInProgress)
        
195     Screen.MousePointer = 0
        
        Exit Sub

Fehler:

200     printJobInProgress = False
205     Call EnableButtonsOverPrintJob(Not printJobInProgress)

210     Me.MousePointer = vbDefault
215     Call FehlerErklärung("frmSP52850", "Ausgabe()")

End Sub

Private Sub EnableButtonsOverPrintJob(blnEnabled As Boolean)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       EnableButtonsOverPrintJob
        ' Description:       [type_description_here]
        ' Created by :       IL
        ' Date-Time  :       03.12.2024-16:26:12
        '
        ' Parameters :       blnEnabled (Boolean)
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler

100     cmd1(0).Enabled = blnEnabled
105     cmd1(1).Enabled = blnEnabled
110     cmd1(5).Enabled = blnEnabled
        
115     mnuBearb1(0).Enabled = blnEnabled
120     mnuBearb1(1).Enabled = blnEnabled
125     mnuBearb1(2).Enabled = blnEnabled
130     mnuBearb1(3).Enabled = blnEnabled
135     mnuBearb1(5).Enabled = blnEnabled
        
        Exit Sub

Fehler:
140     Me.MousePointer = vbDefault
145     Call FehlerErklärung("frmSP52850", "EnableButtonsOverPrintJob()")
End Sub

Private Sub FillPRMValues()

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       FillPRMValues
        ' Description:       PRM Werte holen
        ' Created by :       DFiebach
        ' Date-Time  :       16.12.2024-08:46:28
        '
        ' Parameters :
        '--------------------------------------------------------------------------------
        
        On Error GoTo Fehler

100     strEBelegNone = GetMessage(2361)                                        'E-Beleg-Spalte Bezeichnungen
105     strEBelegZUGFeRDE = GetMessage(2362)
110     strEBelegZUGFeRDS = GetMessage(2363)

        Exit Sub

Fehler:

115     Me.MousePointer = vbDefault
120     Call FehlerErklärung("frmSP52850", "FillPRMValues()")

End Sub

Public Sub InputProcessing(strInput As String)

        On Error GoTo Fehler

100     objERechnung.XMLResult = strInput

        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52831", "InputProcessing()")
End Sub
