VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSP52810 
   Caption         =   "Text-Stamm"
   ClientHeight    =   6075
   ClientLeft      =   4245
   ClientTop       =   2850
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSP52810.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   5805
   Begin VB.TextBox txt1 
      Height          =   2115
      Index           =   0
      Left            =   270
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2940
      Width           =   5310
   End
   Begin TrueOleDBGrid70.TDBGrid TDBG1 
      Height          =   2355
      Index           =   0
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   4154
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
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   -2147483633
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
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   360
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   5280
      Width           =   5800
      _Version        =   65536
      _ExtentX        =   10231
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
         Caption         =   "&Schließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   4477
         TabIndex        =   3
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Drucken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   6660
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "cmd1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   3660
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "cmd1"
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
         Height          =   300
         Index           =   3
         Left            =   7980
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&peichern"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   2650
         TabIndex        =   2
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Löschen"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1340
         TabIndex        =   5
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Neu"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   1260
      End
   End
   Begin MSComctlLib.StatusBar sta1 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      TabIndex        =   10
      Top             =   5730
      Width           =   5805
      _ExtentX        =   10239
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
            Object.Width           =   12541
            MinWidth        =   12541
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDat 
      Caption         =   "Datei"
      Begin VB.Menu mnuDat1 
         Caption         =   "Schließen"
      End
   End
   Begin VB.Menu mnuBearb 
      Caption         =   "Bearbeiten"
      Index           =   0
      Begin VB.Menu mnuBearb1 
         Caption         =   "&Neu"
         Index           =   0
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "&Löschen"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "S&peichern"
         Enabled         =   0   'False
         Index           =   2
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
End
Attribute VB_Name = "frmSP52810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DF 27.01.12 Ver.:6.1.109 ScrollTrack Eigenschaft auf True gesetzt, damit Tabelle sich mit Scrollbalken bewegt
Private objPRM As clsPRM
Private objTDBG As clsTDBG7ole
Private objHlp As SpHlp.clsHlp
Dim CLSADO As New clsADONewGUIDRecord
Private gRsKopf As New ADODB.Recordset
Attribute gRsKopf.VB_VarHelpID = -1
Private gCn As ADODB.Connection
Private Col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private gbInsert As Boolean
Private gbEterNoColChange As Boolean 'Wird bei TDBG_KeyUp abgefragt. (Wenn True, hat betätigen der Enter-Taste kein Column-Wechsel verursacht)
Private gbDataChanged As Boolean 'True -> beim verändern des Textfeldes.
                                 'False -> beim füllen der Daten des aktuellen Datensatzes
                                 'in das Textfeld.

'####### Subclassing ########################
'DeW, Mai 2011
'Variablen notwendig fuer Verwendung der SSubTmr Klasse,
'um eine schoenes Vergroesserung von Fenster und Inhalt und
'Begrenzung der Fenstergroesse zu ermoeglichen!
Implements ISubclass
Private emrConsume As EMsgResponse
'
'DeW, notwendige WM_... Nachrichten fuer das
'Subclassing wurden als Public in SP50000B.bas
'definiert
'############################################


'####### Formular Resizing ##################
'
Dim cResize As FormResize 'HW 03.02.2011
'
'############################################

'############# Subclassing Methoden  ####################
'
Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
    ' This Property Let is not really needed!
  End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
    ' This will tell you which message you are responding to:
    ' Tell the subclasser what to do for this message (here we do all processing):
    ISubClass_MsgResponse = emrConsume
End Property

Private Function ISubClass_WindowProc( _
    ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long _
    ) As Long
    If iMsg = WM_EXITSIZEMOVE Then
        cResize.Resize
    End If

    'emrConsume = emrPostProcess

End Function
'#########################################################






Public Sub AusDBLesen()
                     '***Beginn
10    On Error GoTo Fehler
                     '***Ende
               
        Dim sql As String
        Dim gCn As New ADODB.Connection
        
20      gCn.Mode = adModeUnknown
30      gCn.Provider = "Microsoft.Jet.OLEDB.4.0"
  
        'Isolation und Kapselung der Transaktionen zum Herabsetzen des Transaktionslevels
40      gCn.IsolationLevel = adXactCursorStability
  
50      gCn.CursorLocation = adUseServer
60      gCn.Open GsHauptPfad & "dat\" & CStr(CInt(GsAnwenderNr)) & "\SP50000.dat"
 
70      Set gRsKopf.ActiveConnection = gCn
80      gRsKopf.CursorLocation = adUseServer
90      gRsKopf.CursorType = adOpenKeyset
100     gRsKopf.LockType = adLockOptimistic
  
110     sql = "SELECT * FROM [2800_Texte]"

120     gRsKopf.Open sql
130     If gRsKopf.RecordCount > 0 Then
140       cmd1(1).Enabled = True
150       mnuBearb1(1).Enabled = True
160     End If
  
                     '***Beginn
170           Exit Sub
Fehler:
180           Call FehlerErklärung("frmSP52810", "AusDBLesen")
                     '***Ende
End Sub


Private Function getFilter() As String
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
        'Creates the SQL statement in adodc1.recordset.filter
        'and only filters text currently. It must be modified to filter other data types.
        
        Dim tmp As String
        Dim N As Integer
        Dim operator As String
        
20      For Each Col In cols
30        If Trim(Col.FilterText) <> "" Then
40          If InStr(1, Col.FilterText, Chr(34)) = 0 And InStr(1, Col.FilterText, Chr(39)) = 0 Then
              ' " und ' müssen ausgeschlossen werden. (Um SQL-Fehler zu vermeiden.)
50            N = N + 1
60            If N > 1 Then
70              operator = " AND "
80            End If
              
90            Select Case gRsKopf.Fields(Col.DataField).Type
              Case adBSTR, adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar
100             tmp = tmp & operator & "[" & Col.DataField & "] LIKE '" & Col.FilterText & "*'"
110           Case adDate, adDBDate, adDBTime
120             If InStr(1, Col.FilterText, ">=") = 1 Then
130               If IsDate(Mid(Col.FilterText, 3)) Then
140                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 3)
150               End If
160             ElseIf InStr(1, Col.FilterText, "=>") = 1 Then
170               If IsDate(Mid(Col.FilterText, 3)) Then
180                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 3)
190               End If
200             ElseIf InStr(1, Col.FilterText, "=<") = 1 Then
210               If IsDate(Mid(Col.FilterText, 3)) Then
220                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 3)
230               End If
240             ElseIf InStr(1, Col.FilterText, "<=") = 1 Then
250               If IsDate(Mid(Col.FilterText, 3)) Then
260                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 3)
270               End If
280             ElseIf InStr(1, Col.FilterText, "<>") = 1 Then
290               If IsDate(Mid(Col.FilterText, 3)) Then
300                 tmp = tmp & operator & "[" & Col.DataField & "] <> " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 3)
310               End If
320             ElseIf InStr(1, Col.FilterText, ">") = 1 Then
330               If IsDate(Mid(Col.FilterText, 2)) Then
340                 tmp = tmp & operator & "[" & Col.DataField & "] > " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 2)
350               End If
360             ElseIf InStr(1, Col.FilterText, "<") = 1 Then
370               If IsDate(Mid(Col.FilterText, 2)) Then
380                 tmp = tmp & operator & "[" & Col.DataField & "] < " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 2)
390               End If
400             ElseIf InStr(1, Col.FilterText, "=") = 1 Then
410               If IsDate(Mid(Col.FilterText, 2)) Then
420                 tmp = tmp & operator & "[" & Col.DataField & "] = " & Mid(Format(Col.FilterText, "\#dd\/mm\/yyyy\#"), 2)
430               End If
440             Else
450               If IsDate(Col.FilterText) Then
460                 tmp = tmp & operator & "[" & Col.DataField & "] = " & Format(Col.FilterText, "\#dd\/mm\/yyyy\#")
470               End If
480             End If
              
490           Case Else 'Numerischen Werte
500             If InStr(1, Col.FilterText, ">=") = 1 Then
510               If IsNumeric(Mid(Col.FilterText, 3)) Then
520                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & SQLZahl(Mid(Col.FilterText, 3))
530               End If
540             ElseIf InStr(1, Col.FilterText, "=>") = 1 Then
550               If IsNumeric(Mid(Col.FilterText, 3)) Then
560                 tmp = tmp & operator & "[" & Col.DataField & "] >= " & SQLZahl(Mid(Col.FilterText, 3))
570               End If
580             ElseIf InStr(1, Col.FilterText, "=<") = 1 Then
590               If IsNumeric(Mid(Col.FilterText, 3)) Then
600                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & SQLZahl(Mid(Col.FilterText, 3))
610               End If
620             ElseIf InStr(1, Col.FilterText, "<=") = 1 Then
630               If IsNumeric(Mid(Col.FilterText, 3)) Then
640                 tmp = tmp & operator & "[" & Col.DataField & "] <= " & SQLZahl(Mid(Col.FilterText, 3))
650               End If
660             ElseIf InStr(1, Col.FilterText, "<>") = 1 Then
670               If IsNumeric(Mid(Col.FilterText, 3)) Then
680                 tmp = tmp & operator & "[" & Col.DataField & "] <> " & SQLZahl(Mid(Col.FilterText, 3))
690               End If
700             ElseIf InStr(1, Col.FilterText, ">") = 1 Then
710               If IsNumeric(Mid(Col.FilterText, 2)) Then
720                 tmp = tmp & operator & "[" & Col.DataField & "] > " & SQLZahl(Mid(Col.FilterText, 2))
730               End If
740             ElseIf InStr(1, Col.FilterText, "<") = 1 Then
750               If IsNumeric(Mid(Col.FilterText, 2)) Then
760                 tmp = tmp & operator & "[" & Col.DataField & "] < " & SQLZahl(Mid(Col.FilterText, 2))
770               End If
780             ElseIf InStr(1, Col.FilterText, "=") = 1 Then
790               If IsNumeric(Mid(Col.FilterText, 2)) Then
800                 tmp = tmp & operator & "[" & Col.DataField & "] = " & SQLZahl(Mid(Col.FilterText, 2))
810               End If
820             Else
830               If IsNumeric(Col.FilterText) Then
840                 tmp = tmp & operator & "[" & Col.DataField & "] = " & SQLZahl(Col.FilterText)
850               End If
860             End If
870           End Select
880         End If
890       End If
900     Next Col
        
910     If Trim(tmp) <> "" Then
920       getFilter = tmp
930     End If
        
              '***Beginn
940           Exit Function
Fehler:
950           Call FehlerErklärung("frmSP52712", "getFilter")
              '***Ende
End Function

Private Sub cmd1_Click(Index As Integer)
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
        Dim i As Integer
        
20      Select Case Index
        Case 0 'Neu
30        If gRsKopf.RecordCount > 0 Then
40          For i = 1 To 32000
50            gRsKopf.MoveFirst
60            gRsKopf.Find "Schl = '" & "Text" & CStr(i) & "'"
70            If gRsKopf.EOF Then
                TDBG1(0).HoldFields
80              gRsKopf.AddNew
                'HW 12.07.07 Ver.:5.3.109 hinzugefügt: Klasse clsDAONewGuidRecord eingeführt und dadurch AddNew mit SQL und Access möglich
                CLSADO.SearchAndSetGUID gRsKopf
                
90              gRsKopf!Schl = "Text" & CStr(i)
100             Exit For
110           End If
120         Next
130       Else
140         gRsKopf.AddNew
            'HW 12.07.07 Ver.:5.3.114 hinzugefügt: Klasse clsDAONewGuidRecord eingeführt und dadurch AddNew mit SQL und Access möglich
            CLSADO.SearchAndSetGUID gRsKopf

150         gRsKopf!Schl = "Text1"
160       End If
170       gRsKopf.Update
180     Case 1 'Löschen
190       If gRsKopf.RecordCount > 0 Then
200         Call msgText(1, 13, 0, 0, 0)
210         If MsgBox(GsMsgText(0), vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
              'Wollen Sie den Datensatz löschen?
220           TDBG1(0).Delete
230         End If
240       End If
250     Case 2 'Speichern
260       If gRsKopf.RecordCount > 0 Then
270         TDBG1(0).Update
280       End If
290     Case 6 'Schließen
300       Unload Me
310       Exit Sub
320     End Select
        
330     If gRsKopf.RecordCount = 0 Then
340       cmd1(1).Enabled = False
345       mnuBearb1(1).Enabled = False
350       cmd1(0).SetFocus
360     Else
370       TDBG1(0).SetFocus
380     End If

390     cmd1(2).Enabled = False
395     mnuBearb1(2).Enabled = False
        
              '***Beginn
400           Exit Sub
Fehler:
410           Call FehlerErklärung("frmSP52810", "cmd1_Click")
              '***Ende
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 '***Beginn
On Error GoTo Fehler
 '***Ende
  Dim AltDown
  Const vbAltMask = 4
  
  'Beim clicken auf das Programm-Symbol im Register Aktive Programme (SP51.ProgrammAktivieren)
  'wird SendKeys "%{F12}", True - Befehl ausgeführt.
  'Aktuelle form wird zu aktiven Form in normaler Größe.
  AltDown = (Shift And vbAltMask) > 0
  If KeyCode = vbKeyF12 Then
    If AltDown Then
      Call Main
    End If
  End If

 '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52820", "Form_KeyDown")
 '***Ende

End Sub

Private Sub Form_Load()
  
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
        
      '########### SkinFramework ##############################
      'HW 04.03.2011 - Is ein Windows Skin Tool von Codejock Software
      'SkinFramework1.LoadSkin GsHauptPfad & "\exe\Spedifix.cjstyles", "NormalSilver.ini"
      'SkinFramework1.ApplyWindow Me.hwnd
      '########################################################
  
20      If GsTitel <> "" Then
30        GlSP51000hwnd = FindWindow(vbNullString, GsTitel)
40        SetWindowLong Me.hwnd, GWL_HWNDPARENT, GlSP51000hwnd
50      End If
  
      '60      If left(Me.Name, 3) = "frm" Then
      '70        sta1.Panels(1).Text = Mid(Me.Name, 4)
      '80      Else
      '90        sta1.Panels(1).Text = Me.Name
      '100     End If
60      sta1.Panels(1).Text = "SP62810"

70      sta1.Panels(2).Text = DisplayVerInfo(GsHauptPfad & "exe\" & Gc_strExeFile)
  
80      'Me.Width = 5580 'GetSetting("SP50000", "SP52800", "SP52810Width", "5460")
        Me.Width = 5900 'DH, 16.09.2011, in Folge der Button-Neuausrichtung verbreitert
90      Me.height = 6855 'GetSetting("SP50000", "SP52800", "SP52810Height", "6390")
100     Me.Left = GetSetting("SP50000", "SP52800", "SP52810Left", "6850")
110     Me.Top = GetSetting("SP50000", "SP52800", "SP52810Top", "870")
  
120     SetXPSize Me
  
130     txt1(0).Width = TEXT_BREITE
  
140     Set objPRM = New clsPRM
150     Set objPRM.gForm = Me
160     objPRM.PRM_Alle
  
170     Set objHlp = New SpHlp.clsHlp
180     objHlp.DatabaseName = GsHauptPfad & "hlp\SP50000.hlp"
190     objHlp.Table = Me.Name
200     objHlp.Caption = Me.Name & " - Feldhilfe"
  
210     AusDBLesen
220     Set TDBG1(0).DataSource = gRsKopf
  
230     Set objTDBG = New clsTDBG7ole
240     Set objTDBG.TDBG = TDBG1(0)
250     Set objTDBG.PrmDataBase = GDBprm
  
260     objTDBG.SichtbareZeilen = 8
270     objTDBG.SatzAnzahl = gRsKopf.RecordCount
280     objTDBG.GridParameter ("SELECT * FROM PRM52810 WHERE name = 'TDBG1' AND pos1 = 0 ORDER BY index")
290     Set objTDBG = Nothing
  
300     objPRM.FindFirstString = "name = 'TDBG1' AND pos1 = 9"
310     TDBG1(0).Caption = objPRM.Caption("")
320     TDBG1(0).MarqueeStyle = dbgHighlightCell
330     TDBG1(0).HighlightRowStyle.BackColor = vbWindowBackground '&H80000005
340     TDBG1(0).HighlightRowStyle.ForeColor = vbWindowText  'Farbe des Textes in Fenstern
350     TDBG1(0).AllowColSelect = False
360     TDBG1(0).ExtendRightColumn = False 'Die äußere rechte Spalte wird ans rechte Gridende nicht erweitert.
        'TDBG1(0).Width = TDBG1(0).Width + 300 'Um den H-ScrollBar zu unterdrücken.
370     TDBG1(0).Width = txt1(0).Width
380     TDBG1(0).height = TDBG1(0).height + 320
390     TDBG1(0).FilterBar = True
400     TDBG1(0).HoldFields
410     TDBG1(0).ReBind
420     Set cols = TDBG1(0).Columns
  
430     txt1(0).BackColor = vbWindowBackground '&H80000005
440     txt1(0).ForeColor = vbWindowText  'Farbe des Textes in Fenstern

450     SaveSetting "SP50000", "SP52800", "SP52810", Me.Caption
 
        'DeW 08.08.2011, Testweise eingefuegt, weil Groessen angepasst, wg. Menu
        'und besserem Aussehen!
        SSPanel1(1).Top = sta1.Top - SSPanel1(1).height
        
        '########## Subclassing: Messages festlegen #############
        ' DeW, ZyG Mai 2011
460     AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO 'DeW
470     AttachMessage Me, Me.hwnd, WM_EXITSIZEMOVE 'DeW
        '
        '########################################################
  
        '####### Subclassing: Groessenbegrenzung Formular #######
        ' TODO in Arbeit, MagicNumbers
        'dieser Aufruf kann je nach Programm-Modul woanders in der
        'Form_Load Methode stehen!
        'Zuerst muss der "alte" Code die Zuweisung von Breit und
        'Hoehe korrekt vorgenommen haben!
480     SetMinMaxInfo Me.hwnd, Me.height, (Me.height * 2), Me.Width, (Me.Width * 2)
        '
        '########################################################
         
        '###### Formular Resizing: Parameter setzen#############
        ' DeW, ZyG, Mai 2011
        'Section- oder KeyBezeichnung sind in vielen Faellen in
        'altem Code hart eincodiert worden, manchmal wird auch
        'eine Variable verwendet...
490     Set cResize = New FormResize
500     cResize.setSectionBezeichnung = "SP52810"
510     cResize.setKeyBezeichnung = "SP52810"
520     cResize.setIstUnterFenster = False
        '
        '########################################################
  
  
        '######## Formular Resizing: Formular zuweisen ##########
        ' DeW, Mai 2011
        'Zuweisung von Form erst nach Groessensetzung s.o. Me.Width = ...
        'aber auf jeden Fall nach SetMinMaxInfo ... fuer die
        'Groessenbegrenzung
530     cResize.Form = Me
        '
        'Speichere keine Informationen (Spaltenbreiten usw.) fuer die
        'Tabellen im Form, wenn z.B. nur eine einzelne Tabelle
        'vorhanden ist, die jeweils mit neuen Daten gefuellt
        'und an eine andere Position verschoben wird (z.B. SP51000
        'Mandantenstamm
540     cResize.IgnoreTrueDBGridInfo = True
        ''
        ''########################################################
  
550     cResize.ReResize
  
              '***Beginn
560           Exit Sub
Fehler:
570           Call FehlerErklärung("frmSP52810", "Form_Load")
              '***Ende
  
  

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
  
20      If WindowPosition(Me) Then
30        SaveSetting "SP50000", "SP52800", "SP52810Left", Me.Left
40        SaveSetting "SP50000", "SP52800", "SP52810Top", Me.Top
50        SaveSetting "SP50000", "SP52800", "SP52810Width", Me.Width
60        SaveSetting "SP50000", "SP52800", "SP52810Height", Me.height
70      End If
  
        'Unterrutine in SP50000B
80      Call ProgrammAus("281")
90      Protokoll iAppend, vbCrLf & "Programm beendet: 281 -> " & Now & vbCrLf & "-----"
  
      '####### Subclassing: Messages austragen #############
      'DeW, Mai 2011
100   DetachMessage Me, Me.hwnd, WM_EXITSIZEMOVE
110   DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
      '
      '#####################################################

      '####### Subclassing: Groessenbegrenzung loeschen #######
      'DeW, Mai 2011
120   RemoveMinMaxInfo Me.hwnd
      '
      '########################################################
130   Me.Visible = False
140   Me.Hide
150   DoEvents
160   Sleep (0.5)

              '***Beginn
170           Exit Sub
Fehler:
180           Call FehlerErklärung("frmSP52810", "Form_QueryUnload")
              '***Ende
End Sub


Private Sub Form_Unload(Cancel As Integer)

'########################################################
'########## Formular Resizing: stoppen###################
'
'DeW, folgendes terminiert die Klasse, und loest dort
'das _Terminate Ereigniss aus -> Speicherung der eingestellten
'Vergroesserungswerte und Spaltenbreiten aus den TrueDBGrid
'Info-Daten in der Registry
'
10    Set cResize = Nothing
'
'########################################################

20    Call ProgrammAus("281")

      'HW 26.07.2013
      '########################################################
30    On Error Resume Next

40    Set objPRM = Nothing
50    Set objTDBG = Nothing
60    Set objHlp = Nothing
70    Set CLSADO = Nothing

80    gRsKopf.Close
90    If Err.Number <> 0 Then Err.Clear
100   Set gRsKopf = Nothing

110   gCn.Close
120   If Err.Number <> 0 Then Err.Clear
130   Set gCn = Nothing

140   DisposeObjects Me 'HW 26.07.2013
      '########################################################

End Sub

Private Sub mnuBearb1_Click(Index As Integer)
  'DeW, neu, 27.07.2011, wegen neuem Bearbeiten Menu
        '***Beginn
On Error GoTo Fehler
        '***Ende
  
  cmd1_Click Index
        
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "mnuBearb1_Click")
        '***Ende

End Sub

Private Sub mnuDat1_Click()
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Call cmd1_Click(6)

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "mnuDat1_Click")
        '***Ende

End Sub

Private Sub TDBG1_AfterUpdate(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  cmd1(2).Enabled = False
  mnuBearb1(2).Enabled = False
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_AfterUpdate")
        '***Ende
End Sub

Private Sub TDBG1_BeforeColEdit(Index As Integer, ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  objPRM.FindFirstString = "name = 'TDBG1' AND pos1 = 0 AND index = " & TDBG1(0).Col
  TDBG1(0).Columns(ColIndex).DataWidth = objPRM.EingabeLaenge

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_BeforeColEdit")
        '***Ende
End Sub

Private Sub TDBG1_BeforeInsert(Index As Integer, Cancel As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  gbInsert = True
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_BeforeInsert")
        '***Ende
End Sub


Private Sub TDBG1_BeforeUpdate(Index As Integer, Cancel As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  If gbInsert Then
    TDBG1(0).Columns("ErstVon").Text = GsUser
    gbInsert = False
  End If
  TDBG1(0).Columns("AendVon").Text = GsUser
  TDBG1(0).Columns("AendDat").Text = Now

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_BeforeUpdate")
        '***Ende
End Sub

Private Sub TDBG1_Change(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  cmd1(2).Enabled = True
  mnuBearb1(2).Enabled = True
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_Change")
        '***Ende
End Sub

Private Sub TDBG1_FilterChange(Index As Integer)
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
        Dim c As Integer
        
20      If Index = 0 Then
          'Sperren der Bildschirmausgabe während die Form verendert wird.
30        LockWindowUpdate (Me.hwnd)
40        c = TDBG1(0).Col
50        gRsKopf.Filter = getFilter()
60        TDBG1(0).Col = c
70        TDBG1(0).EditActive = True
          'Entsperren der Bildschirmausgabe während die Form verendert wird.
80        LockWindowUpdate (0&)
90      End If

              '***Beginn
100           Exit Sub
Fehler:
              'Entsperren der Bildschirmausgabe während die Form verendert wird.
110           LockWindowUpdate (0&)
120           Call FehlerErklärung("frmSP52810", "TDBG1_FilterChange")
              '***Ende

End Sub

Private Sub TDBG1_GotFocus(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  TDBG1(Index).HighlightRowStyle.BackColor = &HC0E0FF  'Farbe der aktiven Zeile (orange).
  TDBG1(Index).HighlightRowStyle.ForeColor = &H80000002  'Aktive titelleiste vbBlack

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_GotFocus")
        '***Ende
End Sub

Private Sub TDBG1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
       '***Beginn
10    On Error GoTo Fehler
       '***Ende
        
        
20      Select Case KeyCode
        Case vbKeyF1
30        If Shift = 1 Then
            'Die UMSCHALT-TASTE ist gedrückt.
            'Hilfetexte können erfast oder bearbeitet werden.
40          objHlp.HlpShow HlpWrite, objPRM.HlpID
50        Else
60          objHlp.HlpShow HlpRead, objPRM.HlpID
70        End If
80      Case vbKeyEscape
90        If TDBG1(0).DataChanged = False Then
100         Call objPRM.SprungNeu("Rückwärts", Shift, TDBG1(0).TabIndex)
110       End If
120     Case vbKeyReturn, vbKeyDown
130       gbEterNoColChange = True
140     End Select
       '***Beginn
150           Exit Sub
Fehler:
160           Call FehlerErklärung("frmSP52810", "TDBG1_KeyDown")
       '***Ende

End Sub

Private Sub TDBG1_KeyPress(Index As Integer, KeyAscii As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  
  objPRM.FindFirstString = "name = 'TDBG1' AND pos1 = 0 AND index = " & TDBG1(0).Col
  KeyAscii = objPRM.EingabePrüfung(KeyAscii)

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_KeyPress")
        '***Ende
End Sub


Private Sub TDBG1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  Select Case KeyCode
  Case vbKeyReturn
    If gbEterNoColChange Then
      Call objPRM.SprungNeu("Vorwärts", Shift, TDBG1(0).TabIndex)
    End If
    gbEterNoColChange = False
  End Select

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_KeyUp")
        '***Ende
End Sub

Private Sub TDBG1_LostFocus(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  TDBG1(Index).HighlightRowStyle.BackColor = vbWindowBackground '&H80000005
  TDBG1(Index).HighlightRowStyle.ForeColor = vbWindowText  'Farbe des Textes in Fenstern

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "TDBG1_LostFocus")
        '***Ende
End Sub


Private Sub TDBG1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
20      gbEterNoColChange = False
30      gbDataChanged = False
        
40      If TDBG1(0).bookmark <> LastRow Then
50        If gRsKopf.RecordCount > 0 Then
60          txt1(0).Text = TDBG1(0).Columns("Inhalt").Text
70        Else
80          txt1(0).Text = ""
90        End If
100     End If
          
110     gbDataChanged = True
              '***Beginn
120           Exit Sub
Fehler:
130           Call FehlerErklärung("frmSP52810", "TDBG1_RowColChange")
              '***Ende
End Sub




Private Sub txt1_Change(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  If gbDataChanged Then
    TDBG1(0).Columns("Inhalt") = txt1(0).Text
    If cmd1(2).Enabled = False Then
      cmd1(2).Enabled = True
      mnuBearb1(2).Enabled = True
    End If
  End If
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "txt1_Change")
        '***Ende
End Sub


Private Sub txt1_GotFocus(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  txt1(Index).BackColor = &HC0E0FF
  txt1(Index).ForeColor = &H80000002
        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "txt1_GotFocus")
        '***Ende
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
       '***Beginn
10    On Error GoTo Fehler
       '***Ende
        
        
20      Select Case KeyCode
        Case vbKeyF1
30        If Shift = 1 Then
            'Die UMSCHALT-TASTE ist gedrückt.
            'Hilfetexte können erfast oder bearbeitet werden.
40          objHlp.HlpShow HlpWrite, objPRM.HlpID
50        Else
60          objHlp.HlpShow HlpRead, objPRM.HlpID
70        End If
80      Case vbKeyEscape
90        If TDBG1(0).DataChanged = False Then
100         Call objPRM.SprungNeu("Rückwärts", Shift, txt1(0).TabIndex)
110       End If
120     End Select
       '***Beginn
130           Exit Sub
Fehler:
140           Call FehlerErklärung("frmSP52810", "txt1_KeyDown")
       '***Ende

End Sub

Private Sub txt1_LostFocus(Index As Integer)
        '***Beginn
On Error GoTo Fehler
        '***Ende
  txt1(Index).BackColor = vbWindowBackground '&H80000005
  txt1(Index).ForeColor = vbWindowText  'Farbe des Textes in Fenstern

        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52810", "txt1_LostFocus")
        '***Ende
End Sub


Private Sub mnuAnsicht_ResFak_Click(Index As Integer)
      'DeW 17.03.2011, neu eingefuegt, wegen neuen Ansicht Menues,
      'Schrittweisen Vergroesserung des Formulars
10        Select Case Index
              Case 0  'auf Originale Groesse setzen
20                Call cResize.ResizeAboutPercent(0#, 0)
30            Case 2  'auf 120 Prozent
40                Call cResize.ResizeAboutPercent(20#, 2)
50            Case 4  'auf 140 Prozent
60                Call cResize.ResizeAboutPercent(40#, 4)
70            Case 6  'auf 160 Prozent
80                Call cResize.ResizeAboutPercent(60#, 6)
90            Case 8  'auf 180 Prozent
100               Call cResize.ResizeAboutPercent(80#, 8)
110           End Select
End Sub


Private Sub mnuAnsicht_ResetPosition_Click()
      'DeW 25.03.2011 neu eingebaut
      '
      '25.03.2011, neu Wunsch, nur das aktuelle Fenster
      'resetten... Daher in jedem Programmfenster einen Menueintrag
      'Unterroutine ResetWindowPos() in SP50000B.bas definiert
      '
10    Call ResetWindowPos(Me.hwnd, "SP51000")
      'DeW, neu, 16.05.2011, Loeschen aller Resize Eintraege,
      'falls Werte einmal falsch gespeichert werden
20    cResize.RemoveRegistryKeys

End Sub



Private Sub mnuAnsicht_Alle_Click()
          'DeW 17.03.2011, neu eingefuegt, neuer Eintrag im Ansicht Menue
          'fuer eine Aenderung der Groesse in allen Unterformularen
10        mnuAnsicht_Alle.Checked = Not mnuAnsicht_Alle.Checked
20        cResize.ResizeAllForms = mnuAnsicht_Alle.Checked
End Sub


Private Sub mnuAnsicht_Prop_Click()
          'DeW 17.03.2011, neu eingefuegt, neuer Eintrag im Ansicht Menue
          'fuer eine Proportionale Vergroesserung der Fenster
10        mnuAnsicht_Prop.Checked = Not mnuAnsicht_Prop.Checked
20        cResize.ScalingProportional = mnuAnsicht_Prop.Checked
          'Trigger Resize Event, to update View
30        cResize.Resize

End Sub


