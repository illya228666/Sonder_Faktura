VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmSP52832 
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   5955
   ClientTop       =   2835
   ClientWidth     =   9300
   Icon            =   "frmSP52832.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   9300
   Begin TrueOleDBGrid70.TDBGrid TDBG1 
      Height          =   5565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9816
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
      DeadAreaBackColor=   13160660
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
      _StyleDefs(8)   =   ":id=1,.fontname=Microsoft Sans Serif "
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Microsoft Sans Serif "
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Microsoft Sans Serif "
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
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   5610
      Width           =   9300
      _Version        =   65536
      _ExtentX        =   16404
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
         Caption         =   "Neu"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Löschen"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Speichern"
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   2700
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Kopieren"
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4020
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "cmd1"
         Height          =   300
         Index           =   4
         Left            =   5340
         TabIndex        =   4
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Drucken"
         Height          =   300
         Index           =   5
         Left            =   6660
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Schließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   7980
         TabIndex        =   2
         Top             =   30
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmSP52832"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Private objPRM As clsPRM
  Private objTDBG As clsTDBG7ole
  Private objHlp As SpHlp.clsHlp
  
  Private gRS As New ADODB.Recordset
Attribute gRS.VB_VarHelpID = -1
  Private Col As TrueOleDBGrid70.Column
  Private cols As TrueOleDBGrid70.Columns
  Private gIntSortColIndex As Integer
  Private gstrSortOrder As String
  Private gintPrivBelegArt As Integer
  
Public Function AusDBLesen() As Boolean
                     '***Beginn
10    On Error GoTo Fehler
                     '***Ende
                     
        Dim SQL As String
        Dim cn As New ADODB.Connection
        Dim Felder As String
        
20      cn.Mode = adModeUnknown
30      cn.Provider = "Microsoft.Jet.OLEDB.4.0"
        
        'Isolation und Kapselung der Transaktionen zum Herabsetzen des Transaktionslevels
40      cn.IsolationLevel = adXactCursorStability
        
50      cn.CursorLocation = adUseServer
        
60      Felder = objPRM.GetUseFields("TDBG1")
        
70      If gintPrivBelegArt = 0 Then
80        SQL = "SELECT " & Felder & " FROM 2800_Haupt INNER JOIN 2800_Archiv_Rng ON [2800_Haupt].BelegID = [2800_Archiv_Rng].BelegID WHERE [2800_Haupt].Art=0"
90      Else
100       SQL = "SELECT " & Felder & " FROM 2800_Haupt INNER JOIN 2800_Archiv_Gut ON [2800_Haupt].BelegID = [2800_Archiv_Gut].BelegID WHERE [2800_Haupt].Art=1"
110     End If
        
120     cn.Open GsHauptPfad + "dat\" & CStr(CInt(GsAnwenderNr)) & "\SP50000.dat"
130     Set gRS.ActiveConnection = cn
140     gRS.CursorLocation = adUseServer
150     gRS.CursorType = adOpenKeyset
160     gRS.Open SQL
        
170     Set TDBG1.DataSource = gRS
        
180     If gRS.RecordCount = 0 Then cmd1(0).Enabled = False
        
       
                     '***Beginn
190           Exit Function
Fehler:
200           Call FehlerErklärung("frmSP52832", "AusDBLesen")
                     '***Ende

End Function



Private Sub cmd1_Click(Index As Integer)
        
10      On Error GoTo Fehler
        
20      Select Case Index
        Case 0 'Anzeigen
30        If GbArchiv Then
40          MsgBox "Zugrif auf ein externes Archivierungssystem ist noch nicht realisiert.", vbExclamation
50        Else
60          If gRS.RecordCount > 0 Then
70            If Not gRS.EOF And Not gRS.BOF Then
80              If FileExists(gRS!Datei) Then
90                ShellExecute Me.hwnd, "Open", Datei(gRS!Datei), "", Pfad(gRS!Datei), 1
100             Else
110               Call MsgText(3, 91, 34, 211, 0)
120               MsgBox GsMsgText(0) & ": ''" & gRS!Datei & "'' " & GsMsgText(1) & vbCrLf & GsMsgText(2), vbExclamation
                  'MsgBox "Datei: ''" & gRS!Datei & "'' wurde nicht gefunden. Auswertung ist nicht möglich."
130             End If
140           End If
150         End If
160       End If
170     Case 6
180       Unload Me
190     End Select
        
200     Exit Sub
        
Fehler:
210     Call FehlerErklärung("frmSP52832", "cmd1_Click")
        
End Sub

Private Sub Form_Load()
        '***Beginn
On Error GoTo Fehler
        '***Ende
  
  gintPrivBelegArt = GintBelegArt
  
  If gintPrivBelegArt = 0 Then
    Me.top = frmRechnung.top
    Me.left = frmRechnung.left
  Else
    Me.top = frmGutschrift.top
    Me.left = frmGutschrift.left
  End If
  
  Set objPRM = New clsPRM
  Set objPRM.gForm = Me
  objPRM.PRM_Alle
  

  Set objHlp = New SpHlp.clsHlp
  objHlp.DatabaseName = GsHauptPfad & "hlp\SP50000.hlp"
  objHlp.Table = Me.Name
  objHlp.Caption = Me.Name & " - Feldhilfe"

  Set objTDBG = New clsTDBG7ole
  Set objTDBG.TDBG = TDBG1
  Set objTDBG.PrmDataBase = GDBprm
 
  AusDBLesen
  
  objTDBG.SichtbareZeilen = 23
  objTDBG.SatzAnzahl = gRS.RecordCount
  objTDBG.GridParameter ("SELECT * FROM PRM52832 WHERE name = 'TDBG1' ORDER BY index")
  'TDBG1.Width = 9050
  TDBG1.HoldFields

  TDBG1.MarqueeStyle = dbgHighlightCell
  TDBG1.HighlightRowStyle.BackColor = vbWindowBackground '&H80000005
  TDBG1.HighlightRowStyle.ForeColor = vbWindowText  'Farbe des Textes in Fenstern
  TDBG1.AllowColSelect = False
  TDBG1.ExtendRightColumn = True 'Die äußere rechte Spalte wird ans rechte Gridende erweitert.
  TDBG1.FilterBar = True
  Set cols = TDBG1.Columns
  


        '***Beginn
        Exit Sub
Fehler:
        Call FehlerErklärung("frmSP52832", "Form_Load")
        '***Ende
End Sub

Private Sub Form_Resize()
  If Me.Width > 500 Then TDBG1.Width = Me.Width - 150
  If Me.Height > 5000 Then TDBG1.Height = Me.Height - 825
  SSPanel1(0).top = TDBG1.top + TDBG1.Height + 60

End Sub

Private Sub TDBG1_DblClick()
  cmd1(0).Value = True
End Sub

Private Sub TDBG1_FilterChange()
              '***Beginn
10    On Error GoTo Fehler
              '***Ende
        Dim c As Integer
        

        'Sperren der Bildschirmausgabe während die Form verendert wird.
20      LockWindowUpdate (Me.hwnd)
30      c = TDBG1.Col
40      TDBG1.HoldFields
50      gRS.Filter = getFilter()
60      TDBG1.Col = c
70      TDBG1.EditActive = True
80      If gRS.RecordCount > 0 Then
90        cmd1(0).Enabled = True
100     Else
110       cmd1(0).Enabled = False
120     End If
        'Entsperren der Bildschirmausgabe während die Form verendert wird.
130     LockWindowUpdate (0&)

              '***Beginn
140           Exit Sub
Fehler:
              'Entsperren der Bildschirmausgabe während die Form verendert wird.
150           LockWindowUpdate (0&)
160           Call FehlerErklärung("frmSP52832", "TDBG1_FilterChange")
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
              
90            Select Case gRS.Fields(Col.DataField).Type
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
        
        
910     getFilter = tmp
        
              '***Beginn
920           Exit Function
Fehler:
930           Call FehlerErklärung("frmSP52832", "getFilter")
              '***Ende
End Function






