VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSP52820 
   Caption         =   "Artikel-Stamm"
   ClientHeight    =   5700
   ClientLeft      =   8505
   ClientTop       =   2535
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSP52820.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9270
   Tag             =   "1"
   Begin VB.Frame Frame1 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   60
      TabIndex        =   33
      Top             =   3450
      Width           =   9135
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   1020
         Index           =   11
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   8865
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Index           =   0
      Left            =   60
      TabIndex        =   23
      Top             =   60
      Width           =   9135
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "*Inaktiv"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   2940
         Width           =   1635
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
         Index           =   3
         Left            =   2640
         Picture         =   "frmSP52820.frx":0442
         Style           =   1  'Grafisch
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1050
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Durchlaufend"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   2700
         Width           =   1635
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
         Index           =   10
         Left            =   3030
         Picture         =   "frmSP52820.frx":0524
         Style           =   1  'Grafisch
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2250
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
         Index           =   7
         Left            =   2610
         Picture         =   "frmSP52820.frx":0606
         Style           =   1  'Grafisch
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2010
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
         Index           =   6
         Left            =   2610
         Picture         =   "frmSP52820.frx":06E8
         Style           =   1  'Grafisch
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1770
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
         Index           =   12
         Left            =   3030
         Picture         =   "frmSP52820.frx":07CA
         Style           =   1  'Grafisch
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1290
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
         Left            =   3000
         Picture         =   "frmSP52820.frx":08AC
         Style           =   1  'Grafisch
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   330
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   210
         Index           =   12
         Left            =   2640
         TabIndex        =   5
         Top             =   1290
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Steuer"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   2460
         Width           =   1635
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   10
         Left            =   1530
         TabIndex        =   11
         Top             =   2250
         Width           =   1485
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   1
         Left            =   1530
         TabIndex        =   1
         Top             =   570
         Width           =   7215
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   9
         Left            =   7050
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   8
         Left            =   7050
         TabIndex        =   9
         Top             =   660
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   7
         Left            =   1530
         TabIndex        =   8
         Top             =   2010
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   6
         Left            =   1530
         TabIndex        =   7
         Top             =   1770
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   5
         Left            =   1530
         TabIndex        =   6
         Top             =   1530
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   4
         Left            =   1530
         TabIndex        =   4
         Top             =   1290
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   3
         Left            =   1530
         TabIndex        =   3
         Top             =   1050
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   2
         Left            =   1530
         TabIndex        =   2
         Top             =   810
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   0
         Left            =   1530
         TabIndex        =   0
         Top             =   330
         Width           =   1485
      End
      Begin VB.Label lbl1 
         Caption         =   "Preis/Währung"
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label lbl1 
         Caption         =   "Währung"
         ForeColor       =   &H80000000&
         Height          =   210
         Index           =   12
         Left            =   600
         TabIndex        =   36
         Top             =   1140
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lbl1 
         Caption         =   "Textschlüssel"
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   35
         Top             =   2250
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "FiBu-Konto"
         Enabled         =   0   'False
         Height          =   210
         Index           =   9
         Left            =   5610
         TabIndex        =   34
         Top             =   870
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Bez"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   570
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Kostenst.-Konto"
         Enabled         =   0   'False
         Height          =   255
         Index           =   8
         Left            =   5610
         TabIndex        =   31
         Top             =   630
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Sachkonten-Schl"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   2010
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Kostenst.-Schl"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   1770
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Rabatt"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1530
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Einheit"
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   1050
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Menge"
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label lbl1 
         Caption         =   "Suchbegriff"
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   330
         Width           =   1425
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   22
      Top             =   5010
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
         Caption         =   "Schließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   7920
         TabIndex        =   16
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Drucken"
         Height          =   300
         Index           =   5
         Left            =   1380
         TabIndex        =   21
         Top             =   510
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "cmd1"
         Height          =   300
         Index           =   4
         Left            =   60
         TabIndex        =   20
         Top             =   510
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Kopieren"
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   2163
         TabIndex        =   17
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Speichern"
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   6093
         TabIndex        =   15
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Löschen"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4785
         TabIndex        =   19
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "New"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   0
         Left            =   3475
         TabIndex        =   18
         Top             =   30
         Width           =   1260
      End
   End
   Begin MSComctlLib.StatusBar sta1 
      Align           =   2  'Unten ausrichten
      Height          =   350
      Left            =   0
      TabIndex        =   37
      Top             =   5350
      Width           =   9270
      _ExtentX        =   16351
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
            Object.Width           =   13017
            MinWidth        =   12541
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDat 
      Caption         =   "Datei"
      Index           =   0
      Begin VB.Menu mnuDat1 
         Caption         =   "Schließen"
         Index           =   0
      End
   End
   Begin VB.Menu mnuBearb 
      Caption         =   "Bearbeiten"
      Index           =   0
      Begin VB.Menu mnuBearb1 
         Caption         =   "Neu"
         Index           =   0
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Löschen"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Speichern"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Kopieren"
         Enabled         =   0   'False
         Index           =   3
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
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSP52820"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
Dim gbDataChanged As Boolean

Dim gbSatzNeu     As Boolean
  
Dim gRS           As ADODB.Recordset

Dim gvntMerker

Dim gbEinfg               As Boolean
  
Private objPRM            As clsPRM

Private objSQLAusw        As SPSQLAuswahl.clsSQLAuswahl

Private objHlp            As SpHlp.clsHlp

Private lastControl       As Control                    'DH, 30.01.2015, 6.4.102, Wird bei txt1_LostFocus() gefuellt

Private denyChangeControl As Boolean              'DH, 30.01.2015, 6.4.103, Flag das festlegt, ob das Steuerlement verlassen werden darf
  
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

'<Added by: GW at: 31.01.2019, Ver.: 6.5.109 >
Private shiftPressed As Boolean
'</Added by: GW at: 31.01.2019, Ver.: 6.5.109 >

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   
        On Error GoTo Fehler
    
        '<Added by: GW at: 31.01.2019, Ver.: 6.5.109 >
100     If Shift <> 1 Then
105         shiftPressed = False
        End If
    
        Exit Sub

        '</Added by: GW at: 31.01.2019, Ver.: 6.5.109 >

Fehler:
110     Call FehlerErklärung("frmSP52820", "Form_KeyUp()")
   
End Sub

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

Private Sub Check1_Click(Index As Integer)

        On Error GoTo Fehler

100     If gRS.EOF Or gRS.BOF Then Exit Sub       'DH, 16.01.2018, 6.5.104, Wenn kein Datensatz 'aktiv' ist (z.B. bei Neuanlage), Methode verlassen

105     If gbDataChanged Then
110         If gRS.EditMode <> dbEditInProgress Then 'dbEditInProgress = 1
                'Edit-Modus ist noch nicht eingeschaltet.
                '40         gRS.Edit
115             cmd1(2).Enabled = True
120             mnuBearb1(2).Enabled = True
            End If
        End If

        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52820", "Check1_Click(" & Index & ")")
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        On Error GoTo Fehler

100     Select Case KeyCode

            Case vbKeyF1
105             objPRM.FindFirstString = "name = 'Check1' AND index = " & Index

110             If Shift = 1 Then
                    'Die UMSCHALT-TASTE ist gedrückt.
                    'Hilfetexte können erfast oder bearbeitet werden.
115                 objHlp.HlpShow HlpWrite, objPRM.HlpID
                Else
120                 objHlp.HlpShow HlpRead, objPRM.HlpID
                End If

125         Case vbKeyReturn, vbKeyDown
130             Call objPRM.SprungNeu("Vorwärts", Shift, Check1(Index).TabIndex, True)

135         Case vbKeyEscape, vbKeyUp
140             Call objPRM.SprungNeu("Rückwerts", Shift, Check1(Index).TabIndex, True)
        End Select

        Exit Sub

Fehler:
145     Call FehlerErklärung("frmSP52820", "Check1_KeyDown()")
End Sub

Private Sub cmd1_Click(Index As Integer)

        On Error GoTo Fehler

        Dim i              As Integer

        Dim antwort        As String

        Dim validatedIndex As Integer
        
        '<Added by: GW at: 26.08.2020, Ver.: 6.6.103 >
100     Select Case Index
        
            Case 2, 3

105             validatedIndex = Plausi

110             If validatedIndex <> -1 Then
115                 txt1(validatedIndex).SetFocus

                    Exit Sub

                End If
        
        End Select

        '</Added by: GW at: 26.08.2020, Ver.: 6.6.103 >
  
120     Select Case Index

            Case 0 'Neu TODO

                'DH, 06.10.2011, Abbrechen-Button beim Leeren hinzugefuegt
125             If cmd1(2).Enabled Then
130                 Call msgText(1, 18, 0, 0, 0)
                    '50         If MsgBox(GsMsgText(0), vbYesNo + vbExclamation) = vbYes Then
                    'Msgbox "Sie haben die Daten geändert. Wollen Sie die Änderungen speichern?"
135                 antwort = MsgBox(GsMsgText(0), vbYesNoCancel + vbExclamation + vbDefaultButton3, strMeldungCap)

                    '70          End If
140                 Select Case antwort

                        Case vbYes

                            '<Modified by: GW at 14.08.2020, Ver.: 6.6.103 >
                            '125                         Call SatzSpeichern
145                         If Not Speichern Then

                                Exit Sub

                            End If

                            '</Modified by: GW at 14.08.2020, Ver.: 6.6.103 >

150                     Case vbNo

155                     Case vbCancel

                            Exit Sub

                    End Select

                End If

160             Call MaskeLeeren
                '140       gbSatzNeu = True     'DH, 16.01.2018, 6.5.104, Leeren (ehemals Neu) darf nicht das Flag fuer einen neuen Datensatz setzen
165             txt1(0) = ""
170             txt1(0).SetFocus

175         Case 1 'Löschen
180             Call msgText(1, 14, 0, 0, 0)
185             antwort = MsgBox(GsMsgText(0), vbYesNo + vbQuestion + vbDefaultButton2, strMeldungCap)

                'antwort = MsgBox("Wollen Sie den Datensatz wirklich löschen?", vbYesNo + vbExclamation + vbDefaultButton2)
190             If antwort = vbYes Then
                    'DH, 24.10.2017, 6.5.101, Nach dem Loeschen keinen anderen Datensatz anzeigen, sondern nur die Maske leeren
                    '210         If gRS.RecordCount > 1 Then
                    '220           gRS.Delete
                    '230           gRS.MoveNext
                    '240           If gRS.EOF = True Then
                    '250             gRS.MoveLast
                    '260           End If
                    '270           Call SatzZeigen
                    '240         Else
                    '250           gRS.Delete
                    '260           txt1(0) = ""
                    '270           Call MaskeLeeren
                    '280         End If

195                 gRS.Delete

200                 If gRS.RecordCount > 0 Then gRS.MoveFirst

205                 Call MaskeLeeren
210                 txt1(0).text = ""
                End If

215             txt1(0).SetFocus

220         Case 2 'Speichern

                '<Modified by: GW at 13.08.2020, Ver.: 6.6.103 >
225             Speichern

                '</Modified by: GW at 13.08.2020, Ver.: 6.6.103 >

230         Case 3 'Kopieren

235             If cmd1(2).Enabled Then
240                 Call msgText(1, 18, 0, 0, 0)

245                 If MsgBox(GsMsgText(0), vbYesNo + vbExclamation, strMeldungCap) = vbYes Then

                        'Msgbox "Sie haben die Daten geändert. Wollen Sie die Änderungen speichern?"
                        '<Modified by: GW at 14.08.2020, Ver.: 6.6.103 >
                        '230                     Call SatzSpeichern
250                     If Not Speichern Then

                            Exit Sub

                        End If

                        '</Modified by: GW at 14.08.2020, Ver.: 6.6.103 >
                    End If
                End If

255             frmSP52821.ShowMe cReSize.CurrScaleFactorWidth, cReSize.CurrScaleFactorHeight
260             txt1(0).SetFocus
    
265         Case 6 'Schließen
270             Unload frmSP52821
275             Unload Me

        End Select

        Exit Sub

Fehler:
280     Call FehlerErklärung("frmSP52820", "cmd1_Click()")
End Sub

Private Sub cmdAuswahl_Click(Index As Integer)

        On Error GoTo Fehler

        Dim i         As Integer

        Dim ColLeft   As Long

        Dim RowBottom As Long

        Dim OT        As String

        Dim rc        As rect

100     Call GetWindowRect(txt1(Index).hwnd, rc)

105     ColLeft = rc.Left * Screen.TwipsPerPixelX '- 70
110     RowBottom = rc.bottom * Screen.TwipsPerPixelY '- 340
        
115     objSQLAusw.FilterBar = True
120     objSQLAusw.BorderStyle = 4
125     objSQLAusw.caption = lbl1(Index)
130     objSQLAusw.Top = RowBottom
135     objSQLAusw.Left = ColLeft

        '*****************************************************
        'HW, DeW, Mai 2011, Vergroessern des F2-Fensters
140     objSQLAusw.ScaleFactorHeight = cReSize.CurrScaleFactorHeight
145     objSQLAusw.ScaleFactorWidth = cReSize.CurrScaleFactorWidth
150     objSQLAusw.fontSize = Me.fontSize
        '*****************************************************
  
155     Select Case Index

            Case 0 'Schl.
    
160             If cmd1(2).Enabled = True Then
                    'Der aktuelle Datensatz wurde verändert
165                 Call msgText(1, 18, 0, 0, 0)
    
170                 If MsgBox(GsMsgText(0), vbYesNo + vbInformation, strMeldungCap) = vbYes Then
                        'MsgBox "Sie haben die Daten geändert. Wollen Sie die Änderungen speichern?"
175                     Call SatzSpeichern
                    Else
180                     Call SatzZeigen
185                     cmd1(2).Enabled = False
190                     mnuBearb1(2).Enabled = False
                    End If
                End If
    
                Dim oF2MCodeSchl As ResultF2_TextSchl
    
195             oF2MCodeSchl = GetF2_TextSchl("A", E_DATATYPE.Sonderfaktura_Rechnung, "Schl", txt1(0), 0, Me, cReSize, objPRM, , , , True)   'IL 12.03.2025 , Ver.: 6.7.107 : InaktiveAnzeigen = true
    
200             If oF2MCodeSchl.Canceled = False Then
205                 txt1(Index) = oF2MCodeSchl.Schl
210                 gRS.MoveFirst           'DH, 24.10.2017, 6.5.101, Vor dem Find() immer erst auf den ersten Datensatz stellen
215                 gRS.Find "Schl = '" & txt1(Index).text & "'"
220                 Call SatzZeigen
225                 Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                Else
230                 txt1(Index).SetFocus
                End If
                
                '<Added by: IL at: 8.27.2024-15:47:23 on machine: T017>

235         Case 3 'Verpackung
                
                Dim oF2Verpackung As ResultF2_Verpackung
                        
240             oF2Verpackung = GetF2_Verpackung("Schl", txt1(Index), Index, Me, cReSize, objPRM)
                        
245             If oF2Verpackung.Canceled = False Then

250                 txt1(Index).text = Trim(oF2Verpackung.Schl)

255                 modERechnung.boolVPSchlusselanderung = True

260                 If Not CheckENCodeBeiVerpackung(0, oF2Verpackung.Schl) Then

265                     txt1(Index).text = ""
                        
270                     If txt1(Index).Enabled Then txt1(Index).SetFocus

                    Else

275                     Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)

                    End If

                End If
                
                '</Added by: IL at: 8.27.2024-15:47:23 on machine: T017>

280         Case 6, 7  'KostenSchl, SachkontenSchl
285             objSQLAusw.ColParameter 0, colWidth, txt1(Index).width

                'Datensatz positionieren.
290             If Trim(txt1(Index)) <> "" Then
295                 objSQLAusw.Find = "Schl like '" & txt1(Index) & "*'"
                End If

300             If Index = 6 Then
305                 objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl,Bezeichnung,Erlöse,Kosten FROM [1100_FiBuKostenStellen]"
                Else
310                 objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl,Bezeichnung,Erlöse,Kosten FROM [1100_FiBuSachkonten]"
                End If

315             If objSQLAusw.Abbruch = False Then
320                 txt1(Index) = objSQLAusw.FieldText(0)
325                 Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                Else
330                 txt1(Index).SetFocus
                End If

335         Case 10 'Textschl.

                '<Modified by: IL at 9.17.2024-11:04:22 on machine: T017>
                '# Umgestellt aud modF2
                Dim oF2TextSchl As ResultF2_TextSchl

340             oF2TextSchl = GetF2_TextSchl("T", E_DATATYPE.Sonderfaktura_Rechnung, "Schl", txt1(10), 10, Me, cReSize, objPRM, objSQLAusw.GetIfOnesHit)

345             If oF2TextSchl.Canceled = False Then
350                 txt1(Index) = oF2TextSchl.Schl
355                 txt1(Index + 1) = oF2TextSchl.Inhalt
360                 Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                Else
365                 txt1(Index).SetFocus
                End If
                
                '208 Orig:                 objSQLAusw.ColParameter 2, ColVisible, 0
                '
                '                                'Datensatz positionieren.
                '210                             If Trim(txt1(Index)) <> "" Then
                '212                                 objSQLAusw.Find = "Schl like '" & txt1(Index).text & "%'"
                '                                End If
                '
                '214                             objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl, Bez, Inhalt FROM [2800_Texte] ORDER BY Schl"
                '
                '216                             If objSQLAusw.Abbruch = False Then
                '218                                 txt1(Index) = objSQLAusw.FieldText(0)
                '220                                 txt1(Index + 1) = objSQLAusw.FieldText(2)
                '222                                 Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                '                                Else
                '224                                 txt1(Index).SetFocus
                '                                End If

                '</Modified by: IL at 9.17.2024-11:04:22 on machine: T017>

370         Case 12 'Währung
375             objSQLAusw.ColParameter 0, colWidth, txt1(Index).width + 100

                'Datensatz positionieren.
380             If Trim(txt1(Index)) <> "" Then
385                 objSQLAusw.Find = "ISO like '" & txt1(Index).text & "*'"
                End If

390             objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT DISTINCT ISO FROM [1100_Währungen]"

395             If objSQLAusw.Abbruch = False Then
400                 txt1(Index) = objSQLAusw.FieldText(0)
405                 Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                Else
410                 txt1(Index).SetFocus
                End If

        End Select

        Exit Sub

Fehler:
415     Call FehlerErklärung("frmSP52820", "cmdAuswahl_Click()")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        Dim AltDown

        Const vbAltMask = 4
                
        '<Added by: GW at: 31.01.2019, Ver.: 6.5.109 >
100     If Shift = 1 Then
105         shiftPressed = True
        End If

        '</Added by: GW at: 31.01.2019, Ver.: 6.5.109 >
  
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
140     Call FehlerErklärung("frmSP52820", "Form_KeyDown")
        '***Ende
End Sub

Private Sub Form_Load()

        On Error GoTo Fehler

        Dim i As Integer

        Dim x As Control

        '<Added by: GW at: 31.01.2019, Ver.: 6.5.109 >
100     shiftPressed = False
        '</Added by: GW at: 31.01.2019, Ver.: 6.5.109 >
        
105     SaveSetting "SP50000", App.EXEName, "SP62820_WndHwnd", Me.hwnd 'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

        '########### SkinFramework ##############################
        'HW 04.03.2011 - Is ein Windows Skin Tool von Codejock Software
        'SkinFramework1.LoadSkin GsHauptPfad & "\exe\Spedifix.cjstyles", "NormalSilver.ini"
        'SkinFramework1.ApplyWindow Me.hwnd
        '########################################################

        '########## Subclassing: Messages festlegen #############
        ' DeW, ZyG Mai 2011
110     AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO 'DeW
115     AttachMessage Me, Me.hwnd, WM_EXITSIZEMOVE 'DeW
        '
        '########################################################

        '####### Subclassing: Groessenbegrenzung Formular #######
        ' TODO in Arbeit, MagicNumbers
        'dieser Aufruf kann je nach Programm-Modul woanders in der
        'Form_Load Methode stehen!
        'Zuerst muss der "alte" Code die Zuweisung von Breit und
        'Hoehe korrekt vorgenommen haben!
120     SetMinMaxInfo Me.hwnd, Me.height, (Me.height * 2), Me.width, (Me.width * 2)
        '
        '########################################################
  
        '60      If GsTitel <> "" Then
        '70        GlSP51000hwnd = FindWindow(vbNullString, GsTitel)
        '80        SetWindowLong Me.hwnd, GWL_HWNDPARENT, GlSP51000hwnd
        '90      End If
  
125     Set objPRM = New clsPRM
130     Set objPRM.gForm = Me
135     objPRM.PRM_Alle
        
140     strMeldungCap = mnuDummy.caption
        
        '130     Me.left = GetSetting("SP50000", "SP52800", "SP52820Left", "2550")
        '140     Me.Top = GetSetting("SP50000", "SP52800", "SP52820Top", "870")
  
        'Me.Width = FORM_WIDTH 'Const in SP5000B.bas
145     Me.width = 9400 'DH, 16.09.2011, im Zuge der Button-Neuausrichtung verbreitert
150     Me.height = FORM_HEIGHT
  
155     SetXPSize Me

160     Call setSkinnerBackColor(sta1)
  
        '140     If left(Me.Name, 3) = "frm" Then
        '150       sta1.Panels(1).Text = Mid(Me.Name, 4)
        '160     Else
        '170       sta1.Panels(1).Text = Me.Name
        '180     End If
165     sta1.Panels(1).text = "SP62820"
  
        'txt1(11).Width = TEXT_BREITE
170     txt1(11).width = 8900 'DH, 16.09.2011' TEXT_BREITE definiert die Breite von txt1 aus SP52810
        ' was dann für SP52820 zu klein ist -> verbreitert
175     sta1.Panels(2).text = DisplayVerInfo(GsHauptPfadLokal & "exe\" & Gc_strExeFile)
  
        'Me.Show
180     OPEN_gConn
185     Set gRS = New ADODB.Recordset
190     gRS.Open "SELECT * FROM [2800_Artikel] ORDER BY [Schl]", gConn, adOpenKeyset, adLockOptimistic

        'HW 06.05.2015
195     Set objSQLAusw = New SPSQLAuswahl.clsSQLAuswahl

200     Set objHlp = New SpHlp.clsHlp
205     objHlp.DatabaseName = GsHauptPfadLokal & "hlp\SP50000.hlp"
210     objHlp.table = Me.name
215     objHlp.caption = Me.name & " - Feldhilfe"
        '300     objHlp.ParentFrm = Me
  
220     SaveSetting "SP50000", "SP52800", "SP52820", Me.caption
 
        '###### Formular Resizing: Parameter setzen#############
        ' DeW, ZyG, Mai 2011
        'Section- oder KeyBezeichnung sind in vielen Faellen in
        'altem Code hart eincodiert worden, manchmal wird auch
        'eine Variable verwendet...
225     Set cReSize = New FormResize
230     cReSize.setSectionBezeichnung = "SP52820"
235     cReSize.setKeyBezeichnung = "SP52820"
240     cReSize.setIstUnterFenster = False
        '
        '########################################################

        '######## Formular Resizing: Formular zuweisen ##########
        ' DeW, Mai 2011
        'Zuweisung von Form erst nach Groessensetzung s.o. Me.Width = ...
        'aber auf jeden Fall nach SetMinMaxInfo ... fuer die
        'Groessenbegrenzung
245     cReSize.Form = Me
        '

        'Speichere keine Informationen (Spaltenbreiten usw.) fuer die
        'Tabellen im Form, wenn z.B. nur eine einzelne Tabelle
        'vorhanden ist, die jeweils mit neuen Daten gefuellt
        'und an eine andere Position verschoben wird (z.B. SP51000
        'Mandantenstamm
250     cReSize.IgnoreTrueDBGridInfo = True
        '
        '########################################################

255     cReSize.resize

        '130     Me.left = GetSetting("SP50000", "SP52800", "SP52820Left", "2550")
        '140     Me.Top = GetSetting("SP50000", "SP52800", "SP52820Top", "870")

        'HW 10.07.2014
260     Call readWindowPos(Me, "SP52800", "SP52820Left", "SP52820Top")

265     If GsTitel <> "" Then
270         GlSP51000hwnd = FindWindow(vbNullString, GsTitel)
275         SetWindowLong Me.hwnd, GWL_HWNDPARENT, GlSP51000hwnd
        End If
   
        Exit Sub

Fehler:
280     Call FehlerErklärung("frmSP52820", "Form_Load()")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        On Error GoTo Fehler

        Dim dialogResult As Long

        'DH, 06.10.2011, Abbrechen beim Beenden hinzugefuegt
100     If cmd1(2).Enabled = True Then
            'Der aktuelle Datensatz wurde verändert
105         Call msgText(1, 18, 0, 0, 0)
            'If MsgBox(GsMsgText(0), vbYesNo + vbInformation) = vbYes Then
            'MsgBox "Sie haben die daten geändert. Wollen Sie die Änderungen speichern?"
110         dialogResult = MsgBox(GsMsgText(0), vbYesNoCancel + vbInformation + vbDefaultButton3, strMeldungCap)

115         Select Case dialogResult

                Case vbYes

                    '<Modified by: GW at 14.08.2020, Ver.: 6.6.103 >
                    '120                 If Not SatzSpeichern Then
120                 If Not Speichern Then
                        '</Modified by: GW at 14.08.2020, Ver.: 6.6.103 >
125                     Cancel = True

                        Exit Sub

                    End If

130             Case vbNo

135             Case vbCancel
140                 Cancel = True

                    Exit Sub              'DH, 30.01.2015, 6.4.101, Damit das Fenster bei Abbrechen nicht dennoch verschwindet

            End Select

            'End If
        End If
    
        '####### Subclassing: Messages austragen #############
        'DeW, Mai 2011
145     DetachMessage Me, Me.hwnd, WM_EXITSIZEMOVE
150     DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
        '
        '#####################################################

        '####### Subclassing: Groessenbegrenzung loeschen #######
        'DeW, Mai 2011
155     RemoveMinMaxInfo Me.hwnd
        '
        '########################################################

        'HW 10.07.2014
160     Call writeWindowPos(Me, "SP52800", "SP52820Left", "SP52820Top")

165     Me.Visible = False
170     Me.Hide
175     DoEvents
180     Sleep (0.5)

        Exit Sub

Fehler:
185     Call FehlerErklärung("frmSP52820", "Form_QueryUnload()")
End Sub

Private Sub Form_Resize()

    '***Beginn
    On Error GoTo Fehler

    '***Ende
    sta1.Panels(1).width = Me.width / FORM_WIDTH * FORM_PANELS_1
    sta1.Panels(2).width = Me.width / FORM_WIDTH * FORM_PANELS_2

    '***Beginn
    Exit Sub

Fehler:
    Call FehlerErklärung("frmSP52820", "Form_Resize")
    '***Ende
End Sub

Private Sub Form_Unload(Cancel As Integer)

        On Error GoTo Fehler

100     SaveSetting "SP50000", App.EXEName, "SP62820_WndHwnd", ""               'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

        'Unterrutine in SP50000B.bas
105     Call ProgrammAus("282")

110     Protokoll iAppend, " *****PROGRAMM ENDE*****: 282 -> " & Now & vbCrLf & "-----" ' vbCrLf &

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
120     Set objPRM = Nothing
125     Set objSQLAusw = Nothing
130     Set objHlp = Nothing

135     If gRS.state = adStateOpen Then gRS.Close
140     If Err.number <> 0 Then Err.Clear
145     Set gRS = Nothing

150     CLOSE_gConn True

155     DisposeObjects Me                                                       'HW 26.07.2013
 
        Exit Sub

Fehler:
160     Call FehlerErklärung("frmSP52820", "Form_Unload()")
End Sub

Private Sub mnuBearb1_Click(Index As Integer)

        On Error GoTo Fehler
  
100     Call cmd1_Click(Index)
  
        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52820", "mnuBearb1_Click()")
End Sub

Private Sub mnuDat1_Click(Index As Integer)

        On Error GoTo Fehler

100     Call cmd1_Click(6)

        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52820", "mnuDat1_Click()")
End Sub

Private Sub mnuInfo_Click(Index As Integer)
    
        On Error GoTo Fehler

100     If shiftPressed Then
105         objHlp.HlpShow HlpWrite, "InfoSonderfaktura"
        Else
110         objHlp.HlpShow HlpRead, "InfoSonderfaktura"
        End If

115     shiftPressed = False
120     objPRM.FindFirstString = "name = '" & Me.ActiveControl.name & "' AND index = " & Me.ActiveControl.Index
    
        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52820", "mnuInfo_Click()")
  
End Sub

Private Sub mnuUpdateInfo_Click(Index As Integer)
    
        On Error GoTo Fehler
   
100     objPRM.FindFirstString = "name = 'mnuUpdateInfo' "
105     objHlp.UpdateAnzeigen = True
110     objHlp.UpdateCounter = gi_UpdateAenderung

115     If shiftPressed = True Then

120         objHlp.HlpShow HlpWrite, "UpdateAenderung" & programmNr
125         shiftPressed = False
        Else
130         objHlp.HlpShow HlpRead, "UpdateAenderung" & programmNr
        End If

135     objHlp.UpdateAnzeigen = False
   
        Exit Sub

Fehler:
140     Call FehlerErklärung("frmSP52820", "mnuUpdateInfo_Click()")
   
End Sub

Private Sub txt1_Change(Index As Integer)

        On Error GoTo Fehler
        
100     cmd1(3).Enabled = False
105     mnuBearb1(3).Enabled = True
        
110     If gRS.EOF Or gRS.BOF Then Exit Sub             'DH, 12.01.2018, 6.5.104, Wenn kein Datensatz geoeffnet ist, Methode verlassen
  
115     If Index = 0 Then
120         If gbDataChanged Then
125             gvntMerker = txt1(Index)

130             If cmd1(2).Enabled Then
135                 Call msgText(1, 18, 0, 0, 0)

140                 If MsgBox(GsMsgText(0), vbYesNo + vbExclamation, strMeldungCap) = vbYes Then
                        'Msgbox "Sie haben die Daten geändert. Wollen Sie die Änderungen speichern?"
145                     Call SatzSpeichern
                    End If
                End If

150             Call MaskeLeeren
155             txt1(Index) = gvntMerker
            End If

        Else

160         If gbDataChanged Then
165             If gRS.EditMode <> dbEditInProgress Then 'dbEditInProgress = 1
                    'Edit-Modus ist noch nicht eingeschaltet.
                    '170           gRS.Edit
170                 cmd1(2).Enabled = True
175                 mnuBearb1(2).Enabled = True

180                 If Index = 3 Then modERechnung.boolVPSchlusselanderung = True
                End If
            End If
        End If
  
        Exit Sub

Fehler:

185     Select Case Err.number

            Case 3260 'Aktualisieren nicht möglich. Der Datensatz ist von anderem Benutzer gesperrt.
190             Call FehlerErklärung("frmSP52820", "txt1_Change")
195             Call SatzZeigen

200         Case Else
205             Call FehlerErklärung("frmSP52820", "txt1_Change")
        End Select

End Sub

Private Sub txt1_GotFocus(Index As Integer)

        On Error GoTo Fehler

100     txt1(Index).SelStart = 0
105     txt1(Index).selLength = Len(txt1(Index))
110     txt1(Index).ForeColor = vbWindowText
115     txt1(Index).BackColor = &HC0E0FF  'hellorange

120     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
125     gvntMerker = txt1(Index)

        'DH, 30.01.2015, 6.4.102
        'Hiermit soll verhindert werden, dass mit der Maus innerhalb der Textfelder gesprungen
        'werden kann.
130     If denyChangeControl Then
135         If lastControl.name = "txt1" Then
140             txt1(lastControl.Index).SetFocus
            End If
        End If

145     denyChangeControl = False

        '
        Exit Sub

Fehler:
150     Call FehlerErklärung("frmSP52820", "txt1_GotFocus()")
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        On Error GoTo Fehler

        Dim Cancel As Boolean

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

                    Case 0, 3, 6, 7, 10, 12 'Schl, KostenSchl, SachkontenSchl Text, Währung
130                     Call cmdAuswahl_Click(Index)
                End Select

135         Case vbKeyReturn, vbKeyDown

                '<Added by: GW at: 26.08.2020, Ver.: 6.6.103 >
140             Select Case Index
                
                    Case 6, 7
145                     Call txt1_Validate(Index, Cancel)

150                     If Cancel Then

                            Exit Sub

                        End If

                End Select

                '</Added by: GW at: 26.08.2020, Ver.: 6.6.103 >

155             If SatzAufbereiten(Index) Then
160                 objPRM.FindFirstString = "name = 'txt1' AND index = " & Index

                    'Weil objPRM.SprungNeu Validate-Ereignis nicht auslöst,
                    'muss die Umwandlung und Prüfung an der Stelle stattfinden.
165                 Select Case Index

                        Case 4              'Preis
170                         txt1(Index).text = Format(txt1(Index).text, ZahlFormat(postCommaPreis))

175                     Case Else           'Alle sonstigen Felder
180                         txt1(Index).text = objPRM.EingabeUmwandlung(txt1(Index))

                    End Select

185                 If objPRM.EingabeFehler(txt1(Index)) = False Then

                        '<Added by: IL at: 8.28.2024-09:35:21 on machine: T017>
190                     If Index = 3 Then

195                         If PlausiF2(Index) = -1 Then

200                             If Not CheckENCodeBeiVerpackung(0, txt1(Index).text) Then

205                                 txt1(Index).text = ""
210                                 txt1(Index).SetFocus

                                    Exit Sub

                                End If

                            Else
            
215                             txt1(Index).text = ""
220                             txt1(Index).SetFocus

                                Exit Sub
                                
                            End If
                            
                        End If

                        '</Added by: IL at: 8.28.2024-09:35:21 on machine: T017>

225                     Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                    End If
                End If

230         Case vbKeyEscape, vbKeyUp
235             objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
240             txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

245             If objPRM.EingabeFehler(txt1(Index)) = False Then
250                 Call objPRM.SprungNeu("Rückwerts", Shift, txt1(Index).TabIndex, True)
                End If
    
255         Case vbKeyPageDown 'Blättern in der Datensatzgruppe. Vorwärtsbewegung (Bild Ab-Taste).

260             If gRS.RecordCount > 0 Then
265                 If cmd1(2).Enabled = True Then
                        'Der aktuelle Datensatz wurde verändert
270                     Call msgText(1, 18, 0, 0, 0)

275                     If MsgBox(GsMsgText(0), vbYesNo + vbInformation, strMeldungCap) = vbYes Then
                            'MsgBox "Sie haben die daten geändert. Wollen Sie die Änderungen speichern?"
280                         Call SatzSpeichern
                        Else
285                         cmd1(2).Enabled = False
290                         mnuBearb1(2).Enabled = False
                        End If
                    End If

295                 If cmd1(1).Enabled = False Then
300                     gRS.MoveFirst
                    Else
305                     gRS.MoveNext

310                     If gRS.EOF Then
315                         gRS.MoveLast
320                         Beep
                        End If
                    End If

325                 Call SatzZeigen
                End If

330         Case vbKeyPageUp 'Blättern in der Datensatzgruppe. Rückwärtsbewegung (Bild Auf-Taste).

335             If gRS.RecordCount > 0 Then
340                 If cmd1(2).Enabled = True Then
                        'Der aktuelle Datensatz wurde verändert
345                     Call msgText(1, 18, 0, 0, 0)

350                     If MsgBox(GsMsgText(0), vbYesNo + vbInformation, strMeldungCap) = vbYes Then
                            'MsgBox "Sie haben die daten geändert. Wollen Sie die Änderungen speichern?"
355                         Call SatzSpeichern
                        Else
360                         cmd1(2).Enabled = False
365                         mnuBearb1(2).Enabled = False
                        End If
                    End If

370                 If cmd1(1).Enabled = False Then
375                     gRS.MoveLast
                    Else
380                     gRS.MovePrevious

385                     If gRS.BOF Then
390                         gRS.MoveFirst
395                         Beep
                        End If
                    End If

400                 Call SatzZeigen
                End If

        End Select

        Exit Sub

Fehler:
405     Call FehlerErklärung("frmSP52820", "txt1_KeyDown()")
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)

        On Error GoTo Fehler

100     Select Case KeyAscii

            Case vbKeyReturn, vbKeyEscape
105             KeyAscii = 0
        End Select

        'DH, 24.10.2017, 6.5.101, Feld Menge gesondert behandeln und hier nur EIN Minuszeichen zulassen.
        '                         Das ist erstmal sicherer, als das in den PRM-Funktionen zu aendern.
        '                         ...das Ganze muesste man eigentlich auch noch fuer das Komma machen
        '                         ...oder einfach die PRM-Funktionen mal ueberarbeiten.
110     If Index = 2 And KeyAscii = Asc("-") Then
115         If InStr(1, txt1(Index).text, "-") > 0 Then                     'Wenn bereits ein Minus im Text vorkommt
120             If txt1(Index).selLength <> Len(txt1(Index).text) Then      'Wenn nicht der gesamte Text markiert ist
125                 KeyAscii = 0

                    Exit Sub

                End If
            End If
        End If

130     KeyAscii = objPRM.EingabePrüfung(KeyAscii, txt1(Index).text)

135     If txt1(Index).MaxLength = 0 Then

            'Prüfen, ob die neue Zeichenfolge, die zugelassene Länge übersteigt. (Numerischen Felder. Sonst gilt die MaxLength-Eigenschaft von clsPRM)
            Dim NeuText As String

140         If txt1(Index).selLength = 0 Then NeuText = Mid(txt1(Index), 1, txt1(Index).SelStart) & Chr(KeyAscii) & Mid(txt1(Index), txt1(Index).SelStart + 1)
145         If KeyAscii <> 0 And KeyAscii <> 8 Then
150             If Len(NeuText) > objPRM.EingabeLaenge(NeuText) Then
155                 KeyAscii = 0
                Else
160                 KeyAscii = objPRM.CheckPrePostComma(txt1(Index), KeyAscii, 12, postCommaPreis)
                End If
            End If
        End If

        Exit Sub

Fehler:
165     Call FehlerErklärung("frmSP52820", "txt1_KeyPress()")
End Sub

Private Sub txt1_LostFocus(Index As Integer)

        On Error GoTo Fehler

100     txt1(Index).ForeColor = vbActiveTitleBar
105     txt1(Index).BackColor = vbWindowBackground  'Fensterhintergrund(weiß)
  
        '40    Debug.Print "Loosing Focus"
110     Set lastControl = txt1(Index)
    
        Exit Sub

Fehler:
115     Call FehlerErklärung("frmSP52820", "txt1_LostFocus()")
End Sub

Public Sub SatzZeigen()

        On Error GoTo Fehler
 
100     gbDataChanged = False
105     gbSatzNeu = False

110     txt1(0) = "" & gRS!Schl
115     txt1(1) = "" & gRS!bez
120     objPRM.FindFirstString = "name = 'txt1' AND index = 2"
125     txt1(2) = objPRM.EingabeUmwandlung(gRS!Menge)
130     txt1(3) = "" & gRS!Einheit
135     objPRM.FindFirstString = "name = 'txt1' AND index = 4"
        '90      txt1(4) = objPRM.EingabeUmwandlung(gRS!Preis)
140     txt1(4).text = Format(gRS!Preis, ZahlFormat(postCommaPreis))
145     objPRM.FindFirstString = "name = 'txt1' AND index = 5"
150     txt1(5) = objPRM.EingabeUmwandlung(gRS!Rabatt)
155     txt1(6) = "" & gRS!KostSchl
160     txt1(7) = "" & gRS!FiBuSchl
165     txt1(8) = "" & gRS!KostKonto
170     txt1(9) = "" & gRS!FibuKonto
175     txt1(10) = "" & gRS!TextSchl

180     If txt1(10) <> "" Then
185         objSQLAusw.GetIfOnesHit = True
190         Call cmdAuswahl_Click(10)
195         objSQLAusw.GetIfOnesHit = False
        Else
200         txt1(11) = ""
        End If

205     txt1(12) = "" & gRS!Wrg
210     Check1(0).value = gRS!Steuer
215     Check1(1).value = gRS!Durchlaufend
220     Check1(2).value = IIf(gRS!status = 2, 1, 0)                                             'IL 12.03.2025 , Ver.: 6.7.107
  
225     gbDataChanged = True
230     cmd1(1).Enabled = True
235     mnuBearb1(1).Enabled = True
240     cmd1(3).Enabled = True
245     mnuBearb1(3).Enabled = True

        Exit Sub

Fehler:
250     Call FehlerErklärung("frmSP52820", "SatzZeigen()")
End Sub

Public Sub MaskeLeeren()

        On Error GoTo Fehler

        Dim i As Integer
  
100     gbDataChanged = False

105     For i = 1 To txt1.Count - 1
110         txt1(i) = ""
        Next

115     Check1(0).value = 0
120     Check1(1).value = 0
125     Check1(2).value = 0                                                    'IL 12.03.2025 , Ver.: 6.7.107
  
130     gbDataChanged = True
135     cmd1(1).Enabled = False
140     mnuBearb1(1).Enabled = False
145     cmd1(3).Enabled = False
150     mnuBearb1(3).Enabled = False
  
155     cmd1(2).Enabled = False
160     mnuBearb1(2).Enabled = False

        Exit Sub

Fehler:
165     Call FehlerErklärung("frmSP52820", "MaskeLeeren()")
End Sub

Public Function SatzSpeichern(Optional neu As Boolean, _
                              Optional customRS As ADODB.Recordset = Nothing) As Boolean     'IL 27.02.2025 , Ver.: 6.7.106 : Optional customRS As ADODB.Recordset = gRS

        On Error GoTo Fehler

        'DH, 16.01.2018, 6.5.104, Hier einfach den globalen Parameter verwenden, da diese an
        '                         unterschiedlichen Stellen gesetzt wird.
100     neu = gbSatzNeu

        '<Modified by: IL at 27.02.2025, Ver.: 6.7.106 >
        '# Erlaube bei Bedarf die Übertragung der Funktion auf andere Datensatznetzwerke zum Kopieren von Artikeln auf andere Clients. gRS wird weiterhin standardmäßig eingesetzt
        
105     If customRS Is Nothing Then Set customRS = gRS
        
110     If neu Then

115         customRS.AddNew
120         customRS!Schl = txt1(0)
125         customRS!ErstVon = GetSetting("SP50000", "Settings", "User", "")

        Else
        
            '70        gRS.Edit
            
        End If
  
130     customRS!bez = txt1(1)
135     customRS!Menge = txt1(2)
140     customRS!Einheit = txt1(3)
145     customRS!Preis = txt1(4)
150     customRS!Rabatt = txt1(5)
155     customRS!KostSchl = txt1(6)
160     customRS!FiBuSchl = txt1(7)
165     customRS!KostKonto = txt1(8)
170     customRS!FibuKonto = txt1(9)
175     customRS!TextSchl = txt1(10)
180     customRS!Wrg = txt1(12)
185     customRS!Steuer = Check1(0).value
190     customRS!Durchlaufend = Check1(1).value
195     customRS!status = IIf(Check1(2).value = 1, 2, 0)                             'IL 12.03.2025 , Ver.: 6.7.107 : DLL macht Artikel nur rot, wenn Sperre = 2
  
200     customRS!AendVon = GetSetting("SP50000", "Settings", "User", "")
205     customRS!AendDat = Now
210     customRS.Update
        '</Modified by: IL at 27.02.2025, Ver.: 6.7.106 >
  
        ''250     gRS.bookmark = gRS.LastModified
        '240     gRS.MoveLast
        '250     Call SatzZeigen

        '270     txt1(0).SetFocus   'DH, 07.02.2014, 6.2.102, Das Textfeld kann nicht angesprungen werden, da noch ein Modales Fenster geoeffnet ist.

215     cmd1(2).Enabled = False
220     mnuBearb1(2).Enabled = False
225     SatzSpeichern = True
230     gbSatzNeu = False

235     Call MaskeLeeren
        '280     gbSatzNeu = True
240     txt1(0) = ""

245     If txt1(0).Enabled Then txt1(0).SetFocus

        Exit Function

Fehler:

250     Select Case Err.number

            Case 3260 'Aktualisieren nicht möglich. Der Datensatz ist von anderem Benutzer gesperrt.
255             Call FehlerErklärung("frmSP52820", "SatzSpeichern()")
260             cmd1(2).Enabled = True
265             mnuBearb1(2).Enabled = True
270             SatzSpeichern = False

275         Case Else
280             Call FehlerErklärung("frmSP52820", "SatzSpeichern()")
        End Select

End Function

Public Function SatzAufbereiten(Index As Integer) As Boolean

        On Error GoTo Fehler

        Dim i       As Integer

        Dim antwort As String

        Dim query   As String

        Dim SatzOK  As Boolean

        Dim rs      As ADODB.Recordset
  
100     Set rs = New ADODB.Recordset
     
105     Select Case Index

            Case 0  'Schl

                'DH, 12.01.2018, 6.5.104, Logik bei Neuanlage angepasst.
                '                         Vorhandene MCodes werden geoeffnet,
                '                         nicht vorhandene werden (auf Nachfrage) angelegt.
110             If Trim(txt1(Index).text) <> "" Then
115                 If gRS.RecordCount > 0 Then
120                     gRS.MoveFirst
125                     gRS.Find "[Schl] = '" & Trim(txt1(Index).text) & "'"
  
130                     If Not gRS.EOF And Not gRS.BOF Then                     'Wenn der eingegebene Datensatz gefunden wurde
135                         Call SatzZeigen
140                         gvntMerker = txt1(0)
145                         SatzAufbereiten = True

                            Exit Function

                        Else
150                         GoTo Neuanlage
                        End If

                    Else
155                     GoTo Neuanlage
                    End If

                Else
160                 SatzAufbereiten = False
                End If
  
Neuanlage:

165             Call msgText(2, 39, 29, 0, 0)                                   'Artikel exisitiert nicht. Moechten Sie ihn hinzufuegen ?

170             If MsgBox(GsMsgText(0) & " '' " & txt1(Index) & " '' " & GsMsgText(1), vbYesNo + vbInformation, strMeldungCap) = vbYes Then
                    '                'DH, 24.10.2017, 6.5.101, Bei Neuanlage soll die Waehrung aus dem Mandantenstamm gezogen werden
175                 query = ""
180                 query = query & "SELECT Wrg.Währung "
185                 query = query & "FROM [1100_Mandant] AS Mandant "
190                 query = query & "LEFT JOIN [1100_Währungen] AS Wrg ON(Wrg.Schl = Mandant.EigenWrg)"

195                 rs.Open query, gConn, adOpenStatic, adLockReadOnly
                    '
                    '290             gRS.AddNew
                    '300             gRS.Fields(0).Value = Trim(txt1(0))
                    '310             If rs.RecordCount > 0 Then gRS.Fields("Wrg").Value = rs.Fields("Währung").Value
                    '320             gRS.Update
                    '330             gRS.MoveLast
                    '340             gbSatzNeu = False
                    '350             gvntMerker = txt1(0)
                    '360             Call SatzZeigen

                    'DH, 16.01.2018, 6.5.104, Die Logik der Neuanlage soll jetzt so wie im Kundenstamm erfolgen.
                    '                         Das tatsaechliche Anlegen des Kunden also erst beim Beenden/Speichern.
200                 gbSatzNeu = True
205                 cmd1(2).Enabled = True
210                 objPRM.FindFirstString = "name = 'txt1' AND index = 2"
215                 txt1(2) = objPRM.EingabeUmwandlung(1)
220                 txt1(4).text = Format(0, ZahlFormat(postCommaPreis))                    'Preis
225                 objPRM.FindFirstString = "name = 'txt1' AND index = 5"
230                 txt1(5).text = objPRM.EingabeUmwandlung(0)                              'Rabatt

235                 If rs.RecordCount > 0 Then txt1(12).text = rs.Fields("Währung").value   'Waehrung
240                 txt1(8).text = 0                                                        'Kostenstelle
245                 txt1(9).text = 0                                                        'FiBu Konto

250                 gbDataChanged = False

255                 Check1(0).value = 1

260                 rs.Close

265                 txt1(1).SetFocus
270                 SatzAufbereiten = True

275                 gbDataChanged = True
                Else
280                 SatzAufbereiten = False
285                 gbSatzNeu = False
290                 txt1(0).SetFocus
295                 txt1(0).SelStart = 0
300                 txt1(0).selLength = Len(txt1(0).text)
                End If
  
                'Alte Logik erstmal drin gelassen

                '430       If Trim(txt1(Index)) <> "" Then
                '440         If gRS.RecordCount > 0 Then
                '450           gRS.MoveFirst
                '460           gRS.Find "[Schl] = '" & Trim(txt1(Index).Text) & "'"
                '
                '470           If Not gRS.EOF And Not gRS.BOF Then
                '480             If gbSatzNeu Then
                '490               Call msgText(2, 39, 269, 0, 0)
                '500               MsgBox GsMsgText(0) & " '' " & txt1(Index) & " '' " & GsMsgText(1), vbInformation
                '                  'MsgBox "Der Artikel ist bereits vergeben!"
                '510               Call SatzZeigen
                '520               gbSatzNeu = False
                '530               SatzAufbereiten = True
                '540             Else
                '550               Call SatzZeigen
                '560               gvntMerker = txt1(0)
                '570               SatzAufbereiten = True
                '580             End If
                '590           Else
                '600             If gbSatzNeu Then
                '610               antwort = vbYes
                '620             Else
                '630               Call msgText(2, 39, 29, 0, 0)
                '640               antwort = MsgBox(GsMsgText(0) & " '' " & txt1(Index) & " '' " & GsMsgText(1), vbYesNo + vbInformation)
                '                  'antfort = MsgBox("Der Artikel " & txt1(Index) & " existiert nicht. Möchten Sie ihn Hinzufügen?", vbYesNo + vbInformation)
                '650             End If
                '660             If antwort = vbYes Then
                '                  'DH, 24.10.2017, 6.5.101, Bei Neuanlage soll die Waehrung aus dem Mandantenstamm gezogen werden
                '670               query = ""
                '680               query = query & "SELECT Wrg.Währung "
                '690               query = query & "FROM [1100_Mandant] AS Mandant "
                '700               query = query & "LEFT JOIN [1100_Währungen] AS Wrg ON(Wrg.Schl = Mandant.EigenWrg)"
                '
                '710               rs.Open query, gConn, adOpenStatic, adLockReadOnly
                '
                '720               gRS.AddNew
                '730               gRS.Fields(0).Value = Trim(txt1(0))
                '740               If rs.RecordCount > 0 Then gRS.Fields("Wrg").Value = rs.Fields("Währung").Value
                '750               gRS.Update
                '760               gRS.MoveLast
                '770               gbSatzNeu = False
                '780               gvntMerker = txt1(0)
                '790               Call SatzZeigen
                '800               SatzAufbereiten = True
                '
                '810               rs.Close
                '820             Else
                '830               SatzAufbereiten = False
                '840             End If
                '850           End If
                '860         Else
                '870           If gbSatzNeu Then
                '880             antwort = vbYes
                '890           Else
                '900             Call msgText(2, 39, 29, 0, 0)
                '910             antwort = MsgBox(GsMsgText(0) & " '' " & txt1(Index) & " '' " & GsMsgText(1), vbYesNo + vbInformation)
                '                'antfort = MsgBox("Der Artikel " & txt1(Index) & " existiert nicht. Möchten Sie ihn Hinzufügen?", vbYesNo + vbInformation)
                '920           End If
                '
                '930           If antwort = vbYes Then
                '940             gRS.AddNew
                '
                '950             gRS.Fields(0).Value = Trim(txt1(0))
                '960             gRS.Update
                '970             gRS.MoveLast
                '980             gvntMerker = txt1(0)
                '990             Call SatzZeigen
                '1000            SatzAufbereiten = True
                '1010          Else
                '1020            SatzAufbereiten = False
                '1030          End If
                '1040        End If
                '1050      Else
                '1060        SatzAufbereiten = False
                '1070      End If

305         Case Else
310             SatzAufbereiten = True
        End Select

        Exit Function

Fehler:
315     Call FehlerErklärung("frmSP52820", "SatzAufbereiten()")

320     If rs.state = adStateOpen Then rs.Close
End Function

Private Sub txt1_MouseDown(Index As Integer, _
                           Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           Y As Single)

        On Error GoTo Fehler

        'Debug.Print "Mouse Down"
        'denyChangeControl = True       'DH, 24.10.2017, 6.5.101, Auskommentiert und somit die Maus wieder zugelassen.

        Exit Sub

Fehler:
100     Call FehlerErklärung("frmSP52820", "txt1_MouseDown()")
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)

        On Error GoTo Fehler

100     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index

105     Select Case Index

            Case 4              'Preis
            
110             txt1(Index).text = Format(txt1(Index).text, ZahlFormat(postCommaPreis))

115         Case 6, 7

                Dim result As Integer

120             result = PlausiByIndex(Index) 'Plausi

125             If result <> -1 Then
130                 Cancel = True
135                 txt1(result).SetFocus

                End If

140         Case Else           'Alle sonstigen Felder

145             txt1(Index).text = objPRM.EingabeUmwandlung(txt1(Index))

        End Select

150     If objPRM.EingabeFehler(txt1(Index)) Then

155         Cancel = True

            '<Added by: IL at: 8.28.2024-09:35:52 on machine: T017>
160     ElseIf Index = 3 Then

165         If PlausiF2(Index) = -1 Then

170             If Not CheckENCodeBeiVerpackung(0, txt1(Index).text) Then Cancel = True

            Else
            
175             Cancel = True

            End If

            '</Added by: IL at: 8.28.2024-09:35:52 on machine: T017>
 
        End If
  
        Exit Sub

Fehler:
180     Call FehlerErklärung("frmSP52820", "txt1_Validate()")
End Sub

Public Sub SatzKopieren(NeuerSchluessel As String, _
                        ZielMandant As String) 'IL 27.02.2025 , Ver.: 6.7.106 : ZielMandant As String
                        
        Dim rs As ADODB.Recordset

        On Error GoTo Fehler

100     If Trim(NeuerSchluessel) <> "" Then

105         If ZielMandant = GsAnwenderNr Then
 
110             Set rs = gRS
 
            Else

                Dim conn As ADODB.Connection
                        
115             Set conn = New ADODB.Connection
120             Set rs = New ADODB.Recordset

125             conn.Open GetConnectionString(GsHauptPfad, Spedifix, ZielMandant)
130             rs.Open "SELECT * FROM [2800_Artikel] ORDER BY [Schl]", conn, adOpenKeyset, adLockOptimistic

            End If

135         If rs.RecordCount > 0 Then
                
140             If rs.BOF = False Then rs.MoveFirst
                
145             rs.Find "[Schl] = '" & Trim(NeuerSchluessel) & "'"

                '115             If Not gRS.EOF And gRS.BOF Then
                
150             If rs.EOF = False Then

155                 Call msgText(2, 39, 269, 0, 0)
160                 MsgBox GsMsgText(0) & " '' " & Trim(NeuerSchluessel) & " '' " & GsMsgText(1), vbInformation, strMeldungCap
                    'MsgBox "Der Artikel ist bereits vergeben!"
                    
                Else
                
165                 gbDataChanged = False
170                 txt1(0) = NeuerSchluessel
175                 gbDataChanged = True
                    '<Added by: DFiebach at: 15.01.2019, Ver.: 6.5.109 >
                    ' # Ohne diesen Zeiger, wird ein Fehler beim Speichern geworfen, da der Datensatz nicht neu hizugefügt wird.
180                 gbSatzNeu = True
                    '</Added by: DFiebach at: 15.01.2019, Ver.: 6.5.109 >

                    '<Modified by: IL at 27.02.2025, Ver.: 6.7.106 >

185                 Call SatzSpeichern(True, rs)

                    '</Modified by: IL at 27.02.2025, Ver.: 6.7.106 >

                End If
                
            End If
            
        End If

        Exit Sub

Fehler:
190     Call FehlerErklärung("frmSP52820", "SatzKopieren")
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

Private Function Plausi() As Integer
   
        On Error GoTo Fehler
        
        Dim i      As Integer
        
        Dim result As Integer

100     result = -1
        
105     For i = txt1.Count - 1 To 0 Step -1
            
110         If txt1(i).Index = 6 Or txt1(i).Index = 7 Or txt1(i).Index = 3 Then  'IL 27.08.2024
115             objPRM.FindFirstString = "name = 'txt1' AND index = " & i
120             txt1(i) = objPRM.EingabeUmwandlung(txt1(i))

125             If objPRM.EingabeFehler(txt1(i)) Then

130                 result = i

                Else

135                 result = PlausiF2(i)

                End If

                'result = PlausiF2(i)
                    
140             If result <> -1 Then

                    Exit For
                
                End If
                        
            End If
            
        Next
        
145     Plausi = result
        
        Exit Function

Fehler:
150     Call FehlerErklärung("frmSP52820", "Plausi()")

End Function

Private Function PlausiF2(Index As Integer) As Integer
  
        On Error GoTo Fehler

        Dim sql       As String

        Dim objPlausi As clsPlausi

        Dim result    As Integer
        
100     result = Index

105     If objPlausi Is Nothing Then
110         Set objPlausi = New clsPlausi
115         objPlausi.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
        End If
        
120     Select Case Index

                '<Added by: IL at: 8.27.2024-15:50:27 on machine: T017>
            Case 3

125             If Trim(txt1(Index)) = "" And objPRM.PRM_Inhalt("txt1", "pflicht", Index) = 0 Then
130                 result = -1
                Else

135                 sql = "SELECT Schl, Bezeichnung FROM [1400_Verpackungen] WHERE Schl = '" & Trim(txt1(Index).text) & "'"
                End If

                '</Added by: IL at: 8.27.2024-15:50:27 on machine: T017>
        
140         Case 6

145             If Trim(txt1(Index)) = "" And objPRM.PRM_Inhalt("txt1", "pflicht", Index) = 0 Then
                
150                 result = -1
                Else
155                 sql = "SELECT Schl,Bezeichnung,Erlöse,Kosten FROM [1100_FiBuKostenStellen] WHERE Schl = '" & txt1(Index) & "'"
                End If

160         Case 7

165             If Trim(txt1(Index)) = "" And objPRM.PRM_Inhalt("txt1", "pflicht", Index) = 0 Then
170                 result = -1
                Else

175                 sql = "SELECT Schl,Bezeichnung,Erlöse,Kosten FROM [1100_FiBuSachkonten] WHERE Schl = '" & txt1(Index) & "'"
                End If

        End Select
        
180     If result <> -1 Then
            
185         If objPlausi.RSOpen(sql, True) = True Then

190             result = -1

195             txt1(Index).text = objPlausi.ValueFromRsSQL                    ' IL 29.08.2024

            End If
        
        End If

200     PlausiF2 = result

        Exit Function

Fehler:
205     Call FehlerErklärung("frmSP52820", "PlausiF2()")
   
End Function

Public Function Speichern() As Boolean

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       Speichern
        ' Description:       [type_description_here]
        ' Created by :       GW
        ' Date-Time  :       14.08.2020-15:09:26
        '
        ' Parameters :
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler

        Dim Index     As Integer

        Dim result    As Boolean
        
        Dim objPlausi As clsPlausi

100     result = False
105     Index = Plausi

110     If Index = -1 Then

115         If CheckENCodeBeiVerpackung(0, txt1(3).text) Then

120             result = SatzSpeichern

            Else
            
125             txt1(3).SetFocus

            End If

        Else
        
130         txt1(Index).SetFocus

        End If

135     Speichern = result

        Exit Function

Fehler:
140     Call FehlerErklärung("frmSP52820", "Speichern()")
   
End Function

Private Function PlausiByIndex(Index As Integer) As Integer
   
        On Error GoTo Fehler
                
        Dim result As Integer

100     result = -1
            
105     If txt1(Index).Index = 6 Or txt1(Index).Index = 7 Then
110         objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
115         txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

120         If objPRM.EingabeFehler(txt1(Index)) Then
125             result = Index
            End If

130         result = PlausiF2(Index)
                    
        End If
   
135     PlausiByIndex = result
        
        Exit Function

Fehler:
140     Call FehlerErklärung("frmSP52820", "Plausi()")

End Function
