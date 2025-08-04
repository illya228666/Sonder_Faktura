VERSION 5.00
Begin VB.Form frmSP52821 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Kopieren"
   ClientHeight    =   1500
   ClientLeft      =   5535
   ClientTop       =   5085
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSP52821.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt1 
      BorderStyle     =   0  'Kein
      Height          =   200
      Index           =   1
      Left            =   1770
      MaxLength       =   12
      TabIndex        =   4
      Top             =   720
      Width           =   2540
   End
   Begin VB.TextBox txt1 
      BorderStyle     =   0  'Kein
      Height          =   200
      Index           =   0
      Left            =   1770
      MaxLength       =   12
      TabIndex        =   0
      Top             =   1080
      Width           =   2540
   End
   Begin VB.CommandButton cmdNeu 
      Caption         =   "OK"
      Height          =   300
      Index           =   0
      Left            =   4875
      TabIndex        =   1
      Top             =   1050
      Width           =   1260
   End
   Begin VB.CommandButton cmdNeu 
      Caption         =   "Abbrechen"
      Height          =   300
      Index           =   1
      Left            =   6180
      TabIndex        =   2
      Top             =   1050
      Width           =   1260
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bitte wählen Sie einen M-Code für den neu zu erzeugenden Stammsatz"
      Height          =   390
      Index           =   0
      Left            =   150
      TabIndex        =   6
      Top             =   240
      Width           =   5820
   End
   Begin VB.Label lbl0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Neuer Suchbegriff"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label lbl0 
      Caption         =   "Neuer Suchbegriff"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   1080
      Width           =   1395
   End
End
Attribute VB_Name = "frmSP52821"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objPRM As clsPRM

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
Dim cReSize        As FormResize 'HW 03.02.2011

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

Private Sub cmdNeu_Click(Index As Integer)

100     If Index = 0 Then

105         If Trim(txt1(0)) <> "" Then

110             If IsSQLConnectionAlive(GetConnectionString(GsHauptPfad, Spedifix, txt1(1))) Then

115                 Call frmSP52820.SatzKopieren(Trim(txt1(0)), txt1(1))
                
                Else
                
120                 MsgBox GetMessage(2395), vbExclamation, strMeldungCap
                
                    Exit Sub
                
                End If
                
            End If
            
        End If
  
125     Unload Me
End Sub

Private Sub Form_Load()

        Dim Parenthwnd As Long

        Dim i          As Integer
        
100     Set objPRM = New clsPRM
105     Set objPRM.gForm = Me
110     objPRM.PRM_Alle

        '########### SkinFramework ##############################
        'HW 04.03.2011 - Is ein Windows Skin Tool von Codejock Software
        'SkinFramework1.LoadSkin GsHauptPfad & "\exe\Spedifix.cjstyles", "NormalSilver.ini"
        'SkinFramework1.ApplyWindow Me.hWnd
        '########################################################

        '########## Subclassing: Messages festlegen #############
        ' DeW, ZyG Mai 2011
115     AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO 'DeW
120     AttachMessage Me, Me.hwnd, WM_EXITSIZEMOVE 'DeW
        '
        '########################################################

        '####### Subclassing: Groessenbegrenzung Formular #######
        ' TODO in Arbeit, MagicNumbers
        'dieser Aufruf kann je nach Programm-Modul woanders in der
        'Form_Load Methode stehen!
        'Zuerst muss der "alte" Code die Zuweisung von Breit und
        'Hoehe korrekt vorgenommen haben!
125     SetMinMaxInfo Me.hwnd, Me.height, (Me.height * 2), Me.width, (Me.width * 2)
        '
        '########################################################

130     Parenthwnd = FindWindow(vbNullString, frmSP52820.caption)
135     SetWindowLong Me.hwnd, GWL_HWNDPARENT, Parenthwnd

140     Me.Top = frmSP52820.Top + (frmSP52820.height - Me.height) / 2
145     Me.Left = frmSP52820.Left + (frmSP52820.width - Me.width) / 2

150     txt1(1).text = GsAnwenderNr                                            'IL 27.02.2025 , Ver.: 6.7.106 :

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        On Error GoTo Fehler

        '####### Subclassing: Messages austragen #############
        'DeW, Mai 2011
100     DetachMessage Me, Me.hwnd, WM_EXITSIZEMOVE
105     DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
        '
        '#####################################################

        '####### Subclassing: Groessenbegrenzung loeschen #######
        'DeW, Mai 2011
110     RemoveMinMaxInfo Me.hwnd
        '
        '########################################################

        'DH, 07.02.2014, 6.2.102, Das Fenster nicht mehr Schliessen (Unload) da sich sonst der Skinner verabschiedet.
        '                         Unload geschieht jetzt nur noch, wenn sich das Hauptfenster schliesst (UnloadMode = 5)
115     If UnloadMode <> 5 Then
120         txt1(0).text = ""       'Beim Schliessen des Fensters das Textfeld leeren
125         Me.Hide
        End If

        Exit Sub

Fehler:
130     Call FehlerErklärung("frmSP5281", "Form_QueryUnload()")
End Sub

Private Sub Form_Unload(Cancel As Integer)

        On Error GoTo Fehler

        '########## Formular Resizing: stoppen###################
        '
        'DeW, folgendes terminiert die Klasse, und loest dort
        'das _Terminate Ereigniss aus -> Speicherung der eingestellten
        'Vergroesserungswerte und Spaltenbreiten aus den TrueDBGrid
        'Info-Daten in der Registry
        '
100     Set cReSize = Nothing
        '
        '########################################################

        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52821", "Form_Unload()")
End Sub

Public Sub ShowMe(FactorWidth As Single, FactorHeight As Single)

        '###### Formular Resizing: Parameter setzen#############
        ' DeW, ZyG, Mai 2011
        'Section- oder KeyBezeichnung sind in vielen Faellen in
        'altem Code hart eincodiert worden, manchmal wird auch
        'eine Variable verwendet...
100     Set cReSize = New FormResize
105     cReSize.setSectionBezeichnung = "SP52820"
110     cReSize.setKeyBezeichnung = "SP52821"
115     cReSize.setIstUnterFenster = True
        '
        '########################################################

120     If FactorWidth > 0 Then cReSize.CurrScaleFactorWidth = FactorWidth
125     If FactorHeight > 0 Then cReSize.CurrScaleFactorHeight = FactorHeight

        '
        'Speichere keine Informationen (Spaltenbreiten usw.) fuer die
        'Tabellen im Form, wenn z.B. nur eine einzelne Tabelle
        'vorhanden ist, die jeweils mit neuen Daten gefuellt
        'und an eine andere Position verschoben wird (z.B. SP51000
        'Mandantenstamm
130     cReSize.IgnoreTrueDBGridInfo = True
        '
135     cReSize.Form = Me

140     Me.Show vbModal
   
End Sub

Private Sub txt1_GotFocus(Index As Integer)
 
        On Error GoTo Fehler
    
100     txt1(Index).SelStart = 0
105     txt1(Index).selLength = Len(txt1(Index))
110     txt1(Index).ForeColor = vbBlack
115     txt1(Index).BackColor = &HC0E0FF  'hellorange

125     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index & " AND  pos1 = 0"

        Exit Sub
    
Fehler:
135     Call FehlerErklärung("frmSP52821", "txt1_GotFocus()")

End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 
        On Error GoTo Fehler
   
100     Select Case KeyCode
    
            Case vbKeyF1
           
105         Case vbKeyReturn, vbKeyDown
                
110             If objPRM.EingabeFehler(txt1(Index)) = False Then
115                 Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                End If
    
        End Select
   
        Exit Sub
    
Fehler:
120     Call FehlerErklärung("frmSP52821", "txt1_KeyDown()")
  
End Sub

Private Sub txt1_LostFocus(Index As Integer)
   
        On Error GoTo Fehler
    
100     txt1(Index).ForeColor = vbActiveTitleBar
105     txt1(Index).BackColor = vbWindowBackground  'Fensterhintergrund(weiß)
    
        Exit Sub
    
Fehler:
110     Call FehlerErklärung("frmSP52821", "txt1_LostFocus()")
    
End Sub
