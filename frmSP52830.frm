VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSP52830 
   Caption         =   "Sonderfaktura"
   ClientHeight    =   5670
   ClientLeft      =   17790
   ClientTop       =   9870
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmSP52830.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9315
   Tag             =   "1"
   Begin VB.Frame Frame1 
      Caption         =   "FiBu"
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
      TabIndex        =   71
      Top             =   3450
      Width           =   4665
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   26
         Left            =   2590
         Picture         =   "frmSP52830.frx":0442
         Style           =   1  'Grafisch
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   470
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   25
         Left            =   2590
         Picture         =   "frmSP52830.frx":0488
         Style           =   1  'Grafisch
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   25
         Left            =   1530
         TabIndex        =   33
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   26
         Left            =   1530
         TabIndex        =   34
         Top             =   480
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   27
         Left            =   4020
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   28
         Left            =   1530
         TabIndex        =   36
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auftragsbest."
         Height          =   195
         Index           =   3
         Left            =   3960
         TabIndex        =   129
         Top             =   450
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblDummy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angebot"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   128
         Top             =   210
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblDummy 
         Caption         =   "Rechnung"
         Height          =   195
         Index           =   0
         Left            =   3060
         TabIndex        =   108
         Top             =   210
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblDummy 
         Caption         =   "Gutschrift"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   107
         Top             =   450
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Kostenst.-Schl."
         Height          =   195
         Index           =   25
         Left            =   150
         TabIndex        =   75
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Sachkonten-Schl."
         Height          =   195
         Index           =   26
         Left            =   150
         TabIndex        =   74
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Kostenst.-Konto"
         Height          =   195
         Index           =   27
         Left            =   2640
         TabIndex        =   73
         Top             =   720
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Sach-Konto"
         Height          =   195
         Index           =   28
         Left            =   150
         TabIndex        =   72
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Statistik"
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
      Index           =   3
      Left            =   4800
      TabIndex        =   58
      Top             =   3450
      Width           =   4490
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   36
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1200
         Width           =   225
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   36
         Left            =   1620
         Picture         =   "frmSP52830.frx":04CE
         Style           =   1  'Grafisch
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   1190
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   35
         Left            =   2590
         Picture         =   "frmSP52830.frx":0514
         Style           =   1  'Grafisch
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   950
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   31
         Left            =   2590
         Picture         =   "frmSP52830.frx":055A
         Style           =   1  'Grafisch
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   710
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   30
         Left            =   2590
         Picture         =   "frmSP52830.frx":05A0
         Style           =   1  'Grafisch
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   470
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   29
         Left            =   2590
         Picture         =   "frmSP52830.frx":05E6
         Style           =   1  'Grafisch
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   33
         Left            =   3195
         TabIndex        =   44
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   32
         Left            =   2980
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   33
         Left            =   4260
         Picture         =   "frmSP52830.frx":062C
         Style           =   1  'Grafisch
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1185
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   35
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   960
         Width           =   1185
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   29
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   30
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   480
         Width           =   1185
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   31
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label lbl1 
         Caption         =   "Kosten-Art"
         Height          =   195
         Index           =   36
         Left            =   150
         TabIndex        =   106
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label lbl1 
         Caption         =   "Kfz-Suchbegriff"
         Height          =   195
         Index           =   33
         Left            =   1950
         TabIndex        =   100
         Top             =   1200
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl1 
         Caption         =   "Relation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   32
         Left            =   3180
         TabIndex        =   99
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lbl1 
         Caption         =   "Sendungs-Nr."
         Height          =   195
         Index           =   35
         Left            =   150
         TabIndex        =   97
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label lbl1 
         Caption         =   "Abf.-Datum"
         Height          =   195
         Index           =   29
         Left            =   150
         TabIndex        =   70
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lbl1 
         Caption         =   "Abf.-Zusatz"
         Height          =   195
         Index           =   30
         Left            =   150
         TabIndex        =   63
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lbl1 
         Caption         =   "Abf.-Position"
         Height          =   195
         Index           =   31
         Left            =   150
         TabIndex        =   62
         Top             =   720
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Beleg"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Index           =   2
      Left            =   4800
      TabIndex        =   57
      Top             =   30
      Width           =   4485
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   200
         Index           =   43
         Left            =   2880
         TabIndex        =   132
         Text            =   "wird nicht verwendet"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.CommandButton cmdKlar 
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
         Left            =   2590
         Picture         =   "frmSP52830.frx":0672
         Style           =   1  'Grafisch
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   39
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1190
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
         Index           =   1
         Left            =   2590
         Picture         =   "frmSP52830.frx":0984
         Style           =   1  'Grafisch
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   42
         Left            =   1400
         TabIndex        =   30
         Top             =   2400
         Width           =   225
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   38
         Left            =   4040
         Picture         =   "frmSP52830.frx":09CA
         Style           =   1  'Grafisch
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   2150
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   37
         Left            =   2590
         Picture         =   "frmSP52830.frx":0A10
         Style           =   1  'Grafisch
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   2150
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   37
         Left            =   1400
         TabIndex        =   28
         Top             =   2160
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   38
         Left            =   3100
         TabIndex        =   29
         Top             =   2160
         Width           =   950
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   24
         Left            =   4665
         TabIndex        =   95
         Top             =   3240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   16
         Left            =   1395
         TabIndex        =   32
         Top             =   3000
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   15
         Left            =   1395
         TabIndex        =   31
         Top             =   2880
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ausweisen"
         Height          =   200
         Index           =   1
         Left            =   3340
         TabIndex        =   21
         Top             =   720
         Width           =   1070
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   18
         Left            =   5640
         Picture         =   "frmSP52830.frx":0A56
         Style           =   1  'Grafisch
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   2985
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
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
         Index           =   17
         Left            =   2590
         Picture         =   "frmSP52830.frx":0A9C
         Style           =   1  'Grafisch
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   470
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Vorlage"
         Height          =   200
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   2650
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   17
         Left            =   1400
         TabIndex        =   19
         Top             =   480
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   18
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   19
         Left            =   1400
         TabIndex        =   23
         Top             =   1200
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H8000000A&
         Height          =   200
         Index           =   20
         Left            =   2880
         TabIndex        =   24
         Top             =   960
         Width           =   470
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   21
         Left            =   1400
         TabIndex        =   25
         Top             =   1440
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   22
         Left            =   1400
         TabIndex        =   26
         Top             =   1680
         Width           =   1190
      End
      Begin VB.TextBox txt1 
         Alignment       =   1  'Rechts
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   23
         Left            =   1400
         TabIndex        =   27
         Top             =   1920
         Width           =   1190
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "status"
         Height          =   195
         Index           =   7
         Left            =   1395
         TabIndex        =   127
         Top             =   2640
         Width           =   2745
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Rechnung*"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   126
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Enabled         =   0   'False
         Height          =   200
         Index           =   20
         Left            =   3350
         TabIndex        =   124
         Top             =   960
         Width           =   150
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
         Height          =   200
         Index           =   18
         Left            =   2880
         TabIndex        =   123
         Top             =   720
         Width           =   1310
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Enabled         =   0   'False
         Height          =   200
         Index           =   17
         Left            =   2880
         TabIndex        =   122
         Top             =   480
         Width           =   1310
      End
      Begin VB.Label lbl1 
         Caption         =   "Dez.-Stellen"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   120
         Top             =   2400
         Width           =   1185
      End
      Begin VB.Label lbl1 
         Caption         =   "bis"
         Height          =   200
         Index           =   43
         Left            =   2860
         TabIndex        =   116
         Top             =   2160
         Width           =   260
      End
      Begin VB.Label lbl1 
         Caption         =   "von"
         Height          =   195
         Index           =   42
         Left            =   1100
         TabIndex        =   115
         Top             =   2160
         Width           =   250
      End
      Begin VB.Label lbl1 
         Caption         =   "Zeitraum"
         Height          =   195
         Index           =   37
         Left            =   150
         TabIndex        =   109
         Top             =   2160
         Width           =   700
      End
      Begin VB.Label lbl1 
         Caption         =   "Porto vom Stamm"
         Height          =   195
         Index           =   24
         Left            =   4500
         TabIndex        =   96
         Top             =   3090
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Beleg-Datum"
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   87
         Top             =   3000
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Beleg-Nr."
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   86
         Top             =   2880
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Beleg-Währung"
         Height          =   195
         Index           =   17
         Left            =   150
         TabIndex        =   69
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Mand.-Währung"
         Height          =   195
         Index           =   18
         Left            =   150
         TabIndex        =   68
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Kurs"
         Height          =   195
         Index           =   19
         Left            =   150
         TabIndex        =   67
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Steuersatz"
         Height          =   195
         Index           =   20
         Left            =   150
         TabIndex        =   66
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Skonto %"
         Height          =   195
         Index           =   21
         Left            =   150
         TabIndex        =   61
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Skonto-Tage"
         Height          =   195
         Index           =   22
         Left            =   150
         TabIndex        =   60
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Tage-Netto"
         Height          =   195
         Index           =   23
         Left            =   150
         TabIndex        =   59
         Top             =   1920
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kunde"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3400
      Index           =   0
      Left            =   60
      TabIndex        =   49
      Top             =   30
      Width           =   4665
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
         Left            =   2990
         Picture         =   "frmSP52830.frx":0AE2
         Style           =   1  'Grafisch
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   8
         Left            =   1875
         Picture         =   "frmSP52830.frx":0B28
         Style           =   1  'Grafisch
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   41
         Left            =   1530
         TabIndex        =   17
         Top             =   3120
         Width           =   2925
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   40
         Left            =   5370
         TabIndex        =   110
         Top             =   3355
         Visible         =   0   'False
         Width           =   2025
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
         Index           =   40
         Left            =   7410
         Picture         =   "frmSP52830.frx":0B6E
         Style           =   1  'Grafisch
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3343
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "ändern"
         Height          =   195
         Index           =   2
         Left            =   3660
         TabIndex        =   1
         Top             =   250
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   34
         Left            =   1530
         TabIndex        =   3
         Top             =   480
         Width           =   225
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   34
         Left            =   1760
         Picture         =   "frmSP52830.frx":0DB8
         Style           =   1  'Grafisch
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   470
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   12
         Left            =   2010
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   9
         Left            =   3240
         Picture         =   "frmSP52830.frx":0DFE
         Style           =   1  'Grafisch
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2150
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   11
         Left            =   4240
         Picture         =   "frmSP52830.frx":0E44
         Style           =   1  'Grafisch
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2630
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Left            =   4240
         Picture         =   "frmSP52830.frx":0E8A
         Style           =   1  'Grafisch
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   2390
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Left            =   4240
         Picture         =   "frmSP52830.frx":0ED0
         Style           =   1  'Grafisch
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1670
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Index           =   5
         Left            =   3550
         Picture         =   "frmSP52830.frx":0F16
         Style           =   1  'Grafisch
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1430
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.CommandButton cmdAuswahl 
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
         Left            =   4240
         Picture         =   "frmSP52830.frx":0F5C
         Style           =   1  'Grafisch
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1190
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   3
         Left            =   1530
         TabIndex        =   6
         Top             =   1200
         Width           =   2715
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   13
         Left            =   3030
         TabIndex        =   16
         Top             =   2880
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   14
         Left            =   1530
         TabIndex        =   15
         Top             =   2880
         Width           =   1425
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   11
         Left            =   1530
         TabIndex        =   14
         Top             =   2640
         Width           =   2715
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   1
         Left            =   1530
         TabIndex        =   4
         Top             =   720
         Width           =   2925
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   10
         Left            =   1530
         TabIndex        =   13
         Top             =   2400
         Width           =   2715
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   9
         Left            =   2190
         TabIndex        =   12
         Top             =   2160
         Width           =   1050
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   8
         Left            =   1530
         TabIndex        =   11
         Top             =   2160
         Width           =   345
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   195
         Index           =   7
         Left            =   1530
         TabIndex        =   10
         Top             =   1920
         Width           =   2925
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   6
         Left            =   1530
         TabIndex        =   9
         Top             =   1680
         Width           =   2715
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   5
         Left            =   2580
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   4
         Left            =   1530
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         DataSource      =   "Data1"
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   2
         Left            =   1530
         TabIndex        =   5
         Top             =   960
         Width           =   2925
      End
      Begin VB.TextBox txt1 
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000002&
         Height          =   200
         Index           =   0
         Left            =   1530
         TabIndex        =   0
         Top             =   240
         Width           =   1455
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
         Left            =   2990
         Picture         =   "frmSP52830.frx":0FA2
         Style           =   1  'Grafisch
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   230
         UseMaskColor    =   -1  'True
         Width           =   210
      End
      Begin VB.Label lbl1 
         Caption         =   "Interner Vermerk"
         Height          =   200
         Index           =   41
         Left            =   150
         TabIndex        =   114
         Top             =   3120
         Width           =   1370
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H003F9200&
         Height          =   500
         Left            =   90
         Top             =   1440
         Width           =   1400
      End
      Begin VB.Label lbl1 
         Caption         =   "Steuertext"
         Height          =   200
         Index           =   40
         Left            =   3870
         TabIndex        =   111
         Top             =   3480
         Visible         =   0   'False
         Width           =   1370
      End
      Begin VB.Label lbl1 
         Caption         =   "UID-/Steuer-Nr."
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   65
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Konto-Art / -Nr."
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   93
         Top             =   510
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Ortsteil"
         Height          =   195
         Index           =   11
         Left            =   150
         TabIndex        =   79
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "- Ort"
         ForeColor       =   &H003F9200&
         Height          =   195
         Index           =   6
         Left            =   175
         TabIndex        =   78
         Top             =   1680
         Width           =   1340
      End
      Begin VB.Label lbl1 
         Caption         =   "Name2"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   77
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Ansprechpartner"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   76
         Top             =   1230
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Umsatzsteuer"
         Height          =   200
         Index           =   14
         Left            =   150
         TabIndex        =   64
         Top             =   3480
         Visible         =   0   'False
         Width           =   1370
      End
      Begin VB.Label lbl1 
         Caption         =   "Ort"
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   56
         Top             =   2400
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Name1"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   55
         Top             =   750
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Lkz / Plz"
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   54
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label lbl1 
         Caption         =   "Straße"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   53
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "Postfach/Plz"
         ForeColor       =   &H003F9200&
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   52
         Top             =   1470
         Width           =   1365
      End
      Begin VB.Label lbl1 
         Caption         =   "MCode"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   51
         Top             =   270
         Width           =   1365
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   48
      Top             =   4920
      Width           =   9330
      _Version        =   65536
      _ExtentX        =   16457
      _ExtentY        =   635
      _StockProps     =   15
      ForeColor       =   -2147483630
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
         Caption         =   "Beleg Suchen"
         Height          =   300
         Index           =   8
         Left            =   6105
         TabIndex        =   119
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Beleg-Archiv"
         Height          =   300
         Index           =   7
         Left            =   60
         TabIndex        =   113
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "S&chließen"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   6
         Left            =   7970
         TabIndex        =   47
         Top             =   30
         Width           =   1290
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Rechnung"
         Height          =   300
         Index           =   5
         Left            =   4785
         TabIndex        =   45
         Top             =   30
         Width           =   1260
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "&Leeren"
         CausesValidation=   0   'False
         Height          =   300
         Index           =   0
         Left            =   3465
         TabIndex        =   46
         Top             =   30
         Width           =   1260
      End
   End
   Begin MSComctlLib.StatusBar sta1 
      Align           =   2  'Unten ausrichten
      Height          =   345
      Left            =   0
      TabIndex        =   85
      Top             =   5325
      Width           =   9315
      _ExtentX        =   16431
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
            Object.Width           =   13097
            MinWidth        =   12541
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDruckOption 
      Caption         =   " SPEDIFIX® | Sonderfaktura | Druck-Dialog"
      Height          =   135
      Left            =   9240
      TabIndex        =   112
      Top             =   240
      Width           =   735
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
         Caption         =   "Beleg"
         Index           =   0
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Speichern"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Neu"
         Index           =   5
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuBearb1 
         Caption         =   "Beleg suchen"
         Index           =   7
         Begin VB.Menu mnuSuch 
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
         Begin VB.Menu mnuSuch 
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
         Begin VB.Menu mnuSuch 
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
         Begin VB.Menu mnuSuch 
            Caption         =   "Archiv"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSuch 
            Caption         =   "Gutschrift"
            Index           =   4
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
      Begin VB.Menu mnuBearb1 
         Caption         =   "Beleg-Archiv"
         Index           =   8
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
         Caption         =   "Ansprechpartner übernehmen"
         Index           =   1
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "*Adresse(Beleg-Empf.) auf Folgeseiten drucken"
         Index           =   2
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "Bearbeiter (LoginName) drucken"
         Index           =   3
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "Gesamt = Brutto"
         Index           =   4
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "*Folgemaske autom. [leeren] nach DRUCK/SPEICHERN (nicht bei Vorlage)"
         Index           =   5
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "*Folgemaske autom. [schließen] nach DRUCK/SPEICHERN (nicht bei Vorlage)"
         Index           =   6
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "-"
         Index           =   49
      End
      Begin VB.Menu mnuOpt1 
         Caption         =   "Druckerauswahldialog"
         Index           =   50
      End
   End
   Begin VB.Menu mnuHlp 
      Caption         =   "&?"
      Begin VB.Menu mnuBesch 
         Caption         =   "Programmbeschreibung"
      End
      Begin VB.Menu mnuUpdateInfo 
         Caption         =   "&Update-Info"
      End
   End
   Begin VB.Menu mnuDummy 
      Caption         =   "Dummy"
      Index           =   0
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmSP52830"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Aenderungen ab August 2011
'
'DW 16.08.2011
' objSQLAusw.ColParameter 0, colWidth, (txt1(Index).Width / cResize.CurrScaleFactorWidth)
' den aktuellen ResizeFaktor mit eingebaut, sonst werden die Spalten im F2 Fenster zu gross!
'
Option Explicit
  
Public gintPrivBelegArt   As Integer

Public glngBelegID        As Long

Public glngBelegIDTmp     As Long                                               'HW 03.12.2015 ID-Variable für die Vorschau

Public glngBelegIDVorlage As Long                                               'DH, 01.06.2017, 6.4.126, Bei einer Vorlage wird diese ID bei Rechnungserzeugung nicht auf 0 gesetzt
  
Public gintDruck          As Integer                                            '0=nicht gedruckt, 1=gedruckt

Public gintZwAblage       As Integer
  
Private gbDataChanged     As Boolean

Private gbSatzNeu         As Boolean
  
Private gvntMerker        As Variant

Private gbEinfg           As Boolean

Private gstrEWerk         As String

Public objPRM             As clsPRM

Private objSQLAuswDef     As SPSQLAuswahl.clsSQLAuswahl

Private objDAOSeek        As SPDAOSeek.clsDAOSeek

Private objSQLAusw        As SPSQLAuswahl.clsSQLAuswahl                         'HW 06.05.2015
  
Private objHlp            As SpHlp.clsHlp

Private objLimit          As cLimit

Private objPlausi         As clsPlausi

Private clsUidWeb         As New clsUIDWebValidation                            'DF 11.20.2023 , Ver.: 6.6.128 : Überprüfung der UID

Private connSQL           As ADODB.Connection

Private dblMwSt           As Double                                             'IL 01.08.2024
  
'####### Subclassing ########################
'DeW, Mai 2011
'Variablen notwendig fuer Verwendung der SSubTmr Klasse,
'um eine schoenes Vergroesserung von Fenster und Inhalt und
'Begrenzung der Fenstergroesse zu ermoeglichen!
Implements ISubclass

Private emrConsume              As EMsgResponse
'
'DeW, notwendige WM_... Nachrichten fuer das
'Subclassing wurden als Public in SP50000B.bas
'definiert
'############################################

'<Modified by: GW at 21.02.2020, Ver.: GOBD >
'Private BelegArchiv             As SPBelegArchiv.clsBelegArchiv
Private objEmailSending         As clsEmailSending
'</Modified by: GW at 21.02.2020, Ver.: GOBD >

Public belegDatum               As String

Public BelegNr                  As String

Public ValutaDatum              As String

'####### Formular Resizing ##################
'
Public cReSize                  As FormResize                                   'HW 03.02.2011        'DH, 22.12.2015, 6.4.115, Public gesetzt, damit die Druckschleife darauf zugreifen kann

'
'############################################
Public printDone                As Boolean                                      'DH, 11.07.2013, Flag welches anzeigt ob ein Druck (nicht die Vorschau) ausgefuehrt wurde

Private shiftPressed            As Boolean                                      'DF 14.01.2015

Public gobjUpdateAenderung      As Integer                                      'DF 14.01.2015

Public gobjUpdateAenderungCount As Integer                                      'DF 14.01.2015

Public gi_UpdateAenderung       As Integer                                      'DF 14.01.2015

Public gi_UpdateInfoAngezeigt   As Boolean                                      'DF 14.01.2015

Public blnBelegNeu              As Boolean                                      'DF, 16.01.2015 : anzeigt, ob es sich um ein neuer Beleg handelt oder Beleg wurde aus Gedruckt/Ungedruckt aufgerufen.

'                'blnBelegNeu' muss dem Objekt frmRechnungErf/frmGutschriftErf vor dem Aufruf von frmRechnungErf.FolgeZeigen gesetz werden.

Private gTmpCaption             As String

Public gboolHatKind             As Boolean                                     'IL 05.11.2024: aktivieren, wenn das zweite Fenster geladen wird. Notwendig zur Steuerung der „Weiter“ und "Zurück"-Tasten

Public gboolBelegAngenommen As Boolean

Private Sub Check1_Click(Index As Integer)

        On Error GoTo Fehler

100     SetFormCaption

        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52830", "Check1_Click()")

End Sub

Private Sub cmdKlar_Click()

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       cmdKlar_Click
        ' Description:       Löschen die Felder im Abschnitt "Beleg".
        ' Created by :       IL
        ' Date-Time  :       18.10.2024-15:36:48
        '
        ' Parameters :
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler

        Dim i As Integer

100     For i = 0 To txt1.Count - 1

105         Select Case i
                    
                Case 19, 20, 21, 22, 23, 24                                     'Nummerischen Felder
110                 objPRM.FindFirstString = "name = 'txt1' AND index = " & i
115                 txt1(i) = objPRM.EingabeUmwandlung(0)

120             Case 37, 38
125                 txt1(i) = ""

            End Select

        Next
        
130     KursFuerRechnung "000"

135     txt1(21).SetFocus
        
        Exit Sub

Fehler:

140     Me.MousePointer = vbDefault
145     Call FehlerErklärung("frmSP52830", "cmdKlar_Click()")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

        On Error GoTo Fehler

100     If (Shift And vbShiftMask) = 0 Then shiftPressed = False

        Exit Sub

Fehler:
105     Call FehlerErklärung("frmSP52830", "Form_KeyUp")

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
  
'DH, 31.10.2013, 6.2.100
'Liefert den aktuell eingestellten Vergroesserungsfaktor dieses Fensters
Public Function getResizeFactor() As Double

        On Error GoTo Fehler

100     If mnuAnsicht_ResFak(0).Checked Then getResizeFactor = 1
105     If mnuAnsicht_ResFak(2).Checked Then getResizeFactor = 1.2
110     If mnuAnsicht_ResFak(4).Checked Then getResizeFactor = 1.4
115     If mnuAnsicht_ResFak(6).Checked Then getResizeFactor = 1.6
120     If mnuAnsicht_ResFak(8).Checked Then getResizeFactor = 1.8

        Exit Function

Fehler:

125     Call FehlerErklärung("frmSP56430", "getReiszeFactor()")

End Function

Sub AutomatischSteuerText(Index As Integer)

        On Error GoTo Fehler

        Dim nr          As String

        Dim Sort        As String

        Dim rec1100Text As ADODB.Recordset
  
100     OPEN_gConn
  
105     Set rec1100Text = New ADODB.Recordset

110     rec1100Text.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Sort = 1", gConn, adOpenStatic, adLockReadOnly

115     If rec1100Text.RecordCount > 0 Then
120         txt1(Index).text = "" & rec1100Text!titel
        End If

125     rec1100Text.Close
130     Set rec1100Text = Nothing

        Exit Sub

Fehler:
135     Call FehlerErklärung("frmSP52830", "AutomatischSteuerText(" & Index & ")")
End Sub

Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
  
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
130             objPRM.FindFirstString = "name = 'Check1' AND index = " & Index
135             Call objPRM.SprungNeu("Vorwärts", Shift, Check1(Index).TabIndex, True)

140         Case vbKeyEscape, vbKeyUp
145             objPRM.FindFirstString = "name = 'Check1' AND index = " & Index
150             Call objPRM.SprungNeu("Rückwerts", Shift, Check1(Index).TabIndex, True)
        End Select

        '***Beginn
        Exit Sub

Fehler:
155     Call FehlerErklärung("frmSP52830", "Check1_KeyDown")
        '***Ende

End Sub

Private Sub cmd1_Click(Index As Integer)

        On Error GoTo Fehler

        Dim i          As Integer
        
        Dim strMessage As String
        
        Dim antwort    As String
  
100     Select Case Index

            Case 0 'Leeren

105             If txt1(1) <> "" Then

                    '<Added by: GW at: 08.04.2019, Ver.: 6.5.110 >
                    'If MsgBox("Sie haben einige Daten erfasst. Soll die Eingabemaske trotzdem geleert werden?", vbYesNo + vbQuestion, strMeldungCap) = vbYes Then
110                 If MsgBox(GetMessage(523), vbYesNo + vbQuestion, strMeldungCap) = vbYes Then

115                     MaskeLeeren (False)

120                     txt1(0) = ""
125                     txt1(0).SetFocus
                        
130                     SetFormCaption (False)
                        '</Added by: GW at: 08.04.2019, Ver.: 6.5.110 >
                        
                    End If

                Else
                
135                 MaskeLeeren (False)

140                 txt1(0) = ""
145                 txt1(0).SetFocus

                    '<Added by: GW at: 08.04.2019, Ver.: 6.5.110 >
150                 SetFormCaption (False)
                    '</Added by: GW at: 08.04.2019, Ver.: 6.5.110 >
                    
                End If
                
155             If gboolHatKind Then
                
160                 Select Case gintPrivBelegArt

                        Case 0

165                         frmRechnungErf.MaskeLeeren ohneBelegDaten

170                         Unload frmRechnungErf
    
175                     Case 1

180                         frmGutschriftErf.MaskeLeeren ohneBelegDaten

185                         Unload frmGutschriftErf
    
190                     Case 2

195                         frmAngebotErf.MaskeLeeren ohneBelegDaten

200                         Unload frmAngebotErf
    
205                     Case 3

210                         frmAuftragsbestErf.MaskeLeeren ohneBelegDaten

215                         Unload frmAuftragsbestErf

                    End Select
                
                End If

220         Case 5 'Weiter
                
                'VORLAGE
225             If Check1(0).value = 1 Then                                     'Added by: GW at: 03.04.2019, Ver.: 6.5.110
                    
230                 MsgBox GetMessage(2187), vbExclamation, strMeldungCap       'Voragang kann nur als Vorlage bearbeitet werden

                End If
                
                'KTO-ART/-NUMMER ÜBERPRÜFEN
235             If Me.txt1(1) <> vbNullString Then

240                 If TxtBoxLeer(Me.txt1(34), 2192, objPRM, , , Me.txt1(1).text) Then Exit Sub

245                 If TxtBoxLeer(Me.txt1(12), 2192, objPRM, , , Me.txt1(1).text) Then Exit Sub

                End If
                
                'UID-ÜBERPRÜFUNG
250             If CheckUID(1) = False Then Exit Sub                            'DF 12.19.2023 , Ver.: 6.6.124

                'BELEG-NUMMER ÜBERPRÜFEN
255             If txt1(15).Visible Then

260                 If IsNumeric(txt1(15)) Then

265                     If Not IstBelegNrFrei(CLng(BelegNr), glngBelegID, gintPrivBelegArt) Then
                            
270                         strMessage = GetMessage(2173)
                            
275                         strMessage = Replace$(strMessage, "%1", BelegNr)

280                         MsgBox strMessage, vbExclamation, strMeldungCap

285                         gbDataChanged = False
                            
290                         gbDataChanged = True

                            Exit Sub

                        End If
                        
                    End If
                    
                End If

                'CSBmk <E-BELEG MANDANTEN-PFLICHTFELDER ÜBERPRÜFEN + HINWEISE>
295             If gEnmKudnenERechnungType <> eERechnungType.None And Check1(0).value = 0 And modERechnung.IsEBelegDoc Then         'DF 12.12.2024 , Ver.: 6.7.103
                
300                If modMandant.CheckEBelegMandantenFelder(True) = False Then Exit Sub   'Or CheckTlb = False
                   
305                If gEnmKudnenERechnungType = ZUGFeRD Then                                                                        'DF 04.02.2025 , Ver.: 6.7.105 : Himweis -> SAMMELERFASSUNG
                    
310                    MsgBox GetMessage(2390), vbOKOnly + vbExclamation, strMeldungCap
                    
                   End If
                   
                End If

315             i = Plausi
  
320             If i = 999 Then
                    
325                 If GesamtIstBrutto And gEnmKudnenERechnungType <> eERechnungType.None Then  'DF 04.09.2024 , Ver.: 6.7.101, Bei Brutto/Netto Umrechnung darf keine EREchnung erstellt werden, da die Option für die Privatkunden-Rehcnung gedacht ist.
                        
330                     MsgBox GetMessage(2358), vbInformation + vbOKOnly, strMeldungCap
                        
                    End If

                    '<Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
                    '# Wir prüfen, ob Fenster 2 bereits geöffnet ist, und wenn ja, führen wir die Funktionen aus, um es zu aktualisieren und sichtbar zu machen
335                 Me.Hide
                    
340                 If gboolHatKind Then

                        Dim lngNewBelegId As Long
                        
345                     lngNewBelegId = 0

350                     If gboolBelegAngenommen Then lngNewBelegId = glngBelegID

355                     Select Case gintPrivBelegArt

                            Case 0

360                             frmRechnungErf.TabellenAktualisieren lngNewBelegId
    
365                         Case 1

370                             frmGutschriftErf.TabellenAktualisieren lngNewBelegId
    
                                'frmGutschriftErf.Show vbModal, Me
    
375                         Case 2

380                             frmAngebotErf.TabellenAktualisieren lngNewBelegId
    
                                'frmAngebotErf.Show vbModal, Me
    
385                         Case 3

390                             frmAuftragsbestErf.TabellenAktualisieren lngNewBelegId
    
                                'frmAuftragsbestErf.Show vbModal, Me

                        End Select
                    
                    Else
                        '</Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
                    
395                     MousePointer = 11
                    
400                     gboolHatKind = True
                    
                        '<Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
405                     Select Case gintPrivBelegArt
                    
                            Case 0 'Rechnung
410                             Set frmRechnungErf = New frmSP52831
415                             MousePointer = vbDefault
420                             Set frmRechnungErf.frmParent = Me
425                             frmRechnungErf.BelegNeu = blnBelegNeu                   'DF 16.01.2015 : Beleg Alt\Neu erfasst
430                             frmRechnungErf.FolgeZeigen glngBelegID
435                             frmRechnungErf.Show vbModal, Me
                        
440                         Case 1 'Gutschrift

445                             Set frmGutschriftErf = New frmSP52831
450                             Set frmGutschriftErf.frmParent = Me
455                             frmGutschriftErf.BelegNeu = blnBelegNeu                 'DF 16.01.2015 : Beleg Alt\Neu erfasst
460                             frmGutschriftErf.FolgeZeigen glngBelegID
465                             frmGutschriftErf.Show vbModal, Me
                        
470                         Case 2 'Angebot

475                             Set frmAngebotErf = New frmSP52831
480                             Set frmAngebotErf.frmParent = Me
485                             frmAngebotErf.BelegNeu = blnBelegNeu
490                             frmAngebotErf.FolgeZeigen glngBelegID
495                             frmAngebotErf.Show vbModal, Me
                        
500                         Case 3 'Auftragsbestätigung

505                             Set frmAuftragsbestErf = New frmSP52831
510                             Set frmAuftragsbestErf.frmParent = Me
515                             frmAuftragsbestErf.BelegNeu = blnBelegNeu
520                             frmAuftragsbestErf.FolgeZeigen glngBelegID
525                             frmAuftragsbestErf.Show vbModal, Me
                    
                        End Select
                    
                    End If
                    
                    '265                 If gintPrivBelegArt = 0 Then                                'Rechnung
                    '270                     Set frmRechnungErf = New frmSP52831
                    '275                     Set frmRechnungErf.frmParent = Me
                    '280                     frmRechnungErf.BelegNeu = blnBelegNeu                   'DF 16.01.2015 : Beleg Alt\Neu erfasst
                    '285                     frmRechnungErf.FolgeZeigen glngBelegID
                    '290                     frmRechnungErf.Show vbModal, Me
                    '
                    '                    Else                                                        'Gutschrift
                    '
                    '295                     Set frmGutschriftErf = New frmSP52831
                    '300                     Set frmGutschriftErf.frmParent = Me
                    '305                     frmGutschriftErf.BelegNeu = blnBelegNeu                 'DF 16.01.2015 : Beleg Alt\Neu erfasst
                    '310                     frmGutschriftErf.FolgeZeigen glngBelegID
                    '315                     frmGutschriftErf.Show vbModal, Me
                    
                    'End If
                    '</Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
                    
530                 MousePointer = 0

                    On Error Resume Next

535                 If txt1(0).Enabled Then txt1(0).SetFocus

540                 If Err.number <> 0 Then Err.Clear

                    On Error GoTo Fehler

                Else

545                 If i <> 0 And i <> 25 And i <> 26 And i <> 17 Then          'DF 23.05.2023 , Ver.: 6.6.120 : um 0 für M-Code erweitert

550                     Call msgText(2, 116, 12, 0, 0)                          'Ein Pflichtfeld ist leer! Das Feld muss ausgefüllt werden.

555                     MsgBox GsMsgText(0) & " " & GsMsgText(1), vbExclamation, strMeldungCap

                    End If

                    On Error Resume Next

560                 If txt1(i).Enabled Then txt1(i).SetFocus

565                 If Err.number <> 0 Then Err.Clear

                    On Error GoTo Fehler

                End If

570         Case 6                                                              'Schließen

575             Unload Me

580         Case 7                                                              'Beleg-Archiv

585             Call mnuBearb1_Click(8)

590         Case 8                                                              'Übernahme (Beleg suchen)
       
595             Me.PopupMenu mnuBearb1(7), 0, cmd1(Index).left, cmd1(Index).top + cmd1(Index).height + SSPanel1(0).top    'DH, 19.101.2017, 6.5.101, Das Kontext Menue ging zu weit oben auf
    
        End Select

        Exit Sub

Fehler:
600     Call FehlerErklärung("frmSP52830", "cmd1_Click(" & Index & ")")
605     Err.Clear
End Sub

'DH, 05.10.2016, 6.4.122, Methode ueberarbeitet und Fehlerbehandlung hinzugefuegt
'DF 06.07.12 Erstellt. Liefert SteuerTyp entsprechend der übergebenen Parameter
Private Sub SteuerTyp(Zahl As Integer)

        On Error GoTo Fehler

        Dim connDef As ADODB.Connection

        Dim rsDEF   As ADODB.Recordset

100     Set connDef = New ADODB.Connection
105     Set rsDEF = New ADODB.Recordset

110     connDef.ConnectionString = GetACCESSConnectionString(DEF_CONNECTION)
115     connDef.Open

120     rsDEF.Open "SELECT Knz, " & GsDefFeld & " FROM [Auswahl] WHERE TabName = '1200_GrundKonditionen' AND FeldName = 'Ust' AND Knz = '" & Zahl & "'", connDef, adOpenStatic, adLockReadOnly

125     If rsDEF.RecordCount > 0 Then                       'Wenn die DEF einen entsprechenden Eintrag enthaelt
130         txt1(39).text = rsDEF.Fields(GsDefFeld).value
        Else                                                'Kein Eintrag vorhanden, den Int-Wert des Steuertyps in das Textfeld schreiben (darf eigentlich nie passieren)
135         txt1(39).text = CStr(Zahl)
        End If

140     rsDEF.Close
145     connDef.Close

        'HW 09.07.2012 Ver.: 6.1.114
150     intSteuerTyp = Zahl
        
155     Call ResetSteuerSatz(intSteuerTyp)                                      'DF 11.06.2020 , Ver.: 6.6.102 : Ust Senkung Umstellung
        
        Exit Sub

Fehler:
160     Call FehlerErklärung("frmSP52830", "SteuerTyp(" & Zahl & ")")

165     If rsDEF.state = adStateOpen Then rsDEF.Close
170     If connDef.state = adStateOpen Then connDef.Close
End Sub

Private Sub ResetSteuerSatz(intStKnz As Integer)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       ResetSteuerSatz
        ' Description:       Setzt die Anzeige des Steuersatzes bei Steuerfreien Belegen zurück
        ' Created by :       DFiebach
        ' Date-Time  :       11.06.2020-10:29:16
        '
        ' Parameters :       intStKnz (Integer)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler
        
100     Select Case intStKnz

            Case CInt(C_STR_STRFREI_EU), CInt(C_STR_STRFREI)
                 
105             objPRM.FindFirstString = "name = 'lbl2' AND index = 20"
110             txt1(20).text = objPRM.EingabeUmwandlung(CStr(0))
  
        End Select
            
        Exit Sub
    
Fehler:
    
115     Me.MousePointer = vbDefault
120     Call FehlerErklärung("frmSP52830", "ResetSteuerSatz()")
    
End Sub

Private Sub cmdAuswahl_Click(Index As Integer)

        On Error GoTo Fehler

        Dim i                 As Integer

        Dim ColLeft           As Long

        Dim RowBottom         As Long

        Dim OT                As String

        Dim sql               As String

        Dim Lkz               As String

        Dim strMessage        As String

        Dim intAntwort        As Integer

        Dim dt                As Date

        Dim rc                As rect
        
        Dim intIndexKorrektur As Integer                                        'DF 15.01.2019 , Ver.: 6.5.109 : Korrektur der Indexe der Textboxen, die zur Auswahlbuttons-Indexe nicht passen
        
        Dim bCancel           As Boolean
        
100     Select Case Index                                                       '<Added by: DFiebach at: 15.01.2019, Ver.: 6.5.109
        
            Case 1
                 
105             intIndexKorrektur = 39
                 
110         Case Else
                 
115             intIndexKorrektur = Index
       
        End Select
        
120     Call SetF2Position(txt1(intIndexKorrektur), Index, ColLeft, RowBottom)

125     objDAOSeek.ScaleFactorHeight = cReSize.CurrScaleFactorHeight
130     objDAOSeek.ScaleFactorWidth = cReSize.CurrScaleFactorWidth
135     objDAOSeek.fontSize = Me.fontSize

155     objSQLAusw.ScaleFactorHeight = cReSize.CurrScaleFactorHeight
160     objSQLAusw.ScaleFactorWidth = cReSize.CurrScaleFactorWidth
165     objSQLAusw.fontSize = Me.fontSize
170     objSQLAusw.SperrenFeld = ""

190     objSQLAuswDef.ScaleFactorHeight = cReSize.CurrScaleFactorHeight
195     objSQLAuswDef.ScaleFactorWidth = cReSize.CurrScaleFactorWidth
200     objSQLAuswDef.fontSize = Me.fontSize

220     If Index = 37 Or Index = 38 Then

225         dt = Date

230         g_objCal.BuddyHWnd = txt1(Index).hwnd

235         If g_objCal.GetData(dt) Then

240             txt1(Index).text = dt

245             objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
                'Weil objPRM.SprungNeu Validate-Ereignis nicht auslöst,
                'muss die Umwandlung und Prüfung an der Stelle stattfinden.
250             txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

                Dim erg As Boolean                                               'GW_05.03.2018 Ver.6.5.106 Datumvalidierung

255             erg = DatumUnterschiedRek(2, txt1(Index), txt1(37), 90, "d")

260             If erg = True Then

265                 txt1(Index) = ""

270                 If txt1(Index).Enabled Then txt1(Index).SetFocus
                Else

275                 If objPRM.EingabeFehler(txt1(Index)) = False Then

280                     Call objPRM.SprungNeu("Vorwärts", False, txt1(Index).TabIndex, True)

                    End If
                    
                End If

            Else

285             If txt1(Index).Enabled Then txt1(Index).SetFocus

            End If

            Exit Sub

        End If

290     If Index = 1 Or Index = 5 Or Index = 6 Or Index = 9 Or Index = 10 Or Index = 11 Then

295         If Trim(txt1(8)) = "" Then                                          'Zugrif auf E-Werk

300             Lkz = "D"

            Else
            
305             Lkz = txt1(8)

            End If

310         objDAOSeek.BorderStyle = 4

            'DeW, 03.08.2011, neu, auf "- " am Anfang kontrollieren, da
            'die - Ort / Plz Labels manchmal damit beginnen!
315         If (left(lbl1(Index), 2) Like "- ") Then
320             objDAOSeek.caption = Mid(lbl1(Index), 3)
            Else
325             objDAOSeek.caption = lbl1(Index)
            End If

330         objDAOSeek.top = RowBottom - 340
335         objDAOSeek.left = ColLeft - 70

340         objDAOSeek.ColParameter 0, ColVisible, -1
345         objDAOSeek.ColParameter 1, ColVisible, -1
350         objDAOSeek.ColParameter 2, ColVisible, -1
355         objDAOSeek.ColParameter 3, ColVisible, -1
360         objDAOSeek.ColParameter 2, ColCaption, lbl1(10)
365         objDAOSeek.ColParameter 3, ColCaption, lbl1(11)

370         Select Case Index

                Case 1                                                          'Umsatzsteuer
                
375                 objSQLAuswDef.caption = lbl1(14)
380                 objSQLAuswDef.BorderStyle = 4
385                 objSQLAuswDef.ColumnHeaders = False

390                 objSQLAuswDef.top = RowBottom '+ 2450
395                 objSQLAuswDef.left = ColLeft

400                 objSQLAuswDef.SectionBezeichnung = "cmdAuswahl" & CStr(Index)

410                 objSQLAuswDef.ColParameter 0, ColVisible, False             'HW 09.07.2012 Ver.: 6.1.114  bei der Vorschau/Drucken wird beim Speichern ein Fehler geworfen !!!! Es wird 0,1 oder 2 erwartet als Integer aber ein String wurde übergeben! FALSCH!

415                 objSQLAuswDef.ColParameter 1, colWidth, 3000
                    
420                 If Trim$(txt1(39).text) <> "" Then
                    
425                     objSQLAuswDef.Find = "Knz = '" & CStr(intSteuerTyp) & "'"
                    
                    End If
                    
435                 objSQLAuswDef.RSOpen GetACCESSConnectionString(DEF_CONNECTION), "SELECT Knz, KnzBez1 FROM Auswahl WHERE TabName = '1200_GrundKonditionen' AND FeldName = 'Ust' ORDER BY Knz" 'HW 09.07.2012 Ver.: 6.1.114  Knz wird unsichtbar mitgeladen

440                 If objSQLAuswDef.Abbruch = False Then
                        
                        Dim blnValidateUID As Boolean
                        
445                     intSteuerTyp = objSQLAuswDef.FieldText(0)               'HW 09.07.2012 Ver.: 6.1.114

450                     txt1(39).text = objSQLAuswDef.FieldText(1)
                        
455                     blnValidateUID = CheckUID(0)                            'DF 19.01.2024 , Ver.: 6.6.124 : analog zur der manuellen Änderung
                        
                        Dim oKWaehrung As KundenWaehrung                        'DF 19.06.2020 , Ver.: 6.6.102 : Nach der Umstellung der Ust.-Schl. den SteuerSatz aus Beleg-Währung holen.
                        
460                     oKWaehrung = GetWaehrung(txt1(17).text)

465                     If oKWaehrung.KndWrg <> "" Then
           
                            'Steuer-Satz (der Beleg-Währung)
470                         objPRM.FindFirstString = "name = 'txt1' AND index = 20"
475                         txt1(20).text = objPRM.EingabeUmwandlung(CStr(oKWaehrung.MwSt))
           
                        Else
            
                            'Steuer-Satz (der Beleg-Währung)
480                         objPRM.FindFirstString = "name = 'lbl2' AND index = 20"
485                         txt1(20).text = objPRM.EingabeUmwandlung("0")
            
                        End If
                        
490                     Call ResetSteuerSatz(intSteuerTyp)                      'DF 11.06.2020 , Ver.: 6.6.102 : Ust Senkung Umstellung

495                     objPRM.FindFirstString = "name = 'txt1' AND index = 39"

500                     If blnValidateUID Then

505                         Call objPRM.SprungNeu("Vorwärts", 0, txt1(39).TabIndex)

                        Else

510                         If txt1(14).Enabled Then txt1(14).SetFocus

                        End If
                        
                    Else
                    
515                     If txt1(39).Enabled Then txt1(39).SetFocus

                    End If

520             Case 5                                                          'Postfach-PLZ

530                 If gstrEWerk = "BzgEu" Then
535                     objDAOSeek.RSOpen gstrEWerk, "LkzPlz", Lkz, OrtUmwandlung(txt1(Index))
                    Else
540                     objDAOSeek.RSOpen gstrEWerk, "Plz", txt1(Index)
                    End If

545                 If objDAOSeek.Abbruch = False Then
                        'Auswahl wurde bestätigt. Inhalte ausgeben.
550                     txt1(Index) = objDAOSeek.FieldText(1)
555                     txt1(Index + 1) = objDAOSeek.FieldText(2)
560                     Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index + 1).TabIndex)
                    Else

565                     If txt1(Index).Enabled Then txt1(Index).SetFocus
                    End If

570             Case 6                                                          'Postfach-Ort

580                 If gstrEWerk = "BzgEu" Then
585                     objDAOSeek.RSOpen gstrEWerk, "LkzOrt", Lkz, OrtUmwandlung(txt1(Index))
                    Else
590                     objDAOSeek.RSOpen gstrEWerk, "Ort", txt1(Index)
                    End If

595                 If objDAOSeek.Abbruch = False Then
                        'Auswahl wurde bestätigt. Inhalte ausgeben.
600                     txt1(Index - 1) = objDAOSeek.FieldText(1)
605                     txt1(Index) = objDAOSeek.FieldText(2)
610                     Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index + 1).TabIndex)
                    Else

615                     If txt1(Index).Enabled Then txt1(Index).SetFocus
                    End If

620             Case 9                                                          'PLZ

630                 If gstrEWerk = "BzgEu" Then
635                     objDAOSeek.RSOpen gstrEWerk, "LkzPlz", Lkz, OrtUmwandlung(txt1(Index))
                    Else
640                     objDAOSeek.RSOpen gstrEWerk, "Plz", txt1(Index)
                    End If

645                 If objDAOSeek.Abbruch = False Then
                        'Auswahl wurde bestätigt. Inhalte ausgeben.
650                     txt1(Index - 1) = objDAOSeek.FieldText(0)
655                     txt1(Index) = objDAOSeek.FieldText(1)
660                     txt1(Index + 1) = objDAOSeek.FieldText(2)
665                     txt1(Index + 2) = objDAOSeek.FieldText(3)
670                     Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index + 2).TabIndex)
                    Else

675                     If txt1(Index).Enabled Then txt1(Index).SetFocus
                    End If

680             Case 10                                                         'Ort1

690                 If gstrEWerk = "BzgEu" Then
695                     If Trim(txt1(Index + 1)) = "" Then
700                         objDAOSeek.RSOpen gstrEWerk, "LkzOrt", Lkz, OrtUmwandlung(txt1(Index))
                        Else
705                         objDAOSeek.RSOpen gstrEWerk, "LkzOrtOt", Lkz, OrtUmwandlung(txt1(Index)), OrtUmwandlung(txt1(Index + 1))
                        End If

                    Else

710                     If Trim(txt1(Index + 1)) = "" Then
715                         objDAOSeek.RSOpen gstrEWerk, "Ort", OrtUmwandlung(txt1(Index))
                        Else
720                         objDAOSeek.RSOpen gstrEWerk, "OrtOt", OrtUmwandlung(txt1(Index)), OrtUmwandlung(txt1(Index + 1))
                        End If
                    End If

725                 If objDAOSeek.Abbruch = False Then
                        'Auswahl wurde bestätigt. Inhalte ausgeben.
730                     txt1(Index - 2) = objDAOSeek.FieldText(0)
735                     txt1(Index - 1) = objDAOSeek.FieldText(1)
740                     txt1(Index) = objDAOSeek.FieldText(2)
745                     txt1(Index + 1) = objDAOSeek.FieldText(3)
750                     Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index + 1).TabIndex)
                    Else

755                     If txt1(Index).Enabled Then txt1(Index).SetFocus
                    End If

760             Case 11                                                         'Ort2

765                 If Trim(txt1(Index)) = "" Then
                        'Um die Suche nicht bei leeren Eintragungen zu beginnen.
770                     OT = "A"
                    Else
775                     OT = txt1(Index)
                    End If

785                 If gstrEWerk = "BzgEu" Then
790                     If Trim(txt1(Index - 1)) = "" Then
795                         objDAOSeek.RSOpen gstrEWerk, "LkzOt", Lkz, OrtUmwandlung(OT)
                        Else
800                         objDAOSeek.RSOpen gstrEWerk, "LkzOrtOt", Lkz, OrtUmwandlung(txt1(Index - 1)), OrtUmwandlung(OT)
                        End If

                    Else

805                     If Trim(txt1(Index - 1)) = "" Then
810                         objDAOSeek.RSOpen gstrEWerk, "Ot", OrtUmwandlung(OT)
                        Else
815                         objDAOSeek.RSOpen gstrEWerk, "OrtOt", OrtUmwandlung(txt1(Index - 1)), OrtUmwandlung(OT)
                        End If
                    End If

820                 If objDAOSeek.Abbruch = False Then
                        'Auswahl wurde bestätigt. Inhalte ausgeben.
825                     txt1(Index - 3) = objDAOSeek.FieldText(0)
830                     txt1(Index - 2) = objDAOSeek.FieldText(1)
835                     txt1(Index - 1) = objDAOSeek.FieldText(2)
840                     txt1(Index) = objDAOSeek.FieldText(3)
845                     Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                    Else

850                     If txt1(Index).Enabled Then txt1(Index).SetFocus
                    End If

            End Select

        Else

855         If Index = 34 Or Index = 36 Then                                    'Konto-Kennzeichen, KostenArt

860             objSQLAuswDef.BorderStyle = 4
865             objSQLAuswDef.ColumnHeaders = False
870             objSQLAuswDef.top = RowBottom '+ 80
875             objSQLAuswDef.left = ColLeft
880             objSQLAuswDef.ColParameter 0, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
885             objSQLAuswDef.ColParameter 1, colWidth, 1100
890             objSQLAuswDef.caption = ""

895             If Trim(txt1(Index)) <> "" Then

900                 objSQLAuswDef.Find = "Knz like '" & txt1(Index).text & "*'"

                End If

910             If Index = 34 Then

915                 objSQLAuswDef.RSOpen GetACCESSConnectionString(DEF_CONNECTION), "SELECT Knz, KnzBez1 FROM [Auswahl] WHERE TabName = '1200_AllgDaten' AND FeldName = 'KtoKnz' AND (Knz = 'D' OR KnZ = 'K') ORDER BY Knz"

                Else
                
920                 objSQLAuswDef.RSOpen GetACCESSConnectionString(DEF_CONNECTION), "SELECT Knz, KnzBez1 FROM [Auswahl] WHERE TabName = '5800_SpeditionsBuch' AND FeldName = 'KostenArt' ORDER BY Knz"

                End If

925             If objSQLAuswDef.Abbruch = False Then

930                 txt1(Index) = objSQLAuswDef.FieldText(0)

935                 txt1_Validate Index, bCancel

940                 If bCancel = True Then

                        Exit Sub

                    End If
                    
945                 Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)

                Else

950                 If txt1(Index).Enabled Then txt1(Index).SetFocus

                End If

            Else

955             objSQLAusw.FilterBar = True
960             objSQLAusw.BorderStyle = 4
965             objSQLAusw.caption = lbl1(Index)
970             objSQLAusw.top = RowBottom                                      '+ 340    'DH, 14.02.2018, 6.5.105, Bei der SQLAuswahl duerfen die 340 nicht abgezeogen werden (siehe Zuweisung zu RowBottom oben)
975             objSQLAusw.left = ColLeft

980             Select Case Index

                    Case 0, 12                                                      'MCode/Konto-Nr IL 04.12.2024 6.7.102
                    
                        Dim SortSpalte As String
                        
                        Dim HauptTxt   As Control

985                     objSQLAusw.ColParameter 0, colWidth, 1200
990                     objSQLAusw.ColParameter 0, ColCaption, "M-Code"
995                     objSQLAusw.ColParameter 1, colWidth, 300
1000                    objSQLAusw.ColParameter 1, ColCaption, "Art"
1005                    objSQLAusw.ColParameter 2, colWidth, 900
1010                    objSQLAusw.ColParameter 2, ColCaption, "KtoNr."
1015                    objSQLAusw.ColParameter 3, colWidth, 1900
1020                    objSQLAusw.ColParameter 4, colWidth, 400
1025                    objSQLAusw.ColParameter 5, colWidth, 900
1030                    objSQLAusw.ColParameter 6, colWidth, 1900
1035                    objSQLAusw.ColParameter 7, colWidth, 1900
1040                    objSQLAusw.ColParameter 8, ColVisible, 0
1045                    objSQLAusw.ColParameter 9, ColVisible, 0
1050                    objSQLAusw.ColParameter 10, ColVisible, 0

1055                    objSQLAusw.SperrenFeld = "Sperre"

                        '<Added by: IL at: 04.12.2024, Ver.: 6.7.102 >
                        '# Wählen die Spalte aus, nach der die Tabelle sortiert werden soll
1060                    Select Case Index

                            Case 0
    
1065                            SortSpalte = "MCode"
    
1070                        Case 12
    
1075                            SortSpalte = "KtoNr"

                        End Select

                        '</Added by: IL at: 04.12.2024, Ver.: 6.7.102 >

                        'Datensatz positionieren.
1080                    If Trim(txt1(Index).text) <> "" Then

1085                        objSQLAusw.Find = SortSpalte & " LIKE '" & txt1(Index).text & "'"

                        End If
                        
1095                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT MCode,KtoKnz,KtoNr,Name1 + ' ' + Name2 AS Name ,Lkz,Plz,Ort,Straße,Sperre,Name1,Name2 FROM [1200_AllgDaten] WHERE (KundenKnz = 0 OR KundenKnz = 99) AND (KtoKnz = 'D' OR KtoKnz = 'K') AND Sperre <> '2' ORDER BY " & SortSpalte

1100                    If objSQLAusw.Abbruch = False Then
                        
1105                        If objSQLAusw.FieldText(8) = "0" Then

1110                            txt1(0) = objSQLAusw.FieldText("MCode")

1115                            GstrAuftraggeber = objSQLAusw.FieldText("Name1")

1120                            If KundeZeigen Then

                                    'HW 17.09.2012 Ver.: 6.1.117 Hier abfragen auf Kreditor / Debitor
                                    '####################################################
1125                                If programmNr = "283" Then

1130                                    If objSQLAusw.FieldText(1) = "K" Then

                                            'Meldung
1135                                        MsgBox GetMessage(295, "ZusatzTexte_53100"), vbOKOnly + vbInformation, strMeldungCap

                                        End If

                                    Else

1140                                    If programmNr = "284" Then

1145                                        If objSQLAusw.FieldText(1) = "D" Then

                                                'Meldung
1150                                            MsgBox GetMessage(296, "ZusatzTexte_53100"), vbOKOnly + vbInformation, strMeldungCap

                                            End If
                                            
                                        End If
                                        
                                    End If

                                    '####################################################

1155                                blnBelegNeu = True

1160                                Call objPRM.SprungNeu("Vorwärts", 1, txt1(Index).TabIndex)

                                End If

                            Else
                                
1165                            Call msgText(1, 1541, 0, 0, 0)                  'DF 03.02.12 Meldung für gesperrte Kunden

1170                            strMessage = GsMsgText(0)
1175                            strMessage = Replace(strMessage, "%1", objSQLAusw.FieldText("MCode"))
1180                            strMessage = Replace(strMessage, "%2", vbCrLf)
1185                            intAntwort = MsgBox(strMessage, vbYesNo + vbInformation + vbDefaultButton2, strMeldungCap)    'Gesperrter Stammsatz.... Trotzdem Übernehmen?

1190                            If intAntwort = 7 Then                          ' Falls Nein

1195                                txt1(Index).SetFocus
1200                                txt1(Index).BackColor = &HC0E0FF
1205                                blnBelegNeu = False

                                    Exit Sub

                                End If

1210                            txt1(0) = objSQLAusw.FieldText("MCode")

1215                            GstrAuftraggeber = objSQLAusw.FieldText("Name1")

1220                            If KundeZeigen Then

1225                                blnBelegNeu = True

1230                                Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)

                                End If

                            End If

                        Else

1235                        If txt1(Index).Enabled Then txt1(Index).SetFocus
1240                        blnBelegNeu = False

                        End If

1245                Case 3                                                      'Ansprechpartner

1255                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Anrede,Name1,Name2 FROM [1200_AnsprPartner] WHERE MCode = '" & txt1(0) & "' ORDER BY Standard"

1260                    If objSQLAusw.Abbruch = False Then

1265                        txt1(Index) = objSQLAusw.FieldText(0) & " " & objSQLAusw.FieldText(1) & " " & objSQLAusw.FieldText(2)

1270                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)

                        Else
                        
1275                        If txt1(Index).Enabled Then txt1(Index).SetFocus

                        End If

1280                Case 8                                                      'Lkz

1285                    If Trim(txt1(Index).text) <> "" Then

1290                        objSQLAusw.Find = "Lkz LIKE '" & txt1(Index).text & "%'"

                        End If

1300                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Lkz, Land FROM [1100_Land] ORDER BY Lkz"

1305                    If Not objSQLAusw.Abbruch Then

1310                        txt1(Index).text = objSQLAusw.FieldText(0)

1315                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex, True)

                        End If

1320                Case 17                                                     'Währung

                        '<Removed by: DFiebach at: 10.06.2020, Ver.: 6.6.102 >
                        '
                        ' # OPTIMIERUNG -> Standardfenster
                        '
                        '1155                    objSQLAusw.ColParameter 0, colWidth, (txt1(Index).width / cResize.CurrScaleFactorWidth)
                        '1160                    objSQLAusw.ColParameter 1, ColNumberFormat, "#0.00"
                        '1165                    objSQLAusw.ColParameter 2, ColNumberFormat, "###0.#####0"
                        '
                        '                        'Datensatz positionieren.
                        '1170                    If Trim(txt1(Index)) <> "" Then
                        '
                        '1175                        objSQLAusw.Find = "ISO like '" & txt1(Index) & "*'"
                        '
                        '                        End If
                        '
                        '1180                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT ISO,MwSt,Kurs,Schl,Land FROM [1100_Währungen]"
                        '
                        '1185                    If oSQLAusw.Abbruch = False Then
                        '
                        '1190                        txt1(Index) = oSQLAusw.FieldText(0)
                        '
                        '1195                        txt1(20) = Format(oSQLAusw.FieldText(1), "#0.00")    'Steuersatz
                        '
                        '1200                        KursFuerRechnung "0", txt1(17), txt1(18)    ' "0" ist ein Dummy-Wert
                        '
                        '1205                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                        '
                        '                        Else
                        '
                        '1210                        txt1(Index).SetFocus
                        '
                        '                        End If
                        '
                        '</Removed by: DFiebach at: 10.06.2020, Ver.: 6.6.102 >
                        
                        'DF 11.06.2020 , Ver.: 6.6.102 : UST Senkung Umstellung
                        
                        Dim oF2Waehrung As ResultF2_Waehrung
                        
1325                    oF2Waehrung = GetF2_Waehrung("Schl", txt1(Index), Index, Me, cReSize, objPRM)
                        
1330                    If oF2Waehrung.Canceled = False Then
                            
                            'Währung Schlüssel
1335                        txt1(Index) = oF2Waehrung.Schl
1340                        lbl2(17).caption = oF2Waehrung.ISO
                            
                            'Kurs
1345                        objPRM.FindFirstString = "name = 'txt1' AND index = 19"
1350                        txt1(19) = objPRM.EingabeUmwandlung(CStr(oF2Waehrung.Kurs))
            
                            'Steuer-Satz (der Beleg-Währung)
1355                        objPRM.FindFirstString = "name = 'lbl2' AND index = 20"
1360                        txt1(20).text = objPRM.EingabeUmwandlung(CStr(oF2Waehrung.MwSt))
                            
1365                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)

                        Else
                        
1370                        If txt1(Index).Enabled Then txt1(Index).SetFocus

                        End If

1375                Case 25, 26                                                 'FiBu-Schlüßeln

1380                    objSQLAusw.ColParameter 0, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)

                        'Datensatz positionieren.
1385                    If Trim(txt1(Index)) <> "" Then
1390                        objSQLAusw.Find = "Schl like '" & txt1(Index) & "'"
                        End If

1400                    If Index = 25 Then
1405                        objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl,Bezeichnung,Erlöse,Kosten FROM [1100_FiBuKostenStellen]"
                        Else
1410                        objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT Schl,Bezeichnung,Erlöse,Kosten FROM [1100_FiBuSachkonten]"
                        End If

1415                    If objSQLAusw.Abbruch = False Then
1420                        txt1(Index) = objSQLAusw.FieldText(0)
1425                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                        Else

1430                        If txt1(Index).Enabled Then txt1(Index).SetFocus
                        End If

1435                Case 33                                                     'Kfz

1445                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT KfzSchl,Art,Knz,Nutzlast FROM [1600_Fahrzeuge]"

1450                    If objSQLAusw.Abbruch = False Then
1455                        txt1(Index) = objSQLAusw.FieldText(0)
1460                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                        Else

1465                        If txt1(Index).Enabled Then txt1(Index).SetFocus
                        End If

1470                Case 29                                                     'Abf.-Datum

1475                    objSQLAusw.ColParameter 0, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1480                    objSQLAusw.ColParameter 1, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1485                    objSQLAusw.ColParameter 2, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1490                    objSQLAusw.FilterBar = True

1495                    sql = "SELECT AbfDatum,AbfZus,AbfPos FROM [2800_Abfr_SdgAusw] GROUP BY AbfPos,AbfZus,AbfDatum "
1500                    sql = sql & " ORDER BY AbfDatum,AbfZus,AbfPos"

1505                    If IsDate(txt1(Index)) Then objSQLAusw.Find = "AbfDatum >= CONVERT(Datetime,'" & txt1(Index).text & "')"

1515                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), sql

1520                    If objSQLAusw.Abbruch = False Then

1525                        txt1(Index) = objSQLAusw.FieldText(0)
1530                        txt1(Index + 1) = objSQLAusw.FieldText(1)
1535                        txt1(Index + 2) = objSQLAusw.FieldText(2)
1540                        txt1(35) = ""

1545                        If Trim(txt1(36)) = "" Then txt1(36) = "H"

1550                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)

                        Else
                        
1555                        If txt1(Index).Enabled Then txt1(Index).SetFocus

                        End If

1560                    Call DatumUnterschiedRek(0, txt1(Index), txt1(37), 90, "d") 'GW_06.03.2018 90-Tage überprüfung

1565                Case 30                                                     'Abf.-Zusatz

1570                    objSQLAusw.ColParameter 0, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1575                    objSQLAusw.ColParameter 1, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1580                    objSQLAusw.ColParameter 2, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1585                    objSQLAusw.FilterBar = True

1590                    sql = "SELECT AbfZus,AbfPos,AbfDatum FROM [2800_Abfr_SdgAusw] GROUP BY AbfZus,AbfPos,AbfDatum "
1595                    sql = sql & " ORDER BY AbfZus,AbfPos,AbfDatum"

1600                    If IsDate(txt1(Index)) Then objSQLAusw.Find = "AbfZus like '" & txt1(Index) & "%'"
                       
1610                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), sql

1615                    If objSQLAusw.Abbruch = False Then
1620                        txt1(Index) = objSQLAusw.FieldText(0)
1625                        txt1(Index + 1) = objSQLAusw.FieldText(1)
1630                        txt1(Index - 1) = objSQLAusw.FieldText(2)
1635                        txt1(35) = ""

1640                        If Trim(txt1(36)) = "" Then txt1(36) = "H"
1645                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                        Else

1650                        If txt1(Index).Enabled Then txt1(Index).SetFocus
                        End If

1655                Case 31                                                     'Abf.-Position

1660                    objSQLAusw.ColParameter 0, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1665                    objSQLAusw.ColParameter 1, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1670                    objSQLAusw.ColParameter 2, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1675                    objSQLAusw.FilterBar = True

1680                    sql = "SELECT AbfPos,AbfZus,AbfDatum FROM [2800_Abfr_SdgAusw] GROUP BY AbfPos,AbfZus,AbfDatum "
1685                    sql = sql & " ORDER BY AbfPos,AbfZus,AbfDatum"

1690                    If IsDate(txt1(Index)) Then objSQLAusw.Find = "AbfPos >= " & txt1(Index)

1700                    objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), sql

1705                    If objSQLAusw.Abbruch = False Then
1710                        txt1(Index) = objSQLAusw.FieldText(0)
1715                        txt1(Index - 1) = objSQLAusw.FieldText(1)
1720                        txt1(Index - 2) = objSQLAusw.FieldText(2)
1725                        txt1(35) = ""

1730                        If Trim(txt1(36)) = "" Then txt1(36) = "H"
1735                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                        Else

1740                        If txt1(Index).Enabled Then txt1(Index).SetFocus
                        End If

1745                Case 35                                                     'SendungsNr

1750                    If IsNumeric(txt1(31)) And IsDate(txt1(29)) Then

1755                        objSQLAusw.MaxWidth = 12000
1760                        objSQLAusw.ColParameter 0, colWidth, (txt1(Index).width / cReSize.CurrScaleFactorWidth)
1765                        objSQLAusw.ColParameter 2, colWidth, 1500
1770                        objSQLAusw.ColParameter 3, colWidth, 600
1775                        objSQLAusw.ColParameter 4, colWidth, 1500
1780                        objSQLAusw.ColParameter 5, colWidth, 1500
1785                        objSQLAusw.ColParameter 6, colWidth, 600
1790                        objSQLAusw.ColParameter 7, colWidth, 1500

1795                        sql = "SELECT ErfNr,Datum,AName,APlz,AOrt,EName,EPlz,EOrt FROM [2800_Abfr_SdgAusw] "
1800                        sql = sql & " WHERE AbfPos = " & txt1(31)
1805                        sql = sql & " AND AbfZus = '" & txt1(30) & "'"
                            '3740            sql = sql & " AND AbfDatum = DateValue('" & txt1(29).Text & "')"
1810                        sql = sql & " AND AbfDatum = '" & txt1(29).text & "'"    'DH, 14.12.2017, 6.5.103, DateValue() gibt es im SQL Server nicht.
1815                        sql = sql & " UNION SELECT '<alle>' AS ErfNr, Null AS Datum,'' AS AName,'' AS APlz,'' AS AOrt,'' AS EName,'' AS EPlz,'' AS EOrt FROM [2800_Abfr_SdgAusw]"

1825                        objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), sql

1830                        If objSQLAusw.Abbruch = False Then
1835                            txt1(Index) = objSQLAusw.FieldText(0)
1840                            Call objPRM.SprungNeu("Vorwärts", 0, txt1(Index).TabIndex)
                            Else

1845                            If txt1(Index).Enabled Then txt1(Index).SetFocus
                            End If
                            
                        End If

1850                Case 40

1855                    Call SteuerText(Index)

1860                Case 42                                                    'Dezimal-Stellen

1865                    objSQLAuswDef.caption = lbl1(4)
1870                    objSQLAuswDef.BorderStyle = 4
1875                    objSQLAuswDef.ColumnHeaders = False

1880                    objSQLAuswDef.top = RowBottom '+ 2450
1885                    objSQLAuswDef.left = ColLeft

1890                    objSQLAuswDef.SectionBezeichnung = "cmdAuswahl" & CStr(Index)

                        'HW 09.07.2012 Ver.: 6.1.114  bei der Vorschau/Drucken wird beim Speichern ein Fehler geworfen !!!! Es wird 0,1 oder 2 erwartet als Integer aber ein String wurde übergeben! FALSCH!
1900                    objSQLAuswDef.ColParameter 0, ColVisible, False
1905                    objSQLAuswDef.ColParameter 1, colWidth, 3000

                        'HW 09.07.2012 Ver.: 6.1.114  Knz wird unsichtbar mitgeladen
1915                    objSQLAuswDef.RSOpen GetACCESSConnectionString(DEF_CONNECTION), "SELECT Knz, KnzBez1 FROM Auswahl WHERE TabName = '1200_GrundKonditionen' AND FeldName = 'Ust' ORDER BY Knz"

1920                    If objSQLAuswDef.Abbruch = False Then

1925                        intSteuerTyp = objSQLAuswDef.FieldText(0)           'HW 09.07.2012 Ver.: 6.1.114

1930                        txt1(39).text = objSQLAuswDef.FieldText(1)

1935                        Call objPRM.SprungNeu("Vorwärts", 0, txt1(39).TabIndex)

                        Else
                        
1940                        If txt1(39).Enabled Then txt1(39).SetFocus

                        End If

                End Select

            End If
            
        End If

        Exit Sub

Fehler:
1945    Call FehlerErklärung("frmSP52830", "cmdAuswahl_Click()")
End Sub

Sub SteuerText(Index As Integer)                                                'HW 09.07.2012 Ver.: 6.1.114   eingebaut!

        Dim nr          As String

        Dim Sort        As String

        Dim rec1100Text As ADODB.Recordset

        Dim textX       As SpActiveX1.clsTexte
  
        On Error GoTo Fehler

100     Set textX = New SpActiveX1.clsTexte
105     textX.titel = txt1(Index).text
110     textX.SetResizeParameter cReSize.CurrScaleFactorHeight, cReSize.CurrScaleFactorWidth, Me.Font.Size
115     textX.TexteInit (5)

120     If textX.titel <> "" Then
125         txt1(Index).text = textX.titel
130         nr = textX.nr
135         Sort = textX.Sort
    
            On Error Resume Next

140         txt1(Index).SetFocus 'HW 02.08.2011 Ver.: 6.1.107
145         Err.Clear

            On Error GoTo Fehler

        Else

            Exit Sub

        End If

150     Set textX = Nothing

155     If txt1(Index).text <> "" Then
160         Set rec1100Text = New ADODB.Recordset
165         rec1100Text.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Titel = '" & txt1(Index).text & "' And Lkz = '" & nr & "' AND Sort = " & Sort, gConn, adOpenStatic, adLockReadOnly

170         If rec1100Text.RecordCount > 0 Then
175             If rec1100Text!Sort = 1 Then                      'Wenn automatisch, wird Variable auf Leerstring gesetzt, weil im Formular darauf reagiert wird!
180                 gstrSteuerText = ""
                Else

185                 If rec1100Text!text = "" Then                 'Wenn der Text Leer ist muss ein Leerzeichen gesetzt werden!
190                     gstrSteuerText = " "
                    Else
195                     gstrSteuerText = rec1100Text!text         'Normaler Text wird gespeichert
                    End If
                End If
            End If

200         rec1100Text.Close

205         Set rec1100Text = Nothing
        Else
210         gstrSteuerText = ""
        End If

        Exit Sub

Fehler:
215     Call FehlerErklärung("frmSP56430", "SteuerText(" & Index & ")")
End Sub

Private Sub Form_Activate()
    
        On Error GoTo Fehler

100     frmMsg.LoadSkinner

105     If gi_UpdateInfoAngezeigt = False Then

            '30       DoEvents                                                  'Damit Form im Hintergrung richtig dargestellt werden kann

110         gi_UpdateInfoAngezeigt = True

115         gi_UpdateAenderung = 15                                            'DF 20.11.2024 , Ver.: 6.7.101 : 15
            
120         objHlp.ModulName = App.EXEName & "_283"    'DH, 29.10.2015, 6.4.110, Es soll jetzt immer nur auf den UpdateInfo Text der Rechnung zugegriffen werden

125         gobjUpdateAenderungCount = GetSetting("SP50000", App.EXEName & "_283", "UpdateAenderungCounter", 0)
130         gobjUpdateAenderung = GetSetting("SP50000", App.EXEName & "_283", "UpdateAenderungAnzeigen", 0)

            'Wenn UpdateMeldung unterdrückt wird, muss auf höhere Änderung abgefragt werden!
135         If gobjUpdateAenderung = 1 Then

                'Auf count abfragen
140             If gobjUpdateAenderungCount < gi_UpdateAenderung Then

145                 objHlp.UpdateAnzeigen = True
150                 objHlp.UpdateCounter = gi_UpdateAenderung
155                 objHlp.HlpShow HlpRead, "UpdateAenderung" & "_283"
160                 objHlp.UpdateAnzeigen = False

                End If

            Else
            
165             objHlp.UpdateAnzeigen = True
170             objHlp.UpdateCounter = gi_UpdateAenderung
175             objHlp.HlpShow HlpRead, "UpdateAenderung" & "_283"
180             objHlp.UpdateAnzeigen = False

            End If
            
        End If

        Exit Sub

Fehler:
185     Call FehlerErklärung("frmSP52830", "Form_Activate")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
        On Error GoTo Fehler

        Dim AltDown

        Const vbAltMask = 4
  
        'Beim clicken auf das Programm-Symbol im Register Aktive Programme (SP51.ProgrammAktivieren)
        'wird SendKeys "%{F12}", True - Befehl ausgeführt.
        'Aktuelle form wird zu aktiven Form in normaler Größe.
100     AltDown = (Shift And vbAltMask) > 0

        '<Removed by: DFiebach at: 22.11.2024, Ver.: 6.7.101 >
        ' # was soll hier diese Funktionalität???
        '105     If KeyCode = vbKeyF12 Then
        '110         If AltDown Then
        '115             Call Main
        '            End If
        '        End If
        '</Removed by: DFiebach at: 22.11.2024, Ver.: 6.7.101 >
  
120     If ((Shift And vbShiftMask) > 0) And Not AltDown Then
125         shiftPressed = True
        Else
130         shiftPressed = False
        End If
  
135     If Shift = 2 Then
140         GStrg = True
        End If

        Exit Sub

Fehler:
145     Call FehlerErklärung("frmSP52830", "Form_KeyDown")

End Sub

Private Sub Form_Load()

        On Error GoTo Fehler

        Dim i As Integer

        Dim x As Control
        
        '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
        '# Um neue Module Angebot und AufBest. erweitert.
100     Select Case GintBelegArt
        
            Case 0
            
105             SaveSetting "SP50000", App.EXEName, "SP62830_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
            
110         Case 1
            
115             SaveSetting "SP50000", App.EXEName, "SP62840_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
            
120         Case 2, 3

125             Call SteuerungFiBuUndSpeditionsbuch(False)

130             Select Case GintBelegArt

                    Case 2

135                     SaveSetting "SP50000", App.EXEName, "SP62880_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

140                 Case 3

145                     SaveSetting "SP50000", App.EXEName, "SP62890_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

                End Select
 
        End Select

        '100     If GintBelegArt = 0 Then
        '105         SaveSetting "SP50000", App.EXEName, "SP62830_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
        '        Else
        '110         SaveSetting "SP50000", App.EXEName, "SP62840_WndHwnd", Me.hwnd      'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
        '        End If
        '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
  
150     Set objPRM = New clsPRM
155     Set objPRM.gForm = Me
160     objPRM.PRM_Alle

        '<Added by: IL at: 29.10.2024, Ver.: 6.7.101 >
165     lblDummy(0).caption = objPRM.getColCaption("name = 'stringProgNameRechnung'")
170     lblDummy(1).caption = objPRM.getColCaption("name = 'stringProgNameGutschrift'")
175     lblDummy(2).caption = objPRM.getColCaption("name = 'stringProgNameAngebot'")
180     lblDummy(3).caption = objPRM.getColCaption("name = 'stringProgNameAuftragsbestetigung'")
        '</Added by: IL at: 29.10.2024, Ver.: 6.7.101 >

185     strMeldungCap = mnuDummy(0).caption
       
190     SetFormCaption (False)

195     gintPrivBelegArt = GintBelegArt
200     Frame1(2).caption = lblDummy(GintBelegArt)
205     cmd1(5).caption = objPRM.getColCaption("name = 'singleStringWeiter'")
210     mnuBearb1(0).caption = objPRM.getColCaption("name = 'singleStringWeiter'")                        'DF 19.01.2015

215     Me.height = FORM_HEIGHT
  
220     SetXPSize Me

225     Call setSkinnerBackColor(sta1)
  
230     sta1.Panels(1).text = "SP62830"
235     sta1.Panels(2).text = DisplayVerInfo(GsHauptPfadLokal & "exe\" & Gc_strExeFile)

240     Set objSQLAusw = New SPSQLAuswahl.clsSQLAuswahl                         'HW 06.05.2015
245     objSQLAusw.FilterBar = True

250     Set objSQLAuswDef = New SPSQLAuswahl.clsSQLAuswahl
  
255     Set objDAOSeek = New SPDAOSeek.clsDAOSeek
260     objDAOSeek.DatabaseName = GsHauptPfadLokal & "fix\SP50000.fix"
  
265     Set objHlp = New SpHlp.clsHlp
270     objHlp.DatabaseName = GsHauptPfadLokal & "hlp\SP50000.hlp"
275     objHlp.table = Me.name
280     objHlp.caption = Me.name & " - Feldhilfe"
285     objHlp.parentFrm = Me                                                   '08.12.2015, 6.4.114, Dem HLP-Objekt dieses Formular uebergeben, damit sich das Hilfefenster daran ausrichten kann
    
290     Set objPlausi = New clsPlausi
295     objPlausi.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
300     Set objLimit = New cLimit

305     Set objDruckOptionen = New clsDruckOptionen
310     objDruckOptionen.EnableValutaDatum = True                                               'DH, 16.02.2015, 6.4.103
315     objDruckOptionen.EnableBelegNrSystemparameter = Not GetSysPar_FakturaBelegNrSperren     'DF 05.02.2019 , Ver.: 6.5.109 : SystemParameter -> Faktura BelegNr sperren
        
320     Set objERechnung = New clsERechnung                                     'DF 23.08.2024 , Ver.: 6.7.101 : Evtl. erst instanzieren wenn der Kunde gezogen wird.
        
325     Set connSQL = New ADODB.Connection
330     connSQL.ConnectionString = GetConnectionString(GsHauptPfadLokal, Spedifix, GsAnwenderNr)

335     Call MaskeLeeren(False)

340     Call EWerkErmitteln

345     If gintPrivBelegArt = 0 Then
350         SaveSetting "SP50000", "SP52800", "SP52830", Me.caption
        Else
355         SaveSetting "SP50000", "SP52800", "SP52840", Me.caption
        End If

360     mnuOpt1(50).Checked = GetSetting("SP50000", "SP52800", "SP52830DruckerDialog", "-1")
365     mnuOpt1(1).Checked = GetSetting("SP50000", "SP52800", "SP52830Ansprechpartner", "-1")
370     mnuOpt1(2).Checked = GetSetting("SP50000", "SP52800", "SP52830FolgeseitenKurzDrucken", "0")
375     blnFolgeseitenKurzDrucken = mnuOpt1(2).Checked
380     mnuOpt1(3).Checked = GetSetting("SP50000", "SP52800", "SP52830Bearbeiter", "-1") 'HW 18.11.2015 eingepflegt
385     BearbeiterDrucken = mnuOpt1(3).Checked                                  'HW 29.04.2016

390     mnuOpt1(4).Checked = GetSetting("SP50000", "SP52800", "SP52830GesamtIstBrutto", "0") 'HW 24.05.2016 Ver.: 6.4.120 GesamtIstBrutto eingepflegt
395     GesamtIstBrutto = mnuOpt1(4).Checked                                    'HW 24.05.2016 Ver.: 6.4.120

        '<Added by: GW at: 02.04.2019, Ver.: 6.5.110 >
        
        '<Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
        '# Jetzt schließt sich das zweite Fenster nach dem Drucken immer automatisch
400     mnuOpt1(5).Visible = False
405     mnuOpt1(6).Visible = False
        '</Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
        
        '<Removed by: IL at: 07.11.2024, Ver.: 6.7.101 >

        '400     mnuOpt1(5).Checked = GetSetting("SP50000", "SP52800", "SP52830MaskeLeeren", "0")
        '405     blnMaskeLeeren = mnuOpt1(5).Checked
        '410     mnuOpt1(6).Checked = GetSetting("SP50000", "SP52800", "SP52830MaskeSchliessen", "0")
        '415     blnMaskeSchliessen = mnuOpt1(6).Checked

        '</Removed by: IL at: 07.11.2024, Ver.: 6.7.101 >
        '</Added by: GW at: 02.04.2019, Ver.: 6.5.110 >
  
410     txt1(42).text = CStr(postCommaPreis)                                    'HW, 10.01.2018
  
415     txt1(15).Visible = False
420     txt1(16).Visible = False
425     lbl1(15).Visible = False
430     lbl1(16).Visible = False
435     txt1(15).text = ""                                                      'DH, 13.03.2013, Ehemalige Textfelder fuer BelegNr und -Datum. Werden jetzt nicht mehr genutzt
440     txt1(16).text = ""                                                      '                da dieses im DruckOptionen Dialog einstellbar ist
    
445     Call AutomatischSteuerText(40)                                          'HW 16.07.2012
        
        '########## Subclassing: Messages festlegen #############
        ' DeW, ZyG Mai 2011
450     AttachMessage Me, Me.hwnd, WM_GETMINMAXINFO                             'DeW
455     AttachMessage Me, Me.hwnd, WM_EXITSIZEMOVE                              'DeW
        '
        '########################################################
  
        '####### Subclassing: Groessenbegrenzung Formular #######
        ' TODO in Arbeit, MagicNumbers
        'dieser Aufruf kann je nach Programm-Modul woanders in der
        'Form_Load Methode stehen!
        'Zuerst muss der "alte" Code die Zuweisung von Breit und
        'Hoehe korrekt vorgenommen haben!
460     SetMinMaxInfo Me.hwnd, Me.height, (Me.height * 2), Me.width, (Me.width * 2)
        '
        '########################################################
  
        '###### Formular Resizing: Parameter setzen#############
        ' DeW, ZyG, Mai 2011
        'Section- oder KeyBezeichnung sind in vielen Faellen in
        'altem Code hart eincodiert worden, manchmal wird auch
        'eine Variable verwendet...
465     Set cReSize = New FormResize
470     cReSize.setSectionBezeichnung = "SP52830"
475     cReSize.setKeyBezeichnung = "SP52830"
480     cReSize.setIstUnterFenster = False
        '
        '########################################################
  
        '######## Formular Resizing: Formular zuweisen ##########
        ' DeW, Mai 2011
        'Zuweisung von Form erst nach Groessensetzung s.o. Me.Width = ...
        'aber auf jeden Fall nach SetMinMaxInfo ... fuer die
        'Groessenbegrenzung
485     cReSize.Form = Me
        '
        'Speichere keine Informationen (Spaltenbreiten usw.) fuer die
        'Tabellen im Form, wenn z.B. nur eine einzelne Tabelle
        'vorhanden ist, die jeweils mit neuen Daten gefuellt
        'und an eine andere Position verschoben wird (z.B. SP51000
        'Mandantenstamm
490     cReSize.IgnoreTrueDBGridInfo = True
        '
        '########################################################
495     cReSize.resize
  
500     Call readWindowPos(Me, "SP52800", "SP52830" & gintPrivBelegArt & "Left", "SP52830" & gintPrivBelegArt & "Top")  'HW 10.07.2014

505     If GsTitel <> "" Then
510         GlSP51000hwnd = FindWindow(vbNullString, GsTitel)
515         SetWindowLong Me.hwnd, GWL_HWNDPARENT, GlSP51000hwnd
        End If

        Exit Sub

Fehler:
520     Call FehlerErklärung("frmSP52830", "Form_Load()")
End Sub

Public Sub EWerkErmitteln()

        On Error GoTo Fehler

        Dim rs As ADODB.Recordset

100     Set rs = New ADODB.Recordset
105     rs.Open "SELECT Ewerk FROM [1200_GrundKonditionen] WHERE MCode = '" & txt1(0).text & "'", gConn, adOpenStatic, adLockReadOnly
  
110     If rs.RecordCount = 0 Then
115         If rs.state = adStateOpen Then rs.Close
120         rs.Open "SELECT Ewerk FROM [1200_GrundKonditionen] WHERE MCode = '@System'", gConn, adOpenStatic, adLockReadOnly
        End If
  
125     cmdAuswahl(5).Visible = True
130     cmdAuswahl(6).Visible = True
135     cmdAuswahl(9).Visible = True
140     cmdAuswahl(10).Visible = True
145     cmdAuswahl(11).Visible = True
  
150     Select Case rs!EWerk

            Case "0"
155             gstrEWerk = "" '0=ohne
160             cmdAuswahl(5).Visible = False
165             cmdAuswahl(6).Visible = False
170             cmdAuswahl(9).Visible = False
175             cmdAuswahl(10).Visible = False
180             cmdAuswahl(11).Visible = False

185         Case "1"
190             gstrEWerk = "GftOrt"

195         Case "2"
200             gstrEWerk = "BzgD"

205         Case "3"
210             gstrEWerk = "BzgEu"
        End Select

        Exit Sub

Fehler:
215     Call FehlerErklärung("frmSP52830", "EWerkErmitteln")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

        '<Added by: IL at: 08.11.2024, Ver.: 6.7.101 >
100     If gboolHatKind Then

            Dim lngNewBelegId As Long
                        
105         lngNewBelegId = 0

110         If gboolBelegAngenommen Then lngNewBelegId = glngBelegID

115         Select Case gintPrivBelegArt

                Case 0

120                 frmRechnungErf.TabellenAktualisieren lngNewBelegId, False

125                 Unload frmRechnungErf

130                 If gboolHatKind Then

135                     Cancel = True

                        Exit Sub

                    End If
    
140             Case 1

145                 frmGutschriftErf.TabellenAktualisieren lngNewBelegId, False

150                 Unload frmGutschriftErf

155                 If gboolHatKind Then

160                     Cancel = True

                        Exit Sub

                    End If
    
165             Case 2

170                 frmAngebotErf.TabellenAktualisieren lngNewBelegId, False

175                 Unload frmAngebotErf

180                 If gboolHatKind Then

185                     Cancel = True

                        Exit Sub

                    End If
    
190             Case 3

195                 frmAuftragsbestErf.TabellenAktualisieren lngNewBelegId, False

200                 Unload frmAuftragsbestErf

205                 If gboolHatKind Then

210                     Cancel = True

                        Exit Sub

                    End If

            End Select

        End If

        '</Added by: IL at: 08.11.2024, Ver.: 6.7.101 >

215     Me.Visible = False 'HW 22.09.2011
                
        '####### Subclassing: Messages austragen #############
        'DeW, Mai 2011
220     DetachMessage Me, Me.hwnd, WM_EXITSIZEMOVE
225     DetachMessage Me, Me.hwnd, WM_GETMINMAXINFO
        '
        '#####################################################

        '####### Subclassing: Groessenbegrenzung loeschen #######
        'DeW, Mai 2011
230     RemoveMinMaxInfo Me.hwnd
        '
        '########################################################
  
        'HW 10.07.2014
235     Call writeWindowPos(Me, "SP52800", "SP52830" & gintPrivBelegArt & "Left", "SP52830" & gintPrivBelegArt & "Top")

240     Me.Visible = False 'HW 22.09.2011

245     If Me.Visible Then Me.Hide
250     DoEvents
255     Sleep (0.5)

End Sub

Private Sub Form_Resize()

    On Error GoTo Fehler

    sta1.Panels(1).width = Me.width / FORM_WIDTH * FORM_PANELS_1
    sta1.Panels(2).width = Me.width / FORM_WIDTH * FORM_PANELS_2

    Exit Sub

Fehler:
    Call FehlerErklärung("frmSP52830", "Form_Resize")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

        On Error GoTo Fehler

100     Call CloseBelegArchiv                                                   'Added by: GW at: 09.03.2020, Ver.: GOBD

        '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
        ' # Um neue Module Angebot und AufBest. erweitert.
105     Select Case GintBelegArt

            Case 0
    
110             SaveSetting "SP50000", App.EXEName, "SP62830_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
    
115         Case 1
    
120             SaveSetting "SP50000", App.EXEName, "SP62840_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
    
125         Case 2
    
130             SaveSetting "SP50000", App.EXEName, "SP62880_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
    
135         Case 3
    
140             SaveSetting "SP50000", App.EXEName, "SP62890_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!

        End Select
  
        '105     If GintBelegArt = 0 Then
        '110         SaveSetting "SP50000", App.EXEName, "SP62830_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
        '        Else
        '115         SaveSetting "SP50000", App.EXEName, "SP62840_WndHwnd", ""           'HW 13.03.2014 Ver.: 6.2.104 Wird für SP60000 zum Fensterpositions-Reset benötigt, falls kein Titel gefunden wird!
        '        End If
        '
        '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
        
        '<Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
        ' # Um neue Module Angebot und AufBest. erweitert.
145     Select Case gintPrivBelegArt
        
            Case 0
            
150             Call ProgrammAus("283")
155             Protokoll iAppend, "*****PROGRAMM ENDE*****: 283 -> " & Now & vbCrLf & "***************************************************"  'vbCrLf &
            
160         Case 1
            
165             Call ProgrammAus("284")
170             Protokoll iAppend, "*****PROGRAMM ENDE*****: 284 -> " & Now & vbCrLf & "***************************************************"  'vbCrLf &
            
175         Case 2
            
180             Call ProgrammAus("288")
185             Protokoll iAppend, "*****PROGRAMM ENDE*****: 288 -> " & Now & vbCrLf & "***************************************************"  'vbCrLf &
            
190         Case 3
            
195             Call ProgrammAus("289")
200             Protokoll iAppend, "*****PROGRAMM ENDE*****: 289 -> " & Now & vbCrLf & "***************************************************"  'vbCrLf &
        
        End Select

        'ORIG`: 145     If gintPrivBelegArt = 0 Then                                            'Unterrutine in SP50000B.bas
        '
        '150         Call ProgrammAus("283")
        '155         Protokoll iAppend, "*****PROGRAMM ENDE*****: 283 -> " & Now & vbCrLf & "***************************************************"  'vbCrLf &
        '
        '        Else
        '
        '160         Call ProgrammAus("284")
        '165         Protokoll iAppend, "*****PROGRAMM ENDE*****: 284 -> " & Now & vbCrLf & "***************************************************"  'vbCrLf &
        '
        '        End If
        '</Modified by: IL at 11.10.2024, Ver.: 6.7.101 >

        '########## Formular Resizing: stoppen###################
        '
        'DeW, folgendes terminiert die Klasse, und loest dort
        'das _Terminate Ereigniss aus -> Speicherung der eingestellten
        'Vergroesserungswerte und Spaltenbreiten aus den TrueDBGrid
        'Info-Daten in der Registry
        '
205     Set cReSize = Nothing

        On Error Resume Next                                                    'HW 26.07.2013
  
210     Set objPRM = Nothing
215     Set objSQLAusw = Nothing
220     Set objSQLAuswDef = Nothing
225     Set objDAOSeek = Nothing
230     Set objHlp = Nothing
235     Set objLimit = Nothing
240     Set objPlausi = Nothing
245     Set objEmailSending = Nothing
250     Set objERechnung = Nothing
        
255     DisposeObjects Me                                                       'HW 26.07.2013

        Exit Sub

Fehler:
260     Call FehlerErklärung("frmSP52830", "Form_Unload")

End Sub

Public Function Speichern(Optional BelegID As Long, _
                          Optional tmp As Boolean, _
                          Optional Druck As Boolean) As Boolean

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        'BelegID wird mitgegeben beim Aktualisieren der Rechnung.
        'Sonst wird BelegID ermittelt und neue Rechnungssätze werden hinzugefügt.
        'Tmp = True -> Die Daten der Rechnung werden in die Tmp-Tabellen gespeichert
        'um die Vorschau der noch nicht gespeicherten Rechnung zu ermöglichen.
        Dim rs              As ADODB.Recordset

        Dim rsHaupt         As ADODB.Recordset

        Dim rsMAXTMPID      As ADODB.Recordset 'HW 03.12.2015

        Dim TmpZusatz       As String

        Dim RechnNr         As Long

        Dim SteuerPfl       As Double

        Dim SteuerFr        As Double

        Dim Ust             As Double

        Dim Limit           As Boolean

        Dim Lim             As Double

        Dim Betr            As Double

        Dim MText           As String

        Dim trans           As Boolean

        Dim GotNeueBelegID  As Boolean

        Dim oldBelegID      As Long
        
        Dim blnBelegNrFrei  As Boolean                                          'DF 22.01.2019 , Ver.: 6.5.109 : Zeigt, ob ein BelegNr bereits verwendent wurden (RAB)

        Dim blnBelegNrFortL As Boolean                                          'DF 22.01.2019 , Ver.: 6.5.109 : Zeigt, ob ein BelegNr in einer fortlaufender Reihenfolge sich befindet
        
        Dim lngArt          As Integer                                          'DF 22.01.2019 , Ver.: 6.5.109 :
        
        Dim intSofa         As Integer                                          'DF 24.01.2019 , Ver.: 6.5.109 : Zeigt, ob BelegNr in RAB aus eigenen oder aus Standardnummernkreis überprüft werden soll.
        
        Dim lngKrKreis      As Long                                             'DF 28.01.2019 , Ver.: 6.5.109 : Nummer des NummernKreisen, woher die BelegNr kommnt.
        
        Dim strMessage      As String
        
100     Set connSQL = New ADODB.Connection
105     Set rs = New ADODB.Recordset
110     Set rsHaupt = New ADODB.Recordset
115     Set rsMAXTMPID = New ADODB.Recordset
  
120     connSQL.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
125     connSQL.Open
        
        '<Added by: DFiebach at: 23.01.2019, Ver.: 6.5.109 >
130     If gintPrivBelegArt = 0 Then

135         lngArt = 1   'Ausg.-Rechn.

        Else
        
140         lngArt = 2   'Ausg.-Gutschr.

        End If

        '</Added by: DFiebach at: 23.01.2019, Ver.: 6.5.109 >
        
145     If Not tmp Then

            '***** SQL-Server *****
            'Bonitätslimit Überprüfung. Nur für Stammkunden
150         If Trim(txt1(0)) <> "" Then

155             objLimit.MCode = Trim(txt1(0))
160             Lim = Runden(objLimit.Limit, 0)

165             If Lim > 0 Then

170                 Betr = Runden(objLimit.GesamtBetrag, 0)

                    '<Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
175                 Select Case gintPrivBelegArt

                        Case 0
    
180                         Betr = Betr + frmRechnungErf.GetEndBetrag
    
185                     Case 1
    
190                         Betr = Betr + frmGutschriftErf.GetEndBetrag
    
195                     Case 2
    
200                         Betr = Betr + frmAngebotErf.GetEndBetrag
    
205                     Case 3
    
210                         Betr = Betr + frmAuftragsbestErf.GetEndBetrag

                    End Select

                    '175                 If gintPrivBelegArt = 0 Then
                    '180                     Betr = Betr + frmRechnungErf.GetEndBetrag
                    '                    Else
                    '185                     Betr = Betr + frmGutschriftErf.GetEndBetrag
                    '                    End If
                    '</Modified by: IL at 11.10.2024, Ver.: 6.7.101 >

215                 If Lim < Betr Then

220                     Protokoll iAppend, "Limit für " & Trim(txt1(0)) & " erreicht. Limit: " & Lim & " Gesamt-Betrag: " & Betr
225                     Limit = True
230                     MText = "A C H T U N G" & vbCrLf & "Das Kunden-Limit für Kunde: '" & Trim(txt1(0)) & "' ist überschritten !" & vbCrLf
235                     MText = MText & "Limit" & vbTab & ": " & Format(Lim, "###,###,##0.00") & vbCrLf
240                     MText = MText & "Aktuell" & vbTab & ": " & Format(Betr, "###,###,##0.00") & vbCrLf & vbCrLf
245                     MText = MText & "Soll der Beleg trotzdem gespeichert werden ?"

250                     If MsgBox(MText, vbCritical + vbYesNo + vbDefaultButton2, strMeldungCap) = vbNo Then

                            '<Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
255                         Select Case gintPrivBelegArt

                                Case 0
    
260                                 frmRechnungErf.cmd1(2).Enabled = False
    
265                             Case 1
    
270                                 frmGutschriftErf.cmd1(2).Enabled = False
    
275                             Case 2
    
280                                 frmAngebotErf.cmd1(2).Enabled = False
    
285                             Case 3
    
290                                 frmAuftragsbestErf.cmd1(2).Enabled = False

                            End Select

                            '230                         If gintPrivBelegArt = 0 Then
                            '235                             frmRechnungErf.cmd1(2).Enabled = False
                            '                            Else
                            '240                             frmGutschriftErf.cmd1(2).Enabled = False
                            '                            End If
                            '</Modified by: IL at 11.10.2024, Ver.: 6.7.101 >

                            Exit Function

                        End If
                        
                    End If
                    
                End If
                
            End If
            
295         If IsNumeric(BelegNr) Then

300             If CLng(BelegNr) = 0 Then

305                 BelegNr = ""

                Else
                    
                    '<Removed by: DFiebach at: 22.01.2019, Ver.: 6.5.X >
                    '245                 If Not IstBelegNrFrei(CLng(BelegNr), glngBelegID, gintPrivBelegArt) Then
                    '</Removed by: DFiebach at: 22.01.2019, Ver.: 6.5.X >

                    '<Removed by: GW at: 25.05.2020, Ver.: 6.6.101 >
                    'rausgenommen, damit die Logik bei allen Druckmodulen gleich ist. Bei allen Druckmodulen ist es möglich
                    ' eine Belegnummer mehrmals zu verwenden.
                    '260                 blnBelegNrFrei = IstBelegNrFrei(CLng(BelegNr), glngBelegID, gintPrivBelegArt)
                    '
                    '                    '<Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
                    '265                 If Not blnBelegNrFrei Then
                    '                        '</Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
                    '
                    '270                     strMessage = getMessage(2173)
                    '
                    '275                     strMessage = Replace$(strMessage, "%1", BelegNr)
                    '
                    '280                     MsgBox strMessage, vbInformation, strMeldungCap
                    '285                     Protokoll iAppend, "Speichern: Die vorgeschlagene Beleg-Nummer ''" & BelegNr & "'' wurde bereits vergeben. Die Beleg-Nummer wird automatisch beim Drucken ermittelt."
                    '290                     BelegNr = ""
                    '
                    '                    End If

                    '</Removed by: GW at: 25.05.2020, Ver.: 6.6.101 >
                    
                    '<Removed by: DFiebach at: 24.01.2019, Ver.: 6.5.109 >
                    ' # Überprüfung auf fortlaufende BelegNr bei manueller Eingabe wieder rausgenommen, da nicht klar in welchem NrKreis dann die BelegNr geprüft wird.
                    '                    '<Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
                    '
                    '300                 blnBelegNrFortL = IstBelegNrFortlaufend("5700_Haupt", CLng(BelegNr), glngBelegID, lngArt, intSoFa, True)
                    '
                    '305                 If blnBelegNrFortL = False Then
                    '
                    '310                     Protokoll iAppend, vbCrLf & "Speichern -> BELEGNUMMER NICHT FORTLAUFEND. ABBRUCH DURCH BENUTZER. BelegID " & BelegID & ", BelegNr " & BelegNr
                    '
                    '                        Exit Function
                    '
                    '                    End If
                    '
                    '                    '</Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
                    '</Removed by: DFiebach at: 24.01.2019, Ver.: 6.5.109 >
                    
                End If
                
            End If
            
            '<Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
310         blnBelegNrFrei = False
315         blnBelegNrFortL = False
            '</Added by: DFiebach at: 22.01.2019, Ver.: 6.5.109 >
            
            'Or Check1(0).Value = 1
320         If Druck And (Check1(0).value = 0) Then  'Zwischenablage-Belege bekommen keine automatische Belegnummer und Druckstatus bleibt = 0.

                '430         If txt1(15) = "" Then
325             If BelegNr = "" Then
                    
330                 BelegNr = 0
                    
                    '<Removed by: DFiebach at: 01.02.2019, Ver.: 6.5.X >
                    '                    'In diesem Fall muss eine Belegnummer generiert werden.
                    '                    'Im SQL-Server muss das vor der Transaktion ausgeführt werden
                    '                    '(In der SQL-Transaktion darf nur 1 mal auf die Tabelle zugegrifen werden. Das wäre nicht der Fall, wenn die Belegnummer mehrmals ermittelt werden müsste).
                    '
                    '                    Do
                    '
                    '                        '<Added by: DFiebach at: 28.01.2019, Ver.: 6.5.109 >
                    '                        ' # NrKreis in Variable ausgelagert, da diese untern bei Überprüfung auf "fortlaufend" noch benutzt wird.
                    '315                     lngKrKreis = NummernKreisWaehlenSQL(gintPrivBelegArt + 8)
                    '320                     RechnNr = NummernKreisSQL(lngKrKreis)  'Prüfen, ob allgemainer Nummernkreis für Rechnungen benutzt werden soll.
                    '                        '</Added by: DFiebach at: 28.01.2019, Ver.: 6.5.109 >
                    '
                    '                        '<Removed by: DFiebach at: 28.01.2019, Ver.: 6.5.109 >
                    '                        '315                     RechnNr = NummernKreisSQL(NummernKreisWaehlenSQL(gintPrivBelegArt + 8))  'Prüfen, ob allgemainer Nummernkreis für Rechnungen benutzt werden soll.
                    '                        '</Removed by: DFiebach at: 28.01.2019, Ver.: 6.5.109 >
                    '
                    '325                     If RechnNr = -1 Then
                    '
                    '                            'Fehler beim Ziehen der Rechnungsnummer.
                    '330                         Call msgText(1, 338, 0, 0, 0)
                    '
                    '                            'Der Nummernkreis ist von einem anderen Benutzer kurzzeitig gesperrt. Bitte versuchen Sie den Vorgang nochmals auszuführen.
                    '335                         MsgBox GsMsgText(0), vbInformation, strMeldungCap
                    '340                         Protokoll iAppend, vbCrLf & "Speichern -> Nummernkreis gesperrt. Rollback. BelegID :" & BelegID
                    '
                    '                            Exit Function
                    '
                    '                        End If
                    '
                    '                        'Rechnungsnummer muss genau so ermittelt werden, wie das im Programm 571 ist.
                    '                        'HW 02.05.2011 Ver.: 6.1.102 - hier darf die BelegNr nicht hochgezählt werden, weil in 571 die Logik geändert wurde!
                    '                        '520             RechnNr = RechnNr + 1
                    '
                    '                        'Wenn die ermittelte RechnNr bereits existiert, neue RechnNr ermitteln.
                    '
                    '                        '<Modified by: DFiebach at 21.01.2019, Ver.: 6.5.109 >
                    '                        '
                    '                        ' # <IstBelegNrFortlaufend> in die Scheifenbedingung hinzugefügt.
                    '                        '
                    '                        ' # ORIGINAL
                    '                        '300                 Loop Until IstBelegNrFrei(RechnNr, glngBelegID, gintPrivBelegArt)
                    '                        '
                    '                        ' # NEU, Überprüfung der BelegNummer in RAB, statt eigenen Tabellen
                    '                        '
                    '
                    '                        'Die Variable wird in <NummernKreisWaehlenSQL> gefüllt, und zeigt aus welchem NrKreis die RechnNr stamm, aus Eingenem( = False) oder Standard ( = True)
                    '345                     If SP52800B.gSFBelegStandardNrKreis Then
                    '
                    '350                         intSoFa = 0   'Bei der Überpürfung auf fortlaufende BelegNr wird im Standard-NrKreis (Standard-Faktura) geguckt
                    '
                    '                        Else
                    '
                    '355                         intSoFa = 1   'Bei der Überpürfung auf fortlaufende BelegNr wird im Eigenen (SF) - NrKreis (Standard-Faktura) geguckt
                    '
                    '                        End If
                    '
                    '360                     blnBelegNrFrei = False
                    '365                     blnBelegNrFortL = False
                    '
                    '                        'DF 29.01.2019, Ver.6.5.109: An dieser Stelle war zuerst die Überprüfung, ob BelegNr in den eigenen Tabellen vorhanden ist -> <IstBelegNrFrei>
                    '                        'Diese wurde rausgenommen, da die BelegNr nicht mehr in den Tabellen mit ungedruckten Belegen abgespeichert wird, und auf die Überprüfung der BelegNr in RAB umgestellt -> <CheckBelegNrInAusBuchFrei>
                    '                        'Diese wurde auch rausgenommen, da bei Besprechung als Überflüssig festgestellt wurde.
                    '                        'Die Variable wurde einfach auf True gesetzt. Hier muss dann die Überprüfung wieder aktiviert werden.
                    '370                     blnBelegNrFrei = True 'CheckBelegNrInAusBuchFrei(RechnNr, True)
                    '
                    '375                     If blnBelegNrFrei Then
                    '
                    '380                         blnBelegNrFortL = IstBelegNrFortlaufend("5700_Haupt", RechnNr, glngBelegID, lngArt, intSoFa, True, programmNr, SP52800B.gSFBelegStandardNrKreis, lngKrKreis)
                    '
                    '385                         If blnBelegNrFortL = False Then
                    '
                    '390                             Protokoll iAppend, vbCrLf & "Speichern -> BELEGNUMMER NICHT FORTLAUFEND. ABBRUCH DURCH BENUTZER. BelegID :" & BelegID & ", BelegNr " & CStr(RechnNr)
                    '395                             SP52800B.gSFBelegStandardNrKreis = False
                    '
                    '                                Exit Function
                    '
                    '                            End If
                    '
                    '                        End If
                    '
                    '400                     SP52800B.gSFBelegStandardNrKreis = False
                    '
                    '405                 Loop Until blnBelegNrFrei
                    '
                    '                    '
                    '                    '</Modified by: DFiebach at 21.01.2019, Ver.: 6.5.109 >
                    '
                    '                    '530           txt1(15) = RechnNr
                    '410                 BelegNr = RechnNr
                    '415                 Protokoll iAppend, "Speichern beim Einzelldruck. Automatisch ermittelte Beleg-Nummer: " & RechnNr & " BelegID: " & BelegID & ", BelegNr " & CStr(RechnNr)
                    '</Removed by: DFiebach at: 01.02.2019, Ver.: 6.5.X >

                End If
                
            End If

            '***** SQL-Server *****
        Else
        
335         TmpZusatz = "Tmp"

            '600       GDB.Execute "DELETE FROM [2800_HauptTmp] WHERE [BelegID] = " & BelegID , dbFailOnError
            'HW 03.08.2011 Ver.: 6.1.106 hier habe ich dieses FailOnError weggenommen, so schmeißt er keinen Fehler mehr !
            '590       GDB.Execute "DELETE FROM [2800_HauptTmp] WHERE [BelegID] = " & BelegID
            
340         connSQL.Execute "DELETE FROM [2800_HauptTmp] WHERE [BelegID] = " & BelegID

        End If
  
NeueBelegID:

345     If GotNeueBelegID Then

350         Err.Clear
355         connSQL.Rollback

360         BelegID = 0
365         GotNeueBelegID = False

370         trans = False

        End If

        'DH, 11.11.2015, 6.4.111, Dafuer sorgen, dass keine doppelten BelegIDs in den Tabellen entstehen
        '                         und das beim Beleg-Storno durch das Ausgangsbuch die Rechnung neu
        '                         gedruckt werden kann und auch korrekt an das Ausgangsbuch uebergeben wird
        '#####
375     oldBelegID = BelegID                                                    'Die BelegID speichern mit welcher in die Funktion eingesprungen wird

380     If BelegID = 0 Then                                                     'BelegID = 0, die Rechnung wurde also noch nie gedruckt

385         If tmp Then
                'MAX ID aus TMP
390             BelegID = GetMAXID_FROM_TABLE("2800_HauptTmp", "BelegID")
395             BelegID = BelegID + 1
            Else
400             BelegID = GetNrKreisValue_EXTENDED(connSQL, 32)                 'Neue ID aus dem Nummernkreis ziehen
            End If

        Else                                                                    'Bereits eine BelegID vorhanden

            'HW 03.12.2015 Wenn nicht Tmp dann muss überprüft werden!
405         If Not tmp Then

410             Call CheckBelegID(BelegID, "[5700_Haupt]")                      'Wenn die BelegID bereits in der [5700_Haupt] vorhanden ist, neue BelegID ziehen

                '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
                '# Um neue Module Angebot und AufBest. erweitert.
                Dim strTableName As String

415             Select Case GintBelegArt

                    Case 0
    
420                     strTableName = "[2800_BelegArchiv_Rng]"
    
425                 Case 1
    
430                     strTableName = "[2800_BelegArchiv_Gut]"
    
435                 Case 2
    
440                     strTableName = "[2800_BelegArchiv_Ang]"
    
445                 Case 3
    
450                     strTableName = "[2800_BelegArchiv_Auf]"

                End Select
                
455             Call CheckBelegID(BelegID, strTableName)
                
                'ORIG
                'Call CheckBelegID(BelegID, IIf(GintBelegArt = 0, "[2800_BelegArchiv_Rng]", "[2800_BelegArchiv_Gut]"))
                '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >

            End If

460         If BelegID <> oldBelegID Then                                       'Wenn sich die BelegID geaendert hat, die BelegID in den SoFa-Tabellen aktualisieren

465             connSQL.Execute "UPDATE [2800_Haupt] SET BelegID = " & BelegID & " WHERE BelegID = " & oldBelegID
470             connSQL.Execute "UPDATE [2800_Folge] SET BelegID = " & BelegID & " WHERE BelegID = " & oldBelegID

            End If

        End If
  
        'HW 03.12.2015 Auf TMP abgefragt
        '*
475     If tmp Then

480         glngBelegIDTmp = BelegID
485         rsHaupt.Open "SELECT * FROM [2800_Haupt" & TmpZusatz & "] WHERE BelegID = " & BelegID, connSQL, adOpenKeyset, adLockOptimistic

        Else
        
490         glngBelegID = BelegID                                               'Die BelegID Variable ausserhalb der Funktion aktualisieren
495         rsHaupt.Open "SELECT * FROM [2800_Haupt" & TmpZusatz & "] WHERE BelegID = " & BelegID, connSQL, adOpenKeyset, adLockOptimistic

        End If

        '*
        '#####
        
500     If oldBelegID = 0 Then                                                  'Wenn die BelegID eine 0 WAR, wurde die Rechnung noch nie gedruckt (weder gedruckt noch im Ausgangsbuch storniert)

505         GotNeueBelegID = True

510         connSQL.BeginTrans
515         trans = True

520         rsHaupt.AddNew
525         rsHaupt!BelegID = BelegID
530         rsHaupt.Update                                                      'Falls die ermittelte BelegID in der Zwischenzeit gespeichert wurde, tritt ein Fehler auf. Der Vorgang wird wiederholt.' Sofort Updaten um mehrfache Vergabe der BelegID zu vermeiden
    
535         If tmp = False Then

540             glngBelegID = BelegID
                
545             objPRM.FindFirstString = "name = 'statusBarText' AND index = 2" 'DF 16.01.2015 : StatusBAr Text aus PRM holen.
550             sta1.Panels(3).text = objPRM.caption("Erfassungs-Nr.:") & " " & glngBelegID
555             objPRM.FindFirstString = ""

            End If

            '**********
            
560         rsHaupt.MoveLast
565         rsHaupt!ErstDat = Now
570         rsHaupt!ErstVon = GsUser
575         rsHaupt!AendDat = Now
580         rsHaupt!AendVon = GsUser

        Else
        
585         connSQL.BeginTrans
590         trans = True
    
595         If tmp Then

600             rsHaupt.AddNew
605             rsHaupt!BelegID = BelegID
610             rsHaupt!ErstDat = Now
615             rsHaupt!ErstVon = GsUser
620             rsHaupt.Update                                                 'Da bei AddNew Stehen dem neuen Datensatz noch keine Standardwerte zur Verfügung (so wie bei Access der Fall ist).

625             rsHaupt.MoveLast
    
            End If

630         If rsHaupt.RecordCount = 0 Then                                     'HW 18.11.2015 Wenn keine Daten vorhanden -> neu anlegen
635             rsHaupt.AddNew
            End If

640         rsHaupt!AendDat = Now
645         rsHaupt!AendVon = GsUser

        End If
  
650     rsHaupt!MCode = txt1(0)
655     rsHaupt!Name1 = txt1(1)
660     rsHaupt!Name2 = txt1(2)
665     rsHaupt!AnsprPartner = txt1(3)
670     rsHaupt!Postfach = txt1(4)
675     rsHaupt!PLZ1 = txt1(5)
680     rsHaupt!ort1 = txt1(6)
685     rsHaupt!Straße = txt1(7)
690     rsHaupt!Lkz = txt1(8)
695     rsHaupt!Plz = txt1(9)
700     rsHaupt!Ort = txt1(10)
705     rsHaupt!ORTSTEIL = txt1(11)
710     rsHaupt!KtoNr = txt1(12)
715     rsHaupt!KtoKnz = txt1(34)
720     rsHaupt!SteuerNr = txt1(13)
725     rsHaupt!Uid = txt1(14)
730     rsHaupt!InternerVermerk = txt1(41)                                      'DF 14.01.2015 : Interner Vermerk hinzugefügt.
  
        'HW 09.07.2012 Ver.: 6.1.114  in txt1(38).Text steht z.B. "steuerfrei" aber das Datenfeld ist ein Integer! FEHLER!
        '1310    rsHaupt!Ust = txt1(39)
735     rsHaupt!Ust = intSteuerTyp
  
740     rsHaupt!ZwAblage = Check1(0).value
  
        '1330    If IsDate(txt1(16)) Then rsHaupt!BelegDatum = txt1(16) 'Manuelle Erfassung
        
        'DF 13.02.2019 , Ver.: 6.5.109 : And Druck hinzugefügt, damit das BelegDatum nur beim richtigen Druck gespeichert wird,
        '                                sonst stand das Datum nach der einfachen Vorschau des Beleges drin.
745     If IsDate(belegDatum) And Druck Then rsHaupt!belegDatum = belegDatum 'Manuelle Erfassung
        'HW 23.07.2015 Ver.: 6.4.108
        '* Hier muss abgeprüft werden, wenn die BelegNr schon vergeben ist und Druck = 0 wurde ein belegstorno gemacht!
        '* In diesem Fall muss der Status RABuch auch wieder auf 0 geschaltet werden, damit der Beleg in den Ein/Ausgangs-Büchern importiert werden kann!
        '++++++++
        '1340    If IsNumeric(txt1(15)) Then rsHaupt!BelegNr = txt1(15) 'Manuelle Erfassung oder Drucken
        
750     If IsNumeric(BelegNr) Then

755         If rsHaupt!BelegNr = BelegNr Then

760             If rsHaupt!RABuch > 0 And rsHaupt!Druck = 0 Then

765                 rsHaupt!BelegNr = 0

                End If

            Else
            
770             rsHaupt!BelegNr = BelegNr 'Manuelle Erfassung oder Drucken
                
                '<Added by: DFiebach at: 28.01.2019, Ver.: 6.5.109 >
775             rsHaupt!BelegNrKReis = lngKrKreis
                '</Added by: DFiebach at: 28.01.2019, Ver.: 6.5.109 >
                
            End If

        Else
        
            '++++++++

        End If

780     If Druck And rsHaupt!ZwAblage = 0 And modERechnung.CheckENCodeBeiVerpackung(1, CStr(glngBelegID), False) Then                   'Zwischenablage-Belege bekommen keine automatische Belegnummer und Druckstatus bleibt = 0.

            '</Modified by: Project Administrator at 9.3.2024-09:12:08 on machine: T017>

785         rsHaupt!Druck = 1

790         If Not IsDate(rsHaupt!belegDatum) Then rsHaupt!belegDatum = GdtDatum

795         If IsDate(objDruckOptionen.CurrentValutaDatum) Then rsHaupt.Fields("ValutaDatum").value = objDruckOptionen.CurrentValutaDatum

        End If
        
        'DF 11.06.2020 , Ver.: 6.6.102 : Ust Senkung Umstellung txt1(17/18) auf Labels umgestellt
        '                                Währungs-Schl(Beleg- und Vergleichs werden jetzt auch gespeichert)
800     rsHaupt!Art = gintPrivBelegArt
805     rsHaupt!Wrg1 = lbl2(17).caption 'txt1(17).text
810     rsHaupt!Wrg2 = lbl2(18).caption 'txt1(18).text
815     rsHaupt!WrgSchl = Trim$(txt1(17).text)
820     rsHaupt!WrgSchlVgl = Trim$(txt1(18).text)
825     rsHaupt!WrgAusw = Check1(1).value
830     rsHaupt!Kurs = txt1(19).text
835     rsHaupt!MwSt = txt1(20).text
840     rsHaupt!ZSkto = txt1(21).text
845     rsHaupt!ZSktoTage = txt1(22).text
850     rsHaupt!ZTage = txt1(23).text
855     rsHaupt!PPRechnung = IIf(Trim$(txt1(24).text) = "", 0, Trim$(txt1(24).text))
  
860     rsHaupt!KostSchl = txt1(25).text
865     rsHaupt!FiBuSchl = txt1(26).text

        '<Added by: GW at: 02.04.2019, Ver.: 6.5.110 >
        'Belegdatum bei Vorlage und Ungedruckt speichern
870     If Not IsDate(rsHaupt!belegDatum) Then rsHaupt!belegDatum = GdtDatum
        '</Added by: GW at: 02.04.2019, Ver.: 6.5.110 >
        
        'HW, 26.06.2017, Abprüfungen hinzugefügt
875     If IsNumeric(txt1(27)) Then
880         If Not IsEmpty(txt1(27)) Then
885             rsHaupt!KostKonto = txt1(27)
            Else
890             rsHaupt!KostKonto = 0
            End If
        End If

        'HW, 26.06.2017
895     If IsNumeric(txt1(28)) Then
900         If Not IsEmpty(txt1(28)) Then
905             rsHaupt!FibuKonto = txt1(28)
            Else
910             rsHaupt!FibuKonto = 0
            End If
        End If
  
915     If IsDate(txt1(29)) Then
920         rsHaupt!Datum = txt1(29)
        Else
925         rsHaupt!Datum = Null
        End If

930     rsHaupt!AbfZus = txt1(30)
  
        'HW, 26.06.2017
935     If IsNumeric(txt1(31)) Then
940         If Not IsEmpty(txt1(31)) Then
945             rsHaupt!AbfPos = txt1(31)
            Else
950             rsHaupt!AbfPos = 0
            End If
        End If

955     rsHaupt!KfzSchl = txt1(33)

        'HW, 26.06.2017
960     If IsNumeric(txt1(32)) Then
965         If Not IsEmpty(txt1(32)) Then
970             rsHaupt("Relation") = Abs(DEFAULT_VALUE(txt1(32), 0, True))
            Else
975             rsHaupt("Relation") = 0
            End If
        End If

980     If Trim(txt1(35)) <> "<alle>" Then rsHaupt!ErfNr = txt1(35)
985     rsHaupt!KostenArt = txt1(36)

        '<Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
990     Select Case gintPrivBelegArt

            Case 0
    
995             frmRechnungErf.FolgeSpeichern BelegID, tmp, SteuerPfl, SteuerFr
    
1000        Case 1
    
1005            frmGutschriftErf.FolgeSpeichern BelegID, tmp, SteuerPfl, SteuerFr
    
1010        Case 2
    
1015            frmAngebotErf.FolgeSpeichern BelegID, tmp, SteuerPfl, SteuerFr
    
1020        Case 3
    
1025            frmAuftragsbestErf.FolgeSpeichern BelegID, tmp, SteuerPfl, SteuerFr

        End Select

        '965     If gintPrivBelegArt = 0 Then
        '970         frmRechnungErf.FolgeSpeichern BelegID, tmp, SteuerPfl, SteuerFr
        '        Else
        '975         frmGutschriftErf.FolgeSpeichern BelegID, tmp, SteuerPfl, SteuerFr
        '        End If
        '</Modified by: IL at 11.10.2024, Ver.: 6.7.101 >
  
        'Der Betrag kann erst gespeichert werden nach dem die Folge gespeichert wurde.
        '***** SQL-Server *****
        'EndBetraege "2800_Folge" & TmpZusatz, BelegID, SteuerPfl, SteuerFr
        '***** SQL-Server *****
        
1030    If rsHaupt!MwSt = 0 Then                                               'IL 01.08.2024, bei SteuerPfrei DL soll SteuerSatz der Währung für die Berechnung der SteuerPflichtigen Positionen genommen werden

1035        dblMwSt = GetWaehrung(rsHaupt!WrgSchl, False).MwSt
            
        Else
            
1040        dblMwSt = rsHaupt!MwSt
            
        End If
        
        '<Removed by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >
        '945     Ust = Runden((SteuerPfl * rsHaupt!MwSt / 100), 2)
        '</Removed by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >
        
        '<Added by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >
1045    Ust = SteuerBetrag(CCur(SteuerPfl), Format(CCur(dblMwSt), "#.00"))     'IL 01.08.2024    rsHaupt!MwSt ---> dblMwSt
        '</Added by: DFiebach at: 18.09.2019, Ver.: 6.5.112 >

        'HW 24.05.2016 Ver.: 6.4.120
1050    If GesamtIstBrutto Then
1055        rsHaupt!Betrag = SteuerPfl + SteuerFr
        Else
1060        rsHaupt!Betrag = SteuerPfl + Ust + SteuerFr
        End If
  
        'Die 2800_Haupt muss gesperrt bleiben bis die 2800_Folge upgedatet ist.
        'So wird verhindert, dass ein anderer Benutzer die Tabelle in der Zwischenzeit verändert.
        'HW 29.08.2007 vers.: 5.4.116 rsHaupt!Druck ... Auf Null abfragen, weil sonst ein Fehler passiert und Stadardwert setzen
1065    If (IsNull(rsHaupt!Druck)) Then
1070        gintDruck = 0
        Else
1075        gintDruck = rsHaupt!Druck
        End If
  
1080    If IsDate(txt1(37)) Then rsHaupt!vonDatum = txt1(37)
1085    If IsDate(txt1(38)) Then rsHaupt!bisDatum = txt1(38)
    
1110    rsHaupt.Update

1115    rsHaupt.MoveLast
  
1120    blnBelegNeu = False                                                     'DF 16.01.2015 : Beleg gespeichert -> Zeiger auf False setzen. Im Rechnung Fenster kann Beleg sofort storniert werden.
  
        'If gintDruck = 1 And rsHaupt!ZwAblage = 0 And Not tmp Then SpeditionsBuch BelegID, gintPrivBelegArt
1125    If gintDruck = 1 And rsHaupt!ZwAblage = 0 And Not tmp Then SpeditionsBuch rsHaupt, SteuerPfl, SteuerFr

        '***** SQL-Server *****
1130    Protokoll iAppend, ">SPEICHERN -> BelegNr: " & rsHaupt!BelegNr & " BelegID: " & BelegID & " Druck: " & gintDruck & " Art: " & gintPrivBelegArt

1135    If trans Then connSQL.CommitTrans                                       'HW 26.07.2011 Ver.: 6.1.105 -Abfrage auf trans hinzugefügt

1140    Speichern = True

1145    If connSQL.state = adStateOpen Then connSQL.Close

        Exit Function

Fehler:

1150    If trans Then connSQL.RollbackTrans                                     'HW 16.03.2015 Ver.: 6.4.104  auf trans abgeprüft!
1155    Call FehlerErklärung("frmSP52830", "Speichern()")

1160    If connSQL.state = adStateOpen Then connSQL.Close
End Function


Private Sub mnuBearb1_Click(Index As Integer)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
  
        'Call cmd1_Click(Index)
100     Select Case Index

            Case 0 'Rechnung
105             Call cmd1_Click(5)

110         Case 5 'Leeren
115             Call cmd1_Click(0)

120         Case 7 'Beleg Suchen
125             Call mnuSuch_Click(Index)

130         Case 8 'Beleg-Archiv

                '<Removed by: GW at: 21.02.2020, Ver.: GOBD >
                '135             If objBelegArchiv Is Nothing Then Set objBelegArchiv = New SPBelegArchiv.clsBelegArchiv
                '
                '140             Select Case programmNr
                '
                '                    Case "283", "285"
                '145                     Call objBelegArchiv.Show(E_BelegArt.Sonderfaktura_Rechnung)
                '
                '150                 Case "284", "286"
                '155                     Call objBelegArchiv.Show(E_BelegArt.Sonderfaktura_Gutschrift)
                '
                '                End Select
                '</Removed by: GW at: 21.02.2020, Ver.: GOBD >
                                 
135             Call ShowBelegArchiv

        End Select
  
        '***Beginn
        Exit Sub

Fehler:
140     Call FehlerErklärung("frmSP52830", "mnuBearb1_Click")
        '***Ende
End Sub

Private Sub mnuBesch_Click()

        On Error GoTo Fehler

100     objPRM.FindFirstString = "name = 'mnuBesch' "
    
105     If shiftPressed = True Then
            'Die UMSCHALT-TASTE ist gedrückt.
            'Hilfetexte können erfast oder bearbeitet werden.
110         objHlp.HlpShow HlpWrite, objPRM.HlpID & "_283" '& programmNr
        Else
115         objHlp.HlpShow HlpRead, objPRM.HlpID & "_283" '& programmNr
        End If

120     shiftPressed = False
  
        Exit Sub

Fehler:
125     Call FehlerErklärung("frmSP52830", "mnuBesch_Click")

End Sub

Private Sub mnuDat1_Click(Index As Integer)

        On Error GoTo Fehler
    
100     Select Case Index

            Case 6 'Schließen
105             Call cmd1_Click(6)
        End Select

        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52830", "mnuDat1_Click")
      
End Sub

Private Sub mnuOpt1_Click(Index As Integer)

        On Error GoTo Fehler
        
100     mnuOpt1(Index).Checked = Not mnuOpt1(Index).Checked

105     Select Case Index

            Case 1 'Ansprechpartner übernehmen
            
110             SaveSetting "SP50000", "SP52800", "SP52830Ansprechpartner", CInt(mnuOpt1(Index).Checked)

115         Case 2 'Adresse auf Folgeseiten drucken

                '<Modified by: GW at 24.04.2019, Ver.: 6.5.111 >
                '120             SaveSetting "SP50000", "SP52800", "SP52830Ansprechpartner", CInt(mnuOpt1(Index).Checked)
120             SaveSetting "SP50000", "SP52800", "SP52830FolgeseitenKurzDrucken", CInt(mnuOpt1(Index).Checked)
125             blnFolgeseitenKurzDrucken = mnuOpt1(Index).Checked
                '</Modified by: GW at 24.04.2019, Ver.: 6.5.111 >

130         Case 3 'Bearbeiter drucken

135             SaveSetting "SP50000", "SP52800", "SP52830Bearbeiter", CInt(mnuOpt1(Index).Checked)
140             BearbeiterDrucken = mnuOpt1(3).Checked  'HW 29.04.216

145         Case 4 'HW 24.05.2016 Ver.: 6.4.120 GESAMT = Brutto

150             SaveSetting "SP50000", "SP52800", "SP52830GesamtIstBrutto", CInt(mnuOpt1(Index).Checked)

                '155             If gboolHatKind Then
                '
                '160                 If MsgBox(GetMessage(2382), vbYesNo + vbExclamation, strMeldungCap) = vbYes Then
                '
                '165                     Select Case gintPrivBelegArt
                '
                '                            Case 0 'Rechnung
                '
                '170                             frmRechnungErf.FolgeZeigen 0
                '
                '175                         Case 1 'Gutschrift
                '
                '180                             frmGutschriftErf.FolgeZeigen 0
                '
                '185                         Case 2 'Angebot
                '
                '190                             frmAngebotErf.FolgeZeigen 0
                '
                '195                         Case 3 'Auftragsbestetigung
                '
                '200                             frmAuftragsbestErf.FolgeZeigen 0
                '
                '                        End Select
                '
                '205                     GesamtIstBrutto = mnuOpt1(Index).Checked
                '
                '                    Else
                '
                '210                     mnuOpt1(Index).Checked = Not mnuOpt1(Index).Checked
                '
                '                    End If
                '
                '                Else
                
155             GesamtIstBrutto = mnuOpt1(Index).Checked

                'End If

                '<Added by: GW at: 02.04.2019, Ver.: 6.5.110 >
                
160         Case 5 'MaskeLeeren

165             If mnuOpt1(6).Checked Then
170                 mnuOpt1(6).Checked = Not mnuOpt1(6).Checked
175                 SaveSetting "SP50000", "SP52800", "SP52830MaskeSchliessen", CInt(mnuOpt1(6).Checked)
180                 blnMaskeSchliessen = mnuOpt1(6).Checked
                End If

185             SaveSetting "SP50000", "SP52800", "SP52830MaskeLeeren", CInt(mnuOpt1(Index).Checked)
190             blnMaskeLeeren = mnuOpt1(Index).Checked

195         Case 6 'MaskeSchließen

200             If mnuOpt1(5).Checked Then
205                 mnuOpt1(5).Checked = Not mnuOpt1(5).Checked
210                 SaveSetting "SP50000", "SP52800", "SP52830MaskeLeeren", CInt(mnuOpt1(5).Checked)
215                 blnMaskeLeeren = mnuOpt1(5).Checked
                End If

220             SaveSetting "SP50000", "SP52800", "SP52830MaskeSchliessen", CInt(mnuOpt1(Index).Checked)
225             blnMaskeSchliessen = mnuOpt1(Index).Checked
                
                '</Added by: GW at: 02.04.2019, Ver.: 6.5.110 >
                
230         Case 50 'Druckerauswahl

235             SaveSetting "SP50000", "SP52800", "SP52830DruckerDialog", CInt(mnuOpt1(Index).Checked)

        End Select
        
        Exit Sub

Fehler:
240     Call FehlerErklärung("frmSP52830", "mnuOpt1_Click")
        
End Sub

Private Sub mnuSuch_Click(Index As Integer)
    '<Removed by: IL at: 15.10.2024, Ver.: 6.7.101 >
    '# Aufgrund der Aufteilung in mehrere Tabs entfernt
    '        Dim frm As Form
    '
    '        On Error GoTo Fehler
    '
    '100     Select Case Index
    '
    '            Case 0, 1, 2                                                        'Ungedruckten,Gedruckten,Vorlagen
    '
    '105             BelegSuchen Index
    '
    '        End Select
    '
    '        Exit Sub
    '
    'Fehler:
    '110     Call FehlerErklärung("frmSP52830", "mnuSuch_Click")
    '</Removed by: IL at: 15.10.2024, Ver.: 6.7.101 >
End Sub

Private Sub mnuUbernehmenA_Click(Index As Integer)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       mnuUbernehmenA_Click
        ' Description:       Zeigt die Dokumentauswahltabelle für den ausgewählten Dokumenttyp an
        ' Created by :       IL
        ' Date-Time  :       15.10.2024-14:54:43
        '
        ' Parameters :       Index (Integer)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler
        
100     BelegSuchen Sonderfaktura_Angebot, Index
          
        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52830", "mnuUbernehmenA_Click()")
End Sub

Private Sub mnuUbernehmenB_Click(Index As Integer)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       mnuUbernehmenB_Click
        ' Description:       Zeigt die Dokumentauswahltabelle für den ausgewählten Dokumenttyp an
        ' Created by :       IL
        ' Date-Time  :       15.10.2024-14:54:43
        '
        ' Parameters :       Index (Integer)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler
        
100     BelegSuchen Sonderfaktura_Auftragsbestetigung, Index
        
        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52830", "mnuUbernehmenB_Click()")
End Sub

Private Sub mnuUbernehmenR_Click(Index As Integer)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       mnuUbernehmenR_Click
        ' Description:       Zeigt die Dokumentauswahltabelle für den ausgewählten Dokumenttyp an
        ' Created by :       IL
        ' Date-Time  :       15.10.2024-14:54:43
        '
        ' Parameters :       Index (Integer)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler
        
100     BelegSuchen Sonderfaktura_Rechnung, Index
        
        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52830", "mnuUbernehmenR_Click()")
End Sub

Private Sub mnuUbernehmenG_Click(Index As Integer)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       mnuUbernehmenG_Click
        ' Description:       Zeigt die Dokumentauswahltabelle für den ausgewählten Dokumenttyp an
        ' Created by :       IL
        ' Date-Time  :       15.10.2024-14:54:43
        '
        ' Parameters :       Index (Integer)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler
        
100     BelegSuchen Sonderfaktura_Gutschrift, Index
        
        Exit Sub

Fehler:
105     Me.MousePointer = vbDefault
110     Call FehlerErklärung("frmSP52830", "mnuUbernehmenG_Click()")
End Sub

Private Sub mnuUpdateInfo_Click()

        On Error GoTo Fehler

100     objPRM.FindFirstString = "name = 'mnuUpdateInfo' "
105     objHlp.UpdateAnzeigen = True
110     objHlp.UpdateCounter = gi_UpdateAenderung

115     If shiftPressed = True Then
            'Die UMSCHALT-TASTE ist gedrückt.
            'Hilfetexte können erfast oder bearbeitet werden.
            '60         objHlp.HlpShow HlpWrite, "UpdateAenderung" & "_" & programmNr
120         objHlp.HlpShow HlpWrite, "UpdateAenderung" & "_283"      'DH, 29.10.2015, 6.4.110, Bei der UpdateIfo soll jetzt immer nur auf den Text der Rechnung zugegriffen werden
125         shiftPressed = False
        Else
            '90         objHlp.HlpShow HlpRead, "UpdateAenderung" & "_" & programmNr
130         objHlp.HlpShow HlpRead, "UpdateAenderung" & "_283"
        End If

135     objHlp.UpdateAnzeigen = False
   
140     If Me.ActiveControl.name <> "" Then

            On Error Resume Next

145         objPRM.FindFirstString = "name = '" & Me.ActiveControl.name & "' AND index = " & Me.ActiveControl.Index
150         Err.Clear

            On Error GoTo 0

        End If

        Exit Sub

Fehler:
155     Call FehlerErklärung("frmSP52830", "mnuUpdateInfo_Click")
End Sub

Private Sub txt1_GotFocus(Index As Integer)

        On Error GoTo Fehler

100     If Index <> 18 Then 'Mandantenwährung (gesperrt für die Eingabe)
105         txt1(Index).SelStart = 0
110         txt1(Index).selLength = Len(txt1(Index))
115         txt1(Index).ForeColor = vbWindowText
120         txt1(Index).BackColor = &HC0E0FF  'hellorange
125         objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
            'txt1(Index).MaxLength = objPRM.EingabeLaenge(txt1(Index))
        End If

        '*************Alt
130     gvntMerker = txt1(Index)
  
        Exit Sub

Fehler:
135     Call FehlerErklärung("frmSP52830", "txt1_GotFocus()")
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

        On Error GoTo Fehler

        Dim rsUST   As ADODB.Recordset    'HW 11.03.2014

        Dim ret     As Integer         'HW 11.03.2014

        Dim bCancel As Boolean

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

                    Case 0, 3, 8, 12, 17, 18, 25, 26, 29, 30, 31, 34, 35, 36, 37, 38, 44  'IL 04.12.2024 , Ver.: 6.7.102 : 12(Konto-Nr) Hinzugefügt
                    
130                     Call cmdAuswahl_Click(Index)

135                 Case 39                                                     'HW 11.03.2014 Ver.: 6.2.103 Umsatzsteuer muss bei der Sonderfaktura eingegeben werden können im Fall von "Einmalkunden"

140                     Call cmdAuswahl_Click(1)

145                 Case 5, 6, 9, 10, 11                                        'PLZ, Ort1, Ortsteil

150                     If gstrEWerk <> "" Then Call cmdAuswahl_Click(Index)

                End Select

155         Case vbKeyReturn, vbKeyDown

160             Select Case Index

                    Case 0, 12                                                  'MCode/Konto-Nr                  'IL 04.12.2024 , Ver.: 6.7.102 : 12(Konto-Nr) Hinzugefügt

165                     If Trim(txt1(Index)) <> "" Then
170                         objSQLAusw.GetIfOnesHit = True
175                         Call cmdAuswahl_Click(Index)
180                         objSQLAusw.GetIfOnesHit = False
                        Else

185                         Select Case Index
                        
                                Case 0
                            
190                                 Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                            
195                             Case 12
                            
200                                 Call objPRM.SprungNeu("Vorwärts", Shift, 3, True)
                        
                            End Select

                        End If

205                     EWerkErmitteln

210                 Case 5, 6, 9, 10, 11                                        'PLZ, Ort1, Ortsteil

215                     If Trim(txt1(Index)) = "" Then
220                         Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex)
                        Else

225                         If gstrEWerk <> "" Then
230                             objDAOSeek.RSFind = True                        'Sofort übernehmen, wenn Satz eindeutig ist.
235                             Call cmdAuswahl_Click(Index)
240                             objDAOSeek.RSFind = False
                            Else
245                             Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                            End If
                            
                        End If
                        
250                 Case 17                                                     'Währung

                        'DF 11.06.2020 , Ver.: 6.6.102 : Statt ISO wird Schl geprüft -> ISO -> Schl (UST Senkung Umstellung)

255                     If objPlausi.RSOpen("SELECT Schl FROM [1100_Währungen] WHERE Schl = '" & Trim$(txt1(Index).text) & "'", True) Then

260                         Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)

                        End If

265                 Case 25                                                     'Kostenschlüssel

                        'HW, 22.12.2017, Bei Enter soll F2-Fenster aufgehen, wenn der Wert nicht gefunden wurde! GLeichzeitig soll ein gefundener Wert gleich übernommen werden!
270                     If txt1(Index).text <> "" Then
275                         objSQLAusw.GetIfOnesHit = True
280                         cmdAuswahl_Click Index
285                         objSQLAusw.GetIfOnesHit = False
                        Else
290                         Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                        End If

295                 Case 26                                                     'Sachkonten Schluessel
        
                        'HW, 22.12.2017, Bei Enter soll F2-Fenster aufgehen, wenn der Wert nicht gefunden wurde! GLeichzeitig soll ein gefundener Wert gleich übernommen werden!
300                     If txt1(Index).text <> "" Then
305                         objSQLAusw.GetIfOnesHit = True
310                         cmdAuswahl_Click Index
315                         objSQLAusw.GetIfOnesHit = False
                        Else
320                         Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                        End If

325                 Case 39                                                     'HW 11.03.2014 Textfeld im Form auf Locked = True gestellt!
                
                        'TODO: GW wurde auskommentiert, weil es falsch ist.
                        '      1) analysieren, ob man case 39 braucht, da das feld gesperrt ist
                        '      2) so machen wie case 26 oder wie es in Lademittel gemacht wurde
                        '###########################################
                        '590             If Not rsUST Is Nothing Then
                        '600                 On Error Resume Next
                        '610                 rsUST.Close
                        '620                 If Err.number <> 0 Then Err.Clear
                        '630                 On Error GoTo Fehler
                        '640             End If
                        '
                        '650             Set rsUST = New ADODB.Recordset
                        '
                        '                'TODO : FEHLER
                        '660             'rsUST.Open "SELECT Knz, KnzBez1 FROM Auswahl WHERE TabName = '1200_GrundKonditionen' AND FeldName = 'Ust' AND KnzBez1 Like '" & txt1(39).Text & "*'", gConn, adOpenStatic, adLockReadOnly
                        '                'objDEFAusw.RSOpen GetACCESSConnectionString(DEF_CONNECTION), "SELECT Knz, KnzBez1 FROM Auswahl WHERE TabName = '1200_GrundKonditionen' AND FeldName = 'Ust' ORDER BY Knz"
                        '670             ret = rsUST.RecordCount
                        '
                        '680             If ret = 0 Then
                        '690                 txt1(Index).Text = ""
                        '700                 cmdAuswahl_Click 1
                        '710             Else
                        '720                 If ret = 1 Then    ' bei 1 kann übernommen werden!
                        '730                     intSteuerTyp = CInt(rsUST("Knz"))
                        '740                     txt1(Index).Text = "" & CStr(rsUST("KnzBez1"))
                        '750                     Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                        '760                 Else
                        '770                     txt1(Index).Text = ""
                        '780                     cmdAuswahl_Click 1
                        '790                 End If
                        '800             End If
                        '
                        '810             rsUST.Close
                        '820             Set rsUST = Nothing
                        '###########################################
                        
330                     Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                        
335                 Case 42

340                     If objPRM.EingabeFehler(txt1(Index)) = False Then

                            'HW, 10.01.2018, Überprüfung
                            '*GW_05.03.2018
345                         If Index = 42 Then
350                             txt1_Validate 42, bCancel

355                             If bCancel Then

                                    Exit Sub

                                End If
                            End If

360                         Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)

                        End If
                    
365                 Case Else

370                     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index

                        'Weil objPRM.SprungNeu Validate-Ereignis nicht auslöst,
                        'muss die Umwandlung und Prüfung an der Stelle stattfinden.
375                     If Index = 29 Or Index = 37 Or Index = 38 Then          'Abfertigungs-Datum, Zeitraum

380                         txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

                            'GW_05.03.2018 Ver.6.5.106 -------------------------------------------------------------------
                            'Nicht 29, weil kein DatumBis < DatumVon nötig ist
385                         If Index <> 29 Then

390                             txt1_Validate Index, bCancel

395                             If bCancel = True Then

400                                 KeyCode = 0

                                    Exit Sub

                                Else
405                                 Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)
                                End If

                            End If

                        Else
                        
410                         txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , True, True)

                        End If

                        'GW-------------------------------------------------------------------------------

415                     If objPRM.EingabeFehler(txt1(Index)) = False Then

420                         If Index = 34 Then
                        
425                             Call objPRM.SprungNeu("Vorwärts", Shift, 1, True)
                        
                            Else

430                             Call objPRM.SprungNeu("Vorwärts", Shift, txt1(Index).TabIndex, True)

                            End If

                        End If

                End Select

435         Case vbKeyEscape, vbKeyUp

440             If gvntMerker <> "" Then

445                 If gvntMerker <> txt1(Index) Then

450                     txt1(Index) = gvntMerker

                    End If
                    
                End If

455             Select Case Index

                    Case 10             'Name1, Ort
                    
460                     Call objPRM.SprungNeu("Rückwerts", Shift, txt1(Index).TabIndex, True)

465                 Case 1

470                     Call objPRM.SprungNeu("Rückwerts", Shift, 3, True)

475                 Case 12, 34

480                     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
485                     txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

490                     If objPRM.EingabeFehler(txt1(Index)) = False Then

495                         Select Case Index

                                Case 12
    
500                                 Call objPRM.SprungNeu("Rückwerts", Shift, 4, True)
    
505                             Case 34
    
510                                 Call objPRM.SprungNeu("Rückwerts", Shift, 2, True)

                            End Select

                        End If

515                 Case 17               'Währung
                    
                        'DF 11.06.2020 , Ver.: 6.6.102 : Statt ISO wird Schl geprüft -> ISO -> Schl (UST Senkung Umstellung)
520                     If objPlausi.RSOpen("SELECT Schl FROM [1100_Währungen] WHERE Schl = '" & Trim$(txt1(Index).text) & "'", True) Then

525                         Call objPRM.SprungNeu("Rückwerts", Shift, txt1(Index).TabIndex, True)

                        End If

530                 Case Else

535                     objPRM.FindFirstString = "name = 'txt1' AND index = " & Index
540                     txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

545                     If objPRM.EingabeFehler(txt1(Index)) = False Then
550                         Call objPRM.SprungNeu("Rückwerts", Shift, txt1(Index).TabIndex, True)
                        End If

                End Select

555         Case vbKeyDelete

560             If txt1(Index).selLength = Len(txt1(Index)) Then

565                 Select Case Index

                        Case 29, 30, 31   'Abf-Statistik-Felder
570                         txt1(29) = ""
575                         txt1(30) = ""
580                         txt1(31) = ""
585                         txt1(35) = ""

590                     Case 35, 36
595                         txt1(Index) = ""
                    End Select

                End If

        End Select

        Exit Sub

Fehler:
600     Call FehlerErklärung("frmSP52830", "txt1_KeyDown()")
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
      
        On Error GoTo Fehler

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

        Exit Sub

Fehler:
145     Call FehlerErklärung("frmSP52830", "txt1_KeyPress")
      
End Sub

Private Sub txt1_LostFocus(Index As Integer)
    
        On Error GoTo Fehler

100     txt1(Index).ForeColor = vbActiveTitleBar
105     txt1(Index).BackColor = vbWindowBackground  'Fensterhintergrund(weiß)
        
        Exit Sub

Fehler:
110     Call FehlerErklärung("frmSP52830", "txt1_LostFocus")
        
End Sub

Public Sub MaskeLeeren(blnSaveMCode As Boolean)

        On Error GoTo Fehler
        
        Dim i  As Integer
        
        Dim j  As Integer
        
        Dim zk As ZahlungsKonditionen
        
100     gbDataChanged = False
        
105     j = 0
        
110     If blnSaveMCode Then
            
115         j = 1
        
        End If
        
120     For i = j To txt1.Count - 1

125         Select Case i

                Case 19, 20, 21, 22, 23, 24                                     'Nummerischen Felder
                
130                 objPRM.FindFirstString = "name = 'txt1' AND index = " & i
135                 txt1(i) = objPRM.EingabeUmwandlung(0)

140             Case 40                                                         'HW 16.07.2012 muss gemacht werden damit Steuertext enthalten bleibt!

145             Case 42                                                         'HW, 10.01.2018, Dezimalstellen dürfen nicht geleert werden!
                
165             Case Else

170                 txt1(i) = ""

            End Select

        Next
        
175     KursFuerRechnung "000"
 
        'DF 10.06.2020 , Ver.: 6.6.102 : Mandanten-Währung nicht mehr in Beleg-Währung übernehmen,
        '                                da diese jetz Beleg und Vergleichswährung aus Kundenstamm sind
        '150     txt1(17) = txt1(18) 'Rechnungswährung
  
180     Check1(0).value = 0
185     Check1(1).value = 0
190     Check1(2).value = 0                                                     'MCode ändern
195     Check1(2).Visible = False

200     Frames True

205     gintZwAblage = 0
210     gintDruck = 0
215     glngBelegID = 0
220     glngBelegIDTmp = 0                                                      'HW 03.12.2015 Ver.: 6.4.113 ID fue die Vorschau Leeren!

225     sta1.Panels(3).text = ""

        'lbl2(7) = ""                                                           'IL 27.08.2024

230     gbDataChanged = True
  
235     EWerkErmitteln

240     objDruckOptionen.clearVars                                              'DH, 11.07.2013, Beim Leeren auch die fuer den DruckOptionen Dialog wichtigen Variablen leeren

245     belegDatum = ""
250     BelegNr = ""
255     ValutaDatum = ""
260     printDone = False
265     objDruckOptionen.EnableBelegDatum = True
270     objDruckOptionen.EnableBelegNr = True
275     blnBelegNeu = False
280     gEnmBelegWrgArt = BelegSchlVorhanden                                    'DF 22.06.2020 , Ver.: 6.6.102
285     gEnmKudnenERechnungType = eERechnungType.None                           'DF 23.07.2024 , Ver.: 6.7.101
290     objERechnung.Clear                                                      'DF 29.08.2024 , Ver.: 6.7.101
295     modERechnung.dictRowsNums.RemoveAll

300     Call ClearBelegInfo                                                     'IL 22.10.2024 , Ver.: 6.7.101
        
        'ZAHLUNGSKONDITIONEN DES @SYSTEM-STAMMSATZES
305     zk = GetZahlungsKonditionen(C_STR_SYSTEM_STAMMSATZ_MCODE)               'DF 18.11.2024 , Ver.: 6.7.101 : Beim Leeren die Zahlungskonditionen aus dem System-Stammsatz laden, sowie z.B. E-Beleg Einstellung.
        
310     txt1(21).text = zk.SkontoSatz
315     txt1(22).text = zk.SkontoTage
320     txt1(23).text = zk.SkontoTageNetto
        
        '<Modified by: IL at 22.10.2024, Ver.: 6.7.101 >
        '# Für die Dokumenttypen „Angebot“ und „Auftragdbästetigung“ bleibt die Variable leer, da die E-Rechnung nicht verwendet wird
        Dim tempBlnSetGlobal As Boolean
        
325     If GintBelegArt <> 2 And GintBelegArt <> 3 Then

330         tempBlnSetGlobal = True
335         lbl1(7).Enabled = True
340         lbl2(7).Enabled = True

        Else

345         lbl1(7).Enabled = False
350         lbl2(7).Enabled = False

        End If

355     lbl2(7).caption = modERechnung.GetKundenERechnungTypeAsText(C_STR_SYSTEM_STAMMSATZ_MCODE, 1, tempBlnSetGlobal)
        '</Modified by: IL at 22.10.2024, Ver.: 6.7.101 >

360     Me.caption = objPRM.getColCaption("name = 'frmSP52830' AND index = " & GintBelegArt)
        
365     Call ResetUIDValidationInfo                                             'DF 12.19.2023 , Ver.: 6.6.124
        
        Exit Sub

Fehler:
370     Call FehlerErklärung("frmSP52830", "MaskeLeeren()")
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)

        On Error GoTo Fehler
        
100     If gbDataChanged Then

105         objPRM.FindFirstString = "name = 'txt1' AND index = " & Index

110         Select Case Index
                
                Case 8, 14, 34                                                  'LKZ, UID, KTO-ART
                    
                    'Bei Änderungen dieser Felder muss die UID-Überprüfung erneut angestoßen werden.
                    
115                 Call ResetUIDValidationInfo                                 'DF 12.19.2023 , Ver.: 6.6.124
                    
120                 txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , True, True)
                    
125             Case 17                                                         'Währung

130                 If Trim(txt1(Index)) <> "" Then
                        
135                     If objPlausi.RSOpen("SELECT Schl FROM [1100_Währungen] WHERE Schl = '" & Trim$(txt1(Index).text) & "'", True) = False Then 'DF 11.06.2020 , Ver.: 6.6.102 : Statt ISO wird Schl geprüft -> ISO -> Schl (UST Senkung Umstellung)

140                         Cancel = True

                        End If
                        
                    End If

145             Case 25                                                         'Kostenschlüssel

150                 If GbKostenstellenPflicht Or Trim(txt1(Index)) <> "" Then

155                     If objPlausi.RSOpen("SELECT Schl FROM [1100_FiBuKostenStellen] WHERE Schl = '" & txt1(Index) & "'", True) = False Then
160                         Cancel = True

                            Exit Sub

                        End If
                        
                    End If

165             Case 26                                                         'Sachkonten Schluessel

170                 If Trim(txt1(Index).text) <> "" Then

175                     If objPlausi.RSOpen("SELECT Schl FROM [1100_FiBuSachkonten] WHERE Schl = '" & txt1(Index) & "'", True) = False Then

180                         Cancel = True

                            Exit Sub

                        End If
                        
                    End If

185             Case 29, 37, 38, 44                                             'Abfertigungs-Datum, Zeitraum, Lieferdatum

190                 txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index))

                    Dim erg As Boolean

195                 erg = DatumUnterschiedRek(2, txt1(Index), txt1(37), 90, "d") 'GW_05.03.2018, Ver. 6.5.106 : Fehlermeldung schmeissen, wenn Datum-Bis kleiner als Datum-Von ist

200                 If erg = True Then

205                     txt1(Index) = ""

210                     If txt1(Index).Enabled Then txt1(Index).SetFocus

215                     Cancel = True

                    Else
                    
220                     Cancel = False

                    End If

225             Case 39
                    
230                 objSQLAuswDef.GetIfOnesHit = True
                    
235                 cmdAuswahl_Click 1                                          'HW 11.03.2014
                    
240                 objSQLAuswDef.GetIfOnesHit = False
                    
245                 If txt1(Index).text = "" Then

250                     Cancel = True

                        Exit Sub

                    End If

255             Case 42                                                         'Dezimal-Stellen

260                 If IsNumeric(txt1(Index)) Then                              'HW, 10.01.2018, Überprüfung auf gültige Eingabe

                        Dim intA As Integer

265                     intA = CInt(txt1(Index))

270                     If intA >= 2 And intA < 6 Then

275                         postCommaPreis = intA

                        Else
                        
280                         MsgBox GetMessage(203), vbExclamation + vbOKOnly, strMeldungCap
285                         Cancel = True

                        End If

                    Else
                    
                    End If

290             Case Else

295                 txt1(Index) = objPRM.EingabeUmwandlung(txt1(Index), , True, True)

300                 If Index = 0 Then EWerkErmitteln

            End Select

305         If objPRM.EingabeFehler(txt1(Index)) Then

310             Cancel = True

            End If
            
        End If

        Exit Sub

Fehler:
315     Call FehlerErklärung("frmSP52830", "txt1_Validate()")

End Sub

Public Function KundeZeigen(Optional blnForAbwRE As Boolean = False) As Boolean

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       KundeZeigen
        ' Description:
        ' Created by :       DFiebach
        ' Date-Time  :       14.11.2024-13:41:14
        '
        ' Parameters :       blnForAbwRE (Boolean = False) - Optionaler Parameter um zu zeigen,
        '                                                 dass Stammsatz des Abw.RE geladen wird (für REKURSION)
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler

        Dim rs       As ADODB.Recordset

        Dim sql      As String

        Dim WrgSchl  As String

        Dim Limit    As Boolean

        Dim Lim      As Double

        Dim Betr     As Double

        Dim MText    As String

        Dim Merker   As String
        
        Dim blnAbwRE As Boolean                                                 'DF 16.06.2020 , Ver.: 6.6.102 : Ust Senkung Umstellung
        
        Dim strMCode As String

100     If Check1(2).value = 0 Then MaskeLeeren (True)
        
        'CSBmk <KUNDEN-INFO>
105     Call ShowKundenInfo(Trim(txt1(0)))                                      'DF 30.03.2022 , Ver.: 6.6.113 : Kunden-Info Text
        
        'CSBmk <BONITÄTSLIMIT ÜBERPRÜFEN>
110     objLimit.MCode = Trim(txt1(0))
115     Lim = objLimit.Limit

120     If Lim > 0 Then

125         Betr = objLimit.GesamtBetrag

130         If Lim < Betr Then

135             Protokoll iAppend, "Limit für " & Trim(txt1(0)) & " erreicht. Limit: " & Lim & " Gesamt-Betrag: " & Betr

140             Limit = True

145             MText = "A C H T U N G" & vbCrLf & "Das Kunden-Limit für Kunde: '" & Trim(txt1(0)) & "' ist überschritten !" & vbCrLf
150             MText = MText & "Limit" & vbTab & ": " & Format(Lim, "###,###,##0.00") & vbCrLf
155             MText = MText & "Aktuell" & vbTab & ": " & Format(Betr, "###,###,##0.00") & vbCrLf & vbCrLf
160             MText = MText & "Soll die aktuelle Erfassung abgebrochen werden ?"

165             If MsgBox(MText, vbCritical + vbYesNo + vbDefaultButton1, strMeldungCap) = vbYes Then

                    Exit Function

                End If
                
            End If
            
        End If
 
        'CSBmk <ABW. RECHNUNGS-EMPFÄNGER>
        
170     OPEN_gConn
        
175     Set rs = New ADODB.Recordset
180     rs.Open "SELECT FrMcode FROM [1200_AllgDaten] WHERE [MCode] = '" & txt1(0).text & "'", gConn, adOpenStatic, adLockReadOnly
        
185     If Trim(rs!FrMCode) <> "" Then

190         Call msgText(1, 2356, 0, 0, 0)
195         GsMsgText(0) = Replace(GsMsgText(0), "%1", txt1(0))
200         GsMsgText(0) = Replace(GsMsgText(0), "%2", rs!FrMCode)

205         If MsgBox(GsMsgText(0), vbYesNo + vbQuestion, strMeldungCap) = vbYes Then

210             Merker = txt1(0)

215             txt1(0) = rs!FrMCode

220             blnAbwRE = True
                
            Else
            
225             GstrAuftraggeber = ""

            End If

        Else
            
230        If blnForAbwRE = False Then GstrAuftraggeber = ""                   'DF 14.11.2024 , Ver.: 6.7.101 : Nur leeren wenn nicht Aufruf für den Abw.RE

        End If
        
        '<Added by: DFiebach at: 14.11.2024, Ver.: 6.7.101 >
        ' # Neue Logik der Übernahme des Abw.RE. M-Code muss mitübernommen werden, damit anschließende Übergabe an E-Beleg, OP, FIBu auch an den Abw.RE passiert.
        ' # ACHTUNG REKURSION!!!!
235     If blnAbwRE Then
            
240         Call KundeZeigen(True)
            
245         KundeZeigen = True
            
            Exit Function
            
        End If

        '</Added by: DFiebach at: 14.11.2024, Ver.: 6.7.101 >
        
250     rs.Close

255     rs.Open "SELECT MCode, Name1, Name2, Postfach, Plz1, Ort1, Straße, Lkz, Plz, Ort, Ortsteil, KtoNr, KtoKnz, SteuerNr, UID FROM [1200_AllgDaten] WHERE MCode= '" & Trim(txt1(0).text) & "'", gConn, adOpenStatic, adLockReadOnly

260     If rs.RecordCount > 0 Then

265         txt1(1) = "" & rs!Name1
270         txt1(2) = "" & rs!Name2
275         txt1(4) = "" & rs!Postfach
280         txt1(5) = "" & rs!PLZ1
285         txt1(6) = "" & rs!ort1
290         txt1(7) = "" & rs!Straße
295         txt1(8) = "" & rs!Lkz
300         txt1(9) = "" & rs!Plz
305         txt1(10) = "" & rs!Ort
310         txt1(11) = "" & rs!ORTSTEIL
315         txt1(12) = "" & rs!KtoNr
320         txt1(34) = "" & rs!KtoKnz
325         txt1(13) = "" & rs!SteuerNr
330         txt1(14) = "" & rs!Uid
            
        End If

340     rs.Close
  
        'CSBmk <ZAHLUNGSKONDITIONEN>
345     rs.Open "SELECT Lkz, Ust, ZSkto, ZSktoTage, Ztage, PPRechnung FROM [1200_GrundKonditionen] WHERE MCode= '" & txt1(0).text & "'", gConn, adOpenStatic, adLockReadOnly

350     If rs.RecordCount > 0 Then

355         objPRM.FindFirstString = "name = 'txt1' AND index = 21"
360         txt1(21) = objPRM.EingabeUmwandlung(rs!ZSkto)

365         objPRM.FindFirstString = "name = 'txt1' AND index = 22"
370         txt1(22) = objPRM.EingabeUmwandlung(rs!ZSktoTage)

375         objPRM.FindFirstString = "name = 'txt1' AND index = 23"
380         txt1(23) = objPRM.EingabeUmwandlung(rs!ZTage)

385         objPRM.FindFirstString = "name = 'txt1' AND index = 24"
390         txt1(24) = objPRM.EingabeUmwandlung(rs!PPRechnung)

395         SteuerTyp (rs!Ust)

400         WrgSchl = rs!Lkz
            
        End If

405     rs.Close
        
        'DF 14.11.2024 , Ver.: 6.7.101 : Eigentlich wird die Wrg ab jetzt immer vom RE genommen , da im Falle des Abw.RE die Funktiuon rekursiv aufgerufen wird, und an dieser Stelle steht halt ab jetzt immer
        '                                der "tatsächlicher" RE egal ob abweichender oder nicht. Also kann künftig die Abfrage If blnAbwRE = False weggelassen werden.
410     If blnAbwRE = False Then KursFuerRechnung WrgSchl                      'DF 15.06.2020 , Ver.: 6.6.102 : Währung darf nur vom Rechnungs-Empfänger und nicht vom Abw.Rechnungs-Empfänger genommen werden.

415     If mnuOpt1(1).Checked Then

            'DH, 08.11.2012, 6.1.119, Logik angepasst, welcher Datensatz gezogen werden soll und wann das Feld einfach leer bleibt
420         rs.Open "SELECT Anrede,Name1,Name2 FROM [1200_AnsprPartner] WHERE MCode='" & txt1(0) & "' AND Standard = 1", gConn, adOpenStatic, adLockReadOnly

425         If rs.RecordCount > 0 Then
430             txt1(3) = rs!Anrede & " " & rs!Name1 & " " & rs!Name2
            Else
435             txt1(3) = ""
            End If
            
        End If
  
440     If rs.state = adStateOpen Then rs.Close
  
445     txt1(29) = "" 'GdtDatum
  
450     EWerkErmitteln
        
        'CSBmk <E-RECHNUNG>
        
        '<Modified by: DFiebach at 18.10.2024, Ver.: 6.7.101 >
        'lbl2(7).caption = modERechnung.ERechnungZeigen(txt1(0).text)           'IL 27.08.2024
        'objERechnung.BelegEmpfaengerMCode = txt1(0).text                       'DF 29.08.2024 , Ver.: 6.7.101
        
        '<Modified by: IL at 22.10.2024, Ver.: 6.7.101 >
        '# Für die Dokumenttypen „Angebot“ und „Auftragdbästetigung“ bleibt die Variable leer, da die E-Rechnung nicht verwendet wird
        Dim tempBlnSetGlobal As Boolean
        
455     If GintBelegArt <> 2 And GintBelegArt <> 3 Then

460         tempBlnSetGlobal = True
465         lbl1(7).Enabled = True
470         lbl2(7).Enabled = True

        Else

475         lbl1(7).Enabled = False
480         lbl2(7).Enabled = False

        End If

485     lbl2(7).caption = modERechnung.GetKundenERechnungTypeAsText(txt1(0).text, 1, tempBlnSetGlobal)
        '</Modified by: DFiebach at 18.10.2024, Ver.: 6.7.101 >
        
        '</Modified by: IL at 22.10.2024, Ver.: 6.7.101 >

490     modERechnung.dictRowsNums.RemoveAll                                    'IL 02.09.2024

495     Call ClearBelegInfo                                                    'IL 22.10.2024 , Ver.: 6.7.101
        
500     If Merker <> "" Then txt1(0) = Merker
  
505     KundeZeigen = True

        Exit Function

Fehler:
510     Call FehlerErklärung("frmSP52830", "KundeZeigen()")
End Function

Public Sub BelegKopfZeigen(BelegID As Long)

        On Error GoTo Fehler

        Dim rs  As ADODB.Recordset

        Dim sql As String
  
100     MaskeLeeren (True)

105     txt1(0) = ""
  
110     Set rs = New ADODB.Recordset
  
115     rs.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = " & BelegID, gConn, adOpenStatic, adLockReadOnly

120     If rs.RecordCount > 0 Then

125         txt1(0) = "" & rs!MCode
130         txt1(1) = "" & rs!Name1
135         txt1(2) = "" & rs!Name2
140         txt1(3) = "" & rs!AnsprPartner
145         txt1(4) = "" & rs!Postfach
150         txt1(5) = "" & rs!PLZ1
155         txt1(6) = "" & rs!ort1
160         txt1(7) = "" & rs!Straße
165         txt1(8) = "" & rs!Lkz
170         txt1(9) = "" & rs!Plz
175         txt1(10) = "" & rs!Ort
180         txt1(11) = "" & rs!ORTSTEIL
185         txt1(12) = "" & rs!KtoNr
190         txt1(34) = "" & rs!KtoKnz
195         txt1(13) = "" & rs!SteuerNr
200         txt1(14) = "" & rs!Uid
205         txt1(41) = "" & rs!InternerVermerk                                 'DF 14.01.2015 : Interner Vermerk hinzugefügt.
    
210         SteuerTyp (rs!Ust)

215         Check1(0).value = 0 'rs!ZwAblage

220         gintZwAblage = rs!ZwAblage

225         If rs!ZwAblage = 0 Then

                'DH, 13.03.2013, BelegNr und -Datum des geladenen Belegs jetzt nicht mehr in die Textfelder sonder in die lokale Variable schreiben
                
                'BelegNr = "" & rs!BelegNr                                     'IL 21.10.2024 , Ver.: 6.7.101 :

230             belegDatum = "" & rs!belegDatum
235             ValutaDatum = "" & rs!ValutaDatum                              'DH, 17.02.2015, 6.4.103, ValutaDatum auch aus dem geladenen Beleg holen

240             glngBelegIDVorlage = 0                                         'DH, 01.06.2017, 6.4.126, Auf Null setzen, wenn es sich nicht um eine Vorlage handelt

            Else
            
245             txt1(15) = ""
250             txt1(16) = ""

255             glngBelegIDVorlage = BelegID                                   'DH, 01.06.2017, 6.4.126, Bei einer Vorlage die urspruengliche BelegID merken, damit das Stornieren funktioniert

            End If
            
            'DF 16.06.2020 , Ver.: 6.6.102 : Ust Senkung
            '265         lbl2(17).caption = "" & rs!Wrg1
            '270         lbl2(18).caption = "" & rs!Wrg2
            '
            '275         txt1(17).text = "" & rs!WrgSchl
            '280         txt1(18).text = "" & rs!WrgSchlVgl
            
260         Check1(1).value = rs!WrgAusw

265         objPRM.FindFirstString = "name = 'txt1' AND index = 19"
270         txt1(19) = objPRM.EingabeUmwandlung(rs!Kurs)
275         objPRM.FindFirstString = "name = 'txt1' AND index = 20"
280         txt1(20) = objPRM.EingabeUmwandlung(rs!MwSt)
285         objPRM.FindFirstString = "name = 'txt1' AND index = 21"
290         txt1(21) = objPRM.EingabeUmwandlung(rs!ZSkto)
295         objPRM.FindFirstString = "name = 'txt1' AND index = 22"
300         txt1(22) = objPRM.EingabeUmwandlung(rs!ZSktoTage)
305         objPRM.FindFirstString = "name = 'txt1' AND index = 23"
310         txt1(23) = objPRM.EingabeUmwandlung(rs!ZTage)
315         objPRM.FindFirstString = "name = 'txt1' AND index = 24"
320         txt1(24) = objPRM.EingabeUmwandlung(rs!PPRechnung)

325         txt1(25) = "" & rs!KostSchl
330         txt1(26) = "" & rs!FiBuSchl
335         txt1(27) = "" & rs!KostKonto
340         txt1(28) = "" & rs!FibuKonto
345         txt1(29) = "" & rs!Datum
350         txt1(30) = "" & rs!AbfZus
355         txt1(31) = "" & rs!AbfPos
360         txt1(35) = "" & rs!ErfNr
365         txt1(32) = "" & rs!Relation
370         txt1(33) = "" & rs!KfzSchl
375         txt1(36) = "" & rs!KostenArt
380         txt1(37) = "" & rs!vonDatum
385         txt1(38) = "" & rs!bisDatum
            
400         glngBelegID = BelegID

405         gintDruck = rs!Druck

            'DF 11.06.2020 , Ver.: 6.6.102 : Ust Senkung Umstellung txt1(17/18) auf lbl2(17/18) umgestellt
            '                                Währungs-Schl(Beleg- und Vergleichs werden jetzt in txt1(17/18) angezeigt)
            
            '# INFO!!!###
            '
            'Währung-Schlüssel fehlt -> alte Logik (es wurden nur Wrg-ISO und Wrg-Prozentsatz im Beleg-Datensatz gespeichert) -> lade Wrg-ISO und VglWrg-ISO
            '                           NUR bei bereits gedruckten Belegen und NICHT Vorlagen
            '                           Bei Gespeicherten und Vorlagen, wird neue Logik angewendet, der fehlende Wrg-Schl wird dann ducht <Plausi> gemeldet.
            '
            '############
410         If Trim$(rs.Fields("WrgSchl").value) = "" And gintDruck = 1 And gintZwAblage = 0 Then

415             lbl2(17).caption = ""
420             lbl2(18).caption = ""

425             txt1(17).text = "" & rs!Wrg1
430             txt1(18).text = "" & rs!Wrg2

435             gEnmBelegWrgArt = BelegSchlFehlt

            Else
                
440             lbl2(17).caption = "" & rs!Wrg1
445             lbl2(18).caption = "" & rs!Wrg2

450             txt1(17).text = "" & rs!WrgSchl
455             txt1(18).text = "" & rs!WrgSchlVgl

460             gEnmBelegWrgArt = BelegSchlVorhanden
                
            End If
            
            'If (gintDruck = 0 Or gintZwAblage = 1) And Trim$(txt1(0).text) <> "" Then                           'DF 03.09.2024 , Ver.: 6.7.101 : wenn Beleg gespeichert oder Vorlage -> hole ERechnung Einstellung neu aus dem Stammsatz
            'IL 17.09.2024   ORIG:  gintDruck = 0 Or gintZwAblage = 1
                
            '<Modified by: DFiebach at 18.10.2024, Ver.: 6.7.101 >
            'lbl2(7).caption = modERechnung.ERechnungZeigen(txt1(0).text)
            'objERechnung.BelegEmpfaengerMCode = txt1(0).text
                
            '<Modified by: IL at 22.10.2024, Ver.: 6.7.101 >
            '# Für die Dokumenttypen „Angebot“ und „Auftragdbästetigung“ bleibt die Variable leer, da die E-Rechnung nicht verwendet wird
            Dim tempBlnSetGlobal As Boolean
        
465         If GintBelegArt <> 2 And GintBelegArt <> 3 Then

470             tempBlnSetGlobal = True
475             lbl1(7).Enabled = True
480             lbl2(7).Enabled = True

            Else

485             lbl1(7).Enabled = False
490             lbl2(7).Enabled = False

            End If

495         lbl2(7).caption = modERechnung.GetKundenERechnungTypeAsText(txt1(0).text, 1, tempBlnSetGlobal)

            '</Modified by: IL at 22.10.2024, Ver.: 6.7.101 >

            '</Modified by: DFiebach at 18.10.2024, Ver.: 6.7.101 >
               
            'End If
            
500         objPRM.FindFirstString = "name = 'statusBarText' AND index = 2"     'DF 16.01.2015 : StatusBAr Text aus PRM holen.
505         sta1.Panels(3).text = objPRM.caption("Erfassungs-Nr.:") & " " & glngBelegID
510         objPRM.FindFirstString = ""
    
        End If

515     blnBelegNeu = False                                                     'DF 16.01.2015 : ein alter Beleg wurde geladen

520     If rs.state = adStateOpen Then rs.Close
525     Set rs = Nothing
  
530     EWerkErmitteln
  
        Exit Sub

Fehler:
535     Call FehlerErklärung("frmSP52830", "BelegKopfZeigen()")
End Sub

Public Sub KursFuerRechnung(ByVal WrgSchl As String, _
                            Optional ByVal KunenISO As String, _
                            Optional ByVal MandantenISO As String)

        On Error GoTo Fehler

        '<Removed by: DFiebach at: 10.06.2020, Ver.: 6.6.102 >
        '
        '        Dim rsWrg As ADODB.Recordset
        '
        '        Dim sql   As String
        '
        '        Dim Kurs  As Double
        '
        '100     Kurs = 1
        '
        '105     Set rsWrg = New ADODB.Recordset
        '
        '        'Kunden Währung
        
        '110     OPEN_gConn
        '115     rsWrg.Open "SELECT Schl,Kurs,MwSt,ISO FROM [1100_Währungen]", gConn, adOpenStatic, adLockReadOnly
        '
        '120     If rsWrg.RecordCount > 0 Then
        '125         If KunenISO <> "" Then
        '130             rsWrg.Find "ISO= '" & KunenISO & "'"
        '            Else
        '135             rsWrg.Find "Schl= '" & WrgSchl & "'"
        '            End If
        '
        '140         If Not rsWrg.EOF And Not rsWrg.BOF Then
        '145             txt1(17) = rsWrg!Iso
        '
        '150             If rsWrg!Kurs <> 0 Then
        '155                 Kurs = rsWrg!Kurs
        '                End If
        '
        '            Else
        '160             txt1(17) = ""
        '            End If
        '        End If
        '
        '        'Mandanten Währung
        
        '165     If MandantenISO <> "" Then
        
        '170         If rsWrg.RecordCount > 0 Then
        
        '175             rsWrg.Find "ISO='" & MandantenISO & "'"
        '
        '180             If Not rsWrg.EOF And Not rsWrg.BOF Then
        '185                 If rsWrg!Kurs <> 0 Then
        '190                     objPRM.FindFirstString = "name = 'txt1' AND index = 19"
        '195                     txt1(19) = objPRM.EingabeUmwandlung(CStr(Runden(rsWrg!Kurs / Kurs, 6)))
        '                    End If
        '                End If
        
        '            End If
        '
        '        Else
        '
        '200         If GmandantRS.RecordCount > 0 Then
        
        '205             If rsWrg.RecordCount > 0 Then
        
        '210                 rsWrg.MoveFirst
        '215                 rsWrg.Find "Schl='" & GmandantRS!EigenWrg & "'"
        '
        '220                 If Not rsWrg.EOF And Not rsWrg.BOF Then
        
        '225                     If txt1(17) = rsWrg!Iso Then
        
        '                            'Mandanten- und Kunden-Währung sind gleich.
        '                            'txt1(18) = ""
        
        '230                         objPRM.FindFirstString = "name = 'txt1' AND index = 19"
        '235                         txt1(19) = objPRM.EingabeUmwandlung("1")
        
        '                            'objPRM.FindFirstString = "name = 'txt1' AND index = 20"
        '                            'txt1(20) = objPRM.EingabeUmwandlung(RsWrg!MwSt)
        
        '                        Else
        
        '240                         objPRM.FindFirstString = "name = 'txt1' AND index = 19"
        
        '                            'Hauptwährung hat immer die Basis = 1. Deshalb wird der Kurs = MandantenKurs/RechnungsKurs
        '245                         txt1(19) = objPRM.EingabeUmwandlung(CStr(Runden(rsWrg!Kurs / Kurs, 6)))
        
        '                        End If
        '
        '250                     txt1(18) = rsWrg!Iso
        '255                     objPRM.FindFirstString = "name = 'txt1' AND index = 20"
        '260                     txt1(20) = objPRM.EingabeUmwandlung(rsWrg!MwSt)
        
        '                    Else
        
        '265                     txt1(18) = ""
        
        '270                     objPRM.FindFirstString = "name = 'txt1' AND index = 19"
        '275                     txt1(19) = objPRM.EingabeUmwandlung("0")
        
        '280                     objPRM.FindFirstString = "name = 'txt1' AND index = 20"
        '285                     txt1(20) = objPRM.EingabeUmwandlung("0")
        
        '                    End If
        
        '                End If
        '
        '            Else
        
        '290             txt1(18) = ""
        
        '295             objPRM.FindFirstString = "name = 'txt1' AND index = 19"
        '300             txt1(19) = objPRM.EingabeUmwandlung("0")
        
        '305             objPRM.FindFirstString = "name = 'txt1' AND index = 20"
        '310             txt1(20) = objPRM.EingabeUmwandlung("0")
        
        '            End If
        '        End If
        '
        '</Removed by: DFiebach at: 10.06.2020, Ver.: 6.6.102 >
        
        Dim oKWaehrung As KundenWaehrung
        
100     oKWaehrung = GetKundenWaehrung(txt1(0).text, True)
102     oKWaehrung.Kurs = 1

        '110     If oKWaehrung.KursVgl = 0 Then
        '115         oKWaehrung.KursVgl = oKWaehrung.Kurs
        '        End If
        
108     If oKWaehrung.KndWrg <> "" Then
           
            'Beleg-Währung
110         txt1(17).text = oKWaehrung.WrgSchl
112         lbl2(17).caption = oKWaehrung.ISO
        
            'Vergleichs-Währung
114         txt1(18).text = oKWaehrung.WrgSchlVgl
116         lbl2(18).caption = oKWaehrung.ISOVgl
            
            'Kurs
118         objPRM.FindFirstString = "name = 'txt1' AND index = 19"
120         txt1(19) = objPRM.EingabeUmwandlung(CStr(Runden(oKWaehrung.KursVgl, 6)))
            
            '<Removed by: DFiebach at: 17.06.2020, Ver.: 6.6.102 >
            '
            ' # Kurs darf nicht mehr umgerecnet werden. sonst wird immer von der Vergleichswährung übernommen
            '
            '150         If oKWaehrung.ISO = oKWaehrung.ISOVgl Then
            '
            '                'Beleg- und Vergleichs-Währung sind gleich
            '
            '                'Kurs
            '155             objPRM.FindFirstString = "name = 'txt1' AND index = 19"
            '
            '160             txt1(19).text = objPRM.EingabeUmwandlung("1")
            '
            '            Else
            '
            '                'Kurs
            '165             objPRM.FindFirstString = "name = 'txt1' AND index = 19"
            '
            '170             If oKWaehrung.KursVgl <> 0 Then
            '
            '                    'Hauptwährung hat immer die Basis = 1. Deshalb wird der Kurs = Vergl.Währung Kurs/RechnungsKurs
            '175                 txt1(19) = objPRM.EingabeUmwandlung(CStr(Runden(oKWaehrung.KursVgl / oKWaehrung.Kurs, 6)))
            '
            '                Else
            '
            '                End If
            '
            '            End If
            '
            '</Removed by: DFiebach at: 17.06.2020, Ver.: 6.6.102 >

            'Steuer-Satz (der Beleg-Währung)
122         objPRM.FindFirstString = "name = 'txt1' AND index = 20"
124         txt1(20).text = objPRM.EingabeUmwandlung(CStr(oKWaehrung.MwSt))
           
        Else
        
            'Beleg-Währung
126         txt1(17).text = ""
128         lbl2(17).caption = ""
        
            'Vergleichs-Währung
130         txt1(18).text = ""
132         lbl2(18).caption = ""
            
            'Kurs
134         objPRM.FindFirstString = "name = 'txt1' AND index = 19"
136         txt1(19) = objPRM.EingabeUmwandlung("0")
            
            'Steuer-Satz (der Beleg-Währung)
138         objPRM.FindFirstString = "name = 'lbl2' AND index = 20"
140         txt1(20).text = objPRM.EingabeUmwandlung("0")
            
        End If
        
142     If IsNumeric(oKWaehrung.MwStSchl) Then
        
144         Call SteuerTyp(CInt(oKWaehrung.MwStSchl))
        
        End If
        
146     objPRM.FindFirstString = "name = 'txt1' AND index = 19"
        
        Exit Sub

Fehler:
148     Call FehlerErklärung("frmSP52830", "KursFuerRechnung()")
End Sub

Public Function Plausi() As Integer

        On Error GoTo Fehler

        Dim i      As Integer
        
        Dim strSQL As String
        
100     Plausi = 999
  
105     gbDataChanged = False
  
110     For i = 0 To txt1.Count - 1
                              
115         If txt1(i).Visible Then 'Beleg -Nr und -Datum werden nur erfasst, wenn manuelle Erfassung eingestellt ist.

120             Select Case i
                    
                    Case 0  'M-Code
                        
125                     If Trim(txt1(i)) <> "" Then                            'DF 23.05.2023 , Ver.: 6.6.120 : M-Code überprüfen
                            
135                         If objPlausi.RSOpen("SELECT MCode FROM [1200_AllgDaten] WHERE MCode = '" & Trim$(txt1(i).text) & "'", True) = False Then

140                             Plausi = i

                                Exit For

                            End If
                            
                        End If

145                 Case 17 'Währung

150                     If Trim(txt1(i)) = "" Then

155                         Plausi = i

                            Exit For

                        Else
                        
                            'DF 11.06.2020 , Ver.: 6.6.102 : Statt ISO wird Schl geprüft -> ISO -> Schl (UST Senkung Umstellung)
                            '                                Bei alten Belegen (vor der Umstellung erfasst) wird ISO in den TextFeld geladen (wie bei alter Logik) und anschließend geprüft.
160                         Select Case gEnmBelegWrgArt
                            
                                Case BelegWaehrungArt.BelegSchlVorhanden
                                
165                                 strSQL = "SELECT Schl FROM [1100_Währungen] WHERE Schl = '" & Trim$(txt1(i).text) & "'"
                                
170                             Case BelegWaehrungArt.BelegSchlFehlt
                                     
175                                 strSQL = "SELECT ISO FROM [1100_Währungen] WHERE ISO = '" & Trim$(txt1(i).text) & "'"
                                     
                            End Select
                            
180                         If objPlausi.RSOpen(strSQL, True) = False Then

185                             Plausi = i

                                Exit For

                            End If
                            
                        End If

190                 Case 25 'Kostenschlüssel

195                     If GbKostenstellenPflicht Then

200                         If Trim(txt1(i)) = "" Then

205                             Plausi = i

210                             Call msgText(2, 203, 285, 0, 0)
215                             MsgBox GsMsgText(0) & vbCrLf & GsMsgText(1), vbExclamation, strMeldungCap 'Eingabe ist unvollständig oder falsch! Bitte betätigen Sie F2-Taste oder die Auswahlschaltfläche um eine erlaubte Eintragung auszuwählen.

                                Exit For

                            End If
                            
                        End If

220                     If Trim(txt1(i)) <> "" Then

225                         If objPlausi.RSOpen("SELECT Schl FROM [1100_FiBuKostenStellen] WHERE Schl = '" & txt1(i) & "'", True) = False Then

230                             Plausi = i

                                Exit For

                            End If
                            
                        End If

235                 Case 26 'Sachkonten Schluessel

240                     If Trim(txt1(i)) <> "" Then

245                         If objPlausi.RSOpen("SELECT Schl FROM [1100_FiBuSachkonten] WHERE Schl = '" & txt1(i) & "'", True) = False Then

250                             Plausi = i

                                Exit For

                            End If
                            
                        End If
                        
255                 Case Else

260                     If Trim(txt1(i)) = "" Then

265                         objPRM.FindFirstString = "name = 'txt1' AND index = " & i

270                         If objPRM.EingabePflicht Then

275                             Plausi = i

                                Exit For

                            End If
                            
                        End If

                End Select

            End If

        Next
  
280     If IsDate(txt1(29)) Then

            'KostenArt muss gesetzt sein, wenn die Daten im Speditionsbuch gespeichert werden sollen.
285         If Trim(txt1(36)) = "" Then Plausi = 36

        End If
  
290     gbDataChanged = True
  
        Exit Function

Fehler:
295     Call FehlerErklärung("frmSP52830", "Plausi()")
End Function

Public Sub BelegSuchen(eBelegTyp As E_DATATYPE, Druck As Integer)

        'Druck = 0 -> Ungedruckten
        'Druck = 1 -> Gedruckten
        'Druck = 2 -> Vorlagen
  
        On Error GoTo Fehler

        Dim i           As Integer

        Dim ColLeft     As Long

        Dim RowTop      As Long

        Dim OT          As String
  
        Dim rc          As rect
        
        Dim intBelegArt As Integer

100     Call GetWindowRect(Frame1(0).hwnd, rc)
        '<Modified by: IL at 18.10.2024, Ver.: 6.7.101 >
        '# Mischen die Auswahltabelle näher zur Mitte
        
105     ColLeft = (rc.left * Screen.TwipsPerPixelX - 70) + lbl1(5).left
110     RowTop = (rc.top * Screen.TwipsPerPixelY - 340) + lbl1(5).top + lbl1(5).height

        '       ORIG:
        '       ColLeft = (rc.Left * Screen.TwipsPerPixelX - 70) + lbl1(5).Left
        '       RowTop = (rc.Top * Screen.TwipsPerPixelY - 340)
        
        '</Modified by: IL at 18.10.2024, Ver.: 6.7.101 >
  
        'DF 19.01.15 : aktuelle Resize-Stufe des Hauptfensters übernehmen
115     objSQLAusw.ScaleFactorHeight = cReSize.CurrScaleFactorHeight
120     objSQLAusw.ScaleFactorWidth = cReSize.CurrScaleFactorWidth
125     objSQLAusw.fontSize = Me.fontSize

130     objSQLAusw.FilterBar = True
135     objSQLAusw.BorderStyle = 4
        '100     objSQLAusw.caption = "  " & Replace(cmd1(5).caption, "&", "") & "-Suche"
        
        '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
        '# Um neue Module Angebot und AufBest. erweitert.
140     Select Case eBelegTyp
        
            Case Sonderfaktura_Rechnung
            
145             objPRM.FindFirstString = "name = 'singleStringRechnung'"

150             intBelegArt = 0
            
155         Case Sonderfaktura_Gutschrift
            
160             objPRM.FindFirstString = "name = 'singleStringGutschrift'"

165             intBelegArt = 1
            
170         Case Sonderfaktura_Angebot
            
175             objPRM.FindFirstString = "name = 'singleStringAngebot'"

180             intBelegArt = 2
            
185         Case Sonderfaktura_Auftragsbestetigung
            
190             objPRM.FindFirstString = "name = 'singleStringAuftragsbestetigung'"

195             intBelegArt = 3
        
        End Select

        '140     If GintBelegArt = 0 Then
        '145         objPRM.FindFirstString = "name = 'singleStringRechnung'"
        '        Else
        '150         objPRM.FindFirstString = "name = 'singleStringGutschrift'"
        '        End If
        '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >

        '<Modified by: IL at 15.10.2024, Ver.: 6.7.101 >

200     objSQLAusw.caption = objPRM.caption("") & " - " & Replace(objPRM.getColCaption("name = 'singleStringEditState" & Druck & "'"), "Übern.: ", "")

        'ORIG:
        '205     objSQLAusw.caption = objPRM.caption("") & " - " & mnuSuch(Druck).caption
        
        '</Modified by: IL at 15.10.2024, Ver.: 6.7.11 >
        
205     objSQLAusw.top = RowTop
210     objSQLAusw.left = ColLeft
        '100     objSQLAusw.width = 10000
215     objSQLAusw.MaxWidth = 12000 '9420

220     objSQLAusw.ColParameter 0, colWidth, 800
225     objSQLAusw.ColParameter 1, colWidth, 1000
230     objSQLAusw.ColParameter 2, colWidth, 1000
235     objSQLAusw.ColParameter 3, colWidth, 1000
240     objSQLAusw.ColParameter 4, colWidth, 1300
245     objSQLAusw.ColParameter 5, colWidth, 1300
250     objSQLAusw.ColParameter 6, colWidth, 400
255     objSQLAusw.ColParameter 7, colWidth, 1000
260     objSQLAusw.ColParameter 8, colWidth, 1000
265     objSQLAusw.ColParameter 9, colWidth, 3000
  
        'DF 16.01.2015 : Spaltenüberschriften aus PRM holen
  
270     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 0"
275     objSQLAusw.ColParameter 0, ColCaption, objPRM.caption("BelegID")

280     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 1"
285     objSQLAusw.ColParameter 1, ColCaption, objPRM.caption("BelegNr")

290     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 2"
295     objSQLAusw.ColParameter 2, ColCaption, objPRM.caption("BelegDatum")

300     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 3"
305     objSQLAusw.ColParameter 3, ColCaption, objPRM.caption("MCode")

310     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 4"
315     objSQLAusw.ColParameter 4, ColCaption, objPRM.caption("Name1")

320     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 5"
325     objSQLAusw.ColParameter 5, ColCaption, objPRM.caption("Name2")

330     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 6"
335     objSQLAusw.ColParameter 6, ColCaption, objPRM.caption("Lkz")

340     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 7"
345     objSQLAusw.ColParameter 7, ColCaption, objPRM.caption("Plz")

350     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 8"
355     objSQLAusw.ColParameter 8, ColCaption, objPRM.caption("Ort")

360     objPRM.FindFirstString = "name = 'BelegSuchenF2' AND Index = 9"
365     objSQLAusw.ColParameter 9, ColCaption, objPRM.caption("InternerVermerk")

370     objPRM.FindFirstString = ""
  
375     objSQLAusw.SperrenFeld = ""                                            ' DF 20.12.2012
  
380     If Druck = 2 Then                                                      'WENN VORLAGEN

385         objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT BelegID,BelegNr,BelegDatum,MCode,Name1,Name2,Lkz,Plz,Ort,InternerVermerk,Art,Druck FROM [2800_Haupt] WHERE ZwAblage = 1 AND Art = " & intBelegArt & " AND Storno = '0' ORDER BY BelegID DESC"                               'DF 19.01.2015 : ORDER BY Name1 -> BelegID DESC

        Else
        
390         objSQLAusw.RSOpen GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr), "SELECT BelegID,BelegNr,BelegDatum,MCode,Name1,Name2,Lkz,Plz,Ort,InternerVermerk,Art,Druck FROM [2800_Haupt] WHERE ZwAblage = 0 AND Druck = " & Druck & " AND Art = " & intBelegArt & " AND Storno = '0' ORDER BY BelegID DESC"     'DF 19.01.2015 : ORDER BY Name1 -> BelegID DESC

        End If
  
395     If objSQLAusw.Abbruch = False Then

400         gboolBelegAngenommen = True

405         BelegKopfZeigen CLng(objSQLAusw.FieldText(0))

410         Call FillBelegInfo(True, CLng(objSQLAusw.FieldText(0)), CStr(objSQLAusw.FieldText(1)), CInt(objSQLAusw.FieldText(10)), IIf(Druck = 2, 2, CInt(objSQLAusw.FieldText(11))))     'IL 18.10.2024 , Ver.: 6.7.101 : speichern die Hauptdaten in einem Typ für eine spätere bequemere Arbeit

415         Call msgText(1, 2374, 0, 0, 0)

420         If BelegInfo.Druck = 1 Then Call MsgBox(GsMsgText(0), vbOKOnly + vbInformation, strMeldungCap)
    
            '<Removed by: IL at: 22.10.2024, Ver.: 6.7.101 >
            '# Aufgrund der Änderung in der Rechnungsannahmelogik sind jetzt gedruckte Rechnungen möglich bearbeitet werden kann
            
            '405         If Druck = 1 Then Frames False
            
            '</Removed by: IL at: 22.10.2024, Ver.: 6.7.101 >

            '<Modified by: IL at 22.10.2024, Ver.: 6.7.101 >
            '# Aufgrund der Änderung in der Rechnungsannahmelogik sind jetzt gedruckte Rechnungen möglich bearbeitet werden kann

425         Check1(2).Visible = True

            '410         If Druck = 0 Or Druck = 2 Then
            '
            '415             Check1(2).Visible = True                                       'MCode ändern
            '
            '            End If
            '</Modified by: IL at 22.10.2024, Ver.: 6.7.101 >
  
            'DH, 12.01.2016, 6.4.117, Wenn ein bereits vorhandener Beleg ausgewaehlt wurde, dies als Hinweis in der Titelleiste anzeigen.
            '<Modified by: GW at 05.04.2019, Ver.: 6.5.110 >
            
            '<Modified by: IL at 15.10.2024, Ver.: 6.7.101 >
430         Me.caption = objPRM.getColCaption("name = 'frmSP52830' AND index = " & GintBelegArt) & " (" & Replace(objPRM.getColCaption("name = 'singleStringEditState" & Druck & "'"), " ", " " & objPRM.getColCaption("name = 'singleStringVerkurzung" & eBelegTyp & "'") & " ") & ")"
435         gTmpCaption = Me.caption
            
            'ORIG:
            '420         Select Case Druck#
            '
            '                Case 0      'ungedruckt
            '425                 Me.caption = objPRM.getColCaption("name = 'frmSP52830'") & " (" & objPRM.getColCaption("name = 'singleStringEditState0'") & ")"
            '430                 gTmpCaption = Me.caption
            '
            '435             Case 1      'gedruckt
            '440                 Me.caption = objPRM.getColCaption("name = 'frmSP52830'") & " (" & objPRM.getColCaption("name = 'singleStringEditState1'") & ")"
            '445                 gTmpCaption = Me.caption
            '
            '450             Case 2      'Vorlagen
            '455                 Me.caption = objPRM.getColCaption("name = 'frmSP52830'") & " (" & objPRM.getColCaption("name = 'singleStringEditState2'") & ")"
            '460                 gTmpCaption = Me.caption
            '            End Select
            '</Modified by: IL at 15.10.2024, Ver.: 6.7.101 >

            '</Modified by: GW at 05.04.2019, Ver.: 6.5.110 >

        End If
  
440     cmd1(5).SetFocus
 
        Exit Sub

Fehler:
445     Call FehlerErklärung("frmSP52830", "BelegSuchen()")
End Sub

Public Sub Frames(Enabled As Boolean)

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     Frame1(0).Enabled = Enabled
105     Frame1(1).Enabled = Enabled
110     Frame1(2).Enabled = Enabled
115     Frame1(3).Enabled = Enabled
  
        'DH, 13.03.2013, Wenn ein gedruckter Beleg geladen wird, duerfen BelegDatum und -Nr nicht mehr veraendert werden
120     objDruckOptionen.EnableBelegDatum = Enabled
125     objDruckOptionen.EnableBelegNr = Enabled
130     objDruckOptionen.EnableValutaDatum = Enabled    'DH, 17.02.2015, 6.4.103, Gleiches gilt fuer das neue Feld Valuta Datum
  
        '***Beginn
        Exit Sub

Fehler:
135     Call FehlerErklärung("frmSP52830", "Frames")
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

Private Sub SetFormCaption(Optional mitZusatz As Boolean = True)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       SetFormCaption
        ' Description:       Formüberschrift aus zwei Texten zusammen setzen.
        '                    Texte für die überschrift werden aus PRM geholt.
        ' Created by :       GW
        ' Date-Time  :       02.04.2019-14:56:28
        '
        ' Parameters : mitZusatz <wenn true, dann wird captionZusatz angezeigt>
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler
    
        Dim captionZusatz As String

100     If mitZusatz Then

105         If Check1(0).value = 1 Then
110             captionZusatz = objPRM.getColCaption("name = 'frmSP52830_Zusatz' AND index = 0")
            Else
115             captionZusatz = ""
            End If
    
120         Me.caption = gTmpCaption & " " & captionZusatz
        Else
        
            '<Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
125         Me.caption = objPRM.getColCaption("name = 'frmSP52830' AND index = " & GintBelegArt)
130         gTmpCaption = Me.caption

            '125         If GintBelegArt = 0 Then
            '130             Me.caption = objPRM.getColCaption("name = 'frmSP52830' AND index = 0")
            '135             gTmpCaption = Me.caption
            '            Else
            '140             Me.caption = objPRM.getColCaption("name = 'frmSP52830' AND index = 1")
            '145             gTmpCaption = Me.caption
            '            End If
            '</Modified by: IL at 04.10.2024, Ver.: 6.7.101 >
            
        End If
    
        Exit Sub

Fehler:
135     Call FehlerErklärung("frmSP52830", "SetFormCaption()")
   
End Sub

Private Function ShowBelegArchiv()

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       ShowArchiv
        ' Description:       [type_description_here]
        ' Created by :       GW
        ' Date-Time  :       21.02.2020-08:28:44
        '
        ' Parameters :
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler

        Dim archiveType As E_DATATYPE

        Dim Col         As Collection

100     Set Col = New Collection
                 
105     Col.Add cmd1(7)
110     Col.Add mnuBearb1(8)
        
115     Select Case programmNr

            Case "283", "285"

120             archiveType = E_DATATYPE.Sonderfaktura_Rechnung

125         Case "284", "286"
130             archiveType = E_DATATYPE.Sonderfaktura_Gutschrift

                '<Added by: IL at: 11.10.2024, Ver.: 6.7.101 >
135         Case "288"

140             archiveType = E_DATATYPE.Sonderfaktura_Angebot

145         Case "289"

150             archiveType = E_DATATYPE.Sonderfaktura_Auftragsbestetigung
                '</Added by: IL at: 11.10.2024, Ver.: 6.7.101 >

        End Select

155     Call StartBelegArchiv(archiveType, Col)

        Exit Function

Fehler:
160     Call EnableControls(True)
165     Call FehlerErklärung("frmSP52830", "ShowBelegArchiv()")

End Function

Private Function CheckUID(intIndex As Integer) As Boolean

        '--------------------------------------------------------------------------------
        ' Project    :       SP52830
        ' Procedure  :       CheckUID
        ' Description:       Kunden-UID Überprüfen
        ' Created by :       DFiebach
        ' Date-Time  :       11.21.2023-09:36:20
        '
        ' Parameters :       intIndex : 0 = Manuel, 1 = Automatisch (beim Speichern)
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler
        
        Dim blnValidate            As Boolean
        
        Dim strMessage             As String
        
        Dim intAnswer              As Integer
        
        Dim blnShowStandardMessage As Boolean
        
        Dim strKundenLKZ           As String
        
        Dim strKundenKtoArt        As String
        
        Dim intUIDCheck            As Integer
        
        Dim blnKundenKtoArtFailed  As Boolean
        
        Dim blnIsEULand            As Boolean
        
        Dim blnIsEULandAndNotDE    As Boolean
        
100     blnValidate = True
        
        'GÜLTIGKEIT DER UID ONLINE PRÜFEN
        
        'WENN STEUER = 2
105     If CStr(intSteuerTyp) = C_STR_STRFREI_EU Then
            
            '0. UID ist LEER (nur bei EU-Ausland)
110         strKundenLKZ = UCase$(Trim$(txt1(8).text))
            
115         blnIsEULand = modLand.IsEULand(strKundenLKZ, False)                       'Das Land ist ein EU-Land
            
120         blnIsEULandAndNotDE = blnIsEULand And IsLkzDE(strKundenLKZ, True) = False 'Das Land ist ein EU-Land, ABER nicht Deutschland!
            
125         If Trim$(txt1(14).text) = "" And blnIsEULandAndNotDE Then
            
130             MsgBox GetMessage(297, "ZusatzTexte_53100"), vbOKOnly + vbExclamation, strMeldungCap
                
135             CheckUID = False

140             If txt1(14).Enabled Then txt1(14).SetFocus
                
                Exit Function

            End If
            
            '1. LIZENZ
145         If ModuleInLicence(C_PRG_UID_VALIDIERUNG) = False And blnIsEULandAndNotDE Then
            
                'Die Meldung für beide Arten Manuell/Autom.
150             MsgBox GetMessage(2339), vbOKOnly + vbExclamation, strMeldungCap
            
155             Call ResetUIDValidationInfo
            
160             CheckUID = True
                   
                Exit Function
        
            End If
        
            '2. SYSTEM-PARAMETER EINSTELLUNG(MANDANT)
165         intUIDCheck = GetSysPar_UIDCheck
            
170         If intUIDCheck = 0 And blnIsEULandAndNotDE Then                                             ' Die Überprüfung im Mandantenstamm ist ausgeschaltet -> KEINE UID Überprüfung.
                
175             CheckUID = True
                
                'Die Meldung für beide Arten Manuell/Autom.
180             MsgBox GetMessage(2339), vbOKOnly + vbExclamation, strMeldungCap
                
                Exit Function
            
            End If
            
            '3. NUR EU AUSLAND ÜBERPRÜFEN.
            '   WENN kein EU-land, oder wenn Deutscher Kunde ist, oder Lkz LEER ist  -> KEINE UID Überprüfung.
185         If blnIsEULand = False Or IsLkzDE(strKundenLKZ, True) = True Then
            
190             If intIndex = 0 Then
                
195                 MsgBox GetMessage(2328), vbOKOnly + vbExclamation, strMeldungCap
                
200                 Call ResetUIDValidationInfo
                
                End If
            
205             CheckUID = True
                   
                Exit Function
         
            End If
            
            'SYSTEM-PARAMETER AUSWERTEN
210         strKundenKtoArt = UCase$(Trim$(txt1(34).text))
            
215         Select Case intUIDCheck
        
                Case 0                                                              'Keine Überprüfung
                    'DF 01.18.2024 : soll nicht mehr gemeldet werden.
                    '205                 If intIndex = 0 Then MsgBox GetMessage(2332), vbOKOnly + vbExclamation, strMeldungCap
 
220                 blnKundenKtoArtFailed = True

225             Case 1                                                              'Debitoren
                 
230                 If strKundenKtoArt <> C_STR_DEBITOR_KNZ Then
                        'DF 01.18.2024 : soll nicht mehr gemeldet werden.
                        '225                     If (intIndex = 0 Or intIndex = 1) And clsUidWeb.AllreadyValidated = False Then MsgBox GetMessage(2333), vbOKOnly + vbExclamation, strMeldungCap
                    
235                     blnKundenKtoArtFailed = True
                 
                    End If
                 
240             Case 2                                                              'Kreditoren

245                 If strKundenKtoArt <> C_STR_KREDITOR_KNZ Then
                        'DF 01.18.2024 : soll nicht mehr gemeldet werden.
                        '245                     If (intIndex = 0 Or intIndex = 1) And clsUidWeb.AllreadyValidated = False Then MsgBox GetMessage(2334), vbOKOnly + vbExclamation, strMeldungCap
                    
250                     blnKundenKtoArtFailed = True
                 
                    End If
                 
255             Case 3                                                              'Debitoren und Kreditoren
                
260                 If strKundenKtoArt <> C_STR_DEBITOR_KNZ And strKundenKtoArt <> C_STR_KREDITOR_KNZ Then
                        'DF 01.18.2024 : soll nicht mehr gemeldet werden.
                        '265                     If (intIndex = 0 Or intIndex = 1) And clsUidWeb.AllreadyValidated = False Then MsgBox GetMessage(2335), vbOKOnly + vbExclamation, strMeldungCap
                   
265                     blnKundenKtoArtFailed = True
                 
                    End If
                
            End Select
        
270         If blnKundenKtoArtFailed Then
           
275             CheckUID = True
                
280             clsUidWeb.AllreadyValidated = True
                
                Exit Function
         
            End If
            
            '4. MANUELL / AUTOM. BEIM SPEICHERN
285         Select Case intIndex
        
                Case 0 'Manuell
                
                    'ID's PFlicht
290                 blnShowStandardMessage = True
                 
295             Case 1 'Autom. beim Speichern
                 
                    'ID leer oder wurde bereits manuell überprüft - > nicht prüfen.
300                 If Trim$(txt1(14).text) = "" Or clsUidWeb.AllreadyValidated Then ' Trim$(txt1(14).text) = ""  hier evtl. nicht mehr notwendig da bereits oben überprüft wird.
                   
305                     CheckUID = True
                   
                        Exit Function
                    
                    End If
                 
310                 blnShowStandardMessage = False
                
            End Select
        
315         Call EnableControls(False)
        
320         Set clsUidWeb = Nothing
325         Set clsUidWeb = New clsUIDWebValidation
330         clsUidWeb.MCode = Trim$(txt1(0).text)
335         clsUidWeb.Kennung = "Spedifix Logistiksoftware"
340         clsUidWeb.AnfrageZeitstempel = Format(Now, "yyyy-mm-dd-HH-mm-ss")
345         clsUidWeb.VATnummerDesAnfragers = clsUidWeb.GetMandantUID
350         clsUidWeb.AngefragteVATnummer = Trim$(txt1(14).text)
            
            '5. WEB-ÜBERPRÜFUNG
355         blnValidate = clsUidWeb.DoValidation(blnShowStandardMessage, True, Me, True)
        
360         clsUidWeb.AllreadyValidated = True

365         Call EnableControls(True)
        
370         If intIndex = 1 And blnValidate = False Then                           'Nur beim Speichern und wenn Ergebnis = falsch ist.

375             strMessage = GetMessage(2336)
380             strMessage = Replace$(strMessage, "%1", clsUidWeb.WebErrorResult.ResponseDescription)
385             intAnswer = MsgBox(strMessage, vbOKOnly + vbExclamation, strMeldungCap)
                
                'DF 02.21.2024 : Die Logik mit Ja/Nein geändert -> OK. Der Vorgang kann nicht fortgesetzt werden
                '375             If intAnswer = vbYes Then
                '
                '380                 blnValidate = True
                '
                '                Else
              
390             blnValidate = False
                
395             clsUidWeb.AllreadyValidated = False
                
400             If txt1(14).Enabled And txt1(14).Visible Then txt1(14).SetFocus
                
                '                End If
           
            End If
            
        End If

405     CheckUID = blnValidate
        
        Exit Function
    
Fehler:

410     LockWindowUpdate 0&
415     Call EnableControls(True)
420     Me.MousePointer = vbDefault
425     Call FehlerErklärung("frmSP52830", "CheckUID()")

End Function

Private Sub ResetUIDValidationInfo()

        '--------------------------------------------------------------------------------
        ' Project    :       frmSP52830
        ' Procedure  :       ResetUIDValidationInfo
        ' Description:       Setzt die Anzeige der UID-Überprüfung zurück.
        ' Created by :       DFiebach
        ' Date-Time  :       11.22.2023-09:49:15
        '
        ' Parameters :
        '--------------------------------------------------------------------------------
        On Error GoTo Fehler
        
110     If Not clsUidWeb Is Nothing Then clsUidWeb.AllreadyValidated = False
        
        Exit Sub
    
Fehler:
    
115     Me.MousePointer = vbDefault
120     Call FehlerErklärung("frmSP52830", "ResetUIDValidationInfo()")
End Sub

Private Sub SteuerungFiBuUndSpeditionsbuch(boolZustand As Boolean)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       SteuerungFiBuUndSpeditionsbuch
        ' Description:       Die Funktion steuert die Aktivität der Felder „Fibu“ und „SpeditionsBuch“. Sie müssen mit „Angebot“ und "Auftragsbestetigung" ausgeschaltet werden.
        ' Created by :       IL
        ' Date-Time  :       11.10.2024-08:43:52
        '
        ' Parameters :       boolZustand (Boolean)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler

        Dim i As Integer

100     For i = 25 To 36

105         If i <> 34 Then

110             txt1(i).Enabled = boolZustand

115             lbl1(i).Enabled = boolZustand

                On Error Resume Next

120             cmdAuswahl(i).Enabled = boolZustand

                On Error GoTo Fehler

            End If

125     Next i
    
        Exit Sub

Fehler:

130     Me.MousePointer = vbDefault
135     Call FehlerErklärung("frmSP52830", "SteuerungFiBuUndSpeditionsbuch()")

End Sub

Public Sub ParentSchalter(boolEnabled As Boolean)

        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       ParentSchalter
        ' Description:       Deaktivieren/aktivieren Objekte im Parent-Form, wenn das aktuelle ausblenden
        ' Created by :       IL
        ' Date-Time  :       05.11.2024-14:44:29
        '
        ' Parameters :       boolEnabled (Boolean)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler

100     Frame1(0).Enabled = boolEnabled
105     Frame1(1).Enabled = boolEnabled
110     Frame1(2).Enabled = boolEnabled
115     Frame1(3).Enabled = boolEnabled

120     cmd1(0).Enabled = boolEnabled
125     cmd1(8).Enabled = boolEnabled

130     mnuBearb1(5) = boolEnabled
135     mnuBearb1(7) = boolEnabled

140     mnuOpt1(1) = boolEnabled
145     mnuOpt1(2) = boolEnabled
150     mnuOpt1(3) = boolEnabled
155     mnuOpt1(4) = boolEnabled
160     mnuOpt1(50) = boolEnabled
        
        Exit Sub

Fehler:
165     Me.MousePointer = vbDefault
170     Call FehlerErklärung("frmSP52831", "ParentSchalter()")
End Sub

