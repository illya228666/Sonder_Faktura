Attribute VB_Name = "SP52800B"
Option Explicit

'Public GWS As DAO.Workspace
Public Const Gc_strExeFile = "SP62800.exe"

Public GsNeuerArtikel As String

'Public Const TEXT_BREITE = 4700 'Breite des Eingabefeldes frmSP52810.txt1(0)
'Public Const TEXT_BREITE = 4900 'HW 05.08.2011 Ver.: 6.1.107 Breite des Eingabefeldes frmSP52810.txt1(0)
'Public Const TEXT_BREITE = 5200 'DH, 16.09.2011, in Folge der Button-Neuausrichtung verbreitert

Public Const TEXT_BREITE = 5600                                                 '5518 'HW 09.07.2012 Ver.: 6.1.114  Neuanordnung

Public GlngArbeitsplatz   As Long

Public GintBelegArt       As Integer                                            '0=Rechnung, 1=Gutschrift, 2=Angebot, 3=Auftragsbestetigung

Public GintBelegNrKreisNr As Integer                                            'NrKreis einer Beleges

Public GstrAuftraggeber   As String                                             'Wird benutzt beim abweichenden Rechnungsempfänger. DF 06.02.2019 : wird nicht mehr benötigt, da die Logik des automatischen TEXT weg ist.

Public frmRechnung        As Form

Public frmGutschrift      As Form

Public frmAngebot         As Form

Public frmAuftragsbest    As Form

Public frmRechnungErf     As frmSP52831

'Public frmRechnungErf As Form
Public frmGutschriftErf   As frmSP52831

Public frmAngebotErf      As frmSP52831

Public frmAuftragsbestErf As frmSP52831

Public frmRechnungFakt    As Form

Public frmGutschriftFakt  As Form

Public gMandant000        As Boolean                                           'HW 04.05.2011

Public GblnExternesArchiv As Boolean

'** List & Label *********
Private glRet             As Long

Private glDummy           As Long
'*************************

Public Const STEUERTYP0 = "steuerfrei"

Public Const STEUERTYP1 = "steuerpflichtig"

Public Const STEUERTYP2 = "steuerfrei/EG"

Public programmNr                As String                                     'HW 17.09.2012

Public gstrSteuerText            As String                                     'HW 09.07.2012 Ver.: 6.1.114

Public intSteuerTyp              As Integer                                    'HW 09.07.2012 Ver.: 6.1.114   ein integer wird beim speichern erwartet und kein String!

Public gdbZusatzText             As ADODB.Connection                           'HW 18.03.2013

Public printDone                 As Boolean                                    'DH, 11.07.2013, Flag welches anzeigt ob ein Druck (nicht die Vorschau) ausgefuehrt wurde

Public printJobInProgress        As Boolean                                    'DF 13.04.2023 , Ver.: 6.6.119 : Druck, Archivierung-Vorgang beendet oder nicht.

Public saveDone                  As Boolean

Public stornoDone                As Boolean

'<Modified by: GW at 21.02.2020, Ver.: GOBD >
'Public objBelegArchiv    As clsBelegArchiv 'DH, 26.11.2015, 6.4.112, Globale Variable des BelegArchivs angelegt
Public objEmailSending           As clsEmailSending
'</Modified by: GW at 21.02.2020, Ver.: GOBD >

Public rsBelegArchiv             As ADODB.Recordset                            'DH, 26.11.2015, 6.4.112, Recordset fuer den eMailversand ueber das BelegArchiv

Public BearbeiterDrucken         As Boolean                                    'HW 29.04.2016 globale Variable

Public GesamtIstBrutto           As Boolean                                    'HW 24.05.2016 Ver.: 6.4.120 globale Variable

Public blnMaskeLeeren            As Boolean

Public blnMaskeSchliessen        As Boolean

Public blnFolgeseitenKurzDrucken As Boolean

Public g_objCal                  As clsKalender                                'DH, 30.05.2017, 6.4.126

Public postCommaPreis            As Integer                                    'DH, 27.10.2017, 6.5.101, Legt fest, wieviele Nachkommastellen im Feld Preis erlaubt sind

'<Added by: GW at: 03.04.2019, Ver.: 6.5.110 >
'wird in der Druckmaske zum Anzeigen der BelgeNr verwendet
Public gLngBelegNr               As Long
'</Added by: GW at: 03.04.2019, Ver.: 6.5.110 >

Public vValue                    As Variant                                  'IL 26.07.2024 Wert der Umzatz-Steuer

Public dblUstSatz                As Double                                   'IL 26.07.2024, Steuerprozentsatz

Public Sub SpeditionsBuch(rsH As ADODB.Recordset, SteuerPfl As Double, SteuerFr As Double)

        On Error GoTo Fehler

        Dim rs       As ADODB.Recordset

        Dim sql      As String

        Dim WrgSchl  As String

        Dim BelegArt As String

        Dim ErfNr    As Variant

        Dim mult     As Integer
  
100     If rsH.RecordCount > 0 Then

105         If Not IsNull(rsH!Datum) And Trim(rsH!KostenArt) <> "" Then

110             OPEN_gConn

115             If rsH!Art = 0 Then
120                 BelegArt = C_STR_AUSGANGSRECHNUNG
125                 mult = 1
                Else
130                 BelegArt = C_STR_AUSGANGSGUTSCHRIFT
                    'AusgangsGutschriften müssen als - Betrag gespeichert werden
135                 mult = -1
                End If
                
140             Set rs = New ADODB.Recordset

                'Währungsschlüßel
                '<Removed by: Project Administrator at: 6.17.2020, Ver.: 6.6.102 >
                ' Ab Update 6.6.101 wird Währungs-Schlüssel in SOFA/LM/LAGER Beleg-Datensatz gespeichert und somit wird diese Logik nicht mehr benötigt.

                '145             rs.Open "SELECT Schl FROM [1100_Währungen] WHERE MwSt = " & rsH!MwSt & " AND ISO = '" & rsH!Wrg1 & "'", gConn, adOpenStatic, adLockReadOnly
                '
                '150             If rs.RecordCount > 0 Then
                '
                '155                 WrgSchl = rs!Schl
                '
                '                Else
                '
                '160                 rs.Close
                '165                 rs.Open "SELECT Schl FROM [1100_Währungen] WHERE ISO ='" & rsH!Wrg1 & "'", gConn, adOpenKeyset, adLockReadOnly
                '
                '170                 If rs.RecordCount > 0 Then
                '175                     WrgSchl = rs!Schl
                '                    Else
                '180                     WrgSchl = GmandantRS!EigenWrg
                '                    End If
                '
                '                End If
                '</Removed by: Project Administrator at: 6.17.2020, Ver.: 6.6.102 >

                '<Added by: Project Administrator at: 6.17.2020, Ver.: 6.6.102 >
                ' # Wenn der WrgSchl leer ist, dann alte Logik verwenden (sicherheit, eingentlich wird beim Speichern überprüft),
                ' # wenn nicht neue Logik -> WrgSchl aus Beleg-Datensatz
145             If IsNull(rsH.Fields("WrgSchl").value) Or Trim$(rsH.Fields("WrgSchl").value) = "" Then
                    
150                 rs.Open "SELECT Schl FROM [1100_Währungen] WHERE MwSt = " & rsH!MwSt & " AND ISO = '" & rsH!Wrg1 & "'", gConn, adOpenStatic, adLockReadOnly
                
155                 If rs.RecordCount > 0 Then
                
160                     WrgSchl = rs!Schl
                
                    Else
                
165                     rs.Close
170                     rs.Open "SELECT Schl FROM [1100_Währungen] WHERE ISO ='" & rsH!Wrg1 & "'", gConn, adOpenKeyset, adLockReadOnly
                
175                     If rs.RecordCount > 0 Then
180                         WrgSchl = rs!Schl
                        Else
185                         WrgSchl = GmandantRS!EigenWrg
                        End If
                
                    End If
                    
                Else
                
190                 WrgSchl = rsH.Fields("WrgSchl").value
                    
                End If

                '</Added by: Project Administrator at: 6.17.2020, Ver.: 6.6.102 >
                
195             If rs.state = adStateOpen Then rs.Close

200             rs.Open "SELECT * FROM [5800_SpeditionsBuch] WHERE AbfPos = 999000999", gConn, adOpenKeyset, adLockOptimistic   'Damit keine Sätze zurück gegeben werden.

205             rs.AddNew
210             rs!AbfPos = rsH!AbfPos
215             rs!AbfZus = rsH!AbfZus
220             rs!AbfDat = rsH!Datum

225             If Trim(rsH!ErfNr) <> "" Then rs!ErfNr = rsH!ErfNr
230             rs!BelegArt = BelegArt
235             rs!BelegNr = Trim(Left(CStr(rsH!BelegNr), 12))
240             rs!belegDatum = rsH!belegDatum

245             If Trim(rsH!MCode) <> "" Then rs!MCode = rsH!MCode
250             If Trim(rsH!KtoKnz) <> "" Then rs!KtoKnz = rsH!KtoKnz
255             If Trim(rsH!KtoNr) <> "" Then rs!KtoNr = rsH!KtoNr
260             If Trim(rsH!Uid) <> "" Then rs!Uid = rsH!Uid
265             If Trim(rsH!Name1) <> "" Then rs!name = rsH!Name1
270             If Trim(rsH!Lkz) <> "" Then rs!Lkz = rsH!Lkz
275             If Trim(rsH!Plz) <> "" Then rs!Plz = rsH!Plz
280             If Trim(rsH!Ort) <> "" Then rs!Ort = rsH!Ort

285             rs!Bemerkung = "Sonderfaktura"
290             rs!KostenArt = rsH!KostenArt
295             rs!Rückstellung = "0"

300             If Trim(rsH!KostSchl) <> "" Then rs!KoSt = rsH!KostSchl
305             If Trim(rsH!FiBuSchl) <> "" Then rs!ErlBer = rsH!FiBuSchl
310             rs!Steuer = CStr(rsH!Ust)
315             rs!BetragStpfl = SteuerPfl * mult 'Bis Ver.:5.1-104; 11.05.04 Runden(SteuerPfl * mult * ((100 + rsH!MwSt) / 100), 2)
320             rs!BetragStfrei = SteuerFr * mult
325             rs!WrgSchl = WrgSchl

330             If Trim(rsH!Wrg1) <> "" Then rs!Wrg = rsH!Wrg1
335             rs!Druck = "0"
340             rs!SpedBuchListeID = 0
345             rs!RngEingNr = 0
350             rs!ReBuch = "0"
355             rs!ErstVon = "SF" & rsH!BelegID
360             rs!AendVon = "Sonderfaktura"
365             rs.Update

370             Protokoll iAppend, vbCrLf & "SpedBuch schreiben -> BelegNr: " & rsH!BelegNr & " BelegID: " & rsH!BelegID & " AbfPos: " & rs!AbfPos & " AbfZus: " & rs!AbfZus & " AbfDatum: " & rs!AbfDat
375             Protokoll iAppend, "BetragStpfl: " & rs!BetragStpfl & " BetragStfrei: " & rs!BetragStfrei

380             rs.Close

            End If
        End If

        Exit Sub

Fehler:
385     Call FehlerErklärung("SP52800B", "SpeditionsBuch()")
End Sub

Public Function NummernKreisWaehlen(NummernKreis As Integer) As Integer

        On Error GoTo Fehler
        
        '##### DF 29.01.2019: WIRD EVTL. NICHT MEHR VERWENDET. S. <NummernKreisWaehlenSQL> #####
        
        'Die Funktion wird benutzt um zu überprüfen, ob für
        'Lagerrechnungen gesonderer Nummernkreis benutzt werden soll.
        'Wenn Felder, im Satz 19 (Lager-Rechnungs-Nr), von = 0, aktuell = 0 und bis = 0 sind,
        'wird der allgemaine Nummernkreis für Rechnungen (Satz 3) benutzt.
        Dim rs  As ADODB.Recordset

        Dim sql As String
  
100     NummernKreisWaehlen = NummernKreis
  
105     Set rs = New ADODB.Recordset
  
110     OPEN_gConn
  
        'HW 04.11.2015 Es soll nur aktuell auf 0 überprüft werden
115     rs.Open "SELECT Nr FROM [1100_NummernKreise] WHERE aktuell = 0 AND Nr = " & NummernKreis, gConn, adOpenStatic, adLockReadOnly
  
120     If rs.RecordCount > 0 Then

125         NummernKreisWaehlen = NummernKreis - 5
    
            'HW 04.11.2015 Es soll nur aktuell auf 0 überprüft werden
            'HW 04.05.2011 Ver.: 6.1.102 auf Mandant 0 zugreifen und NrKreis holen!
            '#################################################
130         rs.Close
    
135         rs.Open "SELECT Nr FROM [1100_NummernKreise] WHERE Nr = " & CStr(NummernKreis - 5) & " AND aktuell = 0 ", gConn, adOpenStatic, adLockReadOnly

140         If rs.RecordCount > 0 Then
145             gMandant000 = True
            Else
150             gMandant000 = False
            End If

            '#################################################
        End If

155     rs.Close
 
        Exit Function

Fehler:
160     Call FehlerErklärung("SP52800B", "NummernKreisWaehlen()")
End Function

Public Function NummernKreisWaehlenSQL(NummernKreis As Integer) As Integer

        'Die Funktion wird benutzt um zu überprüfen, ob für
        'Lagerrechnungen gesonderer Nummernkreis benutzt werden soll.
        'Wenn Felder, im Satz 19 (Lager-Rechnungs-Nr), von = 0, aktuell = 0 und bis = 0 sind,
        'wird der allgemaine Nummernkreis für Rechnungen (Satz 3) benutzt.
        Dim rs  As ADODB.Recordset

        Dim RS1 As ADODB.Recordset

        Dim cn  As ADODB.Connection

        Dim sql As String
 
        On Error GoTo Fehler
        
100     NummernKreisWaehlenSQL = NummernKreis
    
105     Set cn = New ADODB.Connection
110     cn.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
115     cn.Open

        'HW 04.11.2015 Es soll nur aktuell auf 0 überprüft werden
120     sql = "SELECT Nr FROM [1100_NummernKreise] WHERE aktuell = 0 AND Nr = " & NummernKreis
125     Set rs = New ADODB.Recordset
130     rs.Open sql, cn, adOpenKeyset, adLockReadOnly
  
135     If rs.RecordCount > 0 Then
            
140         gblnBelegNrStandardNrKreis = True                                          'DF 24.01.2019 , Ver.: 6.5.109 :    Zeiger auf BelegNr as StandardNrKreis
            
145         NummernKreisWaehlenSQL = NummernKreis - 5
    
            'HW 04.11.2015 Es soll nur aktuell auf 0 überprüft werden
            'HW 04.05.2011 Ver.: 6.1.102 auf Mandant 0 zugreifen und NrKreis holen!
            '#################################################
150         sql = "SELECT Nr FROM [1100_NummernKreise] WHERE Nr = " & CStr(NummernKreis - 5) & " AND aktuell = 0 "
155         Set RS1 = New ADODB.Recordset
160         RS1.Open sql, cn, adOpenKeyset, adLockReadOnly

165         If RS1.RecordCount > 0 Then

170             gMandant000 = True
                'DF 29.01.2019 , Ver.: 6.5.109 : <gblnBelegNrMandant000> hat gleiche Funktion wie <gMandant000>, kommt aber aus <modFaktura> und soll künftig verwendet werden,
                '                                weil <gMandant000> in zwei Modulen definiert ist.
175             gblnBelegNrMandant000 = True
                
            Else
            
180             gMandant000 = False
                'DF 29.01.2019 , Ver.: 6.5.109 : <gblnBelegNrMandant000> hat gleiche Funktion wie <gMandant000>, kommt aber aus <modFaktura> und soll künftig verwendet werden,
                '                                weil <gMandant000> in zwei Modulen definiert ist.
185             gblnBelegNrMandant000 = False
                
            End If

190         RS1.Close
195         Set RS1 = Nothing
            '#################################################
            
        Else
        
200         gblnBelegNrStandardNrKreis = False                                         'DF 24.01.2019 , Ver.: 6.5.109 :    Zeiger auf BelegNr as StandardNrKreis
            
        End If
  
205     rs.Close
210     Set rs = Nothing
  
215     cn.Close
220     Set cn = Nothing
  
        Exit Function

Fehler:

225     Call FehlerErklärung("SP52800B", "NummernKreisWaehlenSQL()")

End Function

'***********************************************************************************
'Routine:           NummernKreis
'
'Autor:             mtrdy@native.cz, 29.11.2001
'Beschreibung:      liest den Wert aus 1100_NummernKreise
'Parameter:         Nr
'                   InNewWorkspace - wenn TRUE, neu Workspace benutzt ist
'
'Änderungen:
'***********************************************************************************
Public Function NummernKreisSQL(nr As Long, _
                                Optional StandardFehlerMeldung As Boolean = True) As Long

        On Error GoTo error

        Dim cn           As ADODB.Connection

        Dim rs           As ADODB.Recordset

        Dim ErrorCount   As Long

        Dim TempKeyValue As Long

        Dim NewKeyValue  As Long

        Dim bCloseWs     As Boolean

        Dim bTrans       As Boolean
  
100     Set cn = New ADODB.Connection

105     If gMandant000 Then 'HW 04.05.2011 Ver.: 6.1.102 auf Mandant 0 zugreifen und NrKreis holen!
110         cn.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, "000")
        Else
115         cn.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
        End If

120     cn.Open

125     TempKeyValue = GetNrKreisValue_EXTENDED(cn, CLng(nr), False)

130     cn.Close

135     NummernKreisSQL = TempKeyValue

        Exit Function
    
error:
140     NummernKreisSQL = -1

145     If StandardFehlerMeldung Then
150         If Not IsLockError(Err.number) Then
155             Call FehlerErklärung("SP52800B", "NummernKreis()")
            End If
        End If

End Function

'Änderungen:
'  mtrdy@native.cz 05.04.02 *****
'     - removed error handling - it was clearing Err
Public Function IsLockError(lErr As Long) As Boolean

        '***Beginn
        On Error GoTo Fehler

        '***Ende
100     Select Case lErr

            Case 3008, 3009, 3188, 3189, 3211, 3260, 3261, 3262, 3218
105             IsLockError = True

110         Case Else
115             IsLockError = False
        End Select

        '***Beginn
        Exit Function

Fehler:
120     Call FehlerErklärung("SP52800B", "IsLockError")
        '***Ende
End Function

Public Function LLPrintListe(frm As Form, _
                             LL1 As ListLabel.ListLabel, _
                             BelegID As Long, _
                             Mode As Integer, _
                             Optional tmp As Boolean, _
                             Optional Save As Boolean, _
                             Optional SammelDruck As Boolean) As Long
        'Mode = 4 -> Ablage
        'Mode = 3 -> Archivierung
        'Mode = 2 -> Vorschau
        'Mode = 1 -> Druck
        'Mode = 0 -> Druck-Wiederholung (Druckstatus darf nicht = 0 gesetzt werden.)
        'Tmp = True -> Der Schalter wird nur bei der Vorschau noch nicht gedruckter Belege gesetzt.
        'Die Rechnungsdaten wurden zuvor in Tmp Tabellen gespeichert.
        'Save = True -> LL-Datei wird gesichert.
        'Die Option wird genutzt um die Belege zu archivieren. (Im 2-ten Durchlauf nachdem die Belege gedruckt wurden.)
        'SammelDruck -> Wird von SP52850.LLPrintSammel verwaltet.
  
        On Error GoTo Fehler

        Dim Formular                 As String

        Dim i                        As Integer

        Dim j                        As Integer

        Dim Msg                      As Boolean

        Dim rs                       As ADODB.Recordset

        Dim rsH                      As ADODB.Recordset

        Dim RS1                      As ADODB.Recordset

        Dim ZwSumme                  As Double

        Dim DruckerDialog            As Boolean
  
        Dim SteuerPfl                As Double

        Dim SteuerFr                 As Double

        Dim Ust                      As Double

        Dim UstTMP                   As Double

        Dim Betrag                   As Double

        Dim SteuerPflWrg             As Double

        Dim SteuerFrWrg              As Double

        Dim UStWrg                   As Double

        Dim BetragWrg                As Double

        Dim Kurs                     As Double
  
        Dim sql                      As String

        Dim BelegArt                 As Integer

        Dim belegDatum               As Variant

        Dim BelegNr                  As Long

        Dim Waehrung                 As String

        Dim Skonto                   As Single

        Dim SkontoTage               As Integer

        Dim nettoTage                As Integer

        Dim MwSt                     As Single

        Dim Seite                    As Long

        Dim TmpZusatz                As String
  
        Dim barcodeDaten             As BarcodeData
        
        Dim SteuerText               As Variant                                'HW 05.11.2010 Ver.: 5.4.122 - Ver.: 6.1.101
        
        Dim strSteuerTextAusFormular As String                                 'DF 11.07.2024 , Ver.: 6.7.101 : der entgültige SteuerText wird im Formular ausgewählt. Daher wird es zum Speichern daraus geholt.
        
        Dim strZahlungsTextNetto     As String                                 'DF 11.07.2024 , Ver.: 6.7.101
         
        Dim strZahlungsText          As String                                 'DF 11.07.2024 , Ver.: 6.7.101

        Dim rec1100Texte             As ADODB.Recordset                   'HW 09.07.2012 Ver.: 6.1.114
  
        Dim ArchivierungsModus       As Integer                                'HW 23.09.2015

        Dim idCollection             As Collection

        Dim PercentPosition          As Integer
        
        Dim strLogBelegArt           As String                                 'DF 16.01.2019 , Ver.: 6.5.109 : DruckArt -String für LOG-Datei
        
        Dim lngBelegNrKres           As Integer                                'DF 05.02.2019 , Ver.: 6.5.109 : Nummer des NrKreises
        
        Dim strStCodeH               As String                                 'DF 29.07.2024 , Ver.: 6.7.101 : St.Code des Hauptsatzes (E-Rechnung)
                
        Dim intSteuerTextLkz         As Integer                                'DF 29.07.2024 , Ver.: 6.7.101 : Lkz des SteuerTextes für die ganze Rechnung , wird anhand des Steuer-Schlüssel der Rechnung ermittelt.
        
        'CSBmk <LOG START>
100     Select Case Mode
        
            Case 0         'DRUCK WIEDERHOLUNG
                 
105             Protokoll iAppend, ">DRUCK START (MODUS: DRUCK WIEDERHOLUNG) -> BelegID: " & BelegID & ""
                 
110         Case 1         'DRUCK
            
115             Protokoll iAppend, ">DRUCK START (MODUS: DRUCK) -> BelegID: " & BelegID & ""
            
120         Case 2         'VORSCHAU
            
125             Protokoll iAppend, ">DRUCK START (MODUS: VORSCHAU) -> BelegID: " & BelegID & ""
            
130         Case 3         'ARCHIVIERUNG
        
135             Protokoll iAppend, ">DRUCK START (MODUS: ARCHIVIERUNG) -> BelegID: " & BelegID & ""

140         Case 4         'ABLAGE
        
145             Protokoll iAppend, ">DRUCK START (MODUS: ABLAGE) -> BelegID: " & BelegID & ""
            
        End Select
        
150     If tmp Then

155         TmpZusatz = "Tmp"

        End If

160     Set rs = New ADODB.Recordset
165     Set rsH = New ADODB.Recordset
170     Set RS1 = New ADODB.Recordset
175     Set rec1100Texte = New ADODB.Recordset

180     OPEN_gConn
        
        'CSBmk <HAUPT-RECORDSET VARIABLEN>
185     rsH.Open "SELECT * FROM [2800_Haupt" & TmpZusatz & "] WHERE BelegID = " & BelegID, gConn, adOpenKeyset, adLockOptimistic
 
190     If rsH.RecordCount > 0 Then

195         LL1.LlDefineVariableStart                                           'Variablenpuffer löschen.

200         LL1.LlDefineFieldStart                                              'Variablenpuffer löschen.
            
205         j = 1

210         If Mode = 3 Then
215             ArchivierungsModus = 1
            Else
220             ArchivierungsModus = 0
            End If

225         llCurrentFormNr = CInt(objDruckOptionen.FormularNr)                'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

230         Call LL18GestaltungFormular(LL1, objDruckOptionen.FormularNr, "" & rsH.Fields("MCode").value, MandantArr(1), , , ArchivierungsModus)   'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr
235         j = 2

240         Call LLDefineVariablen(LL1, rsH, "Kd_")
245         j = 3
            
250         Call LLDefineTexte(LL1)                                             'DF 24.10.2024 , Ver.: 6.7.101 : ZusatzTexte usw.
            
            '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
            '# ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
255         If rsH!MwSt = 0 And CInt(vValue) <> 0 Then

260             rsH!MwSt = dblUstSatz

265             rsH.Update

270             Call LLDefineFelder(LL1, rsH, "Kd_")                            'Deklarationen

275             rsH!MwSt = 0

280             rsH.Update

            Else
            
285             Call LLDefineFelder(LL1, rsH, "Kd_")

            End If

            '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
290         j = 4
            
            'CSBmk <OPT:BEARBEITER DRUCKEN>
295         If BearbeiterDrucken Then

300             LL1.LlDefineVariableExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
305             LL1.LlDefineFieldExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN

            Else
            
310             LL1.LlDefineVariableExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
315             LL1.LlDefineFieldExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN

            End If
            
            'CSBmk <OPT:ADRESSE AUF FOLGESEITEN DRUCKEN>
320         If blnFolgeseitenKurzDrucken Then                                   'Added by: GW at: 24.04.2019, Ver.: 6.5.111

325             LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN
330             LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN

            Else
            
335             LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN
340             LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN

            End If
            
            'CSBmk <OPT:BRUTTO-NETTO UMRECHNUNG>
345         If GesamtIstBrutto Then                                             'HW 24.05.2016 Ver.: 6.4.120
350             LL1.LlDefineFieldExt "GesamtIstBrutto", "TRUE", LL_BOOLEAN
            Else
355             LL1.LlDefineFieldExt "GesamtIstBrutto", "FALSE", LL_BOOLEAN
            End If
            
360         If objDruckOptionen.CurrentBelegDatum <> "" Then                    'DH, 11.07.2013, BelegDatum aus den DruckOptionen uebernehmen (sofern eingestellt)

365             LL1.LlDefineVariableExt "Kd_BelegDatum", objDruckOptionen.CurrentBelegDatum, LL_DATE_LOCALIZED

370             LL1.LlDefineFieldExt "Kd_BelegDatum", objDruckOptionen.CurrentBelegDatum, LL_DATE_LOCALIZED

            Else

375             If Mode = 2 Then                                                'Wenn die Vorschau aufgerufen wurde

380                 LL1.LlDefineVariableExt "Kd_BelegDatum", 0, LL_DATE
385                 LL1.LlDefineFieldExt "Kd_BelegDatum", 0, LL_DATE

                End If
                
            End If
  
390         LL1.LlDefineFieldExt "ProbeDruckText", ZusatzText(4, "55710"), LL_TEXT 'HW 16.10.2013
395         LL1.LlDefineVariableExt "ERechnungArt", 0, LL_NUMERIC               'DF 04.11.2024 , Ver.: 6.7.101

400         If Mode = 2 Then                                                    'HW 16.10.2013 Wenn ProbeDruck Dann
                
405             LL1.LlDefineFieldExt "ProbeDruck", 1, LL_NUMERIC                'IL 21.10.2024 , Ver.: 6.7.101 :    mode ----> 1; Um die Beschriftung anzuzeigen, muss der Parameter gleich 1 und nicht 2 sein

            Else
            
410             LL1.LlDefineFieldExt "ProbeDruck", 0, LL_NUMERIC                'HW 16.10.2013

            End If

415         Call DefineZusatztext(rsH, LL1)                                     'MW 13.11.08 Ver.: 5.4.119 Zusatztext
            
420         BelegArt = rsH!Art
425         belegDatum = rsH!belegDatum
430         Waehrung = rsH!Wrg1
435         Skonto = rsH!ZSkto
440         SkontoTage = rsH!ZSktoTage
445         nettoTage = rsH!ZTage
450         MwSt = rsH!MwSt
455         Kurs = rsH!Kurs
460         BelegNr = rsH!BelegNr

465         Select Case GintBelegArt
            
                Case 0
                
470                 strLogBelegArt = "Rechnungsdruck"
                
475             Case 1
                
480                 strLogBelegArt = "Gutschriftsdruck"
                
485             Case 2
                
490                 strLogBelegArt = "Angebot"
                
495             Case 3
                
500                 strLogBelegArt = "Auftragsbestätigung"
            
            End Select                  'Added by: DFiebach at: 16.01.2019, Ver.: 6.5.109
            
            '<Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
            'St.Code des Hauptsatzes ahnahd des gewählten St.Schl
               
            'CSBmk <KUNDEN E-RECHNUNG EINSTELLUNG>
505         If GesamtIstBrutto Then gEnmKudnenERechnungType = eERechnungType.None                            'DF 04.09.2024 , Ver.: 6.7.101, keine ERechnung bei Brutto-Netto Umrechnung
        
510         If IsEBelegDoc Then LL1.LlDefineVariableExt "ERechnungArt", CInt(gEnmKudnenERechnungType), LL_NUMERIC

515         Select Case Mode
            
                Case 1, 4                                                      'IL 07.11.2024 , Ver.: 6.7.101 : Ablage hinzugefügt
                            
520                 intSteuerTextLkz = objDruckOptionen.CurrentSteuerValue
                
                    'CSBmk <STEUER-CODE HAUPT>
525                 strStCodeH = GetStCodeFromSteuerText(CStr(intSteuerTextLkz), "Rng")

            End Select

            '</Added by: DFiebach at: 26.07.2024, Ver.: 6.7.101 >
            
            'CSBmk <FOLGE-RECORDSET VARIABLEN>
530         rs.Open "SELECT * FROM [2800_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID & " ORDER BY Nr", gConn, adOpenStatic, adLockReadOnly

535         j = 5
    
540         If rs.RecordCount > 0 Then

545             Seite = 1

550             Call LLDefineVariablen(LL1, rs, "Re_")

555             LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
560             LL1.LlDefineFieldExt "LetzteSeite", 0, LL_NUMERIC
565             LL1.LlDefineFieldExt "Re_ZwSumme", 0, LL_NUMERIC
570             LL1.LlDefineFieldExt "ZahlungsZiel", "", LL_TEXT
575             LL1.LlDefineFieldExt "ZahlungsZielNetto", "", LL_TEXT

580             j = 6
                
585             EndBetraege "2800_Folge" & TmpZusatz, BelegID, SteuerPfl, SteuerFr  'Betrag für ZahlungsZiel

                'CSBmk <STEUER>
590             If MwSt = 0 Then                                               'IL 25.07.2024, Nehmen den Steuerprozentsatz aus der Datenbank, wenn er im Form 0 ist

595                 MwSt = GetWaehrung(rsH!WrgSchl, False).MwSt

                End If
                
600             If GesamtIstBrutto Then                                         'HW 24.05.2016 Ver.: 6.4.120

605                 UstTMP = SteuerBetrag(CCur(SteuerPfl), Format(CCur(MwSt), "#.00"), False)

610                 Ust = SteuerPfl - UstTMP

                Else

615                 Ust = SteuerBetrag(CCur(SteuerPfl), Format(CCur(MwSt), "#.00"))
       
                End If

620             If GesamtIstBrutto Then                                         'HW 24.05.2016 Ver.: 6.4.120

625                 SteuerPfl = (SteuerPfl - Ust)

                End If
                
                'CSBmk <BERECHNUNG>
630             Betrag = SteuerPfl + Ust + SteuerFr                             'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
635             SteuerPflWrg = RundenMitVz(SteuerPfl * Kurs, 2)
640             SteuerFrWrg = RundenMitVz(SteuerFr * Kurs, 2)
645             UStWrg = Runden(Ust * Kurs, 2)
650             BetragWrg = SteuerPflWrg + UStWrg + SteuerFrWrg

655             LL1.LlDefineFieldExt "Re_EPreisDezStellen", postCommaPreis, LL_NUMERIC       'DH, 27.10.2017, 6.5.101, Einstellung aus den Systemparameter uebergeben
660             LL1.LlDefineFieldExt "Re_SummeSteuerPfl", SteuerPfl, LL_NUMERIC
665             LL1.LlDefineFieldExt "Re_SummeSteuerFr", SteuerFr, LL_NUMERIC
670             LL1.LlDefineFieldExt "Re_USt", Ust, LL_NUMERIC
675             LL1.LlDefineFieldExt "Re_Betrag", Betrag, LL_NUMERIC
680             LL1.LlDefineFieldExt "Re_SummeSteuerPflWrg", SteuerPflWrg, LL_NUMERIC
685             LL1.LlDefineFieldExt "Re_SummeSteuerFrWrg", SteuerFrWrg, LL_NUMERIC
690             LL1.LlDefineFieldExt "Re_UStWrg", UStWrg, LL_NUMERIC
695             LL1.LlDefineFieldExt "Re_BetragWrg", BetragWrg, LL_NUMERIC

                'CSBmk <STEUER-TEXTE>
700             objERechnung.colSteuerTexte.Clear

705             frm.fpSpread1(2).GetText 4, 4, SteuerText                       'HW 05.11.2010 Ver.: 5.4.122 - Ver.: 6.1.101
710             LL1.LlDefineFieldExt "Re_SteuerText", SteuerText, LL_TEXT

715             rec1100Texte.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Sort <= 7", gConn, adOpenStatic, adLockReadOnly 'HW 09.07.2012 Ver.: 6.1.114
     
720             If rec1100Texte.RecordCount > 0 Then

725                 Do While Not rec1100Texte.EOF

730                     LL1.LlDefineFieldExt "Steuertext" & rec1100Texte!Sort, "" & rec1100Texte!text, LL_TEXT
                        
735                     If Not objERechnung.colSteuerTexte.ContainsKey(CStr(rec1100Texte!Sort)) Then
                        
740                         Call objERechnung.colSteuerTexte.Add("" & rec1100Texte!text, CStr(rec1100Texte!Sort)) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung

                        End If
                        
745                     rec1100Texte.MoveNext

                    Loop

                Else
                
750                 LL1.LlDefineFieldExt "Steuertext", "", LL_TEXT

                End If
                
755             rec1100Texte.Close
760             Set rec1100Texte = Nothing
                
765             LL1.LlDefineFieldExt "Steuertext", "" & gstrSteuerText, LL_TEXT 'HW 05.07.2012  Ver.: 6.1.129
770             LL1.LlDefineFieldExt "SteuerSchl", intSteuerTyp, LL_NUMERIC
                                
775             objERechnung.SteuerText = GetSteuerText(intSteuerTyp, SteuerFr, gstrSteuerText, objERechnung.colSteuerTexte.GetItem("2"), objERechnung.colSteuerTexte.GetItem("4"), objERechnung.colSteuerTexte.GetItem("6")) 'DF 30.08.2024 , Ver.: 6.7.101, E-Rechnung
                
                'CSBmk <ANLAGEN-TEXT>
780             LL1.LlDefineFieldExt "AnlagenText", "", LL_TEXT                 'HW 01.07.2013
785             LL1.LlDefineVariableExt "AnlagenText", "", LL_TEXT              'HW 01.07.2013
                
                'CSBmk <VON / BIS DATUM>
790             If IsDate(rsH!vonDatum) Then
795                 LL1.LlDefineVariableExt "Kd_VonDatum", rsH!vonDatum, LL_TEXT
                Else
800                 LL1.LlDefineVariableExt "Kd_VonDatum", "", LL_TEXT
                End If

805             If IsDate(rsH!bisDatum) Then
810                 LL1.LlDefineVariableExt "Kd_BisDatum", rsH!bisDatum, LL_TEXT
                Else
815                 LL1.LlDefineVariableExt "Kd_BisDatum", "", LL_TEXT
                End If
    
820             j = 7
                
                'CSBmk <RABATT SICHTBAR?>
825             sql = "SELECT Max([Rabatt]) AS MaxRabatt "
830             sql = sql & "FROM [2800_Folge" & TmpZusatz & "] WHERE BelegID = " & BelegID
835             RS1.Open sql, gConn, adOpenStatic, adLockReadOnly

840             LL1.LlDefineFieldExt "RabattVisible", RS1!MaxRabatt, LL_NUMERIC
845             RS1.Close
850             j = 8

                'CSBmk <BELEG MIT LIEFERSCHIENARTIKEL?>
855             sql = "SELECT TOP 1 SatzTyp "                                   'MW 26.04.05
860             sql = sql & "FROM [2800_Folge" & TmpZusatz & "] WHERE SatzTyp = 'L' AND BelegID = " & BelegID
865             RS1.Open sql, gConn, adOpenStatic, adLockReadOnly

870             LL1.LlDefineFieldExt "LSArtikel", RS1.RecordCount, LL_NUMERIC
875             RS1.Close

880             LL1.LlDefineFieldExt "KostenstellenDruck", Abs(GbKostenstellenPflicht), LL_NUMERIC 'MW 28.12.07

885             j = 9

            Else
            
890             Msg = True

            End If

        Else
        
895         Msg = True

        End If

900     j = 10
  
905     If Msg = False Then

910         Formular = FormularPfad("SP52800.lst")

            'glRet = LL1.LlPreviewSetTempPath(ArbeitsplatzPfad & "\1\")
            
915         j = 11
    
            'Logik aus 55710 um Belege zu archivieren -> Schleife 2 mal: 1 Vorschau mit LL_PRINT_STORAGE (Datei ins Archiv kopieren), 2 Drucken.
            'glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_STORAGE, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "printing list")
            'ArbeitsplatzPfad
            
920         If Mode < 2 Then
                
                'DRUCK
925             glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_NORMAL, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Druck")

930             j = 12

                '<Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
                '# Druck in die LL Datei.
935         ElseIf Mode = 4 Then

                'ABLAGE
940             glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Druck")

            Else
                '</Added by: IL at: 07.11.2024, Ver.: 6.7.101 >
                    
                'VORSCHAU

945             If Save Then

                    '895                 glRet = LL1.LlPrintStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW)

950                 glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Archivierung")
                    
955                 j = 13

                Else
                
960                 glRet = LL1.LlPrintWithBoxStart(LL_PROJECT_LIST, Formular, LL_PRINT_PREVIEW, LL_BOXTYPE_BRIDGEMETER, frm.hwnd, "Vorschau")

965                 j = 14

                End If

            End If

970         If glRet < 0 Then

975             j = 15

980             GoTo Fehler

            End If
            
            'CSBmk <ANZAHL DER KOPIEN>
985         If Not Save And Mode < 2 Then
990             glRet = LL1.LlPrintSetOption(LL_PRNOPT_COPIES, CLng(GetSetting("SP50000", "SP52800", "SP52830_PRNOPT_COPIES", "1")))
995             glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)
            Else
1000            glRet = LL1.LlPrintSetOption(LL_PRNOPT_COPIES, 1)
1005            glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)
            End If

1010        j = 16
            
1015        If SammelDruck Then
1020            DruckerDialog = CBool(GetSetting("SP50000", "SP52800", "SP52850DruckerDialog", "-1"))
            Else
1025            DruckerDialog = CBool(GetSetting("SP50000", "SP52800", "SP52830DruckerDialog", "-1"))
            End If
    
1030        Call LL18PositionierungFormular(LL1, objDruckOptionen.FormularNr)                   'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

1035        If objDruckOptionen.CurrentKurzrechnung Then                        'DH, 11.07.2013, 6.2.100, Wenn so eingestellt, auf den Folgeseiten nicht den gesamten Kopf drucken

1040            Call LL18ShortHeader(LL1)

            End If

1045        If Mode = 1 Then                                                    'Nur beim richtigen Druck und nicht bei der Wiederholung

1050            Call LL18SetCopies(LL1, 1, objDruckOptionen.FormularNr, "SP52800")  'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

            End If
            
            'CSBmk <DRUCKERAUSWAHL-DIALOG>
1055        If DruckerDialog = True Then                                        'Druckdialog

1060            If Not Save Then

1065                If Mode <> 4 Then glRet = LL1.LlPrintOptionsDialog(frm.hwnd, "Drucker")

1070                Select Case glRet                                           '<Added by: DFiebach at: 27.01.2022, Ver.: 6.6.112
                    
                        Case Is < 0
                            
1075                        Protokoll iAppend, "##### -> FEHLER : LL, NUMMER: " & CStr(glRet) & ", SF-" & strLogBelegArt & ", BelegNr = " & BelegNr & ", BelegID =" & BelegID & ""
                            
1080                        GetMsgFromLLErrorCode glRet
                            
1085                        LL1.LlPrintEnd 0

1090                        LLPrintListe = glRet
                        
1095                        If Not SammelDruck Then

1100                            If Mode = 1 Or Mode = 4 Then                   'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

1105                                j = 17

1110                                rsH!Druck = 0
1115                                rsH!belegDatum = Null                       'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
1120                                rsH.Update

1125                                j = 18

                                End If
                            
                            End If
                            
1130                        glRet = LL1.LlPreviewDeleteFiles(ArbeitsplatzPfad & "\SP52800.LL", "") 'DF 08.03.2023 , Ver.: 6.6.118 : Vorschau-Datei beim Abbruch des DruckerAuswahl-Dialoges löschen.
                            
                            Exit Function
                             
                    End Select

1135                If Mode < 2 Then SaveSetting "SP50000", "SP52800", "SP52830_PRNOPT_COPIES", LL1.LlPrintGetOption(LL_PRNOPT_COPIES)

                End If

1140            j = 16

            End If
    
            'Nach Combit ist es unbedingt notwendig, die von LlPrintSetOption gesetzte Kopienanzahl
            'durch den Aufruf von LL_PRNOPT_COPIES_SUPPORTED zu bestätigen.
1145        glRet = LL1.LlPrintGetOption(LL_PRNOPT_COPIES_SUPPORTED)
            
            '######## 1. HIER BELEG-NR ZIEHEN
            
            'CSBmk <BELEG-NR VERGABE>
1150        If Mode = 1 Or Mode = 4 Then                                       'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

1155            If objDruckOptionen.CurrentBelegNr <> "" And objDruckOptionen.CurrentBelegNr <> "0" Then

1160                If Mode = 2 And tmp Then
1165                    rsH!BelegNr = 0
       
                    Else
1170                    rsH!BelegNr = objDruckOptionen.CurrentBelegNr
      
                    End If
                    
1175                Protokoll iAppend, "Beleg-Nummer aus DruckOptionen uebernommen -> BelegID: " & BelegID & ", BelegNr:" & rsH!BelegNr & ""
                    
1180                rsH.Update

                Else
                    
                    '<Modified by: DFiebach at 01.04.2019, Ver.: 6.5.110 >
                    ' # Überprüfung auf Vorlagen hinzugefügt
1185                If rsH!ZwAblage = 0 Then
                    
1190                    BelegNr = GetBelegNr(GintBelegNrKreisNr, BelegID, GintBelegArt, lngBelegNrKres, False, programmNr) 'DF 14.11.2024 , Ver.: 6.7.101 : GintBelegArt + 8 -> GintBelegNrKreisNr. GintBelegArt passt hier nicht mehr wegen der neune BelegArten (Ang, AufBest), da diese andrere Nr-Bereich in NrKresen haben.

                        '<Added by: DFiebach at: 01.04.2019, Ver.: 6.5.110 >
1195                    If BelegNr > 0 Then

1200                        gLngBelegNr = BelegNr                               'Added by: GW at: 03.04.2019, Ver.: 6.5.110
                 
1205                        rsH!BelegNrKReis = lngBelegNrKres
1210                        rsH!BelegNr = BelegNr
1215                        rsH.Update

                        Else
                    
1220                        rsH!Druck = 0
1225                        rsH!belegDatum = Null                               'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
1230                        rsH.Update
                    
1235                        Protokoll iAppend, "##### -> ABBRUCH DURCH BENUTZER, BELEG-NR NICHT FORTLAUFEND, SF-" & strLogBelegArt & ", BelegNr = " & BelegNr & ", BelegID =" & BelegID & ""
                        
1240                        LL1.LlPrintEnd 0

1245                        LLPrintListe = LL_ERR_USER_ABORTED
                            
1250                        glRet = LL1.LlPreviewDeleteFiles(ArbeitsplatzPfad & "\SP52800.LL", "")
                            
                            Exit Function
                        
                        End If

                        '</Added by: DFiebach at: 01.04.2019, Ver.: 6.5.110 >
                        
                    Else
                       
1255                    BelegNr = 0
                       
                    End If

                    '</Modified by: DFiebach at 01.04.2019, Ver.: 6.5.110 >
                    
                End If

1260            Call LLDefineVariablen(LL1, rsH, "Kd_")                                                                   'Deklarationen

                '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
                '#  ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
1265            If rsH!MwSt = 0 And CInt(vValue) <> 0 Then

1270                rsH!MwSt = dblUstSatz

1275                rsH.Update

1280                Call LLDefineFelder(LL1, rsH, "Kd_")                                                                      'Deklarationen

1285                rsH!MwSt = 0

1290                rsH.Update
               
                Else
                
1295                Call LLDefineFelder(LL1, rsH, "Kd_")

                End If

                '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >                                                                 'Deklarationen

            End If

            'CSBmk <BARCODE DEFINIEREN>
1300        barcodeDaten.Seperator = ";"
1305        barcodeDaten.BelegNr = rsH!BelegNr
1310        barcodeDaten.belegDatum = IIf(IsNull(rsH!belegDatum), "", rsH!belegDatum)

1315        barcodeDaten.Name1 = "" & rsH.Fields("Name1").value
1320        barcodeDaten.Name2 = "" & rsH.Fields("Name2").value
1325        barcodeDaten.Adresse = "" & rsH.Fields("Straße").value
1330        barcodeDaten.Lkz = "" & rsH.Fields("Lkz").value
1335        barcodeDaten.Plz = "" & rsH.Fields("Plz").value
1340        barcodeDaten.Ort = "" & rsH.Fields("Ort").value
1345        barcodeDaten.ORTSTEIL = rsH.Fields("Ortsteil").value

1350        Call LL18DefineBarcode(LL1, barcodeDaten, objDruckOptionen.FormularNr, rsH.Fields("MCode").value) 'Barcode im Formular definieren       'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

1355        Screen.MousePointer = 11

            'CSBmk <VARIABLEN DRUCKEN>
1360        glRet = LL1.LLPrint
            
1365        Protokoll iAppend, "Beleg-Positionen werden gedruckt -> BelegID: " & BelegID & ""
            
1370        While Not rs.EOF                                                    'Solange das Ende der Posten-Tabelle nicht erreicht ist...

1375            j = 19

1380            DoEvents

                'Prozentbalken setzen
1385            PercentPosition = 100 * rs.AbsolutePosition / rs.RecordCount
1390            glRet = LL1.LlPrintSetBoxText("Drucken", PercentPosition)
                
                'Datensatzfelder der Liste bekanntmachen.                       'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
1395            If Trim(rs!Einheit) = "%" Then
                    
                    'ORIG
                    'ZwSumme = ZwSumme + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                    
                    'DF 03.03.2025 , Ver.: 6.7.106: NEU -> rs!Menge / 100 führte zum falschen Ergebnis, -> rs!EPreis / 100 analog zum fpSread-Formel.
1400                ZwSumme = ZwSumme + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)

                Else
                
1405                ZwSumme = ZwSumme + RundenMitVz((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)

                End If

1410            LL1.LlDefineFieldExt "Re_EPreisDezStellen", postCommaPreis, LL_NUMERIC      'DH, 27.10.2017, 6.5.101, Einstellung aus den Systemparametern uebergeben

1415            LL1.LlDefineFieldExt "Re_ZwSumme", ZwSumme, LL_NUMERIC

1420            Call LLDefineFelder(LL1, rs, "Re_")

                'Seitenumbruch
1425            If rs!SatzTyp = "S" Then

1430                Seite = Seite + 1

1435                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1440                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

1445                glRet = LL1.LLPrint

1450                j = 20

                End If
                
                'Felder drucken und wenn Seitenumbruch erfolgt ist,
                'Variablen und Felder erneut drucken
    
                'HW 01.07.2013 Wenn AnlageText eingestellt wurde!
                '##################################################
                '                     ANLAGENTEXT
                '##################################################
1455            LL1.LlDefineFieldExt "Anlage", "", LL_TEXT
1460            LL1.LlDefineVariableExt "Anlage", "", LL_TEXT
                '##################################################

1465            While LL1.LlPrintFields = LL_WRN_REPEAT_DATA

1470                Seite = Seite + 1

1475                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1480                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

1485                glRet = LL1.LLPrint

                Wend

1490            If rs!SatzTyp = "Z" Then ZwSumme = 0
1495            rs.MoveNext

            Wend

1500        LL1.LlDefineFieldExt "LetzteSeite", 1, LL_NUMERIC
            
            'DF 04.06.2020 , Ver.: 6.6.102 : Zahlungs-Konditionen auch bei Gutschriften drucken
            
            'CSBmk <ZAHLUNGS-KONDITIONEN>
            
            '<Modified by: IL at 15.10.2024, Ver.: 6.7.101 >
            '#   BelegDatum -------> objDruckOptionen.CurrentBelegDatum.
1505        strZahlungsText = ZahlungsZiel(objDruckOptionen.CurrentBelegDatum, Betrag, Waehrung, Skonto, SkontoTage, nettoTage, objDruckOptionen.CurrentValutaDatum)
1510        strZahlungsTextNetto = ZahlungsZielNetto(objDruckOptionen.CurrentBelegDatum, Betrag, Waehrung, Skonto, SkontoTage, nettoTage, objDruckOptionen.CurrentValutaDatum)
            '</Modified by: IL at 15.10.2024, Ver.: 6.7.101 >
            
1515        LL1.LlDefineFieldExt "ZahlungsZiel", strZahlungsText, LL_TEXT      'DH, 16.02.2015, 6.4.103, ValutaDatum aus den DruckOptionen uebernehmen
1520        LL1.LlDefineFieldExt "ZahlungsZielNetto", strZahlungsTextNetto, LL_TEXT
            
1525        If objERechnung Is Nothing Then Set objERechnung = New clsERechnung      'DF 28.08.2024 , Ver.: 6.7.101
1530        objERechnung.ZHinweisNetto = strZahlungsTextNetto
1535        objERechnung.ZHinweisBrutto = strZahlungsText
            
            'CSBmk <STEUER-CODE SPEICHERN HAUPT UND FOLGE>
            
            '<Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >

1540        If Mode = 1 Or Mode = 4 Then                                                               'Nur beim Druck.

1545            rsH!ERechnungArt = modERechnung.GetERechnungTypeValueForDB(gEnmKudnenERechnungType)    'DF 23.07.2024 , Ver.: 6.7.101
1550            rsH!StCode = strStCodeH
            
1555            rsH.Update

1560            Call SetStCode(E_DATATYPE.Sonderfaktura_Rechnung, 1, rsH!BelegID, intSteuerTextLkz, intSteuerTyp, tmp, GintBelegArt) ' An der Stelle wird zw. SF-RNG und -GUT nicht unterschieden, da beide in der gelichen Tabelle gespeichert werden.

            End If

            '</Added by: DFiebach at: 11.07.2024, Ver.: 6.7.101 >
            
            'Tabellen-Ausdruck beenden
            Do
            
1565            glRet = LL1.LlPrintFieldsEnd()

1570            If glRet = LL_WRN_REPEAT_DATA Then

1575                Seite = Seite + 1

1580                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1585                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

                    'Neue Seite
1590                LL1.LLPrint

1595                j = 21

                End If

1600        Loop Until glRet <> LL_WRN_REPEAT_DATA
  
            'HW 01.07.2013 Wenn AnlageText eingestellt wurde!
            '##################################################
            
            'CSBmk <ANLAGETEXT>
            
1605        LL1.LlPrintResetProjectState                                                                              'Druck Zurücksetzen damit Lastpage und andere diverse Funktionen für LL18 gültig sind!

1610        Call LL18GestaltungFormular(LL1, objDruckOptionen.FormularNr, "" & rsH.Fields("MCode").value, MandantArr(1), , , ArchivierungsModus)          'IL 25.10.2024 , Ver.: 6.7.101 :  rsH!art + 35 ------> objDruckOptionen.FormularNr

1615        Call LLDefineVariablen(LL1, rsH, "Kd_")                                                                   'Deklarationen

            '<Modified by: IL at 7.26.2024, Ver.: 6.7.101 >
            '# ersetzen den erforderlichen Steuersatz im Datensatz und machen dann 0 zurück
1620        If rsH!MwSt = 0 And CInt(vValue) <> 0 Then

1625            rsH!MwSt = dblUstSatz

1630            rsH.Update

1635            Call LLDefineFelder(LL1, rsH, "Kd_")                                                                  'Deklarationen

1640            rsH!MwSt = 0

1645            rsH.Update
           
            Else
            
1650            Call LLDefineFelder(LL1, rsH, "Kd_")

            End If

            '</Modified by: IL at 7.26.2024, Ver.: 6.7.101 >

1655        If SPLL8.bAnlageAktiv And Trim(SPLL8.strAnlageText) <> "" Then                                            'Wenn ein Anlagetext eingestellt ist!
    
1660            LL1.LlDefineFieldExt "Anlage", SPLL8.strAnlageText, LL_TEXT
1665            LL1.LlDefineVariableExt "Anlage", SPLL8.strAnlageText, LL_TEXT

1670            Seite = Seite + 1

1675            LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1680            LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC

1685            glRet = LL1.LLPrint

1690            While LL1.LlPrintFields = LL_WRN_REPEAT_DATA

1695                Seite = Seite + 1

1700                LL1.LlDefineVariableExt "Seite", Seite, LL_NUMERIC
1705                LL1.LlDefineFieldExt "Seite", Seite, LL_NUMERIC
             
1710                glRet = LL1.LLPrint

                Wend

1715            LL1.LlDefineFieldExt "LetzteSeite", Seite, LL_NUMERIC

            End If

            '##################################################
     
            'CSBmk <DRUCK BEENDEN>
1720        glRet = LL1.LlPrintEnd(0)

1725        j = 22
     
1730        If (Mode = 1 Or Mode = 4) And Not tmp Then                         'IL 07.11.2024 , Ver.: 6.7.101 : Ablage hinzugefügt

1735            Protokoll iAppend, ">EINZELDRUCK BEENDET -> BelegID: " & BelegID & " SteuerPfl: " & SteuerPfl & " SteuerFr: " & SteuerFr & " Ust: " & Ust & " Betrag: " & Betrag

            End If
            
            'CSBmk <ÜBERGABE AN RAB>
            
            'CSBmk <PDF-ARCHIVIERUNG>
            '
            'Beim Preview-Druck Preview anzeigen und dann Preview-Datei (.LL) löschen
            'HW 23.04.2014 Ver.: 6.2.105 3 übergeben! fürs archivieren!
1740        If Mode > 1 And Mode <> 4 Then 'PrintMode = LL_PRINT_PREVIEW       auser Ablage

1745            If Save Then

                    '<Modified by: IL at 09.10.2024, Ver.: 6.7.101 >
1750                Select Case BelegArt

                        Case 0 'RECHNUNG
    
1755                        If rsH!ZwAblage = 0 Then

1760                            Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""

1765                            Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Rechnung, GesamtIstBrutto, CCur(SteuerPfl))       'Belegdaten in die Ausgangsbuch Tabellen schreiben

1770                            Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""

1775                            Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110

1780                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

1785                            Call ArchivierenPDF(LL1, "SFR", BelegNr, rsH, rs)

1790                            Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""

                            End If
    
1795                    Case 1 'GUTSCHRIFT
    
1800                        If rsH!ZwAblage = 0 Then

1805                            Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""

1810                            Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Gutschrift, GesamtIstBrutto, CCur(SteuerPfl))      'Belegdaten in die Ausgangsbuch Tabellen schreiben

1815                            Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""

1820                            Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110

1825                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

1830                            Call ArchivierenPDF(LL1, "SFG", BelegNr, rsH, rs)

1835                            Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""

                            End If
    
1840                    Case 2 'ANGEBOT
    
1845                        If rsH!ZwAblage = 0 Then

1850                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

1855                            Call ArchivierenPDF(LL1, "SFA", BelegNr, rsH, rs)

1860                            Protokoll iAppend, ">UEBERAGABEN AN ARCHIV BEENDET -> BelegID: " & BelegID & ""

                            End If
    
1865                    Case 3 'AUFTRAGSBESTETIGUNG
                            
1870                        If rsH!ZwAblage = 0 Then
                            
1875                            Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""

1880                            Call ArchivierenPDF(LL1, "SFB", BelegNr, rsH, rs)

1885                            Protokoll iAppend, ">UEBERAGABEN AN ARCHIV BEENDET -> BelegID: " & BelegID & ""

                            End If

                    End Select

                    'Orig: 1695                If BelegArt = 0 Then
                    '
                    '                        'RECHNUNG
                    '
                    '1700                    If rsH!ZwAblage = 0 Then
                    '
                    '1705                        Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""
                    '
                    '1710                        Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Rechnung, GesamtIstBrutto, CCur(SteuerPfl))       'Belegdaten in die Ausgangsbuch Tabellen schreiben
                    '
                    '1715                        Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""
                    '
                    '1720                        Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110
                    '
                    '1725                        Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""
                    '
                    '1730                        Call ArchivierenPDF(LL1, "SFR", BelegNr, rsH, rs)
                    '
                    '1735                        Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""
                    '
                    '                        End If
                    '
                    '                    Else
                    '
                    '                        'GUTSCHRIFT
                    '
                    '1740                    If rsH!ZwAblage = 0 Then
                    '
                    '1745                        Protokoll iAppend, ">BELEG AN RAB uebergeben -> BelegID: " & BelegID & ""
                    '
                    '1750                        Call BelegAnAusgangsbuch(BelegID, E_DATATYPE.Sonderfaktura_Gutschrift, GesamtIstBrutto, CCur(SteuerPfl))      'Belegdaten in die Ausgangsbuch Tabellen schreiben
                    '
                    '1755                        Protokoll iAppend, ">BELEG AN OP uebergeben -> BelegID: " & BelegID & ""
                    '
                    '1760                        Call BelegAnOpUebergeben(BelegID)                   'GW 12.3.2021, Ver.: 6.6.110
                    '
                    '1765                        Protokoll iAppend, ">BELEG AN ARCHIV uebergeben -> BelegID: " & BelegID & ""
                    '
                    '1770                        Call ArchivierenPDF(LL1, "SFG", BelegNr, rsH, rs)
                    '
                    '1775                        Protokoll iAppend, ">UEBERAGABEN AN RAB/OP/FIBU BEENDET -> BelegID: " & BelegID & ""
                    '
                    '                        End If
                    '
                    '                    End If

                    Dim currentDocType As E_DATATYPE
                    
1890                Select Case BelegArt
                    
                        Case 0
                        
1895                        currentDocType = E_DATATYPE.Sonderfaktura_Rechnung
                        
1900                    Case 1
                        
1905                        currentDocType = E_DATATYPE.Sonderfaktura_Gutschrift
                        
1910                    Case 2
                        
1915                        currentDocType = E_DATATYPE.Sonderfaktura_Angebot
                        
1920                    Case 3
                        
1925                        currentDocType = E_DATATYPE.Sonderfaktura_Auftragsbestetigung
                    
                    End Select
    
                    'Orig:1830           If BelegArt = 0 Then
                    '1835                    currentDocType = E_DATATYPE.Sonderfaktura_Rechnung
                    '                    Else
                    '1840                    currentDocType = E_DATATYPE.Sonderfaktura_Gutschrift
                    '                    End If
                    
                    '</Modified by: >IL at 09.10.2024, Ver.: 6.7.101 >
                    
                    'CSBmk <EMAIL-VERSAND>
1930                If emailActivated(rsH.Fields("MCode").value, CInt(currentDocType)) Then           'DH, 21.12.2015, 6.4.114, Wenn der eMail-Versand aktiviert ist (Mandanten-/Kundenstamm)
                        
1935                    Protokoll iAppend, ">EMAIL VERSAND -> BelegID: " & BelegID & ""
                        
1940                    If objEmailSending Is Nothing Then Set objEmailSending = New clsEmailSending   'Modified by: GW at 21.02.2020, Ver.: GOBD_EMAIL
    
1945                    Set idCollection = New Collection
1950                    idCollection.Add BelegID

1955                    If UCase(frm.name) = "FRMSP52831" Then                                        'Einzeldruck
1960                        Call objEmailSending.StartEmailSending(frm.frmParent.cReSize.CurrScaleFactorHeight, frm.frmParent.cReSize.CurrScaleFactorWidth, voll_automatik, currentDocType, idCollection)
                        Else                                                                          'Sammeldruck
1965                        Call objEmailSending.StartEmailSending(frm.cReSize.CurrScaleFactorHeight, frm.cReSize.CurrScaleFactorWidth, voll_automatik, currentDocType, idCollection)
                        End If

                    End If

                Else
                                                
                    'CSBmk <VORSCHAU ANZEIGEN>
1970                glRet = LL1.LlPreviewDisplay(ArbeitsplatzPfad & "\SP52800.LL", "", frm.hwnd)

1975                j = 23

                End If
                
                'CSBmk <TEMP DATEI LÖSCHEN>
1980            glRet = LL1.LlPreviewDeleteFiles(ArbeitsplatzPfad & "\SP52800.LL", "")

1985            j = 24

            End If
    
1990        rs.MoveFirst

1995        Screen.MousePointer = 0

        End If
        
2000    rsH.Close

2005    Set rsH = Nothing

        'DH, 11.07.2013, Nach dem Druck/Druckwiederholung muessen BelegNr und -Datum gesperrt werden
2010    If Mode = 0 Or Mode = 1 Or Mode = 4 Then                               'IL 07.11.2024 , Ver.: 6.7.101 : Ablage hinzugefügt

2015        objDruckOptionen.EnableBelegDatum = False
2020        objDruckOptionen.EnableBelegNr = False
2025        objDruckOptionen.EnableValutaDatum = False                          'DH, 17.02.2015, 6.4.103, Das neue Feld Valuta Datum muss bei der Druckwiederholung auch gesperrt werden

2030        printDone = True

        End If
        
        'CSBmk <LOG ENDE>
2035    Select Case Mode
        
            Case 0         'DRUCK WIEDERHOLUNG
                 
2040            Protokoll iAppend, ">DRUCK ENDE (MODUS: DRUCK WIEDERHOLUNG) -> BelegID: " & BelegID & ""
                 
2045        Case 1         'DRUCK
            
2050            Protokoll iAppend, ">DRUCK ENDE (MODUS: DRUCK) -> BelegID: " & BelegID & ""
            
2055        Case 2         'VORSCHAU
            
2060            Protokoll iAppend, ">DRUCK ENDE (MODUS: VORSCHAU) -> BelegID: " & BelegID & ""
            
2065        Case 3         'ARCHIVIERUNG
        
2070            Protokoll iAppend, ">DRUCK ENDE (MODUS: ARCHIVIERUNG) -> BelegID: " & BelegID & ""

2075        Case 4         'ABLAGE                                             'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

2080            Protokoll iAppend, ">DRUCK ENDE (MODUS: ABLAGE) -> BelegID: " & BelegID & ""
            
        End Select
        
        Exit Function
  
Fehler:
   
2085    LLPrintListe = Err.number

2090    If Mode = 1 Or Mode = 4 Then                                           'IL 07.11.2024 , Ver.: 6.7.101 :  Ablage hinzugefügt

2095        If Not SammelDruck Then

2100            If Not rsH Is Nothing Then

2105                If rsH.RecordCount > 0 Then
2110                    rsH!Druck = 0
2115                    rsH!belegDatum = Null                                   'DF 13.02.2019 , Ver.: 6.5.109 : BelegDatum nur wenn Beleg erfolgreich gedruckt wurde
2120                    rsH.Update
                    End If

                End If

            End If
            
        End If

2125    If glRet <> 0 Then
2130        Call FehlerErklärung("SP52800B", "LLPrintListe LL-Fehler: " & glRet & ", j = " & j)
        Else
2135        Call FehlerErklärung("SP52800B", "LLPrintListe")
        End If

End Function

Public Sub Main()

        On Error GoTo Fehler
  
100     GsAnwenderNr = GetSetting("SP50000", "Settings", "AnwnderNr", "")
105     GsTitel = GetSetting("SP50000", "Settings", "Titel", "")
110     GdtDatum = GetSetting("SP50000", "Settings", "ArbeitsDatum", Date)
115     GsUser = GetSetting("SP50000", "Settings", "User", "")
120     GlngArbeitsplatz = CLng(GetSetting("SP50000", "Settings", "Arbeitsplatz", "0"))

125     If GsStartProgNrPerMsg <> "" Then
130         programmNr = GsStartProgNrPerMsg
        Else
135         programmNr = GetSetting("SP50000", "Settings", "StartProgrammNr")
        End If
        
140     gstrProgNrGL = programmNr
        
145     GsHauptPfad = GetSetting("SP50000", "Settings", "Pfad", PfadZrck(App.Path))

150     Call SetGsHauptPfadLokal                                                '17.10.2014
        
155     If GsTitel <> "" Then                                                   'Das Programm wird gestartet nur dann, wenn SP50000 aktiv ist.

160         If GsSprache = "" Then Sprache
    
165         DesignerFrei                                                        'Designer freischalten?

170         ArchivFrei                                                          'Archivierung freischalten?

175         Load frmMsg                                                         'Sorgen dafür das die Form nur 1 mal geladen wird.

180         Protokoll iOpen, vbCrLf & "*****PROGRAMM START*****. Mandant: " & GsAnwenderNr, GsHauptPfad & "log\", "SP62800.LOG" 'DF 16.01.2019 , Ver.: 6.5.109 :LOG-Dateinamen angepasst sp52800.log -> SP62800.LOG

185         KostenstellenPflicht
    
            'Mandantendatensatz
190         If GmandantRS Is Nothing Then

195             Set GmandantRS = New ADODB.Recordset
200             OPEN_gConn
205             GmandantRS.Open "SELECT * FROM [1100_Mandant]", gConn, adOpenStatic, adLockReadOnly

            End If

210         MandantArr(1) = GmandantRS.Fields("DruckOhneBez").value       'DH, 27.06.2017, 6.4.126, Ersetzt den Aufruf von mandant()

215         Protokoll iAppend, "Programm-Nr: " & programmNr & "  -> " & Now
    
220         If GDBlog Is Nothing Then
225             Set GDBlog = DBEngine.OpenDatabase(GsHauptPfadLokal & "prm\SP51000.log")
            End If
    
230         If GDBprm Is Nothing Then
235             Set GDBprm = DBEngine.OpenDatabase(GsHauptPfadLokal & "prm\SP500DE.prm")
            End If
    
240         If GDBdef Is Nothing Then
245             Set GDBdef = DBEngine.OpenDatabase(GsHauptPfadLokal & "exe\SP50000.def")
            End If

250         Call LoadSystemparameter
    
255         Set g_objCal = New clsKalender

260         Select Case programmNr

                Case "281" 'Textstamm

                    '350         frmSP52810.Show
                    
265             Case "282" 'Artikelstamm

270                 frmSP52820.Show

275             Case "283" 'Sonderfaktura-Rechnung

280                 GintBelegArt = 0
285                 GintBelegNrKreisNr = C_INT_NRKREIS_NR_SF_RECHNUNG           'DF 14.11.2024 , Ver.: 6.7.101
                    
290                 Set frmRechnung = New frmSP52830
295                 frmRechnung.Show

300             Case "284" 'Sonderfaktura-Gutschrift

305                 GintBelegArt = 1
310                 GintBelegNrKreisNr = C_INT_NRKREIS_NR_SF_GUTSCHRIFT         'DF 14.11.2024 , Ver.: 6.7.101
                    
315                 Set frmGutschrift = New frmSP52830
320                 frmGutschrift.Show

325             Case "285" 'Sammeldruck-Rechnung

330                 GintBelegArt = 0
335                 GintBelegNrKreisNr = C_INT_NRKREIS_NR_SF_RECHNUNG           'DF 14.11.2024 , Ver.: 6.7.101
                    
340                 Set frmRechnungFakt = New frmSP52850
345                 frmRechnungFakt.Show

350             Case "286" 'Sammeldruck-Gutschrift

355                 GintBelegArt = 1
360                 GintBelegNrKreisNr = C_INT_NRKREIS_NR_SF_GUTSCHRIFT         'DF 14.11.2024 , Ver.: 6.7.101

365                 Set frmGutschriftFakt = New frmSP52850
370                 frmGutschriftFakt.Show

375             Case "288"  'Sonderfaktura-Angebot

380                 GintBelegArt = 2
385                 GintBelegNrKreisNr = C_INT_NRKREIS_NR_SF_ANGEBOT            'DF 14.11.2024 , Ver.: 6.7.101

390                 Set frmAngebot = New frmSP52830
395                 frmAngebot.Show

400             Case "289"  'Sonderfaktura-Auftragsbestetigung

405                 GintBelegArt = 3
410                 GintBelegNrKreisNr = C_INT_NRKREIS_NR_SF_AUFTRAGBEST        'DF 14.11.2024 , Ver.: 6.7.101

415                 Set frmAuftragsbest = New frmSP52830
420                 frmAuftragsbest.Show

            End Select

        Else

425         End

        End If

        Exit Sub

Fehler:

430     Select Case Err.number

            Case 401 'Ungebundenes Formular kann nicht angezeigt werden, während modales Formular angezeigt wird
435             MsgBox "Gleichzeitige Rechnungs- und Gutschrift-Erfassung ist nicht möglich.", vbInformation

440             If programmNr = "283" Then
445                 Set frmRechnung = Nothing
                Else
450                 Set frmGutschrift = Nothing
                End If

455         Case Else
460             Call FehlerErklärung("SP52800B", "Main")
        End Select

End Sub

Public Sub LLDesigner(frm As Form, _
                      LL1 As ListLabel.ListLabel, _
                      BelegID As Long, _
                      Index As Integer, _
                      Optional tmp As Boolean)

        On Error GoTo Fehler

        'Index = 0 Alle Mandanten; Index = 1 Aktueller Mandant

        Dim Formular   As String

        Dim rs         As ADODB.Recordset

        Dim rsMAXID    As ADODB.Recordset

        Dim rec1100    As ADODB.Recordset                                       'HW 09.07.2012 Ver.: 6.1.114

        Dim Msg        As Boolean

        Dim ret        As Long

        Dim TmpZusatz  As String

        Dim SteuerText As Variant                                               'HW 05.11.2010 Ver.: 5.4.122 - Ver.: 6.1.101
  
        Dim BelegIDMAX As Long 'HW 03.12.2015

100     If tmp Then
105         TmpZusatz = "Tmp"
        End If

110     Set rs = New ADODB.Recordset
115     Set rsMAXID = New ADODB.Recordset
120     Set rec1100 = New ADODB.Recordset

125     OPEN_gConn

130     If BelegID = 0 Then
135         rsMAXID.Open "SELECT MAX(BelegID) FROM [2800_Haupt]", gConn, adOpenStatic, adLockReadOnly

140         If rsMAXID.RecordCount > 0 Then
145             BelegIDMAX = rsMAXID(0).value
150             rs.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = " & rsMAXID(0).value, gConn, adOpenStatic, adLockReadOnly
            Else
155             BelegIDMAX = 0
160             rs.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = 0", gConn, adOpenStatic, adLockReadOnly
            End If

165         rsMAXID.Close
        Else
170         BelegIDMAX = BelegID
175         rs.Open "SELECT * FROM [2800_Haupt] WHERE BelegID = " & CStr(BelegIDMAX), gConn, adOpenStatic, adLockReadOnly
        End If

180     If rs.RecordCount > 0 Then

185         LL1.LlDefineVariableStart                                           'Variablenpuffer löschen.
190         LL1.LlDefineFieldStart                                              'Variablenpuffer löschen.
195         Formular = "SP52800.lst"

200         Call LL18GestaltungFormular(LL1, 35, "" & rs.Fields("MCode").value, MandantArr(1)) 'HW 30.03.2012 Ver.: 6.1.111
205         Call LLDefineVariablen(LL1, rs, "Kd_")
210         Call LLDefineFelder(LL1, rs, "Kd_")
215         Call LLDefineTexte(LL1)                                                 'DF 24.10.2024 , Ver.: 6.7.101

            'DH, 22.12.2015, 6.4.115, Beim Aufruf aus dem Sammeldruck existieren diese Steuerelemente nicht.
            '                         Darum einfach stets aus der Registry auslesen.
220         If BearbeiterDrucken Then
225             LL1.LlDefineVariableExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
230             LL1.LlDefineFieldExt "Bearbeiter_Drucken", "TRUE", LL_BOOLEAN
            Else
235             LL1.LlDefineVariableExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
240             LL1.LlDefineFieldExt "Bearbeiter_Drucken", "FALSE", LL_BOOLEAN
            End If
            
            '<Added by: GW at: 24.04.2019, Ver.: 6.5.111 >
245         If blnFolgeseitenKurzDrucken Then
250             LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN
255             LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "TRUE", LL_BOOLEAN
            Else
260             LL1.LlDefineVariableExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN
265             LL1.LlDefineFieldExt "FolgeseitenKurz_Drucken", "FALSE", LL_BOOLEAN
            End If

            '</Added by: GW at: 24.04.2019, Ver.: 6.5.111 >
    
            'HW 03.07.2013 Nach-Deklaration
            '################################################
270         LL1.LlDefineFieldExt "Kd_Tel", "", LL_TEXT
275         LL1.LlDefineVariableExt "Kd_Tel", "", LL_TEXT
    
280         LL1.LlDefineFieldExt "Kd_Email", "", LL_TEXT
285         LL1.LlDefineVariableExt "Kd_Email", "", LL_TEXT
    
290         LL1.LlDefineFieldExt "Kd_AnsprPartner", "", LL_TEXT
295         LL1.LlDefineVariableExt "Kd_AnsprPartner", "", LL_TEXT
    
300         LL1.LlDefineFieldExt "Kd_Fax", "", LL_TEXT
305         LL1.LlDefineVariableExt "Kd_Fax", "", LL_TEXT
            '################################################
    
310         LL1.LlDefineFieldExt "ProbeDruckText", ZusatzText(4, "55710"), LL_TEXT 'HW 16.10.2013
315         LL1.LlDefineFieldExt "ProbeDruck", 0, LL_NUMERIC                       'HW 16.10.2013
320         LL1.LlDefineVariableExt "ERechnungArt", 0, LL_NUMERIC                  'DF 04.11.2024 , Ver.: 6.7.101

325         Call DefineZusatztext(rs, LL1)                                         'MW 13.11.08 Ver.: 5.4.119 Zusatztext

330         LL1.LlDefineFieldExt "ZahlungsZiel", "Bei Zahlung bis...", LL_TEXT
335         LL1.LlDefineFieldExt "ZahlungsZielNetto", "Zahlung bis...", LL_TEXT
340         LL1.LlDefineFieldExt "RabattVisible", 1, LL_NUMERIC
345         LL1.LlDefineVariableExt "Seite", 1, LL_NUMERIC
350         LL1.LlDefineFieldExt "LetzteSeite", 1, LL_NUMERIC
    
355         LL1.LlDefineFieldExt "LSArtikel", 0, LL_NUMERIC                                    'MW 26.04.05
360         LL1.LlDefineFieldExt "KostenstellenDruck", Abs(GbKostenstellenPflicht), LL_NUMERIC 'MW 28.12.07
    
365         LL1.LlDefineFieldExt "Re_EPreisDezStellen", 2, LL_NUMERIC                          'MW 28.12.06 Ver.: 5.3.105
370         LL1.LlDefineFieldExt "Re_SummeSteuerPfl", 0, LL_NUMERIC
375         LL1.LlDefineFieldExt "Re_SummeSteuerFr", 0, LL_NUMERIC
380         LL1.LlDefineFieldExt "Re_USt", 0, LL_NUMERIC
385         LL1.LlDefineFieldExt "Re_Betrag", 0, LL_NUMERIC
    
390         LL1.LlDefineFieldExt "Re_SummeSteuerPflWrg", 0, LL_NUMERIC
395         LL1.LlDefineFieldExt "Re_SummeSteuerFrWrg", 0, LL_NUMERIC
400         LL1.LlDefineFieldExt "Re_UStWrg", 0, LL_NUMERIC
405         LL1.LlDefineFieldExt "Re_BetragWrg", 0, LL_NUMERIC
    
410         LL1.LlDefineVariableExt "Kd_VonDatum", "" & rs!vonDatum, LL_TEXT
415         LL1.LlDefineVariableExt "Kd_BisDatum", "" & rs!bisDatum, LL_TEXT
              
420         LL1.LlDefineFieldExt "Re_SteuerText", "(Laut Angaben des Beleg Empfängers)", LL_TEXT 'HW 05.11.2010 Ver.: 5.4.122 - Ver.: 6.1.101
    
            'HW 09.07.2012 Ver.: 6.1.114
            'SteuerTexte
            '#################################
425         rec1100.Open "SELECT * FROM [1100_Texte] WHERE Textart = 'Rng' AND Sort <= 7", gConn, adOpenStatic, adLockReadOnly

430         If rec1100.RecordCount > 0 Then

435             Do While Not rec1100.EOF

440                 LL1.LlDefineFieldExt "Steuertext" & rec1100!Sort, "" & rec1100!text, LL_TEXT

445                 rec1100.MoveNext

                Loop

            Else
            
450             LL1.LlDefineFieldExt "Steuertext", "", LL_TEXT

            End If

455         rec1100.Close

460         LL1.LlDefineFieldExt "Steuertext", "" & gstrSteuerText, LL_TEXT     'HW 05.07.2012  Ver.: 6.1.129
465         LL1.LlDefineFieldExt "SteuerSchl", intSteuerTyp, LL_NUMERIC
            '#################################

470         LL1.LlDefineFieldExt "AnlagenText", "", LL_TEXT                     'HW 01.07.2013
475         LL1.LlDefineVariableExt "AnlagenText", "", LL_TEXT                  'HW 01.07.2013
    
480         LL1.LlDefineVariableExt "BarcodeText", "", LL_TEXT                  'HW 01.07.2013
485         LL1.LlDefineFieldExt "BarcodeText", "", LL_TEXT                     'HW 01.07.2013
    
            'HW 03.12.20015 Ver.: 6.4.113 MAXID eingepflegt
            '*
            'Folge-Recordset
490         OPEN_gConn

495         If rs.state = adStateOpen Then rs.Close
500         rs.Open "SELECT * FROM [2800_Folge] WHERE BelegID = " & BelegIDMAX & " ORDER BY Nr", gConn, adOpenKeyset, adLockReadOnly

505         If rs.RecordCount > 0 Then
510             Call LLDefineFelder(LL1, rs, "Re_")
515             LL1.LlDefineFieldExt "Re_ZwSumme", 0, LL_NUMERIC

                'Designer starten
520             glRet = LL1.LlDefineLayout(frm.hwnd, "Liste", OBJECT_LIST, FormularBearbeiten(Formular, Index))
            Else
525             Msg = True
            End If

        Else
530         Msg = True
        End If

535     If rsMAXID.state = adStateOpen Then rsMAXID.Close
540     Set rsMAXID = Nothing
        
545     If rs.state = adStateOpen Then rs.Close
550     Set rs = Nothing
        
555     If Msg Then
560         Call msgText(1, 23, 0, 0, 0)
565         MsgBox GsMsgText(0), vbOKOnly + vbInformation
            'MsgBox "Es stehen noch keine Datensätze zu Verfügung.", vbInformation
        End If
  
        Exit Sub

Fehler:
570     Call FehlerErklärung("SP52800B", "LLDesigner")
End Sub

'DH, 27.10.2017, 6.5.101
'Laedt Systemparameter aus dem Mandantenstamm
Public Sub LoadSystemparameter()

        On Error GoTo Fehler

        '<Removed by: DFiebach at: 05.02.2019, Ver.: 6.5.109 >
        ' # Umgestellt auf globale SystemParameterGL
        '        Dim rs As ADODB.Recordset
        '
        '100     Set rs = New ADODB.Recordset
        '
        '105     rs.Open "SELECT SFDezStellen FROM [1100_SystemParameter]", gConn, adOpenStatic, adLockReadOnly
        '
        '110     If rs.RecordCount > 0 Then
        '
        '115         postCommaPreis = rs.Fields("SFDezSTellen").Value          'Verwendete Nachkommastellen in der Sonderfaktura
        '
        '        End If
        '
        '120     rs.Close
        '</Removed by: DFiebach at: 05.02.2019, Ver.: 6.5.109 >
        
        '<Added by: DFiebach at: 05.02.2019, Ver.: 6.5.109 >
100     Call SystemParameterGL
        
105     If g_strSysPrm(37) = "" Then g_strSysPrm(37) = "2"
        
110     postCommaPreis = CInt(g_strSysPrm(37))
            
        '</Added by: DFiebach at: 05.02.2019, Ver.: 6.5.109 >
        
        Exit Sub

Fehler:
115     Call FehlerErklärung("frmSP52820", "LoadSystemparameter()")

        '<Removed by: DFiebach at: 05.02.2019, Ver.: 6.5.109 >
        '115     If rs.State = adStateOpen Then rs.Close
        '</Removed by: DFiebach at: 05.02.2019, Ver.: 6.5.109 >
End Sub

Public Function IstBelegNrFrei(BelegNr As Long, _
                               BelegID As Long, _
                               Art As Integer) As Boolean

        On Error GoTo Fehler

        Dim sql As String

        Dim rs  As ADODB.Recordset
  
100     Set rs = New ADODB.Recordset
  
105     If BelegNr = 0 Then

110         IstBelegNrFrei = True

        Else
        
115         rs.Open "SELECT BelegNr FROM [2800_Haupt] WHERE [Storno] = '0' AND [ZwAblage] = 0 AND [BelegNr] = " & BelegNr & "  AND [BelegID] <> " & BelegID & "  AND [Art] = " & Art, gConn, adOpenStatic, adLockReadOnly

120         If rs.RecordCount = 0 Then IstBelegNrFrei = True

125         rs.Close

        End If
  
        Exit Function

Fehler:
130     Call FehlerErklärung("SP52800B", "IstBelegNrFrei()")
End Function

Public Sub ArchivierenPDF(LL1 As ListLabel.ListLabel, _
                          Optional Prefix As String, _
                          Optional BelegNr As Long, _
                          Optional HauptRS As ADODB.Recordset, _
                          Optional FolgeRS As ADODB.Recordset)

        On Error GoTo Fehler

        Dim SHFO        As SHFILEOPSTRUCT

        Dim Erfolg      As Long

        Dim ZielDatei   As String

        Dim Pfad        As String

        Dim DateiName   As String

        Dim i           As Integer

        Dim sql         As String

        Dim Transaktion As Boolean

        Dim hStg        As Long

        Dim lRet        As Long

        Dim AD          As SPLL8.ArchivDaten 'HW 07.06.2013
    
100     If FileExists(ArbeitsplatzPfad & "\SP52800.LL") Then

105         If Not HauptRS Is Nothing Then

110             If HauptRS.RecordCount > 0 Then

115                 HauptRS.MoveFirst

120                 Pfad = LLArchivPfad(HauptRS!belegDatum)

                Else
                
125                 Pfad = LLArchivPfad(Date)

                End If

            Else
            
130             Pfad = LLArchivPfad(Date)

            End If

            'HW 04.06.2013
            '##################################################
135         AD.BelegNr = "" & BelegNr

140         If Not IsNull(HauptRS) Then

145             If HauptRS.RecordCount > 0 Then                                                        'Wenn Keine Daten zur V3erfügung stehen kann nichtsweiter gespeichert werden!

150                 AD.BelegID = "" & HauptRS.Fields("BelegID").value
155                 AD.belegDatum = "" & HauptRS.Fields("Belegdatum").value
160                 AD.Lkz = "" & HauptRS.Fields("LKZ").value
165                 AD.name = "" & HauptRS.Fields("Name1").value & " " & HauptRS.Fields("Name2").value
170                 AD.MCode = "" & HauptRS.Fields("MCode").value
175                 AD.KtoNr = "" & HauptRS.Fields("KtoNr").value
180                 AD.EVers = ""
185                 AD.EDat = "NULL"

190                 If Trim(HauptRS.Fields("Postfach").value) <> "" Then
195                     AD.Ort = "" & HauptRS.Fields("Ort1").value
200                     AD.Plz = "" & HauptRS.Fields("Plz1").value
                    Else
205                     AD.Ort = "" & HauptRS.Fields("Ort").value
210                     AD.Plz = "" & HauptRS.Fields("Plz").value
                    End If

                Else
                
215                 AD.belegDatum = ""
220                 AD.Lkz = ""
225                 AD.name = ""
230                 AD.MCode = ""
235                 AD.KtoNr = ""
240                 AD.Ort = ""
245                 AD.Plz = ""

                End If

            Else
            
250             AD.belegDatum = ""
255             AD.Lkz = ""
260             AD.name = ""
265             AD.MCode = ""
270             AD.KtoNr = ""
275             AD.Ort = ""
280             AD.Plz = ""

            End If

285         AD.LLSrcPfad = ArbeitsplatzPfad
290         AD.LLSrcDateiName = "SP52800.LL"
295         AD.DateiName = DateiName
300         AD.Pfad = Pfad

            '<Modified by: IL at 09.10.2024, Ver.: 6.7.101 >
305         Select Case Prefix

                Case "SFR"
    
310                 AD.Art = E_DATATYPE.Sonderfaktura_Rechnung
    
315             Case "SFG"
    
320                 AD.Art = E_DATATYPE.Sonderfaktura_Gutschrift
    
325             Case "SFA"
    
330                 AD.Art = E_DATATYPE.Sonderfaktura_Angebot
    
335             Case "SFB"
    
340                 AD.Art = E_DATATYPE.Sonderfaktura_Auftragsbestetigung

            End Select

            '305         If Prefix = "SFR" Then
            '310             AD.Art = E_DATATYPE.Sonderfaktura_Rechnung
            '            Else
            '315             AD.Art = E_DATATYPE.Sonderfaktura_Gutschrift
            '            End If
            '</Modified by: IL at 09.10.2024, Ver.: 6.7.101 >

345         Call LLArchivierenPDF(LL1, AD, Prefix)                              'TODO rückgabe wert abfragen und bei Fehler Rollback durchführen GOBD
                
            '##################################################
    
            'SHFileOperation-Aufruf war erfolgreich.
350         If Erfolg <> 0 Then

355             MsgBox "Der aktuelle Beleg konnte nicht archiviert werden. Bitte überprüfen Sie das Betriebsystem.", vbExclamation

360             GblnExternesArchiv = False

            Else

365             If GbArchiv Then

                    'Archivierung ist lizensiert.
370                 If (Not HauptRS Is Nothing) And (Not FolgeRS Is Nothing) Then

375                     OPEN_gConn
    
380                     If HauptRS.RecordCount > 0 Then

385                         gConn.BeginTrans
390                         Transaktion = True

395                         If UCase(Prefix) = "SFR" Then
                                'Wenn die zuvor stornierte Rechnung nochmals gedruckt wird, muss der Satz der Stornierten Rechnung gelöscht werden.
400                             gConn.Execute "DELETE FROM [2800_Archiv_Rng] WHERE BelegID = " & HauptRS!BelegID
405                             sql = "INSERT INTO [2800_Archiv_Rng] "
                            Else
                                'Wenn die zuvor stornierte Gutschrift nochmals gedruckt wird, muss der Satz der Stornierten Gutschrift gelöscht werden.
410                             gConn.Execute "DELETE FROM [2800_Archiv_Gut] WHERE BelegID = " & HauptRS!BelegID
415                             sql = "INSERT INTO [2800_Archiv_Gut] "
                            End If

420                         sql = sql & " (BelegID,Datei,ErstVon,AendVon) VALUES ('" & HauptRS!BelegID & "','" & ExtractFileName(ZielDatei) & "','" & GsUser & "','" & GsUser & "')"
425                         gConn.Execute sql
430                         gConn.CommitTrans
435                         Transaktion = False
                        Else
440                         GblnExternesArchiv = False
445                         MsgBox "Der Beleg: ''" & ZielDatei & "'' konnte nicht archiviert werden, da die benötigte Datengrundlage fehlt (Haupt-Recordset). ", vbExclamation
                        End If

                    Else
450                     GblnExternesArchiv = False
455                     MsgBox "Der Beleg: ''" & ZielDatei & "'' konnte nicht archiviert werden, da die benötigte Datengrundlage fehlt (Haupt- und Folge-Recordset). ", vbExclamation
                    End If
                End If
            End If
        End If

        Exit Sub

Fehler:
460     Call FehlerErklärung("SP52800B", "ArchivierenPDF()")

465     If Transaktion Then gConn.RollbackTrans
470     GblnExternesArchiv = False
End Sub

Private Function ExtractFileName(Pfad As String) As String

        '***Beginn
        On Error GoTo Fehler

        '***Ende
        Dim i As Integer
  
100     For i = Len(Pfad) To 1 Step -1

105         If Mid(Pfad, i, 1) = "\" Then Exit For
110     Next i

115     ExtractFileName = right(Pfad, Len(Pfad) - i)
  
        '***Beginn
        Exit Function

Fehler:
120     Call FehlerErklärung("SP52800B", "ExtractFileName")
        '***Ende
End Function

Public Sub EndBetraege(Tabelle As String, _
                       BelegID As Long, _
                       SteuerPflichtig As Double, _
                       SteuerFrei As Double)

        On Error GoTo Fehler

        Dim sql As String

        Dim rs  As ADODB.Recordset
  
        '***** MW 30.08.05 Ver: 5.1.117
        'Die SQL-Abfrage muss durch die Schleife ersetzt werden, da die Abfrage das Ergebnis der Multiplikation nicht rundet.
        'In der Abfrage kann die Access-Funktion Round auch nicht benutzt werden, da sie nicht kaufmänisch sonder mathematisch rundet.

        '  sql = "SELECT Sum(IIf([Steuer]=1,([Menge]*[EPreis]/IIf([Einheit]='%',100,1))-([Menge]*[EPreis]/IIf([Einheit]='%',100,1)*[Rabatt]/100),0)) AS SteuerPfl, "
        '  sql = sql & "Sum(IIf([Steuer]=0,([Menge]*[EPreis]/IIf([Einheit]='%',100,1))-([Menge]*[EPreis]/IIf([Einheit]='%',100,1)*[Rabatt]/100),0)) AS SteuerFr "
        '  sql = sql & "FROM [" & Tabelle & "] WHERE BelegID = " & BelegID
        '  sql = sql & " HAVING SatzTyp='A' OR SatzTyp='P' OR SatzTyp='L'" 'MW 26.04.05
        '  Set rs = GDB.OpenRecordset(sql, dbOpenSnapshot)
        '
        '  If rs.RecordCount > 0 Then
        '    If Not IsNull(rs!SteuerPfl) Then SteuerPflichtig = Runden(rs!SteuerPfl, 2)
        '    If Not IsNull(rs!SteuerFr) Then SteuerFrei = Runden(rs!SteuerFr, 2)
        '  End If

100     sql = "SELECT Steuer, Menge, EPreis, Einheit, Rabatt "
105     sql = sql & " FROM [" & Tabelle & "] WHERE BelegID = " & BelegID
110     sql = sql & " AND (SatzTyp='A' OR SatzTyp='P' OR SatzTyp='L')" 'MW 26.04.05
  
115     OPEN_gConn
        
        'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
        
120     Set rs = New ADODB.Recordset
125     rs.Open sql, gConn, adOpenStatic, adLockReadOnly

130     Do Until rs.EOF

135         If Trim(rs!Einheit) = "%" Then

                '<Modified by: DFiebach at 03.03.2025, Ver.: 6.7.106 >
                'ORIG
                '140             If rs!Steuer = 1 Then
                '145                 SteuerPflichtig = SteuerPflichtig + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                '                Else
                '150                 SteuerFrei = SteuerFrei + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                '                End If

                'NEU -> rs!Menge / 100 führte zum falschen Ergebnis, -> rs!EPreis / 100 analog zum fpSread-Formel
140             If rs!Steuer = 1 Then
145                 SteuerPflichtig = SteuerPflichtig + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)
                Else
150                 SteuerFrei = SteuerFrei + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)
                End If

                '</Modified by: DFiebach at 03.03.2025, Ver.: 6.7.106 >

            Else

155             If rs!Steuer = 1 Then
160                 SteuerPflichtig = SteuerPflichtig + RundenMitVz((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                Else
165                 SteuerFrei = SteuerFrei + RundenMitVz((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                End If
                
            End If

170         rs.MoveNext
        Loop

        '***** MW 30.08.05 Ver: 5.1.117
  
175     rs.Close
  
        Exit Sub

Fehler:
180     Call FehlerErklärung("SP52800B", "EndBetraege()")
End Sub

Public Sub EndBetraegeGesamtIstBrutto(Tabelle As String, _
                                      BelegID As Long, _
                                      SteuerPflichtig As Double, _
                                      SteuerFrei As Double)

        On Error GoTo Fehler

        Dim sql As String

        Dim rs  As ADODB.Recordset
  
        '***** MW 30.08.05 Ver: 5.1.117
        'Die SQL-Abfrage muss durch die Schleife ersetzt werden, da die Abfrage das Ergebnis der Multiplikation nicht rundet.
        'In der Abfrage kann die Access-Funktion Round auch nicht benutzt werden, da sie nicht kaufmänisch sonder mathematisch rundet.

        '  sql = "SELECT Sum(IIf([Steuer]=1,([Menge]*[EPreis]/IIf([Einheit]='%',100,1))-([Menge]*[EPreis]/IIf([Einheit]='%',100,1)*[Rabatt]/100),0)) AS SteuerPfl, "
        '  sql = sql & "Sum(IIf([Steuer]=0,([Menge]*[EPreis]/IIf([Einheit]='%',100,1))-([Menge]*[EPreis]/IIf([Einheit]='%',100,1)*[Rabatt]/100),0)) AS SteuerFr "
        '  sql = sql & "FROM [" & Tabelle & "] WHERE BelegID = " & BelegID
        '  sql = sql & " HAVING SatzTyp='A' OR SatzTyp='P' OR SatzTyp='L'" 'MW 26.04.05
        '  Set rs = GDB.OpenRecordset(sql, dbOpenSnapshot)
        '
        '  If rs.RecordCount > 0 Then
        '    If Not IsNull(rs!SteuerPfl) Then SteuerPflichtig = Runden(rs!SteuerPfl, 2)
        '    If Not IsNull(rs!SteuerFr) Then SteuerFrei = Runden(rs!SteuerFr, 2)
        '  End If

100     sql = "SELECT Steuer,Menge,EPreis,Einheit,Rabatt "
105     sql = sql & " FROM [" & Tabelle & "] WHERE BelegID = " & BelegID
110     sql = sql & " AND (SatzTyp='A' OR SatzTyp='P' OR SatzTyp='L')" 'MW 26.04.05

115     OPEN_gConn
        
        'DF 19.02.2020 , Ver.: 6.5.114 : Runden -> RundenMiVz
        
120     Set rs = New ADODB.Recordset
125     rs.Open sql, gConn, adOpenStatic, adLockReadOnly

130     Do Until rs.EOF

135         If Trim(rs!Einheit) = "%" Then

                '<Modified by: DFiebach at 03.03.2025, Ver.: 6.7.106 >

                'ORIG
                '140             If rs!Steuer = 1 Then
                '
                '145                 SteuerPflichtig = SteuerPflichtig + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                '
                '                Else
                '
                '150                 SteuerFrei = SteuerFrei + RundenMitVz((rs!Menge / 100 * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)
                '
                '                End If

                'NEU -> rs!Menge / 100 führte zum falschen Ergebnis, -> rs!EPreis / 100 analog zum fpSread-Formel
140             If rs!Steuer = 1 Then

145                 SteuerPflichtig = SteuerPflichtig + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)

                Else
                
150                 SteuerFrei = SteuerFrei + RundenMitVz((rs!Menge * rs!EPreis / 100 - rs!Menge * rs!EPreis / 100 * rs!Rabatt / 100), 2)

                End If

                '</Modified by: DFiebach at 03.03.2025, Ver.: 6.7.106 >

            Else

155             If rs!Steuer = 1 Then

160                 SteuerPflichtig = SteuerPflichtig + RundenMitVz((rs!Menge * rs!EPreis - rs!Menge * rs!EPreis * rs!Rabatt / 100), 2)

                Else

165                 SteuerFrei = SteuerFrei + RundenMitVz(((rs!Menge * rs!EPreis) - (rs!Menge * rs!EPreis * rs!Rabatt / 100)), 2)

                End If
                
            End If

170         rs.MoveNext
        Loop

        '***** MW 30.08.05 Ver: 5.1.117
  
175     rs.Close
  
        Exit Sub

Fehler:
180     Call FehlerErklärung("SP52800B", "EndBetraege()")
End Sub

Public Sub SetSpreadDezimalStellen(Spread As fpSpread, _
                                   Col As Long, _
                                   Row As Long, _
                                   value As Variant)

        'MW 28.12.06 Ver.: 5.3.105 Nachkommastellen dynamisch gestallten (Abhängig von Nachkommastellen im Artikelstamm).
        Dim DynDezStellen As String
            
100     Spread.Col = Col
105     Spread.Row = Row
110     value = "" & value

115     If InStr(1, value, ",") > 0 Then
120         DynDezStellen = Len(value) - InStr(1, value, ",")

125         If DynDezStellen > 2 Then
130             Spread.TypeNumberDecPlaces = DynDezStellen
            Else
135             Spread.TypeNumberDecPlaces = 2
            End If

        Else
140         Spread.TypeNumberDecPlaces = 2
        End If

End Sub

Public Function GetMin2DezimalStellen(value As Variant) As Integer
        'MW 28.12.06 Ver.: 5.3.105 Nachkommastellen dynamisch gestallten (Abhängig von Nachkommastellen im Artikelstamm).
            
100     value = "" & value

105     If InStr(1, value, ",") > 0 Then
110         GetMin2DezimalStellen = Len(value) - InStr(1, value, ",")

115         If GetMin2DezimalStellen < 2 Then
120             GetMin2DezimalStellen = 2
            End If

        Else
125         GetMin2DezimalStellen = 2
        End If

End Function

Public Sub DefineZusatztext(rsH As ADODB.Recordset, LL1 As ListLabel.ListLabel)

        On Error GoTo Fehler

        Dim MCode As String

        Dim sql   As String

        Dim rs    As ADODB.Recordset
    
100     If rsH.RecordCount > 0 Then
105         If Trim(rsH!MCode) <> "" Then
110             MCode = rsH!MCode
            Else
115             MCode = "@System"
            End If

        Else
120         MCode = "@System"
        End If
  
125     sql = "SELECT [1100_Texte].Text AS Zusatztext"
130     sql = sql & " FROM [1200_GrundKonditionen] INNER JOIN"

135     If rsH!Art = 0 Then
            'Rechnung
140         sql = sql & " [1100_Texte] ON [1200_GrundKonditionen].RDokument = [1100_Texte].Titel"
145         sql = sql & " WHERE ([1200_GrundKonditionen].RDokument <> '')"
        Else
            'Gutschrift
150         sql = sql & " [1100_Texte] ON [1200_GrundKonditionen].GDokument = [1100_Texte].Titel"
155         sql = sql & " WHERE ([1200_GrundKonditionen].GDokument <> '')"
        End If

160     sql = sql & " AND ([1200_GrundKonditionen].MCode = '" & MCode & "')"
  
165     OPEN_gConn
  
170     Set rs = New ADODB.Recordset
175     rs.Open sql, gConn, adOpenStatic, adLockReadOnly
  
180     Call LLDefineVariablen(LL1, rs, "Kd_")
185     Call LLDefineFelder(LL1, rs, "Kd_")
        
190     If Not objERechnung Is Nothing And rs.RecordCount > 0 Then             'DF 02.09.2024 , Ver.: 6.7.101
        
195         objERechnung.ZusatzTextSFSTFaktura = rs.Fields("ZusatzText").value
            
        End If
        
        Exit Sub

Fehler:
200     Call FehlerErklärung("SP52800B", "DefineZusatztext()")
End Sub

'HW 24.07.2013 - Diese Funktion dient dazu, alle globalen Objekte ab zu schließen und räumen
'bzw. auf Nothing zu setzen! Damit der Server die Dateien nicht offen hällt!
Public Sub DisposeObjects(frm As Form)

        Dim i, WS, db, rs As Integer

        On Error Resume Next
    
100     Close

105     Unload frmMsg
    
110     For i = Forms.Count - 1 To 0 Step -1

115         If Forms(i).name <> frm.name Then Unload Forms(i)
120         If Err.number <> 0 Then Err.Clear
125     Next i
    
130     GDBdef.Close

135     If Err.number <> 0 Then Err.Clear
140     GDBprm.Close

145     If Err.number <> 0 Then Err.Clear
150     GDBlog.Close

155     If Err.number <> 0 Then Err.Clear

160     Sup.Close

165     If Err.number <> 0 Then Err.Clear
170     Sup2.Close

175     If Err.number <> 0 Then Err.Clear
180     GmsgRS.Close

185     If Err.number <> 0 Then Err.Clear
190     GprmRS.Close

195     If Err.number <> 0 Then Err.Clear
200     gdbZusatzText.Close

205     If Err.number <> 0 Then Err.Clear
    
210     Set GDBdef = Nothing
215     Set GDBprm = Nothing
220     Set GDBlog = Nothing

225     Set Sup = Nothing
230     Set Sup2 = Nothing
235     Set GmsgRS = Nothing
240     Set GprmRS = Nothing
245     Set gdbZusatzText = Nothing
    
250     If Err.number <> 0 Then Err.Clear
    
255     DoEvents

260     If Err.number <> 0 Then Err.Clear

265     For WS = Workspaces.Count - 1 To 0 Step -1

270         With Workspaces(WS)
275             Debug.Print "Open Workspace: " & .name

280             For db = .Databases.Count - 1 To 0 Step -1

285                 With .Databases(db)
290                     Debug.Print "Open Database : " & .name

295                     For rs = .Recordsets.Count - 1 To 0 Step -1
300                         Debug.Print "Open Recordset: " & .Recordsets(rs).name
305                         .Recordsets(rs).Close

310                         If Err.number <> 0 Then Err.Clear
                        Next ' RS

315                     .Close

320                     If Err.number <> 0 Then Err.Clear
                    End With

                Next ' DB

325             If WS <> 0 Then .Close
330             If Err.number <> 0 Then Err.Clear
                ' WS=0 ist #Default Workspace#, und die kann nicht,
                ' und muss auch nicht geschlossen werden
            End With

        Next 'WS
   
335     If Err.number <> 0 Then Err.Clear
End Sub

Public Function BelegAnOpUebergeben(lngBelegID As Long)
        '--------------------------------------------------------------------------------
        ' Project    :       SP52800
        ' Procedure  :       BelegAnOpUebergeben
        ' Description:       [type_description_here]
        ' Created by :       GW
        ' Date-Time  :       12.3.2021-16:17:11
        '
        ' Parameters :       lngBelegID (Long)
        '--------------------------------------------------------------------------------

        On Error GoTo Fehler

        Dim objOp   As clsOp
        
        Dim logItem As Variant

        Dim connSQL As ADODB.Connection

100     Set connSQL = New ADODB.Connection
105     connSQL.ConnectionString = GetConnectionString(GsHauptPfad, Spedifix, GsAnwenderNr)
110     connSQL.Open

115     Set objOp = New clsOp
120     Set objOp.conn = connSQL
125     objOp.Mandant = GsAnwenderNr
130     objOp.ProgPfad = GsHauptPfad
135     objOp.ProgNr = "283"
140     objOp.SetArbeitsDatum
145     objOp.UserErfasZchn = GsUser

150     objOp.BelegAnOpUebergeben (lngBelegID)
        
155     For Each logItem In objOp.Log

160         Call Logbuch(CStr(logItem))

        Next
        
165     connSQL.Close
        
        Exit Function

Fehler:

170     Call FehlerErklärung("SP52800B", "BelegAnOpUebergeben()")

End Function

