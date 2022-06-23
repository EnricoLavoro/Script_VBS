'NOCACHE

' SARA  18/01/2010 : scollegamento  movimenti di magazzino premendo sul bottone "Conferma Doc"
' pers. file STOPDES\xml_ute\StruttureOp.xml
' creazione nuovo file STOPDES\xml_ute\teste\base\Conferma_doc.xml
' viene utilizzato anche lo script su bcol MovimentiMagazzino.vbs

CONST strPathUp2C = "\\Disenia-up\up\bin\up2c\DIS\"
CONST strDir = "\\vmdisenia\UP\bin\script\IIstBaseStOp.log"
CONST my_debug = true

'------------------------------------------------------------------------------------------------------------------------------
'Autore:Claudia
'data:14/5/18
'descrizione: invio fatture elettroniche cit 197/18
'prerequisiti: modifica WReportRisultatiMaster.xml e AReportRisultatiMaster.xml con bottone ButtonInviaFE
'claudia 23/1/19 inserito messaggio di conferma 
'claudia 24/1/19 disabilitato bottone mentre fa invio
'davide 05/03/19 aggiunto caricamento mail visualizzatore ad ogni salvataggio del documento
'paolo 14/07/21 cancellazione dati OCOP alla cancellazione di un ORC (CIT/21/327)
'enrico 17/06/22 aggiunta caricamento indirizzo e-mail quando si specifica il corrispondente interno
'------------------------------------------------------------------------------------------------------------------------------


Function CaricaMailVisualizzatore( serialsoggetto )
	On Error Resume Next

	CaricaMailVisualizzatore= 0
	
	Dim connection: Set connection = CreateObject("ADODB.Connection")		
	connection.Open  globals.GetEnvironment("DSNAPP"), "grupposga","agsoppurg"
				
	Dim rs:	Set rs = CreateObject("ADODB.RecordSet")
	rs.CursorType = 0 'adOpenForwardOnly
	rs.LockType = 1 'adLockReadOnly
			
	Dim strSQL: strSQL = 	"SELECT  MailVisualizzatore" & _
				" FROM 	[ITV_XFattPA_MailVisualizzatore] WITH (NOLOCK) " & _
				" WHERE  AsogCSer = " & serialsoggetto
	'mytrace " DENTRO "& STRSQL
	Call rs.Open( strSQL, connection )
	'mytrace "dopo query"		
	if not rs.Eof then
		
		rs.MoveFirst()
				
		Dim email
		
		email = rs( "MailVisualizzatore" )
		
		Call this.changeValue("KTTISXDFattPAEmail",email)
		
	end if
			
	Set rs = Nothing
	
	Call connection.Close()
	Set connection = Nothing

	mytrace " fine "
	
End Function
Function OnCommand(Nome)
	On Error Resume Next
	OnCommand=0
	'msgbox "command" & nome
	if  Nome= "ButtonConfermaDoc" then
		dim utente
		utente = ucase(trim(Globals.GetEnvironment("UtenCUtente")))
		'msgbox("Utente: " & utente)
		' se si vuole fare in modo che solo alcuni utenti possano scollegare i movimenti magazzino
		'if  utente <> ucase(trim("administrator")) and utente <> ucase(trim("sara"))Then 
		'		msgbox ("Utente NON autorizzato alla conferma del documento")
		'	exit Function
		'end if
	
		'MsgBox(NomeInterattore & ".OnCommand(" & Nome &  ")")
		call ConfermaDoc()
	else if ucase(Nome) = ucase("ButtonInvioFE") then
		DIM strCommand
		dim objShell
		Set objShell = CreateObject("WScript.Shell")
		
		
		'MsgBox(NomeInterattore & ".OnCommand(" & Nome &  ")  " & this.GetValue("MbaisCSer"))
		if  NOT this.getValue("KTTISFattPAInviata") then
			if vbYes = msgbox("Vuoi inviare la fattura elettronica al sistema di interscambio (SDI)?", vbYesNO,"INVIARE FE?") then
				'disabilita bottone
				dim of: of = false
				dim ov: ov = true
		
				call this.SetControlProperty("ButtonInvioFE","enable", of)	'disabilita bottone

strCommand = strPathUp2C & "Up2c.exe i "& this.GetValue("MbaisCser") & " """ & this.GetValue("KTTISFattPAPath2") &""" """ & this.GetValue("KTTISXDFattPAEmail") & """"
		
		call objShell.Run(strCommand, 0, True)
		msgbox "Invio effettuato, controllare gli esiti."
				 call this.SetControlProperty("ButtonInvioFE","enable", ov)	'abilita bottone
			end if
		else
			msgbox "Non è possibile inviare la fattura perchè è già stata inviata"
		end if
	END IF	
	end if
End Function

Function ConfermaDoc()

            On Error Resume Next

            OnConfermaDoc=0
           
 	' msgbox("Conferma  0")
            dim IRIstBaseStOp

            Set IRIstBaseStOp =this.GetApplication().GetInteractor("IRIstBaseStOp")

   
            dim row

            row = 0 

            Dim collezioneRighe

            Set collezioneRighe = IRIstBaseStOp.GetCollection()

            If ( IsObject( collezioneRighe ) ) Then

			 	'msgbox("Conferma  1_1")

				Dim res: res = collezioneRighe.GetFirst()
				'msgbox("Conferma  121")
                                    Dim RigaCorr

                                    While( res )

				      '  msgbox("Conferma  2")

				        call IRIstBaseStOp.SetCurrentRow( row )
                                                Set RigaCorr=collezioneRighe.GetCurrent()

                                                'serial movmag
					dim sermovmag :	sermovmag = RigaCorr.GetValue("HMGISCRMvmg" )                                          
					'dim serbais :	serbais = RigaCorr.GetValue("HMGISCRMvmg" )                                          
					'msgbox ("SerMag " & cstr(sermovmag))
					Call RigaCorr.SetValue("RbaisBrmmov",0)
 
                                                Call RigaCorr.SetValue("HMGISCRMvmg", clng(0)  )                                          
                                                call IRIstBaseStOp.ChangeValue( "HMGISCRMvmg", clng(0) )
                                               
					
                                                'codice bello movmag
                                                Call RigaCorr.SetValue("NumMovMag", "" )        
                                                call IRIstBaseStOp.ChangeValue( "NumMovMag","" )

                                                'codice brutto movmag
                                                Call RigaCorr.SetValue("CodHMGISCRMvmg","" )
                                                call IRIstBaseStOp.ChangeValue( "CodHMGISCRMvmg","" )
					
					'Apro il document dei movimenti di magazzino 145		
					dim ColSeriga : Set ColSerRiga = globals.MakeBCol(145)
					'Passo il serial dei movimenti di magazzino
					dim aParam : aParam = Array(sermovmag)
					'Apro la collezione personalizzata 696969 
					Call ColSerRiga.open(696969, aParam)
					'Leggo documento seriale di riga e mando in esecuzione l'update
					Dim res2: res2 = ColSerRiga.GetFirst()   

					Set RigaCorr = nothing
                                                row = row + 1
                                                res=collezioneRighe.GetNext()

                                    Wend

                        End If

                        Set collezioneRighe = nothing


End Function


'-----------------------------------------------------------------------------------
'Autore: claudia
'Data: 27/11/15
'Descrizione: gestire l'evento di modifica della data consegna promessa:
'           - in un campo memorizzare utente e data modifica
'           -inviare una mail al resp produzione e al resp consegne
'  quando si carica un documento con data consegna mostrare un messaggio "Attenzione consegna promessa. <data>" 
'-----------------------------------------------------------------------------------

Function OnChange(Nome,Valore)
	'On Error Resume Next
	OnChange=0
	'MsgBox(NomeInterattore & ".OnChange(" & Nome & "," & Valore & ")")
	
	If (ucase(Nome) = ucase("MbaisTConf")) then
		
		this.ChangeValue "MbaisDDescr", ucase(trim(Globals.GetEnvironment("UtenCUtente"))) & " il " & StrGetDate() 
		
		
	end if

	'claudia 5/11/18 gestione contestazione
	If (ucase(Nome) = ucase("MbaisXBContestazione") )then
		if (Valore = 1 )then
		
			this.ChangeValue "MbaisXDUtenteContestazione", ucase(trim(Globals.GetEnvironment("UtenCUtente"))) & " il " & StrGetDate() 
		
		end if
		
	end if
	
	If (ucase(Nome) = ucase("KanisCRAsog") )then
		'	MSGBOX (this.getValue("KanisCRAsog"))
		'	MSGBOX (this.getValue("MbaisCRAsog"))
		'	msgbox valore

		call CaricaMailVisualizzatore(valore)
			
	end if
	
	'enrico 17/06/22 aggiunta caricamento indirizzo e-mail quando si specifica il corrispondente interno
	If (ucase(Nome) = ucase("codLkKgsisCRvgrs_CorrispInt")) then
		call CompilaEmail(this.getValue("codLkKgsisCRvgrs_CorrispInt"))
	end if
End Function

'enrico 17/06/22 aggiunta caricamento indirizzo e-mail quando si specifica il corrispondente interno
Function CompilaEmail(ValoreGruppoStatistico)
	'MsgBox Globals.GetEnvironment("RgrsXIndEmail")
	Dim conn : Set conn = CreateObject("ADODB.Connection")
	conn.Open globals.GetEnvironment("DsnAPP"),"grupposga","agsoppurg"
	
	'Creazione oggetto comando
	Dim rs:	Set rs = CreateObject("ADODB.RecordSet")
	rs.CursorType = 0 'adOpenForwardOnly
	rs.LockType = 1 'adLockReadOnly
	
	Dim strSQL : strSQL = "SELECT RgrsXIndEmail FROM GruppiStatisticiRighe WITH (NOLOCK) WHERE RgrsCRgrs = '" & ValoreGruppoStatistico & "'"
	
	Call rs.Open(strSQL, conn)	
	if not rs.Eof then
		rs.MoveFirst()
		Dim email
		email = rs("RgrsXIndEmail")
		Call this.changeValue("MbaisXEmailCorrispInt",email)
	end if
			
	Set rs = Nothing
	Call conn.Close()
	Set conn = Nothing
End Function

Function OnQueryDrop()
	
	'CIT/21/327: cancellazione dati ordini di produzione alla cancellazione di un ordine clienti
	OnQueryDrop = 0
	
	Dim clStop, Consuntivo
	clStop = this.GetCollection().GetCurrent().GetValue("MbaisCRMcso") 
	
	'msgbox this.GetCollection().GetCurrent().GetValue("MbaisCRMcso")
 	
	Consuntivo = EsisteConsuntivo (this.GetValue("MbaisCMbais"))
	
	If clStop = 1  Then 'ORC seriale
		If not(Consuntivo) then
			'msgbox this.GetValue("MbaisCMbais")
		
			Dim conn: Set conn = CreateObject("ADODB.Connection")  
			Dim DSNAPP
			DSNAPP = globals.GetEnvironment("DSNAPP")
			DSNAPP = Replace (DSNAPP, "SQLNCLI11" , "sqloledb")
			conn.Open  DSNAPP, "grupposga","agsoppurg"
		
			dim cmd,p

			Set cmd = CreateObject("ADODB.Command")
			set p = CreateObject("ADODB.Parameter")
		  
			cmd.ActiveConnection = conn
			cmd.CommandType = 4
			cmd.CommandTimeout = 1200 ' 20 minuti
				
			cmd.CommandText = "ITsp_STOP_EliminaOCOP_ORC"
			
			cmd.Parameters(1).Value = this.GetValue("MbaisCMbais")
			
			cmd.Execute
			
			MyTrace "OnQueryDrop: ITsp_STOP_EliminaOCOP_ORC " & this.GetValue("MbaisCMbais")
			
			MsgBox "Sono stati eliminati eventuali ordini di produzione collegati all'Ordine Cliente.", vbInformation, "Ordine Clienti"
		else
			MsgBox "Ordine di produzione non eliminabile (presente consuntivo).", vbExclamation, "Ordine Clienti"
		end if
	End If

	'*******************************************************************************************
End Function

Function EsisteConsuntivo (parNOrdine)
			Dim connection: Set connection = CreateObject("ADODB.Connection")		
			connection.Open  globals.GetEnvironment("DSNAPP"), "grupposga","agsoppurg"
			
			Dim rs:	Set rs = CreateObject("ADODB.RecordSet")
			rs.CursorType = 0 'adOpenForwardOnly
			rs.LockType = 1 'adLockReadOnly
			Dim strSQL
			
				strSQL =  "SELECT OcbaCser " & _
							" FROM 	Consuntivi " & _		
							" WHERE  left(OcbaCDocRif,6) = '" & parNOrdine & "'"
			
			Call rs.Open( strSQL, connection )
			
			if not rs.Eof then
				EsisteConsuntivo = true
			else
				EsisteConsuntivo = false
			end if
				
			Set rs = Nothing
			
			Call connection.Close()
			Set connection = Nothing
End Function

Function StrGetDate() 
	dim mydata
	mydata = Now()
	'StrGetDate =   Day(mydata)  & "/" & Month(mydata)  & "/" & right( Year(mydata),2)
	StrGetDate =   right("00"+cstr(Day(mydata)),2)  & "/" & right("00"+cstr(Month(mydata)),2)  & "/" & right("00"+cstr(Year(mydata)),2)
End Function

'-----------------------------------------------------------------------------------------------------
Function MyTrace(sMess)
' Funzione: Scrive file su server x file di logging
dim fso, oLog

  	if my_debug Then
		Set fso  = CreateObject("Scripting.FileSystemObject")	
		Set oLog = fso.OpenTextFile(strDir, 8, True)
		oLog.WriteLine Now()& " v. "& vers & ": " & sMess
		oLog.close
		set oLog = nothing
		set fso = Nothing
 	end If
  
 	MyTrace = 0

end Function

