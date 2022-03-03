'NOCACHE

Function OnCommand(Nome)
	On Error Resume Next
	OnCommand=0
	'Se il Campo Interattore selezionato corrisponde a quello richiesto
	if UCASE(Nome) = UCASE("btnCopia") then
		'Se il seriale ha un valore valido
		if this.GetValue("TacaCSer") > 0 then
			'Creazione connessione con Database UP
			Dim conn
			Set conn = CreateObject("ADODB.Connection")
			conn.Open globals.GetEnvironment("DsnAPP"),"grupposga","agsoppurg"
			
			'Creazione oggetto comando
			Dim nIdSdi
			Dim oCmd
			Set oCmd = CreateObject("ADODB.Command")

			'Creazione comando con Stored Procedure
			Set oCmd.ActiveConnection = conn
			oCmd.CommandType = 4				'Stored Procedures
			oCmd.CommandText = "ITsp_InsCopiaCommessaAttributi"

			'Specifica parametro di input
			oCmd.Parameters.Refresh
			oCmd.Parameters(1).Value = this.GetValue("TacaCSer")
			
			'Esecuzione del comando
			oCmd.Execute()

			'Chiusura connessione e comando
			Set oCmd = Nothing
			conn.Close
			Set conn = Nothing

			MsgBox("Commessa copiata con successo!")
		else
			MsgBox("Commessa non trovata")
		end if	
	end if
End Function

'----------------------------------------------------------
'Lista Campi Interattore ( nome [ label ] { tipo } )
'----------------------------------------------------------

' TacaTCrea [ Data creazione ]  { CDateTimeControl }
' TacaCTaca [ Codice ]  { CEditControl }
' TacaTinizioRich [ Inizio richiesta ]  { CDateTimeControl }
' TacaTinizioEff [ Effettiva ]  { CDateTimeControl }
' deslookup_su_TacaCRAsog [  ]  { CEditControl }
' ButtonMultilingua [  ]  { CButtonExecuteControl }
' lookup_su_TacaCRTcom [ Tipo ]  { CLookupControl }
' TacaTfinePrev [ Prevista ]  { CDateTimeControl }
' TacaTinizioPrev [ Prevista ]  { CDateTimeControl }
' lbllookup_su_TacaCRStato [  ]  { CLinkControl }
' btnCopia [  ]  { CButtonExecuteControl }
' TacaFAttivoSer [ Attiva ]  { CComboBoxControl }
' TacaDTaca [  ]  { CEditControl }
' lookup_su_TacaCRAsog [ Cliente ]  { CLookupControl }
' TacaTfineRich [ Fine richiesta ]  { CDateTimeControl }
' StrTacaDTaca [  ]  { CEditControl }
' deslookup_su_TacaCRTcom [  ]  { CEditControl }
' lbllookup_su_TacaCRAsog [  ]  { CLinkControl }
' deslookup_su_TacaCRTaca_ca [  ]  { CEditControl }
' TacaAbilRicalcAna [ Visualizza in Ric. Analitica ]  { CCheckBoxControl }
' lbllookup_su_TacaCRTcom [  ]  { CLinkControl }
' TacaCodRif [ Cod.esterno ]  { CEditControl }
' deslookup_su_TacaCRStato [  ]  { CEditControl }
' lookup_su_TacaCRTaca_ca [ Comm. analitica ]  { CLookupControl }
' lookup_su_TacaCRStato [ Stato ]  { CLookupControl }
' TacaTfineEff [ Effettiva ]  { CDateTimeControl }
' StrTacaCTaca [ Codice ]  { CEditControl }
' codlookup_su_TacaCRTcom [ Tipo ]  { CEditControl }
' buttlookup_su_TacaCRTcom [  ]  { CButtonExecuteControl }
' codlookup_su_TacaCRAsog [ Cliente ]  { CEditControl }
' buttlookup_su_TacaCRAsog [  ]  { CButtonExecuteControl }
' codlookup_su_TacaCRTaca_ca [ Comm. analitica ]  { CEditControl }
' buttlookup_su_TacaCRTaca_ca [  ]  { CButtonExecuteControl }
' codlookup_su_TacaCRStato [ Stato ]  { CEditControl }
' buttlookup_su_TacaCRStato [  ]  { CButtonExecuteControl }
' TacaCSer [  ]  { Interger }
' TacaCRTcom [  ]  { Interger }
' TcomFiltroSogg [  ]  { Interger }
' TacaCRStato [  ]  { Interger }
' TacaFAttivoAzione [  ]  { String }
' TacaCRAsog [  ]  { Interger }
' CodiceDB1 [  ]  { String }
' TacaCRTaca_ca [  ]  { Interger }




'----------------------------------------------------------
NomeInterattore = "ITabAttComm"
'----------------------------------------------------------


'Function OnQueryInitialize()
'	On Error Resume Next
'	OnQueryInitialize=0
'	MsgBox(NomeInterattore & ".OnQueryInitialize")

'End Function


'Function OnInitialize()
'	On Error Resume Next
'	OnInitialize=0
'	MsgBox(NomeInterattore & ".OnInitialize")

'End Function


'Function OnLostFocus(Nome)
'	On Error Resume Next
'	OnLostFocus=0
''	Non usare MsgBox in quanto toglie il focus al campo di arrivo
''	MsgBox(NomeInterattore & ".OnLostFocus(" & Nome & ")")

'End Function


'Function OnQueryChange(Nome,Valore)
'	On Error Resume Next
'	OnQueryChange=0
'	MsgBox(NomeInterattore & ".OnQueryChange(" & Nome & "," & Valore & ")")

'End Function


'Function OnChange(Nome,Valore)
'	On Error Resume Next
'	OnChange=0
'	MsgBox(NomeInterattore & ".OnChange(" & Nome & "," & Valore & ")")

'End Function


'Function OnQueryDrop()
'	On Error Resume Next
'	OnQueryDrop=0
'	MsgBox(NomeInterattore & ".OnQueryDrop")

'End Function


'Function OnDrop()
'	On Error Resume Next
'	OnDrop=0
'	MsgBox(NomeInterattore & ".OnDrop")

'End Function


'Function OnQueryNew()
'	On Error Resume Next
'	OnQueryNew=0
'	MsgBox(NomeInterattore & ".OnQueryNew")

'End Function


'Function OnNew()
'	On Error Resume Next
'	OnNew=0

'	MsgBox(NomeInterattore & ".OnNew")

'End Function


'Function OnAfterNew()
'	On Error Resume Next
'	OnAfterNew=0

'	MsgBox(NomeInterattore & ".OnAfterNew")

'End Function


'Function OnQueryMove()
'	On Error Resume Next
'	OnQueryMove=0

'	MsgBox(NomeInterattore & ".OnQueryMove")

'End Function


'Function OnMove()
'	On Error Resume Next
'	OnMove=0

'	MsgBox(NomeInterattore & ".OnMove")

'End Function


'Function OnQuerySave()
'	On Error Resume Next
'	OnQuerySave=0

'	MsgBox(NomeInterattoe & ".OnQuerySave")

'End Function


'Function OnSave()
'	On Error Resume Next
'	OnSave=0

'	MsgBox(NomeInterattore & ".OnSave")

'End Function




'Function OnAfterSave()
'	On Error Resume Next
'	OnAfterSave=0

'	MsgBox(NomeInterattore & ".OnAfterSave")

'End Function


'Function OnQueryShowWindow(Nome,vParam)
'	On Error Resume Next
'	OnQueryShowWindow=0

'	Dim descr(3)
'	Parametri = NomeInterattore & ".OnQueryShowWindow" & "(" & Nome & ",vParam)" & Chr(13) & Chr(10)
'	descr(0) = "PARAMS_IN"
'	descr(1) = "IN_PARAMS"
'	descr(2) = "OUT_PARAMS"
'	for j=LBound(vParam) to UBound(vParam)
'		for i=LBound(vParam(j)) to UBound(vParam(j))
'			if (i = LBound(vParam(j))) then
'				Parametri = Parametri & "  " & descr(j) & Chr(13) & Chr(10)
'			end if
'			Parametri = Parametri &  "    " & "vParam(" & CStr(j) & ")(" & CStr(i) & ")=" & CStr(vParam(j)(i)) & Chr(13) & Chr(10)
'		next
'	next

''ATTENZIONE:
''Se si modificano i valori di vParam (ad es: vParam(0)(0) = "..."
''e' necessario inserire dopo le modifiche 
''	OnQueryShowWindow = vParam

'	MsgBox(Parametri)
'End Function


'Function OnShowWindow(Nome,vParam)
'	On Error Resume Next
'	OnShowWindow=0

'	Dim descr(3)
'	Parametri = NomeInterattore & ".OnShowWindow" & "(" & Nome & ",vParam)" & Chr(13) & Chr(10)
'	descr(0) = "PARAMS_IN"
'	descr(1) = "IN_PARAMS"
'	descr(2) = "OUT_PARAMS"
'	for j=LBound(vParam) to UBound(vParam)
'		for i=LBound(vParam(j)) to UBound(vParam(j))
'			if (i = LBound(vParam(j))) then
'				Parametri = Parametri & "  " & descr(j) & Chr(13) & Chr(10)
'			end if
'			Parametri = Parametri &  "    " & "vParam(" & CStr(j) & ")(" & CStr(i) & ")=" & CStr(vParam(j)(i)) & Chr(13) & Chr(10)
'		next
'	next

''ATTENZIONE:
''Se si modificano i valori di vParam (ad es: vParam(0)(0) = "..."
''e' necessario inserire dopo le modifiche 
''	OnShowWindow = vParam

'	MsgBox(Parametri)
'End Function


'Function OnQueryUpdateRow(controlName,rowIndex)
'	On Error Resume Next
'	OnQueryUpdateRow=0
'	MsgBox(NomeInterattore & ".OnQueryUpdateRow(" & controlName & "," & rowIndex & ")")

'End Function


'Function OnUpdateRow(controlName,rowIndex)
'	On Error Resume Next
'	OnUpdateRow=0
'	MsgBox(NomeInterattore & ".OnUpdateRow(" & controlName & "," & rowIndex & ")")

'End Function


'Function OnQueryRowsDrop(Nome,vBefore,vParam)
'	On Error Resume Next
'	OnQueryRowsDrop=0
'    Parametri = NomeInterattore & ".OnQueryRowsDrop " & "(" & Nome & "," & vBefore & ",vParam)" & Chr(13) & Chr(10)
'	for i=LBound(vParam) to UBound(vParam)
'	    Parametri = Parametri & "  " & vParam(i) & Chr(13) & Chr(10)
'	next
'	MsgBox(Parametri)
'End Function

'Function OnRowsDrop(Nome,vBefore,vParam)
'	On Error Resume Next
'	OnRowsDrop=0
'    Parametri = NomeInterattore & ".OnRowsDrop " & "(" & Nome & "," & vBefore & ",vParam)" & Chr(13) & Chr(10)
'	for i=LBound(vParam) to UBound(vParam)
'	    Parametri = Parametri & "  " & vParam(i) & Chr(13) & Chr(10)
'	next
'	MsgBox(Parametri)
'End Function

'Function OnQueryUndoRowsDrop(Nome,vBefore,vParam)
'	On Error Resume Next
'	OnQueryUndoRowsDrop=0
'    Parametri = NomeInterattore & ".OnQueryUndoRowsDrop " & "(" & Nome & "," & vBefore & ",vParam)" & Chr(13) & Chr(10)
'	for i=LBound(vParam) to UBound(vParam)
'	    Parametri = Parametri & "  " & vParam(i) & Chr(13) & Chr(10)
'	next
'	MsgBox(Parametri)
'End Function

'Function OnUndoRowsDrop(Nome,vBefore,vParam)
'	On Error Resume Next
'	OnUndoRowsDrop=0
'    Parametri = NomeInterattore & ".OnUndoRowsDrop " & "(" & Nome & "," & vBefore & ",vParam)" & Chr(13) & Chr(10)
'	for i=LBound(vParam) to UBound(vParam)
'	    Parametri = Parametri & "  " & vParam(i) & Chr(13) & Chr(10)
'	next
'	MsgBox(Parametri)
'End Function

'Function OnQueryCommand(Nome)
'	On Error Resume Next
'	OnQueryCommand=0

'	MsgBox(NomeInterattore & ".OnQueryCommand(" & Nome &  ")")

'End Function


