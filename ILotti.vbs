'NOCACHE
'---------------------------------------------------------------------------------------------------
'ButtonMostraSpecifiche: verifica il lotto a quali clienti in base alla tabella specifiche tecniche XSpecTecnR (nuova maschera Capitolati) può andar bene
'l'articolo di confronto è quello del mov mag del lotto se presenti altrimenti è quello inserito
'nella maschera lotti e passato come parametr in questa funzione
'ButtonCopia: copia un lotto
'------------------------------------------------------------------------------
Function OnCommand(Nome)
	On Error Resume Next
	OnCommand=0
	'MsgBox(NomeInterattore & ".OnCommand(" & Nome &  ")")
	if Nome = "ButtonMostraSpecifiche" then
		Set Ilotti= this.GetApplication().GetInteractor("ILotti")
		dim SerLot : SerLot =Ilotti.GetCollection().GetCurrent().getValue("LottCser")
		dim SerArt : SerArt =Ilotti.GetCollection().GetCurrent().getValue("LottCRArtb")
		'msgbox ("SerLot " & cstr(SerLot) & " SerArt " & cstr(serart))
		'msgbox ("SerLot " &	 cstr(SerLot))
		call ShowRighe (SerLot,SerArt)
	end if
	
	if Nome = "ButtonCopia" then
		Set Ilotti= this.GetApplication().GetInteractor("ILotti")
		dim Serlotto : Serlotto =Ilotti.GetCollection().GetCurrent().getValue("LottCser")
		CodNewLotto = CSTR(inputbox("Nuovo Codice Lotto: "))
		if CodNewLotto <> "" then
			err.Clear
			 set bdoc = globals.MakeBDocSP()
			if  err <> 0 then
				msgbox (err.Description)
				err.Clear
			end if      
			parametriStore =  "@Serlotto='" & cstr(Serlotto) &  "' , @CodNewLotto='" & CodNewLotto & "'" 
			aParam1 = Array(parametriStore)
			err.Clear
			call bdoc.New()
			if  err <> 0 then
				msgbox (err.Description)
				err.Clear
			end if      
			err.Clear
			dim result
			result = bdoc.DoStoreProc("NetST_DuplicaLotto",aParam1)
			if  err <> 0 then
				msgbox (err.Description)
				err.Clear
				else
				msgbox "Copia Lotto Completata !!!!! Esci e ricarica il Lotto "
			 end if      
		end if
	end if
End Function

Function ShowRighe(SerLotto,SerArt)
'verifica il lotto a quali clienti in base alla tabella specifiche tecniche XSpecTecnR (nuova maschera Capitolati) può andar bene
'l'articolo di confronto è quello del mov mag del lotto se presenti altrimenti è quello inserito
'nella maschera lotti e passato come parametr in questa funzione

'claudia 7/6/21 salva le compatibilità in lottXNote1

	On Error Resume Next
	ShowRighe=0
	
	dim outCompatibile : outCompatibile =""
	dim outNonCompatibile : outNonCompatibile =""
	'msgbox ("SerLot_1 " & cstr(SerLotto))
	Dim connection: Set connection = CreateObject("ADODB.Connection")	
	connection.Open  globals.GetEnvironment("DSNAPP"), "grupposga","agsoppurg"
						
	Dim rs:Set rs = CreateObject("ADODB.RecordSet")
	rs.CursorType = 0 'adOpenForwardOnly
	rs.LockType = 1 'adLockReadOnly
					
		Dim strSQL: strSQL = "SELECT DISTINCT MvmgCRArtb ,ArtbCartb" & _
				" FROM MovimentiMagazzino " & _
				" inner Join Lotti on MvmgCRLott = LottCser " & _
				" inner Join ArticoliBase on MvmgCRArtb = ArtbCser " & _
				" where LottCser= " & SerLotto & _
				" group by MvmgCRArtb ,ArtbCartb"
				
	'msgbox("Strsql= " & Strsql)

	Call rs.Open( strSQL, connection )
	dim ff : ff= 0				
	if not rs.Eof then
		'caso A) ci sono mov mag associatoi al lotto-->ctrl tutti gli articoli movimentati
		rs.MoveFirst()
		While( not rs.Eof )
			'msgbox ("Ciclo 1")
			'SerArt =  rs("MvmgCRArtb")
			dim CodArt : CodArt = left (rs("ArtbCArtb"),3)

			'msgbox("SerArt : " & cstr(SerArt))
			Dim rs_s:Set rs_s = CreateObject("ADODB.RecordSet")
			Dim connection_s: Set connection_s = CreateObject("ADODB.Connection")	
			connection_s.Open  globals.GetEnvironment("DSNAPP"), "grupposga","agsoppurg"
			rs_s.CursorType = 0 'adOpenForwardOnly
			rs_s.LockType = 1 'adLockReadOnly
							
			Dim strSQL_s: strSQL_s = "SELECT AsogDAsog,XSpecTecnR.* " & _
							" FROM XSpecTecnR " & _
							" Inner Join Soggetti on AsogCSer = XSpTRCRAsog " & _
							" Inner Join ArticoliBase on ArtbCSer = XSpTRXArtb " & _
							" where ArtbCartb like '" & CodArt & "%'"
			'msgbox("Strsql_s= " & Strsql_s)
			
			Call rs_s.Open( strSQL_s, connection_s )

			if not rs_s.Eof then
				rs_s.MoveFirst()
				While( not rs_s.Eof )
					'msgbox ("Ciclo 2")
					call MostraRighe (rs_s, outCompatibile, outNonCompatibile )
					
					rs_s.MoveNext()	
					ff = 1
				Wend
			end if
			Set rs_s = Nothing
			Call connection_s.Close()
			Set connection_s = Nothing
			rs.MoveNext()	
			
		Wend
		if ff = 0 then
			Msgbox ("!!!!!!!!Nessuna Specifica Presente!!!!!!!!")
		end if
	else
		'caso B) non ci sono mov mag considero l'articolo passato come parametro
			'msgbox("SerArt : " & cstr(SerArt))
			'inizio leggere
			Dim rs_c:Set rs_c = CreateObject("ADODB.RecordSet")
			Dim connection_c: Set connection_c = CreateObject("ADODB.Connection")	
			connection_c.Open  globals.GetEnvironment("DSNAPP"), "grupposga","agsoppurg"
			rs_c.CursorType = 0 'adOpenForwardOnly
			rs_c.LockType = 1 'adLockReadOnly
							
			Dim strSQL_c: strSQL_c = "SELECT ArtbCArtb " & _
							"FROM ArticoliBase " & _
							"where ArtbCSer = " & SerArt 
			'msgbox("Strsql_c= " & Strsql_c)
			
			Call rs_c.Open( strSQL_c, connection_c )
			if not rs_c.Eof then
				rs_c.MoveFirst()
				CodArt = left (rs_c("ArtbCArtb"),3)
			end if
			Set rs_c = Nothing
			Call connection_c.Close()
			Set connection_c = Nothing
			'fine leggere
			
			Set rs_s = CreateObject("ADODB.RecordSet")
			Set connection_s = CreateObject("ADODB.Connection")	
			connection_s.Open  globals.GetEnvironment("DSNAPP"), "grupposga","agsoppurg"
			rs_s.CursorType = 0 'adOpenForwardOnly
			rs_s.LockType = 1 'adLockReadOnly
							
			strSQL_s = "SELECT AsogDAsog,XSpecTecnR.* " & _
						" FROM XSpecTecnR " & _
						" Inner Join Soggetti on AsogCSer = XSpTRCRAsog " & _
						" Inner Join ArticoliBase on ArtbCSer = XSpTRXArtb " & _
						" where ArtbCartb like '" & CodArt & "%'"

			'msgbox("Strsql_s= " & Strsql_s)
			Call rs_s.Open( strSQL_s, connection_s )
			
			if not rs_s.Eof then
				rs_s.MoveFirst()
				While( not rs_s.Eof )
					'msgbox ("Ciclo 2")
					call MostraRighe (rs_s, outCompatibile, outNonCompatibile)
					
					rs_s.MoveNext()	
				Wend
			else
				Msgbox ("!!!!!!!!Nessuna Specifica Presente!!!!!!!!")
			end if
			Set rs_s = Nothing
			Call connection_s.Close()
			Set connection_s = Nothing
	end if
	Set rs = Nothing
	Call connection.Close()
	Set connection = Nothing

	'claudia 7/6/21--
	if outNonCompatibile <>"" then
	    msgbox  outNonCompatibile,,"NON COMPATIBILI"
	end if
	
	if outCompatibile <> "" then
		if vbYes = msgbox ("OK: "& outCompatibile  & " Confermi la copia in note?", vbYesNo, "Copio compatibili?") then
			call this.ChangeValue("LottXNote1", "OK: "& outCompatibile)
		end if

	end if
	'--fine 7/6/21
End Function

Function MostraRighe(rs_s,byref outCompatibile, byref outNonCompatibile )
'DESCRIZIONE: controlla il singolo capitolato se rispettato dai valori del lotto
'claudia 6/12/21 se il capitolato  ha il range 0-99999 non viene effettuato il controllo
'                perchè vuol dire che il capitolato non ha un vincolo su quel valore
	On Error Resume Next
	MostraRighe=0
	dim Ilotti
	Set Ilotti= this.GetApplication().GetInteractor("ILotti")
	
	'Controllo Elementi Chimici

	dim ElCh
	dim LottXc 		: LottXc = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXC"))
	dim LottXMn 	: LottXMn = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXMn"))
	dim LottXSi 	: LottXSi = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXSi"))
	dim LottXp 		: LottXP = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXP"))
	Dim LottXS		: LottXS = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXS"))
	Dim LottXCr		: LottXCr = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXCr"))
	Dim LottXNi		: LottXNi = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXNi"))
	Dim LottXMo		: LottXMo = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXMo"))	
	Dim LottXCu		: LottXCu = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXCu"))
	Dim LottXSn 	: LottXSn = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXSn"))
	Dim LottXAs 	: LottXAs = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXAs"))
	Dim LottXPb 	: LottXPb = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXPb"))
	Dim LottXAl 	: LottXAl = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXAl"))
	Dim LottXTi 	: LottXTi = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXTi"))
	Dim LottXNb 	: LottXNb = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXNb"))
	Dim LottXB 		: LottXB = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXB"))
	Dim LottXCe 	: LottXCe = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXCe"))
	Dim LottXCa 	: LottXCa = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXCa"))
	Dim LottXFb 	: LottXFb = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXB"))
	
	'MSGBOX "1"
	if (LottXc >= 0) and (rs_s("XSpTRdaC") > 0 or rs_s("XSpTRaC") < 999999) then
		if LottXc < rs_s("XSpTRdaC")   then
			ElCh = "C: " & cstr(LottXc) & " < "   &  cstr(rs_s("XSpTRdaC")) & " "
		end if	
		if LottXc > rs_s("XSpTRaC")  then
			ElCh = "C: " & cstr(LottXc) & " > "   &  cstr(rs_s("XSpTRaC")) & " "
		end if	
	end if
	
	if (LottXMn >= 0) and (rs_s("XSpTRdaMn") > 0 or rs_s("XSpTRaMn") < 999999) then
		if LottXMn < rs_s("XSpTRdaMn") then
			ElCh =  ElCh + "Mn: " & cstr(LottXMn)   & " < " & cstr(rs_s("XSpTRdaMn"))  & " "
		end if
		if LottXMn > rs_s("XSpTRaMn") then
			ElCh =  ElCh + "Mn: " & cstr(LottXMn)  & " > " & cstr(rs_s("XSpTRaMn"))  & " "
		end if
	end if
	
	if (LottXSi >= 0) and (rs_s("XSpTRdaSi") > 0 or rs_s("XSpTRaSi") < 999999) then
		if LottXSi < rs_s("XSpTRdaSi") then
			ElCh =  ElCh + "Si: " & cstr(LottXSi)  & " < " & cstr(rs_s("XSpTRdaSi")) & " "
		end if
		if LottXSi > rs_s("XSpTRaSi") then
			ElCh =  ElCh + "Si: " & cstr(LottXSi)  & " > " & cstr(rs_s("XSpTRaSi")) & " "
		end if
	end if
	
	if (LottXp >= 0) and (rs_s("XSpTRdaP") > 0 or rs_s("XSpTRaP") < 999999) then
		if LottXp <  rs_s("XSpTRdaP")  then
			ElCh =ElCh +  "P: " &  cstr(LottXp)  & "< " &  cstr(rs_s("XSpTRdaP"))  & " "
		end if
		if LottXp > rs_s("XSpTRaP")  then
			ElCh =ElCh +  "P: " &  cstr(LottXp)  & "> " &  cstr(rs_s("XSpTRaP"))  & " "
		end if
	end if
	
	if (LottXS >= 0) and (rs_s("XSpTRdaS") > 0 or rs_s("XSpTRaS") < 999999) then
		if LottXS < rs_s("XSpTRdaS")  then
			ElCh = ElCh + "S: " &  cstr(LottXS)  & "< " & cstr(rs_s("XSpTRdaS"))  & " "
		end if
		if LottXS > rs_s("XSpTRaS")  then
			ElCh = ElCh + "S: " &  cstr(LottXS)  & "> " & cstr(rs_s("XSpTRaS"))  & " "
		end if
	end if 
	
	if (LottXCr >= 0) and (rs_s("XSpTRdaCr") > 0 or rs_s("XSpTRaCr") < 999999) then
		if LottXCr < rs_s("XSpTRdaCr") then
			ElCh = ElCh + "Cr: " &  cstr(LottXCr) & " < " &  cstr(rs_s("XSpTRdaCr")) & " "
		end if
		if LottXCr > rs_s("XSpTRaCr") then
			ElCh = ElCh + "Cr: " & cstr(LottXCr) & " > " & cstr( this.getValue("LottXCr")) & " "
		end if
	end if
	
	if (LottXNi >= 0) and (rs_s("XSpTRdaMn") > 0 or rs_s("XSpTRaMn") < 999999) then
		if LottXNi < rs_s("XSpTRdaNi") then
			ElCh = ElCh + "Ni: " & cstr(LottXNi)  & " < " &cstr(rs_s("XSpTRdaNi")) & " "
		end if
		if LottXNi > rs_s("XSpTRaNi") then
			ElCh = ElCh + "Ni: " & cstr(LottXNi)  & " > " &cstr(rs_s("XSpTRaNi")) & " "
		end if
	end if
	
	if (LottXMo >= 0) and (rs_s("XSpTRdaMo") > 0 or rs_s("XSpTRaMo") < 999999) then
		if LottXMo < rs_s("XSpTRdaMo") then
			ElCh =ElCh +  "Mo: "  & cstr(LottXMo)  & " < " & cstr(rs_s("XSpTRdaMo"))  & " "
		end if
		if LottXMo > rs_s("XSpTRaMo") then
			ElCh =ElCh +  "Mo: "  & cstr(LottXMo)  & " > " & cstr(rs_s("XSpTRaMo"))  & " "
		end if
	end if
	
	if (LottXCu >= 0) and (rs_s("XSpTRdaCu") > 0 or rs_s("XSpTRaCu") < 999999) then
		if LottXCu < rs_s("XSpTRdaCu") then
			ElCh = ElCh + "Cu: " & cstr(LottXCu) & " < "  & cstr(rs_s("XSpTRdaCu")) &  " "
		end if
		if LottXCu > rs_s("XSpTRaCu") then
			ElCh = ElCh + "Cu: " & cstr(LottXCu) & " > "  & cstr(rs_s("XSpTRaCu")) &  " "
		end if	
	end if
	
	if (LottXSn >= 0) and (rs_s("XSpTRdaSn") > 0 or rs_s("XSpTRaSn") < 999999) then
		if LottXSn < rs_s("XSpTRdaSn") then
			ElCh = ElCh +  "Sn: "  & cstr(LottXSn)  & " < " & cstr(rs_s("XSpTRdaSn")) & " "
		end if
		if LottXSn > rs_s("XSpTRaSn") then
			ElCh = ElCh +  "Sn: "  & cstr(LottXSn)  & " > " & cstr(rs_s("XSpTRaSn")) & " "
		end if
	end if
	
	if (LottXAs >= 0) and (rs_s("XSpTRdaAs") > 0 or rs_s("XSpTRaAs") < 999999) then
		if LottXAs < rs_s("XSpTRdaAs") then
			ElCh = ElCh + "As: " & cstr(LottXAs)  & " < " &  cstr(rs_s("XSpTRdaAs")) & " "
		end if
		if LottXAs > rs_s("XSpTRaAs") then
			ElCh = ElCh + "As: " & cstr(LottXAs)  & " > " &  cstr(rs_s("XSpTRaAs")) & " "
		end if
	end if
	
	if (LottXPb >= 0) and (rs_s("XSpTRdaPb") > 0 or rs_s("XSpTRaPb") < 999999) then
		if LottXPb < rs_s("XSpTRdaPb") then
			ElCh = ElCh + "Pb: " & cstr(LottXPb)  & " < " & cstr(rs_s("XSpTRdaPb"))   & " "
		end if
		if LottXPb > rs_s("XSpTRaPb") then
			ElCh = ElCh + "Pb: " & cstr(LottXPb)  & " > " & cstr(rs_s("XSpTRaPb"))   & " "
		end if
	end if
	
	if (LottXAl >= 0) and (rs_s("XSpTRdaAl") > 0 or rs_s("XSpTRaAl") < 999999) then
		if LottXAl < rs_s("XSpTRdaAl") then
			ElCh = ElCh + "Al: " & cstr(LottXAl)  & " < " & cstr(rs_s("XSpTRdaAl"))  & " "
		end if
		if LottXAl > rs_s("XSpTRaAl") then
			ElCh = ElCh + "Al: " & cstr(LottXAl)  & " > " & cstr(rs_s("XSpTRaAl"))  & " "
		end if
	end if
	
	if (LottXTi >= 0) and (rs_s("XSpTRdaTi") > 0 or rs_s("XSpTRaTi") < 999999) then
	if LottXTi <  rs_s("XSpTRdaTi") then
		ElCh = ElCh + "Ti: " & cstr(LottXTi)  & " < " & cstr(rs_s("XSpTRdaTi")) & " "
	end if
	if LottXTi >  rs_s("XSpTRaTi") then
		ElCh = ElCh + "Ti: " & cstr(LottXTi)  & " > " & cstr(rs_s("XSpTRaTi")) & " "
	end if
	end if
	
	if (LottXNb >= 0) and (rs_s("XSpTRdaNb") > 0 or rs_s("XSpTRaNb") < 999999) then
	if LottXNb < rs_s("XSpTRdaNb")  then
		ElCh = ElCh + "Nb: " & cstr(LottXNb)  & " < " & cstr(rs_s("XSpTRdaNb")) & " "
	end if
	if LottXNb > rs_s("XSpTRaNb")  then
		ElCh = ElCh + "Nb: " & cstr(LottXNb)  & " > " & cstr(rs_s("XSpTRaNb")) & " "
	end if
	end if
	
	if (LottXB >= 0) and (rs_s("XSpTRdaB") > 0 or rs_s("XSpTRaB") < 999999) then
	if LottXB < rs_s("XSpTRdaB") then
		ElCh = ElCh + "B: " & cstr(LottXB)  & " < " & cstr(rs_s("XSpTRdaB")) & " "
	end if
	if LottXB > rs_s("XSpTRaB") then
		ElCh = ElCh + "B: " & cstr(LottXB)  & " > " & cstr(rs_s("XSpTRaB")) & " "
	end if
	end if
	
	if (LottXCe >= 0) and (rs_s("XSpTRdaCe") > 0 or rs_s("XSpTRaCe") < 999999) then
	if LottXCe  < rs_s("XSpTRdaCE") then
		ElCh = ElCh + "Ce: " & cstr(LottXCe)  & " > " & cstr(rs_s("XSpTRdaCE")) & " "
	end if
	if LottXCe  > rs_s("XSpTRaCE") then
		ElCh = ElCh + "Ce: " & cstr(LottXCe)  & " > " & cstr(rs_s("XSpTRaCE")) & " "
	end if
	end if
	
	if (LottXFb >= 0) and (rs_s("XSpTRdaCa") > 0 or rs_s("XSpTRaCa") < 999999) then
		if LottXCa < rs_s("XSpTRdaCa")  then
			ElCh =ElCh + "Ca: " & cstr(LottXCa)  & " < " & cstr(rs_s("XSpTRdaCa"))  & " "
		end if
		if LottXCa > rs_s("XSpTRaCa")  then
			ElCh =ElCh + "Ca: " & cstr(LottXCa)  & " > " & cstr(rs_s("XSpTRaCa"))  & " "
		end if
	end if
	
	if (LottXFb >= 0) and (rs_s("XSpTRdaFb") > 0 or rs_s("XSpTRaFb") < 999999) then
		if LottXFb < rs_s("XSpTRdaFb") then
			ElCh = ElCh + "Fb: " & cstr(LottXFb)  & " < " & cstr(rs_s("XSpTRdaFb"))  &  " "
		end if
		if LottXFb > rs_s("XSpTRaFb") then
			ElCh = ElCh + "Fb: " & cstr(LottXFb)  & " > " & cstr(rs_s("XSpTRaFb"))  &  " "
		end if
	end if
	
		
	'Controllo Temprabilità
	dim Tmpb
	Dim LottXHrc1 	: LottXHrc1 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc1"))	
	Dim LottXHrc1punto5 	: LottXHrc1punto5 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc1punto5"))	
	Dim LottXHrc2 	: LottXHrc2 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc2"))	
	Dim LottXHrc3 	: LottXHrc3 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc3"))	
	Dim LottXHrc5 	: LottXHrc5 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc5"))		
	Dim LottXHrc7 	: LottXHrc7 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc7"))	
	Dim LottXHrc9 	: LottXHrc9 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc9"))	
	Dim LottXHrc10 	: LottXHrc10 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc10"))	
	Dim LottXHrc11 	: LottXHrc11 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc11"))	
	Dim LottXHrc13 	: LottXHrc13 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc13"))
	Dim LottXHrc15 	: LottXHrc15 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc15"))	
	Dim LottXHrc20 	: LottXHrc20 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc20"))	
	Dim LottXHrc25 	: LottXHrc25 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc25"))	
	Dim LottXHrc30 	: LottXHrc30 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc30"))	
	Dim LottXHrc35 	: LottXHrc35 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc35"))
	Dim LottXHrc40 	: LottXHrc40 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc40"))
	Dim LottXHrc45 	: LottXHrc45 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc45"))
	Dim LottXHrc50 	: LottXHrc50 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHrc50"))
	

	
	if (LottXHrc1 >= 0) and (rs_s("XSpTRdaHrc1") > 0 or rs_s("XSpTRaHrc1") < 999999) then
		if LottXHrc1 < rs_s("XSpTRdaHrc1") then
			Tmpb = Tmpb + "Hrc1: " & cstr( LottXHrc1) & " < " &  cstr(rs_s("XSpTRdaHrc1"))  & " "
		end if
		if LottXHrc1 > rs_s("XSpTRaHrc1") then
			Tmpb = Tmpb + "Hrc1: " &  cstr(LottXHrc1) & " > " &  cstr(rs_s("XSpTRaHrc1"))  & " "
		end if
	
	end if
	
	if (LottXHrc1punto5 >= 0) and (rs_s("XSpTRdaHrc1punto5") > 0 or rs_s("XSpTRaHrc1punto5") < 999999) then
	'claudia 17/9/21
	if LottXHrc1punto5 < rs_s("XSpTRdaHrc1punto5") then
		Tmpb = Tmpb + "Hrc1.5: " & cstr( LottXHrc1punto5) & " < " &  cstr(rs_s("XSpTRdaHrc1punto5"))  & " "
	end if
	if LottXHrc1punto5 > rs_s("XSpTRaHrc1punto5") then
		Tmpb = Tmpb + "Hrc1.5: " &  cstr(LottXHrc1punto5) & " > " &  cstr(rs_s("XSpTRaHrc1punto5"))  & " "
	end if
	end if
	
	if (LottXHrc2 >= 0) and (rs_s("XSpTRdaHrc2") > 0 or rs_s("XSpTRaHrc2") < 999999) then

		'claudia 13/9/21 aggiunto hrc2
		if LottXHrc2 < rs_s("XSpTRdaHrc2") then
			Tmpb = Tmpb + "Hrc2: " & cstr( LottXHrc2) & " < " &  cstr(rs_s("XSpTRdaHrc2"))  & " "
		end if
		if LottXHrc1 > rs_s("XSpTRaHrc2") then
			Tmpb = Tmpb + "Hrc2: " &  cstr(LottXHrc2) & " > " &  cstr(rs_s("XSpTRaHrc2"))  & " "
		end if
	end if
	
	if (LottXHrc3 >= 0) and (rs_s("XSpTRdaHrc3") > 0 or rs_s("XSpTRaHrc3") < 999999) then
		if LottXHrc3 < rs_s("XSpTRdaHrc3")  then
			Tmpb = Tmpb + "Hrc3: " & cstr(LottXHrc3) & " < " & cstr(rs_s("XSpTRdaHrc3")) & " "
		end if
		if LottXHrc3 > rs_s("XSpTRaHrc3")  then
			Tmpb = Tmpb + "Hrc3: " & cstr(LottXHrc3) & " > " & cstr(rs_s("XSpTRaHrc3")) & " "
		end if
	end if
	
	if (LottXHrc5 >= 0) and (rs_s("XSpTRdaHrc5") > 0 or rs_s("XSpTRaHrc5") < 999999) then					
		if LottXHrc5 < rs_s("XSpTRdaHrc5") then
			Tmpb = Tmpb + "Hrc5: " & cstr(LottXHrc5) & " < " & cstr(rs_s("XSpTRdaHrc5"))  & " "
		end if
		if LottXHrc5 > rs_s("XSpTRaHrc5") then
			Tmpb = Tmpb + "Hrc5: " & cstr(LottXHrc5) & " > " & cstr(rs_s("XSpTRaHrc5"))  & " "
		end if							
	end if
	
	if (LottXHrc7 >= 0) and (rs_s("XSpTRdaHrc7") > 0 or rs_s("XSpTRaHrc7") < 999999) then
		if LottXHrc7 < rs_s("XSpTRdaHrc7") then
			Tmpb = Tmpb + "Hrc7: " & cstr(LottXHrc7) & " < " & cstr(rs_s("XSpTRdaHrc7"))  & " "
		end if
		if LottXHrc7 > rs_s("XSpTRaHrc7") then
			Tmpb = Tmpb + "Hrc7: " & cstr(LottXHrc7) & " > " & cstr(rs_s("XSpTRaHrc7"))  & " "
		end if	
	end if
	
	if (LottXHrc9 >= 0) and (rs_s("XSpTRdaHrc9") > 0 or rs_s("XSpTRaHrc9") < 999999) then
		if LottXHrc9 < rs_s("XSpTRdaHrc9") then
			Tmpb = Tmpb + "Hrc9: " & cstr(LottXHrc9) & " < " & cstr(rs_s("XSpTRdaHrc9"))  & " "
		end if
		if LottXHrc9 > rs_s("XSpTRaHrc9") then
			Tmpb = Tmpb + "Hrc9: " & cstr(LottXHrc9) & " > " & cstr(rs_s("XSpTRaHrc9"))  & " "
		end if		
	end if
	
	if (LottXHrc10 >= 0) and (rs_s("XSpTRdaHrc10") > 0 or rs_s("XSpTRaHrc10") < 999999) then						
		if LottXHrc10 < rs_s("XSpTRdaHrc10") then
			Tmpb = Tmpb + "Hrc10: " & cstr(LottXHrc10) & " < " & cstr(rs_s("XSpTRdaHrc10"))  & " "
		end if
		if LottXHrc10 > rs_s("XSpTRaHrc10") then
			Tmpb = Tmpb + "Hrc10: " & cstr(LottXHrc10) & " > " & cstr(rs_s("XSpTRaHrc10"))  & " "
		end if		
	end if
	
	if (LottXHrc11 >= 0) and (rs_s("XSpTRdaHrc11") > 0 or rs_s("XSpTRaHrc11") < 999999) then
		if LottXHrc11 < rs_s("XSpTRdaHrc11") then
			Tmpb = Tmpb + "Hrc11: " & cstr(LottXHrc11) & " < " & cstr(rs_s("XSpTRdaHrc11"))  & " "
		end if
		if LottXHrc11 > rs_s("XSpTRaHrc11") then
			Tmpb = Tmpb + "Hrc11: " & cstr(LottXHrc11) & " > " & cstr(rs_s("XSpTRaHrc11"))  & " "
		end if		
	end if
	
	if (LottXHrc13 >= 0) and (rs_s("XSpTRdaHrc13") > 0 or rs_s("XSpTRaHrc13") < 999999) then
		if LottXHrc13 < rs_s("XSpTRdaHrc13") then
			Tmpb = Tmpb + "Hrc13: " & cstr(LottXHrc13) & " < " & cstr(rs_s("XSpTRdaHrc13"))  & " "
		end if
		if LottXHrc13 > rs_s("XSpTRaHrc13") then
			Tmpb = Tmpb + "Hrc13: " & cstr(LottXHrc13) & " > " & cstr(rs_s("XSpTRaHrc13"))  & " "
		end if		
	end if
	
	if (LottXHrc15 >= 0) and (rs_s("XSpTRdaHrc15") > 0 or rs_s("XSpTRaHrc15") < 999999) then
		if LottXHrc15 < rs_s("XSpTRdaHrc15") then
			Tmpb = Tmpb + "Hrc15: " & cstr(LottXHrc15) & " < " & cstr(rs_s("XSpTRdaHrc15"))  & " "
		end if
		if LottXHrc15 > rs_s("XSpTRaHrc15") then
			Tmpb = Tmpb + "Hrc15: " & cstr(LottXHrc15) & " > " & cstr(rs_s("XSpTRaHrc15"))  & " "
		end if		
	end if
	
	if (LottXHrc20 >= 0) and (rs_s("XSpTRdaHrc20") > 0 or rs_s("XSpTRaHrc20") < 999999) then
		if LottXHrc20 < rs_s("XSpTRdaHrc20") then
			Tmpb = Tmpb + "Hrc20: " & cstr(LottXHrc20) & " < " & cstr(rs_s("XSpTRdaHrc20"))  & " "
		end if
		if LottXHrc20 > rs_s("XSpTRaHrc20") then
			Tmpb = Tmpb + "Hrc20: " & cstr(LottXHrc20) & " > " & cstr(rs_s("XSpTRaHrc20"))  & " "
		end if		
	end if
	
	if (LottXHrc25 >= 0) and (rs_s("XSpTRdaHrc25") > 0 or rs_s("XSpTRaHrc25") < 999999) then
		if LottXHrc25 < rs_s("XSpTRdaHrc25") then
			Tmpb = Tmpb + "Hrc25: " & cstr(LottXHrc25) & " < " & cstr(rs_s("XSpTRdaHrc25"))  & " "
		end if
		if LottXHrc25 > rs_s("XSpTRaHrc25") then
			Tmpb = Tmpb + "Hrc25: " & cstr(LottXHrc25) & " > " & cstr(rs_s("XSpTRaHrc25"))  & " "
		end if		
	end if
	err.clear

	if (LottXHrc30 >= 0) and (rs_s("XSpTRdaHrc30") > 0 or rs_s("XSpTRaHrc30") < 999999) then	
		if LottXHrc30 < rs_s("XSpTRdaHrc30") then
			Tmpb = Tmpb + "Hrc30: " & cstr(LottXHrc30) & " < " & cstr(rs_s("XSpTRdaHrc30"))  & " "
		end if
		if LottXHrc30 > rs_s("XSpTRaHrc30") then
			Tmpb = Tmpb + "Hrc30: " & cstr(LottXHrc30) & " > " & cstr(rs_s("XSpTRaHrc30"))  & " "
		end if	
	end if
	
	if (LottXHrc35 >= 0) and (rs_s("XSpTRdaHrc35") > 0 or rs_s("XSpTRaHrc35") < 999999) then

		if LottXHrc35 < rs_s("XSpTRdaHrc35") then
			Tmpb = Tmpb + "Hrc35: " & cstr(LottXHrc35) & " < " & cstr(rs_s("XSpTRdaHrc35"))  & " "
		end if
		if LottXHrc35 > rs_s("XSpTRaHrc35") then
			Tmpb = Tmpb + "Hrc35: " & cstr(LottXHrc35) & " > " & cstr(rs_s("XSpTRaHrc35"))  & " "
		end if		
	end if
	
	if (LottXHrc40 >= 0) and (rs_s("XSpTRdaHrc40") > 0 or rs_s("XSpTRaHrc40") < 999999) then
		if LottXHrc40 < rs_s("XSpTRdaHrc40") then
			Tmpb = Tmpb + "Hrc40: " & cstr(LottXHrc40) & " < " & cstr(rs_s("XSpTRdaHrc40"))  & " "
		end if
		if LottXHrc40 > rs_s("XSpTRaHrc40") then
			Tmpb = Tmpb + "Hrc40: " & cstr(LottXHrc40) & " > " & cstr(rs_s("XSpTRaHrc40"))  & " "
		end if		
	end if
	
	if (LottXHrc45 >= 0) and (rs_s("XSpTRdaHrc45") > 0 or rs_s("XSpTRaHrc45") < 999999) then
		if LottXHrc45 < rs_s("XSpTRdaHrc45") then
			Tmpb = Tmpb + "Hrc45: " & cstr(LottXHrc45) & " < " & cstr(rs_s("XSpTRdaHrc45"))  & " "
		end if
		if LottXHrc45 > rs_s("XSpTRaHrc45") then
			Tmpb = Tmpb + "Hrc45: " & cstr(LottXHrc45) & " > " & cstr(rs_s("XSpTRaHrc45"))  & " "
		end if		
	end if
	
	if (LottXHrc50 >= 0) and (rs_s("XSpTRdaHrc50") > 0 or rs_s("XSpTRaHrc50") < 999999) then
		if LottXHrc50 < rs_s("XSpTRdaHrc50") then
			Tmpb = Tmpb + "Hrc50: " & cstr(LottXHrc50) & " < " & cstr(rs_s("XSpTRdaHrc50"))  & " "
		end if
		if LottXHrc50 > rs_s("XSpTRaHrc50") then
			Tmpb = Tmpb + "Hrc50: " & cstr(LottXHrc50) & " > " & cstr(rs_s("XSpTRaHrc50"))  & " "
		end if	
	end if
	
	
	'Controllo Vari -------------------------------------

	dim Vari
	Dim LottXRm1 	: LottXRm1 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXRm1"))	
	Dim LottXRs1 	: LottXRs1 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXRs1"))
	Dim LottXA51	: LottXA51 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXA51"))
	Dim LottXZ		: LottXZ = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXZ"))
	Dim LottXKcu	: LottXKcu = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXKcu"))
	Dim LottXGrano 	: LottXGrano = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXGrano"))	
	Dim LottXGranoA : LottXGranoA = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXGranoA")) 'claudia 29/7/21
	'DAVIDE 30/08/21 COMMENTATO MICRO SU RICHIESTA DI TOMMASO.
	'Dim LottXMicr	: LottXMicr = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXMicropurezza"))	
	Dim LottXDiam 	: LottXDiam  = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXDiamIdeal"))	
	Dim LottXHb		: LottXHb = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHb"))
	'Claudia 10/9/21 COMMENTATO hb1 SU RICHIESTA DI Stefano.
	'Dim LottXHb1	: LottXHb1 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXHb1"))
	'claudia 10/9/21 nuovi campi
	Dim LottXKU 	: LottXKU = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXKU"))	
	Dim LottXKV 	: LottXKV = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXKV"))
	Dim LottXMicroS	: LottXMicroS = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXMicroS"))
	Dim LottXMicroC	: LottXMicroC = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXMicroC"))
	Dim LottXR		: LottXR = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXR"))
	Dim LottXBandatura : LottXBandatura = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXBandatura"))
	
	if (LottXRm1 >= 0) and (rs_s("XSpTRdaRm1") > 0 or rs_s("XSpTRaRm1") < 999999) then
		if  LottXRm1 < rs_s("XSpTRdaRm1") then
			Tmpb = Tmpb + "Rm1: " & cstr( LottXRm1)  &  " < " & cstr(rs_s("XSpTRdaRm1")) & " "
		end if	
		if  LottXRm1 >  rs_s("XSpTRaRm1") then
			Tmpb = Tmpb + "Rm1: " &  cstr( LottXRm1)  &  " > " & cstr(rs_s("XSpTRaRm1")) & " "
		end if
	end if
	
	if (LottXRs1 >= 0) and (rs_s("XSpTRdars1") > 0 or rs_s("XSpTRaRs1") < 999999) then
		if LottXRs1 < rs_s("XSpTRdaRs1") then
			Tmpb = Tmpb + "Rs1: " &  cstr(LottXRs1) & " < " &  cstr(rs_s("XSpTRdaRs1")) & " "
		end if
		if LottXRs1 > rs_s("XSpTRaRs1") then
			Tmpb = Tmpb + "Rs1: " &  cstr(LottXRs1) & " > " &  cstr(rs_s("XSpTRaRs1")) & " "
		end if
	end if
	
	if (LottXA51 >= 0) and (rs_s("XSpTRdaA51") > 0 or rs_s("XSpTRaA51") < 999999) then
			if LottXA51 < rs_s("XSpTRdaA51") then
				Tmpb = Tmpb + "A51: " & cstr(LottXA51) & " < " & cstr(rs_s("XSpTRdaA51"))  & " "
			end if
			if LottXA51 > rs_s("XSpTRaA51") then
				Tmpb = Tmpb + "A51: " & cstr(LottXA51) & " > " & cstr(rs_s("XSpTRaA51"))  & " "
			end if
	end if
	
	if (LottXZ >= 0) and (rs_s("XSpTRdaZ") > 0 or rs_s("XSpTRaZ") < 999999) then
		if LottXZ < rs_s("XSpTRdaZ")  then
			Tmpb = Tmpb + "Z: "   & cstr(LottXZ)   & " < " & cstr(rs_s("XSpTRdaZ"))   &  " "
		end if
		if LottXZ > rs_s("XSpTRaZ")  then
			Tmpb = Tmpb + "Z: "   & cstr(LottXZ)   & " > " & cstr(rs_s("XSpTRaZ"))    &  " "
		end if
	end if
	
	if (LottXKcu >= 0) and (rs_s("XSpTRdakcu") > 0 or rs_s("XSpTRaKcu") < 999999) then
		if  LottXKcu < rs_s("XSpTRdaKCU") then
			Tmpb = Tmpb + "Kcu: " & cstr(LottXKcu)  & " < " & cstr(rs_s("XSpTRdaKCU"))  & " "
		end if
		if  LottXKcu > rs_s("XSpTRaKCU") then
			Tmpb = Tmpb + "Kcu: " & cstr(LottXKcu)  & " > " & cstr(rs_s("XSpTRaKCU"))  & " "
		end if
	end if
	
	if (LottXGrano >= 0) and (rs_s("XSpTRdaGrano") > 0 or rs_s("XSpTRaGrano") < 999999) then
		if LottXGrano <  rs_s("XSpTRdaGrano")  then
			Tmpb = Tmpb + "Grano: " & cstr(LottXGrano) &  " < " & cstr(rs_s("XSpTRdaGrano")) &  " "
		end if
		'claudia 29/7/21 gestito estremo superiore grano del lotto (nuovo campo)
		if LottXGranoA >  rs_s("XSpTRaGrano")  then
			Tmpb = Tmpb + "Grano: " & cstr(LottXGranoA) &  " > "& cstr(rs_s("XSpTRaGrano")) &  " "
		end if
	end if
	'DAVIDE 30/08/21 COMMENTATO MICRO SU RICHIESTA DI TOMMASO.
	'if LottXMicr < rs_s("XSpTRdaMicro") then
	'	Tmpb = Tmpb + "Micro: " & cstr(LottXMicr) & " < " & cstr(rs_s("XSpTRdaMicro"))  & " "
	'end if
	'if LottXMicr > rs_s("XSpTRaMirco") then
	'	Tmpb = Tmpb + "Micro: " & cstr(LottXMicr) & " > " & cstr(rs_s("XSpTRaMirco"))  & " "
	'end if
	
	
	if (LottXDiam >= 0) and (rs_s("XSpTRdaDiam") > 0 or rs_s("XSpTRaDiam") < 999999) then
		if LottXDiam < rs_s("XSpTRdaDiam") then
			Tmpb = Tmpb + "Diam: "  & cstr(LottXDiam) & " < " & cstr(rs_s("XSpTRdaDiam")) & " "
		end if
		if LottXDiam > rs_s("XSpTRaDiam") then
			Tmpb = Tmpb + "Diam: "  & cstr(LottXDiam) & " > " & cstr(rs_s("XSpTRaDiam")) & " "
		end if
	end if
	
	if (LottXHb >= 0) and (rs_s("XSpTRdaHb") > 0 or rs_s("XSpTRaHb") < 999999) then
		if LottXHb < rs_s("XSpTRdaHb") then
			Tmpb = Tmpb + "Diam: "  & cstr(LottXHb) & " < " & cstr(rs_s("XSpTRdaHb")) & " "
		end if
		if LottXHb > rs_s("XSpTRaHb") then
			Tmpb = Tmpb + "Diam: "  & cstr(LottXHb) & " > " & cstr(rs_s("XSpTRaHb")) & " "
		end if
	end if
	
	
	'Claudia 10/9/21 COMMENTATO hb1 SU RICHIESTA DI Stefano.
	'if LottXHb1 < rs_s("XSpTRdaDiam") then
	'	Tmpb = Tmpb + "Diam: "  & cstr(LottXHb1) & " < " & cstr(rs_s("XSpTRdaHb1")) & " "
	'end if
	'if LottXHb1 > rs_s("XSpTRaDiam") then
	'	Tmpb = Tmpb + "Diam: "  & cstr(LottXHb1) & " > " & cstr(rs_s("XSpTRaHb1")) & " "
	'end if
	'inizio Claudia 10/9/21 nuovi campi
	
	if (LottXKU >= 0) and (rs_s("XSpTRdaKU") > 0 or rs_s("XSpTRaKU") < 999999) then
		if  LottXKU < rs_s("XSpTRdaKU") then
			Tmpb = Tmpb + "KU: " & cstr( LottXKU)  &  " < " & cstr(rs_s("XSpTRdaKU")) & " "
		end if	
		if  LottXKU >  rs_s("XSpTRaKU") then
			Tmpb = Tmpb + "KU: " &  cstr( LottXKU)  &  " > " & cstr(rs_s("XSpTRaKU")) & " "
		end if
	end if
	
	if (LottXKV >= 0) and (rs_s("XSpTRdaKV") > 0 or rs_s("XSpTRaKV") < 999999) then
		if LottXKV < rs_s("XSpTRdaKV") then
			Tmpb = Tmpb + "KV: " &  cstr(LottXKV) & " < " &  cstr(rs_s("XSpTRdaKV")) & " "
		end if
		if LottXKV > rs_s("XSpTRaKV") then
			Tmpb = Tmpb + "KV: " &  cstr(LottXKV) & " > " &  cstr(rs_s("XSpTRaKV")) & " "
		end if
	end if
	
	if (LottXMicroS >= 0) and (rs_s("XSpTRdaMicroS") > 0 or rs_s("XSpTRaMicroS") < 999999) then
	
		if LottXMicroS < rs_s("XSpTRdaMicroS") then
			Tmpb = Tmpb + "Macro S: " & cstr(LottXMicroS) & " < " & cstr(rs_s("XSpTRdaMicroS"))  & " "
		end if
		if LottXMicroS > rs_s("XSpTRaMicroS") then
			Tmpb = Tmpb + "Macro S: " & cstr(LottXMicroS) & " > " & cstr(rs_s("XSpTRaMicroS"))  & " "
		end if
	end if
	
	if (LottXMicroC >= 0) and (rs_s("XSpTRdaMicroc") > 0 or rs_s("XSpTRaMicroc") < 999999) then
		if LottXMicroC < rs_s("XSpTRdaMicroc")  then
			Tmpb = Tmpb + "Macro C: "   & cstr(LottXMicroC)   & " < " & cstr(rs_s("XSpTRdaMicroC"))   &  " "
		end if
		if LottXMicroC > rs_s("XSpTRaMicroC")  then
			Tmpb = Tmpb + "Macro C: "   & cstr(LottXMicroC)   & " > " & cstr(rs_s("XSpTRaMicroC"))    &  " "
		end if
	end if
	
	if (LottXR >= 0) and (rs_s("XSpTRdaR") > 0 or rs_s("XSpTRaR") < 999999) then
		if  LottXR < rs_s("XSpTRdaR") then
			Tmpb = Tmpb + "R: " & cstr(LottXR)  & " < " & cstr(rs_s("XSpTRdaR"))  & " "
		end if
		if  LottXR > rs_s("XSpTRaR") then
			Tmpb = Tmpb + "R: " & cstr(LottXR)  & " > " & cstr(rs_s("XSpTRaR"))  & " "
		end if
	end if
	
	if (LottXBandatura >= 0) and (rs_s("XSpTRdaBandatura") > 0 or rs_s("XSpTRaBandatura") < 999999) then
		if  LottXBandatura < rs_s("XSpTRdaBandatura") then
			Tmpb = Tmpb + "Bandatura: " & cstr(LottXBandatura)  & " < " & cstr(rs_s("XSpTRdaBandatura"))  & " "
		end if
		if  LottXBandatura > rs_s("XSpTRaR") then
			Tmpb = Tmpb + "Bandatura: " & cstr(XSpTRaBandatura)  & " > " & cstr(rs_s("XSpTRaBandatura"))  & " "
		end if
		'fine 10/9/21
	end if
	
	'Controllo Gas ---------------------------------------------
	
	
	dim Gas
	Dim LottXH2 	: LottXH2 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXH2"))
	Dim LottXO2 	: LottXO2 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXO2"))
	Dim LottXN2 	: LottXN2 = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXN2"))
	
	
	
	if (LottXH2 >= 0) and (rs_s("XSpTRdaH2") > 0 or rs_s("XSpTRaH2") < 999999) then
		if LottXH2 < rs_s("XSpTRdaH2") then
			Gas = Gas  + "H2: "  & cstr(LottXH2) & " < " & cstr(rs_s("XSpTRdaH2")) & " "
		end if
		if LottXH2 > rs_s("XSpTRaH2") then
			Gas = Gas  + "H2: "  & cstr(LottXH2) & " > " & cstr(rs_s("XSpTRaH2")) & " "
		end if
	end if
	
	if (LottXO2 >= 0) and (rs_s("XSpTRdaO2") > 0 or rs_s("XSpTRaO2") < 999999) then
		if LottXO2 < rs_s("XSpTRdaO2") then
			Gas = Gas  + "O2: "  & cstr(LottXO2) & " < " & cstr(rs_s("XSpTRdaO2")) & " "
		end if
		if LottXO2 > rs_s("XSpTRaO2") then
			Gas = Gas  + "O2: "  & cstr(LottXO2) & " > " & cstr(rs_s("XSpTRaO2")) & " "
		end if		
	end if
	
	if (LottXN2 >= 0) and (rs_s("XSpTRdaN2") > 0 or rs_s("XSpTRaN2") < 999999) then
		if LottXN2 < rs_s("XSpTRdaN2") then
			Gas = Gas  + "N2: "  & cstr(LottXN2) & " < " & cstr(rs_s("XSpTRdaN2")) & " "
		end if
		if LottXN2 > rs_s("XSpTRaN2") then
			Gas = Gas  + "N2: "  & cstr(LottXN2) & " > " & cstr(rs_s("XSpTRaN2")) & " "
		end if
	end if
	
	'Controllo Micro
	dim Micro
	Dim LottXAFine 		: LottXAFine = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXAFine"))
	Dim LottXASpesso 	: LottXASpesso = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXASpesso"))
	Dim LottXBFine 		: LottXBFine = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXBFine"))
	Dim LottXBSpesso 	: LottXBSpesso = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXBSpesso"))
	Dim LottXCFine 		: LottXCFine = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXCFine"))
	Dim LottXCSpesso 	: LottXCSpesso = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXCSpesso"))
	Dim LottXDFine 		: LottXDFine = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXDFine"))
	Dim LottXDSpesso 	: LottXDSpesso = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXDSpesso"))
	Dim LottXKOxide 	: LottXKOxide = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXKOxide"))
	Dim LottXKSulfide 	: LottXKSulfide = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXKSulfide"))
	Dim LottXKTotal 	: LottXKTotal = CDbl(Ilotti.GetCollection().GetCurrent().getValue("LottXKTotal"))
	
	
	
	if (LottXAFine >= 0) and (rs_s("XSpTRdaAFine") > 0 or rs_s("XSpTRaAFine") < 999999) then
		if LottXAFine < rs_s("XSpTRDaAFine") then
			Micro = Micro  + "A Fine: "  & cstr(LottXAFine) & " < " & cstr(rs_s("XSpTRDaAFine")) & " "
		end if
		if LottXAFine > rs_s("XSpTRAAFine") then
			Micro = Micro  + "A Fine: "  & cstr(LottXAFine) & " > " & cstr(rs_s("XSpTRAAFine")) & " "
		end if
	end if
	
	if (LottXASpesso >= 0) and (rs_s("XSpTRdaASpesso") > 0 or rs_s("XSpTRaASpesso") < 999999) then
		if LottXASpesso < rs_s("XSpTRDaASpesso") then
			Micro = Micro  + "A Spesso: "  & cstr(LottXASpesso) & " < " & cstr(rs_s("XSpTRDaASpesso")) & " "
		end if
		if LottXASpesso > rs_s("XSpTRAASpesso") then
			Micro = Micro  + "A Spesso: "  & cstr(LottXASpesso) & " > " & cstr(rs_s("XSpTRAASpesso")) & " "
		end if
	end if
	
	if (LottXBFine >= 0) and (rs_s("XSpTRdaBFine") > 0 or rs_s("XSpTRaBFine") < 999999) then
		if LottXBFine < rs_s("XSpTRDaBFine") then
			Micro = Micro  + "B Fine: "  & cstr(LottXBFine) & " < " & cstr(rs_s("XSpTRDaBFine")) & " "
		end if
		if LottXBFine > rs_s("XSpTRABFine") then
			Micro = Micro  + "B Fine: "  & cstr(LottXBFine) & " > " & cstr(rs_s("XSpTRABFine")) & " "
		end if
	
	end if
	
	if (LottXBSpesso >= 0) and (rs_s("XSpTRdaBSpesso") > 0 or rs_s("XSpTRaBSpesso") < 999999) then
		if LottXBSpesso < rs_s("XSpTRDaBSpesso") then
			Micro = Micro  + "B Spesso: "  & cstr(LottXBSpesso) & " < " & cstr(rs_s("XSpTRDaBSpesso")) & " "
		end if
		if LottXBSpesso > rs_s("XSpTRABSpesso") then
			Micro = Micro  + "B Spesso: "  & cstr(LottXBSpesso) & " > " & cstr(rs_s("XSpTRABSpesso")) & " "
		end if
		
	end if
	
	if (LottXN2 >= 0) and (rs_s("XSpTRdaCFine") > 0 or rs_s("XSpTRaCFine") < 999999) then
		if LottXCFine < rs_s("XSpTRDaCFine") then
			Micro = Micro  + "C Fine: "  & cstr(LottXCFine) & " < " & cstr(rs_s("XSpTRDaCFine")) & " "
		end if
		if LottXCFine > rs_s("XSpTRACFine") then
			Micro = Micro  + "C Fine: "  & cstr(LottXCFine) & " > " & cstr(rs_s("XSpTRACFine")) & " "
		end if
	end if
	
	if (LottXCSpesso >= 0) and (rs_s("XSpTRdaCSpesso") > 0 or rs_s("XSpTRaCSpesso") < 999999) then
		if LottXCSpesso < rs_s("XSpTRDaCSpesso") then
			Micro = Micro  + "C Spesso: "  & cstr(LottXCSpesso) & " < " & cstr(rs_s("XSpTRDaCSpesso")) & " "
		end if
		if LottXCSpesso > rs_s("XSpTRACSpesso") then
			Micro = Micro  + "C Spesso: "  & cstr(LottXCSpesso) & " > " & cstr(rs_s("XSpTRACSpesso")) & " "
		end if
	
	end if
	
	if (LottXDFine >= 0) and (rs_s("XSpTRdaDFine") > 0 or rs_s("XSpTRaDFine") < 999999) then
		if LottXDFine < rs_s("XSpTRDaDFine") then
			Micro = Micro  + "D Fine: "  & cstr(LottXDFine) & " < " & cstr(rs_s("XSpTRDaDFine")) & " "
		end if
		if LottXDFine > rs_s("XSpTRADFine") then
			Micro = Micro  + "D Fine: "  & cstr(LottXDFine) & " > " & cstr(rs_s("XSpTRADFine")) & " "
		end if
	end if
	
	if (LottXKOxide >= 0) and (rs_s("XSpTRdaKOxide") > 0 or rs_s("XSpTRaKOxide") < 999999) then
			if CDbl(LottXKOxide ) < rs_s("XSpTRDaKOxide") then
				Micro = Micro  + "K Oxide: "  & cstr(LottXKOxide) & " < " & cstr(rs_s("XSpTRDaKOxide")) & " "
			end if
			if CDbl(LottXKOxide ) > rs_s("XSpTRAKOxide") then
				Micro = Micro  + "K Oxide: "  & cstr(LottXKOxide) & " > " & cstr(rs_s("XSpTRAKOxide")) & " "
			end if
	end if
	
	if (LottXKSulfide >= 0) and (rs_s("XSpTRdaKSulfide") > 0 or rs_s("XSpTRaKSulfide") < 999999) then
		if LottXKSulfide < rs_s("XSpTRDaKSulfide") then
			Micro = Micro  + "K Sulfide: "  & cstr(LottXKSulfide) & " < " & cstr(rs_s("XSpTRDaKSulfide")) & " "
		end if
		if LottXKSulfide > rs_s("XSpTRAKSulfide") then
			Micro = Micro  + "K Sulfide: "  & cstr(LottXKSulfide) & " > " & cstr(rs_s("XSpTRAKSulfide")) & " "
		end if
	end if
	
	if (LottXKTotal >= 0) and (rs_s("XSpTRdaKTotal") > 0 or rs_s("XSpTRaKTotal") < 999999) then
	
		if LottXKTotal < rs_s("XSpTRDaKTotal") then
			Micro = Micro  + "K Total: "  & cstr(LottXKTotal) & " < " & cstr(rs_s("XSpTRDaKTotal")) & " "
		end if
		if LottXKTotal > rs_s("XSpTRAKTotal") then
			Micro = Micro  + "K Total: "  & cstr(LottXKTotal) & " > " & cstr(rs_s("XSpTRAKTotal")) & " "
		end if		
	end if
	
	'claudia 21/9/21 visualizza il titolo del capitolato se esiste altrimenti la descrizione del soggetto	
	dim sTitoloSoggetto
	if rs_s("XspTRDTitolo") <> "" then
		sTitoloSoggetto = rs_s("XspTRDTitolo")
	else
		sTitoloSoggetto= rs_s("AsogDAsog")
	end if
	If Gas <> "" or Vari <> "" or Tmpb <> "" or ElCh <> "" or Micro <> "" then
		SpeTec = ElCh & Tmpb & Vari & Gas & Micro
	'Claudia 7/6/21
		outNonCompatibile = outNonCompatibile +sTitoloSoggetto  &  vbCrLf &  SpeTec  & vbCrLf& vbCrLf
	else
		SpeTec = "Compatibile"
		
		'Claudia 7/6/21
		outCompatibile = outCompatibile + sTitoloSoggetto + " - "
		
	end if
	
'msgbox ("Cliente: " & rs_s("AsogDAsog") & " " & SpeTec )

End Function
'----------------------------------------------------------
'Lista Campi Interattore ( nome [ label ] { tipo } )
'----------------------------------------------------------

' LottXBarre2Fasci [ Composizione Fasci ]  { CEditControl }
' LottXH2 [ H2 ]  { CCurrencyEditControl }
' LottXCSpesso [ C Spesso ]  { CCurrencyEditControl }
' LottXC [ C ]  { CCurrencyEditControl }
' LottXDSpesso [ D Spesso ]  { CCurrencyEditControl }
' LottXKSulfide [ K Sulfide ]  { CCurrencyEditControl }
' LottXQuaCert [ Qualita Certificato ]  { CEditControl }
' LottXColataMatricola [ Colata o Matricola ]  { CEditControl }
' LottXNote1 [ Note ]  { CEditControl }
' LottXTipo [ Tipo ]  { CEditControl }
' LottXMo [ Mo ]  { CCurrencyEditControl }
' LottXHrc45 [ Hrc45 ]  { CCurrencyEditControl }
' LottXHrc35 [ Hrc35 ]  { CCurrencyEditControl }
' LottXCFine [ C fine ]  { CCurrencyEditControl }
' LottXAltreProve [ AltreProve ]  { CEditControl }
' LottXHrc10 [ Hrc10 ]  { CCurrencyEditControl }
' LottXAl [ Al ]  { CCurrencyEditControl }
' desLookupSlot [  ]  { CEditControl }
' LottXBarre1Lung [ Lunghezza ]  { CCurrencyEditControl }
' LottXAs [ As ]  { CCurrencyEditControl }
' LottXDFine [ D Fine ]  { CCurrencyEditControl }
' LottXBarre1Fasci [ Composizione Fasci ]  { CEditControl }
' LottXDiamIdeal [ Diametro Ideale mm ]  { CCurrencyEditControl }
' LottXSezione [ Sezione ]  { CCurrencyEditControl }
' desLookupTlot [  ]  { CEditControl }
' LottXBarre2KgCad [ Kg Cad. ]  { CCurrencyEditControl }
' LottXHrc11 [ Hrc11 ]  { CCurrencyEditControl }
' GiorniScadArtb [ GiorniScadArtb ]  { CEditControl }
' LottXHrc13 [ Hrc13 ]  { CCurrencyEditControl }
' LottXKv [ Kv ]  { CCurrencyEditControl }
' LottXHrc15 [ Hrc15 ]  { CCurrencyEditControl }
' LottXKU [ KU in J ]  { CCurrencyEditControl }
' lblLookupArtb [  ]  { CLinkControl }
' LottXKTotal [ K Total ]  { CCurrencyEditControl }
' LottXAFine [ A Fine ]  { CCurrencyEditControl }
' LottXNi [ Ni ]  { CCurrencyEditControl }
' LottXHrc9 [ Hrc9 ]  { CCurrencyEditControl }
' ButtonCopia [  ]  { CButtonExecuteControl }
' LottXHRC2 [ Hrc2 ]  { CCurrencyEditControl }
' LottXFb [ Sb ]  { CCurrencyEditControl }
' LottCLott [ Codice ]  { CEditControl }
' LottXHrc5 [ Hrc5 ]  { CCurrencyEditControl }
' LottXMn [ Mn ]  { CCurrencyEditControl }
' LottXN2 [ N2 ]  { CCurrencyEditControl }
' lblLookupTlot [  ]  { CLinkControl }
' desLookupSoggFor [  ]  { CEditControl }
' LottXA5 [ A5 ]  { CCurrencyEditControl }
' LottXMicroC [ C ]  { CCurrencyEditControl }
' LottXMicroS [ S ]  { CCurrencyEditControl }
' LookupArtb [ Articolo di Riferimento ]  { CLookupControl }
' LottXBarre1Nr [ PRIMA LUNGHEZZA  N.Barre ]  { CCurrencyEditControl }
' LottXPb [ Pb ]  { CCurrencyEditControl }
' LottXRappRid [ Rapp.Rid. ]  { CCurrencyEditControl }
' LottXHrc20 [ Hrc20 ]  { CCurrencyEditControl }
' LottXO2 [ O2 ]  { CCurrencyEditControl }
' LottXHrc1punto5 [ Hrc1,5 ]  { CCurrencyEditControl }
' LottXHrc25 [ Hrc25 ]  { CCurrencyEditControl }
' lblLookupSlot [  ]  { CLinkControl }
' GiorniValidita [ Giorni Validita' ]  { CDblEditControl }
' LottTInizio [ Data Inizio Validita' ]  { CDateTimeControl }
' LottXHrc30 [ Hrc30 ]  { CCurrencyEditControl }
' LottXHrc50 [ Hrc50 ]  { CCurrencyEditControl }
' desLookupArtb [  ]  { CEditControl }
' LottXHrC1 [ Hrc1 ]  { CCurrencyEditControl }
' LottXKCU [ KCU in J/mmq ]  { CCurrencyEditControl }
' LottXNb [ Nb ]  { CCurrencyEditControl }
' ButtonMostraSpecifiche [  ]  { CButtonExecuteControl }
' LottXCa [ Ca ]  { CCurrencyEditControl }
' LottXCe [ CE ]  { CCurrencyEditControl }
' LookupTlot [ Tipo ]  { CLookupControl }
' LottXHrc40 [ Hrc40 ]  { CCurrencyEditControl }
' FAttivo [ Attivo ]  { CComboBoxControl }
' LottXCr [ Cr ]  { CCurrencyEditControl }
' LottXCu [ Cu ]  { CCurrencyEditControl }
' LottXBarre1KgCad [ Kg Cad. ]  { CCurrencyEditControl }
' LottTcrea [ Data Creazione ]  { CDateTimeControl }
' LottXGranoA [ A Grano ]  { CCurrencyEditControl }
' LookupSoggFor [ Fornitore ]  { CLookupControl }
' LottTfine [ Data Fine Validita' ]  { CDateTimeControl }
' LottXKOxide [ K Oxide ]  { CCurrencyEditControl }
' LottXBarre2Nr [  SECONDA LUNGHEZZA  N.Barre ]  { CCurrencyEditControl }
' LottXAltreNote [ AltreNote ]  { CEditControl }
' LottXNumCert [ Numero Certificato ]  { CEditControl }
' FMultiArt [ Multi Articolo ]  { CComboBoxControl }
' LottXBSpesso [ B Spesso ]  { CCurrencyEditControl }
' LottXBarre2Lung [ Lunghezza ]  { CCurrencyEditControl }
' LottXTi [ Ti ]  { CCurrencyEditControl }
' LottCodRif [ Codice Esterno ]  { CEditControl }
' LottXBandatura [ Bandatura ]  { CCurrencyEditControl }
' LottXMateriale [ Materiale ]  { CEditControl }
' LottXBFine [ B Fine ]  { CCurrencyEditControl }
' LottXOrigine [ Origine ]  { CEditControl }
' LottXHb [ Hb ]  { CCurrencyEditControl }
' LottXSi [ Si ]  { CCurrencyEditControl }
' LottXSn [ Sn ]  { CCurrencyEditControl }
' LottXB [ B ]  { CCurrencyEditControl }
' LottXRs [ Rs ]  { CCurrencyEditControl }
' LottXHrc7 [ Hrc7 ]  { CCurrencyEditControl }
' LottXGrano [ Da Grano ]  { CCurrencyEditControl }
' LottXP [ P ]  { CCurrencyEditControl }
' LottXS [ S ]  { CCurrencyEditControl }
' LottXR [ R ]  { CCurrencyEditControl }
' LookupSlot [ Stato ]  { CLookupControl }
' LottXV [ V ]  { CCurrencyEditControl }
' LottXASpesso [ A Spesso ]  { CCurrencyEditControl }
' LottXZ [ Z ]  { CCurrencyEditControl }
' GroupBox1 [  ]  { CFrameControl }
' LottXHRC3 [ Hrc3 ]  { CCurrencyEditControl }
' LottXRm [ Rm ]  { CCurrencyEditControl }
' codLookupArtb [ Articolo di Riferimento ]  { CEditControl }
' buttLookupArtb [  ]  { CButtonExecuteControl }
' codLookupTlot [ Tipo ]  { CEditControl }
' buttLookupTlot [  ]  { CButtonExecuteControl }
' codLookupSoggFor [ Fornitore ]  { CEditControl }
' buttLookupSoggFor [  ]  { CButtonExecuteControl }
' codLookupSlot [ Stato ]  { CEditControl }
' buttLookupSlot [  ]  { CButtonExecuteControl }
' LottCser [  ]  { Interger }
' LottCRTlot [  ]  { Interger }
' LottCRSlot [  ]  { Interger }
' LottCRAsog_for [  ]  { Interger }
' LottFAttivoCSer [  ]  { Interger }
' LottFAttivoCAzione [  ]  { String }
' LottFMultiArtCSer [  ]  { Interger }
' LottFMultiArtCAzione [  ]  { String }
' LottCRArtb [  ]  { Interger }
' LottXProvetta [  ]  { Double }
' LottXUnita1 [  ]  { String }
' LottXUnita2 [  ]  { String }
' LottXUnita3 [  ]  { String }
' LottXUnita4 [  ]  { String }
' LottXTrattamento [  ]  { String }
' LottXTempRes [  ]  { String }
' LottXMicropurezza [  ]  { String }


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


