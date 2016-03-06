* obvykle se  mìní:
*	struktura dbf souborù - upravit na vývojové v Sablony\, pøekopírovat do ostré do Sablony\
*	pøípona dbf souborù - upravit Sablony\pripona.txt, pøekopírovat na ostrou
* neboli všechny zmìny jsou obvykle jen v adresáøi Sablony\

* 2015 v Ucto chtìjí místo xxx04 souborù nyní xxx15 - pøejmenoval jsem tedy Out soubory
* 2016 to vypadá, že to bude pravidelnì, ale nevíme, tak jsem udìlal soubor Sablony\pripona.txt, odkud si pøíponu naète (napø. 16)
*	šablonám zùstává pøípona ...15, pokud by se mìnila, tak zmìnit .cPriponaSabl

* na vývojovém pc mít \mz\cezaucto a vfp
* sestavit standalone exe, pøedat na ostrý exe a zmìnìné šablony

LOCAL llDevelop, lcIn, lcOut, lcProg, lcSablony, loCu
	llDevelop = DIRECTORY( 'c:\sarce' )
	
IF	m.llDevelop
	lcProg = "C:\mz\cezaucto\"    && bylo "C:\sarce\"
	lcIn = m.lcProg + "in\"	&& adr04.txt
	lcOut = m.lcProg + "out\"
ELSE
	ON ERROR DO ErrHnd
	lcProg = "C:\DOKUMENTY\CEZAUCTO\"
	lcIn = "c:\cezar\(export)\"	&& adr04.txt
	* lcOut = "c:\ucto2012\"	&& nahrazeno loCu.GetOutDir() - uctoRRRR s nejvyšším èíslem roku RRRR
ENDIF
lcSablony = m.lcProg + "Sablony\"

loCu = NEWOBJECT( 'CezaUcto', .NULL., .NULL., m.lcSablony )

loCu.cIn = m.lcIn
loCu.cOut = IIF( EMPTY( m.lcOut ), loCu.GetOutDir(), m.lcOut )
loCu.cSablony = m.lcSablony
loCu.cProg = m.lcProg

IF	NOT (m.loCu.SureDir( m.lcIn + m.loCu.cZpracDir ) AND m.loCu.SureDir( m.loCu.cOut + m.loCu.cZpracDir ))
	RETURN .F.
ENDIF

loCu.Odb()
loCu.Dod()
loCu.Den()
loCu.Poh()
loCu.Zav()

MESSAGEBOX( "Exporty byly provedeny do " + m.loCu.cOut, 64, "Dokonèeno" )
IF	m.llDevelop
	CLEAR ALL
ELSE
	QUIT
ENDIF

PROCEDURE ErrHnd
	LOCAL ARRAY laeErr[1], laeStack[1]
	LOCAL lni, lnj, lcMsg
		lcMsg = ''
	lnE = AERROR( laeErr )
	lnS = ASTACKINFO( laeStack )
	FOR lni=1 TO ALEN( laeErr, 1 )
		FOR lnj=1 TO 2	&& 1 TO 7
			IF	NOT ISNULL( laeErr[m.lni,m.lnj] ) AND NOT EMPTY( laeErr[m.lni,m.lnj] )
				lcMsg = IIF( EMPTY( m.lcMsg ), '', m.lcMsg + CHR(13)+CHR(10) ) + TRANSFORM( laeErr[m.lni,m.lnj] )
			ENDIF
		ENDFOR
	ENDFOR
	lcMsg = m.lcMsg + CHR(13)+CHR(10)
	lni = ALEN( laeStack, 1 )-1
	FOR lnj=5 TO 6	&& 1 TO 6
		IF	NOT ISNULL( laeStack[m.lni,m.lnj] ) AND NOT EMPTY( laeStack[m.lni,m.lnj] )
			lcMsg = IIF( EMPTY( m.lcMsg ), '', m.lcMsg + CHR(13)+CHR(10) ) + TRANSFORM( laeStack[m.lni,m.lnj] )
		ENDIF
	ENDFOR
	IF	MESSAGEBOX( m.lcMsg, 1, "Došlo k chybì programu" )=2
		CANCEL
	ENDIF


DEFINE CLASS CezaUcto AS Custom
	cZpracDir = 'Zpracovane'

	cIn = ''
	cOut = ''
	cSablony = ''
	cProg = ''
	
	cPriponaRok = ''
	cPriponaSabl = '15'

	PROCEDURE Init( tcSablony )
		THIS.cPriponaRok = LEFT(FILETOSTR(m.tcSablony + 'pripona.txt'), 2)
		
	PROCEDURE SureDir( tcDir )
		IF	NOT DIRECTORY( m.tcDir )
			TRY
				MD (m.tcDir)
			CATCH
			ENDTRY
			IF	NOT DIRECTORY( m.tcDir )
				MESSAGEBOX( "Nelze vytvoøit adresáø " + m.tcDir + " (zkuste jej vytvoøit ruènì).", 16, "Problém" )
				RETURN .F.
			ENDIF
		ENDIF

	PROCEDURE Odb
		LOCAL lcIn, lcOut
			lcIn = 'Adr04Odb.dbf'
			lcOut = 'Adr' + THIS.cPriponaRok + '.dbf'
			lcSabl = 'Adr' + THIS.cPriponaSabl
		IF	NOT THIS.BkIn( m.lcIn )
			RETURN .F.
		ENDIF
		THIS.BkOut( m.lcOut, m.lcSabl )
		SCAN && FOR AktivniOdb	&& BkOut selectuje Src
			INSERT INTO Trg ;
					(Firma, Jmeno, ;
					Ulice, PSC, Misto, Stat, ;
					Tlf, Mobil, Fax, Email, ;
					ICO10, DIC, Banka, Ucet) ;
					VALUES ;
					(Src.Nazev, Src.Kontakt, ;
					Src.DUlice, Src.DPSC, Src.FMesto, IIF(INLIST(UPPER( Src.KodStatu ),'CZ','CS'), '', Src.KodStatu), ;
					Src.Telefon, Src.Mobil, Src.Fax, Src.E_mail, ;
					IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), Src.DIC, '', Src.CisloUctu)
		ENDSCAN
		THIS.CloseTbls()
		
	PROCEDURE Dod
		LOCAL lcIn, lcOut
			lcIn = 'Adr04Dod.dbf'
			lcOut = 'Adr' + THIS.cPriponaRok + '.dbf'
			lcSabl = 'Adr' + THIS.cPriponaSabl
		IF	NOT THIS.BkIn( m.lcIn )
			RETURN .F.
		ENDIF
		THIS.BkOut( m.lcOut, m.lcSabl, .T. )
		SCAN && FOR AktivniDod	&& BkOut selectuje Src
			INSERT INTO Trg ;
					(Firma, Jmeno, ;
					Ulice, PSC, Misto, Stat, ;
					Tlf, Mobil, Fax, Email, ;
					ICO10, DIC, Banka, Ucet) ;
					VALUES ;
					(Src.Nazev, Src.Kontakt, ;
					Src.Ulice, Src.PSC, Src.Mesto, Src.KodStatu, ;
					Src.Telefon, Src.Mobil, Src.Fax, Src.E_mail, ;
					IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), Src.DIC, Src.Banka, Src.CisloUctu)
		ENDSCAN
		THIS.CloseTbls()

	PROCEDURE Den
		LOCAL lcIn, lcOut, Zvysena, Snizena
			STORE '' TO Zvysena, Snizena
			lcIn = 'Den04.dbf'
			lcOut = 'Den' + THIS.cPriponaRok + '.dbf'
			lcSabl = 'Den' + THIS.cPriponaSabl
		IF	NOT THIS.BkIn( m.lcIn, @m.Zvysena, @m.Snizena )
			RETURN .F.
		ENDIF
		THIS.BkOut( m.lcOut, m.lcSabl )
		SCAN	&& BkOut selectuje Src
			INSERT INTO Trg ;
					(Plat, Doklad, ICO, ;
					Druh, ;
					Text, Pozn, ;
					BezDaneZ, DPHZ, BezDaneS, DPHS, BezDane0, Celkem, ;
					Datum, DatumDPH) ;
					VALUES ;
					('H', Src.CisloPd, IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), ;
					IIF( LEFT( Src.CisloPd, 1 )=='1', 'PZJ', ;
							IIF( LEFT( Src.CisloPd, 1 )=='3', 'PZS', ;
							IIF( LEFT( Src.CisloPd, 1 )=='4', 'PZM', '' ) ) ), ;
					"Prodej zboží", ALLTRIM(Src.Odberatel) + IIF(EMPTY(Src.ICO),'',', ' + TRANSFORM(Src.ICO)), ;
					Src.Zaklad&Zvysena, Src.DPH&Zvysena, Src.Zaklad&Snizena, Src.DPH&Snizena, Src.Zaklad0, Src.Celkem, ;
					Src.Datum, Src.ZdanPlneni)
					* do 2012 bylo takto: Src.Zaklad20, Src.DPH20, Src.Zaklad14, Src.DPH14
		ENDSCAN
		THIS.CloseTbls()

	PROCEDURE Poh
		LOCAL lcIn, lcOut, Zvysena, Snizena
			STORE '' TO Zvysena, Snizena
			lcIn = 'Poh04.dbf'
			lcOut = 'Poh' + THIS.cPriponaRok + '.dbf'
			lcSabl = 'Poh' + THIS.cPriponaSabl
		IF	NOT THIS.BkIn( m.lcIn, @m.Zvysena, @m.Snizena )
			RETURN .F.
		ENDIF
		THIS.BkOut( m.lcOut, m.lcSabl )
		SCAN	&& BkOut selectuje Src
			INSERT INTO Trg ;
					(Plat, Doklad, Pozn, ;
					Text, ICO, ;
					Druh, ;
					BezDaneZ, DPHZ, BezDaneS, DPHS, BezDane0, Celkem, ;
					DatumVyst, DatumSpl, DatumDPH) ;
					VALUES ;
					(UPPER(Src.FormaUhrad), 'f/' + Src.CisloF, ALLTRIM(Src.Odberatel) + IIF(EMPTY(Src.ICO),'',', ' + TRANSFORM(Src.ICO)), ;
					"Prodej zboží", IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), ;
					IIF( LEFT( Src.CisloF, 2 )=='10' OR LEFT( Src.CisloF, 2 )=='12', 'PZJ', ;
							IIF( LEFT( Src.CisloF, 2 )=='13', 'PZS', ;
							IIF( LEFT( Src.CisloF, 2 )=='14', 'PZM', '' ) ) ), ;
					Src.Zaklad&Zvysena, Src.DPH&Zvysena, Src.Zaklad&Snizena, Src.DPH&Snizena, Src.Zaklad0, Src.Celkem, ;
					Src.Vystaveno, Src.Splatnost, Src.ZdanPlneni)
					* FormaUhrad=b/h
					* do 2012 bylo takto: Src.Zaklad20, Src.DPH20, Src.Zaklad14, Src.DPH14
		ENDSCAN
		THIS.CloseTbls()

	PROCEDURE Zav
		LOCAL lcIn, lcOut, Zvysena, Snizena
			STORE '' TO Zvysena, Snizena
			lcIn = 'Zav04.dbf'
			lcOut = 'Poh' + THIS.cPriponaRok + '.dbf'
			lcSabl = 'Poh' + THIS.cPriponaSabl
		IF	NOT THIS.BkIn( m.lcIn, @m.Zvysena, @m.Snizena )
			RETURN .F.
		ENDIF
		THIS.BkOut( m.lcOut, m.lcSabl, .T. )
		SCAN	&& BkOut selectuje Src
			INSERT INTO Trg ;
					(Plat, Doklad, Pozn, ;
					Text, ICO, ;
					Druh, ;
					BezDaneZ, DPHZ, BezDaneS, DPHS, BezDane0, Celkem, ;
					DatumVyst, DatumSpl, DatumDPH) ;
					VALUES ;
					(UPPER(Src.FormaUhrad), 'v/'+Src.CisloF, ALLTRIM(Src.Dodavatel) + IIF(EMPTY(Src.ICO),'',', ' + TRANSFORM(Src.ICO)), ;
					"Nákup zboží", IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), ;
					'NZ', ;
					Src.Zaklad&Zvysena, Src.DPH&Zvysena, Src.Zaklad&Snizena, Src.DPH&Snizena, Src.Zaklad0, Src.Celkem, ;
					Src.Vystaveno, Src.Splatnost, Src.ZdanPlneni)
					* FormaUhrad=b/h
					* do 2012 bylo takto: Src.Zaklad20, Src.DPH20, Src.Zaklad14, Src.DPH14
		ENDSCAN
		THIS.CloseTbls()

	PROCEDURE BkIn( tcTable, tcZvysena, tcSnizena )
		* zkopíruje vstupní soubor do backup adresáøe
		tcTable = THIS.cIn + m.tcTable
		IF	NOT FILE( m.tcTable )
			MESSAGEBOX( "Chybí zdrojový soubor " + m.tcTable, 48, "Pozor" )
			RETURN .F.
		ENDIF
		USE IN SELECT( 'Src' )
		SELECT 0
		USE (m.tcTable) ALIAS Src
		THIS.GetSazbyDane( @m.tcZvysena, @m.tcSnizena )
		COPY TO (THIS.cIn + ADDBS( THIS.cZpracDir ) + ;
						JUSTSTEM( m.tcTable ) + '_' + TTOC( DATETIME(), 1 ) + '.' + JUSTEXT( m.tcTable ))

	PROCEDURE GetSazbyDane( tcZvysena, tcSnizena )
		* @ tcZvysena, @ tcSnizena - sazby danì z názvu polí ZakladXX a DPHXX
		LOCAL lni, lcTatoSazba
		FOR lni=1 TO AFIELDS( lacFlds, 'Src' )
			IF	LOWER( LEFT( lacFlds[m.lni,1], 6 ) )=='zaklad'
				lcTatoSazba = SUBSTR( lacFlds[m.lni,1], 7 )
				IF	VAL( m.lcTatoSazba )>0
					IF	EMPTY( m.tcZvysena )
						tcZvysena = m.lcTatoSazba
					ENDIF
					IF	EMPTY( m.tcSnizena )
						tcSnizena = m.lcTatoSazba
					ENDIF
					IF	VAL( m.lcTatoSazba )>VAL( m.tcZvysena )
						tcZvysena = m.lcTatoSazba
					ENDIF
					IF	VAL( m.lcTatoSazba )<VAL( m.tcSnizena )
						tcSnizena = m.lcTatoSazba
					ENDIF
				ENDIF
			ENDIF
		ENDFOR
		
	PROCEDURE BkOut( tcTable, tcSablStem, tlAppend )
		* tlAppend=True : nezkopíruje se prázdný ze šablon, ale pokraèuje se plnit tento (u Dodavatelù)
		LOCAL llIsFile
		USE IN SELECT( 'Trg' )
		tcTable = THIS.cOut + m.tcTable
		llIsFile = FILE( m.tcTable )
		IF	m.llIsFile
			SELECT 0
			USE (m.tcTable) ALIAS Trg
			COPY TO (THIS.cOut + ADDBS( THIS.cZpracDir ) + ;
						JUSTSTEM( m.tcTable ) + '_' + TTOC( DATETIME(), 1 ) + ;
						IIF( m.tlAppend, '_cast_1', '' ) + ;
						'.' + JUSTEXT( m.tcTable ))
			USE IN SELECT( 'Trg' )
		ENDIF
		ERASE (FORCEEXT( m.tcTable, 'fpt' ))
		IF	NOT m.tlAppend OR NOT m.llIsFile
			* kopíruj šablonu
			ERASE (FORCEEXT( m.tcTable, 'cdx' ))
			ERASE (FORCEEXT( m.tcTable, 'dbt' ))
			ERASE (FORCEEXT( m.tcTable, 'dbf' ))
			SELECT 0
			USE (THIS.cSablony + m.tcSablStem) ALIAS Sablona
			COPY TO (m.tcTable) TYPE FOXPLUS AS 852
			USE IN SELECT( 'Sablona' )
		ENDIF
		SELECT 0
		USE (m.tcTable) ALIAS Trg
		SELECT Src	&& poslední pøíkaz !!!

	PROCEDURE CloseTbls
		USE IN SELECT( 'Src' )
		USE IN SELECT( 'Trg' )
	
	PROCEDURE GetOutDir
		* c:\uctoRRRR s nejvyšším èíslem roku RRRR
		LOCAL ARRAY lacFolders[1]
		LOCAL lni, lcPostFix, lcBeginWith, lcMax
			lcMax = ''
			lcBeginWith = 'C:\UCTO'
		FOR lni=1 TO ADIR( lacFolders, m.lcBeginWith + '*', 'D' )
			lcPostFix = SUBSTR( lacFolders[m.lni,1], 5 )
			IF	m.lcPostFix>m.lcMax OR EMPTY( m.lcMax )
				lcMax = m.lcPostFix
			ENDIF
		ENDFOR
		RETURN m.lcBeginWith + m.lcMax + '\'
ENDDEFINE
