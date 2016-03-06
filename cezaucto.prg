* obvykle se  m�n�:
*	struktura dbf soubor� - upravit na v�vojov� v Sablony\, p�ekop�rovat do ostr� do Sablony\
*	p��pona dbf soubor� - upravit Sablony\pripona.txt, p�ekop�rovat na ostrou
* neboli v�echny zm�ny jsou obvykle jen v adres��i Sablony\

* 2015 v Ucto cht�j� m�sto xxx04 soubor� nyn� xxx15 - p�ejmenoval jsem tedy Out soubory
* 2016 to vypad�, �e to bude pravideln�, ale nev�me, tak jsem ud�lal soubor Sablony\pripona.txt, odkud si p��ponu na�te (nap�. 16)
*	�ablon�m z�st�v� p��pona ...15, pokud by se m�nila, tak zm�nit .cPriponaSabl

* na v�vojov�m pc m�t \mz\cezaucto a vfp
* sestavit standalone exe, p�edat na ostr� exe a zm�n�n� �ablony

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
	* lcOut = "c:\ucto2012\"	&& nahrazeno loCu.GetOutDir() - uctoRRRR s nejvy���m ��slem roku RRRR
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

MESSAGEBOX( "Exporty byly provedeny do " + m.loCu.cOut, 64, "Dokon�eno" )
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
	IF	MESSAGEBOX( m.lcMsg, 1, "Do�lo k chyb� programu" )=2
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
				MESSAGEBOX( "Nelze vytvo�it adres�� " + m.tcDir + " (zkuste jej vytvo�it ru�n�).", 16, "Probl�m" )
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
					"Prodej zbo��", ALLTRIM(Src.Odberatel) + IIF(EMPTY(Src.ICO),'',', ' + TRANSFORM(Src.ICO)), ;
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
					"Prodej zbo��", IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), ;
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
					"N�kup zbo��", IIF(EMPTY(Src.ICO),'',TRANSFORM(Src.ICO)), ;
					'NZ', ;
					Src.Zaklad&Zvysena, Src.DPH&Zvysena, Src.Zaklad&Snizena, Src.DPH&Snizena, Src.Zaklad0, Src.Celkem, ;
					Src.Vystaveno, Src.Splatnost, Src.ZdanPlneni)
					* FormaUhrad=b/h
					* do 2012 bylo takto: Src.Zaklad20, Src.DPH20, Src.Zaklad14, Src.DPH14
		ENDSCAN
		THIS.CloseTbls()

	PROCEDURE BkIn( tcTable, tcZvysena, tcSnizena )
		* zkop�ruje vstupn� soubor do backup adres��e
		tcTable = THIS.cIn + m.tcTable
		IF	NOT FILE( m.tcTable )
			MESSAGEBOX( "Chyb� zdrojov� soubor " + m.tcTable, 48, "Pozor" )
			RETURN .F.
		ENDIF
		USE IN SELECT( 'Src' )
		SELECT 0
		USE (m.tcTable) ALIAS Src
		THIS.GetSazbyDane( @m.tcZvysena, @m.tcSnizena )
		COPY TO (THIS.cIn + ADDBS( THIS.cZpracDir ) + ;
						JUSTSTEM( m.tcTable ) + '_' + TTOC( DATETIME(), 1 ) + '.' + JUSTEXT( m.tcTable ))

	PROCEDURE GetSazbyDane( tcZvysena, tcSnizena )
		* @ tcZvysena, @ tcSnizena - sazby dan� z n�zvu pol� ZakladXX a DPHXX
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
		* tlAppend=True : nezkop�ruje se pr�zdn� ze �ablon, ale pokra�uje se plnit tento (u Dodavatel�)
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
			* kop�ruj �ablonu
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
		SELECT Src	&& posledn� p��kaz !!!

	PROCEDURE CloseTbls
		USE IN SELECT( 'Src' )
		USE IN SELECT( 'Trg' )
	
	PROCEDURE GetOutDir
		* c:\uctoRRRR s nejvy���m ��slem roku RRRR
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
