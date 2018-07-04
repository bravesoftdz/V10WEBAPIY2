unit trcommun;

interface

uses sysutils, Controls, db, dbctrls, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, DBCommon, Mul,
	 Bde, FichList, HDB, extCtrls, hctrls , trConstantes, Hmsgbox, forms, Hent1,
	 HPanel, UTOT, UTOB, ParamSoc ;
//--------------------
//Déclaration Publique
//--------------------
//Calcul de Numéro de transaction
Function GetNum             	(Module : string ; Societe : string; Transac : string) : string ;
Function SetNum             	(Module : string ; Societe : string; Transac : string; Num : string) : string ;
Function InitNumTransac     	(Module : string ; Societe : string ; Transaction : string) : string;
function TrGetNewNumPiece	(souche : string ; TypeSouche : string ; ecrit : boolean ; date : TDateTime ) : integer ;

//Calcul de taux
Function CalcNbJourParAnBase	(Date : TDateTime ; BaseCalcul : integer) : integer ;
Function CalcNbJourReelBase 	(DateDeb : TDate ; DateFin : TDate ; BaseCalcul : integer) : double ;
Function CalcMontantAgios   	(Montant : double; Taux : double; Base : integer ; NbJour : integer; tauxPreCompte : boolean ; AgiosPreCompte : Boolean) : double ;
Function CalcTauxAgios      	(MontantAgios : Double; Base : integer ; Montant : integer ;NbJour : integer; tauxPostCompte : boolean; AgiosPostCompte : Boolean) : double ;
Function CalcTauxPrePost    	(TauxDepartPost : boolean; ValTaux : double ; NbJour : integer ; Base : integer) : double ;

//Calcul fiche suivi
Function calcInteretDEB		(nb : integer ; taux : double; base : integer; date : TdateTime) : double ;

//Calcul avec accés BD
Function GetSoldePrec		(CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
Function GetSoldeValPrec	     (CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
Function GetSolde		     (CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
Function GetSoldeValeur		(CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
Function RetDeviseCompte	     (CompteGeneral : string) : string ;
Function GetCompteAgence	     (Compte : string) : string ;
Function RetDeviseBanque	     (Banque : string) : string ;
Function CalcDecimaleDevise	(Devise  : string) : integer ;
Function RetContreValeur	     (Date: TDateTime; Montant : double; DevMontant : string ; DevContrevalorise : string) : double ;
function RetPariteEuro		(Devise : string; Date : TDateTime ) : double ;
function GetCodeAFB		     (transac : string; TypeFlux : string) : string ;
function GetSensFlux		(transac : string; TypeFlux : string) : string ;
function GetRIB			(var Banque : string ; var Guichet : string ; var Num : string ; var cle : string ; CompteGeneral : string) : string ;
function GetConditionFinPlac  (var ConditionFinPlac : TRCONDITIONFINPLAC; Transac : string ; Banque : string ; Compte : string ; societe : string) : boolean ;
function GetConditionDec 	(Flux : string ;	Agence : string; tobC : TOB ): boolean ;
function GetValTauxRef		(TauxRef : string; Le : string; var Cotation : double) : boolean ;
function GetCodeFlux		(CodeTransac, TypeFlux: String): String;
function GetTypeFlux          (CodeCatTransac : string ): String;
Function calcCMVT		     (nombre : double ; TobCond : TOB) : double ;
function calcIBAN		     (pays : string ; RIB : string) : string ;
function calcRIB           (pays : string ; banque : string ; guichet : string ; compte : string; cle : string) : string ;
//function UpdateTrEcriture		(DateDepart: TDateTime; Compte: String): Boolean;

//Calculs de date de valeur / calendrier
function CalculPaques		(Year: Integer): TDateTime;
function RecupererChamp		(F : TFMul ; Champ : string) : string ;
function GetCalendrier		(CompteGeneral: String): String;
function CalcDateValeur		(CodeAFB, CompteGeneral: String; aDate: TDateTime): TDateTime;
function DateEstOuvre		(aDate: TDateTime; CodeCal: String): Boolean;
function GetPremierOuvre   (aDate: TDateTime; CodeCal: String): TDateTime;
function GetPrecedentOuvre	(aDate: TDateTime; CodeCal: String): TDateTime;
function DateEstOuvrable   (aDate: TDateTime): Boolean;
function GetPremierOuvrable   (aDate: TDateTime): TDateTime;
function GetPrecedentOuvrable (aDate: TDateTime): TDateTime;
FUNCTION DebutTrimestre		(Date : Tdate) : string ;
FUNCTION FinTrimestre		(Date : Tdate) : string ;

//Ecriture DB
Function EcritTREcriture   (TrEcrit : TREcritures ; Add : boolean) : boolean ;
FUNCTION RecalculSolde		(Compte : string; DateDepart : string; soldeInit : double; InitSolde : boolean) : double ;
Function TobSomme		      (TL : Tob ; NomChampRetour : string ; NomChamp : String ; Ope : string; ValeurChamp : variant; deb : integer) : double ;

//Divers
function TrShowMessage		(Caption: String; Num: Integer; P1, P2: String): TModalResult;
procedure ValidCodeOnKeyPress	(var Key: Char);
procedure FillBackground	(Num: Integer; Panel: THPanel);
Procedure AssignDrapeau		(Source : TImage ; CodeIso : string) ;
Function TrGetParamSoc		(champs:string) : variant ;

implementation


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : Rempli le fond d'un Panel avec la texture choisie dans les préférences
Suite ........ : Num = 0 : Fond de l'application ; Num = 1 : Connexion
Mots clefs ... : IMAGE;FOND
*****************************************************************}
procedure FillBackground(Num: Integer; Panel: THPanel);
begin
  if not V_PGI.ModeTSE then // ?
	LoadFond(Panel.Bitmap, Num, Panel);
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Rempli Source d'un drapeau correspondant au pays de la
Suite ........ : devise CodeIso
Mots clefs ... : DEVISE;ISO;DRAPEAU
*****************************************************************}
Procedure AssignDrapeau(Source : TImage; CodeIso : string) ;
Var
	SQL 	: String ;
	Q 		: TQuery ;
Begin
	{SQL := 'Select LO_CODE,LO_TYPE,LO_LIBELLE, LO_OBJET  From LIENSOLE where LO_TYPE="DRA" AND LO_CODE="' + CodeIso +'"';
    Q:= OpenSQL(SQL, true);
	if not Q.Eof then Source.Picture.Assign(Q.FieldByName('LO_OBJET'))
		else
	   Source.Picture.Assign(Nil);
	Ferme(Q);
   }
   //RRO en attente de résolution
   Source.Picture.Assign(Nil);
End ;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Affiche le message passé en paramètre, stocké dans
Suite ........ : trConstantes.pas
Mots clefs ... : MESSAGES
*****************************************************************}
function TrShowMessage(Caption: String; Num: Integer; P1, P2: String): TModalResult;
var	S: String;
begin
  S := IntToStr(Num)+';'+Caption+';'+TrMessage[Num]; // N°;Titre;Message
  Result := HShowMessage(S, P1, P2);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 05/12/2001
Modifié le ... :   /  /    
Description .. : Filtre les caractères saisis pour un code afin de respecter les
Suite ........ : règles d'ergonomie (règle N°7)
Suite ........ : Le caractère est changé en code 0 s'il n'est pas accepté
Mots clefs ... : CODE;TOUCHE;FILTRE
*****************************************************************}
procedure ValidCodeOnKeyPress(var Key: Char);
begin
  if (Key in ['0'..'9', 'A'..'Z', 'a'..'z']) or (Key < ' ') then
  begin	// Caractères de contrôle, lettres majuscules et chiffres seuls autorisés
	if Key >= 'a' then
		Key := Chr(Ord(Key) and $DF); // Passe en majuscule
  end
  else
	Key := #0;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 17/10/2001
Modifié le ... :   /  /
Description .. : Rend la date de Pâques pour une année donnée
Suite ........ : Non je ne l'ai pas inventé ; oui ça marche !
Mots clefs ... : DATE;PAQUES
*****************************************************************}
function CalculPaques(Year: Integer): TDateTime;
var c,n,k,i,j,l,m,d: integer;
begin
  c := Year div 100;
  n := Year mod 19;
  k := (c-17) div 25;
  i := (c-c div 4 - (c-k) div 3 + 19*n +15) mod 30;
  i := i-(i div 28) * (1-(i div 28)*(29 div (i+1)) * ((21-n) div 11));
  j := (Year+ Year div 4 + i + 2 - c + c div 4) mod 7;
  l := i - j;
  m := 3 + (l+40) div 44;
  d := l + 28 - 31 * (m div 4);
  Result := EncodeDate(Year, m, d);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Rend le code calendrier à partir du compte général, via
Suite ........ : l'agence
Mots clefs ... : CALENDRIER
*****************************************************************}
function GetCalendrier(CompteGeneral: String) : String; //CodeCalendrier
var Q: TQuery;
begin
  Result := '';
  Q := OpenSQL('Select TRA_CODECAL from AGENCE left join BANQUECP on TRA_GUICHET=BQ_GUICHET '+
				'where BQ_GENERAL="'+CompteGeneral+'"', True);
  if not Q.Eof then
	Result := Q.Fields[0].AsString;
  Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Rend la date de valeur à partir d'une date et des conditions
Suite ........ : de valeur
Mots clefs ... : VALEUR;GLISSEMENT
*****************************************************************}
function CalcDateValeur(CodeAFB, CompteGeneral: String; aDate: TDateTime): TDateTime;
var Q: TQuery;
	CodeCal: String;
	N: Integer;

	procedure ExecGlissement(TypeGlissement, NbJours : String);
	begin
	  N := Q.FieldByName(NbJours).AsInteger;
	  if N > 0 then
		case Q.FieldByName(TypeGlissement).AsInteger of
			0 : Result := Result + N; // Calendaire
			1 : while N > 0 do // Ouvré
				begin
					Result := GetPremierOuvre(Result+1, CodeCal);
					Dec(N);
				end;
			2 : while N > 0 do // Ouvrable
				begin
					Result := GetPremierOuvrable(Result+1);
					Dec(N);
				end;
		end
	  else if N < 0 then
		case Q.FieldByName(TypeGlissement).AsInteger of
			0 : Result := Result + N; // Calendaire
			1 : while N < 0 do // Ouvré
				begin
					Result := GetPrecedentOuvre(Result-1, CodeCal);
					Inc(N);
				end;
			2 : while N < 0 do // Ouvrable
				begin
					Result := GetPrecedentOuvrable(Result-1);
					Inc(N);
				end;
		end;
	end;

begin
  Result := aDate;
  Q := OpenSQL('Select TCV_TYPEGLISSEMENT,TCV_NBJPREMIER,TCV_TYPEPREMIER,TCV_NBJDEUXIEME,TCV_TYPEDEUXIEME '+
				'from CONDITIONVAL where TCV_CODECIB="'+CodeAFB+'" and TCV_COMPTEGENERAL="'+CompteGeneral+'"', True);
  if not Q.Eof then
  begin
	CodeCal := GetCalendrier(CompteGeneral);
	 // Glissement initial
	N := Q.FieldByName('TCV_TYPEGLISSEMENT').AsInteger;
	// 0 : Pas de report
	if N > 0 then
	begin
		Result := GetPremierOuvre(Result, CodeCal); // Premier ouvré
		if N = 2 then // Lendemain (du premier ouvré)
			Result := Result + 1;
	end;
	 // Premier glissement
	ExecGlissement('TCV_TYPEPREMIER', 'TCV_NBJPREMIER');
	 // Deuxieme glissement
	ExecGlissement('TCV_TYPEDEUXIEME', 'TCV_NBJDEUXIEME');
  end;
  // else !!
  Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Vérifie si une date/Calendrier ouvrée
Mots clefs ... : CALENDRIER;FERME;OUVRE
*****************************************************************}
function DateEstOuvre(aDate: TDateTime; CodeCal: String): Boolean;
var	Week, tmpDate: TDateTime;
	S: String;
	D, M, Y: Word;
begin
  Result := False;
  S := 'Select TJF_DATEFERMEE from JOURFERME where TJF_CODECAL="'+CodeCal+'" and '; // Calendrier concerné
  // Construit date "jour de la semaine"
  Week := CalendarWeekBound + (DayOfWeek(aDate)+5) mod 7; // Dimanche = 1 -> 6, Lundi = 2 -> 0 !
  // Rend l'année neutre
  DecodeDate(aDate, Y, M, D);
  tmpDate := EncodeDate(CalendarClosedYear, M, D);
  // Recherche parmi les dates fermées si existe, ou si jour de la semaine fermé
  if ExisteSQL(S+'(TJF_DATEFERMEE="'+UsDateTime(tmpDate)+'" or TJF_DATEFERMEE="'+UsDateTime(Week)+'")') then
	Exit;

  // Cherche parmi les jours fériés mobiles
  { Optimisation : voir si le décalage est reconnu
  I := Trunc(CalculPaques(Y) - aDate); // Décalage
  if I in [0, 1, 39, 49, 50] then	Voir tablette JFE !
  begin }
	tmpDate := CalendarEasterBound + aDate - CalculPaques(Y);
	if ExisteSQL(S+'TJF_DATEFERMEE="'+UsDateTime(tmpDate)+'"') then
		Exit;

  Result := True; // Ouvré
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Rend la prochaine date ouvrée
Mots clefs ... : CALENDRIER;OUVRE
*****************************************************************}
function GetPremierOuvre(aDate: TDateTime; CodeCal: String): TDateTime;
begin	// Pour optimiser, faire une version de DateEstFerme avec Tob
  Result := aDate;
  while not DateEstOuvre(Result, CodeCal) do
	Result := Result+1;
end;

function GetPrecedentOuvre(aDate: TDateTime; CodeCal: String): TDateTime;
begin	// Pour optimiser, faire une version de DateEstFerme avec Tob
  Result := aDate;
  while not DateEstOuvre(Result, CodeCal) do
	Result := Result-1;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Vérifie si une date est ouvrable
Suite ........ : Pas dimanche ni jour férié
Mots clefs ... : CALENDRIER;OUVRABLE
*****************************************************************}
function DateEstOuvrable(aDate: TDateTime): Boolean;
var	EasterDate, tmpDate: TDateTime;
	Q: TQuery;
	S: String;
	I: Integer;
	D, M, Y: Word;
begin
  Result := False;
  if DayOfWeek(aDate) = 1 then // Dimanche = 1
	Exit;

  // Cherche parmi les jours fériés (voir TOM_CALENDRIER.LoadJourFerie)
  DecodeDate(aDate, Y, M, D);
  EasterDate := CalculPaques(Y);
  tmpDate := EncodeDate(CalendarClosedYear, M, D); // Rend l'année neutre
  Q := OpenSQL('Select CO_LIBRE from COMMUN where CO_TYPE="JFE"', True);
  try
	while not Q.Eof do
	begin
		S := Q.Fields[0].AsString;
		if S[1] = 'P' then	// Date à calculer par rapport à Pâques
		begin	// Extrait le décalage (<200)
			if S = 'P' then
				I := 0
			else	// P+
				I := StrToInt(Copy(S, 3, 9));
			if aDate = (EasterDate+I) then
				Exit;
		end
		else	// Date fixe
			if tmpDate = StrToDate(S) then
				Exit;

		Q.Next;
	end;
  finally
	Ferme(Q);
  end;

  Result := True; // Ouvrable
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Rend la prochaine date ouvrable
Mots clefs ... : CALENDRIER;OUVRABLE
*****************************************************************}
function GetPremierOuvrable(aDate: TDateTime): TDateTime;
begin	// Pour optimiser, faire une version avec Tob
  Result := aDate;
  while not DateEstOuvrable(Result) do
	Result := Result+1;
end;

function GetPrecedentOuvrable(aDate: TDateTime): TDateTime;
begin	// Pour optimiser, faire une version avec Tob
  Result := aDate;
  while not DateEstOuvrable(Result) do
	Result := Result-1;
end;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Calcul le montant des agios/intérets par rapport au type de
Suite ........ : taux, et du type d'agios/intérets
Mots clefs ... : MONTANT;POSTCOMPTE; PRECOMPTE;AGIOS;INTERET
*****************************************************************}
Function CalcMontantAgios(Montant : double; Taux : double;
						  Base    : integer ; NbJour : integer;
						  tauxPreCompte : boolean ;
						  AgiosPreCompte : Boolean) : double ;
var vResult : double ;
begin
  vResult := 0 ;
  Taux := Taux /100 ;
  try
	// Si Agios PostCompté TauxPostCompte
	if Not AgiosPreCompte then vResult := Montant * Taux / Base * NbJour
	// Cas Agios PréCompte TauxPostCompte
	else vResult := Montant/(1+(base/(taux*nbjour)));
  finally
	Result := vResult;
  end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Calcul le Taux des Agios/intérêt par rapport au type de
Suite ........ : taux, et du type d'agios/intérets
Mots clefs ... : MONTANT;POSTCOMPTE; PRECOMPTE;AGIOS;INTERET
*****************************************************************}
Function CalcTauxAgios(MontantAgios : Double; Base : integer ; Montant : integer ;NbJour : integer; tauxPostCompte : boolean; AgiosPostCompte : Boolean) : double ;
Var vResult : Double ;
begin
	vResult := 0;
	try
	if AgiosPostCompte then
	  // Cas Agios PostCompte TauxPostCompte
		// -----------------------------------
	  if tauxPostCompte then
		vResult := (montantAgios* base)/(Montant*NbJour)
	  else
	  // Cas Agios PostCompte TauxPréCompte
	  // -----------------------------------
		vResult := (montantAgios* base)/(NbJour*(Montant+MontantAgios))
	else
	  // Cas Agios PréCompte TauxPréCompte
	  // -----------------------------------
	  if Not tauxPostCompte then
        vResult := (MontantAgios*Base)/(Montant*NbJour)
      else
      // Cas Agios PréCompte TauxPostCompte
      // -----------------------------------
        vResult := (montantAgios* base)/(NbJour*(Montant-MontantAgios))
  finally
  	Result := vResult *100;
  end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... : 01/10/2001
Description .. : Calcul le taux pré ou post compté d'un taux avec comme
Suite ........ : crité si c'est un taux postcompé
Mots clefs ... : TAUX PRECOMPTE; TAUX POSTCOMPTE
*****************************************************************}
Function CalcTauxPrePost (TauxDepartPost : boolean; ValTaux : double ;
							NbJour : integer ; Base : integer) : double ;
Begin
	Try
	valTaux := valTaux / 100;
	if TauxDepartPost then
		//Si c'est un taux postcompté on retourne un taux Précompté
		Result := (ValTaux)/(1+(ValTaux  * NbJour / Base)) * 100
	else
		//Si c'est un taux Précompté on retourne un taux Postcompté
		Result := (ValTaux) /(1-(ValTaux * NbJour / Base)) * 100;
	except
	Result := 0 ;
	end ;
end ;


Function calcInteretDEB(nb : integer ; taux : double; base : integer; date : TdateTime) : double ;
var num : double ;
Begin
		num := CalcNbJourParAnBase(date,base)*100 ;
		if num<>0 then result := (nb*taux)/num;
end ;

//Function calcCMVT(nombre : double ; taux : double ) : double ;
Function calcCMVT(nombre : double ; TobCond : TOB) : double ;
var taux : Double ;
begin
		{GetValTauxRef(GetField('TCT_TAUX'), GetField('TCT_DATECREATION'), Taux );
		Maj := getField('TCT_MAJORATION') ;
		Mul := getField('TCT_MULTIPLE') ;
		taux := (Taux+Maj)* (Mul/100) ;
		}
		if nombre <= TobCond.GetValue('TCN_PLAFOND1') then
			GetValTauxRef( TobCond.GetValue('TCN_TAUXREF1'), DateToStr(date), taux)
		else if nombre <= TobCond.GetValue('TCN_PLAFOND2') then
			GetValTauxRef( TobCond.GetValue('TCN_TAUXREF2'), DateToStr(date), taux)
			else if nombre > TobCond.GetValue('TCN_PLAFOND2') then
				GetValTauxRef( TobCond.GetValue('TCN_TAUXREF3'), DateToStr(date), taux);

		result := nombre * (taux /100);
end ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 31/07/2001
Modifié le ... : 19/09/2001
Description .. : Fonction qui génère un N° Transactionnel
Mots clefs ... : GETNUM; NUMERO;TRANSACTION
*****************************************************************}
function GetNum(Module : string; Societe : string; Transac : string) : string ;
Var  Q : TQuery ;
	 SQL : string ;
   Num : integer ;
BEGIN
	//Acces BD
	SQL := 'SELECT TCP_COMPTEURTRESO, TCP_MODULECOMPT,' +
		   'TCP_NUMCOMPT,TCP_SOCIETECOMPT,TCP_TRANSACCOMPT ' +
		   'FROM COMPTEURTRESO WHERE ' ;
	SQL := SQL + 'TCP_MODULECOMPT="'+ Module + '"';
	SQL := SQL + ' AND TCP_SOCIETECOMPT="' + Societe + '"' ;
	SQL := SQL + ' AND TCP_TRANSACCOMPT="' + Transac + '"' ;
	Q:=OPENSQL(SQL , TRUE) ; //Lecture Seule
	Num := 0 ;
	if not Q.Eof then
	BEGIN
	  Num := Q.Findfield('TCP_NUMCOMPT').Asinteger ;
	END;
	If Num > 0 Then
	BEGIN
	  inc(num) ;
	  result := StringOfChar('0', 4-length(inttostr(num))) + inttostr(num)  ;
	END ELSE result := '0001' ;
	Ferme(Q) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : retourne un numéro de transaction tréso
Mots clefs ... : TRANSACTION;NUMERO;TRESORERIE
*****************************************************************}
function SetNum(Module : string; Societe : string; Transac : string; Num : string) : string ;
Var SQL : string ;
BEGIN
    SQL := 'UPDATE COMPTEURTRESO SET TCP_NUMCOMPT="' + Num ;
    SQL := SQL + '" WHERE' ;
    SQL := SQL + ' TCP_MODULECOMPT="'+ Module + '"';
    SQL := SQL + ' AND TCP_SOCIETECOMPT="' + Societe + '"' ;
    SQL := SQL + ' AND TCP_TRANSACCOMPT="' + Transac + '"' ;
    if ExecuteSQL(SQL)<=0 then
		begin
			 SQL := 'INSERT COMPTEURTRESO (TCP_NUMCOMPT,TCP_MODULECOMPT,TCP_SOCIETECOMPT,TCP_TRANSACCOMPT) ' ;
			 SQL := SQL + ' VALUES (' ;
			 SQL := SQL + '"' + Num + '",' ;
			 SQL := SQL + '"' + Module + '",';
			 SQL := SQL + '"' + Societe + '",' ;
			 SQL := SQL + '"' + Transac + '")' ;
			 ExecuteSQL(SQL);
		end ;
END ;

function TrGetNewNumPiece(souche : string ; TypeSouche : string ; ecrit : boolean ; date : TDateTime ) : integer ;
Var Q : TQuery ;
    Num : Longint ;
    SQL : String ;
BEGIN
Num:=0 ;
SQL := 'Select SH_NUMDEPART from SOUCHE Where SH_TYPE="' + TypeSouche + '" AND SH_SOUCHE="'+Souche+'"' ;
Q:=OpenSQL(SQL,True) ;
if Not Q.EOF then Num:=Q.Fields[0].AsInteger else V_PGI.IoError:=oeUnknown ;
Ferme(Q) ;
if ((Num>0) and (Ecrit)) then
   BEGIN
   SQL:='UPDATE SOUCHE SET SH_NUMDEPART='+IntToStr(Num+1)+' Where SH_TYPE="' + TypeSouche + '" AND SH_SOUCHE="'+Souche+'" AND SH_NUMDEPART='+IntToStr(Num) ;
   if ExecuteSQL(SQL)<=0 then BEGIN Num:=0 ; V_PGI.IoError:=oeSaisie ; END ;
	 END ;
Result:=Num ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Créer le numéro de transaction initial
Mots clefs ... : TRANSACTION; NUMERO TRANSACTION
*****************************************************************}
Function InitNumTransac(Module : string ; Societe : string ; Transaction : string) : string;
var str, Num : string ;
BEGIN
   Result := '' ;
   if length(Module)<>0 then //and length(Societe)<>0 and length(Transaction)<>0 then
   begin
	 Num := GetNum(Module,Societe,Transaction) ;
	 str := Module+Societe+Transaction+num ;
   end ;
   Result:= Str ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 07/09/2001
Modifié le ... : 01/10/2001
Description .. : Fonction de retour du nombre de jour d'une  année par
Suite ........ : rapport à la base de calcul
Mots clefs ... : BASE DE CALCUL; NOMBRE JOUR; ANNEE
*****************************************************************}
Function CalcNbJourParAnBase(Date : TDateTime ; BaseCalcul : integer) : integer ;
var
   base : integer ;
   Annee,mm,jj : word ;
BEGIN
	base := 360 ;
	try
	  DecodeDate(Date, Annee, mm, jj);
	  base := 360 ;
	  case basecalcul of
		 B30_360 : base := 360 ;
		 B30_365 : base := 365 ;
		 BJUSTE_360 : base := 360 ;
		 BJUSTE_365 : base := 365 ;
		 BJUSTE_JUSTE :
		 BEGIN
			base := 365 ;
			//Vérifie si c'est une année bisextile
			if IsLeapYear(Annee) then base := 366 ;
		 END ;
		 B30E_360 : base := 360 ;
	  end ; //Case
   finally
	result := base ;
   end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 07/09/2001
Modifié le ... : 19/09/2001
Description .. : Retourne le nombre de jour entre 2 dates par rapport à une
Suite ........ : base de calcul
Mots clefs ... : NOMBRE DE JOUR; BASE DE CALCUL
*****************************************************************}
Function CalcNbJourReelBase(DateDeb : TDate ; DateFin : TDate ; BaseCalcul : integer) : Double ;
Var jDeb, mDeb, aDeb, jFin, mFin,aFin : word ;
   Nbj : double ;
BEGIN
   Nbj := 0 ;
   try
	  if (DateDeb=DateFin) or (basecalcul<0) then Result := 0
	  ELSE                                  
		BEGIN
		 case basecalcul  of
			B30_360,B30_365,B30E_360 :
			   BEGIN
				  DecodeDate(DateDeb, aDeb, mDeb, jDeb);
				  DecodeDate(DateFin, aFin, mfin, jFin);
				  if jDeb>=30 then jDeb:=30;
				  if jFin>=30 then jFin:=30;
				  Nbj:=((aFin-aDeb)*12+(mFin-mDeb))*30+(jFin-jDeb) ;
				  //Cas particulier
				  if basecalcul=B30E_360 then
				  BEGIN
					 if mDeb=2 then
					 BEGIN
						case jDeb of
						   28 : Nbj:=Nbj-3 ;
						   29 : Nbj:=Nbj-2 ;
						end //Case
					 END ; //if
					 if mFin=2 then
					 BEGIN
						case jfin of
						   28 : Nbj:=Nbj+2 ;
						   29 : Nbj:=Nbj+1 ;
						end //Case
					 END ; //if
				  END ;//if
			   END; //Case

			BJUSTE_360,BJUSTE_365, BJUSTE_JUSTE : Nbj:=DateFin-DateDeb;
		 END; //Case
	  END ; //if DateDeb=DateFin
   finally;
	Result:= Nbj ;
	 end ;
END;

FUNCTION DebutTrimestre(Date : Tdate) : string ;
VAR AA,MM,JJ : word ;
begin
	decodedate(date, AA,MM,JJ);
	case MM of
		1..3 : result:=datetostr(debutdemois(encodedate(AA,1,1)));
		4..6 : result:=datetostr(debutdemois(encodedate(AA,4,1)));
		7..9 : result:=datetostr(debutdemois(encodedate(AA,7,1)));
		10..12 : result:=datetostr(debutdemois(encodedate(AA,9,1)));
	end ;
end ;

FUNCTION FinTrimestre(Date : Tdate) : string ;
VAR AA,MM,JJ : word ;
begin
	decodedate(date, AA,MM,JJ);
	case MM of
		1..3 : result:=datetostr(Findemois(encodedate(AA,3,JJ)));
		4..6 : result:=datetostr(Findemois(encodedate(AA,6,JJ)));
		7..9 : result:=datetostr(Findemois(encodedate(AA,9,JJ)));
		10..12 : result:=datetostr(Findemois(encodedate(AA,12,JJ)));
	end ;
end ;


//Calcul Base de donnée
FUNCTION RecupererChamp(F : TFMul ; Champ : string) : string ;
BEGIN
	If F.Q <> Nil then
	BEGIN
		try
			if (Length(F.Q.Findfield(Champ).AsString)>0) then
				Result:=F.Q.FindField(Champ).AsString
			else Result := '';
		except ;
			Result := '';
		end ;
	END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /    
Description .. : met a jour la table TREcriture
Mots clefs ... :
*****************************************************************}
FUNCTION EcritTREcriture(TrEcrit : TREcritures ; Add : boolean) : boolean ;
Var SQL : string ;
	phrase : string ;
	phrase2 : string ;
	i : integer ;
	Q : TQuery ;
	solde : double ;
	//soldeValeur : double ;
	//mnt : string ;
	AncienSoldeValeur : double ;
Begin
		SQL:= 'SELECT * FROM TRECRITURE WHERE TE_GENERAL="' + TrEcrit.TE_GENERAL +
				'" AND TE_NUMEROPIECE='+ inttostr(TrEcrit.TE_NUMEROPIECE) + ' and TE_NUMTRANSAC="' + TrEcrit.TE_NUMTRANSAC + '"' ;
		Q:=OpenSQL(SQL, True);
		if not Q.Eof then begin
		AncienSoldeValeur := Q.FindField('TE_SOLDEDEVVALEUR').AsFloat;
		//UPDATE
	   { TrEcrit.TE_JOURNAL			:= Q.FindField('TE_JOURNAL').AsString ;
		TrEcrit.TE_NUMEROPIECE		:= Q.FindField('TE_NUMEROPIECE').AsString ;
		TrEcrit.TE_NUMLIGNE 	  	:= Q.FindField('TE_NUMLIGNE').AsString ;
		TrEcrit.TE_EXERCICE 	  	:= Q.FindField('TE_EXERCICE').AsString ;
		TrEcrit.TE_NUMECHE 	  		:= Q.FindField('TE_NUMECHE').AsInteger ;
		TrEcrit.TE_GENERAL 	  		:= Q.FindField('TE_GENERAL').AsString ;
		TrEcrit.TE_CODECIB 	  		:= Q.FindField('TE_CODECIB').AsString ;
		TrEcrit.TE_CODEFLUX 	  	:= Q.FindField('TE_CODEFLUX').AsString ;
		TrEcrit.TE_REFINTERNE 	  	:= Q.FindField('TE_REFINTERNE').AsString ;
		TrEcrit.TE_SOCIETE 	  		:= Q.FindField('TE_SOCIETE').AsString ;
		TrEcrit.TE_QUALIFORIGINE	:= Q.FindField('TE_QUALIFORIGINE').AsString ;
		TrEcrit.TE_NUMTRANSAC		:= Q.FindField('TE_NUMTRANSAC').AsString ;
		TrEcrit.TE_LIBELLE			:= Q.FindField('TE_LIBELLE').AsString ;
		TrEcrit.TE_USERCREATION		:= Q.FindField('TE_USERCREATION').AsString ;
		TrEcrit.TE_USERMODIF		:= Q.FindField('TE_USERMODIF').AsString ;
		TrEcrit.TE_USERVALID		:= Q.FindField('TE_USERVALID').AsString ;
		TrEcrit.TE_USERCOMPTABLE	:= Q.FindField('TE_USERCOMPTABLE').AsString ;
		TrEcrit.TE_DATECREATION		:= Q.FindField('TE_DATECREATION').AsString ;
		TrEcrit.TE_DATEMODIF		:= Q.FindField('TE_DATEMODIF').AsString ;
		TrEcrit.TE_DATEVALID		:= Q.FindField('TE_DATEVALID').AsString ;
		TrEcrit.TE_DATECOMPTABLE	:= Q.FindField('TE_DATECOMPTABLE').AsString ;
		TrEcrit.TE_DATERAPPRO		:= Q.FindField('TE_DATERAPPRO').AsString ;
		TrEcrit.TE_DEVISE			:= Q.FindField('TE_DEVISE').AsString ;
		TrEcrit.TE_MONTANTDEV		:= Q.FindField('TE_MONTANTDEV').AsFloat ;
		TrEcrit.TE_MONTANT			:= Q.FindField('TE_MONTANT').AsFloat ;
		TrEcrit.TE_COTATION			:= Q.FindField('TE_COTATION').AsFloat ;
		TrEcrit.TE_DATEVALEUR		:= Q.FindField('TE_DATEVALEUR').AsString ;
		TrEcrit.TE_SOLDEDEVVALEUR	:= Q.FindField('TE_SOLDEDEVVALEUR').AsFloat ;
		TrEcrit.TE_SOLDEDEV 	  	:= Q.FindField('TE_SOLDEDEV').AsFloat ;
		TrEcrit.TE_CODEBANQUE 	  	:= Q.FindField('TE_CODEBANQUE').AsString ;
		TrEcrit.TE_CODEGUICHET   	:= Q.FindField('TE_CODEGUICHET').AsString ;
		TrEcrit.TE_NUMCOMPTE 	  	:= Q.FindField('TE_NUMCOMPTE').AsString ;
		TrEcrit.TE_CLERIB 		  	:= Q.FindField('TE_CLERIB').AsString ;
		}
		Phrase := 'UPDATE TRECRITURE SET ' ;
		Phrase := Phrase + Q.Fields[0].FieldName + '="'  + TrEcrit.TE_JOURNAL         + '", ' ;
		Phrase := Phrase + Q.Fields[1].FieldName + '=' 	 + strFPoint(TrEcrit.TE_NUMEROPIECE)     + ', ' ;
		Phrase := Phrase + Q.Fields[2].FieldName + '="'  + strFPoint(TrEcrit.TE_NUMLIGNE)  + '", ' ;
		Phrase := Phrase + Q.Fields[3].FieldName + '="'  + TrEcrit.TE_EXERCICE        + '", ' ;
		Phrase := Phrase + Q.Fields[4].FieldName + '=' 	 + strFPoint(TrEcrit.TE_NUMECHE) + ', ' ;
		Phrase := Phrase + Q.Fields[5].FieldName + '="'  + TrEcrit.TE_GENERAL         + '", ' ;
		Phrase := Phrase + Q.Fields[6].FieldName + '="'  + TrEcrit.TE_CODECIB         + '", ' ;
		Phrase := Phrase + Q.Fields[7].FieldName + '="'  + TrEcrit.TE_CODEFLUX        + '", ' ;
		Phrase := Phrase + Q.Fields[8].FieldName + '="'  + TrEcrit.TE_REFINTERNE      + '", ' ;
		Phrase := Phrase + Q.Fields[9].FieldName + '="'  + TrEcrit.TE_SOCIETE         + '", ' ;
		Phrase := Phrase + Q.Fields[10].FieldName + '="' + TrEcrit.TE_QUALIFORIGINE   + '", ' ;
		Phrase := Phrase + Q.Fields[11].FieldName + '="' + TrEcrit.TE_NUMTRANSAC      + '", ' ;
		Phrase := Phrase + Q.Fields[12].FieldName + '="' + TrEcrit.TE_LIBELLE         + '", ' ;
		Phrase := Phrase + Q.Fields[13].FieldName + '="' + TrEcrit.TE_USERCREATION    + '", ' ;
		Phrase := Phrase + Q.Fields[14].FieldName + '="' + TrEcrit.TE_USERMODIF       + '", ' ;
		Phrase := Phrase + Q.Fields[15].FieldName + '="' + TrEcrit.TE_USERVALID       + '", ' ;
		Phrase := Phrase + Q.Fields[16].FieldName + '="' + TrEcrit.TE_USERCOMPTABLE   + '", ' ;
		Phrase := Phrase + Q.Fields[17].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATECREATION))   + '", ' ;
		Phrase := Phrase + Q.Fields[18].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATEMODIF))      + '", ' ;
		Phrase := Phrase + Q.Fields[19].FieldName + '="' + USDATETIME(strTodatetime(TrEcrit.TE_DATEVALID))      + '", ' ;
		Phrase := Phrase + Q.Fields[20].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATERAPPRO))     + '", ' ;
		Phrase := Phrase + Q.Fields[21].FieldName + '="' + TrEcrit.TE_DEVISE          + '", ' ;
		Phrase := Phrase + Q.Fields[22].FieldName + '='  + strFPoint(TrEcrit.TE_MONTANTDEV)      + ', ' ;
		Phrase := Phrase + Q.Fields[23].FieldName + '='  + strFPoint(TrEcrit.TE_MONTANT )        + ', ' ;
		Phrase := Phrase + Q.Fields[24].FieldName + '='	 + strFPoint(TrEcrit.TE_COTATION)        + ', ' ;
		Phrase := Phrase + Q.Fields[25].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATECOMPTABLE))  + '", ' ;
		Solde  := 0;
		Phrase := Phrase + Q.Fields[26].FieldName + '='  + strFPoint(solde)  + ', ' ;
		Phrase := Phrase + Q.Fields[27].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATEVALEUR))     + '", ' ;
		Phrase := Phrase + Q.Fields[2].FieldName + '='   + strFPoint(TrEcrit.TE_NUMLIGNE) + ', ' ;
		Phrase := Phrase + Q.Fields[3].FieldName + '="'  + TrEcrit.TE_EXERCICE        + '", ' ;
		Phrase := Phrase + Q.Fields[4].FieldName + '=' 	 + strFPoint(TrEcrit.TE_NUMECHE) + ', ' ;
		Phrase := Phrase + Q.Fields[5].FieldName + '="'  + TrEcrit.TE_GENERAL         + '", ' ;
		Phrase := Phrase + Q.Fields[6].FieldName + '="'  + TrEcrit.TE_CODECIB         + '", ' ;
		Phrase := Phrase + Q.Fields[7].FieldName + '="'  + TrEcrit.TE_CODEFLUX        + '", ' ;
		Phrase := Phrase + Q.Fields[8].FieldName + '="'  + TrEcrit.TE_REFINTERNE      + '", ' ;
		Phrase := Phrase + Q.Fields[9].FieldName + '="'  + TrEcrit.TE_SOCIETE         + '", ' ;
		Phrase := Phrase + Q.Fields[10].FieldName + '="' + TrEcrit.TE_QUALIFORIGINE   + '", ' ;
		Phrase := Phrase + Q.Fields[11].FieldName + '="' + TrEcrit.TE_NUMTRANSAC      + '", ' ;
		Phrase := Phrase + Q.Fields[12].FieldName + '="' + TrEcrit.TE_LIBELLE         + '", ' ;
		Phrase := Phrase + Q.Fields[13].FieldName + '="' + TrEcrit.TE_USERCREATION    + '", ' ;
		Phrase := Phrase + Q.Fields[14].FieldName + '="' + TrEcrit.TE_USERMODIF       + '", ' ;
		Phrase := Phrase + Q.Fields[15].FieldName + '="' + TrEcrit.TE_USERVALID       + '", ' ;
		Phrase := Phrase + Q.Fields[16].FieldName + '="' + TrEcrit.TE_USERCOMPTABLE   + '", ' ;
		Phrase := Phrase + Q.Fields[17].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATECREATION))   + '", ' ;
		Phrase := Phrase + Q.Fields[18].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATEMODIF))      + '", ' ;
		Phrase := Phrase + Q.Fields[19].FieldName + '="' + USDATETIME(strTodatetime(TrEcrit.TE_DATEVALID))      + '", ' ;
		Phrase := Phrase + Q.Fields[20].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATERAPPRO))     + '", ' ;
		Phrase := Phrase + Q.Fields[21].FieldName + '="' + TrEcrit.TE_DEVISE          + '", ' ;
		Phrase := Phrase + Q.Fields[22].FieldName + '='  + strFPoint(TrEcrit.TE_MONTANTDEV)      + ', ' ;
		Phrase := Phrase + Q.Fields[23].FieldName + '='  + strFPoint(TrEcrit.TE_MONTANT )        + ', ' ;
		Phrase := Phrase + Q.Fields[24].FieldName + '='	 + strFPoint(TrEcrit.TE_COTATION)        + ', ' ;
		Phrase := Phrase + Q.Fields[25].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATECOMPTABLE))  + '", ' ;
		Solde  := 0 ;
		Phrase := Phrase + Q.Fields[26].FieldName + '='  + strFPoint(solde)  + ', ' ;
		Phrase := Phrase + Q.Fields[27].FieldName + '="' + USDATETIME(strTodate(TrEcrit.TE_DATEVALEUR))     + '", ' ;
		//et on recalcule le nouveau
		Solde  := 0;
		Phrase := Phrase + Q.Fields[28].FieldName + '='  + strFPoint(solde)  + ', ' ;
		Phrase := Phrase + Q.Fields[29].FieldName + '="' + TrEcrit.TE_CODEBANQUE      + '", ' ;
		Phrase := Phrase + Q.Fields[30].FieldName + '="' + TrEcrit.TE_CODEGUICHET     + '", ' ;
		Phrase := Phrase + Q.Fields[31].FieldName + '="' + TrEcrit.TE_NUMCOMPTE       + '", ' ;
		Phrase := Phrase + Q.Fields[32].FieldName + '="' + TrEcrit.TE_CLERIB + '"' ;

		//WHERE
		Phrase := Phrase + ' WHERE TE_NUMTRANSAC="'      + TrEcrit.TE_NUMTRANSAC + '"' ;
		Phrase := Phrase + ' AND TE_NUMLIGNE='           + strFPoint(TrEcrit.TE_NUMLIGNE) ;
		ExecuteSQL(phrase);

	end
	else
	begin
		//INSERT
		Phrase:= 'INSERT TRECRITURE (' ;
		for i:=0 to 33 do
        	Phrase := Phrase + Q.Fields[i].FieldName + ',' ;
		Phrase := Phrase + Q.Fields[i].FieldName ;
		//VALUES
      Phrase2 := '"' + TrEcrit.TE_NUMTRANSAC + '", ' ;
		Phrase2 := Phrase2 + strFPoint(TrEcrit.TE_NUMLIGNE)   + ', ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_JOURNAL        + '", ' ;
		Phrase2 := Phrase2 +   strFPoint(TrEcrit.TE_NUMEROPIECE) + ', ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_EXERCICE        + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_GENERAL         + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrECrit.TE_CONTREPARTIETR  + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_CODECIB         + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_CODEFLUX        + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_REFINTERNE      + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_SOCIETE         + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_QUALIFORIGINE   + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_LIBELLE         + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_USERCREATION    + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_USERMODIF       + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_USERVALID       + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_USERCOMPTABLE   + '", ' ;
		Phrase2 := Phrase2 + '"' +  USDATETIME(strTodate(TrEcrit.TE_DATECREATION))    + '", ' ;
		Phrase2 := Phrase2 + '"' +  USDATETIME(strTodate(TrEcrit.TE_DATEMODIF))       + '", ' ;
		Phrase2 := Phrase2 + '"' +  USDATETIME(strTodatetime(TrEcrit.TE_DATEVALID))   + '", ' ;
		Phrase2 := Phrase2 + '"' +  USDATETIME(strTodate(TrEcrit.TE_DATERAPPRO))      + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_DEVISE          + '", ' ;
		Phrase2 := Phrase2 + strFPoint(TrEcrit.TE_MONTANTDEV)  + ', ' ;
		Phrase2 := Phrase2 + strFPoint(TrEcrit.TE_MONTANT )    + ', ' ;
		Phrase2 := Phrase2 + strFPoint(TrEcrit.TE_COTATION)    + ', ' ;
		Phrase2 := Phrase2 + '"' +  USDATETIME(strTodate(TrEcrit.TE_DATECOMPTABLE))   + '", ' ;
		Phrase2 := Phrase2 + strFPoint(TrEcrit.TE_SOLDEDEV) + ', ';
		Phrase2 := Phrase2 + '"' +  USDATETIME(strTodate(TrEcrit.TE_DATEVALEUR))      + '", ' ;
		Phrase2 := Phrase2 + strFPoint(TrEcrit.TE_SOLDEDEVVALEUR) + ', ';
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_CODEBANQUE      + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_CODEGUICHET     + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_NUMCOMPTE       + '", ' ;
		Phrase2 := Phrase2 + '"' +  TrEcrit.TE_CLERIB + '", ' ;
      Phrase2 := Phrase2 + '"' +  TrEcrit.TE_IBAN + '", ' ;
      Phrase2 := Phrase2 + '"' +  TrEcrit.TE_NATURE + '" ' ;
		phrase := Phrase + ') VALUES (' + phrase2 + ')' ;
		Ferme(Q);
		ExecuteSQL(phrase);


	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 08/01/2002
Modifié le ... :   /  /
Description .. : Function sui recalcul les solde à partir d'une date
Mots clefs ... : SOLDE;DATE
*****************************************************************}
FUNCTION RecalculSolde(Compte : string; DateDepart : string; soldeInit : double; InitSolde : boolean) : double ;
Var //phrase : string ;
	SoldePrec : double ;
	//signe : string ;
	I: Integer;
	OK: Boolean;
    QQ: TQuery ;
		soldeMAJ : double ;
    montant : double ;
		dateretour : string ;
		tobTRE, TL : TOB ;
		SQL : String ;
BEGIN
	//MISE A JOUR DES SOLDES SUIVANT
	if Compte <> '' then
	begin
		TobTRE := Tob.Create('_TRE', Nil, -1);
		SQL := 'SELECT ' + 
		'TE_NUMTRANSAC, TE_NUMLIGNE, TE_DATECOMPTABLE,TE_GENERAL,TE_MONTANTDEV, ' +
		'TE_NUMEROPIECE,TE_SOLDEDEV,TE_SOLDEDEVVALEUR ' +
		'FROM TRECRITURE ' +
		'WHERE TE_GENERAL="'+ Compte +
		'" AND TE_DATECOMPTABLE>="' + UsDateTime(strtodate(DateDepart)) + '"' ;
		QQ:=OpenSQL( SQL, true);
		TobTRE.LoadDetailDB('TRECRITURE', '', '', QQ, false);
		ferme(QQ);

		if TobTRE.Detail.Count>0 then
		begin
				if InitSolde then
               SoldePrec := soldeInit
				else
               SoldePrec := GetSoldePrec(compte,datedepart,dateretour);
				//solde en date d'ope
				TobTRE.detail.Sort('TE_DATECOMPTABLE;TE_NUMEROPIECE');
				for i:=0 to TobTRE.Detail.Count-1 do
				begin
				   montant:=TobTRE.Detail[i].GetValue('TE_MONTANTDEV') ;
               dateretour:=TobTRE.Detail[i].GetValue('TE_DATECOMPTABLE') ;
               if i=0 then
                  //on calcul avec la valeur de la base ou celle demandé
                  if InitSolde then
                     soldeMAJ := SoldePrec
                  else	soldeMAJ := montant + soldeprec
               else
                     //on utilise la valeur précédente
                     soldeMAJ:=montant+SoldeMAJ;

               tobTRE.Detail[i].PutValue('TE_SOLDEDEV', soldeMAJ);
				end ;
				//Mise à jour des soldes en date d'opé
				tobTRE.UpdateDB(false, false) ;
				tobTRE.Free ;
		end ; // Detail

		TobTRE := Tob.Create('_TRE', Nil, -1);
		SQL := 'SELECT ' +
		'TE_NUMTRANSAC, TE_NUMLIGNE, TE_DATECOMPTABLE, TE_DATEVALEUR,TE_GENERAL,TE_MONTANTDEV, ' +
		'TE_NUMEROPIECE,TE_SOLDEDEV,TE_SOLDEDEVVALEUR ' +
		'FROM TRECRITURE ' +
		'WHERE TE_GENERAl="'+ Compte +
		'" AND TE_DATEVALEUR>="' + UsDateTime(strtodate(DateDepart)) + '"' ;
		QQ:=OpenSQL( SQL, true);
		TobTRE.LoadDetailDB('TRECRITURE', '', '', QQ, false);
		ferme(QQ) ;

		if TobTRE.Detail.Count>0 then
		begin
			//Solde en date de valeur
			TobTRE.detail.Sort('TE_DATEVALEUR;TE_NUMEROPIECE');
			TL := Tob.Create('_TL', nil,-1);
			TL := TobTRE ;                                   
			//on récupère le dernier solde en valeur connu
			soldeMAJ := GetSoldeValPrec(compte,datedepart,dateretour);
			for i:=0 to TobTRE.Detail.Count-1 do
			begin
				//on calcul avec la valeur de la base ou celle demandé
				dateretour := tobTRE.Detail[i].GetValue('TE_DATECOMPTABLE');
				if InitSolde AND (i=0) then
					SoldeMAJ := TobTRE.Detail[i].GetValue('TE_SOLDEDEV')
				else
					soldeMAJ:= TobSomme(TL, 'TE_MONTANTDEV', 'TE_DATEVALEUR', '<=', dateretour, i);

				tobTRE.Detail[i].PutValue('TE_SOLDEDEVVALEUR', soldeMAJ);
			end ;
			tobTRE.UpdateDB(false, false) ;
			tobTRE.Free;
		end ;
		Result := soldeMAJ ;
	end; //compte
END;

Function TobSomme(TL : Tob ; NomChampRetour : string ; NomChamp : String ; Ope : string; ValeurChamp : variant; deb : integer) : double ;
Var i,j  : integer ;
		Val  : Double  ;
		Comp : variant ;
		dateL : string ;
Begin
	result := 0 ;
	for i:=0 to TL.Detail.count-1 do
	begin
		j:=TL.Detail[i].GetNumChamp(NomChamp) ;
		if varType(TL.Detail[i].Valeurs[j]) in [vardate] then
		begin
			comp := VarAsType(TL.Detail[i].Valeurs[j],vardate) ;
			if Ope = '<=' then
			begin
				//On récupère la valeur deu champ retour dans le bon type
				j:=TL.Detail[i].GetNumChamp(NomChampRetour) ;
				if varType(TL.Detail[i].Valeurs[j]) in [varByte,varSmallint,varInteger,varSingle,varDouble,varCurrency] then
					Val := VarAsType(TL.Detail[i].Valeurs[j],vardouble) ;
				dateL := TL.Detail[i].GetValue(NomChamp) ;
				if strtodate(dateL) <= varastype(ValeurChamp,vardate) then
					result := result + val ;
			end ; // if
		end ; // if
	end ; //For
end ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 04/07/2002
Modifié le ... :   /  /    
Description .. : retourne le dernier solde connu
Mots clefs ... : SOLDE
*****************************************************************}
Function GetSoldePrec(CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
var
	SQL 	: string ;
	Q		: TQuery ;
	Solde	: Double ;
begin
	Solde := 0 ;
	//Le principe est de récupérer à une date, le max d'une numéro de pièce
	//tel que la datecomptable soit au plus proche ancienne à la date.
	SQL := 'select TE_DATEComptable, TE_NUMEROPIECE, TE_SOLDEDEV from trecriture where TE_GEneral like "' + CompteGeneral + '" ';
	SQL := SQL + '	and te_datecomptable<"' + UsDateTime(strtodate(Endate))+ '" order by TE_DATECOMPTABLE desc, TE_NUMEROPIECE desc';
	Q :=OpenSQL(SQL, True);
	solde :=Q.FindField('TE_SOLDEDEV').asFloat ;
	dateCalc := Q.FindField('TE_DATECOMPTABLE').Asstring;
	Ferme(Q);
	Result:=Solde;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : retourne le dernier solde en valeur connu
Mots clefs ... : SOLDE;VALEUR
*****************************************************************}
Function GetSoldeValPrec(CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
var
	SQL 	: string ;
	Q		: TQuery ;
	Solde	: Double ;
begin
	Solde := 0 ;
	//Le principe est de récupérer à une date, le max d'une numéro de pièce
	//tel que la datecomptable soit au plus proche ancienne à la date.
	{SQL := 'Select TE_SOLDEDEV, TE_DATECOMPTABLE from TRECRITURE where TE_GENERAL="' ;
	SQL := SQL + CompteGeneral + '" and TE_NUMEROPIECE=(select max(TE_NUMEROPIECE) from trecriture ' ;
	SQL := SQL + 'where TE_DATECOMPTABLE=(select max(TE_DATECOMPTABLE) from trecriture where te_datecomptable<"' ;
	SQL := SQL + UsDateTime(strtodate(Endate))+ '"))' ;
	}
	SQL := 'select TE_DATEVALEUR,TE_DATECOMPTABLE, TE_NUMEROPIECE, TE_SOLDEDEV from trecriture where TE_GEneral like' ;
	SQL := SQL + ' "' + CompteGeneral +'"';
	SQL := SQL + ' and TE_DATECOMPTABLE<=' ;
	SQL := SQL + ' "' + UsDateTime(strtodate(Endate)) + '"' ;
	SQL := SQL + ' order by TE_DATECOMPTABLE desc, TE_NUMEROPIECE desc';
	Q :=OpenSQL(SQL, True);
	solde :=Q.FindField('TE_SOLDEDEV').asFloat ;
	dateCalc := Q.FindField('TE_DATECOMPTABLE').Asstring;
	Ferme(Q);
	Result:=Solde;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... : 04/07/2002
Description .. : retourne le solde en opé d'un compte général
Mots clefs ... : SOLDE
*****************************************************************}
Function GetSolde(CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
var
	SQL 	: string ;
	Q		: TQuery ;
	Solde	: Double ;
	OK		: Boolean ;
begin
		Solde := 0 ;
		OK := false ;
		//Le principe est de récupérer à une date, le max d'une numéro de pièce
		//tel que la datecomptable soit au plus proche ancienne à la date.
		//ancienne requete
		{
		SQL := 'Select TE_SOLDEDEV, TE_DATECOMPTABLE from TRECRITURE where TE_GENERAL="' ;
		SQL := SQL + CompteGeneral + '" and TE_NUMEROPIECE=(select max(TE_NUMEROPIECE) from trecriture ' ;
		SQL := SQL + 'where TE_DATECOMPTABLE=(select max(TE_DATECOMPTABLE) from trecriture where te_datecomptable<="' ;
		SQL := SQL + UsDateTime(strtodate(Endate))+ '") AND TE_GENERAL="' + CompteGeneral + '")' ;
		}

		SQL := 'select TE_DATEComptable, TE_NUMEROPIECE, TE_SOLDEDEV from trecriture where TE_GEneral like "' + CompteGeneral
                + '" and te_datecomptable<="' + UsDateTime(strtodate(Endate))+ '" order by TE_DATECOMPTABLE desc, TE_NUMEROPIECE desc';
		Q :=OpenSQL(SQL, True);
		if not Q.EOF then
		begin
			solde :=Q.FindField('TE_SOLDEDEV').asFloat ;
			dateCalc := Q.FindField('TE_DATECOMPTABLE').Asstring;
			OK := True ;
		end ;
		Ferme(Q);
		Result:=Solde;
		//RRO if not OK then TrShowMessage('Calcul de solde', 7, '', '') ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... : 04/07/2002
Description .. : retourne le solde en valeur d'un compte général
Mots clefs ... : SOLDE;VALEUR
*****************************************************************}
Function GetSoldeValeur(CompteGeneral : string ; Endate : string; var datecalc : string) : double ;
var
	SQL 	: string ;
	Q		: TQuery ;
	Solde	: Double ;
	OK		: Boolean ;
begin
	Solde := 0 ;
		OK := false ;
		//Le principe est de récupérer à une date, le max d'une numéro de pièce
		//tel que la datecomptable soit au plus proche ancienne à la date.
		{
		SQL := 'Select TE_SOLDEDEV, TE_DATEVALEUR from TRECRITURE where TE_GENERAL="' ;
		SQL := SQL + CompteGeneral + '" and TE_NUMEROPIECE=(select max(TE_NUMEROPIECE) from trecriture ' ;
		SQL := SQL + 'where TE_DATEVALEUR=(select max(TE_DATEVALEUR) from trecriture where TE_DATEVALEUR<="' ;
		SQL := SQL + UsDateTime(strtodate(Endate))+ '"))' ;
    }
		SQL := 'select TE_DATEVALEUR, TE_NUMEROPIECE, TE_SOLDEDEV from trecriture where TE_GEneral like "' + CompteGeneral + '" ';
		SQL := SQL + '	and te_datevaleur<="' + UsDateTime(strtodate(Endate))+ '"	order by TE_DATEVALEUR desc, TE_NUMEROPIECE desc';
		Q :=OpenSQL(SQL, True);
		if not Q.EOF then
		begin
			solde :=Q.FindField('TE_SOLDEDEV').asFloat ;
			dateCalc := Q.FindField('TE_DATEVALEUR').Asstring;
			OK := True ;
		end ;
		Ferme(Q);
		Result:=Solde;
		if not OK then TrShowMessage('Calcul de solde', 7, '', '') ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... : 04/07/2002
Description .. : retourne les conditions de financement et placement
Mots clefs ... : CONDITION;FINANCEMENT;PLACEMENT
*****************************************************************}
function GetConditionFinPlac (var ConditionFinPlac : TRCONDITIONFINPLAC; Transac : string ; Banque : string ; Compte : string ; societe : string) : boolean ;
Var
	Q : TQuery ;
		SQL : string ;
Begin
	Result := false ;
	SQL := 'SELECT TCF_AGIOSDEDUIT,TCF_AGIOSPRECOMPTE, TCF_BASECALCUL,TCF_NBJOURBANQUE,TCF_TAUXFIXE,TCF_TAUXVAR,TCF_VALMAJORATION,TCF_VALMULTIPLE,TCF_VALTAUX,TCF_NBJOURENCAISS,TCF_NBJOURMINAGIOS FROM CONDITIONFINPLAC WHERE ' ;
    SQL := SQL + ' TCF_CONDITIONFP="' + Transac + '"' ;
	if length(Compte)>0 then SQL := SQL + ' AND TCF_GENERAL="' + compte + '"'
    else
    begin
        SQL := SQL + ' AND TCF_BANQUE="' + banque + '"' ;
        SQL := SQL + ' AND TCF_SOCIETE="' + societe + '"' ;
    end ;

    if length(Transac)>0 then
    begin
       Q := OpenSQL(SQL, True);
       if not Q.EOF then
       begin
          with  ConditionFinPlac do
               begin
                    TCF_AGIOSPRECOMPTE  := Q.FindField('TCF_AGIOSPRECOMPTE').asstring ;
                    TCF_AGIOSDEDUIT     := Q.FindField('TCF_AGIOSDEDUIT').asstring ;
                    TCF_TAUXFIXE        := Q.FindField('TCF_TAUXFIXE').asstring ;
                    TCF_BASECALCUL      := Q.FindField('TCF_BASECALCUL').asInteger ;
                    TCF_TAUXPRECOMPTE   := Q.FindField('TCF_TAUXPRECOMPTE').asstring ;
                    TCF_VALTAUX         := Q.FindField('TCF_VALTAUX').AsFloat ;
                    TCF_TAUXVAR         := Q.FindField('TCF_TAUXVAR').asstring ;
                    TCF_VALMAJORATION   := Q.FindField('TCF_VALMAJORATION').AsFloat ;
                    TCF_VALMULTIPLE     := Q.FindField('TCF_VALMULTIPLE').AsFloat ;
                    TCF_NBJOURBANQUE    := Q.FindField('TCF_NBJOURBANQUE').AsFloat ;
                    TCF_NBJOURENCAISS   := Q.FindField('TCF_NBJOURENCAISS').AsFloat ;
                    TCF_NBJOURMINAGIOS  := Q.FindField('TCF_NBJOURMINAGIOS').AsFloat ;
                    Result := True ;
               end;
        end;
        Ferme(Q);
    end ;
End ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... : 06/03/2002
Description .. : retourne une TOB remplie des conditions de découverts
Suite ........ : correspondant à la banque et au flux demandé ...
Mots clefs ... : CONDITION DECOUVERT
*****************************************************************}
function GetConditionDec (Flux : string ; Agence : string; TobC : TOB ) : boolean ;
var SQL : String ;
		Q : TQuery ;
		TobL: TOB;
begin
		SQL := 'SELECT * FROM CONDITIONDEC Where TCN_FLUX="'+ Flux + '"' ;
		SQL := SQL + ' AND TCN_AGENCE="' + Agence + '"' ;
		Q := OpenSQL(SQL,true) ;
		if not Q.EOF then
			TobC.LoadDetailDB( 'cond', '', '', Q, False, True)
		else
		begin
			TobL := TOB.Create('', TobC, -1);
			TobL.AddChampSupValeur('TCN_FLUX',flux);
			TobL.AddChampSupValeur('TCN_AGENCE',agence);
			TobL.AddChampSupValeur('TCN_SOCIETE','');
			TobL.AddChampSupValeur('TCN_DEVISE','EUR');
			TobL.AddChampSupValeur('TCN_AUTORISATION',0);
			TobL.AddChampSupValeur('TCN_TYPECALCFRAIS',0);
			TobL.AddChampSupValeur('TCN_MAJOTAUX1',0);
			TobL.AddChampSupValeur('TCN_MAJOTAUX2',0);
			TobL.AddChampSupValeur('TCN_MAJOTAUX3',0);
			TobL.AddChampSupValeur('TCN_MULTAUX1',0);
			TobL.AddChampSupValeur('TCN_MULTAUX2',0);
			TobL.AddChampSupValeur('TCN_MULTAUX3',0);
			TobL.AddChampSupValeur('TCN_TAUXREF1',0);
			TobL.AddChampSupValeur('TCN_TAUXREF2',0);
			TobL.AddChampSupValeur('TCN_TAUXREF3',0);
			TobL.AddChampSupValeur('TCN_PLAFOND1',0);
			TobL.AddChampSupValeur('TCN_PLAFOND2',0);
			TobL.AddChampSupValeur('TCN_BASECALCUL',0);
			TobL.AddChampSupValeur('TCN_DATECONTRAT',StrToDate(NoDate));
			TobL.AddChampSupValeur('TCN_PERIODE',0);
			TobL.AddChampSupValeur('TCN_CALCULSOLDE',0);
			TobL.AddChampSupValeur('TCN_SOLDEVALEUR',0);
			TobL.AddChampSupValeur('TCN_NBJOUR',0);
			TobL.AddChampSupValeur('TCN_TYPEPREMIER',0);
			TobL.AddChampSupValeur('TCN_CPFD',0);
			TobL.AddChampSupValeur('TCN_PLAFONNEE',0);
			TobL.AddChampSupValeur('TCN_LIEAUTO','-');
		end ;

		Ferme(Q);
		result := true ;
End ;



{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Retourne la devise d'un compte général
Mots clefs ... : DEVISE, COMPTE, GENERAL
*****************************************************************}
Function RetDeviseCompte(CompteGeneral : string) : string ;
Var SQL : String ;
		Q : TQuery ;
BEGIN
		Result := '';
		SQL := 'SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL ="' ;
    SQL := SQL + CompteGeneral + '"';
    Q	 :=OpenSQL(SQL, True);
    if not Q.EOF then result := Q.FindField('BQ_DEVISE').AsString;
		ferme(Q);

END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 04/07/2002
Modifié le ... : 04/07/2002
Description .. : retour l'a agence d'un compte
Mots clefs ... : COMPTE;AGENCE
*****************************************************************}
Function GetCompteAgence(Compte : string) : string ;
Var SQL : String ;
		Q : TQuery ;
BEGIN
		Result := '';
		SQL := 'SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_GENERAL="' ;
		SQL := SQL + Compte + '"';
		Q	 :=OpenSQL(SQL, True);
		if not Q.EOF then result := Q.FindField('BQ_BANQUE').AsString;
		ferme(Q);
END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /    
Description .. : retourne la devise d'une banque
Mots clefs ... : DEVISE; BANQUE
*****************************************************************}
Function RetDeviseBanque(Banque : string) : string ;
Var SQL : String ;
		Q : TQuery ;
BEGIN
	Result := '';

	SQL := 'SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_BANQUE = "' ;
	SQL := SQL + Banque + '"' ;
	Q	 :=OpenSQL(SQL, True);
	if not Q.EOF then result := Q.FindField('BQ_DEVISE').AsString;
    Ferme(Q);

END ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Retourne le nb de décimale d'une devise donnée
Mots clefs ... : DEVICE;DECIMALE
*****************************************************************}
Function CalcDecimaleDevise(Devise  : string) : integer ;
Var SQL : String ;
		Q : TQuery ;
BEGIN
		Result := 2;
		SQL := 'SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="' ;
		SQL := SQL + Devise + '"';
	  Q	 :=OpenSQL(SQL, True);
	  if not Q.EOF then result := Q.FindField('D_DECIMALE').AsInteger;
      Ferme(Q);
END ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Retourne le cours contre l'Euro, le plus proche en date,
Suite ........ : d'une devise à une date donnée
Mots clefs ... : DATE;COURS;DEVISE;PARITE;EURO
*****************************************************************}
function RetPariteEuro(Devise : string; Date : TDateTime ) : double ;
Var SQL 	: String ;
	Q		: TQuery ;
Begin
	Result := 1 ;
	  if Devise <> 'EUR' then
	  begin
		 SQL := 'SELECT D_MONNAIEIN,D_PARITEEURO FROM DEVISE WHERE D_DEVISE = "' + devise + '"' ;
		 Q 	:= OpenSQL(SQL,TRue);
		 if not Q.EOF Then
		 begin
			if Q.FindField('D_MONNAIEIN').asString = 'X' then
			Begin
			   Result := Q.FindField('D_PARITEEURO').AsFloat;
			end
			else
			begin
			   Ferme(Q) ;
			   SQL := 'SELECT H_DATECOURS,H_DEVISE,H_TAUXREEL FROM CHANCELL ' ;
			   SQL := SQL + 'WHERE H_DEVISE = "' + devise + '" ' ;
			   SQL := SQL + 'AND H_DATECOURS=(select max(H_DATECOURS) from chancell ' ;
			   SQL := SQL + 'where H_DATECOURS<="' + usdatetime(Date) +'" ' ;
			   SQL := SQL + 'AND H_DEVISE="' + devise + '" )';
			   Q := OpenSQL(SQL,true);
			   if not Q.EOF Then result := Q.FindField('H_TAUXREEL').AsFloat;
			end ;
		 end ;
		 Ferme(Q) ;
		end ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 25/01/2002
Modifié le ... : 03/09/2002
Description .. : Donne le code flux à partir du code transaction
Mots clefs ... : FLUX;TRANSACTION
*****************************************************************}
function GetCodeFlux(CodeTransac, TypeFlux: String): String;
var	Q : TQuery;
begin
	Result := '';
	Q := OpenSQL('Select TTR_'+TypeFlux+' from TRANSAC where TTR_TRANSAC="'+CodeTransac+'"', True);
	if not Q.EOF Then
		Result := Q.Fields[0].AsString;
	Ferme(Q);

end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 03/09/2002
Modifié le ... : 03/09/2002
Description .. : REtourne le type de flux d'une catégorie de transaction
Mots clefs ... : CATTRANSAC; FLUX
*****************************************************************}
function GetTypeFlux(CodeCatTransac : string ): String;
var	Q : TQuery;
begin
	Result := '';
	Q := OpenSQL('Select TCA_TYPEFLUX from CATTRANSAC where TCA_CATTRANSAC="'+CodeCatTransac+'"', True);
	if not Q.EOF Then
		Result := Q.Fields[0].AsString;
	Ferme(Q);

end;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 17/09/2002
Modifié le ... :   /  /    
Description .. : Mets en forme les RIB 
Mots clefs ... : 
*****************************************************************}
function calcRIB(pays : string ; banque : string ; guichet : string ; compte : string; cle : string) : string ;
begin
   if (pays='ES') then
      result := trim(banque) + trim(guichet) + trim(cle) + trim(compte)
      else if (pays='FR') then
         result := trim(banque) + trim(guichet) + trim(compte) + trim(cle)
         else result := trim(banque) + trim(guichet) + trim(compte) + trim(cle);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 28/03/2002
Modifié le ... : 17/09/2002
Description .. : retourne le IBAN d'un compte
Mots clefs ... : IBAN;COMPTE
*****************************************************************}
function calcIBAN(pays : string ; RIB : string) : string ;
Var St2, St3, St4, cleIBAN, ret, concat, strInter : String ;
		ii : Byte ;
		IBAN, cleL : integer ;
		i : Integer ;
BEGIN
	//lRIB=litnombreIBAN(Banque+guichet+cle);
	//lPays=litnombreIBAN(pays);
	Result:='' ;
	St2:=trim(RIB)+Trim(Pays)+'00' ;
	if Length(St2)<10 then exit ;
	St2:=UpperCase(St2) ;
	//Transforme les lettres en chiffres selon le NEE 5.3
	i:=1 ;
	while i<Length(St2) do
	begin
		if St2[i] in ['A'..'Z'] then
		BEGIN
			ii:=Ord(St2[i])-65 ;
			st4:= copy(st2,1,i-1) + inttostr(10+ii) + copy(st2,i+1, length(st2));
			st2:=st4 ;
		END ;
		inc(i);
	end ;
	
	i := 1 ;
	ret := '' ;
	cleL := 0 ;
	st4:='';
	//On découpe par tranche de 9
	//On calcul la clé via mod 97 puis on fait clé + reste du rib
	//while (i<length(st2)) do
	for i:=1 to (length(st2) div 9)+1 do
	begin
		st4 := copy(st2,1,9) ;
		delete(st2,1,9);
		strInter := inttostr(cleL)+st4 ;
		cleL := strtoint64(strinter) mod 97 ;
	end ;
	//une fois fini, on calcul 98-clé
	cleIBAN := inttostr(98-(cleL  mod 97));
	if length(cleIBAN)=1 then  cleIBAN := '0' + cleIBAN ;
	Result:= trim(Pays)+ trim(cleIBAN)+ trim(RIB) ; 
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 26/12/2001
Modifié le ... :   /  /
Description .. : Retourne le code AFB d'une transaction
Mots clefs ... : TRANSACTION;AFB
*****************************************************************}
function GetCodeAFB(transac : string; TypeFlux : string) : string ;
Var SQL : string ;
	Q : TQuery ;
begin
			result := '' ;
        SQL := 'SELECT TFT_CODECIB,TFT_FLUX FROM FLUXTRESO WHERE TFT_FLUX=(' ;
        SQL := SQL + 'SELECT TTR_'+TypeFlux +' FROM TRANSAC WHERE TTR_TRANSAC="' + Transac + '")';
        Q:=OpenSQL(SQL, True);
        if not Q.EOF Then
        begin
        	result := Q.FindField('TFT_CODECIB').AsString;
				end ;
		Ferme(Q);

end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... : 04/07/2002
Description .. : retourne le sens d'une transac
Mots clefs ... : TRANSAC;SENS
*****************************************************************}
function GetSensFlux(transac : string; TypeFlux : string) : string ;
Var SQL : string ;
	Q : TQuery ;
    //sens : string ;
begin

    	result := '' ;
        SQL := 'SELECT TTL_SENS FROM TYPEFLUX WHERE TTL_TYPEFLUX=(' ;
        SQL := SQL + 'SELECT TFT_TYPEFLUX FROM FLUXTRESO WHERE TFT_FLUX=(' +
        		'SELECT TTR_'+TypeFlux+ ' From transac where TTR_TRANSAC="'   + Transac + '"))';
        Q:=OpenSQL(SQL, True);
        if not Q.EOF Then
        begin
        	result := Q.FindField('TTL_SENS').AsString;      //D=DEBIT;C=CREDIT
        end ;

    	Ferme(Q);

end;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /    
Description .. : retourne le cours d'un taux de référence à un date
Mots clefs ... :
*****************************************************************}
function GetValTauxRef(TauxRef : string; Le : string; var Cotation : double) : boolean ;
Var SQL : String ;
	Q	: TQuery ;
BEGIN
		result := false ;
		SQL:= 'SELECT MAx(TTA_COTATION) COTATION, Max(TTA_DATE), max(TTA_CODE) FROM COTATIONTAUX WHERE TTA_CODE="'+TauxRef+'"' +
				' AND TTA_DATE<="' + USDATETIME(strtodatetime(Le)) + '"' ;
		Q:=OpenSQL(SQL, True);
		if not Q.EOF then
		BEGin
			Cotation := Q.FindField('COTATION').AsFloat;
			Result := True ;
		END ;
		FERME(Q);
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 08/01/2002
Modifié le ... :   /  /    
Description .. : Retourne dans Banque, Guichet, Num, cle
Suite ........ : les valeurs du RIB correspond
Suite ........ : ET
Suite ........ : retourne dans result la concaténation du RIB ENTIER
Mots clefs ... : RIB;BANQUE;GUICHET;COMPTE;NUMERO;CLE;CLERIB
*****************************************************************}
function GetRIB(var Banque : string ; var Guichet : string ; var Num : string ; var cle : string ; CompteGeneral : string) : string ;
Var SQL : string ;
	Q : TQuery ;
begin

  	result := '' ;
	SQL := 'SELECT BQ_ETABBQ,BQ_GUICHET,BQ_NUMEROCOMPTE,BQ_CLERIB FROM BANQUECP WHERE BQ_GENERAL="' + CompteGeneral +'"' ;
   Q:=OpenSQL(SQL, True);
   if not Q.EOF Then
   begin
   Banque := Q.FindField('BQ_ETABBQ').AsString;
      Guichet := Q.FindField('BQ_GUICHET').AsString ;
      Num := Q.FindField('BQ_NUMEROCOMPTE').AsString ;
      Cle := Q.FindField('BQ_CLERIB').AsString ;
   result := Banque ;
      result := result + Guichet;
      result := result + Num ;
      result := result + cle;
       end ;
   Ferme(Q);

end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 25/01/2002
Modifié le ... :   /  /
Description .. : Copie ou met à jour dans la table TRECRITURE les
Suite ........ : écritures comptables nouvelles ou modifiées depuis une
Suite ........ : date, pour un ou tous les comptes.  ABANDONNE
Mots clefs ... : ECRITURES
*****************************************************************}
{function UpdateTrEcriture(DateDepart: TDateTime; Compte: String) : Boolean;
var	Q: TQuery;
	SQL: String;
	TrEcrit: TREcritures;
	Societe, DateUS: string;
	Banque, Guichet, Num, Cle: String;
begin
	DateUS := UsDateTime(DateDepart);
  SQL := 'Select E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE,E_EXERCICE,E_NUMECHE,E_GENERAL,E_MODEPAIE,'+
			'E_REFINTERNE,E_ETABLISSEMENT,E_QUALIFORIGINE,E_LIBELLE,E_CREERPAR,E_UTILISATEUR,'+
			'E_DATECREATION,E_DATEMODIF,E_DATEPOINTAGE,E_DEVISE,(E_DEBIT-E_CREDIT) MONTANTDEV,'+
			'(E_DEBITEURO-E_CREDITEURO) MONTANT,E_COTATION,E_DATECOMPTABLE '+
			'from ECRITURE where E_GENERAL like "51%" and (E_DATECREATION>="'+DateUS+
			'" or E_DATEMODIF>="'+DateUS+'")';
  if Compte <> '' then
	SQL := SQL + ' and E_GENERAL="'+Compte+'"';
  Result := True;
  try
	Q := OpenSql(SQL, True);
	while not Q.Eof do
	begin
		TrEcrit.TE_JOURNAL			:= Q.FindField('E_JOURNAL').AsString ;
		TrEcrit.TE_NUMEROPIECE		:= Q.FindField('E_NUMEROPIECE').AsInteger;
		TrEcrit.TE_NUMLIGNE 	  	:= Q.FindField('E_NUMLIGNE').AsString ;
		TrEcrit.TE_EXERCICE 	  	:= Q.FindField('E_EXERCICE').AsString ;
		TrEcrit.TE_NUMECHE 	  		:= Q.FindField('E_NUMECHE').AsInteger ;
		TrEcrit.TE_GENERAL 	  		:= Q.FindField('E_GENERAL').AsString ;
		TrEcrit.TE_REFINTERNE 	  	:= Q.FindField('E_REFINTERNE').AsString ;
		Societe						:= Q.FindField('E_ETABLISSEMENT').AsString ;
		TrEcrit.TE_SOCIETE 	  		:= Societe;
		TrEcrit.TE_QUALIFORIGINE	:= Q.FindField('E_QUALIFORIGINE').AsString ;
		TrEcrit.TE_LIBELLE			:= Q.FindField('E_LIBELLE').AsString ;
		TrEcrit.TE_USERCREATION		:= Q.FindField('E_UTILISATEUR').AsString ;
		TrEcrit.TE_USERMODIF		:= '';
		TrEcrit.TE_USERVALID		:= '';
		TrEcrit.TE_USERCOMPTABLE	:= Q.FindField('E_UTILISATEUR').AsString ;
		TrEcrit.TE_DATECREATION		:= Q.FindField('E_DATECREATION').AsString ;
		TrEcrit.TE_DATEMODIF		:= Q.FindField('E_DATEMODIF').AsString ;
		TrEcrit.TE_DATEVALID		:= NoDate;
		TrEcrit.TE_DATECOMPTABLE	:= Q.FindField('E_DATECOMPTABLE').AsString ;
		TrEcrit.TE_DATERAPPRO		:= Q.FindField('E_DATEPOINTAGE').AsString ;
		TrEcrit.TE_DEVISE			:= Q.FindField('E_DEVISE').AsString ;
		TrEcrit.TE_MONTANTDEV		:= Q.FindField('MONTANTDEV').AsFloat ;
		TrEcrit.TE_MONTANT			:= Q.FindField('MONTANT').AsFloat ;
		TrEcrit.TE_COTATION			:= Q.FindField('E_COTATION').AsFloat ;

		GetRIB(Banque, Guichet, Num, Cle, TrEcrit.TE_GENERAL);
		TrEcrit.TE_CODEBANQUE 	  	:= Banque;
		TrEcrit.TE_CODEGUICHET   	:= Guichet;
		TrEcrit.TE_NUMCOMPTE 	  	:= Num;
		TrEcrit.TE_CLERIB 		  	:= Cle;

		TrEcrit.TE_CODECIB := RechDom('TAFB', Q.FindField('E_MODEPAIE').AsString, False);
		TrEcrit.TE_DATEVALEUR := DateToStr(CalcDateValeur(TrEcrit.TE_CODECIB, TrEcrit.TE_GENERAL, Q.FindField('E_DATECOMPTABLE').AsDateTime));

		TrEcrit.TE_CODEFLUX := '';
		TrEcrit.TE_NUMTRANSAC := CodeModuleCpt+TrEcrit.TE_SOCIETE+IntToStr(TrEcrit.TE_NUMEROPIECE); // Crée une valeur unique (varchar(17))

		TrEcrit.TE_SOLDEDEV := TrEcrit.TE_MONTANTDEV+GetSoldePrec(TrEcrit.TE_GENERAL, TrEcrit.TE_DATECOMPTABLE, DateUS);
		TrEcrit.TE_SOLDEDEVVALEUR := 0;

		Result := EcritTREcriture(TrEcrit, Q.FieldByName('E_DATEMODIF').AsDateTime < DateDepart);  // True = pas de modif : ajout
		if Result then
		begin		// Recalcule les soldes
			Result := RecalculSolde(TrEcrit.TE_MONTANTDEV, TrEcrit.TE_GENERAL, TrEcrit.TE_DATECOMPTABLE);
			if Result then
				Result := RecalculSoldeVal(TrEcrit.TE_MONTANTDEV, TrEcrit.TE_GENERAL, TrEcrit.TE_DATECOMPTABLE);
		end;
		if Not Result then
			Break;

		Q.Next;
	end;

  finally
	Ferme(Q);
  end;
end; }

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Retourne à une date le montant contrevalorisé d'un montant
Suite ........ : en une devise dans une autre devise
Mots clefs ... : MONTANT;DEVISE;CONTREVALORISE
*****************************************************************}
Function RetContreValeur(Date: TDateTime; Montant : double; DevMontant : string ; DevContrevalorise : string) : double ;
Var parite1 : double ;
Parite2 : Double ;
Begin
		parite1 := RetPariteEuro(DevMontant,Date ) ;
		parite2 := RetPariteEuro(DevContrevalorise,Date ) ;
	  if parite1 <> 0 then
			Result := Montant * parite2/parite1
	  else
		Result := 0 ;
End ;


{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 04/07/2002
Modifié le ... :   /  /    
Description .. : Retourne les paramètre société
Mots clefs ... : 
*****************************************************************}
Function TrGetParamSoc(champs:string) : Variant ;
begin
	result:='' ;
	if length(champs)> 0 then result:= GetParamSoc(champs);
end ;

end.

