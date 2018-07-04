unit AffEcheanceUtil;

interface
uses AffaireUtil,FactUtil,SAISUTIL,HEnt1,HCtrls,SysUtils,EntGC,paramsoc,
			Aglinit,
{$IFDEF EAGLCLIENT}
     MainEagl,
{$ELSE}
		 fe_main,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

{$IFDEF AGLINF545}
     ActiviteUtil,
{$ENDIF}
     UTOB;

Function GenerEcheancesAffaire (CodeAff : string; DateDebCal,DateFinCal : TdateTime): Boolean;

// *** Suppression d'écheances ***
Function SupEcheancesAffaire (Affaire : string; EchFact : Boolean; DateTraite : TDateTime) : Boolean;
Function AjusteDateGenerSurEch (Affaire : string; EchFact,bUpdate : Boolean; DateTraite,DateDebGen : TDateTime) : TDateTime;

// *** Fonctions spécifiques contrat ***
function GetDateDebutPeriode(AcsPeriodicite:string; AcdDateDebut:TDateTime;AciInterval : integer;MethodeCalcul  : string):TDateTime;
function CalculProrata(TypEch,Periodicite : string;NbInter : Integer;DebEch,DebPer,FinPer :TDateTime;var DebFac,Finfac : TdateTime) : double;
// Moteur de calcul des écheances
Procedure CalculMtEch (TobDet: TOB; DEV : RDEVISE; zmont : Double; SaisieContre : Boolean);
    //mcd 03/07/03 ajout profil sur 3fct suivantes
Procedure UtilCalculEcheances(MtGlobal:Double;Profil,ModeCalcul,CodeAff,zRepAct, ztier, ztyp_fac : string;
                              TypeModif:T_TypeModifAff; zinterval: integer;zmont: double;
                              DateDebut , DateFin, DateLiquid,DateDebCal,DateFinCal,DateResil : TdateTime;
                              Dev :RDEVISE; SaisieContre, bMtPeriodique, bPourcentage : Boolean;
                              zMultiEche:string; ModeMontantDoc,MethodeCalcul  : string;
                              DateDebutFac,DateFinFac  : TDateTime; TermeEche : string);
 
procedure ConstitueTOBEcheances (TOBEch : TOB; MtGlobal:Double; Profil, CodeAff, TypeCalcul, zRepAct, ztier, ztyp_fac : string; TypeModif:T_TypeModifAff; zinterval: integer;
             zmont: double; DateDebut , DateFin, DateLiquid,DateDebCal,DateFinCal,DateResil : TdateTime; Dev :RDEVISE; SaisieContre, bMtPeriodique, bPourcentage : Boolean;
             zMultiEche:string; ModeMontantDoc,MethodeCalcul : string; DateDebutFac  : TDateTime; TermeEche : string);

Procedure UtilCalculEcheancesBudget(Profil, CodeAff, zRepAct, zGenerAuto : string;
                                Dev :RDEVISE; SaisieContre : Boolean; zMultiEche:string);
function GetTOBEcheance (TobEch:TOB; Rang:integer; Profil, sLiquidative, CodeAff,  zRepAct, ztier, ztyp_fac : string;
             zmont: double; DateEch,DateFinEche : TdateTime; Dev :RDEVISE; SaisieContre : Boolean; TermeEche,MethodeCalcul,typeCalcul : string):longint;

function RecaleEcheanceFinMois( AcdDateEntree, AcdDateDebEch : TDateTime ) : TDateTime;
function GetLeChoixSurProrata (TypeCalcul,TypeEcart : string;NbMoisInPeriod : double;D1,D2 : TDateTime) : integer;
function LanceRecalcEches (TOBAffaire : TOB; DEV : RDevise) : boolean;
function ConstitueLibEcheContrat ( TOBEche : TOB ) : string;

Type T_GenerEchAff = Class(TObject)
     CodeAff : string;
     DateDebCal,DateFinCal : TdateTime;
     Procedure GenereLesEcheances;
     End ;


implementation
uses dateutils,hmsgbox;

Function GenerEcheancesAffaire (CodeAff : string; DateDebCal,DateFinCal : TdateTime): Boolean;
Var GenEch : T_GenerEchAff;
    io:TIoErr ;
begin
Result := True;
GenEch:=T_GenerEchAff.Create;
GenEch.CodeAff:=CodeAff;
GenEch.DateDebCal:=DateDebCal;
GenEch.DateFinCal:=DateFinCal;

io:=Transactions(GenEch.GenereLesEcheances,1) ;
Case io of
    oeUnknown : BEGIN
                result:=false;
                END;
    oeSaisie  : BEGIN
                result:=false;
                END ;
   END ; 
GenEch.Free;
end;

Procedure T_GenerEchAff.GenereLesEcheances;
Var TobAff : TOB;
    SaisieContre, bMtPeriodique, bPourcentage : Boolean;
    DateDebGen, DateFinGen, DateLiquid : TDateTime;
    zMont : double;
    TypeModif : T_TypeModifAff;
    DEV  : RDEVISE;
    Q :Tquery; //GM
    ret : boolean;
    zchp,req : string;
    ModeMontantDoc,MethodeCalcul : string;
    DateDebutfac,DateFinfac : TdateTime;
begin
if CodeAff = '' then Exit;
{
if DateDebCal = iDate1900 then Exit;
If (DateFinCal = idate2099) then DateFinCal := PlusDate (DateDebCal,1,'A');
}

if DateFinCal = idate2099 then Exit;   // gm 22/08/02


TobAff := TOB.Create('AFFAIRE',Nil,-1);
try
// gm 07/01/2002  Trop lourd, je limite le select
//if Not(RemplirTOBAffaire (CodeAff,TobAff)) then Exit;
	ret := true;
	zchp := 'AFF_AFFAIRE,AFF_ETATAFFAIRE,AFF_DATEDEBGENER,AFF_DATEFINGENER,AFF_DATEFACTLIQUID';
  zchp := zchp+',AFF_DEVISE,AFF_SAISIECONTRE,AFF_GENERAUTO,AFF_POURCENTAGE,AFF_MONTANTECHEDEV';
  zchp := zchp+',AFF_METHECHEANCE,AFF_PERIODICITE,AFF_INTERVALGENER,AFF_TOTALHTGLODEV';
  zchp := zchp+',AFF_REPRISEACTIV,AFF_TIERS,AFF_DATERESIL,AFF_PROFILGENER';

  req := 'SELECT '+ zchp+ ' FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAff+'"';
  Q:=OpenSQL(req,True,-1,'',true) ;
   If (Not Q.EOF) then
   		TOBAff.SelectDB('',Q)
   Else Ret:=False;
   Ferme(Q);

   if ret =false then exit;

	 if (TobAff.GetValue('AFF_ETATAFFAIRE')='CLO') or (TobAff.GetValue('AFF_ADMINISTRATIF')='X') then Exit;  // gm 22/08/02
// Récup des valeurs
DateDebGen  := TobAff.GetValue('AFF_DATEDEBGENER');
DateFinGen  := TobAff.GetValue('AFF_DATEFINGENER');
DateLiquid  := TobAff.GetValue('AFF_DATEFACTLIQUID');
Dev.Code := TobAff.GetValue('AFF_DEVISE'); GetInfosDevise(DEV) ;
SaisieContre := (TobAff.GetValue('AFF_SAISIECONTRE')='X');
bPourcentage := (TobAff.GetValue('AFF_GENERAUTO')='POU') or (TobAff.GetValue('AFF_GENERAUTO')='POT');
ModeMontantDoc := TOBAFF.getValue('AFF_DETECHEANCE');
MethodeCalcul := TOBAFF.getValue('AFF_METHECHEANCE');
DateDebutfac := TOBAFF.GetValue('AFF_DATEDEBGENER');
DateFinfac := TOBAFF.GetValue('AFF_DATEFINGENER');

if ((TobAff.GetValue('AFF_GENERAUTO')='POU') or (TobAff.GetValue('AFF_GENERAUTO')='POT')) then
   zmont :=TobAff.GetValue('AFF_POURCENTAGE')
else
   zmont :=TobAff.GetValue('AFF_MONTANTECHEDEV');

//If (DateFin = idate2099) then DateFin := PlusDate (DateDebut,1,'A');


// En mode de génération Contrat et méthode d'échéance civile, on place la date de début au début de la période
// suivant la périodicité : Annuelle, mensuelle, hebdomadaire
if (TobAff.GetValue('AFF_GENERAUTO')='CON') and (TobAff.GetValue('AFF_METHECHEANCE')='CIV') then
    DateDebGen:= getDateDebutPeriode ( TobAff.GetValue('AFF_PERIODICITE'), DateDebGen, TobAff.GetValue('AFF_INTERVALGENER'),methodeCalcul);

bMtPeriodique := True; // principe : on part du montant périodique
TypeModif := [tmaDate];
  //mcd 03/07/03 ajout profilgener
UtilCalculEcheances(TobAff.GetValue('AFF_TOTALHTGLODEV'),TobAff.GetValue('AFF_PROFILGENER'),TobAff.GetValue('AFF_PERIODICITE'),CodeAff, 
   TobAff.GetValue('AFF_REPRISEACTIV'), TobAff.GetValue('AFF_TIERS'), TobAff.GetValue('AFF_GENERAUTO'),
   TypeModif, TobAff.GetValue('AFF_INTERVALGENER'), zmont, DateDebGen, DateFinGen, DateLiquid, DateDebCal , DateFinCal,
   TobAff.GetValue('AFF_DATERESIL'),DEV, SaisieContre, bMtPeriodique, bPourcentage, TobAff.GetValue('AFF_MULTIECHE'),ModeMontantDoc,MethodeCalcul,DateDebutfac,DateFinfac,TobAff.GetValue('AFF_TERMEECHEANCE'));

Finally
   TobAff.Free;
   end;
end;

//*************** Suppression des écheances ************************************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 10/08/2001
Modifié le ... : 10/08/2001
Description .. : Fonction de suppression des écheances d'affaires
Mots clefs ... : ECHEANCE;AFFAIRE
*****************************************************************}
Function  SupEcheancesAffaire(Affaire : string; EchFact : Boolean; DateTraite : TDateTime) : Boolean;
Var stFact, stWhereDate : string;
BEGIN
Result:=true;
// Attention on part du principe suivant :
// si écheances facturées, on prend tout ce qui est antérieur à la date de traitement
// sinon postérieur à date de traitement
if EchFact then stFact := 'X' else stFact:='-';
if (DateTraite <> iDate1900) And (DateTraite <> iDate2099) then
   begin
   if EchFact then
      stWhereDate := 'AND AFA_DATEECHE <="'+UsDateTime(DateTraite) + '"'
   else
      stWhereDate := 'AND AFA_DATEECHE >"'+UsDateTime(DateTraite) + '"';
   end;
ExecuteSQL('DELETE FROM FACTAFF WHERE AFA_ECHEFACT="'+ stFact
          +'" AND AFA_AFFAIRE ="'+ Affaire + '" '+ stWhereDate) ;
END;


{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 10/08/2001
Modifié le ... : 10/08/2001
Description .. : Ajustement des dates d'écheances affaires en fonction des 
Suite ........ : écheances existantes
Mots clefs ... : ECHEANCE;AFFAIRE
*****************************************************************}
Function AjusteDateGenerSurEch (Affaire : string; EchFact,bUpdate : Boolean; DateTraite,DateDebGen : TDateTime) : TDateTime;
Var NewDate : TDateTime;
    stChamp : string;
    Q : TQuery;
begin

// Réalignement date début ou fin d'échéances de génération
if EchFact then
   begin
   NewDate := PlusDate(DateTraite,1,'J');  // voir pb de calcul des écheances ?????
   stChamp := 'AFF_DATEDEBGENER';
   end
else
   begin
   NewDate := DateTraite;
   if (NewDate=iDate1900) or (NewDate=iDate2099) then
      begin
      Q := OpenSQL('SELECT MAX(AFA_DATEECHE) As DATEECHE FROM FACTAFF Where AFA_AFFAIRE ="'+ Affaire+'"', True,-1,'',true);
      if not Q.EOF and (Q.FindField('DATEECHE').AsString <> '') then
         NewDate := StrToDate(Q.FindField('DATEECHE').AsString)
      else NewDate := DateDebGen; // Fin = Debut
      Ferme(Q);
      end;
   stChamp := 'AFF_DATEFINGENER';
   end;
if bUpdate then ExecuteSQL ('UPDATE Affaire SET '+ stChamp + '="'+UsDateTime(NewDate)+ '"WHERE AFF_AFFAIRE ="'+ Affaire+'"') ;
Result := NewDate;
end;


//*************** Fonction spécifique gestion des contrats *********************

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 10/08/2001
Modifié le ... : 10/08/2001
Description .. : recalcul de la date début de période par rapport à la
Suite ........ : m^éthode de calcul du contrat
Mots clefs ... : CONTRAT;CALCUL ECHEANCE
*****************************************************************}
function GetDateDebutPeriode(AcsPeriodicite:string; AcdDateDebut:TDateTime; AciInterval : integer;MethodeCalcul  : string):TDateTime;
var eAnnee, eMois, eJour, sAnnee, sMois, sJour : word;
		zx, Semaine : Integer;
begin
  Result:=0;
  sAnnee:=eAnnee;
  sMois:=1;
  sJour:=1;
	if (AcdDateDebut=idate1900) then Result:=idate1900;

  DecodeDate(AcdDateDebut, eAnnee, eMois, eJour);
  if MethodeCalcul='AN' then
  begin
    Result := EncodeDate(eAnnee, eMois, eJour);
    exit;
  end;

  if (AcsPeriodicite='A') then
  begin
	  // Annuelle
    sAnnee:=eAnnee;
    sMois:=1;
    if MethodeCalcul='CIV' then  sJour:=1 else sJour := eJour;
  end else if (AcsPeriodicite='M') then
  begin
    // Mensuelle
    // attention, on recherche le jour du 1er mois de la période    (modif GM 1/08/01)
    // cad, si je suis Trimestriel , je dois me caler sur les mois 1,4,7 ou 10
    //      si je suis Semestriel, je dois me caler sur les mois 1 ou 7
    sAnnee:=eAnnee;
    sJour:=1;
    if (AciInterval = 1) then
    begin
      sMois:=eMois;
    end else
    begin
      zx := 1;
      sMois := 0;
      while (sMois = 0) do
      begin
        if ( eMois <=  (zx * AciInterval) ) then
          sMois := (zx * AciInterval) - (AciInterval - 1);
        inc(zx);
      end;
    end;
    if MethodeCalcul='CIV' then  sJour:=1 else sJour := eJour;
  end else if (AcsPeriodicite='S') then
  begin
  	// Hebdomadaire
    {$IFDEF AGLINF545}
    Result:=PremierJourSemaineTempo(NumSemaine(AcdDateDebut), eAnnee);
    {$ELSE}
    // PL le 23/05/02 : réparation de la fonction AGL545
    Semaine :=NumSemaine(AcdDateDebut);
    if (eMois=12) and (Semaine=1) then eAnnee := eAnnee + 1;
    if (eMois=1) and ((Semaine=52) or (Semaine=53)) then eAnnee := eAnnee - 1;
    Result:=PremierJourSemaine(Semaine, eAnnee);
    /////////////////////// AGL545
    {$ENDIF}
  end else if (AcsPeriodicite='NBI') then  //FV1 : Modif pour gestion nombre d'intervention
  begin
    Result := now;
  end;
  if (Result=0) then Result := EncodeDate(sAnnee, sMois, sJour);
end;





{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 10/08/2001
Modifié le ... : 10/08/2001
Description .. : fct de calcul de prorata pour le mode Contrat
Suite ........ :
Mots clefs ... : CONTRAT;CALCUL ECHEANCE
*****************************************************************}
// debech : début échéance
//  debper,finper : début de facturation du contrat et date de fin ou résiliation du contrat
//  debfac,finfac : dates de facturation réelles
function  CalculProrata (TypEch, Periodicite : string; NbInter : Integer;
                        DebEch, DebPer, FinPer : TDateTime; var DebFac, Finfac : TdateTime) : double;
Var
		FinEch : Tdatetime;
    NbJrEch, NbJrReel : Double; 					// dates de fin d'échéance
    DFacJ, DFacM, DFacA : Word; 					// décomposition Date debut facture
    FFacJ, FFacM, FFacA : Word;           // décomposition Date fin facture
    FdmJ, FdmM, FdmA : Word;              // décomposition Fin de mois de la date de fin
    wi : integer;
BEGIN
// conseil	result := 1;
  if (Typech = '') then typech := 'CIV';
  if (Periodicite = 'P') then Periodicite := 'J';
  // Calcul de la date de fin théorique de l'échéance
	FinEch := PlusDate (DebEch, NbInter, Periodicite) - 1;

  // recherche de la date de début réelle à facturer DefFac
  // si anniversaire, les échéances sont callés sur la date réelle ,
  // 		 donc date début = date echeances
  if (TypEch = 'CIV') and (DebPer > DebEch) and (DebPer <= FinEch) then
  	DebFac := DebPer
	else
  	DebFac := DebEch;

  // recherche de la date de Fin réelle à facturer FinFac
	if (FinPer < FinEch) then
  	FinFac := Finper
  else
  	FinFac := FinEch;
  // Nbr de jours théorique de l'échéance
	NbJrEch := FinEch - DebEch + 1;
  // Nbr de jours réels de l'échéance
	NbJrReel := FinFac - DebFac + 1;


  // Si on a des mois entiers, on calcul sur un nombre de mois, sinon on calcul
  // en nombre de jours
  // On ne fait pas de calcul sur nombre de mois, si on à plus de 12 mois et si on n'est pas
  // en périodicite mensuelle

  DecodeDate (DebFac, DFacA, DFacM, DFacJ);
  DecodeDate (FinFac, FFacA, FFacM, FFacJ);

  DecodeDate (FinDeMois(FinFac), FdmA, FdmM, FdmJ);

  if (DfacJ <> 1) or (FdmJ <> FFacJ)  or (Nbinter > 12) or  (Periodicite <> 'M') then
    result := arrondi ((NbJrReel / NbJrEch) * NbInter, 2)
  else
  	begin
    if (FFacM < DFacM) then wi := 12 else wi := 0;
    result := (FFacM + wi) - DFacM + 1;
    end;

END;

Procedure CalculMtEch (TobDet: TOB; DEV : RDEVISE; zmont : Double; SaisieContre : Boolean);
Var MtPivot : Double;
BEGIN
  TobDet.PutValue('AFA_MONTANTECHEDEV',zmont);
  // C.B 18/06/2003 Suppression contrevaleur
  //FV1 - FS#359 - La devise n'étant pas renseignée (reprise du tiers dans le contrat) Mtpivot pas calculé
  //ConvertDevToPivot(DEV,zMont,MtPivot);
  TobDet.PutValue('AFA_MONTANTECHE',zmont);
END;


function CalcDaysInPeriod (DateDeb : TdateTime; NbMois : integer) : double; overload;
var DateFin : TdateTime;
		Year,Month,Day : Word;
begin
	DateFin := PlusDate( DateDeb, NbMois, 'M') -1;
  DeCodeDate (DateFin,year,Month,Day); DateFin := EncodeDateTime ( Year,Month,Day,23,59,59,0);
	result  := Arrondi(DaySpan(DateDeb,DateFin),0);
end;

function CalcDaysInPeriod (DateDeb,DateFin : TdateTime) : double; overload;
begin
	result  := Arrondi(DaySpan(DateDeb,DateFin),0);
end;

//**********************************************************************************
//****************Moteur de calcul des echeances              ************
//**********************************************************************************
{***********A.G.L.***********************************************
Auteur  ...... : G. Merieux

Créé le ...... : 03/02/2000

Modifié le ... :   /  /

Description .. : Calcul des échéances de facturation de l'affaire

Mots clefs ... : CALCUL ECHEANCE AFFAIRE

*****************************************************************}
function getFinPeriodefac(LibelleEche : string) : string;
var ipos : integer;
		Tmp : string;
    yy,mm,dd : word;
begin
	// obligation de passer par la bicose dans la table on à pas la date de fin de l'échéance #?*/;
	result := '';
  iPos :=  Pos(' au ',LibelleEche);
  if iPos = 0 then exit;
  tmp := copy(LibelleEche,iPos+4,10);
  DecodeDate (strtoDate(tmp),yy,mm,dd);
  if IsValidDate (yy,mm,dd) then result := tmp;
end;

procedure ConstitueTOBEcheances (TOBEch : TOB; MtGlobal:Double; Profil, CodeAff, TypeCalcul, zRepAct, ztier, ztyp_fac : string; TypeModif:T_TypeModifAff; zinterval: integer;
             zmont: double; DateDebut , DateFin, DateLiquid,DateDebCal,DateFinCal,DateResil : TdateTime; Dev :RDEVISE; SaisieContre, bMtPeriodique, bPourcentage : Boolean;
             zMultiEche:string; ModeMontantDoc,MethodeCalcul : string; DateDebutFac  : TDateTime; TermeEche : string);
Var QQ : Tquery;
    ii, cpt, iMax: Integer;
    TobDet : TOB;
    MtDejaFac,MtResteAFac : double;
    bDejaFac:boolean;
    DateDebEcheance, DateFinEcheance, DateEcheanceSuivante, DerniereDateFac, DateFinTraite:TDateTime;
    stWhereDateCal : string;
    chp,zorder : string; //gm
    IndiceEcheance:longint;
    AnneeDD,MoisDD,JourDD,AnneeDF,MoisDF,MoisInter,JourDF:word;
    bDateOK, bEcheLiquidExist, bCreerEcheLiquid:boolean;
    Lemontant : double;
    LeChoix : integer;
    year,Month,Day,Yx,Mx,Dx : word;
    yearF,MonthF,DayF : word;
    NbJourInyear : integer;
    CouvertureperiodeIncomplete : boolean;
    DeltaMois,DeltaJour,NbDaysInPeriod,NbMonthInPeriod,NbDaysInselection : double;
    D1,D2 : TdateTime;
    DateTmp : string;
    TypeEcart : string; // D date de debut differente ; F Date de fin
    indiceEche : integer;
    DebNewEcheReele : TDateTime;
Begin
  if MtGlobal = 0 then exit;
  bEcheLiquidExist := false;
  bDejaFac := false;   Cpt := 0;
  MtDejaFac:=0;
  DerniereDateFac := idate1900;
  DateDebutfac := StrToDate(DateToStr(DateDebutfac));
  if (datedebut=idate1900) then exit; // gm 22/08/02
  if (DateFinCal<>iDate2099) And (DateFinCal<DateFin)  then DateFinTraite:= DateFinCal
                                                       else DateFinTraite:= DateFin;
  // PA prise en compte de la date de résil si antérieure à la date de fin
  if (DateResil < DateFinTraite) then DateFinTraite := DateResil;

	if Not bPourcentage then MtResteAFac:= MtGlobal-MtDejaFac
                       else MtResteAFac:= 100-MtDejaFac;
   // création Tob contenant toutes les nouvelles échéances
  if (bDejaFac) then
  begin
  	DateDebut := GetDateDebutPeriode (TypeCalcul, DateDebutFac, zinterval,MethodeCalcul)  ;
  end;
  DateDebEcheance := DateDebut;
  DateFinEcheance := PlusDate( DateDebEcheance, zinterval, TypeCalcul );
  DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
  if Not bMtPeriodique and Not bPourcentage then
      zmont := Arrondi(zmont, Dev.Decimale);

  // recadre la date de debut d'écheance sur la période de calcul
  while (DateDebCal <> iDate1900) And (DateDebCal > DateDebEcheance) And (TypeCalcul <>'P') do
  begin
    Inc(Cpt);
    DateDebEcheance := PlusDate( DateDebut, Cpt * zinterval, TypeCalcul );
    DateDebEcheance := RecaleEcheanceFinMois( DateDebEcheance, DateDebut );
    DateFinEcheance := PlusDate( DateDebEcheance, zinterval, TypeCalcul );
    DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
  end;

  // recup max numeche
  iMax := 1;
  if (DateFinEcheance > DateFinTraite ) then DateFinEcheance := DateFinTraite;
  if (TypeCalcul ='P') then
  begin
    DateDebEcheance := DateDebut;
     DateFinEcheance := PlusDate( DateDebEcheance, zinterval, TypeCalcul );
     DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
    if (DateDebEcheance <= DateFinTraite) and (DateDebEcheance >=DateDebCal) then
       GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac, zmont, DateDebEcheance, DateFinEcheance , Dev,
                      SaisieContre,TermeEche,MethodeCalcul,TypeCalcul);
  end else
  begin
      // calcul des échéances
    indiceEche := 1;
    while (DateDebEcheance <= DateFinTraite)  do
    Begin
      //
      // cas des contrats annuels dont les dates ne couvrent pas des périodes annuelles
      //
      DeCodeDate (DateDebEcheance,year,Month,Day); DateDebEcheance := EncodeDateTime (Year,Month,Day,0,0,0,0);
      DeCodeDate (DateDebutFac,year,Month,Day); DateDebutFac := EncodeDateTime (Year,Month,Day,0,0,0,0);
      DeCodeDate (DateFinEcheance,yearF,MonthF,DayF); DateFinEcheance := EncodeDateTime ( YearF,MonthF,DayF,23,59,59,0);
      if DateDebutFac > DateFinEcheance then break; // cas de figure un poil tordu
      if (indiceEche=1) and (pos(typeCalcul,'A;M')>0) then
      begin
        if TypeCalcul = 'A' then
        begin
          NbDaysInPeriod := DaysInAYear (yearF);
          NbMonthInPeriod := 12;
        end else if TypeCalcul = 'M' then
        begin
          NbDaysInPeriod := CalcDaysInperiod (DateDebutfac,zinterval);
          NbMonthInPeriod := zinterval;
        end;
        CouvertureperiodeIncomplete := (Arrondi(DaySpan(DateDebutFac,DateFinEcheance),0) <> NbDaysInPeriod);
        DeltaMois := Arrondi(MonthSpan(DateDebutFac,DateFinEcheance),0);
        DeltaJour := Arrondi(DaySpan(DateDebutFac,DateFinEcheance),0);
        D1 := DateDebutfac;
        D2 := DateFinEcheance;
      end else if (pos(typeCalcul,'A;M')>0) then
      begin
        if TypeCalcul = 'A' then
        begin
          NbDaysInPeriod := DaysInAYear (yearF);
          NbMonthInPeriod := 12;
        end else if Typecalcul = 'M' then
        begin
          NbDaysInPeriod := CalcDaysInperiod (DateDebEcheance,zinterval);
          NbMonthInPeriod := zinterval;
        end;
        CouvertureperiodeIncomplete := (Arrondi(DaySpan(DateDebEcheance,DateFinEcheance),0) <> NbDaysInPeriod);
        DeltaMois := Arrondi(MonthSpan(DateDebEcheance,DateFinEcheance),0);
        DeltaJour := Arrondi(DaySpan(DateDebEcheance,DateFinEcheance),0);
        D1 := DateDebEcheance;
        D2 := DateFinEcheance;
      end;
      //
      if (pos(TypeCalcul,'A;M')>0) and (CouvertureperiodeIncomplete) then
      begin
        if D1 <> DateDebEcheance then TypeEcart := 'D' else TypeEcart := 'F';
        leChoix := GetLeChoixSurProrata (TypeCalcul,TypeEcart,NbMonthInPeriod,D1,D2);
        if (leChoix = 1) then
        begin
          LeMontant := zMont;
        end else if (leChoix = 2) then
        begin
          // prorata sur nb de mois concerné
          if TypeEcart = 'D' then
          begin
            D1 := DEBUTDEMOIS(D1);
            DeltaJour := Arrondi(DaySpan(D1,D2),0);
            Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
          end else
          begin
            D2 := DebutDemois(D2);
            DeltaJour := Arrondi(DaySpan(D1,D2),0);
            Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
          end;
          //                  LeMontant := Arrondi(zmont * ((DeltaMois)/NbMonthInPeriod),V_PGI.OkDecV );
        end else if (leChoix = 3) then
        begin
          // prorata sur le nombre de jour dans l'année concerné ou la période concernée
          Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
        end else
        begin
          // prorata sur nb de mois suivant le mois de début de facturation
          if TypeEcart = 'D' then
          begin
            D1 := Findemois (D1)+1;
            DeltaJour := Arrondi(DaySpan(D1,D2),0);
            Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
          end else
          begin
            D2 := FinDemois(D2);
            DeltaJour := Arrondi(DaySpan(D1,D2),0);
            Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
          end;
        end;

      end else
      begin
        Lemontant := zMont;
      end;
      //
      //            IndiceEcheance:=GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac, Lemontant, DateEcheance,DateFinEcheance, Dev, SaisieContre);
      IndiceEcheance:=GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac, Lemontant, D1,D2, Dev, SaisieContre,TermeEche,MethodeCalcul,TypeCalcul);

      // PL le 18/06/02 : gestion multi echeances
      if (GetParamSoc('SO_AFMULTIECHE')= true) and (zMultiEche='X') then
      // On gère le multi echeance
      if (ztyp_fac<>'ACT') and (ztyp_fac<>'MAN') and (zRepAct<>'NON') then
      // En cas de facturation mixte
      begin
        TobEch.Detail[IndiceEcheance].PutValue('AFA_REPRISEACTIV','NON');
        Inc(iMax);
        GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, 'ACT', 0, DateDebEcheance, DateFinEcheance, Dev, SaisieContre,TermeEche,MethodeCalcul,TypeCalcul);
      end;
      ///////////////////////////////////////////////

      Inc(iMax); inc(indiceEche);
      DebNewEcheReele := PlusDate( DateFinEcheance, 1, 'J' );
      DateDebEcheance := PlusDate (DateDebEcheance,zinterval,TypeCalcul);
      DateDebEcheance := RecaleEcheanceFinMois( DateDebEcheance, DateDebut );
      DateFinEcheance := PlusDate( DateDebEcheance, zinterval, TypeCalcul );
      DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
      //
      if DebNewEcheReele < DateDebEcheance then DateDebEcheance := DebNewEcheReele;
      //
      if (DateFinEcheance > DateFinTraite ) then DateFinEcheance := DateFinTraite;

      if  Not bMtPeriodique then
      // Si le montant de chaque échéance calculé à partir du montant global
      // on rattrape les écarts d'arrondi sur la dernière échéance
      // Sinon, c'est le même montant sur toutes les échéances
      begin
        DateEcheanceSuivante:=PlusDate (DateDebEcheance,zinterval,TypeCalcul);
        DateEcheanceSuivante := RecaleEcheanceFinMois( DateEcheanceSuivante, DateDebut );
        MtResteAFac := MtResteAFac - zmont;
        if ((MtResteAFac<(2*zmont)) and (DateEcheanceSuivante>=DateFinTraite)) then
        zmont := MtResteAFac;
      end
    End;
  end;
end;

Procedure UtilCalculEcheances(MtGlobal:Double;Profil,ModeCalcul,CodeAff,zRepAct, ztier, ztyp_fac : string;
                              TypeModif:T_TypeModifAff; zinterval: integer;zmont: double;
                              DateDebut , DateFin, DateLiquid,DateDebCal,DateFinCal,DateResil : TdateTime;
                              Dev :RDEVISE; SaisieContre, bMtPeriodique, bPourcentage : Boolean;
                              zMultiEche:string; ModeMontantDoc,MethodeCalcul  : string;
                              DateDebutFac,DateFinFac  : TDateTime; TermeEche : string);
Var QQ : Tquery;
    ii, cpt, iMax: Integer;
    TobEch, TobDet : TOB;
    MtDejaFac,MtResteAFac : double;
    bDejaFac:boolean;
    DateDebEcheance, DateFinEcheance, DateEcheanceSuivante, DerniereDateFac, DateFinTraite:TDateTime;
    stWhereDateCal : string;
    chp,zorder : string; //gm
    IndiceDerniereEche, IndiceEcheance:longint;
    AnneeDD,MoisDD,JourDD,AnneeDF,MoisDF,MoisInter,JourDF:word;
    bDateOK, bEcheLiquidExist, bCreerEcheLiquid:boolean;
    Lemontant : double;
    LeChoix : integer;
    year,Month,Day,Yx,Mx,Dx : word;
    yearF,MonthF,DayF,DayAnniv : word;
    NbJourInyear : integer;
    CouvertureperiodeIncomplete : boolean;
    DeltaMois,DeltaJour,NbDaysInPeriod,NbMonthInPeriod,NbDaysInselection : double;
    D1,D2,DateDebutPeriode,DateFinPeriode : TdateTime;
    DateTmp : string;
    TypeEcart : string; // D date de debut differente ; F Date de fin
    indiceEche : integer;
    DebNewEcheReele : TDateTime;
    IndiceEcheMaxDate : TdateTime;
Begin
  if MtGlobal = 0 then exit;

  if MethodeCalcul = 'AN' then
  begin
    DecodeDate(DateFin,year,Month,DayAnniv);
  end;

  bEcheLiquidExist := false;
  IndiceDerniereEche := 0;
  bDejaFac := false;   Cpt := 0;
  MtDejaFac := 0;
  DerniereDateFac := idate1900;
  DateDebutfac := StrToDate(DateToStr(DateDebutfac));
  CouvertureperiodeIncomplete := False;

  if (datedebut=idate1900) then exit; // gm 22/08/02
  if VH_GC.GCIfDefCEGID then
   if (copy(codeaff,1,1)='P') then exit;

  // création Tob contenant toutes les échéances
  TobEch := Tob.Create('Liste Echéances',nil,-1);

  // PA prise en compte des dates de deb et fin de calcul distinctes des dates de gener
  if DateDebCal <> iDate1900 then
   stWhereDateCal := ' AND AFA_DATEECHE >="'+usdatetime(DateDebCal)+'" '
  else
   stWhereDateCal := '';
  if (DateFinFac<>iDate2099) then DateFinTraite:= DateFinFac
                                                     else DateFinTraite:= DateFin;
  // PA prise en compte de la date de résil si antérieure à la date de fin
//  if (DateResil < DateFinTraite) then DateFinTraite := DateResil;

  chp:= 'AFA_DATEECHE,AFA_AFFAIRE,AFA_NUMECHE,AFA_NUMECHEBIS,AFA_TIERS,AFA_ECHEFACT,AFA_LIQUIDATIVE,'+
        'AFA_POURCENTAGE,AFA_MONTANTECHE,AFA_MONTANTECHEDEV,AFA_REPRISEACTIV,AFA_TYPECHE,'+
        'AFA_GENERAUTO,AFA_DEVISE,AFA_LIBELLEECHE,AFA_DATEFINFAC,AFA_DATEDEBUTFAC ';
  zOrder := ' ORDER BY AFA_DATEECHE,AFA_NUMECHE';
  Try
  {$IFNDEF RECALCULAFF}
  QQ := nil;
  Try
    // PL le 06/03/02 : INDEX 3
    //QQ := OpenSQL('SELECT ' + chp+ ' FROM FACTAFF Where AFA_TYPECHE="NOR" AND AFA_AFFAIRE="'+CodeAff+'" And AFA_LIQUIDATIVE="-"'+stWhereDateCal+zorder,True) ; // gm 22/08/02 (zorder)
    // PL le 10/10/02 : on a besoin de l'ech liquidative car il faut reporter les modifs aussi dessus ex : AFA_REPRISEACTIV ou la date
    QQ := OpenSQL('SELECT ' + chp+ ' FROM FACTAFF Where AFA_TYPECHE="NOR" AND AFA_AFFAIRE="'+CodeAff+'"'+stWhereDateCal+zorder,True) ; // gm 22/08/02 (zorder)
    If Not QQ.EOF then TobEch.LoadDetailDB('FACTAFF','','',QQ,True);
  finally
  	Ferme(QQ);
  end;
    {$ENDIF}

  if (TobEch.Detail.Count > 0) Then
  Begin
    for ii:=0 To TobEch.Detail.Count-1 do
    Begin
      TobDet := TobEch.Detail[ii];

      if (TobDet.Getvalue('AFA_LIQUIDATIVE')= '-') then
      begin
          if (TobDet.GetString('AFA_DATEFINFAC')='') or
             (TobDet.GetString('AFA_DATEDEBUTFAC')='') or
             (TobDet.GetDateTime('AFA_DATEFINFAC')=IDate1900) or
             (TobDet.GetDateTime('AFA_DATEDEBUTFAC')=IDate1900) then
          begin
            PgiError ('IMPOSSIBLE : Des dates ne sont pas correctement renseignées.#13#10 Veuillez contrôler vos échéances');
            break;
          end;

          If (TobDet.Getvalue('AFA_ECHEFACT')= 'X') then   // Non facturé
        Begin  // Facturé
          bDejaFac :=true;
          if Not bPourcentage then MtDejaFac:=MtDejaFac+TobDet.GetValue('AFA_MONTANTECHE')
                              else MtDejaFac:=MtDejaFac+TobDet.GetValue('AFA_POURCENTAGE');

            if (IndiceDerniereEche < TobDet.Getvalue('AFA_NUMECHE')) then
            IndiceDerniereEche := TobDet.Getvalue('AFA_NUMECHE');
            //
            if TobDet.GetDateTime('AFA_DATEFINFAC') > DerniereDateFac then
          begin
              DerniereDateFac := TobDet.GetDateTime('AFA_DATEFINFAC');
              DateDebutFac := TobDet.GetDateTime('AFA_DATEDEBUTFAC');
          end;
          End;
      end;

      // PL le 10/10/02 : on gère l'échéance liquidative, mais à part
      If (TobDet.Getvalue('AFA_LIQUIDATIVE')= 'X') then
      begin
        // PL le 13/11/02 : on sait si l'echeance liquidative existe deja
        bEcheLiquidExist := true;
        if (TobDet.Getvalue('AFA_ECHEFACT')= '-') then   // Non facturé
        If (tmaDfacLiq in TypeModif) or (tmaRepAct in TypeModif) Then
        Begin // mcd 13/11/02 revu pour cas activite et date liquid
          GetTOBEcheance(TobDet, TobDet.getValue('AFA_NUMECHE'),Profil, 'X',
                          CodeAff, zRepAct, ztier, ztyp_fac,TobDet.getValue('AFA_MONTANTECHEDEV'),
                          DateLiquid,IDate1900, Dev,
                          SaisieContre,TermeEche,MethodeCalcul,ModeCalcul);
        End;
      end;
    End;   // for
  End;  /// >0

  finally
    TobEch.free;
  end;

  // PL le 13/11/02 : on force la création de l'echeance liquidative si elle n'existait pas
  // et qu'on l'a modifiée
  bCreerEcheLiquid := (Not bEcheLiquidExist) and (tmaDFacLiq in TypeModif);

  // Gestion d'un type ponctuel
  If (ModeCalcul = 'P') And (bDejaFac) then Exit;

  // Modification de date
  // mcd 13/11/02 if (tmaDate in TypeModif) or (tmaDFacLiq in TypeModif) or Not bMtPeriodique Then
  if (tmaDate in TypeModif) or (tmaMntEch in TypeModif) or Not bMtPeriodique Then
   Begin
   // Modif PA le 18/07/2001 pour suppression directe  sans passage par une tob (charge + suppression)
   ExecuteSQL('DELETE FROM FACTAFF Where AFA_AFFAIRE="'+CodeAff+
               '" And AFA_DATEFINFAC >"'+usdatetime(DerniereDateFac)+
               '" And AFA_ECHEFACT="-" AND AFA_TYPECHE="NOR"' + stWhereDateCal);
   if Not bPourcentage then MtResteAFac:= MtGlobal-MtDejaFac
                       else MtResteAFac:= 100-MtDejaFac;
   // création Tob contenant toutes les nouvelles échéances
   TobEch := Tob.Create('Liste Echéances',nil,-1);
   if (bDejaFac) then
   begin
//      DateDebut := GetDateDebutPeriode (ModeCalcul, DateDebutFac, zinterval,MethodeCalcul)  ;
      Datedebut := PlusDate( DerniereDateFac, 1, 'J');
      DateDebutFac := Datedebut;
   end;
    //
    DateDebEcheance := DateDebut;
    DateFinEcheance := PlusDate( DateDebEcheance, zinterval, ModeCalcul );
   DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
    //
   if Not bMtPeriodique and Not bPourcentage then
        zmont := Arrondi(zmont, Dev.Decimale);

   // recadre la date de debut d'écheance sur la période de calcul
    while (DateDebCal <> iDate1900) And (DateDebCal > DateDebEcheance) And (ModeCalcul <>'P') do
      begin
      Inc(Cpt);
      DateDebEcheance := PlusDate( DateDebut, Cpt * zinterval, ModeCalcul );
      DateDebEcheance := RecaleEcheanceFinMois( DateDebEcheance, DateDebut );
      DateFinEcheance := PlusDate( DateDebEcheance, zinterval, ModeCalcul );
       DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
      end;
    
   // recup max numeche
   iMax := 1;
    if (DateDebEcheance <= DateFinTraite) then
      begin
      // PL le 06/03/02 : INDEX 3
      QQ:=OpenSQL('SELECT MAX(AFA_NUMECHE) FROM FACTAFF WHERE AFA_TYPECHE="NOR" AND AFA_AFFAIRE="'+CodeAff+'"',TRUE) ;
      if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1;
      Ferme(QQ);
      end;
   if (DateFinEcheance > DateFinTraite ) then DateFinEcheance := DateFinTraite;
    if (ModeCalcul ='P') then
      begin
      DateDebEcheance := DateDebut;
      DateFinEcheance := PlusDate( DateDebEcheance, zinterval, ModeCalcul );
       DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
      if (DateDebEcheance <= DateFinTraite) and (DateDebEcheance >=DateDebCal) then
      begin
        GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac, zmont,
                       DateDebEcheance, DateFinEcheance , Dev,SaisieContre,TermeEche,MethodeCalcul,ModeCalcul);
      end;
    end else
    begin
      //
      if (DateResil < DateFinEcheance) then DateFinEcheance := DateResil;
      indiceEche := 1;
      while (DateDebEcheance <= DateFinTraite)  do
            Begin
        // calcul des échéances
        DeCodeDate (DateDebEcheance,year,Month,Day); DateDebEcheance := EncodeDateTime (Year,Month,Day,0,0,0,0);
            DeCodeDate (DateDebutFac,year,Month,Day); DateDebutFac := EncodeDateTime (Year,Month,Day,0,0,0,0);
            DeCodeDate (DateFinEcheance,yearF,MonthF,DayF); DateFinEcheance := EncodeDateTime ( YearF,MonthF,DayF,23,59,59,0);
        // calcul de la période prévue
        DateDebutPeriode := DateDebEcheance;
        DateFinPeriode := DateFinEcheance;
        //
        if ModeCalcul = 'NBI' then DateFinEcheance := DateFinTraite;
        if DateDebutFac <= DateFinEcheance then
            begin
          if (indiceEche=1) and (pos(ModeCalcul,'A;M;NBI')>0) then
              begin
            if ModeCalcul = 'A' then
            begin
              	NbDaysInPeriod := DaysInAYear (yearF);
                NbMonthInPeriod := 12;
            end else if ModeCalcul = 'M' then
              begin
              NbDaysInPeriod := CalcDaysInperiod (DateDebutPeriode,DateFinPeriode);
                NbMonthInPeriod := zinterval;
            end else if ModeCalcul = 'NBI' then   //FV1 ==> ???
              begin
              	NbDaysInPeriod := CalcDaysInperiod (DateDebutfac,zinterval);
                NbMonthInPeriod := 1;
              end;
            if MethodeCalcul = 'CIV' then
            begin
              CouvertureperiodeIncomplete := (Arrondi(DaySpan(DateDebutFac,DateFinEcheance),0) <> NbDaysInPeriod)
            end else if MethodeCalcul = 'AN' then
            begin
              CouvertureperiodeIncomplete := (Arrondi(DaySpan(DateDebutFac,DateFinEcheance),0) <> NbDaysInPeriod);
            end;

              DeltaMois := Arrondi(MonthSpan(DateDebutFac,DateFinEcheance),0);
              DeltaJour := Arrondi(DaySpan(DateDebutFac,DateFinEcheance),0);
              D1 := DateDebutfac;
              D2 := DateFinEcheance;
          end else if (pos(ModeCalcul,'A;M;NBI')>0) then
            begin
            if ModeCalcul = 'A' then
              begin
              	NbDaysInPeriod := DaysInAYear (yearF);
                NbMonthInPeriod := 12;
            end else if ModeCalcul = 'M' then
              begin
              NbDaysInPeriod := CalcDaysInperiod (DateDebEcheance,zinterval);
                NbMonthInPeriod := zinterval;
            end else if ModeCalcul = 'NBI' then
              begin
              	NbDaysInPeriod := DaysInAYear (yearF);
                NbMonthInPeriod := 1;
              end;
            if MethodeCalcul = 'CIV' then
            begin
              CouvertureperiodeIncomplete := (Arrondi(DaySpan(DateDebEcheance,DateFinEcheance),0) <> NbDaysInPeriod)
            end else if MethodeCalcul = 'AN' then
            begin
              CouvertureperiodeIncomplete := (Arrondi(DaySpan(DateDebEcheance,DateFinEcheance),0) <> NbDaysInPeriod);
            end;
            DeltaMois := Arrondi(MonthSpan(DateDebEcheance,DateFinEcheance),0);
            DeltaJour := Arrondi(DaySpan(DateDebEcheance,DateFinEcheance),0);
            D1 := DateDebEcheance;
              D2 := DateFinEcheance;
            end;
            //
          if (pos(ModeCalcul,'A;M')>0) and (CouvertureperiodeIncomplete) then
            begin
            if D1 <> DateDebEcheance then TypeEcart := 'D' else TypeEcart := 'F';
            leChoix := GetLeChoixSurProrata (ModeCalcul,TypeEcart,NbMonthInPeriod,D1,D2);
              if (leChoix = 1) then
              begin
              	LeMontant := zMont;
              end else if (leChoix = 2) then
              begin
                  // prorata sur nb de mois concerné
                if TypeEcart = 'D' then
                begin
                	D1 := DEBUTDEMOIS(D1);
                	DeltaJour := Arrondi(DaySpan(D1,D2),0);
                	Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
                end else
                begin
                	D2 := DebutDemois(D2);
                	DeltaJour := Arrondi(DaySpan(D1,D2),0);
                	Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
                end;
              //                  LeMontant := Arrondi(zmont * ((DeltaMois)/NbMonthInPeriod),V_PGI.OkDecV );
              end else if (leChoix = 3) then
              begin
              	// prorata sur le nombre de jour dans l'année concerné ou la période concernée
                Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
              end else
              begin
                // prorata sur nb de mois suivant le mois de début de facturation
                if TypeEcart = 'D' then
                begin
                	D1 := Findemois (D1)+1;
                	DeltaJour := Arrondi(DaySpan(D1,D2),0);
                	Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
                end else
                begin
                	D2 := FinDemois(D2);
                	DeltaJour := Arrondi(DaySpan(D1,D2),0);
                	Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
                end;
              end;

            end else
            begin
            	Lemontant := zMont;
            end;
            //
          //            IndiceEcheance:=GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac, Lemontant, DateEcheance,DateFinEcheance, Dev, SaisieContre);
          IndiceEcheance:=GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac,
                                         Lemontant, D1,D2, Dev, SaisieContre,TermeEche,MethodeCalcul,
                                         ModeCalcul);

            // PL le 18/06/02 : gestion multi echeances
            if (GetParamSoc('SO_AFMULTIECHE')= true) and (zMultiEche='X') then
                // On gère le multi echeance
                if (ztyp_fac<>'ACT') and (ztyp_fac<>'MAN') and (zRepAct<>'NON') then
          begin
                    // En cas de facturation mixte
                    TobEch.Detail[IndiceEcheance].PutValue('AFA_REPRISEACTIV','NON');
                    Inc(iMax);
            GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, 'ACT', 0,
                          DateDebEcheance, DateFinEcheance, Dev, SaisieContre,TermeEche,
                          MethodeCalcul,ModeCalcul);
                    end;
            ///////////////////////////////////////////////

            Inc(iMax); inc(indiceEche);
        end;
        if ModeCalcul = 'NBI' then
            begin
          DateDebEcheance := DateFinTraite;
          DateDebEcheance := PlusDate (DateDebEcheance,zinterval,'J');
        end else
            begin
          DebNewEcheReele := PlusDate( DateFinEcheance, 1, 'J' );
          DateDebEcheance := PlusDate (DateDebEcheance,zinterval,ModeCalcul);
          DateDebEcheance := RecaleEcheanceFinMois( DateDebEcheance, DateDebut );
          DateFinEcheance := PlusDate( DateDebEcheance, zinterval, ModeCalcul );
              DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
          //
          if MethodeCalcul = 'AN' then
          begin
            //FV1 == test si on ne fait cela que dans le cas d'un calcul annuel ce qui semble logique pour une date anniv'...
            if ModeCalcul = 'A' then
            begin
            if DebNewEcheReele <> DateDebEcheance then DateDebEcheance := DebNewEcheReele;
            if (DayOf(DateFinEcheance)<>DayOf(DateFin)) and
               (DayOf(EndOfTheMonth (DateFinEcheance))>=DayOf(DateFin)) then
            begin
              DeCodeDate (DateFinEcheance,year,Month,Day);
              DateFinEcheance := EncodeDate (year,Month,DayOf(DateFin));
            end;
          end;
          end;
            //
        end;
        //
            if (DateFinEcheance > DateFinTraite ) then DateFinEcheance := DateFinTraite;
        if (DateResil < DateFinEcheance) then
        begin
          DateFinEcheance := DateResil;
          DateFinTraite := DateResil;
        end;

            if  Not bMtPeriodique then
        begin
               // Si le montant de chaque échéance calculé à partir du montant global
               // on rattrape les écarts d'arrondi sur la dernière échéance
               // Sinon, c'est le même montant sur toutes les échéances
          DateEcheanceSuivante:=PlusDate (DateDebEcheance,zinterval,ModeCalcul);
               DateEcheanceSuivante := RecaleEcheanceFinMois( DateEcheanceSuivante, DateDebut );
               MtResteAFac := MtResteAFac - zmont;
               if ((MtResteAFac<(2*zmont)) and (DateEcheanceSuivante>=DateFinTraite)) then
                   zmont := MtResteAFac;
               end
            End;
      end;

    if (TobEch.Detail.Count > 0) Then
    begin
      if TobEch.detail[TobEch.detail.count -1].GetDateTime('AFA_DATEFINFAC') < StrToDate(DateToStr(DateFinTraite)) then
      begin
        DateDebEcheance := PlusDate (TobEch.detail[TobEch.detail.count -1].GetDateTime('AFA_DATEDEBUTFAC'),zinterval,ModeCalcul);
        DateDebEcheance := RecaleEcheanceFinMois( DateDebEcheance, DateDebut );
        DateFinEcheance := PlusDate( DateDebEcheance, zinterval, ModeCalcul );
        DateFinEcheance := PlusDate( DateFinEcheance, -1, 'J' );
        NbDaysInPeriod := CalcDaysInperiod (DateDebEcheance,zinterval);
        //
        D1 := PlusDate( TobEch.detail[TobEch.detail.count -1].GetDateTime('AFA_DATEFINFAC'), 1, 'J' );
        D2 := StrToDate(DateToStr(DateFinTraite));
        //
        DeltaJour := Arrondi(DaySpan(D1,D2),0); if deltaJour = 0 then DelTaJour := 1;
        Lemontant := Arrondi(Zmont * (DeltaJour/NbDaysInPeriod),V_PGI.okdecV);
        //
        iMax := TobEch.detail[TobEch.detail.count -1].GetValue('AFA_NUMECHE') +1;
        if LeMontant > 0 then
        begin
          IndiceEcheance:=GetTOBEcheance(TobEch, iMax,Profil, '-', CodeAff, zRepAct,  ztier, ztyp_fac,
                                         Lemontant, D1,D2, Dev, SaisieContre,TermeEche,MethodeCalcul,
                                         ModeCalcul);
        end;
      end;
      TobEch.InsertOrUpdateDB;
    end;
   TobEch.free;

   // On doit forcer la création de l'echeance liquidative dans ce cas
   bCreerEcheLiquid := true;
   IndiceDerniereEche := iMax - 1;
   End;

  if (GetParamSoc('SO_AFGERELIQUIDE') = true) and bCreerEcheLiquid then
  begin
    TobEch := Tob.Create('Echéance liquidative',nil,-1);
    try
      if (DateLiquid<>iDate1900) and (DateLiquid<>iDate2099) then
      GetTOBEcheance(TobEch, IndiceDerniereEche + 1,Profil, 'X', CodeAff, zRepAct, ztier, ztyp_fac, 0, DateLiquid, IDate1900, Dev, SaisieContre,TermeEche,MethodeCalcul,ModeCalcul);
    finally
      if (TobEch.Detail.Count > 0) Then TobEch.InsertOrUpdateDB;
      TobEch.free;
    End;
  end;

End;
{***********A.G.L.***********************************************
Auteur  ...... : P.LENORMAND

Créé le ...... : 19/06/2002

Modifié le ... :   /  /

Description .. : Calcul des échéances de facturation de l'affaire par rapport au budget

Mots clefs ... : CALCUL ECHEANCE AFFAIRE BUDGET

*****************************************************************}

Procedure UtilCalculEcheancesBudget(Profil, CodeAff, zRepAct, zGenerAuto : string;
                                Dev :RDEVISE; SaisieContre : Boolean; zMultiEche:string);
var
TobBudget, TobEchBud, TobFactAff:TOB;
QQ:TQuery;
IndiceEcheance,ii,iMax,ib:integer;
bMultiEcheOK,bECAExiste:boolean;
Begin
// création Tob contenant toutes les échéances du budget
TobBudget := Tob.Create('Echéances budget',nil,-1);
// création Tob recevant les nouvelles echeances dans factaff
TobFactAff := Tob.Create('Echéances factaff',nil,-1);
Try
Try
QQ := nil;
Try
QQ := OpenSQL('SELECT ABU_TYPEAFBUDGET,ABU_AFFAIRE,ABU_DATEBUD,SUM(ABU_MTPVBUDDEV) AS ABU_MTPVBUDDEV,ABU_TIERS,ABU_ALIMBUDECH FROM AFBUDGET Where ABU_AFFAIRE="'+CodeAff+'" AND ABU_TYPEAFBUDGET="PVT" GROUP BY ABU_TYPEAFBUDGET,ABU_AFFAIRE,ABU_DATEBUD,ABU_TIERS,ABU_ALIMBUDECH ORDER BY ABU_AFFAIRE,ABU_TIERS,ABU_DATEBUD,ABU_ALIMBUDECH',True,-1,'',true) ;
If Not QQ.EOF then TobBudget.LoadDetailDB('AFBUDGET','','',QQ,True);
finally
Ferme(QQ);
end;

if V_PGI.IoError=oeOk then
begin
if (TobBudget.Detail.Count > 0) Then
    Begin
    // Suppression des anciennes lignes de FactAff  pour cette affaire
    ExecuteSQL('DELETE FROM FACTAFF WHERE AFA_AFFAIRE="'+CodeAff+'" AND AFA_ECHEFACT="-" AND AFA_TYPECHE="NOR"');

    // premier indice à 1
    iMax:=1;
    for ii:=0 To TobBudget.Detail.Count-1 do
        Begin
        TobEchBud := TobBudget.Detail[ii];

        IndiceEcheance:=-1;
        if (TobEchBud.GetValue('ABU_ALIMBUDECH')='EC1') then
            IndiceEcheance:=GetTOBEcheance(TobFactAff, iMax,Profil, '-', CodeAff, zRepAct,  TobEchBud.GetValue('ABU_TIERS'), zGenerAuto, TobEchBud.GetValue('ABU_MTPVBUDDEV'), TobEchBud.GetValue('ABU_DATEBUD'), IDate1900, Dev, SaisieContre,'','CIV','M');

        bMultiEcheOK := (GetParamSoc('SO_AFMULTIECHE') = true) and (zMultiEche = 'X') and (zGenerAuto <> 'ACT') and (zGenerAuto <> 'MAN') and (zRepAct <> 'NON');

        if (TobEchBud.GetValue('ABU_ALIMBUDECH')='EC1') and bMultiEcheOK then
            begin
            if (IndiceEcheance<>-1) then
                    TobFactAff.Detail[IndiceEcheance].PutValue('AFA_REPRISEACTIV','NON');
                    
            ib:=ii+1;
            bECAExiste:=false;
            while (ib <= TobBudget.Detail.Count-1) and (bECAExiste=false) and
                (TobEchBud.GetValue('ABU_AFFAIRE')=TobBudget.Detail[ib].GetValue('ABU_AFFAIRE')) and
                (TobEchBud.GetValue('ABU_DATEBUD')=TobBudget.Detail[ib].GetValue('ABU_DATEBUD')) and
                (TobEchBud.GetValue('ABU_TIERS')=TobBudget.Detail[ib].GetValue('ABU_TIERS'))
                do
                begin
                if (TobBudget.Detail[ib].GetValue('ABU_ALIMBUDECH')='ECA') then bECAExiste:= true;
                Inc(ib);
                end;

          if Not bECAExiste then
                begin
                Inc(iMax);

                GetTOBEcheance(TobFactAff, iMax,Profil, '-', CodeAff, zRepAct,  TobEchBud.GetValue('ABU_TIERS'), 'ACT', 0, TobEchBud.GetValue('ABU_DATEBUD'), IDate1900, Dev, SaisieContre,'','CIV','M');
                end;
            end;

        if (TobEchBud.GetValue('ABU_ALIMBUDECH')='ECA') and bMultiEcheOK then
                GetTOBEcheance(TobFactAff, iMax, Profil,'-', CodeAff, zRepAct,  TobEchBud.GetValue('ABU_TIERS'), 'ACT', TobEchBud.GetValue('ABU_MTPVBUDDEV'), TobEchBud.GetValue('ABU_DATEBUD'),IDate1900, Dev, SaisieContre,'','CIV','M');

        Inc(iMax);
        end;

    if (TobFactAff.Detail.Count > 0) Then
        TobFactAff.InsertOrUpdateDB;

    end;
end;

finally
TobFactAff.free;
TobBudget.Free;
end;

except
V_PGI.IoError:=oeSaisie;    
end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 17/10/2001
Modifié le ... :   /  /
Description .. : Calcule la date de fin de mois par rapport à une date de
Suite ........ : début d'échéance en tenant compte que la date de début
Suite ........ : d'échéance peut-être le dernier jour du mois, l'avant dernier
Suite ........ : ou l'antepénultième (seuls cas qui posent problème).
Suite ........ : entrees :
Suite ........ : AcdDateEntree : date de l'échéance calculee brute
Suite ........ : AcdDateDebEch : date de la premiere echeance
Suite ........ :
Suite ........ : ex :
Suite ........ : calcul bête des échéances (la 1ere date est la 1ere ech) :
Suite ........ : 31/01/01, 28/02/01, 28/03/01, 28/04/01....
Suite ........ : calcul recalé en passant les dates dans la fonction :
Suite ........ : 31/01/01, 28/02/01, 31/03/01, 30/04/01...
Suite ........ :
Suite ........ : mais même mieux :
Suite ........ : calcul bête en ne tenant compte que des fins de mois :
Suite ........ : 30/01/01, 28/02/01, 28/03/01.... le 30 n'étant pas la fin de janvier
Suite ........ : alors qu'il faut
Suite ........ : 30/01/01, 28/02/01, 30/03/01, 30/04/01...
Suite ........ :
Suite ........ : encore plus fort :
Suite ........ : si la 1ere echeance est en février, il ne faut gérer que le cas
Suite ........ : du dernier jour du mois (soit 28, soit 29)... à méditer...
Suite ........ :
Suite ........ : Petite subtilité : les échéances peuvent être espacées de n mois !
Suite ........ :
Mots clefs ... : ECHEANCES; DATE;
*****************************************************************}
function RecaleEcheanceFinMois( AcdDateEntree, AcdDateDebEch : TDateTime ) : TDateTime;
var jourDeb, jourEch,  mois, annee : word;
    iNbJFinMois : integer;
begin
  Result := AcdDateEntree;
  iNbJFinMois := trunc(FinDeMois(AcdDateDebEch) - AcdDateDebEch);

  if iNbJFinMois > 2 then exit;
  if (AcdDateEntree = idate1900) or (AcdDateEntree = idate2099) then exit;

  // Gestion 1ere echeance en février
  DecodeDate(AcdDateDebEch, annee, mois, jourDeb);
  // Si la 1ere echeance est en février et que ce n'est pas la fin du mois (28 ou 29)
  // on conserve la date d'entree pour les echeances suivantes
  if (mois=2) and (iNbJFinMois<>0) then exit;


  DecodeDate(AcdDateEntree, annee, mois, jourEch);
  case iNbJFinMois of
    0 : Result := FinDeMois(AcdDateEntree);
    1,2 : if (mois=2) then Result := FinDeMois(AcdDateEntree)
            else Result := EncodeDate(annee, mois, jourDeb);
  end;

end;

function ConstitueLibEcheContrat ( TOBEche : TOB ) : string;
var st : string;
  Posit : Integer;
begin
  St := VH_GC.AFFLibfactAff;
  Posit := Pos ('**',St);            // $$  pour reprendre le numéro d'echeance
  If (Posit <> 0) then
  Begin
    delete (St,Posit,2);
    Insert(TOBEche.Getvalue('AFA_NUMECHEBIS'),St,Posit);
  end;   //mcd 23/05/03 mise bis
  Posit := Pos ('$$',St);            // $$  pour reprendre la date
  If (Posit <> 0) then
  Begin
    delete (St,Posit,2);
    if TOBEche.GetValue('AFA_DATEFINFAC') <> IDate1900 then
    begin
      Insert(' '+DateToStr(TOBEche.GetValue('AFA_DATEDEBUTFAC'))+' au '+DateToStr(TOBEche.GetValue('AFA_DATEFINFAC')),St,Posit);
    end else
    begin
      Insert(DateToStr(TOBEche.GetValue('AFA_DATEECHE')),St,Posit);
    end;
  end;
  Result := st;
end;

function GetTOBEcheance (TobEch:TOB; Rang:integer; Profil,sLiquidative, CodeAff,  zRepAct, ztier, ztyp_fac : string;
             zmont: double; DateEch,DateFinEche : TdateTime; Dev :RDEVISE; SaisieContre : Boolean; TermeEche,MethodeCalcul,typeCalcul : string):longint;
var  TOBDet:TOB;
  st : string;
  Posit : integer;
  Y,M,D : word;
  Dcivile : TdateTime;
begin
  DecodeDate (DateFinEche,Y,M,D); DateFinEche := EncodeDate (Y,M,D);
  TobDet := Tob.Create('FACTAFF',TobEch,-1);
  TobDet.PutValue('AFA_TIERS',ztier);
  TobDet.PutValue('AFA_REPRISEACTIV',zRepAct);
  TobDet.Putvalue('AFA_AFFAIRE',CodeAff);
  TobDet.Putvalue('AFA_TYPECHE','NOR');
  TobDet.Putvalue('AFA_NUMECHE',Rang);
  TobDet.Putvalue('AFA_NUMECHEBIS',Rang);   //mcd 15/05/03
  TobDet.Putvalue('AFA_PROFILGENER',Profil);   //mcd 03/07/03
  TobDet.Putvalue('AFA_GENERAUTO',ztyp_fac);
  If ((ztyp_fac = 'POU') or (ztyp_fac = 'POT')) Then
   Begin
   TobDet.PutValue('AFA_POURCENTAGE',zmont);
   TobDet.PutValue('AFA_MONTANTECHE',0);
   TobDet.PutValue('AFA_MONTANTECHEDEV',0);
  End Else
   Begin
   TobDet.PutValue('AFA_DEVISE',Dev.Code);
   // calcul montant pivot + contrevaleur / devise
   CalculMtEch (TobDet,DEV,zmont, SaisieContre);
   TobDet.PutValue('AFA_POURCENTAGE',0);
   End;
  if TermeEche = 'ECH' then
  begin
	TobDet.PutValue('AFA_DATEECHE',DateFinEche);
  end else
  begin
	if MethodeCalcul = 'CIV' then
  begin
  	DecodeDate (DateEch,Y,M,D);
		if typeCalcul = 'A' then DCivile := EncodeDate (Y,1,1)
    										else DCivile := EncodeDate (Y,M,1);
      //FV1 : 19/03/2018 -  FS#3010 - ACTUACOM - Date d'échéance incorrecte pour la première échéance
      //Si la date d'échéance est différente de Dcivile on recharge avec dateEch.
      if DateEch > DCivile then
        TobDet.PutValue('AFA_DATEECHE',DateEch)
      else
        TobDet.PutValue('AFA_DATEECHE',DCivile);
      //
    end
  else
    begin
      TobDet.PutValue('AFA_DATEECHE',DateEch);
    end;
  end;

  if (sLiquidative='X') then
    begin
    St := GetParamSoc('SO_AFLIBLIQFACTAF');
    Posit := Pos ('**',St);            // $$  pour reprendre le numéro d'echeance
    If (Posit <> 0) then
    Begin
      delete (St,Posit,2);
      Insert(TobDet.Getvalue('AFA_NUMECHEBIS'),St,Posit);
    end;     //mcd 25/05/03 ajout Bis
    Posit := Pos ('$$',St);            // $$  pour reprendre la date
    If (Posit <> 0) then Begin delete (St,Posit,2);Insert(DateToStr(DateEch),St,Posit); end;
  end else
    begin
    St := VH_GC.AFFLibfactAff;
    Posit := Pos ('**',St);            // $$  pour reprendre le numéro d'echeance
    If (Posit <> 0) then
    Begin
      delete (St,Posit,2);
      Insert(TobDet.Getvalue('AFA_NUMECHEBIS'),St,Posit);
    end;   //mcd 23/05/03 mise bis
    Posit := Pos ('$$',St);            // $$  pour reprendre la date
    If (Posit <> 0) then
    Begin
    	delete (St,Posit,2);
      if DateFinEche <> IDate1900 then
      begin
      	Insert(' '+DateToStr(DateEch)+' au '+DateToStr(DateFinEche),St,Posit);
      end else
      begin
      	Insert(DateToStr(DateEch),St,Posit);
      end;
    end;
    end;
  //
  TobDet.PutValue('AFA_DATEDEBUTFAC',DateEch);
  TobDet.PutValue('AFA_DATEFINFAC',DateFineche);
  //
  TobDet.PutValue('AFA_LIBELLEECHE',St);
  TobDet.PutValue('AFA_ECHEFACT','-');
  TobDet.PutValue('AFA_LIQUIDATIVE', sLiquidative);
  Result:= TobDet.GetIndex;
end;

function GetLeChoixSurProrata (TypeCalcul,TypeEcart : string;NbMoisInPeriod : double;D1,D2 : TDateTime) : integer;
var TOBEchange : TOB;
		Mois,MoisSuiv : string;
    year,Month,Day : word;
begin
  if TypeEcart = 'D' then DecodeDate (D1,Year,Month,Day)
  									 else DecodeDate (D2,Year,Month,Day);

	Mois := LongMonthNames[Month]; if Month = 12 then moisSuiv := LongMonthNames[1] else moisSuiv := LongMonthNames[month+1];

	TOBEchange := TOB.Create('ECHANGE INFOS',nil,-1);
  TOBEchange.AddChampSupValeur ('PERIODE','Echéance du '+DateToStr(D1)+' au '+DateToStr(D2));
  TOBEchange.AddChampSupValeur ('TEXTE1','La période de couverture est inférieure');
  if TypeCalcul = 'A' then TOBEchange.AddChampSupValeur ('TEXTE2','à une année compléte')
  											 else TOBEchange.AddChampSupValeur ('TEXTE2','à la période de '+FloatToStr(NbMoisInPeriod)+' mois');
  //
  TOBEchange.AddChampSupValeur ('CHOIX1','Prendre en compte le montant global pour la période du '+DateToStr(D1)+' au '+DateToStr(D2));
  if TypeEcart = 'D' then
  begin
  	if typeCalcul = 'A' then
    begin
      TOBEchange.AddChampSupValeur ('CHOIX2','Calculer un prorata à partir du début du mois de '+Mois);
      TOBEchange.AddChampSupValeur ('CHOIX3','Calculer un prorata à partir du '+IntToStr(Day)+' du mois de '+Mois);
      TOBEchange.AddChampSupValeur ('CHOIX4','Calculer un prorata à partir du début du mois de '+MoisSuiv);
    end else
    begin
      TOBEchange.AddChampSupValeur ('CHOIX2','');
      TOBEchange.AddChampSupValeur ('CHOIX3','Calculer un prorata à partir du '+IntToStr(Day)+' du mois de '+Mois);
      TOBEchange.AddChampSupValeur ('CHOIX4','');
    end;
  end else
  begin
  	if TypeCalcul = 'A' then
    begin
      TOBEchange.AddChampSupValeur ('CHOIX2','Calculer un prorata jusqu''au début du mois de '+Mois);
      TOBEchange.AddChampSupValeur ('CHOIX3','Calculer un prorata jusqu''au '+IntToStr(Day)+' du mois de '+Mois);
      TOBEchange.AddChampSupValeur ('CHOIX4','Calculer un prorata jusqu''à la fin du mois de '+Mois);
    end else
    begin
      TOBEchange.AddChampSupValeur ('CHOIX2','');
      TOBEchange.AddChampSupValeur ('CHOIX3','Calculer un prorata jusqu''au '+IntToStr(Day)+' du mois de '+Mois);
      TOBEchange.AddChampSupValeur ('CHOIX4','');
    end;
  end;
  TOBEchange.AddChampSupValeur ('NBCHOIX',4);
  TOBEchange.AddChampSupValeur ('RETOUR',1);
  TheTOB := TOBEchange;
  AglLanceFiche ('BTP','BTDEMANDECHOIX','','','');
  TheTob := nil;
  result := TOBEchange.getValue('RETOUR');
  TOBEchange.free;
end;

function LanceRecalcEches (TOBAffaire : TOB; DEV : RDevise) : boolean;
var TypeMotif : T_TypeModifAff;
		zMont : double;
    DateDebut,DateFinFac,DateLiquid,DateResil : TDateTime;
    bMtPeriodique, bPourcentage : Boolean;
    ModeMontantDoc: String;
    MethodeCalcul : string;
begin
	TypeMotif := [tmaMntEch,tmaDate];
  bMtPeriodique := true;
  bPourcentage := ((TOBAffaire.GetValue ('AFF_GENERAUTO') = 'POU') or (TOBAffaire.GetValue('AFF_GENERAUTO') = 'POT') ) ;
  //
  if (bPourcentage) then
    zmont := double (TOBAffaire.GetValue('AFF_POURCENTAGE') )
  else
    zmont := double (TOBAffaire.GetValue('AFF_MONTANTECHEDEV') ) ;
  //
  DateDebut := TOBAffaire.GETVALUE('AFF_DATEDEBGENER');
  DateFinFac := TOBAffaire.GETVALUE('AFF_DATEFINGENER');
  DateLiquid := TOBAffaire.GETVALUE('AFF_DATEFACTLIQUID');
  DateResil  := TOBAffaire.GetValue('AFF_DATERESIL');
  ModeMontantDoc := TOBAffaire.getValue('AFF_DETECHEANCE');
	MethodeCalcul := TOBAffaire.getValue('AFF_METHECHEANCE');

  //
  UtilCalculEcheances (TOBAffaire.GetValue('AFF_TOTALHTGLODEV'),
                       TOBAffaire.GetValue('AFF_PROFILGENER'),
                       TOBAffaire.GetValue('AFF_PERIODICITE'),
                       TOBAffaire.GetValue('AFF_AFFAIRE'),
                       TOBAffaire.GetValue('AFF_REPRISEACTIV'),
                       TOBAffaire.GetValue('AFF_TIERS'),
                       TOBAffaire.GetValue('AFF_GENERAUTO'),
                       TypeMotif,
                       Integer (TOBAffaire.GetValue('AFF_INTERVALGENER') ),
                       zmont,
                       DateDebut,
                       DateFinFac,
                       DateLiquid,
                       DateDebut,
                       idate2099,
                       DateResil,
			           			 DEV,
                       false,
                       bMtPeriodique,
                       bPourcentage,
                       TOBAffaire.GetValue('AFF_MULTIECHE'),
                       ModeMontantDoc,
                       MethodeCalcul,
                       DateDebut,DateFinFac,
                       TOBAffaire.getValue('AFF_TERMEECHEANCE') ) ;

end;

end.
