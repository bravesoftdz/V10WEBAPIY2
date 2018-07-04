{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2003
Modifié le ... :   /  /
Description .. : - FQ 12109 - CA - 03/06/2003 - Si dernier jour du mois ==>
Suite ........ : jour = 30 ( correction pb 28/02 )
Suite ........ : BTY 02/06 Anciennes édition ouvertes à Web Access (pour la TVA)
Suite......... : - FQ 17423 - MBO - 23/02/2006 - gettaux : on borne le taux à 100 maxi
Suite......... : BTY 03/06 FQ 17547, 17583 Pb en éditions Web Access ((éditions ne sortent pas
Suite......... : ou certains montants ne sortent pas ou sont faux) à cause des dates lues
Suite......... : dans VHIMMO au lieu de PlanInfo.Exercices
Suite......... : BTY 03/06 Optimisation en Web Access par usage de la class d'édition TAmortissementContext
Suite......... : - FQ 15287 - MBO - 31/03/2006 - calcul du linéaire si exercice < 1 mois
Suite......... : MVG 19/02/2007 NOMBREJOUR360 reconduction du fonctionnement en delphi 5
Suite......... : - FQ 20290 - MBO 30/05/2007 : annulation de la fq 17423 on ne borne plus le taux 
Suite......... : MBO - 28/09/2007 - pour fq 21754  suite nveau paramètre pour amortir jour de cession ou non :
                                    2 nouvelles fonctions : GetDotationLinSortie et GetProrataLinSortie
Mots clefs ... :
*****************************************************************}
unit ImOuPlan;

interface

uses
  SysUtils
  , HEnt1
  , ImEnt
  , HCtrls
  , Math
  , UTOB
  {$IFDEF SERIE1}
  , S1util  //YCP 25/08/05
  {$ELSE}
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ELSE}
  , Ent1
  {$ENDIF}
  {$ENDIF}
  {$IFDEF eAGLClient}
  {$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  {$ENDIF eAGLClient}
  // BTY 03/06 ,ImPlanInfo   // BTY 02/06
  ;


function GetTaux(Methode: string; DateAmort,DateAchat: TDateTime; NbMois: integer):double;
function GetTauxLineaire(DateAchat: TDateTime; NbMois: integer):double;
function GetTauxDegressif(DateAchat: TDateTime; NbMois: integer):double;
function GetProrataLinAvant(DateDebutPer, DateFinPer,DateDeb,DateFinAm ,DateCalcul : TDateTime; bDotationExercice : boolean) : double;
function GetProrataLinApres(DateDebutPer, DateFinPer, DateCalcul : TDateTime; bDotationExercice : boolean) : double;
function GetProrataDegApres(DateDebutPer,DateFinPer,DateCalcul :TDateTime) : double;
function GetProrataDegAvant(DateDebutPer,DateFinPer,DateDebAmort,DateCalcul :TDateTime) : double;
procedure GetDatesExercice(DateCalcul : TDateTime; var DateDebut,DateFin : TDateTime);
// procedure GetDateDebutFinEx(Q : TQuery; var DateDebut,DateFin : TDateTime);
procedure GetDateDebutFinEx(Exo : TImExoDate; var DateDebut,DateFin : TDateTime;LongExo:integer=12); // Modif CA le 15/03/99
procedure CalculDateDebutFinEx(DateCalcul : TDateTime;var DateDebut,DateFin : TDateTime;LongExo:integer=12);
procedure DecrementeExe(DateCalcul,DateRef : TDateTime; var DateDebut,DateFin : TDateTime; LongExo:integer=12);
procedure IncrementExe(DateCalcul,DateRef : TDateTime; var DateDebut,DateFin : TDateTime; LongExo:integer=12);
function GetNombrePeriodeMois(Base,Diviseur : integer) : integer;
 function GetNombrePeriodeDate(DateRef1,DateRef2 : TDateTime) : integer ;
function GetDotationLinAvant(DateDebut,DateFin,DateDebAmort,DateCalcul,DateFinAmort : TDateTime; Annuite : double; bDotationExercice : boolean): double;
function GetDotationLinApres(DateDebut,DateFin,DateCalcul,DateFinAmort : TDateTime; Annuite : double; bDotationExercice : boolean): double;
function GetDotationDegAvant(DateDebut,DateFin,DateDebAmort,DateCalcul : TDateTime; Annuite : double; var Prorata : double;AvecSurAmort : boolean): double;
function GetDotationDegApres(DateDebut,DateFin,DateCalcul : TDateTime; Annuite : double; var Prorata : double): double;
function ExerciceEnCours(DateRef : TDateTime) : boolean;
Procedure TraiteParamDateFinAmort(DateDebAmort,DateFinAmort : TDateTime; var DateDebut,DateFin,DateCalcul : TDateTime);
function GetDateDebutExoRef(DateOp : TDateTime) : TDateTime;
procedure CalculDateDebutFinExByPos(iPosEx : integer;var DateDebut,DateFin : TDateTime);//EPZ 03/11/00

// BTY 03/06 Optimisation
// BTY 02/06 Anciennes éditions à ouvrir à CWAS
{$IFDEF EAGLSERVER}
 //function GetTauxTVAOrganisme (PlanInfo:TPlanInfo; stOrganisme, stGeneTVA : string) : double;
 //function CalculEchTva(PlanInfo:TPlanInfo; Q : TQuery; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
 //procedure GetCoeffTVARable(PlanInfo:TPlanInfo; Q : TQuery; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
 //procedure CalculDateDebutFinExByPosWA(PlanInfo: TPlanInfo; iPosEx : integer;var DateDebut,DateFin : TDateTime; DebutExo :TDateTime);//EPZ 03/11/00
 //function TauxTVAServer (PlanInfo:TPlanInfo; ModeTVA,Tva : String; Achat:Boolean): double;
function TauxTVAServer (ModeTVA,Tva : String; Achat:Boolean): double;
{$ELSE}
 //function GetTauxTVAOrganisme (stOrganisme, stGeneTVA : string) : double;
 //function CalculEchTva(Q : TQuery; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
 //procedure GetCoeffTVARable(Q : TQuery; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
function CalculEchTvaTOB(OB : TOB; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
function AffecteTvaEch(Q : TQuery; laTVADefaut : string;MontantHT : double) : double;
procedure GetCoeffTVARableTOB(OB : TOB; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
{$ENDIF}

function CalculEchTva(Q : TQuery; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
procedure GetCoeffTVARable(Q : TQuery; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
function GetTauxTVAOrganisme (stOrganisme, stGeneTVA : string) : double;

function CalculSurAmortissement ( Dotation : double; DateAchat, DateDebEx,DateFinEx : TDateTime; bIntermediaire : boolean = FALSE; DateFin : TDateTime = 2) : double;
function NOMBREJOUR360 ( DateDebut, DateFin : TDateTime ) : integer;

function AMTaux2Duree (Methode: string; DateAchat: TDateTime; Taux : double ) : integer;
function AMTaux2DureeLineaire( Taux : double ) : integer;
function AMTaux2DureeDegressif ( Taux : double ; DateAchat : TDateTime ) : integer;


//new mbo 28.09.07 - fq 21754
function GetProrataLinSortie(DateDebutPer, DateFinPer,DateDeb,DateFinAm ,DateCalcul : TDateTime; bDotationExercice : boolean) : double;
function GetDotationLinSortie(DateDebut,DateFin,DateDebAmort,DateCalcul,DateFinAmort : TDateTime; Annuite : double; bDotationExercice : boolean): double;




implementation

// BTY 03/06 Pour récupérer le contexte d'édition
// BTY 03/06 Contexte d'édition remonté dans ImEnt
//{$IFDEF eAGLServer}
//uses AmedServer;
//{$ENDIF}


function GetTaux(Methode: string; DateAmort,DateAchat: TDateTime; NbMois: integer):double;
var
	Taux : double;
begin
  Taux:=0.00;
  if NbMois>0 then
  begin
    if Methode = 'LIN' then Taux:=GetTauxLineaire(DateAmort, NbMois)
    else if Methode = 'DEG' then Taux:=GetTauxDegressif(DateAchat, NbMois);
  end;
  // fq 17423 - mbo - on borne le taux à maxi 100 %
  { fq 20290 - mbo - retour arrière sur la fq 17423 on ne borne plus le taux car pb de calcul
  if Taux > 100.00 then
     Taux := 100.00;  }

  result := Taux;
end;

function GetTauxLineaire(DateAchat: TDateTime; NbMois: integer):double;
begin
     result := (100/(NbMois/12));
end;

function GetTauxDegressif(DateAchat: TDateTime; NbMois: integer):double;
var Coeff,Taux,Coef1,Coef2,Coef3 : double;
  Annee : double; DateApplicationNouveauCoef : TDateTime;
  T1: tob;   i: integer ; St,wSt: string ;  btrouve: boolean ;
begin
  Taux:=(100/(NbMois/12));
  Coeff:=1; Coef1:=1 ; Coef2:=1; Coef3:=1;
  DateApplicationNouveauCoef:=idate1900 ;  bTrouve:=false ;
  Annee:=NbMois / 12;
  T1:=TOB.Create('UNKOWN',nil,-1) ;
  T1.LoadDetailDB('CHOIXCOD','"ICD"','CC_CODE',nil,false) ;
  for i:=0 to T1.Detail.Count-1 do
    begin
    St:=T1.Detail[i].GetValue('CC_LIBELLE') ;
    wSt:=READTOKENST(St) ; if IsValidDate(wSt) then DateApplicationNouveauCoef:=StrToDate(wSt) else continue ;
    wSt:=READTOKENST(St) ; if IsNumeric(wSt)   then Coef1:=VALEUR(wSt) else continue ;
    wSt:=READTOKENST(St) ; if IsNumeric(wSt)   then Coef2:=VALEUR(wSt) else continue ;
    wSt:=READTOKENST(St) ; if IsNumeric(wSt)   then Coef3:=VALEUR(wSt) else continue ;
    //if (DateApplicationNouveauCoef<>iDate1900) and (DateAchat >= DateApplicationNouveauCoef) then
    if (DateAchat >= DateApplicationNouveauCoef) then
      begin
      if (Annee>=3) and (Annee<5)       then Coeff:=coef1   // 3 et 4 ans
      else if (Annee>=5) and (Annee<=6) then Coeff:=coef2   // 5 et 6 ans
      else if Annee>6                   then Coeff:=coef3 ; // > 6 ans
      btrouve:=true ;
      end
    end ;
  T1.free ;

  if not btrouve then
    begin
    //Cas gérés par l'application
    if (Annee>=3) and (Annee<5)       then Coeff := 1.5  // 3 et 4 ans
    else if (Annee>=5) and (Annee<=6) then Coeff := 2    // 5 et 6 ans
    else if Annee>6                   then Coeff := 2.5; // > 6 ans
    if (NbMois>12) and (DateAchat>=StrToDate('01/02/1996')) AND (DateAchat<=StrToDate('31/01/1997')) then
      Coeff:=Coeff+1 ;
    end ;
  result:=Taux*Coeff;
end;

function GetDotationLinApres(DateDebut,DateFin,DateCalcul,DateFinAmort : TDateTime; Annuite : double; bDotationExercice : boolean): double;
var
  Prorata : double;
begin
  if DateFinAmort<>iDate1900 then TraiteParamDateFinAmort(iDate1900,DateFinAmort,DateDebut,DateFin,DateCalcul); // fin amort ou cessio
  Prorata:=GetProrataLinApres(DateDebut,DateFin,DateCalcul,bDotationExercice);
  result:=Arrondi (Annuite*Prorata,V_PGI.OkDecV);
end;

function GetDotationLinAvant(DateDebut,DateFin,DateDebAmort,DateCalcul,DateFinAmort : TDateTime; Annuite : double; bDotationExercice : boolean): double;
var
  Prorata : double;
begin
  result:=0.00;
  if DateCalcul<DateDebAmort then exit;
  if DateFinAmort<>iDate1900 then TraiteParamDateFinAmort(DateDebAmort,DateFinAmort,DateDebut,DateFin,DateCalcul); // fin amort ou cessio
  Prorata := GetProrataLinAvant(DateDebut,DateFin,DateDebAmort,DateFinAmort,DateCalcul,bDotationExercice);
  result := Arrondi (Annuite*Prorata,V_PGI.OkDecV);
end;

function GetDotationDegApres(DateDebut,DateFin,DateCalcul :TDateTime; Annuite : double; var Prorata : double): double;
begin
  Prorata := GetProrataDegApres(DateDebut,DateFin,DateCalcul);
  result := Arrondi (Annuite*Prorata,V_PGI.OkDecV);
end;

function GetDotationDegAvant(DateDebut,DateFin,DateDebAmort,DateCalcul :TDateTime; Annuite : double; var Prorata : double; AvecSurAmort : boolean): double;
begin
  result:=0.00;
  if DateCalcul<DateDebAmort then exit;
  Prorata := GetProrataDegAvant(DateDebut,DateFin,DateDebAmort,DateCalcul);
  if AvecSurAmort and (Prorata < 1) and (DateDebAmort < DateDebut) then  // Calcul dotation intermédiaire avec suramortissement
    result :=  Arrondi (CalculSurAmortissement(Annuite,DateDebAmort,DateDebut,DateCalcul,True,DateFin),V_PGI.OkDecV)
  else result := Arrondi (Annuite*Prorata,V_PGI.OkDecV);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 01/04/2004
Modifié le ... : 01/04/2004
Description .. : Calcul du prorata en dégressif pour une période allant
Suite ........ : jusqu'à la fin de l'exercice
Suite ........ :
Suite ........ : Prorata =
Suite ........ : (NbJourExercice/360)*(NbMoisDetention/NbMoisExercice)
Mots clefs ... :
*****************************************************************}
function GetProrataDegApres(DateDebutPer,DateFinPer,DateCalcul :TDateTime) : double;
var
  PremMois,PremAnnee,NbMoisExercice,NbMoisDetention : Word;
begin
  Result := 1;
  if DateDebutPer=DateFinPer then exit;
  NOMBREMOIS (DateDebutPer,DateFinPer,PremMois,PremAnnee,NbMoisExercice);
  NOMBREMOIS (DateCalcul,DateFinPer,PremMois,PremAnnee,NbMoisDetention);
  if (NbMoisExercice <> 0) then
    result := (NOMBREJOUR360 ( DateDebutPer,DateFinPer )/360)*(NbMoisDetention/NbMoisExercice)
  else
  // CA - 22/04/2002 - Cas des exercice sur 1 mois.
  if (NbMoisExercice=0) then  result :=1;
end;

function GetProrataDegAvant(DateDebutPer,DateFinPer,DateDebAmort,DateCalcul :TDateTime) : double;
var
  PremMois,PremAnnee,NbMois : Word;
  NbMoisPeriode, NbMoisDetention : integer;
  tDateRef : TDateTime;
begin
  Result := 1;
  if DateDebutPer=DateFinPer then exit;
//  if (DateCalcul=DateFinPer) and (DateDebutPer=DateDebAmort) then exit;
// CA - 28/08/2002 : deuxième partie ne semble pas utile ...
  if (DateCalcul=DateFinPer) then exit;
  if (DateDebAmort>DateDebutPer) and (DateDebAmort<DateFinPer) then
  begin
    if (DateCalcul=DateFinPer) then exit;
    NOMBREMOIS (DateDebAmort,DateFinPer,PremMois,PremAnnee,NbMois);
  end
  else NOMBREMOIS (DateDebutPer,DateFinPer,PremMois,PremAnnee,NbMois);
  NbMoisPeriode :=NbMois;
  if DateDebutPer > DateDebAmort then tDateRef:=DateDebutPer else tDateRef:=DateDebAmort;
  NOMBREMOIS (tDateRef,DateCalcul,PremMois,PremAnnee,NbMois);
  if FINDEMOIS(DateCalcul)=DateCalcul then NbMoisDetention := NbMois
  else NbMoisDetention := NbMois-1;
  if NbMoisPeriode<>0 then result :=NbMoisDetention/NbMoisPeriode else
  // CA - 22/04/2002 - Cas des exercice sur 1 mois.
  if (NbMoisPeriode=0) then  result :=1;
end;

//recupere le prorata depuis DateCalcul à DateFin
function GetProrataLinApres(DateDebutPer, DateFinPer, DateCalcul : TDateTime; bDotationExercice : boolean) : double;
var
  PremMois,PremAnnee : Word;
  NbMoisPeriode: Word;
  NbJourDetention,NbJourPeriode : integer;
begin
  Result := 1;
  if DateDebutPer=DateFinPer then exit;
  NOMBREMOIS (DateDebutPer,DateFinPer,PremMois,PremAnnee,NbMoisPeriode);
  NbMoisPeriode:=NbMoisPeriode-1;
  NbJourPeriode := NOMBREJOUR360(DateDebutPer,DateFinPer);
  NbJourDetention := NOMBREJOUR360(DateCalcul,DateFinPer);
  try
    if NbMoisPeriode<>0 then
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode
    end else
    // CA - 22/04/2002 - Cas des exercice sur 1 mois.  mbo fq 15287 31/03/06 ajout du <
    if (NbMoisPeriode=0) and (NbJourPeriode<=30) then
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode
    end;
  except
    result := 1;
  end;
end;

//recupere le prorata depuis DateDebut à DateCalcul
function GetProrataLinAvant(DateDebutPer, DateFinPer, DateDeb,DateFinAm, DateCalcul : TDateTime; bDotationExercice : boolean) : double;
var PremMois,PremAnnee,NbMois : Word;
  YDeb, MDeb, DDeb, DFin, MFin, YFin : Word;
  NbMoisPeriode, NbMoisAmort, NbJourDetention,NbJourPeriode : integer;
  tDateRef : TDateTime;
begin
  Result := 1;
  if (DateDebutPer=DateFinPer)
  or ((DateCalcul=DateFinPer) and (DateDebutPer=DateDeb))
  or ((DateCalcul>DateFinAm) and (DateFinAm<>iDate1900))  then exit;
  NOMBREMOIS(DateDebutPer,DateFinPer,PremMois,PremAnnee,NbMois);
  NbMoisPeriode:=NbMois-1;
  DecodeDate(DateDebutPer,YDeb, MDeb, DDeb);
  DecodeDate(DateFinPer,YFin, MFin, DFin);
  if (DFin = DaysPerMonth (YFin,MFin)) then DFin := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if (DDeb = DaysPerMonth (YDeb,MDeb)) then DDeb := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  // NbJourPeriode:=((NbMoisPeriode*30)+(MinIntValue([30,DFin])-MinIntValue([30,DDeb])+1));
  NbJourPeriode:=((NbMoisPeriode*30)+(DFin-MinIntValue([30,DDeb])+1));
  if DateDebutPer>DateDeb then tDateRef:=DateDebutPer else tDateRef:=DateDeb;
  NOMBREMOIS(tDateRef,DateCalcul,PremMois,PremAnnee,NbMois);
  NbMoisAmort:=NbMois-1;
  DecodeDate(tDateRef,YDeb, MDeb, DDeb);
  DecodeDate(DateCalcul,YFin, MFin, DFin);
  if (DFin = DaysPerMonth (YFin,MFin)) then DFin := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if (DDeb = DaysPerMonth (YDeb,MDeb)) then DDeb := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if NbMoisAmort=0 then NbJourDetention:=MinIntValue([30,DFin-DDeb+1])
  else if NbMoisAmort=1 then NbJourDetention:=MinIntValue([30,30-DDeb])+MinIntValue([30,DFin])+1
//  else NbJourDetention:=((NbMoisAmort*30)+(MinIntValue([30,DFin])-MinIntValue([30,DDeb])+1));
  else NbJourDetention:=((NbMoisAmort*30)+(DFin-MinIntValue([30,DDeb])+1));
  try
    if NbMoisPeriode<>0 then
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode ;
    end else
    // CA - 22/04/2002 - Cas des exercice sur 1 mois.
    if (NbMoisPeriode=0) and (NbJourPeriode=30) then
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode;
    end else
    if (NbMoisPeriode=0) then  // CA - 02/10/2003 - FQ 12631
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode;
    end;
  except
    result := 1;
  end;
end;

function GetNombrePeriodeMois(Base,Diviseur : integer) : integer;
var
  TmpCalcInt : integer;
  TmpCalcFlt : double;
begin
  TmpCalcInt := Base div Diviseur;
  TmpCalcFlt := Base / Diviseur;
  if TmpCalcInt <> TmpCalcFlt then result := TmpCalcInt +1
  else result := TmpCalcInt;
end;

function GetNombrePeriodeDate(DateRef1,DateRef2 : TDateTime) : integer ;
var
  PremMois,PremAnnee,NbMois : Word;
  TmpCalcInt : integer;
  TmpCalcFlt : double;
begin
  NOMBREMOIS (DateRef1,DateRef2,PremMois,PremAnnee,NbMois);
  TmpCalcInt := NbMois div 12;
  TmpCalcFlt := NbMois / 12;
  if TmpCalcInt <> TmpCalcFlt then result := TmpCalcInt +1
  else result := TmpCalcInt;
end;

// Modif CA le 15/03/99
// Modif BTY 03/06 Pb en éditions Web Access avec VHIMMO => PlanInfo.Exercices
procedure GetDatesExercice(DateCalcul : TDateTime; var DateDebut,DateFin : TDateTime);
var DateTmp: TDateTime; i: integer ; LongExo: word ;
begin
  // datecalcul : datedebutamort ou datedebut exercice precedent
  if DateFin=iDate1900 then DateTmp:=DateCalcul // debut recherche datecalcul
                       else DateTmp:=DateFin+1;

// BTY 03/06
{$IFDEF EAGLSERVER}
  i := 0; LongExo:=12 ;

  while TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Code<>'' do
{$ELSE}
  i := 1; LongExo:=12 ;

  while VHImmo^.Exercices[i].Code<>'' do
{$ENDIF}
  begin

  // BTY 03/06
{$IFDEF EAGLSERVER}
    //NOMBREMOIS(VHImmo^.Exercices[i].Deb,VHImmo^.Exercices[i].Fin,pMois,pAnnee,LongExo) ;
    if (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Deb<=DateTmp)
    and (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Fin>=DateTmp) then
{$ELSE}
    if (VHImmo^.Exercices[i].Deb<=DateTmp) and (VHImmo^.Exercices[i].Fin>=DateTmp) then
{$ENDIF}
    break;
    Inc(i,1);
  end;

// BTY 03/06
{$IFDEF EAGLSERVER}
  if TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Code<>''
  then
  GetDateDebutFinEx(TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i],
                    DateDebut,
                    DateFin,
                    LongExo)
{$ELSE}
  if  VHImmo^.Exercices[i].Code<>'' then
      GetDateDebutFinEx(VHImmo^.Exercices[i],DateDebut,DateFin,LongExo)
{$ENDIF}
  else if DateFin=iDate1900 then CalculDateDebutFinEx(DateCalcul,DateDebut,DateFin,LongExo)
  else
  begin
    DateDebut := DateTmp;
    DateFin := PlusMois(DateDebut,LongExo)-1;
  end;
end;


{procedure GetDateDebutFinEx(Q : TQuery; var DateDebut,DateFin : TDateTime);
begin
  Q.First;
  DateDebut:=StrToDate(Q.FindField('EX_DATEDEBUT').AsString);
  DateFin:=StrToDate(Q.FindField('EX_DATEFIN').AsString);
end;}

// Modif CA le 15/03/99
procedure GetDateDebutFinEx(Exo : TImExoDate; var DateDebut,DateFin : TDateTime;LongExo:integer=12);
begin
  DateDebut:= Exo.Deb;
  DateFin:= Exo.Fin;
end;

// Modif CA le 15/03/99
// Modif BTY 03/06 Pb en édition Web Access à cause de VHIMMO

procedure CalculDateDebutFinEx(DateCalcul : TDateTime;var DateDebut,DateFin : TDateTime; LongExo:integer=12);
var
  DateMin, DateMax : TDateTime;
  i : integer;
begin
  DateDebut:=iDate1900;
  DateFin  :=iDate1900;

// BTY 03/06
{$IFDEF EAGLSERVER}
  if TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[0].Code='' then exit;
  DateMin:=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[0].Deb;
  i := 0;

  while TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Code<>''
        do  Inc(i,1);
  if (i > 0) then
    DateMax:=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i-1].Fin
  else
    DateMax:=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[0].Fin;

{$ELSE}
  if VHImmo^.Exercices[1].Code='' then exit;
  DateMin:=VHImmo^.Exercices[1].Deb;
  i := 1;

  while VHImmo^.Exercices[i].Code<>'' do Inc(i,1);
  DateMax:=VHImmo^.Exercices[i-1].Fin;
{$ENDIF}

  if DateCalcul<DateMin then DecrementeExe(DateCalcul,DateMin,DateDebut,DateFin,LongExo)
  else if DateCalcul>DateMax then IncrementExe(DateCalcul,DateMax,DateDebut,DateFin,LongExo);
end;

//EPZ 03/11/00
// Modif BTY 03/06 Pb en édition Web Access à cause de VHIMMO

procedure CalculDateDebutFinExByPos(iPosEx : integer;var DateDebut,DateFin : TDateTime);
var i,iCount: integer ; pMois,pAnnee,LongExo: word;
begin

// BTY 03/06
{$IFDEF EAGLSERVER}

  // initialisations
  if (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[0].Code='') then
      exit;
  DateDebut:=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[0].Deb;
  DateFin  :=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[0].Fin;
  i:=0;

  // on parcourt le tableau EXERCICES jusqu'à l'exo correspondant à l'exercice en cours
  while (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Code<>'')
  and (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Deb <=
       TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Deb) do
  begin
    NOMBREMOIS(TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Deb,
               TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Fin,
               pMois, pAnnee, LongExo) ;
    DateDebut:=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Deb;
    DateFin  :=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Fin;
    Inc(i);
  end;

  // on parcourt le tableau EXERCICES à partir de l'exo en cours jusqu'au rang demandé en entrée
  iCount := 0;
  while (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Code<>'')
  and (iCount<=(iPosEx-1)) do
  begin
    NOMBREMOIS(TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Deb,
               TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Fin,
               pMois,pAnnee,LongExo) ;
    DateDebut := TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Deb;
    DateFin := TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Exercices[i].Fin;
    Inc(iCount); Inc(i);
  end;

  // le rang demandé n'est pas occupé par un exo => calcul théorique de l'exo
  if iCount<=(iPosEx-1) then
  begin
    DateDebut:=PlusMois(DateFin,12*((iPosEx-1)-iCount))+1;
    DateFin  :=PlusMois(DateDebut,12)-1;
  end;

{$ELSE}

  if VHImmo^.Exercices[1].Code='' then exit;
  DateDebut:=VHImmo^.Exercices[1].Deb;
  DateFin  :=VHImmo^.Exercices[1].Fin;
  i:=1;
  //on parcours jusqu'a l'exo correspondant à VHImmo^.Encours
  while (VHImmo^.Exercices[i].Code<>'') and (VHImmo^.Exercices[i].Deb<=VHImmo^.Encours.Deb) do
  begin
    NOMBREMOIS(VHImmo^.Exercices[i].Deb,VHImmo^.Exercices[i].Fin,pMois,pAnnee,LongExo) ;
    DateDebut:=VHImmo^.Exercices[i].Deb;
    DateFin  :=VHImmo^.Exercices[i].Fin;
    Inc(i);
  end;
  //on parcours tant que exo créé et pas trouve recherche
  iCount := 1;
  while (VHImmo^.Exercices[i].Code<>'') and (iCount<=iPosEx) do
  begin
    NOMBREMOIS(VHImmo^.Exercices[i].Deb,VHImmo^.Exercices[i].Fin,pMois,pAnnee,LongExo) ;
    DateDebut := VHImmo^.Exercices[i].Deb;
    DateFin := VHImmo^.Exercices[i].Fin;
    Inc(iCount); Inc(i);
  end;

  if iCount<=iPosEx then
  begin
    DateDebut:=PlusMois(DateFin,12*(iPosEx-iCount))+1;
    DateFin  :=PlusMois(DateDebut,12)-1;
  end;

{$ENDIF}

end;

//EPZ 03/11/00

// BTY 02/06 Pour les anciennes éditions, nelles fonctions CWAS
// BTY 02/06 FQ 17547 tableau Exercices démarre à l'indice 0
//{$IFDEF EAGLSERVER}
{procedure CalculDateDebutFinExByPosWA(PlanInfo: TPlanInfo; iPosEx : integer;
          var DateDebut,DateFin : TDateTime; DebutExo :TDateTime);
var i,iCount: integer ; pMois,pAnnee,LongExo: word;
begin
  if PlanInfo.Exercices[0].Code='' then exit;
  DateDebut:=PlanInfo.Exercices[0].Deb;
  DateFin  :=PlanInfo.Exercices[0].Fin;
  i:=0;
  //on parcourt le tableau EXERCICES jusqu'à l'exo correspondant à DebutExo
  while (PlanInfo.Exercices[i].Code <> '') and
        (PlanInfo.Exercices[i].Deb <= DebutExo) do
  begin
    NOMBREMOIS(PlanInfo.Exercices[i].Deb, PlanInfo.Exercices[i].Fin, pMois,pAnnee,LongExo);
    DateDebut:=PlanInfo.Exercices[i].Deb;
    DateFin  :=PlanInfo.Exercices[i].Fin;
    Inc(i);
  end;
  //on parcourt le tableau EXERCICES à partir de l'exo en cours jusqu'à l'indice exo demandé
  iCount := 0;
  while (PlanInfo.Exercices[i].Code<>'') and (iCount<=(iPosEx-1)) do
  begin
    NOMBREMOIS(PlanInfo.Exercices[i].Deb, PlanInfo.Exercices[i].Fin, pMois,pAnnee,LongExo) ;
    DateDebut := PlanInfo.Exercices[i].Deb;
    DateFin := PlanInfo.Exercices[i].Fin;
    Inc(iCount); Inc(i);
  end;
  if iCount<=(iPosEx-1) then
  begin
    DateDebut:=PlusMois(DateFin,12*((iPosEx-1)-iCount))+1;
    DateFin  :=PlusMois(DateDebut,12)-1;
  end;
end; }

// BTY 02/06 FQ 17547 tableau Exercices démarre à l'indice 0
{$IFDEF EAGLSERVER}

 // BTY 03/06
 //function TauxTVAServer (PlanInfo:TPlanInfo; ModeTVA,Tva : String; Achat:Boolean): double;
function TauxTVAServer (ModeTVA, Tva : String; Achat:Boolean): double;
var TACTobTva , TobT: TOB ;
begin
  result:=0 ;

  TACTobTva := TOB.Create('',Nil,-1);
  if TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.TobTva <> nil then
     TACTobTva.Dupliquer
     (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.TobTva,
      True,True);

//  TobT := PlanInfo.TobTva.FindFirst(
  TobT := TACTobTva.FindFirst(
          ['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',ModeTVA,TVA],False);
  if TobT <> nil then
    begin
    if Achat then result:= TobT.GetValue('TV_TAUXACH')/100.0
             else result:= TobT.GetValue('TV_TAUXVTE')/100.0 ;
    end;

  if TACTobTva <> nil then TACTobTva.Free;
end;
{$ENDIF}


procedure DecrementeExe(DateCalcul,DateRef : TDateTime; var DateDebut,DateFin : TDateTime; LongExo:integer=12);
var
  DateTmp : TDateTime;
begin
  DateTmp := DateRef;
  While DateCalcul<DateTmp do DateTmp:=PlusMois(DateTmp,LongExo*-1);
  DateDebut:=DateTmp;
  DateFin  :=PlusMois(DateDebut,LongExo)-1;
end;

procedure IncrementExe(DateCalcul,DateRef : TDateTime; var DateDebut,DateFin : TDateTime; LongExo:integer=12);
var
  DateTmp : TDateTime;
begin
  DateTmp  :=DateRef;
  While DateCalcul>DateTmp do DateTmp:=PlusMois(DateTmp,LongExo);
  DateFin  :=DateTmp;
  DateDebut:=PlusMois(DateFin,LongExo*-1)+1;
end;

// BTY 03/06 Pb en édition Web Access à cause de VHIMMO
function ExerciceEnCours (DateRef : TDateTime) : boolean;
begin
 // BTY 03/06
{$IFDEF EAGLSERVER}
  result := (DateRef >=
        TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Deb)
        and (DateRef <=
        TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Fin);
{$ELSE}
  result := (DateRef>=VHImmo^.Encours.Deb) and (DateRef<=VHImmo^.Encours.Fin);
{$ENDIF}
end;

Procedure TraiteParamDateFinAmort(DateDebAmort,DateFinAmort : TDateTime; var DateDebut,DateFin,DateCalcul : TDateTime);
begin
  if (DateDebAmort>=DateDebut) and (DateDebAmort<=DateFin) then DateDebut:=DateDebAmort;
  if (DateFinAmort>=DateDebut) and (DateFinAmort<=DateFin) then DateFin:=DateFinAmort;
  if DateCalcul>DateFin   then DateCalcul := DateFin;
  if DateCalcul<DateDebut then DateCalcul := DateDebut;
end;

//30/06/99 pour immo saisie sur N avec date achat N+1 (avec N+1 exo ouvert)
// BTY 03/06 Pb en édition Web Access à cause de VHIMMO
function GetDateDebutExoRef(DateOp : TDateTime) : TDateTime;
var
  DateDebut,DateFin : TDateTime;
begin
  // BTY 03/06
 {$IFDEF EAGLSERVER}
  if DateOp <= TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Fin then
     result := TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Deb else
 {$ELSE}
  if DateOp <= VHImmo^.Encours.Fin then result := VHImmo^.Encours.Deb else
 {$ENDIF}
  begin
    DateFin:=iDate1900;
    GetDatesExercice(DateOp,DateDebut,DateFin);
    result := DateDebut;
  end;
end;
//30/06/99 pour immo saisie sur N avec date achat N+1 (avec N+1 exo ouvert)


// BTY 03/06
// BTY 02/06 Anciennes éditions à ouvrir à CWAS
//{$IFDEF EAGLSERVER}
//function CalculEchTva(PlanInfo:TPlanInfo; Q : TQuery; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
//{$ELSE}
//function CalculEchTva(Q : TQuery; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
//{$ENDIF}
//var
// dCoeffTvaRable,dCoeffTvaRee : double;
//begin
//{$IFDEF EAGLSERVER}
//  GetCoeffTVARable(PlanInfo, Q,laTVADefaut,dCoeffTVARable, dCoeffTVARee);
//{$ELSE}
//  GetCoeffTVARable(Q,laTVADefaut,dCoeffTVARable, dCoeffTVARee);
//{$ENDIF}
//  if dCoeffTVARee=1 then result := Montant
//  else if dCoeffTVARee=0 then result := Montant / (1+dCoeffTVARable)
//  else result := Montant / (1+(dCoeffTVARable*dCoeffTVARee));
//end;

// BTY 03/06
// BTY 02/06 Anciennes éditions à ouvrir à CWAS
//{$IFDEF EAGLSERVER}
//procedure GetCoeffTVARable(PlanInfo:TPlanInfo; Q : TQuery; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
//{$ELSE}
//procedure GetCoeffTVARable(Q : TQuery; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
//{$ENDIF}
//var dMontantHT : double;
//begin
//{$IFDEF EAGLSERVER}
//  dCoeffTvaRable := GetTauxTVAOrganisme (PlanInfo, Q.FindField('I_ORGANISMECB').AsString,laTVADefaut);
//{$ELSE}
// dCoeffTvaRable := GetTauxTVAOrganisme (Q.FindField('I_ORGANISMECB').AsString,laTVADefaut);
//{$ENDIF}
//  dCoeffTvaRee := 1;
//if Q.FindField('I_TVARECUPERABLE').AsFloat<>0 then
//begin
//    dCoeffTvaRee := Q.FindField('I_TVARECUPEREE').AsFloat / Q.FindField('I_TVARECUPERABLE').AsFloat;
//    dMontantHT := Q.FindField('I_MONTANTHT').AsFloat;
//    if dMontantHT<>0 then dCoeffTvaRable := Q.FindField('I_TVARECUPERABLE').AsFloat / dMontantHT;
//end;
//end;

// BTY 03/06
// BTY 02/06 Anciennes éditions à ouvrir à CWAS
//{$IFDEF EAGLSERVER}
//function GetTauxTVAOrganisme (PlanInfo:TPlanInfo; stOrganisme, stGeneTVA : string) : double;
//{$ELSE}
function GetTauxTVAOrganisme (stOrganisme, stGeneTVA : string) : double;
//{$ENDIF}
var Q : TQuery;
    TauxTVA : double;
begin
 Q := OpenSQL ('SELECT T_AUXILIAIRE,T_REGIMETVA FROM TIERS WHERE T_AUXILIAIRE="'+stOrganisme+'"',TRUE);
 try
   if not Q.Eof then
   begin
{$IFDEF EAGLSERVER}
// BTY 03/06
//    TauxTVA := TauxTVAServer (PlanInfo, Q.FindField ('T_REGIMETVA').AsString, stGeneTVA, True);
    TauxTVA := TauxTVAServer (Q.FindField ('T_REGIMETVA').AsString, stGeneTVA, True);
{$ELSE}
    TauxTVA := ImTVA2Taux (Q.FindField ('T_REGIMETVA').AsString,stGeneTVA,True);
{$ENDIF}
   end else TauxTVA := 0;
 finally
   Ferme (Q);
 end;
 Result := TauxTVA;
end;

// BTY 03/06
//EPZ 12/12/2000
//{$IFDEF EAGLSERVER}
//{$ELSE}
// BTY 02/06 Ouvrir à CWAS pour anciennes éditions
//function GetTauxTVAOrganisme (stOrganisme, stGeneTVA : string) : double;
//var Q : TQuery;
//    TauxTVA : double;
//begin
// Q := OpenSQL ('SELECT T_AUXILIAIRE,T_REGIMETVA FROM TIERS WHERE T_AUXILIAIRE="'+stOrganisme+'"',TRUE);
// try
//   if not Q.Eof then
//   begin
//     TauxTVA := ImTVA2Taux (Q.FindField ('T_REGIMETVA').AsString,stGeneTVA,True);
//   end else TauxTVA := 0;
// finally
//   Ferme (Q);
// end;
// Result := TauxTVA;
//end;

// BTY 03/06
// BTY 02/06 Ouvrir à CWAS pour anciennes éditions
function CalculEchTva(Q : TQuery; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
var
  dCoeffTvaRable,dCoeffTvaRee : double;
begin
  GetCoeffTVARable(Q,laTVADefaut,dCoeffTVARable, dCoeffTVARee);
  if dCoeffTVARee=1 then result := Montant
  else if dCoeffTVARee=0 then result := Montant / (1+dCoeffTVARable)
  else result := Montant / (1+(dCoeffTVARable*dCoeffTVARee));
end;

// BTY 03/06
{$IFDEF EAGLSERVER}
{$ELSE}

function CalculEchTvaTOB(OB : TOB; laTVADefaut : string;Montant : double ; var dEchHT : double) : double;
var
  dCoeffTvaRable,dCoeffTvaRee : double;
begin
  GetCoeffTVARableTOB(OB,laTVADefaut,dCoeffTVARable, dCoeffTVARee);
  if dCoeffTVARee=1 then result := Montant
  else if dCoeffTVARee=0 then result := Montant / (1+dCoeffTVARable)
  else result := Montant / (1+(dCoeffTVARable*dCoeffTVARee));
end;

function AffecteTvaEch(Q : TQuery; laTVADefaut : string;MontantHT : double) : double;
var
  dEchTTC,dCoeffTvaRable,dCoeffTvaRee : double;
  //dEchHT,dTvaRable,dTVARec : double; //XVI Conseil compile
begin
  GetCoeffTVARable(Q,laTVADefaut,dCoeffTVARable, dCoeffTVARee);
  if dCoeffTVARee=0 then dEchTTC := MontantHT * (1+dCoeffTVARable)
  else if dCoeffTVARee=1 then dEchTTC := MontantHT
  else dEchTTC := MontantHT * (1+(dCoeffTVARable*dCoeffTVARee));

  result := dEchTTC;
  exit;

  {dTvaRable := dEchHT * dCoeffTVARable;
  dTVARec := dTvaRable * dCoeffTvaRee;

  result := MontantHT + (dTvaRable - dTVARec);} //XVI Conseil compile
end;

procedure GetCoeffTVARableTOB(OB : TOB; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
var dMontantHT : double;
begin
  dCoeffTvaRable := GetTauxTVAOrganisme (OB.GetValue('I_ORGANISMECB'),laTVADefaut);
  dCoeffTvaRee := 1;
  if OB.GetValue('I_TVARECUPERABLE')<>0 then
  begin
    dCoeffTvaRee := OB.GetValue('I_TVARECUPEREE') / OB.GetValue('I_TVARECUPERABLE');
    dMontantHT := OB.GetValue('I_MONTANTHT');
    if dMontantHT<>0 then dCoeffTvaRable := OB.GetValue('I_TVARECUPERABLE') / dMontantHT;
  end;
end;
{$ENDIF}


// BTY 03/06
// BTY 02/06 Ouvrir à CWAS pour anciennes éditions
procedure GetCoeffTVARable(Q : TQuery; laTVADefaut : string; var dCoeffTVARable, dCoeffTVARee : double);
var dMontantHT : double;
begin
  dCoeffTvaRable := GetTauxTVAOrganisme (Q.FindField('I_ORGANISMECB').AsString,laTVADefaut);
  dCoeffTvaRee := 1;
  if Q.FindField('I_TVARECUPERABLE').AsFloat<>0 then
  begin
    dCoeffTvaRee := Q.FindField('I_TVARECUPEREE').AsFloat / Q.FindField('I_TVARECUPERABLE').AsFloat;
    dMontantHT := Q.FindField('I_MONTANTHT').AsFloat;
    if dMontantHT<>0 then dCoeffTvaRable := Q.FindField('I_TVARECUPERABLE').AsFloat / dMontantHT;
  end;
end;


// BTY 03/06 Pb en édition Web Access à cause de VHIMMO

function CalculSurAmortissement ( Dotation : double; DateAchat, DateDebEx,DateFinEx : TDateTime; bIntermediaire : boolean = FALSE; DateFin : TDateTime = 2) : double;
var PremMois,PremAnnee,nbMois, nbMoisPer, nbMoisExo : Word;
    ExoAchat : TExoDate;
    DotMens : double;
    Prorata : double;
    bSansExercice : boolean;
begin
  NOMBREMOIS (DateDebEx, DateFinEx,PremMois,PremAnnee,nbMoisExo);
  nbMois := nbMoisExo;
//  QuelDateDeExo(QuelExoDt(DateAchat),ExoAchat);
  ImCQuelExercice(DateAchat,ExoAchat);

  // BTY 03/06
{$IFDEF EAGLSERVER}
  bSansExercice := (ExoAchat.Code = TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Code)
                   and (DateAchat < TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Deb );

  if (((ExoAchat.Fin = DateFinEx) and ( ExoAchat.Deb = DateDebEx)) and (not bSansExercice))
       or (bSansExercice and (DateFinEx < TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Deb)) then
{$ELSE}
  bSansExercice := (ExoAchat.Code = VHImmo^.EnCours.Code) and (DateAchat <VHImmo^.Encours.Deb ); // Pour contourner pb QuelExoDt

  if (((ExoAchat.Fin = DateFinEx) and ( ExoAchat.Deb = DateDebEx)) and (not bSansExercice))
       or (bSansExercice and  (DateFinEx < VHImmo^.Encours.Deb)) then
{$ENDIF}

  begin
    // année 1 du suramortissement
    NOMBREMOIS (DateAchat,DateFinEx,PremMois,PremAnnee,nbMois);
    Result := Dotation * 1.3;
  end
  // BTY 03/06
{$IFDEF EAGLSERVER}
  else if ((DateDebEx=ExoAchat.Fin+1) and (not bSansExercice)) or
          (bSansExercice and (DateFinEx=TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Fin)) then
{$ELSE}
  else if ((DateDebEx=ExoAchat.Fin+1) and (not bSansExercice)) or (bSansExercice and (DateFinEx=VHImmo^.Encours.Fin)) then
{$ENDIF}

  begin
    // année 2 du suramortissement
    NOMBREMOIS (DateAchat,DateDebEx-1,PremMois,PremAnnee,nbMois);
    if nbMois <= 12 then nbMois := 12 - nbMois;
    if (nbMois > 12) then nbMois := 12;
    if bIntermediaire then
    begin
  // BTY 03/06
{$IFDEF EAGLSERVER}
      NOMBREMOIS (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Deb,
                  TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.fEncours.Fin,
                  PremMois,
                  PremAnnee,
                  nbMoisExo);
{$ELSE}
      NOMBREMOIS (VHImmo^.Encours.Deb, VHImmo^.Encours.Fin,PremMois,PremAnnee,nbMoisExo);
{$ENDIF}

      // Si c'est un calcul intermédiaire, la dotation comporte déjà le montant du suramortissement
      // On recalcul donc le montant de l'amortissement mensuel sans suramortissement
      DotMens := ((Dotation * nbMoisExo) / ( nbMoisExo + 0.3*nbMois))/nbMoisExo;  // Dotation mensuelle sans suramortissement
      // Dans ce cas DateFinEx = Date de calcul
      NombreMois(DateDebEx, DateFinEx, PremMois,PremAnnee,nbMoisPer);
      if nbMoisPer > nbMois then Result := (DotMens*(1.3*nbMois))+(DotMens*(nbMoisPer-nbMois))
      else Result := (DotMens*(1.3*nbMoisPer));
    end else
    begin
      Result := Dotation * (1 + (0.3*(nbMois/nbMoisExo)));
    end;
  end
  else
  begin
    Prorata := GetProrataDegAvant(dateDebEx,DateFin,DateAchat,DateFinEx);
    Result := Arrondi (Dotation*Prorata,V_PGI.OkDecV);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Retourne le coefficient à appliquer au taux linéaire pour
Suite ........ : calculer le dégressif en fonction de la durée et de la date
Suite ........ : d'achat.
Suite ........ : BTY 03/06 En éditions WEb Access pb avec VHIMMO
Mots clefs ... :
*****************************************************************}
function AMCoefficientDegressif ( DateAchat : TDateTime; Duree : integer ) : double;
var i : integer;
    T : TOB;
// BTY 03/06
{$IFDEF EAGLSERVER}
   TACCoeffDegressif:TOB;
{$ENDIF}
begin
  Result := 0;
  T := nil;
  i:=1;

  // BTY 03/06
{$IFDEF EAGLSERVER}
  TACCoeffDegressif := TOB.Create('',Nil,-1);
  if TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.CoeffDegressif <> nil then
     TACCoeffDegressif.Dupliquer
     (TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.CoeffDegressif,
      True,True);
  while ((T=nil) and (i<TACCoeffDegressif.Detail.Count))   do
  begin
    if StrToDate(TACCoeffDegressif.Detail[i].GetValue('DATE'))>DateAchat then
      T := TACCoeffDegressif.Detail[i-1]
    else Inc(i,1);
  end;
  { si pas trouvé, on prend les derniers coefficients }
  if T = nil then
     T := TACCoeffDegressif.Detail[TACCoeffDegressif.Detail.Count-1];

{$ELSE}
  while ((T=nil) and (i<VHImmo^.CoeffDegressif.Detail.Count)) do
  begin
    if StrToDate(VHImmo^.CoeffDegressif.Detail[i].GetValue('DATE'))>DateAchat then
      T := VHImmo^.CoeffDegressif.Detail[i-1]
    else Inc(i,1);
  end;
  { si pas trouvé, on prend les derniers coefficients }
  if T = nil then T := VHImmo^.CoeffDegressif.Detail[VHImmo^.CoeffDegressif.Detail.Count-1];
{$ENDIF}

  if T <> nil then
  begin
    { récupération du bon coefficient en fonction de la durée demandée }
    if Duree <= 48 then Result := T.GetValue('COEFF1')
    else if Duree <= 72 then Result := T.GetValue('COEFF2')
    else Result := T.GetValue('COEFF3');
  end;

// BTY 03/06
{$IFDEF EAGLSERVER}
if (TACCoeffDegressif <> nil) then   TACCoeffDegressif.free;
{$ENDIF}

end;



{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Calcul de la durée d'amortissement en fonction du taux
Mots clefs ... :
*****************************************************************}
function AMTaux2Duree (Methode: string; DateAchat: TDateTime; Taux : double ) : integer;
begin
  if Methode = 'LIN' then Result := AMTaux2DureeLineaire( Taux )
  else if Methode = 'DEG' then Result := AMTaux2DureeDegressif( Taux, DateAchat )
  else Result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Calcul de la durée d'amortissement en fonction du taux pour
Suite ........ : un amortissement linéaire
Mots clefs ... :
*****************************************************************}
function AMTaux2DureeLineaire( Taux : double ) : integer;
begin
     Result := Trunc(Arrondi( ( 1200 / Taux ), 0 ));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Calcul de la durée d'amortissement en fonction du taux pour
Suite ........ : un amortissement dégressif
Mots clefs ... :
*****************************************************************}
function AMTaux2DureeDegressif ( Taux : double ; DateAchat : TDateTime ) : integer;
var TauxCalc  : double;
    iMois     : integer;
    Duree     : integer;
    MemDuree  : integer;
begin
  { ne connaissant pas la durée, on est obligé de partir de la valeur du taux pour estimer le coefficient }
  if Taux >= 100 then Duree := AmTaux2DureeLineaire(Taux)
  else
  begin
    Duree := 0;
    MemDuree := 0;
    iMois := 36;
    while (iMois < 1000) do
    begin
      { calcul du taux théorique à pour une durée de iMois }
      TauxCalc := Arrondi(GetTauxLineaire(DateAchat, iMois)*AmCoefficientDegressif(DateAchat,iMois),2);
      if TauxCalc = Arrondi(Taux,2) then
      begin
        { si le taux théorique est égal au taux initial, on a trouvé }
        if (iMois mod 12)=0 then
        begin
          { si on trouve un nombre d'années entier, c'est sûr on a trouvé }
          Duree := iMois;
          break;
          { si on n'a pas un nombre d'années entier, on continue au cas où on trouve une valeur plus 'sympathique' }
          { ex : pour un amortissement antérieur à 2001 ( hors 96 ), 45 mois et 48 donnent 40% : on garde 48 mois }
        end else MemDuree := iMois;
      end;
      { au delà de 7 ans le taux est régulièrement décroissant : si on est inférieur, on ne trouvera plus le taux ! }
      if (iMois>84) and (TauxCalc < Taux) then
      begin
        Duree := MemDuree;
        break;
      end;
      Inc(iMois,1);
    end;
  end;
  { si on ne trouve pas de taux dégressif, on utilise la méthode linéaire }
  if Duree = 0 then Duree := AmTaux2DureeLineaire(Taux);
  Result := Duree;
end;

function NOMBREJOUR360 ( DateDebut, DateFin : TDateTime ) : integer;
var
  PremMois,PremAnnee,NbMoisPeriode : Word;
  YDeb, MDeb, DDeb, DFin, MFin, YFin : Word;
begin
  NOMBREMOIS (DateDebut,DateFin,PremMois,PremAnnee,NbMoisPeriode);
  // MVG 19/02/2007 - reconduction du fonctionnement en DELPHI 5 Ne pas faire -1 sur 0
  if NbMoisPeriode<>0 then NbMoisPeriode:=NbMoisPeriode-1 else NbMoisPeriode:=65535;
  DecodeDate(DateDebut,YDeb, MDeb, DDeb);
  DecodeDate(DateFin,YFin, MFin, DFin);
  if (DFin = DaysPerMonth (YFin,MFin)) then DFin := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if (DDeb = DaysPerMonth (YDeb,MDeb)) then DDeb := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  Result := ((NbMoisPeriode*30)+(DFin-MinIntValue([30,DDeb])+1));
end;

// new mbo 28.09.2007 - fq 21754
// fonction particulière servant uniquement pour la sortie sans prise en compte du jour de cession

function GetProrataLinSortie(DateDebutPer, DateFinPer, DateDeb,DateFinAm, DateCalcul : TDateTime; bDotationExercice : boolean) : double;
var PremMois,PremAnnee,NbMois : Word;
  YDeb, MDeb, DDeb, DFin, MFin, YFin : Word;
  NbMoisPeriode, NbMoisAmort, NbJourDetention,NbJourPeriode : integer;
  tDateRef : TDateTime;
begin
  Result := 1;
  if (DateDebutPer=DateFinPer)
  or ((DateCalcul=DateFinPer) and (DateDebutPer=DateDeb))
  or ((DateCalcul>DateFinAm) and (DateFinAm<>iDate1900))  then exit;
  NOMBREMOIS(DateDebutPer,DateFinPer,PremMois,PremAnnee,NbMois);
  NbMoisPeriode:=NbMois-1;
  DecodeDate(DateDebutPer,YDeb, MDeb, DDeb);
  DecodeDate(DateFinPer,YFin, MFin, DFin);
  if (DFin = DaysPerMonth (YFin,MFin)) then DFin := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if (DDeb = DaysPerMonth (YDeb,MDeb)) then DDeb := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  // NbJourPeriode:=((NbMoisPeriode*30)+(MinIntValue([30,DFin])-MinIntValue([30,DDeb])+1));
  NbJourPeriode:=((NbMoisPeriode*30)+(DFin-MinIntValue([30,DDeb])+1));
  if DateDebutPer>DateDeb then tDateRef:=DateDebutPer else tDateRef:=DateDeb;
  NOMBREMOIS(tDateRef,DateCalcul,PremMois,PremAnnee,NbMois);
  NbMoisAmort:=NbMois-1;
  DecodeDate(tDateRef,YDeb, MDeb, DDeb);
  DecodeDate(DateCalcul,YFin, MFin, DFin);
  if (DFin = DaysPerMonth (YFin,MFin)) then DFin := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if (DDeb = DaysPerMonth (YDeb,MDeb)) then DDeb := 30;  // CA - 03/06/2003 - Si dernier jour du mois ==> jour = 30 ( correction pb 28/02 )
  if NbMoisAmort=0 then NbJourDetention:=MinIntValue([30,DFin-DDeb+1])
  else if NbMoisAmort=1 then NbJourDetention:=MinIntValue([30,30-DDeb])+MinIntValue([30,DFin])+1

  else NbJourDetention:=(((NbMoisAmort*30)+(DFin-MinIntValue([30,DDeb])+1)))-1; // on fait moins un ds ce cas précis
  try
    if NbMoisPeriode<>0 then
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode ;
    end else
    // CA - 22/04/2002 - Cas des exercice sur 1 mois.
    if (NbMoisPeriode=0) and (NbJourPeriode=30) then
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode;
    end else
    if (NbMoisPeriode=0) then  // CA - 02/10/2003 - FQ 12631
    begin
      if bDotationExercice then NbJourPeriode:=360 ;
      result :=NbJourDetention/NbJourPeriode;
    end;
  except
    result := 1;
  end;
end;

function GetDotationLinSortie(DateDebut,DateFin,DateDebAmort,DateCalcul,DateFinAmort : TDateTime; Annuite : double; bDotationExercice : boolean): double;
var
  Prorata : double;
begin
  result:=0.00;
  if DateCalcul<DateDebAmort then exit;
  if DateFinAmort<>iDate1900 then TraiteParamDateFinAmort(DateDebAmort,DateFinAmort,DateDebut,DateFin,DateCalcul); // fin amort ou cessio
  Prorata := GetProrataLinSortie(DateDebut,DateFin,DateDebAmort,DateFinAmort,DateCalcul,bDotationExercice);
  result := Arrondi (Annuite*Prorata,V_PGI.OkDecV);
end;


end.

