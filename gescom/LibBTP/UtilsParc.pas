unit UtilsParc;

interface

uses  M3FP,
      StdCtrls,
      Controls,
      Classes,
      Forms,
      SysUtils,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOB,
      UTOF,
      AglInit,
      Agent,
      EntGC,
      StrUtils,
      DateUtils,
      HeureUtil,
      ParamSoc,
      Constantes,
      UProcGen,
      UtilSaisieconso,
      BTPUtil,
{$IFDEF EAGLCLIENT}
      MaineAGL,HPdfPrev,UtileAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
      HTB97;

Var
  ListeChamps : Array of String;

function AffaireExist (Affaire : string; WhereSql : string='' ) : boolean;

Procedure CalculCout(TobCout : TOB);
Function  CalculDateFinParcMat(NBHJour, Duree : Double; DDeb : TdateTime; TobjFerie : Tob) : String; OverLoad;
Function  CalculDateFinParcMat(DateDeb, Datefin : TDateTime; NBHJour : Double; Duree : Integer; TobjFerie : Tob) : TDateTime; overload
function  CalculDureeEvent(Borne1, Borne2 : TDateTime; TobJFerie, TobCalendrier : TOB) : Double;
function  CalculJourPaques(Annee : Integer) : TDateTime;
Function  CalculNbHeure(DDeb, Dfin : TdateTime) : String;
Function  ChargeInfoAffaire  (code : string; WithMsg : Boolean)  : String;
Function  ChargeInfoArticle  (code, StWhere : string; WithMsg    : Boolean) : String;
Function  ChargeInfoFamMat   (code : string; WithMsg : Boolean)  : String;

Function  ChargeInfoMateriel (code : string; WithMsg : Boolean)  : String  ;OverLoad;
Procedure ChargeInfoMateriel (code : string; TobMateriel : Tob);OverLoad;

Function  ChargeInfoRessource(code, StWhere : string; WithMsg    : Boolean) : String ;OverLoad;
Procedure ChargeInfoRessource(code, Stwhere : string; TobRessource : Tob);overload;

Function  ChargeInfoTiers    (code : string; WithMsg : Boolean)  : String;
Function  ChargeInfoTypeAction(code, Stwhere : string; WithMsg : Boolean) : String;
Procedure ChargeInfoTableComplexe(Table, Condition : String; TobFic : Tob);
Function  ChargeInfoTableSimple(Table, Condition : String; ListeChamps : Array of string) : String;
Procedure ChargeJoursFeries(TobJferie : TOB);
Function  ChargeParametreTemps : Double;
function  CodeMaterielExist (CodeMateriel : string; WhereSql : string=''; Obligatoire : Boolean=True ) : boolean;
Function  ControleCoherence(DateDeb,DateFin : TDateTime; Duree : double; Materiel : string; NoEnCours : Integer) : boolean;
Function  CreateConsoByEventMat(TobEvent : TOB) : Boolean;
procedure CreateLigneCom(TobPiece, TOBL : TOB);
Procedure CreateLigneCompl(TOBL, TobLigne : tob);
procedure CreateLigneDoc(TobPiece, TOBL : TOB);
Procedure CreateTOBPiece(TobPiece : TOB);


function DateEstOuvre(aDate: TDateTime; CodeCalendrier, CodeRessource : String): Boolean;
function DetermineJourTravaille(CodeCalendrier, CodeRessource : string) :  Integer;

function EventMatExists (CodeMateriel,CodeEvt,CodeAffaire : string) : boolean;

function FamilleMatExist (CodeFam : string; WhereSql : string='' ) : boolean;

Function GereConso(BtEtat : string) : Boolean;

Function JourFeries(TobFerie : TOB; DateCtrl : TdateTime) : Boolean;
Function JourNonTravaille(CodeCalendrier, codeRessource : string; NumJour : Integer) : Boolean;

function PrestationExist (CodeArticle : string; WhereSql : string='' ) : boolean;

Procedure RechercheCalendrier(CodeRes : string; TobCalendrier : Tob);
Function  RecupTempsMoyen(CodeCalendrier, CodeRessource : string) : Double;
Procedure ReinitialiseTob(NomTob : Tob);
function  RessourceExist (CodeRessource : string; WhereSql : string='' ) : boolean;

function  TiersExist (Tiers : string; WhereSql : string='' ) : boolean;
function  TraitementConso(NumEventMat : Integer; DateImp : TDateTime) : Boolean;
Procedure TraitementRealise(TobEvent : TOB);
function  TypeActionExist (typeAction : string; WhereSql : string=''; Obligatoire : Boolean=True ) : boolean;

Function VerifResultat(ZoneText : String; WithMsg : Boolean; NumMsg : Integer) : String;

const
  TexteMessage: array[1..14] of string  = (
          {1}   'Code tiers inexistant'
          {2},  'Ressource inexistante'
          {3},  'Chantier inexistant'
          {4},  'Type action inexistant'
          {5},  'Matériel inexistant'
          {6},  'Famille matériel inexistante'
          {7},  'Article inexistant'
          {8},  'Tiers fermé. Veuillez saisir un nouveau Tiers'
          {9},  'Ressource fermée. Veuillez saisir une nouvelle Ressource'
          {10},  'Chantier fermé. Veuillez saisir un nouveau Chantier'
          {11},  'Type d''action fermé. Veuillez saisir un nouveau type d''action'
          {12},  'Matériel fermé. Veuillez saisir un nouveau Matériel'
          {13}, 'Famille matériel fermée. Veuillez saisir un nouvelle Famille Matériel'
          {14}, 'Article fermé. Veuillez saisir un nouvel Article'
            );


implementation
uses  UdateUtils, Math;

//Permet de récupérer le libellé et de vérifier si le matériel n'est pas fermé
Function ChargeInfoMateriel(code : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,2);

  Listechamps[0] := 'BMA_LIBELLE';
  ListeChamps[1] := 'BMA_FERME';

  Tempo  := ChargeInfoTableSimple('BTMATERIEL', 'BMA_CODEMATERIEL = "' + Code + '"', Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 5);

end;

//Permet de récupérer la totalité de l'enreg dans une tob
Procedure ChargeInfoMateriel(code : string; TobMateriel : Tob);
begin

  ReinitialiseTob(TobMateriel);

  if Code = '' then exit;

  ChargeInfoTableComplexe('BTMATERIEL', 'BMA_CODEMATERIEL = "' + Code + '"', TobMateriel);

  if TobMateriel.GetBoolean('BMA_FERME') then
  begin
    PGIError(TexteMessage[12], 'Erreur Saisie');
    ReinitialiseTob(TobMateriel);
  end;

end;

Function ChargeInfoTypeAction(code, StWhere : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,1);

  Listechamps[0] := 'BTA_LIBELLE';

  Tempo := ChargeInfoTableSimple('BTETAT', 'BTA_BTETAT = "' + Code + '"' + StWhere, Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 4);

end;


//Permet de récupérer la totalité de l'enreg dans une tob
Procedure ChargeInfoRessource(code, Stwhere : string; TobRessource : Tob);
begin

  ReinitialiseTob(TobRessource);

  if Code = '' then exit;

  ChargeInfoTableComplexe('RESSOURCE', 'ARS_RESSOURCE = "' + Code + '"' + Stwhere, TobRessource);

  if TobRessource.GetBoolean('ARS_FERME') then
  begin
    PGIError(TexteMessage[9], 'Erreur Saisie');
    ReinitialiseTob(TobRessource);
  end;

end;

Function ChargeInfoRessource(code, StWhere : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,2);

  Listechamps[0] := 'ARS_LIBELLE';
  ListeChamps[1] := 'ARS_FERME';

  if StWhere <> '' then
    Tempo := ChargeInfoTableSimple('RESSOURCE', ' ARS_RESSOURCE = "' + Code + '"' + Stwhere, Listechamps)
  else
    Tempo := ChargeInfoTableSimple('RESSOURCE', ' ARS_RESSOURCE = "' + Code + '"', Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 2);

end;

Function ChargeInfoTiers(code : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,2);

  Listechamps[0] := 'T_LIBELLE';
  ListeChamps[1] := 'T_FERME';

  Tempo := ChargeInfoTableSimple('TIERS', 'T_TIERS = "' + Code + '" AND T_NATUREAUXI ="CLI"', Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 1);

end;

Function ChargeInfoAffaire(code : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,1);

  Listechamps[0] := 'AFF_LIBELLE';

  Tempo := ChargeInfoTableSimple('AFFAIRE', 'AFF_AFFAIRE = "' + Code + '"', Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 3);

end;

Function ChargeInfoFamMat(code : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,1);

  Listechamps[0] := 'BFM_LIBELLE';

  Tempo := ChargeInfoTableSimple('BTFAMILLEMAT', 'BFM_CODEFAMILLE = "' + Code + '"', Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 6);

end;

Function ChargeInfoArticle(Code, StWhere : string; WithMsg : Boolean) : String;
Var Tempo : string;
begin

  Result := '';

  if Code = '' then exit;

  SetLength(ListeChamps,1);

  Listechamps[0] := 'GA_LIBELLE';

  if StWhere <> '' then
    Tempo := ChargeInfoTableSimple('ARTICLE', 'GA_ARTICLE = "' + Code + '" OR GA_CODEARTICLE = "' + Code + '" ' + Stwhere, Listechamps)
  else
    Tempo := ChargeInfoTableSimple('ARTICLE', 'GA_ARTICLE = "' + Code + '" OR GA_CODEARTICLE = "' + Code + '" ', Listechamps);

  Result := VerifResultat(Tempo, WithMsg, 7);

end;

Function ChargeInfoTableSimple(Table, Condition : String; ListeChamps : Array of string) : String;
Var Champs  : String;
    StSQL   : string;
    QQ      : TQuery;
    TobTable: TOB;
    Ind     : Integer;
begin

  //On formate une zone texte pour pouvoir intégrer les champs à une requête
  Champs := '';

  Result := '';

  for ind := 0 to High(ListeChamps) do
  begin
    if Champs = '' then Champs := ListeChamps[Ind]
    else Champs := Champs + ',' + ListeChamps[Ind];
  end;

  //On crée la requete à partir des éléments passés
  StSQl := 'SELECT ' + Champs + ' FROM ' + Table + ' WHERE ' + Condition;
  QQ    := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    TOBTable := Tob.create('TABLE', nil, -1);
    TobTable.SelectDB('ELEMENT', QQ);
    For Ind := 0 To High(Listechamps) do
    begin
      if Ind = 0 then Result := TobTable.GetValue(Listechamps[Ind])
      else Result := Result + ';' + TobTable.GetValue(Listechamps[Ind]);
    end;
    FreeAndNil(TobTable);
  end;

  Ferme(QQ);

end;

Procedure ChargeInfoTableComplexe(Table, Condition : String; TobFic : Tob);
Var StSQL   : string;
    QQ      : TQuery;
begin

  //On crée la requete à partir des éléments passés
  StSQl := 'SELECT * FROM ' + Table + ' WHERE ' + Condition;
  QQ    := OpenSQL(StSQL, False);
  If Not QQ.Eof then
  begin
    TobFic.SelectDB('ELEMENTS', QQ);
  end;

  Ferme(QQ);

end;



Function VerifResultat(ZoneText : String; WithMsg : Boolean; NumMsg : Integer) : String;
Var StVal : string;
    ind   : Integer;
begin

  if ZoneText = '' then
  begin
    Result := '';
    if WithMsg then PGIError(TexteMessage[NumMsg], 'Erreur Saisie');
  end
  else
  begin
    for ind := 0 To High(ListeChamps) do
    begin
      StVal := ReadTokenst(ZoneText);
      Case ind of
        //On charge le libellé
        0 : Result := StVal;
        //On conterole si enreg est fermé
        1 : if StVal = 'X' then
            begin
              if WithMsg then PGIError(TexteMessage[NumMsg+7], 'Erreur Saisie');
              Result := '';
            end;
      End;
    end;

  end;

end;

function CodeMaterielExist (CodeMateriel : string; WhereSql : string=''; Obligatoire : Boolean=True ) : boolean;
var SQl : string;
begin

  Result := False;

  if (Obligatoire) and (CodeMateriel = '') then Exit;

  if (CodeMateriel = '') then 
  begin
     Result := True;
     exit;
  end;
  Sql := 'SELECT 1 FROM BTMATERIEL WHERE BMA_CODEMATERIEL="'+CodeMateriel+'"';
  if WhereSql <> '' then
  begin
    SQL := SQL + ' AND '+WhereSql;
  end;

  Result := ExisteSQL(SQl);

end;

function TypeActionExist (typeAction : string; WhereSql : string=''; Obligatoire : Boolean=True) : boolean;
var SQl : string;
begin

  Result := False;

  if (Obligatoire) and (typeAction = '') then Exit;

  if typeAction = '' then
  begin
    Result := True;
    Exit;
  end;

  Sql := 'SELECT 1 FROM BTETAT WHERE BTA_BTETAT="'+typeAction+'"';
  if WhereSql <> '' then
  begin
    SQL := SQL + ' AND '+WhereSql;
  end;
  Result := ExisteSQL(SQl);
end;

function EventMatExists (CodeMateriel,CodeEvt,CodeAffaire : string) : boolean;
var SQl : string;
begin
  Sql := 'SELECT 1 FROM BTEVENTMAT WHERE BEM_CODEMATERIEL="'+CodeMateriel+'" AND '+
         'BEM_BTETAT="'+CodeEvt+'" AND BEM_CODEETAT ="ARE" AND ';
  if CodeAffaire <> '' then Sql := Sql + 'BEM_AFFAIRE="'+CodeAffaire+'" AND ';
  Sql := Sql + 'BEM_REFPIECE=""';
  Result := ExisteSQL(Sql);
end;


function TiersExist (Tiers : string; WhereSql : string='' ) : boolean;
var SQl : string;
begin

  Result := True;

  if Tiers = '' then Exit;

  Sql := 'SELECT 1 FROM TIERS WHERE T_TIERS="'+Tiers+'"';
  if WhereSql <> '' then
  begin
    SQL := SQL + ' AND '+ WhereSql;
  end;
  Result := ExisteSQL(SQl);

end;

function FamilleMatExist (CodeFam : string; WhereSql : string='' ) : boolean;
var SQl   : string;
begin

  Result  := True;

  if CodeFam = '' then Exit;

  Sql := 'SELECT 1 FROM BTFAMILLEMAT WHERE BFM_CODEFAMILLE="'+CodeFam+'"';
  if WhereSql <> '' then
  begin
    SQL := SQL + ' AND '+ WhereSql;
  end;

  Result := ExisteSQL(SQl);

end;

function PrestationExist (CodeArticle : string; WhereSql : string='' ) : boolean;
var SQl   : string;
begin
  Result  := True;
  if CodeArticle = '' then Exit;
  Sql := 'SELECT 1 FROM ARTICLE LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES ';
  SQl := Sql + ' WHERE (GA_CODEARTICLE="' + CodeArticle + '")';
  SQl := Sql + '    OR (GA_ARTICLE="' +  CodeArticle + '") ';
  if WhereSql <> '' then
  begin
    SQL := SQL + ' AND '+ WhereSql;
  end;
  Result := ExisteSQL(SQl);

end;


function RessourceExist (CodeRessource : string; WhereSql : string='' ) : boolean;
var SQl : string;
begin

  Result := True;
  if CodeRessource = '' then Exit;
  Sql := 'SELECT 1 FROM RESSOURCE WHERE ARS_RESSOURCE="'+CodeRessource+'"';
  if WhereSql <> '' then
  begin
    SQL := SQL + ' AND '+ WhereSql;
  end;
  Result := ExisteSQL(SQl);


end;

function AffaireExist (Affaire : string; WhereSql : string='' ) : boolean;
var StSQl : string;
    Tempo : String;
begin

  Result := True;

  if Affaire = '' then Exit;

  Tempo  := Copy(Affaire,2,Length(Affaire)-1);
  Tempo  := Trim(Copy(Tempo,1, Length(Affaire)-3));

  if Tempo = '' then exit;

  StSql := 'SELECT 1 FROM AFFAIRE WHERE AFF_AFFAIRE="'+Affaire+'"';
  if WhereSql <> '' then
  begin
    StSQL := StSQL + ' AND '+ WhereSql;
  end;

  Result := ExisteSQL(StSQl);

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 09/10/2015
Modifié le ... :   /  /    
Description .. : Calcul du coût de Revient Matériel
Mots clefs ... : 
*****************************************************************}
Procedure CalculCout(Tobcout : Tob);
Var CodeCalendrier  : String;
    CodeRessource   : String;
    StSQL           : String;
    //
    Nbjours         : Integer;
    NbJourouvre     : Integer;
    //
    NbJourConso     : Double;
    NbHeureConso    : Double;
    NbPercentConso  : Double;
    //
    NbJourImmo      : Double;
    NbHeureImmo     : Double;
    NbPercentImmo   : Double;
    //
    NbJourEnt       : Double;
    NbHeureEnt      : Double;
    NbPercentEnt    : Double;
    //
    NbJourNoUse     : Double;
    NbHeureNoUse    : Double;
    NbPercentNoUse  : Double;
    //
    Periode         : Double;
    NbHeure         : Double;
    //
    amortissement   : Double;
    CoutEntretien   : Double;
    TotalCoutPeriode: Double;
    //
    CoutRevientJ    : Double;
    CoutRevientH    : Double;
    //
    Borne1, Borne2  : TDateTime;
    //
    Jour1           : Word;
    Mois1           : Word;
    Annee1          : Word;
    Jour2           : Word;
    Mois2           : Word;
    Annee2          : Word;
    //
    H,Min,Sec,Msec  : Word;
    //
    DateRef         : Tdatetime;
    CalendrierSpecif: Boolean;
    QQ              : TQuery;
    //
    Typeentretien   : string;
    CodeMateriel    : string;
begin


  CodeRessource     := TobCout.GetString('RESSOURCE');
  CodeMateriel      := TobCout.GetString('CODEMATERIEL');

  //Récupération du nombre d'heure travaillée moyenne par jour
  CodeCalendrier    := TobCout.GetString('STANDCALEN');
  CalendrierSpecif  := TobCout.GetBoolean('CALENSPECIF');

  Borne1 := Tobcout.GetDateTime('BORNE1');
  Borne2 := Tobcout.GetDateTime('BORNE2');

  CodeMateriel := Tobcout.GetString('CODEMATERIEL');

  //Calcul du nombre de jour ouvré
  DecodeDate(Borne1, Annee1, Mois1, Jour1);
  DecodeDate(Borne2, Annee2, Mois2, Jour2);

  DateRef := Borne1;

  NbJourOuvre := 0;

  //Contrôle si la ressource dispose d'un calendrier spécifique
  If CalendrierSpecif then
    NbHeure := RecupTempsMoyen(CodeCalendrier, CodeRessource)
  else
    NbHeure := RecupTempsMoyen(CodeCalendrier, '');

  if NbHeure = 0 then NbHeure := ChargeParametreTemps;

  if NbHeure = 0 then NbHeure := StrToTime('07:30:00');

  While DateRef <= Borne2 do
  begin
    if DateEstOuvre(DateRef,CodeCalendrier, CodeRessource) then Inc(NbJourOuvre);
    DateRef := IncDay(DateRef, 1);
  end;

  //Calcul de la période en heure
  //On ramène tout en minute, puis en heures
  DecodeTime(NbHeure,H,Min,Sec,MSec);
  Min := (H*60) + Min;
  Periode := NbJourOuvre * Min;

  //Recherche des lignes conso sur la période poour calcul Nb jours utilisation
  //StSQl := 'SELECT COUNT(BCO_DATEMOUV) AS NBJOURCONSO FROM CONSOMMATIONS';
  //StSQL := StSQL + ' WHERE BCO_RESSOURCE = "' + CodeRessource;
  //StSQL := StSQL + '"  AND (BCO_DATEMOUV BETWEEN "' + USDATETIME(Borne1) + '" AND "' + USDATETIME(Borne2) + '") ';

  StSQl := 'SELECT SUM(BEM_NBHEURE) AS NBHEURECONSO FROM BTEVENTMAT ';
  StSQL := StSQL + 'WHERE  BEM_CODEMATERIEL = "' + CodeMateriel;
  StSQL := StSQL + '"  AND (BEM_DATEDEB    >= "' + USDATETIME(Borne1) + '") ';
  StSQL := StSQL + '   AND (BEM_DATEFIN    <= "' + USDATETIME(Borne2) + '") ';
  QQ := OpenSQL(StSQL, False);
  if QQ.Eof then
    NbHeureConso := 0
  else
    NbHeureConso := QQ.findfield('NBHEURECONSO').AsInteger;

  Ferme(QQ);

  //tranbsformation en Minutes
  NbHeureConso   := NbHeureConso * 60;
  NBJourconso    := Arrondi((NbHeureConso / Min),2);
  NbPercentConso := Arrondi((NbJourConso * 100)/NbJourOuvre,0);

  //Calcul des lignes immobilisation
  NbHeureImmo    := Periode - Nbheureconso;
  NbJourImmo     := NbJourOuvre - NbJourConso;
  NbPercentImmo  := Arrondi((NbJourimmo * 100)/NbJourOuvre,0);

  //On récupère le type d'action paramétré pour l'entretien
  TypeEntretien  := GetParamSocSecur('SO_ACTIONENTRETIEN','');
  TypeEntretien  := FormatMultiValComboForSQL(TypeEntretien);

  //Recherche des lignes événements pour chargement des événement de type entretien
  if Typeentretien <> '' then
  begin
    StSql := 'SELECT SUM(BEM_NBHEURE) AS NBHEUREENT FROM BTEVENTMAT';
    StSQL := STSQL + ' WHERE BEM_DATEDEB >= "'  + USDATETIME(Borne1) + '"';
    StSQL := StSQL + '   AND BEM_DATEFIN <= "'  + USDATETIME(Borne2) + '"';
    StSQL := StSQL + '   AND BEM_BTETAT in ('   + TypeEntretien      + ')';
    StSQL := StSQL + '   AND BEM_CODEMATERIEL="'+ CodeMateriel       + '"';
    QQ := OpenSQL(StSQL, False);
    if QQ.Eof then
      NBHeureEnt := 0
    else
      NbHeureEnt := QQ.findfield('NBHEUREENT').AsInteger;
    Ferme(QQ);

    NbJourEnt    := Arrondi((NbHeureEnt/NbHeure),2);
    NbPercentent := Arrondi((NbJourEnt * 100)/NbJourOuvre,0);
  end
  else
  begin
    NbHeureEnt   := 0;
    NbJourEnt    := 0;
    NbPercentent := 0;
  end;

  //Calcul des zones non Utilisée
  NbHeureNoUse   := NbHeureImmo - NbheureEnt;
  NbJourNoUse    := Arrondi(NbJourImmo - NbJourEnt,2);
  NbPercentNoUse := Arrondi((NbJourNoUse * 100)/NbJourOuvre,0);

  //Chargement de l'amortissement au prorata de la période
  Nbjours       := DaysBetween(Borne2, Borne1) + 1;
  Amortissement := TobCout.GetDouble('AMORTISSEMENT');
  if Nbjours <> DaysInYear(Annee1) then
  begin
    Amortissement := (Amortissement/DaysInYear(Annee1)) * NbJourouvre;
    Amortissement := Arrondi(amortissement, 2);
  end;

  //calcul du coût d'entretien : Somme des facture d'achat sur la période
  StSQL := 'SELECT SUM(GP_TOTALHTDEV) AS COUTENTRETIEN FROM PIECE ';
  StSQL := StSQL + ' WHERE GP_NATUREPIECEG = "BCM" ';
  StSQL := StSQL + ' AND GP_CODEMATERIEL="' + CodeMateriel + '" ';
  StSQL := StSQL + ' AND (GP_DATEPIECE BETWEEN "' + USDATETIME(Borne1) + '" AND "' + USDATETIME(Borne2) +'") ';
  QQ := OpenSQL(StSQL, False);
  if QQ.Eof then
    CoutEntretien := 0
  else
    CoutEntretien := QQ.findfield('COUTENTRETIEN').AsInteger;
  Ferme(QQ);

  //Total Coût de la période
  TotalCoutPeriode := Amortissement + CoutEntretien;

  //coût de revient journalier et heure
  CoutRevientJ := Arrondi(TotalCoutPeriode / NbJourOuvre,2);
  CoutRevientH := Arrondi(TotalCoutPeriode / Periode,2);

  //On repasse tout en heure avant l'affichage à l'écran
  Periode       := (Periode/24/60)*24;
  NbHeureConso  := (NbHeureConso/24/60)*24;
  NbHeureImmo   := (NbHeureImmo/24/60)*24;
  NbHeureNoUse  := (NbHeureNoUse/24/60)*24;

  //Chargement de TobCout dans les Zones Ecran
  TobCout.AddChampsupValeur('PeriodeHeure',     Periode);
  TobCout.AddChampsupValeur('PeriodeJour',      NbJourouvre);
  TobCout.AddChampsupValeur('UtilHeure',        NbHeureConso);
  TobCout.AddChampsupValeur('Utiljour',         NbJourConso);
  TobCout.AddChampsupValeur('UtilPercent',      NbPercentConso);
  TobCout.AddChampsupValeur('ImmoHeure',        NbHeureImmo);
  TobCout.AddChampsupValeur('ImmoJour',         NbJourImmo);
  TobCout.AddChampsupValeur('ImmoPercent',      NbPercentImmo);
  TobCout.AddChampsupValeur('EntretienHeure',   NbHeureEnt);
  TobCout.AddChampsupValeur('entretienJour',    NbJourent);
  TobCout.AddChampsupValeur('entretienPercent', NbPercentent);
  TobCout.AddChampsupValeur('NotUseHeure',      NbHeureNoUse);
  TobCout.AddChampsupValeur('NotUseJour',       NbJourNoUse);
  TobCout.AddChampsupValeur('NotUsePercent',    NbPercentNoUse);
  TobCout.AddChampsupValeur('Amortissement',    Amortissement);
  TobCout.AddChampsupValeur('CoutEntretien',    CoutEntretien);
  TobCout.AddChampsupValeur('TotalCoutPeriode', TotalCoutPeriode);
  TobCout.AddChampsupValeur('CtRevientJour',    CoutRevientJ);
  TobCout.AddChampsupValeur('CtRevientHeure',   CoutRevientH);

end;

Function RecupTempsMoyen(CodeCalendrier, CodeRessource : string) : Double;
Var StSQL : String;
    QQ    : TQuery;
begin

  Result := 0;

  //Lecture du calendier de la ressource
  if CodeCalendrier = '' then Exit;

  //Lecture du calendrier spécifique à la Ressource
  StSQL := 'SELECT SUM(ACA_DUREE) AS NBHEURE FROM CALENDRIER WHERE ACA_STANDCALEN="' + CodeCalendrier + '" ';
  if CodeRessource <> '' then StSQl := StSQL + ' AND ACA_RESSOURCE="' + CodeRessource + '"';

  QQ := OpenSQL(StSQl, false);

  if not QQ.Eof then  Result := QQ.Findfield('NBHEURE').AsFloat;

  Ferme(QQ);

  //Calcul de la période
  //Détermination des jours non travaillés
  Result := Result/DetermineJourTravaille(CodeCalendrier, CodeRessource);

end;

function DetermineJourTravaille(CodeCalendrier, CodeRessource : string) : Integer;
Var TobCal  : TOB;
    StSQL   : string;
    QQ      : TQuery;
    ind     : Integer;
    Duree   : Double;
begin

   //lecture du Calendrier pour Vérification si Samedi et Dimanche Travaillé
  //Ou si Ressource travaille les jours Fériés
  Result := 0;

  if CodeCalendrier = '' then Exit;

  StSQL := 'SELECT ACA_JOUR, ACA_DUREE, ACA_FERIETRAVAIL FROM CALENDRIER ';
  StSQL := StSQL + ' WHERE ACA_STANDCALEN="' + CodeCalendrier + '" ';
  if CodeRessource <> '' then StSQl := StSQL + ' AND ACA_RESSOURCE="' + CodeRessource + '"';

  QQ := OpenSQL(StSQl, false);

  if not QQ.Eof then
  begin
    TobCal := TOB.create('CALENDRIER', nil, -1);
    TobCal.SelectDB('JOURS', QQ);

    For ind := 0 to TobCal.detail.count -1 do
    begin
      Duree   := StrToFloat(TobCal.detail[Ind].GetValue('ACA_DUREE'));
      if duree <> 0 then inc(Result);
    end;

    FreeAndNil(TOBCal);
  end;

  Ferme(QQ);

end;

Function ChargeParametreTemps : Double;
Var HeureDebAM : TTime;
    HeurefinAM : TTime;
    HeureDebPM : TTime;
    HeureFinPM : TTime;
    TpsMatin   : Double;
    TpsAprem   : Double;
begin

  //Result := 0;
  //Faudrait peut-être tenir compte du calendrier Ressource (???)


  HeureDebAM := GetParamSocSecur('SO_BTAMDEBUT', '08:30');
  HeureFinAM := GetParamSocSecur('SO_BTAMFIN',   '12:30');
  HeureDebPM := GetParamSocSecur('SO_BTPMDEBUT', '14:00');
  HeureFinPM := GetParamSocSecur('SO_BTPMFIN',   '17:30');

  //Calcul du nombre heure matin + nombre heure Après-midi
  TpsMatin   := HeureFinAM - HeureDebAM;
  TpsAprem   := HeureFinPM - HeureDebPM;

  Result     := TpsMatin + TpsAprem;

end;
{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Vérifie si une date/Calendrier ouvrée
Mots clefs ... : CALENDRIER;FERME;OUVRE
*****************************************************************}
function DateEstOuvre(aDate: TDateTime; CodeCalendrier, CodeRessource : String): Boolean;
begin
  Result := False;

  //Recherche si jour non travaillé
  // Dimanche = 1
  if JourNonTravaille(CodeCalendrier, CodeRessource, DayOfWeek(aDate)) then Exit;
  if isJourFerie(aDate) then exit;

  Result := True; // Ouvrable

end;

Function JourNonTravaille(CodeCalendrier, codeRessource : string; NumJour : Integer) : Boolean;
Var StSql : string;
    QQ    : TQuery;
    StParam : String;
begin

  Result := False;

  //si on est sur un dimanche on force à 7 le Numéro du jour
  if numjour = 1 then
    StParam := 'SO_JOURFERMETURE'
  else
    StParam := 'SO_JOURFERMETURE' + IntToStr(Numjour -1);

  if CodeCalendrier = '' then
    //recherche dans les paramétres societé
    Result := GetParamSocSecur(StParam, False)
  else
  begin
    //recherche dans le calendrier de la ressource
    StSQl := 'Select ACA_DUREE FROM CALENDRIER WHERE ACA_STANDCALEN="' + CodeCalendrier + '" ';
    StSql := StSql + ' AND ACA_JOUR=' + IntToStr(NumJour);
    if CodeRessource  <> '' then StSQl := StSQL + ' AND ACA_RESSOURCE="' + CodeRessource + '"';
    QQ := OpenSQL(StSQL, False);
    if not QQ.eof then
    begin
      if QQ.FindField('ACA_DUREE').Asfloat = 0 then  Result := true;
    end;
    Ferme(QQ);
  end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 08/10/2015
Modifié le ... :   /  /
Description .. : Génération de la pièce "BCM" à partir d'un événement
Suite ........ : Et génération des consommations asociées
Mots clefs ... :
*****************************************************************}
Procedure CreateEnteteDoc(TobPiece, TOBL : TOB);
Var DateMouv : TDateTime;
begin

  DateMouv := StrToDate(TOBL.GetString('BEM_DATEFIN'));

  TobPiece.PutValue ('NATUREPIECEG', 'BCM');
  TobPiece.PutValue ('AFFAIRE',      TOBL.GetString('BEM_AFFAIRE'));
  TobPiece.PutValue ('TIERS',        TOBL.GetString('BEM_TIERS'));
  TobPiece.PutValue ('BTETAT',       TOBL.GetString('BEM_BTETAT'));
  TobPiece.PutValue ('RESSOURCE',    TOBL.GetString('BEM_RESSOURCE'));
  TobPiece.PutValue ('CODEMATERIEL', TOBL.GetString('BEM_CODEMATERIEL'));

  TobPiece.PutValue ('ETABLISSEMENT',TOBL.GetString('AFF_ETABLISSEMENT'));
  TobPiece.PutValue ('DOMAINE',      TOBL.GetString('AFF_DOMAINE'));
  TobPiece.PutValue ('DATEPIECE',    DateMouv);
  TobPiece.PutValue ('REFINTERNE',   '');
  TobPiece.PutValue ('CREEPAR',      'Coûts Parc/Matériel');
  TobPiece.PutValue ('ORIGINE',      'PARC/MAT');

  CreateLigneCom(TobPiece, TOBL);

end;

procedure CreateLigneCom(TobPiece, TOBL : TOB);
Var TobLigne : TOB;
begin

  // Création TOB pour ligne de pièce
  TobLigne := Tob.Create ('LES LIGNES', TobPiece, -1);
  //
  TobLigne.AddChampSupValeur('TYPELIGNE',   'COM');
  TobLigne.AddChampSupValeur('CODEARTICLE', '');
  TobLigne.AddChampSupValeur('ARTICLE',     '');
  TobLigne.AddChampSupValeur('RESSOURCE',     TOBL.GetString('BEM_RESSOURCE'));

  TobLigne.AddChampSupValeur('LIBELLE',     'Evénement Parc/Matériel N°' + TOBL.GetString('BEM_IDEVENTMAT'));

  TobLigne.AddChampSupValeur('QTEFACT',      0.00);
  TobLigne.AddChampSupValeur('QUALIFQTE',   '');
  TobLigne.AddChampSupValeur('PUHTDEV',      0.00);
  TobLigne.AddChampSupValeur('PAHT',         0.00);
  TobLigne.AddChampSupValeur('DPA',          0.00);
  TobLigne.AddChampSupValeur('PMAP',         0.00);
  TobLigne.AddChampSupValeur('DATELIVRAISON',idate1900);

  TobLigne.AddChampSupValeur('DEPOT', GetParamSocSecur('SO_GCDEPOTDEFAUT', '', False))

end;

procedure CreateLigneDoc(TobPiece, TOBL : TOB);
Var TobLigne  : TOB;
    Qte       : Double;
begin

    //création des lignes qui vont bien....
    TobLigne := TOB.Create('LES LIGNES', TOBPiece,-1);

    TobLigne.AddChampSupValeur('TYPELIGNE', 'ART');
    TobLigne.AddChampSupValeur('AFFAIRE',       TOBL.GetString('BEM_AFFAIRE'));
    TobLigne.AddChampSupValeur('CODEARTICLE',   TOBL.GetString('GA_CODEARTICLE'));
    TobLigne.AddChampSupValeur('ARTICLE',       TOBL.GetString('GA_ARTICLE'));
    TobLigne.AddChampSupValeur('LIBELLE',       TOBL.GetString('BEM_LIBACTION'));
    TobLigne.AddChampSupValeur('LIBCOMPL',      TOBL.GetString('BEM_LIBREALISE'));
    TobLigne.AddChampSupValeur('DATELIVRAISON', TOBL.GetString('BEM_DATEFIN'));
    TobLigne.AddChampSupValeur('RESSOURCE',     TOBL.GetString('BEM_RESSOURCE'));
    TobLigne.AddChampSupValeur('CODEMATERIEL',  TOBL.GetString('BEM_CODEMATERIEL'));
    TobLigne.AddChampSupValeur('BTETAT',        TOBL.GetString('BEM_BTETAT'));
    TobLigne.AddChampSupValeur('NBHEURE',       TOBL.GetString('BEM_NBHEURE'));

    Qte := TOBL.GetDouble('BEM_NBHEURE');

    TobLigne.AddChampSupValeur('QTEFACT',     Qte);
    TobLigne.AddChampSupValeur('QUALIFQTE',   'FT');
    TobLigne.AddChampSupValeur('PUHTDEV',     TOBL.GetDouble('BEM_DPA'));
    TobLigne.AddChampSupValeur('PAHT',        TOBL.GetDouble('BEM_DPA'));
    TobLigne.AddChampSupValeur('DPA',         TOBL.GetDouble('BEM_DPA'));
    TobLigne.AddChampSupValeur('PMAP',        TOBL.GetDouble('GA_PMAP'));

    TobLigne.AddChampSupValeur('DEPOT', GetParamSocSecur('SO_GCDEPOTDEFAUT', '', False));

    CreateLigneCompl(TOBL, TobLigne);

end;

Procedure CreateLigneCompl(TOBL, TobLigne : tob);
Var TobLigneC : TOB;
    Qte       : Double;
Begin

    //création des lignes Complément qui vont pas trop mal....
    TobLigneC := TOB.Create('LES LIGNES COMPL', TOBLigne,-1);

    TobLigneC.AddChampSupValeur('CODEMATERIEL',  TOBL.GetString('BEM_CODEMATERIEL'));
    TobLigneC.AddChampSupValeur('BTETAT',        TOBL.GetString('BEM_BTETAT'));
    TobLigneC.AddChampSupValeur('IDEVENTMAT',    TOBL.GetString('BEM_IDEVENTMAT'));

    Qte := TOBL.GetDouble('BEM_NBHEURE');

    TobLigneC.AddChampSupValeur('QTEAPLANIF',   Qte);

end;

Procedure CreateTOBPiece(TobPiece : TOb);
begin

  TobPiece.AddChampSupValeur ('NATUREPIECEG', 'BCM');
  TobPiece.AddChampSupValeur ('RESSOURCE',    '');
  TobPiece.AddChampSupValeur ('CODEMATERIEL', '');
  TobPiece.AddChampSupValeur ('BTETAT',       '');
  TobPiece.AddChampSupValeur ('AFFAIRE',      '');
  TobPiece.AddChampSupValeur ('TIERS',        '');
  TobPiece.AddChampSupValeur ('ETABLISSEMENT','');
  TobPiece.AddChampSupValeur ('DOMAINE',      '');
  TobPiece.AddChampSupValeur ('DATEPIECE',    '');
  TobPiece.AddChampSupValeur ('REFINTERNE',   '');
  TobPiece.AddChampSupValeur ('CREEPAR',      'Coûts Parc/Matériel');
  TobPiece.AddChampSupValeur ('ORIGINE',      'PARC/MAT');

end;

Procedure ReinitialiseTob(NomTob : Tob);
begin

  NomTob.ClearDetail;

  NomTob.InitValeurs;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 09/10/2015
Modifié le ... :   /  /    
Description .. : Procedure de création des conso à partir de la 
Suite ........ : Tobevénement Parc/Matériel
Mots clefs ... : 
*****************************************************************}

Procedure  TraitementRealise(TobEvent : TOB);
Var StSQL     : String;
    NumEvent  : Integer;
    DateImp   : TdateTime;
begin

  //On contrôle si le type d'action est géré en conso...
  if not GereConso(TobEvent.GetValue('BEM_BTETAT')) then
  begin
    //IError('Ce type d''action : '+ TobEvent.GetValue('BEM_BTETAT') + ' n''est pas géré en consommation !', 'Erreur Mise à jour consommations');
    Exit;
  end;

  //Si aucune affaire de saisie on fait rien
  if TobEvent.GetValue('BEM_AFFAIRE') = ''        then
  begin
    //IError('Le Code Affaire dans l''événement est obligatoire', 'Erreur Mise à jour consommations');
    Exit;
  end;

  NumEvent  := TobEvent.GetValue('BEM_IDEVENTMAT');
  DateImp   := StrToDate(TobEvent.GetValue('DATEIMPUTATION'));

  //Si tout est ok on peut faire la mise à jour des consos
  if TraitementConso(NumEvent, DateImp) then
  begin
    StSQL := 'UPDATE BTEVENTMAT ';
    StSQL := StSQL + 'SET BEM_NUMMOUV=(SELECT BCO_NUMMOUV FROM CONSOMMATIONS WHERE BCO_INTEGRNUMFIC = ' + IntToStr(NumEvent) + ') ';
    if TobEvent.GetValue('BEM_CODEETAT') <> 'REA' then StSQL := StSQL + ',BEM_CODEETAT="REA"';
    if TobEvent.GetValue('BEM_LIBREALISE') =''    then StSQL := StSQL + ',BEM_LIBREALISE="Réalisation de ' + TobEvent.GetValue('BEM_LIBACTION') + '" ';
    //
    StSQL := StSQL + 'WHERE BEM_IDEVENTMAT=' + IntToStr(NumEvent);
    if ExecuteSQL(StSQL)<> 1 then
    begin
      //La mise à jour de la table des événement ne c'est pas fait...
      //il faut remettre le code etat à 'En-cours'
      TobEvent.PutValue('BEM_CODEETAT', 'ECO');
      Tobevent.InsertOrUpdateDB(False);
      PGIError('La Mise à jour des consommations ne s''est pas effectuée', 'Erreur Mise à jour consommations');
      //il faut supprimer la conso associée
      StSQL := 'DELETE CONSOMMATION WHERE BCO_INTEGRNUMFIC=' + IntToStr(Numevent);
      if ExecuteSQL(StSQL)<> 1 then PGIError('Suppression de la consommation associée impossible', 'Erreur Mise à jour consommations');
    end
    else
  end;

end;

function TraitementConso(NumEventMat : Integer; DateImp : TDateTime) : Boolean;
Var StSQL         : string;
    QQ            : TQuery;
    TobTravail    : Tob;
begin

  Result := False;

  //Lecture de la table Affaire pour récupération de l'Etablissement et du domaine
  STSQL :=         'SELECT * FROM BTEVENTMAT ';
  STSQL := STSQL + '  LEFT JOIN BTMATERIEL  ON BEM_CODEMATERIEL=BMA_CODEMATERIEL ';
  STSQL := STSQL + '  LEFT JOIN ARTICLE     ON BMA_CODEARTICLE=GA_CODEARTICLE ';
  STSQL := STSQL + '  LEFT JOIN RESSOURCE   ON BMA_RESSOURCE=ARS_RESSOURCE ';
  StSQL := StSql + '  LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES ';
  STSQL := STSQL + '  LEFT JOIN AFFAIRE     ON BEM_AFFAIRE=AFF_AFFAIRE ';
  STSQL := STSQL + '  LEFT JOIN TIERS       ON AFF_TIERS=T_TIERS ';
  STSQL := STSQL + '  LEFT JOIN BTETAT      ON BEM_BTETAT=BTA_BTETAT ';
  StSQL := StsQL + ' WHERE BEM_IDEVENTMAT=' + IntToStr(NumEventMat);
  StSQL := StSQL + '   AND BTA_GESTIONCONSO = "X"'; //si le type d'action ne doit pas gérer les Conso on fait rien

  QQ    := OpenSQL(StSQL, false);

  if QQ.Eof then
  begin
    Ferme(QQ);
    Exit;
  end;

  TobTravail := Tob.Create('TRAVAIL',nil,-1);

  TobTravail.SelectDB(StSQL, QQ);
  TobTravail.AddChampSupValeur('DATEIMPUTATION', DateToStr(DateImp));
  Ferme(QQ);

  Result := CreateConsoByEventMat(TobTravail);

  FreeandNil(TobTravail);

end;

//les consommations se mettent à jour en fonction de la Ressource saisie en fiche Matériel
//le mouvement devient un mlouvement de type RES et pas de MO
Function CreateConsoByEventMat(TobEvent : TOB) : Boolean;
var TobConso : TOB;
    DateMouv : TDateTime;
    TypRes	 : string;
    Prix     : Double;
    NumConso : double;
    Annee    : word;
    Mois     : word;
    Jour     : word;
    Semaine  : integer;
    ValConsoPresta : Boolean;
begin

  Result := False;

  ValConsoPresta := GetParamSocSecur('SO_BTVALOSCONSO', False);

  //Les consos sont déjà mise à jour on peut donc pas faire de modif...
  if TobEvent.GetInteger('BEM_NUMMOUV') <> 0    then
  begin
    //IError('Il existe déjà une ligne consommation pour cet événement', 'Erreur Mise à jour consommations');
    Exit;
  end;

  if (GetNumUniqueConso (NumConso) = Gncabort)  then
  begin
    //IError('Problème de numérotation des lignes de consommations', 'Erreur Mise à jour consommations');
    exit;
  end;

	TobConso := TOB.Create ('CONSOMMATIONS',TobConso,-1);

  ReinitialiseTob(TobConso);

  TobConso.AddchampSupValeur('COEFMARGE', 0);

  TobConso.putValue('BCO_NUMMOUV', NumConso);

  //chargement et découpage du code affaire
  TobConso.putValue('BCO_AFFAIRE',        TobEvent.GetValue('BEM_AFFAIRE'));
  TobConso.putValue('BCO_AFFAIRE0',       TobEvent.GetValue('AFF_AFFAIRE0'));
  TobConso.putValue('BCO_AFFAIRE1',       TobEvent.GetValue('AFF_AFFAIRE1'));
  TobConso.putValue('BCO_AFFAIRE2',       TobEvent.GetValue('AFF_AFFAIRE2'));
  TobConso.putValue('BCO_AFFAIRE3',       TobEvent.GetValue('AFF_AFFAIRE3'));
  TobConso.putValue('BCO_AVENANT',        TobEvent.GetValue('AFF_AVENANT'));

  TobConso.putValue('BCO_AFFAIRESAISIE',  TobEvent.GetValue('BEM_AFFAIRE'));

  TobConso.putValue('BCO_PHASETRA',     '');

  // Découpage de la date
  // Recherche du Numéro de semaine
  //DateMouv:=StrToDate(TobEvent.GetValue('BEM_DATEFIN'));
  DateMouv:=StrToDate(TobEvent.GetValue('DATEIMPUTATION')); //=> chargement de label date d'imputation

  DecodeDate(DateMouv, Annee, Mois, Jour);

  Semaine := NumSemaine(DateMouv);

  TobConso.putValue('BCO_DATEMOUV',     DateMouv);
  TobConso.PutValue('BCO_SEMAINE',      Semaine);
  TobConso.PutValue('BCO_MOIS',         Mois);

  TobConso.putValue('BCO_NATUREMOUV',   'RES');
  TobConso.putValue('BCO_RESSOURCE',    TobEvent.GetValue('BMA_RESSOURCE'));

  // Si le code article n'est pas renseigné, on récupère la prestation par défaut de la ressource
  TobConso.putValue('BCO_CODEARTICLE',  TobEvent.GetValue('GA_CODEARTICLE'));
  TobConso.putValue('BCO_ARTICLE',      TobEvent.GetValue('GA_ARTICLE'));
  if TobConso.GetValue('BCO_CODEARTICLE') = '' then TobConso.PutValue('BCO_ARTICLE', TobEvent.GetValue('ARS_ARTICLE'));

  TobConso.putValue('BCO_LIBELLE',      TobEvent.GetValue('BEM_LIBREALISE'));
  TobConso.putValue('BCO_QUANTITE',     TobEvent.GetValue('BEM_NBHEURE'));
  TobConso.putValue('BCO_QTEFACTUREE',  TobEvent.GetValue('BEM_NBHEURE'));
  TobConso.putValue('BCO_QTEINIT',      TobEvent.GetValue('BEM_NBHEURE'));

  TobConso.putValue('BCO_QUALIFQTEMOUV','H');

  Prix := TobEvent.GetDouble('BEM_PA');
  TobConso.putValue('BCO_DPA',          Prix);

  Prix := TobEvent.GetDouble('BEM_PR');
  TobConso.putValue('BCO_DPR',          Prix);

  TobConso.putValue('BCO_TYPEHEURE',    '');

  TobConso.PutValue('BCO_FACTURABLE',   'N');
  TobConso.PutValue('BCO_VALIDE',       'X');

  // récupération du type de resssource
  TypRes:= TobEvent.GetValue('BNP_TYPERESSOURCE');

  if ValConsoPresta then
  begin
    if (TobConso.getValue('BCO_DPA')=0)    or
       (TobConso.GetValue('BCO_DPR')=0)    or
       (TobConso.GetValue('BCO_PUHT')=0) then GetValoPrestation(TobConso)
  end
  else
  begin
    if (TobConso.getValue('BCO_DPA')=0)  then SetValoFromRessource (TobConso,TobEvent,[TmvPa]);
    if (TobConso.GetValue('BCO_DPR')=0)  then SetValoFromRessource (TobConso,TobEvent,[TmvPr]);
    if (TobConso.GetValue('BCO_PUHT')=0) then SetValoFromRessource (TobConso,TobEvent,[TmvPV])
  end;

  if TobConso.GetValue('COEFMARGE') = 0 then TobConso.PutValue('COEFMARGE', TobEvent.GetValue('GA_COEFMARGE'));

  TobConso.putValue('BCO_INTEGRNUMFIC', TobEvent.GetValue('BEM_IDEVENTMAT'));

  CalculeLaLigne (TobConso);

	if not TOBConso.InsertOrUpdateDB(False) then V_PGI.ioError := oeUnknown;

  if V_PGI.ioError = OeOK then Result := True;

  FreeAndNil(TobConso);

end;

Function GereConso(BtEtat : string) : Boolean;
var StSQL : string;
    QQ    : TQuery;
begin

  result := False;

  StSQL := 'SELECT BTA_GESTIONCONSO FROM BTETAT WHERE BTa_BTETAT="' + Btetat + '"';
  QQ    := OpenSQL(StSQL, False);

  if not QQ.eof then
  begin
    if QQ.FindField('BTA_GESTIONCONSO').AsString = 'X' then
      Result := True
    else
      Result := False;
  end;

  Ferme(QQ);

end;

Function CalculNbHeure(DDeb, Dfin : TdateTime) : String;
Var Nbjour    : Integer;
    //
    HDeb        : Tdatetime;
    HFin        : Tdatetime;
    NbHJour     : TDateTime;
    //
    Hour      : Word;
    Min       : Word;
    Sec       : Word;
    MSec      : Word;
begin

  Result := '0';

  HDeb    := GetParamsocSecur('SO_HEUREDEB', 0);
  HFin    := GetParamsocSecur('SO_HEUREFIN', 0);
  NbHJour := HFin-Hdeb;

  if Dfin = idate2099 then
    Result := '0'
  else
  begin
    DecodeTime(NBHJour, Hour, Min, Sec, Msec);

    //on recherche le nomnbre de jour entre deux dates
    Nbjour       := DaysBetween(DFin, DDeb);
    Nbjour := Nbjour; // + 1;

    if Nbjour = 0 then Result := IntToStr(Hour)
    Else Result := IntToStr(Nbjour * Hour);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 20/01/2016
Modifié le ... :   /  /
Description .. : Contrôle de cohérence de saisie sur un évènement
Suite ........ :
Mots clefs ... :
*****************************************************************}
Function ControleCoherence(DateDeb,DateFin : TDateTime; Duree : double; Materiel : string; NoEnCours : Integer) : boolean;
Var StSQL   : string;
    QQ      : TQuery;
begin

	result := true;
  //
  //Vérification si jour férié
  if isJourFerie(DateDeb) then
  begin
  	pgiInfo ('Impossible : La date de début est un jour Férié');
    result := false;
    exit;
  end;

  //Vérification si date de début correspond à un jour férié ou non travaillé
  if IsDateClosed (DateDeb) then
  begin
  	pgiInfo ('Attention : La date de début est un jour non travaillé');
    result := false;
    exit;
  end;
  //
  //Vérification si date de fin correspond à un jour férié ou non travaillé
  if IsDateClosed (DateFin) then
  begin
    pgiInfo ('Attention : La date de fin est un jour non travaillé');
    result := false;
    exit;
  end;
  //
  if isJourFerie(Datefin) then
  begin
  	pgiInfo ('Impossible : La date de fin est un jour Férié');
    result := false;
    exit;
  end;
  //
  //Vérification si matériel pas déjà sur la période
  StSQL := 'SELECT BEM_IDEVENTMAT FROM BTEVENTMAT ';
  StSQL := StSQL + ' WHERE  BEM_CODEMATERIEL="' + Materiel + '" ';
  StSQL := StSQL + '   AND (BEM_DATEDEB BETWEEN "' + USDATETIME(DateDeb) + '" ';
  StSQL := StSQL + '   AND "' + USDATETIME(DateFin) + '" ';
  StSQL := StSQL + '    OR BEM_DATEFIN BETWEEN "' + USDATETIME(DateDeb) + '" ';
  StSQL := StSQL + '   AND "' + USDATETIME(DateFin) + '")';
  //
  QQ := OpenSQL(STSQL, False,-1,'',True);
  if not QQ.eof then
  begin
    if (NoEnCours = 0) then
    begin
      pgiInfo ('Impossible : Un évènement est déjà prévu sur cette période pour ce matériel');
      Result := True;
    end
    else
    begin
      if (QQ.FindField('BEM_IDEVENTMAT').AsInteger <> NoEnCours) then
      begin
        pgiInfo ('Impossible : Un évènement est déjà prévu sur cette période pour ce matériel');
        Result := True;
      end;
    end;
  end;

  Ferme(QQ);
  //
  //Il faudrait peut-être un jour faire des contrôle sur le calendrier de la ressource associée au Matos...
  //Si l'on doit le faire cela se passera ici
  //
end;

//Recherche du calendrier associé à la ressource
Procedure RechercheCalendrier(CodeRes : string; TobCalendrier : Tob);
Var StSQL : String;
    QQ    : TQuery;
begin

  TobCalendrier.ClearDetail;

  //On Cherche si un Calendrier standard est associé à cette ressource
  StSQL := 'SELECT ARS_STANDCALEN, ARS_CALENSPECIF FROM RESSOURCE WHERE ARS_RESSOURCE = "' + CodeRes + '" AND ARS_TYPERESSOURCE in ("MAT", "SAL", "OUT")';
  QQ    := OpenSQL(StSQL, False);

  If not QQ.eof then
  Begin
    if QQ.Findfield('ARS_STANDCALEN').AsString <> '' then
    begin
      //Si coche spécifique dans fiche ressource on cherche un calendrier spécifique à la ressource
      if QQ.FindField('ARS_CALENSPECIF').AsBoolean then
        StSQL := 'SELECT * FROM CALENDRIER WHERE ACA_RESSOURCE = "' + CodeRes + '" AND ACA_STANDCALEN="' + QQ.Findfield('ARS_STANDCALEN').AsString + '"'
      else
        //Sinon on charge le calendrier standard
        StSQL := 'SELECT * FROM CALENDRIER WHERE ACA_RESSOURCE = "***" AND ACA_STANDCALEN = "' + QQ.Findfield('ARS_STANDCALEN').AsString + '" ';
      QQ    := OpenSQL(StSQL, False);
      if not QQ.eof then
      begin
        TobCalendrier.LoadDetailFromSQL(StSQL,False,true);
      end;
    end;
  end;

  Ferme(QQ);

end;

//Recherche du calendrier associé à la ressource
Procedure ChargeJoursFeries(TobJferie : TOB);
Var StSQL : String;
    QQ    : TQuery;
begin

  if TobJferie = nil then Exit;

  StSQL := 'SELECT * FROM JOURFERIE';
  QQ    := OpenSQL(StSQL, False);

  If not QQ.eof then
  begin
    TobJFerie.LoadDetailFromSQL(StSQL,False,true);
  end;

  Ferme(QQ);

end;


function CalculDureeEvent(Borne1, Borne2 : TDateTime; TobJFerie, TobCalendrier : TOB) : Double;
Var DateRef : TdateTime;
    Jour    : Integer;
    TobLCal : TOB;
    Duree   : Double;
    H,M,S,Msec : Word;
    StParam : String;
Begin

  Result := 0;

  DateRef := Borne1;

  if Borne2 = idate2099 then
  begin
    Result := 0;
    Exit;
  end;

  While DateRef <= Borne2 do
  begin
    Duree := 0;
    //On vérifie Si nous sommes sur un jour férié
    if not JourFeries(TobJFerie, DateRef) then
    begin
      //Détermination du Numéro du jour
      Jour := DayOfTheWeek(DateRef);
      //si on est sur un dimanche on force à 7 le Numéro du jour car paramètre commence à 1 : dimanche => 7 : Lundi (Méthode Anglaise)
      if Jour = 7 then
        StParam := 'SO_JOURFERMETURE'
      else
        StParam := 'SO_JOURFERMETURE' + InttoStr(Jour);
      //On vérifie dans les paramètres société si c'est un jour travaillé
      if Not GetParamSocSecur(StParam, False) then
      begin
        //Inc(NbJourOuvre); //on cumule pour déterminer le nombre de jours ouvrés
        //on vérifie si la ressource comporte un calendrier, Spécifique ou non...
        if TobCalendrier <> nil then
        begin
          if TobCalendrier.detail.count > 0 then
          begin
            TobLCal := TobCalendrier.FindFirst(['ACA_JOUR'],[Jour],False);
            if TobLCal <> nil then
            begin
              Duree   := TobLCal.GetValue('ACA_DUREE')
            end
            else
            begin
              Duree := ChargeParametreTemps; //on charge la durée en fonction des paramètres société
              //Si on n'a pu déterminer une durée on force à une durée standard 8
              if Duree = 0 then Duree := 8;
            end;
          end
          else
          begin
            Duree := ChargeParametreTemps; //on charge la durée en fonction des paramètres société
            //Si on n'a pu déterminer une durée on force à une durée standard 8
            if Duree = 0 then Duree := 8;
          end;
        end
        else
        begin
          Duree := ChargeParametreTemps; //on charge la durée en fonction des paramètres société
          //Si on n'a pu déterminer une durée on force à une durée standard 8
          if Duree = 0 then Duree := 8;
        end;
      end;
      DecodeTime(duree, H, M, S, MSec);
      Duree := (H*60) + M;
      Result  := Result + Round(Duree); //on multiplie par 60 pour avoir un nombre de minutes...
    end;
    DateRef := IncDay(DateRef, 1); //On passe au jour suivant
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : FV1
Créé le ...... : 06/12/2016
Modifié le ... :   /  /
Description .. : Zone de calcul de temps : date, heure, intervalle entre deux
Suite ........ : date, intervalle entre deux heures
Mots clefs ... :
*****************************************************************}
Function CalculDateFinParcMat(NBHJour, Duree : Double; DDeb : TdateTime; TobjFerie : Tob) : String; overload;
var NBJour    : Integer;
    //
    DFin      : TDateTime;
    Dtemp     : TDatetime;
    //
    Jour      : Integer;
    StParam   : string;
begin

  Result := DateToStr(idate2099);

  //Si la durée est égale à zéro alors la date de fin = 01/01/2099
  if Duree = 0 then
  begin
    Result       := DateTimeToStr(idate2099);
    Exit;
  end;

  //Si le durée est inférieure au nombre d'heure travaillée alors date déb = date fin
  if Duree < NBHJour then
  begin
    Result       := DateTimeToStr(DDeb);
    Exit;
  end;

  //On redétermine le nombre de jours entre les deux dates
  NBJour := Round(Duree/NBHJour);

  Dfin := IncDay(DDeb, -1);;

  While NBJour > 0 do
  begin
    //Par défaut on charge la date de fin avec la date de début...
    Dtemp := DFin;
    Dtemp := IncDay(Dtemp);
    if Jourferies(TobjFerie, Dtemp) then
      Dfin := Dtemp
    else
    begin
      //Détermination du Numéro du jour
      Jour := DayOfTheWeek(Dtemp);
      //si on est sur un dimanche on force le Numéro du jour
      if Jour = 7 then
        StParam := 'SO_JOURFERMETURE'
      else
        StParam := 'SO_JOURFERMETURE' + IntToStr(Jour);
      //On vérifie dans les paramètres société si c'est un jour travaillé
      if not GetParamSocSecur(StParam, False) then NBJour := NBJour-1;
      Dfin := Dtemp;
    end;
  end;

  Result := DateTimeToStr(DFin);

end;

Function CalculDateFinParcMat(DateDeb, Datefin : TDateTime; NBHJour : Double; Duree : Integer; TobjFerie : Tob) : TDateTime; overload;
var HeureDebut    : TTime;
    DateDebut     : TDateTime;
    DateFinal     : TDateTime;
    HeureFin      : TTime;
    H, M, S, Msec : Word;
    An, Mois, Jour: Word;
    Cumul         : Integer;
    StParam       : string;
    TempCumul     : Integer;
    HeureDebAM    : TTime;
    HeureFinAM    : TTime;
    HeureDebPM    : TTime;
    HeureFinPM    : TTime;
begin

  //Si la durée est égale à zéro alors la date de fin = 31/12/2099
  if Duree = 0 then
  begin
    Result       := idate2099;
    Exit;
  end;

  DateDebut := DateDeb;
  DecodeDateTime(DateDeb, An, Mois, Jour, H, M, S, Msec);
  HeureDebut:= EncodeTime(H, M, 0, 0);

  //Si le durée est inférieure au nombre d'heure travaillée alors date fin = date Deb
  if Duree <= (NBHJour * 60) then
  begin
    HeureFin     := IncMinute(DateDeb, Duree);
    Result       := HeureFin;
    Exit;
  end;

  //Il faut un chargeParametre temps avec la gestion des calendriers...
  HeureDebAM := GetParamSocSecur('SO_BTAMDEBUT', '08:30');
  HeureFinAM := GetParamSocSecur('SO_BTAMFIN',   '12:30');
  HeureDebPM := GetParamSocSecur('SO_BTPMDEBUT', '14:00');
  HeureFinPM := GetParamSocSecur('SO_BTPMFIN',   '17:30');

  //Par défaut on charge la date de fin avec la date de début...
  DateFinal := DateDeb;

  //Tenir compte des demie-journée......
  while 1=1 Do
  begin
    if Jourferies(TobjFerie, DateFinal) then
    Begin
      DateFinal := IncDay(DateFinal);
      Continue;
    end
    else
    begin
      //Détermination du Numéro du jour
      Jour := DayOfTheWeek(DateFinal);
      //si on est sur un dimanche on force le Numéro du jour
      if Jour = 7 then
        StParam := 'SO_JOURFERMETURE'
      else
        StParam := 'SO_JOURFERMETURE' + IntToStr(Jour);
      //On vérifie dans les paramètres société si c'est un jour travaillé
      if GetParamSocSecur(StParam, False) then
      Begin
        DateFinal := IncDay(DateFinal);
        Continue;
      end
      else
      Begin
        //On isole l'heure de fin de la date complète
        DecodeTime(DateFinal, H, M, S, Msec);
        HeureFin := Encodetime(H, M, 0, 0);
        //il faut vérifier si l'heure de debut de matinée est dépassée : Si oui on charge l'heure de début d'après midi
        if (HeureFin > HeureFinAM) and (HeureFin < HeureDebPM) then
        Begin
          HeureFin  := HeureDebPM;
          DateFinal := Trunc(DateFinal) + Heurefin;
          Continue;
        end;
        //il faut vérifier si l'heure de fin de journée est dépassée : si oui on augmente d'un jour et on positionne l'heure de début de matinée....
        if (HeureFin > HeureFinPM) then
        Begin
          HeureFin  := HeureDebAM;
          DateFinal := IncDay(DateFinal);
          DateFinal := Trunc(DateFinal) + Heurefin;
          Continue;
        end;
        Inc(TempCumul);
        DateFinal := IncMinute(DateFinal, 1);
        if TempCumul > Duree then
        Begin
          DateFinal := IncMinute(DateFinal, 1);
          Break;
        end;
      end;
    end;
  end;

  DecodeDateTime(DateFinal, An, Mois, Jour, H, M, S, Msec);
  HeureFin  := EncodeTime(H, M, 0, 0);

  Result := Trunc(DateFinal) + HeureFin;


end;

Function JourFeries(TobFerie : TOB; DateCtrl : TdateTime) : Boolean;
var Y,M,D       : word;
    DatePaques  : TDateTime;
    TobLJF      : TOB;
begin

  if TobFerie = nil then
  begin
    Result := False;
    Exit;
  end;

  result := true;

  // Cherche parmi les jours fériés (voir TOM_CALENDRIER.LoadJourFerie)
  DecodeDate(DateCtrl, Y, M, D);

  DatePaques := CalculJourPaques(CurrentYear);
  //
  //Dans un premier temps on vérifie si l'on se trouve sur un jour fixe
  TobLJF := TobFerie.FindFirst(['AJF_JOUR', 'AJF_MOIS'], [D, M], False);

  if TOBLJF <> nil then Exit;

  //Si on a pas trouvé dans les jours fixes on vérifie sur les autres...
  //Pâques
  if DatePaques = DateCtrl then Exit;

  //Lundi de ¨pâques
  if incDay(datePaques, 1) = DateCtrl then Exit;

  //Ascension (Pâques + 39)
  if incDay(datePaques, 39) = DateCtrl then Exit;

  //Lundi de Pentecôte
  if incDay(datePaques, 49) = DateCtrl then Exit;

  //Pentecôte
  if incDay(datePaques, 50) = DateCtrl then Exit;

  result := false;

end;

function CalculJourPaques(Annee : Integer) : TDateTime;
var 
  Y,G,C,X,Z,D,E,N,P,J,M : Integer; 
begin 
  {Algorithme valable pour les dates comprises entre 1583 et 4099}
  Y := Annee; 
  G := (Y mod 19) + 1; 
  C := Trunc((Y/100)) + 1; 
  X := Trunc(3 * C / 4) - 12; 
  Z := Trunc(((8 * C) + 5) / 25) - 5; 
  D := Trunc(((5 * Y) / 4) - X - 10); 
  E := ((11 * G)+ 20 + Z - X) mod 30; 
  if ((E = 25) and (G > 11)) or (E = 24) then 
    E := E + 1; 
  N := 44 - E; 
  if N < 21 then
    N := N + 30; 
  P := N + 7 - ((D + N) mod 7); 
  if P > 31 then 
    J := P - 31 
  else 
    J := P; 
  if J = P then 
    M := 3 
  else 
    M := 4; 
  Result := EncodeDate(Annee, M, J); 
end;

end.
