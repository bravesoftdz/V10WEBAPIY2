unit PGVIGVALABS;
{***********UNITE*************************************************
Auteur  ...... : Paie -MF
Créé le ...... : 21/06/2006
Modifié le ... :
Description .. : Vignette des absences par motif (Camembert)
               : Vignette : PG_VIG_VALABS
               : Tablette : PGPERIODEVIGNETTE
               : Table    : MOTIFABSENCE
               : Vue      : PGMVTABSRESP
Mots clefs ... :
*****************************************************************
PT1  | 29/01/2008 | FLO | Restriction de la liste aux salariés sous la responsabilité du manager connecté
}

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  uToolsOWC,
  PGVignettePaie,
  PGVIGUTIL,
  HCtrls;

type
  PG_VIG_VALABS = class (TGAVignettePaie)
  private
    EnDateDu, EnDateAu  : TDateTime;
    ValidResp           : string;
    TypeConge           : string;
    procedure ValiderAbs;

  protected
    procedure RecupDonnees; override;
    function GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
    function SetInterface : Boolean ; override;
  public
  end;

implementation
uses
  HEnt1,
  SysUtils,
  PgCalendrier;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_VALABS.GetInterface (NomGrille: string): Boolean;
var
  Periode                           : string;
  sN1                               : string;
//  ValidResp                         : string;

begin
  Result := inherited GetInterface (NomGrille);
  EnDateDu := iDate1900;
  EnDateAu := iDate2099;

  // Qd on lance le portail ParamFich = '', par défaut DateJour = date système
  if (ParamFich = '') then DateJour := DateRef;

  Periode := '003';                        // valeur = 003 (mois)
  // calcul des dates de début et fin de période
  DatesPeriode(Datejour, EnDateDu, EnDateAu, Periode,sN1);

  if (ParamFich = 'moins') OR (ParamFich = 'plus') then
  // on fait glisser les périodes
  begin
    GlisserPeriode (ParamFich,DateJour, Periode);

    // calcul des dates de début et fin de période
    DatesPeriode(Datejour, EnDateDu, EnDateAu, Periode,sN1);
  end;

  SetControlValue ('N1',sN1);

  ValidResp := GetControlValue ('VALIDRESP') ;

  if ValidResp = '' then
    ValidResp := 'ATT';

  SetControlValue ('VALIDRESP',ValidResp);

  TypeConge := GetControlValue ('TYPECONGE') ;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_VALABS.GetClauseWhere;
begin
  inherited;
{  if GetControlValue ('VALIDRESP')='' then
  begin
    SetControlValue ('VALIDRESP','ATT');
    ClauseWhere := ' PCN_VALIDRESP="ATT" ';
  end;
   if (GetControlValue ('VALIDRESP') <> '') then
     ClauseWhere := ClauseWhere+' PCN_VALIDRESP="'+GetControlValue ('VALIDRESP')+'" ';}

   if (ValidResp <> '') then
     ClauseWhere := ClauseWhere+' PCN_VALIDRESP="'+ValidResp+'" ';

 {  if (GetControlValue ('TYPECONGE') <> '') then
   begin
     if (ClauseWhere <> '') then ClauseWhere := ClauseWhere + ' AND ';
     ClauseWhere := ClauseWhere+'PCN_TYPECONGE="'+GetControlValue ('TYPECONGE')+'" ';
   end; }
   if (TypeConge <> '') then
   begin
     if (ClauseWhere <> '') then ClauseWhere := ClauseWhere + ' AND ';
     ClauseWhere := ClauseWhere+'PCN_TYPECONGE="'+TypeConge+'" ';
   end; 

   if (ClauseWhere <> '') then ClauseWhere := ClauseWhere + ' AND ';
      ClauseWhere := ClauseWhere+' (PCN_DATEDEBUTABS <="'+UsDateTime(EnDateAu)+'" AND PCN_DATEFINABS >= "'+UsDateTime(EnDateDu)+
   '") AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE is null) ';
//  ' OR PSA_DATESORTIE>="06/07/2006"))

   ClauseWhere := ClauseWhere + ' AND ' + AdaptByTypeResp('RESPONSABS', V_PGI.UserSalarie, False); //PT1

   ClauseWhere := ' WHERE '+ ClauseWhere;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_VALABS.RecupDonnees;
var
  St,MonFormat : String;
  Q                      : TQuery;
  i : Integer;
  maTob                : TOB;
begin
  inherited;
  if Paramfich = 'VALIDABS' then ValiderAbs ()
  else
  begin
  try
  st := 'SELECT PCN_SALARIE, PSA_LIBELLE, PSA_PRENOM, PMA_LIBELLE, PCN_TYPECONGE,'+
  'PCN_DATEDEBUTABS, "XX" DATEDEBUTABS, PCN_DATEFINABS,"XX" DATEFINABS, PCN_TYPEMVT, PCN_JOURS, '+
  'PCN_VALIDRESP, PCN_ETABLISSEMENT, PCN_EXPORTOK, PCN_ORDRE, PCN_VALIDSALARIE, PSA_DATESORTIE '+
  'PCN_VALIDABSENCE, PCN_DATEVALIDITE  FROM PGMVTABSRESP'+
  ' LEFT JOIN MOTIFABSENCE ON PCN_TYPECONGE=PMA_MOTIFABSENCE AND ##PMA_PREDEFINI##';
//  DataTob := Tob.Create ('LES ABS', nil, -1);
  st := St + ClauseWhere + ' ORDER BY PCN_DATEFINABS ,PCN_SALARIE' ;
  Q := OpenSelect (St);
  TobDonnees.LoadDetailDB ('ABS_', '', '', Q, False);
  ToBDonnees.AddChampSup ('DATEDEBUTABS', TRUE);
  ToBDonnees.AddChampSup ('DATEFINABS', TRUE);
  for i := 0 to TOBDonneeS.Detail.count-1 do
  begin
    MaTob := TobDonnees.detail[i];
//    ToBDonnees.AddChampSup ('DATEDEBUTABS', TRUE);
//    ToBDonnees.AddChampSup ('DATEFINABS', TRUE);
    If V_PGI.LanguePrinc = 'UK' then MonFormat := 'mmm dd'
    else MonFormat := 'dd mmm';
    MaTob.putValue ('DATEDEBUTABS', FormatDateTime(MonFormat, maTob.GetValue('PCN_DATEDEBUTABS')));
    MaTob.putValue ('DATEFINABS', FormatDateTime(MonFormat, maTob.GetValue('PCN_DATEFINABS')));
  end;
  Ferme (Q);

//  TobDonnees.savetofile('c:\MaTOB.txt',false,true,true);
  {
  if (TobDonnees.Detail.Count = 0) then
    MessageErreur := TraduireMemoire('Aucune données pour la période');}
//  ConvertFieldValue(DataTob);

  PutGridDetail('FLISTE', TobDonnees);
  finally
//    FreeAndNil (DataTob);
  end;
//
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure PG_VIG_VALABS.DrawGrid (Grille: string);
begin
  inherited;
    SetVisibleCol('FListe', 'PCN_SALARIE', False);
    SetVisibleCol('FListe', 'PCN_TYPEMVT', False);
    SetVisibleCol('FListe', 'PCN_ETABLISSEMENT', False);
    SetVisibleCol('FListe', 'PCN_EXPORTOK', False);
    SetVisibleCol('FListe', 'PCN_ORDRE', False);
    SetVisibleCol('FListe', 'PCN_VALIDSALARIE', False);
    SetVisibleCol('FListe', 'PCN_DATESORTIE', False);
    SetVisibleCol('FListe', 'PCN_DATEVALIDITE', False);
    SetVisibleCol('FListe', 'PCN_TYPECONGE', FALSE);
    SetVisibleCol('FListe', 'PCN_VALIDABSENCE', False);

    If V_PGI.LanguePrinc = 'UK' then
    begin
    SetTitreCol('FListe' , 'PSA_LIBELLE', 'Last name');
        SetTitreCol('FListe' , 'PSA_PRENOM', 'First name');
        SetTitreCol('FListe' , 'PMA_LIBELLE', 'Absence');
        SetTitreCol('FListe' , 'PCN_DATEDEBUTABS', 'Start');
        SetTitreCol('FListe' , 'PCN_DATEFINABS', 'End');
        SetTitreCol('FListe' , 'DATEDEBUTABS', 'Start');
        SetTitreCol('FListe' , 'DATEFINABS', 'End');
        SetTitreCol('FListe' , 'PCN_JOURS', 'Days');
        SetTitreCol('FListe' , 'PCN_VALIDRESP', 'Status');
    end else
    begin
        SetTitreCol('FListe' , 'PSA_LIBELLE', 'Nom');
        SetTitreCol('FListe' , 'PSA_PRENOM', 'Prénom');
        SetTitreCol('FListe' , 'PMA_LIBELLE', 'Absence');
        SetTitreCol('FListe' , 'PCN_DATEDEBUTABS', 'Début');
        SetTitreCol('FListe' , 'PCN_DATEFINABS', 'Fin');
        SetTitreCol('FListe' , 'DATEDEBUTABS', 'Début');
        SetTitreCol('FListe' , 'DATEFINABS', 'Fin');
        SetTitreCol('FListe' , 'PCN_JOURS', 'Jours');
        SetTitreCol('FListe' , 'PCN_VALIDRESP', 'Etat');
    end;

    SetWidthCol('FListe' , 'PSA_LIBELLE',  26);
    SetWidthCol('FListe' , 'PSA_PRENOM', 24);
    SetWidthCol('FListe' , 'PCN_DATEDEBUTABS', -1);
    SetWidthCol('FListe' , 'PCN_DATEFINABS', -1);
    SetWidthCol('FListe' , 'DATEDEBUTABS', 17);
    SetWidthCol('FListe' , 'DATEFINABS', 17);
    SetWidthCol('FListe' , 'PCN_JOURS', 12);
    SetWidthCol('FListe' , 'PMA_LIBELLE', 19);
    SetWidthCol('FListe' , 'PCN_VALIDRESP', -1);
    SetWidthCol('FListe', 'PCN_SALARIE', -1);
    SetWidthCol('FListe', 'PCN_TYPEMVT', -1);
    SetWidthCol('FListe', 'PCN_ETABLISSEMENT', -1);
    SetWidthCol('FListe', 'PCN_EXPORTOK', -1);
    SetWidthCol('FListe', 'PCN_ORDRE', -1);
    SetWidthCol('FListe', 'PCN_VALIDSALARIE', -1);
    SetWidthCol('FListe', 'PCN_DATESORTIE', -1);
    SetWidthCol('FListe', 'PCN_DATEVALIDITE', -1);
    SetWidthCol('FListe', 'PCN_TYPECONGE', -1);
    SetWidthCol('FListe', 'PCN_VALIDABSENCE', -1);

    SetFormatCol('FListe' , 'DATEDEBUTABS', 'C.0O ---');
    SetFormatCol('FListe' , 'DATEFINABS', 'C.0O ---');
    SetFormatCol('FListe' , 'PCN_JOURS', 'C.0O ---');
    SetFormatCol('FListe' , 'PCN_VALIDRESP', 'C.0O ---');
    SetFormatCol('FListe' , 'PMA_LIBELLE', 'C.0O ---');

end;

function PG_VIG_VALABS.SetInterface: Boolean;
begin
 inherited SetInterface;

  result:=true;
end;

procedure PG_VIG_VALABS.ValiderAbs;
var
  F : TOB;
  n,Ordre : Integer;
  Sal,TypMvt,ValidResp : String;
begin
  {On commence par remplir la TobSelection}
  AvecSelection := True;
  GetInterface('FListe');
  AvecSelection := False;
  {Traitement proprement dit}
  if TobSelection.Detail.Count = 0 then
    MessageErreur := TraduireMemoire('Veuillez sélectionner au moins une ligne.')

  else begin
    for n := 0 to TobSelection.Detail.Count - 1 do begin
      F := TobSelection.Detail[n];
      Sal := F.GetString ('PCN_SALARIE');
      TypMvt := F.GetString ('PCN_TYPEMVT');
      ValidResp := F.GetString ('PCN_VALIDRESP');
      Ordre := StrToInt(F.GetString ('PCN_ORDRE'));
      PgMajAbsEtatValidSal(Sal,'VAL',TypMvt,V_Pgi.UserSalarie,ValidResp,ordre);
      {st := 'UPDATE ABSENCESALARIE SET PCN_VALIDRESP = "'+'ATT'+'",PCN_VALIDABSENCE="'+''+'", '+
      'PCN_DATEMODIF = "'+UsDateTime(Date)+'" '+
      'WHERE PCN_TYPEMVT = "'+F.GetString ('PCN_TYPEMVT')+'" AND PCN_SALARIE ="'+F.GetString ('PCN_SALARIE')+'" '+
      'AND PCN_ORDRE ='+F.GetString ('PCN_ORDRE');
      ExecuteSQL(St); }
    end;
  end;
end;

end.
