unit PGVIGTOOLSIDE;
{***********UNITE*************************************************
Auteur  ...... : Paie - MF - PH
Créé le ...... : 27/06/2006
Modifié le ... :
Description .. : Vignette des absences à valider pour toolsidebar VISTA
               : Vignette : PG_VIG_TOOLSIDE
               : Tablette : PGPERIODEVIGNETTE
               : Table    : MOTIFABSENCE
               : Vue      : PGMVTABSRESP
Mots clefs ... :
*****************************************************************}

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
  PG_VIG_TOOLSIDE = class (TGAVignettePaie)
  private
    EnDateDu, EnDateAu  : TDateTime;
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
  SysUtils;

{-----Lit les critères ------------------------------------------------------------------}

function PG_VIG_TOOLSIDE.GetInterface (NomGrille: string): Boolean;
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

end;

{-----Critère de la requète ------------------------------------------------------------}

procedure PG_VIG_TOOLSIDE.GetClauseWhere;
begin
  inherited;

   ClauseWhere := ' PCN_VALIDRESP="ATT" AND';
   ClauseWhere := ClauseWhere+' (PCN_DATEDEBUTABS <="'+UsDateTime(EnDateAu)+'" AND PCN_DATEFINABS >= "'+UsDateTime(EnDateDu)+
   '") AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE is null) ';
   ClauseWhere := ' WHERE '+ ClauseWhere;
end;

{-----Chargement des données -----------------------------------------------------------}

procedure PG_VIG_TOOLSIDE.RecupDonnees;
var
  St : String;
  Q                      : TQuery;
//  DataTob                : TOB;
begin
  inherited;
  try
  st := 'SELECT PSA_LIBELLE, PSA_PRENOM, PMA_MOTIFABSENCE, PMA_LIBELLE,PCN_DATEFINABS,PCN_DATEDEBUTABS FROM PGMVTABSRESP'+
  ' LEFT JOIN MOTIFABSENCE ON PCN_TYPECONGE=PMA_MOTIFABSENCE AND ##PMA_PREDEFINI##';
  st := St + ClauseWhere + ' ORDER BY PCN_DATEFINABS ,PCN_SALARIE' ;
  Q := OpenSelect (St);
  TobDonnees.LoadDetailDB ('ABS_', '', '', Q, False);
  Ferme (Q);
//  TobDonnees.savetofile('c:\MaTOB.txt',false,true,true);
  except
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

function PG_VIG_TOOLSIDE.SetInterface: Boolean;
var
   i                                   : integer;
   Suite, LeHint, MonFormat            : string;
   TT                                  : TOB;
begin
 inherited SetInterface;
  Suite := '';
  for i := 1 to 5 do
  begin
    SetControlValue ('LBL'+intToStr(i),'');
    SetControlValue ('LBL'+intToStr(i)+'1','');
  end;

  for i:=0 to TobDonnees.Detail.Count-1 do
  begin
       if i > 4 then break;
       if (i = 4) and (TobDonnees.Detail.Count > 5) then
          Suite := '             ...';

       SetControlValue ('LBL'+intToStr(i+1),TobDonnees.Detail[i].GetValue('PSA_PRENOM')+' '+TobDonnees.Detail[i].GetValue('PSA_LIBELLE')+Suite);
       TT := TobDonnees.Detail[i];
       if (TT.GetValue('PMA_MOTIFABSENCE')= 'PRI') then LeHint := 'CP'
       else leHint := TT.GetValue('PMA_MOTIFABSENCE');
       If V_PGI.LanguePrinc = 'UK' then MonFormat := 'mmm dd'
         else MonFormat := 'dd mmm';
       LeHint := LeHint + ' '+FormatDateTime(MonFormat, TT.GetValue('PCN_DATEDEBUTABS'))+ ' - '+
          FormatDateTime(MonFormat, TT.GetValue('PCN_DATEFINABS'));
       SetControlValue('LBL'+intToStr(i+1)+'1', LeHint);
  end;
  result:=true;
end;


procedure PG_VIG_TOOLSIDE.DrawGrid(Grille: string);
begin
  inherited;

end;

end.
