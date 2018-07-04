{***********UNITE*************************************************
Auteur  ...... : Paie
Créé le ...... : 15/02/2006
Modifié le ... : 15/02/2006
Description .. : Vignette des effectifs / absences
Mots clefs ... : Critère Date
*****************************************************************}

unit AFVignetteListePaie;

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  TGAListePAIE = class (TGAVignettePaie)
  private
    EnDateDu, EnDateAu: TDateTime;
    Contrat, Abs, Quoi : string;
  protected
    procedure RecupDonnees; override;
    function GetInterface (NomGrille: string = ''): Boolean; override;
    procedure GetClauseWhere; override;
    procedure DrawGrid (Grille: string); override;
  public
  end;

implementation
uses
  SysUtils,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function TGAListePAIE.GetInterface (NomGrille: string): Boolean;
var LeTyp : Boolean;
begin
  Result := inherited GetInterface ('EnDateDu'); // indique si une date de validité
  EnDateDu := iDate1900;
  EnDateAu := iDate2099;
  Contrat := '';
  Abs := '';
  EnDateDu := GetDateTime ('ENDATEDU'); // Date début du
  EnDateAu := GetDateTime ('ENDATEAU'); // Date fin au
  leTyp := inherited GetInterface ('CONTRAT'); // indique si contrat
  if LeTyp then
  begin
    Contrat := GetString ('CONTRAT'); // Type de contrat
    Quoi := 'CONTRAT';
  end
  else
  begin
    leTyp := inherited GetInterface ('ABS'); // indique si contrat
    if LeTyp then
    begin
      abs := GetString ('ABS'); // motif d'absence
      Quoi := 'ABS';
    end
  end;
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure TGAListePAIE.GetClauseWhere;
var Date1 : TDateTime;
begin
  inherited;
  if Quoi = 'ABS' then
  begin
   Date1 := EnDateDu - 7; // Prise en compte de la semaine dans 1 premier temps
   ClauseWhere := 'where PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI")'+
                   'AND (PCN_DATEDEBUTABS >= "'+USDATETIME (Date1)+'") AND (PCN_DATEFINABS <= "'+
                    USDATETIME (EnDateDu)+'") group BY PCN_TYPECONGE';
  end;
  if Quoi = 'CONTRAT' then
   ClauseWhere := 'WHERE ((PCI_FINCONTRAT >= "'+UsDateTime (EnDateDu)+
                  '") OR (PCI_FINCONTRAT >= "'+UsdateTime(Idate1900)+'"))'+
                  'AND (PCI_DEBUTCONTRAT  <= "'+UsDateTime (EnDateDu)+'") group by PCI_TYPECONTRAT';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure TGAListePAIE.RecupDonnees;
var
  Q: TQuery;
  TobL, MaTob: TOB;
  i_ind: integer;
  St : String;
begin
  inherited;
  if Quoi = 'ABS' then
   st := 'select PCN_TYPECONGE, SUM (PCN_JOURS) from Absencesalarie'
  else if Quoi = 'CONTRAT' then
   St := 'select Count(PCI_TYPECONTRAT), PCI_TYPECONTRAT from contrattravail ';

  MaTob := TOB.Create ('les infos', nil, -1);
  try
    try
      ddWriteLN ('PAIE : Debut de RecupDonnees');
      Q := OpenSelect (st + ClauseWhere);
      if Quoi = 'CONTRAT' Then MaTob.LoadDetailDB ('CONTRATTRAVAIL', '', '', Q, False)
      else MaTob.LoadDetailDB ('Absencesalarie', '', '', Q, False);
      Ferme (Q);
      for i_ind := 0 to MaTob.detail.count - 1 do
      begin
{        TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
        TobL.AddChampSupValeur ('TIERS', TobAffaire.detail [i_ind] .GetString ('AFF_TIERS'));
        TobL.AddChampSupValeur ('LIBELLETIERS', TobAffaire.detail [i_ind] .GetString ('T_LIBELLE'));
        TobL.AddChampSupValeur ('CODEAFFAIRE', TobAffaire.detail [i_ind] .GetString ('AFF_AFFAIRE'));
        TobL.AddChampSupValeur ('LIBELLEAFF', TobAffaire.detail [i_ind] .GetString ('AFF_LIBELLE'));
        TobL.AddChampSupValeur ('DATEDEBUT', TobAffaire.detail [i_ind] .GetString ('AFF_DATEDEBUT'));
        TobL.AddChampSupValeur ('TOTAL', TobAffaire.detail [i_ind] .GetDouble ('AFF_TOTALHT'));
        TobL.AddChampSupValeur ('ETATAFFAIRE', TobAffaire.detail [i_ind] .GetString ('AFF_ETATAFFAIRE'));
        }
      end;
      ddWriteLN ('PAIE : Fin de RecupDonnees');

    except
      on E: Exception do
        MessageErreur := 'Erreur lors du traitement des données avec le message :'#13#13 + E.Message;
    end;
  finally
    MaTob.free;
  end;
end;

{-----Formate les données de la grille -------------------------------------------------}

procedure TGAListePAIE.DrawGrid (Grille: string);
begin
  inherited;
  ddWriteLN ('GAListePAIE : DRAWGRID');
  SetTitreCol (Grille, 'TIERS', 'Client');
  SetTitreCol (Grille, 'LIBELLETIERS', 'Nom Client');
  SetTitreCol (Grille, 'CODEAFFAIRE', 'Code Affaire');
  SetTitreCol (Grille, 'LIBELLEAFF', 'descriptif');
  SetTitreCol (Grille, 'DATEDEBUT', 'Date début');
  SetTitreCol (Grille, 'TOTAL', 'Total HT');
  SetTitreCol (Grille, 'ETATAFFAIRE', 'Etat');

  SetFormatCol (Grille, 'TIERS', 'G.0 ---');
  SetFormatCol (Grille, 'LIBELLETIERS', 'G.0 ---');
  SetFormatCol (Grille, 'CODEAFFAIRE', 'G.0 ---');
  SetFormatCol (Grille, 'LIBELLEAFF', 'G.0 ---');
  SetFormatCol (Grille, 'DATEDEBUT', 'C.0 ---');
  SetFormatCol (Grille, 'TOTAL', 'D.2 ---');
  SetFormatCol (Grille, 'ETATAFFAIRE', 'C.0 ---');

  SetWidthCol (Grille, 'TIERS', 30);
  SetWidthCol (Grille, 'LIBELLETIERS', 40);
  SetWidthCol (Grille, 'CODEAFFAIRE', 40);
  SetWidthCol (Grille, 'LIBELLEAFF', 50);
  SetWidthCol (Grille, 'DATEDEBUT', 30);
  SetWidthCol (Grille, 'TOTAL', 30);
  SetWidthCol (Grille, 'ETATAFFAIRE', 20);

  {  if not Assigned(TobDonnees.Detail[0]) or
       (TobDonnees.Detail[0].GetString('SYSDOSSIER') = '') then
      SetVisibleCol('GRID', 'SYSDOSSIER', False);}
end;

end.
