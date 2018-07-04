{***********UNITE*************************************************
Auteur  ...... : Paie
Créé le ...... : 15/02/2006
Modifié le ... : 15/02/2006
Description .. : Vignette des effectifs
Mots clefs ... : Critère Date
*****************************************************************}

unit AFVignetteEffectifPaie;

interface

uses
  Classes,
  UTob,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  PGVignettePaie,
  HCtrls;

type
  TGEffectifPAIE = class (TGAVignettePaie)
  private
    EnDateDu, EnDateAu: TDateTime;
    Contrat : string;
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
  SysUtils,
  HEnt1;

{-----Lit les critères ------------------------------------------------------------------}

function TGEffectifPAIE.GetInterface (NomGrille: string): Boolean;
begin
  Result := inherited GetInterface ('EnDateDu'); // indique si une date de validité
  EnDateDu := iDate1900;
  EnDateAu := iDate2099;
  Contrat := '';
  EnDateDu := GetDateTime ('ENDATEDU'); // Date début du
  Contrat := GetString ('CONTRAT'); // Type de contrat
end;

{-----Critère de la requète ------------------------------------------------------------}

procedure TGEffectifPAIE.GetClauseWhere;
begin
  inherited;
  ClauseWhere := 'WHERE ((PCI_FINCONTRAT >= "'+UsDateTime (EnDateDu)+
                  '") OR (PCI_FINCONTRAT >= "'+UsdateTime(Idate1900)+'"))'+
                  'AND (PCI_DEBUTCONTRAT  <= "'+UsDateTime (EnDateDu)+'") group by PCI_TYPECONTRAT';
end;

{-----Chargement des données -----------------------------------------------------------}

procedure TGEffectifPAIE.RecupDonnees;
var
  Q: TQuery;
  TobL, MaTob: TOB;
  i_ind: integer;
  St : String;
begin
  inherited;
   ddWriteLN ('PAIE : ' + ClauseWhere);
  St := 'select Count(PCI_TYPECONTRAT) NBREC, PCI_TYPECONTRAT from contrattravail ';
//  st := St + ClauseWhere;
  MaTob := TOB.Create ('les infos', nil, -1);
  try
    try
      ddWriteLN ('PAIE : Debut de RecupDonnees');
      ddWriteLN ('PAIE : ' + ClauseWhere);
      Q := OpenSelect (st + ClauseWhere);
      ddWriteLN ('PAIE Le SQl :'+st);
      MaTob.LoadDetailDB ('CONTRATTRAVAIL', '', '', Q, False);
      Ferme (Q);
      for i_ind := 0 to MaTob.detail.count - 1 do
      begin
        TobL := TOB.Create ('£REPONSE', TobDonnees, -1);
        TobL.AddChampSupValeur ('NBRE', MaTob.detail [i_ind] .GetDouble ('NBREC'));
        TobL.AddChampSupValeur ('CONTRAT', MaTob.detail [i_ind] .GetString ('PCI_TYPECONTRAT'));
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

procedure TGEffectifPAIE.DrawGrid (Grille: string);
begin
  inherited;
  ddWriteLN ('GEffectifPAIE : DRAWGRID');
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

function TGEffectifPAIE.SetInterface: Boolean;
var St : String;
    NCDD, NCDI, NAutre : Double;
    i_Num : Integer;
    UneTob : TOB;
begin
    NCDD := 0;
    NCDI := 0;
    NAUTRE := 0;
    ddWriteLN ('PAIE Deb Setinterface ');
//    TObDonnees.SaveToFile ('C:\temp\TOBDONNEE.TXT',FALSE,TRUE, TRUE);
    For i_Num := 0 to TobDonnees.detail.count-1  do
    begin
      UneTob := TobDonnees.detail [I_Num];
      if UneTob.GetValue ('CONTRAT')='CCD' then NCDD := UneTob.getValue ('NBRE')
      else if UneTob.GetValue ('CONTRAT')='CDI' then NCDI := UneTob.getValue ('NBRE')
        else NAutre := NAutre + UneTob.getValue ('NBRE');
     end;
   st := FloatToStr (NCDI);
   SetControlValue ('NBCDI', St);
   st := FloatToStr (NCDD);
   SetControlValue ('NBCDD', St);
   st := FloatToStr (NAutre);
   SetControlValue ('NBAUTRE', St);
   st := FloatToStr (NCDI+NCDD+NAutre);
   SetControlValue ('NBTOTAL', St);
   result := True;
end;

end.
