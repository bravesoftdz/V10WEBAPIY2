{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  22/11/05   JP   Création de l'unité : La vignette Indicateur d'En-Cours
--------------------------------------------------------------------------------------}
unit CPEnCoursTiers;

interface


uses
  Classes, UTob,
  {$IFNDEF DBXPRESS}
  dbtables,
  {$ELSE}
  uDbxDataSet,
  {$ENDIF}
  uCPVignettePlugIn, HCtrls;

type
  TObjetEnCoursTiers = class(TCPVignettePlugIn)
  private

  protected
    procedure RecupDonnees ;  override;
    procedure GetClauseWhere; override;
    procedure DrawGrid(Grille : string); override;
    procedure GetClauseGC;
  public

  end;

implementation

uses
  SysUtils, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCoursTiers.RecupDonnees;
{---------------------------------------------------------------------------------------}
const
  s = 'SELECT SUM(E_DEBIT - E_COUVERTURE) DEB, SUM(E_CREDIT - E_COUVERTURE) AS CRED, ' +
      'E_AUXILIAIRE, T_LIBELLE FROM ECRITURE ' +
      'LEFT JOIN TIERS ON E_AUXILIAIRE = T_AUXILIAIRE ';
  g = 'GROUP BY E_AUXILIAIRE, T_LIBELLE ORDER BY E_AUXILIAIRE';
  {Jointure sur les journaux et les généraux pour ne traiter que les compte et journaux d'effets}
  e = 'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
      'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ';
  {Requête pour récupérer les commandes non facturées : cf TiersUtil.RisqueTiersGC}
  c = 'SELECT SUM(GL_TOTALTTC * ABS(GL_QTERESTE) / ABS(GL_QTEFACT)) AS CMD, GL_TIERS, T_LIBELLE  FROM LIGNE ' +
      'LEFT JOIN TIERS ON GL_TIERS = T_AUXILIAIRE WHERE ' +
      'GL_VIVANTE = "X" AND GL_ARTICLE <> "" AND GL_ETATSOLDE = "ENC" AND GL_QTEFACT <> 0 AND GL_TOTALTTC <> 0 ';
  r = 'GROUP BY GL_TIERS, T_LIBELLE ORDER BY GL_TIERS';

var
  T : TOB;{Fille des soldes}
  w : string;
  Gescom : TOB;
  Effets : TOB;

    {----------------------------------------------------------------------}
    function CreeFille(Aux, Lib : string) : TOB;
    {----------------------------------------------------------------------}
    begin
      Result := TOB.Create('£REPONSE', TobDonnees, -1);
      Result.AddChampSupValeur('TIERS', Aux);
      Result.AddChampSupValeur('LIBELLE', Lib);
      Result.AddChampSup('SOLDE'   , False);
      Result.AddChampSup('EFFETS', False);
      Result.AddChampSup('COMMANDES', False);
      Result.AddChampSup('TOTAL' , False);
    end;

    {----------------------------------------------------------------------}
    procedure CreeEnreg(Aux : string; Bcl : Byte);
    {----------------------------------------------------------------------}
    var
      A : TOB;{Fille de la TOB réponse à la "requête"}
      F : TOB;{Fille des effets non échus}
      R : TOB;{Fille des commandes non facturés}
    begin
      A := CreeFille(Aux, T.GetString('T_LIBELLE'));
      if Bcl = 1 then begin
        A.SetString('SOLDE', Format('%.2n', [Arrondi(T.GetDouble('DEB') - T.GetDouble('CRED'), 2)]));
        F := Effets.FindFirst(['E_AUXILIAIRE'], [Aux], True);
      end;

      if Bcl <= 2 then begin
        if Bcl = 2 then F := T;
        if Assigned(F) then begin
          A.SetString('EFFETS', Format('%.2n', [Arrondi(F.GetDouble('DEB') - F.GetDouble('CRED'), 2)]));
          {On ôte du solde les effets non échus}
          A.SetString('SOLDE', Format('%.2n', [Arrondi(A.GetDouble('SOLDE') - A.GetDouble('EFFETS'), 2)]));
          {On détruit la tob, pour voir en fin de boucle sur la Tob solde les tiers non traités}
          if Bcl = 1 then FreeAndNil(F);
        end;
        R := Gescom.FindFirst(['GL_TIERS'], [Aux], True);
      end;

      if Bcl = 3 then R := T;
      if Assigned(R) then begin
        A.SetString('COMMANDES', Format('%.2n', [R.GetDouble('CMD')]));
        {On détruit la tob, pour voir en fin de boucle sur la Tob solde les tiers non traités}
        if Bcl <= 2 then FreeAndNil(R);
      end;

      {Calcul du total de l'encours Tiers}
      if Bcl <= 2 then
        A.SetString('TOTAL', Format('%.2n', [Arrondi(Valeur(T.GetString('DEB')) - Valeur(T.GetString('CRED')) + Valeur(A.GetString('COMMANDES')), 2)]))
      else
        A.SetString('TOTAL', Format('%.2n', [Arrondi(T.GetDouble('COMMANDES'), 2)]));
    end;

var
  Solde  : TOB;
  n      : Integer;
  Q      : TQuery;
begin
  inherited;
  ddWriteLN('PGICPENCOURSTIERS : Debut des traitements des En-Cours des tiers');
  try
    {Clause Where sur la date d'échéance, pour ne traiter que les effets non echus}
    w := ' AND E_DATEECHEANCE > "' + UsDateTime(Now) + '" AND G_EFFET = "X" AND J_EFFET = "X" ';

    Solde  := TOB.Create('µSOLDE' , nil, -1);
    Gescom := TOB.Create('µGESCOM', nil, -1);
    Effets := TOB.Create('µEFFETS', nil, -1);
    try

      ddWriteLN('PGICPENCOURSTIERS : ' + s + ClauseWhere + g);
      Q := OpenSelect(s + ClauseWhere + g);
      Solde.LoadDetailDB('µSOLDE', '', '', Q, False);
      Ferme(Q);

      ddWriteLN('PGICPENCOURSTIERS : ' + s + e + ClauseWhere + w + g);
      Q := OpenSelect(s + e + ClauseWhere + w + g);
      Effets.LoadDetailDB('µEFFETS', '', '', Q, False);
      Ferme(Q);

      GetClauseGC;
      ddWriteLN('PGICPENCOURSTIERS : ' + c + ClauseWhere + r);
      Q := OpenSelect(c + ClauseWhere + r);
      Gescom.LoadDetailDB('µGESCOM', '', '', Q, False);
      Ferme(Q);

      for n := 0 to Solde.Detail.Count - 1 do begin
        T := Solde.Detail[n];
        CreeEnreg(T.GetString('E_AUXILIAIRE'), 1);
      end;

      for n := 0 to Effets.Detail.Count - 1 do begin
        T := Effets.Detail[n];
        CreeEnreg(T.GetString('E_AUXILIAIRE'), 2);
      end;

      for n := 0 to Gescom.Detail.Count - 1 do begin
        T := Gescom.Detail[n];
        CreeEnreg(T.GetString('GL_TIERS'), 3);
      end;
    finally
      if Assigned(Solde ) then FreeAndNil(Solde );
      if Assigned(Gescom) then FreeAndNil(Gescom);
      if Assigned(Effets) then FreeAndNil(Effets);
      ddWriteLN('PGICPENCOURSTIERS : Fin des traitements des En-Cours des tiers');
    end;
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCoursTiers.GetClauseWhere;
{---------------------------------------------------------------------------------------}
begin
  inherited;

  if GetString('E_AUXILIAIRE') <> '' then
    SetClause('E_AUXILIAIRE >= "' + GetString('E_AUXILIAIRE') + '"');

  if GetString('E_AUXILIAIRE_') <> '' then
    SetClause('E_AUXILIAIRE <= "' + GetString('E_AUXILIAIRE_') + '"');

  if GetString('ETABLISSEMENT') <> '' then
    SetClause('E_ETABLISSEMENT = "' + GetString('ETABLISSEMENT') + '"');

  if GetString('NATUREAUXI') <> '' then
    SetClause('T_NATUREAUXI = "' + GetString('NATUREAUXI') + '" ');

  SetClause('E_QUALIFPIECE = "N" ');
  SetClause('E_ECRANOUVEAU = "N" AND (E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL")');
  ClauseWhere := ' WHERE ' + ClauseWhere;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCoursTiers.GetClauseGC;
{---------------------------------------------------------------------------------------}
begin
  ClauseWhere := '';

  if GetString('E_AUXILIAIRE') <> '' then
    SetClause('GL_TIERS >= "' + GetString('E_AUXILIAIRE_') + '"');

  if GetString('E_AUXILIAIRE_') <> '' then
    SetClause('GL_TIERS <= "' + GetString('E_AUXILIAIRE_') + '"');

  if GetString('ETABLISSEMENT') <> '' then
    SetClause('GL_ETABLISSEMENT = "' + GetString('ETABLISSEMENT') + '"');

  if GetString('NATUREAUXI') <> '' then
    SetClause('T_NATUREAUXI = "' + GetString('NATUREAUXI') + '" ');

  if ClauseWhere <> '' then
    ClauseWhere := ' AND ' + ClauseWhere;
end;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCoursTiers.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ddWriteLN('PGICPENCOURSTIERS : DrawGrid');

  SetTitreCol('GRID', 'TIERS'    , 'Tiers');
  SetTitreCol('GRID', 'LIBELLE'  , 'libellé');
  SetTitreCol('GRID', 'SOLDE'    , 'Solde cpt.');
  SetTitreCol('GRID', 'EFFETS'   , 'Eff. non échus');
  SetTitreCol('GRID', 'COMMANDES', 'Commande');
  SetTitreCol('GRID', 'TOTAL'    , 'Tot. en-cours');

  SetFormatCol('GRID','SOLDE'    , 'D.2 ---');
  SetFormatCol('GRID','EFFETS'   , 'D.2 ---');
  SetFormatCol('GRID','COMMANDES', 'D.2 ---');
  SetFormatCol('GRID','TOTAL'    , 'D.2 ---');
end;

end.


