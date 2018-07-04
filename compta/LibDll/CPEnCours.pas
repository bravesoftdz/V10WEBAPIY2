{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 7.00.001.001  22/11/05   JP   Création de l'unité : La vignette Indicateur d'En-Cours
--------------------------------------------------------------------------------------}
unit CPEnCours;

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
  TObjetEnCours = class(TCPVignettePlugIn)
  private

  protected
    procedure RecupDonnees ;             override;
    procedure GetClauseWhere;            override;
    procedure DrawGrid(Grille : string); override;
  public

  end;

implementation

uses
  SysUtils, HEnt1;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCours.DrawGrid(Grille: string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ddWriteLN('PGICPENCOURS : DRAWGRID');
  SetTitreCol('GRID' , 'EN-COURS', 'En-Cours');
  SetTitreCol('GRID' , 'ÉCHÉANCE', 'Échéance');
  SetFormatCol('GRID', 'EN-COURS', 'D.2 ---');
  SetFormatCol('GRID', 'ÉCHÉANCE', 'C.0 ---');
end;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCours.GetClauseWhere;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  inherited;

  if GetString('E_AUXILIAIRE') <> '' then
    SetClause('E_AUXILIAIRE >= "' + GetString('E_AUXILIAIRE') + '"');
  if GetString('E_AUXILIAIRE_') <> '' then
    SetClause('E_AUXILIAIRE <= "' + GetString('E_AUXILIAIRE_') + '"');
  if GetString('NATUREAUXI') <> '' then
    SetClause('T_NATUREAUXI = "' + GetString('NATUREAUXI') + '" ');
  if GetString('GENERAL') <> '' then
    SetClause('E_GENERAL >= "' + GetString('GENERAL') + '"');
  if GetString('GENERAL_') <> '' then
    SetClause('E_GENERAL <= "' + GetString('GENERAL_') + '"');
  ddWriteLN('PGICPBUDGET : Dossier (' + GetString('GENERAL') + ')');
  if GetString('ETABLISSEMENT') <> '' then
    SetClause('E_ETABLISSEMENT = "' + GetString('ETABLISSEMENT') + '"');
  ddWriteLN('PGICPBUDGET : Dossier (' + GetString('ETABLISSEMENT') + ')');
(*  if TobEntree.GetString('E_DATEECHEANCE') <> '' then
    SetClause('E_DATEECHEANCE >= "' + USDateTime(TobEntree.GetDateTime('E_DATEECHEANCE')) + '"');
  if TobEntree.GetString('E_DATEECHEANCE_') <> '' then
    SetClause('E_DATEECHEANCE <= "' + USDateTime(TobEntree.GetDateTime('E_DATEECHEANCE_')) + '"');
    *)
  if GetString('MVTYPE') <> '' then begin
    if Pos('N', GetString('MVTYPE')) > 0 then s := '(E_QUALIFPIECE = "N"';
    if Pos('S', GetString('MVTYPE')) > 0 then begin
      if s <> '' then s := s + ' OR E_QUALIFPIECE = "S"'
                 else s := '(E_QUALIFPIECE = "S"';
    end;
    if Pos('U', GetString('MVTYPE')) > 0 then begin
      if s <> '' then s := s + ' OR E_QUALIFPIECE = "U"'
                 else s := '(E_QUALIFPIECE = "U"';
    end;
    if s <> '' then s := s + ')';

    SetClause(s);
  end
  else
    SetClause('(E_QUALIFPIECE = "N" OR E_QUALIFPIECE = "S" OR E_QUALIFPIECE = "U")');

  SetClause('E_ECRANOUVEAU = "N" AND (E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL")');
end;

{---------------------------------------------------------------------------------------}
procedure TObjetEnCours.RecupDonnees;
{---------------------------------------------------------------------------------------}
const
  s = 'SELECT SUM(E_DEBIT - E_COUVERTURE) DEB, SUM(E_CREDIT - E_COUVERTURE) AS CRED FROM ECRITURE ' +
      'LEFT JOIN TIERS ON E_AUXILIAIRE = T_AUXILIAIRE WHERE ';

    {----------------------------------------------------------------------}
    function GetEnCours(Lib : string; Deb, Fin : TDateTime) : string;
    {----------------------------------------------------------------------}
    var
      e : string;
      T : TQuery;
      C : Double;
      D : Double;
      F : TOB;
    begin
      e := ' AND E_DATEECHEANCE BETWEEN "' + UsDateTime(Deb) + '" AND "' + UsDateTime(Fin) + '"';
      T := OpenSelect(s + ClauseWhere + e);
      try
        C := T.FindField('CRED').AsFloat;
        D := T.FindField('DEB').AsFloat;
        F := TOB.Create('£ECHE', TobDonnees, -1);
        F.AddChampSupValeur('ÉCHÉANCE', Lib);
        if C < D then F.AddChampSupValeur('EN-COURS', FloatToStr(D - C) + ' D')
                 else F.AddChampSupValeur('EN-COURS', FloatToStr(C - D) + ' C')
      finally
        Ferme(T);
      end;
    end;

begin
  inherited;
  try
    ddWriteLN('PGICPENCOURS : Debut des traitements des En-Cours');
    {<-90 jours}
    GetEnCours('< 90', iDate1900, Now - 91);
    {-60 jours}
    GetEnCours('< 60', Now - 90, Now - 61);
    {-30 jours}
    GetEnCours('< 30', Now - 60, Now - 31);
    {-0 jour}
    GetEnCours('< 0' , Now - 30, Now - 1);
    {Aujourd'hui}
    GetEnCours('= 0' , Now, Now);
    {+0 jours}
    GetEnCours('> 0' , Now + 1, Now + 30);
    {+30 jours}
    GetEnCours('> 30' , Now + 31, Now + 60);
    {+60 jours}
    GetEnCours('> 60' , Now + 61, Now + 90);
    { + 90 jours}
    GetEnCours('> 90' , Now + 91, iDate2099);
    ddWriteLN('PGICPENCOURS : Fin des traitements des En-Cours');
  except
    on E : Exception do
      MessageErreur := 'Erreur lors du traitement des données avec le message :'#13 + E.Message;
  end;
end;

end.


