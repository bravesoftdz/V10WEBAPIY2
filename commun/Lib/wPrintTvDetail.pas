{***********UNITE*************************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 14/01/2004
Modifié le ... :   /  /
Description .. : Renseigne une table temporaire liée à l'utilisateur
Description .. : afin de gérer un treeview dans les éditions
Mots clefs ... :
*****************************************************************}
unit wPrintTvDetail;

interface

uses
  uTob
  ;

function wSetPrintTvDetail(TobTV: Tob; const Prefixe, ChampNiveau: String; const Delete: Boolean = True; const OffSetNiveau: Integer = 0; const Cle1ToDelete: String = ''): Boolean;
function WhereWTD(const Prefixe: String; const Cle1ToDelete: String = ''): String;
function wGetCle1FromTobTV(T: Tob; const Prefixe: String): String;

const
  TableName = 'WPRINTTVDETAIL';

implementation

uses
  HCtrls
  ,HEnt1
  ,SysUtils
  ,wCommuns
  ;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 14/01/2004
Modifié le ... :   /  /    
Description .. : Fabrique un where sur la table WPRINTTVDETAIL
Mots clefs ... : 
*****************************************************************}
function WhereWTD(const Prefixe: String; const Cle1ToDelete: String = ''): String;
begin
  Result := 'WTD_CREATEUR="' + V_PGI.User + '" AND WTD_PREFIXE="' + Prefixe + '"' + iif(Cle1ToDelete <> '', ' AND WTD_CLE1 LIKE "' + Cle1ToDelete + '%"', '')
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 14/01/2004
Modifié le ... :   /  /    
Description .. : Compose un champ à partir de la clé1 des enregistrement 
Suite ........ : de la TobTV
Mots clefs ... :
*****************************************************************}
function wGetCle1FromTobTV(T: Tob; const Prefixe: String): String;
var
  StFieldsCle1: String;
begin
  Result := '';
  StFieldsCle1 := wGetFieldsClef1(PrefixeToTable(Prefixe));
  StFieldsCle1 := StringReplace(StFieldsCle1, ',', ';', [rfIgnoreCase, rfReplaceAll]);
  while StFieldsCle1 <> '' do
    Result := Result + iif(Result <> '', ';', '') + T.GetString(ReadTokenSt(StFieldsCle1))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 14/01/2004
Modifié le ... :   /  /
Description .. : Renseigne la table WPRINTTVDETAIL afin de gérer des
Suite ........ : treeview dans les éditions
Mots clefs ... :
*****************************************************************}
function wSetPrintTvDetail(TobTV: Tob; const Prefixe, ChampNiveau: String; const Delete: Boolean = True; const OffSetNiveau: Integer = 0; const Cle1ToDelete: String = ''): Boolean;
var
  i, j            : Integer;
  TobWTD, T, TNext: Tob;

  function GetNiv(Niv: integer): string;
  begin
    Result := '0' + IntToStr(Niv);
    Result := wRight(Result, 2);
  end;

  function GetNiveauTob(T: Tob; const Champ: String; const OffSet: Integer): Integer;
  begin
    Result := T.GetInteger(Champ) + OffSet
  end;

begin
  Result := False;
  if Delete then
    ExecuteSQL('DELETE FROM ' + TableName + ' WHERE ' + WhereWTD(Prefixe, Cle1ToDelete));
  TobWTD := Tob.Create(PrefixeToTable(TableName), nil, -1);
  try
    { Passe 1: de Haut en bas : préinitialise tous les niveaux de toutes les lignes de la tob & crée la tob basée sur WTD }
    for i := 0 to Pred(TobTV.Detail.Count) do
    begin
      Result := True;
      T := Tob.Create(TableName, TobWTD, -1);
      T.SetString('WTD_PREFIXE', Prefixe);
      T.SetInteger('WTD_ORDRE', i);
      T.SetString('WTD_CLE1', wGetCle1FromTobTV(TobTV.Detail[i], Prefixe));
      if i < Pred(TobTV.Detail.Count) then
      begin
        for j := 0 to Pred(GetNiveauTob(TobTV.Detail[i], Prefixe + '_' + ChampNiveau, OffSetNiveau)) do
        begin
          if j < Pred(GetNiveauTob(TobTV.Detail[i], Prefixe + '_' + ChampNiveau, OffSetNiveau)) then
            T.SetString('WTD_CADRE' + GetNiv(j), 'VER')
          else
            T.SetString('WTD_CADRE' + GetNiv(j), 'TEG')
        end
      end
      else if GetNiveauTob(TobTV.Detail[i], Prefixe + '_' + ChampNiveau, OffSetNiveau) > 0 then
        T.SetString('WTD_CADRE' + GetNiv(GetNiveauTob(TobTV.Detail[i], Prefixe + '_' + ChampNiveau, OffSetNiveau) - 1), 'CBG')
    end;

    { Passe 2: de Bas en haut de la tob basée sur WTD }
    for i := Pred(TobTV.Detail.Count) downto 1 do
    begin
      Result := True;
      T := TobWTD.Detail[Pred(i)];
      TNext := TobWTD.Detail[i];
      for j := 0 to Pred(GetNiveauTob(TobTV.Detail[Pred(i)], Prefixe + '_' + ChampNiveau, OffSetNiveau)) do
      begin
        if (T.GetString('WTD_CADRE' + GetNiv(j)) = 'VER') and (TNext.GetString('WTD_CADRE' + GetNiv(j)) = '') then
          T.SetString('WTD_CADRE' + GetNiv(j), '')
        else if (T.GetString('WTD_CADRE' + GetNiv(j)) = 'TEG') and (TNext.GetString('WTD_CADRE' + GetNiv(j)) = '') then
          T.SetString('WTD_CADRE' + GetNiv(j), 'CBG')
      end
    end;

    { Insert dans la table WPRINTTVDETAIL }
    if Result then
      TobWTD.InsertDB(nil, True)
  finally
    TobWTD.Free;
  end;
end;

end.
