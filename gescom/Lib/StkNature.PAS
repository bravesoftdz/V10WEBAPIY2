{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Fonctions & procédure métier
Mots clefs ... : STKNATURE
*****************************************************************}
unit StkNature;

interface

uses
  Windows,
  Classes,
  SysUtils,
  Hctrls,
  Hent1,
  uTob,
  uTom,
  Dialogs,
  Menus,
  Graphics,
  SaisieList,
  uTableFiltre,
  messages,
  Hmsgbox,

  {$IFNDEF EAGLCLIENT}
  DBtables,
  Db,
  FE_Main,
  {$ELSE}
  MainEAGL,
  {$ENDIF}

  wCommuns
  ;

{ Get }
{$IFDEF STK}
function GetTobGSN: boolean;
function GetTobCSK: boolean;
function GetFieldFromGSN(FieldName, QualifMvt: string):Variant;
function GetQualifMvtRef(QualifMvt: string): string;
function GetQualifMvtSuiv(QualifMvt, StkTypeMvt: string): string;
{$ENDIF}

const
	TableName = 'STKNATURE';

implementation

Uses
  EntGP
  ;

{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Génération de la TobGSN
Mots clefs ... :
*****************************************************************}
function GetTobGSN: boolean;
var
  sql : string;
begin
  Result := false;
  if VH_GP.TobGSN.detail.Count = 0 then
  begin
    sql := 'SELECT *'
         + ' FROM ' + TableName
         + ' ORDER BY GSN_QUALIFMVT'
         ;
    Result := wLoadTobFromSql(TableName, sql, VH_GP.TobGSN);
  end;
end;
{$ENDIF}

{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/04/2003
Modifié le ... : 12/06/2003
Description .. : Génération de la TobCSK
Mots clefs ... :
*****************************************************************}
function GetTobCSK: boolean;
var
  sql : string;
begin
  Result := false;
  if Vh_Gp.TobCSK.Detail.Count = 0 then
  begin
    sql := 'SELECT *'
         + ' FROM COMMUN'
         + ' WHERE CO_TYPE="CSK"'
         + ' ORDER BY CO_CODE'
         ;
    Result := wLoadTobFromSql('CSK', sql, VH_GP.TobCSK);
  end;
end;
{$ENDIF}


{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Renvoi un champ de la table GSN
Suite ........ : Celle ci est en mémoire
Mots clefs ... :
*****************************************************************}
function GetFieldFromGSN(FieldName, QualifMvt: string):Variant;
var
  t         : tob;
  TypeField : string;
begin
  if VH_GP.TobGSN.Detail.Count = 0 then GetTobGSN;

  t := VH_GP.TobGSN.FindFirst(['GSN_QUALIFMVT'], [QualifMvt], false);
  if Assigned(t) then
  begin
    Result := t.G(FieldName)
  end
  else
  begin
    TypeField := wGetTypeField(FieldName);
    if (TypeField = 'INTEGER') or (TypeField = 'SMALLINT') then Result := 0
    else if (TypeField = 'DOUBLE') or (TypeField = 'RATE') or (TypeField = 'EXTENDED') then Result := 0.0
    else if (TypeField = 'DATE') then Result := iDate1900
    else Result := ''
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. : Renvoie le mouvement de référence
Mots clefs ... :
*****************************************************************}
function GetQualifMvtRef(QualifMvt: string): string;
var
  i : integer;
begin
  Result := '';

  GetTobGSN;
  for i := 0 to VH_GP.TobGSN.Detail.Count-1 do
  begin
    if Pos(QualifMvt, VH_GP.TobGSN.Detail[i].G('GSN_QUALIFMVTSUIV')) > 0 then
    begin
      Result := VH_GP.TobGSN.Detail[i].G('GSN_QUALIFMVT');
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 11/08/2003
Modifié le ... :   /  /    
Description .. : Renvoie le QualifMvt suivant du type demandé.
Mots clefs ... :
*****************************************************************}
function GetQualifMvtSuiv(QualifMvt, StkTypeMvt: string): string;
var
  QualifMvtSuiv, QualifMvtSilo: string;
begin
  Result := '';
  QualifMvtSuiv := GetFieldFromGSN('GSN_QUALIFMVTSUIV', QualifMvt);
  Repeat
    QualifMvtSilo := ReadTokenSt(QualifMvtSuiv);
    if (QualifMvtSilo <> '') and (GetFieldFromGSN('GSN_STKTYPEMVT', QualifMvtSilo) = StkTypeMvt) then
      Result := QualifMvtSilo;
  until (Result <> '') or (QualifMvtSuiv = '');
end;
{$ENDIF}



end.