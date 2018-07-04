{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 08/10/2002
Modifié le ... :   /  /
Description .. : Méthodes sur les jetons
Mots clefs ... :
*****************************************************************}
unit wJetons;

interface

Type
  tCleWJT = Record
              Prefixe: string;
              Contexte: string;
            end;

{ Where }
function WhereWJT(Const CleWJT:tCleWJT): string;

{ Exist }
function wExistWJT(Const CleWJT: tCleWJT; Const WithAlert: Boolean = false): boolean;

{ Create }
function wCreateWJT(Const CleWJT: tCleWJT; Const Jeton: integer): boolean;
procedure wCreateWJTForAnyWAN(Const Prefixe: string);

{ Modify }
function wUpdateWJT(Const CleWJT: tCleWJT; Const Jeton: integer):boolean;

{$IFNDEF EAGLSERVER}
{ Recalcul }
procedure wResetWJT(Const Question: Boolean=true);
{$ENDIF EAGLSERVER}

{ Get }
function wGetMaxIdentifiant(Const Prefixe: string):integer;
Function wGetJeton(Const Prefixe:string; Const Contexte: string = ''): integer;

{ Set }
function wSetJeton(Const Prefixe:string; Const Contexte: string = ''; Const Plage: integer=1): integer;
procedure SwitchToSequences;

implementation

//GP_20080204_TS_GP14791 >>>
uses
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    Db,
    {$IF not(Defined(EAGLSERVER) or Defined(ERADIO))}
      FE_Main,
    {$IFEND not(EAGLSERVER or ERADIO)}
  {$ELSE}
    MainEagl,
  {$ENDIF !EAGLCLIENT}
	HMsgBox,
  Variants,
	HCtrls,
	Hent1,
	uTob,
  SysUtils,
  Controls,
  M3FP,
	wCommuns
  ;
//GP_20080204_TS_GP14791 <<<

Const
  TableName = 'WJETONS';


procedure SwitchToSequences;
begin
  //
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 09/10/2002
Modifié le ... :   /  /
Description .. : Création d'une clef
Mots clefs ... :
*****************************************************************}
function wCreateWJT(Const CleWJT: tCleWJT; Const Jeton: integer): boolean;
var
  tobWJT: Tob;
begin
  tobWJT := Tob.Create(Tablename, nil, -1);
  try
    tobWJT.SetString('WJT_PREFIXE', CleWJT.Prefixe);
    tobWJT.SetString('WJT_CONTEXTE', CleWJT.Contexte);
    tobWJT.SetInteger('WJT_JETON', Jeton);

    { Insert }
    Result := tobWJT.InsertDB(nil);
  finally
    TobWJT.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 09/10/2002
Modifié le ... :   /  /
Description .. : Construction du where
Mots clefs ... :
*****************************************************************}
function WhereWJT(Const CleWJT:tCleWJT): string;
begin
  Result := '('
          + 'WJT_PREFIXE = "' + CleWJT.Prefixe + '"'
          + ' AND WJT_CONTEXTE = "' + CleWJT.Contexte + '"'
          + ')'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 09/10/2002
Modifié le ... :   /  /
Description .. : Test d'existance d'une clef
Mots clefs ... :
*****************************************************************}
function wExistWJT(Const CleWJT: tCleWJT; Const WithAlert: Boolean = false): boolean;
var
	Sql: string;
begin
	Sql := 'SELECT 1'
			 + ' FROM ' + TableName
			 + ' WHERE ' + WhereWJT(CleWJT)
     	 ;
	Result := ExisteSQL(Sql);

  if WithAlert and (not Result) then
  begin
    PgiError(Format(TraduireMemoire('L''identifiant pour le prefixe %s et la nature de travail %s n''existe pas.'), [CleWJT.Prefixe,CleWJT.Contexte]), 'Identifiant');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 09/10/2002
Modifié le ... :   /  /
Description .. : Création des WJT basé sur une nature pour un prefixe en
Suite ........ : particulier
Mots clefs ... :
*****************************************************************}
procedure wCreateWJTForAnyWAN(Const Prefixe: string);
var
  i     : integer;
  Sql   : string;
  TobWNA: Tob;

  function GetCleWJT: tCleWJT;
  begin
    Result.Prefixe := Prefixe;
    Result.Contexte := TobWNA.Detail[i].GetString('WNA_NATURETRAVAIL');
  end;
begin
  { On vire les lignes en trop }
  Sql := 'DELETE FROM ' + TableName
       + ' WHERE WJT_CONTEXTE IN (SELECT WNA_NATURETRAVAIL FROM WNATURETRAVAIL WHERE WNA_ACTIF="-")'
       ;
  ExecuteSQL(Sql);

  { On crée les nouvelles lignes }
  Sql := 'SELECT WNA_NATURETRAVAIL'
       + ' FROM WNATURETRAVAIL'
       + ' WHERE WNA_ACTIF="' + wTrue + '"'
       + ' ORDER BY WNA_NATURETRAVAIL'
       ;
  TobWNA := Tob.Create('WNA', nil, -1);
  try
    if TobWNA.LoadDetailDBFromSql('WNATURETRAVAIL', Sql) then
    begin
      for i := 0 to TobWNA.Detail.Count - 1 do
        if not wExistWJT(GetCleWJT) then wCreateWJT(GetCleWJT, 1);
    end;
  finally
    TobWNA.free
  end;
end;

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 18/07/2002
Modifié le ... :   /  /
Description .. : Appel de la procédure de re-calcul des jetons de
Suite ........ : production -> pour chaque préfixe distinct contenu dans DECHAMPS.
Mots clefs ... :
*****************************************************************}
Procedure wResetWJT(Const Question: Boolean=true);
Var
  i              : integer;
  TobDH          : Tob;
  Sql, ErrPrefixe: string;

  function GetCleWJT: tCleWJT;
  begin
    Result.Prefixe := TobDH.Detail[i].GetString('DH_PREFIXE');
    Result.Contexte := '';
  end;
begin
{$IFNDEF PGIMAJVER}
  if (not Question) or (PgiAsk('Confirmez-vous le lancement du recalcul des identifiants ', '') = MrYes) then
  begin
{$ENDIF PGIMAJVER}
    Sql := 'SELECT DISTINCT DH_PREFIXE'
         + ' FROM DECHAMPS'
         + ' WHERE (DH_NOMCHAMP LIKE "%_IDENTIFIANT"'
         + ' AND DH_TYPECHAMP = "INTEGER")'
         + ' OR (DH_NOMCHAMP="GTE_CODEPTRF")'
         + ' OR (DH_NOMCHAMP="GPX_NUMPREPA")'
         ;
    TobDH := Tob.Create('DH', nil, -1);
    try
      if TobDH.LoadDetailDBFromSql('DECHAMPS', Sql) then
      begin
        wInitProgressForm(nil, traduireMemoire('Recalcul des identifiants'), '', TobDH.Detail.Count, False, True);
        try
          for i := 0 to TobDH.Detail.Count-1 do
          begin
            wMoveProgressForm(TraduireMemoire('Prefixe : ') + TobDH.Detail[i].GetString('DH_PREFIXE'));
            if not wUpdateWJT(GetCleWJT, wGetMaxIdentifiant(GetCleWJT.Prefixe) + 1) then
            begin
              ErrPrefixe := GetCleWJT.Prefixe;
            end;
          end;
        finally
          wFiniProgressForm;
        end;
      end;
    finally
      TobDH.free;
    end;

{$IFNDEF PGIMAJVER}
    if ErrPrefixe = '' then PgiInfo('Traitement terminé', 'Recalcul des identifiants')
                       else PgiError(Format(TraduireMemoire('Traitement interrompu, problème sur la table %s'), [PrefixeToTable(ErrPrefixe)]), 'Recalcul des identifiants');
  end;
{$ENDIF PGIMAJVER}
end;
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 17/07/2002
Modifié le ... :   /  /
Description .. : Recherche du dernier n° de jeton utilisé une table, à partir du préfixe
Mots clefs ... :
*****************************************************************}
function wGetMaxIdentifiant(Const Prefixe: string): integer;
Var
  Sql: string;
  Q  : tQuery;
begin
  Result := 0;

  if Prefixe = 'GTE' then
    Sql := 'SELECT MAX(CAST(GTE_CODEPTRF AS INTEGER)) IDENTIFIANT'
  else if Prefixe = 'GPX' then
    Sql := 'SELECT MAX(GPX_NUMPREPA) IDENTIFIANT'
  else
    Sql := 'SELECT MAX(' + Prefixe + '_IDENTIFIANT) IDENTIFIANT'
  ;
  Sql := Sql + ' FROM ' + PrefixeToTable(Prefixe);

  Q := OpenSQL(Sql,true, 1);
  try
    if not Q.eof then
    begin
      result := Q.FindField('IDENTIFIANT').AsInteger;
    end;
  finally
    ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 17/07/2002
Modifié le ... : 09/10/2002 (JLS)
Description .. : Mise à jour du n° de jeton avec le prochain n° à uriliser
Suite ........ : après vérification du dernier n° utilisé dans la table
Suite ........ : correspondant au préfixe
Mots clefs ... :
*****************************************************************}
function wUpdateWJT(Const CleWJT: tCleWJT; Const Jeton: integer): boolean;
var
  TobWJT: Tob;
  Sql   : string;
begin
  Result := false;
  if wExistWJT(CleWJT) then
  begin
    Sql := 'SELECT *'
         + ' FROM ' + TableName
         + ' WHERE ' + WhereWJT(CleWJT)
         ;
    TobWJT := Tob.Create('WJT', nil, -1);
    try
      if TobWJT.LoadDetailDBFromSql(TableName, Sql) then
      begin
        TobWJT.Detail[0].SetInteger('WJT_JETON', Jeton);

        { Update }
        Result := TobWJT.UpdateDb;
      end;
    finally
      TobWJT.free;
    end;
  end
  else
  begin
    Result := wCreateWJT(CleWJT, Jeton);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. : Lit le jeton courant
Mots clefs ... : JETON
*****************************************************************}
Function wGetJeton(Const Prefixe:string; Const Contexte: string = ''): integer;
var
  Sql: string;
  Q  : tQuery;

  function GetCleWJT: tCleWJT;
  begin
    Result.Prefixe := Prefixe;
    Result.Contexte := Contexte;
  end;
begin
  Sql := 'SELECT WJT_JETON'
       + ' FROM ' + TableName
       + ' WHERE ' + WhereWJT(GetCleWJT)
       ;
  Q := OpenSql(Sql, True, 1);
  try
    if not Q.eof then
    begin
      if Q.FindField('WJT_JETON').Value=Null then
        Result := 1
      else
        Result := Q.FindField('WJT_JETON').Value
    end
    else
    begin
      wCreateWJT(GetCleWJT, 1);
   	  Result := WGetJeton(Prefixe, Contexte);
    end;
  finally
    Ferme(Q);
  end;
end;

Function AGLwGetJeton(Parms : array of variant; nb : integer): Variant;
begin
	Result := wGetJeton(VarToStr(Parms[0]), VarToStr(Parms[1]));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 26/09/2001
Modifié le ... :   /  /
Description .. : Renvoie le jeton courant et l'incrémente
Mots clefs ... : JETON
*****************************************************************}
//GP_20080204_TS_GP14791 >>>
function wSetJeton(Const Prefixe:string; Const Contexte: string = ''; Const Plage: integer=1): integer;

  function GetCleWJT: tCleWJT;
  begin
    Result.Contexte := Contexte;
    Result.Prefixe := Prefixe;
  end;

var
	Sql: String;
//GP_20080715_TS_GP15294 >>>
const
  firstJeton: Boolean = False;

  function _ExecuteUpdate: Boolean;
  begin
    if not firstJeton then
      Result := wExecuteSqlOnClonedDB(Sql) = 1
    else
      Result := ExecuteSql(Sql) = 1
  end;
//GP_20080715_TS_GP15294 <<<

begin
  Repeat
    Result := wGetJeton(Prefixe, Contexte);
//GP_20080715_TS_GP15294 >>>
    if not firstJeton then
      firstJeton := Result = 1;
//GP_20080715_TS_GP15294 <<<
    Sql := 'UPDATE ' + TableName
         + ' SET WJT_JETON  = ' + IntToStr(Result + Plage)
         + ' WHERE ' + WhereWJT(GetCleWJT)
         + ' AND (WJT_JETON=' + IntToStr(Result) + ' OR (WJT_JETON IS NULL))'
//GP_20080715_TS_GP15294
  until _ExecuteUpdate()
end;
//GP_20080204_TS_GP14791 <<<

Function AGLwSetJeton(Parms : array of variant; nb : integer): Variant;
begin
	Result := wSetJeton(VarToStr(Parms[0]), VarToStr(Parms[1]));
end;

Initialization
	if (FindAglFunc('wSetJeton') = nil) then RegisterAglFunc('wSetJeton', False , 2, AGLwSetJeton);
	if (FindAglFunc('wGetJeton') = nil) then RegisterAglFunc('wGetJeton', False , 2, AGLwGetJeton);
	if (FindAglFunc('xSetJeton') = nil) then RegisterAglFunc('xSetJeton', False , 2, AGLwSetJeton);
	if (FindAglFunc('xGetJeton') = nil) then RegisterAglFunc('xGetJeton', False , 2, AGLwGetJeton);
end.
