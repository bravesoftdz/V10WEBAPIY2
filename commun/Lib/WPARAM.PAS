unit wParam;

interface

uses
  Variants,
	Windows,
	Classes,
	SysUtils,
	HCtrls,
	Hent1,
	uTob,
	Dialogs,
	Forms,
	Controls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
  	MUL,
    Fiche,
  	EdtEtat,
    FE_Main,
  {$ENDIF !ERADIO}
  {$ENDIF EAGLSERVER}
	EdtREtat,
	{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
	Db,
{$ELSE}
	eMul,
	eFiche,
	utilEagl,
  MainEagl,
{$ENDIF}
	HMsgBox,
	UTom,
	Hdb,
	M3FP,
	Menus,
  {$IFDEF GPAO}
    wArtNat,
    EntGP,
  {$ELSE GPAO}
    {$IFDEF SAV}
      {$IFNDEF CRM}
//      wArtNat,
//      EntGP,
    {$ENDIF CRM}
    {$ENDIF SAV}
  {$ENDIF GPAO}
  wCommuns
  ;

Type tWpaField = Record
                   Name   : String;
                   Value  : Variant;
                   Libelle: string;
                 end;

{ Gestion pour un utilisateur}
function WhereWPA(Const CodeParam: string): string;
function wExistWPA(Const CodeParam: string; Const WithAlert: Boolean = false): boolean;
function wIniWPA(Const CodeParam: string): boolean;
function wSetWPA(Const CodeParam: string; Const FieldsValues: Array of Variant): Boolean;
function wGetTobWPA(Const CodeParam: string; Const TobWPA: Tob): Boolean;
function wGetFieldFromWPA(Const CodeParam, FieldName: string): variant;
function wGetFieldsFromWPA(Const CodeParam: String; Const FieldsName: Array of String): MyArrayValue;
function wGetLequelWPA(Const CodeParam: string): string;

{ Demandes }
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
//  BBI Web services
function wAskCodeDate(Var Contexte: string; Var Date:tDateTime; Const Title: string = ''): boolean;
//  BBI Fin Web services
{$IFDEF GPAO}
function wAskArtNat(var CleWAN: tCleWAN; const Title: string = ''): Boolean;
{$ELSE GPAO}
  {$IFDEF SAV}
    {$IFNDEF CRM}
    {$IFNDEF BTP}
    function wAskArtNat(var CleWAN: tCleWAN; const Title: string = ''): Boolean;
    {$ENDIF}
    {$ENDIF CRM}
  {$ENDIF SAV}
{$ENDIF GPAO}

function wAskCodeArticle(Const Param: string = ''; Const Title: string = ''): String;
function wAskDepot(Const Title: String = ''): String;
function wAskBlocNote(Const Title: string = ''): string;
function wAskBlocNoteWithEscape(Var bValide: boolean; Const Title: string = ''): string;
procedure wAskDateBlocNote(Var Date: tDateTime; Var BlocNote: string; Const Title: string = '');
procedure wAskDate(var Date: TDateTime; Const Title: string = '');
procedure wAskString(var St: String; const Title: String = '');
function wAskDataType(var Value: String; const DataType, Plus, Title: String): Boolean;
function wAskTransUser(var Value: String; const DataType, Plus, Title: String): String;
{$IFDEF QUALITE}
function AskClotureRQN(Var Date: tDateTime; Var CodeName,BlocNote: string; Const Title: string = ''):boolean;
{$ENDIF QUALITE}
{$IFDEF SAV}
function wAskVersion(Const Article, VersionActuelle: string; Var BlocNote:String): string;
{$ENDIF SAV}
function wAskCodeOrdre(Const NatureTravail: string; Const Title: string = ''): integer;
procedure wAskParamStkProjete;
//GP_20080507_DKZ_GP14995
function wAskParamGenerationPrevCom: string;
procedure wAskDuplicParamConsoPdR;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

const FieldTitle = 'WPA_LONGVARCHAR01';

implementation

{$IFDEF GPAO}
  uses
    {$IFDEF SAV}
      wVersion,
    {$ENDIF SAV}
	  UtilArticle;
{$ELSE GPAO}
  {$IFDEF SAV}
    uses
      {$IFNDEF CRM}
       	UtilArticle,
      {$ENDIF CRM}
      wVersion;
  {$ENDIF SAV}
{$ENDIF GPAO}

Const
  TableName = 'WPARAM';

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Renvoie un Where WPA
Mots clefs ... :
*****************************************************************}
function WhereWPA(Const CodeParam: string): string;
begin
	Result := 'WPA_CODEPARAM  = "' + CodeParam  + '"'
          + ' AND WPA_UTILISATEUR = "' + V_Pgi.User + '"'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Répond Oui si le COdeParam + Usder existe déjà dans la
Suite ........ : table
Mots clefs ... :
*****************************************************************}
function wExistWPA(Const CodeParam: string; Const WithAlert: Boolean = false): boolean;
var
  Sql	: string;
begin
	Sql := 'SELECT 1'
       + ' FROM ' + TableName
       + ' WHERE ' + WhereWPA(CodeParam)
       ;
  Result := existeSql(Sql);

  if WithAlert and (not Result) then
  begin
    PgiError(Format(TraduireMemoire('Le paramètre utilisateur %s n''existe pas.'), [CodeParam]), 'Paramètre utilisateur');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 27/11/2001
Modifié le ... :   /  /
Description .. : Crée la ligne si elle n'existe pas.
Mots clefs ... : WPARAM
*****************************************************************}
Function wIniWPA(Const CodeParam: string): Boolean;
var
  TobWPA: Tob;
  TomWPA: Tom;
begin
  Result := false;
  if not wExistWPA(CodeParam) then
  begin
   	Result := true;
    TobWPA := Tob.Create(TableName, nil, -1);
    TomWPA := CreateTOM(TableName, nil, false, true);
    try
      { NewRecord }
      TomWPA.InitTOB(TobWPA);

      { Key }
      TobWPA.SetString('WPA_CODEPARAM', CodeParam);

      { UpDateRecord }
      TobWPA.AddChampSupValeur('IKC', 'C', false);
      if TomWPA.VerifTOB(TobWPA) and TobWPA.InsertDB(nil) then
        TomWPA.AfterVerifTOB(TobWPA)
    finally
      TobWPA.free;
      TomWPA.free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 16/11/2001
Modifié le ... : 16/11/2001
Description .. : Crée ou modifie l'enregistrement demandé avec la liste des
Suite ........ : champs et des valeurs données.
Suite ........ : ATTENTION AUX DATES !!! Il faut passer un tDateTime si
Suite ........ : l'on veut que la date soit stockée dans la BD. :o)
Mots clefs ... : WPARAM
*****************************************************************}
function wSetWPA(Const CodeParam: string; Const FieldsValues: Array of Variant): Boolean;
var
	i 		: integer;
  TobWPA: Tob;
  TomWPA: Tom;
begin
	Result := WIniWPA(CodeParam);

  TobWPA := Tob.Create(TableName, nil, -1);
  TomWPA := CreateTOM(TableName, nil, false, true);
  try
    wGetTobWPA(CodeParam, TobWPA);

    { Sauvegarde de la Tob }
    TomWPA.LoadBufferAvantModif(TobWPA);

    i := 0;
    while i < high(FieldsValues) do
    begin
      TobWPA.PutValue(FieldsValues[i], FieldsValues[i+1]);
      i := i + 2;
    end;

    { UpDateRecord }
    TobWPA.AddChampSupValeur('IKC', 'M', false);
    if TomWPA.VerifTOB(TobWPA) and TobWPA.UpdateDB(true) then
      TomWPA.AfterVerifTOB(TobWPA);
  finally
    TobWPA.free;
    TomWPA.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 27/11/2001
Modifié le ... :   /  /
Description .. : Récupère un paramètre user
Mots clefs ... : WPARAM
*****************************************************************}
Function wGetTobWPA(Const CodeParam: string; Const TobWPA: Tob): Boolean;
begin
	Result := WIniWPA(CodeParam);

  TobWPA.SetString('WPA_CODEPARAM', CodeParam);
  TobWPA.SetString('WPA_UTILISATEUR', V_Pgi.User);

  TobWPA.LoadDb;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 03/07/2002
Modifié le ... :   /  /
Description .. : renvoie la valeur d'un champ de WPA
Mots clefs ... :
*****************************************************************}
function wGetFieldFromWPA(Const CodeParam, FieldName: string): variant;
begin
	Result := wGetSqlFieldValue(FieldName, TableName, WhereWPA(CodeParam));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 21/02/2003
Modifié le ... :   /  /
Description .. : Renvoie un tableau de valeurs de champs WPA
Mots clefs ... :
*****************************************************************}
function wGetFieldsFromWPA(Const CodeParam: String; Const FieldsName: Array of String): MyArrayValue;
begin
  Result := wGetSqlFieldsValues(FieldsName, TableName, WhereWPA(CodeParam));
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Demande un article
Mots clefs ... :
*****************************************************************}
function wAskCodeArticle(Const Param: string = ''; Const Title: string = ''): String;
Const
	CodeParam     = 'WPA_CODEARTICLE';
	CodeArticle   = 'WPA_VARCHAR01';
begin
  wSetWPA(CodeParam,[CodeArticle, '', FieldTitle, Title]);
  AglLanceFiche ('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + Param + ';' + iif(Title <> '', 'TITLE=' + Title, ''));
  Result := wGetFieldFromWPA(CodeParam, CodeArticle);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
function wAskCodeDate(Var Contexte: string; Var Date:tDateTime; Const Title: string = ''): boolean;
Var
  WPAValues: MyArrayValue;
	Retour:string;
Const
	  CodeParam      = 'WPA_GETAFCONTEXTE';
    ContexteAfWPA  = 'WPA_VARCHAR01';
    DateAfWPA      = 'WPA_DATE01';
begin
	result:=false;

  wSetWPA(CodeParam,[DateAfWPA, Date, FieldTitle, Title]);
  Retour:= AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION');
  if (Retour<>wFalse) and (Copy(Retour,1,17)<>'WPA_GETAFCONTEXTE') then
  begin
    result:=True;
    WPAValues := wGetFieldsFromWPA(CodeParam, [ContexteAfWPA, DateAfWPA]);
    Date    := WPAValues[1];
    Contexte:= WPAValues[0];
  end;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
function wAskDepot(Const Title: String = ''): String;
Const
	CodeParam     = 'WPA_DEPOT';
	CodeArticle   = 'WPA_COMBO01';
begin
  wSetWPA(CodeParam,[CodeArticle, '', FieldTitle, Title]);
  AglLanceFiche ('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + iif(Title <> '', 'TITLE=' + Title, ''));
  Result := wGetFieldFromWPA(CodeParam, CodeArticle);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 14/11/2002
Modifié le ... : 14/11/2002
Description .. : Demande un article
Mots clefs ... :
*****************************************************************}
{$IFDEF GPAO}
function wAskArtNat(var CleWAN: tCleWAN; Const Title: string = ''): Boolean;
Const
	CodeParam     = 'WPA_ARTNAT';
	CodeArticle   = 'WPA_VARCHAR01';
  NatureTravail = 'WPA_VARCHAR02';
var
  Param: String;
begin
	Param := 'NATURETRAVAIL=' + CleWAN.NatureTravail
  		   + ';CODEARTICLE='  + wGetCodeArticleFromArticle(CleWAN.Article)
         ;
  wSetWPA(CodeParam,[CodeArticle, wGetCodeArticleFromArticle(CleWAN.Article), NatureTravail, CleWAN.NatureTravail, FieldTitle, Title]);
  Result := AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + Param + ';' + iif(Title <> '', 'TITLE=' + FieldTitle, '')) = wTrue;
  CleWAN.NatureTravail := wGetFieldFromWPA(CodeParam, NatureTravail);
  CleWAN.Article       := wGetArticleFromCodeArticle(wGetFieldFromWPA(CodeParam, CodeArticle));
end;
{$ELSE GPAO}
  {$IFDEF SAV}
   {$IFNDEF CRM}
    function wAskArtNat(var CleWAN: tCleWAN; const Title: string = ''): Boolean;
    Const
      CodeParam     = 'WPA_ARTNAT';
      CodeArticle   = 'WPA_VARCHAR01';
      NatureTravail = 'WPA_VARCHAR02';
    var
      Param: String;
    begin
      Param := 'NATURETRAVAIL=' + CleWAN.NatureTravail
             + ';CODEARTICLE='  + wGetCodeArticleFromArticle(CleWAN.Article)
             ;
      wSetWPA(CodeParam,[CodeArticle, wGetCodeArticleFromArticle(CleWAN.Article), NatureTravail, CleWAN.NatureTravail, FieldTitle, Title]);
      Result := AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + Param + ';' + iif(Title <> '', 'TITLE=' + FieldTitle, '')) = wTrue;
      CleWAN.NatureTravail := wGetFieldFromWPA(CodeParam, NatureTravail);
      CleWAN.Article       := wGetArticleFromCodeArticle(wGetFieldFromWPA(CodeParam, CodeArticle));
    end;
   {$ENDIF CRM}
  {$ENDIF SAV}
{$ENDIF GPAO}
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Demande un mémo
Mots clefs ... :
*****************************************************************}
function wAskBlocNote(Const Title: string = ''): string;
Const
	CodeParam = 'WPA_BLOCNOTE';
	BlocNote  = 'WPA_BLOCNOTE';
begin
  wSetWPA(CodeParam,[BlocNote, '', FieldTitle, Title]);
  AglLanceFiche ('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + iif(Title <> '', 'TITLE=' + FieldTitle, ''));
  Result := wGetFieldFromWPA(CodeParam, BlocNote);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Nicolas Chavanne
Créé le ...... : 14/05/2007
Modifié le ... :   /  /
Description .. : Demande un mémo avec retour si fermeture ou validation
Mots clefs ... :
*****************************************************************}
function wAskBlocNoteWithEscape(var bValide: boolean; Const Title: string = ''): string;
Const
	CodeParam = 'WPA_BLOCNOTE';
	BlocNote  = 'WPA_BLOCNOTE';
Var
  sRetour : string;
begin
  { Initialisation }
  bValide := True;

  wSetWPA(CodeParam,[BlocNote, '', FieldTitle, Title]);
  sRetour := AglLanceFiche ('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + iif(Title <> '', 'TITLE=' + FieldTitle, '') + ';WITHCANCELINFO=X');
  { sRetour = 'X' si Bloc Note <>'' et clic sur Valider }
  Result := wGetFieldFromWPA(CodeParam, BlocNote);
  if sRetour=wFalse then bValide := False;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}


{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Demande une date et un mémo
Mots clefs ... :
*****************************************************************}
procedure wAskDateBlocNote(var Date: tDateTime; var BlocNote: string; Const Title: string = '');
Const
	CodeParam      = 'WPA_DATEBLOCNOTE';
	FieldBlocNote  = 'WPA_BLOCNOTE';
  FieldDate      = 'WPA_DATE01';
begin
  wSetWPA(CodeParam,[FieldBlocNote, BlocNote, FieldDate, Date, FieldTitle, Title]);
  AglLanceFiche ('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + iif(Title <> '', 'TITLE=' + FieldTitle, ''));
  Date := wGetFieldFromWPA(CodeParam, FieldDate);
  BlocNote := wGetFieldFromWPA(CodeParam, FieldBlocNote);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFDEF QUALITE}
function AskClotureRQN(Var Date: tDateTime; Var CodeName,BlocNote: string; Const Title: string = ''): boolean;
Var
	Retour:string;
Const
	CodeParam      = 'RQGETCLOTURERQN';
	FieldBlocNote  = 'WPA_BLOCNOTE';
  FieldDate      = 'WPA_DATE01';
  FieldChar			 = 'WPA_VARCHAR01';
begin
	result:=false;
  wSetWPA(CodeParam,[FieldBlocNote, BlocNote, FieldChar, CodeName, FieldDate, Date, FieldTitle, Title]);
  Retour  := AglLanceFiche ('RT', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + iif(Title <> '', 'TITLE=' + FieldTitle, ''));
  if strtobool_(Retour) then     // Fermeture avec validation
  begin
    result:=True;
	  Date 		:= wGetFieldFromWPA(CodeParam, FieldDate);
		CodeName:= wGetFieldFromWPA(CodeParam, FieldChar);
  	BlocNote:= wGetFieldFromWPA(CodeParam, FieldBlocNote);
  end;
end;
{$ENDIF QUALITE}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 28/02/2003
Modifié le ... :   /  /
Description .. : Demande un code version (Module SAV)
Mots clefs ... :
*****************************************************************}
{$IFDEF SAV}
function wAskVersion(Const Article, VersionActuelle: string; Var BlocNote:String): string;
Const
	CodeParam             = 'WPA_CHGVERSION';
	ChampBlocNote         = 'WPA_BLOCNOTE';
  ChampVersionActuelle  = 'WPA_VARCHAR01';
  ChampNewCodeVersion   = 'WPA_VARCHAR02';
Var
  Param       : string;
  lBoucle     : Boolean;
  snewversion : String;
  sTmp : string;

  function GetCleWVS: tCleWVS;
  begin
    Result.Article := Article;
    Result.CodeVersion := sNewVersion;
  end;
begin
  wSetWPA(CodeParam,[ChampVersionActuelle, VersionActuelle, ChampBlocNote, '']);
  Param := 'ARTICLE=' + Article;
  lBoucle := True;
  sNewversion := VersionActuelle;
  while lBoucle do
  begin
    sTmp := AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + Param );
    if sTmp = 'X' then
    begin
      sNewVersion := wGetFieldFromWPA(CodeParam, ChampNewCodeVersion);
      if wVersionPerimee(getCleWVS) then
      begin
        lBoucle := True;
        PgiError('Version périmée. Vous devez saisir une version active','Attention');
      end
      else
        lBoucle := False;
    end
    else
    begin
      lBoucle := False;
      sNewversion := VersionActuelle;
    end;
  end;
  result := sNewversion;
  BlocNote := wGetFieldFromWPA(CodeParam, ChampBlocNote);
end;
{$ENDIF SAV}
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

function wGetLequelWPA(Const CodeParam: string): string;
begin
  Result := CodeParam + ';' + V_Pgi.User
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 06/02/2003
Modifié le ... :   /  /
Description .. : Appel une fiche et renvoi la date modifiée en paramètre par
Suite ........ : adresse
Mots clefs ... :
*****************************************************************}
procedure wAskDate(var Date: TDateTime; Const Title: string = '');
Const
	CodeParam = 'WPA_GETDATE';
  FieldDate = 'WPA_DATE01';
begin
  wSetWPA(CodeParam,[FieldDate, Date, FieldTitle, Title]);
  AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;' + iif(Title <> '', 'TITLE=' + FieldTitle, ''));
  Date := wGetFieldFromWPA(CodeParam, FieldDate);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
function wAskDataType(var Value: String; const DataType, Plus, Title: String): Boolean;
const
	CodeParam = 'WPA_GETSTRING';
  FieldString = 'WPA_VARCHAR01';

  function GetArgument: String;
  begin
    Result := 'MONOFICHE;ACTION=MODIFICATION;FIELDSTRING=' + FieldString
            + iif(Title <> '', ';TITLE=' + FieldTitle, '')
            + iif(DataType <> '', ';DATATYPE=' + DataType, '')
            + iif(Plus <> '', ';PLUS=' + Plus, '')
  end;

begin
  wSetWPA(CodeParam, [FieldString, Value, FieldTitle, Title]);
  AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), GetArgument());
  Value := wGetFieldFromWPA(CodeParam, FieldString);
  Result := Value <> '';
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Nicolas CHAVANNE
Créé le ...... : 31/01/2007
Modifié le ... :   /  /
Description .. : Ouverture de l'écran de sélection de l'utilisateur pour le
Suite ........ : transfert de proposition validées
Mots clefs ... : CBN TRANSFERT UTILISATEUR
*****************************************************************}
function wAskTransUser(var Value: String; const DataType, Plus, Title: String): String;
const
	CodeParam = 'WPA_CBNTRANSUSER';
  FieldString = 'WPA_COMBO01';

  function GetArgument: String;
  begin
    Result := 'MONOFICHE;ACTION=MODIFICATION;FIELDSTRING=' + FieldString
            + iif(Title <> '', ';TITLE=' + FieldTitle, '')
            + iif(DataType <> '', ';DATATYPE=' + DataType, '')
            + iif(Plus <> '', ';PLUS=' + Plus, '')
  end;

begin
  wSetWPA(CodeParam, [FieldString, Value, FieldTitle, Title]);
  Result := AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), GetArgument());
  Value := wGetFieldFromWPA(CodeParam, FieldString);
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}


{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure wAskString(var St: String; const Title: String = '');
begin
  wAskDataType(St, '', '', Title);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Dominique SCLAVOPOULOS
Créé le ...... : 29/07/2003
Modifié le ... :   /  /
Description .. : Sélection d'un n° de regroupement
Mots clefs ... :
*****************************************************************}
function wAskCodeOrdre(Const NatureTravail: string; Const Title: string = ''): integer;
Const
	CodeParam = 'WPA_GETCODEORDRE';
  FieldCodeOrdre = 'WPA_INTEGER01';
begin
  wSetWPA(CodeParam,[FieldCodeOrdre,0]);
  AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;NATURETRAVAIL= ' + NatureTravail );
  Result := wGetFieldFromWPA(CodeParam, FieldCodeOrdre);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Denis KOZA
Créé le ...... : 09/03/2004
Modifié le ... :   /  /
Description .. : > Paramètres stock projeté consolidé
Mots clefs ... :
*****************************************************************}
procedure wAskParamStkProjete;
Const
	CodeParam = 'WPA_GETSTKPROJETE';
begin
  { La 1ère fois initialisation de certains paramètres par défaut }
  wIniWPA(CodeParam);
  AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;');
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : KOZA Denis
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : > Paramètres génération automatique lignes de prévisions
Mots clefs ... :
*****************************************************************}
//GP_20080507_DKZ_GP14995
function wAskParamGenerationPrevCom: string;
Const
	CodeParam = 'WPA_GETPREVCOM';
begin
  { La 1ère fois initialisation de certains paramètres par défaut }
  wIniWPA(CodeParam);
//GP_20080507_DKZ_GP14995
  Result := AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION;');
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : KOZA Denis
Créé le ...... : 16/06/2004
Modifié le ... :   /  /
Description .. : > Paramètres génération automatique lignes de prévisions
Mots clefs ... :
*****************************************************************}
procedure wAskDuplicParamConsoPdR;
Const
	CodeParam = 'WPA_TYPEPDR';
begin
  { La 1ère fois initialisation de certains paramètres par défaut }
  wIniWPA(CodeParam);
  AglLanceFiche('W', CodeParam, '', wGetLequelWPA(CodeParam), 'MONOFICHE;ACTION=MODIFICATION');
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

initialization
end.


