{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 15/01/2007
Modifié le ... :   /  /
Description .. : Fonctions & procédure métier
Mots clefs ... : YPLANNING
*****************************************************************}
unit YPlanning;

interface

uses
  Variants,
  SysUtils,
  Hctrls,
  Hent1,
  uTob,
  uTom,
  Hmsgbox,

{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    FE_Main,
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
{$ELSE}
  MainEAGL,
{$ENDIF}
{$IFNDEF EAGLSERVER}
  wParam,
{$ENDIF EAGLSERVER}
  wCommuns
  ;

  { Méthodes }
  function GestionYPlanning : boolean;
  function WriteYPLanning(pStTypePla : String) : Boolean;
  function ReadYPLanning(pStTypePla : String) : String;
  function GetLibYPlanning(pStTypePla : String; var pStLibelleCourt, pStLibelleLong : String) : Boolean;
  function GetCleOrigine(pStTypePla : String; pStGUID : String) : String;
{$IFNDEF EAGLSERVER}
  procedure LanceFicheOrigine(pStTypePla : String; pStCle : String);
{$ENDIF !EAGLSERVER}
  function CreateYPL(TobYPL: Tob): integer;                                               { Création d'enregistrements dans YPLANNING }
  function UpdateYPL(TobYPL: Tob): integer;                                               { Modification d'enregistrements dans YPLANNING }
  function DeleteYPL(Const Prefixe, GuidOri: string): integer;                            { Suppression d'enregistrements dans YPLANNING }
  function GetLastErrorMsgYPL(Const LastError: integer): String;                          { Renvoie le texte de l'erreur }
  function GetFieldFromYPL(Const FieldName, Prefixe, GuidOri: string): variant;           { Renvoie la valeur du champ demandé }
  function LoadTobYPL(Const Where: String; TobYPL:Tob): Boolean;                          { Chargement des données dans une TobYPL }
  function IsFreeYPL(pStTypePla : String; Const Datedebut, DateFin: tDateTime; Const Where: string): Boolean;  { La période est-elle libre? }

  function RecalculYPL(Const Prefixe: string; Const Sql: string=''; TobYPL: Tob=nil;Silence:Boolean=False) : Boolean; { recalcul de YPLANNING }
                                                                                                           
  function createTobYPL(pTob : Tob; pStPrefixe : String) : Integer;
  function updateTobYPL(pTob : Tob; pStPrefixe : String) : Integer;
  function CreateOrUpdateTobYPL(pTob : Tob; pStPrefixe : String; pBoCreate : Boolean) : Integer;
  function InitTobYPL(pTobYPL, pTob : Tob; pStPrefixe : String) : Integer;           

  const
    cStTableGRC     = 'RAI';
    cStTablePaie    = 'PCN';
    cStTableGa      = 'APL';
    cStTableBureau  = 'JEV';

  TexteMessage: array[1..10] of string = (
    {1} 'ACT',
    {2} 'ACTION',
    {3} 'INT',
    {4} 'INTERVENTION',
    {5} 'AG',
    {6} 'AGENDA',
    {7} 'ABS',
    {8} 'ABSENCE',
    {9} 'Traitement de la table yplanning en cours, veuillez patienter...', //20080828_C.B_GA15215
    {10}'Traitement en cours ...'
    );         

implementation

uses                             
  yPlanning_TOM, paramsoc, yressource,
  {$IFDEF AFFAIRE}
  entgc, // C.B 15/11/2007 mis temporairement tant que la paie de sic ne passe pas par yplanning
  {$ENDIF}
  ed_tools;

function WhereYPL(Const Guid, Prefixe, GuidOri: string): String; overload;
begin
  Result := '';

  if Guid <> '' then
    Result := 'YPL_GUID="' + Guid + '"';

  if Prefixe <> '' then
    Result := Result + iif(Result <> '', ' AND ', '') + 'YPL_PREFIXE="' + Prefixe + '"';

  if GuidOri <> '' then
    Result := Result + iif(Result <> '', ' AND ', '') + 'YPL_GUIDORI="' + GuidOri + '"';
end;

function WhereYPL(TobYPL: Tob): String; overload;
var
  Guid, Prefixe, GuidOri: string;
begin
  if TobYPL.FieldExists('YPL_GUID') then
    Guid := TobYPL.GetString('YPL_GUID')
  else
    Guid := '';

  if TobYPL.FieldExists('YPL_PREFIXE') then
    Prefixe := TobYPL.GetString('YPL_PREFIXE')
  else
    Prefixe := '';

  if TobYPL.FieldExists('YPL_GUIDORI') then
    GuidOri := TobYPL.GetString('YPL_GUIDORI')
  else
    GuidOri := '';

  Result := WhereYPL(Guid, Prefixe, GuidOri);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Dis si la tob est valide:
Suite ........ : Tob virtuelle
Suite ........ : Champs YPF_PREFIXE, YPF_GUIDORi existent
Mots clefs ... :
*****************************************************************}
function ValideYPL(TobYPL: Tob): integer;
begin
  Result := 0;

  if TobYPL.NomTable = 'YPLANNING' then Result := 1
  else if not TobYPL.FieldExists('YPL_PREFIXE') then Result := 2
  else if not TobYPL.FieldExists('YPL_GUIDORI') then Result := 3
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 26/04/2007
Modifié le ... :
Description .. : Cette fonction est utilisée pour savoir si il y a une gestion
Description .. : du planning intégré, c'est à dire si un des paramsoc métiers
Description .. : est <> de <<tous>> qui signifie <<aucun>>
Suite ........ :
Mots clefs ... :
*****************************************************************}
function GestionYPlanning : Boolean;

  function TestParamSoc(pStParamsoc : String) : Boolean;
  var
    vStParam : String;
  begin
    vStParam := GetParamSocSecur(pStParamsoc, '');
    result := (vStParam <> '');
  end;
 
begin
  Result := False;
  //C.B 15/11/2007 optimisation SIC
  {$IFDEF AFFAIRE}
  if TestParamSoc('SO_YPLGA') and (not VH_GC.GCIfDefCEGID) then
  {$ELSE}
  if TestParamSoc('SO_YPLGA') then
  {$ENDIF}
    result := True
  else if TestParamSoc('SO_YPLGRC') then
    result := True
  else if TestParamSoc('SO_YPLPAIE') then
    result := True
  else if TestParamSoc('SO_YPLBUREAU') then
    result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 26/04/2007
Modifié le ... :
Description .. : teste si un des produits lit le planning du produit passé en paramètre,
Description .. : si c'est le cas, il faut ecrire dans la table YPlanning
Suite ........ :
Mots clefs ... :
*****************************************************************}
function WriteYPLanning(pStTypePla : String) : Boolean;

  function EstDansParamsoc(pStTypePla, pStParamsoc : String) : Boolean;
  var
    vStParam : String;
  begin
    result := False;
    vStParam := GetParamSocSecur(pStParamsoc, '');
    if Pos(pStTypePla, vStParam) > 0 then
      result := True;
  end;

begin
  result := False;
  if EstDansParamsoc(pStTypePla, 'SO_YPLGA') then
    result := True

  else if EstDansParamsoc(pStTypePla, 'SO_YPLGRC') then
    result := True

  else if EstDansParamsoc(pStTypePla, 'SO_YPLPAIE') then
    result := True

  else if EstDansParamsoc(pStTypePla, 'SO_YPLBUREAU') then
    result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 27/04/2007
Modifié le ... :
Description .. : retourne la liste des produits a afficher dans le produit
Description .. : passé en paramètre
Description .. : '' si aucun
Mots clefs ... :
*****************************************************************}
function ReadYPLanning(pStTypePla : String) : String;
begin
  result := '';
  if pStTypePla = cStTableGa then
    result := GetParamSocSecur('SO_YPLGA', '')

  else if pStTypePla = cStTableGRC then
    result := GetParamSocSecur('SO_YPLGRC', '')

  else if pStTypePla = cStTablePaie then
    result := GetParamSocSecur('SO_YPLPAIE', '')

  else if pStTypePla = cStTableBureau then
    result := GetParamSocSecur('SO_YPLBUREAU', '');
                                       
  if (result = TraduireMemoire('<<Tous>>')) then
    result := '';
end;
                             
{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 27/04/2007
Modifié le ... :
Description .. : retourne un libellé court et un libellé long liste des produits
Description .. : a afficher dans le produit passé en paramètre
Description .. : '' si aucun
Mots clefs ... :
*****************************************************************}
function GetLibYPlanning(pStTypePla : String; var pStLibelleCourt, pStLibelleLong : String) : Boolean;
begin
  result := False;
  pStLibelleCourt := '';
  pStLibelleLong := '';

  if pStTypePla = cStTableGRC then
  begin
    pStLibelleCourt := traduireMemoire(texteMessage[1]); // ACT
    pStLibelleLong := traduireMemoire(texteMessage[2]);  // ACTION
    result := True;
  end
  else if pStTypePla = cStTableGa then
  begin
    pStLibelleCourt := traduireMemoire(texteMessage[3]); // INT
    pStLibelleLong := traduireMemoire(texteMessage[4]);  // INTERVENTION
    result := True;
  end
  else if pStTypePla = cStTableBureau then
  begin
    pStLibelleCourt := traduireMemoire(texteMessage[5]); // AG
    pStLibelleLong := traduireMemoire(texteMessage[6]);  // AGENDA
    result := True;
  end
  else if pStTypePla = cStTablePaie then
  begin
    pStLibelleCourt := traduireMemoire(texteMessage[7]); // ABS
    pStLibelleLong := traduireMemoire(texteMessage[8]);  // ABSENCE
    result := True;
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 09/08/2007
Modifié le ... : Permet de récuperer la clé de la table d'origine à
Description .. : partir de YPL_GUID
Description .. : fait pour la table 'RAI/RAC' dans un premier temps
Mots clefs ... :
*****************************************************************}
function GetCleOrigine(pStTypePla : String; pStGUID : String) : String;
var
  vSt : String;
  vTob : Tob;

begin
  if pStTypePla = cStTableGRC then
  begin
    vSt := 'SELECT RAC_AUXILIAIRE, RAC_NUMACTION from ACTIONS, ACTIONINTERVENANT, YPLANNING ';
    vSt := vSt + ' WHERE RAC_AUXILIAIRE = RAI_AUXILIAIRE ';
    vSt := vSt + ' AND RAC_NUMACTION = RAI_NUMACTION ';
    vSt := vSt + ' AND YPL_GUIDORI = RAI_GUID ';
    vSt := vSt + ' AND YPL_GUID = "' + pStGUID + '"';

    vTob := TOB.Create('#ACTIONINTERVENANT', nil, -1);
    try
      vTob.LoadDetailFromSQL(vSt);
      if vTob.detail.count = 1 then
        result := vTob.detail[0].GetString('RAC_AUXILIAIRE') +';'+ vTob.detail[0].GetString('RAC_NUMACTION')
      else
        result := '';
    finally
      FreeAndNil(vTob);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 09/08/2007
Modifié le ... : Permet de lancer la fiche d'origine
Description .. : fait pour la GRC dans un premier temps
Mots clefs ... :                  
*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure LanceFicheOrigine(pStTypePla : String; pStCle : String);
begin
  if pStTypePla = cStTableGRC then
    AGLLanceFiche('RT','RTACTIONS','', pStCle, 'ACTION=CONSULTATION');
end;
{$ENDIF !EAGLSERVER}
 
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Création d'un enregistrement dans YPL:
Description .. : -TobYPL contient une tob virtuelle
Description .. : -Les champs YPL_PREFIXE,YPL_GUIDORI sont obligatoires
Description .. :
Description .. :
Mots clefs ... :
*****************************************************************}
function CreateYPL(TobYPL: Tob): integer;
var
  iYPL: integer;

  function MyCreate(TobData: Tob): integer;
  var
    iChamp  : integer;
    MyTobYPL: Tob;
    TomYPL  : Tom;
  begin
    Result := 0;

    TomYPL := CreateTOM('YPLANNING', nil, false, true);
    MyTobYPL := Tob.Create('YPLANNING', nil, -1);
    try
      { OnNewRecord }
      TomYPL.InitTOB(MyTobYPL);

      { Saisie }
      for iChamp := 1000 to (1000 + TobData.ChampsSup.Count - 1) do
        MyTobYPL.PutValue(TobData.GetNomChamp(iChamp), TobData.GetValue(TobData.GetNomChamp(iChamp)));

      { OnUpdateRecord }
      MyTobYPL.AddChampSupValeur('IKC', 'C', false);
      if TomYPL.VerifTOB(MyTobYPL) then
      begin
        if MyTobYPL.InsertDb(nil) then
        begin
          { OnAfterUpdaterecord }
          TomYPL.AfterVerifTOB(MyTobYPL);

          { Retour des informations }
          for iChamp := 1 to MyTobYPL.NbChamps do
            TobData.AddChampSupValeur(MyTobYPL.GetNomChamp(iChamp), MyTobYPL.GetValue(MyTobYPL.GetNomChamp(iChamp)));
        end;
      end
      else
        Result := TomYPL.LastError;
    finally
      MyTobYPL.free;
      TomYPL.free;
    end;
  end;
begin
  Result := 0;
  if not GestionYPlanning then exit;

  Result := ValideYPL(TobYPL);
  if Result = 0 then
    Result := MyCreate(TobYPL)
  else if TobYPL.Detail.Count > 0 then
  begin
    for iYPL := 0 to TobYPL.Detail.Count-1 do
    begin
      MoveCurProgressForm();
      Result := CreateYPL(TobYPL.Detail[iYPL]);
      if Result <> 0 then
        Break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : mise à jour d'un enregistrement dans YPL
Description .. :
Mots clefs ... :
*****************************************************************}
function UpdateYPL(TobYPL: Tob): integer;
var
  iYPL: integer;

  function MyUpdate(TobData: Tob): integer;
  var
    iChamp  : integer;
    Sql     : String;
    MyTobYPL: Tob;
    TomYPL  : Tom;
  begin
    Result := 0;

    TomYPL := CreateTOM('YPLANNING', nil, false, true);
    MyTobYPL := Tob.Create('YPLANNING', nil, -1);
    try
      Sql := 'SELECT ##TOP 1## *'
           + ' FROM YPLANNING'
           + ' WHERE ' + WhereYPL(Tobdata)
           ;
      if wSelectTobFromSQL(Sql, MyTobYPL, false) then
      begin
        { Sauvegarde de la Tob }
        TomYPL.LoadBufferAvantModif(MyTobYPL);

        { Saisi }
        for iChamp := 1000 to (1000 + TobData.ChampsSup.Count - 1) do
          MyTobYPL.PutValue(TobData.GetNomChamp(iChamp), TobData.GetValue(TobData.GetNomChamp(iChamp)));

        MyTobYPL.AddChampSupValeur('IKC', 'M', false);
        if TomYPL.VerifTOB(MyTobYPL) then
        begin
          if MyTobYPL.UpdateDb then      
          begin
            TomYPL.AfterVerifTOB(MyTobYPL);

            { Retour des informations }
            for iChamp := 1 to MyTobYPL.NbChamps do
              TobData.AddChampSupValeur(MyTobYPL.GetNomChamp(iChamp), MyTobYPL.GetValue(MyTobYPL.GetNomChamp(iChamp)));
          end;
        end
        else
          Result := TomYPL.lastError;
      end;
    finally
      MyTobYPL.free;
      TomYPL.free;
    end;
  end;
begin
  Result := 0;
  if not GestionYPlanning then exit;

  Result := ValideYPL(TobYPL);
  if Result = 0 then
    Result := MyUpdate(TobYPL)
  else if TobYPL.Detail.Count > 0 then
  begin
    for iYPL := 0 to TobYPL.Detail.Count-1 do
    begin
      Result := UpdateYPL(TobYPL.Detail[iYPL]);
      if Result <> 0 then
        Break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : suppression d'un enregistrement dans YPL
Description .. :
Mots clefs ... :
*****************************************************************}
function DeleteYPL(Const Prefixe, GuidOri: string): integer;
begin
  Result := 0;
  if not GestionYPlanning then exit;
  wDeleteTable('YPLANNING', WhereYPL('', Prefixe, GuidOri))
end;

function GetLastErrorMsgYPL(Const LastError: integer): String;
var
  TomYPL: Tom;
begin
  TomYPL := CreateTOM('YPLANNING', nil, false, true);
  try
    Result := Tom_YPlanning(TomYPL).GetLastErrorMsg(LastError);
  finally
    TomYPL.free;
  end;
end;

function GetFieldFromYPL(Const FieldName, Prefixe, GuidOri: string): variant;
begin
  Result := wGetSqlFieldValue(FieldName, 'YPLANNING', WhereYPL('', Prefixe, GuidOri));
end;

function LoadTobYPL(Const Where: String; TobYPL:Tob): Boolean;
var
  Sql: string;
begin
  Sql := 'SELECT *'
       + ' FROM YPLANNING'
       + ' WHERE ' + Where
       ;
  Result := TobYPL.LoadDetailDBFromSql('YPL', Sql);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 03/04/2007
Modifié le ... :
Description .. : teste si une période est libre
Description .. :
Description .. :
Suite ........ :
Mots clefs ... :
*****************************************************************}
function IsFreeYPL(pStTypePla : String; Const DateDebut, DateFin: tDateTime; Const Where: string): Boolean;
var
  Sql : string;
  //vStType : String;
  //vStCode : String;

begin

  result := true;
  if not GestionYPlanning then exit;

  sql := 'SELECT 1'
       + ' FROM YPLANNING'
       + ' WHERE YPL_DATEDEBUT<"' + UsDateTime(DateFin) + '"'
       + ' AND YPL_DATEFIN>"' + UsDateTime(DateDebut) + '"'
       + iif(Where <> '', ' AND (' + Where + ')', '')
       ;
  Result := not ExisteSql(Sql);

  {vStType := ReadYPLanning(pStTypePla);

  vStCode := (Trim(ReadTokenSt(vStType)));
  if vStCode <> '' then Sql := Sql + ' AND YPL_PREFIXE = "' + vStCode + '"';

  vStCode := (Trim(ReadTokenSt(vStType)));
  while (vStCode <> '') do
  begin
    Sql := Sql + ' OR YPL_PREFIXE = "' + vStCode + '"';
    vStCode := (Trim(ReadTokenSt(vStType)));
  end;}
                    
end;

{function IsFreeYPL(Const DateDebut, DateFin: tDateTime; Const Where: string): Boolean;
var
  Sql: string;
begin
  sql := 'SELECT 1'
       + ' FROM YPLANNING'
       + ' WHERE YPL_DATEDEBUT>="' + UsDateTime(DateDebut) + '"'
       + ' AND YPL_DATEFIN<="' + UsDateTime(DateFin) + '"'
       + iif(Where <> '', ' AND (' + Where + ')', '')
       ;
  Result := not ExisteSql(Sql);
end;}

{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 22/01/2007
Modifié le ... :
Description .. : création d'un ou plusieurs enregistrement(s) dans la table YPlanning
                 depuis une tob metier avec preparation de la tob a enregistrer
Mots clefs ... :
*****************************************************************}
function createTobYPL(pTob : Tob; pStPrefixe : String) : Integer;
begin
  result := CreateOrUpdateTobYPL(pTob, pStPrefixe, True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 22/01/2007
Modifié le ... :
Description .. : mise à jour d'un ou plusieurs enregistrement(s) dans la table YPlanning
                 depuis une tob metier avec preparation de la tob a enregistrer
Mots clefs ... :
*****************************************************************}
function updateTobYPL(pTob : Tob; pStPrefixe : String) : Integer;
begin
  result := CreateOrUpdateTobYPL(pTob, pStPrefixe, False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 22/01/2007
Modifié le ... :
Description .. : appel de createTobYPL ou updateTobYPL
Mots clefs ... :
*****************************************************************}
function CreateOrUpdateTobYPL(pTob : Tob; pStPrefixe : String; pBoCreate : Boolean) : Integer;
var
  i             : Integer;
  vTobYPL       : Tob;
  vTobYPLFille  : Tob;
  vInLastError  : Integer;

begin
  result := 0;
  if WriteYPLanning(pStPrefixe) then
  begin
    vTobYPL := TOB.create('#YPLANNING', nil, -1);
    try                   
      vInLastError := 0;
      if (pTob <> nil) and (pTob.Detail.Count > 0) then
      begin
        // prepare la tob a enregistrer
        for i := 0 to pTob.Detail.Count-1 do
        begin
          MoveCurProgressForm(); //20080828_C.B_GA15215
          vTobYPLFille := TOB.create('#YPLANNING', vTobYPL, -1);
          vInLastError := InitTobYPL(vTobYPLFille, pTob.Detail[i], pStPrefixe);

          if vInLastError <> 0 then
          begin
            result := -1;
            if vInLastError = 8 then
              PGIBox(format(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), [pStPrefixe]), traduireMemoire(GetLastErrorMsgYPL(7)))
            else
              PGIBox(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), traduireMemoire(GetLastErrorMsgYPL(7)));
            break;
          end;
        end;
      end
      //
      else if (pTob <> nil) then
      begin
        vInLastError := InitTobYPL(vTobYPL, pTob, pStPrefixe);
        if vInLastError <> 0 then
        begin
          result := -1;
          if vInLastError = 8 then
            PGIBox(format(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), [pStPrefixe]), traduireMemoire(GetLastErrorMsgYPL(7)))
          else
            PGIBox(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), traduireMemoire(GetLastErrorMsgYPL(7)));
        end;
      end;

      if vInLastError = 0 then
      begin
        if pBoCreate then
          result := CreateYPL(vTobYPL)
        else
          result := UpdateYPL(vTobYPL);
        if result <> 0 then
          PGIBox(traduireMemoire(GetLastErrorMsgYPL(result)), traduireMemoire(GetLastErrorMsgYPL(7)));
      end

    finally
      vTobYPL.Free;
    end;
  end
  else
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 15/02/2007
Modifié le ... : 15/02/2007
Description .. : initialisation de l'enregistrement YPLANNING pour tous
                 les produits
Mots clefs ... :
*****************************************************************}
function InitTobYPL(pTobYPL, pTob : Tob; pStPrefixe : String) : Integer;
var
  vStGuidYRS : String;
begin

  result := 8;

  // gestion d'affaire
  if (pStPrefixe = cStTableGa) then
  begin
    vStGuidYRS := GetYRS_GUID('', pTob.getValue('APL_RESSOURCE'), '');
    if pTob.getValue('APL_RESSOURCE') = '' then
      result :=  9

    else if vStGuidYRS = '' then
      result := 10
    else

    begin
      result := 0;
      pTobYPL.AddChampSupValeur('YPL_PREFIXE', 'APL');
      pTobYPL.AddChampSupValeur('YPL_GUIDORI', pTob.getValue('APL_GUID'));
      pTobYPL.AddChampSupValeur('YPL_GUIDYRS', vStGuidYRS);
      pTobYPL.AddChampSupValeur('YPL_DATEDEBUT', pTob.getValue('APL_HEUREDEB_PLA'));
      pTobYPL.AddChampSupValeur('YPL_DATEFIN', pTob.getValue('APL_HEUREFIN_PLA'));
      pTobYPL.AddChampSupValeur('YPL_LIBELLE', pTob.getValue('APL_LIBELLEPLA'));
      pTobYPL.AddChampSupValeur('YPL_ABREGE', pTob.getValue('APL_LIBELLEPLA'));
      pTobYPL.AddChampSupValeur('YPL_STATUTYPL', pTob.getValue('APL_ETATLIGNE'));
      pTobYPL.AddChampSupValeur('YPL_PRIVE', '-');
    end;
  end

  // grc
  else if (pStPrefixe = cStTableGrc) then
  begin

    vStGuidYRS := GetYRS_GUID('', pTob.getValue('RAI_RESSOURCE'), '');
    if pTob.getValue('RAI_RESSOURCE') = '' then
      result :=  9

    else if vStGuidYRS = '' then
      result := 10
    else

    begin
      result := 0;
      pTobYPL.AddChampSupValeur('YPL_PREFIXE', 'RAI');
      pTobYPL.AddChampSupValeur('YPL_GUIDORI', pTob.getValue('RAI_GUID'));
      pTobYPL.AddChampSupValeur('YPL_GUIDYRS', vStGuidYRS);
      pTobYPL.AddChampSupValeur('YPL_DATEDEBUT', pTob.getValue('DATEDEBUT'));
      pTobYPL.AddChampSupValeur('YPL_DATEFIN', pTob.getValue('DATEFIN'));
      pTobYPL.AddChampSupValeur('YPL_LIBELLE', pTob.getValue('RAC_LIBELLE'));
      pTobYPL.AddChampSupValeur('YPL_ABREGE', pTob.getValue('RAC_LIBELLE'));
      pTobYPL.AddChampSupValeur('YPL_STATUTYPL', pTob.getValue('RAC_ETATACTION'));
      pTobYPL.AddChampSupValeur('YPL_PRIVE', '-');
    end
  end

  // paie
  else if pStPrefixe = cStTablePaie then
  begin
    vStGuidYRS:= GetYRS_GUID (pTob.GetValue ('PCN_SALARIE'), '', '');
    if (pTob.getValue('PCN_SALARIE')='') then
      result :=  9

    else if vStGuidYRS = '' then
      result := 10
    else

    begin
      result := 0;
      pTobYPL.AddChampSupValeur('YPL_PREFIXE', 'PCN');
      pTobYPL.AddChampSupValeur('YPL_GUIDORI', pTob.getValue('PCN_GUID'));
      pTobYPL.AddChampSupValeur('YPL_GUIDYRS', vStGuidYRS);
      pTobYPL.AddChampSupValeur('YPL_DATEDEBUT', pTob.getValue('PCN_DATEDEBUTABS'));
      pTobYPL.AddChampSupValeur('YPL_DATEFIN', pTob.getValue('PCN_DATEFINABS'));
      pTobYPL.AddChampSupValeur('YPL_LIBELLE', pTob.getValue('PCN_LIBELLE'));
      pTobYPL.AddChampSupValeur('YPL_ABREGE', pTob.getValue('PCN_LIBELLE'));
      pTobYPL.AddChampSupValeur('YPL_STATUTYPL', pTob.getValue(''));
      pTobYPL.AddChampSupValeur('YPL_PRIVE', '-');
    end;
  end

  // bureau
  else if pStPrefixe = cStTableBureau then
  begin

    vStGuidYRS := GetYRS_GUID('', '', pTob.getValue('JEV_USER1'));

    if pTob.getValue('JEV_USER1') = '' then
      result :=  9

    else if vStGuidYRS = '' then
      result := 10
    else

    begin
      result := 0;
      pTobYPL.AddChampSupValeur('YPL_PREFIXE', 'JEV');
      pTobYPL.AddChampSupValeur('YPL_GUIDORI', pTob.getValue('JEV_GUIDEVT'));
      pTobYPL.AddChampSupValeur('YPL_GUIDYRS', vStGuidYRS);
      pTobYPL.AddChampSupValeur('YPL_DATEDEBUT', pTob.getValue('JEV_DATE'));
      pTobYPL.AddChampSupValeur('YPL_DATEFIN', pTob.getValue('JEV_DATEFIN'));
      pTobYPL.AddChampSupValeur('YPL_LIBELLE', pTob.getValue('JEV_EVTLIBELLE'));
      pTobYPL.AddChampSupValeur('YPL_ABREGE', pTob.getValue('JEV_EVTLIBABREGE'));
      pTobYPL.AddChampSupValeur('YPL_STATUTYPL', '');
      pTobYPL.AddChampSupValeur('YPL_PRIVE', '-');  
    end;
  end
                           
end;

{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 22/01/2007
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
function Recalcul_CreateTobYPL(pTob : Tob; pStPrefixe : String) : Integer;
var
  i             : Integer;
  vTobYPL       : Tob;
  vTobYPLFille  : Tob;
  vInLastError  : Integer;
  Heureact : TDateTime;
  Dureeact : Double;
  Heures,Minutes : integer;

begin
  result := 0;
  if WriteYPLanning(pStPrefixe) then
  begin
    vTobYPL := TOB.create('#YPLANNING', nil, -1);
    try
      vInLastError := 0;
      if pTob.Detail.Count > 0 then
      begin
        // prepare la tob a enregistrer
        for i := 0 to pTob.Detail.Count-1 do
        begin
          Heureact := 0;
          if pTob.Detail[i].getValue('RAC_HEUREACTION') <> IDate1900 then Heureact := StrToTime(pTob.Detail[i].getValue('RAC_HEUREACTION'));
          pTob.Detail[i].AddChampSupValeur('DATEDEBUT',pTob.Detail[i].GetValue('RAC_DATEACTION')+ HeureAct);
          Dureeact := pTob.Detail[i].getValue('RAC_DUREEACTION');
          Heures:=trunc(Dureeact/60);
          Minutes:=trunc(Dureeact-(Heures*60));
          pTob.Detail[i].AddChampSupValeur('DATEFIN',pTob.Detail[i].GetValue('DATEDEBUT')+EncodeTime(Heures,Minutes,0,0));

          vTobYPLFille := TOB.create('#YPLANNING', vTobYPL, -1);
          vInLastError := InitTobYPL(vTobYPLFille, pTob.Detail[i], pStPrefixe);

          if vInLastError <> 0 then
          begin
            result := -1;
            if vInLastError = 8 then
              PGIBox(format(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), [pStPrefixe]), traduireMemoire(GetLastErrorMsgYPL(7)))
            else
              PGIBox(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), traduireMemoire(GetLastErrorMsgYPL(7)));
            break;
          end;
        end;
      end
      //
      else if pTob <> nil then
      begin
        vInLastError := InitTobYPL(vTobYPL, pTob, pStPrefixe);
        if vInLastError <> 0 then
        begin
          result := -1;
          if vInLastError = 8 then
            PGIBox(format(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), [pStPrefixe]), traduireMemoire(GetLastErrorMsgYPL(7)))
          else
            PGIBox(traduireMemoire(GetLastErrorMsgYPL(vInLastError)), traduireMemoire(GetLastErrorMsgYPL(7)));
        end;
      end;

      if vInLastError = 0 then
      begin
        result := CreateYPL(vTobYPL);
        if result <> 0 then
          PGIBox(traduireMemoire(GetLastErrorMsgYPL(result)), traduireMemoire(GetLastErrorMsgYPL(7)));
      end

    finally
      vTobYPL.Free;
    end;
  end
  else
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 26/04/2007
Modifié le ... :
Description .. : recalcul la table YPLANNING à partir des tables
Suite ........ : AFPLANNING, ABSENCESALARIE, JUEVENEMENT et ACTIONINTERVENANT
Suite ........ : le recalcul est effectué pour chaque produit avec en paramètre
Suite ........ : la requête correspondante
Suite ........ : si on ne peut pas passer la requête Sql, on utilisera une tob
Mots clefs ... :
*****************************************************************}
function RecalculYPL(Const Prefixe: string; Const Sql: string=''; TobYPL: Tob=nil;Silence:Boolean=False) : Boolean;
var
  vTob      : Tob;
  vStTable  : String;

begin
  result := True;

  executeSql('DELETE FROM YPLANNING WHERE YPL_PREFIXE = "' + Prefixe + '"');

  if Sql <> '' then
  begin
    if Prefixe = cStTableGa then vStTable := ' AFPLANNING'
    else if Prefixe = cStTableGRC then vStTable := ' ACTIONINTERVENANT'
    else if Prefixe = cStTableBureau then vStTable := ' JUEVENEMENT'
    else if Prefixe = cStTablePaie then vStTable := ' ABSENCESALARIE';

    vTob := Tob.Create('#' + vStTable, nil, -1);
    try
      vTob.LoadDetailDBFromSql('#' + vStTable, Sql);

      if not Silence then
      begin
        SourisSablier;
        InitMoveProgressForm(nil, '',traduireMemoire(TexteMessage[9]), vTob.detail.count, False, True);
      end;

      if vTob.detail.count > 0 then
        result := createTobYPL(vTob, Prefixe) = 0; //C.B 06/05/2008 ajout du test pour stopper le traitement
 
    finally
      vTob.Free;
      if not Silence then
      begin
        //MoveCurProgressForm(traduireMemoire(TexteMessage[10])); //20080828_C.B_GA15215
        FiniMoveProgressForm;
        SourisNormale;
      end;
    end;
  end
  // A faire quand ce sera utile
  else
  begin
    if Prefixe = cStTableGRC then
      Recalcul_createTobYPL(TobYPL, Prefixe);

  end;

end;

end.
