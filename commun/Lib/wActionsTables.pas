unit wActionsTables;

interface

uses
  wRapport,
  uTob,
  EntGC,
  HEnt1,
  M3FP
  ,uEntCommun ,UtilTOBPiece
  ;

{ Existe }
function ExistsTable(Const TableName: String; TobData: Tob): Boolean;
{ Create }
function CreateTable    (Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function CreateTableZ   (Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function CreateInfoCompl(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function CreateTT(const Datatype, Code, Libelle, Abrege: String; TobResult: Tob): Boolean;
{ Update }
function UpdateTable    (Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function UpdateTableZ   (Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function UpdateInfoCompl(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
{ Delete }
function DeleteTable    (Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function DeleteTableZ   (Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
function DeleteInfoCompl(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
{ Duplic }
function DuplicTable(Const TableName: String; TobOldCle, TobNewCle: Tob; Rapport: TWRapport = nil): Boolean;

{ Fonctions spéciales pièces/lignes...(!) }
function UpdateInfosPiece(const CleDoc: R_CleDoc; const Argument: String; out LastErrorMsg: HString; const Separateur: String = ';'): Boolean;
function UpdateInfosLigne(const CleDoc: R_CleDoc; const Argument: String; out LastErrorMsg: HString; const Separateur: String = ';'): Boolean;

{ Mouvements }
{$IFDEF STK}
function CreateGSM(TobData: Tob; const Rapport: TWRapport = nil): Boolean;
function AffectATT(const QualifMvt, Guid: string; Const QPrevue: Double; Const RefAffectation, Emplacement, LotInterne: string; var LastErrorMsg: String): String;
{$ENDIF STK}

implementation

uses
  wCommuns
  ,Variants
  ,HCtrls
  ,uTom
  ,UtilArticle
  ,SysUtils
  ,UtilPgi
  ,wParamFonction
//GP_20080627_MM_GP15251
  {$IFDEF GPAO}
  ,wParamFonction_Tom
  {$ENDIF GPAO}
  ,FactTob
  ,EntRT
  ,ParamSoc
  ,wJetons
  ,FactUtil
  ,HMsgBox
  ,uTomComm
//GP20070228 GP14542 MM
//  ,yTarifs_TOM
  {$IFDEF STK}
    ,StkMouvement
  {$ENDIF STK}
  	,factcomm
   ,CbpMCD
   ,CbpEnumerator
  ;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/03/2005
Modifié le ... :   /  /    
Description .. : Teste d'existence d'un tuple
Mots clefs ... : 
*****************************************************************}
function ExistsTable(Const TableName: String; TobData: Tob): Boolean;
var
  sWhere, sSql: String;
begin
  if TableToNum(TableName) > 0 then
  begin
    sWhere := wMakeWhereSQL(wMakeFieldString(TableName, ';'), wGetValueClef1(TableName, TobData));
    if sWhere <> '' then
    begin
      sSql := 'SELECT 1'
            + ' FROM ' + TableName
            + ' WHERE ' + sWhere
            ;
      Result := ExisteSQL(sSql)
    end
    else
      Result := False
  end
  else
    Result := False
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. : Crée un enregistrement dans la table "TableName"
Suite ........ : Renvoie l'enregistrement inséré sous la TobData
Suite ........ : Utilise les TOM métiers
Mots clefs ... : CREATE;TOM;TABLE
*****************************************************************}
//GP_20071116_MM_GP14542 Déb
function CreateTable(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
var
  TomTABLE: Tom;
  TobTABLE: Tob;
  i: Integer;
  NomChamp: String;
  SuffixesToRaZ : string;
  Prefixe       : string;

  procedure SetError(const St: HString);
  begin
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end;

  function IsInitValue(const FieldName: String): Boolean;
  begin
    Case wGetSimpleTypeField(FieldName) of
      'I': Result := TobTABLE.GetInteger(FieldName) = 0;
      'N': Result := EgaliteDouble(TobTABLE.GetDouble(FieldName), 0);
      'D': Result := TobTABLE.GetDateTime(FieldName) = iDate1900;
      'B': Result := not TobTABLE.GetBoolean(FieldName);
    else
      Result := TobTABLE.GetString(FieldName) = ''
    end
  end;

  { Result := CHAMP1[;CHAMPN...] }
  function GetSuffixesToRAZ: String;
  begin
    if TableName = 'TTTTTTTT' then
      Result := 'CCCCCCCC'
    else
      Result := ''
  end;

begin
  Result := False;
  if TableToNum(TableName) > 0 then
  begin
    if not Result then
    begin
      TobTABLE := Tob.Create(TableName, nil, -1);
      TomTABLE := CreateTOM(TableName, nil, False, True);
      try
        { Argument (Spécifique à certaines tom uniquement) }
        if TableName = 'LISTEINVLIG' then
          TTomComm(TomTABLE).StArgument := GetArgumentByTob(TobData, True, True)
(*
				else if (TableName='YTARIFS') then
        begin
          TOM_yTarifs(TomTABLE).sAppel           := 'CBS';
          TOM_yTarifs(TomTABLE).sFonctionnalite  := TobData.GetString('YTS_FONCTIONNALITE');
          TOM_yTarifs(TomTABLE).lFicheEntete     := False;
          TOM_yTarifs(TomTABLE).lModifClefEntete := True;
          TOM_yTarifs(TomTABLE).lModifClefLignes := True;
          TOM_yTarifs(TomTABLE).lModifConditions := True;
        end
*)
//GP_20080117_DS_14895
        else if TableName = 'WGAMMELIG' then
          TTomComm(TomTABLE).StArgument := GetArgumentByTob(TobData, True, True)
//GP_20080627_MM_GP15251 Déb
        {$IFDEF GPAO}
        else if (TableName = 'WPARAMFONCTION') then
          TOM_WPARAMFONCTION(TomTABLE).StArgument := 'CODEFONCTION='+TobData.GetString('WPF_CODEFONCTION');
        {$ENDIF GPAO}
//GP_20080627_MM_GP15251 Fin
        ;

        { NewRecord }
        TomTABLE.InitTOB(TobTABLE);

        { Suppression des champs système + IDENTIFIANT & GUID + Compléments paramétrés ainsi que de champs particuliers en fonction de certaines tables}
        Prefixe := TableToPrefixe(TableName);
        SuffixesToRAZ := GetSuffixesToRAZ;
        SuffixesToRAZ := Prefixe+'_IDENTIFIANT;'
                        +Prefixe+'_GUID;'
                        +Prefixe+'_DATECREATION;'
                        +Prefixe+'_DATEMODIF;'
                        +Prefixe+'_CREATEUR;'
                        +Prefixe+'_UTILISATEUR'
                        +iif(SuffixesToRAZ<>'', ';', '')+SuffixesToRAZ;

        { Init. des données }
        for i := 1000 to 1000 + Pred(TobData.ChampsSup.Count) do
        begin
          NomChamp := TobData.GetNomChamp(i);
          if TobTABLE.FieldExists(NomChamp) and (TobTABLE.GetValue(NomChamp)<>TobData.GetValeur(i)) and (pos(NomChamp, SuffixesToRaZ)=0) then
            TobTABLE.PutValue(NomChamp, TobData.GetValeur(i))
          { Reporte les champs supp. }
          else if not TobTABLE.FieldExists(NomChamp) then
            TobTABLE.AddChampSupValeur(NomChamp, TobData.GetValeur(i))
        end;
//GP_20071116_MM_GP14542 Fin

        { Rapport }
        if Assigned(Rapport) then
          TobTABLE.AddChampSupValeur(WRPTobFieldName, LongInt(Rapport), False);

        { UpdateRecord }
        TobTABLE.AddChampSupValeur('IKC', 'C', False);
        TobTABLE.AddChampSupValeur('Error', '', False);
        Result := TomTABLE.VerifTOB(TobTABLE);
        if Result then
        begin
          Result := not ExistsTable(TableName, TobTABLE);
          if Result then
          begin
            Result := TobTABLE.InsertDB(nil);
            if Result then
            begin
              TomTABLE.AfterVerifTOB(TobTABLE);
              { Résultat }
              TobData.VirtuelleToReelle(TableName);
              TobData.InitValeurs();
              TobData.Dupliquer(TobTABLE, False, True);
              if Assigned(Rapport) then
                Rapport.Add(TWRP_Done, wGetValueClef1(TobData), TraduireMemoire('Enregistrement créé avec succès.'))
            end
            else
              SetError(TraduireMemoire('Echec lors de la création de l''enregistrement !'))
          end
          else
            SetError(TraduireMemoire('Cet enregistrement existe déjà !'))
        end
        else
          if TobTABLE.GetString('Error') <> '' then
            SetError(TobTABLE.GetString('Error'))
      finally
        TomTABLE.Free;
        TobTABLE.Free
      end
    end
  end
  else
    SetError(Format(TraduireMemoire('La table %s n''existe pas !'), [TableName]))
end;

{ Permet de créer des enregistrements dans les tables "Z" uniquement }
function CreateTableZ(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
var
  St: String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  if Mcd.getTable(TableName).Domain = 'Z' then
    Result := CreateTable(TableName, TobData, Rapport)
  else
  begin
    Result := False;
    St := TraduireMemoire('Vous n''avez pas le droit d''écrire dans cette table.');
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end
end;

{ Permet de créer des enregistrements dans les tablettes CHOIXCOD & CHOIXEXT }
function CreateTT(const Datatype, Code, Libelle, Abrege: String; TobResult: Tob): Boolean;
var
  iTT: Integer;
  TobTT: Tob;
  Prefixe, sTT: String;

  procedure SetError(const ErrorMsg: String);
  begin
    CreateTT := False;
    TobResult.AddChampSupValeur('Error', ErrorMsg, False)
  end;

begin
  Result := False;
  sTT := Datatype;
  iTT := TTToNum(sTT);
  { Contrôle existence de la tablette }
  if iTT > 0 then
  begin
    Prefixe := V_Pgi.DECombos[iTT].Prefixe;
    { Tablette paramétrable ? dans CHOIXEXT ou CHOIXCOD }
    if Pos(Prefixe + ';', 'CC;YX;') > 0 then
    begin
      { Données valides }
      if (Trim(Code) <> '') and (Trim(Libelle) <> '') then
      begin
        { Taille du code valide }
        if ((Prefixe = 'CC') and (Length(Trim(Code)) <= 3 ))
        or ((Prefixe = 'YX') and (Length(Trim(Code)) <= 17)) then
        begin
          TobTT := Tob.Create(PrefixeToTable(Prefixe), nil, -1);
          try
            TobTT.SetString(Prefixe + '_TYPE', V_Pgi.DECombos[iTT].Tipe);
            TobTT.SetString(Prefixe + '_CODE', UpperCase(Trim(Code)));
            { Existence de l'enregistrement }
            if not TobTT.ExistDB then
            begin
              TobTT.SetString(Prefixe + '_LIBELLE', Trim(Libelle));
              TobTT.SetString(Prefixe + '_ABREGE' , iif(Trim(Abrege) <> '', Trim(Abrege), Trim(Libelle)));
              if TobTT.InsertDB(nil) then
              begin
                Result := True;
                TobResult.Dupliquer(TobTT, False, True);
              end
              else
                SetError(TraduireMemoire('Echec lors de la création d''un enregistrement dans la tablette'))
            end
            else
              SetError(TraduireMemoire('Ce code existe déjà dans cette tablette'))
          finally
            TobTT.Free
          end
        end
        else
        begin
          if Prefixe = 'CC' then
            SetError(TraduireMemoire('Le code de cette tablette ne doit pas dépasser 3 caractères'))
          else
            SetError(TraduireMemoire('Le code de cette tablette ne doit pas dépasser 17 caractères'))
        end
      end
      else
        SetError(TraduireMemoire('Le code et le libellé de cette tablette doivent être renseignés'))
    end
    else
      SetError(TraduireMemoire('Cette tablette n''est pas paramétrable'))
  end
  else
    SetError(TraduireMemoire('Tablette inexistante'))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. : Met à jour un enregistrement dans la table "TableName"
Suite ........ : Renvoie l'enregistrement mis à jour sous la TobData
Suite ........ : Utilise les TOM métiers
Mots clefs ... : UPDATE;TOM;TABLE
*****************************************************************}
function UpdateTable(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
var
  TomTABLE: Tom;
  TobTABLE: Tob;
  i: Integer;
  sClef1,
  NomChamp,
  sFieldClef1: String;

  procedure SetError(const St: HString);
  begin
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end;

begin
  Result := False;
  if TableToNum(TableName) > 0 then
  begin
    if not Result then
    begin
      TobTABLE := Tob.Create(TableName, nil, -1);
      TomTABLE := CreateTOM(TableName, nil, False, True);
      try
        { Argument (Spécifique à certaines tom uniquement) }
        if TableName = 'LISTEINVLIG' then
          TTomComm(TomTABLE).StArgument := GetArgumentByTob(TobData, True, True);

        { Init. de la clef }
        sClef1 := wMakeFieldString(TableName, ';');
        while sClef1 <> '' do
        begin
          sFieldClef1 := ReadTokenSt(sClef1);
          if TobData.FieldExists(sFieldClef1) then
            TobTABLE.PutValue(sFieldClef1, TobData.GetValue(sFieldClef1));
        end;

        { Load }
        Result := TobTABLE.LoadDb();
        if Result then
        begin
          TomTABLE.LoadBufferAvantModif(TobTABLE);

          { Init. des données }
          for i := 1000 to 1000 + Pred(TobData.ChampsSup.Count) do
          begin
            NomChamp := TobData.GetNomChamp(i);
            if TobTABLE.FieldExists(NomChamp) and (TobTABLE.GetValue(NomChamp) <> TobData.GetValeur(i)) then
              TobTABLE.PutValue(NomChamp, TobData.GetValeur(i))
            { Reporte les champs supp. }
            else if not TobTABLE.FieldExists(NomChamp) then
              TobTABLE.AddChampSupValeur(NomChamp, TobData.GetValeur(i))
          end;

          { Rapport }
          if Assigned(Rapport) then
            TobTABLE.AddChampSupValeur(WRPTobFieldName, LongInt(Rapport), False);

          { UpdateRecord }
          TobTABLE.AddChampSupValeur('IKC', 'M', False);
          TobTABLE.AddChampSupValeur('Error', '', False);
          TobTABLE.SetDateModif(V_Pgi.DateEntree + Time);
          Result := TomTABLE.VerifTOB(TobTABLE);
          if Result then
          begin
            Result := TobTABLE.UpdateDB();
            if Result then
            begin
              TomTABLE.AfterVerifTOB(TobTABLE);
              { Résultat }
              TobData.VirtuelleToReelle(TableName);
              TobData.InitValeurs();
              TobData.Dupliquer(TobTABLE, False, True);
              if Assigned(Rapport) then
                Rapport.Add(TWRP_Done, wGetValueClef1(TobData), TraduireMemoire('Enregistrement créé avec succès.'))
            end
            else
              SetError(TraduireMemoire('Echec lors de la mise à jour de l''enregistrement'))
          end
          else
            if TobTABLE.GetString('Error') <> '' then
              SetError(TobTABLE.GetString('Error'))
        end
        else
          SetError(TraduireMemoire('Cet enregistrement n''existe pas ou la clé est incomplète !'));
      finally
        TomTABLE.Free;
        TobTABLE.Free
      end
    end
  end
  else
    SetError(Format(TraduireMemoire('La table %s n''existe pas !'), [TableName]));
end;

{ Permet de créer des enregistrements dans les tables "Z" uniquement }
function UpdateTableZ(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
var
  St: String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  if Mcd.getTable(TableName).Domain = 'Z' then
    Result := UpdateTable(TableName, TobData, Rapport)
  else
  begin
    Result := False;
    St := TraduireMemoire('Vous n''avez pas le droit d''écrire dans cette table.');
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. : Suppression d'un enregistrement de la table "TableName"
Suite ........ : Renvoie l'enregistrement supprimé
Suite ........ : Utilise les TOM métiers
Mots clefs ... : DELETE;TOM;TABLE
*****************************************************************}
function DeleteTable(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
var
  RapportNotAssigned: Boolean;
begin
  RapportNotAssigned := not Assigned(Rapport);

  if RapportNotAssigned then
    Rapport := twRapport.Create(TraduireMemoire('Suppression table'));
  try
    Result := wDeleteTable(TableName, wMakeWhereSQL(wMakeFieldString(TableName, ';'), wGetValueClef1(TableName, TobData)), False, Rapport);

    if not Result then
    begin
      TobData.AddChampSup('Error', false);
      if Rapport.CountError = 0 then
      begin
        if ExistsTable(TableName, TobData) then
          TobData.SetString('Error', TraduireMemoire('Erreur non définie !'))
        else
          TobData.SetString('Error', TraduireMemoire('Enregistement inexistant !'))
      end
      else
        TobData.SetString('Error', Rapport.GetFirstError)
    end;
  finally
    if RapportNotAssigned then
      FreeAndNil(Rapport);
  end;
end;

function DeleteTableZ(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
var
  St: String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  if Mcd.getTable(TableName).Domain = 'Z' then
    Result := DeleteTable(TableName, TobData, Rapport)
  else
  begin
    Result := False;
    St := TraduireMemoire('Vous n''avez pas le droit de supprimer des enregistrements de cette table.');
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /
Description .. : Duplication d'un enregistrement de la table "TableName"
Suite ........ : Renvoie l'enregistrement Créé
Suite ........ : Utilise les TOM métiers
Mots clefs ... : DUPLIC;TOM;TABLE
*****************************************************************}
function DuplicTable(Const TableName: String; TobOldCle, TobNewCle: Tob; Rapport: TWRapport = nil): Boolean;
var
  sClef1,
  Prefixe,
  SuffixesToRAZ,
//GP_20080414_TS_GP15004
  SuffixesToDEL,
  sField: String;
  TobOld,
  TobNew: Tob;

  procedure SetError(const St: HString);
  begin
    TobNewCle.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobNewCle), St)
  end;

//GP_20080414_TS_GP15004 >>>
  { Liste des champs à réinitialiser : (champs générés automatiquement par le système si non initialisés au préalable) }
  { Result := CHAMP1[;CHAMPN...] }
  function GetSuffixesToRAZ: String;
  begin
    if TableName = 'ARTICLE' then
      Result := 'CODEBARRE'
    else
      Result := ''
  end;

  { Liste des champs à ignorer totalement : (champs générés automatiquement et systématiquement par le système) }
  { Result := CHAMP1[;CHAMPN...] }
  function GetSuffixesToDEL: String;
  begin
    if TableName = 'WNOMELIG' then
      Result := 'LIENNOME'
    else if TableName = 'WGAMMELIG' then
      Result := 'NUMOPERGAMME'
// GC_20080124_GM_GP14599 DEBUT
    else if TableName = 'PIECEDA' Then
      Result := 'NUMERO;SOUCHE;CODESERVICE;RESPONSVAL;SERVICESUP;STATUTDA;DEPOT;DATELIVRAISON;DERNIERVAL;DERNIERSERV;PROCHAINSTATUT;ETABLISSEMENT;CTRLBUDGET;CODESESSION;CODECIRCUIT;CIRCUITCTRL;DATE;TOTALHT;'
// GC_20080124_GM_GP14599 FIN
    else
      Result := ''
  end;
//GP_20080414_TS_GP15004 <<<

  procedure AddChampsSupSpecifs;
  begin
    if TableName = 'ARTICLE' then
      TobNew.AddChampSupValeur('DUPLICART', TobOld.GetString('GA_ARTICLE'))
  end;

//GP_20080414_TS_GP15004 >>>
  { suppression des champs à ignorer lors de la duplication
   (champs générés automatiquement et systématiquement par le système) }
  procedure DeleteField(const Suffixe: String);
  begin
    if TobNew.FieldExists(Prefixe + '_' + Suffixe) then
      TobNew.DelChampSup(Prefixe + '_' + Suffixe, False)
  end;

  { initialisation des champs automatiquement calculés par le système mais non forcé
   (champs générés automatiquement par le système si non initialisés au préalable) }
  procedure InitField(const Suffixe: String);
  begin
    if TobNew.FieldExists(Prefixe + '_' + Suffixe) then
      TobNew.PutValue(Prefixe + '_' + Suffixe, wGetInitValue(Prefixe + '_' + Suffixe))
  end;
//GP_20080414_TS_GP15004 <<<

begin
  Result := False;

  if TableToNum(TableName) > 0 then
  begin
    Prefixe := TableToPrefixe(TableName);
    TobOld := Tob.Create(TableName, nil, -1);
    try
      sClef1 := wMakeFieldString(TableName, ';');
      while sClef1 <> '' do
      begin
        sField := ReadTokenSt(sClef1);
        if TobOldCle.FieldExists(sField) then
          TobOld.PutValue(sField, TobOldCle.GetValue(sField))
      end;
      if TobOld.LoadDB() then
      begin
        TobNew := Tob.Create(Prefixe, nil, -1);
        try
          wCreateChampSupFromTable(TobNew, TableName);
          AddChampsSupSpecifs();
          wCopyTobBySuffixe(TobOld, TobNew, Prefixe, Prefixe);

//GP_20080414_TS_GP15004 >>>
          { Suppression des champs système + IDENTIFIANT & GUID + Compléments paramétrés }
          SuffixesToDEL := GetSuffixesToDEL();
          SuffixesToDEL := 'IDENTIFIANT;GUID;DATECREATION;DATEMODIF;CREATEUR;UTILISATEUR'
                         + iif(SuffixesToDEL <> '', ';', '') + SuffixesToDEL;
          while (SuffixesToDEL <> '') do
            DeleteField(ReadTokenSt(SuffixesToDEL));

          { R.A.Z. des champs paramétrés }
          SuffixesToRAZ := GetSuffixesToRAZ();
          while (SuffixesToRAZ <> '') do
            InitField(ReadTokenSt(SuffixesToRAZ));
//GP_20080414_TS_GP15004 <<<

          wCopyTobBySuffixe(TobNewCle, TobNew, Prefixe, Prefixe);
          Result := CreateTable(TableName, TobNew, Rapport);
          TobNewCle.VirtuelleToReelle(TableName);
          TobNewCle.Dupliquer(TobNew, False, True)
        finally
          TobNew.Free
        end
      end
      else
        SetError(TraduireMemoire('L''enregistrement dupliqué n''existe pas ou sa clé est incomplète !'))
    finally
      TobOld.Free
    end
  end
  else
    SetError(Format(TraduireMemoire('La table %s n''existe pas !'), [TableName]))
end;

function DoUpdateInfosGP_GL(const Prefixe: String; const CleDoc: R_CleDoc; const Argument: String;
                            out LastErrorMsg: HString; const Separateur: String = ';'): Boolean;

  function GetSQLPieceLigne: String;
  begin
    if Prefixe = 'GP' then
      Result := 'SELECT * FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False)
    else
      Result := 'SELECT * FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, True, True)
  end;

  function EstVivante(T: Tob): Boolean;
  begin
    if Prefixe = 'GP' then
      Result := T.GetBoolean('GP_VIVANTE')
    else
      Result := GetEtatSolde(T) = eslEnCours
  end;

  function isChampTable(const Prefixe, Suffixe: String): Boolean;
  begin
    Result := ChampToNumDicho(TableToNum(PrefixeToTable(Prefixe)), Prefixe + '_' + Suffixe) >= 0
  end;

  procedure SetError(const Msg: HString);
  begin
    DoUpdateInfosGP_GL := False;
    if Prefixe = 'GP' then
      LastErrorMsg := TraduireMemoire('Mise à jour des informations de la pièce impossible.')+#13#10
    else
      LastErrorMsg := TraduireMemoire('Mise à jour des informations de la ligne impossible.')+#13#10;
    LastErrorMsg := LastErrorMsg + Msg
  end;

var
  TobArg, TobG, TobRD,
  TobChampsProFille,
  TobChampsProDescr: Tob;
  iField: Integer;
  FieldName, RefLigne: String;
  CharCompl: Char;
  PrefixeCompl: String;
const
  GP_AuthorizedFields = 'GP_BLOCNOTE;GP_LIBREPIECE1;GP_LIBREPIECE2;GP_LIBREPIECE3;GP_DATELIBREPIECE1;'
                      + 'GP_DATELIBREPIECE2;GP_DATELIBREPIECE3;';
  GL_AuthorizedFields = 'GL_BLOCNOTE;';
begin
  Result := True;
  CharCompl := iif(Prefixe = 'GP', 'D', '8');
  PrefixeCompl := iif(Prefixe = 'GP', 'RDD', 'RD8');
  TobChampsProFille := VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [CharCompl], False);
  TobG := Tob.Create(PrefixeToTable(Prefixe), nil, -1);
  TobRD := Tob.Create('RTINFOS00' + CharCompl, nil, -1);
  try
    { Chargement de la pièce }
    if wSelectTobFromSQL(GetSQLPieceLigne(), TobG) then
    begin
      if EstVivante(TobG) then
      begin
        TobG.Modifie := False;
        { Chargement (ou non : pas obligatoire) des infos compl }
        if Prefixe = 'GP' then
        begin
          if TobG.GetInteger('GP_NBTRANSMIS') > 0 then
          begin
            TobRD.SetString('RDD_CLEDATA', TobG.GetString('GP_NBTRANSMIS'));
            if TobRD.LoadDB() then
              TobRD.Modifie := False
          end
          else
            TobRD.SetString('RDD_CLEDATA', '0');
        end
        else
        begin
          RefLigne := EncodeRefPiece(TobG);
          TobRD.SetString('RD8_CLEDATA', RefLigne);
          if TobRD.LoadDB() then
            TobRD.Modifie := False
        end;

        { Chargement de l'argument dans une tob }
        TobArg := Tob.Create('_TobArguments_', nil, -1);
        try
          LoadTobFromArgument(TobArg, Argument, Separateur);
          iField := 1000;
          while Result and (iField < 1000 + TobArg.ChampsSup.Count) do
          begin
            FieldName := TobArg.GetNomChamp(iField);
            { Mise à jour des champs de la pièce }
            if isChampTable(Prefixe, ExtractSuffixe(FieldName)) then
            begin
              if Pos(FieldName + ';', iif(Prefixe = 'GP', GP_AuthorizedFields, GL_AuthorizedFields)) > 0 then
                TobG.PutValue(FieldName, TobArg.GetValeur(iField))
              else
                SetError(Format(TraduireMemoire('Mise à jour du champ %s interdite'), [FieldName]))
            end
            { Mise à jour des infos compl. de la pièce }
            else if isChampTable(PrefixeCompl, ExtractSuffixe(FieldName)) then
            begin
              if GetParamSocSecur('SO_RTGESTINFOS00' + CharCompl, False)
              and StrToBool_(GetInfoParPiece(CleDoc.NaturePiece, 'GPP_INFOSCPLPIECE'))
              and Assigned(TobChampsProFille) then
              begin
                TobChampsProDescr := TobChampsProFille.FindFirst(['RDE_NOMCHAMP'], [StringReplace(FieldName, PrefixeCompl + '_' + PrefixeCompl, 'RPR_RPR', [rfIgnoreCase])], False);
                if Assigned(TobChampsProDescr) then
                  TobRD.PutValue(FieldName, TobArg.GetValeur(iField))
                else
                  SetError(Format(TraduireMemoire('Mise à jour du champ %s non paramétrée'), [FieldName]))
              end
              else
                SetError(TraduireMemoire('Cette pièce ne gère pas les informations complémentaires'))
            end
            else
              SetError(Format(TraduireMemoire('Nom de champ incorrect : %s'), [FieldName]));

            Inc(iField)
          end
        finally
          TobArg.Free
        end
      end
      else
        SetError(Format(TraduireMemoire('La %s n''est pas modifiable'), [iif(Prefixe = 'GP', TraduireMemoire('pièce'), TraduireMemoire('ligne'))]));
    end
    else
      SetError(Format(TraduireMemoire('La %s n''existe pas'), [iif(Prefixe = 'GP', TraduireMemoire('pièce'), TraduireMemoire('ligne'))]));

    if Result then
    begin
      if (Prefixe = 'GP') and TobRD.Modifie then
      begin    
        if TobRD.GetInteger(PrefixeCompl + '_CLEDATA') > 0 then
          TobRD.UpdateDB()
        else
        begin
          TobRD.SetString('RDD_CLEDATA', IntToStr(wSetJeton('GP')));
          TobG.SetInteger('GP_NBTRANSMIS', TobRD.GetInteger('RDD_CLEDATA'));
          TobRD.InsertDB(nil)
        end
      end
      else if TobRD.Modifie then
        TobRD.InsertOrUpdateDB();

      if TobG.Modifie then
        TobG.UpdateDB()
    end
  finally
    TobRD.Free;
    TobG.Free
  end
end;

{ Mise à jour d'informations non sensibles dans les pièces + infos compl pièce }
function UpdateInfosPiece(const CleDoc: R_CleDoc; const Argument: String; out LastErrorMsg: HString; const Separateur: String = ';'): Boolean;
begin
  Result := DoUpdateInfosGP_GL('GP', CleDoc, Argument, LastErrorMsg, Separateur)
end;

{ Mise à jour d'informations non sensibles dans les lignes + infos compl ligne }
function UpdateInfosLigne(const CleDoc: R_CleDoc; const Argument: String; out LastErrorMsg: HString; const Separateur: String = ';'): Boolean;
begin
  Result := DoUpdateInfosGP_GL('GL', CleDoc, Argument, LastErrorMsg, Separateur)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /    
Description .. : Fonction commune à la création/modification d'infos compl.
Mots clefs ... :
*****************************************************************}
function CreateOrUpdateInfoCompl(Const TableName: String; TobData: Tob; Const Ikc: String = 'CM'; Rapport: TWRapport = nil): Boolean;

  procedure SetError(const St: HString);
  begin
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end;

var
  TobInfoCompl,
  TobChampsProFille: Tob;
  Prefixe, SocNom,
  FieldName        : String;
  NumTable, i      : Integer;
begin
  Result := False;
  NumTable := TableToNum(TableName);
  if NumTable > 0 then
  begin
    SocNom := StringReplace(TableName, 'RTINFOS', 'SO_RTGESTINFOS', [rfIgnoreCase]);
    { Test d'autorisation de gestion des informations complémentaires }
    Result := not ExisteSQL('SELECT 1 FROM PARAMSOC WHERE SOC_NOM="' + SocNom + '"') or GetParamSocSecur(SocNom, False);
    if Result then
    begin
      Prefixe := TableToPrefixe(TableName);
      TobChampsProFille := VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [wRight(Prefixe, 1)], False);
      { Test si le paramétrage de saisie des informations complémentaires a été fait }
      Result := Assigned(TobChampsProFille) and (TobChampsProFille.Detail.Count > 0);
      if Result then
      begin
        TobInfoCompl := Tob.Create(TableName, nil, -1);
        try
          if Pos('M', Ikc) > 0 then
          begin
            { Chargement de l'enregistrement si on est en mode modification }
            TobInfoCompl.SetString(Prefixe + '_CLEDATA', TobData.GetString(Prefixe + '_CLEDATA'));
            Result := TobInfoCompl.LoadDB();
            if not Result then
              SetError(TraduireMemoire('Cet enregistrement n''existe pas ou la clé est incomplète !'))
          end
          else { Mode création }
            Result := True;

          { Remplisage de la tob des informations complémentaires }
          if Result then
          begin
            i := 1000;
            while Result and (i < 1000 + TobData.ChampsSup.Count) do
            begin
              FieldName := TobData.GetNomChamp(i);
              Result := TobInfoCompl.FieldExists(FieldName);
              if Result then
              begin
                { Test d'autorisation de gestion du champ }
                Result := (Pos('_CLEDATA', FieldName) > 0) or
                          (TobChampsProFille.FindFirst(['RDE_NOMCHAMP'], [StringReplace(FieldName, Prefixe + '_' + Prefixe, 'RPR_RPR', [rfIgnoreCase])], False) <> nil);
                if Result then
                  TobInfoCompl.PutValue(FieldName, TobData.GetValeur(i))
                else
                  SetError(Format(TraduireMemoire('Champ non saisissable : %s'), [FieldName]))
              end
              else
                SetError(Format(TraduireMemoire('Nom de champ incorrect : %s'), [FieldName]));

              Inc(i)
            end;

            if Result then
            begin
              { Enregistrement des informations complémentaires }
              if Ikc = 'C' then
                Result := TobInfoCompl.InsertDB(nil)
              else if Ikc = 'M' then
                Result := TobInfoCompl.UpdateDB()
              else
                Result := TobInfoCompl.InsertOrUpdateDB();

              if not Result then
                SetError(TraduireMemoire('Erreur lors de l''enregistrement !'))
            end
          end
        finally
          TobInfoCompl.Free
        end
      end
      else
        SetError(TraduireMemoire('Le paramètrage de cette saisie n''a pas été effectué'))
    end
    else
      SetError(TraduireMemoire('Gestion des informations complémentaires non autorisée'))
  end
  else
    SetError(Format(TraduireMemoire('La table %s n''existe pas !'), [TableName]))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Création de données dans les tables RTx : Infos
Suite ........ : complémentaires.
Suite ........ : Renvoie l'enregistrement créé.
Mots clefs ... :
*****************************************************************}
function CreateInfoCompl(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
begin
  Result := CreateOrUpdateInfoCompl(TableName, TobData, 'C', Rapport)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Mise à jour d'un enregistrement dans les tables RTx :
Suite ........ : Infos complémentaires.
Suite ........ : Renvoie l'enregistrement mis à jour.
Mots clefs ... :
*****************************************************************}
function UpdateInfoCompl(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;
begin
  Result := CreateOrUpdateInfoCompl(TableName, TobData, 'M', Rapport)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Suppression d'un enregistrement dans les tables RTx :
Suite ........ : Infos complémentaires.
Mots clefs ... :
*****************************************************************}
function DeleteInfoCompl(Const TableName: String; TobData: Tob; Rapport: TWRapport = nil): Boolean;

  procedure SetError(const St: HString);
  begin
    TobData.AddChampSupValeur('Error', St);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, wGetValueClef1(TableName, TobData), St)
  end;

var
  Prefixe: String;
begin
  Result := False;
  if TableToNum(TableName) > 0 then
  begin
    Prefixe := TableToPrefixe(TableName);
    if TobData.GetString(Prefixe + '_CLEDATA') <> '' then
      Result := ExecuteSQL('DELETE FROM ' + TableName + ' WHERE ' + Prefixe + '_CLEDATA="' + TobData.GetString(Prefixe + '_CLEDATA') + '"') >= 0
    else
      SetError(Format(TraduireMemoire('Suppression impossible, vous devez renseigner %s'), [Prefixe + '_CLEDATA']))
  end
  else
    SetError(Format(TraduireMemoire('La table %s n''existe pas !'), [TableName]))
end;

{ Mouvements }

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 28/06/2007
Modifié le ... :   /  /
Description .. : Création de mouvements physiques en flux de stock
Mots clefs ... :
*****************************************************************}
{$IFDEF STK}
function CreateGSM(TobData: Tob; const Rapport: TWRapport = nil): Boolean;

  function SetError(const Error: String): Boolean;
  begin
    Result := False;
    TobData.AddChampSupValeur('Error', Error);
    if Assigned(Rapport) then
      Rapport.Add(TWRP_Error, TobData, Error)
  end;

//GP_20080409_TS_GP14986 >>>
var
  ActionGSM : TActionGSM;
  APR: Boolean;
begin
  ActionGSM := TActionGSM.Create;
  try
    TobData.ChangeParent(ActionGSM.TobData, -1);

    APR := TobData.GetString('GSM_QUALIFMVT') = 'APR'; //-> CBS affectation d'attendus de prod.
    if not APR and (Pos(GetStkTypeMvt(TobData.GetString('GSM_QUALIFMVT')), 'PHY;CET;') = 0) then
      Result := SetError(TraduireMemoire('Création de mouvement interdit avec ce type de mouvement.'))
    else if not APR and (GetStkFlux(TobData.GetString('GSM_QUALIFMVT')) <> 'STO') then
      Result := SetError(TraduireMemoire('Création de mouvement interdit avec ce flux.'))
    else
    begin
      V_PGI.IoError := Transactions(ActionGSM.DoCreate, 0);
      Result := (V_PGI.IoError = oeOk) and (not TobData.FieldExists('Error') or (TobData.GetString('Error') = ''));
      if (V_PGI.IoError <> oeOk) and not TobData.FieldExists('Error') then
        SetError(TraduireMemoire('Erreur de transaction.'))
    end;
  finally
    TobData.ChangeParent(nil, -1);
    ActionGSM.Free;
  end;
//GP_20080409_TS_GP14986 <<<
end;
{$ENDIF STK}

{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 06/10/2008
Modifié le ... : 06/10/2008
Description .. : Explose un mouvement attendus en fonction d'une quantité
Suite ........ : et de nouvelles valeurs (lot, série,
Suite ........ : emplacement, affectation, ...)
Suite ........ :
Suite ........ : >
Suite ........ :
Suite ........ : Retour: Guid du mouvement crée ou mis à jour
Mots clefs ... :
*****************************************************************}
function AffectATT(const QualifMvt, Guid: string; Const QPrevue: Double; Const RefAffectation, Emplacement, LotInterne: string; var LastErrorMsg: String): String;
var
  TobOrig, TobDest, t: Tob;
  Sql: string;
  ActionGSM : TActionGSM;
begin
  Result := '';
  LastErrorMsg := '';

  TobOrig := Tob.Create('GSM', nil, -1);
  TobDest := Tob.Create('GSM', nil, -1);
  ActionGSM := TActionGSM.Create;
  Try
    { Chargement de la tob origine }
    Sql := 'SELECT *'
         + ' FROM STKMOUVEMENT'
         + ' WHERE GSM_STKTYPEMVT="ATT"'
         + ' AND GSM_QUALIFMVT="' + QualifMvt + '"'
         + ' AND GSM_GUID="' + Guid + '"'
         ;
    if TobOrig.LoadDetailDbFromSql('GSM', Sql) then
    begin
      { Contrôles }
      if TobOrig.Detail[0].GetString('GSM_ETATMVT') = 'SOL' then
        LastErrorMsg := TraduireMemoire('Mouvement soldé. Changement d''affectation impossible')
      else if QPrevue > GetQBesoin(TobOrig.Detail[0]) then
        LastErrorMsg := TraduireMemoire('Quantité demandée trop importante')
      else if QPrevue <= 0 then
        LastErrorMsg := TraduireMemoire('Quantité demandée inférieure ou égale à 0')
      else if (TobOrig.Detail[0].GetString('GSM_REFAFFECTATION') = RefAffectation)
           and (TobOrig.Detail[0].GetString('GSM_EMPLACEMENT') = Emplacement)
           and (TobOrig.Detail[0].GetString('GSM_LOTINTERNE') = LotInterne)
           then
        LastErrorMsg := TraduireMemoire('demand esur la ligne d''origine')
      else
      begin
        Sql := 'SELECT ##TOP 1## *'
             + ' FROM STKMOUVEMENT'
             + ' WHERE GSM_STKTYPEMVT="ATT"'
             + ' AND GSM_QUALIFMVT="' + QualifMvt + '"'
             + ' AND GSM_GUID<>"' + Guid + '"'
             + ' AND ' + WhereGSM(GetRefOrigineGSM(TobOrig.Detail[0]))
             + iif(RefAffectation <> TobOrig.Detail[0].GetString('GSM_REFAFFECTATION'), ' AND GSM_REFAFFECTATION="' + RefAffectation + '"', '')
             + iif(Emplacement    <> TobOrig.Detail[0].GetString('GSM_EMPLACEMENT')   , ' AND GSM_EMPLACEMENT="'    + Emplacement    + '"', '')
             + iif(LotInterne     <> TobOrig.Detail[0].GetString('GSM_LOTINTERNE')    , ' AND GSM_LOTINTERNE="'     + LotInterne     + '"', '')
             ;
        BeginTrans;
        try
          { TobOrig }
          TobOrig.Detail[0].SetDouble('GSM_QPREVUE', TobOrig.Detail[0].GetDouble('GSM_QPREVUE') - QPrevue);
          TobOrig.Detail[0].ChangeParent(ActionGSM.TobData, -1);
          try
            if ActionGSM.TobData.Detail[0].GetDouble('GSM_QPREVUE') > 0 then
            begin
              { Le mouvement d'origine a toujours lieu d'être: Il faut le modifier }
              ActionGSM.DoUpdate;
            end
            else
            begin
              { le mouvement d'origine n'a plus lieu d'être: il faut le supprimer }
              ActionGSM.DoDelete;
            end;
          finally
            ActionGSM.TobData.Detail[0].ChangeParent(TobOrig, -1);
          end;

          { TobDest }
          if TobDest.LoadDetailDBFromSql('GSM', Sql) then
          begin
            { Le mouvement de destination existe déjà: Il faut le modifier }
            TobDest.Detail[0].SetDouble('GSM_QPREVUE', TobDest.Detail[0].GetDouble('GSM_QPREVUE') + QPrevue);
            TobDest.Detail[0].ChangeParent(ActionGSM.TobData, -1);
            try
              ActionGSM.DoUpdate;
            finally
              ActionGSM.TobData.Detail[0].ChangeParent(TobDest, -1);
            end;
          end
          else
          begin
            { Le mouvement de destination n'existe déjà: Il faut le créer }
            T := Tob.Create('GSM', TobDest, -1);
            T.Dupliquer(TobOrig.Detail[0], false, true);

            { Values }
            TobDest.Detail[0].SetString('GSM_EMPLACEMENT'   , Emplacement);
            TobDest.Detail[0].SetString('GSM_REFAFFECTATION', RefAffectation);
            TobDest.Detail[0].SetString('GSM_LOTINTERNE'    , LotInterne);

            TobDest.Detail[0].SetDouble('GSM_QPREVUE' , QPrevue);
            TobDest.Detail[0].SetDouble('GSM_QPREPA'  , 0.0);
            TobDest.Detail[0].SetDouble('GSM_PHYSIQUE', 0.0);

            TobDest.Detail[0].ChangeParent(ActionGSM.TobData, -1);
            try
              ActionGSM.DoCreate;
            finally
              ActionGSM.TobData.Detail[0].ChangeParent(TobDest, -1);
            end;
          end;

          CommitTrans;
        except
          RollBack;
        end;
      end;
    end
  Finally
    ActionGSM.Free;
    TobDest.Free;
    TobOrig.Free;
  End;
end;
{$ENDIF STK}

{ ************************************************************************************************ }
{                                        Fonctions utiles                                          }
{ ************************************************************************************************ }


function CheckTable(const TableName: String; TobData: Tob): Boolean;
var
  Prefixe: String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Result := (Mcd.getTable(TableName).Domain = 'Z');
  if not Result and Assigned(TobData) then
    TobData.AddChampSupValeur('Error', Format('Table %s non valide', [TableName]))
end;

function CheckTableZ(const TableName: String; TobData: Tob): Boolean;
var
  Prefixe: String;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Result := (Mcd.getTable(TableName).Domain = 'Z');
  if not Result and Assigned(TobData) then
    TobData.AddChampSupValeur('Error', Format('La table %s n''appartient pas au domaine "Z"', [TableName]))
end;

{ Droit }
function JaiLeDroitActionTable(const TableName: String; const ActionCUDX: Char): Boolean;
const
  StRights = 'GA=CUX;GIA=CUD;GTA=CUD;WAN=CUDX;WNT=CUDX;WNL=CUDX;WGT=CUDX;WGL=CUDX;WGR=CUDX;WOT=CUD'
           + ';WOL=CUD;WOP=CUD;WOB=CUD;WOG=CUD;WOR=CUD;GM=CUD;GCA=CUD;WPF=CUDX'
           + ';GNE=CUD;GNL=CUD;YTS=CUD;YTF=CUD;GQ=CU;Z=CU;QSM=CUDX;'
           ;
var
  Prefixe: String;
begin
  Prefixe := TableToPrefixe(TableName);
  if CheckTableZ(TableName, nil) then
    Prefixe := 'Z';
  Result := (Pos(ActionCUDX, GetArgumentString(StRights, Prefixe)) > 0) or (Pos(Prefixe + '=', StRights) = 0)
end;

{ Fonction générique de création de tuple }
function ActionTableByTob(const Action: Char; const TableName: String; TobData: Tob; const Domaine: Char = '*'): Boolean;

  procedure CheckTableWPF;
  begin
    if (TableName = 'WPARAMFONCTION') and (Pos(Action, 'UD') > 0) and (TobData.GetString('WPF_GUID') = '') then
      TobData.SetString('WPF_GUID', wGetSqlFieldValue('WPF_GUID', TableName, WhereWPF(TobData.GetString('WPF_CODEFONCTION'), GetCleWPF(GetArgumentByTob(TobData, False)))))
  end;

begin
  if JaiLeDroitActionTable(TableName, Action) then
  begin
    CheckTableWPF();

    case Action of
      'C': case Domaine of
             'Z': Result := CreateTableZ(TableName, TobData);
           else
             Result := CreateTable(TableName, TobData);
           end;
      'U': case Domaine of
             'Z': Result := UpdateTableZ(TableName, TobData);
           else
             Result := UpdateTable(TableName, TobData);
           end;
      'D': Result := DeleteTable(TableName, TobData);
    else
      Result := False
    end
  end
  else
    Result := False
end;

procedure CheckTobReel(TobData: Tob; var TobVirtual: Tob; const iTableOri: Integer);
begin
  if iTableOri > 0 then
  begin
    { Tob réelle à transformer en virtuelle }
    TobVirtual := Tob.Create(TableToPrefixe(TobData.NomTable), nil, -1);
    TobRealToVirtual(TobData, TobVirtual);
  end
  else
    TobVirtual := TobData
end;

procedure RestoreTobVirtual(TobData: Tob; var TobVirtual: Tob; const iTableOri: Integer);
begin
  if (iTableOri > 0) and (TobData <> TobVirtual) then
  begin
    TobData.Dupliquer(TobVirtual, false, true);
    FreeAndNil(TobVirtual)
  end
end;

{ ************************************************************************************************ }
{                                                                                                  }
{                                    PUBLICATIONS SCRIPT AGL                                       }
{                                                                                                  }
{ ************************************************************************************************ }

{***********A.G.L.***********************************************
Description .. : Création d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> Nom de la table
Suite ........ : Params[1] : POINTER -> TobData
*****************************************************************}
function AGLCreateTable(Params: array of variant; Nb: integer): variant;
var
  TobData, TobVirtual: Tob;
  TableName: String;
  iTableOri: Integer;
begin
  TobData := VarAsTob(Params[1]);
  iTableOri := TobData.NumTable;
  CheckTobReel(TobData, TobVirtual, iTableOri);
  try
    TableName := VarToStr(Params[0]);
    Result := CheckTable(TableName, TobVirtual) and ActionTableByTob('C', TableName, TobVirtual)
  finally
    RestoreTobVirtual(TobData, TobVirtual, iTableOri)
  end
end;

{***********A.G.L.***********************************************
Description .. : Modification d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> Nom de la table
Suite ........ : Params[1] : POINTER -> TobData
*****************************************************************}
function AGLUpdateTable(Params: array of variant; Nb: integer): variant;
var
  TobData, TobVirtual: Tob;
  TableName: String;
  iTableOri: Integer;
begin
  TobData := VarAsTob(Params[1]);
  iTableOri := TobData.NumTable;

  CheckTobReel(TobData, TobVirtual, iTableOri);
  try
    TableName := VarToStr(Params[0]);
    Result := CheckTable(TableName, TobVirtual) and ActionTableByTob('U', TableName, TobVirtual)
  finally
    RestoreTobVirtual(TobData, TobVirtual, iTableOri)
  end
end;

{***********A.G.L.***********************************************
Description .. : Duplication d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> Nom de la table
Suite ........ : Params[1] : POINTER -> TobData1 contenant la clé à dupliquer
Suite ........ : Params[2] : POINTER -> TobData2 contenant la clé dupliquée
*****************************************************************}
function AglDuplicTable(Params: Array of Variant; Nb: Integer): Variant;
var
  TableName: String;
  TobVirtualOld,
  TobVirtualNew,
  TobCleOld,
  TobCleNew: Tob;
  iTableOriOld,
  iTableOriNew: Integer;
begin
  TableName := VarToStr(Params[0]);
  TobCleOld := VarAsTob(Params[1]);
  TobCleNew := VarAsTob(Params[2]);
  iTableOriOld := TobCleOld.NumTable;
  iTableOriNew := TobCleNew.NumTable;

  CheckTobReel(TobCleOld, TobVirtualOld, iTableOriOld);
  CheckTobReel(TobCleNew, TobVirtualNew, iTableOriNew);
  try
    if Assigned(TobVirtualOld) and Assigned(TobVirtualNew) then
      Result := CheckTable(TableName, TobVirtualOld) and
                JaiLeDroitActionTable(TableName, 'X') and
                DuplicTable(TableName, TobVirtualOld, TobVirtualNew)
    else
      Result := False
  finally
    RestoreTobVirtual(TobCleOld, TobVirtualOld, iTableOriOld);
    RestoreTobVirtual(TobCleNew, TobVirtualNew, iTableOriNew);
  end
end;

{***********A.G.L.***********************************************
Description .. : Suppression d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> Nom de la table
Suite ........ : Params[1] : POINTER -> TobData
*****************************************************************}
function AGLDeleteTable(Params: array of variant; Nb: integer): variant;
var
  TableName: String;
  TobData: Tob;
begin
  TableName := VarToStr(Params[0]);
  TobData := VarAsTob(Params[1]);
  Result := CheckTable(TableName, TobData) and ActionTableByTob('D', TableName, TobData)
end;

function AGLExistsTable(Params: array of variant; Nb: integer): variant;
begin
  Result := ExistsTable(VarToStr(Params[0]), VarAsTob(Params[1]));
end;

{***********A.G.L.***********************************************
Description .. : Création d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> Nom de la table
Suite ........ : Params[1] : POINTER -> TobData
*****************************************************************}
function AglCreateTableZ(Params: Array of Variant; Nb: Integer): Variant;
var
  TableName: String;
  TobData, TobVirtual: Tob;
  iTableOri: Integer;
begin
  TableName := VarToStr(Params[0]);
  TobData := VarAsTob(Params[1]);
  iTableOri := TobData.NumTable;

  CheckTobReel(TobData, TobVirtual, iTableOri);
  try
    if Assigned(TobVirtual) then
      Result := CheckTableZ(TableName, TobVirtual) and ActionTableByTob('C', TableName, TobVirtual, 'Z')
    else
      Result := False
  finally
    RestoreTobVirtual(TobData, TobVirtual, iTableOri)
  end
end;

{***********A.G.L.***********************************************
Description .. : Modification d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> Nom de la table
Suite ........ : Params[1] : POINTER -> TobData (Virtuelle!)
*****************************************************************}
function AglUpdateTableZ(Params: Array of Variant; Nb: Integer): Variant;
var
  TableName: String;
  TobData, TobVirtual: Tob;
  iTableOri: Integer;
begin
  TableName := VarToStr(Params[0]);
  TobData := VarAsTob(Params[1]);
  iTableOri := TobData.NumTable;

  CheckTobReel(TobData, TobVirtual, iTableOri);
  try
    if Assigned(TobVirtual) then
      Result := CheckTableZ(TableName, TobVirtual) and ActionTableByTob('U', TableName, TobVirtual, 'Z')
    else
      Result := False
  finally
    RestoreTobVirtual(TobData, TobVirtual, iTableOri)
  end
end;

{ ************************** Création tablettes ************************* }

{***********A.G.L.***********************************************
Description .. : Modification d'un tuple dans la table concernée
Suite ........ : Params[0] : STRING -> DataType (Nom de la Tablette)
Suite ........ : Params[1] : STRING -> Code (3  caractères pour une tablette domaine CC (CHOIXCOD))
                                            (17 caractères pour une tablette domaine YX (CHOIXEXT))
Suite ........ : Params[2] : STRING -> Libellé
Suite ........ : Params[3] : STRING -> Abregé : si '' <=> Libellé tronqué
Suite ........ : Params[4] : POINTER -> TobResult (doit être créée avant appel afin de récupérer l'erreur dans le champ 'Error'
*****************************************************************}
function AglCreateTT(Params: Array of Variant; Nb: Integer): Variant;
begin
  Result := CreateTT(VarToStr(Params[0]), VarToStr(Params[1]), VarToStr(Params[2]),
                     iif(VarToStr(Params[3]) <> '', VarToStr(Params[3]), VarToStr(Params[2])), VarAsTob(Params[4]))
end;

{ Mise à jour des infos indicatives et non sensibles des pièces / lignes + infos compl. }

{***********A.G.L.***********************************************
Description .. : Mise à jour d'informations non sensibles dans
Suite ........ : les pièces + infos compl pièce
Suite ........ : Params[0] : STRING  : Nature de pièce
Suite ........ : Params[1] : STRING  : Souche de pièce
Suite ........ : Params[2] : INTEGER : N° de pièce
Suite ........ : Params[3] : INTEGER : Indice pièce
Suite ........ : Params[4] : STRING  : Argument contenant les MàJ
*****************************************************************}
function AglUpdateInfosPiece(Params: Array of Variant; Nb: Integer): Variant;
var
  CleDoc: R_CleDoc;
  ErrorMsg: HString;
begin
  CleDoc.NaturePiece := VarToStr(Params[0]);
  CleDoc.Souche := VarToStr(Params[1]);
  CleDoc.NumeroPiece := VarAsType(Params[2], varInteger);
  CleDoc.Indice := VarAsType(Params[3], varInteger);
  Result := UpdateInfosPiece(CleDoc, VarToStr(Params[4]), ErrorMsg);
  if not Result then
    PgiError(ErrorMsg)
end;

{***********A.G.L.***********************************************
Description .. : Mise à jour d'informations non sensibles dans
Suite ........ : les lignes + infos compl ligne
Suite ........ : Params[0] : STRING  : Nature de pièce
Suite ........ : Params[1] : STRING  : Souche de pièce
Suite ........ : Params[2] : INTEGER : N° de pièce
Suite ........ : Params[3] : INTEGER : Indice pièce
Suite ........ : Params[4] : INTEGER : Numéro d'ordre ligne
Suite ........ : Params[5] : STRING  : Argument contenant les MàJ
*****************************************************************}
function AglUpdateInfosLigne(Params: Array of Variant; Nb: Integer): Variant;
var
  CleDoc: R_CleDoc;
  ErrorMsg: HString;
begin
  CleDoc.NaturePiece := VarToStr(Params[0]);
  CleDoc.Souche := VarToStr(Params[1]);
  CleDoc.NumeroPiece := VarAsType(Params[2], varInteger);
  CleDoc.Indice := VarAsType(Params[3], varInteger);
  CleDoc.NumOrdre := VarAsType(Params[4], varInteger);
  Result := UpdateInfosLigne(CleDoc, VarToStr(Params[5]), ErrorMsg);
  if not Result then
    PgiError(ErrorMsg)
end;

{$IFDEF STK}
{***********A.G.L.***********************************************
Description .. : Création de mouvement PHY CET sur flux STO
Suite ........ : Params[0] : STRING  : QualifMvt
Suite ........ : Params[1] : STRING  : Dépôt
Suite ........ : Params[2] : STRING  : Article
Suite ........ : Params[3] : DOUBLE  : Physique
Suite ........ : Params[4] : POINTER : TobData
*****************************************************************}
function AglCreateGSM(Params: Array of Variant; Nb: Integer): Variant;
var
  TobData: Tob;
begin
  TobData := VarAsTob(Params[4]);
  TobData.AddChampSupValeur('GSM_QUALIFMVT', VarToStr(Params[0]));
  TobData.AddChampSupValeur('GSM_DEPOT'    , VarToStr(Params[1]));
  TobData.AddChampSupValeur('GSM_ARTICLE'  , VarToStr(Params[2]));
  TobData.AddChampSupValeur('GSM_PHYSIQUE' , VarAsType(Params[3], varDouble));
  Result := CreateGSM(TobData);
  if not Result then
    PgiError(TobData.GetString('Error'))
end;
{$ENDIF STK}

{$IFDEF STK}
{***********A.G.L.***********************************************
Description .. : Création de mouvement PHY CET sur flux STO
Suite ........ : Params[0] : STRING  : QualifMvt
Suite ........ : Params[1] : STRING  : Guyd
Suite ........ : Params[2] : DOUBLE  : QPrevue
Suite ........ : Params[3] : STRING  : RefAffectation
Suite ........ : Params[4] : STRING  : Emplacement
Suite ........ : Params[5] : STRING  : LotInterne
*****************************************************************}
function AglAffectATT(Params: Array of Variant; Nb: Integer): Variant;
var
  LastResult: String;
begin
  AffectATT(VarToStr(Params[0]), VarToStr(Params[1]), VarAsType(Params[2], varDouble), VarToStr(Params[3]), VarToStr(Params[4]), VarToStr(Params[5]), LastResult);
  if lastResult <> '' then
    PgiError(LastResult)
end;
{$ENDIF STK}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Création de données dans les tables RTx : Infos
Suite ........ : complémentaires.
Suite ........ : Renvoie l'enregistrement créé.
Suite ........ : Params[0] : STRING  : Nom table RTxxxxx
Suite ........ : Params[1] : POINTER : TobData
Mots clefs ... :
*****************************************************************}
function AglCreateInfoCompl(Params: Array of Variant; Nb: Integer): Variant;
begin
  Result := CreateInfoCompl(VarToStr(Params[0]), VarAsTob(Params[1]))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /    
Description .. : Mise à jour d'un enregistrement dans les tables RTx :
Suite ........ : Infos complémentaires.
Suite ........ : Renvoie l'enregistrement mis à jour.
Suite ........ : Params[0] : STRING  : Nom table RTxxxxx
Suite ........ : Params[1] : POINTER : TobData
Mots clefs ... :
*****************************************************************}
function AglUpdateInfoCompl(Params: Array of Variant; Nb: Integer): Variant;
begin
  Result := UpdateInfoCompl(VarToStr(Params[0]), VarAsTob(Params[1]))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 26/02/2007
Modifié le ... :   /  /    
Description .. : Suppression d'un enregistrement dans les tables RTx :
Suite ........ : Infos complémentaires.
Suite ........ : Params[0] : STRING  : Nom table RTxxxxx
Suite ........ : Params[1] : POINTER : TobData
Mots clefs ... :
*****************************************************************}
function AglDeleteInfoCompl(Params: Array of Variant; Nb: Integer): Variant;
begin
  Result := DeleteInfoCompl(VarToStr(Params[0]), VarAsTob(Params[1]))
end;

Initialization
  if FindAglFunc('AglCreateTable')      = nil then RegisterAglFunc('AglCreateTable'     , False, 2, AglCreateTable);
  if FindAglFunc('AglCreateInfoCompl')  = nil then RegisterAglFunc('AglCreateInfoCompl' , False, 2, AglCreateInfoCompl);
  if FindAglFunc('AglUpdateTable')      = nil then RegisterAglFunc('AglUpdateTable'     , False, 2, AglUpdateTable);
  if FindAglFunc('AglUpdateInfoCompl')  = nil then RegisterAglFunc('AglUpdateInfoCompl' , False, 2, AglUpdateInfoCompl);
  if FindAglFunc('AglDuplicTable')      = nil then RegisterAglFunc('AglDuplicTable'     , False, 3, AglDuplicTable);
  if FindAglFunc('AglDeleteTable')      = nil then RegisterAglFunc('AglDeleteTable'     , False, 2, AglDeleteTable);
  if FindAglFunc('AglDeleteInfoCompl')  = nil then RegisterAglFunc('AglDeleteInfoCompl' , False, 2, AglDeleteInfoCompl);
  if FindAglFunc('AglExistsTable')      = nil then RegisterAglFunc('AglExistsTable'     , False, 2, AglExistsTable);
  if FindAglFunc('AglCreateTableZ')     = nil then RegisterAglFunc('AglCreateTableZ'    , False, 2, AglCreateTableZ);
  if FindAglFunc('AglUpdateTableZ')     = nil then RegisterAglFunc('AglUpdateTableZ'    , False, 2, AglUpdateTableZ);
  if FindAglFunc('AglCreateTT')         = nil then RegisterAglFunc('AglCreateTT'        , False, 5, AglCreateTT);
  if FindAglFunc('AglUpdateInfosPiece') = nil then RegisterAglFunc('AglUpdateInfosPiece', False, 5, AglUpdateInfosPiece);
  if FindAglFunc('AglUpdateInfosLigne') = nil then RegisterAglFunc('AglUpdateInfosLigne', False, 6, AglUpdateInfosLigne);
{$IFDEF STK}
  if FindAglFunc('AglCreateGSM')        = nil then RegisterAglFunc('AglCreateGSM'       , False, 5, AglCreateGSM);
  if FindAglFunc('AglAffectATT')        = nil then RegisterAglFunc('AglAffectATT'       , False, 6, AglAffectATT);
{$ENDIF STK}  
end.
