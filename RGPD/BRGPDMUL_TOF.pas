{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
unit BRGPDMUL_TOF;

interface

uses
  StdCtrls
  , Controls
  , Classes
  , mul
  {$IFNDEF EAGLCLIENT}
  , db
  , uDbxDataSet
  , FE_Main
  {$ENDIF EAGLCLIENT}
  , uTob
  , forms
  , sysutils
  , ComCtrls
  , HCtrls
  , HEnt1
  , HMsgBox
  , UTOF
  , CBPMcd
  , Htb97
  , uTOFComm
  , BRGPDUtils
  , HDB
  , Menus
  ;

function BLanceFiche_RGPDThirdMul(Nat, Cod, Range, Lequel, Argument: string): string;

type
  TOF_BRGPDMUL = class(tTOFComm)
  private
    bOuvrir          : TToolbarButton97;
    bSelectAll       : TToolbarButton97;
    RgpdPopulation   : T_RGPDPopulation;
    FListe           : THDBGrid;
    TobRG2           : TOB;
    TobRG3           : TOB;
    PZoom            : TPopupMenu;
    mnJnalEvent      : TMenuItem;
    mnFormPopulation : TMenuItem;

    procedure bOuvrir_OnClick(Sender: TObject);
    procedure FListe_OnDblClick(Sender : TObject);
    procedure LoadTobRG2;
    procedure LoadTobRG3;
    function GetWhereTablesL(Where, FieldName: string): string;
    function GetSelectedFieldsFromTable(TableName: string): string;
    function GetExportAnonymizationWhere(KeyValue : string) : string;
    procedure InsertJnal(PathFile : string; AdditionalInformation : string='');
    procedure ExportDatas(PathFile : string);
    procedure Anonymization(PathFile : string);
    procedure Rectification(PathFile : string);
    procedure ConsentRequest(PathFile : string);
    procedure ConsentResponse(PathFile : string);
    procedure Zoom_OnClick(Sender : TObject);
//    procedure FListe_OnMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

  protected
    RgpdAction : T_RGPDActions;
    IsConsent  : Boolean;

  public
    sPopulationCode : string;
    sFieldCode      : string;
    sFieldCode2     : string;
    sFieldCode3     : string;
    sFieldLabel     : string;
    sFieldLabel2nd  : string;

    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  end;

implementation

uses
  TntStdCtrls
  , wCommuns
  , UtilPGI
  , BRGPDVALIDTRT_TOF
  , FormsName
  , Windows
  , ShellAPI
  , UtilGC
  , BTPUtil
  , TntDBGrids
  , ed_Tools
  , Grids
  , CommonTools
  ;

function BLanceFiche_RGPDThirdMul(Nat, Cod, Range, Lequel, Argument: string): string;
begin
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
end;

procedure TOF_BRGPDMUL.OnNew;
begin
  inherited;
end;

procedure TOF_BRGPDMUL.OnDelete;
begin
  inherited;
end;

procedure TOF_BRGPDMUL.OnUpdate;
begin
  inherited;
end;

procedure TOF_BRGPDMUL.OnLoad;
begin
  inherited;
end;

procedure TOF_BRGPDMUL.OnArgument(S: string);
begin
  inherited;
  fMulDeTraitement := True;
  TobRG2           := TOB.Create('BRGPDTABLESL', nil, -1);
  TobRG3           := TOB.Create('BRGPDTABLESL', nil, -1);
  RgpdAction       := RGPDUtils.GetActionFromCode(GetArgumentString(S, 'ACTION'));
  RgpdPopulation   := RGPDUtils.GetPopulationFromCode(sPopulationCode);
  FListe           := THDBGrid(GetControl('FListe'));
  bOuvrir          := TToolbarButton97(GetControl('BOUVRIR'));
  bSelectAll       := TToolbarButton97(GetControl('BSELECTALL'));
  PZoom            := TPopupMenu(GetControl('PZOOM'));
  mnJnalEvent      := TMenuItem(GetControl('mnJnalEvent'));
  mnFormPopulation := TMenuItem(GetControl('mnFormPopulation'));
  IsConsent        := ((RgpdAction = rgdpaConsentRequest) or (RgpdAction = rgdpaConsentResponse));
  bOuvrir.OnClick       := bOuvrir_OnClick;
  bOuvrir.Caption       := TraduireMemoire('Valider');
  FListe.MultiSelection := (RgpdAction = rgdpaConsentRequest);
  FListe.OnDblClick     := FListe_OnDblClick;
//  FListe.OnMouseDown    := FListe_OnMouseDown;
  TToolbarButton97(GetControl('BSELECTALL')).Visible := (RgpdAction = rgdpaConsentRequest);
  mnJnalEvent.Tag          := 0;
  mnJnalEvent.OnClick      := Zoom_OnClick;
  mnFormPopulation.Tag     := 1;
  mnFormPopulation.OnClick := Zoom_OnClick;
  mnFormPopulation.Caption := RGPDUtils.GetLabelFromPopulation(RgpdPopulation);
end;

procedure TOF_BRGPDMUL.OnClose;
begin
  inherited;
  FreeAndNil(TobRG2);
  FreeAndNil(TobRG3);
end;

procedure TOF_BRGPDMUL.OnDisplay();
begin
  inherited;
end;

procedure TOF_BRGPDMUL.OnCancel();
begin
  inherited;
end;

procedure TOF_BRGPDMUL.bOuvrir_OnClick(Sender: TObject);
var
  LastParam   : string;
  Params      : string;
  PathFile    : string;
  CanContinue : boolean;
begin
  if    (RgpdAction = rgdpaDataRectification)
     or ((RgpdAction <> rgdpaDataRectification) and (PGIAsk('Veuillez confirmer le traitement.', Ecran.Caption) = mrYes))
  then
  begin
    CanContinue := True;
    LastParam   := ';QUI='
                 + GetString(sFieldCode)
                 + iif(GetString(sFieldCode2) <> '', '|' + GetString(sFieldCode2), '')
                 + '~' + GetString(sFieldLabel)
                 + '~' + GetString(sFieldLabel2nd)
                 + ';CODEFILENAME=' + sFieldCode + iif(sFieldCode2 <> '', '|' + sFieldCode2, '');
    if RgpdAction = rgdpaConsentRequest then
    begin
      CanContinue := ((TFMul(Ecran).FListe.nbSelected > 0) or (TFMul(Ecran).FListe.AllSelected));
      LastParam   := LastParam + ';QTE=' + IntToStr(TFMul(Ecran).FListe.nbSelected)
    end;
    if CanContinue then
    begin
      Params := 'ORIGINE=' + RGPDUtils.GetCodeFromPopulation(RgpdPopulation) + ';ACTION=' + RGPDUtils.GetCodeFromAction(RgpdAction) + LastParam;
      PathFile := BLanceFiche_RGPDValidTrt('BTP', frm_RGPDTrtValid, '', '', Params);
      if ReadTokenSt(PathFile) = 'OK' then
      begin
        case RgpdAction of
          rgpdaDataExport        : ExportDatas(GetArgumentString(PathFile, 'INPUT'));
          rgpdaAnonymization     : Anonymization(GetArgumentString(PathFile, 'INPUT'));
          rgdpaDataRectification : Rectification(GetArgumentString(PathFile, 'INPUT'));
          rgdpaConsentRequest    : ConsentRequest(PathFile);
          rgdpaConsentResponse   : ConsentResponse(PathFile);
        end;
      end;
    end else
      PGIError(TraduireMemoire('Veuillez effectuer une sélection.'), Ecran.Caption);
  end;
end;

procedure TOF_BRGPDMUL.FListe_OnDblClick(Sender : TObject);
begin
  if RgpdAction <> rgdpaConsentRequest then
    bOuvrir_OnClick(Self);
end;
  
procedure TOF_BRGPDMUL.LoadTobRG2;
var
  Sql : string;
begin
  TobRG2.ClearDetail;
  Sql := 'SELECT RG1_TABLENAME'
       + '     , RG1_KEY'
       + iif(RgpdPopulation <> rgpdpContact, ', BRGPDTABLESL.*', '')
       + ' FROM BRGPDTABLESP'
       + iif(RgpdPopulation <> rgpdpContact, ' LEFT JOIN BRGPDTABLESL ON RG2_IDRG1 = RG1_ID' + RGPDUtils.GetSqlTablesException, '')
       + ' WHERE RG1_TABLENAME = "' + RGPDUtils.GetTableNameFromPopulation(RgpdPopulation) + '"';
  TobRG2.LoadDetailFromSQL(Sql);
end;

procedure TOF_BRGPDMUL.LoadTobRG3;
var
  Sql : string;

  function GetWhere: string;
  begin
    Result := ' WHERE RG1_TABLENAME = "' + RGPDUtils.GetTableNameFromPopulation(RgpdPopulation) + '"';
    case RgpdAction of
      rgpdaDataExport    : Result := Result + ' AND RG3_EXPORT = "X"';
      rgpdaAnonymization : Result := Result + ' AND RG3_RESET  = "X"';
    end;
  end;

begin
  TobRG3.ClearDetail;
  Sql := 'SELECT 1 AS SORT'
       + '     , BRGPDCHAMPS.*'
       + '     , DH_LIBELLE'
       + ' FROM BRGPDTABLESP'
       + ' JOIN BRGPDCHAMPS ON RG3_TABLENAME = RG1_TABLENAME'
       + ' JOIN DECHAMPS    ON DH_NOMCHAMP   = RG3_FIELDNAME' + GetWhere
       + 'UNION '
       + 'SELECT 2 AS SORT, BRGPDCHAMPS.*, DH_LIBELLE'
       + ' FROM BRGPDTABLESP'
       + ' JOIN BRGPDTABLESL ON RG2_IDRG1     = RG1_ID'
       + ' JOIN BRGPDCHAMPS  ON RG3_TABLENAME = RG2_NOMTABLE'
       + ' JOIN DECHAMPS     ON DH_NOMCHAMP   = RG3_FIELDNAME'
       + GetWhere;
  Sql := Sql + ' ORDER BY SORT, RG3_TABLENAME, RG3_FIELDNAME';
  TobRG3.LoadDetailFromSQL(Sql);
end;

function TOF_BRGPDMUL.GetWhereTablesL(Where, FieldName: string): string;
var
  Value       : string;
  ValueLength : Integer;
begin
  Result := Where;
  Value  := '|' + FieldName + '|';
  if Pos(Value, Where) > 0 then
  begin
    while Pos(Value, Result) > 0 do
    begin
      ValueLength := length(Value);
      Result := Copy(Result, 1, Pos(Value, Result) - 1) + GetString(FieldName) + Copy(Result, Pos(Value, Result) + ValueLength, Length(Result));
    end;
  end;
end;

function TOF_BRGPDMUL.GetSelectedFieldsFromTable(TableName: string): string;
var
  TobRG3L: TOB;
begin
  TobRG3L := TobRG3.FindFirst(['RG3_TABLENAME'], [TableName], True);
  while assigned(TobRG3L) do
  begin
    Result := Result + ',' + TobRG3L.GetString('RG3_FIELDNAME');
    TobRG3L := TobRG3.FindNext(['RG3_TABLENAME'], [TableName], True);
  end;
  Result := copy(Result, 2, length(Result));
end;

function TOF_BRGPDMUL.GetExportAnonymizationWhere(KeyValue : string) : string;
begin
  Result := GetWhereTablesL(KeyValue, sFieldCode);
  if sFieldCode2 <> '' then
    Result := GetWhereTablesL(Result, sFieldCode2);
  if sFieldCode3 <> '' then
    Result := GetWhereTablesL(Result, sFieldCode3);
end;


procedure TOF_BRGPDMUL.InsertJnal(PathFile : string; AdditionalInformation : string='');
var
  InfoJnal : string;
begin
  InfoJnal := RGPDUtils.GetLabelFromPopulation(RgpdPopulation)
            + ' '
            + GetString(sFieldCode)
            + iif(GetString(sFieldCode2)    <> '', ' - ' + GetString(sFieldCode2), '')
            + iif(GetString(sFieldLabel)    <> '', ' - ' + GetString(sFieldLabel), '')
            + iif(GetString(sFieldLabel2nd) <> '', ' ' + GetString(sFieldLabel2nd), '')
            + iif(AdditionalInformation     <> '', ' - ' + AdditionalInformation, '')
            ;
  MAJJnalEvent(RGPDCodeJnal, 'OK', RGPDUtils.GetLabelFromAction(RgpdAction), InfoJnal, PathFile);
end;

procedure TOF_BRGPDMUL.ExportDatas(PathFile : string);
var
  TobRG2L    : TOB;
  TobResult  : TOB;
  TobResultL : TOB;
  Cpt        : integer;
  Where      : string;
  TempPath   : string;
  TableName  : string;
  KeyValue   : string;
  tsFile     : TStringList;

  procedure AddTableValue(TableName, Where: string);
  var
    Sql        : string;
    FieldName  : string;
    FieldLabel : string;
    FieldsList : string;
    TobData    : TOB;
    CptData    : integer;
    CptFields  : integer;
  begin
    if (TableName <> '') and (Where <> '') then
    begin
      TobData := TOB.Create('_DATA', nil, -1);
      try
        FieldsList := GetSelectedFieldsFromTable(TableName);
        if FieldsList <> '' then
        begin
          Sql := 'SELECT ' + FieldsList + ' FROM ' + TableName + ' ' + Where;
          TobData.LoadDetailFromSQL(Sql);
          for CptData := 0 to pred(TobData.Detail.count) do
          begin
            for CptFields := 0 to pred(TobData.Detail[CptData].NombreChampSup) do
            begin
              TobResultL := TOB.Create('_VALUE', TobResult, -1);
              FieldName := TobData.Detail[CptData].GetNomChamp(1000 + CptFields);
              FieldLabel := TobRG3.FindFirst(['RG3_FIELDNAME'], [FieldName], True).GetString('DH_LIBELLE');
              TobResultL.AddChampSupValeur('_VALUE', TableName + ';' + FieldName + ';' + FieldLabel + ';' + TobData.Detail[CptData].GetString(FieldName));
            end;
          end;
        end;
      finally
        FreeAndNil(TobData);
      end;
    end;
  end;

begin
  TobResult := TOB.Create('_RESULT', nil, -1);
  try
    LoadTobRG2;
    LoadTobRG3;
    { Chargement des données }
    for Cpt := 0 to pred(TobRG2.detail.count) do
    begin
      TobRG2L   := TobRG2.Detail[Cpt];
      KeyValue  := TobRG2L.GetString(iif(Cpt = 0, 'RG1_KEY'      , 'RG2_FILTRE'));
      TableName := TobRG2L.GetString(iif(Cpt = 0, 'RG1_TABLENAME', 'RG2_NOMTABLE'));
      if KeyValue <> '' then
      begin
        Where := GetExportAnonymizationWhere(KeyValue);
        AddTableValue(TableName, ' WHERE ' + Where);
      end;
      Where := '';
    end;
    if TobResult.Detail.Count > 0 then
    begin
      TempPath := GetMyTempPath + GetString(sFieldCode) + '.CSV';
      tsFile := TStringList.Create;
      try
        tsFile.Add('TABLE;CHAMP;LIBELLE;VALEUR');
        for Cpt := 0 to pred(TobResult.Detail.Count) do
          tsFile.Add(TobResult.Detail[Cpt].GetString('_VALUE'));
        tsFile.SaveToFile(TempPath);
      finally
        FreeAndNil(tsFile);
      end;
      InsertJnal(PathFile);
      ShellExecute(0, pchar('open'), pchar(TempPath), nil, nil, SW_SHOW);
    end;
  finally
    FreeAndNil(TobResult);
  end;
  TobRG2.ClearDetail;
  TobRG3.ClearDetail;
end;

procedure TOF_BRGPDMUL.Anonymization(PathFile : string);
var
  TobResult : TOB;
  TobRG2L   : TOB;
  Cpt       : integer;
  Where     : string;
  TableName : string;
  KeyValue  : string;

  procedure AddUpdate(TableName, Where: string);
  var
    TobResultL : TOB;
    Sql        : string;
    FieldsList : string;
    FieldName  : string;
    FieldValue : string;
  begin
    if (TableName <> '') and (Where <> '') then
    begin
      Sql        := '';
      FieldsList := GetSelectedFieldsFromTable(TableName);
      if FieldsList <> '' then
      begin
        if ExisteSQL('SELECT 1 FROM ' + TableName + Where) then
        begin
          while FieldsList <> '' do
          begin
            FieldName := ReadTokenPipe(FieldsList, ',');
            case GetFieldType(FieldName) of
              ttfText            : FieldValue := 'iif(' + FieldName +  ' <> "", ' + '"' + wStringRepeat('Z', GetFieldSize(FieldName)) + '", "")';
              ttfInt, ttfNumeric : FieldValue := 'iif(' + FieldName +  ' <> 0, ' + '"0", "")';
              ttfMemo            : FieldValue := 'iif(LEN(CAST(' + FieldName +  ' AS NVARCHAR(1))) = 1, ' + '"' + wStringRepeat('Z', 35) + '", "")';
              ttfDate            : FieldValue := '"' + DateToStr(iDate1900) + '"';
            end;
            Sql := Sql + ', ' + FieldName + ' = ' + FieldValue;
          end;
          Sql := copy(Sql, 3, Length(Sql));
          { Ferme la fiche si nécessaire }
          case Tools.CaseFromString(TableName, ['TIERS', 'RESSOURCE', 'UTILISAT', 'SUSPECTS', 'CONTACT']) of
            {TIERS}     0 : Sql := Sql + ', T_FERME = "X", T_DATEFERMETURE = "' + UsDateTime(Date) + '", T_DATEMODIF = "' + UsDateTime(Date) + '", T_UTILISATEUR = "' + V_PGI.User + '"';
            {RESSOURCE} 1 : Sql := Sql + ', ARS_FERME = "X", ARS_DATEMODIF = "' + UsDateTime(Date) + '", ARS_UTILISATEUR = "' + V_PGI.User + '"';
            {UTILISAT}  2 : Sql := Sql + ', US_DESACTIVE = "X"';
            {SUSPECTS}  3 : Sql := Sql + ', RSU_FERME = "X", RSU_DATEFERMETURE = "' + UsDateTime(Date) + '", RSU_DATEMODIF = "' + UsDateTime(Date) + '", RSU_UTILISATEUR = "' + V_PGI.User + '"';
            {CONTACT}   4 : Sql := Sql + ', C_FERME = "X", C_DATEFERMETURE = "' + UsDateTime(Date) + '", C_DATEMODIF = "' + UsDateTime(Date) + '", C_UTILISATEUR = "' + V_PGI.User + '"';
          end;
          Sql := 'UPDATE ' + TableName + ' SET ' + Sql + Where;
          TobResultL := TOB.Create('_UPDATE', TobResult, -1);
          TobResultL.AddChampSupValeur('_UPDATE', Sql);
        end;
      end;
    end;
  end;

begin
  TobResult := TOB.Create('_RESULT', nil, -1);
  try
    LoadTobRG2;
    LoadTobRG3;
    { Chargement des données }
    for Cpt := 0 to pred(TobRG2.detail.count) do
    begin
      TobRG2L   := TobRG2.Detail[Cpt];
      KeyValue  := TobRG2L.GetString(iif(Cpt = 0, 'RG1_KEY'      , 'RG2_FILTRE'));
      TableName := TobRG2L.GetString(iif(Cpt = 0, 'RG1_TABLENAME', 'RG2_NOMTABLE'));
      if KeyValue <> '' then
      begin
        Where := GetExportAnonymizationWhere(KeyValue);
        AddUpdate(TableName, ' WHERE ' + Where);
      end;
      Where := '';
    end;
    BeginTrans;
    try
      { Anonymisation des tables }
      for Cpt := 0 to pred(TobResult.detail.count) do
        ExecuteSQL(TobResult.Detail[Cpt].GetString('_UPDATE'));
      { Alimentation du journal }
      InsertJnal(PathFile);
      CommitTrans;
      PGIInfo('Traitement effecuté avec succès.', Ecran.Caption);
    except
      Rollback;
      PGIInfo('Erreur durant l''exécution du traitement.', Ecran.Caption);
    end;
  finally
    FreeAndNil(TobResult);
  end;
  TFMul(Ecran).BChercheClick(nil);
end;

procedure TOF_BRGPDMUL.Rectification(PathFile : string);
var
  OkUpdate : string;
  Return   : string;  
begin
  case RgpdPopulation of
    rgpdpThird    : begin
                      Return   := GetString('T_TIERS');
                      OkUpdate := OpenForm.CliPro(GetString('T_AUXILIAIRE'), GetString('T_NATUREAUXI'));
                    end;
    rgpdpResource : begin
                      Return   := GetString('ARS_RESSOURCE');
                      OkUpdate := OpenForm.Resource(Return, '', 'ORIGINE=RGP');
                    end;
    rgpdpUser     : begin
                      Return   := GetString('US_UTILISATEUR');
                      OkUpdate := OpenForm.User(Return);
                    end;
    rgpdpSuspect  : begin
                      Return   := GetString('RSU_SUSPECT');
                      OkUpdate := OpenForm.Suspect(Return);
                    end;
    rgpdpContact  : begin
                      Return   := GetString('C_TYPECONTACT') + ';' + GetString('C_AUXILIAIRE') + ';' + GetString('C_NUMEROCONTACT');
                      OkUpdate := OpenForm.Contact(GetString('C_TYPECONTACT'), GetString('C_AUXILIAIRE'), GetInteger('C_NUMEROCONTACT'), 'FROMSAISIE;ORIGIN=RGPD');
                    end;
  end;
  if OkUpdate = Return then
   InsertJnal(PathFile);
end;

procedure TOF_BRGPDMUL.ConsentRequest(PathFile : string);
var
  NbSelected  : integer;
  Cpt         : integer;
  FieldKey    : string;
  TemplateDoc : string;
  PathFiles   : string;
  AllSelected : boolean;
  Form        : TFmul;

  procedure GenerateDocument;
  var
    Where       : string;
    FileName    : string;
    FieldValues : string;
    Value       : string;
    FieldValue  : string;
    CptValues   : integer;
    PosV        : integer;
  begin
    case RgpdPopulation of
      rgpdpThird    : FieldValue := Fliste.DataSource.DataSet.FindField('T_TIERS').AsString;
      rgpdpResource : FieldValue := Fliste.DataSource.DataSet.FindField('ARS_RESSOURCE').AsString;
      rgpdpUser     : FieldValue := Fliste.DataSource.DataSet.FindField('US_UTILISATEUR').AsString;
      rgpdpSuspect  : FieldValue := Fliste.DataSource.DataSet.FindField('RSU_SUSPECT').AsString;
      rgpdpContact  : FieldValue := Fliste.DataSource.DataSet.FindField('C_TYPECONTACT').AsString
                                  + '-' + Fliste.DataSource.DataSet.FindField('C_AUXILIAIRE').AsString
                                  + '-' + Fliste.DataSource.DataSet.FindField('C_NUMEROCONTACT').AsString;
    end;
    MoveCurProgressForm(Format('%s (%s/%s)', [FieldValue, IntToStr(Cpt + 1), IntToStr(NbSelected + 1)]));
    FileName := IncludeTrailingBackslash(PathFiles)
              + FormatDateTime('yyyymmdd',Date)
              + '_' + RGPDUtils.GetLabelFromPopulation(RgpdPopulation)
              + '_' + FieldValue + '.DOC';;
    if pos('|V1|', FieldKey) = 0 then
      Where := FieldKey + '"' + FieldValue + '"'
    else
    begin
      CptValues   := 0;
      Where       := FieldKey;
      FieldValues := FieldValue;
      while FieldValues <> '' do
      begin
         Inc(CptValues);                                                                                   
         Value := ReadTokenPipe(FieldValues, '-');
         PosV  := Pos('|V' + IntToStr(Cptvalues), Where) - 1;
         Where := Copy(Where, 1, PosV) + Value + Copy(Where, PosV + 5, Length(Where));
      end;
    end;
    { Pour les contacts, il faut trouver l'éventuelle 1ère adresse associée si demandé }
    if (RgpdPopulation = rgpdpContact) then
      RGPDUtils.MergExecuteContactAndAdressFromContact(Fliste.DataSource.DataSet.FindField('C_TIERS').AsString
                                                     , Fliste.DataSource.DataSet.FindField('C_NUMEROCONTACT').AsInteger
                                                     , TemplateDoc
                                                     , FileName
                                                     )
    { Pour les utilisateurs, il faut trouver l'adresse du salarié associée si existe }
    else if (RgpdPopulation = rgpdpUser) then
      RGPDUtils.MergExecuteUtilisatAndAdressFromSalarie(Fliste.DataSource.DataSet.FindField('US_UTILISATEUR').AsString
                                                      , TemplateDoc
                                                      , FileName
                                                      )
    else
      Publipost.MergeExecute(RGPDUtils.GetTableNameFromPopulation(RgpdPopulation), Where, TemplateDoc, FileName);
    InsertJnal(FileName);
  end;

begin
  Form        := TFMul(Ecran);
  AllSelected := FListe.AllSelected;
  NbSelected  := Pred(iif(AllSelected, Form.Q.RecordCount, FListe.nbSelected));
  InitMoveProgressForm(Ecran, 'Génération en cours.', Ecran.Caption, NbSelected + 1, False, True);
  try
    TemplateDoc := GetArgumentString(PathFile, 'INPUT');
    PathFiles   := GetArgumentString(PathFile, 'OUTPUT');
    case RgpdPopulation of
      rgpdpThird    : FieldKey := ' T_TIERS = ';
      rgpdpResource : FieldKey := ' ARS_RESSOURCE = ';
      rgpdpUser     : FieldKey := ' US_UTILISATEUR = ';
      rgpdpSuspect  : FieldKey := ' RSU_SUSPECT = ';
      rgpdpContact  : FieldKey := ' C_TYPECONTACT = "|V1|" AND C_AUXILIAIRE = "|V2|" AND C_NUMEROCONTACT = |V3|';
    end;
    if not AllSelected then
    begin
      for Cpt := 0 to NbSelected do
      begin
        FListe.GotoLeBookmark(Cpt);
        GenerateDocument;
      end;
    end else
    begin
      Cpt := 0;
      Form.Q.First;
      While not Form.Q.Eof do
      begin
         GenerateDocument;
         Form.Q.Next;
         Inc(Cpt);
      end;
    end;
  finally
    FiniMoveProgressForm;
    PGIInfo(Format('Document(s) généré(s) dans %s.', [PathFiles]), Ecran.Caption);
    Form.FListe.ClearSelected;
    if AllSelected then
    begin
      bSelectAll.Down := False;
      Form.FListe.AllSelected := False;
    end;
  end;
end;

procedure TOF_BRGPDMUL.ConsentResponse(PathFile: string);
var
  PathFiles  : string;
  Response   : string;
  Complement : string;
begin
  PathFiles := GetArgumentString(PathFile, 'INPUT');
  Response  := GetArgumentString(PathFile, 'RESPONSE');
  case Tools.CaseFromString(Response, ['V', 'R']) of
    {Validée} 0 : Complement := 'Demande validée.';
    {Refusée} 1 : Complement := 'Demande refusée.';
  end;
  InsertJnal(PathFiles, Complement);
  PGIInfo('Réponse enregistrée.', Ecran.Caption);
end;

procedure TOF_BRGPDMUL.Zoom_OnClick(Sender: TObject);
begin
  case THMenuItem(Sender).Tag of
    {JnalEvent}  0 : OpenForm.JnalEvent('TYPEEVENT=RGP;LABEL=' + RGPDUtils.GetLabelFromAction(RgpdAction));
    {Zoom fiche} 1 : case RgpdPopulation of
                       rgpdpThird    : OpenForm.CliPro(GetString('T_AUXILIAIRE'), GetString('T_NATUREAUXI'), '', True);
                       rgpdpUser     : OpenForm.User(GetString('US_UTILISATEUR'), True);
                       rgpdpResource : OpenForm.Resource(GetString('ARS_RESSOURCE'), '', 'ORIGINE=RGP', True);
                       rgpdpSuspect  : OpenForm.Suspect(GetString('RSU_SUSPECT'), True);
                       rgpdpContact  : OpenForm.Contact(GetString('C_TYPECONTACT'), GetString('C_AUXILIAIRE'), GetInteger('C_NUMEROCONTACT'), 'ORIGIN=RGPD', True);
                     end;
  end;
end;

(*
procedure TOF_BRGPDMUL.FListe_OnMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  TopCell : Integer;
  Line    : TGridRect;
  FListel : THGrid;
begin
  if (Shift = [ssRight]) then
  begin
    FListel := THGRid(TFMul(Ecran).FListe);
    TopCell := FListel.MouseCoord(X, Y).Y;
    if TopCell > 0 then
    begin
//      if (Shift = [ssRight]) then
//      begin
        Line.Left := 1;
        Line.Right := 1;
        Line.Top := TopCell;
        Line.Bottom := TopCell;
        FListel.Selection := Line;
        FListel.PopupMenu := TPopupMenu(GetControl('PZOOM'));
//      end
    end else
      FListel.PopupMenu :=  nil;
  end;
end;
*)

initialization
  registerclasses([TOF_BRGPDMUL]);
end.

