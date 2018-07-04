{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BRGPDREFERENTIEL ()
Mots clefs ... : TOF;BRGPDREFERENTIEL
*****************************************************************}
Unit BRGPDREFERENTIEL_TOF ;

Interface

Uses
  StdCtrls
  , Controls
  , Classes
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
  , HSysMenu
  ;

function BLanceFiche_RGPDReferentiel(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BRGPDREFERENTIEL = Class (tTOFComm)
  private
    Population       : THValCombobox;
    LstTables        : THGrid;
    LstFields        : THGrid;
    TobTables        : TOB;
    TobFields        : TOB;
    ColsTables       : string;
    ColsFields       : string;
    AddField         : TToolbarButton97;
    DelField         : TToolbarButton97;
    AdvancedSetting  : TToolbarButton97;
    NewField         : THEdit;
    HMTrad           : THSystemMenu;

    procedure GridsManagement;
    procedure ButtonManagement;
    procedure AdvancedSettingsManagment;
    procedure LoadTobTable(Sender : TObject);
    procedure LoadTobFields(TableName : string);
    procedure LstTables_OnClick(Sender : TObject);
    procedure LstFields_OnClick(Sender : TObject);
    procedure LstFields_OnDlbclick(Sender : TObject);
    procedure AddField_OnClick(Sender : TObject);
    procedure DelField_OnClick(Sender : TObject);
    procedure AdvancedSetting_OnClick(Sender : TObject);
    function GetTobFromGrid(CurrentGrid : THGrid) : TOB;

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  TntStdCtrls
  , wCommuns
  , LookUp
  , BRGPDUtils
  , ParamSoc
  , AglInit
  , BTCONFIRMPASS_TOF
  , Windows
  , UtilPGI
  , CommonTools
  ;

const
  LstNameTables = 'LSTTABLESL';
  LstNameFields = 'LSTCHAMPS';
  FieldColInternal = 1;
  FieldColField    = 2;
  FieldColLabel    = 3;
  FieldColExport   = 4;
  FieldColAnonym   = 5;

function BLanceFiche_RGPDReferentiel(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  V_PGI.ZoomOle := True;
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
  V_PGI.ZoomOle := False;
end;

procedure TOF_BRGPDREFERENTIEL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDREFERENTIEL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDREFERENTIEL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDREFERENTIEL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDREFERENTIEL.OnArgument (S : String ) ;
begin
  Inherited ;
  Population      := THValCombobox(GetControl('POPULATION'));
  LstTables       := THGrid(GetControl(LstNameTables));
  LstFields       := THGrid(GetControl(LstNameFields));
  AddField        := TToolbarButton97(GetControl('ADDFIELD'));
  DelField        := TToolbarButton97(GetControl('DELFIELD'));
  AdvancedSetting := TToolbarButton97(GetControl('PARAMAVANCE'));
  NewField        := THEdit(GetControl('NEWFIELD'));
  TobTables       := TOB.Create('_LinkedTables', nil, -1);
  TobFields       := TOB.Create('BRGPDCHAMPS', nil, -1);
  Population.OnChange     := LoadTobTable;
  LstTables.OnClick       := LstTables_OnClick;
  LstFields.OnClick       := LstFields_OnClick;
  LstFields.OnDblClick    := LstFields_OnDlbclick;
  AddField.OnClick        := AddField_OnClick;
  DelField.OnClick        := DelField_OnClick;
  AdvancedSetting.OnClick := AdvancedSetting_OnClick;
  GridsManagement;
  AdvancedSettingsManagment;
end ;

procedure TOF_BRGPDREFERENTIEL.OnClose ;
begin
  Inherited ;
  FreeAndNil(TobTables);
  FreeAndNil(TobFields);
end ;

procedure TOF_BRGPDREFERENTIEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDREFERENTIEL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BRGPDREFERENTIEL.LoadTobTable(Sender: TObject);
var
  Sql : string;
  Id  : string;
begin
  LstTables.VidePile(False);
  LstFields.VidePile(False);
  TobTables.ClearDetail;
  TobFields.ClearDetail;
  Id  := Population.Value;
  Sql := ' SELECT 1              as SORT'
       + '       , RG1_TABLENAME as TABLENAME'
       + '       , RG1_LABEL     as LABEL'
       + ' FROM BRGPDTABLESP'
       + ' WHERE RG1_ID = ' + Id
       + ' UNION'
       + ' SELECT 2              as SORT'
       + '       , RG2_NOMTABLE  as TABLENAME'
       + '       , RG2_LIBELLE   as LABEL'
       + ' FROM BRGPDTABLESL'
       + ' WHERE RG2_IDRG1 = ' + Id
       +   RGPDUtils.GetSqlTablesException
       + ' ORDER BY Sort, Label'
       ;
  TobTables.LoadDetailFromSQL(Sql);
  if TobTables.Detail.Count > 0 then
  begin
    LoadTobFields(TobTables.Detail[0].GetString('TableName'));
    TobTables.PutGridDetail(LstTables, False, False, ColsTables);
  end;
end;

procedure TOF_BRGPDREFERENTIEL.LoadTobFields(TableName : string);
var
  Sql : string;
begin
  LstFields.VidePile(False);
  TobFields.ClearDetail;
  if TobTables.Detail.Count > 0 then
  begin
    Sql := 'SELECT BRGPDCHAMPS.*, DH_LIBELLE'
         + ' FROM BRGPDCHAMPS'
         + ' JOIN DECHAMPS ON DH_NOMCHAMP=RG3_FIELDNAME'
         + ' WHERE RG3_TABLENAME = "' + TableName + '"'
         + ' ORDER BY RG3_FIELDNAME'
         ;
    TobFields.LoadDetailFromSQL(Sql);
    TobFields.PutGridDetail(LstFields, False, False, ColsFields);
    ButtonManagement;
  end;
end;

procedure TOF_BRGPDREFERENTIEL.GridsManagement;
var
  Cpt    : integer;

  procedure ManageBoolean;
  begin
    LstFields.ColWidths[Cpt] := 55 ;
    LstFields.ColAligns[Cpt]  := taCenter;
    LstFields.ColTypes[Cpt]   := 'B';
    LstFields.ColFormats[Cpt] := IntToStr(Integer(csCoche));
  end;

begin
  ColsTables := 'Label';
  ColsFields := 'RG3_INTERNE;RG3_FIELDNAME;DH_LIBELLE;RG3_EXPORT;RG3_RESET';
  for Cpt := 1 to 5 do
  begin
    case Cpt of
      FieldColInternal :  if RGPDUtils.AdvancedSettingsEnabled then
                          begin
                            LstFields.ColWidths[Cpt] := 55;
                            ManageBoolean;
                          end else
                            LstFields.ColWidths[Cpt] := -1;
      FieldColExport
      , FieldColAnonym :  ManageBoolean;
    end;
  end;
end;

procedure TOF_BRGPDREFERENTIEL.ButtonManagement;
var
  TobField : TOB;
begin
  TobField := GetTobFromGrid(LstFields);
  if assigned(TobField) then
  begin
    AddField.Enabled := (TobTables.Detail.Count > 0);
    DelField.Enabled := (not TobField.GetBoolean('RG3_INTERNE'));
  end;
end;

procedure TOF_BRGPDREFERENTIEL.AdvancedSettingsManagment;
begin
  TGroupBox(GetControl('GBFIELDS')).Height := iif(RGPDUtils.AdvancedSettingsEnabled, 400, 435);
  AddField.Visible        := (RGPDUtils.AdvancedSettingsEnabled);
  DelField.Visible        := (RGPDUtils.AdvancedSettingsEnabled);
  AdvancedSetting.Visible := (not RGPDUtils.AdvancedSettingsEnabled);
  GridsManagement;
  if Population.Value <> '' then
    LoadTobTable(Self);
  HMTrad.ResizeGridColumns(LstFields);
  LstFields.Refresh;
end;

function TOF_BRGPDREFERENTIEL.GetTobFromGrid(CurrentGrid : THGrid): TOB;
var
  TobCurrent : TOB;
begin
  if Assigned(CurrentGrid) then
  begin
    case Tools.CaseFromString(CurrentGrid.Name, [LstNameTables, LstNameFields]) of
      {LstNameTables} 0 : TobCurrent := TobTables;
      {LstNameFields} 1 : TobCurrent := TobFields;
    else
      TobCurrent := nil;
    end;
    if TobCurrent.Detail.Count > 0 then
      Result := TobCurrent.Detail[CurrentGrid.Row - 1]
    else
      Result := nil;
  end else
    Result := nil;
end;

procedure TOF_BRGPDREFERENTIEL.LstTables_OnClick(Sender: TObject);
var
  TobTable : TOB;
begin
  TobTable := GetTobFromGrid(LstTables);
  if Assigned(TobTable) then
    LoadTobFields(TobTable.GetString('TABLENAME'));
  ButtonManagement;
end;

procedure TOF_BRGPDREFERENTIEL.LstFields_OnClick(Sender : TObject);
begin
  ButtonManagement;
end;

procedure TOF_BRGPDREFERENTIEL.AddField_OnClick(Sender : TObject);
var
  TobTable  : TOB;
  TobField  : TOB;
  TableName : string;
  Prefix    : string;
  Where     : string;
  Cpt       : integer;
begin
  TobTable := GetTobFromGrid(LstTables);
  if Assigned(TobTable) then
  begin
    TableName := TobTable.GetString('TABLENAME');
    Prefix    := TableToPrefixe(TableName);
    Where     := '     DH_PREFIXE = "' + Prefix + '"'                                                         // Préfixe de la table en cours
               + ' AND DH_CONTROLE LIKE "%L%"'                                                                // Les champs ayant L dans DH_CONTROLE
               + ' AND DH_TYPECHAMP NOT IN ("BOOLEAN", "COMBO")'                                              // Pas de Boolean ou Combo
               + ' AND DH_NOMCHAMP <> (SELECT DT_CLE1 FROM DETABLES WHERE DT_NOMTABLE = "' + TableName + '")' // Pas champs de la clé 1
               + ' AND NOT EXISTS (SELECT 1 FROM BRGPDCHAMPS WHERE RG3_FIELDNAME = DH_NOMCHAMP)'              // Pas déjà présent dans le référentiel
               ;
    NewField.Text := '';
    LookupList(THCritMaskEdit(NewField), TraduireMemoire('Ajouter un champ'), 'DECHAMPS', 'DH_NOMCHAMP', 'DH_LIBELLE', Where, 'DH_NOMCHAMP', False, 0);
    if NewField.Text <> '' then
    begin
      TobField := TOB.Create('BRGPDCHAMPS', TobFields, -1);
      TobField.SetString('RG3_TABLENAME', TableName);
      TobField.SetString('RG3_FIELDNAME', NewField.Text);
      TobField.SetBoolean('RG3_INTERNE' , False);
      TobField.SetBoolean('RG3_EXPORT'  , True);
      TobField.SetBoolean('RG3_RESET'   , RGPDUtils.CanAnonymizableField(TableName, NewField.Text));
      TobField.InsertDB(nil);
      LoadTobFields(TableName);
      for Cpt := 0 to pred(LstFields.RowCount) do
      begin
        if LstFields.Cells[FieldColField, Cpt] = NewField.Text then
        begin
          LstFields.Col := FieldColField;
          LstFields.Row := Cpt;
          break;
        end;
      end;
    end;
  end;
end;

procedure TOF_BRGPDREFERENTIEL.DelField_OnClick(Sender: TObject);
var
  TobField  : TOB;
  Sql       : string;
begin
  TobField := GetTobFromGrid(LstFields);
  if Assigned(TobField) then
  begin
    if PGIAsk(Format(TraduireMemoire('Veuillez confirmer la suppression de %s'), [TobField.GetString('RG3_FIELDNAME')]), Ecran.Caption) = mrYes then
    begin
      Sql := 'DELETE FROM BRGPDCHAMPS'
           + ' WHERE RG3_TABLENAME = "' + TobField.GetString('RG3_TABLENAME') + '"'
           + '   AND RG3_FIELDNAME = "' + TobField.GetString('RG3_FIELDNAME') + '"'
           ;
      ExecuteSQL(Sql);
      LoadTobFields(TobField.GetString('RG3_TABLENAME'));
    end;
  end;
end;

procedure TOF_BRGPDREFERENTIEL.AdvancedSetting_OnClick(Sender : TObject);
begin
  if IsOkDayPass(GetMyComputerName) then
  begin
    SetParamSoc('SO_RGPDPARAMAVANCE', 'X');
    AdvancedSettingsManagment;
  end;
end;

procedure TOF_BRGPDREFERENTIEL.LstFields_OnDlbclick(Sender: TObject);
var
  TobField   : TOB;
  CurrentCol : integer;
  FieldName  : string;
  Sql        : string;
  TableValue : string;
  FieldValue : string;
begin
  CurrentCol := LstFields.Col;
  if (CurrentCol = FieldColExport) or (CurrentCol = FieldColAnonym) then
  begin
    TobField := GetTobFromGrid(LstFields);
    if Assigned(TobField) then
    begin
      case CurrentCol of
        FieldColExport : FieldName := 'RG3_EXPORT';
        FieldColAnonym : FieldName := 'RG3_RESET';
      end;
      TableValue := TobField.GetString('RG3_TABLENAME');
      FieldValue := TobField.GetString('RG3_FIELDNAME');
      if (FieldName = 'RG3_RESET') and (not RGPDUtils.CanAnonymizableField(TableValue, FieldValue)) then
        PGIError(TraduireMemoire('Vous ne pouvez pas anonymiser ce champ.'), Ecran.Caption)
      else
      begin
        TobField.SetBoolean(FieldName, iif(TobField.GetBoolean(FieldName), False, True));
        TobField.PutLigneGrid(LstFields, LstFields.Row, False, False, ColsFields);
        Sql := 'UPDATE BRGPDCHAMPS'
             + ' SET ' + FieldName + ' = "' + TobField.GetString(FieldName) + '"'
             + ' WHERE RG3_TABLENAME = "' + TableValue + '"'
             + '   AND RG3_FIELDNAME = "' + FieldValue + '"'
             ;
        ExecuteSQL(Sql);
      end;
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BRGPDREFERENTIEL ] ) ;
end.

