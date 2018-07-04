{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/04/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTABLEAUTOWS ()
Mots clefs ... : TOF;BTTABLEAUTOWS
*****************************************************************}
Unit BTWSTABLEAUTO_TOF ;

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

function BLanceFiche_WSTableasAutorisees(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BTTABLEAUTOWS = Class (tTOFComm)
  private
    LstTables        : THGrid;
    TobTables        : TOB;
    ColsTables       : string;
    ColsFields       : string;
    HMTrad           : THSystemMenu;

    procedure GridsManagement;
    procedure LoadTobTable(Sender : TObject);
    procedure LoadTobFields(TableName : string);
    procedure LstTables_OnClick(Sender : TObject);
    procedure LstFields_OnClick(Sender : TObject);
    procedure LstTables_OnDlbclick(Sender : TObject);
    procedure AddField_OnClick(Sender : TObject);
    procedure AdvancedSetting_OnClick(Sender : TObject);

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
  , UConnectWSConst
  ;

function BLanceFiche_WSTableasAutorisees(Nat, Cod, Range,Lequel,Argument : string) : string;
begin
  V_PGI.ZoomOle := True;
  Result := AglLanceFiche(Nat, Cod, Range, Lequel, Argument);
  V_PGI.ZoomOle := False;
end;

procedure TOF_BTTABLEAUTOWS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTABLEAUTOWS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTTABLEAUTOWS.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTTABLEAUTOWS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTTABLEAUTOWS.OnArgument (S : String ) ;
var
  Sql : string;
  Cpt : integer;
  TobTableL : TOB;
begin
  Inherited ;
(*
  LstTables := THGrid(GetControl('LSTTABLES'));
  TobTables := TOB.Create('BTWSTABLEAUTO', nil, -1);
  Sql := 'SELECT BWT_NOMTABLE'
       + '     , DT_LIBELLE AS LABEL'
       + '     , "" AS DATA'
       + '     , BWT_AUTORISEE'
       + ' FROM BTWSTABLEAUTO'       + ' JOIN DETABLES ON DT_NOMTABLE = BWT_NOMTABLE'       + ' ORDER BY BWT_NOMTABLE';
  TobTables.LoadDetailDBFromSQL(Sql)
  for Cpt := 0 to pred(TobTables.Detail.count) do
  begin
    TobTableL := TobTables.Detail[Cpt];
    TobTableL.SetString('DATA', Tools.GetExtractTypeFromTableName(TobTableL.GetString('BWT_NOMTABLE')));
  end;
  LstTables.OnDblClick := LstTables_OnDlbclick;
//  GridsManagement;
*)
end ;

procedure TOF_BTTABLEAUTOWS.OnClose ;
begin
  Inherited ;
  FreeAndNil(TobTables);
end ;

procedure TOF_BTTABLEAUTOWS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTABLEAUTOWS.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTTABLEAUTOWS.LstTables_OnDlbclick(Sender: TObject);
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
  registerclasses ( [ TOF_BTTABLEAUTOWS ] ) ;
end.

