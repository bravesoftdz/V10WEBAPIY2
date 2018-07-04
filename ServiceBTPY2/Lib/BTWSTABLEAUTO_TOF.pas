unit BTWSTABLEAUTO_TOF;

interface

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

function BLanceFiche_WSTablesAutorisees(Nat, Cod, Range, Lequel, Argument : string) : string;

Type
  TOF_BTTABLEAUTOWS = Class (tTOFComm)
  private
    LstTables   : THGrid;
    TobTables   : TOB;
    ColsFields  : string;
    CheckNumCol : integer;

    procedure LstTables_OnDlbclick(Sender : TObject);

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

function BLanceFiche_WSTablesAutorisees(Nat, Cod, Range,Lequel,Argument : string) : string;
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
  LstTables  := THGrid(GetControl('LSTTABLES'));
  TobTables  := TOB.Create('BTWSTABLEAUTO', nil, -1);
  ColsFields := 'BWT_NOMTABLE;LABEL;DATA;BWT_AUTORISEE';
  Sql := 'SELECT BWT_NOMTABLE'
       + '     , DT_LIBELLE AS LABEL'
       + '     , "" AS DATA'
       + '     , BWT_AUTORISEE'
       + ' FROM BTWSTABLEAUTO'
       + ' JOIN DETABLES ON DT_NOMTABLE = BWT_NOMTABLE'
       + ' ORDER BY BWT_NOMTABLE';
  TobTables.LoadDetailDBFromSql('BTWSTABLEAUTO', Sql);
  for Cpt := 0 to pred(TobTables.Detail.count) do
  begin
    TobTableL := TobTables.Detail[Cpt];
    TobTableL.SetString('DATA', Tools.GetExtractTypeFromTableName(TobTableL.GetString('BWT_NOMTABLE')));
  end;
  CheckNumCol := 4;
  LstTables.OnDblClick := LstTables_OnDlbclick;
  LstTables.ColAligns[CheckNumCol]  := taCenter;
  LstTables.ColTypes[CheckNumCol]   := 'B';
  LstTables.ColFormats[CheckNumCol] := IntToStr(Integer(csCoche));
  TobTables.PutGridDetail(LstTables, False, True, ColsFields);
end ;

procedure TOF_BTTABLEAUTOWS.OnClose ;
begin
  Inherited ;
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
  TableName  : string;
  Sql        : string;
  NewValue   : Boolean;
begin
  CurrentCol := LstTables.Col;
  if CurrentCol = CheckNumCol then
  begin
    TobField := TobTables.Detail[LstTables.Row - 1];
    if Assigned(TobField) then
    begin
      NewValue  := iif(TobField.GetBoolean(FieldName), False, True);
      TableName := TobField.GetString('BWT_NOMTABLE');
      FieldName := 'BWT_AUTORISEE';
      TobField.SetBoolean(FieldName, iif(TobField.GetBoolean(FieldName), False, True));
      TobField.PutLigneGrid(LstTables, LstTables.Row, False, False, ColsFields);
      Sql := 'UPDATE BTWSTABLEAUTO'
           + ' SET ' + FieldName + ' = "' + TobField.GetString(FieldName) + '"'
           + '     , BWT_DATEMODIF   = "' + UsDateTime(Now) + '"'
           + '     , BWT_UTILISATEUR = "' + V_PGI.User + '"'
           + ' WHERE BWT_NOMTABLE   = "' + TobField.GetString('BWT_NOMTABLE') + '"'
           ;
      ExecuteSQL(Sql);
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTTABLEAUTOWS ] ) ;
end.

