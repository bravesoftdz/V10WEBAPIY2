{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/06/2007
Modifié le ... :   /  /
Description .. : Source TOT de la TABLETTE : PGLIBGED (PGLIBGED)
Mots clefs ... : TOT;PGLIBGED
*****************************************************************}
Unit UTOTPGLIBGED;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HDB,
     HEnt1,
     HMsgBox,
     UTob,
     hqry,HTB97,
     UTOT ;

Type
  TOT_PGLIBGED = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
    Procedure Set3row;
  end ;

  TOT_PGLIBGED1 = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnDeleteRecord           ; override ;
  private
{$IFDEF EAGLCLIENT}
    Fliste   : THGrid;
{$ELSE}
    Fliste   : THDBGrid;
{$ENDIF}
    New : Boolean;
    IndexTable : Integer;
  end ;

  TOT_PGLIBGED2 = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnAfterUpdateRecord      ; override ;
  private
{$IFDEF EAGLCLIENT}
    Fliste   : THGrid;
{$ELSE}
    Fliste   : THDBGrid;
{$ENDIF}
    New : Boolean;
    IndexTable : Integer;
  end ;

  TOT_PGLIBGED3 = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnAfterUpdateRecord      ; override ;
  private
{$IFDEF EAGLCLIENT}
    Fliste   : THGrid;
{$ELSE}
    Fliste   : THDBGrid;
{$ENDIF}
    New : Boolean;
    IndexTable : Integer;
  end ;


Implementation

//uses DBGrids;

procedure ChangeTitle(Index : Integer;Const Ecran: TForm);
var
  Q : TQuery;
  begin
  if Assigned(Ecran) then
  begin
    //On récupère le libellé de la table libre
    Q := OpenSQL('Select CC_LIBELLE from CHOIXCOD where CC_TYPE="PG0" and CC_CODE ="00'+IntToStr(Index)+'"',True,1);
    Ecran.Caption := TraduireMemoire('Table libre ')+ IntToStr(Index)+ ' : ' + Q.Fields[0].AsString;
    Ferme(Q);
  end;
end;


procedure Initialise(IndexTable : Integer; Const Ecran: TForm;
                     {$IFNDEF EAGLCLIENT}var Fliste : THDBGrid;{$ELSE}var Fliste : THGrid;{$ENDIF}
                     var isNew : Boolean);
begin
  isNew := False;
{$IFNDEF EAGLCLIENT}
  Fliste   := THDBGrid(Ecran.FindComponent('FListe'));
{$ELSE}
  Fliste   := THGrid(Ecran.FindComponent('FListe'));
{$ENDIF}
  ChangeTitle(IndexTable,Ecran);
end;


procedure ShowMsgNewRecord(Const Ecran: TForm; var New : Boolean);
begin
  if Assigned(Ecran) then
  begin
    if New then
    begin
      PGIInfo(TraduireMemoire('Il est nécéssaire de mettre à jour la table des liaisons hiérarchique pour cette nouvelle valeur.'),Ecran.Caption);
      New := False;
    end;
  end;
end;

Function DeleteHierarchieAssocie(NumTable : Integer; IndexLigne : String) : Boolean;
var
  stSQL : String;
begin
  //On test si cette ligne est utilisé en tant que propriété pour certains documents
  if ExisteSQL('Select 1 from RTDOCUMENT where RTD_TABLELIBREGED'+IntToStr(NumTable)+' = "'+IndexLigne+'"') then
  begin
    result := False;
  end else begin
    result := True;
    stSQL := 'Delete from YDATATYPETREES where ';
    if NumTable = 1 then
      stSQL := stSQL + '(YDT_CODEHDTLINK = "PAIEGED12" and YDT_MCODE = "'+IndexLigne+'" )';
    if NumTable = 2 then
    begin
      stSQL := stSQL + '(   (YDT_CODEHDTLINK = "PAIEGED12" and YDT_SCODE = "'+IndexLigne+'" )';
      stSQL := stSQL + ' or (YDT_CODEHDTLINK = "PAIEGED23" and YDT_MCODE = "'+IndexLigne+'" ))';
    end;
    if NumTable = 3 then
      stSQL := stSQL + '(YDT_CODEHDTLINK = "PAIEGED23" and YDT_SCODE = "'+IndexLigne+'" )';
    ExecuteSQL(stSQL);
  end;
end;

{ TOT_PGLIBGED }

procedure TOT_PGLIBGED.OnNewRecord () ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBGED.OnDeleteRecord () ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBGED.OnUpdateRecord () ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBGED.OnAfterUpdateRecord () ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBGED.OnClose () ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBGED.OnArgument (S : String ) ;
Var
  CoControl: TControl;
begin
  Inherited ;
  if Ecran <> nil then
  begin
    CoControl := TControl(Ecran.FindComponent('BINSERT'));
    if CoControl <> nil then
      CoControl.Visible := FALSE;
    CoControl := TControl(Ecran.FindComponent('BDELETE'));
    if CoControl <> nil then
      CoControl.Visible := FALSE;
  end;
  Set3row;
end ;

procedure TOT_PGLIBGED.Set3row;
var
  TobLibGed, TobFilleLibGed : Tob;
  BoModified : Boolean;
  i :Integer;
begin
  BoModified := False;
  TobLibGed := TOB.Create('liste des tablettes GED',nil,-1);
  TobLibGed.LoadDetailDBFromSQL('CHOIXCOD','select * from CHOIXCOD where CC_TYPE="PG0"');
  //On met toujours 3 lignes dans la tablette
  For i := 1 to 3 do// TobLibGed.FillesCount(0) < 3 do
  begin
    if not assigned(TobLibGed.FindFirst(['CC_CODE'],['00'+IntToStr(i)],False)) then
    begin
      TobFilleLibGed := TOB.Create('CHOIXCOD',TobLibGed,-1);
      TobFilleLibGed.PutValue('CC_TYPE','PG0');
      TobFilleLibGed.PutValue('CC_CODE','00'+IntToStr(i));
      TobFilleLibGed.PutValue('CC_LIBELLE','');
      TobFilleLibGed.PutValue('CC_ABREGE','');
      BoModified := True;
    end;
  end;
  if TobLibGed.FillesCount(0) > 3 then
  begin
    ExecuteSQL('Delete CHOIXCOD where (CC_TYPE="PG0") and (CC_CODE not in ("001","002","003"))');
    BoModified := True;
  end;
  if BoModified then
  begin
    TobLibGed.InsertOrUpdateDB();
    AvertirTable('PGLIBGED');
  end;
  FreeAndNil(TobLibGed);
end;

{ TOT_PGLIBGED1 }

procedure TOT_PGLIBGED1.OnArgument(S: String);
begin
  inherited;
  IndexTable := 1;
  Initialise(IndexTable,Ecran,Fliste,New );
end;

procedure TOT_PGLIBGED1.OnDeleteRecord;
var
  IndexLigne : String;
begin
  inherited;
  if Ecran <> nil then
  begin
{$IFNDEF EAGLCLIENT}
    IndexLigne := Fliste.Fields[0].AsString;
{$ELSE}
    IndexLigne := Fliste.CellValues[0,Fliste.row];
{$ENDIF}
    if not DeleteHierarchieAssocie(IndexTable,IndexLigne) then
    begin
      LastErrorMsg := TraduireMemoire('Cette valeur a été utilisée pour la caractérisation d''un document.#10#13Vous ne pouvez pas la supprimer.');
      LastError := 1;
    end;
  end;
end;

procedure TOT_PGLIBGED1.OnNewRecord;
begin
  inherited;
  New := True;
end;

procedure TOT_PGLIBGED1.OnAfterUpdateRecord;
begin
  inherited;
  ShowMsgNewRecord(Ecran, New);
end;

{ TOT_PGLIBGED2 }

procedure TOT_PGLIBGED2.OnAfterUpdateRecord;
begin
  inherited;
  ShowMsgNewRecord(Ecran, New);
end;

procedure TOT_PGLIBGED2.OnArgument(S: String);
begin
  inherited;
  IndexTable := 2;
  Initialise(IndexTable,Ecran,Fliste,New);
end;

procedure TOT_PGLIBGED2.OnDeleteRecord;
var
  IndexLigne : String;
begin
  inherited;
  if Ecran <> nil then
  begin
{$IFNDEF EAGLCLIENT}
    IndexLigne := Fliste.Fields[0].AsString;
{$ELSE}
    IndexLigne := Fliste.CellValues[0,Fliste.row];
{$ENDIF}
    if not DeleteHierarchieAssocie(IndexTable,IndexLigne) then
    begin
      LastErrorMsg := TraduireMemoire('Cette valeur a été utilisée pour la caractérisation d''un document.#10#13Vous ne pouvez pas la supprimer.');
      LastError := 1;
    end;
  end;
end;

procedure TOT_PGLIBGED2.OnNewRecord;
begin
  inherited;
  New := True;
end;

{ TOT_PGLIBGED3 }

procedure TOT_PGLIBGED3.OnAfterUpdateRecord;
begin
  inherited;
  ShowMsgNewRecord(Ecran, New);
end;

procedure TOT_PGLIBGED3.OnArgument(S: String);
begin
  inherited;
  IndexTable := 3;
  Initialise(IndexTable,Ecran,Fliste,New);
end;

procedure TOT_PGLIBGED3.OnDeleteRecord;
var
  IndexLigne : String;
begin
  inherited;
  if Ecran <> nil then
  begin
{$IFNDEF EAGLCLIENT}
    IndexLigne := Fliste.Fields[0].AsString;
{$ELSE}
    IndexLigne := Fliste.CellValues[0,Fliste.row];
{$ENDIF}
    if not DeleteHierarchieAssocie(IndexTable,IndexLigne) then
    begin
      LastErrorMsg := TraduireMemoire('Cette valeur a été utilisée pour la caractérisation d''un document.#10#13Vous ne pouvez pas la supprimer.');
      LastError := 1;
    end;
  end;
end;

procedure TOT_PGLIBGED3.OnNewRecord;
begin
  inherited;
  New := True;
end;

Initialization
  registerclasses ( [ TOT_PGLIBGED,TOT_PGLIBGED1,TOT_PGLIBGED2,TOT_PGLIBGED3 ] ) ;
end.
