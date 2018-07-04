{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/11/2015
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DEPOTS (DEPOTS)
Mots clefs ... : TOM;DEPOTS
*****************************************************************}
Unit DEPOTS_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,HDb,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     HCtrls,
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     HTB97,
     UTob ;

Type
  TOM_DEPOTS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;

  private

    CodeDepot : THEdit;

    BDelete   : TToolBarButton97;

    StSQL     : string;
    end ;

Implementation
uses Paramsoc,ENt1;

procedure TOM_DEPOTS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnDeleteRecord ;
begin
  Inherited ;

end ;

procedure TOM_DEPOTS.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnLoadRecord ;
begin
  Inherited ;

  //contrôle si code dépôt sur une pièce
  StSQL := 'SELECT GP_DEPOT FROM PIECE WHERE GP_DEPOT="' + CodeDepot.Text + '"';
  if ExisteSQL(StSQL) then Bdelete.Visible := False else Bdelete.Visible := True;

end ;

procedure TOM_DEPOTS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnArgument ( S: String ) ;
var codeAxe : string;
begin
  Inherited ;

  CodeDepot := THEdit(Getcontrol('GDE_DEPOT'));
  Bdelete   := TToolBarButton97(Getcontrol('BDELETE'));
  CodeAxe := GetparamSocSecur('SO_BTAXEANALSTOCK','TX1');
  if VH^.LiaisonY2ViaShare then
  begin
    THDbEdit(GetControl('GDE_SECTION')).DataType := 'TZCSECTION'+Copy(CodeAxe,2,1);
  end else
  begin
    THDbEdit(GetControl('GDE_SECTION')).DataType := 'TZSECTION'+Copy(CodeAxe,2,1);
  end;

end ;

procedure TOM_DEPOTS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_DEPOTS.OnAfterRollBack;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_DEPOTS ] ) ; 
end.
