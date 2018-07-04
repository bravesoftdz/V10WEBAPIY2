{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CDETICKETS (CDETICKETS)
Mots clefs ... : TOM;CDETICKETS
*****************************************************************}
Unit CDETICKETS_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     Fiche, 
     FichList, 
     UTob,
     vierge ;

Type
  TOM_CDETICKETS = Class (TOM)
    private
    QMul : TQUERY;     // Query recuperee du mul
    public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

procedure TOM_CDETICKETS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnArgument ( S: String ) ;
var
  F                     : TFVierge ;
begin
  Inherited ;
if not (Ecran is TFVierge) then exit;
F:=TFVierge(Ecran) ;
if F <> NIL then
   BEGIN
{$IFDEF EAGLCLIENT}
  QMUL:=THQuery(F.FMULQ).TQ ;
{$ELSE}
  QMUL:=F.FMULQ ;
{$ENDIF}
   END ;
if QMUL = NIL then exit;
end ;

procedure TOM_CDETICKETS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CDETICKETS.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_CDETICKETS ] ) ; 
end.
