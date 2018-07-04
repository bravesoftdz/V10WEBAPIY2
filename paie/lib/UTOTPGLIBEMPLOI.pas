{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/05/2007
Modifié le ... :   /  /
Description .. : Source TOT PGLIBEMPLOI
Mots clefs ... :
*****************************************************************}
{

}
Unit UTOTPGLIBEMPLOI ;

Interface

Uses StdCtrls, Controls, Classes,forms,sysutils,ComCtrls,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,
{$ELSE}
       UTOB,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOT ;

Type
  TOT_PGLIBEMPLOI = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation
uses P5def;

procedure TOT_PGLIBEMPLOI.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBEMPLOI.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBEMPLOI.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBEMPLOI.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBEMPLOI.OnClose ;
begin
  Inherited ;
end ;

procedure TOT_PGLIBEMPLOI.OnArgument (S : String ) ;
begin
  Inherited ;
  PaieConceptTabMinPaie(Ecran);
end ;

Initialization
  registerclasses ( [ TOT_PGLIBEMPLOI ] ) ;
end.
