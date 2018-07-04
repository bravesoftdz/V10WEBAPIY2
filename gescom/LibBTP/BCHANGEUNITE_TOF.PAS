{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/09/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BCHANGEUNITE ()
Mots clefs ... : TOF;BCHANGEUNITE
*****************************************************************}
Unit BCHANGEUNITE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UCoefsUnite,
     UTOF ;

Type
  TOF_BCHANGEUNITE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    GS : THGrid;
    TOBLocals : TOB;
    TOBUnites : TOB;
    ChangeUnite : TChangeUnite;
    procedure   Definigrid;

  end ;

Implementation
uses SAISUTIL,Math;

procedure TOF_BCHANGEUNITE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BCHANGEUNITE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BCHANGEUNITE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BCHANGEUNITE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BCHANGEUNITE.OnArgument (S : String ) ;
begin
  Inherited ;
  ChangeUnite := TChangeUnite.create;
  TOBLocals := ChangeUnite.TOBCoefs;
  TOBUnites := ChangeUnite.TOBunites;
  //
  GS := THGrid(GetControl('GS'));
  Definigrid;
end ;

procedure TOF_BCHANGEUNITE.OnClose ;
begin
  Inherited ;
  ChangeUnite.Free;
end ;

procedure TOF_BCHANGEUNITE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BCHANGEUNITE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BCHANGEUNITE.Definigrid;
var indice,Col,decalage : Integer;
begin
	GS.ColCount := TOBUnites.Detail.Count + 1;
	GS.RowCount := TOBUnites.Detail.Count + 1;
	for Indice := 0 to TOBUnites.detail.count -1 do
  begin
		GS.cells[0,Indice+1] := RechDom('GCQUALUNITPIECE',TOBUnites.detail[Indice].GetString('GME_MESURE'),false);
    GS.Cells[Indice+1,0] := RechDom('GCQUALUNITPIECE',TOBUnites.detail[Indice].GetString('GME_MESURE'),false);
		for Col := 0 to TOBUnites.Detail.count -1 do
    begin
      if Indice > 0 then Decalage := trunc(Power(TOBUnites.Detail.Count,Indice))+Col+1
      							else decalage := Col;
    	GS.Cells[Indice+1,Col+1] := Strf00(TOBLocals.detail[Decalage].GetDouble('BCU_COEFCHANGE'),4);
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BCHANGEUNITE ] ) ;
end.
