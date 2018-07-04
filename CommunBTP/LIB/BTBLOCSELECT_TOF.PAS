{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/03/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTBLOCSELECT ()
Mots clefs ... : TOF;BTBLOCSELECT
*****************************************************************}
Unit BTBLOCSELECT_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
     uTob,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     grids,
     vierge,
     UTOF ;

Type
  TOF_BTBLOCSELECT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBTABLES : TOB;
    GS : Thgrid;
    ListeSaisie : string;
    procedure Constituegrille;
    procedure AfficheLagrille;

  end ;

Implementation

procedure TOF_BTBLOCSELECT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCSELECT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCSELECT.OnUpdate ;
var Indice : integer;
begin
  Inherited ;
  for Indice := 1 to GS.RowCount-1 do
  begin
  	if GS.IsSelected(indice) then
    begin
      TOBTABLES.detail[Indice-1].setString('SELECT','X');
    end;
  end;
  TOBTABLES.setString('OKOK','X');
end ;

procedure TOF_BTBLOCSELECT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCSELECT.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBTABLES := LaTOB;
  GS := THgrid(GetControl('GS'));
  ListeSaisie := 'OPEN;LIBELLE;';
  Constituegrille;
  AfficheLagrille;
end ;

procedure TOF_BTBLOCSELECT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCSELECT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCSELECT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTBLOCSELECT.Constituegrille;
begin
	GS.ColCount := 2;
  //
  GS.Cells[0,0] := ' ';
  GS.ColWidths [0] := 10;
  //
  GS.Cells[1,0] := 'Désignation';
  GS.ColWidths [1] := 120;
  //
end;

procedure TOF_BTBLOCSELECT.AfficheLagrille;
var indice : integer;
begin
  GS.RowCount := TOBTABLES.detail.count +1;
	FOR Indice := 0 to TOBTABLES.detail.count -1 do
  begin
    TOBTABLES.detail[Indice].PutLigneGrid(GS,Indice+1,false,false,ListeSaisie);
  end;
  TFvierge(ecran).HMTrad.ResizeGridColumns(GS); 
end;

Initialization
  registerclasses ( [ TOF_BTBLOCSELECT ] ) ;
end.
