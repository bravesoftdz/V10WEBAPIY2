{***********UNITE*************************************************
Auteur  ...... : gha
Créé le ...... : 16/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DETAILPOPULATION ()
Mots clefs ... : TOF;DETAILPOPULATION
*****************************************************************}
Unit DETAILPOPULATION_TOF ;

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
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_DETAILPOPULATION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    FListeDet : THGrid;
  end ;

Implementation

uses Grids;

procedure TOF_DETAILPOPULATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DETAILPOPULATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_DETAILPOPULATION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_DETAILPOPULATION.OnLoad ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : GHA
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. : Ce traitement permet d'afficher le détail des critères
Suite ........ : sélectionnés sur la fiche "Saisie des populations".
Mots clefs ... :
*****************************************************************}
procedure TOF_DETAILPOPULATION.OnArgument (S : String ) ;
var
  ReadList : TstringList;
  id_col,id_lig,num_rec : integer;
begin
  Inherited ;
  FListeDet := THGrid(GetControl('FLISTE'));
  if (FListeDet <> nil) and (s <> '') then
  begin
    ReadList := TstringList.Create;
    ReadList.Clear;
    ReadList.Text := s;

    num_rec := 0;
    FListeDet.RowCount := (ReadList.Count div FListeDet.ColCount)+1;

    for id_lig :=1 to (ReadList.Count div FListeDet.ColCount) do
    begin
      for id_col :=0 to FListeDet.ColCount-1 do
      begin
        FListeDet.Cells[id_col,id_lig] := ReadList.Strings[num_rec];
        inc(num_rec);
      end;
    end;
  end;
  ReadList.Free;
end ;

procedure TOF_DETAILPOPULATION.OnClose ;
begin
  Inherited ;
  FListeDet.Free;
end ;

procedure TOF_DETAILPOPULATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_DETAILPOPULATION.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_DETAILPOPULATION ] ) ;
end.
