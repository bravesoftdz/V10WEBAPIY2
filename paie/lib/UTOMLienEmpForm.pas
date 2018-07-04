{***********UNITE*************************************************
Auteur  ...... : FL
Créé le ...... : 25/05/2007
Modifié le ... :   /  /
Description .. : Source TOM de la table LIENEMPFORM
Mots clefs ... : TOM;LIENEMPFORM
*****************************************************************
}
Unit UTOMLienEmpForm ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     eFiche,eFichList,   
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOM,UTob,PgOutils2,uTableFiltre,SaisieList,HMsgBox ;

Type
  TOM_LIENEMPFORM = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnUpdateRecord             ; override ;
    private
    TF : TTableFiltre;
    end ;

Implementation

procedure TOM_LIENEMPFORM.OnNewRecord ;
begin
  Inherited ;
       // Mise à blanc du libellé pour éviter les informations fantômes
       TF.LaGrid.CellValues[2,tf.LaGrid.Row] := '';
end ;


procedure TOM_LIENEMPFORM.OnArgument ( S: String ) ;
begin
  Inherited ;
       TF := TFSaisieList(Ecran).LeFiltre;
end ;

procedure TOM_LIENEMPFORM.OnUpdateRecord;
begin
       If (TF.LaGrid.CellValues[1,tf.LaGrid.Row] = '')  Then
       Begin
               PGIBox('Vous devez renseigner l''emploi associé au stage', 'Information obligatoire');
               LastError := 1;
               Exit;
       End;
  inherited;
end;

Initialization
  registerclasses ( [ TOM_LIENEMPFORM ] ) ;
end.

