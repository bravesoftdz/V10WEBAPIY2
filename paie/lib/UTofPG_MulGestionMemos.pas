{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 12/02/2004
Modifié le ... :   /  /
Description .. : Unit gestion du multi critère gestion des memos de la paie
Mots clefs ... : PAIE
*****************************************************************}
{
}
unit UTofPG_MULGESTIONMEMOS;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Mul, Fe_Main,
  {$ELSE}
  MaineAgl, emul,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, UTOM,
  AGLInit;

type
  TOF_PGMULGESTIONMEMOS = class(TOF)
  private
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation


procedure TOF_PGMULGESTIONMEMOS.OnArgument(Arguments: string);
var
  Arg, Identif : string;
begin
  inherited;

  Arg := Readtokenst (Arguments);
  Identif := Readtokenst (Arguments);
  if Arg = '' then exit;

  SetControlText ('XX_WHERE', ' LO_TABLEBLOB = "'+Arg+'" AND LO_IDENTIFIANT ="'+Identif+'"' );

end;


procedure TOF_PGMULGESTIONMEMOS.OnLoad;
begin
  inherited;

end;

initialization
  registerclasses([TOF_PGMULGESTIONMEMOS]);
end.

