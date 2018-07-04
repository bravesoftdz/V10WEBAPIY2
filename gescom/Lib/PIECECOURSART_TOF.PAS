{***********UNITE*************************************************
Auteur  ...... : Joël TRIFILIEFF                                             
Créé le ...... : 29/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PIECECOURSART ()
Suite..........: Limite l'affichage des natures de pièces dans liste des
Suite..........: pièces d'un article
Mots clefs ... : TOF;PIECECOURSART
*****************************************************************}
Unit PIECECOURSART_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF NOMADE}
     UtilPOP,         
{$ENDIF}
     forms,menus, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF ;

Type
  TOF_PIECECOURSART = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;
Implementation

procedure TOF_PIECECOURSART.OnArgument (S : String ) ;
{$IFDEF NOMADE} //Limite nature des documents pour la sélection
var
    NatureDoc : THMultiValComboBox;
{$ENDIF} //NOMADE
{$IFDEF BTP}
var
    i : integer;
    Mpop : TPopupMenu ;
{$ENDIF} //BTP
begin
  Inherited ;
{$IFDEF NOMADE} //Limite nature des documents pour la sélection
NatureDoc := THMultiValComboBox(GetControl('GL_NATUREPIECEG'));
NatureDoc.Plus := GetNaturePOP('GPP_NATUREPIECEG');
{$ENDIF} //NOMADE

{$IFDEF BTP}
Mpop:=tpopupmenu(GetControl('POPUPMENUZ'));
for i:=0 to Mpop.Items.Count-1 do
  BEGIN
  if uppercase(Mpop.Items[i].Name)='MNZTIERS' then Begin Mpop.Items[i].Visible:=False;  end;
  END;
{$ENDIF} //BTP

end ;


Initialization
  registerclasses ( [ TOF_PIECECOURSART ] ) ;
end.
