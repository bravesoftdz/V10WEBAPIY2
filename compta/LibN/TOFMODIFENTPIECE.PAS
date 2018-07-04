{***********UNITE*************************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : MODIFENTETEPIECE ()
Mots clefs ... : TOF;MODIFENTETEPIECE
*****************************************************************}
Unit TOFMODIFENTPIECE ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF ;

Type
  TOF_MODIFENTPIECE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

procedure TOF_MODIFENTPIECE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIECE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIECE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIECE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MODIFENTPIECE.OnArgument (S : String ) ;
begin
  Inherited ;

  THValComboBox(GetControl('J_NATUREJAL')).ItemIndex := 0 ;

end ;

procedure TOF_MODIFENTPIECE.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_MODIFENTPIECE ] ) ;
end.
