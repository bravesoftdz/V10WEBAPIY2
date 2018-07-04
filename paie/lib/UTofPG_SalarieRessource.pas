{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/08/2003
Modifié le ... :   /  /
Description .. : Unit de saisie du code ressource en creation de salarié
Suite ........ :
Mots clefs ... : PAIE
*****************************************************************}
unit UTofPG_SalarieRessource;

interface
uses  Controls,
      Classes,
      HTB97,
{$IFNDEF EAGLCLIENT}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      HDB,
      DBCtrls,
      Fe_Main,
      UTOB,
      AGLInit,
      HEnt1,
      ComCtrls,
      sysutils,
      Graphics,
      forms,
      StdCtrls,
{$ELSE}

{$ENDIF}
      HCtrls,
      HMsgBox,
      UTOF,
      Vierge;
Type
     TOF_PG_SalarieRessource = Class (TOF)
       private
       BV     : TToolbarButton97;
       procedure ControlRessource (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnClose                  ; override ;
     END ;

implementation

procedure TOF_PG_SalarieRessource.ControlRessource (Sender: TObject);
begin
if GetControlText ('CODERESSOURCE') = '' then
begin
  PgiBox ('Vous devez saisir un code ressource', Ecran.Caption );
  SetFocusControl ( 'CODERESSOURCE') ;
end
else
begin
  if ExisteSQL ('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetControlText ('CODERESSOURCE')+'"') then
  begin
    PgiBox ('La ressource '+ GetControlText ('CODERESSOURCE') + ' existe déjà ? ', Ecran.Caption );
    SetFocusControl ( 'CODERESSOURCE') ;
  end
  else if BV <> NIL then BV.Click ;
end ;

end ;

procedure TOF_PG_SalarieRessource.OnArgument (Arguments: String) ;
var Res    : THEdit ;
    Btn    : TToolbarButton97;
begin
  inherited ;
  Res := THEdit (GetControl ('CODERESSOURCE')) ;
  if Res <> NIL then Res.OnExit := ControlRessource ;
  Btn :=  TToolbarButton97 ( GetControl ('VALIDATION')) ;
  if Btn <> NIL then Btn.OnClick := ControlRessource ;
  BV :=  TToolbarButton97 ( GetControl ('BVALIDER')) ;

  SetFocusControl ( 'CODERESSOURCE') ;
end ;

procedure TOF_PG_SalarieRessource.OnClose;
begin
  inherited ;
// on recupère le code ressource saisi
 TFVierge(Ecran).Retour := GetControlText ('CODERESSOURCE') ;
end ;

Initialization
registerclasses ([TOF_PG_SalarieRessource]) ;
end.
