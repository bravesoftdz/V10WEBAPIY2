{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 22/08/2003
Modifié le ... : 22/08/2003
Description .. : Source TOF de la FICHE : EDIT_TARIF ()
Mots clefs ... : TOF;EDIT_TARIF
*****************************************************************}
Unit EDIT_TARIF_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     MainEagl,
{$else}
     db,
     dbtables,
     mul,
     Fe_Main,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF, ParamSoc, EntGc;

function GCLanceFiche_EditionTarif (Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_EDIT_TARIF = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

function GCLanceFiche_EditionTarif (Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche (Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_EDIT_TARIF.OnArgument (S : String ) ;
var starg, st : string;
begin
  Inherited ;
  st := S;
  stArg := ReadTokenSt (st);

  if Ecran.Name = 'GCEDITTARIFTIE' then
  begin
    SetControlText ('XX_WHERE', 'AND GF_NATUREAUXI<>"FOU"');
  end else
  begin
    SetControlText ('XX_WHERE', 'AND (GF_NATUREAUXI="' + stArg + '" OR GF_NATUREAUXI="")');
  end;

  SetControlText ('NATUREAUXI', stArg);
  if stArg = 'CLI' then
  begin
    SetControlProperty ('GF_TIERS', 'DataType', 'GCTIERSCLI');
    SetControlText ('TGF_TIERS', 'Client');
  end else
  begin
    SetControlProperty ('GF_TIERS', 'DataType', 'GCTIERSFOURN');
    SetControlText ('TGF_TIERS', 'Fournisseur');
  end;

  SetControlProperty ('DEVISE', 'Value', GetParamSoc ('SO_DEVISEPRINC'));
  SetControlText ('GF_DEVISE', GetParamSoc ('SO_DEVISEPRINC'));
  SetControlText('GF_REGIMEPRIX','GLO');
  SetControlText ('GF_REGIMEPRIX', 'GLO');
if not VH_GC.GCMultiDepots then
   begin    //mcd 07/10/03 pour cacher dépot is pas gérer (GI/GA)
   SetControlVisible ('TGF_DEPOT', false);
   SetControlVisible ('GF_DEPOT', false);
   end;
end ;

procedure TOF_EDIT_TARIF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_TARIF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_EDIT_TARIF.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_EDIT_TARIF ] ) ; 
end.
