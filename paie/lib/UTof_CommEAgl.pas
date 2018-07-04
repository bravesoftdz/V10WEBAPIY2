{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 18/06/2001
Modifié le ... :   /  /
Description .. : Unité de gestion des absences validées
Suite ........ : Uniquement pour le Manager de la base
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************}
unit Utof_CommEAgl;

interface
uses  StdCtrls,Classes, HTB97,
{$IFDEF EAGLCLIENT}
      UtileAGL,
{$ENDIF}
      UTOF,PgOutilsEAgl;

Type
     tof_CommEAgl = Class (TOF)
       private
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure BValiderClick(Sender: TObject);
     END ;
implementation

procedure tof_CommEAgl.OnArgument(Arguments: String);
var B : Ttoolbarbutton97;
begin
B := Ttoolbarbutton97(getcontrol('BVALIDER'));
if B <> nil then B.onclick := BValiderClick;
end;

// Administration des absences au niveau administrateur dans la base eAGL
procedure tof_CommEAgl.BValiderClick(Sender: TObject);
var RBTN : TRadioButton;
begin
RBTN := TRadioButton(GetControl ('RBTNEXPORT'));
if RBTN <> NIL then
   if RBTN.Checked = TRUE then RecupAbsFromeAgl(GetControlText('DATEINTEGRATION'));
end;


Initialization
registerclasses([tof_CommEAgl]);
end.
