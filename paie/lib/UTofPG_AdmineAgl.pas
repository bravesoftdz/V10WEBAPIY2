{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 14/06/2001
Modifié le ... :   /  /
Description .. : Unit administration de la base Saisie déportée
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************
PT1 28/10/2002 V585 On ne doit pas tenir compte du champ pcn_validsalarie pour l'export des données
PT2 05/12/2002 V591 SB Etat des mouvements exportés modifiés
}
unit UtofPG_AdmineAgl;

interface
uses  StdCtrls,Controls,Classes,HTB97,HCtrls,HMsgBox,UTOF,ParamSoc;



Type
     tof_PG_AdmineAgl = Class (TOF)
       private
       procedure BValideClick(Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
     END ;
implementation


procedure tof_PG_AdmineAgl.OnArgument(Arguments: String);
var B : Ttoolbarbutton97;
begin
B := Ttoolbarbutton97(getcontrol('Bvalider'));
if B <> nil then
   B.onclick := BValideClick;
end;

procedure tof_PG_AdmineAgl.BValideClick(Sender: TObject);
var RBTN : TRadioButton;
    rep : Integer;
begin
RBTN := TRadioButton(GetControl ('RBTNABS'));
if RBTN <> NIL then
   begin // On revalide pour l'export vers la base de production les absences exportées
   if RBTN.Checked = TRUE then
    begin
    Rep := PGIAsk ('Cette procédure doit être utilisée uniquement pour réactiver #13#10 '+
               'Les absences exportées dont la récupération n''a pas été faite', Ecran.Caption);
    if rep = MrOk then       //PT2 Les mouvements exportés sont topés ENC et non plus "X", recap faussé
     ExecuteSQL ('UPDATE ABSENCESALARIE SET PCN_EXPORTOK = "-" WHERE PCN_SAISIEDEPORTEE = "X" '+
     'AND PCN_VALIDRESP = "VAL" AND PCN_EXPORTOK = "ENC"'); {PT1 mise en commentaire PCN_VALIDSALARIE = "VAL" AND}
    end;
   end;
RBTN := TRadioButton(GetControl ('RBTNBLOQ'));
if RBTN <> NIL then
   begin // On bloque l'accès à la base car on est en train de faire un export
   if RBTN.Checked = TRUE then SetParamSoc('SO_PGBASEEPAIE','-');
   end;
RBTN := TRadioButton(GetControl ('RBTNBLOQ'));
if RBTN <> NIL then
   begin // On débloque l'accès à la base
   if RBTN.Checked = TRUE then SetParamSoc('SO_PGBASEEPAIE','X');
   end;
end;

Initialization
registerclasses([tof_PG_AdmineAgl]);
end.
