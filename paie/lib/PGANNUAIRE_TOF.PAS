{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 24/08/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF Muticritere specifique gestion de infos de
Suite ........ : l'annuaire concernant uniquement le social Accessible
Suite ........ : depuis la paie sans prendre les autres critères contenus
Suite ........ : dans la table ANNUAIRE du DP
Mots clefs ... : PAIE;ANNUAIRE
*****************************************************************
---- JL 20/03/2006 modification clé annuaire ----
PT1 : 28/03/2006 MF V_65 : correction modification clé annuaire
PT2 : 17/10/2006 JL V_70 : FQ 13044 Rafraichissement tablette annuaire pour bénéficiaire saisie arrêt
PT3 : 20/11/2006 JL V_70 : FQ 13693 Permettre multi sélection en CWAS pour récup organisme.
}

Unit PGANNUAIRE_TOF ;

Interface

Uses StdCtrls, Controls, Classes,forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} mul,
{$ELSE}
     eMul, UtileAGL,
{$ENDIF}
     HCtrls, HEnt1,HStatus, HMsgBox, UTOF,HTB97 ;

Type
  TOF_PGANNUAIRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    Trait,TypePersonne : String;                   // Type de traitement avec code Retour
    procedure RecupValeurCombo(Sender : TObject);
  end ;

Implementation

procedure TOF_PGANNUAIRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGANNUAIRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGANNUAIRE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGANNUAIRE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGANNUAIRE.OnArgument (S : String ) ;
var
     Bt :TToolBarButton97;
begin
  Inherited ;
TFmul(Ecran).retour := 'VIDE';
SetControlText('ANN_FAMPER','SOC');
Trait := ReadTokenPipe (S,';');
typePersonne := ReadTokenPipe (S,';');
If TypePersonne <> 'RET' then SetControltext('ANN_TYPEPER',TypePersonne);
if Trait = 'R' then
   begin
   {$IFDEF EAGLCLIENT}
   SetControlProperty('Fliste', 'MultiSelect', true);   //PT3
   {$ELSE}
   SetControlProperty('Fliste', 'Multiselection', true);
   {$ENDIF}
   SetControlVisible ('BTNVALIDANN', TRUE);
   SetControlVisible ('BAnnuler', FALSE);
   end
   else
   begin
   If Trait = 'COMBO' then
   begin
        SetControlVisible ('BTNVALIDANN', FALSE);
        SetControlVisible ('BOuvrir', False);
        SetControlVisible ('BRECUP', True);
        Bt := TToolBarButton97(GetControl('BRECUP'));
        If Bt <> Nil then Bt.OnClick := RecupValeurCombo;
        TFmul(Ecran).FListe.onDblClick := RecupValeurCombo;
   end
   else
   begin
        SetControlVisible ('BTNVALIDANN', FALSE);
        SetControlVisible ('BAnnuler', TRUE);
   end;
   end;
end ;

procedure TOF_PGANNUAIRE.OnClose ;
var
  i : Integer;
  ordre       : string;
begin
  Inherited ;
if Trait <> 'R' then exit;
TFmul(Ecran).retour := 'VIDE';  // initialisation avec une valeur significative
if TFmul(Ecran).FListe.NbSelected-1 > 1 then
 begin
 PgiBox ('Vous ne devez sélectionner qu''une seule ligne', Ecran.Caption);
 end
 else
 BEGIN
 InitMove(TFmul(Ecran).FListe.NbSelected,'');
 for i:=0 to TFmul(Ecran).FListe.NbSelected-1 do
    BEGIN
    TFmul(Ecran).FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
    TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row-1) ;
{$ELSE}
{$ENDIF}
//d PT1
{   Ordre  :=TFmul(Ecran).Q.findfield('ANN_GUIDPER').asinteger;
   TFmul(Ecran).retour := IntToStr (Ordre);}
//f PT1
   Ordre  :=TFmul(Ecran).Q.findfield('ANN_GUIDPER').asstring;
   TFmul(Ecran).retour := Ordre;
   END;
 END ;
end ;

procedure TOF_PGANNUAIRE.RecupValeurCombo(Sender : TObject);
begin
     If TypePersonne = 'DDT' then AvertirTable('PGDDTEFP');
     If TypePersonne = 'MED' then AvertirTable('PGMEDTRAVAIL');
     If TypePersonne = 'RET' then AvertirTable('ANNUAIRE');       //PT2
     TFmul(Ecran).retour := TFmul(Ecran).Q.findfield('ANN_GUIDPER').asstring;
     TFMul(Ecran).BAnnulerCLick(TFMul(Ecran).BAnnuler);
end;

Initialization
  registerclasses ( [ TOF_PGANNUAIRE ] ) ;
end.
