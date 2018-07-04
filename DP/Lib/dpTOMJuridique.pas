{***********UNITE*************************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. : TOM de la table JURIDIQUE
Mots clefs ... : fiche principale : DP FICHJURIDIQUE -> JUR JURIDIQUE
*****************************************************************}
unit dpTOMJuridique;

interface

uses
   SysUtils, Controls, //LM20070516
   {$IFNDEF EAGLCLIENT}
   Fe_Main, Fiche,
   {$ELSE}
   MaineAgl, eFiche,
   {$ENDIF}
   UTOMYYJuridique, db, HTB97, hctrls, StdCtrls, dpOutils, ComCtrls,
   Classes,
   utob,  //LM20070516
   forms;  //LM20070516

  Type
     TOM_JURIDIQUE = Class (TOM_YYJuridique)
         procedure OnArgument (stArgument : String ) ; override ;
         procedure OnClose; override;
         procedure OnNewRecord; override;
         procedure OnLoadRecord ; override ;
         procedure OnChangeField(F: TField);         
         procedure OnUpdateRecord ; override ;
         procedure OnAfterUpdateRecord; override;
        private
     end;



////////////// IMPLEMENTATION //////////////
implementation
uses DpJurOutils, uSatUtil, HEnt1, entDP, Annoutils ;//LM20060611 LM20070516
{ TOM_JURIDIQUE }


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Juridique.OnArgument(stArgument : String);
var
    p : integer;
    s : String;
    bVoirLiens : Boolean;
begin
   inherited;
   //+LM20070406
   //Onglet proposé par défaut.
   p := pos('TABSHEET=',uppercase(stArgument)) ;
   if p>0 then
   begin
    s:=copy(stArgument, p+length('TabSheet='), 100) ;
    s:=gtfs (s,';',1);
    ActiveTabSheet :=s ;
   end ;
   //-LM20070406

   bVoirLiens := JaiLeDroitConceptBureau(ccVoirLesLiens); // suite FQ 11763 ! and JaiLeDroitConceptBureau(ccModifAnnuaireOuLiens);
   SetControlVisible('BLIENS', bVoirLiens);
   SetControlEnabled('NOMPROPRIETAIRE', bVoirLiens);
   SetControlVisible('BLIENPROPRIETAIRE1', bVoirLiens);
   // SetControlEnabled('TANN_PERASS2GUID', bVoirLiens);
   SetControlEnabled('TANN_PERASS2GUID', False); // MD c'est non modifiable, il faut passer par l'annuaire
   SetControlProperty('TANN_PERASS2GUID', 'ElipsisButton', False);
   // SetControlVisible('BTSUPPERASS2GUID', bVoirLiens and JaiLeDroitConceptBureau(ccSupprAnnuaireOuLiens));
   SetControlVisible('BTSUPPERASS2GUID', False); // MD bouton inutile car non branché !

   SetControlVisible('INFODOSSIER_JUR', false);

   SetControlVisible('BINSERT', false);
   SetControlVisible('BDELETE', false);
//   SetControlVisible('BIMPRIMER', false);
   SetControlVisible('BTYPEINFOS', false);
   SetControlVisible('BEXERCICE', false);
   SetControlVisible('BCFE', false);
   SetControlVisible('BEVENEMENTS', false);
   SetControlVisible('BOPERATIONS', false);
   SetControlVisible('BCTRDIVERS', false);

   SetControlVisible('TJUR_IMPODIR', false);
   SetControlVisible('JUR_IMPODIR', false);
   SetControlVisible('TJUR_ISBAIL', false);
   SetControlVisible('JUR_ISBAIL', false);
   SetControlVisible('TJUR_ARCHIVE', false);
   SetControlVisible('JUR_ARCHIVE', false);
   SetControlVisible('BVOIRPERSONNEJUR', false);
   SetControlVisible('BTITRES', false);

   if sValeurFiche_c = 'SYNTHESE' then
   begin
      TTabSheet(GetControl('PGeneral')).TabVisible :=  FALSE;
      TTabSheet(GetControl('PACTIVITE')).TabVisible :=  FALSE;
      TTabSheet(GetControl('TSGERANCE')).TabVisible :=  FALSE;
      TTabSheet(GetControl('PINFOSTAT')).TabVisible :=  FALSE;
      TTabSheet(GetControl('TSACTIONNAIRES')).TabVisible :=  FALSE;
      TTabSheet(GetControl('PBLOCNOTE')).TabVisible :=  FALSE;
      TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl('TSCAPITAL'));
   end;

   setControlVisible('PSECRETARIAT', vh_dp.Group);//LM20070611
   TFFiche(ecran).DisabledMajCaption := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Juridique.OnClose;
begin
   inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Juridique.OnNewRecord;
begin
   inherited;
   if GetField('JUR_CODEDOS') = '' then
      SetField('JUR_CODEDOS', '&#@');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Juridique.OnLoadRecord;
var
    ctrl : TControl;
begin
  inherited;
{   SetControlText('TJUR_REFGERANTDP', RecupNomAnnulien(sGuidPerDos_c,'GRT') );
   GereCac;
   //GereEleveNomin;
   GereNbAdmin;}

   if ActiveTabSheet <> '' then
      SetActiveTabSheet(ActiveTabSheet); //LM20070412

   ctrl := getControl('Pages');
   if (ctrl <> nil) and (ThPageControl2(ctrl).activepage = nil) then
   begin
      if getControl('PGeneral') <> nil then
         ThPageControl2(ctrl).activepage := TTabSheet(getControl('PGeneral')) ; //LM20070611
   end;
   //MD20070621 - Verrue, car pas moyen de planquer PSECRETARIAT malgré le code LMO
   ctrl := GetControl('PAGES');
   if (ctrl <> nil) and (Not vh_dp.Group) and (ThPageControl2(ctrl).activepage <> nil) then
   begin
      if (ThPageControl2(ctrl).activepage.Name='PSECRETARIAT') then
      begin
         ThPageControl2(ctrl).activepage := TTabSheet(getControl('PGeneral')) ;
         TTabSheet(GetControl('PSECRETARIAT')).TabVisible := False;
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JURIDIQUE.OnChangeField(F: TField);
begin
   inherited;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Juridique.OnUpdateRecord;
begin
   inherited;
   if LastError = 1 then exit;
   // Ici : traitements éventuels spécifiques au mode DP
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_Juridique.OnAfterUpdateRecord;
begin
   inherited;
end;


Initialization
  registerclasses([TOM_Juridique]) ;
end.
