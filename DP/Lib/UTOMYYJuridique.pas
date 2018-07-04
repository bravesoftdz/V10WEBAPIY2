{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 11/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
unit UTOMYYJuridique;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   {$IFNDEF EAGLCLIENT}
   Fe_Main, Fiche,
   {$ELSE}
   MaineAgl, eFiche,
   {$ENDIF}
   sysutils, ComCtrls, Classes, Controls, HMsgBox, HEnt1, db, hctrls,
   HTB97, windows, messages, stdctrls, extctrls, HRichOle,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}
   {$IFDEF DP}
   EntDP,
   {$ENDIF}

   DpJurOutilsBlob, DpJurOutils, UTOB, CLASSCapital, CLASSGrilleSansMul,
   dpOutils, UOutilsAnnuLienTri, UTOMZonesLibres, UTOFAnnuLienTri, USupprime,
   LittleTom,    //LM20070516
   EventDecla,   //LM20070516
   Forms;        //LM20070704

/////////////////////////////////////////////////////////////////
Type
     TOM_YYJURIDIQUE = Class (TOM_ZONESLIBRES)

         procedure OnArgument (sArgument_p : String ) ; override ;
         procedure OnClose; override;
         procedure OnNewRecord; override;
         procedure OnLoadRecord ; override ;
         procedure OnChangeField (F : TField); override ;
         procedure OnUpdateRecord ; override ;
         procedure OnAfterUpdateRecord; override;

      public
         procedure OnFormKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState); virtual;
         procedure OnChange_BLOB(Sender : TObject); virtual;
         procedure OnChange_RCSDATE(Sender : TObject);
         procedure OnDblClick_FLISTE(Sender : TObject); virtual;
         procedure OnClick_BVOIRPERSONNE(Sender : TObject); virtual;
         procedure OnClick_BTITRES(Sender : TObject); virtual;
         procedure OnClick_BLIENS(Sender: TObject);
         procedure OnClick_BTORGDIRECTION(Sender : TObject);
         procedure OnClick_BTGERANT(Sender : TObject);
         procedure OnClick_BTCAC(Sender : TObject);
         procedure OnClick_CBCAC(Sender : TObject);
         procedure OnClick_BIMPRIMER(Sender : TObject) ;
         procedure OnClick_BEVENEMENTS(Sender: TObject);
         procedure OnClick_BTYPEINFOS(Sender: TObject);
         procedure OnClick_BDelete(Sender: TObject);
         procedure OnElipsisClick_Gerant(Sender: TObject);
         procedure OnElipsisClick_CC(Sender: TObject);
         procedure OnClick_FLISTEFONCTION(Sender : TObject);

      protected
         sTypeFiche_c, sValeurFiche_c, sPPPM_c, sAnnCodeInsee_c : string;
         sTypeDos_c, sGuidPerDos_c, sNoOrdre_c, sAnnForme_c : string;
         bChangeForme_c, bCharge_c : boolean;
         dtDateDeb_c, dtDateFin_c : TDateTime;
         bDossierDP_c, bSuppCAC_c : boolean;
         sCAC_c : string;
         ActiveTabSheet : string ;//LM20070406

         function  DateExpCalcul( dtDateExp_p, dtRcsDate_p : TDateTime; nDuree_p : integer) : TDateTime;
         procedure CapitalCalcule(dDateDeb_p, dDateFin_p : TDateTime; sDomaine_p : string);
         function  CapitalRefresh : boolean;
         procedure CapitalMAJ;
         procedure GereActionPreference;
         procedure GereGerance;
         procedure GereOrganeDir;
         procedure GereEleveNomin;
//         procedure GereNbAdmin;
         procedure GereCac;
//         function  ExisteCAC : string;

         procedure JURDateExercice(var dDateDeb_p, dDateFin_p : TDateTime);
         //LM20070412 plus utilisé visiblement procedure DPDateExercice(var dDateDeb_p, dDateFin_p : TDateTime);
         procedure ActionnairesAffiche;
//         procedure VraiOnClose (Sender:TObject; var Action:TCloseAction);
      private
         OBBlob_c : TOB;
         OBFonctions_c, OBCurFonction_c, OBCurLien_c : TOB;
         sFonctDefaut_c : string;
         gdActionnaires_c : TGrilleSansMul;
         cCapital_c : TCapital;

         evt : TEventDecla ;                         //LM20070516
         ANN_ : TLittleTOM ;                         //LM20070516
         DFI_ : TLittleTOM ;                         //LM20070516
         DOR_ : TLittleTOM ;                         //MD20070622

         codeNafIni : string;                        //LM20070516

         procedure DossierChargeEtAffiche(sGuidPerDos_p : string; bInsert_p : boolean );
         function  InfosTransmisesPrepare : string;
         procedure InfosTransmisesRecupere( sInfosDossier_p : string);
         procedure BlobsCharge(sGuidPerDos_p : string );
         procedure BlobsReCharge(sGuidPerDos_p, sTypeDos_p : string );
         procedure Change_BLOB(sControlName_p : string);
         procedure VoirPersonne;
         procedure LiensChargeEtAffiche;
         procedure ForceRegimeFiscal;         
         procedure AfficheMasqueSelonFormeJur;

         // procedure ActionnaireFormateLibelles(var sFonction_p, sFonctAbrege_p : string);
         // function  ActionnaireAutorise : boolean;
         procedure AfficheLien(sZone_p, sFonction_p : String);
         function InFonctionsTitres : string;
         // procedure GestionActionPreference;
         procedure AfficheOngletGerance ;//LM20060516

         procedure initAutreTom  ;                   //LM20070516
         procedure Ann_OnExit (sender : TObject) ;   //LM20070516
         procedure ANN_RCSOnClick(sender:TObject) ;  //LM20070516
         procedure ANN_RMOnClick(sender:TObject) ;   //LM20070516
         procedure rcs;
         procedure rm ;

         procedure VoirSCI ;                         //LM20070516
         procedure bFormeOnClick (sender:TObject) ;   //LM20070516
         procedure ANN_CODENAFOnEnter (sender:TObject) ;   //LM20070516
         procedure ElipsisPerAssClick(Sender: TObject);  //LM20070516
         procedure JUR_LOCAGERANCE_OnClick(Sender: TObject); //de TomOrga ... LM20070516
         procedure GereLocagerance; //de TomOrga ... LM20070516
         procedure BLIENPROPRIETAIRE1_OnClick(Sender: TObject); //de TomOrga ... LM20070516
         procedure ElipsisPropriClick(Sender: TObject); //de TomOrga ... LM20070516

   end;
/////////////////////////////////////////////////////////////////
implementation

uses uSatUtil, Annoutils , UTOM;// LM20070516

/////////////////////////////////////////////////////

{ TOM_YYJURIDIQUE }

/////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnArgument(sArgument_p: String);
var
   sAction_l : string;
begin
   inherited;
   sAction_l := ReadTokenSt(sArgument_p);
   sGuidPerDos_c := ReadTokenSt(sArgument_p);
   sTypeDos_c := ReadTokenSt(sArgument_p);
   sNoOrdre_c := ReadTokenSt(sArgument_p);
   sTypeFiche_c := ReadTokenSt(sArgument_p);
   sValeurFiche_c := ReadTokenSt(sArgument_p);

   bCharge_c := true;

   THRichEditOle(GetControl('OLE_OBJETSOC')).onChange := OnChange_BLOB;
   THRichEditOle(GetControl('OLE_BLOCNOTES')).onChange := OnChange_BLOB;

   THEdit(GetControl('ANN_RCSDATE')).OnChange := OnChange_RCSDATE;

   TToolbarButton97(GetControl('BLIENS')).OnClick := OnClick_BLIENS;

   if GetControl('BVOIRPERSONNE')<>nil then //LM20070315
      TToolbarButton97(GetControl('BVOIRPERSONNE')).Onclick := OnClick_BVOIRPERSONNE;

   if GetControl('BEVENEMENTS')<> nil then
      TToolbarButton97(GetControl('BEVENEMENTS')).Onclick := OnClick_BEVENEMENTS;

   if GetControl('BTYPEINFOS')<>nil then
      TToolbarButton97(GetControl('BTYPEINFOS')).Onclick := OnClick_BTYPEINFOS;

   if GetControl('BDELETE')<>nil then
      TToolbarButton97(GetControl('BDELETE')).Onclick := OnClick_BDelete;

   TToolbarButton97(GetControl('BIMPRIMER')).OnClick := OnClick_BIMPRIMER;
   THGrid(GetControl('LISTEFONCTION')).OnClick := OnClick_FLISTEFONCTION;

   Ecran.OnKeyDown := OnFormKeyDown;

   SetControlProperty('ANN_SIREN', 'Enabled', False);
   SetControlProperty('ANN_CLESIRET', 'Enabled', False);
   TCheckBox(GetControl('JUR_CAC')).OnClick := OnClick_CBCAC;
   SetControlVisible ('BFORME', False);

   cCapital_c := TCapital.Create;
   OBBlob_c   := TOB.Create('La table des blobs', nil,-1);
   OBFonctions_c := TOB.Create ('les fonctions jur', nil, -1); // $$$ JP 21/04/06 - TOB.Create('ANNUAIRE', nil, -1);

   gdActionnaires_c := TGrilleSansMul.Create(THGrid(GetControl('FLISTE')),
                                             GetControlText('DBLISTE'),
                                             TPanel(GetControl('PCUMUL')));

   //+LM20070516
   ANN_ := TLittleTOM.Create(Ecran, 'ANNUAIRE', self, @ds);
   DFI_ := TLittleTOM.Create(Ecran, 'DPFISCAL', self, @ds);
   DOR_ := TLittleTOM.Create(Ecran, 'DPORGA',   self, @ds);

   evt:=TEventDecla.create(Ecran);
   initAutreTom ;

   Evt.Rebranche('ANN_RCS', 'OnClick', ANN_RCSOnClick);
   Evt.Rebranche('ANN_RM', 'OnClick', ANN_RMOnClick);
   Evt.Rebranche('ANN_CODENAF', 'OnEnter', ANN_CODENAFOnEnter);
   Evt.Rebranche('ANN_CODENAF2', 'OnEnter', ANN_CODENAFOnEnter);
   Evt.Rebranche('TANN_PERASS2GUID', 'OnElipsisClick', ElipsisPerAssClick);


   Evt.Rebranche('bForme', 'OnClick', bFormeOnClick);
   Evt.Rebranche('JUR_LOCAGERANCE', 'OnClick', JUR_LOCAGERANCE_OnClick);
   Evt.Rebranche('NOMPROPRIETAIRE', 'OnElipsisClick', ElipsisPropriClick);
   Evt.Rebranche('BLIENPROPRIETAIRE1', 'OnClick', BLIENPROPRIETAIRE1_OnClick);
//   Evt.Rebranche('Juridique', 'OnClose', VraiOnClose); //LM20070704
   //-LM20070516


   // MD 09/07/07 - FQ 11566 : on passera par identité pour changer la forme

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnNewRecord;
begin
   inherited;
   SetField('JUR_GUIDPERDOS', sGuidPerDos_c);
   SetField('JUR_TYPEDOS', sTypeDos_c);
   SetField('JUR_NOORDRE', sNoOrdre_c);

   if sPPPM_c = 'PM' then
      SetField('JUR_ISBAIL', true);

   ANN_.NewEnreg ;//LM20070516
   DFI_.NewEnreg ;//LM20070516
   DOR_.NewEnreg ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/08/2007
Modifié le ... :   /  /    
Description .. : Valeur par défaut issue du fiscal
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.ForceRegimeFiscal;
begin
   // MD 20080114 - FQ 11929 - Ne pas forcer IR !
   {if (GetField('JUR_FORME') = 'SA') or (GetField('JUR_FORME') = 'SAS') then
   begin
      if (DFI_.Get('DFI_IMPODIR') <> 'IR') or (GetField('JUR_IMPODIR') <> 'IR') then
      begin
         ModeEdition(DS);
         SetField('JUR_IMPODIR', 'IR');
      end;
      SetControlEnabled('JUR_IMPODIR', false);
   end
   else} if (DFI_.Get('DFI_IMPODIR') <> GetField('JUR_IMPODIR')) then
   begin
      ModeEdition(DS);
      SetField('JUR_IMPODIR', DFI_.Get('DFI_IMPODIR'));
      SetControlEnabled('JUR_IMPODIR', true);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.AfficheMasqueSelonFormeJur;
var
   bVisible1_l, bVisible2_l : boolean;
begin
   SetControlVisible('JUR_APPELPUB', (GetField('JUR_FORME') = 'SA'));
   SetControlVisible('TJUR_NREDG', (GetField('JUR_FORME') = 'SA') or (GetField('JUR_FORME') = 'SAS'));
   SetControlVisible('JUR_NREDG', (GetField('JUR_FORME') = 'SA') or (GetField('JUR_FORME') = 'SAS'));


   // Onglet Particularités Statutaires
   if (GetField('JUR_FORME') = 'SA') or (GetField('JUR_FORME') = 'SAS') then
      SetControlCaption('JUR_POUVTRFSS', 'Pouvoir au conseil pour transfert siège social')
   else
      SetControlCaption('JUR_POUVTRFSS', 'Pouvoir à la gérance pour transfert siège social');

   // Onglet Infos statutaires
   // Nombre minimum de titres
   bVisible1_l := (GetField('JUR_FORME') <> 'SA') and (GetField('JUR_FORME') <> 'SADIR') and
                  (GetField('JUR_FORME') <> 'SAS') and (GetField('JUR_FORME') <> 'SASU');
   bVisible2_l := (GetField('JUR_FORME') = 'SA') or (GetField('JUR_FORME') = 'SADIR') or
                  (GetField('JUR_FORME') = 'SAS') or (GetField('JUR_FORME') = 'SASU');

   SetControlVisible('JUR_TTNBMINASS', bVisible1_l);
   SetControlVisible('TJUR_TTNBMINASS', bVisible1_l);

   SetControlVisible('JUR_TTNBMINADM', bVisible2_l);
   SetControlVisible('TJUR_TTNBMINADM', bVisible2_l);

   SetControlVisible('JUR_NBGRTMIN', bVisible1_l);
   SetControlVisible('JUR_NBGRTMAX', bVisible1_l);
   SetControlVisible('TJUR_NBGRTMIN', bVisible1_l);

   SetControlVisible('JUR_NBADMMIN', bVisible2_l);
   SetControlVisible('JUR_NBADMMAX', bVisible2_l);
   SetControlVisible('TJUR_NBADMMIN', bVisible2_l);

   SetControlVisible('JUR_NBDGMIN', bVisible2_l);
   SetControlVisible('JUR_NBDGMAX', bVisible2_l);
   SetControlVisible('TJUR_NBDGMIN', bVisible2_l);

   // Durée des mandats
{   bVisible1_l := (GetField('JUR_FORME') <> 'SA') and (GetField('JUR_FORME') <> 'SADIR') and
                  (GetField('JUR_FORME') <> 'SAS') and (GetField('JUR_FORME') <> 'SASU');
   bVisible2_l := (GetField('JUR_FORME') = 'SA') or (GetField('JUR_FORME') = 'SADIR') or
                  (GetField('JUR_FORME') = 'SAS') or (GetField('JUR_FORME') = 'SASU');}

   SetControlVisible('JUR_DURMANDGRT', bVisible1_l);
   SetControlVisible('TJUR_DURMANDGRT', bVisible1_l);

   SetControlVisible('JUR_DURMANDADM', bVisible2_l);
   SetControlVisible('TJUR_DURMANDADM', bVisible2_l);

   SetControlVisible('JUR_DURMANDPCA', bVisible2_l);
   SetControlVisible('TJUR_DURMANDPCA', bVisible2_l);

   SetControlVisible('JUR_DURMANDDG', bVisible2_l);
   SetControlVisible('TJUR_DURMANDDG', bVisible2_l);

   // Limite d'âge

   SetControlVisible('JUR_LIMAGEGRT', bVisible1_l);
   SetControlVisible('TJUR_LIMAGEGRT', bVisible1_l);

   SetControlVisible('JUR_LIMAGEADM', bVisible2_l);
   SetControlVisible('TJUR_LIMAGEADM', bVisible2_l);
   SetControlVisible('JUR_LIMAGEADMPC', bVisible2_l);
   SetControlVisible('TJUR_LIMAGEADMPC', bVisible2_l);

   SetControlVisible('JUR_LIMAGEPCA', bVisible2_l);
   SetControlVisible('TJUR_LIMAGEPCA', bVisible2_l);

   SetControlVisible('JUR_LIMAGEDG', bVisible2_l);
   SetControlVisible('TJUR_LIMAGEDG', bVisible2_l);

   // Onglet titres
   SetControlVisible('PBOTTOM', (GetField('JUR_FORME') = 'SA'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnLoadRecord;
var
   ctrl:TControl ;
   sNoDossier_l : string;
begin
   inherited;
   bDossierDP_c := (GetField('JUR_CODEDOS') = '&#@');

   sNoDossier_l := GetValChamp('DOSSIER', 'DOS_NODOSSIER', 'DOS_GUIDPER = "' + sGuidPerDos_c + '"');
   SetControlText('DOS_NODOSSIER', sNoDossier_l);
   SetControlVisible('DOS_NODOSSIER', (sNoDossier_l <> ''));
   SetControlVisible('TDOS_NODOSSIER', (sNoDossier_l <> ''));

   DossierChargeEtAffiche(GetField('JUR_GUIDPERDOS'), (Ds.State = dsInsert));
   BlobsCharge(GetField('JUR_GUIDPERDOS'));

   ANN_RCSOnClick(nil); //LM20070516
   ANN_RMOnClick(nil);  //LM20070516
//   ANN_OnExit(getControl('ANN_CODENAF'));//LM20070516 En attendant... Il faut revoir TOM_Annuaire.EnregBlob... pour un meilleur partage entre annuaire et juridique
//   ANN_OnExit(getControl('ANN_CODENAF2'));
   ANN_OnExit(getControl('ANN_SIREN'));

//   ANN_.ChargeEnreg(GetField('JUR_GUIDPERDOS')); //Dans DossierChargeEtaffiche
   DFI_.ChargeEnreg(GetField('JUR_GUIDPERDOS')); //LM20070516
   DOR_.ChargeEnreg(GetField('JUR_GUIDPERDOS')); //MD20070622

   if GetField('JUR_NOTEOUV') = 'X' then
   begin
      SetActiveTabSheet( 'PBLOCNOTE' );
      SetControlVisible( 'PBLOCNOTE', true );
   end;

   SetControlEnabled('BLIENS', (DS.State <> dsInsert));


  AfficheInfoLienPersonne(self, GetControlText('ANN_PERASS2GUID'), 'ANN_PERASS2GUID', 'TANN_PERASS2GUID'); //LM20070516

  voirSci ;//LM20070516
  GereLocagerance;//LM20070516

  if VerrouilleSiBaseComptable(self, GetField('JUR_GUIDPERDOS')) then //LM20070620
    setControlEnabled('JUR_DATEDEBUTEX', false);

   bCharge_c := false;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnChangeField(F: TField);
var
   dtDateExp_l, dtRcsDate_l : TDateTime;
   nDuree_l : integer;
   bAct_l : boolean;
   sCaption_l : string;
begin
   inherited;
	If (F.FieldName = 'JUR_DOSLIBELLE') then
   begin
      sCaption_l := 'Fiche dossier : ';
      if (sTypeFiche_c = 'JUR') and (GetField('JUR_CODEDOS') <> '') and (GetField('JUR_CODEDOS') <> '&@#') then
         sCaption_l := sCaption_l + GetField('JUR_CODEDOS') + ' ';
      sCaption_l := sCaption_l + GetField('JUR_DOSLIBELLE');
      TFFiche(ecran).Caption := sCaption_l;
   end
   //////
   else if (F.FieldName = 'JUR_FORME') then
   begin
      bAct_l := (FctDesPropDeTitres(GetField('JUR_FORME'), 'ACT') <> '');
      SetControlEnabled('BACTIONNAIRE', bAct_l);

      bChangeForme_c := false;
      if (Ds.State = dsBrowse) then
      begin
         if (GetField('JUR_FORME') = '') and (sAnnForme_c <> '') then
         begin
            ModeEdition(DS);
            SetField('JUR_FORME', sAnnForme_c);
         end;
      end
      else if not (Ds.State in [dsInsert]) then
      begin
         if (GetField('JUR_FORME') <> sAnnForme_c) then
         begin
            PGIInfo( 'Attention : la forme juridique du dossier#13#10' +
                     'et la forme juridique de l''annuaire sont différentes.#13#10' +
                     'Vous devez modifier les fonctions des personnes rattachées #13#10' +
                     'par la fiche annuaire.#13#10',
                   Ecran.Caption );
            bChangeForme_c := true;
         end;
      end;
      LiensChargeEtAffiche;
      ActionnairesAffiche;
      GereCac;
      GereGerance;
      GereOrganeDir;
      GereActionPreference;
      ForceRegimeFiscal;
      AfficheMasqueSelonFormeJur;

//      SetControlEnabled('JUR_FORME', false);
   end
   //*****//
   else If (F.FieldName = 'JUR_NBDROITSVOTE') then
   begin
      if GetField('JUR_NBTITRESCLOT')<>GetField('JUR_NBDROITSVOTE') then
      begin
         ModeEdition(DS);
         SetField('JUR_NBTITRESCLOT', GetField('JUR_NBDROITSVOTE'));
      end;
   end
   //*****//
   else if(F.FieldName = 'JUR_DUREE') then
   begin
      if GetControlText('ANN_RCSDATE') = null then
         dtRcsDate_l := iDate1900
      else
         dtRcsDate_l := ANN_.Get('ANN_RCSDATE');

      if GetField('JUR_DATEXP') = null then
         dtDateExp_l := iDate1900
      else
         dtDateExp_l := GetField('JUR_DATEXP');

      if GetField('JUR_DUREE') = null then
         nDuree_l := 0
      else
         nDuree_l := GetField('JUR_DUREE');

      dtDateExp_l := DateExpCalcul( dtDateExp_l, dtRcsDate_l, nDuree_l);
      if GetField('JUR_DATEXP')<>dtDateExp_l then
      begin
         ModeEdition(DS);
         SetField('JUR_DATEXP', dtDateExp_l);
      end;
   end
   else if F.FieldName = 'JUR_RESERVESTATUT' then
   begin
      SetControlVisible('JUR_MONTANTSTATUT', (GetField('JUR_RESERVESTATUT') = 'X'));
      SetControlVisible('TJUR_MONTANTSTATUT', (GetField('JUR_RESERVESTATUT') = 'X'));
   end
   else if F.FieldName = 'JUR_NATUREGERANCE' then
   begin
      SetControlVisible('TJUR_REFCOGERANTDP', (GetField('JUR_NATUREGERANCE') = 'COG'));
      SetControlVisible('TCOGERANT', (GetField('JUR_NATUREGERANCE') = 'COG'));
   end;

//   else if F.FieldName = 'ANN_PERASS1GUID' then
//    AfficheInfoLienPersonne(self, GetField('ANN_PERASS1GUID'), 'ANN_PERASS1GUID', 'TANN_PERASS1GUID')//LM20070516

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnUpdateRecord ;
begin
   inherited;
   if (GetField('JUR_FORME') = '') and (GetField('JUR_TYPEDOS') = 'STE') then
   begin
      LastError := 1 ;
      LastErrorMsg := 'La saisie d''une forme juridique est obligatoire.#10#13' +
                      'Vous devez modifier la fiche de la personne sélectionnée.' ;
      SetFocusControl('JUR_FORME');
      exit;
   end;
   if (GetField('JUR_GUIDPERDOS') = '') then
   begin
      LastError := 1 ;
      LastErrorMsg:='La saisie d''une personne est obligatoire.' ;
      SetFocusControl('JUR_GUIDPERDOS');
      exit;
   end;

   BlobsEnreg( OBBlob_c );
   if bSuppCAC_c then
      ExecuteSQL('DELETE from ANNULIEN ' +
                 'WHERE ANL_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                 '  and (ANL_FONCTION = "CCT" OR ANL_FONCTION = "CCS")');

   // Synchronisation
   DFI_.Put('DFI_IMPODIR', GetField('JUR_IMPODIR'));
   
   ANN_.Sauve(getField('JUR_GUIDPERDOS'));//LM20070516
   DFI_.Sauve(getField('JUR_GUIDPERDOS'));//LM20070516
   DOR_.Sauve(getField('JUR_GUIDPERDOS'));//MD20070622
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnAfterUpdateRecord;
begin
   inherited;
   SetControlEnabled('BLIENS',true);
   if IsInside(Ecran) then
      ReloadTomInsideAfterInsert(TFFiche(Ecran), DS,
                     ['JUR_GUIDPERDOS', 'JUR_TYPEDOS', 'JUR_NOORDRE'],
                     [sGuidPerDos_c, sTypeDos_c, sNoOrdre_c]);
end;

{procedure TOM_YYJURIDIQUE.VraiOnClose (Sender:TObject; var Action:TCloseAction); //LM20070704
begin

   ANN_.Free ;
   DFI_.Free ;
   DOR_.Free ;

   if gdActionnaires_c <> nil then
      gdActionnaires_c.Destroy;

   if cCapital_c <> nil then
      cCapital_c.Free;

   if OBBlob_c <> nil then
      OBBlob_c.Free;
   if OBFonctions_c <> nil then
      OBFonctions_c.Free;

  Evt.Exec('Juridique', 'OnClose', sender, action);

  //evt.Free ; Effacer par le destroy de la forme.
end ;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClose;
begin
   inherited;
   ANN_.Free ;
   DFI_.Free ;
   DOR_.Free ;

   if gdActionnaires_c <> nil then
      gdActionnaires_c.Destroy;

   if cCapital_c <> nil then
      cCapital_c.Free;

   if OBBlob_c <> nil then
      OBBlob_c.Free;
   if OBFonctions_c <> nil then
      OBFonctions_c.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnChange_BLOB( Sender : TObject );
var
   sControlName_l : string;
begin
   sControlName_l := TControl( Sender ).Name;
   Change_BLOB(sControlName_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 31/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnChange_RCSDATE(Sender : TObject);
var
   dtDateRCS_l, dtDateExp_l : TDateTime;
begin
   if bCharge_c then exit;
   try
      dtDateRCS_l := StrToDate(GetControlText('ANN_RCSDATE'));
   except
      exit;
   end;
   ModeEdition(DS);
   SetControlText('ANN_RCSDATE1', GetControlText('ANN_RCSDATE'));
   dtDateExp_l := DateExpCalcul(GetField('JUR_DATEXP'), dtDateRCS_l, GetField('JUR_DUREE'));
   SetField('JUR_DATEXP', dtDateExp_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.Change_BLOB(sControlName_p : string);
var
   sPrefixe_l, sEmploiBlob_l, sRangBlob_l, sLibelle_l : string;
begin
   BlobGetCode(sControlName_p, sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);

   if BlobChange(OBBlob_c, sPrefixe_l, GetField('JUR_GUIDPERDOS'), sRangBlob_l,
                  TRichEdit(GetControl(sControlName_p)) ) then
   begin
      ModeEdition(DS);
      // Pour forcer la mise à jour des données
      SetField('JUR_NOMDOS', GetField('JUR_NOMDOS'));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.VoirPersonne;
var
   sCle_l, sInfosDossier_l : string;
begin
   if (GetField('JUR_GUIDPERDOS') <> '' ) then
   begin
      if sTypeFiche_c = 'JUR' then
         sCle_l := ';;;;' + GetField('JUR_CODEDOS') + ';JUR;'
      else
         sCle_l := ';;;;;JUR;';

      sInfosDossier_l := InfosTransmisesPrepare;

      sInfosDossier_l := AGLLanceFiche('YY','ANNUAIRE', GetField('JUR_GUIDPERDOS'),
                  GetField('JUR_GUIDPERDOS'), sCle_l + sInfosDossier_l);

      if (sInfosDossier_l <> '0') and (sInfosDossier_l <> '' ) then
      begin
         ReadTokenSt(sInfosDossier_l); // on saute la chaine '1;'
         InfosTransmisesRecupere(sInfosDossier_l);
      end;
      DossierChargeEtAffiche(GetField('JUR_GUIDPERDOS'), false);
      BlobsReCharge(GetField('JUR_GUIDPERDOS'), GetField('JUR_TYPEDOS'));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BVOIRPERSONNE(Sender : TObject);
begin
   VoirPersonne;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BTITRES(Sender : TObject);
begin
   OnDblClick_FLISTE(Sender);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 13/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnDblClick_FLISTE(Sender : TObject);
var
   iTTNBPP_l : integer;
   sGuidPer_l, sCle1_l, sCle2_l, sFonction_l : string;
begin
   if not GetControlEnabled('BTITRES') then exit;

   // Vérif pour traitement de la clé
   If Not gdActionnaires_c.ListeCorrecte('JMT_GUIDPER', 'Guid de la fiche personne') then exit;
   If Not gdActionnaires_c.ListeCorrecte('JMT_FONCTION', 'Fonction personne') then exit;
   If Not gdActionnaires_c.ListeCorrecte('JMT_TTNBPP', 'Nombre de titres') then exit;

   sGuidPer_l  := gdActionnaires_c.OnGetField('JMT_GUIDPER');
   sFonction_l := gdActionnaires_c.OnGetField('JMT_FONCTION');
   if sFonction_l = '' then
      sFonction_l := sFonctDefaut_c;

   iTTNBPP_l := gdActionnaires_c.OnGetField('JMT_TTNBPP');

   sCle1_l := 'JMT_GUIDPERDOS=' + GetField('JUR_GUIDPERDOS') + ';' +
              'JMT_TYPEDOS=' + GetField('JUR_TYPEDOS') + ';' +
              'JMT_GUIDPER=' + sGuidPer_l + ';' +
              'JMT_NATURECPT=1';  // Compte principal uniquement

   sCle2_l := GetField('JUR_GUIDPERDOS') + ';' +
              GetField('JUR_TYPEDOS') + ';' +
              sGuidPer_l + ';' +
              '1';  // Compte principal uniquement

   AGLLanceFiche('JUR', 'JUMVTTITRES_MUL', sCle1_l, '',
                        'ACTION=MODIFICATION;' + sCle2_l + ';' +
                        'Dossier ' + GetField('JUR_DOSLIBELLE') + ';' +
                         sFonction_l +  ';' +
                         VarToStr(GetField('JUR_VALNOMINCLOT')) + ';' +
                         IntToStr(iTTNBPP_l));

   if CapitalRefresh then
      CapitalMAJ;

   gdActionnaires_c.OnReload;
   SetControlEnabled('BTITRES', (gdActionnaires_c.NbRows > 0)); //and ActionnaireAutorise);
   THGrid(GetControl('FLISTE')).ClearSelected;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BIMPRIMER(Sender: TObject);
begin
   AGLLanceFiche('JUR', 'PRINTINFODOS', '', '', GetField('JUR_GUIDPERDOS'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 27/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BLIENS(Sender: TObject);
var
   sRetour_l : string;
begin
   if (GetField('JUR_FORME') <> '') then
   begin
      if DS.State <> dsBrowse then
      begin
         TFFiche(Ecran).BValiderClick(nil);
	      Ecran.ModalResult := mrNone;
      end;

      sRetour_l := AGLLanceFiche('JUR','ANNULIEN_TRI', '', '',
                                 GetField('JUR_GUIDPERDOS') + ';' +
                                 GetField('JUR_TYPEDOS') + ';' +
                                 IntToStr(GetField('JUR_NOORDRE')) + ';' +
                                 GetField('JUR_FORME') + ';' +
                                 GetField('JUR_CODEDOS') + ';' +
                                 sTypeFiche_c+ ';'+
                                 {+GHA - 12/2007 - concept KPMG}
                                 TForm(globalfichePile[0]).Name);
                                 {-GHA - 12/2007 - concept KPMG}


      gdActionnaires_c.OnReload;
      SetControlEnabled('BTITRES', (gdActionnaires_c.NbRows > 0)); //and ActionnaireAutorise);
      
//      if sRetour_l = 'X' then
      LiensChargeEtAffiche;

      GereActionPreference;
      GereGerance;
      GereOrganeDir;
      GereCac;
  end
  else
      PGIInfo('La forme juridique n''est pas renseignée : ' + #13#10 +
              ' vous ne pouvez pas accéder aux liens du dossier');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BEVENEMENTS(Sender: TObject);
begin
   if GetField('JUR_CODEDOS') <> '' then
   begin
      if DS.State <> dsBrowse then
      begin
         TFFiche(Ecran).BValiderClick(nil);
	      Ecran.ModalResult := mrNone;
      end;
      AGLLanceFiche('JUR', 'EVENEMENT_MUL', 'JEV_CODEDOS=' + GetField('JUR_CODEDOS'), '',
                    GetField('JUR_CODEDOS'));
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BTYPEINFOS(Sender: TObject);
begin
   if GetField('JUR_CODEDOS') <> '' then
   begin
      if DS.State <> dsBrowse then
      begin
         TFFiche(Ecran).BValiderClick(nil);
	      Ecran.ModalResult := mrNone;
      end;
      AGLLanceFiche('JUR', 'TYPEINFO', '', '', GetField('JUR_CODEDOS'));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BDelete(Sender: TObject);
begin
   if GetField('JUR_CODEDOS') <> '' then
      SupprimeDossier('JURIDIQUE', GetField('JUR_CODEDOS'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.DossierChargeEtAffiche(sGuidPerDos_p : string; bInsert_p : boolean );
Var
   sNom1_l, sNom2_l, sNom3_l, sAnnCapDev_l, sAnnMoisCloture_l : string;
   vMontant_l : variant;
   nAnnCapital_l, nAnnCapNbTitre_l, sAnnCapVn_l : integer;
//   bVisible_l : boolean;
begin
   // Entete
   ANN_.ChargeEnreg(sGuidPerDos_p);
   SetControlText('ANN_RCSDATE1', ANN_.Get('ANN_RCSDATE'));

   sNom1_l := ANN_.Get('ANN_NOM1');
   sNom2_l := ANN_.Get('ANN_NOM2');
   sNom3_l := ANN_.Get('ANN_NOM3');
   sPPPM_c := ANN_.Get('ANN_PPPM');

   if sNom2_l <> '' then
      sNom1_l := sNom1_l + ' ' + sNom2_l;
   if sNom3_l <> '' then
      sNom1_l := sNom1_l + ' ' + sNom3_l;

   if GetField('JUR_DOSLIBELLE') <> Copy(sNom1_l, 1, 70) then
   begin
      ModeEdition(DS);
      SetField('JUR_DOSLIBELLE', Copy(sNom1_l, 1, 70));
   end;

   if GetControl('TJUR_GUIDPERDOS') <> nil then
      SetControlText('TJUR_GUIDPERDOS', sNom1_l);

   // Synchronisation
   if GetField('JUR_NOMDOS') <> ANN_.Get('ANN_NOMPER') then
   begin
      ModeEdition(DS);
      SetField('JUR_NOMDOS', ANN_.Get('ANN_NOMPER'));
   end;
   
   sAnnForme_c := ANN_.Get('ANN_FORME');
   sAnnCapDev_l := ANN_.Get('ANN_CAPDEV');
   nAnnCapital_l := ANN_.Get('ANN_CAPITAL');
   nAnnCapNbTitre_l := ANN_.Get('ANN_CAPNBTITRE');
   sAnnMoisCloture_l := ANN_.Get('ANN_MOISCLOTURE');
   sAnnCapVn_l := ANN_.Get('ANN_CAPVN');
   sAnnCodeInsee_c := ANN_.Get('ANN_CODEINSEE');

//   if sTypeFiche_c <> 'DP' then
//   begin
   if bInsert_p or (GetField('JUR_CODEDOS') = '&#@') then
   begin
      // En création, initialisation avec les données de l'annuaire
      ModeEdition(DS);
      SetField('JUR_FORME', sAnnForme_c );
      SetField('JUR_CAPDEV', sAnnCapDev_l);
      SetField('JUR_CAPITAL', nAnnCapital_l);
      SetField('JUR_NBDROITSVOTE', nAnnCapNbTitre_l);
      SetField('JUR_NBTITRESCLOT', nAnnCapNbTitre_l);
      SetField('JUR_MOISCLOTURE', sAnnMoisCloture_l);
      SetField('JUR_VALNOMINCLOT', sAnnCapVn_l);
   end;
//   end;

   //à vérifier
   vMontant_l := VarAsType(ANN_.Get('ANN_CAPLIB'), varDouble);
   SetControlText('ANN_CAPLIB', vMontant_l);

   vMontant_l := VarAsType(ANN_.Get('ANN_CAPVN'), varDouble);
   if GetControl('ANN_CAPVN') <> nil then
      SetControlText('ANN_CAPVN', vMontant_l);
   if GetControl('ANN_CAPDEV_') <> nil then
      SetControlText('ANN_CAPDEV_', GetField('JUR_CAPDEV'));
   if GetControl('ANN_CAPDEV__') <> nil then
      SetControlText('ANN_CAPDEV__', GetField('JUR_CAPDEV'));

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. : Données à prendre en compte par l'annuaire
Mots clefs ... :
*****************************************************************}
function TOM_YYJURIDIQUE.InfosTransmisesPrepare : string;
var
   sInfosDossier_l : string;
begin
   sInfosDossier_l := 'JUR_NOMDOS=' + GetField('JUR_NOMDOS') + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_FORME=' + GetField('JUR_FORME') + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_CAPDEV=' + GetField('JUR_CAPDEV') + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_CAPITAL=' + IntToStr(GetField('JUR_CAPITAL')) + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_NBDROITSVOTE=' + IntToStr(GetField('JUR_NBDROITSVOTE')) + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_NBTITRESCLOT=' + IntToStr(GetField('JUR_NBTITRESCLOT')) + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_MOISCLOTURE=' + GetField('JUR_MOISCLOTURE') + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_VALNOMINCLOT=' + IntToStr(GetField('JUR_VALNOMINCLOT')) + ';';
   sInfosDossier_l := sInfosDossier_l + 'JUR_NOMPER=' + GetField('JUR_NOMPER') + ';';

   result := sInfosDossier_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.InfosTransmisesRecupere( sInfosDossier_p : string);
var
   sNomChamp_l, sValChamp_l : string;
begin
   if sInfosDossier_p <> '' then
      ModeEdition(DS);

   // Mise à jour des Fields
   while sInfosDossier_p <> '' do
   begin
      sValChamp_l := ReadTokenSt(sInfosDossier_p);
      sNomChamp_l := ReadTokenPipe(sValChamp_l, '=');
      if sNomChamp_l = 'JUR_FORME' then
        sAnnForme_c := sValChamp_l;
      SetField(sNomChamp_l, sValChamp_l);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.JURDateExercice(var dDateDeb_p, dDateFin_p : TDateTime);
var
   iAnnee_l : integer;
   sDateDeb_l, sDateFin_l, sMois_l : string;
begin
   sMois_l := GetField('JUR_MOISCLOTURE');
   if sMois_l = '' then
      exit;

   iAnnee_l := StrToInt(Copy(DateToStr(Date), 7, 4));
   sDateDeb_l := '01/' + sMois_l + '/' + IntToStr(iAnnee_l);

   if StrToDate(sDateDeb_l) >= Date then
      Dec(iAnnee_l);

   sDateDeb_l := '01/' + sMois_l + '/' + IntToStr(iAnnee_l);
   dDateDeb_p := StrToDateTime(sDateDeb_l);

   sDateFin_l := '01/' + sMois_l + '/' + IntToStr(iAnnee_l + 1);
   dDateFin_p := StrToDateTime(sDateFin_l);
   dDateFin_p := PlusDate(dDateFin_p, -1, 'J');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
(*LM20070412
procedure TOM_YYJURIDIQUE.DPDateExercice(var dDateDeb_p, dDateFin_p : TDateTime);
var
   OBExercice_l : TOB;
begin
   dDateDeb_p := iDate1900;
   dDateFin_p := iDate1900;
   OBExercice_l := TOB.Create('DPORGA', nil, -1);
   OBExercice_l.LoadDetailFromSQL('select DOR_DATEDEBUTEX, DOR_DATEFINEX ' +
                                  'from DPORGA ' +
                                  'where DOR_GUIDPER = "' + sGuidPerDos_c + '"') ;

   if OBExercice_l.Detail.Count > 0 then
   begin
      dDateDeb_p := OBExercice_l.Detail[0].GetValue('DOR_DATEDEBUTEX');
      dDateFin_p := OBExercice_l.Detail[0].GetValue('DOR_DATEFINEX');
   end;
   OBExercice_l.Free;
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.CapitalCalcule(dDateDeb_p, dDateFin_p : TDateTime; sDomaine_p : string);
begin
   cCapital_c.Init(GetField('JUR_GUIDPERDOS'), GetField('JUR_NBTITRESCLOT'),
                   dDateDeb_p, dDateFin_p, sDomaine_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 06/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function TOM_YYJURIDIQUE.CapitalRefresh : boolean;
begin
   result := false;
   if cCapital_c.CompareCapital(true) then
   begin
      ModeEdition(DS);
      result := true;
      // Maj du capital juridique
//         SetField('JUR_NBTITRESOUV', cCapital_c.iNbTitresOuv_c);
      SetField('JUR_NBTITRESAUGM', cCapital_c.iNbTitresAugm_c);
      SetField('JUR_NBTITRESRED', cCapital_c.iNbTitresRed_c);
      SetField('JUR_NBTITRESCLOT', cCapital_c.iNbTitresClo_c);
      SetField('JUR_NBDROITSVOTE', cCapital_c.iNbTitresClo_c);
      SetField('JUR_CAPITAL', GetField('JUR_VALNOMINCLOT') * cCapital_c.iNbTitresClo_c);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.CapitalMAJ;
begin
   if PGIAsk('Voulez-vous mettre à jour le Capital social de l''annuaire ?', TitreHalley ) = mrYes then
   begin
      // Maj capital annuaire;
      ANN_.Put('ANN_CAPVN', GetField('JUR_VALNOMINCLOT'));
      ANN_.Put('ANN_CAPNBTITRE', GetField('JUR_NBTITRESCLOT'));
      ANN_.Put('ANN_CAPITAL', GetField('JUR_CAPITAL'));
      ANN_.tb_.PutEcran(Ecran);
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.BlobsCharge(sGuidPerDos_p : string);
var
  OBBlob_l : TOB;
  sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l : string;
begin
   OBBlob_c.ClearDetail;

   if GetControl('OLE_OBJETSOC') <> nil then
   begin
      BlobGetCode('OLE_OBJETSOC', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidPerDos_p, sRangBlob_l, sEmploiBlob_l,
                                    sLibelle_l, TRichEdit(GetControl('OLE_OBJETSOC'))) ;
      OBBlob_l.ChangeParent(OBBlob_c, -1);
   end;
   if GetControl('OLE_BLOCNOTES') <> nil then
   begin
      BlobGetCode('OLE_BLOCNOTES', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidPerDos_p, sRangBlob_l, sEmploiBlob_l,
                              sLibelle_l, TRichEdit(GetControl('OLE_BLOCNOTES'))) ;
      OBBlob_l.ChangeParent(OBBlob_c, -1);
   end;
   if GetControl('OLE_ACTIVITE') <> nil then
   begin
      BlobGetCode('OLE_ACTIVITE', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidPerDos_p, sRangBlob_l, sEmploiBlob_l,
                              sLibelle_l, TRichEdit(GetControl('OLE_ACTIVITE'))) ;
      OBBlob_l.ChangeParent(OBBlob_c, -1);
   end;
   if GetControl('OLE_ACTIVITE2') <> nil then//LM20070315
   begin
      BlobGetCode('OLE_ACTIVITE2', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidPerDos_p, sRangBlob_l, sEmploiBlob_l,
                              sLibelle_l, TRichEdit(GetControl('OLE_ACTIVITE2'))) ;
      OBBlob_l.ChangeParent(OBBlob_c, -1);
   end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 28/04/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.BlobsReCharge(sGuidPerDos_p, sTypeDos_p : string );
var
  OBBlob_l : TOB;
  sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l : string;
begin
   BlobGetCode('OLE_ACTIVITE', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
   OBBlob_l := OBBlob_c.FindFirst(['LO_TABLEBLOB', 'LO_IDENTIFIANT', 'LO_RANGBLOB'],
                                  ['ANN', sGuidPerDos_p, 1], true);

   if OBBlob_l = nil then
      Exit;

   OBBlob_l.Free;

   OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidPerDos_p, sRangBlob_l, sEmploiBlob_l,
                                 sLibelle_l, TRichEdit(GetControl('OLE_ACTIVITE'))) ;

   OBBlob_l.ChangeParent(OBBlob_c, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_YYJURIDIQUE.DateExpCalcul( dtDateExp_p, dtRcsDate_p : TDateTime; nDuree_p : integer) : TDateTime;
var
  dtDate_l : TDateTime;
  sDate_l : string;
  wYear_l, wMonth_l, wDay_l : word;

begin
   dtDate_l := Now;
   DecodeDate(dtDate_l, wYear_l, wMonth_l, wDay_l);
   sDate_l := IntToStr(wDay_l) + '/' + IntToStr(wMonth_l) + '/' + IntToStr(wYear_l);
   // dtDate_l := strToDate(sDate_l);

   if (dtRcsDate_p = iDate1900) then
   begin
      dtDateExp_p := iDate1900;
   end
   else //if (dtDateExp_p < dtDate_l) and (nDuree_p <> 0) then
   begin
      nDuree_p := nDuree_p * 12;
      dtDateExp_p := PLUSMOIS(dtRcsDate_p, nDuree_p) - 1;
   end;
   result := dtDateExp_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnFormKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
var
   FListe_l : THGrid;
begin
   FListe_l := THGrid(GetControl('FLISTE'));

   if (MyKey_p = VK_F5) then  // Sélection ouvrir
   begin
      MyKey_p := 0;
      if (FListe_l <> nil) and (FListe_l.Focused) and
         (TToolbarButton97(GetControl('BTITRES')).Enabled) and
         (TToolbarButton97(GetControl('BTITRES')).Visible) then
         TToolbarButton97(GetControl('BTITRES')).OnClick(nil);
   end
   else if (MyKey_p = VK_F3) and (ssCtrl in MyShift_p) then
   begin
      MyKey_p := 0;
      if (FListe_l <> nil) and (FListe_l.Focused) then
         Postmessage(FListe_l.Handle, WM_KEYDOWN, vk_home, 0);
   end
   else if (MyKey_p = VK_F4) and (ssCtrl in MyShift_p) then
   begin
      MyKey_p := 0;
      if (FListe_l <> nil) and (FListe_l.Focused) then
         Postmessage(FListe_l.Handle, WM_KEYDOWN, vk_end, 0);
   end
   else if (MyKey_p = VK_F3) and (MyShift_p = []) then
   begin
      MyKey_p := 0;
      if (FListe_l <> nil) and (FListe_l.Focused) then
         Postmessage(FListe_l.Handle, WM_KEYDOWN, vk_up, 0);
   end
   else if (MyKey_p = VK_F4) and (MyShift_p = []) then
   begin
      MyKey_p := 0;
      if (FListe_l <> nil) and (FListe_l.Focused) then
         Postmessage(FListe_l.Handle, WM_KEYDOWN, vk_down, 0);
   end
   else
      TFFiche(Ecran).FormKeyDown(MySender_p, MyKey_p, MyShift_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.GereActionPreference;
var
   bActionPref_l : boolean;
   sRequete_l : string;
begin
   bActionPref_l := false;

   if GetField('JUR_GUIDPERDOS') <> '' then
   begin
      sRequete_l := 'SELECT ANL_TITNATURE FROM ANNULIEN ' +
                    'WHERE ANL_GUIDPERDOS = "' + GetField('JUR_GUIDPERDOS') + '" ' +
                    'AND EXISTS (SELECT YNT_CODE FROM YNATURETIT ' +
                    '            WHERE YNT_CODE LIKE "AP%" and ANL_TITNATURE = YNT_CODE)';

      bActionPref_l := ExisteSQL(sRequete_l);
   end;
   SetControlVisible('LACTIONPREF_F', bActionPref_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{function TOM_YYJURIDIQUE.ExisteCAC : string;
var
   sRequete_l : string;
   OBCCT_l, OBCCS_l : TOB;
begin
   result := '-';
   sRequete_l := 'SELECT 1 FROM ANNULIEN ' +
                 'WHERE ANL_GUIDPERDOS = "' + sGuidPerDos_c + '"' +
                 '  AND (ANL_FONCTION = "CCT" OR ANL_FONCTION = "CCS")';

   if ExisteSQL(sRequete_l) then
      result := 'X';

   OBCCT_l := OBFonctions_c.FindFirst(['JTF_FONCTION'], ['CCT'], true);
   OBCCS_l := OBFonctions_c.FindFirst(['JTF_FONCTION'], ['CCS'], true);
   if ((OBCCT_l = nil) and (OBCCS_l = nil)) or
      ((OBCCT_l.Detail.count = 0) and (OBCCS_l.Detail.count = 0)) then
      result := '-'
   else
      result := 'X';
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 19/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYJURIDIQUE.InFonctionsTitres : string;
var
   sFonction_l : string;
   iInd_l, iFirst_l : integer;
begin
   iInd_l := 0;
   iFirst_l := 0;
   sFonction_l := '""';
   while iInd_l < OBFonctions_c.Detail.Count do
   begin
      if OBFonctions_c.Detail[iInd_l].GetValue('JFT_TITRE') = 'X' then
      begin
         if iFirst_l = 0 then
            sFonctDefaut_c := OBFonctions_c.Detail[0].GetValue('JFT_FONCTION');
         sFonction_l := sFonction_l + ',"' + OBFonctions_c.Detail[iInd_l].GetValue('JFT_FONCTION') + '"';
         Inc(iFirst_l);
      end;
      Inc(iInd_l);
   end;
   result := sFonction_l;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.ActionnairesAffiche;
var
   sRequete_l, sFonction_l, sTmp_l : string;
   OBFonctions_l : TOB;
   iInd_l : integer;
begin
   SetControlEnabled('BTITRES', false);
//   ActionnaireFormateLibelles(sFonction_l, sFonctAbrege_l);
   TToolbarButton97(GetControl('BTITRES')).Caption := 'Compte de titres';// + sFonction_l;
   TTabSheet(GetControl('TSACTIONNAIRES')).Caption := 'Titres';

//   sFonctionsTitres_c := InFonctionsTitres;

   OBFonctions_l := TOB.Create('JUFONCTION', nil, -1);
   OBFonctions_l.LoadDetailFromSQL('select JFT_FONCTION from JUFONCTION ' +
                                   ' where JFT_TYPEDOS = "' + GetField('JUR_TYPEDOS') + '" ' +
                                   '   and JFT_FORME = "' + GetField('JUR_FORME') + '" ' +
                                   '   AND JFT_TITRE = "X"');

   iInd_l := 0;
   sFonction_l := '""';
   while iInd_l < OBFonctions_l.Detail.Count do
   begin
      if iInd_l = 0 then
         sFonctDefaut_c := OBFonctions_l.Detail[0].GetValue('JFT_FONCTION');
      sFonction_l := sFonction_l + ',"' + OBFonctions_l.Detail[iInd_l].GetValue('JFT_FONCTION') + '"';
      Inc(iInd_l);
   end;
   OBFonctions_l.Free;
{   sFonction_l := 'AND JMT_FONCTION IN (' + sFonction_l + ')';

   if GetField('JUR_GUIDPERDOS') <> 0 then  // sinon création dossier
      sRequete_l := 'JMT_GUIDPERDOS = ' + IntToStr(GetField('JUR_GUIDPERDOS'))
   else
      sRequete_l := 'JMT_GUIDPERDOS = 0';

   sRequete_l := sRequete_l + ' AND JMT_TYPEDOS = "' + GetField('JUR_TYPEDOS') + '" ' + sFonction_l;
   gdActionnaires_c.OnLoad(sRequete_l);}

   sTmp_l := GetField('JUR_GUIDPERDOS');

   sRequete_l := 'select distinct jmt_GUIDperdos, jmt_GUIDper, ' +
                 '       jmt_typedos, jmt_noordre, ' +
                 '       (ANN_NOM1|| " " ||ANN_NOM2) as jmt_nomper, ' +
                 '       ANN_CV as jmt_cv, anl_fonction as jmt_fonction, ' +
                 '       jmt_date, yno_libelle, ' +
                 '       ANL_TTNBNP As jmt_TTNBNP, ANL_TTNBPP As jmt_TTNBPP, ' +
                 '       ANL_TTNBUS As jmt_TTNBUS, ANL_TTNBTOT As jmt_TTNBTOT ' +
                 'from annulien, annuaire, jumvttitres ' +
                 'left join ynatureop on jmt_natureop = yno_code ' +
                 'where anl_GUIDperdos = jmt_GUIDperdos and anl_GUIDper = jmt_GUIDper ' +
                 '  and jmt_typedos = anl_typedos and anl_GUIDper = ann_GUIDper ' +
                 '  and jmt_noordre = (select max(jmt_noordre) from jumvttitres ' +
                 '                     where jmt_GUIDper = anl_GUIDper and jmt_GUIDperdos = anl_GUIDperdos ' +
                 '                     and jmt_typedos = anl_typedos) ' +
                 ///
                 '  and JMT_GUIDPERDOS = "' + sTmp_l + '"' +
                 '  AND JMT_TYPEDOS = "' + GetField('JUR_TYPEDOS') + '" ' +
                 '  AND anl_FONCTION IN (' + sFonction_l + ')';
   sRequete_l := sRequete_l + ' union ';

   sRequete_l := sRequete_l +
                 'select distinct anl_GUIDperdos as jmt_GUIDperdos, anl_GUIDper as jmt_GUIDper, ' +
                 '       anl_typedos as jmt_typedos, 0 as jmt_ordre, ' +
                 '       (ANN_NOM1|| " " ||ANN_NOM2) as jmt_nomper, ' +
                 '       ANN_CV as jmt_cv, anl_fonction as jmt_fonction, ' +
                 '       "" as jmt_date, "" as yno_libelle, ' +
                 '       ANL_TTNBNP As jmt_TTNBNP, ANL_TTNBPP As jmt_TTNBPP, ' +
                 '       ANL_TTNBUS As jmt_TTNBUS, ANL_TTNBTOT As jmt_TTNBTOT ' +
                 'from annulien, annuaire ' +
                 'where anl_GUIDper = ann_GUIDper ' +
                 '  and not exists (select jmt_noordre from jumvttitres ' +
                 '                  where jmt_GUIDper = anl_GUIDper and jmt_GUIDperdos = anl_GUIDperdos ' +
                 '                  and jmt_typedos = anl_typedos) ' +
                 ///
                 '  and anl_GUIDPERDOS = "' + sTmp_l + '"' +
                 '  AND anl_TYPEDOS = "' + GetField('JUR_TYPEDOS') + '" ' +
                 '  AND anl_FONCTION IN (' + sFonction_l + ')';

   sRequete_l := sRequete_l + ' union ';

   sRequete_l := sRequete_l +
                 'select distinct jmt_GUIDperdos, jmt_GUIDper, ' +
                 '       jmt_typedos as typedos, jmt_noordre, ' +
                 '       (ANN_NOM1|| " " ||ANN_NOM2) as jmt_nomper, ' +
                 '       ANN_CV as jmt_cv, "" as jmt_fonction, ' +
                 '       jmt_date, yno_libelle, ' +
                 '       0 As jmt_TTNBNP, 0 As jmt_TTNBPP, ' +
                 '       0 As jmt_TTNBUS, 0 As jmt_TTNBTOT ' +
                 'from annuaire, jumvttitres ' +
                 'left join ynatureop on jmt_natureop = yno_code ' +
                 'where jmt_GUIDper = ann_GUIDper ' +
                 '  and jmt_noordre = (select max(jmt_noordre) from jumvttitres TIT ' +
                 '                   where jmt_GUIDper = TIT.jmt_GUIDper and jmt_GUIDperdos = TIT.jmt_GUIDperdos ' +
                 '                    and jmt_typedos = TIT.jmt_typedos) ' +
                 '  and not exists (select anl_noordre from annulien ' +
                 '                  where jmt_GUIDper = anl_GUIDper and jmt_GUIDperdos = anl_GUIDperdos ' +
                 '                  and jmt_typedos = anl_typedos) ' +
                 ///
                 '  and JMT_GUIDPERDOS = "' + sTmp_l + '"' +
                 '  AND JMT_TYPEDOS = "' + GetField('JUR_TYPEDOS') + '" ';
                 ///
   sRequete_l := sRequete_l + ' order by 6';
                 ///
   gdActionnaires_c.OnLoad(sRequete_l,
                           'Dossier;Personne;Type dossier;Ordre;' +
                           'Nom;Civilité;Fonction;Der. Mvt;Nat. der. opération;' +
                           'Nue-propriété;Pleine propriété;Usufruit;Titres total;Formule;',
                           '8;8;8;8;' +
                           '20;8;8;10;13;' +
                           '8;8;8;8;16;',
                           'G.0O X--;G.0O X--;G.0O X--;G.00 X--;' +
                           'G.0  -X-;G.0  -X-;G.0  -X-;G.0  -X-;G.0  -X-;' +
                           'D.0  ---;D.0  ---;D.0  ---;D.0  ---;D/2  X--;');
   SetControlEnabled('BTITRES', (gdActionnaires_c.NbRows > 0)); //and ActionnaireAutorise);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 15/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
{ procedure TOM_YYJURIDIQUE.ActionnaireFormateLibelles(var sFonction_p, sFonctAbrege_p : string);
var
   iInd_l : integer;
   OBFonctions_l : TOB;
begin
   sFonction_p := '';
   sFonctAbrege_p := '';

   OBFonctions_l := TOB.Create('JUTYPEFONCT', nil, -1);
   OBFonctions_l.LoadDetailDBFromSQL('JUTYPEFONCT',
              'select JTF_FONCTABREGE ' +
              'from (JUTYPEFONCT left join JUFONCTION on JTF_FONCTION = JFT_FONCTION) ' +
              'where (JFT_FORME = "' + GetField('JUR_FORME') + '" ' +
              '  AND JFT_TYPEDOS = "' + GetField('JUR_TYPEDOS') + '" ' +
              '  AND JFT_TITRE = "X") ' +
              'ORDER BY JFT_TRI');

   iInd_l := 0;
   while iInd_l < OBFonctions_l.Detail.Count do
   begin
      if sFonction_p <> '' then
         sFonction_p := sFonction_p + '/';
      if sFonctAbrege_p <> '' then
         sFonctAbrege_p := sFonctAbrege_p + '/';
      sFonction_p := sFonction_p + OBFonctions_l.Detail[iInd_l].GetValue('JTF_FONCTABREGE');
      sFonctAbrege_p := sFonctAbrege_p + Copy(OBFonctions_l.Detail[iInd_l].GetValue('JTF_FONCTABREGE'), 1, 5) + '.';
      Inc(iInd_l);
   end;
   OBFonctions_l.Free;
end; }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 15/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{ function TOM_YYJURIDIQUE.ActionnaireAutorise : boolean;
begin
   result := (GetField('JUR_FORME') = 'SA')  or (GetField('JUR_FORME') = 'SADIR') or
             (GetField('JUR_FORME') = 'SAS') or (GetField('JUR_FORME') = 'SASU');
end; }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... : 17/02/2006
Description .. : fonction ADM en dur, n'est-ce pas assez moyen ?
Suite ........ : et les gérants, ne sont-ils pas les "administrateurs" de 
Suite ........ : certaines formes ? (auquel cas le bouton TBGERANT 
Suite ........ : suffirait)
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BTORGDIRECTION(Sender: TObject);
begin
   AGLLanceFiche('DP', 'LIENDETJURIDIQUE',
                 'ANL_GUIDPERDOS=' + sGuidperdos_c + ';' + 'ANL_FONCTION=ADM;ANL_TYPEDOS=STE',
                 '', sGuidperdos_c + ';ADM');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BTGERANT(Sender: TObject);
var
   sNomper_l : String;
begin
   sNomper_l := AGLLanceFiche('DP', 'LIENDETJURIDIQUE',
                              'ANL_GUIDPERDOS=' + sGuidPerDos_c + ';ANL_FONCTION=GRT;ANL_TYPEDOS=STE',
                              '', sGuidPerDos_c + ';GRT');

   if (sNomper_l <> '') then
   begin
      ModeEdition(DS);
      SetControlText('TJUR_REFGERANTDP', sNomper_l);
   end;
   GereGerance;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_BTCAC(Sender: TObject);
begin
     AGLLanceFiche('DP', 'LIENDETJURIDIQUE',
                   'ANL_GUIDPERDOS=' + sGuidperdos_c + ';ANL_FONCTION=CCT;ANL_TYPEDOS=STE',
                   '', sGuidperdos_c + ';CCT');
     // GereCac => en fait on se contente de rafraichir les noms :
     // CAC titulaire
     SetControlText('CCT', RecupNomAnnulien(sGuidperdos_c,'CCT') );
     // CAC suppléant
     SetControlText('CCS', RecupNomAnnulien(sGuidperdos_c,'CCS') );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_CBCAC(Sender: TObject);
begin
   // évite le passage lors du chargement
   if bCharge_c then
      exit;

   // si on décoche existence de CAC, faut supprimer les CAC
   if (GetCheckBoxState('JUR_CAC') <> cbChecked) and
      (sCAC_c = 'X') then
   begin
      bSuppCAC_c := false;

      if PGIAsk('Cette opération entraînera la suppression de tous' + #13#10 +
                'les liens correspondants lors de la validation.' + #13#10 +
                'Confirmez-vous?', TitreHalley) = mrYes then
         bSuppCAC_c := true;
      if bSuppCAC_c then
      begin
         SetControlText('CCT', '');
         SetControlText('CCS', '');
      end
      else
         SetControlText('JUR_CAC', 'X');
   end;
   // affichages
   GereCac;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.GereCac;
var
   OBCCT_l, OBCCS_l : TOB;
begin
   OBCCT_l := OBFonctions_c.FindFirst(['JTF_FONCTION'], ['CCT'], true);
   OBCCS_l := OBFonctions_c.FindFirst(['JTF_FONCTION'], ['CCS'], true);

   if (OBCCT_l = nil) and (OBCCS_l = nil) then
   begin
      SetControlVisible('GRPBCAC', false);
      AfficheOngletGerance ;//LM20060516
      exit;
   end;

   SetControlVisible('GRPBCAC', true);
   SetControlText('CCT', '');
   SetControlText('CCS', '');

   if (OBCCT_l.Detail.count = 0) and (OBCCS_l.Detail.count = 0) then
   begin
      if GetField('JUR_CAC') = 'X' then
      begin
         ModeEdition(DS);
         SetField('JUR_CAC', '-');
      end;
      sCAC_c := '-';
   end
   else
   begin
      if GetField('JUR_CAC') = '-' then
      begin
         ModeEdition(DS);
         SetField('JUR_CAC', 'X');
      end;
      sCAC_c := 'X';

       // CAC titulaire
      if OBCCT_l.Detail.count > 0 then
         SetControlText('CCT', OBCCT_l.Detail[0].GetString('ANL_NOMPER'));// RecupNomAnnulien(sGuidPerDos_c, 'CCT') );
       // CAC suppléant
      if OBCCS_l.Detail.count > 0 then
         SetControlText('CCS', OBCCS_l.Detail[0].GetString('ANL_NOMPER'));//RecupNomAnnulien(sGuidPerDos_c, 'CCS') );
   end;
   AfficheOngletGerance ;//LM20060516
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnElipsisClick_Gerant(Sender: TObject);
begin
   AfficheLien('TJUR_REFGERANTDP', 'GRT');
   GereGerance;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnElipsisClick_CC(Sender: TObject);
var
   sFonction_l : string;
begin
   // CCT, CCS
   sFonction_l := THEdit(Sender).Name;
   AfficheLien(sFonction_l, sFonction_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.AfficheLien(sZone_p, sFonction_p : String);
var
   sNomPer_l, sRange_l : string;
   sGuidPer_l   : string;
begin
   sNomPer_l := ''; // nom de la personne liée
   sGuidPer_l := ''; // guid personne

   sRange_l := 'ANL_GUIDPERDOS=' + sGuidPerDos_c + ';ANL_TYPEDOS=STE;ANL_FONCTION=' + sFonction_p;
   sRange_l := AGLLANCEFICHE('JUR', 'ANNULIEN_SEL', sRange_l, '', '');
   if (sRange_l <> '') then
   begin
      READTOKENST(sRange_l);
      READTOKENST(sRange_l);
      READTOKENST(sRange_l);
      sGuidPer_l := READTOKENST(sRange_l);
      SetControlText(sZone_p, RecupNomAnnulien(sGuidPer_l, sFonction_p) );
   end;

{   // Nb de liens de ce type
   iNb_l := NbEnreg('ANNULIEN', '*',
                    'WHERE ANL_GUIDPERDOS = "' + sGuidPerDos_c + '"' +
                    ' AND ANL_FONCTION = "' + sFonction_p + '"');

   // si un seul lien
   if (iNb_l=1) then
   begin
      // récup du lien
      OBAnnulien_l := TOB.Create('ANNULIEN', nil, -1);
      OBAnnulien_l.LoadDetailDBFromSQL('ANNULIEN',
                                       'SELECT ANL_NOMPER, ANL_GUIDPER ' +
                                       'FROM ANNULIEN ' +
                                       'WHERE ANL_GUIDPERDOS = "' + sGuidPerDos_c + '"' +
                                       '  AND ANL_FONCTION = "' + sFonction_p + '"');

      if OBAnnulien_l.Detail.Count > 0 then
      begin
         sNomPer_l := OBAnnulien_l.Detail[0].GetString('ANL_NOMPER');
         sGuidPer_l := OBAnnulien_l.Detail[0].GetString('ANL_GUIDPER');
      end;
      OBAnnulien_l.Free;

      // affiche directement sa fiche annuaire
      AGLLanceFiche('YY', 'ANNUAIRE', sGuidPer_l, sGuidPer_l,
                    'ACTION=MODIFICATION');
   end
   else
   begin
       // sinon la liste des liens pour faire un choix
       sNomPer_l := AGLLanceFiche('DP', 'LIENDETJURIDIQUE',
                                  'ANL_GUIDPERDOS=' + sGuidPerDos_c + ';ANL_FONCTION=' + sFonction_p + ';ANL_TYPEDOS=STE',
                                  '',
                                  sGuidPerDos_c + ';' + sFonction_p + ';JUR');

   end;

   // lien choisi
   if (sNomPer_l <> '') then
      SetControlText(sZone_p, sNomPer_l)
   else
      SetControlText(sZone_p, RecupNomAnnulien(sGuidPer_l, sFonction_p) );    }
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.GereEleveNomin;
var
   bActif_l : boolean;
begin
     bActif_l := (GetCheckBoxState('JUR_ELEVENOMIN') = cbChecked);
     SetControlEnabled('JUR_DATEELEVNOMIN', bActif_l);
     SetControlEnabled('JUR_NOUVNOMIN', bActif_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure TOM_YYJURIDIQUE.GereNbAdmin;
var
   bActif_l : boolean;
begin
     bActif_l := (GetField('JUR_TTNBMINADM') < 20);
     SetControlEnabled('JUR_ORGDIRECTION', bActif_l);
//     SetControlEnabled('BTORGDIRECTION', bActif_l);
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.GereGerance;
var
   OBGrt_l, OBTitre_l : TOB;
   sAssocies_l, sGerant_l, sCoGerant_l, sTypeGerance_l, sNatureGerance_l : string;
   iInd_l, iNbGer_l : integer;
   dPctBenef_l : double;
begin
   OBGrt_l := OBFonctions_c.FindFirst(['JTF_FONCTION'], ['GRT'], true);
   SetControlVisible('GBGERANCE', (OBGrt_l <> nil));
   if OBGrt_l = nil then
   begin
    	AfficheOngletGerance ;//LM20060516
    	Exit;
   end ;

   // Nature de gérance
   if OBGrt_l.Detail.Count = 0 then
   begin
      sNatureGerance_l := 'NON';
      sTypeGerance_l := 'NON';
   end
   else
   begin
     	iNbGer_l := 0;
      sGerant_l := '';
      sCoGerant_l := '';
      sAssocies_l := '';

      for iInd_l := 0 to OBGrt_l.Detail.Count -1 do
      begin
	      // Celui dont les dates de mandat sont valides : ANL_NOMDATE et ANL_EXPDATE
         if (OBGrt_l.Detail[iInd_l].GetDateTime('ANL_NOMDATE') <= Date) and
            ((OBGrt_l.Detail[iInd_l].GetDateTime('ANL_EXPDATE') = iDate1900) or
             (OBGrt_l.Detail[iInd_l].GetDateTime('ANL_EXPDATE') > Date)) then
         begin
         	Inc(iNbGer_l);

            if (iNbGer_l = 1) then
            	sGerant_l := OBGrt_l.Detail[iInd_l].GetString('ANL_NOMPER')
            else if (iNbGer_l = 2) then
            	sCoGerant_l := OBGrt_l.Detail[iInd_l].GetString('ANL_NOMPER');

            if sAssocies_l <> '' then
               sAssocies_l := sAssocies_l + ', ';
            sAssocies_l := sAssocies_l + '"' + OBGrt_l.Detail[iInd_l].GetString('ANL_GUIDPER') + '"';

         end;
      end;

      if sAssocies_l <> '' then
         sAssocies_l := '  AND ANL_GUIDPER IN (' + sAssocies_l + ')';

      if iNbGer_l = 1 then
         sNatureGerance_l := 'GER'
      else
         sNatureGerance_l := 'COG';

      // Recherche des gérants également associés
      OBTitre_l := TOB.Create('ANNULIEN', nil, -1);
      OBTitre_l.LoadDetailDBFromSQL('ANNULIEN',
                     'SELECT ANL_FONCTION, ANL_NOMPER, ANL_TTPCTVOIX ' +
                     'FROM ANNULIEN ' +
                     'INNER JOIN JUFONCTION ON JFT_TYPEDOS = ANL_TYPEDOS ' +
                     '  AND JFT_FORME = ANL_FORME ' +
                     'WHERE ANL_GUIDPERDOS = "' + sGuidPerDos_c + '" ' +
                     sAssocies_l +
                     '  AND ANL_NOORDRE = 1 AND ANL_TYPEDOS = "STE" ' +
                     '  AND ANL_FONCTION <> "GRT" AND JFT_TITRE = "X"');

      if OBTitre_l.Detail.Count = 0 then
      begin
         sTypeGerance_l := 'HAS';
      end
      else
      begin
         dPctBenef_l := 0;
         for iInd_l := 0 to OBTitre_l.Detail.Count - 1 do
         begin
            dPctBenef_l := dPctBenef_l +
                           (OBTitre_l.Detail[iInd_l].GetDouble('ANL_TTPCTVOIX') * 100);
         end;

         if dPctBenef_l >= 50 then
            sTypeGerance_l := 'MAJ'
         else if dPctBenef_l < 50 then
            sTypeGerance_l := 'MIN';
      end;
   end;

   OBTitre_l.free;

   if GetControl('TJUR_REFGERANTDP') <> nil then
	   SetControlText('TJUR_REFGERANTDP', sGerant_l);

   if GetControl('TJUR_REFCOGERANTDP') <> nil then
   begin
	   SetControlText('TJUR_REFCOGERANTDP', sCoGerant_l);
   end;

   if GetField('JUR_NATUREGERANCE') <> sNatureGerance_l then
   begin
      ModeEdition(DS);
      SetField('JUR_NATUREGERANCE', sNatureGerance_l);
   end;

   if GetField('JUR_TYPEGERANCE') <> sTypeGerance_l then
   begin
      ModeEdition(DS);
      SetField('JUR_TYPEGERANCE', sTypeGerance_l);
   end;

   AfficheOngletGerance ;//LM20060516
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/06/2006
Modifié le ... :   /  /
Description .. : Organe de direction
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.GereOrganeDir;
var
   bVisible_l : boolean;
begin
   bVisible_l := (OBFonctions_c.FindFirst(['JTF_FONCTION'], ['ADM'], true) <> nil) or
                 (OBFonctions_c.FindFirst(['JTF_FONCTION'], ['MDI'], true) <> nil) or
                 (OBFonctions_c.FindFirst(['JTF_FONCTION'], ['MCD'], true) <> nil);

   SetControlVisible('GBORGANEDIR', bVisible_l);

   if (OBFonctions_c.FindFirst(['JTF_FONCTION'], ['ADM'], true) <> nil) then
   begin
      if GetField('JUR_ORGDIRECTION') <> 'CA' then
      begin
         ModeEdition(DS);
         SetField('JUR_ORGDIRECTION', 'CA');
      end;
   end
   else if (OBFonctions_c.FindFirst(['JTF_FONCTION'], ['MDI'], true) <> nil) then
   begin
      if GetField('JUR_ORGDIRECTION') <> 'DIR' then
      begin
         ModeEdition(DS);
         SetField('JUR_ORGDIRECTION', 'DIR');
      end;
   end
   else if (OBFonctions_c.FindFirst(['JTF_FONCTION'], ['MCD'], true) <> nil) then
   begin
      if GetField('JUR_ORGDIRECTION') <> 'CD' then
      begin
         ModeEdition(DS);
         SetField('JUR_ORGDIRECTION', 'CD');
      end;
   end
   else
   begin
      if GetField('JUR_ORGDIRECTION') <> '' then
      begin
         ModeEdition(DS);
         SetField('JUR_ORGDIRECTION', '');
      end;
   end;

   // Demande KPMG du 22/02/2008
   {$IFDEF DP}
   if (V_PGI.RunFromLanceur) and (vh_dp.SeriaKPMG)  then
   begin
      SetControlEnabled('JUR_ORGDIRECTION', false);
      SetControlEnabled('JUR_CAC', false);
      SetControlEnabled('CCT', false);
      SetControlEnabled('CCS', false);
      SetControlEnabled('JUR_TYPEGERANCE', false);
      SetControlEnabled('JUR_NATUREGERANCE', false);
      SetControlEnabled('TJUR_REFGERANTDP', false);
      SetControlEnabled('TJUR_REFCOGERANTDP', false);
      SetControlEnabled('TSGERANCE', false);
   end
   else
   {$ENDIF}
   begin
      if bVisible_l then
         SetControlEnabled('JUR_ORGDIRECTION', (GetField('JUR_TTNBMINADM') < 20));
   end;
   AfficheOngletGerance ;//LM20060516
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYJURIDIQUE.LiensChargeEtAffiche;
var
  iFRow_l, iLRow_l, iInd_l : integer;
begin
   iFRow_l := THgrid(GetControl('LISTEFONCTION')).Row;
   iLRow_l := THGrid(GetControl('LISTELIEN')).Row;

   for iInd_l := 0 to OBFonctions_c.Detail.Count - 1 do
      OBFonctions_c.Detail[iInd_l].ClearDetail;
   OBFonctions_c.ClearDetail;

   ChargeFonctionsEtLiens(sGuidPerDos_c, sNoOrdre_c,
                           sTypeDos_c, sTypeFiche_c, GetField('JUR_FORME'),
                           OBFonctions_c, OBCurFonction_c, OBCurLien_c);
   AfficheFonctionsEtLiens(sTypeDos_c, 
                           THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                           OBFonctions_c, OBCurFonction_c, OBCurLien_c);
   TFFiche(Ecran).HMTrad.ResizeGridColumns(THGrid(GetControl('LISTEFONCTION')));

   if THGrid(GetControl('LISTEFONCTION')).RowCount <= iFRow_l then
      iFRow_l := THGrid(GetControl('LISTEFONCTION')).RowCount - 1;
   THgrid(GetControl('LISTEFONCTION')).Row := iFRow_l;
   OnClick_FListeFonction(TObject(GetControl('LISTEFONCTION')));

   if THGrid(GetControl('LISTELIEN')).RowCount <= iLRow_l then
      iLRow_l := THGrid(GetControl('LISTELIEN')).RowCount - 1;
   THGrid(GetControl('LISTELIEN')).Row := iLRow_l;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYJURIDIQUE.OnClick_FLISTEFONCTION(Sender : TObject);
begin
   AfficheCurLiens(THGrid(GetControl('LISTEFONCTION')), THGrid(GetControl('LISTELIEN')),
                   OBCurFonction_c, OBCurLien_c);
   TFFiche(Ecran).HMTrad.ResizeGridColumns(THGrid(GetControl('LISTELIEN')));
end;

procedure TOM_YYJURIDIQUE.AfficheOngletGerance ;//LM20060516
    function isVisible(groupBox:string):boolean ;
    var ctrl:TControl ;
    begin
      result:=true ;
      ctrl:=getControl(groupBox) ;
      if ctrl=nil then exit ;
      result:=ctrl.visible ;
    end ;

begin
  setControlVisible('TSGERANCE',   isVisible('GBORGANEDIR')
                                or isVisible('GRPBCAC')
                                or isVisible('GBGERANCE') );
end ;


procedure TOM_YYJURIDIQUE.initAutreTom ;
var i : integer ;
    nom: string ;
begin
  for i:=0 to Ecran.ComponentCount-1 do
  begin
    nom:= upperCase(Ecran.Components[i].name) ;
    if (champToNum(nom)>-1) and (ExtractPrefixe(nom)='ANN') then
      evt.rebranche (nom, 'OnExit', Ann_OnExit ) ;
  end ;
end ;



procedure TOM_YYJURIDIQUE.Ann_OnExit (sender : TObject) ; //LM20070516
var nom, st : string ;
begin
  if (sender=nil) or (ANN_=nil) then exit ;//LM20070704
  nom:=TControl(sender).name ;
  if champToNum(nom)=0 then exit ;

  // éxécute l'évènement ancêtre
//  evt.Exec(nom, 'OnExit', sender);
  ann_.OnClick(Sender);
  if inString(nom, ['ANN_RCSVILLE', 'ANN_RCSDATE', 'ANN_RCSGEST', 'ANN_RCS']) then
    rcs
  else if inString(nom, ['ANN_RMDEP', 'ANN_RM']) then
    rm
  else if inString(nom, ['ANN_SIREN']) then
  begin
    rcs ;
    rm ;
  end
  else if inString(nom, ['ANN_CODENAF', 'ANN_CODENAF2']) then
  begin
    st := GetControlText(TControl(sender).name) ; //code naf
    ChangeCodeNaf(self, TControl(sender), codeNafIni)  ;
  end ;

end ;

procedure TOM_YYJURIDIQUE.rcs;
var st : string ;
begin
   if Boolean(GetControlText('ANN_RCS') = 'X') then
      st := GetControlText('ANN_SIREN') + ' RCS ' + GetControlText('ANN_RCSVILLE')
   else
      st := '' ;
   setControlText('TANN_RCSNO1', st )
end ;

procedure TOM_YYJURIDIQUE.rm ;
var st : string ;
begin
   if Boolean(GetControlText('ANN_RM') = 'X') then
      st := GetControlText('ANN_SIREN') + ' RM ' + GetControlText('ANN_RMDEP')
   else
      st:='' ;
   SetControlText('TANN_RMNO', st);
end ;

procedure TOM_YYJURIDIQUE.ANN_RCSOnClick(sender:TObject) ; //LM20070516
begin
   GroupBoxEnabled(Ecran, getControl('GBRCS'), 'ANN_RCS', getControlText('ANN_RCS')='X') ;
   if bCharge_c then exit;
   ann_.OnClick(Sender);
end ;

procedure TOM_YYJURIDIQUE.ANN_RMOnClick(sender:TObject) ; //LM20070516
begin
   GroupBoxEnabled(Ecran, getControl('GBRM'), 'ANN_RM', getControlText('ANN_RM')='X') ;
   if bCharge_c then exit;
   ann_.OnClick(Sender);
end ;

procedure TOM_YYJURIDIQUE.VoirSCI ; //LM20070516 doublon uTomAnn
var b:boolean ;
begin
  b:=inString (getField('ANN_CODEINSEE'), ['6541','6542','6540']) ;
  setControlVisible('ANN_TYPESCI', b) ;
  setControlVisible('TANN_TYPESCI', b) ;
end ;

procedure TOM_YYJURIDIQUE.bFormeOnClick (sender:TObject) ;   //LM20070516
begin
//   evt.Exec('bForme', 'OnClick', sender);
   if ChangeFormeJuridique (self, 'JUR_GUIDPERDOS', @DS) then
      VoirSCI ;
end ;

procedure TOM_YYJURIDIQUE.ANN_CODENAFOnEnter (sender:TObject) ;   //LM20070516
begin
  if sender = nil then exit ;
  codeNafIni := GetControlText(TControl(sender).name) ;
end ;

procedure TOM_YYJURIDIQUE.ElipsisPerAssClick(Sender: TObject);  //LM20070516
begin
  if sender=nil then exit ;
  cElipsisPerAssClick(Self, Sender) ;
  ann_.OnClick(sender);
end ;

procedure TOM_YYJURIDIQUE.JUR_LOCAGERANCE_OnClick(Sender: TObject); //de TomOrga ... LM20070516
var res: Integer;
    s : string ;
begin
  // si on décoche alors qu'il y a un propriétaire
  if (GetCheckBoxState('JUR_LOCAGERANCE') <> cbChecked) then
  begin
    res := 0;
    if RecupNomGerant(GetField('JUR_GUIDPERDOS'), s)<>'' then
      res := SupprimeAttach(GetField('JUR_GUIDPERDOS'), 'PFC', '');
    // si on n'a pas pu le supprimer, on recoche
    if res <> 0 then
      SetControlChecked('JUR_LOCAGERANCE', True);
  end;
  // affichages
  GereLocagerance;
end;

procedure TOM_YYJURIDIQUE.GereLocagerance; //de TomOrga ... LM20070516
var actif: Boolean;
    s:string ;
begin
  actif := (GetCheckBoxState('JUR_LOCAGERANCE')=cbChecked);
  SetControlEnabled('BLIENPROPRIETAIRE1', actif);
  SetControlEnabled('TNOMPROPRETAIRE', actif);
  SetControlEnabled('NOMPROPRIETAIRE', actif);
  SetControlText('NOMPROPRIETAIRE', RecupNomGerant(GetField('JUR_GUIDPERDOS'), s));
end;


procedure TOM_YYJURIDIQUE.BLIENPROPRIETAIRE1_OnClick(Sender: TObject); //de TomOrga ... LM20070516
var s :string ;
begin
  // on remplit aussi param Lequel pour remplir les critères du mul
  AGLLanceFiche('DP', 'LIENINTERPROP', 'ANL_GUIDPERDOS=' + GetField('JUR_GUIDPERDOS') + ';ANL_FONCTION=PFC',
                '', GetField('JUR_GUIDPERDOS') + ';PFC');
  SetControlText('NOMPROPRIETAIRE', RecupNomGerant(GetField('JUR_GUIDPERDOS'), s ));
end;

procedure TOM_YYJURIDIQUE.ElipsisPropriClick(Sender: TObject); //de TomOrga ... LM20070516
var ValChamp, RefGerant : String;
    Nb       : Integer;
begin
  // Inactif si on n'a pas les droits "voir les liens"
  if not JaiLeDroitConceptBureau(ccVoirLesLiens) then exit;
  
  ModeEdition(DS);
  Nb := NbEnreg('ANNULIEN', '*',
                'where ANL_GUIDPERDOS="' + GetField('JUR_GUIDPERDOS')+'"' +' and ANL_FONCTION="PFC"');
  ValChamp := RecupNomGerant(GetField('JUR_GUIDPERDOS'), RefGerant);
  if (ValChamp = '') or (Nb > 1) then
    begin
    // plusieurs propriétaires => affiche la liste pour en choisir un à afficher
    ValChamp := AGLLanceFiche('DP','LIENINTERPROP',
                              'ANL_GUIDPERDOS=' + GetField('JUR_GUIDPERDOS') + ';ANL_FONCTION=PFC','', GetField('JUR_GUIDPERDOS'));
    if (ValChamp = '') then ValChamp := RecupNomGerant(GetField('JUR_GUIDPERDOS'), RefGerant);
    end
  else
    // 1 seul prop, affiche directement sa fiche
    AGLLanceFiche('YY', 'ANNUAIRE', RefGerant, RefGerant,'ACTION=MODIFICATION')  ;

  if (ValChamp <> '') then
    SetControlText('NOMPROPRIETAIRE', ValChamp);
end;


Initialization
   registerclasses([TOM_YYJURIDIQUE]) ;
end.
