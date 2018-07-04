{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 06/12/2002
Modifié le ... : 18/11/2003
Description .. : TOF de base pour les multi critères Evenement
Suite ........ : CECI EST UNE TOF ANCETRE
Suite ........ : NE PAS METTRE DE SOURCE SPECIFIQUE A UNE
Suite ........ : APPLICATION PRECISE OU ME DEMANDER
Suite ........ : SINON CA POSE PROBLEME DANS LA TOF HERITANTE
Suite ........ : (BEATRICE MERIAUX)
Mots clefs ... :
*****************************************************************}

unit UTOFEvenement;

interface

uses
{$IFDEF EAGLCLIENT}
   MaineAGL, eMul, utob,
{$ELSE}
   {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, Mul,
{$ENDIF}
   hqry, UTOF, UCLASSMyMenu, UCLASSMySubMenu, DPTofAnnuSel,
   hmsgbox, HTB97, Menus, HCtrls, HEnt1, HDB, sysutils, AGLInit, Classes, controls;

{$IFDEF EAGLCLIENT}
//type THMyGrid = THGrid;
{$ELSE}
//type THMyGrid = THDBGrid;
{$ENDIF}

//////////// DECLARATION //////////////
const
  cstFiltre_g : string = 'XX_WHERE0';

Type
   TOF_Evenement = Class (TOF)
      public
         QRYMul_c : THQuery;
         sCodeDos_c, sGuidPer_c : string;
         bClose_c : boolean;

         // Pour savoir si on charge tel ou tel menu click droit
         procedure OnArgument(Arguments : String ) ; override ;
         procedure OnClose; override ;

      protected

         // Boutons
         procedure BINSERT_OnClick(Sender: TObject); virtual;
         procedure BDELETE_OnClick(Sender: TObject); virtual;
         procedure BDUPLIQUER_OnClick(Sender: TObject); virtual;
         procedure BOUTLOOK_OnClick(Sender: TObject); virtual;
         procedure BNouvRech_OnClick(Sender: TObject); virtual;
         // Grille
         procedure FLISTE_OnDblClick(Sender: TObject); virtual;
         // Critères
         procedure JUR_NOMDOS_OnElipsisClick(Sender: TObject); virtual;
         procedure JEV_CODEOP_OnElipsisClick(Sender: TObject); virtual;
         procedure ANN_NOMPER_OnElipsisClick(Sender: TObject); virtual;
         procedure JUR_NOMDOS_OnExit(Sender: TObject); virtual;
         procedure JEV_CODEOP_OnExit(Sender: TObject); virtual;
         procedure ANN_NOMPER_OnExit(Sender: TObject); virtual;

         Procedure InitClickDroit( bChargeOutlook_p : boolean = false;
                              bChargeTaches_p : boolean = false; bChargeEtats_p : boolean = false);
         procedure FiltreCombo(sControl_p, sRestriction_p : string);
         procedure FiltreEnreg(sRestriction_p : string);
         procedure FiltreMenuBouton(sRestriction_p : string);
         procedure OuvreFicheEvt(sGuidEvt_p, sFamEvt_p : string);

         procedure OnClickEvtTaches(Sender : TObject);
         procedure OnClickEvtEtats(Sender : TObject);
         procedure OnClickEvtOpe(Sender : TObject);
         procedure OnPopup_Action(Sender : TObject);

      private
         FMenuNewEvt_c : TMyMenu;     // Menu sur bouton NEW
         FMenuEvt_c : TMySubMenu;           // Click droit : Sous menus tâches, états et messagerie
         sFamEvtMul_c : string;              // Famille d'évènements associé au multi critère courant
         
   end;

//////////// IMPLEMENTATION //////////////

implementation

uses AGLInitDpJur, DpJurOutils, DpJurOutilsEve;

{*****************************************************************
Auteur ....... : BM
Date ......... : 06/12/02
Procédure .... : OnArgument
Description .. : Initialisation des éléments, traitement de la confidentialité, menus clic droit
Paramètres ... : Mode fiche (DOS) pour le DP, JEV_CODEDOS pour juri...
*****************************************************************}
procedure TOF_Evenement.OnArgument(Arguments : String ) ;
var
   nPosFam_l : integer;
begin
   inherited;
   bClose_c := false;
   QRYMul_c := TFMul(Ecran).Q;
   Ecran.HelpContext := 851100100;

   // Famille d'évènement depuis le nom de la fiche "YYEVEN_DOC_MUL", "EVEN_DOC_MUL" ou "EVENEMENT_MUL"
   sFamEvtMul_c := Ecran.name;
   READTOKENPipe( sFamEvtMul_c, '_');    // 1er "_" avant la famille
   nPosFam_l := Pos('_', sFamEvtMul_c);
   if ( nPosFam_l > 0 ) then
      // 2ème "_" après la famille, si "YYEVEN_DOC_MUL", "EVEN_DOC_MUL"
      sFamEvtMul_c := Copy( sFamEvtMul_c, 1, nPosFam_l - 1 )
   else
      // Si "EVENEMENT_MUL"
      sFamEvtMul_c := '';
   //
   SetControlVisible( 'PCOMPLEMENT', MontreOnglet( TFMul(Ecran), 'PCOMPLEMENT' ));

{$IFDEF EAGLCLIENT}
   THGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblClick;
{$ELSE}
   THDBGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblClick;
{$ENDIF}
   // Boutons
   if GetControl('BNEWEVT') <> nil then
   begin
      // Bouton nouveau avec menu déroulant
      //TToolbarButton97(GetControl('BNEWEVT')).OnClick := BINSERT_OnClick;
      // Menu déroulant
      FMenuNewEvt_c := TMyMenu.Create(TPopUpMenu(GetControl('POPUPEVT')),
                                      'SELECT CC_CODE, CC_LIBELLE ' +
                                      'FROM CHOIXCOD ' +
                                      'WHERE CC_TYPE = "' + TTToTipe('JUFAMEVT') + '" ' +
                                       ' AND CC_ABREGE <> ""',
                                       'CC_LIBELLE', 'CC_CODE',
                                       BINSERT_OnClick);

      SetControlVisible( 'BNEWEVT', true );
      SetControlVisible( 'BINSERT', false );
   end
   else
      // Bouton nouveau classique
      TToolbarButton97(GetControl('BINSERT')).OnClick := BINSERT_OnClick;

   if GetControl('BOUTLOOK') <> nil then
      TToolbarButton97(GetControl('BOUTLOOK')).OnClick := BOUTLOOK_OnClick;

   TToolbarButton97(GetControl('BDELETE')).OnClick := BDELETE_OnClick;
   TMenuItem(GetControl('BNOUVRECH')).OnClick := BNouvRech_OnClick;
   if GetControl('BDUPLIQUER')<>Nil then
     TMenuItem(GetControl('BDUPLIQUER')).OnClick := BDUPLIQUER_OnClick;

   // Champs
   if (GetControl('JUR_NOMDOS' ) <> nil) then
   begin
      THEdit(GetControl('JUR_NOMDOS')).OnElipsisClick := JUR_NOMDOS_OnElipsisClick;
      THEdit(GetControl('JUR_NOMDOS')).OnExit := JUR_NOMDOS_OnExit;
   end;

   if (GetControl('JEV_CODEOP' ) <> nil) then
   begin
      THEdit(GetControl('JEV_CODEOP')).OnElipsisClick := JEV_CODEOP_OnElipsisClick;
      THEdit(GetControl('JEV_CODEOP')).OnExit := JEV_CODEOP_OnExit;
   end;

   if (GetControl('ANN_NOMPER' ) <> nil) then
   begin
      THEdit(GetControl('ANN_NOMPER')).OnElipsisClick := ANN_NOMPER_OnElipsisClick;
      THEdit(GetControl('ANN_NOMPER')).OnExit := ANN_NOMPER_OnExit;
   end;

   // Filtrage des combos
   if (GetControl('JEV_FAMEVT' ) <> nil) then
   begin
      if sFamEvtMul_c <> '' then
      begin
         SetControlProperty('JEV_FAMEVT', 'Plus', 'AND CC_CODE = "' + sFamEvtMul_c + '"');
         SetControlText('JEV_FAMEVT', sFamEvtMul_c);
         SetControlEnabled('JEV_FAMEVT', false);
      end
      else
      begin
         SetControlProperty('JEV_FAMEVT', 'Plus', 'AND (CC_ABREGE <> "")');
      end;
   end;
   // NE PAS CODER JTE_DOMAINEACT = "" (SPECIFIQUE DP??) ICI MAIS DANS LA TOF HERITANTE
   // SINON PROBLEME EN HERITAGE PAR JURI CAR JTE_DOMAINEACT DOIT Y ETRE A "JUR"
   if (GetControl('JEV_CODEEVT' ) <> nil) then
   begin
      if sFamEvtMul_c <> '' then       // Filtre sur UNE SEULE famille
         SetControlProperty('JEV_CODEEVT', 'Plus', 'JTE_FAMEVT = "' + sFamEvtMul_c + '"')
      else
         SetControlProperty('JEV_CODEEVT', 'Plus', 'JTE_FAMEVT in (SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "JFE" and CC_ABREGE <> "")');
   end;

   // Filtrage des enregistrements
   with THDBEdit.Create( Ecran ) do
   begin
      Name := cstFiltre_g;
      Text := 'JEV_FAMEVT in (SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "JFE" and CC_ABREGE <> "")';
      Parent := TWinControl( GetControl('PCritere' ) );
      Visible := false;
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 25/02/2003
Modifié le ... :   /  /    
Description .. : Initialisation du menu sur click droit
Mots clefs ... : 
*****************************************************************}
procedure TOF_Evenement.InitClickDroit( bChargeOutlook_p : boolean = false;
                        bChargeTaches_p : boolean = false; bChargeEtats_p : boolean = false);
var
   sCaption_l, sReqEtat_l, sReqFait_l : string;
   iInd_l : integer;
begin
   if bChargeTaches_p then
   begin
      sCaption_l := 'JEV_ETATDOC';
   end;

   if bChargeEtats_p then
   begin
      if sCaption_l <> '' then
         sCaption_l := sCaption_l + ';';
      sCaption_l := sCaption_l + 'JEV_FAIT';
   end;

   FMenuEvt_c := TMySubMenu.Create(TPopUpMenu(GetControl('POPUP')),
                                TFmul(Ecran), sCaption_l, true);

   iInd_l := FMenuEvt_c.GetSubIndice;
   if bChargeTaches_p then
   begin

      sReqEtat_l := 'select CO_CODE, CO_ABREGE from COMMUN ' +
                    'WHERE CO_TYPE = "' + TTToTipe('JUETATDOC') + '"' ;
      FMenuEvt_c.SubMenuAdd(iInd_l, sReqEtat_l, 'CO_ABREGE', 'CO_CODE', OnClickEvtTaches);
      Inc(iInd_l);
   end;

   if bChargeEtats_p then
   begin
      sReqFait_l := 'select CO_LIBRE, CO_ABREGE from COMMUN ' +
                    'WHERE CO_TYPE = "' + TTToTipe('JUFAIT') + '"';
      FMenuEvt_c.SubMenuAdd(iInd_l, sReqFait_l, 'CO_ABREGE', 'CO_LIBRE', OnClickEvtEtats);
      Inc(iInd_l);
   end;

   if (Ecran.Name = 'EVEN_DOC_MUL') then
   begin
      TPopupMenu(GetControl('POPUP')).OnPopup := OnPopup_Action;
      TMenuItem(GetControl('PMDOC')).OnClick := FLISTE_OnDblClick;
      TMenuItem(GetControl('PMOPE')).OnClick := OnClickEvtOpe;
   end;

end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : OnClickEvtTache
Description .. : Procédure exécutée lors du click sur les éléments
                 du sous menu "Tâches"
Paramètres ... : L'objet
*****************************************************************}

procedure TOF_Evenement.OnClickEvtTaches(Sender : TObject);
var
   sValeur_l, sData_l : string;
   nRow_l : integer;
begin
   sValeur_l := FMenuEvt_c.GetHint;
   sData_l := 'JUEVENEMENT_MUL|';                           // 0 : Table
   sData_l := sData_l + 'JEV_ETATDOC=' + sValeur_l + '|';   // 1 : Champ à mettre à jour et sa valeur
   sData_l := sData_l + 'JEV_GUIDEVT';                        // 2 : Clé de l'enregistrement

   nRow_l := GrilleParcours(TFmul(Ecran),
                            sData_l, 3,
                            @InitSetValChamp, @AGLSetValChamp);
   if nRow_l > 0 then
   begin
      PGIInfo( IntToStr(nRow_l) + ' éléments modifiés', FMenuEvt_c.GetSubCaption);
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 18/09/02
Procédure .... : OnClickEvtEtat
Description .. : Procédure exécutée lors du click sur les éléments
                 du sous menu "Etat"
Paramètres ... :
*****************************************************************}

procedure TOF_Evenement.OnClickEvtEtats(Sender : TObject);
var
   sValeur_l, sData_l : string;
   nRow_l : integer;
begin
   sValeur_l := FMenuEvt_c.GetHint;
   sData_l := 'JUEVENEMENT_MUL|';                        // 0 : Table
   sData_l := sData_l + 'JEV_FAIT=' + sValeur_l + '|';   // 1 : Champ à mettre à jour et sa valeur
   sData_l := sData_l + 'JEV_GUIDEVT';                     // 2 : Clé de l'enregistrement

   nRow_l := GrilleParcours(TFmul(Ecran),
                            sData_l, 3,
                            @InitSetValChamp, @AGLSetValChamp );
   if nRow_l > 0 then
   begin
      PGIInfo( IntToStr(nRow_l) + ' éléments modifiés', FMenuEvt_c.GetSubCaption);
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 15/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Evenement.OnPopup_Action(Sender : TObject);
var
   sCodeOp_l : string;
begin
   {$IFDEF EAGLCLIENT}
   GetDataSet;
   {$ENDIF}
   sCodeOp_l := TFMul(Ecran).Q.FindField('JEV_CODEOP').AsString;
   TMenuItem(GetControl('PMOPE')).Enabled := (sCodeOp_l <> '');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 15/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Evenement.OnClickEvtOpe(Sender : TObject);
var
   sCodeOp_l, sCodeDos_l : string;
begin
   {$IFDEF EAGLCLIENT}
   GetDataSet; // repositionne le query (semble-t'il) sur l'enreg sélectionné dans la liste !
   {$ENDIF}

   sCodeOp_l  := TFMul(Ecran).Q.FindField('JEV_CODEOP').AsString;
   sCodeDos_l := TFMul(Ecran).Q.FindField('JEV_CODEDOS').AsString;

   AGLLanceFiche('JUR', 'OPEDOSSIER', '', sCodeOp_l + ';' + sCodeDos_l,
                 'ACTION=MODIFICATION;' + sCodeOp_l + ';' + sCodeDos_l + ';');

   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Evenement.FiltreCombo(sControl_p, sRestriction_p : string);
var
   sPlus_l : string;
begin
   sPlus_l := THValComboBox(GetControl(sControl_p)).Plus;
   if sPlus_l <> '' then
      sPlus_l := sPlus_l + ' AND ';
   sPlus_l := sPlus_l + '(' + sRestriction_p + ')';
   SetControlProperty(sControl_p, 'Plus', sPlus_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Evenement.FiltreEnreg(sRestriction_p : string);
var
   sFiltre_l : string;
begin
   sFiltre_l := GetControlText(cstFiltre_g);
   if sFiltre_l <> '' then
      sFiltre_l := sFiltre_l + ' AND ';
   sFiltre_l := sFiltre_l + '(' + sRestriction_p + ')';
   SetControlText(cstFiltre_g, sFiltre_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_Evenement.FiltreMenuBouton(sRestriction_p : string);
begin
   if FMenuNewEvt_c <> nil then
      FMenuNewEvt_c.OnMenuFiltre(sRestriction_p);
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 5/12/02
Procédure .... : OnClose
Description .. : A la fermeture de la fenêtre, détruit les menus clic droit
Paramètres ... :
*****************************************************************}
procedure TOF_Evenement.OnClose;
begin
   inherited;
   bClose_c := true;
   if FMenuEvt_c <> nil then
      freeandnil(FMenuEvt_c);//OnMenuDetruit ;

   if FMenuNewEvt_c <> nil then
      freeandnil(FMenuNewEvt_c);//OnMenuDetruit;
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : FLISTE_OnDblClick
Description .. : Ouvre une fiche Evènement(différente selon DP/Juri).
                 Lancé aussi par la mouette verte.
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.FLISTE_OnDblClick(Sender: TObject);
var
   sFamEvt_l, sGuidEvt_l : string;
begin
   {$IFDEF EAGLCLIENT}
   GetDataSet; // repositionne le query (semble-t'il) sur l'enreg sélectionné dans la liste !
   {$ENDIF}
   sGuidEvt_l := QRYMul_c.FindField('JEV_GUIDEVT').AsString;
   sFamEvt_l := QRYMul_c.FindField('JEV_FAMEVT').AsString;

   OuvreFicheEvt(sGuidEvt_l, sFamEvt_l);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_Evenement.OuvreFicheEvt(sGuidEvt_p, sFamEvt_p : string);
var
   sEcran_l, sCle_l : string;
begin
   if sFamEvt_p = '' then
   begin
      PGIInfo('Aucune famille d''évènement définie pour cet évènement.', Ecran.Caption);
      exit;
   end;
   sEcran_l := 'YYEVEN_' + sFamEvt_p;
   sCle_l := 'ACTION=MODIFICATION;;;' + sFamEvt_p;

   AGLLanceFiche( 'YY', sEcran_l, sGuidEvt_p, sGuidEvt_p, sCle_l );
   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);

end;
{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : BINSERT_OnClick
Description .. : Création d'un nouvel évènement.
                 Ouvre une fiche Evènement(différente selon DP/Juri).
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.BINSERT_OnClick(Sender: TObject);
var
   sEcran_l, sCle_l, sFamEvtMul_l : string;
begin
   if FMenuNewEvt_c <> nil then
      sFamEvtMul_l := FMenuNewEvt_c.GetHint
   else
      sFamEvtMul_l := sFamEvtMul_c;

   if sFamEvtMul_l = '' then exit;

   sEcran_l := 'YYEVEN_' + sFamEvtMul_l;

   sCle_l := 'ACTION=CREATION;' + GetControlText('JEV_CODEDOS') + ';' +
             GetControlText('JEV_GUIDPER') + ';' + sFamEvtMul_l;

   AGLLanceFiche( 'YY', sEcran_l, '', '', sCle_l );
   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : BDELETE_OnClick
Description .. : Suppression d'une sélection d'enreg dans le mul
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.BDELETE_OnClick(Sender: TObject);
begin
  SupprimeListeEnreg(TFMul(Ecran), 'JUEVENEMENT', true);
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : BDUPLIQUER_OnClick
Description .. : Duplique un évènement
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.BDUPLIQUER_OnClick(Sender: TObject);
var
   sNewGuidEvt_l, sFamEvt_l : string;
begin
   sFamEvt_l := QRYMul_c.FindField('JEV_FAMEVT').AsString;
   sNewGuidEvt_l := DupliquerElement('JUEVENEMENT', QRYMul_c.FindField('JEV_GUIDEVT').AsString, '');
//   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
   OuvreFicheEvt(sNewGuidEvt_l, sFamEvt_l);
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : JUR_NOMDOS_OnElipsisClick
Description .. : Sélection d'un dossier
Paramètres ... : L'obejt
*****************************************************************}
procedure TOF_Evenement.JUR_NOMDOS_OnElipsisClick(Sender: TObject);
var
   nouveau: String;
begin
   nouveau := AGLLanceFiche('JUR','JURIDIQUE_SEL',GetControlText('JEV_CODEDOS'),'',';');
   if nouveau  <> '' then
   begin
      SetControlText('JEV_CODEDOS', nouveau);
      SetControlText('JUR_NOMDOS', RechDom('JURIDIQUE', nouveau, False));
      TFMul(Ecran).BChercheClick( Sender );
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 5/12/02
Procédure .... : JEV_CODEOP_OnElipsisClick
Description .. : Sélection d'une opération
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.JEV_CODEOP_OnElipsisClick(Sender: TObject);
var
   nouveau : String;
begin
   nouveau := AGLLanceFiche('JUR','DOSOPER_SEL',GetControlText('JEV_CODEOP'),'','');
   if nouveau <> '' then
   begin
      SetControlText('JEV_CODEOP', nouveau);
      SetControlText('JEV_CODEOP', RechDom('JUDOSOPER', nouveau, False));
      TFMul(Ecran).BChercheClick( Sender );
   end;
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : ANN_NOMPER_OnElipsisClick
Description .. : Sélection d'une personne
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.ANN_NOMPER_OnElipsisClick(Sender: TObject);
var
   nouveau : String;
begin
//   nouveau := AGLLanceFiche('YY','ANNUAIRE_SEL', '','','');
   nouveau := LancerAnnuSel ('','','');
//   nouveau := AGLLanceFiche('YY','ANNUAIRE_SEL',GetControlText('JEV_GUIDPER'),'','');
   if nouveau<>'' then
   begin
      //SetControlText('JEV_GUIDPER', nouveau);
      SetControlText('ANN_NOMPER', RechDom('ANNUAIRE', nouveau, False));
      TFMul(Ecran).BChercheClick( Sender );
   end;
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : JUR_NOMDOS_OnExit
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.JUR_NOMDOS_OnExit(Sender: TObject);
begin
   // Sinon problème dans JURI lorsqu'on quitte la fiche à cause du focus sur ce champ
   if bClose_c then exit;
// Recherche sur nom
   if GetControlText('JUR_NOMDOS') = '' then
      SetControlText('JEV_CODEDOS', sCodeDos_c);
   TFMul(Ecran).BChercheClick( Sender );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 5/12/02
Procédure .... : JEV_CODEOP_OnExit
Description .. : Validation de l'opération
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.JEV_CODEOP_OnExit(Sender: TObject);
begin
   TFMul(Ecran).BChercheClick( Sender );
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : ANN_NOMPER_OnExit
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.ANN_NOMPER_OnExit(Sender: TObject);
begin
   if GetControlText('ANN_NOMPER') = '' then
      SetControlText('JEV_GUIDPER', '');
   TFMul(Ecran).BChercheClick( Sender );
end;

{*****************************************************************
Auteur ....... : MD
Date ......... : ??/??/??
Procédure .... : BOUTLOOK_OnClick
Description .. :
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.BOUTLOOK_OnClick(Sender: TObject);
var
   sData_l : string;
   nRow_l : integer;
begin
   sData_l := 'JEV_FAMEVT;JEV_GUIDEVT'; // 0 famille évènement, 1 n° évènement
   nRow_l := GrilleParcours(TFMul(Ecran),
                            sData_l, 2,
                            @InitAjoutOutlook, @AGLAjoutOutlook);
   if nRow_l > 0 then
      PGIInfo( IntToStr(nRow_l) + ' éléments ajoutés', 'Liaison Outlook' );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 16/12/02
Procédure .... : BNouvRech_OnClick
Description .. : RAZ pour nouvelle recherche
Paramètres ... : L'objet
*****************************************************************}
procedure TOF_Evenement.BNouvRech_OnClick(Sender: TObject);
begin
   SetControlText( 'JEV_CODEDOS', '' );
   SetControlText( 'JEV_CODEOP', '' );
   SetControlText( 'JEV_GUIDPER', '' );

   SetControlText( 'JUR_NOMDOS', '' );
   SetControlText( 'JEV_CODEOP', '' );
   SetControlText( 'ANN_NOMPER', '' );
   TFMul(Ecran).BNouvRechClick(Sender);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure TOF_Evenement.OnFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
   Point : TPoint;
begin
   if (Key = VK_F10) then  // Enregistrement ou validation
   begin
      Key := 0;
      if GetControl('BVALIDER') <> nil then
         TToolbarButton97(GetControl('BVALIDER')).OnClick(nil);
   end
   else if (Key = VK_F11) then  // Menu contextuel
   begin
      Key := 0;
      if THDBGrid(GetControl('FLISTE')).PopupMenu <> nil then
      begin
         Point.x := THDBGrid(GetControl('FLISTE')).Left +
                    round((THDBGrid(GetControl('FLISTE')).Width) / 2);
         Point.y := THDBGrid(GetControl('FLISTE')).Top +
                    (THDBGrid(GetControl('FLISTE')).Row + 1) *
                    (THDBGrid(GetControl('FLISTE')).RowHeights[1]);


         ClientToScreen(Ecran.Handle, Point);
         THDBGrid(GetControl('FLISTE')).PopupMenu.Popup(Point.x, Point.y);
      end;
   end
   else if (Key = vk_delete) and (Shift = [ssCtrl]) then  // Ctrl + suppr : supprimer
   begin
      Key := 0;
      if GetControl('BSUPPRIMER') <> nil then
         TToolbarButton97(GetControl('BSUPPRIMER')).OnClick(nil);
   end
   else if (Key = vk_nouveau) and (Shift = [ssCtrl]) then  // Ctrl + N : nouveau
   begin
      Key := 0;
      if GetControl('BNEWEVT') <> nil then
      begin
         Point.x := THDBGrid(GetControl('FLISTE')).Left;
         ClientToScreen(Ecran.Handle, Point);

         Point.x := TToolbarButton97(GetControl('BNEWEVT')).Left - Point.x;
         Point.y := TToolbarButton97(GetControl('BNEWEVT')).Top +
                    TToolbarButton97(GetControl('BNEWEVT')).Height;
         Point := TFMul(Ecran).ScreenToClient(TToolbarButton97(GetControl('BNEWEVT')).ClientToScreen(Point));
         ClientToScreen(TFMul(Ecran).Handle, Point);
         TToolbarButton97(GetControl('BNEWEVT')).DropdownMenu.Popup(Point.X, Point.Y);
      end
      else
         TToolbarButton97(GetControl('BINSERT')).OnClick(nil);
   end
   else if (Key = vk_imprime)  and (Shift = [ssCtrl]) then  // Ctrl + P : imprimer
   begin
      Key := 0;
      TToolbarButton97(GetControl('BIMPRIMER')).OnClick(nil);
   end
   else
      TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
end;}

Initialization
   registerclasses([TOF_Evenement]) ;
end.
