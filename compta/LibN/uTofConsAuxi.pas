{***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 15/05/2002
Modifié le ... : 23/08/2005 - GCO - FQ 11950 et 16467
Description .. : Source TOF de la FICHE : CPCONSGENE ()
Mots clefs ... : TOF;CPCONSGENE
*****************************************************************}
Unit uTofConsAuxi ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
{$ELSE}
     db,
     dbGrids,
  {$IFDEF DBXPRESS}
     uDbxDataSet,
  {$ELSE}
     dbtables,
  {$ENDIF}
     Fe_Main,         // AGLLanceFiche
     Mul,
{$ENDIF}
     HDB,
     Forms,
     SysUtils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     uTob,
     HQry,             // HQuery
     Htb97,
     HSysMenu,
     Menus,            // PopUpMenu
     ent1,             // VH^.
     Filtre,
     Windows,          // VK_
     Grids,            // TGridDrawState
     Graphics,
     AGLInit,          // TheData
     HRichOle,         // THRichEditOle
     CalcOle,          // TabloExt
     LookUp,           // LookUpList
{$IFDEF MODENT1}
     CPTypeCons,
{$ENDIF MODENT1}

{$IFDEF COMPTA}
     CritEdt,          // TCritEdt
     ComObj,           // CreateOleObject
{$ENDIF}
     UTOFViergeMul;

Type

  TOF_CPCONSAUXI = Class (TOF_ViergeMul)

    TabTablesLibres  : TTabSheet;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure AfterShow                ; override;
    procedure OnKeyDownEcran   ( Sender : TObject; var Key: Word; Shift: TShiftState); override ;
    procedure OnPopUpPopF11    ( Sender : TObject ); override ;

    private
      FToutAccorde          : Boolean; // Pour déterminer les droits de modif du USER
      FBoCriterePonderation : Boolean; // Doit on prendre en compte les critères de la Pondéeration
      FBoCritereVariation   : Boolean; // Doit on prendre en compte les critères de la Variation

      FAction            : TActionFiche;

      FStSelect          : string;
      FStArgumentTOF     : string;

      FColDebitE         : integer;
      FColCreditE        : integer;
      FColDebitS         : integer;
      FColCreditS        : integer;
      FColSoldeP         : integer;
      FColSoldeE         : integer;
      FColVariation      : integer;
      FColPourcentage    : integer;

      FValeurPonderation : Double;  //
      FValeurVariation   : Double;  //

      T_Auxiliaire       : THEdit;
      T_Auxiliaire_      : THEdit;
      T_Collectif        : THEdit;
      T_DateDernMvt      : THEdit;
      T_DateDernMvt_     : THEdit;

      TCorresp1          : THMultiValComboBox;
      TModeSelection     : THValComboBox;
      T_NatureAuxi       : THMultiValComboBox;
      TSensReel          : THValComboBox;

      // Révision
      TSituationCompte   : THValComboBox;
      TTypePonderation   : THValComboBox;
      ZCPonderation      : THValComboBox;
      TValeurPonderation : THNumEdit;

      TTypeVariation     : THValComboBox;
      ZCVariation        : THValComboBox;
      TValeurVariation   : THNumEdit;
      TNatureVariation   : THValComboBox;
      BEffaceRevision    : TToolBarButton97;
      // Fin Révision

      BBlocNote          : TToolBarButton97;

      PopUpComptable     : TPopUpMenu;
      PopUpSpecifique    : TPopUpMenu;
      PopUpAideRevision  : TPopUpMenu;
      PopUpRevInteg      : TPopUpMenu;
      PopUpParametre     : TPopUpMenu;
      PopUpEdition       : TPopUpMenu;

      FStPrefixeE : string;
      FStPrefixeP : string;
      FStCodeExoE : string ;
      FStCodeExoP : string ;

      FExoDate : TExoDate;

      FBoJustifSolde : Boolean; // Ouverture du CPCONSECR en justif de solde
      FBoOkRevision  : Boolean; // Gestion de la révision

      procedure IndiceColonne;
      procedure TransactionDelettreMvtExoRef;
      procedure TransactionDeLettrageTotal;
      procedure MiseAJourCaptionEcran;
      procedure OnExitT_Auxiliaire        ( Sender : TObject );
      procedure OnElipsisClickT_Auxiliaire( Sender : TObject );

      // Initalisation des PopUp
      procedure InitAllPopUp            ( vActivation : Boolean );
      procedure OnPopUpComptable        ( Sender : TObject );
      procedure OnPopUpSpecifique       ( Sender : TObject );
      procedure OnPopUpAideRevision     ( Sender : TObject );
      procedure OnPopUpRevInteg         ( Sender : TObject );
      procedure OnPopUpParametre        ( Sender : TObject );
      procedure OnPopUpEdition          ( Sender : TObject );

      procedure InitPopUpComptable      ( vActivation : Boolean ); //
      procedure InitPopUpSpecifique     ( vActivation : Boolean ); //
      procedure InitPopUpAideRevision   ( vActivation : Boolean ); //
      procedure InitPopUpRevInteg       ( vActivation : Boolean ); //
      procedure InitPopUpParametre      ( vActivation : Boolean ); //
      procedure InitPopUpEdition        ( vActivation : Boolean ); //

      procedure GetCellCanvasFListe     ( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

      procedure RecupCritereRevision;
      procedure TripotageTSituationCompte;
      function  RecupAutreWhere           : string;
      function  RecupWhereSituationCompte : string;
      function  CompareOperateur( vValeur, vValeurAComparer : Double; vStOperateur : string ) : Boolean;

      // Révision avancée
      procedure OnClickBEffaceRevision         ( Sender : TObject );
      procedure OnChangeTNatureVariation       ( Sender : TObject );
      // Fin Révision avancée

      // Outils specifiques
      procedure OnClickDetailCompte            ( Sender : TObject );
      procedure OnClickDetailCompte2           ( Sender : TObject );

      procedure LettrageMClick         ( Sender : TObject ); // Lettrage manuel
      procedure LettrageAClick         ( Sender : TObject ); // Lettrage automatique
      procedure DelettreMvtExoRef      ( Sender : TObject ); // Delettrage exercice de référence
      procedure DelettreTotal          ( Sender : TObject ); // Delettrage Total du Compte

      // Traitements Comptables
      procedure SaisieEcrClick         ( Sender : TObject ); // Saisie d'une Piece
      procedure SaisieBorClick         ( Sender : TObject ); // Saisie d'un Bordereau

      // Utilitaires
      procedure OnClickParamCollectif  ( Sender : TObject ); // Paramètres du compte collectif
      procedure OnClickParamTiers      ( Sender : TObject ); // Paramètres du compte auxiliaire
      procedure OnClickCompteRib       ( Sender : TObject ); // Relevé d'indentité bancaire

      // Fonctions Complémetaires
      procedure OnClickCommentaire           ( Sender : TObject ); // Commentaires de l'auxliaire
      procedure OnClickCumulsCollectif       ( Sender : TObject ); // Cumuls du compte collectif par défaut
      procedure OnClickCumulsAuxiliaire      ( Sender : TObject ); // Cumuls Auxiliaires
      procedure OnClickTableChancellerie     ( Sender : TObject ); // Tables de chancelleries ( PAS OK CWAS )
//      procedure OnClickAnalyseComplementaire ( Sender : TObject ); // Analyse Complémentaire (**)

{$IFDEF COMPTA}
      procedure OnClickGCD                   ( Sender : TObject );
      procedure OnClickProgrammeTravail      ( Sender : TObject ); // Programme de Travail
      procedure OnClickMemoCycle             ( Sender : TObject ); // Mémo cycle de révision de la documentation des travaux
      procedure OnClickMemoObjectif          ( Sender : TObject ); // Mémo Objectif de révision de la documentation des travaux
      procedure OnClickMemoSynthese          ( Sender : TObject ); // Mémo synthèse
      procedure OnClickMemoMillesime         ( Sender : TObject ); // Mémo millesime
      procedure OnClickMemoCompte            ( Sender : TObject ); // Mémo compte général

      procedure OnClickAPG                   ( Sender : TObject ); // Appréciation générale
      procedure OnClickSCY                   ( Sender : TObject ); // Synthèse des cycles
      procedure OnClickEXP                   ( Sender : TObject ); // Attestation Expert

      procedure OnClickDossPermanent         ( Sender : TObject ); // Dossier Permanent
      procedure OnClickDossAnnuel            ( Sender : TObject ); // Dossier Annuel
      procedure OnClickRechercheDoc          ( Sender : TObject ); // Recherche de document
      procedure OnClickNoteTravail           ( Sender : TObject ); // Note de Travail
      procedure OnClickTableauVariation      ( Sender : TObject ); // Tableau de Variation
      procedure OnClickRechercheEcritures    ( Sender : TObject ); // Recherche d'écritures
      procedure OnClickFeuilleExcel          ( Sender : TObject ); // Feuille Excel

      // Editions
      procedure PrepareArgumentGL       ( vBoJustifSolde : Boolean);
      procedure PrepareArgumentBAL;
   {$ENDIF}

      procedure OnClickJustifSolde      ( Sender : TObject );
      procedure OnClickGLAuxi           ( Sender : TObject );
      procedure OnClickGLAuxiParGene    ( Sender : TObject );
      procedure OnClickBALAuxi          ( Sender : TObject );
      procedure OnClickBALAuxiParGene   ( Sender : TObject );
      procedure AuxiElipsisClick        ( Sender : TObject );

   protected
      procedure InitControl                              ; override; // Init des composants de la fiche
      procedure RemplitATobFListe                        ; override;
      procedure RefreshFListe( vBoFetch : Boolean )      ; override; //
      function  BeforeLoad : Boolean                     ; override;
      function  AjouteATobFListe( vTob : Tob ) : Boolean; override;
   public

  end ;


function CPLanceFiche_CPCONSAUXI( vStParam : string = '' ) : string ;

Implementation

Uses
{$IFDEF MODENT1}
     CPProcMetier,
{$ENDIF MODENT1}
     ParamSoc,             // GetParamSocSecur
{$IFDEF COMPTA}
     RappAuto,             // RapprochementAuto
     SaisBor,              // SaisieFolio
     uTofCPGLAuxi,         // CPLanceFiche_CPGLAuxi
     CPBALAUXI_TOF,        // CPLanceFiche_BalanceAuxiliaire
     CPBALAUXIGEN_Tof,     // CPLanceFiche_BalanceAuxiGen
     Lettrage,             //
     uTOFHistoCpte,        // CC_LanceFicheHistoCpte;
     CPGestionCReance ,    // 
     CPRevProgTravail_TOF, // CPLanceFiche_CPREVProgTravail
     CPREVDocTravaux_TOF,  // CPLanceFiche_CPRevDOCTRAVAUX
     CRevInfoDossier_TOM,  // CPLanceFiche_CREVNOTESUPERVISION
     uTofCPTableauVar,     // CPLanceFiche_CPTableauVar
     uTofCPNoteTravail,    // CPLanceFiche_CPNoteTravail

{$ENDIF}
     SaiSUtil,         //
     CPGeneraux_TOM,   // FicheGene
     CPTiers_TOM,      // FicheTiers
     Cummens,          // CumulCpteMensuel
     FichComm,         // FicheRIB_AGL
     Devise_TOM,       // FicheDevise
     CPChancell_TOF,   // FicheChancel
     uTofConsEcr,      // OperationsSurComptes
     uTofConsGene,     // CPLanceFiche_CPCONSGENE
     uTofCPMulMvt,     // MultiCritereMvt
     SaisComm,         // GetO
     AGLFctSuppr,      // CPSupprimeListeEnreg
     LettUtil,         // CExecDelettrage
     uLibWindows,      // IIF
     uLibEcriture,     // CEstAutoriseDelettrage
     uImportBob_TOF,   // CPLanceFiche_CPIMPFICHEXCEL
     uTofRechercheEcr, // CPLanceFiche_CPRechercheEcr
     Variants,         // unAssigned
     UTofMulParamGen; {23/04/07 YMO F5 sur Auxiliaire }

const cFI_TABLE    = 'CPCONSAUXI';
      cOrder       = 'ORDER BY T_AUXILIAIRE';
      cStChampsSup = ',T_TOTDEBE TOTDEBE, T_TOTCREE TOTCREE, ' +
                     'T_TOTDEBP TOTDEBP, T_TOTCREP TOTCREP, ' +
                     'T_TOTDEBS TOTDEBS, T_TOTCRES TOTCRES, ' +
                     'T_COLLECTIF TCOLLECTIF, T_LETTRABLE LETTRABLE, ' +
                     'T_MULTIDEVISE MULTIDEVISE, T_DEVISE LADEVISE';

      cNS = -999999999.99;               

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPCONSAUXI( vStParam : string = '' ) : string ;
begin
{$IFDEF CCSTD}
  // GCO - 26/10/2004 - FQ 14859
  Result := AGLLanceFiche('CP', 'CPCONSAUXI', '', '', vStParam);
{$ELSE}
  if CtxPcl in V_Pgi.PgiContexte then
  begin
    if(VH^.CpExoRef.Code = VH^.Encours.Code) or (VH^.CpExoRef.Code = VH^.Suivant.Code) then
      Result := AGLLanceFiche('CP', 'CPCONSAUXI', '', '', vStParam)
    else
      PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
              'être l''exercice en cours ou le suivant.', 'Consultation des comptes auxiliaires');
  end
  else
    Result := AGLLanceFiche('CP', 'CPCONSAUXI', '', '', vStParam);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/05/2002
Modifié le ... : 22/01/2003
Description .. : - LG - 22/01/2003 - appel de init auto pour gere
Suite ........ : l'autocomplementation
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnArgument (S : String ) ;
begin
  FStArgumentTOF := S; // Récupération des arguments
  FFI_Table      := cFI_TABLE;
  FStListeParam  := 'CPCONSAUXI';

  inherited ;
  Ecran.HelpContext := 7603200 ;

  // Elements de la fiche Ancêtre
  FListe.OnDblClick  := OnClickDetailCompte;
  FListe.Options     := FListe.Options + [goRangeSelect] ;
  FListe.MultiSelect := False; // GCO - 28/07/2005

  IndiceColonne;

  FStSelect    := FStListeChamps;
  FStSelect    := CSqlTextFromList(FStSelect) ;

  Ecran.OnKeyDown := OnKeyDownEcran;

  T_Auxiliaire    := THEdit(GetControl('T_AUXILIAIRE', True));
  T_Auxiliaire_   := THEdit(GetControl('T_AUXILIAIRE_', True));
  T_Collectif     := THEdit(GetControl('T_COLLECTIF', True));
  T_DateDernMvt   := THEdit(GetControl('T_DATEDERNMVT', True));
  T_DateDernMvt_  := THEdit(GetControl('T_DATEDERNMVT_', True));

  TModeSelection  := THValComboBox(GetControl('TMODESELECTION', True));
  TCorresp1       := THMultiValComboBox(GetControl('TCORRESP1', True));
  TSituationCompte:= THValComboBox(GetControl('TSITUATIONCOMPTE', True));
  TSensReel       := THValComboBox(GetControl('TSENSREEL', True));
  T_NatureAuxi    := THMultiValComboBox(GetControl('T_NATUREAUXI', True));

  BBlocNote       := TToolBarButton97(GetControl('BBLOCNOTE', True));

  PopUpComptable     := TPopUpMenu(GetControl('PopUpComptable', True));
  PopUpSpecifique    := TPopUpMenu(GetControl('PopUpSpecifique', True));
  PopUpAideRevision  := TPopUpMenu(GetControl('PopUpAideRevision', True));
  PopUpRevInteg      := TPopUpMenu(GetControl('PopUpRevInteg', True));
  PopUpParametre     := TPopUpMenu(GetControl('PopUpParametre', True));
  PopUpEdition       := TPopUpMenu(GetControl('PopUpEdition', True));

  TabTablesLibres := TTabSheet(GetControl('TABTABLESLIBRES', True));

  // Révision avancée
  TTypePonderation   := THValComboBox(GetControl('TTYPEPONDERATION', True));
  ZCPonderation      := THValComboBox(GetControl('ZCPONDERATION', True));
  TValeurPonderation := THNumEdit(GetControl('TVALEURPONDERATION', True));

  TTypeVariation     := THValComboBox(GetControl('TTYPEVARIATION', True));
  ZCVariation        := THValComboBox(GetControl('ZCVARIATION', True));
  TValeurVariation   := THNumEdit(GetControl('TVALEURVARIATION', True));
  TNatureVariation   := THValComboBox(GetControl('TNATUREVARIATION', True));

  BEffaceRevision    := TToolBarButton97(GetControl('BEFFACEREVISION', True));
  // Fin de Révision avancée

  //----------------------- Branchement des événements -------------------------
  PopF11.OnPopup            := OnPopUpPopF11;
  PopUpComptable.OnPopup    := OnPopUpComptable;
  PopUpSpecifique.OnPopup   := OnPopUpSpecifique;
  PopUpAideRevision.OnPopup := OnPopUpAideRevision;
  PopUpRevInteg.OnPopup     := OnPopUpRevInteg;
  PopUpParametre.OnPopup    := OnPopUpParametre;
  PopUpEdition.OnPopUp      := OnPopUpEdition;

  FListe.GetCellCanvas      := GetCellCanvasFListe;

  TNatureVariation.OnChange := OnChangeTNatureVariation;
  BEffaceRevision.OnClick   := OnClickBEffaceRevision;

  TripotageTSituationCompte;
  InitControl; // Valeur par défaut de tous les controles de la fiche

  T_Auxiliaire.MaxLength       := VH^.Cpta[fbAux].lg;
  T_Auxiliaire_.MaxLength      := VH^.Cpta[fbAux].lg;
  T_Collectif.MaxLength        := VH^.Cpta[fbGene].lg;
  T_Auxiliaire.OnElipsisClick  := OnElipsisClickT_Auxiliaire;
  T_Auxiliaire_.OnElipsisClick := OnElipsisClickT_Auxiliaire;
  T_Auxiliaire.OnExit          := OnExitT_Auxiliaire;
  T_Auxiliaire_.OnExit         := OnExitT_Auxiliaire;

  if CtxPcl in V_Pgi.PgiContexte then
  begin
    FExoDate.Code := VH^.CPExoRef.Code;
    FExoDate.Deb  := VH^.CPExoRef.Deb;
    FExoDate.Fin  := VH^.CPExoRef.Fin;
  end
  else
  begin
    if (VH^.Entree.Code = VH^.EnCours.Code) or
       (VH^.Entree.Code = VH^.Suivant.Code) then
    begin
      FExoDate.Code := VH^.Entree.Code;
      FExoDate.Deb  := VH^.Entree.Deb;
      FExoDate.Fin  := VH^.Entree.Fin;
    end
    else
    begin
      FExoDate.Code := VH^.EnCours.Code;
      FExoDate.Deb  := VH^.EnCours.Deb;
      FExoDate.Fin  := VH^.EnCours.Fin;
    end;
  end;

  // Voir si besoin de gérer la fiche en mode TaConsult lors de certains appels
  FAction := TaModif;

  // GCO - 11/04/2007
  FBoOkRevision := VH^.Revision.Plan <> '';

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('T_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('T_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

  // GCO - 24/07/2007 - FQ 20984          
  BSelectAll.Visible := False;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnLoad ;
var lStArgument : string;
begin
  if FStArgumentTOF <> '' then
  begin
    lStArgument := FStArgumentTOF;
    T_Collectif.Text := ReadTokenSt(lStArgument);
  end;

  LibellesTableLibre(TabTablesLibres, 'TT_TABLE', 'T_TABLE', 'T');
  Inherited ;
  InitAllPopUp( False );
  FToutAccorde := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False);
  InitAutoSearch ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2004
Modifié le ... : 27/04/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.AfterShow;
var lNbMoisSoldeE : word ;
    lNbMoisSoldeP : word ;
    lPremMois     : word ;
    lPremAnnee    : word ;
begin
  inherited; // Lancement d'un BchercheClick
  if T_Auxiliaire.CanFocus then
    T_Auxiliaire.SetFocus;

  // GCO - 10/06/2005 - FQ 15985
  if FExoDate.Code = VH^.Suivant.Code then
  begin
    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Suivant.Deb, VH^.Suivant.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);
  end
  else
  begin
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Precedent.Deb, VH^.Precedent.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);
  end;

  if (lNbMoisSoldeE = lNbMoisSoldeP) then
  begin
    TTypeVariation.Enabled := False;
    TTypevariation.Value := 'ABSOLUE';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/08/2002
Modifié le ... :   /  /
Description .. : Mise à jour du Caption de l' écran en fonction de la recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.MiseAJourCaptionEcran;
begin
  Ecran.Caption := TraduireMemoire('Consultation des comptes auxiliaires') + ' : ' + RechDom('TTEXERCICE', FExoDate.Code, False);
  UpDateCaption(Ecran);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnNew ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnDelete ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnUpdate ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.IndiceColonne ;
var
 lStChamps : string;
 lStListe  : string;
 lInIndex  : integer;
begin
  lStListe        := FStListeChamps;
  lInIndex        := 1;
  FColDebitE      := -1;
  FColCreditE     := -1;
  FColDebitS      := -1;
  FColCreditS     := -1;
  FColSoldeP      := -1;
  FColSoldeE      := -1;
  FColVariation   := -1;
  FColPourcentage := -1;

  while lStListe <> '' do
  begin
   lStChamps := ReadTokenSt(lStListe);

   if ( lStChamps = 'T_TOTDEBE' ) then
     FColDebitE := lInIndex
   else
     if ( lStChamps = 'T_TOTCREE' ) then
       FColCreditE := lInIndex
     else
       if ( lStChamps = 'T_TOTDEBS' ) then
         FColDebitS := lInIndex
       else
         if ( lStChamps = 'T_TOTCRES' ) then
           FColCreditS := lInIndex
         else
           if ( lStChamps = 'SOLDEE' ) then
             FColSoldeE := lInIndex
           else
             if ( lStChamps = 'SOLDEP' ) then
               FColSoldeP := lInIndex
             else
               if ( lStChamps = 'VARIATION' ) then
                 FColVariation := lInIndex
               else
                 if ( lStChamps = 'POURCENTAGE' ) then
                   FColPourcentage := lInIndex;

    Inc(lInIndex);
  end; // while
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/06/2002
Modifié le ... : 06/04/2005
Description .. : Initialise les composants dans leur état par défaut à l' ouverture
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitControl;
begin
  // Init des valeurs
  T_Auxiliaire.Text      := '';
  T_Auxiliaire_.Text     := '';
  T_NatureAuxi.Value     := '<<' + TraduireMemoire('Tous') + '>>';
  TCorresp1.Value        := '<<' + TraduireMemoire('Tous') + '>>';

  TModeSelection.Value   := 'ALL';
  TSensReel.Value        := '0';
  TSituationCompte.Value := '';

  T_DateDernMvt.Text     := DateToStr( iDate1900 );
  T_DateDernMvt_.Text    := DateToStr( iDate2099 );

  BBlocNote.Visible      := False;

  OnClickBEffaceRevision(nil);

  // GCO - 25/10/2006 - FQ 16894
  if VH^.Precedent.Code = '' then
  begin
    TTypeVariation.Enabled   := False;
    ZCVariation.Enabled      := False;
    TValeurVariation.Enabled := False;
    TNatureVariation.Enabled := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/12/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnExitT_Auxiliaire(Sender: TObject);
begin
  if (Trim(THEdit(Sender).Text) = '') or (TestJoker(THEdit(Sender).Text)) then Exit;
  if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
    THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2005
Modifié le ... : 02/10/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnElipsisClickT_Auxiliaire(Sender: TObject);
begin
  // GCO - 02/10/2007 - FQ 21544 - 2 au lieu de 1 dans le dernier paramètre
  LookUpList( THEdit(Sender),
             'Comptes auxiliaires',
             'TIERS',
             'T_AUXILIAIRE',
             'T_LIBELLE, T_NATUREAUXI',
             CGenereSQLConfidentiel('T') +
             ' AND (T_NATUREAUXI <> "NCP" AND T_NATUREAUXI <> "CON" AND ' +
             'T_NATUREAUXI <> "PRO" AND T_NATUREAUXI <> "SUS")',
             'T_AUXILIAIRE',
             True,
             2 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/05/2002
Modifié le ... :   /  /
Description .. : Charge la Tob avec la requete grille avec la requête
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.RemplitATobFListe;
var lStSelect   : string ;
    lStWhere    : string ;
    lSt         : string;
begin
  if (FExoDate.Code = VH^.Encours.Code) or (FExoDate.Code = VH^.Suivant.Code) then
  begin
    lSt := RecupWhereCritere(PageControl);
    if lSt <> '' then
      lStWhere := lSt + ' AND ' + RecupAutreWhere
    else
      lStWhere := ' WHERE ' + RecupAutreWhere;

    lSt := RecupWhereSituationCompte;
    if lSt <> '' then
      lStWhere := lStWhere + lSt;
  
    // Cas particulier pour la requête générée, on recherche les comptes
    // mouvementés sur les deux exercices sans les ANO
    if TModeSelection.Value = 'EXS' then
    begin
      lStSelect := 'SELECT DISTINCT(E_GENERAL),' + CSqlTextFromList(FStListeChamps) + ' ' +
                   cStChampsSup + ' FROM ECRITURE ' +
                   'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
                   'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
                   lStWhere + ' AND E_QUALIFPIECE = "N"' +
                   'AND (E_EXERCICE = "' + FStCodeExoE + '" OR E_EXERCICE = "' + FStCodeExoP + '")' +
                   ' AND J_NATUREJAL <> "ANO" '  + cOrder;
    end
    else
    begin
      lStSelect := 'SELECT ' + CSqlTextFromList(FStListeChamps) + ' ' + cStChampsSup +
                 ' FROM TIERS ' + lStWhere + ' ' + cOrder;
    end;

    // On envoie le texte SQL de la requête à la fiche ANCETRE pour qu'elle fasse le OPENSQL
    AStSqlTobFListe := lStSelect;

    // Remplit FRevision avec les valeurs saisies dans l'onglet "Révision Avancée"
    RecupCritereRevision;

  end
  else
  begin
    PgiInfo('Pour pouvoir utiliser cette fonction, l''exercice de référence doit ' +
             'être l''exercice en cours ou le suivant.', 'Consultation des comptes généraux');

    AStSqlTobFListe := '';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/11/2002
Modifié le ... : 20/01/2003
Description .. : Récupère la condition WHERE des critères qui
Suite ........ : n'appartiennent
Suite ........ : pas à la table
Suite ........ : - FB 11804 - la gestion des plans de correspond ne fct pas
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSAUXI.RecupAutreWhere : string;
var lStValCorresp1 : string;
    lStLibCorresp1 : string;
begin
  // GCO - 15/09/2004 - FQ 14087 - Gestion du T_CONFIDENTIEL
  if (V_Pgi.Confidentiel = '0') then
    Result := '((T_CONFIDENTIEL = "0") OR (T_CONFIDENTIEL = "-"))'
  else
  begin
    Result := '(((T_CONFIDENTIEL = "-") OR (T_CONFIDENTIEL = "X")) OR ' +
              '(T_CONFIDENTIEL <= "' + V_PGI.Confidentiel + '"))';
  end;

  // Mode de sélection des comptes
  // le Cas où Value = 'EXS' est géré à la création du texte SQL
  if TModeSelection.Value = 'EXT' then
  begin // Comptes mouvementés sur les deux exercices (avec ANO)
    Result := Result + ' AND ((T_TOTDEB' + FStPrefixeE + ' <> 0 OR T_TOTCRE' + FStPrefixeE + ' <> 0) OR' +
                       ' (T_TOTDEB' + FStPrefixeP + ' <> 0 OR T_TOTCRE' + FStPrefixeP + ' <> 0))';
  end
  else
    if TModeSelection.Value = 'NSL' then
    begin // Comptes non soldés sur les deux exercices
      Result := Result + ' AND ((T_TOTDEB' + FStPrefixeE + ' <> T_TOTCRE' + FStPrefixeE + ') OR' +
                         ' (T_TOTDEB' + FStPrefixeP + ' <> T_TOTCRE' + FStPrefixeP + '))';
    end;

  // Sens Réel du solde '0' = Tous les Comptes
  if TSensReel.Value = '1' then // Débiteur
    Result := Result + ' AND (T_TOTDEB' + FStPrefixeE + ' > T_TOTCRE' + FStPrefixeE + ')'
  else
    if TSensReel.Value = '2' then // Créditeur
      Result := Result + ' AND (T_TOTDEB' + FStPrefixeE + ' < T_TOTCRE' + FStPrefixeE + ')'
    else
      if TSensReel.Value = '3' then // Soldé
        Result := Result + ' AND (T_TOTDEB' + FStPrefixeE + ' = T_TOTCRE' + FStPrefixeE + ')';

  // GCO - 23/08/2005 - FQ 11950 et 16467
  // Compte de Correspondance
  TraductionTHMultiValComboBox( TCorresp1, lStValCorresp1, lStLibCorresp1, 'T_CORRESP1', True );
  if lStValCorresp1 <> '' then
    Result := Result + ' AND ' + lStValCorresp1;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSAUXI.RecupWhereSituationCompte: string;
begin
  Result := '';

  if TSituationCompte.Value = 'SI1' then
  begin // Comptes mouvementés sur l'exercice mais sans mouvements sur le précédent
    Result := ' AND T_AUXILIAIRE IN (SELECT DISTINCT (E1.E_AUXILIAIRE) FROM ECRITURE E1 WHERE' +
              ' E1.E_EXERCICE = "' + FStCodeExoE + '" AND' +
              ' E1.E_ECRANOUVEAU = "N" AND' +
              ' E1.E_QUALIFPIECE = "N" AND E1.E_AUXILIAIRE NOT IN' +
              ' (SELECT E2.E_AUXILIAIRE FROM ECRITURE E2 WHERE' +
              ' E2.E_EXERCICE = "' + FStCodeExoP + '" AND' +
              ' E2.E_ECRANOUVEAU = "N" AND' +
              ' E2.E_QUALIFPIECE = "N" ))';
  end
  else
  if TSituationCompte.Value = 'SI2' then
  begin // Comptes sans mouvements sur l'exercice mais mouvementés sur le précédent
    Result := ' AND T_AUXILIAIRE IN (SELECT DISTINCT (E1.E_AUXILIAIRE) FROM ECRITURE E1 WHERE' +
              ' E1.E_EXERCICE = "' + FStCodeExoP + '" AND' +
              ' E1.E_ECRANOUVEAU = "N" AND' +
              ' E1.E_QUALIFPIECE = "N" AND E1.E_AUXILIAIRE NOT IN' +
              ' (SELECT E2.E_AUXILIAIRE FROM ECRITURE E2 WHERE' +
              ' E2.E_EXERCICE = "' + FStCodeExoE + '" AND' +
              ' E2.E_ECRANOUVEAU = "N" AND' +
              ' E2.E_QUALIFPIECE = "N" ))';
  end
  else
  if TSituationCompte.Value = 'SI3' then
  begin // Comptes non soldés sur l'exercice mais soldés sur le précédent
    Result := ' AND (T_TOTDEB' + FStPrefixeE + ' <> T_TOTCRE' + FStPrefixeE + ')' +
              ' AND (T_TOTDEB' + FStPrefixeP + ' = T_TOTCRE' + FStPrefixeP + ')';
  end
  else
  if TSituationCompte.Value = 'SI4' then
  begin // Comptes soldés sur l'exercice mais son soldés sur le précédent
    Result := ' AND (T_TOTDEB' + FStPrefixeE + ' = T_TOTCRE' + FStPrefixeE + ')' +
              ' AND (T_TOTDEB' + FStPrefixeP + ' <> T_TOTCRE' + FStPrefixeP + ')';
  end
  else
  if TSituationCompte.Value = 'SI5' then
  begin //Comptes n'ayant pas le même sens sur l'exercice que sur le précédent
    Result := ' AND (((T_TOTDEB' + FStPrefixeE + ' - T_TOTCRE' + FStPrefixeE + ' > 0) AND ' +
                    '(T_TOTDEB' + FStPrefixeP + ' - T_TOTCRE' + FstPrefixeP + ' < 0)) OR ' +
                    '((T_TOTDEB' + FStPrefixeE + ' - T_TOTCRE' + FStPrefixeE + ' < 0) AND ' +
                    '(T_TOTDEB' + FStPrefixeP + ' - T_TOTCRE' + FstPrefixeP + ' > 0)))';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.RecupCritereRevision;
begin
  // GCO - 07/06/2007 - FQ 20613
  //FValeurPonderation := Valeur(TValeurPonderation.Text);
  FValeurPonderation := TValeurPonderation.Value;

  if TNatureVariation.ItemIndex <> -1 then
    //FValeurVariation := Valeur(TValeurVariation.Text);
    FValeurVariation := TValeurVariation.Value;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpPopF11(Sender: TObject);
begin
  InitAllPopUp( True );
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitAllPopUp( vActivation : Boolean );
begin
  InitPopUpComptable    ( vActivation );
  InitPopUpSpecifique   ( vActivation );
  InitPopUpAideRevision ( vActivation ); 
  InitPopUpRevInteg     ( vActivation );
  InitPopUpParametre    ( vActivation );
  InitPopUpEdition      ( vActivation );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpComptable(Sender: TObject);
begin
  InitPopUpComptable( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitPopUpComptable( vActivation : Boolean );
var i : integer;
    lMenuitem : TMenuItem;
begin
  for i := 0 to PopUpComptable.Items.Count -1 do
  begin
    lMenuItem := PopUpComptable.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'DETAILCOMPTE' then
      begin
        if vActivation then
          lMenuItem.Enabled := True
        else
          lMenuItem.OnClick := OnClickDetailCompte;
        Continue;
      end;

      if lMenuItem.Name = 'DETAILCOMPTE2' then
      begin
        if vActivation then
          lMenuItem.Enabled := True
        else
          lMenuItem.OnClick := OnClickDetailCompte2;
        Continue;
      end;

      if lMenuItem.Name = 'SAISIEBOR' then
      begin
        if vActivation then
          lMenuItem.Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
        else
          lMenuItem.OnClick := SaisieBorClick ;
        Continue;
      end;

      if lMenuitem.Name = 'SAISIEPIECE' then
      begin
        if vActivation then
          lMenuItem.Enabled := ExJaiLeDroitConcept(TConcept(ccSaisEcritures),False)
        else
          lMenuItem.OnClick := SaisieEcrClick ;
        Continue;
      end;

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpSpecifique( Sender : TObject );
begin
  InitPopUpSpecifique( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitPopUpSpecifique( vActivation : Boolean );
var i,j : integer;
    lMenuItem : TMenuItem;
    lBoOkLettrable : Boolean;
begin
  lBoOkLettrable := False;

  if vActivation then
  begin
    if ATobRow = nil then Exit;
    lBoOkLettrable := (ATobRow.GetValue('LETTRABLE') = 'X');
  end;

  for i := 0 to PopUpSpecifique.Items.Count -1 do
  begin
    lMenuItem := PopUpSpecifique.Items[i];
    try

      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'GCD' then
        begin // Créer une Créance dans GCD
          if vActivation then
            lMenuItem.Enabled := VH^.OkModGCD and ( ATobRow.GetString('T_NATUREAUXI') = 'CLI' )
          else
            lMenuItem.OnClick := OnClickGCD;
          Continue;
        end;
      {$ENDIF}

      if lMenuItem.Name = 'LETTRAGE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'LETMANUEL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := lBoOkLettrable // FQ 16349
            else
              lMenuItem.Items[j].OnClick := LettrageMClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'LETAUTO' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := lBoOkLettrable
            else
              lMenuItem.Items[j].OnClick := LettrageAClick;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DELETEXOREF' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := lBoOkLettrable
            else
              lMenuItem.Items[j].OnClick := DelettreMvtExoRef;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'DELETTOTAL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := lBoOkLettrable
            else
              lMenuItem.Items[j].OnClick := DelettreTotal;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'ETATJUSTIFSOLDE' then
          begin // Justificatif de solde
            if vActivation then
              lMenuItem.Items[j].Enabled := lBoOkLettrable
            else
              lMenuItem.Items[j].OnClick := OnClickJustifSolde;
            Continue;
          end;

        end; // FOR

      end; // LETTRAGE

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpAideRevision(Sender: TObject);
begin
  InitPopUpAideRevision( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitPopUpAideRevision(vActivation: Boolean);
var i,j : integer;
    lMenuItem : TMenuItem;
begin
  if (ATobRow = nil) and (vActivation) then Exit;

  for i := 0 to PopUpAideRevision.Items.Count -1 do
  begin
    lMenuItem := PopUpAideRevision.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'OUTILREVISION' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'NOTETRAVAIL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobRow.GetString('TCOLLECTIF') <> '') and VH^.OkModRic and JaiLeRoleCompta( rcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickNoteTravail;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'TABLEAUVARIATION' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobRow.GetString('TCOLLECTIF') <> '') and VH^.OkModRic and JaiLeRoleCompta( rcReviseur )
            else
             lMenuItem.Items[j].OnClick := OnClickTableauVariation;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'RECHERCHEECR' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := True
            else
              lMenuItem.Items[j].OnClick := OnClickRechercheEcritures;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'FEUILLEEXCEL' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModRic and JaiLeRoleCompta( rcReviseur )
            else
              lMenuItem.Items[j].OnClick := OnClickFeuilleExcel;
            Continue;
          end;

        {$ENDIF}

        end; // FOR
      end; // LETTRAGE

      if lMenuItem.Name = 'ANALYSE' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'CUMULCOLLECTIF' then
          begin
            if vActivation then
             lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0)
            else
              lMenuItem.Items[j].OnClick := OnClickCumulsCollectif ;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'CUMULAUXI' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0)
            else
              lMenuItem.Items[j].OnClick := OnClickCumulsAuxiliaire ;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'COMMENTAIRE' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0)
            else
              lMenuItem.Items[j].OnClick := OnClickCommentaire;
            Continue;
          end;
        end; // FOR
      end; // ANALYSE
    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpRevInteg( Sender : TObject );
begin
  InitPopUpRevInteg( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitPopUpRevInteg( vActivation : Boolean );
var i,j : integer;
    lMenuItem : TMenuItem;
{$IFDEF COMPTA}
    lBoCycleRevision : Boolean;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lBoCycleRevision := False;
{$ENDIF}

  if vActivation then
  begin
    if ATobRow = nil then Exit;
  {$IFDEF COMPTA}
    if FBoOkRevision and (ATobRow.GetString('TCOLLECTIF') <> '') then
      lBoCycleRevision := GetColonneSQL('GENERAUX', 'G_CYCLEREVISION', 'G_GENERAL = "' + ATobRow.GetString('TCOLLECTIF') + '"') <> '';
  {$ENDIF}
  end;

  for i := 0 to PopUpRevInteg.Items.Count -1 do
  begin
    lMenuItem := PopUpRevInteg.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      {$IFDEF COMPTA}
      if lMenuItem.Name = 'PROGTRAVAIL' then
      begin
        // GCO - 25/07/2007 - FQ 20944
        lMenuItem.Caption := TraduireMemoire('Programme de &travail');

        if vActivation then
          lMenuItem.Enabled := VH^.OkModRic and FBoOkRevision and
                               ((JaiLeRoleCompta(rcSuperviseur)) or
                                (JaiLeRoleCompta(rcReviseur) and
                                 not JaileRoleCompta(rcSuperviseur))
                               )
        else
          lMenuItem.OnClick := OnClickProgrammeTravail;
        Continue;
      end;
      {$ENDIF}

      if lMenuItem.Name = 'DOCTRAVAUX' then
      begin
        lMenuItem.Enabled := VH^.OkModRIC;

        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'MEMOCYCLE' then
          begin // Mémo cycle
            if vActivation then
              lMenuItem.Items[j].Enabled := JaiLeRoleCompta( RcReviseur ) and
                                            FBoOkRevision and lBoCycleRevision
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCycle;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'MEMOOBJECTIF' then
          begin // Mémo objectif de révision
            if vActivation then
              lMenuItem.Items[j].Enabled := (CtxPcl In V_Pgi.PgiContexte) and
                                            JaiLeRoleCompta( RcReviseur ) and
                                            FBoOkRevision and lBoCycleRevision
            else
              lMenuItem.Items[j].OnClick := OnClickMemoObjectif;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'MEMOSYNTHESE' then
          begin // Mémo synthèse du cycle
            if vActivation then
              lMenuItem.Items[j].Enabled := JaiLeRoleCompta( RcReviseur ) and
                                            FBoOkRevision and lBoCycleRevision
            else
              lMenuItem.Items[j].OnClick := OnClickMemoSynthese;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'MEMOMILLESIME2' then
          begin // Mémo commantire millésimé
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobRow.GetString('TCOLLECTIF') <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoMillesime;
            Continue;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'MEMOCOMPTE2' then
          begin // Mémo compte général
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobRow.GetString('TCOLLECTIF') <> '')
            else
              lMenuItem.Items[j].OnClick := OnClickMemoCompte;
            Continue;
          end;
          {$ENDIF}
        end; // FOR
      end; // DOCTRAVAUX

      {$IFDEF COMPTA}
      {$IFNDEF CCMP}
      if lMenuItem.Name = 'SUPERVISIONTRAVAUX' then
      begin // GCO - 21/06/2007 - FQ 20584 + FQ 21366
        lMenuItem.Enabled := (CtxPcl In V_Pgi.PgiContexte) and VH^.OkModRIC and
                             FBoOkRevision and JaiLeRoleCompta( rcSuperviseur );
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := lMenuItem.Enabled; // GCO - 13/09/2007 - FQ 21366
          if lMenuItem.Items[j].Name = 'APG' then
          begin // Appréciation générale du dossier
            if not vActivation then
              lMenuItem.Items[j].OnClick := OnClickAPG;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'SCY' then
          begin // Synthèse des cycles
            if not vActivation then
              lMenuItem.Items[j].OnClick := OnClickSCY;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'EXP' then
          begin // Attestation Expert
            if not vActivation then
              lMenuItem.Items[j].OnClick := OnClickEXP;
            Continue;
          end;
        end;
      end;
      {$ENDIF}
      {$ENDIF}

      if lMenuItem.Name = 'DOSSIERCLIENT' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'DOSPERMANENT' then
          begin // Dossier permanent
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickDossPermanent;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'DOSANNUEL' then
          begin // Dossier annuel
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickDossAnnuel;
          end;
          {$ENDIF}

          {$IFDEF COMPTA}
          if lMenuItem.Items[j].Name = 'RECHERCHEDOC' then
          begin // Recherche de document
            if vActivation then
              lMenuItem.Items[j].Enabled := VH^.OkModGed
            else
              lMenuItem.Items[j].OnClick := OnClickRechercheDoc;
          end;
          {$ENDIF}

        end; // FOR
      end; // DOSSIERCLIENT
    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpParametre( Sender : TObject );
begin
  InitPopUpParametre( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitPopUpParametre( vActivation : Boolean );
var i,j : integer;
    lMenuItem : TMenuItem;
    lBoOkTableChancellerie : Boolean;
begin
  lBoOkTableChancellerie := False;

  if VActivation then
  begin
    if ATobRow = nil then Exit;
    lBoOkTableChancellerie := (ATobRow.GetString('LADEVISE') <> V_PGI.DevisePivot) or
                              (ATobRow.GetString('MULTIDEVISE') = 'X');
  end;

  for i := 0 to PopUpParametre.Items.Count -1 do
  begin
    lMenuItem := PopUpParametre.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      if lMenuItem.Name = 'PARAMETRES' then
      begin
        lMenuItem.Enabled := True;
        for j := 0 to lMenuItem.Count - 1 do
        begin
          lMenuItem.Items[j].Enabled := False;

          if lMenuItem.Items[j].Name = 'PARAMCOLLECTIF' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled  := ATobRow.GetString('TCOLLECTIF') <> ''
            else
              lMenuItem.Items[j].OnClick  := OnClickParamCollectif;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'PARAMTIERS' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := (ATobFListe.Detail.Count > 0)
            else
              lMenuItem.Items[j].OnClick  := OnClickParamTiers;
            Continue;
          end;

          if lMenuItem.Items[j].Name = 'COMPTERIB' then
          begin
            if vActivation then
              lMenuItem.Items[j].Enabled := ATobRow.GetValue('T_COLLECTIF') <> ''
            else
              lMenuItem.Items[j].OnClick  := OnClickCompteRib;
            Continue;
          end;
        end;
      end;

      if lMenuItem.Name = 'TABLECHANCELLERIE' then
      begin
        if vActivation then
          lMenuItem.Enabled := lBoOkTableChancellerie
        else
          lMenuItem.OnClick := OnClickTableChancellerie;
        Continue;
      end;

    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnPopUpEdition( Sender : TObject );
begin
  InitPopUpEdition( True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.InitPopUpEdition( vActivation : Boolean );
var i : integer;
    lMenuItem : TMenuItem;
begin
  for i := 0 to PopUpEdition.Items.Count -1 do
  begin
    lMenuItem := PopUpEdition.Items[i];
    try
      lMenuItem.Enabled := False;
      if (vActivation and ABoMultiSelected) then Continue;

      // GCO - 23/10/2007 - FQ 21677
      if lMenuItem.Name = 'EETATJUSTIFSOLDE' then
      begin // Justificatif de solde PGE
        if vActivation then
          lMenuItem.Enabled := not (CtxPcl in V_Pgi.PgiContexte) and (ATobRow.GetValue('LETTRABLE') = 'X')
        else
          lMenuItem.OnClick := OnClickJustifSolde;
        Continue;
      end;

      if lMenuItem.Name = 'GLAUXI' then
      begin // Grand livre Auxiliaire
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLAuxi;
        Continue;
      end;

      if lMenuItem.Name = 'GLAUXIPARGENE' then
      begin // Grand Livre Auxiliaire par Général
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickGLAuxiParGene;
        Continue;
      end;

      if lMenuItem.Name = 'BALAUXI' then
      begin // Balance Auxiliaire
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickBALAuxi;
        Continue;
      end;

      if lMenuItem.Name = 'BALAUXIPARGENE' then
      begin // Balance Auxiliaire par Général
        if vActivation then
          lMenuItem.Enabled := (ATobFListe.Detail.Count > 0)
        else
          lMenuItem.OnClick := OnClickBALAuxiParGene;
        Continue;
      end;
      
    finally
      if vActivation then ActivationMenuItem(lMenuItem);
    end;
  end;  
end;

////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  inherited;
  case Key of

    VK_F5 :    if (FListe.Focused) and ( ssShift in Shift) then
    	           OnClickDetailCompte2(nil);

    Ord('H') : if ssAlt in Shift then // ALT + H -> Table de Chancellerie
                 OnClickTableChancellerie(nil);

{$IFDEF COMPTA}
    Ord('O') : if ssAlt in Shift then // ALT + O -> Commentaire millésimé
                 OnClickMemoMillesime(nil);
{$ENDIF}

    Ord('R') : if ssAlt in Shift then // ALT + R -> Accès au RIB
                 OnClickCompteRib( nil );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSAUXI.BeforeLoad : Boolean;
begin
  FBoCriterePonderation := (TTypePonderation.ItemIndex <> -1) and
                           (ZCPonderation.ItemIndex <> -1);
                           //(TValeurPonderation.Text <> '');

  FBoCritereVariation   := (TTypeVariation.ItemIndex <> -1) and
                           (ZCVariation.ItemIndex <> -1) and
                           //(TValeurVariation.Text <> '') and
                           (TNatureVariation.ItemIndex <> -1);

  FStPrefixeE := IIF(FExoDate.Code = VH^.Encours.Code, 'E', 'S');
  FStPrefixeP := IIF(FExoDate.Code = VH^.Encours.Code, 'P', 'E');

  FStCodeExoE := IIF(FExoDate.Code = VH^.Encours.Code, VH^.Encours.Code, VH^.Suivant.Code);
  FStCodeExoP := IIF(FExoDate.Code = VH^.Encours.Code, VH^.Precedent.Code, VH^.Encours.Code);

  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSAUXI.AjouteATobFListe(vTob : Tob) : Boolean;
var lRdSoldeE       : double ;
    lRdSoldeP       : double ;
    lTempSoldeE     : double ;
    lTempSoldeP     : double ;
    lVariation      : double ;
    lPourcentageVariation : double ;
    lStPourcentageVariation : string ;
    lNbMoisSoldeE   : word   ;
    lNbMoisSoldeP   : word   ;
    lPremMois       : word   ;
    lPremAnnee      : word   ;

    lTotalDebit     : double ;
    lTotalCredit    : double ;
    lSoldeAbsolu    : double ;

    function _CalculRevision : Boolean;
    var lBoResultPonderation : Boolean;
        lBoResultVariation   : Boolean;
    begin
      lBoResultPonderation := False;
      lBoResultVariation   := False;

      if FBoCriterePonderation then
      begin
        lBoResultPonderation := False;

        // Type de la pondération
        if TTypePonderation.Value = 'SOLDEDEB' then
        begin // Solde Débiteur
          if lRdSoldeE > 0 then
            lBoResultPonderation := CompareOperateur(lRdSoldeE, FValeurPonderation, ZCPonderation.Value);
        end
        else
        if TTypePonderation.Value = 'SOLDECRE' then
        begin // Solde Créditeur
          if lRdSoldeE < 0 then
            lBoResultPonderation := CompareOperateur(Abs(lRdSoldeE), FValeurPonderation, ZCPonderation.Value);
        end
        else
        if TTypePonderation.Value = 'SOLDE' then
        begin // Solde
          lBoResultPonderation := CompareOperateur(lSoldeAbsolu, FValeurPonderation, ZCPonderation.Value);
        end
        else
        if TTypePonderation.Value = 'TOTALDEB' then
        begin // Total débit
          lBoResultPonderation := CompareOperateur(lTotalDebit, FValeurPonderation, ZCPonderation.Value);
        end
        else
        if TTypePonderation.Value = 'TOTALCRE' then
        begin // Total Crédit
          lBoResultPonderation := CompareOperateur(lTotalCredit, FValeurPonderation, ZCPonderation.Value);
        end;
      end;

      if FBoCritereVariation then
      begin
        lBoResultVariation := False;

        if TTypeVariation.Value = 'ABSOLUE' then
        begin // Variation Absolue (Brute)
          if TNatureVariation.Value = 'POURCENTAGE' then
          begin
            if (lRdSoldeP <> 0) then
              lBoResultVariation := CompareOperateur( lPourcentageVariation, FValeurVariation, ZCVariation.Value);
          end
          else
            lBoResultVariation := CompareOperateur( lRdSoldeE - lRdSoldeP, FValeurVariation, ZCVariation.Value);
        end
        else
        if TTypeVariation.Value = 'RELATIVE' then
        begin // Variation Annuelle (Les sommes sont ramenées sur 12 mois)
          if TNatureVariation.Value = 'POURCENTAGE' then
          begin
            if (lRdSoldeP <> 0) then
              lBoResultVariation := CompareOperateur( lPourcentageVariation, FValeurVariation, ZCVariation.Value);
          end
          else
            lBoResultVariation := CompareOperateur( lVariation, FValeurVariation, ZCVariation.Value);
        end;
      end;

      if FBoCriterePonderation and FBoCritereVariation then
      begin
        Result := lBoResultPonderation and lBoResultVariation
      end
      else
      begin
        if FBoCriterePonderation then
          Result := lBoResultPonderation
        else
          Result := lBoResultVariation;
      end;
    end;
begin
  if FExoDate.Code = VH^.Suivant.Code then
  begin
    // Colonne de G_TOTDEBE et G_TOTCREE sont cachées
    if FColDebitE  <> -1 then FListe.ColWidths[FColDebitE]  := -1;
    if FColCreditE <> -1 then FListe.ColWidths[FColCreditE] := -1;

    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Suivant.Deb, VH^.Suivant.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);

    lRdSoldeE := vTob.GetValue('TOTDEBS') - vTob.GetValue('TOTCRES');
    lRdSoldeP := vTob.GetValue('TOTDEBE') - vTob.GetValue('TOTCREE');

    lTotalDebit   := vTob.GetValue('TOTDEBS');
    lTotalCredit  := vTob.GetValue('TOTCRES');
    lSoldeAbsolu  := Abs( vTob.GetValue('TOTDEBS') - vTob.GetValue('TOTCRES'));
  end
  else
  begin
    // Colonne de G_TOTDEBS et G_TOTCRES sont cachées
    if FColDebitS  <> -1 then FListe.ColWidths[FColDebitS]  := -1;
    if FColCreditS <> -1 then FListe.ColWidths[FColCreditS] := -1;

    // Nombre de mois des exercices pour le calcul de la variation
    NombreMois( VH^.Encours.Deb, VH^.Encours.Fin, lPremMois, lPremAnnee, lNbMoisSoldeE);
    NombreMois( VH^.Precedent.Deb, VH^.Precedent.Fin, lPremMois, lPremAnnee, lNbMoisSoldeP);

    lRdSoldeE := vTob.GetValue('TOTDEBE') - vTob.GetValue('TOTCREE');
    lRdSoldeP := vTob.GetValue('TOTDEBP') - vTob.GetValue('TOTCREP');

    lTotalDebit  := vTob.GetValue('TOTDEBE');
    lTotalCredit := vTob.GetValue('TOTCREE');
    lSoldeAbsolu := Abs( vTob.GetValue('TOTDEBE') - vTob.GetValue('TOTCREE'));
  end;

  // Calcul de la Variation
  // Calcul de la Variation
  lTempSoldeE := Abs(lRdSoldeE);
  lTempSoldeP := Abs(lRdSoldeP);

  //if (vTob.GetString('SENS') = 'D') or (vTob.GetString('SENS') = 'M') then
  //begin
  //  if (lRdSoldeE < 0) then lTempSoldeE := -1 * lTempSoldeE;
  //  if (lRdSoldeP < 0) then lTempSoldeP := -1 * lTempSoldeP;
  //end
  //else
  //begin
    if (lRdSoldeE > 0) then lTempSoldeE := -1 * lTempSoldeE;
    if (lRdSoldeP > 0) then lTempSoldeP := -1 * lTempSoldeP;
  //end;

  // GCO - 24/05/2006 - FQ 18197
  if TTypeVariation.Value = 'ABSOLUE' then
    lVariation := (lTempSoldeE - lTempSoldeP)
  else
    lVariation := ((lTempSoldeE / lNbMoisSoldeE) * 12) - ((lTempSoldeP / lNbMoisSoldeP) * 12);

  // Calcul du Pourcentage
  if (lRdSoldeP = 0) and (lRdSoldeE <> 0) then // Division par 0
    lPourcentageVariation := cNS
  else
  begin
    if (lRdSoldeP = 0) and (lRdSoldeE = 0) then
      lPourcentageVariation := 0
    else
    begin
      lPourcentageVariation := ((lVariation) / Abs(lTempSoldeP)) * 100;
      lPourcentageVariation := Arrondi(lPourcentageVariation, V_PGI.OkDecV);
    end;
  end;

  if Abs(lPourcentageVariation) > 300 then
    lStPourcentageVariation := 'NS'
  else
    lStPourcentageVariation := StrFMontant( lPourcentageVariation, 15, V_PGI.OkDecV, '', True) + ' %';

  if FBoCriterePonderation or FBoCritereVariation then
    Result := _CalculRevision
  else
    Result := True;

  if Result then
  begin
    vTob.AddChampSupValeur('SOLDEE', lRdSoldeE, False);
    vTob.AddChampSupValeur('SOLDEP', lRdSoldeP, False);
    vTob.AddChampSupValeur('VARIATION', lVariation, False);
    vTob.AddChampSupValeur('POURCENTAGE', lStPourcentageVariation, False);
    vTob.AddChampSupValeur('XPOURCENTAGE', lPourcentageVariation, False);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Charge ATOBFLISTE dans la grille avec le PutGridDetail
Suite ........ : Retaille la largeur des colonnes de FLISTE
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.RefreshFListe;
begin
  inherited;
  MiseAJourCaptionEcran;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.GetCellCanvasFListe(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickBEffaceRevision(Sender: TObject);
begin
  TTypePonderation.ItemIndex := 2;  // Init à "SOLDE"
  ZCPonderation.ItemIndex    := -1;
  TValeurPonderation.Value   := 0;

  if TTypeVariation.Enabled then
    TTypeVariation.ItemIndex := -1;
  ZCVariation.ItemIndex      := -1;
  TValeurVariation.Value     := 0;
  TNatureVariation.ItemIndex := 1;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/01/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPCONSAUXI.CompareOperateur(vValeur, vValeurAComparer: Double; vStOperateur: string): Boolean;
begin
  if vStOperateur = 'EGA' then
    Result := ( vValeur = vValeurAComparer)
  else
    if vStOperateur = 'IEG' then
      Result := ( vValeur <= vValeurAComparer)
    else
      if vStOperateur = 'INF' then
        Result := ( vValeur < vValeurAComparer)
      else
      if vStOperateur = 'SEG' then
        Result := ( vValeur >= vValeurAComparer)
      else
        if vStOperateur = 'SUP' then
          Result := ( vValeur > vValeurAComparer)
        else
          Result := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnChangeTNatureVariation(Sender: TObject);
begin
  if TNatureVariation.Value = 'POURCENTAGE' then
  begin
    TValeurVariation.NumericType := ntPercentage;
    //TValeurVariation.Text := StrFMontant(Valeur(TValeurVariation.Text)/100, 15, V_PGI.OkDecV, '', True) + '%';
  end
  else
  begin
    TValeurVariation.NumericType := ntGeneral;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/02/2005
Modifié le ... : 07/02/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.TripotageTSituationCompte;
begin
  TSituationCompte.Items[1] := 'Mouvementés sur l''exercice N mais sans mouvements sur N-1';
  TSituationCompte.Items[2] := 'Sans mouvement sur l''exercice N mais mouvementés sur N-1';
  TSituationCompte.Items[3] := 'Non soldés sur l''exercice N mais soldés sur N-1';
  TSituationCompte.Items[4] := 'Soldés sur l''exercice N mais non soldés sur N-1';
  TSituationCompte.Items[5] := 'N''ayant pas le même sens sur l''exercice N que sur N-1';
end;
////////////////////////////////////////////////////////////////////////////////










////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Lance la consultation des écritures du compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickDetailCompte(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  OperationsSurComptes('', FExoDate.Code, '', ATobRow.GetString('T_AUXILIAIRE'), False, '', FBoJustifSolde, TModeSelection.Value );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickDetailCompte2(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  OperationsSurComptes('', FExoDate.Code, '', ATobRow.GetValue('T_AUXILIAIRE'), False, '', True );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... : 04/05/2006
Description .. : Lettrage Manuel d'un compte
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.LETTRAGEMClick(Sender: TObject);
{$IFDEF COMPTA}
var R : RLETTR;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if ATobRow = nil then Exit;

  FillChar(R, Sizeof(R), #0);
  // GCO - 04/05/25006 - FQ 16349
  if T_Collectif.Text <> '' then
    R.General     := T_Collectif.Text
  else
    R.General     := ATobRow.GetString('T_COLLECTIF');
  // FIN GCO  

  R.Auxiliaire    := ATobRow.GetString('T_AUXILIAIRE');
  R.Appel         := tlMenu;
  R.CritDev       := V_PGI.DevisePivot;
  R.DeviseMvt     := V_PGI.DevisePivot;
  R.GL            := nil;
  R.CritMvt       := ' AND ( E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL" ) ';
  CEtudieModeLettrage(R) ;
  LettrageManuel(R, True, FAction);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Lettrage Automatique
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.LettrageAClick(Sender: TObject);
begin
{$IFDEF COMPTA}
  if ATobRow = nil then Exit;
  if not CEstAutoriseDelettrage(False, False) then Exit;
  RapprochementAuto( ATobRow.GetString('T_COLLECTIF'), ATobRow.GetString('T_AUXILIAIRE') );
  RefreshPclPge;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/11/2004
Modifié le ... : 28/07/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.DelettreMvtExoRef(Sender: TObject);
begin
  if not CEstAutoriseDelettrage(False, False) then Exit ;

  if PgiAsk('Attention, vous allez supprimer tout le lettrage de l''exercice de référence.' + #13 + #10 +
            'Confirmez vous le traitement ?', 'Delettrage complet pour l''exercice de référence') = MrYes then
  begin
    TRANSACTIONS(TransactionDelettreMvtExoRef,1) ;
    RefreshPclPge;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.TransactionDelettreMvtExoRef ;
begin
{$IFDEF COMPTA}
  if ATobRow = nil then Exit;
  CExecDelettreMvtExoRef(ATobRow.GetString('T_COLLECTIF'), ATobRow.GetString('T_AUXILIAIRE'));
  RefreshPclPge;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.DelettreTotal(Sender: TObject);
begin
  if not CEstAutoriseDelettrage(False, False) then Exit;

  if PGIAsk('Attention, vous allez supprimer tout le lettrage du compte.' + #13 + #10 +
            'Confirmez-vous le traitement ?', 'Delettrage automatique') = mrYes then
  begin
    TRANSACTIONS(TransactionDeLettrageTotal, 1);
    RefreshPclPge;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/11/2004
Modifié le ... : 28/07/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.TransactionDeLettrageTotal ;
begin
{$IFDEF COMPTA}
  if ATobRow = nil then Exit;
  CExecDeLettrage( ATobRow.GetString('T_COLLECTIF'), ATobRow.GetString('T_AUXILIAIRE'));
  RefreshPclPge;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
//---------------------------- Traitements Comptables --------------------------
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Saisie d'une Piece
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.SaisieEcrClick(Sender: TObject);
var OldSC: Boolean;
begin
  OldSC := VH^.BouclerSaisieCreat;
  VH^.BouclerSaisieCreat := False;
  MultiCritereMvt(taCreat, 'N', False);
  VH^.BouclerSaisieCreat := OldSC;
  RefreshPclPge;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... :   /  /
Description .. : Saisie d'un Bordereau
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.SaisieBorClick(Sender: TObject);
begin
{$IFDEF COMPTA}
  SaisieFolio(taModif);
  RefreshPclPge;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
//---------------------------- Utilitaires -------------------------------------
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/11/2002
Modifié le ... :   /  /
Description .. : Affiche la fenêtre des paramètres du compte général
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickParamCollectif(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  FicheGene( nil, '', ATobRow.GetString('TCOLLECTIF'), taModif, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/09/2002
Modifié le ... : 29/11/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickParamTiers(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  FicheTiers( nil, '', ATobRow.GetString('T_AUXILIAIRE'), taModif, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickCompteRib(Sender: TObject);
var lOkExiste : Boolean;
begin
  if ATobRow = nil then Exit;

  lOkExiste := ExisteSQL('SELECT R_AUXILIAIRE FROM RIB WHERE R_AUXILIAIRE="' +
               ATobRow.GetString('T_AUXILIAIRE') + '"');

  if (not lOkExiste) then
  begin
    if FAction = TaConsult then
      PgiInfo('Le compte n''a pas de RIB associé.', 'Consultation des auxiliaires')
    else
    begin
      if PgiAsk('Le compte n''a pas de RIB associé.', 'Voulez-vous le créer ?') = MrYes then
      begin
        FAction := TaCreatOne;
        FicheRIB_AGL( ATobRow.GetValue('T_AUXILIAIRE'), FAction, False, '', False);
      end;
    end;
  end
  else // Ouverture du rib
    FicheRIB_AGL( ATobRow.GetValue('T_AUXILIAIRE'), FAction, False, '', False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickNoteTravail(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  CPLanceFiche_CPNoteTravail( ATobRow.GetString('TCOLLECTIF') );
  RefreshPclPge;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickTableauVariation(Sender: TObject);
begin
  if ATobRow = nil then Exit;
  CPLanceFiche_CPTableauVar( ATobRow.GetString('TCOLLECTIF') );
  RefreshPclPge;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickRechercheEcritures(Sender: TObject);
begin
  if CtxPcl in V_Pgi.PgiContexte then
    CPLanceFiche_CPRechercheEcr( False )
  else
    MultiCritereMvt(taConsult,'N', False );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickFeuilleExcel(Sender: TObject);
var lStCycle : string;
begin
  lStCycle := GetColonneSql('GENERAUX', 'G_CYCLEREVISION', 'G_GENERAL = "' + ATobRow.GetString('T_COLLECTIF') + '"');
  CPLanceFiche_CPIMPFICHEXCEL ('FTS;' + ATobRow.GetString('T_COLLECTIF') + ';' + lStCycle);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickGCD( Sender : TObject) ;
begin
  if ATobRow = nil then Exit;
  CPLanceFiche_GCDOperation('' , 'AUX;CREANCE;' + ATobRow.GetString('T_AUXILIAIRE') + ';' + ATobRow.GetString('T_COLLECTIF') + ';' + ATobRow.GetString('SOLDEE') ) ;
  RefreshPclPge ;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickProgrammeTravail(Sender: TObject);
begin
  CPLanceFiche_CPRevProgTravail( VH^.EnCours.Code );
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickMemoCycle(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPRevDocTravaux( lTob.GetString('TCOLLECTIF'), '', VH^.EnCours.Code, 0 );
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickMemoObjectif(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPRevDocTravaux( lTob.GetString('TCOLLECTIF'), '', VH^.EnCours.Code, 1 );
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickMemoSynthese(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPRevDocTravaux( lTob.GetString('TCOLLECTIF'), '', VH^.EnCours.Code, 2 );
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickMemoMillesime(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPRevDocTravaux( lTob.GetString('TCOLLECTIF'), '', VH^.EnCours.Code, 3 );
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickMemoCompte(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe);
  if lTob = nil then Exit;
  CPLanceFiche_CPRevDocTravaux( lTob.GetString('TCOLLECTIF'), '', VH^.EnCours.Code, 4);
  RefreshPclPGE;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickAPG(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 1);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickSCY( Sender : TObject );
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 2);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickEXP( Sender : TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 3);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.OnClickDossPermanent(Sender: TObject);
var lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Dossier permanent');
  lObj := unassigned;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CPCONSAUXI.OnClickDossAnnuel(Sender: TObject);
var lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Dossier annuel');
  lObj := unassigned;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEf COMPTA}
procedure TOF_CPCONSAUXI.OnClickRechercheDoc(Sender: TObject);
var lObj : Variant;
begin
  lObj := CreateOleObject( 'CegidPgi.CegidPgi' );
  lObj.ZoomDP(V_PGI.NoDossier,'Recherche documentaire');
  lObj := unassigned;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
//---------------------------- Fonctions Complémentaires -----------------------
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickCommentaire(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  FicheTiers( nil, '', lTob.GetValue('T_AUXILIAIRE'), taModif, 2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickCumulsCollectif(Sender: TObject);
var lTob : TOB ;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CumulCpteMensuel(fbGene, lTob.GetValue('TCOLLECTIF'), '', VH^.CPExoRef);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/09/2002
Modifié le ... : 07/12/2004
Description .. : Cumuls du compte auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickCumulsAuxiliaire(Sender: TObject);
var lTob : Tob;
begin
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  CumulCpteMensuel(fbAux, lTob.GetValue('T_AUXILIAIRE'), lTob.GetValue('T_LIBELLE'), VH^.CPExoRef);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickTableChancellerie(Sender: Tobject);
var lTob : Tob;
begin
  lTob := GetO(FListe) ;
  if (lTob = nil) then Exit;

  if (lTob.GetString('LADEVISE') <> V_PGI.DevisePivot) or
     (lTob.GetString('MULTIDEVISE') = 'X') then
  begin
    if (lTob.GetValue('MULTIDEVISE') = 'X') {or
       (EstMonnaieIN(lTob.GetValue('LADEVISE')))} then
      FicheDevise(lTob.GetValue('LADEVISE'), FAction, False)
    else
      FicheChancel(lTob.GetValue('LADEVISE'), False, 0, FAction, True);
  end;
end;

(*
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickAnalyseComplementaire(Sender: TObject);
begin
  AGLLanceFiche('CP', 'CPCONSULTREVIS', '', '', '');
end;*)

////////////////////////////////////////////////////////////////////////////////
//---------------------------- Editions ----------------------------------------
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.PrepareArgumentGL(vBoJustifSolde: Boolean);
begin
  // Exercice
  ACritEdt.CritEdt.Exo.Code := FExoDate.Code;
  // Début de la Date Comptable (1)
  ACritEdt.CritEdt.Date1 := FExoDate.Deb;
  // Fin de la Datecomptable (2)
  ACritEdt.CritEdt.Date2 := FExoDate.Fin;
  // Justificatif de Solde
  ACritEdt.CritEdt.GL.EnDateSituation := vBoJustifSolde;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... : 06/09/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickJustifSolde(Sender: TObject);
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if (lTob = nil) then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( True );
  TheData := ACritEdt;
  ACritEdt.CritEdt.Cpt1 := lTob.GetString('T_AUXILIAIRE');
  ACritEdt.CritEdt.Cpt2 := lTob.GetString('T_AUXILIAIRE');
  CPLanceFiche_CPGLAUXI('');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... : 06/09/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickGLAuxi(Sender: TObject);
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if (lTob = nil) then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1  := lTob.GetString('T_AUXILIAIRE');
  ACritEdt.CritEdt.Cpt2  := lTob.GetString('T_AUXILIAIRE');
  TheData := ACritEdt;
  CPLanceFiche_CPGLAuxi('');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... : 06/09/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickGLAuxiParGene(Sender: TObject);
{$IFDEF COMPTA}
var lTob : Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if (lTob = nil) then Exit;

  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentGL( False );
  ACritEdt.CritEdt.Cpt1  := lTob.GetString('T_AUXILIAIRE');
  ACritEdt.CritEdt.Cpt2  := lTob.GetString('T_AUXILIAIRE');
  ACritEdt.CritEdt.SCpt1 := '';
  ACritEdt.CritEdt.SCpt2 := '';
  TheData := ACritEdt;
  CPLanceFiche_CPGLAuxiParGene('');
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFDEF COMPTA}
procedure TOF_CPCONSAUXI.PrepareArgumentBAL;
begin
  // Exercice
  ACritEdt.CritEdt.Exo.Code := FExoDate.Code;
  // Début de la Date Comptable (1)
  ACritEdt.CritEdt.Date1 := FExoDate.Deb;
  // Fin de la Datecomptable (2)
  ACritEdt.CritEdt.Date2 := FExoDate.Fin;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickBALAuxi(Sender: TObject);
begin
{$IFDEF COMPTA}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1 := T_Auxiliaire.Text;
  ACritEdt.CritEdt.Cpt2 := T_Auxiliaire_.Text;
  ACritEdt.CritEdt.NatureCpt := T_NATUREAUXI.Text ; // FQ 15844 SBO 27/09/2005
  TheData := ACritEdt;
  CPLanceFiche_BalanceAuxiliaire;
  TheData := nil;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.OnClickBALAuxiParGene(Sender: TObject);
{$IFDEF COMPTA}
var lTob: Tob;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lTob := GetO(FListe) ;
  if lTob = nil then Exit;
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  PrepareArgumentBAL;
  ACritEdt.CritEdt.Cpt1  := lTob.GetString('T_AUXILIAIRE');
  ACritEdt.CritEdt.Cpt2  := lTob.GetString('T_AUXILIAIRE');
  ACritEdt.CritEdt.SCpt1 := '';
  ACritEdt.CritEdt.SCpt2 := '';
  TheData := ACritEdt;
  CPLanceFiche_BalanceAuxiGen;
  TheData := nil;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 23/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONSAUXI.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


Initialization
  registerclasses ( [ TOF_CPCONSAUXI ] ) ;
end.
