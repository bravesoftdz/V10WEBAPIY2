unit MDispCCSTD;

interface

uses
  HEnt1,
  Forms,
  sysutils,
  HMsgBox,
  Classes,
  Controls,
  HPanel,
  UIUtil,
  ImgList,
{$IFDEF EAGLCLIENT}
  eTablette,
  MaineAGL, //
  MenuOLX,  //
{$ELSE}
  Tablette,
  FE_Main,
  EdtREtat,
  uEdtComp,
  AglInitPlus,
  AGLInit,
{$ENDIF}
  Windows,
  Messages,
  Graphics,
  Dialogs,
  RecupPlantype,
  Menus,
  ComCtrls,
  HStatus,
  HMenu97,
  Hctrls,
  CalcOle,
  UtilPGI,
  EntPGI,
  HTB97,
  Hout,
  uTOB;

procedure InitApplication;

type
  TFMenuDisp = class(TDatamodule)
    LargeImages: TImageList;
  end;

var
  FMenuDisp: TFMenuDisp;

procedure AfterChangeModule(NumModule: integer);
procedure AssignZoom;
function LaunchAuto : boolean;
procedure InitSerialisation;

implementation

uses
  FichComm,
  Paramsoc,
  Utilsoc,
  RefAuto_TOM,       // ParamLibelle
{$IFDEF EAGLCLIENT}
  Pays_TOM,
  CPRegion_TOF,
  CPCodePostal_TOF,
  SUIVCPTA_TOM,      // CPLanceFiche_Scenario
  CPMulAna_TOF,      // MultiCritereAna
{$ELSE}
  MenuOLG,
  Pays,
  Region,
  CodePost,
  Rupture,           // PlanRupture
  Scenario,          // ParamScenario
  MulAna,            // MultiCritereAna
{$ENDIF}
  CPAxe_TOM,               // FicheAxeAna('') ;
  Devise_TOM,        //
  Corresp_TOM,       // CCLanceFiche_Correspondance
  CPTABLIBRELIB_TOF, // CPLanceFiche_ParamTablesLibres
  CPCHOIXTALI_TOF,   // CPLanceFiche_ChoixTableLibre
  Saisutil,          //
  AssitStd,          //
  UtotVentilType,    // ParamVentilType
  GalEnv,            //
  GalOutil,          // RajouteCaptionDossier
  Ent1,              //
  CritEdt,           //
  uLibStdCpta,       // TTraitementFiltre
  uTofStdChoixPlan,  // ChargeListeDossierType
  EdtQR,             //
  CpteUtil,          // GetDate
  LGCOMPTE_TOF,      // CCLanceFiche_LongueurCompte
  uTofCPGLGENE,      // CPLanceFiche_CPGLGENE
  uTofCPGLAUXI,      // CPLanceFiche_CPGLAUXI
  uTofConsGene,      // CPLanceFiche_CPCONSGENE
  uTofEtatsChaines,  // CPLanceFiche_EtatsChaines
  uTofCPJalEcr,      // CPLanceFiche_CPJALECR
  uTofCPMulMvt,      // MultiCritereMvt
  CPBALGEN_TOF,      // CPLanceFiche_BalanceGeneral
  CPBALAUXI_TOF,     // CPLanceFiche_BalanceAuxiliaire
  CRITSYNTH,         // EtatPCL
  uTofConsEcr,       // OperationsSurComptes
  uTofCPMulGuide,    // ParamGuide
  MulLettr,          // LanceLettrage
  CPGeneraux_TOM,    // FicheGene
  CPTiers_TOM,       // FicheTiers
  CPJournal_TOM,     // FicheJournal
  CPSection_TOM,     // FicheSection
  AMEdition_TOF,     // AMLanceFiche_Edition
  AMAnalyse_TOF,     // AMLanceFiche_StatPrevisionnel
  CPBalAnal_TOF,     // CPLanceFiche_BalanceAnalytique
  CPBalGenAnal_TOF,  // CPLanceFiche_BalanceGenAnal
  CPBalAnalGen_TOF,  // CPLanceFiche_BalanceAnalGen
  CPBalGenAuxi_TOF,  // CPLanceFiche_BalanceGenAuxi
  CPBalAuxiGen_TOF,  // CPLanceFiche_BalanceAuxiGen
  uTofCPGLAna,       // CPLanceFiche_CPGLGANA
  TofBalAgee,        // CPLanceFiche_BalanceAgee, CPLanceFiche_BalanceVentilee
  TofEdJal,          // Journal Centralisateur
  CroiseAxeTof,      // Analyse Statistique
  NatCpte_TOM,       // FicheNatCpte
  CPMulRub_TOF,      // MultiCritereRubriqueV2
  CPSuppRub_TOF,     // SuppressionRubriqueV2
  CPControleRub_TOF, // ControleRubrique
  CPQRTRubrique_TOF, // ImpRubrique
  Rubrique_TOM,      // FicheRubrique
  CPTVATPF_TOF,      // ParamTvaTpf
  uTOFStdEnregPlan,  // CPLanceFiche_EnregistrementSurStandard
  uTofRechercheEcr,  // CPLanceFiche_CPRechercheEcr(False)
  CPRevMulParamCycl_TOF, // CPLanceFiche_CPREVMULPARAMCYCL
  uImportBOB_TOF,    // CPLanceFiche_CPIMPFICHEXCEL (TypS : string);
  CPLOIVENTILMUL_TOF,// CPLanceFiche_LoiVentilMul
  CPMASQUEMUL_TOF,   // CPLanceFiche_MasqueMul
  CPControleLiasse_TOF // CPLanceFiche_CPControleCycle ( vStArgument : string );
  ,BobGestion
  ;


{$R *.DFM}

var
  { CA - 12/07/2005 - Variable globale pour indiquer si un blocage a empêché l'entrée dans l'application }
  gBlocage : boolean;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 29/11/2007
Modifié le ... :   /  /
Description .. : Vérification des maquettes des états de synthèse : cas ou
Suite ........ : des maquettes sont doublées avec NODOSSIER=""
Mots clefs ... :
*****************************************************************}
procedure VerificationEtatSynth;
  function _MajNoDossierEtatSynth (pstGuid : string) : boolean;
  begin
    Result := (ExecuteSQL ('UPDATE YFILESTD SET YFS_NODOSSIER="000000" '+
      ' WHERE YFS_FILEGUID="'+pstGuid+'"')=1);
  end;
  function _DetruireEtatSynth (pstGuid : string) : boolean;
  begin
    Result := False;
    if (ExecuteSQL ('DELETE FROM YFILESTD WHERE YFS_FILEGUID="'+pstGuid+'"') >=1) then
    begin
      ExecuteSQL ('DELETE FROM NFILES WHERE NFI_FILEGUID="'+pstGuid+'"');
      ExecuteSQL ('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID="'+pstGuid+'"');
      Result := True;
    end;
  end;
var
  lT : TOB;
  i : integer;
begin
  // Est-ce qu'il existe des états de synthèse avec YFS_NODOSSIER="" ?
  if ExisteSQL ('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" '
                  +' AND YFS_CRIT1="ETATSYNTH" AND (YFS_NODOSSIER="" OR YFS_NODOSSIER IS NULL)') then
  begin
    // Chargement des états qui posent problème
    lT := TOB.Create ('', nil, - 1);
    try
      lT.LoadDetailDBFromSQL('YFILESTD','SELECT * FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" '
                  +' AND YFS_CRIT1="ETATSYNTH" AND (YFS_NODOSSIER="" OR YFS_NODOSSIER IS NULL)');
      for i := 0 to lT.Detail.Count - 1 do
      begin
        // Est-ce que l'état existe par ailleurs dans la base
        if ExisteSQL('SELECT * FROM YFILESTD WHERE YFS_CODEPRODUIT="COMPTA" '
          +' AND YFS_CRIT1="ETATSYNTH" AND YFS_NODOSSIER<>"" AND YFS_NOM="'+lT.Detail[i].GetString('YFS_NOM')+'"'
          +' AND YFS_CRIT2="'+lT.Detail[i].GetString('YFS_CRIT2')+'"'
          +' AND YFS_CRIT3="'+lT.Detail[i].GetString('YFS_CRIT3')+'"') then
        begin
          // On détruit l'état qui a NODOSSIER=""
          _DetruireEtatSynth (lT.Detail[i].GetString('YFS_FILEGUID'));
        end else
        begin
          // On met à jour le numéro de dossier
          _MajNoDossierEtatSynth(lT.Detail[i].GetString('YFS_FILEGUID'));
        end;
      end;
    finally
      lT.Free;
    end;
  end;
end;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var
    Texte: string;
    lSt, St : string;
begin
  case Num of
    10: //Apres connection
      begin
        // GCO - 13/09/2004 - FQ 12441
        if EstBloque( 'nrStandard',True) then
        begin
          HShowMessage('0;' + NomHALLEY + ';Fonction bloquée, recommencez ultérieurement...;W;O;O;O', '', '');
          gBlocage := True;
          FMenuG.FermeSoc;
          FMenuG.Quitter;
          Exit;
        end else
        begin
          Bloqueur( 'nrStandard',True);
          gBlocage := False;
        end;

        // AjOUT ME 12/06/01 pour supprimer les plans ref > 100 et < 21 si défini en STD
        SuppressionPlanNonreference;
        // On bascule les standards en euro
        BasculeStandardEuro;
        if not StandardEnEuro then
        begin
          FMenuG.ForceClose := True;
          FMenuG.Close;
          Exit;
        end;
        // CA - 29/04/2002 - Le standard est systématiquement créé en euro
        InitDossierStandardEuro;
        { Création d'exercices fictifs }
        InitTableExerciceStandard;
        // Rechargement des paramsoc
        RechargeParamSoc;
        // Vérification des états de synthèse
        VerificationEtatSynth;
      end;

    11: begin
          //Après deconnection
        end;

    12: begin
          // Interdiction de lancer en direct
          // GCO - 30/11/2004 - FQ 14991
          // GCO - 10/05/2005 - Correction PGE car V_PGI.RunFromLanceur = TRUE
          if (not (CtxPCl in V_Pgi.PgiContexte)) or (not V_PGI.RunFromLanceur) then
          begin
            FMenuG.FermeSoc;
            FMenuG.Quitter;
            Exit;
          end;
          VH^.ModeSilencieux := True; // Avant connection ou seria

          // GCO - 05/06/2007 - Test de la séria RIC
          InitSerialisation;
        end;
        
    13 : // à la fermeture
      begin
        if V_PGI.Menucourant = 64 then
        begin
          Texte := 'Attention,Si vous n''avez pas enregistré vos modifications,'
            + 'elles ne seront pas prises en compte#13#10'
            + 'Voulez-vous enregistrer votre dossier type ?';
          // CA - 28/11/2001
          if HShowMessage('0;Enregistrer le dossier type;' + Texte +
            ';Q;YN;N;N;', '', '') = mrYes then
            begin
            // LanceAssistantEnreg;
              V_PGI.ZoomOLE := TRUE;
              CPLanceFiche_EnregistrementSurStandard( NumPlanCompte);
              V_PGI.ZoomOLE := False;
            end;
        end;
        if not gBlocage then
          Bloqueur('nrStandard',False) ;
        DeBlocageMonoPoste(True);
      end;

    15: ;
    16:
      begin
        (*
        {$IFNDEF EAGLCLIENT}
          { Importation des bobs CCS5 }
          BOB_IMPORT_PCL('CCS5');
          { Importation des BOBS communes en Entreprise }
          if not (ctxPCL in V_PGI.PGIContexte) then BOB_IMPORT_PCL('COMM',True);
          // PGI_IMPORT_BOB('CSTD'); CA - 25/06/2007 - Désormais les bobs standards sont identiques à la comptabilité
        {$ENDIF}
        *)
        St := 'CCS5';
        lSt := 'Paramétrage Comptabilité';
        if not (ctxPCL in V_PGI.PGIContexte) then begin
          St := St+';COMM';
          lSt := lSt + ';Paramétrage Commun';
        end;
        BOB_IMPORT_PCL_STD (St,lSt);
      end;
    100:
    begin
      RetourForce := True; // execution depuis lanceur
    end;

    // restauration depuis SISCOII
    64030 : RecupPlansisco;

    64710 : begin // Chargement d'un standard
              FMenuG.ChoixModule;
              V_PGI.ZoomOLE := TRUE;
              // FQ 19295 - CA - Proposition d'enregistrement du dossier type
              if NumPlanCompte <> 0 then
              begin
                Texte := 'Attention,Si vous n''avez pas enregistré vos modifications,'
                         + 'elles ne seront pas prises en compte#13#10'
                         + 'Voulez-vous enregistrer votre dossier type '
                         + IntToStr(NumPlanCompte) + '?';
                if PGIAsk (Texte, 'Choix du dossier type') = mrYes then
                   CPLanceFiche_EnregistrementSurStandard( NumPlanCompte );
              end;
              AGLLanceFiche('CP','STDCHOIXPLAN','','','');
              V_PGI.ZoomOLE := FALSE;
              if NumPlanCompte = 0 then FMenuG.ChangeModule(69);
            end;

    64711 : begin //Passage du module 64 au module 69 ==> proposition d'enregistrement du standard
              FMenuG.ChoixModule;
              if NumPlanCompte <> 0 then
              begin
                V_PGI.ZoomOLE := TRUE;
                Texte := 'Attention,Si vous n''avez pas enregistré vos modifications,'
                         + 'elles ne seront pas prises en compte#13#10'
                         + 'Voulez-vous enregistrer votre dossier type '
                         + IntToStr(NumPlanCompte) + '?';
                if PGIAsk (Texte, 'Choix du dossier type') = mrYes then
                  CPLanceFiche_EnregistrementSurStandard( NumPlanCompte );
                RajouteCaptionDossier('Dossier type : ' + IntToStr(NumPlanCompte));
                V_PGI.ZoomOLE := FALSE;
              end;
            end;

    64720 : begin // Enregistrement du standard
              if NumPlanCompte <> 0 then
              begin
                FMenuG.ChoixModule;
                V_PGI.ZoomOLE := TRUE;
                CPLanceFiche_EnregistrementSurStandard( NumplanCompte );
                RajouteCaptionDossier('Dossier type : ' + IntToStr(NumPlanCompte));
                V_PGI.ZoomOLE := FALSE;
              end;
            end;

    // MultiCritère GENERAUX
    7112 : begin
             V_PGI.ExtendedFieldSelection := '';
             AGLLanceFiche('CP', 'MULGENERAUX', '', '', 'C;7112000');
           end;

    // MultiCritère TIERS
    7145 : begin
             V_PGI.ExtendedFieldSelection := '';
             AGLLanceFiche('CP', 'MULTIERS', '', '', 'C;7145000');
           end;

    // MultiCritère JOURNAL       
    7211 : begin
             V_PGI.ExtendedFieldSelection := '';
             AGLLanceFiche('CP', 'MULJOURNAL', '', '', 'C;7211000');
           end;

    // GCO - 25/05/2004 7241: SelectGuide('', '', '', '', True);

    (* GCO - 25/05/2004
    7244:
      begin
        V_PGI.ListeByUser := FALSE; // ajout me 05/06/01
        MultiCritereMvt(taCreat, 'N', False);
      end;
    *)

    //societe
    1104: ParamSociete(False,
          'SCO_COORDONNEES;SCO_CPREVISION;SCO_DATESDIVERS;SCO_CPFOLIO;SCO_CPEDI;SCO_OLAP',
          'SCO_PARAMETRESGENERAUX', '', ChargeSocieteHalley, ChargePageSoc,
          SauvePageSoc, InterfaceSoc, 1105000);

    // GCO - 04/05/2005 - Menu : Paramètrages - Société - TVA
    1165 : ParamTable('TTREGIMETVA',taCreat,1165000,PRien) ; // Régimes Fiscaux
    1170 : ParamTvaTpf( True  );                             // TVA par régime fiscal
    1175 : ParamTvaTpf( False );                             // TPF par régime fiscal
    // FIN GCO

    1110: FicheEtablissement_AGL(taModif);

    7352: CPLanceFiche_MulGuide('','','');  //ParamGuide('', 'NOR', taModif);
    7355: ParamScenario('', '');
    17420 : CPLanceFiche_MasqueMul('','','ACTION=MODIFICATION;CMS_TYPE=SAI') ;
    // changement longueur compte
    3150: CCLanceFiche_LongueurCompte;

    1185: FicheModePaie_AGL('');
    1190: FicheRegle_AGL('', FALSE, taModif);

    //    Comptes de correspondance
    1325: CCLanceFiche_Correspondance('GE'); // Plan de correspondance Généraux
    1330: CCLanceFiche_Correspondance('AU'); // Plan de correspondance Auxiliaire

    // Qualifiants des quantités
    64241 : ParamTable('ttQualUnitMesure', taCreat, 1190030, PRien);

{$IFDEF EAGLCLIENT}

{$ELSE}
    //    rupture des comptes
    1375: PlanRupture('RUG');
    1380: PlanRupture('RUT');
{$ENDIF}

    //    Libelle automatique
    1425 : ParamLibelle;

    //    table libre
    1194 : CPLanceFiche_ParamTablesLibres;  // Tables libres // Personalisation
    1196 : CPLanceFiche_ChoixTableLibre;    // Tables libres // Saisie

    //    Etat libres comptabilité
    // MBAMF
    //1270 : EditEtatS5S7('E', 'UCO', '', True, nil, '', '');

    // standard d'éditions
    7445 : CPLanceFiche_BalanceGeneral ;
    7415 : CPLanceFiche_CPGLGENE('');
    7435 : EtatPCL(esbil);
    7434 : EtatPCL(esCR);
    7453 : EtatPCL(esSIG);
    7448 : CPLanceFiche_BalanceAuxiliaire ;
    7418 : CPLanceFiche_CPGLAUXI('');
    7560 : CPLanceFiche_CPGLAUXI(''); //GLAuxiliaireL;

    // GCO MBAMF - plus de raison d'existe puisque integre dans le GL
    //7535 : JustSolde;

      // Journal divisionnaire
    7602 : OperationsSurComptes('', '', '');

    // Recherche des écritures
    7259 : CPLanceFiche_CPRechercheEcr(False);

    // Etats chainés
    7444 : CPLanceFiche_EtatsChaines('CP');

    7508 : LanceLettrage;

    //      7514 : RapprochementAuto('','') ;

    // Analytique
    1105 : FicheAxeAna('');


    7178 : begin
             V_PGI.ExtendedFieldSelection := '';
             AGLLanceFiche('CP', 'MULSECTION', '', '', 'C;7178000');
           end;

    1460 : ParamVentilType;
    7383 : CPLanceFiche_LoiVentilMul('', '', 'CLV_LOITYPE=LOI');
    7370 :
      begin
        // GCO 25/05/2204 toujours a True desormais
        //V_PGI.ListeByUser := FALSE; // ajout me 05/06/01
        MultiCritereAna(taModif);
      end;

    7113 : begin
             V_PGI.ExtendedFieldSelection := '1';
             AGLLanceFiche('CP', 'MULGENERAUX', '', '', 'C;7112000');
           end;
    //    assistant
    3140:
      begin
        LanceAssistantStd(FALSE);
      end;
    { Journal des écritures }
    7394 : CPLanceFiche_CPJALECR ;

    // GCO - 21/07/2004 - Brancement des nouvelles Editions analytiques
    7451 : CPLanceFiche_BalanceAnalytique; // Balance analytique
    7466 : CPLanceFiche_BalanceGenAnal ;   // Balance par Comptes Généraux
    7469 : CPLanceFiche_BalanceAnalGen ;   // Balance Analytique par Général

    7421 : CPLanceFiche_CPGLANA;          // Grand livre analytique

    7787 : MultiCritereRubriqueV2(taModif, False);
    7790 : SuppressionRubriqueV2(False);

    7796 : ImpRubrique('', False, False);


    7775 : ParamTable('CPRUBFAMILLE', taCreat, 7775000, PRien, 17);
    7811 : ParamTable('TTCONSTANTE', taCreat,0, PRien, 3, 'Constantes'); // Constantes;
    7802 : ControleRubrique('');

    // Consultation des comptes
    52110 : CPLanceFiche_CPCONSGENE();

    // Maquettes
    69110: AGLLanceFiche('CP', 'STDMAQ', 'CR', '', 'ACTION=MODIFICATION;CR');
    69210: AGLLanceFiche('CP', 'STDMAQ', 'SIG', '', 'ACTION=MODIFICATION;SIG');
    69310: AGLLanceFiche('CP', 'STDMAQ', 'BIL', '', 'ACTION=MODIFICATION;BIL');
    {    69110 : // CR
              CreationMaquette (1);
        69120 : // suppression
        AGLLanceFiche('CP','STDMAQUETTE','CR', '','OK');
        69210 : // SIG
              CreationMaquette (2);
        69220 : // suppression
        AGLLanceFiche('CP','STDMAQUETTE','SIG', '','OK');

        69310 : // BIL
              CreationMaquette (3);

        69320 : // suppression
        AGLLanceFiche('CP','STDMAQUETTE','BIL', '','OK');
              }
    // GCO - 19/07/2004 Ajout des éditions des immo
    2150  : CPLanceFiche_EtatsChaines('AM');
    2544  : AMLanceFiche_Edition ( 'ITE', 'ITE' );
    20232 : AMLanceFiche_Edition ( 'ITN', 'ITN' );
    20234 : AMLanceFiche_Edition ( 'ITF', 'ITF' );
    20236 : AMLanceFiche_Edition ( 'ITR', 'ITR' );
    20238 : AMLanceFiche_Edition ( 'ITO', 'ITO' );
    2541  : AMLanceFiche_Edition ( 'ILS', 'ILS' ); // PCL
    2542  : AMLanceFiche_Edition ( 'IAC', 'IAC' );
    2543  : AMLanceFiche_Edition ( 'ISO', 'ISO' );
    20222 : AMLanceFiche_Edition ( 'IMU', 'IMU' );
    2549  : AMLanceFiche_Edition ( 'IPM', 'IPM' );
    20262 : ; // Je sais pas quoi en faire
    20252 : AMLanceFiche_Edition ( 'ITP', 'ITP' );
    2547  : AMLanceFiche_Edition ( 'IPR', 'PRE' );
    20272 : AMLanceFiche_Edition ( 'IVE', 'IVE' );
    20274 : AMLanceFiche_Edition ( 'IVL', 'IVL' );
    20276 : ; // A faire avec RemoveFromMenu
    20278 : ; // A faire avec RemoveFromMenu
    2180  : AMLanceFiche_StatPrevisionnel;

    // GCO - 18/08/2004 - FQ 14094
    7457  : CPLanceFiche_BalanceGenAuxi ;
    7460  : CPLanceFiche_BalanceAuxiGen ;
    7547  : CPLanceFiche_BalanceAgee;
    7556  : CPLanceFiche_BalanceVentilee;

    7427  : CPLanceFiche_CPGLGENEPARAUXI;
    7430  : CPLanceFiche_CPGLAUXIPARGENE;
    // Journal centralisateur par période-journal
    7412  : AGLLanceFiche('CP', 'EPJALGEN', '', 'JCD', 'JAC') ;

    // Analyses statistiques
    8405  : AGLLanceFiche('CP', 'CPECRITURE_TOBV', '', '', '') ;
    8410  : AGLLanceFiche('CP', 'CPANALYTIQ_TOBV', '', '', '') ;
    // FIN - GCO - 18/08/2004 - FQ 14094

    // GCO 16/03/2005 Ajout du Module Révision
    66110 : AGLLanceFiche('CP', 'CPREVMULPLAN', '', '', '') ;
            //ParamTable('CPPLANREVISION', taCreat, 0, PRien, 3, 'Plans de révision') ;

    66120 : CPLanceFiche_CPREVMULPARAMCYCL('');
    66130 : CPLanceFiche_CPCONTROLECYCLE('');
    // FIN GCO

    // FTS
    66141 : CPLanceFiche_CPIMPFICHEXCEL ('FTS');
    66142 : CPLanceFiche_CPIMPFICHEXCEL ('EXL');
  end;

end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range:
  string);
begin
  case Num of
    1: FicheGene(nil, '', LeQuel, Action, 0);
    2: FicheTiers(nil, '', LeQuel, Action, 1);
    4: FicheJournal(nil, '', Lequel, Action, 0);
    7701 : ParametrageRubrique(Lequel,Action, CtxRubrique) ;
    64101:
      //     AGLLanceFiche('CP','STDPLCPTE','','','');
      AGLLanceFiche('CP', 'STDCHOIXPLAN', '', '', '');
  end;
end;

Procedure RemoveItemTN(i : Integer) ;
Var HOI : THOutItem ;
    TN : TTreeNode ;
BEGIN
  HOI:=FMenuG.OutLook.GetItemByTag(i) ; If HOI<>NIL Then FMenuG.OutLook.RemoveItem(HOI) ;
  TN:=GetTreeItem(FMenuG.TreeView,i) ; If TN<>NIL Then TN.Delete ;
END ;

procedure TripotageMenu(NumModule: integer);
begin
  if NumModule = 64 then
  begin
    RemoveItemTN(7560);
    RemoveItemTN(7535);
    RemoveItemTN(7355); // GCO - 18/08/2004 - FQ 14145

    // GCO - 24/09/2004 - FQ 13231
    //Suppression des éditions qui ne sont pas encore implémentées }
    { Etat des sorties }
    FMenuG.RemoveItem(20224);
    { Branche crédit-bail }
    FMenuG.RemoveItem(20262);
    FMenuG.RemoveItem(-20260);
    { Etat comparatif des bases TP }
    FMenuG.RemoveItem(20254);
    { Etiquettes }
    FMenuG.RemoveItem(20276);
    { Fiche Immobilisation }
    FMenuG.RemoveItem(20278);
    { Plans de ruptures }
    FMenuG.RemoveItem(1375);
    FMenuG.RemoveItem(1380);
    FMenuG.RemoveItem(-10);        
  end;
end;

procedure AfterProtec(sAcces: string);
begin
  V_PGI.VersionDemo := False;
{$IFDEF GIL}
  VH^.OkModRIC      :=True;
{$ELSE}
  VH^.OkModRIC      :=(sAcces[1]='X');
{$ENDIF}
end;

procedure AfterChangeModule(NumModule: integer);
begin
  UpdateSeries;
  V_PGI.ZoomOLE := FALSE;
  TripotageMenu(NumModule);
end;

procedure AssignZoom;
begin
  ProcGetVH   := GetVH;
  ProcGetDate := GetDate;
{$IFDEF EAGLCLIENT}
{$ELSE}
  // MBAMF
  //ProcZoomEdt := ZoomEdtEtat;
  //ProcCalcEdt := CalcOLEEtat;
  if not Assigned(ProcZoomGene)    then ProcZoomGene    := FicheGene;
  if not Assigned(ProcZoomTiers)   then ProcZoomTiers   := FicheTiers;
  if not Assigned(ProcZoomSection) then ProcZoomSection := FicheSection;
  if not Assigned(ProcZoomJal)     then ProcZoomJal     := FicheJournal;
  if not Assigned(ProcZoomCorresp) then ProcZoomCorresp := ZoomCorresp;
  if not Assigned(ProcZoomRub)     then ProcZoomRub     := FicheRubrique;
  if not Assigned(ProcZoomNatCpt)  then ProcZoomNatCpt  := FicheNatCpte;
{$ENDIF}
end;

procedure AfterUpdateBarreOutils ( Sender : TObject );
begin
  TToolbarButton97 ( Sender ).Visible :=      V_PGI.MenuCourant = 64;
end;

function LaunchAuto : boolean;
begin
  if FMenuG.GetLastNumClick<>0 then
  begin
    if V_PGI.MenuCourant=64 then FMenuG.LanceDispatch(64710);
    if V_PGI.MenuCourant=69 then FMenuG.LanceDispatch(64711);
  end;
  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure InitSerialisation;
var sDom : String ;
begin
  sDom           := '00280011';
  VH^.SerProdRIC := '00579080';
  FMenuG.SetSeria(sDom, [VH^.SerProdRIC], ['Révision intégrée']);
end;

////////////////////////////////////////////////////////////////////////////////

procedure InitApplication;
begin
  V_PGI.DispatchTT := DispatchTT;
  AssignZoom;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  FMenuG.OnAfterProtec:=AfterProtec ;
  FMenuG.OnChangeModule := AfterChangeModule;
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;

{$IFDEF EAGLCLIENT}
   V_PGI.DispatchTT:=DispatchTT ;
   AssignZoom ;
{$ELSE}
   FMenuG.OnMajPendant:=nil;
   FMenuG.OnAfterUpdateBarreOutils := AfterUpdateBarreOutils; // MBAMF
{$ENDIF}
   FMenuG.OnLaunchAuto := LaunchAuto ;
   FMenuG.SetModules([69, 64, 66], []);
   //FMenuG.SetPreferences(['Saisies'],False) ;

{$IFDEF EAGLCLIENT}

{$ELSE}
  //if EstSpecif('51502') then FMenuG.SetBouton(10, 'Récupération d''un dossier type depuis Sisco II', 'Restauration du dossier type à partir standards SISCOII', 64030, 1, FMenuDisp.LargeImages); //LargeImages) ;
  FMenuG.bUser.Visible := FALSE;
{$ENDIF}

  V_PGI.VersionDemo := False;
end;

Procedure InitLaVariablePGI;
Begin
  RenseigneLaSerie ( exeCCSTD );

  V_PGI.OutLook := TRUE;
  V_PGI.OfficeMsg := TRUE;
  V_PGI.ToolsBarRight := TRUE;
  ChargeXuelib;
  V_PGI.VersionDemo := TRUE;
  V_PGI.SAV := False;
  V_PGI.VersionReseau := True;

  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.NiveauAccesConf := 0;

  /////////////

  {Généralités}
  V_PGI.VersionDemo := True;
  V_PGI.SAV := False;
  V_PGI.VersionReseau := True;
  V_PGI.PGIContexte := [ctxCompta, ctxPCL, ctxStandard];
  V_PGI.CegidAPalatys := FALSE;
  V_PGI.CegidBureau := TRUE;
  V_PGI.StandardSurDP := True;
  V_PGI.MajPredefini := False;
  V_PGI.MultiUserLogin := False;
  V_PGI.BlockMAJStruct := True;
  V_PGI.EuroCertifiee := True;

  ChargeXuelib;

  {Série}
  V_PGI.OfficeMsg := True;
  V_PGI.OutLook := TRUE;
  V_PGI.NoModuleButtons := False;
  V_PGI.NbColModuleButtons := 1;

  {Divers}
  V_PGI.MenuCourant := 0;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  V_PGI.NiveauAccesConf := 1;
  V_PGI.VersionDEV := TRUE;
  V_PGI.ImpMatrix := True;
  V_PGI.DispatchTT := DispatchTT;
  V_PGI.ParamSocLast := False;
  V_PGI.RAZForme := TRUE;
End;


Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;

end.

