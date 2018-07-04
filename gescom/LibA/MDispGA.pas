unit MDispGA;
interface

Uses HEnt1,Forms,sysutils,HMsgBox, Classes, Controls, HPanel, UIUtil,ImgList
     , AGLInit , ParamSoc, Hctrls, DicoAF, MDispGAUtil ,choix ,Windows,Messages
{$IFNDEF SANSCOMPTA}
      ,devise
{$ENDIF}
{$IFDEF EAGLCLIENT}
      ,Maineagl,MenuOLX,M3FP,UtileAGL, eTablette

{$ELSE}                                          
      ,MenuOLG, UTOFAFSUIVIEACT,FE_Main,EdtEtat,EdtQr,tablette,pays, region,codepost,UtilSocAF
      ,PGIAppli, PieceEpure_Tof, db, dbtables

  {$IFNDEF V530}
      ,EdtREtat,UEdtComp
  {$ENDIF}
{$ENDIF}


{$IFDEF GIGI}
      ,UtomParBoni,UtofApprecon_Mul ,UtofAppreconAnu_Mul
      ,UtofApprecAff_Mul, UtofAfModifBoni_mul
 {$IFNDEF EAGLCLIENT}
      ,galoutil
 {$ENDIF}
{$ENDIF}

{$IFNDEF EAGLCLIENT}
 {$IFDEF DP}
     ,Utoftraitement_Sur_Annuaire1 ,Utoftraitement_Sur_Annuaire2
 {$ENDIF}
     , UtofGcMulSupprArt  ,UtomBlocageAffaire ,UtofAfEactivite_Mul
        ,UtofAfAffair_ModifLot,UtofAfVentilAna
      ,UtofAfModifLot ,UtofGcTiers_ModifLot
      ,UtofAffaireExport_Mul ,UtofDocExterne
      , UtofAFGENCUTOFFAFFMUL ,UtofFact_Eclat,
     //  , UTOFAFTIERS_MUL_JP
      UtofAfModif_unite, UtofAfModif_PremSem
     ,GalTomUserConf // mcd 27/06/03 pour fiche MDUSERCONF
     ,GalTofUserConf  // pour nouvelle saisie groupe travail
     ,UtofAchatCube ,Edit_Tarif_Tof
{$ENDIF}
{$IFDEF GRC}
      ,RTMenuGRC
{$ENDIF}

     ,UtofAfFormule, UtomFerie
     ,GrpFavoris ,UtilDispGC,HSysMenu ,PGIEnv,HCalc,AffaireUtil,LicUtil,math
     ,UtofEtatsAffaire ,UtofEtatsTache, UtofAftableauBord ,UtofEtatsActivite ,UtofAffaire_Mul
     ,UtofAfJournalAct ,UtofTarif ,UtofAfTarifArt_Mul ,UtofAfPiece, UtofAfPieceProAnu_mul
     ,UtofAfSaisDate ,UtofAfActiviteMul,UtofAfActiviteVisaMul,UtilMenuAff
     ,UtofRessource_Mul, UtofReval_Ress_Mul , UtofAfModifCutOff_Mul
     ,UtofAFARTICLEREVALMUL,  UtofAFFACECLAT_MUL ,UtofAFRessourceOccup,utofAfActiviteCube
     ,UtofAFPlanningViewer, UtofAFPlanningRes,UTofAFPlanning_Mul, UTOFAFTACHE_MUL
     ,UtofAFPLANNINGCOMPAR, UtofAFPlanningRemp, UtofAFPlanningAct, UTofAFPlanGenerer_Mul
     ,Utoftraitement_Sur_Affaire,UtofAfAugmen ,UtofAfTiers_Mul
     ,UtofAfDuplic ,UtofAfArticle_Mul ,UTomCompetence
     ,UtofEtats_Pieces,UtofAfBaseLigne_Mul ,UtofAfprepFact
     ,UtofPrintCalRegle,UtofAfProposEditMul,UtofAfProfilGener
     ,UtofAfVisaPiecPro_Mul, UtilGa ,UtofMBOVentilAna
     ,UtofAfVentesCube ,UtofAfValidePropos
     ,UtofHoraireStd,UTofAFPlanningPDC_Mul,UtofAFPlanningRes_Mul
     ,UTofAftachelig_mul ,UtofAfRegen_FactEclat,UTofAFLIGPLANNING_Mul
     ,UTofAfRessou_Synchro
     ,GcFourn_mul_tof, UtofGcEtabliss_mul, //mcd 06/12/02 obli pour OK fct commune
     UtofGraphPalm, UtofGraphStat, 
     initGA, ENTPGI, HistorisePiece_Tof, tradArticle, UtofGcTraiteAdresse, UtofAfPlanningReval,
     UtofAfTacheModele,UtomAFModeleTache, UtofAfEcheanceCube,
     UtomCommentaire,   //mcd 06/05/03 pour OK commentaire multiligne
     UtOfSuiviPiece, // mcd 26/06/03 pour zoom sur piece
     Graphics, UTob, UtofCPJalEcr, Utreelinks, AssistMajCompteurSouche,
     uTofAfPublicationMul,uTofAfIndiceMul,uTofAFValIndiceMul,uTofAFRevisionMul,
     uTofAfRevFormule_Mul, uTofAfRevModifCoef,utofAformuleVar_Mul,utofAformuleVar,
     utofAfRevRegulFacture, UtilRevision,utofAfRevEtatFacture,UtomAffairePiece,
     utofafRevImportIndice, utofAfRevExportIndice, utofAfRevPrepaImport,
     {$ifdef GIGI}
     YNewDocument,  UGedFiles, Ydocuments_Tom, // pour GED
       {$ifndef EAGLCLIENT}
       DpTofDOssierSel, // pour GED pas encore eagl surtout inclu Galoutil ..
       {$ENDIF}
     {$ENDIF}
     GcArticle_Tof, //mcd 26/09/03 sinon manque ds projet
     UTOFAFREVCALCCOEF,
     UTofAfActivitePai_Mul,UtofAfPaieAnnul,UTofAfPaieGener,UTofAFPaieGenerChoix,UtofPgImportFic
     ,utomactivitepaie,UtofAfRessou_modiflot,Confidentialite_TOF;
                                                      
Procedure InitApplication ;
procedure AfterChangeModule ( NumModule : integer ) ;
procedure AppelTitreLibre ( Name:String;PRien : THPanel  ) ;
function MyProcCalcMul(Nom, params, WhereSQL : string; TT : TDataSet; Total : Boolean) : string;
{$IFDEF CCS3}
Procedure TripoteMenuGAS3 ( NumModule : integer ) ;
{$ENDIF}

type
  TFMenuDisp = class(TDatamodule)
{$ifdef GIGI}
    private
     SeriaGedOk : Boolean;
{$endif}
    end ;

      TFEvtPopUpS5 = Class(TObject)
    public
        procedure PopupMenuS5Popup(Sender: TObject);
    END ;
{$ifdef GIGI}
procedure MyAfterImport (Sender: TObject; FileID: Integer; var Cancel: Boolean) ;
{$endif}

Var
   FMenuDisp         :TFMenuDisp;
   FEvtPopUpS5       :TFEvtPopUpS5;
   ModulesInterdits  :string;

   // gestion du module onyx
   FNumMenus    : Array of integer;
   FNumIcones   : Array of integer;
   FNextModule  : Integer;

   //JP: chemin des standards (pathstd + sousrep de l'appli), des dat (pathdat ...)
   // auparavant dans affaireole.pas
//   AppliPathStd, AppliPathDat   :string;



implementation

uses Ent1,Facture,dimension,UTomPiece, UtilPGI,Entgc,utilgc,
     TarifArticle,UtilArticle, TarifCliArt, TarifTiers, TarifCatArt, FichCOmm

{$IFNDEF SANSCOMPTA}
            ,SuprAuxi,Souche,TVA,InitSoc,OuvrFerm
  {$IFDEF GIGI}
            ,AssistPL
  {$ENDIF}
{$ENDIF}
            ,TiersUtil
            ,ZoomEdtAffaire, ConfidentAffaire,TraducAffaire
            ,AFActivite, ActiviteUtil,CalcOleGescom 
{$IFDEF GRC}
            ,ZoomEdtProspect,UtilRT //,InfoVersion, AssistImportProsp ,RecupSuspect
{$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
            ,CreerSoc ,CopySoc,Reseau,Mullog,Arrondi,EdtDoc ,UserGrp_tom
           ,UtomUtilisat  ,ListeDeSaisie,RecupTempoAssist,AssistCodeAffaire, AffaireOle
  {$IFDEF GIGI}                                                  
            ,AssistGI  // pas en EAGL ...
  {$ENDIF}
{$ENDIF}
            ,utofafactivite_suivi; // $$$JP 23/06/2003
{$R *.DFM}

Procedure LanceEtatLibreGIGA ;
Var CodeD : String ;
BEGIN
CodeD:=Choisir('Choix d''un état libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="ALI" AND MO_MENU="X"','') ;
if CodeD<>'' then LanceEtat('E','ALI',CodeD,True,False,False,Nil,'','',False) ;

END ;

// ****************************************************************************
// ******************* Dispatch gestion interne + Gestion d'affaire  **********
// ****************************************************************************

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
   dDateAP:TDateTime;
   ListMenu : string;
{$IFNDEF EAGLCLIENT}
   stParamsocAffiche, stParamsocVire : String;
      StMessage:string;
{$endif}
   i : Integer;
      
BEGIN
if CtxScot in V_PGI.PGIContexte then
 begin  //on change le nom du cpation à l'écran le NomHalley ne doit en aucun cas changer car utilisé dans registry ..
 //if Num > 100 then FMenuG.Caption := 'Gestion Interne PGI Expert';
 FMenuG.Caption := 'Gestion Interne PGI Expert';
 end;
if Not(TraiteObjActivite(Num,PRien)) then Exit; // PL après V535 : gestion particulière de la saisie d'activité

   CASE Num of
   10 : BEGIN   //  après connexion
//PgiInfo ('Dispatch 10  ',TitreHalley);
        UpdateCombosGC ; GCModifLibelles;  // remplace libellé libre dans table + combo
        VH_GC.GcMarcheVentilAna:='AFF';
{$ifndef EAGL}
{$ifdef GIGI}
        V_PGI.QRPdf :=True; // GED à forcer si on veut que l'archivage GED fctionne dans les QRS1
        If FmenuDisp.seriaGedOk then InitializeGedFiles(V_PGI_Env.DBComName, MyAfterImport); // GED  A supprimer si on ne veut pas la GED
{$endif}
{$endif}
        InitDicoAffaire;
        TripoteStatusAF; //  affiche les info complémentaire sur barre status /paramsoc préférence compta
        If (V_PGI.Superviseur)  then
          begin
          if Not(ExisteSql ( 'SELECT GDA_AXE FROM DECOUPEANA')) then
             begin
             if ExisteSql ('SELECT AST_AXE FROM STRUCRANAAFFAIRE') then
                 PgiInfo ('Vous devez mettre en place la nouvelle analytique. Merci d''appeler l''assistance téléphonique','');
             end;
          end;

        // bob planning
        // execute seulement si la table planning est vide
{$ifndef GIGI}
{$ifndef EAGLCLIENT}
       if not ExisteSQL('SELECT APP_CODEPARAM FROM AFPLANNINGPARAM') then
          if Import_Bob('GAS5') = 1 then
          begin
            StMessage:='Le système doit initialiser le planning.';
            StMessage:=StMessage+#13#10+'Pour prendre en compte ce traitement';
            StMessage:=StMessage+#13#10+'Le logiciel va se fermer automatiquement';
            StMessage:=StMessage+#13#10+'Vous devez relancer l''application';
            PgiInfo(StMessage+':'+V_PGI_Env.DbsocName,'Initialisation planning');
  {$ifndef AGL560E}
            FMenuG.ForceClose := True;
            FMenuG.Close;
  {$else}
            FMenuG.Quitter;  //mcd 17/06/03
  {$endif}
            PostMessage (Fmenug.handle,WM_CLOSE,0,0);
            Retourforce:=True;
            SortieHalley:=true;
          end;
{$endif}
{$endif}

        if ctxScot in V_PGI.PGIContexte then
             begin
             AfterChangeModule(141);
               // mcd 16/09/2002 ajout traitement des bob de l'appli en auto
{$ifdef GIGI}
  {$ifndef EAGLCLIENT}
             if Pcl_Import_Bob('GIS5') = 1 then
                begin
                StMessage:='Le système à déterminé une MAJ de menu';
                StMessage:=StMessage+#13#10+'Pour être prise en compte';
                StMessage:=StMessage+#13#10+'Le logiciel va se fermer automatiquement';
                StMessage:=StMessage+#13#10+'Vous devez relancer l''application';
                PgiInfo(StMessage+':'+V_PGI_Env.DbsocName,'MAJMenu');
  {$ifndef AGL560E}
                FMenuG.ForceClose := True;
                FMenuG.Close;
  {$else}
                FMenuG.Quitter;  //mcd 17/06/03
  {$endif}
                PostMessage (Fmenug.handle,WM_CLOSE,0,0);
                Retourforce:=True;
                SortieHalley:=true;
                end;
{$endif}// fin eagl
{$endif}
             end
        else
            AfterChangeModule(71);
        CHARGESOUSPLANAXE;

        if EstBloque('AffToutSeul',TRUE) then
           begin
             PgiInfo ('Vous ne pouvez pas utiliser l''application, une opération monoposte est en cours.',TitreHalley);
  {$ifndef AGL560E}
            FMenuG.ForceClose := True;
            FMenuG.Close;
  {$else}
            FMenuG.Quitter;  //mcd 17/06/03
  {$endif}
             Exit;
           end;
        //gestion de la confidentialité.
        if (not V_PGI.Superviseur) and (not V_PGI.Controleur) then
         begin
          if(VH_GC.RessourceUser='') and (GetParamsoc('SO_AFTYPECONF')<>'A00') then
             begin
               PgiInfo ('Vous n''êtes pas autorisé(e) à utiliser l''application.',TitreHalley);
  {$ifndef AGL560E}
            FMenuG.ForceClose := True;
            FMenuG.Close;
  {$else}
            FMenuG.Quitter;  //mcd 17/06/03
  {$endif}
               Exit;
             end;
         end;
        Bloqueur ('Affaire',True); // permet d'indiquer que quelqu'un est dans l'applicatif. Utiliser pour blocage mono affaire
{$IFNDEF EAGLCLIENT}
        AjouteApplisPopOffice;     // si appel depuis lanceur, affiche les appli utile pour ce dossier
{$ELSE}
   // AFAIREEAGL
{$ENDIF}

         // JP: chemin des std et dat (pathstd et pathdat, suivi du nom rep' de l'application: GI, GA, ...)
//{$IFNDEF EAGLCLIENT}   // gmeagl à vérifier
//$$$JP         if V_PGI_ENV <> nil then
//        begin
//{$IFDEF GIGI}
//              AppliPathStd := V_PGI_ENV.PathStd + '\GIS5';
//              AppliPathDat := V_PGI_ENV.PathDat + '\GIS5';
//{$ELSE}
//              AppliPathStd := V_PGI_ENV.PathStd + '\GAS5';
//             AppliPathDat := V_PGI_ENV.PathDat + '\GAS5';
//{$ENDIF}
//         end
//         else
//         begin
//              // V_PGI_ENV inexistant, pas normal!
//              AppliPathStd := 'C:\PGI00\STD';
 //             AppliPathDat := 'C:\PGI01\STD';
//         end;
//{$ENDIF}
       END;

   11 :begin
//   PgiInfo ('Dispatch 11  ',TitreHalley);
  end; // après déconnection
   12 :  //  avant connexion et séria
      begin
           // si scot et depuis lanceur, si appli pas active
           // il faut passer en mode silencieux pour ne pas afficher les messages
           // des codes compta non renseignés...
{$IFNDEF EAGLCLIENT}  //AFAIREEAGL
  {$IFDEF GIGI}
        if ctxscot in V_PGI.PGIContexte then VH^.ModeSilencieux := not GetFlagAppli ('CGIS5.EXE');  // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
  {$ENDIF}
{$ENDIF}
      end ;
   13 :  begin // avant déconnection
//PgiInfo ('Dispatch 13  ',TitreHalley);
      bloqueur ('Affaire',False);
			ChargeMenuPop(-1,FMenuG.DispatchX) ;
      // Libère les tablettes à traduire
      DeleteDicoAffaire;
      // Libère la variable interne Affaire
      VH_GC.AFTobInterne.Free;
      VH_GC.AFTobInterne := nil;

      // si lien avec le lanceur, il faut sortir appli, sans passer par sélection dossier
      if (V_PGI.CegidApalatys=False) And (ctxscot in V_PGI.PGIContexte) and (not FMenuG.ForceClose) then
         begin
  {$ifndef AGL560E}
            FMenuG.ForceClose := True;
            FMenuG.Close;
  {$else}
            FMenuG.Quitter;  //mcd 17/06/03
  {$endif}
         Exit;
         end;
  {$ifdef GIGI}
      FinalizeGedFiles; //GED
  {$endif}
      end;
   15 : ; //  ?????? sinon err avec AGL 5.2.3.C

   16 : if (not V_PGI.VersionDemo) and
           GetParamSoc('SO_AFREVISIONPRIX') and not (ctxscot in V_PGI.PGICOntexte) then
        begin
          SetLength(FNumMenus,FNextModule);   // on crée l'array du nbre voulu
          SetLength(FNumIcones,FNextModule);
 
          // placement du module revision après les ventes
          for i := FNextModule - 1 downto 4 do
          begin
            FNumMenus[i] := FNumMenus[i-1];
            FNumIcones[i] := FNumIcones[i-1];
          end;
                      
          FNumMenus[3] := 262;
          FNumIcones[3] := 3;
   
          if (FNextModule > V_PGI.NbRowModuleButtons * V_PGI.NbColModuleButtons) then
          begin
            V_PGI.NbColModuleButtons := 2;
            V_PGI.NbROwModuleButtons := (FNextModule div 2) + (FNextModule mod 2);
          end;
      
          FMenuG.SetModules(FNumMenus, FNumIcones);
          FNumIcones := nil;
          FNumMenus := nil;
       end;
   
   100 : begin      // appel depuis le lanceur
{$IFDEF GIGI}
{$IFNDEF EAGLCLIENT}
  {$IFNDEF SANSCOMPTA}
//PgiInfo ('Dispatch 100  ',TitreHalley);
 If (fmenug.forceclose) then Exit;      // Assistant Compta
       if Not GetFlagAppli ('CCS5.EXE') then   // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
            if Not LanceAssistantComptaPCL (True) then
                begin
                FMenuG.ForceClose := True;
                FMenuG.Close;
                exit;
                end;
  {$ENDIF}
       // Assistant GI
       if not LanceAssistantGI_PCL (True) then
          begin
  {$ifndef AGL560E}
            FMenuG.ForceClose := True;
            FMenuG.Close;
  {$else}
            FMenuG.Quitter;  //mcd 17/06/03
  {$endif}
            exit;
          end;

              // Si on est dans la base 00, on force lien DP à VRAI, à Faux sinon
      if ctxScot in V_PGI.PGIContexte then
          begin
          If (V_PGI_Env.InBaseCommune) then SetParamsoc('SO_AFLIENDP',TRUE)
                                       else SetParamsoc('SO_AFLIENDP',FALSE);
          end;
            // si modif voir aussi UtilSocAf ,ConfigSaisiePreferences
           //mcd 28/11/00 voir si il ne faudra pas le supprimer, si on permet par la suite
           // de faire des liens DP de n'importe quel dossier ... else  SetParamsoc('SO_AFLIENDP',FALSE);
 {on peut aussi débrayer l'accès à la base DB000 (car fait en automatique
 par l'AGL sur la table ANNUAIRE (si on veut pe, un dossier en lien DP, mais dont le DP sera dans la base en cours
 et non pas dans la DB0000 !!!) :
  mettre : V_PGI.StandardSurDp := False
  dans le prog après connexion
  si tu es dans un dossier qui ne doit pas taper dans la base commune.  }

                // Mettre par la suite les appels directs de fonctions
                // si accès pour les temps, planning ....
         RetourForce := True;
{$ENDIF} // fin Non Eagl
{$ENDIF}
         end;


// *********************Module Gestion Interne : Fiches de bases **************
//////// Fiches ///////////
   71101 :  AFLanceFiche_Mul_tiers('T_NATUREAUXI=CLI','');   // Clients en mono fiche avec tierscompl
   71106 :  AFLanceFiche_Mul_tiers('T_NATUREAUXI=CLI','MULTI=VRAI');   // Clients en multi fiche sans tierscompl
   71102 : AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=PRO','STATUT=PRO;NOCHANGESTATUT');  // proposition de mission

   //$$$JP - mul sur clients avec champs dp (essai)
{$IFNDEF EAGLCLIENT}
  // 71105 : AFLanceFiche_Mul_tiers_DP ('T_NATUREAUXI=CLI', '');
{$ENDIF}
   //$$$JP FIN

   71103 : AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT');  // mission
   71107 : AFLanceFiche_MulAffaire ('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT;PASMULTI');  // mission (avec stat cli)
   // 71107 : AglLanceFiche ('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=AFF','','STATUT=AFF;NOCHANGESTATUT;PASMULTI');  // mission (avec stat cli)


   // C.B les avenants avaient été supprimés !
   71111 : AFLanceFiche_MulAvenantAff('AFF_STATUTAFFAIRE=PRO','STATUT=PRO;NOCHANGESTATUT');  // avenant sur proposition de mission
   71112 : AFLanceFiche_MulAvenantAff('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT');  // aveanant mission   (mcd 14/04/03 appel statut avec AFF au lieu PRO)
   //71111 : AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=PRO','STATUT=PRO;NOCHANGESTATUT');  // proposition de mission
   //71112 : AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT');  // mission


   71104 : AFLanceFiche_Mul_Ressource;   // Assistants
//////// Editions ////////
   71211 : AFLanceFiche_EtatsClient;    // Etats sur clients
   71212 : AFLanceFiche_FicheClient;
   71213 : AFLanceFiche_EditStatClient;
   71214 : AFLanceFiche_EditAdresseClient;

   71221 : AFLanceFiche_EtatsAffaire;   // Etats sur affaires
   71222 : AFLanceFiche_FicheAffaire;
   71223 : AFLanceFiche_EditStatAffaire;

   71231 : AFLanceFiche_ETatsRess; // Etats sur ressources
   71232 : AFLanceFiche_FicheREssource;
   71233 : AFLanceFiche_EditStatRess;
   71234 : AFLanceFiche_EditAdresseRess;

   71271 : AFLanceFiche_EtatsPrestation; // Etats sur prestations
   71272 : AFLanceFiche_EditStatPrest;

   71204 : AFLanceFiche_EtiquetteClient;
         
/////// Calendrier  71240
   71241 : AFLanceFiche_Edit_Calendrier;
   71242 : AFLanceFiche_Edit_CalRegle;

   71250 : LanceEtatLibreGIGA;
   71261 : AFLanceFiche_FicheTache; // fiche signalétique tâche


////// Documents clients ///
   71301 : AFLanceFiche_MulEditPropos('AFF_STATUTAFFAIRE=PRO','STATUT=PRO') ; // Edition des propositions
   71302 : AFLanceFiche_MulEditPropos('AFF_STATUTAFFAIRE=AFF','STATUT=AFF') ; // Edition des affaires

{$IFNDEF EAGLCLIENT}
//$$$jp: mul sur documents excel OLE, pour exécution
   72408 : AFLanceFiche_DocExterneMul ('', 'TYPE=EXC;NATURE=ESO;MODEEXEC');
{$endif}

////  Traitements /////////
   71401 : AFLanceFiche_ValidePropos('AFF_STATUTAFFAIRE=PRO','PRO');  // acceptation de missions
{$IFNDEF SANSCOMPTA}
   71404 : begin
           // fct monoutilisatuer. il faut enlever le blocage Affaire mis systématique
         bloqueur ('Affaire',False);
         OuvreFermeCpte(fbAux,FALSE,'CLI') ;  // fermeture des comptes tiers
         Bloqueur ('Affaire',True); // permet d'indiquer que quelqu'un est dans l'applicatif. Utiliser pour blocage mono affaire
         end;
   71405 : OuvreFermeCpte(fbAux,True,'CLI') ;   // ouverture compte tiers
   71406 : SuppressionCpteAuxi('CLI') ;     //supression CLi seulement
   // C.B 05/12/02
   71418 : EntreeTradArticle (taModif) ;  // traduction libellé articles

{$ENDIF}
   71411 : AFLanceFiche_ALign_AffaireSurClt ;                    // Alignement client sur affaire
   71412 : AFLanceFiche_EpurEchFAct('') ;                // epuration échéances facturées
   71413 : AFLanceFiche_EpurEchFAct('GENERATION');       // génération échéances sur tacite reconduction
{$IFNDEF EAGLCLIENT}
   71403 : AFLanceFiche_ModifLot_TiersAff('CLI') ;  // modification en série des tiers
   71407 : AFLanceFiche_Mul_Suppr_Art ;                 // épuration articles
   71410 : AFLanceFiche_Modiflot_Article ;             // modification en série
   71415 : AFLanceFiche_Modiflot_Affaire('') ;             // modif en série sur affaire
   71416 : AFLanceFiche_Modiflot_Affaire('TRA')  ;             // modif en série sur affaire avec traitement
{$ENDIF}
   71408 : AFLanceFiche_Mul_DuplicMiss('AFF_STATUTAFFAIRE=AFF') ;      // Duplication mission N en N
   71409 : AFLanceFiche_Mul_AugmentAff('AFF_STATUTAFFAIRE=AFF') ;      // Augmentation affaire
   71441 : AFLanceFiche_ModifLog_Ress ('');             // Modif en série des ressources
   71442 : AFLanceFiche_ModifLog_Ress ('TRA');          // Modif en série des ressources  avec traitement
   71445 : AFLanceFiche_Synchro_RessSalarie;            // Synchro ressources / salariés

//////// Structures ///////
   71511 : AFLanceFiche_Mul_Article('GA_TYPEARTICLE=PRE','PRE');    // Prestations
   71512 : AFLanceFiche_Mul_Article('GA_TYPEARTICLE=FRA','FRA');    // Frais
   71513 : AFLanceFiche_Mul_Article('GA_TYPEARTICLE=MAR','MAR');    // Fournitures
   71514 : AFLanceFiche_Mul_Article('GA_TYPEARTICLE=CTR','CTR');    // articles de contrat
   71515 : AFLanceFiche_Mul_Article('GA_TYPEARTICLE=POU','POU');    // articles de pourcentage
   71501 : AFLanceFiche_Mul_Apporteur ;                  // Apporteurs Attention num utilisé AfterChangeModule
   71502 : AFLanceFiche_HoraireStd('TYPE:STD');
   71503 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=CLI','','');                      // Annuaire des contacts
   71504 : AFLanceFiche_IntervAff;                   // intervenants sur affaires
   71505 : AGLLanceFiche('GC','GCPRIXREVIENT','','','') ;                       // Prix de revient articles
   // C.B 05/12/02
   71506 : AGLLanceFiche ('GC','GCCOMMERCIAL_MUL','','','') ; // Commerciaux

////// Tarifs
   71601 : EntreeTarifArticle(taModif) ;
   71602 : EntreeTarifTiers (taModif) ;
   71603 : EntreeTarifCliArt (taModif) ;
   71604 : EntreeTarifCatArt (taModif) ;
   71605 : (* pas de tof*) AGLLanceFiche('AFF','AFTARIFCON_MUL','GF_REGIMEPRIX=GLO','','');       //consultation tarif
   71606 : AFLanceFiche_Mul_Maj_Tarif('GF_REGIMEPRIX=GLO','');      // maj tarif
   // C.B 05/12/02
   71608 : AGLLanceFiche('GC','GCASSISTANTTARIF','','','VEN') ; // Assistant création tarif article
   71615 : AFLanceFiche_Mul_Maj_Tarif('GF_REGIMEPRIX=GLO','DUP');  // duplication tarif
   71616 : AFLanceFiche_Mul_Maj_Tarif('GF_REGIMEPRIX=GLO','SUP;VEN');  // suppression tarif
   71617 : AGLLanceFiche('GC','GCEDITTARIFTIE','','','CLI') ; // Edition des tarifs par client
   71618 : AGLLanceFiche('GC','GCEDITTARIFART','','','CLI') ; // Edition des tarifs par article
{$IFNDEF EAGLCLIENT}
   71607 : AFLanceFiche_Mul_TarifArt('','') ;  // Maj BAses Tarifaires
{$endif}
   // Tarifs TTC
   71611 : EntreeTarifArticle(taModif,TRUE) ;
   71612 : EntreeTarifTiers (taModif,TRUE) ;
   71613 : EntreeTarifCliArt (taModif,TRUE) ;

// **************************Module Activité et reporting ******************
//////// Saisie d'activité ///////
   // Par Assistant
   72102 : AFCreerActivite ( tsaRess, tacTemps, 'REA' ) ; // Saisie d'activités Temps
   72103 : AFCreerActivite ( tsaRess, tacFrais, 'REA' ) ; // Saisie d'activités Frais
   72104 : AFCreerActivite ( tsaRess, tacFourn, 'REA' ) ; // Saisie d'activités Fournitures
   72105 : AFCreerActivite ( tsaRess, tacGlobal, 'REA' ) ; // Saisie d'activités Globale
   // Par Client/Mission
   72111 :  AFCreerActivite ( tsaClient, tacTemps, 'REA' ); // Saisie d'activités Temps
   72112 : AFCreerActivite ( tsaClient, tacFrais, 'REA' ); // Saisie d'activités Frais
   72113 : AFCreerActivite ( tsaClient, tacFourn, 'REA' ); // Saisie d'activités Fournitures
   72114 : AFCreerActivite ( tsaClient, tacGlobal, 'REA' ); // Saisie d'activités Globale

/////// Editions sur activités ////
    72203 : AFLanceFiche_Synthese_F_NF;     // synthèse fact/nf
    72211 : AFLanceFiche_JournalActivite;
    72212 : AFLanceFiche_JournalClient;
    72221 : AFLanceFiche_Synthese_Ress_Aff;    // Synthèse assistant detaillée par affaire
    72222 : AFLanceFiche_Synthese_Ress_Art;  // Synthèse assistant detaillée par article
    72231 : AFLanceFiche_Synthese_Clt_Ress;   // Synthèse Client affaire détaillée par ressource
    72232 : AFLanceFiche_Synthese_Clt_Art;   // Synthèse Client affaire détaillée par article

////// Historique ///////////

////// Tableaux de bord //////
    72401 : AFLanceFiche_TbAffaire('');    // Consult tableau de bord
    72405 : AFLanceFiche_TbMultiAff('MULTI:X');     // Consult TB Multi affaires
    72402 : AFLanceFiche_GrandLivre('');  //grand livre ;
    72403 : AFLanceFiche_Balance('') ;     //balance;
    72406 : AFLanceFiche_Balance('ETAT:AAN') ;     //etat activité avec AN;
    72404 : AFLanceFiche_TableauBord; //AlimTableauBordAffaire;
    72407 : AFLanceFiche_VentesCUbe('','CONSULTATION;AFFAIRE;TABLE:GL;') ;

////// Traitements //////////
    72501 : AFLanceFiche_Mul_VIsa_Activ;    // Visas sur l'activité
    72502 : AFLanceFiche_RevalActiv;        // Revalorisation de l'activité par ressources
    72503 : AFLanceFiche_Modif_Activite('AFA_REPRISEACTIV:TOU;');    // Modif ligne activité
    72504 : BEGIN InitParpieceAffaire; CreerPiece('FRE'); END;                                                        //Reprise facturation externe
    72505 : AFLanceFiche_Mul_Piece('GP_NATUREPIECEG=FRE','TABLE:GP;NOCHANGE_NATUREPIECE;NATURE:FRE;ACTION=MODIFICATION;');  // COnsultation facturation externe
    72506 : begin    // Arrêté de période
             V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
// PL le 04/11/02 : la date d'arrêté de période proposée par défaut doit être la veille
// de la date de début d'activité
//             AFLanceFiche_SaisieDate('ZORI:ARRETE;TITRE:Arrêté de période;LIBELLE:Date de l''arrêté de période;TEST:X;ZZDATE:'
//              +format('%s',[DateToStr(GetParamsoc('SO_AFDATEDEBUTACT'))])) ;
             dDateAP := GetParamsoc('SO_AFDATEDEBUTACT');
             if (dDateAP > iDate1900) then dDateAP := dDateAP - 1;
             AFLanceFiche_SaisieDate('ZORI:ARRETE;TITRE:Arrêté de période;LIBELLE:Date de l''arrêté de période;TEST:X;ZZDATE:'
              +format('%s',[DateToStr(dDateAP)])) ;
///////////////////////////////////////              
             RetourForce := True;
             V_PGI.ZoomOLE := False;
            end;
    72508 :  ;  //modification facture éclatéee par assistant
    72509 :  AFLanceFiche_RevalActivArticle;  // Revalorisation de l'activité par article
    72510 :  AFLanceFiche_ActivGenerPaie; //Génération de l'activité vers la paie
    72511 :  AFLanceFiche_AnnulGenerPaie; //Annulation des lignes d'activité générées vers la paie
    
    ////// Consultation //////////
    72610 : AFLanceFiche_Consult_Activite;    // Modif ligne activité

    // $$$JP 23/06/03
    72615 : AFLanceFiche_Suivi_Activite ('TYPEARTICLE=PRE');      // Suivi activité par planning
    // $$$JP fin

    72620 : AFLanceFiche_ActiviteCube;        // Cube de l'activité
    ////// Cut off //////////
    72702 : AFLanceFiche_AfGenCutOffGlobale;  // Génération CutOff totale

{$IFNDEF EAGLCLIENT}
    72705 : AFLanceFiche_AfGenCutOffAFF_Mul ('')  ;  // Génération CutOff par affaire
{$endif}
    72704 : AFLanceFiche_ModifCutOffMul('TYPE:CVE;SAISIE:GLOBAL');    // Modification CutOff  cumulée
    72703 : AFLanceFiche_ModifCutOffMul('TYPE:CVE;SAISIE:DETAIL');    // Modification CutOff  détaillé
    72701 : AFLanceFiche_EtatAppreciation('ETAT:ACU');  // etat cut off
    72706 : AFLanceFiche_EtatAppreciation('ETAT:AFR');  // etat cut off par mois
    72707 : AFLanceFiche_AfGenCutCompta; //Génération Comptable des cut-off

// **************************  Module Honoraires  *********************
    ////// Factures et avoirs
    73101 : CreerPiece('FAC') ;  // Saisie de factures
    73103 : CreerPiece('FPR') ;  // Saisie de factures provisoires
    73104 : AFLanceFiche_Mul_Piece('GP_NATUREPIECEG=FPR','TABLE:GP;NOCHANGE_NATUREPIECE;NATURE:FPR;ACTION=MODIFICATION;');   //Modif Factures pro
    73102 : CreerPiece('AVC') ;  // Saisie d'Avoirs
    73105 : CreerPiece('APR') ;  // Saisie d'Avoirs Provisoires
    73106 : AFLanceFiche_Mul_Piece('GP_NATUREPIECEG=APR','TABLE:GP;NOCHANGE_NATUREPIECE;NATURE:APR;ACTION=MODIFICATION;');   //Modif avoirs pro
    73111 : AFLanceFiche_DuplicPiece('GP_NATUREPIECEG=FPR','NATURE:FPR;DUPPLIC:FPR;');  // duplication Factures  provisoires
    73112 : AFLanceFiche_DuplicPiece('GP_NATUREPIECEG=FAC;','NATURE:FAC;DUPPLIC:FAC;');  // duplication Factures
    73113 : AFLanceFiche_DuplicPiece('GP_NATUREPIECEG=AVC','NATURE:AVC;DUPPLIC:AVC;');  // duplication avoirs définitifs
    73114 : AFLanceFiche_DuplicPiece('GP_NATUREPIECEG=APR','NATURE:APR;DUPPLIC:APR;');  // duplication avoir provisoires
    // C.B 05/12/02
    73120 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','VEN') ; //Historisation des pièces

{$IFNDEF EAGLCLIENT}
    73121 : GCLanceFiche_PieceEpure('GC','GCPIECEEPURE_MUL','','','VEN') ; //Epuration des pièces
{$ENDIF}


    ////// Consultation
    73201 : AFLanceFiche_Mul_Piece('','TABLE:GP;ACTION=CONSULTATION;');   // COnsultation par piece
    73202 : AFLanceFiche_Mul_LignePiece('TABLE:GL;');   // consultation par ligne

    ////// Editions
    73301 : AFLanceFiche_ListeFact;
          //mcd 27/11/01 fiche qui n'existe pas ni ligne au menu !!!    73303 : AGLLanceFiche('GC','GCEDTPIECE','','','');
    73302 : AGLLanceFiche('GC','GCEDITTRAITEDIF','','','') ;      // Edition traites différées
    73303 : AFLanceFiche_EDitPrepFact ;     // Edition liste préparatoire factures
    73304 : AFLanceFiche_Portefeuille('VENTELIG');  // Portefeuille pièces
    73305 : AFLanceFiche_Portefeuille('');  // Portefeuille pièces
    73306 : AFLanceFiche_EtatCommission; // commissionnement commerciaux (rétrocession apporteur)

    //    73307 : AFLanceFiche_Mul_Edition_Piece('TABLE:GP;STATUT=FAC');                                                         // Attention num utilisé AfterChangeModule
    73307 : AGLLanceFiche('GC','GCEDITDOCDIFF_MUL','','','VEN') ;     // Editions différées
     

    73308 : AFLanceFiche_EtatAppreciation('ETAT:AFP');   // etat preparatoire facturation                                                       // Attention num utilisé AfterChangeModule
    73309 : CPLanceFiche_CPJALECR('NATUREJAL=ACH,VTE');   // journal de ventes comptable. appel fct Compta                                                       // Attention num utilisé AfterChangeModule
    ////// Génération
    73401 : AFLanceFiche_Mul_PrepFac;   // Prépa facture

    // fiche inexistante au 22/11/01 .... 73402 : AglLanceFiche ('AFF','AFVALIDFACT_MUL','','','');  // Liste des échéances par date
    73403 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FPR;QUEPROVISOIRES'); //;NOCHANGE_NATUREPIECE validation documents provisoires
    73404 : AFLanceFiche_Mul_Annu_PiecePro('TABLE:GP;NATURE:FPR;QUEPROVISOIRES'); //;NOCHANGE_NATUREPIECE  // annulation de documents provisoires
    73405 : AGLLanceFiche ('GC' ,'GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','FAC'); // Tranfo cde => factures / import E-Commandes
    73406 : AFLanceFiche_Mul_VisaPiecePro('TABLE:GP;NATURE:FPR;QUEPROVISOIRES'); //visas factures provisoires
    73408 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FAC;NOCHANGE_NATUREPIECE'); //;Facture en avoir
           
    // revision de prix
    73409 : AFLanceFicheRegulFacture;
    73203 : AFLanceFicheModifCoef('CONSULTER');

    //traitement
    73501 : begin
            V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
            AFLanceFiche_Regen_FActEclat ;   // Re génération fact ecalt sur une préiode
             V_PGI.ZoomOLE := False;
           end;
    73407 :  AFLanceFiche_ModifEclatFactMul('TYPE:FAC'); // modif facture éclatées
{$IFDEF GIGI}
    ////// Appréciation
    73601 : AFLanceFiche_ModifBoni_Mul('');   // modification des boni/mali
    73605 : AFLanceFiche_ParamBoni;   // Paramètrage des boni/mali
    73607 : AFLanceFiche_TbMultiAff('MULTI:X;BUT:APP');   // Apprec depuis TB
    73610 : AFLanceFiche_Mul_ApprecAff('AFF_STATUTAFFAIRE=AFF','ZORI:COM;');  // Appréc complete
    73611 : AFLanceFiche_EtatContrProd('ETAT:ACR') ;  // etat contrôle de production
    73615 : AFLanceFiche_Mul_PrepFacApp;  // Generation facture apprec
    73618 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FPR;NOCHANGE_NATUREPIECE'); //validation factures provisoires
    73619 : AFLanceFiche_EtatAppreciation ('ETAT:APF');  // etat préparatire facturation
    73620 : AFLanceFiche_Mul_Consult_Apprec('');  // Consultation apprec
    73625 : AFLanceFiche_Mul_Annu_Apprec;  //Suppression apprec
{$ENDIF}
    ////// Statistiques
    73510 : AFLanceFiche_VentesCUbe('','CONSULTATION;PIECE;TABLE:GL;') ;
    73520 : AFLanceFiche_EtatPrevCa('');  // Prévision de CA
    73521 : AFLanceFiche_EcheanceCube('');  // Prévision de CA EN CUBE
    // Palmarés
    73531 : AGLLanceFiche('GC','GCGRAPHPALM','','','VEN') ;
    73532 : AGLLanceFiche('GC','GCPALMARES','','','VEN') ;
    // stat
    73541 : AGLLanceFiche('GC','GCGRAPHSTAT','','','VEN') ;
    73542 : AGLLanceFiche('GC','GCSTAT','','','VEN') ;

// *****************************************************************************
// *****************************  Module Relation clients***********************
// *****************************************************************************
{$IFDEF GRC}
    92100..92999 : RTMenuDispatch (Num,PRien);
    /////// Menu Pop GRC ////////////////
    27080 : AGLLanceFiche('GC','GCPIECE_MUL','GP_VENTEACHAT=VEN;GP_NATUREPIECEG=DE','','CONSULTATION') ;  // devis pour GRC
{$ENDIF}

// *****************************************************************************
// *****************************  Module achat *********************************
// *****************************************************************************
    ////// Saisie des pièces
   // Saisie pièces
   138101 : CreerPiece('CF') ;
   138102 : CreerPiece('BLF') ;
   138103 : CreerPiece('FF') ;
   138104 : CreerPiece('AF') ;   //avoir fournissuer
   // modifications
   138111 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=CF','','MODIFICATION')  ;  //commandes fournisseur
   138112 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=BLF','','MODIFICATION') ;  //bons de livraison fournisseur
   // Duplication unitaire pièces
   138121 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=CF','','CF') ;
   138122 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=BLF','','BLF') ;
   138123 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=FF','','FF') ;
   138124 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AF','','AF') ; // avoir fournissuer

   138130 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','ACH') ; //Historisation des pièces
{$IFNDEF EAGLCLIENT}
   138140 : GCLanceFiche_PieceEpure('GC','GCPIECEEPURE_MUL','','','ACH') ; //Epuration des pièces
{$ENDIF}

   // Consultation
   138201 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_VENTEACHAT=ACH','','CONSULTATION;VENTEACHAT=ACH') ;
   138202 : AGLLanceFiche('GC','GCLIGNEACH_MUL','','','CONSULTATION') ;
     // génération
   138301 : AGLLanceFiche('GC','GCPIECEVISA_MUL','GP_VENTEACHAT=ACH','','VENTEACHAT=ACH') ; // Visa des pièces
   138302 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=CF','','BLF') ;
   138303 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=BLF','','FF') ;
      // Génération manuelle pièces
   138331 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=CF','','ACHAT;CFR;BLF') ;  // Comm --> Livr
   138332 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=BLF','','ACHAT;LFR;FF') ; // Livr --> Fact
      // Génération automatique des pièces
   138321 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CF','','ACHAT;BLF') ;
   138322 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLF','','ACHAT;FF') ;

   // tarif fournisseurs
{$IFNDEF SANSPARAM}
   138401 : EntreeTarifFouArt (taModif) ;                       // saisie
   138402 : AFLanceFiche_Mul_Consult_TarifFou;  // consultation
   138403 : AFLanceFiche_Mul_Maj_TarifFou;  // mise à jour
   // C.B 05/12/02
   138404 : AGLLanceFiche('GC','GCPARFOU_MUL','','','') ;  // Import
{$ENDIF}

   // Editions
   138501 : AGLLanceFiche('GC','GCEDITDOCDIFF_MUL','','','ACH') ;     // Editions différées
   138502 : AGLLanceFiche('GC','GCPTFPIECE','','','ACHAT') ;  // Portefeuille pièces
   138503 : AFLanceFiche_Portefeuille('ACHATS;NOFILTRE');  // Portefeuille pièces
   138511 : AGLLanceFiche('GC','GCFOURN_ARTICLE','','','') ;
   138512 : AGLLanceFiche('GC','GCART_NONREF','','','')  ;
   138513 : AGLLanceFiche('GC','GCARTICLE_FOURN','','','') ;
   // Fichiers
   138601 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','') ;
   // C.B 05/12/02                             
   138602 : AGLLanceFiche('GC','GCMULCATALOGUE','','','') ;
   138603 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=FOU','','') ;  // contacts
   138604 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','FOU') ;  // modification en série des tiers
{$IFNDEF SANSCOMPTA}
   138605 : OuvreFermeCpte(fbAux,FALSE,'FOU') ;  // fermeture des comptes tiers
   138606 : OuvreFermeCpte(fbAux,True,'FOU') ;   // ouverture compte tiers
   138607 : SuppressionCpteAuxi('FOU') ;
{$ENDIF}
    // CutOff fournisseur
   138701 : ;// Génération totale
   138702 : ;// Génération par mission
   138703 : ;// Modification globale
   138704 : ;// Modification détaillée
   138705 : ;// Editions par mois
   138706 : ;// Edition

/////// Module ANALYSES
   // Vente cf module vente GI ...
   // Achats ////
   139211 : AGLLanceFiche('GC','GCGRAPHSTAT','','','ACH') ;     // Statistiques des ventes graphe
   139212 : AGLLanceFiche('GC','GCSTAT','','','ACH') ;           // Statistiques Etat
   139221 : AGLLanceFiche('GC','GCGRAPHPALM','','','ACH') ;     // Palmares des achats graphe
   139222 : AGLLanceFiche('GC','GCPALMARES','','','ACH') ;      // Palmares des ventes Etat
   // C.B 05/12/02
   139230 : AGLLanceFiche('GC','GCACHAT_CUBE','','','') ;
   // Affaire cf Module activité / reporting GI
   // Refacturation interne
   139401 : AFLanceFiche_RefacturationRess;
   139402 : AFLanceFiche_RefacturationAffaire;

   // C.B 05/12/02
   73550 : AGLLanceFiche('GC','GCANALYSE_VENTES','GPP_VENTEACHAT=VEN','','CONSULTATION') ;

// *****************************************************************************
// ******************************  Module Planning  ****************************
// *****************************************************************************

  // Tâches
  153102 : AFLanceFicheAFTache_Mul('', '');
  153103 : AFLanceFicheAFPlanningPDC_mul('', 'PLANCHARGEQTERAF');
  153101 : AFLanceFicheAFPlanningPDC_mul('', 'TACHES');
  153104 : AFLanceFicheAFPlanningPDC_mul('', 'PLANCHARGEMNTRAF');
  153105 : AFLanceFicheAFTacheModele_Mul; // Modèles de tâches

  // Suivi du planning
  153201 : AFLanceFicheAFPlanning_Mul('', 'PLANNING1' + ';NUMPLANNING:' + intToStr(Num));
  153203 : AFLanceFicheAFPlanning_Mul('', 'PLANNING3' + ';NUMPLANNING:' + intToStr(Num));
  153202 : AFLanceFicheAFPlanningRes_Mul('', 'PLANNING2' + ';NUMPLANNING:' + intToStr(Num));
  153204 : AFLanceFicheAFPlanningRes_Mul('', 'PLANNING4' + ';NUMPLANNING:' + intToStr(Num));
  153205 : AFLanceFicheAFPlanningPDC_mul('', 'PLANCHARGEQTE');
  153206 : AFLanceFicheAFPlanningPDC_mul('', 'PLANCHARGEMNT');

  // Traitements
  153301 : AFLanceFicheAFPlanGenerer_Mul;
  153302 : AFLanceFiche_GenPlanningAct;
  153303 : AFLanceFiche_PlanningRemp;
  153305 : AFLanceFicheAFPlanningRevalAffaire;
  153306 : AFLanceFicheAFPlanningRevalTache;
  153307 : AFLanceFicheAFPlanningRevalRess;
  153308 : AFLanceFicheAFPlanningRevalArticle;
  153309 : AFLanceFicheAFTacheModele_Gen; // Affectation de tâches à partir des modèles
  153310 : AFLanceFicheAFPlanningModiflot; // Modification par lots des lignes de planning

  // paramétrage
  // pas de sources associés
  153401 : AGLLanceFiche('AFF','AFPLANPARAM_MUL','','',''); // paramétrage du planning
  153402 : AGLLanceFiche('AFF','AFPLANETAT_MUL','','',''); // paramétrage des états

  // analyse
  // plan de charge
  153501 : AFLanceFicheAFPlanningViewer('', 'MONOFICHE');
  153502 : AFLanceFicheAFPlanningRes('', 'MONOFICHE');
  // planning jour
  153503 : AFLanceFiche_DetailTaches; // mul des lignes de taches;
  153504 : AFLanceFiche_DetailPlanning; // mul des lignes de planning
  153506 : AFLanceFiche_OccupRessSem;
  153507 : AFLanceFiche_OccupRessMois;
  153508 : AFLanceFiche_PlanningCompar;

// *****************************************************************************
// ******************************  Module Paramétrage  *************************
// *****************************************************************************
{$IFNDEF EAGLCLIENT}
    /////// Société
    74101 : BEGIN
            BrancheParamSocAffiche (stParamsocVire,stParamsocAffiche);
            ParamSociete (FALSE,stParamsocVire,stParamsocAffiche,'',ChargeSocieteHalley,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
{ en prépa pour gestion domaine
            ParamSociete (FALSE,stParamsocVire,stParamsocAffiche,'A',ChargeSocieteHalley,ChargePageSocGA,SauvePageSocGA,InterfaceSocGA,110000220);
}
           END;
    74102 : AGLLanceFiche('GC','GCETABLISS_MUL','','','') ;          // FicheEtablissement_AGL(taModif);
  {$IFNDEF SANSCOMPTA}
    74103 : FicheSouche('GES');                                  // Compteurs de pièces
  {$ENDIF}
    74111 : EntreeListeSaisie('') ;
    74106 : AglLanceFiche ('GC','GCPARPIECECOMPL','FAC','','NATURE=FAC') ; // Complément Etablissement facture
    74107 : AglLanceFiche ('GC','GCPARPIECECOMPL','AVC','','NATURE=AVC') ; // Complément Etablissement Avoir
  {$IFDEF GIGI}
    74108 : LanceAssistantGI_PCL (false);
//    74108 : LanceAssistantComptaPCL (True);
  {$ENDIF}
    74109 : BEGIN LanceAssistCodeAffaire; RetourForce := True;  END;
    74104 : BEGIN
            if (V_PGI.PassWord = CryptageSt(DayPass(V_PGI.DateEntree))) then
               AglLanceFiche ('GC','GCPARPIECE','','','') // LanceparPiece
            else AglLanceFiche ('GC','GCPARPIECEUSR','','','') ; // LanceparPiece
            END;
{$ELSE}
    74104 : AglLanceFiche ('GC','GCPARPIECEUSR','','','') ;
{$ENDIF}   // fin ifndef EAGLCLIENT
 
    ////// Paramètres généraux
  {$IFNDEF SANSCOMPTA}
    74151 : ParamTable('TTREGIMETVA',taCreat,AfNoAideTablette('TTREGIMETVA'),PRien) ; //RegimesTva ;
    74152 : ParamTvaTpf(true) ;
  {$ENDIF}
    74153 : FicheModePaie_AGL('');
    74154 : FicheRegle_AGL ('',False,taModif);
  {$IFNDEF SANSCOMPTA}
    74155 : FicheDevise('',tamodif,False);
  {$ENDIF}
    74156 : ParamTable('ttFormeJuridique',tacreat,AfNoAideTablette('TTFORMEJURIDIQUE'),PRien) ;
    74157 : ParamTable('ttCivilite',taCreat,AfNoAideTablette('TTCIVILITE'),PRien) ;
    74158 :  AglLanceFiche ('YY','YYPAYS','','','') ; //ouvrePays; supprimer le 26/11/01 idem gescom
    74159 : AglLanceFiche('GC','GCREGION_FSL','','','');
    {$IFNDEF EAGLCLIENT}
    74160 : OuvrePostal(PRien) ;
    {$ENDIF}
    74161 : ParamTable('ttLangue',taCreat,AfNoAideTablette('TTLANGUE'),PRien) ;

    ////// Paramétres honoraires & commerciaux
    74201 : AFLanceFiche_ProfilGener('TYPE:DEF');
    74202 : if VH_GC.GAAchatSeria =False then AglLanceFiche ('GC','GCCODECPTA','','','NOACHAT')                // Ventilation comptable GCCOMPTESHT_MUL
                      else AglLanceFiche ('GC','GCCODECPTA','','','') ;               // Ventilation comptable GCCOMPTESHT_MUL
    74203 : AGLLanceFiche('GC','GCUNITEMESURE','','','');
    74204 : ParamTable ('GCEMPLOIBLOB',taCreat,AfNoAideTablette('GCEMPLOIBLOB'),PRien) ;
    {$IFNDEF EAGLCLIENT}
    74205 : EntreeArrondi (taModif) ;
    74206 : if GetParamSoc('SO_GCAxeAnalytique') then AGLLanceFiche('MBO','VENTILANA','','','') // ventil générique
            else  begin
              if V_PGI.SAV then AFLanceFiche_VentilAna // ancienne ventilation GI/GA
              else  PgiInfo('Il faut utiliser la nouvelle analytique. Merci d''appeler l''assistance pour mise en place','');
              end;
    {$ENDIF}
    74207 : ParamTable ('AFJUSTIFBONI',taCreat,AfNoAideTablette('AFJUSTIFBONI'),PRien,3);  // justificatif boni/mali
    74211 : ParamTable ('GCCOMMENTAIREENT',taCreat,AfNoAideTablette('GCCOMMENTAIREENT'),PRien,6);   // commentaires entete
    74212 : begin  //mcd 23/01/02
         if GetParamSoc('SO_GCCOMMENTAIRE') then AGLLanceFiche ('GC','GCCOMMENTAIRE','','','')
                                                 else ParamTable ('GCCOMMENTAIRELIGNE',taCreat,AfNoAideTablette('GCCOMMENTAIRELIGNE'),PRien,6) ;
            end;
    74213 : ParamTable ('GCCOMMENTAIREPIED',taCreat,AfNoAideTablette('GCCOMMENTAIREPIED'),PRien,6);  // commentaires pied
    74208 : AGLLanceFiche('GC','GCPORT_MUL','','','') ;   // Port et frais
    ////// Paramétres affaires
    {$IFNDEF EAGLCLIENT}
    74751 : AFLanceFiche_BlocageAff ;      // blocage affaire
    {$ENDIF}
    74752 : AFLanceFiche_FormuleCutOff (''); // saisie de la formule de calcul des cut off

    //ONYX
    74754 : AFLanceFicheFormuleVar_Mul; // saisie de la formule de calcul des cut off

    ////// Modèles d'édition
    {$IFNDEF EAGLCLIENT}
    74251 : begin EditDocument('L','ADE','ADE',True) ; RetourForce:=True ;end; // Propositions d'affaire
    74252 : begin EditEtat('E','APE','AP1',True, nil, '', ''); RetourForce:=True ;end;
    74253 : begin
            AFLanceFiche_DocExterneMul ('', 'TYPE=WOR;NATURE=PMI'); // gestion documents parser pour propositions missions
            end;
     74261 : begin EditDocument('L','AFF','AF1',True) ; RetourForce:=True ;end; // affaire
    74262 : begin EditEtat('E','AFE','AF1',True, nil, '', ''); RetourForce:=True ;end;
    74263 : begin
            AFLanceFiche_DocExterneMul ('', 'TYPE=WOR;NATURE=LMI'); // gestion documents parser pour lettres missions
            end;
    74271 : begin EditDocument('L','GPI','AFC',True) ; RetourForce:=True ;end; // Factures
    74272 : begin EditEtat('E','GPJ','AFC',True, nil, '', ''); RetourForce:=True ;end;
    74280 : begin EditEtat('E','ALI','',True, nil, '', ''); RetourForce:=True ;end;  // etats libres GI/GA

    //$$$JP: gestion document excel OLE
    74281 : AFLanceFiche_DocExterneMul ('', 'TYPE=EXC;NATURE=ESO'); // gestion documents excel pour lien OLE
    {$ENDIF} // eaglClient

    ////// Paramétres liés aux Fiches de bases
    // Client
    //74311 : ParamTable('DPCODENAF',taCreat,0,PRien,4) ;
    74311 : ParamTable('YYCODENAF',taCreat,0,PRien,3,'Code NAF') ;
    74312 : ParamTable('GCCOMPTATIERS',taCreat,AfNoAideTablette('GCCOMPTATIERS'),PRien) ;
    74313 : ParamTable ('GCZONECOM',taCreat,AfNoAideTablette('GCZONECOM'),PRien);
    74314 : ParamTable('TTSECTEUR',taCreat,AfNoAideTablette('TTSECTEUR'),PRien) ;       // secteurs d'activité
    74315 : ParamTable('TTTARIFCLIENT',taCreat,AfNoAideTablette('TTTARIFCLIENT'),PRien) ;
    74316 : ParamTable('GCORIGINETIERS',taCreat,AfNoAideTablette('GCORIGINETIERS'),PRien) ;
    // Contacts
    74321 : ParamTable('ttFonction',taCreat,AfNoAideTablette('TTFONCTION'),PRien) ;
    74322 : ParamTable('ttLienParent',taCreat,AfNoAideTablette('TTLIENPARENT'),PRien) ;
    74323 : ParamTable('YYSERVICE',taCreat,AfNoAideTablette('YYSERVICE'),PRien) ;

    // C.B 05/12/02
    74330 : ParamTable('GCCOTEARTFOURN',taCreat,AfNoAideTablette('GCCOTEARTFOURN'),PRien) ;   // cote fournisseur
    74340 : ParamTable('TTTARIFCLIENT',taCreat,AfNoAideTablette('TTTARIFCLIENT'),PRien) ;    // code tarif client

    // Affaire
    74331 : ParamTable('AFCOMPTAAFFAIRE',taCreat,AfNoAideTablette('AFCOMPTAAFFAIRE'),PRien) ;
    74332 : ParamTable ('AFDEPARTEMENT',taCreat,AfNoAideTablette('AFDEPARTEMENT'),PRien) ;
    74333 : ParamTable ('AFTLIENAFFTIERS',taCreat,AfNoAideTablette('AFTLIENAFFTIERS'),PRien);
    74334 : ParamTable ('AFTRESILAFF',taCreat,AfNoAideTablette('AFTRESILAFF'),PRien);
    74335 : ParamTable ('AFFAIREPART1',taCreat,AfNoAideTablette('AFFAIREPART1'),PRien);
    74336 : ParamTable ('AFETATAFFAIRE',taCreat,AfNoAideTablette('AFETATAFFAIRE'),PRien);
    74337 : ParamTable ('AFTREGROUPEFACT',taCreat,AfNoAideTablette('AFTREGROUPEFACT'),PRien);
    // Articles
    74341..74343 : ParamTable('GCFAMILLENIV'+IntToStr(Num-74340),taCreat,AfNoAideTablette('GCFAMILLENIV'+IntToStr(Num-74340)),PRien) ;
    74344 : ParamTable('GCLIBFAMILLE',taModif,AfNoAideTablette('GCLIBFAMILLE'),PRien) ;
    74345 : ParamTable('GCCOMPTAARTICLE',taCreat,AfNoAideTablette('GCCOMPTAARTICLE'),PRien) ;
    74346 : ParamTable('GCTARIFARTICLE',taCreat,AfNoAideTablette('GCTARIFARTICLE'),PRien) ;
    // Ressources
    74351 : ParamTable('AFTTARIFRESSOURCE',taCreat,0,PRien) ;
    74352 : AFLanceFiche_Fonction('','');
    74353 : AFLanceFiche_Competence;
    74354 : ParamTable ('AFTNIVEAUDIPLOME',taCreat,AfNoAideTablette('AFTNIVEAUDIPLOME'),PRien);
    74355 : ParamTable ('AFTTYPERESSOURCE',taCreat,AfNoAideTablette('AFTTYPERESSOURCE'),PRien);
    // Calendrier
    74361 : ParamTable ('AFTSTANDCALEN',taCreat,AfNoAideTablette('AFTSTANDCALEN'),PRien);
    74362 : AFLanceFiche_JourFerie;
    // Activité
    74371 : ParamTable ('AFTTYPEHEURE',taCreat,AfNoAideTablette('AFTTYPEHEURE'),PRien);
    // Tâches
    74381 : ParamTable ('AFFAMILLETACHE',taCreat,0,PRien);

    /////// Tables libres //////////
    74402 : AppelTitreLibre('CT',Prien);  // Libellé tables libres client
    74403 : AppelTitreLibre('CM',Prien);  // Libellé mtt libres client
    74404 : AppelTitreLibre('CD',Prien);  // Libellé dates libres client
    74405 : AppelTitreLibre('CC',Prien);  // Libellé textes libres client
    74406 : AppelTitreLibre('CB',Prien);  // Libellé bool libres client
    74407 : AppelTitreLibre('CR',Prien);  // Libellé ressources libres client
    74456 : AppelTitreLibre('PT',Prien);  // Libellé tables libres pièces
    74457 : AppelTitreLibre('PD',Prien);  // Libellé dates libres pieces
    74481 : AppelTitreLibre('MT',Prien);  // Libellé tables libres affaire
    74482 : AppelTitreLibre('MM',Prien);  // Libellé mtt libres affaire
    74483 : AppelTitreLibre('MD',Prien);  // Libellé dates libres affaire
    74484 : AppelTitreLibre('MC',Prien);  // Libellé textes libres affaire
    74485 : AppelTitreLibre('MB',Prien);  // Libellé bool libres affaire
    74486 : AppelTitreLibre('MR',Prien);  // Libellé ressource libres affaire
    74491 : AppelTitreLibre('RT',Prien);  // Libellé tables libres ressource
    74492 : AppelTitreLibre('RM',Prien);  // Libellé mtt libres ressource
    74493 : AppelTitreLibre('RD',Prien);  // Libellé dates libres ressource
    74494 : AppelTitreLibre('RC',Prien);  // Libellé textes libres ressource
    74495 : AppelTitreLibre('RB',Prien);  // Libellé bool libres ressource
    74501 : AppelTitreLibre('AT',Prien);  // Libellé tables libres article
    74502 : AppelTitreLibre('AM',Prien);  // Libellé mtt libres article
    74503 : AppelTitreLibre('AD',Prien);  // Libellé dates libres article
    74504 : AppelTitreLibre('AC',Prien);  // Libellé textes libres article
    74505 : AppelTitreLibre('AB',Prien);  // Libellé bool libres article
    74511 : AppelTitreLibre('ET',Prien);  // Libellé tables libres etablissement
    74512 : AppelTitreLibre('EM',Prien);  // Libellé mtt libres etablissement
    74513 : AppelTitreLibre('ED',Prien);  // Libellé dates libres etablissement
    74514 : AppelTitreLibre('EC',Prien);  // Libellé textes libres aetablissement
    74515 : AppelTitreLibre('EB',Prien);  // Libellé bool libres etablissement
    74522 : AppelTitreLibre('VM',Prien);  // Libellé mttlibres apporteur
    74523 : AppelTitreLibre('VD',Prien);  // Libellé dates libres apporteur
    74531 : AppelTitreLibre('BT',Prien);  // Libellé tables libres contact
    74532 : AppelTitreLibre('BM',Prien);  // Libellé mtt libres contact
    74533 : AppelTitreLibre('BD',Prien);  // Libellé dates libres contact
    74534 : AppelTitreLibre('BC',Prien);  // Libellé textes libres contact
    74535 : AppelTitreLibre('BB',Prien);  // Libellé bool libres contact
    74551 : AppelTitreLibre('FT',Prien);  // Libellé tables libres fournissuer
    74552 : AppelTitreLibre('FM',Prien);  // Libellé mtt libres fournisseur
    74553 : AppelTitreLibre('FD',Prien);  // Libellé dates libres fournisseur
    74571 : AppelTitreLibre('TT',Prien);  // Libellé tables libres Tâches
    74572 : AppelTitreLibre('TD',Prien);  // Libellé dates libres taches
    74573 : AppelTitreLibre('TC',Prien);  // Libellé textes libres taches
    74574 : AppelTitreLibre('TB',Prien);  // Libellé bool libres taches
    74575 : AppelTitreLibre('TM',Prien);  // Libellé montants taches C.B 05/08/02

    74411..74419 : ParamTable('GCLIBRETIERS'+IntToStr(Num-74410),taCreat,AfNoAideTablette('GCLIBRETIERS'+IntToStr(Num-74410)),PRien,6,RechDom('GCZONELIBRE','CT'+IntToStr(Num-74410),FALSE)) ;  // Stats clients
    74420 : ParamTable('GCLIBRETIERSA',taCreat,AfNoAideTablette('GCLIBRETIERSA'),PRien,6,RechDom('GCZONELIBRE','CTA',FALSE)) ;  // Stats clients
    74461..74469 : ParamTable('AFTLIBREAFF'+IntToStr(Num-74460),taCreat,AfNoAideTablette('AFTLIBREAFF'+IntToStr(Num-74460)),PRien,6,RechDom('GCZONELIBRE','MT'+IntToStr(Num-74460),FALSE)) ;   // Stats Affaire
    74489 : ParamTable('AFTLIBREAFFA',taCreat,AfNoAideTablette('AFTLIBREAFFA'),PRien,6,RechDom('GCZONELIBRE','MTA',FALSE)) ;   // Stats Affaire
    74431..74439 : ParamTable('AFTLIBRERES'+IntToStr(Num-74430),taCreat,AfNoAideTablette('AFTLIBRERES'+IntToStr(Num-74430)),PRien,6,RechDom('GCZONELIBRE','RT'+IntToStr(Num-74430),FALSE)) ;   // Stats Ressource
    74499 : ParamTable('AFTLIBRERESA',taCreat,AfNoAideTablette('AFTLIBRERESA'),PRien,6,RechDom('GCZONELIBRE','RTA',FALSE)) ;   // Stats Ressource
    74441..74449 : ParamTable('GCLIBREART'+IntToStr(Num-74440),taCreat,AfNoAideTablette('GCLIBREART'+IntToStr(Num-74440)),PRien,6,RechDom('GCZONELIBRE','AT'+IntToStr(Num-74440),FALSE)) ;    // Stats Article
    74450 : ParamTable('GCLIBREARTA',taCreat,AfNoAideTablette('GCLIBREARTA'),PRien,6,RechDom('GCZONELIBRE','ATA',FALSE)) ;  // Stats article 10
    74451..74453 : ParamTable('GCLIBREPIECE'+IntToStr(Num-74450),taCreat,AfNoAideTablette('GCLIBREPIECE'+IntToStr(Num-74450)),PRien,6,RechDom('GCZONELIBRE','PT'+IntToStr(Num-74450),FALSE)) ;  // Stats Pièces
    74470..74478 : ParamTable('YYLIBREET'+IntToStr(Num-74469),taCreat,AfNoAideTablette('YYLIBREET'+IntToStr(Num-74469)),PRien,6,RechDom('GCZONELIBRE','ET'+IntToStr(Num-74469),FALSE)) ;
    74479 : ParamTable('YYLIBREETA',taCreat,AfNoAideTablette('YYLIBREETA'),PRien,6,RechDom('GCZONELIBRE','ETA',FALSE)) ;
    74541..74549 : ParamTable('YYLIBRECON'+IntToStr(Num-74540),taCreat,AfNoAideTablette('YYLIBRECON'+IntToStr(Num-74540)),PRien,3,RechDom('GCZONELIBRE','BT'+IntToStr(Num-74540),FALSE)) ;
    74560 : ParamTable('YYLIBRECONA',taCreat,AfNoAideTablette('YYLIBRECONA'),PRien,6,RechDom('GCZONELIBRE','BTA',FALSE)) ;  // Stats contatc 10
    74561..74563 : ParamTable('GCLIBREFOU'+IntToStr(Num-74560),taCreat,AfNoAideTablette('GCLIBREFOU'+IntToStr(Num-74560)),PRien,6,RechDom('GCZONELIBRE','FT'+IntToStr(Num-74560),FALSE)) ;  // Stats clients
    74581..74589 : ParamTable('AFLIBRETACHE'+IntToStr(Num-74580),taCreat,0,PRien,6,RechDom('GCZONELIBRE','TT'+IntToStr(Num-74580),FALSE)) ;   // Stats tache
    74590        : ParamTable('AFLIBRETACHEA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','TTA',FALSE)) ;   // Stats tache
  {$IFNDEF SANSCOMPTA}
    {$IFDEF V530}
   74801 : GestionSociete(PRien,@InitSociete) ;
    {$ELSE}
   74801 : GestionSociete(PRien,@InitSociete,Nil) ;
    {$ENDIF}
  {$ENDIF}

   {$IFNDEF EAGLCLIENT}
   74802 : RecopieSoc(PRien,True) ;                 // recopie société à société

    /////// Utilitaires
    74601 : AGLLanceFiche('YY','YYJNALEVENT','','','') ; // Journal évènements
    74602 : Begin AssistRecupTempo; RetourForce := True; End;
    74603 : AGLLanceFiche('GC','GCCPTADIFF','','','') ;  // comptabilisation différée
  {$IFDEF DP}
    74604 : AFLanceFiche_MulAnnuTiersFaux ;  // Lien annuaire/tiers erroné
    74605 : AFLanceFiche_CltSansLienAnn ;  // Lien annuaire/tiers plus rapide
    74606 : AFLanceFiche_MulAnnTiersDOuble ;  // Lien annuaire/tiers erroné
    74607 : AFLanceFiche_MulDoublonsAnn;  // Doublons annuaire
  {$ENDIF}
    74608 : begin   // pour chgmt fact éclatée
            V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
            AFLanceFiche_FAct_Eclat ;
//             RetourForce := True;
             V_PGI.ZoomOLE := False;
             end;
    74612 : begin   // pour chgmt option unité activite
            V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
            AFLanceFiche_Modif_Unite ;
//             RetourForce := True;
             V_PGI.ZoomOLE := False;
             end;
    74614 : Entree_TraiteAdresse; // option pour basculer de table ADRESSE sur PIECEADRESSE
    74617 : begin   // pour chgmt option première semaine
            V_PGI.ZoomOLE := True;   //pour passer la fiche en modal
            AFLanceFiche_Modif_PremSem ;
//             RetourForce := True;
             V_PGI.ZoomOLE := False;
             end;
    74618 : AglLanceFiche('YY','EDIT_MCD','','',''); // impression MCD Affaire
    74615  : AnnuleCutOff;  //mcd 27/05/03 annulation table cut off pour changer option param
    74616  : EntreeMajCompteurSouche; // mcd 25/07/03 chgmt souche pour les pièces
 {$ifdef GIGI}
    74619 : begin  //fct pour annuler la GI
            AnnuleGi;
    {$ifndef AGL560E}
            FMenuG.ForceClose := True;
            FMenuG.Close;
    {$else}
            FMenuG.Quitter;  //mcd 17/06/03
    {$endif}
            PostMessage (Fmenug.handle,WM_CLOSE,0,0);
            Retourforce:=True;
            SortieHalley:=true;
            end;
  {$endif}
    74613 :
        if GetParamSoc('SO_GCAxeAnalytique') then GIGABasculeAna (True)  // pour chgmt analytique  GA en Analmytique Générique
            else PgiBoxAf('Option non disponible','');
  {$ifdef GIGI}
    74609 : AGlLanceFiche ('YY','YYBOBINTEGRE','','','GIS5');
  {$endif}
           // Test en debug uniquement ...
        74610 : begin
                  AGLLanceFiche('AFF','EXERCICE','','','');
                   end;
        74611 : EntreeListeSaisie ('');
    74620 : AGLLanceFiche('YY','PROFILETABL','','','ETA') ; //Restrictions utilisateurs
    74699 : AGLLanceFiche('AFF','OUTILSDEBUG','','','') ;  // outils debug
    {$ENDIF}
     
    /////// Gestion des Utilisateurs
    {$IFNDEF EAGLCLIENT}
    74701 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ; //utilisatuers
    74702 :  AGLLanceFiche('YY','YYUSERGROUP','','','')  {mcd 21/03/02 FicheUserGrp} ;                          // Groupes d'utilisateurs
    74703 : ReseauUtilisateurs(False) ;             // utilisateurs connectés
    74704 : VisuLog ;                               // Suivi d'activité
    74705 : ReseauUtilisateurs(True) ;              // RAZ connexions
    //mcd 12/09/03 74706 : AGLLanceFiche('YY','MDUSERCONF','','',''); // Confidentialité
    74706 : AGLLanceFiche('YY','YYMULUSERCONF','','',''); // Confidentialité
    {$ENDIF}

    74707 : ParamTable('YYGROUPECONF',taCreat,0,PRien,3); // Groupes de conf
    74708 : AGLLanceFiche('GC','GCPARAMOBLIG','','','') ;      // Champs obligatoires
    74709 : AGLLanceFiche('GC','GCPARAMCONFID','','','') ;       // restricyions fiches
    74710 : ParamFavoris (ModulesInterdits,TagsInterdits,FALSE,False) ;  // Gestion des favoris : tout module confondu et on voit tout
                                                         // True/false : pour favori par module en voyant tous les module spour choix,
                                                         // true/true pour favori par module en ne voyant que les options du module
//    74711 : if ctxscot in V_PGI.PgiContexte then AGLLanceFiche('YY','YYCONF_VUE','','',' AND(mn_2 in(141,142,143,144,154,152,151))')
//              else AGLLanceFiche('YY','YYCONF_VUE','','','and (mn_2 in(71,72,73,74,153,138,139,140))') ; // consultation global confidentialié
    74711 : if ctxscot in V_PGI.PgiContexte then
              begin
              ListMenu :='141;142;143;144;151';
              If Vh_GC.GAPlanningSeria then ListMenu := ListMenu+ ';154';
              If Vh_GC.GAAchatSeria then ListMenu := ListMenu+ ';152';
              GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','ACTION=CONSULTATION;'+ListMenu); // consultation global confidentialié
              end
            else
              begin
              ListMenu :='71;72;73;74';
              If Vh_GC.GAPlanningSeria then ListMenu := ListMenu+ ';153';
              If Vh_GC.GAAchatSeria then ListMenu := ListMenu+ ';138';
              If VH_GC.GRCSeria then ListMenu := ListMenu+ ';92';
              ListMenu := ListMenu+ ';139;140';
              {$IFNDEF CCS3}
              ListMenu := ListMenu+ ';262';
              {$ENDIF}
              GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','',ListMenu); // consultation global confidentialié
             end;


    74712 : ShowAdminHdtLinks(Prien); // saisie des tablettes hierarchiques
    74713 : ShowTreeLinks(Prien); // saisie des hierarchie dans les tablettes
   ////// Imports          
    {$IFNDEF EAGLCLIENT}
    74651 : AGLLanceFiche('E','EIMPORTTIERS','','','');      // import E-Tiers
    74652 : AGLLanceFiche('E','EIMPORTCMD','','','');        // import E-Commandes
    74653 : AFLanceFiche_Mul_EActivite (''); // import E-Activité
    74654 : AFLanceFiche_Suivi_EActivite; //suivi saisie décentralisée
// gm 15/05/03 plus utilisé    74655 : AFLanceFiche_Mul_ExportAff; // remplacé par 74654 en GI, à supprimer qd n'existe plus en GA export par affaires
    74656 : AFLanceFiche_MulEAffaire; // import E_affaires
    {$ENDIF}

    74753 : AFLanceFiche_MulActivitePaie; // Paramétrage lien Activité avec Rubriques Paie
    74998 : CalculetteSimple;

   // onyx
   262101 : AFLanceFiche_MulPublication ;
   262201 : AFLanceFiche_MulFormule  ;
   262301 : AFLanceFiche_MulIndice ;
   262302 : AFLanceFiche_MulValIndice ;
   262303 : AFLanceFiche_AFREVPREPAIMPORT;
   262304 : AFLanceFiche_AFREVEXPORTINDICE;

   262401 : AFLanceFicheRevCalcCoef;
   262402 : AFLanceFicheModifCoef('MODIFIER');
   262403 : AFLanceFicheModifCoef('APPLIQUER');
   262404 : FlagueLesLignesRegularisables(false);
   262405 : AFLanceFicheRegulFacture;

   262501 : LanceEtatDernieresValeursIndicesSaisies;
   262502 : LanceEtatIndicesNonSaisis;
   262503 : LanceEtatFactureARegulariser;

// Par défaut
		else if not TraiteMenuSpecif(Num) then HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT, Range : String) ;
var
 ii : integer;

BEGIN

  if ((Num=8) and (TT='GCTIERSSAISIE'))  then
  BEGIN
    ii:=TTToNum(TT) ;
    if (ii>0) and (pos('T_NATUREAUXI="FOU"',V_PGI.DECombos[ii].Where)>0) then num:=12  ;
  END;

  // alimente messgae plus titre ==> peut être changer d'un seul coup
  case Num of
    1: PgiInfo ('Vous n''avez pas le droit d''utiliser cette option. Passer par la comptabilité','Acces Interdit');
  {$IFNDEF GCGC}
    1 : {Compte gene} FicheGene(Nil,'',LeQuel,Action,0) ;
    2 : {Tiers compta} FicheTiers(Nil,'',LeQuel,Action,1) ;
    4 : {Journal} FicheJournal(Nil,'',Lequel,Action,0) ;
  {$ENDIF}
    5 : {affaires}affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    6 : {ressources}  affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    7 : {article}affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    8 : {Clients par T_TIERS} affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    9 : {commerciaux} If AglJaiLeDroitFiche (['COMMERCIAL',ActionToString(Action)],2)then AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)) ;
    10 : {conditionnement} AGLLanceFiche('GC','GCCONDITIONNEMENT','',Lequel,ActionToString(Action)) ;
    11 : {catalogue} AGLLAnceFiche('GC','GCCATALOGU_SAISI3','',Lequel,ActionToString(Action)) ;
    12 : {Fournisseurs via T_TIERS} affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    13 : {dépots} AGLLanceFiche('GC','GCDEPOT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
    14 : {apporteurs} If AglJaiLeDroitFiche (['COMMERCIAL',ActionToString(Action)],2)then  AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)) ;
    15 : {Règlement}  FicheRegle_AGL ('',False,taModif) ;
    16 : {contacts} affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    17 : {Fonction des ressources} If AglJaiLeDroitFiche (['FONCTION',ActionToString(Action)],2) then AFLanceFiche_Fonction(Lequel,ActionToString(Action)) ;
    18 : {Etablissements} AGLLanceFiche('GC','GCETABLISS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
    19 : {Contenu tables libres} AGLLanceFiche('GC','GCCODESTATPIECE','','','ACTION=CREATION;YX_CODE='+Range) ;
    21 : {Ports} AGLLanceFiche('GC','GCPORT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
    22 : {GRC-action} AGLLanceFiche('RT','RTACTIONS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
    23 : {GRC-operation} AGLLanceFiche('RT','RTOPERATIONS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
    28 : {Clients par T_AUXILIAIRE}  affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
    29 : {Fournisseurs via T_AUXILIAIRE} AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=FOU') ;
{$IFDEF GRC}
   30..40 : RTMenuDispatchTT ( Num, Action, Lequel, TT, Range );
{$ENDIF}
    // ****** maj des tablettes *******
    900 : ParamTable('GCCOMMENTAIRELIGNE',taCreat,AfNoAideTablette(TT),Nil,6);  // mis de manière explicite car appellé depuis un Grid
       // pour mise à jour par defaut des tablettes
    996 : if JaiLeDroitTable(TT) then ParamTable(TT,taModif,AfNoAideTablette(TT),Nil,6) ;   {choixext}
    997 : if JaiLeDroitTable(TT) then ParamTable(TT,taCreat,AfNoAideTablette(TT),Nil,6) ;   {choixext}
    998 : if JaiLeDroitTable(TT) then ParamTable(TT,taModif,AfNoAideTablette(TT),Nil,3) ;
    999 : if JaiLeDroitTable(TT) then ParamTable(TT,taCreat,AfNoAideTablette(TT),Nil,3) ;
    // pour Annuaire
    85230 : AGLLanceFiche('JUR','GRPSOC',Range,Lequel,'') ;

    85942 : ParamTable('YYANNULIBRE1',taCreat,0,nil,3) ;
    85943 : ParamTable('YYANNULIBRE2',taCreat,0,nil,3) ;
    85944 : ParamTable('YYANNULIBRE3',taCreat,0,nil,3) ;

    85951 : ParamTable('YYVOIENOCOMPL',taCreat,0,nil,3) ;
    85952 : ParamTable('YYVOIETYPE',taCreat,0,nil,3) ;
  end;
END;

procedure AppelTitreLibre (Name : string;PRien : THPanel );
var aide : integer;
begin
  // fct qui permet de ne saisir que la partie voulue de la tablette des titres libres
  // similaire à la fct FactoriseZl de mdispgc.
  // on pourra par la suite tester Name mpour alimenter n° aide
  GCModifTitreZonesLibres (Name);
  // mcd 11/09/02 ajout n° aide
  if copy(Name,1,1)='M' then Aide:=120000152
  else if copy(Name,1,1)='R' then Aide:=120000154
  else if copy(Name,1,1)='C' then Aide:=110000242
  else if copy(Name,1,1)='A' then Aide:=110000241
  else if copy(Name,1,1)='E' then Aide:=110000245
  else if copy(Name,1,1)='P' then Aide:=110000248
  else if copy(Name,1,1)='B' then Aide:=110000244
  else if copy(Name,1,1)='F' then Aide:=110000243
  else if copy(Name,1,1)='V' then Aide:=0
  else if copy(Name,1,1)='T' then Aide:=0
  else  Aide:=0 ;
  ParamTable('GCZONELIBRE',taModif,Aide,PRien,3);
  GCModifTitreZonesLibres ('');
  if ctxScot in V_PGI.PGIContexte then AfterChangeModule(144)
                                  else AfterChangeModule(74);
end;

procedure AfterChangeModule ( NumModule : integer ) ;
Var MenuPar : boolean;
BEGIN
/// Menu Pop général
TraduitMenu(NumModule);
    // mcd 22/05/02, il ne faut pas de menu contextuel sur les menus paramètres/administration
MenuPar:=false;
if (ctxscot in V_PGI.PGIContexte) then begin if (NumModule = 144) then Menupar:=true;   end
   else if NumModule=140 then MenuPar:=true;

If MenuPar then   ChargeMenuPop(0,FMenuG.DispatchX)
 else ChargeMenuPop(7,FMenuG.DispatchX) ;
if NumModule=92 then ChargeMenuPop(8,FMenuG.DispatchX) ;

{$IFDEF CCS3}
TripoteMenuGAS3(NumModule) ;
{$ENDIF}
{$IFDEF GRC}
RTMenuTraiteMenusGRC(NumModule);
{$ENDIF}
END;

Function ChargeFavoris : boolean;
begin
AddGroupFavoris(FMenuG,'') ;
result:=true;
end;

Procedure AfterProtec ( sAcces : String ) ;
var SS,SS1  : Array of integer;
    ii      : integer;
    i       : Integer;
             
BEGIN

  // cas ou on coche "version non sérialiser' dans
  // l'écran de séria. On force la non sérialisation
  If V_PGI.NoProtec then V_PGI.VersionDemo:= TRUE
                   else V_PGI.VersionDemo :=False;

    // si le 1er module n'est pas sérialiser, on est en démo, sur tous les modules
  if (sAcces[1]='-') then   V_PGI.VersionDemo:= TRUE; // on repasse en demo dans ce cas


  VH_GC.GAPlanningSeria:=False;
  VH_GC.GAStdSeria :=False;
  VH_GC.GAAchatSeria:=False;
  VH_GC.GASaisieDecSeria:=False;
  VH_GC.GRCSeria:=False ;
  V_PGI.NbColModuleButtons:=1 ;
  V_PGI.NbROwModuleButtons:=6 ;  // valeurs par defaut

  if ctxScot in V_PGI.PGIContexte then
     begin
       if (sAcces[2]='X') then   VH_GC.GAPlanningSeria:= TRUE;
       if (sAcces[3]='X') then   VH_GC.GAAchatSeria:= TRUE;
       if (sAcces[4]='X') then   VH_GC.GASaisieDecSeria:= TRUE;
{$ifdef GIGI}
       FMenuDisp.SeriaGEdOK:= False;
       if (sAcces[5]='X') then   FMenuDisp.SeriaGEdOK:= TRUE;
{$endif}

  {$IFDEF GRC}
       VH_GC.GRCSeria:=(sAcces[6]='X')  ;
       if V_PGI.VersionDemo then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC]
        else if VH_GC.GRCSeria then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC] ;
  {$ENDIF}
  {$ifdef DIFUSGI}
   VH_GC.GAPlanningSeria :=False; // pas de planning en GI 01/2003
  {$endif}

     end
     else begin  // GA
       VH_GC.GAseria := false ;
       VH_GC.GASaisieDecSeria :=false;
       VH_GC.GAPlanningSeria := false;
       VH_GC.GAAchatSeria:=false;

			 if (sAcces[1]='X') then   VH_GC.GAseria:= TRUE;
       {$IFDEF CCS3}
       if (sAcces[3]='X') then   VH_GC.GAPlanningSeria:= TRUE;
       if (sAcces[4]='X') then   VH_GC.GAAchatSeria:=True;
       {$ELSE}
       if (sAcces[3]='X') then   VH_GC.GASaisieDecSeria:=True;
       VH_GC.GAAchatSeria:=True;
       VH_GC.GAPlanningSeria := True;
       {$ENDIF}

        if Length(sAcces)>1 then VH_GC.GRCSeria:=(sAcces[2]='X') ;  // gestion de la Relation Client
        {$IFDEF GRC}
        if V_PGI.VersionDemo then
          V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC]
        else
          if VH_GC.GRCSeria then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxGRC] ;
        {$ENDIF}
     end;
 
  //  71/141 Fiches de base, 72/142 Activite reporting,
  //  73/143 facturation, 74/144 Paramétrage  ,  153/154 planning
  // 138/152 achats, 139/151 Analyses,  140/.. Administration
  if V_PGI.VersionDemo then
  begin
        // tout est accessible ,même module non sérialisé
      VH_GC.GAPlanningSeria:=True;
      {$ifdef DIFUSGI}
      VH_GC.GAPlanningSeria :=False; // pas de planning en GI 01/2003
      {$endif}
      VH_GC.GAAchatSeria:=True;
      {$IFNDEF CCS3}
      VH_GC.GASaisieDecSeria:=True;
      {$ENDIF}
      {$IFDEF GRC}
       VH_GC.GRCSeria:=True  ;
      {$endif}
      if ctxScot in V_PGI.PGIContexte then
         begin //liste moduel GI
  {$IFDEF GRC}
       FMenuG.SetModules([141,142,143,154,152,92,151,144],[1,8,35,73,38,77,11,71]);
  {$else}
       FMenuG.SetModules([141,142,143,152,151,144],[1,8,35,38,11,71]);
  {$endif}
         V_PGI.NbColModuleButtons:=2 ;
         V_PGI.NbRowModuleButtons:= ceil(7/2)
         end
       else
       begin //cas module GA
  {$ifdef EAGLCLIENT}
          //FMenuG.SetModules([71,72,73,138,139,153,92],[1,8,35,18,11,73,77]) ;
  {$else}
          //FMenuG.SetModules([71,72,73,262,138,139,153,92,74,140],[1,8,35,3,18,11,73,77,71,49]) ;
  {$endif}
         // AB-20031020 Menu eagl complèté - alignement avec menu avec séria- ajout menu spécifique
         FMenuG.SetModules([71,72,73,262,153,138,92,139,160,74,140],
                           [1, 8, 35,3,  73, 18, 77,11, 60, 71,49]) ;
         V_PGI.NbColModuleButtons:=2 ;
         V_PGI.NbRowModuleButtons:= ceil(11/2);

     end;
  end
  else
      begin  // cas version sérailisée .......
      if ctxScot in V_PGI.PGIContexte then
         begin
         ii:=5;
         if VH_GC.GAPlanningSeria then inc(ii);
         if VH_GC.GAAchatSeria then inc(ii);
         if VH_GC.GRCSeria then inc(ii);
         SetLength(SS,ii);   // on crée l'array du nbre voulu
         SetLength(SS1,ii);
                           // on affecte les modules à gérer
         SS[0]:=141; SS1[0]:=1;   // module fichier obligaotire
         SS[1]:=142; SS1[1]:=8;  // module activité obligaotire
         SS[2]:=143; SS1[2]:=35;  // module facturation obligatoire
         II:=3;
         if VH_GC.GAPlanningSeria  then begin   SS[II]:= 154; SS1[II]:= 73; inc (II);end;
         if VH_GC.GAAchatSeria  then begin   SS[II]:= 152; SS1[II]:= 38;inc (II);end;
         if VH_GC.GRCSeria then begin   SS[II]:= 92; SS1[II]:= 77;inc (II);end;
         SS[ii]:=151; SS1[ii]:=11;  // module analyse obligatoire
         inc(ii); 
         SS[ii]:=144; SS1[ii]:=71;  // module paramètres obligatoire
         end
      else begin    // cas GA
  //       FMenuG.SetModules([71,72,73,138,139,74,140],[1,8,35,18,11,71,49]) ;
         ii:=6;
         if VH_GC.GAPlanningSeria then inc(ii);
         if VH_GC.GAAchatSeria then inc(ii);
         if VH_GC.GRCSeria then inc(ii);
         inc(ii); //AB-20031021 module Spécifique
         SetLength(SS,ii);   // on crée l'array du nbre voulu
         SetLength(SS1,ii);

         // onyx
         SetLength(FNumMenus,ii);
         SetLength(FNumIcones,ii);
         FNextModule := ii + 1;

                           // on affecte les modules à gérer
         SS[0]:=71; SS1[0]:=1;   // module fichier obligaotire
         SS[1]:=72; SS1[1]:=8;  // module activité obligaotire
         SS[2]:=73; SS1[2]:=35;  // module facturation obligatoire
         II:=3;
         if VH_GC.GAPlanningSeria  then begin   SS[II]:= 153; SS1[II]:= 73; inc (II);end;
         if VH_GC.GAAchatSeria  then begin   SS[II]:= 138; SS1[II]:= 18;inc (II);end;
         if VH_GC.GRCSeria  then begin   SS[II]:= 92; SS1[II]:= 77;inc (II);end;
         SS[ii]:=139; SS1[ii]:=11;  // module analyse obligatoire
         inc(II);
         SS[II]:=160; SS1[II]:=60;  //AB-20031021 module Spécifique

         // C.B 29/09/03
         // module paramètres et Administration en eagl
         inc(ii);
         SS[ii]:=74; SS1[ii]:=71;
         inc(ii);
         SS[ii]:=140; SS1[ii]:=49;
             
    //     inc(ii);    gm pour etat libre en cours
    //    SS[ii]:=160; SS1[ii]:=1;  // module Administration obligatoire
         end;

      if (ii+1 >5) then
      begin
        V_PGI.NbColModuleButtons:=2 ;
        V_PGI.NbRowModuleButtons:= ceil((ii+1)/2)
      end;

      //onyx
      for i := 0 to FNextModule - 2 do
      begin
        FNumMenus[i] := SS[i];
        FNumIcones[i] := SS1[i];
      end;

      FMenuG.SetModules(SS,SS1);
      SS:=  Nil; // on libere array dynamique
      SS1:=Nil;

    end;

      // on renseigne les mdoules interdits en fct séria
  If CtxScot in V_PGI.PGICOntexte then
     begin
     ModulesInterdits:='144;';
     if not VH_GC.GAPlanningSeria then ModulesInterdits:= ModulesInterdits+'154;' ;
     if not VH_GC.GAAchatSeria then ModulesInterdits:= ModulesInterdits+'152;';
     if not VH_GC.GRCSeria then ModulesInterdits:= ModulesInterdits+'92;';
     end
  else begin
     ModulesInterdits:='140;74;';
     if not VH_GC.GAPlanningSeria then ModulesInterdits:= ModulesInterdits+'153;' ;
     if not VH_GC.GAAchatSeria then ModulesInterdits:= ModulesInterdits+'138;';
     if not VH_GC.GRCSeria then ModulesInterdits:= ModulesInterdits+'92;';
     end;
END;

procedure TFEvtPopUpS5.PopupMenuS5Popup(Sender: TObject);
var
  j:integer;
begin
  for j:=0 to fmenug.MenuPop.Items.Count-1 do
  begin
    fmenug.MenuPop.items[j].Caption :=TraduitGA(fmenug.MenuPop.items[j].Caption);
    fmenug.MenuPop.items[j].Enabled := JaiLeDroitTag(fmenug.MenuPop.items[j].tag);
  end;
end;

function MyProcCalcMul(Nom, params, WhereSQL : string; TT : TDataSet; Total : Boolean) : string;
var
  vSt : String;
  vQr : TQuery;

BEGIN                        
  if Nom='LIENOLE' then
  begin
    vSt := 'SELECT LO_LIBELLE FROM LIENSOLE WHERE LO_EMPLOIBLOB = "' + params + '"';
    vSt := vSt + ' AND LO_IDENTIFIANT = "' + TT.findfield('AFF_AFFAIRE').asstring + '"';
    Try
      vQR := OpenSql(vSt, True);
      if Not vQR.Eof then
        result := vQR.findfield('LO_LIBELLE').AsString
      else
        result := '';
    Finally
      if vQR <> nil then ferme(vQR);
    End
  end;
END;

  //mcd 11/09/03 pour gestion de la GED
{$ifdef GIGI}
procedure MyAfterImport (Sender: TObject; FileID: Integer; var Cancel: Boolean) ;
// Evenement après scannérisation : retourne le FileID du fichier scanné
var ParGed : TParamGedDoc;
begin
  if FileId<0 then exit;
  // ---- Insertion classique dans la GED avec boite de dialogue ----
  // Propose le rangement dans la GED
  ParGed.DocId := 0;
  ParGed.NoDossier := V_PGI_Env.NoDossier;
  ParGed.CodeGed := '';
  // FileId est le n° de fichier obtenu par la GED suite à l'insertion
  ParGed.FileId := FileId;
  // Description par défaut du document à archiver...
  if Sender is TForm then ParGed.DefName := TForm(Sender).Caption;
  Application.BringToFront;
  if ShowNewDocument(ParGed)='##NONE##' then
    // Fichier refusé, suppression dans la GED
  V_GedFiles.Erase(FileId);
end;
{$endif}


procedure InitApplication;
BEGIN

  ProcZoomEdt := AFZoomEdtEtat;
  ProcCalcEdt := GCCalcOLEEtat;

  {$IFDEF GIGI}
       GlobalBeforeTraduc:= TraduitEcranAffaire ;
       GlobalBeforeTraducTablette:= TraduitTabletteAffaire ;
  {$ENDIF}

  //ProcGetVH:=GetVS1 ; //ProcGetDate:=GetDate ;
  FMenuG.OnDispatch:=Dispatch ;
  FMenuG.OnChargeMag:=ChargeMagHalley ;
  Fmenug.OnChargeFavoris:=CHargeFavoris;
  {$IFDEF EAGLCLIENT}
       FMenuG.OnMajAvant:=Nil ;            
       FMenuG.OnMajApres:=Nil ;
  {$ELSE}
  (*FMenuG.OnMajAvant:=MAJHalleyAvant ;
  FMenuG.OnMajApres:=MAJHalleyApres ;
  FMenuG.OnMajPendant:=MajHalleyPendant ;*)
  {$ENDIF}

  FMenuG.OnChangeModule:=AfterChangeModule ;
  FMenuG.OnAfterProtec:=AfterProtec ;
  FMenuG.MenuPop.OnPopup:=FEvtPopUpS5.PopupMenuS5Popup;
  {$IFNDEF EAGLCLIENT}
  FMenuG.OnChangeUser:=OnChangeUserPGI ;
  {$ENDIF}

  {$IFDEF GIGI}  // Gestion interne
    {$IFDEF GRC}
   FMenuG.SetSeria('00280011',['00270050','00271050','00272050','00273060','00119050','05003042'],
                ['Gestion Interne','Planning','Fact. Fournisseur','Saisie décentralisée','GED','Gestion de la Relation Client']) ;  //PCL
    {$ELSE}
   FMenuG.SetSeria('00280011',['00270050','00271050','00272050','00273050','00119050'],
                ['Gestion Interne','Planning','Fact. Fournisseur','Saisie décentralisée','GED']) ;  //PCL
    {$ENDIF}
  {$ELSE} // **** Gestion d'affaires *****  CODESERIA
    {$IFDEF CCS3}
    FMenuG.SetSeria('00016011', ['00107050','00060050','00109050','00108050'],
                                ['Gestion d''Affaires S3',
                              	'Option Gestion de la Relation Client S3',
                                'Option Planning S3',
     														'Option Achat S3']) ;
   {$ELSE}
  	 FMenuG.SetSeria('00059010',['00110050','05003050','00091050'],
															  ['Gestion d''Affaires S5',
                              	'Option Gestion de la Relation Client S5',
     														'Option Saisie d''activité décentralisée S5']) ;
    {$ENDIF}
  {$ENDIF}

  FMenuG.SetPreferences(['Documents'],False) ;
  V_PGI.DispatchTT:=DispatchTT ;
  if ctxscot in V_PGI.PGIContexte then V_PGI.WidthModuleButtons := 260;     // mcd 13/01/03 pour OK largeur en mode XP

  // cb 25/02/03
  ProcCalcMul := MyProcCalcMul;

end;

{$IFDEF CCS3}
Procedure TripoteMenuGAS3 ( NumModule : integer ) ;

          Procedure RemoveTablesLibresS3 ( TT : Array of integer ) ;
          Var i : integer ;
          BEGIN
          for i:=Low(TT) to High(TT) do FMenuG.RemoveItem(TT[i]) ;
          END ;
BEGIN
        RemoveTablesLibresS3([74403,74404,74405,74406,74407,74414,74415,74416,74417,74418,74419,74420]) ;
        RemoveTablesLibresS3([74482,74483,74484,74485,74486,74464,74465,74466,74467,74468,74469,74489]) ;
        RemoveTablesLibresS3([74492,74493,74494,74495,74434,74435,74436,74437,74438,74439,74499]) ;
        RemoveTablesLibresS3([74502,74503,74504,74505,74444,74445,74446,74447,74448,74449,74450]) ;
        RemoveTablesLibresS3([74512,74513,74514,74515,74473,74474,74475,74476,74477,74478,74479]) ;
        RemoveTablesLibresS3([74457,74532,74533,74534,74535,74544,74545,74546,74547,74548,74549,74560]) ;
        RemoveTablesLibresS3([74523,74553,74552,74553]) ;
        RemoveTablesLibresS3([74572,74573,74574,74575,74584,74585,74586,74587,74588,74589,74590]) ;
 END ;
{$ENDIF}



initialization
//mcd 10/04/2003 suppression de tout ce qui est initilaisation qui est maintenant dans initga

  {$IFNDEF EAGLCLIENT}   // gmeagl à vérifier
  RegisterAffairePGI;                 // Déclaration du serveur OLE GI PGI
  {$ENDIF}

end.


