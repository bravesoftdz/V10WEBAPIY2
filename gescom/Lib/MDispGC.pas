unit MDispGC;

interface

Uses Forms,sysutils ;

Procedure InitApplication ;    // initialisation de FMenuG
//procedure InitLaVariablePGI ;  // initialisation de la variable globale V_PGI
//procedure InitAGL ;
procedure AfterChangeModule ( NumModule : integer ) ;

type
  TFMenuDisp = class(TDatamodule)
  private
    procedure HDTDrawNode (Sender: TObject; var Caption: string) ;
  end ;

Var FMenuDisp : TFMenuDisp ;
    //InitProcAGL : TProcedure ;

implementation

Uses Windows,HMsgBox, Classes, Hctrls, Controls, HPanel, UIUtil,
     ImgList,ParamSoc,AGLInit,
{$IFDEF EAGLCLIENT} ////// Unités uniquement CWAS
     Maineagl,MenuOLX,eTablette,M3FP,UtileAGL,CegidIEU,
{$ELSE}    ////// Unités uniquement NON CWAS
     FE_Main,MenuOLG,tablette,EdtQR,
     region,codepost,
     CreerSoc,CopySoc,Reseau,Mullog,LicUtil,
     EdtEtat,EdtREtat,UEdtComp,EdtDOC,EdtRdoc,
     {$IFNDEF SANSCOMPTA} devise,{$ENDIF}
     AssistInitSoc,       //Transfert, 
     UtofConsoEtat,UTofGcTraiteAdresse,UtilSoc,
     AssistDEBC,
     galOutil, // Import BOB
 {$IFNDEF SANSCOMPTA}
     Souche,Suprauxi, Tva, OuvrFerm, InitSoc, UTofRTSuppressionTiers,
 {$ENDIF}
 {$IFNDEF SANSPARAM}
     eTransferts,AssistStockAjust,
     Arrondi,ListeDeSaisie,
 {$ENDIF}
 {$IFNDEF GCGC}
     MulGene, General, Tiers, Journal,TVA,OuvrFerm,
     EtbMce,EtbUser,TeleTrans,banquecp,
 {$ENDIF}
 {$IFDEF AFFAIRE}
     ConfidentAffaire,
     {$IFNDEF EAGLCLIENT}
     GalTomUserConf,
     {$ENDIF}
 {$ENDIF}
 {$IFDEF GRC}
     RegGRCPGI,
 {$ENDIF}
 {$IFNDEF CCS3}
 {$IFNDEF MODE}
    AssistCodeAffaire,UTOFAFSUIVIEACT, UtofAFFAIREEXPORT_MUL, UTOFAFEACTIVITE_MUL,
 {$ENDIF}
 {$ENDIF}
{$ENDIF} // EAGLCLIENT

  ////// Unités commune CWAS et autres
  Transfert, 
  Facture, EntGC,Ent1,AssistSuggestion,TiersUtil,FichComm,
  ZoomEdtGescom,GCOleAuto, CalcOleGescom, PGIOleAuto,
  HistorisePiece_Tof,ListeInvContrem_TOF,InvContrem_TOF,UTOFGCLISTEINV_MUL,
  ListeInvPreCon_TOF,UTOFListeInvVal,MouvStkExContr_TOF,UTofMouvStkEx,
  UtilGC,UtomArticle,UTomPiece,ArtNonMvtes_Tof,
 {$IFNDEF SANSPARAM}
     TarifArticle,TarifTiers,TarifCliArt,TarifCatArt, UtomModePaie,
     tradarticle, dimension,
 {$ENDIF}
 {$IFDEF AFFAIRE}
     AFActivite,ActiviteUtil,
     UtofAfRessou_synchro,UTofAFRes_SyncWhat,  
     UtofpgImportFic,UTofAfActivitePai_Mul,UTomActivitePaie,
     UTofAfPaieGener,UTofAFPaieGenerChoix,UtofAfPaieAnnul, //Affaire - Lien Paie
 {$ENDIF}
 {$IFDEF CAISSE}
     MenuCaisse ,
 {$ENDIF}
 {$IFNDEF CCS3}
     VerifLigneSerie_TOF,
 {$IFNDEF MODE}
     UTOFPRINTCALREGLE,
     UTofAffaire_Mul, UTofEtatsAffaire, UTofAfJournalAct,
     UTOFTRAITEMENT_SUR_AFFAIRE, UTOFAFAFFAIR_MODIFLOT,
     UTOFAFRESSOU_MODIFLOT, UTofAfActiviteCube, UTofEtatsActivite,
     UTofAfPiece, UTOFREVAL_RESS_MUL, UtofAFARTICLEREVALMUL, AffaireUtil,
     UTOFAFSAISDATE, UtofAFGENCUTOFFAFFMUL, UtofAFMODIFCUTOFF_MUL,
     UTofAfTableauBord, UTOFETATS_PIECES, UtofAFFormule,UtofAfVentilAna,   // gm 09/01/03
     UtomAffairepiece,UTofAfEcheanceCube,UTOFCPJALECR,
 {$ENDIF}
 {$ENDIF}
 {$IFDEF MODE}
     MODEMenu,
 {$ENDIF}
 {$IFDEF GRC}
     RTMenuGRC, CalcMulProspect,RTRappel_Actions,UTreeLinks,UtofProspect_mul,
 {$ENDIF}
 {$IFDEF GPAODEV}
     wTobDebug_Tof,
 {$ENDIF}
    entPGI,
    GrpFavoris,
    UtilDispGC,UtilArticle,HEnt1,
    UtomUtilisat,UtilPGI,
    Choix,Confidentialite_TOF,
    AFFCDEENTETE_TOM,AFFCDEMODIF_TOF,AFFCDESAISIEDIM_TOF,
    AFFCDESAISIE_TOF,MBOAFFCDELANCE_TOF,MBOAFFCDESELECT_TOF,UTILAFFCDE,
    PieceEpure_tof,AssistPieceEpure,
    AssistRazActivite, AssistMajCompteurSouche, AssistAffecteDepot, VerifAuxiTiers,
    UtilULst_Tof,VerifNatErrImp,CalcMulPGI,
    {$IFDEF EDI}
      EntEDI,
      EDIPlanif,
    {$ENDIF EDI}

    CalcMulGPAO,
    {$IFDEF GPAO}
      DB,
      FactUtil,
      SaisUtil,
      uTob,
      {$IFNDEF GRC}
        uTreeLinks,
      {$ENDIF}
      wArtNat,
      wCbnInterface,
      //wCbnPropMul_Tof,  
      wChgComposant,
      wFaisabilite,
      wGammeLig,
      wGammeRes,
      wJetons,
      wNomeTet,
      wOrdreLig,
      wOrdrePhase,
      wPrixDeRevient,
      xAdminMetaSet,
      ZoomEdtGPAO,

    {$ENDIF GPAO}
    wCommuns,
{$IFDEF NOMADESERVER}
    uToxConst,   //Agl
    UtoxFiches, UtoxConf, GCTOXWORK,  ToxMoteur, GCToxCtrl, ToxSim, ToxConsultBis,ToxDetop,  //LibCO
{$ENDIF}
    Tarifs,
    wCbnPropMul_Tof,UtomFerie,
    Edit_Tarif_Tof
    ;

    const lesTagsToRemove = ''
          + '92403;92850;30700;30610;30611;-30620;-30640;30609;33106;33107;-33140;-33150;'
          + '-33160;-33170;33106;33181;33182;92403;2104;2114;'
{$IFNDEF CEGID}
          + '-33180;'
{$ENDIF}
{$IFDEF CCS3}
          + '32700;32602;33300;-33130;60206;60207;65204;65408;65305;65306;-65524;-65250;65600;92107;'
          + '92106;92108;92201;92270;92271;92272;92291;92288;-92340;92413;92414;92423;92424;92425;92440;92912;'
          + '92917;92913;92914;-92995;92946;92947;-92919;92915;92500;-92700;-92720;92740;92750;92945;-92931;92948;'
          + '-60410;-60420;'
{$ENDIF}
          ;
    const lesTagsToRemoveCegid = '33201;'
{$IFDEF CCS3}
          + '32700;32602;33300;-33130;60206;60207;65204;65408;65305;65306;-65524;-65250;65600;92107;'
          + '92106;92108;92201;92270;92271;92272;92291;92288;-92340;92413;92414;92423;92424;92425;92440;92912;'
          + '92917;92913;92914;-92995;92946;92947;-92919;92915;92500;-92700;-92720;92740;92750;92945;-92931;92948;'
{$ENDIF}
          ;
{$R *.DFM}

function ChargeFavoris : boolean ;
begin
  if not VH_GC.GCIfDefCEGID then
     AddGroupFavoris( FMenuG , lesTagsToRemove)
  else
     AddGroupFavoris( FMenuG , lesTagsToRemoveCegid) ;
  result := true;
end;

Procedure FactoriseZL ( Tag : integer ; PRien : THPanel ) ;
Var St : String ;
    iHelpContext : Longint ;
BEGIN
St:='' ;
Case Tag of
   65561 : St:='AT' ; 65562 : St:='AM' ; 65563 : St:='AD' ; 65564 : St:='AC' ; 65565 : St:='AB' ;
   65571 : St:='CT' ; 65572 : St:='CM' ; 65573 : St:='CD' ; 65574 : St:='CC' ; 65575 : St:='CB' ;
   65581 : St:='ET' ; 65582 : St:='EM' ; 65583 : St:='ED' ; 65584 : St:='EC' ; 65585 : St:='EB' ;
   65592 : St:='VM' ; 65593 : St:='VD' ;
   65595 : St:='FT' ; 65596 : St:='FM' ; 65597 : St:='FD' ;
   65525 : St:='MT' ; 65526 : St:='MM' ; 65527 : St:='MD' ; 65528 : St:='MC' ; 65529 : St:='MB' ; 65530 : St:='MR' ;
   65535 : St:='RT' ; 65536 : St:='RM' ; 65537 : St:='RD' ; 65538 : St:='RC' ; 65539 : St:='RB' ;
   65586,105901 : St:='BT' ; 65587,105902 : St:='BM' ; 65588,105903 : St:='BD' ;
   65589,105904 : St:='BC' ; 65590,105905 : St:='BB' ;
   105909 : St:='AS';
   105911 : St:='DT' ; 105912 : St:='DM' ; 105913 : St:='DD' ; 105914 : St:='DC' ; 105915 : St:='DB' ;
   else Exit ;
   End ;
iHelpContext:=0;
if (Length(St)>0) then
   Case St[1] of
   'A' :  iHelpContext:=110000241;        //Articles
   'C' :  iHelpContext:=110000242;        //Clients
   'E' :  iHelpContext:=110000245;        //Etablissements
   'V' :  iHelpContext:=0;                //Commerciaux
   'F' :  iHelpContext:=110000243;        //Fournisseurs
   'M' :  iHelpContext:=0;                //Affaire
   'R' :  iHelpContext:=0;                //Ressources
   'B' :  iHelpContext:=110000244;                //Contacts
   End;
GCModifTitreZonesLibres(St) ;
ParamTable('GCZONELIBRE',taModif,iHelpContext,PRien,3) ;
GCModifTitreZonesLibres('') ;
AvertirTable ('GCZONELIBRE');
AfterChangeModule(65); AfterChangeModule(105);
END ;

Procedure LanceEtatLibreGIGA ;
Var CodeD : String ;
BEGIN
CodeD:=Choisir('Choix d''un état libre','MODELES','MO_CODE || " " || MO_LIBELLE','MO_CODE','MO_NATURE="ALI" AND MO_MENU="X"','') ;
if CodeD<>'' then LanceEtat('E','ALI',CodeD,True,False,False,Nil,'','',False) ;
END ;

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
{$IFNDEF EAGLCLIENT}
Var StVire,StSeule,stMessage : String ;
{$ENDIF}
BEGIN
{$IFDEF AFFAIRE}
if Not(TraiteObjActivite(Num,PRien)) then Exit; // PL après V535 : gestion particulière de la saisie d'activité
{$ENDIF}

{$IFDEF MISESOUSPLI}  // Version specifique CEGID ne sait faire que les editions différées
if ((num>100) and (num <>30609)) then
   begin
   PGIBox ('Version verrouillée, uniquement édition factures pour mise sous pli');
   Exit ;
   end;
{$ENDIF}
   Case Num of
   10 : begin
{$IFNDEF EAGLCLIENT}
        if PCL_IMPORT_BOB(NomHalley) = 1 then
        begin
          stMessage := 'LE SYSTEME A DETERMINE UNE MAJ DE MENU';
          stMessage := stMessage +#13#10 +'POUR ETRE PRISE EN COMPTE';
          stMessage := stMessage +#13#10 +'VOUS DEVREZ RELANCER L''APPLICATION';
          PGIInfo(stMessage,'MAJ MENU');
        end;
{$ENDIF}
        UpdateCombosGC; GCModifLibelles; GCTripoteStatus ;
{$IFDEF GRC}
        if ctxGRC in V_PGI.PGIContexte then
        begin
          if (GetParamSoc('SO_RTAFFICHEINFOS')) then
          begin
            if (GetSynRegKey('RTMesActions','X',TRUE)='X') then
            begin
              V_PGI.ZoomOLE := true;  //Affichage en mode modal
              AGLLanceFiche ('RT','RTINFOSRESPONS','','','');
              V_PGI.ZoomOLE := False;  //Affichage en mode modal
              { Affichage automatiques des Rappels des Actions }
              if (GetParamSoc('SO_RTACTRAPPELAUTO')) then
                if (GetSynRegKey('RTRappelAuto','X',TRUE)='X') then
                  RTInitTHreadActions(False);
            end
            else
            begin
              { Affichage automatiques des Rappels des Actions }
              if (GetParamSoc('SO_RTACTRAPPELAUTO')) then
                if (GetSynRegKey('RTRappelAuto','X',TRUE)='X') then
                  RTInitTHreadActions(True);
            end;
          end
          else
            { Affichage automatiques des Rappels des Actions }
            if (GetParamSoc('SO_RTACTRAPPELAUTO')) then
              if (GetSynRegKey('RTRappelAuto','X',TRUE)='X') then
                RTInitTHreadActions(true);
        end;
{$ENDIF}
{$IFDEF GESCOM}
        if GetParamsoc('SO_MONOFOURNISS') then SetParamSoc('SO_MONOFOURNISS',False) ;
{$ENDIF}
{$IFDEF MODE}
        VH_GC.GcMarcheVentilAna:='MODE';
{$ELSE}
         VH_GC.GcMarcheVentilAna:='AFF';     // gm 09/01/03 même pour la GC ,il faut mettre AFF
{$ENDIF}
        GCTestDepotOblig;
        end;       //  après connexion
   11 :  ; //  après déconnexion
   12 : ; //  avant connexion et séria
{$IFDEF EAGLCLIENT}
   13 : begin
       {$IFDEF CAISSE}
          LiberationCaisse;
       {$ENDIF}
          ChargeMenuPop(-1,FMenuG.DispatchX) ;
        end;
{$ELSE}
   13 : begin
       {$IFDEF CAISSE}
          LiberationCaisse;
       {$ENDIF}
          ChargeMenuPop(-1,FMenuG.DispatchX) ;
        end;
{$ENDIF}
   15 : ; //  ??????
   16 : ; // entre connexion et chargemodule
   100 : ;  // Si lanceur
/////// Menu Pop ////////////////
   27080 : AGLLanceFiche('GC','GCPIECE_MUL','GP_VENTEACHAT=VEN;GP_NATUREPIECEG=DE','','CONSULTATION') ;  // devis pour GRC
///////  Ventes /////////////
   // Saisies pièces
   30201 : CreerPiece('DE') ; // Devis
   30202 : CreerPiece('PRO') ; // Proforma
   30241 : CreerPiece('CC') ;  // Commande
   30242 : CreerPiece('CCE') ;  // Commande échantillon
   30209 : CreerPiece('PRE') ;  // Préparation livraison
   30251 : CreerPiece('BLC') ; // Livraison
   30252 : CreerPiece('LCE') ; // Livraison échantillon
   30210 : CreerPiece('FPR') ; // factures provisoires
   30205 : CreerPiece('FAC') ; // Facture
   30261 : CreerPiece('AVC') ; // Avoir
   30262 : CreerPiece('AVS') ; // Avoir sur stock
   30263 : CreerPiece('APR') ; // Avoir Provisoire
   // mofification pièces
   30221 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=DE','','MODIFICATION') ;   //devis
   30222 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=PRO','','MODIFICATION') ;  //Proforma
   30223 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=CC','','MODIFICATION') ;   //Commandes
   30226 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=CCE','','MODIFICATION') ;   //Commandes échantillon
   30229 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=PRE','','MODIFICATION') ;   //Prépa
   30224 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=BLC','','MODIFICATION') ;  //Livraison
   30227 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=LCE','','MODIFICATION') ;  //Livraison échantillon
   30225 : AGLLanceFiche('AFF','AFPIECE_MUL','GP_NATUREPIECEG=FPR','','TABLE:GP;NOCHANGE_NATUREPIECE;NATURE:FPR;ACTION=MODIFICATION;');   //Factures provisoires
   // Duplication unitaire pièces
   30231 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=DE','','DE') ;   // Devis
   30232 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=PRO','','PRO') ;  // Proforma
   30233 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=CC','','CC') ;   // Commandes
   30239 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=PRE','','PRE') ;   // Prépas
   30234 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=BLC','','BLC') ;  // Livraisons
   30235 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=FAC','','FAC') ;  // Factures
   30236 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AVC','','AVC') ;  // Avoirs
   30238 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AVS','','AVS') ;  // Avoirs sur stock
   30237 : AGLLanceFiche('AFF','AFDUPLICPIECE_MUL','GP_NATUREPIECEG=FPR','','NATURE:FPR;DUPPLIC:FPR;');  // Factures  provisoires
   30211 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','VEN') ; //Histirisation des pièces
   // Consultation
   30301 : AGLLanceFiche('GC','GCPIECE_MUL','GP_VENTEACHAT=VEN;GP_NATUREPIECEG=FAC','','CONSULTATION') ;
   30302 : AGLLanceFiche('GC','GCLIGNE_MUL','GL_NATUREPIECEG=FAC','','CONSULTATION') ;
   // Génération unitaire pièces
   30401 : AGLLanceFiche('GC','GCPIECEVISA_MUL','GP_NATUREPIECEG=CC;GP_VENTEACHAT=VEN','','VENTEACHAT=VEN') ; // Visa des pièces
   30403 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=DE','','CC') ;   // Commandes
   30409 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=CC','','PRE') ;   // Prépas
   30421 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=PRE','','BLC') ;  // Livraisons
   30422 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=CCE','','LCE') ;  // Livraisons échantillon
   30405 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=BLC','','FAC') ;  // Factures
   // Génération manuelle pièces
   30431 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=CC','','VENTE;CCR;BLC') ;  // Comm --> Livr
   30432 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=BLC','','VENTE;LCR;FAC') ; // Livr --> Fact
   // Génération automatique pièces
   30413 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','VENTE;PRE');  // Com --> Prep
   30415 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','VENTE;BLC');  // Com --> Livr
   30411 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=PRE','','VENTE;BLC') ; // Prep --> Livr
   30417 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=PRE','','VENTE;FAC') ; // Prep --> Fact
   30412 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLC','','VENTE;FAC') ; // Livr --> Fact
   30414 : AglLanceFiche ('AFF','AFPIECEPRO_MUL','','','TABLE:GP;NATURE:FPR;QUEPROVISOIRES');
   30416 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','VENTE;FAC');  // Com --> Fact
   30407 : AGLLanceFiche('AFF','AFPREPFACT_MUL','','','');   // Prépa facture par affaires
   30408 :AglLanceFiche ('AFF','AFPIECEPROANU_MUL','','','TABLE:GP;NATURE:FPR;QUEPROVISOIRES'); // Annulation facture provisoire
   // Affectation des commandes clients
   30441: AGLLanceFiche('MBO', 'AFFCDESELECT', '', '', ''); // sélection des lignes de commandes
   30442: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'RESERVATION'); // réservation
   30443: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'AFFECTATION'); // affectation
   30444: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'PREPARATION'); // préparation
   30445: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'FIN'); // fin de l'affectation
   30446: ;
   30447: AGLLanceFiche('MBO', 'AFFCDEMODIF', '', '', 'AFFECTATION'); // Modififcation des affectations
   30449: AGLLanceFiche('MBO', 'AFFCDEENTETE', '', '', ''); // paramètres d'affectation
   // PasRel
   30406 : AGLLanceFiche('GC','GCEXPORT_PASREL','GP_VENTEACHAT=VEN','','') ;  //Service Pas.Rel
{$IFNDEF CCS3}
{$IFNDEF MODE}
   // editions Affaires
   30801 : AFLanceFiche_ListeFact;
   30802 : AFLanceFiche_EDitPrepFact ;     // liste préparatoire factures
   30803 : AFLanceFiche_Portefeuille('');  // Portefeuille pièces
   30804 : CPLanceFiche_CPJALECR('NATUREJAL=ACH,VTE'); // journal de ventes comptable. appel fct Compta                                                       // Attention num utilisé AfterChangeModule
{$ENDIF}
{$ENDIF}


   // Tarifs HT
{$IFNDEF SANSPARAM}
   30101 : AGLLanceFiche('GC','GCASSISTANTTARIF','','','VEN') ; // Assistant création tarif article
   30102 : EntreeTarifArticle(taModif) ;
   30103 : EntreeTarifTiers(taModif) ;
   30104 : EntreeTarifCliArt(taModif) ;
   30105 : EntreeTarifCatArt(taModif) ;
   30106 : AGLLanceFiche('GC','GCTARIFCON_MUL','','','') ;
   30107 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','') ;
   30108 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','DUP') ; // duplication de tarif
//   30109 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','SUP') ; // suppression de tarif
   30109 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','SUP;VEN') ; // suppression de tarif
   30111 : GCLanceFiche_EditionTarif ('GC', 'GCEDITTARIFTIE', '', '', 'CLI');
   30112 : GCLanceFiche_EditionTarif ('GC', 'GCEDITTARIFART', '', '', 'CLI');
   // Tarifs TTC
   30902 : EntreeTarifArticle(taModif,TRUE);
   30903 : EntreeTarifTiers(taModif,TRUE);
   30904 : EntreeTarifCliArt(taModif,TRUE) ;
{$ENDIF}
   // Editions
   30601 : AGLLanceFiche('GC','GCEDITDOCDIFF_MUL','','','VEN') ;     // Editions différées
   30602 : AGLLanceFiche('GC','GCEDITTRAITEDIF','','','') ;     // Edition traites différées
   30603 : AGLLanceFiche('GC','GCPTFPIECE','','','VENTE') ;  // Portefeuille pièces
   30604 : AGLLanceFiche('GC','GCETATCOMMISSION','','','') ; // commissionnement commerciaux
   30605 : AGLLanceFiche('GC','GCETIQCLI','','','') ;        // Etiquettes clients
   // spécifs CEGID
   30609 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCEDITDOCDIFF_MP','','','VEN') ;     // Editions différées pour mise sous pli
   30610 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE') ;        // portefeuille articles
   30611 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID02','','','VENTE') ;        // portefeuille commercial
   30621 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;005') ;        // portefeuille articles
   30622 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;006') ;        // portefeuille articles
   30623 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;007') ;        // portefeuille articles
   30624 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;008') ;        // portefeuille articles
   30625 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;009') ;        // portefeuille articles
   30626 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;010') ;        // portefeuille articles
   30627 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;011') ;        // portefeuille articles
   30628 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;012') ;        // portefeuille articles
   30641 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;001') ;        // portefeuille articles
   30642 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;002') ;        // portefeuille articles
   30643 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;003') ;        // portefeuille articles
   30644 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPTFPIECECEGID01','','','VENTE;004') ;        // portefeuille articles
   // Suivi clients
   30701 : AGLLanceFiche('CP','EPBALAGEE','','','') ;     // Balance agée
   30702 : AGLLanceFiche('CP','EPGLAGE','','','') ;     // Grand livre
   30703 : AGLLanceFiche('CP','EPECHEANCIER','','','') ;     // Echéancier
   // fichiers
   30501 : if ctxMode in V_PGI.PGIContexte then AGLLanceFiche('GC','GCCLIMUL_MODE','T_NATUREAUXI=CLI','','')
           {$IFDEF GRC} else if ctxGRC in V_PGI.PGIContexte then RTLanceFiche_Prospect_Mul('RT','RTPROSPECT_MUL','T_NATUREAUXI=CLI','','GC') {$ENDIF}
           else AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=CLI','','') ;    // Clients
   30502 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=CLI','','') ; // contacts
   30503 : AGLLanceFiche ('GC','GCCOMMERCIAL_MUL','','','') ; // Commerciaux
   // GRC
   //30504 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','CLI') ;  // modification en série des tiers
   30504 : AGLLanceFiche('RT','RTQUALITE','','','CLI;GC') ;  // modification en série des tiers
{$IFNDEF SANSCOMPTA}
   //30505 : OuvreFermeCpte(fbAux,FALSE,'CLI') ;  // fermeture des comptes tiers
   //30506 : OuvreFermeCpte(fbAux,True,'CLI') ;   // ouverture compte tiers
   30507 : RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'CLI');


{$ENDIF}
   30506 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;CLI') ;  // ouverture compte tiers
   30505 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;CLI') ;  // fermeture compte tiers


{$IFDEF CAISSE}
///////  Ventes Comptoir  /////////////
   36100..36999 : MenuDispatchCaisse(Num);
{$ENDIF}




///////  Achats  /////////////
   // Saisie pièces
   31201 : CreerPiece('CF') ;
   31202 : CreerPiece('BLF') ;
   31203 : CreerPiece('FF') ;
   31234 : CreerPiece('BFA') ; // retour fournisseur positionné
   31231 : CreerPiece('AF') ; // Avoir financier
   31232 : CreerPiece('AFS') ; // Avoir sur stock
   31233 : CreerPiece('AFP') ; // Avoir valorisé
   31206 : CreerPiece('DEF') ;
   {$IFDEF GPAO} 
     211611 : CreerPiece('CSA') ;
   {$ENDIF GPAO }
   // Modifications
   31211 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=CF','','MODIFICATION')  ;  //commandes fournisseur
   31212 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=BLF','','MODIFICATION') ;  //bons de livraison fournisseur
   31213 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=DEF','','MODIFICATION') ;  //Proposition d'achat fournisseur
   31214 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=BFA','','MODIFICATION') ;  //Retour fournisseur
   {$IFDEF GPAO} 
     211612 : AGLLanceFiche('GC' , 'GCPIECEACH_MUL' , 'GP_NATUREPIECEG=CSA' ,'','MODIFICATION')  ;  //commandes sous-traitance d'achat
     211613 : AGLLanceFiche('GC' , 'GCPIECEACH_MUL' , 'GP_NATUREPIECEG=BSA' ,'','MODIFICATION')  ;  //RECEPTION sous-traitance d'achat
   {$ENDIF GPAO} 
   // Duplication unitaire pièces
   31221 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=CF','','CF') ;
   31222 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=BLF','','BLF') ;
   31223 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=FF','','FF') ;
   31224 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=DEF','','DEF') ;
   31225 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AF','','AF') ;
   31226 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AFS','','AFS') ;  // avoir fournisseur avec mise à jour du stock
   31227 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AFP','','AFP') ;  // avoir fournisseur avec mise à jour stock+prix
   31228 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=BFA','','BFA') ;  // retour fournisseur
   31204 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','ACH') ; //Histirisation des pièces
   // Consultation
   31101 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=FF','','CONSULTATION') ;
   31102 : AGLLanceFiche('GC','GCLIGNEACH_MUL','GL_NATUREPIECEG=FF','','CONSULTATION') ;
   // génération
   31501 : AGLLanceFiche('GC','GCPIECEVISA_MUL','GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH','','VENTEACHAT=ACH') ; // Visa des pièces
   31513 : AGLLanceFiche('GC','GCCONTREMVTOA','','','MODE=VTOA') ; // génération commande fournisseur de contremarque
   31511 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=CF','','BLF') ;
   {$IFDEF GPAO} 
     211623: AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=CSA','','BSA') ;
   {$ENDIF GPAO}
   31512 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=BLF','','FF') ;
   31514 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=DEF','','CF') ;
   // Génération manuelle pièces
   31531 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=CF','','ACHAT;CFR;BLF') ;  // Comm --> Livr
   31532 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=BLF','','ACHAT;LFR;FF') ; // Livr --> Fact
   {$IFDEF GPAO} 
     211621: AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=CSA','','ACHAT;CFR;BSA') ;
   {$ENDIF GPAO}
   // génération automatique
   31521 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CF','','ACHAT;BLF') ;
   31522 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLF','','ACHAT;FF') ;
   {$IFDEF GPAO} 
     211622: AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CSA','','ACHAT;BSA') ;
   {$ENDIF GPAO}
   //  avoirs
   31541 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=FF','','AFS') ; // Avoir sur stock
   31542 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=FF','','AFP') ; // Avoir valorisé

   // tarif fournisseurs
{$IFNDEF SANSPARAM}
   31401 : EntreeTarifFouArt (taModif) ;                       // saisie
   31402 : AGLLanceFiche('GC','GCTARIFFOUCON_MUL','','','') ;  // consultation
   31403 : AGLLanceFiche('GC','GCTARIFFOUMAJ_MUL','','','') ;  // mise à jour
   31404 : AGLLanceFiche('GC','GCPARFOU_MUL','','','') ;  // Import
   31405 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','SUP;ACH') ;  // Suppression des tarifs fournisseurs
{$ENDIF}
   // génération
   31301 : BEGIN Assist_Suggestion;  RetourForce:=TRUE;  END ;  // Suggestion de réappro
   31302 : AGLLanceFiche('GC','GCLANCEREA_MUL','','','') ;     // Lancement réappro
   // Editions
   31601 : AGLLanceFiche('GC','GCEDITDOCDIFF_MUL','','','ACH') ;     // Editions différées
   31602 : AGLLanceFiche('GC','GCPTFPIECE','','','ACHAT') ;  // Portefeuille pièces
   31612 : AGLLanceFiche('GC','GCART_NONREF','','','')  ;
   31611 : AGLLanceFiche('GC','GCFOURN_ARTICLE','','','') ;
   31613 : AGLLanceFiche('GC','GCARTICLE_FOURN','','','') ;
   // Fichiers
   31701 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','') ;
   31702 : AGLLanceFiche('GC','GCMULCATALOGUE','','','') ;
   31703 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=FOU','','') ;  // contacts
   // GRC
   //31704 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','FOU') ;  // modification en série des tiers
   31704 : AGLLanceFiche('RT','RTQUALITE','','','FOU;GC') ;  // modification en série des tiers
{$IFNDEF SANSCOMPTA}
//   31705 : OuvreFermeCpte(fbAux,FALSE,'FOU') ;  // fermeture des comptes tiers
   31705 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;FOU') ;
//   31706 : OuvreFermeCpte(fbAux,True,'FOU') ;   // ouverture compte tiers
   31706 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;FOU') ;
   31707 : RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'FOU');

{$ENDIF}
   {$IFDEF GPAO} 
     21163 : AGLLanceFiche('U','UCOMPFOURNI','','','ACTION=CREATION;DROIT=CDAMV');
     21164 : AGLLanceFiche('U','UORDRELASST','','','ACTION=MODIFICATION;LAS=ACH');
     211652 : AGLLanceFiche ('U','US0_QR1', '', '', 'ACTION=MODIFICATION');
     211653 : AGLLanceFiche ('U','US0_QR2', '', '', 'ACTION=MODIFICATION');
     211651 : AGLLanceFiche ('U','US0_QR3', '', '', 'ACTION=MODIFICATION');
     211661 : AglLanceFiche('U', 'UVIEWERST2','', '', '');
     211662 : AglLanceFiche('U', 'UVIEWERST1','', '', '');
     21167  : AglLanceFiche('U', 'UORDRELIG_MUL','', '', 'ACTION=MODIFICATION');
     21168  : AglLanceFiche('U', 'UORDREBES_MUL','', '', 'ACTION=CONSULTATION');
   {$ENDIF GPAO} 

///////  Articles et Stocks  /////////////////
   // Articles
   {$IFDEF GPAO}
     32501  : LanceMulArticle('','','WARTICLE;ACTION=MODIFICATION;ONLYPROFIL=-');
     325021 : AglLanceFiche('W','WPROFILART','','','');
     325022 : LanceMulArticle('', '', 'WARTICLE;ACTION=MODIFICATION;ONLYPROFIL=X');
   {$ELSE GPAO}
     32501 : AGLLanceFiche('GC','GCARTICLE_MUL','','','GCARTICLE') ; // article
     32502 : AglLanceFiche ('GC','GCPROFILART','','','') ;           // Profil article
   {$ENDIF GPAO}
   32503 : AGLLanceFiche('GC','GCARTICLE_MODIFLO','','','') ;  // modification en série
   32504 : AGLLanceFiche('GC','GCMULSUPPRART','','','') ;  // épuration articles
   32506 : AGLLanceFiche('GC','GCTARIFART_MUL','','','') ;  // Maj BAses Tarifaires
   32505 : EntreeTradArticle (taModif) ;  // traduction libellé articles
    // Consultation
   32110 : AGLLanceFiche('GC','GCDISPO_MUL','','','') ;
   32111 : if ctxMode in V_PGI.PGIContexte then
                AGLLanceFiche('MBO','MOUVSTK','','','')  // stock prévisionnel
           else AGLLanceFiche('GC','GCMOUVSTK','','','') ; // stock prévisionnel
   32112 : AGLLanceFiche('GC','GCMULCONSULTSTOCK','','','') ; // stock prévisionnel
   32113 : AGLLanceFiche('GC','GCCONTREAVT','','','') ; // suivi contremarque
   32114 : GCLanceFiche_MouvStkEx('GC','GCMOUVINTERNE_MUL','','','') ; // Pièces internes
   32115 : GCLanceFiche_ArtNonMvtes('GC','GCARTNONMVTES_MUL','','','') ; // Consult des articles non mouvementés
   // Traitement
   //32201 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;TEM'); // Transferts Inter-Dépôt  
   32201 : CreerTransfert('TEM', 'TEM_TRE'); // Transferts Inter-Dépôt  

   {$IFDEF STK}
     32221 : AGLLanceFiche('GC', 'GSMPHYSIQUE_FIC', '', '', 'MONOFICHE;ACTION=CREATION;ORIG=MNU;FLUX=STO');
     32222 : ;
   {$ELSE}
     32221 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;EEX');
     32222 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;SEX');
   {$ENDIF STK}
   32231 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;RCC'); // Réajustement réservé client
   32232 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;RCF'); // Réajustement réservé fournisseur
   32233 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;RPR'); // Réajustement préparé client
   32203 : AGLLanceFiche('GC','GCTTFINMOIS','','','');
   32204 : AGLLanceFiche('GC','GCTTFINEXERCICE','','','');
   32205 : AGLLanceFiche('GC','GCAFFECTSTOCK','','','');
   32206 : AGLLanceFiche('GC','GCINITSTOCK','','','');  // Initialisation des stocks
   32207 : GCLanceFiche_MouvStkExContr('GC','GCMOUVSTKEXCONTR','','','');  // Mouvements exceptionnels contremarque
{$IFDEF GPAO}
   { 2121: CBN }
      { 21211: Calcul }
      21211 : wAppelFicheCalculBesoinsNets;
      { 21212: Tableaux de bord }
         { 212121: TobViewer besoins }
         212121 : wAppelTobVWBesoinsNets;
         { 212122: TobViewer évolutions des stocks / global}
         212122 : wAppelTobVWEvolutionStocks('G');
         { 212123: TobViewer évolutions des stocks / Dépôts}
         212123 : wAppelTobVWEvolutionStocks('D');
         { 212124: TobViewer propositions }
         212124 : wAppelTobVWPropositions;
      { 21213: Validation des propositions}
         212131 : wAppelMulPropositionsCBN('PRA','CBN');
         212132 : wAppelMulPropositionsCBN('PRF','CBN');
         212133 : wAppelMulPropositionsCBN('PRS','CBN');
         212134 : wAppelMulPropositionsCBN('PRT','CBN');
      { 21214: Fichiers}
         21214 : AGLLanceFiche('W','WDEPOTAPP_FSL','','','ACTION=CREATION;DROIT=CDAMV');
{$ENDIF GPAO}

   // inventaire ////
   32310 : AGLLanceFiche('GC','GCLISTEINV','','','') ;
   32315 : AGLLanceFiche('GC','GCLISTEINVPRE','','','') ;
   32317 : GCLanceFiche_ListeInvVal('GC','GCLISTEINVVAL','','','') ;
   32320 : AGLLanceFiche('GC','GCSAISIEINV_MUL','','','') ;
   32330 : AGLLanceFiche('GC','GCVALIDINV_MUL','','','') ;
   32340 : GCLanceFiche_ListeInvMul('GC','GCLISTEINV_MUL','','','') ;

   32351 : AGLLanceFiche('MBO','TRANSTPI','','','') ;       // Transmission inventaire depuis un TPI (Mode 103111)
   32352 : AGLLanceFiche('MBO','TRANSINV_MUL','','','') ;   // Saisie / Consultation des inventaires transmis   (Mode 103112)
   32353 : AGLLanceFiche('GC','GCINTEGRINV_MUL','','','') ;  // Intégration des inventaires transmis  (Mode 103113)

   // inventaire contremarque ////
   32810 : GCLanceFiche_ListeInvContre('GC','GCSAISIEINV_MUL','','','CONTREMARQUE') ; // Consultation listes contremarque
   32811 : EntreeSaisieInvContrem ; // Saisie inventaire contremarque
   32812 : GCLanceFiche_ListeInvPreCon('GC','GCLISTEINVPRECON','','','') ; //Edition inventaire de contremarque non valorisé
   32813 : GCLanceFiche_ListeInvVal('GC','GCLISTEINVVALCON','','','') ; //Edition inventaire de contremarque valorisé
   32815 : GCLanceFiche_ListeInvMul('GC','GCLISTEINV_MUL','','','CONTREMARQUE') ; //Suppression liste inventaire de contremarque
   // Editions ////
   32401 : AGLLanceFiche('GC','GCARTDISPO','','','') ;
   32403 : AGLLanceFiche('GC','GCINVPERM','','','')  ;   // inventaire permanent
   32404 : AGLLanceFiche('GC','GCETIQART','','','')  ;
   {$IFNDEF EAGLCLIENT}
   32405 : if not AppelConsostock then Retourforce:=true;
   {$ENDIF}
   // Fichiers
   32601 : AGLLanceFiche('GC','GCDEPOT_MUL','','','') ;
   32602 : AGLLanceFiche('GC','GCEMPLACEMENT_MUL','','','') ;
   32603 : AGLLanceFiche('GC','GCPRIXREVIENT','','','') ;
   // Dimensions ////
   32701 : begin ParamTable('GCCATEGORIEDIM',taModif,110000045,PRien,3); AfterChangeModule(32); end;  // categorie
   32702 : AGLLanceFiche('GC','GCMASQUEDIM','','','') ;
{$IFNDEF SANSPARAM}
   32703 : ParamDimension ;
{$ENDIF}
   { 21220 : Logistique des transferts inter-dépôts }
   { 21221: Validation des propositions}
   212211 : wAppelMulPropositionsCBN('PRA','LOG');
   212212 : wAppelMulPropositionsCBN('PRT','LOG');
   { 21222: Scan}
   212221 : AGLLanceFiche('GC','PROPTRANSF_FSL','','','');
   212222 : AGLLanceFiche('GC','REACREECOLIS_FSL','','','');
   212223 : AGLLanceFiche('GC','PIECETRANSF_FSL','','','');
   { 21223: Saisie des transferts inter-dépôts}
   21223  : CreerTransfert('TEM', 'TEM_TRV_NUMDIFFERENT');
   { 21224: Validation des transferts inter-dépôts}
   21224  : AGLLanceFiche('GC','GCTRANSFERT_MUL','GP_NATUREPIECEG=TRV','','VALIDATION');
   { 21225: Editions}
   212251 : AGLLanceFiche ('GC','LOGISTIQUE0_QR1', '', '', 'ACTION=MODIFICATION');
   212252 : AGLLanceFiche ('GC','LOGISTIQUE1_QR1', '', '', 'ACTION=MODIFICATION');
   212253 : AGLLanceFiche ('GC','LOGISTIQUE2_QR1', '', '', 'ACTION=MODIFICATION');
   212254 : AGLLanceFiche ('GC','LOGISTIQUE3_QR1', '', '', 'ACTION=MODIFICATION');

   32711..32715 : ParamTable('GCGRILLEDIM'+IntToStr(Num-32710),taCreat,110000045,PRien) ;
   {$IFDEF STK}
   // Paramétrage //
   21291: AglLanceFiche('GC', 'GCSTKNATURE', '', '', '');
   {$ENDIF STK}
///////////// Analyses ///////////////////////
   // Ventes ////
   33101 : AGLLanceFiche('GC','GCCTRLMARGE','','','') ;
   33102 : AGLLanceFiche('GC','GCANALYSE_VENTES','','','CONSULTATION') ;
   33105 : AGLLanceFiche('GC','GCVENTES_CUBE','','','CONSULTATION') ;
   33106 : AGLLanceFiche('GC','GQR1SCC','','','') ;
   33111 : AGLLanceFiche('GC','GCGRAPHSTAT','','','VEN') ;
   33112 : AGLLanceFiche('GC','GCSTAT','','','VEN') ;
   33113 : AGLLanceFiche('GC','GCSTATCEGID','','','VEN') ;
   33114 : AGLLanceFiche('GC','GCSTATNOMEN','','','VEN') ;
   33107 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCNVXCLIENTS','','','') ; // Nouveaux Clients // AGLLanceFiche('GC','GQR1SCN','','','') ;
   33171 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;002') ; // Stats cegid par Marché
   33172 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;003') ; // Stats cegid par Marché
   33173 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;009') ; // Stats cegid par Marché
   33174 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;008') ; // Stats cegid par Marché
   33151 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;001') ; // Stats cegid par Agence
   33152 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;007') ; // Stats cegid par Agence
   33153 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;006') ; // Stats cegid par Agence
   33161 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;004') ; // Stats cegid par Canal
   33162 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;010') ; // Stats cegid par Canal
   33163 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCEGID01','','','VEN;005') ; // Stats cegid par Canal
   33121 : AGLLanceFiche('GC','GCGRAPHPALM','','','VEN') ;
   33122 : AGLLanceFiche('GC','GCPALMARES','','','VEN') ;
   33132 : AGLLanceFiche('AFF','AFVENTES_CUBE','','','CONSULTATION;PIECE') ;
   33131 : AGLLanceFiche('AFF','ETATPREVCA','','','');
{$IFDEF AFFAIRE}
{$IFNDEF MODE}
   33133 : AFLanceFiche_EcheanceCube('');
{$ENDIF}
{$ENDIF}
   33141 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GQR1GPA','','','') ;
   33142 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GQR1GII','','','') ;
   33143 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GQR1GCI','','','') ;
   33144 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GQR1GTV','','','') ;
   33181 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCSTATCOMM','','','') ;
   33182 : if VH_GC.GCIfDefCEGID then AGLLanceFiche('GC','GCPIECEECR_CUBE','','','') ;
   // Achats ////
   33201 : AGLLanceFiche('GC','GCCATANA','','','') ;
   33202 : AGLLanceFiche('GC','GCACHAT_CUBE','','','') ;
   33211 : AGLLanceFiche('GC','GCGRAPHSTAT','','','ACH') ;
   33212 : AGLLanceFiche('GC','GCSTAT','','','ACH') ;
   33214 : AGLLanceFiche('GC','GCSTATNOMEN','','','ACH') ;
   33221 : AGLLanceFiche('GC','GCGRAPHPALM','','','ACH') ;
   33222 : AGLLanceFiche('GC','GCPALMARES','','','ACH') ;
   // Tableaux de bord Affaires
   33401 : AGLLanceFiche('AFF','AFTBVIEWER','','','');     // Consult tableau de bord
   33402 : AGLLanceFiche('AFF','AFTBVIEWERMULTI','','','MULTI:X');     // Consult TB Multi affaires
   33403 : AGLLanceFiche('AFF','AFGRANDLIVRE','','','') ;  //grand livre ;
   33404 : AGLLanceFiche('AFF','AFBALANCE','','','') ;     //balance;
   33405 : AGLLanceFiche('AFF','AFTABLEAUBORD','','','') ; //AlimTableauBordAffaire;
   33406 : AGLLanceFiche('AFF','AFVENTES_CUBE','','','CONSULTATION;AFFAIRE') ;// cube affaire gm 09/04/02
   // Stocks
   33311 : AGLLanceFiche('GC','GCTRACABILITE','','','LOT') ;   // Traçabilité des lots
   33312 : AGLLanceFiche('GC','GCTRACABILITE','','','SERIE') ; // Traçabilité des numéros de série
{$IFNDEF CCS3}
   33321 : GCLanceFiche_VerifLigneSerie('GC','GCVERIFLIGNESERIE','','','CLI') ; // Vérification des ventes des numéros de série
   33322 : GCLanceFiche_VerifLigneSerie('GC','GCVERIFLIGNESERIE','','','FOU') ; // Vérification des achats des numéros de série
{$ENDIF}

   // Divers
{$IFNDEF EAGLCLIENT}
   33501 : Assist_DEBC ; // Declaration d'échange de biens (DEB)
{$ENDIF}

{$IFNDEF SANSPARAM}
   {Fonctions E-COMMERCE}
   {$IFNDEF CCS3}
   {$IFNDEF MODE}
   {$IFNDEF EAGLCLIENT}
   59101 : ExportAllTables;
   59102 : ImportAllTables;
   59103 : BalanceTradTablette;
   59201 : ImporteECommande;
   {$ENDIF}
   59202 : AGLLanceFiche('E','EIMPORTCMD','','','');
   {$IFNDEF EAGLCLIENT}
   59203 : AGLLanceFiche('E','ECMDVALID_MUL','','','');
   // import affaire

   59204 : AFLanceFiche_Mul_EActivite (''); // import E-Activité
   59205 : AFLanceFiche_Mul_ExportAff; // remplacé par 74654 en GI, à supprimer qd n'existe plus en GA export par affaires
   59206 : AFLanceFiche_MulEAffaire; // import E_affaires
   59207 : AGLLanceFiche('E','EIMPORTTIERS','','','');      // import E-Tiers
   59208 : AGLLanceFiche('E','EIMPORTCMD','','','');        // import E-Commandes
   59209 : AFLanceFiche_Suivi_EActivite; //suivi saisie décentralisée
   //
   59301 : AGLLanceFiche('E','EEXPORTARTICLE','','','');
   59302 : AGLLanceFiche('E','EEXPORTTIERS','','','');
   59303 : AGLLanceFiche('E','EEXPORTTARIF','','','');
   59304 : AGLLanceFiche('E','EEXPORTCOMM','','','');
   59305 : AGLLanceFiche('E','EEXPORTPAYS','','','');
   59401 : AGLLanceFiche('E','ETRADTABLETTE','','','');
   {$ENDIF} //CCS3
   {$ENDIF} //MODE
   {$ENDIF}
{$ENDIF}

////// Affaires
{$IFNDEF CCS3}
{$IFNDEF MODE}
   // Fichiers affaires
//   70101 : AglLanceFiche ('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=PRO','','STATUT=PRO;NOCHANGESTATUT');  // proposition d'affaires
   70101: AFLanceFiche_MulAffaire ('AFF_STATUTAFFAIRE=PRO','STATUT=PRO;NOCHANGESTATUT;PASMULTI');  // mission (avec stat cli)

   // 70102 remis pour le menu popup en attendant que le menu 27 soit mis a jour
   // en remplacant 70102 par 70122
   70102 : AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT');  // mission
   70122 : AFLanceFiche_MulAffaire('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT');  // mission
   70123 : AFLanceFiche_MulAffaire ('AFF_STATUTAFFAIRE=AFF','STATUT=AFF;NOCHANGESTATUT;PASMULTI');  // mission (avec stat cli)

   70111 : AglLanceFiche ('AFF','AFFAIREAVT_MUL','AFF_STATUTAFFAIRE=PRO','','STATUT=PRO;NOCHANGESTATUT');  // proposition de mission
   70112 : AglLanceFiche ('AFF','AFFAIREAVT_MUL','AFF_STATUTAFFAIRE=AFF','','STATUT=AFF;NOCHANGESTATUT');  // mission
   70103 : AglLanceFiche ('AFF','RESSOURCE_MUL','','','');                        // ressources
   70104 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=PRE','','PRE');    // Prestations
   70105 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=FRA','','FRA');    // Frais
   70107 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=CTR','','CTR');    // articles de contrat
   70106 : AGLLanceFiche ('AFF','HORAIRESTD','','','TYPE:STD');                   // Calendriers
   70108 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=POU','','POU');    // articles de cpourcentage

   // apporteur
   70130 : AFLanceFiche_Mul_Apporteur;

   // Editions Afaires
   70211 : AGLLanceFiche ('AFF','ETATSAFFAIRE','','','');   // Etats sur affaires
   70212 : AGLLanceFiche ('AFF','FICHEAFFAIRE','','','');
   70213 : AGLLanceFiche ('AFF','STATAFFAIRE','','','');
   70221 : AGLLanceFiche ('AFF','ETATSRESSOURCE','','',''); // Etats sur ressources
   70222 : AGLLanceFiche ('AFF','FICHERESSOURCE','','','');
   70223 : AGLLanceFiche ('AFF','STATRESSOURCE','','','');
   70224 : AGLLanceFiche ('AFF','ADRESSERESSOURCE','','','');

   // C.B 05/12/02
   // calendriers
   70231 : AFLanceFiche_Edit_Calendrier;
   70232 : AFLanceFiche_Edit_CalRegle;

   // clients
   70241 : AFLanceFiche_EtatsClient;
   70242 : AFLanceFiche_FicheClient;
   70243 : AFLanceFiche_EditAdresseClient;
   70244 : AFLanceFiche_EditStatClient;

   // affaires et propositions
   70251 : AFLanceFiche_EtatsAffaire;
   70252 : AFLanceFiche_FicheAffaire;
   70253 : AFLanceFiche_EditStatAffaire;

   //articles
   70261 : AFLanceFiche_EtatsPrestation; // Etats sur prestations
   70262 : AFLanceFiche_EditStatPrest;

   //Etiquette
   70270 : AFLanceFiche_EtiquetteClient;

   //suivi des propositions
   70280 : ;

   //etats libres
   70289 : LanceEtatLibreGIGA;

   70291 : AFLanceFiche_RefacturationRess;
   70292 : AFLanceFiche_RefacturationAffaire;

   // Documents clients
   70301 : AGLLanceFiche('AFF','AFPROPOSEDITMUL','AFF_STATUTAFFAIRE=PRO','','STATUT=PRO') ; // Edition des propositions
   70302 : AGLLanceFiche('AFF','AFPROPOSEDITMUL','AFF_STATUTAFFAIRE=AFF','','STATUT=AFF') ; // Edition des affaires
   // Saisie d'activité Par ressources
{$IFDEF AFFAIRE}
   70411 : AFCreerActivite ( tsaRess, tacTemps, 'REA' ) ; // Saisie d'activités Temps
   70412 : AFCreerActivite ( tsaRess, tacFrais, 'REA' ) ; // Saisie d'activités Frais
   70413 : AFCreerActivite ( tsaRess, tacFourn, 'REA' ) ; // Saisie d'activités Fournitures
   70414 : AFCreerActivite ( tsaRess, tacGlobal, 'REA' ) ; // Saisie d'activités Globale
   // Saisie d'activité Par affaires
   70421 : AFCreerActivite ( tsaClient, tacTemps, 'REA' ); // Saisie d'activités Temps
   70422 : AFCreerActivite ( tsaClient, tacFrais, 'REA' ); // Saisie d'activités Frais
   70423 : AFCreerActivite ( tsaClient, tacFourn, 'REA' ); // Saisie d'activités Fournitures
   70424 : AFCreerActivite ( tsaClient, tacGlobal, 'REA' ); // Saisie d'activités Globale
   70430 : AGLLanceFiche('AFF','AFACTIVITECON_MUL','','','');    // Modif ligne activité
   70440 : AFLanceFiche_ActiviteCube;
{$ENDIF}

   // Traitements affaires
   70501 : AglLanceFiche ('AFF','AFVALIDEPROPOS','AFF_STATUTAFFAIRE=PRO','','PRO'); // acceptation de missions
   70502 : AGLLanceFiche('AFF','AFAUGMEN_MUL','AFF_STATUTAFFAIRE=AFF','','') ;      // Augmentation affaire
   70503 : AGLLanceFiche('AFF','AFALIGNAFF','','','') ;                             // Alignement client sur affaire

   70511 : AFLanceFiche_EpurEchFAct('') ;                // epuration échéances facturées
   70512 : AFLanceFiche_EpurEchFAct('GENERATION');       // génération échéances sur tacite reconduction

   70504 : AGLLanceFiche('AFF','AFALIGN_MONNAIE','','','');
   70505 : AGLLanceFiche('AFF','AFAFFAIR_MODIFLOT','','','') ;

   70506 : AFLanceFiche_Modiflot_Affaire('') ;             // modif en série sur affaire
   70507 : AFLanceFiche_Modiflot_Affaire('TRA')  ;             // modif en série sur affaire avec traitement

   // Traitements sur activité
   70551 : AGLLanceFiche('AFF','AFACTIVITEGLO_MUL','','','AFA_REPRISEACTIV:TOU;');  // Modif ligne activité
   70552 : AGLLanceFiche('AFF','AFACTIVITEVISAMUL','','','');                       // Visas sur l'activité
   70553 : AFLanceFiche_RevalActiv;           // Revalorisation de ressources
   70554 : AFLanceFiche_RevalActivArticle;    // Revalorisation de l'activité par article
   70555 : BEGIN InitParpieceAffaire; CreerPiece('FRE'); END; //Reprise facturation externe
   70556 : AFLanceFiche_Mul_Piece('GP_NATUREPIECEG=FRE','TABLE:GP;NOCHANGE_NATUREPIECE;NATURE:FRE;ACTION=MODIFICATION;');  // COnsultation facturation externe

   // Génération CutOff totale
   70561 : AFLanceFiche_AfGenCutOffGlobale;  // Génération CutOff totale
{$IFNDEF EAGLCLIENT}
   70562 : AFLanceFiche_AfGenCutOffAFF_Mul ('')  ;  // Génération CutOff par affaire
{$endif}

   70563 : AFLanceFiche_ModifCutOffMul('TYPE:CVE;SAISIE:GLOBAL');    // Modification CutOff  cumulée
   70564 : AFLanceFiche_ModifCutOffMul('TYPE:CVE;SAISIE:DETAIL');    // Modification CutOff  détaillé
   70565 : AFLanceFiche_EtatAppreciation('ETAT:AFR');  // etat cut off par mois
   70566 : AFLanceFiche_EtatAppreciation('ETAT:ACU');  // etat cut off
{$IFDEF AFFAIRE}
   70570 :  AFLanceFiche_ActivGenerPaie; //Génération de l'activité vers la paie
   70571 :  AFLanceFiche_AnnulGenerPaie; //Annulation des lignes d'activité générées vers la paie
   70585 :  AFLanceFiche_Synchro_RessSalarie; //Affaire - Lien Paie - Synchro Ressource Salarié
{$ENDIF}
   // Editions sur activités
   70601 : AGLLanceFiche('AFF','AFJOURNALACT','','','') ;
   70602 : AGLLanceFiche('AFF','AFJOURNALCLIENT','','','') ;
   70611 : AGLLanceFiche('AFF','AFSYNTHRESS','','','') ;   // Synthèse assistant detaillée par affaire
   70612 : AGLLanceFiche('AFF','AFSYNTHRESSART','','','') ;// Synthèse assistant detaillée par article
   70621 : AGLLanceFiche('AFF','AFSYNTHCLIENT','','','') ; // Synthèse Client affaire détaillée par ressource
   70622 : AGLLanceFiche('AFF','AFSYNTHCLIART','','','') ; // Synthèse Client affaire détaillée par article

   70630 : AFLanceFiche_Synthese_F_NF; // synthese facturable / non facturable

   // Consultations
   70701 : AGLLanceFiche('AFF','AFINTERVENANTSAFF',' ','','');                      // intervenants sur affaires

   // Traitements ressources
   70582 : AFLanceFiche_ModifLog_Ress ('');             // Modif en série des ressources
   70583 : AFLanceFiche_ModifLog_Ress ('TRA');          // Modif en série des ressources  avec traitement
{$ENDIF}  //CCS3
{$ENDIF}  //MODE

    //===== ADMINISTRATION ================
   60201 :  AGLLanceFiche('YY','YYUSERGROUP','','','')  {FicheUserGrp} ;
   60202 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;   // Utilisateurs
   60208 :  GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','26;27;30;31;32;33;36;59;60;65;70;92;160;267;111;260'  );

      //// menus  /////////
   60501 : // Gestion des favoris
       if not VH_GC.GCIfDefCEGID then
          ParamFavoris('',lesTagsToRemove,False,False)
       else
          ParamFavoris('',lesTagsToRemoveCegid,False,False);
 {$IFNDEF EAGLCLIENT}
   {$IFNDEF SANSCOMPTA}
   60101 : GestionSociete(PRien,@InitSocietePGI,Nil) ;
   {$ENDIF}

   60102 : RecopieSoc(PRien,True) ;         // recopie société à société
   //////// Utilisateurs et accès
   3172  : AGLLanceFiche('YY','PROFILETABL','','','ETA') ;
   60203 : ReseauUtilisateurs(False) ;      // utilisateurs connectés
   60204 : VisuLog ;                        // Suivi d'activité
   60205 : ReseauUtilisateurs(True) ;       // RAZ connexions
   60206 : AGLLanceFiche('GC','GCPARAMOBLIG','','','') ;      // Champs obligatoires
   60207 : AGLLanceFiche('GC','GCPARAMCONFID','','','') ;       // restricyions fiches
    /////// Outils
{$IFDEF MODE}
   60301 : AGLLanceFiche('GC','GCJNALEVENT_MUL','','','') ; // Journal évènements
{$ELSE}
   60301 : AGLLanceFiche('YY','YYJNALEVENT','','','') ; // Journal évènements *
{$ENDIF}
   60401 : AGLLanceFiche('GC','GCCPTADIFF','','','') ;
      {$IFNDEF SANSPARAM}
   60402 : EntreeStockAjust;
   60403 : Entree_TraiteAdresse ;
      {$ENDIF}
   // epurations et historisations
   60411 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','VEN') ; //Historisation des pièces
   60412 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','ACH') ; //Historisation des pièces
   60421 : GCLanceFiche_PieceEpure('GC','GCPIECEEPURE_MUL','','','VEN') ; //Epuration des pièces
   60422 : GCLanceFiche_PieceEpure('GC','GCPIECEEPURE_MUL','','','ACH') ; //Epuration des pièces
   // utilitaires
   60602 : Assist_RazActivite;
   60603 : GCLanceFiche_UtilULst('GC','GCUTILULST','','','1;0');
   60605 : EntreeMajCompteurSouche;
   60606 : GCLanceFiche_UtilULst('GC','GCUTILULST','','','2;1');
   60608 : Assist_AffecteDepot;
   60609 : GCLanceFiche_UtilULst('GC','GCUTILULST','','','3;1');
   60610 : EntreeVerifAuxiTiers;
   60611 : EntreeVerifNatErrImp;

    /////// Utilitaires
    // PCS les imports sont dans GCIMPORT.exe
      {$IFNDEF SANSPARAM}
   // 60901 : BEGIN Assist_Import;  RetourForce:=TRUE;  END ;        // Import negoce
         {$IFDEF GRC}
   //60904 : BEGIN Assist_ImportProspect;  RetourForce:=TRUE;  END ;        // Import Prospect II
         {$ENDIF} {GRC}
      {$ENDIF} {SansParam}
///////// Paramètres /////////////
    /////// Paramètres - Société //////////
   65201 : BEGIN
           BrancheParamSocAffiche(StVire,StSeule) ;
           ParamSociete(False,StVire,StSeule,'',Nil,ChargePageSoc,SauvePageSoc,InterfaceSoc,110000220) ;
           END ;
   {$IFNDEF CCS3}
   {$IFNDEF MODE}
   65204 : BEGIN LanceAssistCodeAffaire; RetourForce := True;  END;
   {$ENDIF}
   {$ENDIF}
 {$ENDIF} {eAGL CLient}
   65203 : AGLLanceFiche('GC','GCETABLISS_MUL','','','') ;
    /////// Gestion Commerciale /////////
   {$IFDEF VRELEASE}
      {$IFNDEF EAGLCLIENT}
   65471 : if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
         AglLanceFiche ('GC','GCPARPIECE','','','') // LanceparPiece
         else AglLanceFiche ('GC','GCPARPIECEUSR','','','') ; // LanceparPiece  utilisateur
      {$ELSE}
   65471 :     AglLanceFiche ('GC','GCPARPIECEUSR','','','') ; // LanceparPiece  utilisateur
      {$ENDIF} // EAGLCLIENT
   {$ELSE}
   65471 : AglLanceFiche ('GC','GCPARPIECE','','','') ;// LanceparPiece ;
   {$ENDIF} // VRELEASE
   {$IFNDEF EAGLCLIENT}
   65472 : begin EditDocument('L','GPI','',True) ;RetourForce:=True ;end;
   65474 : begin EditEtat('E','GPJ','',True, nil, '', '');RetourForce:=True ;end;
    {$IFNDEF SANSPARAM} 65473 : EntreeListeSaisie('') ; {$ENDIF}
  {$ENDIF} // EAGLCLIENT
   65475 : ParamTable('GCEQUIPIECE',taCreat,0,PRien) ; // Natures de regroupement
   65401 : AGLLanceFiche('GC','GCPORT_MUL','','','') ;   // Port et frais
  {$IFNDEF SANSCOMPTA} 65402 : FicheSouche('GES') ; {$ENDIF}
   65403 : AGLLanceFiche('GC','GCUNITEMESURE','','','') ;
   65404 : AglLanceFiche ('GC','GCCODECPTA','','','') ;     //parametrage ventilations comptables
//   65408 : AGLLanceFiche ('AFF','VENTILANA','','','');      // Ventil ana par affaire
// gm 09/01/03
   {$IFNDEF CCS3}
   {$IFNDEF MODE}
   65408 : if GetParamSoc('SO_GCAxeAnalytique') then AGLLanceFiche('MBO','VENTILANA','','','') // ventil générique
           else
            AFLanceFiche_VentilAna; // ancienne ventilation GI/GA
   {$ENDIF}
   {$ENDIF}
(*  à revoir GM le 10/01/03
             begin
              if V_PGI.SAV then AFLanceFiche_VentilAna // ancienne ventilation GI/GA
              else  PgiInfo('Il faut utiliser la nouvelle analytique. Merci d''appeler l''assistance pour mise en place','');
              end;
*)
   {$IFNDEF EAGLCLIENT}
   {$IFNDEF SANSPARAM} 65405 : EntreeArrondi (taModif) ; {$ENDIF}
   {$ENDIF}
   65410 : ParamTable('YYDOMAINE',taCreat,110000217,PRien) ; // Domaines activité
   65409 : ParamTable('GCMOTIFMOUVEMENT',taCreat,110000284,PRien) ;          // Motif mouvement stock
   65406 : ParamTable('GCEMPLOIBLOB',taCreat,110000172,PRien) ;          // Emploi blob
   65451 : ParamTable ('GCCOMMENTAIREENT',taCreat,110000167,PRien,6) ;   // commentaires entete
   65453 : ParamTable ('GCCOMMENTAIREPIED',taCreat,110000167,PRien,6) ;  // commentaires pied
   65452 : begin  // commentaires ligne
           if (ctxMode in V_PGI.PGIContexte) or (ctxChr in V_PGI.PGIContexte) or (ctxFO in V_PGI.PGIContexte) then
              begin
              ParamTable ('GCCOMMENTAIRELIGNE',taCreat,110000167,PRien,6) ; // commentaires ligne
              end else
              begin
              if GetParamSoc('SO_GCCOMMENTAIRE') then AGLLanceFiche ('GC','GCCOMMENTAIRE','','','')
                                                 else ParamTable ('GCCOMMENTAIRELIGNE',taCreat,110000167,PRien,6) ;
              end;
           end;
   // Affaires
   65422 : AGLLanceFiche('AFF','AFBLOCAGEAFFAIRE','','','') ;      // blocage affaire
   65421 : AGLLanceFiche ('AFF','AFPROFILGENER','','','TYPE:DEF'); // profils de facturation par affaire
{$IFNDEF CCS3}
  {$IFNDEF MODE}
    65425 : AFLanceFiche_FormuleCutOff(''); // formule cut off
    65426 : AFLanceFiche_MulActivitePaie; // Paramétrage lien Activité avec Rubriques Paie
  {$ENDIF MODE}
{$ENDIF CCS3}
   /// Confident Affaire
{$IFNDEF EAGLCLIENT}
   65423 : AGLLanceFiche('YY','YYMULUSERCONF','','',''); // Confidentialité
{$ENDIF}   
   65424 : ParamTable('YYGROUPECONF',taCreat,0,PRien,3); // Groupes de conf
   // Paramètres Tiers
   65101 : ParamTable('GCZONECOM',taCreat,110000156,PRien) ;        // zones commerciales
   65102 : ParamTable('GCEXPEDITION',taCreat,110000156,PRien) ;     // mode d'expedition
   65103 : ParamTable('GCCOTEARTFOURN',taCreat,110000156,PRien) ;   // cote fournisseur
   65104 : ParamTable('TTTARIFCLIENT',taCreat,110000156,PRien) ;    // code tarif client
   65105 : ParamTable('GCCOMPTATIERS',taCreat,110000156,PRien) ;    // code comptable tiers
   65106 : ParamTable('TTSECTEUR',taCreat,110000156,PRien) ;        // secteurs d'activité
   65107 : ParamTable('GCORIGINETIERS',taCreat,110000156,PRien) ;   // Origine du tiers
   65108 : ParamTable('TTTARIFCLIENT',taCreat,0,PRien) ;            // Famille tarif client
   65109 : ParamTable('TTTARIFFOURNISSEUR',taCreat,0,PRien) ;       // Famille tarif fournisseur   
   // Paramètres articles
   65301 : ParamTable('GCCOLLECTION',taCreat,110000182,PRien,3,'Collections') ;     // collection
   65302 : ParamTable('GCTARIFARTICLE',taCreat,110000182,PRien) ;   // code tarif  article
   65303 : Begin              // libellés des codes familles
           ParamTable('GCLIBFAMILLE',taModif,110000182,PRien,3,'Titre des familles/sous-familles');
           AfterChangeModule(65); AfterChangeModule(105);
           end;
   65304 : ParamTable('GCCOMPTAARTICLE',taCreat,110000182,PRien) ;  // code comptable article
   // type et côte emplacement se trouvant dans le module établissement
   65305 : ParamTable('GCTYPEEMPLACEMENT',taCreat,110000182,PRien) ;    // type emplacement
   65306 : ParamTable('GCCOTEEMPLACEMENT',taCreat,110000182,PRien) ;    // côte emplacement
   65321..65323 : ParamTable('GCFAMILLENIV'+IntToStr(Num-65320),taCreat,110000182,PRien,3,RechDom('GCLIBFAMILLE','LF'+IntToStr(Num-65320),FALSE)) ;  // familles
   //// Paramètres généraux ///
   65901 : ParamTable('ttFormeJuridique',taCreat,110000187,PRien,3,'Abréviations postales') ;
   65903 : ParamTable('ttCivilite',taCreat,110000189,PRien) ;
   65902 : ParamTable('ttFonction',taCreat,110000188,PRien) ;
   65904 : AglLanceFiche ('YY','YYPAYS','','','') ;  // OuvrePays ;
   65905 : AglLanceFiche('GC','GCREGION_FSL','','','');
   65909 : ParamTable('YYSERVICE',taCreat,110000193,PRien) ;
   65910 : ParamTable('TTPROFILGESTION',taCreat,0,PRien) ;
   {$IFNDEF EAGLCLIENT}
     65906 : OuvrePostal(PRien) ;
   {$ENDIF EAGLCLIENT }
   {$IFNDEF SANSCOMPTA}
   65907 : FicheDevise('',tamodif,False) ;
   65908 : ParamTable('TTREGIMETVA',taCreat,110000183,PRien) ;
   65911 : ParamTvaTpf(true) ;
   65912 : ParamTvaTpf(false) ;
   {$ENDIF}
   65921 : FicheModePaie_AGL('');
   65922 : FicheRegle_AGL ('',False,taModif);
   65923 : AFLanceFiche_JourFerie;
   65924 : ParamTable('ttLangue',taCreat,110000190,PRien) ;
   659251 : ParamTable('GCMODEEXP' , taCreat, 110000184, PRien, 3, 'Modes d''expéditions'); { Export / Mode d'expédition }
   659252 : ParamTable('GCINCOTERM', taCreat, 110000184, PRien, 3, 'Incoterms');            { Export / Incoterm }
//   659253 : ParamTable('GCREGIME', taCreat, 0, PRien, 3, 'Régimes');                      { Export / Régime }
//   659254 : ParamTable('GCNATURE', taCreat, 0, PRien, 3, 'Natures');                      { Export / Nature }
   /////// Paramètres - Tables libres //////////
   65501..65509 : ParamTable('GCLIBREART'+IntToStr(Num-65500),taCreat,110000251,PRien,6,RechDom('GCZONELIBRE','AT'+IntToStr(Num-65500),FALSE)) ;
   65510 : ParamTable('GCLIBREARTA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','ATA',FALSE));
   65511..65519 : ParamTable('GCLIBRETIERS'+IntToStr(Num-65510),taCreat,110000252,PRien,6,RechDom('GCZONELIBRE','CT'+IntToStr(Num-65510),FALSE)) ;
   65520 : ParamTable('GCLIBRETIERSA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','CTA',FALSE));
   65251..65259 : ParamTable('AFTLIBREAFF'+IntToStr(Num-65250),taCreat,0,PRien,6,RechDom('GCZONELIBRE','MT'+IntToStr(Num-65250),FALSE)) ;
   65260 : ParamTable('AFTLIBREAFFA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','MTA',FALSE));
   65261..65269 : ParamTable('AFTLIBRERES'+IntToStr(Num-65260),taCreat,0,PRien,6,RechDom('GCZONELIBRE','RT'+IntToStr(Num-65260),FALSE));
   65270 : ParamTable('AFTLIBRERESA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','RTA',FALSE)) ;
   65640 : ParamTable ('AFTRESILAFF',taCreat,0,PRien);
   65541..65543 : ParamTable('GCLIBREPIECE'+IntToStr(Num-65540),taCreat,110000256,PRien,6) ;
   65545..65547 : ParamTable('GCLIBREFOU'+IntToStr(Num-65544),taCreat,110000253,PRien,6) ;
   65551..65559 : ParamTable('YYLIBREET'+IntToStr(Num-65550),taCreat,110000255,PRien,6,RechDom('GCZONELIBRE','ET'+IntToStr(Num-65550),FALSE)) ;
   65671..65679 : ParamTable('YYLIBRECON'+IntToStr(Num-65670),taCreat,110000254,PRien,6,RechDom('GCZONELIBRE','BT'+IntToStr(Num-65670),FALSE)) ; //tables libres contacts
   65680 : ParamTable('YYLIBRECONA',taCreat,110000254,PRien,6,RechDom('GCZONELIBRE','BTA',FALSE)) ;
   65560 : ParamTable('YYLIBREETA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','ETA',FALSE)) ;
   /////// Paramètres Affaires
   65601 : ParamTable('AFCOMPTAAFFAIRE',taCreat,110000214,PRien) ;
   65602 : ParamTable ('AFDEPARTEMENT',taCreat,110000214,PRien) ;
   65603 : ParamTable ('AFTLIENAFFTIERS',taCreat,110000214,PRien);
   65604 : ParamTable ('AFTRESILAFF',taCreat,110000214,PRien);
   65605 : ParamTable ('AFETATAFFAIRE',taCreat,110000214,PRien);
   65606 : ParamTable ('AFTREGROUPEFACT',taCreat,0,PRien);
    /////// Paramètres Ressources
    65651 : ParamTable ('AFTTYPERESSOURCE',taCreat,110000219,PRien);
    65652 : ParamTable('AFTTARIFRESSOURCE',taCreat,0,PRien) ;
    65653 : AglLanceFiche ('AFF','FONCTION','','','');
    65654 : AglLanceFiche ('AFF','COMPETENCE','','','');
    65655 : ParamTable ('AFTNIVEAUDIPLOME',taCreat,0,PRien);
    65656 : ParamTable ('AFTSTANDCALEN',taCreat,0,PRien);
    {Tables libres généralisées}
   65561..65565, 65571..65575, 65581..65590, 65592..65593, 65595..65597,
    65525..65530, 65535..65539 : FactoriseZL(Num,PRien) ;
// {$ENDIF} // EAGLCLIENT
{$IFDEF GRC}
   92949 : ShowTreeLinks(PRien,'Tablettes hiérarchiques','RT%',FMenuDisp.HDTDrawNode);
   92100..92948,92950..92999 : RTMenuDispatch (Num,PRien);
{$ENDIF}
{$IFDEF MODE}
   101000..112999 : begin
                    if (num >= 105921) and (num <=105929) then ParamTable('YYLIBREDEP'+IntToStr(Num-105920),taCreat,0,PRien,6,RechDom('GCZONELIBRE','DT'+IntToStr(Num-105920),FALSE)) ;
                    if (num = 105930) then ParamTable('YYLIBREDEPA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','DTA',FALSE)) ;
                    MODEMenuDispatch (Num,PRien,retourforce,sortiehalley);
                    if (num = 105133) then AfterChangeModule(105)
                    else if (num >= 105901) and (num <= 105905) then FactoriseZL(Num,PRien)
                    else if (num = 105909) then FactoriseZL(Num,PRien)
                    else if (num >= 105911) and (num <= 105915) then FactoriseZL(Num,PRien);
                    end;
   115000..115999 : MODEMenuDispatch (Num,PRien,retourforce,sortiehalley);
   59000..59999 : MODEMenuDispatch (Num,PRien,retourforce,sortiehalley);

{$ELSE}

    //////////////////////////////  Menus PCP Serveur
{$IFDEF NOMADESERVER}
    111100: AglToxSaisieParametres; // Paramètres par défaut
    111110: AglToxSaisieVariables; // Variables
    111120: AglToxMulSites; // Sites
    111130: AglToxSaisieGroupes; //  Groupes
    111140: AglToxMulConditions; // Requêtes
    111150: AglToxMulEvenements; // Evénements
    111200: // Démarrage/Arrêt des échanges
      begin
        //
        // 2 cas pour l'instant en fonction des paramètres sociétés
        //     A - On demande confirmation
        //     B - On intègre sans demander
        //
        if GetParamSoc('SO_GCTOXCONFIRM') = True then
        begin
          AglToxConf(aceConfigure, '', nil, GescomModeTraitementTox, nil, AvantArchivageTOX);
        end else
        begin
          AglToxConf(aceConfigure, '', GescomModeNotConfirmeTox, GescomModeTraitementTox, nil, AvantArchivageTOX);
        end;
      end;
    111210: AglToxMulChronos; // Consultation des échanges
    111220: AfficheCorbeille; // Consultation des corbeilles
    111230: AglLanceFiche('MBO', 'CONSOLIDATION', '', '', '');
    111240: LanceEtat('E','MST','IVE',True,False,False,Nil,'','',False);
    111300: ToxSimulation; // Intégration d'un fichier
    111310: ConsultManuelTox; // Visualisation d'un fichier
    111320: DetopeTox; // Détopage d'un fichier
    //111330: LanceGenereBE; // Génération des BE
    //111340: LanceMiniFTP; // Mini FTP pour récupérer des fichiers sur le central
{$ENDIF}



{$ENDIF}

{$IFDEF GPAO}
 	{ 98 : Configurateur }
    98001: AGLLanceFiche('X', 'XWIZARDS_MUL', '', '', '');
    98002: AGLLanceFiche('X', 'XDISPATCH_MUL', '', '', '');
    {$IFDEF AGL560}
      219002: ShowTreeLinks(PRien);
      219001: ShowAdminHDTLinks(PRien);
    {$ENDIF AGL560}
    98005: AGLLanceFiche('X', 'XMACROWIZARDS_MUL', '', '', '');


  105113 : ParamTable('GCTYPESTATPIECE',taCreat,0,PRien,3);
  105114 : AGLLanceFiche('GC','GCCODESTATPIECE','','','ACTION=MODIFICATION') ;

	{ 120: Production }
      { 1200: Gestion }
   		12001: AGLLanceFiche('W', 'WORDRETET_MUL'   , '', '', 'ACTION=MODIFICATION');	{ Entêtes d'ordres }
   		12002: AGLLanceFiche('W', 'WORDRELIG2_MUL'  , '', '', 'ACTION=MODIFICATION');	{ Lignes d'ordres }
   		12003: AGLLanceFiche('W', 'WORDREPHASE2_MUL', '', '', 'ACTION=MODIFICATION'); { Phases d'ordres }
   		12004: AGLLanceFiche('W', 'WORDREBES2_MUL'  , '', '', 'ACTION=MODIFICATION');	{ Besoins de production }
      12006: AGLLanceFiche('W', 'WORDRECMP_MUL'   , '', '', 'ACTION=MODIFICATION'); { Carnet de mise en production }
      { 1201: Consultation }
   		12011: AGLLanceFiche('W', 'WORDRETET_MUL'   , '', '', 'ACTION=' + iif(JAiLeDroitTag(12001), 'MODIFICATION', 'CONSULTATION'));	{ Entêtes d'ordres }
   		12012: AGLLanceFiche('W', 'WORDRELIG2_MUL'  , '', '', 'ACTION=' + iif(JAiLeDroitTag(12002), 'MODIFICATION', 'CONSULTATION'));	{ Lignes d'ordres }
   		12013: AGLLanceFiche('W', 'WORDREPHASE2_MUL', '', '', 'ACTION=' + iif(JAiLeDroitTag(12003), 'MODIFICATION', 'CONSULTATION')); { Phases d'ordres }
   		12014: AGLLanceFiche('W', 'WORDREBES2_MUL'  , '', '', 'ACTION=' + iif(JAiLeDroitTag(12004), 'MODIFICATION', 'CONSULTATION'));	{ Besoins de production }
      { 1202: Traitements }
         { 12020: Regroupement }
            120200: AGLLanceFiche('W', 'WOT_TRAITEMENT' , 'WOT_ETATTET=OUV'    , '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Fermeture }
            120201: AGLLanceFiche('W', 'WOT_TRAITEMENT' , 'WOT_ETATTET=OUV~FER', '', 'ACTION=MODIFICATION;TRAITEMENT=2'); { Solde }
         { 12021: Ordres de production }
            120210: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=DCL'    , '', 'ACTION=MODIFICATION;TRAITEMENT=1');  { Décomposition }
            120211: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=DEC'    , '', 'ACTION=MODIFICATION;TRAITEMENT=2');  { Lancement }
            120212: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=LAN~REC', '', 'ACTION=MODIFICATION;TRAITEMENT=3');  { Réception }
            120213: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=REC'    , '', 'ACTION=MODIFICATION;TRAITEMENT=4');  { Termine }
            120214: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=TER'    , '', 'ACTION=MODIFICATION;TRAITEMENT=5');  { Solde }
         { 12022: Phases }
            120220: AGLLanceFiche('W', 'WOP_TRAITEMENT', 'WOP_ETATPHASE=DEC~LAN', '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Lancement }
            120221: AGLLanceFiche('W', 'WOP_TRAITEMENT', 'WOP_ETATPHASE=LAN~REC', '', 'ACTION=MODIFICATION;TRAITEMENT=2'); { Réception }
            120222: AGLLanceFiche('W', 'WOP_TRAITEMENT', 'WOP_ETATPHASE=REC'    , '', 'ACTION=MODIFICATION;TRAITEMENT=3'); { Termine }
            120223: AGLLanceFiche('W', 'WOP_TRAITEMENT', 'WOP_ETATPHASE=TER'    , '', 'ACTION=MODIFICATION;TRAITEMENT=4'); { Solde }
         { 12023: Besoins }
            120230: AGLLanceFiche('W', 'WORDRELAS_MUL'   , '', '', 'ACTION=MODIFICATION');   { Listes à servir }
            120231: AGLLanceFiche('W', 'WOB_TRAITEMENT', 'WOB_ETATBES=BES~CON', '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Solde }
         { 12024: Ruptures }
            120240: AGLLanceFiche('W', 'WOB_RUPTURES'  , '', '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Consommation }
            120241: AGLLanceFiche('W', 'WOB_RUPTURES' , '' , '', 'ACTION=MODIFICATION;TRAITEMENT=2'); { Solde }
         { 12025: Gammes d'ordres }
            120250: AGLLanceFiche('W', 'WOG_TRAITEMENT'  , 'WOG_ETATOPE=LAN~COM', '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Termine}
            120251: AGLLanceFiche('W', 'WOG_TRAITEMENT' , 'WOG_ETATOPE=DCL~LAN~COM~TER' , '', 'ACTION=MODIFICATION;TRAITEMENT=2'); { Solde }
      { 1205: Analyse }
         12051: AGLLanceFiche('W', 'WORDRELIG_STA' , '' , '', '');                          { Analyse production }
         12052: wCallParamFaisabilite;                                                    { Paramètres calcul de faisabilité }
         12053: AGLLanceFiche('W', 'WTAUXSERVICE_MUL', '', '', 'ACTION=MODIFICATION');    { Paramètres taux de service }
         12054: AGLLanceFiche('W', 'WTAUXSERVICEGRAPH', '', '', 'ACTION=MODIFICATION');   { Graphique taux de service }
         12055: ; { Bilan d'of }

      { 1206: Editions }
         {12060 : Ordres de production}
         {12061 : Lignes de production}
            120610: wEditWOL;
       			120611: AGLLanceFiche ('W','WO5_QR1', '', '', 'ACTION=MODIFICATION');
            120612: AGLLanceFiche ('W','WO7_QR1', '', '', 'ACTION=MODIFICATION');
         {12062 : Phases de production}
            120620: wEditWOP;
     			120621: AGLLanceFiche ('W','WP0_QR1', '', '', 'ACTION=MODIFICATION');
         {12063 : Besoins de production}
            120630: AGLLanceFiche ('W','WO6_QR1', '', '', 'ACTION=MODIFICATION');
            120631: AGLLanceFiche ('W','WO4_QR1', '', '', 'ACTION=MODIFICATION');
      { 1207: Fichiers }
      { 1208: Champs libres }
         { 12080: Libellés regroupements/ordres }
            120800: begin
                      ParamTable('WLIBELLELIBREWOT', taModif, 0, Nil, 17); { Libellé tablettes WOT }
                      AfterChangeModule(120);
                    end;
            120801: ParamTable('WVALLIBREWOT' , taModif, 0, Nil, 17); { Libellé valeurs WOT }
            120802: ParamTable('WDATELIBREWOT', taModif, 0, Nil, 17); { Libellé dates WOT }
            120803: ParamTable('WBOOLLIBREWOT', taModif, 0, Nil, 17); { Libellé décisions WOT }
            120804: ParamTable('WCHARLIBREWOT', taModif, 0, Nil, 17); { Libellé textes WOT }
         { 12082: phases }
            120820: begin
                      ParamTable('WLIBELLELIBREWOP', taModif, 0, Nil, 17); { Libellé tablettes WOP }
                      AfterChangeModule(120);
                    end;
            120821: ParamTable('WVALLIBREWOP' , taModif, 0, Nil, 17); { Libellé valeurs WOP }
            120822: ParamTable('WDATELIBREWOP', taModif, 0, Nil, 17); { Libellé dates WOP }
            120823: ParamTable('WBOOLLIBREWOP', taModif, 0, Nil, 17); { Libellé décisions WOP }
            120824: ParamTable('WCHARLIBREWOP', taModif, 0, Nil, 17); { Libellé textes WOP }

      { 1209: Paramètres }
         { 12090: Automatismes }
            120900: AGLLanceFiche('W', 'WORDREAUTO_MUL', '', '', 'ACTION=MODIFICATION;WOA_CODEAUTO=WOL');
         { 12091: Numéros }
            120910: AGLLanceFiche('W', 'WJETONS_MUL', '', '', 'ACTION=MODIFICATION;WJT_PREFIXE=WOT'); { Regroupements }
            120911: AGLLanceFiche('W', 'WJETONS_MUL', '', '', 'ACTION=MODIFICATION;WJT_PREFIXE=WOL'); { Ordres de production }
            120912: AGLLanceFiche('W', 'WJETONS_MUL', '', '', 'ACTION=MODIFICATION;WJT_PREFIXE=WLS'); { Listes à servir }

	{ 121: Pilotage de flux }
   	{ 1211: Gestion }
        12111 : AGLLanceFiche('Q', 'QUFMITI'     , '', '', 'ACTION=CREATION');  { Itinéraires }
        12112 : AGLLanceFiche('Q', 'QUFMCIRCUIT' , '', '', 'ACTION=CREATION');  { Circuits }
      { 1212: Consultation }
        12121 : AGLLanceFiche('Q', 'QUFMITI'     , '', '', 'ACTION=' + iif(JAiLeDroitTag(12111), 'MODIFICATION', 'CONSULTATION'));  { Itinéraires }
        12122 : AGLLanceFiche('Q', 'QUFMCIRCUIT' , '', '', 'ACTION=' + iif(JAiLeDroitTag(12112), 'MODIFICATION', 'CONSULTATION'));  { Circuits }
      { 1218: Editions }
      { 1219: Fichiers }
        12191 : AGLLanceFiche('Q', 'QUFMPHASE'   , '', '', '');                              { Phases }
        12192 : AGLLanceFiche('Q', 'QUFLPOLE'    , '', '', '');                              { Pôle }
        12193 : AGLLanceFiche('Q', 'QUFLSITE'    , '', '', '');                              { Site / Atelier }
        12194 : AGLLanceFiche('Q', 'QUFMGROUPE'  , '', '', '');                              { Groupe }
        12195 : AGLLanceFiche('Q', 'QUFLTYPTRANS', '', '', '');                              { Cas de transport }
        12196 : AGLLanceFiche('Q', 'QUFMCALPOLE', '', '', '') ;                              { Calcul calendrier - Pôles }
        12197 : AGLLanceFiche('Q', 'QUFMGRPTYPE3', '', '', '');                              { Calcul calendrier - Groupe Gestion 3 [Jour calendaire] }
        12198 : AGLLanceFiche('Q', 'QUFMCALTYPTRANS', '', '', '');                           { Calcul calendrier - Cas de transport }

  { 125: SAV }
    { 1250: Gestion }
      12500: AglLanceFiche('W', 'WPARC_MUL'   , '' , '' , 'ACTION=MODIFICATION'); { Parc machines }
      12501: AglLanceFiche('W', 'WVERSION_MUL', '' , '' , 'ACTION=MODIFICATION'); { Versions }
    { 1251: Consultation }
      12510: AglLanceFiche('W', 'WPARC_MUL'   , '' , '' , 'ACTION=' + iif(JAiLeDroitTag(12500), 'MODIFICATION', 'CONSULTATION')); { Parc machines }
      12511: AglLanceFiche('W', 'WVERSION_MUL', '' , '' , 'ACTION=' + iif(JAiLeDroitTag(12501), 'MODIFICATION', 'CONSULTATION')); { Versions }
    { 1252: Traitements }
      {12520: Parc }
        125200: AGLLanceFiche('W', 'WPC_TRAITEMENT' , 'WPC_ETATPARC=ES'    , '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Mise Hors Service }
      {12521: Version }
        125210: AGLLanceFiche('W', 'WVS_TRAITEMENT' , 'WVS_ETATVERC=ACT'    , '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Péremption }
    { 1253: Analyse }
    { 1254: Editions }
    { 1255: Fichiers }
    { 1256: Champs libres }
        { 12560: Libellés fiches parc }
          125600: begin
                    ParamTable('WLIBELLELIBREWPC', taModif, 0, Nil, 17); { Libellés tablettes WPC }
                    AfterChangeModule(125);
                  end;
          125601: ParamTable('WVALLIBREWPC' , taModif, 0, Nil, 17); { Libellés valeurs WPC }
          125602: ParamTable('WDATELIBREWPC', taModif, 0, Nil, 17); { Libellés dates WPC }
          125603: ParamTable('WBOOLLIBREWPC', taModif, 0, Nil, 17); { Libellés décisions WPC }
          125604: ParamTable('WCHARLIBREWPC', taModif, 0, Nil, 17); { Libellés textes WPC }
        { 12562: Libellés fiches versions }
          125620: begin
                    ParamTable('WLIBELLELIBREWVS', taModif, 0, Nil, 17); { Libellés tablettes WVS }
                    AfterChangeModule(125);
                  end;
          125621: ParamTable('WVALLIBREWVS' , taModif, 0, Nil, 17); { Libellés valeurs WVS }
          125622: ParamTable('WDATELIBREWVS', taModif, 0, Nil, 17); { Libellés dates WVS }
          125623: ParamTable('WBOOLLIBREWVS', taModif, 0, Nil, 17); { Libellés décisions WVS }
          125624: ParamTable('WCHARLIBREWVS', taModif, 0, Nil, 17); { Libellés textes WVS }
      { 1257: Paramètres }

  { 126: Bases techniques }
	  { 1260: Gestion }
		  12600: LanceMulArticle('','','WARTICLE;ACTION=MODIFICATION');                                 { Articles }
   	  12601: AGLLanceFiche('W', 'WNOMETET_MUL', 'WNT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION');	{ Nomenclatures }
		  12602: AGLLanceFiche('W', 'WGAMMETET_MUL', 'WGT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION'); { Gammes }
		  12603: AGLLanceFiche('W', 'WRESSOURCEFAM_MUL', '', '', 'ACTION=MODIFICATION');                { Familles de ressources }
      { 1261: Consultation }
		   12610: LanceMulArticle('','','WARTICLE;ACTION=' + iif(JAiLeDroitTag(12600), 'MODIFICATION', 'CONSULTATION'));              { Articles }
   	   12611: AGLLanceFiche('W', 'WNOMETET_MUL', '', '', 'ACTION=' + iif(JAiLeDroitTag(12601), 'MODIFICATION', 'CONSULTATION'));	  { Nomenclatures }
		   12612: AGLLanceFiche('W', 'WGAMMETET_MUL', '', '', 'ACTION=' + iif(JAiLeDroitTag(12602), 'MODIFICATION', 'CONSULTATION'));  { Gammes }
       12613: AGLLanceFiche('W', 'WRESSOURCEFAM_MUL', '', '', 'ACTION=' + iif(JAiLeDroitTag(12603), 'MODIFICATION', 'CONSULTATION'));   { Familles de ressources }
       { 1262: Traitement }
         { 12620: Révision nomenclatures }
            126200: AGLLanceFiche('W', 'WREVISIONWNT_MUL', 'WNT_ETATREV=MOD'    , '', 'ACTION=MODIFICATION;REVISION=1');  { Validation }
            126201: AGLLanceFiche('W', 'WREVISIONWNT_MUL', 'WNT_ETATREV=MOD'    , '', 'ACTION=MODIFICATION;REVISION=2');  { Validation à date }
            126202: AGLLanceFiche('W', 'WREVISIONWNT_MUL', 'WNT_ETATREV=VAL'    , '', 'ACTION=MODIFICATION;REVISION=3');  { Révision mineure }
            126203: AGLLanceFiche('W', 'WREVISIONWNT_MUL', 'WNT_ETATREV=VAL'    , '', 'ACTION=MODIFICATION;REVISION=4');  { Révision majeure }
            126204: AGLLanceFiche('W', 'WREVISIONWNT_MUL', 'WNT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION;REVISION=5');  { Péremption }
            126205: AGLLanceFiche('W', 'WREVISIONWNT_MUL', 'WNT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION;REVISION=6');  { Péremption à date}
         { 12621: Révision gammes }
            126210: AGLLanceFiche('W', 'WREVISIONWGT_MUL', 'WGT_ETATREV=MOD'    , '', 'ACTION=MODIFICATION;REVISION=1');  { Validation }
            126211: AGLLanceFiche('W', 'WREVISIONWGT_MUL', 'WGT_ETATREV=MOD'    , '', 'ACTION=MODIFICATION;REVISION=2');  { Validation à date }
            126212: AGLLanceFiche('W', 'WREVISIONWGT_MUL', 'WGT_ETATREV=VAL'    , '', 'ACTION=MODIFICATION;REVISION=3');  { Révision mineure }
            126213: AGLLanceFiche('W', 'WREVISIONWGT_MUL', 'WGT_ETATREV=VAL'    , '', 'ACTION=MODIFICATION;REVISION=4');  { Révision majeure }
            126214: AGLLanceFiche('W', 'WREVISIONWGT_MUL', 'WGT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION;REVISION=5');  { Péremption }
            126215: AGLLanceFiche('W', 'WREVISIONWGT_MUL', 'WGT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION;REVISION=6');  { Péremption à date}

         12622: wCallParamChgComposant; { Changement de composant }
	   { 1265: Prix de revient }
		   12650 : wAppelCalculPrixDeRevient;
		   12651 : AGLLanceFiche ( 'W', 'WPDR_MUL'          , '', '', '' );	{ Consultation prix de revient }

		   { 12652 : Fichiers }
		      126520 : AGLLanceFiche ( 'W', 'WPDRTYPE_FIL'      , '', '', '' );	{ Type de Prix de Revient }
		      126521 : AGLLanceFiche ( 'W', 'WPDRTYPECOMPO_FIL' , '', '', '' );	{ Type de Composant }
		      126522 : AGLLanceFiche ( 'W', 'WPDRSECTION_FIL'   , '', '', '' );	{ Section  d'analyse }
		      126523 : AGLLanceFiche ( 'W', 'WPDRRUBRIQUE_FIL'  , '', '', '' );	{ Rubrique d'analyse }

   	{ 1266: Editions }
   		{ 12660: Nomenclatures }
   			126600 : AGLLanceFiche ('W','WN0_QR1', '', '', 'ACTION=MODIFICATION');
        126601 : AGLLanceFiche ('W','WN1_QR1', '', '', 'ACTION=MODIFICATION');
        126602 : AGLLanceFiche ('W','WN2_QR1', '', '', 'ACTION=MODIFICATION');
        126603 : AGLLanceFiche ('W','WN3_QR1', '', '', 'ACTION=MODIFICATION');
      { 12661: Gammes }
        126610 : AGLLanceFiche ('W','WG0_QR1', '', '', 'ACTION=MODIFICATION');
        126611 : AGLLanceFiche ('W','WG1_QR1', '', '', 'ACTION=MODIFICATION');
        126612 : AGLLanceFiche ('W','WG2_QR1', '', '', 'ACTION=MODIFICATION');
      { 12662: Ressources }
				126620 : AGLLanceFiche ('W','WR0_QR1', '', '', 'ACTION=MODIFICATION');
        126621 : AGLLanceFiche ('W','WR1_QR1', '', '', 'ACTION=MODIFICATION');
        126622 : AGLLanceFiche ('W','WR2_QR1', '', '', 'ACTION=MODIFICATION');
      { 12663: Nature de travail }
				126630 : AGLLanceFiche ('W','WT0_QR1', '', '', 'ACTION=MODIFICATION');
        126631 : LanceEtat ('E','WT0', 'NT1', True, False, False, Nil, '', '', False);
    { 1267: Fichiers }
      12670: AGLLanceFiche('W', 'WNATURETRAVAIL', '', '', '');                      { Nature de travail }
      12671: AGLLanceFiche('W', 'WGAOPER_FIC'   , '', '', 'ACTION=MODIFICATION');   { Catalogue des opérations }
      12672: AGLLanceFiche('W', 'WRESSOURCE_MUL', '', '', 'ACTION=MODIFICATION');   { Ressources }
    { 1268: Champs libres }
       { 12680: Libellés entêtes nomenclatures }
          126800: begin
                    ParamTable('WLIBELLELIBREWNT', taModif, 0, Nil, 17); { Libellé tablettes WNT }
                    AfterChangeModule(126);
                  end;
          126801: ParamTable('WVALLIBREWNT' , taModif, 0, Nil, 17); { Libellé valeurs WNT }
          126802: ParamTable('WDATELIBREWNT', taModif, 0, Nil, 17); { Libellé dates WNT }
          126803: ParamTable('WBOOLLIBREWNT', taModif, 0, Nil, 17); { Libellé décisions WNT }
          126804: ParamTable('WCHARLIBREWNT', taModif, 0, Nil, 17); { Libellé textes WNT }

       { 12682: Libellés lignes nomenclatures }
          126820: begin
                    ParamTable('WLIBELLELIBREWNL', taModif, 0, Nil, 17); { Libellé tablettes WNL }
                    AfterChangeModule(126);
                  end;
          126821: ParamTable('WVALLIBREWNL' , taModif, 0, Nil, 17); { Libellé valeurs WNL }
          126822: ParamTable('WDATELIBREWNL', taModif, 0, Nil, 17); { Libellé dates WNL }
          126823: ParamTable('WBOOLLIBREWNL', taModif, 0, Nil, 17); { Libellé décisions WNL }
          126824: ParamTable('WCHARLIBREWNL', taModif, 0, Nil, 17); { Libellé textes WNL }

       { 12684: Libellés entêtes gammes }
          126840: begin
                    ParamTable('WLIBELLELIBREWGT', taModif, 0, Nil, 17); { Libellé tablettes WGT }
                    AfterChangeModule(126);
                  end;
          126841: ParamTable('WVALLIBREWGT'    , taModif, 0, Nil, 17); { Libellé valeurs WGT }
          126842: ParamTable('WDATELIBREWGT'   , taModif, 0, Nil, 17); { Libellé dates WGT }
          126843: ParamTable('WBOOLLIBREWGT'   , taModif, 0, Nil, 17); { Libellé décisions WGT }
          126844: ParamTable('WCHARLIBREWGT'   , taModif, 0, Nil, 17); { Libellé textes WGT }

       { 12686: Libellés lignes gammes }
          126860: begin
                    ParamTable('WLIBELLELIBREWGL', taModif, 0, Nil, 17); { Libellé tablettes WGL }
                    AfterChangeModule(126);
                  end;
          126861: ParamTable('WVALLIBREWGL'    , taModif, 0, Nil, 17); { Libellé valeurs WGL }
          126862: ParamTable('WDATELIBREWGL'   , taModif, 0, Nil, 17); { Libellé dates WGL }
          126863: ParamTable('WBOOLLIBREWGL'   , taModif, 0, Nil, 17); { Libellé décisions WGL }
          126864: ParamTable('WCHARLIBREWGL'   , taModif, 0, Nil, 17); { Libellé textes WGL }

    { 1269: Paramètres }
        12691: AGLLanceFiche('W', 'WNATURETRAVAIL', '', '', '');      						{ Nature de travail }


  	{ 21520 : Gestion de production : 21530 }
      215201 :
              begin
                {$IFNDEF EAGLCLIENT}
                if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
                  AglLanceFiche ('GC', 'GCPARPIECE', '', '', 'PARAMPIECESGP=OUI') // LanceparPiece
                else
                  AglLanceFiche ('GC', 'GCPARPIECEUSR', '', '', 'PARAMPIECESGP=OUI'); // LanceparPiece  utilisateur
                {$ELSE EAGLCLIENT}
                  AglLanceFiche ('GC', 'GCPARPIECEUSR', '', '', 'PARAMPIECESGP=OUI');
                {$ENDIF EAGLCLIENT}
              end;
      215202 : AGLLanceFiche('W', 'WINITCHAMPTET_MUL', '', '', 'ACTION=MODIFICATION'); { Initilisation de champs }
    { 21530 : Conversion d'unités }
      215301 : AGLLanceFiche ('W','WUNITES_MUL','','','') ;	   		    { Unité de mesures - Fiche GPAO }
      215302 : AGLLanceFiche ('W', 'WCARA_FIC', '', '', '' );	  		  { Caractéristiques }
      215303 : ParamTable('WCODEFORME', taCreat, 0, PRien, 3); 	  		{ Code formes }
      215304 : AGLLanceFiche ('W', 'WFORME_VIE', '', '', '' );			  { Formes et caractéristiques }
      215305 : AGLLanceFiche ('W', 'WCUQUALIFIANT_FIC', '', '', '' ); { Qualifiants }
      215306 : AGLLanceFiche ('W', 'WFORMULECHAMP_VIE', '', '', '' ); { Champs mis à jour par les formules }
      215307 : AGLLanceFiche ('W', 'WFORMULE_MUL', '', '', '' );		  { Formules }
      215308 : AGLLanceFiche ('W', 'WCUARRONDI_FIC', '', '', '' );	  { Arrondis }
      215309 : ParamTable('WCONTEXTE', taCreat, 0, PRien, 3); 	  		{ Contextes de conversion }

   	{ 2153: Tablettes Articles }
		  215401 : AGLLanceFiche ( 'W', 'WCOLORIS_FSL'   , '', '', '' );	{ Coloris Articles }

      { Données des tables libres production }
            120810..120819: ParamTable('WLIBREWOT' + IntToHex(Num - 120809, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWOT', 'TWOT_LIBREWOT' + intToHex(Num - 120809, 1), false)); { Données des tables libre WOT }
            120830..120839: ParamTable('WLIBREWOP' + IntToHex(Num - 120829, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWOP', 'TWOP_LIBREWOP' + intToHex(Num - 120829, 1), false)); { Données des tables libre WOP }

      { Données des tables libres bases techniques }
            126810..126819: ParamTable('WLIBREWNT' + IntToHex(Num - 126809, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWNT', 'TWNT_LIBREWNT' + intToHex(Num - 126809, 1), false)); { Données des tables libre WNT }
            126830..126839: ParamTable('WLIBREWNL' + IntToHex(Num - 126829, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWNL', 'TWNL_LIBREWNL' + intToHex(Num - 126829, 1), false)); { Données des tables libre WNL }
            126850..126859: ParamTable('WLIBREWGT' + IntToHex(Num - 126849, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWGT', 'TWGT_LIBREWGT' + intToHex(Num - 126849, 1), false)); { Données des tables libre WGT }
            126870..126879: ParamTable('WLIBREWGL' + IntToHex(Num - 126869, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWGL', 'TWGL_LIBREWGL' + intToHex(Num - 126869, 1), false)); { Données des tables libre WGL }

      { Données des tables libres SAV }
            125610..125619: ParamTable('WLIBREWPC' + IntToHex(Num - 125609, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWPC', 'TWPC_LIBREWPC' + intToHex(Num - 125609, 1), false)); { Données des tables libre WPC }
            125630..125639: ParamTable('WLIBREWPC' + IntToHex(Num - 125629, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWVS', 'TWVS_LIBREWVS' + intToHex(Num - 125629, 1), false)); { Données des tables libre WVS }

  { GPAO : Administration : 219 }
    { Traitement : 21904 }
      219045: wResetWJT;    { Procédure de re-calcul de identifiants }

{$ENDIF GPAO}

  {$IFDEF EDI}
    { 260: EDI }
      { 26010: Gestion }
        260110: AGLLanceFiche('GC', 'EDIMESSAGES_MUL', '', '', 'ACTION=MODIFICATION');{ Messages EDI }
        260120: AGLLanceFiche('GC', 'EDITRANSFERT_MUL', 'ETR_TYPETRANS=EMI', '', 'ACTION=MODIFICATION');{ Transferts EDI }
        260130: AGLLanceFiche('GC', 'EDIRECEPTION_MUL', 'ERE_ETATRECEPTION=NON', '', 'ACTION=MODIFICATION');{ Réceptions EDI }
        260140: AGLLanceFiche('GC', 'EDIPIECE_MUL', '', '', 'ACTION=MODIFICATION');{ Pièces EDI }
        260150: AGLLanceFiche('GC', 'EDILIGNE_MUL', '', '', 'ACTION=MODIFICATION');{ Lignes EDI }
        260160: AGLLanceFiche('GC', 'EDITIERS_MUL', '', '', 'ACTION=MODIFICATION');{ Tiers EDI }
        260170: AGLLanceFiche('GC', 'EDITESTSCRIPT_FIC', '', '', 'ACTION=MODIFICATION');{ Tiers EDI }
      { 26020: Consultation }
        260210: AGLLanceFiche('GC', 'EDIMESSAGES_MUL', '', '', 'ACTION=CONSULTATION');{ Messages EDI }
        260220: AGLLanceFiche('GC', 'EDITRANSFERT_MUL', 'ETR_TYPETRANS=EMI', '', 'ACTION=CONSULTATION');{ Transferts EDI }
        260230: AGLLanceFiche('GC', 'EDIRECEPTION_MUL', 'ERE_ETATRECEPTION=NON', '', 'ACTION=CONSULTATION');{ Réceptions EDI }
        260240: AGLLanceFiche('GC', 'EDIPIECE_MUL', '', '', 'ACTION=CONSULTATION');{ Pièces EDI }
        260250: AGLLanceFiche('GC', 'EDILIGNE_MUL', '', '', 'ACTION=CONSULTATION');{ Lignes EDI }
        260260: AGLLanceFiche('GC', 'EDILIGNE_MUL', '', '', 'ACTION=CONSULTATION');{ Tiers EDI }
      { 26030: Analyse }
        260310: AGLLanceFiche('GC', 'EDITRANSHISTO_STA', '', '', 'ACTION=MODIFICATION');{ Historique des transferts EDI }
      { 26040: Traitements }
        260410: EDICallFicPlanif;{ Planificateur de tâches EDI }
      { 26050: Edition }
{$ENDIF EDI}


  { 210: Ventes }
    { 2104 : Tarifs Ventes }
         210401 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=201','','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE=201;APPEL=MENU');
         210402 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=201','','ACTION=CONSULTATION;DROIT=....V;MONOFICHE;YTA_FONCTIONNALITE=201;APPEL=MENU');
         210403 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=201','','ACTION=MODIFICATION;DROIT=...MV;MONOFICHE;YTA_FONCTIONNALITE=201;APPEL=MENU');
         210404 : AppelTarifsRecherche('201', '1', 'RECHERCHE'); // Recherche de Tarif celon un contexte
    { 2105 : Indirect  Ventes }
    { 2106 : Commissionnement Ventes }
         210601 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=202','','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE=202;APPEL=MENU');
         210602 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=202','','ACTION=CONSULTATION;DROIT=....V;MONOFICHE;YTA_FONCTIONNALITE=202;APPEL=MENU');
         210603 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=202','','ACTION=MODIFICATION;DROIT=...MV;MONOFICHE;YTA_FONCTIONNALITE=202;APPEL=MENU');
         210604 : AppelTarifsRecherche('202', '2', 'RECHERCHE');

  { 211: Achats }
    { 2114 : Tarifs Fournisseurs }
      { 21141 : Tarifs Achats }
           211411 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=101','','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE=101;APPEL=MENU');
           211412 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=101','','ACTION=CONSULTATION;DROIT=....V;MONOFICHE;YTA_FONCTIONNALITE=101;APPEL=MENU');
           211413 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=101','','ACTION=MODIFICATION;DROIT=...MV;MONOFICHE;YTA_FONCTIONNALITE=101;APPEL=MENU');
           211414 : AppelTarifsRecherche('101', '1', 'RECHERCHE');
      { 21142 : Tarifs Sous-traitance d'achat }
           211421 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=301','','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE=301;APPEL=MENU');
           211422 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=301','','ACTION=CONSULTATION;DROIT=....V;MONOFICHE;YTA_FONCTIONNALITE=301;APPEL=MENU');
           211423 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=301','','ACTION=MODIFICATION;DROIT=...MV;MONOFICHE;YTA_FONCTIONNALITE=301;APPEL=MENU');
           211424 : AppelTarifsRecherche('301', '1', 'RECHERCHE');
      { 21143 : Tarifs Sous-traitance de phase }
           211431 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=401','','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE=401;APPEL=MENU');
           211432 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=401','','ACTION=CONSULTATION;DROIT=....V;MONOFICHE;YTA_FONCTIONNALITE=401;APPEL=MENU');
           211433 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=401','','ACTION=MODIFICATION;DROIT=...MV;MONOFICHE;YTA_FONCTIONNALITE=401;APPEL=MENU');
           211434 : AppelTarifsRecherche('401', '1', 'RECHERCHE');
    { 2115 : Indirects Achats }                                                                                                                                    
//GPAO_TARIFS à intégrer une fois validé par Lyon  211501 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=103','','ACTION=CREATION    ;DROIT=CDAMV;MONOFICHE;YTA_FONCTIONNALITE=103;APPEL=MENU');
//GPAO_TARIFS à intégrer une fois validé par Lyon  211502 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=103','','ACTION=CONSULTATION;DROIT=....V;MONOFICHE;YTA_FONCTIONNALITE=103;APPEL=MENU');
//GPAO_TARIFS à intégrer une fois validé par Lyon  211503 : AGLLanceFiche('Y','YTARIFSAQUI_MUL','YTA_FONCTIONNALITE=103','','ACTION=MODIFICATION;DROIT=...MV;MONOFICHE;YTA_FONCTIONNALITE=103;APPEL=MENU');
//GPAO_TARIFS à intégrer une fois validé par Lyon  211504 : AppelTarifsRecherche('103', '1', 'RECHERCHE');

	{ GPAO : Paramètres : 215 }
   	{ 21510 : Gestion Commercial }
      215101 : AGLLanceFiche ( 'Y', 'YTARIFSPECIAL_FSL'   , '', '', '');  { Conditions spéciales } 

{$IFDEF GPAOLIGHT}
  { 267: Production }
    { 267000: Gestion }
      267001: AGLLanceFiche('W', 'WORDRETET_MUL'   , '', '', 'ACTION=MODIFICATION');	{ Entêtes d'ordres }
      267002: AGLLanceFiche('W', 'WORDRELIG2_MUL'  , '', '', 'ACTION=MODIFICATION');	{ Lignes d'ordres }
      267003: AGLLanceFiche('W', 'WORDREBES2_MUL'  , '', '', 'ACTION=MODIFICATION');	{ Besoins de production }
      267004: AGLLanceFiche('W', 'WORDRECMP_MUL'   , '', '', 'ACTION=MODIFICATION'); { Carnet de mise en production }
    { 267100: Consultation }
      267101: AGLLanceFiche('W', 'WORDRETET_MUL'   , '', '', 'ACTION=CONSULTATION');	{ Entêtes d'ordres }
      267102: AGLLanceFiche('W', 'WORDRELIG2_MUL'  , '', '', 'ACTION=CONSULTATION');	{ Lignes d'ordres }
      267103: AGLLanceFiche('W', 'WORDREBES2_MUL'  , '', '', 'ACTION=CONSULTATION');	{ Besoins de production }
    { 267200: Traitements }
       { 267210: Regroupement }
          267211: AGLLanceFiche('W', 'WOT_TRAITEMENT' , 'WOT_ETATTET=OUV'    , '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Fermeture }
          267212: AGLLanceFiche('W', 'WOT_TRAITEMENT' , 'WOT_ETATTET=OUV~FER', '', 'ACTION=MODIFICATION;TRAITEMENT=2'); { Solde }
       { 267220: Ordres de production }
          267221: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=DCL'    , '', 'ACTION=MODIFICATION;TRAITEMENT=1');  { Décomposition }
          267222: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=DEC'    , '', 'ACTION=MODIFICATION;TRAITEMENT=2');  { Lancement }
          267223: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=LAN~REC', '', 'ACTION=MODIFICATION;TRAITEMENT=3');  { Réception }
          267224: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=REC'    , '', 'ACTION=MODIFICATION;TRAITEMENT=4');  { Termine }
          267225: AGLLanceFiche('W', 'WOL_TRAITEMENT', 'WOL_ETATLIG=TER'    , '', 'ACTION=MODIFICATION;TRAITEMENT=5');  { Solde }
       { 267230: Besoins }
          267231: AGLLanceFiche('W', 'WORDRELAS_MUL'   , '', '', 'ACTION=MODIFICATION');   { Listes à servir }
          267232: AGLLanceFiche('W', 'WOB_TRAITEMENT', 'WOB_ETATBES=BES~CON', '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Solde }
       { 267240: Ruptures }
          267241: AGLLanceFiche('W', 'WOB_RUPTURES'  , '', '', 'ACTION=MODIFICATION;TRAITEMENT=1'); { Consommation }
          267242: AGLLanceFiche('W', 'WOB_RUPTURES' , '' , '', 'ACTION=MODIFICATION;TRAITEMENT=2'); { Solde }
    { 267300: Editions }
       {26730 : Ordres de production}
       {267310 : Lignes de production}
          // 267311: wEditWOL;
          267312: AGLLanceFiche ('W','WO5_QR1', '', '', 'ACTION=MODIFICATION');
          267313: AGLLanceFiche ('W','WO7_QR1', '', '', 'ACTION=MODIFICATION');
       {267320 : Besoins de production}
          267321: AGLLanceFiche ('W','WO6_QR1', '', '', 'ACTION=MODIFICATION');
          267322: AGLLanceFiche ('W','WO4_QR1', '', '', 'ACTION=MODIFICATION');
       {267330 : Nomenclatures }
          267331 : AGLLanceFiche ('W','WN0_QR1', '', '', 'ACTION=MODIFICATION');
          267332 : AGLLanceFiche ('W','WN1_QR1', '', '', 'ACTION=MODIFICATION');
          267333 : AGLLanceFiche ('W','WN2_QR1', '', '', 'ACTION=MODIFICATION');
          267334 : AGLLanceFiche ('W','WN3_QR1', '', '', 'ACTION=MODIFICATION');
    {267400: Fichiers }
          267401: LanceMulArticle('GA_TYPEARTICLE="NOM"','','WARTICLE;ACTION=MODIFICATION');                               { Articles }
          267402: AGLLanceFiche('W', 'WNOMETET_MUL', 'WNT_ETATREV=MOD~VAL', '', 'ACTION=MODIFICATION');	{ Nomenclatures }
    {2678: Champs libres }
       {267810: Libellés regroupements/ordres }
          267811: begin
                    ParamTable('WLIBELLELIBREWOT', taModif, 0, Nil, 17); { Libellé tablettes WOT }
                    AfterChangeModule(267);
                  end;
          267812: ParamTable('WVALLIBREWOT' , taModif, 0, Nil, 17); { Libellé valeurs WOT }
          267813: ParamTable('WDATELIBREWOT', taModif, 0, Nil, 17); { Libellé dates WOT }
          267814: ParamTable('WBOOLLIBREWOT', taModif, 0, Nil, 17); { Libellé décisions WOT }
          267815: ParamTable('WCHARLIBREWOT', taModif, 0, Nil, 17); { Libellé textes WOT }
       {267820: Tables libres regroupements }
          267821..267830: ParamTable('WLIBREWOT' + IntToHex(Num - 267820, 1), taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWOT', 'TWOT_LIBREWOT' + intToHex(Num - 267820, 1), false)); { Données des tables libre WOT }

    { 2679: Paramètres }
       { 267900: Automatismes }
          267911: AGLLanceFiche('W', 'WORDREAUTO_MUL', '', '', 'ACTION=MODIFICATION;WOA_CODEAUTO=WOL');
       { 267910: Numéros }
          267921: AGLLanceFiche('W', 'WJETONS_MUL', '', '', 'ACTION=MODIFICATION;WJT_PREFIXE=WOT'); { Regroupements }
          267922: AGLLanceFiche('W', 'WJETONS_MUL', '', '', 'ACTION=MODIFICATION;WJT_PREFIXE=WOL'); { Ordres de production }
          267923: AGLLanceFiche('W', 'WJETONS_MUL', '', '', 'ACTION=MODIFICATION;WJT_PREFIXE=WLS'); { Listes à servir }
{$ENDIF}
      else if not TraiteMenuSpecif(Num) then HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
var ii : integer ;
    Arg5LanceFiche : string ;

    InfosTable : MyArrayValue;
BEGIN
  InfosTable := nil;
if ((Num=8) and (TT='GCTIERSSAISIE'))  then
    BEGIN
    ii:=TTToNum(TT) ;
    if (ii>0) and (pos('T_NATUREAUXI="FOU"',V_PGI.DECombos[ii].Where)>0) then num:=12  ;
    END ;
Case Num of
{$IFNDEF GCGC}
    1 : {Compte gene} FicheGene(Nil,'',LeQuel,Action,0) ;
    2 : {Tiers compta} FicheTiers(Nil,'',LeQuel,Action,1) ;
    4 : {Journal} FicheJournal(Nil,'',Lequel,Action,0) ;
{$ENDIF}
    5 : BEGIN {Affaires}
        Arg5LanceFiche:= ActionToString(Action);
        if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
        if (lequel ='') then lequel:='AFF';
        AGLLanceFiche('AFF','AFFAIRE','',Lequel,Arg5LanceFiche);
        END;
    6 : BEGIN {Ressources}
        Arg5LanceFiche:= ActionToString(Action);
        if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
        AGLLanceFiche('AFF','RESSOURCE','',Lequel,Arg5LanceFiche);
        END;
    7 : {Article} DispatchTTArticle ( Action ,Lequel,TT,Range) ;
    8 : {Clients par T_TIERS}
        BEGIN
        Lequel:=TiersAuxiliaire(Lequel,false);
        Arg5LanceFiche:=ActionToString(Action)+';MONOFICHE' ;
        If Range='' then Arg5LanceFiche:=Arg5LanceFiche+ ';T_NATUREAUXI=CLI' else Arg5LanceFiche:=Arg5LanceFiche+Range ;
        AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche) ;
        END ;
    9 : {Commerciaux} AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)+';GCL_TYPECOMMERCIAL="REP"') ;
   10 : {Conditionnement} AGLLanceFiche('GC','GCCONDITIONNEMENT',Range,Lequel,ActionToString(Action)) ;
   //11 : {Catalogue} AGLLAnceFiche('GC','GCCATALOGU_NVFOUR','',Lequel,ActionToString(Action)) ;
   11 : {Catalogue} AGLLAnceFiche('GC','GCCATALOGU_SAISI3','',Lequel,ActionToString(Action)) ;
   12 : {Fournisseurs via T_TIERS} begin
            Lequel:=TiersAuxiliaire(Lequel,false);
            AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=FOU') ;
            end;
   13 : {Dépôts}
        if (CtxMode in V_PGI.PGIContexte) then
             AGLLanceFiche('MBO','DEPOT','',Lequel,ActionToString(Action)+';MONOFICHE')
        else AGLLanceFiche('GC','GCDEPOT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   14 : {Apporteurs} AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)+';GCL_TYPECOMMERCIAL="APP"') ;
   15 : {Règlement}  FicheRegle_AGL ('',False,taModif) ;
   16 : {contacts}
    begin    
        if TT = 'YYCONTACTTIERS' then
        begin
        InfosTable := wGetSqlFieldsValues(['T_NATUREAUXI', 'T_PARTICULIER', 'T_LIBELLE'],'TIERS','T_AUXILIAIRE="' + Range + '"');
        if (Lequel <> '') and (InfosTable <> nil) then Lequel := 'T;'+ Range + ';' + Lequel;
        AGLLanceFiche('YY','YYCONTACT', 'C_TYPECONTACT=T;C_AUXILIAIRE=' + Range, Lequel, ActionToString(Action)+';MONOFICHE'
                      + ';TYPE=T;TYPE2=' + InfosTable[0]  + ';TIERS=' + TiersAuxiliaire(Range, True)
                      + ';PART=' + InfosTable[1]
                      + ';TITRE='+ InfosTable[2] +';ALLCONTACT' ) ;
        end
        else AGLLanceFiche('YY','YYCONTACT',Lequel,Range,ActionToString(Action)+';MONOFICHE') ;
    end;
   17 : {Fonction des ressources} AGLLanceFiche('AFF','FONCTION','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   18 : {Etablissements}
        if (CtxMode in V_PGI.PGIContexte) then
          AGLLanceFiche('MBO','ETABLISS','',Lequel,ActionToString(Action)+';MONOFICHE')
        else
          AGLLanceFiche('GC','GCETABLISS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   19 : {Contenu tables libres} AGLLanceFiche('GC','GCCODESTATPIECE','','','ACTION=CREATION;YX_CODE='+Range) ;
   20 :  ;
   21 : {Ports} AGLLanceFiche('GC','GCPORT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
{$IFDEF GRC}
   22..24 : RTMenuDispatchTT ( Num, Action, Lequel, TT, Range );
{$ENDIF}
   25 : {Démarques} AGLLanceFiche('MFO','DEMARQUE','',Lequel,'') ;
   28 : {Clients par T_AUXILIAIRE}
        BEGIN
        Arg5LanceFiche:=ActionToString(Action)+';MONOFICHE' ;
        If Range='' then Arg5LanceFiche:=Arg5LanceFiche+ ';T_NATUREAUXI=CLI'  else Arg5LanceFiche:=Arg5LanceFiche+';'+Range ;
        AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche) ;
        END ;
   29 : {Fournisseurs via T_AUXILIAIRE}
        begin
        Arg5LanceFiche:=ActionToString(Action)+';MONOFICHE' ;
        If Range='' then Arg5LanceFiche:=Arg5LanceFiche+ ';T_NATUREAUXI=FOU'  else Arg5LanceFiche:=Arg5LanceFiche+';'+Range ;
        AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,Arg5LanceFiche) ;
        end;
{$IFDEF GRC}
   30..40 : RTMenuDispatchTT ( Num, Action, Lequel, TT, Range );
{$ENDIF}
   41 : begin
        if Action=taConsult then Action := taModif;
        AGLLanceFiche('MBO','EXPORT',Range,Range+';'+Lequel,ActionToString(Action)+';MONOFICHE');
        end;
   42 : begin
          InfosTable := wGetSqlFieldsValues(['ADR_TYPEADRESSE', 'ADR_NUMEROADRESSE'],'ADRESSES',Range + 'and ADR_NADRESSE = ' + Lequel);
          if InfosTable <> nil then Lequel := InfosTable[0] + ';' + IntToStr(InfosTable[1]);
          AGLLanceFiche('GC','GCADRESSES',WMakeRangeFromFplus(Range) ,Lequel, ActionToString(Action)+';MONOFICHE' + ';' + WMakeRangeFromFplus(Range) + ';NATUREAUXI=CLI') ;

        end;
   126521 : AGLLanceFiche ('W', 'WPDRTYPECOMPO_FIL' , '', Lequel, ActionToString(Action)+';MONOFICHE');	{ Type de composant }
   126522 : AGLLanceFiche ('W', 'WPDRSECTION_FIL'   , '', Lequel, ActionToString(Action)+';MONOFICHE');	{ Section  d'analyse }
   126523 : AGLLanceFiche ('W', 'WPDRRUBRIQUE_FIL'   , '', Lequel, ActionToString(Action)+';MONOFICHE');	{ Rubrique d'analyse }
   12671  : AGLLanceFiche ('W', 'WGAOPER_FIC'       ,'', Lequel, ActionToString(Action)+';MONOFICHE');   { Catalogue des opérations }
   { HYPER-ZOOMS : De 600 à 699 => Réservé ZoomEdtGPAO => Voir avec TS }
   600 : AGLLanceFiche('W', 'WNOMETET_FIC'  , '', Lequel, ActionToString(Action)+';MONOFICHE'); { En-tête de nomenclature }
   601 : AGLLanceFiche('W', 'WNOMELIG_FIC'  , '', Lequel, ActionToString(Action)+';MONOFICHE'); { Lignes de nomenclature }
   602 : AGLLanceFiche('W', 'WGAMMETET_FIC' , '', Lequel, ActionToString(Action)+';MONOFICHE'); { En-tête de gamme }
   603 : AGLLanceFiche('W', 'WGAMMELIG_FIC' , '', Lequel, ActionToString(Action)+';MONOFICHE'); { Lignes de gamme }
   604 : AGLLanceFiche('W', 'WARTNAT_FIC'   , Lequel, '', ActionToString(Action)+';MONOFICHE'); { Nature / Article }
   605 : AGLLanceFiche('W', 'WNATURETRAVAIL', Lequel, '', ActionToString(Action)+';MONOFICHE'); { Natures de travail }


   // *** maj des tablettes ***
   900 : ParamTable('GCCOMMENTAIRELIGNE',taCreat,0,Nil,6);  // mis de manière explicite car appellé depuis un Grid

   // pour mise à jour par defaut des tablettes
   994 : ParamTable(TT,taModif,0,Nil,9) ;   {choixext}
   995 : ParamTable(TT,taCreat,0,Nil,9) ;   {choixext}
   996 : ParamTable(TT,taModif,0,Nil,6) ;   {choixext}
   997 : ParamTable(TT,taCreat,0,Nil,6) ;   {choixext}
   998 : ParamTable(TT,taModif,0,Nil,3) ;   {choixcode ou commun}
   999 : ParamTable(TT,taCreat,0,Nil,3) ;   {choixcode ou commun}

   { 126000: Tablettes Bases Techniques }
   126001..126010: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWNT', 'TWNT_LIBREWNT' + IntToHex(Num - 126000, 1), false)); { Tables libres WNT }
   126011..126020: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWNL', 'TWNL_LIBREWNL' + IntToHex(Num - 126010, 1), false)); { Tables libres WNL }
   126101..126110: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWGT', 'TWGT_LIBREWGT' + IntToHex(Num - 126100, 1), false)); { Tables libres WGT }
   126111..126120: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWGL', 'TWGL_LIBREWGL' + IntToHex(Num - 126110, 1), false)); { Tables libres WGL }
   126121:  begin {Familles de ressources}
              Arg5LanceFiche:= ActionToString(Action);
              if (Range<>'') then
                Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
              AGLLanceFiche('W','WRESSOURCEFAM_FIC','',Lequel,Arg5LanceFiche + ';MONOFICHE;ESTFAMILLE=OUI');
            end;

   { 120000: Tablettes Production }
   120001..120010: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWOT', 'TWOT_LIBREWOT' + IntToHex(Num - 120000, 1), false)); { Tables libres WOT }
   120011..120020: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWOP', 'TWOP_LIBREWOP' + IntToHex(Num - 120010, 1), false)); { Tables libres WOP }

   { 125000: Tablettes SAV }
   125001..125010: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWPC', 'TWPC_LIBREWPC' + IntToHex(Num - 125000, 1), false)); { Tables libres WPC }
   125011..125020: ParamTable(TT, taCreat, 0, Nil, 17, RechDom('WLIBELLELIBREWVS', 'TWVS_LIBREWVS' + IntToHex(Num - 125010, 1), false)); { Tables libres WVS }

   215401 : AGLLanceFiche ( 'W', 'WCOLORIS_FSL'   , Lequel, '', '' );	{ Coloris Articles }

{$IFDEF MODE}
   112200 : MODEMenuDispatchTT ( Num, Action, Lequel, TT, Range );
{$ENDIF}
{$IFDEF CAISSE}
   112200 : CAISSEMenuDispatchTT ( Num, Action, Lequel, TT, Range );
{$ENDIF}
  end ;
END ;

procedure TFMenuDisp.HDTDrawNode (Sender: TObject; var Caption: string) ;
begin
{$IFDEF GRC}
  RTModifTabHie (Sender,Caption);
{$ENDIF}
end;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
AjusteMenu(NumModule);
TripoteMenuGCS3(NumModule) ;
TraiteMenusMode(NumModule) ;
TraiteMenusModeEAGL(NumModule) ;
{$IFDEF GRC}
RTMenuTraiteMenusGRC(NumModule);
{$ENDIF}
TraiteMenusGescom(NumModule) ;
TraiteMenusAffaire(NumModule) ;
TraiteSpecifsCEGID(NumModule) ;
TraiteMenusEAGL(NumModule) ;
TraiteMenusTarifs;
{$IFDEF GPAO} 
  TraiteMenusGPAO(NumModule);
{$ENDIF GPAO}

{$IFDEF EAGLCLIENT}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(5,FMenuG.DispatchX) ;
  92 : ChargeMenuPop(8,FMenuG.DispatchX) ;
   else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  END ;
{$ELSE}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(5,FMenuG.DispatchX) ;
  92 : ChargeMenuPop(8,FMenuG.DispatchX) ;
  101,102,103,105,106,110,111 : ChargeMenuPop(24,FMenuG.DispatchX) ;
   else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  END ;
{$ENDIF}

{$IFDEF GPAO}
{ Correspondance Menu GPAO - autres modules }
  Case NumModule of
    210: AfterChangeModule(30); { Ventes }
    211: AfterChangeModule(31); { Achats }
    212: AfterChangeModule(32); { Articles Zet stocks }
    213: AfterChangeModule(33); { Analyses }
    219: AfterChangeModule(60); { Administration }
    215: AfterChangeModule(65); { Paramètres }
  end;
{$ENDIF GPAO}

{$IFDEF CCS3}
FMenuG.RemovePopupItem(92107) ;
FMenuG.RemovePopupItem(70102);
{$ENDIF CCS3}
{$IFDEF GRC}
if (GetParamSoc('SO_RTAFFICHEINFOS')=False) then
     FMenuG.RemovePopupItem(92296) ;
{$ENDIF GRC}
END ;

procedure InitApplication ;
BEGIN
Application.Title := TitreHalley;
{$IFDEF GPAODEV} 
  www(Nil);
  V_PGI.SAV := True;
{$ENDIF}
 ProcZoomEdt:=GCZoomEdtEtat ;
{$IFDEF GRC}
  RTMenuInitApplication;
//  ProcCalcMul:=RTProcCalcMul ;
{$ENDIF GRC}
ProcCalcMul := PGIProcCalcMul;
ProcCalcEdt:=GCCalcOLEEtat ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChargeMag:=ChargeMagHalley ;

{$IFDEF EAGLCLIENT}
  FMenuG.OnMajAvant:=Nil ;
  FMenuG.OnMajApres:=Nil ;
{$ELSE EAGLCLIENT}
  FMenuG.OnMajAvant:=Nil ;
  FMenuG.OnMajApres:=Nil ;
  FMenuG.OnMajPendant:=Nil ;
  FMenuG.OnChangeUser:=OnChangeUserPGI ;
{$ENDIF EAGLCLIENT}
FMenuG.OnAfterProtec:=AfterProtecGC ;
{if VH_GC.GCIfDefCEGID then
     AfterProtecGC('')  // Pas de seria CEGID
else }
FMenuG.SetSeria(GCCodeDomaine,GCCodesSeria,GCTitresSeria) ;
FMenuG.OnChangeModule:=AfterChangeModule ;
V_PGI.DispatchTT:=DispatchTT ;
{$IFNDEF MODE}
  FMenuG.SetModules([60],[49]) ;
  V_PGI.NbColModuleButtons:=1 ; V_PGI.NbRowModuleButtons:=6 ;
{$ENDIF MODE}
{$IFNDEF EAGLCLIENT}
  FMenuG.SetPreferences(['Pièces'],False) ;
{$ENDIF}
FMenuG.OnChargeFavoris := ChargeFavoris;

{$IFDEF GPAO} { Mis à la fin pour pour écraser les autres : Je ne comprends pas comment cela est géré en tant qu'intégré... }
//	ProcCalcMul := wProcCalcMul;
  ProcZoomEdt := GPAOZoomEdtEtat;
  RegisterAdminMetaSet;   
{$ENDIF GPAO}

{$IFDEF TESTSIC}
  {$IFDEF EAGLCLIENT}
    SaveSynRegKey('eAGLHost', 'inf-eagl:80', true);
    FCegidIE.HostN.Enabled := false;
  {$ENDIF EAGLCLIENT}
{$ENDIF TESTSIC}

END ;

initialization

RegisterGCPGI ;
{$IFNDEF EAGLCLIENT}
{$IFDEF GRC}
RegisterGRCPGI ;
{$ENDIF}
{$ENDIF}
RegisterPGIOLEAUTO ;

end.
