unit MDispPOP;

interface

uses Forms, sysutils;

procedure InitApplication;
procedure AfterChangeModule(NumModule: integer);

type
  TFMenuDisp = class(TDatamodule)
  end;

var FMenuDisp: TFMenuDisp;

implementation

uses Windows, HMsgBox, Classes, Hctrls, Controls, HPanel, UIUtil,
  ImgList, ParamSoc,
  {$IFDEF EAGLCLIENT}
  Maineagl, MenuOLX, M3FP, UtileAGL, EntGC, eTablette, AGLInit,
  CalcOleGescom, UtilGC, Facture, TarifArticle, TarifTiers, TarifCliArt, TarifCatArt,
  VerifLigneSerie_TOF,
  {$IFDEF AFFAIRE}
  AFActivite, ActiviteUtil,
  {$ENDIF}
  {$IFDEF GRC}
  RTMenuGRC, CalcMulProspect,
  {$ENDIF}
  {$IFDEF GESCOM}
  MenuCaisse,
  {$ENDIF}
  {$IFDEF MODE}MODEMenu, {$ENDIF}
  Ent1, AssistSuggestion, Tradarticle, TiersUtil, FichComm, ZoomEdtGescom, GCOleAuto,
  HistorisePiece_Tof, ListeInvContrem_TOF, InvContrem_TOF, UTOFGCLISTEINV_MUL,
  ListeInvPreCon_TOF, UTOFListeInvVal, MouvStkExContr_TOF, UTofMouvStkEx,
  {$ELSE}
  FE_Main, MenuOLG,
  AGLInit, tablette, pays, region, codepost, UtomArticle,
  EdtQR, CreerSoc, CopySoc, Reseau, Mullog, LicUtil, FactUtil, EdtEtat,
  {$IFNDEF V530}EdtREtat, UEdtComp, {$ENDIF}
  {$IFNDEF SANSCOMPTA}devise, {$ENDIF}
  AssistInitSoc, Ent1, Transfert, saisutil,
  UTomPiece, EntGC, UtilGC,
  FichComm, ZoomEdtGescom, CalcOLEGescom,
  UtofConsoEtat, UTofGcTraiteAdresse,
  UtilSoc, AssistSuggestion, TiersUtil, AssistCodeAffaire,
  AssistDEBC, EdtDOC, {$IFNDEF V530}EdtRdoc, {$ENDIF}
  GCOleAuto,
  {$IFNDEF SANSCOMPTA}
  Souche, Suprauxi, Tva, OuvrFerm, InitSoc,
  {$ENDIF}
  {$IFNDEF SANSPARAM}
  dimension, UtomModePaie, Arrondi, ListeDeSaisie, tradarticle,
  TarifArticle, TarifCliArt, TarifTiers, TarifCatArt, eTransferts,
  AssistStockAjust, VerifLigneSerie_TOF,
  {$ENDIF}
  {$IFNDEF GCGC}
  MulGene, General, Tiers, Journal, TVA, OuvrFerm,
  EtbMce, EtbUser, TeleTrans, banquecp,
  {$ENDIF}
  {$IFDEF AFFAIRE}
  AFActivite, ActiviteUtil, ConfidentAffaire,
  {$ENDIF}
  {$IFDEF MODE}MODEMenu, {$ENDIF}
  {$IFDEF GRC}
  RegGRCPGI, RTMenuGRC, CalcMulProspect,
  {$ENDIF}
  {$IFDEF GESCOM}
  MenuCaisse,
  {$ENDIF}
  Facture, HistorisePiece_Tof, PieceEpure_Tof, ListeInvContrem_TOF, InvContrem_TOF,
  UTOFGCLISTEINV_MUL, ListeInvPreCon_TOF, UTOFListeInvVal, MouvStkExContr_TOF,
  UTofMouvStkEx,
  {$ENDIF}
  {$IFDEF AGL550B}entPGI, {$ENDIF}
  {$IFNDEF V530}GrpFavoris, {$ENDIF}
  {$IFDEF NOMADE}
  UtilPop, FOExportCais, {AssistImport,}
  uToxClasses, uToxConst, uToxFiches, GCTOXWORK, ToxSim, ToxDetop, uToxConf, ToxConsultBis, GCToxCtrl,
  {$ENDIF}
  StatDim_Tof, // DCA - Ajout menu stocks dans PCP
  ArtPrestation_Tof,
  UtilDispGC, UtilArticle, HEnt1,
  UtomUtilisat, UtilPGI;

{$R *.DFM}

function ChargeFavoris: boolean;
var lesTagsToRemove: string;
begin
  {$IFNDEF V530}
  lesTagsToRemove := '';
  AddGroupFavoris(FMenuG, lesTagsToRemove);
  result := true;
  {$ENDIF}
end;

procedure FactoriseZL(Tag: integer; PRien: THPanel);
var St: string;
  iHelpContext: Longint;
begin
  St := '';
  case Tag of
    65561: St := 'AT';
    65562: St := 'AM';
    65563: St := 'AD';
    65564: St := 'AC';
    65565: St := 'AB';
    65571: St := 'CT';
    65572: St := 'CM';
    65573: St := 'CD';
    65574: St := 'CC';
    65575: St := 'CB';
    65581: St := 'ET';
    65582: St := 'EM';
    65583: St := 'ED';
    65584: St := 'EC';
    65585: St := 'EB';
    65592: St := 'VM';
    65593: St := 'VD';
    65595: St := 'FT';
    65596: St := 'FM';
    65597: St := 'FD';
    65525: St := 'MT';
    65526: St := 'MM';
    65527: St := 'MD';
    65528: St := 'MC';
    65529: St := 'MB';
    65530: St := 'MR';
    65535: St := 'RT';
    65536: St := 'RM';
    65537: St := 'RD';
    65538: St := 'RC';
    65539: St := 'RB';
    65586, 105901: St := 'BT';
    65587, 105902: St := 'BM';
    65588, 105903: St := 'BD';
    65589, 105904: St := 'BC';
    65590, 105905: St := 'BB';
    105909: St := 'AS';
  else Exit;
  end;
  iHelpContext := 0;
  if (Length(St) > 0) then
    case St[1] of
      'A': iHelpContext := 110000241; //Articles
      'C': iHelpContext := 110000242; //Clients
      'E': iHelpContext := 110000245; //Etablissements
      'V': iHelpContext := 0; //Commerciaux
      'F': iHelpContext := 110000243; //Fournisseurs
      'M': iHelpContext := 0; //Affaire
      'R': iHelpContext := 0; //Ressources
      'B': iHelpContext := 110000244; //Contacts
    end;
  GCModifTitreZonesLibres(St);
  ParamTable('GCZONELIBRE', taModif, iHelpContext, PRien, 3);
  GCModifTitreZonesLibres('');
  AvertirTable('GCZONELIBRE');
  AfterChangeModule(65);
  AfterChangeModule(105);
end;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var StVire, StSeule: string;
begin
  case Num of
    10: //  après connexion
      begin
        UpdateCombosGC;
        GCModifLibelles;
        GCTripoteStatus;
        {$IFDEF NOMADE}
        // Chargement des sites au démarrage de l'appli.
        PCP_LesSites := TCollectionSites.Create(TCollectionSite, True);

        if VerifEtatBasePCP() < 3 then
        begin
          // Lancement automatique de l'assistant PCP
          PCPAssistInit('');
        end;
        {$ENDIF}
      end;
    11: ; //  après déconnexion
    12: ; //  avant connexion et séria
    {$IFDEF EAGLCLIENT}
    13:
      begin
        {$IFDEF GESCOM}
        LiberationCaisse;
        {$ENDIF}
        {$IFDEF AGL550B}
        ChargeMenuPop(-1, FMenuG.DispatchX);
        {$ELSE}
        ChargeMenuPop(TTypeMenu(-1), FMenuG.DispatchX);
        {$ENDIF}
      end;
    {$ELSE}
    {$IFDEF AGL545}
    13:
      begin
        {$IFDEF GESCOM}
        LiberationCaisse;
        {$ENDIF}
        {$IFDEF AGL550B}
        ChargeMenuPop(-1, FMenuG.DispatchX);
        {$ELSE}
        ChargeMenuPop(TTypeMenu(-1), FMenuG.DispatchX);
        {$ENDIF}
      end;
    {$ELSE}
    13:
      begin
        {$IFDEF GESCOM}
        LiberationCaisse;
        {$ENDIF}
        ChargeMenuPop(TTypeMenu(-1), FMenuG.Dispatch);
      end;
    {$ENDIF}
    {$ENDIF}
    15: ; //  ??????
    16: ; // DC 22/05/2003 le nouvel agl par paquet provoque une fonction 16 non disponible !!
    100: ; // Si lanceur
    /////// Menu Pop ////////////////
    27080: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_VENTEACHAT=VEN;GP_NATUREPIECEG=DE', '', 'CONSULTATION'); // devis pour GRC
    ///////  Ventes /////////////
       // Saisies pièces
    30201: CreerPiece('DE'); // Devis
    30202: CreerPiece('PRO'); // Proforma
    30241: CreerPiece('CC'); // Commande
    30242: CreerPiece('CCE'); // Commande échantillon
    30209: CreerPiece('PRE'); // Préparation livraison
    30251: CreerPiece('BLC'); // Livraison
    30252: CreerPiece('LCE'); // Livraison échantillon
    30210: CreerPiece('FPR'); // factures provisoires
    30205: CreerPiece('FAC'); // Facture
    30261: CreerPiece('AVC'); // Avoir
    30262: CreerPiece('AVS'); // Avoir sur stock
    // mofification pièces
    30221: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=DE', '', 'MODIFICATION'); //devis
    30222: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=PRO', '', 'MODIFICATION'); //Proforma
    30223: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=CC', '', 'MODIFICATION'); //Commandes
    30226: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=CCE', '', 'MODIFICATION'); //Commandes échantillon
    30229: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=PRE', '', 'MODIFICATION'); //Prépa
    30224: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=BLC', '', 'MODIFICATION'); //Livraison
    30227: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=LCE', '', 'MODIFICATION'); //Livraison échantillon
    30225: AGLLanceFiche('AFF', 'AFPIECE_MUL', 'GP_NATUREPIECEG=FPR', '', 'TABLE:GP;NOCHANGE_NATUREPIECE;NATURE:FPR;ACTION=MODIFICATION;');
      //Factures provisoires
    // Duplication unitaire pièces
    30231: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=DE', '', 'DE'); // Devis
    30232: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=PRO', '', 'PRO'); // Proforma
    30233: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=CC', '', 'CC'); // Commandes
    30239: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=PRE', '', 'PRE'); // Prépas
    30234: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=BLC', '', 'BLC'); // Livraisons
    30235: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=FAC', '', 'FAC'); // Factures
    30236: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=AVC', '', 'AVC'); // Avoirs
    30238: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=AVS', '', 'AVS'); // Avoirs sur stock
    30237: AGLLanceFiche('AFF', 'AFDUPLICPIECE_MUL', 'GP_NATUREPIECEG=FPR', '', 'NATURE:FPR;DUPPLIC:FPR;'); // Factures  provisoires
    // Consultation
    {$IFDEF NOMADE}
    30301: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=CC', '', 'CONSULTATION');
    30302: AGLLanceFiche('GC', 'GCLIGNE_MUL', 'GL_NATUREPIECEG=CC', '', 'CONSULTATION');
    {$ELSE} // NOMADE
    30301: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_VENTEACHAT=VEN;GP_NATUREPIECEG=FAC', '', 'CONSULTATION');
    30302: AGLLanceFiche('GC', 'GCLIGNE_MUL', 'GL_NATUREPIECEG=FAC', '', 'CONSULTATION');
    {$ENDIF} // NOMADE
    // Génération unitaire pièces
    30401: AGLLanceFiche('GC', 'GCPIECEVISA_MUL', 'GP_NATUREPIECEG=CC;GP_VENTEACHAT=VEN', '', 'VENTEACHAT=VEN'); // Visa des pièces
    30403: AGLLanceFiche('GC', 'GCTRANSPIECE_MUL', 'GP_NATUREPIECEG=DE', '', 'CC'); // Commandes
    30409: AGLLanceFiche('GC', 'GCTRANSPIECE_MUL', 'GP_NATUREPIECEG=CC', '', 'PRE'); // Prépas
    30421: AGLLanceFiche('GC', 'GCTRANSPIECE_MUL', 'GP_NATUREPIECEG=PRE', '', 'BLC'); // Livraisons
    30422: AGLLanceFiche('GC', 'GCTRANSPIECE_MUL', 'GP_NATUREPIECEG=CCE', '', 'LCE'); // Livraisons échantillon
    30405: AGLLanceFiche('GC', 'GCTRANSPIECE_MUL', 'GP_NATUREPIECEG=BLC', '', 'FAC'); // Factures
    // Génération manuelle pièces
    30431: AGLLanceFiche('GC', 'GCGROUPEMANPIECE', 'GP_NATUREPIECEG=CC', '', 'VENTE;CCR;BLC'); // Comm --> Livr
    30432: AGLLanceFiche('GC', 'GCGROUPEMANPIECE', 'GP_NATUREPIECEG=BLC', '', 'VENTE;LCR;FAC'); // Livr --> Fact
    // Génération automatique pièces
    30413: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=CC', '', 'VENTE;PRE'); // Com --> Prep
    30415: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=CC', '', 'VENTE;BLC'); // Com --> Livr
    30411: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=PRE', '', 'VENTE;BLC'); // Prep --> Livr
    30412: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=BLC', '', 'VENTE;FAC'); // Livr --> Fact
    30414: AGLLanceFiche('AFF', 'AFPIECEPRO_MUL', '', '', 'TABLE:GP;NATURE:FPR;NOCHANGE_NATUREPIECE'); //validation factures provisoires
    30407: AGLLanceFiche('AFF', 'AFPREPFACT_MUL', '', '', ''); // Prépa facture par affaires
    30408: AglLanceFiche('AFF', 'AFPIECEPROANU_MUL', '', '', 'TABLE:GP;NATURE:FPR;NOCHANGE_NATUREPIECE'); // Annulation facture provisoire
    // PasRel
    30406: AGLLanceFiche('GC', 'GCEXPORT_PASREL', 'GP_VENTEACHAT=VEN', '', ''); //Service Pas.Rel

    // Tarifs HT
    {$IFNDEF SANSPARAM}
    30101: AGLLanceFiche('GC', 'GCASSISTANTTARIF', '', '', 'VEN'); // Assistant création tarif article
    30102: EntreeTarifArticle(taModif);
    30103: EntreeTarifTiers(taModif);
    30104: EntreeTarifCliArt(taModif);
    30105: EntreeTarifCatArt(taModif);
    30106: AGLLanceFiche('GC', 'GCTARIFCON_MUL', '', '', '');
    30107: AGLLanceFiche('GC', 'GCTARIFMAJ_MUL', '', '', '');
    30108: AGLLanceFiche('GC', 'GCTARIFMAJ_MUL', '', '', 'DUP'); // duplication de tarif
    30109: AGLLanceFiche('GC', 'GCTARIFMAJ_MUL', '', '', 'SUP'); // suppression de tarif
    // Tarifs TTC
    30902: EntreeTarifArticle(taModif, TRUE);
    30903: EntreeTarifTiers(taModif, TRUE);
    30904: EntreeTarifCliArt(taModif, TRUE);
    {$ENDIF}
    // Editions
    {$IFDEF NOMADE}
    30601: AGLLanceFiche('GC', 'GCEDITDOCDIFF_MUL', 'GP_NATUREPIECEG=CC', '', 'VEN'); // Editions différées
    {$ELSE} // NOMADE
    30601: AGLLanceFiche('GC', 'GCEDITDOCDIFF_MUL', '', '', 'VEN'); // Editions différées
    {$ENDIF} // NOMADE
    30602: AGLLanceFiche('GC', 'GCEDITTRAITEDIF', '', '', ''); // Edition traites différées
    30603: AGLLanceFiche('GC', 'GCPTFPIECE', '', '', 'VENTE'); // Portefeuille pièces
    30604: AGLLanceFiche('GC', 'GCETATCOMMISSION', '', '', ''); // commissionnement commerciaux
    30605: AGLLanceFiche('GC', 'GCETIQCLI', '', '', ''); // Etiquettes clients
    {$IFDEF CEGID}
    // spécifs CEGID
    30610: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE'); // portefeuille articles
    30611: AGLLanceFiche('GC', 'GCPTFPIECECEGID02', '', '', 'VENTE'); // portefeuille commercial
    30621: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;005'); // portefeuille articles
    30622: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;006'); // portefeuille articles
    30623: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;007'); // portefeuille articles
    30624: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;008'); // portefeuille articles
    30625: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;009'); // portefeuille articles
    30626: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;010'); // portefeuille articles
    30627: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;011'); // portefeuille articles
    30628: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;012'); // portefeuille articles
    30641: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;001'); // portefeuille articles
    30642: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;002'); // portefeuille articles
    30643: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;003'); // portefeuille articles
    30644: AGLLanceFiche('GC', 'GCPTFPIECECEGID01', '', '', 'VENTE;004'); // portefeuille articles
    {$ENDIF}
    // Suivi clients
    30701: AGLLanceFiche('CP', 'EPBALAGEE', '', '', ''); // Balance agée
    30702: AGLLanceFiche('CP', 'EPGLAGE', '', '', ''); // Grand livre
    30703: AGLLanceFiche('CP', 'EPECHEANCIER', '', '', ''); // Echéancier
    // fichiers
    {$IFDEF NOMADE}
    30501: if ctxMode in V_PGI.PGIContexte
      //              then AGLLanceFiche('GC','GCCLIMUL_MODE','T_NATUREAUXI=CLI','','ACTION=CONSULTATION')
      //              else AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=CLI','','ACTION=CONSULTATION') ;    // Clients
      then AGLLanceFiche('GC', 'GCCLIMUL_MODE', 'T_NATUREAUXI=CLI', '', '')
      else AGLLanceFiche('GC', 'GCTIERS_MUL', 'T_NATUREAUXI=CLI', '', ''); // Clients
    {$ELSE} // NOMADE
    30501: if ctxMode in V_PGI.PGIContexte
      then AGLLanceFiche('GC', 'GCCLIMUL_MODE', 'T_NATUREAUXI=CLI', '', '')
      else AGLLanceFiche('GC', 'GCTIERS_MUL', 'T_NATUREAUXI=CLI', '', ''); // Clients
    {$ENDIF} // NOMADE
    30502: AGLLanceFiche('YY', 'YYCONTACTTIERS', 'T_NATUREAUXI=CLI', '', ''); // contacts
    30503: AGLLanceFiche('GC', 'GCCOMMERCIAL_MUL', '', '', ''); // Commerciaux
    // GRC
    //30504 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','CLI') ;  // modification en série des tiers
    30504: AGLLanceFiche('RT', 'RTQUALITE', '', '', 'CLI;GC'); // modification en série des tiers
    {$IFNDEF SANSCOMPTA}
    //30505 : OuvreFermeCpte(fbAux,FALSE,'CLI') ;  // fermeture des comptes tiers
    //30506 : OuvreFermeCpte(fbAux,True,'CLI') ;   // ouverture compte tiers
    30507: SuppressionCpteAuxi('CLI');
    {$ENDIF}
    30506: AGLLanceFiche('RT', 'RTOUVFERMTIERS', '', '', 'OUVRE'); // ouverture compte tiers
    30505: AGLLanceFiche('RT', 'RTOUVFERMTIERS', '', '', 'FERME'); // fermeture compte tiers

    ///////  Achats  /////////////
       // Saisie pièces
    31201: CreerPiece('CF');
    31202: CreerPiece('BLF');
    31203: CreerPiece('FF');
    // Modifications
    31211: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=CF', '', 'MODIFICATION'); //commandes fournisseur
    31212: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=BLF', '', 'MODIFICATION'); //bons de livraison fournisseur
    // Duplication unitaire pièces
    31221: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=CF', '', 'CF');
    31222: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=BLF', '', 'BLF');
    31223: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=FF', '', 'FF');
    // Consultation
    31101: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=FF', '', 'CONSULTATION');
    31102: AGLLanceFiche('GC', 'GCLIGNEACH_MUL', 'GL_NATUREPIECEG=FF', '', 'CONSULTATION');
    // génération
    31501: AGLLanceFiche('GC', 'GCPIECEVISA_MUL', 'GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH', '', 'VENTEACHAT=ACH'); // Visa des pièces
    31513: AGLLanceFiche('GC', 'GCCONTREMVTOA', '', '', 'MODE=VTOA'); // génération commande fournisseur de contremarque
    31511: AGLLanceFiche('GC', 'GCTRANSACH_MUL', 'GP_NATUREPIECEG=CF', '', 'BLF');
    31512: AGLLanceFiche('GC', 'GCTRANSACH_MUL', 'GP_NATUREPIECEG=BLF', '', 'FF');
    // Génération manuelle pièces
    31531: AGLLanceFiche('GC', 'GCGROUPEMANPIECE', 'GP_NATUREPIECEG=CF', '', 'ACHAT;CFR;BLF'); // Comm --> Livr
    31532: AGLLanceFiche('GC', 'GCGROUPEMANPIECE', 'GP_NATUREPIECEG=BLF', '', 'ACHAT;LFR;FF'); // Livr --> Fact
    // génération automatique
    31521: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=CF', '', 'ACHAT;BLF');
    31522: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=BLF', '', 'ACHAT;FF');
    // tarif fournisseurs
    {$IFNDEF SANSPARAM}
    31401: EntreeTarifFouArt(taModif); // saisie
    31402: AGLLanceFiche('GC', 'GCTARIFFOUCON_MUL', '', '', ''); // consultation
    31403: AGLLanceFiche('GC', 'GCTARIFFOUMAJ_MUL', '', '', ''); // mise à jour
    31404: AGLLanceFiche('GC', 'GCPARFOU_MUL', '', '', ''); // Import
    {$ENDIF}
    // génération
    31301:
      begin
        Assist_Suggestion;
        RetourForce := TRUE;
      end; // Suggestion de réappro
    31302: AGLLanceFiche('GC', 'GCLANCEREA_MUL', '', '', ''); // Lancement réappro
    // Editions
    31601: AGLLanceFiche('GC', 'GCEDITDOCDIFF_MUL', '', '', 'ACH'); // Editions différées
    31602: AGLLanceFiche('GC', 'GCPTFPIECE', '', '', 'ACHAT'); // Portefeuille pièces
    31612: AGLLanceFiche('GC', 'GCART_NONREF', '', '', '');
    31611: AGLLanceFiche('GC', 'GCFOURN_ARTICLE', '', '', '');
    31613: AGLLanceFiche('GC', 'GCARTICLE_FOURN', '', '', '');
    // Fichiers
    {$IFDEF NOMADE}
    //   31701 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','ACTION=CONSULTATION') ;
    31701: AGLLanceFiche('GC', 'GCFOURNISSEUR_MUL', 'T_NATUREAUXI=FOU', '', '');
    31702: AGLLanceFiche('GC', 'GCMULCATALOGUE', '', '', 'ACTION=CONSULTATION');
    {$ELSE}
    31701: AGLLanceFiche('GC', 'GCFOURNISSEUR_MUL', 'T_NATUREAUXI=FOU', '', '');
    31702: AGLLanceFiche('GC', 'GCMULCATALOGUE', '', '', '');
    {$ENDIF} //NOMADE
    31703: AGLLanceFiche('YY', 'YYCONTACTTIERS', 'T_NATUREAUXI=FOU', '', ''); // contacts
    // GRC
    //31704 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','FOU') ;  // modification en série des tiers
    31704: AGLLanceFiche('RT', 'RTQUALITE', '', '', 'FOU;GC'); // modification en série des tiers
    {$IFNDEF SANSCOMPTA}
    31705: OuvreFermeCpte(fbAux, FALSE, 'FOU'); // fermeture des comptes tiers
    31706: OuvreFermeCpte(fbAux, True, 'FOU'); // ouverture compte tiers
    31707: SuppressionCpteAuxi('FOU');
    {$ENDIF}
    ///////  Articles et Stocks  /////////////////
       // Articles
    {$IFDEF NOMADE}
    32501: AGLLanceFiche('GC', 'GCARTICLE_MUL', '', '', 'GCARTICLE;ACTION=CONSULTATION'); // article
    {$ELSE}
    32501: AGLLanceFiche('GC', 'GCARTICLE_MUL', '', '', 'GCARTICLE'); // article
    {$ENDIF}
    32502: AglLanceFiche('GC', 'GCPROFILART', '', '', ''); // Profil article
    32503: AGLLanceFiche('GC', 'GCARTICLE_MODIFLO', '', '', ''); // modification en série
    32504: AGLLanceFiche('GC', 'GCMULSUPPRART', '', '', ''); // épuration articles
    32506: AGLLanceFiche('GC', 'GCTARIFART_MUL', '', '', ''); // Maj BAses Tarifaires
    32505: EntreeTradArticle(taModif); // traduction libellé articles
    // Consultation
    32110: AGLLanceFiche('GC', 'GCDISPO_MUL', '', '', '');
    32111: AGLLanceFiche('GC', 'GCMOUVSTK', '', '', ''); // stock prévisionnel
    32112: AGLLanceFiche('GC', 'GCMULCONSULTSTOCK', '', '', ''); // stock prévisionnel
    32113: AGLLanceFiche('GC', 'GCCONTREAVT', '', '', ''); // suivi contremarque
    // Traitement
    {$IFNDEF EAGLCLIENT}
    //32201 : RetourForce:=Not CreerTransfert('TEM');        // Saisie des Transferts Inter-Boutiques
    32201: AGLLanceFiche('GC', 'GCMOUVSTKEX', '', '', 'ACTION=CREATION;TEM'); // Transferts Inter-Dépôt
    {$ENDIF}
    32221: AGLLanceFiche('GC', 'GCMOUVSTKEX', '', '', 'ACTION=CREATION;EEX');
    32222: AGLLanceFiche('GC', 'GCMOUVSTKEX', '', '', 'ACTION=CREATION;SEX');
    32231: AGLLanceFiche('GC', 'GCMOUVSTKEX', '', '', 'ACTION=CREATION;RCC'); // Réajustement réservé client
    32232: AGLLanceFiche('GC', 'GCMOUVSTKEX', '', '', 'ACTION=CREATION;RCF'); // Réajustement réservé fournisseur
    32233: AGLLanceFiche('GC', 'GCMOUVSTKEX', '', '', 'ACTION=CREATION;RPR'); // Réajustement préparé client
    32203: AGLLanceFiche('GC', 'GCTTFINMOIS', '', '', '');
    32204: AGLLanceFiche('GC', 'GCTTFINEXERCICE', '', '', '');
    32205: AGLLanceFiche('GC', 'GCAFFECTSTOCK', '', '', '');
    32206: AGLLanceFiche('GC', 'GCINITSTOCK', '', '', ''); // Initialisation des stocks
    // inventaire ////
    32310: AGLLanceFiche('GC', 'GCLISTEINV', '', '', '');
    32315: AGLLanceFiche('GC', 'GCLISTEINVPRE', '', '', '');
    32317: AGLLanceFiche('GC', 'GCLISTEINVVAL', '', '', '');
    32320: AGLLanceFiche('GC', 'GCSAISIEINV_MUL', '', '', '');
    32330: AGLLanceFiche('GC', 'GCVALIDINV_MUL', '', '', '');
    32340: AGLLanceFiche('GC', 'GCLISTEINV_MUL', '', '', '');
    // Editions ////
    32401: AGLLanceFiche('GC', 'GCARTDISPO', '', '', '');
    32403: AGLLanceFiche('GC', 'GCINVPERM', '', '', ''); // inventaire permanent
    32404: AGLLanceFiche('GC', 'GCETIQART', '', '', '');
    {$IFNDEF EAGLCLIENT}
    32405: if not AppelConsostock then Retourforce := true;
    {$ENDIF}
    // Fichiers
    32601: AGLLanceFiche('GC', 'GCDEPOT_MUL', '', '', '');
    32602: AGLLanceFiche('GC', 'GCEMPLACEMENT_MUL', '', '', '');
    32603: AGLLanceFiche('GC', 'GCPRIXREVIENT', '', '', '');
    // Dimensions ////
    32701:
      begin
        ParamTable('GCCATEGORIEDIM', taModif, 110000045, PRien, 3);
        AfterChangeModule(32);
      end; // categorie
    32702: AGLLanceFiche('GC', 'GCMASQUEDIM', '', '', '');
    {$IFNDEF SANSPARAM}
    {$IFNDEF EAGLCLIENT}32703: ParamDimension;
    {$ENDIF}
    {$ENDIF}
    32711..32715: ParamTable('GCGRILLEDIM' + IntToStr(Num - 32710), taCreat, 110000045, PRien);
    ///////////// Analyses ///////////////////////
       // Ventes ////
    33101: AGLLanceFiche('GC', 'GCCTRLMARGE', '', '', '');
    33102: AGLLanceFiche('GC', 'GCANALYSE_VENTES', 'GPP_VENTEACHAT=VEN', '', 'CONSULTATION');
    33105: AGLLanceFiche('GC', 'GCVENTES_CUBE', 'GPP_VENTEACHAT=VEN', '', 'CONSULTATION');
    33106: AGLLanceFiche('GC', 'GQR1SCC', '', '', '');
    33107: AGLLanceFiche('GC', 'GCNVXCLIENTS', '', '', ''); // Nouveaux Clients // AGLLanceFiche('GC','GQR1SCN','','','') ;
    33111: AGLLanceFiche('GC', 'GCGRAPHSTAT', '', '', 'VEN');
    33112: AGLLanceFiche('GC', 'GCSTAT', '', '', 'VEN');
    {$IFDEF CEGID}
    33113: AGLLanceFiche('GC', 'GCSTATCEGID', '', '', 'VEN');
    33171: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;002'); // Stats cegid par Marché
    33172: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;003'); // Stats cegid par Marché
    33173: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;009'); // Stats cegid par Marché
    33174: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;008'); // Stats cegid par Marché
    33151: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;001'); // Stats cegid par Agence
    33152: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;007'); // Stats cegid par Agence
    33153: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;006'); // Stats cegid par Agence
    33161: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;004'); // Stats cegid par Canal
    33162: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;010'); // Stats cegid par Canal
    33163: AGLLanceFiche('GC', 'GCSTATCEGID01', '', '', 'VEN;005'); // Stats cegid par Canal
    {$ENDIF}
    33121: AGLLanceFiche('GC', 'GCGRAPHPALM', '', '', 'VEN');
    33122: AGLLanceFiche('GC', 'GCPALMARES', '', '', 'VEN');
    33132: AGLLanceFiche('AFF', 'AFVENTES_CUBE', 'GPP_VENTEACHAT=VEN', '', 'CONSULTATION');
    33131: AGLLanceFiche('AFF', 'ETATPREVCA', '', '', '');
    {$IFDEF CEGID}
    33141: AGLLanceFiche('GC', 'GQR1GPA', '', '', '');
    33142: AGLLanceFiche('GC', 'GQR1GII', '', '', '');
    33143: AGLLanceFiche('GC', 'GQR1GCI', '', '', '');
    33144: AGLLanceFiche('GC', 'GQR1GTV', '', '', '');
    {$ENDIF}
    // Achats ////
    33201: AGLLanceFiche('GC', 'GCCATANA', '', '', '');
    33202: AGLLanceFiche('GC', 'GCACHAT_CUBE', 'GPP_VENTEACHAT=ACH', '', '');
    33211: AGLLanceFiche('GC', 'GCGRAPHSTAT', '', '', 'ACH');
    33212: AGLLanceFiche('GC', 'GCSTAT', '', '', 'ACH');
    33221: AGLLanceFiche('GC', 'GCGRAPHPALM', '', '', 'ACH');
    33222: AGLLanceFiche('GC', 'GCPALMARES', '', '', 'ACH');
    // Tableaux de bord Affaires
    33401: AGLLanceFiche('AFF', 'AFTBVIEWER', '', '', ''); // Consult tableau de bord
    33402: AGLLanceFiche('AFF', 'AFTBVIEWERMULTI', '', '', 'MULTI:X'); // Consult TB Multi affaires
    33403: AGLLanceFiche('AFF', 'AFGRANDLIVRE', '', '', ''); //grand livre ;
    33404: AGLLanceFiche('AFF', 'AFBALANCE', '', '', ''); //balance;
    33405: AGLLanceFiche('AFF', 'AFTABLEAUBORD', '', '', ''); //AlimTableauBordAffaire;
    // Stocks
    33311: AGLLanceFiche('GC', 'GCTRACABILITE', '', '', 'LOT'); // Traçabilité des lots
    33312: AGLLanceFiche('GC', 'GCTRACABILITE', '', '', 'SERIE'); // Traçabilité des numéros de série
    33321: GCLanceFiche_VerifLigneSerie('GC', 'GCVERIFLIGNESERIE', '', '', 'CLI'); // Vérification des ventes des numéros de série
    33322: GCLanceFiche_VerifLigneSerie('GC', 'GCVERIFLIGNESERIE', '', '', 'FOU'); // Vérification des achats des numéros de série
    // Divers
    {$IFNDEF EAGLCLIENT}
    33501: Assist_DEBC; // Declaration d'échange de biens (DEB)
    {$ENDIF}

    {$IFNDEF SANSPARAM}
    {Fonctions E-COMMERCE}
    {$IFNDEF EAGLCLIENT}
    59101: ExportAllTables;
    59102: ImportAllTables;
    59103: BalanceTradTablette;
    59201: ImporteECommande;
    59202: AGLLanceFiche('E', 'EIMPORTCMD', '', '', '');
    59203: AGLLanceFiche('E', 'ECMDVALID_MUL', '', '', '');
    59301: AGLLanceFiche('E', 'EEXPORTARTICLE', '', '', '');
    59302: AGLLanceFiche('E', 'EEXPORTTIERS', '', '', '');
    59303: AGLLanceFiche('E', 'EEXPORTTARIF', '', '', '');
    59304: AGLLanceFiche('E', 'EEXPORTCOMM', '', '', '');
    59305: AGLLanceFiche('E', 'EEXPORTPAYS', '', '', '');
    59401: AGLLanceFiche('E', 'ETRADTABLETTE', '', '', '');
    {$ENDIF}
    {$ENDIF}
    ////// Affaires
       // Fichiers affaires
    70101: AglLanceFiche('AFF', 'AFFAIRE_MUL', 'AFF_STATUTAFFAIRE=PRO', '', 'STATUT=PRO;NOCHANGESTATUT'); // proposition d'affaires
    70102: AglLanceFiche('AFF', 'AFFAIRE_MUL', 'AFF_STATUTAFFAIRE=AFF', '', 'STATUT=AFF;NOCHANGESTATUT'); // affaires
    70111: AglLanceFiche('AFF', 'AFFAIREAVT_MUL', 'AFF_STATUTAFFAIRE=PRO', '', 'STATUT=PRO;NOCHANGESTATUT'); // proposition de mission
    70112: AglLanceFiche('AFF', 'AFFAIREAVT_MUL', 'AFF_STATUTAFFAIRE=AFF', '', 'STATUT=AFF;NOCHANGESTATUT'); // mission
    70103: AglLanceFiche('AFF', 'RESSOURCE_MUL', '', '', ''); // ressources
    70104: AGLLanceFiche('AFF', 'AFARTICLE_MUL', 'GA_TYPEARTICLE=PRE', '', 'PRE'); // Prestations
    70105: AGLLanceFiche('AFF', 'AFARTICLE_MUL', 'GA_TYPEARTICLE=FRA', '', 'FRA'); // Frais
    70107: AGLLanceFiche('AFF', 'AFARTICLE_MUL', 'GA_TYPEARTICLE=CTR', '', 'CTR'); // articles de contrat
    70106: AGLLanceFiche('AFF', 'HORAIRESTD', '', '', 'TYPE:STD'); // Calendriers
    // Editions Afaires
    70211: AGLLanceFiche('AFF', 'ETATSAFFAIRE', '', '', ''); // Etats sur affaires
    70212: AGLLanceFiche('AFF', 'FICHEAFFAIRE', '', '', '');
    70213: AGLLanceFiche('AFF', 'STATAFFAIRE', '', '', '');
    70221: AGLLanceFiche('AFF', 'ETATSRESSOURCE', '', '', ''); // Etats sur ressources
    70222: AGLLanceFiche('AFF', 'FICHERESSOURCE', '', '', '');
    70223: AGLLanceFiche('AFF', 'STATRESSOURCE', '', '', '');
    70224: AGLLanceFiche('AFF', 'ADRESSERESSOURCE', '', '', '');
    70230: AGLLanceFiche('AFF', 'AFPRINTCALENDRIER', '', '', '');
    // Documents clients
    70301: AGLLanceFiche('AFF', 'AFPROPOSEDITMUL', 'AFF_STATUTAFFAIRE=PRO', '', 'STATUT=PRO'); // Edition des propositions
    70302: AGLLanceFiche('AFF', 'AFPROPOSEDITMUL', 'AFF_STATUTAFFAIRE=AFF', '', 'STATUT=AFF'); // Edition des affaires
    // Saisie d'activité Par ressources
    {$IFDEF AFFAIRE}
    70411: AFCreerActivite(tsaRess, tacTemps, 'REA'); // Saisie d'activités Temps
    70412: AFCreerActivite(tsaRess, tacFrais, 'REA'); // Saisie d'activités Frais
    70413: AFCreerActivite(tsaRess, tacFourn, 'REA'); // Saisie d'activités Fournitures
    70414: AFCreerActivite(tsaRess, tacGlobal, 'REA'); // Saisie d'activités Globale
    // Saisie d'activité Par affaires
    70421: AFCreerActivite(tsaClient, tacTemps, 'REA'); // Saisie d'activités Temps
    70422: AFCreerActivite(tsaClient, tacFrais, 'REA'); // Saisie d'activités Frais
    70423: AFCreerActivite(tsaClient, tacFourn, 'REA'); // Saisie d'activités Fournitures
    70424: AFCreerActivite(tsaClient, tacGlobal, 'REA'); // Saisie d'activités Globale
    70430: AGLLanceFiche('AFF', 'AFACTIVITECON_MUL', '', '', ''); // Modif ligne activité
    {$ENDIF}
    // Consultations
    70701: AGLLanceFiche('AFF', 'AFINTERVENANTSAFF', ' ', '', ''); // intervenants sur affaires
    // Traitements affaires
    70501: AglLanceFiche('AFF', 'AFVALIDEPROPOS', 'AFF_STATUTAFFAIRE=PRO', '', 'PRO'); // acceptation de missions
    70502: AGLLanceFiche('AFF', 'AFAUGMEN_MUL', 'AFF_STATUTAFFAIRE=AFF', '', ''); // Augmentation affaire
    70503: AGLLanceFiche('AFF', 'AFALIGNAFF', '', '', ''); // Alignement client sur affaire
    70511: AGLLanceFiche('AFF', 'AFEPUR_ECH_AFF', '', '', ''); // Epuration échéances facturées
    70512: AGLLanceFiche('AFF', 'AFEPUR_ECH_AFF', '', '', 'GENERATION'); // génération échéances
    70504: AGLLanceFiche('AFF', 'AFALIGN_MONNAIE', '', '', '');
    70505: AGLLanceFiche('AFF', 'AFAFFAIR_MODIFLOT', '', '', '');
    // Traitements ressources
    70581: AGLLanceFiche('AFF', 'AFRESSOU_MODIFLOT', '', '', ''); // Modif en série des ressources
    // Traitements sur activité
    70551: AGLLanceFiche('AFF', 'AFACTIVITEGLO_MUL', '', '', 'AFA_REPRISEACTIV:TOU;'); // Modif ligne activité
    70552: AGLLanceFiche('AFF', 'AFACTIVITEVISAMUL', '', '', ''); // Visas sur l'activité
    70553: AGLLanceFiche('AFF', 'REVAL_RESS_MUL', '', '', ''); // Revalorisation de ressources

    // Editions sur activités
    70601: AGLLanceFiche('AFF', 'AFJOURNALACT', '', '', '');
    70602: AGLLanceFiche('AFF', 'AFJOURNALCLIENT', '', '', '');
    70611: AGLLanceFiche('AFF', 'AFSYNTHRESS', '', '', ''); // Synthèse assistant detaillée par affaire
    70612: AGLLanceFiche('AFF', 'AFSYNTHRESSART', '', '', ''); // Synthèse assistant detaillée par article
    70621: AGLLanceFiche('AFF', 'AFSYNTHCLIENT', '', '', ''); // Synthèse Client affaire détaillée par ressource
    70622: AGLLanceFiche('AFF', 'AFSYNTHCLIART', '', '', ''); // Synthèse Client affaire détaillée par article
    {$IFNDEF SANSCOMPTA}
    {$IFDEF V530}
    60101: GestionSociete(PRien, @InitSociete);
    {$ELSE}
    60101: GestionSociete(PRien, @InitSociete, nil);
    {$ENDIF}
    {$ENDIF}

    60201: AGLLanceFiche('YY', 'YYUSERGROUP', '', '', '') {FicheUserGrp};
    60202:
      begin
        FicheUSer(V_PGI.User);
        ControleUsers;
      end; // Utilisateurs

    {$IFNDEF EAGLCLIENT}
    60102: RecopieSoc(PRien, True); // recopie société à société
    //////// Utilisateurs et accès
    3172: AGLLanceFiche('YY', 'PROFILETABL', '', '', 'ETA');
    60203: ReseauUtilisateurs(False); // utilisateurs connectés
    60204: VisuLog; // Suivi d'activité
    60205: ReseauUtilisateurs(True); // RAZ connexions
    60206: AGLLanceFiche('GC', 'GCPARAMOBLIG', '', '', ''); // Champs obligatoires
    60207: AGLLanceFiche('GC', 'GCPARAMCONFID', '', '', ''); // restricyions fiches
    /////// Outils
    60301: AGLLanceFiche('GC', 'GCJNALEVENT_MUL', '', '', '');
    60401: AGLLanceFiche('GC', 'GCCPTADIFF', '', '', '');
    {$IFNDEF SANSPARAM}
    60402: EntreeStockAjust;
    60403: Entree_TraiteAdresse;
    {$ENDIF}
    /////// Utilitaires
    {$IFNDEF SANSPARAM}
    //60901 : BEGIN Assist_Import;  RetourForce:=TRUE;  END ;        // Import negoce
    {$IFDEF GRC}
    //60904 : BEGIN Assist_ImportProspect;  RetourForce:=TRUE;  END ;        // Import Prospect II
    {$ENDIF} {GRC}
    {$ENDIF} {SansParam}
    ///////// Paramètres /////////////
        /////// Paramètres - Société //////////
    65201:
      begin
        BrancheParamSocAffiche(StVire, StSeule);
        ParamSociete(False, StVire, StSeule, '', nil, ChargePageSoc, SauvePageSoc, InterfaceSoc, 110000220);
      end;
    65203: AGLLanceFiche('GC', 'GCETABLISS_MUL', '', '', '');
    65204:
      begin
        LanceAssistCodeAffaire;
        RetourForce := True;
      end;
    {$ENDIF} {eAGL CLient}
    /////// Gestion Commerciale /////////
    {$IFDEF VRELEASE}
    {$IFNDEF EAGLCLIENT}
    65471: if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
        AglLanceFiche('GC', 'GCPARPIECE', '', '', '') // LanceparPiece
      else AglLanceFiche('GC', 'GCPARPIECEUSR', '', '', ''); // LanceparPiece  utilisateur
    {$ELSE}
    65471: AglLanceFiche('GC', 'GCPARPIECEUSR', '', '', ''); // LanceparPiece  utilisateur
    {$ENDIF} // EAGLCLIENT
    {$ELSE}
    65471: AglLanceFiche('GC', 'GCPARPIECE', '', '', ''); // LanceparPiece ;
    {$ENDIF} // VRELEASE
    {$IFNDEF EAGLCLIENT}
    65472:
      begin
        EditDocument('L', 'GPI', '', True);
        RetourForce := True;
      end;
    {$IFNDEF SANSPARAM}65473: EntreeListeSaisie('');
    {$ENDIF}
    65474:
      begin
        EditEtat('E', 'GPJ', '', True, nil, '', '');
        RetourForce := True;
      end;
    65475: ParamTable('GCEQUIPIECE', taCreat, 0, PRien); // Natures de regroupement
    65401: AGLLanceFiche('GC', 'GCPORT_MUL', '', '', ''); // Port et frais
    {$IFNDEF SANSCOMPTA}65402: FicheSouche('GES');
    {$ENDIF}
    65403: AGLLanceFiche('GC', 'GCUNITEMESURE', '', '', '');
    65404: AglLanceFiche('GC', 'GCCODECPTA', '', '', ''); //parametrage ventilations comptables
    65408: AGLLanceFiche('AFF', 'VENTILANA', '', '', ''); // Ventil ana par affaire
    {$IFNDEF SANSPARAM}65405: EntreeArrondi(taModif);
    {$ENDIF}
    65410: ParamTable('YYDOMAINE', taCreat, 110000217, PRien); // Domaines activité
    65409: ParamTable('GCMOTIFMOUVEMENT', taCreat, 110000284, PRien); // Motif mouvement stock
    65406: ParamTable('GCEMPLOIBLOB', taCreat, 110000172, PRien); // Emploi blob
    65451: ParamTable('GCCOMMENTAIREENT', taCreat, 110000167, PRien, 6); // commentaires entete
    65453: ParamTable('GCCOMMENTAIREPIED', taCreat, 110000167, PRien, 6); // commentaires pied
    65452:
      begin // commentaires ligne
        if (ctxMode in V_PGI.PGIContexte) or (ctxChr in V_PGI.PGIContexte) or (ctxFO in V_PGI.PGIContexte) then
        begin
          ParamTable('GCCOMMENTAIRELIGNE', taCreat, 110000167, PRien, 6); // commentaires ligne
        end else
        begin
          if GetParamSoc('SO_GCCOMMENTAIRE') then AGLLanceFiche('GC', 'GCCOMMENTAIRE', '', '', '')
          else ParamTable('GCCOMMENTAIRELIGNE', taCreat, 110000167, PRien, 6);
        end;
      end;
    65422: AGLLanceFiche('AFF', 'AFBLOCAGEAFFAIRE', '', '', ''); // blocage affaire
    65421: AGLLanceFiche('AFF', 'AFPROFILGENER', '', '', 'TYPE:DEF'); // profils de facturation par affaire
    // Paramètres Tiers
    65101: ParamTable('GCZONECOM', taCreat, 110000156, PRien); // zones commerciales
    65102: ParamTable('GCEXPEDITION', taCreat, 110000156, PRien); // mode d'expedition
    65103: ParamTable('GCCOTEARTFOURN', taCreat, 110000156, PRien); // cote fournisseur
    65104: ParamTable('TTTARIFCLIENT', taCreat, 110000156, PRien); // code tarif client
    65105: ParamTable('GCCOMPTATIERS', taCreat, 110000156, PRien); // code comptable tiers
    65106: ParamTable('TTSECTEUR', taCreat, 110000156, PRien); // secteurs d'activité
    65107: ParamTable('GCORIGINETIERS', taCreat, 110000156, PRien); // Origine du tiers
    // Paramètres articles
    65301: ParamTable('GCCOLLECTION', taCreat, 0, PRien, 3, 'Collections'); // collection
    65302: ParamTable('GCTARIFARTICLE', taCreat, 0, PRien); // code tarif  article
    65303:
      begin // libellés des codes familles
        ParamTable('GCLIBFAMILLE', taModif, 0, PRien, 3, 'Titre des familles/sous-familles');
        AfterChangeModule(65);
        AfterChangeModule(105);
      end;
    65304: ParamTable('GCCOMPTAARTICLE', taCreat, 0, PRien); // code comptable article
    65305: ParamTable('GCTYPEEMPLACEMENT', taCreat, 0, PRien); // type emplacement
    65306: ParamTable('GCCOTEEMPLACEMENT', taCreat, 0, PRien); // cote emplacement
    65321..65323: ParamTable('GCFAMILLENIV' + IntToStr(Num - 65320), taCreat, 0, PRien, 3, RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Num - 65320), FALSE));
      // familles
    //// Paramètres généraux ///
    65901: ParamTable('ttFormeJuridique', taCreat, 110000187, PRien);
    65903: ParamTable('ttCivilite', taCreat, 110000189, PRien);
    65902: ParamTable('ttFonction', taCreat, 110000188, PRien);
    65904: AglLanceFiche('YY', 'YYPAYS', '', '', ''); // OuvrePays ;
    65905: FicheRegion('', '', False);
    65906: OuvrePostal(PRien);
    {$IFNDEF SANSCOMPTA}
    65907: FicheDevise('', tamodif, False);
    65908: ParamTable('TTREGIMETVA', taCreat, 110000183, PRien);
    65911: ParamTvaTpf(true);
    65912: ParamTvaTpf(false);
    {$ENDIF}
    65921: FicheModePaie_AGL('');
    65922: FicheRegle_AGL('', False, taModif);
    65923: AglLanceFiche('AFF', 'JOURFERIE', '', '', '');
    65924: ParamTable('ttLangue', taCreat, 110000190, PRien);
    /////// Paramètres - Tables libres //////////
    65501..65509: ParamTable('GCLIBREART' + IntToStr(Num - 65500), taCreat, 110000251, PRien, 6, RechDom('GCZONELIBRE', 'AT' + IntToStr(Num - 65500), FALSE));
    65510: ParamTable('GCLIBREARTA', taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'ATA', FALSE));
    65511..65519: ParamTable('GCLIBRETIERS' + IntToStr(Num - 65510), taCreat, 110000252, PRien, 6, RechDom('GCZONELIBRE', 'CT' + IntToStr(Num - 65510),
      FALSE));
    65520: ParamTable('GCLIBRETIERSA', taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'CTA', FALSE));
    65251..65259: ParamTable('AFTLIBREAFF' + IntToStr(Num - 65250), taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'MT' + IntToStr(Num - 65250), FALSE));
    65260: ParamTable('AFTLIBREAFFA', taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'MTA', FALSE));
    65261..65269: ParamTable('AFTLIBRERES' + IntToStr(Num - 65260), taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'RT' + IntToStr(Num - 65260), FALSE));
    65270: ParamTable('AFTLIBRERESA', taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'RTA', FALSE));
    65640: ParamTable('AFTRESILAFF', taCreat, 0, PRien);
    65541..65543: ParamTable('GCLIBREPIECE' + IntToStr(Num - 65540), taCreat, 110000256, PRien, 6);
    65545..65547: ParamTable('GCLIBREFOU' + IntToStr(Num - 65544), taCreat, 110000253, PRien, 6);
    65551..65559: ParamTable('YYLIBREET' + IntToStr(Num - 65550), taCreat, 110000255, PRien, 6, RechDom('GCZONELIBRE', 'ET' + IntToStr(Num - 65550), FALSE));
    65566..65568: ParamTable('YYLIBRECON' + IntToStr(Num - 65565), taCreat, 110000254, PRien, 6, RechDom('GCZONELIBRE', 'BT' + IntToStr(Num - 65565), FALSE));
      //tables libres contacts
    65560: ParamTable('YYLIBREETA', taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'ETA', FALSE));
    /////// Paramètres Affaires
    65601: ParamTable('AFCOMPTAAFFAIRE', taCreat, 110000214, PRien);
    65602: ParamTable('AFDEPARTEMENT', taCreat, 110000214, PRien);
    65603: ParamTable('AFTLIENAFFTIERS', taCreat, 110000214, PRien);
    65604: ParamTable('AFTRESILAFF', taCreat, 110000214, PRien);
    65605: ParamTable('AFETATAFFAIRE', taCreat, 110000214, PRien);
    65606: ParamTable('AFTREGROUPEFACT', taCreat, 0, PRien);
    /////// Paramètres Ressources
    65651: ParamTable('AFTTYPERESSOURCE', taCreat, 110000219, PRien);
    65652: ParamTable('AFTTARIFRESSOURCE', taCreat, 0, PRien);
    65653: AglLanceFiche('AFF', 'FONCTION', '', '', '');
    65654: AglLanceFiche('AFF', 'COMPETENCE', '', '', '');
    65655: ParamTable('AFTNIVEAUDIPLOME', taCreat, 0, PRien);
    65656: ParamTable('AFTSTANDCALEN', taCreat, 0, PRien);
    {Tables libres généralisées}
    65561..65565, 65571..65575, 65581..65590, 65592..65593, 65595..65597,
      65525..65530, 65535..65539: FactoriseZL(Num, PRien);
    {$ENDIF} // EAGLCLIENT
    {$IFDEF GRC}
    {$IFNDEF NOMADE}
    //////// Fiches ///////////
    92101: AGLLanceFiche('RT', 'RTPROSPECT_MUL', '', '', '');
    92102: AGLLanceFiche('YY', 'YYCONTACTTIERS', '', '', 'COURRIER');
    92103: AGLLanceFiche('RT', 'RTOPERATIONS_MUL', '', '', '');
    92104: AGLLanceFiche('RT', 'RTACTIONS_MUL', '', '', '');
    92107: AGLLanceFiche('RT', 'RTPERSPECTIVE_MUL', '', '', '');
    92108: AGLLanceFiche('RT', 'RTCONCURRENT_MUL', '', '', '');

    92105: if GetParamSoc('SO_RTPROJGESTION') = True then
        AGLLanceFiche('RT', 'RTPROJETS_MUL', '', '', '')
      else
        PGIInfo('Les projets ne sont pas gérés (paramètres société)', 'Projets');
    92106: AGLLanceFiche('RT', 'RTSUSPECT_MUL', '', '', '');
    92110: AGLLanceFiche('RT', 'RTOUVFERMTIERS', '', '', 'OUVRE'); // ouverture compte tiers
    92109: AGLLanceFiche('RT', 'RTOUVFERMTIERS', '', '', 'FERME'); // fermeture compte tiers

    ///// Editions ///////////////////
    92311: AGLLanceFiche('RT', 'RTACTIONS_ETAT2', '', '', '');
    92312: AGLLanceFiche('RT', 'RTTIERS_VISITE', '', '', '');
    92313: AGLLanceFiche('RT', 'RTTIERS_VISITE', '', '', 'COMMERCIAL');

    92320: AGLLanceFiche('RT', 'RTOPER_ETAT', '', '', '');

    92331: AGLLanceFiche('RT', 'RTFICHEPROSPECT', '', '', '');
    92332: AGLLanceFiche('GC', 'GCETIQCLI', '', '', '');
    92333: AGLLanceFiche('RT', 'RTLISTEHIERARCHIQ', '', '', '');

    92340: AGLLanceFiche('RT', 'RTPERSP_ETAT', '', '', '');

    92360: AGLLanceFiche('RT', 'RTOPER_ETAT_COUT', '', '', '');
    92370: AGLLanceFiche('RT', 'RTPROSSACTION_TV', '', '', '');

    ////// Traitements///////////////
    92202: AGLLanceFiche('RT', 'RTGENERE_ACTIONS', '', '', ''); // lancement opération
    92203: AGLLanceFiche('RT', 'RTPROS_MUL_MAILIN', '', '', '');
    //92204 : AGLLanceFiche('RT','RTPERSPECTIV_ETAT','','','') ;
    92205: AGLLanceFiche('RT', 'RTPROS_MAILIN_FIC', '', '', '');

    92270: AGLLanceFiche('RT', 'RTRECUP_SUSPECT', '', '', '');
    // suspects
    92271: AGLLanceFiche('RT', 'RTSUSP_MUL_MAILIN', '', '', '');
    92272: AGLLanceFiche('RT', 'RTSUSP_MAILIN_FIC', '', '', '');

    // mng
    92282: AGLLanceFiche('RT', 'RTQUALITE', '', '', 'CLI;GRC');
    92284: AGLLanceFiche('RT', 'RTQUALITE', '', '', 'PRO;GRC');
    92286: AGLLanceFiche('RT', 'RTQUALITE', '', '', 'CON;GRC');
    92288: AGLLanceFiche('RT', 'RTQUALITE_SUSP', '', '', 'SUS;GRC');

    92290: AGLLanceFiche('RT', 'RTDOUBLONS', '', '', '');
    92291: AGLLanceFiche('RT', 'RTSUSP_DOUBLE', '', '', '');

    ///// Analyse ////////
    92402: AGLLanceFiche('RT', 'RTANAPROSP', '', '', '');
    92411: AGLLanceFiche('RT', 'RTPROSPECTS_TV', '', '', '');
    92412: AGLLanceFiche('RT', 'RTACTIONS_TV', '', '', '');
    92413: AGLLanceFiche('RT', 'RTPERSPECTIVE_TV', '', '', '');
    {$IFDEF CEGID}
    92403: if (VH_RT.RTCodeCommercial <> '') and (VH_RT.RTNomCommercial <> '') then
        AGLLanceFiche('RT', 'RTCOMMERCIAL_PRO', '', '', ''); // synthese activite du mois
    {$ENDIF}
    //   92414 : AGLLanceFiche('RT','RTTVENCOURS','','','');

    92421: AGLLanceFiche('RT', 'RTTIERS_CUBE', '', '', '');
    92422: AGLLanceFiche('RT', 'RTACTIONS_CUBE', '', '', '');
    92423: AGLLanceFiche('RT', 'RTPERSP_CUBE', '', '', '');
    92424: AGLLanceFiche('RT', 'RTPERSPDETAIL_CUB', '', '', '');
    92425: AGLLanceFiche('RT', 'RTSUSPECTS_CUBE', '', '', '');

    92440: AGLLanceFiche('RT', 'RTPERSPSSDEV_TV', '', '', '');

    {$IFNDEF EAGLCLIENT}
    ////// Outils //////////
    92801: MajPhonetiqueTiers;
    92810: RTConversionPerspEuros;
    92802: CommercialToRessource;
    92820: RTCreerProspectPourClient;
    92830: IntegrationContactsRCC;
    92840: IntegrationAppelsRCC;
    {$IFDEF CEGID}
    92850: RTIntegrationRessourcePaie;
    //   92850 : RTIntegrationRessource;

    {$ENDIF}

    {$ENDIF}

    ////// Tablettes //////////
   { 92911 : begin ParamTable('RTTYPEACTION',taCreat,0,PRien) ;
                  AvertirTable ('RTTYPEACTION'); end;}
    92911:
      begin
        AGLLanceFiche('RT', 'RTTYPEACTIONS', '', '', '');
        ;
        AvertirTable('RTTYPEACTION');
      end;

    92912:
      begin
        ParamTable('RTTYPEPERSPECTIVE', taCreat, 0, PRien);
        AvertirTable('RTTYPEPERSPECTIVE');
      end;
    { suspects92913 : begin ParamTable('RTETATACTION',taModif,0,PRien) ;
                  AvertirTable ('RTETATACTION'); end;

    92914 : begin ParamTable('RTETATPERSPECTIVE',taModif,0,PRien) ;
                  AvertirTable ('RTETATPERSPECTIVE'); end;
    }
    92913:
      begin
        ParamTable('RTLIBCHPLIBSUSPECTS', taModif, 0, PRien);
        AvertirTable('RTLIBCHPLIBSUSPECTS');
      end;
    92914:
      begin
        ParamTable('RTMOTIFFERMETURE', taCreat, 0, PRien);
        AvertirTable('RTMOTIFFERMETURE');
      end;

    92915:
      begin
        ParamTable('RTLIBCLACTIONS', taModif, 0, PRien);
        AvertirTable('RTLIBCLACTIONS');
      end;
    92750:
      begin
        ParamTable('RTLIBCLPROPOSITIONS', taModif, 0, PRien);
        AvertirTable('RTLIBCLPROPOSITIONS');
      end;
    // suspect : rien à voir mais correction tacreat
    92916:
      begin
        ParamTable('RTOBJETOPE', taCreat, 0, PRien);
        AvertirTable('RTOBJETOPE');
      end;
    92917:
      begin
        ParamTable('RTMOTIFS', tacreat, 0, PRien);
        AvertirTable('RTMOTIFS');
        AvertirTable('RTMOTIFSIGNATURE');
      end;
    92950:
      begin
        ParamTable('RTIMPORTANCEACT', taCreat, 0, PRien); // Niveau d'importance des actions
        AvertirTable('RTIMPORTANCEACT');
      end;

    // suspect92920 : AglLanceFiche ('AFF','RESSOURCE_MUL','','','');  // ressources
    92918: AglLanceFiche('AFF', 'RESSOURCE_MUL', '', '', ''); // ressources
    92920..92929: ParamTable('RTRSCLIBTABLE' + IntToStr(Num - 92920), taCreat, 0, PRien, 3, RechDom('RTLIBCHPLIBSUSPECTS', 'CL' + IntToStr(Num - 92920),
      FALSE));

    92930: AGLLanceFiche('RT', 'RTPARAMCL', '', '', 'ZERO'); //ParamSaisieTL ();
    92945: AGLLanceFiche('RT', 'RTPARAMPROSPCOMPL', '', '', '');
    92946: if GetParamSoc('SO_RTCONFIDENTIALITE') = True then
        AGLLanceFiche('RT', 'RTUTILISATEUR_MUL', '', '', '')
      else
        PGIInfo('La confidentialité n''est pas gérée (paramètres société)', 'Confidentialité');

    92951..92960: ParamTable('RTRPRLIBTABLE' + IntToStr(Num - 92951), taCreat, 0, PRien, 3, RechDom('RTLIBCHAMPSLIBRES', 'CL' + IntToStr(Num - 92951), FALSE));
    92961..92976: ParamTable('RTRPRLIBTABLE' + IntToStr(Num - 92951), taCreat, 0, PRien, 3, RechDom('RTLIBCHAMPSLIBRES', 'CL' + Chr(Num - 92951 + 55), FALSE));
    92980..92989: ParamTable('RTRPRLIBTABMUL' + IntToStr(Num - 92980), taCreat, 0, PRien, 3, RechDom('RTLIBCHAMPSLIBRES', 'ML' + IntToStr(Num - 92980),
      FALSE));
    92991..92993: ParamTable('RTRPRLIBACTION' + IntToStr(Num - 92990), taCreat, 0, PRien, 3, RechDom('RTLIBCHAMPSLIBRES', 'AL' + IntToStr(Num - 92990),
      FALSE));
    92996..92998: ParamTable('RTRPRLIBPERSPECTIVE' + IntToStr(Num - 92995), taCreat, 0, PRien, 3, RechDom('RTLIBCHAMPSLIBRES', 'PL' + IntToStr(Num - 92995),
      FALSE));
    // mng : now 12 tables
    92932..92941: ParamTable('RTRPCLIBTABLE' + IntToStr(Num - 92932), taCreat, 0, PRien, 3, RechDom('RTLIBTABLECOMPL', 'TL' + IntToStr(Num - 92932), FALSE));
    92942: ParamTable('RTRPCLIBTABLEA', taCreat, 0, PRien, 3, RechDom('RTLIBTABLECOMPL', 'TLA', FALSE));
    92943: ParamTable('RTRPCLIBTABLEB', taCreat, 0, PRien, 3, RechDom('RTLIBTABLECOMPL', 'TLB', FALSE));
    92740: AGLLanceFiche('RT', 'RTPARAMCL', '', '', 'UN');
    92701..92705: ParamTable('RT001LIBTABLE' + IntToStr(Num - 92701), taCreat, 0, PRien, 3, RechDom('RT001LIBTABLE', 'TL' + IntToStr(Num - 92701), FALSE));
    92721..92725: ParamTable('RT001LIBTABMUL' + IntToStr(Num - 92721), taCreat, 0, PRien, 3, RechDom('RT001LIBTABMUL', 'TL' + IntToStr(Num - 92721), FALSE));
    {$ENDIF} // {$IFNDEF NOMADE}
    {$ENDIF}
    {$IFDEF MODE}
    ////////////////////////////// Menus de la Mode - Back-Office

       // 101 - Achats
    101101: AGLLanceFiche('MBO', 'ACHATCUB', '', '', 'ACH');
    101102: CreerPiece('ALF');
    101103: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=ALF', '', 'MODIFICATION');
    101104: AGLLanceFiche('GC', 'GCTRANSACH_MUL', 'GP_NATUREPIECEG=CF', '', 'ALF');
    101105: AGLLanceFiche('MBO', 'ACHATVUE', '', '', 'ACH');
    101106: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=BLF;GP_VENTEACHAT=ACH', '', 'CONSULTATION');
    101107: AGLLanceFiche('GC', 'GCLIGNEACH_MUL', 'GL_NATUREPIECEG=BLF', '', 'CONSULTATION');
    101108: AGLLanceFiche('MBO', 'ECARTALFBLF', 'NATURE=CF', '', 'CF'); // Edition des écarts entre commande fournisseur et réception
    101109: AGLLanceFiche('MBO', 'ECARTALFBLF', 'NATURE=ALF', '', 'ALF'); // Edition des écarts entre annonce de livraison et réception
    101110: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=FF', '', 'MODIFICATION');
    101111: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=ALF', '', 'ALF');

    // 102 - Ventes
    102101: AGLLanceFiche('MBO', 'MVTCLI_MODE', '', '', ''); // Liste des Pièces Clients/Articles
    102102: AGLLanceFiche('GC', 'GCPIEDECHET_MUL', '', '', 'MODIFICATION'); // Modif des règlements
    102103: CreerPiece('FFO'); // Saisie des Ventes F.O.
    102104: AGLLanceFiche('GC', 'GCTICKET_MUL', '', '', 'MODIFICATION'); //Modif des tickets F.O.
    102105: AGLLanceFiche('GC', 'GCEDTCRV_MODE', '', '', ''); // Compte-rendu des ventes détail
    102106: AGLLanceFiche('GC', 'GCEDTJCA_MODE', '', '', ''); // Journal de caisse
    102107: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'VEN');
    102108: AGLLanceFiche('GC', 'GCTICKET_MUL', '', '', 'CONSULTATION'); //Consultation des tickets F.O.
    102109: AGLLanceFiche('GC', 'GCLIGTICKET_MUL', 'GL_TYPEARTICLE=MAR', '', 'CONSULTATION'); // Consultation par lignes de ventes FFO
    102110: AGLLanceFiche('MBO', 'VENTEVUE', '', '', 'TYPESTA=VEN;TYPEART=ART');
    102111: AGLLanceFiche('GC', 'GCEDSYVTEA', 'TYPSYN=AD', 'AND', 'AD');
    102112: AGLLanceFiche('GC', 'GCEDSYVTEA', 'TYPSYN=AA', 'ANN', 'AA');
    102113: AGLLanceFiche('GC', 'GCGRAPHSYNA', '', '', 'AD');
    102114: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=FAC', '', 'MODIFICATION'); //Modification Factures
    102115: AGLLanceFiche('MBO', 'ANALREG', '', '', '');
    102116: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=AVC', '', 'MODIFICATION'); //Modification Avoir client
    102121: AGLLanceFiche('GC', 'GCEDSYVTEM', 'TYPSYN=MD', 'MOD', 'MD');
    102122: AGLLanceFiche('GC', 'GCEDSYVTEM', 'TYPSYN=MA', 'MON', 'MA');
    102123: AGLLanceFiche('GC', 'GCGRAPHSYNM', '', '', 'MD');
    102131: AGLLanceFiche('GC', 'GCGRAPHVIEPROD', '', '', '');
    102141: AGLLanceFiche('GC', 'GCEDTETVTE', '', '', '');
    102142: AGLLanceFiche('GC', 'GCGRAPHETVTE', '', '', '');
    102151: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'TYPECUB=VEN;TYPEART=ART');
    102152: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'TYPECUB=VEN;TYPEART=DIM');

    102201: EntreeTarifArticleMode(taModif, TRUE);
    102202: EntreeTarifTiersMode(taModif, TRUE);
    102203: AGLLanceFiche('MBO', 'TARIF_MUL', '', '', ''); // Tarif Mode consultation
    102204: AGLLanceFiche('MBO', 'TARIFPER', '', '', 'CREATION');
    102205: AGLLanceFiche('MBO', 'TARIFTYPE', '', '', 'CREATION');
    102206: RecopieTarifMode;
    102207: DispatchArtMode(4, '', '', 'GCARTICLE'); // Tarif détail - Maj tarifs mode
    102211: DispatchArtMode(7, '', '', 'ETATTARIF'); // Tarifs détail - Edition - tarifs
    102212: DispatchArtMode(7, '', '', 'COMPTARIF'); // Tarifs détail - Edition - comparatif
    102213: AGLLanceFiche('MBO', 'PANIER_MOYEN', '', '', ''); // Panier Moyen
    102214: AGLLanceFiche('MBO', 'ARTNONVENDUS', '', '', ''); // Articles non vendus

    // 103 - Stocks
    103101: AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULDISTANT'); //Consultation du stock à distance
    103102: DispatchArtMode(3, '', '', ''); // Consultation - Disponible article
    103103: DispatchArtMode(8, '', '', ''); // Stocks - Consultation - Statistique stock disponible
    103104: DispatchArtMode(9, '', '', ''); // Stocks - Consultation - Cube des stocks
    103105: DispatchArtMode(5, '', '', ''); // Edition - Etat du stock par taille
    103106: AGLLanceFiche('GC', 'GCTRANSFERT_MUL', 'GP_NATUREPIECEG=TEM', '', 'CONSULTATION'); // Consultation des Transferts Inter-Boutiques
    103107: AGLLanceFiche('GC', 'GCTRANSFERT_MUL', 'GP_NATUREPIECEG=TEM', '', 'MODIFICATION'); // Modification des Transferts Inter-Boutiques
    103108: AGLLanceFiche('MBO', 'PR_PROPTRF', '', '', ''); // Proposition de transfert
    103109: AGLLanceFiche('MBO', 'PR_MINMAX', '', '', ''); // Saisie des quantités Mini et Maxi par masque de dimension
    103110: ParamTable('GCTYPEAFFECTATION', taCreat, 0, PRien, 3);
    103111: AGLLanceFiche('MBO', 'TRANSTPI', '', '', ''); // Transmission inventaire depuis un TPI
    103112: AGLLanceFiche('MBO', 'TRANSINV_MUL', '', '', ''); // Saisie / Consultation des inventaires transmis
    103113: AGLLanceFiche('GC', 'GCINTEGRINV_MUL', '', '', ''); // Intégration des inventaires transmis
    103114: AGLLanceFiche('GC', 'GCLIGNEINV_MUL', '', '', 'CONSULTATION'); // Consultation des lignes d'écarts d'inventaire générées
    103115: AGLLanceFiche('GC', 'GCDISPODIST_MUL', '', '', '');
    103116: AGLLanceFiche('MBO', 'LISTEINV', '', '', ''); // Génération liste préparatoire d'inventaire
    103117: DispatchArtMode(1, 'TYPETRAIT=STK', '', 'ARTICLE;STOCKMIN'); // Génération du stock Mini et Maxi
    103118: AGLLanceFiche('GC', 'GCTRANSFERT_MUL', 'GP_NATUREPIECEG=TRV', '', 'VALIDATION');
    103119: Assist_FinInvent; // Assistant de fin d'inventaire
    103120: AGLLanceFiche('MBO', 'EDTECARTINV', '', '', ''); // Edition des écarts d'inventaire générés
    103121: AGLLanceFiche('MBO', 'EDTLISTEINV', '', '', ''); // Edition de l'inventaire
    103122: AGLLanceFiche('MBO', 'EDTCOMPSTKINV', '', '', ''); // Etat comparatif d'inventaire
    103123: AGLLanceFiche('MBO', 'ECARTTEMTRE', 'TRANSFERT=TEM;SITEDISTANT=', '', 'TEM'); // Edition des écarts entre transferts émis et reçus
    103124: AGLLanceFiche('MBO', 'ECARTTEMTRE', 'TRANSFERT=TRV;SITEDISTANT=X', '', 'TRV'); // Edition des écarts entre transferts à valider et reçus
    103125: CreerPiece('EEX');
    103126: CreerPiece('SEX');
    103127: AGLLanceFiche('GC', 'GCMOUVSTKEX_MUL', 'GP_NATUREPIECEG=EEX', '', 'CONSULTATION'); // Consultation des mouvements exceptionnels
    103128: AGLLanceFiche('GC', 'GCPIECEVISA_MUL', 'GP_NATUREPIECEG=TEM', '', 'TRANSFERT=TEM;VENTEACHAT=TRF'); // Visa des pièces
    103129: AGLLanceFiche('GC', 'GCEDITDOCDIFF_MUL', 'GP_NATUREPIECEG=TEM', '', 'STOCK'); // Editions différées
    103130: AGLLanceFiche('MBO', 'EDTTBBORD', '', '', ''); // Edition du tableau de bord
    103131: AGLLanceFiche('MBO', 'EXPORTPIECE_MUL', 'GP_NATUREPIECEG=TEM', '', ''); // Export ASCII des transferts
    103132: CreerTransfert('TEM'); // Saisie des Transferts Inter-Boutiques

    // 105 - Paramètres
 //   105101 : DispatchArtMode(10,'PRE','','TYPEARTICLE=PRE') ;  // Données de base - Articles - Prestations
    105101: LanceMulPrestation('ACTION=MODIFICATION'); // Données de base - Articles - Prestations
    105102: AGLLanceFiche('MBO', 'ARTFINANCIER', 'FI', '', 'TYPEARTICLE=FI'); // mode operation caisse AC
    105103: ParamTable('TTLIENPARENT', taCreat, 0, PRien, 3, 'Lien de parenté');
    105104: AGLLanceFiche('GC', 'GCFOUMUL_MODE', 'T_NATUREAUXI=FOU', '', '');
    105105: AGLLanceFiche('YY', 'YYCONTACTS_MUL', '', '', 'MODIFICATION');
    105106: AglLanceFiche('GC', 'GCVENDEUR', 'VEN', '', ''); // Vendeurs
    105107: AGLLanceFiche('MBO', 'EDTCLI_MODE', '', '', ''); // Edition des clients
    105108: AGLLanceFiche('MBO', 'EDTCPTCLI_MODE', '', '', ''); // Compte client
    105109: TraitClientSerie; // formatage des zones clients
    105110: if (not ExisteSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")'))
      then AGLLanceFiche('GC', 'GCCODEBARRES', 'SO;...', '', 'ACTION=CREATION')
      else AglLanceFiche('GC', 'GCCODEBARRES', '', 'SO;...', 'ACTION=MODIFICATION'); // Code à barres;
    105111: ParamTable('GCORIGINETIERS', taCreat, 0, PRien, 3, 'Origine des clients');
    105112: ParamTable('YYCODENAF', taCreat, 0, PRien, 3, 'Code NAF');
    105113: ParamTable('GCTYPESTATPIECE', taCreat, 0, PRien, 3);
    105114: AGLLanceFiche('GC', 'GCCODESTATPIECE', '', '', 'ACTION=MODIFICATION');
    105115: AGLLanceFiche('MBO', 'TYPEMASQUE', '', '', ''); // Type de masque
    105116: DispatchArtMode(1, '', '', 'ARTICLE'); // Mul Article Mode
    105117: AGLLanceFiche('GC', 'GCFAMHIER', '', '', 'ACTION=MODIFICATION;NIV1=FN1;NIV2=FN2;NIV3=FN3');
    105118: DispatchArtMode(1, '', '', 'GENERECAB'); // Génération codes à barres
    105119: AGLLanceFiche('MBO', 'EDTCOM', '', '', ''); // Liste des Commerciaux
    105120: AGLLanceFiche('MBO', 'EDTETA', '', '', '');
    105121: DispatchArtMode(1, '', '', 'ETIQARTCAT'); // Etiquettes articles sur catalogue
    105122: AGLLanceFiche('GC', 'GCETIARTDEM_MODE', '', '', ''); // Etiquettes articles à la demande
    105123: AGLLanceFiche('GC', 'GCETIARTFFO_MODE', 'GL_NATUREPIECEG=FFO', '', 'FFO'); // Etiquettes articles sur retour de vente
    105124: AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=CF', '', 'CF'); // Etiquettes articles sur commandes
    105125: AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=BLF', '', 'BLF'); // Etiquettes articles sur receptions
    105126: AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=TRE', '', 'TRE'); // Etiquettes articles sur transferts
    105127: DispatchArtMode(2, '', '', ''); // Etiquettes articles sur stock
    105128: AGLLanceFiche('GC', 'GCETIARTDOC_MODE', 'GP_NATUREPIECEG=ALF', '', 'ALF'); // Etiquettes articles sur annonce de livraison
    105129: DispatchArtMode(6, '', '', ''); // Article - Liste des articles
    105130: AGLLanceFiche('GC', 'GCVENTILANA', '', '', ''); // Ventil ana par mode
    {$IFNDEF EAGL}
    105131:
      begin
        LanceEtatLibreGC;
        RetourForce := True
      end; // Lanceur d'états libres utilisateurs
    105132:
      begin
        EditEtatS5S7('E', 'UBO', '', True, nil, '', '');
        RetourForce := True
      end; // Editeur d'états libres utilisateurs
    {$ENDIF}
    105133:
      begin // categories article
        {$IFDEF MODES3}
        ParamTable('GCCATEGORIEDIMS3', taModif, 110000045, PRien, 3);
        {$ELSE}
        ParamTable('GCCATEGORIEDIM', taModif, 110000045, PRien, 3);
        {$ENDIF}
        AfterChangeModule(105);
      end;
    {$IFNDEF EAGL}
    105134:
      begin
        EditEtat('E', 'GFA', GetParamSoc('SO_GCETATARTICLE'), TRUE, nil, '', '');
        RetourForce := TRUE;
      end;
    105135:
      begin
        EditEtat('E', 'RPF', GetParamSoc('SO_GCETATFICHETIERS'), TRUE, nil, '', '');
        RetourForce := TRUE;
      end;
    {$ENDIF}

    105201: AGLLanceFiche('MBO', 'EDTFOU', '', '', 'LISTE'); // Edition de la liste des fournisseurs
    105202: AGLLanceFiche('MBO', 'EDTFOU', '', '', 'ETIQ'); // Edition des étiquettes fournisseurs

    105304..105308: ParamTable('GCFAMILLENIV' + IntToStr(Num - 105300), taCreat, 0, PRien, 3, RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Num - 105300), FALSE));
      // familles

    105501..105502: ParamTable('GCSTATART' + IntToStr(Num - 105500), taCreat, 0, PRien, 9, RechDom('GCZONELIBRE', 'AS' + IntToStr(Num - 105500), False));
      // Statistiques

    105901..105905: FactoriseZL(Num, PRien);
    105906..105908: ParamTable('YYLIBRECON' + IntToStr(Num - 105905), taCreat, 0, PRien, 6);

    105909: FactoriseZL(Num, PRien);

    //   105999 : UtilTables_CtrlChamps;
    {$IFNDEF EAGL}
    // 106 - Administration
    106101: AGLLanceFiche('GC', 'GCSERIA_MODE', '', '', ''); // Etat de la sérialisation Mode
    106102:
      begin
        Assist_ImportGB2000;
        RetourForce := TRUE;
      end; // Import Mode GB 2000
    106103: AGLLanceFiche('GC', 'GCREPRISEDONNEES', '', '', ''); // Paramétrage reprise GPAO MODE
    106104:
      begin
        Assist_ImportGPAO;
        RetourForce := TRUE;
      end; // Import GPAO
    106105: EntreeStockAjustMODE; // vérification des stocks
    {$IFNDEF V530}
    106106: ParamFavoris('', '', False, False); // Gestion des favoris
    {$ENDIF}
    106107: MajPhonetiqueTiers;
    106108: AGLLanceFiche('GC', 'GCCPTADIFF', '', '', ''); // Compta différée
    106109: AGLLanceFiche('GC', 'GCCPTAPIECE', '', '', ''); // Regénération des écritures compta.
    106110: AGLLanceFiche('MBO', 'RECUPFICHIER', '', '', ''); //JD Download fichiers de MAJ PGI
    106111: AGLLanceFiche('MBO', 'PARAMCOMPTA', '', '', ''); //AC Fiche de test pour génération compta

    106201: ParamTraduc(TRUE, nil);
    106202: ParamTraduc(FALSE, nil);

    {   107410 :      // Situation Flash
                BEGIN
                if VerifEtatCaisse (Num, True) then
                   BEGIN
                   sCaisse := FOCaisseCourante ;
                   NumZ    := IntToStr(FOGetNumZCaisse(sCaisse)) ;
                   FOConsultationCaisse(sCaisse, NumZ, NumZ, PRien, 'LF1') ;
                   END else RetourForce := True ;
                END ;
    }
    110101: AglLanceFiche('MBO', 'ARTPHOTO', '', '', ''); // Création d'articles à partir de photos
    //////////////////////////////  Fin des menus de la Mode - Back-Office
    //////////////////////////////  Menus de la Mode - COMMUNS FO ET BO
    {$IFNDEF SANSTOX}
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
          AglToxConf(aceConfigure, '', nil, GescomModeTraitementTox, nil);
        end else
        begin
          AglToxConf(aceConfigure, '', GescomModeNotConfirmeTox, GescomModeTraitementTox, nil);
        end;
      end;
    111210: AglToxMulChronos; // Consultation des échanges
    111220: AfficheCorbeille; // Consultation des corbeilles
    111300: ToxSimulation; // Intégration d'un fichier
    111310: ConsultManuelTox; // Visualisation d'un fichier
    111320: DetopeTox; // Détopage d'un fichier
    {$ENDIF}

    {$ENDIF}
    //////////////////////////////  Fin des menus de la Mode - - COMMUNS FO ET BO
    ////////////////////////////// Menus de la Mode - Front-Office

       // 112 - Paramétres du Front-Office
    112120: ParamTable('GCCODEEVENT', taCreat, 0, PRien, 3, 'Evènements de la journée');
    112130: ParamTable('GCTYPEMETEO', taCreat, 0, PRien, 3, 'Météo');
    112140: ParamTable('GCTYPEREMISE', taCreat, 0, PRien, 3, 'Démarques');
    112180: AglLanceFiche('MFO', 'DETAILESPECES', V_PGI.DevisePivot + ';BIL', '', 'MODE=BO;DEVISE=' + V_PGI.DevisePivot);

    // Fin des menus de la Mode - Back-Office
    {$ENDIF}

    {$IFDEF NOMADE}
    ////////////////////////////// Menus POP - Prise d'Ordre par Portable

       // 168200 - Clients
    92411: AGLLanceFiche('RT', 'RTPROSPECTS_TV', '', '', '');
    92421: AGLLanceFiche('RT', 'RTTIERS_CUBE', '', '', '');

    // 102 - Ventes
    102101: AGLLanceFiche('MBO', 'MVTCLI_MODE', '', '', ''); // Liste des Pièces Clients/Articles
    102151: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'TYPECUB=VEN;TYPEART=ART');
    102152: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'TYPECUB=VEN;TYPEART=DIM');
    102110: AGLLanceFiche('MBO', 'VENTEVUE', '', '', 'TYPESTA=VEN;TYPEART=ART');
    102119: AGLLanceFiche('MBO', 'VENTEVUE', '', '', 'TYPESTA=VEN;TYPEART=DIM');
    102203: AGLLanceFiche('MBO', 'TARIF_MUL', '', '', ''); // Tarif Mode consultation
    102211: DispatchArtMode(7, '', '', 'ETATTARIF'); // Edition Tarifs Détail
    102161: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=DE', '', 'VENTE;CC'); // Gen.auto Devis --> Commande

    // DCA - Ajout du stock dans PCP Mode
    // 170 - Stocks
    103102: DispatchArtMode(3, '', '', ''); // Consultation - Disponible article
    103141: LanceStat(''); // Stats dispo article
    103142: LanceStat('DIM'); // Stats dispo article à la dimension
    103105: DispatchArtMode(5, '', '', ''); // Edition - Etat du stock par taille

    // 105 - Paramètres
    105101: LanceMulPrestation('ACTION=CONSULTATION'); // Données de base - Articles - Prestations
    105104: AGLLanceFiche('GC', 'GCFOUMUL_MODE', 'T_NATUREAUXI=FOU', '', 'ACTION=CONSULTATION');
    105107: AGLLanceFiche('MBO', 'EDTCLI_MODE', '', '', ''); // Edition des clients
    105116: DispatchArtMode(1, '', '', 'ARTICLE;ACTION=CONSULTATION'); // Mul Article Mode
    105129: DispatchArtMode(6, '', '', ''); // Article - Liste des articles
    105201: AGLLanceFiche('MBO', 'EDTFOU', '', '', 'LISTE'); // Edition de la liste des fournisseurs

    // 111 - Transmissions
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
          AglToxConf(aceConfigure, '', nil, GescomModeTraitementTox, nil);
        end else
        begin
          AglToxConf(aceConfigure, '', GescomModeNotConfirmeTox, GescomModeTraitementTox, nil);
        end;
      end;
    111210: AglToxMulChronos; // Consultation des échanges
    111220: AfficheCorbeille; // Consultation des corbeilles
    // DCA - Ajout des lignes de menus 111230/111240 Etat de consolidation/intégrations
    111230: AglLanceFiche('MBO', 'CONSOLIDATION', '', '', '');
    111240: LanceEtat('E','MST','IVE',True,False,False,Nil,'','',False);
    111300: ToxSimulation; // Intégration d'un fichier
    111310: ConsultManuelTox; // Visualisation d'un fichier
    111320: DetopeTox; // Détopage d'un fichier

    ////////////////////////////// Fin menus POP - Prise d'Ordre par Portable
    {$ENDIF}
  else if not TraiteMenuSpecif(Num) then
    HShowMessage('2;?caption?;' + TraduireMemoire('Fonction non disponible : ') + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;

end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range: string);
var ii: integer;
  Arg5LanceFiche: string;
begin
  if ((Num = 8) and (TT = 'GCTIERSSAISIE')) then
  begin
    ii := TTToNum(TT);
    if (ii > 0) and (pos('T_NATUREAUXI="FOU"', V_PGI.DECombos[ii].Where) > 0) then num := 12;
  end;
  case Num of
    {$IFNDEF GCGC}
    1: {Compte gene} FicheGene(nil, '', LeQuel, Action, 0);
    2: {Tiers compta} FicheTiers(nil, '', LeQuel, Action, 1);
    4: {Journal} FicheJournal(nil, '', Lequel, Action, 0);
    {$ENDIF}
    5:
      begin {Affaires}
        Arg5LanceFiche := ActionToString(Action);
        if (Range <> '') then Arg5LanceFiche := Arg5LanceFiche + ';' + Range;
        if (lequel = '') then lequel := 'AFF';
        AGLLanceFiche('AFF', 'AFFAIRE', '', Lequel, Arg5LanceFiche);
      end;
    6:
      begin {Ressources}
        Arg5LanceFiche := ActionToString(Action);
        if (Range <> '') then Arg5LanceFiche := Arg5LanceFiche + ';' + Range;
        AGLLanceFiche('AFF', 'RESSOURCE', '', Lequel, Arg5LanceFiche);
      end;
    7: {Article} DispatchTTArticle(Action, Lequel, TT, Range);
    8: {Clients par T_TIERS}
      begin
        Lequel := TiersAuxiliaire(Lequel, false);
        Arg5LanceFiche := ActionToString(Action) + ';MONOFICHE';
        if Range = '' then Arg5LanceFiche := Arg5LanceFiche + ';T_NATUREAUXI=CLI' else Arg5LanceFiche := Arg5LanceFiche + Range;
        AGLLanceFiche('GC', 'GCTIERS', '', Lequel, Arg5LanceFiche);
      end;
    9: {Commerciaux} AGLLanceFiche('GC', 'GCCOMMERCIAL', '', Lequel, ActionToString(Action) + ';GCL_TYPECOMMERCIAL="REP"');
    10: {Conditionnement} AGLLanceFiche('GC', 'GCCONDITIONNEMENT', Range, Lequel, ActionToString(Action));
    //11 : {Catalogue} AGLLAnceFiche('GC','GCCATALOGU_NVFOUR','',Lequel,ActionToString(Action)) ;
    11: {Catalogue} AGLLAnceFiche('GC', 'GCCATALOGU_SAISI3', '', Lequel, ActionToString(Action));
    12: {Fournisseurs via T_TIERS}
      begin
        Lequel := TiersAuxiliaire(Lequel, false);
        AGLLanceFiche('GC', 'GCFOURNISSEUR', '', Lequel, ActionToString(Action) + ';MONOFICHE;T_NATUREAUXI=FOU');
      end;
    13: {Dépots} AGLLanceFiche('GC', 'GCDEPOT', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    14: {Apporteurs} AGLLanceFiche('GC', 'GCCOMMERCIAL', '', Lequel, ActionToString(Action) + ';GCL_TYPECOMMERCIAL="APP"');
    15: {Règlement} FicheRegle_AGL('', False, taModif);
    16: {contacts} AGLLanceFiche('YY', 'YYCONTACT', Lequel, Range, ActionToString(Action) + ';MONOFICHE');
    17: {Fonction des ressources} AGLLanceFiche('AFF', 'FONCTION', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    18: {Etablissements}
      if (CtxMode in V_PGI.PGIContexte) then
        AGLLanceFiche('MBO', 'ETABLISS', '', Lequel, ActionToString(Action) + ';MONOFICHE')
      else
        AGLLanceFiche('GC', 'GCETABLISS', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    19: {Contenu tables libres} AGLLanceFiche('GC', 'GCCODESTATPIECE', '', '', 'ACTION=CREATION;YX_CODE=' + Range);
    20: ;
    21: {Ports} AGLLanceFiche('GC', 'GCPORT', '', Lequel, ActionToString(Action) + ';MONOFICHE');

    22: {GRC-action} AGLLanceFiche('RT', 'RTACTIONS', '', Lequel, ActionToString(Action) + ';MONOFICHE');
    23: {GRC-operation} AGLLanceFiche('RT', 'RTOPERATIONS', '', Lequel, ActionToString(Action) + ';MONOFICHE');

    28: {Clients par T_AUXILIAIRE}
      begin
        Arg5LanceFiche := ActionToString(Action) + ';MONOFICHE';
        if Range = '' then Arg5LanceFiche := Arg5LanceFiche + ';T_NATUREAUXI=CLI' else Arg5LanceFiche := Arg5LanceFiche + Range;
        AGLLanceFiche('GC', 'GCTIERS', '', Lequel, Arg5LanceFiche);
      end;
    29: {Fournisseurs via T_AUXILIAIRE} AGLLanceFiche('GC', 'GCFOURNISSEUR', '', Lequel, ActionToString(Action) + ';MONOFICHE;T_NATUREAUXI=FOU');
    30: {GRC-Projet} AGLLanceFiche('RT', 'RTPROJETS', '', Lequel, ActionToString(Action) + ';MONOFICHE');

    // *** maj des tablettes ***
    900: ParamTable('GCCOMMENTAIRELIGNE', taCreat, 0, nil, 6); // mis de manière explicite car appellé depuis un Grid

    // pour mise à jour par defaut des tablettes
    994: ParamTable(TT, taModif, 0, nil, 9); {choixext}
    995: ParamTable(TT, taCreat, 0, nil, 9); {choixext}
    996: ParamTable(TT, taModif, 0, nil, 6); {choixext}
    997: ParamTable(TT, taCreat, 0, nil, 6); {choixext}
    998: ParamTable(TT, taModif, 0, nil, 3); {choixcode ou commun}
    999: ParamTable(TT, taCreat, 0, nil, 3); {choixcode ou commun}
  end;
end;

procedure AfterChangeModule(NumModule: integer);
begin
  AjusteMenu(NumModule);
  TraiteMenusMode(NumModule);
  {$IFDEF GRC}
  RTMenuTraiteMenusGRC(NumModule);
  {$ENDIF}
  TraiteMenusGescom(NumModule);
  TraiteMenusAffaire(NumModule);
  TraiteSpecifsCEGID(NumModule);
  TraiteMenusEAGL(NumModule);
  {$IFDEF EAGLCLIENT}
  {$IFDEF AGL550B}
  case NumModule of
    30, 31, 32, 33, 34, 65, 70: ChargeMenuPop(5, FMenuG.DispatchX);
    92: ChargeMenuPop(8, FMenuG.DispatchX);
  else ChargeMenuPop(NumModule, FMenuG.DispatchX);
  end;
  {$ELSE}
  case NumModule of
    30, 31, 32, 33, 34, 65, 70: ChargeMenuPop(hm5, FMenuG.DispatchX);
    92: ChargeMenuPop(hm8, FMenuG.DispatchX);
  else ChargeMenuPop(TTypeMenu(NumModule), FMenuG.DispatchX);
  end;
  {$ENDIF}
  {$ELSE}
  {$IFDEF AGL545}
  {$IFDEF AGL550B}
  case NumModule of
    //30, 31, 32, 33, 34, 65, 70: ChargeMenuPop(5, FMenuG.DispatchX);
    //92: ChargeMenuPop(8, FMenuG.DispatchX);
    30,31,32 : ChargeMenuPop(26,FMenuG.DispatchX) ;
    167,168,170 : ChargeMenuPop(27,FMenuG.DispatchX) ;
    169,171 : if ctxMode in V_PGI.PGIContexte
                then ChargeMenuPop(27,FMenuG.DispatchX)
                else ChargeMenuPop(26,FMenuG.DispatchX) ;
  else ChargeMenuPop(NumModule, FMenuG.DispatchX);
  end;
  {$ELSE}
  case NumModule of
    30, 31, 32, 33, 34, 65, 70: ChargeMenuPop(hm5, FMenuG.DispatchX);
    92: ChargeMenuPop(hm8, FMenuG.DispatchX);
  else ChargeMenuPop(TTypeMenu(NumModule), FMenuG.DispatchX);
  end;
  {$ENDIF}
  {$ELSE}
  case NumModule of
    30, 31, 32, 33, 34, 65, 70: ChargeMenuPop(hm5, FMenuG.Dispatch);
    92: ChargeMenuPop(hm8, FMenuG.Dispatch);
  else ChargeMenuPop(TTypeMenu(NumModule), FMenuG.Dispatch);
  end;
  {$ENDIF}
  {$ENDIF}

  {$IFDEF NOMADE}
  if not (ctxMode in V_PGI.PGIContexte) then TraiteMenusNomade(NumModule);
  {$ENDIF} // NOMADE

end;

///////////////////////////////////////////////////////////////////////////////////////
//  AfterProtec
///////////////////////////////////////////////////////////////////////////////////////

procedure AfterProtec(sAcces: string);
var PCPSeria: boolean;
begin
  PCPSeria := False;
  if Length(sAcces) > 0 then PCPSeria := (sAcces[1] = 'X'); // PCP sérialisé

  if V_PGI.NoProtec then
  begin
    V_PGI.VersionDemo := True;
    PCPSeria := False;
  end
  else V_PGI.VersionDemo := False;

  if (PCPSeria = False) then V_PGI.VersionDemo := True; // on repasse en demo dans ce cas

  if (V_PGI.VersionDemo) and (Pos(' (DEMO)', TitreHalley) = 0) then TitreHalley := TitreHalley + ' (DEMO)';
end;

procedure InitApplication;
begin
  ProcZoomEdt := GCZoomEdtEtat;
  {$IFDEF GRC}
  RTMenuInitApplication;
  ProcCalcMul := RTProcCalcMul;
  {$ENDIF}
  ProcCalcEdt := GCCalcOLEEtat;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  {$IFDEF EAGLCLIENT}
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
  {$ELSE}
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
  FMenuG.OnMajPendant := nil;
  {$ENDIF}
  FMenuG.OnAfterProtec := AfterProtecGC;
  {$IFDEF CEGID}
  AfterProtecGC(''); // Pas de seria CEGID
  {$ELSE}
  FMenuG.SetSeria(GCCodeDomaine, GCCodesSeria, GCTitresSeria);
  {$ENDIF}
  FMenuG.OnChangeModule := AfterChangeModule;
  V_PGI.DispatchTT := DispatchTT;
  {$IFNDEF MODE}
  FMenuG.SetModules([60], [49]);
  V_PGI.NbColModuleButtons := 1;
  V_PGI.NbRowModuleButtons := 6;
  {$ENDIF}
  {$IFNDEF EAGLCLIENT}
  FMenuG.SetPreferences(['Pièces'], False);
  {$ENDIF}
  {$IFNDEF V530}
  FMenuG.OnChargeFavoris := ChargeFavoris;
  {$ENDIF}
end;

{$IFNDEF AGL550B}
initialization
  InitLaVariablePGI;
  {$ENDIF}

end.
