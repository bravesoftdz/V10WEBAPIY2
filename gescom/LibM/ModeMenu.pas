unit ModeMenu;

interface

uses Windows, Forms, HMsgBox, Classes, Hctrls, Controls, HPanel, UIUtil,
  ImgList, ParamSoc,
  {$IFDEF EAGLCLIENT}
  Maineagl, MenuOLX, M3FP, UtileAGL, EntGC, eTablette, AGLInit,
  {$ELSE}
  FE_Main, MenuOLG,
  AGLInit, tablette, pays, region, codepost, UtomArticle,
  EdtQR, CreerSoc, CopySoc, Reseau, Mullog, LicUtil, FactUtil, EdtEtat,
  EdtREtat, UEdtComp, devise, AssistInitSoc,
  {$ENDIF}
  GrpFavoris, UtilDispGC, UtilArticle, HEnt1, sysutils,
  UtomUtilisat, UtilPGI, CPTACORRESP_TOF, Confidentialite_TOF;

procedure MODEMenuDispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
procedure MODEMenuDispatchTT(Num: integer; Action: TActionFiche; Lequel, TT, Range: string);

implementation

uses
  Facture, Transfert, TiersUtil, FOUtil,
  {$IFDEF EAGLCLIENT}
  GcTraitCli, TarifArticleMode, TarifTiersMode, AssistRecopiTarf,
  ArtPrestation_Tof,
  {$ELSE}
  UtilGC, SaisUtil,
  GcTraitCli, TarifArticleMode, TarifTiersMode, AssistRecopiTarf, ETransferts,
  {$IFNDEF SANSTOX}
  uToxConst, uToxFiches, GCTOXWORK, ToxSim, ToxDetop, uToxConf, ToxConsultBis, GCToxCtrl,
  GenereBE, ToxMoteur, RecupFichier_Tof,
  {$ENDIF}
  AssistImportGPAO, AssistImportGB2000,
  Traduc, TradMenu, AssistStockAjust, ArtPrestation_Tof, CubeDim_Tof, StatDim_Tof,
  {$ENDIF}
  MBOEXPORTPIECE_TOF, MBOEPURATIONALF_TOF, EXPORTDISPO_TOF, AssistFinInv;

procedure MODEMenuDispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
begin
  case Num of
////////////////////////////// Menus E-Commerce
   {Fonctions E-COMMERCE}
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

   //59204 : AFLanceFiche_Mul_EActivite (''); // import E-Activité
   //59205 : AFLanceFiche_Mul_ExportAff; // remplacé par 74654 en GI, à supprimer qd n'existe plus en GA export par affaires
   //59206 : AFLanceFiche_MulEAffaire; // import E_affaires
   59207 : AGLLanceFiche('E','EIMPORTTIERS','','','');      // import E-Tiers
   59208 : AGLLanceFiche('E','EIMPORTCMD','','','');        // import E-Commandes
   //59209 : AFLanceFiche_Suivi_EActivite; //suivi saisie décentralisée
   //
   59301 : AGLLanceFiche('E','EEXPORTARTICLE','','','');
   59302 : AGLLanceFiche('E','EEXPORTTIERS','','','');
   59303 : AGLLanceFiche('E','EEXPORTTARIF','','','');
   59304 : AGLLanceFiche('E','EEXPORTCOMM','','','');
   59305 : AGLLanceFiche('E','EEXPORTPAYS','','','');
   59401 : AGLLanceFiche('E','ETRADTABLETTE','','','');
   {$ENDIF}

    ////////////////////////////// Menus de la Mode - Back-Office

    // 101 - Achats
    101121: AGLLanceFiche('MBO', 'ACHATCUB', '', '', 'TYPECUB=ACH;TYPEART=ART');
    101122: AGLLanceFiche('MBO', 'ACHATCUB', '', '', 'TYPECUB=ACH;TYPEART=DIM');
    101102: CreerPiece('ALF');
    101103: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=ALF', '', 'MODIFICATION');
    101104: AGLLanceFiche('GC', 'GCTRANSACH_MUL', 'GP_NATUREPIECEG=CF', '', 'ALF');
    101106: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=BLF;GP_VENTEACHAT=ACH', '', 'CONSULTATION');
    101107: AGLLanceFiche('GC', 'GCLIGNEACH_MUL', 'GL_NATUREPIECEG=BLF', '', 'CONSULTATION');
    101108: AGLLanceFiche('MBO', 'ECARTALFBLF', 'NATURE=CF', '', 'CF'); // Edition des écarts entre commande fournisseur et réception
    101109: AGLLanceFiche('MBO', 'ECARTALFBLF', 'NATURE=ALF', '', 'ALF'); // Edition des écarts entre annonce de livraison et réception
    101110: AGLLanceFiche('GC', 'GCPIECEACH_MUL', 'GP_NATUREPIECEG=FF', '', 'MODIFICATION');
    101111: AGLLanceFiche('GC', 'GCDUPLICPIECE_MUL', 'GP_NATUREPIECEG=ALF', '', 'ALF');
    101112: LanceExportPiece('EXPORTPIECEAC_MUL', 'ALF', 'PAC'); // Export ASCII des pièces d'achat
    101113: LanceEpurationAnnonceDeLivraison;
    // génération automatique
    101114: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=FCF', '', 'ACHAT;CF'); // Commande Réassort => CF
    101131: AGLLanceFiche('MBO', 'ACHATVUE', '', '', 'TYPESTA=ACH;TYPEART=ART');
    101132: AGLLanceFiche('MBO', 'ACHATVUE', '', '', 'TYPESTA=ACH;TYPEART=DIM');
    // Tarif d'achat
    101201: EntreeTarifArticleMode(taModif);
    101202: AGLLanceFiche('MBO', 'TARIF_MUL', '', '', 'TYPE=ACH'); // Tarif Mode consultation
    101203: AGLLanceFiche('MBO', 'TARIFTYPE', '', '', 'TYPE=ACH');
    101204: DispatchArtMode(4, '', '', 'GCARTICLE;TYPE=ACH'); // Tarif détail - Maj tarifs mode
    101205: DispatchArtMode(7, '', '', 'ETAT=ETATTARIF;TYPE=ACH'); // Tarifs détail - Edition - tarifs
    101206: AGLLanceFiche('MBO', 'TARIF_SUP', '', '', 'TYPE=ACH'); // Tarif Mode suppression

    // 102 - Ventes
    102101: AGLLanceFiche('MBO', 'MVTCLI_MODE', '', '', ''); // Liste des Pièces Clients/Articles
    102102: AGLLanceFiche('GC', 'GCPIEDECHET_MUL', '', '', 'MODIFICATION'); // Modif des règlements
    102103: CreerPiece('FFO'); // Saisie des Ventes F.O.
    102104: AGLLanceFiche('GC', 'GCTICKET_MUL', '', '', 'MODIFICATION'); //Modif des tickets F.O.
    102105: AGLLanceFiche('GC', 'GCEDTCRV_MODE', '', '', ''); // Compte-rendu des ventes détail
    102106: AGLLanceFiche('GC', 'GCEDTJCA_MODE', '', '', ''); // Journal de caisse
    102151: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'TYPECUB=VEN;TYPEART=ART');
    102152: AGLLanceFiche('MBO', 'VENTECUB', '', '', 'TYPECUB=VEN;TYPEART=DIM');
    102108: AGLLanceFiche('GC', 'GCTICKET_MUL', '', '', 'CONSULTATION'); //Consultation des tickets F.O.
    102109: AGLLanceFiche('GC', 'GCLIGTICKET_MUL', 'GL_TYPEARTICLE=MAR', '', 'CONSULTATION'); // Consultation par lignes de ventes FFO
    102110: AGLLanceFiche('MBO', 'VENTEVUE', '', '', 'TYPESTA=VEN;TYPEART=ART');
    102111: AGLLanceFiche('GC', 'GCEDSYVTEA', 'TYPSYN=AD', 'AND', 'AD');
    102112: AGLLanceFiche('GC', 'GCEDSYVTEA', 'TYPSYN=AA', 'ANN', 'AA');
    102113: AGLLanceFiche('GC', 'GCGRAPHSYNA', '', '', 'AD');
    102114: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=FAC', '', 'MODIFICATION'); //Modification Factures
    102115: AGLLanceFiche('MBO', 'ANALREG', 'GPE_NATUREPIECEG=FFO', '', '');
    102116: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=AVC', '', 'MODIFICATION'); //Modification Avoirs client
    102118: AGLLanceFiche('GC', 'GCPIECE_MUL', 'GP_NATUREPIECEG=AVS', '', 'MODIFICATION'); //Modification Avoirs sur Stock
    102119: AGLLanceFiche('MBO', 'VENTEVUE', '', '', 'TYPESTA=VEN;TYPEART=DIM');
    102140: AGLLanceFiche('MBO', 'SYNVENTES', '', '', '');
    102121: AGLLanceFiche('GC', 'GCEDSYVTEM', 'TYPSYN=MD', 'MOD', 'MD');
    102122: AGLLanceFiche('GC', 'GCEDSYVTEM', 'TYPSYN=MA', 'MON', 'MA');
    102123: AGLLanceFiche('GC', 'GCGRAPHSYNM', '', '', 'MD');
    102131: AGLLanceFiche('GC', 'GCGRAPHVIEPROD', '', '', '');
    102141: AGLLanceFiche('GC', 'GCEDTETVTE', '', '', '');
    102142: AGLLanceFiche('GC', 'GCGRAPHETVTE', '', '', '');
    102117: LanceExportPiece('EXPORTPIECEVT_MUL', 'FAC', 'PVT'); // Export ASCII des pièces de vente
    102161: AGLLanceFiche('GC', 'GCGROUPEPIECE_MUL', 'GP_NATUREPIECEG=DE', '', 'VENTE;CC'); // Devis --> Commande

    102201: EntreeTarifArticleMode(taModif, TRUE);
    102202: EntreeTarifTiersMode(taModif, TRUE);
    102203: AGLLanceFiche('MBO', 'TARIF_MUL', '', '', 'TYPE=VTE'); // Tarif Mode consultation
    102204: AGLLanceFiche('MBO', 'TARIFPER', '', '', '');
    102205: AGLLanceFiche('MBO', 'TARIFTYPE', '', '', 'TYPE=VTE');
    102206: RecopieTarifMode;
    102207: DispatchArtMode(4, '', '', 'GCARTICLE;TYPE=VTE'); // Tarif détail - Maj tarifs mode
    102211: DispatchArtMode(7, '', '', 'ETAT=ETATTARIF;TYPE=VTE'); // Tarifs détail - Edition - tarifs
    102212: DispatchArtMode(7, '', '', 'ETAT=COMPTARIF;TYPE=VTE'); // Tarifs détail - Edition - comparatif
    102213: AGLLanceFiche('MBO', 'PANIER_MOYEN', '', '', ''); // Panier Moyen
    102214: AGLLanceFiche('MBO', 'ARTNONVENDUS', '', '', ''); // Articles non vendus
    102215: AGLLanceFiche('MBO', 'TARIF_SUP', '', '', 'TYPE=VTE'); // Tarif Mode suppression

    // Affectation des commandes clients
    102221: AGLLanceFiche('MBO', 'AFFCDESELECT', '', '', ''); // sélection des lignes de commandes
    102222: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'RESERVATION'); // réservation
    102223: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'AFFECTATION'); // affectation
    102224: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'PREPARATION'); // préparation
    102225: AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'FIN'); // fin de l'affectation
    102226: ;
    102227: AGLLanceFiche('MBO', 'AFFCDEMODIF', '', '', 'AFFECTATION'); // Modififcation des affectations
    102229: AGLLanceFiche('MBO', 'AFFCDEENTETE', '', '', ''); // paramètres d'affectation

    // 103 - Stocks
    103101: AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULDISTANT'); //Consultation du stock à distance
    103102: DispatchArtMode(3, '', '', ''); // Consultation - Disponible article
    103103: AGLLanceFiche('MBO', 'DISPODIST_MUL', '', '', 'MULLOCAL'); //Consultation du stock local
    103104: ;
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
    103131: LanceExportPiece('EXPORTPIECE_MUL', 'TEM', 'TID'); // Export ASCII des transferts
    103132: CreerTransfert('TEM', 'TEM_TRV_NUMIDENTIQUE'); // Saisie des Transferts Inter-Boutiques
    103133: AGLLanceFiche('GC', 'GCMOUVSTKEX_MUL', 'GP_NATUREPIECEG=EEX', '', 'MODIFICATION'); // Modification des mouvements exceptionnels
    // -103140 entête du sous-menu des stat
    {$IFNDEF EAGLCLIENT}
    103141: LanceStat(''); // Stats dispo article
    103142: LanceStat('DIM'); // Stats dispo article à la dimension
    {$ENDIF}
    // -103145 entête du sous-menu des stat
    103146: AGLLanceFiche('MBO', 'STOCKCUB', '', '', 'ART'); //DispatchArtMode(9,'','','') ;  // Stocks - Consultation - Cube des stocks
    103147: AGLLanceFiche('MBO', 'STOCKCUB', '', '', 'DIM'); // Stats vente (cube) à la taille
    103148: LanceExportDispo('DIS');

    103150: AGLLanceFiche('GC', 'GCREPRISEDONNEES', 'SAISIE_INVENTAIRE=X', '', ''); // Paramétrage du fichier inventaire
    {$IFNDEF EAGLCLIENT}
    103151:
      begin
        Assist_ImportGPAO(True);
        RetourForce := TRUE;
      end; // Reprise du fichier inventaire
    {$ENDIF}
    // 105 - Paramètres
    105101: LanceMulPrestation('ACTION=MODIFICATION'); // Données de base - Articles - Prestations
    105102: LanceMulArtFinancier('ACTION=MODIFICATION'); // Paramètres - Gestion - Opérations de
    105103: ParamTable('TTLIENPARENT', taCreat, 0, PRien, 3, 'Lien de parenté');
    105104: AGLLanceFiche('GC', 'GCFOUMUL_MODE', 'T_NATUREAUXI=FOU', '', '');
    105105: AGLLanceFiche('YY', 'YYCONTACTS_MUL', '', '', 'MODIFICATION');
    105106: AglLanceFiche('GC', 'GCVENDEUR', 'VEN', '', ''); // Vendeurs
    105107: AGLLanceFiche('MBO', 'EDTCLI_MODE', '', '', ''); // Edition des clients
    105108: AGLLanceFiche('MBO', 'EDTCPTCLI_MODE', '', '', ''); // Compte client
    105109: TraitClientSerie; // formatage des zones clients
    105110:
      if (not ExisteSQL('Select * from CODEBARRES where (GCB_NATURECAB="SO") and (GCB_IDENTIFCAB="...")'))
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
    105130: AGLLanceFiche('MBO', 'VENTILANA', '', '', ''); // Ventil ana par mode
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
    105136: AGLLanceFiche('MBO', 'ARTICLE_MODIFLOT', '', '', '');
    {$IFNDEF EAGL}
    105137:
      begin
        EditEtatS5S7('E', 'UFO', '', True, nil, '', '');
        RetourForce := True
      end; // Editeur d'états libres FO utilisateurs
    105138:
      begin
        LanceEtatLibreGC('UFO');
        RetourForce := True
      end; // Lanceur d'états libres FO utilisateurs
    {$ENDIF}
    // type et côte emplacement se trouvant dans le module dépôts
    105139: ParamTable('GCTYPEEMPLACEMENT', taCreat, 0, PRien); // type emplacement
    105140: ParamTable('GCCOTEEMPLACEMENT', taCreat, 0, PRien); // cote emplacement
    105141: ParamTable('GCCATEGORIETAXE', taModif, 110000045, PRien, 3);

    {$IFNDEF EAGL}
    105151:
      begin
        EditEtat('E', 'MST', 'CUA', False, nil, '', '');
        RetourForce := True
      end;
    105152:
      begin
        EditEtat('E', 'MST', 'TVA', False, nil, '', '');
        RetourForce := True
      end;
    105153:
      begin
        EditEtat('E', 'MST', 'CUV', False, nil, '', '');
        RetourForce := True
      end;
    105154:
      begin
        EditEtat('E', 'MST', 'TVV', False, nil, '', '');
        RetourForce := True
      end;
    105155:
      begin
        EditEtat('E', 'MST', 'CUS', False, nil, '', '');
        RetourForce := True
      end;
    105156:
      begin
        EditEtat('E', 'MST', 'TVS', False, nil, '', '');
        RetourForce := True
      end;
    {$ENDIF}

    // Menu Export
    105161 : ParamTable('GCMODEEXP' , taCreat, 0, PRien, 3, 'Modes d''expéditions'); { GPAO1_INCOTERM Export / Mode d'expédition }
    105162 : ParamTable('GCINCOTERM', taCreat, 0, PRien, 3, 'Incoterms');            { GPAO1_INCOTERM Export / Incoterm }

    105201: AGLLanceFiche('MBO', 'EDTFOU', '', '', 'LISTE'); // Edition de la liste des fournisseurs
    105202: AGLLanceFiche('MBO', 'EDTFOU', '', '', 'ETIQ'); // Edition des étiquettes fournisseurs

    105304..105308: ParamTable('GCFAMILLENIV' + IntToStr(Num - 105300), taCreat, 0, PRien, 3, RechDom('GCLIBFAMILLE', 'LF' + IntToStr(Num - 105300), FALSE));
      // familles

    105500: ParamTable('GCCOLLECTION', taCreat, 0, PRien, 3, RechDom('GCZONELIBRE', 'AS' + IntToStr(Num - 105500), False)); // collection
    105501..105502: ParamTable('GCSTATART' + IntToStr(Num - 105500), taCreat, 0, PRien, 3, RechDom('GCZONELIBRE', 'AS' + IntToStr(Num - 105500), False));
      // Statistiques

    105906..105908: ParamTable('YYLIBRECON' + IntToStr(Num - 105905), taCreat, 0, PRien, 6, RechDom('GCZONELIBRE', 'BT' + IntToStr(Num - 1105905), False));
      // Statistiques) ;

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
    106105: EntreeStockAjust; // vérification des stocks
    106106: ParamFavoris('', '', False, False); // Gestion des favoris
    106107: MajPhonetiqueTiers;
    106108: AGLLanceFiche('GC', 'GCCPTADIFF', '', '', ''); // Compta différée
    106109: AGLLanceFiche('GC', 'GCCPTAPIECE', '', '', ''); // Regénération des écritures compta.
    106110: AGLLanceFiche('MBO', 'RECUPFICHIER', '', '', ''); //JD Download fichiers de MAJ PGI
    106111: AGLLanceFiche('MBO', 'TRANSFERTCOMPTA', '', '', ''); //AC Pour génération compta
    106112: AGLLanceFiche('MBO', 'RECALCULPIECE_MUL', 'GP_NATUREPIECEG=FFO', '', ''); // Recalcul des pièces
    //106113 : AppelCalculStock();   // MODIF LM pour compile sans CalculStock
    106115: GCLanceFiche_CorrespCompta('GC', 'GCCPTACORRESP', '', '', 'MODEREGL;MDR;MR_MODEREGLE;MR_LIBELLE'); //Saisie correspondance compta sur MODEREGL
    106116: GCLanceFiche_CorrespCompta('GC', 'GCCPTACORRESP', '', '', 'MODEPAIE;MDP;MP_MODEPAIE;MP_LIBELLE'); //Saisie correspondance compta sur MODEPAIE
    106117: // Gestion droits d'accès
      GCLanceFiche_Confidentialite( 'YY', 'YYCONFIDENTIALITE', '', '', '26;101;102;103;105;106;107;108;109;110;111;112;115;27' );

    106201: ParamTraduc(TRUE, nil);
    106202: ParamTraduc(FALSE, nil);

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
    111330: LanceGenereBE; // Génération des BE
    111340: LanceMiniFTP; // Mini FTP pour récupérer des fichiers sur le central
    {$ENDIF}
    {$ENDIF}

    110101: AglLanceFiche('MBO', 'ARTPHOTO', '', '', ''); // Création d'articles à partir de photos
    // Fidélité
    110010 : AGLLanceFiche('MBO','MBOREGLEFID','','','') ;
    110011 : AGLLanceFiche('MBO','FIDELITEENT','','','') ;

    ////////////////////////////// Menus de la Mode - Front-Office

    // 112 - Paramétres du Front-Office
    112110: AGLLanceFiche('MFO', 'PCAISSE', '', '', ActionToString(taModif)); // Paramétrage des caisses
    112115: AGLLanceFiche('MFO', 'MODIFCAIS_MUL', '', '', ActionToString(taModif)); // Modification en série du paramétrage des caisses
    112120: ParamTable('GCCODEEVENT', taCreat, 0, PRien, 3, 'Evènements de la journée');
    112130: ParamTable('GCTYPEMETEO', taCreat, 0, PRien, 3, 'Météo');
    112140: AGLLancefiche('MFO', 'DEMARQUE', '', '', '');
    112180: AglLanceFiche('MFO', 'DETAILESPECES', V_PGI.DevisePivot + ';BIL', '', 'MODE=BO;DEVISE=' + V_PGI.DevisePivot);
  end;
end;

procedure MODEMenuDispatchTT(Num: integer; Action: TActionFiche; Lequel, TT, Range: string);
begin
  case Num of
    112200: {Lancement de l'éditeur de ticket} FODispatchParamModele(Num, Action, Lequel, TT, Range);
  end;
end;

end.
