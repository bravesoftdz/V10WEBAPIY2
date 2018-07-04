unit MDispGCS3;

interface

Uses Windows,Forms,HMsgBox, Classes, Hctrls, Controls, HPanel, UIUtil,ImgList,
{$IFDEF EAGLCLIENT}
     Maineagl,MenuOLX,M3FP,UtileAGL,eTablette,
{$ELSE}
     FE_Main,MenuOLG,EdtEtat,
{$IFNDEF V530}
     EdtREtat,UedtComp,
{$ENDIF}
     Tablette,ParamSoc,UtomArticle,UtilArticle,pays,region,codepost,
     EdtQR,CreerSoc,CopySoc,Reseau,Mullog,LicUtil,FactUtil,
  {$IFNDEF SANSCOMPTA}
     devise,
  {$ENDIF}
{$ENDIF}
     UtilDispGC,AGLInit,HEnt1,sysutils,
      UtomUtilisat,UtilPGI,UserGrp_tom,UtilRT,UtilRCC,AssistInitSoc,tradarticle  ;

Procedure InitApplication ;
procedure AfterChangeModule ( NumModule : integer ) ;

type
  TFMenuDisp = class(TDatamodule)
    end ;

Var FMenuDisp : TFMenuDisp ;

implementation

{$IFDEF EAGLCLIENT}
Uses CalcOleGescom,UtilGC,Facture,TarifArticle,TarifTiers,TarifCliArt,TarifCatArt,
    {$IFDEF GRC}
     ZoomEdtProspect,
    {$ENDIF}
 {$IFDEF GESCOM}
    MenuCaisse ,
 {$ENDIF}
     Ent1,AssistSuggestion,TiersUtil,FichComm,EntGC, ZoomEdtGescom,
     HistorisePiece_Tof,ListeInvContrem_TOF,InvContrem_TOF, UTOFGCLISTEINV_MUL,
     ListeInvPreCon_TOF,UTOFListeInvVal,MouvStkExContr_TOF ;
{$ELSE}
uses Ent1,
     UTomPiece, EntGC, UtilGC,
     FichCOmm,ZoomEdtGescom ,CalcOLEGescom,
      EdtDoc,UtofConsoEtat,Transfert,AssistDEBC,
     UtilSoc,AssistSuggestion,TiersUtil,GCOLEAUTO,
 {$IFNDEF SANSCOMPTA}
     Souche,Suprauxi, Tva, OuvrFerm, InitSoc,
 {$ENDIF}
 {$IFNDEF SANSPARAM}
     dimension,UtomModePaie,Arrondi,ListeDeSaisie,
     TarifArticle, TarifCliArt, TarifTiers, TarifCatArt,
     AssistStockAjust,eTransferts,
 {$ENDIF}
 {$IFNDEF GCGC}
     MulGene, General, Tiers, Journal,TVA,OuvrFerm,
     EtbMce,EtbUser,TeleTrans,banquecp,
 {$ENDIF}
 {$IFDEF AFFAIRE}
     AFActivite,ActiviteUtil,ConfidentAffaire,   AssistCodeAffaire,
 {$ENDIF}
 {$IFDEF GRC}
     ZoomEdtProspect,
 {$ENDIF}
 {$IFDEF GESCOM}
    MenuCaisse ,
 {$ENDIF}
    Facture,HistorisePiece_Tof,PieceEpure_Tof,ListeInvContrem_TOF,InvContrem_TOF,
    UTOFGCLISTEINV_MUL,ListeInvPreCon_TOF,UTOFListeInvVal,MouvStkExContr_TOF ;
{$ENDIF}


{$R *.DFM}

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
   65525 : St:='MT' ; 65526 : St:='MM' ; 65527 : St:='MD' ; 65528 : St:='MC' ; 65529 : St:='MB' ;
   65535 : St:='RT' ; 65536 : St:='RM' ; 65537 : St:='RD' ; 65538 : St:='RC' ; 65539 : St:='RB' ;
   65586,105901 : St:='BT' ; 65587,105902 : St:='BM' ; 65588,105903 : St:='BD' ;
   65589,105904 : St:='BC' ; 65590,105905 : St:='BB' ;
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
   'B' :  iHelpContext:=0;                //Contacts
   End;
GCModifTitreZonesLibres(St) ;
ParamTable('GCZONELIBRE',taModif,iHelpContext,PRien,3) ;
GCModifTitreZonesLibres('') ; AfterChangeModule(65) ;
END ;

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
Var StVire,StSeule,StParPiece : String ;
BEGIN
   Case Num of
   10 : begin UpdateCombosGC; GCModifLibelles; GCTripoteStatus ; end;       //  après connexion
   11 : ; //  après déconnexion
   12 : ; //  avant connexion et séria
{$IFDEF EAGLCLIENT}
   13 : begin
       {$IFDEF GESCOM}
          LiberationCaisse;
       {$ENDIF}
       {$IFDEF AGL550B}
          ChargeMenuPop(-1,FMenuG.DispatchX) ;
       {$ELSE}
          ChargeMenuPop(TTypeMenu(-1),FMenuG.DispatchX) ;
       {$ENDIF}
        end;
{$ELSE}
   {$IFDEF AGL545}
   13 : begin
       {$IFDEF GESCOM}
          LiberationCaisse;
       {$ENDIF}
       {$IFDEF AGL550B}
          ChargeMenuPop(-1,FMenuG.DispatchX) ;
       {$ELSE}
          ChargeMenuPop(TTypeMenu(-1),FMenuG.DispatchX) ;
       {$ENDIF}
        end;
   {$ELSE}
   13 : begin
        {$IFDEF GESCOM}
          LiberationCaisse;
        {$ENDIF}
          ChargeMenuPop(TTypeMenu(-1),FMenuG.Dispatch) ;
        end;
   {$ENDIF}
{$ENDIF}
   15 : ;
   100 : ;  // Si lanceur   
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
{$IFNDEF EAGLCLIENT}
   30212 : GCLanceFiche_PieceEpure('GC','GCPIECEEPURE_MUL','','','VEN') ; //Epuration des pièces
{$ENDIF}
   // Consultation
   30301 : AGLLanceFiche('GC','GCPIECE_MUL','GP_VENTEACHAT=VEN;GP_NATUREPIECEG=FAC','','CONSULTATION') ;
   30302 : AGLLanceFiche('GC','GCLIGNE_MUL','GL_NATUREPIECEG=FAC','','CONSULTATION') ;
   // Génération unitaire pièces
   30401 : AGLLanceFiche('GC','GCPIECEVISA_MUL','GP_NATUREPIECEG=CC;GP_VENTEACHAT=VEN','','') ; // Visa des pièces
   30403 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=DE','','CC') ;   // Commandes
   30409 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=CC','','PRE') ;   // Prépas
   30421 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=PRE','','BLC') ;  // Livraisons
   30422 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=CCE','','LCE') ;  // Livraisons échantillon
   30405 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=BLC','','FAC') ;  // Factures
   // Génération manuelle pièces
   30431 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=CC','','CCR;BLC') ;  // Comm --> Livr
   30432 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=BLC','','LCR;FAC') ; // Livr --> Fact
   // Génération automatique pièces
   30413 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','PRE');  // Com --> Prep
   30415 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','BLC');  // Com --> Livr
   30411 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=PRE','','BLC') ; // Prep --> Livr
   30412 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLC','','FAC') ; // Livr --> Fact
   30414 : AGLLanceFiche('AFF','AFPIECEPRO_MUL','','','TABLE:GP;NATURE:FPR;NOCHANGE_NATUREPIECE');  //validation factures provisoires
   30416 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CC','','VENTE;FAC');  // Com --> Fact
   30407 : AGLLanceFiche('AFF','AFPREPFACT_MUL','','','');   // Prépa facture par affaires
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
   30109 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','SUP') ; // suppression de tarif
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
   // fichiers
   30501 : if ctxMode in V_PGI.PGIContexte
              then AGLLanceFiche('GC','GCCLIMUL_MODE','T_NATUREAUXI=CLI','','')
              else AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=CLI','','') ;    // Clients
   30502 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=CLI','','') ; // contacts
   30503 : AGLLanceFiche ('GC','GCCOMMERCIAL_MUL','','','') ; // Commerciaux
   // GRC
   //30504 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','CLI') ;  // modification en série des tiers
   30504 : AGLLanceFiche('RT','RTQUALITE','','','CLI;GC') ;  // modification en série des tiers
{$IFNDEF SANSCOMPTA}
   30507 : SuppressionCpteAuxi('CLI') ;
{$ENDIF}
   30506 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE') ;  // ouverture compte tiers
   30505 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME') ;  // fermeture compte tiers

{$IFDEF GESCOM}
///////  Ventes Comptoir  /////////////
   36100..36999 : MenuDispatchCaisse(Num);
{$ENDIF}

///////  Achats  /////////////
   // Saisie pièces
   31201 : CreerPiece('CF') ;
   31202 : CreerPiece('BLF') ;
   31203 : CreerPiece('FF') ;
   // Modifications
   31211 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=CF','','MODIFICATION')  ;  //commandes fournisseur
   31212 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=BLF','','MODIFICATION') ;  //bons de livraison fournisseur
   // Duplication unitaire pièces
   31221 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=CF','','CF') ;
   31222 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=BLF','','BLF') ;
   31223 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=FF','','FF') ;
   31204 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','ACH') ; //Histirisation des pièces
{$IFNDEF EAGLCLIENT}
   31205 : GCLanceFiche_PieceEpure('GC','GCPIECZEPURE_MUL','','','ACH') ; //Epuration des pièces
{$ENDIF}
   // Consultation
   31101 : AGLLanceFiche('GC','GCPIECEACH_MUL','GP_NATUREPIECEG=FF;GP_VENTEACHAT=ACH','','CONSULTATION') ;
   31102 : AGLLanceFiche('GC','GCLIGNEACH_MUL','GL_NATUREPIECEG=FF','','CONSULTATION') ;
   // génération
   31501 : AGLLanceFiche('GC','GCPIECEVISA_MUL','GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH','','') ; // Visa des pièces
   31513 : AGLLanceFiche('GC','GCCONTREMVTOA','','','MODE=VTOA') ; // génération commande fournisseur de contremarque
   31511 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=CF','','BLF') ;
   31512 : AGLLanceFiche('GC','GCTRANSACH_MUL','GP_NATUREPIECEG=BLF','','FF') ;
   // Génération manuelle pièces
   31531 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=CF','','CFR;BLF') ;  // Comm --> Livr
   31532 : AGLLanceFiche('GC','GCGROUPEMANPIECE','GP_NATUREPIECEG=BLF','','LFR;FF') ; // Livr --> Fact
   // génération automatique
   31521 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CF','','BLF') ;
   31522 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLF','','FF') ;
   // tarif fournisseurs
{$IFNDEF SANSPARAM}
   31401 : EntreeTarifFouArt (taModif) ;                       // saisie
   31402 : AGLLanceFiche('GC','GCTARIFFOUCON_MUL','','','') ;  // consultation
   31403 : AGLLanceFiche('GC','GCTARIFFOUMAJ_MUL','','','') ;  // mise à jour
   31404 : AGLLanceFiche('GC','GCPARFOU_MUL','','','') ;  // Import
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
   31704 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','FOU') ;  // modification en série des tiers
{$IFNDEF SANSCOMPTA}
   31705 : OuvreFermeCpte(fbAux,FALSE,'FOU') ;  // fermeture des comptes tiers
   31706 : OuvreFermeCpte(fbAux,True,'FOU') ;   // ouverture compte tiers
   31707 : SuppressionCpteAuxi('FOU') ;
{$ENDIF}
///////  Articles et Stocks  /////////////////
   // Articles
   32501 : AGLLanceFiche('GC','GCARTICLE_MUL','','','GCARTICLE') ; // article
   32502 : AglLanceFiche ('GC','GCPROFILART','','','') ;      // Profil article
   32503 : AGLLanceFiche('GC','GCARTICLE_MODIFLO','','','') ;  // modification en série
   32504 : AGLLanceFiche('GC','GCMULSUPPRART','','','') ;  // épuration articles
   32506 : AGLLanceFiche('GC','GCTARIFART_MUL','','','') ;  // Maj BAses Tarifaires
   32505 : EntreeTradArticle (taModif) ;  // traduction libellé articles
    // Consultation
   32110 : AGLLanceFiche('GC','GCDISPO_MUL','','','') ;
   32111 : AGLLanceFiche('GC','GCMOUVSTK','','','') ; // stock prévisionnel
   32112 : AGLLanceFiche('GC','GCMULCONSULTSTOCK','','','') ; // stock prévisionnel
   32113 : AGLLanceFiche('GC','GCCONTREAVT','','','') ; // suivi contremarque
   {$IFNDEF EAGLCLIENT}
   32201 : AGLLanceFiche('GC','GCMOUVSTKEX','','','ACTION=CREATION;TEM'); // Transferts Inter-Dépôt
   {$ENDIF}
   // Traitement
   32221 : AGLLanceFiche('GC','GCMOUVSTKEX','','','ACTION=CREATION;EEX');
   32222 : AGLLanceFiche('GC','GCMOUVSTKEX','','','ACTION=CREATION;SEX');
   32231 : AGLLanceFiche('GC','GCMOUVSTKEX','','','ACTION=CREATION;RCC'); // Réajustement réservé client
   32232 : AGLLanceFiche('GC','GCMOUVSTKEX','','','ACTION=CREATION;RCF'); // Réajustement réservé fournisseur
   32233 : AGLLanceFiche('GC','GCMOUVSTKEX','','','ACTION=CREATION;RPR'); // Réajustement préparé client
   32203 : AGLLanceFiche('GC','GCTTFINMOIS','','','');
   32204 : AGLLanceFiche('GC','GCTTFINEXERCICE','','','');
   32205 : AGLLanceFiche('GC','GCAFFECTSTOCK','','','');
   32206 : AGLLanceFiche('GC','GCINITSTOCK','','','');  // Initialisation des stocks
   32207 : GCLanceFiche_MouvStkExContr('GC','GCMOUVSTKEXCONTR','','','');  // Mouvements exceptionnels contremarque
   // inventaire ////
   32310 : AGLLanceFiche('GC','GCLISTEINV','','','') ;
   32315 : AGLLanceFiche('GC','GCLISTEINVPRE','','','') ;
   32317 : GCLanceFiche_ListeInvVal('GC','GCLISTEINVVAL','','','') ;
   32320 : AGLLanceFiche('GC','GCSAISIEINV_MUL','','','') ;
   32330 : AGLLanceFiche('GC','GCVALIDINV_MUL','','','') ;
   32340 : GCLanceFiche_ListeInvMul('GC','GCLISTEINV_MUL','','','') ;
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
   {$IFNDEF EAGLCLIENT} 32703 : ParamDimension ; {$ENDIF}
{$ENDIF}
   32711..32715 : ParamTable('GCGRILLEDIM'+IntToStr(Num-32710),taCreat,110000045,PRien) ;
///////////// Analyses ///////////////////////
   // Ventes ////
   33101 : AGLLanceFiche('GC','GCCTRLMARGE','','','') ;
   33102 : AGLLanceFiche('GC','GCANALYSE_VENTES','GPP_VENTEACHAT=VEN','','CONSULTATION') ;
   33105 : AGLLanceFiche('GC','GCVENTES_CUBE','GPP_VENTEACHAT=VEN','','CONSULTATION') ;
   33111 : AGLLanceFiche('GC','GCGRAPHSTAT','','','VEN') ;
   33112 : AGLLanceFiche('GC','GCSTAT','','','VEN') ;
   33113 : AGLLanceFiche('GC','GCSTATCEGID','','','VEN') ;
   33121 : AGLLanceFiche('GC','GCGRAPHPALM','','','VEN') ;
   33122 : AGLLanceFiche('GC','GCPALMARES','','','VEN') ;
   33132 : AGLLanceFiche('AFF','AFVENTES_CUBE','GPP_VENTEACHAT=VEN','','CONSULTATION') ;
   33131 : AGLLanceFiche('AFF','ETATPREVCA','','','');

   // Achats ////
   33201 : AGLLanceFiche('GC','GCCATANA','','','') ;
   33202 : AGLLanceFiche('GC','GCACHAT_CUBE','GPP_VENTEACHAT=ACH','','') ;
   33211 : AGLLanceFiche('GC','GCGRAPHSTAT','','','ACH') ;
   33212 : AGLLanceFiche('GC','GCSTAT','','','ACH') ;
   33221 : AGLLanceFiche('GC','GCGRAPHPALM','','','ACH') ;
   33222 : AGLLanceFiche('GC','GCPALMARES','','','ACH') ;
   // Tableaux de bord Affaires
   33401 : AGLLanceFiche('AFF','AFTBVIEWER','','','');     // Consult tableau de bord
   33402 : AGLLanceFiche('AFF','AFTBVIEWERMULTI','','','MULTI:X');     // Consult TB Multi affaires
   33403 : AGLLanceFiche('AFF','AFGRANDLIVRE','','','') ;  //grand livre ;
   33404 : AGLLanceFiche('AFF','AFBALANCE','','','') ;     //balance;
   33405 : AGLLanceFiche('AFF','AFTABLEAUBORD','','','') ; //AlimTableauBordAffaire;
{$IFNDEF EAGLCLIENT}
   33501 : Assist_DEBC ; // Declaration d'échange de biens (DEB)
{$ENDIF}
{$IFNDEF SANSPARAM}
   {Fonctions E-COMMERCE}
   {$IFNDEF EAGLCLIENT}
   59101 : ExportAllTables;
   59102 : ImportAllTables;
   59103 : BalanceTradTablette;
   59201 : ImporteECommande;
   59202 : AGLLanceFiche('E','EIMPORTCMD','','','');
   59301 : AGLLanceFiche('E','EEXPORTARTICLE','','','');
   59302 : AGLLanceFiche('E','EEXPORTTIERS','','','');
   59303 : AGLLanceFiche('E','EEXPORTTARIF','','','');
   59304 : AGLLanceFiche('E','EEXPORTCOMM','','','');
   59305 : AGLLanceFiche('E','EEXPORTPAYS','','','');
   59401 : AGLLanceFiche('E','ETRADTABLETTE','','','');
   {$ENDIF}
{$ENDIF}
////// Affaires
   // Fichiers affaires
   70101 : AglLanceFiche ('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=PRO','','STATUT=PRO;NOCHANGESTATUT');  // proposition d'affaires
   70102 : AglLanceFiche ('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=AFF','','STATUT=AFF;NOCHANGESTATUT');  // affaires
   70111 : AglLanceFiche ('AFF','AFFAIREAVT_MUL','AFF_STATUTAFFAIRE=PRO','','STATUT=PRO;NOCHANGESTATUT');  // proposition de mission
   70112 : AglLanceFiche ('AFF','AFFAIREAVT_MUL','AFF_STATUTAFFAIRE=AFF','','STATUT=AFF;NOCHANGESTATUT');  // mission
   70103 : AglLanceFiche ('AFF','RESSOURCE_MUL','','','');                        // ressources
   70104 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=PRE','','PRE');    // Prestations
   70105 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=FRA','','FRA');    // Frais
   70107 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=CTR','','CTR');    // articles de contrat
   70106 : AGLLanceFiche ('AFF','HORAIRESTD','','','TYPE:STD');                   // Calendriers
   // Editions Afaires
   70211 : AGLLanceFiche ('AFF','ETATSAFFAIRE','','','');   // Etats sur affaires
   70212 : AGLLanceFiche ('AFF','FICHEAFFAIRE','','','');
   70213 : AGLLanceFiche ('AFF','STATAFFAIRE','','','');
   70221 : AGLLanceFiche ('AFF','ETATSRESSOURCE','','',''); // Etats sur ressources
   70222 : AGLLanceFiche ('AFF','FICHERESSOURCE','','','');
   70223 : AGLLanceFiche ('AFF','STATRESSOURCE','','','');
   70224 : AGLLanceFiche ('AFF','ADRESSERESSOURCE','','','');
   70230 : AGLLanceFiche ('AFF','AFPRINTCALENDRIER','','','');
   // Documents clients
   70301 : AGLLanceFiche('AFF','AFPROPOSEDITMUL','AFF_STATUTAFFAIRE=PRO','','STATUT=PRO') ; // Edition des propositions
   70302 : AGLLanceFiche('AFF','AFPROPOSEDITMUL','AFF_STATUTAFFAIRE=AFF','','STATUT=AFF') ; // Edition des affaires
   // Saisie d'activité Par ressources
   // Consultations
   70701 : AGLLanceFiche('AFF','AFINTERVENANTSAFF',' ','','');                      // intervenants sur affaires
   // Traitements affaires
   70501 : AglLanceFiche ('AFF','AFVALIDEPROPOS','AFF_STATUTAFFAIRE=PRO','','PRO'); // acceptation de missions
   70502 : AGLLanceFiche('AFF','AFAUGMEN_MUL','AFF_STATUTAFFAIRE=AFF','','') ;      // Augmentation affaire
   70503 : AGLLanceFiche('AFF','AFALIGNAFF','','','') ;                             // Alignement client sur affaire
   70511 : AGLLanceFiche('AFF','AFEPUR_ECH_AFF','','','') ;                         // Epuration échéances facturées
   70512 : ; // // génération échéances
   // Traitements sur activité
   70551 : AGLLanceFiche('AFF','AFACTIVITEGLO_MUL','','','AFA_REPRISEACTIV:TOU;');  // Modif ligne activité
   70552 : AGLLanceFiche('AFF','AFACTIVITEVISAMUL','','','');                       // Visas sur l'activité
   70553 : AGLLanceFiche('AFF','REVAL_RESS_MUL','','','');                          // Revalorisation de ressources

   // Editions sur activités
    70601 : AGLLanceFiche('AFF','AFJOURNALACT','','','') ;
    70602 : AGLLanceFiche('AFF','AFJOURNALCLIENT','','','') ;
    70611 : AGLLanceFiche('AFF','AFSYNTHRESS','','','') ;   // Synthèse assistant detaillée par affaire
    70612 : AGLLanceFiche('AFF','AFSYNTHRESSART','','','') ;// Synthèse assistant detaillée par article
    70621 : AGLLanceFiche('AFF','AFSYNTHCLIENT','','','') ; // Synthèse Client affaire détaillée par ressource
    70622 : AGLLanceFiche('AFF','AFSYNTHCLIART','','','') ; // Synthèse Client affaire détaillée par article
   {$IFNDEF SANSCOMPTA}
     {$IFDEF V530}
   60101 : GestionSociete(PRien,@InitSocietePGI) ;
     {$ELSE}
   60101 : GestionSociete(PRien,@InitSocietePGI,Nil) ;
     {$ENDIF}
   {$ENDIF}

   60201 : FicheUserGrp ;
   60202 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;   // Utilisateurs

   {$IFNDEF EAGLCLIENT}
   60102 : RecopieSoc(PRien,True) ;         // recopie société à société
   //////// Utilisateurs et accès
   3172  : AGLLanceFiche('YY','PROFILETABL','','','ETA') ;
   60203 : ReseauUtilisateurs(False) ;      // utilisateurs connectés
   60204 : VisuLog ;                        // Suivi d'activité
   60205 : ReseauUtilisateurs(True) ;       // RAZ connexions
    /////// Outils
   60301 : AGLLanceFiche('GC','GCJNALEVENT_MUL','','','') ;
   60401 : AGLLanceFiche('GC','GCCPTADIFF','','','') ;
      {$IFNDEF SANSPARAM}
   60402 : EntreeStockAjust;
      {$ENDIF}
    /////// Utilitaires
      {$IFNDEF SANSPARAM}
   //60901 : BEGIN Assist_Import;  RetourForce:=TRUE;  END ;        // Import negoce
      {$ENDIF} {SansParam}
///////// Paramètres /////////////
    /////// Paramètres - Société //////////
   65201 : BEGIN
           BrancheParamSocAffiche(StVire,StSeule) ;
           ParamSociete(False,StVire,StSeule,'',Nil,ChargePageSoc,SauvePageSoc,InterfaceSoc,110000220) ;
           END ;
   65203 : AGLLanceFiche('GC','GCETABLISS_MUL','','','') ;
   //65204 : BEGIN LanceAssistCodeAffaire; RetourForce := True;  END;
   {$ENDIF} {eAGL CLient}
    /////// Gestion Commerciale /////////
   {$IFDEF VRELEASE}
   65471 : if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
         AglLanceFiche ('GC','GCPARPIECE','','','') // LanceparPiece
         else AglLanceFiche ('GC','GCPARPIECEUSR','','','') ; // LanceparPiece
   {$ELSE}
   65471 : AglLanceFiche ('GC','GCPARPIECE','','','') ;// LanceparPiece ;
   {$ENDIF} // VRELEASE
   {$IFNDEF EAGLCLIENT}
   65472 : begin EditDocument('L','GPI','',True) ;RetourForce:=True ;end;
   {$IFNDEF SANSPARAM} 65473 : EntreeListeSaisie('') ; {$ENDIF}
   65474 : begin EditEtat('E','GPJ','',True, nil, '', '');RetourForce:=True ;end;
   65475 : ParamTable('GCEQUIPIECE',taCreat,0,PRien) ; // Natures de regroupement
   65401 : AGLLanceFiche('GC','GCPORT_MUL','','','') ;   // Port et frais
  {$IFNDEF SANSCOMPTA} 65402 : FicheSouche('GES') ; {$ENDIF}
   65403 : AGLLanceFiche('GC','GCUNITEMESURE','','','') ;
   65404 : AglLanceFiche ('GC','GCCODECPTA','','','') ;     //parametrage ventilations comptables
   65408 : AGLLanceFiche ('AFF','VENTILANA','','','');      // Ventil ana par affaire
   {$IFNDEF SANSPARAM} 65405 : EntreeArrondi (taModif) ; {$ENDIF}
   65410 : ParamTable('YYDOMAINE',taCreat,110000217,PRien) ; // Domaines activité
   65409 : ParamTable('GCMOTIFMOUVEMENT',taCreat,110000284,PRien) ;          // Motif mouvement stock
   65406 : ParamTable('GCEMPLOIBLOB',taCreat,110000172,PRien) ;          // Emploi blob
   65451 : ParamTable ('GCCOMMENTAIREENT',taCreat,110000167,PRien,6) ;   // commentaires entete
   65453 : ParamTable ('GCCOMMENTAIREPIED',taCreat,110000167,PRien,6) ;  // commentaires pied
   65452 : if GetParamSoc('SO_GCCOMMENTAIRE') then AGLLanceFiche ('GC','GCCOMMENTAIRE','','','')
                                              else ParamTable ('GCCOMMENTAIRELIGNE',taCreat,110000167,PRien,6) ;
   65422 : AGLLanceFiche('AFF','AFBLOCAGEAFFAIRE','','','') ;      // blocage affaire
   65421 : AGLLanceFiche ('AFF','AFPROFILGENER','','','TYPE:DEF'); // profils de facturation par affaire
   // Paramètres Tiers
   65101 : ParamTable('GCZONECOM',taCreat,110000156,PRien) ;        // zones commerciales
   65102 : ParamTable('GCEXPEDITION',taCreat,110000156,PRien) ;     // mode d'expedition
   65103 : ParamTable('GCCOTEARTFOURN',taCreat,110000156,PRien) ;   // cote fournisseur
   65104 : ParamTable('TTTARIFCLIENT',taCreat,110000156,PRien) ;    // code tarif client
   65105 : ParamTable('GCCOMPTATIERS',taCreat,110000156,PRien) ;    // code comptable tiers
   65106 : ParamTable('TTSECTEUR',taCreat,110000156,PRien) ;        // secteurs d'activité
   65107 : ParamTable('GCORIGINETIERS',taCreat,110000156,PRien) ;        // Origine du tiers
   // Paramètres articles
   65301 : ParamTable('GCCOLLECTION',taCreat,0,PRien,3,'Collections') ;     // collection
   65302 : ParamTable('GCTARIFARTICLE',taCreat,0,PRien) ;   // code tarif  article
   65303 : Begin              // libellés des codes familles
           ParamTable('GCLIBFAMILLE',taModif,0,PRien,3,'Titre des familles/sous-familles');
           AfterChangeModule(65);
           end;
   65304 : ParamTable('GCCOMPTAARTICLE',taCreat,0,PRien) ;  // code comptable article
   65305 : ParamTable('GCTYPEEMPLACEMENT',taCreat,0,PRien) ;    // type emplacement
   65306 : ParamTable('GCCOTEEMPLACEMENT',taCreat,0,PRien) ;    // cote emplacement
   65321..65323 : ParamTable('GCFAMILLENIV'+IntToStr(Num-65320),taCreat,0,PRien,3,RechDom('GCLIBFAMILLE','LF'+IntToStr(Num-65320),FALSE)) ;  // familles
   //// Paramètres généraux ///
   65901 : ParamTable('ttFormeJuridique',taCreat,110000187,PRien) ;
   65903 : ParamTable('ttCivilite',taCreat,110000189,PRien) ;
   65902 : ParamTable('ttFonction',taCreat,110000188,PRien) ;
   65904 : AglLanceFiche ('YY','YYPAYS','','','') ;  // OuvrePays ;
   65905 : FicheRegion('','',False) ;
   65906 : OuvrePostal(PRien) ;
   {$IFNDEF SANSCOMPTA}
   65907 : FicheDevise('',tamodif,False) ;
   65908 : ParamTable('TTREGIMETVA',taCreat,110000183,PRien) ;
   65911 : ParamTvaTpf(true) ;
   65912 : ParamTvaTpf(false) ;
   {$ENDIF}
   65921 : FicheModePaie_AGL('');
   65922 : FicheRegle_AGL ('',False,taModif);
   65923 : AglLanceFiche ('AFF','JOURFERIE','','','');
   65924 : ParamTable('ttLangue',taCreat,110000190,PRien) ;
   /////// Paramètres - Tables libres //////////
   65501..65509 : ParamTable('GCLIBREART'+IntToStr(Num-65500),taCreat,110000251,PRien,6,RechDom('GCZONELIBRE','AT'+IntToStr(Num-65500),FALSE)) ;
   65510 : ParamTable('GCLIBREARTA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','ATA',FALSE));
   65511..65519 : ParamTable('GCLIBRETIERS'+IntToStr(Num-65510),taCreat,110000252,PRien,6,RechDom('GCZONELIBRE','CT'+IntToStr(Num-65510),FALSE)) ;
   65520 : ParamTable('GCLIBRETIERSA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','CTA',FALSE));
   65521..65523 : ParamTable('AFTLIBREAFF'+IntToStr(Num-65520),taCreat,0,PRien,6,RechDom('GCZONELIBRE','MT'+IntToStr(Num-65520),FALSE)) ;
   65530 : ParamTable('AFTLIBREAFFA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','MTA',FALSE));
   65531..65533 : ParamTable('AFTLIBRERES'+IntToStr(Num-65530),taCreat,0,PRien,6,RechDom('GCZONELIBRE','RT'+IntToStr(Num-65530),FALSE));
   65540 : ParamTable('AFTLIBRERESA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','RTA',FALSE)) ;
   65640 : ParamTable ('AFTRESILAFF',taCreat,0,PRien);
   65541..65543 : ParamTable('GCLIBREPIECE'+IntToStr(Num-65540),taCreat,0,PRien,6) ;
   65545..65547 : ParamTable('GCLIBREFOU'+IntToStr(Num-65544),taCreat,110000253,PRien,6) ;
   65551..65559 : ParamTable('YYLIBREET'+IntToStr(Num-65550),taCreat,110000255,PRien,6,RechDom('GCZONELIBRE','ET'+IntToStr(Num-65550),FALSE)) ;
   65566..65568 : ParamTable('YYLIBRECON'+IntToStr(Num-65565),taCreat,0,PRien,6,RechDom('GCZONELIBRE','BT'+IntToStr(Num-65565),FALSE)) ; //tables libres contacts
   65560 : ParamTable('YYLIBREETA',taCreat,0,PRien,6,RechDom('GCZONELIBRE','ETA',FALSE)) ;
   /////// Paramètres Affaires
   65601 : ParamTable('AFCOMPTAAFFAIRE',taCreat,110000214,PRien) ;
   65602 : ParamTable ('AFDEPARTEMENT',taCreat,110000214,PRien) ;
   65603 : ParamTable ('AFTLIENAFFTIERS',taCreat,110000214,PRien);
   65604 : ParamTable ('AFTRESILAFF',taCreat,110000214,PRien);
   65605 : ParamTable ('AFETATAFFAIRE',taCreat,110000214,PRien);
    /////// Paramètres Ressources
    65651 : ParamTable ('AFTTYPERESSOURCE',taCreat,110000219,PRien);
    65652 : ParamTable('AFTTARIFRESSOURCE',taCreat,0,PRien) ;
    65653 : AglLanceFiche ('AFF','FONCTION','','','');
    65654 : AglLanceFiche ('AFF','COMPETENCE','','','');
    65655 : ParamTable ('AFTNIVEAUDIPLOME',taCreat,0,PRien);
    65656 : ParamTable ('AFTSTANDCALEN',taCreat,0,PRien);
    65657 : ParamTable ('AFDEPARTEMENT',taCreat,110000214,PRien) ;
    {Tables libres généralisées}
   65561..65565, 65571..65575, 65581..65590, 65592..65593, 65595..65597,
    65525..65529, 65535..65539 : FactoriseZL(Num,PRien) ;
   {$ENDIF} // EAGLCLIENT
{$IFDEF GRC}
   //////// Fiches ///////////
   92101 : AGLLanceFiche('RT','RTPROSPECT_MUL','','','');
   92102 : AGLLanceFiche('YY','YYCONTACTTIERS','','','COURRIER');
   92103 : AGLLanceFiche('RT','RTOPERATIONS_MUL','','','');
   92104 : AGLLanceFiche('RT','RTACTIONS_MUL','','','');
   92107 : AGLLanceFiche('RT','RTPERSPECTIVE_MUL','','','');
   92108 : AGLLanceFiche('RT','RTCONCURRENT_MUL','','','');

   92105 : if GetParamSoc('SO_RTPROJGESTION') = True then
              AGLLanceFiche('RT','RTPROJETS_MUL','','','')
           else
              PGIInfo('Les projets ne sont pas gérés (paramètres société)','Projets');
   92106 : AGLLanceFiche('RT','RTSUSPECT_MUL','','','');
   92110 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE') ;  // ouverture compte tiers
   92109 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME') ;  // fermeture compte tiers


   ///// Editions ///////////////////
   92311 : AGLLanceFiche('RT','RTACTIONS_ETAT2','','','') ;
   92312 : AGLLanceFiche('RT','RTTIERS_VISITE','','','') ;
   92313 : AGLLanceFiche('RT','RTTIERS_VISITE','','','COMMERCIAL') ;

   92320 : AGLLanceFiche('RT','RTOPER_ETAT','','','') ;

   92331 : AGLLanceFiche('RT','RTFICHEPROSPECT','','','') ;
   92332 : AGLLanceFiche('GC','GCETIQCLI','','','') ;
   92333 : AGLLanceFiche('RT','RTLISTEHIERARCHIQ','','','') ;

   92340 : AGLLanceFiche('RT','RTPERSP_ETAT','','','') ;

   92360 : AGLLanceFiche('RT','RTOPER_ETAT_COUT','','','') ;
   92370 : AGLLanceFiche('RT','RTPROSSACTION_TV','','','') ;

   ////// Traitements///////////////
   92202 : AGLLanceFiche('RT','RTGENERE_ACTIONS','','','');  // lancement opération
   92212 : AGLLanceFiche('RT','RTPROS_MAILIN_FIC','','','') ;
   92214 : AGLLanceFiche('RT','RTPRO_MAILIN_CONT','','','FIC') ;
   92222 : AGLLanceFiche('RT','RTPROS_MUL_MAILIN','','','MAI') ;
   92224 : AGLLanceFiche('RT','RTPRO_MAILIN_CONT','','','MAI') ;

   // suspects
   92270 : AGLLanceFiche('RT','RTRECUP_SUSPECT','','','') ;
   92271 : AGLLanceFiche('RT','RTSUSP_MUL_MAILIN','','','MAI') ;
   92272 : AGLLanceFiche('RT','RTSUSP_MAILIN_FIC','','','') ;
  

   92282 : AGLLanceFiche('RT','RTQUALITE','','','CLI;GRC') ;
   92284 : AGLLanceFiche('RT','RTQUALITE','','','PRO;GRC') ;
   92286 : AGLLanceFiche('RT','RTQUALITE','','','CON;GRC') ;
   92288 : AGLLanceFiche('RT','RTQUALITE_SUSP','','','SUS;GRC') ;
   
   92290 : AGLLanceFiche('RT','RTDOUBLONS','','','') ;
   92291 : AGLLanceFiche('RT','RTSUSP_DOUBLE','','','') ;
   
     ///// Analyse ////////
   92402 : AGLLanceFiche('RT','RTANAPROSP','','','');
   92411 : AGLLanceFiche('RT','RTPROSPECTS_TV','','','');
   92412 : AGLLanceFiche('RT','RTACTIONS_TV','','','');
   92413 : AGLLanceFiche('RT','RTPERSPECTIVE_TV','','','');
// mng   
{$IFDEF CEGID}
   92403 : if (VH_RT.RTCodeCommercial <>'') and (VH_RT.RTNomCommercial <>'') then
              AGLLanceFiche('RT','RTCOMMERCIAL_PRO','','',''); // synthese activite du mois
{$ENDIF}
//   92414 : AGLLanceFiche('RT','RTTVENCOURS','','','');

   92421 : AGLLanceFiche('RT','RTTIERS_CUBE','','','');
   92422 : AGLLanceFiche('RT','RTACTIONS_CUBE','','','');
   92423 : AGLLanceFiche('RT','RTPERSP_CUBE','','','');
   92424 : AGLLanceFiche('RT','RTPERSPDETAIL_CUB','','','');
   92425 : AGLLanceFiche('RT','RTSUSPECTS_CUBE','','','');

   {$IFNDEF EAGLCLIENT}
   ////// Outils //////////
   92801 : MajPhonetiqueTiers ;
   92810 : RTConversionPerspEuros;
   92802 : CommercialToRessource;
   92820 : RTCreerProspectPourClient;
   92830 : IntegrationContactsRCC;
   92840 : IntegrationAppelsRCC;
   //92850 : RTIntegrationRessourcePaie;
{$IFDEF CEGID}
//   92850 : RTIntegrationRessource;
//   92860 : RTRecupPrescripteur ;
{$ENDIF}

   {$ENDIF}

   ////// Tablettes //////////
  { 92911 : begin ParamTable('RTTYPEACTION',taCreat,0,PRien) ;
                 AvertirTable ('RTTYPEACTION'); end;}
   92911 : begin AGLLanceFiche('RT','RTTYPEACTIONS','','','');;
                 AvertirTable ('RTTYPEACTION'); end;
   92912 : begin ParamTable('RTTYPEPERSPECTIVE',taCreat,0,PRien) ;
                 AvertirTable ('RTTYPEPERSPECTIVE'); end;
   { suspects92913 : begin ParamTable('RTETATACTION',taModif,0,PRien) ;
                 AvertirTable ('RTETATACTION'); end;
   92914 : begin ParamTable('RTETATPERSPECTIVE',taModif,0,PRien) ;
                 AvertirTable ('RTETATPERSPECTIVE'); end;
   }
   92913 : begin ParamTable('RTLIBCHPLIBSUSPECTS',taModif,0,PRien) ;
                 AvertirTable ('RTLIBCHPLIBSUSPECTS'); end;
   92914 : begin ParamTable('RTMOTIFFERMETURE',taCreat,0,PRien) ;
                 AvertirTable ('RTMOTIFFERMETURE'); end;
   92915 : begin
           ParamTable('RTLIBCLACTIONS',taModif,0,PRien) ;
           AvertirTable ('RTLIBCLACTIONS') ;
           AfterChangeModule(92);
           end ;
   // suspect : rien à voir mais correction tacreat        
   92916 : begin ParamTable('RTOBJETOPE',taCreat,0,PRien) ;
                 AvertirTable ('RTOBJETOPE'); end;
   92917 : begin ParamTable('RTMOTIFS',tacreat,0,PRien) ;
                 AvertirTable ('RTMOTIFS'); AvertirTable ('RTMOTIFSIGNATURE'); end;
   92950 : begin ParamTable('RTIMPORTANCEACT',taCreat,0,PRien) ;  // Niveau d'importance des actions
                 AvertirTable ('RTIMPORTANCEACT'); end;                 
   // suspect92920 : AglLanceFiche ('AFF','RESSOURCE_MUL','','','');  // ressources
   92918 : AglLanceFiche ('AFF','RESSOURCE_MUL','','','');  // ressources
   92920..92929 : ParamTable('RTRSCLIBTABLE'+IntToStr(Num-92920),taCreat,0,PRien,3,RechDom('RTLIBCHPLIBSUSPECTS','CL'+IntToStr(Num-92920),FALSE)) ;
   92930 : AGLLanceFiche('RT','RTPARAMCL','','',''); //ParamSaisieTL ();
   92945 : AGLLanceFiche('RT','RTPARAMPROSPCOMPL','','','');
   92946 : if GetParamSoc('SO_RTCONFIDENTIALITE') = True then
              AGLLanceFiche('RT','RTUTILISATEUR_MUL','','','')
           else
              PGIInfo('La confidentialité n''est pas gérée (paramètres société)','Confidentialité');

   92951..92960 : ParamTable('RTRPRLIBTABLE'+IntToStr(Num-92951),taCreat,0,PRien,3,RechDom('RTLIBCHAMPSLIBRES','CL'+IntToStr(Num-92951),FALSE)) ;
   92961..92976 : ParamTable('RTRPRLIBTABLE'+IntToStr(Num-92951),taCreat,0,PRien,3,RechDom('RTLIBCHAMPSLIBRES','CL'+Chr(Num-92951+55),FALSE)) ;
   92980..92989 : ParamTable('RTRPRLIBTABMUL'+IntToStr(Num-92980),taCreat,0,PRien,3,RechDom('RTLIBCHAMPSLIBRES','ML'+IntToStr(Num-92980),FALSE)) ;
   92991..92993 : ParamTable('RTRPRLIBACTION'+IntToStr(Num-92990),taCreat,0,PRien,3,RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(Num-92990),FALSE)) ;
   92996..92998 : ParamTable('RTRPRLIBPERSPECTIVE'+IntToStr(Num-92995),taCreat,0,PRien,3,RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(Num-92995),FALSE)) ;
   // mng : now 12 tables
   92932..92941 : ParamTable('RTRPCLIBTABLE'+IntToStr(Num-92932),taCreat,0,PRien,3,RechDom('RTLIBTABLECOMPL','TL'+IntToStr(Num-92932),FALSE)) ;
   92942 : ParamTable('RTRPCLIBTABLEA',taCreat,0,PRien,3,RechDom('RTLIBTABLECOMPL','TLA',FALSE)) ;
   92943 : ParamTable('RTRPCLIBTABLEB',taCreat,0,PRien,3,RechDom('RTLIBTABLECOMPL','TLB',FALSE)) ;
{$ENDIF}
   
      else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
var ii : integer ;
    Arg5LanceFiche : string ;
BEGIN
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
   13 : {Dépots} AGLLanceFiche('GC','GCDEPOT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   14 : {Apporteurs} AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)+';GCL_TYPECOMMERCIAL="APP"') ;
   15 : {Règlement}  FicheRegle_AGL ('',False,taModif) ;
   16 : {contacts} AGLLanceFiche('YY','YYCONTACT',Lequel,Range,ActionToString(Action)+';MONOFICHE') ;
   17 : {Fonction des ressources} AGLLanceFiche('AFF','FONCTION','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   18 : {Etablissements} AGLLanceFiche('GC','GCETABLISS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   19 : {Contenu tables libres} AGLLanceFiche('GC','GCCODESTATPIECE','','','ACTION=CREATION;YX_CODE='+Range) ;
   20 :  ;
   21 : {Ports} AGLLanceFiche('GC','GCPORT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   22 : {GRC-action} AGLLanceFiche('RT','RTACTIONS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   23 : {GRC-operation} AGLLanceFiche('RT','RTOPERATIONS','',Lequel,ActionToString(Action)+';MONOFICHE') ;

   28 : {Clients par T_AUXILIAIRE}
        BEGIN
        Arg5LanceFiche:=ActionToString(Action)+';MONOFICHE' ;
        If Range='' then Arg5LanceFiche:=Arg5LanceFiche+ ';T_NATUREAUXI=CLI'  else Arg5LanceFiche:=Arg5LanceFiche+Range ;
        AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche) ;
        END ;
   29 : {Fournisseurs via T_AUXILIAIRE} AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=FOU') ;
   30 : {GRC-Projet} AGLLanceFiche('RT','RTPROJETS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   // *** maj des tablettes ***
   900 : ParamTable('GCCOMMENTAIRELIGNE',taCreat,0,Nil,6);  // mis de manière explicite car appellé depuis un Grid
   // pour mise à jour par defaut des tablettes
   996 : ParamTable(TT,taModif,0,Nil,6) ;   {choixext}
   997 : ParamTable(TT,taCreat,0,Nil,6) ;   {choixext}
   998 : ParamTable(TT,taModif,0,Nil,3) ;   {choixcode ou commun}
   999 : ParamTable(TT,taCreat,0,Nil,3) ;   {choixcode ou commun}
  end ;
END ;

Procedure RemoveTablesLibresS3 ( TT : Array of integer ) ;
Var i : integer ;
BEGIN
for i:=Low(TT) to High(TT) do FMenuG.RemoveItem(TT[i]) ;
END ;

Procedure TripoteMenuGCS3 ( NumModule : integer ) ;
BEGIN
Case NumModule of
   30 : BEGIN
        FMenuG.RemoveItem(30237) ; {Duplication Facture provisoire}
        FMenuG.RemoveItem(30401) ; {Visa des pièces}
        END ;
   31 : BEGIN
        FMenuG.RemoveItem(31501) ; {Visa des pièces}
        END ;
   32 : BEGIN
        FMenuG.RemoveItem(32505) ;       {Traductions articles}
        FMenuG.RemoveItem(32602) ;       {Emplacements}
        FMenuG.RemoveGroup(32700,True) ; {Dimensions}
        END ;
   33 : BEGIN
        FMenuG.RemoveItem(-33130) ; {Analyses par affaires}
        END ;
   65 : BEGIN
        FMenuG.RemoveItem(65102)  ; {Mode expédition}
        FMenuG.RemoveItem(65204)  ; {Assistant code affaire}
        FMenuG.RemoveItem(65305)  ; {Types emplacement}
        FMenuG.RemoveItem(65306)  ; {Cotes emplacements}
        FMenuG.RemoveItem(65410)  ; {Domaines activité}
        FMenuG.RemoveItem(65406)  ; {Emploi mémos/photos}
        FMenuG.RemoveItem(65473)  ; {Listes de saisie}
        FMenuG.RemoveItem(65475)  ; {Natures de regoupement}
        FMenuG.RemoveItem(-65420) ; {Affaires}
        FMenuG.RemoveItem(-65520) ; {Libres affaires}
        FMenuG.RemoveItem(-65524) ; {Lib Libres affaires}
        FMenuG.RemoveItem(-65250) ; {Champs libres Affaires}
        FMenuG.RemoveItem(-65530) ; {Libres ressources}
        FMenuG.RemoveItem(-65534) ; {Lib Libres ressources}
        FMenuG.RemoveItem(-65260) ; {Champs libres Ressources}
        FMenuG.RemoveItem(-65540) ; {Pièces}
        FMenuG.RemoveItem(-65550) ; {Etablissements}
        FMenuG.RemoveItem(-65580) ; {Libellés Etab.}
        FMenuG.RemoveItem(65924)  ; {Langues}
        FMenuG.RemoveItem(65923)  ; {Jours fériés}
        FMenuG.RemoveGroup(65600,True) ;
        // grc
        //FMenuG.RemoveGroup(65650,True) ;{Tablettes affaire, ressources}
        FMenuG.RemoveItem(65652)  ; {Codes tarifs}
        FMenuG.RemoveItem(65654)  ; {Compétences}
        FMenuG.RemoveItem(65655)  ; {Niveau de diplome}
        FMenuG.RemoveItem(65656)  ; {Type de calendrier}

        RemoveTablesLibresS3([65562,65563,65564,65565,65504,65505,65506,65507,65508,65509,65510]) ;
        RemoveTablesLibresS3([65572,65573,65574,65575,65514,65515,65516,65517,65518,65519,65520]) ;
        RemoveTablesLibresS3([65596,65597]) ;
        RemoveTablesLibresS3([65587,65588,65589,65590]) ;
        RemoveTablesLibresS3([65587,65588,65589,65590]) ;
        RemoveTablesLibresS3([65582,65583,65584,65585,65555,65556,65557,65558,65559,65560]) ;
        END ;
   END ;
END ;

procedure AfterChangeModule ( NumModule : integer ) ;
Var i : integer ;
    St : String; // GRC
BEGIN
AjusteMenu(NumModule);
TripoteMenuGCS3(NumModule) ;
TraiteMenusGRC(NumModule) ;
TraiteMenus303265(NumModule) ;
TraiteMenusAffaire(NumModule) ;
TraiteSpecifsCEGID(NumModule) ;
TraiteMenusEAGL(NumModule) ;
{$IFDEF EAGLCLIENT}
   {$IFDEF AGL550B}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(5,FMenuG.DispatchX) ;
  92 : ChargeMenuPop(8,FMenuG.DispatchX) ;
   else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  END ;
    {$ELSE}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(hm5,FMenuG.DispatchX) ;
  92 : ChargeMenuPop(hm8,FMenuG.DispatchX) ;
   else ChargeMenuPop(TTypeMenu(NumModule),FMenuG.DispatchX) ;
  END ;
  {$ENDIF}
{$ELSE}
 {$IFDEF AGL545}
    {$IFDEF AGL550B}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(5,FMenuG.DispatchX) ;
  92 : ChargeMenuPop(8,FMenuG.DispatchX) ;
   else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  END ;
    {$ELSE}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(hm5,FMenuG.DispatchX) ;
  92 : ChargeMenuPop(hm8,FMenuG.DispatchX) ;
   else ChargeMenuPop(TTypeMenu(NumModule),FMenuG.DispatchX) ;
  END ;
    {$ENDIF}
 {$ELSE}
Case NumModule of
  30,31,32,33,34,65,70 : ChargeMenuPop(hm5,FMenuG.Dispatch) ;
  92 : ChargeMenuPop(hm8,FMenuG.Dispatch) ;
   else ChargeMenuPop(TTypeMenu(NumModule),FMenuG.Dispatch) ;
  END ;
 {$ENDIF}
{$ENDIF}
FMenuG.RemovePopupItem(70102) ;
FMenuG.RemovePopupItem(92107) ;
END ;


procedure InitApplication ;
BEGIN
     ProcZoomEdt:=GCZoomEdtEtat ;
     ProcCalcEdt:=GCCalcOLEEtat ;
{$IFDEF GRC} ProcZoomEdt:=RTZoomEdtEtat ; {$ENDIF}
     FMenuG.OnDispatch:=Dispatch ;
     FMenuG.OnChargeMag:=ChargeMagHalley ;
    {$IFDEF EAGLCLIENT}
     FMenuG.OnMajAvant:=Nil ;
     FMenuG.OnMajApres:=Nil ;
    {$ELSE}
     FMenuG.OnMajAvant:=Nil;
     FMenuG.OnMajApres:=Nil ;
     FMenuG.OnMajPendant:=Nil ;
    {$ENDIF}
     FMenuG.OnAfterProtec:=AfterProtecGC ;
     FMenuG.OnChangeModule:=AfterChangeModule ;
     V_PGI.DispatchTT:=DispatchTT ;

     FMenuG.SetModules([60],[49]) ;
     V_PGI.NbColModuleButtons:=1 ; V_PGI.NbRowModuleButtons:=6 ;

     FMenuG.SetPreferences(['Pièces'],False) ;

    // definition des constantes dans UtilDispGC
    FMenuG.SetSeria(GCCodeDomaine,GCCodesSeria,GCTitresSeria) ;  // PCS
END ;

initialization
RegisterGCPGI ;

end.

