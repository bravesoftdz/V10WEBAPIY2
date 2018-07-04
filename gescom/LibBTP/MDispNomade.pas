unit MDispNomade;

interface

uses Forms, sysutils;

procedure InitApplication;
procedure AfterChangeModule(NumModule: integer);

type
  TFMenuDisp = class(TDatamodule)
  end;

var FMenuDisp: TFMenuDisp;


implementation

uses Windows, HMsgBox, Classes, Hctrls, Controls, HPanel, UIUtil, ImgList, ParamSoc,
     FE_Main, MenuOLG, AGLInit, tablette, pays, region, codepost, UtomArticle, EdtQR,
     CreerSoc, CopySoc, Reseau, Mullog, LicUtil, FactUtil, EdtEtat, AssistInitSoc,
     Ent1, saisutil, UTomPiece, EntGC, UtilGC, FichComm, ZoomEdtGescom, CalcOLEGescom,
     UtofConsoEtat, UTofGcTraiteAdresse, UtilSoc, AssistSuggestion, TiersUtil, AssistCodeAffaire,
     EdtDOC, Facture, HistorisePiece_Tof, PieceEpure_Tof, 
     UTOFGCLISTEINV_MUL, UTOFListeInvVal, MouvStkExContr_TOF,regtva,
     UTofMouvStkEx, UtilNomade, FOExportCais, uToxClasses, uToxConst, uToxFiches, GCTOXWORK, ToxSim,
     ToxDetop, uToxConf, ToxConsultBis, GCToxCtrl, ArtPrestation_Tof, UtilDispGC,
     UtilArticle, HEnt1, UtomUtilisat, UTofProspect_Mul, InitPCP, GCOleAuto, EdtREtat, UEdtComp, EdtRdoc,
     RegGRCPGI, RTMenuGRC, CalcMulProspect, entPGI, GrpFavoris, UtilPGI, ChangeVersions,UserGrp_tom,UtilEvent,
     Tva,Arrondi,Reconduction_tof,UtilMenuAff,
  {$IFNDEF SANSCOMPTA}
     devise,
  {$ENDIF}
{$IFDEF BTP}
		 BTPUTIL,PARAMCATTAXE_TOF, 
{$ENDIF}
     ToxMoteur,UTofGCAff_Mul,Souche, Confidentialite_TOF,TarifArticle,TarifCatArt,TarifTiers,TarifCliArt,Dimension,InitSoc;

var lesTagsToRemove: string;

{$R *.DFM}


procedure AppelTitreLibre (Name : string;PRien : THPanel );
begin
     // fct qui permet de ne saisir que la partie voulue de la tablette des titres libres
GCModifTitreZonesLibres (Name);
ParamTable('GCZONELIBRE',taModif,0,PRien,3);
GCModifTitreZonesLibres ('');
if ctxScot in V_PGI.PGIContexte then AfterChangeModule(144)
                                else AfterChangeModule(148);
end;

function ChargeFavoris: boolean;
begin
  AddGroupFavoris(FMenuG, lesTagsToRemove);
  result := true;
end;

procedure Dispatch(Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var StVire,StSeule,stMessage : string;
begin
  case Num of
    10  : begin //  après connexion
            VH_GC.GcMarcheVentilAna:='AFF';
            TraiteChangementVersions;

{$IFNDEF EAGLCLIENT}
            if PGI_IMPORT_BOB('BAT3') = 1 then
            begin
              stMessage := 'LE SYSTEME A DETERMINE UNE MAJ DE MENU';
              stMessage := stMessage +#13#10 +'POUR QU''ELLE SOIT PRISE EN COMPTE';
              stMessage := stMessage +#13#10 +'VOUS DEVREZ RELANCER L''APPLICATION';
              PGIInfo(stMessage,'MAJ MENU');
            end;
{$ENDIF}
        		InitParPieceBTP ;
//
            InitMetier;
            UpdateCombosGC;
//            GCModifLibelles;
            GCTripoteStatus;
            PCP_LesSites := TCollectionSites.Create(TCollectionSite, True); // Chargement des sites au démarrage de l'appli.
            TestSeriaNomade;
            TrouvePiecesPCP;
            if VerifEtatBasePCP() < 3 then // Lancement automatique de l'assistant PCP si 1ère connexion
              PCPAssistInit('');
            TraiteMenusPCP(285);
           end;
    11  : ; //  après déconnexion
    12  : ; //  avant connexion et séria
    13  : ChargeMenuPop(-1,FMenuG.DispatchX) ;
    15  : ;
    16  : begin
            lesTagsToRemove := '111210;111230;111240;111100;111110;111120;111130;111140;111150;' +
                               '111300;111310;111320;60501';
            GetVarPCP(V_PGI.User);
            ChargeModules_PCP;
          end;
    100 : ;
// VENTES
    //Saisie
    30201  : AGLLanceFiche('BTP','BTDEVIS_MUL','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN','','MODIFICATION') ;   //devis;
    //Consultations
    30301  : AGLLanceFiche('GC','GCPIECE_MUL','GP_VENTEACHAT=VEN;GP_NATUREPIECEG=DBT','','CONSULTATION') ; //Pièces
    30302  : AGLLanceFiche('GC','GCLIGNE_MUL','GL_NATUREPIECEG=DBT','','CONSULTATION;VENTEACHAT:VEN') ; //Lignes
    //Editions et Analyses
    30601  : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=DBT','','VEN') ;

// DONNEES DE BASE
		// AFFAIRE
    145110 : if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
          			AGLLanceFiche('BTP', 'BTAFFAIREINT_MUL','AFF_AFFAIRE=A','','AFF')
    				 else	AglLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE=A','','AFF');   // Affaires
    145120 : AGLLanceFiche('BTP','BTETATSAFFAIRES','AFF_AFFAIRE0=A','','');   // Etats sur affaires
		// ARTICLE
    149210 : AGLLanceFiche('BTP','BTARTICLE_MUL','','','MAR');
    149220 : AGLLanceFiche('BTP','BTARTICLE_MUL','GA_TYPEARTICLE=PRE','','PRE');
    149230 : AGLLanceFiche('BTP','BTARTICLE_MUL','','','NOM');
    149245 : AGLLanceFiche('BTP', 'BTVARIABLE', 'G;GENERAL','', 'GENERALES');
    149290 : AGLLanceFiche('BTP','BTARTICLE_MUL','GA_TYPEARTICLE=FRA','','FRA');
    149250 : AGLLanceFiche('GC','GCFAMHIER','','','ACTION=MODIFICATION;NIV1=FN1;NIV2=FN2;NIV3=FN3') ;
    149240 : AglLanceFiche ('BTP','BTPROFILART','','','') ;      // Profil article
    149260 : AglLancefiche ('BTP','BTNATPREST_MUL','','','');
	  // TARIF
{$IFNDEF SANSPARAM}
		// TARIF HT
    30102 : EntreeTarifArticle(taModif) ;
    30105 : EntreeTarifCatArt(taModif) ;
    30103 : EntreeTarifTiers(taModif) ;
    30104 : EntreeTarifCliArt(taModif) ;
    30106 : AGLLanceFiche('GC','GCTARIFCON_MUL','','','') ;
    30107 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','') ;
    30108 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','DUP') ; // duplication de tarif
    30109 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','SUP') ; // suppression de tarif
   // Tarifs TTC
    30902 : EntreeTarifArticle(taModif,TRUE);
    30903 : EntreeTarifTiers(taModif,TRUE);
    30904 : EntreeTarifCliArt(taModif,TRUE) ;
{$ENDIF}
    // DIMENSION
    149110 : begin ParamTable('GCCATEGORIEDIM',taModif,0,PRien,3); AfterChangeModule(149); end;  // categorie
    149130 : AGLLanceFiche('GC','GCMASQUEDIM','','','') ;
{$IFNDEF SANSPARAM}
    149140 : BEGIN
             ParamDimension ;  // valeurs
             AvertirTable ('GCGRILLEDIM1');
             AvertirTable ('GCGRILLEDIM2');
             AvertirTable ('GCGRILLEDIM3');
             AvertirTable ('GCGRILLEDIM4');
             AvertirTable ('GCGRILLEDIM5');
             END;
{$ENDIF}
    149121..149125 : BEGIN
                     ParamTable('GCGRILLEDIM'+IntToStr(Num-149120),taCreat,0,PRien,3,RechDom('GCCATEGORIEDIM','DI'+IntToStr(Num-149120),FALSE)) ;
                     AvertirTable ('GCGRILLEDIM'+IntToStr(Num-149120));
                     END;
	  // Clients
    30501 : if ctxMode in V_PGI.PGIContexte then
              AGLLanceFiche('GC','GCCLIMUL_MODE','T_NATUREAUXI=CLI','','')
{$IFDEF GRC}
           else if ctxGRC in V_PGI.PGIContexte then
              RTLanceFiche_Prospect_Mul('RT','RTPROSPECT_MUL','T_NATUREAUXI=CLI','','GC')
{$ENDIF}
           else
              AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=CLI','','') ;    // Clients
    30502 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=CLI','','') ; // contacts
    105107 : AGLLanceFiche('MBO','EDTCLI_MODE','','','') ;      // Edition des clients
    // prospect
    149351 : AGLLanceFiche('RT','RTPROSPECT_MUL','T_NATUREAUXI=PRO','','');
    // fournisseurs
    31701 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','') ;
    31702 : AGLLanceFiche('GC','GCMULCATALOGUE','','','') ;
    138512 : AGLLanceFiche('GC','GCFOURN_ARTICLE','','','') ;
    31703 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=FOU','','') ;  // contacts
    // Tarif fournisseurs
    31401 : EntreeTarifFouArt (taModif) ;                       // saisie
    31402 : AGLLanceFiche('GC','GCTARIFFOUCON_MUL','','','') ;  // consultation
    31403 : AGLLanceFiche('GC','GCTARIFFOUMAJ_MUL','','','') ;  // mise à jour
    // Commerciaux
    30503 : AGLLanceFiche ('GC','GCCOMMERCIAL_MUL','','','') ; // Commerciaux

// TRANSMISSIONS ET PARAMETRES
    111200: begin // Démarrage/Arrêt des échanges
    				if V_PGI.superviseur then
            begin
             if VH_GC.PCPVenteSeria then ActiveEvent('PCPS_REMONTEE_VTE');
             if VH_GC.PCPAchatSeria then ActiveEvent('PCPS_REMONTEE_ACH');
            end else
            begin
             if VH_GC.PCPVenteSeria then ActiveEvent('PCPU_REMONTEE_VTE');
             if VH_GC.PCPAchatSeria then ActiveEvent('PCPU_REMONTEE_ACH');
            end;
             if GetParamSoc('SO_GCTOXCONFIRM') = True then
               AglToxConf(aceConfigure, '', nil, GescomModeTraitementTox, nil, ScruteTobTox)
               else
               AglToxConf(aceConfigure, '', GescomModeNotConfirmeTox, GescomModeTraitementTox, nil, ScruteTobTox);
            end;
    111210 : AglToxMulChronos; // Consultation des échanges
    111220 : AfficheCorbeille; // Consultation des corbeilles
    111230 : AglLanceFiche('MBO', 'CONSOLIDATION', '', '', '');
    111240 : LanceEtat('E','MST','IVE',True,False,False,Nil,'','',False);
    111100 : AglToxSaisieParametres; // Paramètres par défaut
    111110 : AglToxSaisieVariables; // Variables
    111120 : AglToxMulSites; // Sites
    111130 : AglToxSaisieGroupes; //  Groupes
    111140 : AglToxMulConditions; // Requêtes
    111150 : AglToxMulEvenements; // Evénements
    111300 : ToxSimulation; // Intégration d'un fichier
    111310 : ConsultManuelTox; // Visualisation d'un fichier
    111320 : DetopeTox; // Détopage d'un fichier

// ADMINISTRATION
    60501 : ParamFavoris('',LesTagsToRemove,False,False); // Gestion des favoris
    60208 : if ctxMode in V_PGI.PGIContexte then
              GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','26;167;168;169;170;171')
              else
              GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','26;27;149;169;285;60');
    60202 : begin // Utilisateurs
              FicheUSer(V_PGI.User) ;
              ControleUsers ;
            end;
    60205 : ReseauUtilisateurs(True); // RAZ connexions
    65201 : begin  // Paramètres société
              BrancheParamSocAffiche(StVire,StSeule) ;
              ParamSociete(False,StVire,StSeule,'',Nil,ChargePageSoc,SauvePageSoc,InterfaceSoc,110000220) ;
            end;
    65203 : AGLLanceFiche('GC','GCETABLISS_MUL','','','') ; // Etablissements
    65471 : if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then // ParPièce
              AglLanceFiche ('GC','GCPARPIECE','','','');
    65402 : FicheSouche('GES') ; // Souche
  // ---- Administration ---- //
{$IFNDEF SANSCOMPTA} // aujourd'hui InitSociété= compta
{$IFNDEF V530}
   60101 : GestionSociete(PRien,@InitSociete,Nil) ; // Gestionnaire société
{$ELSE}
   60101 : GestionSociete(PRien,@InitSociete) ; // Gestionnaire société
{$ENDIF}
{$ENDIF}
{$IFNDEF SANSPARAM}
    60201 : FicheUserGrp ;                   // groupe utilisateurs
    60403 : Entree_TraiteAdresse ;
{$ENDIF}
{$IFNDEF EAGLCLIENT}
    60211 : AGLLanceFiche('YY','PROFILETABL','','','ETA;GRP') ; // JTR - Fonction en CWAS
    3172  : AGLLanceFiche('YY','PROFILETABL','','','ETA;UTI') ; // JTR - Fonction en CWAS
{$ELSE}
    60211 : AGLLanceFiche('YY','PROFILETABL','ETA;GRP','','ETA;GRP') ; // JTR - Fonction en CWAS
    3172  : AGLLanceFiche('YY','PROFILETABL','ETA;UTI','','ETA;UTI') ; // JTR - Fonction en CWAS
{$ENDIF EAGLCLIENT}
   // --
   60203 : ReseauUtilisateurs(False) ;      // utilisateurs connectés
   60204 : VisuLog ;                        // Suivi d'activité
   60206 : AGLLanceFiche('GC','GCPARAMOBLIG','','','') ;      // Champs obligatoires
   60207 : AGLLanceFiche('GC','GCPARAMCONFID','','','') ;       // restricyions fiches

// PARAMETRES
	 74101 : BEGIN
              BrancheParamSocAffiche(StVire,StSeule) ;
              ParamSociete(False,StVire,StSeule,'',Nil,ChargePageSoc,SauvePageSoc,InterfaceSoc,110000220) ;
           END;
   74109 : BEGIN LanceAssistCodeAffaire; RetourForce := True;  END;
   74102 : AGLLanceFiche('GC','GCETABLISS','','','') ;          // FicheEtablissement_AGL(taModif);
   74103 : FicheSouche('GES');                                  // Compteurs de pièces
   74104 : AglLanceFiche ('GC','GCPARPIECE','','','') ;
   74105 : begin EditEtat('E','GPJ','',True, nil, '', '');RetourForce:=True ;end;
   148100 : AglLanceFiche('BTP','BTVARIABLE','A;APPLICATION','','APPLICATION');
   74151 : RegimesTva ;
   148801 : ParametrageTypeTaxe;
   74152 : ParamTvaTpf(true) ;
   74153 : FicheModePaie_AGL('');
   74154 : FicheRegle_AGL ('',False,taModif);
   74155 : FicheDevise('',tamodif,False);
   74156 : ParamTable('ttFormeJuridique',taCreat,1120000,PRien) ;
   74157 : ParamTable('ttCivilite',taCreat,1122000,PRien) ;
   74158 : OuvrePays ;
   74159 : FicheRegion('','',False) ;
   74160 : OuvrePostal(PRien) ;
   74161 : ParamTable('ttLangue',taCreat,0,PRien) ;
   74202 : AglLanceFiche ('GC','GCCODECPTA','','','') ; // Ventilation comptable GCCOMPTESHT_MUL
   74206 : begin
             SetParamSOc ('SO_GCAXEANALYTIQUE',TRUE);  //brl 14/10 force paramétrage des axes analytiques
             AGLLanceFiche('MBO','VENTILANA','','','') // ventil générique
           end;
   74208 : AGLLanceFiche('GC','GCPORT_MUL','','','') ;   // Port et frais
   74203 : AGLLanceFiche('GC','GCUNITEMESURE','','','');
   74205 : EntreeArrondi (taModif) ;
   74204 : ParamTable ('GCEMPLOIBLOB',taCreat,0,PRien) ;
   // Emplacement de stocks
   105139 : ParamTable('GCTYPEEMPLACEMENT',taCreat,0,PRien) ;    // type emplacement
   105140  : ParamTable('GCCOTEEMPLACEMENT',taCreat,0,PRien) ;    // côte emplacement
   /// contrats
   14851 : ParametrageTypeAction;
   14852 : ParamTable ('AFTRESILAFF',taCreat,0,PRien);
   14853 : ParamTable ('TTPRIOCONTRAT',taCreat,0,PRien);
   92988 : begin
   						AGLLanceFiche('RT','RTTYPEACTIONS','---;GRC','','GRC');;
              AvertirTable ('RTTYPEACTION');
   				 end;
   148251 : AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=DBT;NUMERO=0') ;
   148252 : AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=FBT;NUMERO=0') ;
   148253 : AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=ABT;NUMERO=0') ;
   148254 : AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=ETU;NUMERO=0') ;
   148255 : AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=DAC;NUMERO=0') ;
   74311 : ParamTable('YYCODENAF',taCreat,0,PRien,3,'Code NAF') ;
   74312 : ParamTable('GCCOMPTATIERS',taCreat,0,PRien) ;
   74313 : ParamTable ('GCZONECOM',taCreat,0,PRien);
   74314 : ParamTable('TTSECTEUR',taCreat,0,PRien) ;       // secteurs d'activité
   74315 : ParamTable('TTTARIFCLIENT',taCreat,0,PRien) ;
   74316 : ParamTable('GCORIGINETIERS',taCreat,0,PRien) ;
   // Contacts
   74321 : ParamTable('ttFonction',taCreat,1125000,PRien) ;
    // Affaire
   74331 : ParamTable('AFCOMPTAAFFAIRE',taCreat,0,PRien) ;
   74332 : ParamTable ('AFDEPARTEMENT',taCreat,0,PRien) ;
   74333 : ParamTable ('AFTLIENAFFTIERS',taCreat,0,PRien);
   74337 : ParamTable ('AFETATAFFAIRE',taCreat,120000158,PRien);
   // Articles
   74341..74343 : ParamTable('GCFAMILLENIV'+IntToStr(Num-74340),taCreat,0,PRien) ;
   74344 : ParamTable('GCLIBFAMILLE',taModif,0,PRien) ;
   74345 : ParamTable('GCCOMPTAARTICLE',taCreat,0,PRien) ;
   74346 : ParamTable('GCTARIFARTICLE',taCreat,0,PRien) ;
   65305 : ParamTable('GCTYPEEMPLACEMENT',taCreat,0,PRien) ;    // type emplacement
   65306 : ParamTable('GCCOTEEMPLACEMENT',taCreat,0,PRien) ;    // côte emplacement
   // Appel d'offre
   74351 : BEGIN ParamTable ('BTNATUREDOC',taCreat,0,PRien); AvertirTable ('BTNATUREDOC'); END;
   148651 : ParamTable('GCTYPECOMMERCIAL',taCreat,0,PRien) ;
   148661 : ParamTable('AFTTYPEHEURE',taCreat,0,PRien) ;
   74456 : AppelTitreLibre('PT',Prien);  // Libellé tables libres pièces
   74457 : AppelTitreLibre('PD',Prien);  // Libellé dates libres pieces
   74451..74453 : ParamTable('GCLIBREPIECE'+IntToStr(Num-74450),taCreat,0,PRien,6) ;  // Stats Pièces
   /////// Tables libres //////////
   74402 : AppelTitreLibre('CT',Prien);  // Libellé tables libres client
   74403 : AppelTitreLibre('CM',Prien);  // Libellé mtt libres client
   74404 : AppelTitreLibre('CD',Prien);  // Libellé dates libres client
   74405 : AppelTitreLibre('CC',Prien);  // Libellé textes libres client
   74406 : AppelTitreLibre('CB',Prien);  // Libellé bool libres client
   74411..74419 : ParamTable('GCLIBRETIERS'+IntToStr(Num-74410),taCreat,0,PRien,6) ;  // Stats clients
   74420 : ParamTable('GCLIBRETIERSA',taCreat,0,PRien,6) ;  // Stats clients
   74481 : AppelTitreLibre('MT',Prien);  // Libellé tables libres affaire
   74482 : AppelTitreLibre('MM',Prien);  // Libellé mtt libres affaire
   74483 : AppelTitreLibre('MD',Prien);  // Libellé dates libres affaire
   74484 : AppelTitreLibre('MC',Prien);  // Libellé textes libres affaire
   74485 : AppelTitreLibre('MB',Prien);  // Libellé bool libres affaire
   74461..74469 : ParamTable('AFTLIBREAFF'+IntToStr(Num-74460),taCreat,0,PRien,6) ;   // Stats Affaire
   74489 : ParamTable('AFTLIBREAFFA',taCreat,0,PRien,6) ;   // Stats Affaire
   74491 : AppelTitreLibre('RT',Prien);  // Libellé tables libres ressource
   74492 : AppelTitreLibre('RM',Prien);  // Libellé mtt libres ressource
   74493 : AppelTitreLibre('RD',Prien);  // Libellé dates libres ressource
   74494 : AppelTitreLibre('RC',Prien);  // Libellé textes libres ressource
   74495 : AppelTitreLibre('RB',Prien);  // Libellé bool libres ressource
   74431..74439 : ParamTable('AFTLIBRERES'+IntToStr(Num-74430),taCreat,0,PRien,6) ;   // Stats Ressource
   74499 : ParamTable('AFTLIBRERESA',taCreat,0,PRien,6) ;   // Stats Ressource
   74501 : AppelTitreLibre('AT',Prien);  // Libellé tables libres article
   74502 : AppelTitreLibre('AM',Prien);  // Libellé mtt libres article
   74503 : AppelTitreLibre('AD',Prien);  // Libellé dates libres article
   74504 : AppelTitreLibre('AC',Prien);  // Libellé textes libres article
   74505 : AppelTitreLibre('AB',Prien);  // Libellé bool libres article
   74441..74449 : ParamTable('GCLIBREART'+IntToStr(Num-74440),taCreat,0,PRien,6) ;    // Stats Article
   74450 : ParamTable('GCLIBREARTA',taCreat,0,PRien,6) ;    // Stats Article
   74511 : AppelTitreLibre('ET',Prien);  // Libellé tables libres etablissement
   74512 : AppelTitreLibre('EM',Prien);  // Libellé mtt libres etablissement
   74513 : AppelTitreLibre('ED',Prien);  // Libellé dates libres etablissement
   74514 : AppelTitreLibre('EC',Prien);  // Libellé textes libres aetablissement
   74515 : AppelTitreLibre('EB',Prien);  // Libellé bool libres etablissement
   74470..74478 : ParamTable('YYLIBREET'+IntToStr(Num-74469),taCreat,AfNoAideTablette('YYLIBREET'+IntToStr(Num-74469)),PRien,6,RechDom('GCZONELIBRE','ET'+IntToStr(Num-74469),FALSE)) ;
   74479 : ParamTable('YYLIBREETA',taCreat,AfNoAideTablette('YYLIBREETA'),PRien,6,RechDom('GCZONELIBRE','ETA',FALSE)) ;
   74551 : AppelTitreLibre('FT',Prien);  // Libellé tables libres fournissuer
   74552 : AppelTitreLibre('FM',Prien);  // Libellé mtt libres fournisseur
   74553 : AppelTitreLibre('FD',Prien);  // Libellé dates libres fournisseur
   74561..74563 : ParamTable('GCLIBREFOU'+IntToStr(Num-74560),taCreat,AfNoAideTablette('GCLIBREFOU'+IntToStr(Num-74560)),PRien,6,RechDom('GCZONELIBRE','FT'+IntToStr(Num-74560),FALSE)) ;  // Stats clients
  else if not TraiteMenuSpecif(Num) then
    HShowMessage('2;?caption?;' + TraduireMemoire('Fonction non disponible : ') + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT, Range: string);
var ii: integer;
  Arg5LanceFiche: string;
TypeArticle,TypeTiers:string;
ArticleAff:string;
ChampMul,ValMul,Critere : string;
Arguments : String;
x: Integer;

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
//    7: {Article} DispatchTTArticle(Action, Lequel, TT, Range);
    7 : BEGIN {article}
        TypeArticle := '';
        if not IsCodeArticleUnique(Lequel) then Lequel:=CodeArticleUnique(Lequel,'','','','','') ;
        // Dans le cas de la création, il ne faut pas passer le code article en entrée
        if (Action=taCreat) then ArticleAff:='' else ArticleAff:=Lequel;

        if ArticleAff <> '' then
           TypeArticle := GetChampsArticle(Lequel, 'GA_TYPEARTICLE')
        else
        begin
           Arguments := Range;
           Repeat
               Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
               if Critere<>'' then
               begin
                  x:=pos('=',Critere);
                  if x<>0 then
                  begin
                     ChampMul:=copy(Critere,1,x-1);
                     ValMul:=copy(Critere,x+1,length(Critere));
                     if ChampMul='TYPEARTICLE' then TypeArticle:=ValMul ;// AC:traitment mar ou nom
                  end;
               end;
           until  Critere='';
        end;
        // MODIF LS POUR CORRECTION Suite a fiche de transmission
        if TypeArticle = '' then TypeArticle := 'MAR';   // Evite de lancer la fiche affaire
        // --
        Arg5LanceFiche:= ActionToString(Action);
        // Si on n'a ni article, ni type d'article, on suppose que l'on passe l'info manquante dans le range
        if (ArticleAff='') and (TypeArticle='') then
            Arg5LanceFiche:= Arg5LanceFiche + ';' + Range
        else
            begin
            Arg5LanceFiche:= Arg5LanceFiche + ';TYPEARTICLE='+TypeArticle;
            if (Range<>'') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
            end;

        // PIL le 16/06/2000 car le monofiche ne convient pas aux MUL
        if (IsArticleSpecif(TypeArticle)='FicheAffaire') then
            AGLLanceFiche('AFF', 'AFARTICLE', '', ArticleAff, Arg5LanceFiche)
        else
            // Modif BTP
            if (IsArticleSpecif(TypeArticle)='FicheBtp') then
                begin
                if TypeArticle = 'POU' then
                    AGLLanceFiche('BTP', 'BTARTPOURCENT', '', ArticleAff, Arg5LanceFiche)
                else if TypeArticle = 'FRA' then
                    AGLLanceFiche('BTP', 'BTPRESTATION', '', ArticleAff, Arg5LanceFiche)
                else if TypeArticle = 'PRE' then
                    AGLLanceFiche('BTP', 'BTPRESTATION', '', ArticleAff, Arg5LanceFiche)
                else
                    AGLLanceFiche('BTP', 'BTARTICLE', '', ArticleAff, Arg5LanceFiche);
                end
            else AGLLanceFiche('GC', 'GCARTICLE', '', ArticleAff, Arg5LanceFiche) ;
        END ;

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
  TraiteMenusPCP(NumModule);
  ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
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
  RTMenuInitApplication;
  ProcCalcMul := RTProcCalcMul;
  ProcCalcEdt := GCCalcOLEEtat;
  FMenuG.OnDispatch := Dispatch;
  FMenuG.OnChargeMag := ChargeMagHalley;
  FMenuG.OnMajAvant := nil;
  FMenuG.OnMajApres := nil;
  FMenuG.OnMajPendant := nil;
  FMenuG.OnAfterProtec := AfterProtecGC;
  FMenuG.SetSeria(GCCodeDomaine, GCCodesSeria, GCTitresSeria);
  FMenuG.OnChangeModule := AfterChangeModule;
  V_PGI.DispatchTT := DispatchTT;
  FMenuG.SetPreferences(['Pièces'], False);
  FMenuG.OnChargeFavoris := ChargeFavoris;
end;

initialization
  InitLaVariablePGI;

end.

