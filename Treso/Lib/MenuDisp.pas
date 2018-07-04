{***********UNITE*************************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Dans cette unité, on trouve LA FONCTION qui sera
Suite ........ : appelée par l'AGL pour lancer les options.
Suite ........ : Ainsi que la gestion de l'HyperZoom, de la gestion de
Suite ........ : l'arobase, ...
Suite ........ : C'est aussi dans cette unité que l'on défini le fichier ini
Suite ........ : utilisé, le nom de l'application, sa version, que l'on lance la
Suite ........ : sérialisation, les différentes possibilités d'action sur la mise à
Suite ........ : jour de structure, ...
Mots clefs ... : IMPORTANT;STRCTURE;MENU;SERIALISATION
*****************************************************************}
unit MenuDisp;

interface
Uses
  HEnt1,
  EntPGI,
  Ent1, hCtrls,
  Constantes, ComCtrls, HOut,
{$IFDEF EAGLCLIENT}
  MenuOLX,MaineAGL,eTablette,CEGIDIEU, 
{$ELSE}
  MenuOLG,Tablette,FE_Main,
  UEdtComp,
{$ENDIF EAGLCLIENT}
  Forms,sysutils,HMsgBox, Classes, Controls, HPanel, UIUtil,
  AGLInit, ParamSoc, UtilSoc, Reseau,
  TofCalculSolde ,
  TofEchellesInterets	,
  TofEquilibrage ,
  TofFicheSuivi	 ,
  TofGraphChancel,
  TofGraphTaux	 ,
  TofLanceEtat	 ,
  TofLanceJournaux,
  TofTaux	 ,
  TomAgence	 ,
  UTomBanque     ,
  TomCalendrier	 ,
  TomCatTransac	 ,
  TomCIB	 ,
  TomCourtsTermes,
  TomFluxTreso	 ,
  TomTrancheFrais,
  TomTransac	 ,
  CPCHANCELL_TOF ,
  TRQRPREVREAL_TOF,
  TRMULINTEGREES_TOF,
  TRINTEGRECOMPTA_TOF,
  TRSYNCHRONISATION_TOF,
  TRMULSYNCHRO_TOF,
  TRCONTROLEVALEUR_TOF,
  CPSUPPRUB_TOF,
  TRMULCOMPTEBQ_TOF,
  TRMULVIREMENT_TOF,
  REGLEACCRO_TOM,
  UTOMUTILISAT,
  USERGRP_TOM,
  TOMPROFILUSER,
  UTILPGI,
  TRGRAPHSOLDE_TOF,
  TRMULVALEUR_TOF,
  TRMULFINPLAC_TOF,
  TRMULDECOUVERT_TOF,
  TRMULEQUILIBRAGE_TOF,
  TRQRVIREMENT_TOF,
  TRMULPREVISIONNEL_TOF,
  TRMULFLUXTRESO_TOF,
  TRQREXTRAITBQ_TOF,
  TRQRRELEVEBQ_TOF,
  TRPREVMENSUEL_TOF,
  TRPREVANNUEL_TOF,
  TRANALYSEECRITURE_TOF,
  TRANALYSEVIREMENT_TOF,
  TRANALYSETRANSACT_TOF,
  TRCUBEECRITURE_TOF,
  TRCUBEVIREMENT_TOF,
  TRCUBETRANSACT_TOF,
  TRSUPPRECRITURE_TOF,
  CONFIDENTIALITE_TOF,
  TRSUIVISOLDE_TOF,
  TRSUIVITRESO_TOF,
  TRMULCIB_TOF,
  CPRAPPRODET_TOF,
  TRMULCONTRAT_TOF,
  TRCONTRAT_TOM,
  TRMULFRAIS_TOF,
  TRQRCAHIERTARIFS_TOF,
  TRQRJALCOMMISSION_TOF,
  TRSUIVICOMMISSION_TOF,
  uMultiDossier,
  //uMultiDossierUtil,
  GrpFavoris,
  BANQUECP_TOM,
  PAYS_TOM,
  TRMODIFIERUBRIQUE_TOF,
  Commun,
  TRTYPEFRAIS_TOT,
  TRMULOPCVM_TOF,
  TRTRANSACOPCVMACH_TOF,
  TRTRANSACOPCVMVEN_TOF,
  TRVENTEMULTIOPCVM_TOF,
  TROPCVMREF_TOM,
  TRCOTATIONOPCVM_TOF,
  TRSUIVIOPCVM_TOF,
  TRGRAPHOPCVM_TOF,
  TRQRJALOPCVM_TOF,
  TRMULVENTEOPCVM_TOF,
  TRPORTEFEUILLE_TOT,
  TRQRACHATOPCVM_TOF,
  TRQRVENTEOPCVM_TOF,
  TRSUIVIGROUPE_TOF,
  TRREMISEAZERO_TOF {10/10/06 : FQ 10331},
  TRMULICC_TOF, {12/10/06 : Gestion des Interêts de comptes courants}
  TRMULTAUXICC_TOF,{17/10/06 : Gestion des Interêts de comptes courants}
  TRQRJALCC_TOF, {17/10/06 : Journal des comptes courants}
  TRREPAREDOSSIER_TOF, {07/11/06 : Réparation des dossier}
  TRCONTROLESOLDE_TOF, {23/11/06 : Contrôle des soldes}
  uTOFPointageMul,  {06/03/07 : Nouveau pointage : Pointage}
  CPMULCETEBAC_TOF, {06/03/07 : Nouveau pointage : Mul des relevés bancaires}
  CPMULEEXBQLIG_TOF,{06/03/07 : Nouveau pointage : Mul de mouvement bancaires}
  dpTOFCETEBAC, {06/03/07 : Nouveau pointage : Intégration des relevés}
  TomIdentificationBancaire, {International}
  TRGUIDE_TOM, {29/03/07 : guides de saisie automatiques}
  ULibPointage,
  {$IFDEF TRCONF}
  TRPARAMGENERAUX_TOF,
  TRMULCONFIDENTIEL_TOF,
  {$ENDIF TRCONF}
  TRQRJALPORTEFEUIL_TOF,
  TRADDCASHPOOLING_TOF,
  TomTypeFlux, uProcEtat, ImgList;


Procedure InitApplication ;
procedure InitSerialisation;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1 : TImageList;
  end ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
procedure TripotageMenuTRESO(const TNumModule : array of integer; var stAExclure : string);
procedure ChargeObjetExercice;
procedure LibereObjetExercice;

Var
  FMenuDisp : TFMenuDisp;


implementation

uses
  BobGestion,
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  {$ENDIF MODENT1}
  {$IFDEF EAGLCLIENT}
  UtileAgl,
  {$ELSE}
  MajTable,
  {$ENDIF EAGLCLIENT}
  uRecupDos, LicUtil, PwdDossier, UProcEcriture;

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette fonction est appellée par l'AGL à chaque sélection
Suite ........ : d'une option de menu, en lui indiquant le TAG du menu en
Suite ........ : question. Ce qui déclenche l'action en question.
Suite ........ : L'AGL lance aussi cette fonction directement afin d'offrir à
Suite ........ : l'application la possibilité d'agir avant ou après la connexion,
Suite ........ : et avant ou après la déconnexion.
Suite ........ : Cette fonction prend aussi en paramètre retourForce et
Suite ........ : SortieHalley. Si RetourForce est à True, alors l'AGL
Suite ........ : remontera au niveau des modules, si SortieHalley est à
Suite ........ : True, alors ...
Mots clefs ... : MENU;OPTION;DISPATCH
*****************************************************************}
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
  stAExclureFavoris : string;
begin
   case Num of
     10 : //Apres connection
          begin
            
            {On s'assure que l'on est sur la base de Trésorerie}
            if (GetParamSocSecur('SO_TRBASETRESO', '') <> '') and
               (GetParamSocSecur('SO_TRBASETRESO', '') <> V_PGI.SchemaName) then begin
              PGIError(TraduireMemoire('Pour utiliser la trésorerie, veuillez vous connecter à la base ' + GetParamSocSecur('SO_TRBASETRESO', '')));
              {On se remet sur la mire de la connection}
              FMenuG.FermeSoc;
              Exit;
            end;

            {Chargement de V_PGI.NoDossier}
            ChargeNoDossier;

            VG_ObjetExo.ChargeObjet;

            {Moulinette pour remplacer la valeur des champs _GENERAL par BQ_CODE}
            if not GetParamSocSecur('SO_TRMOULINEBQCODE', False, True) then begin
              if PGIAsk(TraduireMemoire('La nouvelle version de la Trésorerie nécessite un traitement'#13 +
                                        'de mise à jour des structures qui peut s''avérer long.'#13#13 +
                                        'Souhaitez-vous poursuivre ?')) = mrYes then begin
                if not MonoToMultiSoc then begin
                  FMenuG.ForceClose := True;
                  FMenuG.FermeSoc;
                  {FMenuG.Quitter; 21/09/07 : Provodque en eAGL provoque une erreur de transtypage : remplacer par Close}
                  FMenuG.Close;
                  Exit;
                  Exit;
                end
                else
                  PGIInfo(TraduireMemoire('La mise à jour s''est correctement effectuée.'));
              end
              else begin
                FMenuG.ForceClose := True;
                FMenuG.FermeSoc;
                {FMenuG.Quitter; 21/09/07 : Provodque en eAGL provoque une erreur de transtypage : remplacer par Close}
                FMenuG.Close;
                Exit;
              end;
            end;
          end;
     11 : ; //Après deconnection
     12 : ; //Avant connection ou seria
     13 : ; //Avant deconnection
     15 : ; //Avant formshow
     16 : begin
            { Importation des bobs CTS5 }
            BOB_IMPORT_PCL_STD('CTS5;CCS5;', 'Trésorerie;Comptabilité;');
          end; //Avant le 10
     31 : ; {WebSat}
     32 : ; {Téléchargement}
    100 : begin {FQ 10183}
            if not (ctxPCL in V_PGI.PGIContexte) then begin
              { vérif du mot de passe lors d'une connexion automatique sur un dossier PCL protégé }
              if not CheckPwdDossier( V_PGI.NoDossier, RecupPwdDossierFromLgcde) then begin
                FMenuG.FermeSoc;
                FMenuG.Close;
                exit;
              end;
              { On force le mode CascadeForms }
              ForceCascadeForms;
              { Vérifie si le dossier a été paramétré}
              if (not IsFlagAppliOk) or
                 (not ExisteSQL('SELECT * FROM EXERCICE')) then begin
                HShowMessage('0;' + '' + ';Le dossier n''est pas paramétré.'#13 +
                             'Veuillez passer par la comptabilité pour le configurer.;W;O;O;O;', '', '');
                FMenuG.ForceClose := True;
                FMenuG.Close;
                Exit;
              end;
            end;
          end;
   1104 : ParamSociete(False,'','SCO_TRTRESORERIE;','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,500014);
   1750 : ParamRegroupementMultiDossier; {Multi-dossiers}
   3133 : begin
            stAExclureFavoris := '';
            TripotageMenuTRESO([130, 131, 132, 133, 134, 135, 136], stAExclureFavoris);
            ParamFavoris('134;135', stAExclureFavoris, False, True) ;
          end ;
//-------------
// Fiches Tréso
//-------------
      {$IFDEF JOHN}
      111111 : AGLLanceFiche('CP','RLVETEBAC','','',';;;I');
      {$ENDIF JOHN}
//Analyse Menu 130
      //Previsions 130000
      130099 : TRLanceFiche_FicheSuivi('TR','TRFICHESUIVI','','',''); {Menu Popup}
      130098 : TRLanceFiche_Previsionnel('TR','TRMULPREVISIONNEL','','', ''); {Menu Popup}
      {JP 05/10/05 : FQ 10301 : le suivi des dates de valeur est ajouté dans le Popup}
      130097 : TRLanceFiche_ControlValeur('TR','TRCONTROLEVALEUR', '', '', ''); {JP 08/09/2003 : Contrôle des dates de valeur et de pointage}

      130051 : TRLanceFiche_TRPREVANNUEL ('TR','TRPREVANNUEL','','', ''); {JP 15/10/2003 : Prévisions annuelles}
      130055 : TRLanceFiche_TRPREVMENSUEL('TR','TRPREVMENSUEL','','', ''); {JP 15/10/2003 : Prévisions mensuelles}
      130059 : TRLanceFiche_Previsionnel('TR','TRMULPREVISIONNEL','','', ''); {JP 15/10/2003 : Saisie d'écritures budgétaires}
      //Suivi 130100
      130101 : TRLanceFiche_FicheSuivi('TR','TRFICHESUIVI'  , '', '', ''); //Suivi en valeur
      130105 : TRLanceFiche_TRSUIVISOLDE('TR','TRSUIVISOLDE', '', '', ''); {suivi des soldes}
      130110 : TRLanceFiche_TRSUIVITRESO('TR','TRSUIVITRESO', '', '', ''); {suivi de trésorerie / comptabilité}
      130115 : TRLanceFiche_TRSUIVISOLDE('TR','TRSUIVISOLDE', '', '', 'SUIBQE'); {12/10/04 : suivi bancaire}
      130120 : TRLanceFiche_TrSuiviCommission('TR','TRSUIVICOMMISSION', '', '', ''); {29/07/04 : Suivi des commissions}
      130130 : TRLanceFiche_SuiviPortefeuille('TR','TRSUIVIOPCVM', '', '', ''); {04/11/04 : Suivi des Portefeuilles}
      {JP 05/10/05 : FQ 10301 : le suivi des dates de valeur est Déplacé de 133 à 130}
      130150 : TRLanceFiche_ControlValeur('TR','TRCONTROLEVALEUR', '', '', ''); {JP 08/09/2003 : Contrôle des dates de valeur et de pointage}
      {08/08/09 : mise en place de la fiche de suivi de groupe}
      130160 : TRLanceFiche_TRSUIVIGROUPE('');

      //Echelle d'intéret 130300
      130301 : TRLanceFiche_EchellesInterets('TR','TRECHELLESINT','','',''); //Echelle d'intérêts
      // Analyses Stat 130400
      130401 : TRLanceFiche_TRANALYSEECRITURE('TR', 'TRANALYSEECRITURE', '', '', ''); {25/11/03 : Analyse sur la table TRECRITURE}
      130405 : TRLanceFiche_TRANALYSETRANSACT('TR', 'TRANALYSETRANSACT', '', '', ''); {25/11/03 : Analyse sur la table TRTRANSACT}
      130410 : TRLanceFiche_TRANALYSEVIREMENT('TR', 'TRANALYSEVIREMENT', '', '', ''); {25/11/03 : Analyse sur la table EQUILIBRAGE}
      // Analyses Cube 130500
      130501 : TRLanceFiche_TRCUBEECRITURE('TR', 'TRCUBEECRITURE', '', '', ''); {26/11/03 : Cube sur la table TRECRITURE}
      130505 : TRLanceFiche_TRCUBETRANSACT('TR', 'TRCUBETRANSACT', '', '', ''); {26/11/03 : Cube sur la table TRTRANSACT}
      130510 : TRLanceFiche_TRCUBEVIREMENT('TR', 'TRCUBEVIREMENT', '', '', ''); {26/11/03 : Cube sur la table EQUILIBRAGE}
      // Analyses Stat 130400 Mutli-dossiers
      130411 : TRLanceFiche_TRANALYSEECRITURE('TR', 'TRANALYSEECRITMD', '', '', ''); {24/06/04 : Analyse sur la table TRECRITURE}
      130415 : TRLanceFiche_TRANALYSETRANSACT('TR', 'TRANALYSETRANSAMD', '', '', ''); {24/06/04 : Analyse sur la table TRTRANSACT}
      130420 : TRLanceFiche_TRANALYSEVIREMENT('TR', 'TRANALYSEVIREMTMD', '', '', ''); {24/06/04 : Analyse sur la table EQUILIBRAGE}
      // Analyses Cube 130500 Mutli-dossiers
      130511 : TRLanceFiche_TRCUBEECRITURE('TR', 'TRCUBEECRITUREMD', '', '', ''); {24/06/04 : Cube sur la table TRECRITURE}
      130515 : TRLanceFiche_TRCUBETRANSACT('TR', 'TRCUBETRANSACTMD', '', '', ''); {24/06/04 : Cube sur la table TRTRANSACT}
      130520 : TRLanceFiche_TRCUBEVIREMENT('TR', 'TRCUBEVIREMENTMD', '', '', ''); {24/06/04 : Cube sur la table EQUILIBRAGE}

//Editions Menu 131
      //Document bancaire 131100
      131099 : TRLanceFiche_PrevRealise('TR','TRETATPREVREAL','','',''); {Menu Popup}

      131511 : TRLanceFiche_EtatsCourtsTermes('TCT', tye_TicketOpe); {Ticket d'Opération}
      131513 : TRLanceFiche_EtatsCourtsTermes('TCL', tye_LettreCon); {Lettre de confirmation}
      131515 : TRLanceFiche_EtatsCourtsTermes('TCO', tye_OrdrePaie); {Ordre de paiement}

      131521 : TRLanceFiche_EtatsAchatsOPCVM('TAT', tye_TicketOpe); {Ticket d'opération d'achat d'OPCVM}
      131523 : TRLanceFiche_EtatsAchatsOPCVM('TAL', tye_LettreCon); {Lettre de confirmation d'achat d'OPCVM}
      131525 : TRLanceFiche_EtatsAchatsOPCVM('TAO', tye_OrdrePaie); {Ordre de paiement d'achat d'OPCVM}

      131531 : TRLanceFiche_EtatsVenteOPCVM('TVT', tye_TicketOpe); {Ticket d'opération de vente d'OPCVM}
      131533 : TRLanceFiche_EtatsVenteOPCVM('TVL', tye_LettreCon); {Lettre de confirmation de vente d'OPCVM}
      131535 : TRLanceFiche_EtatsVenteOPCVM('TVO', tye_OrdrePaie); {Ordre de paiement de vente d'OPCVM}

      131541 : TRLanceFiche_EtatVirement('TR','TRQRVIREMENT','','','CONFI'); {Lettre de confirmation}
      131545 : TRLanceFiche_EtatVirement('TR','TRQRVIREMENT','','','ORDRE'); {Ordre de virement}

      131114 : TRLanceFiche_Extrait('TR','TRQREXTRAITBQ','','',''); // Relevé bancaire
      131115 : TRLanceFiche_Releve('TR','TRQRRELEVEBQ','','',''); // Relevé de compte
      131120 : TRLanceFiche_CahierTarifs('TR','TRQRCAHIERTARIFS','','',''); {Cahiers des tarifs}

      {Tableaux 131200}
      131201 : AglLanceFiche('TR','TRCREATION','','',''); // Echéanciers
      131205 : TRLanceFiche_PrevRealise('TR','TRETATPREVREAL','','',''); // Etat  prévisionnel ou réalisé
      131206 : TRLanceFiche_HistoSolde('TR', 'TRGRAPHSOLDE', '', '', ''); // Histogramme du solde général de tous les comptes

      {Journaux 1314000}
      131401 : TRLanceFiche_LanceJourneaux('TR','TRLANCEJOURNAUX', '', '', 'TRJ;JTR'); //Journal des mouvements bancaires
      131410 : TRLanceFiche_JalCommission('TR', 'TRQRJALCOMMISSION', '', '', ''); {29/07/04 : Journal des commissions}
      131420 : TRLanceFiche_JalVenteOPCVM('TR', 'TRQRJALOPCVM', '', '', ''); {04/11/04 : Journal des ventes d'OPCVM}
      131430 : TRLance_JournalCC('CC;');
      131440 : TRLanceFiche_JalPortefeuille('', '', '');

      {24/06/04 : Multi-Dossiers 131300}
      131301 : TRLanceFiche_EtatVirement('TR','TRQRVIREMENTMD','','','CONFI'); {Lettre de confirmation}
      131302 : TRLanceFiche_EtatVirement('TR','TRQRVIREMENTMD','','','ORDRE'); {Ordre de virement}
      131304 : TRLanceFiche_Extrait('TR','TRQREXTRAITBQMD','','',''); {Relevé bancaire}
      131305 : TRLanceFiche_Releve('TR','TRQRRELEVEBQMD','','',''); {Relevé de compte}
      131351 : TRLanceFiche_PrevRealise('TR','TRETATPREVREALMD','','',''); {Etat  prévisionnel ou réalisé}
      131352 : TRLanceFiche_LanceJourneaux('TR','TRLANCEJOURNAUXMD', '', '', 'TRJ;JTR'); {Journal des mouvements bancaires}

//Transaction Financière Menu 132
      //Financement 132100
      132099 : TRLanceFiche_CourtsTermes ('TR','TRTRANSACTION','','', 'TCT;D'); {Menu Popup; TCT : Court terme, D : débit}
      132101 : TRLanceFiche_CourtsTermes('TR','TRTRANSACTION','','','TCT;D'); {TCT : Court terme, D : débit}
      132105 : AglLanceFiche('TR','TRCREATION','','','MLTFIN'); // Moyen long termes
      //Placement 132200
      132201 : TRLanceFiche_CourtsTermes('TR','TRTRANSACTION','','','TCT;C'); {TCT : Court terme, C : crédit}
      132205 : AglLanceFiche('TR','TRCREATION','','','MLTPLA'); // Moyen long termes
      {OPCVM 132300}
      132301 : TRLanceFiche_TransacOpcvmAch('TR', 'TRTRANSACOPCVMACH', '', '', ''); {24/11/04 : Achat d'OPCVM}
      132320 : TRLanceFiche_TransacOpcvmVen('TR', 'TRTRANSACOPCVMVEN', '', '', ''); {24/11/04 : Création de ventes simples / partielles d'OPCVM}
      132340 : TRLanceFiche_VenteMultiOpcvm('TR', 'TRVENTEMULTIOPCVM', '', '', ''); {15/12/04 : Création de ventes multiples d'OPCVM}
      132360 : TRLanceFiche_MulVenteOPCVM  ('TR', 'TRMULVENTEOPCVM'  , '', '', ''); {23/11/04 : Réalisation de vente d'OPCVM}

// Gestion bancaire
      //Equilibre 133000
      133099 : TRLanceFiche_Equilibrage ('TR','TRMULSOLDE','','', ''); {Menu Popup}
      133001 : TRLanceFiche_Equilibrage('TR','TRMULSOLDE','','',''); //Equilibrage
      133010 : TRLanceFiche_MulVirements('TR', 'TRMULVIREMENT', '', '', ''); {JP 02/10/2003 : Mul des virements effectués}
      133020 : AglLanceFiche('TR','TRCREATION','','',''); //Equilibre structure
      // Rapprochement 133100
      133101 : CPLanceFiche_ImportCEtebac('');
      133103 : CPLanceFiche_MulCEtebac('');
      133105 : CPLanceFiche_MulEexBqLig('', '', '');
      133109 : CPLanceFiche_PointageMul(CODENEWPOINTAGE + ';');
      (*
      133101 : AGLLanceFiche('CP', 'RLVINTEGRE', '', '', '');// en attendant CPLanceFiche_RecupReleve;
      133105 : CP_LancePointageAuto;
      133110 : CPLanceFiche_DPointageMul;
      *)
      { Attention à ne pas réutiliser ce Tag, car les droits figures dans les Bases anciennes
      133115 : TRLanceFiche_ControlValeur('TR','TRCONTROLEVALEUR', '', '', ''); {JP 08/09/2003 : Contrôle des dates de valeur et de pointage}
      133120 : CC_LanceFicheEtatRapproDet; {JP 12/03/04 : Nouvel état de rapprochement Bancaire}
      133130 : TRLanceFiche_Guide('', '', ''); {29/03/07 : guides de saisie automatiques lors du pintage}

      {12/10/06 : 133200 : Gestion des intérêts de comptes courants}
      133210 : TRLanceFiche_ControlValeur('TR','TRCONTROLEVALEUR', '', '', 'C/C');
      133220 : TRLanceFiche_MulIcc('');
      133230 : TRLance_FicheTauxIcc('');
      133240 : TRLance_JournalCC('ICC;');

//Paramètres Trésorerie Menu 134
      //Banques 134100
      134099 : TRLanceFiche_BanqueCP ('TR','TRMULCOMPTEBQ','','', ''); {Menu Popup}
      134098 : TRLanceFiche_MulEquilibrage('TR','TRMULEQUILIBRAGE','','',''); {Menu Popup}
      134101 : TRLanceFiche_Agence('TR','TRAGENCE','','','');   //Agence
      134103 : LanceFicheIdentificationBancaire;
      134105 : TRLanceFiche_Banques('TR','TRCTBANQUE','','',''); //Banque
      134110 : TRLanceFiche_Calendrier('TR','TRCALENDRIER','','',''); //CALENDRIER
      134111 : ParamTable('TTJOURFERIER', taCreat, 500011, prien); {JP 17/09/03 : Tablette des jours fériers}
      134125 : TRLanceFiche_BanqueCP('TR','TRMULCOMPTEBQ','','','');  {JP 23/09/03 : Liste des comptes bancaires}
      134131 : AglLanceFiche('YY','DEVISE','','','');   {Devise}
      134132 : TRLanceFiche_GraphChancel('TR','TRGRAPHCHANCEL','','','');   //Graph des cours de devise
      134134 : FicheChancel('', False, DebutDeMois(Date), taModif, True);{ Saisie de la chancellerie}
      134135 : OuvrePAYS;//AglLanceFiche('YY','YYPAYS','','','');   //Drapeaux / DEVISE
      //134135 : AglLanceFiche('TR','TRDRAPEAUX','','','');   //Drapeaux
      //134140 : TRLanceFiche_Regroupement('TR','TRREGROUPEMENT','','','');    //Regroupement des comptes bancaires
      134145 : ParamTable('TRTAUXREF', taCreat, 150, pRien) ;  //AglLanceFiche('TR','TRTAUXREF','','',''); // Saisie TAUX
      134146 : TRLanceFiche_CotationTaux('TR','TRTAUX','','',''); // Saisie cotation
      134148 : TRLanceFiche_GraphTaux('TR','TRGRAPHTAUX', '', '', ''); // Graph des taux
      //Conditions 134200
      134201 : TRLanceFiche_MulValeur     ('TR','TRMULVALEUR'     ,'','',''); //Condition de Valeur
      134211 : TRLanceFiche_MulDecouvert  ('TR','TRMULDECOUVERT'  ,'','',''); //Condition de découvert
      134205 : TRLanceFiche_MulEquilibrage('TR','TRMULEQUILIBRAGE','','',''); //Condition d'équilibrage
      134210 : TRLanceFiche_MulFinPlac    ('TR','TRMULFINPLAC'    ,'','',''); //Condition de Fin-plac
      //Société 134300
      134301 : SuppressionRubrique; {JP 04/03/04 : remplacemet de CPMULRUB par CPSUPPRUB}
      134115 : TRLanceFiche_MulCIB('TR', 'TRMULCIB', '','', tc_Reference);{paramétrage des cib de référence}
      134116 : TRLanceFiche_CIB('TR','TRCIB','','','ACTION=MODIFICATION;' + tc_CIB);      // Code InterBancaire
      134305 : TRLanceFiche_TypeFlux('TR','TRTYPEFLUX', '', '', ''); // type de flux
      134307 : TRLanceFiche_MulFluxTreso('TR','TRMULFLUXTRESO','','',''); // Flux
      134310 : TRLanceFiche_CatTransac('TR','TRCATTRANSAC','','',''); // Transactions
      134315 : TRLanceFiche_Transac('TR','TRTRANSACT','','',''); // Transactions
      //Général 134400
      134401 : AglLanceFiche('YY','YYCONTACT', 'TRE', '', 'TYPE=TRE;TYPE2=TRE');{JP 04/03/04 : Fiche contacts mise en place du type TRE}
      134405 : AglLanceFiche('YY','YYMODEREGL','','',''); //Mode de règlement
      134410 : AglLanceFiche('YY','YYMODEPAIE','','',''); //Mode de paiement
      134414 : TRLanceFiche_ParamRegleAccro('TR', 'TRFICHEACCROCHAGE', '', '', ''); {JP 10/10/03 : paramétrage des Règles d'accrochage }
      (* 11/08/04 FQ 10084 le Menu 134415 est fusionné avec 134116
      134415 : TRLanceFiche_MulCIB('TR', 'TRMULCIB', '', '', tc_Accrochage);{Attachement d'une Règle d'accrochage à un CIB}
      *)

      {134500 : Société}
      {$IFDEF TRCONF}
      134501 : TRLanceFiche_ParamGeneraux('');
      {$ENDIF TRCONF}
      134510 : TRLanceFiche_AddCashPooling('');

      //Commission 134600
      134601 : TRLanceFiche_MulContrat('TR', 'TRMULCONTRAT', '', '', ''); {30/06/04 : Saisie des contrats}
      134611 : LanceTot_TypeFrais(taCreat, pRien); {30/06/04 : Saisie des types de frais dans choix code}
      134621 : TRLanceFiche_MulFrais('TR', 'TRMULFRAIS', '', '', ''); {Paramétrage des frais}
      134631 : TRLanceFiche_TrancheFrais('TR','TRTRANCHEFRAIS','','',''); {Tranche de Frais}
      {134700 : gestion des OPCVM}
      134701 : TRLanceFiche_MulOPCVM('TR', 'TRMULOPCVM', '', '', ''); {03/11/04 : Création de code OPCVM}
      134710 : TRLanceFiche_TRCotationOPCVM('TR','TRCOTATIONOPCVM','','','');{03/11/04 : Saisie des cours des OPCVM}
      134720 : TRLanceFiche_OPCVMBase100('TR', 'TRGRAPHOPCVM', '', '', '');{04/11/04 : Graph comparatif des cours en base 100}
      134730 : LanceTot_Portefeuille(taCreat, pRien); {04/11/04 : Portefeuilles}

//Outils 135000
      // Utilisateur 135200
      135201 : FicheUserGrp; {!!!! 15/01/2004 : CETTE FICHE EST AMENÉE À ÊTRE MODIFIÉE !!!!!!}
      135202 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;
      135203 : YYLanceFiche_ProfilUser;
      135204 : ReseauUtilisateurs(False);
      {$IFDEF TRCONF}
      135205 : TRLanceFiche_MulConfidentialite('', '', '');
      {$ENDIF TRCONF}
      135208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','136;133;130;132;131;134;135;26;27');
      135310 : Commun.CreerAgence; {JP 10/03/04 : Création automatiques des agences bancaires}
      135315 : TRLanceFiche_ModifieRubrique('TR', 'TRMODIFIERUBRIQUE', '', '', ''); {12/10/04 : Affinage des codes rubriques}
      // Calcul des soldes
      135320 : TRLanceFiche_CalculSolde('TR','TRCALCULSOLDE','','',''); // Recalcul des soldes
      135325 : TRLanceFiche_PrepareDossier('');
      135330 : RemiseAZero;
      135340 : TRLanceFiche_TRCONTROLESOLDE('');
      135111 : AglLanceFiche('YY', 'YYMULJOBS', '', '', 'ACTION=CONSULTATION');

// Transfert / traitement
      136099 : TRLanceFiche_Synchro('TR','TRSYNCHRONISATION','','', ''); {Menu Popup}
   // Synchronisation
      136101 : TRLanceFiche_Synchro('TR','TRSYNCHRONISATION', '', '', ''); {JP 02/09/03 : Récupération des écritures comptables}
      136105 : TRLanceFiche_VisuSynchro('TR','TRMULSYNCHRO', '', '', ''); {JP 04/09/03 : Visualisation des écritures Récupérées}
   // Intégration
      136201 : TRLanceFiche_Integration('TR','TRINTEGRECOMPTA', '', '', ''); {JP 20/07/03 : Intégration en comptabilité}
      136205 : TRLanceFiche_VisuInteg('TR','TRMULINTEGREES', '', '', ''); {JP 01/08/03 : Écritures intégrées en comptabilité}
   // Traitement des écritures
      //136201 : ; ENVISAGER UN TRAITEMENT DE CHANGEMENT DE NATURE DES ECRITURES PREVISIONNNELLES
      136305 : TRLanceFiche_Suppression('TR', 'TRSUPPRECRITURE', '', '', '') {08/12/03 : suppression des écritures}


{$IFDEF EAGLCLIENT}
{$ENDIF}
	else HShowMessage('2;?caption?;Fonction non disponible : ;W;O;O;O;',TitreHalley,IntToStr(Num)) ;
	end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette fonction est appelée par l'AGL a chaque fois qu'un
Suite ........ : utilisateur click sur un bouton de paramètrage d'un combo.
Suite ........ : Le paramètre NUM est la valeur qui est affectée dans la
Suite ........ : zone Tag des tablettes.
Mots clefs ... : TABLETTE;COMBO
*****************************************************************}
Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
begin
  case Num of
    1      : FicheBanqueCPCode(Lequel, Action); {Tablette TTBanqueCp}
    900    : ParamRegroupementMultiDossier(Lequel);
    999    : ParamTable(TT, taCreat, 150, nil, 3); {Tablettes YYSERVICE et TTFONCTION dans la fiche contact}
    130301 : ;
    134401 : AglLanceFiche('YY','YYCONTACT', 'TRE', '', 'TYPE=TRE;TYPE2=TRE;');{Fiche contact}
    133305 : TRLanceFiche_TypeFlux('TR','TRTYPEFLUX', Range, Lequel, ActionToString(Action)); {Tablette : TRTYPEFLUX}
    133307 : TRLanceFiche_FluxTreso('TR','TRFLUXTRESO', Range, Lequel, ActionToString(Action)); {Tablette : TRCODEFLUX}
    (* 24/05/05 : Cette tablette n'a plus de raison d'être ...
    133308 : TRLanceFiche_FluxTreso('TR','TRFLUXTRESO', Range, Lequel, ActionToString(Action)); {Tablette : TRTFLUXTRESO}
    133309 : TRLanceFiche_FluxTreso('TR','TRFLUXTRESO', Range, Lequel, ActionToString(Action)); {Tablette : TRTFLUXTRESO}*)
    133310 : TRLanceFiche_CatTransac('TR','TRCATTRANSAC', Range, Lequel, ActionToString(Action)); {Tablette : TRCATTRANSAC}
    134102 : TRLanceFiche_Agence('TR','TRAGENCE', Range, Lequel, ActionToString(Action));   {Agences bancaires}
    133414 : TRLanceFiche_ParamRegleAccro('TR', 'TRFICHEACCROCHAGE', Range, Lequel, ActionToString(Action)); {Tablette : TRREGLEACCRO}
    134602 : TRLanceFiche_Contrat('TR','TRCONTRAT', Range, Lequel, ActionToString(Action)); {Contrat bancaire}
    134612 : LanceTot_TypeFrais(taCreat, nil);  {Type de Frais}
    134632 : TRLanceFiche_TrancheFrais('TR','TRTRANCHEFRAIS', Range, Lequel, ActionToString(Action)); {Tranche de Frais}
    134133 : AglLanceFiche('YY','DEVISE', Range, Lequel, ActionToString(Action)); {Devises}
    134701 : TRLanceFiche_OPCVMREF('TR', 'TROPCVMREF', Range, Lequel, ActionToString(Action)); {Fiche des OPCVM}
    134315 : TRLanceFiche_Transac ('TR', 'TRTRANSACT', ''{Range cf FQ AGL 11257}, Lequel, ActionToString(Action)); {Transactions}
    134715 : LanceTot_Portefeuille(taCreat, nil);//ParamTable(TT, taCreat, 150, nil, 3, 'Portefeuilles de valeurs mobilières');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure RemoveFromMenu(const iTag : integer; const bGroup : boolean; var stAExclure : string);
{---------------------------------------------------------------------------------------}

    {----------------------------------------------------------------------}
    procedure RemoveGroupTN(i : Integer; b : Boolean);
    {----------------------------------------------------------------------}
    var
      TN : TTreeNode;
    begin
      FMenuG.OutLook.RemoveGroup(i,b);
      TN := GetTreeItem(FMenuG.TreeView,i);
      if TN <> nil then TN.Delete;
    end;

    {----------------------------------------------------------------------}
    procedure RemoveItemTN(i : Integer);
    {----------------------------------------------------------------------}
    var
      HOI : THOutItem;
      TN  : TTreeNode;
    begin
      HOI := FMenuG.OutLook.GetItemByTag(i);
      if HOI <> nil then FMenuG.OutLook.RemoveItem(HOI);
      TN  := GetTreeItem(FMenuG.TreeView,i);
      if TN <> nil then TN.Delete;
    end;

begin
  if bGroup then FMenuG.RemoveGroup (iTag, True)
            else RemoveItemTN (iTag);
  stAExclure := stAExclure + IntToStr(iTag) + ';';
end;

{---------------------------------------------------------------------------------------}
procedure TripotageMenuTRESO(const TNumModule : array of integer; var stAExclure : string);
{---------------------------------------------------------------------------------------}
var
  i,
  NumModule : Integer;
begin
  for i := 0 to High(TNumModule) do begin
    NumModule := TNumModule[i];
    case NumModule of
      130 : begin
              RemoveFromMenu(130058, False, stAExclure);{Budget quotidien}
              {Le suivi du groupe n'est visible qu'en multi sociétés}
              if not IsTresoMultiSoc then
                RemoveFromMenu(130160, False, stAExclure);
            end;
      131 : begin
              RemoveFromMenu(131201, False, stAExclure);{Échéanciers}
              {09/07/07 : on retire le graph des soldes qu'il faudrait redévelopper ...}
              RemoveFromMenu(131206, False, stAExclure);
              if not IsTresoMultiSoc then
                RemoveFromMenu(131430, False, stAExclure);{24/10/06 : Journal des C/C}
            end;
      132 : begin
              RemoveFromMenu(132105, False, stAExclure);{Financement Moyen/Long - Termes}
              RemoveFromMenu(132205, False, stAExclure);{Placement Moyen/Long - Termes}
            end;
      133 : begin
              RemoveFromMenu(133020, False, stAExclure);{Equilibre structure}
              {12/10/06 : On ne gère les comptes courants que si on est en Tréso multi sociétés}
              if not IsTresoMultiSoc then RemoveFromMenu(133200, True, stAExclure);
            end;
      134 : begin
              {$IFNDEF TRCONF}
              RemoveFromMenu(134501, False, stAExclure);{Paramètres généraux}
              {$ENDIF TRCONF}
              RemoveFromMenu(134102, False, stAExclure);{Paramètres - Généraux}
              RemoveFromMenu(134120, False, stAExclure);{CodeAFB}
              RemoveFromMenu(134140, False, stAExclure);{Regroupement de comptes}
              {26/10/07 : SO_TRNOCASHPOOLING = True : on a bloqué le CashPooling dans uRecupDos,
                          car BanqueCp n'était pas partagée}
// GP le 23/06/2008 : On cache le cash
(*
              if not EstMultiSoc or GetParamSocSecur('SO_TRNOCASHPOOLING', False) then
*)
              If CacheCash Then RemoveFromMenu(134510, False, stAExclure);{Cash Pooling}
              stAExclure := stAExclure +  '134100;134200;134300;134400;134500;';
            end;
      135 : begin
              {$IFNDEF TRCONF}
              RemoveFromMenu(135205, False, stAExclure);{Les confidentialités}
              {$ENDIF TRCONF}
              stAExclure := stAExclure +  '135200;135300;135400;';
              //if (not IsTresoMultiSoc) or (not V_PGI.Superviseur) or (not (V_PGI.PassWord = CryptageSt(DayPass(Date)))) then
              RemoveFromMenu(135325, False, stAExclure);{Fusion des comptes bancaires}
            end;
      136 : begin
            end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure AfterChangeModule(NumModule : Integer);
{---------------------------------------------------------------------------------------}
var
  stAExclure : string;
begin
  UpdateSeries ;
  Case NumModule of
    130, 131, 132, 133, 134, 135, 136 : ChargeMenuPop(integer(hm18),FMenuG.DispatchX) ; // FQ 10360
    else ChargeMenuPop(NumModule,FMenuG.DispatchX) ;
  end ;
  TripotageMenuTRESO([130, 131, 132, 133, 134, 135, 135, 136], stAExclure);
end;

{---------------------------------------------------------------------------------------}
function ChargeFavoris : Boolean ;
{---------------------------------------------------------------------------------------}
begin
  AddGroupFavoris( FMenuG, '') ;
  Result := True;
end;

{---------------------------------------------------------------------------------------}
procedure AfterProtec(sAcces : string);
{---------------------------------------------------------------------------------------}
begin
  {JP 11/10/04 : en attente de plus amples informations FQ 10172}
  VH^.OkModTreso := (sAcces[1] = 'X');
  if ((not VH^.OkModTreso) and not V_PGI.VersionDemo) then
  begin
    FMenuG.SetModules([134,135], [62,74]);
    V_PGI.NumMenuPop := 0;
  end
  else
  begin
    FMenuG.SetModules([136,133,130,132,131,134,135],[90,3,10,94,122,60,49]) ;
    V_PGI.NumMenuPop := 27;
  end ;
end;

procedure InitSerialisation;
var
  sDom : string;
begin
  If Not CS3 Then
    BEGIN
    sDom:='05990011' ;
    VH^.SerProdTreso := '00155080';
    FMenuG.SetSeria(sDom,[VH^.SerProdTreso], ['Trésorerie Business Place']);
    END Else
    BEGIN
    sDom:='00396010' ;
    VH^.SerProdTreso := '00398080';
    FMenuG.SetSeria(sDom,[VH^.SerProdTreso], ['Trésorerie Business Suite']);
    END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette procédure permet d'initiliaser certaines référence de
Suite ........ : fonction, les modules des menus gérés par l'application, ...
Suite ........ :
Suite ........ : Cette procédure est appelée directement dans le source du
Suite ........ : projet.
Mots clefs ... : INITILISATION
*****************************************************************}
procedure InitApplication ;
begin
//ProcZoomEdt:=ZoomEdtEtat ;
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  ProcCalcEdt := CalcAroBaseEtat ;
  {$ENDIF}
  InitSerialisation;
  FMenuG.OnChargeFavoris := ChargeFavoris;
  FMenuG.OnDispatch:=Dispatch ;
  FMenuG.OnMajAvant:=Nil ;
  FMenuG.OnMajApres:=Nil ;
  FMenuG.OnChargeMag:=ChargeMagHalley ;
  FMenuG.OnAfterProtec:=AfterProtec ;
  V_PGI.NbColModuleButtons := 2;
  V_PGI.NbRowModuleButtons := 4;
  FMenuG.SetModules([136,133,130,132,131,134,135],[90,63,10,94,122,60,49]) ;
  FMenuG.OnChangeModule:=AfterChangeModule ;

  V_PGI.DispatchTT:=DispatchTT ;
  {$IFDEF TESTSIC}
  {$IFDEF EAGLCLIENT}
  SaveSynRegKey('eAGLHost', 'CWAS-DEV3:80', true);
  FCegidIE.HostN.Enabled := false;
  {$ENDIF EAGLCLIENT}
  {$ENDIF TESTSIC}
end ;

procedure InitLaVariablePGI;
begin
  {Version}
  HalSocIni:='CEGIDPGI.INI' ;
   { Ce nom apparaît dans le caption de l'application, c'est aussi le clef dans la BDR pour le
     stockage des informations préférence de l'application. }
//  NomHalley := 'Trésorerie S5';
   { Ce nom apparaît en bas à gauche de l'application }
 // TitreHalley := 'CEGID Trésorerie S5';
   { Précise le nom du fichier ini utilisé par l'application. Normalement CEGIDPGI.INI }
  HalSocIni := 'cegidpgi.ini';
   { Ce nom apparaît en bas à gauche de l'application }
  Copyright := '© Copyright ' + Apalatys;
   { Précise la série. Cela modifie l'affichage OutLook }
  If CS3 Then V_PGI.LaSerie := S3 Else V_PGI.LaSerie := S5;
  V_PGI.NumMenuPop := 27;

  {Généralités}
  V_PGI.VersionDemo   := True ;
  V_PGI.SAV           := False ;
  V_PGI.VersionReseau := True ;
  V_PGI.PGIContexte   := [ctxTreso] ;
  V_PGI.CegidAPalatys := False ;
  V_PGI.CegidBureau   := True ;
  V_PGI.StandardSurDP := True ;
  V_PGI.MajPredefini  := False ;
  V_PGI.MultiUserLogin:= False ;
  V_PGI.BlockMAJStruct:= True ;
  V_PGI.EuroCertifiee := True ;

  V_PGI.OutLook       := True ;
  V_PGI.OfficeMsg     := True ;
  V_PGI.ToolsBarRight := True ;

  If CS3 Then RenseignelaSerie(ExeCTS3) Else RenseignelaSerie(ExeCTS5  );

  {$IFDEF JOHN}
  {$IFDEF EAGLCLIENT}
//  V_PGI.NumVersionBase := 740;
//  V_PGI.NumVersion:='8.0.0' ;
//  V_PGI.NumBuild:='001.017';
  {$ENDIF EAGLCLIENT}
  {$ENDIF JOHN}

  ChargeXuelib;

  V_PGI.VersionReseau := False;
  V_PGI.ImpMatrix     := True;
  V_PGI.OKOuvert      := False;
  V_PGI.Halley        := True;
  V_PGI.NiveauAccesConf:=0;
  V_PGI.MenuCourant   := 0;
  V_PGI.DispatchTT    := DispatchTT ;
  V_PGI.ParamSocLast  := False;
  V_PGI.RAZForme      := True;
  V_PGI.CodeProduit   := '008';
  V_PGI.CegidUpdateServerParams:= 'www.update.cegid.fr';
  V_PGI.PortailWeb := 'http://utilisateurs.cegid.fr';
end;

{---------------------------------------------------------------------------------------}
procedure ChargeObjetExercice;
{---------------------------------------------------------------------------------------}
begin
  VG_ObjetExo := TObjetExercice.Create;
end;

{---------------------------------------------------------------------------------------}
procedure LibereObjetExercice;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(VG_ObjetExo) then FreeAndNil(VG_ObjetExo);
end;


initialization
CS3:=FALSE ;
If Pos('CTS3',ParamStr(0))>0 Then CS3:=TRUE ;
  ProcChargeV_PGI := InitLaVariablePGI ;
  ChargeObjetExercice;

finalization
  LibereObjetExercice;

end.

