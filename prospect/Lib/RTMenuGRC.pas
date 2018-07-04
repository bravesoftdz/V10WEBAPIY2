unit RTMenuGRC;

interface

uses
{$IFDEF EAGLCLIENT}
     Maineagl,eTablette,MenuOLX,
{$ELSE}
     Tablette,FE_Main, MenuOLG,
{$ENDIF}
     StdCtrls, ExtCtrls,Dialogs, {$IFNDEF GIGI} ctiAlerte, {$ENDIF} SaisUtil,  EntGC,
{$IFDEF CTI}
     CtiGrc, UTofTiersCti_Mul, UtilCti,CtiPCB,
{$ENDIF}
{$IFDEF AFFAIRE}
   UtofRTAffaire_TV ,
   ConfidentAffaire ,
   AffaireUtil,
{$ENDIF}
{$IFDEF CHR}
   EntCHR,
{$ENDIF}
   HPanel,sysutils,HMsgBox,Hctrls, Controls,
   UtilDispGC,HEnt1,AGLInit,paramsoc,UtilPgi,UtilRT,TiersUtil, UtomAction,
   UTofActions_TV, UTofReportActions,Utomactionchaine,
   Genereactions, UtofMailing, UtofOperations, utomoperation,
{$IFNDEF GRCLIGHT}
   utofprojets, UtofMailingaff,
   UtofOPERATIONS_TV, UtofCubeOPERATIONS, AssistCiblageVersOperation,
   UTOFCIBLAGE_TOF,UTOFVISAMUL_TOF,UtomCIBLAGE,
   UtofRTRECHMIXTE_TOF,UtofRTSUSPECTPROSPECT,
   UTOFLISTEPOURVISA_TOF,UTofListeVisaCon_TOF,UtofMixte,uTofRTCIBLESUS_TOF,UtofRTDUPLIMPORT,
   UTofRTTOPCIBLE,uTofRTTOPCONTACT_TOF,UtofRTTOPSUSCON_TOF,UTofRTTOPTIERS_TOF,
   utofRTTOPTIERSCON_TOF,{BASECONNAISSANCE_TOF,BASECONNAISSANCE_TOM,UtilBASECONNAISSANCE,}
{$IFNDEF BUREAU}
   UTofParSuspect,recupsuspect,TrtRecupSuspect,UtofRECUPSUSPECT,
   UTofSuspect_Mul,UtomSuspect,UtofRtSuspects_Cube,UtofRtFicheSuspect,UtofRTSUPPRSUSP,
{$ENDIF BUREAU}
{$ENDIF GRCLIGHT}
   UtofCubeActions,UtofRTTIERS_RECH,
   UtofGenere_Actions, Utofactions,EntRT, UtofChainages,
   UtofPerspective,UtofPerspective_TV,UtofRtCourrier,UtofCubePerspectives
   ,UtofCubeProspect, UtomParActions, UtomParChainages,UtofCONTACTS_TV
   ,UtofPros_ssactions,UTreeLinks,RTPlanning,UtofProspect_TV,UtofRtInfosRespons
   ,UtotRT,UtofPersp_Modiflot,UTofAct_ModifLot
	,UtofProspectCOmpl,UtofProspect_mul   ,UtofRTSuiviAction

{$ifndef BUREAU}  //mcd 06/07/2005 pour nepas tout avoir dans le bureau
   ,ZoomEdtProspect,UtilRCC,
   UtofParamCl, UtofPerspective_etat,
   UtofOper_Etat, UTofRTDoublons,UtofRTSuppressionTiers,
   UtofTypesChainages,
   Utof_RTSelectRessources,
   UtofEtiqCli,UTofRTProspect_ModifLot,UtofRTPiece_TV,UtofProdCommercial,
   UtofFicheProspect,
   utomRTEtatPlan,UtofRTParamPlanning,UtofInfosPlanning,
   UtofCubePropositions,UtofParamTableCompl,UtofPerspective_his,UtofPerspective_var,
   UtofProspectConf,UtofSuiviActivite,
   UtomPerspective
   ,UtofRtConcurrent_Mul
{$endif BUREAU}
{$IFNDEF CRM}
{$ifndef BUREAU}
,UTofGCFournisseur_Mul
{$endif BUREAU}
{$endif !CRM}
{$ifdef GIGI}
  ,UtofAfParamDP
{$else GIGI}
   ,UtofRECHDOCGED
{$ENDIF GIGI}
 ;

Procedure RTMenuDispatch (Num : Integer;PRien : THPanel);
Procedure RTMenuDispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
procedure RTMenuTraiteMenusGRC ( NumModule : integer; bNumModule : boolean = true ) ;
procedure RTMenuInitApplication;

implementation
uses UtilGC;

procedure AppelTitreLibre(Name: string; PRien: THPanel; tag: integer);
var
   intitule,TypeCl: string;
begin
  // fct qui permet de ne saisir que la partie voulue de la tablette des titres libres
  // similaire à la fct FactoriseZl de mdispgc.
  // on pourra par la suite tester Name mpour alimenter n° aide
  //Voir fonction  GCModifTitreZonesLibres (Name);
  //mcd 07/07/08 pour afficher le type de champ libre
if (Length(Name)>0) then
   Case Name[1] of
   'T' :  TypeCl:='tables' ;
{CRM_20090710_CD_010;12758}
   'M' :  TypeCl:='valeurs' ;
   'D' :  TypeCl:='dates' ;
   'C' :  TypeCl:='textes' ;
   'B' :  TypeCl:='décisions' ;
   'R' :  TypeCl:='ressources' ;
   End;
if (TypeCl<>'') then Intitule := TraduireMemoire('Libellés libres ') + TraduireMemoire(TypeCl);
  // fin mcd
  if ((Tag >= 92628) and (Tag <= 92629)) or ((Tag >= 92527) and (Tag <= 92528)) or (tag = 92620) then
  Begin
    GCModifTitreZonesLibres(Name, 'RRP');
{CRM_20090710_CD_010;12758}   // pr language CRM
    Intitule :=Intitule + ' ' + traduireMemoire ('projets');
    ParamTable('RTLIBPROJET', taModif, 111000700, Prien, 17, Intitule,0,true)
  end;
   if ((Tag >= 92648) and (Tag <= 92649)) or ((Tag >= 92547) and (Tag <= 92548)) or  (tag = 92640) then
  Begin
    GCModifTitreZonesLibres(Name, 'RRO');
{CRM_20090710_CD_010;12758} // pr language CRM
    Intitule :=Intitule + ' ' + traduireMemoire ('opérations');
    ParamTable('RTLIBOPERATION', taModif,111000216, Prien, 17, Intitule,0,true)
  end;
   if ((Tag >= 92638) and (Tag <= 92639)) or ((Tag >= 92537) and (Tag <= 92538)) or (tag = 92630) then
  Begin
    GCModifTitreZonesLibres(Name, 'RRS');
{CRM_20090710_CD_010;12758}  // pr language CRM
    Intitule :=Intitule + ' ' + traduireMemoire ('propositions');

{$IFDEF AFFAIRE}
    ParamTable('RTLIBPERSPECTIVE', taModif, 111000260, Prien, 17, Intitule,0,true)
{$ENDIF AFFAIRE}
  end;
   if ((Tag >= 92551) and (Tag <= 92555))   then
  Begin
    GCModifTitreZonesLibres(Name, 'RRC');
    Intitule :=Intitule + traduireMemoire (' actions');
{$IFDEF AFFAIRE}

    ParamTable('RTLIBACTIONS', taModif, 111000251, Prien, 17, Intitule,0,true)
{$ENDIF AFFAIRE}
  end;
   if ((Tag >= 92511) and (Tag <= 92512)) or (tag=92609) then
  Begin
    GCModifTitreZonesLibres(Name, 'RCI');
    Intitule :=Intitule + traduireMemoire (' ciblages');
    ParamTable('RTLIBCIBLAGE', taModif, 111000750, Prien, 17, Intitule,0,true)
  end;
end;


Procedure RTMenuDispatch (Num : Integer;PRien : THPanel);
var DateDebut : TDateTime;
    NumSem, AnneeSem : integer;
{$IFNDEF GRCLIGHT }
    StrSql              : String;
{$ENDIF}

begin
{$IFDEF CHR}
  if Not VH_GC.GRCSeria then
  Begin
    PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Gestion de la Relation Client');
    Exit;
  End;
{$ENDIF}
   Case Num of
   92109..92111, 92282,92284,92286,92287:
     if GetParamSocSecur('SO_RTCONFIDENTIALITE',False) = True and VH_RT.RTExisteConfident = False then
     begin
     PgiBox('Fonction indisponible. Vous avez les droits en consultation uniquement#13#10 liés au module "Relation clients"','Gestion de la Relation Client');
     Exit;
     end;
   92289:
     if (GetParamSocSecur('SO_RTCONFIDENTIALITE',False) = True) and (VH_RT.RTCreatPropositions = False) and (VH_RT.RTExisteConfident = False) then
     begin
     PgiBox('Fonction indisponible. Vous avez les droits en consultation uniquement#13#10 liés au module "Relation clients"','Gestion de la Relation Client');
     Exit;
     end;
   93130:
     if GetParamSocSecur('SO_RFCONFIDENTIALITE',False) = True and VH_RT.RFExisteConfident = False then
     begin
     PgiBox('Fonction indisponible. Vous avez les droits en consultation uniquement#13#10 liés au module "Relation fournisseurs"','Gestion de la Relation Fournisseur');
     Exit;
     end;
   end;
   Case Num of
// Fichiers
   92101 : RTLanceFiche_Prospect_Mul('RT','RTPROSPECT_MUL','','','');
   92102 : AGLLanceFiche('YY','YYCONTACTTIERS','','','MODIFICATION;GRC');
{$IFNDEF GRCLIGHT}
   92105 : RTLanceFiche_Projets('RT','RTPROJETS_MUL','','','');
{$ENDIF GRCLIGHT}   
   92107 : AGLLanceFiche('RT','RTPERSPECTIVE_MUL','','','');
   92104 : AGLLanceFiche('RT','RTACTIONS_MUL','','','');
{$IFNDEF GRCLIGHT}
   92103 : RTLanceFiche_Operations('RT','RTOPERATIONS_MUL','','','');
{$ENDIF GRCLIGHT}   
{$IFNDEF BUREAU}
   92918 : AglLanceFiche ('AFF','RESSOURCE_MUL','','','');  // ressources
{$IFNDEF GRCLIGHT}
   92106 : RTLanceFiche_Suspect_Mul('RT','RTSUSPECT_MUL','','','');
{$ENDIF GRCLIGHT}
{$ENDIF BUREAU}

   //92120 : AGLLanceFiche('RT','RTACTIONS_MULCHAI','','','');
   92130 : RTLanceFiche_Chainages('RT','RTCHAINAGES','','','GRC');

{$IFNDEF BUREAU}
   92108 : RTLanceFiche_RTConcurrent_Mul('RT','RTCONCURRENT_MUL','','','');
   92110 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;GRC') ;  // ouverture compte tiers
   92109 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;GRC') ;  // fermeture compte tiers
   92111 :
{$ifdef GIGI}   //mcd 10/09/07 14491
 //mcd 21/11/2007 plus de blocage monoposte en GI          if not ToutSeulAff then
{$else}
           if _BlocageMonoPoste(False, '', TRUE) then
{$endif}
              begin
              RTLanceFiche_RTSupprTiers('RT','RTSUPPRTIERS','','','') ;
              end;
{$ENDIF BUREAU}
 {$ifndef GIGI}
    92112 : RTLanceFiche_RECHDOCGED('RT','RTRECHDOCGEDGLO','','','');
 {$ENDIF GIGI}
{$IFNDEF GRCLIGHT}
    92113 : ;//RTLanceFiche_BaseConnaissance('RT','RBCONNAISSANC_MUL','','','');
{$ENDIF GRCLIGHT}    
// Traitements
{$IFDEF CTI}
   92205 : if GetParamSocSecur ('SO_RTCTIGESTION', False) then
             begin
             FMenuG.ptitre.caption:=TraduireMemoire('Couplage Téléphonie Informatique');
             VH_RT.AncienModeCTI:=true;
             FreeAndNil (VH_RT.ctiAlerte);
             if GetParamSocSecur('SO_RTINTERFACECTI','ISP') = 'CCA' then RTLanceCCA
               else RTLancePBC ;
             VH_RT.AncienModeCTI:=false;
             if (VH_RT.ctiAlerte = nil) and (VH_RT.CTISeria) then
                VH_RT.ctiAlerte := TFCtiAlerte.Create (nil);
             end;
{$ENDIF}
{$IFNDEF BUREAU}
{$IFNDEF GRCLIGHT}
   92212 : RTLanceFiche_Mailing('RT','RTPROS_MAILIN_FIC','','','FIC') ;
   92214 : RTLanceFiche_Mailing('RT','RTPRO_MAILIN_CONT','','','FIC') ;
   92218 : RTLanceFiche_Mailing('RT','RTLIGN_MAILIN_FIC','','','FIC') ;
   92272 : AGLLanceFiche('RT','RTSUSP_MAILIN_FIC','','','') ;
{$ENDIF GRCLIGHT}
   92284 : RTLanceFiche_RTProspect_ModifLot('RT','RTQUALITE','','','PRO;GRC') ;
   92286 : RTLanceFiche_RTProspect_ModifLot('RT','RTQUALITE','','','CON;GRC') ;
   92287 : RTLanceFiche_RTProspect_ModifLot('RT','RTQUALITE','','','CCO;GRC') ;
   92288 : RTLanceFiche_RTProspect_ModifLot('RT','RTQUALITE_SUSP','','','SUS;GRC') ;
   92290 : AGLLanceFiche('RT','RTDOUBLONS','','','') ;
   92282 : RTLanceFiche_RTProspect_ModifLot('RT','RTQUALITE','','','CLI;GRC') ;
{$IFNDEF GRCLIGHT}
   92291 : AGLLanceFiche('RT','RTSUSP_DOUBLE','','','') ;
   92270 : RTLanceFiche_RECUPSUSPECT('RT','RTRECUP_SUSPECT','','','') ;
   92135 : RTLanceFiche_RTSUPPRSUSP('RT','RTSUPPRSUSP','','','');
{$ENDIF GRCLIGHT}
   92289 : RTLanceFiche_Persp_ModifLot('RT','RTPERSP_MODIFLOT','','','GRC') ;
   92281 : RTLanceFiche_Act_ModifLot('RT','RTACT_MODIFLOT','','','GRC') ;
   92255,92256 : RTLanceFiche_Act_ModifLot('RT','RTACT_MODIFLOT','','','GRCCOMPL') ;
   92257 : RTLanceFiche_Persp_ModifLot('RT','RTPERSP_MODIFLOT','','','RPE;GRC') ;
{$ENDIF BUREAU}
{$IFNDEF GRCLIGHT}
   92232 : RTLanceFiche_Genere_Actions('RT','RTGENERE_ACTIONS','','','CLI');  // lancement opération
   92234 : RTLanceFiche_Genere_Actions('RT','RTGENERE_ACTIONS','','','CONT');  // lancement opération
   92238 : RTLanceFiche_Genere_Actions('RT','RTGENERE_ACTIONS','','','LIGN');  // lancement opération
   92222 : RTLanceFiche_Mailing('RT','RTPROS_MUL_MAILIN','','','MAI') ;
   92224 : RTLanceFiche_Mailing('RT','RTPRO_MAILIN_CONT','','','MAI') ;
   92228 : RTLanceFiche_Mailing('RT','RTLIGN_MUL_MAILIN','','','MAI') ;
   92271 : RTLanceFiche_Mailing('RT','RTSUSP_MUL_MAILIN','','','MAI') ;
{$ENDIF GRCLIGHT}
   92241 : RTLanceFiche_RtActionsLot('RT','RTACTIONS_LOT','','','CLI;PRODUITPGI=GRC') ;
   92242 : RTLanceFiche_RtActionsLot('RT','RTACTIONS_LOT','','','CONT;PRODUITPGI=GRC') ;
   92243 : RTLanceFiche_RtActionsLot('RT','RTACTIONS_LOT','','','FIC;PRODUITPGI=GRC') ;
   92294,92471,92481 : AGLLanceFiche('RT','RTACTIONS_MUL_RAP','','','') ;
   92292 : RTLanceFiche_ReportActions('RT','RTACTIONS_REPORT','','','');
   92295,92472,92482 : AGLLanceFiche('RT','RTACTIONS_CHRAP','','','') ;
   92296 : begin
           //SaveSynRegKey('RTMesActions','X',TRUE);
           V_PGI.ZoomOLE := true;  //Affichage en mode modal
           AGLLanceFiche ('RT','RTINFOSRESPONS','','','');
           V_PGI.ZoomOLE := False;  //Affichage en mode modal
           end;
   92251,92474,92484 : begin
           FMenuG.ptitre.caption:=TraduireMemoire('Agenda');
           if VH_RT.RTResponsable <> '' then
             begin
             NumSem := NumSemaine(V_PGI.DateEntree,AnneeSem);
             DateDebut := PremierJourSemaine(NumSem, AnneeSem);
           	 V_PGI.ZoomOLE := true;  //Affichage en mode modal
             ExecutePlanning ('INTERVENANT='+VH_RT.RTResponsable+'|',DateDebut,DateDebut+6,1,True,'AGENDA');
             V_PGI.ZoomOLE := False;  //Affichage en mode modal
             end
           else PgiBox('Ressource non définie','Relation Clients');
           end;
   92252,92473,92483 : AGLLanceFiche('RT','RTINFOSPLANNING','','','') ;
   92254,92475,92485 : RTLanceFiche_RTSuiviAction('RT', 'RTSUIVIACTION', '', '', '');
   92297 : RTLanceFiche_RTSuiviAction('RT', 'RTSUIVIACTION', '', '', '');
{$IFNDEF GRCLIGHT}
   92261 : RTLanceFiche_utofciblage('RT','RTCIBLAGEMUL','','','');
   92264 : RTLanceFiche_utofciblage('RT','RTCIBLAGEMUL','','','SUPPRESSION');
   92265 : RTLanceFiche_utofvisamul('RT','RTVISAMUL','','','');
   92266 : RTLanceFiche_RTRECHMIXTE('RT','RTRECHMIXTE','','','');
   92262 : RTLanceFiche_RTSUSPECTPROSPECT('RT','RTSUSPECTPROSPECT','','','');
   92999 : ParamTable('RTCORRESPIMPORT', taModif,111000214,PRien);
   92990 :
            //TJA 06/07/2007  2puration tablette correspondance pour import suspects
            begin
              V_PGI.ZoomOle   := False;
              FMenuG.PRien.Refresh;
              StrSql          := 'DELETE FROM CHOIXCOD WHERE CC_TYPE="CIM"';
              if HShowMessage ('0;' + 'RAZ correspondance pour import suspects;' +
                                      'Cette commande va vider TOTALEMENT la tablette des correspondances.#13#10' +
                                      'Voulez vous continuer ?' + ';Q;YN;N;N;','','') = MrYes then
                if HShowMessage ('0;' + 'RAZ correspondance pour import suspects;' +
                                        'Cette action va supprimer TOUTES les correspondances.#13#10' +
                                        'Voulez vous interrompre le traitement ?' + ';Q;YN;Y;Y;','','') = MrNo then
                  if HShowMessage ('0;' + 'RAZ correspondance pour import suspects;' +
                                          'Etes-vous certain de vouloir continuer ?' + ';Q;YN;N;N;','','') = MrYes then
                    if ExecuteSQL(StrSql) <> -1 then
                      PgiBox('Tablette des correspondances vidée !')
                    else
                      PgiBox('Erreur sur la suppression des correspondances');
            end;
{$IFDEF AFFAIRE}
{$IFNDEF GRCLIGHT}
   92216 : RTLanceFiche_MailingAffContactExp;
   92226 : RTLanceFiche_MailingAffContact ;
   92236 : RTLanceFiche_GenActionsAffContact;
{$ENDIF GRCLIGHT}
{$ENDIF}
{$ENDIF GRCLIGHT}

// Editions
{$IFNDEF BUREAU}
   92331 : RTLanceFiche_FicheProspect('RT','RTFICHEPROSPECT','','','') ;
   92332 : RTLanceFiche_EtiqCli('RT','RTETIQCLI','','','') ;
   92333 : AGLLanceFiche('RT','RTLISTEHIERARCHIQ','','','') ;
   92311 : AGLLanceFiche('RT','RTACTIONS_ETAT2','','','') ;
   92312 : AGLLanceFiche('RT','RTTIERS_VISITE','','','') ;
   92313 : AGLLanceFiche('RT','RTTIERS_VISITE','','','COMMERCIAL') ;
   92341 : AGLLanceFiche('RT','RTPERSP_ETAT','','','') ;
   92342 : AGLLanceFiche('RT','RTPERSP_ETABLC','','','') ;
{$IFNDEF GRCLIGHT}
   92320 : RTLanceFiche_OperEtat('RT','RTOPER_ETAT','','','') ;
   92360 : AGLLanceFiche('RT','RTOPER_ETAT_COUT','','','') ;
   92370 : RtLanceFiche_RTFIcheSuspect('RT','RTFICHESUSPECT','','','') ;
{$ENDIF GRCLIGHT}
{$ENDIF BUREAU}

// Analyses
   92402 : RTLanceFiche_PROS_SSACTIONS('RT','RTPROSSACTION_TV','','','') ;   // Suivi d'activité
   92411 : RTLanceFiche_PROSPECTS_TV('RT','RTPROSPECTS_TV','','','');
   92412 : RTLanceFiche_Actions_TV('RT','RTACTIONS_TV','','','');
   92413 : AGLLanceFiche('RT','RTPERSPECTIVE_TV','','','');
   92414 : AGLLanceFiche('RT','RTPERSPCON_TV','','','');
   92415 : AGLLanceFiche('RT','RTACTIONS_TVCHA','','','');
   92416 : RTLanceFiche_Contacts_TV('RT','RTCONTACTS_TV','','','');
   92417 : AGLLanceFiche('RT','RTPIECES_TV','GP_VENTEACHAT=VEN;GP_NATUREPIECEG=DE','','');
   92418 : AGLLanceFiche('RT','RTAFFAIRES_TV','AFF_STATUTAFFAIRE=PRO','','');
{$IFNDEF GRCLIGHT}
   92419 : RTLanceFiche_Operations_TV('RT','RTOPERATIONS_TV','','','');
{$ENDIF GRCLIGHT}
   92421 : RTLanceFiche_CubeProspect('RT','RTTIERS_CUBE','','','');
   92422 : RTLanceFiche_CubeActions('RT','RTACTIONS_CUBE','','','');
   92423 : AGLLanceFiche('RT','RTPERSP_CUBE','','','');
   92424 : AGLLanceFiche('RT','RTPERSPDETAIL_CUB','','','');
{$IFNDEF BUREAU}
{$IFNDEF GRCLIGHT}
   92425 : RTLanceFiche_RTSuspects_Cube('RT','RTSUSPECTS_CUBE','','','');
{$ENDIF GRCLIGHT}
{$ENDIF BUREAU}
{$IFNDEF GRCLIGHT}
   92426 : RTLanceFiche_CubeOperations('RT','RTOPERATIONS_CUBE','','','');
{$ENDIF GRCLIGHT}
   92440 : AGLLanceFiche('RT','RTPERSPSSDEV_TV','','','');
   92450 : RTLanceFiche_CubeProspect('RT','RTTIERS_PIECE_CUB','','','');
   92460,92461 : AGLLanceFiche('RT','RTTIERS_PILOTACTC','','','');
   92403 : if (VH_RT.RTCodeCommercial <>'') and (VH_RT.RTNomCommercial <>'') then
              AGLLanceFiche('RT','RTCOMMERCIAL_PRO','','',''); // synthese activite du mois

// Analyses Multi-dossier
   33631 : RTLanceFiche_Actions_TV('RT','RTACTIONS_TVMULTI','','','') ;
   ////// Outils /////////
{$IFNDEF BUREAU}
   92802 : CommercialToRessource;
{$ENDIF BUREAU}

   //92830 : IntegrationContactsRCC;
   //92840 : IntegrationAppelsRCC;
{$IFNDEF BUREAU}
   92850 : RTIntegrationRessourcePaie;
{$ENDIF BUREAU}
//   92850 : RTIntegrationRessource;
   //92870 : RTCreerLiensConcurrents ;


{$IFDEF TABLETTEHIE}
   92830 : ShowAdminHDTLinks(PRien);
{$ENDIF}

// Paramètres généraux
{$IFNDEF BUREAU}
   92930 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=GCTIERS');
 {$ifdef GIGI}
   92944 : AfLanceFiche_ParamDP('DP'); //MCD 04/07/2005 paramcompl pour champs DP
 {$ENDIF GIGI}
   92951..92960 : ParamTable('RTRPRLIBTABLE'+IntToStr(Num-92951),taCreat,111000203,PRien,3,RechDom('RTLIBCHAMPSLIBRES','CL'+IntToStr(Num-92951),FALSE)) ;
   92961..92985 : ParamTable('RTRPRLIBTABLE'+IntToStr(Num-92951),taCreat,111000203,PRien,3,RechDom('RTLIBCHAMPSLIBRES','CL'+Chr(Num-92951+55),FALSE)) ;
   92902..92911 : ParamTable('RTRPRLIBTABMUL'+IntToStr(Num-92902),taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','ML'+IntToStr(Num-92902),FALSE)) ;
    //mcd 20/04/07 ajout 92690 à 92699 pour nouvelles tablettes
   92690 : ParamTable('RTRPRLIBTABMULA',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLA',FALSE)) ;
   92691 : ParamTable('RTRPRLIBTABMULB',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLB',FALSE)) ;
   92692 : ParamTable('RTRPRLIBTABMULC',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLC',FALSE)) ;
   92693 : ParamTable('RTRPRLIBTABMULD',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLD',FALSE)) ;
   92694 : ParamTable('RTRPRLIBTABMULE',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLE',FALSE)) ;
   92695 : ParamTable('RTRPRLIBTABMULK',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLF',FALSE)) ;
   92696 : ParamTable('RTRPRLIBTABMULG',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLG',FALSE)) ;
   92697 : ParamTable('RTRPRLIBTABMULH',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLH',FALSE)) ;
   92698 : ParamTable('RTRPRLIBTABMULI',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLI',FALSE)) ;
   92699 : ParamTable('RTRPRLIBTABMULJ',taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPSLIBRES','MLJ',FALSE)) ;

   92945 : AGLLanceFiche('RT','RTPARAMPROSPCOMPL','','','');
   92932..92941 : ParamTable('RTRPCLIBTABLE'+IntToStr(Num-92932),taCreat,111000206,PRien,3,RechDom('RTLIBTABLECOMPL','TL'+IntToStr(Num-92932),FALSE)) ;
   92942 : ParamTable('RTRPCLIBTABLEA',taCreat,111000206,PRien,3,RechDom('RTLIBTABLECOMPL','TLA',FALSE)) ;
   92943 : ParamTable('RTRPCLIBTABLEB',taCreat,111000206,PRien,3,RechDom('RTLIBTABLECOMPL','TLB',FALSE)) ;
{$IFNDEF GRCLIGHT}
   92913 : begin
(* Mis en commentaire par BRL 15112010 car la tablette RTLIBCHPLIBSUSLIGHT n'existe pas !!
            // Pl le 23/08/07 : FQ10622 gestion des champs libres pour tous mais avec paramsoc
            if (GetParamSocSecur ('SO_RTPARAMAVANCES', false) = true ) then
              begin
                ParamTable('RTLIBCHPLIBSUSPECTS',taModif,111000207,PRien) ;
                AvertirTable ('RTLIBCHPLIBSUSPECTS');
              end
            else
              begin
                ParamTable('RTLIBCHPLIBSUSLIGHT',taModif,111000207,PRien) ;
                AvertirTable ('RTLIBCHPLIBSUSLIGHT');
              end;
*)
            ParamTable('RTLIBCHPLIBSUSPECTS',taModif,111000207,PRien) ;
            AvertirTable ('RTLIBCHPLIBSUSPECTS');
          end;
   92920..92929 : ParamTable('RTRSCLIBTABLE'+IntToStr(Num-92920),taCreat,111000208,PRien,3,RechDom('RTLIBCHPLIBSUSPECTS','CL'+IntToStr(Num-92920),FALSE)) ;
   //gestion des champs libres ciblage
   92609 :  begin
              ParamTable('RTLIBCIBLAGE',taModif,0,PRien,3,'',0,true) ;
              AvertirTable ('RTLIBCIBLAGE');
            end;
   92612..92616 : ParamTable('RTRCBTABLELIBRE'+IntToStr(Num-92611),taCreat,0,PRien,6,RechDom('RTLIBCIBLAGE','TL'+IntToStr(Num-92611),FALSE)) ;

   // gestion des champs libres Projets
   92620: begin
              ParamTable('RTLIBPROJET',taModif,0,PRien,3,'',0,true) ;
              AvertirTable ('RTLIBPROJET');
            end;
   92622..92626 : ParamTable('RTRRPJTABLELIBRE'+IntToStr(Num-92621),taCreat,0,PRien,6,RechDom('RTLIBPROJET','TL'+IntToStr(Num-92621),FALSE)) ;

   // gestion des champs libres Propositions marketing
(*PL le 27/09/07 : on passe le bloc en dehors du Ifndef GrcLight car on veut récupérer les tables libres sur 6c dans la grcLight
   92630: begin
              ParamTable('RTLIBPERSPECTIVE',taModif,0,PRien) ;
              AvertirTable ('RTLIBPERSPECTIVE');
            end;
   92632..92636 : ParamTable('RTRRPETABLELIBRE'+IntToStr(Num-92631),taCreat,0,PRien,6,RechDom('RTLIBPERSPECTIVE','TL'+IntToStr(Num-92631),FALSE)) ;
*)
   // gestion des champs libres Opérations
   92640: begin
              ParamTable('RTLIBOPERATION',taModif,0,PRien,3,'',0,true) ;
              AvertirTable ('RTLIBOPERATION');
            end;
   92642..92646 : ParamTable('RTRROPTABLELIBRE'+IntToStr(Num-92641),taCreat,0,PRien,6,RechDom('RTLIBOPERATION','TL'+IntToStr(Num-92641),FALSE)) ;


  // Infos compl Projet
   92670 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=RTPROJETS');
   92653..92657 : ParamTable('RT00QLIBTABLE'+IntToStr(Num-92653),taCreat,0,PRien,3,RechDom('RTLIBCHAMPS00Q','CL'+IntToStr(Num-92653),FALSE)) ;
   92658..92662 : ParamTable('RT00QLIBTABMUL'+IntToStr(Num-92658),taCreat,0,PRien,3,RechDom('RTLIBCHAMPS00Q','ML'+IntToStr(Num-92658),FALSE)) ;


   92914 : begin ParamTable('RTMOTIFFERMETURE',taCreat,111000209,PRien) ;
                 AvertirTable ('RTMOTIFFERMETURE'); end;
   92947 : AGLLanceFiche('RT','RTCORRESP_SUSP','','','');

   92770 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=RTOPERATIONS');
   92711..92715 : ParamTable('RT002LIBTABLE'+IntToStr(Num-92711),taCreat,0,PRien,3,RechDom('RTLIBCHAMPS002','CL'+IntToStr(Num-92711),FALSE)) ;
   92731..92735 : ParamTable('RT002LIBTABMUL'+IntToStr(Num-92731),taCreat,0,PRien,3,RechDom('RTLIBCHAMPS002','ML'+IntToStr(Num-92731),FALSE)) ;

   92916 : begin ParamTable('RTOBJETOPE',taCreat,111000211,PRien) ;
                 AvertirTable ('RTOBJETOPE'); end;

{$ENDIF GRCLIGHT}

   92946 : if GetParamSocSecur('SO_RTCONFIDENTIALITE',False) = True then
              AGLLanceFiche('RT','RTUTILISATEUR_MUL','','','')
           else
              PGIInfo('La confidentialité n''est pas gérée (paramètres société)','Confidentialité');

   { infos complémentaires des contacts }
   92765 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=YYCONTACT');
   92786..92790 : ParamTable('RT006LIBTABLE'+IntToStr(Num-92786),taCreat,0,PRien,3,RechDom('RTLIBCHAMPS006','CL'+IntToStr(Num-92786),FALSE)) ;
   92793..92797 : ParamTable('RT006LIBTABMUL'+IntToStr(Num-92793),taCreat,0,PRien,3,RechDom('RTLIBCHAMPS006','ML'+IntToStr(Num-92793),FALSE)) ;

   92815 : ParamTable('RBBOOLLIBREBC',taModif,0,PRien,3,'Décisions libres base de connaissance',0,true) ;
   92820 : ParamTable('RBMULTILIBREBC',taModif,0,PRien,3,'Multi-choix libre base de connaissance',0,true) ;
   92831 : begin ParamTable('RBMULTICHOIXBC1',taCreat,0,PRien,3,RechDom('RBMULTILIBREBC','BM1',FALSE)) ;
           AvertirTable ('RBMULTICHOIXBC1') ; end ;
   92840 : begin ParamTable('RBCATEGORIE',taCreat,0,PRien) ;
           AvertirTable ('RBCATEGORIE') ; end ;
   92841 : begin ParamTable('RBTYPEINFORMATION',taCreat,0,PRien) ;
           AvertirTable ('RBTYPEINFORMATION') ; end ;
// Paramètres Actions/Propositions

   92988 : begin AGLLanceFiche('RT','RTTYPEACTIONS','---;GRC','','GRC');;
                 AvertirTable ('RTTYPEACTION'); end;
   92915 : begin
           ParamTable('RTLIBCLACTIONS',taModif,111000251,PRien,3,'',0,true) ;
           AvertirTable ('RTLIBCHAMPSLIBRES') ; end ;
   92991..92993 : ParamTable('RTRPRLIBACTION'+IntToStr(Num-92990),taCreat,111000252,PRien,3,RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(Num-92990),FALSE)) ;
   92950 : begin ParamTable('RTIMPORTANCEACT',taCreat,111000253,PRien) ;  // Niveau d'importance des actions
                 AvertirTable ('RTIMPORTANCEACT'); end;
   92740 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=RTACTIONS');
   92701..92705 : ParamTable('RT001LIBTABLE'+IntToStr(Num-92701),taCreat,111000257,PRien,3,RechDom('RTLIBCHAMPS001','CL'+IntToStr(Num-92701),FALSE)) ;
   92721..92725 : ParamTable('RT001LIBTABMUL'+IntToStr(Num-92721),taCreat,111000258,PRien,3,RechDom('RTLIBCHAMPS001','ML'+IntToStr(Num-92721),FALSE)) ;

   92760 :
       begin
       AGLLanceFiche ('RT','RTINFOS_ACTGEN','MODELES D''ACTIONS','','MODELES;CODEOPER=MODELES D''ACTIONS');
       VH_RT.RTModActionsCTI.ClearDetail;
       end;
   92775 : begin
           ParamTable('RTLIBCHAINAGES',taModif,111000249,PRien,3,'',0,true) ;
           AvertirTable ('RTLIBCHAINAGES') ; end ;
   92781..92783 : ParamTable('RTLIBCHAINAGE'+IntToStr(Num-92780),taCreat,111000248,PRien,3,RechDom('RTLIBCHAINAGES','RH'+IntToStr(Num-92780),FALSE)) ;

   92912 : begin ParamTable('RTTYPEPERSPECTIVE',taCreat,111000259,PRien,3,'',0,true) ;
                 AvertirTable ('RTTYPEPERSPECTIVE'); end;

   // PL le 27/09/07 : on passe le bloc en dehors du Ifndef GrcLight car on veut récupérer les tables libres sur 6c dans la grcLight
   // gestion des champs libres Propositions marketing sur 6c
   92630: begin
					(*
              ParamTable('RTLIBPERSPECTIVE',taModif,0,PRien,3,'',0,true) ;
              AvertirTable ('RTLIBPERSPECTIVE');
          *)
						AppelTitreLibre('TL', Prien, num); // Libellé tables libres perspectives
            end;
   92632..92636 : ParamTable('RTRRPETABLELIBRE'+IntToStr(Num-92631),taCreat,0,PRien,6,RechDom('RTLIBPERSPECTIVE','TL'+IntToStr(Num-92631),FALSE)) ;

   // gestion des champs libres Propositions marketing sur 3c
   92750 : begin
           ParamTable('RTLIBCLPROPOSITIONS',taModif,111000260,PRien,3,'',0,true) ;
           AvertirTable ('RTLIBCHAMPSLIBRES') ; end ;
   92996..92998 : ParamTable('RTRPRLIBPERSPECTIVE'+IntToStr(Num-92995),taCreat,111000261,PRien,3,RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(Num-92995),FALSE)) ;

   92917 : begin ParamTable('RTMOTIFS',tacreat,111000262,PRien) ;
                 AvertirTable ('RTMOTIFS'); AvertirTable ('RTMOTIFSIGNATURE'); end;
  // Infos compl Proposition
   92865 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=RTPERSPECTIVES');
   92871..92875 : ParamTable('RT00VLIBTABLE'+IntToStr(Num-92871),taCreat,111000263,PRien,3,RechDom('RTLIBCHAMPS00V','CL'+IntToStr(Num-92871),FALSE)) ;
   92881..92885 : ParamTable('RT00VLIBTABMUL'+IntToStr(Num-92881),taCreat,111000263,PRien,3,RechDom('RTLIBCHAMPS00V','ML'+IntToStr(Num-92881),FALSE)) ;

   // Modèles Chainages
   92610 : RTLanceFiche_TypesChainages('RT','RTTYPESCHAINAGES','RPG_PRODUITPGI=GRC','','GRC');
// Paramètres Planning
   92791 : AGLLanceFiche('RT','RTPARAMPLANNING','','','PL2;'+V_PGI.User);
   92792 : Agllancefiche('RT','RTETATPLANNING',V_PGI.User,'',V_PGI.User );
   92798 : AGLLanceFiche('RT','RTPARAMAGENDA','','','PL1;'+V_PGI.User);
// GED
   92681 : begin
           ParamTable('RTLIBGED',taModif,0,PRien,3,'',0,true) ;
           AvertirTable ('RTLIBGED') ; end ;
   92683..92685 : ParamTable('RTLIBGED'+IntToStr(Num-92682),taCreat,0,PRien,3,RechDom('RTLIBGED','RD'+IntToStr(Num-92682),FALSE)) ;
   { GRF : menu fichiers }
   93110 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','') ;
   93120 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=FOU','','MODIFICATION') ;  // contacts
   93130 : RTLanceFiche_RTProspect_ModifLot('RT','RTQUALITE','','','ICF;GC') ;
   93140 : RTLanceFiche_Actions_Mul('RT','RFACTIONS_MUL','','','');
   93150 : RTLanceFiche_Actions_Mul('RT','RTCHAINAGES','','','GRF');
   93160 : RTLanceFiche_Operations('RT','RFOPERATIONS_MUL','','','GRF');
   { GRF : menu traitement }
   93211 : RTLanceFiche_Genere_Actions('RT','RFGENERE_ACTIONS','','','FOU;GRF');  // lancement opération
   93212 : RTLanceFiche_Genere_Actions('RT','RFGENERE_ACTIONS','','','LIGN;GRF');  // lancement opération
   93221 : RTLanceFiche_Mailing('RT','RFPROS_MUL_MAILIN','','','MAI;GRF') ;
   93231 : RTLanceFiche_Mailing('RT','RFPROS_MAILIN_FIC','','','FIC;GRF') ;
   93222 : RTLanceFiche_Mailing('RT','RFLIGN_MUL_MAILIN','','','MAI;GRF') ;
   93232 : RTLanceFiche_Mailing('RT','RFLIGN_MAILIN_FIC','','','FIC;GRF') ;
   93240 : RTLanceFiche_ReportActions('RT','RFACTIONS_REPORT','','','');
   93281 : RTLanceFiche_Act_ModifLot('RT','RFACT_MODIFLOT','','','GRF') ;
   //93240 : AFLancefiche_RecodifTiers ;
   { GRF : Editions }
   93320 : RTLanceFiche_FicheProspect('RT','RFFICHEFOUR','','','GRF') ;
   93310 : AGLLanceFiche('RT','RFACTIONS_ETAT','','','GRF');
   93330 : RTLanceFiche_OperEtat('RT','RFOPER_ETAT','','','') ;
   { GRF : Analyses }
   93411 : RTLanceFiche_PROSPECTS_TV('RT','RFPROSPECTS_TV','','','GRF');
   93412 : AGLLanceFiche('RT','RFACTIONS_TV','','','GRF');
   93413 : AGLLanceFiche('RT','RFACTIONS_TVCHA','','','GRF');   
   93421 : RTLanceFiche_CubeProspect('RT','RFTIERS_CUBE','','','GRF');
   93422 : AGLLanceFiche('RT','RFACTIONS_CUBE','','','GRF');
   93430 : RTLanceFiche_PROS_SSACTIONS('RT','RFPROSSACTION_TV','','','GRF') ;
   93440 : RTLanceFiche_CubeProspect('RT','RFTIERS_PIECE_CUB','','','GRF');
   { GRF : Paramètres généraux }
   93510 : RTLanceFiche_ParamCl('RT','RTPARAMCL','','','FICHEPARAM=GCFOURNISSEURS');
   93521..93525 : ParamTable('RT003LIBTABLE'+IntToStr(Num-93521),taCreat,111000203,PRien,3,RechDom('RTLIBCHAMPS003','CL'+IntToStr(Num-93521),FALSE)) ;
   93541..93545 : ParamTable('RT003LIBTABMUL'+IntToStr(Num-93541),taCreat,111000204,PRien,3,RechDom('RTLIBCHAMPS003','ML'+IntToStr(Num-93541),FALSE)) ;
   93560 : begin ParamTable('RFOBJETOPE',taCreat,111000211,PRien) ;
                 AvertirTable ('RFOBJETOPE'); end;
   93570 : if GetParamSocSecur('SO_RFCONFIDENTIALITE',False) = True then
              AGLLanceFiche('RT','RTUTILISATEUR_MUL','','','')
           else
              PGIInfo('La confidentialité n''est pas gérée (paramètres société)','Confidentialité');
   { GRF : Paramètres actions }
   93610 : begin AGLLanceFiche('RT','RTTYPEACTIONS','---;GRF','','GRF');;
                 AvertirTable ('RTTYPEACTIONFOU'); end;
   { tables libres actions fournisseurs }
   93620 : begin
           ParamTable('RFLIBCLACTIONS',taModif,111000251,PRien) ;
           AvertirTable ('RFLIBCLACTIONS') ; end ;
   93631..93633 : ParamTable('RFRPRLIBACTION'+IntToStr(Num-93630),taCreat,111000252,PRien,3,RechDom('RFLIBCLACTIONS','FA'+IntToStr(Num-93630),FALSE)) ;
   93650 : RTLanceFiche_TypesChainages('RT','RTTYPESCHAINAGES','RPG_PRODUITPGI=GRF','','GRF');
   { tables libres chainages fournisseurs }
   93660 : begin
           ParamTable('RFLIBCHAINAGES',taModif,111000251,PRien) ;
           AvertirTable ('RFLIBCHAINAGES') ; end ;
   93671..93673 : ParamTable('RFLIBCHAINAGE'+IntToStr(Num-93670),taCreat,111000252,PRien,3,RechDom('RFLIBCHAINAGES','RF'+IntToStr(Num-93670),FALSE)) ;
{$ENDIF BUREAU}
   end;
end;

Procedure RTMenuDispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT,Range : String ) ;
var Arg5LanceFiche : string ;
    z_V_PGI_RibbonBar : Boolean;
begin
   Case Num of
   22 : {GRC-action} AGLLanceFiche('RT','RTACTIONS','',Lequel,ActionToString(Action)+';MONOFICHE'+Range) ;
{$IFNDEF GRCLIGHT}
   23 : {GRC-operation} AGLLanceFiche('RT','RTOPERATIONS','',Lequel,ActionToString(Action)+';MONOFICHE;'+Range) ;
{$ENDIF GRCLIGHT}   
   24 : {GRC-actionsChainées}
        BEGIN
        Arg5LanceFiche:=ActionToString(Action)+';MONOFICHE;'+Range;
        //Arg5LanceFiche:='MONOFICHE;'+Range;
        AGLLanceFiche('RT','RTACTIONSCHAINE','',Lequel,Arg5LanceFiche) ;
        END;
{$IFNDEF GRCLIGHT}
   30 : {GRC-Projet} AGLLanceFiche('RT','RTPROJETS','',Lequel,ActionToString(Action)+';MONOFICHE') ;
{$ENDIF GRCLIGHT}
   31..33 : ParamTable('RTRPRLIBACTION'+IntToStr(Num-30),Action,111000252,Nil,3,RechDom('RTLIBCHAMPSLIBRES','AL'+IntToStr(Num-30),FALSE)) ;
   34..36 : ParamTable('RTRPRLIBPERSPECTIVE'+IntToStr(Num-33),Action,111000261,Nil,3,RechDom('RTLIBCHAMPSLIBRES','PL'+IntToStr(Num-33),FALSE)) ;
   37..39 : ParamTable('RTLIBCHAINAGE'+IntToStr(Num-36),Action,0,Nil,3,RechDom('RTLIBCHAINAGES','RH'+IntToStr(Num-36),FALSE)) ;
   40 : {Proposition}
        begin
        Arg5LanceFiche:=ActionToString(Action)+';MONOFICHE';
        if Range <> '' then Arg5LanceFiche := Arg5LanceFiche +';'+Range;
        AGLLanceFiche('RT','RTPERSPECTIVES','',Lequel,Arg5LanceFiche) ;
        end;
   43 : begin ParamTable('RTENSEIGNE',taCreat,110000156,Nil) ;
        AvertirTable ('RTENSEIGNE'); AvertirTable ('RTENSEIGNECODE'); end;
{$IFNDEF GRCLIGHT}
   46 : {GRC-suspect}
        {$IFDEF AFFAIRE}
        if AglJaiLeDroitFiche(['RTSUSPECTS', ActionToString(Action)], 2) then
        {$ENDIF}
            AGLLanceFiche('RT','RTSUSPECTS','',Lequel,ActionToString(Action)+';MONOFICHE'+Range) ;
{$ENDIF GRCLIGHT}
   47 : {Concurrent}
        begin
        z_V_PGI_RibbonBar := V_PGI.RibbonBar;
        V_PGI.RibbonBar := False;
        Arg5LanceFiche:=ActionToString(Action);
        if Range <> '' then Arg5LanceFiche := Arg5LanceFiche +';'+Range;
        AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche) ;
        V_PGI.RibbonBar := z_V_PGI_RibbonBar;
        end;
   end;
end;

procedure RTMenuTraiteMenusGRC ( NumModule : integer; bNumModule : boolean = true ) ;
begin
  if bNumModule then
  begin
    VH_GC.NumModuleCourant := NumModule;
    TraiteMenusGRCModule (NumModule);
  end;
  TraiteMenusGRC(NumModule, bNumModule) ;
end;

procedure RTMenuInitApplication;
begin
// ProcZoomEdt:=RTZoomEdtEtat ;
end;


end.

