unit MDispBTP;
interface

Uses Windows,Forms,HMsgBox, Classes, Hctrls, Controls, HPanel, UIUtil,ImgList,
{$IFDEF EAGLCLIENT}
     Maineagl,MenuOLX,M3FP,UtileAGL,eTablette,
{$ELSE}
     FE_Main,MenuOLG,
     EdtEtat,Tablette,ParamSoc,UtomArticle,UtilArticle,pays,region,codepost,
     EdtQR,CreerSoc,CopySoc,Reseau,Mullog,LicUtil,FactUtil,SOUCHE_TOM,
  {$IFNDEF SANSCOMPTA}
     devise,
  {$ENDIF}
{$ENDIF}
{$IFNDEF V530}
      EdtREtat,UEdtComp,GrpFavoris,
{$ENDIF}
      Confidentialite_TOF, UTOMCompetence,
      Ent1,EntGC,EntRT,EntPGI,Messages,
      ModifTableLigOuv,UTOFGCTRAITEADRESSE,
      UtilDispGC,AGLInit,HEnt1,sysutils, DicoBTP, AffaireUtil,AssistCodeAffaire,
      AssistCodeAppel,UtomUtilisat,UtilPGI,UserGrp_tom, tradarticle,menus,UtilMenuAff,
      HistorisePiece_Tof,PieceEpure_Tof,AssistSuggestion,UTofMouvStkEx,UTOFListeInvVal,UTOFGCLISTEINV_MUL,UtofConsoEtat,
      AssistRazActivite,AssistMajCompteurSouche,AssistAffecteDepot,
      VerifAuxiTiers,VerifNatErrImp,UTILULST_TOF,LGCOMPTE_TOF,UTofAfVentilana,
      UtofEtatsActivite, UTofAfJournalAct, UTofAfTableauBord, UTofAfventescube,
      UTofAfActiviteMul, utofafactivite_suivi, UTofAfActiviteCube, AGLInitBTP,
      ArtNonMvtes_Tof, UtofBTpaieGener, UTofAfActivitePai_Mul, UtofBTPaieAnnul,
      BTPrepaLivr,UtofRessource_Mul,UTOF_VideInside,UgenereAffaire,UTOFAFRESSOU_SYNCHRO,
      ChangeVersions,UTofAfPrepFact,UTOFAFVISAPIECPRO_MUL,UtofAfPiece,UTofAfPieceProAnu_Mul,
      AssistStockAjust,UtofGCOuvFerm, PARAMCATTAXE_TOF, RECONDUCTION_TOF, UtofMailing,
//Integration pour augmentation et revision de prix
      uTofAfPublicationMul,uTofAfIndiceMul,uTofAFValIndiceMul,uTofAFRevisionMul,
      uTofAfRevFormule_Mul, uTofAfRevModifCoef,utofAformuleVar_Mul,utofAformuleVar,
      utofAfRevRegulFacture, UtilRevision,utofAfRevEtatFacture,UtomAffairePiece,
      utofafRevImportIndice, utofAfRevExportIndice, utofAfRevPrepaImport,UTOFAFAUGMEN,
      UTOFAFREVCALCCOEF, UTofReportActions,BTAssistDecisionAch,BTTRARTHIER_TOF,
      UtilDuplicBordereaux,ImportInnovpro,BTCONTROLECBT_TOF,UtilOutilsLivraisons
{$IFDEF CTI}
			,CtiGRC{, BTPCTIPCB}
{$ENDIF}
//uniquement en line
			,AssistParamSocBtp
      ,OuvreExo
      ,UtilLine
      ,ShellAPI
      ,BTWebIE
      ,AssistExport
      ,BTAssistImport
      ,UIUtil3
//
			,Reindex
      ,UTob
      {$IFNDEF DBXPRESS} ,dbTables {$ELSE} ,uDbxDataSet {$ENDIF}
      ,UtilGepBTP
      ,UintegreBatiprix
      ,UCdeBatiprix
{$IFNDEF EAGLCLIENT}
//      ,BTMajStruct
			,YYBUNDLE_TOF
{$ENDIF}
      ,BtPlanning
      ,yalertes_tof
      ,UtilBlocage
      ,UTofPARCTIERS_MUL
      ,BTBLOCFONT_TOF
      ,BTSpigao
      ,UApplication
//      ,UFicheLienVP
      ,UMenusNeg
      ,UMenusGestionParc
      ,BTRechargeFamTarif
      ,UplanningBTP
      //,LiaisonOutlook
      ;

Var TOBParamSoc	: Tob;

Procedure InitApplication ;
procedure AfterChangeModule ( NumModule : integer ) ;
procedure AppelTitreLibre (Name : string;PRien : THPanel;Prefixe : string = 'ZLI' );

Procedure FermeApplication ;

type

  TFMenuDisp = class(TDatamodule)

  private

    procedure HDTDrawNode(Sender: TObject; var Caption: string);
    procedure InitializeCTI;
    procedure TerminateCTI;

  end ;

Var FMenuDisp : TFMenuDisp ;

implementation


uses Facture,(*Souche,*)dimension,UTomPiece, Traduc
            ,TarifArticle, TarifCliArt, TarifTiers, TarifCatArt, FichCOmm,Arrondi
            ,EdtDoc, BTPutil
{compta}    {,OuvrFerm,Suprauxi},InitSoc,RegTva{,General, Tiers, Journal}
{Etebac}    ,EtbMce,EtbUser{,TeleTrans}
            ,banquecp
{Negoce}    ,ListeDeSaisie{,AssistImport}
{affaires}
            ,ZoomEdtAffaire, ConfidentAffaire
            ,RecupTempoAssist

            ,AFActivite, ActiviteUtil
            ,UtilSoc
            ,UtilGc
            ,CalcOleBTP
{$IFNDEF EAGLCLIENT}
{$IFDEF GRC}
     ,RegGRCPGI
{$ENDIF}
{$ENDIF}
{$IFDEF GRC}
     ,RTMenuGRC
     ,CalcMulProspect
     ,RTRappel_Actions
     ,UTreeLinks
     ,UtofProspect_mul
{$ENDIF}
		 ,UTofRTSuppressionTiers
     ,CtiAlerte
     ,RTPlanning
     ,UCotraitance
     ,CalcMulBTP
     ,BTPGetVersions
     ,FactPiece
     ,UdroitUtilisation
     ,UTraiteTables
     ,UspecifPOC
     , BRGPDREFERENTIEL_TOF
     , BRGPDTIERSMUL_TOF
     , FormsName
     , BRGPDUtils
     , BRGPDRESSOURCEMUL_TOF
     , BRGPDUTILISATMUL_TOF
     , BRGPDSUSPECTMUL_TOF
     , BTPARAMWS_TOF
     , BTWSTABLEAUTO_TOF
     ;
{$R *.DFM}

var lesTagsToRemove : string;
    TypeFacDetail		: String;
//    PresenceSpigao  : Boolean;

function InitLesTagsToRemoveFavoris(sParam : string) : string;
begin
(*
  if      (sParam='STANDARD') then
  begin
    Result := ''
            + '92403;92850;30700;30610;30611;-30620;-30640;30609;33106;33107;-33140;-33150;'
            + '-33160;-33170;33106;33181;33182;92403;'
            {$IFNDEF CEGID}
              + '-33180;'
            {$ENDIF}
            + '36600;'
            {$IFDEF CCS3}
              + '32700;32602;33300;-33130;60206;60207;65204;65408;65305;65306;-65524;-65250;65600;92107;'
              + '92106;92108;92201;92270;92271;92272;92291;92288;-92340;92413;92414;92423;92424;92425;92440;92912;'
              + '92917;92913;92914;-92995;92946;92947;-92919;92915;92500;-92700;-92720;92740;92750;92945;-92931;92948;'
              + '-60410;-60420;'
            {$ENDIF}
            ;
      if not GetParamSoc('SO_TARIFSAVANCES') then
        Result := Result + '210400;211400;211500;211700;';
      if not GetParamSoc('SO_FRAISAVANCES') then
        Result := Result + '32603;';
      if not GetParamSoc('SO_TARIFSPRIXMARGES') then
        Result := Result + '210500;';
      if not GereFraisAvances then
        Result := Result + '211800;32000;';
      if not StkGereTrf then
        Result := Result + '32208;32209;';

    Result := Result + '2106;';
  end
  else if (sParam='CEGID') then
  begin
    Result := '33201;'
              {$IFDEF CCS3}
                + '32700;32602;33300;-33130;60206;60207;65204;65408;65305;65306;-65524;-65250;65600;92107;'
                + '92106;92108;92201;92270;92271;92272;92291;92288;-92340;92413;92414;92423;92424;92425;92440;92912;'
                + '92917;92913;92914;-92995;92946;92947;-92919;92915;92500;-92700;-92720;92740;92750;92945;-92931;92948;'
              {$ENDIF}
  end
  ;
*)
end;


function ChargeFavoris : boolean ;
begin
  if not VH_GC.GCIfDefCEGID then
     AddGroupFavoris( FMenuG , InitLesTagsToRemoveFavoris('STANDARD'));
  result := true;
end;

procedure ChangeLibelleMenus;
var i : integer;
begin
  For i := 1 to 3 do
  begin
    FMEnuG.RenameItem (74340+i,RechDom ('GCLIBFAMILLE','LF'+inttostr(i), False));
    FMEnuG.RenameItem (148900+i,RechDom ('BTLIBOUVRAGE','BO'+inttostr(i), False));
    FMEnuG.RenameItem (323150+i,RechDom ('BTLIBARTPARC','PA'+inttostr(i), False));
  end;
end;

procedure ChangeLibelleMenusParc;
var i : integer;
begin
  For i := 1 to 3 do
  begin
    FMEnuG.RenameItem (323150+i,RechDom ('BTLIBARTPARC','PA'+inttostr(i), False));
  end;
end;

// ****************************************************************************
// ******************* Dispatch gestion interne + Gestion d'affaire  **********
// ****************************************************************************
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
  stParamsocAffiche, stParamsocVire : String;
  stLesExclus : String;
  Droits    : string;
  oldvalue  : boolean;
  DateDebut : TDateTime;
  NumSem    : Integer;
  AnneeSem  : integer;
  stMessage : String;
  Souche    : string;
  Sql       : string;
BEGIN
	if VersionInterne then V_PGI.VersionDemo:=false;
    ReinitTOBAffaires ; // Eviter les effets de bords

	  if ((Num >= 284000) And (Num <= 284999)) and
       (VH_GC.BTSeriaContrat = False) And
       (V_PGI.VersionDemo = False)    Then
    	 Begin
       PgiBox(TraduireMemoire('Ce module n''est pas sérialisé. Il ne peut être utilisé.'),traduirememoire('Contrats'));
	     RetourForce:=True;
  	   Exit;
	     End
    else if ((Num >= 304000) And (Num <= 304999)) and
       (VH_GC.BTSeriaIntervention = False) and
       (V_PGI.VersionDemo = False) then
    	 Begin
       PgiBox(TraduireMemoire('Ce module n''est pas sérialisé. Il ne peut être utilisé.'),traduirememoire('Interventions'));
	     RetourForce:=True;
  	   Exit;
	     End
    else if ((Num >= 329000) And (Num <= 329999)) and
       (VH_GC.BTSeriaParcMateriel = False) and
       (V_PGI.VersionDemo = False) then
    	 Begin
       PgiBox(TraduireMemoire('Ce module n''est pas sérialisé. Il ne peut être utilisé.'),traduirememoire('Gestion Parc Materiel'));
	     RetourForce:=True;
  	   Exit;
	     End;


    // Gestion des accès selon la sérialisation
    if (((num > 145200) and (num < 145510))  or
        ((num > 145530) and (num < 145810))  or
        ((num > 145850) and (num < 146000))) and
        (VH_GC.BTSeriaES = False) and (V_PGI.VersionDemo = False) then
       Begin
//uniquement en line
//       PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Devis - Factures');
       PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Etudes && Situations');
       RetourForce:=True;
       Exit;
       End;
    //
    if (((num > 31000) and (num < 31401))  or
        ((num > 31404) and (num < 31701))  or
        ((num > 31708) and (num < 32000))) and
        (VH_GC.BTSeriaLogistique = False)  and (V_PGI.VersionDemo = False) then
       Begin
       PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Achats');
       RetourForce:=True;
       Exit;
       End;
    //
    if (((num > 147000) and (num < 147211))  or
       ((num > 147220)  and (num < 147310))  or
       ((num > 147320)  and (num < 148000))) and
       (num <> 147810)                       and
       (VH_GC.BTSeriaChantiers = False)      and
       (V_PGI.VersionDemo = False) then
       Begin
       PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Gestion de chantiers');
       RetourForce:=True;
       Exit;
       End;
    //
    if (num > 92000) and (num <> 92102) and (num < 93000) and (VH_GC.BTSeriaGRC = False) and (V_PGI.VersionDemo = False) then
      Begin
      PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Relations Clients');
      RetourForce:=True;
      Exit;
      End;
    //                                                                      
    if (num > 331000) and (num < 331999) then
      Begin
{$IFDEF V10}                                             
        if (VH_GC.SeriaPlanCharge = False) and (V_PGI.VersionDemo = False) then
        begin
          PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','plannings');
          RetourForce:=True;
          Exit;
        end;
{$ELSE}
        if (VH_GC.SeriaPlanCharge = False) then
        begin
          PgiBox('Ce module n''est disponible qu''en version 10.Veuillez prendre contact avec votre commercial','plannings');
          RetourForce:=True;
          Exit;
        end;
{$ENDIF}
      End;

   case Num of
   10 : BEGIN   //  après connexion
          VH_GC.GcMarcheVentilAna:='AFF';

          // verification du droit d'utilisation
          if not VerificationDroitsOk then
          begin
              stMessage := 'Vous ne disposez plus des droits d''utilisation de ce logiciel';
              stMessage := stMessage +#13#10 +' Merci de contacter votre service commercial';
              stMessage := stMessage +#13#10 +' Votre N° d''identification : '+ getIdDroits;
              PGIInfo(stMessage,'DROITS D''UTILISATION DU LOGICIEL');
              Fmenug.Quitter;
              Exit;
          end;
          // -----

          if not DelphiRunning then
          begin
            // Controle que le programme est en bonne version pour se connecter a la base de donnée
  {$IFDEF PGIMAJVER}
            if BTPGetVersionRef <> NumRefObligatoire then
  {$ELSE}
            if BTPGetVersion <> NumRefObligatoire then
  {$ENDIF}
            begin
              stMessage := 'Merci de contacter votre administrateur';
              stMessage := stMessage +#13#10 +'afin de mettre la version de produit correcte sur votre poste';
              PGIInfo(stMessage,'INFORMATION VERSION DE PROGRAMME');
              Fmenug.Quitter;
              Exit;
            end;
          end;
          InitParPieceBTP ;
  {$IFDEF GRC}
          if GetParamSocSecur ('SO_RTCTIGESTION', False) then
             FMenuDisp.InitializeCTI;
  {$ENDIF}

          TraiteChangementGamme; // Passage S1->S3 ou S3->S1

          //uniquement en line
          if NomHalley = 'CBTPS1' then
          begin
            // Appel automatique de l'assistant de paramétrage société
            if not GetParamSocSecur ('SO_BTPARAMSOCOK',false) then
            begin
              if AppelParamSociete (True) then
              begin
                // Si l'on a importé des données --> Sortie
                RetourForce:=True;
                Fmenug.Quitter;
                Exit;
              end else
              begin
                RetourForce:=True;
              end;
            end;
          end;

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

          SPIGAOOBJ.MonteToaster;

          { Si module RGPD visible et si l'utilisateur n'a pas encore validé le message de sensibilisation RGPD, affiche le message }
          Sql := 'SELECT 1 FROM MENU'
               + ' WHERE MN_1 = 0'
               + '   AND MN_2 = 280'
               + '   AND SUBSTRING(MN_ACCESGRP, 1, (SELECT COUNT(UG_GROUPE) FROM USERGRP)) = REPLICATE("-", (SELECT COUNT(UG_GROUPE) FROM USERGRP))'
                 ;
          if not ExisteSQL(Sql) then
          begin
            Sql := 'SELECT 1 FROM UTILISAT WHERE US_UTILISATEUR = "' + V_PGI.User + '" AND US_RGPDOK <> "X"';
            if ExisteSql(Sql) then
            begin
              Sql := OpenForm.RGPDSensibilisation;
              if Sql <> 'X' then
                FMenuG.Quitter
              else
                ExecuteSQL('UPDATE UTILISAT SET US_RGPDOK = "X" WHERE US_UTILISATEUR = "' + V_PGI.User + '"');
            end;
          end;
          { Test date du dernier changement de mot de passe }
          if GetParamSocSecur('SO_BTDELAIPASSWORD', 0) > 0 then
          begin
            Sql := 'SELECT 1 FROM UTILISAT WHERE US_UTILISATEUR = "' + V_PGI.User + '" AND US_DATECHANGEPWD <= "' + UsDateTime(IncMonth(Date, -3)) + '"';
            if ExisteSql(Sql) then
              PGIInfo(TraduireMemoire('Il est conseillé de changer régulièrement son mot de passe .'), TraduireMemoire('Information.'));
          end;

        END ;
   11 : BEGIN
        //uniquement en line
				//Gestion de la sauvegarde automatique !!!!!
				//
        if NomHalley = 'CBTPS1' then
        begin
          If V_PGI.Superviseur then
           Begin
           if GetParamSocSecur('SO_SAUVEAUTO', False) then
              Begin
              if GetParamSocSecur('SO_ASKBEFORESAVE', True) Then
                 Begin
                 BackupPGIS1(False);
                 end;
              end;
           End;
				//
        end;
   			END; //  après déconnexion
   12 : begin
        end;
   13 : begin
{$IFDEF EAGLCLIENT}
        ChargeMenuPop(TTypeMenu(-1),FMenuG.DispatchX) ;
{$ELSE}
        ChargeMenuPop(-1,FMenuG.DispatchX) ;
{$ENDIF}
				{ Mng 05/03/2007 - Finalisation interface CTI  }
        FMenuDisp.TerminateCTI;
        end;
	 15 : begin
   			end;
   16 : begin
   (*
{$IFNDEF EAGLCLIENT}
          Coderetour := MajStructure;
          if CodeRetour <> 0 then
          begin
      			RetourForce:=True;
            Fmenug.Quitter;
            Exit;
          end;
   				TraiteChangementGamme;
{$ENDIF}
*)
   				stLesExclus := '';
   				LesTagsToRemove := LesTagsToRemove + ';' + stLesExclus;
          LesTagsToRemove      := InitLesTagsToRemoveFavoris('STANDARD');
   		  end; // entre connexion et chargemodule
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

   // Module Révision de Prix
   284510 : AFLanceFiche_MulPublication ;
	 284530 : AFLanceFiche_MulFormule;
   284521 : AFLanceFiche_MulIndice;
   284522 : AFLanceFiche_MulValIndice;
   284523 : AFLanceFiche_AFREVPREPAIMPORT;
   284524 : AFLanceFiche_AFREVEXPORTINDICE;
   284541 : AFLanceFicheRevCalcCoef;
   284542 : AFLanceFicheModifCoef('MODIFIER');
   284543 : AFLanceFicheModifCoef('APPLIQUER');
   284544 : FlagueLesLignesRegularisables(false);
   284545 : AFLanceFicheRegulFacture;
   284551 : LanceEtatDernieresValeursIndicesSaisies;
   284552 : LanceEtatIndicesNonSaisis;
   284553 : LanceEtatFactureARegulariser;

   // Module Contrats
   284110 :	AglLanceFiche('BTP','BTCONTRAT_MUL','AFF_AFFAIRE0=I','','ETAT=ENC;STATUT=INT');   // Contrats
   284120 : AGLLanceFiche('BTP','BTRESCONTRATS','','','STATUT=INT') ;
   284130 : AGLLanceFiche('RT','RTACTIONS_MUL','','','STATUT=INT');
   284140 : AGLLanceFiche('BTP','BTTACHEMODELE_MUL','','','');
   284160 : AGLLanceFiche('BTP','BTEDITCONDIFF_MUL','GP_NATUREPIECEG=AFF','','VEN;STATUT=INT') ;
   284150 : AGLLanceFiche('BTP','BTTYPEAFFAIRE','','','TYPEAFFAIRE=I') ;
   284201 : AFLanceFiche_Mul_PrepFac;   // Prépa facture
//   284202 : AFLanceFiche_Mul_VisaPiecePro('TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT=INT'); //visas factures provisoires
   284203 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT=INT'); //;NOCHANGE_NATUREPIECE validation documents provisoires
   284204 : AFLanceFiche_Mul_Annu_PiecePro('TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT=INT'); //;NOCHANGE_NATUREPIECE
   284205 : AglLanceFiche('BTP','BTMULAFFECH','AFF_AFFAIRE0=I','','STATUT=INT;ANNEE=' + FormatDateTime('yyyy',Now));   // Affaires
// 284209 : AFLanceFicheRegulFacture;
   284210 : AGLLanceFiche('BTP','BTVENTES_CUBE','','','I') ;
   284301 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FAC','','VEN;STATUT=INT') ;
   284302 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FPR;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=INT');   //Modif Factures
	 284303 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FPR','','VEN') ;
   284304 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=INT');   //Modif Factures
   284305 : AGLLanceFiche('BTP','BTREGLPIECE_MUL','GP_NATUREPIECEG=FAC','','STATUT=INT') ; // Enregistrement des reglements non comptabilisés
   284410 : RTLanceFiche_Mailing('BTP','BTMULCONTREN','','','REC') ;
   284420 : RTLanceFiche_ReportActions('RT','RTACTIONS_REPORT','','','');
   284430 : RTLanceFiche_Mailing('RT','RTPROS_MAILIN_FIC','','','REC') ;
   284440 : AGLLanceFiche('RT','RTACTIONS_MUL_RAP','','','') ;
   284281 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FAC;NOCHANGE_NATUREPIECE;STATUT:INT;VIVANTE: ;');
   284282 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=AVC;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=INT');   //Modif Avoirs
   284283 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=AVC','','VEN;STATUT=INT') ;
   284284 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN','','STATUT=INT;') ; // Viser les pièces
   284285 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN','','STATUT=INT;RESTORE') ; // Enlever Visa
   284286 : AGLLanceFiche('BTP','BTTVACONTRAT_MUL','AFF_AFFAIRE0=I','','ETAT=ENC;STATUT=INT');
   // Module Interventions
   304110 :  if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
       		    AGLLanceFiche('BTP', 'BTMULAPPELSINT','AFF_AFFAIRE0=W','','ETAT=ECO;STATUT=APP;NOCHANGESTATUT')
       		 else
       		    AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W','','ETAT=ECO;STATUT=APP;NOCHANGESTATUT'); // Affaires
	 304120 : AGLLanceFiche('BTP','BTACCDEV_MUL','GP_NATUREPIECEG=DAP; AFF_ETATAFFAIRE=ACD','','MODIFICATION;STATUT=APP') ;   //devis
	 304140 : if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
       		     AGLLanceFiche('BTP', 'BTMULAPPELSINT','AFF_AFFAIRE0=W','','ETAT=AFF;STATUT=APP;REALISATION')
       		 else
       		 	AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W','','ETAT=AFF;STATUT=APP;REALISATION'); // Affaires
   304150 : BEGIN
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
   					END;
   304170 : if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
       		     AGLLanceFiche('BTP', 'BTMULAPPELSINT','AFF_AFFAIRE0=W','','ETAT=REA;STATUT=APP;TERMINE')
       		  else
         		 	 AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W','','ETAT=REA;STATUT=APP;TERMINE'); // Affaires

   304180 : AGLLanceFiche('BTP','BTDEVISINT_MUL','GP_NATUREPIECEG=DAP;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=APP');   //Modif Avoirs
   304190 : AGLLanceFiche('BTP','BTIMPORTMULDEV','','','');   // Intégration des devis en attente 

	 304210 : AGLLanceFiche('BTP','BTMULEDIBI','','','STATUT=APP');
	 304211 : AGLLanceFiche('BTP','BTMULEDIBI','','','ETAT=REA;STATUT=APP');
	 304220 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=DAP','','VEN;STATUT=APP') ;
	 304310 : Begin
    	      if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
       		     AGLLanceFiche('BTP', 'BTMULAPPELSINT','AFF_AFFAIRE0=W','','ETAT=ECO;STATUT=APP;AFFECTATION')
       		  else
       		     AglLanceFiche('BTP','BTMULAPPELS','AFF_AFFAIRE0=W','','ETAT=ECO;STATUT=APP;AFFECTATION'); // Affaires
			      end;

   304320 : Agllancefiche ('BTP', 'BTPARAMPLAN_MUL', 'HPP_MODEPLANNING=PLA', '', 'TYPEPLANNING=PLA') ;

   304330	:
           begin
            OldValue := V_Pgi.ZoomOle ;
            V_Pgi.ZoomOle := True ;
            try
           		Saisieplanning ('PLA', 0, taModif) ;
            finally
              V_Pgi.ZoomOle := OldValue ;
            end;
    {$IFNDEF EAGLCLIENT}
            retourforce := True;
    {$ENDIF}
		       end;
   304340	:
           begin
//
            OldValue := V_Pgi.ZoomOle ;
            V_Pgi.ZoomOle := True ;
            try
           		Saisieplanning ('PLA', 0, taConsult) ;
            finally
              V_Pgi.ZoomOle := OldValue ;
            end;
    {$IFNDEF EAGLCLIENT}
            retourforce := True;
    {$ENDIF}
//
		       end;
   304350 : //Appel Planning Type d'action
           begin
//
            OldValue := V_Pgi.ZoomOle ;
            V_Pgi.ZoomOle := True ;
            try
           		Saisieplanning ('PTA', 0, taModif) ;
            finally
              V_Pgi.ZoomOle := OldValue ;
            end;
    {$IFNDEF EAGLCLIENT}
            retourforce := True;
    {$ENDIF}
//

		       end;

   304360 : //Appel Planning Sous-Famille
           begin
//
            OldValue := V_Pgi.ZoomOle ;
            V_Pgi.ZoomOle := True ;
            try
           		Saisieplanning ('PSF', 0, taModif) ;
            finally
              V_Pgi.ZoomOle := OldValue ;
            end;
    {$IFNDEF EAGLCLIENT}
            retourforce := True;
    {$ENDIF}
//

		       end;
   304370 : BEGIN SaisirHorsInside('BTPARAMMAIL','TYPEMAIL=P'); RetourForce := TRUE ; end;

   304410 :
           Begin
             SaisirHorsInside('SAISIEINTERVENTION');
             RetourForce:=TRUE;
           End;

   //304510 : AGLLanceFiche('BTP','BTPREFAC_MUL','GP_NATUREPIECEG=DAP','','MODIFICATION;STATUT=APP') ;   																// Prépa facture

   304510 : AGLLanceFiche ('BTP','BTPREPFACAPP','AFF_AFFAIRE0=W','','STATUT=APP'); //AFLanceFiche_Mul_PrepFacApp;
   304520 : AFLanceFiche_Mul_VisaPiecePro('TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT=APP'); 																//Visa sur facture
   304530 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT:'+ TypeFacDetail); 														//Validation définitive facture provisoire
   304540 : AGLLanceFiche('AFF', 'AFPIECEPROANU_MUL','GP_VIVANTE=X', '', 'TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT='+TypeFacDetail);
   304551 : AFLanceFiche_Mul_Piece_Provisoire('TABLE:GP;NATURE:FAC;NOCHANGE_NATUREPIECE;STATUT=APP;VIVANTE=-');
   304552 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=AVC;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=APP');			//Saisie Avoirs Direct
   304553 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=AVC','','VEN;STATUT=APP') ; 																				//Edition des Avoirs
   304561:	if GetParamSoc('SO_FACTPROV') = False then
               AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN;AFFAIRE0=W','','MODIFICATION;STATUT=' + TypeFacDetail)
          	else
               AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FPR;GP_VENTEACHAT=VEN;AFFAIRE0=W','','MODIFICATION;STATUT=' + TypeFacDetail);
   304562 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN;AFFAIRE0=W','','MODIFICATION;STATUT=' + TypeFacDetail);   		//Modif Factures
   304563 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FAC','','VEN;STATUT='+ TypeFacDetail) ;
   304564 : AGLLanceFiche('BTP','BTREGLPIECE_MUL','GP_NATUREPIECEG=FAC','','STATUT=APP') ; // Enregistrement des reglements non comptabilisés
   //
   304570 : AGLLanceFiche('BTP','BTVENTES_CUBE','','','W') ;
   304580 : AGLLanceFiche('BTP','BTFACCONSO_MUL', 'BCO_AFFAIRE0=W;BCO_FACTURABLE=A', '', 'TABLE:GP;NATURE:FPR;QUEPROVISOIRES;STATUT=' + TypeFacDetail);  //Visas sur Avoir

   //: AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FPR','','VEN;STATUT=APP') ;

   304610 : AglLanceFiche('BTP', 'BTFAMRES_MUL', '', '', '') ;
   304620 : AglLanceFiche('BTP', 'BTTYPERES_MUL', '', '', '') ;
   304630 : AGLLanceFiche('BTP', 'BTTYPEACTION_MUL','','','TYPEACTION=INT'); //Agllancefiche('BTP', 'BTETAT_MUL', '', '', '') ;
   304640 : Agllancefiche('BTP', 'BTEVENEMENT_MUL', '', '', '') ;
   304670 : AGLLanceFiche('BTP', 'BTHORAIRESTD','','', 'TYPE=STD');
   304650 : AGLLanceFiche('BTP', 'BTTYPEAFFAIRE','','','TYPEAFFAIRE=W') ;

   304701 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN','','STATUT=APP') ; 														// Visa des factures
   304702 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=FAC;GP_VENTEACHAT=VEN','','STATUT=APP;RESTORE') ; 														// Visa des factures
   
   // Tarifs TTC
   30902 : EntreeTarifArticle(taModif,TRUE);
   30903 : EntreeTarifTiers(taModif,TRUE);
   30904 : EntreeTarifCliArt(taModif,TRUE) ;
{$ENDIF}
   // Clients
   30501 : if ctxMode in V_PGI.PGIContexte then
              AGLLanceFiche('GC','GCCLIMUL_MODE','T_NATUREAUXI=CLI','','')
{$IFDEF GRC}
           else if ctxGRC in V_PGI.PGIContexte then
              RTLanceFiche_Prospect_Mul('RT','RTPROSPECT_MUL','T_NATUREAUXI=CLI','','GC')
{$ENDIF}
           else
              AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=CLI','','') ;    // Clients
   30502 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=CLI','','MODIFICATION') ; // contacts
   30504 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','CLI') ;  // modification en série des tiers
   105107 : AGLLanceFiche('MBO','EDTCLI_MODE','','','') ;      // Edition des clients
   33102 : AGLLanceFiche('GC','GCANALYSE_VENTES','GPP_VENTEACHAT=VEN','','CONSULTATION') ;
   //33105 : AGLLanceFiche('BTP','BTVENTES_CUBE','','','CONSULTATION') ;
   33105 : AGLLanceFiche('BTP','BTVENTES_CUBE','','','A') ;

{$IFDEF CTAITMIEUXAVANTTT}
{$IFNDEF SANSCOMPTA}
   30505 : OuvreFermeCpte(fbAux,FALSE,'CLI') ;  // fermeture des comptes tiers
   30506 : OuvreFermeCpte(fbAux,True,'CLI') ;   // ouverture compte tiers
   30507 : SuppressionCpteAuxi('CLI') ;
{$ENDIF}
{$ELSE}
   30507 : if _BlocageMonoPoste(False,'',TRUE) then RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'CLI');
   30506 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;CLI') ;  // ouverture compte tiers
   30505 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;CLI') ;  // fermeture compte tiers
{$ENDIF}
    // FOurnisseurs
{$IFNDEF SANSPARAM}
   31401 : EntreeTarifFouArt (taModif) ;                       // saisie
   31402 : AGLLanceFiche('GC','GCTARIFFOUCON_MUL','','','') ;  // consultation
   31403 : AGLLanceFiche('GC','GCTARIFFOUMAJ_MUL','','','') ;  // mise à jour
   31404 : AGLLanceFiche('GC','GCPARFOU_MUL','','','') ;  // Import
{$ENDIF}
   31701 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','') ;
   31702 : AGLLanceFiche('GC','GCMULCATALOGUE','','','') ;
   31703 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=FOU','','') ;  // contacts
   31704 : AGLLanceFiche('RT','RTQUALITE','','','FOU;GC') ;  // modification en série des tiers
   31705 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;FOU') ;
   31706 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;FOU') ;
   31707 : if _BlocageMonoPoste(False,'',TRUE) then RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'FOU');
   //
   31708 : AGLLanceFiche('MBO','EDTFOU','','','') ;      // Edition des Fournisseur
   //
// Commerciaux
   30503 : AGLLanceFiche ('GC','GCCOMMERCIAL_MUL','','','') ; // Commerciaux
   30604 : AGLLanceFiche('GC','GCETATCOMMISSION','','','') ; // commissionnement commerciaux
   // NEw ONE
   // ACHATS
   // Saisie pièces
//   31201 : CreerPiece('CF',True) ;
//   31202 : CreerPiece('BLF',True) ;
//   31203 : CreerPiece('FF',True) ;
//   31206 : CreerPiece('DEF',True) ;
	 // Cube des achats
   146501 : AGLLanceFiche('BTP','BTACHAT_CUBE','','','') ;
   // Modifications
   31211  : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH','','CF;MODIFICATION;CHANTIER')  ;  //commandes fournisseur Chantier
   31215  : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH','','CF;MODIFICATION;STOCK')  ;  //commandes fournisseur Stock
   146285 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=CFR;GP_VENTEACHAT=ACH','','MODIFICATION')  ;  //commandes fournisseur regr.
   146286 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=LFR;GP_VENTEACHAT=ACH','','MODIFICATION')  ;  //réceptions fournisseur regr.
   31212  : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=BLF;GP_VENTEACHAT=ACH','','MODIFICATION') ;  //bons de livraison fournisseur
   31213  : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=DEF;GP_VENTEACHAT=ACH','','MODIFICATION') ;  //Proposition d'achat fournisseur
   31214  : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FF;GP_VENTEACHAT=ACH','','MODIFICATION') ;  //Proposition d'achat fournisseur
   // Duplication unitaire pièces
   31221 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH','','CF') ;
   31222 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=BLF;GP_VENTEACHAT=ACH','','BLF') ;
   31223 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=FF;GP_VENTEACHAT=ACH','','FF') ;
   31224 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=DEF;GP_VENTEACHAT=ACH','','DEF') ;
   31225 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=AF;GP_VENTEACHAT=ACH','','AF') ;
   31226 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=AFS;GP_VENTEACHAT=ACH','','AFS') ;
   31228 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=BFA;GP_VENTEACHAT=ACH','','BFA') ;
   // AJOUT LS
   31234 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=BFA;GP_VENTEACHAT=ACH','','MODIFICATION')  ;  //retour fournisseur
   31231 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=AF;GP_VENTEACHAT=ACH','','MODIFICATION')  ;  //Avoir financier
   31232 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=AFS;GP_VENTEACHAT=ACH','','MODIFICATION')  ;  //Avoir sur stock
   //
   31204 : GCLanceFiche_HistorisePiece('GC','GCHISTORISEPIECE','','','ACH') ; //Histirisation des pièces
{$IFNDEF EAGLCLIENT}
   31205 : GCLanceFiche_PieceEpure('GC','GCPIECEEPURE_MUL','','','ACH') ; //Epuration des pièces
{$ENDIF}
   // Consultation
   31101 : AGLLanceFiche('BTP','BTPIECEACH_MUL','GP_NATUREPIECEG=FF','','CONSULTATION') ;
   31102 : AGLLanceFiche('BTP','BTLIGNEACH_MUL','GL_NATUREPIECEG=FF','','CONSULTATION') ;
   // génération
   31501 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=CF;GP_VENTEACHAT=ACH','','VENTEACHAT=ACH') ; // Visa des pièces
//   146340 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=BLF;GP_VIVANTE=X','','BFA') ;  // Retour chantier
	 146340 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_VENTEACHAT=ACH','','BFA') ;
   31511 : AGLLanceFiche('BTP','BTTRANSACH_MUL','GP_VENTEACHAT=ACH','','BLF') ;
   31512 : AGLLanceFiche('BTP','BTTRANSACH_MUL','GP_VENTEACHAT=ACH','','FF') ;

{   146510 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AF;GP_VENTEACHAT=ACH','','AF') ;}
{   146520 : AGLLanceFiche('GC','GCDUPLICPIECE_MUL','GP_NATUREPIECEG=AFS;GP_VENTEACHAT=ACH','','AFS') ;}
   146510 : AGLLanceFiche('BTP','BTTRANSACH_MUL','GP_VENTEACHAT=ACH','','AF') ;
   146520 : AGLLanceFiche('BTP','BTTRANSACH_MUL','GP_VENTEACHAT=ACH','','AFS') ;
    // Visa des pièces
   146610 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_VENTEACHAT=ACH','','VENTEACHAT=ACH') ; // Visa des pièces; // Mise en place du visa
   146620 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_VENTEACHAT=ACH','','VENTEACHAT=ACH;RESTORE') ; // enlever le visa
   146630 : AGLLanceFiche('BTP','BTIMPORTBSV','','','') ; // Importation BSV (Fichier Xml)
   146640 : AGLLanceFiche('BTP','BTIMPORTMULBSV','','','') ; // Importation BSV par lot(Fichier Xml)
   146641 : ConstitueDocsFromDatasBSV(true);
   146642 : AGLLanceFiche('BTP','BTGENFACBAST','','','') ; // génération des factures achat depuis BAST définitifs
   146643 : BEGIN
              AppelsBAST;
              RetourForce:=TRUE;
            END;

   146311 : AGLLanceFiche('BTP','BTIMPORTMULGIE','','','') ; // Importation des factures d'achat GIE
//   31515 : AGLLanceFiche('BTP','BTTRANSACH_MUL','','','BFA') ;
   31514 : AGLLanceFiche('BTP','BTTRANSACH_MUL','GP_NATUREPIECEG=DEF;GP_VENTEACHAT=ACH','','CF') ;
   // Génération manuelle pièces
   31531 : AGLLanceFiche('BTP','BTGROUPEMANPIECE','GP_VENTEACHAT=ACH','','LIVRAISON;ACHAT;CFR;BLF') ;  // Comm --> Livr
   31532 : AGLLanceFiche('BTP','BTGROUPEMANPIECE','GP_VENTEACHAT=ACH','','FACTURATION;ACHAT;LFR;FF') ; // Livr --> Fact
   // génération automatique
   31521 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CF','','ACHAT;BLF') ;
   31522 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLF','','ACHAT;FF') ;
   146110 : BEGIN DecisionStockReappro;  RetourForce:=TRUE;  END ; // Décisionnel stock sur réappro
   //
   146121 : BEGIN LanceAssistantDecisionnelAch;  RetourForce:=TRUE;  END ; // Décisionnel stock sur réappro
   146122 : AGLLanceFiche('BTP','BTDECISIONACH_MUL','','','MODIFICATION') ;
   146123 : AGLLanceFiche('BTP','BTDECISIONACH_MUL','','','VALIDATION') ;

   31301 : BEGIN LanceAssistantSuggestion ;  RetourForce:=TRUE;  END ;  // Suggestion de réappro
   31302 : AGLLanceFiche('BTP','BTLANCEREA_MUL','','','SELECTION') ;     // Lancement réappro
   146131 : AGLLanceFiche('BTP','BTLANCEREA_MUL','','','MODIFICATION') ;     // Modification réappro
   //
   146241  : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FF;GP_VENTEACHAT=ACH','','MODIFICATION;REAJUSTEANA') ;  //Réajustement analytique
   // Editions
   31601 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','','','ACH') ;
   31602 : AGLLanceFiche('BTP','BTPTFPIECEACH','GP_NATUREPIECEG=CF','','ACHAT') ;  // Portefeuille pièces
   31612 : AGLLanceFiche('GC','GCART_NONREF','','','')  ;
   31611 : AGLLanceFiche('GC','GCFOURN_ARTICLE','','','') ;
   31613 : AGLLanceFiche('GC','GCARTICLE_FOURN','','','') ;
   // ---
   // Article
   32503 : AGLLanceFiche('GC','GCARTICLE_MODIFLO','','','') ;  // modification en série
   32505 : EntreeTradArticle (taModif) ;  // traduction libellé articles
   //32506 : AGLLanceFiche('GC','GCTARIFART_MUL','','','') ;  // Maj BAses Tarifaires
   65106 : ParamTable('TTSECTEUR',taCreat,0,PRien) ;        // secteurs d'activité
//   32206 : AGLLanceFiche('GC','GCINITSTOCK','','','');  // Initialisation des stocks
   // type et côte emplacement se trouvant dans le module établissement
   65305 : ParamTable('GCTYPEEMPLACEMENT',taCreat,0,PRien) ;    // type emplacement
   65306 : ParamTable('GCCOTEEMPLACEMENT',taCreat,0,PRien) ;    // côte emplacement
   65409 : ParamTable('GCMOTIFMOUVEMENT',taCreat,0,PRien) ;          // Motif mouvement stock
// *********************Module Gestion Interne : Fiches de bases **************
//   32504 : AGLLanceFiche('GC','GCMULSUPPRART','','','') ;  // épuration articles
{STOCKS}
    // Consultation
   32110 : AGLLanceFiche('BTP','BTDISPO_MUL','','','') ;
   32111 : AGLLanceFiche('BTP','BTMOUVSTK','','','') ; // stock prévisionnel
   32116 : AGLLanceFiche('BTP','BTRECALCULPMAP','','','') ; // Recalcul PMAP

   32112 : AGLLanceFiche('GC','GCMULCONSULTSTOCK','','','') ; // stock prévisionnel
   32113 : AGLLanceFiche('GC','GCCONTREAVT','','','') ; // suivi contremarque
   32114 : GCLanceFiche_MouvStkEx('GC','GCMOUVINTERNE_MUL','','','') ; // Pièces internes
   32115 : GCLanceFiche_ArtNonMvtes('GC','GCARTNONMVTES_MUL','','','') ; // Consult des articles non mouvementés
   // Traitement
   32201 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;TEM'); // Transferts Inter-Dépôt
   32221 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;EEX');
   32222 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;SEX');
   32231 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;RCC'); // Réajustement réservé client
   32232 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;RCF'); // Réajustement réservé fournisseur
   32233 : GCLanceFiche_MouvStkEx('GC','GCMOUVSTKEX','','','ACTION=CREATION;RPR'); // Réajustement préparé client
   32203 : AGLLanceFiche('GC','GCTTFINMOIS','','','');
   32204 : AGLLanceFiche('GC','GCTTFINEXERCICE','','','');
   32205 : AGLLanceFiche('GC','GCAFFECTSTOCK','','','');
   32206 : AGLLanceFiche('BTP','BTINITSTOCK','','','');  // Initialisation des stocks
   32207 : AGLLanceFiche('BTP','BTREVALPMAP','','','TRAITEDEPOT');  // Initialisation des stocks
//   32207 : GCLanceFiche_MouvStkExContr('GC','GCMOUVSTKEXCONTR','','','');  // Mouvements exceptionnels contremarque
   // inventaire ////
   32310 : AGLLanceFiche('BTP','BTLISTEINV','','','') ;
   32315 : AGLLanceFiche('BTP','BTLISTEINVPRE','','','') ;
   32317 : GCLanceFiche_ListeInvVal('BTP','BTLISTEINVVAL','','','') ;
   32320 : AGLLanceFiche('BTP','BTSAISIEINV_MUL','','','') ;
   32330 : AGLLanceFiche('GC','GCVALIDINV_MUL','','','') ;
   32340 : GCLanceFiche_ListeInvMul('GC','GCLISTEINV_MUL','','','') ;

   32351 : AGLLanceFiche('MBO','TRANSTPI','','','') ;       // Transmission inventaire depuis un TPI (Mode 103111)
   32352 : AGLLanceFiche('MBO','TRANSINV_MUL','','','') ;   // Saisie / Consultation des inventaires transmis   (Mode 103112)
   32353 : AGLLanceFiche('GC','GCINTEGRINV_MUL','','','') ;  // Intégration des inventaires transmis  (Mode 103113)
   // Editions ////
   32401 : AGLLanceFiche('GC','GCARTDISPO','','','') ;
   32403 : AGLLanceFiche('GC','GCINVPERM','','','')  ;   // inventaire permanent
   32402 : AGLLanceFiche('GC','GCETIQART','','','')  ;
   32404 : AGLLanceFiche('BTP','BTETIQUETTE','','','')  ;
   32406 : AGLLanceFiche('BTP','BTINTEGREFICTPI','','','');
   32407 : AGLLanceFiche('BTP','BTARTDISPO','','','');
   32408 : AGLLanceFiche('BTP','BTETATSTKPAYE','','','');
   {$IFNDEF EAGLCLIENT}
   32405 : if not AppelConsoStock then Retourforce:=true;
   {$ENDIF}
   // Fichiers
   32601 : AGLLanceFiche('GC','GCDEPOT_MUL','','','') ;
   32602 : AGLLanceFiche('GC','GCEMPLACEMENT_MUL','','','') ;
   32603 : AGLLanceFiche('GC','GCPRIXREVIENT','','','') ;
{------}


//------ Boîte à Outils Business Line BTP
//uniquement en line
//Prevoir la mise à jour du Paramsoc en fonction si assistant ou menu (zone de paramètre)
   320111 : Begin //
				   	TheTOB := TOBparamSoc;
            SaisirHorsInside('BTPARAMSOCS');
            TheTOB := nil;
            RetourForce:=True;
            end;
   320112 : Begin
            TheTOB := TOBparamSoc;
            SaisirHorsInside('BTPARAMGENS1');
            TheTOB := nil;
            RetourForce:=True;
            end;
   320113 : Begin
            TheTOB := TOBparamSoc;
            SaisirHorsInside('BTPARSOCCPTAGS1');
            TheTOB := nil;
            RetourForce:=True;
            end;
   320114	: Begin
            SaisirHorsInside('BTCPTPIECE_S1') ;
            RetourForce:=True;
            end;
	 320115	: Begin
					  SaisirHorsInside('BTPARAMSOCPART');
            RetourForce:=True;
   					End;
   320120 : Begin
						SaisirHorsInside ('PARAMSOC');
            RetourForce:=True;
   					end;
   320131 : Begin
            if PrepareOuvreExoSansCompta then
               OuvertureExo
            Else
	  			     RetourForce:=True;
   					End;
   320132 : Begin
            SaisirHorsInside ('BTCLOTUREEX');
            RetourForce:=TRUE;
   					End;
   320141 : Begin
            RegimesTva ;
            End;
   320143	: Begin
					  //ParamTvaTpf(true) ;
            SaisirHorsInside ('BTTVA');
            RetourForce:=TRUE;
   					End;
   320151 : Begin
       			FicheDevise('',tamodif,False);
            end;
   320152 : Begin
            FicheModePaie_AGL('');
            end;
   320153 : Begin
            FicheRegle_AGL ('',False,taModif);
            end;
   320154 : Begin
            OuvrePostal(PRien) ;
            end;
   320155 : Begin
    				FicheRegion('','',False) ;
            end;
   320156 : Begin
    			  OuvrePays ;
            end;
   320157 : Begin
						ParamTable('ttFormeJuridique',taCreat,1120000,PRien) ;
            end;
   320158 : Begin
						ParamTable('ttCivilite',taCreat,1122000,PRien) ;
            end;
   320159 : Begin
            AgLlanceFIche('BTP','BDOMACT_MUL','','','');
            end;
   320161 : Begin
   					AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=DBT;SOUCHE=DBT;NUMERO=0') ;
   					end;
   320162 : Begin
   					AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=FBT;SOUCHE=FBT;NUMERO=0') ;
   					end;
   320163 : Begin
   					AGLLanceFiche ('BTP','BTPARIMPDOC','','','NATURE=ABT;SOUCHE=ABT;NUMERO=0') ;
   					end;
   320164 : Begin
						EditEtat('E','GPJ','',True, nil, '', '');
   					RetourForce:=TRUE;
   					end;
	 320211 : Begin
            AppelAssistExport;
   					RetourForce:=TRUE;
   					end;
   320212 : Begin
            AppelAssistImport;
   					RetourForce:=TRUE;
   					end;
   320221 : Begin
            BackupPGIS1;
            RetourForce:=TRUE;
            end;
   320222 : Begin
            //RestorePGIS1;
            PGIInfoAF('Si vous désirez restaurer une sauvegarde, veuillez utiliser l''administrateur SQL ou appeller le service de maintenance', 'Restauration');
            RetourForce:=TRUE;
            end;
   320230 : Begin
            if DetectionConnexionInternet then
               OuvertureWebBrowser;
   					end;
   320241 : Begin
            FicheUSer(V_PGI.User) ;
				  	ControleUsers;
            End;
   320242 : Begin
				    ReseauUtilisateurs(False)
            End;
   320321 : Begin
  					ParamTable('GCCOMPTAARTICLE',taCreat,0,nil) ;
					  AvertirTable ('GCCOMPTAARTICLE');
  					end;
   320322 : Begin
            paramTable('GCCOMPTATIERS',taCreat,0,nil) ;
					  AvertirTable ('GCCOMPTATIERS');
  					end;
   320323 : Begin
            ParamTable('AFCOMPTAAFFAIRE',taCreat,0,nil) ;
            AvertirTable ('AFCOMPTAAFFAIRE');
  					end;
   320331 : Begin
            AglLanceFiche ('GC','GCCODECPTA','','','') ;
            End;
   320332 : Begin
				   	TheTOB := TOBparamSoc;
					  AglLanceFiche ('BTP','BTVENTILCPTAS1','','','');
					  TheTOB := nil;
   					End;

{$IFNDEF SANSCOMPTA} // aujourd'hui InitSociété= compta
// Administration
{$IFNDEF V530}
   60101 : GestionSociete(PRien,@InitSociete,Nil) ; // Gestionnaire société
{$ELSE}
   60101 : GestionSociete(PRien,@InitSociete) ; // Gestionnaire société
{$ENDIF}
{$ENDIF}

   60102 : RecopieSoc(PRien,True) ;         // recopie société à société
  {$IFNDEF EAGLCLIENT}
   60103 : YYLanceFiche_Bundle ('YY', 'YYBUNDLE', '', '', '');
  {$ENDIF EAGLCLIENT}
   //////// Utilisateurs et accès
{$IFNDEF SANSPARAM}
   60201 : FicheUserGrp ;                   // groupe utilisateurs
   60403 : Entree_TraiteAdresse ;
{$ENDIF}
   60202 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;   // Utilisateurs
   // --
   60206 : AGLLanceFiche('GC','GCPARAMOBLIG','','','') ;      // Champs obligatoires
   60207 : AGLLanceFiche('GC','GCPARAMCONFID','','','') ;       // restricyions fiches
   // FAVORIS
   60501 : // Gestion des favoris
     if not VH_GC.GCIfDefCEGID then
        ParamFavoris('',LesTagsToRemove,False,False)
     else
{        ParamFavoris('',LesTagsToRemove
Cegid,False,False)};

   // MODIF LS Confidentialité
   60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','26;27;60;92;111;145;146;147;148;149;150;160;279;283;284;285;304;319;320;321;322;323;325;327;328;329;331;280');
   // --
   60203 : ReseauUtilisateurs(False) ;      // utilisateurs connectés
   60204 : VisuLog ;                        // Suivi d'activité
   60205 : ReseauUtilisateurs(True) ;       // RAZ connexions
    /////// Outils
//   60301 : AGLLanceFiche('GC','GCJNALEVENT_MUL','','','') ;
   60301: OpenForm.JnalEvent; // Journal evenements

   60401 : AGLLanceFiche('GC','GCCPTADIFF','','','') ;
   60402 : EntreeStockAjust;

   ///60406 : MajPhonetiqueTiers ;
   60407 : RTCreerProspectPourClient;
   60408 : RTAlimCleTelephonTiers ;
   60409 : RTAlimTelephonContact ;
 {$IFNDEF EAGLCLIENT}
   60211 : AGLLanceFiche('BTP','BTPROFILTAB','','','ETA;GRP') ; // JTR - Fonction en CWAS
   3172  : AGLLanceFiche('BTP','BTPROFILTAB','','','ETA;UTI') ; // JTR - Fonction en CWAS
 {$ELSE}
   60211 : AGLLanceFiche('BTP','BTPROFILTAB','ETA;GRP','','ETA;GRP') ; // JTR - Fonction en CWAS
   3172  : AGLLanceFiche('BTP','BTPROFILTAB','ETA;UTI','','ETA;UTI') ; // JTR - Fonction en CWAS
 {$ENDIF EAGLCLIENT}

 	 //Gestion menu propre Line --> $LINE
   3145: RecopieSoc(PRien, True);
   3220: ReindexSoc;
    { Alertes }
   60209 : RTLanceFiche_Alertes_Mul('Y','YALERTES_MUL','','','');
   60212 : begin
           AGLLanceFiche('Y','YTABLEALERTES','','','');
           AvertirTable ('YTABLEALERTES');
           VH_EntPgi.TobTableAlertes.ClearDetail;
           VH_EntPgi.TobTablesLiees.ClearDetail;
           end;
   60213 : AGLLanceFiche('Y','YALERTES_LOG','','','');
   { Imprimantes par défaut }
   4006  : BEGIN
              if V_Pgi.Superviseur then Droits := 'DROIT=ADMIN' else Droits := '';
              AglLanceFiche('YY','YYDEFPRT','','',Droits,'');
           END;
   4010: AGLLanceFiche('BTP', 'BTIMPORT', '', '', ''); // Import d'objets
   4020: AGLLanceFiche('YY', 'YYMYINDEXES', '', '', ''); // Indexes supplémentaires

   // utilitaires
   60602 : Assist_RazActivite;
   60603 : GCLanceFiche_UtilULst('GC','GCUTILULST','','','1;0');
   //
   60604 : AglLanceFiche ('BTP','BTRECALCPIECE_MUL','','','ACTION=MODIFICATION');
   //
   60605 : EntreeMajCompteurSouche;
   60606 : GCLanceFiche_UtilULst('GC','GCUTILULST','','','2;1');
   60607 : AglLanceFiche('BTP','BTRECALCPMAPNOSTK','','','ACTION=MODIFICATION');
   60608 : Assist_AffecteDepot;
   60609 : GCLanceFiche_UtilULst('GC','GCUTILULST','','','3;1');
   60610 : EntreeVerifAuxiTiers;
   60611 : EntreeVerifNatErrImp;
   60612 : AglLanceFiche ('BTP','BTMANPIECE_MUL','','','ACTION=MODIFICATION'); // Mise a niveau des documents
   // EPURATIONS
   60651 : BEGIN
            if V_PGI.Superviseur then
            begin
            	AGLLanceFiche('BTP', 'BTFERMEPIECE_MUL', '', '', ''); // Fermeture/reouverture de pièces
            end else
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
            end;
   				 END;
   60652 : BEGIN
            if V_PGI.Superviseur then
            begin
            	AGLLanceFiche('BTP', 'BTEPUREPIECE_MUL', '', '', ''); // 2puration de pièces
            end else
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
            end;
   				 END;
   60653 : BEGIN
            if V_PGI.Superviseur then
            begin
            	AGLLanceFiche('BTP', 'BTEPUREAFF_MUL', 'AFF_AFFAIRE0=A;AFF_ETATAFFAIRE=TER', '', ''); // 2puration de pièces
            end else
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
            end;
   				 END;
   60654 : BEGIN
   					if V_PGI.Superviseur then
            begin
              AGLLanceFiche('BTP','BTCLOTUREAFF_MUL','AFF_AFFAIRE0=A;AFF_ETATAFFAIRE=ACP','','');  // Clôture de chantier
            end else
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
            end;
   				 END;																																																 // Antériorite
   60655 : BEGIN
   					if V_PGI.Superviseur then
            begin
              AGLLanceFiche('BTP','BTCLOTUREAFF_MUL','AFF_AFFAIRE0=A;AFF_ETATAFFAIRE=TER','','REACTIVE');  // Réactivation de chantier
            end else
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
            end;
   				 END;																																																 // Antériorite
   60681 : BEGIN
   					if not V_PGI.Superviseur then
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
              exit;
            end;
            AGLLanceFiche('BTP', 'BTEPUREPIECE_MUL', '', '', 'SUPFACTURE;SANSCONTROLE'); // Epuration de factures sans controle
   				 END;
   60682 : BEGIN
   					if not V_PGI.Superviseur then
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
              exit;
            end;
            AGLLanceFiche('BTP','BTSELFAC_MUL','AFF_ETATAFFAIRE=ACP','','MODIFICATION;STATUT=AFF;REAJUSTEAVANC') ;   
   				 END;
   60683 : BEGIN
   					if not V_PGI.Superviseur then
            begin
              PgiInfo ('Vous n''avez pas les droits nécessaires pour effectuer cette opération');
              exit;
            end;
            AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF;REAJUSTESIT');
           END;
   // ----
   60700 : BEGIN
//   						AGLLanceFiche('BTP','BTSYNCCONTACT','','','');  // Synchronisation des contacts / Outlook
   				 END;
   1210 : BEGIN
   						ParamTraduc(False,nil);
   				 END;

//////// Fiches ///////////
   71101 : AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=CLI','','');   // Clients
   71102 : AglLanceFiche ('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=PRO','','PRO');  // proposition de mission
   71103 : AglLanceFiche ('AFF','AFFAIRE_MUL','AFF_STATUTAFFAIRE=AFF','','AFF');  // mission
   71104 : AglLanceFiche ('BTP','BTRESSOURCE_MUL','','','');   // Assistants
//////// Editions ////////
   71211 : AGLLanceFiche ('AFF','ETATSCLIENT','','','');    // Etats sur clients
   71212 : AGLLanceFiche ('AFF','FICHECLIENT','','','');
   71213 : AGLLanceFiche ('AFF','STATCLIENT','','','');
   71214 : AGLLanceFiche ('AFF','ADRESSECLIENT','','','');

   71221 : AGLLanceFiche ('AFF','ETATSAFFAIRE','','','');   // Etats sur affaires
   71222 : AGLLanceFiche ('AFF','FICHEAFFAIRE','','','');
   71223 : AGLLanceFiche ('AFF','STATAFFAIRE','','','');

   71231 : AGLLanceFiche ('AFF','ETATSRESSOURCE','','',''); // Etats sur ressources
   71232 : AGLLanceFiche ('AFF','FICHERESSOURCE','','','');
   71233 : AGLLanceFiche ('AFF','STATRESSOURCE','','','');
   71234 : AGLLanceFiche ('AFF','ADRESSERESSOURCE','','','');

   71271 : AGLLanceFiche ('AFF','ETATSPRESTATION','','',''); // Etats sur prestations
   71272 : AGLLanceFiche ('AFF','STATPRESTATION','','','');

   71204 : AGLLanceFiche('AFF','ETIQUETTECLIENT ','','','') ;
////// Documents clients ///

////  Traitements /////////
   71401 : AglLanceFiche ('AFF','AFVALIDEPROPOS','AFF_STATUTAFFAIRE=PRO','','PRO');  // acceptation de missions
   71403 : AGLLanceFiche('RT','RTQUALITE','','','CLI;GC') ;  // modification en série des tiers
   71405 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;CLI') ;  // ouverture compte tiers
   71404 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;CLI') ;  // fermeture compte tiers
   71406 : if _BlocageMonoPoste(False,'',TRUE) then RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'CLI');
(*
   71403 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','') ;  // modification en série des tiers
   71404 : OuvreFermeCpte(fbAux,FALSE) ;  // fermeture des comptes tiers
   71405 : OuvreFermeCpte(fbAux,True) ;   // ouverture compte tiers
   71406 : SuppressionCpteAuxi ;
   *)
   71407 : AGLLanceFiche('GC','GCEPURATIONART','','','') ;      // Epuration des articles
   71409 : AFLanceFiche_Mul_AugmentAff('AFF_STATUTAFFAIRE=INT');
//////// Structures ///////
   71511 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=PRE','','PRE');    // Prestations
   71512 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=FRA','','FRA');    // Frais
   71513 : AGLLanceFiche('AFF','AFARTICLE_MUL','GA_TYPEARTICLE=MAR','','MAR');    // Fournitures
   71501 : AglLanceFiche ('AFF','AFAPPORTEUR_MUL','','','') ;                  // Apporteurs Attention num utilisé AfterChangeModule
   71502 : AGLLanceFiche ('AFF','HORAIRESTD','','','TYPE:STD');
   71503 : AGLLanceFiche('YY','YYCONTACTTIERS','','','');                      // Annuaire des contacts
{
   ////// Tarifs  pour la gestion d'affaire uniquement
   71601 : EntreeTarifArticle(taModif) ;
   71602 : EntreeTarifTiers (taModif) ;
   71603 : EntreeTarifCliArt (taModif) ;
   71604 : EntreeTarifCatArt (taModif) ;
   71605 : AGLLanceFiche('GC','GCTARIFCON_MUL','','','');
   71606 : AGLLanceFiche('GC','GCTARIFMAJ_MUL','','','');
   // Tarifs TTC
   71611 : EntreeTarifArticle(taModif,TRUE) ;
   71612 : EntreeTarifTiers (taModif,TRUE) ;
   71613 : EntreeTarifCliArt (taModif,TRUE) ;
}

// **************************Module Activité et reporting ******************
//////// Saisie d'activité ///////
   // Par Assistant
{$IFDEF PL}
   72102 : AFCreerActivite ( tsaRess, tacTemps, 'REA' ) ; // Saisie d'activités Temps
   72103 : AFCreerActivite ( tsaRess, tacFrais, 'REA' ) ; // Saisie d'activités Frais
   72104 : AFCreerActivite ( tsaRess, tacFourn, 'REA' ) ; // Saisie d'activités Fournitures
   72105 : AFCreerActivite ( tsaRess, tacGlobal, 'REA' ) ; // Saisie d'activités Globale
   // Par Client/Mission
   72111 :  begin
            if SaisieActiviteManager then
                  AFCreerActivite ( tsaClient, tacTemps, 'REA' ) // Saisie d'activités Temps
              else
                  HShowMessage('2;?caption?;'+TraduireMemoire('Les droits d''accès ne sont pas suffisants : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
            end;
   72112 : begin
            if SaisieActiviteManager then
                  AFCreerActivite ( tsaClient, tacFrais, 'REA' ) // Saisie d'activités Frais
              else
                  HShowMessage('2;?caption?;'+TraduireMemoire('Les droits d''accès ne sont pas suffisants : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
            end;
   72113 : begin
            if SaisieActiviteManager then
                  AFCreerActivite ( tsaClient, tacFourn, 'REA' ) // Saisie d'activités Fournitures
              else
                  HShowMessage('2;?caption?;'+TraduireMemoire('Les droits d''accès ne sont pas suffisants : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
            end;
   72114 : begin
            if SaisieActiviteManager then
                  AFCreerActivite ( tsaClient, tacGlobal, 'REA' ) // Saisie d'activités Globale
              else
                  HShowMessage('2;?caption?;'+TraduireMemoire('Les droits d''accès ne sont pas suffisants : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
            end;
{$ENDIF}
/////// Editions sur activités ////
    72211 : AGLLanceFiche('AFF','AFJOURNALACT','','','') ;
    72212 : AGLLanceFiche('AFF','AFJOURNALCLIENT','','','') ;
    72221 : AGLLanceFiche('AFF','AFSYNTHRESS','','','') ; // Synthèse assistant detaillée par affaire
    72222 : AGLLanceFiche('AFF','AFSYNTHRESSART','','','') ; // Synthèse assistant detaillée par article
    72231 : AGLLanceFiche('AFF','AFSYNTHCLIENT','','','') ; // Synthèse Client affaire détaillée par ressource
    72232 : AGLLanceFiche('AFF','AFSYNTHCLIART','','','') ; // Synthèse Client affaire détaillée par article

////// Historique ///////////


////// Tableaux de bord //////
    72404 : AGLLanceFiche('AFF','AFTABLEAUBORD','','','') ; //AlimTableauBordAffaire;

////// Traitements //////////

// **************************  Module Honoraires  *********************
    ////// Factures et avoirs
    73101 : CreerPiece('FAC') ;  // Saisie de factures
    73102 : CreerPiece('AVC') ;  // Saisie d'Avoirs
    73103 : CreerPiece('FPR') ;  // Saisie de factures provisoires
    73104 : AGLLanceFiche('AFF','AFPIECE_MUL','GP_NATUREPIECEG=FPR','','NOCHANGE_NATUREPIECE');   //Modif Factures pro
    ////// Duplication
    73111 : AGLLanceFiche('AFF','AFDUPLICPIECE_MUL','GP_NATUREPIECEG=FPR','','FPR');  // duplication Factures  provisoires
    73112 : AGLLanceFiche('AFF','AFDUPLICPIECE_MUL','GP_NATUREPIECEG=FAC','','FAC');  // duplication Factures

    ////// Consultation
    73201 : AGLLanceFiche('AFF','AFPIECE_MUL','','','');   // COnsultation par piece
    73202 : AGLLanceFiche('AFF','AFLIGNE_MUL','','','');   // consultation par ligne

    ////// Editions
    73301 : AGLLanceFiche('AFF','AFLISTEFACT','','','');
    73303 : AGLLanceFiche('GC','GCEDTPIECE','','','');
    73304 : AGLLanceFiche('AFF','ETATPREVCA','','','');
    73305 : AGLLanceFiche('AFF','AFPORTPIECE','','','');  // Portefeuille pièces
    73306 : AGLLanceFiche('AFF','ETATCOMMISSION','','',''); // commissionnement commerciaux (rétrocession apporteur)
                                                             // Attention num utilisé AfterChangeModule
    ////// Génération
    73402 : AglLanceFiche ('AFF','FACTAFF_MUL','','','');          //Liste des échéances par date
    73403 : AglLanceFiche ('BTP','BTPIECEPRO_MUL','','','NATURE:FPR;NOCHANGE_NATUREPIECE;STATUT=APP');          //validation factures provisoires
    73404 : AglLanceFiche ('AFF','AFPIECEPROANU_MUL','','','NATURE:FPR;NOCHANGE_NATUREPIECE;STATUT=APP');
    ////// Statistiques

    ////// Règlements
    73601 : AGLLanceFiche('GC','GCREGLEMENT_MUL','','','');  // saisie des reglements

// *****************************************************************************
// ******************************  Module Paramétrage  *************************
// *****************************************************************************
    /////// Société
    74101 : BEGIN
            BrancheParamSocAffiche (stParamsocVire,stParamsocAffiche);
            ParamSociete (FALSE,stParamsocVire,stParamsocAffiche,'',ChargeSocieteHalley,ChargePageSoc,SauvePageSoc,InterfaceSoc,0);
            END;
    74102 : AGLLanceFiche('GC','GCETABLISS','','','') ;          // FicheEtablissement_AGL(taModif);
//    74103 : FicheSouche('GES');                                  // Compteurs de pièces
    74103 : YYLanceFiche_Souche('0;GES','','ACTION=MODIFICATION;GES'); //Compteurs
    74104 : AglLanceFiche ('GC','GCPARPIECE','','','') ;
    74105 : begin EditEtat('E','GPJ','',True, nil, '', '');RetourForce:=True ;end;

    74106 : AglLanceFiche ('GC','GCPARPIECECOMPL','','','NATURE=FAC') ; // Complément Etablissement facture
    74107 : AglLanceFiche ('GC','GCPARPIECECOMPL','','','NATURE=AVC') ; // Complément Etablissement Avoir
	  74108 : BEGIN LanceAssistCodeAppel; RetourForce := True;  END;
    74109 : BEGIN LanceAssistCodeAffaire; RetourForce := True;  END;
    74111 : EntreeListeSaisie('') ;

    ////// Paramètres généraux
    74151 : RegimesTva ;
    74152 : BEGIN
    				//ParamTvaTpf(true) ;
            SaisirHorsInside ('TVA');
            RetourForce:=TRUE;
            END;
    148801 : ParametrageTypeTaxe;
    148821 : AglLanceFiche ('BTP','BSVSERVER','','','ACTION=MODIFICATION') ;
    148822 : BEGIN
              V_PGI.ZoomOle := true;
              AglLanceFiche ('BTP','BTPARAMLIGBAST','','','ACTION=MODIFICATION') ;
              V_PGI.ZoomOle := false;
             END;
    148823 : BEGIN
              V_PGI.ZoomOle := true;
              AglLanceFiche ('BTP','BDEPENSESPOC','','','ACTION=MODIFICATION') ;
              V_PGI.ZoomOle := false;
             END;
    // 74XXX : ParamTvaTpf(false) ;  TPF non gérée
    74153 : FicheModePaie_AGL('');
    74154 : FicheRegle_AGL ('',False,taModif);
    74155 : FicheDevise('',tamodif,False);
    74156 : ParamTable('ttFormeJuridique',taCreat,1120000,PRien) ;
    74157 : ParamTable('ttCivilite',taCreat,1122000,PRien) ;
    74158 : OuvrePays ;
    74159 : FicheRegion('','',False) ;
    //FV1 : 16/03/2017 -  FS#2452 - Viviane - Ajouter la recherche par Code Postal ou Dépt dans liste Codes Postaux - Ville
    74160 : AGLLanceFiche( 'BTP', 'BTCODEPOSTAL' , '' , '' , '') ; //OuvrePostal(PRien) ;
    74161 : ParamTable('ttLangue',taCreat,0,PRien) ;
    74162 : ParamFont('BLOC') ;
    74163 : BEGIN
    				SaisirHorsInside('BTPARAFFDOC');
            RetourForce:=TRUE;
            END;
{$IFDEF V10}     
    74164 : AglLanceFiche ('BTP','BTCOMMENTAIRE','','','') ;
{$ELSE}
    74164 : HShowMessage('2;?caption?;'+TraduireMemoire('Fonction disponible uniquement à partir de la Version 10 : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)); 
{$ENDIF}
    /// contrats
    14851 : ParametrageTypeAction;
    14852 : ParamTable ('AFTRESILAFF',taCreat,0,PRien,3,'Résiliation contrat / Rejet Devis');
    14853 : ParamTable ('TTPRIOCONTRAT',taCreat,0,PRien);

    ////// Paramétres honoraires & commerciaux
    74201 : AGLLanceFiche ('AFF','AFPROFILGENER','','','TYPE:DEF');
    74202 : AglLanceFiche ('GC','GCCODECPTA','','','') ; // Ventilation comptable GCCOMPTESHT_MUL
    74203 : AGLLanceFiche('GC','GCUNITEMESURE','','','');
    74204 : ParamTable ('GCEMPLOIBLOB',taCreat,0,PRien) ;
    74205 : EntreeArrondi (taModif) ;
    74206 : Begin
              SetParamSOc ('SO_GCAXEANALYTIQUE',TRUE);  //brl 14/10 force paramétrage des axes analytiques
              AGLLanceFiche('BTP','BTVENTILANA','','','') // ventil générique
            end;
    74208 : AGLLanceFiche('GC','GCPORT_MUL','','','') ;   // Port et frais
    74751 : AGLLanceFiche('AFF','AFBLOCAGEAFFAIRE','','','') ;      // blocage affaire
    74211 : ParamTable ('GCCOMMENTAIREENT',taCreat,0,PRien,6);   // commentaires entete
//    74212 : ParamTable ('GCCOMMENTAIRELIGNE',taCreat,0,PRien,6); // commentaires ligne
    74212 : if GetParamSoc('SO_GCCOMMENTAIRE') then AGLLanceFiche ('GC','GCCOMMENTAIRE','','','')
                                              else ParamTable ('GCCOMMENTAIRELIGNE',taCreat,110000167,PRien,6) ;
    74213 : ParamTable ('GCCOMMENTAIREPIED',taCreat,0,PRien,6);  // commentaires pied

    ////// Modèles d'édition
    74251 : EditDocument('L','ADE','ADE',True) ;  // Propositions d'affaire
    74252 : EditDocument('L','AFF','',True) ;  // affaire
    74253 : EditDocument('L','GPI','AFC',True) ;  // Factures

    ////// Paramétres liés aux Fiches de bases
    // Client
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
    //74334 : ParamTable ('AFTRESILAFF',taCreat,0,PRien);
    //74335 : ParamTable ('TTRECONDUCABO',taCreat,0,PRien);
    //74336 : ParamTable ('TTPRIOCONTRAT',taCreat,0,PRien);
    74337 : ParamTable ('AFETATAFFAIRE',taCreat,120000158,PRien);
    // Appel d'offre
    74351 : BEGIN ParamTable ('BTNATUREDOC',taCreat,0,PRien); AvertirTable ('BTNATUREDOC'); END;
    // Articles
    74341..74343 : ParamTable('GCFAMILLENIV'+IntToStr(Num-74340),taCreat,0,PRien) ;
    74344 : BEGIN ParamTable('GCLIBFAMILLE',taModif,0,PRien) ; AvertirTable ('GCLIBFAMILLE'); AfterChangeModule (148); END;
    74345 : ParamTable('GCCOMPTAARTICLE',taCreat,0,PRien) ;
    //74346 : ParamTable('GCTARIFARTICLE',taCreat,0,PRien) ;
    //148210 : FamilleTarifHierarchique;
    74347   : InitialisationTARIF;
    74346   : AglLanceFiche('BTP','BTFAMTARIF','','','') ;
    148210  : AglLanceFiche('BTP','BTVIDEINSIDE','','','FICHE=BTSFAMTARIF') ;
    // Emplacement de stocks
    105139 : ParamTable('GCTYPEEMPLACEMENT',taCreat,0,PRien) ;    // type emplacement
    105140  : ParamTable('GCCOTEEMPLACEMENT',taCreat,0,PRien) ;    // côte emplacement
    // Ressources
    74317 : ParamTable('AFEQUIPERESS', taCreat,0, PRien); //Equipe Ressource
    //74351 : ParamTable('AFTTARIFRESSOURCE',taCreat,0,PRien) ;
    //74352 : AglLanceFiche ('AFF','FONCTION','','','');
    //74353 : AglLanceFiche ('AFF','COMPETENCE','','','');
    //74354 : ParamTable ('AFTNIVEAUDIPLOME',taCreat,0,PRien);
    //74355 : ParamTable ('AFTTYPERESSOURCE',taModif,0,PRien);
    // Calendrier
    74361 : ParamTable ('AFTSTANDCALEN',taCreat,0,PRien);
    74362 : AglLanceFiche ('AFF','JOURFERIE','','','');


    /////// Tables libres //////////
    74402 : AppelTitreLibre('CT',Prien,'ZLT');  // Libellé tables libres client
    74403 : AppelTitreLibre('CM',Prien,'ZLT');  // Libellé mtt libres client
    74404 : AppelTitreLibre('CD',Prien,'ZLT');  // Libellé dates libres client
    74405 : AppelTitreLibre('CC',Prien,'ZLT');  // Libellé textes libres client
    74406 : AppelTitreLibre('CB',Prien,'ZLT');  // Libellé bool libres client
    74407 : AppelTitreLibre('CR',Prien,'ZLT');  // Libellé ressources libres client
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
    74501 : AppelTitreLibre('AT',Prien,'ZLA');  // Libellé tables libres article
    74502 : AppelTitreLibre('AM',Prien,'ZLA');  // Libellé mtt libres article
    74503 : AppelTitreLibre('AD',Prien,'ZLA');  // Libellé dates libres article
    74504 : AppelTitreLibre('AC',Prien,'ZLA');  // Libellé textes libres article
    74505 : AppelTitreLibre('AB',Prien,'ZLA');  // Libellé bool libres article
    74511 : AppelTitreLibre('ET',Prien);  // Libellé tables libres etablissement
    74512 : AppelTitreLibre('EM',Prien);  // Libellé mtt libres etablissement
    74513 : AppelTitreLibre('ED',Prien);  // Libellé dates libres etablissement
    74514 : AppelTitreLibre('EC',Prien);  // Libellé textes libres aetablissement
    74515 : AppelTitreLibre('EB',Prien);  // Libellé bool libres etablissement
    74522 : AppelTitreLibre('VM',Prien);  // Libellé mttlibres apporteur
    74523 : AppelTitreLibre('VD',Prien);  // Libellé dates libres apporteur
    74531 : AppelTitreLibre('BT',Prien,'ZLC');  // Libellé tables libres contact
    74532 : AppelTitreLibre('BM',Prien,'ZLC');  // Libellé mtt libres contact
    74533 : AppelTitreLibre('BD',Prien,'ZLC');  // Libellé dates libres contact
    74534 : AppelTitreLibre('BC',Prien,'ZLC');  // Libellé textes libres contact
    74535 : AppelTitreLibre('BB',Prien,'ZLC');  // Libellé bool libres contact
    74551 : AppelTitreLibre('FT',Prien,'ZLT');  // Libellé tables libres fournissuer
    74552 : AppelTitreLibre('FM',Prien,'ZLT');  // Libellé mtt libres fournisseur
    74553 : AppelTitreLibre('FD',Prien,'ZLT');  // Libellé dates libres fournisseur
    74571 : AppelTitreLibre('TT',Prien);  // Libellé tables libres Tâches

    /////// Tables libres //////////
    74411..74419 : ParamTable('GCLIBRETIERS'+IntToStr(Num-74410),taCreat,0,PRien,6) ;  // Stats clients
    74420 : ParamTable('GCLIBRETIERSA',taCreat,0,PRien,6) ;  // Stats clients
    74461..74469 : ParamTable('AFTLIBREAFF'+IntToStr(Num-74460),taCreat,0,PRien,6) ;   // Stats Affaire
    74489 : ParamTable('AFTLIBREAFFA',taCreat,0,PRien,6) ;   // Stats Affaire
    74431..74439 : ParamTable('AFTLIBRERES'+IntToStr(Num-74430),taCreat,0,PRien,6) ;   // Stats Ressource
    74499 : ParamTable('AFTLIBRERESA',taCreat,0,PRien,6) ;   // Stats Ressource
    74441..74449 : ParamTable('GCLIBREART'+IntToStr(Num-74440),taCreat,0,PRien,6) ;    // Stats Article
    74450 : ParamTable('GCLIBREARTA',taCreat,0,PRien,6) ;    // Stats Article
    74451..74453 : ParamTable('GCLIBREPIECE'+IntToStr(Num-74450),taCreat,0,PRien,6) ;  // Stats Pièces
    74456 : AppelTitreLibre('PT',Prien);  // Libellé tables libres pièces
    74457 : AppelTitreLibre('PD',Prien);  // Libellé dates libres pieces
    74470..74478 : ParamTable('YYLIBREET'+IntToStr(Num-74469),taCreat,AfNoAideTablette('YYLIBREET'+IntToStr(Num-74469)),PRien,6,RechDom('GCZONELIBRE','ET'+IntToStr(Num-74469),FALSE)) ;
    74479 : ParamTable('YYLIBREETA',taCreat,AfNoAideTablette('YYLIBREETA'),PRien,6,RechDom('GCZONELIBRE','ETA',FALSE)) ;
    74541..74549 : ParamTable('YYLIBRECON'+IntToStr(Num-74540),taCreat,0,PRien,6) ;    // Stats Contacts
    74550 : ParamTable('YYLIBRECONA',taCreat,0,PRien,6) ;    // Stats Contacts
    74561..74563 : ParamTable('GCLIBREFOU'+IntToStr(Num-74560),taCreat,AfNoAideTablette('GCLIBREFOU'+IntToStr(Num-74560)),PRien,6,RechDom('GCZONELIBRETIE','FT'+IntToStr(Num-74560),FALSE)) ;  // Stats clients

    /////// Administration
{   74501 : GestionSociete(PRien,@InitSociete) ;     // Gestionnaire société
   74502 : RecopieSoc(PRien,True) ;                 // recopie société à société
}

    /////// Gestion des Utilisateurs
    74701 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;
    //74702 : FicheUserGrp;                           // Groupes d'utilisateurs
    74703 : ReseauUtilisateurs(False) ;             // utilisateurs connectés
    74704 : VisuLog ;                               // Suivi d'activité
    74705 : ReseauUtilisateurs(True) ;              // RAZ connexions

    74999 : if (PRien.InsideForm=nil) then RetourForce := True;
// *****************************************************************************
// *****************************  Module Relation client ***********************
// *****************************************************************************
{$IFDEF GRC}
   92949 : ShowTreeLinks(PRien,'Tablettes hiérarchiques','RT%',FMenuDisp.HDTDrawNode);
   92100..92948,92950..92999 : RTMenuDispatch (Num,PRien);
{$ENDIF}
    //////////////////////////////  Menus PCP Serveur

// *****************************************************************************
// *****************************  Module achat *********************************
// *****************************************************************************
    ////// Saisie des pièces
   // Saisie pièces
   138101 : CreerPiece('CF') ;
   138102 : CreerPiece('BLF') ;
   138103 : CreerPiece('FF') ;
   // modifications
   138111 : AGLLanceFiche('BTP','BTPIECEACH_MUL','GP_NATUREPIECEG=CF','','MODIFICATION')  ;  //commandes fournisseur
   138112 : AGLLanceFiche('BTP','BTPIECEACH_MUL','GP_NATUREPIECEG=BLF','','MODIFICATION') ;  //bons de livraison fournisseur
   // Consultation
   138201 : AGLLanceFiche('BTP','BTPIECEACH_MUL','GP_VENTEACHAT=ACH','','CONSULTATION') ;
   138202 : AGLLanceFiche('BTP','BTLIGNEACH_MUL','GPP_VENTEACHAT=ACH','','CONSULTATION') ;
   // génération
   138301 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_VENTEACHAT=ACH','','') ; // Visa des pièces
   138302 : AGLLanceFiche('GC','GCTRANSACH_MUL','','','BLF') ;   // Livraisons fournisseurs
   138303 : AGLLanceFiche('GC','GCTRANSACH_MUL','','','FF') ;    // LFactures fournisseurs
   138321 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=CF','','BLF') ;
   138322 : AGLLanceFiche('GC','GCGROUPEPIECE_MUL','GP_NATUREPIECEG=BLF','','FF') ;
   // tarif fournisseurs
{$IFNDEF SANSPARAM}
   138401 : EntreeTarifFouArt (taModif) ;                       // saisie
   138402 : AGLLanceFiche('GC','GCTARIFFOUCON_MUL','','','') ;  // consultation
   138403 : AGLLanceFiche('GC','GCTARIFFOUMAJ_MUL','','','') ;  // mise à jour
   // 138404 : AGLLanceFiche('GC','GCPARFOU_MUL','','','') ;  // Import
{$ENDIF}
   // Editions
   138501 : AGLLanceFiche('GC','GCEDITDOCDIFF_MUL','','','ACH') ;     // Editions différées
   138502 : AGLLanceFiche('BTP','BTPTFPIECEACH','','','ACHAT') ;  // Portefeuille pièces
   138511 : AGLLanceFiche('GC','GCART_NONREF','','','')  ;
   138512 : AGLLanceFiche('GC','GCFOURN_ARTICLE','','','') ;
   138513 : AGLLanceFiche('GC','GCARTICLE_FOURN','','','') ;
   // Fichiers
   138601 : AGLLanceFiche('GC','GCFOURNISSEUR_MUL','T_NATUREAUXI=FOU','','') ;
   138602 : AGLLanceFiche('GC','GCCATALOGUE_MUL','','','') ;
   138603 : AGLLanceFiche('YY','YYCONTACTTIERS','T_NATUREAUXI=FOU','','') ;  // contacts
   138604 : AGLLanceFiche('GC','GCTIERS_MODIFLOT','','','FOU') ;  // modification en série des tiers
{$IFNDEF SANSCOMPTA}
   138607 : if _BlocageMonoPoste(False,'',TRUE) then RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'CLI');
   138606 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','OUVRE;CLI') ;  // ouverture compte tiers
   138605 : AGLLanceFiche('RT','RTOUVFERMTIERS','','','FERME;CLI') ;  // fermeture compte tiers
   (*
   138605 : OuvreFermeCpte(fbAux,FALSE,'FOU') ;  // fermeture des comptes tiers
   138606 : OuvreFermeCpte(fbAux,True,'FOU') ;   // ouverture compte tiers
   138607 : SuppressionCpteAuxi('FOU') ;
   *)
{$ENDIF}

    /////// AJOUTS BTP
    145110 : if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
          			AGLLanceFiche('BTP', 'BTAFFAIREINT_MUL','AFF_AFFAIRE=A','','STATUT=AFF')
    				 else
               	AglLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE=A','','NOCHANGESTATUT;STATUT=AFF'); // Affaires
    //
    145120 : AGLLanceFiche('BTP','BTETATSAFFAIRES','AFF_AFFAIRE0=A','','NOCHANGESTATUT;STATUT=AFF');   // Etats sur affaires
    145130 : AGLLanceFiche('BTP','BTRESTDUAFF','GP_NATUREPIECEG=FBT','','NOCHANGESTATUT');  // reste dû par affaires
    // Fiche Qualité 11380
    147810 : AGLLanceFiche ('BTP','BTACCEPTAFF_MUL','AFF_AFFAIRE=A;AFF_ETATAFFAIRE=ENC','','STATUT=AFF'); // Acceptation d'affaire;
    // ---
    145221 : AGLLanceFiche('BTP','BTDEVIS_MUL','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF') ;   // Devis
    145240 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=DBT','','VEN;STATUT=AFF') ;
    145260 : AGLLanceFiche('BTP','BTPTFPIECE','GP_NATUREPIECEG=DBT','','DEVIS;STATUT=AFF') ;  // Portefeuille pièces
    145262 : AGLLanceFiche('GC','GCTIERS_MUL','T_NATUREAUXI=FOU','','');   // Fournisseurs
    145300 : AGLLanceFiche('BTP','BTACCDEV_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;STATUT=AFF') ;   // Devis
    //145320 : AGLLanceFiche('BTP','BTACCDEV_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;ACCEPTINVERSE=X') ;   // Devis
    // NEW LS
    145310 : AGLLanceFiche('BTP','BTDEVIS_MUL','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN;ZEACTION=PLANNIFICATION;AFF_ETATAFFAIRE=ACP;AFF_PREPARE=-','','MODIFICATION') ;   //devis
    145320 : AGLLanceFiche('BTP','BTACCDEV_MUL','GP_NATUREPIECEG=DBT','','MODIFICATION;ACCEPTINVERSE=X;STATUT=AFF') ;   // Devis
    145321 : AGLLanceFiche('BTP','BTREJDEVNEG_MUL','GP_NATUREPIECEG=DBT;AFF_ETATAFFAIRE=ENC','','MODIFICATION;STATUT=AFF') ;
    // ----------- GESTION DES FACTURES/STITUATIONS PROVISOIRES
    145411 : AGLLanceFiche('BTP','BTSITPROV_MUL','','','MODIFICATION') ;
    145413 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FBP;BAC;','','VEN') ;
    145415 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=ABP;','','VEN') ;
    //FV1 : 29/07/2016 - 145414 : AGLLanceFiche('BTP','BTVALSITPROV_MUL','','','') ; (???? pas de nature de piece)
    145414 : AGLLanceFiche('BTP','BTVALSITPROV_MUL','','','VEN') ;
    // -- Gestion des Demande de Prix
    145352 : AGLLanceFiche('BTP','BTENTDDEPRIX_MUL','','','BPP_NATUREPIECEG=DBT;STATUT=AFF') ;   // Demande de prix
    145353 : BEGIN
              SaisirHorsInside('BTPARAMMAIL', 'TYPEMAIL=F');
              retourforce := true;
              exit;
             end; //Envoi de Mail fournisseur
    145354 : AGLLanceFiche('BTP','BTINTEGREDOC','','','TYPEIMPORT=DDE');
    // --
    //145412 : AGLLanceFiche('BTP','BTPREFAC_MUL','','','MODIFICATION;STATUT=AFF') ;   // Prépa facture
    145412 : AGLLanceFiche('BTP','BTSELFAC_MUL','AFF_ETATAFFAIRE=ACP','','MODIFICATION;STATUT=AFF') ;   // Prépa facture

    145421 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF');   //Modif Factures
    145422 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FBT','','VEN;STATUT=AFF') ;
    145423 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','STATUT=AFF;BTP;') ; // Visa des factures
    //FV1 : 15/02/2018 - FS#2947 - ESPACS - Restriction sur le code affaire ce qui empêche certains traitements des pièces
    //145424 : AGLLanceFiche('BTP','BTREGLPIECE_MUL','GP_NATUREPIECEG=FBT','','STATUT=AFF') ; // Enregistrement des reglements non comptabilisés
    145424 : AGLLanceFiche('BTP','BTREGLPIECE_MUL','GP_NATUREPIECEG=FBT','','') ; // Enregistrement des reglements non comptabilisés
    145425 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN;AFF_GENERAUTO=AVA','','MODIFICATION;MODIFDATSIT;STATUT=AFF');   //Modif Factures
    145426 : if (VH_GC.SeriaCoTraitance) OR (VH_GC.SeriaSousTraitance) then AGLLanceFiche('BTP','BTREGLCOTRAIT_MUL','GP_NATUREPIECEG=FBT','','INTERVENANT;STATUT=AFF') ; // Saisie des reglements Cotraitants
    145427 : if (VH_GC.SeriaCoTraitance) OR (VH_GC.SeriaSousTraitance) then AGLLanceFiche('BTP','BTEDTBONPAIEMENT','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','INTERVENANT;MODIFICATION;STATUT=AFF');  //Edition lettre d'éclatement et bon de paiement
    145428 : AGLLanceFiche('BTP','BTRECTSIT_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF;AFF_GENERAUTO=AVA');   //Modif Factures

    145431 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=ABT;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF');   //Modif Avoirs
    145432 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=ABT','','VEN;STATUT=AFF') ;
    145433 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=ABT;GP_VENTEACHAT=VEN','','STATUT=AFF;BTP;') ; // Visa des factures
    //
    145434 : AGLLanceFiche('BTP','BTRECTSIT_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN;AVOIRUNIQUEMENT=X','','MODIFICATION;STATUT=AFF;AFF_GENERAUTO=AVA');   //génération d'avoir depuis facture
    //
    145440 : AGLLanceFiche('BTP','BTPTFPIECE','GP_NATUREPIECEG=FBT','','FACTURE;STATUT=AFF') ;  // Portefeuille pièces
    // Demande d'acompte
    145451 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=DAC;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF') ;   //Modification
    145452 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=DAC','','VEN;STATUT=AFF') ;
    145453 : AGLLanceFiche('BTP','BTREGLPIECE_MUL','GP_NATUREPIECEG=DAC','','STATUT=AFF') ; // Enregistrement des reglements

    // Fin de travaux
//    145460 : AGLLanceFiche ('BTP','BTCLOTURE_MUL','GP_NATUREPIECEG=DBT','','MEMOIRE');
    145460 : AGLLanceFiche ('BTP','BTSELFAC_MUL','AFF_ETATAFFAIRE=ACP','','CLOTURE;STATUT=AFF');
    145461 : AGLLanceFiche ('BTP','BTIMPDGD_MUL','AFF_ETATAFFAIRE=TER','','STATUT=AFF');
    // verification de la facturation
    145470 : AGLLanceFiche('BTP','BTISFACTURE','GP_NATUREPIECEG=DBT','','MODIFICATION');

		// Etats spécifiques
    145491 : AGLLanceFiche('BTP','BTETATRECAP','GP_NATUREPIECEG=FBT','','') ;  // etat recap HLM DURAND SAS
    // 
    145492 : AGLLanceFiche('BTP','BTVISAPIECES_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','STATUT=AFF;BTP;') ; // Visa des factures type BTP
    // Contre etudes
    145710 : AGLLanceFiche('BTP','BTDEVIS_MUL','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN;ZEACTION=GENCONTRETU;AFF_ETATAFFAIRE=ACP;AFF_PREPARE=-','','MODIFICATION;STATUT=AFF') ;   //devis
    145720 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=BCE;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF');   //Modif Contre etude
    // Visa des pièces
    145910 : AGLLanceFiche('BTP','BTVISAPIECES_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','STATUT=AFF;') ; // Visa des pièces; // Mise en place du visa
    145920 : AGLLanceFiche('BTP','BTVISAPIECES_MUL','GP_NATUREPIECEG=FBT;GP_VENTEACHAT=VEN','','STATUT=AFF;RESTORE') ; // enlever le visa
    // Appel d'offre
    145510 : if (VH_GC.BTSeriaAO = TRUE) or (V_PGI.VersionDemo = True) then
                begin
                if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
                   AGLLanceFiche('BTP', 'BTAFFAIREINT_MUL','AFF_AFFAIRE=P','','STATUT=PRO')
                 else
                 	 AglLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE=P','','STATUT=PRO');   // Appel offre
                end
             else
             		 PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Récupération des Appels d''Offres');
    145520 : if (VH_GC.BTSeriaAO = TRUE) or (V_PGI.VersionDemo = True) then
    						AglLanceFiche('BTP','BTAPPOFF_MUL','AFF_AFFAIRE=P','','STATUT=PRO;ETAT=ENC;APPOFFACCEPT')   // Appel offre
						 else
             		PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Récupération des Appels d''Offres');
    145530 : if (VH_GC.BTSeriaAO = TRUE) or (V_PGI.VersionDemo = True) then
    						AGLLanceFiche('BTP','BTETATSAFFAIRES','AFF_AFFAIRE0=P','','STATUT=PRO')   // Etats sur affaires - Appels D'offre
             else
             		PgiBox('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Récupération des Appels d''Offres');
    // Etudes
    145610 : AGLLanceFiche('BTP','BTDEVIS_MUL','GP_NATUREPIECEG=ETU;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF') ;   //Etudes
    145620 : AGLLanceFiche('BTP','BTACCDEV_MUL','GP_NATUREPIECEG=ETU','','MODIFICATION;STATUT=AFF') ;   //acceptation études
    145830 : AGLLanceFiche('BTP','BTTRANSPIECE_MUL','GP_NATUREPIECEG=CBT;GP_VIVANTE=X','','LBT') ;  // Livraisons
    145840 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=LBT','','VEN') ;
    145850 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=LBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    145860 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=CBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    145870 : BEGIN RecalculLivrNonValorise ; RetourForce:=True; END;
//    145810 : AGLLanceFiche('GC','GCTRANSPIECE_MUL','GP_NATUREPIECEG=BLC;GP_VIVANTE=X','','ALV') ;  // Retour chantier
   	145810 : AGLLanceFiche('BTP','BTDUPLICPIECE_MUL','GP_NATUREPIECEG=LBT;GP_VENTEACHAT=VEN','','BFC') ;
    145820 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=BFC;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    //    Module de facturation "Comptoir" (Sur stock)
    145151 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FBC;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION;STATUT=AFF');
    145152 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=FBC','','VEN;VIVANTE;STATUT=AFF') ;
    145153 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=ABC;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION;STATUT=AFF');
    145154 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=ABC','','VEN;VIVANTE;STATUT=AFF') ;
    145155 : AGLLanceFiche('BTP','BTPTFPIECE','GP_NATUREPIECEG=FBC','','COMPTOIR;STATUT=AFF') ;  // Portefeuille pièces de type facture / stock

    145156 : AGLLanceFiche('BTP','BTVISAPIECES_MUL','GP_NATUREPIECEG=FBC;GP_VENTEACHAT=VEN','','STATUT=AFF;NEGOCE;') ; // Visa des factures type negoce

    // - FACTURE D'ACOMPTES
    145481 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=B00;GP_VENTEACHAT=VEN','','MODIFICATION;STATUT=AFF');   //Modif Factures
    145482 : AGLLanceFiche('BTP','BTEDITDOCDIFF_MUL','GP_NATUREPIECEG=B00','','VEN;STATUT=AFF') ;
    145483 : AGLLanceFiche('BTP','BTPIECEVISA_MUL','GP_NATUREPIECEG=B00;GP_VENTEACHAT=VEN','','STATUT=AFF') ; // Visa des factures

    // ---------------
    145955 : BEGIN
              V_PGI.ZoomOle := true;
              AGLLanceFiche('BTP','BTGETREGLCPTA','','','') ;  // récupéaryion des reglemenst depuis la compta
              V_PGI.ZoomOle := false;
              retourforce := True;
             end;
    // ------------------
    // SPIGAO
    145961 : Begin
               if Not SPIGAOOBJ.connected then
               begin
               	SPIGAOOBJ.Infos
               end else
               begin
                 if SPIGAOOBJ.IsNewGestion then
                 begin
                   SPIGAOOBJ.ChargeNewAffaires;
                end else
                 begin
                 		SPIGAOOBJ.VoirAffairesNew;
                 end;
               end;
               RetourForce:=True;
             end;
    145962 : Begin
               if Not SPIGAOOBJ.connected then
               begin
               	SPIGAOOBJ.Infos;
               end else
               begin
               	SPIGAOOBJ.VoirAffairesMod;
               end;
               RetourForce:=True;
             End;
    145963 : Begin
               if Not SPIGAOOBJ.connected then SPIGAOOBJ.Infos
               else
               begin
                 SPIGAOOBJ.IntegreDie;
               end;
               RetourForce:=True;
             End;
    145964 : Begin
               if Not SPIGAOOBJ.connected then SPIGAOOBJ.Infos
                                          else SPIGAOOBJ.VoirMessages;
  						 FMenuG.PRien.Refresh;
               RetourForce:=True;
             end;
    145965 : Begin
               if Not SPIGAOOBJ.connected then SPIGAOOBJ.Infos
                                          else SPIGAOOBJ.LienBTP;
							  FMenuG.PRien.Refresh;
               RetourForce:=True;
             end;
    145967 : Begin
               if Not SPIGAOOBJ.connected then SPIGAOOBJ.Infos
               else
               begin
                 SPIGAOOBJ.ConfDetection;
               end;
               RetourForce:=True;
             end;
    145968 : Begin
               if Not SPIGAOOBJ.connected then SPIGAOOBJ.Infos
                                          else SPIGAOOBJ.ConfTraitement;
  							FMenuG.PRien.Refresh;
               RetourForce:=True;
             end;
    145969 : Begin
               if Not SPIGAOOBJ.connected then SPIGAOOBJ.Infos
                                          else SPIGAOOBJ.APropos;
 							 FMenuG.PRien.Refresh;
               RetourForce:=True;
             end;

    //
    // Consultation de pièce
    //283210 : AGLLanceFiche('GC','GCPIECE_MUL','GP_NATUREPIECEG=CBT','','CONSULTATION') ;
    //AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=CBT','','CONSULTATION') ;
    283210 : AGLLanceFiche('BTP','BTCONSULTPIEC_MUL','GP_NATUREPIECEG=CBT','','CONSULTATION') ;
    // Consultation des lignes
    283220 : AGLLanceFiche('BTP','BTLIGNE_MUL','GL_NATUREPIECEG=CBT','','CONSULTATION') ;
    //283220 : AGLLanceFiche('GC','GCLIGNE_MUL','GL_NATUREPIECEG=CBT','','CONSULTATION') ;
    // GESTION DE CHANTIER
    // NEW LS
//  147110 : SaisieConsommation;
    147110 : BEGIN SaisirHorsInside('SAISIECONSOMMATIONS'); RetourForce:=TRUE; END;
    147120 : AGLLanceFiche('BTP','BTJOUCON','','','') ;
    147130 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=PBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X;AVANC=X','','MODIFICATION') ;
    147140 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=FF;GP_VENTEACHAT=ACH','','MODIFICATION') ;  //factures fournisseur
    147145 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=AF;GP_VENTEACHAT=ACH','','MODIFICATION')  ;  //Avoir financier
    147151 : ImportConsoInnovPro;
    147152 : AnnulImportConsoInnovPro;
    // Ajout des saisies via Modèles
    147161 : BEGIN SaisirHorsInside ('BTSAISIEHEBDO'); RetourForce := true; END;
    147162 : BEGIN SaisirHorsInside ('BTMODELHEBDO'); RetourForce := true; END;
    147163 : BEGIN
              // saisie des avancements POC
              AGLLanceFiche('BTP','BTSELPIECE_POC','GP_NATUREPIECEG=BCE;GP_VENTEACHAT=VEN;GP_VIVANTE=X;AVANCPOC=X','','MODIFICATION') ;
             END;
    //
//uniquement en line
//  147211 : AGLLanceFiche('BTP','BTDEVIS_MUL_S1','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN;ZEACTION=PLANNIFICATION;AFF_ETATAFFAIRE=ACP;AFF_PREPARE=-','','MODIFICATION') ;   //devis
    147211 : AGLLanceFiche('BTP','BTPREPACHANTIER','GP_NATUREPIECEG=DBT;GP_VENTEACHAT=VEN;ZEACTION=PLANNIFICATION;AFF_ETATAFFAIRE=ACP;AFF_PREPARE=-','','MODIFICATION;STATUT=AFF') ;   //devis
    147212 : AGLLanceFiche('BTP','BTPREPACHANTIER','GP_NATUREPIECEG=BCE;GP_VENTEACHAT=VEN;ZEACTION=PLANNIFICATION;AFF_ETATAFFAIRE=ENC;AFF_PREPARE=-','','MODIFICATION;STATUT=AFF') ; // Validation
    147220 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=PBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION;STATUT=AFF') ;
    147310 : GenereCommandeChantier;
    147320 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=CBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION;AVANC');
    147330 : AGlLanceFiche('BTP','BTCONTROLECBT','','','MODIFICATION');

    147410 : AGLLanceFiche('BTP','BTRESULTATS','','','') ;
    147420 : AGLLanceFiche('BTP','BTTABLEAUBORD','','','') ; //CHANGESTATUT;
    147430 : AGLLanceFiche('BTP','BTCOMPARATIF','','','') ;
    147440 : AGLLanceFiche ('BTP','BTSAISRESTDEP_MUL','AFF_AFFAIRE=A;AFF_ETATAFFAIRE=ACP','','STATUT=AFF'); // Saisie des reste à dépenser par affaire
    //
    147510 : BTLanceFiche_ActivGenerPaie; //Génération de l'activité vers la paie
    147520 : BTLanceFiche_AnnulGenerPaie; //Annulation des lignes d'activité générées vers la paie
    147530 : AFLanceFiche_MulActivitePaie; // Paramétrage lien Activité avec Rubriques Paie
    147610 : BEGIN SaisirHorsInside ('BTREGENERECONSO'); RetourForce := True; END;
    147620 : BEGIN DeblocageVerrouValidation; RetourForce := True; END;
    147630 : AGLLanceFiche('BTP','BTTRSCONSO_MUL','','','') ;
    147640 : AGLLanceFiche('BTP','BETATPREPPAIE','','','') ;
    147650 : AGLLanceFiche('BTP','BTGENODANAL','','','') ;
    // -- VISUAL PROJETS
//    147713 : AGLLanceFiche('BTP','BTVISUALP_MUL','GP_NATUREPIECEG=PBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION') ;


    //PARAMETRES
    148100 : AglLanceFiche('BTP','BTVARIABLE','A;APPLICATION','','APPLICATION');
    //FV1
    65923  : AglLanceFiche('BTP','BTJOURFERIE','','','');
    148110 : CCLanceFiche_LongueurCompte;
    148120 : GenereAffaireFromAnal;
    148251 : begin
              Souche := GetSoucheG('DBT','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=DBT;SOUCHE='+Souche+';NUMERO=0') ;
             end;
    148252 : BEGIN
              Souche := GetSoucheG('FBT','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=FBT;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148253 : BEGIN
              Souche := GetSoucheG('ABT','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=ABT;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148254 : BEGIN
              Souche := GetSoucheG('ETU','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=ETU;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148255 : BEGIN
              Souche := GetSoucheG('DAC','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=DAC;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148256 : BEGIN
              Souche := GetSoucheG('FBC','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=FBC;SOUCHE='+Souche+';NUMERO=0') ;  // facture de négoce
             END;
    148257 : BEGIN
              Souche := GetSoucheG('ABC','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=ABC;SOUCHE='+Souche+';NUMERO=0') ;  // Avoir sur facture de négoce
             END;
    148258 : BEGIN
              Souche := GetSoucheG('BCE','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=BCE;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    //
    148259 : BEGIN
              Souche := GetSoucheG('DE','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=DE;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148260 : BEGIN
              Souche := GetSoucheG('CC','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=CC;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148261 : BEGIN
              Souche := GetSoucheG('BLC','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=BLC;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148262 : BEGIN
              Souche := GetSoucheG('FBP','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=FBP;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148263 : BEGIN
              Souche := GetSoucheG('ABP','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=ABP;SOUCHE='+Souche+';NUMERO=0') ;
             END;
    148264 : BEGIN
              Souche := GetSoucheG('B00','','');
              AGLLanceFiche('BTP','BTPARIMPDOC','','','NATURE=B00;SOUCHE='+Souche+';NUMERO=0') ;
             END;

    148651 : ParamTable('GCTYPECOMMERCIAL',taCreat,0,PRien) ;
    148661 : ParamTable('AFTTYPEHEURE',taCreat,0,PRien) ;
    148662 : AFLanceFiche_Fonction('','ACTION=MODIFICATION') ;
    148802 : AglLanceFiche('BTP','BTIMBRESPMUL','','','ACTION=MODIFICATION');
    //
    148901..148903 : ParamTable('BTFAMILLEOUV'+IntToStr(Num-148900),taCreat,0,PRien) ;
    148904 : BEGIN ParamTable('BTLIBOUVRAGE',taModif,0,PRien) ; AvertirTable ('GCLIBFAMILLE'); AfterChangeModule (148); END;
    148905 : BTLanceFicheParamWSCegid('BTP', 'BTPARAMWS', '', '', '');
    148906 : OpenForm.WSCreationAllowedTable;
    //
    // BIBLIOTHEQUE
		149204 : AGLLanceFiche('BTP','BBPRIXMATFAM_MUL','','','ACTION=MODIFICATION');
    149205 : AGLLanceFiche('BTP','BBPRIXMAT_MUL','','','ACTION=MODIFICATION');
    149210 : AGLLanceFiche('BTP','BTARTICLE_MUL','','','MAR');
    149220 : AGLLanceFiche('BTP','BTARTICLE_MUL','GA_TYPEARTICLE=PRE','','PRE');


//  149230 : AGLLanceFiche('BTP','BTARTICLE_MUL','','','NOM');
    149230 : AGLLanceFiche('BTP','BTOUVRAGE_MUL','','','');
//uniquement en line
//  149240 : AglLanceFiche ('BTP','BTPROFILART_S1','','','') ;      // Profil article
    149240 : AglLanceFiche ('BTP','BTPROFILART','','','') ;      // Profil article
    149245 : AGLLanceFiche('BTP', 'BTVARIABLE', 'G;GENERAL','', 'GENERALES');
    149250 : begin
              if not IsMasterOnShare('GCFAMILLENIV1') then
              begin
                PgiInfo ('Vous devez réaliser cette opération sur la base principale');
                RetourForce:=True;
              end else
              begin
                //AGLLanceFiche('GC','GCFAMHIER','','','ACTION=MODIFICATION;NIV1=FN1;NIV2=FN2;NIV3=FN3') ;
                AGLLanceFiche('BTP','BTFAMHIER','','','ACTION=MODIFICATION;NIV1=FN1;NIV2=FN2;NIV3=FN3') ;
              end;
             end;
    149251 : begin
              if not IsMasterOnShare('GCFAMILLENIV1') then
              begin
                PgiInfo ('Vous devez réaliser cette opération sur la base principale');
                RetourForce:=True;
              end else
              begin
                //AGLLanceFiche('GC','GCFAMHIER','','','ACTION=MODIFICATION;NIV1=BO1;NIV2=BO2;NIV3=BO3') ;
                AGLLanceFiche('BTP','BTFAMHIER','','','ACTION=MODIFICATION;NIV1=BO1;NIV2=BO2;NIV3=BO3') ;    
              end;
             end;
    149260 : AglLancefiche ('BTP','BTNATPREST_MUL','','','');
    149271 : BEGIN AgLlanceFIche('BTP','BDOMACT_MUL','','',''); AvertirTable ('BTDOMAINEACT'); END;
    149290 : AGLLanceFiche('BTP','BTARTICLE_MUL','GA_TYPEARTICLE=FRA','','FRA');
    149292 : AGLLanceFiche('GC','GCMULSUPPRART','','','') ;  // épuration articles
    149295 : AGLLanceFiche('GC','GCTARIFART_MUL','','','') ;  // Maj BAses Tarifaires
    149296 : AGLLanceFiche('BTP','BTARTICLE_MODIFLO','','','') ;  // modification en série
    // Dimensions ////
    149110 : begin ParamTable('GCCATEGORIEDIM',taModif,0,PRien,3); AfterChangeModule(149); end;  // categorie
    149130 : AGLLanceFiche('GC','GCMASQUEDIM','','','') ;
    149150 : AGLLanceFiche('MBO','TYPEMASQUE','','','');  // Type de masque

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
    149351 : AGLLanceFiche('RT','RTPROSPECT_MUL','T_NATUREAUXI=PRO;T_NATUREAUXI_=PRO','','');
    149352 : if _BlocageMonoPoste(False,'',TRUE) then RTLanceFiche_RTSupprTiers ('RT', 'RTSUPPRTIERS', '', '', 'PRO');

// Ressources
//uniquement en line
{*
    149930 : AFLanceFiche_Mul_RessourceLine('TYPERESSOURCE=SAL');   // Salaries
    149940 : AFLanceFiche_Mul_RessourceLine('TYPERESSOURCE=ST');   // Sous-Traitants
    149950 : AFLanceFiche_Mul_RessourceLine('TYPERESSOURCE=INT');   // Intérims
    149960 : AFLanceFiche_Mul_RessourceLine('TYPERESSOURCE=MAT');   // Matériels
    149970 : AFLanceFiche_Mul_RessourceLine('TYPERESSOURCE=OUT');   // Outils
    149980 : AFLanceFiche_Mul_RessourceLine('TYPERESSOURCE=LOC');   // Location
*}
    149910 : AFLanceFiche_Mul_Ressource;   // Assistants
    149920 : AFLanceFiche_Synchro_RessSalarie; // Synchro ressources / salariés
    149710 : DuplicBordereau; // duplication de bordereau
// Utilitaires
    149610 : AGLLanceFiche('BTP','BTARTICLE_RECH','','','MODIFCODE');
    149620 : ReinitAncOuvrage;
    149640 : Begin ChangementCodeTarif; RetourForce:=True; End;
    //FV1 : 07/01/2014 - Etat Préparatoire de TVA
    7647   : AGLLanceFiche('CP','CPQRTVACTRL','','','');            // Etat de contrôle TVA


//uniquement en line
    149630 :Begin AppelParamSociete; RetourForce:=True; End;
    149297 : BEGIN CommandeBiblioBatiprix; RetourForce := true; END;
    149298 : BEGIN IntegrationDonneeBatibrix; RetourForce := true; END;
//  49298 : AglLanceFiche('BTP','BATIPRIXINTERF','','','');
		149299 : AglLanceFiche('BTP','BTREFCDEBATIPRIX','','','');
    // ----
//  149999 : GCvoirtob(Nil); // Uniquement pour pourvoir utiliser le voirtob en debug
    150111 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=BLF;GP_VENTEACHAT=ACH;GP_VIVANTE=X','','MODIFICATION');
    150112 : AGLLanceFiche('BTP','BTPIECE_MUL','GP_NATUREPIECEG=LBT;GP_VENTEACHAT=VEN;GP_VIVANTE=X','','MODIFICATION');
    219021 : AGLLanceFiche('BTP','BTRECALCPIECE_MUL','GP_NATUREPIECEG=DBT;GP_VIVANTE=X;GP_VENTEACHAT=VEN','','MODIFICATION');
    // PARC
    323210 : RTLanceFiche_PARCTIERS_MUL('BTP','BTPARCTIERS_MUL','T_NATUREAUXI=CLI','','GC');
    323120 : AGLLancefiche('BTP','BTARTPARC_MUL','','','ACTION=MODIFICATION');
    323151..323153 : ParamTable('BTFAMILLEPARC'+IntToStr(Num-323150),taCreat,0,PRien) ;
    323131 : AGLLancefiche('BTP','BARTVERSIONNABLE','','','ACTION=MODIFICATION');
    323132 : BEGIN ParamTable('BTTYPEVERSION',taCreat,0,PRien) ; AvertirTable ('BTTYPEVERSION'); END;
    323140 : BEGIN ParamTable('BTLIBARTPARC',taModif,0,PRien) ; AvertirTable ('BTLIBARTPARC'); AfterChangeModule (323); END;
    //----- GESTION DE LA COTRAITANCE ----
    325110..325999 : BEGIN
    								   if (GestionMenuCotraitance (Num) < 0) then
                       begin
                         retourforce := true;
                         exit;
                       end;
    							   END;
    //----- GESTION DE LA SOUS-TRAITANCE ----
    327110..327999 : BEGIN
    								   if (GestionMenuSousTraitance (Num) < 0) then
                       begin
                         retourforce := true;
                         exit;
                       end;
    							   END;
    // ----- Gestion Menus Negoce de matériaux
    328110..328999 : BEGIN
    								   if (GestionMenuNegoce (Num) < 0) then
                       begin
                         retourforce := true;
                         exit;
                       end;
    							   END;
    // ----- Gestion Menus Gestion Parc Matériel
    329001..329999 : BEGIN
    								   if (GestionMenuParcMateriel (Num) < 0) then
                       begin
                         retourforce := true;
                         exit;
                       end;
    							   END;

    // -- DEBUT PLAN DE CHARGE -¨PLANNING
    331110 :
      begin
        OldValue := V_Pgi.ZoomOle ;
        V_Pgi.ZoomOle := True ;
        try
          Saisieplanning ('PRA', 0, taModif) ;
        finally
          V_Pgi.ZoomOle := OldValue ;
        end;
      end;
    331210 :
      begin
        OldValue := V_Pgi.ZoomOle ;
        V_Pgi.ZoomOle := True ;
        try
          Saisieplanning ('PRA', 0, taConsult) ;
        finally
          V_Pgi.ZoomOle := OldValue ;
        end;
      end;
    331310 : AGLLanceFiche('BTP','BTAFFAIRE_SEL','','','ACTION=MODIFICATION;MODEPC=CHANTIER');
    331320 : AGLLanceFiche('BTP','BTAFFAIRE_SEL','','','ACTION=MODIFICATION;MODEPC=FONCTION');
    331330 : AGLLanceFiche('BTP','BTPREVFACT_MUL','AFF_ETATAFFAIRE=ACP','','');
    331340 : AGLLanceFiche('BTP','BTRESTITUPREVFACT','AFF_ETATAFFAIRE=ACP','','');

    331410 : Agllancefiche('BTP','BTPARAMPLAN_MUL', 'HPP_MODEPLANNING=PRA', '', 'TYPEPLANNING=PCA') ;
    331510 : AGLLanceFiche('BTP','BTTYPEACTION_MUL','','','TYPEACTION=PCA');
    // -- FIN PLAN DE CHARGE -¨PLANNING

    60999 : BEGIN AglLanceFiche ('BTP','BTESTFORMULE','','',''); END;

    { RGPD }
    {  Export données personnelles }
    280011 : OpenForm.RGPDWindows(rgpdpThird   , -280010, Num, rgpdaDataExport);        // Export données personnelles - Clients
    280012 : OpenForm.RGPDWindows(rgpdpResource, -280010, Num, rgpdaDataExport);        // Export données personnelles - Ressources
    280013 : OpenForm.RGPDWindows(rgpdpUser    , -280010, Num, rgpdaDataExport);        // Export données personnelles - Utilisateurs
    280014 : OpenForm.RGPDWindows(rgpdpSuspect , -280010, Num, rgpdaDataExport);        // Export données personnelles - Suspects
    280015 : OpenForm.RGPDWindows(rgpdpContact , -280010, Num, rgpdaDataExport);        // Export données personnelles - Contacts
    {  Droit à l'oubli }
    280021 : OpenForm.RGPDWindows(rgpdpThird   , -280020, Num, rgpdaAnonymization);     // Droit à l'oubli - Clients
    280022 : OpenForm.RGPDWindows(rgpdpResource, -280020, Num, rgpdaAnonymization);     // Droit à l'oubli - Ressources
    280023 : OpenForm.RGPDWindows(rgpdpUser    , -280020, Num, rgpdaAnonymization);     // Droit à l'oubli - Utilisateurs
    280024 : OpenForm.RGPDWindows(rgpdpSuspect , -280020, Num, rgpdaAnonymization);     // Droit à l'oubli - Suspects
    280025 : OpenForm.RGPDWindows(rgpdpContact , -280020, Num, rgpdaAnonymization);     // Droit à l'oubli - Contacts
    {  Demande de rectification }
    280031 : OpenForm.RGPDWindows(rgpdpThird   , -280030, Num, rgdpaDataRectification); // Demande de rectification - Clients
    280032 : OpenForm.RGPDWindows(rgpdpResource, -280030, Num, rgdpaDataRectification); // Demande de rectification - Ressources
    280033 : OpenForm.RGPDWindows(rgpdpUser    , -280030, Num, rgdpaDataRectification); // Demande de rectification - Utilisateurs
    280034 : OpenForm.RGPDWindows(rgpdpSuspect , -280030, Num, rgdpaDataRectification); // Demande de rectification - Suspects
    280035 : OpenForm.RGPDWindows(rgpdpContact , -280030, Num, rgdpaDataRectification); // Demande de rectification - Contacts
    {  Demande de consentement }
    280042 : OpenForm.RGPDWindows(rgpdpThird   , -280040, Num, rgdpaConsentRequest);    // Demande de consentement - Clients
    280045 : OpenForm.RGPDWindows(rgpdpResource, -280040, Num, rgdpaConsentRequest);    // Demande de consentement - Ressources
    280048 : OpenForm.RGPDWindows(rgpdpUser    , -280040, Num, rgdpaConsentRequest);    // Demande de consentement - Utilisateurs
    280051 : OpenForm.RGPDWindows(rgpdpSuspect , -280040, Num, rgdpaConsentRequest);    // Demande de consentement - Suspects
    280054 : OpenForm.RGPDWindows(rgpdpContact , -280040, Num, rgdpaConsentRequest);    // Demande de consentement - Contacts
    {  Retour demande de consentement }
    280043 : OpenForm.RGPDWindows(rgpdpThird   , -280040, Num, rgdpaConsentResponse);   // Retour demande de consentement - Clients
    280046 : OpenForm.RGPDWindows(rgpdpResource, -280040, Num, rgdpaConsentResponse);   // Retour demande de consentement - Ressources
    280049 : OpenForm.RGPDWindows(rgpdpUser    , -280040, Num, rgdpaConsentResponse);   // Retour demande de consentement - Utilisateurs
    280052 : OpenForm.RGPDWindows(rgpdpSuspect , -280040, Num, rgdpaConsentResponse);   // Retour demande de consentement - Suspects
    280055 : OpenForm.RGPDWindows(rgpdpContact , -280040, Num, rgdpaConsentResponse);   // Retour demande de consentement - Contacts
    {  Référentiel }
    280061 : BLanceFiche_RGPDReferentiel('BTP', frm_RGPDRepository , '', '', OpenForm.SetWindowCaption(-280060, Num)); // Référentiel
    // --
    else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

Procedure DispatchTT ( Num : Integer ; Action : TActionFiche ; Lequel,TT, Range : String) ;
var TypeRessource	: String;
		TypeArticle		: String;
		TypeTiers			:	String;

    ArticleAff		:string;
    Arg5LanceFiche:string;
    ChampMul			: String;
    ValMul				: String;
    Critere 			: string;
    Arguments 		: String;

    X							: Integer;

    PresentTypeArt: Boolean;
    PresentGATypeArt : boolean;

BEGIN
  case Num of

//    1 : {Compte gene} FicheGene(Nil,'',LeQuel,Action,0) ;
//    2 : {Tiers compta} FicheTiers(Nil,'',LeQuel,Action,1) ;
//    4 : {Journal} FicheJournal(Nil,'',Lequel,Action,0) ;
    5 : {affaires}
        BEGIN
        if Lequel[1] = 'W' then AGLLanceFiche ('BTP','BTAPPELINT','','','CODEAPPEL:' + Lequel +';ACTION=CONSULTATION')
        else if Lequel[1] = 'I' then AGLLanceFiche('BTP','BTAFFAIRE','',Lequel,'STATUT:INT;ACTION=CONSULTATION')
        else AGLLanceFiche('BTP','BTAFFAIRE','',Lequel,'ACTION=CONSULTATION');
        END;
//Ressources
    6 : BEGIN
        TypeRessource := '';
        Arg5LanceFiche:= ActionToString(Action);
        TypeRessource := GetArgumentValue(Range, 'TYPERESSOURCE');
        if (Range <> '') then Arg5LanceFiche:= Arg5LanceFiche + ';' + Range;
        if Typeressource='' then Range := range + ';RESSOURCE="SAL"';

        if AGLJaiLeDroitFiche(['RESSOURCE',ActionToString(Action)],2) Then
           Begin
//uniquement en line
{*
	         If TypeRessource = 'SAL' then AGLLanceFiche('BTP','BTRESSOURCE_S1','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'ST' then AGLLanceFiche('BTP','BTRESSOURCE_S1','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'LOC' then AGLLanceFiche('BTP','BTMATERIEL_S1','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'INT' then AGLLanceFiche('BTP','BTRESSOURCE_S1','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'MAT' then AGLLanceFiche('BTP','BTMATERIEL_S1','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'OUT' then AGLLanceFiche('BTP','BTMATERIEL_S1','',Lequel,Arg5LanceFiche)
           Else AGLLanceFiche('BTP','BTRESSOURCE_S1','',Lequel,Arg5LanceFiche) ;
*}
	         If (TypeRessource = 'SAL') then AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'ST' then AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'LOC' then AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'INT' then AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'MAT' then AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche)
           Else If TypeRessource = 'OUT' then AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche)
           Else AGLLanceFiche('BTP','BTRESSOURCE','',Lequel,Arg5LanceFiche);
           end;

        END;
//Articles
    7 : BEGIN
    		presentTypeArt := false;
        presentGATypeArt := false;
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
                     if (ChampMul='TYPEARTICLE') then
                     begin
                     	TypeArticle:=ValMul ;// AC:traitment mar ou nom
                      presentTypeArt := true;
                     end else if (ChampMul = 'GA_TYPEARTICLE') then
                     begin
                     	TypeArticle:=ValMul ;// AC:traitment mar ou nom
                      presentGATypeArt := true;
                     end;
                  end;
               end;
           until  Critere='';
        end;
        if Copy(TypeArticle,1,1)='''' then
        begin
        	TypeArticle := Copy(TypeArticle,2,length(TypeArticle)-2);
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
        if (presentGATypeArt) and (not presentTypeArt) then Arg5LanceFiche:= Arg5LanceFiche + ';TYPEARTICLE='+TypeArticle;
        // PIL le 16/06/2000 car le monofiche ne convient pas aux MUL
        if (IsArticleSpecif(TypeArticle)='FicheAffaire') then
            AGLLanceFiche('AFF', 'AFARTICLE', '', ArticleAff, Arg5LanceFiche)
        else
            // Modif BTP
            if (IsArticleSpecif(TypeArticle)='FicheBtp') then
            begin
            		if (copy(TypeArticle,1,2)='PA') then
                	AGLLanceFiche('BTP', 'BTARTPARC', '', ArticleAff, Arg5LanceFiche)
                else if TypeArticle = 'POU' then
                	AGLLanceFiche('BTP', 'BTARTPOURCENT', '', ArticleAff, Arg5LanceFiche)
                else if TypeArticle = 'FRA' then
//uniquement en line
//                AGLLanceFiche('BTP', 'BTARTICLE_S1', '', ArticleAff, Arg5LanceFiche)
                	AGLLanceFiche('BTP', 'BTPRESTATION', '', ArticleAff, Arg5LanceFiche)
                else if TypeArticle = 'PRE' then
//uniquement en line
//              	AGLLanceFiche('BTP', 'BTARTICLE_S1', '', ArticleAff, Arg5LanceFiche)
                	AGLLanceFiche('BTP', 'BTPRESTATION', '', ArticleAff, Arg5LanceFiche)
                else
//uniquement en line
//              	AGLLanceFiche('BTP', 'BTARTICLE_S1', '', ArticleAff, Arg5LanceFiche);
                	AGLLanceFiche('BTP', 'BTARTICLE', '', ArticleAff, Arg5LanceFiche);
						end else AGLLanceFiche('GC', 'GCARTICLE', '', ArticleAff, Arg5LanceFiche) ;
        END ;

    8 : {clients}
        begin
        Arg5LanceFiche := '';
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
                     if ChampMul='T_NATUREAUXI' then TypeTiers:=ValMul ;
                  end;
               end;
        until  Critere='';
        if Range <> '' then Arg5LanceFiche := Range + ';'+ActionToString(Action)+';MONOFICHE'
                       else Arg5LanceFiche :=  ActionToString(Action)+';MONOFICHE';
        if (TypeTiers <> 'FOU') or (TypeTiers='') or (TypeTiers='PRO') then AGLLanceFiche('GC','GCTIERS','',Lequel,Arg5LanceFiche)
                                                  										 else AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,Arg5LanceFiche);
        end ;
    9 : {commerciaux} AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)) ;
   10 : {conditionnement} AGLLanceFiche('GC','GCCONDITIONNEMENT','',Lequel,ActionToString(Action)) ;
   11 : {catalogue} AGLLAnceFiche('GC','GCCATALOGU_NVFOUR','',Lequel,ActionToString(Action)) ;
//   12 : {fournisseurs} AGLLanceFiche('GC','GCTIERS','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=FOU') ;
   12 : BEGIN {fournisseurs}
   				AGLLanceFiche('GC','GCFOURNISSEUR','',Lequel,ActionToString(Action)+';MONOFICHE;T_NATUREAUXI=FOU') ;
   			END;
   13 : {dépots} AGLLanceFiche('BTP','BTDEPOT','',Lequel,ActionToString(Action)+';MONOFICHE') ;
   14 : {apporteurs} AGLLanceFiche('GC','GCCOMMERCIAL','',Lequel,ActionToString(Action)) ;
   15 : {Règlement}  FicheRegle_AGL ('',False,taModif) ;
   16 : {contacts} affDispatchTT ( Num , Action , Lequel,TT, Range ) ;
   17 : {Fonction des ressources} If AglJaiLeDroitFiche (['FONCTION',ActionToString(Action)],2) then AFLanceFiche_Fonction(Lequel,ActionToString(Action)) ;
   //
   21 : {Appel fiche Pied Ports} AGLLanceFiche('GC','GCPORT','',Lequel,ActionToString(Action)+';MONOFICHE') ;   // Port et frais
{$IFDEF GRC}
   22..24 : RTMenuDispatchTT ( Num, Action, Lequel, TT, Range );
{$ENDIF}
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

   46 : Begin
      	AGLLanceFiche ('RT', 'RTSUSPECTS', '', Lequel, 'MONOFICHE;ACTION=MODIFICATION');
        end;
	 47 : begin
   			Arg5LanceFiche := Range + ';'+ActionToString(Action)+';MONOFICHE';
   			AGLLanceFiche('GC','GCTIERS','',lequel,Arg5LanceFiche) ;
        end;
	 50 :  {Appels du module interventions}
        Begin
        AGLLanceFiche('BTP','BTAPPELINT','','','CODEAPPEL='+ Lequel+ ';'+ ActionToString(Action));
        End;
   //**********************************
   //Gestion Parc/Materiel
   //**********************************
   60 : {Gestion Famille matériel/Parc}
        Begin
          AGLLanceFiche('BTP','BTFAMILLEMATERIEL','','','CODEFAMILLE='+ Lequel+ ';'+ ActionToString(Action));
        end;
   61 : {Gestion matériel/Parc}
        Begin
          AGLLanceFiche('BTP','BTMATERIEL','','','CODEMATERIEL='+ Lequel+ ';'+ ActionToString(Action));
        end;

(*
 	 997 : ParamTable(TT,taCreat,0,Nil) ;
   998 : ParamTable(TT,taModif,0,Nil) ;
   999 : ParamTable(TT,taCreat,0,Nil) ;
*)
   // mise en adequation avec gescom et affaire
   994 : ParamTable(TT,taModif,0,Nil,9) ;   {choixext}
   995 : ParamTable(TT,taCreat,0,Nil,9) ;   {choixext}
   996 : ParamTable(TT,taModif,0,Nil,6) ;   {choixext}
   997 : ParamTable(TT,taCreat,0,Nil,6) ;   {choixext}
   998 : ParamTable(TT,taModif,0,Nil,3) ;   {choixcode ou commun}
   999 : ParamTable(TT,taCreat,0,Nil,3) ;   {choixcode ou commun}
  end ;
END ;


procedure RTModifTabHie (Sender: TObject; var Caption: string) ;
// Tablettes hiérarchiques pour paramètrer le libellé
begin
  if not V_PGI.Sav then
    begin
      // on veut juste le libellé (pas de code entre [])
      if Sender is TYTreeLinkItem then
        Caption := TYTreeLinkItem(Sender).Libelle
      else if Sender is TYTreeDataType then
        Caption := TYTreeDataType(Sender).Libelle ;
    end;
end;

procedure TFMenuDisp.HDTDrawNode (Sender: TObject; var Caption: string) ;
begin
{$IFDEF GRC}
  RTModifTabHie (Sender,Caption);
{$ENDIF}
end;

{ MNG 05/03/2007 - initialisation et démarrage interface CTI }
procedure TFMenuDisp.InitializeCTI;
begin
   // Fenêtre d'alerte cti, même si le cti n'est pas démarré
   if (VH_RT.ctiAlerte = nil) and (VH_RT.CTISeria) then
      VH_RT.ctiAlerte := TFCtiAlerte.Create (nil);
end;

procedure TFMenuDisp.TerminateCTI;
begin
     // On supprime la fenêtre d'alerte CTI
     if Assigned (VH_RT.ctiAlerte) then
       FreeAndNil (VH_RT.ctiAlerte);
end;

procedure AfterChangeModule ( NumModule : integer ) ;
Var i       : Integer;
    Nbnew   : Integer;
    NbMod   : Integer;
    NbImp   : Integer;
    NbMes   : integer ;
//uniquement en line
    QQ			: TQuery;
    UneTOB	: TOB;
    OkReglCpta : boolean;
BEGIN

/// Menu Pop général
TraduitMenu(NumModule);

ChargeMenuPop(16,FMenuG.DispatchX);

{$IFDEF GRC}
  RTMenuTraiteMenusGRC(NumModule);
{$ENDIF}

  //uniquement en line
	if TobParamSoc <> nil then TobParamSoc.ClearDetail;

Case NumModule of

  145 : BEGIN
        if not GetParamSocSecur('SO_BTFACTUREACOMPTE',false) then
        begin
          FMenuG.RemoveGroup(-145480,True);  // facture d'acomptes
          FMenuG.RemoveItem  (-145480);
        end;
        if not GetParamSocSecur('SO_BTSITPROVISOIRE',false) then
        begin
          FMenuG.RemoveGroup(-145410,True);  // Situations provisoires
          FMenuG.RemoveItem  (-145410); 
        end;
        //---- WHYYYYYYYYYY
        FMenuG.RemoveItem  (145951); // Envoie Ecriture
        FMenuG.RemoveItem  (145952); // Récupération infos
        FMenuG.RemoveItem  (145953); // Ouverture Nouvel exercice
        FMenuG.RemoveItem  (145954); // Clôture d'exercice
        // ----
        OkReglCpta := GetParamSocSecur ('SO_BTGETREGLCPTA',false);
        if not OkReglCpta then
        begin
          FMenuG.removeGroup (-145950,true); // Echanges ou traitements comptables //if GetParamSocSecur('SO_BTLIENCPTAS1','')='AUC' then
          FMenuG.RemoveItem  (145955); // Récupéartion des reglements depuis la compta
        end;
        if not ExisteSql('SELECT MO_CODE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="ZEL" AND MO_CODE="ZDU"') then
         		FMenuG.removeItem (145491); // etat recap HLM DURAND SAS

        if SPIGAOOBJ.connected then
        begin
          SPIGAOOBJ.RecupNbAffaires (Nbnew, NbMod, NbImp, NbMes);
          FMenuG.RenameItem(145961,'Affaires Nouvelles ('+IntToStr(NbNew)+')');
          FMenuG.RenameItem(145962,'Affaires Modifiées ('+IntToStr(NbMod)+')');
          FMenuG.RenameItem(145963,'Affaires à Importer ('+IntToStr(NbImp)+')');
          FMenuG.RenameItem(145964,'Messages SPIGAO ('+IntToStr(NbMes)+')');
          if SPIGAOOBJ.IsNewGestion then
          begin
          	FMenuG.RenameItem (145967,'Configuration');
        		FMenuG.RemoveItem (145968); // Clôture d'exercice
          end;
        end;
          if (VH_GC.BTCODESPECIF <> '001') then FMenuG.RemoveItem(147163);
       END;
//
  146 : BEGIN

          if not (VH_GC.BTCODESPECIF = '001') then
          begin
            FMenuG.RemoveItem(146241);
            FMenuG.RemoveItem(146641);
          end;

          if not GetParamSocSecur('SO_CHANTIEROBLIGATOIRE',false) then
          begin
            FMenuG.RemoveItem (31211);
            FMenuG.RenameItem (31215,'Saisie - Modification');
          end;

        end;
  147 : BEGIN
          FmenuG.RemoveGroup (-147300, True); //On enlève préparation de gestion de chantier en suite
          if not V_PGI.Sav then
          begin
            FMenuG.removeItem (147610);
            FMenuG.removeItem (147620);
          end;
          if (VH_GC.BTCODESPECIF = '001') then FMenuG.RemoveItem(147130);
        end;
  149 : begin
         {Categories de dimensions }
         for i:=1 to 5 do FMenuG.RenameItem(149120+i,RechDom('GCCATEGORIEDIM','DI'+IntToStr(i),FALSE)) ;

  //uniquement en line
  {*
        //contrôle si le produit est sérialisé
        If (VH_GC.BTSeriaChantiers = False) and (V_PGI.VersionDemo = False) then
	         Begin
           FMenuG.removeItem (149290); 				// Frais
         	 FMenuG.removeGroup(149400,True); 	// Fournisseur
         	 FMenuG.removeGroup (-149900,True); // Ressources
           End;

        //Modification libellé menu affaire en Chantier
        FMenuG.RenameItem(147800,'Gestion des Chantiers');
        FMenuG.RenameItem(145110,'Chantiers');
       	FMenuG.removeItem (149245); // Variables Générales
       	FMenuG.removeItem (149260); // Nature de prestations
       	FMenuG.removeItem (-149270); // Tarif H.T
       	FMenuG.removeItem (-149280); // Tarif TTC
        //
       	FMenuG.removeGroup(149100, True); // Dimensionnement article
       	FMenuG.removeGroup (-147700,TRue); // Stock
        //
        //Modification Libellé Création-Modification en Client
        //
        FMenuG.RenameItem(30501,'Clients');
       	FMenuG.removeItem (30502); // Contact clients
       	FMenuG.removeItem (33105); // Cubes
        FMenuG.removeItem (30504); // Modification en serie Clients
       	FMenuG.removeItem (30505); // Mise en sommeil
       	FMenuG.removeItem (30506); // Activation
        //
       	FMenuG.removeGroup (149350,True); // prospects
        //
        //Modification Libellé Création-Modification en Fournisseurs
        //
        FMenuG.RenameItem(31701,'Fournisseurs');
        FMenuG.removeItem (31702); // Catalogue
        FMenuG.removeItem (138512); //Edition Catalogue
        FMenuG.removeItem (-31400); // Tarif
        FMenuG.removeItem (31703); // Contact
        FMenuG.removeItem (31704); // Modificatione en série Fournisseurs
        FMenuG.removeItem (31705); // Mise en Sommeil
        FMenuG.removeItem (31706); // Activation
        //
       	FMenuG.removeGroup (149500,True); // Commerciaux
        //
        FMenuG.removeItem (149910); // Saisie de Ressource
        FMenuG.removeItem (149920); // Synchronisation ressource
        //
       	FMenuG.removeGroup (-149700,True); // Bordereaux de prix
        FMenuG.RemoveGroup(-149290,True);  // Batiprix
        //
        FMenuG.OutLook.RemoveGroup(-149600, True);  //Paramètres
       	FMenuG.removeItem (149610); // Utilitaire de recodification article
        //
*}
        //Affichage saisies divers types ressources specif Line - 15092008
        {
        FMenuG.removeGroup (149930, True);   // Salaries
	      FMenuG.removeGroup (149940, True);   // Sous-Traitants
   		  FMenuG.removeGroup (149950, True);   // Intérims
		    FMenuG.removeGroup (149960, True);   // Matériels
    		FMenuG.removeGroup (149970, True);   // Outils
		    FMenuG.removeGroup (149980, True);   // Location
        }
        //
        FMenuG.removeItem (149930);   // Salaries
	      FMenuG.removeItem (149940);   // Sous-Traitants
   		  FMenuG.removeItem (149950);   // Intérims
		    FMenuG.removeItem (149960);   // Matériels
    		FMenuG.removeItem (149970);   // Outils
		    FMenuG.removeItem (149980);   // Location
		    FMenuG.removeItem (149630);   // Assistant de paramétrage société
//
        if (Not GetParamSocSecur('SO_GCFAMHIERARCHIQUE',False)) then
        begin
          FMenuG.removeItem (149250);   // Familles hiérarchiques Articles
          FMenuG.removeItem (149251);   // Familles hiérarchiques Ouvrages
        end;
        end ;

  150 : begin
         FMenuG.removeItem (32203); // Fin de mois
         FMenuG.removeItem (32204); // Fin d'exercice
         FMenuG.removeItem (32233); // réajustements préparé client
         FMenuG.removeItem (32405); // Consommations
         end;
   60:  BEGIN
        //uniquement en line
        {*
         FMenuG.RemoveGroup (60100,true) ;
         FMenuG.RemoveItem (60208) ;
         FMenuG.RemoveItem (-60210) ;
         FMenuG.removeItem (60209);
//         FMenuG.RemoveGroup (60800,true) ;
         FMenuG.removeItem (1205); // Personnalisation multiligne
         FMenuG.RemoveGroup (60400,true) ;
         FMenuG.RemoveItem (60402) ;
         FMenuG.RemoveItem (60608) ;
         FMenuG.RemoveGroup (60698,true) ;
         FMenuG.RemoveItem (-219040) ;
         FMenuG.RemoveItem (60611) ;
         FMenuG.RemoveGroup (60700,true) ;
         FMenuG.RemoveGroup (60500,True) ;
        *}

         FMenuG.RemoveGroup (-60417,true) ; // somptabilisation différée
         FMenuG.removeItem (60401); // comptabilisation différée
         FMenuG.removeItem (60404); // Assemblage
         FMenuG.removeItem (60413); // Comptanilisation des stocks
         FMenuG.removeItem (60612); // Conversion de nomenclature
         //
         FMenuG.RemoveGroup (-60410,true) ; // Historisation des pièces
         FMenuG.RemoveGroup (-60420,true) ; // épuration des pièces
         //
         FMenuG.RemoveGroup (60400,true) ; // épuration des pièces
         //
         FMenuG.RemoveGroup (-60613,true) ;
         FMenuG.RemoveGroup (60800,true) ; // multilingue
         FMenuG.RemoveGroup (-219030,true) ; // vérifications...
         FMenuG.RemoveGroup (-219046,true) ;
         FMenuG.removeItem (-219046); // Bascule tarif en mode Gescom
         FMenuG.removeItem (219050); // Bascule tarif en mode Gescom
         FMenuG.removeItem (60206); // restrictions
         FMenuG.removeItem (60207); // restrictions
   			END;
    71:  BEGIN
         if Not(VH_GC.AFGestionCom) then FMenuG.removeItem (71501);
         END;
    72:  BEGIN
         END;
    73:  BEGIN
         if Not(VH_GC.AFGestionCom) then FMenuG.removeItem (73306);
         END;
{$IFDEF NOMADE}
 	 101,102,103,105,106,110,111 : ChargeMenuPop(24,FMenuG.DispatchX) ;
{$ENDIF}
   148: BEGIN
   			ChangeLibelleMenus;
        if not (VH_GC.BTCODESPECIF = '001') then
        begin
          FMenuG.RemoveItem(-148820);
          FMenuG.RemoveItem(148821);
          FMenuG.RemoveItem(148822);
          FMenuG.RemoveItem(148823);
          FMenuG.RemoveGroup(-148820,true);
        end;
        For i:=1 to 9 do   // tablTiers
            begin
            FMEnuG.RenameItem (74410+i,RechDomZoneLibre ('CT'+IntToStr(i), False));
            FMEnuG.RenameItem (74469+i,RechDomZoneLibre ('ET'+IntToStr(i), False));
            FMEnuG.RenameItem (74440+i,RechDomZoneLibre ('AT'+IntToStr(i), False));
            FMEnuG.RenameItem (74460+i,RechDomZoneLibre ('MT'+IntToStr(i), False));
            FMEnuG.RenameItem (74430+i,RechDomZoneLibre ('RT'+IntToStr(i), False));
            FMEnuG.RenameItem (74540+i,RechDomZoneLibre ('BT'+IntToStr(i), False));
            end;
        For i:=1 to 3 do
            begin
            FMEnuG.RenameItem (74560+i,RechDomZoneLibre ('FT'+IntToStr(i), False));
            end;
        FMEnuG.RenameItem (74420,RechDomZoneLibre ('CTA', False));
        FMEnuG.RenameItem (74479,RechDomZoneLibre ('ETA', False));
        FMEnuG.RenameItem (74450,RechDomZoneLibre ('ATA', False));
        FMEnuG.RenameItem (74489,RechDomZoneLibre ('MTA', False));
        FMEnuG.RenameItem (74499,RechDomZoneLibre ('RTA', False));
        FMEnuG.RenameItem (74550,RechDomZoneLibre ('BTA', False));
        //uniquement en line
        {*
  			FMenuG.RemoveGroup(74100,True);  // parametres societes
  			FMenuG.RemoveGroup(-149200,True);  // Depots
  			FMenuG.RemoveGroup(1485,True);  // Contrats
  			FMenuG.RemoveGroup(74400,True);  // Tables libres
  			FMenuG.RemoveGroup(-74350,True);  // Appels d'offres
  			FMenuG.RemoveGroup(-148650,True);  // Commerciaux
        //
        FMenuG.removeItem (148100); // variables
        FMenuG.removeItem (148801); // Categorie de taxe
        FMenuG.removeItem (74206); // Ventilation analytiques
        FMenuG.removeItem (74208); // Frais et ports
        FMenuG.removeItem (74205); // Methodes d'arrondi
        FMenuG.removeItem (74204); // Multi-memos
        FMenuG.removeItem (148255); // Modeles Demandes d'acomptes
        FMenuG.removeItem (148254); // Modeles Etudes
        *}
        if (GetParamSocSecur('SO_GCFAMHIERARCHIQUE',False)) then
        begin
          // Tablettes Familles Articles
          FMenuG.removeItem (74341);
          FMenuG.removeItem (74342);
          FMenuG.removeItem (74343);
          // Tablettes Familles Ouvrages
          FMenuG.removeItem (148901);
          FMenuG.removeItem (148902);
          FMenuG.removeItem (148903);
        end;
        END ;
   283 : begin
           if ISPrepaLivFromAppro then
           begin
            FMenuG.removeItem (147310);
           end;
					 if not LivraisonVisible Then FMenuG.RemoveGroup (-145800,true) ;
   				 if not V_PGI.Sav then
           begin
           	FMenuG.removeItem (145870);
         		FMenuG.removeItem (147330);
           end;
         end;
	 284 : begin
         //verrue Révision de Prix
	       FMenuG.removegroup(284500, true);
   			 end;

//uniquement en line
   320 : Begin
         //Chargement de la TOB Paramètres
         QQ := OpenSql ('SELECT * FROM PARAMSOC',true,-1, '', True);
				 TOBParamSoc.LoadDetailDB ('PARAMSOC','','',QQ,false);
				 ferme (QQ);

         //Affichage des Ventilation comptable Article, tiers, affaire en fonction du paramètre
         UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTAART'],true);
         if UneTOB <> Nil then
            if UneTOB.GetValue('SOC_DATA')<>'X' then FMenuG.removeItem (320321);

         UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTATIERS'],true);
         if UneTOB <> Nil then
            if UneTOB.Getvalue('SOC_DATA')<> 'X' then FMenuG.removeItem (320322);

         UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTAAFF'],true);
         if UneTOB <> Nil then
            if UneTOB.Getvalue('SOC_DATA')<> 'X' then FMenuG.removeItem (320323);

			   //Permet de gérer en line le mode SAV pour affichage Administration PGI
   			 if Not V_PGI.Sav then
            Begin
            FmenuG.RemoveGroup(320400, True);
         		FMenuG.removeItem (320164);
						FMenuG.removeItem (60208); //Droits d'accès
            End;

         //Affichage des options au menu si pas de compta
			   UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
			   if UneTOB <> Nil then
				    if UneTOB.GetValue('SOC_DATA')='AUC' then
        			 begin
		           FmenuG.RemoveGroup(320300, True);
    		    	 FMenuG.removeItem (-320210);
               end;
         End;

   321 : Begin
         FmenuG.RemoveItem(321110);
         FmenuG.RemoveItem(321120);
         FmenuG.RemoveItem(321130);
         FmenuG.RemoveGroup(321200, True);
         FmenuG.RemoveItem(321210);
         FmenuG.RemoveItem(321220);
         FmenuG.RemoveItem(321230);
         FmenuG.RemoveItem(321240);       
   			 End;
    323: BEGIN
					ChangeLibelleMenusParc
    		 END;
    327: BEGIN
            if not (VH_GC.BTCODESPECIF = '001') then
            begin
              FMenuG.RemoveGroup(-327600,true);
            end else
            begin
              FMenuG.RemoveGroup(327200,true);
              FMenuG.RemoveGroup(327300,true);
              FMenuG.RemoveGroup(327400,true);
              FMenuG.RemoveGroup(327500,true);
            end;
         END;

End;
{$IFDEF BTP}
//FMenuG.removeItem (284202);
FMenuG.removeItem (284209);
FMenuG.removeItem (304520);
FMenuG.RemoveItem (320440);

//FMenuG.removeItem (284140);  // MODIF BRL 8/07/08 : en attendant les tests et la validation ...
//Chargement paramètre Facturation groupée (Spécifique BTP gestion des appels/Intervention)
if GetParamSoc('SO_BTFACAPPELDETAIL') then
   TypeFacDetail := 'APP'
Else
   TypeFacDetail := 'GRP';

//Module cti
if GetParamSoc('SO_RTCTIGESTION') = False then FMenuG.RemoveItem(304130) ;

if (GetParamSoc('SO_FACTPROV')=False) then
Begin
   FMenuG.removeItem (284203);
   FMenuG.removeItem (284204);
   FMenuG.RemoveItem (284302); //FV1 : 25/07/2013 ==> FS#298 - ESPACS : Ne pas faire apparaître "modification de factures" si pas de gestion factures provisoires
   FMenuG.removeItem (304530);
	 FMenuG.removeItem (304540);
   FmenuG.removeItem (304561); //FV1 : 25/07/2013 ==> FS#298 - ESPACS : Ne pas faire apparaître "modification de factures" si pas de gestion factures provisoires
end;
{$ENDIF}

{$IFDEF GRC}
if (GetParamSoc('SO_RTAFFICHEINFOS')=False) then
     FMenuG.RemovePopupItem(92296) ;
{$ENDIF GRC}
END;

procedure AppelTitreLibre (Name : string;PRien : THPanel; Prefixe : string );
begin
     // fct qui permet de ne saisir que la partie voulue de la tablette des titres libres
GCModifTitreZonesLibres (Name,Prefixe);
if Prefixe = 'ZLI' then ParamTable('GCZONELIBRE',taModif,0,PRien,3)
else if prefixe = 'ZLA' Then ParamTable ('GCZONELIBREART',taModif,0,PRien,3)
else if prefixe = 'ZLT' Then BEGIN ParamTable ('GCZONELIBRETIE',taModif,0,PRien,3); END
else if prefixe = 'ZLC' Then BEGIN ParamTable ('GCZONELIBRECON',taModif,0,PRien,3); END;
GCModifTitreZonesLibres ('',Prefixe);
if ctxScot in V_PGI.PGIContexte then AfterChangeModule(144)
                                else AfterChangeModule(148);
AvertirTable('GCZONELIBRE');                                
end;

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
   // cas ou on décoche "utilisation de la seria du programme' dans
   // l'écran de séria. On force la non sérialisation
V_PGI.nolock := true;
V_PGI.CADEnabled := false;
{$IFDEF DEMOBTP}
V_PGI.VersionDemo:=False;
VH_GC.BTSeriaAO:= TRUE;
VH_GC.BTSeriaES:= TRUE;
VH_GC.BTSeriaLogistique:= TRUE;
VH_GC.BTSeriaChantiers:= TRUE;
VH_GC.BTSeriaGRC:= TRUE;
VH_GC.BTSeriaContrat:=True;
VH_GC.BTSeriaIntervention:=True;
VH_GC.BTSeriaParcMateriel:= True;
VH_GC.SeriaNEGOCE:=True;
VH_GC.SeriaPlanCharge := true;
{$ELSE}
  If V_PGI.NoProtec then V_PGI.VersionDemo:=TRUE
                  else V_PGI.VersionDemo:=False;

  // verif si module sérialisé, sinon on passe en version démo
  //uniquement en line
  //  if (sAcces[1]='X') then   VH_GC.BTSeriaES:= TRUE
  //                     else   VH_GC.BTSeriaES:= FALSE;
  //  if (sAcces[2]='X') then   VH_GC.BTSeriaChantiers:= TRUE
  //                     else   VH_GC.BTSeriaChantiers:= FALSE;
  //
  if (sAcces[1]='X') then   VH_GC.BTSeriaES:= TRUE
                     else   VH_GC.BTSeriaES:= FALSE;
  if (sAcces[2]='X') then   VH_GC.BTSeriaAO:= TRUE
                     else   VH_GC.BTSeriaAO:= FALSE;
  if (sAcces[3]='X') then   VH_GC.BTSeriaLogistique:= TRUE
                     else   VH_GC.BTSeriaLogistique:= FALSE;
  if (sAcces[4]='X') then   VH_GC.BTSeriaChantiers:= TRUE
                     else   VH_GC.BTSeriaChantiers:= FALSE;
  if (sAcces[5]='X') then   VH_GC.BTSeriaGRC:= TRUE
                     else   VH_GC.BTSeriaGRC:= FALSE;
  if (sAcces[6]='X') then   VH_GC.BTSeriaContrat:= TRUE
                     else   VH_GC.BTSeriaContrat:= FALSE;
  if (sAcces[7]='X') then   VH_GC.BTSeriaIntervention:= TRUE
                     else   VH_GC.BTSeriaIntervention:= FALSE;
  if (sAcces[8]='X') then   VH_GC.SeriaCoTraitance:= TRUE
                     else   VH_GC.SeriaCoTraitance:= FALSE;
  if (sAcces[9]='X') then   VH_GC.SeriaSousTraitance:= TRUE
                     else   VH_GC.SeriaSousTraitance:= FALSE;
  if (sAcces[10]='X') then   VH_GC.SeriaNEGOCE:= TRUE
                      else   VH_GC.SeriaNEGOCE:= FALSE;
  if (sAcces[11]='X') then   VH_GC.SeriaPlanCharge:= TRUE
                      else   VH_GC.SeriaPlanCharge:= FALSE;
  if (sAcces[12]='X') then   VH_GC.BTSeriaParcMateriel:= TRUE
                      else   VH_GC.BTSeriaParcMateriel:= FALSE;

{$ENDIF}

END ;

procedure InitApplication ;
BEGIN

if Not IsLibrary then ;
{$IFDEF GRC}
  RTMenuInitApplication;
//  ProcCalcMul:=RTProcCalcMul ;
{$ENDIF GRC}

ProcZoomEdt:=AFZoomEdtEtat ;
ProcCalcEdt:=BTCalcOLEEtat ;
ProcCalcMul:=BTProcCalcMul ;
FMenuG.OnDispatch:=Dispatch ;
FMenuG.OnChargeMag:=ChargeMagHalley ;
//
//
//
//uniquement en line
// Création de la tob Paramètres
TOBParamSoc := TOB.create('PARAMSS1', nil, -1);
if NomHalley = 'CBTPS1' then
begin
	FMenuG.SetModules([145,325,327,283,146,150,147,92,284,304,323,149,160,320,148,60],[24,77,21,74,121,124,41,77,127,73,9,110,69,99,59,34,49]) ;
end else
begin
	FMenuG.SetModules([145,325,327,283,328,146,150,147,92,329,284,304,323,331,149,160,148,280,60],[24,77,21,74,72,121,124,41,77,99,127,73,9,110,45,69,99,34,78,49]) ;
end;
V_PGI.NbColModuleButtons:=2 ; V_PGI.NbRowModuleButtons:=9 ;

FMenuG.OnChangeModule:=AfterChangeModule ;

if not VersionInterne then
begin
  FMenuG.OnAfterProtec:=AfterProtec ;
  FMenuG.SetSeria(GCCodeDomaine,GCCodesSeria,GCTitresSeria) ;
end else
begin
	V_PGI.NoProtec := true;
	V_PGI.nolock := true;
	V_PGI.CADEnabled := false;
  VH_GC.BTSeriaAO:= TRUE;
  VH_GC.BTSeriaES:= true;
  VH_GC.BTSeriaLogistique:= TRUE;
  VH_GC.BTSeriaChantiers:= TRUE;
  VH_GC.BTSeriaGRC:= TRUE;
  VH_GC.BTSeriaContrat:=True;
  VH_GC.BTSeriaIntervention:=True;
  VH_GC.SeriaCoTraitance:=True;
  VH_GC.SeriaSousTraitance:=True;
  VH_GC.SeriaNEGOCE := True;
  VH_GC.BTSeriaParcMateriel := True;
  VH_GC.SeriaPlanCharge := True;
	V_PGI.VersionDemo:=False;
end;
FMenuG.OnChargeFavoris := ChargeFavoris;
FmenuG.bCAD.OnClick := nil;

//
FMenuG.SetPreferences(['Documents'],False) ;
V_PGI.DispatchTT:=DispatchTT ;
InitDicoAffaire ;

SPIGAOOBJ := T_SPIGAOCONNECT.create;
SPIGAOOBJ.Connecte;

END;

Procedure FermeApplication ;
begin
  SPIGAOOBJ.Free;
  FreeAndNil(TOBParamSoc);

  //uniquement en line
  if NomHalley = 'CBTPS1' then
  begin
    //Gestion de la sauvegarde automatique !!!!!
    if (PGIAsk(TraduireMemoire('Voulez-vous lancer la sauvegarde automatique de ' + V_PGI.NomSociete + ' ?'), 'Sauvegarde') = mrYes) then
    begin
      BackupPGIS1;
    end;
  end;

end;

                             
initialization
Apalatys:='LSE';
V_PGI.PGIContexte:=[ctxGescom,ctxAffaire,ctxBTP,ctxGRC];
V_PGI.StandardSurDp := False;  // indispensable pour l'aiguillage entre le DP et les bases dossiers (gamme Expert), voir SQL025
Copyright:='© Copyright ' + Apalatys ;
V_PGI.OutLook:=TRUE ;
V_PGI.OfficeMsg:=TRUE ;
V_PGI.ToolsBarRight:=TRUE ;
ChargeXuelib ;
V_PGI.VersionDemo:=True ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionReseau:=True ;
V_PGI.NumVersion:='10.0.0' ;
V_PGI.NumVersionBase:=998 ;
V_PGI.NumBuild:='000.155';
V_PGI.CodeProduit:='034' ;
V_PGI.DateVersion:=EncodeDate(2018,05,04);
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.BlockMAJStruct:=TRUE ;
V_PGI.PortailWeb:='http://www.lse.fr/' ;
V_PGI.CegidApalatys:=FALSE ;
V_PGI.QRMultiThread:=TRUE;


//V_PGI.NiveauAcces:=1 ;
V_PGI.MenuCourant:=0 ;
V_PGI.RazForme:=TRUE;
V_PGI.NumMenuPop:=27 ;
V_PGI.CegidBureau:= True;
V_PGI.LookUpLocate:= True;   // utile pour les ellipsis_click sur THEDIT
V_PGI.QRPDFOptions:=4;       // pour tronquer les libellés en impression
V_PGI.QRPdf:=True;           // Mode PDF Par défaut pour totalisations situations
V_PGI.MultiUserLogin := True;
//V_PGI.ToolBarreUp := true;

end.



