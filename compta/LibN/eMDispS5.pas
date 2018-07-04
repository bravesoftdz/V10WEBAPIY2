unit eMDispS5 ;

interface

uses
    Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
    HMsgBox, Hctrls, UIUtil, ImgList, ParamSoc, Ent1,
    HEnt1,IniFiles, Prefs,Buttons,About, HPanel, UtilPGI,
    Maineagl,MenuOLX,M3FP,UtileAGL,eTablette,AGLInit,
    CPGRANDLIVRETL_TOF, CPBALANCETL_TOF,                                // Etats sur TL
    TofBalagee,                                                         // Etats suivi des tiers
    CPJALCENTRAL_TOF,                                                   // Journal centralisateur
    CPBALGEN_TOF, CPBALAUXI_TOF,                                        // Balance géné et Auxi
    eSaisie,
    UTOFCPMULMVT,
    Hout,
    UTOFConsEcr,
    ComCtrls,                                                           // Pour le removeitem
    EntPGI,
    UTOMUTILISAT,                                                       // pour la fiche utilisateur
    USERGRP_TOM                                                         // pour les groupe d'utilisateur
    ;

Procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
procedure TripotageMenu (NumModule : integer);


implementation

// PGI Communs

// Uses ;

procedure UpdateSeries ;
var
    ii : integer ;
    St : String ;
begin
    if ((EstSerie(S3)) or (EstSerie(S5))) then
    begin
        {Axes}
        St:='TTAXE' ; ii:=TTToNum(St) ;
        if ii>0 then
        begin
            if EstSerie(S3) then V_PGI.DECombos[ii].Where:='X_AXE<"A2" AND X_FERME="-"' ;
        end;

        {Qualifiants pièce}
        St:='TTQUALPIECECRIT' ; ii:=TTToNum(St) ;
        if ii>0 then
        begin
            if EstSerie(S5) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"PRE" AND CO_CODE<>"TOU"' else
            if EstSerie(S3) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"PRE" AND CO_CODE<>"NSS" AND CO_CODE<>"SSI"' ;
        end;

        St:='TTQUALPIECE' ; ii:=TTToNum(St) ;
        if ii>0 then
        begin
            if EstSerie(S5) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"P" AND CO_CODE<>"R"' else
            if EstSerie(S3) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"P" AND CO_CODE<>"R" AND CO_CODE<>"U"' ;
        end;

        {$IFDEF CCS3}

        {Natures de journaux}
        St:='TTNATJAL' ; ii:=TTToNum(St) ;
        if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"ANA"' ;

        {Types de rubrique}
        St:='TTRUBTYPE' ; ii:=TTToNum(St) ;
        if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+ ' AND CO_CODE<>"A/B" AND CO_CODE<>"A/G" AND CO_CODE<>"B/A" AND CO_CODE<>"BUD" AND CO_CODE<>"G/A" AND CO_CODE<>"G/T" AND CO_CODE<>"T/G"' ;

        {Niveaux de relance}
        St:='TTNIVEAURELANCE' ; ii:=TTToNum(St) ;
        if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<="003"' ;

        {Journaux}
        St:='TTJOURNAL' ; ii:=TTToNum(St) ;
        if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND J_NATUREJAL<>"ANA"' ;
        St:='TTJOURNAUX' ; ii:=TTToNum(St) ;
        if ii>0 then V_PGI.DECombos[ii].Where:=' AND J_NATUREJAL<>"ANA"' ;

        {$ENDIF}
    end;
end;

Procedure AfterProtec ( sAcces : String ) ;
begin
    V_PGI.NbColModuleButtons:=2;
    V_PGI.NbRowModuleButtons:=5;
    FMenuG.SetModules([9,11,13,16,14,12,20,8,17,18],[108,78,81,90,58,22,1,11,28,49]);
    V_Pgi.NoModuleButtons:=false;
end;

procedure AssignLesZoom ;
begin
    ProcZoomEdt:=Nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 15/11/2002
Modifié le ... :   /  /
Description .. : -15/11/2002 - on n'appelle plus tripotage menu ( on affiche
Suite ........ : tous les menus )
Mots clefs ... :
*****************************************************************}
procedure AfterChangeModule ( NumModule : integer ) ;
begin
    AssignLesZoom ;
    UpdateSeries;
    // On cache les menus non gérés en eAGL
    //  TripotageMenu (NumModule);
end;

procedure AssignZoom ;
begin
    AssignLesZoom ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/11/2002
Modifié le ... :   /  /
Description .. : -13/11/2002- deplacement de l'init des menu ds afterprotec
Mots clefs ... :
*****************************************************************}
procedure InitApplication ;
begin
    AssignZoom ;
    //-------------
    FMenuG.OnDispatch:=Dispatch ;
    FMenuG.OnChargeMag:=ChargeMagHalley ;
    FMenuG.OnMajAvant:=Nil ;
    FMenuG.OnMajApres:=Nil ;
    FMenuG.OnChangeModule:=AfterChangeModule ;
    V_PGI.DispatchTT:=DispatchTT ;
    FMenuG.OnAfterProtec:=AfterProtec ;
end;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
end;

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
    prout : integer;
begin
    Case Num of
        10 : ChargeHalleyUser ;                                     //  après connexion
        11 : ;                                                      //  après déconnexion
        12 : ;                                                      //  avant connexion et séria
        13 : ChargeMenuPop(-1,FMenuG.DispatchX) ;
        15 : ;                                                      //  ??????
        100 : ;                                                     // Si lanceur

    //==== Ecritures (9)  ====
        // Courantes cwas
        72440 : eLanceSaisieDirecte;
        7244 : MultiCritereMvt(taCreat,'N',False) ;                 // Creation
        7256 : MultiCritereMvt(taConsult,'N',False) ;               // Consultation
        7259 : MultiCritereMvt(taModif,'N',False) ;                 // Modification

    //==== Tiers (11) ====
        // --> Suivi des tiers
        75470 : CPLanceFiche_BalanceAgee;                           // Balance âgée eAGL
        75560 : CPLanceFiche_BalanceVentilee;                       // Balance Ventilée eAGL

    //==== Editions (13) ====
        // --> Journaux
        7394 : AGLLanceFiche('CP','EPJALDIV','','JDP','JDI') ;      // Journal divisionnaire
        7397 : CPLanceFiche_JournalCentral;                         // Journal centralisateur eAGL
        7403 : AGLLanceFiche('CP','EPJALGEN','','JPP','JPE') ;      // Journal périodique par Journal-Période TOFEDJAL
        7406 : AGLLanceFiche('CP','EPJALGEN','','JGE','JAL') ;      // Journal général
        // --> Grand livre
        7415 : AGLLanceFiche('CP','EPGLIVRE','','GLG','GLG') ;      // Grand livre général
        7418 : AGLLanceFiche('CP','EPGLIVRE','','GLA','GLA') ;      // Grand livre auxiliaire
        7419 : CPLanceFiche_EtatGrandLivreSurTables;                // Grand Livre sur table libre
        // --> Balance
        7445 : CPLanceFiche_BalanceGeneral;                         // Balance générale eAGL
        7448 : CPLanceFiche_BalanceAuxiliaire;                      // Balance auxiliaire eAGL
        7429 : CPLanceFiche_EtatBalanceSurTables;                   // Balance sur table libre
        // --> Etats de synthèses
        7440 : AGLLanceFiche('CP','EPETASYN','','RCP','RES') ;      // Document de synthèse - Résultat
        7441 : AGLLanceFiche('CP','EPETASYN','','BAF','BIL') ;      // Document de synthèse - Actif
        7442 : AGLLanceFiche('CP','EPETASYN','','SIG','SIG') ;      // Document de synthèse - SIG

    //==== Traitements (16) ====
        7602 : CPLanceFiche_CONSECR('');                            // Consultation des comptes

    //==== Analyses (8) ====
        // Cubes
        8305 : AGLLanceFiche('CP','CPECRITURE_CUBE','','','') ;
        8307 : AGLLanceFiche('CP','CPECRITURE_CUBE','','','CONTREPARTIE') ;
        8310 : AGLLanceFiche('CP','CPANALYTIQ_CUBE','','','') ;
        8312  : AGLLanceFiche('CP','CPBALAGEE_CUBE','','','') ;
        // Analyses statistiques
        8405 : AGLLanceFiche('CP','CPECRITURE_TOBV','','','') ;
        8410 : AGLLanceFiche('CP','CPANALYTIQ_TOBV','','','') ;
        8412 : AGLLanceFiche('CP','CPANAGENE_TOBV','','','') ;
        // Analytiques multi-axes
        8205 : AGLLanceFiche('CP','CPGENERECROISEAXE','','','') ;
        8210 : AGLLanceFiche('CP','CPCROISEAXE_CUBE','','','') ;
        8215 : AGLLanceFiche('CP','CPCROISEAXE_STAT','','','') ;
        8220 : AGLLanceFiche('CP','CPCROISEAXE_ETAT','','','') ;

    //==== Administration / outils (18) ====
        // Utilisateur
        3165 : FicheUserGrp;
        3170 : begin FicheUSer(V_PGI.User); ControleUsers; end;

    //==== Non encore implementé ====
        else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
    end;
end;


procedure RemoveGroupTN(i : Integer ; b : Boolean) ;
begin
    FMenuG.OutLook.RemoveGroup(i,b) ;
    //  TN:=GetTreeItem(FMenuG.TreeView,i) ;
    //  If TN<>NIL Then TN.Delete ;
end;

Procedure RemoveItemTN(i : Integer) ;
var
    HOI : THOutItem ;
begin
    HOI:=FMenuG.OutLook.GetItemByTag(i) ;
    If HOI<>NIL Then FMenuG.OutLook.RemoveItem(HOI) ;
    //  TN:=GetTreeItem(FMenuG.TreeView,i) ;
    //  If TN<>NIL Then TN.Delete ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 13/11/2002
Modifié le ... :   /  /
Description .. : -13/11/2002- ajout des menu de lettrage
Mots clefs ... :
*****************************************************************}
procedure TripotageMenu (NumModule : integer);
begin
    // ===> Ecritures <===
    if NumModule = 9 then
    begin
        // Courantes
        RemoveItemTN(7245) ;        // Effets
        RemoveItemTN(7247) ;        // Règlements
        RemoveItemTN(7241) ;        // Guidée
        RemoveItemTN(7005) ;        // Bordereau
        RemoveItemTN(7264) ;        // Entête de pièce
        RemoveItemTN(-82) ;         // Suppression       ??
        RemoveItemTN(7268) ;        // Brouillard
        // Autres groupes
        RemoveGroupTN(7271,True) ;  // Simulations
        RemoveGroupTN(7292,True) ;  // situations
        RemoveGroupTN(7325,True) ;  // Abonnements
        RemoveGroupTN(7361,True) ;  // Analytiques
    end
    else
    // ===> Tiers <===
    if NumModule = 11 then
    begin
        // Editions de suivi
        // RemoveItemTN(7245) ;                 //

        // Autres groupes
        RemoveGroupTN(7496,True);   // Encaissements
        RemoveGroupTN(-74999,True); // Décaissements
        RemoveGroupTN(7562,True);   // Relances de suivi
        RemoveGroupTN(-7230,True);  // Tiers Payeurs
        RemoveGroupTN(-7589,True);  // Emissions
        RemoveGroupTN(7532,True);   // edition de suivi

        //module lettrage   7505
        RemoveItemTN(7511) ;        // delettrage
        RemoveItemTN(7514) ;        // lettrage auto
        RemoveItemTN(7520) ;        // regul de lettrage
        RemoveItemTN(7523) ;        // diff de change
        RemoveItemTN(7524) ;        // ecart de conversion
        RemoveItemTN(7529) ;        // modif echeance
    end
    else
    // ===> Editions <===
    if NumModule = 13 then
    begin
        // Journaux
        RemoveItemTN(7409) ;        // Périodique par général

        // Grands livres
        RemoveItemTN(7421) ;        // Analytique
        RemoveItemTN(7427) ;        // Généraux par Auxiliaires
        RemoveItemTN(7430) ;        // Auxiliaires par Généraux
        RemoveItemTN(7436) ;        // Généraux par Analytiques
        RemoveItemTN(7439) ;        // Analytiques par Généraux

        // Balances
        RemoveItemTN(7451) ;        // Analytique
        RemoveItemTN(7457) ;        // Généraux par Auxiliaires
        RemoveItemTN(7460) ;        // Auxiliaires par Généraux
        RemoveItemTN(7466) ;        // Généraux par Analytiques
        RemoveItemTN(7469) ;        // Analytiques par Généraux

        // Autres groupes
        RemoveGroupTN(7475,True) ;  // Cumuls périodiques
        RemoveGroupTN(0,True) ;     // Etats de synthèses
        RemoveGroupTN(-7410,True) ; // Editions paramétrables
        RemoveGroupTN(-7456,True) ; // Etats Libres
    end
    else
    // ===> Traitements <===
    if NumModule = 16 then
    begin
        // Courantes
        RemoveItemTN(7245) ;        // Effets
        // Autres groupes
        RemoveGroupTN(7271,True);   // Simulations
        RemoveGroupTN(-96110,True); // ICC
        RemoveGroupTN(-7772,True);  // Releves
        RemoveGroupTN(-7710,True);  // Balances
        RemoveGroupTN(-7728,True);  // A Nouveaux
        RemoveGroupTN(-7691,True);  // Validation
        // suppression des modules sans tag
        RemoveItemTN(7604);         // validation
        RemoveItemTN(7607);         // depointage
        RemoveItemTN(7616);         // etat de pointage
        RemoveItemTN(7619);         // etat de rappro
        RemoveItemTN(7622);         //  justif de solde
        RemoveItemTN(7631);         //  controle de caisse
        RemoveItemTN(7736);
        RemoveItemTN(7737);
        RemoveItemTN(7742);
        RemoveItemTN(7745);
        RemoveItemTN(7751);
        RemoveItemTN(7757);
        RemoveItemTN(7760);
        RemoveItemTN(7766);
        RemoveItemTN(7766);
        RemoveItemTN(7769);
        RemoveItemTN(7683);
        RemoveItemTN(7717);
        RemoveItemTN(7717);
        RemoveItemTN(7261);         // modif du journal centralisateur
        RemoveItemTN(7677);         // extourne
        RemoveItemTN(7719);         // affection sur table libre
        RemoveItemTN(7724);
        RemoveItemTN(7727);
        RemoveItemTN(7718);
    end
    else
    // ===> Analyses <===
    if NumModule = 8 then
    begin
    end;
end;


{$IFDEF AGL550B}
Procedure InitLaVariablePGI;
Begin
    {Version}
    Apalatys:='CEGID' ;
    NomHalley:='Comptabilité S5' ;
    TitreHalley:='CEGID Comptabilité S5' ;
    HalSocIni:='CEGIDPGI.INI' ;
    Copyright:='© Copyright ' + Apalatys ;
    V_PGI.NumVersion:='4.2.0' ;
    V_PGI.NumBuild:='2' ; V_PGI.NumVersionBase:=595 ;
    V_PGI.DateVersion:=EncodeDate(2003,01,13) ;

    {Généralités}
    V_PGI.VersionDemo:=True ;
    V_PGI.SAV:=False ;
    V_PGI.VersionReseau:=True ;
    V_PGI.PGIContexte:=[ctxCompta] ;
    V_PGI.CegidAPalatys:=FALSE ;
    V_PGI.CegidBureau:=TRUE ;
    V_PGI.StandardSurDP:=True ;
    V_PGI.MajPredefini:=False ;
    V_PGI.MultiUserLogin:=False ;
    V_PGI.BlockMAJStruct:=True ;
    V_PGI.EuroCertifiee:=True ;

    ChargeXuelib ;

    {Série}
    V_PGI.LaSerie:=S5 ;
    V_PGI.NumMenuPop:=27 ;
    V_PGI.OfficeMsg:=True ;
    V_PGI.OutLook:=TRUE ;
    V_PGI.NoModuleButtons:=False ;
    V_PGI.NbColModuleButtons:=1 ;

    {Divers}
    V_PGI.MenuCourant:=0 ;
    V_PGI.OKOuvert:=FALSE ;
    V_PGI.Halley:=TRUE ;
    V_PGI.VersionDEV:=TRUE ;
    V_PGI.ImpMatrix := True ;
    V_PGI.DispatchTT:=DispatchTT ;
    V_PGI.ParamSocLast:=False ;
    V_PGI.RAZForme:=TRUE ;
End;

Initialization

ProcChargeV_PGI :=  InitLaVariablePGI ;
{$ELSE}

initialization

{Version}
Apalatys:='CEGID' ;
NomHalley:='Comptabilité S5' ;
TitreHalley:='CEGID Comptabilité S5' ;
HalSocIni:='CEGIDPGI.INI' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.NumVersion:='4.2.0' ;
V_PGI.NumBuild:='2' ; V_PGI.NumVersionBase:=595 ;
V_PGI.DateVersion:=EncodeDate(2003,01,13) ;

{Généralités}
V_PGI.VersionDemo:=True ;
V_PGI.SAV:=False ;
V_PGI.VersionReseau:=True ;
V_PGI.PGIContexte:=[ctxCompta] ;
V_PGI.CegidAPalatys:=FALSE ;
V_PGI.CegidBureau:=TRUE ;
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=False ;
V_PGI.MultiUserLogin:=False ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=True ;

ChargeXuelib ;

{Série}
V_PGI.LaSerie:=S5 ;
V_PGI.NumMenuPop:=27 ;
V_PGI.OfficeMsg:=True ;
V_PGI.OutLook:=TRUE ;
V_PGI.NoModuleButtons:=False ;
V_PGI.NbColModuleButtons:=1 ;

{Divers}
V_PGI.MenuCourant:=0 ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.VersionDEV:=TRUE ;
V_PGI.ImpMatrix := True ;
V_PGI.DispatchTT:=DispatchTT ;
V_PGI.ParamSocLast:=False ;
V_PGI.RAZForme:=TRUE ;

{$ENDIF}

end.


