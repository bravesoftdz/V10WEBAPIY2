unit eMDispS5 ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
	HMsgBox, Hctrls, UIUtil, ImgList, ParamSoc, Ent1,
  HEnt1,IniFiles, Prefs,Buttons,About, HPanel, UtilPGI,
  Maineagl,MenuOLX,M3FP,UtileAGL,eTablette,AGLInit,
	CPGRANDLIVRETL_TOF, CPBALANCETL_TOF, 							// Etats sur TL
	TofBalagee,																				// Etats suivi des tiers
	CPJALCENTRAL_TOF,																	// Journal centralisateur
	CPBALGEN_TOF																			// Balance géné et Auxi
  ;



Procedure InitApplication ;
procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);


implementation

// PGI Communs

// Uses ;


procedure UpdateSeries ;
Var ii : integer ;
    St : String ;
BEGIN
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   {Axes}
   St:='TTAXE' ; ii:=TTToNum(St) ;
   if ii>0 then
      BEGIN
// JLD5Axes       if EstSerie(S5) then V_PGI.DECombos[ii].Where:='X_AXE<"A3" AND X_FERME="-"' else
         if EstSerie(S3) then V_PGI.DECombos[ii].Where:='X_AXE<"A2" AND X_FERME="-"' ;
      END ;
   {Qualifiants pièce}
   St:='TTQUALPIECECRIT' ; ii:=TTToNum(St) ;
   if ii>0 then
      BEGIN
      if EstSerie(S5) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"PRE" AND CO_CODE<>"TOU"' else
       if EstSerie(S3) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"PRE" AND CO_CODE<>"NSS" AND CO_CODE<>"SSI"' ;
      END ;
   St:='TTQUALPIECE' ; ii:=TTToNum(St) ;
   if ii>0 then
      BEGIN
      if EstSerie(S5) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"P" AND CO_CODE<>"R"' else
       if EstSerie(S3) then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"P" AND CO_CODE<>"R" AND CO_CODE<>"U"' ;
      END ;
  {$IFDEF CCS3}
  (* GP le 15/03/2002
   {Natures de journaux}
   St:='TTNATJAL' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"ANA" AND CO_CODE<>"ODA"' ;
   *)
   {Natures de journaux}
   St:='TTNATJAL' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<>"ANA"' ;
   {Types de rubrique}
   St:='TTRUBTYPE' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+
      ' AND CO_CODE<>"A/B" AND CO_CODE<>"A/G" AND CO_CODE<>"B/A" AND CO_CODE<>"BUD" AND CO_CODE<>"G/A" AND CO_CODE<>"G/T" AND CO_CODE<>"T/G"' ;
   {Niveaux de relance}
   St:='TTNIVEAURELANCE' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND CO_CODE<="003"' ;
  (* GP le 15/03/2002
   {Journaux}
   St:='TTJOURNAL' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA"' ;
   St:='TTJOURNAUX' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:='J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA"' ;
   *)
   {Journaux}
   St:='TTJOURNAL' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=V_PGI.DECombos[ii].Where+' AND J_NATUREJAL<>"ANA"' ;
   St:='TTJOURNAUX' ; ii:=TTToNum(St) ;
   if ii>0 then V_PGI.DECombos[ii].Where:=' AND J_NATUREJAL<>"ANA"' ;
  {$ENDIF}
   END ;
END ;

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
V_PGI.NbColModuleButtons:=1 ; V_PGI.NbRowModuleButtons:=3 ;
END ;

procedure AssignLesZoom ;
BEGIN
ProcZoomEdt:=Nil ;
{
   ProcZoomEdt:=ZoomEdtEtat ;}
//   ProcCalcEdt:=CalcOLEEtat ;

END ;

procedure AfterChangeModule ( NumModule : integer ) ;
BEGIN
AssignLesZoom ;
UpdateSeries;
END ;

procedure AssignZoom ;
BEGIN
AssignLesZoom ;
END ;

{function ChargeFavoris : boolean ;
var lesTagsToRemove : string;
begin
  lesTagsToRemove := '';
  AddGroupFavoris( FMenuG , lesTagsToRemove) ;
  result := true;
end;}

procedure InitApplication ;
BEGIN
	AssignZoom ;
  //-------------
    FMenuG.OnDispatch:=Dispatch ;
    FMenuG.OnChargeMag:=ChargeMagHalley ;
    FMenuG.OnMajAvant:=Nil ;
    FMenuG.OnMajApres:=Nil ;
    FMenuG.OnChangeModule:=AfterChangeModule ;
    V_PGI.DispatchTT:=DispatchTT ;
    FMenuG.SetModules([8,21],[11,49]) ;
    V_PGI.NbColModuleButtons := 1 ;
    V_PGI.NbRowModuleButtons := 2 ;

//    FMenuG.SetPreferences(['Pièces'],False) ;
//    FMenuG.OnChargeFavoris := ChargeFavoris;
		FMenuG.OnAfterProtec:=AfterProtec ;
END ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
end ;

Procedure Dispatch ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
BEGIN
  Case Num of
    10 : begin
           ChargeHalleyUser ;
        end;       //  après connexion
    11 : ; //  après déconnexion
    12 : ; //  avant connexion et séria
    13 : ChargeMenuPop(TTypeMenu(-1),FMenuG.DispatchX) ;
    15 : ; //  ??????
    100 : ;  // Si lanceur
  //==== Analyses ====
    // Cubes
   8305  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','') ;
    8307  : AGLLanceFiche('CP','CPECRITURE_CUBE','','','CONTREPARTIE') ;
    8310  : AGLLanceFiche('CP','CPANALYTIQ_CUBE','','','') ;
    8312  : AGLLanceFiche('CP','CPBALAGEE_CUBE','','','') ;
    // Analyses statistiques
    8405  : AGLLanceFiche('CP','CPECRITURE_TOBV','','','') ;
    8410  : AGLLanceFiche('CP','CPANALYTIQ_TOBV','','','') ;
    // Analytiques multi-axes
    8205  : AGLLanceFiche('CP','CPGENERECROISEAXE','','','') ;
    8210  : AGLLanceFiche('CP','CPCROISEAXE_CUBE','','','') ;
    8215  : AGLLanceFiche('CP','CPCROISEAXE_STAT','','','') ;
    8220  : AGLLanceFiche('CP','CPCROISEAXE_ETAT','','','') ;

 // ==== Editions ====
    // --> Journaux
    7394 : AGLLanceFiche('CP', 'EPJALDIV', '', 'JDP', 'JDI') ; 	// Journal divisionnaire
    7397 : CPLanceFiche_JournalCentral; 												// Journal centralisateur eAGL
    7403 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JPP', 'JPE') ; 	// Journal périodique par Journal-Période TOFEDJAL
    7406 : AGLLanceFiche('CP', 'EPJALGEN', '', 'JGE', 'JAL') ; 	// Journal général
    // --> Grand livre
    7415 : AGLLanceFiche('CP', 'EPGLIVRE', '', 'GLG', 'GLG') ; 	// Grand livre général
    7418 : AGLLanceFiche('CP', 'EPGLIVRE', '', 'GLA', 'GLA') ; 	// Grand livre auxiliaire
    7419 : CPLanceFiche_EtatGrandLivreSurTables; 								// Grand Livre sur table libre
    // --> Balance
    7445 : CPLanceFiche_Balance('GENE'); 												// Balance générale eAGL
    7448 : CPLanceFiche_Balance('AUXI'); 												// Balance auxiliaire eAGL
    7429 : CPLanceFiche_EtatBalanceSurTables; 									// Balance sur table libre
    // --> Etats de synthèses
    7440 : AGLLanceFiche('CP', 'EPETASYN', '', 'RCP', 'RES') ; // Document de synthèse - Résultat
    7441 : AGLLanceFiche('CP', 'EPETASYN', '', 'BAF', 'BIL') ; // Document de synthèse - Actif
    7442 : AGLLanceFiche('CP', 'EPETASYN', '', 'SIG', 'SIG') ; // Document de synthèse - SIG
    // --> Suivi des tiers
    7547 : CPLanceFiche_BalanceAgee;   			// Balance âgée eAGL
    7556 : CPLanceFiche_BalanceVentilee;   	// Balance Ventilée eAGL

      else HShowMessage('2;?caption?;'+TraduireMemoire('Fonction non disponible : ')+';W;O;O;O;',TitreHalley,IntToStr(Num)) ;
     end ;
END ;

initialization

{Version}
Apalatys:='CEGID' ;
NomHalley:='Comptabilité S5' ;
TitreHalley:='CEGID Comptabilité S5' ;
HalSocIni:='CEGIDPGI.INI' ;
Copyright:='© Copyright ' + Apalatys ;
V_PGI.NumVersion:='4.00' ;
V_PGI.NumBuild:='1' ; V_PGI.NumVersionBase:=591 ;
V_PGI.DateVersion:=EncodeDate(2001,12,13) ;

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
V_PGI.Decla100:=TRUE ;
V_PGI.RAZForme:=TRUE ;

end.


