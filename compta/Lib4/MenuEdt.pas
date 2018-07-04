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
unit MenuEdt;

interface
Uses
  Windows, Messages,
  HEnt1,EntPGI,
{$ifdef eAGLClient}
  MenuOLX,MaineAGL,eTablette, UtilEagl,
{$else}
  MenuOLG, FE_Main,
{$endif eAGLClient}

  Forms,sysutils,HMsgBox, Classes, Controls, HPanel,
  hctrls, MajTable,
  ExtCtrls, inifiles, UTOB,

  Reseau, ImgList
 ;

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1  : TImageList;
    end ;

Var FMenuDisp : TFMenuDisp ;
    Soc,User,Motdepass    : string;
    OkExport,Minimise : Boolean;
    OkSic : Boolean ;

procedure RemoveFromMenu ( const iTag : integer; const bGroup : boolean; var stAExclure : string );
procedure AfterChangeModule ( NumModule : integer ) ;

implementation

{$R *.DFM}

Uses LicUtil,Ent1,
     ParamSoc, QRGLGen, QRBLGESE, QRBLSEGE, QRGLANA, QRGLAuGe, QRGLAux,
     QRGLGeAu, QRGLGESE, QRGLSEGE, QRBaAuGe, QRBaGeAu, QRBalAux, QRBalGen,
     QRBlAna, QRJCpte, QRJDivis, QRCumulG, CRITEDT, QRTvaAcc, QRTvaAcE,
     QRTvaFac, uTofCPGLGene, QRAff, QRBrouil, JustiSol, QRGLAge, QRGLVen,
     QrBalAge, QrBalVen, QRBAGTP, QRBAVTP, QRGLATP, QRGLVTP, QrBroAna, QrBrt,
     utilPGI, UserGrp_tom, UTOMUTILISAT, Confidentialite_TOF, MulLog1,
     TomProfilUser;

Procedure RenseignelaSerieGG ;

begin
HalSocIni:='CEGIDPGI.INI' ;
If OkSic Then HalSocIni:='CEGIDPGISIC.INI' ;

V_PGI.NumVersionBase:=0 ;
V_PGI.NumVersion:='1.0.0' ;
V_PGI.NumBuild:='002.001' ;
V_PGI.DateVersion:=EncodeDate(2007,02,16) ;


{$IFDEF EAGLCLIENT}
NomHalley:='EDT23' ;
TitreHalley:='Editions comptables' ;
{$ELSE}
NomHalley:='EDT23' ;
TitreHalley:='Editions comptables' ;
{$ENDIF}
If OkSIC Then
 BEGIN
{$IFDEF EAGLCLIENT}
NomHalley:='EDT23' ;
TitreHalley:='Editions comptables' ;
{$ELSE}
NomHalley:='EDT23' ;
TitreHalley:='Editions comptables' ;
{$ENDIF}
 END ;
Application.Title := TitreHalley ;

END ;

//exemple rem c:\pgi00\app\COMSX.EXE /TRF=répertoire;Fichier d'import;IMPORT;SOCIETE;UTILISTATEUR;Adresse Email;TYPE écriture;TYPE écriture;Fichier RAPPORT;Gestion DOUBLON;Controle Paramsoc; Mode SANSECHEC;Option Minimisé
//exemple C:\PGI00\APP\COMSX.EXE /TRF=C:\tmp\FFF\;JOURNALGI.TRA;IMPORT;SECALDOS;CEGID;mentressangle@cegid.fr;N;N;RAPPORT.LOG;FALSE;FALSE;SANSECHEC;Minimized

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

Procedure DispatchUtil ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
Numopt         : integer;
BEGIN
     Numopt := Num;
   case Numopt of
     16 :
     begin
     V_PGI.StopCourrier:=TRUE ;
     end;
     10 :
     begin
{$IFNDEF EAGLCLIENT}
{$IFDEF SCANGED}
        if V_PGI.RunFromLanceur then
          InitializeGedFiles(V_PGI.DefaultSectionDbName, nil)
        else
          InitializeGedFiles(V_PGI.DbName, nil);
{$ENDIF}
{$ENDIF}
          V_PGI.SAV:=FALSE ; //Apres connection
      If Estspecif('YO') Then
        BEGIN
        FMenuG.BSQL.Tag:=21 ;
        END ;

     end;
     11 :
          begin
               V_PGI.SAV:=FALSE ;
//               TCPContexte.Release;

          end;//Après deconnection
     12 : BEGIN
             // Interdiction de lancer en direct
{$IFNDEF EAGLCLIENT}
          RenseignelaSerieGG ;
{$ENDIF}

          V_PGI.SAV:=FALSE ; //Avant connection ou seria
{$ifndef eAGLClient}
          FMenuG.OkVerifStructure:=FALSE ;
{$endif}
          END ;
     13 : V_PGI.SAV:=FALSE ; //Avant deconnection
     15 : ;
      21 :  ;
     100 : ;

     //45  Ecritures
     45101 : Brouillard('N');
     45103 : Brouillard('S');
     45105 : Brouillard('R');
     45107 : BrouillardAna('N');

     45211 : CumulPeriodique(fbJal,neJaG)     ;
     45213 : CumulPeriodique(fbJal,neJalPer)  ;
     45215 : JalCpteGe  ;
     45217 : CumulPeriodique(fbJal,neJalCentr);
     45219 : JalDivisio;
     45251 : CumulPeriodique(fbGene,neJalRien) ;
     45253 : CumulPeriodique(fbAux,neJalRien) ;
     45255 : CumulPeriodique(fbSect,neJalRien) ;

     //46  Tiers
     46111 : BalanceAuxi ;
     46113 : BalGeAu ;
     46115 : BalAuGe ;
     46117 : BalanceAge ;
     46119 : BalVentile ;

     46211 : GLAuxiliaire ;
     46213 : GLGenAuxi ;
     46215 : GLAuxiGen ;
     46217 : GLivreAge ;
     46219 : GLVentile ;
     46221 : GLAuxiliaire ;
     46223 : JustifSolde('',fbAux);

     46311 : BrouillardTP ;
     46313 : BalanceAge2TP;
     46315 : GLivreAgeTP;
     46317 : BalVentileTP;
     46319 : GLVentileTP;

     //47  Autres éditions
     47111 : BalanceGene ;
     47113 : BLAnal ;
     47115 : BLGESE ;
     47117 : BLSEGE ;
     47119 : BLAffaire;

     47211 : GLGeneral ;
     47213 : GLGESE ;
     47215 : GLSEGE ;
     47217 : GLAnal ;

     47311 : EditionTvaAcE;           // Etat des acomptes
     47313 : EditionTvaFac(False);    // Factures par étapes
     47315 : EditionTvaFac(True);     // Gestion directe des factures

     //
     45600 : BEGIN
             V_PGI.NumVersionBase:=750 ;
             AGLLanceFiche('YY', 'YYIMPORTOBJET', '', '', ''); // Import d'objets
{$IFDEF EAGLCLIENT}
             AvertirCacheServer('FORMES') ;
             AvertirCacheServer('FORMESP') ;
             AvertirCacheServer('MENU') ;
             AvertirCacheServer('MODELE') ;
             AvertirCacheServer('MODEDATA') ;
{$ENDIF}
             V_PGI.NumVersionBase:=0 ;
             END ;

//=======================================
//==== Administrations / Outils (18) ====
//=======================================

    // Utilisateurs et accès
    3165  : FicheUserGrp ;
    3170 : BEGIN FicheUSer(V_PGI.User) ; ControleUsers ; END ;
    60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','8;9;11;12;13;14;16;17;18;20;26;27;324;361');

    3172 : YYLanceFiche_ProfilUser;
    3180 : ReseauUtilisateurs(False) ;

    3185 : VisuLog1 ;

    3195 : ReseauUtilisateurs(True) ;
{$IFDEF CERTIFNF}
    { BVE 03.09.07 : Journal des évenements }
    3169 : CPLanceFicheCPJNALEVENT('','','');
{$ENDIF}

    3198 : AGLLanceFiche('CP', 'CPREINITFNCT', '', '', '');

     // liste des mouvement modifies et pointées
    else HShowMessage('2;?caption?;Fonction non disponible : ;W;O;O;O;',TitreHalley,IntToStr(Num)) ;
  end ;
END ;

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

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
VH^.OkModCPPackAvance := True ;
if ( VH^.OkModCPPackAvance ) or ( V_PGI.VersionDemo ) then
begin
  V_PGI.LaSerie := S7 ;
  //TEST Préférences dans agl pour style du menu :
  V_PGI.ChangeModeOutlook := True ;
end ;
VH^.OkModCompta:=True ; VH^.OkModBudget:=TRUE ; VH^.OkModImmo:=TRUE ;
//V_PGI.Monoposte:=TRUE ;
V_PGI.VersionDemo:=FALSE ;

V_PGI.NbColModuleButtons := 3;
V_PGI.NbRowModuleButtons := 6;
FMenuG.SetModules( [363,18], []);

END ;


procedure InitApplication ;
begin
     { Référence à la fonction PRINCIPALE qui permet de lancer les actions en fonction de la
       sélection d'une option des menus.
      }
     FMenuG.OnDispatch:=DispatchUtil ;

     { Référence à une fonction qui est lancée après la connexion à la base et le chargement du dictionnaire}
     FMenuG.OnChargeMag:=ChargeMagHalley ;

     { Référence à une fonction qui est lancée avant la mise à jour de structure}
     FMenuG.OnMajAvant:=Nil ;

{$IFNDEF EAGLCLIENT}
     { Référence à une fonction qui est lancée pendant la mise à jour de structure}
     FMenuG.OnMajpendant:=Nil ;
{$ENDIF}

     { Référence à une fonction qui est lancée après la mise à jour de structure}
     FMenuG.OnMajApres:=Nil ;
     FMenuG.OnChangeModule:=AfterChangeModule ;
     OkExport := FALSE;
     Minimise := FALSE;

FMenuG.OnAfterProtec:=AfterProtec ;
RenseignelaSerieGG ;
END ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin

end ;

Procedure InitLaVariablePGI;
Begin
Apalatys:='CEGID' ;
Copyright:='© Copyright ' + Apalatys ;

HalSocIni:='CEGIDPGI.INI' ;
If OkSic Then HalSocIni:='CEGIDPGISIC.INI' ;
V_PGI.OutLook:=FALSE ;
V_PGI.VersionDemo:=FALSE ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionDEV:=TRUE ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.OfficeMsg:=True ;
V_PGI.NumMenuPop:=27 ;
V_PGI.DispatchTT:=DispatchTT ;
V_PGI.ParamSocLast:=False ;
V_PGI.RAZForme:=TRUE ;
V_PGI.NoModuleButtons:=False ;
V_PGI.PGIContexte:=[ctxCompta] ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;

V_PGI.SAV:=False ;
V_PGI.VersionReseau:=True ;
V_PGI.CegidAPalatys:=FALSE ;
V_PGI.CegidBureau:=TRUE ;
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=False ;
V_PGI.MultiUserLogin:=False ;

V_PGI.NumMenuPop:=27 ;
V_PGI.OfficeMsg:=True ;
V_PGI.NoModuleButtons:=FALSE ;
V_PGI.NbColModuleButtons:=1 ;
V_PGI.LaSerie := S5;
RenseignelaSerieGG ;
if GetSynRegKey('Outlook',2,True)=2 then
  SaveSynRegKey('Outlook',False,True);
ChargeXuelib ;
end;

procedure RemoveFromMenu ( const iTag : integer; const bGroup : boolean; var stAExclure : string );
begin
  if bGroup then FMenuG.RemoveGroup (iTag, True)
  else FMenuG.RemoveItem (iTag);
  stAExclure := stAExclure+IntToStr(iTag)+';';
end;

procedure AfterChangeModule ( NumModule : integer ) ;
var stAExclure : string;
BEGIN
  RemoveFromMenu (3100, True, stAExclure) ;
  RemoveFromMenu (3125, True, stAExclure) ;
  RemoveFromMenu (-3200, True, stAExclure) ;
  RemoveFromMenu (-3240, True, stAExclure) ;
  RemoveFromMenu (-31, True, stAExclure) ;
END;


Initialization
{Version}
If Pos('SIC',AnsiUpperCase(ParamStr(0)))>0 Then
  BEGIN
  OKSIC:=TRUE ;
  If PARAMCOUNT=2 Then If paramstr(1)='CEGID' Then OkSIC:=FALSE ;
  END Else OKSIC:=FALSE ;
OkSic:=FALSE ;
ProcChargeV_PGI :=  InitLaVariablePGI ;

end.

