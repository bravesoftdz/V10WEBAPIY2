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
unit MenuAdm;

interface
Uses
  UtilPGI,
  Windows, Messages,
  HEnt1,EntPGI,
{$ifdef eAGLClient}
  MaineAGL,eTablette,
{$else}
  MenuOLG, FE_Main,
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$endif eAGLClient}

  Forms,sysutils,HMsgBox, Classes, Controls, HPanel,
  hctrls, MajTable,
  ExtCtrls, inifiles, UTOB,
  Utilsoc,
  Reseau, ImgList,
{$IFDEF ADRESSE}
    PARAMADRE_TOM,    // Gestion des adresses.
    SAISIEADR_TOF,
{$ENDIF}
    EdtEtat,
{$IFDEF EAGLCLIENT}
    utileagl, // ConnecteHalley
    MenuOLX,
    About,
    Buttons,
    ComCtrls,
    Graphics,
    Hout,
    M3FP,
    Prefs,
    {$IFDEF CONTREPARTIE}
    UtofReparationfic,  // ajout me
    {$ENDIF}
{$ELSE}
    uedtcomp,
    Tablette,
{$ENDIF EAGLCLIENT}
    Vierge,
    dialogs,
    HTB97,
    UIUtil,
    AglInitPlus,
    AGLInit,
    uMultiDossier,
    GalOutil,
    AglMail,
    Exercice_Tom,
    Souche_Tom,
    FichComm,
    CPREGION_TOF,
    CPCODEPOSTAL_TOF,
    CPTVATPF_TOF,
    UTOFMULPARAMGEN,
    PAYS_TOM,
    LGCOMPTE_TOF,
    CPAXE_TOM,
    CPSTRUCTURE_TOF,
    UTOTVENTILTYPE,
    TomIdentificationBancaire,
    CPCODEAFB_TOF,
    OuvreExo,
    Ent1,
    UTOFCPJALECR,
    UTOMBanque,
{$IFDEF EAGLCLIENT}
{$IFNDEF ENTREPRISE}
    BobGestion,
{$ENDIF}
{$ELSE}
    BobGestion,
{$ENDIF}
    Confidentialite_Tof
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

Uses LicUtil,
     ParamSoc {,UserGrp_tom, UTOMUTILISAT, Confidentialite_TOF, MulLog1,}
     {TomProfilUser};

Procedure VerifInstall (BobPrefixe,BobSuffixe,NomChampParamSoc : String) ;
Type terrbob = (bobOK, bobPasla, bobLaMaisPasIntegre, bobPasLaMaisIntegre, bobAInteger) ;
Var Q : tQuery ;
    BobName,St1 : String ;
    FindIt : Boolean ;
    datecreation : tDateTime ;
    errb : terrbob ;
BEGIN
//BobName:='CIS50751M60.BOB' ;
// SO_I_FISACTIVTPTF
{$IFDEF EAGLCLIENT}
AvertirCacheServer('YMYBOBS') ;
AvertirCacheServer('PARAMSOC') ;
{$ENDIF}
BobName:='CCAD'+BobPrefixe+BobSuffixe ;
ErrB:=BobOK ;
If BobName='' Then Exit ;
//If NomChampParamSoc='' Then Exit ;
(*
Q := OpenSql('SELECT SOC_NOM FROM PARAMSOC WHERE SOC_NOM ="'+NomChampParamSoc+'" ',True, 1);
If Q.Eof Then ErrB:=bobPasLaMaisIntegre ;
Ferme(Q) ;
*)
//If ErrB=BobOK Then Exit ;

Q := OpenSql('SELECT * FROM YMYBOBS WHERE YB_BOBNAME="' + BobName + '"',True, 1);
DateCreation:=iDate1900 ;
if not q.eof then datecreation := q.findfield('YB_BOBDATECREAT').asdatetime;
FindIt := not Q.Eof;
Ferme(Q);
If FindIt Then
  BEGIN
  ErrB:=bobOk ;
  END Else
  BEGIN
  ErrB:=bobPasla ;
  END ;

Case ErrB Of
  bobPasla : BEGIN
             St1 := 'UN FICHIER SYSTEME ('+BobName+'.BOB) N''A PAS ETE INTEGRE. ';
             St1 := St1 +#13#10 +#13#10 +'CELA PEUT DECLENCHER DES DYSFONCTIONNEMENTS.';
             St1 := St1 +#13#10 +#13#10 +'NOUS VOUS CONSEILLONS D''APPELER LE SUPPORT TELEPHONIQUE CEGID.';
             PGIInfo(St1,'AVERTISSEMENT');
             END ;
  bobLaMaisPasIntegre : BEGIN
                        St1 := 'UN FICHIER SYSTEME ('+BobName+'.BOB intégré le '+DateToStr(DateCreation)+ ') N''EST PAS CORRECT. ';
                        St1 := St1 +#13#10 +#13#10 +'CELA PEUT DECLENCHER DES DYSFONCTIONNEMENTS.';
                        St1 := St1 +#13#10 +#13#10 +'NOUS VOUS CONSEILLONS D''APPELER LE SUPPORT TELEPHONIQUE CEGID.';
                        PGIInfo(St1,'PROBLEME DETECTE ! ');
                        END ;
  END ;
END ;

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
St,lSt         : String ;
BEGIN
     Numopt := Num;
   case Numopt of
     16 :
     begin
// ??? GP le 15/09/2008 vu ce que faut la procédure appelée, ca peut pas être un mal ...       PGI_IMPORT_BOB('CCAD');
        St := 'CCAD';
        lSt := 'Administration comptable';
// Modif GP : plus d'import auto en mode CWAS.
//     BOB_IMPORT_PCL_STD
{$IFDEF EAGLCLIENT}
{$IFDEF ENTREPRISE}
        VerifInstall ('0851M','001','') ;
{$ELSE}
        BOB_IMPORT_PCL_STD (St,lSt);
{$ENDIF}
{$ELSE}
        BOB_IMPORT_PCL_STD (St,lSt);
{$ENDIF}
     end;
     10 :
     begin
{$IFNDEF EAGLCLIENT}

{$ENDIF}
     V_PGI.SAV:=FALSE ; //Apres connection

     end;
     11 :
          begin
               V_PGI.SAV:=FALSE ;
//               TCPContexte.Release;

          end;//Après deconnection
     12 : BEGIN
             // Interdiction de lancer en direct

          V_PGI.SAV:=FALSE ; //Avant connection ou seria
{$ifndef eAGLClient}
          FMenuG.OkVerifStructure:=FALSE ;
{$endif}
          END ;
     13 : V_PGI.SAV:=FALSE ; //Avant deconnection
     15 : ;
      21 :  ;
     100 : ;

      // PARAMETRES GENERAUX
      1104 : begin
             ParamSociete(False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX', '',
             RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 1105000);
           end;
      1290 : YYLanceFiche_Exercice('0');  // Exercice
      // Compteurs
      1300 : YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT'); //Compteurs
      1110 : FicheEtablissement_AGL(taModif) ;

      // Tables
      1120 : ParamTable('ttFormeJuridique',taCreat,1120000,PRien) ;
      1122 : ParamTable('ttCivilite',taCreat,1122000,PRien) ;
      1125 : ParamTable('ttFonction',taCreat,1125000,PRien) ;
      1135 : OuvrePays ;
      1140 : FicheRegion('','',False) ;
      1145 : OuvrePostal(PRien) ;

      1165 : ParamTable('TTREGIMETVA',taCreat,1165000,PRien) ;
      1170 : ParamTvaTpf(true) ;   // Tables // TVA par régime fiscal
      1175 : ParamTvaTpf(false) ;  // Tables // TPF par régime fiscal
      1185 : FicheModePaie_AGL('');
      1190 : FicheRegle_AGL('',FALSE,taModif) ;
      1193 : ParamTable('ttNivCredit',taCreat,1190040,PRien) ;


      // STRUCTURES
      7211 : CPLanceFiche_MULJournal('C;7211000');  //MulticritereJournal(taModif);
      // Structures
      7112 : CCLanceFiche_MULGeneraux('C;7112000'); //MulticritereCpteGene(taModif);
      7145 : CPLanceFiche_MULTiers('C;7145000');    //MultiCritereTiers(taModif);
      7118 : CCLanceFiche_MULGeneraux('S;7118000');  // Suppressions // Comptes généraux
      7214 : CPLanceFiche_MULJournal('S;7214000');   // Suppressions // Journaux
      7394 : CPLanceFiche_CPJALECR ; // 26/11/2002 nv journal des ecritures
      3205 : CCLanceFiche_LongueurCompte ;

      // ANALYTIQUE
      1105 : FicheAxeAna ('') ;
      7178 : begin
//             if VersionAnalytique=2 then
//               CPLanceFiche_MULSectSP('C;7178000')  //MulticritereSection(taModif);
//             else
               CPLanceFiche_MULSection('C;7178000');  //MulticritereSection(taModif);
             end;
      1450 : ParamPlanAnal('');      // Analytiques / Structures analytiques
      1460 : ParamVentilType(PRien); // Analytiques / Ventilations types

      //BANQUE
      1701 : LanceFicheIdentificationBancaire;
      1705 : if EstComptaTreso then TRLanceFiche_Banques('TR','TRCTBANQUE', '', '', '')
                               else FicheBanque_AGL('',taModif,0);
      1720 : ParamTable('ttCFONB',tacreat,500005,PRien);   // Codes CFONB
      1725 : ParamCodeAFB;                                 // Code AFB
      1730 : ParamTable('ttCodeResident',taCreat,500007,PRien) ;

      // TRAITEMENTS
      7718 : OuvertureExo ; // Ouverture d'exercice

      //ACCES
      60208 : GCLanceFiche_Confidentialite( 'YY','YYCONFIDENTIALITE','','','365');

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
FMenuG.SetModules( [365], []);

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
V_PGI.OutLook:=True ;
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
RenseignelaSerie(exeCCADM);
if GetSynRegKey('Outlook',2,True)=2 then
  SaveSynRegKey('Outlook',True,True);
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

