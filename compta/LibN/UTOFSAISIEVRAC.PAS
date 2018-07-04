{***********UNITE*************************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... :   /  /
Description .. : Source de la saisie en vrac
Mots clefs ... :
*****************************************************************}
Unit UTOFSAISIEVRAC ;

Interface

Uses StdCtrls, Controls, Classes,    // grids
     {$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}

     {$IFDEf VER150}
     Variants,
     {$ENDIF}
     forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,
     // composant AGL
     HTB97,
     HSysMenu,
     HFLabel,
     UIUtil,
     // VCL
     Windows,
     Messages,
     Graphics,
     Dialogs,
     Menus,
     Grids,          // pour TGridDrawState
     extctrls,       // pour le TBevel
     HPanel,
     // lib
     Ent1,
     {$IFNDEF EAGLCLIENT}
     FE_Main,        // pour AGLLanceFiche()
     {$ELSE}
     MaineAGL,
     {$ENDIF}
     // Objet de gestion
     ULibEcriture,
     UObjFiltres,
     // AGL
     UTOB,
     ZGuide,
{$IFDEF EAGLCLIENT}
{$ELSE}
     MenuOLG,
{$ENDIF}
     UTOF ;

Const
   cStTexteTitreFenetre = 'Saisie en vrac' ;

Type

  TZSaisieVrac       = Class ;
  TZIntegrationVrac  = Class ;

  TOF_CPSAISIEVRAC = Class (TOF)
   private
    // variable interne
    FStatut                 : TActionFiche ;          // Etat de la grille ( modif,creation...)
    FBoGuideRun             : boolean;                // True la saisie par guide est active
    FExceptionEvent         : TExceptionEvent;        // evenement gerant toutes les exceptions non gérées de la saisies
    FRechCompte             : TRechercheGrille ;      // objet de recherche des comptes
    FInfoEcr                : TInfoEcriture;          // objet contenant la liste des journaux,... partage entre par l ZEtebac et les guides
    FMessCompta             : TMessageCompta ;        // affichage des messages
    FZSaisieVrac            : TZSaisieVrac ;          // objet de gestion
    FZIntegration           : TZIntegrationVrac ;     // objet d'integration des ecritures
    FZGuide                 : TZGuide;                // objet de gestion des guides
    FZNature                : TZNature ;              // objet de gestion des natures ( gere
    FBoAfficherNature       : boolean ;               // false la case nature n'apparit pas : parametre SO_NATSAISIEVRAC
    FBoAfficherEtablis      : boolean ;               // false la case etablissement n'apparait pas ( si un seul etablisssment )
    FDtDateDeb              : TDatetime ;             // date de debut, valeur par defaut pour e_datecomptable
    FDtDateFin              : TDatetime ;             // date de fin ( pour le mul )
    FStTitre                : string ;                // titre de la fenetre
    FObjFiltre              : TObjFiltre ;            // objet de gestion du filtre. On lui donne un panel et il s'occupe de tout
    FRdSoldeDebit           : double ;
    FRdSoldeCredit          : double ;
    FRdSolde                : double ;
    FBoValidationEnCours    : boolean ;
    FBoAffMessErr           : boolean ;
    _Plantage               : integer ;
    FInLigneCourante        : integer ;
    FInIndexComptaAuto      : integer ;
    FStLastJal              : string ;
    FInNumGroupEcr          : integer ;
    FBoGrilleVide           : boolean ;               // true si la grille est en creation mais que l'on n'a rien enregistré
    FBoGrilleModif          : boolean ;               // true si on a modifie la grille, different du statut ( consult si on a enregsitrer )
    FStGenHT                : string ;
//    FInNbLigneGuide         : integer ;
    // liste des controle
    FPageControl            : TPageControl;           // Page à onglet de recherche
    TbsStandard             : TTabSheet;              // Onglet standard
    FListe                  : THGrid;                 // Grille de saisie
    BInsert                 : TToolbarButton97;       // Bouton BInsert
    BCherche                : TToolbarButton97;       // Bouton 'Mouette'
    BDelete                 : TToolbarButton97;       // Bouton BCherche
    BValider                : TToolbarButton97;       // Bouton Valider : coche verte
    BImprimer               : TToolbarButton97;       // Bouton imprimer
    BDefaire                : TToolbarButton97;       // Bouton Annulation
    BAgrandir               : TToolbarButton97;       // Bouton Agrandir la liste
    BMenuZoom               : TToolbarButton97;       // Bouton Zoom
    BIntegre                : TToolbarButton97;       // Bouton Integration
    E_JOURNAL               : THCritMaskEdit;
    E_ETABLISSEMENT         : THCritMaskEdit;
    E_UTILISATEUR           : THCritMaskEdit;
    E_DATECOMPTABLE         : THCritMaskEdit;
    E_DATECOMPTABLE_        : THCritMaskEdit;
    HSystemMenu1            : THSystemMenu;
    FFiltres                : THValComboBox;          // Liste des filtres disponibles
    POPF                    : TPopupMenu;
    POPS                    : TPopupMenu;
    POPZOOM                 : TPopupMenu;
    BSelectionner           : TToolbarButton97;       // Bouton BCherche
    BDeSelectionner         : TToolbarButton97;
    SOLDEDEBIT              : THLabel;
    SOLDECREDIT             : THLabel;
    SOLDE                   : THLabel;
    CSOLDEDEBIT             : TBevel;
    CSOLDECREDIT            : TBevel;
    CSOLDE                  : TBevel;
    LSOLDE                  : THLabel;
    G_LIBELLE               : THLabel;
    T_LIBELLE               : THLabel;
    FlashGuide              : TFlashingLabel;
    procedure ChargeParam;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormKeyDown(Sender : TObject; var Key : Word;Shift : TShiftState);
    procedure FormKeyDownFliste(Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure FormKeyDownForms(Sender : TObject; var Key : Word; Shift : TShiftState);
//    procedure GestionException(Sender : TObject; E: Exception);
    procedure POPSPopup(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure FListeCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure FListeCellExitModeGuide(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure FListeRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure FListeRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FListeElipsisClick(Sender: TObject);
    procedure FListePostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;AState: TGridDrawState);
    procedure FListeDeleteRow ( Row : integer = - 1 );
    procedure FListeKeyPress (Sender : TObject ; var Key : Char);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValiderClick(Sender: TObject);
 //   procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure E_JOURNALElipsisClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject) ;
    procedure BDeleteClick(Sender: TObject) ;
    procedure BAgrandirClick(Sender: TObject);        // Bouton Agrandir Liste
    procedure BIntegreClick(Sender: TObject);        // Bouton Agrandir Liste
    procedure E_JOURNALExit(Sender: TObject);

    // gestion du MUL et enregistrement en base
    procedure BChercheClick(Sender: TObject) ;
    procedure ChercheClick( vBoEnMemoire : boolean = true ) ;
    procedure UpdateGrille ;
    function  RemplitGrille : boolean ;
    function  Validation : boolean ;
    function  RefreshFListe( vBoEnMemoire : boolean = true ) : integer ;
    function  RefreshLigne( vIndexGrille : integer ; vIndexTOB : integer = - 1 ) : boolean ;
    function  ControleCritere(vTOB : TOB): boolean ;
    // gestion des erreurs
    procedure OnErrorRech(sender : TObject ; Error : TRecError);
    procedure OnErrorTOB(sender : TObject ; Error : TRecError);
    function  RecupTOB( ARow : integer = - 1 ) : TTOBEcriture;
    procedure CompleteCeQueJePeux( ACol : integer = - 1 ; ARow : integer = - 1 );
    // gestion des deplacemennts
    procedure NextRow;
    function  ControleColPrecedent(var ACol, ARow: integer ) : boolean ;
    procedure CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );
    procedure InitRow ;
    procedure AfficheTitreAvecCommeInfo (vInfo: string = '' ) ;
    procedure AfficheInfoLigne ;
    procedure AfficheSolde;
    function  LigneCorrecte ( ARow : integer = - 1 ) : boolean;
    procedure SetStatut(Value: TActionFiche);
    procedure NextNature ;
    procedure EnableButtons( Ou : integer=-1) ;
    function  GetNature : string ; // retourne la nature de la grille ou OD si la nature n'est pas affichées
    procedure InitGrid ;
    procedure TrouveIndexPiece ( vInNumGroupeEcr : integer ; var vInIndexDeb : integer ; var vInIndexFin : integer) ;
    function  CalculSeparateur(ACol, ARow: Integer ; var vBoHaut: boolean ; vBo : boolean = false  ) : boolean ;
    procedure AppelCalculatrice ;
    // fonction
    procedure GereAna( ARow : integer = - 1 ; vBoAfficher : boolean = false );
    procedure GetAna( vTOB : TOB );
    procedure VK_F7FListe ;
    procedure SoldeFolio;
    procedure CalculSoldeFolio( vInNumGroupeEcr : integer = -1 );
    procedure CalculInfoSupp( vTOB,vTOBPrec,vTOBSuiv : TOB  ; vBoInsert : boolean = false );
    procedure NumeroteGrille;
    procedure GetCompteAuto( vInIndice : Integer) ;
    function  AccelerateurDeSaisie : integer ; //( Ou : integer = - 1 ) : integer ;
    function  GetCompteAcc ( Ou : integer ; E_AUXILIAIRE : string ) : string ;
    function  IsRowHT ( vRow : Integer ) : Boolean ;
    // gestion des guides
    procedure GuiClick ;
    procedure NextColStop ( var ACol , ARow : integer ) ;
    procedure SetGuideRun ( Value : boolean ) ;
    procedure FinGuide ;

    property  Statut              : TActionFiche read FStatut      write SetStatut;
    property  BoGuideRun          : boolean      read FBoGuideRun  write SetGuideRun ;
    function GetFormulePourCalc(Formule: hstring): variant;


   protected
    function  GetControlTOF           : boolean; // recupere les contrôles de la grille
    function  AssignEvent             : boolean; // Affecte les evenements au contrôle
    function  InitControl             : boolean; // initialise les grilles
    function  CreateControl           : boolean; // creation de TOB
  public
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  end ;

  TZSaisieVrac = Class(TOBCompta)
  private
   FListeASupprimer : TStringList ;
   FTOBAna          : TOB ;
   FStSQLFiltre     : string ;
   function    ProchainCompteur : integer;
   procedure   Numerote ;
   function    InternalAdd ( vIndex : integer = - 1) : TTOBEcriture;
  protected
  public
   destructor  Destroy ; override ;
   procedure   Initialize ; override ;
   function    Add : TTOBEcriture ;
   function    Insert(vIndex : integer) : TTOBEcriture;
   procedure   Delete(vTOB : TOB) ;
   procedure   Save ;
   procedure   Load ;
   procedure   LoadAna( vTOB : TOB ) ;
   procedure   RemiseAZeroControleur ;
   procedure   AddAna(  vTOB : TOB ) ;
   property    StSQLFiltre : string read FStSQLFiltre write FStSQLFiltre ;
  end;

 TZIntegrationVrac = Class(TObjetCompta)
 private
  FTOBBorN1   : TOB ;
  FTOBPieceN1 : TOB ;
  FTOBLibN1   : TOB ;
  FTOBBorN3   : TOB ;
  FTOBPieceN3 : TOB ;
  FTOBLibN3   : TOB ;
  FTOBN       : TOB ;
  FTOBGarbage : TOB ;
  FCR         : TList ;
  FEcran      : TForm ;
  procedure   AjouteN3aN1( vTOBEcrN1, vTOBEcrN3 : TOB );
  procedure   AjouteN2aN1( vTOBEcrN1, vTOBEcrN2 : TOB );
  function    ChercheN2dansN1 ( vTOBEcrN1 , vTOBEcrN2 : TOB ; var vInNumGroupeEcr : integer ) : TOB;
  procedure   StockeLesEcrituresIntegrees(vTOB : TOB) ;
  procedure   RAZFLC;
 public
  constructor Create( vInfoEcr : TInfoEcriture ) ; override ;
  destructor  Destroy ; override ;
  procedure   Save ;
  function    RecupereLigne( vTOB : TOB ) : boolean;
  procedure   ReAffecteLigne ;
  property    CR    : TList read FCR ;
  property    Ecran : TForm read FEcran  write FEcran ;

 end;


  procedure CPLanceFiche_SaisieVrac;

  function   GetE ( FListe : THGrid ; Lig : integer = -1 ) :  TTOBEcriture ;
  procedure  SetE ( FListe : ThGrid ; O : TTOBEcriture ; Lig : integer = -1 ) ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  ImRapInt,
  SaisUtil,      // pour RDevise
  UtilSais,      // pour les constantes de grille
  SaisComm,      // pour le getO
  ParamSoc,      // pour le GetParam
  Filtre,        // gestion du filtre
  AGLInit,
  Vierge,        // pour le form closeQuery
  ed_tools,      // pour le TQRProgressForm
  ParamDat,
  LookUp,
  eSaisAnal,
  ZCompte,
  Choix,
  ULibWindows,
  UTOFLOOKUPTOB,
  HQry ,        // pour le recupWhereCritere
  ULibAnalytique,
  Echeance,
  ZEch,
  hCalc,
  Formule ,
  ULibExercice,
  HCompte,      // pour la saisie analytique
  //unit1 ,
  UtilPGI ;      // pour le _Blocage

{$IFDEF TT}

var
 TheLastError : string ;

var
 TheLog : TStringList;

procedure DelEvenement ;
begin
 TheLog.Clear ;
 TheLog.SaveToFile('c:\SaisieVrac.txt') ;
end;

procedure AddEvenement( value : string );
begin
 if TheLog = nil then exit ;
 TheLog.Add(Value) ;
 TheLastError:=value;
 TheLog.SaveToFile('c:\SaisieVrac.txt') ;
end;
{$ENDIF}

{ TZSaisieVrac }


function CGetCleVrac : string ;
begin
 result := Copy( 'V' +  V_PGI.User , 1 , 3 ) ;
end;

function CGetNumAxe( vAxe : string ) : Integer ;
begin
 Result:=StrToInt(Copy(vAxe,2,1)) ;
end ;

function _GetCleBor ( vTOB : TOB ) : string ;
begin
 result := 'E_EXERCICE="'           + vTOB.GetValue('E_EXERCICE')                  + '" AND ' +
           'E_JOURNAL="'            + vTOB.GetValue('E_JOURNAL')                   + '" AND ' +
           'E_DATECOMPTABLE ="'     + UsDateTime(vTOB.GetValue('E_DATECOMPTABLE')) + '" AND ' +
           'E_NUMEROPIECE='         + IntToStr(vTOB.GetValue('E_NUMEROPIECE'))     + ' AND '  +
           'E_NUMECHE='             + IntToStr(vTOB.GetValue('E_NUMECHE'))         + ' AND '  +
           'E_QUALIFPIECE="'        + vTOB.GetValue('E_QUALIFPIECE')               + '" AND ' +
           'E_CONTROLEUR="'         + vTOB.GetValue('E_CONTROLEUR')                + '" '     ;
end;



procedure _CSuppAna( vTOB : TOB ) ;
begin
 if vTOB.GetValue('E_ANA') = 'X' then
  executeSql( 'DELETE ANALYTIQ WHERE Y_JOURNAL="' + vTOB.GetValue('E_JOURNAL') + '" ' +
              'AND Y_EXERCICE="'            + vTOB.GetValue('E_EXERCICE')                  + '" ' +
              'AND Y_DATECOMPTABLE="'       + usDateTime(vTOB.GetValue('E_DATECOMPTABLE')) + '" ' +
              'AND Y_NUMEROPIECE='          + intToStr(vTOB.GetValue('E_NUMEROPIECE'))     + ' '  +
              'AND Y_NUMLIGNE='             + intToStr(vTOB.GetValue('E_NUMLIGNE'))        + ' '  +
              'AND Y_QUALIFPIECE="'         + vTOB.GetValue('E_QUALIFPIECE')               + '" ' ) ;
end;

function _AvecAna( vTOB : TOB ) : boolean ;
begin
 result := ( vTOB.Detail <> nil ) and ( vTOB.Detail.Count <> 0 ) ;
end;

function _MakeCrit ( vTOB : TOB ) : string ;
var
 lStCritDate : string ;
begin

 if ( UpperCase(vTOB.GetValue('E_MODESAISIE')) = 'LIB' ) then
  lStCritDate := intToStr(vTOB.GetValue('E_PERIODE'))
   else
    lStCritDate := DatetoStr(vTOB.GetValue('E_DATECOMPTABLE')) ;

 result := vTOB.GetValue('E_EXERCICE')                 +
           vTOB.GetValue('E_ETABLISSEMENT')            +
           vTOB.GetValue('E_JOURNAL')                  +
           lStCritDate                                 +
           vTOB.GetValue('E_NATUREPIECE') ;
end;


procedure _AddChampsSuppVrac( vTOB : TOB ) ;
begin
 vTOB.AddChampSupValeur('LIBCOMPTE'  , '-') ; // si 'X' => le, libelle de l'ecriture est le libelle du compte
 vTOB.AddChampSupValeur('ANA'        , '-') ;       // si 'X' => analytique sur la ligne
 vTOB.AddChampSupValeur('SUPPRIMER'  , '-') ; // si 'X' => ligne a supprimer
 vTOB.AddChampSupValeur('CODEERREUR' , RC_PASERREUR) ;
 vTOB.AddChampSupValeur('SOLDE'      , 0) ; // solde de chaque ecriture
 vTOB.AddChampSupValeur('CRIT'       , '') ; // crit pour les rupture de piece
 vTOB.AddChampSupValeur('GRILLE'     , 'X') ; // si X l'enregistrement est affiche ds la grille
end ;


///////////////////////////////////////////////////////////////////////////////
//
///////////////////////////////////////////////////////////////////////////////

procedure TZSaisieVrac.Initialize ;
begin
 inherited ;
 FTOBAna          := TOB.Create('',nil,-1) ;     
 FListeASupprimer := TStringList.Create ;
end ;


function TZSaisieVrac.InternalAdd ( vIndex : integer = -1 ) : TTOBEcriture;
begin
 result         := TTOBEcriture.Create('ECRITURE',self,vIndex) ;
 result.Info    := Info ;
 Result.OnError := OnError ;
 result.PutDefautEcr ;
 result.AddChampSupValeur('ACC'        , '-' ) ;
 result.AddChampSupValeur('_GUIDE'     , '000' ) ;
 result.AddChampSupValeur('_NUMLIGNE'  , '0' ) ;
 result.PutValue('E_NUMEROPIECE'       , -1 );
 result.PutValue('E_NUMLIGNE'          , -1 );
 result.PutValue('E_QUALIFORIGINE'     , 'VRA' );
 result.PutValue('E_QUALIFPIECE'       , 'V' );
 result.PutValue('E_NUMGROUPEECR'      , - 1 );
 result.PutValue('E_UTILISATEUR'       , V_PGI.User );
 result.PutValue('E_CONTROLEUR'        , CGetCleVrac );
 _AddChampsSuppVrac(result) ;
end;


function TZSaisieVrac.Add : TTOBEcriture;
begin
 result := InternalAdd ;
end;

function TZSaisieVrac.Insert( vIndex : integer ) : TTOBEcriture;
begin
 result := InternalAdd(vIndex) ;
end;


procedure TZSaisieVrac.AddAna(  vTOB : TOB ) ;
var
 lStCle : string ;
begin
 lStCle := CGetCleAna(vTOB) ;
 if FListeASupprimer.IndexOf(lStCle) = -1 then
  FListeASupprimer.Add(lStCle) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 09/08/2004
Modifié le ... :   /  /
Description .. : - FB 13868 - LG - 09/08/2004 - on reappliqe les critures du
Suite ........ : filtre sinon on supprime toutes les ecritures vrac presentes
Mots clefs ... :
*****************************************************************}
procedure TZSaisieVrac.Load ;
var
 lTOB : TOB ;
	lStSQL    : string;

  lQ   : TQuery ;

begin

  // Construction requête
 lStSQL    := CGetSQLFromTable('ECRITURE',['E_BLOCNOTE','E_TRACE'] ) ;
 lStSQL    := lStSQL + ' ' + FStSQLFiltre ;
 lStSQL    := lStSQL + ' AND E_CONTROLEUR ="" AND E_QUALIFPIECE="V" ORDER BY E_NUMGROUPEECR,E_NUMEROPIECE,E_NUMLIGNE' ;


 // chargement des donnéees
  lQ := OpenSql( lStSQL, True ) ;
  ClearDetail ;
  try
   while not lQ.EOF do
    begin
     lTOB := Add ;
     lTOB.SelectDB('',lQ) ;
     lTOB.PutValue('E_CONTROLEUR' , CGetCleVrac );
     lQ.Next ;
    end; // while
  finally
   Ferme(lQ) ;
  end; //try


 ExecuteSQL('UPDATE ECRITURE set E_CONTROLEUR="' + CGetCleVrac + '" ' + FStSQLFiltre + ' AND E_CONTROLEUR ="" AND E_QUALIFPIECE="V" ') ;

end;

procedure TZSaisieVrac.Delete( vTOB : TOB ) ;
begin
 vTOB.ChangeParent(FTOBAna,-1) ;
end;

procedure TZSaisieVrac.LoadAna( vTOB : TOB ) ;
begin
 CChargeAna(vTOB) ;
 AddAna(vTOB) ;
end;

function TZSaisieVrac.ProchainCompteur : integer;
begin
 // on calcul le prochain numero et on le met à jour en meme temps
 result := CCalculProchainNumeroSouche ('VRA','NOR');
 if result = -1 then
  begin // creation de la souche
   ExecuteSQL('insert into SOUCHE( '                 +
              'SH_TYPE, '                            +
              'SH_SOUCHE, '                          +
              'SH_LIBELLE, '                         +
              'SH_ABREGE, '                          +
              'SH_NATUREPIECE, '                     +
              'SH_NUMDEPART, '                       +
              'SH_JOURNAL, '                         +
              'SH_MASQUENUM, '                       +
              'SH_SOCIETE, '                         +
              'SH_DATEDEBUT, '                       +
              'SH_DATEFIN, '                         +
              'SH_FERME, '                           +
              'SH_ANALYTIQUE, '                      +
              'SH_NATUREPIECEG, '                    +
              'SH_SIMULATION, '                      +
              'SH_NUMDEPARTS, '                      +
              'SH_NUMDEPARTP, '                      +
              'SH_SOUCHEEXO, '                       +
              'SH_RESERVEWEB ) '                     +
              ' values ('                            +
              '"VRA", '                              +
              '"NOR", '                              +
              '"Compteur saisie en vrac", '          +
              '"Vrac", '                             +
              '"", '                                 +
              '1, '                                  +
              '"", '                                 +
              '"", '                                 +
              '"'+V_PGI.CodeSociete+'", '            +
              '"'+USDateTime(iDate1900)+'", '        +  // FQ 14538 - CA - 09/09/2004
              '"'+USDateTime(iDate2099)+'", '        +  // FQ 14538 - CA - 09/09/2004
              '"-", '                                +
              '"-", '                                +
              '"", '                                 +
              '"-", '                                +
              '1, '                                  +
              '1, '                                  +
              '"-", '                                +
              '"-" ) ');
   result := 1;
  end; // if

end;


procedure TZSaisieVrac.Numerote ;
var
 i           : integer ;
 lInNumPrec  : integer ;
 lInNumLigne : integer ;
begin

 if ( Detail = nil ) or ( Detail.Count = 0 ) then exit ;

 lInNumPrec  := Detail[0].GetValue('E_NUMEROPIECE') ;
 lInNumLigne := Detail[0].GetValue('E_NUMLIGNE') ;

 if lInNumPrec = - 1 then
  begin
   lInNumPrec  := ProchainCompteur ;
   lInNumLigne := 1 ;
   Detail[0].PutValue('E_NUMEROPIECE' , lInNumPrec ) ;
   Detail[0].PutValue('E_NUMLIGNE'    , lInNumLigne ) ;
  end; // if

 for i := 1 to ( Detail.Count - 1 ) do
  begin

   lInNumPrec := Detail[i-1].GetValue('E_NUMEROPIECE') ;

   if Detail[i].GetValue('E_NUMEROPIECE') = - 1 then
    begin
     inc(lInNumLigne) ;
     Detail[i].PutValue('E_NUMEROPIECE' , lInNumPrec ) ;
     Detail[i].PutValue('E_NUMLIGNE'    , lInNumLigne ) ;
    end // for
     else
      if ( Detail[i].GetValue('E_NUMEROPIECE') <> lInNumPrec ) then
       begin
       // lInNumPrec  := Detail[i].GetValue('E_NUMEROPIECE') ;
        lInNumLigne := Detail[i].GetValue('E_NUMLIGNE') ;
       end// if
        else
         begin
          inc(lInNumLigne) ;
          Detail[i].PutValue('E_NUMLIGNE' , lInNumLigne ) ;
         end;

  end; // for

 CNumeroPiece(self) ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 09/08/2004
Modifié le ... : 28/11/2005
Description .. : - FB 13868 - LG - 09/08/2004 - on reappliqe les critures du
Suite ........ : filtre sinon on supprime toutes les ecritures vrac presentes
Suite ........ : - FB 17091 - LG - 28/11/2005 - coorection de l'integration, 
Suite ........ : on ignore l'erreur rc_qualifpiecems et on affiche les message 
Suite ........ : d'erreur ( voir pq on affichait - 2 )
Mots clefs ... :
*****************************************************************}
procedure TZSaisieVrac.Save ;
var
 i         : integer ;
 lTOB      : TOB ;
 lRecError : TRecError ;
begin

 {$IFDEF TT}
  AddEvenement('save');
{$ENDIF}

 if ( Detail = nil ) and ( Detail.Count = 0 ) then exit ;

 for i := 0 to FTOBAna.Detail.Count - 1 do
  _CSuppAna(FTOBAna.Detail[i] ) ;

 CSupprimeLigneSaisieVide(self) ;
 i    := 0 ;

 while i < Detail.Count do
  begin

   lTOB      := Detail[i] ;
   lRecError := TTOBEcriture(lTOB).isValid ;

   if ( lRecError.RC_Error <> RC_PASERREUR ) and ( lRecError.RC_Error <> RC_QUALIFPIECEMS ) then
    begin
     NotifyError(lRecError.RC_Error,intToStr(i+1) ) ;
     exit ;
    end; // if

//   lTOB.PutValue('E_IO','X') ;

   if _AvecAna(lTOB) then _CSuppAna(lTOB ) ;

   inc(i) ;

  end; // while

 {$IFDEF TT}
  AddEvenement('Numerote');
{$ENDIF}

 Numerote ;

 ExecuteSQL('DELETE FROM ECRITURE ' + FStSQLFiltre + ' AND E_QUALIFPIECE="V" and E_CONTROLEUR="' + CGetCleVrac + '"' ) ;
 SetAllModifie(true) ;
 InsertDB(nil) ;

 FTOBAna.ClearDetail ;

end;

procedure TZSaisieVrac.RemiseAZeroControleur ;
begin
 ExecuteSQL('UPDATE ECRITURE SET E_CONTROLEUR="" WHERE AND E_QUALIFPIECE="V" AND E_CONTROLEUR="' + CGetCleVrac + '"'  ) ;
end;


destructor TZSaisieVrac.Destroy;
begin

 try

  if assigned(FListeASupprimer) then FListeASupprimer.Free;
  if assigned(FTOBAna)          then FTOBAna.Free;

 finally
  FListeASupprimer              :=nil ;
  FTOBAna                       :=nil ;
 end;

 inherited;

end;



{ Fin TZSaisieVrac }

function GetE ( FListe : THGrid ; Lig : integer = -1 ) :  TTOBEcriture ;
begin
 result := nil ;
 if Lig = - 1 then Lig := FListe.Row ;
 if FListe.Objects[SA_Exo,Lig] is TTOBEcriture then
  Result := TTOBEcriture(FListe.Objects[SA_Exo,Lig]) ;
end;

procedure SetE ( FListe : THGrid ; O : TTOBEcriture ; Lig : integer = -1 ) ;
begin
 if Lig=-1 then Lig:=FListe.Row ;
 FListe.Objects[SA_Exo,Lig] :=O ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/07/2001
Modifié le ... : 20/07/2001
Description .. : Poinr d'entre de la saisie de en vrac
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_SaisieVrac;
begin

 if not ExJaiLeDroitConcept(TConcept(ccSaisEcritures), TRUE) then
  exit;

 if not CExoRefOuvert(true) then
  exit;

 if not VH^.ZACTIVEPFU then
  begin // l'option activé les nouvelles saisie n'est pas active
   PGIInfo('Vous ne pouvez pas accéder à cette fonction',cStTexteTitreFenetre);
   Exit;
  end ;

 if VH^.EtablisDefaut = '' then
  begin
   PGIINfo('Aucun etablisssemnt',cStTexteTitreFenetre);
   Exit ;
  end ;

 if ( VH^.DateCloturePer > 0 ) and ( FinDeMois(VH^.Entree.Fin) <= VH^.DateCloturePer ) then
  begin
   PGIInfo(Format('Toutes les périodes sont clôturées entre le %s et le %s', [DateToStr(VH^.Entree.Deb), DateToStr(VH^.Entree.Fin)]), cStTexteTitreFenetre);
   exit;
  end ;  

 AGLLanceFiche('CP','CPSAISIEVRAC','','','') ;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... :   /  /
Description .. : Rem : on rentre directement en création
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEVRAC.OnArgument (S : String ) ;
begin
 Inherited ;
 {$IFDEF TT}
 TheLog:=TStringList.Create ;
 TheLog.SaveToFile('c:\SaisieVrac.txt') ;
{$ENDIF}

  // on sauvegarde la valeur courante
 FExceptionEvent         := Application.OnException ;
 // on branche le nouveau gestionnaire d'exception
// Application.OnException := GestionException ; FB 13915 ( voir fiche )

 if not GetControlTOF then exit ;
 if not CreateControl then exit ;
 if not InitControl   then exit ;
 if not AssignEvent   then exit ;

 if Transactions(FZSaisieVrac.RemiseAZeroControleur,1) <> oeOK then
  begin
   ShowMessage( 'bug' + V_PGI.LastSQLError );
  end;

 ChercheClick(false) ;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 01/12/2004
Description .. : Création des objets de la TOF
Suite ........ : - LG - 01/12/2004 - FB 15031 - corefction des filtres
Mots clefs ... : 
*****************************************************************}
function TOF_CPSAISIEVRAC.CreateControl : boolean;
 var
  lComposants : TControlFiltre ;
begin
 {$IFDEF TT}
  AddEvenement('CreateControl' );
 {$ENDIF}
 // initialisation de l'objet de recherche ( avant le readFolio qui l'utilise )
 FInfoEcr              := TInfoEcriture.Create ;
 FRechCompte           := TRechercheGrille.Create(FInfoEcr) ;
  // Objet d'affichage des messages
 FMessCompta           := TMessageCompta.Create(Ecran.Caption) ;
 FZSaisieVrac          := TZSaisieVrac.Create('',nil,-1) ;
 FZSaisieVrac.Info     := FInfoEcr ;
 FZSaisieVrac.OnError  := OnErrorTOB ;
 FZSaisieVrac.Initialize ;

 FZIntegration         := TZIntegrationVrac.Create(FInfoEcr) ;
 FZIntegration.OnError := OnErrorTOB ;
 FZIntegration.Ecran   := Ecran ;
 FZIntegration.Initialize ;

 FZGuide               := TZGuide.Create(FInfoEcr);

 FZNature              := TZNature.Create ;
 lComposants.PopupF    := POPF ;
 lComposants.Filtres   := FFILTRES ;
 lComposants.Filtre    := TToolbarButton97(GetControl('BFILTRE'));
 lComposants.PageCtrl  := FPageControl  ;
 FObjFiltre            := TObjFiltre.Create(lComposants, 'SAISIEVRAC') ;
 FlashGuide            := TFlashingLabel.Create(Ecran) ;
 result                := true;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Destruction des objets
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEVRAC.OnClose ;
begin

 if Transactions(FZSaisieVrac.RemiseAZeroControleur,1) <> oeOK then
  begin
   ShowMessage( 'bug' + V_PGI.LastSQLError );
  end;

 //PurgePopup(POPF);
 PurgePopup(POPZOOM);
 PurgePopup(POPS);

 // on rebranche l'evenement par defaut
 Application.OnException := FExceptionEvent;

 try

  if assigned(FRechCompte)       then FRechCompte.Free ;
  if assigned(FMessCompta)       then FMessCompta.Free;
  if assigned(FInfoEcr)          then FInfoEcr.Free;
  if assigned(FZSaisieVrac)      then FZSaisieVrac.Free;
  if assigned(FZNature)          then FZNature.Free;
  if assigned(FZIntegration)     then FZIntegration.Free;
  if assigned(FZGuide)           then FZGuide.Free ;
  FreeAndNil(FObjFiltre) ; 
  FreeAndNil(FlashGuide) ;

 finally
  FRechCompte         := nil ;
  FMessCompta         := nil ;
  FInfoEcr            := nil ;
  FZSaisieVrac        := nil ;
  FZNature            := nil ;
  FZIntegration       := nil ;
  FZGuide             := nil ;
 end; //try

 Inherited ;

 {$IFDEF TT}
  FreeAndNil(TheLog) ;
{$ENDIF}

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 08/11/2001
Modifié le ... :   /  /
Description .. : Gestion des exceptions globales pour la fiche CRELBQE
Mots clefs ... :
*****************************************************************}
{procedure TOF_CPSAISIEVRAC.GestionException ( Sender : TObject ; E : Exception);
begin
 MessageAlerte( 'Quitter la saisie en vrac ! ' + #13#10 + #13#10 + TheLastError + #13#10 + #13#10 +
               E.Message );
 FListe.SynEnabled := true ;
 FListe.EndUpdate ;
 FListe.Refresh ;
 FListe.Enabled            := false ;
 FPageControl.Enabled      := false ;
 BInsert.Enabled           := false ;
 BCherche.Enabled          := false ;
 BDelete.Enabled           := false ;
 BValider.Enabled          := false ;
 BImprimer.Enabled         := false ;
 BDefaire.Enabled          := false ;
 BAgrandir.Enabled         := false ;
 BMenuZoom.Enabled         := false ;
 BIntegre.Enabled          := false ;
 BSelectionner.Enabled     := false ;
 BDeSelectionner.Enabled   := false ;
end;
}

procedure TOF_CPSAISIEVRAC.ChargeParam;
begin

 {$IFDEF TT}
 AddEvenement('ChargeParam' );
 {$ENDIF}

 FBoAfficherEtablis := not CUnSeulEtablis ;
 FBoAfficherNature  := not GetParamSocSecur('SO_NATSAISIEVRAC',False) ;
 CDateParDefautPourSaisie(FDtDateDeb,FDtDateFin) ; // init des dates de saisie

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 20/07/2001
Description .. : Recupération des contrôles de la fiche
Mots clefs ... :
*****************************************************************}
function TOF_CPSAISIEVRAC.GetControlTOF: boolean;
begin

 FPageControl            := TPageControl(GetControl('FE__PAGECONTROL_CRITERE'));
 TbsStandard             := TTabSheet(GetControl('FE__TABSHEET_STANDARD'));
 FListe                  := THGrid(GetControl('GS'));
 BInsert                 := TToolbarButton97(GetControl('BInsert')) ;
 BDelete                 := TToolbarButton97(GetControl('BDelete')) ;
 BCherche                := TToolbarButton97(GetControl('BCherche')) ;
 BValider                := TToolbarButton97(GetControl('BValider')) ;
 BImprimer               := TToolbarButton97(GetControl('BImprimer')) ;
 BDefaire                := TToolbarButton97(GetControl('BDefaire')) ;
 BAgrandir               := TToolbarButton97(GetControl('BAgrandir')) ;
 BMenuZoom               := TToolbarButton97(GetControl('BMenuZoom')) ;
 BIntegre                := TToolbarButton97(GetControl('BIntegre')) ;
 E_JOURNAL               := THCritMaskEdit(GetControl('E_JOURNAL'));
 E_ETABLISSEMENT         := THCritMaskEdit(GetControl('E_ETABLISSEMENT'));
 E_JOURNAL               := THCritMaskEdit(GetControl('E_JOURNAL'));
 E_UTILISATEUR           := THCritMaskEdit(GetControl('E_UTILISATEUR'));
 E_DATECOMPTABLE         := THCritMaskEdit(GetControl('E_DATECOMPTABLE'));
 E_DATECOMPTABLE_        := THCritMaskEdit(GetControl('E_DATECOMPTABLE_'));
 SOLDEDEBIT              := THLabel(GetControl('SOLDEDEBIT'));
 SOLDECREDIT             := THLabel(GetControl('SOLDECREDIT'));
 SOLDE                   := THLabel(GetControl('SOLDE'));
 LSOLDE                  := THLabel(GetControl('LSOLDE'));
 CSOLDE                  := TBevel(GetControl('CSOLDE'));
 CSOLDEDEBIT             := TBevel(GetControl('CSOLDEDEBIT'));
 CSOLDECREDIT            := TBevel(GetControl('CSOLDECREDIT'));
 BDefaire                := TToolbarButton97(GetControl('BDefaire')) ;
 BAgrandir               := TToolbarButton97(GetControl('BAgrandir')) ;
 G_LIBELLE               := THLabel(GetControl('G_LIBELLE'));
 T_LIBELLE               := THLabel(GetControl('T_LIBELLE'));

 POPF                    := TPopupMenu(GetControl('POPF'));
 POPS                    := TPopupMenu(GetControl('POPS'));
 POPZOOM                 := TPopupMenu(GetControl('POPZOOM'));
 FFiltres                := THValComboBox(GetControl('FFiltres'));
 BDeSelectionner         := TToolbarButton97(GetControl('BDETAG')) ;
 BSelectionner           := TToolbarButton97(GetControl('BTAG')) ;
 HSystemMenu1            := TFVierge(Ecran).HMTrad;

 result := true;

end;


procedure TOF_CPSAISIEVRAC.InitGrid ;
begin
 // parametrage de la grille de saisie des relevées
 FListe.ColCount   := 11 ;
 FListe.ScrollBars := ssVertical ;
 if ( not FBoAfficherEtablis ) then FListe.ColCount := FListe.ColCount - 1 ;
 if ( not FBoAfficherNature ) then FListe.ColCount := FListe.ColCount - 1 ;

 if FListe.ColCount = 11 then
  begin
   SA_Numl     := 0 ;
   SA_DateC    := 1 ;
   SA_Jal      := 2 ;
   SA_Etabl    := 3 ;
   SA_Nat      := 4 ;
   SA_Gen      := 5 ;
   SA_Aux      := 6 ;
   SA_Lib      := 8 ;
   SA_RefI     := 7 ;
   SA_Debit    := 9 ;
   SA_Credit   := 10 ;
  end; // if

 if ( FListe.ColCount = 10 ) and ( not FBoAfficherEtablis ) then
  begin
   SA_Numl     := 0 ;
   SA_DateC    := 1 ;
   SA_Jal      := 2 ;
   SA_Etabl    := -1 ;
   SA_Nat      := 3 ;
   SA_Gen      := 4 ;
   SA_Aux      := 5 ;
   SA_Lib      := 7 ;
   SA_RefI     := 6 ;
   SA_Debit    := 8 ;
   SA_Credit   := 9 ;
  end; // if

 if ( FListe.ColCount = 10 ) and ( not FBoAfficherNature ) then
  begin
   SA_Numl     := 0 ;
   SA_DateC    := 1 ;
   SA_Jal      := 2 ;
   SA_Etabl    := 3 ;
   SA_Nat      := -1 ;
   SA_Gen      := 4 ;
   SA_Aux      := 5 ;
   SA_Lib      := 7 ;
   SA_RefI     := 6 ;
   SA_Debit    := 8 ;
   SA_Credit   := 9 ;
  end; // if

 if ( FListe.ColCount = 9 ) then
  begin
   SA_Numl     := 0 ;
   SA_DateC    := 1 ;
   SA_Jal      := 2 ;
   SA_Etabl    := -1 ;
   SA_Nat      := -1 ;
   SA_Gen      := 3 ;
   SA_Aux      := 4 ;
   SA_Lib      := 6 ;
   SA_RefI     := 5 ;
   SA_Debit    := 7 ;
   SA_Credit   := 8 ;
  end; // if
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/06/2001
Modifié le ... : 03/02/2006
Description .. : Initialisation des contrôles
Suite ........ : - 09/08/2004 - LG - suppression de l'option de trie de la 
Suite ........ : grille de saisie
Suite ........ : - 03/02/2006 - LG - reactivation de la gestion du formatage 
Suite ........ : de la grille
Mots clefs ... : 
*****************************************************************}
function TOF_CPSAISIEVRAC.InitControl : boolean;
var
  R : TRect;
begin
 {$IFDEF TT}
  AddEvenement('InitControl' );
  FListe.Options := FListe.Options + [goColSizing] ;
 {$ENDIF}
 FListe.SortEnabled                               := false ;   
 BImprimer.Visible                                := false ;
 Ecran.Caption                                    := ' Saisie en vrac' ;
 // suppression des boutons impossible a supprimer
 TToolbarButton97(GetControl('bDefaire')).Visible := false ;
 FPageControl.Pages[1].TabVisible                 := false ;
 FPageControl.Pages[2].TabVisible                 := false ;
  // recup du titre de la fenetre
 FStTitre                                         := Ecran.Caption ;
 Statut                                           := taConsult ;
 FBoGrilleModif                                   := false ;
 FBoGrilleVide                                    := true ;
 FBoValidationEnCours                             := false ;
 FBoAffMessErr                                    := true ;
 // init des num de colonnes - constantes def dans UtilSais
//  _initSA_SaisieVrac ;
 FRechCompte.OnError                              := OnErrorRech ;
 // chargement des parametres
 ChargeParam;
 {$IFDEF TT}
  AddEvenement('InitControl' );
 {$ENDIF}
 // chargement des info sur la devise pivot
 FInfoEcr.Devise.Load([v_pgi.DevisePivot]) ;

 InitGrid ;

 FRechCompte.COL_GEN                              := SA_GEN ;
 FRechCompte.COL_AUX                              := SA_AUX ;

 FListe.ColWidths[SA_Numl]                        := 20;
 FListe.ColWidths[SA_DateC]                       := 45;
 FListe.ColWidths[SA_Jal]                         := 30;
 FListe.ColWidths[SA_Gen]                         := 75;
 FListe.ColWidths[SA_Aux]                         := 80;
 FListe.ColWidths[SA_Lib]                         := 80;
 FListe.ColWidths[SA_RefI]                        := 70;  
 FListe.ColWidths[SA_Debit]                       := 50;
 FListe.ColWidths[SA_Credit]                      := 50;
 if FBoAfficherEtablis then FListe.ColWidths[SA_Etabl] := 30 ;
 if FBoAfficherNature then  FListe.ColWidths[SA_Nat]   := 30 ;

 FListe.Cells[SA_Numl,0]                          := TraduireMemoire('N°');
 FListe.Cells[SA_DateC,0]                         := TraduireMemoire('Date');
 FListe.Cells[SA_Jal,0]                           := TraduireMemoire('Jal');
 if FBoAfficherEtablis then FListe.Cells[SA_Etabl,0] := TraduireMemoire('Ets');
 if FBoAfficherNature then   FListe.Cells[SA_Nat,0]  := TraduireMemoire('Nat');
 FListe.Cells[SA_Gen,0]                           := TraduireMemoire('Général');
 FListe.Cells[SA_Aux,0]                           := TraduireMemoire('Auxiliaire');
 FListe.Cells[SA_Lib,0]                           := TraduireMemoire('Libellé');
 FListe.Cells[SA_RefI,0]                          := TraduireMemoire('Référence');
 FListe.Cells[SA_Debit,0]                         := TraduireMemoire('Débit');
 FListe.Cells[SA_Credit,0]                        := TraduireMemoire('Crédit');

 // limitation de la longeur des champs
 FListe.ColLengths[SA_Lib]                        := 35;
 FListe.ColLengths[SA_RefI]                       := 17;
 FListe.ColLengths[SA_Gen]                        := 30 ;
 FListe.ColLengths[SA_Aux]                        := VH^.Cpta[fbAux].Lg ;
 if FBoAfficherNature then
 FListe.ColLengths[SA_Nat]                        := 3;
 if FBoAfficherEtablis then
 FListe.ColLengths[SA_Etabl]                      := 3;
 if FBoAfficherNature then
 FListe.ColLengths[SA_Jal]                        := 3;

 // formatage de la cellule Date
 FListe.ColAligns[SA_Numl]                        := taCenter ;
 FListe.ColTypes[SA_DateC]                        := 'D';
 FListe.ColFormats[SA_DateC]                      := ShortDateFormat;
 FListe.ColAligns[SA_Jal]                         := taCenter ;
 if FBoAfficherEtablis then FListe.ColAligns[SA_Etabl] := taCenter ;
 FListe.ColTypes[SA_Debit]                        := 'R';
 FListe.ColAligns[SA_Debit]                       := taRightJustify ;
 FListe.ColFormats[SA_Debit]                      := '#,##0.00';
 FListe.ColTypes[SA_Credit]                       := 'R';
 FListe.ColFormats[SA_Credit]                     := '#,##0.00';
 FListe.ColAligns[SA_Credit]                      := taRightJustify ; 

 // redimensionnement de la grille
 HSystemMenu1.ResizeGridColumns(FListe);

 R                                                := FListe.CellRect(SA_Debit ,0) ;
 SOLDEDEBIT.Left                                  := R.Left - 7 ;
 SOLDEDEBIT.Width                                 := R.Right - R.Left ;
 CSOLDEDEBIT.Left                                 := R.Left - 2 ;
 CSOLDEDEBIT.Width                                := R.Right - R.Left ;

 R                                                := FListe.CellRect(SA_Credit ,0) ;
 SOLDECREDIT.Left                                 := R.Left - 7 ;
 SOLDECREDIT.Width                                := R.Right - R.Left ;
 CSOLDECREDIT.Left                                := R.Left - 2 ;
 CSOLDECREDIT.Width                               := R.Right - R.Left ;
 SOLDE.Left                                       := R.Left - 7 ;
 SOLDE.Width                                      := R.Right - R.Left ;
 CSOLDE.Left                                      := R.Left - 2 ;
 CSOLDE.Width                                     := R.Right - R.Left ;
 LSOLDE.Left                                      := CSOLDEDEBIT.Left + 3 ;

 G_LIBELLE.Font.Color                             := clBlack ;
 G_LIBELLE.Font.Style                             := G_LIBELLE.Font.Style-[fsBold] ;
 G_LIBELLE.Caption                                := '' ;
 T_LIBELLE.Caption                                := '' ;
 FlashGuide.Parent                                := T_LIBELLE.Parent ;
 FlashGuide.Left                                  := T_LIBELLE.Left ;
 FlashGuide.Height                                := T_LIBELLE.Height ;
 FlashGuide.Top                                   := T_LIBELLE.Top + 18 ;
 FlashGuide.Caption                               := 'GUIDE' ;
 FlashGuide.Transparent                           := true ;
 FlashGuide.Color                                 := clRed ;
 FlashGuide.Font.Style                            := [fsBold] ;                                                                  
 FlashGuide.Visible                               := false ;

 Ecran.PopupMenu                                  := POPS;

 FPageControl.ActivePageIndex                     := 0;

 FInfoEcr.Etabliss.LoadAll ; // chargement de tous les etablissement
 FZNature.LoadAll ; // chargement des natures

 E_DATECOMPTABLE.Text                            := DateToStr(FDtDateDeb) ;
 E_DATECOMPTABLE_.Text                           := DateToStr(FDtDateFin) ;
 FInLigneCourante                                := -1 ;
 FInIndexComptaAuto                              := 1 ;
 FStLastJal                                      := '' ;
 FInNumGroupEcr                                  := 0 ;
 BoGuideRun                                      := false ;
 PositionneEtabUser(E_ETABLISSEMENT) ;

 FObjFiltre.Charger ;

 result                                          := true;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 20/06/2001
Modifié le ... : 20/07/2001
Description .. : Affectation des evenements au contrôle
Mots clefs ... :
*****************************************************************}
function TOF_CPSAISIEVRAC.AssignEvent : boolean;
begin
 FListe.OnCellEnter           := FListeCellEnter;    // affichage du Bouton sur les col. Imputation et Date
 FListe.OnElipsisClick        := FListeElipsisClick; // Click sur Bouton dans les col. Imputation et Date
 FListe.PostDrawCell          := FListePostDrawCell;
 FListe.OnRowEnter            := FListeRowEnter;
 FListe.OnRowExit             := FListeRowExit;
 FListe.OnCellEnter           := FListeCellEnter;
 FListe.OnCellExit            := FListeCellExit;
 FListe.OnEnter               := FListeEnter;
 FListe.OnKeyPress            := FListeKeyPress;
 FListe.OnKeyDown             := FListeKeyDown;
 Ecran.OnKeyDown              := FormKeyDown ;
 BValider.OnClick             := BValiderClick ;
 BCherche.OnClick             := BChercheClick ;
 BIntegre.OnClick             := BIntegreClick ;
 E_JOURNAL.OnElipsisClick     := E_JOURNALElipsisClick;
 E_JOURNAL.OnExit             := E_JOURNALExit ;
 BInsert.OnClick              := BInsertClick ;
 BDelete.OnClick              := BDeleteClick ;
 BAgrandir.OnClick            := BAgrandirClick;
 Ecran.OnCloseQuery           := FormCloseQuery ;
 POPS.OnPopup                 := POPSPopup;
 result                       := true;
end;


{procedure TOF_CPSAISIEVRAC.BMenuZoomMouseEnter(Sender: TObject);
begin
 PopZoom97(BMenuZoom,POPZOOM);
end; }


procedure SetGridSep(g: thgrid ; ACol, ARow : Integer ; Canvas : TCanvas; bHaut : Boolean) ;
var R : TRect ;
begin
Canvas.Brush.Color := clRed ;
Canvas.Brush.Style := bsSolid ;
Canvas.Pen.Color   := clRed ;
Canvas.Pen.Mode    := pmCopy ;
Canvas.Pen.Style   := psSolid ;
Canvas.Pen.Width   := 1 ;
R:=G.CellRect(ACol, ARow) ;
if bHaut then begin Canvas.MoveTo(R.Left, R.Top) ; Canvas.LineTo(R.Right+1, R.Top) end
         else begin Canvas.MoveTo(R.Left, R.Bottom-1) ; Canvas.LineTo(R.Right+1, R.Bottom-1) end ;
end ;


function TOF_CPSAISIEVRAC.CalculSeparateur ( ACol, ARow : Integer ; var vBoHaut : boolean ; vBo : boolean = false  ) : boolean ;
var
 lInNum            : integer ;
 lInNumPrec        : integer ;
 lInNumSuiv        : integer ;
 lTOB              : TOB ;
begin

 lInNum        := -1 ;
 lInNumPrec    := -1 ;
 lInNumSuiv    := -1 ;
 vBoHaut       := false ;
 lTOB          := GetE(Fliste,ARow) ;
 result        := false ;

 if lTOB <> nil then
 lInNum := lTOB.GetValue('E_NUMGROUPEECR') ;
 lTOB := GetE(Fliste,ARow-1) ;
 if lTOB <> nil then
 lInNumPrec := lTOB.GetValue('E_NUMGROUPEECR') ;
 lTOB := GetE(Fliste,ARow+1) ;
 if lTOB <> nil then
 lInNumSuiv := lTOB.GetValue('E_NUMGROUPEECR') ;

 if vBo then
  begin
   if ( ACol > 0 ) and ( lInNum <> lInNumPrec )  then result := true ;
  end
   else
    begin
     if ( ACol > 0 ) and ( lInNum <> lInNumSuiv ) and (FListe.Row <> ARow) then result := true ;
     if ( ACol > 0)  and ( lInNum <> lInNumPrec ) and ( FListe.Row = ARow - 1) then begin result := true ; vBoHaut := true ; end ;
    end ; 

end ;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 28/06/2001
Modifié le ... : 20/07/2001
Description .. : Grise les cellules débit/crédit en fonction de leur état
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FListePostDrawCell( ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState );
var
 lBoGrise          : boolean ;
 lBoSeparateur     : boolean ;
 lBoHaut           : boolean ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 if ( ARow < FListe.FixedRows ) or ( ARow > ( FListe.RowCount - 1 ) ) then Exit ;

 try

 lBoSeparateur := CalculSeparateur ( ACol, ARow , lBoHaut ) ;

 // gestion des cases debit et crédit - une seul des deux doit être renseignée
 lBoGrise := ( ( ACol = SA_Debit )  and ( FListe.Cells[SA_Credit, ARow] <> '' ) and  ( ARow > 0 ) ) or
             ( ( ACol = SA_Credit ) and ( FListe.Cells[SA_Debit, ARow]  <> '' ) and  ( ARow > 0 ) ) or
             ( ( ACol = SA_Aux )    and ( trim(FListe.Cells[SA_Aux, ARow] ) = ''    ) );

 if lBoGrise then
  begin
   FListe.PostDrawCell  := nil ; // on debranche l'évènement lors du dessin de la grille
   SetGridGrise(ACol, ARow, FListe) ;
   FListe.PostDrawCell  := FListePostDrawCell ;
  end;

 if lBoSeparateur then
  begin
   FListe.PostDrawCell  := nil ; // on debranche l'évènement lors du dessin de la grille
   CSetGridSep(ACol, ARow, FListe, Canvas , lBoHaut ) ;
   FListe.PostDrawCell  := FListePostDrawCell ;
  end;

 except
  on E : exception do
   begin
     FListe.PostDrawCell := nil;
     PGIBox( 'Problème lors de l''affichage des données !' + #13#10 +
             E.Message ,
             cStTexteTitreFenetre);
     V_PGI.IoError := oeSaisie;
   end; // on
 end;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 13/01/2004
Modifié le ... : 16/09/2004
Description .. : -13/01/2004 - FB 13095 - message sur echap
Suite ........ : - 13/05/2004 - FB 12892 - en validation oin ne peut pas 
Suite ........ : sortir
Suite ........ : - 16/09/2004 - FB 14613 - LG - des que l'on appuie sur echap on sort 
Suite ........ : ds la saisie
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin

 if FBoValidationEnCours then
  begin
   CanClose := false ;
   exit ;
  end ;

 if BoGuideRun then
  begin
   BoGuideRun := PGIAskCancel(FMessCompta.GetMessage(RC_GUIDESTOP),cStTexteTitreFenetre) <> mrYes ;
   if not BoGuideRun then
    FinGuide ;
   CanClose   := false ;
   Exit ;
  end ;


 if CanClose and FBoGrilleModif then
  begin
   if ( PGIAskCancel(FMessCompta.GetMessage(RC_ABANDON),cStTexteTitreFenetre) = mrYes ) then
    begin
     CanClose := true ;
     TFVierge(Ecran).FormCloseQuery(Sender,CanClose);
     exit ;
    end
     else
      begin
       CanClose := false ;
       exit ;
      end ;
  end;

  TFVierge(Ecran).FormCloseQuery(Sender,CanClose); 

end;


procedure TOF_CPSAISIEVRAC.FormKeyDownForms( Sender : TObject ; var Key : Word; Shift : TShiftState);
begin

 case Key of
  VK_F5  : begin {$IFDEF TT} AddEvenement('FormKeyDownForms : KEY_LOOKUP'); {$ENDIF} Key :=0 ; CVK_F5OnEdit(Ecran) ;  end ; // if
 { VK_END      : if (Shift = [ssCtrl] ) and (Statut=taConsult) then
                 begin // on ne place a la fin quand consultation
                  Key:=0 ;
                  SetGridOptions(nbLignes) ;
                  GS.Row:=nbLignes ;
                  GS.col:=SF_GEN ;
                 end;  }
 end; // case

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 09/10/2003
Modifié le ... : 27/07/2005
Description .. : - 09/10/2003 - FB 12887 - entre sur la dernier ligne une
Suite ........ : nouvelle ligne, mais en insertion on provoque un VK_TAB
Suite ........ : - 17/10/2003 - FB 12911 - F12 place le curseur ds le multi
Suite ........ : critere
Suite ........ : - 10/05/2004 - FB 12903 - gestion de la touche fin
Suite ........ : - 10/05/2004 - FB 12895 - on supprime la ligne uniquement
Suite ........ : si le compte general n'est pas rensiugné
Suite ........ : - 13/05/2004 - FB 12892 - si on est en validation on sort
Suite ........ : - 18/06/2004 - FB 12903 - solde credit on vas au debit etc 
Suite ........ : etc
Suite ........ : - 18/06/2004 - FB 13174 - correction de la fermeture qd 
Suite ........ : plus 
Suite ........ : aucune ecriture ds la grille
Suite ........ : - 21/12/2004 - FB 12903 - correction de la gestion de la 
Suite ........ : touche fin; Qd un montent est present on se positionne 
Suite ........ : dessus
Suite ........ : - LG - 27/07/2005 - FB 16324 - supression de l'ellispis qd on 
Suite ........ : fait vk_end
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FormKeyDownFliste( Sender : TObject ; var Key : Word; Shift : TShiftState);
var
 Vide         : Boolean ;
 ACol,ARow    : integer ;
 lCancel : boolean ;
begin

 Vide        := ( Shift = [] ) ;

 if FBoValidationEnCours then exit ;

 case Key of
  VK_RETURN  : if (Vide) then
                begin {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_RETURN'); {$ENDIF}
                 FRechCompte.InKey := Key ;
                 if ( FListe.Col = SA_Credit ) and ( FListe.Row = FListe.RowCount -1 ) then
                  begin
                    Key := 0 ; NextRow ;
                  end
                   else
                    Key := VK_TAB ;
                end ;
  VK_F5 : begin {$IFDEF TT} AddEvenement('FormKeyDown : KEY_LOOKUP'); {$ENDIF} Key :=0 ; FListeElipsisClick(nil) ; end ; // if
  VK_TAB     : if Vide and not BoGuideRun then
               begin
                {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_TAB'); {$ENDIF}
                if ( FListe.Row = FListe.RowCount - 1 ) and ( FListe.Col = FListe.Colcount - 1 ) then
                  begin
                   key := 0 ;
                   NextRow ;
                  end;
               end;  // if
  VK_DOWN   : if ( FListe.Row = FListe.RowCount - 1 )  then
               begin
                {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_DOWN'); {$ENDIF}
                Key := 0 ;
                NextRow ;
               end;
 { VK_UP     :  begin
               FBoAffMessErr := false ;
               if ( Statut <> taConsult ) and not ( LigneCorrecte ) and ( FListe.Row = FListe.RowCount - 1 ) then
                begin

                 if not FBoValidationEnCours then
                  begin
                   Key := 0 ;
                   if ( FListe.Cells[SA_Gen,FListe.Row] = '' ) then
                    FListeDeleteRow ;
                  end ;
                end ; // if
                FBoAffMessErr := true ;
                end; // case  }
  VK_INSERT : if (Vide) then
               begin
                {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_INSERT'); {$ENDIF}
                Key := 0 ;
                if not BoGuideRun then CreateRow(FListe.Row,true) ;
               end; // if
  VK_DELETE : if  Shift=[ssCtrl] then
               begin
                {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_DELETE'); {$ENDIF}
                Key := 0 ;
                FListeDeleteRow ;
               end ;
  VK_SPACE : if Vide and (FListe.Col=SA_NAT) then
               begin
                {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_DELETE'); {$ENDIF}
                Key := 0 ;
                NextNature ;
               end ;
 VK_F2     : if Shift=[ssAlt] then
              begin
                AppelCalculatrice ;
              end ;
 VK_F10    : if Vide and not BoGuideRun then
              begin
               {$IFDEF TT}   DelEvenement ; {$ENDIF}
               {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_F10'); {$ENDIF}
               Key := 0 ;
               if not FBoValidationEnCours then
                begin
                 if ( Statut = taModif) and not LigneCorrecte then exit ;
                 Validation ;
                 if FBoGrilleVide then
                  begin
                   {$IFDEF EAGLCLIENT}
                     Ecran.Close ;
                    if IsInside(Ecran) then CloseInsidePanel(Ecran) ;
                   {$ELSE}
                    if not IsInside(Ecran) then Ecran.Close
                     else
                      if PrepareInside then FMenuG.AfficheSoc(True,TRUE) ; //CloseInsidePanel(Self) ;
                   {$ENDIF}
                  end ;
                end ;
              end;
 VK_F6    : if Vide then
              begin
               {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_F6'); {$ENDIF}
               Key := 0 ;
               if FListe.Cells[SA_Gen,FListe.Row] <> '' then SoldeFolio
                else
                 begin
                  GetCompteAuto(1) ;
                  if FListe.Cells[SA_Gen,FListe.Row] <> '' then SoldeFolio ;
                 // FinSaisie ;
                 end ;
              end ;
 VK_F7     : if Vide then
              begin
               {$IFDEF TT} AddEvenement('FormKeyDownFliste : VK_F7'); {$ENDIF}
               Key := 0 ;
               VK_F7FListe ;
              end;
 VK_F11     : begin
               Key:=0;
               TPopUpMenu(GetControl('POPS')).Popup(Mouse.CursorPos.x,Mouse.CursorPos.y);
              end; // VK_F11
 VK_F4 : begin
                if FListe.Row = FInLigneCourante then
                 begin // sur la meme ligne on increment le compte auto
                  Inc(FInIndexComptaAuto) ;
                  if FInIndexComptaAuto > FInfoEcr.Journal.NombreDeCompteAuto then FInIndexComptaAuto:=1 ;
                  GetCompteAuto(FInIndexComptaAuto) ;
                  Key := 0 ;
                 end // if
                  else
                   begin
                    FInLigneCourante := FListe.Row ;
                    GetCompteAuto(1) ;
                    Key := 0 ;
                   end;
              end ;
VK_F3 : begin
               if FListe.Row = FInLigneCourante then
                begin // sur la meme ligne on increment le compte auto
                 Dec(FInIndexComptaAuto) ;
                 if FInIndexComptaAuto < 1 then FInIndexComptaAuto := FInfoEcr.Journal.NombreDeCompteAuto ;
                 GetCompteAuto(FInIndexComptaAuto) ;
                 Key := 0 ;
               end
                else
                 begin // sur une nouvelle ligne on prend le premier
                  FInLigneCourante := FListe.Row ;
                  GetCompteAuto(1) ;
                  Key := 0 ;
                 end;
              end ;
 VK_F12     : if Vide and ( FPageControl.CanFocus ) then begin Key := 0 ; FPageControl.SetFocus ; end;
 VK_END     : if Vide and not BoGuideRun then
              begin
             //  if ( FListe.Cells[SA_Credit,FListe.Row] <> '' ) then
              ACol:=FListe.Col ; ARow:=FListe.Row ;
              FListeCellExit(nil,ACol,ARow,lCancel) ; if lCancel then exit ;
              FListe.ElipsisButton := false ;
              if  FListe.Cells[SA_Credit, ARow] <> '' then
               FListe.Col := SA_Credit
                else
                 if  FListe.Cells[SA_Debit, ARow] <> '' then
                  FListe.Col := SA_Debit
                   else
                    if FRdSolde > 0 then
                     FListe.Col := SA_Credit
                      else
                       FListe.Col := SA_Debit ;
               Key := 0 ;
              end
               else
                if ( Shift = [ssCtrl] ) and ( Statut = taConsult ) then
                 begin
                  Key:=0 ;
                  FListe.Row := FListe.RowCount - 1 ;
                  Fliste.col := SA_DateC ;
                  AfficheInfoLigne ;
                 end ;
{CTRL+A} 65 : if Shift=[ssCtrl] then begin Key := 0 ; GereAna(FListe.Row,true) ; end ;
{ALT+G}  71 : if Shift=[ssAlt]  then begin Key:=0 ; GuiClick ; end
 end; // case
end;



procedure TOF_CPSAISIEVRAC.FormKeyDown( Sender : TObject ; var Key : Word; Shift : TShiftState);
begin

 case Key of
  VK_F9       : if not BoGuideRun then
                 begin
                  Key := 0 ;
                  ChercheClick ;
                  exit ;
                 end ;
 end; // case

 if Screen.ActiveControl = FListe then
  FormKeyDownFListe( Sender , Key , Shift )
   else
    FormKeyDownForms( Sender , Key , Shift ) ;
end;

procedure TOF_CPSAISIEVRAC.FListeKeyPress (Sender : TObject ; var Key : Char);
begin
 if not FListe.SynEnabled then
  Key:=#0
   else
    begin
     if Key = #13 then Key := #9 ;
     if (Key='+') and ((FListe.Col=SA_GEN) or (FListe.Col=SA_AUX)) then Key := #0 ;
     if (Key='+') and ((FListe.Col=SA_GEN) or (FListe.Col=SA_AUX)) then Key := #0 ;
     if (Key='+') and (FListe.Col=SA_NAT) then Key := #0 ;
     if (Key=' ') and (FListe.Col=SA_NAT) then Key := #0 ;
   end ;
 if ( Key <> #9 ) and ( Key <> #0 ) then
  begin
    Statut         := taModif ;
    FBoGrilleModif := true ;
  end ;

 if FListe.Col = SA_DateC then
  ParamDate(ecran,sender,key);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 23/10/2003
Modifié le ... :   /  /    
Description .. : LG - 23/10/2003 - FB 12965 - on passe lagrille en modif qd 
Suite ........ : on utilise la touche suppr
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
 if ( Shift = [] ) and ( key = VK_DELETE ) then
  begin
   Statut         := taModif ;
   FBoGrilleModif := true ;
  end ;
end;

procedure TOF_CPSAISIEVRAC.POPSPopup ( Sender : TObject) ;
var
 i   : integer;
 lT  : TMenuItem ;
 lRS : TShortCut ;
begin

 InitPopUp(Ecran);

 // Ajout des shortcuts au menu cree par InitPopUp
 for i := 0 to ( POPS.Items.Count - 1 ) do
  begin
   lT  := POPS.Items[i];
   lRS := 0;

   if UpperCase( lT.Name ) = 'PTAG' then
    lRS := ShortCut( ord('A') , [ssCtrl] )
     else
      if UpperCase( lT.Name ) = 'PINTEGRER' then
       lRS := ShortCut( VK_F10 , [] )
        else
         if UpperCase( lT.Name ) = 'PDELETE' then
          lRS := ShortCut( VK_DELETE , [ssCtrl] );

   if lRS <> 0 then
    lT.Caption:= lT.Caption + #9 + ShortCutToText(lRS) ;

  end; // for

end;

procedure TOF_CPSAISIEVRAC.AppelCalculatrice ;
var
 ACol,ARow : integer ;
 lChg,lCancel : boolean ;
begin
 Application.ProcessMessages ; LaCalculette(Screen.ActiveForm) ;
 ACol:=FListe.Col ; ARow:=FListe.Row ;
 FListe.Cells[FListe.Col,FListe.Row] := FListe.InplaceEditor.Text ; // la grille n'a pas encore ete remise a jour avec l'inplaceEditor et les cellExit ne se faisaient pas avec les bonnes valeurs
 FListeCellExit(nil,ACol,ARow,lCancel) ; lChg:=true ;
 if not lCancel then FListeRowExit(nil,ARow,lCancel,lChg) ;
 if not lCancel then NextRow ;
// Statut:=taModif ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 12/01/2004
Modifié le ... :   /  /    
Description .. : - 12/01/2004 - FB 13093 - rajout d'un arrondi sur le solde
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.CalculInfoSupp(vTOB,vTOBPrec,vTOBSuiv : TOB  ; vBoInsert : boolean = false ) ;
var
 lRdSolde : double ;
begin

 {$IFDEF TT}
   AddEvenement('CalculInfoSupp') ;
 {$ENDIF}

 if BoGuideRun then exit ;

 vTOB.PutValue('CRIT',_MakeCrit(vTOB) ) ;

 if vBoInsert then
  begin
   if vTOBPrec = nil then
    vTOB.PutValue('E_NUMGROUPEECR',vTOBSuiv.GetValue('E_NUMGROUPEECR'))
    else
      vTOB.PutValue('E_NUMGROUPEECR',vTOBPrec.GetValue('E_NUMGROUPEECR')) ;
  end
   else
    if vTOBPrec = nil then
     begin
      Inc(FInNumGroupEcr) ;
      vTOB.PutValue('E_NUMGROUPEECR', FInNumGroupEcr ) ;
      vTOB.PutValue('SOLDE', vTOB.GetValue('E_DEBIT') - vTOB.GetValue('E_CREDIT') );
     end
      else
       if vTOBPrec.GetValue('CRIT') = vTOB.GetValue('CRIT') then
        begin
         vTOB.PutValue('E_NUMGROUPEECR',FInNumGroupEcr) ;
         lRdSolde := Arrondi( vTOBPrec.GetValue('SOLDE') + vTOB.GetValue('E_DEBIT') - vTOB.GetValue('E_CREDIT') , V_PGI.OkDecV ) ;
         vTOB.PutValue('SOLDE', lRdSolde ) ;
         if lRdSolde = 0 then
          Inc(FInNumGroupEcr) ;
        end
         else
          begin
           Inc(FInNumGroupEcr) ;
           vTOB.PutValue('E_NUMGROUPEECR',FInNumGroupEcr) ;
           vTOB.PutValue('SOLDE', vTOB.GetValue('E_DEBIT') - vTOB.GetValue('E_CREDIT') ) ;
          end;

 {$IFDEF TT}
   if vTOB <> nil then
    FListe.Cells[SA_RefI,FListe.Row] := _MakeCrit(vTOB) ; // inttostr(vTOB.GetValue('E_NUMGROUPEECR')) ;
 {$ENDIF}

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 16/09/2004
Modifié le ... : 20/09/2005
Description .. : - LG - 16/09/2004 - FB 14590 - utilistion d'une fct de 
Suite ........ : ULibEcriture pour l'affectation des echeances
Suite ........ : - LG - 20/09/2005 - FB 16740 - les test sur le modede saisie 
Suite ........ : et la qualifpice doit etre ignoree ici ( e_qualipiece = 'V' )
Mots clefs ... : 
*****************************************************************}
function TOF_CPSAISIEVRAC.LigneCorrecte  ( ARow : integer = - 1 ) : boolean;
var
 lTOB      : TTOBEcriture ;
 ACol      : integer ;
 lBoCancel : boolean ;
 lRecError : TRecError ;
begin

 result := true ;

 if ( csDestroying in Ecran.ComponentState ) or ( Statut = taConsult ) or ( FBoGrilleVide ) then Exit ;

 {$IFDEF TT}
   AddEvenement('LigneCorrecte Ou:='+intToStr(FListe.Row)) ;
 {$ENDIF}

 result     := false ;
 lBoCancel  := false ;
 if ARow = - 1 then ARow := FListe.Row ;

 ACol       := FListe.Col ;
 if not BoGuideRun then FListeCellExit(nil,ACol,ARow,lBoCancel) ;
 if lBoCancel then exit ;

 lTOB     := RecupTOB(ARow) ;

 if lTOB = nil then
  begin  //
   MessageAlerte('gravec erreur dans la saisie en vrac' ) ;
   Exit ;
  end;

 FInfoEcr.Load(FListe.Cells[SA_Gen,ARow],FListe.Cells[SA_Aux,ARow],FListe.Cells[SA_Jal,ARow]) ;
 if FInfoEcr.Journal.inIndex = -1 then
  begin
   AfficheTitreAvecCommeInfo( FMessCompta.GetMessage( RC_JALINEXISTANT ) ) ;
   Exit ;
  end;

 FBoValidationEnCours := true ;

 try

 lTOB.PutValue( 'E_MODESAISIE' , FInfoEcr.GetString('J_MODESAISIE') ) ;
 lTOB.CompleteInfo ;
 lTOB.OnError := OnErrorTOB ;

 lRecError := lTOB.IsValid ;
 result := ( lRecError.RC_Error = RC_PASERREUR ) or ( lRecError.RC_Error = RC_QUALIFPIECEMS ) ;
 if result then
  begin
   GereAna(ARow) ;
   CalculSoldeFolio ;
   AfficheSolde ;
   if not BoGuideRun then Statut := taConsult ;
  end
   else
    begin
     FMessCompta.Execute(lRecError.RC_Error) ;
     exit ;
    end ;

 CompleteCeQueJePeux ;

 finally
  FBoValidationEnCours := false ;
  EnableButtons ;
 end;

end;

procedure TOF_CPSAISIEVRAC.FListeEnter( Sender : TObject );
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 {$IFDEF TT}
  AddEvenement('FListeEnter');
 {$ENDIF}

 if FBoGrilleVide then NextRow ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 17/09/2004
Modifié le ... : 13/12/2004
Description .. : - LG - 17/09/2004 - FB 14615 - plantage quand on cliquait 
Suite ........ : ds la case auxi a partir de la der ligne de la grille ( qui ce 
Suite ........ : supprime )
Suite ........ : - LG - 13/12/2004 - suppression des elipsis bouton en 
Suite ........ : eAGL
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FListeCellEnter(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean);
var
 lc      : integer;
 lr      : integer;
 lInSens : integer ;
 SC      : integer ;
 lStSens : string ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

{$IFDEF TT}
 AddEvenement('FListeCellEnter ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow)+' : FListe.Col:='+intToStr(FListe.Col)+' FListe.Row:='+intToStr(FListe.Row));
{$ENDIF}

 Fliste.BeginUpdate ;

 try

 if ACol > SA_Jal then
  FBoGrilleVide := false ; // on considere que la grille n'est plus vide si on est sorti de la case journal. Le LigneCorrecte sera donc executer

 lr      := FListe.Row;
 lc      := FListe.Col;
 lInSens := CGetGridSens(FListe , ACol, ARow) ;

 if _Plantage < 15 then
  begin
   Cancel := not ControleColPrecedent(ACol,ARow) ;
   if Cancel then
    begin
     inc(_Plantage) ;
     FListe.Col := ACol ;
     exit ;
    end; // if
 end;
  // gestion de l'affichage des boutons
 {$IFNDEF EAGLCLIENT}
 FListe.ElipsisButton := not ( ( lc = SA_Lib ) or ( lc = SA_Debit) or ( lc = SA_Credit ) or ( lc = SA_RefI ) ) ;
 {$ELSE}
 FListe.ElipsisButton := not ( ( lc = SA_Lib ) or ( lc = SA_Debit) or ( lc = SA_Credit ) or ( lc = SA_RefI ) or ( lc = SA_Aux ) or ( lc = SA_Gen )  ) ;
 {$ENDIF}
 //if ( lc = SA_Nat )

 if ( lc = SA_DEBIT ) and ( GetNature <> 'OD' ) and ( FListe.Cells[SA_DEBIT,lr] = '' ) and ( FListe.Cells[SA_CREDIT,lr] = '' ) then
  begin
   // QuelSens : Dernier paramètre : 1=Débit ; 2=Crédit ; 3=Mixte
   if FInfoEcr.LoadCompte(FListe.Cells[SA_Gen,FListe.Row]) then
    begin

     if lInSens = RC_GAUCHE_DROITE then
      begin
       lStSens := FInfoEcr.GetString('G_SENS') ;
       if lStSens = '' then lStSens := 'D' ;
        case lStSens[1] of
         'D' : SC := 1 ;
         'C' : SC := 2 ;
         else  SC := 3 ;
        end; // case
       if ( QuelSens(GetNature, FInfoEcr.GetString('G_NATUREGENE'), SC)) <> 1 then
        begin // on force le positionnement au credit
          ACol      := SA_CREDIT ;
          ARow      := lr ;
          Cancel    := true ;
          exit ;
        end; // if
     end;
    end;
  end;

 if ( lc = SA_CREDIT) and (FListe.Cells[SA_DEBIT,lr] = '' ) and ( FListe.Cells[SA_CREDIT,lr] = '' ) then
  begin
   if FInfoEcr.LoadCompte(FListe.Cells[SA_Gen,FListe.Row]) then
    begin
     if ( lInSens = RC_DROITE_GAUCHE ) and not ( FInfoEcr.GetString('G_NATUREGENE') = 'D' ) then
      begin // on force le positionnement au debit
       ACol    := SA_DEBIT ;
       ARow    := lr ;
       Cancel  := true ;
       exit ;
      end; // if
    end; // if
  end; // if

 // on peut saisir soit un debit, soit un credit
 if ( lc = SA_Debit ) then
  begin
   if ( FListe.Cells[SA_Credit,lr] <> '' ) then
    begin
     PasseColSuivante(ACol,ARow,FListe) ;
     Cancel := true ;
    end;
   end; // Debit
 if ( lc = SA_Credit ) then
  begin
   if ( FListe.Cells[SA_Debit,lr] <> '' ) then
    begin
     if not LigneCorrecte then
      begin
       ACol   := SA_Debit ;
       Cancel := true;
      end  // if
       else
        PostMessage(FListe.Handle, WM_KEYDOWN, VK_TAB, 0) ;
    end;
  end; // if


 // Interdir la saisie d'un auxiliaire sur un compte non collectif
 if ( lc = SA_AUX ) then
  begin
   if FInfoEcr.LoadCompte(FListe.Cells[SA_Gen,FListe.Row]) and not (FInfoEcr.GetString('G_COLLECTIF') = 'X' ) then
    begin
     if ( lInSens = RC_GAUCHE_DROITE ) then ACol := SA_AUX + 1 else ACol := SA_AUX - 1 ;
     if ARow < FListe.RowCount then
      Cancel := true
       else
        begin // on vient d'une ligne qui n'existe plus
         ARow := FListe.Row ;
         Cancel := true ;
        end ; // if
     Exit ;
    end; // if
  end; // if Aux

 finally
  FListe.EndUpdate ;
 end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/01/2004
Modifié le ... : 15/05/2006
Description .. : - LG - 19/01/2004 - FB 13207 - l'option "ne prends pas en 
Suite ........ : compte le code natureds la saisie en vrac" ne fonctionnait 
Suite ........ : pas
Suite ........ : - LG - 16/09/2004 - FB 13563 - controel du formet de la 
Suite ........ : date avant le recup TOB
Suite ........ : - LG - 19/01/2005- FB 12900 - utilisation direct de la 
Suite ........ : calculatrice ds la zone debit credit
Suite ........ : - LG  - 07/07/2005 - FB 16221 - on ne ait plus le der test sur 
Suite ........ : la date ( qui empeche de saisie sur n+1 avec une date 
Suite ........ : d'entre ds N )
Suite ........ : - LG  - 15/05/2006 - on affecte le type d'exo du tinfo pur le 
Suite ........ : controle des comptes vise ( uniquement valable sur l'exo en 
Suite ........ : cours )
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FListeCellExit( Sender : TObject; var ACol,ARow : Integer ; var Cancel : Boolean);
var
 lTOB : TTOBEcriture ;
 lInErr : integer ;
begin

 if ( csDestroying in Ecran.ComponentState ) or ( Statut = taConsult ) then Exit ;

 {$IFDEF TT}
  AddEvenement('FListeCellExit ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow)+' : FListe.Row:='+intToStr(FListe.Col)+' FListe.Row:='+intToStr(FListe.Row)) ;
 {$ENDIF}
 lInErr := 1 ;
  // Validité de la Date
 if ( ACol = SA_DateC) then
  begin
   Cancel := (length(trim(FListe.Cells[SA_DateC,ARow]))) < 10 ;   
   if not Cancel then
    begin
    // Cancel := not CControleDateBor(FListe.Cells[SA_DateC,ARow],FInfoEcr.Exercice,true) ;
     lInErr := ControleDate(FListe.Cells[SA_DateC,ARow]) ;
     Cancel := ( lInErr <> 0 ) and ( lInErr <> 5 )  ;
    end ; // if
    if Cancel then
     begin
      case lInErr of
       1 : PGIInfo('La date que vous avez renseignée n''est pas valide' ,cStTexteTitreFenetre) ;
       2 : PGIInfo('La date que vous avez renseignée n''est pas dans un exercice ouvert' ,cStTexteTitreFenetre) ;
       3 : PGIInfo('La date que vous avez renseignée est antérieure à une clôture' ,cStTexteTitreFenetre) ;
       4 : PGIInfo('La date que vous avez renseignée est antérieure à une clôture' ,cStTexteTitreFenetre) ;
       5 : PGIInfo('La date que vous avez renseignée est en dehors des limites autorisées' ,cStTexteTitreFenetre) ;
      end ;  // case
      exit ;
     end ; // if
  end; // if Date


 lTOB := RecupTOB(ARow) ;
 if lTOB = nil then
  begin
   {$IFDEF TT}
    MessageAlerte('FListeCellExit : vTOB égale à nil ! ');
   {$ELSE}
    MessageAlerte('Quitter la saisie en vrac !');
   {$ENDIF}
   exit ;
  end; // if

 // on recharge systemtiquement les comptes, aux et journaux a voir par la suite si on peut etre plus fin
 FInfoEcr.Load(FListe.Cells[SA_Gen,ARow],FListe.Cells[SA_Aux,ARow],FListe.Cells[SA_Jal,ARow]) ;
 if FInfoEcr.Compte.InIndex  <> -1 then FListe.Cells[SA_Gen,ARow] := FInfoEcr.StCompte ;
 if FInfoEcr.Aux.InIndex     <> -1 then FListe.Cells[SA_Aux,ARow] := FInfoEcr.StAux ;
 if FInfoEcr.Journal.InIndex <> -1 then FListe.Cells[SA_Jal,ARow] := FInfoEcr.StJournal ;

 // validite du journal
 if (ACol=SA_Etabl) then
  begin
   if not lTOB.IsValidEtabliss then
    begin
     Cancel:=true ;
     PostMessage(FListe.Handle , WM_KEYDOWN , VK_F5 , 0) ;
     exit ;
    end;
  end; // if SA_NAT

 //J_FERME="-" AND ( J_NATUREJAL="ACH" OR J_NATUREJAL="BQE" OR J_NATUREJAL="OD" OR J_NATUREJAL="VTE" OR J_NATUREJAL="CAI")

 // validite du journal
 if (ACol=SA_Jal) then
  begin
   if not lTOB.IsValidJournal or not ( (FInfoEcr.GetString('J_NATUREJAL') = 'ACH') or
                                    (FInfoEcr.GetString('J_NATUREJAL') = 'BQE') or
                                    (FInfoEcr.GetString('J_NATUREJAL') = 'OD') or
                                    (FInfoEcr.GetString('J_NATUREJAL') = 'VTE') or
                                    (FInfoEcr.GetString('J_NATUREJAL') = 'CAI') ) then
    begin
     Cancel := true ;
     PostMessage(FListe.Handle , WM_KEYDOWN , VK_F5 , 0) ;
     exit ;
    end;
   if FBoAfficherNature then
    FListe.Cells[SA_Nat,ARow] := lTOB.Info.Journal.NatureParDefaut ; 
   if FStLastJal <> lTOB.GetValue('E_JOURNAL') then
    begin
     FInIndexComptaAuto := 1 ;
     FStLastJal         := lTOB.GetValue('E_JOURNAL') ;
    end; // if
  end; // if SA_NAT

 // validite de la nature
 if (ACol=SA_Nat) then
  begin
   if not lTOB.IsValidNat then
    begin
     Cancel:=true ;
     PostMessage(FListe.Handle , WM_KEYDOWN , VK_F5 , 0) ;
     exit ;
    end;
   FListe.Cells[SA_Nat,ARow] := lTOB.GetValue('E_NATUREPIECE') ;
  end; // if SA_NAT

 // Validité de la Date
 if ( ACol = SA_DateC) then
  begin
   Cancel := not lTOB.IsValidDateComptable ;
   if Cancel then
    exit
     else
       FRechCompte.Info.TypeExo := CGetTypeExo(StrToDate(FListe.Cells[SA_DateC,ARow])) ;

  end; // if Date

 // Validité du compte général
// if ( ACol = SA_GEN ) and ( FListe.Cells[SA_Gen,FListe.Row] <> '' ) and ( ACol <> ( FListe.Col - 1 ) ) then
 if ( ACol = SA_GEN ) and ( FListe.Cells[SA_Gen,ARow] <> '' ) then //and ( ACol <> ( FListe.Col - 1 ) ) then
  begin
   {$IFDEF TT} AddEvenement('*** FRechCompte.CellExitGen ***') ; {$ENDIF}
   FRechCompte.Ecr := lTOB ;
   HGBeginUpdate(FListe) ;
   try
    FRechCompte.CellExitGen( FListe , ACol , ARow , Cancel ) ;
    if not Cancel and not FRechCompte.Info.Compte.IsVentilable and _AvecAna(lTOB) then
     FZSaisieVrac.Delete(lTOB) ;

   finally
    HGEndUpdate(FListe) ;
   end;
   if Cancel then exit ;
  end; // if

 // Validité du compte auxiliaire
 if ( ACol = SA_AUX ) then
  begin
   {$IFDEF TT} AddEvenement('*** FRechCompte.CellExitAux ***') ; {$ENDIF}
   FRechCompte.Ecr := lTOB ;
   FRechCompte.CellExitAux( FListe , ACol , ARow , Cancel ) ;
   if Cancel then exit ;
  end; // if

 // Validite des montants
 if ( ACol = SA_Debit ) or ( ACol = SA_Credit ) then
  begin
   //calculs
   if (pos('+',FListe.Cells[ACol, ARow])>0) or (pos('-',FListe.Cells[ACol, ARow])>0) or (pos('/',FListe.Cells[ACol, ARow])>0) or (pos('*',FListe.Cells[ACol, ARow])>0) then
    begin
     if Pos(FListe.Cells[ACol, ARow][1],'/')>-1 then
      FListe.Cells[ACol, ARow]:='{"#.###,0"'+FListe.Cells[ACol, ARow]+'}';
     FListe.Cells[ACol, ARow]:=GFormule('{'+FListe.Cells[ACol, ARow]+'}', GetFormulePourCalc, nil, 1) ;
    end ;
   if not lTOB.IsValidMontant( Valeur( FListe.Cells[ACol,ARow] ) ) then
    begin
     Cancel                   := true ;
     FListe.Cells[ACol,ARow]  := '' ;
     exit ;
    end;

  end; // if

 CompleteCeQueJePeux(ACol,ARow) ;
 FListeCellExitModeGuide(Sender,ACol,ARow,Cancel) ;

 AfficheTitreAvecCommeInfo ;

end;

function TOF_CPSAISIEVRAC.GetFormulePourCalc(Formule : hstring) : variant ;
begin
 result := #0 ;
end ;


function _CChoixCpteAuto ( vListeDesComptes : String ) : String ;
var
 lStWhere, lStListeDesComptes ,lStCompte : String ;
begin


 result     := '' ;
 lStWhere   := '' ;

 if vListeDesComptes = '' then Exit ;

 lStListeDesComptes := vListeDesComptes ;

 repeat

   lStCompte := ReadTokenSt(lStListeDesComptes) ;

   if lStCompte <> '' then
    begin
      if Length(lStCompte)=VH^.Cpta[fbGene].Lg then
       lStWhere := lStWhere +' OR G_GENERAL="' + lStCompte + '" '
        else
         lStWhere := lStWhere+' OR G_GENERAL LIKE "'+lStCompte+'%"' ;
    end ; // if

 until ( (lStListeDesComptes = '' ) or ( lStCompte = '') ) ;

 if lStWhere = '' then Exit ;

 Delete(lStWhere,1,4) ;

 Result := Choisir('Choix d''un compte automatique','GENERAUX','G_GENERAL||"  "||G_LIBELLE','G_GENERAL',lStWhere,'G_GENERAL') ;

END ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/01/2005
Modifié le ... :   /  /    
Description .. : - LG - FB 14590 - 03/01/2005 - affectation du libelle apres 
Suite ........ : l'affectation auto du compte de contrepartie du compte
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.GetCompteAuto( vInIndice : integer ) ;
var
 lStCompte : string ;
begin

 if ( FListe.Col <> SA_GEN  ) then exit ;

 FInfoEcr.LoadJournal(FListe.Cells[SA_Jal,FListe.Row]) ;

 if FInfoEcr.Journal.CompteAuto = '' then exit ;

 if vInIndice <=0 then
  lStCompte := _CChoixCpteAuto( FInfoEcr.Journal.CompteAuto )
   else
    lstCompte := TrouveAuto(FInfoEcr.Journal.CompteAuto , vInIndice ) ;  // fct de SaisUtil

 if lStCompte='' then Exit ;

 if EstInterdit(FInfoEcr.GetString('J_COMPTEINTERDIT'), lStCompte, 0) > 0 then Exit ;

 FListe.Cells[SA_GEN, FListe.Row] := lStCompte ;
 Statut                           := taModif ;
 FBoGrilleModif                   := true ;
 FInfoEcr.Load(FListe.Cells[SA_Gen,FListe.Row],FListe.Cells[SA_Aux,FListe.Row],FListe.Cells[SA_Jal,FListe.Row]) ;
 CompleteCeQueJePeux(SA_Lib,FListe.Row ) ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 21/06/2004
Modifié le ... :   /  /
Description .. : - LG - 21/06/2004 - FB 13539 - qd on remonte ds la grille,
Suite ........ : on supprime le ligne courante s'il elle etait vide
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEVRAC.FListeRowExit( Sender : TObject ; Ou : Integer ; var Cancel : Boolean ; Chg : Boolean);
var
 lInSens : integer ;
begin

 {$IFDEF TT}
  AddEvenement('FListeRowExit Ou:='+intToStr(Ou)) ;
 {$ENDIF}

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 lInSens := CGetGridSens(FListe, FListe.Col, Ou) ;

 if lInSens = - 1 then
  begin
   if FListe.Cells[SA_GEN,Ou] = '' then
    begin
     FListeDeleteRow(Ou) ;
     exit ;
    end ;
  end ;

 if not BoGuideRun then
  begin
   Cancel := not LigneCorrecte(Ou) ;
   if not Cancel and  ( FListe.Row = FListe.RowCount - 1 ) then
   AccelerateurDeSaisie ;
  end  // if
   else
    Cancel := not LigneCorrecte(Ou) ;

 AfficheTitreAvecCommeInfo ;

end;

procedure TOF_CPSAISIEVRAC.FListeRowEnter( Sender : TObject; Ou : Integer ; var Cancel : Boolean ; Chg : Boolean);
begin

{$IFDEF TT}
 AddEvenement('FListeRowEnter Ou:='+intToStr(Ou));
{$ENDIF}

 AfficheInfoLigne ;
 EnableButtons ;

 _Plantage     := 0 ;

end;


procedure TOF_CPSAISIEVRAC.FListeElipsisClick (Sender: TObject) ;
var
 ARow      : integer ;
 ACol      : integer ;
 lBoCancel : boolean ;
begin

 if csDestroying in Ecran.ComponentState then exit ;

 {$IFDEF TT}
  AddEvenement('FListeElipsisClick Row:=' + intToStr(FListe.Row)) ;
 {$ENDIF}

 Statut         := taModif ;
 FBoGrilleModif := true ;

 if FListe.Col = SA_DateC then
  begin
   V_PGI.ParamDateProc(FListe);
   Statut         := taModif;
   FBoGrilleModif := true ;
  end; // if

 // Journal
 if ( FListe.Col = SA_Jal ) then
  begin
   // uniquement les journaux de type classiques
   LookupList(FListe,'Journal','JOURNAL','J_JOURNAL','J_LIBELLE',
              'J_FERME="-" AND ( J_NATUREJAL="ACH" OR J_NATUREJAL="BQE" OR J_NATUREJAL="OD" OR J_NATUREJAL="VTE" OR J_NATUREJAL="CAI") ',
              '',true,0) ;
   exit ;
  end;

 // Etablissement
 if FListe.Col = SA_Etabl then
  begin
  // LookupList(FListe,'Etablissement','ETABLISS','ET_ETABLISSEMENT','ET_LIBELLE','','',true,0) ;
   LookUpTob(FListe,FInfoEcr.Etabliss.Items,'Etablissement','ET_ETABLISSEMENT;ET_LIBELLE','Code;Libellé',true ) ;
   exit ;
  end;

  // Nature de pièce
 if FListe.Col = SA_NAT then
  begin
   FZNature.MakeTag( FInfoEcr.GetString('J_NATUREJAL') ) ;
   LookUpTob(FListe,FZNature.Items,'Nature','CO_CODE;CO_LIBELLE','Code;Libellé',true,'1') ;
   exit ;
  end ; // if

 Statut                := taModif ;
 FBoGrilleModif        := true ;
 FRechCompte.Ecr       := RecupTOB ;
 ARow                  := FListe.Row ;
 ACol                  := FListe.Col ;
 if FRechCompte.ElipsisClick(FListe) then
  begin // exit ;
   lBoCancel  := false ;
   FListeCellEnter(nil,ACol,ARow,lBoCancel) ;
   if lBoCancel then
    begin
     FListe.Row := ARow ;
     FListe.Col := ACol ;
    end ;
  end;


end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 04/03/2003
Modifié le ... :   /  /
Description .. : - 04/03/2003 - fct affichage des msg d'erreur. les msg
Suite ........ : d'erreur ne sont afficher que lorsque l'on valide la ligne.
Suite ........ : Pendant la saisie on ouvre les lookup sans msg
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEVRAC.OnErrorRech (sender : TObject; Error : TRecError ) ;
var
 lStMethode : string ;
 lStTitre   : string ;
begin
 lStTitre := Ecran.Caption ;
 if Error.RC_Methode <> ''  then lStMethode := 'Méthode : ' + Error.RC_Methode ;
 if ( Error.RC_Error <> RC_PASERREUR ) then
      AfficheTitreAvecCommeInfo( FMessCompta.GetMessage( Error.RC_Error ) ) ;
      //PGIInfo( FMessCompta.GetMessage( Error.RC_Error ) );
 if ( trim(Error.RC_Message) <> '' ) then PGIInfo(Error.RC_Message + #10#13 + Error.RC_Methode , lStTitre ) ;
end;

procedure TOF_CPSAISIEVRAC.OnErrorTOB (sender : TObject; Error : TRecError ) ;
var
 lStMethode : string ;
begin
 if ( Error.RC_Error = - 1 ) and ( Error.RC_Message = '' ) then exit ;
 if Error.RC_Error = - 2 then
  begin
   FListe.Row := strToInt(Error.RC_Message) ;
   exit ;
  end;
 if FBoAffMessErr and FBoValidationEnCours then
  FMessCompta.Execute(Error.RC_Error)
   else
    begin
     if Error.RC_Methode <> ''  then lStMethode := 'Méthode : ' + Error.RC_Methode ;
     if ( Error.RC_Error <> RC_PASERREUR ) then
      //FMessCompta.Execute(Error.RC_Error) ;
      AfficheTitreAvecCommeInfo( FMessCompta.GetMessage( Error.RC_Error ) ) ;
   end ;
 if ( trim(Error.RC_Message) <> '' ) then PGIInfo(Error.RC_Message + #10#13 + Error.RC_Methode , Ecran.Caption ) ;
end;


function TOF_CPSAISIEVRAC.RecupTOB ( ARow : integer = - 1 ) : TTOBEcriture ;
var
  lStMess : string ;
begin

 {$IFDEF TT}
  AddEvenement('RecupTOB');
{$ENDIF}

 if ARow = - 1 then ARow := FListe.Row ;
 result :=  GetE(FListe,ARow) ;
 if result = nil then exit ;
 if result.NomTable = '' then
  begin
   MessageAlerte('RecupTOB : nom de table vide') ;
   result := nil ;
   exit ;
  end;

 {$IFDEF TT}
  AddEvenement('RecupTOB ARow:='+intToStr(ARow));
 {$ENDIF}

 lStMess :='Date comptable' ;

 try

  //mise à jour des donnees de la TOB
  CRemplirDateComptable(result,strToDate(FListe.Cells[SA_DateC,ARow])) ; lStMess :='Journal' ;
  FInfoEcr.Devise.AffecteTaux(result.GetValue('E_DATECOMPTABLE')) ;
  result.PutValue('E_JOURNAL'        , FListe.Cells[SA_Jal      , ARow] ) ; lStMess :='Etablissement' ;
  if FBoAfficherEtablis then
   begin result.PutValue('E_ETABLISSEMENT'  , FListe.Cells[SA_Etabl    , ARow] ) ; lStMess :='Nature de pièce' ; end
   else result.PutValue('E_ETABLISSEMENT'  , VH^.ETABLISDEFAUT ) ;
  result.PutValue('E_NATUREPIECE'    , GetNature ) ; lStMess :='Général' ;
  result.PutValue('E_GENERAL'        , FListe.Cells[SA_Gen      , ARow] ) ; lStMess :='Auxiliaire' ;
  result.PutValue('E_AUXILIAIRE'     , FListe.Cells[SA_Aux      , ARow] ) ; lStMess :='Libelle' ;
  result.PutValue('E_LIBELLE'        , FListe.Cells[SA_Lib      , ARow] ) ; lStMess :='Référence' ;
  result.PutValue('E_REFINTERNE'     , FListe.Cells[SA_RefI     , ARow] ) ; lStMess :='Montant' ;

  CSetMontants( result,
               Valeur(FListe.Cells[SA_Debit  , ARow]) ,
               Valeur(FListe.Cells[SA_Credit , ARow]) ,
               FInfoEcr.Devise.Dev ,
               true );

 {$IFDEF TT}
//  FListe.Cells[0 , ARow] := VH^.ETABLISDEFAUT ;
 {$ENDIF}

 except
  on E : Exception do
   begin
    MessageAlerte('Erreur de récupération de la colonne :' + lStMess + #10#13 + E.Message ) ;
    result := nil ;
  end; // on
 end; // try

  {$IFDEF TT}
  AddEvenement('Fin de RecupTOB');
{$ENDIF}

end;

procedure TOF_CPSAISIEVRAC.CompleteCeQueJePeux( ACol : integer = - 1 ; ARow : integer = - 1 ) ;
var
 lTOB : TTOBEcriture ;
begin

 {$IFDEF TT}
  AddEvenement('CompleteCeQueJePeux');
 {$ENDIF}

 if ACol = -1 then ACol := FListe.Col ;
 if ARow = -1 then ARow := FListe.Row ;

 lTOB := GetE(FListe) ;

 if Valeur( FListe.Cells[SA_Debit,ARow] ) <> 0 then
  FListe.Cells[SA_Debit,ARow]  := STRFMONTANT ( Valeur( FListe.Cells[SA_Debit,ARow] ) , 15 , V_PGI.OkDecV, '' , true);

 if Valeur( FListe.Cells[SA_Credit,ARow] ) <> 0 then
  FListe.Cells[SA_Credit,ARow]  := STRFMONTANT ( Valeur( FListe.Cells[SA_Credit,ARow] ) , 15 , V_PGI.OkDecV, '' , true);

 if (trim(GetNature) = '') and (trim(FListe.Cells[SA_Jal,ARow]) <> '') then
  FListe.Cells[SA_Nat,ARow] := FInfoEcr.Journal.NatureParDefaut ;

 if ( ACol = SA_Lib ) then
  begin

   if FListe.Cells[SA_Lib,ARow] = '' then
    begin
     if trim(FListe.Cells[SA_Aux,ARow]) <> ''  then
      FListe.Cells[SA_Lib,ARow] := FInfoEcr.GetString('T_LIBELLE')
       else
        if not FInfoEcr.Compte.isCollectif then
         FListe.Cells[SA_Lib,ARow] := FInfoEcr.GetString('G_LIBELLE') ;
     if ( lTOB <> nil ) then lTOB.PutValue('LIBCOMPTE' , 'X') ;
     
  end; // if

   if ( FListe.Cells[SA_Lib,ARow] <> '' ) then
    begin
     if ( lTOB <> nil ) and ( FListe.Cells[SA_Lib,ARow] = FInfoEcr.GetString('G_LIBELLE') ) then lTOB.PutValue('LIBCOMPTE' , 'X') ;
     if ( lTOB <> nil ) and ( FListe.Cells[SA_Lib,ARow] = FInfoEcr.GetString('T_LIBELLE') ) then lTOB.PutValue('LIBCOMPTE' , 'X') ;
    end;

  end; // if

 
 G_LIBELLE.Caption := FInfoEcr.GetString('G_LIBELLE') ;
 T_LIBELLE.Caption := FInfoEcr.GetString('T_LIBELLE') ;

end;

function TOF_CPSAISIEVRAC.ControleColPrecedent( var ACol, ARow : integer ) : boolean;
var
 i : integer ;
begin
 result := true ;
 if ( Statut = taConsult ) or ( ACol >= SA_Gen ) or ( ARow <> FListe.Row ) then exit ;
 for i := 1 to ( FListe.Col - 1 ) do
  begin
   if FListe.ColWidths[i] = - 1 then continue ;
   result := FListe.Cells[i,ARow] <> '' ;
   if not result then
    begin
     ACol := i ;
     exit ;
    end; // if
  end; // for
 result := true ;
end;



procedure TOF_CPSAISIEVRAC.AfficheInfoLigne ;
begin
 {$IFDEF TT}
  AddEvenement('AfficheInfoLigne');
 {$ENDIF}

 {$IFDEF TT}
 // FListe.Cells[SA_RefI,FListe.Row] :=  floattostr(Arrondi(FRdSoldeDebit - FRdSoldeCredit , V_PGI.OkDecV ) ) ;
 {$ENDIF}

 FInfoEcr.Load(FListe.Cells[SA_Gen,FListe.Row],FListe.Cells[SA_Aux,FListe.Row],FListe.Cells[SA_Jal,FListe.Row]) ;

 G_LIBELLE.Caption := FInfoEcr.GetString('G_LIBELLE') ;
 T_LIBELLE.Caption := FInfoEcr.GetString('T_LIBELLE') ;

end;

procedure TOF_CPSAISIEVRAC.NextRow ;
var
 lBoCancel : boolean ;
 lBoChg    : boolean ;
begin
  {$IFDEF TT}
  AddEvenement('NextRow');
 {$ENDIF}

 lBoChg := true ;
 FListeRowExit( nil , FListe.Row , lBoCancel, lBoChg );
 if lBoCancel then exit ;
 CreateRow ;

 FBoGrilleVide := FZSaisieVrac.Detail.Count = 0 ;
 EnableButtons ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 21/10/2003
Modifié le ... :   /  /    
Description .. : - LG - 22/10/2003 - FB 12897 - on reprenais 
Suite ........ : systematiquement le libelle du compte sur la ligne siuvante
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.InitRow ;
var
 lTOB       : TTOBEcriture ;
 lRowRef    : integer ;
 lBoHaut    : boolean ;
begin

 {$IFDEF TT}
  AddEvenement('InitRow FListe.Row:='+IntToStr(FListe.Row));
 {$ENDIF}

 if FListe.Row = 1 then
  begin
   FListe.Cells[SA_DateC , FListe.Row] := DateToStr(FDtDateDeb) ;
   if FBoAfficherEtablis then
    FListe.Cells[SA_Etabl , FListe.Row] := VH^.ETABLISDEFAUT ;
   if FBoAfficherNature then
    FListe.Cells[SA_Nat   , FListe.Row] := 'OD' ;
   if FListe.Row < ( FListe.RowCount -1 ) then
    FListe.Cells[SA_Jal   , FListe.Row] :=  FListe.Cells[SA_Jal   , FListe.Row + 1] ;
  end
   else
    begin

     if FListe.Row = ( FListe.Rowcount - 1 ) then
      lRowRef := FListe.Row - 1
       else
        lRowRef := FListe.Row + 1 ;

      lTOB := GetE(FListe,lRowRef) ;
      FListe.Cells[SA_DateC , FListe.Row ] := FListe.Cells[SA_DateC , lRowRef ] ;
      if FBoAfficherNature then  FListe.Cells[SA_Nat , FListe.Row ] := FListe.Cells[SA_Nat , lRowRef ] ;
      if FBoAfficherEtablis then FListe.Cells[SA_Etabl , FListe.Row]  := FListe.Cells[SA_Etabl , lRowRef ] ;
      FListe.Cells[SA_Jal , FListe.Row]  := FListe.Cells[SA_Jal , lRowRef ] ;
      if not CalculSeparateur(FListe.Col,FListe.Row,lBoHaut,true) then
       begin
        if ( lTOB <> nil ) and ( lTOB.GetValue('LIBCOMPTE') = '-' ) then FListe.Cells[SA_Lib , FListe.Row]   := FListe.Cells[SA_Lib   , lRowRef ] ;
        FListe.Cells[SA_RefI , FListe.Row]  := FListe.Cells[SA_RefI , lRowRef ] ;
       end; // if
    end;//if

end;


procedure TOF_CPSAISIEVRAC.CreateRow( ARow : integer = -1 ; vBoInsert : boolean = false );
var
 lTOB : TOB ;
 lTOBSuiv : TOB ;
 lTOBPrec : TOB ;

 procedure _Init ;
  begin
   InitRow ;
   FListe.Col      := SA_DateC ;
   if FListe.CanFocus then FListe.SetFocus ;
   FListe.ShowEditor ;
   if lTOBPrec <> nil then
    begin
     if lTOBPrec.GetValue('SOLDE') = 0 then
      FListe.Col      := SA_DateC
       else
        FListe.Col      := SA_Gen ;
    end ; // if
   {$IFNDEF TT}
   if FListe.Cells[SA_Numl,FListe.Row] = '' then FListe.Cells[SA_Numl,FListe.Row] := intToStr(FListe.RowCount-1) ;
   {$ENDIF}
  end;

 procedure _InsererLigne ;
  begin
   FListe.InsertRow(ARow) ;
   FListe.Row      := ARow ;          // on se place sur cette nouvelle ligne
   SetE(FListe,FZSaisieVrac.Insert(FListe.Row - 1)) ;
   NumeroteGrille ;
   FListe.Refresh ;
  end;

 procedure _AjouterLigne ;
  begin
   FListe.RowCount := FListe.RowCount + 1 ;
   FListe.Row      := FListe.RowCount - 1 ;          // on se place sur cette nouvelle ligne
   SetE(FListe,FZSaisieVrac.Add) ;
   ARow            := FListe.Row ;
   {$IFNDEF TT}
   FListe.Cells[SA_Numl,FListe.Row] := intToStr(FListe.RowCount-1) ;
   {$ENDIF}
  end;

begin

  if ( Statut <> taConsult )  then exit; // on ajout une ligne uniquement si on etait en consultation

 {$IFDEF TT}
  AddEvenement('CreateRow Row:='+intToStr(FListe.Row));
 {$ENDIF}

 Statut         := taCreat ;
 FBoGrilleModif := true ;

 if ARow = - 1 then ARow := FListe.Row ;

 HGBeginUpdate(FListe) ;

 try

  if vBoInsert then
   _InsererLigne
    else
     begin
      if ( FListe.Row = 1 ) and ( GetE(FListe) = nil ) then
       SetE(FListe,FZSaisieVrac.Add)
        else
         if ( FListe.Row = FListe.RowCount - 1 )  then
           _AjouterLigne
        else
         if ( FListe.Row > 1 ) and ( FListe.Row < FListe.RowCount  ) then
          _InsererLigne ;
     end; // if

  AfficheTitreAvecCommeInfo ;
  EnableButtons ;

  // on enleve les montants de la ligne ds le solde
  // mis a jour ds le rowexit
  lTOB      := GetE(FListe) ;
  lTOBprec  := nil ;
  lTOBSuiv  := nil ;

  if ARow < FListe.RowCount - 1 then
   lTOBSuiv := GetE(FListe,ARow + 1) ;

  if ARow > 0 then
   lTOBPrec := GetE(FListe,ARow - 1) ;

  if lTOB = nil then exit ;
  FRdSoldeDebit  := FRdSoldeDebit  - lTOB.GetValue('E_DEBIT') ;
  FRdSoldeCredit := FRdSoldeCredit - lTOB.GetValue('E_CREDIT') ;
  FRdSolde       := Arrondi(FRdSoldeDebit - FRdSoldeCredit , V_PGI.OkDecV ) ;

  if vBoInsert then
   lTOB.PutValue('E_NUMGROUPEECR',lTOBSuiv.GetValue('E_NUMGROUPEECR'))
    else
     if lTOBPrec <> nil then
      lTOB.PutValue('E_NUMGROUPEECR',FInNumGroupEcr) ;

  FListe.OnEnter := nil ;    
  _Init ;
  FListe.OnEnter := FListeEnter;

  {$IFDEF TT}
  // if lTOB <> nil then
  //  FListe.Cells[SA_Numl,FListe.Row] := inttostr(lTOB.GetValue('e_numligne')) ;
   FListe.Cells[SA_Numl,FListe.Row] := floatToStr(FRdSolde ) ;
  {$ENDIF}

 AfficheInfoLigne ;

 finally
  HGEndUpdate(FListe) ;
 end; // try

end;


procedure TOF_CPSAISIEVRAC.AfficheTitreAvecCommeInfo( vInfo : string ='' );
var
 lTOB : TOB ;
begin

  if FBoValidationEnCours then exit ;

 {$IFDEF TT}
  AddEvenement('AfficheTitreAvecCommeInfo' );
 {$ENDIF}

 Ecran.Caption := FStTitre ;
 if trim(vInfo)<>'' then Ecran.Caption := Ecran.Caption + ' : ' + TraduireMemoire(vInfo) ;

 lTOB          := GetE(Fliste,FListe.Row) ;
 if ( lTOB <> nil ) and ( lTOB.NomTable <> '' ) and ( lTOB.GetValue('CODEERREUR') <> RC_PASERREUR ) then
  begin
   Ecran.Caption := Ecran.Caption + ' : ' + FMessCompta.GetMessage(lTOB.GetValue('CODEERREUR')) ;
  end;

 {$IFDEF TT}
 case Statut of
  taCreat   : Ecran.Caption := Ecran.Caption + ' ( Creation )' ;
  taConsult : Ecran.Caption := Ecran.Caption + ' ( Consult )' ;
  taModif   : Ecran.Caption := Ecran.Caption + ' ( Modif )' ;
 end; // case
 if FBoGrilleModif then Ecran.Caption := Ecran.Caption + ' Modifiee ' else Ecran.Caption := Ecran.Caption + ' Pas modifiee ' ;
 {$ENDIF}

 UpdateCaption(Ecran) ;

end;

procedure TOF_CPSAISIEVRAC.FListeDeleteRow( Row : integer ) ;
var
 lTOB : TOB ;
begin

 if BoGuideRun then exit ;

 if Row = - 1 then Row := FListe.Row ;
 {$IFDEF TT}
  AddEvenement('FListeDeleteRow Row ' + IntToStr(Row) );
 {$ENDIF}

 HGBeginUpdate(FListe) ;

 try

 lTOB           := GetE(FListe,Row) ;
 if lTOB = nil then exit ;
 FRdSoldeDebit  := FRdSoldeDebit  - lTOB.GetValue('E_DEBIT') ;
 FRdSoldeCredit := FRdSoldeCredit - lTOB.GetValue('E_CREDIT') ;
 FRdSolde       := Arrondi(FRdSoldeDebit - FRdSoldeCredit , V_PGI.OkDecV ) ;
 AfficheSolde ;

 FZSaisieVrac.Delete(lTOB) ;
 SetE(FListe,nil,Row) ;

 if FListe.RowCount = 2 then
  begin
   Statut        := taConsult ;
   FListe.VidePile(False) ;
   FBoGrilleVide := true ;
   CreateRow ;
  end
   else
    begin
     if ( FListe.Row <> 1 ) and ( FListe.Row = Row ) then // si on n'est pas sur la derniere cellule on remonte d'une ligne
      FListe.Row := FListe.Row - 1 ;
     FListe.DeleteRow(Row) ; // on supprime la ligne
     Statut := taConsult ;
     EnableButtons ;
    end; // if

 AfficheInfoLigne ;

 finally
  HGEndUpdate(FListe) ;
 end; // try

end;

procedure TOF_CPSAISIEVRAC.NextNature ;
begin
 FZNature.MakeTag( FInfoEcr.GetString('J_NATUREJAL') ) ;
 FListe.Cells[SA_Nat,FListe.Row] := FZNature.NextNature ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 09/08/2004
Modifié le ... :   /  /    
Description .. : - LG - 09/08/2004 - FB 13785 13871 - des que l'on
Suite ........ : modofoe le statu on valide les boutons, pour empecher de 
Suite ........ : supprime puis valider juste der.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.SetStatut ( Value : TActionFiche );
begin
// if not ( ( Value = taModif ) and ( Statut = taCreat ) ) then // le F5 place la grille en modification
 FStatut := Value ;
 AfficheTitreAvecCommeInfo ;
 EnableButtons ;
end;

//
// gestion du bp
//

procedure TOF_CPSAISIEVRAC.BValiderClick(Sender: TObject);
begin
 Validation ;
end;


//
//  gestion du mul
//

procedure TOF_CPSAISIEVRAC.UpdateGrille;
begin
 if ((THSystemMenu(GetControl('HMTrad')).ActiveResize) and (V_PGI.Outlook)) then
  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 30/09/2003
Modifié le ... :   /  /    
Description .. : - LG - 30/09/2003 - on cache l'editeur avant de relancer la 
Suite ........ : recherche
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSAISIEVRAC.ChercheClick( vBoEnMemoire : boolean = true);
begin

 {$IFDEF TT}
  AddEvenement('BChercheClick');
 {$ENDIF}

  FListe.HideEditor ;

  if ( Statut <> taConsult ) and not ( FListe.Row = 1 ) then
   begin
    //FListeDeleteRow ;
    PGIInfo('Pas de recherche en mode création',Ecran.Caption );
    exit ;
   end; // if}

  if vBoEnMemoire and ( Fliste.Cells[SA_GEN,FListe.Row] = '' ) and ( FListe.Row = 1 ) then exit ;

  SourisSablier ;
  HGBeginUpdate(Fliste) ;

  try

   if not vBoEnMemoire and not RemplitGrille then exit ;
   CVerifCritere(Ecran);
   RefreshFListe(vBoEnMemoire) ;
   THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
   UpdateGrille ;

  finally
   HGEndUpdate(Fliste) ;
   SourisNormale ;
  end;

end;




function TOF_CPSAISIEVRAC.RemplitGrille : boolean;
begin

 {$IFDEF TT}
  AddEvenement('RemplitGrille');
 {$ENDIF}

 result := false ;

 // blocage
 if _Blocage(['nrSaisieVrac'], true, 'nrSaisieVrac') then exit ;

 try

   FZSaisieVrac.StSQLFiltre := RecupWhereCritere(FPageControl) ;
   result                   := Transactions(FZSaisieVrac.load,1) = oeOK  ;

   if not result then MessageAlerte('Erreur lors de l''enregistrement ! ' + #10#13 + V_PGI.LastSQLError  ) ;

 finally
  _Bloqueur('nrSaisieVrac',False) ;
 end ;

end;

function TOF_CPSAISIEVRAC.ControleCritere( vTOB : TOB ) : boolean ;
begin
 result := vTOB.GetValue('SUPPRIMER') <> 'X' ;
 if E_JOURNAL.Text <> '' then
  result := vTOB.GetValue('E_JOURNAL') = AnsiUpperCase(E_JOURNAL.Text) ;
 if E_ETABLISSEMENT.Text <> '' then
  result := result and ( vTOB.GetValue('E_ETABLISSEMENT') = AnsiUpperCase(E_ETABLISSEMENT.Text ) ) ;
 if E_UTILISATEUR.Text <> '' then
  result := result and ( vTOB.GetValue('E_UTILISATEUR') = AnsiUpperCase(E_UTILISATEUR.Text ) ) ;
 if E_DATECOMPTABLE.Text <> '' then
  result := result and ( ( vTOB.GetValue('E_DATECOMPTABLE') >= StrToDate(E_DATECOMPTABLE.Text) ) and ( vTOB.GetValue('E_DATECOMPTABLE') <= StrToDate(E_DATECOMPTABLE_.Text) ) ) ;
end;


function TOF_CPSAISIEVRAC.RefreshLigne( vIndexGrille : integer ; vIndexTOB : integer = - 1 ) : boolean ;
var
 lTOB     : TOB ;
 lTOBPrec : TOB ;
 lTOBSuiv : TOB ;
begin

 lTOBPrec := nil ;
 lTOBSuiv := nil ;
 result   := true ;

 if vIndexTOB = -1 then
  vIndexTOB := vIndexGrille ;

 lTOB := FZSaisieVrac.Detail[vIndexTOB] ;
 
 if vIndexTOB > 0                             then lTOBPrec := FZSaisieVrac.Detail[vIndexTOB-1] ;
 if vIndexTOB < FZSaisieVrac.Detail.Count - 1 then lTOBSuiv := FZSaisieVrac.Detail[vIndexTOB+1] ;

 if ( vIndexGrille + 1 ) > ( FListe.RowCount - 1 ) then
  FListe.RowCount  := FListe.RowCount + 1 ;

 FListe.Row                       := vIndexGrille + 1 ;
 FListe.Cells[SA_Numl,FListe.Row] := intToStr(FListe.RowCount - 1) ;
 SetE(FListe,TTOBEcriture(lTOB)) ;
 CalculInfoSupp( lTOB , lTOBPrec , lTOBSuiv ) ;

 //mise à jour des donnees de la TOB
 FListe.Cells[SA_DateC    , vIndexGrille+1] := DateToStr( lTOB.GetValue('E_DATECOMPTABLE') ) ;
 FListe.Cells[SA_Jal      , vIndexGrille+1] := lTOB.GetValue('E_JOURNAL'      ) ;
 if FBoAfficherEtablis then
  FListe.Cells[SA_Etabl    , vIndexGrille+1] := lTOB.GetValue('E_ETABLISSEMENT') ;
 if FBoAfficherNature then
  FListe.Cells[SA_Nat      , vIndexGrille+1] := lTOB.GetValue('E_NATUREPIECE'  ) ;
 FListe.Cells[SA_Gen      , vIndexGrille+1] := lTOB.GetValue('E_GENERAL'      ) ;
 FListe.Cells[SA_Aux      , vIndexGrille+1] := lTOB.GetValue('E_AUXILIAIRE'   ) ;
 FListe.Cells[SA_Lib      , vIndexGrille+1] := lTOB.GetValue('E_LIBELLE'      ) ;
 FListe.Cells[SA_RefI     , vIndexGrille+1] := lTOB.GetValue('E_REFINTERNE'   ) ;

 if lTOB.GetValue('E_DEBIT') <> 0 then // si debit = 0 la cellule doit afficher ''
  FListe.Cells[SA_Debit  , vIndexGrille+1] := STRFMONTANT ( lTOB.GetValue('E_DEBIT') , 15 , V_PGI.OkDecV, '' , true)
   else
    FListe.Cells[SA_Debit  , vIndexGrille+1] := '' ;

  if lTOB.GetValue('E_CREDIT') <> 0 then // si debit = 0 la cellule doit afficher ''
   FListe.Cells[SA_Credit  , vIndexGrille+1] := STRFMONTANT ( lTOB.GetValue('E_CREDIT') , 15 , V_PGI.OkDecV, '' , true)
    else
     FListe.Cells[SA_Credit  , vIndexGrille+1] := '' ;

 FRdSoldeDebit  := FRdSoldeDebit  + lTOB.GetValue('E_DEBIT') ;
 FRdSoldeCredit := FRdSoldeCredit + lTOB.GetValue('E_CREDIT') ;
 lTOB.PutValue('GRILLE','X') ;

end ;

function TOF_CPSAISIEVRAC.RefreshFListe( vBoEnMemoire : boolean = true ) : integer;
var
 i        : integer ;
 lInIndex : integer ;
 lTOB     : TOB ;
begin

 {$IFDEF TT}
   AddEvenement('RefreshFListe ARow:='+intToStr(FListe.Row)) ;
   if ( FZSaisieVrac.Detail.Count = 0 ) then AfficheTitreAvecCommeInfo('RefreshFListe : TOB vide') ;
 {$ENDIF}

 result := 0 ;

 if ( FZSaisieVrac = nil ) or ( FZSaisieVrac.Detail = nil ) then exit ;

 FRdSoldeDebit  := 0 ;
 FRdSoldeCredit := 0 ;
 FListe.VidePile(false) ;
 FInNumGroupEcr := 0 ;
 //i              := 0 ;

 lInIndex := 0 ;

 for i := 0 to FZSaisieVrac.Detail.Count - 1 do
  begin

    lTOB := FZSaisieVrac.Detail[i] ;

    if vBoEnMemoire and ( not ControleCritere(lTOB) ) then
     lTOB.PutValue('GRILLE','-') 
     else
      begin
       RefreshLigne(lInIndex,i) ;
       Inc(lInIndex) ;
      end ;

  end ;


 FRdSolde       := Arrondi(FRdSoldeDebit - FRdSoldeCredit , V_PGI.OkDecV ) ;

 AfficheSolde ;

 FBoGrilleVide  := FZSaisieVrac.Detail.Count = 0 ;
 result         := FZSaisieVrac.Detail.Count - 1 ;

  if FBoGrilleVide then
    CreateRow
     else
      begin
       if FListe.CanFocus then FListe.SetFocus ;
       PostMessage(FListe.Handle, WM_KEYDOWN, VK_LEFT, 0) ;
       PostMessage(FListe.Handle, WM_KEYDOWN, VK_RIGHT, 0) ;
       FListe.ShowEditor ;
      end ;// if


end;


//
//  enregistrement
//

function TOF_CPSAISIEVRAC.Validation : boolean ;
begin

 result               := false ;

 if FBoValidationEnCours then exit ;

 if ( Statut <> taConsult ) and ( FListe.Row <  ( FListe.Rowcount - 1 ) ) and ( not BValider.Enabled ) then exit ;

 if ( Statut = taModif) and not LigneCorrecte then exit ;

{$IFDEF TT}
  AddEvenement('Validation');
{$ENDIF}

 if _Blocage(['nrSaisieVrac'], true, 'nrSaisieVrac') then exit ;

 HGBeginUpdate(FListe) ;

 AfficheTitreAvecCommeInfo('Enregistrement en cours ! ') ;
 EnableButtons ;
 FListe.PostDrawCell  := nil ;
 FBoValidationEnCours := true ;

 try

  result := Transactions(FZSaisieVrac.Save,1) = oeOK  ;

 {$IFDEF TT}
  AddEvenement('******************Validation');
{$ENDIF}

// FZSaisieVrac.Save ; result := true ;

  if not result then
   MessageAlerte('Erreur lors de l''enregistrement ! ' + #10#13 + V_PGI.LastSQLError  )
    else
     RefreshFListe ;

 finally
  _Bloqueur('nrSaisieVrac',False) ;
  AfficheTitreAvecCommeInfo ;
  HGEndUpdate(FListe) ;
  EnableButtons ;
  FListe.PostDrawCell  := FListePostDrawCell ;
  FListe.SynEnabled    := true ;
  FBoGrilleModif       := false ;
  AssignEvent ;
  FBoValidationEnCours := false ;
 end; //

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 21/10/2003
Modifié le ... : 15/09/2004
Description .. : - LG - 22/10/2003 - FB 12896 - la bouton valide est tout le 
Suite ........ : temps actifs !
Suite ........ : - LG - 07/05/2004 - FB 13206 - le bouton valider est actif 
Suite ........ : que que l'on a saisie une ligne
Suite ........ : - onne peux integrer le ecr qd qd elles sont enr.
Suite ........ : - LG  - 15/09/2004 - FB 13206 - on peu integrer a tout 
Suite ........ : moment
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.EnableButtons( Ou : integer = - 1);
begin
 BValider.Enabled := not FBoValidationEnCours and not BoGuideRun  ; //not FBoGrilleVide ;
 BInsert.Enabled  := not FBoValidationEnCours and ( Statut = taConsult );
 BDelete.Enabled  := not FBoValidationEnCours and not BoGuideRun ;
 BIntegre.Enabled := not BoGuideRun ;
 TToolbarButton97(GetControl('BFILTRE')).Enabled :=  not BoGuideRun ;
 BCherche.Enabled :=  not BoGuideRun ;
end;

procedure TOF_CPSAISIEVRAC.E_JOURNALElipsisClick(Sender: TObject);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 LookupList(E_JOURNAL,'Journal','JOURNAL','J_JOURNAL','J_LIBELLE','J_FERME="-" AND ( J_NATUREJAL="ACH" OR J_NATUREJAL="BQE" OR J_NATUREJAL="OD" OR J_NATUREJAL="VTE" ) ','',true,0) ;

end;


procedure TOF_CPSAISIEVRAC.E_JOURNALExit(Sender: TObject);
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 if ( trim(E_JOURNAL.Text) <> '' ) and ( not FInfoEcr.LoadJournal(E_JOURNAL.Text) ) then
  begin
   if E_JOURNAL.CanFocus then E_JOURNAL.SetFocus;
   E_JOURNALElipsisClick(nil) ;
  end ;

end;

//
// gestion de l'analytique
//

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 10/10/2003
Modifié le ... : 19/09/2007
Description .. : - LG - 10/10/2003 - FB 12891 - l'affectation automatique de 
Suite ........ : l'analytiq ne fct pas ( les lignes de tob d'analytiq n'eatit pas 
Suite ........ : creer, la fct ventilerTOB plantait )
Suite ........ : - LG - 19/09/2007 - FB 21447 - ne ne teste plus l'anlytiq sur 
Suite ........ : le ligne vide
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.GereAna ( ARow : integer = -1 ; vBoAfficher : boolean = false );
var
 lTOB      : TOB ;
begin

 if ARow = - 1 then ARow := FListe.Row ;

 if trim(FListe.Cells[SA_Gen,ARow]) = '' then exit ;

 lTOB := GetE(FListe, ARow) ;
 if lTOB = nil then exit ;

 if ( lTOB.GetValue('E_ANA') = 'X' ) and not ( _AvecAna(lTOB) ) then
  FZSaisieVrac.LoadAna(lTOB) ;


 if FInfoEcr.LoadCompte( FListe.Cells[SA_Gen,ARow]) then exit ;

 if not FInfoEcr.Compte.IsVentilable then
  begin
   lTOB.PutValue('E_ANA','-' ) ;
   exit ;
  end
   else
    lTOB.PutValue('E_ANA','X' ) ;

 // test si l'on a deja saisie de l'analytique sans demander l'affichage
 if ( not vBoAfficher ) and ( lTOB.GetValue('ANA') = 'X') then exit ;

 if vBoAfficher or ( (VH^.ZGEREANAL) and (VH^.ZSAISIEANAL) )then
  GetAna(lTOB)
   else
    AlloueAxe(lTOB) ;

 CVentilerTOB(lTOB,FInfoEcr) ;
 //VentilerTOB(lTOB, '', 0, FInfoEcr.Devise.Dev.Decimale , false ) ;
 lTOB.PutValue('ANA','X' ) ;
 FZSaisieVrac.AddAna(lTOB) ;

end;

procedure TOF_CPSAISIEVRAC.GetAna( vTOB : TOB );
var
 lRecA     : ARG_ANA ;
 lGGeneral : TGGeneral ;
 lNumAxe   : integer ;
begin

 lGGeneral    := nil ;

 try

  lGGeneral             := CZompteVersTGGeneral ( FInfoEcr.Compte ) ;
  lRecA.Action	        := taCreat ;
  lRecA.CC              := lGGeneral ;
  lRecA.GuideA          := '' ;
  lRecA.VerifVentil     := True ;
  lRecA.ControleBudget  := false ;
  lRecA.DernVentilType  := '' ;
  lRecA.DEV             := FInfoEcr.Devise.Dev ;
  lRecA.QuelEcr         := EcrGen ;
  lReCA.NumGeneAxe      := 0 ;
  lRecA.VerifQte        := False ;
  lRecA.Info            := FInfoEcr ;
  lRecA.MessageCompta   := FMessCompta ;

  // Préremplissage de l'analytique ( la saisie analytique ne le fait plus...)
  if CAOuvrirVentil( vTOB, nil, FInfoEcr, lNumAxe ) then
   begin
    if VentilationExiste( vTOB ) then
     lRecA.Action	:= taModif ;
    // Appel de la saisie analytique
    eSaisieAnal(vTOB , lRecA ) ;
   end ;

 finally
  if assigned(lGGeneral) then lGGeneral.Free;
 end;

end;

procedure TOF_CPSAISIEVRAC.BDeleteClick(Sender: TObject);
begin
 FListeDeleteRow ;
end;

procedure TOF_CPSAISIEVRAC.BInsertClick(Sender: TObject);
begin
 CreateRow(FListe.Row,true) ;
end;


procedure TOF_CPSAISIEVRAC.BAgrandirClick(Sender: TObject);
begin
 FPageControl.Visible := not FPageControl.Visible;
 if FPageControl.Visible then
  BAgrandir.Hint := TraduireMemoire('Agrandir')
   else
    BAgrandir.Hint := TraduireMemoire('Réduire')
end;


procedure TOF_CPSAISIEVRAC.AfficheSolde;
begin
 SOLDEDEBIT.Caption  := STRFMONTANT ( FRdSoldeDebit  , 15 , V_PGI.OkDecV, '' , true) ;
 SOLDECREDIT.Caption := STRFMONTANT ( FRdSoldeCredit , 15 , V_PGI.OkDecV, '' , true) ;
 SOLDE.Caption       := AfficheDBCR ( FRdSolde ) ;
 Solde.Font.Color    := IIF(Pos('C', SOLDE.Caption) > 0, ClGreen, ClRed) ;
end;

function TOF_CPSAISIEVRAC.GetNature : string ;
begin
 if FBoAfficherNature then
  result := FListe.Cells[SA_Nat , FListe.Row]
   else result := 'OD' ;
end;

procedure TOF_CPSAISIEVRAC.NumeroteGrille ;
{$IFNDEF TT}
var
 i : integer ;
{$ENDIF}
begin
 HGBeginUpdate(FListe) ;
 {$IFNDEF TT}
 for i:= 1 to FListe.RowCount - 1 do
  FListe.Cells[SA_Numl,i] := intToStr(i) ;
 {$ENDIF}
 HGEndUpdate(FListe) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 10/12/2003
Modifié le ... :   /  /    
Description .. : - 10/12/2003 - FB 13090 - on recopie le smontant aussi
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.VK_F7FListe ;
begin
{ if ( FListe.Col <> SA_DEBIT) and ( FListe.Col <> SA_CREDIT ) and ( FListe.Row > 1 ) then
  begin
   FListe.Cells[FListe.Col,FListe.Row] := FListe.Cells[FListe.Col , FListe.Row - 1] ;
   Statut                              := taModif ;
  end; // if   }
if ( FListe.Row > 1 ) then
  begin
   FListe.Cells[FListe.Col,FListe.Row] := FListe.Cells[FListe.Col , FListe.Row - 1] ;
   FListe.InplaceEditor.SelStart       := Length(FListe.Cells[FListe.Col,FListe.Row]) ;
   Statut                              := taModif ;
   FBoGrilleModif                      := true ;
  end; // if
end;


procedure TOF_CPSAISIEVRAC.TrouveIndexPiece( vInNumGroupeEcr : integer ; var vInIndexDeb : integer ; var vInIndexFin : integer) ;
var
 i         : integer ;
 lBoTrouve : boolean ;
begin

 {$IFDEF TT}
  AddEvenement('TrouveIndexPiece');
 {$ENDIF}

 i           := 0 ;
 lBoTrouve   := false ;
 vInIndexDeb := 0 ;
 vInIndexFin := FZSaisieVrac.Detail.Count - 1 ;

 if vInNumGroupeEcr = -1 then exit ;

 while ( i < FZSaisieVrac.Detail.Count ) and not lBoTrouve do
  begin
   lBoTrouve := FZSaisieVrac.Detail[i].GetValue('E_NUMGROUPEECR') = vInNumGroupeEcr ;
   inc(i) ;
  end; // while

 if lBoTrouve then
  begin
   vInIndexDeb := i - 1 ;
   lBoTrouve   := false ;

   while ( i < FZSaisieVrac.Detail.Count ) and not lBoTrouve do
    begin
     lBoTrouve := FZSaisieVrac.Detail[i].GetValue('E_NUMGROUPEECR') <> vInNumGroupeEcr ;
     inc(i) ;
    end; // while

   if lBoTrouve then
    vInIndexFin := i - 2
     else
      vInIndexFin := i - 1 ;

  end ; // if

end;


procedure TOF_CPSAISIEVRAC.CalculSoldeFolio ( vInNumGroupeEcr : integer = -1 ) ;
var
 i             : integer ;
 lInIndexDeb   : integer ;
 lInIndexFin   : integer ;
 lTOBPrec      : TOB ;
 lTOBSuiv      : TOB ;
begin

 {$IFDEF TT}
  AddEvenement('CalculSoldeFolio');
 {$ENDIF}

 FRdSoldeDebit  := 0 ;
 FRdSoldeCredit := 0 ;

 TrouveIndexPiece (vInNumGroupeEcr,lInIndexDeb,lInIndexFin) ;
 FInNumGroupEcr := 0 ;
 lTOBPrec       := nil ;
 lTOBSuiv       := nil ;

 for i := lInIndexDeb to lInIndexFin  do
  begin
   FRdSoldeDebit  := FRdSoldeDebit  + FZSaisieVrac.Detail[i].GetValue('E_DEBIT') ;
   FRdSoldeCredit := FRdSoldeCredit + FZSaisieVrac.Detail[i].GetValue('E_CREDIT') ;
   if i > 0 then
    lTOBPrec := FZSaisieVrac.Detail[i-1] ;
   if i < FZSaisieVrac.Detail.Count - 1 then
    lTOBSuiv             := FZSaisieVrac.Detail[i+1] ;
   CalculInfoSupp(FZSaisieVrac.Detail[i],lTOBPrec,lTOBSuiv) ;
  end; // for

 FRdSolde       := Arrondi(FRdSoldeDebit - FRdSoldeCredit , V_PGI.OkDecV ) ;

 FListe.Refresh ;

end;


procedure TOF_CPSAISIEVRAC.SoldeFolio ;
var
 ACol,ARow    : integer ;
 Cancel       : boolean ;
 lTOB         : TOB ;
 lRdSolde     : double ;
 lBoLigneVide : boolean ;
begin

 {$IFDEF TT}
  AddEvenement('SoldeFolio');
 {$ENDIF}


 ACol           := FListe.Col ;
 ARow           := FListe.Row ;

 //if FListe.Row = 1 then PGIInfo('La calcul de solde ne fonctionne pas sur la première ligne ! ', Ecran.Caption ) ;

 Cancel         := false ;
 FListeCellExit( nil,ACol,ARow,Cancel) ;
 if Cancel then exit ;

 FListe.Cells[SA_DEBIT  , FListe.Row] := '' ;
 FListe.Cells[SA_CREDIT , FListe.Row] := '' ;

 lTOB           := GetE(Fliste) ;

 CalculSoldeFolio(lTOB.GetValue('E_NUMGROUPEECR')) ;

 lBoLigneVide   := ( lTOB.GetValue('E_DEBIT') = 0 ) and ( lTOB.GetValue('E_CREDIT') = 0 ) ;

 if ( FRdSolde > 0 ) and lBoLigneVide then
  begin
   FListe.Cells[SA_CREDIT , FListe.Row] := StrFMontant(Abs(FRdSolde),15,FInfoEcr.Devise.Dev.Decimale,'',TRUE) ;
  end; // if

 if ( FRdSolde < 0 ) and lBoLigneVide then
  begin
   FListe.Cells[SA_DEBIT , FListe.Row] := StrFMontant(Abs(FRdSolde),15,FInfoEcr.Devise.Dev.Decimale,'',TRUE) ;
  end; // if


 if not lBoLigneVide then
  begin
   lRdSolde := FRdSolde - lTOB.GetValue('E_DEBIT') + lTOB.GetValue('E_CREDIT') ;
   if lRdSolde < 0 then
    begin
     FListe.Cells[SA_DEBIT , FListe.Row] := StrFMontant(Abs(lRdSolde),15,FInfoEcr.Devise.Dev.Decimale,'',TRUE) ;
     FRdSoldeDebit                       := FRdSoldeDebit  - Abs(lRdSolde) ;
    end
     else
      begin
       FListe.Cells[SA_CREDIT , FListe.Row] := StrFMontant(Abs(lRdSolde),15,FInfoEcr.Devise.Dev.Decimale,'',TRUE) ;
       FRdSoldeCredit                       := FRdSoldeCredit - Abs(lRdSolde) ;
      end;
  end; // if


 Statut         := taModif ;
 FBoGrilleModif := true ;
 FRdSolde       := Arrondi(0 , V_PGI.OkDecV ) ;

 // ligne correcte recalcul le solde
  if LigneCorrecte(FListe.Row) and ( FListe.Row = FListe.RowCount - 1 ) then
    NextRow ;

 AfficheSolde ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 10/10/2003
Modifié le ... : 08/06/2006
Description .. : - 10/10/2003 - FB 12890 - on respecte l'ordre de saisie
Suite ........ : - LG - 17/05/2004 -FB13206 - on passe la grille en consult 
Suite ........ : quand on lance l'integration
Suite ........ : - LG - 08/06/2006 - FB 18274 - on valide la ligne avant de 
Suite ........ : l'integrer
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.BIntegreClick(Sender : TObject);
var
 i            : integer ;
 ARow    : integer ;
 lChg,lCancel : boolean ;
begin

 FListe.PostDrawCell  := nil;
 FBoValidationEnCours := true ;

 lCancel := false ; lChg := false ;
 ARow:=FListe.Row ;
 FListeRowExit(nil,ARow,lCancel,lChg) ;
 if lCancel then exit ;

 Statut               := taConsult ;

 try

 // on supprime les pointeurs d'objets sur la grille, des enregistrements font etre supprimer du fait de l'integration
  for i := 0 to FListe.RowCount - 1 do
   SetE(FListe,nil,i) ;

  FZSaisieVrac.Numerote ;
  if not FZIntegration.RecupereLigne(FZSaisieVrac) then exit ;
  if Transactions(FZIntegration.Save,1) <> oeOK then
   begin
    PGIInfo('Erreur à l''enregistrement' + #10#13 + V_PGI.LastSQLError , Ecran.Caption ) ;
    exit ;
   end
    else
     if FZIntegration.FCR.count > 0 then
      begin
       AfficheCRIntegration (FZIntegration.FCR); // fct de libI\ImRapInt.pas
       FZIntegration.RAZFLC ;
       InitGrid ;
      end; // if

 finally
  FBoValidationEnCours := false ;
  Validation ;
  FListe.PostDrawCell  := FListePostDrawCell;
 end ;

end;


{ TZIntegrationVrac }


function _MakeCritN2 ( vTOB : TOB ) : string ;
var
 lStCritDate : string ;
begin

 if ( UpperCase(vTOB.GetValue('E_MODESAISIE')) = 'LIB' ) or ( UpperCase(vTOB.GetValue('E_MODESAISIE')) = 'BOR' ) then
  lStCritDate := intToStr(vTOB.GetValue('E_PERIODE'))
   else
    lStCritDate := DatetoStr(vTOB.GetValue('E_DATECOMPTABLE')) ;

 result := vTOB.GetValue('E_EXERCICE')                 +
           vTOB.GetValue('E_JOURNAL')                  +
           lStCritDate                                 +
           vTOB.GetValue('E_NATUREPIECE') ;
end;


constructor TZIntegrationVrac.Create( vInfoEcr : TInfoEcriture ) ;
begin
 inherited ;
 FTOBGarbage := TOB.Create('',nil,-1) ;
 FTOBLibN1   := TOB.Create('', FTOBGarbage, -1) ;
 FTOBLibN1.AddChampSupValeur('NIV',1);
 FTOBBorN1   := TOB.Create('', FTOBGarbage, -1) ;
 FTOBBorN1.AddChampSupValeur('NIV',1);
 FTOBPieceN1 := TOB.Create('', FTOBGarbage, -1) ;
 FTOBPieceN1.AddChampSupValeur('NIV',1);
 FTOBLibN3   := TOB.Create('', FTOBGarbage, -1) ;
 FTOBLibN3.AddChampSupValeur('NIV',3);
 _AddChampsSuppVrac(FTOBLibN3) ;
 FTOBBorN3   := TOB.Create('', FTOBGarbage, -1) ;
 FTOBBorN3.AddChampSupValeur('NIV',3);
 _AddChampsSuppVrac(FTOBBorN3) ;
 FTOBPieceN3 := TOB.Create('', FTOBGarbage, -1) ;
 FTOBPieceN3.AddChampSupValeur('NIV',3);
 _AddChampsSuppVrac(FTOBPieceN3) ;
 FCR := TList.Create ;
end;

destructor TZIntegrationVrac.Destroy;
begin
 RAZFLC ;
 FCR.Free ;
 FreeAndNil(FTOBGarbage) ;
 inherited;
end;


procedure TZIntegrationVrac.RAZFLC ;
var
 i       : integer ;
 ARecord : TCRPiece ;
begin
 for i := 0 to (FCr.Count - 1) do
  begin
   ARecord := FCR.Items[i];
   ARecord.Free;
  end; // for
 FCR.Clear ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/01/2004
Modifié le ... :   /  /    
Description .. : - 22/01/2004 - on reaffecte E_IO a '-' sur les ecritures non 
Suite ........ : integres
Mots clefs ... : 
*****************************************************************}
procedure TZIntegrationVrac.ReAffecteLigne ;
var
 i    : integer ;
 lTOBN2 : TOB ;
begin

 for i := 0  to FTOBGarbage.Detail.Count - 1 do
  begin
   lTOBN2 := FTOBGarbage.Detail[i] ; // N2

   while ( lTOBN2.Detail.Count > 0 ) and ( lTOBN2.Detail[0].GetValue('NIV') = 3 ) do
    begin
     lTOBN2.Detail[0].PutValue('E_QUALIFPIECE' , 'V' ) ;
     lTOBN2.Detail[0].PutValue('E_IO','-') ;
     lTOBN2.Detail[0].ChangeParent(FTOBN,-1) ;
    end ;
  end; // for

 FTOBN.Detail.Sort('E_NUMGROUPEECR;E_NUMEROPIECE;E_NUMLIGNE') ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 03/11/2003
Modifié le ... : 07/05/2004
Description .. : LG - 03/10/2003 - FB 12937 - changement du mode de 
Suite ........ : recuperation des ecritures. On ne recupere que les ecritures 
Suite ........ : affichées
Suite ........ : - LG - 07/05/2004 - FB 13206 - suppression des lignes 
Suite ........ : vides lors de la validation
Mots clefs ... : 
*****************************************************************}
function TZIntegrationVrac.RecupereLigne( vTOB : TOB ) : boolean;
var
 lTOB   : TOB ;
 lNbEcr : integer ;
 lIndex : integer ;
 i      : integer ;
begin

 FTOBN  := vTOB ;
 lNbEcr := FTOBN.Detail.Count ;
 lIndex := 0 ;
 i      := 0 ;
 result := false ;

 try

 while ( FTOBN.Detail.Count  >  0 ) and ( i < lNbEcr ) do
  begin
   lTOB := TTOBEcriture(FTOBN.detail[lIndex]) ;
   if ( lTOB.GetValue('GRILLE') = '-' ) or ( ( lTOB.GetValue('E_DEBIT') = 0 ) and ( lTOB.GetValue('E_CREDIT') = 0 ) ) then
    begin
     Inc(lIndex) ;
     Inc(i) ;
     continue ;
    end ;
   lTOB.AddChampSupValeur('NIV',3);
   lTOB.PutValue('E_QUALIFPIECE' , 'N' );
   lTOB.PutValue('E_IO', 'X') ;
   if ( lTOB.GetValue('E_MODESAISIE') = '' ) or ( lTOB.GetValue('E_MODESAISIE') = '-' ) then
    lTOB.ChangeParent(FTOBPieceN3,-1)
     else
      if ( lTOB.GetValue('E_MODESAISIE') = 'BOR' )  then
       lTOB.ChangeParent(FTOBBorN3,-1)
        else
         if ( lTOB.GetValue('E_MODESAISIE') = 'LIB' )  then
          lTOB.ChangeParent(FTOBLibN3,-1) ;
   Inc(i) ;
  end;

  except
   on E : Exception  do
    begin
     MessageAlerte('Erreur' + #10#13 + E.Message ) ;
     exit ;
    end ;
  end;

  FTOBPieceN3.Detail.Sort('E_EXERCICE;E_JOURNAL;E_DATECOMPTABLE;E_NATUREPIECE;E_NUMGROUPEECR;E_NUMLIGNE') ;
  FTOBBorN3.Detail.Sort('E_EXERCICE;E_JOURNAL;E_DATECOMPTABLE;E_NATUREPIECE;E_NUMGROUPEECR;E_NUMLIGNE') ;
  FTOBLibN3.Detail.Sort('E_EXERCICE;E_JOURNAL;E_DATECOMPTABLE;E_NUMEROPIECE;E_NUMGROUPEECR;E_NUMLIGNE') ;

  result := true ;

end;

function TZIntegrationVrac.ChercheN2dansN1 ( vTOBEcrN1 , vTOBEcrN2 : TOB ; var vInNumGroupeEcr : integer ) : TOB;
begin
 vInNumGroupeEcr := 1 ;
 result          := vTOBEcrN1.FindFirst(['CRIT'],[vTOBEcrN2.GetValue('CRIT')],false) ;
 if result = nil then
  result := TOB.Create('', vTOBEcrN1 , -1 )
   else
    if result.detail[result.Detail.Count-1].GetValue('E_MODESAISIE') = 'BOR' then
      vInNumGroupeEcr := result.detail[result.Detail.Count-1].GetValue('E_NUMGROUPEECR') + 1;
 result.AddChampSupValeur('CRIT',vTOBEcrN2.GetValue('CRIT')) ;
 result.AddChampSupValeur('NIV',2);
end ;

procedure TZIntegrationVrac.AjouteN2aN1( vTOBEcrN1 , vTOBEcrN2 : TOB ) ;
var
 lTOBN1 : TOB ;
 InNumGroupeEcr : integer ;
begin
 if ( UpperCase(vTOBEcrN2.Detail[0].GetValue('E_MODESAISIE')) = '-' ) or ( UpperCase(vTOBEcrN2.Detail[0].GetValue('E_MODESAISIE')) = '' ) then
  vTOBEcrN2.ChangeParent(FTOBPieceN1,-1)
   else
    begin
     lTOBN1 := ChercheN2dansN1(vTOBEcrN1,vTOBEcrN2,InNumGroupeEcr) ;
     while vTOBEcrN2.Detail.Count > 0 do
      begin
       vTOBEcrN2.Detail[0].PutValue('E_NUMGROUPEECR',InNumGroupeEcr) ;
       vTOBEcrN2.Detail[0].ChangeParent(lTOBN1,-1) ;
      end;
     vTOBEcrN2.Free ;
    end; // if
end;


procedure TZIntegrationVrac.AjouteN3aN1( vTOBEcrN1 , vTOBEcrN3 : TOB ) ;
var
 lRdSolde      : double ;
 lInLastCrit   : string ;
 lTOBN2        : TOB ;
 lRecErreur    : TRecError ;
 i             : integer ;
begin

 lInLastCrit := vTOBEcrN3.Detail[0].GetValue('E_NUMGROUPEECR') ;
 lRdSolde    := 0 ;
 lTOBN2      := TOB.Create('',FTOBGarbage,-1) ;
 lTOBN2.AddChampSupValeur('CRIT','-') ;
 lTOBN2.AddChampSupValeur('NIV',2);

 while vTOBEcrN3.Detail.Count <> 0 do
  begin

   if lInLastCrit = vTOBEcrN3.Detail[0].GetValue('E_NUMGROUPEECR')  then
    begin
     // calcul du solde
     lRdSolde := Arrondi(lRdSolde + vTOBEcrN3.Detail[0].GetValue('E_DEBIT') - vTOBEcrN3.Detail[0].GetValue('E_CREDIT') , V_PGI.OkDecV ) ;
     vTOBEcrN3.Detail[0].PutValue('SOLDE',lRdSolde) ;

     if lRdSolde = 0 then
      begin

       lTOBN2.PutValue('CRIT',_MakeCritN2(vTOBEcrN3.Detail[0])) ;
       vTOBEcrN3.Detail[0].ChangeParent(lTOBN2,-1) ;
       CAffectRegimeTva(lTOBN2) ;
       CAffectCompteContrePartie(lTOBN2,Info) ;
       lRecErreur := CIsValidSaisiePiece(lTOBN2,Info) ;
       if lRecErreur.RC_Error = RC_PASERREUR then // si le piece est valide
         AjouteN2aN1(vTOBEcrN1,lTOBN2) // on l'ajoute a la liste des pieces
          else
           for i := 0 to lTOBN2.Detail.Count - 1 do
            lTOBN2.Detail[i].PutValue('CODEERREUR',lRecErreur.RC_Error) ; // on la marque en erreur

       // on recreer une nouvelle pieces. la precedentes est sois ds N1 sois ds TOBGarbage
       lTOBN2 := TOB.Create('',FTOBGarbage,-1) ;
       lTOBN2.AddChampSupValeur('CRIT','-') ;
       lTOBN2.AddChampSupValeur('NIV',2);

      end // if
       else
        vTOBEcrN3.Detail[0].ChangeParent(lTOBN2,-1) ;
    end // if
     else
       begin
        lTOBN2      := TOB.Create('',FTOBGarbage,-1) ;
        lTOBN2.AddChampSupValeur('CRIT','-') ;
        lTOBN2.AddChampSupValeur('NIV',2);
        lRdSolde    := 0 ;
        lInLastCrit := vTOBEcrN3.Detail[0].GetValue('E_NUMGROUPEECR') ;
       end;

  end; // while

end;


procedure TZIntegrationVrac.Save ;
var
 lBoAucunEnr : boolean ;
 lStTitre    : string ;

 procedure _internalSave ( vTOBEcrN1, vTOBEcrN3 : TOB ) ;
 var
  i           : integer ;
 begin

  if vTOBEcrN3.Detail.Count <>  0 then
   begin

    AjouteN3aN1(vTOBEcrN1,vTOBEcrN3) ;

    for i := 0 to vTOBEcrN1.Detail.Count - 1 do
     begin
      lBoAucunEnr := false ;
      vTOBEcrN1.Detail[i].Detail[0].PutValue('E_QUALIFPIECE', 'V') ;  // astuce pour supprime l'analytique de la saisie en vrac ( plus propre que de mettre un parametre uniquement pour ce cas particulier )
      _CSuppAna(vTOBEcrN1.Detail[i].Detail[0]) ;
      vTOBEcrN1.Detail[i].Detail[0].PutValue('E_QUALIFPIECE', 'N') ;
      CAffectCompteContrePartie(vTOBEcrN1.Detail[i],Info) ;
       if CEnregistreSaisie(vTOBEcrN1.Detail[i],true,false,true,Info) then
         begin
          StockeLesEcrituresIntegrees(vTOBEcrN1.Detail[i]) ;
          MoveCurProgressForm('Intégration en cours...') ;
         end;
     end; // for

   end ;// if

 end;

begin
 lStTitre    := Ecran.Caption ;
 lBoAucunEnr := true ;
 InitMoveProgressForm(FEcran,'Génération en cours...','Génération en cours',FTOBBorN3.Detail.count+FTOBLibN3.Detail.count+FTOBPieceN1.Detail.count,true,false) ;

 try

  _internalSave(FTOBBorN1   , FTOBBorN3) ;
  _internalSave(FTOBLibN1   , FTOBLibN3) ;
  _internalSave(FTOBPieceN1 , FTOBPieceN3) ;
//  if lBoAucunEnr then
//   PGIInfo('Aucun enregistrement intégré',lStTitre) ;

 finally
  FTOBBorN1.ClearDetail ;
  FTOBLibN1.ClearDetail ;
  FTOBPieceN1.ClearDetail ;
  ReAffecteLigne ;
  FiniMoveProgressForm ;
 end ;

end;


procedure TZIntegrationVrac.StockeLesEcrituresIntegrees( vTOB : TOB);
var
 lLigneCR : TCRPiece ; // defini ds ImRapInt
 lTOB     : TOB ;
begin
 if vTOB.Detail.Count > 0 then
  begin
   lLigneCR             := TCRPiece.Create ;
   lTOB                 := vTOB.Detail[0] ;
   lLigneCR.NumPiece    := lTOB.GetValue('E_NUMEROPIECE') ;
   lLigneCR.Journal     := lTOB.GetValue('E_JOURNAL') ;
   lLigneCR.Date        := lTOB.GetValue('E_DATECOMPTABLE') ;
   lLigneCR.QualifPiece := lTOB.GetValue('E_QUALIFPIECE') ;
   lLigneCR.Debit       := vTOB.Somme('E_DEBIT', ['E_NUMEROPIECE'],[lLigneCR.NumPiece],TRUE);
   lLigneCR.Credit      := vTOB.Somme('E_CREDIT',['E_NUMEROPIECE'],[lLigneCR.NumPiece],TRUE);
   FCR.Add (lLigneCR);
  end; // for
end;



procedure TOF_CPSAISIEVRAC.BChercheClick(Sender: TObject);
begin
 ChercheClick ;
end;


function TOF_CPSAISIEVRAC.IsRowHT ( vRow : LongInt) : Boolean ;
begin
 result  := false ;
 if ( FListe.Cells[SA_GEN, vRow] = '') or ( Length(FListe.Cells[SA_GEN, vRow]) <> VH^.Cpta[fbGene].Lg)  then Exit ;
 if FInfoEcr.LoadCompte(FListe.Cells[SA_Gen,FListe.Row]) then
  result := FInfoEcr.Compte.isHT ;
end ;

function TOF_CPSAISIEVRAC.GetCompteAcc( Ou : integer ; E_AUXILIAIRE : string ) : string;
var
 i : integer ;

 function _GetContre : string ;
 var
  lTOB : TOB ;
  i : integer ;
 begin
   for i:=Ou-1 downto 1 do
      begin
       if E_AUXILIAIRE = FListe.Cells[SA_AUX, i]  then
        begin
         lTOB := GetE(FListe,i) ;
         if lTOB = nil then Exit ;
         result := lTOB.GetValue('E_CONTREPARTIEGEN') ;
         Break ;
       end;
      end ; // for
 end ;


begin

 result := '' ;

 if not FInfoEcr.Compte.IsCollectif then exit ;

 if ( FInfoEcr.GetString('YTC_ACCELERATEUR') = 'X' ) then
  begin
   result := VarToStr( FInfoEcr.GetString('YTC_SCHEMAGEN') ) ;
   if result = '' then result := _GetContre ;
  end  // if
   else
    begin
     result := _GetContre ;
     if result = '' then
      result:= VarToStr(FInfoEcr.GetString('YTC_SCHEMAGEN')) ;
    end ;

 if result = '' then
  begin
   if result = '' then result := FStGenHT ;
   if result = '' then
    for i := Ou - 1 downto 1 do
        begin
        if isRowHT(i) then
          begin
           result := FListe.Cells[SA_GEN, i] ;
           Break ;
          end;
        end ;
  end ;

end;


function TOF_CPSAISIEVRAC.AccelerateurDeSaisie : integer ;//( Ou : integer = - 1 ) : integer ;
var
 lTOB          : TTOBEcriture ;
 lEcrRef       : TTOBEcriture ;
 lStGen        : string ;
 lStNatureAuxi : string ;
 lStRegimeTVA  : string ;
 lStCompteTVA  : string ;
 lRdTauxTva    : double ;
 lBoMvtDebit   : boolean ;
 lRdVal        : double ;
 lRdTva        : double ;
 C             : double;
begin

 result := -1 ;

 if ( FRdSolde - Valeur(FListe.Cells[SA_DEBIT,FListe.Row]) + Valeur(FListe.Cells[SA_CREDIT,FListe.Row] ) ) <> 0 then exit ;

 lEcrRef := RecupTOB ; if lEcrRef = nil then exit ;
 FInfoEcr.Load(FListe.Cells[SA_Gen,FListe.Row],FListe.Cells[SA_Aux,FListe.Row],FListe.Cells[SA_Jal,FListe.Row]) ;

 if ( FInfoEcr.Journal.InIndex = -1 ) or ( FInfoEcr.Compte.InIndex = -1 ) or ( FInfoEcr.Aux.InIndex = -1 ) then exit ;
 if ( FInfoEcr.GetString('J_ACCELERATEUR') = '-' ) or ( lEcrRef.GetValue('ACC') = 'X' ) or ( FInfoEcr.GetString('YTC_ACCELERATEUR') = '-' ) then exit ;

 if not LigneCorrecte then exit ;

 lStGen  := GetCompteAcc( FListe.Row, lEcrRef.GetValue('E_GENERAL')) ;
 if lStGen = '' then exit ;

 lStNatureAuxi := '' ;
 lStRegimeTVA  := '' ;

 FInfoEcr.Compte.RecupInfoTVA (       lStGen,
                                      lEcrRef.GetValue('E_DEBIT') ,
                                      lEcrRef.GetValue('E_NATUREPIECE') ,
                                      FInfoEcr.GetString('J_NATUREJAL') ,
                                      FInfoEcr.GetString('T_NATUREAUXI') ,
                                      lStCompteTVA,
                                      lRdTauxTva,
                                      FInfoEcr.GetString('T_REGIMETVA') );

 lBoMvtDebit := lEcrRef.GetValue('E_DEBIT') <> 0 ;
 if lBoMvtDebit then C := lEcrRef.GetValue('E_DEBIT') else C := lEcrRef.GetValue('E_CREDIT') ;

 if lStCompteTVA <> '' then
  begin
   lRdVal := Arrondi(C / ( 1 + ( lRdTauxTva / 100 ) ),FInfoEcr.Devise.Dev.Decimale) ;
   lRdTva := Arrondi(C - lRdVal,FInfoEcr.Devise.Dev.Decimale) ;
  end
   else
    begin
     lRdVal := Arrondi(C ,FInfoEcr.Devise.Dev.Decimale) ;
     lRdTva := 0 ;
    end;

 // creation de la ligne sur le compte de contrepartie
 CreateRow(FListe.Row) ;
 lTOB := GetE(FListe) ;
 CDupliquerTOBEcr(lEcrRef,lTOB) ;
 lTOB.PutValue('E_GENERAL',lStGen) ;
 if lBoMvtDebit then
  CSetMontants( lTOB , 0 , lRdVal , FInfoEcr.Devise.Dev , true )
   else
    CSetMontants( lTOB , lRdVal , 0 , FInfoEcr.Devise.Dev , true ) ;
 lTOB.PutValue('E_LIBELLE'     , lEcrRef.GetValue('E_LIBELLE')) ;
 lTOB.PutValue('E_REFINTERNE'  , lEcrRef.GetValue('E_REFINTERNE')) ;
 RefreshLigne(FListe.Row-1) ;
 NextRow ;

 // creation de la ligne de tva
  if lStCompteTVA <> '' then
  begin
   CreateRow(FListe.Row) ;
   lTOB := GetE(FListe) ;
   CDupliquerTOBEcr(lEcrRef,lTOB) ;
   lTOB.PutValue('E_GENERAL',lStCompteTVA) ;
   if lBoMvtDebit then
    CSetMontants( lTOB , 0 , lRdTva , FInfoEcr.Devise.Dev , true )
     else
     CSetMontants( lTOB , lRdTva , 0 , FInfoEcr.Devise.Dev , true ) ;
   lTOB.PutValue('E_LIBELLE'     , lEcrRef.GetValue('E_LIBELLE')) ;
   lTOB.PutValue('E_REFINTERNE'  , lEcrRef.GetValue('E_REFINTERNE')) ;
   RefreshLigne(FListe.Row-1) ;
   NextRow ;
  end ; // if

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 24/11/2005
Modifié le ... :   /  /    
Description .. : - 24/11/2005 - LG - FB 17067 - les guides plantaient qd 
Suite ........ : l'etablissement n'etait aps affiché
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.NextColStop( var ACol , ARow : integer ) ;
var
 i,j         : integer ;
 lTOB        : TOB ;
 lStArret    : string ;
 lTOBGuide   : TOB  ;
begin

 if FBoAfficherEtablis then
  ACol := ACol - 3
  else
   ACol := ACol - 2 ;

 for i := ARow to FListe.RowCount - 1 do
  begin
   lTOB      := RecupTOB(i) ;
   lTOBGuide := FZGuide.LoadLigne(['NOR', lTOB.GetValue('_GUIDE')], lTOB.GetValue('_NUMLIGNE') ) ;
   if lTOBGuide = nil then exit ;
   for j := ACol to 6 do
    begin
     lStArret := copy(lTOBGuide.GetValue('EG_ARRET') ,1,6); // la grille de saisie ne comporte que 6 cases
     if ( length(lStArret) >= ACol ) and ( lStArret <> '' ) and ( trim(lStArret[j]) = 'X' ) then
      begin
       if FBoAfficherEtablis then
        ACol := j + 4  // Col
         else
          ACol := j + 3 ; // Col
       ARow := i ; // Row
       exit ;
      end ; // if
    end ; // for
    ACol := 1 ;
  end; // for
 // il n'y a plus de point d'arret, on se replace au debut
 ACol  := 1 ;
 ARow  := 1 ;

end;

procedure TOF_CPSAISIEVRAC.FListeCellExitModeGuide(Sender: TObject; var ACol, ARow : Integer; var Cancel: Boolean) ;
var
 lBoFirstCell : boolean ;
 lInRow       : integer ;
 i            : integer ;
begin

 if ( csDestroying in Ecran.ComponentState ) then Exit ;

 if not BoGuideRun then exit ;

 {$IFDEF TT}
  AddEvenement('FListeCellExitModeGuide ACol:='+intToStr(ACol)+' ARow:='+intToStr(ARow)) ;
 {$ENDIF}

 RecupTOB(ARow) ;

 Statut := taModif ;
 lInRow := ARow ;
 NextColStop( ACol , ARow ) ;
 lBoFirstCell := ( ARow = 1 ) and ( ACol = 1 ) ;
 Cancel       := not lBoFirstCell ;

 if ( ARow <> FListe.Row ) and not LigneCorrecte then
  Cancel := true
   else
    if Cancel then
     begin
      FListe.Col := ACol ;
      FListe.Row := ARow ;
     end ;

 HGBeginUpdate(FListe) ;
 try
  FZGuide.RecalculGuide ( FZSaisieVrac , lInRow ) ;
  for i := lInRow - 1 to ( FZGuide.InNumDep + FZGuide.InNbLigneGuide - 2 ) do
   RefreshLigne( i ) ;
  CompleteCeQueJePeux(SA_Lib,ARow ) ;

 if lBoFirstCell then
  begin
   CompleteCeQueJePeux(SA_Lib,FListe.Row ) ;
   FinGuide ;
  end ;

 finally
  HGEndUpdate(FListe) ;
 end ;

end;

procedure TOF_CPSAISIEVRAC.FinGuide ;
var
 lInIndex : integer ;
begin

 {$IFDEF TT}
  AddEvenement('FinGuide') ;
 {$ENDIF}

  lInIndex   := FZGuide.InNumDep + FZGuide.InNbLigneGuide - 1;

  HGBeginUpdate(FListe) ;
  BoGuideRun := false ;

  try

  while lInIndex >= FZGuide.InNumDep do
   begin
    if not LigneCorrecte(lInIndex) then
     FListeDeleteRow(lInIndex) ;
    Dec(lInIndex) ;
    Statut := taModif ;
   end ; // while

  finally
   HGEndUpdate(FListe) ;
   NextRow ;
  end ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 24/11/2005
Modifié le ... :   /  /    
Description .. : - 24/11/2005 - LG - FB 17067 - les guides plantaient qd 
Suite ........ : l'etablissement n'etait aps affiché
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPSAISIEVRAC.GuiClick ;
var
 lTOBEcr   : TOB ;
 lTOB      : TOB ;
 i         : integer ;
 lEcrRef   : TOB ;
 lInRowRef : integer ;
 lACol     : integer ;
begin
 if (FRdSolde <> 0) or BoGuideRun then exit ;
 lEcrRef   := RecupTOB ; if lEcrRef = nil then exit ;
 lTOBEcr   := TOB.Create('',nil,-1) ;
 lInRowRef := FListe.Row ;
 try
  if FZGuide.RechercheGuideEnBase(FZSaisieVrac,lTOBEcr,FListe) then
   begin
    Statut     := taModif ;
    BoGuideRun := true ;
    for i := 0 to lTOBEcr.Detail.Count - 1 do
     begin
      CreateRow(FListe.Row) ;
      lTOB   := GetE(FListe) ;
      if lTOB = nil then break ;
      lTOB.PutValue('E_GENERAL'       , lTOBEcr.Detail[i].GetValue('E_GENERAL') ) ;
      lTOB.PutValue('E_JOURNAL'       , lTOBEcr.Detail[i].GetValue('E_JOURNAL') ) ;
      lTOB.PutValue('E_AUXILIAIRE'    , lTOBEcr.Detail[i].GetValue('E_AUXILIAIRE') ) ;
      lTOB.PutValue('E_NATUREPIECE'   , lTOBEcr.Detail[i].GetValue('E_NATUREPIECE') ) ;
      lTOB.PutValue('E_REFINTERNE'    , lTOBEcr.Detail[i].GetValue('E_REFINTERNE') ) ;
      lTOB.PutValue('_GUIDE'          , lTOBEcr.Detail[i].GetValue('_GUIDE') ) ;
      lTOB.PutValue('_NUMLIGNE'       , lTOBEcr.Detail[i].GetValue('_NUMLIGNE') ) ;
      lTOB.PutValue('E_NUMGROUPEECR'  , - 1 );
      CRemplirDateComptable( lTOB     , lEcrRef.GetValue('E_DATECOMPTABLE') ) ;
      CSetMontants( lTOB,
                    lTOBEcr.Detail[i].GetValue('E_DEBIT')  ,
                    lTOBEcr.Detail[i].GetValue('E_CREDIT') ,
                    FInfoEcr.Devise.Dev ,
                    true );
      RefreshLigne(FListe.Row - 1 ) ;
      Statut := taConsult ; // pour passer ds le createRow qui verif qd l'on soie bien en consultation avant de creeer une ligne
     end ; // if
   end ; // if

   if FBoAfficherEtablis then
     lACol := 4
      else
       lACol := 3 ;
   NextColStop(lACol,lInRowRef);
   FListe.Col       := lACol ;
   FListe.Row       := lInRowRef ;
   Statut           := taModif ;
   FBoGrilleVide    := false ;

 finally
  lTOBEcr.Free ;
 end ;
end;

procedure TOF_CPSAISIEVRAC.SetGuideRun ( Value : boolean) ;
begin
 FBoGuideRun          := Value ;
 FlashGuide.Flashing  := Value ;
 FlashGuide.Visible   := Value ;
end;

Initialization
  registerclasses ( [ TOF_CPSAISIEVRAC ] ) ;
end.
