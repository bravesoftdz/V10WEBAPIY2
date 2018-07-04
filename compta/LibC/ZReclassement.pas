{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 03/03/2004
Description .. :
Suite ........ : GCO - 03/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Mots clefs ... :
*****************************************************************}
unit ZReclassement;

interface

uses
  Classes,
  Forms,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB,
  {$ENDIF}
  UTOB ,
  Ent1 , // pour le ListePeriode
  UlibEcriture,
{$IFDEF AMORTISSEMENT}
  ImmoCpte_Tom, // TesteExistenceComptesAssocies
  ImOutGen,     // TCompteAss
{$ENDIF}
  StdCtrls,
  Mask,
  Controls,
  HTB97,
  ExtCtrls,
  ComCtrls,
  ToolWin,
  Menus,
  ImgList,
  Hctrls,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HSysMenu, HImgList, TntStdCtrls
  ;

type

  TReclass  = ( trJal, trPeriode , trGen , trFolio , trEta , trExo, trGenLigne , trAux , trAuxLigne ) ;
  TTriTw    = ( ttExo, ttJal ) ; // stocke le mode de tri courant pour le TreeView, par defaut TriExo
  TTypeNode = ( tnExo , tnJal, tnPer ) ;

  TZreclassement  = Class ;

  // structure contenant les info pour la gestion de l'interface.
  // ces données sont rattachées au Data du TreeView et du ListView
  PInfoTW = ^RInfoTW;
  RInfoTW = record
   TExo            : TExoDate ;
   E_NUMECHE       : integer ;
   E_DATECOMPTABLE : TDatetime ;
   EX_ABREGE       : string ;
   InAnnee         : Word ;
   InMois          : Word ;
   E_JOURNAL       : string ;
   J_LIBELLE       : string ;
   J_NATUREJAL     : string ;
   J_CONTREPARTIE  : string ;
   E_QUALIFPIECE   : string ;
   E_MODESAISIE    : string ;
   J_VALIDEEN      : string ;
   J_VALIDEEN1     : string ;
   E_ETABLISSEMENT : string ;
   BoClos          : boolean ;
   BoValide        : boolean ;
   E_NUMEROPIECE   : integer ;
   E_NUMLIGNE      : integer ;
   InId            : integer ;
   BoVide          : boolean ;
   InNiv           : integer ;
   BoTriExo        : boolean ;
   BoASupprimer    : boolean ;
   TypeNode        : TTypeNode ;
   StInfoLibre     : string ;
   BoUtilise       : boolean ;
   E_IMMO          : string;
  end;


  TFReclassement = class(TForm)
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    IM32: THImageList;
    CoolBar1: TCoolBar;
    TW: TTreeView;
    POPB: TPopupMenu;
    Parjournal: TMenuItem;
    Parexercice: TMenuItem;
    CoolBar2: TCoolBar;
    BTree: TToolbarButton97;
    Splitter1: TSplitter;
    LW: TListView;
    POPC: TPopupMenu;
    MMJal: TMenuItem;
    MMGen: TMenuItem;
    MMFolio: TMenuItem;
    MMPer: TMenuItem;
    FParam: TToolWindow97;
    L1: TLabel;
    E1: THCritMaskEdit;
    L2: TLabel;
    E2: THCritMaskEdit;
    Panel1: TPanel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    FRech: TToolWindow97;
    E3: THCritMaskEdit;
    Label1: TLabel;
    BSuivantRech: TButton;
    BAnnulerRech: TButton;
    RadioGroup1: TRadioGroup;
    BChoix: TToolbarButton97;
    MMOuv: TMenuItem;
    N1: TMenuItem;
    HE: THValComboBox;
    BAff: TToolbarButton97;
    HMTrad: THSystemMenu;
    POPAFF: TPopupMenu;
    MMCodeExo: TMenuItem;
    N2: TMenuItem;
    MMGrand: TMenuItem;
    MMPetit: TMenuItem;
    MMAFFJAL: TMenuItem;
    N3: TMenuItem;
    BZoom: TToolbarButton97;
    POPZoom: TPopupMenu;
    MMZoomJal: TMenuItem;
    CBDuplic: TCheckBox;
    MMEta: TMenuItem;
    HV: THValComboBox;
    MMAna: TMenuItem;
    MMContr: TMenuItem;
    MMRecup: TMenuItem;
    MMsauv: TMenuItem;
    MMInsert: TMenuItem;
    MMAux: TMenuItem;
    L4: THLabel;
    E4: THCritMaskEdit;
    IM16: THImageList;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TWExpanding(Sender: TObject; Node: TTreeNode ; var AllowExpansion: Boolean);
    procedure ParexerciceClick(Sender: TObject);
    procedure ParjournalClick(Sender: TObject);
    procedure TWClick(Sender: TObject);
    procedure MMJalClick(Sender: TObject);
    procedure TWDragOver(Sender, Source: TObject; X, Y: Integer;State: TDragState; var Accept: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure TWDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure BFermeClick(Sender: TObject);
    procedure LWDblClick(Sender: TObject);
    procedure MMPerClick(Sender: TObject);
    procedure MMGenClick(Sender: TObject);
    procedure POPCPopup(Sender: TObject);
    procedure MMFolioClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
    procedure BSuivantRechClick(Sender: TObject);
    procedure BAnnulerRechClick(Sender: TObject);
    procedure TWChange(Sender: TObject; Node: TTreeNode);
    procedure MMOuvClick(Sender: TObject);
    procedure MMGrandClick(Sender: TObject);
    procedure MMPetitClick(Sender: TObject);
    procedure MMCodeExoClick(Sender: TObject);
    procedure MMAFFJALClick(Sender: TObject);
    procedure MMVerifClick(Sender: TObject);
    procedure MMZoomJalClick(Sender: TObject);
    procedure MMEtaClick(Sender: TObject);
    procedure E2Exit(Sender: TObject);
    procedure E1Exit(Sender: TObject);
    procedure MMAnaClick(Sender: TObject);
    procedure MMContrClick(Sender: TObject);
    procedure MMRecupClick(Sender: TObject);
    procedure MMsauvClick(Sender: TObject);
    procedure MMInsertClick(Sender: TObject);
    procedure E1ElipsisClick(Sender: TObject);
    procedure E2ElipsisClick(Sender: TObject);
    procedure E4ElipsisClick(Sender: TObject);
    procedure MMAuxClick(Sender: TObject);
    procedure E4Exit(Sender: TObject);
    procedure E2Change(Sender: TObject);
  private
    { Déclarations privées }
   FTri        : TTriTw ;         // stocke le mode de tri courant pour le TreeView, par defaut TriExo
   FCurrentId  : integer ;        // utilise pour identifie un noueds ds la treeview
   FTReclass   : TReclass ;       // type de reclassement : journal,compte
   FIndexRech  : integer ;        // index de depart ds le treeview pour la recherche
   FZReclass   : TZReclassement ; // objet gerant le reclassement
   FInfoEcr    : TInfoEcriture ;  // objet contenant la liste des journaux,... partage entre par l ZEtebac et les guides
   FMessCompta : TMessageCompta ; // affichage des messages
   FBoToutEta  : boolean ;        // FBoToutEta := true le reclassement se fait sur tout les etablissments ( on utilise pas les info du noued selectionne )
   FDtPer      : TDateTime ;      // stocke la date de destination lors d'une operation de drag and drop
   FStNatGene  : string ;
   FBoLW       : boolean ;
   procedure AjouteItemVide( vNode : TTreeNode) ;               // ajoute un noeud a supprimer par la suite au noeud courant, sert a faire apparaitre le '+' sur celui ci, le treeview etant remplis que qd on le selectionne
   procedure SuppItemVide( vNode : TTreeNode ) ;                // supprime le noeud a supprimer avant de rajoute un element un noeud
   procedure RemplirTreeviewExo ;                               // remplit le treeview avec l'ensemble des exo du dossiers
   procedure RemplirTreeviewPeriode( vRootNode : TTreeNode ) ;  // remplit le treeview avec l'ensemble des periodes de l'exo
   procedure RemplirTreeviewMoisJal( vRootNode : TTreeNode ) ;  // remplit le treeview avec l'ensemble des journal des periodes de l'exo
   procedure RemplirTreeviewJal( vRootNode : TTreeNode ) ;      // remplit le treeview avec l'ensemble des journaux de l'exo
   procedure RemplirTreeviewJalMois(vRootNode : TTreeNode) ;    // remplit le treeview avec l'ensemble des periodes de l'exo
   procedure RemplirListViewFolio(vRootNode : TTreeNode) ;
   procedure VideTW ;
   procedure VideLW ;
   procedure ChangeTri(vTri : TTriTw) ;
   procedure AddItemLW( vP : PInfoTW ; vCaption : string ) ;
   procedure RemplirLW( vNode : TTreeNode) ;
   procedure RemplirTW(vNode: TTreeNode);
   procedure InitParam( vPSource : PInfoTW ; vPDestination : PInfoTW = nil) ;
   procedure ToutSelectionnerLW;
   procedure Chercher;
   procedure EnabledControl;
   function  GetLibJal( vP : PInfoTW ) : string;
   function  GetLibExo( vP : PInfoTW ) : string;
   procedure ZoomJal;
   procedure OnErrorTOB(sender : TObject ; Error : TRecError );
   procedure ExecuteToutEta ;
   function  ExecuteTW : boolean ;
   function  ExecuteLW : boolean ;
  public
    { Déclarations publiques }
  end;

 TZreclassement = Class(TObjetCompta)
  private
   FListPiece       : TList ;
   FCR              : TList ;
   FPSource         : PInfoTW ;
   FTOBSource       : TOB ;
   FTOBDestination  : TOB ;
   FBoDuplication   : boolean ;
   FDtPeriodeDest   : TDateTime ;
   FStJournalDest   : string ;
   FStJournalSource : string ;
   FStEtaDest       : string ;
   FStEtaSource     : string ;
   FInFolioDest     : integer ;
   FStGenSource     : string ;
   FStGenDest       : string ;
   FStAuxDest       : string ;
   FStAuxSource     : string ;
   FTypeTrait       : TReclass ;
   {$IFDEF AMORTISSEMENT}
   FBoMajImmo       : Boolean;
   {$ENDIF}

   {$IFDEF TT}
   _flisteD         : tstringlist ;
   _flisteS         : tstringlist ;
   _flisteT         : tstringlist ;
   _flisteSQL       : tstringlist ;
   _FTOB            : TOB ;
   _FTOBAna         : TOB ;
   _FBoEnr          : boolean ;
   {$ENDIF}
   function  LoadParFolio : boolean;
   procedure ExecutePer ;
   procedure ExecuteEta ;
   procedure ExecuteFolio ;
   procedure ExecuteGen    ;
   function  ValideGen : boolean ;
   procedure ExecuteAux ;
   procedure ExecuteToutEta ;
   procedure RAZFLC;
   procedure StockeLesEcrituresIntegrees(vTOB: TOB ; vBoDetail : boolean = true) ;
   function  ExecuteTransfertJal(vTOB: TOB) : boolean ;
   procedure ExecuteJal ;
   procedure ExecuteJalBor ;
   procedure ExecuteJalPiece ;
   function  ControleBor : boolean ;
   procedure LoadParJournal( vP : PInfoTW ) ;
   procedure LoadParJal( vP : PInfoTW ) ;
   function  IsValidGen  ( lTOB : TOB ) : boolean ;

   {$IFDEF TT}
   procedure _RenumAna ;
   function  _ChargeBorSansAna(vP : PInfoTW ; vReclass : TReclass ) : TOB ;
   {$ENDIF}

  public
   constructor Create( vInfoEcr : TInfoEcriture ) ; override ;
   destructor  Destroy ; override ;
   procedure   Add( vP : PInfoTW ) ;
   function    Execute : boolean ;
   {$IFDEF TT}
   procedure    _ExecuteAna ( vP : PInfoTW ; vBoEnr : boolean = true ) ;
   {$ENDIF}
   function    IsValid( vPSource : PInfoTW = nil ) : boolean ;
   procedure   VideLaListe ;

   function  EstCorrectJournal : boolean ;
   function  EstCorrectCompteSource : boolean ;
   function  EstCorrectCompteDest : boolean ;

{$IFDEF AMORTISSEMENT}
   function  EstCorrectCompteImmo : boolean;
   function  TraitementImmo : Boolean;
{$ENDIF}

   function  EstCorrectAux : boolean ;

   property    DtPeriodeDest   : TDateTime read FDtPeriodeDest   write FDtPeriodeDest ;
   property    StJournalDest   : string    read FStJournalDest   write FStJournalDest ;
   property    StJournalSource : string    read FStJournalSource write FStJournalSource ;
   property    StEtaDest       : string    read FStEtaDest       write FStEtaDest ;
   property    StEtaSource     : string    read FStEtaSource     write FStEtaSource ;
   property    StGenDest       : string    read FStGenDest       write FStGenDest ;
   property    StAuxDest       : string    read FStAuxDest       write FStAuxDest ;
   property    StAuxSource     : string    read FStAuxSource     write FStAuxSource ;
   property    StGenSource     : string    read FStGensource     write FStGenSource ;
   property    InFolioDest     : integer   read FInFolioDest     write FInFolioDest ;
   property    BoDuplication   : boolean   read FBoDuplication   write FBoDuplication ;
   property    TypeTrait       : TReclass  read FTypeTrait       write FTypeTrait ;

 end ;


procedure CPLanceFiche_Reclassement ;
procedure CPInitPInfoTW ( var vP : PInfoTW ) ;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
   ParamSoc,
   ULibAnalytique,
   ImRapInt,
   ed_tools,
   CPJournal_TOM,
   hmsgbox, // pour le JaiLeDroitConcept
   SaisUtil, // pour QuellePeriode
   UtilSais,
   Saisie,
   HEnt1, // traduire memoire
   HPanel,
   UIUtil,
   SysUtils,
   Windows ,
   Lookup ;

const
 cImageBureau      = 0 ;
 cImageDoss        = 1 ;
 cImageDossSelec   = 2 ;
 cImageVide        = 3 ;
 cImageFeuille     = 7 ;
 cImageClot        = 9 ;
 cImageFolioClot   = 9 ;
 cImageValide      = 5 ;

 cStTitreTWDoss    = 'Dossier' ;
 cStTitreTWExo     = 'Exercice ' ;
 cStBordereau      = 'Bordereau ' ;
 cStPiece          = 'Pièce ' ;
 cStLibre          = 'Bor. libre ' ;
 cStJournalDeb     = 'Journal de départ' ;
 cStJournalArr     = 'Journal d''arrivée' ;
 cStPerDeb         = 'Période de départ' ;
 cStPerArr         = 'Période d''arrivée' ;
 cStGenDeb         = 'Compte de départ' ;
 cStGenArr         = 'Compte d''arrivée' ;
 cStEtaDeb         = 'Etablissement de départ' ;
 cStEtaArr         = 'Etablissement d''arrivée' ;
 cStPer            = '  Période : ' ;
 cStExo            = 'Exercice : ' ;
 cStJournal        = '  journal : ' ;
 cStExoClos        = 'Exercice clos  ' ;
 cStPeriodeVal     = 'Période validée ' ;

 cErrPasEnr        = 'Aucun Enregistrment sélectionné' ;
 cErrPasInfo       = 'Aucune info. de dispo pour ce noeud' ;
 cErrPasEnrT       = 'Aucun enregistrement ! ' ;
 cErrMvtValide     = 'Vous ne pouvez pas modifier ce bordereau,'+#10#13+'il contient des éléments pointés, lettrés ou validés !' ;
 cErrEquil         = 'Reclassement impossible : le folio n''est pas équilibré au jour' ;
 cErrParamJal      = 'Reclassement impossible : les journaux n''ont pas le même paramètrage multi-devises' ;
 cErrJalVal        = 'Reclassement impossible : le journal de destination est validé pour cette periode' ;
 cErrBorExist      = 'Reclassement impossible : ce numéro de bordereau existe déjà !' ;
 cErrMonoDev       = 'Reclassement impossible : le bordereau n''est pas mono devise !' ;
 cErrCompte        = 'Reclassement impossible : vous ne pouvez pas choisir de compte ventilable ou collectif !' ;
 cErrNatCompte     = 'Reclassement impossible : les comptes n''ont pas la même nature !' ;
// cErrLettCompte    = 'Reclassement impossible : les comptes n''ont pas les mêmes paramètres de lettrage !' ;
 cErrJalNat        = 'Reclassement impossible : la nature des journaux est différente !' ;
 cErrCompteContre  = 'Reclassement impossible : les comptes de contrepartie sont différents !' ;
 cErrCompteL       = 'Reclassement impossible : les comptes n''ont pas le même paramètrage de lettrage !' ;
 cErrCompteP       = 'Reclassement impossible : les comptes n''ont pas le même paramètrage de pointage !' ;
 cErrCompteI       = 'Reclassement impossible : les comptes n''ont pas le même paramètrage d''immobilisation !' ;
 cErrCompteIdent   = 'Reclassement impossible : les comptes ne peuvent pas être identiques !' ;
 cErrCompteInc     = 'Reclassement impossible : compte inconnu !' ;
 cErrAuxInc        = 'Reclassement impossible : auxiliaire inconnu !' ;
 cErrCompteDeC     = 'Reclassement impossible. Le compte sélectionné est le compte de contrepartie' ;
 cErrCompteAux     = 'Reclassement impossible. Le compte doit être un compte divers !' ;
 cErrCollectifObl  = 'Reclassement impossible. Le compte doit être un compte collectif !' ;
 cErrAucunErr      = 'Reclassement impossible. Aucun enregistrement ne correspond aux critères sélectionés' ;
 cErrCompteContreJ = 'Reclassement impossible sur le compte de contrepartie du journal !' ;
 cErrLigneValide   = 'Vous ne pouvez pas modifier cette ligne,elle contient des éléments validés !' ;
 cErrLignePointe   = 'Vous ne pouvez pas modifier cette ligne,elle contient des éléments pointés !' ;
 cErrLigneLettre   = 'Vous ne pouvez pas modifier cette ligne,elle contient des éléments lettrés !' ;

 cErrPrg           = - 100 ;
 cErrTransfert     = - 101 ;


function _AvecEchange : boolean ;
begin
  result :=( GetParamsocSecur('SO_CPLIENGAMME','') <> 'AUC' ) and ( GetParamsocSecur('SO_CPLIENGAMME','') <> '' ) ;
end ;

procedure CPLanceFiche_Reclassement ;
var
 lF : TFReclassement ;
 PP : THPanel ;
begin

if LienS1 then
 begin
  PGIInfo('Fonction non disponible !' +#10#13 + 'Vous êtes en synchro Business Line !','Reclassement');
  exit ;
 end;

if not BlocageMonoPoste(true) then
 begin
  PGIInfo('Vous devez etre seul a travailler sur ce dossier pour utiliser ce module') ;
  exit ;
 end ;

try

lF:=TFReclassement.Create(Application) ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     lF.ShowModal ;
   finally
     lF.Free ;
   end ;
 // Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(lF, PP) ;
  lF.Show ;
  end ;
finally
 DeblocageMonoPoste(true) ;
end;
end ;

function _FabriqTitre( vP : PInfoTW ; vTri : TTriTw ) : string ;
begin

 if vP = nil then exit ;

 if vP^.BoClos then
  result := cStExoClos
    else
     if vP^.BoValide then
       result := cStPeriodeVal ;

    result :=  result + cStExo + vP^.TExo.Code ;

   if vTri = ttExo then
    begin
     if ( vP^.InAnnee <> 0 ) and ( vP^.InMois <> 0 ) then
       result :=  result + cStPer + FormatDateTime('mmmm yyyy',EncodeDateBor(vP^.InAnnee,vP^.InMois,vP^.TExo)) ;
      if length(vP^.J_LIBELLE) > 0 then
        result :=  result + cStJournal + vP^.J_LIBELLE ;
    end // if
     else
      begin
       if length(vP^.J_LIBELLE) > 0 then
         result :=  result + cStJournal + vP^.J_LIBELLE ;
       if ( vP^.InAnnee <> 0 ) and ( vP^.InMois <> 0 ) then
         result :=  result + cStPer + FormatDateTime('mmmm yyyy',EncodeDateBor(vP^.InAnnee,vP^.InMois,vP^.TExo)) ;
      end;

end;

function _DecodeModeSaisie( vModeSaisie : string ) : string ;
begin
 if vModeSaisie = 'BOR' then
  result := cStBordereau
   else
    if ( vModeSaisie = '-' ) or ( vModeSaisie = '' ) then
     result := cStPiece
      else
       result := cStLibre ;

 result := TraduireMemoire(result) ;
end;

procedure _AffecteImageNode( vNode : TTreeNode ; vP : PInfoTW ) ;
begin
 if vP^.BoClos then
  begin
   vNode.SelectedIndex    := cImageClot ;
   vNode.ImageIndex       := cImageClot ;
  end
   else
    if vP^.BoValide then
     begin
      vNode.SelectedIndex    := cImageValide ;
      vNode.ImageIndex       := cImageValide ;
     end
      else
       begin
        vNode.SelectedIndex    := cImageDossSelec ;
        vNode.ImageIndex       := cImageDoss ;
       end;
end ;


procedure CPInitPInfoTW ( var vP : PInfoTW ) ;
begin
 vP^.TExo.Code         := '' ;
 vP^.InAnnee           := 0 ;
 vP^.InMois            := 0 ;
 VP^.TExo.Deb          := 0 ;
 vP^.TExo.Fin          := 0 ;
 vP^.E_JOURNAL         := '' ;
 vP^.J_LIBELLE         := '' ;
 vP^.J_NATUREJAL       := '' ;
 vP^.J_CONTREPARTIE    := '' ;
 vP^.E_QUALIFPIECE     := '' ;
 vP^.E_MODESAISIE      := '' ;
 vP^.J_VALIDEEN        := '' ;
 vP^.J_VALIDEEN1       := '' ;
 vP^.BoClos            := false ;
 vP^.BoValide          := false ;
 vP^.E_NUMEROPIECE     := 0 ;
 vP^.E_NUMLIGNE        := 0 ;
 vP^.InId              := -1 ;
 vP^.BoVide            := false ;
 vP^.InNiv             := -1 ;
 vP^.BoTriExo          := true ;
 vP^.BoASupprimer      := false ;
 vP^.TypeNode          := tnExo ;
 vP^.BoUtilise         := false ;
 vP^.E_Immo            := '';
end;


function _MakeSQLBor ( vP : PInfoTW ) : string ;
begin
 result := 'SELECT E_EXERCICE,E_JOURNAL,E_NUMEROPIECE,E_QUALIFPIECE,E_MODESAISIE,E_NUMECHE,E_DATECOMPTABLE,E_ETABLISSEMENT,E_VALIDE FROM ECRITURE ' +
           'WHERE E_EXERCICE="' + vP^.TExo.Code + '" ' +
           'AND E_JOURNAL="' + vP^.E_JOURNAL + '" ' +
           'AND E_NUMLIGNE=1 AND E_NUMECHE<=1 '  +
           'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDateBor(vP^.InAnnee, vP^.InMois , vP^.TExo))+'" ' +
           'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDateBor(vP^.InAnnee, vP^.InMois , vP^.TExo))) + '" ' +
           'AND E_QUALIFPIECE="N" ' +
           ' ORDER BY E_NUMEROPIECE' ;
end ;

function _IsValidJournalPeriode( vStExo , vStValideN , vStValideN1 : string ;  vDtDate : TDateTime ; vInfo : TInfoEcriture) : integer ;
var
 lBoExoN     : boolean ;
 lInIndexPer : integer ; // indice de la periode pour les champs valideen ( chaine de 24 car representant deux annees )
 lDtDate     : TDateTime ;
 lStValide   : string ;
begin

 result  := RC_PASERREUR ;

 lBoExoN := vStExo = vInfo.Exercice.EnCours.Code ;
 if ( not lBoExoN ) and ( vStExo <>vInfo.Exercice.Suivant.Code ) then Exit ; // on est pas sur l'exo en cours ni le suivant, il est donc forcement clos

 if lBoExoN then
  begin
   lStValide := vStValideN ;
   lDtDate   := vInfo.Exercice.EnCours.Deb ;
  end
   else
    begin
     lStValide := vStValideN1 ;
     lDtDate   := vInfo.Exercice.Suivant.Deb ;
    end ;

 lInIndexPer := QuellePeriode(vDtDate , lDtDate ) ;
 if ( lInIndexPer > 0 ) and ( lStValide[lInIndexPer] = 'X' ) then
  result := RC_JALVALID ;

end;




{ TFReclassement }
procedure TFReclassement.FormShow(Sender: TObject);
begin
 FInfoEcr            := TInfoEcriture.Create ;
 FZReclass           := TZReclassement.Create(FInfoEcr) ;
 FMessCompta         := TMessageCompta.Create(Self.Caption) ;
 FZReclass.OnError   := OnErrorTOB ;
 FInfoEcr.OnError    := OnErrorTOB ;
 RemplirTreeviewExo ;
 FParam.Visible     := false ;
 FCurrentId         := 0 ;
 EnabledControl ;
end;

procedure TFReclassement.FormClose(Sender: TObject; var Action: TCloseAction);
begin

 RegSaveToolbarPos(Self, 'ZReclassement') ;

 VideTW ;
 VideLW ;

 try
  if FInfoEcr    <> nil then FInfoEcr.Free ;
  if FMessCompta <> nil then FMessCompta.Free ;
  if FZReclass   <> nil then FZReclass.Free ;
 finally
  FInfoEcr    := nil ;
  FMessCompta := nil ;
  FZReclass   := nil ;
 end; //

end;

procedure TFReclassement.OnErrorTOB (sender : TObject; Error : TRecError ) ;
var
lStJournal : string ;
lStCompte  : string ;
begin
 if Error.RC_Error = RC_CINTERDIT then
  begin
   lStJournal := ReadtokenSt(Error.RC_Message) ;
   lStCompte  := ReadtokenSt(Error.RC_Message) ;
   PGIInfo(FMessCompta.GetMessage(Error.RC_Error) +#10#13 +  'Pour le journal ' + lStJournal + ' et le compte ' + lStCompte ) ;
  end
   else
    if ( trim(Error.RC_Message) <> '' ) then
     PGIInfo(Error.RC_Message , Self.Caption )
      else
       FMessCompta.Execute(Error.RC_Error) ;
end;


procedure TFReclassement.RemplirTreeviewExo ;
var
 lQ        : TQuery ;
 lNodeDoss : TTreeNode ;
 lNode     : TTreeNode ;
 lP        : PInfoTW ;
begin

 VideTW ;
 VideLW ;

 lNodeDoss            := TW.Items.Add(nil,TraduireMemoire(cStTitreTWDoss));
 lNodeDoss.ImageIndex := cImageBureau ;

 lQ := OpenSQL('SELECT EX_EXERCICE,EX_ETATCPTA,EX_VALIDEE,EX_ABREGE,EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE ' , true) ;
 inc(FCurrentId ) ;
 
 while not lQ.EOF do
  begin
   New(lP) ;
   CPInitPInfoTW(lP) ;
   lP^.TExo.Code          := lQ.findField('EX_EXERCICE').asString ;
   lP^.EX_ABREGE          := lQ.findField('EX_ABREGE').asString ;
   lP^.TExo.Deb           := lQ.findField('EX_DATEDEBUT').asDateTime ;
   lP^.TExo.Fin           := lQ.findField('EX_DATEFIN').asDateTime ;
   lP^.BoClos             := lQ.findField('EX_ETATCPTA').asString <> 'OUV' ;
   lP^.InNiv              := 0 ;
   lP^.BoTriExo           := true ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := FCurrentId ;
   lP^.TypeNode           := tnExo ;
   lNode                  := TW.Items.AddChildObject(lNodeDoss,GetLibExo(lP),lP) ;
   _AffecteImageNode(lNode,lP) ;
   AjouteItemVide(lNode) ;
   lQ.Next ;
  end ; // while

 Ferme(lQ) ;

end;

procedure TFReclassement.RemplirTreeviewPeriode( vRootNode : TTreeNode ) ;
var
 lQ        : TQuery ;
 lNode     : TTreeNode ;
 lPRoot    : PInfoTW ;
 lP        : PInfoTW ;
begin

 lPRoot := vRootNode.Data ;
 if not assigned(lPRoot) then exit ;

 lPRoot^.BoVide := false ;

 lQ := OpenSQL('SELECT YEAR(E_DATECOMPTABLE) ANNEE,MONTH(E_DATECOMPTABLE) MOIS FROM ECRITURE ' +
                'WHERE E_EXERCICE="' + lPRoot^.TExo.Code + '" ' +
                'AND E_QUALIFPIECE="N" ' +
                'GROUP BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE) ' +
                'ORDER BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE) ' , true ) ;

 if not lQ.Eof then SuppItemVide(vRootNode);
 inc(FCurrentId) ;

 while not lQ.EOF do
  begin
   New(lP) ;
   CPInitPInfoTW( lP ) ;
   lP^.TExo               := lPRoot^.TExo ;
   lP^.InAnnee            := lQ.findField('ANNEE').asInteger ;
   lP^.InMois             := lQ.findField('MOIS').asInteger ;
   lP^.InNiv              := 1 ;
   lP^.BoTriExo           := true ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := FCurrentId ;
   lP^.BoClos             := lPRoot^.BoClos ;
   lP^.TypeNode           := tnPer ;
   lP^.BoValide           := ( VH^.DateCloturePer > 0 ) and ( FinDeMois(EncodeDateBor(lP^.InAnnee, lP^.InMois,lP^.TExo )) <= VH^.DateCloturePer ) ;
   lNode                  := TW.Items.AddChildObject(vRootNode,FormatDateTime('mmmm yyyy',EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo)),lP) ;
   _AffecteImageNode(lNode,lP) ;
   AjouteItemVide(lNode) ;
   lQ.Next ;
  end ; // while

 Ferme(lQ) ;

end;


procedure TFReclassement.RemplirTreeviewJal( vRootNode : TTreeNode ) ;
var
 lQ        : TQuery ;
 lNode     : TTreeNode ;
 lPRoot    : PInfoTW ;
 lP        : PInfoTW ;
begin

 lPRoot := vRootNode.Data ;
 if not assigned(lPRoot) then exit ;

 lPRoot^.BoVide := false ;

 lQ := nil ;

 try

 lQ := OpenSQL('SELECT DISTINCT E_JOURNAL,J_LIBELLE,J_NATUREJAL,J_CONTREPARTIE,J_VALIDEEN,J_VALIDEEN1 FROM ECRITURE, JOURNAL ' +
               'WHERE E_JOURNAL=J_JOURNAL ' +
               'AND E_QUALIFPIECE="N" ' +
               'AND J_NATUREJAL<>"CLO" ' +
               'AND E_EXERCICE="' + lPRoot^.TExo.Code + '" ' , true ) ;

 if not lQ.Eof then SuppItemVide(vRootNode);
 inc(FCurrentId ) ;

 while not lQ.EOF do
  begin
   New(lP) ;
   CPInitPInfoTW( lP ) ;
   lP^.TExo               := lPRoot^.TExo ;
   lP^.InAnnee            := lPRoot^.InAnnee ;
   lP^.InMois             := lPRoot^.InMois ;
   lP^.E_JOURNAL          := lQ.findField('E_JOURNAL').asString ;
   lP^.J_LIBELLE          := lQ.findField('J_LIBELLE').asString ;
   lP^.J_NATUREJAL        := lQ.findField('J_NATUREJAL').asString ;
   lP^.J_VALIDEEN         := lQ.findField('J_VALIDEEN').asString ;
   lP^.J_VALIDEEN1        := lQ.findField('J_VALIDEEN1').asString ;
   lP^.J_CONTREPARTIE     := lQ.findField('J_CONTREPARTIE').asString ;
   lP^.InNiv              := 1 ;
   lP^.BoTriExo           := false ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := FCurrentId ;
   lP^.BoClos             := lPRoot^.BoClos ;
   lP^.BoValide           := lP^.J_NATUREJAL = 'ANO' ;
   lP^.TypeNode           := tnJal ;
   lNode                  := TW.Items.AddChildObject(vRootNode,GetLibJal(lP),lP) ;
   _AffecteImageNode(lNode,lP) ;
   AjouteItemVide(lNode) ;
   lQ.Next ;
  end ; // while

 finally
  Ferme(lQ) ;
 end;

end;


procedure TFReclassement.RemplirTreeviewMoisJal( vRootNode : TTreeNode ) ;
var
 lQ        : TQuery ;
 lNode     : TTreeNode ;
 lPRoot    : PInfoTW ;
 lP        : PInfoTW ;
begin

 lPRoot := vRootNode.Data ;
 if not assigned(lPRoot) then exit ;

 lPRoot^.BoVide := false ;

 lQ :=  nil ;

 try

 lQ := OpenSQL('SELECT DISTINCT E_JOURNAL,J_LIBELLE,J_NATUREJAL,J_VALIDEEN,J_VALIDEEN1,J_CONTREPARTIE FROM ECRITURE, JOURNAL ' +
                'WHERE E_JOURNAL=J_JOURNAL ' +
                'AND E_QUALIFPIECE="N" ' +
                'AND J_NATUREJAL<>"CLO" ' +
                'AND E_EXERCICE="' + lPRoot^.TExo.Code + '" ' +
                'AND YEAR(E_DATECOMPTABLE)=' + IntToStr(lPRoot^.InAnnee) +
                'AND MONTH(E_DATECOMPTABLE)=' + IntToStr(lPRoot^.InMois) , true ) ;

 if not lQ.Eof then SuppItemVide(vRootNode);
 inc(FCurrentId ) ;

 while not lQ.EOF do
  begin
   New(lP) ;
   CPInitPInfoTW( lP ) ;
   lP^.TExo               := lPRoot^.TExo ;
   lP^.InAnnee            := lPRoot^.InAnnee ;
   lP^.InMois             := lPRoot^.InMois ;
   lP^.E_JOURNAL          := lQ.findField('E_JOURNAL').asString ;
   lP^.J_LIBELLE          := lQ.findField('J_LIBELLE').asString ;
   lP^.J_NATUREJAL        := lQ.findField('J_NATUREJAL').asString ;
   lP^.J_VALIDEEN         := lQ.findField('J_VALIDEEN').asString ;
   lP^.J_VALIDEEN1        := lQ.findField('J_VALIDEEN1').asString ;
   lP^.J_CONTREPARTIE     := lQ.findField('J_CONTREPARTIE').asString ;
   lP^.InNiv              := 2 ;
   lP^.BoTriExo           := true ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := FCurrentId ;
   lP^.BoClos             := lPRoot^.BoClos ;
   lP^.BoValide           := lPRoot^.BoValide or ( lP^.J_NATUREJAL = 'ANO' ) ;
   lP^.TypeNode           := tnJal ;
   if not lP^.BoValide then  // si le journal est pas clot, on regarde s'il n'est pas valide pour cette periode
    lP^.BoValide          := _IsValidJournalPeriode( lP^.TExo.Code , lP^.J_VALIDEEN  , lP^.J_VALIDEEN1 , EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo), FInfoEcr ) <> RC_PASERREUR ;
   lNode                  := TW.Items.AddChildObject(vRootNode,GetLibJal(lP) ,lP) ;
   _AffecteImageNode(lNode,lP) ;
   lQ.Next ;
  end ; // while

 finally
  Ferme(lQ) ;
 end; // try

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : - LG - FB 13413 - 
Suite ........ : impossible de reclasser les écritures d'un journal de banque 
Suite ........ : à un autre journal de banque. On a pas les journaux de type 
Suite ........ : banque dans le combo de destination.
Mots clefs ... : 
*****************************************************************}
procedure TFReclassement.RemplirTreeviewJalMois( vRootNode : TTreeNode ) ;
var
 lQ        : TQuery ;
 lNode     : TTreeNode ;
 lPRoot    : PInfoTW ;
 lP        : PInfoTW ;
begin

 lPRoot := vRootNode.Data ;
 if not assigned(lPRoot) then exit ;

 lPRoot^.BoVide := false ;

 lQ := nil ;

 try

 lQ := OpenSQL('SELECT YEAR(E_DATECOMPTABLE) ANNEE,MONTH(E_DATECOMPTABLE) MOIS FROM ECRITURE ' +
                'WHERE E_EXERCICE="' + lPRoot^.TExo.Code + '" ' +
                'AND E_JOURNAL="' + lPRoot^.E_JOURNAL + '" ' +
                'AND E_QUALIFPIECE="N" ' +
                'GROUP BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE) ' +
                'ORDER BY YEAR(E_DATECOMPTABLE),MONTH(E_DATECOMPTABLE) ' , true ) ;

 if not lQ.Eof then SuppItemVide(vRootNode) ;
 inc(FCurrentId ) ;

 while not lQ.EOF do
  begin
   New(lP) ;
   CPInitPInfoTW( lP ) ;
   lP^.TExo               := lPRoot^.TExo ;
   lP^.InAnnee            := lQ.findField('ANNEE').asInteger ;
   lP^.InMois             := lQ.findField('MOIS').asInteger ;
   lP^.J_VALIDEEN         := lPRoot^.J_VALIDEEN ;
   lP^.J_VALIDEEN1        := lPRoot^.J_VALIDEEN1 ;
   lP^.E_JOURNAL          := lPRoot^.E_JOURNAL ;
   lP^.J_NATUREJAL        := lPRoot^.J_NATUREJAL ;
   lP^.J_CONTREPARTIE     := lPRoot^.J_CONTREPARTIE ;
   lP^.InNiv              := 2 ;
   lP^.BoTriExo           := true ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := FCurrentId ;
   lP^.BoClos             := lPRoot^.BoClos ;
   lP^.BoValide           := lPRoot^.BoValide ;
   lP^.TypeNode           := tnPer ;
   if not lP^.BoValide then
    lP^.BoValide          := ( _IsValidJournalPeriode( lP^.TExo.Code , lP^.J_VALIDEEN  , lP^.J_VALIDEEN1 , EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo),FInfoEcr )  <> RC_PASERREUR ) or ( lP^.J_NATUREJAL = 'ANO' ) ;
   lNode                  := TW.Items.AddChildObject(vRootNode,FormatDateTime('mmmm yyyy',EncodeDateBor(lP^.InAnnee,lP^.InMois,lP^.TExo)),lP) ;
   _AffecteImageNode(lNode,lP) ;
   lQ.Next ;
  end ; // while

 finally
  Ferme(lQ) ;
 end ;

end;

procedure TFReclassement.RemplirListViewFolio( vRootNode : TTreeNode ) ;
var
 lQ        : TQuery ;
 lPRoot    : PInfoTW ;
 lP        : PInfoTW ;
 {$IFDEF TT}
// lTOB      : TOB ;
// lTOBBor   : TOB ;
/// i         : integer ;
 {$ENDIF}
begin

 lPRoot := vRootNode.Data ;
 if not assigned(lPRoot) then exit ;

 {$IFDEF TT}
// lTOBBor := CGetListeBordereauBloque ;
 {$ENDIF}

 LW.Items.BeginUpdate ;
 VideLW ;
 lQ := nil ;

 try

 lPRoot^.BoVide    := false ;

 lQ := OpenSQL( _MakeSQLBor(lPRoot) , true ) ;
 {$IFDEF TTA}
 lTOB := TOB.Create('',nil,-1) ;
 lTOB.LoadDetailDB('ECRITURE','','',lQ,true ) ;
 Ferme(lQ) ;
 for i := 0  to lTOB.Detail.Count - 1 do
  begin
   New(lP) ;
   CPInitPInfoTW( lP ) ;
   lP^.TExo               := lPRoot^.TExo ;
   lP^.InAnnee            := lPRoot^.InAnnee ;
   lP^.InMois             := lPRoot^.InMois ;
   lP^.E_JOURNAL          := lPRoot^.E_JOURNAL ;
   lP^.J_NATUREJAL        := lPRoot^.J_NATUREJAL ;
   lP^.BoClos             := lPRoot^.BoClos ;
   lP^.BoValide           := lPRoot^.BoValide ;
   lP^.E_NUMEROPIECE      := lTOB.Detail[i].GetValue('E_NUMEROPIECE') ;
   lP^.E_QUALIFPIECE      := lTOB.Detail[i].GetValue('E_QUALIFPIECE') ;
   lP^.E_MODESAISIE       := lTOB.Detail[i].GetValue('E_MODESAISIE') ;
   lP^.E_NUMECHE          := lTOB.Detail[i].GetValue('E_NUMECHE') ;
   lP^.E_DATECOMPTABLE    := lTOB.Detail[i].GetValue('E_DATECOMPTABLE') ;
   lP^.E_ETABLISSEMENT    := lTOB.Detail[i].GetValue('E_ETABLISSEMENT') ;
   lP^.InNiv              := 3 ;
   lP^.BoTriExo           := true ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := lPRoot^.InId ;
   lP^.BoUtilise          := CEstBloqueEcritureBor ( lTOB.Detail[i] , lTOBBor ) ;
   AddItemLW(lP,_DecodeModeSaisie(lP^.E_MODESAISIE) + intToStr(lP^.E_NUMEROPIECE ) ) ;
  end ; // for
  lTOB.Free ;

  finally
   LW.Items.EndUpdate ;
  end;

 {$ELSE}
 while not lQ.EOF do
  begin
   New(lP) ;
   CPInitPInfoTW( lP ) ;
   lP^.TExo               := lPRoot^.TExo ;
   lP^.InAnnee            := lPRoot^.InAnnee ;
   lP^.InMois             := lPRoot^.InMois ;
   lP^.E_JOURNAL          := lPRoot^.E_JOURNAL ;
   lP^.J_NATUREJAL        := lPRoot^.J_NATUREJAL ;
   lP^.BoClos             := lPRoot^.BoClos ;
   lP^.BoValide           := lPRoot^.BoValide or ( lQ.findField('E_VALIDE').asString ='X' ) ;
   lP^.E_NUMEROPIECE      := lQ.findField('E_NUMEROPIECE').asInteger ;
   lP^.E_QUALIFPIECE      := lQ.findField('E_QUALIFPIECE').asString ;
   lP^.E_MODESAISIE       := lQ.findField('E_MODESAISIE').asString ;
   lP^.E_NUMECHE          := lQ.findField('E_NUMECHE').asInteger ;
   lP^.E_DATECOMPTABLE    := lQ.findField('E_DATECOMPTABLE').asDateTime ;
   lP^.E_ETABLISSEMENT    := lQ.findField('E_ETABLISSEMENT').asString ;
   lP^.InNiv              := 3 ;
   lP^.BoTriExo           := true ;
   lP^.BoVide             := true ;
   lP^.BoASupprimer       := false ;
   lP^.InId               := lPRoot^.InId ;
   AddItemLW(lP,_DecodeModeSaisie(lP^.E_MODESAISIE) + intToStr(lP^.E_NUMEROPIECE ) ) ;
   lQ.Next ;
  end ; // while

 finally
  LW.Items.EndUpdate ;
  Ferme(lQ) ;
 end;

{$ENDIF}
end;

procedure TFReclassement.AjouteItemVide( vNode : TTreeNode) ;
var
 lP     : PInfoTW ;
 lNode  : TTreeNode ;
begin
 New(lP) ;
 CPInitPInfoTW( lP ) ;
 lP^.InNiv              := 0 ;
 lP^.BoVide             := true ;
 lP^.BoASupprimer       := true ;
 lNode                  := TW.Items.AddChildObject(vNode,'<Vide>' ,lP) ;
 lNode.SelectedIndex    := cImageVide ;
 lNode.ImageIndex       := cImageVide ;
end;

procedure TFReclassement.VideTW ;
var
 i  : integer ;
 lP : PInfoTW ;
begin
 for i := 0 to TW.Items.Count - 1  do
  begin
   lP := (TW.Items[i].Data) ;
   if lP <> nil  then
    begin  // correction d'une fuite memoire
     Dispose(lP) ; TW.Items[i].Data := nil ;
    end ;
  end ;
 TW.Items.Clear ;
end;

procedure TFReclassement.SuppItemVide( vNode : TTreeNode);
var
 lP    : PInfoTW ;
 lNode : TTreeNode ;
begin
 lNode := vNode.getFirstChild ;
 if lNode = nil then exit ;
 lP := lNode.Data ;
 if not assigned(lP) then exit ;
 if lP^.BoASupprimer then
  begin
   Dispose(lP) ;
   lNode.Delete ;
  end; // if
end;

procedure TFReclassement.RemplirTW( vNode : TTreeNode) ;
var
 lP : PInfoTW ;
begin
 lP := vNode.Data ;
 if not assigned(lP) then exit ;
 if not lP^.BoVide then exit ;
 if ( lP^.InNiv = 0  ) and ( FTri = ttExo ) then RemplirTreeviewPeriode(vNode) ;
 if ( lP^.InNiv = 0  ) and ( FTri = ttJal ) then RemplirTreeviewJal(vNode) ;
 if ( lP^.InNiv = 1  ) and ( FTri = ttExo ) then RemplirTreeviewMoisJal(vNode) ;
 if ( lP^.InNiv = 1  ) and ( FTri = ttJal ) then RemplirTreeviewJalMois(vNode) ;
end;


procedure TFReclassement.TWExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean) ;
begin
 if Node.Data = nil then exit ;
 RemplirTW(Node) ;
end;

procedure TFReclassement.TWClick(Sender: TObject);
begin
 if TW.Selected = nil then exit ;
 RemplirLW(TW.Selected) ;
end;

procedure TFReclassement.LWDblClick(Sender: TObject);
var
 lItem      : TListItem ;
 lP         : PInfoTW ;
 lQ         : TQuery ;
 lStCols    : string ;
 lStWhere   : string ;
 lStOrderBy : string ;
begin

 if LW.SelCount > 1 then exit ;

 lItem      := LW.Selected ;
 if ( lItem = nil ) or ( lItem.Data = nil ) then exit ;
 lP         := PInfoTW( lItem.Data ) ;

 lQ         := nil ;

 lStCols    := 'E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE' ;

 lStWhere   := 'E_DATECOMPTABLE>="' + USDateTime(EncodeDateBor(lP^.InAnnee, lP^.InMois, lP^.TExo)) + '" ' +
               ' AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDateBor(lP^.InAnnee, lP^.InMois, lP^.TExo))) + '"' +
               ' AND E_JOURNAL="' + lP^.E_JOURNAL + '"' +
               ' AND E_NUMEROPIECE='+ IntToStr(lP^.E_NUMEROPIECE) +
               ' AND E_NUMLIGNE=1 AND E_NUMECHE<=1 ' +
               ' AND E_QUALIFPIECE="' + lP^.E_QUALIFPIECE + '" ' ;

 lStOrderBy := 'E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE' ;

 try
  lQ         :=OpenSQL('SELECT ' + lStCols + ' FROM ECRITURE WHERE ' + lStWhere + ' ORDER BY ' + lStOrderBy , TRUE) ;
// if (GS.Cells[SC_TYPEJAL, GS.Row]<>'-') and (GS.Cells[SC_TYPEJAL, GS.Row]<>'') then
//   LanceSaisieFolio(Q, taModif)
 //    else
      TrouveEtLanceSaisie(lQ, taModif, 'N') ;
 finally
  ferme(lQ) ;
 end;

end;


procedure TFReclassement.RemplirLW( vNode : TTreeNode ) ;
var
 i : integer ;
begin

 if ( vNode = nil ) or ( vNode.Data = nil )   then exit ;

 // on remplit le treeview
 RemplirTW(vNode) ;

 // qd on est au dernier niv du treeview
 if PInfoTW(vNode.Data).InNiv = 2 then
  RemplirListViewFolio(vNode) // on charge les info sur le folio
   else
    begin
     // on recopie de le treeview ds le listview
     LW.Items.BeginUpdate ;
     VideLW ;
     try
      for i:=0 to vNode.Count - 1 do
        AddItemLW( vNode.Item[i].Data , vNode.Item[i].Text ) ;
     finally
      LW.Items.EndUpdate ;
     end; // try
    end; // if

end;


procedure TFReclassement.AddItemLW( vP : PInfoTW ; vCaption : string ) ;
var
 lItem  : TListItem;
begin
 if vP^.BoASupprimer then exit ;
 lItem            := LW.Items.Add ;
 if vP^.TypeNode = tnJal then
  lItem.Caption := vP^.J_LIBELLE
   else
    lItem.Caption := vCaption ;
 lItem.Data       := vP ;
 if vP^.BoUtilise then
 lItem.ImageIndex := cImageValide
 else
 case vP^.InNiv of
  0..2 : if vP^.BoClos then
          lItem.ImageIndex := cImageClot
           else
            if vP^.BoValide then
             lItem.ImageIndex := cImageValide
              else
               lItem.ImageIndex := cImageDoss ;
  3    : if vP^.BoClos or vP^.BoValide then
          lItem.ImageIndex := cImageFolioClot
           else
            lItem.ImageIndex := cImageFeuille ;
 end; // case
end;

procedure TFReclassement.ParexerciceClick( Sender: TObject ) ;
begin
 ChangeTri(ttExo) ;
end;

procedure TFReclassement.ParjournalClick( Sender : TObject ) ;
begin
 ChangeTri(ttJal) ;
end;

procedure TFReclassement.ChangeTri( vTri : TTriTw ) ;
begin
 if FTri = vTri then exit ;
 FTri := vTri ;
 Parexercice.Checked := FTri = ttExo ;
 ParJournal.Checked  := FTri = ttJal ;
 RemplirTreeviewExo ;
end;


procedure TFReclassement.VideLW ;
var
 i  : integer ;
 lP : PInfoTW ;
begin
 for i := 0 to LW.Items.Count - 1 do
  begin
   lP := (LW.Items[i].Data) ;
    if assigned(lP)  and ( lP^.InNiv = 3 ) then
     begin // LG correction d'une fuite memoire
      Dispose(lP) ;
      LW.Items[i].Data := nil ;
     end ;
  end ;
 LW.Items.Clear ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 22/12/2004
Modifié le ... : 02/09/2005
Description .. : - LG - 22/12/2004 - FB 13412 - on empeche de 
Suite ........ : reselectionner le meme compte
Suite ........ : - LG - 02/09/2005 - FB 16264 - limitation du chiffre max 
Suite ........ : pour la numero de folio
Mots clefs ... : 
*****************************************************************}
procedure TFReclassement.InitParam( vPSource : PInfoTW ; vPDestination : PInfoTW = nil) ;
begin

 E2.Text            := '' ;
 E2.EditMask        := '' ;
 E2.ElipsisButton   := true ;
 E2.Enabled         := true ;
 E2.DataType        := '' ;
 E2.OpeType         := otString ;
 E1.Enabled         := false ;
 E1.ElipsisButton   := false ;
 E2.Visible         := true ;
 E2.Plus            := '' ;
 HE.Visible         := false ;
 TW.Enabled         := false ;
 LW.Enabled         := false ;
 E1.Visible         := true ;
 HV.Visible         := false ;
 E4.Visible         := false ;
 L4.Visible         := false ;
 E2.ElipsisButton   := true ;
 E2.MaxLength       := 35 ;
 E4.Text            := '' ;
 E2.OnExit          := nil ;
 E1.OnExit          := nil ;
 E4.OnExit          := nil ;
 E2.OnElipsisClick  := nil ;
 E1.OnElipsisClick  := nil ;
 E4.OnElipsisClick  := E4ElipsisClick ;

 try

 FParam.Visible := true ;

 if FTReclass = trJal then
  begin
   L1.Caption             := TraduireMemoire(cStJournalDeb) ;
   E1.Text                := vPSource^.E_JOURNAL ;
   L2.Caption             := TraduireMemoire(cStJournalArr) ;
   E2.DataType            := 'TTJALCRIT' ;
   if vPSource^.J_CONTREPARTIE = '' then
    E2.Plus  := ' J_JOURNAL<>"' + vPSource^.E_JOURNAL + '" AND ( J_NATUREJAL="OD" OR J_NATUREJAL="' + vPSource^.J_NATUREJAL + '") '
    else
     E2.Plus  := ' J_JOURNAL<>"' + vPSource^.E_JOURNAL + '" AND ( ( J_NATUREJAL="' + vPSource^.J_NATUREJAL + '" ' +
                 ' AND J_CONTREPARTIE="' +  vPSource^.J_CONTREPARTIE + '" ) OR J_NATUREJAL="OD" ) ' ;

   if vPDestination <> nil then
    begin
     E2.Text              := vPDestination^.E_JOURNAL ;
     E2.Enabled           := false ;
     E2.ElipsisButton     := false ;
    end; // if
   if E2.canFocus then E2.Setfocus ;
  end  // if TReclassJal
   else
     if FTReclass = trPeriode then
      begin
       L1.Caption           := TraduireMemoire(cStPerDeb) ;
       E1.Text              := FormatDateTime('mmmm yyyy',EncodeDateBor(vPSource^.InAnnee,vPSource^.InMois,vPSource^.TExo)) ;
       L2.Caption           := TraduireMemoire(cStPerArr) ;
       if vPDestination <> nil then
        begin
         E2.EditMask        := '' ;
         E2.ElipsisButton   := false ;
         E2.Text            := FormatDateTime('mmmm yyyy',EncodeDateBor(vPDestination^.InAnnee,vPDestination^.InMois,vPDestination^.TExo)) ;
         FDtPer             := EncodeDateBor(vPDestination^.InAnnee, vPDestination^.InMois, vPDestination^.TExo) ;
         E2.Enabled         := false ;
        end // if
         else
           begin
            ListePeriode(vPSource^.TExo.Code, HE.Items, HE.Values, true, false) ;
            E2.Visible       := false ;
            HE.Visible       := true ;
            HE.Style         := csDropDownList ;
           end

      end  // if  TReclassPeriode
       else
        if FTReclass = trGen then
          begin
           L1.Caption           := TraduireMemoire(cStGenDeb) ;
           E1.Enabled           := true ;
           if E1.canFocus then E1.Setfocus ;
           E1.ElipsisButton     := true ;
           E1.Text              := '' ;
           L2.Caption           := TraduireMemoire(cStGenArr) ;
           E2.DataType          := 'TZGENERAL' ;
           E1.DataType          := 'TZGENERAL' ;
           E2.OnExit            := E2Exit ;
           E1.OnExit            := E1Exit ;
           E2.OnElipsisClick    := E2ElipsisClick ;
           E1.OnElipsisClick    := E1ElipsisClick ;
          end  // if  TReclassPeriode
           else
            if FTReclass = trFolio then
              begin
               L1.Caption           := TraduireMemoire(_DecodeModeSaisie(vPSource^.E_MODESAISIE) + ' de départ') ;
               E1.Text              := intToStr(vPSource^.E_NUMEROPIECE) ;
               E2.Visible           := true ;
               E1.Visible           := true ;
               E2.ElipsisButton     := false ;
               E2.MaxLength         := 6 ;
               L2.Caption           := TraduireMemoire(_DecodeModeSaisie(vPSource^.E_MODESAISIE) + ' d''arrivé') ;
               if E2.canFocus then E2.Setfocus ;
              end
               else
                if FTReclass = trEta then
                  begin
                  if FBoToutEta then
                   begin
                     HV.Visible         := true ;
                     HV.ItemIndex       := 0 ;
                     E1.Visible         := false ;
                   end
                    else
                     begin
                      HV.Visible        := false ;
                      E1.Visible        := true ;
                      E1.Text           := RechDom('TTETABLISSEMENT',vPSource^.E_ETABLISSEMENT,false) ;
                      E1.Enabled        := false ;
                     end ;
                   L1.Caption           := TraduireMemoire(cStEtaDeb) ;
                   L2.Caption           := TraduireMemoire(cStEtaArr) ;
                   HE.Visible           := true ;
                   HE.Style             := csDropDownList ;
                   HE.Items.Clear ;
                   HE.Values.Clear ;
                   HE.DataType          := 'TTETABLISSEMENT' ;
                   if HE.CanFocus and not FBoToutEta then HE.SetFocus ;
                   HE.ReLoad;
                  end
                   else
                    if ( FTReclass = trAux ) then
                     begin
                      L1.Caption           := TraduireMemoire(cStGenDeb) ;
                      E1.Enabled           := true ;
                      if E1.canFocus then E1.Setfocus ;
                      E1.ElipsisButton     := true ;
                      E1.Text              := '' ;
                      L2.Caption           := TraduireMemoire(cStGenArr) ;
                      E2.DataType          := 'TZGENERAL' ;
                      E1.DataType          := 'TZGENERAL' ;
                      E2.OnExit            := E2Exit ;
                      E1.OnExit            := E1Exit ;
                      E4.OnExit            := E4Exit ;
                      E2.OnElipsisClick    := E2ElipsisClick ;
                      E1.OnElipsisClick    := E1ElipsisClick ;
                      E4.Visible           := true ;
                      L4.Visible           := true ;
                     end ;


 BChoix.Enabled := false ;
 BAff.Enabled   := false ;
 BZoom.Enabled  := false ;
 BTree.Enabled  := false ;

 except
  on e:exception do
   begin
    MessageAlerte(e.message ) ;
    raise ;
   end ;
 end ;

end;

procedure TFReclassement.MMJalClick(Sender: TObject);
begin
 if ( TW.Selected = nil ) or ( TW.Selected.Data = nil ) then exit ;
 FTReclass := trJal ;
 InitParam(PInfoTW(TW.Selected.Data)) ;
end;

procedure TFReclassement.ExecuteToutEta ;
begin
 FBoToutEta             := false ;
 FZReclass.StEtaSource  := HV.Value ;
 FZReclass.StEtaDest    := HE.Value ;
 if Transactions(FZReclass.ExecuteToutEta,1) = oeOk then
  begin
   RemplirTreeviewExo ;
   PGIInfo('Traitement réussi','Reclassement des données' ) ;
  end ;
end;

function TFReclassement.ExecuteTW : boolean ;
var
 lP  : PInfoTW ;
// lNb : integer ;
begin

 result := false ;
 if TW.Selected = nil then exit ;
 lP := PInfoTW(TW.Selected.Data) ;
 if lP = nil then exit ;

 if ( lP^.InNiv = 1 ) then
  FZReclass.LoadParJal(lP)
   else
    FZReclass.LoadParJournal(lP) ;

 FZReclass.TypeTrait := FTReclass ;
 result              := FZReclass.Execute ;

 if result then
  RemplirTreeviewExo ;

end;

function TFReclassement.ExecuteLW : boolean ;
var
 i : integer ;
// lNb : integer ;
begin
 FZReclass.VideLaListe ;
 for i := 0 to LW.Items.Count - 1 do
  if LW.Items[i].Selected then
   FZReclass.Add(LW.Items[i].Data) ;
 FZReclass.TypeTrait := FTReclass ;
 result := FZReclass.Execute ;
 if result and ( ( FTReclass <> trGen ) and ( FTReclass <> trAux ) ) then
  RemplirTreeviewExo ;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/09/2004
Modifié le ... :   /  /
Description .. : - LG - 27/09/2004 - FB 14648 - on pouvait reclasser sur des
Suite ........ : folio a zero
Mots clefs ... :
*****************************************************************}
procedure TFReclassement.BValiderClick(Sender: TObject) ;
var
 lP    : PInfoTW ;
begin

 FZReclass.BoDuplication := CBDuplic.Checked ;

 try

 // récupération des valeurs saisies
 if FTReclass = trPeriode then
  begin
   if HE.Visible and ( HE.Text = '' ) then exit ;
   if E2.Visible and ( E2.Text = '' ) then exit ;
   if HE.Visible then
    FZReclass.DtPeriodeDest := StrToDate(HE.Value)
     else
      FZReclass.DtPeriodeDest := FDtPer ;
  end // if
   else
    if FTReclass = trJal then
      begin
       FZReclass.StJournalSource := UpperCase(E1.Text) ;
       FZReclass.StJournalDest   := UpperCase(E2.Text) ;
      end // if
       else
        if FTReclass = trEta then
          begin
           if FBoToutEta then
            begin
             if HE.Visible and ( HE.Text = '' ) then exit ;
             if HV.Visible and ( HV.Text = '' ) then exit ;
            end
             else
              begin
               if HE.Visible and ( HE.Text = '' ) then exit ;
               FZReclass.StEtaDest   := HE.Value ;
              end ;
          end // if
           else
            if FTReclass = trFolio then
             begin
              if E2.Visible and ( E2.Text = '' ) then exit ;
              if E2.Visible and ( Valeur(E2.Text) = 0 ) then exit ;
              FZReclass.InFolioDest := StrToInt(E2.Text) ;
             end
              else
               if ( FTReclass = trGen ) or ( FTReclass = trAux ) then
                begin
                 FZReclass.StGenDest   := E2.Text ;
                 FZReclass.StGenSource := E1.Text ;
                 FZReclass.StAuxDest   := E4.Text ;
                end
                 else
                  begin
                   PGIInfo('Fonction non disponible') ;
                   exit ;
                  end;

 if ( TW.Selected <> nil ) and not FBoToutEta then
  begin
   lP := PInfoTW(TW.Selected.Data) ;
   if lP = nil then exit ;
  end ;// if

 E2.OnExit  := nil ;
 E1.OnExit  := nil ;

 if FBoToutEta then
  ExecuteToutEta
   else
    if FBoLW then
     FParam.Visible := not ExecuteLW
      else
       FParam.Visible := not ExecuteTW ;

// FParam.Visible := false ;

 finally
  E2.OnExit       := E2Exit ;
  E1.OnExit       := E1Exit ;
  TW.Enabled     := true ;
  LW.Enabled     := true ;
  BChoix.Enabled := true ;
  BAff.Enabled   := true ;
  BZoom.Enabled  := true ;
  BTree.Enabled  := true ;
 end ;

end;

procedure TFReclassement.BFermeClick(Sender: TObject);
begin
 FParam.Visible := false ;
 TW.Enabled     := true ;
 LW.Enabled     := true ;
 FBoToutEta     := false ;
 BChoix.Enabled := true ;
 BAff.Enabled   := true ;
 BZoom.Enabled  := true ;
 BTree.Enabled  := true ;
end;

procedure TFReclassement.TWDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean) ;
var
 lNode : TTreeNode ;
 lItem : TListItem ;
begin
 if Source = Sender then exit ;
 Accept := false ;
 lNode := TW.GetNodeAt(X,Y) ;
 if ( lNode = nil ) or ( lNode.Data = nil ) then exit ;
 lItem := (Source as TListView).Selected ;
 if ( lItem = nil ) or ( lItem.Data = nil ) then exit ;
 if PInfoTW(lItem.Data)^.InNiv = 3 then
  Accept := ( Source is TListView ) and ( Sender is TTreeview ) and
           ( PInfoTW(lNode.Data)^.TExo.Code = PInfoTW(lItem.Data)^.TExo.Code ) and
           ( not PInfoTW(lNode.Data)^.BoClos ) and ( not PInfoTW(lNode.Data)^.BoValide ) and
           ( not PInfoTW(lItem.Data)^.BoClos ) and ( not PInfoTW(lItem.Data)^.BoValide )
  else
   Accept := false ;
end;


procedure TFReclassement.TWDragDrop(Sender, Source: TObject; X,Y: Integer);
var
 lNode           : TTreeNode ;
 lItem           : TListItem ;
 lPSource        : PInfoTW ;
 lPDestination   : PInfoTW ;
begin
 if Source = Sender then exit ;
 lNode := TW.GetNodeAt(X,Y) ;
 if ( lNode = nil ) or ( lNode.Data = nil ) then exit ;
 lItem := (Source as TListView).Selected ;
 if ( lItem = nil ) or ( lItem.Data = nil ) then exit ;
 lPSource       := PInfoTW(lItem.Data) ;
 lPDestination  := PInfoTW(lNode.Data) ;
 if lPDestination^.InNiv = 0 then exit ;
 if lPSource^.InNiv = 3 then
  begin
   if lPDestination^.InNiv = 2 then
    begin
     if lPSource^.InId <> lPDestination^.InID then
      FTReclass := trPeriode
       else
        FTReclass := trJal ;
    end
    else
     FTReclass := trPeriode ;
  end;

 InitParam(lPSource,lPDestination) ;

end;

procedure TFReclassement.MMPerClick(Sender: TObject);
begin
 FTReclass := trPeriode ;
 InitParam(PInfoTW(TW.Selected.Data)) ;
end;

procedure TFReclassement.MMEtaClick(Sender: TObject);
begin
 FTReclass := trEta ;
 if LW.Focused then
  InitParam(PInfoTW(LW.Selected.Data))
   else
    if TW.Focused then
     begin
      FBoToutEta := true ;
      InitParam(PInfoTW(TW.Selected.Data)) ;
     end; // if
end;

procedure TFReclassement.MMGenClick(Sender: TObject);
begin
 FTReclass := trGen ;
 if LW.Focused then
  InitParam(PInfoTW(LW.Selected.Data))
   else
    if TW.Focused then
     InitParam(PInfoTW(TW.Selected.Data)) ;
end;

procedure TFReclassement.MMAuxClick(Sender: TObject);
begin
 FTReclass := trAux ;
 if LW.Focused then
  InitParam(PInfoTW(LW.Selected.Data))
   else
    if TW.Focused then
     InitParam(PInfoTW(TW.Selected.Data)) ;
end;

procedure TFReclassement.POPCPopup(Sender: TObject);
begin
 EnabledControl ;
 FBoLW := LW.Focused ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 12/10/2004
Modifié le ... : 05/06/2007
Description .. : LG - 12/10/2004 - FB 14764 - debracnchement du
Suite ........ : changement de general sur l'exercice
Suite ........ : - LG - 02/05/2005 - FB 15323 - on ne peut plus selectinner 
Suite ........ : journal qd on est ds le treeview
Suite ........ : - LG - 07/09/2005 - FB 16262 - plantage qd on changait de 
Suite ........ : mode d'affichage
Suite ........ : - LG - 05/06/2007 - FB 19109 - suppression de chg de 
Suite ........ : compte en tri par journal
Mots clefs ... : 
*****************************************************************}
procedure TFReclassement.EnabledControl ;
var
 lP              : PInfoTW ;
 lBoRienSelectLW : boolean ;
 lBoInfo         : boolean ;
begin


 lBoRienSelectLW := LW.SelCount = 0 ;
 lBoInfo         := not lBoRienSelectLW and ( PInfoTW(LW.Selected.Data) <> nil ) ;

 if lBoInfo then
  lP := PInfoTW(LW.Selected.Data)
   else
    lP := nil ;

 MMJal.Enabled   := ( not TW.Focused ) and lBoInfo and not lP^.BoClos and not lP^.BoValide and ( lP^.InNiv = 3 );
 MMFolio.Enabled := lBoInfo and  ( LW.SelCount = 1 ) and not lP^.BoClos and not lP^.BoValide and ( lP^.E_MODESAISIE <> '-' ) and ( lP^.InNiv = 3 );

 if TW.Focused then
  begin
   MMPer.Enabled   := false ;
   MMAux.Enabled   := false ;
   if TW.Selected <> nil then
    begin
     MMEta.Enabled   := ( TW.Selected.ImageIndex =  cImageBureau ) ;//lBoInfo and ( TW.Selected.ImageIndex =  cImageBureau ) ;
     MMGen.Enabled   := ( PInfoTW(TW.Selected.Data) <> nil ) and not PInfoTW(TW.Selected.Data)^.BoClos and not PInfoTW(TW.Selected.Data)^.BoValide and ( PInfoTW(TW.Selected.Data)^.InNiv = 1 ) and ( FTri = ttExo ) ;
   end ;
  end
   else
    begin
     MMEta.Enabled   := lBoInfo and not lP^.BoClos and not lP^.BoValide and ( lP^.InNiv = 3 ) ;
     MMPer.Enabled   := lBoInfo and not lP^.BoClos and not lP^.BoValide and ( lP^.InNiv = 3 ) ;
     MMGen.Enabled   := lBoInfo and not lP^.BoClos and not lP^.BoValide and ( lP^.InNiv = 3 ) ;
     MMAux.Enabled   := MMGen.Enabled ;
    end ;

 MMOuv.Enabled   := lBoInfo and  ( LW.SelCount = 1 ) and ( lP^.InNiv = 3 ) ;

 if TW.Selected = nil then exit ;
 Self.caption := _FabriqTitre(PInfoTW(TW.Selected.Data),FTri) ;
 UpdateCaption(Self) ;

{$IFDEF TT1}
 MMAna.Enabled := true ;
 MMAna.Visible := true ;
 MMContr.Enabled := true ;
 MMContr.Visible := true ;
 MMRecup.Enabled := true ;
 MMRecup.Visible := true ;
 MMsauv.Visible  := true ;
 MMsauv.Enabled  := true ;
 MMInsert.Visible  := true ;
 MMInsert.Enabled  := true ;
 MMOuv.Enabled   := false ;
 MMGen.Enabled   := false ;
 MMPer.Enabled   := false ;
 MMEta.Enabled   := false ;
 MMFolio.Enabled := false ;
 MMJal.Enabled := false ;
{$ENDIF}

end;

procedure TFReclassement.MMFolioClick(Sender: TObject);
begin
 FTReclass := trFolio ;
 InitParam(PInfoTW(LW.Selected.Data)) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 21/06/2005
Modifié le ... : 22/12/2006
Description .. : - LG - 21/06/2005 - FB 14770 - F5 ouvre la fenetre de 
Suite ........ : saisie ( bordereau etc etc )
Suite ........ : - LG - 22/12/2006 - LG - le escap faisaint plante en sortie
Mots clefs ... : 
*****************************************************************}
procedure TFReclassement.FormKeyDown(Sender: TObject; var Key: Word ; Shift: TShiftState);
var
 Vide : Boolean ;
begin
 Vide:=(Shift=[]) ;
 case Key of
  VK_F10          : if Vide and FParam.Visible then begin Key:=0 ; BValiderClick(nil) ; end ;
  VK_F3           : if Vide and ( FIndexRech > 0 ) then begin Key:=0 ; BSuivantRechClick(nil) ; end ;
  VK_F11          : begin Key:=0 ; PopC.PopUp(Mouse.CursorPos.X,Mouse.CursorPos.Y) ; end ;
  VK_RETURN       : if LW.Focused and Vide and ( not FParam.Visible) then begin Key := 0 ; LWDblClick(nil) ; end ;
  VK_F5           : if not FPAram.Visible then begin key := 0 ; LWDblClick(nil) ; end ;
  VK_ESCAPE       : if FParam.Visible then
                     begin
                      Key := 0 ;
                      BFermeClick(nil) ;
                     end
                      else
                       if Shift=[] then
                         begin
                          key := 0 ;
                          Close ;
                          if IsInside(Self) then
                           begin
                            CloseInsidePanel(Self) ;
                            THPanel(Self.parent).InsideForm := nil;
                            THPanel(Self.parent).VideToolBar;
                           end ; 
                         end ; 
{CTRL+A}     65   : if Shift=[ssCtrl] then begin Key:=0 ; ToutSelectionnerLW ; end ;
{CTRL+F}     70   : if Shift=[ssCtrl] then begin Key:=0 ; Chercher ; end ;
{ALT+J}      74   : if Shift=[ssAlt]  then begin Key:=0 ; ZoomJal ; end ;
 end; // case
end;

procedure TFReclassement.ToutSelectionnerLW ;
var
 i : integer ;
begin
 for i := 0 to LW.Items.Count - 1 do
 // if PInfoTW(LW.Items[i]).InNiv = 3 then
   LW.Items[i].Selected := not LW.Items[i].Selected ;
end ;

procedure TFReclassement.Chercher ;
begin
 if FParam.Visible then exit ;
 FRech.Visible := true ;
 FIndexRech    := 0 ;
 E3.Text       := '' ;
end ;

procedure TFReclassement.BSuivantRechClick(Sender: TObject);
var
 i      : integer ;
 lP     : PInfoTW ;
begin
 for  i:= FIndexRech to TW.Items.Count - 1 do
  begin
   lP := PInfoTW(TW.Items[i].Data) ; if lP = nil then continue ;
   if lP^.E_JOURNAL = UpperCase(E3.Text) then
    begin
     TW.Items[i].selected := true ;
     FIndexRech           := i + 1 ;
     exit ;
    end;
  end ;
end;

procedure TFReclassement.BAnnulerRechClick(Sender: TObject);
begin
 FRech.Visible := false ;
 FIndexRech    := 0 ;
 E3.Text       := '' ;
end;

procedure TFReclassement.TWChange(Sender: TObject; Node: TTreeNode);
begin
 EnabledControl ;
 TWClick(nil);
end;

procedure TFReclassement.MMOuvClick(Sender: TObject);
begin
 LWDblClick(nil) ;
end;

procedure TFReclassement.MMGrandClick(Sender: TObject);
begin
 LW.ViewStyle    := Comctrls.vsIcon ;
 MMGrand.Checked := true ;
 MMPetit.Checked := false ;
end;

procedure TFReclassement.MMPetitClick(Sender: TObject);
begin
 LW.ViewStyle    := Comctrls.vsList ;
 MMGrand.Checked := false ;
 MMPetit.Checked := true ;
end;

procedure TFReclassement.MMCodeExoClick(Sender: TObject) ;
var
 i  : integer ;
 lP : PInfoTW ;
begin
 MMCodeExo.Checked := not MMCodeExo.Checked ;
 for i := 0 to TW.Items.Count - 1 do
  begin
   lP := PInfoTW(TW.Items[i].Data) ;
   if lP = nil then continue ;
   if lP^.TypeNode = tnExo then
    TW.Items[i].Text := GetLibExo(lP) ;
  end; // if
end;

procedure TFReclassement.MMAFFJALClick(Sender: TObject);
var
 i  : integer ;
 lP : PInfoTW ;
begin
 MMAffJal.Checked := not MMAffJal.Checked ;
 for i := 0 to TW.Items.Count - 1 do
  begin
   lP := PInfoTW(TW.Items[i].Data) ;
   if lP = nil then continue ;
   if lP^.TypeNode = tnJal then
    TW.Items[i].Text := GetLibJal(lP);
  end; // for
end;

function TFReclassement.GetLibJal( vP : PInfoTW ) : string ;
begin
 if not MMAFFJAL.Checked then   // CA - 24/10/2003 - fonction inversée
  result := vP^.E_JOURNAL
   else
    result := vP^.J_LIBELLE ;
end;

function TFReclassement.GetLibExo( vP : PInfoTW ) : string ;
begin
 if not MMCodeExo.Checked then  // CA - 24/10/2003 - fonction inversée
  result := cStTitreTWExo + vP^.TExo.Code
   else
    result := vP^.EX_ABREGE ;
end;

procedure TFReclassement.MMVerifClick(Sender: TObject);
var
 i  : integer ;
 lP : PInfoTW ;
begin
 for i := 0 to LW.Items.Count -1 do
  begin
   lP := PInfoTW(LW.Items[i].Data) ;
   if lP = nil then continue ;
   if not FZReclass.IsValid(lP) then
    begin
     lP^.StInfoLibre        := 'Erreur' ;
     LW.items[i].ImageIndex := cImageFolioClot ;
    end; // if
  end; // if
end;

procedure TFReclassement.ZoomJal ;
var
 lP      : PInfoTW ;
 lAction : TActionFiche ;
begin

 if TW.Selected = nil then exit ;
 lP := PInfoTW(TW.Selected.Data) ;
 if lP = nil then exit ;

 if lP^.E_JOURNAL = '' then exit ;
 lAction := taModif ;
 if not ExJaiLeDroitConcept(TConcept(ccJalModif), FALSE) then lAction := taConsult ;

 FicheJournal(nil, '', lP^.E_JOURNAL,lAction,0) ;

end ;

procedure TFReclassement.MMZoomJalClick(Sender: TObject);
begin
 ZoomJal ;
end;

procedure TFReclassement.MMAnaClick(Sender: TObject);
var
 lP : PInfoTW ;
begin

 if TW.Selected = nil then exit ;
 lP := PInfoTW(TW.Selected.Data) ;
 if lP = nil then exit ;

{$IFDEF TT}
 FZReclass._ExecuteAna(lP) ;
{$ENDIF}

end;





{ TZreclassement }

constructor TZreclassement.Create(vInfoEcr: TInfoEcriture);
begin
  inherited;
  FListPiece        := TList.Create ;
  FTOBSource        := TOB.Create('',nil,-1) ;
  FTOBDestination   := TOB.Create('',nil,-1) ;
  FBoDuplication    := false ;
  FCR               := TList.Create ;
end;

destructor TZreclassement.Destroy ;
var
 i : integer ;
begin

 for i := 0 to FListPiece.Count - 1 do
  if assigned(FListPiece[i]) then
   Dispose(PInfoTW(FListPiece[i])) ;

  RAZFLC ;

  FListPiece.Free ;
  FTObSource.Free ;
  FTOBDestination.Free ;
  FCR.Free ;

  // GCO - 31/08/2006 - FQ 18578
  V_PGI.IoError := OeOk;

  inherited;
  
end;

procedure TZreclassement.Add ( vP : PInfoTW );
begin
 if vP <> nil then
  FListPiece.Add(vP) ;
end;

procedure TZreclassement.RAZFLC ;
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
Auteur  ...... : Laurent Gendreau
Créé le ...... : 09/05/2005
Modifié le ... :   /  /    
Description .. : - LG - 09/05/2005 - FB 14641 - on n'affiche les montants 
Suite ........ : s'ils sont nuls
Mots clefs ... : 
*****************************************************************}
procedure TZReclassement.StockeLesEcrituresIntegrees( vTOB : TOB ; vBoDetail : boolean = true);
var
 lLigneCR : TCRPiece ; // defini ds ImRapInt
 lTOB     : TOB ;
begin
 lTOB := nil ;
 if not vBoDetail then lTOB := vTOB
  else
   if vTOB.Detail.Count > 0 then
    lTOB := vTOB.Detail[0] ;
 if lTOB <> nil then
  begin
   lLigneCR             := TCRPiece.Create ;
   lLigneCR.NumPiece    := lTOB.GetValue('E_NUMEROPIECE') ;
   lLigneCR.Journal     := lTOB.GetValue('E_JOURNAL') ;
   lLigneCR.QualifPiece := lTOB.GetValue('E_QUALIFPIECE') ;
   lLigneCR.Date        := lTOB.GetValue('E_DATECOMPTABLE') ;
   if vTOB.Detail.Count > 0 then
    begin
     lLigneCR.Debit      := vTOB.Somme('E_DEBIT', ['E_NUMEROPIECE'],[lLigneCR.NumPiece],TRUE);
     lLigneCR.Credit     := vTOB.Somme('E_CREDIT',['E_NUMEROPIECE'],[lLigneCR.NumPiece],TRUE);
    end  // if
     else
      begin
       lLigneCR.Debit      := -1 ;
       lLigneCR.Credit     := -1 ;
     end ; // if
   FCR.Add (lLigneCR);
  end; // for
end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 30/09/2005
Modifié le ... : 27/07/2007
Description .. : - LG - 30/09/2005 - FB 16765 - suppresion du controle sur
Suite ........ : le pointage, fait ds une autre fct
Suite ........ : - LG - 04/11/2005 - FB 16993 - suppresion du 16765
Suite ........ : - LG - 21/02/2007 - on bloque les reclassment pour S1 qd 
Suite ........ : on appelle la fct depuis la consultation des comptes
Suite ........ : - LG - 27/07/2007 - Fb 21178 - FB 21081 - on remonte le try 
Suite ........ : pour detruite a chq fois les objets du finally
Mots clefs ... : 
*****************************************************************}
function TZReclassement.Execute : boolean ;
var
 i      : integer ;
 lTri   : TTriTw ;
begin

 result     := false ;

 if LienS1 then
 begin
  PGIInfo('Fonction non disponible !' +#10#13 + 'Vous êtes en synchro Business Line !','Reclassement');
  exit ;
 end;

 if FListPiece.Count = 0 then
  begin
   NotifyError(-1,cErrPasEnr,'TZReclassement.Execute') ;
   exit ;
  end; // if

  try

  case FTypeTrait of
     trJal             : if not EstCorrectJournal then exit ;
     trGen,trGenLigne  : begin
                          if not EstCorrectCompteSource then exit ;
                          if not EstCorrectCompteDest then exit ;
                         end ;
     trAux, trAuxLigne : begin
                          if not EstCorrectCompteSource then exit ;
                          if not EstCorrectCompteDest then exit ;
                          if not EstCorrectAux then exit ;
                         end ;
     trFolio           : if InFolioDest < 1 then exit ;
    end; // case

{$IFDEF AMORTISSEMENT}
  FBoMajImmo := False ;
  if ( FStGenSource <> '' ) and ( FStGenDest <> '' ) then
   begin
    if (GetColonneSQL('GENERAUX', 'G_NATUREGENE', 'G_GENERAL = "' + FStGenSource + '"') = 'IMO') and
       (GetColonneSQL('GENERAUX', 'G_NATUREGENE', 'G_GENERAL = "' + FStGenDest + '"') = 'IMO') then
    begin
      if EstCorrectCompteImmo then
       FBoMajImmo := True
        else
         Exit;
    end;
   end ;
{$ENDIF}

  InitMoveProgressForm (nil, 'Reclassement', 'Traitement en cours', FListPiece.Count , true , true ) ;

  //try

  for i := 0 to FListPiece.Count - 1 do
   begin

    FPSource := PInfoTW( FListPiece.Items[i] ) ;
    if not IsValid(FPSource) then
     begin
      NotifyError(-1,cErrMvtValide,'TZReclassement.Execute') ;
      V_PGI.IOError := oeSaisie ;
      continue ;
     end ;

    if not LoadParFolio then
     begin
      //Result := false ;
      NotifyError(-1,cErrAucunErr,'TZReclassement.Execute') ;
      continue ;
     end ;

    if FPSource^.BoTriExo then
     lTri := ttExo
      else
       lTri := ttJal ;

    if not MoveCurProgressForm(_FabriqTitre(FPSource,lTri)) then break ;

    case FTypeTrait of
     trGen,trGenLigne   : if not ValideGen then exit ;
    end ; // case

    case FTypeTrait of
     trPeriode            : result := Transactions(ExecutePer   ,0 ) = oeOk ;
     trJal                : result := Transactions(ExecuteJal   ,0 ) = oeOk ;
     trEta                : result := Transactions(ExecuteEta   ,0 ) = oeOk ;
     trFolio              : result := Transactions(ExecuteFolio ,0 ) = oeOk ;
     trGen,trGenLigne     : result := Transactions(ExecuteGen   ,0 ) = oeOk ;
     trAux,trAuxLigne     : result := Transactions(ExecuteAux   ,0 ) = oeOk ;
    end; // case

   if not result then exit ;

   end; // for

 finally
  if not result and ( V_PGI.LastSQLError <> '' ) then
    PGIInfo('Erreur lors du transfert : ' + V_PGI.LastSQLError , 'Reclassement' ) ;
  FiniMoveProgressForm ;
  if result and (FCR.Count <> 0 ) then AfficheCRIntegration (FCR); // fct de libI\ImRapInt.pas
  RAZFLC ;
  VideLaListe;
 end ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 22/09/2004
Modifié le ... : 25/07/2007
Description .. : - LG - 22/09/2004 - FB 14628 - gestion des exo decale 
Suite ........ : pour la date de destiantion
Suite ........ : - LG - 22/06/2006 - FB 18280 - on additionnait au lieu de 
Suite ........ : soustraire pour la mise ajour des comptes 
Suite ........ : - LG - 22/02/2007 - E_IO=X vu avec manou
Suite ........ :  - on interdit de changé la  cle en liaison avec S1
Suite ........ : - LG - FB 19104 - msg sur les immo en chg de periode
Mots clefs ... : 
*****************************************************************}
procedure TZReclassement.ExecutePer ;
var
 i                : integer ;
 lTOB             : TOB ;
 lYearSource      : word ;
 lMonthSource     : word ;
 lDaySource       : word ;
 lYearDest        : word ;
 lMonthDest       : word ;
 lDayDest         : word ;
 lMaxMonth        : integer ;
 j                : integer ;
 k                : integer ;
 lTOBAna          : TOB ;
 lBoFirst         : boolean ;
 lBoRenum         : boolean ;
 lQ               : TQuery ;
 lRecError        : TRecError ;
 lDtLaDate        : TDatetime ;
 lBoImmo          : boolean ;
begin

 lBoFirst := true ;
 lBoRenum := false ;

 if _AvecEchange then
  begin
   PGiBox ('Le dossier est en liaison avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) + ', le reclassement est interdit', 'Reclassement');
   exit ;
  end ;


 // controle de la validite de la periode de destination
  if Info.Journal.Load([FTOBDestination.Detail[0].GetValue('E_JOURNAL')]) = - 1 then
  begin
   NotifyError( RC_JALINEXISTANT , '' ) ;
   exit ;
  end ;

  if _IsValidJournalPeriode( FTOBDestination.Detail[0].GetValue('E_EXERCICE') ,
                    Info.Journal.GetValue('J_VALIDEEN')  ,
                    Info.Journal.GetValue('J_VALIDEEN1') ,
                    FTOBDestination.Detail[0].GetValue('E_DATECOMPTABLE') ,Info )  <> RC_PASERREUR then
  begin
   NotifyError( cErrTransfert , cErrJalVal ) ;
   exit ;
  end;

 lBoImmo := false ;

 try

 for i := 0 to FTOBDestination.Detail.Count - 1 do
  begin
   lTOB   := FTOBDestination.Detail[i] ;

   DecodeDate(lTOB.GetValue('E_DATECOMPTABLE'),lYearSource,lMonthSource,lDaySource) ;
   DecodeDate(FDtPeriodeDest,lYearDest,lMonthDest,lDayDest) ;

   lBoImmo := lBoImmo or ( lTOB.GetString('E_IMMO') <> '' ) ;

    if lBoFirst then
    begin
     lBoFirst := false ;
     lQ := OpenSql('SELECT COUNT(*) N FROM ECRITURE WHERE ' +
                   'E_EXERCICE="' + lTOB.GetValue('E_EXERCICE')     + '" ' +
                   'AND E_JOURNAL="' + lTOB.GetValue('E_JOURNAL')   + '" ' +
                   'AND E_NUMEROPIECE='  + IntToStr(lTOB.GetValue('E_NUMEROPIECE'))  + ' '  +
                   'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(lYearDest , lMonthDest, 1)) + '" ' +
                   'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(lYearDest , lMonthDest, 1))) + '" ' +
                   'AND E_QUALIFPIECE="'+ lTOB.GetValue('E_QUALIFPIECE')     + '" ' , true );
     lBoRenum := lQ.FindField('N').asInteger <> 0 ;
    end; // if

   lMaxMonth    := DaysPerMonth(lYearDest,lMonthDest) ;
   lMonthSource := lYearDest ;
   lYearSource  := lYearDest ;
   if lDaySource > lMaxMonth then
    lDaySource := lMaxMonth ;

   lDtLaDate   := EncodeDate(lYearDest, lMonthDest, lDaySource) ;

   if lDtLaDate < FPSource^.TExo.Deb then
    lDtLaDate := FPSource^.TExo.Deb
     else
      if lDtLaDate > FPSource^.TExo.Fin then
       lDtLaDate := FPSource^.TExo.Fin ;

   lTOB.PutValue('E_DATECOMPTABLE',lDtLaDate) ;
   if lTOB.GetValue('E_DATECOMPTABLE') > lTOB.GetValue('E_DATEECHEANCE') then
    begin
     lTOB.PutValue('E_DATEECHEANCE',lTOB.GetValue('E_DATECOMPTABLE')) ;
     lTOB.PutValue('E_ORIGINEPAIEMENT', lTOB.GetValue('E_DATECOMPTABLE') ) ;
    end;
   CSetPeriode(lTOB);
   lTOB.PutValue('E_IO'      , 'X') ;
   lTOB.PutValue('E_CREERPAR', 'REC') ;
   for j := 0 to lTOB.Detail.Count - 1 do
    for k := 0 to lTOB.Detail[j].Detail.Count - 1 do
     begin
      lTOBAna := lTOB.Detail[j].Detail[k] ;
      lTOBAna.PutValue('Y_DATECOMPTABLE',lTOB.GetValue('E_DATECOMPTABLE')) ;
      CSetPeriode(lTOBAna) ;
     end; // for

   lRecError := CIsValidLigneSaisie(lTOB,Info,true) ;
   if lRecError.RC_Error <> RC_PASERREUR then
    begin
     NotifyError(lRecError.RC_Error,'') ;
     V_PGI.IOError := oeSaisie ;
     exit ;
    end ;

  end; // for

 if not CEnregistreSaisie(FTOBDestination,lBoRenum,false,true,Info) then
  begin
   V_PGI.IOError := oeSaisie ;
   exit ;
  end;

 StockeLesEcrituresIntegrees(FTOBDestination) ;

 if not FBoDuplication then
  begin
   FTOBSource.DeleteDB(false) ;
   MajSoldesEcritureTOB (FTOBSource,false) ;
  end; // if

  V_PGI.IOError := oeOk ;

  if lBoImmo then
   PgiInfo('Suite au reclassement, veuillez modifier la date d''achat de l''immobilisation.', 'Reclassement des données');

 except
  on E : Exception do
   begin
    NotifyError( cErrPrg ,e.message) ;
    raise ;
   end;
 end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/08/2004
Modifié le ... : 22/02/2007
Description .. : LG - 19/08/2004 - FRB 13274 - Le changement 
Suite ........ : d'établissement n'impacte pas les écritures analytiques
Suite ........ : lorsqu'il est fait pour l'ensemble du dossier
Suite ........ : - LG - 21/12/2004 - FB 13274 - coorection de la rq de 
Suite ........ : l'analytiq
Suite ........ : - LG - 22/02/2007 - E_IO=X vu avec manou
Mots clefs ... : 
*****************************************************************}
procedure TZReclassement.ExecuteToutEta ;
begin
 if StEtaSource ='' then
  begin
   ExecuteSQL('UPDATE ECRITURE SET E_ETABLISSEMENT="'+StEtaDest+'" , E_IO="X" ' ) ;
   ExecuteSQL('UPDATE ANALYTIQ SET Y_ETABLISSEMENT="'+StEtaDest+'" ' ) ;
   ExecuteSQL('UPDATE IMMO SET I_ETABLISSEMENT="'+StEtaDest+'" ' ) ;
  end
   else
    begin
     ExecuteSQL('UPDATE ECRITURE SET E_ETABLISSEMENT="'+StEtaDest+'" , E_IO="X" WHERE E_ETABLISSEMENT="'+StEtaSource+'" ') ;
     ExecuteSQL('UPDATE ANALYTIQ SET Y_ETABLISSEMENT="'+StEtaDest+'" WHERE Y_ETABLISSEMENT="'+StEtaSource+'" ') ;
     ExecuteSQL('UPDATE IMMO SET I_ETABLISSEMENT="'+StEtaDest+'" WHERE I_ETABLISSEMENT="'+StEtaSource+'" ') ;
    end ;
end ;

function TZReclassement.ValideGen : boolean ;
var
 i               : integer ;
 lTOB            : TOB ;
 lBoLettrable    : boolean ;
 lRecError       : TRecError ;
begin

 result := false ;

 if Info.Compte.GetCompte(FStGenDest) < 0 then exit ;
 lBoLettrable := Info.Compte.GetValue('G_LETTRABLE') = 'X' ;

 for i :=  0 to FTOBDestination.Detail.Count - 1 do
  begin

   lTOB := FTOBDestination.Detail[i] ;

   lTOB.PutValue('E_GENERAL'    , FStGenDest) ;
   if lBoLettrable then
    CRemplirInfoLettrage(lTOB) ;

  // lInCodeErreur := CIsValidcompte(lTOB,Info) ;
   lRecError := CIsValidLigneSaisie(lTOB,Info,true) ;
   if ( lRecError.RC_Error <> RC_PASERREUR ) then
     begin
      NotifyError( lRecError.RC_Error , '' ) ;
      exit ;
     end ;

   if not IsValidGen(lTOB) then
    exit ;

  end ; // for

 result := true ;
  
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/02/2007
Modifié le ... : 06/07/2007
Description .. : - 20/02/2007 - FB 19707 - on utilise une fct du noyau pour
Suite ........ : valider la piece
Suite ........ : - LG - 22/02/2007 - E_IO=X vu avec manou
Suite ........ : - LG - 13/06/2007 - FB 20610 - le TOP 1 ne fct en eAGL
Suite ........ : - LG - 06/07/2007 - FB 19314 - ajout du transfert non 
Suite ........ : lettrable vers lettrable
Mots clefs ... : 
*****************************************************************}
procedure TZReclassement.ExecuteGen ;
var
 i               : integer ;
 lStSQL          : string ;
 lStContrePartie : string ;
 lQ              : TQuery ;
 lBoLettrable    : boolean ;
begin

 if Info.Compte.GetCompte(FStGenDest) < 0 then exit ;
  lBoLettrable := Info.Compte.GetValue('G_LETTRABLE') = 'X' ;


 try

  for i :=  0 to FTOBDestination.Detail.Count - 1 do
   begin

    lQ     := OpenSql('SELECT ##TOP 1## E_GENERAL FROM ECRITURE ' +
                      ' WHERE E_NUMEROPIECE=' + intToStr(FTOBDestination.Detail[i].getValue('E_NUMEROPIECE')) +
                      ' AND E_JOURNAL="'               + FTOBDestination.Detail[i].getString('E_JOURNAL')               + '"' +
                      ' AND E_EXERCICE="'              + FTOBDestination.Detail[i].getString('E_EXERCICE')              + '"' +
                      ' AND E_DATECOMPTABLE="'         + USDateTime(FTOBDestination.Detail[i].getValue('E_DATECOMPTABLE') ) + '"' +
                      ' AND E_QUALIFPIECE="'           + FTOBDestination.Detail[i].GetString('E_QUALIFPIECE') + '"' +
                      ' AND E_CONTREPARTIEGEN="'       + FStGenSource + '" ' , true ) ;


    if not lQ.EOF then
     lStContrePartie := lQ.FindField('E_GENERAL').asString ;
    Ferme(lQ) ;

    // on ne pas a jour les auxi, on ne peu pas modifie un compte auxi
    lStSQL := 'UPDATE ECRITURE SET E_GENERAL="' + FStGenDest + '", E_IO="X",E_CREERPAR="REC",E_CONTREPARTIEGEN="' + lStContrePartie + '" ' ;

    if lBoLettrable then
     begin
      lStSQL :=  lStSQL + ',E_ECHE="X", E_ETATLETTRAGE="AL",E_NUMECHE=1' ;
      {$IFNDEF TRESO}
      lStSQL :=  lStSQL + ',E_TRESOSYNCHRO="CRE"' ;
      {$ENDIF}
     end ; // if


    lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FTOBDestination.Detail[i].getValue('E_NUMEROPIECE')) +
               ' AND E_JOURNAL="'              + FTOBDestination.Detail[i].getValue('E_JOURNAL')               + '"' +
               ' AND E_EXERCICE="'             + FTOBDestination.Detail[i].getValue('E_EXERCICE')              + '"' +
               ' AND E_DATECOMPTABLE="'        + USDateTime(FTOBDestination.Detail[i].getValue('E_DATECOMPTABLE') ) + '"' +
               ' AND E_QUALIFPIECE="'          + FTOBDestination.Detail[i].getValue('E_QUALIFPIECE') + '"' +
               ' AND E_NUMLIGNE='              + intToStr(FTOBDestination.Detail[i].getValue('E_NUMLIGNE')) ;

    if ExecuteSql(lStSQL) <> 1 then
     begin
      V_PGI.IOError := oeSaisie ;
      NotifyError(cErrPrg ,v_pgi.LastSQLError,'TZReclassement.ExecuteGen') ;
      exit ;
     end ;

    // on ne pas a jour les auxi, on ne peu pas modifie un compte auxi
     if lStContrePartie <> '' then
      begin
       lStSQL := 'UPDATE ECRITURE SET E_CONTREPARTIEGEN="' + FStGenDest + '", E_IO="X",E_CREERPAR="REC" ' ;

       lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FTOBDestination.Detail[i].getValue('E_NUMEROPIECE')) +
               ' AND E_JOURNAL="'               + FTOBDestination.Detail[i].getValue('E_JOURNAL')               + '"' +
               ' AND E_EXERCICE="'              + FTOBDestination.Detail[i].getValue('E_EXERCICE')              + '"' +
               ' AND E_DATECOMPTABLE="'         + USDateTime(FTOBDestination.Detail[i].getValue('E_DATECOMPTABLE') ) + '"' +
               ' AND E_QUALIFPIECE="'           + FTOBDestination.Detail[i].getValue('E_QUALIFPIECE') + '"' +
               ' AND E_CONTREPARTIEGEN="'       + FStGenSource + '" '  ;
     end ;
     
    ExecuteSql(lStSQL) ;
    
  end ;

  // Mise à jour des IMMO pour les Comptes IMMO
{$IFDEF AMORTISSEMENT}
  // GCO - 06/04/2006 - FQ 15240 - Mise à jour des immobilisations
  if FBoMajImmo then
  begin
    // GCO - 30/11/2006 - 19103   FB 20616
  //  PgiInfo('Suite au reclassement, veuillez modifier la date d''achat de l''immobilisation.', 'Reclassement des données');
    if not TraitementImmo then
    begin
      V_PGI.IOError := oeSaisie;
      NotifyError(cErrPrg , v_pgi.LastSQLError, 'TZReclassement.TraitementImmo');
      Exit;
    end;
  end;
{$ENDIF}

 StockeLesEcrituresIntegrees(FTOBDestination) ;
 MajSoldesEcritureTOB (FTOBSource,false) ;
 MajSoldesEcritureTOB (FTOBDestination,true) ;

 V_PGI.IOError := oeOk ;

 except
  on E : Exception do
   begin
    NotifyError(cErrPrg ,E.Message,'TZReclassement.ExecuteGen') ;
     V_PGI.IOError := oeSaisie ;
    raise ;
   end;
 end;

end;


function TZReclassement.IsValidGen ( lTOB : TOB ) : boolean ;
var
 lBoLettrage     : boolean ;
 lBoPointage     : boolean ;
 lBoValide       : boolean ;
 lBoImmo         : boolean ;
 lStErr          : string ;
begin

 result      := false ;
 lBoPointage := false ;

 if EstSpecif('51209') then
  lBoPointage := ( lTOB.GetString('E_REFPOINTAGE') <> '' ) ;
 lBoLettrage := ( lTOB.GetString('E_LETTRAGE') <> '' ) ;
 lBoValide   := ( lTOB.GetString('E_VALIDE') = 'X' ) ;
 lBoImmo     := ( lTOB.GetString('E_IMMO') <> '' ) ;

 if lBoPointage or lBoLettrage or lBoValide or lBoImmo then
  begin

    lStErr := 'Journal : ' + lTOB.GetString('E_JOURNAL') + ' Pièce : ' + lTOB.GetString('E_NUMEROPIECE') +
              ' Date : ' + lTOB.GetString('E_DATECOMPTABLE') + ' Ligne : ' + lTOB.GetString('E_NUMLIGNE') ;

    if lBoPointage then
     PGIError(cErrLignePointe + #10#13 + lStErr + ' Référence de pointage : ' + lTOB.GetString('E_REFPOINTAGE') ) 
      else
       if lBoLettrage then
        PGIError(cErrLigneLettre + #10#13 + lStErr + ' Lettrage : ' + lTOB.GetString('E_LETTRAGE') )
         else
          if lBoValide then
           PGIError(cErrLigneValide + #10#13 + lStErr) ;

   end
    else
     result := true ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 13/06/2007
Modifié le ... :   /  /
Description .. : - LG - 13/06/2007 - FB 20614 - le UpdateDB sur la mere ne
Suite ........ : fct pas, lancment sur toute les filles
Mots clefs ... :
*****************************************************************}
procedure TZReclassement.ExecuteAux ;
var
 i               : integer ;
// lInCodeErreur   : integer ;
 lTOB            : TOB ;
 lStContrePartie : string ;
 lStSQL          : string ;
 lQ              : TQuery ;
 lRecErreur      : TRecError ;
begin

 try

 for i :=  0 to FTOBDestination.Detail.Count - 1 do
  begin

   lTOB := FTOBDestination.Detail[i] ;

   if not IsValidGen(lTOB) then
    begin
     V_PGI.IOError := oeSaisie ;
     exit ;
    end ;

   lQ := OpenSql('SELECT ##TOP 1## E_GENERAL FROM ECRITURE ' +
                 ' WHERE E_NUMEROPIECE=' + intToStr(lTOB.getValue('E_NUMEROPIECE')) +
                 ' AND E_JOURNAL="'               + lTOB.getValue('E_JOURNAL')               + '"' +
                 ' AND E_EXERCICE="'              + lTOB.getValue('E_EXERCICE')              + '"' +
                 ' AND E_DATECOMPTABLE="'         + USDateTime(lTOB.getValue('E_DATECOMPTABLE') ) + '"' +
                 ' AND E_QUALIFPIECE="'           + lTOB.getValue('E_QUALIFPIECE') + '"' +
                 ' AND E_CONTREPARTIEGEN="'       + FStGenSource + '" '
                 , true ) ;
   if not lQ.EOF then
    lStContrePartie := lQ.FindField('E_GENERAL').asString ;

   Ferme(lQ) ;

   lTOB.PutValue('E_GENERAL'         , FStGenDest) ;
   lTOB.PutValue('E_AUXILIAIRE'      , FStAuxDest) ;
   lTOB.PutValue('E_CONTREPARTIEGEN' , lStContrePartie) ;
   lTOB.PutValue('E_CREERPAR'        , 'REC' ) ;
   CGetEch(lTOB,Info) ;
   CGetTVA(lTOB,Info) ;
   CGetConso(lTOB,Info) ;

   lRecErreur := CIsValidLigneSaisie(lTOB,Info,true) ;
   if not ( lRecErreur.RC_Error = RC_PASERREUR ) then
     begin
      NotifyError( lRecErreur.RC_Error , '' ) ;
      V_PGI.IOError := oeSaisie ;
      exit ;
     end ;

   lStSQL := 'UPDATE ECRITURE SET E_CONTREPARTIEGEN="' + FStGenDest + '", E_CONTREPARTIEAUX="' + FStAuxDest + '" ,E_IO="X",E_CREERPAR="REC" ' ;

   lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(lTOB.getValue('E_NUMEROPIECE')) +
              ' AND E_JOURNAL="'               + lTOB.getValue('E_JOURNAL')               + '"' +
              ' AND E_EXERCICE="'              + lTOB.getValue('E_EXERCICE')              + '"' +
              ' AND E_DATECOMPTABLE="'         + USDateTime(lTOB.getValue('E_DATECOMPTABLE') ) + '"' +
              ' AND E_QUALIFPIECE="'           + lTOB.getValue('E_QUALIFPIECE') + '"' +
              ' AND E_CONTREPARTIEGEN="'       + FStGenSource + '" '  ;
   if lStContrePartie <> '' then
    ExecuteSql(lStSQL) ;

 end ; // for

// FTOBDestination.SetAllModifie(true) ;
 {$IFDEF EAGLCLIENT}
 for i := 0 to FTOBDestination.Detail.Count - 1 do
  begin
   FTOBDestination.Detail[i].UpdateDB() ;
  end ;
 {$ELSE}
   FTOBDestination.UpdateDB ;
 {$ENDIF}

 StockeLesEcrituresIntegrees(FTOBDestination) ;
 MajSoldesEcritureTOB (FTOBSource,false) ;
 MajSoldesEcritureTOB (FTOBDestination,true) ;

 V_PGI.IOError := oeOk ;

 except
  on E : Exception do
   begin
    NotifyError(cErrPrg ,E.Message,'TZReclassement.ExecuteGen') ;
     V_PGI.IOError := oeSaisie ;
    raise ;
   end;
 end;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/02/2007
Modifié le ... : 22/02/2007
Description .. : - LG - 22/02/2007 - E_IO=X vu avec manou
Suite ........ :  - on interdit de changé la  cle en liaison avec S1
Mots clefs ... : 
*****************************************************************}
procedure TZReclassement.ExecuteFolio ;
var
 lStSQL   : string ;
 lLigneCR : TCRPiece ; // defini ds ImRapInt
begin

 if _AvecEchange then
  begin
   PGiBox ('Le dossier est en liaison avec ' + RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME' , '' ), false ) + ', le reclassement est interdit', 'Reclassement');
   exit ;
  end ;


 if existeSQL     ('SELECT E_NUMEROPIECE FROM ECRITURE WHERE ' +
                   'E_EXERCICE="' + FPSource^.TExo.Code      + '" ' +
                   'AND E_JOURNAL="' + FPSource^.E_JOURNAL    + '" ' +
                   'AND E_NUMEROPIECE='  + IntToStr(FInFolioDest)  + ' '  +
                   'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, 1)) + '" ' +
                   'AND E_DATECOMPTABLE<="' + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '" ' +
                   'AND E_QUALIFPIECE="'+ FPSource^.E_QUALIFPIECE      + '" ' ) then
 begin
   NotifyError( cErrTransfert , cErrBorExist ) ;
   exit ;
 end;

 try

 lStSQL := 'UPDATE ECRITURE SET E_NUMEROPIECE=' + intToStr(FInFolioDest) + ', ' +
           'E_IO="X",E_CREERPAR="REC" ' ;

  if ( FPSource^.E_MODESAISIE = '-' ) or ( FPSource^.E_MODESAISIE = '' ) then
  begin // mode piece
   lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
             ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
             ' AND E_EXERCICE="'              + FPSource^.TExo.Code               + '"' +
             ' AND E_DATECOMPTABLE="'         + USDateTime(FPSource^.E_DATECOMPTABLE) + '"' +
             ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
  end
   else
    begin  // mode bordereau
     lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
               ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
               ' AND E_EXERCICE="'              + FPSource^.TExo.Code               + '"' +
               ' AND E_DATECOMPTABLE>="'        + USDateTime(EncodeDateBor(FPSource^.InAnnee, FPSource^.InMois, FPSource^.TExo)) + '"' +
               ' AND E_DATECOMPTABLE<="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '"' +
               ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
    end; // if

 ExecuteSql(lStSQL) ;

 lStSQL := 'UPDATE ANALYTIQ SET Y_NUMEROPIECE=' + intToStr(FInFolioDest) + ' ' ;

 if ( FPSource^.E_MODESAISIE = '-' ) or ( FPSource^.E_MODESAISIE = '' ) then
  begin // mode piece
   lStSQL := lStSQL +
             'WHERE Y_JOURNAL="'      + FPSource^.E_JOURNAL                    + '" ' +
             'AND Y_EXERCICE="'       + FPSource^.TExo.Code                    + '" ' +
             'AND Y_NUMEROPIECE='     + intToStr(FPSource^.E_NUMEROPIECE)      + ' ' +
             'AND Y_DATECOMPTABLE="'  + USDateTime(FPSource^.E_DATECOMPTABLE)  + '" ' +
             'AND Y_QUALIFPIECE="'    + FPSource^.E_QUALIFPIECE                + '" '   ;
  end
   else
    begin  // mode bordereau
      lStSQL := lStSQL +
                'WHERE Y_JOURNAL="'      + FPSource^.E_JOURNAL                    + '" ' +
                'AND Y_EXERCICE="'       + FPSource^.TExo.Code                    + '" ' +
                'AND Y_NUMEROPIECE='     + intToStr(FPSource^.E_NUMEROPIECE)      + ' ' +
                'AND Y_DATECOMPTABLE>="' + USDateTime(EncodeDateBor(FPSource^.InAnnee, FPSource^.InMois, FPSource^.TExo)) + '" ' +
                'AND Y_DATECOMPTABLE<="' + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '" ' +
                'AND Y_QUALIFPIECE="'    + FPSource^.E_QUALIFPIECE                + '" '   ;
    end; // if

 ExecuteSql(lStSQL) ;

 // on stocke l'ecriture integree
 lLigneCR             := TCRPiece.Create ;
 lLigneCR.NumPiece    := FInFolioDest ;
 lLigneCR.Journal     := FPSource^.E_JOURNAL ;
 lLigneCR.QualifPiece := FPSource^.E_QUALIFPIECE ;
 lLigneCR.Date        := FPSource^.E_DATECOMPTABLE ;
 lLigneCR.Debit       := -1 ;
 lLigneCR.Credit      := -1 ;
 FCR.Add (lLigneCR);

 except
  on E : Exception do
   begin
    NotifyError(cErrPrg ,e.message) ;
    raise ;
   end;
 end ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/02/2007
Modifié le ... :   /  /    
Description .. : - LG - 22/02/2007 - E_IO=X vu avec manou
Mots clefs ... : 
*****************************************************************}
procedure TZReclassement.ExecuteEta ;
var
 lStSQL   : string ;
 lLigneCR : TCRPiece ; // defini ds ImRapInt
begin

 try

 lStSQL := 'UPDATE ECRITURE SET E_ETABLISSEMENT="' + FStEtaDest + '", ' +
           'E_IO="X",E_CREERPAR="REC" ' ;

  if ( FPSource^.E_MODESAISIE = '-' ) or ( FPSource^.E_MODESAISIE = '' ) then
  begin // mode piece
   lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
             ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
             ' AND E_EXERCICE="'              + FPSource^.TExo.Code               + '"' +
             ' AND E_DATECOMPTABLE="'         + USDateTime(FPSource^.E_DATECOMPTABLE) + '"' +
             ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
  end
   else
    begin  // mode bordereau
     lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
               ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
               ' AND E_EXERCICE="'              + FPSource^.TExo.Code              + '"' +
               ' AND E_DATECOMPTABLE>="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, 1)) + '"' +
               ' AND E_DATECOMPTABLE<="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '"' +
               ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
    end; // if

 ExecuteSql(lStSQL) ;

 lStSQL := 'UPDATE ANALYTIQ SET Y_ETABLISSEMENT="' + FStEtaDest + '" ' ;

 if ( FPSource^.E_MODESAISIE = '-' ) or ( FPSource^.E_MODESAISIE = '' ) then
  begin // mode piece
   lStSQL := lStSQL +
             'WHERE Y_JOURNAL="'      + FPSource^.E_JOURNAL                    + '" ' +
             'AND Y_EXERCICE="'       + FPSource^.TExo.Code                    + '" ' +
             'AND Y_NUMEROPIECE='     + intToStr(FPSource^.E_NUMEROPIECE)      + ' ' +
             'AND Y_DATECOMPTABLE="'  + USDateTime(FPSource^.E_DATECOMPTABLE)  + '" ' +
             'AND Y_QUALIFPIECE="'    + FPSource^.E_QUALIFPIECE                + '" '   ;
  end
   else
    begin  // mode bordereau
      lStSQL := lStSQL +
                'WHERE Y_JOURNAL="'      + FPSource^.E_JOURNAL                    + '" ' +
                'AND Y_EXERCICE="'       + FPSource^.TExo.Code                    + '" ' +
                'AND Y_NUMEROPIECE='     + intToStr(FPSource^.E_NUMEROPIECE)      + ' ' +
                'AND Y_DATECOMPTABLE>="' + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, 1)) + '" ' +
                'AND Y_DATECOMPTABLE<="' + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '" ' +
                'AND Y_QUALIFPIECE="'    + FPSource^.E_QUALIFPIECE                + '" '   ;
    end; // if

 ExecuteSql(lStSQL) ;


   // Mise à jour des IMMO pour les Comptes IMMO
{$IFDEF AMORTISSEMENT}
  // GCO - 06/04/2006 - FQ 15240 - Mise à jour des immobilisations

   lStSQL := 'UPDATE IMMO SET I_ETABLISSEMENT="' + FStEtaDest + '" ' ;
   lStSQL := lStSQL + ' WHERE I_IMMO IN ( SELECT E_IMMO FROM ECRITURE ' ;

    if ( FPSource^.E_MODESAISIE = '-' ) or ( FPSource^.E_MODESAISIE = '' ) then
     begin // mode piece
      lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
                ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
                ' AND E_EXERCICE="'              + FPSource^.TExo.Code               + '"' +
                ' AND E_DATECOMPTABLE="'         + USDateTime(FPSource^.E_DATECOMPTABLE) + '"' +
                ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
     end
      else
       begin  // mode bordereau
        lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
                  ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
                  ' AND E_EXERCICE="'              + FPSource^.TExo.Code              + '"' +
                  ' AND E_DATECOMPTABLE>="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, 1)) + '"' +
                  ' AND E_DATECOMPTABLE<="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '"' +
                  ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
       end; // if

    lStSQL := lStSQL + ') ' ;

    ExecuteSql(lStSQL) ;


{$ENDIF}


 // on stocke l'ecriture integree
 lLigneCR             := TCRPiece.Create ;
 lLigneCR.NumPiece    := FPSource^.E_NUMEROPIECE ;
 lLigneCR.Journal     := FPSource^.E_JOURNAL ;
 lLigneCR.QualifPiece := FPSource^.E_QUALIFPIECE ;
 lLigneCR.Date        := FPSource^.E_DATECOMPTABLE ;
 lLigneCR.Debit       := -1 ;
 lLigneCR.Credit      := -1 ;
 FCR.Add (lLigneCR);

 except
  on E : Exception do
   begin
    NotifyError(cErrPrg ,e.message) ;
    raise ;
   end;
 end ;

end;


function _MakeCrit ( vTOB : TOB ) : string ;
begin
 result := vTOB.GetValue('E_EXERCICE')                 +
           vTOB.GetValue('E_JOURNAL')                  +
           vTOB.GetValue('E_ETABLISSEMENT')            +
           DatetoStr(vTOB.GetValue('E_DATECOMPTABLE')) +
           vTOB.GetValue('E_NATUREPIECE') ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 10/05/2005
Modifié le ... :   /  /    
Description .. : - 10/05/2005 - FB 15652 - on arrrondi le solde avant de le 
Suite ........ : tester
Mots clefs ... : 
*****************************************************************}
function TZReclassement.ControleBor : boolean ;
var
 lInNumGroupEcr : integer ;
 lRdSolde      : double ;
 lInLastCrit   : string ;
 i             : integer ;
begin

 lRdSolde       := 0 ;
 lInLastCrit    := _MakeCrit(FTOBDestination.Detail[0]) ;
 lInNumGroupEcr := 1 ;
 result         := false ;

 for i :=  0 to FTOBDestination.Detail.Count - 1 do
  begin
   if lInLastCrit = _MakeCrit(FTOBDestination.Detail[i]) then
    begin
     FTOBDestination.Detail[i].PutValue('E_NUMGROUPEECR', lInNumGroupEcr ) ;
     lRdSolde := Arrondi( lRdSolde +  FTOBDestination.Detail[i].GetValue('E_DEBIT') - FTOBDestination.Detail[i].GetValue('E_CREDIT') , V_PGI.OkDecV ) ;
    end
     else
       if lRdSolde <> 0 then
        begin
         NotifyError( cErrTransfert ,  cErrEquil  ) ;
         exit ;
        end
        else
         begin
          Inc(lInNumGroupEcr) ;
          FTOBDestination.Detail[i].PutValue('E_NUMGROUPEECR', lInNumGroupEcr ) ;
          lInLastCrit := _MakeCrit(FTOBDestination.Detail[i]) ;
          lRdSolde    := Arrondi( FTOBDestination.Detail[i].GetValue('E_DEBIT') - FTOBDestination.Detail[i].GetValue('E_CREDIT') , V_PGI.OkDecV ) ;
         end ;
  end ; // for

 CNumeroPiece(FTOBDestination) ;

 result := true ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : - FB 13412 - LG - 19/08/2004 
Mots clefs ... : 
*****************************************************************}
function TZReclassement.EstCorrectJournal : boolean ;
var
 lIndexSource : integer ;
 lIndexDest   : integer ;
begin

 result := false ;

 lIndexSource := Info.Journal.Load([FStJournalSource]) ;
 lIndexDest   := Info.Journal.Load([FStJournalDest]) ;

 if lIndexSource = - 1 then  // journal de depart
  begin
   NotifyError( RC_JALINEXISTANT , '' ) ;
   exit ;
  end ;

 if lIndexDest = - 1 then // journal d'arrive
  begin
   NotifyError( RC_JALINEXISTANT , '' ) ;
   exit ;
  end ;

 if ( (Info.Journal.GetValue('J_NATUREJAL',lIndexSource)='BQE') or (Info.Journal.GetValue('J_NATUREJAL',lIndexSource)='CAI')) then
  begin
  if  (Info.Journal.GetValue('J_NATUREJAL',lIndexDest)<>'OD') and ( Info.Journal.GetValue('J_CONTREPARTIE',lIndexDest) <> Info.Journal.GetValue('J_CONTREPARTIE',lIndexSource) ) then
   begin
     NotifyError( cErrTransfert , cErrCompteContre ) ;
    exit ;
   end;
   if (Info.Journal.GetValue('J_NATUREJAL',lIndexDest)<>'OD') and ( Info.Journal.GetValue('J_NATUREJAL',lIndexDest) <> Info.Journal.GetValue('J_NATUREJAL',lIndexSource)) then
    begin
     NotifyError( cErrTransfert , 'Reclassement impossible : la nature du compte d''arrivée doit être banque ou OD' ) ;
     exit ;
    end ;
  end
   else
    begin
    if (Info.Journal.GetValue('J_NATUREJAL',lIndexDest)<>'OD') and ( Info.Journal.GetValue('J_NATUREJAL',lIndexDest) <> Info.Journal.GetValue('J_NATUREJAL',lIndexSource)) then
    begin
     NotifyError( cErrTransfert , cErrJalNat ) ;
     exit ;
    end;
  end ;

 if (Info.Journal.GetValue('J_FERME',lIndexDest) = 'X') then
  begin
   NotifyError( RC_JALFERME , '' ) ;
   exit ;
  end; // if

 if Info.Journal.GetValue('J_MULTIDEVISE',lIndexSource) <> Info.Journal.GetValue('J_MULTIDEVISE',lIndexDest) then
  begin
   NotifyError( cErrTransfert , cErrParamJal ) ;
   exit ;
  end; // if

 result := true ;

end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 08/09/2005
Modifié le ... : 08/09/2005
Description .. : -suppression du test sur le journal ds contrepartie
Mots clefs ... : 
*****************************************************************}
function TZReclassement.EstCorrectCompteSource : boolean ;
var
 lIndex : integer ;
begin

 result      := false ;

 if FStGenSource = '' then exit ;

 if ( Info.Compte.GetCompte(FStGenSource) < 0 ) then
  begin
   NotifyError( cErrTransfert , cErrCompteInc ) ;
   exit ;
  end ;// if

 if ( EstSpecif('51209') and Info.Compte.IsCollectif ) or Info.Compte.IsVentilable then
  begin
   NotifyError( cErrTransfert , cErrCompte ) ;
   exit ;
  end ;

 lIndex := Info.Journal.Load([FStJournalSource]) ;
 if ( lIndex > - 1 ) and ( Info.Journal.GetValue('J_CONTREPARTIE' , lIndex) = FStGenSource ) then
  begin
   NotifyError( cErrTransfert , cErrCompteContreJ ) ;
   exit ;
  end ;

 result := true ;

end;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 23/11/2006
Modifié le ... : 29/11/2007
Description .. : - LG - 23/11/2006 - FB 16403 - on autorise de choisir le 
Suite ........ : meme compte pour unreclassement par auxi
Suite ........ : - LG - 29/11/2007 - FB 21957 - chg du test pour les 
Suite ........ : TIC/TID sur le lettrage
Mots clefs ... : 
*****************************************************************}
function TZReclassement.EstCorrectCompteDest : boolean ;
var
 lIndexSource : integer ;
 lIndexDest   : integer ;
begin

 result := false ;

 if FStGenDest = '' then
  begin
   NotifyError( cErrTransfert , cErrCompteInc ) ;
   exit ;
  end ;

 lIndexSource  := Info.Compte.GetCompte(FStGenSource) ;
 lIndexDest    := Info.Compte.GetCompte(FStGenDest) ;

 if ( FStGenDest = FStGenSource ) and not ( FTypeTrait in [trAuxLigne,trAux] )  then
  begin
   NotifyError( cErrTransfert , cErrCompteIdent ) ;
   exit ;
  end ;

 if ( lIndexDest < 0 ) or (  lIndexSource < 0  ) then
  begin
   NotifyError( cErrTransfert , cErrCompteInc ) ;
   exit ;
  end ;// if


 if ( ( Info.Compte.GetValue('G_NATUREGENE',lIndexSource) = Info.Compte.GetValue('G_NATUREGENE',lIndexDest) ) and
      ( Info.Compte.GetValue('G_NATUREGENE',lIndexSource) = 'TID' ) or ( Info.Compte.GetValue('G_NATUREGENE',lIndexSource) = 'TIC' ) ) and
  ( Info.Compte.GetValue('G_LETTRABLE',lIndexSource) <> Info.Compte.GetValue('G_LETTRABLE',lIndexDest) ) then
  begin
   NotifyError( cErrTransfert , cErrCompteL  ) ;
   exit ;
  end;



 if EstSpecif('51209') and ( ( Info.Compte.GetValue('G_POINTABLE',lIndexSource) <> Info.Compte.GetValue('G_POINTABLE',lIndexDest) ) ) and
     ( not VH^.PointageJal )  then
  begin
   NotifyError( cErrTransfert , cErrCompteP  ) ;
   exit ;
  end;


 if ( ( Info.Compte.GetValue('G_NATUREGENE',lIndexSource) = 'IMO' ) or ( Info.Compte.GetValue('G_NATUREGENE',lIndexDest) = 'IMO' ) ) and
    ( Info.Compte.GetValue('G_NATUREGENE',lIndexSource) <> Info.Compte.GetValue('G_NATUREGENE',lIndexDest) ) then
   begin
    NotifyError( cErrTransfert , cErrCompteI  ) ;
    exit ;
   end ; 

 Result := true ;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF AMORTISSEMENT}
function TZreclassement.EstCorrectCompteImmo: boolean;
var lCompteAss : TCompteAss;
    lList      : TList;
begin
  Result := True;
  // Vérification des comptes associés du compte que l'on désire enregistrer
  immocpte_tom.RecupereComptesAssocies(nil, FStGenDest, lCompteAss);
  lList := TList.Create;
  immocpte_tom.TesteExistenceComptesAssocies( lCompteAss, lList);
  if lList.Count <> 0 then
  begin
    PgiInfo('Traitement impossible. Certains comptes associés du compte ' + FStGenDest + ' n''existent pas.',
            'Mise à jour des immobilisations');
    Result := False;
  end;
  FreeAndNil(lList);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TZReclassement.TraitementImmo : Boolean;
var i : integer;
    lBoImmoPasTraite : Boolean;
    lQuery : TQuery;
    lCompteAss : TCompteAss;
begin
  Result := True;
  lBoImmoPasTraite := False;
  immocpte_tom.RecupereComptesAssocies(nil, FStGenDest, lCompteAss);

  for i := 0 to FTOBDestination.Detail.Count -1 do
  begin
    if Trim(FTOBDestination.Detail[i].GetString('E_IMMO')) = '' then Continue;

    lQuery := OpenSql('SELECT * FROM IMMO WHERE I_IMMO = "' + FTOBDestination.Detail[i].GetString('E_IMMO') + '"', False);

    if not lQuery.Eof then
    begin
      if (lQuery.FindField('I_DATEPIECEA').AsDateTime >= VH^.Encours.Deb) and
         (lQuery.FindField('I_DATEPIECEA').AsDateTime <= VH^.Encours.Fin) and
         (lQuery.FindField('I_ETAT').AsString = 'OUV') and
         (lQuery.FindField('I_OPERATION').AsString = '-') then
      begin
        if ExecuteSQL('UPDATE IMMO SET ' +
                      'I_COMPTEIMMO = "' + FStGenDest + '",' +
                      'I_COMPTEAMORT = "' + lCompteAss.Amort + '",' +
                      'I_COMPTEDOTATION = "' + lCompteAss.Dotation + '",' +
                      'I_COMPTEDEROG = "' + lCompteAss.Derog + '", ' +
                      'I_REPRISEDEROG = "' + lCompteAss.RepriseDerog + '", ' +
                      'I_PROVISDEROG = "' + lCompteAss.ProvisDerog + '", ' +
                      'I_DOTATIONEXC = "' + lCompteAss.DotationExcep + '", ' +
                      'I_VACEDEE = "' + lCompteAss.VaCedee + '", ' +
                      'I_AMORTCEDE = "' + lCompteAss.AmortCede + '", ' +
                      'I_REPEXPLOIT = "' + lCompteAss.RepriseExploit + '", ' +
                      'I_REPEXCEP = "' + lCompteAss.RepriseExcep + '", ' +
                      // GCO - 30/11/2006 - 19103
                      'I_ETABLISSEMENT = "' + FTobDestination.Detail[i].GetString('E_ETABLISSEMENT') + '", ' +
                      'I_VAOACEDEE = "' + lCompteAss.VoaCede + '" ' +
                      'WHERE I_IMMO = "' + lQuery.FindField('I_IMMO').AsString + '"') <> 1 then
        begin
          Result := False;
          Ferme(lQuery);
          Break;
          PgiError('Erreur pendant le traitement de l''immobilisation.', 'Mise à jour des immobilisations');
        end;
      end
      else
      begin
        if (not lBoImmoPasTraite) then
          lBoImmoPasTraite := True;
      end;
    end;

    Ferme(lQuery);
  end; // for i := 0

  if lBoImmoPasTraite then
    PgiInfo('Les immobilisations qui ont subi des opérations n''ont pas été traitées.','Mise à jour des immobilisations');
end;

{$ENDIF}
////////////////////////////////////////////////////////////////////////////////

function TZReclassement.EstCorrectAux : boolean ;
begin

 result := false ;

 if ( FStAuxDest = '' ) or ( not Info.LoadAux(FStAuxDest)  ) then
  begin
   NotifyError( cErrTransfert , cErrAuxInc  ) ;
   exit ;
  end ;

 FStAuxDest := Info.StAux ;
 Result     := true ;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/02/2007
Modifié le ... :   /  /    
Description .. : - LG - 22/02/2007 - E_IO=X vu avec manou
Mots clefs ... : 
*****************************************************************}
function TZReclassement.ExecuteTransfertJal( vTOB : TOB ) : boolean ;
var
 lInCodeErreur    : integer ;
 i                : integer ;
 j                : integer ;
 lStError         : string ;
begin

 result := false ;

 try

 vTOB.PutValue('E_JOURNAL'    , FStJournalDest) ;
 vTOB.PutValue('E_IO'         , 'X') ;
 vTOB.PutValue('E_CREERPAR'   , 'REC') ;
 vTOB.PutValue('E_MODESAISIE' , Info.Journal.GetValue('J_MODESAISIE') ) ;

  if _IsValidJournalPeriode( vTOB.GetValue('E_EXERCICE') ,
                    Info.Journal.GetValue('J_VALIDEEN')  ,
                    Info.Journal.GetValue('J_VALIDEEN1') ,
                    vTOB.GetValue('E_DATECOMPTABLE') , Info)  <> RC_PASERREUR then
  begin
   NotifyError( cErrTransfert , cErrJalVal ) ;
   exit ;
  end; // if


 lInCodeErreur := CIsValidNat(vTOB,Info) ;
 if not ( lInCodeErreur = RC_PASERREUR ) then
   begin
    NotifyError( lInCodeErreur , '' ) ;
    V_PGI.IOError := oeSaisie ;
    exit ;
   end ;

 if EstInterdit( Info.Journal.GetValue('J_COMPTEINTERDIT') , vTOB.GetValue('E_GENERAL') , 0 ) > 0 then
  begin
   lStError      := FStJournalDest + ';' + vTOB.GetValue('E_GENERAL') ;
   NotifyError(RC_CINTERDIT , lStError );
   V_PGI.IOError := oeSaisie ;
   exit;
  end ;

  // GCO - 04/12/2006 - FQ 19118
  if (vTob.GetString('E_IMMO') <> '') and (Info.Journal.GetValue('J_NATUREJAL') <> 'ACH') then
  begin
    lStError := 'Reclassement impossible : L''écriture à reclasser contient un lien avec une fiche d''immobilisation.' + #13#10 +
                'Le journal choisi doit être de nature achat.';
    NotifyError(-1 , lStError );
    V_PGI.IOError := oeSaisie ;
    exit;
  end;

 // changement de l'analytique
 for i := 0 to vTOB.Detail.Count - 1 do
  for j := 0 to vTOB.Detail[i].Detail.Count - 1 do
   vTOB.Detail[i].Detail[j].PutValue('Y_JOURNAL',FStJournalDest) ;

 result := true ;

 except
  on E : Exception do NotifyError(cErrPrg ,'Erreur de transfert ' + E.message,'ExecuteTransfertJal');
 end; // try

end ;


procedure TZReclassement.ExecuteJal ;
begin

 if FTOBDestination.Detail.Count = 0 then exit ;

 if Info.Journal.Load([FStJournalDest]) = -1  then  // journal de destination
  begin
   NotifyError( RC_JALINEXISTANT , '' ) ;
   exit ;
  end ;

  if ( Info.Journal.GetValue('J_MODESAISIE') = 'BOR' ) or ( Info.Journal.GetValue('J_MODESAISIE') = 'LIB' ) then
   ExecuteJalBor
    else
     if Info.Journal.GetValue('J_MODESAISIE') = '-' then
      ExecuteJalPiece ;

end;

procedure TZReclassement.ExecuteJalBor ;
var
 i             : integer ;
 lStCodeDevise : string ;
begin

 // info.Journal est possitionner sur le journal de destination

 if ( Info.Journal.GetValue('J_MODESAISIE') = 'BOR' ) and not ControleBor then exit ; // controle les soldes au jours

 lStCodeDevise := FTOBDestination.Detail[0].GetValue('E_DEVISE') ;

 // controle de la validite du changement
 for i := 0 to FTOBDestination.Detail.Count - 1 do
  begin

   if not ExecuteTransfertJal(FTOBDestination.Detail[i]) then exit ; // change le journal ds les tob

   if ( Info.Journal.GetValue('J_MODESAISIE') = 'LIB' ) then
    begin
     if lStCodeDevise <> FTOBDestination.Detail[i].GetValue('E_DEVISE') then
      begin
       NotifyError( cErrTransfert , cErrMonoDev ) ;
       exit ;
      end ;

      FTOBDestination.Detail[i].PutValue('E_NUMGROUPEECR', 1 ) ;
    end ;

  end; // for

 if not CEnregistreSaisie(FTOBDestination,true,false,true,Info) then
  begin
   V_PGI.IOError := oeSaisie ;
   exit ;
  end;

 StockeLesEcrituresIntegrees(FTOBDestination) ;

 if not FBoDuplication then
  begin
   MajSoldesEcritureTOB (FTOBSource,false) ;
   FTOBSource.DeleteDB(false) ;
  end; // if

  V_PGI.IOError := oeOk ;

end;

procedure TZReclassement.ExecuteJalPiece ;
var
 lTOBPiece     : TOB ;
 lTOBGarbage   : TOB ;
 lRdSolde      : double ;
 lInLastCrit   : string ;
 i             : integer ;
 lRecError     : TRecError ;
begin

 lTOBPiece      := TOB.Create('',nil,-1) ;
 lTOBGarbage    := TOB.Create('',nil,-1) ;
 lRdSolde       := 0 ;
 lInLastCrit    := _MakeCrit(FTOBDestination.Detail[0]) ;

 try

  while FTOBDestination.Detail.Count <> 0 do
   begin

    FTOBDestination.Detail[0].PutValue('E_NUMGROUPEECR', 1 ) ;

    if lInLastCrit = _MakeCrit(FTOBDestination.Detail[0]) then
     begin

      lRdSolde    := lRdSolde + Arrondi( FTOBDestination.Detail[0].GetValue('E_DEBIT') - FTOBDestination.Detail[0].GetValue('E_CREDIT') , V_PGI.OkDecV ) ;
      lInLastCrit := _MakeCrit(FTOBDestination.Detail[0]) ;
      if not ExecuteTransfertJal(FTOBDestination.Detail[0]) then exit ; // change le journal ds les tob

      if Arrondi(lRdSolde ,V_PGI.OkDecV ) = 0 then
       begin
        if FTOBDestination.Detail.Count > 1 then
         lInLastCrit := _MakeCrit(FTOBDestination.Detail[1]) ;
        FTOBDestination.Detail[0].ChangeParent(lTOBPiece,-1) ;
        lRecError := CIsValidSaisiePiece(lTOBPiece,Info) ;
        if lRecError.RC_Error <> RC_PASERREUR then
         begin
          NotifyError(lRecError.RC_Error,'') ;
          V_PGI.IOError := oeSaisie ;
          exit ;
         end ;
         
        if not CEnregistreSaisie(lTOBPiece,true,false,true,Info) then
         begin
          V_PGI.IOError := oeSaisie ;
          exit ;
         end ;

       // while lTOBPiece.Detail.Count <> 0 do
        lTOBPiece.Detail[0].ChangeParent(lTOBGarbage,-1) ;
        lTOBPiece.ClearDetail ;
       end
        else
         FTOBDestination.Detail[0].ChangeParent(lTOBPiece,-1) ;

     end  // if
      else
       begin
        NotifyError( cErrTransfert ,  cErrEquil ) ;
        V_PGI.IOError := oeSaisie ;
        exit ;
       end ;

   end ; // while

  if not FBoDuplication then
   begin
    MajSoldesEcritureTOB  (FTOBSource,false) ;
    FTOBSource.DeleteDB(false) ;
   end; // if

 for i := 0 to lTOBGarbage.Detail.Count - 1 do
  begin
   lTOBGarbage.detail[i].PutValue('E_DEBIT',-1) ;
   lTOBGarbage.detail[i].PutValue('E_CREDIT',-1) ;
   StockeLesEcrituresIntegrees(lTOBGarbage.detail[i],false) ;
  end;

 finally
  lTOBPiece.Free ;
  lTOBGarbage.Free ;
 end ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 07/12/2006
Modifié le ... :   /  /
Description .. : - LG - 07/12/2006 - on bloque le chg de journal si est mvt
Suite ........ : est pointe
Mots clefs ... :
*****************************************************************}
function TZReclassement.IsValid ( vPSource : PInfoTW = nil ) : boolean ;
var
 lP : PInfoTW ;
 lQ : TQuery ;
begin

 lP := vPSource ;
 if lP = nil then lP := FPSource ;
 lQ := nil ;

 if ( not EstSpecif('51209') ) then
  begin
   if FTypeTrait in [trGen , trFolio , trEta ,trGenLigne , trAux , trAuxLigne ] then
    begin
     result := true ;
     exit ;
    end ;
  end
   else
    if ( FTypeTrait = trGen ) then begin result := true ; exit ; end ;

 try
  // recherche si il y a pas une ecriture de lettre, pointe
  lQ := OpenSql('SELECT COUNT(*) N FROM ECRITURE WHERE ' +
               'E_EXERCICE="'        + lP^.TExo.Code                + '" ' +
               'AND E_JOURNAL="'     + lP^.E_JOURNAL                + '" ' +
               'AND E_NUMEROPIECE='  + intToStr(lP^.E_NUMEROPIECE)  + ' '  +
               'AND E_QUALIFPIECE="' + lP^.E_QUALIFPIECE            + '" ' +
               'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(lP^.InAnnee, lP^.InMois, 1)) + '" ' +
               'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(lP^.InAnnee, lP^.InMois, 1))) + '" ' +
               'AND ( (E_REFPOINTAGE<>"' + '" ) OR (E_LETTRAGE<>"' + '") OR (E_VALIDE="X" ) ) ', true ) ;
               // GCO - 15240
             //  'AND ( (E_REFPOINTAGE<>"' + '" ) OR (E_LETTRAGE<>"' + '") OR (E_VALIDE="X" ) OR ( E_IMMO<>"' + '"  ) ) ', true ) ;
 result := lQ.findField('N').asInteger = 0 ;
 finally
  Ferme(lQ) ;
 end;
 
end;

procedure TZReclassement.LoadParJournal( vP : PInfoTW ) ;
var
 lQ     : TQuery ;
 lP     : PInfoTW ;
 lYear  : Word ;
 lMonth : Word ;
 lDay   : Word ;
begin

// lStSQL :=  ;

 lQ := nil ;

 try

  lQ := OpenSql( _MakeSQLBor( vP ),true) ;

  while not lQ.Eof do
   begin
    New(lP) ;
    CPInitPInfoTW( lP ) ;
    lP^.TExo.Code          := lQ.findField('E_EXERCICE').asString ;
    lP^.E_JOURNAL          := lQ.findField('E_JOURNAL').asString ;
    lP^.E_NUMEROPIECE      := lQ.findField('E_NUMEROPIECE').asInteger ;
    lP^.E_QUALIFPIECE      := lQ.findField('E_QUALIFPIECE').asString ;
    lP^.E_MODESAISIE       := lQ.findField('E_MODESAISIE').asString ;
    lP^.E_NUMECHE          := lQ.findField('E_NUMECHE').asInteger ;
    lP^.E_DATECOMPTABLE    := lQ.findField('E_DATECOMPTABLE').asDateTime ;
    lP^.E_ETABLISSEMENT    := lQ.findField('E_ETABLISSEMENT').asString ;
    DecodeDate(lQ.findField('E_DATECOMPTABLE').asDateTime,lYear,lMonth,lDay) ;
    lP^.InAnnee            := lYear ;
    lP^.InMois             := lMonth ;
    lP^.InNiv              := 3 ;
    Add(lP) ;
    lQ.Next ;
   end ;// while

 finally
  Ferme(lQ) ;
 end;

end ;

procedure TZReclassement.LoadParJal( vP : PInfoTW ) ;
var
 lQ     : TQuery ;
 lP     : PInfoTW ;
 lYear  : Word ;
 lMonth : Word ;
 lDay   : Word ;
begin

// lStSQL :=  ;

 if vP^.InAnnee = 0 then exit ;

 lQ     := nil ;

 try

  lQ := OpenSql('SELECT E_EXERCICE,E_JOURNAL,E_NUMEROPIECE,E_QUALIFPIECE,E_MODESAISIE,E_NUMECHE,E_DATECOMPTABLE,E_ETABLISSEMENT,E_VALIDE FROM ECRITURE, JOURNAL ' +
                'WHERE E_EXERCICE="' + vP^.TExo.Code + '" ' +
                'AND E_NUMLIGNE=1 AND E_NUMECHE<=1 ' +
                'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(vP^.InAnnee, vP^.InMois, 1))+'" ' +
                'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(vP^.InAnnee, vP^.InMois, 1))) + '" ' +
                'AND E_QUALIFPIECE="N" ' +
                'AND E_JOURNAL=J_JOURNAL ' +
                'AND J_NATUREJAL<>"CLO" ' +
                ' ORDER BY E_NUMEROPIECE' , true );

  while not lQ.Eof do
   begin
    New(lP) ;
    CPInitPInfoTW( lP ) ;
    lP^.TExo.Code          := lQ.findField('E_EXERCICE').asString ;
    lP^.E_JOURNAL          := lQ.findField('E_JOURNAL').asString ;
    lP^.E_NUMEROPIECE      := lQ.findField('E_NUMEROPIECE').asInteger ;
    lP^.E_QUALIFPIECE      := lQ.findField('E_QUALIFPIECE').asString ;
    lP^.E_MODESAISIE       := lQ.findField('E_MODESAISIE').asString ;
    lP^.E_NUMECHE          := lQ.findField('E_NUMECHE').asInteger ;
    lP^.E_DATECOMPTABLE    := lQ.findField('E_DATECOMPTABLE').asDateTime ;
    lP^.E_ETABLISSEMENT    := lQ.findField('E_ETABLISSEMENT').asString ;
    DecodeDate(lQ.findField('E_DATECOMPTABLE').asDateTime,lYear,lMonth,lDay) ;
    lP^.InAnnee            := lYear ;
    lP^.InMois             := lMonth ;
    lP^.InNiv              := 3 ;
    Add(lP) ;
    lQ.Next ;
   end ;// while

 finally
  Ferme(lQ) ;
 end;

end ;



{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 06/06/2006
Modifié le ... :   /  /    
Description .. : LG - 06/06/2006 - FB 16403 - il mq un relation entre la table 
Suite ........ : ecriture et journal. la mqj etait multiplie par la nbnr de compta
Mots clefs ... : 
*****************************************************************}
function TZreclassement.LoadParFolio : boolean;
var
 lQ     : TQuery ;
 lStSQL : string;
 i      : integer ;
begin

 if ( FTypeTrait = trEta ) or ( FTypeTrait = trFolio ) then
  begin
   result := true ;
   exit ;
  end;

 result := false ;

 if FPSource = nil then
  begin
   NotifyError(-1,cErrPasInfo,'TZreclassement.Load') ;
   exit ;
  end;

 try

 FTOBSource.ClearDetail ;
 FTOBDestination.ClearDetail ;

 lStSQL := CGetSQLFromTable('ECRITURE',['E_BLOCNOTE','E_TRACE'] ) ;

 if FTypeTrait in [trGen,trAux,trAuxLigne] then
  lStSQL := lStSQL + ' ,JOURNAL ' ;

 if ( FPSource^.E_MODESAISIE = '-' ) or ( FPSource^.E_MODESAISIE = '' ) then
  begin // mode piece
   lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
             ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
             ' AND E_EXERCICE="'              + FPSource^.TExo.Code               + '"' +
             ' AND E_DATECOMPTABLE="'         + USDateTime(FPSource^.E_DATECOMPTABLE) + '"' +
             ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
  end
   else
    begin  // mode bordereau
     lStSQL := lStSQL + ' WHERE E_NUMEROPIECE=' + intToStr(FPSource^.E_NUMEROPIECE) +
               ' AND E_JOURNAL="'               + FPSource^.E_JOURNAL               + '"' +
               ' AND E_EXERCICE="'              + FPSource^.TExo.Code               + '"' +
               ' AND E_DATECOMPTABLE>="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, 1)) + '"' +
               ' AND E_DATECOMPTABLE<="'        + USDateTime(EncodeDate(FPSource^.InAnnee, FPSource^.InMois, DaysPerMonth(FPSource^.InAnnee, FPSource^.InMois))) + '"' +
               ' AND E_QUALIFPIECE="'           + FPSource^.E_QUALIFPIECE + '"' ;
    end; // if

  if FTypeTrait in [trGen,trAux,trAuxLigne] then
   lStSQL := lStSQL + ' AND E_JOURNAL=J_JOURNAL ' ;

  if FTypeTrait in [trGen,trAux] then
   lStSQL := lStSQL + ' AND E_GENERAL="'     + FStGenSource + '" AND J_CONTREPARTIE<>E_GENERAL AND J_JOURNAL=E_JOURNAL ' ;
  if FTypeTrait in [trGenLigne,trAuxLigne] then
   lStSQL := lStSQL + ' AND E_NUMLIGNE='     + intToStr(FPSource^.E_NUMLIGNE) + ' ' ;

 lStSQL := lStSQL + ' ORDER BY E_JOURNAL, E_EXERCICE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ' ;

 lQ := nil ;

 try
  lQ := OpenSql(lStSQL,true) ;
  FTOBSource.LoadDetailDB('ECRITURE','','',lQ,true) ;
 finally
  Ferme(lQ) ;
 end;

 if FTypeTrait <> trGen then // on ne peut pas changer un compte ventilable
 for i := 0 to FTOBSource.Detail.Count - 1 do
  if FTOBSource.Detail[i].GetValue('E_ANA') = 'X' then
   CChargeAna(FTOBSource.Detail[i]) ;

 FTOBDestination.Dupliquer(FTOBSource,true,true) ;
 result := FTOBDestination.Detail.Count > 0 ;

 if result then
  begin
   FTOBDestination.Detail[0].AddChampSupValeur('SOLDE',0,true) ;
   FTOBDestination.Detail[0].AddChampSupValeur('CRIT',0,true) ;
  end ;

  if FPSource = nil then
  begin
   NotifyError(-1,cErrPasEnrT) ;
   exit ;
  end;

 except
  On E : Exception do
   NotifyError(cErrPrg ,E.Message,'TZreclassement.Load') ;
 end ;

end;


procedure TZreclassement.VideLaListe;
begin
 FListPiece.Clear ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 03/11/2006
Modifié le ... : 31/08/2007
Description .. : - FB 19037 - 03/11/2006 - plus de postmessage si le
Suite ........ : compte n'est pas correct
Suite ........ : - FB 19037 - 31/08/2007 - on affiche pas de msg d'erreur
Mots clefs ... : 
*****************************************************************}
procedure TFReclassement.E2Exit(Sender: TObject);
begin

 if ( csDestroying in self.ComponentState ) then Exit ;
    
 if not ( FTReclass in [trGen,trAux] ) or (E2.Text = '' ) then exit ;

 FZReclass.StGenSource := E1.Text ;
 FZReclass.StGenDest   := E2.Text ;

 FZReclass.TypeTrait   := FTReclass ;

 if not FZReclass.EstCorrectCompteSource then
  begin
   if E1.CanFocus then E1.SetFocus ;
   E1.Text := '' ;
   exit ;
  end ;

 FZReclass.OnError   := nil ;

 try

 if not FZReclass.EstCorrectCompteDest then
  begin
   if E2.CanFocus then E2.SetFocus ;
   E2.Text := '' ;
   exit ;
  end ;

 finally
  FZReclass.OnError   := OnErrorTOB ;
 end ;

 E1.Text := FZReclass.StGenSource ;
 E2.Text := FZReclass.StGenDest ;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/07/2007
Modifié le ... :   /  /    
Description .. : - LG - FB 19037 - 06/07/2007 - plus de controle au niv de
Suite ........ : la saisei, mais lors de la validation
Mots clefs ... :
*****************************************************************}
procedure TFReclassement.E2ElipsisClick(Sender: TObject);
var
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lStCompte     : string ;
 FsStNatGene   : string ;
begin

 if ( csDestroying in self.ComponentState ) then Exit ;

 lStCompte := E1.Text ;
 // on charge les info du compte de depart
 if FInfoEcr.Compte.GetCompte(lStCompte) < 0 then
  begin
   E1.ElipsisClick(nil) ;
   exit ;
  end ;// if 
 FsStNatGene   := FInfoEcr.Compte.GetValue('G_NATUREGENE') ;
 CMakeSQLLookupGen(lStWhere,lStColonne,lStOrder,lStSelect) ;
 lStWhere := lStWhere + //+ 'AND G_GENERAL<>"' + E1.Text + '" ' +
             'AND G_VENTILABLE="-" '  +
             'AND G_LETTRABLE="'      + FInfoEcr.Compte.GetValue('G_LETTRABLE')   + '"' ;
 //  else
 if FTReclass = trAux then
  lStWhere := lStWhere + ' AND ( G_COLLECTIF="X" ) ' ;

 if not LookupList(E2,TraduireMemoire('Comptes'),'GENERAUX',lStColonne,lStSelect,lStWhere,lStOrder,true, 1) then
  begin
   E2.Text := '' ;
   if E2.CanFocus then
    E2.SetFocus ;
  end ;
  
end;

procedure TFReclassement.E1ElipsisClick(Sender: TObject);
var
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lP            : PInfoTW ;
begin
 if ( FTReclass <> trGen ) and ( FTReclass <> trAux ) then exit ;
 lP := nil ;
 if ( TW.Selected <> nil ) and ( TW.Selected.Data <> nil ) then
  lP:= PInfoTW(TW.Selected.Data) ;
 CMakeSQLLookupGen(lStWhere,lStColonne,lStOrder,lStSelect) ;
 lStWhere := lStWhere + ' AND G_VENTILABLE="-" AND G_GENERAL<>"' + lP^.J_CONTREPARTIE + '" ' ;
// if FTReclass = trAux then
//  lStWhere := lStWhere + ' AND G_NATUREGENE="DIV" ' ;
 LookupList(E1,TraduireMemoire('Comptes'),'GENERAUX',lStColonne,lStSelect,lStWhere,lStOrder,true, 1) ;
end;


procedure TFReclassement.E4ElipsisClick(Sender: TObject);
var
 lStSelect     : string ;
 lStColonne    : string ;
 lStOrder      : string ;
 lStWhere      : string ;
 lStCompte     : string ;
begin

 if not ( FTReclass in [trGen,trAux] ) then exit ;

 lStCompte := E2.Text ;

 if FInfoEcr.Compte.GetCompte(lStCompte) < 0 then
  begin
   E1.ElipsisClick(nil) ;
   exit ;
  end ;// if
 FStNatGene   := FInfoEcr.Compte.GetValue('G_NATUREGENE') ;
 CMakeSQLLookupAux(lStWhere,lStColonne,lStOrder,lStSelect,'OD',FStNatGene) ;
 LookupList(E4,TraduireMemoire('Auxiliaire'),'TIERS',lStColonne,lStSelect,lStWhere,lStOrder,true, 2,'',tlLocate) ;

end;


procedure TFReclassement.E1Exit(Sender: TObject);
var
 lP        : PInfoTW ;
begin

 if ( csDestroying in self.ComponentState ) then Exit ;

 lP := nil ;
 if ( TW.Selected <> nil ) and ( TW.Selected.Data <> nil ) then
  lP:= PInfoTW(TW.Selected.Data) ;

 if lP = nil then exit ;

 FZReclass.StJournalSource := UpperCase(lP^.E_JOURNAL) ;
 FZReclass.StGenSource     := E1.Text ;
 FZReclass.TypeTrait       := FTReclass ;

 if not ( FTReclass in [trGen,trAux] ) then exit ;

 FZReclass.OnError   := nil ;

 try

 if not FZReclass.EstCorrectCompteSource then
  begin
   // LG - 31/08/2007 - ne pas appeler le lookupList ca fait planter sur le echap
   E1.Text := '' ;
   E1.SetFocus ;
   exit ;
  end ;

 finally
  FZReclass.OnError   := OnErrorTOB ;
 end ;

 E1.Text := FZReclass.StGenSource ;

end;


{$IFDEF TT}

function _ChargeLigneAna( vTOB : TOB ) : TOB ;
var
 lQ      : TQuery ;
 lStSQL  : string ;
begin
 lStSQL := 'select distinct y_numligne, sum(y_debit)- sum(y_credit) S from analytiq ' +
           'where y_numeropiece=' + intToStr(vTOB.GetValue('e_numeropiece')) +
           'and y_exercice="' + vTOB.GetValue('e_exercice') + '" ' +
           'and y_journal="' + vTOB.GetValue('e_journal') + '" ' +
           'and y_datecomptable="' + usdatetime(vTOB.GetValue('e_datecomptable') ) + '" ' +
           'and y_axe="a1" ' +
           'group by y_axe,y_section,y_numeropiece,y_exercice,y_numligne ' ;
 lQ := openSQL(lStSQL,true) ;
 result := TOB.Create('',nil,-1) ;
 result.LoadDetailDB('ANALYTIQ','','',lQ,true) ;
 ferme(lQ) ;
end ;

function _ChargeLigneAna1( vTOB : TOB ) : TOB ;
var
 lQ      : TQuery ;
 lStSQL  : string ;
begin

lStSQL := 'select distinct y_numligne, sum(y_debit)- sum(y_credit) S from analytiq ' +
          'where y_numeropiece=' + intToStr(vTOB.GetValue('e_numeropiece')) +
          'and y_exercice="' + vTOB.GetValue('e_exercice') + '" ' +
          'and y_journal="' + vTOB.GetValue('e_journal') + '" ' +
          'and y_datecomptable="' + usdatetime(vTOB.GetValue('e_datecomptable') ) + '" ' +
          'and y_axe="a1" ' +
          'and not exists ( select e_numligne from ecriture where e_journal=y_journal and e_exercice=y_exercice ' +
          'and e_numeropiece=y_numeropiece and e_qualifpiece=y_qualifpiece and e_NUMLIGNE=y_NUMLIGNE ' +
          'and e_datecomptable=y_datecomptable ) ' +
          'group by y_axe,y_section,y_numeropiece,y_exercice,y_numligne ' +
          'having( sum(y_debit)- sum(y_credit) =' + VariantToSql(Arrondi( vTOB.GetValue('E_DEBIT') - vTOB.GetValue('E_CREDIT') , V_PGI.OkDecV )) + ') ' ;

 lQ := openSQL(lStSQL,true) ;
 result := TOB.Create('',nil,-1) ;
 result.LoadDetailDB('ANALYTIQ','','',lQ,true) ;
 ferme(lQ) ;
end ;

procedure TZReclassement._ExecuteAna ( vP : PInfoTW ; vBoEnr : boolean = true  ) ;
var
 i    : integer ;
 lTOB : TOB ;
begin

 lTOB := nil ;
 if vP = nil then exit ;

 _flisteD:= tstringlist.Create ;
 _flisteS:= tstringlist.Create ;
 _flisteT:= tstringlist.Create ;
 _flisteSQL:= tstringlist.Create ;

 _FBoEnr := vBoEnr ;

 try

  if vP^.InNiv = 0 then lTOB := _ChargeBorSansAna(vP,trExo)
   else
    if vP^.InNiv = 1 then lTOB := _ChargeBorSansAna(vP,trPeriode)
     else
      if vP^.InNiv = 2 then lTOB := _ChargeBorSansAna(vP,trJal) ;

  if ( lTOB = nil ) or ( lTOB.Detail.Count = 0 ) then exit ;

  InitMoveProgressForm (nil, 'Reclassement', 'Traitement en cours', lTOB.Detail.Count , true , true ) ;

  for i := 0 to lTOB.Detail.Count - 1 do
   begin
    _FTOBAna := _ChargeLigneAna( lTOB.Detail[i] ) ;
    _FTOB    := lTOB.Detail[i] ;
    if not MoveCurProgressForm('traitements du bordereau : Exo ' +
           _FTOB.GetValue('e_exercice') +
           ' Journal : ' + _FTOB.GetValue('e_journal') +
           ' Folio : '   + intToStr(_FTOB.GetValue('e_numeropiece')) +
           ' Date : ' +  usdatetime(_FTOB.GetValue('e_datecomptable') ) ) then exit ;
    if _FTOBAna = nil then continue ;
    if Transactions(_RenumAna,1) <> oeOk then
     begin
      MessageAlerte('erreur') ;
      _FTOBAna.Free ;
      continue ;
     end;
    _FTOBAna.Free ;
   end; // for

  finally
   lTOB.Free ;
   FiniMoveProgressForm ;
   PGIInfo('Traitement terminé.' +#10#13 +#10#13 +
           'Nbr de lignes traitées      : '+intToStr(_fListeT.Count)+#10#13 +
           'Nbr de lignes en doublon    : '+intToStr(_fListeD.Count)+#10#13 +
           'Nbr de lignes sans analytiq : '+intToStr(_fListeS.Count) , 'Erreur' ) ;
   {$I-}
   _flisteD.SaveToFile('c:\pgi00\app\rapport doublon.txt') ;
   _flisteT.SaveToFile('c:\pgi00\app\rapport traite.txt') ;
   _flisteS.SaveToFile('c:\pgi00\app\rapport sans.txt') ;
   _flisteSQL.SaveToFile('c:\pgi00\app\rapport sql.txt') ;
   {$I+}
   _flisteD.Free ;
   _flisteT.Free ;
   _flisteS.Free ;
   _flisteSQL.Free;
  end ;

end;


function TZReclassement._ChargeBorSansAna(vP : PInfoTW ; vReclass : TReclass ) : TOB ;
var
 lQ      : TQuery ;
 lStSQL  : string ;
begin
 lStSQL := 'select e_exercice,e_datecomptable,e_journal,e_numeropiece,e_numligne,e_debit,e_credit,e_datemodif,e_qualifpiece,e_general,e_auxiliaire from ecriture ' +
           'where e_ana="x" and ( e_modesaisie="BOR" or e_modesaisie="LIB" ) and not exists ( select y_numligne from analytiq where e_journal=y_journal and e_exercice=y_exercice ' +
           'and e_numeropiece=y_numeropiece and e_qualifpiece=y_qualifpiece and e_NUMLIGNE=y_NUMLIGNE ' +
           'and e_datecomptable=y_datecomptable ) ' ;
 if vReclass = trJal then
  lStSQL := lStSQL + ' and e_exercice="' + vP^.TExo.Code + '" and e_journal="' + vp^.E_JOURNAL + '" ' +
           'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(vP^.InAnnee, vP^.InMois, 1))+'" ' +
           'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(vP^.InAnnee, vP^.InMois, 1))) + '" '
   else
    if vReclass = trExo then
     lStSQL := lStSQL + ' and e_exercice="' + vP^.TExo.Code + '" '
      else
       if vReclass = trPeriode then
        lStSQL := lStSQL + ' and e_exercice="' + vP^.TExo.Code + '" ' +
        'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(vP^.InAnnee, vP^.InMois, 1))+'" ' +
        'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(vP^.InAnnee, vP^.InMois, 1))) + '" ' ;

 lStSQL := lStSQL + 'order by e_exercice,e_journal,e_numeropiece,e_numligne' ;

 lQ := openSQL(lStSQL,true) ;
 result := TOB.Create('',nil,-1) ;
 result.LoadDetailDB('ECRITURE','','',lQ,true) ;
 ferme(lQ) ;
end ;

function _FaireRapport ( vTOB : TOB ) : string ;
begin
 result := vTOB.GetValue('e_exercice') + ';' +
           vTOB.GetValue('e_journal')  + ';' +
           usdatetime(vTOB.GetValue('e_datecomptable')) + ';' +
           intToStr(vTOB.GetValue('e_numeropiece')) + ';' +
           intToStr(vTOB.GetValue('e_numligne')) + ';' +
           'where E_JOURNAL="'+vTOB.GetValue('E_JOURNAL')+'" AND E_EXERCICE="'+vTOB.GetValue('E_EXERCICE')+'"'
           +' AND E_DATECOMPTABLE="'+UsDateTime(vTOB.GetValue('E_DATECOMPTABLE'))+
           '" AND E_NUMEROPIECE='+InttoStr(vTOB.GetValue('E_NUMEROPIECE'))
           +' AND E_NUMLIGNE='+IntToStr(vTOB.GetValue('E_NUMLIGNE'))  ;
end ;


procedure TZReclassement._RenumAna ;
var
 i : integer ;
 lSol : integer ;
 lInNum : integer ;
 lInNumAna : integer ;

begin

 try
 lInNum := 0 ;
 lInNumAna := 0 ;
 lSol := 0 ;
 for i := 0 to _FTOBAna.Detail.count - 1 do
  begin
    if _FTOBAna.Detail[i].GetValue('S') = Arrondi(_FTOB.GetValue('e_debit')-_FTOB.GetValue('e_credit'),2) then
     begin
      lInNum := _FTOB.GetValue('e_numligne') ;
      lInNumAna := _FTOBAna.Detail[i].GetValue('y_numligne') ;
      Inc(lSol) ;
     end;
  end ;

 if lSol > 1 then
  begin
   _flisteD.Add(_FaireRapport (_FTOB) ) ;
   exit ;
  end
   else
    if lSol = 0 then
     begin
       _flisteS.Add(_FaireRapport (_FTOB) ) ;
       _flisteSQL.Add('WHERE E_JOURNAL="' + _FTOB.GetValue('E_JOURNAL')      + '" ' +
       'AND E_EXERCICE="'            + _FTOB.GetValue('E_EXERCICE')                  + '" ' +
        'AND E_DATECOMPTABLE="'       + usDateTime(_FTOB.GetValue('E_DATECOMPTABLE')) + '" ' +
        'AND E_NUMEROPIECE='          + intToStr(_FTOB.GetValue('E_NUMEROPIECE'))     + ' '  +
        'AND E_QUALIFPIECE="'         + _FTOB.GetValue('E_QUALIFPIECE')               + '" ' +
        'AND E_DEBIT='                + VariantToSql(Arrondi(_FTOB.GetValue('E_DEBIT'),3)) + ' ' +
        'AND E_CREDIT='               + VariantToSql(Arrondi(_FTOB.GetValue('E_CREDIT'),3)) + ' ' +
        'AND E_GENERAL="'             + _FTOB.GetValue('E_GENERAL') + '" ' +
        'AND E_AUXILIAIRE="'          + _FTOB.GetValue('E_AUXILIAIRE') + '" ' ) ;
       exit ;
     end;

 if _FBoEnr then
  begin
     if executeSQL('UPDATE ECRITURE SET E_DATEMODIF="'+UsTime(NowH)+'"'
      +' Where '
      +' E_JOURNAL="'+_FTOB.GetValue('E_JOURNAL')+'" AND E_EXERCICE="'+_FTOB.GetValue('E_EXERCICE')+'"'
      +' AND E_DATECOMPTABLE="'+UsDateTime(_FTOB.GetValue('E_DATECOMPTABLE'))+'" AND E_NUMEROPIECE='+InttoStr(_FTOB.GetValue('E_NUMEROPIECE'))
      +' AND E_QUALIFPIECE="'+_FTOB.GetValue('E_QUALIFPIECE')+'" '
      +' AND E_NUMLIGNE='+IntToStr(_FTOB.GetValue('E_NUMLIGNE'))
      +' AND E_DATEMODIF="'+UsTime(_FTOB.GetValue('E_DATEMODIF'))+'" ' ) <> 1 then
     begin
      messageAlerte('Bordereau en cours d''utilisation : Exo ' +
           _FTOB.GetValue('e_exercice') +
           ' Journal : ' + _FTOB.GetValue('e_journal') +
           ' Folio : '   + intToStr(_FTOB.GetValue('e_numeropiece')) +
           ' Date : ' +  usdatetime(_FTOB.GetValue('e_datecomptable') ) );
      exit ;
     end;

   executeSQL('update analytiq set y_numligne=' + intToStr(lInNum) +
            ' where y_numeropiece=' + intToStr(_FTOB.GetValue('e_numeropiece')) +
            'and y_exercice="' + _FTOB.GetValue('e_exercice') + '" ' +
            'and y_journal="' + _FTOB.GetValue('e_journal') + '" ' +
            'and y_datecomptable="' + usdatetime(_FTOB.GetValue('e_datecomptable') ) + '" ' +
            'and y_numligne=' + intToStr(lInNumAna) ) ;
  end ; // if

 _flisteT.Add(_FaireRapport (_FTOB) ) ;


 except
  on e : exception do
   begin
    PGIInfo( e.message ) ;
    raise ;
   end;
 end;


end;


procedure _CreerTOBR ;
var
 lListe : TStringList ;
 lErr : TStringList ;
 i : integer ;
 lQ : TQuery ;
 //lq1 : tquery ;
 lTOB : TOB ;
 lTOBAna : TOB ;
 lNb : integer ;
begin

 lListe :=TStringList.Create ;
 lErr :=TStringList.Create ;
 lListe.LoadFromFile('c:\pgi00\app\rapport sql.txt') ;

 lTOB := TOB.Create('',nil,-1) ;
 lTOBAna:= TOB.Create('',nil,-1) ;

 try

 for i := 0 to lListe.Count - 1 do
  begin

    try
    lQ := OpenSql( 'SELECT * FROM ECRITURE ' + lListe[i] , true ) ;
    {$ifndef eaglclient}
    lNb := QCount(lQ) ;
    {$ENDIF}
    if lNb = 1 then
     begin
      lTOB.SelectDB('',lQ) ;
     end
      else
       if lNb = 0 then
        begin
         lErr.Add('Ligne non retrouvee : ' + lListe[i] ) ;
         continue ;
        end

        else
         begin
          lErr.Add('Ligne en double : ' + lListe[i] ) ;
          continue ;
         end;

     finally
      Ferme(lQ) ;
     end;

     lQ := OpenSql('SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="' + lTOB.GetValue('E_JOURNAL')      + '" ' +
        'AND Y_EXERCICE="'            + lTOB.GetValue('E_EXERCICE')                  + '" ' +
        'AND Y_DATECOMPTABLE="'       + usDateTime(lTOB.GetValue('E_DATECOMPTABLE')) + '" ' +
        'AND Y_NUMEROPIECE='          + intToStr(lTOB.GetValue('E_NUMEROPIECE'))     + ' '  +
        'AND Y_NUMLIGNE='             + intToStr(lTOB.GetValue('E_NUMLIGNE'))        + ' '  +
        'AND Y_QUALIFPIECE="'         + lTOB.GetValue('E_QUALIFPIECE')               + '" ' , true ) ;

    lTOBAna.LoadDetailDB('ANALYTIQ','','',lQ,true) ;
    Ferme(lQ) ;

  end;

 finally

   PGIInfo('Traitement terminé.' +#10#13 +#10#13 +
           'Nbr de lignes récupérees : '+intToStr(lTOBAna.Detail.Count)+#10#13 +
           'Nbr de lignes en doublon ou en erreur  : '+intToStr(lErr.Count)  ) ;

 lTOBAna.SaveToFile('c:\pgi00\app\ana recup.bob',false,true,true) ;
 lErr.SaveToFile('c:\pgi00\app\ana err.txt') ;

 lTOB.Free ;
 lTOBAna.Free ;
 lListe.Free ;
 lErr.Free ;

 end; 

end;

procedure _Insert ;
var
 lTOB : TOB ;
begin


 lTOB := TOB.Create('',nil,-1) ;

 TOBLoadFromFile('c:\pgi00\app\ana recup.bob',nil, lTOB ) ;

 lTOB.InsertDb(nil) ;

 lTOB.Free ;


end;

procedure _Sauve( vP : PInfoTW ) ;
var
 lQ : TQuery ;
 lSQL : string ;
 lTOB : TOB ;
begin

 lSQL := 'select * from ecriture where e_EXERCICE="' + vP^.TExo.Code + '" ' ;
 lSQL := lSQL + 'AND E_DATECOMPTABLE>="' + USDateTime(EncodeDate(vP^.InAnnee, vP^.InMois, 1))+'" ' +
          'AND E_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(vP^.InAnnee, vP^.InMois, 1))) + '" ' ;
 if vP^.InNiv = 2 then
  lSQL := lSQL + 'and e_JOURNAL="' + vP^.E_JOURNAL  + '" ' ;

 lQ := OpenSQL(lSQL,true) ;
 lTOB := TOB.Create('',nil,-1) ;
 lTOB.LoaddetailDB('ECRITURE','','',lQ,true) ;
 Ferme(lQ) ;
 lTOB.SaveToFile('c:\pgi00\app\ecr.bob',false,true,true);
 lTOB.Free ;

 lSQL := 'select * from analytiq where y_EXERCICE="' + vP^.TExo.Code + '" ' ;
 lSQL := lSQL + 'AND y_DATECOMPTABLE>="' + USDateTime(EncodeDate(vP^.InAnnee, vP^.InMois, 1))+'" ' +
          'AND y_DATECOMPTABLE<="' + USDateTime(FinDeMois(EncodeDate(vP^.InAnnee, vP^.InMois, 1))) + '" ' ;
 if vP^.InNiv = 2 then
  lSQL := lSQL + 'and y_JOURNAL="' + vP^.E_JOURNAL  + '" ' ;

 lQ := OpenSQL(lSQL,true) ;
 lTOB := TOB.Create('',nil,-1) ;
 lTOB.LoaddetailDB('analytiq','','',lQ,true) ;
 lTOB.SaveToFile('c:\pgi00\app\ana.bob',false,true,true);
 lTOB.Free ;
 Ferme(lQ) ;

end;

{$ENDIF}



procedure TFReclassement.MMContrClick(Sender: TObject);
var
 lP : PInfoTW ;
begin

 if TW.Selected = nil then exit ;
 lP := PInfoTW(TW.Selected.Data) ;
 if lP = nil then exit ;

{$IFDEF TT}
 FZReclass._ExecuteAna(lP,false) ;
{$ENDIF}
end;

procedure TFReclassement.MMRecupClick(Sender: TObject);
begin
 {$IFDEF TT}
 _CreerTOBR ;
{$ENDIF}
end;

procedure TFReclassement.MMsauvClick(Sender: TObject);
var
 lP : PInfoTW ;
begin
 if TW.Selected = nil then exit ;
 lP := PInfoTW(TW.Selected.Data) ;
 if lP = nil then exit ;
 {$IFDEF TT}
 _Sauve(lP) ;
{$ENDIF}
end;

procedure TFReclassement.MMInsertClick(Sender: TObject);
begin
{$IFDEF TT}
 _Insert ;
{$ENDIF}

end;


{***********A.G.L.***********************************************
Auteur  ...... : LG
Créé le ...... : 03/11/2006
Modifié le ... :   /  /    
Description .. : - FB 19037 - 03/11/2006 - pas de msg si la zone est vide
Mots clefs ... : 
*****************************************************************}
procedure TFReclassement.E4Exit(Sender: TObject);
begin

 if E4.Text = '' then exit ;

 if ( E2.Text = '' ) and ( FInfoEcr.LoadAux(E4.Text) ) then
  E2.Text := FInfoEcr.Aux_Getvalue('T_COLLECTIF') ;

 FZReclass.StAuxDest := E4.Text ;

 if not FZReclass.EstCorrectAux then
  begin
   E4.SetFocus ;
   E4.Text := '' ;
   exit ;
  end ;

 E4.Text := FZReclass.StAuxDest ;

end;

procedure TFReclassement.E2Change(Sender: TObject);
begin
 if not ( FTReclass in [trGen,trAux] ) then exit ;

 E4.Text := '' ;

end;

end.

