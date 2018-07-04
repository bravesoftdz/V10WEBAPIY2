{ Unité : Source TOF de la FICHE : eSaisAnal
--------------------------------------------------------------------------------------
    Version  |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 6.50.001.024  24/10/05   JP  FQ 16879 : gestion du caption de la fiche
 6.50.001.026  13/01/06   JP  FQ 17160 : Indice de liste Hors limite sur le Ctrl + Suppr
 8.10.001.004  14/08/07   JP  FQ 20881 : Enregistrement d'une ventilation type sur tous les axes
--------------------------------------------------------------------------------------}
unit eSaisAnal;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    Grids,
    Hctrls,
    ComCtrls,
    StdCtrls,
    Hspliter,
    ExtCtrls,
    Hcompte,
    Ent1,
    Buttons,
    hmsgbox,
    Menus,
    HEnt1,
    Filtre,
    HSysMenu,
    HTB97,
    ed_tools,
    paramsoc,
{$IFNDEF CCS3}
    UTOFAffGrille,   // CPLanceFicheAFFVentilana
{$ENDIF}
{$IFDEF EAGLCLIENT}
    Maineagl,
{$ELSE}
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    DB,
    Fe_Main,
{$ENDIF}
{$IFNDEF GCGC}
    SaisComp,
{$ENDIF}
{$IFDEF VER150}
    Variants,
{$ENDIF}
    CPSECTION_TOM,
    SaisUtil,
    SaisComm,
 //   LettUtil,
    UtilSais,
    UtilPGI,
    UTob,
    ULibAnalytique,
    ULibExercice,
    ParamDat,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
    ULibEcriture, TntButtons, TntGrids, TntStdCtrls;
    //CPMODRESTANA_TOM;          {FP 29/12/2006: TRestrictionAnalytique}

Type T5D = Array[1..MaxAxe] of Double ;

Type T5LL = Array[1..MaxAxe] Of HTStringList ;

{Type REC_ANA = RECORD
               QuelEcr     : TTypeEcr ;
               CC          : TObject ;
               MontantEcrP : double ;
               MontantEcrD : double ;
               DEV         : RDevise ;
               RefEcr      : String ;
               LibEcr      : String ;
               OBA         : TObjAna ;
               Action      : TActionFiche ;
               NumGeneAxe  : integer ;
               VerifVentil : boolean ;
               END ; }

// Nouvelle structure utilisée pour appeler la saisie analytique
Type ARG_ANA = RECORD
	       // Anciennement dans REC_ANA
               QuelEcr 	       : TTypeEcr ;
               CC 	       : TGGeneral ;	 // TGGeneral
               DEV 	       : RDevise ;	 // Object devise
               Action 	       : TActionFiche ;  // Saisie ou consult ?
               NumGeneAxe      : Integer ;	 // Axe du scénario
               VerifVentil     : Boolean ;	 // Vrai en SaisiBor et Saisie !
               VerifQte        : Boolean ;	 // Dev3946
	       // Anciennement dans TObjANA
               NumLigneDecal   : Integer ;	 // ??
               DernVentilType  : String3 ;	 // Ventilation type
               GuideA 	       : String3 ;	 // Guide utilisé
	       ControleBudget  : Boolean ;	 // Uniquement pour les Ecr budget
               ModeConf	       : String[1] ;	 // Confidentialité
               // New Dev 3946
               Info            : TInfoEcriture ; // Pour récupérer les TInfoEcriture de la pièce
               MessageCompta   : TMessageCompta ;// Pour récupérer les Messages d'erreurs de la pièce
               // New Dev 96xx : Saisie en montant
               EnMontant       : boolean ;
               END ;

Type TOBMSupp = Class(TObject)
                  CompS : Boolean ;
                  end ;

Procedure eSaisieAnal( var ObjEcr : TOB ; var Arguments : ARG_ANA ) ;
Procedure eSaisieAnalEff( ObjEcr : TOB ; LL : T5LL ; Arguments : ARG_ANA ) ;

type
  TFeSaisAnal = class(TForm)
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
   PEntete       : TPanel;
   PPied         : TPanel;
    H_GENERAL    : THLabel;
    G_GENERAL    : THLabel;
    H_REFINTERNE : THLabel;
    E_REFINTERNE : THLabel;
    E_LIBELLE    : THLabel;
    H_SOLDE      : THLabel;
    S_SOLDE      : THNumEdit;
    H_TOTAL      : THLabel;
    H_RESTE      : THLabel;
    E_TOTAL      : THNumEdit;
    E_RESTE      : THNumEdit;
    E_TOTPOURC: THNumEdit;
    E_RESTEPOURC: THNumEdit;
    H_MONTANTECR : THLabel;
    E_MONTANTECR : THNumEdit;
    Cache        : THCpteEdit;
    HMessLigne   : THMsgBox;
    HMessPiece   : THMsgBox;
    POPS         : TPopupMenu;
    PGA          : TPanel;
    GSA          : THGrid;
    TA           : TTabControl;
    CVType       : THValComboBox ;
    TFType       : THLabel;
   PVentil       : TPanel;
    H_CODEVENTIL : THLabel;
    HLabel1      : THLabel;
    H_TITREVENTIl: TLabel;
    Y_CODEVENTIL    : TEdit;
    Y_LIBELLEVENTIL : TEdit;
    BSauveVentil: TToolbarButton97;
    BNewVentil: THBitBtn;
    BInsert: THBitBtn;
    BSDel: THBitBtn;
    PFenCouverture: TPanel;
    HMTrad: THSystemMenu;
    Outils: TToolbar97;
    BSolde: TToolbarButton97;
    BVentilType: TToolbarButton97;
    BComplement: TToolbarButton97;
    BZoom: TToolbarButton97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    ISigneEuro: TImage;
    BCALCULQTE: TToolbarButton97;

    procedure TAChanging(Sender: TObject; var AllowChange: Boolean);
    procedure TAChange(Sender: TObject);

    procedure GSAMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSADblClick(Sender: TObject);
    procedure GSACellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSACellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure GSARowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSARowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
    procedure GSAKeyPress(Sender: TObject; var Key: Char);

    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

    procedure BZoomClick(Sender: TObject);
    procedure BSoldeClick(Sender: TObject);
    procedure BComplementClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BSDelClick(Sender: TObject);
    procedure BVentilTypeClick(Sender: TObject);
    procedure CVTypeChange(Sender: TObject);
    procedure BSauveVentilClick(Sender: TObject);
    procedure BNewVentilClick(Sender: TObject);

    procedure BAideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);

    procedure E_MontantEcrEnter(Sender: TObject);
    procedure E_MontantEcrExit(Sender: TObject);

    procedure POPSPopup(Sender: TObject);
    procedure BCALCULQTEClick(Sender: TObject);
  private
    {JP 13/01/06 : FQ 17160 : Copie de ObjEcr dans sa forme en entrée de fiche pour le cas
                   où l'on annulerait la saisie}
    ObjTemp           : TOB;
    IsCanceled        : Boolean; {Pour dire que l'on annule}
    FMessageCompta    : TMessageCompta ;
    FInfo             : TInfoEcriture ;

    // Objets utilisés pour gérer les axes
    CurAxe            : integer ;      // Axe en cours d'édition
    SI_TotD  				  : T5D ;          // Total devise pour chaque axe
    SI_TotP      		  : T5D ;	         // Total pivot pour chaque axe
    SI_TotPourc       : T5D ; 	       // Total Pourcentage pour chaque axe
    SI_TotSais        : T5D ;    		   // Total devise pour chaque axe
    Axes              : Array[1..MaxAxe] of Boolean ;	 // Axes ventilables ?
    Conf              : Array[1..MaxAxe] of String[1]; // Confidentialité de chaque axe
    LL                : T5LL ;			 // TList su chaque axe ???
    ModeS             : Byte ;			 // Saisie pourcentage ou montant
    MTEntree          : Double ;			 // Montant Ecr Gen saisie
    Pf                : String[2] ;			 // Préfixe de la table Y ou ...? TOUJOURS Y !!!
    Sens              : Byte ;			 // Ecriture au débit (=1) ou au crédit
    ModeVentil : Boolean ;                       // Enregistrement d'une ventilation type en cours
    MontantEcrP,				 // Montant écriture
    MontantEcrD : double; 			 // Montant écriture en devise
    FRestriction: TRestrictionAnalytique;        // Modèlde de restriction ana FP 29/12/2005}
    GX,GY         : integer ;			 // Cooredonnées souris ??
    FInAxeTva     : integer ;
    FBoOuiTvaEnc  : boolean ;
// Positionnement fenêtre
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
// Init et defaut
    procedure LectureSeule ;
    procedure DefautEntete ;
    procedure DesalloueA( Lig : Longint ) ;
    procedure InitGrid ;                                               // OK
    procedure AlloueMvt (Lig : integer ) ;                             // OK
    procedure AlloueOS  (Lig : integer ) ;
    procedure DeAlloueObj( vNumLigne : integer = - 1 ) ;
    procedure ChargeGrid (NumAxe : integer);
// Calculs lignes
    procedure DefautLigne (Lig : integer ) ;                           // OK
    procedure CalculSoldeCompte ( Sect : String ; DIS,CIS : Double ) ;
    procedure CalculTotal ( NumA : integer ) ;
    procedure TraiteMontant ( Lig : integer ; Calc : boolean ) ;
    function  ValM          ( Lig : integer ) : double ;                          // OK
    function  ValP          ( Lig : integer ) : double ;                          // OK
    function  ValTobM       ( Lig : integer ) : double ;                          // OK
    function  ValTobP       ( Lig : integer ) : double ;                          // OK
    procedure GereArrondi   ( Lig : integer ; ColTest : integer ) ;
// Barre Outils
    procedure ClickSolde ;
    procedure ClickInsert ;
    procedure ClickDel ;
    procedure ClickValide ;
    procedure ClickAbandon ;
    procedure ClickComplement ( Lig : integer ) ;
    procedure ClickZoom ;
{$IFNDEF CCS3}
    procedure ClickZoomSousPlan ;
{$ENDIF}
    procedure ZoomVentilType ;
    procedure ClickVentilType ;
    procedure ClickSauveVentil ;
// Appel de sections en saisie, Ventil type
    Function  ChercheSect ( L : integer ; Force : boolean ) : byte ;
    procedure FinVentil ;
// Affichages
    procedure AffichePied  ;
    procedure InsereLigne ( Lig : integer ) ;
    procedure FormatMontant ( ACol,ARow : integer ) ;
    procedure GereComplements ( Lig : integer ) ;
    procedure GereOptionsGrid ( Lig : integer ) ;
    function  PasToucheLigne ( Lig : integer ) : boolean ;
// Controles
    procedure ControleLignes ;
    procedure ErreurSaisie ( Err : integer ) ;
    Function  Equilibre (NumA : integer) : boolean ;
    Function  EstVentilModifiee : boolean ;
    Function  AxeCorrect    ( NumA : integer ; Parole : boolean ) : Boolean ;
    Function  LigneCorrecte ( Lig : integer ; Alim : boolean = True ) : Boolean ;
    function  ValideLeAxe   ( NumA : integer ; Parole : boolean ) : Boolean ;
    procedure OnError       ( sender : TObject; Error : TRecError ) ; // fonction d'affichage des erreurs

    Procedure AlimObjetMvt ( Lig : integer ) ;
    procedure ParamGrille ;

  public
    ObjEcr 	  : TOB ;		 // TOB de l'écriture appelante
    Arguments 	  : ARG_ANA ;		 // Arguments d'appel de la fiche
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  CPVENTIL_TOF, //SG6 16/12/2004 eSaisieAnalCroise
  SousSect;

{$R *.DFM}

Procedure eSaisieAnal( var ObjEcr : TOB ; var Arguments : ARG_ANA ) ;
Var X : TFeSaisAnal ;
    i : integer ;
BEGIN
  //SG6
  if  VH^.AnaCroisaxe then
    begin
    eSaisieAnalCroise(ObjEcr,Arguments);
    end
  else
    begin
    X:=TFeSaisAnal.Create(Application) ;
    Try
      X.ObjEcr    := ObjEcr ;
      X.Arguments := Arguments ;
      for i:=1 to MaxAxe do
        X.LL[i]:=NIL ;
      if Arguments.QuelEcr=EcrGen then
        //  si TInfo renseigné alors on l'utilise
        if Arguments.Info <> nil then
          begin
          if Arguments.Info.LoadCompte( ObjEcr.GetString('E_GENERAL') ) then
            for i:=1 to MaxAxe do
              X.Axes[i] := Arguments.Info.GetString('G_VENTILABLE' + intToStr(i) ) = 'X' ;
          end
        //  sinon TGGeneral...
        else
          for i:=1 to MaxAxe do
            X.Axes[i]:=Arguments.CC.Ventilable[i] ;
       X.Showmodal ;
    Finally
      Arguments := X.Arguments ;
      X.Free ;
     End ;

    Screen.Cursor:=SyncrDefault ;
    end;
END ;

Procedure eSaisieAnalEff( ObjEcr : TOB  ; LL : T5LL ; Arguments : ARG_ANA) ;
Var X : TFeSaisAnal ;
    i : integer ;
BEGIN
X:=TFeSaisAnal.Create(Application) ;
 Try
  X.ObjEcr:=ObjEcr ;
 	X.Arguments := Arguments ;
  for i:=1 to MaxAxe do
    X.LL[i]:=LL[i] ;

  if Arguments.QuelEcr=EcrGen then
    //  si TInfo renseigné alors on l'utilise
    if Arguments.Info <> nil then
      begin
      if Arguments.Info.LoadCompte( ObjEcr.GetString('E_GENERAL') ) then
        for i:=1 to MaxAxe do
          X.Axes[i] := Arguments.Info.GetString('G_VENTILABLE' + intToStr(i) ) = 'X' ;
      end
    else
      for i:=1 to MaxAxe do
        X.Axes[i]:=Arguments.CC.Ventilable[i] ;

  X.Showmodal ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

{=================================== Init et defaut ====================================}
procedure TFeSaisAnal.LectureSeule ;
BEGIN
  E_MONTANTECR.Enabled:=False ;
  BInsert.enabled:=False ;
  BSDel.Enabled:=False ;
  BSolde.Enabled:=False ;
  BVentilType.Enabled:=False ;
  BSauveVentil.Enabled:=False ;
  CVType.Enabled:=False ;
  BCalculQte.Enabled:=False ;
END ;

// Affectation d'une TOB analytiq à la ligne
// intégré directement dans l'architecture de l'OBM écriture
procedure TFeSaisAnal.AlloueMvt ( Lig : integer ) ;
VAR OBA : TOB;
BEGIN
  if GSA.Objects[AN_NumL,Lig]<>Nil then exit ;
  if GSA.Cells[AN_NumL,Lig]='' then Exit ;
  if Arguments.QuelEcr=EcrGen then
    begin
    OBA := CGetNewTOBAna(CurAxe) ;
    InitCommunObjAnalNew(ObjEcr, OBA) ;
//    OBA.ChangeParent(ObjEcr.Detail[CurAxe-1],-1);
    GSA.Objects[AN_NumL,Lig]:=TObject(OBA) ;
    end;
END ;

procedure TFeSaisAnal.InitGrid ;
BEGIN
DefautLigne(GSA.RowCount-1) ;
END ;

procedure TFeSaisAnal.DefautEntete ;
BEGIN
  if Arguments.QuelEcr=EcrGen
    then G_GENERAL.Caption := ObjEcr.GetValue('E_GENERAL') ;
  E_REFINTERNE.Caption	:= ObjEcr.GetValue('E_REFINTERNE') ;
  E_LIBELLE.Caption	:= ObjEcr.GetValue('E_LIBELLE') ;
  E_MONTANTECR.Value    := GetMontantDev(ObjEcr) ;
  E_TOTAL.Value		:= 0 ;
  E_RESTE.Value		:= E_MONTANTECR.Value ;
  E_RESTEPOURC.Value	:= 1 ;
  E_TOTPOURC.Value	:= 0 ;
  E_MONTANTECR.DEBIT	:= (Sens=1) ;
  ChangeFormatDevise(Self,Arguments.DEV.Decimale,Arguments.DEV.Symbole) ;
END ;

{=============================== GRIDS DE SAISIE ==================================}

// Charge la grille pour l'axe NumAxe
procedure TFeSaisAnal.ChargeGrid ( NumAxe : integer ) ;
Var i,Lig  : integer ;
    OBA    : TOB ;
    LeFb   : TFichierBase ;
    CSect  : TGSection ;       // FQ 16037 SBO 23/09/2005
    lPourc : double ;
BEGIN
  // Init longueur saisie section
  ParamGrille ;
  LeFB:=TFichierBase(Ord(fbAxe1)+NumAxe-1) ;
  GSA.CacheEdit ;
  GSA.ColLengths[AN_Sect]:=VH^.Cpta[LeFb].Lg ;
  GSA.MontreEdit ;
  // Remplissage de la grille à partir de l'OBM
  for i:=0 to ObjEcr.Detail[numAxe-1].Detail.Count-1 do
    BEGIN
    Lig := i+1 ;
    OBA := ObjEcr.Detail[numAxe-1].Detail[i] ;
    CAddChampSupAna( OBA ) ; // ajoute les champs supp du TOBM SBO 02/11/2004 sinon violation d'accès au calcul des soldes de la section
    // Numéro Ligne + TOB
    GSA.Cells[AN_NumL,Lig]	:= IntToStr(Lig) ;
    GSA.Objects[AN_NumL,Lig]	:= TObject(OBA);
    // Section
    GSA.Cells[AN_Sect,Lig]	:= OBA.GetValue(Pf+'_SECTION') ;
    PositionneTz(Cache,NumAxe) ; // Param HComte encore valable ?
    CSect := TGSection.Create( GSA.Cells[AN_Sect,Lig], Cache.ZoomTable ) ;
    GSA.Cells[AN_Lib,Lig]    := CSect.Libelle ;
    GSA.Objects[AN_Sect,Lig] := TObject(CSect) ; // FQ 16037 SBO 23/09/2005
//    ChercheSect(Lig,False) ;     // Param HComte encore valable ?
    // Ajout Info Comp OBM
    AlloueOS(Lig) ;
    // Montant
    if Sens=1
      then GSA.Cells[AN_Montant,Lig]:=StrS(OBA.GetValue(Pf+'_DEBITDEV'),Arguments.DEV.Decimale)
      else GSA.Cells[AN_Montant,Lig]:=StrS(OBA.GetValue(Pf+'_CREDITDEV'),Arguments.DEV.Decimale) ;
    // Pourcentage
    if ( Arguments.action=taModif ) and Arguments.EnMontant then
      begin
      lPourc := 100.0*ValM(Lig)/E_MontantEcr.Value ;
      GSA.Cells[AN_Pourcent,Lig] := StrS(lPourc,ADecimP) ; // OBA.GetValue(Pf+'_POURCENTAGE')
      end
    else
      begin
      GSA.Cells[AN_Pourcent,Lig] := StrS(OBA.GetValue(Pf+'_POURCENTAGE'),ADecimP) ;
      end ;
    // formatage
//FormatMontant(AN_Montant,Lig) ;
    FormatMontant(AN_Pourcent,Lig) ;
    TraiteMontant(Lig,True) ;
    if ( FBoOuiTvaEnc ) and ( CurAxe = FInAxeTva ) then
     GSA.Cells[AN_Date,Lig] := DateToStr(OBA.GetValue(Pf+'_DATEREFEXTERNE')) ;
    // ??
    GereNewLigne(GSA) ;
    END ;

  // Calcul totaux
  CalculTotal(NumAxe) ;
  // Positionnement grille
  GSA.Row:=1 ;
  GSA.Col:=An_Sect ;
  // Test si axe ok
  Equilibre(NumAxe) ;
END ;

procedure TFeSaisAnal.GSADblClick(Sender: TObject);
BEGIN
ClickZoom ;
END ;

procedure TFeSaisAnal.DesalloueA( Lig : Longint ) ;
begin
if GSA.Objects[AN_Sect,Lig]=Nil then Exit ;
TGSection(GSA.Objects[AN_Sect,Lig]).Free ; GSA.Objects[AN_Sect,Lig]:=Nil ;
end ;

{=============================== APPELS DE COMPTE ==================================}
Function TFeSaisAnal.ChercheSect( L : integer ; Force : boolean ) : byte ;
Var St    : String ;
    CSect : TGSection ;
    C     : integer ;
    Changed,Idem : boolean ;
  {$IFNDEF CCS3}
    St1   : String ;
    fb : TFichierBase ;
  {$ENDIF}
    OkSous : boolean ;
  {b FP 29/12/2005}
  CompteAna:   array[1..MaxAxe] of String;
  {e FP 29/12/2005}
BEGIN
  Result:=0 ; C:=AN_Sect ; Changed:=False ;
  St:=uppercase(GSA.Cells[C,L]) ; Cache.Text:=St ;

  if PasToucheLigne( GSA.Row ) then Exit ; // FQ 16037

  {$IFDEF SPEC302}
  OkSous:=VH^.GereSousPlan ;
  {$ELSE}
    {$IFDEF CCS3}
    OkSous:=False ;
    {$ELSE}
    OkSous:=((VH^.GereSousPlan) or (VH^.SaisieTranche[CurAxe])) ;
    {$ENDIF}
  {$ENDIF}

  {$IFNDEF CCS3}
If ((OkSous) and (St='') and (Arguments.Action in [taCreat,taModif])) Then
   BEGIN
   fb:=AxeTofb('A'+IntToStr(CurAxe)) ;
   If VH^.Cpta[fb].Structure Then
      BEGIN
      St1:=ChoisirSSection(fb,St,FALSE,Arguments.Action) ;
      If (St1<>'') And (Arguments.Action in [taCreat,taModif]) Then
        BEGIN
        St:=St1 ;
        Cache.Text:=St1 ;
        END ;
      END ;
   END ;
  {$ENDIF}
  
{b FP 29/12/2005: Ajoute le filtre pour les restrictions analytiques}
  FillChar(CompteAna, sizeof(CompteAna), #0);
  Cache.SynPlus := FRestriction.GetClauseCompteAutorise(
      G_GENERAL.Caption, 'A'+IntToStr(CurAxe), 'SECTION', CompteAna);
{e FP 29/12/2005}
if GChercheCompte(Cache,nil) then
   BEGIN
   if GSA.Objects[C,L]<>Nil then if TGSection(GSA.Objects[C,L]).Sect<>St then Changed:=True ;
   Idem:=((St=Cache.Text) and (GSA.Objects[C,L]<>Nil) and (Not Changed)) ;
   if ((Not Idem) or (Force)) then
      BEGIN
      GSA.Cells[C,L]:=Cache.Text ; DesalloueA(L) ;
      CSect:=TGSection.Create(Cache.Text,Cache.ZoomTable) ; GSA.Objects[C,L]:=TObject(CSect) ;
      if CSect.Ferme then
         BEGIN
         // Impossible de saisir sur des sections fermées FQ 16037 SBO 22/09/2005
         If Arguments.action=taCreat Then
            BEGIN
            ErreurSaisie(6) ;
            DesalloueA(L) ;
            GSA.Cells[C,L]:='' ;
            Exit ;
            END
          else
            BEGIN
            ErreurSaisie(5) ;
            END ;
         END ;
      GSA.Cells[AN_Lib,L]:=CSect.Libelle ;
      Result:=1 ;
      AlimObjetMvt(L) ;
      END else
      BEGIN
      Result:=2 ;
      END ;
   END ;
END ;

{================================== Calculs lignes ==================================}
function TFeSaisAnal.ValM ( Lig : integer ) : Double ;
BEGIN Result:=Valeur(GSA.Cells[AN_Montant,Lig]) ; END ;

function TFeSaisAnal.ValP ( Lig : integer ) : Double ;
BEGIN Result:=Valeur(GSA.Cells[AN_Pourcent,Lig]) ; END ;

procedure TFeSaisAnal.TraiteMontant ( Lig : integer ; Calc : boolean ) ;
Var X,XD,XC : Double ;
    OBA     : TOB ;
BEGIN
  OBA := GetTOB(GSA,Lig) ;
  if OBA = Nil then Exit ;
  X := ValP(Lig) ;
  OBA.PutValue(Pf+'_POURCENTAGE',X) ;
  X := ValM(Lig) ;
  if Sens=1 then
    BEGIN
    XD:=X ;
    XC:=0 ;
    END
  else
    BEGIN
    XD:=0 ;
    XC:=X ;
    END ;
  CSetMontants(OBA,XD,XC,Arguments.DEV,False) ;
  if Calc then CalculTotal(CurAxe) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Calcul des cumuls pour l'axe NumA
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.CalculTotal ( NumA : integer ) ;
Var i   : integer ;
    OBA : TOB ;
    CSect : TGSection ;
BEGIN
  // Init
  Conf[NumA]:='0' ;
  SI_TotD[NumA]:=0 ;
  SI_TotP[NumA]:=0 ;
  SI_TotPourc[NumA]:=0 ;
  SI_TotSais[NumA]:=0 ;
  // Parcours des lignes pour cumuls
  for i:=1 to GSA.RowCount-1 do
    if Trim(GSA.Cells[AN_Sect,i])<>'' then
      BEGIN
      SI_TotSais[NumA] 	:= SI_TotSais[NumA]  + ValM(i) ;
      SI_TotPourc[NumA] := SI_TotPourc[NumA] + ValP(i) ;
      OBA:=GetTOB(GSA,i) ;
      if OBA<>Nil then
        BEGIN
        SI_TotP[NumA] := SI_TotP[NumA] + OBA.GetDouble(Pf+'_CREDIT') 	+ OBA.GetDouble(Pf+'_DEBIT') ;
        SI_TotD[NumA] := SI_TotD[NumA] + OBA.GetDouble(Pf+'_CREDITDEV') + OBA.GetDouble(Pf+'_DEBITDEV') ;
        END ;
      if Arguments.Action<>taConsult then
        BEGIN
        CSect:=GetGSect(GSA,i) ;
        if ((CSect<>Nil) and (CSect.Confidentiel>Conf[NumA]))	then Conf[NumA]:=CSect.Confidentiel ;
        END ;
      END ;
  // Affichage
  SI_TotPourc[NumA] 	:= SI_TotPourc[NumA]/100.0 ;
  E_TotPourc.Value 	  := SI_TotPourc[NumA] ;
  E_RestePourc.Value	:= 1.0-SI_TotPourc[NumA] ;
  E_Total.Value		    := SI_TotSais[NumA] ;
  E_Reste.Value		    := E_MontantEcr.Value-SI_TotSais[NumA] ;
  BValide.Enabled	    := Equilibre(NumA) ;
  BCalculQte.Enabled  := Equilibre(NumA) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Affecter un numéro et une TOB vide à la ligne Lig
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.DefautLigne ( Lig : integer ) ;
BEGIN
  GSA.Cells[AN_NumL,Lig]:=IntToStr(Lig) ;
  AlloueMvt(Lig) ;
  AlloueOS( Lig ) ;
  if ( FBoOuiTvaEnc ) and ( CurAxe = FInAxeTva ) then
   GSA.Cells[AN_Date,Lig] := DatetoStr(ObjEcr.GetValue('E_DATECOMPTABLE')) ;
END ;

procedure TFeSaisAnal.CalculSoldeCompte ( Sect : String ; DIS,CIS : Double ) ;
Var i : integer ;
    TDS,TCS : Double ;
BEGIN
  TDS:=0 ;
  TCS:=0 ;
  for i:=1 to GSA.RowCount-1 do
    BEGIN
    if ((Sect<>'') and (GSA.Cells[AN_Sect,i]=Sect)) then
      BEGIN
      if Sens=1
        then TDS := TDS + ValeurPivot(GSA.Cells[AN_Montant,i],Arguments.DEV.Taux,Arguments.DEV.Quotite)-GetO(GSA,i).GetValue('OLDDEBIT')
        else TCS := TCS + ValeurPivot(GSA.Cells[AN_Montant,i],Arguments.DEV.Taux,Arguments.DEV.Quotite)-GetO(GSA,i).GetValue('OLDCREDIT') ;
      END ;
    END ;
  AfficheLeSolde(S_Solde,TDS+DIS,TCS+CIS) ;
END ;

{================================== Affichages ==================================}
procedure TFeSaisAnal.AffichePied  ;
Var CSect   : TGSection ;
    Sect     : String ;
    DIS,CIS :  double ;
BEGIN
  Sect:='' ;
  DIS:=0 ;
  CIS:=0 ;
  if GSA.Objects[AN_Sect,GSA.Row]<>Nil then
    BEGIN
    CSect	:= TGSection(GSA.Objects[AN_Sect,GSA.Row]) ;
    Sect	:= CSect.Sect ;
    DIS		:= CSect.TotalDebit ;
    CIS		:= CSect.TotalCredit ;
    END ;
  CalculSoldeCompte(Sect,DIS,CIS) ;
END ;

procedure TFeSaisAnal.InsereLigne ( Lig : integer ) ;
BEGIN
  if Arguments.Action=taConsult then Exit ;
  // Insertion et init d'une nouvelle ligne dans la grille
  if Lig<GSA.RowCount-1 then
    BEGIN
    GSA.InsertRow(Lig) ;
    DefautLigne(Lig) ;
    END ;
  NumeroteLignes(GSA) ;
  // MAJ Affichage Cumul
  AffichePied ;
END ;

procedure TFeSaisAnal.FormatMontant ( ACol,ARow : integer ) ;
Var X : Double ;
BEGIN
  X := Valeur(GSA.Cells[ACol,ARow]) ;
  if X=0 then GSA.Cells[ACol,ARow]:='' else
    BEGIN
    if ACol=AN_Montant
      then GSA.Cells[ACol,ARow] := StrS(X,Arguments.DEV.Decimale)
      else GSA.Cells[ACol,ARow] := StrS(X,ADecimP) ;
    END ;
END ;

{========================================= Controles=====================================}
procedure TFeSaisAnal.ControleLignes ;
Var Lig : integer ;
BEGIN
  for Lig:=1 to GSA.RowCount-1 do
    if Not EstRempli(GSA,Lig)
      then DefautLigne(Lig) ;
END ;


procedure TFeSaisAnal.ErreurSaisie ( Err : integer ) ;
BEGIN
	if Err<100 then HMessLigne.Execute(Err-1,'','')
  	else if Err<200 then HMessPiece.Execute(Err-101,'','') ;
END ;

Function TFeSaisAnal.Equilibre (NumA : integer) : Boolean ;
Var Diff,D,C : Double ;
    OBA : TOB ;
BEGIN
  Result := False ;
  OBA := GetTOB(GSA,1) ;
  if OBA=Nil then Exit ;
  Result := Arrondi(SI_TotSais[NumA]-E_MontantEcr.Value,Arguments.DEV.Decimale)=0 ;
  if Not Result then Exit ;
{Si équilibre, forcer les autres équilibres}
// Pivot
  Diff := Arrondi( SI_TotP[NumA]-MontantEcrP , V_PGI.OkDecV ) ;
  if Diff<>0 then
    BEGIN
    D := OBA.GetValue(Pf+'_DEBIT') ;
    C := OBA.GetValue(Pf+'_CREDIT') ;
    if D<>0
      then OBA.PutValue(Pf+'_DEBIT',D-Diff)
      else OBA.PutValue(Pf+'_CREDIT',C-Diff) ;
    SI_TotP[NumA] := MontantEcrP ;
    END ;
// Devise
  Diff:=Arrondi(SI_TotD[NumA]-MontantEcrD,Arguments.DEV.Decimale) ;
  if Diff<>0 then
    BEGIN
    D:=OBA.GetValue(Pf+'_DEBITDEV') ;
    C:=OBA.GetValue(Pf+'_CREDITDEV') ;
    if D<>0
      then OBA.PutValue(Pf+'_DEBITDEV',D-Diff)
      else OBA.PutValue(Pf+'_CREDITDEV',C-Diff) ;
    SI_TotD[NumA]:=MontantEcrD ;
    END ;
END ;


function TFeSaisAnal.LigneCorrecte( Lig : integer ; Alim : boolean = True ): Boolean;
Var lErr : TRecError ;
    lTob : Tob ;
begin

  if Arguments.Action=taConsult then
    begin
    result := true ;
    Exit ;
    end ;

  if (length(trim(GSA.Cells[AN_Date,Lig])) = 10 ) then
   begin
    if not CControleDateBor(strToDate(GSA.Cells[AN_Date,Lig]), ctxExercice ) then
     GSA.Cells[AN_Date,lig] := DatetoStr(ObjEcr.GetValue('E_DATECOMPTABLE')) ;
   end
    else
     GSA.Cells[AN_Date,lig] := DatetoStr(ObjEcr.GetValue('E_DATECOMPTABLE')) ;

  // alimente TOB de la grille avant test
  if Alim then
    AlimObjetMvt(Lig) ;

  // Verif Section + Verif Montant => Verif ligne
  lTob := GetTOB( GSA, Lig ) ;
  lErr := CIsValidLigneAna( lTob, FInfo ) ;

  // Gestion spéciale des section fermée :
  if (lErr.RC_Error = RC_YSECTIONFERMEE) and (Arguments.Action<>taCreat) then
    begin
    OnError( nil, lErr ) ;
    result := True ;
    end
  else
    begin
    result  := lErr.RC_Error = RC_PASERREUR ;
    GSA.Col := AN_Sect ;
    end ;

  // Affichage erreur
  if Not Result
    then OnError( nil, lErr ) ;

end;

function TFeSaisAnal.AxeCorrect( NumA: integer; Parole: boolean ) : Boolean ;
Var lErr : TRecError ;
    lTob : Tob ;
    lTobAna : Tob ;
    i : integer ;
begin

  if Arguments.Action=taConsult then
    begin
    result := true ;
    Exit ;
    end ;

  lErr.RC_Error := RC_PASERREUR ;
  lErr.RC_Axe   := NumA ;
  
  // Ventilation équilibrée ?
  if Not Equilibre(NumA) then
    lErr.RC_Error := RC_YSOLDEPOURC ;

  // Axe correct ?
  lTob   := ObjEcr.Detail[ NumA - 1 ] ; // Tob de l'axe !!

  // Test affiné à la ligne pour détecter la ligne qui pose problème
  // Test des lignes
  if (lErr.RC_Error = RC_PASERREUR) then
    for i := 0 to lTob.Detail.count - 1 do
      begin
      lTobAna := lTob.Detail[i] ;
      // Test ligne
      lErr.RC_Error := CIsValidLigneAna( lTobAna, FInfo ).RC_Error ;
      if lErr.RC_Error <> RC_PASERREUR then
        begin
        GSA.Row := i + 1 ;
        Break ;
        end ;
      end ;

  // Ancien test global sur axe...en attente validation modif ci dessus
//  if (lErr.RC_Error = RC_PASERREUR) then
//    lErr   := CIsValidAxe( NumA, lTob, FInfo, Arguments.VerifQte, False ) ;

  result := (lErr.RC_Error = RC_PASERREUR) ;

  // Verif bufget non bloquant
  if result and Arguments.ControleBudget and Parole then
    if CIsValidAxeBudget( NumA, lTob, FInfo ) <> RC_PASERREUR then
      result := HMessPiece.Execute(6,'','') = mrYes ;

  // Message si demandé
  if (not Result) and Parole and ( lErr.RC_Error <> RC_YSECTIONBUDGET ) then
    OnError( nil, lErr ) ;

end;

function TFeSaisAnal.ValideLeAxe( NumA : integer ; Parole : boolean ) : Boolean;
begin
  // Test validité des données et équilibrage
  Result := AxeCorrect( NumA, Parole ) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Alimente la tob avec les données de la grille
Suite ........ : Création si besoin
Mots clefs ... :
*****************************************************************}
Procedure TFeSaisAnal.AlimObjetMvt ( Lig : integer ) ;
Var OBA : TOB ;
BEGIN
  OBA := GetTOB(GSA,Lig) ;
  if OBA=Nil then
    BEGIN
    DefautLigne(Lig) ;
    OBA := GetTOB(GSA,Lig) ;
    END ;
// ?? MONTANT ??
  if OBA.Parent = nil then
    OBA.ChangeParent( ObjEcr.Detail[CurAxe-1] , -1 );
  OBA.PutValue(Pf+'_NUMVENTIL',	Lig ) ;
  OBA.PutValue(Pf+'_AXE',	'A'+IntToStr(CurAxe) ) ;
  OBA.PutValue(Pf+'_SECTION',	GSA.Cells[AN_Sect,Lig] ) ;
  if ( FBoOuiTvaEnc ) and ( CurAxe = FInAxeTva ) then
   OBA.PutValue(Pf+'_DATEREFEXTERNE' , StrToDate(GSA.Cells[AN_Date,Lig]) ) ;
  if Arguments.Action = taCreat then OBA.PutValue ( Pf + '_DATECREATION', Date ) ;
END ;

{================================= Barre d'outils ==================================}
procedure TFeSaisAnal.ClickAbandon ;
BEGIN
  IsCanceled := True;
  Close ;
END ;

procedure TFeSaisAnal.ClickValide ;
Var lBoCancel : boolean ;
    lBoChange : boolean ;
    i         : integer ;
    lErr      : TRecError ;
begin

  if Not GSA.SynEnabled then Exit ;

  if Arguments.Action=taConsult then
    begin
    Close ;
    Exit ;
    end ;

  // Achèvement de la saisie
  lBoCancel := False ;
  lBoChange := True ;
  GSARowExit(Nil,GSA.Row,lBoCancel,lBoChange) ;
  CalculTotal(CurAxe) ;

  // L'axe courant doit être valide
  if lBoCancel then Exit ;
  if Not ValideLeAxe(CurAxe,True) then Exit ;

  // Recherche du premier axe non valide
  lErr := CIsValidVentil( ObjEcr, FInfo ) ;
  if lErr.RC_Error = RC_PASERREUR then
    begin
    // Si Tous les axes sont valides, on sort
    if Arguments.Action<>taConsult then
      begin
      Arguments.ModeConf:='0' ;
      for i:=1 to MaxAxe do
        if Conf[i]>Arguments.ModeConf then Arguments.ModeConf:=Conf[i] ;
      end ;

    //SG6 21/12/2004 FQ 14731 Vérification des quantités
    if Arguments.VerifQte then // DEV3946
      CheckQuantite(ObjEcr);

    IsCanceled := False;
    Close ;

    end
  else
    begin
    CurAxe      := lErr.RC_Axe ;
    TA.TabIndex := CurAxe - 1 ;
    TAChange(Nil) ;
    end ;

  Cursor:=crDefault ;
end;

procedure TFeSaisAnal.ClickSolde ;
Var C,R,Lig,Acol,ARow : Longint ;
    b,Cancel : boolean ;
    Diff,X   : double ;
begin
if Arguments.Action=taConsult then Exit ; if E_MontantEcr.Value=0 then Exit ;
if GSA.Cells[AN_Sect,GSA.Row]='' then Exit ;
// ajout me 18/07/2007 fiche 19750 restrcition ana n'est pas prise en compte avec F6
if (ChercheSect(GSA.Row,False) <= 0) then Exit ; 

if Not ExJaiLeDroitConcept(TConcept(ccSaisSolde),True) then Exit ;
if PasToucheLigne( GSA.Row ) then Exit ; // FQ 16037
if GSA.RowCount<=2 then BEGIN Acol:=GSA.Col ; ARow:=GSA.Row ; Cancel:=False ; GSACellEnter(Nil,ACol,ARow,Cancel) ; END ;
C:=GSA.Col ; R:=GSA.Row ; b:=False ; Lig:=GSA.Row ;

if ((GSA.Col=AN_Pourcent) or (GSA.Col=AN_Montant)) then GSACellExit(Nil,C,R,b) ;

Diff:=Arrondi(E_MontantEcr.Value-SI_TotSais[CurAxe],Arguments.DEV.Decimale) ;
if Diff=0 then Exit ;
GSA.Cells[AN_Montant,Lig] := StrS( ValM(Lig) + Diff, Arguments.DEV.Decimale ) ;

X:=100.0*ValM(Lig)/E_MontantEcr.Value ;
GSA.Cells[AN_Pourcent,Lig]:=StrS(X,ADecimP) ;

FormatMontant(AN_Montant,Lig) ;
FormatMontant(AN_Pourcent,Lig) ;
TraiteMontant(Lig,True) ;

// gestion de l'arrondi pour solde
GereArrondi( Lig, AN_Montant ) ;

if Equilibre(CurAxe) then ;

end;

procedure TFeSaisAnal.ClickDel ;
Var
  R : integer;
  O : TOB;
BEGIN
  if GSA.RowCount<=3 then Exit ;
  if Arguments.Action=taConsult then Exit ;
  if ((GSA.Row=GSA.RowCount-1) and (Not EstRempli(GSA,GSA.Row))) then BEGIN GSA.SetFocus ; Exit ; END ;
  if PasToucheLigne( GSA.Row ) then Exit ; // FQ 16037

  GSA.CacheEdit ;
  R:=GSA.Row ;

  // BPY le 19/10/2004 => delete de la ligne dans la TOB !
  {JP 13/01/06 : FQ 17160 : en fait, l'ObjEcr est mis à jour dans le FormShow, lors du changement
               d'onglet et de la validation : ce n'est donc pas ObjEcr qu'il faut vider mais la tob
               associée à la ligne de la grille.
  ObjEcr.Detail[TA.TabIndex].Detail.Delete(GSA.Row-1);}
  O := GetTOB(GSA, R);
  if Assigned(O) then FreeAndNil(O);
  DeAlloueObj( GSA.Row ) ;
  GSA.DeleteRow(GSA.Row) ;

  CalculTotal(CurAxe) ;
  NumeroteLignes(GSA) ;
  GSA.SetFocus ;
  if R>1 then GSA.Row:=R
  else GSA.Row:=1 ;
  GSA.Col:=AN_Sect ;
  GereNewLigne(GSA) ;
  GSA.Invalidate ;
  GSA.MontreEdit ;
  ControleLignes ;

  GereOptionsGrid( GSA.Row ) ;

END ;

procedure TFeSaisAnal.ClickInsert ;
Var R : integer ;
BEGIN
if Not EstRempli(GSA,GSA.Row) then Exit ; if Arguments.Action=taConsult then Exit ;
if Not LigneCorrecte(GSA.Row,True) then Exit ;
GSA.CacheEdit ;
R:=GSA.Row ; InsereLigne(R) ;
GSA.SetFocus ; GSA.Row:=R ; GSA.Col:=AN_Sect ; GSA.Cells[GSA.Col,GSA.Row]:='' ;
GereNewLigne(GSA) ;
//GSA.MontreEdit ;
GereOptionsGrid( GSA.Row ) ;
END ;

procedure TFeSaisAnal.ZoomVentilType ;
{$IFNDEF CCS3}
Var OkPourCalcul : Boolean ;
    vv : Variant ;
{$ENDIF}
BEGIN
if not vh^.cpifdefcegid then exit ;
{$IFNDEF CCS3}
OkPourCalcul:=TRUE ;
if Arguments.Action=taConsult then OkPourCalcul:=FALSE ;
if not CVType.Enabled then OkPourCalcul:=FALSE ;
vv:=CPLanceFicheAFFVentilana;
If OkpourCalcul Then
  BEGIN
  CVType.Value:=vv ;
  END ;
{$ENDIF}
END ;


procedure TFeSaisAnal.ClickVentilType ;
BEGIN
if Arguments.Action=taConsult then Exit ; if not CVType.Enabled then Exit ;
if Arguments.DernVentilType='' then Exit ;
CVType.Value:=Arguments.DernVentilType ;
END ;

procedure TFeSaisAnal.ClickComplement ( Lig : integer );
{$IFNDEF GCGC}
Var ModBN : Boolean ;
    LeEcr : TTypeEcr ;
    RC      : R_COMP ;
    lAction : TActionFiche ;
{$ENDIF}
BEGIN
{$IFNDEF GCGC}
if Not EstRempli(GSA,Lig) then Exit ;
if Not LigneCorrecte(Lig,True) then Exit ;
LeEcr:=EcrAna ;
RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ; RC.StComporte:='--XXXXXXXX' ;
RC.Conso:=True ; RC.Attributs:=False ; RC.MemoComp:=Nil ; RC.Origine:=-1 ; RC.DateC:=0 ;
RC.TOBCompl := nil ;
lAction := Arguments.Action ;
if PasToucheLigne( Lig ) then lAction := taConsult ; // FQ 16037
SaisieComplement(GetO(GSA,Lig),LeEcr,lAction,ModBN,RC) ;
// Mémorisation test // DEV3946
if ObjEcr.GetNumChamp('CHECKQTE') < 0
  then ObjEcr.AddChampSupValeur('CHECKQTE', '-' )
  else ObjEcr.PutValue('CHECKQTE', '-' ) ;
{$ENDIF}
END ;

procedure TFeSaisAnal.ClickZoom ;
Var b   : byte ;
    C,R : Longint ;
    A   : TActionFiche ;
begin
if Arguments.Action=taConsult then
   BEGIN
   if ((GX=0) and (GY=0)) then BEGIN C:=AN_Sect ; R:=1 ; END else GSA.MouseToCell(GX,GY,C,R) ;
   A:=taConsult ;
   END else
   BEGIN
   R:=GSA.Row ; C:=GSA.Col ; A:=taModif ;
   END ;
if R<1 then exit ; if C<AN_Sect then exit ;
if ((Arguments.Action=taConsult) and (GSA.Cells[C,R]='')) then Exit ;
if Not ExJaiLeDroitConcept(TConcept(ccSecModif),False) then A:=taConsult ;
if C=AN_Sect then
   BEGIN
   b:=ChercheSect(R,False) ;
   if b=2 then
     begin
     FicheSection(Nil,'A'+Chr(48+CurAxe),GSA.Cells[GSA.Col,GSA.Row],A,0) ;
     if Arguments.Action<>taConsult
       then ChercheSect(R,True) ;
     end ;
   GereNewLigne(GSA) ;
   END ;
end;

{$IFNDEF CCS3}
procedure TFeSaisAnal.ClickZoomSousPlan ;
Var C,R : Longint ;
    A   : TActionFiche ;
    fb  : TFichierBase ;
    Sect : String ;
begin
if Arguments.Action=taConsult then
   BEGIN
   if ((GX=0) and (GY=0)) then BEGIN C:=AN_Sect ; R:=1 ; END else GSA.MouseToCell(GX,GY,C,R) ;
   A:=taConsult ;
   END else
   BEGIN
   R:=GSA.Row ; C:=GSA.Col ; A:=taModif ;
   END ;
if R<1 then exit ; if C<AN_Sect then exit ;
if ((Arguments.Action=taConsult) and (GSA.Cells[C,R]='')) then Exit ;
if Not ExJaiLeDroitConcept(TConcept(ccSecModif),False) then A:=taConsult ;
if C=AN_Sect then
   BEGIN
   fb:=AxeTofb('A'+IntToStr(CurAxe)) ;
   If VH^.Cpta[fb].Structure Then
      BEGIN
      Sect:=ChoisirSSection(fb,GSA.Cells[AN_SECT,R],FALSE,A) ;
      If (Sect<>'') And (A in [taCreat,taModif]) Then GSA.Cells[AN_SECT,R]:=Sect ;
      END ;
   ChercheSect(R,TRUE) ;
   GereNewLigne(GSA) ;
   END ;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Affiche le panel d'enregistrement de la ventilation type
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.ClickSauveVentil ;
BEGIN
  // Tests préliminaires
  if Arguments.Action=taConsult then Exit ;
  if GSA.RowCount<=2 then Exit ;
  if E_TOTPOURC.Value=0 then Exit ;
  if Not ExJaiLeDroitConcept(TConcept(ccSaisCreatVentil),True) then Exit ;
  if ((EstRempli(GSA,GSA.Row)) and (Not LigneCorrecte(GSA.Row,True))) then Exit ;
  // Affichage du panel d'enregistrement de la ventilatin type
  PVentil.Left:=GSA.Left+(GSA.Width-PVentil.Width) div 2 ;
  PVentil.Visible:=True ;
  Y_CODEVENTIL.SetFocus ;
  ModeVentil:=True ;
  PEntete.Enabled:=False ;
  Outils.Enabled:=False ;
  Valide97.Enabled:=False ;
  PGA.Enabled:=False ;
  PPied.Enabled:=False ;
  GSA.Enabled:=False ;
END ;

procedure TFeSaisAnal.BSDelClick(Sender: TObject);
begin ClickDel ; end;

procedure TFeSaisAnal.BAbandonClick(Sender: TObject);
begin
  ClickAbandon ;
end;

procedure TFeSaisAnal.BValideClick(Sender: TObject);
BEGIN ClickValide ; END ;

procedure TFeSaisAnal.BSoldeClick(Sender: TObject);
BEGIN ClickSolde ; END ;

procedure TFeSaisAnal.BInsertClick(Sender: TObject);
begin ClickInsert ; end;

procedure TFeSaisAnal.BComplementClick(Sender: TObject);
begin ClickComplement(GSA.Row) ; end;

procedure TFeSaisAnal.BZoomClick(Sender: TObject);
begin ClickZoom ; end;

procedure TFeSaisAnal.BVentilTypeClick(Sender: TObject);
begin ClickVentilType ; end;

procedure TFeSaisAnal.BSauveVentilClick(Sender: TObject);
begin ClickSauveVentil ; end;

{================================= Création ventil type ================================}
procedure TFeSaisAnal.FinVentil ;
BEGIN
PEntete.Enabled:=True ; Outils.Enabled:=True ; Valide97.Enabled:=True ; PGA.Enabled:=True ;
PPied.Enabled:=True ; GSA.Enabled:=True ;
CVType.SetFocus ; PVentil.Visible:=False ; ModeVentil:=False ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Création d'une nouvelle ventilation type :
Suite ........ :  -> MAJ tablette CHOIXCODE
Suite ........ :  -> MAJ table VENTIL
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.BNewVentilClick(Sender: TObject);
Var Existe    	: boolean ;
    i, n      	: integer ;
    OBA         : TOB ;						// Tob de la ligne de ventil
    TOBCC,												// TOB pour tablette
    TOBVentil		: TOB;						// TOB pour Ventil type
begin
  // Vérif saisie code + libellé
  if ((Y_CODEVENTIL.Text='') or (Y_LIBELLEVENTIL.Text='')) then
    BEGIN
    HMessPiece.Execute(4,'','') ;
    Exit ;
    END ;
  // Vérif existence
  Existe:=ExisteSQL( 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="VTY" AND '
  		       	+ '(CC_CODE="' + Y_CODEVENTIL.Text + '" OR '
                        +  'UPPER(CC_LIBELLE)="' + uppercase(Y_LIBELLEVENTIL.Text) + '")' ) ;
  if Existe then
    BEGIN
    HMessPiece.Execute(5,'','') ;
    Exit ;
    END ;
  // Mise à jour de la tablette
  //SG6 24/01/05 FQ 15265
  TOBCC := TOB.Create('CHOIXCOD',nil,-1);
  TOBCC.InitValeurs(false);
  TOBCC.PutValue('CC_TYPE',	'VTY');
  TOBCC.PutValue('CC_CODE',	Y_CODEVENTIL.Text);
  TOBCC.PutValue('CC_LIBELLE',	Y_LIBELLEVENTIL.Text);
  TOBCC.InsertDB(nil);
  TOBCC.Free ;

  {JP 14/08/07 : FQ 20881 : En travaillant sur la grille on se condamne à n'enregistrer que la ventilation
                 de l'onglet courant => il faut travailler sur la TOB
  // Mise à jour de la ventilation type
  for i:=1 to GSA.RowCount-2 do
    if GSA.Cells[AN_Sect,i]<>'' then
      Begin
      OBA := GetTOB(GSA,i) ;
      TOBVentil := TOB.Create('VENTIL', nil, -1);
      TOBVentil.InitValeurs(false);
      TOBVentil.PutValue('V_NATURE',	    'TY'+IntToStr(CurAxe));
      TOBVentil.PutValue('V_COMPTE',	    Y_CODEVENTIL.text);
      TOBVentil.PutValue('V_SECTION',	    OBA.GetValue(Pf+'_SECTION'));
      TOBVentil.PutValue('V_TAUXMONTANT',   OBA.GetValue(Pf+'_POURCENTAGE'));
      TOBVentil.PutValue('V_TAUXQTE1',	    OBA.GetValue(Pf+'_POURCENTQTE1'));
      TOBVentil.PutValue('V_TAUXQTE2',	    OBA.GetValue(Pf+'_POURCENTQTE2'));
      TOBVentil.PutValue('V_NUMEROVENTIL',  i);
      TOBVentil.InsertDB(nil);
      TOBVentil.Free ;
      end; }

  for n := 1 to MaxAxe do
    if Axes[n] then begin
      for i := 0 to ObjEcr.Detail[n - 1].Detail.Count - 1 do begin
        OBA := ObjEcr.Detail[n - 1].Detail[i] ;
        TOBVentil := TOB.Create('VENTIL', nil, -1);
        TOBVentil.InitValeurs(false);
        TOBVentil.PutValue('V_NATURE',	       'TY'+IntToStr(n));
        TOBVentil.PutValue('V_COMPTE',	       Y_CODEVENTIL.text);
        TOBVentil.PutValue('V_SECTION',	      OBA.GetValue(Pf+'_SECTION'));
        TOBVentil.PutValue('V_TAUXMONTANT',   OBA.GetValue(Pf+'_POURCENTAGE'));
        TOBVentil.PutValue('V_TAUXQTE1',	     OBA.GetValue(Pf+'_POURCENTQTE1'));
        TOBVentil.PutValue('V_TAUXQTE2',	     OBA.GetValue(Pf+'_POURCENTQTE2'));
        TOBVentil.PutValue('V_NUMEROVENTIL',  i + 1);
        TOBVentil.InsertDB(nil);
        TOBVentil.Free ;
      end; {for i := 0}
    end; {if Axes[n]}

  // Mise à jour interface
  FinVentil ;
  AvertirTable('ttVentilType') ;
  CVType.DataType:='' ;
  CVType.DataType:='ttVentilType' ;
end;

{================================= Méthodes de la form ==================================}
{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... : 25/10/2002
Description .. : Affichage de la forme :
Suite ........ :  -> Initialisation interface,
Suite ........ :  -> Gestion Guide, Ventilation type
Suite ........ :  -> Chargement Grille
Suite ........ :  -> Validation auto si scénario
Suite ........ :  -> Affichage axe courant
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.FormShow(Sender: TObject);
Var
 Prem          : integer ;
 Cpte          : String ;
 lErr          : TRecError ;
 lQ            : TQuery ;
 lIndex        : integer ;
 lBoUnSeulAxe  : boolean ;
 lBoOuvreAna   : boolean ;
 lBoVentilTva  : boolean ;
begin

  // Recup info et messages passés en parmaètres, création sinon.
  if Assigned(Arguments.Info)
    then FInfo := Arguments.Info
    else FInfo := TInfoEcriture.Create ;
  if Assigned(Arguments.MessageCompta)
    then FMessageCompta := Arguments.MessageCompta
    else FMessageCompta := TMessageCompta.Create( TraduireMemoire('Ventilations analytiques') ) ;

  // init interface
  if Arguments.QuelEcr=EcrGen then
    BEGIN
    Pf           := 'Y' ;
    GSA.TypeSais := tsAnal ;
    Cpte         := ObjEcr.GetString('E_GENERAL') ;
    MontantEcrP  := GetMontant(ObjEcr);
    MontantEcrD  := GetMontantDev(ObjEcr);
    Sens         := GetSens(ObjEcr) ;

    // FQ 18147 Remise en place de l'annulation
    if (Arguments.Action <> taConsult) and ( ObjEcr.GetNumChamp('NEWVENTIL') > 0 ) then
      if ( ObjEcr.GetString('NEWVENTIL') = 'X' )
        then Arguments.Action := taCreat
        else Arguments.Action := taModif ;

    END ;

  {JP 13/01/05 : FQ 17160 : On duplique la tob pour le cas d'une annulation de saisie}
  if (Arguments.Action = taModif) and ( CIsValidVentil( ObjEcr, FInfo ).RC_Error = RC_PASERREUR) then
    begin
    ObjTemp := TOB.Create('ECRITURE', nil, -1);
    ObjTemp.Dupliquer(ObjEcr, True, True);
    end ;
  IsCanceled := True;

  DefautEntete ;
  ModeVentil           := False ;
  GSA.DefaultRowHeight := 18 ;
  //GSA.ListeParam       := 'SASAISIE1' ;
  ParamGrille ;
  GSA.RowCount         := 2 ;
  AffecteGrid(GSA,Arguments.Action) ;
  Prem:=-1 ;


  // init libellé du Tabcontrol + 1er Axe ventilable
  lIndex := 1 ;
  lQ     := OpenSql('select X_LIBELLE from AXE',true) ;
  lQ.first ;
  while not lQ.EOF do
   begin
    if not Axes[lIndex] then
     TA.Tabs[lIndex-1] := ' '
      else
       begin
        TA.Tabs[lIndex-1] := lQ.Findfield('X_LIBELLE').asString ;
        if Prem = -1 then Prem := lIndex ;
       end ;
    lQ.Next ;
    Inc(lIndex) ;
   end ; // while

  FInfo.LoadJournal(ObjEcr.GetString('E_GENERAL')) ;

  lBoUnSeulAxe := false ;
  lBoOuvreAna  := false ;
  CGereFenAna ( ObjEcr.GetString('E_GENERAL') , FInfo.Journal.GetValue('J_NATUREJAL') , lBoUnSeulAxe , lBoOuvreAna, FBoOuiTvaEnc ,lBoVentilTva) ;

  Ferme(lQ) ;

  // Si scénario, placement de l'axe courant, sinon 1er axe ventilable
  if ((Arguments.NumGeneAxe>0) and (Axes[Arguments.NumGeneAxe])) then
   CurAxe:=Arguments.NumGeneAxe
    else
     CurAxe:=Prem ;

 // préventilation uniquement si guide analytique
  if (Arguments.GuideA <> '') then
    begin
    ObjEcr.ClearDetail ;
    AlloueAxe( ObjEcr ) ;
    // FQ 18536 SBO 18/08/2006 : Pb sur ventilation des guides
    VentilerTOB( ObjEcr, Arguments.GuideA, ObjEcr.GetInteger('E_NUMLIGNE'), Arguments.DEV.Decimale, False, FInfo.Dossier, False ) ;
    end ;

  // Tests --> placement sur le premier axe qui pose problème
  lErr := CIsValidVentil( ObjEcr, FInfo ) ;
  if lErr.RC_Error <> RC_PASERREUR then
    CurAxe := lErr.RC_Axe ;

  // Affichage de l'axe CurAxe
  TA.TabIndex := CurAxe-1 ;
  TAChange(Nil) ;

  Case Arguments.Action of
  	taConsult : LectureSeule ;
   	end ;
  // Affichage montants calculés
  if Arguments.Action<>taCreat then AffichePied ;

  // pour S3 : Un seul axe
//  DelTabsSerie(TA) ;

{$IFDEF GCGC}
  BComplement.Visible:=False ;
  BComplement.Enabled:=False ;
{$ENDIF}

  if ctxPCL in V_PGI.PGIContexte then
    BCalculQte.Visible := False ;

  GereOptionsGrid ( 1 ) ; // FQ 16037 : SBO 23/09/2005

  {JP 24/10/05 : FQ 16879 : gestion du caption de la fiche}
  Caption := TraduireMemoire('Ventilations analytiques');

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... : 25/10/2002
Description .. : Sortie de cellule :
Suite ........ :  -> test saisie section
Suite ........ :  -> calcul montant / pourcent en fonction saisie
Suite ........ :       + affectation à la tob
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.GSACellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
Var X,VV : Double ;
begin
  if csDestroying in ComponentState then Exit ;
  if Arguments.Action = taConsult then exit ;
  VV:=E_MontantEcr.Value ;
  // Section saisie ?
  if (GSA.Row=ARow) and (ACol=AN_Sect) then
    BEGIN
    if LaMeme(GSA,ARow) then Exit ;
    if ChercheSect(ARow,False)<=0 then BEGIN Cancel:=True ; Exit ; END ;
    END ;
  // Saisie pourcentage ou montant ?
  if ((ARow=1) and (ACol=AN_Pourcent) and (ValP(ARow)<>0) and (ValM(ARow)=0)) then
    BEGIN
    // SG6 10/11/04 : Problème lors de la tabulation (pour avoir un fonctionnement identique entre saise et modification
    if ModeS=0 then
      ModeS:=1 ;
    END ;
  if ((ARow=1) and (ACol=AN_Montant) and (ValM(ARow)<>0) and (ValP(ARow)=0)) then
    BEGIN
    if ModeS=0 then
      ModeS:=2 ;
    END ;
  // Calcul pourcentage ou montant et affectation TOB
  if ((ACol=AN_Montant) or (ACol=AN_Pourcent)) then
    BEGIN
    if (ACol=AN_Pourcent) and
       ( ( Arrondi( ValP(ARow)-ValTobP(ARow) , ADecimP ) <> 0 ) // Modif manuelle
         or not (Arguments.EnMontant)                           // pas de saisie en montant
       ) then
      begin
      X:=ValP(ARow)*VV/100.0 ;
      GSA.Cells[AN_Montant,ARow]:=StrS(X,Arguments.DEV.Decimale) ;
      end
    else if (ACol=AN_Montant) and
       ( ( Arrondi( ValM(ARow)-ValTobM(ARow) , Arguments.DEV.Decimale ) <> 0 ) // Modif manuelle
         or (Arguments.EnMontant)                                              // saisie en montant
       ) then
      begin
      X:=100.0*ValM(ARow)/VV ;
      GSA.Cells[AN_Pourcent,ARow]:=StrS(X,ADecimP) ;
      end ;
    FormatMontant(AN_Montant,ARow) ;
    FormatMontant(AN_Pourcent,ARow) ;
    TraiteMontant(ARow,True) ;
    GereArrondi( ARow, ACol ) ;
    END ;

  if ( ACol = AN_Date ) and ((length(trim(GSA.Cells[AN_Date,ARow]))) > 10 ) then
   begin
    if not CControleDateBor(strToDate(GSA.Cells[AN_Date,ARow]), ctxExercice ) then
     GSA.Cells[AN_Date,ARow] := DatetoStr(ObjEcr.GetValue('E_DATECOMPTABLE')) ;
   end ;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : A l'entrée d'une cellule :
Suite ........ :  -> Blocage sur section
Suite ........ :  -> Gestion déplacement
Mots clefs ... : 
*****************************************************************}
procedure TFeSaisAnal.GSACellEnter(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
begin

  if ARow<>GSA.Row then GereOptionsGrid(GSA.Row) ;

  // Si section non saisie alors replacement sur section
  if ((Arguments.Action<>taConsult) and (Not EstRempli(GSA,GSA.Row)) and (GSA.Col<>AN_Sect)) then
    BEGIN
    ACol:=AN_Sect ;
    Cancel:=True ;
    Exit ;
    END ;
  GereNewLigne(GSA) ;
  if Arguments.Action=taConsult then exit ;
  // Gestion du déplacement automatique dans la grille
  if GSA.Col=AN_Lib then
    BEGIN
    if Acol<=AN_Sect
      then Acol:=AN_Pourcent
      else ACol:=AN_Sect ;
    ARow := GSA.Row ;
    Cancel:=True ;
    Exit ;
    END;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : A l'entrée d'une ligne :
Suite ........ :  -> si besoin création d'une nouvelle ligne
Suite ........ :  -> MAJ du pied
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.GSARowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
if Not EstRempli(GSA,Ou) then DefautLigne(Ou) ;
GereOptionsGrid( Ou ) ;
AffichePied ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Appel la fenêtre de saisie complémentaire analytique
Suite ........ : UNIQUEMENT EN S7
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.GereComplements ( Lig : integer ) ;
{$IFNDEF GCGC}
Var OBA    : TOBM ;
    ModBN  : boolean ;
    RC     : R_COMP ;
    CSect  : TGSection ;
    lOS    : TOBMSupp ;
{$ENDIF}
BEGIN
{$IFNDEF GCGC}
  if Arguments.Action=taConsult then Exit ;
  if Not EstSerie(S7) then Exit ;
  if Not VH^.CPLibreAnaObli then Exit ;
  if PasToucheLigne( Lig ) then Exit ; // FQ 16037

  OBA:=GetO(GSA,Lig) ;
  if OBA=Nil then Exit ;

  lOS := TOBMSupp( GSA.Objects[AN_Lib,Lig] ) ;
  if (lOS <> nil) and lOS.CompS then Exit ;

  CSect:=GetGSect(GSA,Lig) ;
  if CSect=Nil then Exit ;

  RC.StLibre:='XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX' ;
  RC.StComporte:='--XXXXXXXX' ;
  RC.Conso:=True ;
  RC.Attributs:=False ;
  RC.MemoComp:=Nil ;
  RC.Origine:=-1 ;
  RC.DateC:=0 ;
  RC.TOBCompl := nil ;

  SaisieComplement(OBA,EcrAna,Arguments.Action,ModBN,RC,VH^.CPLibreAnaObli) ;

  if lOS <> nil then
    lOS.CompS := True ;
  if GSA.CanFocus then GSA.SetFocus ;

{$ENDIF}
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : A la sortie de la ligne :
Suite ........ :  -> Vérification validité donnée sinon blocage
Suite ........ :  -> Si ok, appel commplément
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.GSARowExit(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
Var OkL : Boolean ;
begin
  if csDestroying in ComponentState then Exit ;
  if Arguments.Action=taConsult then Exit ;
  if Ou>=GSA.RowCount-1 then Exit ;
  // Grille inacessible pendant vérification
  GridEna(GSA,False) ;
  // Vérif ligne
  OkL:=LigneCorrecte(Ou,True) ;
  // Appel complément
  if OkL
    then GereComplements(Ou)
    else Cancel := True ;
  // Grille de nouveau accessible
  GridEna(GSA,True) ;
end;

procedure TFeSaisAnal.FormKeyPress(Sender: TObject; var Key: Char);
begin
if Not GSA.SynEnabled then Key:=#0 else
if Key=#127 then Key:=#0 else
if Key='=' then Key:=#0 ;
end;

procedure TFeSaisAnal.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var OkG,Vide : Boolean ;
    lInIdx : integer ;
    lBoOk : boolean ;
begin
if Not GSA.SynEnabled then BEGIN Key:=0 ; Exit ; END ;
OkG:=(ActiveControl=GSA) ; Vide:=(Shift=[]) ;
if ((ModeVentil) and (Key<>VK_F10) and (Key<>VK_ESCAPE)) then Exit ;
Case Key of
   VK_F5     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickZoom ; END else
{$IFDEF CCS3}
                ;
{$ELSE}
                  If OkG And ((Shift=[ssAlt]) Or (Shift=[ssCtrl]))then BEGIN Key:=0 ; ClickZoomSousPlan ; END ;
{$ENDIF}
   VK_F6     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickSolde ; END ;
   VK_F10    : if ModeVentil
                 then BNewVentilClick(Nil)
                else if ((OkG) and (Vide)) then
                  begin
                  Key:=0 ;
                  if BValide.Enabled then
                    ClickValide ;
                  end ;
   VK_RETURN : if ((OkG) and (Vide)) then Key:=VK_TAB ;
   VK_INSERT : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickInsert ; END ;
   VK_DELETE : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; ClickDel ; END ;
   VK_ESCAPE : if ModeVentil then FinVentil else if Vide then ClickAbandon ;
   VK_BACK   : if ((OkG) and (Shift=[ssCtrl])) then BEGIN Key:=0 ; VideZone(GSA) ; END ;
   VK_TAB    : if OkG and ( (Shift=[ssCtrl]) or (Shift=[ssShift,ssCtrl]) ) then
                 begin
                 lInIdx := TA.TabIndex ;
                 lBoOk := True ;
                 if Shift=[ssCtrl] then
                   begin
                   if TA.TabIndex < (TA.Tabs.Count - 1)
                     then lInIdx := TA.TabIndex + 1
                     else lInIdx := 0 ;
                   end
                 else if Shift=[ssShift,ssCtrl] then
                   begin
                   if TA.TabIndex > 0
                     then lInIdx := TA.TabIndex - 1
                     else lInIdx := TA.Tabs.Count - 1 ;
                   end ;
                 if lInIdx <> TA.TabIndex then
                   begin
                   TAChanging( nil, lBoOk ) ;
                   if lBoOk then
                     begin
                     TA.TabIndex := lInIdx ;
                     TAChange(Nil) ;
                     end ;
                   end ;
                 end ;

 {=}     187 : if ((OkG) and (Vide)) then BEGIN Key:=0 ; ClickSolde ; END ;
{AC}      67 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickComplement(GSA.Row) ; END ;
{AQ}      81 :
                if vh^.cpifdefcegid then
                begin
                     if Shift=[ssAlt] then BEGIN Key:=0 ; ZoomVentilType ; END ;
                end ;
{AV}      86 : if Shift=[ssAlt] then BEGIN Key:=0 ; ClickVentilType ; END ;
   END ;
end;

procedure TFeSaisAnal.FormClose(Sender: TObject; var Action: TCloseAction);
begin

  if Arguments.Action <> taConsult then
    if ( ObjEcr.GetNumChamp('NEWVENTIL') > 0 ) and ( ObjEcr.GetString('NEWVENTIL') = 'X' ) then
      ObjEcr.PutValue('NEWVENTIL', '-') ;

  DeAlloueObj ;
  GSA.VidePile(False) ;
  PurgePopup(POPS) ;
  RegSaveToolbarPos(Self,'SaisAnal') ;
  {JP 13/01/05 : FQ 17160 : Destruction de la tob de "sauvegarde"}
  if Assigned(ObjTemp) then FreeAndNil(ObjTemp);

  FreeAndNil(FRestriction);                   {FP 29/12/2005}

  // Libération nouveaux objets si non passé en paramètre
  if not Assigned(Arguments.Info) and Assigned(FInfo) then
    FreeAndNil(FInfo);
  if not Assigned(Arguments.MessageCompta) and Assigned(FMessageCompta) then
    FreeAndNil(FMessageCompta);

end;

procedure TFeSaisAnal.GSAMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
GX:=X ; GY:=Y ;
end;

procedure TFeSaisAnal.POPSPopup(Sender: TObject);
begin
InitPopUp(Self) ;
end;

procedure TFeSaisAnal.TAChanging(Sender: TObject; var AllowChange: Boolean);
Var lBoCancel : boolean ;
    lBoChange : boolean ;
begin

  if Arguments.Action=taConsult then Exit ;

  // Validation de la ligne courante
  lBoCancel := False ;
  lBoChange := True ;
  GSARowExit(Nil,GSA.Row,lBoCancel,lBoChange) ;
  if lBoCancel then
    AllowChange := False
  else if not ValideLeAxe(CurAxe,True) then
    AllowChange := False ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Sur changement d'axe :
Suite ........ :  -> sur pas ventilable : retour en arrière
Suite ........ :  -> sinon maj grille
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.TAChange(Sender: TObject);
begin
  // Si l'axe sélectionner n'est pas ventilable, on reste sur l'axe courant
  if Not Axes[TA.TabIndex+1] then
    BEGIN
    TA.TabIndex:=CurAxe-1 ;
    Exit ;
    END ;
  // Changement d'axe courant
  CurAxe:=TA.TabIndex+1 ;
  // Init Grille
  PositionneTz(Cache,CurAxe) ;
  GSA.CacheEdit ;
  GSA.Row:=1 ;
  GSA.Col:=AN_Sect ;
  DeAlloueObj ;
  GSA.VidePile(false) ;
  GSA.MontreEdit ;
  // Chargement des données pour la grille
  ChargeGrid(CurAxe) ;
  InitGrid ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... : 25/10/2002
Description .. : Sauvegarde du montant de l'écriture dans MTEntree avant
Suite ........ : modif
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.E_MontantEcrEnter(Sender: TObject);
begin
  MTEntree := E_MONTANTECR.Value ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Sur modification du montant écriture à répartir :
Suite ........ :  -> Réaffectation écriture,
Suite ........ :  -> Recalcul ventilation sur tous les axes
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.E_MontantEcrExit(Sender: TObject);
Var XX,XP,XD,X : Double ;
    i,R       : integer ;
begin
  if csDestroying in ComponentState then Exit ;

  // Y'a-t-il eu modification ?
  XX:=E_MONTANTECR.Value ;
  if Arrondi(XX-MtEntree,Arguments.DEV.Decimale)=0 then Exit ;

  if ObjEcr.GetValue('E_DEVISE')=V_PGI.DevisePivot then
    BEGIN // saisie pivot ???
    XP:=XX ;
    XD:=XX ;
    END
  else
    BEGIN
    XP:=DeviseToEuro(XX,Arguments.DEV.Taux,Arguments.DEV.Quotite) ;
    XD:=XX ;
    END ;

  // Affectation nouveaux montants
  MontantEcrP:=XP ; MontantEcrD:=XD ;
  // Ré-affectation totaux sur l'OBM
  if Sens = 1 then
    begin
    ObjEcr.PutValue('E_DEBIT',XP);
    ObjEcr.PutValue('E_DEBITDEV',XD);
    end
  else
    begin
    ObjEcr.PutValue('E_CREDIT',XP);
    ObjEcr.PutValue('E_CREDITDEV',XD);
    end;
  // Recalcul ventil pour les autres axes
  for i:=1 to MaxAxe do
    if ((Axes[i]) and (i<>CurAxe)) then
      begin
      RecalculProrataAnalNEW(Pf,ObjEcr,i,Arguments.DEV) ;
      end;
  // Recalcul ventil pour axe courant
  for R:=1 to GSA.RowCount-2 do
    BEGIN
    X:=ValP(R)*XX/100.0 ;
    GSA.Cells[AN_Montant,R]:=StrS(X,Arguments.DEV.Decimale) ;
    TraiteMontant(R,False) ;
    FormatMontant(AN_Montant,R) ;
    AlimObjetMvt(R) ;
    END ;
  CalculTotal(CurAxe) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 25/10/2002
Modifié le ... : 25/10/2002
Description .. : Sur changement de ventilation "type" :
Suite ........ :  -> re-préventilation de tous les axes
Suite ........ :  -> rechargement de la grille
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.CVTypeChange(Sender: TObject);
begin
  if CVType.Value='' then exit ;
  DeAlloueObj ;
  GSA.VidePile(False) ;
  // Ventilation type
  ObjEcr.ClearDetail ;
  AlloueAxe( ObjEcr ) ;

  //Execution de la ventilation : FQ 18723
  VentilerTob( ObjEcr, '', 0, Arguments.DEV.Decimale, False, '', False, CVType.Value) ;

  TAChange(nil);
  GSA.SetFocus ;
  Arguments.DernVentilType := CVType.Value ;
end;

procedure TFeSaisAnal.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFeSaisAnal.FormCreate(Sender: TObject);
Var i : integer ;
begin
  WMinX:=Width ; WMinY:=228 ;
  for i:=1 to MaxAxe do Conf[i]:='0' ;
  RegLoadToolbarPos(Self,'SaisAnal') ;
  FRestriction   := TRestrictionAnalytique.Create;          {FP 29/12/2005}
  FInAxeTva      := - 1 ;
  if GetParamSocSecur('SO_CPPCLSAISIETVA',false) then
   FInAxeTva      := StrToInt(Copy(GetParamSocSecur('SO_CPPCLAXETVA', ''),2,1) ) ;
end;

procedure TFeSaisAnal.GSAKeyPress(Sender: TObject; var Key: Char);
begin
if Not GSA.SynEnabled then Key:=#0 ;
if GSA.Col = AN_Date then
  ParamDate(self,sender,key);
end;

procedure TFeSaisAnal.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFeSaisAnal.AlloueOS(Lig: integer);
var lOS : TOBMSupp ;
begin
 if Not EstSerie(S7) then Exit ;
  if not Assigned( GSA.Objects[AN_Lib,Lig] ) then
    begin
    lOS := TOBMSupp.Create ;
    lOS.CompS := False ;
    GSA.Objects[AN_Lib,Lig] := TObject(lOS) ;
    end ;
end;

procedure TFeSaisAnal.DeAlloueObj( vNumLigne : integer = - 1 ) ;
var lOS    : TOBMSupp ;
    lig    : integer ;
    OBA    : TOB;
    lInDeb : integer ;
    lInFin : integer ;
begin

  if vNumLigne > 0 then
    begin
    lInDeb := vNumLigne ;
    lInFin := vNumLigne ;
    end
  else
    begin
    lInDeb := 1 ;
    lInFin := GSA.RowCount - 1 ;
    end ;

 for lig := lInDeb to lInFin do
  begin
   if Assigned( GSA.Objects[AN_Lib,Lig] ) then
    begin
     lOS := TOBMSupp(GSA.Objects[AN_Lib,Lig]) ;
     if Assigned(lOS) then lOS.Free ;
     GSA.Objects[AN_Lib,Lig] := nil ;
    end ;
    DesalloueA(lig) ;
    OBA := GetTOB(GSA,Lig) ;
    if (OBA<>nil) and (OBA.Parent = nil) then
     begin
      GSA.Objects[AN_NumL,Lig] := nil ;
      OBA.Free ;
     end ;
  end ; // for
end;


{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 23/09/2005
Modifié le ... :   /  /
Description .. : Ajout pour FQ16037 et gestion des sections fermées :
Suite ........ : Change dynamiquement les options d'accès à la grille en
Suite ........ : fonction du résultat de la fonction PasToucheLigne.
Mots clefs ... :
*****************************************************************}
procedure TFeSaisAnal.GereOptionsGrid ( Lig : integer ) ;
Var Okok   : boolean ;
    lBoPTL : boolean ;
begin

  Okok := (GoRowSelect in GSA.Options) ;

  if ( Arguments.Action = taModif ) then
    begin
    lBoPTL := PasToucheLigne(Lig) ;
    if lBoPTL then
      begin
      GSA.CacheEdit ;
      GSA.Options:=GSA.Options-[GoEditing,GoTabs,GoAlwaysShowEditor] ;
      GSA.Options:=GSA.Options+[GoRowSelect] ;
      end
    else
      begin
      GSA.Options:=GSA.Options-[GoRowSelect] ;
      GSA.Options:=GSA.Options+[GoEditing,GoTabs,GoAlwaysShowEditor] ;
      GSA.MontreEdit ;
      end ;
    end ;

  if Okok <> (GoRowSelect in GSA.Options) then
    GSA.Refresh ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 23/09/2005
Modifié le ... :   /  /
Description .. : Ajout pour FQ16037 et gestion des sections fermées :
Suite ........ : Retourne Vrai si la ligne de doit pas pouvoir être modifiée.
Mots clefs ... :
*****************************************************************}
function TFeSaisAnal.PasToucheLigne(Lig: integer): boolean;
//Var CSect : TGSection ;
begin

  Result := False ;

  if Arguments.Action = taCreat then
    Exit ;

  // SBO 23/09/2005 FQ 16037 modif interdite sur les sections fermées ;
{ SBO 20/09/2006 De coté car bloquant en modif de ventil si appel d'une ventil type avec section fermée
  CSect    := GetGSect( GSA, Lig ) ;
  if CSect <> Nil then
    Result := CSect.Ferme ;
}
end;

{JP 13/01/06 : Si on annule, on récupère les données de l'ouverture de la fiche
{---------------------------------------------------------------------------------------}
procedure TFeSaisAnal.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
var
 i : integer ;
begin
  {$IFDEF TT}
  Close ;
  exit ;
  {$ENDIF}

 // DEV3946 - test ventil correcte
 if Arguments.VerifVentil and ( (Arguments.Action = taCreat) or
                                ((Arguments.Action = taModif) and (not Assigned( ObjTemp )) )
                              )
 then
   begin
   // SBO 22/10/2007 : FQ 21681 gestion des erreurs ignorées pour pb avec sections fermées (cf uLibAnalytique.pas)
   if (Arguments.Action <> taCreat) then
     FInfo.AjouteErrIgnoree( [RC_YSECTIONFERMEE] ) ;
   for i:=1 to MAXAXE do
     if (Axes[i]) then begin
       CalculTotal(i) ;
       if (not ValideLeAxe(i, TRUE)) then begin
         CanClose := False ;
         break ;
       end ;
     end ;
   if (Arguments.Action <> taCreat) then
     FInfo.AjouteErrIgnoree( [] ) ;
   if not CanClose then
     Exit ;
   // Fin FQ 21681 gestion des erreurs ignorées pour pb avec sections fermées
   end ;

  if Assigned( ObjTemp ) and IsCanceled and (Arguments.Action = taModif)
    and ( EstVentilModifiee ) then begin
    CanClose := HShowMessage('0;' + Caption + ';Êtes-vous sûr de vouloir abandonner votre saisie ?;I;YNC;N;C', '', '') = mrYes;
    IsCanceled := False;
    if CanClose then begin
      // FQ 18147 Remise en place de l'annulation :
      // Gestion de la libération mémoire
      for i := 1 to (GSA.RowCount - 1) do begin
        if GSA.Objects[AN_NumL,i]<>nil then
          GSA.Objects[AN_NumL,i].free ;
        GSA.Objects[AN_NumL,i] := nil ;
        end ;
      ObjEcr.ClearDetail;
      ObjEcr.Dupliquer(ObjTemp, True, True);
    end;
  end;
end;

procedure TFeSaisAnal.OnError(sender: TObject; Error: TRecError);
var lStText : string ;
begin
  if ( Error.RC_Error <> RC_PASERREUR ) then
    begin
    lStText := FMessageCompta.GetMessage( Error.RC_Error ) ;
    if (Error.RC_Axe > 0) and (Error.RC_Axe <= MAXAXE) then
      lStText := 'AXE n°' + IntToStr(Error.RC_Axe) + ' : ' + RechDom('TTAXE', 'A'+IntToStr(Error.RC_Axe), False )
                 + #10#13 + lStText ;
    end ;
  if Trim( lStText ) <> '' then
    PgiBox( lStText, Caption ) ;
end;

function TFeSaisAnal.EstVentilModifiee : boolean ;
Var lTobE : Tob ;
    lTobT : Tob ;
    i     : integer ;
    j     : integer ;
begin

  result := False ;
  if Arguments.Action = taConsult then Exit ;
  if not Assigned( ObjTemp ) then Exit ;
  if ObjTemp.Detail.Count <> MAXAXE then Exit ;
  if ObjEcr.Detail.Count <> MAXAXE then Exit ;

  // On parcours les axes et à la 1ère modi, on sort....
  result := True ;
  for i := 1 to MAXAXE do
    begin
    // Nb de lignes ?
    lTobT := ObjEcr.Detail[i-1] ;
    lTobE := ObjTemp.Detail[i-1] ;

    if lTobT.Detail.Count <> lTobe.Detail.Count then Exit ;

    // parcours des ventilations
    for j := 0 to ( lTobT.Detail.Count - 1 ) do
      if ( lTobT.Detail[j].GetString('Y_SECTION') <> lTobE.Detail[j].GetString('Y_SECTION') ) or
         ( lTobT.Detail[j].GetInteger('Y_NUMVENTIL')  <> lTobE.Detail[j].GetInteger('Y_NUMVENTIL') ) or
         ( lTobT.Detail[j].GetDouble('Y_POURCENTAGE') <> lTobE.Detail[j].GetDouble('Y_POURCENTAGE') ) or
         ( GetMontant( lTobT.Detail[j] ) <> GetMontant( lTobE.Detail[j] ) )
         then Exit ;

    end ;

  // Ventil identique
  result := False ;
end;

procedure TFeSaisAnal.BCALCULQTEClick(Sender: TObject);
var lTobVentil  : TOB ;
    i           : integer ;
    lPourc      : double ;
    lTotalQte1  : double ;
    lTotalQte2  : double ;
    lTotalPourc : double ;
    lEcrQte1    : double ;
    lEcrQte2    : double ;
begin
  if Arguments.Action=taConsult then Exit ;
  if Not EstRempli(GSA,GSA.Row) then Exit ;
  if Not LigneCorrecte(GSA.Row,True) then Exit ;

  lTotalPourc := 0 ;
  lTotalQte1  := 0 ;
  lTotalQte2  := 0 ;
  lEcrQte1    := ObjEcr.GetDouble('E_QTE1') ;
  lEcrQte2    := ObjEcr.GetDouble('E_QTE2') ;

  for i := 0 to ObjEcr.Detail[CurAxe - 1].Detail.Count - 1 do
    begin

    lTobVentil  := ObjEcr.Detail[CurAxe - 1].Detail[ i ] ;
    lPourc      := lTobVentil.GetDouble('Y_POURCENTAGE') ;
    lTotalPourc := lTotalPourc + lPourc ;

    // Qté 1
    if lEcrQte1<>0 then
      begin
      lTobVentil.PutValue('Y_POURCENTQTE1',	      lPourc                           ) ;
      lTobVentil.PutValue('Y_QUALIFECRQTE1',      ObjEcr.GetString('E_QUALIFQTE1') );
      lTobVentil.PutValue('Y_QUALIFQTE1',         ObjEcr.GetString('E_QUALIFQTE1') );
      lTobVentil.PutValue('Y_TOTALQTE1',          lEcrQte1                         );
      if Arrondi( lTotalPourc - 100.0,ADecimP) <> 0
        then lTobVentil.PutValue('Y_QTE1', Arrondi( lEcrQte1 * lPourc / 100.0 , V_PGI.OkDecQ ) )
        else lTobVentil.PutValue('Y_QTE1', Arrondi( lEcrQte1 - lTotalQte1     , V_PGI.OkDecQ ) ) ;
      lTotalQte1 := lTotalQte1 + lTobVentil.GetDouble('Y_QTE1') ;
      end
    else
      begin
      lTobVentil.PutValue('Y_POURCENTQTE1', 0 ) ;
      lTobVentil.PutValue('Y_QTE1',         0 ) ;
      lTobVentil.PutValue('Y_TOTALQTE1',    0 ) ;
      end ;

    // Qté 2
    if lEcrQte2<>0 then
      begin
      lTobVentil.PutValue('Y_POURCENTQTE2',	      lPourc                          ) ;
      lTobVentil.PutValue('Y_QUALIFECRQTE2',      ObjEcr.GetString('E_QUALIFQTE2') );
      lTobVentil.PutValue('Y_QUALIFQTE2',         ObjEcr.GetString('E_QUALIFQTE2') );
      lTobVentil.PutValue('Y_TOTALQTE2',          lEcrQte2                        );
      if Arrondi( lTotalPourc - 100.0,ADecimP) <> 0
        then lTobVentil.PutValue('Y_QTE2', Arrondi( lEcrQte2 * lPourc / 100.0 , V_PGI.OkDecQ ) )
        else lTobVentil.PutValue('Y_QTE2', Arrondi( lEcrQte2 - lTotalQte2     , V_PGI.OkDecQ ) ) ;
      lTotalQte2 := lTotalQte2 + + lTobVentil.GetDouble('Y_QTE2') ;
      end
    else
      begin
      lTobVentil.PutValue('Y_POURCENTQTE2', 0 ) ;
      lTobVentil.PutValue('Y_QTE2',         0 ) ;
      lTobVentil.PutValue('Y_TOTALQTE2',    0 ) ;
      end ;

    end ;

  // Mémorisation test // DEV3946
  if ObjEcr.GetNumChamp('CHECKQTE') < 0
    then ObjEcr.AddChampSupValeur('CHECKQTE', '-' )
    else ObjEcr.PutValue('CHECKQTE', '-' ) ;

end;

procedure TFeSaisAnal.ParamGrille ;
begin

 GSA.ColWidths[AN_NumL]       := 2 ;
 GSA.ColWidths[AN_Sect]       := 6 ;
 GSA.ColWidths[AN_Lib]        := 18 ;
 GSA.ColWidths[AN_Pourcent]   := 7 ;
 GSA.ColWidths[AN_Montant]    := 7 ;

 GSA.Cells[AN_NumL,0]         := TraduireMemoire('Lig') ;
 GSA.Cells[AN_Sect,0]         := TraduireMemoire('Section') ;
 GSA.Cells[AN_Lib,0]          := TraduireMemoire('Intitulé') ;
 GSA.Cells[AN_Pourcent,0]     := TraduireMemoire('%') ;
 GSA.Cells[AN_Montant,0]      := TraduireMemoire('Montant') ;

// if Arguments.OuiTvaEnc then
 if ( FBoOuiTvaEnc ) and ( CurAxe = FInAxeTva ) then
  begin
   GSA.ColCount              := 6 ;
   GSA.ColWidths[AN_Date]    := 10 ;
   GSA.Cells[AN_Date,0]      := TraduireMemoire('Date') ;
   GSA.ColTypes[AN_Date]     := 'D';
   GSA.ColFormats[AN_Date]   := ShortDateFormat;
  end
   else
    begin
     GSA.ColCount := 5 ;
    end ;

  GSA.ColTypes[AN_Pourcent]    := 'R';
  GSA.ColTypes[AN_Montant]     := 'R';

 HMTrad.ResizeGridColumns(GSA) ;

end ;



function TFeSaisAnal.ValTobM(Lig: integer): double;
Var OBA     : TOB ;
begin
  result := 0 ;
  OBA := GetTOB(GSA,Lig) ;
  if OBA = Nil then Exit ;
  result := OBA.GetDouble(Pf+'_CREDITDEV') + OBA.GetDouble(Pf+'_DEBITDEV') ;
end;

function TFeSaisAnal.ValTobP(Lig: integer): double;
Var OBA     : TOB ;
begin
  result := 0 ;
  OBA := GetTOB(GSA,Lig) ;
  if OBA = Nil then Exit ;
  result := OBA.GetDouble(Pf+'_POURCENTAGE') ;
end;

procedure TFeSaisAnal.GereArrondi( Lig, ColTest : integer);
var lDiff : double ;
begin
  if ColTest = AN_Pourcent then
    begin
    if ( Arrondi( 100.0 - (SI_TotPourc[CurAxe])*100 , ADecimP ) = 0 ) and // total pourcentage 100%
       ( Arrondi( E_MontantEcr.Value - SI_TotSais[CurAxe], Arguments.DEV.Decimale ) <> 0 ) // Mais total saisie <> 100%
       then
      begin
//      lDiff := Arrondi( E_MontantEcr.Value - SI_TotSais[CurAxe], Arguments.DEV.Decimale ) ;
      lDiff := Arrondi( E_MontantEcr.Value - ( SI_TotSais[CurAxe] - ValM(Lig) ), Arguments.DEV.Decimale ) ;
      if lDiff=0 then Exit ;
      GSA.Cells[AN_Montant,Lig] := StrS( lDiff, Arguments.DEV.Decimale ) ;//StrS( ValM(Lig) + lDiff, Arguments.DEV.Decimale ) ;
      TraiteMontant( Lig, True ) ;
      end ;
    end
  else if ColTest = AN_Montant then
    begin
    if ( Arrondi( E_MontantEcr.Value - SI_TotSais[CurAxe], Arguments.DEV.Decimale ) = 0 ) // total saisie = 100%
       and ( Arrondi( 100.0 - (SI_TotPourc[CurAxe])*100 , ADecimP ) <> 0 )                    // Mais total pourcent <> 100%
       then
      begin
//      lDiff := Arrondi( (1.0 - SI_TotPourc[CurAxe])*100.0, ADecimP ) ;
      lDiff := Arrondi( 100 - ((SI_TotPourc[CurAxe]*100.0)-ValP(Lig)), ADecimP ) ;
      if lDiff=0 then Exit ;
      GSA.Cells[AN_Pourcent,Lig] := StrS( lDiff, ADecimP ) ; //StrS( ValP(Lig) + lDiff, ADecimP ) ;
      TraiteMontant( Lig, True ) ;
      end ;
    end ;
end;

end.





