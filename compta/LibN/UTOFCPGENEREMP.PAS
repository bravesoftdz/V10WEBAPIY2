{-------------------------------------------------------------------------------------
    Version  |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
               07/04/03  SBO  Création de l'unité
               18/02/05  JP   Gestion du champ E_BANQUEPREVI
               09/08/05  JP   FQ 15815 : Gestion du champ E_BANQUEPREVI dans la remise en banque de BOR
 7.00.001.001  12/01/06  JP   FQ 17299 : Suppression de EstSerie(S3) notamment pour les banques previ
 8.00.001.014  10/05/07  JP   FQ 19608 : dessin des carrés rouges en CWAS si présence d'un RIB cf DrawCell
 8.00.001.015  15/05/07  JP   FQ 17309 : Gestion des cumuls en devise pour les virements internationaux
 8.00.001.017  29/05/07  JP   FQ 20441 : gestion du multi-sélection en eAGL suite à la FQ 19608
 8.00.001.018  01/06/07  JP   FQ 20308 : Branchement de la saisie Boredereau en eAGL
 8.00.001.019  07/06/07  JP   FQ 15793 : Remplacement du THEDIT BANQUEPREVI par un MultiValComboBox
--------------------------------------------------------------------------------------}
Unit UTOFCPGENEREMP ;

Interface

Uses
    Windows,
     {$IFDEF VER150} Variants, {$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    ExtCtrls,  // pour le TPanel
    menus,     // pour le TMenuItem
    Classes,
    StdCtrls,
    Controls,
    HTB97,     // pour TToolBarButton97
    Grids,     // Pour le TGridDrawState
    Graphics,  // pour clRed
{$IFDEF EAGLCLIENT}
    eMul,
    maineagl,
    CPLETREGUL_TOF,
{$ELSE}
    FE_MAIN,
    Mul,
    db,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    HDB,
    dbGrids,
    LetRegul,
{$ENDIF}
    SaisBor,       // Pour lanceSaisieFolio FQ 20308
    Saisie,        // pour TrouveEtLanceSaisie
    HQry,
    HEnt1,
    Ent1,
    HMsgBox,
    HStatus,       // pour InitMove et MoveCur et FiniMove
    ParamSoc,      // pour le GetParamSocSecur
    UTOF,
    UTOB,
    uLibWindows,   // pour CDessineTriangle, CVireLigne
    UtilPGI,       // pour Resolution
    GenerMP,       // pour tGenereMP
    GenereMP,      // pour ParamsMP
    MPEscompte,    // pour SaisieMPEscompte
    EncUtil,       // pour CatToMP
    FichComm,      // pour
    SaisLot,       // pour
    DetruitLot,    // Pour DetruitLeLot
    Filtre,        // pour VideFiltre
    TofVerifRib,   // pour CPLanceFiche_VerifRib
    SaisUtil,      // Pour TSuiviMP
    SaisComm,      // pour TOBToIdent
    LettUtil,      // Pour RempliSelectEuro
    Lookup,
    HCtrls
    ;

{$IFDEF EAGLCLIENT}
const dgEditing = goEditing;
{$ENDIF}

Procedure GenereSuiviMP ( smp : TSuiviMP ) ;
Procedure CPLanceFiche_GenereMP ( vStRange, vStLequel, vStArgs : String ) ;

Type
  TOF_CPGENEREMP = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnDisplay                ; override ;
      
    // EVT FILTRES
    procedure E_EXERCICEChange(Sender: TObject);      // ok  C
    procedure CATEGORIEChange(Sender: TObject);       // ok  C
    procedure FTICChange(Sender: TObject);            // ok  C
    procedure cFactCreditChange(Sender: TObject);     // ok
    procedure TraiteEditeChange(Sender: TObject);     // ok  C
    procedure ESCMETHChange(Sender: TObject);         // new
    procedure ESCCBTVAChange(Sender: TObject);        // new
    procedure FEscompteChange(Sender: TObject);       // new
    procedure ESCCBPRORATAChange(Sender: TObject);    // new
    procedure ESCTVAChange(Sender: TObject);          // new
    procedure ESCHTChange(Sender: TObject);           // new
    procedure _chk_RegSupChange(Sender: TObject);     // new
    procedure _E_NATUREPIECEChange(Sender: TObject);  // new
    procedure SpoolerChange(Sender: TObject);         // new

    // EVT BOUTONS ANCETRE MUL
    procedure bSelectAllClick(Sender: TObject);     // ok C
    procedure BParamListeClick(Sender: TObject);    // new
    procedure BOuvrirClick(Sender: TObject);        // new
    // EVT BOUTONS
    procedure BCtrlRibClick(Sender: TObject);       // ok C
    procedure BSwapSelectClick(Sender: TObject);    // ok C
    procedure BParamsClick(Sender: TObject);        // new
    procedure BEscompteClick(Sender: TObject);      // new
    procedure BLotEcrClick(Sender: TObject);        // new

    // EVT Grille
    procedure FListeDblClick(Sender: TObject);      // ok  C
{$IFDEF EAGLCLIENT}
    procedure FListeDrawCell(Sender: TObject; ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);{JP 10/05/07 : FQ 19608}
{$ELSE}
    procedure FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FListeSortColumn(Sender : TObject);
{$ENDIF}
    procedure FListeRowEnter(Sender: TObject);      // ok  C
    procedure FListeFlipSelection(Sender: TObject); // ok  C
    procedure FListeColEnter(Sender: TObject);      // new
    {JP 18/02/05 : Affectation de la banque prévisionnelle}
    procedure AffecteBanqueClick(Sender : TObject);
    // EVT Form
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);   // ok
    procedure FetchLesTous;
{b fb 27/10/2005 FQ12487}
    procedure SelectionElipsisClick(Sender : TObject);
    procedure AuxiliaireElipsisClick(Sender : TObject);
{e fb 27/10/2005 FQ12487}

    procedure OnAfterSelectFiltreEcran;
    procedure AuxiElipsisClick         ( Sender : TObject );
  private
  	Q               : THQuery ;
{$IFDEF EAGLCLIENT}
    FListe          : THGrid ;
{$ELSE}
    FListe          : THDBGrid ;
{$ENDIF}
    Pages           : TPageControl ;
    FFiltres        : THValComboBox ;
    {JP 19/08/04 : FQ 14212}
    lMenuItem : TMenuItem ;

    // Composants géés manuellement // commun
    FTIC            : TCheckBox ;       // ok  C
    CATEGORIE       : THValComboBox ;   // ok  C
    cFactCredit     : TCheckBox ;       // ok  C
    TraiteEdite     : TCheckBox ;       // ok  C

    // onglet standard
    Selection, Auxiliaire : THCritMaskEdit;  {fb 27/10/2005 FQ 12487}
    // onglet critères (avec banque / lot...)
    // onglet compléments
    _chk_RegSup     : TCheckBox;

    // Boutons Toolbar
    BSwapSelect     : TToolbarButton97 ;  // ok C

    // Messages
    HDiv            : TStringList ;
    HM              : TStringList ;

    // Evt Ancêtre
{$IFDEF EAGLCLIENT}
{$ELSE}
    FListeDrawColumnCellParent  : TDrawColumnCellEvent ;
{$ENDIF}
    FormKeyDownParent         : TKeyEvent ;
    bSelectAllClickParent     : TNotifyEvent ;
    BParamListeClickParent    : TNotifyEvent ;

    // Variables locales
    GereEscompte   : boolean ;
    LesCombi       : TStringList ;
    LesCritsRupt   : TStringList ;
    TOBEscomptes   : TOB ;
    TOBGHT         : TOB ;
    TOBGTVA        : TOB ;
    JalEsc         : String ;
    SwapSelect     : Boolean ;		// Gestion du mode selection inversé
    gszCaption     : String;      // Pour les messages
    bgFetchAll     : Boolean;     // Pour le bouton Sélection inversion et Select All
    
    AllSelecting   : boolean;

    // init des messagebox et menuzoom
    procedure InitVariables ;                         // ok
    procedure InitEscompte;
    procedure InitHelpContext ;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitMsgBox;                             // ok
    procedure InitCriteres ;                          // ok
    procedure InitCriteresPourLot ;
    procedure InitConditions ;
    procedure InitSelectionLot ;
    procedure UpdateConditions ;
    procedure InitE_NUMTRACHQ(Invisible : Boolean) ;  // ok  C
    procedure InitContexte ;                          // ok
    procedure InitCATEGORIE ;
    procedure ReinitWhereNatCpt(RAZ : Boolean) ;      // ok  C
    procedure PreciseMP ;                             // ok  C
    procedure MarqueOrigine ;                         // ok
    procedure PrechargeOrigines ;                     // ok
//    procedure ClickModifRib ;                         // ok  C
    Function  CalculSoldeSelection : Double ;         // ok  C
    Procedure AfficheSoldeSelection ;                 // ok  C

    // new
    Procedure Enleve(pszNat : String);                // new
    Function  ConstitueOrigines : Boolean ;
    procedure IncNumLot ;
    Function  VerifEcheMP ( Alerte : Boolean ) : boolean ;

    {Escompte}
    Procedure ClearEscomptes ;
    Procedure SwapEscompteLigne ;
    Function  FindCreerTOBEscompte ( QueFind : boolean ) : TOB ;

    procedure AlimTobEscomptes ;
    Function  AvecEscompte : Boolean ;
    Procedure MarquerLigneSansEscompte ;

    {Mode non eAGL}
    procedure ToutMarquer ;
    Function  OkLignePourEscompte : Boolean ;
    Procedure ChangeModeGrille ;
    function LInsertDB(T : Tob) : Boolean;
    procedure AfficheBanquePrevi(OnArgOk : Boolean); {JP 18/02/05 :  Gestion du champ E_BANQUEPREVI}
    procedure WhereBqPrevi; {JP 18/02/05 :  Gestion du filtre E_BANQUEPREVI}
    {Cette fonction renvoie la valeur du champ sChamp, que l'on travaille avec Q.FindField et
     GetField : cela évite de dupliquer les traitements entre NbSelected et AllSelected}
    function GetChamp(Q : THQuery; sChamp : string) : Variant;
    procedure FFiltresChange(Sender: TObject);

    // Test le nombre d'échéances sélectionnées.
    Function  YEnATrop : boolean ;

  public
    ParmsMP                    : tGenereMP ;
    TOBORig,TOBDest            : TOB ;
    smp                        : TSuiviMP ;
    SorteLettre                : TSorteLettre ;
    SelectOnLot                : Boolean ;
    NomCritLot, NomCritReste   : String ;
    InChargement               : Boolean;
    FinChargeEcran             : Boolean;
    Affichage                  : Boolean;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  {JP 18/02/05 : Affectation de la banque prévisionnelle}
  CPAFFECTEBQPREVI_TOF, AglInit, UTofMulParamGen,
  Commun, {JP 07/06/07 : FQ 15793 : SetPlusBanqueCp}
  UProcGen,  {JP 07/06/07 : FQ 15793 : GetClauseIn}
  MulSMPUtil; {13/04/07 YMO F5 sur Auxiliaire }

// =============================================================================
Procedure GenereSuiviMP ( smp : TSuiviMP ) ;
var lStArgs : String ;
begin
  lStArgs := SmpToStr(smp) ;
  CPLanceFiche_GenereMP ( '', '', lStArgs ) ;
end ;

Procedure CPLanceFiche_GenereMP ( vStRange, vStLequel, vStArgs : String ) ;
begin
  AGLLanceFiche ( 'CP' , 'CPGENEREMP' , vStRange , vStLequel , vStArgs ) ;
end ;
// =============================================================================

procedure TOF_CPGENEREMP.OnUpdate ;
begin

  inherited;

//  V_PGI.MiseSousPli := (GetCheckBoxState('CODEOMR') = cbChecked);

  ClearEscomptes ;

  AfficheSoldeSelection ;
  FListeRowEnter(nil); // 10930

end ;

procedure TOF_CPGENEREMP.OnLoad ;
begin
  Inherited ;
  UpdateConditions;
end ;

procedure TOF_CPGENEREMP.OnArgument (S : String ) ;
var vStSmp  : String ;
begin
  Inherited ;
  TFMul(Ecran).OnAfterSelectFiltre := OnAfterSelectFiltreEcran;
  // Recup smp + param fiche :
  vStSmp                    := S ;
  smp                       := StrToSmp(vStSmp) ;
  SorteLettre               := AttribSL(smp) ;
  {JP 29/11/05 : FQ 17030 : Avec les derniers agl, il faut utiliser SetDBListe, le filtre,
                 ParamDBG et la liste doivent avoir le même nom
  TFMul(Ecran).FNomFiltre   := AttribFiltre(smp) ;
  TFMul(Ecran).Q.Liste      := AttribListe(smp) ;}
  TFMul(Ecran).SetDBListe(AttribListe(smp));

  bgFetchAll := True;

  if isEncMP(smp) then
  begin
    VH^.CCMP.LotCli := True; VH^.CCMP.LotFou := False;
  end
  else
  begin
    VH^.CCMP.LotCli := False; VH^.CCMP.LotFou := True;
  end;

  AllSelecting := false;

  // Init variables locales
  InitVariables;
  // init messages
  InitMsgBox ;
  // Constexte d'aide
  InitHelpContext ;
  // Récup composants
  InitComposants ;
  // Afectation EVT Grille
  InitEvenements ;
  // Init contexte smp
  InitContexte ;

  // Init affichage des onglets
  TTabSheet(GetControl('INVISIBLE',True)).TabVisible := False ;
  InitEscompte ;
  InitSelectionLot ;

//  SetControlChecked('CODEOMR',V_PGI.MiseSousPli);
//  if (smp in [smpEncTraEdt,smpEncTraEdtNC,smpDecBorEdt,smpDecChqEdt,smpDecVirEdt,smpDecVirInEdt,smpDecBorEdtNC,smpDecChqEdtNC,smpDecVirEdtNC,smpDecVirInEdtNC]) then
//  else SetControlVisible('CODEOMR',false);

  // Init des critères du mul
  InitCriteres ;
  InitConditions ;

  {JP 18/02/05 :  Gestion de l'affichage des contrôles E_BANQUEPREVI}
  AfficheBanquePrevi(True);

//  TFMul(Ecran).Q.Manuel     := True ; // Trop tard....

  // Gestion filtre (??)
//  SauveConceptFiltre          := NivConcept[ccActionsFiltre] ;
//  NivConcept[ccActionsFiltre] := 0 ;
//  SaveFiltre( NomCritLot,   FFiltres, Pages ) ;
//  SaveFiltre( NomCritReste, FFiltres, Pages ) ;
  THValComboBox(GetControl('FFiltres')).OnChange := FFiltresChange;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

  {JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVVert"}
  TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0184_S16G1';
end ;

procedure TOF_CPGENEREMP.OnClose ;
begin
  Inherited ;
  // Libération des TOB
  FreeAndNil(TOBOrig) ;
  FreeAndNil(TOBDest) ;
  FreeAndNil(TOBEscomptes) ;
  FreeAndNil(TOBGHT) ;
  FreeAndNil(TOBGTVA) ;
  // Libération des TList
  LesCombi.Clear ;
  LesCombi.Free ;
  LesCritsRupt.Clear ;
  LesCritsRupt.Free ;

//  NivConcept[ccActionsFiltre] := SauveConceptFiltre ;

  FreeAndNil(HM) ;
  FreeAndNil(HDiv) ;

  {JP 05/06/07 : FQ 19256 : pour que si l'on ne passe pas par le Popup pour aller à CPModifEcheMP
                 on n'est pas le mul pré-renseigné}
  VH^.MPPop.MPGenPop  := '';
  VH^.MPPop.MPAuxPop  := '';
  VH^.MPPop.MPJalPop  := '';
  VH^.MPPop.MPExoPop  := '';
  VH^.MPPop.MPNumPop  := 0;
  VH^.MPPop.MPNumLPop := 0;
  VH^.MPPop.MPNumEPop := 0;
  VH^.MPPop.MPDatePop := 0;
end ;

procedure TOF_CPGENEREMP.OnDisplay;
begin
  FListeRowEnter(nil);
  FListe.SetFocus ;
end;

procedure TOF_CPGENEREMP.InitMsgBox;
begin
  // Init HM
  HM := TStringList.Create ;
  HM.Add('0;?Caption?;Vous devez renseigner la zone "Compte de sélection".;W;O;O;O;');
  HM.Add('1;?Caption?;Les lignes d''écritures ont des modes de paiement et des échéances qui diffèrent. Confirmez-vous le traitement ?;E;YN;N;N;');
  HM.Add('2;?Caption?;Les lignes d''écritures ont des modes de paiement qui diffèrent. Confirmez-vous le traitement ?;E;YN;N;N;');
  HM.Add('3;?Caption?;Les lignes d''écritures ont des echéances qui diffèrent. Confirmez-vous le traitement ?;E;YN;N;N;');
  HM.Add('4;?Caption?;Le compte de génération ne doit pas être identique au compte des échéances sélectionnées.;W;O;O;O;');
// En attente modif lettrage SBO :
//  HM.Add('4;?Caption?;Le compte de génération est identique au compte des échéances sélectionnées. Confirmez-vous le traitement ?;Q;YN;N;N;');
  HM.Add('5;?Caption?;Certains comptes de banque n''ont aucun journal associé.;W;O;O;O;');
  HM.Add('6;?Caption?;Le traitement sur ces comptes n''a pas été effectué.;W;O;O;O;');
  HM.Add('7;?Caption?;Les lignes d''écritures ont des modes de paiement qui diffèrent. Veuillez saisir le mode de paiement de contrepartie.;W;O;O;O;');
  HM.Add('8;?Caption?;Les lignes d''écritures ont des modes de paiement et des échéances qui diffèrent. Vous devez renseigner un mode de paiement.;W;O;O;O;');
  HM.Add('9;?Caption?;Vous avez demandé une génération avec escompte. Confirmez-vous le traitement ?;Q;YN;N;N;');
  HM.Add('10;?Caption?;Certaines lignes ont une échéance antérieure à la date du jour (Pas d''escompte). Confirmez-vous le traitement ?;Q;YN;N;N;');
  HM.Add('11;?Caption?;Voulez-vous rééditer ces traites ?;Q;YN;N;N;');

  // Init HDIV
  // FQ 21468 - CA - 19/09/2007 - Suppression des termes clients et fournisseurs
  HDiv := TStringList.Create ;
  HDiv.Add('Euro');
  HDiv.Add('Comptabilisation des prélèvements');
  HDiv.Add('Edition et comptabilisation des lettres-chèque');
  HDiv.Add('Comptabilisation des virements');
  HDiv.Add('Edition des Lettres-BOR'); {4}
  HDiv.Add('Edition et comptabilisation des lettres-traite');
  HDiv.Add('Remise en banque de traites');
  HDiv.Add('Remise en banque de traites');
  HDiv.Add('Remise à l''encaissement des traites');
  HDiv.Add('Remise à l''escompte des traites');
  HDiv.Add('Remise en banque');{10}
  HDiv.Add('Escompte des traites');
  HDiv.Add('Encaissements des lots');
  HDiv.Add('Décaissements des lots');
  HDiv.Add('Edition de lettres-traite');
  HDiv.Add('Factures au débit');
  HDiv.Add('Mise en portefeuille des traites');
  HDiv.Add('Mise en portefeuille des traites');
  HDiv.Add('Edition et comptabilisation des lettres-virement');
  HDiv.Add('Edition et comptabilisation de lettres-prélèvement');
  HDiv.Add('Encaissements divers');{20}
  HDiv.Add('Décaissements divers');
  HDiv.Add('Escompte à');
  HDiv.Add('TVA sur escompte');
  HDiv.Add('Numéro &BOR de');
  HDiv.Add('Numéro &Chèques de');
  HDiv.Add('Edition de lettres-traite');
  HDiv.Add('Edition et comptabilisation des Lettres-BOR');{27}
  HDiv.Add('Mise en portefeuille des chèques');
  HDiv.Add('Mise en portefeuille des cartes bleues');
  HDiv.Add('Remise en banque de chèques');{30}
  HDiv.Add('Remise en banque de cartes bleues');
  HDiv.Add('Edition des lettres-chèque');
  HDiv.Add('Edition des lettres-virement');
  HDiv.Add('Comptabilisation des virements internationaux');
  HDiv.Add('Edition des lettres-virement international');
  HDiv.Add('Edition et comptabilisation des lettres-virement international');
  HDiv.Add('Edition de lettres-prélèvement');
  HDiv.Add('Comptabilisation de prélèvements');
  HDiv.Add('39');
end;

procedure TOF_CPGENEREMP.AfficheSoldeSelection;
Var Solde : Double;
begin
	// Calcul du solde
	Solde := CalculSoldeSelection;
  // Ajout affichage du cumul
  if Solde < 0
  	then Ecran.Caption := gszCaption + ' (Sélection : ' + StrFMontant(-1*Solde,0,V_PGI.OkDecV,'',True) + ' C)'
  	else Ecran.Caption := gszCaption + ' (Sélection : ' + StrFMontant(Solde,0,V_PGI.OkDecV,'',True) + ' D)';
	// Affichage de mode "selection inversé" si besoin
  if SwapSelect then
  	Ecran.Caption := Ecran.Caption + ' [SELECTION INVERSEE] ';

	// Raffraichissement Titre
	UpdateCaption(Ecran) ;
end;

function TOF_CPGENEREMP.CalculSoldeSelection: Double;
Var
  i : Integer ;
		debit,
  credit : Double ;
  PCumul : TPanel ;
  ChpCre : string; {15/05/07}
  ChpDeb : string; {15/05/07}
begin
 	Result := 0 ;
  PCumul := TPanel( GetControl('PCumul', True) ) ;

  {JP 15/05/07 : FQ 17309 : Gestion des cumul en devises pour les virements internationaux}
  if IsVirInternational(Smp) then begin
    {En sélection globale, on travail sur les cumuls de la grille ...}
    if FListe.AllSelected or SwapSelect then begin
      {... On regarde donc s'ils existent pour les montants en devise}
      if (PCumul.FindChildControl('__QRYPCumul_E_DEBITDEV') = nil) or
         (PCumul.FindChildControl('__QRYPCumul_E_CREDITDEV') = nil) then begin
        ChpCre := 'E_CREDIT';
        ChpDeb := 'E_DEBIT';
      end
      else begin
        ChpCre := 'E_CREDITDEV';
        ChpDeb := 'E_DEBITDEV';
      end;
    end

    {... On n'est pas avec AllSelected à True}
    else begin
      {Avec la SocRef 839, les champs figureront dans les DBListes}
      ChpCre := 'E_CREDITDEV';
      ChpDeb := 'E_DEBITDEV';
    end;
  end
  else begin
    ChpCre := 'E_CREDIT';
    ChpDeb := 'E_DEBIT';
  end;

  if (PCumul.FindChildControl('__QRYPCumul_' + ChpDeb) = nil)
  	 or (PCumul.FindChildControl('__QRYPCumul_' + ChpCre) = nil) then Exit ;

  if FListe.AllSelected then
    begin
	  debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_' + ChpDeb)).text ) ;
  	credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_' + ChpCre)).text ) ;
	  Result := debit - credit ;
    end
  else	// Sinon, on parcours le Bookmark
    begin
    for i:=0 to FListe.NbSelected-1 do
      BEGIN
      FListe.GotoLeBookmark(i) ;
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.row - 1) ;
{$ENDIF}
      Result := Result + Q.FindField(ChpDeb).asFloat - Q.FindField(ChpCre).asFloat ;
      END ;

    if SwapSelect then  // en mode inversé, on retranche le cumul des lignes sélectionnée du solde total
    begin
      debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_' + ChpDeb)).text ) ;
      credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_' + ChpCre)).text ) ;
      Result := (debit - credit) - Result ;
    end ;
  end ;
end;

(*procedure TOF_CPGENEREMP.ClickModifRib;
var RJal,RExo : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
{$IFDEF EAGLCLIENT}
  if ModifRibSurMul( Q.TQ, FTIC.Checked, TRUE ) then
{$ELSE}
  if ModifRibSurMul( Q, FTIC.Checked, TRUE ) then
{$ENDIF}

    begin
//    Application.ProcessMessages ;
    TFMul(Ecran).BChercheClick(Nil);
{$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.row - 1) ;
{$ENDIF}
    RJal      := Q.FindField('E_JOURNAL').AsString ;
    RExo      := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
    RDate     := Q.FindField('E_DATECOMPTABLE').AsDateTime ;
    RNumP     := Q.FindField('E_NUMEROPIECE').AsInteger ;
    RNumL     := Q.FindField('E_NUMLIGNE').AsInteger ;
    RNumEche  := Q.FindField('E_NUMECHE').AsInteger ;
{$IFDEF EAGLCLIENT}
    Q.TQ.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
            VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[]) ;
{$ELSE}
    Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
            VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[]) ;
{$ENDIF}
    end ;
end;*)

procedure TOF_CPGENEREMP.InitContexte;
begin
  InitE_NUMTRACHQ( smp in [smpEncPreEdt,smpEncTous,smpDecVirEdt,smpDecVirBqe,smpDecVirInEdt,smpDecVirInBqe,smpEncPreBqe] ) ;
  SetControlProperty('E_NOMLOT',  'PLUS', AttribPlus(smp) );
  SetControlProperty('E_NOMLOT_', 'PLUS', AttribPlus(smp) );
  Case smp of
   smpEncDiv    : begin // Encaissement divers
                  Ecran.Caption       := HDiv[20] ;
                  end ;
   smpEncPreEdt : begin // Prélèvements clients
                  Ecran.Caption   := HDiv[19] ;
                  end ;
   smpEncPreEdtNC : begin // Prélèvements clients
                    Ecran.Caption   := HDiv[37] ;
                    end ;
   smpEncPreBqe : begin // Prélèvements clients
                  Ecran.Caption   := HDiv[38] ;
                  end ;
   smpEncChqPor : begin // Mise en portefeuille chèques clients
                  Ecran.Caption   := HDiv[28] ;
                  end ;
    smpEncCBPor : begin // Mise en portefeuille cartes bleues clients
                  Ecran.Caption   := HDiv[16] ;
                  end ;
   smpEncTraPor : begin // Mise en portefeuille traite clients
                  Ecran.Caption   := HDiv[16] ;
                  end ;
   smpEncTraEdt : begin // Lettres-traite clients
                  Ecran.Caption   := HDiv[5] ;
                  end ;
 smpEncTraEdtNC : begin // Lettres-traite clients
                  Ecran.Caption   := HDiv[26] ;
                  end ;
   smpEncTraEnc : begin // Traite clients à l'encaissement
                  Ecran.Caption   := HDiv[8] ;
                  end ;
   smpEncTraEsc : begin // Traite clients à l'escompte
                  Ecran.Caption   := HDiv[9] ;
                  end ;
   smpEncChqBqe : begin // Chèque clients en banque
                  Ecran.Caption   := HDiv[31] ;
                  end ;
    smpEncCBBqe : begin // CB clients en banque
                  Ecran.Caption   := HDiv[6] ;
                  end ;
   smpEncTraBqe : begin // Traite clients en banque
                  Ecran.Caption   := HDiv[6] ;
                  end ;
     smpEncTous : begin // Prélèvements clients
                  Ecran.Caption   := HDiv[12] ;
                  end ;
   smpDecChqEdt : begin // Lettres-chèques fournisseurs
                  Ecran.Caption   := HDiv[2] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[25] ) ;
                  end ;
 smpDecChqEdtNC : begin // Lettres-chèques fournisseurs
                  Ecran.Caption   := HDiv[32] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[25] ) ;
                  end ;
   smpDecVirEdt : begin // Virements fournisseurs
                  Ecran.Caption     := HDiv[18] ;
                  end ;
 smpDecVirEdtNC : begin // Virements fournisseurs
                  Ecran.Caption     := HDiv[33] ;
                  end ;
   smpDecVirBqe : begin // Virements fournisseurs
                  Ecran.Caption     := HDiv[3] ;
                  end ;
 smpDecVirInEdt : begin // Virements internationales fournisseurs
                  Ecran.Caption     := HDiv[36] ;
                  end ;
 smpDecVirInEdtNC : begin // Virements internationales fournisseurs
                  Ecran.Caption     := HDiv[35] ;
                  end ;
 smpDecVirInBqe : begin // Virements internationales fournisseurs
                  Ecran.Caption     := HDiv[34] ;
                  end ;
   smpDecBorEdt : begin // Lettres-Bor fournisseurs
                  Ecran.Caption   := HDiv[27] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  end ;
 smpDecBorEdtNC : begin // Lettres-Bor fournisseurs
                  Ecran.Caption   := HDiv[4] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  end ;
   smpDecborDec : begin // Traites fournisseurs en encaissement
                  Ecran.Caption   := HDiv[10] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  end ;
   smpDecBorEsc : begin // Traites fournisseurs à l'escompte
                  Ecran.Caption   := HDiv[11] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  end ;
   smpDecTraBqe : begin // Traites fournisseurs en banque
                  Ecran.Caption   := HDiv[7] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  end ;
   smpDecTraPor : begin // Mise en portefeuille traite (clients ??)
                  Ecran.Caption   := HDiv[17] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  end ;
      smpDecDiv : begin // Décaissements divers
                  Ecran.Caption       := HDiv[21] ;
                  end ;
     smpDecTous : begin //
                  Ecran.Caption   := HDiv[13] ;
                  end ;
   end ;

  if isEnc(smp) then
    begin
    SetControlProperty('E_GENERAL','DataType','TZGENCAIS') ;
    SetControlProperty('E_AUXILIAIRE','DataType','TZTTOUTDEBIT') ;
    end
  else
    begin
    SetControlProperty( 'E_GENERAL',     'DataType', 'TZGDECAIS'     ) ;
    SetControlProperty( 'E_AUXILIAIRE',  'DataType', 'TZTTOUTCREDIT' ) ;
    cFactCredit.Caption := HDiv[15] ;
    end ;

  gszCaption  := Ecran.Caption;  // Pour les messages
  UpdateCaption(Ecran) ;

  InitCATEGORIE ;

end;

procedure TOF_CPGENEREMP.InitCriteres;
var i       : integer ;
begin

  if ((smp<>smpEncTous) and (smp<>smpDecTous)) then
    begin
    // Param onglet tables libres
    LibellesTableLibre( TTabSheet(GetControl('PLibres', True)), 'TT_TABLE', 'T_TABLE', 'T') ;
    // Exo et date comptable
    if VH^.CPExoRef.Code<>'' then
      begin
      SetControlText('E_EXERCICE',        VH^.CPExoRef.Code ) ;
      SetControlText('E_DATECOMPTABLE',   DateToStr(VH^.CPExoRef.Deb) ) ;
      SetControlText('E_DATECOMPTABLE_',  DateToStr(VH^.CPExoRef.Fin) ) ;
      end
    else
      begin
      SetControlText('E_EXERCICE',        VH^.Entree.Code ) ;
      SetControlText('E_DATECOMPTABLE',   DateToStr(VH^.Entree.Deb) ) ;
      SetControlText('E_DATECOMPTABLE_',  DateToStr(VH^.Entree.Fin) ) ;
      end ;
    // Dates d'échéances
    SetControlText('E_DATEECHEANCE',    StDate1900 ) ;
    SetControlText('E_DATEECHEANCE_',   StDate2099 ) ;
    SetControlText('E_DEVISE' ,         V_PGI.DevisePivot ) ;
    // Gestion spooler
    TTabSheet(GetControl('PGestionPDF')).TabVisible := smp in
                                     [smpEncPreEdt,smpEncTraEdt,smpEncTraEdtNC,smpEncPreEdtNC,
                                      smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,
                                      smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC,
                                      smpDecBorEdt,smpDecBorEdtNC] ;
    // pas de gestion par lot
    TTabSheet(GetControl('PBanqueLot')).TabVisible := False ;
    Pages.ActivePage := TTabSheet(GetControl('PCritere')) ;

    if (smp=smpDecChqEdtNC) or (smp=smpDecVirInEdtNC) or (smp=smpDecVirEdtNC) or (smp=smpEncTraEdtNC) or (smp=smpDecBorEdtNC) then begin // VL 07/03/05 FQ 13746
      SetControlVisible('CGenereAuto',     False);
      SetControlChecked('CGenereAuto',     False) ;
    end;
    end
  else
    begin
    Pages.ActivePage := TTabSheet(GetControl('PBanqueLot')) ;
    for i:=0 to Pages.PageCount - 1 do
      if Pages.Pages[i].Name<>'PBanqueLot' then
        Pages.Pages[i].TabVisible := False ;
    end ;

  // Gestion n° Traite
  if (smp<>smpEncTraEdt) And (smp<>smpEncTraEdtNC) And
     (smp<>smpDecBorEdt) And (smp<>smpDecBorEdtNC) Then
    begin
    TraiteEdite.State := cbGrayed ;
    SetControlVisible('E_NUMTRAITECHQ_', FALSE) ;
    SetControlEnabled('TE_NUMTRAITECHQ_', FALSE) ;
    SetControlEnabled('E_NUMTRAITECHQ', FALSE) ;
    SetControlEnabled('TE_NUMTRAITECHQ', FALSE) ;
    SetControlText('E_NUMTRAITECHQ_', '') ;
    SetControlText('E_NUMTRAITECHQ', '') ;
    end ;
  TraiteEditeChange(Nil) ;

  // Nature de pièce
  THMultiValComboBox(GetControl('_E_NATUREPIECE', True)).SelectAll ;

  // MAJ des zones des comptes/taux tva pour escompte
  ESCCBTVAChange(nil);

end;

procedure TOF_CPGENEREMP.InitE_NUMTRACHQ( Invisible : Boolean );
begin
  SetControlVisible('E_NUMTRAITECHQ_',  Not Invisible ) ;
  SetControlVisible('TE_NUMTRAITECHQ_', Not Invisible ) ;
  SetControlVisible('E_NUMTRAITECHQ',   Not Invisible ) ;
  SetControlVisible('TE_NUMTRAITECHQ',  Not Invisible ) ;
end;

procedure TOF_CPGENEREMP.MarqueOrigine;
var TOBL : TOB ;
begin
	TOBL:=TOBOrig.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
  	                      [Q.FindField('E_JOURNAL').AsString,Q.FindField('E_EXERCICE').AsString,
    	                     Q.FindField('E_DATECOMPTABLE').AsDateTime,Q.FindField('E_NUMEROPIECE').AsInteger,
      	                   Q.FindField('E_NUMLIGNE').AsInteger,Q.FindField('E_NUMECHE').AsInteger],False) ;
	if TOBL<>Nil then
  	begin
    TOBL.PutValue('MARQUE','X') ;
    // Gestion escompte fournisseur à la ligne
    if (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC]) then
    	begin
      // Recherche TobEscompte, Création si non trouvé
  	  if not GereEscompte then Exit ;
      If Not AvecEscompte Then Exit ;
		  FindCreerTOBEscomptePourLigne( TOBEscomptes, TOBL, False ) ;
      end ;
    end ;
end;

procedure TOF_CPGENEREMP.PrechargeOrigines;
Var QQ : TQuery ;
    st : String ;
begin
  TOBOrig.ClearDetail ;
  St:='SELECT ECRITURE.*, G_NATUREGENE,T_NATUREAUXI,T_ESCOMPTE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '+
      'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
  QQ:=OpenSQL(St+RecupWhereCritere(Pages),True) ;
  TOBOrig.LoadDetailDB('ECRITURE','','',QQ,False,True) ;
  Ferme(QQ) ;
  if TOBOrig.Detail.Count>0 then
    TOBOrig.Detail[0].AddChampSup('MARQUE',True) ;
end;

procedure TOF_CPGENEREMP.PreciseMP;
var E_MODEPAIE : THValComboBox ;
begin
  E_MODEPAIE := THValComboBox(GetControl('E_MODEPAIE',True)) ;
  if Categorie.itemIndex = 0
    then CatToMP ( '',              E_MODEPAIE.Items, E_MODEPAIE.Values, tslAucun, True)
    else CatToMP ( CATEGORIE.Value, E_MODEPAIE.Items, E_MODEPAIE.Values, tslAucun, True ) ;
  E_MODEPAIE.itemIndex := 0 ;
end;

procedure TOF_CPGENEREMP.ReinitWhereNatCpt( RAZ : Boolean ) ;
begin
  if RAZ then
    SetControlText('XX_WHERENATCPT','')
  else
    begin
    if FTIC.Checked then
      begin
      if IsEnc(SMP)
        then SetControlText('XX_WHERENATCPT','E_AUXILIAIRE="" AND G_NATUREGENE<>"TIC"')
        else SetControlText('XX_WHERENATCPT','E_AUXILIAIRE="" AND G_NATUREGENE<>"TID"');
      end
    else
      begin
      if IsEnc(SMP)
        then SetControlText('XX_WHERENATCPT','E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"FOU" AND T_NATUREAUXI<>"AUC"')
        else SetControlText('XX_WHERENATCPT','E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"CLI" AND T_NATUREAUXI<>"AUD"');
      end ;
    end ;
end;

procedure TOF_CPGENEREMP.ToutMarquer;
var i : integer ;
    TOBL : TOB ;
begin
  for i:=0 to TOBORig.Detail.Count-1 do
    BEGIN
    TOBL:=TOBOrig.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    // Gestion escompte fournisseur à la ligne
    if (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC]) then
    	begin
      // Recherche TobEscompte, Création si non trouvé
  	  if not GereEscompte then Exit ;
      If Not AvecEscompte Then Exit ;
		  FindCreerTOBEscomptePourLigne( TOBEscomptes, TOBL, False ) ;
      end ;
    end ;
end;

procedure TOF_CPGENEREMP.CATEGORIEChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  // Le 28/01/2003 BPY correction des bug de la fiche 11794
  if Sender <> nil then
    SetControlVisible('BCtrlRib', not (THValComboBox(Sender).Value = 'TRI') ) ;
  // Fin BPY
  {JP 18/02/05 :  Gestion de l'affichage des contrôles E_BANQUEPREVI}
  AfficheBanquePrevi(False);
  PreciseMP ;
end;

procedure TOF_CPGENEREMP.E_EXERCICEChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  if GetControlText('E_EXERCICE') = '' then
    begin
    SetControlText('E_DATECOMPTABLE',     StDate1900) ;
    SetControlText('E_DATECOMPTABLE_',    StDate2099) ;
    end
  else
    ExoToDates( GetControlText('E_EXERCICE') , THEdit(GetControl('E_DATECOMPTABLE', True)) ,
                                               THEdit(GetControl('E_DATECOMPTABLE_', True))  ) ;
end;
  
procedure TOF_CPGENEREMP.FTICChange(Sender: TObject);
begin

  TFMul(Ecran).CritModified := True;

  { YMO 01/06/06 FQ17447 Attention on ne vide pas ici lors du chargement du filtre à l'ouverture du MUL,
   car on passe ici 3 fois, dont 2 en dehors du chargement du filtre ;
   'InChargement' n'est donc pas suffisant}
  If Not InChargement and FinChargeEcran then SetControlText('E_GENERAL', '') ;

  if FTIC.Checked then
    begin
    SetControlEnabled('E_AUXILIAIRE', FALSE ) ;
    SetControlEnabled('TE_AUXILIAIRE', FALSE ) ;
    SetControlText('E_AUXILIAIRE', '' ) ;
    if IsEnc(smp)
      then SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGTID')
      else SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGTIC') ;
    end
  else
    begin
    SetControlEnabled('E_AUXILIAIRE', TRUE ) ;
    SetControlEnabled('TE_AUXILIAIRE', TRUE ) ;
    if IsEnc(smp)
      then SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGENCAIS')
      else SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGDECAIS') ;
    end ;

end;

procedure TOF_CPGENEREMP.cFactCreditChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  if isEncMP(smp) then
    begin
    if cFactCredit.State=cbGrayed
      then SetControlText('_E_NATUREPIECE', GetControlText('_E_NATUREPIECE') + 'AC;OC;FC;' )
      else if cFactCredit.State=cbUnChecked
        then Enleve('AC')
        else  // cbChecked
          begin
          SetControlText('_E_NATUREPIECE', 'AC;OC;FC;' ) ;
          _chk_RegSup.State     := cbUnChecked;
          end;
    end
  else
    begin
    if cFactCredit.State = cbGrayed
      then SetControlText('_E_NATUREPIECE', 'AF;OF;FF;' )
      else if cFactCredit.State=cbUnChecked
        then Enleve('AF')
        else  // cbChecked
          begin
          SetControlText('_E_NATUREPIECE', 'AF;OF;FF;' ) ;
          _chk_RegSup.State := cbUnChecked ;
          end;
    end;
end;

procedure TOF_CPGENEREMP.TraiteEditeChange(Sender: TObject);
var lBoEtat : Boolean ;
begin
  TFMul(Ecran).CritModified := True;
  lBoEtat := TraiteEdite.State=cbChecked ;
  SetControlEnabled('E_NUMTRAITECHQ_', lBoEtat ) ;
  SetControlEnabled('FE_NUMTRAITECHQ_', lBoEtat ) ;
  SetControlEnabled('E_NUMTRAITECHQ', lBoEtat ) ;
  SetControlEnabled('FE_NUMTRAITECHQ', lBoEtat ) ;

  SetControlText('E_NUMTRAITECHQ_', '' ) ;
  SetControlText('E_NUMTRAITECHQ', '' ) ;
end;

procedure TOF_CPGENEREMP.BCtrlRibClick(Sender: TObject);
Var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(Pages) ;
  if (StWRib = '') then Exit;
  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
  if ({$IFDEF EAGLCLIENT}    // FQ18935 SBO 07/11/2006
      // pas de tout sélectionner en CWAS
      {$ELSE}
      (Not FListe.AllSelected) and
      {$ENDIF}
      (FListe.NbSelected>0) and (FListe.NbSelected<100)) then
    begin
    // Rajoute une clause au WHERE
    StWRib := StWRib+' AND (';
    for i:=0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookmark(i) ;
      {$IFDEF EAGLCLIENT}    // FQ18935 SBO 07/11/2006
        Q.TQ.Seek(FListe.row - 1) ;
      {$ENDIF}
      StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
      end;
    // Efface le dernier OR et rajoute ')'
    StWRib := Copy( StWRib, 1, length(StWRib) - 2 );
    StWRib := StWRib +')';
    end;

//  If StWRib<>'' Then CPLanceFiche_VerifRib('WHERE='+StWRib);
  if (Categorie.Value = 'TRI')
    then CPLanceFiche_VerifRib('IBAN;WHERE='+StWRib)
    else CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

{JP 01/06/07 : FQ 20308 : Réimplémentation de la fonction pour la rendre compatible 2/3 - eAgl
               Saisie pièce et saisie bordereau.
{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREMP.FListeDblClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
{$IFNDEF IMP}
var
  sMode : string;
  lMulQ : TQuery;
  {$ENDIF}
begin
  {$IFNDEF IMP}
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;
  {$IFDEF EAGLCLIENT}
    lMulQ := TFMul(Ecran).Q.TQ;
    lMulQ.Seek(TFMul(Ecran).FListe.Row - 1);
  {$ELSE}
    lMulQ := TFMul(Ecran).Q;
  {$ENDIF}

  sMode := Q.FindField('E_MODESAISIE').AsString ;
  if ((sMode<>'') and (sMode<>'-')) then LanceSaisieFolio(lMulQ,taConsult)
                                    else TrouveEtLanceSaisie(lMulQ,taConsult,'N');
  {$ENDIF}
end;

procedure TOF_CPGENEREMP.BOuvrirClick(Sender: TObject);
Var MM                  : RMVT ;
    TOBE,TOBO2,TOBEsc   : TOB ;
    St,sBqe,sCat,OldBqe : String ;
    i,k                 : integer ;
    PasBon,OkJal,PbJal  : boolean ;
    CarePDF             : Boolean ;
    ShunteParamMP       : Boolean ;
    IsTicTid : Boolean;
    SvgGeneral : string;
    bReeditionTraite : Boolean;
begin
  if TFmul(Ecran).CritModified then begin // FQ 12480
    PgiInfo('Vous n''avez pas appliqué les critères.');
    exit;
  end;

  TOBEsc := nil;
// Le résultat de la recherche est-il vide ?
{$IFDEF EAGLCLIENT}
	if (Q.TQ.Detail.Count = 0) then
{$ELSE}
	if (Q.Eof and Q.Bof) then
{$ENDIF}
    begin
    PGIInfo('Le résultat de la recherche est vide. Veuillez relancer une recherche.',gszCaption) ;
    exit ;
    end ;
// Au moins une ligne de sélectionnée en mode normale, (pas de test en mode inverse pour l'instant)
	if ( FListe.NbSelected = 0 )
    and ( (not FListe.AllSelected) and (not swapSelect) )
    then
      begin
      PGIInfo('Aucune ligne à traiter. Vous devez sélectionner au moins une écriture.',gszCaption) ;
      exit ;
      end ;
// Au plus 4999 lignes de sélecionnées
	if ( YEnATrop ) then
      begin
      PGIInfo( GetYenATropMsg(smp) ,gszCaption) ;
      exit ;
      end ;
// Test et avertissement en cas d'escompte
  if (GetCheckBoxState('FEscompte') = cbChecked) and AvecEscompte then
  	begin
    // Blocage sur le compte/taux d'escompte.
    if ((GetControlText('ESCHT')='') or (not IsNumeric(GetControlText('ESCTAUXESC'))) or (Valeur(GetControlText('ESCTAUXESC'))=0)) then
      begin
    	PGIBox('Vous n''avez pas renseigné le compte et/ou le taux d''escompte.',gszCaption);
      exit;
      end ;
    // Blocage sur le compte/taux de TVA Si Coche "TVA sur escompte" cochée.
    if (GetCheckBoxState('ESCCBTVA') = cbChecked) and ((GetControlText('ESCTVA')='') or (not IsNumeric(GetControlText('ESCTAUXTVA'))) or (Valeur(GetControlText('ESCTAUXTVA'))=0)) then
      begin
    	PGIBox('Vous n''avez pas renseigné le compte et/ou le taux de TVA.',gszCaption);
      exit;
      end ;
    end ;
// Message avertissement pour sélection inversée
	if swapSelect then
  	if PGIAsk('Vous êtes en mode "sélection inversée", Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes
    then  exit;

  // Réédition des traites FQ 12480
  bReeditionTraite := False;
  if (smp in [smpEncTRAEDTNC, smpDecBorEdtNC]) and (TraiteEdite.State = cbChecked) then
    bReeditionTraite := (HShowMessage(HM[11],gszCaption,'') = mrYes);

  CarePdf   := FALSE ;

  // MAJ des paramètres de génération
  ParmsMP.DEV.Code  := GetControlText('E_DEVISE') ;
  GetInfosDevise(ParmsMP.DEV) ;
  // BPY le 21/01/2004 => autre bug soulevé par le point 109 de la FFF
  SvgGeneral := GetControlText('E_GENERAL');
  if (not (SvgGeneral = '')) then
  begin
    IsTicTid := ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + SvgGeneral + '" AND G_NATUREGENE IN ("TIC","TID")');
//    if ((IsTicTid) and (not (GetCheckBoxState('FTIC') = cbChecked))) then if (PGIAsk('Le compte sélectionner est un compte tiers créditeur ou débiteur.' + #10 + 'Vous devriez indiquer cette spécificité dans votre sélection.' + #10 + 'Voulez vous continuer ?',gszCaption) <> mrYes) then exit;
//    if ((not IsTicTid) and (GetCheckBoxState('FTIC') = cbChecked)) then if (PGIAsk('Le compte sélectionner n''est pas un compte tiers créditeur ou débiteur.' + #10 + 'Vous devriez indiquer cette spécificité dans votre sélection.' + #10 + 'Voulez vous continuer ?',gszCaption) <> mrYes) then exit;
//    SetControlChecked('FTIC',IsTicTid);
//    SetControlText('E_GENERAL',SvgGeneral);
    if ((IsTicTid) and (not (GetCheckBoxState('FTIC') = cbChecked))) then
    begin
        PGIBox('Votre sélection contient des TIC/TID : veuillez indiquer cette spécificité dans votre sélection.');
        lMenuItem.Click; //BNouvRechClick(nil); FQ 14212
{$IFDEF EAGLCLIENT}
        FListe.VidePile(true);
{$ENDIF}
        exit;
    end;
    if ((not IsTicTid) and (GetCheckBoxState('FTIC') = cbChecked)) then
    begin
        PGIBox('Votre sélection ne contient aucun TIC/TID : veuillez indiquer cette spécificité dans votre sélection.');
        lMenuItem.Click; //BNouvRechClick(nil); FQ 14212
{$IFDEF EAGLCLIENT}
        FListe.VidePile(true);
{$ENDIF}
        exit;
    end;
  end;
  // Fin BPY

  {JP 18/02/05 : On renseigne la banque prévisionnelle en Enca/Deca
   JP 09/08/05 : FQ 15815 : gestion de la banque prévisionnelle dans la remise en banque des BOR (smpDecBorDec)
   JP 07/06/07 : FQ 15793 : Passage à un MultiValComboBox}
  if (ParmsMP.smp in [smpEncDiv, smpDecDiv, smpDecBorDec]) then begin
    St := GetControlText('BANQUEPREVI');
    if Pos('<<', St) = 0 then
      ParmsMP.BanquePrevi := ReadTokenSt(St);
    St := '';
  end;

  ParmsMP.TIDTIC    := (GetCheckBoxState('FTIC') = cbChecked) ;
  ShunteParamMP := smp In [smpEncTraEdtNC,smpDecBorEdtNC] ;
  if ShunteparamMP Then
    ShunteparamMP := ParamsMPAuto(ParmsMP) ;
  // Si l'affectation auto n'a pas fonctionné : Ouvre la fenêtre
  if not ShunteParamMP Then
    if Not ParamsMP(True,ParmsMP) then Exit ;

  // MAJ des conditions de ruptures --> TOUJOURS E_QUALIFPIECE !!!
  // OR E_QUALIFPIECE FIXE a N PAR LES CRITERES DE SELECTION
  // DONC NE SERT A RIEN !!!!
  // DE PLUS ParmsMP.ChampsRupt TOUJOURS A VIDE !!!!!!!
{
  if ((smp in [smpEncTous,smpDecTous]) and (ParmsMP.ChampsRupt='')) then
    ParmsMP.ChampsRupt := 'E_QUALIFPIECE;' ;
  if Pos('E_QUALIFPIECE',ParmsMP.ChampsRupt)<=0 then
    ParmsMP.ChampsRupt := 'E_QUALIFPIECE;' + ParmsMP.ChampsRupt ;
}
  ParmsMP.ChampsRupt := 'E_QUALIFPIECE;' ;

  if (GetCheckBoxState('FEscompte') = cbChecked) And AvecEscompte then
    if HShowMessage(HM[9],gszCaption,'')<>mrYes then Exit ;

//  inherited; // ???

  if Not ConstitueOrigines then Exit ;
  OldBqe := '' ;
  if TOBOrig.Detail.Count<=0 then Exit ;

  PasBon  := False ;
  PbJal   := FALSE ;
  if (GetCheckBoxState('Spooler') = cbChecked) and (not V_PGI.QRPDF) then
    begin
    V_PGI.QRPDF := TRUE ;
    CarePdf     := TRUE ;
    end ;
  if smp in [smpEncTous,smpDecTous] then
    begin
    for i:=0 to LesCombi.Count-1 do
      begin
      St   := LesCombi[i] ;
      sBqe := ReadtokenSt(St) ;
      sCat := ReadtokenSt(St) ;
      // Tripoter ParmsMP
      ParmsMP.CptG  := sBqe ;
      OkJal         := TRUE ;
      ParmsMP.JalG  := CompteToJal(sBqe,OkJal) ;
      If Not OkJal Then
        begin
        PbJal := TRUE ;
        Continue ;
        end ;
      if i=0 then
        OldBqe := sBqe ;
      if OldBqe<>sBqe then
        IncNumLot ;
      for k:=0 to LesCritsRupt.Count-1 do
        begin
        ConstitueDest( sBQE, sCat, LesCritsRupt[k], TOBOrig, TOBDest, ParmsMP ) ;
        if VerifEcheMP( ParmsMP.AlerteEcheMP ) then
          begin
          if LInsertDB(TOBDest) then
            begin
            TOBE   := TOBDest.Detail[0] ;
            MM     := TOBToIdent(TOBE,False) ;
            CompleteMM ( ParmsMP, SorteLettre,
                         GetCheckBoxState('Spooler') = cbChecked,
                         GetCheckBoxState('XFichierSpooler') = cbChecked,
                         GetControlText('RepSpooler'),
                         MM ) ;
            TOBO2  := ExtractDiscerne(TOBOrig,sBqe,sCat,LesCritsRupt[k],ParmsMP) ;
            if GereEscompte then
              TOBEsc := GenereEscompteMP( TOBO2, TOBEscomptes, TOBGHT, TOBGTVA,
                                          GetCheckBoxState('ESCCBTVA') = cbChecked,
                                          JalEsc , ParmsMP ) ;
            if TOBO2<>Nil then
              begin
              if LanceSaisieMP( MM, TOBO2, TOBEsc,
                                GetCheckBoxState('CGenereAuto') = cbChecked,
                                ParmsMP.LettrAuto,
                                TOBEscomptes, bReeditionTraite, ParmsMP.ModeleBordereau) then
                if GereEscompte then
                  FlagOrigEsc(TOBO2) ;
              TOBO2.Free ;
              END ;
            TOBEsc.Free ;
            END ;
          END
        else
          BEGIN
          Break ;
          PasBon:=True ;
          END ;
        END ; {boucle k}
      OldBqe:=sBqe ;
      if PasBon then Break ;
      END ; {Boucle i}
    LesCombi.Clear ;
    LesCritsRupt.Clear ;
    END
  else
    if LesCritsRupt.Count>0 then
      BEGIN
      CreerCodeLot(Copy(ParmsMP.NumEncaDeca,1,17)); // Cré le code lot : Demande FFF
      ParmsMP.DEV.Code := GetControlText('E_DEVISE') ;
      GetInfosDevise( ParmsMP.DEV ) ;
      ParmsMP.CptS     := GetControlText('E_GENERAL') ;
      for k:=0 to LesCritsRupt.Count-1 do
        BEGIN
        ConstitueDest( sBQE, sCat, LesCritsRupt[k], TOBOrig, TOBDest, ParmsMP ) ;
        if VerifEcheMP( ParmsMP.AlerteEcheMP ) then
          BEGIN
          if LInsertDB(TOBDest) then
            BEGIN
            TOBE   := TOBDest.Detail[0] ;
            MM     := TOBToIdent(TOBE,False) ;
            CompleteMM ( ParmsMP, SorteLettre,
                         GetCheckBoxState('Spooler') = cbChecked,
                         GetCheckBoxState('XFichierSpooler') = cbChecked,
                         GetControlText('RepSpooler'),
                         MM ) ;
            TOBO2  := ExtractDiscerne(TOBOrig,'','',LesCritsRupt[k],ParmsMP) ;

            if GereEscompte then
              TOBEsc := GenereEscompteMP( TOBO2, TOBEscomptes, TOBGHT, TOBGTVA,
                                          GetCheckBoxState('ESCCBTVA') = cbChecked,
                                          JalEsc , ParmsMP ) ;
            if TOBO2<>Nil then
              BEGIN
              if LanceSaisieMP( MM, TOBO2, TOBEsc,
                                GetCheckBoxState('CGenereAuto') = cbChecked,
                                ParmsMP.LettrAuto,
                                TOBEscomptes, bReeditionTraite, ParmsMP.ModeleBordereau) then
                if GereEscompte then
                  FlagOrigEsc(TOBO2) ;
              TOBO2.Free ;
              END ;
            TOBEsc.Free ;
            END ;
          END
        else
          Break ;
        END ; {boucle k}
    LesCritsRupt.Clear ;
    END
  else
    BEGIN
    ParmsMP.DEV.Code := GetControlText('E_DEVISE') ;
    GetInfosDevise(ParmsMP.DEV) ;
    ParmsMP.CptS     := GetControlText('E_GENERAL') ;
    ConstitueDest( '', '', '', TOBOrig, TOBDest, ParmsMP ) ;
    if VerifEcheMP( ParmsMP.AlerteEcheMP ) then
      BEGIN
      if LInsertDB(TOBDest) then
        BEGIN
        TOBE  := TOBDest.Detail[0] ;
        MM    := TOBToIdent(TOBE,False) ;
        CompleteMM ( ParmsMP, SorteLettre,
                     GetCheckBoxState('Spooler') = cbChecked,
                     GetCheckBoxState('XFichierSpooler') = cbChecked,
                     GetControlText('RepSpooler'),
                     MM ) ;
        LanceSaisieMP( MM, TOBOrig, Nil, False, ParmsMP.LettrAuto, TOBEscomptes, bReeditionTraite, ParmsMP.ModeleBordereau);
        If MM.MSED.Spooler Then
          Case smp Of
            smpDecChqEdt,smpDecChqEdtNC : SetparamSoc('SO_CPCHEMINCHEQUE',    MM.MSED.RepSpooler) ;
            smpEncTraEdt,smpEncTraEdtNC : SetparamSoc('SO_CPCHEMINTRAITE',    MM.MSED.RepSpooler) ;
            smpDecBorEdt,smpDecBorEdtNC : SetparamSoc('SO_CPCHEMINBOR',       MM.MSED.RepSpooler) ;
            smpDecVirEdt,smpDecVirEdtNC : SetparamSoc('SO_CPCHEMINVIREMENT',  MM.MSED.RepSpooler) ;
            smpDecVirInEdt,smpDecVirInEdtNC : SetparamSoc('SO_CHEMINVIRIN',MM.MSED.RepSpooler) ; // VL 301003 FQ 12958
            smpEncPreEdt,smpEncPreEdtNC : SetparamSoc('SO_CPCHEMINPRELEVEMENT', MM.MSED.RepSpooler) ;
            End ;
        END ;
      END ;
    END ;
  If PbJal Then
    BEGIN
    HShowMessage(HM[5],gszCaption,'') ;
    HShowMessage(HM[6],gszCaption,'') ;
    END ;

    // YMO 17/05/2006 FQ17523 Pas de regul lettrage si 'case Lettrage auto' décochée dans les paramètres
    If ParmsMP.LettrAuto Then
// bpy le 21/07/2004 => Fiche 12340 : gestion des ecart de change !
    for i:=0 to TOBOrig.Detail.Count-1 do
    begin
        if (not (TOBOrig.Detail[i].GetValue('E_DEVISE') = V_PGI.DevisePivot)) then
        begin
            PGIBox('Attention certaines écritures sont en devise.' + #10 + #13 + 'Vous devez comptabiliser les différences de change.');
            RegulLettrageMP(False,False,prCLient);
            break;
        end;
    end
// Fin BPY
    else
      PGIBox('Ces écritures de règlement ont été générées sans lettrage.');


  If CarePDF Then
    V_PGI.QRPDF := FALSE ;
  TFMul(Ecran).BChercheClick(Nil) ;
end;

{$IFDEF EAGLCLIENT}
{JP 10/05/07 : FQ 19608 : gestion de Carré rouge pour les RIB
{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREMP.FListeDrawCell(Sender: TObject; ACol, ARow: Longint; ARect: TRect; AState: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  sRIB, St : String ;
  TOBE     : TOB ;
  TheColor : TColor ;
  DD : TDateTime ;
  Text   : array[0..255] of Char;
  F      : TAlignment ;
  Grille : THgrid ;
  N, K   : Integer;
begin
  Grille := THgrid(Sender);
  StrPCopy(Text,Grille.Cells[ACol, ARow]);
  Grille.Canvas.Font.Style := Grille.Canvas.Font.Style - [fsItalic];

  if FTic.Checked then St := 'E_GENERAL'
                  else St := 'E_AUXILIAIRE';

  if (gdFixed in AState) then begin
    Grille.Canvas.Brush.Color := Grille.FixedColor;
    Grille.Canvas.Font.Color  := Grille.Font.Color;
    F := taCenter;
  end
  else begin
    F := Grille.ColAligns[ACol];
    {JP 29/05/07 : FQ 20441 : j'avais oublié de redessiner les lignes sélectionnées}
    if Grille.IsSelected(ARow) then begin
      Grille.Canvas.Brush.Color := clHighlight;
      Grille.Canvas.Font.Color  := clHighlightText;
      Grille.Canvas.Font.Style  := [fsItalic] ;
    end
    else if (gdSelected in AState) then begin
      Grille.Canvas.Brush.Color := clHighlight;
      Grille.Canvas.Font.Color  := clHighlightText;
    end
    else begin
      if ((ARow mod 2) = 0) and (V_PGI.NumAltCol <> 0)then
        Grille.Canvas.Brush.Color:= AltColors[V_PGI.NumAltCol]
      else
        Grille.Canvas.Brush.Color:=clWindow;
      Grille.Canvas.Font.Color:=clWindowText;
    end;
  end;

  case F of
    taRightJustify : ExtTextOut(Grille.Canvas.Handle, ARect.Right - Grille.Canvas.TextWidth(Grille.Cells[ACol,ARow])-3,
                               ARect.Top + 2,ETO_OPAQUE or ETO_CLIPPED, @ARect, Text, StrLen(Text), nil);
    taCenter       : ExtTextOut(Grille.Canvas.Handle,ARect.Left + ((ARect.Right-ARect.Left-Grille.canvas.TextWidth(Grille.Cells[ACol,ARow])) div 2),
                         ARect.Top + 2, ETO_OPAQUE or ETO_CLIPPED, @ARect, Text, StrLen(Text), nil);
    else
      ExtTextOut(Grille.Canvas.Handle, ARect.Left + 2, ARect.Top + 2, ETO_OPAQUE or ETO_CLIPPED,
                 @ARect, Text, StrLen(Text), nil);
  end;

  if ((gdfixed in AState) and Grille.Ctl3D) then begin
    DrawEdge(Grille.Canvas.Handle, ARect, BDR_RAISEDINNER, BF_BOTTOMRIGHT);
    DrawEdge(Grille.Canvas.Handle, ARect, BDR_RAISEDINNER, BF_TOPLEFT);
  end;

  if ARow = 0 then Exit;

  {Gestion des RIB}
  if (Grille.ColNames[ACol] = St) then begin
    K := -1;
    for n := 0 to Grille.ColCount - 1 do
      if Grille.ColNames[n] = 'E_RIB' then begin
        K := n;
        Break;
      end;
    if K >= 0 then begin
      sRIB := Grille.Cells[K, ARow];
      if sRib = '' then begin
        Grille.Canvas.Brush.Color := clRed;
        Grille.Canvas.Brush.Style := bsSolid;
        Grille.Canvas.Pen.Color   := clRed;
        Grille.Canvas.Pen.Mode    := pmCopy;
        Grille.Canvas.Pen.Width   := 1;
        Grille.Canvas.Rectangle(ARect.Right-5,ARect.Top+1,ARect.Right-1,ARect.Top+5);
      end;
    end;
  end;

  {pour les Chèques fournisseurs : Escompte par ligne}
  if (FListe.ColNames[ACol]='E_NUMEROPIECE') and
     (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC]) then begin
    TOBE := FindCreerTOBEscompte(True);
    if TOBE <> nil then
      {Indicateur carré jaune si ligne sans escompte}
      if TOBE.GetValue('SANSESCOMPTE')='X' then begin
        FListe.Canvas.Brush.Color := clYellow;
        FListe.Canvas.Brush.Style := bsSolid;
        FListe.Canvas.Pen.Color   := clYellow;
        FListe.Canvas.Pen.Mode    := pmCopy;
        FListe.Canvas.Pen.Width   := 1;
        FListe.Canvas.Rectangle(ARect.Right - 5, ARect.Top + 1, ARect.Right - 1, ARect.Top + 5);
      end;
  end;

  if (FListe.ColNames[ACol]='E_DATEECHEANCE') then begin
    K := -1;
    for n := 0 to Grille.ColCount - 1 do
      if Grille.ColNames[n] = 'E_DATEECHEANCE' then begin
        K := n;
        Break;
      end;

    if K >= 0 then begin
      if OkLignePourEscompte then begin
        TheColor:=clYellow ;
        if (GetCheckBoxState('ESCCBProrata') = cbChecked) then begin
          DD:=Q.FindField('E_DATEECHEANCE').AsDateTime;
          if Arrondi(DD - Date, 0) <= 0 then TheColor := clred;
        end;
        FListe.Canvas.Brush.Color := TheColor;
        FListe.Canvas.Brush.Style := bsSolid;
        FListe.Canvas.Pen.Color   := clGreen;
        FListe.Canvas.Pen.Mode    := pmCopy;
        FListe.Canvas.Pen.Width   := 1;
        FListe.Canvas.Rectangle(ARect.Right - 5, ARect.Top + 1, ARect.Right - 1, ARect.Top + 5);
      end;
    end;
  end;
end;

{$ELSE}
{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 27/10/2005
Modifié le ... :   /  /
Description .. : Report de l'évènement OnDrawDataCell, non déclenché sur
Suite ........ : la Grid
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENEREMP.FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var sRIB, St : String ;
    TOBE     : TOB ;
    TheColor : TColor ;
    DD : TDateTime ;
begin
  // partie commune
  if Assigned(FListeDrawColumnCellParent) then
    FListeDrawColumnCellParent(Sender,Rect,DataCol,Column,State);
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  if FTic.Checked
    then St:='E_GENERAL'
    else St:='E_AUXILIAIRE' ;
  if ((Column.FieldName=St) and (Q.FindField('E_RIB')<>Nil)) then
    begin
    sRIB:=Q.FindField('E_RIB').AsString ;
    if sRib<>'' then Exit ;
    FListe.Canvas.Brush.Color := clRed ;
    FListe.Canvas.Brush.Style := bsSolid ;
    FListe.Canvas.Pen.Color   := clRed ;
    FListe.Canvas.Pen.Mode    := pmCopy ;
    FListe.Canvas.Pen.Width   := 1 ;
    FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
    end ;
  // Uniquement pour les Chèques fournisseurs : Escompte par ligne
  if (Column.FieldName='E_NUMEROPIECE') and (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC]) then
    BEGIN
    TOBE := FindCreerTOBEscompte(True) ;
    if TOBE <> Nil then
      // Indicateur carré jaune si ligne sans escompte
      if TOBE.GetValue('SANSESCOMPTE')='X' then
        begin
        FListe.Canvas.Brush.Color:= clYellow ;
        FListe.Canvas.Brush.Style:=bsSolid ;
        FListe.Canvas.Pen.Color:=clYellow ;
        FListe.Canvas.Pen.Mode:=pmCopy ;
        FListe.Canvas.Pen.Width:= 1 ;
        FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
        end;
    END ;
  // ??
  if ((Column.FieldName='E_DATEECHEANCE') and (Q.FindField('E_DATEECHEANCE')<>Nil)) then
    BEGIN
    If OkLignePourEscompte Then
      BEGIN
      TheColor:=clYellow ;
      If (GetCheckBoxState('ESCCBProrata') = cbChecked) And (Q.FindField('E_DATEECHEANCE')<>NIL) Then
        BEGIN
        DD:=Q.FindField('E_DATEECHEANCE').AsDateTime ;
        If Arrondi(DD-Date,0)<=0 Then TheColor:=clred ;
        END ;
      FListe.Canvas.Brush.Color := TheColor ;
      FListe.Canvas.Brush.Style := bsSolid ;
      FListe.Canvas.Pen.Color   := clGreen ;
      FListe.Canvas.Pen.Mode    := pmCopy ;
      FListe.Canvas.Pen.Width   := 1 ;
      FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
      END ;
    END;
end;

// BPY le 18/10/2004 : demande de RR => trie des liste dans CCMP
procedure TOF_CPGENEREMP.FListeSortColumn(Sender: TObject);
begin
  FListe.ClearSelected;
end;
// Fin BPY
{$ENDIF}

procedure TOF_CPGENEREMP.FListeRowEnter(Sender: TObject);
begin
  VH^.MPModifFaite:=FALSE ;
{$IFDEF EAGLCLIENT}
  Q.TQ.Seek(FListe.Row - 1) ;
{$ENDIF}
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then
    VH^.MPPop.MPExoPop:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
  If Q.FindField('E_GENERAL')<>NIL Then
    VH^.MPPop.MPGenPop:=Q.FindField('E_GENERAL').AsString ;
  If Q.FindField('E_AUXILIAIRE')<>NIL Then
    VH^.MPPop.MPAuxPop:=Q.FindField('E_AUXILIAIRE').AsString ;
  If Q.FindField('E_JOURNAL')<>NIL Then
    VH^.MPPop.MPJalPop:=Q.FindField('E_JOURNAL').AsString ;
  If Q.FindField('E_NUMEROPIECE')<>NIL Then
    VH^.MPPop.MPNumPop:=Q.FindField('E_NUMEROPIECE').AsInteger ;
  If Q.FindField('E_NUMLIGNE')<>NIL Then
    VH^.MPPop.MPNumLPop:=Q.FindField('E_NUMLIGNE').AsInteger ;
  If Q.FindField('E_NUMECHE')<>NIL Then
    VH^.MPPop.MPNumEPop:=Q.FindField('E_NUMECHE').AsInteger ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL Then
    VH^.MPPop.MPDatePop:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
end;

procedure TOF_CPGENEREMP.bSelectAllClick(Sender: TObject);
begin
    AllSelecting := true;

  FetchLesTous;
  if Assigned(bSelectAllClickParent) then
    bSelectAllClickParent(Sender) ;

  // Si on sélectionne tout, on annule le mode "sélection inversée"
  if FListe.AllSelected then
  	begin
    SwapSelect       := False ;
    BSwapSelect.Down := False ;
		BSwapSelect.Hint := 'Activer le mode "sélection inversée"' ;
    end;
	AfficheSoldeSelection ;

    AllSelecting := false;
end;

procedure TOF_CPGENEREMP.BSwapSelectClick(Sender: TObject);
begin
  // Si tout sélectionné, on déselectionne tout
  FetchLesTous;
  if FListe.AllSelected then bSelectAllClick(nil) ;
	// Message d'avertissement
  if not SwapSelect	then
	  if PGIAsk('Vous allez passer en mode "sélection inversée", Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes then
	    begin
	    BSwapSelect.Down := False ;
      Exit;
  	 	end;
  // Inverse la sélection
  SwapSelect := not SwapSelect ;
  // Modification du Hint
  if SwapSelect
  	then BSwapSelect.Hint := 'Désactiver le mode "sélection inversée"'
  	else BSwapSelect.Hint := 'Activer le mode "sélection inversée"' ;
  // Recalcul du solde
 	AfficheSoldeSelection ;
end;

procedure TOF_CPGENEREMP.FListeFlipSelection(Sender: TObject);
begin
    // BPY le 28/05/2004 => demande FFF : pb de selection avec une deselection apres un selectall ...
    // la variable FListe.AllSelected n'est pas desseté
    if ((not AllSelecting) and (FListe.AllSelected)) then FListe.AllSelected := false;

	AfficheSoldeSelection ;
end;

procedure TOF_CPGENEREMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  Case Key of
{AE}   69 : BEGIN // Active-désactive les escomptes sur une ligne
            If (Shift=[ssCtrl]) And AvecEscompte Then
            	BEGIN
              Key:=0 ;
              SwapEscompteLigne ;
              FListe.Refresh ;
              END ;
					  If Shift=[ssAlt] then // ALT-E sur une ligne
           	  BEGIN
              Key:=0 ;
              MarquerLigneSansEscompte ;
              FListe.Refresh ;
              END ;
            END ;
{^M}   77 : if (Shift=[ssCtrl]) then
              BEGIN
              Key:=0 ;
              ChangeModeGrille ;
              END ;
// BPY le 23/08/2004 => Fiche n° 10445 => f10 lance le traitement et non la recherche !
       VK_F10 :
            begin
                BOuvrirClick(nil);
            end;
  else
    begin
      if Assigned(FormKeyDownParent) then
        FormKeyDownParent(Sender, Key, Shift);
    end;
// Fin BPY
  end ;
end;

procedure TOF_CPGENEREMP.InitVariables;
begin
  // Init TOB
  TOBOrig       := TOB.Create('',Nil,-1) ;
  TOBDest       := TOB.Create('',Nil,-1) ;

  TOBEscomptes  := TOB.Create('LESESCOMPTES',Nil,-1) ;
  TOBGHT        := TOB.Create('',Nil,-1) ;
  TOBGTVA       := TOB.Create('',Nil,-1) ;

  // Init TOB escomptes
  InitEscompteSup(TOBEscomptes) ;

  // Init TList
  LesCombi                := TStringList.Create ;
  LesCombi.Sorted         := True ;
  LesCombi.Duplicates     := dupIgnore ;

  LesCritsRupt            := TStringList.Create ;
  LesCritsRupt.Sorted     := True ;
  LesCritsRupt.Duplicates := dupIgnore ;

  // Init Booleens
  SwapSelect  := False ;   // Gestion du mode "selection inversé", non opérationnel par défaut

  {JP 12/01/06 : FQ 17299 : Suppression de EstSerie(S3)
   If EstSerie(S3)
    Then GereEscompte := FALSE
    Else }GereEscompte := True ;

  // init pour gestion lots
  NomCritLot    := 'LOT' + V_PGI.User ;
  NomCritReste  := 'CRIT' + V_PGI.User ;

  // Init variable Halley
  VH^.MPModifFaite := FALSE ;
  Fillchar(VH^.MPPOP,SizeOf(VH^.MPPOP),#0) ;

  // ParmsMP
  FillChar(ParmsMP,Sizeof(ParmsMP),#0) ;
  ParmsMP.NomFSelect := TFMul(Ecran).FNomFiltre ;
  ParmsMP.smp := smp ;

end;

procedure TOF_CPGENEREMP.Enleve(pszNat: String);
var szNatPie, szTemp, szResult : String ;
begin
  szNatPie := GetControlText('_E_NATUREPIECE') ;
  if (szNatPie = '') then
    szNatPie := 'AC;AF;ECC;FC;FF;OC;OD;OF;RC;RF;' ;
  while szNatPie<>'' do begin
    szTemp := ReadTokenSt(szNatPie);
    if (szTemp<>pszNat) then szResult := szResult + sztemp +';';
  end;
  SetControlText('_E_NATUREPIECE', szResult ) ;
end;

procedure TOF_CPGENEREMP.BParamsClick(Sender: TObject);
begin
  ParmsMP.CptS      := GetControlText('E_GENERAL') ;
  ParmsMP.DEV.Code  := GetControlText('E_DEVISE') ;
  ParmsMP.TIDTIC    := (GetCheckBoxState('FTID') = cbChecked) ;
  ParamsMP(False,ParmsMP) ;
end;

procedure TOF_CPGENEREMP.BEscompteClick(Sender: TObject);
Var TOBE : TOB ;
begin
  if not GereEscompte then Exit ;
  TOBE:=FindCreerTOBEscompte(False) ;
  if TOBE<>Nil then
    SaisieMPEscompte(TOBE) ;
end;

procedure TOF_CPGENEREMP.BLotEcrClick(Sender: TObject);
Var i       : Integer ;
    testPDF : Boolean ;
begin
  SelectOnLot := Not SelectOnLot ;
  testPDF := smp in [smpEncPreEdt,smpEncTraEdt,smpEncTraEdtNC,smpDecChqEdt,smpEncPreEdtNC,
                     smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,
                     smpDecVirInEdtNC,smpDecBorEdt,smpDecBorEdtNC] ;

  // Visibiilté des onglets
  for i:=0 to Pages.PageCount-1 do
    Pages.Pages[i].TabVisible := Not SelectOnLot ;
  TTabSheet(GetControl('PAvance', True)).TabVisible     := True ;
  TTabSheet(GetControl('PSQL', True)).TabVisible        := True ;
  TTabSheet(GetControl('INVISIBLE',True)).TabVisible    := False ;
  TTabSheet(GetControl('PBANQUELOT',True)).TabVisible   := SelectOnLot ;
  TTabSheet(GetControl('PGestionPDF', True)).TabVisible := testPDF ;
  FEscompteChange(NIL) ; // ??


  // Passage à la sélection par lot
  if SelectOnLot then
    begin
    SetActiveTabSheet('PBANQUELOT') ;
    InitCriteresPourLot ;
    CFactCredit.State := cbGrayed;
    if (smp=smpEncTraEdt) or (smp=smpEncTraEdtNC) Or
       (smp=smpDecBorEdt) Or (smp=smpDecBorEdtNC) then
      begin
      TraiteEdite.State := cbUnchecked ;
      TraiteEditeChange(Nil) ;
      end ;
    end
  else   // Retour à la Sélection standard
    begin
    SetActiveTabSheet('PCRITERE') ;
    SetControlText('E_NOMLOT', '') ;
    SetControlText('E_NOMLOT_', '') ;
    SetControlText('E_BANQUEPREVI', '') ;
    SetControlText('E_BANQUEPREVI_', '') ;
    end ;

  InitCATEGORIE ;
  InitConditions ;
end;

procedure TOF_CPGENEREMP.BParamListeClick(Sender: TObject);
begin
  // Annule le CTRL+M avant de passer en paramètrage de la liste
  if (dgEditing in FListe.Options) then
    SwapModeGrid(FListe,Q) ;
  if Assigned(BParamListeClickParent) then
    BParamListeClickParent(Sender) ;
end;

procedure TOF_CPGENEREMP.ESCCBPRORATAChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  FListe.Refresh ;
end;

procedure TOF_CPGENEREMP.ESCCBTVAChange(Sender: TObject);
Var AvecTVA : Boolean ;
begin
  TFMul(Ecran).CritModified := True;
  If Not AvecEscompte Then exit ;
  AvecTVA := GetCheckBoxState('ESCCBTVA') = cbChecked ;
  SetControlEnabled('ESCTVA',       AvecTVA ) ;
  SetControlEnabled('TESCTVA',      AvecTVA ) ;
  SetControlEnabled('ESCTAUXTVA',   AvecTVA ) ;
  SetControlEnabled('TESCTAUXTVA',  AvecTVA ) ;
end;

procedure TOF_CPGENEREMP.ESCMETHChange(Sender: TObject);
Var Avec : Boolean ;
begin
  TFMul(Ecran).CritModified := True;
  Avec := AvecEscompte ;
  SetControlEnabled('ESCHT',        Avec) ;
  SetControlEnabled('TESCHT',       Avec) ;
  SetControlEnabled('ESCTAUXESC',   Avec) ;
  SetControlEnabled('TESCTAUXESC',  Avec) ;
  SetControlEnabled('ESCTVA',       Avec) ;
  SetControlEnabled('TESCTVA',      Avec) ;
  SetControlEnabled('ESCTAUXTVA',   Avec) ;
  SetControlEnabled('TESCTAUXTVA',  Avec) ;
  SetControlEnabled('ESCCBTVA',     Avec) ;
  SetControlEnabled('ESCCBPRORATA', Avec) ;
  If Not Avec Then
    BEGIN
    SetControlText('ESCHT',       '') ;
    SetControlText('ESCTAUXESC',  '') ;
    SetControlText('ESCTVA',      '') ;
    SetControlText('ESCTAUXTVA',  '') ;
    SetControlText('ESCCBTVA',    '-') ;
    SetControlText('ESCCBPRORATA','-') ;
    END ;
end;

procedure TOF_CPGENEREMP.FEscompteChange(Sender: TObject);
var testESC : Boolean ;
begin
  TFMul(Ecran).CritModified := True;
  testESC := GetCheckBoxState('FEscompte') = cbChecked ;
  TTabSheet(GetControl('PEscompte',True)).tabVisible := testESC ;
  If Not testESC Then
    begin
    SetControlText('ESCMETH', 'RIE' ) ;
    ESCMETHChange(Nil) ;
    end ;
end;

procedure TOF_CPGENEREMP.ESCHTChange(Sender: TObject);
Var ESCTAUXESC : THEdit ;
begin
  TFMul(Ecran).CritModified := True;
  ESCTAUXESC := THEdit(GetControl('ESCTAUXESC', True)) ;
	ESCTAUXESC.Enabled := GetControlText('ESCHT') <> '' ;
  if ESCTAUXESC.Enabled
  	then ESCTAUXESC.Color := ClWindow
 		else ESCTAUXESC.Color := ClBtnFace;
end;

procedure TOF_CPGENEREMP.ESCTVAChange(Sender: TObject);
Var ESCTAUXTVA : THEdit ;
begin
  TFMul(Ecran).CritModified := True;
  ESCTAUXTVA := THEdit(GetControl('ESCTAUXTVA', True)) ;
  ESCTAUXTVA.Enabled := GetControlText('ESCTVA') <> '' ;
  if ESCTAUXTVA.Enabled
  	then ESCTAUXTVA.Color := ClWindow
	 	else ESCTAUXTVA.Color := ClBtnFace;
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 18/07/2003
Modifié le ... :   /  /    
Description .. : Non utilisé pour l'instant...
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGENEREMP._chk_RegSupChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  if isEncMP(smp) then
    begin
    if _chk_RegSup.State=cbGrayed
      then SetControlText('_E_NATUREPIECE', GetControlText('_E_NATUREPIECE') + 'RC;' )
      else if _chk_RegSup.State=cbUnChecked
        then Enleve('RC')
        else // cbChecked
          begin
          SetControlText('_E_NATUREPIECE', 'RC;' ) ;
          cFactCredit.State     := cbUnChecked;
          end ;
    end
  else
    begin
    if _chk_RegSup.State = cbGrayed
      then SetControlText('_E_NATUREPIECE', 'RF;' )
      else if _chk_RegSup.State=cbUnChecked
        then Enleve('RF')
        else // cbChecked
          begin
          SetControlText('_E_NATUREPIECE', 'RF;' );
          cFactCredit.State := cbUnChecked;
          end;
    end;
end;

procedure TOF_CPGENEREMP._E_NATUREPIECEChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  if (GetControlText('_E_NATUREPIECE')='') then
    begin
    cFactCredit.State := cbGrayed;
    _chk_RegSup.State := cbGrayed;
    end;
//  RempliWhereEnc;
end;

procedure TOF_CPGENEREMP.SpoolerChange(Sender: TObject);
Var testSpooler : Boolean ;
begin
  testSpooler := GetCheckBoxState('Spooler') = cbChecked ;
  SetControlEnabled('RepSpooler',       testSpooler) ;
  SetControlEnabled('TRepSpooler',      testSpooler) ;
  SetControlEnabled('XFichierSpooler',  testSpooler) ;
  If testSpooler And (GetControlText('RepSpooler')='') Then
    begin
    {$IFDEF SPEC350}
    SetControlText('RepSpooler', ExtractFilePath(Application.EXEName) ) ;
    {$ELSE}
    Case smp Of
      smpDecChqEdt,smpDecChqEdtNC     : SetControlText('RepSpooler', GetParamSocSecur('SO_CPCHEMINCHEQUE','') ) ;
      smpEncTraEdt,smpEncTraEdtNC     : SetControlText('RepSpooler', GetParamSocSecur('SO_CPCHEMINTRAITE','') ) ;
      smpDecBorEdt,smpDecBorEdtNC     : SetControlText('RepSpooler', GetParamSocSecur('SO_CPCHEMINBOR','') ) ;
      smpDecVirEdt,smpDecVirEdtNC     : SetControlText('RepSpooler', GetParamSocSecur('SO_CPCHEMINVIREMENT','') ) ;
      smpDecVirInEdt,smpDecVirInEdtNC : SetControlText('RepSpooler', GetParamSocSecur('SO_CHEMINVIRIN','') ) ; // VL 301003 FQ 12958
      smpEncPreEdt,smpEncPreEdtNC     : SetControlText('RepSpooler', GetParamSocSecur('SO_CPCHEMINPRELEVEMENT','') ) ;
      end ;
    {$ENDIF}
    end ;
end;

procedure TOF_CPGENEREMP.FListeColEnter(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
// A FAIRE ??
{$ELSE}
  If (dgEditing in Fliste.Options) Then
    Fliste.SelectedField:=Q.FindField('E_SAISIMP') ;
{$ENDIF}
end;

procedure TOF_CPGENEREMP.AlimTobEscomptes;
begin
  if not GereEscompte then Exit ;
  TOBEscomptes.PutValue('ESCOMPTABLE',  'X') ;
  TOBEscomptes.PutValue('COMPTEHT',     GetControlText('ESCHT') ) ;
  TOBEscomptes.PutValue('TAUXESC',      Valeur(GetControlText('ESCTAUXESC')) ) ;
  // SBO correction fiche 12360 : si "tva sur escompte" non cochée, ne doit pas appliqué de tva à l'escompte 
  if GetCheckBoxState('ESCCBTVA') = cbChecked then
    begin
    TOBEscomptes.PutValue('COMPTETVA',    GetControlText('ESCTVA') ) ;
    TOBEscomptes.PutValue('TAUXTVA',      Valeur(GetControlText('ESCTAUXTVA')) ) ;
    end
  else
    begin
    TOBEscomptes.PutValue('COMPTETVA',    '' ) ;
    TOBEscomptes.PutValue('TAUXTVA',      0.00 ) ;
    end ;
  // Fin correction 12360
  TOBEscomptes.PutValue('MONTANTESC',   0.00) ;
// Ajout Champ SANSESCOMPTE
	TOBEscomptes.PutValue('SANSESCOMPTE', '-') ;
end;

function TOF_CPGENEREMP.AvecEscompte: Boolean;
begin
  Result:=FALSE ;
  If (GetControlText('ESCMETH')<>'RIE') And GereEscompte Then
    Result:=TRUE ;
end;

procedure TOF_CPGENEREMP.ChangeModeGrille;
Var
    i : integer;
{$IFDEF EAGLCLIENT}
    ListeCols : String;
{$ELSE}
    RJal,RExo : String;
    RDate : TDateTime;
    RNumP,RNumL,RNumEche : Integer;
    Pb : Boolean;
{$ENDIF}
begin
// bpy le 20/08/2004 => Fiche n° 11941 : amelioration de l'ergonomie de la fonction
    If (Q.FindField('E_SAISIMP') = nil) then exit;

    // Initialisation des variables locales
{$IFNDEF EAGLCLIENT}
    Pb := false;
    RDate := iDate1900;
    RNumP := 0;
    RNumL := 0;
    RNumEche := 0;
    RJal := '';
    RExo := '';

    // svg des donné pour retrouvé la ligne !
    if (Q.FindField('E_JOURNAL') <> nil) then RJal := Q.FindField('E_JOURNAL').AsString else Pb := true;
    if (Q.FindField('E_DATECOMPTABLE') <> nil) then RExo := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) else Pb := true;
    if (Q.FindField('E_DATECOMPTABLE') <> nil) then RDate := Q.FindField('E_DATECOMPTABLE').AsDateTime else Pb := true;
    if (Q.FindField('E_NUMEROPIECE') <> nil) then RNumP := Q.FindField('E_NUMEROPIECE').AsInteger else Pb := true;
    if (Q.FindField('E_NUMLIGNE') <> nil) then RNumL := Q.FindField('E_NUMLIGNE').AsInteger else Pb := true;
    if (Q.FindField('E_NUMECHE') <> nil) then RNumEche := Q.FindField('E_NUMECHE').AsInteger else Pb := true;
{$ENDIF}

{$IFDEF EAGLCLIENT}
    if (goEditing in Fliste.Options) then
    begin
        ListeCols := '';
        for i := 0 to FListe.ColCount-1 do
        begin
            if (FListe.ColNames[i] = 'E_SAISIMP') then ListeCols := ListeCols + 'E_SAISIMP;'
            else ListeCols := ListeCols + ';'
        end;

        for i := 0 to Q.TQ.Detail.count-1 do
        begin
            Q.TQ.Detail[i].VirtuelleToReelle('ECRITURE');
            Q.TQ.Detail[i].GetLigneGrid(FListe,i+1,ListeCols);
            if (Q.TQ.Detail[i].IsFieldModified('E_SAISIMP')) then Q.TQ.Detail[i].UpdateDB();
        end;
    end;
{$ENDIF}

    // changement du mode de la grille
    SwapModeGrid(FListe,Q);

    // relance la recherche pour reremplire la grille !
    TFMul(Ecran).BChercheClick(Nil);

{$IFDEF EAGLCLIENT}
    if (goEditing in Fliste.Options) then
    begin
        for i := 0 to FListe.ColCount-1 do
        begin
            if (FListe.ColNames[i] = 'E_SAISIMP') then FListe.ColEditables[i] := true
            else FListe.ColEditables[i] := false;
        end;

// ca marche pas ....    if (not Pb) then Q.TQ.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[]);
    end;
{$ELSE}
    if (dgEditing in Fliste.Options) then
    begin
        for i := 0 to FListe.Columns.Count-1 do
        begin
            if (FListe.Columns.Items[i].FieldName = 'E_SAISIMP') then FListe.Columns.Items[i].ReadOnly := false
            else FListe.Columns.Items[i].ReadOnly := true;
        end;

        Fliste.SelectedField := Q.FindField('E_SAISIMP');
    end;

    if (not Pb) then Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[]);
{$ENDIF}
// fin BPY
end;

procedure TOF_CPGENEREMP.ClearEscomptes;
begin
  if not GereEscompte then Exit ;
  TOBEscomptes.ClearDetail ;
  TOBGHT.ClearDetail ;
  TOBGTVA.ClearDetail ;
end;

function TOF_CPGENEREMP.FindCreerTOBEscompte(QueFind: boolean): TOB;
Var TOBE : TOB ;
    Jal,Exo  : String ;
    NumP,NumL,NumE : integer ;
begin
  Result := Nil ;
  if not GereEscompte then Exit ;
  if Not AvecEscompte then Exit ;
{$IFDEF EAGLCLIENT}
  if Q.TQ.Detail.Count = 0 then Exit ;
  if not Q.TQ.FieldExists('E_JOURNAL') then Exit ;
  Q.TQ.Seek(FListe.Row - 1) ;
{$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  if Q.FindField('E_JOURNAL')=nil then Exit ;
{$ENDIF}
  Jal  := Q.FindField('E_JOURNAL').AsString ;
  Exo  := Q.FindField('E_EXERCICE').AsString ;
  NumP := Q.FindField('E_NUMEROPIECE').AsInteger ;
  NumL := Q.FindField('E_NUMLIGNE').AsInteger ;
  NumE := Q.FindField('E_NUMECHE').AsInteger ;
  TOBE := TOBEscomptes.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                                 [Jal,Exo,NumP,NumL,NumE],False) ;
  if ((TOBE=Nil) and (Not QueFind)) then
    BEGIN
    TOBE:=TOB.Create('ESCOMPTE',TOBEscomptes,-1) ;
    // Infos ecriture
    TOBE.AddChampSup('E_JOURNAL',False)     ; TOBE.PutValue('E_JOURNAL',Jal) ;
    TOBE.AddChampSup('E_EXERCICE',False)    ; TOBE.PutValue('E_EXERCICE',Exo) ;
    TOBE.AddChampSup('E_NUMEROPIECE',False) ; TOBE.PutValue('E_NUMEROPIECE',NumP) ;
    TOBE.AddChampSup('E_NUMLIGNE',False)    ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
    TOBE.AddChampSup('E_NUMECHE',False)     ; TOBE.PutValue('E_NUMECHE',NumE) ;
    // infos Escompte
    InitEscompteSup(TOBE) ;
    END ;
  Result:=TOBE ;
end;

procedure TOF_CPGENEREMP.MarquerLigneSansEscompte;
Var TOBE : TOB ;
begin
  if not (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC])
  	then Exit ;
	TOBE:=FindCreerTOBEscompte(FALSE) ;
	If TOBE<>NIL Then
  	begin
  	if TOBE.GetValue('SANSESCOMPTE')='X'
    	Then TOBE.PutValue('SANSESCOMPTE','-')
			Else TOBE.PutValue('SANSESCOMPTE','X') ;
	  end ;
end;

function TOF_CPGENEREMP.OkLignePourEscompte: Boolean;
Var TOBE : TOB ;
    Jal,Exo  : String ;
    NumP,NumL,NumE : integer ;
begin
  Result:=FALSE ;
  if not GereEscompte then Exit ;
  If Not AvecEscompte Then Exit ;
{$IFDEF EAGLCLIENT}
  if Q.TQ.Detail.Count = 0 then Exit ;
  if not Q.TQ.FieldExists('E_JOURNAL') then Exit ;
  Q.TQ.Seek(FListe.Row - 1) ;
{$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  if Q.FindField('E_JOURNAL')=nil then Exit ;
{$ENDIF}
  Jal  := Q.FindField('E_JOURNAL').AsString ;
  Exo  := Q.FindField('E_EXERCICE').AsString ;
  NumP := Q.FindField('E_NUMEROPIECE').AsInteger ;
  NumL := Q.FindField('E_NUMLIGNE').AsInteger ;
  NumE := Q.FindField('E_NUMECHE').AsInteger ;
  TOBE := TOBEscomptes.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                                 [ Jal,        Exo,         NumP,           NumL,        NumE],
                                 False) ;
  if (TOBE=Nil)
    then Result := TRUE
    Else if TOBE.GetValue('ESCOMPTABLE')='X'
           Then Result := TRUE ;
end;

procedure TOF_CPGENEREMP.SwapEscompteLigne;
var TOBE : TOB ;
begin
  TOBE := FindCreerTOBEscompte(FALSE) ;
  if TOBE<>NIL Then
    if TOBE.GetValue('ESCOMPTABLE')='X'
      Then TOBE.PutValue('ESCOMPTABLE','-')
      Else TOBE.PutValue('ESCOMPTABLE','X') ;
end;

function TOF_CPGENEREMP.ConstitueOrigines: Boolean;
var i                         : Integer ;
    TOBL                      : TOB ;
    StBQE, StCAT, StMP        : String ;
    St, LesChamps, NomChamp   : String ;
    LesVals, OldMP            : String ;
    Premier, RuptMP           : Boolean ;
    ResaisieMP, PasPris       : Boolean ;
//    lBoTestCPT                : Boolean ;
Label 0 ;
begin
  0 :
  Result    := TRUE ;
  PrechargeOrigines ;

  if FListe.AllSelected then
    ToutMarquer
  else begin
    for i:=0 to FListe.NbSelected-1 do
    begin
      FListe.GotoLeBookmark(i) ;
      {$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.row - 1) ;
      {$ENDIF}
      MarqueOrigine;
    end ;
    VireInutiles(TOBOrig, swapSelect);
  end;

  Premier := TRUE ;
  RuptMP  := FALSE ;
  AlimTobEscomptes ; // Alim entete de tobescomptes avec valeurs du critère
  PasPris := FALSE ;

  // Vérifications : Compte génération <> compte pièce comptable
//  lBoTestCPT := True ;
  for i:=0 to TOBOrig.Detail.Count-1 do
    begin
    TOBL:=TOBOrig.Detail[i] ;
    if TOBL.GetValue('E_GENERAL')=ParmsMP.CptG then
      begin
      Result  := FALSE ;
      HShowMessage(HM[4],gszCaption,'') ;
      Break ;
      end ;
    // Ce test n'est plus bloquant
  // EN ATTENTE MODIF LETTRAGE    SBO
{    if TOBL.GetValue('E_GENERAL')=ParmsMP.CptG
      then lBoTestCPT := False ;
}    if Premier
      then OldMP  := TOBL.GetValue('E_MODEPAIE')
      Else RuptMP := TOBL.GetValue('E_MODEPAIE')<>OldMP ;
    Premier := FALSE ;
    if GereEscompte then
      TraiteTOBEscompte( TOBL,
                         TOBEscomptes, TOBGHT, TOBGTVA ,
                         GetCheckBoxState('ESCCBPRORATA') = cbChecked ,
                         GetControlText('ESCMETH') ,
                         ParmsMP , PasPris ) ;
    end ;

  // Affichage message si compte géné = compte pièce origine MAIS non bloquant
  if Not Result Then Exit ;
  // EN ATTENTE MODIF LETTRAGE    SBO
{  if not lBoTestCPT then
    if HShowMessage(HM[4],gszCaption,'')<>mrYes Then
      begin
      result := false ;
      Exit ;
      end ;
}
  if PasPris And GereEscompte And (GetCheckBoxState('ESCCBProrata') = cbChecked) Then
    If HShowMessage(HM[10],gszCaption,'')<>mrYes Then Exit ;

  If RuptMP And (ParmsMP.MPG='') Then
    begin
    ResaisieMP := FALSE ;
    // Début 11640
    case smp Of
      smpEncTraEdt,smpEncTraPor : If HShowMessage(HM[2],gszCaption,'')<>mrYes Then
                                    begin
                                    Result := False;
                                    Exit;
                                    end;
      else
        begin
        //SG6 16.03.05 FQ 14831
        //En mode detail on reprend les modes de paiement des factures pour les lignes de reglement
        if ParmsMP.GroupeEncadeca <> 'DET' then
        begin
          HShowMessage(HM[8],gszCaption,'') ;
          ResaisieMP:=TRUE ;
        end
        else
        begin
          If HShowMessage(HM[2],gszCaption,'')<>mrYes Then
          begin
            Result := False;
            Exit;
          end;
        end;
        end ;
      end ; // fin du caase
    // Fin 11640

    if ResaisieMp Then
      if Not ParamsMPSup(ParmsMP,TRUE) then
        begin
        Result := False;
        Exit;
        end
      else
        Goto 0 ;
    end ;

  AjouteLesAnas( TOBOrig ) ;

  if smp in [smpEncTous,smpDecTous] then
    for i:=0 to TOBOrig.Detail.Count-1 do
      BEGIN
      TOBL  := TOBOrig.Detail[i] ;
      StBQE := TOBL.GetValue('E_BANQUEPREVI') ;
      StMP  := TOBL.GetValue('E_MODEPAIE') ;
      StCat := MPToCategorie(StMP) ;
      if StCat<>'' then
        BEGIN
        St := StBQE + ';' + StCat + ';' ;
        if LesCombi.IndexOf(St)<0 then LesCombi.Add(St) ;
        LesChamps  := ParmsMP.ChampsRupt ;
        LesVals    := '' ;
        Repeat
          NomChamp := ReadTokenSt(LesChamps) ;
          if NomChamp<>'' then
            LesVals := LesVals + TOBL.GetValue(NomChamp) + ';' ;
        Until ((LesChamps='') or (NomChamp=''));
        LesCritsRupt.Add(LesVals) ;
        END ;
      END
  else
    if ParmsMP.ChampsRupt<>'' then
      for i:=0 to TOBOrig.Detail.Count-1 do
        BEGIN
        TOBL      := TOBOrig.Detail[i] ;
        LesChamps := ParmsMP.ChampsRupt ;
        LesVals   := '' ;
        Repeat
          NomChamp:=ReadTokenSt(LesChamps) ;
          if NomChamp<>'' then
            LesVals := LesVals + TOBL.GetValue(NomChamp) + ';' ;
          Until ((LesChamps='') or (NomChamp=''));
          if LesCritsRupt.IndexOf(LesVals)<0 then
            LesCritsRupt.Add(LesVals) ;
        END ;
end;

procedure TOF_CPGENEREMP.IncNumLot;
Var NL : String ;
begin
  NL := IncNumLotTraChq(ParmsMP.NumEncaDeca) ;
  ParmsMP.NumEncaDeca := Copy(NL,1,17) ;
  CreerCodeLot(ParmsMP.NumEncaDeca) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 27/06/2003
Modifié le ... : 27/06/2003
Description .. : 27/06/2003 : Cette fonction ne servait à rien : au final,
Suite ........ : retournait le résultat du paramètre ! Algo corrigé.
Mots clefs ... : 
*****************************************************************}
function TOF_CPGENEREMP.VerifEcheMP(Alerte: Boolean): boolean;
Var i, ind          : integer ;
    TOBE            : TOB ;
    Premier, RuptMP : Boolean ;
    RuptEche        : boolean ;
    ExRuptMP        : boolean ;
    ExRuptEche      : boolean ;
    OldEche         : TDateTime ;
    OldMP           : String ;
begin
  Result := True ;
  if TobDest.Detail.Count = 0 then
    begin
    Result := False ;
    PGIInfo('Aucun auxiliaire à traiter ! Le solde de la sélection est nul.', Ecran.Caption ) ;
    Exit ;
    end ;
  if Not Alerte then Exit ;
  // Initialisation des variables locales
	ExRuptEche  := False ;
  ExRuptMP    := False ;
  RuptMP		  := False ;
  RuptEche    := False ;
  Premier 	  := False ;
  OldEche     := iDate1900 ;
  OldMP       := '' ;

  // Parcours des lignes
  for i:=0 to TOBDest.Detail.Count-1 do
    begin
    TOBE := TOBDest.Detail[i] ;

    // Par du principe que la première ligne est un compte de banque --> totalement faux !
    // D'où l'error
    if TOBE.GetValue('E_GENERAL') = ParmsMP.CptG then
      begin
      Premier  := True ;
      RuptMP   := False ;
      RuptEche := False ;
      end ;

    // sur les lignes de banque -> remise a zéro des données de ruptures 
    if Premier then
      begin
      OldMP   := TOBE.GetValue('E_MODEPAIE') ;
      OldEche := TOBE.GetValue('E_DATEECHEANCE') ;
      end
    else
      begin
      // sur les lignes de tiers on test les champs
      if Not ParmsMP.ForceEche then
        RuptEche := (TOBE.GetValue('E_DATEECHEANCE')<>OldEche) ;
      if ParmsMP.MPG='' then
        RuptMP   := (TOBE.GetValue('E_MODEPAIE')<>OldMP) ;
      end ;
    Premier := False ;

    // on met de cote le resultat des tests...
    if RuptMP then ExRuptMP := True ;
    if RuptEche then ExRuptEche := True ;

    // On sort de la boucle quand les 2 types de ... ont été détecté.
    if ((ExRuptEche) and (ExRuptMP)) then Break ;
    end ;

  // Mode de paiement et date d'échéance identique pour toutes les lignes
  // Pas de problème...on sort
  if ((Not ExRuptEche) and (Not ExRuptMP)) then Exit ;

  // Mode de paiement et/ou date d'échéance différentes pour certaines lignes
  // problème...message...
  if ((ExRuptEche) and (ExRuptMP))
    then ind:=1
    else if ExRuptMP
           then ind:=2
           else ind:=3 ;
  if HShowMessage(HM[ind],gszCaption,'')<>mrYes then
    Result := False ;
end;

procedure TOF_CPGENEREMP.InitComposants;
begin

  // composants du Mul
 	Q               := THQuery(GetControl('Q', True)) ;
{$IFDEF EAGLCLIENT}
  FListe          := THGrid(GetControl('FListe', True));
{$ELSE}
  FListe          := THDBGrid(GetControl('FListe', True));
{$ENDIF}
  Pages           := TPageControl(GetControl('Pages', True)) ;
  FFiltres        := THValComboBox(GetControl('FFiltres', True)) ;

  // onglet standard
  FTIC            := TCheckBox(GetControl('FTIC', True)) ;            // C
  CATEGORIE       := THValComboBox(GetControl('CATEGORIE', True)) ;   // C

{b fb 27/10/2005 FQ 12487}
  Selection       := THCritMaskEdit(GetControl('E_GENERAL', true));
  Auxiliaire      := THCritMaskEdit(GetControl('E_AUXILIAIRE', true));  
{e fb 27/10/2005 FQ 12487}
  // onglet compléments
  cFactCredit     := TCheckBox(GetControl('cFactCredit', True)) ;     // C
  TraiteEdite     := TCheckBox(GetControl('TraiteEdite', True)) ;     // C
  _chk_RegSup     := TCheckBox(GetControl('_chk_RegSup', True)) ;
  // VL 02122003
  TraiteEdite.AllowGrayed := True;
  cFactCredit.AllowGrayed := True;
  _CHK_REGSUP.AllowGrayed := True;

  // Boutons Toolbar
  BSwapSelect     := TToolbarButton97(GetControl('BSwapSelect', True)) ;

  // 10271
  SetControlVisible('BPARAMS', not (smp in [smpDecVirInEdtNC, smpDecVirEdtNC, smpDecBorEdtNC, smpEncTraEdtNC, smpEncPreEdtNC, smpDecChqEdtNC]));

  // FQ 15584 : SBO 04/04/2005
  SetCOntrolProperty('REPSPOOLER', 'DATATYPE', 'SAVEFILE(');  
end;

procedure TOF_CPGENEREMP.InitHelpContext;
begin
  { CA - 30/08/2007 - Tous les contextes d'aide sont définis dans la fonction
    AttribHelp de l'unité MulSMPUtil.pas }
  TFMul(Ecran).HelpContext  := AttribHelp(smp,False) ;
end;

procedure TOF_CPGENEREMP.InitConditions;
var StXP, StXP2 : String ;
    StXN, StXN2 : String ;
    St, StV8    : String ;
begin
  // Conditions sur les à nouveaux                  XX_WHEREAN
  St   := 'E_ECRANOUVEAU="N"' ;
  StV8 := LWhereV8 ;
  if StV8<>'' then
    begin
    St:= '(' + St + ' OR E_ECRANOUVEAU="H")  ' ;
    St:= St + ' AND (' + StV8 + ') ' ;
    end ;

  {YMO 06/03/07 FQ17124 Gestion du V_Pgi.Confidentiel}
  St := St + ' AND ' + CGenereSQLConfidentiel('E'); {:Ent1}
  {Utilisation du XX_WHEREAN car il ne bouge plus ensuite}
  
  SetControlText('XX_WHEREAN', St ) ;

  // Conditions nature de compte et montants        XX_WHERENATCPT
  ReinitWhereNatCpt( SelectOnLot ) ;                // FQ19410 : SBO 04/01/2007 Ne pas filtrer les tiers en recherche par lot 

  // Condition sur le numéro de traite              XX_WHERETRACHQ
  SetControlText('XX_WHERETRACHQ', '') ;

  // Condition sur ???                              XX_WHEREMP
  SetControlText('XX_WHEREMP', '') ;

  // Condition sur les lots                         XX_WHERELOT
  if smp = smpEncTous then
    SetControlText('XX_WHERELOT', 'E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' ) ;

  // XX_WHEREENC
  SetControlText('XX_WHEREENC','');

  // XX_WHEREMONTANT
  StXP  := StrFPoint(  9 * Resolution( V_PGI.OkDecV + 1 ) ) ;
  StXN  := StrFPoint( -9 * Resolution( V_PGI.OkDecV + 1 ) ) ;
  StXP2 := StrFPoint(  9 * Resolution( V_PGI.OkDecE + 1 ) ) ;
  StXN2 := StrFPoint( -9 * Resolution( V_PGI.OkDecE + 1 ) ) ;
  St    := '(E_DEBIT+E_CREDIT-E_COUVERTURE not between ' + StXN + ' AND ' + StXP + ')';
  SetControlText('XX_WHEREMONTANT', St ) ;
end;

procedure TOF_CPGENEREMP.InitCATEGORIE ;
begin
  Case smp of
    smpEncDiv,
    smpDecDiv           : CATEGORIE.ItemIndex := 0 ;

    smpEncPreBqe,
    smpEncPreEdt,
    smpEncPreEdtNC      : CATEGORIE.Value := 'PRE' ;


    smpEncCBBqe         : CATEGORIE.Value := 'CB' ;

    smpEncCBPor,
    smpEncTraPor,
    smpEncTraEdt,
    smpEncTraEdtNC,
    smpEncTraEnc,
    smpEncTraEsc,
    smpEncTraBqe,
    smpDecBorEdt,
    smpDecBorEdtNC,
    smpDecborDec,
    smpDecBorEsc,
    smpDecTraBqe,
    smpDecTraPor        : CATEGORIE.Value := 'LCR' ;

    smpEncChqPor,
    smpEncChqBqe,
    smpDecChqEdt,
    smpDecChqEdtNC      : CATEGORIE.Value := 'CHQ' ;

    smpEncTous,
    smpDecTous          : CATEGORIE.Value := '' ;

    smpDecVirEdt,
    smpDecVirEdtNC,
    smpDecVirBqe        : CATEGORIE.Value   := 'VIR' ;

    smpDecVirInEdt,
    smpDecVirInEdtNC,
    smpDecVirInBqe      : CATEGORIE.Value   := 'TRI' ;
   end ;

  ParmsMP.Cat := CATEGORIE.Value ;
  PreciseMP ;
end;

procedure TOF_CPGENEREMP.UpdateConditions;
var szTemp, szLib : String;
begin

  // XX_WHEREMONTANT et XX_WHEREAN sont initialisé une fois
  //  pour toute dans le initConditions

  // XX_WHEREENC
  // VL 02122003
  TraductionTHMultiValComboBox(THMultiValComboBox(GetControl('_E_NATUREPIECE')), szTemp, szLib, 'E_NATUREPIECE');
  if (szTemp <> '') then begin
    // Point 172 FFF (La sélection sur la nature de pièce ne fonctionne pas)
    SetControlText('XX_WHEREENC',sztemp);
{    szTemp := ' OR ' + szTemp;

    if isEncMP(smp) then
      begin
      if cFactCredit.State=cbGrayed
        then SetControlText('XX_WHEREENC','E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"'+szTemp+'))' )
        else if cFactCredit.State=cbChecked
          then SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"'+szTemp+'))' )
          else SetControlText('XX_WHEREENC','E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"'+szTemp+'))' ) ;   // cbUnChecked
      end
    else
      begin
      if cFactCredit.State=cbGrayed
        then SetControlText('XX_WHEREENC','E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"'+szTemp+'))' )
        else if cFactCredit.State=cbChecked
          then SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"'+szTemp+'))' )
          else SetControlText('XX_WHEREENC','E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"'+szTemp+'))' ) ;   // cbUnChecked
      end}
    end
  else
    SetControlText('XX_WHEREENC','');

  // XX_WHERELOT
  if SelectOnLot
    then SetControlText('XX_WHERELOT', '(E_NOMLOT<>"")' )
    else if smp = smpEncTous
      then SetControlText('XX_WHERELOT', 'E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' )
      else SetControlText('XX_WHERELOT', '' ) ;

  // XX_WHERETRACHQ
  if TraiteEdite.State=cbGrayed
    then SetControlText('XX_WHERETRACHQ', '')
    else if TraiteEdite.State=cbChecked
      then SetControlText('XX_WHERETRACHQ', 'E_NUMTRAITECHQ<>""')
      else SetControlText('XX_WHERETRACHQ', 'E_NUMTRAITECHQ=""' );   // cbUnChecked

  // XX_WHERENATCPT
  ReinitWhereNatCpt( SelectOnLot or                                  // FQ19410 : SBO 04/01/2007 Ne pas filtrer les tiers en recherche par lot
                     ( Trim(GetControlText('E_GENERAL')) <> '' ) or
                     ( Trim(GetControlText('E_AUXILIAIRE')) <> '' ) ) ;

  // XX_WHEREMP
  if SelectOnLot
    then LaCategorieChange( THValCombobox( GetControl('CATEGORIE1', True) ),
                            THEdit(GetControl('XX_WHEREMP',True)),
                            SorteLettre )
    else LaCategorieChange( THValCombobox( GetControl('CATEGORIE', True) ),
                            THEdit(GetControl('XX_WHEREMP',True)),
                            SorteLettre ) ;

  {JP 18/02/05 :  Gestion du filtre E_BANQUEPREVI}
  WhereBqPrevi;
end;


procedure TOF_CPGENEREMP.InitEscompte;
begin
  // Gestion bouton escompte (??)
  If Not GereEscompte Then
    SetControlVisible('BEscompte', FALSE ) ;

  // Gestion de l'escompte
  if smp in [smpDecChqEdt,smpDecVirEdt,smpDecVirInEdt] then
    begin
    // indicateurs
    GereEscompte         := True ;
    // si ParamSoc
    JalEsc := GetParamSocSecur('SO_CPJALESCOMPTE','') ;
    if JalEsc='' then GereEscompte       := False ;
    end
  else
    begin
    // indicateurs
    GereEscompte         := False ;
    // ParamSoc ?
    JalEsc               := '' ;
    end ;

  // cb visibles
  SetControlVisible('FEscompte', GereEscompte) ;
  // Init WhereEsc
  SetControlChecked('FEscompte',False);
  FEscompteChange(nil) ;
{  SetControlText( 'ESCMETH' ,   'RIE' ) ;
  ESCMETHChange(NIL) ; }

end;

procedure TOF_CPGENEREMP.InitSelectionLot;
begin

  // Page de sélection sur lot non apparrente au lancement
  TTabSheet(GetControl('PBanqueLot',True)).TabVisible := FALSE ;
  SetActiveTabSheet('PCritere') ;

  //{$IFNDEF CCMP} JP 22/02/05 : cela n'a aucune raison d'être
{  if ( smp in [smpEncDiv,smpDecDiv] )
    then SetControlVisible('BLotEcr', FALSE) ;}
  //{$ENDIF}

  // Page de sélection sur lot non apparrente au lancement
  if IsEnc(SMP)
    then if (not VH^.CCMP.LotCli)
          then BLotEcrClick(Nil)
          else
    else if (Not VH^.CCMP.LotFou)
          then BLotEcrClick(Nil) ;
  selectOnLot   := FALSE ;
end;

procedure TOF_CPGENEREMP.InitCriteresPourLot;
begin

    SetControlText('E_GENERAL',     '') ;
    SetControlText('E_AUXILIAIRE',  '') ;
    // On ne doit pas renseigner la devise lors de la gestion par lot
    // car comment afficher les lots en devises si la devise est forcé à V_PGI.DevisePivot ?
    SetControlText('E_DEVISE',      '') ; // 14737

    SetControlProperty('CATEGORIE',   'VIDE',      TRUE ) ;
    SetControlProperty('E_MODEPAIE',  'VIDE',      TRUE ) ;
    SetControlText('CATEGORIE',       '' ) ;
    SetControlText('E_MODEPAIE',      '' ) ;
    CategorieChange(CATEGORIE) ;

    SetControlText('E_EXERCICE',      '' ) ;
    SetControlText('E_JOURNAL',       '' ) ;

    THMultiValComboBox(GetControl('_E_NATUREPIECE',True)).SelectAll ;

    SetControlText('E_DATECOMPTABLE',     StDate1900) ;
    SetControlText('E_DATECOMPTABLE_',    StDate2099) ;
    SetControlText('E_DATEECHEANCE',      StDate1900) ;
    SetControlText('E_DATEECHEANCE_',     StDate2099) ;
    SetControlText('E_NUMEROPIECE',       '') ;
    SetControlText('E_NUMEROPIECE_',      '') ;
    SetControlText('E_ETABLISSEMENT',     '' ) ;

    SetControlText('E_NUMENCADECA',   '') ;
    SetControlText('E_DEBIT',         '') ;
    SetControlText('E_DEBIT_',        '') ;
    SetControlText('E_CREDIT',        '') ;
    SetControlText('E_CREDIT_',       '') ;
    SetControlText('E_NUMTRAITECHQ',  '') ;
    SetControlText('E_NUMTRAITECHQ_', '') ;
    SetControlText('T_TABLE0', '') ;
    SetControlText('T_TABLE1', '') ;
    SetControlText('T_TABLE2', '') ;
    SetControlText('T_TABLE3', '') ;
    SetControlText('T_TABLE4', '') ;
    SetControlText('T_TABLE5', '') ;
    SetControlText('T_TABLE6', '') ;
    SetControlText('T_TABLE7', '') ;
    SetControlText('T_TABLE8', '') ;
    SetControlText('T_TABLE9', '') ;

end;

procedure TOF_CPGENEREMP.InitEvenements;
Var
  lBouton   : TButton ;
begin

{$IFDEF EAGLCLIENT}
  FListe.DefaultDrawing := False;
  FListe.OnDrawCell := FListeDrawCell;
//  FListe.GetCellCanvas := GetCellCanvasFListe;
{$ELSE}
  if Assigned(FListe.OnDrawColumnCell)
    then FListeDrawColumnCellParent := FListe.OnDrawColumnCell
    else FListeDrawColumnCellParent := nil ;
  FListe.OnDrawColumnCell    := FListeDrawColumnCell ;
  // BPY le 18/10/2004 : demande de RR => trie des liste dans CCMP
  FListe.SortEnabled       := true;
  FListe.OnSortColumn      := FListeSortColumn ;
  // Fin BPY
{$ENDIF}
  FListe.OnDblClick        := FListeDblClick ;
  FListe.OnFlipSelection   := FListeFlipSelection ;

  // Affectation EVT Boutons
    // -> tout sélectionner
    lBouton := TButton(GetControl('bSelectAll',True)) ;
    if Assigned(lBouton.OnClick)
      then bSelectAllClickParent    := lBouton.OnClick
      else bSelectAllClickParent    := nil ;
    lBouton.OnClick := bSelectAllClick ;
    // -> Param liste
    lBouton := TButton(GetControl('bParamListe',True)) ;
    if Assigned(lBouton.OnClick)
      then BParamListeClickParent    := lBouton.OnClick
      else BParamListeClickParent    := nil ;
    lBouton.OnClick := BParamListeClick ;
    // -> Ouvrir
    TButton(GetControl('bOuvrir',True)).OnClick    := bOuvrirClick ;
    // -> Nouvelle recherche
    lMenuItem := TMenuItem(GetControl('BNouvRech',True)) ;
    {JP 19/08/04 : FQ 14212 : ce n'est pas utile
    if Assigned(lMenuItem.OnClick)
      then BNouvRechClickParent    := lMenuItem.OnClick
      else BNouvRechClickParent    := nil ;
    lMenuItem.OnClick := BNouvRechClick ;}

  // EVT Form
  if Assigned(TFMul(Ecran).OnKeyDown)
    then FormKeyDownParent := TFMul(Ecran).OnKeyDown
    else FormKeyDownParent := nil ;
  TFMul(Ecran).OnKeyDown := FormKeyDown ;

  // Ftn spécifique
  TToolbarButton97(GetControl('BCtrlRib', True)).OnClick      := BCtrlRibClick ;
  TToolbarButton97(GetControl('BParams', True)).OnClick       := BParamsClick ;
  TToolbarButton97(GetControl('BEscompte', True)).OnClick     := BEscompteClick ;
  TToolbarButton97(GetControl('BLotEcr', True)).OnClick       := BLotEcrClick ;
  BSwapSelect.OnClick   := BSwapSelectClick ;

  // onglet standard
  THValComboBox(GetControl('E_EXERCICE',  True)).OnChange := E_EXERCICEChange ;
  TCheckBox(GetControl('FEscompte', True)).OnClick         := FEscompteChange ;
  CATEGORIE.OnChange     := CATEGORIEChange ;
  FTIC.OnClick          := FTICChange ;

  cFactCredit.OnClick   := cFactCreditChange ;
  TraiteEdite.OnClick   := TraiteEditeChange ;
{b fb 27/10/2005 FQ12487}
//  Selection.OnElipsisClick := SelectionElipsisClick;
//  Auxiliaire.OnElipsisClick := AuxiliaireElipsisClick;
{e fb 27/10/2005 FQ12487}
  // onglet écritures
  THMultiValComboBox(GetControl('_E_NATUREPIECE', True)).OnChange := _E_NATUREPIECEChange ;
  // onglet critères (avec banque / lot...)
  THValComboBox(GetControl('CATEGORIE1', True)).OnChange          := CATEGORIEChange ;
  // Onglet escompte
  THValComboBox(GetControl('ESCMETH', True)).OnChange  := ESCMETHChange ;      // méthode d'escompte
  TCheckBox(GetControl('ESCCBTVA', True)).OnClick     := ESCCBTVAChange ;      // avec TVA ?
  TCheckBox(GetControl('ESCCBPRORATA', True)).OnClick := ESCCBPRORATAChange ;  // au prorata ?
  THEdit(GetControl('ESCHT', True)).OnChange            := ESCHTChange ;        // compte HT
  THEdit(GetControl('ESCTVA', True)).OnChange           := ESCTVAChange ;       // compte TVA
  // onglet Spooler
  TCheckBox(GetControl('Spooler', True)).OnClick       := SpoolerChange ;
end;

procedure TOF_CPGENEREMP.FetchLesTous;
begin
  // Applique les critères s'ils ne le sont pas
  if TFMul(Ecran).CritModified then
  begin
    bgFetchAll := True;
    TFMul(Ecran).BChercheClick(Nil);
    Exit;
  end;

  // Récupère tous les enregistrement s'ils ne sont pas tous présents
  if bgFetchAll then
  begin
{$IFDEF EAGLCLIENT}
    TFMul(Ecran).FetchLesTous;
{$ENDIF}
    bgFetchAll := False;
  end;
end;

function TOF_CPGENEREMP.LInsertDB(T: Tob): Boolean;
var
  i : Integer;
begin
  try
    Result := T.InsertDB(Nil);
  except
    // On se trouve ici si l'enregistrement qu'on tente d'insérer en base existe
    // dû à un plantage précédent
    for i := 0 to T.Detail.count-1 do
      T.Detail[i].DeleteDB;
      Result := T.InsertDB(Nil);
  end;
end;

{JP 18/02/05 : Affichage des zones concernant la Banque prévisionnelle. Elles sont visibles si :
               1/ On est en impression et/ou comptabilisation de BOR et que la catégorie de
                  paiement est "LCR"
               2/ en Enca / Deca
 JP 09/08/05 : FQ 15815 : gestion de la banque prévisionnelle dans la remise en banque des BOR (smpDecBorDec)
{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREMP.AfficheBanquePrevi(OnArgOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  OkVisible : Boolean;
  Pop       : TPopupMenu;
begin
  {JP 12/01/06 : FQ 17299 : Suppression de EstSerie(S3)}
  OkVisible := (((smp in [smpDecBorEdt, smpDecBorEdtNC]) and (GetControlText('CATEGORIE') = 'LCR')) or
                (smp in [smpDecDiv, smpEncDiv, smpDecBorDec])) {and not EstSerie(S3)};

  {Gestion de la combo des banques prévisionnelles}
  SetControlVisible('BANQUEPREVI', OkVisible);
  SetControlVisible('TE_BANQUEPREVI__', OkVisible);

  SetPlusBanqueCp(GetControl('E_BANQUEPREVI'));

  {Gestion du popup qui permet l'affectation de la baqnue prévisionnelle}
  Pop := TPopupMenu(GetControl('POPUPMENU'));
  if Assigned(Pop) then
    if Assigned(Pop.Items[0]) then begin
      Pop.Items[0].Visible := OkVisible;
      if OnArgOk then begin
        Pop.Items[0].OnClick := AffecteBanqueClick;
        AddMenuPop(Pop, '', '');
      end;
    end;
end;

{JP 18/02/05 : On filtre sur les banques prévisionnelles si la catégorie de paiement est
               LCR (=> GetControlVisible('BANQUEPREVI')) et si l'on n'est pas par Lot
 JP 07/06/07 : FQ 15793 : Passage à un MultiValComboBox
{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREMP.WhereBqPrevi;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  if GetControlVisible('BANQUEPREVI') and not SelectOnLot then begin
    s := GetControlText('BANQUEPREVI');
    if (s <> '') and (Pos('<<', s) = 0) then
      s := 'E_BANQUEPREVI IN (' + GetClauseIn(s) + ')'
    else
      s := '';
  end;
  SetControlText('XX_WHEREBQPREVI', s);
end;

{---------------------------------------------------------------------------------------}
function TOF_CPGENEREMP.GetChamp(Q : THQuery; sChamp : string) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if Q = nil then Result := GetField(sChamp)
             else Result := TFMul(Ecran).Q.FindField(sChamp).AsVariant;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPGENEREMP.AffecteBanqueClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  L : TStringList; {FQ TRESO 10294}

    {----------------------------------------------------------------}
    procedure CreerTob(Q : THQuery = nil);
    {----------------------------------------------------------------}
    var
      F : TOB;
    begin
      {JP 27/09/05 : FQ TRESO 10294 : On ne traite pas les écritures de réglement}
      if L.IndexOf(GetChamp(Q, 'E_JOURNAL')) = -1 then begin
        F := TOB.Create('$$$$', T, -1);
        F.AddChampSupValeur('E_EXERCICE'     , GetChamp(Q, 'E_EXERCICE'));
        F.AddChampSupValeur('E_JOURNAL'      , GetChamp(Q, 'E_JOURNAL'));
        F.AddChampSupValeur('E_DATECOMPTABLE', GetChamp(Q, 'E_DATECOMPTABLE'));
        F.AddChampSupValeur('E_NUMEROPIECE'  , GetChamp(Q, 'E_NUMEROPIECE'));
        F.AddChampSupValeur('E_NUMLIGNE'     , GetChamp(Q, 'E_NUMLIGNE'));
        F.AddChampSupValeur('E_NUMECHE'      , GetChamp(Q, 'E_NUMECHE'));
        F.AddChampSupValeur('E_QUALIFPIECE'  , GetChamp(Q, 'E_QUALIFPIECE'));
        {Pour la mise à jour dans CPAFFECTEBQPREVI_TOF
        JP 21/07/06 : Modification de la fonction WhereEcritureTob : prefixe est devenu inutile
        F.AddChampSupValeur('PREFIXE', 'E');}
      end;
    end;

    {JP 27/09/05 : FQ TRESO 10294 : Chargement des journaux de banque
    {----------------------------------------------------------------}
    function ChargeListeJournaux : TStringList;
    {----------------------------------------------------------------}
    var
      Q : TQuery;
    begin
      Result := TStringList.Create;
      Q := OpenSQl('SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL = "BQE"', True);
      try
        while not Q.EOF do begin
          Result.Add(Q.FindField('J_JOURNAL').AsString);
          Q.Next;
        end;
      finally
        Ferme(Q);
      end;
    end;

var
  n : Integer;
begin
  {Soit rien n'est sélectionné, soit on est en mode inversé et tut est sélectionné}
  if (not TFmul(Ecran).FListe.AllSelected and (TFmul(Ecran).FListe.nbSelected = 0)) or
     (SwapSelect and TFmul(Ecran).FListe.AllSelected) then begin
    HShowMessage('1;' + Ecran.Caption + ';Veuillez sélectionner au moins une ligne.;W;O;O;O;', '', '');
    Exit;
  end;

  {JP 27/09/05 : FQ TRESO 10294 : Chargement des journaux de banque afin de pouvoir empécher
                 l'affectation de la banque prévisionnelle sur les écritures de réglement}
  L := ChargeListeJournaux;

  {Constitution de la Tob contenant la clef des écritures sélectionnées}
  T := TOB.Create('$$$$', nil, -1);
  try
    {Soit tout est sélectionné, soit on est en mode inversé et il n'y a aucune sélection}
    if TFmul(Ecran).FListe.AllSelected or
       (SwapSelect and (TFmul(Ecran).FListe.nbSelected = 0)) then begin
      TFMul(Ecran).Q.First;
      while not TFMul(Ecran).Q.EOF do begin
        CreerTob(TFMul(Ecran).Q);
        TFMul(Ecran).Q.Next;
      end;
    end

    else begin
      {Si l'on est en mode inversé ...}
      if SwapSelect then begin
        TFMul(Ecran).Q.First;
        {$IFDEF EAGLCLIENT}
        for n := 1 to TFmul(Ecran).FListe.RowCount - 1 do begin
          if not TFmul(Ecran).FListe.IsSelected(n) then
        {$ELSE}
        for n := 1 to TFmul(Ecran).Q.RecordCount do begin
          if not TFmul(Ecran).FListe.IsCurrentSelected then
        {$ENDIF EAGLCLIENT}
            CreerTob(TFMul(Ecran).Q);
          TFMul(Ecran).Q.Next;
        end;
      end

      {... Ou si l'on est en mode mode normal}
      else
        for n := 0 to TFmul(Ecran).FListe.nbSelected - 1 do begin
          TFmul(Ecran).FListe.GotoLeBookmark(n);
          CreerTob;
        end;
    end;

    TheTob := T;
    {Appel de la fiche d'affectation de la banque prévisionnelle}
    CP_AffecteBqPrevi(';');
    {Rafraîchissement de la liste}
    TFMul(Ecran).BCherche.Click;
  finally
    if Assigned(T) then FreeAndNil(T);
    if Assigned(L) then FreeAndNil(L);
  end;
end;
{ YMO 24/05/2006 Suppression de l'activation manuelle des listes suite à
la modif eAGL sui ramène bien le titre dans les fenêtres de recherche}

{FB 19/06/2006 Il ne s'agit pas du titre des fenêtres de recherche mais d'un titre
générique}

{b fb 27/10/2005 FQ12487}
procedure TOF_CPGENEREMP.SelectionElipsisClick(Sender : TObject);
begin
  if isEnc(smp) then
    LookUpList(TControl(Sender),'Recherche d''un compte collectif','GENERAUX','G_GENERAL','G_LIBELLE',
    '(G_SUIVITRESO="ENC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" OR G_LETTRABLE="X")','G_GENERAL',true,-1)
  else
    LookUpList(TControl(Sender),'Recherche d''un compte collectif','GENERAUX','G_GENERAL','G_LIBELLE',
    '(G_SUIVITRESO="DEC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" OR G_LETTRABLE="X")','G_GENERAL',true,-1);
end;

procedure TOF_CPGENEREMP.AuxiliaireElipsisClick(Sender : TObject);
begin
  if isEnc(smp) then
    LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
    '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")','T_AUXILIAIRE',true,-1)
  else
    LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
    '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")','T_AUXILIAIRE',true,-1);
end;
{e fb 27/10/2005 FQ 12487}


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 18/05/2006
Modifié le ... :   /  /
Description .. : FQ17447 Pb de mise à jour de E_GENERAL quand on
Suite ........ : coche TIC/TID
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENEREMP.FFiltresChange(Sender: TObject);
begin
 InChargement := True;
 TFMul(Ecran).FFiltresChange(FFiltres);
 InChargement := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 01/06/2006
Modifié le ... :   /  /    
Description .. : Evènement déclenché à la fin du chargement de l'écran ;
Suite ........ : permet de gérer le lancement intempestif du 'FTICChange'
Suite ........ : à l'ouverture de l'écran
Mots clefs ... : FQ17447
*****************************************************************}
procedure TOF_CPGENEREMP.OnAfterSelectFiltreEcran;
begin
If Not InChargement then FinChargeEcran:=True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENEREMP.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


function TOF_CPGENEREMP.YEnATrop : boolean ;
var lQCount : TQuery ;
    lQte    : integer ;
begin
  result := True ;

  // Mode inversé
  if swapSelect then
    begin
    if (not FListe.AllSelected) then
      begin
      lQCount := OpenSQL( 'SELECT COUNT(*) QTE FROM SUIVIMP ' + RecupWhereCritere(Pages) , True ) ;
      if not lQCount.Eof
        then lQte := lQCount.FindField('QTE').AsInteger
        else lQte := 0 ;
      Ferme( lQCount ) ;
      lQte := lQte - FListe.NbSelected ;
      if lQte > MAXSELECTECHE then Exit ;
      end ;
    end
  // Mode normal
  else
    begin
    if ( FListe.NbSelected > MAXSELECTECHE ) then Exit ;
    if FListe.AllSelected then
      begin
      lQCount := OpenSQL( 'SELECT COUNT(*) QTE FROM SUIVIMP ' + RecupWhereCritere(Pages) , True ) ;
      if not lQCount.Eof
        then lQte := lQCount.FindField('QTE').AsInteger
        else lQte := 0 ;
      Ferme( lQCount ) ;
      if lQte > MAXSELECTECHE then Exit ;
      end ;
    end ;

  result := False ;

end;


Initialization
  registerclasses ( [ TOF_CPGENEREMP ] ) ;
end.

