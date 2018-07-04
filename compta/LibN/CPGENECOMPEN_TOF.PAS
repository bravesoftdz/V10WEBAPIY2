{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPGENECOMPEN ()
Mots clefs ... : TOF;CPGENECOMPEN
*****************************************************************}
Unit CPGENECOMPEN_TOF ;

Interface

Uses Windows,
     {$IFDEF VER150} Variants, {$ENDIF}
     StdCtrls,
     Controls,
     Classes,
     Grids,           // pour le TGridDrawState
     ExtCtrls,        // pour le TPanel
     Graphics,        // pour clRed
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF},
     mul,
     FE_MAIN,
     dbGrids,
     EdtREtat,        // pour LanceEtatTob
{$else}
     eMul,
     maineagl,
     UtileAGL,        // pour LanceEtatTob
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     Ent1,
     HEnt1,
     HMsgBox,
     UTOF,
     uTob,            // pour TOB
     uLibWindows,     // pour TraductionTHMultiValComboBox
     SaisBor,         // pour lanceSaisieFolio
     Saisie,          // pour TrouveEtLanceSaisie
     UtilPGI,         // pour Resolution
     EncUtil,         // pour CatToMP
     HDB,             // pour THDBGrid
     HQry,            // pour THQuery
     MulSMPUtil,      // pour SmpToStr
     GenerMP,         // pour tGenereMP
     menus,           // pour le TMenuItem
     HTB97,           // pour TToolBarButton97
     Lookup,          // pour LookupList
     SaisComm,        // pour TOBToIdent
     CPLETREGUL_TOF,  // pour RegulLettrageMP
     UtilSais,        // pour CaseNata
     SaisUtil,        // pour TSuiviMP
     ParamSoc,		    // GetParamSocSecur YMO
     RepDevEur;       // pour ChangeLeTauxDevise FP 17/07/2006

Procedure GenererCompensation(smp : TSuiviMP);

type
  TMode = (mdEdition, mdGeneration);

{$IFDEF EAGLCLIENT}
const dgEditing = goEditing;
{$ENDIF}
  
Type
  TOF_CPGENECOMPEN = Class (TOF)
  private
    LesCritsRupt   : TStringList ;
    SwapSelect     : Boolean ;		// Gestion du mode selection inversé
    gszCaption     : String;      // Pour les messages
    bgFetchAll     : Boolean;     // Pour le bouton Sélection inversion et Select All
    AllSelecting   : boolean;
    FEcrLibelle    : string;      {FP 25/04/2006 FQ17982}

    Q              : THQuery ;
{$IFDEF EAGLCLIENT}
    FListe         : THGrid ;
{$ELSE}
    FListe         : THDBGrid ;
{$ENDIF}
    Pages          : TPageControl ;
    FFiltres       : THValComboBox ;
    lMenuItem      : TMenuItem ;

    // Composants géés manuellement // commun
    CATEGORIE      : THValComboBox ;

    // onglet standard
    Auxiliaire     : THCritMaskEdit;

    // Boutons Toolbar
    BSwapSelect     : TToolbarButton97 ;
    BChangeTaux     : TToolbarButton97;    {FP 17/07/2006}
    AfficPieComp : TCheckBox; //fb 15/09/2006 FQ18581

    // Messages
    HDiv           : TStringList ;
    HM             : TStringList ;

    // Evt Ancêtre
{$IFDEF EAGLCLIENT}
{$ELSE}
    FListeDrawColumnCellParent  : TDrawColumnCellEvent ;
{$ENDIF}
    FormKeyDownParent         : TKeyEvent ;
    bSelectAllClickParent     : TNotifyEvent ;
    BParamListeClickParent    : TNotifyEvent ;
    CBDeviseChangeParent      : TNotifyEvent;    {FP 17/07/2006}

    // init des messagebox et menuzoom
    procedure InitVariables;
    procedure InitMsgBox;
    procedure InitHelpContext;
    procedure InitComposants;
    procedure InitEvenements;
    procedure InitContexte;
    procedure InitSelectionLot;
    procedure InitCriteres;
    procedure InitConditions;
    procedure InitCATEGORIE;
    procedure PreciseMP;
// CA - 31/08/2007 - Inutile
//    procedure ReinitWhereNatCpt(RAZ : Boolean) ;
    procedure InitCriteresPourLot;
    Procedure AfficheSoldeSelection;
    Function  CalculSoldeSelection : Double ;
    Procedure ChangeModeGrille;
    procedure UpdateConditions;
    Function  ConstitueLots(Mode: TMode) : Boolean ;
    Function  ConstitueOrigine(ValsRupt : String; Mode: TMode; RuptEtab: TStrings; var DeviseDest: RDEVISE): Boolean ;
    procedure ConstitueDest(ValsRupt : String; TOBLot, TOBDest : TOB; DeviseDest: RDEVISE);
    function  LInsertDB(T : Tob) : Boolean;
    function  CalculSoldeARegler(TOBL: TOB; DeviseDest: RDEVISE): Double;
    function  IsMultiTiers: Boolean;
    function  LanceSaisie(Etablissement: String): Boolean;
    procedure LanceEdition(TOBEdt: TOB);
    function  VerifSelectionGrille: Boolean;
    function  InitParmsMP: Boolean;
    procedure InitRDevise(TOBL: TOB; DateCompta: TDateTime; var DeviseDest: RDEVISE);
    procedure RAZNomLot(Etablissement: String);

    // EVT FILTRES
    procedure E_EXERCICEChange(Sender: TObject);
    procedure CATEGORIEChange(Sender: TObject);
    procedure _E_NATUREPIECEChange(Sender: TObject);

    // EVT BOUTONS ANCETRE MUL
    procedure bSelectAllClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    // EVT BOUTONS
    procedure BSwapSelectClick(Sender: TObject);
    procedure BLotEcrClick(Sender: TObject);
    procedure BEditionClick(Sender: TObject);
    procedure BChangeTauxClick(Sender: TObject);       {FP 17/07/2006}
    procedure CBDeviseChange(Sender: TObject);       {FP 17/07/2006}

    // EVT Grille
    procedure FListeDblClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
    procedure FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FListeSortColumn(Sender : TObject);
{$ENDIF}
    procedure FListeRowEnter(Sender: TObject);
    procedure FListeFlipSelection(Sender: TObject);
    // EVT Form
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);   // ok
    procedure FetchLesTous;
    procedure AuxiliaireElipsisClick(Sender : TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  public
    ParmsMP                    : tGenereMP ;
    TOBLot,TOBOrig,TOBDest     : TOB ;
    smp                        : TSuiviMP ;
    SelectOnLot                : Boolean ;
    NomCritLot, NomCritReste   : String ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPObjetGen,
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  CPTPAYEURFACTURE_TOF // Pour ComposantToCritere
  , UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

Procedure GenererCompensation(smp : TSuiviMP);
var lStArgs : String ;
begin
  if not TCompensation.IsCompensation then
    begin
    PGIInfo(TraduireMemoire('La gestion de la compensation n''est pas active')); {FP 17/07/2006}
    Exit;
    end;

  lStArgs := SmpToStr(smp) ;
  AGLLanceFiche('CP', 'CPGENECOMPEN_MUL', '', '', lStArgs);
end;

procedure TOF_CPGENECOMPEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPGENECOMPEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPGENECOMPEN.OnUpdate ;
begin
  Inherited ;
  AfficheSoldeSelection ;
  FListeRowEnter(nil); // 10930
end ;

procedure TOF_CPGENECOMPEN.OnLoad ;
begin
  Inherited ;
  UpdateConditions;
end ;

procedure TOF_CPGENECOMPEN.OnArgument (S : String ) ;
var vStSmp  : String ;
begin
  Inherited ;

  vStSmp                    := S ;
  smp                       := StrToSmp(vStSmp) ;
  TFMul(Ecran).SetDBListe(AttribListe(smp));

  bgFetchAll := True;

  if smp = smpCompenCli then
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
  InitSelectionLot ;

  // Init des critères du mul
  InitCriteres ;
  InitConditions ;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    Auxiliaire.OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPGENECOMPEN.OnClose ;
begin
  Inherited ;

  // Libération des TOB
  FreeAndNil(TOBLot) ;
  FreeAndNil(TOBOrig) ;
  FreeAndNil(TOBDest) ;

  // Libération des TList
  LesCritsRupt.Clear ;
  LesCritsRupt.Free ;

  FreeAndNil(HM) ;
  FreeAndNil(HDiv) ;
end ;

procedure TOF_CPGENECOMPEN.OnDisplay () ;
begin
  {b FP 25/04/2006 FQ17945}
  if THValComboBox(GetControl('FFiltres')).Text <> '' then
    TFmul(Ecran).CritModified := False;
  {e FP 25/04/2006}
  FListeRowEnter(nil);
  FListe.SetFocus ;
end ;

procedure TOF_CPGENECOMPEN.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPGENECOMPEN.InitVariables;
begin
  // Init TOB
  TOBLot        := TOB.Create('',Nil,-1) ;
  TOBOrig       := TOB.Create('',Nil,-1) ;
  TOBDest       := TOB.Create('',Nil,-1) ;

  // Init TList
  LesCritsRupt            := TStringList.Create ;
  LesCritsRupt.Sorted     := True ;
  LesCritsRupt.Duplicates := dupIgnore ;

  // Init Booleens
  SwapSelect  := False ;   // Gestion du mode "selection inversé", non opérationnel par défaut

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
  FEcrLibelle := '';      {FP 25/04/2006 FQ17982}
end;

procedure TOF_CPGENECOMPEN.InitMsgBox;
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
  HM.Add('12;?Caption?;Le journal de compensation doit être en multi devise.;W;O;O;O;');

  // Init HDIV
  HDiv := TStringList.Create ;
  HDiv.Add('Gestion de la Compensation fournisseurs');
  HDiv.Add('Gestion de la Compensation clients');
  HDiv.Add('Compensation');     {2}
  HDiv.Add('Règlements Fournisseurs - Solde  des comptes pour la compensation');
  HDiv.Add('Règlements Clients - Solde  des comptes pour la compensation');
  HDiv.Add('Traiter les soldes débiteurs');                     {5}    {FP 25/04/2006 FQ17781}
  HDiv.Add('Traiter les soldes créditeurs');                           {FP 25/04/2006 FQ17781}
  HDiv.Add('TOTAL A REGLER');
  HDiv.Add('TOTAL A RECEVOIR');                                 {8}
end;

procedure TOF_CPGENECOMPEN.InitHelpContext;
begin
  TFMul(Ecran).HelpContext  := AttribHelp(smp,TRUE) ;
end;

procedure TOF_CPGENECOMPEN.InitComposants;
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
  CATEGORIE       := THValComboBox(GetControl('CATEGORIE', True)) ;   // C
  Auxiliaire      := THCritMaskEdit(GetControl('_E_AUXILIAIRE', true));  {FP 15/04/2006 FQ17967}

  // onglet compléments

  // Boutons Toolbar
  BSwapSelect     := TToolbarButton97(GetControl('BSwapSelect', True)) ;
  BChangeTaux     := TToolbarButton97(GetControl('BChangeTaux', True));    {FP 17/07/2006}
  AfficPieComp    := TCheckBox(GetControl('AFFICPIECOMP', true));   //fb 15/09/2006 FQ18581
end;

procedure TOF_CPGENECOMPEN.InitEvenements;
Var
  lBouton   : TButton ;
begin

{$IFDEF EAGLCLIENT}
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
  // -> Imprimer
  TButton(GetControl('BEDITION',True)).OnClick    := BEditionClick;  {FP 25/04/2006 FQ17953}

  // EVT Form
  if Assigned(TFMul(Ecran).OnKeyDown)
    then FormKeyDownParent := TFMul(Ecran).OnKeyDown
    else FormKeyDownParent := nil ;
  TFMul(Ecran).OnKeyDown := FormKeyDown ;

  // Ftn spécifique
  TToolbarButton97(GetControl('BLotEcr', True)).OnClick       := BLotEcrClick ;
  BSwapSelect.OnClick   := BSwapSelectClick ;
  BChangeTaux.OnClick   := BChangeTauxClick;    {FP 17/07/2006}

  // onglet standard
  THValComboBox(GetControl('E_EXERCICE',  True)).OnChange := E_EXERCICEChange ;
  CATEGORIE.OnChange     := CATEGORIEChange ;
  Auxiliaire.OnElipsisClick := AuxiliaireElipsisClick;
  // onglet écritures
  THMultiValComboBox(GetControl('_E_NATUREPIECE', True)).OnChange := _E_NATUREPIECEChange ;
  // onglet critères (avec banque / lot...)
  THValComboBox(GetControl('CATEGORIE1', True)).OnChange          := CATEGORIEChange ;
  {b FP 17/07/2006 Changement de devise}
  CBDeviseChangeParent  := THValComboBox(GetControl('E_DEVISE', True)).OnChange;
  THValComboBox(GetControl('E_DEVISE', True)).OnChange := CBDeviseChange;
  {e FP 17/07/2006}

  {Retire l'évènement OnClick pour ne pas modifier TFMul(Ecran).CritModified}
  TCheckBox(GetControl('FECRDEBITCREDIT1', True)).OnClick := nil;
  TCheckBox(GetControl('FECRDEBITCREDIT', True)).OnClick := nil;
end;

procedure TOF_CPGENECOMPEN.InitContexte;
begin
  SetControlProperty('E_NOMLOT',  'PLUS', AttribPlus(smp) );
  SetControlProperty('E_NOMLOT_', 'PLUS', AttribPlus(smp) );
  Case smp of
     smpCompenCli : Ecran.Caption := HDiv[1] ;
     smpCompenFou : Ecran.Caption := HDiv[0] ;
   end ;

  if smp = smpCompenCli then
    SetControlProperty('_E_AUXILIAIRE','DataType','TZTTOUTDEBIT')         {FP 15/04/2006 FQ17967}
  else
    SetControlProperty( '_E_AUXILIAIRE',  'DataType', 'TZTTOUTCREDIT' );  {FP 15/04/2006 FQ17967}

  gszCaption  := Ecran.Caption;  // Pour les messages
  UpdateCaption(Ecran) ;

  InitCATEGORIE ;

  case smp of
    smpCompenCli: TCheckBox(GetControl('FECRDEBITCREDIT1', True)).Caption := HDiv[6];
    smpCompenFou: TCheckBox(GetControl('FECRDEBITCREDIT1', True)).Caption := HDiv[5];
    end;
  TCheckBox(GetControl('FECRDEBITCREDIT', True)).Caption := TCheckBox(GetControl('FECRDEBITCREDIT1', True)).Caption;
end;

procedure TOF_CPGENECOMPEN.InitSelectionLot;
begin
  // Page de sélection sur lot non apparrente au lancement
  TTabSheet(GetControl('PBanqueLot',True)).TabVisible := FALSE ;
  SetActiveTabSheet('PCritere') ;

  // Page de sélection sur lot non apparrente au lancement
  if smp=smpCompenCli
    then if (not VH^.CCMP.LotCli)
          then BLotEcrClick(Nil)
          else
    else if (Not VH^.CCMP.LotFou)
          then BLotEcrClick(Nil) ;
  selectOnLot   := FALSE ;
end;

procedure TOF_CPGENECOMPEN.InitCriteres;
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

  // pas de gestion par lot
  TTabSheet(GetControl('PBanqueLot')).TabVisible := False ;
  Pages.ActivePage := TTabSheet(GetControl('PCritere')) ;

  // Nature de pièce
  THMultiValComboBox(GetControl('_E_NATUREPIECE', True)).SelectAll ;
end;

procedure TOF_CPGENECOMPEN.InitConditions;
var
    (*
    StXP, StXP2 : String ;
    StXN, StXN2 : String ;
    *)
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
  SetControlText('XX_WHEREAN', St ) ;

  // Conditions nature de compte et montants        XX_WHERENATCPT
  // CA - 31/08/2007 - Inutile
  // ReinitWhereNatCpt(FALSE) ;                     // XX_WHEREMONTANT

  // Condition sur le numéro de traite              XX_WHERETRACHQ
  // CA - 31/08/2007 - Inutile
  //SetControlText('XX_WHERETRACHQ', '') ;

  // Condition sur ???                              XX_WHEREMP
  SetControlText('XX_WHEREMP', '') ;

  // Condition sur les lots                         XX_WHERELOT
  SetControlText('XX_WHERELOT', TCompensation.GetChampPlan+'<>""' ) ;

  // XX_WHEREENC
  SetControlText('XX_WHEREENC','');

  // XX_WHEREMONTANT
  (* - CA - 31/08/2007 - Inutile
  StXP  := StrFPoint(  9 * Resolution( V_PGI.OkDecV + 1 ) ) ;
  StXN  := StrFPoint( -9 * Resolution( V_PGI.OkDecV + 1 ) ) ;
  StXP2 := StrFPoint(  9 * Resolution( V_PGI.OkDecE + 1 ) ) ;
  StXN2 := StrFPoint( -9 * Resolution( V_PGI.OkDecE + 1 ) ) ;
  St    := '(E_DEBIT+E_CREDIT-E_COUVERTURE not between ' + StXN + ' AND ' + StXP + ')';
  SetControlText('XX_WHEREMONTANT', St ) ;
  *)
end;

procedure TOF_CPGENECOMPEN._E_NATUREPIECEChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
end;

procedure TOF_CPGENECOMPEN.AuxiliaireElipsisClick(Sender: TObject);
begin
  if smp=smpCompenCli then
    LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
    '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI")','T_AUXILIAIRE',true,-1)
  else
    LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
    '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU")','T_AUXILIAIRE',true,-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGENECOMPEN.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPGENECOMPEN.BLotEcrClick(Sender: TObject);
Var i       : Integer ;
begin
  SelectOnLot := Not SelectOnLot ;

  // Visibiilté des onglets
  for i:=0 to Pages.PageCount-1 do
    Pages.Pages[i].TabVisible := Not SelectOnLot ;
  TTabSheet(GetControl('PAvance', True)).TabVisible     := True ;
  TTabSheet(GetControl('PSQL', True)).TabVisible        := True ;
  TTabSheet(GetControl('INVISIBLE',True)).TabVisible    := False ;
  TTabSheet(GetControl('PBANQUELOT',True)).TabVisible   := SelectOnLot ;

  // Passage à la sélection par lot
  if SelectOnLot then
    begin
    SetActiveTabSheet('PBANQUELOT') ;
    InitCriteresPourLot ;
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

procedure TOF_CPGENECOMPEN.BOuvrirClick(Sender: TObject);
var
  k, i:              Integer;
  DeviseDest:        RDEVISE;
  OkInsert:          Boolean;
  RuptEtab:          TStringList;
begin
  if not VerifSelectionGrille then
    Exit;

  ParmsMP.ChampsRupt := TCompensation.GetChampPlan;

  if Not ConstitueLots(mdGeneration) then Exit ;
  if TOBLot.Detail.Count<=0 then Exit ;
  if not InitParmsMP then Exit;

  RuptEtab := TStringList.Create;
  RuptEtab.Sorted     := True ;
  RuptEtab.Duplicates := dupIgnore ;

  OkInsert := False;
  for k:=0 to LesCritsRupt.Count-1 do
    begin
    if ConstitueOrigine(LesCritsRupt[k], mdGeneration, RuptEtab, DeviseDest) then
      begin
      {Rupture par établissement}
      for i:=0 to RuptEtab.Count-1 do
        begin
        ConstitueDest(RuptEtab[i], TOBOrig, TOBDest, DeviseDest) ;
        if TOBDest.Detail.Count > 0 then     {FP 25/04/2006 FQ17948}
          begin
          if LInsertDB(TOBDest) then
             OkInsert := LanceSaisie(RuptEtab[i]);
          if OkInsert then
            RAZNomLot(RuptEtab[i]);
          end;
        end;
      end;
    end;
  RuptEtab.Free;

  for k:=0 to TOBLot.Detail.Count-1 do
    begin
    if OkInsert and (not (TOBLot.Detail[k].GetValue('E_DEVISE') = V_PGI.DevisePivot)) then
      begin
      {Provoque un 'index de grille hors limite' car la liste REGULLETEURO n'a aucun champ
      PGIBox('Attention certaines écritures sont en devise.' + #10 + #13 + 'Vous devez comptabiliser les différences de change.');
      if smp = smpCompenCli then
        RegulLettrageMP(False,False,prClient)
      else
        RegulLettrageMP(False,False,prFournisseur);}
      break;
      end;
    end;
  TFMul(Ecran).BChercheClick(Nil) ;
end;

procedure TOF_CPGENECOMPEN.BParamListeClick(Sender: TObject);
begin
  // Annule le CTRL+M avant de passer en paramètrage de la liste
  if (dgEditing in FListe.Options) then
    SwapModeGrid(FListe,Q) ;
  if Assigned(BParamListeClickParent) then
    BParamListeClickParent(Sender) ;
end;

procedure TOF_CPGENECOMPEN.bSelectAllClick(Sender: TObject);
begin
  AllSelecting := true;

  FetchLesTous;
  if Assigned(bSelectAllClickParent) then
    bSelectAllClickParent(Sender);

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

procedure TOF_CPGENECOMPEN.BSwapSelectClick(Sender: TObject);
begin
  // Si tout sélectionné, on déselectionne tout
  FetchLesTous;
  if FListe.AllSelected then bSelectAllClick(nil) ;
  // Message d'avertissement
  if not SwapSelect then
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

procedure TOF_CPGENECOMPEN.CATEGORIEChange(Sender: TObject);
begin
  TFMul(Ecran).CritModified := True;
  PreciseMP ;
end;

procedure TOF_CPGENECOMPEN.E_EXERCICEChange(Sender: TObject);
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

procedure TOF_CPGENECOMPEN.FetchLesTous;
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

{$IFDEF EAGLCLIENT}
procedure TOF_CPGENECOMPEN.FListeDblClick(Sender: TObject);
var
  TOBListe, TOBLigne : TOB ;
  sMode : String ;
begin
  {$IFNDEF IMP}
  if (Q.TQ.Detail.Count = 0) then Exit ;
  Q.TQ.Seek(FListe.row - 1) ;
  sMode := Q.FindField('E_MODESAISIE').AsString ;
  if ((sMode<>'') and (sMode<>'-'))
    then
      PGIInfo(TraduireMemoire('Saisie borderau non implémentée en mode CWAS'), Ecran.Caption )  {FP 17/07/2006}
      // A FAIRE Voir SAISBOR.PAS
    else
      begin
      TOBListe := TOB.Create('Liste Ecriture',nil,-1) ;
      TOBLigne := Q.TQ.Detail[FListe.row-1] ;
      TOB.Create('ECRITURE',TOBListe,-1) ;
      TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
      TrouveEtLanceSaisie( TOBListe, taConsult, 'N') ;
      TOBListe.ClearDetail ;
      FreeAndNil( TOBListe ) ;
      end ;
  {$ENDIF}
end;
{$ELSE}
procedure TOF_CPGENECOMPEN.FListeDblClick(Sender: TObject);
var sMode : String ;
begin
  {$IFNDEF IMP}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  sMode := Q.FindField('E_MODESAISIE').AsString ;
  if ((sMode<>'') and (sMode<>'-'))
    then
      LanceSaisieFolio(Q,taConsult)
    else
      begin
      TrouveEtLanceSaisie(Q,taConsult,'N') ;
      end ;
  {$ENDIF}
end;
{$ENDIF}

{$IFNDEF EAGLCLIENT}
procedure TOF_CPGENECOMPEN.FListeSortColumn(Sender: TObject);
begin
  FListe.ClearSelected;
end;

procedure TOF_CPGENECOMPEN.FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var sRIB, St : String ;
begin
  // partie commune
  if Assigned(FListeDrawColumnCellParent) then
    FListeDrawColumnCellParent(Sender,Rect,DataCol,Column,State);
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  St:='E_AUXILIAIRE';
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
end;
{$ENDIF}

procedure TOF_CPGENECOMPEN.FListeFlipSelection(Sender: TObject);
begin
  if ((not AllSelecting) and (FListe.AllSelected)) then
    FListe.AllSelected := false;
  AfficheSoldeSelection ;
end;

procedure TOF_CPGENECOMPEN.FListeRowEnter(Sender: TObject);
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

procedure TOF_CPGENECOMPEN.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  Case Key of
{AE}   69 : BEGIN // Active-désactive les escomptes sur une ligne
	    If Shift=[ssAlt] then // ALT-E sur une ligne
              BEGIN
              Key:=0 ;
              FListe.Refresh ;
              END ;
            END ;
{^M}   77 : if (Shift=[ssCtrl]) then
              BEGIN
              Key:=0 ;
              ChangeModeGrille ;
              END ;
       VK_F10 :
            begin
            BOuvrirClick(nil);
            end;
  else
    begin
    if Assigned(FormKeyDownParent) then
      FormKeyDownParent(Sender, Key, Shift);
    end;
  end ;
end;

procedure TOF_CPGENECOMPEN.InitCATEGORIE;
begin
  ParmsMP.Cat := CATEGORIE.Value ;
  PreciseMP ;
end;

procedure TOF_CPGENECOMPEN.PreciseMP;
var E_MODEPAIE : THValComboBox ;
begin
  E_MODEPAIE := THValComboBox(GetControl('E_MODEPAIE',True)) ;
  if Categorie.itemIndex = 0
    then CatToMP ( '',              E_MODEPAIE.Items, E_MODEPAIE.Values, tslAucun, True)
    else CatToMP ( CATEGORIE.Value, E_MODEPAIE.Items, E_MODEPAIE.Values, tslAucun, True ) ;
  E_MODEPAIE.itemIndex := 0 ;
end;

(*
  // CA - 31/08/2007 - Inutile
procedure TOF_CPGENECOMPEN.ReinitWhereNatCpt(RAZ: Boolean);
begin
  SetControlText('XX_WHERENATCPT','E_AUXILIAIRE<>""'+
      ' AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC")');
end;
*)

procedure TOF_CPGENECOMPEN.InitCriteresPourLot;
begin
  SetControlText('E_GENERAL',     '') ;
  SetControlText('_E_AUXILIAIRE',  '') ;    {FP 15/04/2006 FQ17967}
  SetControlText('E_DEVISE',      '') ;

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

  SetControlText('E_DEBIT',         '') ;
  SetControlText('E_DEBIT_',        '') ;
  SetControlText('E_CREDIT',        '') ;
  SetControlText('E_CREDIT_',       '') ;
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

procedure TOF_CPGENECOMPEN.AfficheSoldeSelection;
Var Solde : Double;
begin
  // Calcul du solde
  Solde := CalculSoldeSelection;
  // Ajout affichage du cumul
  if Solde < 0
    then Ecran.Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(-1*Solde,0,V_PGI.OkDecV,'',True) + ' C)'
    else Ecran.Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(Solde,0,V_PGI.OkDecV,'',True) + ' D)';
  // Affichage de mode "selection inversé" si besoin
  if SwapSelect then
    Ecran.Caption := Ecran.Caption + ' [SELECTION INVERSEE] ';

  // Raffraichissement Titre
  UpdateCaption(Ecran) ;
end;

function TOF_CPGENECOMPEN.CalculSoldeSelection: Double;
Var i               : Integer ;
    debit, credit   : Double ;
    PCumul          : TPanel ;
begin
  Result := 0 ;
  PCumul := TPanel( GetControl('PCumul', True) ) ;
  if (PCumul.FindChildControl('__QRYPCumul_E_DEBIT') = nil)
  	 or (PCumul.FindChildControl('__QRYPCumul_E_CREDIT') = nil) then Exit ;

  if FListe.AllSelected then
    begin
    debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).text) ;
    credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).text) ;
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
      Result := Result + Q.FindField('E_DEBIT').asFloat - Q.FindField('E_CREDIT').asFloat ;
      END ;

    if SwapSelect then  // en mode inversé, on retranche le cumul des lignes sélectionnée du solde total
      begin
      debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).text ) ;
      credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).text ) ;
      Result := (debit - credit) - Result ;
      end ;
    end ;
end;

procedure TOF_CPGENECOMPEN.ChangeModeGrille;
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
{$ENDIF EAGLCLIENT}

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
{$ELSE}
  // svg des donné pour retrouvé la ligne !
  if (Q.FindField('E_JOURNAL') <> nil) then RJal := Q.FindField('E_JOURNAL').AsString else Pb := true;
  if (Q.FindField('E_DATECOMPTABLE') <> nil) then RExo := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) else Pb := true;
  if (Q.FindField('E_DATECOMPTABLE') <> nil) then RDate := Q.FindField('E_DATECOMPTABLE').AsDateTime else Pb := true;
  if (Q.FindField('E_NUMEROPIECE') <> nil) then RNumP := Q.FindField('E_NUMEROPIECE').AsInteger else Pb := true;
  if (Q.FindField('E_NUMLIGNE') <> nil) then RNumL := Q.FindField('E_NUMLIGNE').AsInteger else Pb := true;
  if (Q.FindField('E_NUMECHE') <> nil) then RNumEche := Q.FindField('E_NUMECHE').AsInteger else Pb := true;
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
end;

procedure TOF_CPGENECOMPEN.UpdateConditions;
var szTemp, szLib : String;
    {b FP 25/04/2006 FQ17967}
    ChpClient:      String;
    ChpFourn:       String;
    ChpCorres:      String;
    {e FP 25/04/2006}
begin
  // XX_WHEREMONTANT et XX_WHEREAN sont initialisé une fois
  //  pour toute dans le initConditions

  // XX_WHEREENC
  TraductionTHMultiValComboBox(THMultiValComboBox(GetControl('_E_NATUREPIECE')), szTemp, szLib, 'E_NATUREPIECE');

  {b FP 25/04/2006 FQ17939 Ajoute une condition OR pour les filtres sur le débit et le crédit}
  if (szTemp <> '') and ((GetControlText('_E_DEBIT_') <> '') or (GetControlText('_E_CREDIT_') <> '')) then
    sztemp := sztemp+' AND';
  if (GetControlText('_E_DEBIT_') <> '') and (GetControlText('_E_CREDIT_') <> '') then
    sztemp := sztemp+' (';
  if (GetControlText('_E_DEBIT_') <> '') then
    sztemp := sztemp+' (E_DEBIT>='+GetControlText('_E_DEBIT')+' AND E_DEBIT<='+GetControlText('_E_DEBIT_')+')';
  if (GetControlText('_E_DEBIT_') <> '') and (GetControlText('_E_CREDIT_') <> '') then
    sztemp := sztemp+' or';
  if (GetControlText('_E_CREDIT_') <> '') then
    sztemp := sztemp+' (E_CREDIT>='+GetControlText('_E_CREDIT')+' AND E_CREDIT<='+GetControlText('_E_CREDIT_')+')';
  if (GetControlText('_E_DEBIT_') <> '') and (GetControlText('_E_CREDIT_') <> '') then
    sztemp := sztemp+')';
  SetControlText('XX_WHEREENC',sztemp);
  {e FP 25/04/2006}

  {b FP 25/04/2006 FQ17967}
  sztemp    := GetControlText('XX_WHEREENC');
  if smp = smpCompenCli then
    begin
    ChpClient := 'CR_LIBRETEXTE1';
    ChpFourn  := 'CR_LIBRETEXTE2';
    end
  else
    begin
    ChpClient := 'CR_LIBRETEXTE2';
    ChpFourn  := 'CR_LIBRETEXTE1';
    end;
  ChpCorres := TCompensation.GetChampPlan;
  if (szTemp <> '') and (GetControlText('_E_AUXILIAIRE')<>'') then
    sztemp := sztemp+' AND';
  if GetControlText('_E_AUXILIAIRE')<>'' then
    sztemp := sztemp+' (T_AUXILIAIRE LIKE "'+GetControlText('_E_AUXILIAIRE')+'%"'+  {FP 13/07/2006 FQ18589}
                     ' or T_AUXILIAIRE in (select '+ChpFourn+' from CORRESP'+       {FP 13/07/2006 FQ18589}
      ' WHERE '+ChpCorres+' = CR_CORRESP AND CR_TYPE = "AU'+IntToStr(TCompensation.GetNumPlanCorrespondance)+'"'+
      ' AND '+ChpClient+' LIKE "'+GetControlText('_E_AUXILIAIRE')+'%"))';
  SetControlText('XX_WHEREENC',sztemp);
  {e FP 25/04/2006}

  // XX_WHERELOT
  if SelectOnLot
    then SetControlText('XX_WHERELOT', 'E_NOMLOT<>""' )
    else SetControlText('XX_WHERELOT', '' ) ;
  SetControlText('XX_WHERELOT', GetControlText('XX_WHERELOT') +
    ' AND '+TCompensation.GetChampPlan+'<>""');

  // XX_WHERENATCPT
  // CA - 31/08/2007 - Inutile
  // ReinitWhereNatCpt(True);

  // XX_WHEREMP
  if SelectOnLot
    then LaCategorieChange( THValCombobox( GetControl('CATEGORIE1', True) ),
                            THEdit(GetControl('XX_WHEREMP',True)),
                            tslAucun)
    else LaCategorieChange( THValCombobox( GetControl('CATEGORIE', True) ),
                            THEdit(GetControl('XX_WHEREMP',True)),
                            tslAucun) ;
{b fb 15/09/2006 FQ18581}
  if not (AfficPieComp.State=cbGrayed) then begin
    if (AfficPieComp.State=cbChecked) then
      sztemp := sztemp + ' AND EXISTS ( '
    else
      sztemp := sztemp + ' AND NOT EXISTS ( ';

    if TCompensation.IsPlanCorresp(1) then
      sztemp := sztemp + ' SELECT T1.T_CORRESP1 '
    else
      sztemp := sztemp + ' SELECT T1.T_CORRESP2 ';

    sztemp := sztemp + ' FROM ECRITURE E1 LEFT JOIN TIERS T1 ON (T_AUXILIAIRE=E_AUXILIAIRE)';
    szTemp := szTemp + ' LEFT JOIN CORRESP ';

    if TCompensation.IsPlanCorresp(1) then
      sztemp := sztemp + ' ON T1.T_CORRESP1=CR_CORRESP AND CR_TYPE="AU1" '
    else
      sztemp := sztemp + ' ON T1.T_CORRESP2=CR_CORRESP AND CR_TYPE="AU2" ';

    sztemp := sztemp + RecupWhereCritere(Pages);
    sztemp := StringReplace(szTemp,'(E_','(E1.E_',[rfReplaceAll]);
    sztemp := StringReplace(szTemp,'(T_','(T1.T_',[rfReplaceAll]);
    sztemp := StringReplace(szTemp,' E_',' E1.E_',[rfReplaceAll]);
    sztemp := StringReplace(szTemp,' T_',' T1.T_',[rfReplaceAll]);
    sztemp := sztemp +' AND (   (TIERS.T_AUXILIAIRE=CORRESP.CR_LIBRETEXTE1 ' +
                              ' AND T1.T_AUXILIAIRE=CORRESP.CR_LIBRETEXTE2) ' +
                              ' OR (TIERS.T_AUXILIAIRE=CORRESP.CR_LIBRETEXTE2 ' +
                              ' AND T1.T_AUXILIAIRE=CORRESP.CR_LIBRETEXTE1) ) ';

    sztemp := sztemp + ')';
    SetControlText('XX_WHEREMP',sztemp);
    end;
{e fb 15/09/2006 FQ18581}
end;

function TOF_CPGENECOMPEN.ConstitueLots(Mode: TMode): Boolean;
  procedure AjouteChampSup(TOBL:  TOB);
  begin
    TOBL.PutValue('MARQUE','X') ;

    {Si compensation client, il faut placer les écritures des comptes clients en premier}
    if (smp = smpCompenCli) and ((TOBL.GetValue('T_NATUREAUXI')='CLI') or (TOBL.GetValue('T_NATUREAUXI')='AUD')) then
      TOBL.PutValue('TRINATURE','1')
    else if (smp = smpCompenFou) and ((TOBL.GetValue('T_NATUREAUXI')='FOU') or (TOBL.GetValue('T_NATUREAUXI')='AUC')) then
      TOBL.PutValue('TRINATURE','1')
    else
      TOBL.PutValue('TRINATURE','2');

    if Mode = mdEdition then
      TOBL.PutValue('TOTALREGLER', 1);
  end;

  procedure PrechargeOrigines;
  Var QQ:    TQuery ;
      st:    String ;
      i:     Integer;
      TOBL:  TOB;
  begin
    TOBLot.ClearDetail ;
    {b FP 13/07/2006 FQ18589
    St:='SELECT ECRITURE.*,G_NATUREGENE,G_VENTILABLE,';
    for i:=1 to MaxAxe do
      St:=St+'G_VENTILABLE'+IntToStr(i)+',';
    St:=St+'T_NATUREAUXI,T_CORRESP1,T_CORRESP2'+
        ' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+
                      ' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE';}
    (*
    St:='SELECT CPSUIVICOMPEN.*,';
    for i:=1 to MaxAxe do
      St:=St+'G_VENTILABLE'+IntToStr(i)+',';
    St:=St+'G_VENTILABLE' +
           ' FROM CPSUIVICOMPEN LEFT JOIN GENERAUX ON E_GENERAL=GENERAUX.G_GENERAL';
    {e FP 13/07/2006}*)
    // CA - 05/09/2007 - On revient en arrière pour ne plus utiliser la vue
    St:='SELECT ECRITURE.*,G_NATUREGENE,G_VENTILABLE,';
    for i:=1 to MaxAxe do
      St:=St+'G_VENTILABLE'+IntToStr(i)+',';
    St:=St+'T_NATUREAUXI,T_CORRESP1,T_CORRESP2'+
        ' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+
                      ' LEFT JOIN TIERS ON (T_AUXILIAIRE=E_AUXILIAIRE OR (E_AUXILIAIRE=T_PAYEUR AND E_CONTREPARTIEAUX=T_AUXILIAIRE AND E_PIECETP<>""))';
    QQ:=OpenSQL(St+' '+RecupWhereCritere(Pages),True) ;
    TOBLot.LoadDetailDB('ECRITURE','','',QQ,False,True) ;
    Ferme(QQ) ;
    if TOBLot.Detail.Count>0 then
      begin
      TOBL := TOBLot.Detail[0];
      TOBL.AddChampSup('MARQUE',      True);
      TOBL.AddChampSup('TRINATURE',   True);
      if Mode = mdEdition then
        begin
        TOBL.AddChampSup('ERRCOMPEN', True);  {1 = La ligne ne doit pas être utilisée pour la compensation}
        TOBL.AddChampSup('SOLDEEUR',  True);  {Solde en euro}
        TOBL.AddChampSup('TAUX',      True);  {Taux devise de la compensation}
        TOBL.AddChampSup('DEVISE',    True);  {Devise de la compensation}
        TOBL.AddChampSup('SOLDEDEV',  True);  {Solde en devise}
        TOBL.AddChampSup('TOTALREGLER',True); {1 si le montant doit être cumulé dans Total à Règler
                                               0 si le montant doit être cumulé dans Total à Recevoir
                                              -1 si le montant ne doit pas être cumulé}
        end;
      end;

    if FListe.AllSelected then
      begin
      for i:=0 to TOBLot.Detail.Count-1 do
        AjouteChampSup(TOBLot.Detail[i]);
      end
    else
      begin
      for i:=0 to FListe.NbSelected-1 do
        begin
        FListe.GotoLeBookmark(i) ;
        {$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.row - 1) ;
        {$ENDIF}
        TOBL:=TOBLot.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                              [Q.FindField('E_JOURNAL').AsString,Q.FindField('E_EXERCICE').AsString,
                             Q.FindField('E_DATECOMPTABLE').AsDateTime,Q.FindField('E_NUMEROPIECE').AsInteger,
                           Q.FindField('E_NUMLIGNE').AsInteger,Q.FindField('E_NUMECHE').AsInteger],False) ;
        if TOBL<>Nil then
          AjouteChampSup(TOBL);
        end ;
        VireInutiles(TOBLot, swapSelect);
      end;
  end;

  function VerifLot: Boolean;
  var
    SAJAL: TSAJAL;
    i:     Integer;
  begin
    {Si le lot n'est pas en euro, le journal de compensation doit être multi devise}
    Result := True;
    SAJAL := TSAJAL.Create(TCompensation.GetJalCompensation, False);
    if not SAJAL.MULTIDEVISE then
      begin
      for i:=0 to TOBLot.Detail.Count-1 do
        if TOBLot.Detail[i].GetValue('E_DEVISE') <> V_PGI.DevisePivot then
          begin
          HShowMessage(HM[12],gszCaption,'');
          Result := False;
          break;
          end;
      end;

    FreeAndNil(SAJAL);
  end;

var i                     : Integer ;
    TOBL                  : TOB ;
    LesChamps, NomChamp   : String ;
    LesVals               : String ;
begin
  LesCritsRupt.Clear;
  PrechargeOrigines;

  if Mode = mdGeneration then
    AjouteLesAnas( TOBLot ) ;

  Result := VerifLot;

  if Result then
    begin
    for i:=0 to TOBLot.Detail.Count-1 do
      BEGIN
      TOBL      := TOBLot.Detail[i] ;
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
end;

procedure TOF_CPGENECOMPEN.ConstitueDest(ValsRupt: String; TOBLot, TOBDest: TOB; DeviseDest: RDEVISE);

  procedure PassageDevise(TOBDest : TOB);
  Var TOBL:       TOB ;
      i:          integer ;
      MntDstDev:  Double;
      MntDstEur:  Double;
  begin
    for i:=0 to TOBDest.Detail.Count-1 do
      begin
      TOBL := TOBDest.Detail[i] ;

      MntDstDev := CalculSoldeARegler(TOBL, DeviseDest);
      if DeviseDest.Code = V_PGI.DevisePivot then
        MntDstEur := MntDstDev
      else
        MntDstEur := DeviseToEuro(MntDstDev, DeviseDest.Taux, DeviseDest.Quotite);

      if MntDstDev > 0 then
        begin
        TOBL.PutValue('E_DEBIT',     Abs(MntDstEur));
        TOBL.PutValue('E_DEBITDEV',  Abs(MntDstDev));
        end
      else begin
        TOBL.PutValue('E_CREDITDEV', Abs(MntDstDev));
        TOBL.PutValue('E_CREDIT',    Abs(MntDstEur));
        end;
      TOBL.PutValue('E_TAUXDEV',       DeviseDest.Taux);
      TOBL.PutValue('E_COTATION',      DeviseDest.Cotation);
      TOBL.PutValue('E_DEVISE',        DeviseDest.Code);
      TOBL.PutValue('E_COUVERTURE',    0);
      TOBL.PutValue('E_COUVERTUREDEV', 0);
      end;
  end;

  procedure ModifieEcr(TOBL: TOB; TOBRef: TOB);
  var
    i:      Integer;
    TOBAxe: TOB;
  begin
    for i:=1 to MaxAxe do
      ParmsMP.Ventils[i] := TOBRef.GetValue('G_VENTILABLE' + IntToStr(i)) = 'X';

    ParmsMP.CptG    := TOBRef.GetValue('E_GENERAL');
    ParmsMP.GenVent := TOBRef.GetValue('G_VENTILABLE')='X';
    TripoteDetail(TOBL, ParmsMP);

    {Retire l'analytique des axes non suivi par le compte général}
    for i:=0 to TOBL.Detail.Count-1 do
      begin
      TOBAxe := TOBL.Detail[i] ;
      if not ParmsMP.Ventils[i+1] then
        TOBAxe.Detail.Clear;
      end ;
  end;

  {b FP 01/06/2006 FQ18193}
  procedure ModifieContrepassation(TOBN: TOB; Indice: Integer; EchEnc: array of double);
  var
    Auxi:     String;                  {FP 30/05/2006 FQ18195}
    i:        Integer;                 {FP 14/06/2006 FQ18351}
  begin
    TOBN.AddChampSup('INDICE',False) ;
    TOBN.PutValue('INDICE',Indice) ;

    ModifieEcr(TOBN, TOBOrig.Detail[0]);

    {b FP 30/05/2006 FQ18195}
    Auxi := TOBOrig.Detail[0].GetValue('E_AUXILIAIRE');
    TOBN.PutValue('E_AUXILIAIRE', Auxi);
    TOBN.PutValue('E_RIB', GetRIBPrincipal(Auxi));
    {e FP 30/05/2006 FQ18195}

    {b FP 14/06/2006 FQ18351}
    // FQ 21139 SBO : mise en place en négatif de la base ht de la tva pour déduction des autres bases HT 
    for i:=0 to 3 do
      TOBN.PutValue('E_ECHEENC'+IntToStr(i+1), EchEnc[i]*-1);
    TOBN.PutValue('E_ECHEDEBIT', EchEnc[4]*-1 );  // FQ 21139 SBO : prise en charge de E_ECHEDEBIT
    {e FP 14/06/2006 FQ18351}
  end;
  {e FP 01/06/2006 FQ18193}

  procedure GenererContrepassation(TOBR: TOB);
  var
    TOBN:     TOB;
    i:        Integer;
    {b FP 01/06/2006 FQ18193}
    TOBL:     TOB;
    XD, XC:   Double;
    DP, CP:   Double;
    DD, CD:   Double;
    {e FP 01/06/2006 FQ18193}
    {b FP 14/06/2006 FQ18351}
    j:        Integer;
    EchEnc:   array[0..4] of double; // FQ 21139 SBO : prise en charge de E_ECHEDEBIT
    {e FP 14/06/2006 FQ18351}
  begin
    {b FP 01/06/2006 FQ18193}
    DP := 0; CP := 0;
    DD := 0; CD := 0;
    Fillchar(EchEnc, sizeof(EchEnc), #0); {FP 14/06/2006 FQ18351}

    if not ParmsMP.CumulEcrGene then
      begin
      for i := 0 to TOBDest.Detail.Count-1 do
        begin
        TOBL := TOBDest.Detail[i];
        TOBN := TOB.Create( 'ECRITURE', TOBR, -1 );
        TOBN.Dupliquer(TOBL,True,True,False);
        {b FP 14/06/2006 FQ18351}
        for j:=0 to 3 do
          EchEnc[j] := TOBL.GetValue('E_ECHEENC'+IntToStr(j+1));
        EchEnc[4] := TOBL.GetValue('E_ECHEDEBIT'); // FQ 21139 SBO : prise en charge de E_ECHEDEBIT
        {e FP 14/06/2006 FQ18351}
        ModifieEcr(TOBL, TOBL);
        ModifieContrepassation(TOBN, i+1, EchEnc);  {FP 14/06/2006 FQ18351}
        end;
      end
    else
      begin
      for i := 0 to TOBDest.Detail.Count-1 do
        begin
        TOBL := TOBDest.Detail[i];

        XD := GetLeMontant(TOBL.GetValue('E_DEBIT'), TOBL.GetValue('E_COUVERTURE'), TOBL.GetValue('E_SAISIMP'), tsmPivot);
        XC := GetLeMontant(TOBL.GetValue('E_CREDIT'), TOBL.GetValue('E_COUVERTURE'), TOBL.GetValue('E_SAISIMP'), tsmPivot);
        DP := Arrondi(DP+XD, V_PGI.OkDecV);
        CP := Arrondi(CP+XC, V_PGI.OkDecV);

        XD := GetLeMontant(TOBL.GetValue('E_DEBITDEV'), TOBL.GetValue('E_COUVERTUREDEV'), TOBL.GetValue('E_SAISIMP'), tsmDevise);
        XC := GetLeMontant(TOBL.GetValue('E_CREDITDEV'), TOBL.GetValue('E_COUVERTUREDEV'), TOBL.GetValue('E_SAISIMP'), tsmDevise);
        DD := Arrondi(DD+XD, ParmsMP.DEV.Decimale);
        CD := Arrondi(CD+XC, ParmsMP.DEV.Decimale);

        {b FP 14/06/2006 FQ18351}
        for j:=0 to 3 do
          EchEnc[j] := Arrondi(EchEnc[j]+TOBL.GetValue('E_ECHEENC'+IntToStr(j+1)), ParmsMP.DEV.Decimale);
        EchEnc[4] := Arrondi(EchEnc[4]+TOBL.GetValue('E_ECHEDEBIT'), ParmsMP.DEV.Decimale);// FQ 21139 SBO : prise en charge de E_ECHEDEBIT
        {e FP 14/06/2006 FQ18351}
        end;
      for i := TOBDest.Detail.Count-1 downto 1 do {Garde la première ligne: échéance la plus ancienne}
        TOBDest.Detail[i].Free;
      TOBL := TOBDest.Detail[0];
      TOBN := TOB.Create( 'ECRITURE', TOBR, -1 );
      TOBN.Dupliquer(TOBL,True,True,False);

      ModifieEcr(TOBL, TOBL);
      ModifieContrepassation(TOBN, 1, EchEnc);    {FP 14/06/2006 FQ18351}

      if ((DP<>0) and (CP<>0)) then
        begin
        if Abs(DP)>Abs(CP) then
          begin
          DP := DP-CP; CP:=0;
          DD := DD-CD; CD:=0;
          end
        else
          begin
          CP := CP-DP; DP:=0;
          CD := CD-DD; DD:=0;
          end;
        end;
      TOBN.PutValue('E_DEBIT',     DP);
      TOBN.PutValue('E_CREDIT',    CP);
      TOBN.PutValue('E_DEBITDEV',  DD);
      TOBN.PutValue('E_CREDITDEV', CD);
      {Ne pas Inverser le sens}
      TOBL.PutValue('E_DEBIT',     DP);
      TOBL.PutValue('E_CREDIT',    CP);
      TOBL.PutValue('E_DEBITDEV',  DD);
      TOBL.PutValue('E_CREDITDEV', CD);
      end;
    ParmsMP.CptG    := '';
    ParmsMP.GenVent := False;
    {e FP 01/06/2006 FQ18193}
  end;
Var
  StSort:             String;
  TOBR, TOBL:         TOB;
  TOBAxe, TOBA:       TOB;
  Indice:             Integer;
  ia, i:              Integer;
begin
  // Duplication
  TOBDest.ClearDetail ;

  ParmsMP.ChampsRupt := 'TRINATURE;E_ETABLISSEMENT';
  ValsRupt := '2;'+ValsRupt;
  DupliquerTOBsmp('', '', ValsRupt, TOBOrig, TOBDest, ParmsMP) ;
  StSort:='E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE';
  TOBDest.Detail.Sort(StSort) ;

  if TOBDest.Detail.Count > 0 then       {FP 25/04/2006 FQ17948}
    begin
    PassageDevise(TOBDest);
    ParmsMP.DEV := DeviseDest;

    {Génère la contrepassation Client ou Fournisseur}
    TOBR := TOB.Create('',Nil,-1) ;
    GenererContrepassation(TOBR);
    FinirDestination(TOBDest, TOBR, ParmsMP);
    FreeAndNil(TOBR) ;

    for Indice:=0 to TOBDest.Detail.Count-1 do
      begin
      TOBL := TOBDest.Detail[Indice];
      TOBL.PutValue('E_NUMLIGNE',        Indice+1);
      TOBL.PutValue('E_TYPEMVT',         '');
      TOBL.PutValue('E_CONTREPARTIEGEN', '');
      TOBL.PutValue('E_CONTREPARTIEAUX', '');
      TOBL.PutValue('E_LIBELLE',         FEcrLibelle);    {FP 25/04/2006 FQ17982}
      AffecteG(TOBL, ParmsMP);
      for ia:=0 to TOBL.Detail.Count-1 do
        begin
        TOBAxe := TOBL.Detail[ia] ;
        for i:=0 to TOBAxe.Detail.Count-1 do
          begin
          TOBA:=TOBAxe.Detail[i] ;
          TOBA.PutValue('Y_CONTREPARTIEGEN', '');
          TOBA.PutValue('Y_CONTREPARTIEAUX', '');
          end;
        end;
      end;
    end;
end;

function TOF_CPGENECOMPEN.ConstitueOrigine(ValsRupt: String; Mode: TMode; RuptEtab: TStrings; var DeviseDest: RDEVISE): Boolean;

  function VerifTOBOrig(StSort: String): Boolean;
  var
    i:                 Integer;
    TOBL:              TOB;
    SoldeCompensation: Double;
    InverseCompte:     Boolean;
  begin
    InverseCompte := False;

    {Vérifie la présence d'un compte client et d'un compte fournisseur}
    Result := False;
    for i:=0 to TOBOrig.Detail.Count-1 do
      begin
      TOBL := TOBOrig.Detail[i];
      if TOBL.GetValue('E_AUXILIAIRE') <> TOBOrig.Detail[0].GetValue('E_AUXILIAIRE') then
        begin
        Result := True;
        break;
        end;
      end;
    if (not Result) and (Mode = mdGeneration) then
      begin
      if CaseNata(TOBOrig.Detail[0].GetValue('T_NATUREAUXI')) = 1 {client} then
        PGIError('Il faut sélectionner au moins une facture fournisseur pour le compte '+
          TOBOrig.Detail[0].GetValue('E_AUXILIAIRE'))
      else
        PGIError('Il faut sélectionner au moins une facture client pour le compte '+
          TOBOrig.Detail[0].GetValue('E_AUXILIAIRE'));
      end;

    {Calcule le solde à règler}
    if Result then
      begin
      InitRDevise(TOBOrig.Detail[0], ParmsMP.DateC, DeviseDest);
      SoldeCompensation := 0;
      for i:=0 to TOBOrig.Detail.Count-1 do
        begin
        TOBL := TOBOrig.Detail[i];
        SoldeCompensation := SoldeCompensation+CalculSoldeARegler(TOBL, DeviseDest);
        end;
      if (smp = smpCompenCli) and (SoldeCompensation <= 0) then
        Result := False
      else if (smp = smpCompenFou) and (SoldeCompensation >= 0) then
        Result := False;

      ParmsMP.smp := smp;
      if (not Result) and (not IsMultiTiers) and (Mode = mdGeneration) then
        begin
        if (smp = smpCompenCli) then
          PGIError('Le solde des factures sélectionnées pour le tiers '+TOBOrig.Detail[0].GetValue('E_AUXILIAIRE')+
            ' est créditeur, aucun traitement n''est effectué')
        else
          PGIError('Le solde des factures sélectionnées pour le tiers '+TOBOrig.Detail[0].GetValue('E_AUXILIAIRE')+
            ' est débiteur, aucun traitement n''est effectué');
        end
      else if (not Result) and IsMultiTiers then
        begin
        if (SoldeCompensation=0) and (Mode = mdGeneration) then
          begin
          PGIError('Les factures sélectionnées pour le tiers '+TOBOrig.Detail[0].GetValue('E_AUXILIAIRE')+
            ' se soldent, aucun traitement n''est effectué');
          end
        else if (SoldeCompensation<>0) then
          begin
          {Il faut inverser le type de compensation: passer de la compensation client vers fournisseur ou inversement}
          if smp = smpCompenCli then
            ParmsMP.smp := smpCompenFou
          else
            ParmsMP.smp := smpCompenCli;
          for i:=0 to TOBOrig.Detail.Count-1 do
            begin
            TOBL := TOBOrig.Detail[i];
            if TOBL.GetValue('TRINATURE') = 1 then
              TOBL.PutValue('TRINATURE', 2)
            else
              TOBL.PutValue('TRINATURE', 1);
            if Mode = mdEdition then
              TOBL.PutValue('TOTALREGLER', 0);
            end;
          TOBOrig.Detail.Sort(StSort) ;
          InverseCompte := True;
          Result        := True;
          end;
        end;
      end;

    {Vérifie que les écritures du compte à règler ont la même devise}
    if Result then
      begin
      if InverseCompte then {Comme il est possible d'inverser l'ordre client/fournisseur, il faut retrouver la devise}
        InitRDevise(TOBOrig.Detail[0], ParmsMP.DateC, DeviseDest);
      for i:=0 to TOBOrig.Detail.Count-1 do
        begin
        TOBL := TOBOrig.Detail[i];
        if (TOBL.GetValue('E_AUXILIAIRE') = TOBOrig.Detail[0].GetValue('E_AUXILIAIRE')) and
           (TOBL.GetValue('E_DEVISE') <> TOBOrig.Detail[0].GetValue('E_DEVISE')) then
          begin
          if Mode = mdGeneration then
            PGIError('Les écritures du compte '+TOBOrig.Detail[0].GetValue('E_AUXILIAIRE')+' n''ont pas la même devise');
          Result := False;
          break;
          end;
        end;
      end;
  end;
var
  StSort: String;
  i:      Integer;
  TOBL:   TOB;
begin
  // Duplication
  TOBOrig.ClearDetail ;
  RuptEtab.Clear;

  ParmsMP.ChampsRupt := TCompensation.GetChampPlan;
  DupliquerTOBsmp('', '', ValsRupt, TOBLot, TOBOrig, ParmsMP) ;
  StSort:='TRINATURE;E_AUXILIAIRE;E_GENERAL;E_DATEECHEANCE;E_NUMEROPIECE';
  TOBOrig.Detail.Sort(StSort) ;

  Result := (TOBOrig.Detail.Count > 0) and VerifTOBOrig(StSort);

  for i:=0 to TOBOrig.Detail.Count-1 do
    begin
    TOBL := TOBOrig.Detail[i];
    if (Mode = mdEdition) then
      begin
      if Result then
        TOBL.PutValue('ERRCOMPEN', 0)
      else
        TOBL.PutValue('ERRCOMPEN', 1);
      end;
    if RuptEtab.IndexOf(TOBL.GetValue('E_ETABLISSEMENT')) = -1 then
      RuptEtab.Add(TOBL.GetValue('E_ETABLISSEMENT'));
    end;
end;
                                       
function TOF_CPGENECOMPEN.LInsertDB(T: Tob): Boolean;
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

function TOF_CPGENECOMPEN.CalculSoldeARegler(TOBL: TOB; DeviseDest: RDevise): Double;
var
  XD, XC:     Double;           {FP 01/06/2006 FQ18193}
  MntSrcEur:  Double;
  MntSrcDev:  Double;
begin
  {b FP 01/06/2006 FQ18193}
  {Retire le montant déjà règlé}
  XD := GetLeMontant(TOBL.GetValue('E_DEBIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot);
  XC := GetLeMontant(TOBL.GetValue('E_CREDIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot);
  MntSrcEur := Arrondi(XD-XC, V_PGI.OkDecV);

  XD := GetLeMontant(TOBL.GetValue('E_DEBITDEV'), TOBL.GetValue('E_COUVERTUREDEV'), TOBL.GetValue('E_SAISIMP'), tsmDevise);
  XC := GetLeMontant(TOBL.GetValue('E_CREDITDEV'), TOBL.GetValue('E_COUVERTUREDEV'), TOBL.GetValue('E_SAISIMP'), tsmDevise);
  MntSrcDev := Arrondi(XD-XC, DeviseDest.Decimale);
  {e FP 01/06/2006 FQ18193}

  if (DeviseDest.Code = TOBL.GetValue('E_DEVISE')) then
    begin
    Result := MntSrcDev;
    end
  else if DeviseDest.Code = V_PGI.DevisePivot then
    begin
    Result := MntSrcEur;
    end
  else
    begin
    Result := EuroToDevise(MntSrcEur, DeviseDest.Taux, DeviseDest.Quotite, DeviseDest.Decimale);
    end;
end;

function TOF_CPGENECOMPEN.IsMultiTiers: Boolean;
begin
  if SelectOnLot then
    Result := GetCheckBoxState('FECRDEBITCREDIT1') = cbChecked
  else
    Result := GetCheckBoxState('FECRDEBITCREDIT') = cbChecked;
end;

function TOF_CPGENECOMPEN.LanceSaisie(Etablissement: String): Boolean;
var
  MM:             RMVT;
  TOBN, TOBLet:   TOB;
  i:              Integer;
  GenereAuto:     Boolean;    {FP 25/04/2006 FQ17966}
begin
  MM     := TOBToIdent(TOBDest.Detail[0],False) ;
  CompleteMM(ParmsMP, tslAucun, False, False, '', MM);

  {Retire les écritures qui ne doivent pas être lettrées}
  TOBLet := TOB.Create('',Nil,-1) ;
  for i:=0 to TOBOrig.Detail.Count-1 do
    begin
    if (TOBOrig.Detail[i].GetValue('E_AUXILIAIRE') <> TOBOrig.Detail[0].GetValue('E_AUXILIAIRE')) and
       (TOBOrig.Detail[i].GetValue('E_ETABLISSEMENT') = Etablissement) then
      begin
      TOBN := TOB.Create('ECRITURE',TOBLet,-1) ;
      TOBN.Dupliquer(TOBOrig.Detail[i],True,True,False) ;
      end;
    end;
  GenereAuto := not ParmsMP.VisuEcrGene;         {FP 25/04/2006 FQ17966}
  Result := LanceSaisieMP(MM, TOBLet, nil, GenereAuto, ParmsMP.LettrAuto, nil, False, ParmsMP.ModeleBordereau);    {FP 25/04/2006 FQ17966}
  FreeAndNil(TOBLet);
end;

procedure TOF_CPGENECOMPEN.LanceEdition(TOBEdt: TOB);
var
  CritEdition:    String;
  TOBL:           TOB;
  i:              Integer;
  DeviseDest:     RDevise;
  RupturePlan:    String;
  XD, XC:         Double;        {FP 01/06/2006 FQ18193}
  Solde:          Double;
begin
  {Recopie les valeurs de l'onglet 'Sélection du lots' dans l'onglet 'Standard'}
  if SelectOnLot then
    begin
    SetControlText('CATEGORIE', GetControlText('CATEGORIE1'));
    SetControlChecked('FECRDEBITCREDIT', TCheckBox(GetControl('FECRDEBITCREDIT1')).Checked);
    end;
  SetControlText('NATUREPIECE', GetControlText('_E_NATUREPIECE'));
  TFmul(Ecran).CritModified := False;

  RupturePlan   := '';
  for i:=0 to TOBEdt.Detail.Count-1 do
    begin
    TOBL := TOBEdt.Detail[i];

    if (RupturePlan='') or (RupturePlan<>TOBL.GetValue(TCompensation.GetChampPlan)) then
      begin
      InitRDevise(TOBL, ParmsMP.DateC, DeviseDest);
      RupturePlan   := TOBL.GetValue(TCompensation.GetChampPlan);
      end;

    {b FP 01/06/2006 FQ18193}
    XD   := GetLeMontant(TOBL.GetValue('E_DEBIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot);
    XC   := GetLeMontant(TOBL.GetValue('E_CREDIT'),TOBL.GetValue('E_COUVERTURE'),TOBL.GetValue('E_SAISIMP'),tsmPivot);
    Solde := Arrondi(XD-XC, V_PGI.OkDecV);
    {e FP 01/06/2006 FQ18193}
    if smp = smpCompenFou then
      Solde := -1*Solde;
    TOBL.PutValue('SOLDEEUR', Solde);
    if TOBL.GetValue('ERRCOMPEN') = 0 then      {Pas d'erreur}
      begin
      if TOBL.GetValue('E_DEVISE') <> DeviseDest.Code then
        TOBL.PutValue('TAUX',   DeviseDest.Taux)
      else
        TOBL.PutValue('TAUX',   TOBL.GetValue('E_COTATION'));
      TOBL.PutValue('DEVISE',   DeviseDest.Code);
      Solde := CalculSoldeARegler(TOBL, DeviseDest);
      if smp = smpCompenFou then
        Solde := -1*Solde;
      TOBL.PutValue('SOLDEDEV', Solde);
      end
    else                                        {Compensation impossible}
      begin
      TOBL.PutValue('TAUX',     0);
      TOBL.PutValue('DEVISE',   '');
      TOBL.PutValue('SOLDEDEV', 0);
      end;
    {Le générateur d'état utilise uniquement le champ T_CORRESP1}
    TOBL.PutValue('T_CORRESP1', TOBL.GetValue(TCompensation.GetChampPlan));
    end;

  CritEdition := ComposantToCritere(Pages);
  if CritEdition <> '' then
    begin
    CritEdition := Copy(CritEdition, 2, Length(CritEdition));
    CritEdition := CritEdition+'`';
    end;
  CritEdition := CritEdition+'NUMPLAN='+IntToStr(TCompensation.GetNumPlanCorrespondance)+
                 '`TITREETAT=';
  if smp = smpCompenFou then
    CritEdition := CritEdition+HDiv[3]
  else
    CritEdition := CritEdition+HDiv[4];
  CritEdition := CritEdition+'`ECRDEBITCREDIT=';
  if smp = smpCompenFou then
    CritEdition := CritEdition+HDiv[5]
  else
    CritEdition := CritEdition+HDiv[6];
  CritEdition := CritEdition+'`TOTALGENERAL1=';
  if smp = smpCompenFou then
    CritEdition := CritEdition+HDiv[7]
  else
    CritEdition := CritEdition+HDiv[8];
  CritEdition := CritEdition+'`TOTALGENERAL2=';
  if smp = smpCompenFou then
    CritEdition := CritEdition+HDiv[8]
  else
    CritEdition := CritEdition+HDiv[7];
  CritEdition := CritEdition+'`DATEGENERATION='+DateToStr(ParmsMP.DateC);

  LanceEtatTob('E', 'CMP', 'CMP', TOBEdt, True, False, False, nil, '', Ecran.Caption,
    False, 0, CritEdition);
end;

function TOF_CPGENECOMPEN.VerifSelectionGrille: Boolean;
begin
  Result := True;
  if TFmul(Ecran).CritModified then
    begin
    PgiInfo(TraduireMemoire('Vous n''avez pas appliqué les critères.'));  {FP 17/07/2006}
    Result := False;
    exit;
    end;

  // Le résultat de la recherche est-il vide ?
{$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then
{$ELSE}
  if (Q.Eof and Q.Bof) then
{$ENDIF}
    begin
    PGIInfo(TraduireMemoire('Le résultat de la recherche est vide. Veuillez relancer une recherche.'),gszCaption) ;   {FP 17/07/2006}
    Result := False;
    exit ;
    end ;

  // Au moins 2 lignes de sélectionnées en mode normale, (pas de test en mode inverse pour l'instant)
  if ( FListe.NbSelected < 2 ) and ( (not FListe.AllSelected) and (not swapSelect) ) then
    begin
    PGIInfo(TraduireMemoire('Aucune ligne à traiter. Vous devez sélectionner au moins 2 écritures.'),gszCaption) ; {FP 17/07/2006}
    Result := False;
    exit ;
    end ;

  // Message avertissement pour sélection inversée
  if swapSelect then
    if PGIAsk('Vous êtes en mode "sélection inversée", Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes then
      begin
      Result := False;
      exit;
      end;
end;

procedure TOF_CPGENECOMPEN.BEditionClick(Sender: TObject);
var
  i, k:              Integer;           
  DeviseDest:        RDEVISE;
  TOBEdt, TOBL:      TOB;
  RuptEtab:          TStringList;
begin
  if not VerifSelectionGrille then
    Exit;

  RuptEtab := TStringList.Create;
  RuptEtab.Sorted     := True ;                
  RuptEtab.Duplicates := dupIgnore ;


  ParmsMP.ChampsRupt := TCompensation.GetChampPlan;

  if Not ConstitueLots(mdEdition) then Exit ;
  if TOBLot.Detail.Count<=0 then Exit ;
  if not InitParmsMP then Exit;

  TOBEdt := TOB.Create('',Nil,-1) ;
  try                                              
    for i:=0 to LesCritsRupt.Count-1 do
      begin
      ConstitueOrigine(LesCritsRupt[i], mdEdition, RuptEtab, DeviseDest);
      for k:=0 to TOBOrig.Detail.Count-1 do
        begin
        TOBL := TOB.Create('', TOBEdt, -1);
        TOBL.Dupliquer(TOBOrig.Detail[k], True, True, False);
        end;
      end;
    LanceEdition(TOBEdt);
  finally
    FreeAndNil(TOBEdt);
    end;
end;

procedure TOF_CPGENECOMPEN.InitRDevise(TOBL: TOB; DateCompta: TDateTime; var DeviseDest: RDEVISE);
var
  Lot:      String;
  DateTaux: TDateTime;
begin
  Lot := Trim(TOBL.GetValue('E_NOMLOT'));
  DateTaux := DateCompta;

  DeviseDest.Code     := TOBL.GetValue('E_DEVISE');
  DeviseDest.DateTaux := DateTaux;
  GetInfosDevise(DeviseDest);
  DeviseDest.Taux     := GetTaux(DeviseDest.Code, DeviseDest.DateTaux, Date);
end;

procedure TOF_CPGENECOMPEN.RAZNomLot(Etablissement: String);
var
  i:     Integer;
  TOBL:  TOB;
  SQL:   String;
begin
  BEGINTRANS;
  try
    for i:=0 to TOBOrig.Detail.Count-1 do
      begin
      TOBL := TOBOrig.Detail[i];
      if TOBL.GetValue('E_ETABLISSEMENT') = Etablissement then
        begin
        SQL:= 'UPDATE ECRITURE SET E_NOMLOT=""' +
              ' WHERE E_GENERAL="'+TOBL.GetValue('E_GENERAL')+'"'+
                 ' AND E_AUXILIAIRE="'+TOBL.GetValue('E_AUXILIAIRE')+'"'+
                 ' AND E_EXERCICE="'+TOBL.GetValue('E_EXERCICE')+'"'+
                 ' AND E_JOURNAL="'+TOBL.GetValue('E_JOURNAL')+'"'+
                 ' AND E_DATECOMPTABLE="'+USDATETIME(TOBL.GetValue('E_DATECOMPTABLE'))+'"'+
                 ' AND E_NUMEROPIECE='+IntToStr(TOBL.GetValue('E_NUMEROPIECE'))+
                 ' AND E_NUMLIGNE='+IntToStr(TOBL.GetValue('E_NUMLIGNE'))+
                 ' AND E_NUMECHE='+IntToStr(TOBL.GetValue('E_NUMECHE')) ;
        ExecuteSQL(SQL);
        end;
      end;
    COMMITTRANS;
  except
    ROLLBACK;
    raise;
    end;
end;


function TOF_CPGENECOMPEN.InitParmsMP: Boolean;
  function GetNatureJal(Jal: String): String;
  var
    Q: TQuery;
  begin
    Result := '';
    Q := OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+Jal+'" ', TRUE);
    if not Q.Eof then
      Result := Q.FindField('J_NATUREJAL').AsString;
    Ferme(Q);
  end;
begin             {InitParmsMP}
  ParmsMP.JalG        := TCompensation.GetJalCompensation;
  ParmsMP.DateC       := V_PGI.DateEntree;
  ParmsMP.NatJalG     := GetNatureJal(ParmsMP.JalG);
  ParmsMP.ExportCFONB := False;
  {FP 25/04/2006 FQ17982: Utilisation de la zone ParmsMP.NumEncaDeca pour stocker le libellé}
  ParmsMP.NumEncaDeca := '';//HDiv[2]; //fb 22/09/2006 FQ18850
  ParmsMP.ForceEche   := False;
  ParmsMP.LettrAuto   := True;
  ParmsMP.GenColl     := True;
  ParmsMP.MPG         := '';      {Force le mode de paiement}
  ParmsMP.Cat         := CATEGORIE.Value;
  ParmsMP.VisuEcrGene := True;    {FP 25/04/2006 FQ17966}
  {ParmsMP.DEV     est modifié dans ConstitueDest}
  {ParmsMP.smp     est modifié dans VerifTOBOrig}
  {ParmsMP.CptG    est modifié dans GenererContrepassation}
  {ParmsMP.GenVent est modifié dans GenererContrepassation}

  ParmsMP.DEV.Code := V_PGI.DevisePivot;  {Pour que ParamsMPSup ne gère pas la devise}
  Result := ParamsMPSup(ParmsMP);

  {b FP 25/04/2006 FQ17982: Utilisation de la zone ParmsMP.NumEncaDeca pour stocker le libellé}
  FEcrLibelle         := Trim(ParmsMP.NumEncaDeca);
  ParmsMP.NumEncaDeca := '';
  {e FP 25/04/2006}
end;

procedure TOF_CPGENECOMPEN.BChangeTauxClick(Sender: TObject);
var TOBOrig : Tob ;
    TobL : Tob;
    i     : Integer ;
    lQEcr : TQuery ;
ReqEcriture : String;
begin
  {b FP 17/07/2006: Provient de UTOFCPMULMVT}

  if TFmul(Ecran).CritModified then
    begin
    PgiInfo(TraduireMemoire('Vous n''avez pas appliqué les critères.'));
    exit;
    end;

 if FListe.nbSelected > 0 then
 begin
     TOBOrig := TOB.Create( 'TOB_ORIGINE', nil, -1 ) ;

     For i:=0 to FListe.NbSelected-1 do
     begin
        FListe.GotoLeBookmark(i) ;

        ReqEcriture := 'SELECT ECRITURE.* FROM ECRITURE '
                 + 'WHERE E_JOURNAL="' + GetField('E_JOURNAL')  + '" AND '
                 + 'E_EXERCICE="' + GetField('E_EXERCICE') + '" AND '
                 + 'E_DATECOMPTABLE="' + UsDateTime(GetField('E_DATECOMPTABLE')) + '" AND '
                 + 'E_NUMEROPIECE='     + IntToStr(GetField('E_NUMEROPIECE')) + ' AND '
                 + 'E_NUMLIGNE=' + IntToStr(GetField('E_NUMLIGNE')) + ' AND '
                 + 'E_NUMECHE=' + IntToStr(GetField('E_NUMECHE')) ;

        TOBL := TOB.Create('ECRITURE', TOBOrig, -1 ) ;
        lQEcr := OpenSQL( ReqEcriture, True ) ;
        TOBL.SelectDB( '', lQEcr ) ;
        Ferme( lQEcr ) ;
     end ;

     ChangeLeTauxDevise(TOBOrig, TaModif) ;
     FreeAndNil(TOBOrig) ;
 end
 else
     PGIInfo(TraduireMemoire('Veuillez sélectionner une pièce'),'Gestion du taux');

  {e FP 17/07/2006}
end;

procedure TOF_CPGENECOMPEN.CBDeviseChange(Sender: TObject);
begin
  {b FP 17/07/2006: Provient de UTOFCPMULMVT}
  If (GetControlText('E_DEVISE')<>'')
      and (GetControlText('E_Devise')<>V_PGI.DevisePivot)
      and (GetControlText('E_Devise')<>V_PGI.DeviseFongible)
      {and (Not EstMonnaieIn(GetControlText('E_Devise')))} Then
    begin
      SetControlVisible('BChangeTaux', TRUE );
    end
  Else
    begin
      SetControlVisible('BChangeTaux', FALSE );
    end;

  if Assigned(CBDeviseChangeParent) then
    CBDeviseChangeParent(Sender);
  {e FP 17/07/2006}
end;
Initialization
  registerclasses ( [ TOF_CPGENECOMPEN ] ) ;
end.

