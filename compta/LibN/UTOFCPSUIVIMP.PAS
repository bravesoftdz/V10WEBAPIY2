{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/03/2003
Modifié le ... : 23/04/2003
Description .. : Source TOF de la FICHE : CPSUIVIMP ()
Mots clefs ... : TOF;CPSUIVIMP
*****************************************************************}
Unit UTOFCPSUIVIMP ;

Interface

Uses Windows,
     {$IFDEF VER150} Variants, {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     ExtCtrls,  // pout le TPanel
     Classes,
     StdCtrls,
     Controls,
     HTB97,     // pour TToolBarButton97
     Grids,     // Pour le TGridDrawState
     Graphics,  // pour clRed
{$IFDEF EAGLCLIENT}
    eMul,
    maineagl,
{$ELSE}
    FE_MAIN,
    Mul,
    db,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    SaisBor,       // Pour lanceSaisieFolio
    HDB,
    dbGrids,
{$ENDIF}
     Saisie,        // pour TrouveEtLanceSaisie
     HQry,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     HStatus,       // pour InitMove et MoveCur et FiniMove
     UTOF,
     UTOB,
     uLibWindows,   // pour CDessineTriangle, CVireLigne
     MulSMPUtil,    // pour
     EncUtil,       // pour CatToMP
     FichComm,      // pour
     SaisLot,       // pour
     DetruitLot,    // Pour DetruitLeLot
     Filtre,        // pour VideFiltre
     TofVerifRib,   // pour CPLanceFiche_VerifRib
     SaisUtil,      // Pour TSuiviMP
     LettUtil,     // Pour RempliSelectEuro
     ParamSoc,		// GetParamSocSecur YMO
     Lookup;  //fb 19/06/2006 FQ12487

Procedure SelectSuiviMP ( smp : TSuiviMP ) ;
Procedure CPLancefiche_SuiviMP ( vStRange, vStLequel, vStArgs : String ) ;

Type
  TOF_CPSUIVIMP = Class (TOF)

    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnDisplay                ; override ;

    // EVT FILTRES
    {JP 13/09/05 : FQ 16395 : Je remplace E_EXERCICEExit par E_EXERCICEChange pour que l'évènement
                   ne soit pas systématiquement exécuté lors de la recherche
    procedure E_EXERCICEExit(Sender: TObject);}
    procedure E_EXERCICEChange(Sender: TObject);
    {JP 19/08/04 : FQ 14212 : L'avantage du OnChange sur le OnExit, c'est qu'il est exécuté
                              sur les nouvelles recherches
    procedure CATEGORIEExit(Sender: TObject);}
    procedure CATEGORIEChange(Sender: TObject);

    procedure FTICExit(Sender: TObject);
    procedure XNOMLOT1Exit(Sender: TObject);
    procedure TraiteEditeExit(Sender: TObject);

    // EVT BOUTONS
    procedure BRAZLotClick(Sender: TObject);
    procedure BCtrlRibClick(Sender: TObject);
    procedure BKillLotClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure BSwapSelectClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);

    // EVT Grille
    procedure FListeDblClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
    procedure FListeDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FListeSortColumn(Sender : TObject);
{$ENDIF}
    procedure FListeRowEnter(Sender: TObject);
    procedure FListeFlipSelection(Sender: TObject);

    // EVT Form
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FetchLesTous;
{b fb 19/06/2006 FQ12487}
    procedure SelectionElipsisClick(Sender : TObject);
    procedure AuxiliaireElipsisClick(Sender : TObject);
{e fb 19/06/2006 FQ12487}

    procedure AuxiElipsisClick(Sender : TObject);
  private
  	Q               : THQuery ;
{$IFDEF EAGLCLIENT}
    FListe          : THGrid ;
{$ELSE}
    FListe          : THDBGrid ;
{$ENDIF}

    // Composants géés manuellement
    FTIC            : TCheckBox ;
    CATEGORIE       : THValComboBox ;
    cFactCredit     : TCheckBox ;
    TraiteEdite     : TCheckBox ;
    XNOMLOT1        : THEdit ;
    XNOMLOT2        : THEdit ;
    CRechVide       : TCheckBox ;

    // Boutons Toolbar
    BCtrlRib        : TToolbarButton97 ;
    BRAZLot         : TToolbarButton97 ;
    BKillLot        : TToolbarButton97 ;
    BSwapSelect     : TToolbarButton97 ;

    // Messages
      HDiv          : TStringList ;

    // Evt Grille
{$IFDEF EAGLCLIENT}
{$ELSE}
    FListeDrawColumnCellParent  : TDrawColumnCellEvent ;
{$ENDIF}
    bSelectAllClickParent     : TNotifyEvent ;
    FormKeyDownParent         : TKeyEvent ;

    // Variables locales
    TobMvt      : Tob ;
    CodeLot     : String ;
    SwapSelect  : Boolean ;						 // Gestion du mode selection inversé
    gszCaption  : String ;             // Pour les messages
    bgFetchAll     : Boolean;     // Pour le bouton Sélection inversion et Select All
    Selection, Auxiliaire : THEdit; //FP 13/07/2006 FQ18589 THCritMaskEdit;  {fb 19/06/2006 FQ12487}

    // init des messagebox et menuzoom
    procedure InitVariables ;
    procedure InitHelpContext ;
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitMsgBox;
    procedure InitCriteres ;
    procedure InitE_NUMTRACHQ(Invisible : Boolean) ;
    procedure InitContexte ;
    procedure InitCATEGORIE ;
    procedure InitConditions ;
    procedure ReinitWhereNatCpt(RAZ : Boolean) ;
    procedure UpdateConditions ;
    procedure PreciseMP ;
    Function  ConstitueTobMvt : Boolean ;
    procedure MarqueOrigine ;
    procedure PrechargeOrigines ;
    procedure ValideLot ;
    procedure ClickModifRib ;
    Function  GereXNomLot : Boolean ;
    Function  AvecXNomLot : Boolean ;
    Function  CalculSoldeSelection : Double ;
    Procedure AfficheSoldeSelection ;
    procedure ToutMarquer ;
    Function  YEnATrop : boolean ;

  public
    smp         : TSuiviMP ;
    SorteLettre : TSorteLettre ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPObjetGen,
  ULibExercice,
  {$ENDIF MODENT1}
  UTofMulParamGen, utilPGI; {26/04/07 YMO F5 sur Auxiliaire }

// =============================================================================
Procedure SelectSuiviMP ( smp : TSuiviMP ) ;
var lStArgs : String ;
begin
  {b FP 21/02/2006}
  if (smp in [smpCompenCli, smpCompenFou]) and (not TCompensation.IsCompensation) then
    begin
    PGIInfo('La gestion de la compensation n''est pas active');
    Exit;
    end;
  {e FP 21/02/2006}
  lStArgs := SmpToStr(smp) ;
  CPLanceFiche_SuiviMP ( '', '', lStArgs ) ;
end ;

Procedure CPLancefiche_SuiviMP ( vStRange, vStLequel, vStArgs : String ) ;
begin                       
  AGLLanceFiche ( 'CP' , 'CPSUIVIMP' , vStRange , vStLequel , vStArgs ) ;
end ;
// =============================================================================

procedure TOF_CPSUIVIMP.OnUpdate ;
begin

  inherited;

  AfficheSoldeSelection ;
  FListeRowEnter(nil); // 10930

end ;

procedure TOF_CPSUIVIMP.OnLoad ;
begin
  Inherited ;
  UpdateConditions ;
end ;

procedure TOF_CPSUIVIMP.OnArgument (S : String ) ;
var vStSmp  : String ;
begin

  Inherited ;

  // Recup smp + param fiche :
  vStSmp      := S ;
  smp         := StrToSmp(vStSmp) ;
  SorteLettre := AttribSL(smp) ;
  {JP 29/11/05 : FQ 17030 : Avec les derniers agl, il faut utiliser SetDBListe, le filtre,
                 ParamDBG et la liste doivent avoir le même nom
  TFMul(Ecran).FNomFiltre  := AttribFiltre(smp, 'SV') ;}

  bgFetchAll := True;

  // context d'aide
  InitHelpContext ;
  // Récup composants
  InitComposants ;
  // Afectation EVT Grille
  InitEvenements ;
  // init messages
  InitMsgBox ;
  // Init variables locales
  InitVariables ;
  // Init des critères du mul
  InitCriteres ;
  // Init zones XX_... et critères cachés
  InitConditions ;
  // ??
  InitContexte ;

  // Rend invisible l'onglet contenant les zones cachées
  TTabSheet(GetControl('INVISIBLE',True)).TabVisible := False ;

  {JP 29/11/05 : FQ 17030 : Avec les derniers agl, il faut utiliser SetDBListe, le filtre,
                 ParamDBG et la liste doivent avoir le même nom
                 Dans l'immédiat, je pare au plus pressé, mais il faudrait doubler toutes les
                 listes pour que CPSUIVIMP et CPGENEREMP aient des filtres différents
  TFMul(Ecran).Q.Manuel     := True ;
  TFMul(Ecran).Q.Liste      := AttribListe(smp) ;
  TFMul(Ecran).Q.Manuel     := False ;}
  TFMul(Ecran).SetDBListe(AttribListe(smp, 'SV'));

  TFMul(Ecran).bSelectAll.Visible := True; // 14698

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('_E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPSUIVIMP.OnClose ;
begin
  Inherited ;
  TobMvt.ClearDetail ;
  FreeAndNil(TobMvt) ;
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

procedure TOF_CPSUIVIMP.OnDisplay;
begin
  FListeRowEnter(nil);
end;

procedure TOF_CPSUIVIMP.InitMsgBox;
begin
  // Init HDIV
  HDiv := TStringList.Create ;
  HDiv.Add('Euro');
  HDiv.Add('Préparation des prélèvements');
  HDiv.Add('Préparation des lettres-chèque');
  HDiv.Add('Préparation des virements');
  HDiv.Add('Préparation des Lettres-BOR');
  HDiv.Add('Préparation des lettres-traite');
  HDiv.Add('Préparation des remises en banque de traites');
  HDiv.Add('Préparation des remises en banque de traites');
  HDiv.Add('Préparation des remises à l''encaissement des traites');
  HDiv.Add('Préparation des remise à l''escompte des traites');
  HDiv.Add('Préparation des remises en banque');
  HDiv.Add('Escompte des traites');
  HDiv.Add('Encaissements des lots');
  HDiv.Add('Décaissements des lots');
  HDiv.Add('Préparation des lettres-traite');
  HDiv.Add('Factures au débit');
  HDiv.Add('Préparation des mises en portefeuille des traites');
  HDiv.Add('Préparation des mises en portefeuille des traites');
  HDiv.Add('Préparation des lettres-virement');
  HDiv.Add('Préparation des lettres-prélèvement');
  HDiv.Add('Encaissements divers');
  HDiv.Add('Décaissements divers');
  HDiv.Add('Escompte à');
  HDiv.Add('TVA sur escompte');
  HDiv.Add('Numéro &BOR de');
  HDiv.Add('Numéro &Chèques de');
  HDiv.Add('Préparation des lettres-traite');
  HDiv.Add('Préparation des Lettres-BOR');
  HDiv.Add('Préparation des mises en portefeuille des chèques');
  HDiv.Add('Préparation des mises en portefeuille des cartes bleues');
  HDiv.Add('Préparation des remises en banque de chèques');
  HDiv.Add('Préparation des remise en banque de cartes bleues');
  HDiv.Add('ATTENTION : Certaines échéances en cours de traitement n''ont pas été mises à jour !');
  HDiv.Add('33;?Caption?;Les codes lots sur les écritures sélectionnées vont être remis à zéro. Confirmez-vous ?;E;YN;N;N;');
  HDiv.Add('Préparation des lettres-virement international');
  HDiv.Add('Préparation des virements internationaux');
  HDiv.Add('36');
  {b FP 21/02/2006}
  HDiv.Add('Préparation de la compensation');   {37 ème élément}
  HDiv.Add('Préparation de la compensation');   {38 ème élément}
  {e FP 21/02/2006}
end;

procedure TOF_CPSUIVIMP.AfficheSoldeSelection;
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

function TOF_CPSUIVIMP.CalculSoldeSelection: Double;
var
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
    END;

    if SwapSelect then  // en mode inversé, on retranche le cumul des lignes sélectionnée du solde total
    begin
      debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_' + ChpDeb)).text ) ;
      credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_' + ChpCre)).text ) ;
      Result := (debit - credit) - Result ;
    end ;
  end ;
end;

procedure TOF_CPSUIVIMP.ClickModifRib;
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
    TFMul(Ecran).BChercheClick(Nil) ;
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
end;

function TOF_CPSUIVIMP.ConstitueTobMvt: Boolean;
var i : integer ;
begin
  Result:=TRUE ;
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
      MarqueOrigine ;
    end ;
    VireInutiles(TobMvt, swapSelect);
  end;

end;

function TOF_CPSUIVIMP.GereXNomLot: Boolean;
var St,
    XAnd,
    St1,
    St2 : String ;
begin
  Result := False ;
  SetControlText('XX_WHERELOT', ' E_NOMLOT="" ') ;
  St1:='' ;
  St2:='' ;
  if XNOMLOT1.Text<>'' then
    St1:=' (E_NOMLOT>="'+XNOMLOT1.Text+'") ' ;
  if XNOMLOT2.Text<>'' then
    St2:=' (E_NOMLOT<="'+XNOMLOT2.Text+'") ' ;
  if (St1<>'') and (St2<>'')
    then XAnd:=' AND '
    else XAnd:='' ;
  if (St1<>'') or (St2<>'') then
    begin
    Result:=TRUE ;
    if not CRechVide.Checked
      then St:=' (E_NOMLOT<>"" AND ('+St1+XAnd+St2+')) '
      else St:=' (E_NOMLOT="" OR ('+St1+XAnd+St2+')) ' ;
    SetControlText('XX_WhereLot', St ) ;
    end ;
end;

procedure TOF_CPSUIVIMP.InitContexte;
begin
  InitE_NUMTRACHQ( smp in [smpEncPreEdt,smpEncTous,smpDecVirEdt,smpDecVirBqe,smpDecVirInEdt,smpDecVirInBqe,smpEncPreBqe,
                           smpCompenCli, smpCompenFou] ) ;  {FP 21/02/2006}
  XNomLot1.Plus := AttribPlus(smp) ;
  XNomLot2.Plus := AttribPlus(smp) ;
  Case smp of
   smpEncDiv    : Ecran.Caption:=HDiv[20] ; // Encaissement divers
   smpEncPreEdt : Ecran.Caption:=HDiv[19] ; // Prélèvements clients
   smpEncPreBqe : Ecran.Caption:=HDiv[1] ; // Prélèvements clients
   smpEncChqPor : Ecran.Caption:=HDiv[28] ; // Mise en portefeuille chèques clients
    smpEncCBPor : Ecran.Caption:=HDiv[16] ; // Mise en portefeuille cartes bleues clients
   smpEncTraPor : Ecran.Caption:=HDiv[16] ; // Mise en portefeuille traite clients
   smpEncTraEdt : Ecran.Caption:=HDiv[5] ; // Lettres-traite clients
 smpEncTraEdtNC : Ecran.Caption:=HDiv[26] ; // Lettres-traite clients
   smpEncTraEnc : Ecran.Caption:=HDiv[8] ; // Traite clients à l'encaissement
   smpEncTraEsc : Ecran.Caption:=HDiv[9] ; // Traite clients à l'escompte
   smpEncChqBqe : Ecran.Caption:=HDiv[31] ; // Chèque clients en banque
    smpEncCBBqe : Ecran.Caption:=HDiv[6] ; // CB clients en banque
   smpEncTraBqe : Ecran.Caption:=HDiv[6] ; // Traite clients en banque
     smpEncTous : Ecran.Caption:=HDiv[12] ; // Prélèvements clients
   smpDecChqEdt : BEGIN // Lettres-chèques fournisseurs
                  Ecran.Caption:=HDiv[2] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[25] ) ;
                  END ;
 smpDecChqEdtNC : BEGIN // Lettres-chèques fournisseurs
                  Ecran.Caption:=HDiv[2] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[25] ) ;
                  END ;
   smpDecVirEdt : Ecran.Caption:=HDiv[18] ; // Virements fournisseurs
 smpDecVirEdtNC : Ecran.Caption:=HDiv[18] ; // Virements fournisseurs
   smpDecVirBqe : Ecran.Caption:=HDiv[3] ; // Virements fournisseurs
 smpDecVirInEdt : Ecran.Caption:=HDiv[34] ; // Virements internationales fournisseurs
 smpDecVirInEdtNC : Ecran.Caption:=HDiv[34] ; // Virements internationales fournisseurs
 smpDecVirInBqe : Ecran.Caption:=HDiv[35] ; // Virements internationales fournisseurs
   smpDecBorEdt : BEGIN // Lettres-Bor fournisseurs
                  Ecran.Caption:=HDiv[27] ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  END ;
 smpDecBorEdtNC : BEGIN // Lettres-Bor fournisseurs
                  Ecran.Caption:=HDiv[4] ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  END ;
   smpDecborDec : BEGIN // Traites fournisseurs en encaissement
                  Ecran.Caption:=HDiv[10] ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  END ;
   smpDecBorEsc : BEGIN // Traites fournisseurs à l'escompte
                  Ecran.Caption:=HDiv[11] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  END ;
   smpDecTraBqe : BEGIN // Traites fournisseurs en banque
                  Ecran.Caption:=HDiv[7] ;
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  END ;
   smpDecTraPor : BEGIN // Mise en portefeuille traite clients
                  Ecran.Caption:=HDiv[17] ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  SetControlText('TE_NUMTRAITECHQ', HDiv[24] ) ;
                  END ;
      smpDecDiv : Ecran.Caption:=HDiv[21] ; // Décaissements divers
     smpDecTous : Ecran.Caption:=HDiv[13] ; //
   {b FP 21/02/2006}
   smpCompenCli : Ecran.Caption:=HDiv[37]; {Compensation client}
   smpCompenFou : Ecran.Caption:=HDiv[38]; {Compensation fournisseur}
   {e FP 21/02/2006}
   END ;

  {b FP 21/02/2006}
  if IsCompensation(smp) then
    begin
    SetControlVisible('TE_GENERAL',      False);
    SetControlVisible('E_GENERAL',       False);
    SetControlVisible('FTIC',            False);
    SetControlVisible('TE_NUMENCADECA',  False);
    SetControlVisible('E_NUMENCADECA',   False);
    SetControlVisible('TE_BANQUEPREVI',  False);
    SetControlVisible('E_BANQUEPREVI',   False);
    SetControlVisible('TE_BANQUEPREVI_', False);
    SetControlVisible('E_BANQUEPREVI_',  False);
    cFactCredit.Visible := False;
    TraiteEdite.Visible := False;
    BCtrlRib.Visible    := False;
    {b FP 13/07/2006 FQ18589}
    if smp = smpCompenCli then
      SetControlProperty('_E_AUXILIAIRE','DataType','TZTTOUTDEBIT')
    else
      SetControlProperty( '_E_AUXILIAIRE',  'DataType', 'TZTTOUTCREDIT' );
    {e FP 13/07/2006 FQ18589}
    SetControlProperty('TE_DEVISE',     'Top', GetControl('TE_AUXILIAIRE').Top);
    SetControlProperty('E_DEVISE',      'Top', GetControl('_E_AUXILIAIRE').Top);  {FP 15/04/2006 FQ17967}
    SetControlProperty('TE_AUXILIAIRE', 'Top', GetControl('TE_GENERAL').Top);
    SetControlProperty('_E_AUXILIAIRE',  'Top', GetControl('E_GENERAL').Top);     {FP 15/04/2006 FQ17967}
    THValComboBox(GetControl('E_DEVISE')).Vide :=True;
    THValComboBox(GetControl('E_DEVISE')).ReLoad;
    end
  else {e FP 21/02/2006}
  if isEnc(smp) then
    begin
    SetControlProperty('E_GENERAL','DataType','TZGENCAIS') ;
    SetControlProperty('_E_AUXILIAIRE','DataType','TZTTOUTDEBIT') ;        {FP 15/04/2006 FQ17967}
    end
  else
    begin
    SetControlProperty( 'E_GENERAL',     'DataType', 'TZGDECAIS'     ) ;
    SetControlProperty( '_E_AUXILIAIRE', 'DataType', 'TZTTOUTCREDIT' ) ;   {FP 15/04/2006 FQ17967}
    cFactCredit.Caption := HDiv[15] ;
    end ;

  Ecran.Caption:=traduireMemoire(Ecran.Caption) ;
  gszCaption  := Ecran.Caption;  // Pour les messages
  UpdateCaption(Ecran) ;

  InitCATEGORIE ;

end;

procedure TOF_CPSUIVIMP.InitCriteres;
var i       : integer ;
    Pages   : TPageControl ;
begin
  Pages := TPageControl(GetControl('Pages', True)) ;
  if ((smp<>smpEncTous) and (smp<>smpDecTous)) then
    begin
    LibellesTableLibre( TTabSheet(GetControl('PLibres', True)), 'TT_TABLE', 'T_TABLE', 'T') ;
    if VH^.CPExoRef.Code<>'' then
      begin
      SetControlText('E_EXERCICE', VH^.CPExoRef.Code ) ;
      SetControlText('E_DATECOMPTABLE', DateToStr(VH^.CPExoRef.Deb) ) ;
      SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin) ) ;
      end
    else
      begin
      SetControlText('E_EXERCICE', VH^.Entree.Code ) ;
      SetControlText('E_DATECOMPTABLE', DateToStr(VH^.Entree.Deb) ) ;
      SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.Entree.Fin) ) ;
      end ;
    SetControlText('E_DATEECHEANCE', StDate1900 ) ;
    SetControlText('E_DATEECHEANCE_', StDate2099 ) ;
    SetControlText('E_DEVISE' , V_PGI.DevisePivot ) ;
    Pages.ActivePage := TTabSheet(GetControl('PCritere')) ;
//    SetActiveTabSheet('PCritere') ;
    end
  else
    begin
    Pages.ActivePage := TTabSheet(GetControl('PBanqueLot')) ;
//    SetActiveTabSheet('PBanqueLot') ;
    for i:=0 to Pages.PageCount - 1 do
      if Pages.Pages[i].Name<>'PBanqueLot' then
        Pages.Pages[i].TabVisible := False ;
    end ;

  // ???
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
  TraiteEditeExit(Nil) ;

end;

procedure TOF_CPSUIVIMP.InitE_NUMTRACHQ( Invisible : Boolean );
begin
  SetControlVisible('E_NUMTRAITECHQ_',  Not Invisible ) ;
  SetControlVisible('TE_NUMTRAITECHQ_', Not Invisible ) ;
  SetControlVisible('E_NUMTRAITECHQ',   Not Invisible ) ;
  SetControlVisible('TE_NUMTRAITECHQ',  Not Invisible ) ;
end;

procedure TOF_CPSUIVIMP.MarqueOrigine;
var TOBL : TOB ;
begin
  TOBL := TobMvt.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],
                           [Q.FindField('E_JOURNAL').AsString,Q.FindField('E_EXERCICE').AsString,
                            Q.FindField('E_DATECOMPTABLE').AsDateTime,Q.FindField('E_NUMEROPIECE').AsInteger,
                            Q.FindField('E_NUMLIGNE').AsInteger,Q.FindField('E_NUMECHE').AsInteger],False) ;
  if TOBL<>Nil then
    TOBL.PutValue('MARQUE','X') ;
end;

procedure TOF_CPSUIVIMP.PrechargeOrigines;
var QQ : TQuery ;
    st : String ;
    wh : string;
begin
  TobMvt.ClearDetail ;
  {b FP 13/07/2006 FQ18589}
  St := 'SELECT E_EXERCICE,E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE,E_QUALIFPIECE';
  // CA - 05/09/2007 - On n'utilise plus la vue de la compensation
  (*if IsCompensation(smp) then
    St := St + ' FROM CPSUIVICOMPEN LEFT JOIN GENERAUX ON E_GENERAL=GENERAUX.G_GENERAL '
  else*)
  {JP 25/10/07 : FQ 21732 : Pas de OR sur les jointures sous Oracle}
  if IsCompensation(smp) then begin
    if isMssql then
      St := St + ' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+
                 ' LEFT JOIN TIERS ON (T_AUXILIAIRE=E_AUXILIAIRE OR (E_AUXILIAIRE=T_PAYEUR AND E_CONTREPARTIEAUX=T_AUXILIAIRE AND E_PIECETP<>'')) '
    else
      St := St + ' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+
                 ' LEFT JOIN TIERS ON T_AUXILIAIRE=E_AUXILIAIRE ';
  end
  else
    St := St + ' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+
               ' LEFT JOIN TIERS ON T_AUXILIAIRE=E_AUXILIAIRE ';
  {e FP 13/07/2006}
  wh := RecupWhereCritere(TPageControl(GetControl('Pages',True)));
  {La jointure sur les auxiliaires n'a d'intérêt que pour la gestion des tables libres sur les tiers}
  if (wh <> '') and IsCompensation(smp) and not isMssql then
    wh := wh + ' AND E_PIECETP IS NULL ';
    
  QQ := OpenSQL( St + wh , True) ;
  TobMvt.LoadDetailDB( 'ECRITURE' , '' , '' , QQ , False , True ) ;
  Ferme( QQ ) ;

  {JP 25/10/07 : FQ 21732 : Pas de OR sur les jointures sous Oracle. Pour la compensation avec Tiers
                 payeurs, on relance une autre requête}
  if IsCompensation(smp) and not isMssql and (wh <> '') then begin
    St := 'SELECT E_EXERCICE,E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE,E_QUALIFPIECE';
    St := St + ' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL'+
               ' LEFT JOIN TIERS ON (E_AUXILIAIRE=T_PAYEUR AND E_CONTREPARTIEAUX=T_AUXILIAIRE) ';

    wh := RecupWhereCritere(TPageControl(GetControl('Pages',True)));
    if wh <> '' then
      wh := wh + ' AND E_PIECETP IS NOT NULL ';
    QQ := OpenSQL( St + wh , True) ;
    TobMvt.LoadDetailDB( 'ECRITURE' , '' , '' , QQ, True, True ) ;
    Ferme( QQ ) ;
  end;

  if TobMvt.Detail.Count>0 then
    TobMvt.Detail[0].AddChampSup('MARQUE',True) ;
end;

procedure TOF_CPSUIVIMP.PreciseMP;
var E_MODEPAIE : THValComboBox ;
begin
  E_MODEPAIE := THValComboBox(GetControl('E_MODEPAIE',True)) ;
  if Categorie.itemIndex = 0
    then CatToMP ( '', E_MODEPAIE.Items, E_MODEPAIE.Values, tslAucun, True)
    else CatToMP ( CATEGORIE.Value, E_MODEPAIE.Items, E_MODEPAIE.Values, tslAucun, True ) ;
end;

procedure TOF_CPSUIVIMP.ReinitWhereNatCpt( RAZ : Boolean ) ;
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
      {b FP 21/02/2006}
      if IsCompensation(smp)
        then SetControlText('XX_WHERENATCPT','E_AUXILIAIRE<>""'+
             ' AND (T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC"' +
                   ' OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD")' +
             ' AND ('+TCompensation.GetChampPlan+'<>"")')
      else
      {e FP 21/02/2006}
      if IsEnc(SMP)
        then SetControlText('XX_WHERENATCPT','E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"FOU" AND T_NATUREAUXI<>"AUC"')
        else SetControlText('XX_WHERENATCPT','E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"CLI" AND T_NATUREAUXI<>"AUD"');
      end ;
    end ;

end;

procedure TOF_CPSUIVIMP.ToutMarquer;
var i : integer ;
    TOBL : TOB ;
begin
  for i:=0 to TobMvt.Detail.Count-1 do
    begin
    TOBL:=TobMvt.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    end ;
end;

procedure TOF_CPSUIVIMP.ValideLot;
Var i : Integer ;
    TobL : TOB ;
    St : String ;
begin
  InitMove(TobMvt.Detail.Count,'') ;
  for i:=TobMvt.Detail.Count-1 downto 0 do
    begin
    MoveCur(FALSE) ;
    TOBL:=TobMvt.Detail[i] ;
    St:='UPDATE ECRITURE SET E_NOMLOT="'+CodeLot+'" WHERE E_JOURNAL="'+TOBL.GetValue('E_JOURNAL')+'" '
       +' AND E_EXERCICE="'+TOBL.GetValue('E_EXERCICE')+'" AND E_DATECOMPTABLE="'+UsDateTime(TOBL.GetValue('E_DATECOMPTABLE'))+'" '
       +' AND E_NUMEROPIECE='+IntToStr(TOBL.GetValue('E_NUMEROPIECE'))+' AND E_NUMLIGNE='+IntToStr(TOBL.GetValue('E_NUMLIGNE'))
       +' AND E_NUMECHE='+IntToStr(TOBL.GetValue('E_NUMECHE'))
       +' AND E_QUALIFPIECE="'+TOBL.GetValue('E_QUALIFPIECE')+'" ' ;
    ExecuteSQL(St) ;
    END ;
  FiniMove ;
  AvertirTable('CPNOMLOT');
end;

procedure TOF_CPSUIVIMP.CATEGORIEChange(Sender: TObject);
begin
  { Le 28/01/2003 BPY correction des bug de la fiche 11794 }
  if (Categorie.Value = 'TRI') then
    BCtrlRib.visible := false;
  { Fin BPY }
  PreciseMP ;
end;

{JP 13/09/05 : FQ 16395 : Je remplace E_EXERCICEExit par E_EXERCICEChange pour que l'évènement
               ne soit pas systématiquement exécuté lors de la recherche}
procedure TOF_CPSUIVIMP.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates( GetControlText('E_EXERCICE') , THEdit(GetControl('E_DATECOMPTABLE', True)) ,
                                             THEdit(GetControl('E_DATECOMPTABLE_', True))  ) ;
end;

procedure TOF_CPSUIVIMP.FTICExit(Sender: TObject);
begin
  SetControlText('E_GENERAL', '') ;
  if FTIC.Checked then
    begin
    SetControlEnabled('_E_AUXILIAIRE', FALSE ) ;       {FP 15/04/2006 FQ17967}
    SetControlEnabled('TE_AUXILIAIRE', FALSE ) ;
    SetControlText('_E_AUXILIAIRE', '' ) ;             {FP 15/04/2006 FQ17967}
    if IsEnc(smp)
      then SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGTID')
      else SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGTIC') ;
    end
  else
    begin
    SetControlEnabled('_E_AUXILIAIRE', TRUE ) ;        {FP 15/04/2006 FQ17967}
    SetControlEnabled('TE_AUXILIAIRE', TRUE ) ;
    if IsEnc(smp)
      then SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGENCAIS')
      else SetControlProperty('E_GENERAL', 'DATATYPE', 'TZGDECAIS') ;
    end ;
end;

procedure TOF_CPSUIVIMP.XNOMLOT1Exit(Sender: TObject);
begin
  if AvecXNomLot then
    CRechVide.Enabled := TRUE
  else
    begin
    CRechVide.Enabled := FALSE ;
    CRechVide.Checked := FALSE ;
    end ;
end;

procedure TOF_CPSUIVIMP.TraiteEditeExit(Sender: TObject);
var lBoEtat : Boolean ;
begin
  lBoEtat := TraiteEdite.State=cbChecked ;
  SetControlEnabled('E_NUMTRAITECHQ_', lBoEtat ) ;
  SetControlEnabled('FE_NUMTRAITECHQ_', lBoEtat ) ;
  SetControlEnabled('E_NUMTRAITECHQ', lBoEtat ) ;
  SetControlEnabled('FE_NUMTRAITECHQ', lBoEtat ) ;

  SetControlText('E_NUMTRAITECHQ_', '' ) ;
  SetControlText('E_NUMTRAITECHQ', '' ) ;
end;

procedure TOF_CPSUIVIMP.BCtrlRibClick(Sender: TObject);
Var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(TPageControl(GetControl('Pages',True))) ;
  if (StWRib = '') then Exit;
  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
  if ({$IFDEF EAGLLCIENT}
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
      {$IFDEF EAGLLCIENT}
        Q.TQ.Seek(FListe.row - 1) ;
      {$ENDIF}
      StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
      end;
    // Efface le dernier OR et rajoute ')'
    StWRib := Copy( StWRib, 1, length(StWRib) - 2 );
    StWRib := StWRib +')';
    end;
  if (Categorie.Value = 'TRI')
    then CPLanceFiche_VerifRib('IBAN;WHERE='+StWRib)
    else CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

procedure TOF_CPSUIVIMP.BKillLotClick(Sender: TObject);
begin
  DetruitLeLot(smp) ;
  AvertirTable('CPNOMLOT');
end;

procedure TOF_CPSUIVIMP.BRAZLotClick(Sender: TObject);
begin
  if HShowMessage( HDiv[33], gszCaption, '') <> mrYes then Exit ;
  CodeLot:='' ;
  ConstitueTobMvt ;
  if Transactions(ValideLot,1)<>oeOk then
    begin
    MessageAlerte(HDiv[32]) ;
    end ;
  TFMul(Ecran).BChercheClick(Nil) ;
  AvertirTable('CPNOMLOT');
end;

{$IFDEF EAGLCLIENT}
procedure TOF_CPSUIVIMP.FListeDblClick(Sender: TObject);
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
      PGIInfo('Saisie borderau non implémentée en mode CWAS', Ecran.Caption )
      // A FAIRE Voir SAISBOR.PAS
    else
      begin
      TOBListe := TOB.Create('Liste Ecriture',nil,-1) ;
      TOBLigne := Q.TQ.Detail[FListe.row-1] ;
      TOB.Create('ECRITURE',TOBListe,-1) ;
      TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
      TrouveEtLanceSaisie( TOBListe, taConsult, 'N') ;
      end ;
  {$ENDIF}
end;
{$ELSE}
procedure TOF_CPSUIVIMP.FListeDblClick(Sender: TObject);
var
  sMode : String ;
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

procedure TOF_CPSUIVIMP.BOuvrirClick(Sender: TObject);
var Abrege,Libre : String ;
begin

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
      PGIInfo( GetYenATropMsg(smp) ,gszCaption ) ;
      exit ;
      end ;
  // Message avertissement si sélection inversée
	if swapSelect then
  	if PGIAsk('Vous êtes en mode "sélection inversée". Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes
    then  exit;

  CodeLot:='' ;
  Abrege:='' ;
  Libre:='' ;
  Abrege := AttribListe(Smp) ;{JP 29/11/05 : FQ 17030}
  {b FP 21/02/2006: Pour la compensation les listes sont identiques qq soit l'axe}
  if (smp = smpCompenCli) then
    Abrege := 'CPCOMPENSCLI'
  else if (smp = smpCompenFou) then
    Abrege := 'CPCOMPENSFOU';
  {e FP 21/02/2006}

  if SaisieCodelot(CodeLot,smp,Abrege,'') then
    begin
    ConstitueTobMvt ;
    if Transactions(ValideLot,1)<>oeOk then
      MessageAlerte(HDiv[32]) ;
    TFMul(Ecran).BChercheClick(Nil) ;
    end ;
end;

{$IFNDEF EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 27/10/2005
Modifié le ... :   /  /
Description .. : Report de l'évènement OnDrawDataCell, non déclenché sur
Suite ........ : la Grid
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSUIVIMP.FListeDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var sRIB, St : String ;
begin
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

end;

// BPY le 18/10/2004 : demande de RR => trie des liste dans CCMP
procedure TOF_CPSUIVIMP.FListeSortColumn(Sender: TObject);
begin
  FListe.ClearSelected;
end;
// Fin BPY
{$ENDIF}



procedure TOF_CPSUIVIMP.FListeRowEnter(Sender: TObject);
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

procedure TOF_CPSUIVIMP.bSelectAllClick(Sender: TObject);
begin
  FetchLesTous;
  if Assigned(bSelectAllClickParent) then
    bSelectAllClickParent(Sender) ;

  // Si on sélectionne tout, on annule le mode "sélection inversée"
  if FListe.AllSelected then
  	begin
    SwapSelect := False ;
    BSwapSelect.Down := False ;
		BSwapSelect.Hint := 'Activer le mode "sélection inversée"' ;
    end;
	AfficheSoldeSelection ;
end;

procedure TOF_CPSUIVIMP.BSwapSelectClick(Sender: TObject);
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

procedure TOF_CPSUIVIMP.FListeFlipSelection(Sender: TObject);
begin
	AfficheSoldeSelection ;
end;

procedure TOF_CPSUIVIMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Assigned(FormKeyDownParent) then
    FormKeyDownParent(Sender, Key, Shift);
  Case Key of
    VK_F5 : begin Key:=0 ; ClickModifRib ; end
  end ;
end;

procedure TOF_CPSUIVIMP.InitConditions;
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
  SetControlText('XX_WHEREAN', St ) ;

  // Conditions nature de compte et montants        XX_WHERENATCPT
  ReinitWhereNatCpt(FALSE) ;

  // Condition sur le numéro de traite              XX_WHERETRACHQ
  SetControlText('XX_WHERETRACHQ', '') ;

  // Condition sur ???                              XX_WHEREMP
  SetControlText('XX_WHEREMP', '') ;

  // Condition sur les lots                         XX_WHERELOT
  if smp=smpDecTous
    then SetControlText('XX_WHERELOT','E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""')
    else if smp=smpEncTous
      then SetControlText('XX_WHERELOT','E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' ) ;

  // XX_WHEREENC                                                                     
  {b FP 21/02/2006}
  if IsCompensation(smp) then
    SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="ENC" OR E_ENCAISSEMENT="DEC")'+
                   ' AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF")' )
  else
  {e FP 21/02/2006}
  if IsEnc(SMP)
    then if cFactCredit.Checked
           then SetControlText('XX_WHEREENC','E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC" OR E_NATUREPIECE="OD"))' )  {FP 25/04/2006 FQ17814: Ajout de la nature OD pour la compensation}
           else SetControlText('XX_WHEREENC','E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"))' )
    else if cFactCredit.Checked
           then SetControlText('XX_WHEREENC','E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF" OR E_NATUREPIECE="OD"))')  {FP 25/04/2006 FQ17814: Ajout de la nature OD pour la compensation}
           else SetControlText('XX_WHEREENC','E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"))') ;

  // XX_WHEREMONTANT
  StXP  := StrFPoint(  9 * Resolution( V_PGI.OkDecV + 1 ) ) ;
  StXN  := StrFPoint( -9 * Resolution( V_PGI.OkDecV + 1 ) ) ;
  StXP2 := StrFPoint(  9 * Resolution( V_PGI.OkDecE + 1 ) ) ;
  StXN2 := StrFPoint( -9 * Resolution( V_PGI.OkDecE + 1 ) ) ;
  St    := '(E_DEBIT+E_CREDIT-E_COUVERTURE not between ' + StXN + ' AND ' + StXP + ')';
  SetControlText('XX_WHEREMONTANT', St ) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPSUIVIMP.UpdateConditions;
{---------------------------------------------------------------------------------------}

    {JP 06/06/07 : FQ 18309 : gestion du format SQL des montants
    {-----------------------------------------------------------}
    function _GetValeur(Chp : string) : string;
    {-----------------------------------------------------------}
    begin
      Result := StrfPoint(Valeur(GetControlText(Chp)));
    end;

var
  sztemp: String;           {FP 25/04/2006 FQ17939}
  {b FP 25/04/2006 FQ17967}
  ChpClient:     String;
  ChpFourn:      String;
  ChpCorres:     String;
  {e FP 25/04/2006}
begin
  {JP 13/09/05 : FQ 16395 : On vérifie la cohérance des dates et de l'exercice}
  if (QuelExo(GetControlText('E_DATECOMPTABLE' )) <> GetControlText('E_EXERCICE')) or
     (QuelExo(GetControlText('E_DATECOMPTABLE_')) <> GetControlText('E_EXERCICE')) then
    E_EXERCICEChange(GetControl('E_EXERCICE'));

  // XX_WHEREMONTANT et XX_WHEREAN sont initialisé une fois
  //  pour toute dans le initConditions

  // XX_WHEREENC
  {b FP 21/02/2006}
  if IsCompensation(smp) then
    SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="ENC" OR E_ENCAISSEMENT="DEC")'+
               ' AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC" OR E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF")' )
  else
  {e FP 21/02/2006}
  if isEncMP(smp) then
    if cFactCredit.Checked
      then SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC" OR E_NATUREPIECE="OD")))' )  {FP 25/04/2006 FQ17814: Ajout de la nature OD pour la compensation}
      else SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC")))')     {FP 25/04/2006 FQ17939 Ajout de parenthèses
                                                                                                                                                JP 14/06/06 : FQ 18156 : manquait une parenthèse}
  else
    if cFactCredit.Checked
      then SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF" OR E_NATUREPIECE="OD")))' )  {FP 25/04/2006 FQ17814: Ajout de la nature OD pour la compensation}
      else SetControlText('XX_WHEREENC','(E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF")))' ) ; {FP 25/04/2006 FQ17939 Ajout de parenthèses}


  {b FP 25/04/2006 FQ17939: Ajoute une condition OR pour les filtres sur le débit et le crédit}
  if IsCompensation(smp) then
    begin
    sztemp := GetControlText('XX_WHEREENC');
    if (szTemp <> '') and ((GetControlText('_E_DEBIT_') <> '') or (GetControlText('_E_CREDIT_') <> '')) then
      sztemp := sztemp+' AND';
    if (GetControlText('_E_DEBIT_') <> '') and (GetControlText('_E_CREDIT_') <> '') then
      sztemp := sztemp+' (';
    if (GetControlText('_E_DEBIT_') <> '') then
      sztemp := sztemp+' (E_DEBIT>='+_GetValeur('_E_DEBIT')+' AND E_DEBIT<='+_GetValeur('_E_DEBIT_')+')';
    if (GetControlText('_E_DEBIT_') <> '') and (GetControlText('_E_CREDIT_') <> '') then
      sztemp := sztemp+' or';
    if (GetControlText('_E_CREDIT_') <> '') then
      sztemp := sztemp+' (E_CREDIT>='+_GetValeur('_E_CREDIT')+' AND E_CREDIT<='+_GetValeur('_E_CREDIT_')+')';
    if (GetControlText('_E_DEBIT_') <> '') and (GetControlText('_E_CREDIT_') <> '') then
      sztemp := sztemp+')';
    SetControlText('XX_WHEREENC',sztemp);
    end
  else
    begin
    sztemp := GetControlText('XX_WHEREENC');
    if (szTemp <> '') and (GetControlText('_E_DEBIT_') <> '') then
      sztemp := sztemp+' AND';
    if (GetControlText('_E_DEBIT_') <> '') then
      sztemp := sztemp+' E_DEBIT>='+_GetValeur('_E_DEBIT')+' AND E_DEBIT<='+_GetValeur('_E_DEBIT_');
    if (szTemp <> '') and (GetControlText('_E_CREDIT_') <> '') then
      sztemp := sztemp+' AND';
    if (GetControlText('_E_CREDIT_') <> '') then
      sztemp := sztemp+' E_CREDIT>='+_GetValeur('_E_CREDIT')+' AND E_CREDIT<='+_GetValeur('_E_CREDIT_');
    SetControlText('XX_WHEREENC',sztemp);
    end;
  {e FP 25/04/2006}

  {b FP 25/04/2006 FQ17967}
  sztemp    := GetControlText('XX_WHEREENC');
  if (szTemp <> '') and (GetControlText('_E_AUXILIAIRE')<>'') then
    sztemp := sztemp+' AND';
  if GetControlText('_E_AUXILIAIRE')<>'' then
    begin
    if IsCompensation(smp) then
      begin
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
      sztemp := sztemp+' (T_AUXILIAIRE LIKE "'+GetControlText('_E_AUXILIAIRE')+'%"'+  {FP 13/07/2006 FQ18589}
                       ' or T_AUXILIAIRE in (select '+ChpFourn+' from CORRESP'+       {FP 13/07/2006 FQ18589}
        ' WHERE '+ChpCorres+' = CR_CORRESP AND CR_TYPE = "AU'+IntToStr(TCompensation.GetNumPlanCorrespondance)+'"'+
        ' AND '+ChpClient+' LIKE "'+GetControlText('_E_AUXILIAIRE')+'%"))';
      end
    else
      begin
      sztemp := sztemp+' E_AUXILIAIRE LIKE "'+GetControlText('_E_AUXILIAIRE')+'%"';
      end;
    end;
  SetControlText('XX_WHEREENC',sztemp);
  {e FP 25/04/2006}

  // XX_WHERELOT
  GereXNomLot ;

  // XX_WHERETRACHQ
  if TraiteEdite.State=cbGrayed
    then SetControlText('XX_WHERETRACHQ', '')
    else if TraiteEdite.State=cbChecked
      then SetControlText('XX_WHERETRACHQ', 'E_NUMTRAITECHQ<>""')
      else SetControlText('XX_WHERETRACHQ', 'E_NUMTRAITECHQ=""' );   // cbUnChecked

  // XX_WHERENATCPT
  ReinitWhereNatCpt( ( Trim(GetControlText('E_GENERAL')) <> '' ) or
                     ( Trim(GetControlText('_E_AUXILIAIRE')) <> '' ) ) ;  // FQ19410 : SBO 04/01/2007

  // XX_WHEREMP
  LaCategorieChange( THValCombobox( GetControl('CATEGORIE', True) ),
                     THEdit(GetControl('XX_WHEREMP',True)),
                     SorteLettre ) ;

end;

function TOF_CPSUIVIMP.AvecXNomLot: Boolean;
begin
  Result := (XNOMLOT1.Text<>'') or (XNOMLOT2.Text<>'') ;
end;

procedure TOF_CPSUIVIMP.InitComposants;
begin
 	Q               := THQuery(GetControl('Q', True)) ;
{$IFDEF EAGLCLIENT}
  FListe          := THGrid(GetControl('FListe', True));
{$ELSE}
  FListe          := THDBGrid(GetControl('FListe', True));
{$ENDIF}
  // -> Composants gérés manuellement
  FTIC            := TCheckBox(GetControl('FTIC', True)) ;
  CATEGORIE       := THValComboBox(GetControl('CATEGORIE', True)) ;
  cFactCredit     := TCheckBox(GetControl('cFactCredit', True)) ;
  TraiteEdite     := TCheckBox(GetControl('TraiteEdite', True)) ;
  XNOMLOT1        := THEdit(GetControl('XNOMLOT1', True)) ;
  XNOMLOT2        := THEdit(GetControl('XNOMLOT1_', True)) ;
  CRechVide       := TCheckBox(GetControl('CRechVide', True)) ;
  // Boutons Toolbar
  BCtrlRib        := TToolbarButton97(GetControl('BCtrlRib', True)) ;
  BRAZLot         := TToolbarButton97(GetControl('BRAZLot', True)) ;
  BKillLot        := TToolbarButton97(GetControl('BKillLot', True)) ;
  BSwapSelect     := TToolbarButton97(GetControl('BSwapSelect', True)) ;

{b fb 19/06/2006 FQ12487 + FP 13/07/2006 FQ18589}
  Selection       := THEdit(GetControl('E_GENERAL', true));
  Auxiliaire      := THEdit(GetControl('_E_AUXILIAIRE', true));
{e fb 19/06/2006 FQ12487}
end;

procedure TOF_CPSUIVIMP.InitEvenements;
var lBouton : TButton ;
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
  if Assigned(FListe.OnDrawColumnCell)
    then FListeDrawColumnCellParent := FListe.OnDrawcolumnCell
    else FListeDrawColumnCellParent := nil ;

FListe.OnDrawColumnCell    := FListeDrawColumnCell ;
  // BPY le 18/10/2004 : demande de RR => trie des liste dans CCMP
  FListe.SortEnabled       := true;
  FListe.OnSortColumn      := FListeSortColumn ;
  // Fin BPY
{$ENDIF}
  FListe.OnDblClick        := FListeDblClick ;
  FListe.OnFlipSelection   := FListeFlipSelection ;

{b fb 19/06/2005 FQ12487  + FP 13/07/2006 FQ18589
  Selection.OnElipsisClick := SelectionElipsisClick;
  Auxiliaire.OnElipsisClick := AuxiliaireElipsisClick;
e fb 19/06/2005 FQ12487}

  // Affectation EVT Boutons
    // -> tout sélectionner
    lBouton := TButton(GetControl('bSelectAll',True)) ;
    if Assigned(lBouton.OnClick)
      then bSelectAllClickParent    := lBouton.OnClick
      else bSelectAllClickParent    := nil ;
    lBouton.OnClick := bSelectAllClick ;
    // -> Ouvrir
    TButton(GetControl('bOuvrir',True)).OnClick    := bOuvrirClick ;
    // -> Nouvelle recherche
    {JP 19/08/04 : FQ 14212 : ce n'est pas utile et qui plus est le transtypage est erroné car
                              il s'agit d'un TMenuItem
    lBouton := TButton(GetControl('BNouvRech',True)) ;
    if Assigned(lBouton.OnClick)
      then BNouvRechClickParent    := lBouton.OnClick
      else BNouvRechClickParent    := nil ;
    lBouton.OnClick := BNouvRechClick ;
  }
  // Ftn spécifique
  BCtrlRib.OnClick      := BCtrlRibClick ;
  BRazLot.OnClick       := BRazLotClick ;
  BKillLot.OnClick      := BKillLotClick ;
  BSwapSelect.OnClick   := BSwapSelectClick ;

  // Affectation EVT Filtre
  {JP 13/09/05 : FQ 16395 : Je remplace E_EXERCICEExit par E_EXERCICEChange pour que l'évènement
                 ne soit pas systématiquement exécuté lors de la recherche
  THValComboBox(GetControl('E_EXERCICE',  True)).OnExit := E_EXERCICEExit;}
  THValComboBox(GetControl('E_EXERCICE',  True)).OnChange := E_EXERCICEChange;

  CATEGORIE.OnChange := CATEGORIEChange; {FQ 14212}
  FTIC.OnExit := FTICExit ;
  TraiteEdite.OnExit := TraiteEditeExit ;
  XNOMLOT1.OnExit := XNOMLOT1Exit ;
  XNOMLOT2.OnExit := XNOMLOT1Exit ;
  // EVT Form
  if Assigned(TFMul(Ecran).OnKeyDown)
    then FormKeyDownParent := TFMul(Ecran).OnKeyDown
    else FormKeyDownParent := nil ;
  TFMul(Ecran).OnKeyDown := FormKeyDown ;
end;

procedure TOF_CPSUIVIMP.InitHelpContext;
begin
  TFMul(Ecran).HelpContext := AttribHelp(smp,TRUE) ;
end;

procedure TOF_CPSUIVIMP.InitVariables;
begin
  TOBMvt      := TOB.Create('',Nil,-1) ;
  // Gestion du mode "selection inversé", non opérationnel par défaut
  SwapSelect  := False ;
  // ??
  VH^.MPModifFaite:=FALSE ;
  Fillchar(VH^.MPPOP,SizeOf(VH^.MPPOP),#0) ;
end;

procedure TOF_CPSUIVIMP.InitCATEGORIE;
begin
  Case smp of
    smpEncDiv,
    smpDecDiv                 : CATEGORIE.ItemIndex := 0 ;

    smpEncPreEdt,smpEncPreBqe : CATEGORIE.Value := 'PRE' ;

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

  PreciseMP ;
end;

procedure TOF_CPSUIVIMP.FetchLesTous;
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

{b fb 19/06/2006 FQ12487}
procedure TOF_CPSUIVIMP.SelectionElipsisClick(Sender : TObject);
begin
  if isEnc(smp) then
    LookUpList(TControl(Sender),'Recherche d''un compte collectif','GENERAUX','G_GENERAL','G_LIBELLE',
    '(G_SUIVITRESO="ENC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" OR G_LETTRABLE="X")','G_GENERAL',true,-1)
  else
    LookUpList(TControl(Sender),'Recherche d''un compte collectif','GENERAUX','G_GENERAL','G_LIBELLE',
    '(G_SUIVITRESO="DEC" OR G_SUIVITRESO="MIX") AND (G_COLLECTIF="X" OR G_LETTRABLE="X")','G_GENERAL',true,-1);
end;

procedure TOF_CPSUIVIMP.AuxiliaireElipsisClick(Sender : TObject);
begin
  if isEnc(smp) then
    LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
    '(T_NATUREAUXI="AUD" OR T_NATUREAUXI="CLI" OR T_NATUREAUXI="DIV")','T_AUXILIAIRE',true,-1)
  else
    LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE',
    '(T_NATUREAUXI="AUC" OR T_NATUREAUXI="FOU" OR T_NATUREAUXI="DIV" OR T_NATUREAUXI="SAL")','T_AUXILIAIRE',true,-1);
end;
{e fb 19/06/2006 FQ12487}

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPSUIVIMP.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


function TOF_CPSUIVIMP.YEnATrop : boolean ;
var lQCount : TQuery ;
    lQte    : integer ;
begin
  result := True ;

  // Mode inversé
  if swapSelect then
    begin
    if (not FListe.AllSelected) then
      begin
      lQCount := OpenSQL( 'SELECT COUNT(*) QTE FROM SUIVIMP ' + RecupWhereCritere( TPageControl(GetControl('Pages',True)) ) , True ) ;
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
      lQCount := OpenSQL( 'SELECT COUNT(*) QTE FROM SUIVIMP ' + RecupWhereCritere( TPageControl(GetControl('Pages',True)) ) , True ) ;
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
  registerclasses ( [ TOF_CPSUIVIMP ] ) ;
end.

