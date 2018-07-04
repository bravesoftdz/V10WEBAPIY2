{***********UNITE*************************************************
Auteur  ...... : SBO
Créé le ...... : 19/03/2007
Modifié le ... : 29/03/2007
Description .. : Source TOF de la FICHE : CPLOIVENTIL ()
Suite ........ : 
Suite ........ : Traitement de répartition secondaire. La même fiche est
Suite ........ : utilisée pour créer les lois de répartition et pour exécuter les
Suite ........ : traitement.
Suite ........ : Le tout est stocké dans CLOIVENTIL
Suite ........ :   CLV_LOITYPE = LOI  --> Loi de répartition
Suite ........ :   CLV_LOITYPE = SEC --> Répartition secondaire
Mots clefs ... : TOF;CPLOIVENTIL
*****************************************************************}
Unit CPLOIVENTIL_TOF ;

Interface

Uses StdCtrls,
     Controls,          // TKeyEvent
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     SaisUtil,             // GetNewNumJal, RDevise
     uLibEcriture,
     uTob,
     UTOF ;


function CPLanceFiche_LoiVentil( vStRange, vStLequel, vStArgs : string ) : string;

Type
  TOF_CPLOIVENTIL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure OnAfterFormShow          ;
  private

    // Composants
    FListe           : THGrid ;

    // Enregistrement
    FTobLoi          : Tob ;
    FInfo            : TInfoECriture ;
    FTobODA          : Tob ;
    FTobCumuls       : Tob ;

    // Evt AGL
    FOldKeyDown      : TKeyEvent ;

    // Indicateur
    FStType          : string ;
    FInNumero        : integer ;
    FStAxe           : string ;

    FStSection       : string ;
    FDtDateGener     : TDateTime ;
    FSoldeSection    : Double ;
    FSoldeTraite     : Double ;

    FBoGrilleModif   : boolean ;
    FBoOnLoad        : boolean ;
    FBoDateOk        : boolean ;

    FDevise          : RDevise ;

    // ======================
    // === Initialisation ===
    // ======================
    procedure AnalyseParametre      (vStArg: string);
    procedure InitComposants        ;
    procedure InitVariables         ;
    procedure InitEvenements        ;
    procedure InitEcran             ;
    procedure InitGrille            ;

    // ===============
    // === Données ===
    // ===============
    procedure GetEcran              ;
    procedure PutEcran              ;
    procedure NewRecord             ;
    function  LoadRecord            : boolean ;
    procedure ChargeODA             ( vFromTob: Tob = nil ) ;
    function  _SaveRecord           : boolean ;
    function  SaveRecord            : boolean ;
    function  TestRecord            : boolean ;
    procedure UpdateRecord          ;
    procedure UpdateNumero          ( vInNum : integer ) ;
    function  EstModifEnCours       : boolean ;
    function  GetRefVentil          ( vInNum : integer ; vStType : string = '' ) : string ;
    function  GetMontantCalc        : Double ;
    procedure SetMontantCalc        ( vMontant : Double ) ;
    function  GetMontantTotal       : Double ;
    function  GetMontantTotalAbs    : Double ;
    function  GetSens               ( vMontant : Double = 0 ) : String ;
    function  GetValeur             ( vStControl : string ) : Double ;
    procedure SetValeur             ( vStControl : string ; vMontant : double ) ;

    // =================
    // === Affichage ===
    // =================
    procedure AfficheTob            ;
    procedure Accesboutons          ;
    procedure LayoutCumuls          ;
    procedure SetDataType           ( vStAxe : string ) ;
    function  GetAction             : TActionFiche ;
    procedure SetAction             ( vAction : TActionFiche ) ;
    procedure AfficheLesMontants    ;
    procedure RecalculTableau       ( vTotal : Double ) ;
    function  GetNumDisp            : integer ;

    // ==================
    // === Evènements ===
    // ==================
    // --> Evènements de la fiche
    procedure FormKeyDown           ( Sender : TObject ; var Key: Word ; Shift: TShiftState ) ;
    procedure FormCloseQuery        ( Sender : TObject ; var CanClose: Boolean);
    procedure EcranClose            ;

    // --> Evènements de gestion de la grille
    procedure BLigneInsertClick     ( Sender : TObject );
    procedure BLigneSupprClick      ( Sender : TObject );
    procedure FListeDblClick        ( Sender : TObject );
    procedure FListeCellExit        ( Sender : TObject; var ACol, ARow: Longint; var Cancel: Boolean );
    procedure FListeCellEnter       ( Sender : TObject; var ACol, ARow: Longint; var Cancel: Boolean );
    procedure FListeKeyDown         ( Sender : TObject; var Key: Word; Shift: TShiftState );
    procedure FListeElipsisClick    ( Sender : TObject );

    // --> Evènements boutons
    procedure BFermeClick           ( Sender : TObject );
    procedure BCopierClick          ( Sender : TObject );
    procedure BCalculerClick        ( Sender : TObject );
    procedure BRecalculerClick      ( Sender : TObject );
    procedure BGenererClick         ( Sender : TObject );
    procedure BVentilTypeClick      ( Sender : TObject );
    {$IFNDEF CCSTD}
    procedure BZoomODAClick         ( Sender : TObject ) ;
    {$ENDIF}
    procedure BImprimerClick        ( Sender : TObject ) ;

    // Evènement des critères
    procedure AxeChanged            ( Sender : TObject );
    procedure SectionChanged        ( Sender : TObject );
    procedure SectionExit           ( Sender : TObject );
    procedure LoiChanged            ( Sender : TObject );
    procedure DateGenerChanged      ( Sender : TObject );
    procedure DateGenerExit         ( Sender : TObject );
    procedure LoiMethodeChanged     ( Sender : TObject );
    procedure MontantCorrChanged    ( Sender : TObject );
    procedure LoiVentilClick        ( Sender : TObject );
    procedure JournalChanged        ( Sender : TObject );

    // ============================
    // === Gestion de la grille ===
    // ============================
    procedure InsertRow             ( vRow : integer ) ;
    procedure DeleteRow             ( vRow : integer ) ;
    function  GetDecimale           ( vCol : integer ) : integer ;
    function  GetMontant            ( vCol, vRow : integer ) : double ;
    procedure setMontant            ( vCol, vRow : integer ; vMontant : double) ;
    procedure RecalculLigne         ( vRow : integer ) ;
    procedure RecalculGrille        ( vRow : integer ) ;
    procedure GetColSuivante        ( var ACol, ARow : integer ; vInSens : integer = 0 ) ;
    procedure SoldeVentil           ( vRow : integer ) ;
    function  IsOut                 ( vRow : integer ) : Boolean ;
    function  GetTotalCol           ( vCol : integer ) : Double ;
    procedure SetFormatGrille       ( vStMeth : string ) ;
    procedure AfficheGrille         ( vTobG : Tob ; vBoFromVT : boolean = False ; vBoFromLoi : boolean = False  ) ;
    procedure ResetGrille           ;
    procedure GereAccesGrille       ;
    procedure GrilleEnabled         ( vBoAvec : Boolean );
    function  EstGrilleEnabled      : boolean ;
    procedure GereArrondi           ( vRow : integer ) ;
    procedure InitLigne             ( vRow : integer ) ;
    procedure NumeroteLignes        ;
    procedure UpdateClefs           ;

    // ==================
    // === Traitement ===
    // ==================
    function  ClickValide           : boolean;
    function  PretACalculer         : Boolean ;
    function  PretAGenerer          ( vBoParle : boolean = True ) : boolean ;
    procedure CalculMontants        ;
    function  EstGrilleValide       ( vBoParle : Boolean = true ) : boolean;
    function  TestPresenceDate      ( vDtDate : TDateTime ; vBoForce : boolean = False ; vBoParle : boolean = True ) : boolean ;
    function  GetSoldeSection       ( vStSection : String ) : Double ;
    function  GetSoldeSectionParCpt ( vStSection : String ) : Double ;
    function  GetInitDate           : TDateTime ;
    function  GetDebutExo           : TDateTime ;
    function  GenerePieceODA        ( vStCptGeneral : string ; vMontantPiece : Double ; vTobParent : Tob = nil ) : TOB ;
    function  GenereLigneODA        ( vTobPiece : Tob ; vStCptGeneral : string ) : TOB ;
    procedure SetInfosODA           ;
    procedure SetCorrectionDansCumuls ;

  end ;



Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  HPanel ,
  {$IFDEF VER150} Variants, {$ENDIF}    // VarToStr
  {$IFDEF EAGLCLIENT}
     MainEagl,          // AglLanceFiche
  {$ELSE}
     FE_Main,           // AglLanceFiche
  {$ENDIF}
  HTB97,                // TToolBarButton97
  HSysMenu,
  lookup,
  grids,                // TGridDrawState
  Choix,                // choisir
  AglInit,
  uLibExercice,
  uLibWindows,
  Vierge,
  DateUtils,
  Windows,
  Messages,             // WM_CLOSE
  CPSECTION_TOM,        // FicheSection
  {$IFNDEF CCSTD}
  DelVisuE,
  SaisODA,              // LanceSaisieODA
  SaisComm,             // MvtToIdent
  {$ENDIF}
  utilSais,             // MajSoldeSectionTOB
  utilPGI,              // GetSelectAll
  uLibANalytique,       // CPutDefautAna
  uObjEtats,            // TObjEtats.GenereEtatGrille
  Ent1;                 // GetInfoCpta

Const  _NbChamps = 7 ;
       _TabMsgChamps : Array[ 1.._NbChamps ] of string = (
                     {1} 'CLV_AXE',
                     {2} 'CLV_SECTION',
                     {3} 'CLV_LIBELLE',
                     {4} 'CLV_LOINIVEAU',
                     {5} 'CLV_LOIMETHODE',
                     {6} 'CLV_SECTIONATT',
                     {7} 'CLV_JOURNAL'
        ) ;
       _TabMsgErreur : Array[ 1.._NbChamps ] of string = (
                     {1} 'Vous devez sélectionner un axe.',
                     {2} 'Vous devez sélectionner une section.',
                     {3} 'Vous devez saisir un libellé.',
                     {4} 'Vous devez sélectionner un niveau de répartition.',
                     {5} 'Vous devez sélectionner une règle de calcul.',
                     {6} 'Vous devez sélectionner une section d''équilibrage.',
                     {7} 'Vous devez sélectionner un journal de génération.'
        ) ;

        // Index des colonnes de la grille de répartition
        IDX_NO       = 1 ;
        IDX_SECTION  = 2 ;
        IDX_LIBELLE  = 3 ;
        IDX_CLEF     = 4 ;
        IDX_MONTANT  = 5 ;

function CPLanceFiche_LoiVentil( vStRange, vStLequel, vStArgs : string ) : string;
begin
  result := AGLLanceFiche('CP', 'CPLOIVENTIL', vStRange, vStLequel, vStArgs ) ;
end ;

procedure TOF_CPLOIVENTIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTIL.OnUpdate ;
begin
  Inherited ;
  LastError := 0 ;
  if not ClickValide
    then LastError := 1
    else FBoGrilleModif := False ;
end ;

procedure TOF_CPLOIVENTIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTIL.OnArgument (S : String ) ;
var lInTT : integer ;
    lStTT : string ;
begin
  Inherited ;

  // en attendant modif socref
  lStTT := 'CPLOIMODELE' ;
  lInTT := TTToNum( lStTT );
  if lInTT > 0 then
    V_PGI.DECombos[lInTT].ChampLib := 'CLV_SECTION || " - "  || CLV_LIBELLE' ;

  InitVariables ;
  AnalyseParametre(S);
  InitComposants;
  InitEvenements;
  InitEcran;

end ;

procedure TOF_CPLOIVENTIL.OnClose ;
begin
  Inherited ;

  if Assigned( FTobLoi ) then
    FreeAndnil( FTobLoi ) ;

  if Assigned( FInfo ) then
    FreeAndnil( FInfo ) ;

  if Assigned( FTobODA ) then
    FreeAndnil( FTobODA ) ;

  if Assigned( FTobCumuls ) then
    FreeAndnil( FTobCumuls ) ;

end ;

procedure TOF_CPLOIVENTIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTIL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPLOIVENTIL.AnalyseParametre(vStArg: string);
var
  lStChaine: string;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Type de répartition concernée
    if Pos('CLV_LOITYPE=', Uppercase(lStChaine)) > 0 then
      FStType := Copy(lStChaine, 13, length(lStChaine))

    else if Pos('CLV_AXE=', Uppercase(lStChaine)) > 0 then
      FStAxe := Copy(lStChaine, 9, length(lStChaine))

    else if Pos('CLV_SECTION=', Uppercase(lStChaine)) > 0 then
      FStSection := Copy(lStChaine, 13, length(lStChaine))

    else if Pos('CLV_NUMERO=', Uppercase(lStChaine)) > 0 then
      begin
      FInNumero := ValeurI( Copy(lStChaine, 12, length(lStChaine) ) ) ;
      end

    else if Pos('ACTION=', Uppercase(lStChaine)) > 0 then
      begin
      SetAction( StringToAction( lStChaine ) ) ;
      end
    ;

    end;

  if FStType='' then
    FStType := 'SEC' ;

  if FStAxe <> '' then
    begin
    SetControlEnabled( 'CLV_AXE', False ) ;
    SetDataType( FStAxe ) ;
    end ;

end;

procedure TOF_CPLOIVENTIL.InitComposants;
begin
  FListe := THGrid( GetControl('FLISTE', True) ) ;
  
//  TToolBarButton97( GetControl('BVALIDER', True) ).Default := False ;
//  THEdit( GetControl('CLV_DATETRAITEMENT', True) ).OpeType  := otString ;
//  THEdit( GetControl('CLV_DATETRAITEMENT', True) ).EditMask := '' ;
//  THEdit( GetControl('TCLV_NUMEROPIECE', True) ).Left := THEdit( GetControl('TCLV_NUMEROPIECE', True) ).Left - 10 ;
//  SetControlProperty( 'LOIVENTILLIB', 'LEFT', THEdit(GetControl('LOIVENTIL')).Left ) ;
//  SetControlProperty( 'BVENTILTYPE', 'HINT', 'Loi de répartition (Ctrl-L)' ) ; //FStType = 'LOI' ) ;

end;

procedure TOF_CPLOIVENTIL.InitEcran;
begin

  // ===== TITRE ====
  Ecran.ShowHint := True ;
  if FStType = 'LOI'
    then Ecran.caption := 'Lois de répartition'
    else Ecran.caption := 'Répartitions secondaires' ;
  UpdateCaption( Ecran ) ;

  // ===== CARACTERISTIQUES =====
  if GetAction<>taCreat then
    begin
    SetControlEnabled('CLV_AXE',      False ) ;
    SetControlEnabled('CLV_SECTION',  False ) ;
    end ;

  SetControlVisible('CLV_LOITYPE',        False ) ;

//  SetControlVisible('LOIVENTIL',      (FStType = 'SEC') and (GetAction=taCreat) ) ;
  SetControlVisible('LOIVENTIL',      False ) ;
  SetControlVisible('TLOIVENTIL',     False ) ;

  SetControlVisible('CLV_DATEGENER',  FStType = 'SEC' ) ;
  SetControlVisible('TCLV_DATEGENER', FStType = 'SEC' ) ;

  SetControlEnabled('CLV_STRICTE',    FStType = 'LOI' ) ;

  // ===== INFORMATIONS =====
  SetControlVisible('PINFOS',       FStType = 'SEC' ) ;
  SetControlText('TSAISIM',         'en ' + V_PGI.DevisePivot ) ;

  // ===== GRILLE =====
  GrilleEnabled(True) ;

  // ===== DERNIERTRAITEMENT =====
//  SetControlVisible('PPIED',          (FStType = 'SEC') and (GetAction<>taCreat) ) ;
  // zoom oda
  SetControlVisible('BZOOMODA',       False ) ;

  // === BOUTONS ===
  // Traitement
  SetControlVisible( 'BRECALCULER',  (FStType = 'SEC') and (GetAction=taModif) ) ;
  SetControlVisible( 'BCALCULER',    False ) ; //(FStType = 'SEC') and (GetAction=taCreat) ) ;
  SetControlVisible( 'BGENERER',     FStType = 'SEC' ) ;

  // loi de répartition
  SetControlVisible(  'BVENTILTYPE',  (FStType = 'SEC') and (GetAction=taCreat) ) ; //FStType = 'LOI' ) ;

  SetControlVisible( 'BLIGNEINSERT', True ) ;
  SetControlVisible( 'BLIGNESUPPR',  True ) ;
  // Impression
  SetControlVisible( 'BIMPRIMER',  True ) ;
  SetControlVisible( 'BCOPIER',    (FStType='LOI') and (GetAction=taCreat) ) ;

end;

procedure TOF_CPLOIVENTIL.InitEvenements;
begin
  // EVT Fiche
  FOldKeyDown                       := TFVierge(Ecran).OnKeyDown ;
  TFVierge(Ecran).OnKeyDown         := FormKeyDown ;
  TFVierge(Ecran).OnCloseQuery      := FormCloseQuery ;
  TFVierge(Ecran).OnAfterFormShow   := OnAfterFormShow ;

  // Boutons
  TToolBarButton97(GetControl('BLIGNESUPPR', True)).OnClick   := BLigneSupprClick;
  TToolBarButton97(GetControl('BLIGNEINSERT', True)).OnClick  := BLigneInsertClick;
  TToolBarButton97(GetControl('BCOPIER', True)).OnClick       := BCopierClick;

  TToolBarButton97(GetControl('BCALCULER', True)).OnClick     := BCalculerClick;
  TToolBarButton97(GetControl('BRECALCULER', True)).OnClick   := BRecalculerClick;
  TToolBarButton97(GetControl('BGENERER', True)).OnClick      := BGenererClick;
  TToolBarButton97(GetControl('BVENTILTYPE', True)).OnClick   := BVentilTypeClick;
  {$IFDEF CCSTD}
  SetControlVisible ('BZOOMODA',False);
  {$ELSE}
  TToolBarButton97(GetControl('BZOOMODA', True)).OnClick      := BZoomODAClick;
  {$ENDIF}
  TToolBarButton97(GetControl('BFERME', True)).OnClick        := BFermeClick;
  TToolBarButton97(GetControl('BIMPRIMER', True)).OnClick     := BImprimerClick;

  // Caractéristiques
  THValComboBox( GetControl('CLV_AXE', True) ).OnChange       := AxeChanged ;
  THEdit( GetControl('CLV_SECTION', True) ).OnChange          := SectionChanged ;
  THEdit( GetControl('CLV_SECTION', True) ).OnExit            := SectionExit ;
  THEdit( GetControl('LOIVENTIL', True) ).OnExit              := LoiChanged ;
  THEdit( GetControl('LOIVENTIL', True) ).OnElipsisClick      := LoiVentilClick;
  THEdit( GetControl('CLV_DATEGENER', True) ).OnExit          := DateGenerExit ;

  // Règles
  THValComboBox(GetControl('CLV_LOIMETHODE', True)).OnChange  := LoiMethodeChanged ;
  THValComboBox(GetControl('CLV_JOURNAL', True)).OnChange     := JournalChanged ;

  // Montants
  THNumEdit( GetControl('CLV_MONTANTCORR', True) ).OnChange   := MontantCorrChanged ;


  // Grille
  FListe.OnDblClick          := FListeDblClick ;
  FListe.OnCellExit          := FListeCellExit ;
  FListe.OnCellEnter         := FListeCellEnter ;
  FListe.OnKeyDown           := FListeKeyDown ;
  FListe.OnElipsisClick      := FListeElipsisClick ;

end;

procedure TOF_CPLOIVENTIL.AxeChanged(Sender: TObject);
var lStAxe : string ;
begin

  lStAxe := GetControlText('CLV_AXE') ;
  if FStAxe = lStAxe then Exit ;

  FStAxe := lStAxe ;

  // MAJ des DataType section et journal ODA...
  SetDataType( lStAxe ) ;

  if FBoOnLoad then Exit ;

  // init des zones Section d'attente et journal
  SetControlText( 'CLV_SECTIONATT', GetInfoCpta( AxeToFb( lStAxe )).Attente ) ;
  SetControlProperty( 'CLV_JOURNAL', 'ITEMINDEX', 0 ) ;
end;

procedure TOF_CPLOIVENTIL.LoiChanged(Sender: TObject);
var lTobModele : Tob ;
    lStRefV    : string ;
begin
  if Trim( GetControlText( 'LOIVENTIL' ) ) = '' then exit ;
  if FStType = 'LOI' then Exit ;

  lTobModele := Tob.Create( 'CLOIVENTIL', nil, -1 ) ;
  lTobModele.PutValue('CLV_LOITYPE', 'LOI' ) ;
  lTobModele.PutValue('CLV_NUMERO', GetControlText('LOIVENTIL') ) ;

  if lTobModele.LoadDB then
    begin
    resetGrille ;
    lStRefV := lTobModele.GetString('CLV_VENTIL') ;
    if lStRefV <> '' then
      lTobModele.LoadDetailDBFromSQL( 'VENTIL', 'SELECT * FROM VENTIL WHERE V_NATURE="CLV" AND V_COMPTE="' + lStRefV + '"' ) ;

    // ===> Caractéristiques
    // Axe deja renseigné normalement
    // Section si pas encore renseignée
    if GetControlText( 'CLV_SECTION' ) = '' then
      SetControlText( 'CLV_SECTION' , lTobModele.GetString('CLV_SECTION') ) ;
    // Libellé
    // Loi stricte
    SetControlText( 'CLV_STRICTE' , lTobModele.GetString('CLV_STRICTE') ) ;

    // ===> Règles
    SetControlText( 'CLV_LOINIVEAU' ,  lTobModele.GetString('CLV_LOINIVEAU') ) ;
    SetControlText( 'CLV_LOIMETHODE' , lTobModele.GetString('CLV_LOIMETHODE') ) ;
    SetControlText( 'CLV_SECTIONATT' , lTobModele.GetString('CLV_SECTIONATT') ) ;
    SetControlText( 'CLV_JOURNAL' ,    lTobModele.GetString('CLV_JOURNAL') ) ;

    // ===> grille de répartition
    if lTobModele.Detail.count > 0
      then begin
           AfficheGrille( lTobModele , False, True) ;
           FBoGrilleModif := True ;
           end ;

    BCalculerClick( nil ) ;
    RecalculGrille(0);

    end ;

  FreeAndNil( lTobModele ) ;
  GereAccesGrille ;
  Accesboutons ;

end;

procedure TOF_CPLOIVENTIL.SectionChanged(Sender: TObject);
var lStVal : string ;
begin
  if FBoOnLoad then Exit ;
  if FStType<>'SEC' then Exit ;
  lStVal := GetControlText('CLV_SECTION') ;
  if length(lStVal)<>GetInfoCpta( AxeTofb( FStAxe ) ).Lg then Exit ;
  FDtDateGener := iDate1900 ;
  FBoDateOk    := False ;
  BCalculerClick( nil ) ;
end;

procedure TOF_CPLOIVENTIL.SetDataType(vStAxe: string);
begin
  if length(vStAxe)<>2 then Exit ;
  SetControlProperty( 'CLV_SECTION',
                      'PLUS',
                      ' S_AXE="' + vStAxe + '" AND S_FERME="-" AND S_INDIRECTE="X" ' ) ;
  SetControlProperty( 'CLV_JOURNAL',
                      'PLUS',
                      ' J_NATUREJAL="ODA" AND J_AXE="' + vStAxe + '" AND J_FERME="-" ' ) ;
  SetControlProperty( 'CLV_SECTIONATT',
                      'PLUS',
                      ' S_AXE="' + vStAxe + '" AND S_FERME="-" ' ) ;

end;

procedure TOF_CPLOIVENTIL.InitGrille;
var lStMeth : string ;
begin
  FListe.RowCount      := 50 ;
  FListe.ElipsisButton := True ;
  FListe.FlipBool      := True ;
  FListe.CalcInCell    := True ;

  FListe.ColCount      := 6 ;
  FListe.DbIndicator   := True ;
  FListe.FixedCols     := 2 ;
  FListe.RowHeights[0] := 18 ; // hauteur de la ligne de titre

  // Colonne N°
  FListe.Cells    [ IDX_NO, 0 ] := TraduireMemoire('N°') ;
  FListe.ColWidths[ IDX_NO ]    := 10 ;
  FListe.ColAligns[ IDX_NO ]    := taCenter ;

  // Colonne Section
  FListe.Cells      [ IDX_SECTION, 0 ] := TraduireMemoire('Section') ;
  FListe.ColWidths  [ IDX_SECTION ]    := 50 ;
  FListe.ColLengths [ IDX_SECTION ]    := GetInfoCpta( AxeToFb( GetControlText('CLV_AXE') ) ).Lg ;

  // Colonne libellé de la section
  FListe.Cells        [ IDX_LIBELLE, 0 ] := TraduireMemoire('Libellé') ;
  FListe.ColWidths    [ IDX_LIBELLE ]    := 80 ;
  FListe.ColEditables [ IDX_LIBELLE ]    := False ;

  // Les autres colonnes sont paramétrés en fonction de la règle utilisée
  lStMeth := GetControlText('CLV_LOIMETHODE');
  if lStMeth = '' then
    lStMeth := '1MO' ;
  SetFormatGrille( lStMeth ) ;

end;

procedure TOF_CPLOIVENTIL.SetFormatGrille(vStMeth: string);
var // Var colonne "Clef"
    lStTitreC      : string ;
    lBoColVisuC    : boolean ;
    lBoColEnaC     : boolean ;
    lStColTypeC    : Char ;
    lStColFormatC  : string ;
    // Var colonne "Montant"
    lStTitreM      : string ;
    lBoColVisuM    : boolean ;
    lBoColEnaM     : boolean ;
    lStColTypeM    : Char ;
    lStColFormatM  : string ;
    HMT            : TComponent ;
    lInDec         : integer ;
begin

  // colonne "Clef"
  lStTitreC      := TraduireMemoire('Pourcentage') ;
  lBoColVisuC    := True ; //FStType <> 'LOI' ;
  lBoColEnaC     := True ;
  lStColTypeC    := 'R' ;
  lStColFormatC  := '#,##0';
  lInDec         := GetDecimale( IDX_CLEF ) ;
  if lInDec > 0
    then lStColFormatC := lStColFormatC + '.' + Copy('000000000000000000000', 1, lInDec) ;

  // colonne "Montant"
  lStTitreM      := TraduireMemoire('Montant') ;
  lBoColVisuM    := FStType <> 'LOI' ;
  lBoColEnaM     := True ;
  lStColTypeM    := 'R' ;
  lStColFormatM  := '#,##0';
  lInDec         := GetDecimale( IDX_MONTANT ) ;
  if lInDec > 0
    then lStColFormatM := lStColFormatM + '.' + Copy('000000000000000000000', 1, lInDec) ;

  // sur les montants
  if vStMeth = '1MO' then
    begin
    lBoColEnaC := False ;
    lBoColVisuC := FStType = 'SEC' ;
    end
  // sur les %
  else if vStMeth = '2PO' then
    begin
    lBoColEnaM  := False ;
    end
  // sur les UO
  else if vStMeth = '3UO' then
    begin
    lStTitreC   := TraduireMemoire('Unité d''oeuvre') ;
    lBoColEnaM  := False ;
    end
  // sur les montants des sections
  else if vStMeth = '4SE' then
    begin
    lStTitreC   := TraduireMemoire('Total section') ; //au ' + GetControlText('CLV_DATEGENER') ) ;
    lBoColEnaC  := False ;
    lBoColVisuC := FStType = 'SEC' ;
    lBoColEnaM  := False ;
    end
  // sur les UO des sections
  else if vStMeth = '5CR' then
    begin
    // colonnes clefs
    lStTitreC   := TraduireMemoire('Unité d''oeuvre') ;
    lBoColEnaC  := False ;
    lBoColVisuC := FStType = 'SEC' ;
    // colonnes montants
    lBoColEnaM := False ;
    end ;

  // colonne "Clef"
  FListe.Cells         [ IDX_CLEF, 0 ] := lStTitreC ;
  FListe.ColEditables  [ IDX_CLEF ]    := lBoColVisuC and lBoColEnaC ;
  if lBoColVisuC then
    begin
    FListe.ColWidths   [ IDX_CLEF ]    := 50 ;
    FListe.Collengths  [ IDX_CLEF ]    := 12 ;
    FListe.ColFormats  [ IDX_CLEF ]    := lStColFormatC ;
    FListe.ColTypes    [ IDX_CLEF ]    := lStColTypeC ;
    FListe.ColAligns   [ IDX_CLEF ]    := taRightJustify ;
    end
  else
    begin
    FListe.Collengths  [ IDX_CLEF ]    := -1 ;
    FListe.ColWidths   [ IDX_CLEF ]    := -1 ;
    end ;

  // colonnes "Montant"
  FListe.Cells         [ IDX_MONTANT, 0 ] := lStTitreM ;
  FListe.ColEditables  [ IDX_MONTANT ]    := lBoColVisuM and lBoColEnaM ;
  if lBoColVisuM then
    begin
    FListe.ColWidths   [ IDX_MONTANT ]    := 50 ;
    FListe.Collengths  [ IDX_MONTANT ]    := 12 ;
    FListe.ColTypes    [ IDX_MONTANT ]    := lStColTypeM ;
    FListe.ColFormats  [ IDX_MONTANT ]    := lStColFormatM ;
    FListe.ColAligns   [ IDX_MONTANT ]    := taRightJustify ;
    end
  else
    begin
    FListe.Collengths  [ IDX_MONTANT ]    := -1 ;
    FListe.ColWidths   [ IDX_MONTANT ]    := -1 ;
    end ;

  HMT := TForm( Ecran ).FindComponent('HMTrad');
  if (HMT<>nil) then
    THSystemMenu(HMT).ResizeGridColumns( FListe );

  FListe.Invalidate ;

  LayoutCumuls ;

end;


procedure TOF_CPLOIVENTIL.InitVariables;
begin
  FInfo := TInfoEcriture.Create ;

  FTobLoi    := TOB.Create('CLOIVENTIL', nil, -1 );
  FTobLoi.InitValeurs ;

  FStType    := '' ;
  FStAxe     := '' ;
  FStSection := '' ;
  LastError  := 0 ;
  FTobODA    := nil ;

  FBoOnLoad  := False ;

end;

function TOF_CPLOIVENTIL.LoadRecord : boolean ;
var lStRefV : string ;
    lBoODA  : boolean ;
begin

  FBoOnLoad  := True ;
  result     := False ;
  FTobODA    := nil ;

  FDtDateGener := iDate1900 ;
  FBoDateOk    := False ;

  if FInNumero > 0 then
    begin
    FTobLoi.InitValeurs;
    FTobLoi.PutValue('CLV_LOITYPE',   FStType ) ;
    FTobLoi.PutValue('CLV_NUMERO', FInNumero ) ;
    result := FTobLoi.LoadDB ;

    if result then
      begin
      // MAJ des DataType section et journal ODA...
      FStAxe := FTobLoi.GetString('CLV_AXE') ;
      SetDataType( FStAxe ) ;

      // Update variables
      SetMontantCalc( FTobLoi.GetDouble('CLV_MONTANTCALC') ) ;
//      FDtDateGener := FTobLoi.GetDateTime('CLV_DATEGENER');

      // MAJ infos
      FSoldeTraite := 0 ;
      ChargeODA ;
      if FTobODA.Detail.count > 0 then
        FSoldeTraite := FTobLoi.GetDouble('CLV_MONTANTTOTAL') ;
      SetInfosODA ;

      // chargement de la grille de répartition
      lStRefV := FTobLoi.GetString('CLV_VENTIL') ;
      if lStRefV <> '' then
        FTobLoi.LoadDetailDBFromSQL( 'VENTIL', 'SELECT * FROM VENTIL WHERE V_NATURE="CLV" AND V_COMPTE="' + lStRefV + '"' ) ;

      end ; // if result

    end ;  // if FInNumero

  AfficheTob ;
  FBoGrilleModif := False ;
  FTobLoi.SetAllModifie(False);
  GereAccesGrille ;
  Accesboutons     ;
  FBoOnLoad  := False ;

  // Validation de la date
  TestPresenceDate( FTobLoi.GetDateTime('CLV_DATEGENER'), False, False ) ;
  lBoODA := Assigned( FTobODA ) and ( FTobODA.Detail.count > 0 ) ;

  if PretACalculer and (not lBoODA) and FBoDateOk then
    CalculMontants ;

end;

procedure TOF_CPLOIVENTIL.NewRecord;
begin

  FTobLoi.InitValeurs ;
  FTobLoi.PutValue('CLV_LOITYPE', FStType ) ;

  // ============= CARACTERISTIQUES =============
  if FStAxe = '' then
    FStAxe := 'A1' ;
  FTobLoi.PutValue('CLV_AXE', FStAxe ) ;
  // MAJ des DataType section et journal ODA...
  SetDataType( FStAxe ) ;

  if FStSection<>'' then
    FTobLoi.PutValue('CLV_SECTION', FStSection ) ;

  FTobLoi.PutValue('CLV_STRICTE',   '-' ) ;

  if FStType = 'LOI'
    then FTobLoi.PutValue('CLV_DATEGENER', iDate1900 )
    else FTobLoi.PutValue('CLV_DATEGENER', GetInitDate ) ;

  // ============= REGLES =============
  FTobLoi.PutValue('CLV_LOINIVEAU',   '1SE' ) ;
  FTobLoi.PutValue('CLV_LOIMETHODE',  '1MO' ) ;
  FTobLoi.PutValue('CLV_SECTIONATT',  GetInfoCpta( AxeToFb( FStAxe )).Attente ) ;
  if THValComboBox( GetControl( 'CLV_JOURNAL')).Values.Count > 0 then
    FTobLoi.PutValue('CLV_JOURNAL',     THValComboBox( GetControl( 'CLV_JOURNAL')).Values[0] ) ;

  // ============= TRAITEMENT =========
  FTobLoi.PutValue('CLV_VENTIL',      '' ) ;
  FTobLoi.PutValue('CLV_NUMEROPIECE', 0  ) ;

  FDtDateGener := iDate1900 ;
  FBoDateOk    := False ;

  FSoldeTraite := 0 ;

  AfficheTob ;

  FBoGrilleModif := False ;
  FTobLoi.SetAllModifie(False);
  GereAccesGrille ;
  Accesboutons    ;

  if PretACalculer then
    CalculMontants ;

end;

function TOF_CPLOIVENTIL._SaveRecord : boolean ;
var lQNum          : TQuery ;
    lInNum         : integer ;
    lStRefV        : string ;
begin

  result := False ;

  try

    beginTrans ;

    FTobLoi.SetAllModifie(True);

    // Affectation numero interne
    lInNum := FTobLoi.GetInteger('CLV_NUMERO') ;
    if lInNum = 0  then
      begin
      lQNum := OpenSQL('SELECT MAX(CLV_NUMERO) AS MAXNUM FROM CLOIVENTIL WHERE CLV_LOITYPE="' + FStType + '"', True) ;
      if lQNum.Eof
        then lInNum := 1
        else lInNum := lQNum.FindField('MAXNUM').AsInteger + 1 ;
      Ferme( lQNum ) ;
      FTobLoi.PutValue('CLV_NUMERO', lInNum ) ;
      UpdateNumero( lInNum ) ;
      end
    else
      begin
      // Suppression des anciennes ventil
      lStRefV := GetRefVentil( lInNum ) ;
      ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE="CLV" AND V_COMPTE="' + lStRefV + '" ') ;
      ExecuteSQL('DELETE FROM CLOIVENTIL WHERE CLV_LOITYPE="' + FStType + '" AND CLV_NUMERO=' + FTobLoi.GetString('CLV_NUMERO') ) ;
      UpdateNumero( lInNum ) ;
      end ;

    // Enregistrement
    FTobLoi.InsertDBByNivel(False) ;
    FTobLoi.SetAllModifie(False);

    commitTrans ;
    result := True ;

    Except

      on E:Exception do
        begin
        Rollback ;
        PgiError( TraduireMemoire( 'Une erreur est survenue lors de l''enregistrement :')
                  +#10+#13 + E.Message , Ecran.Caption ) ;
        end ;

    end ;

end;

function TOF_CPLOIVENTIL.TestRecord: boolean;
var lInErr : integer ;

  function _TestChampsOblig : integer ;
  var i : integer ;
    begin
    result := -1 ;
    for i:= 1 to _NbChamps do
      if GetControlText( _TabMsgChamps[i] ) = '' then
        begin
        result := i ;
        break ;
        end ;
    end ;

begin

  result := False ;

  // === CHAMP OBLIGATOIRE ===
  lInErr := _TestChampsOblig ;
  if lInErr > 0 then
    begin
    PgiInfo( TraduireMemoire( _TabMsgErreur[ lInErr ] ), Ecran.Caption ) ;
    SetFocusControl( _TabMsgChamps[ lInErr ] ) ;
    Exit ;
    end ;

  // test de la date saisie en mode répartition
  if (FStType='SEC') then
    if (FDtDateGener = iDate1900) then
      begin
      PgiInfo( TraduireMemoire('Vous devez saisir une date de génération.') ) ;
      SetFocusControl('CLV_DATEGENER');
      Exit ;
      end
    else
      if not FBoDateOk then
        begin
        PgiInfo( TraduireMemoire('La date de génération n''est pas correcte.') ) ;
        SetFocusControl('CLV_DATEGENER');
        Exit ;
        end ;

  if not EstGrilleValide then
    Exit ;

  result := True ;

end;

procedure TOF_CPLOIVENTIL.UpdateRecord;
var lTobV      : Tob ;
    i          : integer ;
    iLig       : integer ;
    lStRefV    : string ;
    lStSection : string ;
begin

  // Récup info TOB principal
  GetEcran ;

  // Constitution de la ventilation type
  if FBoGrilleModif then
    begin
    FTobLoi.ClearDetail ;
    iLig   := 1 ;
    if FTobLoi.GetString('CLV_VENTIL') <> ''
      then lStRefV := FTobLoi.GetString('CLV_VENTIL') ;
    for i := 1 to FListe.RowCount - 1 do
      begin
      lStSection := FListe.Cells[ IDX_SECTION,i] ;
      // si la section est renseignée et valide...
      if ( Trim(lStSection ) <> '' ) and FInfo.LoadSection( lStSection, GetControlText('CLV_AXE') ) then
        begin
        lTobV := Tob.Create('VENTIL', FTobLoi, -1 ) ;

        lTobV.PutValue('V_NATURE',       'CLV' );
        lTobV.PutValue('V_COMPTE',       lStRefV );
        lTobV.PutValue('V_NUMEROVENTIL', iLig );

        lTobV.PutValue('V_SECTION',      FInfo.GetString('S_SECTION') );
        if FListe.ColLengths[ IDX_CLEF ] > 0 then
          lTobV.PutValue('V_TAUXMONTANT', GetMontant( IDX_CLEF, i ) );
        if FListe.ColLengths[ IDX_MONTANT ] > 0 then
          lTobV.PutValue('V_MONTANT',     GetMontant( IDX_MONTANT, i ) );

        Inc(iLig) ;
        end ;
      end ;

    if FTobLoi.Detail.Count > 0
      then FTobLoi.PutValue('CLV_VENTIL', lStRefV )
      else FTobLoi.PutValue('CLV_VENTIL', '' ) ;

    end ;

end;

procedure TOF_CPLOIVENTIL.AfficheTob;
begin
  // Entete
  PutEcran ;

  // Affichage de grille de répartition depuis une ventil type
  if FTobLoi.GetString('CLV_VENTIL')<>'' then
    AfficheGrille( FTobLoi )
  else
    ResetGrille ; // Rien dans la grille

end;


procedure TOF_CPLOIVENTIL.ResetGrille;
begin
  FListe.RowCount := 2 ;
  InitLigne( 1 ) ;
  FListe.RowCount := 50 ;
  NumeroteLignes ;
end;

procedure TOF_CPLOIVENTIL.UpdateNumero( vInNum : integer );
var i      : integer ;
    lStRef : string ;
begin

  FTobLoi.PutValue('CLV_NUMERO', vInNum ) ;

  lStRef := GetRefVentil( vInNum ) ;

  if FTobLoi.Detail.Count > 0 then
    begin
    FTobLoi.PutValue('CLV_VENTIL', lStRef ) ;
    for i:=0 to (FTobLoi.Detail.Count-1) do
      FTobLoi.Detail[i].PutValue('V_COMPTE', lStRef ) ;
    end ;

end;

function TOF_CPLOIVENTIL.GetRefVentil( vInNum: integer ; vStType : string ): string;
begin
  if vStType='' then
    vStType:=FStType ;
  if vInNum > 0
    then result := vStType + Copy( '000000000000', 1, 10-length(IntToStr( vInNum ))) + IntToStr( vInNum )
    else result := vStType + '0000000000' ;
end;

procedure TOF_CPLOIVENTIL.OnAfterFormShow;
begin

  if FInNumero > 0
    then LoadRecord
    else if GetAction<>taConsult
            then NewRecord
            else begin
                 PgiError('Erreur au chargement de l''enregistrement', Ecran.Caption ) ;
                 end ;

  GereAccesGrille ;

end;

procedure TOF_CPLOIVENTIL.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var rep : integer ;
begin
  if GetAction = taConsult then Exit ;

  CanClose := LastError = 0 ;
  if CanClose then
    if EstModifEnCours then
      begin
      Rep      := PGIAskCancel('Voulez-vous enregistrer les modifications ?',Ecran.Caption) ;
      Case Rep of
        mrYes : CanClose := SaveRecord ;
        mrNo  : CanClose := True ;
        else    CanClose := False ;
        end ;
      end ;
      
end;

procedure TOF_CPLOIVENTIL.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  Case Key of
    // Validation ligne
    VK_F5    : if Shift = [] then
                  begin
                  if Ecran.ActiveControl.Name = 'CLV_SECTION' then
                    begin
                    Key := 0 ;
                    THEdit( GetControl('CLV_SECTION')).ElipsisClick(nil) ;
                    end ;
                  end ;

    // Validation ligne
    VK_F10    : if Shift = [] then
                  begin
                  Key := 0 ;
                  ClickValide;
                  end ;

    // Recalcul
    VK_F9    : if Shift = [] then
                  begin
                  Key := 0 ;
                  BCalculerClick(nil);
                  end ;

    // nouvelle ligne
    VK_ESCAPE : if Shift = [] then
                  begin
                  Key := 0 ;
                  BFermeClick(nil) ;
                  end ;

    // raccourci Boutons
       // Loi de répartition
      {^L} 76 : if Shift = [ssCtrl] then
                  begin
                  Key := 0 ;
                  BVentilTypeClick(nil) ;
                  end ;


    // Impression
(*      {^I} 73 : if Shift = [Ctrl] then
                  begin
                  Key := 0 ;
                  Imprimer(nil) ;
                  end ;
*)
  end ;

  if Assigned( FOldKeyDown ) then
    FOldKeyDown( Sender, Key, Shift ) ;

end;

function TOF_CPLOIVENTIL.ClickValide: boolean;
begin
  result := True ;
  if GetAction = taConsult then Exit ;

  // MAJ de l'objet + Modif ?
  if not EstModifEnCours then
    Exit ;

  if Assigned(FTobODA) and (FTobODA.Detail.count > 0 ) then
    begin
    if FBoGrilleModif or
       (FSoldeTraite <> FTobLoi.GetDouble('CLV_MONTANTTOTAL') ) or
       (FDtDateGener <> FTobODA.Detail[0].GetDateTime('Y_DATECOMPTABLE') ) then
      begin
      PgiInfo( TraduireMemoire('Des pièces d''Od Analytiques ont déjà été générées pour cette répartition, vous ne pouvez plus modifier la répartition sans régénérer les pièces.') ) ;
      result := False ;
      Exit ;
      end ;
    end ;

  // Validation
  result := SaveRecord ;

  Accesboutons     ;

end;

function TOF_CPLOIVENTIL.GetAction: TActionFiche;
begin
  result := TFVierge( Ecran ).TypeAction ;
end;

procedure TOF_CPLOIVENTIL.BLIGNEINSERTClick(Sender: TObject);
begin
  InsertRow( FListe.Row ) ;
end;

procedure TOF_CPLOIVENTIL.BLIGNESUPPRClick(Sender: TObject);
begin
  DeleteRow( FListe.row ) ;
end;

procedure TOF_CPLOIVENTIL.BCopierClick(Sender: TObject);
var lStNum : string ;
begin

  if GetAction <> taCreat then Exit ;

  // choix du masque de référence
  lStNum := Choisir( 'Sélection d''une loi de répartiton à dupliquer','CLOIVENTIL','CLV_LIBELLE','CLV_NUMERO','CLV_LOITYPE="LOI" AND CLV_AXE="' + FStAxe + '"','', False, False, 0) ;

  // Chargement du masque de référence
  if (lStNum<>'') and (ValeurI( lStNum )>0) then
    begin
    FTobLoi.ClearDetail ;
    FInNumero := ValeurI( lStNum ) ;
    LoadRecord ;

    FTobLoi.PutValue('CLV_LOITYPE', FStType ) ;
    FTobLoi.PutValue('CLV_AXE',     FStAxe ) ;
    FTobLoi.PutValue('CLV_NUMERO',  0 ) ;
    FInNumero := 0 ;

    // Affichage
    AfficheTob ;

    FTobLoi.SetAllModifie(True);
    FBoGrilleModif := True ;

    AccesBoutons ;

    end ;

end;

procedure TOF_CPLOIVENTIL.BCalculerClick(Sender: TObject);
begin
  if FBoOnLoad then Exit ;
  CalculMontants  ;
  Accesboutons    ;
end;

procedure TOF_CPLOIVENTIL.BGenererClick(Sender: TObject);
var lTobListePieces : Tob ;
    lTobPieceODA    : Tob ;
    lStGene         : string ;
    lDtDateModif    : TDateTime ;
    i,j             : integer ;
    lInNumeroMin    : integer ;
    lInNumeroMax    : integer ;
    lDtDateGener    : TDateTime ;
    lStExo          : string ;
    lMontant        : Double ;
begin


  if not PretAGenerer(True) then Exit ;
  if not EstGrilleValide then Exit ;

  // Test section d'attente sur axe sélectionné
  if GetInfoCpta( AxeToFb( FStAxe )).Attente = '' then
    begin
    PgiInfo( TraduireMemoire('La section d''attente n''est pas paramétrée sur l''axe sélectionné.'), Ecran.Caption ) ;
    Exit ;
    end ;

  lDtDateGener := FTobLoi.GetDateTime('CLV_DATEGENER') ;
  lStExo       := ctxExercice.QuelExoDt( lDtDateGener ) ;
  lDtDateModif := NowH ;

  // MAJ devise
  FDevise.Code    := V_PGI.DevisePivot ;
  GetInfosDevise( FDevise ) ;
  FDevise.Taux    := GetTaux( FDevise.Code, FDevise.DateTaux, lDtDateGener ) ;

  // répartition du montant de correction sur les soldes par comptes généraux
  SetCorrectionDansCumuls ;

  // === Generation des pièces ===
  lTobListePieces := Tob.Create('LISTE_PIECES', nil, -1 ) ;
  For i:=0 to FTobCumuls.Detail.count - 1 do
    begin
    lStGene       := FTobCumuls.Detail[i].GetString('Y_GENERAL') ;
//    lMontant      := FTobCumuls.Detail[i].GetDouble('DEB') - FTobCumuls.Detail[i].GetDouble('CRE') ;
    lMontant      := FTobCumuls.Detail[i].GetDouble('MONTANT') ;

    // Si solde du compte en cours à 0, il n'est pas pris en compte dans l'algo
    if lMontant = 0 then Continue ;

    lTobPieceODA  := GenerePieceODA( lStGene , lMontant, lTobListePieces ) ;

    if lTobPieceODA.Detail.count < 2 then
      begin
      PGIInfo( TraduireMemoire('Génération interrompue : La pièce d''ODA est incomplète !'), Ecran.Caption ) ;
      lTobListePieces.ClearDetail ;
      FreeAndNil( lTobListePieces ) ;
      Exit ;
      end ;

    end ;

  // Au moins 1 pièce générée ?
  if lTobListePieces.Detail.count = 0 then
    begin
    PGIInfo( TraduireMemoire('Génération interrompue : Aucune pièce d''ODA n''a pu être générée !'), Ecran.Caption ) ;
    lTobListePieces.ClearDetail ;
    FreeAndNil( lTobListePieces ) ;
    Exit ;
    end ;


    // === ENREGISTREMENT ===
    try

      beginTrans ;

      // MAJ des Numero de pièce
      lInNumeroMin := 0 ;
      lInNumeroMax := 0 ;
      for i:= 0 to lTobListePieces.Detail.count - 1 do
        begin
        lTobPieceODA := lTobListePieces.Detail[i] ;
        lInNumeroMax := GetNewNumJal( FTobLoi.GetString('CLV_JOURNAL'), True, lDtDateGener ) ;
        if lInNumeroMax <= 0 then
          raise Exception.Create( TraduireMemoire('Erreur à la détermination du numéro de la pièce d''ODA') ) ;
        if lInNumeroMin = 0 then
          lInNumeroMin := lInNumeroMax ;
        for j:= 0 to lTobPieceODA.Detail.count - 1 do
          lTobPieceODA.Detail[j].PutValue('Y_NUMEROPIECE', lInNumeroMax ) ;
        end ;

      // Enregistrement des pièces
      lTobListePieces.InsertDBByNivel(False);

      // MAJ DES SOLDES DES COMPTES
      for i:= 0 to lTobListePieces.Detail.count - 1 do
        begin
        lTobPieceODA := lTobListePieces.Detail[i] ;
        for j:= 0 to lTobPieceODA.Detail.count - 1 do
          begin
          //Maj Solde
          MajSoldeSectionTOB(lTobPieceODA.Detail[j], True);
          //Maj journaux
          MajJournalAnaTob(lTobPieceODA.Detail[j], False, True);
          end ;
        end ;

      // MAJ DE LA LOI
      FTobLoi.PutValue('CLV_DATETRAITEMENT', lDtDateModif ) ;
      FTobLoi.PutValue('CLV_NUMEROPIECE',    lInNumeroMax ) ;
      ExecuteSQL('UPDATE CLOIVENTIL SET CLV_DATETRAITEMENT="' + UsDateTime(lDtDateModif) + '" , CLV_NUMEROPIECE=' + IntToStr(lInNumeroMax)
                         + ' WHERE CLV_LOITYPE="SEC" AND CLV_NUMERO=' + FTobLoi.GetString('CLV_NUMERO') ) ;


      CommitTrans ;

      if lTobListePieces.Detail.count=1
        then PGIInfo( TraduireMemoire('L''OD Analytique n°') + FTobLoi.GetString('CLV_NUMEROPIECE') + TraduireMemoire(' a été générée sur le journal ') + RechDom( 'TTJOURNAUX', FTobLoi.GetString('CLV_JOURNAL'), False ), Ecran.Caption )
        else PGIInfo( TraduireMemoire('Les OD Analytiques n°') + IntToStr(lInNumeroMin) + ' à ' + IntToStr(lInNumeroMax) + TraduireMemoire(' ont été générées sur le journal ') + RechDom( 'TTJOURNAUX', FTobLoi.GetString('CLV_JOURNAL'), False ), Ecran.Caption ) ;

      // gestion état des tobs / fiche
      FTobLoi.SetAllModifie( False );
      FBoGrilleModif := False ;
      if GetAction=taCreat then
        SetAction( taModif ) ;

//      SetControlEnabled('BVALIDER',    False ) ;
      SetControlEnabled('BGENERER',    False ) ;
      SetControlVisible('BRECALCULER', True ) ;
      FSoldeTraite  := FTobLoi.GetDouble('CLV_MONTANTTOTAL') ;

      ChargeODA( lTobListePieces ) ;

      // MAJ infos pieds
      SetInfosODA ;
      AccesBoutons ;

      lTobListePieces.ClearDetail ;
      FreeAndNil( lTobListePieces ) ;

      Except

        on E:Exception do
          begin
          Rollback ;
          PgiError( TraduireMemoire( 'Une erreur est survenue lors de la génération de la pièce d''ODA !')
                    +#10+#13 + E.Message , Ecran.Caption ) ;
          end ;

    end ; // ===> Try

end;

procedure TOF_CPLOIVENTIL.BRecalculerClick(Sender: TObject);
var lBoResult    : Boolean ;
    lTobPieceODA : Tob ;
    i,j          : integer ;
begin
  // suppression ancienne ODA
  if Assigned(FTobODA) and ( FTobODA.Detail.count > 0 ) then
    begin
    if PgiAskCancel( TraduireMemoire('Le traitement a déjà été effectué, les pièces d''ODA existantes vont être supprimées pour pouvoir regénérer les calculs. Voulez vous continuer ?.')  ) <> mrYes
      then Exit ;
    // suppression des ODA ;
    try
      beginTrans ;

      lTobPieceODA := Tob.Create('OLD_ODA', nil, -1) ;
      for i:=FTobOda.Detail.count-1 downto 0 do
        begin
        // Chargement de la pièce entière
        lTobPieceODA.ClearDetail ;
        lTobPieceODA.LoadDetailFromSQL('SELECT ' + GetSelectAll('Y', True) + ' FROM ANALYTIQ  WHERE ' + WhereEcritureTOB( tsODA, FTobODA.Detail[i], False, False, 'Y') ) ;
        // Maj des solde
        for j:= 0 to lTobPieceODA.Detail.count - 1 do
          begin
          //Maj Solde
          MajSoldeSectionTOB(lTobPieceODA.Detail[j], False);
          //Maj journaux
          MajJournalAnaTob(lTobPieceODA.Detail[j], False, False);
          end ;
        // Delete
        ExecuteSQL( 'DELETE FROM ANALYTIQ WHERE ' + WhereEcritureTOB( tsODA, FTobODA.Detail[i], False, False, 'Y') ) ;
        end ;

        FTobOda.ClearDetail ;
        lTobPieceODA.ClearDetail ;
        FreeAndNil( lTobPieceODA ) ;

        CommitTrans ;

        FSoldeTraite := 0 ;
        lBoResult    := True ;

        Except
          on E:Exception do
            begin
            lBoResult := False ;
            Rollback ;
            PgiError( TraduireMemoire( 'Une erreur est survenue lors de la génération de la pièce d''ODA !')
                      +#10+#13 + E.Message , Ecran.Caption ) ;
            end ;
      end ; // ===> Try

    SetInfosODA ;
    if not lBoResult then Exit ;
    AccesBoutons ;

    end ;

  BCalculerClick(nil);

end;

function TOF_CPLOIVENTIL.EstModifEnCours: boolean;
begin

  // recup dernière modif
//  FBoOnLoad  := True ;
  NextPrevControl( Ecran ) ;

  // MAJ de l'objet
  UpdateRecord ;
//  FBoOnLoad  := False ;

  // test
  result := FBoGrilleModif or FTobLoi.IsOneModifie(False);

end;

procedure TOF_CPLOIVENTIL.AfficheGrille( vTobG : Tob ; vBoFromVT : boolean ; vBoFromLoi : boolean ) ;
var lTobVentil : tob ;
    i          : integer ;
    lStRegle   : string ;
    lStSection : string ;
begin

  try

    FListe.BeginUpdate ;

    if vTobG.Detail.count = 0 then
      begin
      ResetGrille ;
      Exit ;
      end ;

    if vTobG.Detail.count > 50
      then FListe.RowCount := vTobG.Detail.count + 2
      else FListe.RowCount := 50 ;
    lStRegle        := GetControlText('CLV_LOIMETHODE') ;

    for i := 0 to vTobG.Detail.count-1 do
      begin

      lTobVentil := vTobG.Detail[i] ;
      lStSection := lTobVentil.GetString('V_SECTION') ;

      if not FInfo.LoadSection( lStSection, GetControlText('CLV_AXE') ) then
        Continue ;

      // Section
      FListe.Cells[ IDX_SECTION, i+1 ] := lStSection ;

      // Libelle
      FListe.Cells[ IDX_LIBELLE, i+1 ] := FInfo.GetString('S_LIBELLE') ;

      // 3ème / 4ème colonne suivant méthode
      setMontant( IDX_CLEF,    i+1 , 0 ) ;
      setMontant( IDX_MONTANT, i+1 , 0 ) ;
      if vBoFromVT then
        begin
        if lStRegle = '2PO' then
          setMontant( IDX_CLEF,  i+1 , lTobVentil.GetDouble('V_TAUXMONTANT') )
        else if lStRegle = '3UO' then
          setMontant( IDX_CLEF,  i+1 , lTobVentil.GetDouble('V_MONTANT') ) ;
        end
      else if vBoFromLoi or (GetAction<>taCreat) then
        begin
        setMontant( IDX_CLEF,    i+1 , lTobVentil.GetDouble('V_TAUXMONTANT') ) ;
        setMontant( IDX_MONTANT, i+1 , lTobVentil.GetDouble('V_MONTANT') ) ;
        end
      // sur critères UO des sections...
      else if lStRegle = '5CR' then
        begin
        // A recalculer a chaque fois...
        if IsNumeric( FInfo.GetString('S_UO') )
          then  setMontant( IDX_CLEF,    i+1 , Valeur( FInfo.GetString('S_UO') ) )
          else  setMontant( IDX_CLEF,    i+1 , 0 ) ;
        end ;

      end ;

    NumeroteLignes ;
    RecalculGrille( 0 ) ;

  finally

    FListe.EndUpdate ;

  end ;

end;

procedure TOF_CPLOIVENTIL.FListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var lInSens : integer ;
begin
  FListe.ElipsisButton := (FListe.Row >= 1) and ( FListe.Col = IDX_SECTION ) ;
  if not EstGrilleEnabled then Exit ;
  if not FListe.ColEditables[ FListe.Col ] then
    begin
    lInSens := CGetGridSens( FListe, ACol, ARow ) ;
    ACol := FListe.Col ;
    ARow := FListe.Row ;
    GetColSuivante( ACol, ARow, lInSens ) ;
    Cancel := True ;
    end ;
end;

procedure TOF_CPLOIVENTIL.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var lStSection : string ;
    lMontant   : double ;
begin
  if ARow = 0 then Exit ;

  // Gestion de la section
  if ACol = IDX_SECTION then
    begin
    lStSection := FListe.Cells[ACol,ARow] ;
    if ( lStSection <> '' ) then
      if FInfo.LoadSection( lStSection, GetControlText('CLV_AXE') ) then
        begin
        FListe.Cells[IDX_SECTION,ARow] := FInfo.GetString('S_SECTION') ;
        FListe.Cells[IDX_LIBELLE,ARow] := FInfo.GetString('S_LIBELLE') ;
        if GetControlText('CLV_LOIMETHODE') = '5CR' then
          begin
          if IsNumeric( FInfo.GetString('S_UO') )
            then begin
                 SetMontant( IDX_CLEF, ARow, Valeur( FInfo.GetString('S_UO') ) ) ;
                 RecalculGrille( ARow ) ;
                 end
            else SetMontant( IDX_CLEF, ARow, 0 ) ;
          end
        else if GetControlText('CLV_LOIMETHODE') = '4SE' then
          begin
          SetMontant( IDX_CLEF, ARow, GetSoldeSection(lStSection) ) ;
          RecalculGrille( ARow ) ;
          end
        end
      else
        begin
        InitLigne( ARow ) ;
        end ;
    end

  // Gestion des % / UO
  else if (ACol = IDX_CLEF) and (FListe.ColEditables[IDX_CLEF]) then
    begin
    lMontant := GetMontant( IDX_CLEF, ARow ) ;
    SetMontant( IDX_CLEF, ARow, lMontant ) ;
    RecalculGrille( ARow ) ;
    end

  // Gestion des montants
  else if (ACol = IDX_MONTANT) and (FListe.ColEditables[IDX_MONTANT]) then
    begin
    lMontant := GetMontant( IDX_MONTANT, ARow ) ;
    SetMontant( IDX_MONTANT, ARow, lMontant ) ;
    RecalculGrille( ARow ) ;
    end ;

end;

procedure TOF_CPLOIVENTIL.FListeDblClick(Sender: TObject);
var lAction : TActionFiche ;
    lStCpt  : string ;
    lStAxe  : string ;
begin
  lStCpt := FListe.cells[ IDX_SECTION, FListe.Row ] ;
  lStAxe := GetControlText('CLV_AXE') ;
  if (lStCpt = '') or (lStAxe='') or not FInfo.LoadSection( lStCpt, lStAxe ) then
    Exit ;

  if ( GetAction <> taConsult ) and ExJaiLeDroitConcept( TConcept(ccSecModif), False )
    then lAction := taModif
    else lAction := taConsult ;

  FicheSection(nil, lStAxe, lStCpt, lAction, 0) ;

  if lAction=taModif then
    FInfo.LoadSection( lStCpt, lStAxe, True ) ;
end;

procedure TOF_CPLOIVENTIL.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  if (Key in [ 65..90,96..105 ]) and not ((Shift = [ssAlt]) or (Shift=[ssCtrl])) then
    FBoGrilleModif := True ;

    case Key of
      VK_F5 : if (Shift=[]) then
                begin
                Key := 0;
                FListeElipsisClick(nil);
                end
              else if (Shift=[ssCtrl]) then
                begin
                Key := 0;
                FListeDblClick(nil);
                end ;
      VK_F6 : if (Shift=[]) then
                begin
                Key := 0;
                SoldeVentil( FListe.Row ) ;
                end;
      VK_INSERT : if (Shift=[]) then
                    begin
                    Key:=0;
                    InsertRow( FListe.Row ) ;
                    end ;
      VK_DELETE : if Shift=[ssCtrl] then
                    begin
                    Key:=0 ;
                    DeleteRow( FListe.Row ) ;
                    end ;
      VK_RETURN  : if (Shift=[]) then
                      begin
                      Key := VK_TAB  ;
                      PostMessage( FListe.Handle, WM_KEYDOWN, VK_TAB, 0) ;
                      end ;
      VK_ESCAPE : if (Shift=[]) then
                    begin
                    Key:=0 ;
                    EcranClose ;
                    end ;

      end;

end;

procedure TOF_CPLOIVENTIL.FListeElipsisClick(Sender: TObject);
var lStAxe   : string ;
    lStTitre : string ;
    lStWhere : string ;
begin

  lStAxe   := GetControlText('CLV_AXE') ;
  lStTitre := TraduireMemoire('Section analytique') ;
  lStWhere := 'S_AXE="'+lStAxe+'" AND S_FERME="-"' ;

  if GetControlText('CLV_LOINIVEAU') = '1SE' then
    lStWhere := lStWhere + ' AND S_INDIRECTE="-" ' ;

  if LookupList( FListe,
                 lStTitre,
                 'SECTION',
                 'S_SECTION',
                 'S_LIBELLE',
                 lStWhere ,
                 'S_SECTION',
                 true, GetNumDisp)
   then FInfo.LoadSection( FListe.Cells[ IDX_SECTION, FListe.row], lStAxe) ;

end;

procedure TOF_CPLOIVENTIL.SoldeVentil ( vRow : integer ) ;
var lMontant : double ;
    lTotalG  : double ;
    lStMeth  : string ;
    lInCol   : integer ;
    lMax     : double ;
begin
  if isOut( vRow ) then Exit ;

  lStMeth := GetControlText('CLV_LOIMETHODE') ;
  if (lStMeth = '1MO') then
    begin
    if FStType='LOI' then Exit ;
    lInCol   := IDX_MONTANT ;
    lMax     := GetMontantTotalAbs ;
    end
  else if (lStMeth = '2PO') then
    begin
    lInCol   := IDX_CLEF ;
    lMax     := 100 ;
    end
  else Exit ;

  lTotalG  := GetTotalCol( lInCol ) ;
  lMontant := lMax - ( lTotalG - GetMontant( lInCol, vRow ) ) ;

  SetMontant( lInCol, vRow, lMontant ) ;
  RecalculGrille( vRow ) ;
  FBoGrilleModif := True ;

end;

procedure TOF_CPLOIVENTIL.InitLigne(vRow: integer);
begin
  if isOut( vRow ) then Exit ;

  FListe.Cells[ IDX_NO,      vRow] := IntToStr( vRow ) ;
  FListe.Cells[ IDX_SECTION, vRow] := '' ;
  FListe.Cells[ IDX_LIBELLE, vRow] := '' ;
  FListe.Cells[ IDX_CLEF,    vRow] := '' ;
  FListe.Cells[ IDX_MONTANT, vRow] := '' ;

end;

procedure TOF_CPLOIVENTIL.NumeroteLignes;
var i : integer ;
begin
  for i := 1 to FListe.RowCount-1 do
    FListe.Cells[ IDX_NO, i ] := IntToStr(i) ;
end;

procedure TOF_CPLOIVENTIL.BVentilTypeClick(Sender: TObject);
begin
  if not GetcontrolEnabled('BVENTILTYPE') then Exit ;
  LoiVentilClick ( nil ) ;
end;

procedure TOF_CPLOIVENTIL.Accesboutons;
var lBoStricte   : boolean ;
    lBoGrilleEna : boolean ;
    lBoAvecODA   : boolean ;
begin

  lBoGrilleEna := EstGrilleEnabled ;
  lBoStricte   := (FStType = 'SEC') and ( GetControlText('CLV_STRICTE')='X' ) ;

  if FStType = 'SEC' then
    begin
    lBoAvecODA := Assigned(FTobOda) and ( FTobOda.Detail.count > 0 ) ;
    // Acces règles
//    SetControlEnabled('CLV_LOIMETHODE', not lBoStricte ) ;
    SetControlEnabled('CLV_LOINIVEAU',  not lBoStricte ) ;

//    SetControlEnabled('BVALIDER',      not lBoAvecODA ) ;
    SetControlEnabled( 'BGENERER',     PretAGenerer(False) ) ;
    SetControlVisible( 'BRECALCULER',  (GetAction=taModif) and lBoAvecODA ) ;
    SetControlEnabled( 'BVENTILTYPE',  (GetAction=taCreat) ) ;

    // Données de calcul non modifiable si des ODA existent
    SetControlEnabled( 'CLV_DATEGENER',    not lBoAvecODA ) ;
    SetControlEnabled( 'CLV_SECTION',      not lBoAvecODA ) ;
    SetControlEnabled( 'CLV_JOURNAL',      not lBoAvecODA ) ;
    SetControlEnabled( 'FLISTE',           not lBoAvecODA ) ;
    SetControlEnabled( 'CLV_LOIMETHODE',   not lBoAvecODA ) ;
    SetControlEnabled( 'CLV_MONTANTCORR',  not lBoAvecODA ) ;

    // zoom oda
    SetControlVisible('BZOOMODA',      lBoAvecODA ) ;
    end ;

  // Edition de la grille
  SetControlEnabled( 'BLIGNEINSERT', lBoGrilleEna and not lBoStricte ) ;
  SetControlEnabled( 'BLIGNESUPPR',  lBoGrilleEna and not lBoStricte ) ;

  // Copier
  SetControlEnabled('BCOPIER',       (FStType = 'LOI') and (GetAction=taCreat) ) ;

end;

procedure TOF_CPLOIVENTIL.DateGenerChanged(Sender: TObject);
var lStDate : string ;
    lDtDate : TDateTime ;
    lStSect : string ;
begin

  if FBoOnLoad then Exit ;
  if FStType<>'SEC' then Exit ;

  lStDate := GetControlText('CLV_DATEGENER') ;
  if not HEnt1.isValidDate( lStDate ) then
    begin
    PgiInfo(TraduireMemoire('La date saisie n''est pas valide.'), Ecran.Caption ) ;
    SetControlText( 'CLV_DATEGENER', FTobLoi.GetString('CLV_DATEGENER') ) ;
    SetFocusControl( 'CLV_DATEGENER' ) ;
    Exit ;
    end ;

  lDtDate := StrToDate( lStDate ) ;
  if lDtDate = iDate1900 then Exit ;

  // Les autres traitements ne sont fait que si les infos sont renseignées
  lStSect := GetControlText('CLV_SECTION') ;
  if lStSect='' then Exit ;

  // Test d'existence en création
  if TestPresenceDate(lDtDate) then Exit ;

  // Calcul du montant
  BCalculerClick( nil ) ;

end;

procedure TOF_CPLOIVENTIL.CalculMontants;
var lStDate : string ;
    lDtDate : TDateTime ;
    lStSect : string ;
begin

  if not PretACalculer then Exit ;

  lStDate := GetControlText('CLV_DATEGENER') ;
  lDtDate := StrToDate( lStDate ) ;
  lStSect := GetControlText('CLV_SECTION') ;

  // Date non testé ?
  if TestPresenceDate(lDtDate, False, False) then Exit ;

  // Calcul du solde
  FSoldeSection := GetSoldeSectionParCpt( lStSect ) ;
  if FSoldeTraite <> 0 then
    FSoldeSection := FSoldeSection + FSoldeTraite ;
  FTobLoi.PutValue('CLV_MONTANTCALC', FSoldeSection);

  // Calcul sur solde des sections !
  UpdateClefs ;

  // Affichage des montants
  AfficheLesMontants ;
  RecalculGrille(0);

end;

procedure TOF_CPLOIVENTIL.AfficheLesMontants;
var lMontant : Double ;
begin

  lMontant := GetMontantCalc ;
  SetValeur('CLV_MONTANTCALC', Abs( lMontant ) ) ;
  SetControlText('SENSCALC',   GetSens( lMontant )           ) ;

  lMontant := GetMontantTotalAbs ;
  SetValeur('CLV_MONTANTTOTAL', lMontant ) ;
  RecalculTableau( lMontant ) ;

  SetControlText('SENSTOTAL',   GetSens( GetMontantTotal ) ) ;

end;

function TOF_CPLOIVENTIL.GetMontantTotalAbs: Double;
begin
  result := 0 ;
  if FStType='LOI' then Exit ;
  // Montant Total
  result  := Abs( Arrondi( Abs(FSoldeSection) + GetValeur('CLV_MONTANTCORR'), V_PGI.OkDecV ) ) ;
end;

procedure TOF_CPLOIVENTIL.RecalculTableau(vTotal: Double);
var lStMeth : String ;
    lSaiP   : Double ;
    lSaiM   : Double ;
    lResP   : Double ;
    lResM   : Double ;
begin

  if vTotal = 0 then
    begin
    // Affichage
    SetControlText('SAISIM', StrFMontant(0, 15, V_PGI.OkDecV, '', True )  ) ;
    SetControlText('SAISIP', StrFMontant(0, 15, 4,            '', True )  ) ;
    SetControlText('RESTEM', StrFMontant(0, 15, V_PGI.OkDecV, '', True )  ) ;
    SetControlText('RESTEP', StrFMontant(0, 15, 4,            '', True )  ) ;

    // Gestion des boutons
    SetControlEnabled( 'BGENERER', False ) ;

    Exit ;
    end
  else if vTotal < 0 then
    vTotal := Abs( vTotal ) ;

  lSaiM   := GetValeur( 'TOTALCOL2' ) ;
  lResM   := vTotal - lSaiM ;

  // sur les montants
  lStMeth := GetControlText('CLV_LOIMETHODE') ;
  if lStMeth = '1MO' then
    begin
    lSaiP   := GetValeur( 'TOTALCOL1' ) ;
    lResP   := 100.0 - lSaiP ;
    end
  // sur les %
  else if lStMeth = '2PO' then
    begin
    lSaiP   := GetValeur( 'TOTALCOL1' ) ;
    lResP   := 100.0 - lSaiP ;
    end
  else
    begin
    if lSaiM<>0 then
      begin
      lSaiP := 100.0 ;
      lResP := 0 ;
      end
    else
      begin
      lSaiP := 0 ;
      lResP := 100.0 ;
      end
    end ;

  // Affichage
  SetControlText('SAISIM', StrFMontant(lSaiM, 15, V_PGI.OkDecV, '', True )  ) ;
  SetControlText('SAISIP', StrFMontant(lSaiP, 15, 4,            '', True )  ) ;
  SetControlText('RESTEM', StrFMontant(lResM, 15, V_PGI.OkDecV, '', True )  ) ;
  SetControlText('RESTEP', StrFMontant(lResP, 15, 4,            '', True )  ) ;

  // Gestion des boutons
//  SetControlEnabled( 'BGENERER', lResM=0 ) ;

end;

function TOF_CPLOIVENTIL.PretACalculer: Boolean;
var lStDate : String ;
    lDtdate : TDateTime ;
begin
  result := False ;

  // En traitement uniquement
  if FStType<>'SEC' then Exit ;
  // Axe doit être renseigné
  if FStAxe = '' then Exit ;
  // section doit être renseignée
  if GetControlText('CLV_SECTION')='' then Exit ;
  // règle de calcul obligatoire (normalement toujours le cas)
  if GetControlText('CLV_LOIMETHODE')='' then Exit ;

  // Date correcte
  lStDate := GetControlText('CLV_DATEGENER') ;
  if not HEnt1.isValidDate( lStDate ) then Exit ;

  lDtDate := StrToDate( lStDate ) ;
  if lDtDate = iDate1900 then Exit ;

  result := true ;

end;

procedure TOF_CPLOIVENTIL.LoiMethodeChanged(Sender: TObject);
begin

  InitGrille;

  if FStType <> 'SEC' then Exit ;
  if PretACalculer then
    begin
    // ReCalcul des cles pour solde des sections ou UO
    UpdateClefs ;
    AfficheLesMontants ;
    RecalculGrille(0);
    end ;
end;

function TOF_CPLOIVENTIL.GetSens ( vMontant : Double ) : String ;
begin
  if vMontant = 0 then
    vMontant := GetMontantCalc ;
  if vMontant >= 0
    then result := 'D'
    else result := 'C' ;
end;

function TOF_CPLOIVENTIL.GetValeur(vStControl: string): Double;
begin
  result := 0 ;
  if IsNumeric( GetControlText( vStControl ) ) then
    result := Valeur( GetControlText( vStControl ) ) ;
end;

function TOF_CPLOIVENTIL.GetMontantCalc: Double;
begin
  result  := Arrondi( FSoldeSection, V_PGI.OkDecV ) ;
end;

procedure TOF_CPLOIVENTIL.LayoutCumuls;
var lRect : TRect ;
    lEdit : THEdit ;
begin
  // Placement des cumuls
  // --> 1. Cumul Colonne Clef IDX_CLEF
  lEdit          := THEdit( GetControl('TOTALCOL1') ) ;
  lEdit.Visible  := FListe.ColWidths [ IDX_CLEF ] > 0 ;
  if lEdit.Visible then
    begin
    lRect          := FListe.CellRect(IDX_CLEF,0) ;
    lEdit.Left     := lRect.Left ;
    lEdit.Width    := lRect.Right - lRect.Left ;
    end ;
  // --> 2. Cumul Clef IDX_MONTANT
  lEdit          := THEdit( GetControl('TOTALCOL2') ) ;
  lEdit.Visible  := FListe.ColWidths [ IDX_MONTANT ] > 0 ;
  if lEdit.Visible then
    begin
    lRect          := FListe.CellRect(IDX_MONTANT,0) ;
    lEdit.Left     := lRect.Left + 1 ;
    lEdit.Width    := lRect.Right - lRect.Left ;
    end ;
end;


procedure TOF_CPLOIVENTIL.DeleteRow(vRow: integer);
begin

  if GetAction = taConsult then Exit ;
  if not EstGrilleEnabled then Exit ;
  if not GetControlEnabled( 'BLIGNESUPPR' ) then Exit ; // FQ 21576 SBO interdire la modif sur loi stricte

  if vRow >= 1 then
    FListe.DeleteRow( vRow ) ;
  if FListe.RowCount < 50 then
    FListe.RowCount  := 50 ;
  NumeroteLignes ;

  FListe.Col     := IDX_SECTION ;
  FBoGrilleModif := True ;

  // --> Zone de caclul
  RecalculGrille( vRow ) ;

end;

procedure TOF_CPLOIVENTIL.InsertRow(vRow: integer);
begin

  if GetAction = taConsult then Exit ;
  if not EstGrilleEnabled then Exit ;
  if not GetControlEnabled( 'BLIGNEINSERT' ) then Exit ;// FQ 21576 SBO interdire la modif sur loi stricte

  FListe.InsertRow( vRow ) ;
  InitLigne( vRow ) ;
  NumeroteLignes ;
  FBoGrilleModif := True ;
  FListe.Col     := IDX_SECTION ;
  FListe.Row     := vRow ;

end;

procedure TOF_CPLOIVENTIL.EcranClose;
begin
  TFVierge(Ecran).BFermeClick(nil);
end;

procedure TOF_CPLOIVENTIL.RecalculGrille(vRow: integer);
var lTotalCol1 : Double ;
    lTotalCol2 : Double ;
    i          : integer ;
    lInDec     : integer ;
    iLastCol   : integer ;
    lStMeth    : string ;
begin

//  FListe.Enabled := False ;

  iLastCol := 0 ;

  // Recalcul de la ligne en cours
  lStMeth := GetControlText('CLV_LOIMETHODE') ;
  if (lStMeth = '1MO') or (lStMeth = '2PO') then
    RecalculLigne( vRow ) ;

  // Calcul des cumuls CLEF
  lTotalCol1 := 0 ;
  for i := 1 to FListe.RowCount - 1 do
    if FListe.Cells[ IDX_SECTION, i ] <> '' then
      begin
      if (lStMeth = '1MO') then
        RecalculLigne( i ) ;
      if GetMontant( IDX_CLEF, i ) <> 0 then
        begin
        lTotalCol1  := lTotalCol1 + GetMontant( IDX_CLEF, i ) ;
        iLastCol    := i ;
        end ;
      end ;

  // Affichage Total Clé
  lInDec     := GetDecimale( IDX_SECTION ) ;
  lTotalCol1 := Arrondi( lTotalCol1, lInDec ) ;
  SetControlText( 'TOTALCOL1', StrFMontant( lTotalCol1, 15, lInDec, '', True ) ) ;

  // Calcul des cumuls MONTANTS
  lTotalCol2 := 0 ;
  for i := 1 to FListe.RowCount - 1 do
    begin
    if FListe.Cells[ IDX_SECTION, i ] <> '' then
      begin
      // Recalcul du montant suivant ratio CLEF
      if lStMeth <> '1MO' then
        RecalculLigne( i ) ;
      // Total MONTANT
      lTotalCol2  := lTotalCol2 + GetMontant( IDX_MONTANT, i ) ;
      end ;
    end ;

  // Affichage Total Montant
  lInDec    := GetDecimale( IDX_MONTANT ) ;
  lTotalCol2 := Arrondi( lTotalCol2, lInDec ) ; 
  SetValeur( 'TOTALCOL2', lTotalCol2 ) ;

  // Gestion des arrondis...
  if iLastCol > 0 then
    GereArrondi( iLastCol ) ;

  // --> Zone de caclul
  RecalculTableau( GetMontantTotalAbs ) ;
  Accesboutons     ;

  if not FBoOnLoad then
    FBoGrilleModif := True ;

end;

function TOF_CPLOIVENTIL.GetDecimale(vCol: integer): integer;
var lStMeth : string ;
begin

  result := V_PGI.OkDecV ;

  if vCol = IDX_CLEF then
    begin
    lStMeth := GetControlText('CLV_LOIMETHODE') ;
    // pourcentage
    if lStMeth = '2PO' then
     result := 4
      // unité d'oeuvre
      else if (lStMeth = '3UO') or (lStMeth = '5CR') then
       result := V_PGI.OkDecQ ;
    end ;

end;

function TOF_CPLOIVENTIL.GetMontant(vCol, vRow: integer): double;
begin

  result := 0 ;

  if isOut( vRow ) then Exit ;
  if not (vCol in [IDX_CLEF, IDX_MONTANT]) then Exit ;

  if isNumeric( FListe.Cells[vCol,vRow] ) then
    result := Valeur( FListe.Cells[vCol,vRow] ) ;

end;

procedure TOF_CPLOIVENTIL.setMontant(vCol, vRow: integer; vMontant: double);
begin
  if isOut( vRow ) then Exit ;
  if not (vCol in [IDX_CLEF, IDX_MONTANT]) then Exit ;
  if Trim(FListe.Cells[IDX_SECTION,vRow])='' then
    FListe.Cells[vCol,vRow] := '' 
  else
    FListe.Cells[vCol,vRow] := StrFMontant( vMontant, 15, GetDecimale( vCol ), '', True ) ;
end;

procedure TOF_CPLOIVENTIL.GetColSuivante ( var ACol, ARow : integer ; vInSens : integer ) ;
var lInSens : integer ;
    lBoFin      : boolean ;
    lInCol      : integer ;
    lInRow      : integer ;

    procedure _PasseColSuivante ;
      begin
        // Déplacement vers l'avant
        if lInSens > 0 then
          begin
          if (lInCol + 1) < FListe.ColCount then  // Fin de ligne ?
            Inc( lInCol )
          else
            begin
            if ( lInRow + 1 ) < FListe.RowCount then // Fin de la grille ?
              begin
              // passage à la ligne suivante, 1ère colonne
              Inc( lInRow ) ;
              lInCol := FListe.FixedCols ;
              end
            else
              lBoFin := True ; // on s'arrête
            end ;
          end
        // Déplacement vers l'arrière
        else
          begin
          if (lInCol - 1) >= FListe.FixedCols then  // debut de ligne ?
            Dec( lInCol )
          else
            begin
            if ( lInRow - 1 ) >= 1 then // Début de la grille ?
              begin
              // passage à la ligne suivante, 1ère colonne
              Dec( lInRow ) ;
              lInCol := FListe.ColCount - 1 ;
              end
            else
              lBoFin := True ; // on s'arrête
            end ;
          end ;
      end ;
begin

  lBoFin  := False ;
  lInCol  := ACol ;
  lInRow  := ARow ;
  if vInSens <> 0
    then lInSens := vInSens
    else lInSens := CGetGridSens( FListe, ACol, ARow ) ;

  // Détermination de la cellule de destination
  While (not lBoFin) and ( not FListe.ColEditables[ lInCol ] ) do
    _PasseColSuivante ;

  // résultat
  ARow := lInRow ;
  Acol := lInCol ;

end;

function TOF_CPLOIVENTIL.IsOut(vRow: integer): Boolean;
begin
  result := (vRow < 1) or (vRow > FListe.rowCount - 1) ;
end;

function TOF_CPLOIVENTIL.GetTotalCol(vCol: integer): Double;
begin
  result := 0 ;

  if vCol=IDX_CLEF then
    result := Valeur( GetControlText( 'TOTALCOL1' ) )

  else if vCol=IDX_MONTANT then
    result := Valeur( GetControlText( 'TOTALCOL2' ) ) ;

end;

procedure TOF_CPLOIVENTIL.BFermeClick(Sender: TObject);
begin
  LastError := 0 ;
  EcranClose ;
end;

function TOF_CPLOIVENTIL.EstGrilleValide ( vBoParle : Boolean ) : boolean;
var lTotal     : double ;
    i          : integer ;
    lStSection : string ;
    lStMeth    : string ;
begin

  result  := False ;
  lTotal  := 0 ;
  lStMeth := GetControlText('CLV_LOIMETHODE') ;

  for i:=1 to FListe.RowCount-1 do
    begin

    lStSection := FListe.Cells[ IDX_SECTION,i] ;

    // si la section est renseignée et valide...
    if ( Trim(lStSection ) <> '' ) then
      begin

      // Section ok ?
      if not FInfo.LoadSection( lStSection, GetControlText('CLV_AXE') ) then
        begin
        FListe.Row := i ;
        FListe.Col := IDX_SECTION ;
        if vBoParle then
          PgiInfo( TraduireMemoire('La section ' + lStSection + ' n''existe pas !'), Ecran.Caption ) ;
        Exit ;
        end ;
      // Section Fermée ?
      if (GetAction=taCreat) and (FInfo.GetString('S_FERME')='X') then
        begin
        FListe.Row := i ;
        FListe.Col := IDX_SECTION ;
        if vBoParle then
          PgiInfo( TraduireMemoire('La section ' + lStSection + ' est fermée !'), Ecran.Caption ) ;
        Exit ;
        end ;
      // Cumuls sur la clef ?
      if lStMeth = '2PO'
        then lTotal := lTotal + GetMontant( IDX_CLEF, i ) ;
      end ;

    end ;

  // test du total
  if (lStMeth = '2PO') then
    begin
    // Stricte : doit être exacte à 100% !
    if (FStType='LOI') and ( GetControlText('CLV_STRICTE')='X' ) then
      begin
      if Arrondi( 100 - lTotal, 4 ) <> 0 then
        begin
        if vBoParle then
          PgiInfo( TraduireMemoire('La somme des ventilations doit être égale à 100% !'), Ecran.Caption ) ;
        Exit ;
        end
      end
    else
      if lTotal > 100 then
        begin
        if vBoParle then
          PgiInfo( TraduireMemoire('La somme des ventilations ne peut pas dépasser 100% !'), Ecran.Caption ) ;
        Exit ;
        end
    end ;

  result := True ;

end;

procedure TOF_CPLOIVENTIL.MontantCorrChanged(Sender: TObject);
begin
  if FBoOnLoad then Exit ;
  if (FStType <> 'SEC') then Exit ;
//  FListe.Enabled := False ;
  AfficheLesMontants ;
  RecalculGrille(0)  ;
  AfficheLesMontants ;
//  GereAccesGrille    ;
//  FListe.Enabled := True ;
end;

procedure TOF_CPLOIVENTIL.GereAccesGrille;
begin
  FListe.ColEditables[ IDX_SECTION ] := (FStType='LOI') or (GetControlText('CLV_STRICTE')='-') ;
end;

procedure TOF_CPLOIVENTIL.GrilleEnabled(vBoAvec: Boolean);
var lBoEdit : Boolean ;
begin

  // Option de départ
  lBoEdit := (GoRowSelect in FListe.Options) ;

  // Accès cellules
  if vBoAvec then
      begin
      FListe.Options := FListe.Options - [ GoRowSelect ] ;
      FListe.Options := FListe.Options + [ GoEditing, GoTabs, GoAlwaysShowEditor ] ;
      end
    else
      begin
      FListe.Options := FListe.Options - [ GoEditing, GoTabs, GoAlwaysShowEditor ] ;
      FListe.Options := FListe.Options + [ GoRowSelect ] ;
      end ;

  // Rafraichissement si nécessaire
  if lBoEdit <> (GoRowSelect in FListe.Options) then
    FListe.Invalidate ;

end;

function TOF_CPLOIVENTIL.EstGrilleEnabled: boolean;
begin
//  result := not (GoRowSelect in FListe.Options) ;
  result := FListe.ColEditables[IDX_SECTION] or
            FListe.ColEditables[IDX_CLEF] or
            FListe.ColEditables[IDX_MONTANT];
end;

function TOF_CPLOIVENTIL.PretAGenerer( vBoParle : boolean ) : boolean;
begin
  result := False ;

  // En traitement uniquement
  if not PretACalculer then Exit ;
  if GetMontantTotalAbs=0 then Exit ;

  if GetValeur('RESTEM')<>0 then Exit ;

  if Assigned(FTobOda) and (FTobOda.Detail.count>0) then Exit ;

  if vBoParle then
    if EstModifEnCours then
      if not SaveRecord then
        Exit ;

  result := true ;

end;

function TOF_CPLOIVENTIL.SaveRecord: boolean;
begin
  // Tests
  result := TestRecord ;

  // Enregistrement
  if result then
    begin
    result := _SaveRecord ;
    if result then
      begin
      FTobLoi.SetAllModifie( False );
      FBoGrilleModif := False ;
      if GetAction=taCreat then
        SetAction( taModif ) ;
      end ;
    end ;

end;

procedure TOF_CPLOIVENTIL.SetAction(vAction: TActionFiche);
begin
  TFVierge(Ecran).TypeAction := vAction ;
end;

procedure TOF_CPLOIVENTIL.RecalculLigne(vRow: integer);
var lStMeth   : string ;
    lMontant  : double ;
    lPourcent : double ;
    lTotal    : double ;
begin

  if isOut( vRow ) then Exit ;
  lStMeth := GetControlText('CLV_LOIMETHODE') ;

  if GetMontantTotalAbs = 0 then
    begin
    if lStMeth = '1MO'
      then SetMontant( IDX_CLEF, vRow, 0 )
      else SetMontant( IDX_MONTANT, vRow, 0 ) ;
    exit ;
    end ;

  if lStMeth = '1MO' then
    begin
    // Recup montant
    lMontant  := GetMontant( IDX_MONTANT, vRow ) ;
    // Calcul %
    lPourcent := Arrondi( (lMontant / GetMontantTotalAbs) * 100 , 4 ) ;
    SetMontant( IDX_CLEF, vRow, lPourcent ) ;
    end
  else if lStMeth = '2PO' then
    begin
    // Récup %
    lPourcent  := GetMontant( IDX_CLEF, vRow ) ;
    // calcul montant
    lMontant   := Arrondi( GetMontantTotalAbs * ( lPourcent / 100 ) , V_PGI.OkDecV ) ;
    SetMontant( IDX_MONTANT, vRow, lMontant ) ;
    end
  else
    begin
    // Recup UO
    lPourcent := GetMontant( IDX_CLEF, vRow ) ;
    // Récup Total UO
    lTotal    := GetValeur( 'TOTALCOL1' ) ;
    if lTotal <> 0 then
      begin
      // Calcul ratio
      lPourcent := ( lPourcent / lTotal ) ;
      // calcul montant
      lMontant   := Arrondi( GetMontantTotalAbs * lPourcent , V_PGI.OkDecV ) ;
      end
    else
      lMontant := 0 ;
    SetMontant( IDX_MONTANT, vRow, lMontant ) ;
    end ;

end;

procedure TOF_CPLOIVENTIL.GereArrondi(vRow: integer);
var lBoAGerer : boolean ;
    lStMeth   : string ;
    lPourc    : double ;
    lMontant  : double ;
begin
  if FStType<>'SEC' then Exit ;

  lStMeth := GetcontrolText('CLV_LOIMETHODE') ;

  if lStMeth = '2PO' then
    lBoAGerer := Arrondi( GetValeur('TOTALCOL1') - 100.0 , 4 ) = 0
  else if lStMeth = '1MO' then
    lBoAGerer := Arrondi( GetValeur('TOTALCOL2') - GetMontantTotalAbs , V_PGI.OkDecV ) = 0
  else
    lBoAGerer := True ;

  if lBoAGerer then
    begin
    // Arrondi sur le %
    if lStMeth = '1MO' then
      begin
      lPourc   := Arrondi( 100 - ( GetValeur('TOTALCOL1') - GetMontant( IDX_CLEF , vRow ) ), 4 ) ;
      setMontant( IDX_CLEF, vRow, lPourc ) ;
      SetControlText( 'TOTALCOL1', StrFMontant(100, 15, 4, '', True ) ) ;
      end
    // Arrondi sur le Montant
    else
      begin
      lMontant := Arrondi( GetMontantTotalAbs - ( GetValeur('TOTALCOL2') - GetMontant( IDX_MONTANT, vRow ) ), V_PGI.OkDecV ) ;
      setMontant( IDX_MONTANT, vRow, lMontant ) ;
      SetValeur( 'TOTALCOL2', GetMontantTotalAbs ) ;
      end ;
    end ;

end;

procedure TOF_CPLOIVENTIL.LoiVentilClick(Sender: TObject);
var lStWhere  : string ;
    lStLib    : string ;
    lStSect   : string ;
    lStTitre  : string ;
    lStNumLoi : string ;
begin

  lStWhere := 'CLV_AXE="'+FStAxe+'" AND CLV_LOITYPE="LOI"' ;
  lStSect  := GetControlText('CLV_SECTION') ;
  if lStSect<>'' then
    begin
    lStWhere := lStWhere + ' AND CLV_SECTION="' + lStSect + '"' ;
    lStLib   := 'CLV_LIBELLE' ;
    lStTitre := TraduireMemoire('Loi de répartition pour la section ')  + lStSect;
    end
  else
    begin
    lStLib   := 'CLV_SECTION || " - "  || CLV_LIBELLE' ;
    lStTitre := TraduireMemoire('Loi de répartition pour l''axe n°' + FStAxe[2] ) ;
    end ;

  lStNumLoi := Choisir( lStTitre,
                        'CLOIVENTIL',
                        lStLib,
                        'CLV_NUMERO',
                        lStWhere,
                        '') ;
  if lStNumLoi <> '' then
    begin
    SetControlText('LOIVENTIL', lStNumLoi) ;
    LoiChanged ( nil )
    end ;

  SetControlVisible('TLOIVENTIL', GetControlText('LOIVENTIL')<>''  ) ;

end;

{$IFNDEF CCSTD}
procedure TOF_CPLOIVENTIL.BZoomODAClick(Sender: TObject);
var lTLOda : TList ;
    i      : integer;
   lBoEof : boolean ;
    lQODA  : TQuery ;
    M      : RMVT ;
begin
  if not Assigned( FTobOda ) then Exit ;
  if FTobOda.Detail.count <= 0 then Exit ;

  // Si il n'y a qu'une ODA alors on ouvre la saisie directement
  if FTobOda.Detail.count = 1 then
    begin
    lQODA  := OpenSQL('SELECT ' + SQLForIdent( fbSect ) + ' FROM ANALYTIQ WHERE ' + WhereEcritureTob( tsAnal, FTobOda.Detail[0] , False ), True ) ;
    lBoEof := lQODA.Eof ;
    if not lBoEof then
      M := MvtToIdent(lQODA, fbSect, False) ;
    Ferme(lQODA) ;
    if not lBoEof then
      LanceSaisieODA( Nil, taConsult, M ) ;
    end
  // Sinon on ouvre la liste des ODA
  else
    begin
    lTLOda := TList.Create;
    for i:=0 to FTobOda.Detail.count-1 do
      lTlOda.Add( FTobOda.Detail[i] ) ;
    VisuPiecesGenere ( lTlOda, EcrAna, 13 ) ;
    FreeAndNil( lTlOda ) ;
    end ;
end;
{$ENDIF}

procedure TOF_CPLOIVENTIL.SetMontantCalc(vMontant: Double);
begin
  FSoldeSection := vMontant ;
end;

procedure TOF_CPLOIVENTIL.SetValeur(vStControl: string; vMontant: double) ;
begin
  SetControlText( vStControl, StrFMontant( vMontant, 15, V_PGI.OkDecV, '', True ) ) ;
end;

procedure TOF_CPLOIVENTIL.DateGenerExit(Sender: TObject);
begin
  DateGenerChanged( Sender ) ;
end;

function TOF_CPLOIVENTIL.TestPresenceDate( vDtDate : TDateTime ; vBoForce : boolean ; vBoParle : boolean ) : boolean ;
var lQTest    : TQuery ;
    lStMess   : string ;
    lStReq    : string ;
begin

  result := False ;
  if FStType<>'SEC' then Exit ;
  if (vDtDate = FDtDateGener) and (not vBoForce) then Exit ;

  lStReq := 'SELECT CLV_NUMERO, CLV_JOURNAL, CLV_NUMEROPIECE, CLV_DATEGENER, CLV_DATETRAITEMENT FROM CLOIVENTIL '
               + 'WHERE CLV_LOITYPE="SEC" '
                 + 'AND CLV_AXE="' + FStAxe + '" '
                 + 'AND CLV_SECTION="' + GetControlText('CLV_SECTION') + '" '
                 + 'AND CLV_JOURNAL="' + GetControlText('CLV_JOURNAL') + '" '
                 + 'AND CLV_NUMERO <>' + FTobLoi.GetString('CLV_NUMERO') + ' '
                 + 'AND (CLV_DATEGENER="' + UsDateTime( strToDate(GetControlText('CLV_DATEGENER')))+'" )' ;

  lQTest := OpenSQL( lStReq, True ) ;
  result := not lQTest.EOF ;
  if result and vBoParle then
    begin
    // Traitement existant ou déjà généré ?
    if lQTest.FindField('CLV_NUMEROPIECE').asInteger > 0
      then lStMess := TraduireMemoire('Un traitement de répartition secondaire a déjà été effectué au ' ) + GetControlText('CLV_DATEGENER')
      else lStMess := TraduireMemoire('Un traitement de répartition secondaire a déjà été saisi au ' ) + GetControlText('CLV_DATEGENER') ;
    lStMess := lStMess + TraduireMemoire(' pour la section ') + GetControlText('CLV_SECTION')
                       + TraduireMemoire(' sur le journal sélectionné.') ;

    lStMess := lStMess + #10 + #13 + TraduireMemoire('Vous ne pouvez pas saisir une nouvelle répartition sur ce même journal au ') + GetControlText('CLV_DATEGENER') + '.' ;
    PgiBox(  lStMess, Ecran.Caption ) ;
    end ;
  Ferme( lQTest ) ;

  // On a rien trouvé sur le journal sélectyionné, on recherche quand même sur les autres journaux,
  if not result then
    begin
    if ExisteSQL('SELECT CLV_NUMERO FROM CLOIVENTIL WHERE CLV_LOITYPE="SEC" '
                             + 'AND CLV_AXE="' + FStAxe + '" '
                             + 'AND CLV_SECTION="' + GetControlText('CLV_SECTION')+'"'
                             + 'AND CLV_JOURNAL<>"' + GetControlText('CLV_JOURNAL') + '" '
                             + 'AND CLV_NUMERO <>' + FTobLoi.GetString('CLV_NUMERO') + ' '
                             + 'AND CLV_DATEGENER="' + UsDateTime( strToDate(GetControlText('CLV_DATEGENER')))+'"') then
      begin
      lStMess := TraduireMemoire('Un traitement de répartition secondaire a déjà été saisi au ' ) ;
      lStMess := lStMess + GetControlText('CLV_DATEGENER') + TraduireMemoire(' pour la section ') + GetControlText('CLV_SECTION') + TraduireMemoire(' pour un autre journal d''ODA.')
                         + #10 + #13 + TraduireMemoire('Voulez-vous continuer ?') ;
      // question
      if vBoParle
        then result := PgiAsk( lStMess, Ecran.Caption) <> mrYes
        else result := False ;
      end ;
    end ;

  FBoDateOk := not result ;
  if FBoDateOk then
    FDtDateGener := vDtDate ;

end;

function TOF_CPLOIVENTIL.GetSoldeSection(vStSection: String): Double;
var lStReq  : string ;
    lQCumul : TQuery ;
begin

  result := 0 ;

  if FStType <> 'SEC' then Exit ;
  if FDtDateGener=iDate1900 then Exit ;
  if not FInfo.LoadSection( vStSection, FStAxe ) then Exit ;


  lStReq  := 'SELECT SUM(Y_DEBIT) as DEB, SUM(Y_CREDIT) as CRE FROM ANALYTIQ WHERE'
             + ' Y_AXE="' + FStAxe + '" AND Y_SECTION="' + vStSection + '"'
             + ' AND Y_DATECOMPTABLE>="' + UsDateTime( GetDebutExo ) + '"'
             + ' AND Y_DATECOMPTABLE<="' + UsDateTime( FDtDateGener ) + '"'
             + ' AND Y_QUALIFPIECE="N" AND Y_ECRANOUVEAU="N" ';

  lQCumul := OpenSQL( lStReq, True ) ;
  if not lQCumul.Eof
    then result  := lQCumul.Findfield('DEB').AsFloat - lQCumul.Findfield('CRE').AsFloat ;
  Ferme( lQCumul ) ;

end;

procedure TOF_CPLOIVENTIL.SectionExit(Sender: TObject);
var lStVal : string ;
begin
  lStVal := GetControlText('CLV_SECTION') ;
  if Trim(lStVal)='' then Exit ;
  if length(lStVal)<>GetInfoCpta( AxeTofb( FStAxe ) ).Lg then
    THEdit( GetControl('CLV_SECTION', True) ).ElipsisClick(Sender) ;
end;

procedure TOF_CPLOIVENTIL.ChargeODA( vFromTob : Tob ) ;
var lStReq : string ;
    i      : integer ;
begin

  if not Assigned( FTobOda ) then
    FTobOda := Tob.Create('LISTE_PIECEODA', nil, -1) ;

  FTobOda.ClearDetail ;

  if vFromTob = nil then
    begin
    // clé primaire : Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE
    lStReq := 'SELECT ' + GetSelectAll('Y', True) + ' FROM ANALYTIQ '
              + ' WHERE Y_JOURNAL="'       + FTobLoi.GetString('CLV_JOURNAL') + '"' +
                  ' AND Y_EXERCICE="'      + ctxExercice.QuelExoDt( FTobLoi.GetDateTime('CLV_DATEGENER') ) + '"' +
                  ' AND Y_DATECOMPTABLE="' + UsDateTime( FTobLoi.GetDateTime('CLV_DATEGENER') ) + '"' +
                  ' AND Y_NUMLIGNE=0'      +
                  ' AND Y_AXE="'           + FTobLoi.GetString('CLV_AXE') + '"' +
                  ' AND Y_NUMVENTIL=1'     +
                  ' AND Y_SECTION="'       + FTobLoi.GetString('CLV_SECTION') + '"' +
                  ' AND Y_CREERPAR="CLV" AND Y_TYPEANALYTIQUE="X"' ;

    FTobOda.LoadDetailFromSQL( lStReq ) ;
    end
  else
    begin
    for i:=0 to vFromTob.Detail.count - 1 do
      vFromTob.Detail[i].Detail[0].ChangeParent(FTobODA, -1) ;
    if FTobODA.Detail.count > 1 then
      FTobODA.Detail.sort('Y_NUMEROPIECE');
    end ;

end;

procedure TOF_CPLOIVENTIL.GetEcran;
var lBoODAPresente : Boolean ;
begin

  lBoODAPresente := (FStType = 'SEC') and Assigned( FTobOda) and (FTobOda.Detail.Count > 0 ) ;

  if GetAction = taCreat then
    begin
    FTobLoi.PutValue('CLV_AXE',              FStAxe ) ;
    FTobLoi.PutValue('CLV_SECTION',          GetControlText('CLV_SECTION') ) ;
    if FStType <> 'SEC' then
      FTobLoi.PutValue('CLV_DATEGENER', iDate1900 ) ;
    end ;

  FTobLoi.PutValue('CLV_STRICTE',          GetControlText('CLV_STRICTE') ) ;
  FTobLoi.PutValue('CLV_LIBELLE',          GetControlText('CLV_LIBELLE') ) ;
  FTobLoi.PutValue('CLV_LOINIVEAU',        GetControlText('CLV_LOINIVEAU') ) ;
  FTobLoi.PutValue('CLV_SECTIONATT',       GetControlText('CLV_SECTIONATT') ) ;

  // le reste n'est modifiable que si des ODA ne sont pas rattaché à la répazrtition
  if lBoODAPresente then Exit ;

  FTobLoi.PutValue('CLV_LOIMETHODE',       GetControlText('CLV_LOIMETHODE') ) ;
  if FStType = 'SEC'
    then FTobLoi.PutValue('CLV_DATEGENER', GetControlText('CLV_DATEGENER') ) ;
  FTobLoi.PutValue('CLV_JOURNAL',          GetControlText('CLV_JOURNAL') ) ;

  if FStType = 'SEC' then
    begin
    FTobLoi.PutValue('CLV_MONTANTCALC',      GetMontantCalc ) ;
    FTobLoi.PutValue('CLV_MONTANTCORR',      Valeur( GetControlText('CLV_MONTANTCORR') ) ) ;
    FTobLoi.PutValue('CLV_MONTANTTOTAL',     GetMontantTotal ) ;
    end
  else
    begin
    FTobLoi.PutValue('CLV_MONTANTCALC',      0 ) ;
    FTobLoi.PutValue('CLV_MONTANTCORR',      0 ) ;
    FTobLoi.PutValue('CLV_MONTANTTOTAL',     0 ) ;
    end ;

end;

procedure TOF_CPLOIVENTIL.PutEcran;
begin
  SetControlText('CLV_AXE',                FTobLoi.GetString('CLV_AXE')        ) ;
  SetControlText('CLV_SECTION',            FTobLoi.GetString('CLV_SECTION')    ) ;
  SetControlText('CLV_LIBELLE',            FTobLoi.GetString('CLV_LIBELLE')    ) ;
  SetControlText('CLV_STRICTE',            FTobLoi.GetString('CLV_STRICTE')    ) ;
  SetControlText('CLV_LOINIVEAU',          FTobLoi.GetString('CLV_LOINIVEAU')  ) ;
  SetControlText('CLV_LOIMETHODE',         FTobLoi.GetString('CLV_LOIMETHODE') ) ;
  SetControlText('CLV_SECTIONATT',         FTobLoi.GetString('CLV_SECTIONATT') ) ;
  SetControlText('CLV_JOURNAL',            FTobLoi.GetString('CLV_JOURNAL')    ) ;

  // Spécif traitement
  if FStType = 'SEC' then
    begin
    // date de génération
    SetControlText('CLV_DATEGENER',    FTobLoi.GetString('CLV_DATEGENER')    ) ;
    // montants
    SetValeur('CLV_MONTANTCORR',       FTobLoi.GetDouble('CLV_MONTANTCORR')  ) ;
    AfficheLesMontants;

    // infos de traitements
    SetInfosODA ;
    end ;

end;

function TOF_CPLOIVENTIL.GetNumDisp: integer;
begin
  result := 71780 + ValeurI( FStAxe[2]) ;
end;

procedure TOF_CPLOIVENTIL.UpdateClefs;
var lStSection : string ;
    i          : integer ;
begin
    if GetControlText('CLV_LOIMETHODE') = '4SE' then
      begin
      for i:= 1 to FListe.RowCount-1 do
        begin
        lStSection := FListe.Cells[IDX_SECTION,i] ;
        if Trim(lStSection)='' then continue ;
        SetMontant( IDX_CLEF, i, GetSoldeSection( FListe.Cells[IDX_SECTION,i] ) ) ;
        end ;
      end
    // Calcul sur UO des sections !
    else if GetControlText('CLV_LOIMETHODE') = '5CR' then
      begin
      for i:= 1 to FListe.RowCount-1 do
        begin
        lStSection := FListe.Cells[IDX_SECTION,i] ;
        if Trim(lStSection)='' then continue ;
        if not FInfo.LoadSection( lStSection, GetControlText('CLV_AXE') ) then Continue ;
        if IsNumeric( FInfo.GetString('S_UO') )
          then SetMontant( IDX_CLEF, i, Valeur( FInfo.GetString('S_UO') ) )
          else SetMontant( IDX_CLEF, i, 0 ) ;
        end ;
      end ;
end;

function TOF_CPLOIVENTIL.GetMontantTotal: Double;
begin
  result := 0 ;
  if FStType='LOI' then Exit ;
  // Montant Total
  if FSoldeSection<0
    then result := FSoldeSection - GetValeur('CLV_MONTANTCORR')
    else result := FSoldeSection + GetValeur('CLV_MONTANTCORR') ;
  result  := Arrondi( result, V_PGI.OkDecV ) ;
end;

procedure TOF_CPLOIVENTIL.BImprimerClick(Sender: TObject);
var lStTitre : string ;
begin
  if FStType = 'LOI'
    then lStTitre := 'Loi de répartition '
    else lStTitre := 'Répartition secondaire ' ;
  if GetControlText('CLV_SECTION')<>''
    then lStTitre := lStTitre + 'pour ' + GetControlText('CLV_SECTION')  + ' (Axe ' + FStAxe[2] + ')'
    else lStTitre := lStTitre + 'pour l''Axe ' + FStAxe[2] ;
  if FStType = 'SEC' then
    lStTitre := lStTitre + ' au ' + GetControlText('CLV_DATEGENER') ;

  TObjEtats.GenereEtatGrille( FListe,  lStTitre, False ) ;
end;

function TOF_CPLOIVENTIL.GetInitDate: TDateTime;
begin
  result := GetEnCours.Fin ;
{
  if (ctxPCL in V_PGI.PGIContexte) then
    result := GetEnCours.Fin //GetCPExoRef.Fin
  else if ( GetEntree.code = GetEncours.Code ) then
    result := GetEnCours.Fin
  else if ( GetEntree.code = GetSuivant.Code ) then
    result := GetSuivant.Fin
  else
    result := V_PGI.DateEntree ;}
end;

function TOF_CPLOIVENTIL.GenerePieceODA( vStCptGeneral: string ; vMontantPiece : Double ; vTobParent : Tob ) : Tob ;
var lTobV        : Tob ;
    lTobL        : Tob ;
    lDebit       : double ;
    lCredit      : double ;
    lTotal       : double ;
    lSens        : string ;
    lMontantCpt  : Double ;
    i            : integer ;
//    lDtDateModif : TDateTime ;
begin

  result := Tob.create('PIECE_ODA', vTobParent, -1) ;

  lTotal       := 0 ;
  lSens        := GetControlText('SENSTOTAL') ;
//  lDtDateModif := NowH ;

  // === Ligne de contrepartie sur compte indirecte reprenant le total du compte ===
  lTobL := GenereLigneODA( result, vStCptGeneral ) ;
//  lTobL.PutValue( 'Y_DATECREATION' ,   lDtDateModif) ;
  lTobL.PutValue( 'Y_SECTION',    FTobLoi.GetString('CLV_SECTION') );
  //Maj Montants
  if lSens='D' then
    begin
    lDebit  := 0 ;
    lCredit := vMontantPiece ;
    end
  else
    begin
    lDebit  := vMontantPiece ;
    lCredit := 0 ;
    end ;
  CSetMontants( lTobL, lDebit, lCredit, FDevise) ;


  // === Generation des lignes sur les sections directes ===
  For i:=0 to FTobLoi.Detail.count - 1 do
    begin

    lTobV := FTobLoi.Detail[i] ;
    if (lTobV.GetString('V_SECTION')='') or (lTobV.GetDouble('V_MONTANT')=0) then
      Continue ;

    lTobL := GenereLigneODA( result , vStCptGeneral ) ;
//    lTobL.PutValue( 'Y_DATECREATION' ,   lDtDateModif) ;
    lTobL.PutValue( 'Y_SECTION',         lTobV.GetString('V_SECTION') );

    // calcul du montant
    lMontantCpt := Arrondi( vMontantPiece * (lTobV.GetDouble('V_MONTANT') / GetMontantTotalAbs), V_PGI.OkDecV ) ;

    // Affectation
    if lSens='D' then
      begin
      lDebit  := lMontantCpt ;
      lCredit := 0 ;
      end
    else
      begin
      lDebit  := 0 ;
      lCredit := lMontantCpt ;
      end ;

    lTotal := lTotal + lDebit + lCredit ;

    //Maj Montants
    CSetMontants( lTobL, lDebit, lCredit, FDevise) ;

    end; // ===> For

  // === Gestion Arrondi sur la dernière ligne ===
  if Arrondi(lTotal-vMontantPiece, V_PGI.OkDecV)<>0 then
    begin
    lTobL := result.Detail[ result.Detail.count-1 ] ;
    lMontantCpt := lTobL.GetDouble('Y_DEBIT') + lTobL.GetDouble('Y_CREDIT') ;
    lMontantCpt := lMontantCpt + (vMontantPiece - lTotal) ;
    if lSens='D' then
      begin
      lDebit  := lMontantCpt ;
      lCredit := 0 ;
      end
    else
      begin
      lDebit  := 0 ;
      lCredit := lMontantCpt ;
      end ;
    CSetMontants( lTobL, lDebit, lCredit, FDevise) ;
    end ; // ==> if

end;

function TOF_CPLOIVENTIL.GenereLigneODA ( vTobPiece : Tob ; vStCptGeneral : string ) : TOB ;
var lDtDateC : TDateTime ;
    lStExo   : string ;
    lNumV    : integer ;
begin

  lDtDateC := FTobLoi.GetDateTime('CLV_DATEGENER') ;
  lStExo   := ctxExercice.QuelExoDt( lDtDateC ) ;

  result   := TOB.Create('ANALYTIQ', vTobPiece, -1);
  lNumV    := result.GetIndex+1 ;

  CPutDefautAna( result ) ;

  result.PutValue( 'Y_AXE',             FTobLoi.GetString('CLV_AXE') );
  result.PutValue( 'Y_CREERPAR',        'CLV'                        );
  result.PutValue( 'Y_LIBELLE',         FTobLoi.GetString('CLV_LIBELLE') );
  result.PutValue( 'Y_DATECOMPTABLE',   lDtDateC );

  CSetPeriode( result ) ;
  result.PutValue( 'Y_EXERCICE',        lStExo );

  result.PutValue( 'Y_NUMVENTIL',       lNumV );
  result.PutValue( 'Y_NUMLIGNE',        0 );
  result.PutValue( 'Y_NATUREPIECE',     'OD');
  result.PutValue( 'Y_TYPEANALYTIQUE',  'X');
  result.PutValue( 'Y_TOTALECRITURE',   0.0);
  result.PutValue( 'Y_TOTALDEVISE',     0.0);
  result.PutValue( 'Y_JOURNAL',         FTobLoi.GetString('CLV_JOURNAL') );
  result.PutValue( 'Y_POURCENTAGE',     0.0);
  if lNumV = 1
    then result.PutValue('Y_TYPEMVT', 'AE')
    else result.PutValue('Y_TYPEMVT', 'AL') ;

  result.PutValue( 'Y_GENERAL',         vStCptGeneral );

end;

procedure TOF_CPLOIVENTIL.SetInfosODA;
var lInMin : integer ;
    lInMax : integer ;
begin


  if Assigned(FTobODA) and ( FTobODA.Detail.Count > 0 ) then
    begin
    SetControlVisible('BZOOMODA', True ) ; // bouton Zoom
    // indicateur des numéro
    if FTobOda.Detail.count = 1 then
      begin
      SetControlText('CLV_NUMEROPIECE',    FTobLoi.GetString('CLV_NUMEROPIECE')    )  ;
      SetControlText('TCLV_NUMEROPIECE',   'Pièce générée' ) ;
      end
    else
      begin
      lInMin := FTobODA.Detail[0].GetInteger('Y_NUMEROPIECE') ;
      lInMax := FTobODA.Detail[ FTobODA.Detail.count - 1 ].GetInteger('Y_NUMEROPIECE') ;
      SetControlText('CLV_NUMEROPIECE',    IntToStr(lInMin) + ' à ' + IntToStr(lInMax) ) ;
      SetControlText('TCLV_NUMEROPIECE',   'Pièces générées' ) ;
      end ;
    // date de traitement
    SetControlText('CLV_DATETRAITEMENT',        DateToStr( FTobLoi.GetDateTime('CLV_DATETRAITEMENT')) ) ; // FTobLoi.GetString('CLV_DATETRAITEMENT') ) ;
    // montant traité
    SetValeur('MONTANTTRAITE',                  Abs(FTobLoi.GetDouble('CLV_MONTANTTOTAL') ) ) ;
    end
  else
    begin
    SetControlVisible('BZOOMODA',          False ) ; // bouton Zoom

    SetControlText('CLV_DATETRAITEMENT',   '' ) ;
    SetControlText('TCLV_NUMEROPIECE',     'Pièce générée' ) ;
    SetControlText('CLV_NUMEROPIECE',      '' )  ;
    SetValeur('MONTANTTRAITE',             0 ) ;
    end ;

end;

function TOF_CPLOIVENTIL.GetSoldeSectionParCpt(vStSection: String): Double;
var lStReq  : string ;
    i       : integer ;
begin

  result := 0 ;
  if not Assigned( FTobCumuls ) then
    FTobCumuls := Tob.Create('CPT_CUMULS', nil, -1) ;
  FTobCumuls.ClearDetail ;

  if FStType <> 'SEC' then Exit ;
  if FDtDateGener=iDate1900 then Exit ;
  if not FInfo.LoadSection( vStSection, FStAxe ) then Exit ;


  lStReq  := 'SELECT Y_GENERAL, SUM(Y_DEBIT) as DEB, SUM(Y_CREDIT) as CRE FROM ANALYTIQ WHERE'
             + ' Y_AXE="' + FStAxe + '" AND Y_SECTION="' + vStSection + '"'
             + ' AND Y_DATECOMPTABLE>="' + UsDateTime( GetDebutExo ) + '"'
             + ' AND Y_DATECOMPTABLE<="' + UsDateTime( FDtDateGener ) + '"'
             + ' AND Y_QUALIFPIECE="N" AND Y_ECRANOUVEAU="N" '
             + ' GROUP BY Y_GENERAL ' ;

  FTobCumuls.LoadDetailFromSQL( lStReq ) ;
  if FTobCumuls.Detail.count > 0 then
    for i:=0 to FTobCumuls.Detail.count-1 do
      result := result + FTobCumuls.Detail[i].GetDouble('DEB') - FTobCumuls.Detail[i].GetDouble('CRE') ;

  // Ajout Champ Supp pour le total
  if FTobCumuls.Detail.count > 0 then
    FTobCumuls.Detail[0].AddChampSup('MONTANT', True);


end;

procedure TOF_CPLOIVENTIL.JournalChanged(Sender: TObject);
begin
  if FBoOnLoad then Exit ;
  if (FStType <> 'SEC') then Exit ;

  if not HEnt1.isValidDate( GetControlText('CLV_DATEGENER') ) then Exit ;

  if (FStType='SEC') then
    if TestPresenceDate( StrToDate( GetControlText('CLV_DATEGENER') ), True ) then
       SetFocusControl('CLV_DATEGENER') ;

end;

procedure TOF_CPLOIVENTIL.SetCorrectionDansCumuls;
var lMontantCorr : double ;
    lTotalCorr   : double ;
    lMontant     : double ;
    i            : integer ;
    lCoef        : double ;
    lSens        : string ;
    lCorrSaisie  : double ;
    lInLast      : integer ;
    lTobCumuls   : Tob ;
begin
//  lMontantCorr := GetMontantCorr * () ;

  lTotalCorr  := 0 ;
  lInLast     := -1;
  lSens       := GetControlText('SENSTOTAL') ;
  lCorrSaisie := GetValeur('CLV_MONTANTCORR') ;

  For i:=0 to FTobCumuls.Detail.count - 1 do
    begin
    lTobCumuls  := FTobCumuls.Detail[i] ;
    if lSens = 'D'
      then lMontant    := lTobCumuls.GetDouble('DEB') - lTobCumuls.GetDouble('CRE')
      else lMontant    := lTobCumuls.GetDouble('CRE') - lTobCumuls.GetDouble('DEB') ;
    lCoef       := Abs( lMontant / GetMontantCalc ) ;

    // Si solde du compte en cours à 0, il n'est pas pris en compte dans l'algo
    if lMontant = 0 then
      begin
      lTobCumuls.PutValue('MONTANT', 0 ) ;
      Continue ;
      end ;

    lMontantCorr  := Arrondi( lCorrSaisie * lCoef , V_PGI.OkDecV ) ;
    lTobCumuls.PutValue('MONTANT', lMontant + lMontantCorr ) ;

    lTotalCorr    := lTotalCorr + lMontantCorr ;
    lInLast       := i ;
    end ;

  // gestion de l'arrondi sur la correction
  if ( lTotalCorr <> lCorrSaisie ) and (lInLast >= 0 ) then
    begin
    // affectation de la diff sur le dernier compte
    lMontant := FTobCumuls.Detail[ lInLast ].GetDouble('MONTANT') + lCorrSaisie - lTotalCorr ;
    FTobCumuls.Detail[ lInLast ].PutValue('MONTANT', lMontant );
    end ;


end;

function TOF_CPLOIVENTIL.GetDebutExo: TDateTime;
var lStExo : string ;
begin
  result := GetEnCours.Deb ;
  if not FBoDateOk then Exit ;
  lStExo := QuelExoDt( FDtDateGener ) ;
{  if lStExo = GetEnCours.Code
    then result := GetEnCours.Deb
  else}
  if lStExo = GetSuivant.Code 
    then result := GetSuivant.Deb ;
end;

Initialization
  registerclasses ( [ TOF_CPLOIVENTIL ] ) ;
end.
