{$A8,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W-,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{$WARN SYMBOL_DEPRECATED OFF}
{$WARN SYMBOL_LIBRARY ON}
{$WARN SYMBOL_PLATFORM OFF}
{$WARN UNIT_LIBRARY ON}
{$WARN UNIT_PLATFORM OFF}
{$WARN UNIT_DEPRECATED OFF}
{$WARN HRESULT_COMPAT ON}
{$WARN HIDING_MEMBER ON}
{$WARN HIDDEN_VIRTUAL ON}
{$WARN GARBAGE ON}
{$WARN BOUNDS_ERROR ON}
{$WARN ZERO_NIL_COMPAT ON}
{$WARN STRING_CONST_TRUNCED ON}
{$WARN FOR_LOOP_VAR_VARPAR ON}
{$WARN TYPED_CONST_VARPAR ON}
{$WARN ASG_TO_TYPED_CONST ON}
{$WARN CASE_LABEL_RANGE ON}
{$WARN FOR_VARIABLE ON}
{$WARN CONSTRUCTING_ABSTRACT ON}
{$WARN COMPARISON_FALSE ON}
{$WARN COMPARISON_TRUE ON}
{$WARN COMPARING_SIGNED_UNSIGNED ON}
{$WARN COMBINING_SIGNED_UNSIGNED ON}
{$WARN UNSUPPORTED_CONSTRUCT ON}
{$WARN FILE_OPEN ON}
{$WARN FILE_OPEN_UNITSRC ON}
{$WARN BAD_GLOBAL_SYMBOL ON}
{$WARN DUPLICATE_CTOR_DTOR ON}
{$WARN INVALID_DIRECTIVE ON}
{$WARN PACKAGE_NO_LINK ON}
{$WARN PACKAGED_THREADVAR ON}
{$WARN IMPLICIT_IMPORT ON}
{$WARN HPPEMIT_IGNORED ON}
{$WARN NO_RETVAL ON}
{$WARN USE_BEFORE_DEF ON}
{$WARN FOR_LOOP_VAR_UNDEF ON}
{$WARN UNIT_NAME_MISMATCH ON}
{$WARN NO_CFG_FILE_FOUND ON}
{$WARN MESSAGE_DIRECTIVE ON}
{$WARN IMPLICIT_VARIANTS ON}
{$WARN UNICODE_TO_LOCALE ON}
{$WARN LOCALE_TO_UNICODE ON}
{$WARN IMAGEBASE_MULTIPLE ON}
{$WARN SUSPICIOUS_TYPECAST ON}
{$WARN PRIVATE_PROPACCESSOR ON}
{$WARN UNSAFE_TYPE OFF}
{$WARN UNSAFE_CODE OFF}
{$WARN UNSAFE_CAST OFF}
{***********UNITE*************************************************
Auteur  ...... : SBO
Créé le ...... : 14/11/2006
Modifié le ... : 21/06/2007
Description .. : Source TOF de la FICHE : CPMASQUESAISIE ()
Suite ........ : 
Suite ........ : 21/06/2007:   L'onglet Test est caché en CCSTD pour ne 
Suite ........ : pas récupéré toutes les unités compta
Mots clefs ... : TOF;CPMASQUESAISIE
*****************************************************************}
Unit CPMASQUESAISIE_TOF ;

Interface

Uses StdCtrls,          // TListBox
     Controls,          // TKeyEvent
     Classes,           // TShiftState
     Windows,           // TRect, VK_INSERT, ...
     HEnt1,             // TActionFiche
     HTB97,
     ComCtrls,          // TPageControl
     HCtrls,            // THValComboBox
     uTob,
{$IFNDEF CCSTD }
     uLibSaisiePiece,   // TSaisiePiece
{$ENDIF CCSTD }
     UTOF ;

Type

  TOF_CPMASQUESAISIE = Class (TOF)
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

    // =================
    // === Variables ===
    // =================
    FTobMasque      : Tob ;             // Enregistrement en cours
    FTobMasqueCrit  : Tob ;             // Enregistrement en cours

    FStType     : string;           // Stocke le type de masque concerné
    FInNumero   : integer ;         // Identifiant de l'enregistrement en cours

    FTobChamps  : TOB ;             // Stocke les champs disponibles
    FBoModifEnCours : Boolean ;     // Indicateur de modif

    gFormuleSaisie : integer;       // Nombre de formules dans la grille saisie
    gFormuleCompl  : integer;       // Nombre de formules dans la grille complémentaire

    // Ecran de test
{$IFNDEF CCSTD }
    TSP         : TSaisiePiece ;
{$ENDIF CCSTD }

    // ==================
    // === Composants ===
    // ==================
    TLBDispo    : TListBox ;        // grille des champs Disponible
    TLBSaisie   : TListBox ;        // grille des champs de Saisie
    TLBCompl    : TListBox ;        // grille des champs d'infos Complémentaires
    ComboTable  : THValComboBox ;   // Combo des tables utilisables
    Affichage   : THRadioGroup ;    // Affichage en libellé ou en champs
    PCMasque    : TPageControl ;    // Page control principal

    SAdd        : TToolBarButton97 ;
    SDel        : TToolBarButton97 ;
    SUp         : TToolBarButton97 ;
    SDown       : TToolBarButton97 ;
    CAdd        : TToolBarButton97 ;
    CDel        : TToolBarButton97 ;
    CUp         : TToolBarButton97 ;
    CDown       : TToolBarButton97 ;

    // Evt AGL
    FOldKeyDown   : TKeyEvent ;

    // ==================================
    // === Initialisation de la fiche ===
    // ==================================
    procedure AnalyseParametre     ( vStArg: string);
    procedure InitComposants       ;
    procedure InitVariables        ;
    procedure InitEvenements       ;
    procedure InitEcran            ;

{$IFNDEF CCSTD }
    procedure InitPageTest         ;
{$ENDIF CCSTD }

    procedure ChargeChampsDispo    ;
    procedure RemplirComboTable    ;
    function  GetAction            : TActionFiche ;

    // ==================
    // === Evènements ===
    // ==================
    // --> Evènements des boutons
    procedure BUpClick            ( Sender : TObject ) ;
    procedure BDownClick          ( Sender : TObject ) ;
    procedure BAddClick           ( Sender : TObject ) ;
    procedure BDelClick           ( Sender : TObject ) ;
    procedure BCopierClick        ( Sender : TObject ) ;


    // --> Evènements de la fiche
    procedure FormKeyDown         ( Sender : TObject ; var Key: Word ; Shift: TShiftState ) ;
    procedure FormCloseQuery      ( Sender : TObject ; var CanClose: Boolean);
    function  IsVisibleFormule    ( vTLB : TListBox ) : boolean;
    function  IsVisibleCondition  ( vTLB : TListBox ) : boolean;

    procedure PCMasqueChange      ( Sender : TObject );


    // --> Evènements de liste
    procedure TLBClick            ( Sender : TObject ) ;
    procedure TLBDblClick         ( Sender : TObject ) ;
    procedure TLBDrawItem         ( Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState) ;

    // --> Evts autres compo
    procedure ChangeCondition     ( Sender : TObject ) ;
    procedure ComboTableChanged   ( Sender : TObject ) ;
    procedure AffichageClick      ( Sender : TObject ) ;
    procedure ParamColChanged     ( Sender : TObject ) ;
    procedure EnteteChanged       ( Sender : TObject ) ;
{$IFDEF SAISIEDOC}
    procedure ParamChanged        ( Sender : TObject ) ;
{$ENDIF SAISIEDOC}
    procedure LanceFormuleZone    ( Sender : TObject ) ;
    procedure ChargeTreeView      ( Tree : TTreeView ; vGrille : TListBox );
    procedure DechargeTreeView    ( Tree : TTreeView );   
    function  GetNbFormuleGrille  ( vTLB : TListBox ) : integer;
    function  GetNiemeFormuleGrille ( vTLB : TListBox ; n : integer ) : TOB;

    // ===========================
    // === Gestion des données ===
    // ===========================
    procedure NewRecord           ;
    function  LoadRecord          : boolean ;
    procedure SaveRecord          ;
    function  TestRecord          : boolean ;
    procedure UpdateTobMasque     ;
    procedure CalculIndex         ;
    function  IsFormule           ( vStNom : String ) : boolean;
    function  GetNbFormule        ( vTLB : TListBox ) : integer;
    procedure DelFormuleList      ( vTLB : TListBox ; var gFormule : integer) ;


    // Affichage
    procedure AfficheMasque       ;

    // ==================
    // === Validation ===
    // ==================
    function  ClickValide         : boolean;

    // ============================
    // === Gestion des colonnes ===
    // ============================
    function  CreerTobCol         ( vStNom,vStLib,vStTipe,vStTable: string; vParent: Tob): Tob;
    function  DupliquerTobCol     ( vTobCol : Tob ; vParent : Tob )      : Tob ;
    function  EstDansListe        ( vStNom : string ; vTLB : TListBox ) : boolean ;
    function  QuelleTLB           ( vControl : TControl ) : TListBox ;
    procedure AjouterTobCol       ( vTLBDest : TListBox; vDispoIdx: integer) ;
    procedure RemplirGrille       ( vTLB : TListBox ; vTobListeCol : Tob ) ;
    procedure SetListeChampsDispo ;
    procedure RaffraichirTLB      ( vTLB : TListBox ; vBoAvecTri : Boolean = false ) ;
    procedure RaffraichirCombo    ;
    procedure AfficheParamCol     ( vStPref : string ; vTobCol : Tob ) ;
    procedure SetTitreColSpecif   ( vTobCol : TOB ) ;
    function  GetColName          ( vTLB : TListBox ; vIdx : integer ) : string ;

    function  GetLargeurDefaut    ( vTobCol : TOB ) : integer ;


  end ;

// =============================================================================
// lancement de la fiche de paramétrage des masques de saisie
function CPLanceFiche_MasqueSaisie( vStRange, vStLequel, vStArgs : string ) : string;
// =============================================================================

// Init des zones de critères paramétrables dans un écran piloté par une TOF :
procedure CInitCritMasque         ( laTof : TOF ; vStPref : string = 'CMC' ) ;
// Init des libellé des champs du MCD
procedure CInitMasqueDEChamps ;
// Recup des multi-critères de sélection
procedure CChargeMultiCrit        ( var vTobMasque : Tob ) ;
// Enregistrement
procedure CEnregistreMasque( vTobMasque : Tob ; vTobCrit : Tob = nil ) ;
// Recup des multi-critères de sélection
function  CGetListFromCritMulti ( vTobMasque : Tob ) : TOB ; //vStCrit1, vStCrit2, vStCrit3 : string ) : TOB ;


// =============================================================================

// vérification de l'édition possible d'un champs dans la table écriture
function  CEstColEditable         ( vStNom : string ) : boolean ;
// Le paramétrage du dossier permet-il l'utilisation de la saisie paramétrable
Function  CPresenceMasque         ( vBoParle : boolean = True ) : boolean ;
// initialisation d'une tob receptrice d'un masque de saisie
procedure CInitTobMasque          ( var vTob : Tob ) ;
{$IFDEF SAISIEDOC}
function  CInitTobMasqueParam     ( var vTob : Tob ) : Tob ;
{$ENDIF SAISIEDOC}
// Chargement d'une masque de saisie type SAI suivant critères paramétré
function  CChargeMasqueSaisie     ( vStCrit1, vStCrit2, vStCrit3 : string ; var vTob : Tob ) : boolean ;
// Chargement d'une masque de saisie selon sa clé primaire
function  CChargeMasque           ( vInNum : integer ; var vTob : Tob ; vStType : string = 'SAI' ) : boolean ;
// Remplissage de la tob suivant description XML
procedure CRemplitMasqueFromXML    ( var vTob : Tob ) ;

// =============================================================================


Implementation

uses Messages,            // WM_CLOSE
     Graphics,            // fsBold
     sysutils,            // FreeAndNil

     HMsgBox,             // PgiError
     AglInit,             // StringToAction
     Vierge,              // TFVierge
     ParamSoc,            // GetParamSocSecur
     Choix,               // choisir
     ed_formu,            // EditeFormule
     ed_tools,            // TOedt
     uTobView,            // LoadLig
{$IFDEF SAISIEDOC}
     CBPPAth,             // TCbpPath
{$ENDIF SAISIEDOC}
{$IFDEF EAGLCLIENT}
     MainEagl, // AglLanceFiche
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main, // AglLanceFiche
{$ENDIF}
//     ULibMasqueSaisie,
     Ent1,                // ChargeSocieteHalley
     uLibEcriture,        // TInfoEcriture
{$IFNDEF CCSTD }
     uLibPieceCompta,     // TPieceCompta
{$ENDIF CCSTD }
     utilSoc;             // ChargePageSoc



Const
  _VersionXML : string = '01.00' ;

  // Liste des champs de la table écriture éditables en grille lors de la saisie en pièce
 _InMaxChampsEdit = 48 ; // 49 ; oter la devise
 _RecChampsEdit : array[1.._InMaxChampsEdit] of string =
 // Saisie pièce
 ('E_GENERAL'     ,'E_AUXILIAIRE'    ,'E_REFINTERNE'     ,'E_LIBELLE'       ,'E_DEBITDEV'      ,'E_CREDITDEV'
 // Saisie BOR
{ ,'E_DEVISE'}      ,'E_NATUREPIECE'   ,'E_DATECOMPTABLE'  ,'E_MODEPAIE'      ,'E_DATEVALEUR'    ,'E_DATEECHEANCE'
 // Multi-étab , banque prévi, ...
 ,'E_BANQUEPREVI' ,'E_ETABLISSEMENT'
 // Champs spécif saisie mono-section
 ,'SECTIONA1'     ,'SECTIONA2'      ,'SECTIONA3'         ,'SECTIONA4'       ,'SECTIONA5'
 // Champs complémentaires
 ,'E_REFLIBRE'    ,'E_REFEXTERNE'    ,'E_DATEREFEXTERNE' ,'E_AFFAIRE'       ,'E_QTE1'          ,'E_QTE2'
 ,'E_QUALIFQTE1'  ,'E_QUALIFQTE2'    ,'E_LIBRETEXTE0'    ,'E_LIBRETEXTE1'   ,'E_LIBRETEXTE2'   ,'E_LIBRETEXTE3'
 ,'E_LIBRETEXTE4' ,'E_LIBRETEXTE5'   ,'E_LIBRETEXTE6'    ,'E_LIBRETEXTE7'   ,'E_LIBRETEXTE8'   ,'E_LIBRETEXTE9'
 ,'E_TABLE0'      ,'E_TABLE1'        ,'E_TABLE2'         ,'E_TABLE3'        ,'E_LIBREDATE'     ,'E_LIBREBOOL0'
 ,'E_LIBREBOOL1'  ,'E_CONSO'         ,'E_LIBREMONTANT0'  ,'E_LIBREMONTANT1' ,'E_LIBREMONTANT2' ,'E_LIBREMONTANT3'
 ) ;
 //E_MODEPAIE;E_DATEVALEUR;E_DATEECHEANCE ==> Ajout 30/01/2007

// =============================================================================
function CPLanceFiche_MasqueSaisie( vStRange, vStLequel, vStArgs : string ) : string;
begin
  result := '' ;
  if not CPresenceMasque then Exit ;
  if vStLequel <> '' then
    vStArgs := vStArgs + ';CMS_NUMERO=' + vStLequel ;
  result := AGLLanceFiche('CP', 'CPMASQUESAISIE', '', '', vStArgs ) ;
end ;
// =============================================================================


// =============================================================================
// Init des zones de critères paramétrables dans un écran piloté par une TOF
procedure CInitCritMasque ( laTof : TOF ; vStPref : string ) ;
var lTobCrit : Tob ;
    lStDataType : string ;
    lStLibelle  : string ;
    lBoVisu     : boolean ;

    function _ChercheParamCrit( vStNom : string ; var vStLib, vStDT : string ) : boolean ;
    var lStVal   : string ;
        lTobV    : Tob ;
    begin
      vStLib := '' ;
      vStDT  := '' ;
      lStVal := GetParamSocSecur( vStNom, '000' ) ;
      result := False ;
      if lStVal <> '000' then
        begin
        lTobV  := lTobCrit.FindFirst( ['CO_CODE'], [lStVal], True ) ;
        result := Assigned( lTobV ) ;
        if result then
          begin
          vStLib := lTobV.GetString('CO_LIBRE') ;
          vStDT  := lTobV.GetString('CO_ABREGE') ;
          end ;
        end ;
    end ;

    procedure _InitControl( vStNom, vStLib, vStDT : string ; vBoVisu : boolean ) ;
      begin
      // affichage
      laTof.SetControlVisible( vStNom, vBoVisu ) ;           // combo
      laTof.SetControlVisible( 'T' + vStNom, vBoVisu ) ;     // THLabel
      if vBoVisu then
        begin
        // tablette
        laTof.SetControlProperty( vStNom, 'DATATYPE', vStDT ) ;
        // libellé
        laTof.SetControlText( 'T' + vStNom, vStLib ) ;
        end ;

      end ;

begin

  lTobCrit := TOB.Create('CPMASQUECRIT', nil, -1 ) ;
  lTobCrit.LoadDetailDBFromSQL('COMMUN', 'SELECT * FROM COMMUN WHERE CO_TYPE="CMC" AND CO_CODE<>"000" ') ;

  // 1er Critère
  lBoVisu := _ChercheParamCrit( 'SO_CPMASQUECRIT1' , lStLibelle, lStDataType ) ;
  _InitControl( vStPref + '_CRITERE1', lStLibelle, lStDataType, lBoVisu ) ;

  // 2ème Critère
  lBoVisu := _ChercheParamCrit( 'SO_CPMASQUECRIT2' , lStLibelle, lStDataType ) ;
  _InitControl( vStPref + '_CRITERE2', lStLibelle, lStDataType, lBoVisu ) ;

  // 3ème Critère
  lBoVisu := _ChercheParamCrit( 'SO_CPMASQUECRIT3' , lStLibelle, lStDataType ) ;
  _InitControl( vStPref + '_CRITERE3', lStLibelle, lStDataType, lBoVisu ) ;

  FreeAndNil( lTobCrit ) ;

end ;
// =============================================================================


// =============================================================================
procedure CInitMasqueDEChamps ;
var lTobCrit    : Tob ;
    lTobV       : Tob ;
    lStVal      : string ;

    procedure _SetChampsLib( vStNom , vStLib : string ) ;
    var lNumTable   : integer ;
        lNumChamp   : integer ;
      begin
      lNumTable := TableToNum( 'CMASQUECRITERES' );
      if lNumTable <= 0 then Exit ;
      lNumChamp := ChampToNum( vStNom );
      if lNumChamp <= 0 then Exit ;
        V_PGI.DEChamps[ lNumTable, lNumChamp ].Libelle := vStLib ;
      end ;

begin
  lTobCrit := TOB.Create('CPMASQUECRIT', nil, -1 ) ;
  lTobCrit.LoadDetailDBFromSQL('COMMUN', 'SELECT * FROM COMMUN WHERE CO_TYPE="CMC"') ;

  // 1er Critère
  lStVal := GetParamSocSecur( 'SO_CPMASQUECRIT1' , '000' ) ;
  lTobV  := lTobCrit.FindFirst( ['CO_CODE'], [lStVal], True ) ;
  if Assigned( lTobV )
    then lStVal := lTobV.GetString('CO_LIBRE')
    else lStVal := 'Critère non défini' ;
  _SetChampsLib( 'CMC_CRITERE1', lStVal ) ;

  // 2ème Critère
  lStVal := GetParamSocSecur( 'SO_CPMASQUECRIT2' , '000' ) ;
  lTobV  := lTobCrit.FindFirst( ['CO_CODE'], [lStVal], True ) ;
  if Assigned( lTobV )
    then lStVal := lTobV.GetString('CO_LIBRE')
    else lStVal := 'Critère non défini' ;
  _SetChampsLib( 'CMC_CRITERE2', lStVal ) ;

  // 3ème Critère
  lStVal := GetParamSocSecur( 'SO_CPMASQUECRIT3' , '000' ) ;
  lTobV  := lTobCrit.FindFirst( ['CO_CODE'], [lStVal], True ) ;
  if Assigned( lTobV )
    then lStVal := lTobV.GetString('CO_LIBRE')
    else lStVal := 'Critère non défini' ;
  _SetChampsLib( 'CMC_CRITERE3', lStVal ) ;

  FreeAndNil( lTobCrit ) ;

end ;
// =============================================================================

// Recup des multi-critères de sélection
procedure CChargeMultiCrit        ( var vTobMasque : Tob ) ;
var lQCrit         : TQuery ;
    lStCrit1       : string ;
    lStCrit2       : string ;
    lStCrit3       : string ;
    lStCritMulti1  : string ;
    lStCritMulti2  : string ;
    lStCritMulti3  : string ;
begin
  lQCrit := OpenSQL('SELECT * FROM CMASQUECRITERES WHERE '
                       + 'CMC_NUMERO=' + vTobMasque.GetString('CMS_NUMERO')
                  + ' AND CMC_TYPE="' + vTobMasque.GetString('CMS_TYPE') + '" ', True) ;

  lStCritMulti1  := '' ;
  lStCritMulti2  := '' ;
  lStCritMulti3  := '' ;
  while not lQCrit.Eof do
    begin
    lStCrit1  := lQCrit.FindField('CMC_CRITERE1').AsString ;
    lStCrit2  := lQCrit.FindField('CMC_CRITERE2').AsString ;
    lStCrit3  := lQCrit.FindField('CMC_CRITERE3').AsString ;
    if (lStCrit1 <> '') and ( pos( lStCrit1, lStCritMulti1 ) = 0 )  then
      lStCritMulti1 := lStCritMulti1 + lStCrit1 + ';' ;
    if (lStCrit2 <> '') and ( pos( lStCrit2, lStCritMulti2 ) = 0 )  then
      lStCritMulti2 := lStCritMulti2 + lStCrit2+ ';' ;
    if (lStCrit3 <> '') and ( pos( lStCrit3, lStCritMulti3 ) = 0 )  then
      lStCritMulti3 := lStCritMulti3 + lStCrit3 + ';' ;
    lQCrit.Next ;
    end ;
  Ferme( lQCrit ) ;

  if ( lStCritMulti1='' ) and ( GetParamSocSecur( 'SO_CPMASQUECRIT1' , '000' ) <> '000' ) then
    lStCritMulti1 := TraduireMemoire('<<Tous>>') ;
  if ( lStCritMulti2='' ) and ( GetParamSocSecur( 'SO_CPMASQUECRIT2' , '000' ) <> '000' ) then
    lStCritMulti2 := TraduireMemoire('<<Tous>>') ;
  if ( lStCritMulti3='' ) and ( GetParamSocSecur( 'SO_CPMASQUECRIT3' , '000' ) <> '000' ) then
    lStCritMulti3 := TraduireMemoire('<<Tous>>') ;

  vTobMasque.PutValue('CMS_CRITERE1', lStCritMulti1);
  vTobMasque.PutValue('CMS_CRITERE2', lStCritMulti2);
  vTobMasque.PutValue('CMS_CRITERE3', lStCritMulti3);

end ;
// =============================================================================

// MAJ en base du masque et des lignes de critères correspondantes
procedure CEnregistreMasque( vTobMasque : Tob ; vTobCrit : Tob ) ;
var lTobCrit       : Tob ;
    lAction        : TActionFiche ;
    lQNum          : TQuery ;
    lBoFree        : Boolean ;
    i              : integer ;
begin

  // Affectation numero interne
  if vTobMasque.GetInteger('CMS_NUMERO') = 0  then
    begin
    lQNum := OpenSQL('SELECT MAX(CMS_NUMERO) AS MAXNUM FROM CMASQUESAISIE WHERE CMS_TYPE="' + vTobMasque.GetString('CMS_TYPE') + '"', True) ;
    if lQNum.Eof
      then vTobMasque.PutValue('CMS_NUMERO', 1)
      else vTobMasque.PutValue('CMS_NUMERO', lQNum.FindField('MAXNUM').AsInteger + 1 ) ;
    Ferme( lQNum ) ;
    lAction := taCreat ;
    end
  else
    begin
    lAction := taModif ;
    // Suppression des anciens critères
    ExecuteSQL('DELETE FROM CMASQUECRITERES WHERE CMC_NUMERO="' + vTobMasque.GetString('CMS_NUMERO')
                                         + '" AND CMC_TYPE="'   + vTobMasque.GetString('CMS_TYPE') + '" ') ;
    end ;


  // Récupération des critères
  lBoFree := not Assigned( vTobCrit ) ;
  if lBoFree
    then lTobCrit := CGetListFromCritMulti( vTobMasque )
    else begin
         lTobCrit := vTobCrit ;
         // maj du numéro
         for i := 0 to lTobCrit.Detail.count - 1 do
           lTobCrit.Detail[i].PutValue('CMC_NUMERO', vTobMasque.GetInteger('CMS_NUMERO') ) ;
         end ;

  // Enregistrement du masque
  if lAction = taCreat
    then vTobMasque.InsertDB(nil)
    else vTobMasque.UpdateDB ;

  // Enregistrement des lignes de critères
  lTobCrit.InsertDB(nil) ;
  if lBoFree then
    FreeAndNil( lTobCrit ) ;

end ;

// =============================================================================

// Recup des multi-critères de sélection
function  CGetListFromCritMulti ( vTobMasque : Tob ) : TOB ;
var lStCritMulti1  : string ;
    lStCritMulti2  : string ;
    lStCritMulti3  : string ;
    lTObRef        : Tob ;
    i              : integer ;
    lInCount       : integer ;

    procedure _AjouteTobCrit( vStCrit : string ; vInNiv : integer ; vTobRef : Tob ) ;
      var lTobTmp : Tob ;
          i       : integer ;
      begin
        lTobTmp := Tob.Create('CMASQUECRITERES', result, -1 ) ;
        lTobTmp.PutValue( 'CMC_NUMERO',                    vTobMasque.GetInteger('CMS_NUMERO') );
        lTobTmp.PutValue( 'CMC_TYPE',                      vTobMasque.GetString('CMS_TYPE')    );
        // Valeur du critère en cours de traitement
        lTobTmp.PutValue( 'CMC_CRITERE'+IntToStr(vInNiv),  vStCrit                             );
        // copie des valeurs des critères inférieures
        if Assigned( vTobRef ) then
          for i := 1 to (vInNiv - 1) do
            lTobTmp.PutValue( 'CMC_CRITERE'+IntToStr(i),   vTobRef.GetValue( 'CMC_CRITERE'+IntToStr(i) ) );
      end ;

    procedure _TraiteNiveau(vInNiv : integer; vStCritMulti : String ; vTobRef : Tob = nil ) ;
      var lStCrit    : string ;
          lStChp     : string ;
          lBoPremier : boolean ;
      begin
        lStChp     := 'CMC_CRITERE' + IntToStr( vInNiv ) ;
        lBoPremier := Assigned( vTobRef ) ;
        if vStCritMulti <> '' then
          begin
          while (vStCritMulti <> '') do
            begin
            lStCrit := ReadTokenSt( vStCritMulti ) ;
            if lStCrit<> '' then
              begin
              if lBoPremier
                then vTobRef.PutValue ( lStChp,  lStCrit )
                else _AjouteTobCrit   ( lStCrit, vInNiv, vTobRef ) ;
              lBoPremier := False ;
              end ;
            end
          end
        else
          begin
          if lBoPremier
            then vTobRef.PutValue ( lStChp, '' )
            else _AjouteTobCrit   ( '', vInNiv, vTobRef ) ;
          end ;
      end ;

begin

  lStCritMulti1  := '';
  lStCritMulti2  := '';
  lStCritMulti3  := '';

  result := Tob.Create('V_LISTECRIT', nil, -1 );

  // Récupération des critères
  if ( vTobMasque.GetString('CMS_CRITERE1') <> '' ) and ( vTobMasque.GetString('CMS_CRITERE1') <> TraduireMemoire('<<Tous>>') ) then
      lStCritMulti1  := vTobMasque.GetString('CMS_CRITERE1') ;
  if GetParamSocSecur('SO_CPMASQUECRIT2', '000') <> '000' then
    begin
    if ( vTobMasque.GetString('CMS_CRITERE2') <> '' ) and ( vTobMasque.GetString('CMS_CRITERE2') <> TraduireMemoire('<<Tous>>') ) then
      lStCritMulti2  := vTobMasque.GetString('CMS_CRITERE2') ;
    end ;
  if GetParamSocSecur('SO_CPMASQUECRIT3', '000') <> '000' then
    begin
    if ( vTobMasque.GetString('CMS_CRITERE3') <> '' ) and ( vTobMasque.GetString('CMS_CRITERE3') <> TraduireMemoire('<<Tous>>') ) then
      lStCritMulti3  := vTobMasque.GetString('CMS_CRITERE3') ;
    end ;

  // MISE EN PLACE DES CRITERES 1ER NIVEAU
  _TraiteNiveau( 1, lStCritMulti1, nil ) ;

  // MISE EN PLACE DES CRITERES 2EME NIVEAU
  lInCount := result.Detail.count ;
  for i := 0 to lInCount - 1 do
    begin
    lTobRef := result.Detail[i] ;
    _TraiteNiveau( 2, lStCritMulti2, lTobRef ) ;
    end ;

  // MISE EN PLACE DES CRITERES 3EME NIVEAU
  lInCount := result.Detail.count ;
  for i := 0 to lInCount - 1 do
    begin
    lTobRef := result.Detail[i] ;
    _TraiteNiveau( 3, lStCritMulti3, lTobRef ) ;
    end ;

end ;

// =============================================================================

function  CEstColEditable         ( vStNom : string ) : boolean ;
var i : integer ;
begin
  result := False ;

  for i := 1 to _InMaxChampsEdit do
    if _recChampsEdit[i] = vStNom then
      begin
      result := True ;
      Exit ;
      end ;
end ;
// =============================================================================

Function  CPresenceMasque( vBoParle : boolean ) : boolean ;
begin
result := False ;
if not ExisteSQL('SELECT CMS_NUMERO FROM CMASQUESAISIE WHERE CMS_TYPE = "CEG"') then
  begin
  if vBoParle then
    PGIBox( TraduireMemoire('Les masques de saisie CEGID sont manquants ! Veuillez importer les masques CEGID dans votre dossier.') ) ;
  end
else if GetParamSocSecur('SO_CPMASQUEDEFAUT', '') = '' then
  begin
  if vBoParle then
    begin
    PGIBox( TraduireMemoire('Vous devez renseigner le masque de saisie par défaut pour pourvoir utiliser ce module ! ') ) ;
    ParamSociete(False,'','SCO_CPSAISIE','',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000) ;
    result := GetParamSocSecur('SO_CPMASQUEDEFAUT', '', True) <> '' ;
    end ;
  end
else
  result := true ;
end ;

// =============================================================================
procedure CInitTobMasque( var vTob : Tob ) ;
begin

  if not Assigned( vTob )
    then vTob := Tob.Create('CMASQUESAISIE', nil, -1)
    else vTob.ClearDetail ;

  // Description grille de saisie
  Tob.Create('GRILLE', vTob, -1);
  vTob.Detail[0].AddChampSupValeur('TYPE', 'SAISIE') ;

  // Description grille de compléments
  Tob.Create('GRILLE', vTob, -1);
  vTob.Detail[1].AddChampSupValeur('TYPE', 'COMPLEMENTS') ;

{$IFDEF SAISIEDOC}
  // Paramètres
  CInitTobMasqueParam( vTob ) ;
{$ENDIF SAISIEDOC}

  vTob.AddChampSup('CMS_CRITERE1', False);
  vTob.AddChampSup('CMS_CRITERE2', False);
  vTob.AddChampSup('CMS_CRITERE3', False);

end ;

{$IFDEF SAISIEDOC}
function CInitTobMasqueParam ( var vTob : Tob ) : Tob ;
begin
  // Paramètres
  result := Tob.Create('GRILLE', vTob, -1);
  result.AddChampSupValeur('TYPE',            'PARAMETRES') ;
  result.AddChampSupValeur('VIEWERACTIF',     '-');
  result.AddChampSupValeur('VIEWERPOSITION',  '');
  result.AddChampSupValeur('VIEWERDIRECTORY', TCbpPath.GetCommonDocuments);
  result.AddChampSupValeur('VIEWEREXTENSION', '*');
  result.AddChampSupValeur('VIEWERIMPORTGED', '-');
end ;
{$ENDIF SAISIEDOC}
// =============================================================================

// =============================================================================
function  CChargeMasqueSaisie( vStCrit1, vStCrit2, vStCrit3 : string ; var vTob : Tob ) : boolean ;
var lQMasque  : TQuery ;
    lInNum    : integer ;
    lStSQL    : String ;
begin
  CInitTobMasque( vTob ) ;
  lInNum    := 0 ;

  // Valeur Exacte
  lStSQL := 'SELECT CMC_NUMERO FROM CMASQUECRITERES WHERE CMC_TYPE="SAI" AND CMC_CRITERE1="' + vStCrit1 + '" ' ;
  if GetParamSocSecur('SO_CPMASQUECRIT2', '000') <> '000' then
    lStSQL := lStSQL + ' AND CMC_CRITERE2="' + vStCrit2 + '" ' ;
  if GetParamSocSecur('SO_CPMASQUECRIT3', '000') <> '000' then
    lStSQL := lStSQL + ' AND CMC_CRITERE3="' + vStCrit3 + '" ' ;

  lQMasque := OpenSQL( lStSQL, True ) ;
  result   := not lQMasque.Eof ;
  if result then
    lInNum := lQMasque.FindField('CMC_NUMERO').AsInteger ;
  Ferme( lQMasque ) ;

  if lInNum>0 then
    CChargeMasque( lInNum, vTOb ) ;

end ;
// =============================================================================

// =============================================================================
function  CChargeMasque( vInNum : integer ; var vTob : Tob ; vStType : string ) : boolean ;
begin
  CInitTobMasque( vTob ) ;

  vTob.PutValue( 'CMS_NUMERO', vInNum  ) ;
  vTob.PutValue( 'CMS_TYPE',   vStType ) ;

  result := vTob.LoadDB ;

  if result then
    CRemplitMasqueFromXML( vTob ) ;
end ;
// =============================================================================
// =============================================================================
procedure CRemplitMasqueFromXML    ( var vTob : Tob ) ;
var lStream   : TStringStream ;
    lBoEntete : boolean ;
    lStDecode : string ;
    lTob      : Tob ;
    i         : integer ;
begin

  // Descriptif XML
  lTob := Tob.Create('DESCRIPTION', nil, -1) ;
  lStream := TStringStream.Create( vTob.GetString('CMS_DESCRIPTION') ) ;
  lTob.LoadFromXMLStream(lStream, lBoEntete, lStDecode ) ;
  FreeAndNil( lStream ) ;

  vTob.ClearDetail ;
  for i := lTob.Detail.Count - 1 downto 0 do
    lTob.Detail[i].ChangeParent( vTob, 0 ) ;

  // informations de version
  vTob.AddChampSupValeur('VERSION',   lTob.GetString('VERSION') ) ;
  vTob.AddChampSupValeur('DATEMODIF', lTob.GetString('DATEMODIF') ) ;

  FreeAndnil( lTob ) ;

end ;
// =============================================================================

// =============================================================================
// ======                         SOURCE DE La TOF                         =====
// =============================================================================

procedure TOF_CPMASQUESAISIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUESAISIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUESAISIE.OnUpdate ;
begin
  Inherited ;

  LastError := 0 ;
  if not ClickValide
    then LastError := 1
    else FBoModifEnCours := False ;

end ;

procedure TOF_CPMASQUESAISIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUESAISIE.OnArgument (S : String ) ;
begin
  Inherited ;

  AnalyseParametre( S );
  InitComposants ;
  InitVariables  ;
  InitEvenements ;
  InitEcran      ;

end ;

procedure TOF_CPMASQUESAISIE.OnClose ;
{$IFNDEF CCSTD }
var lInfo  : TInfoEcriture ;
    lPiece : TPieceCompta ;
{$ENDIF CCSTD }
begin
  Inherited ;

  if assigned(FTobMasque) then
    FreeAndNil(FTobMasque) ;

  if Assigned( FTobMasqueCrit ) then
    FreeAndNil( FTobMasqueCrit ) ;

  if Assigned( FTobChamps ) then
    FreeAndNil(FTobChamps) ;

{$IFNDEF CCSTD }
  if Assigned( TSP ) then
    begin
    lPiece := TSP.Piece ;
    if Assigned( lPiece ) then
      begin
      lInfo := lPiece.Info ;
      FreeAndNil( lPiece ) ;
      if Assigned( lInfo ) then
        FreeAndNil( lInfo  ) ;
      end ;
    FreeAndNil( TSP ) ;
    end ;
{$ENDIF CCSTD }

end ;

procedure TOF_CPMASQUESAISIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUESAISIE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMASQUESAISIE.AnalyseParametre(vStArg: string);
var
  lStChaine: string;
begin

  while vStArg <> '' do
    begin
    lStChaine := ReadTokenSt(vStArg);

    // Type de masque concerné
    if Pos('CMS_NUMERO=', Uppercase(lStChaine)) > 0 then
      FInNumero := ValeurI( Copy(lStChaine, 12, length(lStChaine)) )

    // Type de masque concerné
    else if Pos('CMS_TYPE=', Uppercase(lStChaine)) > 0 then
      FStType := Copy(lStChaine, 10, length(lStChaine))

    // Type de masque concerné
    else if Pos('ACTION=', Uppercase(lStChaine)) > 0 then
      begin
      TFVierge(Ecran).TypeAction := StringToAction( lStChaine ) ;
      end

    ;

    end ;

  if FStType='' then
    FStType := 'SAI' ;

end;

procedure TOF_CPMASQUESAISIE.InitComposants;
begin

  // Liste
  TLBDispo   := TListBox( GetControl('TLBDispo', True ) ) ;      // grille des champs Disponible
  TLBSaisie  := TListBox( GetControl('TLBSaisie', True ) ) ;     // grille des champs de Saisie
  TLBCompl   := TListBox( GetControl('TLBCompl', True ) ) ;      // grille des champs d'infos Complémentaires

  // Autres
  ComboTable := ThValComboBox( GetControl('CBTABLEDISPO', True) ) ;   // Combo des tables utilisables
  Affichage  := THRadioGroup( GetControl('AFFICHAGE', True) ) ;       // Affichage en libellé ou en champs
  PCMasque   := TPageControl( GetControl('PCMASQUE', True) ) ;        // Page control principal

  // Boutons
  SAdd       := TToolBarButton97( GetControl('SADD',  True) ) ;
  SDel       := TToolBarButton97( GetControl('SDEL',  True) ) ;
  SUp        := TToolBarButton97( GetControl('SUP',   True) ) ;
  SDown      := TToolBarButton97( GetControl('SDOWN', True) ) ;

  CAdd       := TToolBarButton97( GetControl('CADD',  True) ) ;
  CDel       := TToolBarButton97( GetControl('CDEL',  True) ) ;
  CUp        := TToolBarButton97( GetControl('CUP',   True) ) ;
  CDown      := TToolBarButton97( GetControl('CDOWN', True) ) ;

  TToolBarButton97( GetControl('BVALIDER', True) ).Default := False ;

end;

procedure TOF_CPMASQUESAISIE.InitEcran;
begin

  Ecran.ShowHint := True ;

  // Paramétrage des libellé des critères
  CInitCritMasque( self, 'CMS' ) ;

  // Remplissage de la liste des tables
  RemplirComboTable ;

  // Sélection par défaut de la table affichée
  ComboTable.ItemIndex := 0 ;
  ComboTableChanged(nil) ;

  // Onglet guide non utilisé pour le moment
  if GetControl('PGUIDE')<>nil then
    SetControlVisible('PGUIDE', False ) ;

{$IFDEF CCSTD}
  // L'onglet Test est caché en CCSTD pour ne pas récupéré toutes les unités compta
  SetControlVisible('PTEST', False ) ;
{$ENDIF CCSTD}
{$IFNDEF SAISIEDOC}
  // L'onglet Paramètres n'est visible que si l'on compile avec la directive pour le moment
  if GetControl('PPARAM')<>nil then
    SetControlVisible('PPARAM', False ) ;
{$ENDIF SAISIEDOC}

  SetControlEnabled ('BCOPIER', GetAction = taCreat) ;
  SetControlProperty('BCOPIER', 'HINT', TraduireMemoire('Copier un masque de saisie')) ;


end;

procedure TOF_CPMASQUESAISIE.InitEvenements;
begin

  // EVT Fiche
  FOldKeyDown                       := TFVierge(Ecran).OnKeyDown ;
  TFVierge(Ecran).OnKeyDown         := FormKeyDown ;
  TFVierge(Ecran).OnCloseQuery      := FormCloseQuery ;
  TFVierge(Ecran).OnAfterFormShow   := OnAfterFormShow ;

  //  Evt composant
  ComboTable.OnClick                := ComboTableChanged ;
  Affichage.OnClick                 := AffichageClick ;
  PCMasque.OnChange                 := PCMasqueChange ;  
  THEdit( GetControl('SCONDITION', True) ).OnDblClick     := LanceFormuleZone ;
  THEdit( GetControl('SFORMULE', True) ).OnDblClick       := LanceFormuleZone ;
  THEdit( GetControl('CFORMULE', True) ).OnDblClick       := LanceFormuleZone ;
  THEdit( GetControl('SCONDITION', True) ).OnElipsisClick := LanceFormuleZone ;
  THEdit( GetControl('SFORMULE', True) ).OnElipsisClick   := LanceFormuleZone ;
  THEdit( GetControl('CFORMULE', True) ).OnElipsisClick   := LanceFormuleZone ;
  THEdit( GetControl('SCONDITION', True) ).OnChange       := ChangeCondition ;

  // zone d'entête
  THEdit( GetControl('CMS_LIBELLE', True ) ).OnExit := EnteteChanged ;
  THMultiValComboBox( GetControl('CMS_CRITERE1', True ) ).OnExit := EnteteChanged ;
  THMultiValComboBox( GetControl('CMS_CRITERE2', True ) ).OnExit := EnteteChanged ;
  THMultiValComboBox( GetControl('CMS_CRITERE3', True ) ).OnExit := EnteteChanged ;

{$IFDEF SAISIEDOC}
  // onglet paramètres
  TCheckBox( GetControl('VIEWERACTIF' )).OnExit        := ParamChanged ;
  THEdit( GetControl('VIEWERPOSITION' )).OnExit        := ParamChanged ;
  THEdit( GetControl('VIEWERDIRECTORY' )).OnExit       := ParamChanged ;
  THEdit( GetControl('VIEWEREXTENSION' )).OnExit       := ParamChanged ;
  TCheckBox( GetControl('VIEWERIMPORTGED' )).OnExit    := ParamChanged ;
{$ENDIF SAISIEDOC}

  //  Evts Grille
  TLBSaisie.OnClick                 := TLBClick ;
  TLBCompl.OnClick                  := TLBClick ;
  TLBDispo.OnDblClick               := TLBDblClick ;
  TLBSaisie.OnDblClick              := TLBDblClick ;
  TLBCompl.OnDblClick               := TLBDblClick ;
  TLBSaisie.OnDrawItem              := TLBDrawItem ;

  // Evts Paramètres de col
  THEdit( GetControl('STITRE', True) ).OnExit        := ParamColChanged ;
  THEdit( GetControl('SLARGEUR', True) ).OnExit      := ParamColChanged ;
  THEdit( GetControl('SFORMULE', True) ).OnExit      := ParamColChanged ;
  THEdit( GetControl('SCONDITION', True) ).OnExit    := ParamColChanged ;
  THEdit( GetControl('SAVECLIBELLE', True) ).OnClick := ParamColChanged ;
  THEdit( GetControl('SREADONLY', True) ).OnClick    := ParamColChanged ;
  THEdit( GetControl('SFORMULE', True) ).OnClick     := ParamColChanged ;
  THEdit( GetControl('SCONDITION', True) ).OnClick   := ParamColChanged ;

  THEdit( GetControl('CTITRE', True) ).OnExit        := ParamColChanged ;
  THEdit( GetControl('CLARGEUR', True) ).OnExit      := ParamColChanged ;
  THEdit( GetControl('CFORMULE', True) ).OnExit      := ParamColChanged ;
  THEdit( GetControl('CAVECLIBELLE', True) ).OnClick := ParamColChanged ;
  THEdit( GetControl('CFORMULE', True) ).OnClick     := ParamColChanged ;

  //  Evts Boutons
  SAdd.OnClick   := BAddClick ;
  SDel.OnClick   := BDelClick ;
  SUp.OnClick    := BUpClick ;
  SDown.OnClick  := BDownClick ;

  CAdd.OnClick   := BAddClick ;
  CDel.OnClick   := BDelClick ;
  CUp.OnClick    := BUpClick ;
  CDown.OnClick  := BDownClick ;

  TToolBarButton97( GetControl('BCOPIER',True) ).OnClick := BCopierClick;

end;

function TOF_CPMASQUESAISIE.LoadRecord: boolean;
begin
  // chargement
  result := CChargeMasque( FInNumero, FTobMasque ) ;

  FTobMasqueCrit := nil ;

  // Recup des multi-critères de sélection
  CChargeMultiCrit( FTobMasque ) ;

  // Affichage écran
  AfficheMasque ;

  FBoModifEnCours := False ;

end;

procedure TOF_CPMASQUESAISIE.SaveRecord ;
var lTobTemp : Tob ;
    i        : integer ;
begin

  // ======================
  // === ENREGISTREMENT ===
  // ======================
  // On vire les tob virtuelles pour éviter pb lors de la maj base
  lTobTemp := Tob.Create('DESCRIPTION', nil, -1 );
  for i := FTobMasque.Detail.count-1 downto 0 do
    FTobMasque.Detail[i].ChangeParent(lTobTemp, 0) ;

  // Enregistrement
  CEnregistreMasque( FTobMasque, FTobMasqueCrit ) ;

  // Récup numéro généré
  FInNumero := FTobMasque.GetInteger('CMS_NUMERO') ;
  SetControlText('CMS_NUMERO', IntTostr(FInNumero) ) ;

  // Maj indicateurs de modif
  FBoModifEnCours := False ;

  // -------------------------
  // ReMAJ des TOB description
  // -------------------------
  if Assigned( lTobTemp ) then
    begin
    // Récup des tob virtuelles
    for i := lTobTemp.Detail.count-1 downto 0 do
      lTobTemp.Detail[i].ChangeParent(FTobMasque, 0) ;
    // Libération mémoire
    FreeAndNil( lTobTemp ) ;
    end ;

  if Assigned( FTobMasqueCrit ) then
    FreeAndNil( FTobMasqueCrit ) ;

end;

function TOF_CPMASQUESAISIE.TestRecord: boolean;
var lTobCrit : Tob ;
    i        : integer ;
    lStMsg   : string ;
begin

  result := False ;

  if Trim( GetControlText('CMS_LIBELLE') ) = '' then
    begin
    PgiError('Vous devez saisir un libellé !', Ecran.Caption ) ;
    SetFocusControl( 'CMS_LIBELLE' ) ;
    Exit ;
    end ;

  // Pour les masques , Un seul par clé crit1/2/3
  if (GetAction = taCreat) and (FStType = 'SAI') then
    begin
    // Récupération des critères
    if Assigned( FTobMasqueCrit ) then
      FreeAndNil( FTobMasqueCrit ) ;

    FTobMasqueCrit := CGetListFromCritMulti( FTobMasque ) ;
    // test existence pour chacun des trio
    for i := 0 to FTobMasqueCrit.Detail.count - 1 do
      begin
      lTobCrit := FTobMasqueCrit.Detail[i] ;

      if ExisteSQL('SELECT CMC_NUMERO FROM CMASQUECRITERES WHERE CMC_TYPE="SAI"'
                           + ' AND CMC_CRITERE1="' + lTobCrit.GetString('CMC_CRITERE1') + '" '
                           + ' AND CMC_CRITERE2="' + lTobCrit.GetString('CMC_CRITERE2') + '" '
                           + ' AND CMC_CRITERE3="' + lTobCrit.GetString('CMC_CRITERE3') + '" '
                           + ' AND CMC_NUMERO<>'   + FTobMasque.GetString('CMS_NUMERO') ) then
        begin
        lStMsg := 'Un masque de saisie a déjà été saisi avec les critères suivants : ' ;
        if lTobCrit.GetString('CMC_CRITERE1') <> ''
          then lStMsg := lStMsg + lTobCrit.GetString('CMC_CRITERE1')
          else lStMsg := lStMsg + '<<Tous>>' ;
        if GetParamSocSecur('SO_CPMASQUECRIT2', '000')<>'000' then
          if lTobCrit.GetString('CMC_CRITERE2') <> ''
            then lStMsg := lStMsg + ' / ' + lTobCrit.GetString('CMC_CRITERE2')
            else lStMsg := lStMsg + ' / ' + '<<Tous>>' ;
        if GetParamSocSecur('SO_CPMASQUECRIT3', '000')<>'000' then
          if lTobCrit.GetString('CMC_CRITERE3') <> ''
            then lStMsg := lStMsg + ' / ' + lTobCrit.GetString('CMC_CRITERE3')
            else lStMsg := lStMsg + ' / ' + '<<Tous>>' ;
        lStMsg := lStMsg + ' ! ' ;
        PgiError( lStMsg , Ecran.Caption ) ;
        SetFocusControl( 'CMS_CRITERE1' ) ;
        FreeAndNil( FTobMasqueCrit ) ;
        Exit ;
        end ;
      end ;

    end ;

  result := True ;

end;

procedure TOF_CPMASQUESAISIE.BDownClick(Sender: TObject);
var lBouton : TToolBarButton97 ;
    lTLB    : TListBox ;
    lOldIdx : integer ;
    lNbDep  : integer ;
    i       : integer ;
    lTobG   : Tob ;
begin
  if not assigned(sender) then Exit ;
  if GetAction = taConsult then Exit ;

  lBouton := TToolBarButton97( Sender ) ;
  if (not lBouton.Enabled) or (not lBouton.visible) then Exit ;

  if Copy(lBouton.Name, 1, 1) = 'S' then
    begin
    lTLB  := TLBSaisie ;
    lTobG := FTobMasque.Detail[0] ;
    end
  else
    begin
    lTLB := TLBCompl ;
    lTobG := FTobMasque.Detail[1] ;
    end ;

  // Test des cellules indissociables pour la grille de saisie
  lOldIdx := lTLB.ItemIndex;
  if ( lTLB = TLBSaisie ) and ( pos(GetColName( lTLB, lOldIdx), 'E_NUMLIGNE;E_GENERAL;E_AUXILIAIRE;E_DEBITDEV;E_CREDITDEV')>0 )
    then Exit ;

  if (lOldIdx >= 0) and ( lOldIdx < lTLB.Items.Count - 1) then
    begin
    // Test des cellules indissociables pour la grille de saisie
    if ( lTLB = TLBSaisie ) and ( pos(GetColName( lTLB, lOldIdx+1), 'E_GENERAL;E_DEBITDEV')>0 )
      then lNbDep := 2
      else lNbDep := 1 ;

    // Déplacement
    if (lOldIdx + lNbDep) <= (lTLB.Items.Count - 1) then
      for i := 1 to lNbDep do
        begin
        lTLB.Items.Exchange(lOldIdx, lOldIdx + 1);
        lTobG.Detail.Move(lOldIdx, lOldIdx + 1);
        Inc( lOldIdx ) ;
        end ;

    lTLB.ItemIndex := lOldIdx ;

    end;

  FBoModifEnCours := True ;

end;

procedure TOF_CPMASQUESAISIE.BUpClick(Sender: TObject);
var lBouton : TToolBarButton97 ;
    lTLB    : TListBox ;
    lOldIdx : integer ;
    lNbDep  : integer ;
    i       : integer ;
    lTobG   : TOB ;
    lInMin  : integer ;
begin
  if not assigned(sender) then Exit ;
  if GetAction = taConsult then Exit ;

  lBouton := TToolBarButton97( Sender ) ;
  if (not lBouton.Enabled) or (not lBouton.visible) then Exit ;

  if Copy(lBouton.Name, 1, 1) = 'S' then
    begin
    lTLB   := TLBSaisie ;
    lTobG  := FTobMasque.Detail[0] ;
    lInMin := 1 ;
    end
  else
    begin
    lTLB   := TLBCompl ;
    lTobG  := FTobMasque.Detail[1] ;
    lInMin := 0 ;
    end ;

  // Test des cellules indissociables pour la grille de saisie
  lOldIdx := lTLB.ItemIndex;
  if ( lTLB = TLBSaisie ) and ( pos( GetColName( lTLB, lOldIdx), 'E_NUMLIGNE;E_GENERAL;E_AUXILIAIRE;E_DEBITDEV;E_CREDITDEV')>0 )
    then Exit ;

  if ( lOldIdx > 0) then
    begin
    // Test des cellules indissociables pour la grille de saisie
    if ( lTLB = TLBSaisie ) and ( pos( GetColName( lTLB, lOldIdx-1), 'E_AUXILIAIRE;E_CREDITDEV')>0 )
      then lNbDep := 2
      else lNbDep := 1 ;

    // Déplacement
    if (lOldIdx - lNbDep) >= lInMin then
      for i := 1 to lNbDep do
        begin
        lTLB.Items.Exchange( lOldIdx, lOldIdx - 1 );
        lTobG.Detail.Move(lOldIdx, lOldIdx - 1);
        Dec( lOldIdx );
        end ;

    // maj index
    lTLB.ItemIndex := lOldIdx ;

    end;

  FBoModifEnCours := True ;

end;

procedure TOF_CPMASQUESAISIE.BAddClick(Sender: TObject);
var lBouton : TToolBarButton97 ;
    lTLB    : TListBox ;
    lIndex  : integer ;
begin
  if not assigned(sender) then Exit ;
  if GetAction = taConsult then Exit ;

  lBouton := TToolBarButton97( Sender ) ;
  if (not lBouton.Enabled) or (not lBouton.visible) then Exit ;

  lTLB := QuelleTLB( lBouton ) ;

  // Aucun champs sélectionné ?
  lIndex := TLBDispo.ItemIndex;
  if lIndex < 0 then Exit ;

//  if EstDansListe( GetColName( TLBDispo, lIndex), lTLB ) then Exit ;
  if not(IsFormule( Tob(TLBDispo.Items.Objects[lIndex]).GetString('NOM') )) and
     EstDansListe(  Tob(TLBDispo.Items.Objects[lIndex]).GetString('NOM'), lTLB ) then Exit ;

  AjouterTOBCol( lTLB, lIndex );
{
  if lIndex < TLBDispo.Items.Count
    then TLBDispo.itemIndex := lIndex
    else TLBDispo.itemIndex := lIndex - 1;
}
  lTLB.ItemIndex := lTLB.Items.Count - 1;
  TLBClick( lTLB );

  FBoModifEnCours := True ;

end;

procedure TOF_CPMASQUESAISIE.BDelClick(Sender: TObject);
var lBouton : TToolBarButton97 ;
    lTLB    : TListBox ;
    lOldIdx : integer ;
    lTobCol : Tob ;
begin
  if not assigned(sender) then Exit ;
  if GetAction = taConsult then Exit ;

  lBouton := TToolBarButton97( Sender ) ;
  if (not lBouton.Enabled) or (not lBouton.visible) then Exit ;

  if Copy(lBouton.Name, 1, 1) = 'S'
    then lTLB := TLBSaisie
    else lTLB := TLBCompl ;

  if lTLB.ItemIndex < 0 then Exit ;

  lOldIdx := lTLB.ItemIndex;
  lTobCol := Tob(lTLB.Items.Objects[lOldIdx]);

  // Test des cellules obligatoire pour la grille de saisie
  if ( lTLB = TLBSaisie ) then
    if lTobCol.GetString('OBLIGATOIRE')='X' then
      begin
      PgiInfo('Vous ne pouvez pas enlever ce champ, il est obligatoire.', Ecran.Caption ) ;
      Exit ;
      end ;

  if IsFormule(lTobCol.GetString('NOM')) then
     if lTLB = TLBSaisie then
        DelFormuleList(lTLB,gFormuleSaisie)
     else            
        DelFormuleList(lTLB,gFormuleCompl);


  // Suppression de l'élement de la TLIstBix
  lTLB.Items.Delete( lOldIdx );

  // Libération de la tob colonne
   FreeAndNil( lTobCol ) ;

  // replacement dans la liste
  if lOldIdx < lTLB.Items.Count
    then lTLB.itemIndex := lOldIdx
    else lTLB.itemIndex := lOldIdx - 1;

  TLBClick( lTLB );

  FBoModifEnCours := True ;

end;

procedure TOF_CPMASQUESAISIE.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var rep : integer ;
begin
  if GetAction = taConsult then Exit ;

  // recup dernière modif
  NextPrevControl( Ecran ) ;

  if FBoModifEnCours then
    begin
    Rep      := PGIAskCancel('Voulez-vous enregistrer les modifications ?',Ecran.Caption) ;
    Case Rep of
      mrYes : CanClose := ClickValide ;
      mrNo  : CanClose := True ;
      else    CanClose := False ;
      end ;
    end
  else
    CanClose := LastError <> 1 ;

  if CanClose then
    TFVierge(Ecran).FormCloseQuery( Sender, CanClose ) ;
    
end;

function TOF_CPMASQUESAISIE.IsVisibleFormule    ( vTLB : TListBox ) : boolean;
var
  Id : string;
begin
  Result := false;
  if vTLB.ItemIndex > 0 then
  begin
     Id := copy(vTLB.Name,4,1);
     if IsFormule(Tob(vTLB.Items.Objects[vTLB.ItemIndex]).GetString('NOM')) then
        Result := true;
  end;

  SetControlVisible('T' + Id + 'FORMULE', Result);
  SetControlVisible(Id + 'FORMULE', Result);
end;

function TOF_CPMASQUESAISIE.IsVisibleCondition  ( vTLB : TListBox ) : boolean;
var
  Id : string;
begin            
  Result := false;
  if vTLB.ItemIndex > 0 then
  begin
     Id := copy(vTLB.Name,4,1);
     if (GetControlVisible(Id + 'READONLY')) and
        (GetCheckBoxState (Id + 'READONLY') = cbUnChecked) then
        Result := true;
  end; 
  SetControlVisible('T' + Id + 'CONDITION', Result);
  SetControlVisible(Id + 'CONDITION', Result);
end;

procedure TOF_CPMASQUESAISIE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var	Vide     : boolean;
    AvecCtrl : boolean;
    AvecAlt  : boolean;
begin

  Vide     := ( Shift = [] ) ;
  AvecCtrl := ( Shift = [ssctrl] ) ;
  AvecAlt  := ( Shift = [ssAlt] ) ;

  Case Key of
    // Validation ligne
    VK_F10    : if Vide then
                  begin
                  Key := 0 ;
                  ClickValide;
                  end ;
    VK_F5     : if (Ecran.ActiveControl is THEdit) then
                begin
                   if (THEdit(Ecran.ActiveControl).Name = 'SFORMULE') or
                      (THEdit(Ecran.ActiveControl).Name = 'CFORMULE') or
                      (THEdit(Ecran.ActiveControl).Name = 'SCONDITION') then  
                      Key := 0;
                      LanceFormuleZone(Ecran.ActiveControl);
                end;
    // Validation ligne
{    VK_RETURN  : if Vide then
                   begin
                   Exit ;
//                   Key := 0 ;
//                   NextControl(Ecran);
//                   Key := VK_TAB  ;
//                   PostMessage( FListe.Handle, WM_KEYDOWN, VK_TAB, 0) ;
                   end ;
}
{
    // nouvelle ligne
    VK_INSERT : if Vide then
                  begin
                  Key := 0 ;
                  ClickInsert;
                  end ;
}
    // raccourci Boutons
    {^A} 65 :   begin
                if not (Ecran.ActiveControl is THMultiValComboBox) then
                  begin
                  Key := 0 ;
                  if AvecCtrl then
                    BAddClick( SAdd )
                  else if AvecAlt then
                    BAddClick( CAdd ) ;
                  end ;
                end ;
    {^E} 69 :   begin
                Key := 0 ;
                if AvecCtrl then
                  BDelClick( SDel )
                else if AvecAlt then
                  BDelClick( CDel ) ;
                end ;
    {^M} 77 :   begin
                Key := 0 ;
                if AvecCtrl then
                  BUpClick( SUp )
                else if AvecAlt then
                  BUpClick( CUp ) ;
                end ;
    {^D} 68 :   begin
                Key := 0 ;
                if AvecCtrl then
                  BDownClick( SDown )
                else if AvecAlt then
                  BDownClick( CDown ) ;
                end ;
(*
    {^J} 74 :   begin
                Key := 0 ;
                end ;
    {^N} 78 :   begin
                Key := 0 ;
                end ;
    {^O} 79 :   begin
                Key := 0 ;
                end ;
    {^S} 83 :   begin
                Key := 0 ;
                end ;

    // Impression
      {^I} 73 : if AvecCtrl then
                  begin
                  Key := 0 ;
                  Imprimer(nil) ;
                  end ;
*)
  end ;

  if Assigned( FOldKeyDown ) then
    FOldKeyDown( Sender, Key, Shift ) ;

end;

procedure TOF_CPMASQUESAISIE.InitVariables;
begin
  FTobMasque := Tob.Create('CMASQUESAISIE', nil, -1) ;
  CInitTobMasque( FTobMasque ) ;

  // Chargements des champs
  FTobChamps := TOB.Create('LISTETABLES', nil, -1) ;
  ChargeChampsDispo ;

  FTobMasqueCrit := nil ;  

end;

procedure TOF_CPMASQUESAISIE.OnAfterFormShow;
begin

  if FInNumero > 0
    then LoadRecord
    else if GetAction<>taConsult
            then NewRecord
            else begin
                 PgiError('Erreur au chargement du masque', Ecran.Caption ) ;
//               TFVierge(Ecran).BFermeClick(nil);
                 end ;       
  gFormuleSaisie := GetNbFormule(TLBSaisie);
  gFormuleCompl  := GetNbFormule(TLBCompl);
end;

procedure TOF_CPMASQUESAISIE.ComboTableChanged(Sender: TObject);
begin
  // MAJ de la liste des champs dispo
  SetListeChampsDispo ;

  { BVE On autorise les champs TIERS et GENERAUX dans la grille de saisie
        Et les champs spécifiques dans la grille de complément 
  // MAJ accès au boutons de sélections des champs dans les grilles
  SetControlEnabled('SADD', (ComboTable.Value = 'ECRITURE') or (ComboTable.Value = 'SPECIFIQUES') ) ;
  SetControlEnabled('SDEL', (ComboTable.Value = 'ECRITURE') or (ComboTable.Value = 'SPECIFIQUES') ) ;

  SetControlEnabled('CADD', (ComboTable.Value <> 'SPECIFIQUES') ) ;
  SetControlEnabled('CDEL', (ComboTable.Value <> 'SPECIFIQUES') ) ;     }

end;


procedure TOF_CPMASQUESAISIE.NewRecord ;
var lStNum   : string ;
    i        : integer ;
begin

  // choix du masque de référence
  lStNum := Choisir( 'Sélection d''un modèle de masque','CMASQUESAISIE','CMS_LIBELLE','CMS_NUMERO','CMS_TYPE="CEG"','', False, False, 0) ;

  // Chargement du masque de référence
  if lStNum<>''
    then CChargeMasque( StrToInt( lStNum ), FTobMasque, 'CEG' )
    else CChargeMasque( GetParamSocSecur( 'SO_CPMASQUEDEFAUT', 2 ) , FTobMasque, 'CEG' ) ;

  // Reset des zones
  FTobMasque.PutValue('CMS_TYPE',     FStType ) ;
  FTobMasque.PutValue('CMS_NUMERO',   0 ) ;
  FTobMasque.PutValue('CMS_LIBELLE',  '' ) ;
  FTobMasque.PutValue('CMS_CRITERE1', '' ) ;
  FTobMasque.PutValue('CMS_CRITERE2', '' ) ;
  FTobMasque.PutValue('CMS_CRITERE3', '' ) ;

  // MAJ des colonnes
  for i := 0 to FTobMasque.Detail[0].Detail.count - 1 do
    FTobMasque.Detail[0].Detail[ i ].PutValue('OBLIGATOIRE', 'X' ) ;

  // Affichage
  AfficheMasque ;

  FBoModifEnCours := True ;
  FTobMasqueCrit  := nil ;
  FTobMasque.SetAllModifie(True);
  SetControlEnabled('BCopier', True);

end;

function TOF_CPMASQUESAISIE.CreerTobCol( vStNom,vStLib,vStTipe,vStTable: string; vParent: Tob): Tob;
begin

  result := Tob.Create( 'COLONNE',    vParent,   -1 ) ;

  // champs pour clé XML
  result.AddChampSupValeur( 'NOM',         vStNom ) ;
  result.AddChampSupValeur( 'LIBELLE',     vStLib )  ;
  result.AddChampSupValeur( 'TYPE',        vStTipe ) ;
  result.AddChampSupValeur( 'INDEX',       0 )      ;
  result.AddChampSupValeur( 'TITRE',       vStLib )  ;
  result.AddChampSupValeur( 'LARGEUR',     101 ) ;
  result.AddChampSupValeur( 'AVECLIBELLE', '-' ) ;
  result.AddChampSupValeur( 'READONLY',    '-' ) ;
  result.AddChampSupValeur( 'OBLIGATOIRE', '-' ) ;
  result.AddChampSupValeur( 'FORMAT',      '' ) ;
  result.AddChampSupValeur( 'DATATYPE',    '' ) ;

  // attributs
  if (vStTable = 'ECRITURE') and not CEstColEditable( vStNom ) then
    result.PutValue('READONLY', 'X') ;

  result.PutValue( 'LARGEUR', GetLargeurDefaut( result ) ) ;    

end;

function TOF_CPMASQUESAISIE.EstDansListe( vStNom: string; vTLB: TListBox): boolean;
var
  i : integer;
begin 
  result := false;
  { BVE On ne sert plus du libelle mais du nom du champ
  result := vTLB.Items.IndexOf(vStNom) >= 0 ;           }
  for i := 0 to vTLB.Items.Count - 1 do
  begin
     if Tob(vTLB.Items.Objects[i]).GetString('NOM') = vStNom then
     begin
        result := true;
        break;
     end;
  end;
end;

procedure TOF_CPMASQUESAISIE.TLBClick ( Sender : TObject ) ;
var lTLIdx  : integer ;
    lTobCol : Tob ;
    lTLB    : TListBox ;
    lStPref : string ;
begin

  if not Assigned( Sender ) then Exit ;
  if GetAction = taConsult then Exit ;


  lTLB   := TListBox(Sender) ;
  lTLIdx := lTLB.ItemIndex ;

  if lTLB = TLBSaisie
    then lStPref := 'S'
    else lStPref := 'C' ;

  if lTLIdx >= 0
    then lTobCol := Tob( lTLB.Items.Objects[lTLIdx] )
    else lTobCol := nil ;

  AfficheParamCol( lStPref, lTobCol ) ;

  IsVisibleFormule(lTLB);
  IsVisibleCondition(lTLB);
end;


function TOF_CPMASQUESAISIE.QuelleTLB( vControl : TControl ) : TListBox ;
begin
    if vControl.Parent.Name = 'GBSAISIE'
      then result := TLBSaisie
      else result := TLBCompl ;
end;

procedure TOF_CPMASQUESAISIE.AjouterTobCol(vTLBDest: TListBox; vDispoIdx: integer);
var vParent : Tob ;
    lTobCol : Tob ;
    lStText : string ;
begin

  if vTLBDest.Name = 'TLBSAISIE'
    then vParent := FTobMasque.Detail[0]
    else vParent := FTobMasque.Detail[1] ;

  lTobCol := DupliquerTobCol( Tob( TLBDispo.Items.Objects[vDispoIdx] ), vParent ) ;

  if IsFormule(lTobCol.GetString('NOM')) then
  begin
     if vTLBDest.Name = 'TLBSAISIE' then
     begin
        Inc(gFormuleSaisie);
        lTobCol.SetString('NOM','FORMULE' + IntToStr(gFormuleSaisie));
     end
     else
     begin
        Inc(gFormuleCompl);
        lTobCol.SetString('NOM','FORMULE' + IntToStr(gFormuleCompl));
     end;
  end;

  if Affichage.ItemIndex = 0
    then lStText := lTobCol.GetString('NOM')
    else lStText := lTobCol.GetString('LIBELLE') ;

  lTobCol.PutValue( 'INDEX', vTLBDest.Items.Count + 1 ) ;

  vTLBDest.AddItem( lStText, lTobCol );

end;


procedure TOF_CPMASQUESAISIE.RemplirGrille( vTLB : TListBox ; vTobListeCol : Tob ) ;
var lTobCol    : Tob ;
    lStText    : string ;
    i          : integer ;
begin

  vTLB.Items.Clear ;

  For i := 0 to vTobListeCol.Detail.count-1 do
    begin
    lTobCol := Tob( vTobListeCol.Detail[ i ] ) ;
    if Affichage.ItemIndex = 0
      then lStText := lTobCol.GetString('NOM')
      else lStText := lTobCol.GetString('LIBELLE') ;
    vTLB.Items.AddObject( lStText, lTobCol ) ;
    end ;

end;

procedure TOF_CPMASQUESAISIE.SetListeChampsDispo ;
var lTobTable  : Tob ;
    lStTable   : string ;
begin

  TLBDispo.Items.clear ;

  lStTable  := ComboTable.Value ;
  lTobTable := FTobChamps.FindFirst( ['CO_ABREGE'], [ lStTable], True ) ;
  if not Assigned( lTobTable ) then Exit ;

  RemplirGrille( TLBDispo, lTobTable ) ;

end;

procedure TOF_CPMASQUESAISIE.ChargeChampsDispo;
var i,j        : integer ;
    lNumTable  : integer ;
    lNumChamps : integer ;
    lStTable   : string ;
    lTobTable  : Tob ;
    lTobCol    : Tob ;
    FTobSpecif : Tob ;
begin
  FTobChamps.ClearDetail ;

  // On récupère la liste des tables
  FTobChamps.LoadDetailDBFromSQL('COMMUN', 'SELECT * FROM COMMUN WHERE CO_TYPE="CMB" ORDER BY CO_CODE' ) ;

  for i := 0 to FTobChamps.Detail.Count - 1 do
    begin
    lTobTable := FTobChamps.Detail[ i ] ;
    lStTable  := lTobTable.GetString('CO_ABREGE') ;

    if lStTable = 'SPECIFIQUES' then
      // Champs spécifiques ( commme la saisie mono-section par ex )
      begin
      FTobSpecif := Tob.Create('SPECIFIQUES', nil, -1) ;
      FTobSpecif.LoadDetailDBFromSQL('COMMUN', 'SELECT * FROM COMMUN WHERE CO_TYPE="' + lTobTable.GetString('CO_LIBRE') + '" ORDER BY CO_CODE') ;
      for j := 0 to FTobSpecif.Detail.count-1 do
        begin
        lTobCol          := CreerTobCol( FTobSpecif.Detail[j].GetString('CO_ABREGE'),
                                         FTobSpecif.Detail[j].GetString('CO_LIBELLE'),
                                         'SPECIF',
                                         lStTable,
                                         lTobTable ) ;
        lTobCol.PutValue('DATATYPE', FTobSpecif.Detail[j].GetString('CO_LIBRE') ) ;
        SetTitreColSpecif( lTobCol ) ;
        end ;
      FreeAndNil( FTobSpecif ) ;
      end
    else
      // Récup champs du MCD avec control à "L"
      begin
      lNumTable := TableToNum( lStTable );
      if (lNumTable > 0) and (lNumTable<High(V_PGI.DEChamps)) then
        begin
        {$IFDEF EAGLCLIENT}
        lTobCol := Tob.Create( lStTable, nil, -1 ) ;
        FreeAndNil( lTobCol ) ;
        {$ENDIF EAGLCLIENT}
        for lNumChamps := 1 to High( V_PGI.DEChamps[lNumTable] ) do
          begin
          if pos( 'L', V_PGI.DEChamps[lNumTable, lNumChamps ].Control ) > 0 then
            if V_PGI.DEChamps[lNumTable, lNumChamps ].Tipe <> 'BLOB' then
              CreerTobCol( V_PGI.DEChamps[lNumTable, lNumChamps ].Nom,
                           V_PGI.DEChamps[lNumTable, lNumChamps ].Libelle,
                           V_PGI.DEChamps[lNumTable, lNumChamps ].Tipe,
                           lStTable,
                           lTobTable ) ;
          end ;
        end ;
      end ;

    end ;

end;
procedure TOF_CPMASQUESAISIE.ChangeCondition ( Sender : TObject ) ;
var
  vide : boolean;
begin
  vide := (Trim(GetControlText(THEdit(Sender).Name)) = '');
  SetControlEnabled('SREADONLY',vide);
end;

procedure TOF_CPMASQUESAISIE.AffichageClick(Sender: TObject);
begin
  RaffraichirCombo ;
  RaffraichirTLB( TLBDispo, True ) ;
  RaffraichirTLB( TLBSaisie ) ;
  RaffraichirTLB( TLBCompl  ) ;
end;

procedure TOF_CPMASQUESAISIE.RaffraichirTLB(vTLB: TListBox; vBoAvecTri: Boolean);
var i       : integer ;
    lTobCol : Tob ;
    lStText : string ;
    lIdx    : integer ;
begin
  // sauvegarde index
  lIdx := vTLB.ItemIndex ;

  // Modif des libellés
  vTLB.Sorted := False ;
  for i := 0 to vTLB.Items.Count - 1 do
    begin
    lTobCol := Tob( vTLB.Items.Objects[i] ) ;
    if Affichage.ItemIndex = 0
      then lStText := lTobCol.GetString('NOM')
      else lStText := lTobCol.GetString('LIBELLE') ;

    vTLB.Items[i] := lStText ;
    end;

  // replacement
  vTLB.ItemIndex := lIdx ;

  // Tri
  if vBoAvecTri then
    vTLB.Sorted := True ;
end ;


function TOF_CPMASQUESAISIE.DupliquerTobCol( vTobCol: Tob ; vParent: Tob ) : Tob ;
begin
  result := Tob.Create( 'COLONNE',    vParent,   -1 ) ;
  result.Dupliquer( vTobCol, False, True);
end;

procedure TOF_CPMASQUESAISIE.RemplirComboTable;
var i         : integer ;
    lTobTable : Tob ;
    lStText   : string ;
begin

  ComboTable.Items.Clear ;

  for i := 0 to FTobChamps.Detail.count - 1 do
    begin
    lTobTable := FTobChamps.Detail[ i ] ;

    if Affichage.ItemIndex = 0
      then lStText := lTobTable.GetString('CO_ABREGE')
      else lStText := lTobTable.GetString('CO_LIBELLE') ;

    ComboTable.Items.AddObject( lStText, lTobTable) ;
    ComboTable.Values.Add( lTobTable.GetString('CO_ABREGE') ) ;
    end ;

end;

function TOF_CPMASQUESAISIE.GetAction: TActionFiche;
begin
  result := TFVierge( Ecran ).TypeAction ;
end;

procedure TOF_CPMASQUESAISIE.ParamColChanged(Sender: TObject);
var lTLB      : TListBox ;
    lTobCol   : Tob ;
    lControl  : TControl ;
    lStChp    : string ;
    lStNewV   : string ;
begin
  if not Assigned(Sender) then Exit ;
  if GetAction = taConsult then Exit ;
  lControl := TControl( Sender ) ;
  if (not lControl.Enabled) or (not lControl.visible) then Exit ;

  lTLB    := QuelleTLB( lControl ) ;
  if lTLB.ItemIndex < 0 then Exit ;

  lTobCol := Tob( lTLB.Items.Objects[ lTLB.ItemIndex ] ) ;

  lStChp  := Copy( lControl.Name, 2, length( lControl.Name ) ) ;
  lStNewV := GetControlText( lControl.Name ) ;

  if not(lTobCol.FieldExists( lStChp )) then
  begin
    lTobCol.AddChampSupValeur( lStChp, lStNewV ); 
  end
  else if lTobCol.GetString( lStChp ) <> lStNewV then
  begin
    lTobCol.PutValue( Copy(lControl.Name, 2, length( lControl.Name ) ), GetControlText(lControl.Name) ) ;
    FBoModifEnCours := True ;
  end ;

  IsVisibleFormule(lTLB);
  IsVisibleCondition(lTLB);  
end;


function TOF_CPMASQUESAISIE.ClickValide : boolean;
begin
  result := True ;
  if GetAction = taConsult then Exit ;

  // recup dernière modif
  NextPrevControl( Ecran ) ;

  if not FBoModifEnCours then Exit ;

  // MAJ de l'objet
  UpdateTobMasque ;

  // Tests
  result := TestRecord ;

  // Enregistrement
  if result then
    if transactions( SaveRecord, 1 ) <> oeOk then
      begin
      result := False ;
      PgiBox( TraduireMemoire( 'Une erreur est survenue lors de l''enregistrement !') , Ecran.Caption ) ; 
      end ;

  SetControlEnabled('BCOPIER', False) ;

  // fermeture de la fenêtre en création
  if result and ( GetAction = taCreat ) then
      SendMessage(Ecran.Handle, WM_CLOSE, 0, 0);


end;

procedure TOF_CPMASQUESAISIE.AfficheParamCol( vStPref : string ; vTobCol: Tob );
begin

  if vTobCol=nil then
  begin
    SetControlText(     vStPref + 'TITRE',       '' ) ;
    SetControlText(     vStPref + 'LARGEUR',     '0' ) ;
    SetControlVisible(  vStPref + 'AVECLIBELLE', False ) ;
    SetControlText(     vStPref + 'AVECLIBELLE', '-' ) ;
    if (vStPref = 'S') then
    begin
      SetControlVisible( 'SREADONLY', False ) ;
      SetControlText(    'SREADONLY',    '-' ) ;
      SetControlText(   vStPref + 'CONDITION', '');
    end ;
    SetControlText(     vStPref + 'FORMULE', ''); 
  end
  else
  begin
    SetControlText( vStPref + 'TITRE',   vTobCol.GetString('TITRE') ) ;
    SetControlText( vStPref + 'LARGEUR', vTobCol.GetString('LARGEUR') ) ;
    if vTobCol.GetString('TYPE') = 'COMBO' then
    begin
      SetControlVisible( vStPref + 'AVECLIBELLE', True ) ;
      SetControlText(    vStPref + 'AVECLIBELLE', vTobCol.GetString('AVECLIBELLE') ) ;
    end
    else
    begin
      SetControlVisible( vStPref + 'AVECLIBELLE', False ) ;
      SetControlText(    vStPref + 'AVECLIBELLE', '-' ) ;
    end ;

    if (vStPref = 'S') then
    begin
      if CEstColEditable( vTobCol.GetString('NOM') ) then
      begin
        SetControlVisible( 'SREADONLY',  True ) ;
        SetControlText(    'SREADONLY',  vTobCol.GetString('READONLY') ) ;
      end
      else
      begin
        SetControlVisible( 'SREADONLY', False ) ;
        SetControlText(    'SREADONLY',    '-' ) ;
      end ;
      if vTobCol.FieldExists('CONDITION') then
         SetControlText(   vStPref + 'CONDITION', vTobCol.GetString('CONDITION'))
      else
         SetControlText(   vStPref + 'CONDITION', '');
    end ;
    if vTobCol.FieldExists('FORMULE') then
       SetControlText(     vStPref + 'FORMULE', vTobCol.GetString('FORMULE'))
    else
       SetControlText(     vStPref + 'FORMULE', '');
  end ;
end;

procedure TOF_CPMASQUESAISIE.AfficheMasque;
{$IFDEF SAISIEDOC}
var lTobG : Tob ;
{$ENDIF SAISIEDOC}
begin
  // Entete
  FTobMasque.PutEcran( Ecran, TWinControl(GetControl('GBENTETE', True)) ) ;

{$IFDEF SAISIEDOC}
  // ------------------------
  // Affichage des paramètres
  // ------------------------
  lTobG := FTobMasque.FindFirst( ['TYPE'], ['PARAMETRES'], True ) ;
  if not Assigned(lTobG) then
    lTobG := CInitTobMasqueParam( FTobMasque ) ;

  lTobG.PutEcran(Ecran, TWinControl(GetControl('PPARAM')) ) ;
{$ENDIF SAISIEDOC}

  // Affichage de la liste de SAISIE
  RemplirGrille( TLBSaisie, FTobMasque.Detail[0] ) ;
  TLBSaisie.ItemIndex := 1 ;
  TLBClick( TLBSaisie ) ;

  // Affichage de la liste de COMPLEMENTS
  RemplirGrille( TLBCompl, FTobMasque.Detail[1] ) ;
  TLBCompl.ItemIndex := 1 ;
  TLBClick( TLBCompl ) ;

end;

procedure TOF_CPMASQUESAISIE.TLBDblClick(Sender: TObject);
begin
  if GetAction = taConsult then Exit ;
  if not Assigned(Sender) then Exit ;

  if Sender = TLBDispo then
    begin
    if SAdd.Enabled
      then BAddClick( SAdd )
    else if CAdd.Enabled
      then BAddClick( CAdd ) ;   
    end
  else if Sender = TLBSaisie then
    BDelClick( SDel )
  else if Sender = TLBCompl then
    BDelClick( CDel ) ;

end;

procedure TOF_CPMASQUESAISIE.TLBDrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var lTLB    : TListBox ;
    lTobCol : Tob ;
begin

  lTLB    := TListBox( Control ) ;
  lTobCol := Tob(lTLB.Items.Objects[Index]);

  // Test des cellules obligatoire pour la grille de saisie
  if lTobCol.GetString('OBLIGATOIRE')='X' then
    lTLB.Canvas.Font.Style := [ fsBold ]
  else
    lTLB.Canvas.Font.Style := [  ] ;   //TFontStyles
//      Canvas.Font.Color := clWindowText ;


end;

procedure TOF_CPMASQUESAISIE.CalculIndex;
var lTob : Tob ;
    i    : integer ;
begin

  for i := 1 to TLBSaisie.Items.count do
    begin
    lTob := TOb(TLBSaisie.items.objects[i-1]) ;
    lTob.PutValue('INDEX', i ) ;
    end ;

  for i := 1 to TLBCompl.Items.count do
    begin
    lTob := TOb(TLBCompl.items.objects[i-1]) ;
    lTob.PutValue('INDEX', i ) ;
    end ;

end;

function  TOF_CPMASQUESAISIE.IsFormule ( vStNom : String ) : boolean;
begin
  Result := ( Pos('FORMULE',vStNom) = 1 );
end;

function  TOF_CPMASQUESAISIE.GetNbFormule ( vTLB : TListBox ) : integer;
var
  i : integer;
begin
  Result := 0;
  for i := 0 to vTLB.Items.Count - 1 do
     if IsFormule(Tob(vTLB.Items.Objects[i]).GetString('NOM')) then
        Inc(Result);
end;

procedure TOF_CPMASQUESAISIE.DelFormuleList ( vTLB : TListBox ; var gFormule : integer) ;
var
  id   : integer;
  id2  : integer;
  nom  : string;
  nom2 : string;
  i    : integer;
begin
  nom := Tob(vTLB.Items.Objects[vTLB.ItemIndex]).GetString('NOM');
  id := StrToInt(copy(nom,8,length(nom)-7));

  if id <> gFormule then
  begin
     for i := 0 to vTLB.Items.Count - 1 do
     begin                
        nom2 := Tob(vTLB.Items.Objects[i]).GetString('NOM');
        if IsFormule(nom2) then
        begin
           id2 := StrToInt(copy(nom2,8,length(nom2)-7));
           if id2 > id then
           begin
              Dec(id2);
              Tob(vTLB.Items.Objects[i]).SetString('NOM','FORMULE' + IntToStr(id2));
              if Affichage.ItemIndex = 0 then
                 vTLB.Items[i] := 'FORMULE' + IntToStr(id2);
           end;
        end;
     end;
  end;                            
  Dec(gFormule);
end;

procedure TOF_CPMASQUESAISIE.RaffraichirCombo;
var i         : integer ;
    lTobTable : Tob ;
    lStText   : string ;
    lIdx    : integer ;
begin

  // sauvegarde index
  lIdx := ComboTable.ItemIndex ;

  for i := 0 to ComboTable.Items.count - 1 do
    begin
    lTobTable := Tob( ComboTable.Items.Objects[ i ] ) ;

    if Affichage.ItemIndex = 0
      then lStText := lTobTable.GetString('CO_ABREGE')
      else lStText := lTobTable.GetString('CO_LIBELLE') ;

    ComboTable.Items[i] := lStText ;

    end ;

  ComboTable.refresh ;
  ComboTable.ItemIndex := lIdx ;

end;

{$IFNDEF CCSTD }
procedure TOF_CPMASQUESAISIE.InitPageTest;
var lPiece   : TPieceCompta ;
    lGrid    : THGrid ;
begin

  if not assigned( TSP ) then
    begin
    lGrid  := THGrid( GetControl( 'FLISTE', TRue ) ) ;
    lPiece := TPieceCompta.CreerPiece ;
    TSP := TSaisiePiece.CreerPourMasque( FTobMasque, lGrid, lPiece, Ecran);
    lGrid  := THGrid( GetControl( 'FCOMPL', TRue ) ) ;
    TSP.SetGrilleComplement( lGrid );
    end ;

  if GetAction <> taConsult then
    begin
    UpdateTobMasque ;
    TSP.UpdateMasqueSaisie ;
    TSP.AfficheLignes ;
    end ;

end;
{$ENDIF CCSTD }

procedure TOF_CPMASQUESAISIE.UpdateTobMasque ;
var lTobTemp : Tob ;
    lStream  : TStringStream ;
    i        : integer ;
{$IFDEF SAISIEDOC}
    lTobG    : Tob ;
{$ENDIF SAISIEDOC}

begin

  CalculIndex ;

{$IFDEF SAISIEDOC}
  // ---------------------------
  // Récupération des paramètres
  // ---------------------------
  lTobG := FTobMasque.FindFirst( ['TYPE'], ['PARAMETRES'], True ) ;
  if not Assigned(lTobG) then
    lTobG := CInitTobMasqueParam( FTobMasque ) ;

  lTobG.GetEcran( Ecran, TWinControl(GetControl('PPARAM')) );
{$ENDIF SAISIEDOC}

  // ---------------------------------
  // Préparation de la tob description
  // ---------------------------------
  lTobTemp := Tob.Create('DESCRIPTION', nil, -1 );
  for i := FTobMasque.Detail.count-1 downto 0 do
    FTobMasque.Detail[i].ChangeParent(lTobTemp, 0) ;

  lTobTemp.AddChampSupValeur('VERSION',   _VersionXML );
  lTobTemp.AddChampSupValeur('DATEMODIF', DateToStr( Now ) );

  // ------------------------
  // Récupération de l'entête
  // ------------------------
  FTobMasque.PutValue('CMS_LIBELLE',  GetControlText('CMS_LIBELLE') ) ;
  FTobMasque.PutValue('CMS_CRITERE1', GetControlText('CMS_CRITERE1') ) ;
  FTobMasque.PutValue('CMS_CRITERE2', GetControlText('CMS_CRITERE2') ) ;
  FTobMasque.PutValue('CMS_CRITERE3', GetControlText('CMS_CRITERE3') ) ;

  // ---------------------
  // MAJ du descriptif XML
  // ---------------------
  lStream := TStringStream.Create( '' ) ;
  lTobTemp.SaveToXmlStream( lStream, False, True );
  FTobMasque.SetString('CMS_DESCRIPTION', lStream.DataString );

  // -------------------------
  // ReMAJ des TOB description
  // -------------------------
  for i := lTobTemp.Detail.count-1 downto 0 do
    lTobTemp.Detail[i].ChangeParent(FTobMasque, 0) ;

  // ------------------
  // Libération mémoire
  // ------------------
  if Assigned( lTobTemp ) then
    FreeAndNil( lTobTemp ) ;

  if Assigned( lStream ) then
    FreeAndNil( lStream ) ;

end;

procedure TOF_CPMASQUESAISIE.PCMasqueChange(Sender: TObject);
begin

  if (PCMasque.ActivePage.Name <> 'PTEST') then Exit ;

{$IFNDEF CCSTD }
  InitPageTest ;
{$ENDIF CCSTD }

end;

procedure TOF_CPMASQUESAISIE.SetTitreColSpecif(vTobCol: TOB);
var lStAxe : string ;
    lQAxe  : TQuery ;
begin
  if pos( 'SECTION', vTobCol.GetString('NOM') ) > 0 then
    begin
    lStAxe := Copy( vTobCol.GetString('NOM'), 8, 2 ) ;
    lQAxe := OpenSQL('SELECT X_LIBELLE FROM AXE WHERE X_AXE = "' + lStAxe + '"', True ) ;
    if not lQAxe.Eof then
      vTobCol.PutValue('TITRE', lQAxe.FindField('X_LIBELLE').AsString ) ;
    Ferme( lQAxe ) ;
    end ;
end;

procedure TOF_CPMASQUESAISIE.EnteteChanged(Sender: TObject);
var lStChp : string ;
begin
  if not Assigned(Sender) then Exit ;

  lStChp := TControl(Sender).Name ;

  if GetControlText( lStChp ) <> FTobMasque.GetString( lStChp ) then
    FBoModifEnCours := True ;

end;

procedure TOF_CPMASQUESAISIE.LanceFormuleZone ( Sender : TObject ) ;
var
  tv : TTreeView;
begin
   if not(Sender is THEdit) then Exit;
   tv := TTreeView.Create(ecran);
   try
      tv.Name := 'CHAMPSDISPO';
      tv.Parent := ecran;
      if copy(THEdit(Sender).Name,1,1) = 'S' then
         ChargeTreeView(tv,TLBSaisie)
      else if copy(THEdit(Sender).Name,1,1) = 'C' then
         ChargeTreeView(tv,TLBCompl);
      SetControlText(THEdit(Sender).Name,EditeFormule(tv,GetControlText(THEdit(Sender).Name),true));
   finally
      DechargeTreeView(tv);
      tv.Free;
   end;
end;

procedure TOF_CPMASQUESAISIE.ChargeTreeView ( Tree : TTreeView ; vGrille : TListBox );
var
  OEdt     : TOedt;
  Index    : integer;
  i,j      : integer;
  p,n,o    : TTreeNode;
  TobTable : TOB;
  TobChamp : TOB;
  TobFormule : TOB;
begin
  Tree.Items.Clear;
  p := Tree.Items.Add(nil, ' Champs disponibles');
  for Index := 0 to FTobChamps.Detail.Count - 1 do
  begin
    // On est sur les différentes tables
    TobTable := FTobChamps.Detail[Index];
    if (TobTable.GetString('CO_ABREGE') = '') then continue;
    OEdt := TOedt.Create;
    OEdt.Quoi := oeChamp;
    OEdt.Code := TobTable.GetString('CO_ABREGE');
    OEdt.Libelle := TobTable.GetString('CO_LIBELLE');
    OEdt.TypeChamp := ChampToType(OEdt.Code);
    if (OEdt.Libelle = '') or (OEdt.Libelle = '??') then OEdt.Libelle := OEdt.Code;
    OEdt.Controle := nil;
    n := Tree.Items.AddChildObject(p, OEdt.Libelle, OEdt);
    n.ImageIndex := 2;
    n.SelectedIndex := n.ImageIndex;
    for i := 0 to TobTable.Detail.Count - 1 do
    begin
      // On est sur les champs de la table
      TobChamp := TobTable.Detail[i];
      if (TobTable.GetString('CO_ABREGE') = 'SPECIFIQUES') and
         (TobChamp.GetString('NOM') = 'FORMULE') then
      begin
         // Gestion des formules à part.
         for j := 1 to GetNbFormuleGrille(vGrille) do
         begin
            TobFormule := GetNiemeFormuleGrille(vGrille,j);
            OEdt := TOedt.Create;
            OEdt.Quoi := oeChamp;
            OEdt.Code := TobFormule.GetString('NOM');
            OEdt.Libelle := TobFormule.GetString('TITRE');
            OEdt.TypeChamp := ChampToType(OEdt.Code);
            if (OEdt.Libelle = '') or (OEdt.Libelle = '??') then OEdt.Libelle := OEdt.Code;
            OEdt.Controle := nil;
            o := Tree.Items.AddChildObject(n, OEdt.Libelle, OEdt);
            o.ImageIndex := 2;
            o.SelectedIndex := o.ImageIndex;
         end; 
      end
      else
      begin
         OEdt := TOedt.Create;
         OEdt.Quoi := oeChamp;
         OEdt.Code := TobChamp.GetString('NOM');
         OEdt.Libelle := TobChamp.GetString('LIBELLE');
         OEdt.TypeChamp := ChampToType(OEdt.Code);
         if (OEdt.Libelle = '') or (OEdt.Libelle = '??') then OEdt.Libelle := OEdt.Code;
         OEdt.Controle := nil;
         o := Tree.Items.AddChildObject(n, OEdt.Libelle, OEdt);
         o.ImageIndex := 2;
         o.SelectedIndex := o.ImageIndex;
      end;
    end;
  end;
end;

function TOF_CPMASQUESAISIE.GetNbFormuleGrille ( vTLB : TListBox ) : integer;
var
  i : integer;
begin
  Result := 0;
  for i := 0 to vTLB.Items.Count - 1 do
  begin
     if IsFormule(Tob(vTLB.Items.Objects[i]).GetString('NOM')) then
        Inc(Result);
  end;
end;

function TOF_CPMASQUESAISIE.GetNiemeFormuleGrille ( vTLB : TListBox ; n : integer ) : TOB;
var
  i,j : integer;
begin
  Result := nil;
  j := 0;
  for i := 0 to vTLB.Items.Count - 1 do
  begin
     if IsFormule(Tob(vTLB.Items.Objects[i]).GetString('NOM')) then
     begin
        Inc(j);
        if (j = n) then
        begin
           Result := Tob(vTLB.Items.Objects[i]);
           Exit;
        end;
     end;
  end;
end;

procedure TOF_CPMASQUESAISIE.DechargeTreeView ( Tree : TTreeView );
var
  Index: integer;   
  OEdt: TOedt;
begin
  for Index := 1 to Tree.Items.Count - 1 do
  begin
    OEdt := TOedt(Tree.Items[Index].Data);
    if OEdt <> nil then OEdt.Free;
  end;
end;

function TOF_CPMASQUESAISIE.GetColName( vTLB : TListBox ; vIdx : integer ) : string ;
var lTobCol : Tob ;
begin

  result := '' ;

  if vIdx < 0 then Exit ;
  if vIdx > (vTLB.Items.Count - 1) then Exit ;

  lTobCol := TOB( vTLB.Items.Objects[ vIdx ] ) ;
  result  := lTObCol.GetString('NOM') ;

end;

procedure TOF_CPMASQUESAISIE.BCopierClick(Sender: TObject);
var lStNum   : string ;
begin

  if GetAction <> taCreat then Exit ;

  // choix du masque de référence
  lStNum := Choisir( 'Sélection d''un masque à dupliquer','CMASQUESAISIE','CMS_LIBELLE','CMS_NUMERO','CMS_TYPE="SAI"','', False, False, 0) ;

  // Chargement du masque de référence
  if lStNum<>''
    then CChargeMasque( StrToInt( lStNum ), FTobMasque )
    else Exit ;

  // Reset des zones
  FTobMasque.PutValue('CMS_TYPE',     FStType ) ;
  FTobMasque.PutValue('CMS_NUMERO',   0 ) ;
  FTobMasque.PutValue('CMS_LIBELLE',  '' ) ;
  FTobMasque.PutValue('CMS_CRITERE1', '' ) ;
  FTobMasque.PutValue('CMS_CRITERE2', '' ) ;
  FTobMasque.PutValue('CMS_CRITERE3', '' ) ;
  FTobMasque.SetAllModifie(True);

  // Affichage
  AfficheMasque ;

  FBoModifEnCours := True ;
  FTobMasqueCrit  := nil ;

end;

function TOF_CPMASQUESAISIE.GetLargeurDefaut(vTobCol: TOB): integer;
var lStType : string ;
begin

  lStType := vTobCol.GetString('TYPE') ;

  if lStType = 'COMBO' then
    result := 80
  else if lStType = 'DOUBLE' then
    result := 100
  else if lStType = 'DATE' then
    result := 80
  else if lStType = 'INTEGER' then
    result := 50
  else if lStType = 'VARCHAR(35)' then
    result := 160
  else if lStType = 'VARCHAR(17)' then
    result := 80
  else if lStType = 'VARCHAR(1)' then
    result := 20
  else if lStType = 'CHAR(1)' then
    result := 20
  else if lStType = 'BOOLEAN' then
    result := 20
  else if lStType = 'SPECIF' then
    result := 80
  else result := 80 ;

end;

{$IFDEF SAISIEDOC}
procedure TOF_CPMASQUESAISIE.ParamChanged(Sender: TObject);
var lStChp    : string ;
    lTobParam : Tob ;
begin
  if not Assigned(Sender) then Exit ;

  lTobParam := FTobMasque.FindFirst( ['TYPE'], ['PARAMETRES'], True ) ;
  if not Assigned(lTobParam) then Exit ;

  lStChp := TControl(Sender).Name ;
  if GetControlText( lStChp ) <> lTobParam.GetString( lStChp ) then
    FBoModifEnCours := True ;

end;
{$ENDIF SAISIEDOC}

Initialization
  registerclasses ( [ TOF_CPMASQUESAISIE ] ) ;
end.
