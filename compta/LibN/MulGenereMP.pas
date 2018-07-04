unit MulGenereMP;

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
    Mul,
    StdCtrls,
    Hcompte,
    Mask,
    Hctrls,
    hmsgbox,
    Menus,
    DB,
    DBTables,
    Hqry,
    Grids,
    DBGrids,
    ExtCtrls,
    ComCtrls,
    Buttons,
    Hent1,
    Ent1,
    Saisie,
    SaisUtil,
    SaisComm,
    HRichEdt,
    HSysMenu,
    HDB,
    HTB97,
    ColMemo,
    HPanel,
    UiUtil,
    EncUtil,
    LettUtil,
    GenereMP,
    UTOB,
    {FE_MAIN,}
    UtilPGI,
    FileCtrl,
    ParamSoc,
    HRichOLE,
    MulSMPUtil,
    Filtre,
    GenerMP,
    TofVerifRib,
    uLibWindows, // pour CVireLigne
{$IFNDEF IMP}
    SaisBor,
{$ENDIF}
    LetRegul
    ;

Const GereEscompte : boolean = True ;

Procedure GenereSuiviMP ( smp : TSuiviMP ) ;

type
  TFMulGenereMP = class(TFMul)
    HM: THMsgBox;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    TE_DEBIT: TLabel;
    TE_DEBIT_: TLabel;
    TE_CREDIT: TLabel;
    TE_CREDIT_: TLabel;
    TE_ETABLISSEMENT: THLabel;
    E_DEBIT: THCritMaskEdit;
    E_DEBIT_: THCritMaskEdit;
    E_CREDIT: THCritMaskEdit;
    E_CREDIT_: THCritMaskEdit;
    E_ETABLISSEMENT: THValComboBox;
    PEcritures: TTabSheet;
    Bevel6: TBevel;
    TE_JOURNAL: THLabel;
    TE_NATUREPIECE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    TE_EXERCICE: THLabel;
    TE_DATEECHEANCE: THLabel;
    TE_DATEECHEANCE2: THLabel;
    E_EXERCICE: THValComboBox;
    E_JOURNAL: THValComboBox;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    XX_WHEREAN: TEdit;
    E_QUALIFPIECE: THCritMaskEdit;
    E_ECHE: THCritMaskEdit;
    E_ETATLETTRAGE: THCritMaskEdit;
    E_TRESOLETTRE: THCritMaskEdit;
    XX_WHEREENC: TEdit;
    XX_WHEREVIDE: TEdit;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    HLabel4: THLabel;
    E_GENERAL: THCritMaskEdit;
    HLabel17: THLabel;
    E_DEVISE: THValComboBox;
    HLabel7: THLabel;
    CATEGORIE: THValComboBox;
    Label14: TLabel;
    E_MODEPAIE: THValComboBox;
    TE_AUXILIAIRE: THLabel;
    E_AUXILIAIRE: THCritMaskEdit;
    HDiv: THMsgBox;
    XX_WHEREMP: TEdit;
    BParams: TToolbarButton97;
    HLabel1: THLabel;
    E_NUMENCADECA: THCritMaskEdit;
    PBanqueLot: TTabSheet;
    Bevel7: TBevel;
    HLabel3: THLabel;
    E_NOMLOT: THCritMaskEdit;
    HLabel5: THLabel;
    E_NOMLOT_: THCritMaskEdit;
    TE_BANQUEPREVI: THLabel;
    E_BANQUEPREVI: THCritMaskEdit;
    TE_BANQUEPREVI_: THLabel;
    E_BANQUEPREVI_: THCritMaskEdit;
    XX_WHERELOT: TEdit;
    CGenereAuto: TCheckBox;
    TCATEGORIE1: THLabel;
    CATEGORIE1: THValComboBox;
    BCtrlRib: TToolbarButton97;
    FE_NUMTRAITECHQ: TLabel;
    FE_NUMTRAITECHQ_: TLabel;
    E_NUMTRAITECHQ_: THCritMaskEdit;
    E_NUMTRAITECHQ: THCritMaskEdit;
    XX_WHERETRACHQ: TEdit;
    BRecapTraite: TToolbarButton97;
    PGestionPDF: TTabSheet;
    Bevel8: TBevel;
    LNBLIGNES: THLabel;
    TREPSPOOLER: THLabel;
    CAutoEnreg1: TCheckBox;
    NBLIGNES: THCritMaskEdit;
    Spooler: TCheckBox;
    RepSpooler: THCritMaskEdit;
    XFichierSpooler: TCheckBox;
    XX_WHERENATCPT: TEdit;
    FTID: TCheckBox;
    TE_NUMEROPIECE: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    cFactCredit: TCheckBox;
    TraiteEdite: TCheckBox;
    XX_WHEREPROFIL: TEdit;
    XX_WHEREMONTANT: TEdit;
    BLotEcr: TToolbarButton97;
    PEscompte: TTabSheet;
    Bevel9: TBevel;
    TESCHT: THLabel;
    ESCHT: THCritMaskEdit;
    TESCTVA: THLabel;
    ESCTVA: THCritMaskEdit;
    TESCTAUXTVA: TLabel;
    ESCTAUXTVA: THCritMaskEdit;
    ESCCBTVA: TCheckBox;
    ESCMETH: THValComboBox;
    TESCMETH: THLabel;
    TESCTAUXESC: TLabel;
    ESCTAUXESC: THCritMaskEdit;
    XX_WHEREESC: TEdit;
    FEscompte: TCheckBox;
    ESCCBPRORATA: TCheckBox;
    BSwapSelect: TToolbarButton97;
    _chk_RegSup: TCheckBox;
    _E_NATUREPIECE: THMultiValComboBox;
    CODEOMR: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);														override ;
    procedure CATEGORIEChange(Sender: TObject);
    procedure BParamsClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); 												override ;
    procedure BOuvrirClick(Sender: TObject);       										override ;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BCtrlRibClick(Sender: TObject);
    procedure FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
    procedure cFactCreditClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure SpoolerClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure BChercheClick(Sender: TObject);													override ;
    procedure FTIDClick(Sender: TObject);
    procedure TraiteEditeClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BLotEcrClick(Sender: TObject);
    procedure ESCMETHChange(Sender: TObject);
    procedure ESCCBTVAClick(Sender: TObject);
    procedure FEscompteClick(Sender: TObject);
    procedure ESCCBPRORATAClick(Sender: TObject);
    procedure FListeColEnter(Sender: TObject);
    procedure FListeFlipSelection(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure BSwapSelectClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
    procedure ESCTVAChange(Sender: TObject);
    procedure ESCHTChange(Sender: TObject);
    procedure _chk_RegSupClick(Sender: TObject);
    procedure _E_NATUREPIECEChange(Sender: TObject);
//    procedure FListeTitleClick(Column: TColumn);
  private    { Déclarations privées }
//    ANouveau  : boolean ;
    LesCombi,LesCritsRupt : TStringList ;
    TOBEscomptes,TOBGHT,TOBGTVA : TOB ;
    JalEsc : String ;
    SauveConceptFiltre : Byte ;
    SwapSelect : Boolean ;						 // Gestion du mode selection inversé
    gszRegSup : String;                // Pour la case à cocher : Factures avec un réglement supérieur
    gszCaption : String;               // Pour les messages

    saveCodeOmr : boolean;
    
    procedure InitCriteres ;
    procedure InitE_NUMTRACHQ(Invisible : Boolean) ;
    procedure InitContexte ;
    procedure ReinitWhereNatCpt(RAZ : Boolean) ;
    procedure InitEnc ;
    procedure InitDec ;
    procedure PrechargeOrigines ;
    procedure MarqueOrigine ;
    procedure ToutMarquer ;
    Function  ConstitueOrigines : Boolean ;
    procedure PreciseMP ;
    procedure IncNumLot ;
    Function  VerifEcheMP ( Alerte : Boolean ) : boolean ;
    {Escompte}
    Procedure ClearEscomptes ;
    Procedure SwapEscompteLigne ;
    Function  FindCreerTOBEscompte ( QueFind : boolean ) : TOB ;
    Function  OkLignePourEscompte : Boolean ;
    procedure AlimTobEscomptes ;
//    procedure ClickModifRib ;
    Function  AvecEscompte : Boolean ;
    Procedure ChangeModeGrille ;
    Procedure MarquerLigneSansEscompte ;
    Function  CalculSoldeSelection : Double ;
    Procedure AfficheSoldeSelection ;
    Procedure RempliWhereEnc;
    Procedure Enleve(pszNat : String);
    function LInsertDB(T : Tob) : Boolean;
 public
    ParmsMP                    : tGenereMP ;
    TOBORig,TOBDest            : TOB ;
    smp                        : TSuiviMP ;
    SorteLettre                : TSorteLettre ;
    ExNormal, ExOppose, DejaOP : Boolean ;
    SelectOnLot    : Boolean ;
    NomCritLot, NomCritReste   : String ;
  end;

implementation

uses FichComm;

{$R *.DFM}

Var SmpRes : String = '' ;

Procedure GenereSuiviMP ( smp : TSuiviMP ) ;
var
    X : TFMulGenereMP ;
    PP : THPanel ;
//    ind : tProfilTraitement ;
//    OkOk : Boolean ;
begin
    If EstSerie(S3) Then GereEscompte := FALSE;
    If Not ProfilOk(smp) Then
    BEGIN
        HShowMessage('0;Encaissements/Décaissements;Vous n''avez défini de profil !;E;O;O;O;','','') ;
        HShowMessage('1;Encaissements/Décaissements;Menu autres traitements / Profils utilisateur!;E;O;O;O;','','') ; Exit ;
    END ;
    (*
    Ind:=QuelProfil(smp,OkOk) ;
    If Not MonProfilOk(Ind) Then
    BEGIN
        smpRes:='SMP'+IntToStr(Ord(smp)) ;
        if Blocage([smpRes],True,smpRes) then Exit ;
    END ;
    *)
    PP:=FindInsidePanel ;

    X:=TFMulGenereMP.Create(Application) ;

    X.Q.Manuel:=True ;
    X.smp:=smp ;
    X.SorteLettre:=AttribSL(smp) ;
    X.FNomFiltre:=AttribFiltre(smp) ;
    X.Q.Liste:=AttribListe(smp) ;
    X.HelpContext:=AttribHelp(smp,FALSE) ;

    if PP=Nil then
    BEGIN
        try
            X.ShowModal ;
        finally
            X.Free ;
//            If Not MonProfilOk(Ind) Then Bloqueur(smpRes,False) ;
        end;
        Screen.Cursor:=SyncrDefault ;
    END else
    BEGIN
        InitInside(X,PP) ;
        X.Show ;
    END ;
END ;

(*========================= METHODES DE LA FORM ==============================*)
procedure TFMulGenereMP.FormCreate(Sender: TObject);
begin
  inherited;
  MemoStyle := msBook ;

  TOBOrig       := TOB.Create('',Nil,-1) ;
  TOBDest       := TOB.Create('',Nil,-1) ;
  TOBEscomptes  := TOB.Create('LESESCOMPTES',Nil,-1) ;
  TOBGHT        := TOB.Create('',Nil,-1) ;
  TOBGTVA       := TOB.Create('',Nil,-1) ;

  LesCombi                := TStringList.Create ;
  LesCombi.Sorted         := True ;
  LesCombi.Duplicates     := dupIgnore ;

  LesCritsRupt            := TStringList.Create ;
  LesCritsRupt.Sorted     := True ;
  LesCritsRupt.Duplicates := dupIgnore ;

  DejaOP      := False ;
  SwapSelect  := False ;   // Gestion du mode "selection inversé", non opérationnel par défaut
end;

procedure TFMulGenereMP.FormShow(Sender: TObject);
begin
  // FQ 10572
  Case smp Of
    smpEncDiv,smpDecDiv : BEGIN
                          BLotEcr.Visible:=FALSE ;
                          PBanqueLot.TabVisible:=FALSE ;
                          Pages.ActivePage:=PCritere ;
                          END ;
    Else BEGIN
         // Page de sélection sur lot non apparrente au lancement
         PBanqueLot.TabVisible:=FALSE ;
         Pages.ActivePage:=PCritere ;
    END ;
  End;

    // BPY le 11/12/2003 => Fiche 11869 : etant dans l'imposibilité de suprimé dans la fiche ... je cache ici !
    _chk_RegSup.Visible := false;
    // fin BPY ...

    saveCodeOmr := V_PGI.MiseSousPli;
{$IFDEF CCMP}
    if EstSerie(S5) then
    begin
        CODEOMR.Checked := V_PGI.MiseSousPli;
        if (smp in [smpEncTraEdt,smpEncTraEdtNC,smpDecBorEdt,smpDecChqEdt,smpDecVirEdt,smpDecVirInEdt,smpDecBorEdtNC,smpDecChqEdtNC,smpDecVirEdtNC,smpDecVirInEdtNC]) then
        CODEOMR.Visible := True;
    end
    else
    begin
        CODEOMR.Visible := false;
        V_PGI.MiseSousPli:=False ;
    end ;
{$ELSE}
    CODEOMR.Visible := false;
{$ENDIF}

  // Constexte d'aide
  if (smp=smpDecVirEdt)     then HelpContext:=999999433;
  if (smp=smpDecVirEdtNC)   then HelpContext:=999999435;
  if (smp=smpDecVirBqe)     then HelpContext:=999999437;
  if (smp=smpDecVirInEdt)   then HelpContext:=999999439;
  if (smp=smpDecVirInEdtNC) then HelpContext:=999999440;
  if (smp=smpDecVirInBqe)   then HelpContext:=999999442;

  SauveConceptFiltre:=NivConcept[ccActionsFiltre] ;
  NivConcept[ccActionsFiltre]:=0 ;
  If smp In [smpDecChqEdt,smpDecVirEdt,smpDecVirInEdt] Then
    BEGIN
    GereEscompte         := TRUE ;
    JalEsc:=GetParamSocSecur('SO_CPJALESCOMPTE','') ;
    if JalEsc='' then
      GereEscompte       := False ;
    FEscompte.Visible    := TRUE ;
    FEscompteClick(nil);
    END
  Else
    BEGIN
    JalEsc               := '' ;
    GereEscompte         := False ;
    XX_WHEREESC.Text     := '' ;
    FEscompte.Checked    := FALSE ;
    FEscompte.Visible    := False ;
    FEscompteClick(nil);
    END ;
  NomCritLot    := 'LOT' + V_PGI.User ;
  NomCritReste  := 'CRIT' + V_PGI.User ;
  InitCriteres ;
  if (smp<>smpEncTraEdt) And (smp<>smpEncTraEdtNC) And
     (smp<>smpDecBorEdt) And (smp<>smpDecBorEdtNC) Then
     BEGIN
     TraiteEdite.State        := cbGrayed ;
     E_NUMTRAITECHQ_.Visible  := FALSE ;
     FE_NUMTRAITECHQ_.Enabled := FALSE ;
     E_NUMTRAITECHQ.Enabled   := FALSE ;
     FE_NUMTRAITECHQ.Enabled  := FALSE ;
     E_NUMTRAITECHQ_.Text     := '' ;
     E_NUMTRAITECHQ.Text      := '' ;
     XX_WHERETRACHQ.Text      := '' ;
     END ;
  TraiteEditeClick(Nil) ;
  inherited;
  _E_NATUREPIECE.Value := '';
  InitContexte ;
  selectOnLot:=FALSE ;
  VH^.MPModifFaite := FALSE ;
  Fillchar(VH^.MPPOP,SizeOf(VH^.MPPOP),#0) ;
  Q.Manuel := FALSE ;
  Q.UpdateCriteres ;
  CentreDBGrid(FListe) ;
  XX_WHEREVIDE.Text:='' ;
  InitEscompteSup(TOBEscomptes) ;
  SaveFiltre(NomCritLot,FFiltres,Pages) ;
  SaveFiltre(NomCritReste,FFiltres,Pages) ;
  // Page de sélection sur lot non apparrente au lancement
  If IsEnc(SMP) Then
    If (Not VH^.CCMP.LotCli) Then BLotEcrClick(Nil) ;
  If Not IsEnc(SMP) Then
    If (Not VH^.CCMP.LotFou) Then BLotEcrClick(Nil) ;
  // MAJ des zones des comptes/taux tva pour escompte
  ESCCBTVAClick(nil);
end;

(*============================ INITIALISATIONS ===============================*)
procedure TFMulGenereMP.PreciseMP ;
BEGIN
  If Categorie.itemIndex=0
    Then CatToMP('',E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True)
    Else CatToMP(CATEGORIE.Value,E_MODEPAIE.Items,E_MODEPAIE.Values,tslAucun,True) ;
  // On ôte les virement internationnaux dans le menu BOR // SBO fiche 10594
  if smp in [smpDecBorEdt,smpDecBorEdtNC,smpDecborDec,smpDecTraPor] then
    CVireLigne( E_MODEPAIE , 'VIN' ) ;
  // Fin SBO fiche 10594
END ;

procedure TFMulGenereMP.ReinitWhereNatCpt(RAZ : Boolean) ;
Var StXP,StXP2,StXN,StXN2,St : String ;
BEGIN
  If RAZ Then
    XX_WHERENATCPT.Text:=''
  Else
    BEGIN
    If FTID.Checked Then
      BEGIN
        If IsEnc(SMP)
          Then XX_WHERENATCPT.Text:='E_AUXILIAIRE="" AND G_NATUREGENE<>"TIC"'
          Else XX_WHERENATCPT.Text:='E_AUXILIAIRE="" AND G_NATUREGENE<>"TID"';
      END
    Else
      BEGIN
      If IsEnc(SMP)
        Then XX_WHERENATCPT.Text:='E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"FOU" AND T_NATUREAUXI<>"AUC"'
        Else XX_WHERENATCPT.Text:='E_AUXILIAIRE<>"" AND T_NATUREAUXI<>"CLI" AND T_NATUREAUXI<>"AUD"';
      END ;
  END ;

  StXP  := StrFPoint(9*Resolution(V_PGI.OkDecV+1))  ;
  StXN  := StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
  StXP2 := StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ;
  StXN2 := StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
  St:='(E_DEBIT+E_CREDIT-E_COUVERTURE not between '+StXN+' AND '+StXP+')' ;
  XX_WHEREMONTANT.Text := St ;
{ // 11673
If _chk_RegSup.State=cbGrayed Then begin // AVEC
  if (VH^.CCMP.LotCli) then gSzRegSup := ' OR E_NATUREPIECE="RC"'
                       else gSzRegSup := ' OR E_NATUREPIECE="RF"';
  cFactCreditClick(nil);
end;
If _chk_RegSup.State=cbChecked Then begin // UNIQUEMENT
  gSzRegSup:= '';
  if (VH^.CCMP.LotCli) then XX_WHEREENC.Text := 'E_ENCAISSEMENT="DEC" AND E_NATUREPIECE="RC"'
                       else XX_WHEREENC.Text := 'E_ENCAISSEMENT="ENC" AND E_NATUREPIECE="RF"';
end;
if _chk_RegSup.State = cbUnChecked then begin // SANS
//  Ne fait rien
end;}
END ;

procedure TFMulGenereMP.InitEnc ;
var
  szNatPie,szTemp : String;
begin
  szNatPie := _E_NATUREPIECE.Value;
  if (szNatPie <> '') then
    begin
    while szNatPie<>'' do
      szTemp := szTemp + ' OR E_NATUREPIECE="'+ReadTokenSt(szNatPie)+'"';
    if (gSzRegSup<>'')
      then szTemp:='';
    If cFactCredit.State=cbGrayed then    // AVEC
      XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"'+gSzRegSup+szTemp+'))';
    If cFactCredit.State=cbChecked then   // UNIQUEMENT
      XX_WHEREENC.Text:='(E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC" OR E_NATUREPIECE="FC"'+gSzRegSup+szTemp+'))';
    If cFactCredit.State=cbUnChecked then // SANS
      XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND (E_NATUREPIECE="AC" OR E_NATUREPIECE="OC"'+gSzRegSup+szTemp+'))';
    end
  else
    XX_WHEREENC.Text := '';

  E_GENERAL.DataType    := 'TZGENCAIS' ;
  E_AUXILIAIRE.DataType := 'TZTTOUTDEBIT' ;
  if smp = smpEncTous then
    XX_WHERELOT.Text := 'E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' ;
end ;

procedure TFMulGenereMP.InitDec ;
var
  szNatPie,szTemp : String;
begin
  szNatPie := _E_NATUREPIECE.Value;
  if (szNatPie <> '') then
    begin
    while szNatPie<>'' do
      szTemp := szTemp + ' OR E_NATUREPIECE="'+ReadTokenSt(szNatPie)+'"';
    if (gSzRegSup<>'') then
      szTemp:='';
    If cFactCredit.State=cbGrayed    Then
      XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"'+gSzRegSup+szTemp+'))';
    If cFactCredit.State=cbChecked   Then
      XX_WHEREENC.Text:='(E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF" OR E_NATUREPIECE="FF"'+gSzRegSup+szTemp+'))';
    If cFactCredit.State=cbUnChecked Then
      XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND (E_NATUREPIECE="AF" OR E_NATUREPIECE="OF"'+gSzRegSup+szTemp+'))';
    end
  else
    XX_WHEREENC.Text := '';

  E_GENERAL.DataType    := 'TZGDECAIS' ;
  E_AUXILIAIRE.DataType := 'TZTTOUTCREDIT' ;
  if smp=smpDecTous then
    XX_WHERELOT.Text  := 'E_NOMLOT<>"" AND E_BANQUEPREVI<>"" AND E_MODEPAIE<>""' ;
  cFactCredit.Caption := HDiv.Mess[15] ;
end ;

procedure TFMulGenereMP.InitE_NUMTRACHQ(Invisible : Boolean) ;
begin
  E_NUMTRAITECHQ_.Visible  := Not Invisible ;
  FE_NUMTRAITECHQ_.Visible := Not Invisible ;
  E_NUMTRAITECHQ.Visible   := Not Invisible ;
  FE_NUMTRAITECHQ.Visible  := Not Invisible ;
end ;

procedure TFMulGenereMP.InitContexte ;
begin
  InitE_NUMTRACHQ( smp in [smpEncPreEdt,smpEncTous,smpDecVirEdt,smpDecVirBqe,smpDecVirInEdt,smpDecVirInBqe,smpEncPreBqe] ) ;
  E_NOMLOT.Plus  :=AttribPlus(smp) ;
  E_NOMLOT_.Plus :=AttribPlus(smp) ;
  Case smp of
    smpEncDiv    : BEGIN // Encaissement divers
                  InitEnc ; CATEGORIE.ItemIndex:=0 ; Caption:=HDiv.Mess[20] ;
                  END ;
    smpEncPreEdt : BEGIN // Prélèvements clients
                  InitEnc ; CATEGORIE.Value:='PRE' ; Caption:=HDiv.Mess[19] ;
                  END ;
    smpEncPreBqe : BEGIN // Prélèvements clients
                  InitEnc ; CATEGORIE.Value:='CB' ; Caption:=HDiv.Mess[29] ;
                  END ;
    smpEncChqPor : BEGIN // Mise en portefeuille chèques clients
                  InitEnc ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[28] ;
                  END ;
    smpEncCBPor : BEGIN // Mise en portefeuille cartes bleues clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[16] ;
                  END ;
    smpEncTraPor : BEGIN // Mise en portefeuille traite clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[16] ;
                  END ;
    smpEncTraEdt : BEGIN // Lettres-traite clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[5] ;
                  END ;
    smpEncTraEdtNC : BEGIN // Lettres-traite clients
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[26] ;
                  END ;
    smpEncTraEnc : BEGIN // Traite clients à l'encaissement
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[8] ;
                  END ;
    smpEncTraEsc : BEGIN // Traite clients à l'escompte
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[9] ;
                  END ;
    smpEncChqBqe : BEGIN // Chèque clients en banque
                  InitEnc ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[31] ;
                  END ;
    smpEncCBBqe : BEGIN // CB clients en banque
                  InitEnc ; CATEGORIE.Value:='CB' ; Caption:=HDiv.Mess[6] ;
                  END ;
    smpEncTraBqe : BEGIN // Traite clients en banque
                  InitEnc ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[6] ;
                  END ;
    smpEncTous : BEGIN // Prélèvements clients
                  InitEnc ; CATEGORIE.Value:='' ; Caption:=HDiv.Mess[12] ;
                  END ;
    smpDecChqEdt : BEGIN // Lettres-chèques fournisseurs
                  InitDec ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[2] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[25] ;
                  END ;
    smpDecChqEdtNC : BEGIN // Lettres-chèques fournisseurs
                  InitDec ; CATEGORIE.Value:='CHQ' ; Caption:=HDiv.Mess[32] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[25] ;
                  END ;
    smpDecVirEdt : BEGIN // Virements fournisseurs
                  InitDec ; CATEGORIE.Value:='VIR' ; Caption:=HDiv.Mess[18] ;
                  END ;
    smpDecVirEdtNC : BEGIN // Virements fournisseurs
                  InitDec ; CATEGORIE.Value:='VIR' ; Caption:=HDiv.Mess[33] ;
                  END ;
    smpDecVirBqe : BEGIN // Virements fournisseurs
                  InitDec ; CATEGORIE.Value:='VIR' ; Caption:=HDiv.Mess[3] ;
                  END ;
    smpDecVirInEdt : BEGIN // Virements internationaux fournisseurs
                  InitDec ; CATEGORIE.Value:='TRI' ; Caption:=HDiv.Mess[36] ;
                  END ;
    smpDecVirInEdtNC : BEGIN // Virements internationaux fournisseurs
                  InitDec ; CATEGORIE.Value:='TRI' ; Caption:=HDiv.Mess[35] ;
                  END ;
    smpDecVirInBqe : BEGIN // Virements internationaux fournisseurs
                  InitDec ; CATEGORIE.Value:='TRI' ; Caption:=HDiv.Mess[34] ;
                  END ;
    smpDecBorEdt : BEGIN // Lettres-Bor fournisseurs
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[27] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
    smpDecBorEdtNC : BEGIN // Lettres-Bor fournisseurs
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[4] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
    smpDecborDec : BEGIN // Traites fournisseurs en encaissement
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[10] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
    smpDecBorEsc : BEGIN // Traites fournisseurs à l'escompte
                  InitDec ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[11] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
    smpDecTraBqe : BEGIN // Traites fournisseurs en banque
                  InitDec ; CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[7] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
    smpDecTraPor : BEGIN // Mise en portefeuille traite clients
                  InitDec ;
                  CVireLigne(Categorie, 'TRI') ; // SBO fiche 10594 virer les virements internationaux
                  CATEGORIE.Value:='LCR' ; Caption:=HDiv.Mess[17] ;
                  FE_NUMTRAITECHQ.Caption:=HDiv.Mess[24] ;
                  END ;
    smpDecDiv    : BEGIN // Décaissements divers
                  InitDec ; CATEGORIE.ItemIndex:=0 ; Caption:=HDiv.Mess[21] ;
                  END ;
    smpDecTous : BEGIN //
                  InitDec ; CATEGORIE.Value:='' ; Caption:=HDiv.Mess[13] ;
                  END ;
    END ;

  gszCaption  := Caption;  // Pour les messages
  ParmsMP.Cat := CATEGORIE.Value ;
  ParmsMP.smp := smp ;
  PreciseMP ;
  UpdateCaption(Self) ;
END ;

procedure TFMulGenereMP.InitCriteres ;
Var i : integer ;
    StV8,St : String ;
BEGIN
  if ((smp<>smpEncTous) and (smp<>smpDecTous)) then
    BEGIN
    LibellesTableLibre(PzLibre,'TT_TABLE','T_TABLE','T') ;
    if VH^.CPExoRef.Code<>'' then
      BEGIN
      E_EXERCICE.Value      := VH^.CPExoRef.Code ;
      E_DATECOMPTABLE.Text  := DateToStr(VH^.CPExoRef.Deb) ;
      E_DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin) ;
      END
    else
      BEGIN
      E_EXERCICE.Value      := VH^.Entree.Code ;
      E_DATECOMPTABLE.Text  := DateToStr(VH^.Entree.Deb) ;
      E_DATECOMPTABLE_.Text := DateToStr(VH^.Entree.Fin) ;
      END ;
    E_DATEECHEANCE.Text  := StDate1900 ;
    E_DATEECHEANCE_.Text := StDate2099 ;
    E_DEVISE.Value       := V_PGI.DevisePivot ;
    PGestionPDF.TabVisible := smp in [smpEncPreEdt,smpEncTraEdt,smpEncTraEdtNC,smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC,smpDecBorEdt,smpDecBorEdtNC] ;
    If (smp=smpDecChqEdtNC) or (smp=smpDecVirInEdtNC) or (smp=smpDecVirEdtNC) or (smp=smpEncTraEdtNC) or (smp=smpDecBorEdtNC) Then begin // 13746
      CGenereAuto.Visible   := FALSE ;
      end;
    If (smp=smpDecChqEdt) or (smp=smpDecVirEdt) or (smp=smpDecVirInEdt) or (smp=smpEncPreEdt) Then
      BEGIN
      PBanqueLot.TabVisible := TRUE ;
      CGenereAuto.Visible   := FALSE ;
      CATEGORIE1.Visible    := FALSE ;
      TCATEGORIE1.Visible   := FALSE ;
      END ;
    END
  else
   BEGIN
   Pages.ActivePage := PBanqueLot ;
   for i:=0 to Pages.PageCount-1 do
    if Pages.Pages[i]<>PBanqueLot then
      Pages.Pages[i].TabVisible := False ;
   END ;
  // ParmsMP
  FillChar(ParmsMP,Sizeof(ParmsMP),#0) ;
  ParmsMP.NomFSelect := FNomFiltre ;
  St    := 'E_ECRANOUVEAU="N"' ;
  StV8  := LWhereV8 ;
  if StV8<>'' then
    BEGIN
    St:='('+St+' OR E_ECRANOUVEAU="H")  ' ;
    St:=St+' AND ('+StV8+') ' ;
    END ;
  XX_WHEREAN.Text := St ;
  ReinitWhereNatCpt(FALSE) ;
  XX_WHEREPROFIL.Text := '' ;
END ;

(*================================ CRITERES ==================================*)
procedure TFMulGenereMP.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
  ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFMulGenereMP.CATEGORIEChange(Sender: TObject);
Var HCAT : THValComboBox ;
begin
  inherited;
  // Le 28/01/2003 BPY correction des bug de la fiche 11794
  if (Categorie.Value = 'TRI')
    then BCtrlRib.visible := false
    else BCtrlRib.visible := true;
  // Fin BPY
  HCAT := THValCombobox(Sender) ;
  LaCategorieChange( HCat, XX_WHEREMP, SorteLettre ) ;
  PreciseMP ;
end;

procedure TFMulGenereMP.BParamsClick(Sender: TObject);
begin
  inherited;
  ParmsMP.CptS      := E_GENERAL.Text ;
  ParmsMP.DEV.Code  := E_DEVISE.Value ;
  ParmsMP.TIDTIC    := FTID.Checked ;
  ParamsMP(False,ParmsMP) ;
  DejaOP            := True ;
end;

procedure TFMulGenereMP.FListeDblClick(Sender: TObject);
Var sMode : String ;
begin
  inherited;
  {$IFNDEF IMP}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  sMode := Q.FindField('E_MODESAISIE').AsString ;
  if ((sMode<>'') and (sMode<>'-'))
    then LanceSaisieFolio(Q,TypeAction)
    else TrouveEtLanceSaisie(Q,taConsult,'N') ;
  {$ENDIF}
end;

{procedure TFMulGenereMP.ClickModifRib ;
Var RJal,RExo : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
begin
  If ModifRibSurMul(Q,FTID.Checked,TRUE) Then
    BEGIN
    Application.ProcessMessages ; BChercheClick(Nil) ;
    RJal:=Q.FindField('E_JOURNAL').AsString ;
    RExo:=QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime)) ;
    RDate:=Q.FindField('E_DATECOMPTABLE').AsDateTime ;
    RNumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
    RNumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
    RNumEche:=Q.FindField('E_NUMECHE').AsInteger ;
    Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
            VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]),[])
    END ;
end;}

procedure TFMulGenereMP.PrechargeOrigines ;
Var QQ : TQuery ;
    st : String ;
BEGIN
  TOBOrig.ClearDetail ;
  St:='SELECT ECRITURE.*, G_NATUREGENE,T_NATUREAUXI,T_ESCOMPTE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '+
      'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;
  QQ:=OpenSQL(St+RecupWhereCritere(Pages),True) ;
  TOBOrig.LoadDetailDB('ECRITURE','','',QQ,False,True) ;
  Ferme(QQ) ;
  if TOBOrig.Detail.Count>0 then
    TOBOrig.Detail[0].AddChampSup('MARQUE',True) ;
END ;

procedure TFMulGenereMP.MarqueOrigine ;
Var TOBL : TOB ;
BEGIN
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
END ;

procedure TFMulGenereMP.ToutMarquer ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
  for i:=0 to TOBORig.Detail.Count-1 do
    BEGIN
    TOBL:=TOBOrig.Detail[i] ;
    TOBL.PutValue('MARQUE','X') ;
    // Gestion escompte fournisseur à la ligne
    if (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC]) then
    	begin
      // Recherche TobEscompte, Création si non trouvé
      if not GereEscompte then Continue; // Ne prenait qu'une
      If Not AvecEscompte Then Continue; // seule ligne de la sélection
      FindCreerTOBEscomptePourLigne( TOBEscomptes, TOBL, False ) ;
      end ;
    END ;
END ;

Function TFMulGenereMP.ConstitueOrigines : Boolean ;
Var i {,k} : integer ;
    TOBL : TOB ;
    StBQE,StCAT,StMP,St,LesChamps,NomChamp,LesVals : String ;
    Premier : Boolean ;
    OldMP : String ;
    RuptMP : Boolean ;
    ResaisieMP,PasPris : Boolean ;
//    lBoTestCPT         : Boolean ;
Label 0 ;
begin
0 :
  Result    := True ;
  ExNormal  := False ;
  ExOppose  := False ;
  PrechargeOrigines ;
  // ============== DEBUT GESTION MODE SELECTION INVERSEE
  if swapSelect then
  	// Nouveau code : ajout mode selection inversée
  	begin
    Q.First ;
    While not Q.Eof do
      begin
      if  not FListe.IsCurrentSelected then MarqueOrigine ;
      Q.Next ;
      end ;
    end
  else
  	// Ancien code : avant ajour mode selection inversée
    begin
    if Not FListe.AllSelected then
    	begin
      for i:=0 to FListe.NbSelected-1 do
        begin
        FListe.GotoLeBookmark(i) ;
        MarqueOrigine ;
        end ;
      end
    else
      ToutMarquer ;
    end ;
  // ============== FIN GESTION MODE SELECTION INVERSEE
  VireInutiles( TOBOrig ) ;
  Premier := TRUE ;
  RuptMP  := FALSE ;
  AlimTobEscomptes ; // Alim entete de tobescomptes avec valeurs du critère
  PasPris := FALSE ;

  // Vérifications : Compte génération <> compte pièce comptable
//  lBoTestCPT := True ;
  for i:=0 to TOBOrig.Detail.Count-1 do
    BEGIN
    TOBL := TOBOrig.Detail[i] ;
    if TOBL.GetValue('E_GENERAL')=ParmsMP.CptG then
      BEGIN
      Result := FALSE ;
      HM.Execute(4,gszCaption,'') ;
      Break ;
      END ;
    // Ce test n'est plus bloquant
    // EN ATTENTE MODIF LETTRAGE : SBO
{    if TOBL.GetValue('E_GENERAL')=ParmsMP.CptG
      then lBoTestCPT := False ;
}   ExNormal:=True ;
    if Premier
      then OldMP  := TOBL.GetValue('E_MODEPAIE')
      Else RuptMP := (TOBL.GetValue('E_MODEPAIE')<>OldMP) ;
    Premier:=FALSE ;
    TraiteTOBEscompte( TOBL, TOBEscomptes, TOBGHT, TOBGTVA , ESCCBPRORATA.Checked, ESCMETH.Value, ParmsMP, PasPris ) ;
    END ;

  // Affichage message si compte géné = compte pièce origine MAIS non bloquant
  if Not Result Then Exit ;
    // EN ATTENTE MODIF LETTRAGE : SBO
{  if not lBoTestCPT then
    if HM.Execute(4,gszCaption,'')<>mrYes Then
      begin
      result := false ;
      Exit ;
      end ;
}
  If PasPris And GereEscompte And (ESCCBPRORATA.Checked) Then
    If HM.Execute(10,gszCaption,'')<>mrYes Then Exit ;

  If RuptMP And (ParmsMP.MPG='') Then
    BEGIN
    ResaisieMP:=FALSE ;
    // Début 11640
    case smp Of
      smpEncTraEdt,smpEncTraPor : If HM.Execute(2,gszCaption,'')<>mrYes Then
                                    begin
                                    Result := False;
                                    Exit;
                                    end;
      else
        begin
        HM.Execute(8,gszCaption,'') ;
        ResaisieMP:=TRUE ;
        end ;
    end ;
    // Fin 11640
    If ResaisieMp Then
      if Not ParamsMPSup(ParmsMP,TRUE) then
        begin
        Result := False;
        Exit;
        end
      else Goto 0;
    end ;

  AjouteLesAnas( TOBOrig ) ;
  if smp in [smpEncTous,smpDecTous] then
    BEGIN
    for i:=0 to TOBOrig.Detail.Count-1 do
      BEGIN
      TOBL  := TOBOrig.Detail[i] ;
      StBQE := TOBL.GetValue('E_BANQUEPREVI') ;
      StMP  := TOBL.GetValue('E_MODEPAIE') ;
      StCat := MPToCategorie(StMP) ;
      if StCat<>'' then
        BEGIN
        St  := StBQE+';'+StCat+';' ;
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
      END ;
    END
  else
    if ParmsMP.ChampsRupt<>'' then
      BEGIN
      for i:=0 to TOBOrig.Detail.Count-1 do
        BEGIN
        TOBL      := TOBOrig.Detail[i] ;
        LesChamps := ParmsMP.ChampsRupt ;
        LesVals   := '' ;
        Repeat
          NomChamp:=ReadTokenSt(LesChamps) ;
          if NomChamp<>'' then
            LesVals:=LesVals+TOBL.GetValue(NomChamp)+';' ;
        Until ((LesChamps='') or (NomChamp=''));
        if LesCritsRupt.IndexOf(LesVals)<0 then
          LesCritsRupt.Add(LesVals) ;
        END ;
      END ;
END ;

procedure TFMulGenereMP.IncNumLot ;
Var NL : String ;
begin
  NL := IncNumLotTraChq(ParmsMP.NumEncaDeca) ;
  ParmsMP.NumEncaDeca:=Copy(NL,1,17) ;
  CreerCodeLot(ParmsMP.NumEncaDeca) ;
end ;

Function TFMulGenereMP.VerifEcheMP ( Alerte : Boolean ) : boolean ;
Var i : integer ;
    TOBE : TOB ;
    Premier,RuptMP,RuptEche,ExRuptEche,ExRuptMP : boolean ;
    OldEche : TDateTime ;
    OldMP   : String ;
    ind     : integer ;
begin
  Result:=True ;
  if TobDest.Detail.Count = 0 then
    begin
    Result := False ;
    PGIInfo('Aucun auxiliaire à traiter ! Le solde de la sélection est nul.', Caption ) ;
    Exit ;
    end ;
  if Not Alerte then Exit ;
  // Initialisation des variables locales
	ExRuptEche := False ;
  ExRuptMP   := False ;
  RuptMP		 := False ;
  RuptEche   := False ;
  Premier 	 := False ;
  OldEche    := iDate1900 ;
  OldMP      := '' ;
  for i:=0 to TOBDest.Detail.Count-1 do
    begin
    TOBE := TOBDest.Detail[i] ;
    if TOBE.GetValue('E_GENERAL')=ParmsMP.CptG then
      begin
      Premier  := True ;
      RuptMP   := False ;
      RuptEche := False ;
      end ;
    if Premier then
      begin
      OldMP   := TOBE.GetValue('E_MODEPAIE') ;
      OldEche := TOBE.GetValue('E_DATEECHEANCE') ;
      end
    else
      begin
      if Not ParmsMP.ForceEche then
        RuptEche := (TOBE.GetValue('E_DATEECHEANCE')<>OldEche) ;
      if ParmsMP.MPG='' then
        RuptMP := (TOBE.GetValue('E_MODEPAIE')<>OldMP) ;
      end ;
    Premier := False ;
    if RuptMP then   ExRuptMP   := True ;
    if RuptEche then ExRuptEche := True ;
    if ((ExRuptEche) and (ExRuptMP)) then Break ;
    end ;
  if ((Not ExRuptEche) and (Not ExRuptMP)) then Exit ;
  if ((ExRuptEche) and (ExRuptMP))
    then ind:=1
    else if ExRuptMP
           then ind:=2
           else ind:=3 ;
  if HM.Execute(ind,gszCaption,'')<>mrYes then
    Result:=False ;
end ;

procedure TFMulGenereMP.BOuvrirClick(Sender: TObject);
Var MM : RMVT ;
    TOBE,TOBO2,TOBEsc : TOB ;
    i,k  : integer ;
    AlerteEcheMP,PasBon,OkJal,PbJal : boolean ;
    St,sBqe,sCat,{sM,}OldBqe : String ;
//    SoucheSpooler : Integer ;
    CarePDF : Boolean ;
    ShunteParamMP : Boolean ;
begin
// Le résultat de la recherche est-il vide ?
  	if (Q.Eof and Q.Bof) then
    begin
    PGIInfo('Le résultat de la recherche est vide. Veuillez relancer une recherche.',gszCaption) ;
    exit ;
    end ;
// Au moins une ligne de sélectionnée en mode normale, (pas de test en mode inverse pour l'instant)
	if (((FListe.NbSelected=0) and not(FListe.AllSelected)) and (not swapSelect)) then
    begin
    PGIInfo('Aucune ligne à traiter. Vous devez sélectionner au moins une écriture.',gszCaption) ;
    exit ;
    end ;
// Test et avertissement en cas d'escompte
  if FEscompte.Checked and AvecEscompte then
  	begin
    // Blocage sur le compte/taux d'escompte.
    if ((ESCHT.Text='') or (not IsNumeric(ESCTAUXESC.Text)) or (Valeur(ESCTAUXESC.Text)=0)) then
      begin
    	PGIBox('Vous n''avez pas renseigné le compte et/ou le taux d''escompte.',gszCaption);
      exit;
      end ;
    // Blocage sur le compte/taux de TVA Si Coche "TVA sur escompte" cochée.
    if ESCCBTVA.Checked and ((ESCTVA.Text='') or (not IsNumeric(ESCTAUXTVA.Text)) or (Valeur(ESCTAUXTVA.Text)=0)) then
      begin
    	PGIBox('Vous n''avez pas renseigné le compte et/ou le taux de TVA.',gszCaption);
      exit;
      end ;
{   ---- MIS EN COMMENTAIRE CAR UTILITE FONCTIONELLE A VALIDER ---- Ne pas effacer
    // Avertissement pour le compte/taux de TVA si Coche "TVA sur escompte" non cochée.
    if ((ESCTVA.Text='') or (not IsNumeric(ESCTAUXTVA.Text)) or (Valeur(ESCTAUXTVA.Text)=0)) then
      begin
      if PGIAsk('Vous n''avez pas renseigné le compte et/ou le taux de TVA. Voulez-vous continuez ?',gszCaption) <> mrYes
        then Exit ;
      end ;}
    end ;
// Message avertissement pour sélection inversée
	if swapSelect then
  	if PGIAsk('Vous êtes en mode "sélection inversée", Le traitement peut être long, voulez-vous continuez ?',gszCaption) <> mrYes
    then  exit;

    V_PGI.MiseSousPli := CODEOMR.Checked;

CarePdf:=FALSE ;
ExNormal:=False ; ExOppose:=False ;
if smp in [smpEncTous,smpDecTous] then else
   BEGIN
   if E_GENERAL.Text='' then BEGIN (* GP A VALIDER HM.Execute(0,gszCaption,'') ; Exit ; *) END ;
   END ;
ParmsMP.DEV.Code:=E_DEVISE.Value ; GetInfosDevise(ParmsMP.DEV) ;
ParmsMP.TIDTIC:=FTID.Checked ;
ShunteParamMP:=smp In [smpEncTraEdtNC,smpDecBorEdtNC] ;
If ShunteparamMP Then ShunteparamMP:=ParamsMPAuto(ParmsMP) ;
If ShunteParamMP Then Else
  BEGIN
  if Not ParamsMP(True,ParmsMP) then Exit ;
  END ;
if ((smp in [smpEncTous,smpDecTous]) and (ParmsMP.ChampsRupt='')) then ParmsMP.ChampsRupt:='E_QUALIFPIECE;' ;
if Pos('E_QUALIFPIECE',ParmsMP.ChampsRupt)<=0 then ParmsMP.ChampsRupt:='E_QUALIFPIECE;'+ParmsMP.ChampsRupt ;
//ParmsMP.ChampsRupt:='E_AUXILIAIRE;' ;
AlerteEcheMP:=ParmsMP.AlerteEcheMP ;
If FEscompte.Checked And AvecEscompte Then If HM.Execute(9,gszCaption,'')<>mrYes Then Exit ;
  inherited;
If Not ConstitueOrigines Then Exit ; OldBqe:='' ;
if TOBOrig.Detail.Count<=0 then Exit ;
DejaOP:=True ; PasBon:=False ; PbJal:=FALSE ;
If Spooler.Checked And (Not V_PGI.QRPDF) Then BEGIN V_PGI.QRPDF:=TRUE ; CarePdf:=TRUE ; END ;
if smp in [smpEncTous,smpDecTous] then
   BEGIN
   for i:=0 to LesCombi.Count-1 do
       BEGIN
       St:=LesCombi[i] ; sBqe:=ReadtokenSt(St) ; sCat:=ReadtokenSt(St) ;
       // Tripoter ParmsMP
       ParmsMP.CptG:=sBqe ; OkJal:=TRUE ;
       ParmsMP.JalG:=CompteToJal(sBqe,OkJal) ; If Not OkJal Then PbJal:=TRUE ;
       If Not OkJal Then Continue ;
       if i=0 then OldBqe:=sBqe ;
       if OldBqe<>sBqe then IncNumLot ;
       for k:=0 to LesCritsRupt.Count-1 do
           BEGIN
           ConstitueDest( sBQE, sCat, LesCritsRupt[k], TOBOrig, TOBDest, ParmsMP ) ;
           if VerifEcheMP(AlerteEcheMP) then
              BEGIN
              if LInsertDB(TobDest) then
                 BEGIN
                 TOBE:=TOBDest.Detail[0] ;
                 MM:=TOBToIdent(TOBE,False) ;
                 CompleteMM ( ParmsMP, SorteLettre, Spooler.Checked, XFichierSpooler.Checked, RepSpooler.Text, MM ) ;
                 TOBO2:=ExtractDiscerne(TOBOrig,sBqe,sCat,LesCritsRupt[k],ParmsMP) ;
                 TOBEsc := GenereEscompteMP( TOBO2, TOBEscomptes, TOBGHT, TOBGTVA,
                                             ESCCBTVA.Checked , JalEsc , ParmsMP ) ;
                 if TOBO2<>Nil then
                    BEGIN
                    if LanceSaisieMP(MM,TOBO2,TOBEsc,CGenereAuto.Checked,TOBEscomptes) then
                      if GereEscompte then
                        FlagOrigEsc(TOBO2) ;
                    TOBO2.Free ;
                    END ;
                 TOBEsc.Free ;
                 END ;
              END else
              BEGIN
              Break ; PasBon:=True ;
              END ;
           END ; {boucle k}
       OldBqe:=sBqe ;
       if PasBon then Break ;
       END ; {Boucle i}
   LesCombi.Clear ; LesCritsRupt.Clear ;
   END
else if LesCritsRupt.Count>0 then
   BEGIN
   ParmsMP.DEV.Code:=E_DEVISE.Value ; GetInfosDevise(ParmsMP.DEV) ;
   ParmsMP.CptS:=E_GENERAL.Text ;
   for k:=0 to LesCritsRupt.Count-1 do
       BEGIN
       ConstitueDest( sBQE, sCat, LesCritsRupt[k], TOBOrig, TOBDest, ParmsMP ) ;
       if VerifEcheMP(AlerteEcheMP) then
          BEGIN
          if LInsertDB(TobDest) then
             BEGIN
             TOBE:=TOBDest.Detail[0] ;
             MM:=TOBToIdent(TOBE,False) ;
             CompleteMM ( ParmsMP, SorteLettre, Spooler.Checked, XFichierSpooler.Checked, RepSpooler.Text, MM ) ;
             TOBO2:=ExtractDiscerne(TOBOrig,'','',LesCritsRupt[k],ParmsMP) ;
             TOBEsc := GenereEscompteMP( TOBO2, TOBEscomptes, TOBGHT, TOBGTVA,
                                         ESCCBTVA.Checked , JalEsc , ParmsMP ) ;
             if TOBO2<>Nil then
                BEGIN
                if LanceSaisieMP(MM,TOBO2,TOBEsc,CGenereAuto.Checked,TOBEscomptes) then
                  if GereEscompte then
                    FlagOrigEsc(TOBO2) ;
                TOBO2.Free ;
                END ;
             TOBEsc.Free ;
             END ;
          END else
          BEGIN
          Break ;
          END ;
       END ; {boucle k}
   LesCritsRupt.Clear ;
   END
else
   BEGIN
//   SoucheSpooler:=0 ;
   ParmsMP.DEV.Code:=E_DEVISE.Value ; GetInfosDevise(ParmsMP.DEV) ;
   ParmsMP.CptS:=E_GENERAL.Text ;
   ConstitueDest( '', '', '', TOBOrig, TOBDest, ParmsMP ) ;
   if VerifEcheMP(AlerteEcheMP) then
      BEGIN
      if LInsertDB(TobDest) then
         BEGIN
         TOBE:=TOBDest.Detail[0] ;
         MM:=TOBToIdent(TOBE,False) ;
         CompleteMM ( ParmsMP, SorteLettre, Spooler.Checked, XFichierSpooler.Checked, RepSpooler.Text, MM ) ;
//         MM.MSED.SessionFaite:=SessionFaite ;
//         If MM.MSED.MultiSessionEncaDeca Then M.MSED.ModeleMultiSession:=ModeleMultiSession ;
         LanceSaisieMP(MM,TOBOrig,Nil,False,TOBEscomptes) ;
         If MM.MSED.Spooler Then
           Case smp Of
             smpDecChqEdt,smpDecChqEdtNC : SetparamSoc('SO_CPCHEMINCHEQUE',MM.MSED.RepSpooler) ;
             smpEncTraEdt,smpEncTraEdtNC : SetparamSoc('SO_CPCHEMINTRAITE',MM.MSED.RepSpooler) ;
             smpDecBorEdt,smpDecBorEdtNC : SetparamSoc('SO_CPCHEMINBOR',MM.MSED.RepSpooler) ;
             smpDecVirEdt,smpDecVirEdtNC : SetparamSoc('SO_CPCHEMINVIREMENT',MM.MSED.RepSpooler) ;
             smpDecVirInEdt,smpDecVirInEdtNC : SetparamSoc('SO_CHEMINVIRIN',MM.MSED.RepSpooler) ; // VL 301003 FQ 12958
             smpEncPreEdt :  SetparamSoc('SO_CPCHEMINPRELEVEMENT',MM.MSED.RepSpooler) ;
             End ;
         END ;
      END ;
   END ;
// bpy le 21/07/2004 => Fiche 12340 : gestion des ecart de change !
for i:=0 to TOBOrig.Detail.Count-1 do
begin
    if (not (TOBOrig.Detail[i].GetValue('E_DEVISE') = V_PGI.DevisePivot)) then
    begin
        PGIBox('Attention certaines écritures sont en devise.' + #10 + #13 + 'Vous devez comptabiliser les différences de change.');
        RegulLettrageMP(False,False,prCLient);
        break;
    end;
end;
// Fin BPY
If PbJal Then BEGIN HM.Execute(5,gszCaption,'') ; HM.Execute(6,gszCaption,'') ; END ;
If CarePDF Then V_PGI.QRPDF:=FALSE ;
BChercheClick(Nil) ;
end;

procedure TFMulGenereMP.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
    V_PGI.MiseSousPli := saveCodeOmr;

  // Libération des TOB
  TOBOrig.Free ;
  TOBDest.Free ;
  TOBEscomptes.Free ;
  TOBGHT.Free ;
  TOBGTVA.Free ;
  // Libération des TList
  LesCombi.Clear ;
  LesCombi.Free ;
  LesCritsRupt.Clear ;
  LesCritsRupt.Free ;

  NivConcept[ccActionsFiltre] := SauveConceptFiltre ;

end;

procedure TFMulGenereMP.BCtrlRibClick(Sender: TObject);
Var
  StWRib : String ;
  i : Integer;
begin
  inherited;
  StWRib := RecupWhereCritere(Pages) ;
  If (StWRib = '') Then Exit;
//  if not SwapSelect then begin  // Si on n'est pas en Sélection inversée
    if ((Not FListe.AllSelected) and (FListe.NbSelected>0) and (FListe.NbSelected<100)) then begin  // Si on n'a pas tous sélectionné ET qu'il y a au moins 1 et 100 au plus lignes sélectionnées
      // Rajoute une clause au WHERE
      StWRib := StWRib+' AND (';
      for i:=0 to FListe.NbSelected-1 do begin
        FListe.GotoLeBookmark(i) ;
        StWRib := StWRib +' (E_NUMEROPIECE='+ Q.FindField('E_NUMEROPIECE').AsString +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
      end;
      // Efface le dernier OR et rajoute ')'
      delete(StWRib,length(StWRib)-2,3);
      StWRib := StWRib +')';
    end;
{
  // Ce code ne fonctionne pas
  end
  else begin // On est en mode inversé
    if FListe.AllSelected then exit;  // Si tous sélectionné : Sort
    // Rajoute une clause au WHERE
    if (FListe.NbSelected>0) then begin
      szSwap := ' AND (';
      FListe.Visible := False; Q.First;
      while not Q.EOF do begin
        if not FListe.IsCurrentSelected then
          szSwap := szSwap +' (E_NUMEROPIECE='+  +' AND E_NUMLIGNE='+ Q.FindField('E_NUMLIGNE').AsString +' AND E_JOURNAL="'+ Q.FindField('E_JOURNAL').AsString +'") OR';
        Q.Next;
      end;
      // Efface le dernier OR et rajoute ')'
      delete(szSwap,length(szSwap)-2,3);
      StWRib := StWRib + szSwap+')';
      Q.First; FListe.Visible := True;
    end;
  end;
}
  If StWRib<>'' Then CPLanceFiche_VerifRib('WHERE='+StWRib);
end;

procedure TFMulGenereMP.FListeDrawDataCell(Sender: TObject; const Rect: TRect; Field: TField; State: TGridDrawState);
Var sRIB,St : String ;
//    RJal,RExo,RQualif : String ;
//    RDate : TDateTime ;
//    RNumP,RNumL,RNumEche : Integer ;
//    OkLocate : Boolean ;
    TOBE     : TOB ;
    TheColor : TColor ;
    DD : TDateTime ;
begin
  inherited;
if ((Q.EOF) and (Q.BOF)) then Exit ;
If FTid.Checked Then St:='E_GENERAL' Else St:='E_AUXILIAIRE' ;
if ((Field.FieldName=St) and (Q.FindField('E_RIB')<>Nil)) then
   BEGIN
   sRIB:=Q.FindField('E_RIB').AsString ; if sRib<>'' then Exit ;
   FListe.Canvas.Brush.Color:= clRed ; FListe.Canvas.Brush.Style:=bsSolid ; FListe.Canvas.Pen.Color:=clRed ;
   FListe.Canvas.Pen.Mode:=pmCopy ; FListe.Canvas.Pen.Width:= 1 ;
   FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
   END ;
// Uniquement pour les Chèques fournisseurs : Escompte par ligne
if (Field.FieldName='E_NUMEROPIECE') and (smp in [smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC]) then
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
//if ((Field.FieldName='E_DATECOMPTABLE') and (Q.FindField('E_DATECOMPTABLE')<>Nil)) then
if ((Field.FieldName='E_DATEECHEANCE') and (Q.FindField('E_DATEECHEANCE')<>Nil)) then
   BEGIN
   (*
   TOBE:=FindCreerTOBEscompte(True) ;
   if TOBE<>Nil then if TOBE.GetValue('ESCOMPTABLE')='X' then
   *)
   If OkLignePourEscompte Then
      BEGIN
      TheColor:=clYellow ;
      If ESCCBProrata.Checked And (Q.FindField('E_DATEECHEANCE')<>NIL) Then
        BEGIN
        DD:=Q.FindField('E_DATEECHEANCE').AsDateTime ;
        If Arrondi(DD-Date,0)<=0 Then TheColor:=clred ;
        END ;
      FListe.Canvas.Brush.Color:= TheColor ; FListe.Canvas.Brush.Style:=bsSolid ; FListe.Canvas.Pen.Color:=clGreen ;
      FListe.Canvas.Pen.Mode:=pmCopy ; FListe.Canvas.Pen.Width:= 1 ;
      FListe.Canvas.Rectangle(Rect.Right-5,Rect.Top+1,Rect.Right-1,Rect.Top+5);
      END ;
   END;
end;

procedure TFMulGenereMP._chk_RegSupClick(Sender: TObject);
begin
  inherited;
  if isEncMP(smp) then BEGIN
    if _chk_RegSup.State=cbGrayed    then
      _E_NATUREPIECE.Value := _E_NATUREPIECE.Value + 'RC;';
    if _chk_RegSup.State=cbChecked   then
      begin
      _E_NATUREPIECE.Value := 'RC;';
      cFactCredit.State    := cbUnChecked;
      end;
    if _chk_RegSup.State=cbUnChecked then
      Enleve('RC');
    end
  else
    begin
    if _chk_RegSup.State=cbGrayed    then
      _E_NATUREPIECE.Value := _E_NATUREPIECE.Value + 'RF;';
    if _chk_RegSup.State=cbChecked   then
      begin
      _E_NATUREPIECE.Value := 'RF;';
      cFactCredit.State := cbUnChecked;
      end;
    if _chk_RegSup.State=cbUnChecked then
      Enleve('RF');
    end;
  RempliWhereEnc;
end;

procedure TFMulGenereMP.cFactCreditClick(Sender: TObject);
begin
  inherited;
  if isEncMP(smp) then
    begin
    if cFactCredit.State=cbGrayed    then
      _E_NATUREPIECE.Value := _E_NATUREPIECE.Value + 'AC;OC;FC;';
    if cFactCredit.State=cbChecked   then
      begin
      _E_NATUREPIECE.Value  := 'AC;OC;FC;';
      _chk_RegSup.State     := cbUnChecked;
      end;
    if cFactCredit.State=cbUnChecked then
      Enleve('AC');
    end
  else
    begin
    if cFactCredit.State=cbGrayed then
      _E_NATUREPIECE.Value := 'AF;OF;FF;';
    if cFactCredit.State=cbChecked then
      begin
      _E_NATUREPIECE.Value := 'AF;OF;FF;';
      _chk_RegSup.State := cbUnChecked;
      end;
    if cFactCredit.State=cbUnChecked then
      Enleve('AF');
    end;
  RempliWhereEnc;
end;

procedure TFMulGenereMP.Enleve(pszNat: String);
var szNatPie, szTemp, szResult : String ;
begin
  szNatPie := _E_NATUREPIECE.Value;
  if (szNatPie = '') then szNatPie := 'AC;AF;ECC;FC;FF;OC;OD;OF;RC;RF;';

  while szNatPie<>'' do begin
    szTemp := ReadTokenSt(szNatPie);
    if (szTemp<>pszNat) then szResult := szResult + sztemp +';';
  end;
  _E_NATUREPIECE.Value := szResult;
end;

procedure TFMulGenereMP.RempliWhereEnc;
var
  szNatPie,szTemp : String;
BEGIN
  szNatPie := _E_NATUREPIECE.Value;
  if (szNatPie <> '') then begin
    szTemp := '(';
    while szNatPie<>'' do begin
      szTemp := szTemp + 'E_NATUREPIECE="'+ReadTokenSt(szNatPie)+'" OR ';
    end;
    Delete(szTemp,length(szTemp)-2,2);
    szTemp := szTemp + ')';

    if isEncMP(smp) then begin
      XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" OR (E_ENCAISSEMENT="DEC" AND '+szTemp+')';
      if cFactCredit.State=cbChecked then XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" AND '+szTemp;
      end
    else begin
      XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR (E_ENCAISSEMENT="ENC" AND '+szTemp+')';
      if cFactCredit.State=cbChecked   then XX_WHEREENC.Text:='E_ENCAISSEMENT="ENC" AND '+szTemp;
    end;
    end
  else begin
    XX_WHEREENC.Text:='E_ENCAISSEMENT="DEC" OR E_ENCAISSEMENT="ENC"';
  end;
end;

procedure TFMulGenereMP.BNouvRechClick(Sender: TObject);
begin
  inherited;
  CategorieChange(CATEGORIE) ;
end;

procedure TFMulGenereMP.SpoolerClick(Sender: TObject);
begin
  inherited;
  RepSpooler.Enabled      := Spooler.Checked ;
  TRepSpooler.Enabled     := Spooler.Checked ;
  XFichierSpooler.Enabled := Spooler.Checked ;
  If Spooler.Checked And (RepSpooler.text='') Then
    begin
    {$IFDEF SPEC350}
    RepSpooler.text:=ExtractFilePath(Application.EXEName)
    {$ELSE}
    Case smp Of
      smpDecChqEdt,smpDecChqEdtNC     : RepSpooler.text:=GetParamSocSecur('SO_CPCHEMINCHEQUE','') ;
      smpEncTraEdt,smpEncTraEdtNC     : RepSpooler.text:=GetParamSocSecur('SO_CPCHEMINTRAITE','') ;
      smpDecBorEdt,smpDecBorEdtNC     : RepSpooler.text:=GetParamSocSecur('SO_CPCHEMINBOR','') ;
      smpDecVirEdt,smpDecVirEdtNC     : RepSpooler.text:=GetParamSocSecur('SO_CPCHEMINVIREMENT','') ;
      smpDecVirInEdt,smpDecVirInEdtNC : RepSpooler.text:=GetParamSocSecur('SO_CHEMINVIRIN','') ; // VL 301003 FQ 12958
      smpEncPreEdt : RepSpooler.text:=GetParamSocSecur('SO_CPCHEMINPRELEVEMENT','') ;
      end ;
    {$ENDIF}
    end ;
end;

procedure TFMulGenereMP.FListeRowEnter(Sender: TObject);
begin
  inherited;
  VH^.MPModifFaite:=FALSE ;
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

procedure TFMulGenereMP.BChercheClick(Sender: TObject);
begin
  ReinitWhereNatCpt(FALSE) ;
  If (Trim(E_GENERAL.Text)<>'') or (Trim(E_AUXILIAIRE.Text)<>'') Then
    ReinitWhereNatCpt(TRUE) ;
  inherited;
  DejaOP:=False ;
  ClearEscomptes ;
  AfficheSoldeSelection ;
  FListeRowEnter(nil); // 10930
end;

procedure TFMulGenereMP.FTIDClick(Sender: TObject);
begin
  inherited;
  E_GENERAL.Text:='' ;
  If FTID.Checked Then
    BEGIN
    E_AUXILIAIRE.Enabled:=FALSE ;
    TE_AUXILIAIRE.Enabled:=FALSE ;
    E_AUXILIAIRE.Text:='' ;
    If IsEnc(smp)
      Then E_GENERAL.DataType:='TZGTID'
      Else E_GENERAL.DataType:='TZGTIC' ;
    END
  Else
    BEGIN
    E_AUXILIAIRE.Enabled:=TRUE ;
    TE_AUXILIAIRE.Enabled:=TRUE ;
    If IsEnc(smp)
      Then E_GENERAL.DataType:='TZGENCAIS'
      Else E_GENERAL.DataType:='TZGDECAIS';
    END ;
end;

Procedure TFMulGenereMP.ClearEscomptes ;
begin
  if not GereEscompte then Exit ;
  TOBEscomptes.ClearDetail ;
  TOBGHT.ClearDetail ;
  TOBGTVA.ClearDetail ;
end ;

Procedure TFMulGenereMP.SwapEscompteLigne ;
var TOBE : TOB ;
begin
  TOBE := FindCreerTOBEscompte(FALSE) ;
  if TOBE<>NIL Then
    if TOBE.GetValue('ESCOMPTABLE')='X'
      Then TOBE.PutValue('ESCOMPTABLE','-')
      Else TOBE.PutValue('ESCOMPTABLE','X') ;
end ;

Function TFMulGenereMP.FindCreerTOBEscompte ( QueFind : boolean ) : TOB ;
Var TOBE : TOB ;
    Jal,Exo  : String ;
    NumP,NumL,NumE : integer ;
    FF             : TField ;
BEGIN
{TOBE:=Nil ;} Result:=Nil ;
if not GereEscompte then Exit ; If Not AvecEscompte Then Exit ;
if ((Q.EOF) and (Q.BOF)) then Exit ;
FF:=Q.FindField('E_JOURNAL') ; if FF=Nil then Exit else Jal:=FF.AsString ;
Exo:=Q.FindField('E_EXERCICE').AsString ;
NumP:=Q.FindField('E_NUMEROPIECE').AsInteger ;
NumL:=Q.FindField('E_NUMLIGNE').AsInteger ;
NumE:=Q.FindField('E_NUMECHE').AsInteger ;
TOBE:=TOBEscomptes.FindFirst(['E_JOURNAL','E_EXERCICE','E_NUMEROPIECE','E_NUMLIGNE','E_NUMECHE'],[Jal,Exo,NumP,NumL,NumE],False) ;
if ((TOBE=Nil) and (Not QueFind)) then
   BEGIN
   TOBE:=TOB.Create('ESCOMPTE',TOBEscomptes,-1) ;
   {Infos ecriture}
   TOBE.AddChampSup('E_JOURNAL',False)     ; TOBE.PutValue('E_JOURNAL',Jal) ;
   TOBE.AddChampSup('E_EXERCICE',False)    ; TOBE.PutValue('E_EXERCICE',Exo) ;
   TOBE.AddChampSup('E_NUMEROPIECE',False) ; TOBE.PutValue('E_NUMEROPIECE',NumP) ;
   TOBE.AddChampSup('E_NUMLIGNE',False)    ; TOBE.PutValue('E_NUMLIGNE',NumL) ;
   TOBE.AddChampSup('E_NUMECHE',False)     ; TOBE.PutValue('E_NUMECHE',NumE) ;
   {infos Escompte}
   InitEscompteSup(TOBE) ;
   END ;
Result:=TOBE ;
END ;

Function TFMulGenereMP.OkLignePourEscompte : Boolean ;
Var TOBE : TOB ;
    Jal,Exo  : String ;
    NumP,NumL,NumE : integer ;
    FF             : TField ;
BEGIN
  Result:=FALSE ;
  if not GereEscompte then Exit ;
  If Not AvecEscompte Then Exit ;
  if ((Q.EOF) and (Q.BOF)) then Exit ;

  FF   := Q.FindField('E_JOURNAL') ;
  if FF=Nil
    then Exit
    else Jal:=FF.AsString ;

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
END ;

procedure TFMulGenereMP.AlimTobEscomptes ;
BEGIN
  if not GereEscompte then Exit ;
  TOBEscomptes.PutValue('ESCOMPTABLE',  'X') ;
  TOBEscomptes.PutValue('COMPTEHT',     ESCHT.Text) ;
  TOBEscomptes.PutValue('TAUXESC',      Valeur(ESCTAUXESC.text)) ;
  // SBO correction fiche 12360 : si "tva sur escompte" non cochée, ne doit pas appliqué de tva à l'escompte
  if ESCCBTVA.Checked then
    begin
    TOBEscomptes.PutValue('COMPTETVA',    ESCTVA.Text) ;
    TOBEscomptes.PutValue('TAUXTVA',      Valeur(ESCTAUXTVA.text)) ;
    end
  else
    begin
    TOBEscomptes.PutValue('COMPTETVA',    '' ) ;
    TOBEscomptes.PutValue('TAUXTVA',      0.00 ) ;
    end ;
  // Fin correction 12360
  TOBEscomptes.PutValue('MONTANTESC',   0.00) ;
// Ajout Champ AESCOMPTER
	TOBEscomptes.PutValue('SANSESCOMPTE', '-') ;
END ;

procedure TFMulGenereMP.TraiteEditeClick(Sender: TObject);
begin
  inherited;
  E_NUMTRAITECHQ_.Enabled:=TraiteEdite.State=cbChecked ;
  FE_NUMTRAITECHQ_.Enabled:=TraiteEdite.State=cbChecked ;
  E_NUMTRAITECHQ.Enabled:=TraiteEdite.State=cbChecked ;
  FE_NUMTRAITECHQ.Enabled:=TraiteEdite.State=cbChecked ;
  XX_WHERETRACHQ.Text:='' ;
  If TraiteEdite.State=cbGrayed Then
    BEGIN
    E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
    XX_WHERETRACHQ.Text:='' ;
    END ;
  If TraiteEdite.State=cbChecked Then
    BEGIN
    E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
    XX_WHERETRACHQ.Text:='E_NUMTRAITECHQ<>""' ;
    END ;
  If TraiteEdite.State=cbUnChecked Then
    BEGIN
    E_NUMTRAITECHQ_.Text:='' ; E_NUMTRAITECHQ.Text:='' ;
    XX_WHERETRACHQ.Text:='E_NUMTRAITECHQ=""' ;
    END ;
end;

Procedure TFMulGenereMP.ChangeModeGrille ;
Var RJal,RExo : String ;
    RDate : TDateTime ;
    RNumP,RNumL,RNumEche : Integer ;
    Pb : Boolean ;
BEGIN
  If Q.FindField('E_SAISIMP')=NIL then Exit;
  // Initialisation des variables locales
	Pb:=FALSE ;
	RDate := iDate1900 ;
  RNumP := 0 ; RNumL := 0 ; RNumEche := 0 ;
  RJal := '' ;  RExo := '' ;
  If Q.FindField('E_JOURNAL')<>NIL
    Then RJal := Q.FindField('E_JOURNAL').AsString
    Else Pb   := TRUE ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL
    Then RExo := QuelExo(DateToStr(Q.FindField('E_DATECOMPTABLE').AsDateTime))
    Else Pb   := TRUE ;
  If Q.FindField('E_DATECOMPTABLE')<>NIL
    Then RDate := Q.FindField('E_DATECOMPTABLE').AsDateTime
    Else Pb    := TRUE ;
  If Q.FindField('E_NUMEROPIECE')<>NIL
    Then RNumP := Q.FindField('E_NUMEROPIECE').AsInteger
    Else Pb    := TRUE ;
  If Q.FindField('E_NUMLIGNE')<>NIL
    Then RNumL := Q.FindField('E_NUMLIGNE').AsInteger
    Else Pb    := TRUE ;
  If Q.FindField('E_NUMECHE')<>NIL
    Then RNumEche := Q.FindField('E_NUMECHE').AsInteger
    Else Pb       := TRUE ;
  SwapModeGrid(FListe, Q) ;
  BChercheClick(NIL) ;
  If (dgEditing in Fliste.Options) Then
    Fliste.SelectedField := Q.FindField('E_SAISIMP') ;
  If Not Pb Then
    Q.Locate('E_JOURNAL;E_EXERCICE;E_DATECOMPTABLE;E_QUALIFPIECE;E_NUMEROPIECE;E_NUMLIGNE;E_NUMECHE',
              VarArrayOf([RJal,RExo,RDate,'N',RNumP,RNumL,RNumEche]), []) ;
END ;

procedure TFMulGenereMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var A : Boolean ;
begin
  inherited;
  A:=(Shift=[ssCtrl]) ;
Case Key of
   VK_F5 : BEGIN
                Key:=0 ;
//                ClickModifRib ;
                E_AUXILIAIRE.ElipsisClick(E_AUXILIAIRE);
           END ;
{AE}  69 : BEGIN // Active-désactive les escomptes sur une ligne
           If A And AvecEscompte Then
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
{^M}  77 : if A then BEGIN Key:=0 ; ChangeModeGrille ; END ;
  END ;
end;

procedure TFMulGenereMP.BLotEcrClick(Sender: TObject);
Var i : Integer ;
begin
  inherited;
  SelectOnLot := Not SelectOnLot ;
  for i:=0 to Pages.PageCount-1 do
    BEGIN
    if Pages.Pages[i]=PBanqueLot then
      Pages.Pages[i].TabVisible := SelectOnLot
    Else if Pages.Pages[i]=PGestionPDF then
           PGestionPDF.TabVisible := smp in [smpEncPreEdt,smpEncTraEdt,smpEncTraEdtNC,smpDecChqEdt,smpDecChqEdtNC,smpDecVirEdt,smpDecVirEdtNC,smpDecVirInEdt,smpDecVirInEdtNC,smpDecBorEdt,smpDecBorEdtNC]
         Else
           Pages.Pages[i].TabVisible := Not SelectOnLot ;
    FEscompteClick(NIL) ;
    END ;
  PAvance.TabVisible := TRUE ;
  PSQL.TabVisible    := TRUE ;
  If EstSerie(S3) Then
    BEGIN
    E_BANQUEPREVI.Text     := '' ;
    E_BANQUEPREVI_.Text    := '' ;
    E_BANQUEPREVI.Visible  := FALSE ;
    E_BANQUEPREVI_.Visible := FALSE ;
    TE_BANQUEPREVI.Visible := FALSE ;
    TE_BANQUEPREVI_.Visible:= FALSE ;
    CGenereAuto.Visible    := FALSE ;
    CGenereAuto.Checked    := FALSE ;
    END ;
  If SelectOnLot Then
    BEGIN
{
    SaveFiltre(NomCritReste,FFiltres,Pages) ;
    LoadFiltre(NomCritLot,FFiltres,Pages) ;
}
    Pages.ActivePage  := PBanqueLot ;
    E_GENERAL.Text    := '' ;
    E_AUXILIAIRE.Text := '' ;
    E_DEVISE.Value    := V_PGI.DevisePivot ;
    CATEGORIE.Vide    := TRUE ;
    E_MODEPAIE.Vide   := TRUE ;
    CATEGORIE.ItemIndex := 0 ;
    CategorieChange(CATEGORIE) ;
    E_MODEPAIE.ItemIndex := 0 ;
    E_EXERCICE.ItemIndex := 0 ;
    E_JOURNAL.ItemIndex  := 0 ;
    _E_NATUREPIECE.SelectAll ;
    E_DATECOMPTABLE.Text  := StDate1900 ;
    E_DATECOMPTABLE_.Text := StDate2099 ;
    E_DATEECHEANCE.Text   := StDate1900 ;
    E_DATEECHEANCE_.Text  := StDate2099 ;
    E_NUMEROPIECE.Text    := '' ;
    E_NUMEROPIECE_.Text   := '' ;
    E_ETABLISSEMENT.ItemIndex := 0 ;
    E_NUMENCADECA.Text    := '' ;
    E_DEBIT.Text          := '' ;
    E_DEBIT_.Text         := '' ;
    E_CREDIT.Text         := '' ;
    E_CREDIT_.Text        := '' ;
    E_NUMTRAITECHQ.Text   := '' ;
    E_NUMTRAITECHQ_.Text  := '' ;
    T_TABLE0.Text := '' ;
    T_TABLE1.Text := '' ;
    T_TABLE2.Text := '' ;
    T_TABLE3.Text := '' ;
    T_TABLE4.Text := '' ;
    T_TABLE5.Text := '' ;
    T_TABLE6.Text := '' ;
    T_TABLE7.Text := '' ;
    T_TABLE8.Text := '' ;
    T_TABLE9.Text := '' ;
    CFactCredit.State := cbGrayed;
    XX_WHERELOT.text := '(E_NOMLOT<>"")' ;
    if (smp=smpEncTraEdt) or (smp=smpEncTraEdtNC) Or
       (smp=smpDecBorEdt) Or (smp=smpDecBorEdtNC) Then
       TraiteEdite.State:=cbUnchecked ; TraiteEditeClick(Nil) ;
    END
  Else
    BEGIN
    SaveFiltre(NomCritLot,FFiltres,Pages) ;
    LoadFiltre(NomCritReste,FFiltres,Pages) ;
    Pages.ActivePage    := PCritere ;
    E_NOMLOT.Text       := '' ;
    E_NOMLOT_.Text      := '' ;
    E_BANQUEPREVI.Text  := '' ;
    E_BANQUEPREVI_.Text := '' ;
    XX_WHERELOT.text    := '' ;
    InitContexte ;
    END ;
end;

procedure TFMulGenereMP.ESCMETHChange(Sender: TObject);
Var Avec : Boolean ;
begin
  inherited;
  Avec := AvecEscompte ;
  ESCHT.enabled        := Avec ;
  TESCHT.enabled       := Avec ;
  ESCTAUXESC.enabled   := Avec ;
  TESCTAUXESC.enabled  := Avec ;
  ESCTVA.enabled       := Avec ;
  TESCTVA.enabled      := Avec ;
  ESCTAUXTVA.enabled   := Avec ;
  TESCTAUXTVA.enabled  := Avec ;
  ESCCBTVA.Enabled     := Avec ;
  ESCCBPRORATA.Enabled := Avec ;
  If Not Avec Then
    BEGIN
    ESCHT.Text           := '' ;
    ESCTAUXESC.Text      := '' ;
    ESCTVA.Text          := '' ;
    ESCTAUXTVA.Text      := '' ;
    ESCCBTVA.checked     := FALSE ;
    ESCCBPRORATA.checked := FALSE ;
    END ;
end;

procedure TFMulGenereMP.ESCCBTVAClick(Sender: TObject);
Var AvecTVA : Boolean ;
begin
  inherited;
  If Not AvecEscompte Then exit ;
  AvecTVA             := ESCCBTVA.Checked ;
  ESCTVA.enabled      := AvecTVA ;
  TESCTVA.enabled     := AvecTVA ;
  ESCTAUXTVA.enabled  := AvecTVA ;
  TESCTAUXTVA.enabled := AvecTVA ;
end;

Function  TFMulGenereMP.AvecEscompte : Boolean ;
BEGIN
  Result:=FALSE ;
  If (ESCMETH.Value<>'RIE') And GereEscompte Then
    Result:=TRUE ;
END ;

procedure TFMulGenereMP.FEscompteClick(Sender: TObject);
begin
  inherited;
  PEscompte.tabVisible:=FEscompte.Checked ;
  If Not FEscompte.Checked Then
    BEGIN
    ESCMETH.Value := 'RIE' ;
    ESCMETHChange(Nil) ;
    END ;
end;

procedure TFMulGenereMP.ESCCBPRORATAClick(Sender: TObject);
begin
  inherited;
  FListe.Refresh ;
end;

procedure TFMulGenereMP.FListeColEnter(Sender: TObject);
begin
  inherited;
{$IFDEF EAGLCLIENT}
  // Equivalent EAGL ????
{$ELSE}
  If (dgEditing in Fliste.Options) Then
    Fliste.SelectedField:=Q.FindField('E_SAISIMP') ;
{$ENDIF}
end;

procedure TFMulGenereMP.MarquerLigneSansEscompte ;
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

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Retourne le solde (débit positif) du cumul des lignes
Suite ........ : sélectionnées
Mots clefs ... :
*****************************************************************}
function TFMulGenereMP.CalculSoldeSelection: Double;
Var i 							: Integer ;
		debit, credit  	: Double ;
begin
	Result := 0 ;
  if (PCumul.FindChildControl('__QRYPCumul_E_DEBIT') = nil)
  	 or (PCumul.FindChildControl('__QRYPCumul_E_CREDIT') = nil) then Exit ;
  if FListe.AllSelected then	// Si tout sélectionné, on parcours le Query
    begin
	  debit  := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_DEBIT')).text ) ;
  	credit := Valeur(THNumEdit(PCumul.FindChildControl('__QRYPCumul_E_CREDIT')).text ) ;
	  Result := debit - credit ;
    end
  else	// Sinon, on parcours le Bookmark
    begin
    for i:=0 to FListe.NbSelected-1 do
      BEGIN
      FListe.GotoLeBookmark(i) ;
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

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Affiche le cumul des lignes sélectionnées dans le titre
Mots clefs ... :
*****************************************************************}
procedure TFMulGenereMP.AfficheSoldeSelection;
Var Solde : Double;
begin
	// Calcul du solde
	Solde := CalculSoldeSelection;
  // Ajout affichage du cumul
  if Solde < 0
  	then Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(-1*Solde,0,V_PGI.OkDecV,'',True) + ' C)'
  	else Caption := gszCaption + ' (Cumul de la sélection : ' + StrFMontant(Solde,0,V_PGI.OkDecV,'',True) + ' D)';
	// Affichage de mode "selection inversé" si besoin
  if SwapSelect then
  	Caption := Caption + ' [SELECTION INVERSEE] ';
	// Raffraichissement Titre
	UpdateCaption(Self) ;
end;

procedure TFMulGenereMP.FListeFlipSelection(Sender: TObject);
begin
  inherited;
	AfficheSoldeSelection ;
end;

procedure TFMulGenereMP.bSelectAllClick(Sender: TObject);
begin
  inherited;
  // Si on sélectionne tout, on annule le mode "sélection inversée"
  if FListe.AllSelected then
  	begin
    SwapSelect := False ;
    BSwapSelect.Down := False ;
		BSwapSelect.Hint := 'Activer le mode "sélection inversée"' ;
    end;
	AfficheSoldeSelection ;
end;

procedure TFMulGenereMP.BSwapSelectClick(Sender: TObject);
begin
  // Si tout sélectionné, on déselectionne tout
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

procedure TFMulGenereMP.BParamListeClick(Sender: TObject);
begin
  // Annule le CTRL+M avant de passer en paramètrage de la liste
  if (dgEditing in FListe.Options) then SwapModeGrid(FListe,Q) ;
  inherited;
end;

procedure TFMulGenereMP.ESCTVAChange(Sender: TObject);
begin
  inherited;
  ESCTAUXTVA.Enabled := ESCTVA.Text <> '' ;
  if ESCTAUXTVA.Enabled
  	then ESCTAUXTVA.Color := ClWindow
	 	else ESCTAUXTVA.Color := ClBtnFace;
end;

procedure TFMulGenereMP.ESCHTChange(Sender: TObject);
begin
  inherited;
	ESCTAUXESC.Enabled := ESCHT.Text <> '' ;
  if ESCTAUXESC.Enabled
  	then ESCTAUXESC.Color := ClWindow
 		else ESCTAUXESC.Color := ClBtnFace;
end;

procedure TFMulGenereMP._E_NATUREPIECEChange(Sender: TObject);
begin
  inherited;
  if (_E_NATUREPIECE.Value = '') then begin
    cFactCredit.State := cbGrayed;
    _chk_RegSup.State := cbGrayed;
  end;
  RempliWhereEnc;
end;

function TFMulGenereMP.LInsertDB(T: Tob): Boolean;
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

end.


