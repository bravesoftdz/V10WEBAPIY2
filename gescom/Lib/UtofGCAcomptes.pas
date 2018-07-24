unit UtofGCAcomptes;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HSysMenu
  , HCtrls, HEnt1, HMsgBox, UTOF, vierge, UTOB, AglInit, LookUp, EntGC, SaisUtil, graphics
  , grids, windows, utilPGI, ent1, FactRG, ParamSoc
  {$IFDEF EAGLCLIENT}
  , MaineAGL
  {$ELSE}
  , db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main
  {$ENDIF}
  , M3FP, Messages, FactCpta,TiersUtil,
{$IFDEF BTP}
  FactAcompte,
{$ENDIF}
	utilcb,
	HPanel,
  HTB97;

type
  TOF_GCACOMPTE = class(TOF)
  private
    Erreur, OkFermer: boolean;
    PassCompta : boolean;
    FIndexFille: integer;
    TobPiece, TobTiers, TobAcomptes: Tob;
    TOBPieceTrait : TOB;
    Fournisseur : string;
    procedure BdeleteClick (Sender : Tobject);
    procedure ChangeDateAcc (Sender : Tobject);
    function GetLibelleRegl: string;
    procedure ChangeTypeSaisie (Sender : TObject);
//    AuxiFacture : String;
  public
    Action: TActionFiche;
    IsReglement: boolean;
    // MODIF BTP
    TOBPieceRG: TOB;
    // ----
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
    procedure OnNew; override;
    procedure OnLoad; override;
  end;

  // Un onglet d'affectation des règlements
  TOnglet = class (TObject)
  	private
      TabSheet : TTabSheet;
      TOBAcomptes : TOB;
      fournisseur : string;
    	GS: Thgrid;
      Name : string;
      ARegler : Thedit;
      Total : THEdit;
      MtTotal : double;
    public
    	constructor create ;
      destructor destroy; override;
  end;

  // La liste des onglets intervenants
  TlistTabSheet = Class (TList)
  	private
      function Add(AObject: TOnglet): Integer;
      function GetItems(Indice: integer): TOnglet;
      procedure SetItems(Indice: integer; const Value: Tonglet);
    public
      constructor create;
    	destructor destroy; override;
      property Items [Indice : integer] : TOnglet read GetItems write SetItems;
      function findOnglet (NomOnglet : string ): TOnglet;
      procedure clear; override;
  end;

  TOF_GCACOMPTES = class(TOF)
  private
    fCurrentOnglet : TOnglet;
    NbCol : Integer;
    PGACOMPTES : TTabControl;
    LesColonnes: string;
    listeOnglet : TlistTabSheet;
//    GS: THGRID;
    ColMontantdev, colRgt, ColDispo, ColJal, ColModR, ColLib, ColAffected: integer;
    TobPiece, TobTiers, TobAcomptes, TobCompta, TobAcomptesOrigine: Tob;
    (*Total,*) OldMontant: double;
    TOBPIECETRAIT : TOB;
    // MODIF BTP
    ModeGestion: T_TraitAcompte;
    TOBPieceRG: TOB;
    BTotalite, BProrata: TRadioButton;
    Titre: string;
    TransformationPiece : boolean;
    PassCompta : boolean;
    TOBPieceDetail : TOB;
    // ----
    procedure GSEnter(Sender: TObject);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSLigneDClick(Sender: TObject);
    procedure DessineCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GridSetEditing;
    procedure CreerAcompte;
    procedure ChargeComptaEnt(TheOnglet : TOnglet);
    procedure BFermerClick;
    procedure GSOnBeforeFlip(Sender: TObject; ARow: Longint; var Cancel: Boolean);
    function DispoSiTransfoPiece(MontantDispo: double; TobA: Tob): double;
    {$IFDEF BTP}
    function ExisteAffectation(TOBL: TOB): boolean;
    {$ENDIF}
    procedure TotaliteClick(Sender: TObject);
    procedure ProrataClick(Sender: TObject);
    procedure CalculeMontantTotal ( TheOnglet : Tonglet);
    procedure TraitementAcomptes;
    procedure MultiChargeCompta;
    procedure GetReglementsFromCompta;
    procedure SetMtAcomptes;
    function FindLaPiece (TOBA : TOB) : TOB;
    procedure DefiniAcomptesGlobaux;
    //
    procedure ConstitueTabSheets(TOBPIECETRAIT : TOB);
    procedure ConfigureGrillesSaisie;
    procedure DefiniTOBsSaisie;
    procedure BeforeChangeVue(Sender: TObject; var AllowChange: Boolean);
    procedure ChangeVue(Sender: Tobject);
    procedure SetEventsGrid(GS: THgrid; Etat: boolean);
    procedure GetAcomptesFromgene(TheOnglet : TOnglet);
    procedure AfficheGrid(TheOnglet: Tonglet);
    function AcomptesSaisie : boolean;
    function ControleLesSaisies: boolean;
    procedure RecomposeAcomptes;
    procedure AffichelesGrids;
    procedure ConstituelaTOBlocale(fCurrentOnglet : TOnglet; LaTOBLocale: TOB; Position : integer);
    procedure RecupAcomptesFromTOBlocale(fCurrentOnglet: TOnglet; LaTOBLocale: TOB; Ligne : integer);
    procedure CalculelesMontants;
    procedure ChargeLesComptas;
    procedure ChargeComptaSST(TheOnglet: Tonglet);
    procedure PositionneLesSelections;
    procedure PositionneSelected(TT: Tonglet);
    function GetmontantPorts (TOBpiece : TOB) : Double;
  public
    Action: TActionFiche;
    IsReglement, EnErreur: boolean;
    SANSCOMPTA: boolean;
    VisuSelection: boolean;
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnCancel; override;
    procedure OnClose; override;
  end;

  // Modif BTP
procedure MajMontants(TobPiece, TobA: Tob);
function NewAcompte(TOBTiers, TOBPiece, TOBAcompte: TOB): boolean;
function ModifAcompte(TOBTiers, TOBLPiece: TOB; var TOBregl: TOB; Action: string): boolean;
procedure ConsultAcompte(TOBTiers, TOBLPiece: TOB; TOBregl: TOB; Action: string);
procedure SuprimeAcompte(TOBPIece, TOBRegl: TOB);
// ---
function VerifDateExpireCB(dateExpire: string): boolean;

implementation
uses BTPUtil,UCotraitance,uEntCommun,FactTob,UtilTOBPiece;

// Modif BTP

function NewAcompte(TOBTiers, TOBPiece, TOBAcompte: TOB): boolean;
var TOBPASS, TOBC, TOBP, TOBACC: TOB;
begin
  Result := false;
  TOBPass := TOB.create('', nil, -1);
  TOBC := TOB.Create('TIERS', TOBPAss, -1);
  TOBC.Dupliquer(TOBTiers, true, true);
  TOBP := TOB.Create('PIECE', TOBC, -1);
  TOBP.Dupliquer(TOBPIece, true, true);
  TOBACC := TOB.Create('ACOMPTES', TOBP, -1);
  try
    TheTob := TOBPass;
    AGLLanceFiche('BTP', 'BTACOMPTE', '', '', 'ACTION=CREATION;ISREGLEMENT=X');
    TobAcc := TheTob;
    TheTob := nil;
    if TobAcc <> nil then
    begin
      TobAcc.addchampsup('MONTANTDISPO', false);
      TobAcc.putValue('MONTANTDISPO', TobAcc.GetValue('GAC_MONTANTDEV'));
      TobAcc.ChangeParent(TobAcompte, -1);
    end;
    if TOBAcompte.detail.count > 0 then result := true;
  finally
    TOBPASS.free;
  end;
end;

function ModifAcompte(TOBTiers, TOBLPiece: TOB; var TOBregl: TOB; Action: string): boolean;
var TOBPASS, TOBC, TOBP, TOBACC, TOBDACC: TOB;
begin
  Result := false;
  TOBPass := TOB.create('', nil, -1);
  TOBC := TOB.Create('TIERS', TOBPAss, -1);
  TOBC.Dupliquer(TOBTiers, true, true);
  TOBP := TOB.Create('PIECE', TOBC, -1);
  TOBP.Dupliquer(TOBLPIece, true, true);
  TOBACC := TOB.Create('LESACOMPTES', TOBP, -1);
  TOBDACC := TOB.Create('ACOMPTES', TOBACC, -1);
  TOBDACC.dupliquer(TOBRegl, true, true);
  TOBDACC.putvalue('GAC_MONTANTDEV', TOBDACC.GetValue('MONTANTDISPO'));
  try
    TheTob := TOBPass;
    AGLLanceFiche('BTP', 'BTACOMPTE', '', '', 'ACTION=' + Action + ';LATOB=' + inttostr(TOBACC.detail[0].GetIndex));
    TobAcc := TheTob;
    TheTob := nil;
    if TobAcc <> nil then
    begin
      TOBRegl.dupliquer(TOBACC, false, true);
      TobRegl.putValue('MONTANTDISPO', TobRegl.GetValue('GAC_MONTANTDEV'));
      TobRegl.putValue('MONTANTINIT', TobRegl.GetValue('GAC_MONTANTDEV'));
      TOBACC.free;
      result := true;
    end;
  finally
    TOBPASS.free;
  end;
end;

procedure ConsultAcompte(TOBTiers, TOBLPiece: TOB; TOBregl: TOB; Action: string);
var TOBPASS, TOBC, TOBP, TOBACC: TOB;
begin
  TOBPass := TOB.create('', nil, -1);
  TOBC := TOB.Create('TIERS', TOBPAss, -1);
  TOBC.Dupliquer(TOBTiers, true, true);
  TOBP := TOB.Create('PIECE', TOBC, -1);
  TOBP.Dupliquer(TOBLPIece, true, true);
  TOBACC := TOB.Create('LESACOMPTES', TOBP, -1);
  try
    TOBACC.Dupliquer(TOBRegl, true, true);
    TheTob := TOBPass;
    AGLLanceFiche('BTP', 'BTACOMPTES', '', '', 'ACTION=CONSULTATION;SANSCOMPTA');
    TheTob := nil;
  finally
    TOBPASS.free;
  end;
end;

procedure SuprimeAcompte(TOBPIece, TOBRegl: TOB);
begin
  DetruitAcompteGCCpta(TOBpiece, TobRegl);
end;
// --
//++++++++++++++++++++++++++++++++
//  Saisie d'un acompte
//++++++++++++++++++++++++++++++++

procedure TOF_GCACOMPTE.OnArgument(Arguments: string);
var St, Critere: string;
  x: integer;
  ChampMul, ValMul: string;
  // Modif BTP
  TRG, TRD, PaieDirect : double;
  TypeFacturation : String;
  ComptaACCReglements : boolean;
  PieceComptaAccRegl : boolean;
  LibReglement : string;
  TOBPT : TOB;
  // --
begin
  TOBPieceTrait := nil;
  Fournisseur := '';
  TOBPT := nil;
  TobTiers := LaTob.Detail[0];
  TobPiece := LaTob.Detail[0].Detail[0];
  TobAcomptes := LaTob.Detail[0].Detail[0].Detail[0];
  PassCompta := true;

  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  ComptaACCReglements := GetParamSocSecur ('SO_BTCOMPTAREGL',false);
//  PieceComptaAccRegl := (GetInfoParPiece (TOBpiece.getValue('GP_NATUREPIECEG'),'GPP_TYPEPASSACC')<>'AUC');
  if (not ComptaACCReglements) {or (not PieceComptaAccRegl)} then
  begin
  	PassCompta := false;
  end;
  // Modif BTP
  TOBPieceRG := TOB(LaTob.data);
  if TOBPieceRg <> nil then TOBPieceTrait := TOB(TOBPieceRg.data);
  // Modif JT le 09/10/03 -- Gestion du tiers Facturé
//  AuxiFacture := TiersAuxiliaire(TobPiece.GetValue('GP_TIERSFACTURE'));
  //
  St := Arguments;
  Action := taModif;
  IsReglement := false;
  FIndexFille := (-1);
  repeat
    Critere := uppercase(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'ACTION' then
        begin
          if ValMul = 'CREATION' then
          begin
            Action := taCreat;
          end;
          if ValMul = 'MODIFICATION' then
          begin
            Action := taModif;
          end;
          if ValMul = 'CONSULTATION' then
          begin
            Action := taConsult;
          end;
        end
        else if ChampMul = 'LATOB' then FIndexFille := strToInt(ValMul)
        else if ChampMul = 'ISREGLEMENT' then
        begin
          IsReglement := true;
          SetControlEnabled('GAC_ISREGLEMENT', False);
        end
        else if ChampMul = 'ISACOMPTE' then
        begin
          IsReglement := false;
          SetControlEnabled('GAC_ISREGLEMENT', False);
        end
        else if ChampMul = 'TITRE' then
        begin
          Ecran.caption := ValMul;
          UpdateCaption(Ecran);
          Ecran.refresh;
        end else if ChampMul = 'FOURNISSEUR' then
        begin
					Fournisseur := ValMul;
        end;
      end;

    end;
  until Critere = '';
  if Action = taConsult then
  begin
    FicheReadOnly(Ecran);
  end else
  begin
    if Action = TaCreat then
    begin
      SetControlText('GAC_CBLIBELLE', TOBTiers.GetValue('T_LIBELLE'));
      //JT - Affiche libellé acomptes

      //fv1 : 06/01/2014 - FS#798 - DELABOUDINIERE - date par défaut lors création d'un acompte dans une piece
      //SetControlText ('GAC_DATEECR',TOBPiece.GetString('GP_DATEPIECE'));
      //SetControlText ('GAC_DATEECHEANCE',TOBPiece.GetString('GP_DATEPIECE'));
      SetControlText ('GAC_DATEECR','');
      SetControlText ('GAC_DATEECHEANCE','');

      if not IsReglement then
      begin
        LibReglement := Copy('Acompte ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) + ' Pce. '+ IntToStr(TOBPiece.GetValue('GP_NUMERO')), 1, 35);
      end else
      begin
        LibReglement := Copy('Règlement ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE')) + ' Pce. '+ IntToStr(TOBPiece.GetValue('GP_NUMERO')), 1, 35);
      end;
      SetControlText('GAC_LIBELLE', LibReglement );
      ThEdit(GetControl('GAC_DATEECR')).OnChange := ChangeDateAcc;
    end;
    {$IFDEF BTP}
    if Action = TaModif then
    begin // On autorise que la modif du montant
      THRadioGRoup(GetControl('GAC_ISREGLEMENT')).Enabled := false;
      THEdit(GetControl('GAC_MODEPAIE')).Enabled := false;
      THEdit(GetControl('GAC_DATEECR')).Enabled := false;
      THEdit(GetControl('GAC_DATEECHEANCE')).Enabled := false;
      THValComboBox(GetControl('GAC_TYPECARTE')).Enabled := false;
      THEdit(GetControl('GAC_CBLIBELLE')).Enabled := false;
      THEdit(GetControl('GAC_CBINTERNET')).Enabled := false;
      THEdit(GetControl('GAC_DATEEXPIRE')).Enabled := false;
      THEdit(GetControl('GAC_NUMCHEQUE')).Enabled := false;
      THEdit(GetControl('GAC_JALECR')).Enabled := false;
      SetFocusControl('GAC_MONTANTDEV');
    end else SetFocusControl('GAC_MODEPAIE');
    {$ELSE}
    SetFocusControl('GAC_MODEPAIE');
    {$ENDIF}
  end;
  if not PassCompta then
  begin
  	THEdit(GetControl('GAC_JALECR')).Visible := false;
    THLabel(GetControl('TJALECR')).Visible := false;
    if (Action <> taConsult) then
    begin
      TToolBarButton97(GetControl('BDELETE')).Visible := true;
      TToolBarButton97(GetControl('BDELETE')).onclick := BdeleteClick;
    end;
  end;
  if (TOBPieceTrait <> nil) and (TOBPieceTrait.detail.count > 0) then
  begin
    TOBPT := TOBPieceTrait.FindFirst(['BPE_FOURNISSEUR'],[Fournisseur],true);
    if TOBPT <> nil then
    begin
    	THNumEdit(GetControl('GP_TOTALTTCDEV')).value := TOBPT.GetDouble('BPE_MONTANTREGL');
    end;
    THNumEdit(GetControl('TOTALTTCRG')).value := 0;
  end;
  if TOBPT = nil then
  begin
  	THNumEdit(GetControl('GP_TOTALTTCDEV')).value := TobPiece.GetValue('GP_TOTALTTCDEV') - TobPiece.GetValue('GP_ACOMPTEDEV');
    // Modif BTP
    GetMontantRGReliquat(TOBPIeceRG, TRD, TRG,true);

//    GetCumulRG(TOBPIeceRG, , );
    // AVANT - Uniquement Retenue Garantie TTC
    // MAINTENANT - Toutes les retenues (y compris pairment direct des sous traitants)
    THNumEdit(GetControl('TOTALTTCRG')).value := TRD {+ GetCumulPaiementDirect (TOBPieceTrait)};
    // --
  end;
end;

procedure TOF_GCACOMPTE.OnLoad;
var RG: THRadioGroup;
begin
  inherited;
  Erreur := False;
  OkFermer := False;
  if (Action = taConsult) or (Action = TaModif) then
  begin
    OkFermer := True;
    if FIndexFille < 0 then Exit;
    TobAcomptes.detail[FIndexFille].PutEcran(Ecran);
    exit;
  end;
  if Action = taCreat then
  begin
    RG := THRadioGroup(GetControl('GAC_ISREGLEMENT'));
    if IsReglement then RG.Value := 'X' else RG.Value := '-';
  end;
	THRadioGroup(GetControl('GAC_ISREGLEMENT')).OnClick := ChangeTypeSaisie;
end;

procedure TOF_GCACOMPTE.OnNew;
begin
  inherited;
end;

procedure TOF_GCACOMPTE.OnClose;
begin
  inherited;
  if (not OkFermer) and (valeur(getcontroltext('GAC_MONTANTDEV')) <> 0) then
  begin
    if PGIAsk('Confirmez-vous l''abandon ?', Ecran.Caption) = mrNo then
    begin
      LastError := (-1);
      LastErrorMsg := '';
      Exit;
    end else
    begin
      LastError := 0;
      Exit;
    end;
  end;

  if Erreur then
  begin
    LastError := (-1);
    LastErrorMsg := '';
    Erreur := false;
    OkFermer := False;
  end;
end;

{$IFDEF AVANT_LIBVALIDATION}

procedure TOF_GCACOMPTE.OnUpdate;
var X: T_GCAcompte;
  QQ: TQuery;
  RG: THRadioGroup;
begin
  inherited;
  OkFermer := True;
  if Action = taConsult then Exit;
  LastErrorMsg := '';
  LastError := 0;
  Erreur := False;
  if GetControlText ('GAC_DATEECR')='' then
  begin
    LastError := 5;
  	gereError(5,'La date de paiement est obligatoire');
  end else if getcontroltext('GAC_LIBELLE') = '' then
  begin
    LastError := 1;
    LastErrorMsg := 'Le libellé est obligatoire';
  end
  else if getcontroltext('GAC_MODEPAIE') = '' then
  begin
    LastError := 2;
    LastErrorMsg := 'Le mode de paiement est obligatoire';
  end
  else if (getcontroltext('GAC_JALECR') = '') and (PassCompta) then
  begin
    LastError := 3;
    LastErrorMsg := 'Le journal est obligatoire';
  end
  else if valeur(getcontroltext('GAC_MONTANTDEV')) <= 0 then
  begin
    LastError := 4;
    LastErrorMsg := 'Le montant doit être supérieur à zéro';
  end
  else if (MPTocategorie(getcontroltext('GAC_MODEPAIE')) = 'CB') and (VerifNoCarteCB(getcontroltext('GAC_CBInternet')) = false) then
  begin
    LastError := 5;
    LastErrorMsg := 'Le numéro de carte n''est pas correct';
  end
  else if (MPToCategorie(getcontroltext('GAC_MODEPAIE')) = 'CB') and (VerifDateExpireCB(getcontroltext('GAC_DateExpire')) = false) then
  begin
    LastError := 6;
    LastErrorMsg := 'La date d''expiration n''est pas correcte ou la carte est expirée';
  end
    ;
  Erreur := (LastError <> 0);
  if Erreur then Exit;
  FillChar(X, Sizeof(X), #0);
  X.JalRegle := getcontroltext('GAC_JALECR');
  X.CpteRegle := '';
  QQ := OpenSql('Select J_CONTREPARTIE,J_MODESAISIE from JOURNAL where J_JOURNAL="' + X.JalRegle + '" AND (J_NATUREJAL="BQE" OR J_NATUREJAL="CAI")', TRUE);
  if not QQ.EOF then
  begin
  	if QQ.findField('J_MODESAISIE').asString <> '-' then
    begin
      LastError := 7;
      LastErrorMsg := 'Le journal doit être en mode pièce';
    end;
  	X.CpteRegle := QQ.fields[0].asstring;
  end;
  ferme(QQ);
  if LastError <> 0 then
  begin
    exit;
  end;
  if X.CpteRegle = '' then
  begin
    QQ := OpenSql('Select MP_CPTEREGLE from MODEPAIE where MP_MODEPAIE="' + getcontroltext('GAC_MODEPAIE') + '"', TRUE);
    if not QQ.EOF then X.CpteRegle := QQ.fields[0].asstring;
    ferme(QQ);
  end;
  X.Libelle := getcontroltext('GAC_LIBELLE');
  X.ModePaie := getcontroltext('GAC_MODEPAIE');
  ;
  X.Montant := valeur(getcontroltext('GAC_MONTANTDEV'));
  X.CBLibelle := getcontroltext('GAC_CBLIBELLE');
  X.CBInternet := getcontroltext('GAC_CBINTERNET');
  X.DateExpire := getcontroltext('GAC_DATEEXPIRE');
  X.TypeCarte := getcontroltext('GAC_TYPECARTE');
  X.NumCheque := getcontroltext('GAC_NUMCHEQUE');
  RG := THRadioGroup(GetControl('GAC_ISREGLEMENT'));
  X.IsReglement := (RG.value = 'X');
  X.IsContrepartie := false;
  X.CpteContre := '';
  X.LibelleContre := '';
  X.DateEche := '';
  X.DateECr := getcontroltext('GAC_DATEECR');
  X.Date
  // Modif BTP
  X.IsModif := (Action = TAModif);
  if FIndexFille >= 0 then // cas d'une modif
  begin
    X.NumEcr := TobAcomptes.detail[FIndexFille].GetValue('GAC_NUMECR');
    X.LaTobAcc := TobAcomptes.detail[FIndexFille];
  end else
  begin
    X.numecr := 0;
    X.LaTobAcc := nil;
  end;
  // --
  TheTob := EnregistreAcompte(TOBPiece, TOBTiers, X);
end;
{$ELSE} // AVANT_LIBVALIDATION

procedure TOF_GCACOMPTE.OnUpdate;
var X: T_GCAcompte;
  QQ: TQuery;
  RG: THRadioGroup;
  okDate  : Boolean;
  DateEcr : TdateTime;

  procedure GereError(Num: integer; Msg: string);
  begin
    LastError := Num;
    LastErrorMsg := TraduireMemoire(Msg);
    TToolBarButton97(GetControl('BVALIDER')).Enabled := True;
  end;

begin
  inherited;
  TToolBarButton97(GetControl('BVALIDER')).Enabled := False;
  OkFermer  := True;
  if Action = taConsult then Exit;
  LastErrorMsg := '';
  LastError := 0;
  Erreur    := False;

  //FV1 : 07/01/2014 - FS#798 - DELABOUDINIERE - date par défaut lors création d'un acompte dans une piece
  if Not IsValidDate(GetControlText ('GAC_DATEECR')) then GereError(5,'La date de paiement est obligatoire')
  else if GetControlText('GAC_LIBELLE')  = '' then GereError(1, 'Le libellé est obligatoire')
  else if GetControlText('GAC_MODEPAIE') = '' then GereError(2, 'Le mode de paiement est obligatoire')
  else if (GetControlText('GAC_JALECR')  = '') and (PassCompta) then GereError(3, 'Le journal est obligatoire')
  else if Valeur(GetControlText('GAC_MONTANTDEV')) <= 0 then GereError(4, 'Le montant doit être supérieur à zéro')
  // JT 09/10/2003, si CB, test si on doit faire le contrôle du n° et date
  else if (not GetParamSoc('SO_GCCBFACULTATIF')) and (MPTocategorie(GetControlText('GAC_MODEPAIE')) = 'CB') then
  begin
    if not VerifNoCarteCB(GetControlText('GAC_CBInternet')) then GereError(5, 'Le numéro de carte n''est pas correct')
    else if not VerifDateExpireCB(GetControlText('GAC_DateExpire')) then GereError(6, 'La date d''expiration n''est pas correcte ou la carte est expirée');
  end;
  if not PassCompta then
  begin
    if Valeur(GetControlText('GAC_MONTANTDEV')) > TOBPiece.GetValue('GP_TOTALTTCDEV') then
    begin
			GereError(8, 'Le montant de l''acompte excède celui du document');
		end;
  end;

  Erreur := (LastError <> 0);
  if Erreur then Exit;

  //FV1 : 07/01/2014 - FS#798 - DELABOUDINIERE - date par défaut lors création d'un acompte dans une piece
   SetControlText('GAC_DATEECHEANCE',GetControlText ('GAC_DATEECR'));

  FillChar(X, Sizeof(X), #0);
  if PassCompta then
  begin
    X.JalRegle := getcontroltext('GAC_JALECR');
    X.CpteRegle := '';
    QQ := OpenSql('Select J_CONTREPARTIE,J_MODESAISIE from JOURNAL where J_JOURNAL="' + X.JalRegle + '" AND (J_NATUREJAL="BQE" OR J_NATUREJAL="CAI")', TRUE);
    if not QQ.EOF then
    begin
      if QQ.findField('J_MODESAISIE').asString <> '-' then
      begin
        LastError := 7;
        LastErrorMsg := 'Le journal doit être en mode pièce';
      end;
      X.CpteRegle := QQ.fields[0].asstring;
    end;
  	ferme(QQ);
  end;
  if LastError <> 0 then
  begin
  	erreur := (LastError <> 0);
		ecran.ModalResult := MrCancel;
  	TToolBarButton97(GetControl('BVALIDER')).Enabled := true;
    exit;
  end;
  if X.CpteRegle = '' then
  begin
    QQ := OpenSql('Select MP_CPTEREGLE from MODEPAIE where MP_MODEPAIE="' + getcontroltext('GAC_MODEPAIE') + '"', TRUE);
    if not QQ.EOF then X.CpteRegle := QQ.fields[0].asstring;
    ferme(QQ);
  end;
  X.Libelle := getcontroltext('GAC_LIBELLE');
  X.ModePaie := getcontroltext('GAC_MODEPAIE');
  X.Montant := valeur(getcontroltext('GAC_MONTANTDEV'));
  X.CBLibelle := getcontroltext('GAC_CBLIBELLE');
  X.CBInternet := getcontroltext('GAC_CBINTERNET');
  X.DateExpire := getcontroltext('GAC_DATEEXPIRE');
  X.TypeCarte := getcontroltext('GAC_TYPECARTE');
  X.NumCheque := getcontroltext('GAC_NUMCHEQUE');
  RG := THRadioGroup(GetControl('GAC_ISREGLEMENT'));
  X.IsReglement := (RG.value = 'X');
  X.IsContrepartie := false;
  X.CpteContre := '';
  X.LibelleContre := '';
  X.DateECr := getcontroltext('GAC_DATEECR');
  X.DateEche := getcontroltext('GAC_DATEECHEANCE');
  // Modif JT le 09/10/03 -- Gestion du tiers Facturé
  //  X.AuxiFacture := AuxiFacture;
  // Modif BTP
  X.IsComptabilisable := PassCompta;

  X.IsModif := (Action = TAModif);
  if FIndexFille >= 0 then // cas d'une modif
  begin
    X.NumEcr := TobAcomptes.detail[FIndexFille].GetValue('GAC_NUMECR');
    X.LaTobAcc := TobAcomptes.detail[FIndexFille];
  end else
  begin
    X.numecr := 0;
    X.LaTobAcc := nil;
  end;
  // --
  TheTob := EnregistreAcompte(TOBPiece, TOBTiers, X);
  if LaTOB.FieldExists('VALIDSOLO') then LaTOB.SetString('VALIDSOLO','X'); 
end;
{$ENDIF} // AVANTLIBVALIDATION

function VerifDateExpireCB(dateExpire: string): boolean;
var Mois, Annee: string;
  i, iMois, iAnnee: integer;
  Year, Month, Day: Word;
begin
  Result := False;
  Mois := copy(DateExpire, 1, 2);
  Annee := copy(DateExpire, 4, 4);
  for i := 1 to Length(Mois) do
    if (not (Mois[i] in ['0'..'9'])) then exit;
  for i := 1 to Length(Annee) do
    if (not (Annee[i] in ['0'..'9'])) then exit;
  DecodeDate(Now, Year, Month, Day);
  iMois := strToInt(Mois);
  if (iMois < 1) or (iMois > 12) then exit;
  iAnnee := strToInt(Annee);
  if (iAnnee < 2000) or (iAnnee > 2099) then exit;
  if iAnnee < Year then exit;
  if (iannee = year) and (iMois < Month) then exit;
  result := True;
end;

//++++++++++++++++++++++++++++++++++++
//  Liste des acomptes et selection
//++++++++++++++++++++++++++++++++++++

procedure TOF_GCACOMPTES.TotaliteClick(Sender: TObject);
var TOBA,TOBP : TOB;
    i: integer;
begin
  for i := fCurrentOnglet.GS.fixedrows to fCurrentOnglet.GS.RowCount - 1 do
  begin
    TOBA := Tob(fCurrentOnglet.GS.Objects[0, i]);
    if TOBA = nil then exit;
    if TOBPieceDetail = nil then TOBP := TOBPiece
                            else TOBP := FindLaPiece (TOBA);
    if TOBP <> nil then
    begin
      if TOBA.GetValue('MONTANTDISPO') > (TOBP.getValue('GP_TOTALTTCDEV') ) then
        TobA.putValue('GAC_MONTANTDEV',TOBP.getValue('GP_TOTALTTCDEV'))
      else // déja affecté
        TobA.putValue('GAC_MONTANTDEV', TOBA.GetValue('MONTANTDISPO'));
    end else
    begin
      TobA.putValue('GAC_MONTANTDEV', TOBA.GetValue('MONTANTDISPO'));
    end;
//
    TOBA.PutLigneGrid(fCurrentOnglet.GS, i, false, false, LesColonnes);
  end;
  CalculeMontantTotal(fCurrentOnglet);
end;

procedure TOF_GCACOMPTES.ProrataClick(Sender: TObject);
var TOBA,TOBP: TOB;
  Ratio, MontantProrata: double;
  i: integer;
begin
  for i := fCurrentOnglet.GS.fixedrows to fCurrentOnglet.GS.RowCount - 1 do
  begin
    TOBA := Tob(fCurrentOnglet.GS.Objects[0, i]);
    if TOBA = nil then Exit;
//
    if TOBPieceDetail = nil then TOBP := TOBPiece
                            else TOBP := FindLaPiece (TOBA);
    if TOBP <> nil then
    begin
      Ratio := (TOBP.GetValue('AVANCSAISIE') - TOBP.GetValue('AVANCPREC')) / (TOBP.GetValue('GP_TOTALHTDEV') +
        TOBP.GetValue('GP_TOTALREMISEDEV') + TOBP.GetValue('GP_TOTALESCDEV'));
  //
      if Ratio = 0 then Ratio := 1;
      MontantProrata := TOBA.GetValue('MONTANTINIT') * Ratio;
      if MontantProrata > TOBA.GetValue('MONTANTDISPO') then MontantProrata := TOBA.GetValue('MONTANTDISPO');
      TobA.putValue('GAC_MONTANTDEV', MontantProrata);
      TOBA.PutLigneGrid(fCurrentOnglet.GS, i, false, false, LesColonnes);
    end;
  end;
  CalculeMontantTotal(fCurrentOnglet);
end;

procedure TOF_GCACOMPTES.OnArgument(Arguments: string);
var i: integer;
  St, Nam, Critere: string;
  x: integer;
  ChampMul, ValMul: string;
  depart : integer;
  TypeFacturation : string;
begin
  inherited;
	listeOnglet := TlistTabSheet.create;
  TOBPieceDetail := nil;
  TransformationPiece := false;
  St := Arguments;
  Action := taModif;
  ModeGestion := TtaNormal; // gestion par defaut atd gescom
  VisuSelection := false;
  SansCOmpta := false;
  repeat
    Critere := uppercase(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
      end else ChampMul := Critere;
      if ChampMul = 'ACTION' then
      begin
        if ValMul = 'CREATION' then
        begin
          Action := taCreat;
        end;
        if ValMul = 'MODIFICATION' then
        begin
          Action := taModif;
        end;
        if ValMul = 'CONSULTATION' then
        begin
          Action := taConsult;
        end;
      end
        {$IFDEF BTP}
      else if ChampMul = 'SANSCOMPTA' then
      begin
        SansCompta := true;
      end
      else if ChampMul = 'TRANSFOPIECE' then
      begin
        TransformationPiece := true;
      end
        {$ENDIF}
      else if ChampMul = 'ISREGLEMENT' then
      begin
        IsReglement := true;
      end
      else if ChampMul = 'ISACOMPTE' then
      begin
        IsReglement := false;
      end
      else if ChampMul = 'AVANCEMENT' then
      begin
        ModeGestion := TtaProrata;
      end
      else if ChampMul = 'DIRECTE' then
      begin
        ModeGestion := TtaReliquat;
      end
      else if ChampMul = 'MEMOIRE' then
      begin
        ModeGestion := TtaProrata;
      end
      else if ChampMul = 'CHOIXSEL' then
      begin
        VisuSelection := true;
      end
      else if ChampMul = 'LIBELLE' then
      begin
        Ecran.Caption := ValMul;
        UpdateCaption(Ecran);
        Titre := Valmul;
      end
    end;
  until Critere = '';
  TobTiers := LaTob.Detail[0];
  TobPiece := LaTob.Detail[0].Detail[0];
  if TOBpiece.data <> nil then TOBPieceDetail := TOB(TOBpiece.data);
  //
  TobAcomptes := LaTob.Detail[0].Detail[0].Detail[0];
  TobAcomptesOrigine := tob.create('Acomptes originaux', nil, -1);
  TobAcomptesOrigine.Dupliquer(TobAcomptes, True, True, True);
  //
  TobCompta := Tob.create('Ecritures acomptes', nil, -1);
  //
  PassCompta := true;
  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  (*
  if (not GetParamSoc ('SO_BTCOMPTAREGL')) or
  	 (GetInfoParPiece (TOBpiece.getValue('GP_NATUREPIECEG'),'GPP_TYPEPASSACC')='RIE') or
  	 (TOBpiece.getValue('GP_NATUREPIECEG')='DAC') or
     (TypeFacturation='DAC') then
  begin
  	PassCompta := false;
  end;
  *)
  if (not GetParamSoc ('SO_BTCOMPTAREGL')) then
  begin
  	PassCompta := false;
  end;
  // MODIF BTP
  TOBPieceRG := TOB(LaTob.data);
  if TOBPieceRG <> nil then TOBPIECETRAIT := TOB(TOBPieceRG.Data);
  // --
  {$IFDEF BTP}
  NbCol := 8;
  LesColonnes := 'FIXED;GAC_NATUREPIECEG;GAC_JALECR;GAC_MODEPAIE;GAC_LIBELLE;GAC_MONTANTDEV;MONTANTDISPO;GAC_ISREGLEMENT';
  {$ELSE}
  NbCol := 7;
  LesColonnes := 'FIXED;GAC_NATUREPIECEG;GAC_JALECR;GAC_MODEPAIE;GAC_MONTANTDEV;MONTANTDISPO;GAC_ISREGLEMENT';
  {$ENDIF}
  PGACOMPTES := TTabControl(GetControl('PGACOMPTES'));

  // Nouveauté ---> Onglets pour paiement direct sous traitant
  ConstitueTabSheets(TOBPIECETRAIT);
  // --------------
  Btotalite := TRadioButton(GetCOntrol('CTOTALITE'));
  if BTotalite <> nil then Btotalite.OnClick := TotaliteClick;
  BProrata := TRadioButton(GetCOntrol('CPRORATA'));
  if Bprorata <> nil then BProrata.OnClick := ProrataClick;
  //
  ConfigureGrillesSaisie;
  DefiniTOBsSaisie;
  //
  if BTotalite <> nil then
  begin
    if ((ModeGestion = ttaNormal) or (ModeGestion = ttAReliquat)) then
    begin
      Btotalite.Checked := true;
      if VisuSelection then
      begin
        Btotalite.visible := true;
        Bprorata.visible := true;
      end;
    end else
    begin
      Bprorata.Checked := true;
      if VisuSelection then
      begin
        Btotalite.visible := true;
        Bprorata.visible := true;
      end;
    end;
  end;

  if PassCompta then
  begin
    if TOBPieceDetail<>nil then
    begin
      MultiChargeCompta;
    end else
    begin
  	  ChargeLesComptas;
    end;
  end else
  begin
  	TraitementAcomptes;
  end;
  //
  if fCurrentOnglet.TobAcomptes.detail.count = 0 then
  begin
    CreerAcompte;
  end;
  //
  AffichelesGrids;
  PositionneLesSelections;
	CalculelesMontants;
  AffichelesGrids;
  SetEventsGrid (fCurrentOnglet.GS,true);
	PGACOMPTES.OnChange := ChangeVue;
  PGACOMPTES.OnChanging := BeforeChangeVue;
end;

procedure TOF_GCACOMPTES.TraitementAcomptes;
var Indice : integer;
		ratio,MontantProrata,MontantDispo,MontantUsed : double;
    QQ : TQuery;
    XD,XP,MontantAregler : double;
begin
  GetMontantRGReliquat(TOBPIeceRG, XD, XP,true);
  MontantAregler := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD;
  fCurrentOnglet.ARegler.text := FloatToStrF(MontantAregler, ffNumber, 15, 2);
  //
  QQ := OpenSql ('SELECT SUM(BST_MONTANTACOMPTE) AS MTUSED FROM BSITUATIONS WHERE '+
                 'BST_SSAFFAIRE="'+TOBPiece.getValue('GP_AFFAIREDEVIS')+'" AND BST_VIVANTE="X"',true);
  if not QQ.eof then MontantUsed := QQ.findField('MTUSED').AsFloat;
  ferme (QQ);

  if (Bprorata <> nil) and (Bprorata.Checked) then
  begin
    Ratio := (TobPiece.GetValue('AVANCSAISIE') - TOBPiece.GetValue('AVANCPREC')) / (TOBPiece.GetValue('GP_TOTALHTDEV') +
      				TOBPiece.GetValue('GP_TOTALREMISEDEV') + TOBPiece.GetValue('GP_TOTALESCDEV'));
    if Ratio = 0 then Ratio := 1;
  end;

  for Indice := 0 to fCurrentOnglet.TOBAcomptes.detail.count -1 do
  begin
    if not fCurrentOnglet.TobAComptes.detail[Indice].FieldExists ('MONTANTINIT') then
    begin
      fCurrentOnglet.TobAComptes.detail[Indice].AddChampSupValeur ('MONTANTINIT', 0);
      fCurrentOnglet.TobAComptes.detail[Indice].AddChampSupValeur ('MONTANTDISPO', 0);
    end;
    fCurrentOnglet.TobAComptes.detail[Indice].SetDouble ('MONTANTINIT',fCurrentOnglet.TobAComptes.detail[Indice].Getdouble('GAC_MONTANTDEV'));
    fCurrentOnglet.TobAComptes.detail[Indice].SetDouble ('MONTANTDISPO',fCurrentOnglet.TobAComptes.detail[Indice].Getdouble('GAC_MONTANTDEV')-MontantUsed);
    if (Bprorata <> nil) and (Bprorata.Checked) then
    begin
      MontantDispo := fCurrentOnglet.TobAComptes.detail[Indice].GetValue('MONTANTDISPO');
      MontantProrata := (fCurrentOnglet.TobAComptes.detail[Indice].GetValue('MONTANTINIT') * Ratio);
      if MontantProrata > MontantDispo then MontantProrata := MontantDispo;
      fCurrentOnglet.TobAComptes.detail[Indice].PutValue('GAC_MONTANTDEV',MontantProrata);
    end;
  end;
end;

procedure TOF_GCACOMPTES.ChargeLesComptas;
var II : Integer;
		TT : Tonglet;
begin
	for II := 0 to listeOnglet.Count -1 do
  begin
		TT := listeOnglet.items[II];
    if ii = 0 then ChargeComptaEnt (TT)
    					else ChargeComptaSST (TT);
  end;
end;


procedure TOF_GCACOMPTES.ChargeComptaSST (TheOnglet : Tonglet);

	function getAuxiSSTraitant (SSTrait : string) : string;
  var QQ : Tquery;
  begin
    Result := '';
  	QQ := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+SSTrait+'" AND T_NATUREAUXI="FOU"',True,1,'',true);
    if not QQ.eof then
    begin
      Result := QQ.findField('T_AUXILIAIRE').AsString;
    end;
    Ferme (QQ);
  end;

  function GetMtPaieSSTrait (TOBPIECETRAIT: TOB; SStrait : string) : double;
  var TOBT : TOB;
  begin
    result := 0;
  	TOBT := TOBPIECETRAIT.FindFirst(['BPE_FOURNISSEUR'],[SSTrait],true);
    if TOBT <> NIL then Result := TOBT.GEtDouble('BPE_MONTANTREGL');
  end;

var i,Indice: integer;
  SQL,QualifP  : string;
  Q: TQuery;
  TOBE, TOBA, TobTiersCpta : Tob;
  Achat: boolean;
  stDebit, stCredit, stCouverture, WhereNatEcr, NotWhereNatEcr: string;
  MtAcompte, MtLettre, MontantDispo, MontantInit,mtAcomptePiece,MtLettrePiece,MtCOnsomme: double;
  {$IFDEF BTP}
  MontantProrata, Ratio : double;
  {$ENDIF BTP}
  CalcDebCred,CalcDebCredDev : string;
  XX : RMVT ;
  RefCP : string;
  AuxiSSt : string;
  MontantAregler : double;
begin
  AuxiSSt := getAuxiSSTraitant (TheOnglet.fournisseur);
  MontantAregler := GetMtPaieSSTrait (TOBPIECETRAIT ,TheOnglet.fournisseur);
  if AuxiSSt = '' then Exit;
  // initalisation sur TobAcomptes
  for Indice := 0 to TheOnglet.TobAComptes.detail.count -1 do
  begin
    if not TheOnglet.TobAComptes.detail[Indice].FieldExists ('TRAITEVIACPTA') then
    begin
      TheOnglet.TobAComptes.detail[Indice].AddChampSup ('TRAITEVIACPTA',false);
    end;
    TheOnglet.TobAComptes.detail[Indice].PutValue ('TRAITEVIACPTA','-');
  end;
  TheOnglet.ARegler.text := FloatToStrF(MontantAregler, ffNumber, 15, 2);

  TobCompta.ClearDetail;
  // champs devise suivant type de saisie de la pièce
  if TobPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then
  begin
    stDebit := 'E_DEBITDEV';
    stCredit := 'E_CREDITDEV';
    stCouverture := 'E_COUVERTUREDEV';
  end
  else if TOBPiece.GetValue('GP_SAISIECONTRE') = '-' then
  begin
    stDebit := 'E_DEBITDEV';
    stCredit := 'E_CREDITDEV';
    stCouverture := 'E_COUVERTUREDEV';
  end;

  if TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH' then Achat := true else Achat := False;
  if Achat then
  begin
    WhereNatEcr := '(E_NATUREPIECE="OD")';
    NotWhereNatEcr := '(E_NATUREPIECE<>"OD")';
  end
  else
  begin
    WhereNatEcr := '(E_NATUREPIECE="OD")';
    NotWhereNatEcr := '(E_NATUREPIECE<>"OD")';
  end;
  //Modif JT 13/10/2003 - test vente ou achat
  CalcDebCred := 'E_CREDIT-E_DEBIT';
  CalcDebCredDev := 'E_CREDITDEV-E_DEBITDEV';
//JT 29/10/2003 - Recherche pièce en simu
	QualifP:='N';
  // Modif JT le 09/10/03 -- Gestion du tiers facturé
  TobTiersCpta := TOB.Create('TIERS',Nil,-1) ;
  TobTiersCpta.Dupliquer(TobTiers,False,True) ;
  RenseigneTiersFact(TOBPiece,TOBTiers,TOBTiersCpta) ;
  // Fin JT
  SQL := 'SELECT E_JOURNAL,E_NUMEROPIECE,E_AUXILIAIRE,E_NATUREPIECE,E_MODEPAIE,E_QUALIFPIECE,E_DATECOMPTABLE,'+
  		'('+ CalcDebCred +') as SOLDECR '
    + ',('+ CalcDebCredDev +') as SOLDECRDEV, '
//    + '(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE '
    + '(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE AND GAC_AUXILIAIRE=ECR.E_AUXILIAIRE ';
(* retire
    + '         and NOT(GAC_NATUREPIECEG="' + TOBPiece.GetValue('GP_NATUREPIECEG') + '" AND GAC_SOUCHE= "' + TOBPiece.GetValue('GP_SOUCHE') + '" '
    + '               AND GAC_NUMERO=' + intToStr(TOBPiece.GetValue('GP_NUMERO')) + ' AND GAC_INDICEG=' + intToStr(TOBPiece.GetValue('GP_INDICEG')) + ' ) ';
*)
    if TOBPiece.getValue('GP_NATUREPIECEG')<>'DBT' then
    begin
    SQL := SQL + '         AND (GAC_NATUREPIECEG<>"DBT")';
    end;
    SQL := SQL
    + ' ) as MONTANTACOMPTE ';
    // RAJOUT LS
(*
    if TOBPiece.getValue('GP_NATUREPIECEG')<>'DBT' then
    begin
*)
      SQL := SQL + ',(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE AND GAC_AUXILIAIRE=ECR.E_AUXILIAIRE '
      + '         and GAC_NATUREPIECEG="' + TOBPiece.GetValue('GP_NATUREPIECEG') + '" AND GAC_SOUCHE= "' + TOBPiece.GetValue('GP_SOUCHE') + '" '
      + '               AND GAC_NUMERO=' + intToStr(TOBPiece.GetValue('GP_NUMERO')) + ' AND GAC_INDICEG=' + intToStr(TOBPiece.GetValue('GP_INDICEG'))
      + ') as MONTANTACOMPTEPIECE ';
(*
    end;
*)
    SQL := SQL
    // ----
    + ', J_MODESAISIE' // AJout LS
// BRL 06/06   + ',(SELECT sum(' + CalcDebCredDev + ') FROM ECRITURE ECC WHERE ECC.E_AUXILIAIRE=ECR.E_AUXILIAIRE AND ECC.E_GENERAL=ECR.E_GENERAL '
    + ',(SELECT sum(E_COUVERTURE) FROM ECRITURE ECC WHERE ECC.E_AUXILIAIRE=ECR.E_AUXILIAIRE AND ECC.E_GENERAL=ECR.E_GENERAL '
    + '         AND ECC.E_QUALIFPIECE=ECR.E_QUALIFPIECE AND ECC.E_ECHE=ECR.E_ECHE AND ' + NotWhereNatEcr
    + ' AND ECC.E_DEVISE=ECR.E_DEVISE '
// BRL 06/06   + '         AND ECC.E_REFGESCOM="" AND ECC.E_LETTRAGE<>"" AND ECC.E_LETTRAGE=ECR.E_LETTRAGE '
    + '         AND ECC.E_LETTRAGE<>"" AND ECC.E_LETTRAGE=ECR.E_LETTRAGE '
    + ' ) as MONTANTLETTRE ';
// LS GRRRRRR
    RefCP:=TOBPiece.GetValue('GP_REFCOMPTABLE'); //reference comptable de la pièce
    if refcp <> '' then
    begin
      XX:=DecodeRefGCComptable(RefCP) ;
      XX.CodeD:=TOBPiece.GetValue('GP_DEVISE') ;
      XX.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
      SQL := SQL + ', (SELECT SUM(EEC.E_COUVERTURE) FROM ECRITURE EEC WHERE EEC.E_JOURNAL="'+XX.Jal+'" AND EEC.E_QUALIFPIECE="'+XX.Simul+'" '
                 + 'AND EEC.E_NUMEROPIECE='+IntToStr(XX.Num)+' AND EEC.E_LETTRAGE=ECR.E_LETTRAGE) AS MTLETTRAGEPIECE';
    end else
    begin
    	SQL := SQL + ',(E_COUVERTURE-E_COUVERTURE) AS MTLETTRAGEPIECE'; // la on est sur que c a zero
    end;
    SQL := SQL
// --
    + ',E_LIBELLE, E_NUMTRAITECHQ, E_DATEECHEANCE '
    + ' FROM ECRITURE ECR '
    + ' LEFT JOIN JOURNAL ON J_JOURNAL=E_JOURNAL ' // Ajout LS
    + ' WHERE E_AUXILIAIRE="' + TobTiersCpta.GetValue('T_AUXILIAIRE') + '" '
    + ' AND E_GENERAL="' + TobTiersCpta.GetValue('T_COLLECTIF') + '" '
    + ' AND E_CONTREPARTIEAUX="'+AuxiSSt +'" '
    + ' AND J_MODESAISIE="-" '
// ATTENTION le jour ou il y aura des pièces en simulation (idem CEGID)  Ben voila c'est fait (JT/PCS)
    + '       AND E_QUALIFPIECE="' + QualifP + '" AND E_ECHE="X" AND ' + WhereNatEcr + ' AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") '
    + '       AND E_DEVISE="' + TOBPiece.GetValue('GP_DEVISE') + '" ';

  // Modif BTP (consultation)
  if not SansCompta then
  begin
    Q := OpenSQL(SQL, True);
    if not Q.EOF then TOBCompta.LoadDetailDB('', '', '', Q, True, True);
    Ferme(Q);
  end;
  FreeAndNil(TobTiersCpta);

  if TOBCompta.Detail.Count > 0 then
  begin
    for i := 0 to TOBCompta.Detail.Count - 1 do
    begin
      TOBE := TOBCompta.Detail[i];
      // montant acompte deja affecté sur autre pièce
      mtAcompte := VarToDouble(TOBE.GetValue('MONTANTACOMPTE'));
      // montant acompte deja affecté par lettrage sans affectation a une pièce
      mtLettre := VarToDouble(TOBE.GetValue('MONTANTLETTRE'));
      // montant acompte affecte a la piece
      if MtAcompte > 0 then mtAcomptePiece := VarToDouble(TOBE.GetValue('MONTANTACOMPTEPIECE'))
      								 else mtAcomptePiece := 0;
      // mt de lettrage assosice a la piece
      MtLettrePiece := VarToDouble(TOBE.GetValue('MTLETTRAGEPIECE'));
      // --
      TOBA := TheOnglet.TOBAcomptes.FindFirst(['GAC_JALECR', 'GAC_NUMECR', 'GAC_AUXILIAIRE', 'GAC_MODEPAIE', 'GAC_QUALIFPIECE' ],
        [TOBE.GetValue('E_JOURNAL'), TOBE.GetValue('E_NUMEROPIECE'), TOBE.GetValue('E_AUXILIAIRE'), TOBE.GetValue('E_MODEPAIE'),TOBE.GetValue('E_QUALIFPIECE')],
        True);
{$IFDEF BTP}
      if (TOBA<>Nil) and (TransformationPiece) and (TheOnglet.TOBAcomptes.detail.count > 0) then
      begin
        // cas de la facturation a l'avancement
        MontantDispo := GetMontantReliquatAcompte (TOBPiece);
        MontantInit := GetMontantAcompteInit (TOBPiece);
      end else
      begin
{$ENDIF}
				if (MtAcompte - MtAcomptePiece) > (Mtlettre - MtLettrePiece) then
        begin
          MontantInit := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece ;
          MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece;
        end else
        begin
          MontantInit := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece ;
          MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece;
        end;
{$IFDEF BTP}
      end;
{$ENDIF}
      if MontantDispo < 0 then MontantDispo := 0;
      MontantDispo := Arrondi(Montantdispo, 6);

      if TOBA = nil then
      begin
        if (MontantDispo) > 0 then
        begin
          TOBA := Tob.create('ACOMPTES', TheOnglet.TOBAcomptes, -1);

          TobA.AddChampSupValeur ('TRAITEVIACPTA','X');
          TobA.addchampsupValeur('MONTANTINIT', MontantInit);
          TOBA.addchampsupValeur('MONTANTDISPO', 0);
          TobA.putValue('GAC_JALECR', TOBE.GetValue('E_JOURNAL'));
          TobA.putValue('GAC_NUMECR', TOBE.GetValue('E_NUMEROPIECE'));
          TobA.putValue('GAC_AUXILIAIRE', TOBE.GetValue('E_AUXILIAIRE'));
          TobA.putValue('GAC_QUALIFPIECE', TOBE.GetValue('E_QUALIFPIECE'));
          if MontantDispo > (MontantAregler ) then
            TobA.putValue('GAC_MONTANTDEV',MontantAregler)
          else // déja affecté
            TobA.putValue('GAC_MONTANTDEV', MontantDispo);

          TobA.putValue('MONTANTDISPO', MontantDispo);
          TobA.putValue('GAC_DATEECR', TOBE.GetValue('E_DATECOMPTABLE'));
          TobA.putValue('GAC_DATEECHEANCE', TOBE.GetValue('E_DATEECHEANCE'));
          TobA.putValue('GAC_MODEPAIE', TOBE.GetValue('E_MODEPAIE'));
          TobA.putValue('GAC_CBLIBELLE', '');
          TobA.putValue('GAC_CBINTERNET', '');
          TobA.putValue('GAC_DATEEXPIRE', '');
          TobA.putValue('GAC_TYPECARTE', '');
          TobA.putValue('GAC_NUMCHEQUE', '');
          TobA.SetBoolean ('GAC_ISREGLEMENT',true);
          // Modif BTP
          TobA.putValue('GAC_LIBELLE', TOBE.GetValue('E_LIBELLE'));
          TobA.putValue('GAC_NUMCHEQUE', TOBE.GetValue('E_NUMTRAITECHQ'));
          // --
//          if copy(TOBE.GetValue('E_NATUREPIECE'), 1, 1) = 'R' then TobA.putValue('GAC_ISREGLEMENT', 'X') else TobA.putValue('GAC_ISREGLEMENT', '-');
        end;
      end else
      begin
        TOBA.addchampsupValeur('MONTANTDISPO', 0);
        TobA.addchampsupValeur('MONTANTINIT', MontantInit);
        TobA.addchampsupValeur('MONTANTDISPO', MontantDispo);
        TobA.AddChampSupValeur ('TRAITEVIACPTA','-');
        if (Bprorata <> nil) and (Bprorata.Checked) then
        begin
          Ratio := (TOBPiece.GetValue('AVANCSAISIE') - TOBPiece.GetValue('AVANCPREC')) / (MontantAregler );
          if Ratio = 0 then Ratio := 1;
          MontantProrata := TOBA.GetValue('MONTANTINIT') * Ratio;
          if MontantProrata > MontantDispo then MontantProrata := MontantDispo;
          TobA.putValue('GAC_MONTANTDEV', MontantProrata);
        end else
        begin
          if TOBA.GetValue('GAC_NATUREPIECEG') = '' then // non affecté
          begin
            if MontantDispo > (MontantAregler)  then
              TobA.putValue('GAC_MONTANTDEV',MontantAregler)
            else // déja affecté
              TobA.putValue('GAC_MONTANTDEV', MontantDispo);
          end;
        end;
      end;
    end;
  end;
  // MODIF LS
  if (not TransformationPiece) or (TheOnglet.TOBAcomptes.detail.count = 0) then
  begin
    for Indice := 0 to TheOnglet.TOBAcomptes.detail.count -1 do
    begin
      if TheOnglet.TOBAcomptes.detail[Indice].GetValue('TRAITEVIACPTA')<>'X' then
      begin
        if not TheOnglet.TOBAcomptes.detail[Indice].FieldExists ('MONTANTINIT') then
        begin
          TheOnglet.TOBAcomptes.detail[Indice].addchampsupValeur('MONTANTINIT',0);
        end;
        if not TheOnglet.TOBAcomptes.detail[Indice].FieldExists ('MONTANTDISPO') then
        begin
          TheOnglet.TOBAcomptes.detail[Indice].addchampsupValeur('MONTANTDISPO',0);
        end;
        TheOnglet.TOBAcomptes.detail[Indice].PutValue('MONTANTINIT', TheOnglet.TOBAcomptes.detail[Indice].GetValue('GAC_MONTANTDEV'));
        TheOnglet.TOBAcomptes.detail[Indice].PutValue('MONTANTDISPO', TheOnglet.TOBAcomptes.detail[Indice].GetValue('GAC_MONTANTDEV'));
      end;
    end;
  end;

  // Modif BTP
  if SansCompta then
  begin
    for i := 0 to TheOnglet.TOBAcomptes.detail.count - 1 do
    begin
      TOBA := TheOnglet.TOBAcomptes.detail[i];
      TobA.addchampsupValeur('MONTANTDISPO', 0, false);
      TobA.addchampsupValeur('MONTANTINIT', TOBA.GetValue('GAC_MONTANTDEV'));
    end;
  end;
end;

procedure TOF_GCACOMPTES.ChargeComptaEnt (TheOnglet : TOnglet);
var i,Indice: integer;
  SQL,QualifP  : string;
  Q: TQuery;
  TOBE, TOBA, TobTiersCpta : Tob;
  Achat: boolean;
  stDebit, stCredit, stCouverture, WhereNatEcr, NotWhereNatEcr: string;
  MtAcompte, MtLettre, MontantDispo, MontantInit,mtAcomptePiece,MtLettrePiece,MtCOnsomme: double;
  {$IFDEF BTP}
  MontantProrata, Ratio : double;
  {$ENDIF BTP}
  CalcDebCred,CalcDebCredDev : string;
  XX : RMVT ;
  RefCP : string;
  Xd,Xp,MontantAregler : double;
begin
  if (TOBPIECETRAIT <> nil) and (TOBpieceTrait.detail.count > 0) and (ExisteReglDirect (TobPieceTrait)) then
  begin
  	MontantAregler := GetMontantEntreprisePDir (TOBPIECETRAIT);
  end else
  begin
    GetMontantRGReliquat(TOBPIeceRG, XD, XP,true);
    MontantAregler := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD;
  end;
  TheOnglet.ARegler.text := FloatToStrF(MontantAregler, ffNumber, 15, 2);
  // initalisation sur TobAcomptes
  for Indice := 0 to TheOnglet.TobAComptes.detail.count -1 do
  begin
    if not TheOnglet.TobAComptes.detail[Indice].FieldExists ('TRAITEVIACPTA') then
    begin
      TheOnglet.TobAComptes.detail[Indice].AddChampSup ('TRAITEVIACPTA',false);
    end;
    TheOnglet.TobAComptes.detail[Indice].PutValue ('TRAITEVIACPTA','-');
  end;

  TobCompta.ClearDetail;
  // champs devise suivant type de saisie de la pièce
  if TobPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then
  begin
    stDebit := 'E_DEBITDEV';
    stCredit := 'E_CREDITDEV';
    stCouverture := 'E_COUVERTUREDEV';
  end
  else if TOBPiece.GetValue('GP_SAISIECONTRE') = '-' then
  begin
    stDebit := 'E_DEBITDEV';
    stCredit := 'E_CREDITDEV';
    stCouverture := 'E_COUVERTUREDEV';
  end;
  if TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH' then Achat := true else Achat := False;
  if Achat then
  begin
    WhereNatEcr := '(E_NATUREPIECE="OF" OR E_NATUREPIECE="RF")';
    NotWhereNatEcr := '(E_NATUREPIECE<>"OF" AND E_NATUREPIECE<>"RF")';
  end
  else
  begin
    WhereNatEcr := '(E_NATUREPIECE="OC" OR E_NATUREPIECE="RC")';
    NotWhereNatEcr := '(E_NATUREPIECE<>"OC" AND E_NATUREPIECE<>"RC")';
  end;
  //Modif JT 13/10/2003 - test vente ou achat
  if TobPiece.GetValue('GP_VENTEACHAT') = 'ACH' then
  begin
    CalcDebCred := 'E_DEBIT-E_CREDIT';
    CalcDebCredDev := 'E_DEBITDEV-E_CREDITDEV';
  end else
  begin
    CalcDebCred := 'E_CREDIT-E_DEBIT';
    CalcDebCredDev := 'E_CREDITDEV-E_DEBITDEV';
  end;
//JT 29/10/2003 - Recherche pièce en simu
  if VH_GC.GCIfDefCEGID then
    QualifP:='S'
    else
    QualifP:='N';
  // Modif JT le 09/10/03 -- Gestion du tiers facturé
  TobTiersCpta := TOB.Create('TIERS',Nil,-1) ;
  TobTiersCpta.Dupliquer(TobTiers,False,True) ;
  RenseigneTiersFact(TOBPiece,TOBTiers,TOBTiersCpta) ;
  // Fin JT
  SQL := 'SELECT E_JOURNAL,E_NUMEROPIECE,E_AUXILIAIRE,E_NATUREPIECE,E_MODEPAIE,E_QUALIFPIECE,E_DATECOMPTABLE,'+
  		'('+ CalcDebCred +') as SOLDECR '
    + ',('+ CalcDebCredDev +') as SOLDECRDEV, '
//    + '(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE '
    + '(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE AND GAC_AUXILIAIRE=ECR.E_AUXILIAIRE ';
(* retire
    + '         and NOT(GAC_NATUREPIECEG="' + TOBPiece.GetValue('GP_NATUREPIECEG') + '" AND GAC_SOUCHE= "' + TOBPiece.GetValue('GP_SOUCHE') + '" '
    + '               AND GAC_NUMERO=' + intToStr(TOBPiece.GetValue('GP_NUMERO')) + ' AND GAC_INDICEG=' + intToStr(TOBPiece.GetValue('GP_INDICEG')) + ' ) ';
*)
    if TOBPiece.getValue('GP_NATUREPIECEG')<>'DBT' then
    begin
    SQL := SQL + '         AND (GAC_NATUREPIECEG<>"DBT")';
    end;
    SQL := SQL
    + ' ) as MONTANTACOMPTE ';
    // RAJOUT LS
(*
    if TOBPiece.getValue('GP_NATUREPIECEG')<>'DBT' then
    begin
*)
      SQL := SQL + ',(SELECT sum(GAC_MONTANTDEV) from ACOMPTES where GAC_JALECR=ECR.E_JOURNAL and GAC_NUMECR=ECR.E_NUMEROPIECE AND GAC_AUXILIAIRE=ECR.E_AUXILIAIRE '
      + '         and GAC_NATUREPIECEG="' + TOBPiece.GetValue('GP_NATUREPIECEG') + '" AND GAC_SOUCHE= "' + TOBPiece.GetValue('GP_SOUCHE') + '" '
      + '               AND GAC_NUMERO=' + intToStr(TOBPiece.GetValue('GP_NUMERO')) + ' AND GAC_INDICEG=' + intToStr(TOBPiece.GetValue('GP_INDICEG'))
      + ') as MONTANTACOMPTEPIECE ';
(*
    end;
*)
    SQL := SQL
    // ----
    + ', J_MODESAISIE' // AJout LS
// BRL 06/06   + ',(SELECT sum(' + CalcDebCredDev + ') FROM ECRITURE ECC WHERE ECC.E_AUXILIAIRE=ECR.E_AUXILIAIRE AND ECC.E_GENERAL=ECR.E_GENERAL '
    + ',(SELECT sum(E_COUVERTURE) FROM ECRITURE ECC WHERE ECC.E_AUXILIAIRE=ECR.E_AUXILIAIRE AND ECC.E_GENERAL=ECR.E_GENERAL '
    + '         AND ECC.E_QUALIFPIECE=ECR.E_QUALIFPIECE AND ECC.E_ECHE=ECR.E_ECHE AND ' + NotWhereNatEcr
    + ' AND ECC.E_DEVISE=ECR.E_DEVISE '
// BRL 06/06   + '         AND ECC.E_REFGESCOM="" AND ECC.E_LETTRAGE<>"" AND ECC.E_LETTRAGE=ECR.E_LETTRAGE '
    + '         AND ECC.E_LETTRAGE<>"" AND ECC.E_LETTRAGE=ECR.E_LETTRAGE '
    + ' ) as MONTANTLETTRE ';
// LS GRRRRRR
    RefCP:=TOBPiece.GetValue('GP_REFCOMPTABLE'); //reference comptable de la pièce
    if refcp <> '' then
    begin
      XX:=DecodeRefGCComptable(RefCP) ;
      XX.CodeD:=TOBPiece.GetValue('GP_DEVISE') ;
      XX.Etabl:=TOBPiece.GetValue('GP_ETABLISSEMENT') ;
      SQL := SQL + ', (SELECT SUM(EEC.E_COUVERTURE) FROM ECRITURE EEC WHERE EEC.E_JOURNAL="'+XX.Jal+'" AND EEC.E_QUALIFPIECE="'+XX.Simul+'" '
                 + 'AND EEC.E_NUMEROPIECE='+IntToStr(XX.Num)+' AND EEC.E_LETTRAGE=ECR.E_LETTRAGE) AS MTLETTRAGEPIECE';
    end else
    begin
    	SQL := SQL + ',(E_COUVERTURE-E_COUVERTURE) AS MTLETTRAGEPIECE'; // la on est sur que c a zero
    end;
    SQL := SQL
// --
    + ',E_LIBELLE, E_NUMTRAITECHQ, E_DATEECHEANCE '
    + ' FROM ECRITURE ECR '
    + ' LEFT JOIN JOURNAL ON J_JOURNAL=E_JOURNAL ' // Ajout LS
    + ' WHERE E_AUXILIAIRE="' + TobTiersCpta.GetValue('T_AUXILIAIRE') + '" '
    + ' AND E_GENERAL="' + TobTiersCpta.GetValue('T_COLLECTIF') + '" '
    + ' AND J_MODESAISIE="-" '
// ATTENTION le jour ou il y aura des pièces en simulation (idem CEGID)  Ben voila c'est fait (JT/PCS)
    + '       AND E_QUALIFPIECE="' + QualifP + '" AND E_ECHE="X" AND ' + WhereNatEcr + ' AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") '
    + '       AND E_DEVISE="' + TOBPiece.GetValue('GP_DEVISE') + '" ';

  // Modif BTP (consultation)
  if not SansCompta then
  begin
    Q := OpenSQL(SQL, True);
    if not Q.EOF then TOBCompta.LoadDetailDB('', '', '', Q, True, True);
    Ferme(Q);
  end;
  FreeAndNil(TobTiersCpta);

  if TOBCompta.Detail.Count > 0 then
  begin
    for i := 0 to TOBCompta.Detail.Count - 1 do
    begin
      TOBE := TOBCompta.Detail[i];
      // montant acompte deja affecté sur autre pièce
      mtAcompte := VarToDouble(TOBE.GetValue('MONTANTACOMPTE'));
      // montant acompte deja affecté par lettrage sans affectation a une pièce
      mtLettre := VarToDouble(TOBE.GetValue('MONTANTLETTRE'));
      // montant acompte affecte a la piece
      if MtAcompte > 0 then mtAcomptePiece := VarToDouble(TOBE.GetValue('MONTANTACOMPTEPIECE'))
      								 else mtAcomptePiece := 0;
      // mt de lettrage assosice a la piece
      MtLettrePiece := VarToDouble(TOBE.GetValue('MTLETTRAGEPIECE'));
      // --
      TOBA := TheOnglet.TOBAcomptes.FindFirst(['GAC_JALECR', 'GAC_NUMECR', 'GAC_AUXILIAIRE', 'GAC_MODEPAIE', 'GAC_QUALIFPIECE' ],
        [TOBE.GetValue('E_JOURNAL'), TOBE.GetValue('E_NUMEROPIECE'), TOBE.GetValue('E_AUXILIAIRE'), TOBE.GetValue('E_MODEPAIE'),TOBE.GetValue('E_QUALIFPIECE')],
        True);
{$IFDEF BTP}
      if (TOBA<>Nil) and (TransformationPiece) and (TheOnglet.TOBAcomptes.detail.count > 0) then
      begin
        // cas de la facturation a l'avancement
        MontantDispo := GetMontantReliquatAcompte (TOBPiece);
        MontantInit := GetMontantAcompteInit (TOBPiece);
      end else
      begin
{$ENDIF}
				if (MtAcompte - MtAcomptePiece) > (Mtlettre - MtLettrePiece) then
        begin
          MontantInit := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece ;
          MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece;
        end else
        begin
          MontantInit := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece ;
          MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece;
        end;
{$IFDEF BTP}
      end;
{$ENDIF}
      if MontantDispo < 0 then MontantDispo := 0;
      MontantDispo := Arrondi(Montantdispo, 6);

      if TOBA = nil then
      begin
        if (MontantDispo) > 0 then
        begin
          TOBA := Tob.create('ACOMPTES', TheOnglet.TOBAcomptes, -1);

          TobA.AddChampSupValeur ('TRAITEVIACPTA','X');
          TobA.addchampsupValeur('MONTANTINIT', MontantInit);
          TOBA.addchampsupValeur('MONTANTDISPO', 0);
          TobA.putValue('GAC_JALECR', TOBE.GetValue('E_JOURNAL'));
          TobA.putValue('GAC_NUMECR', TOBE.GetValue('E_NUMEROPIECE'));
          TobA.putValue('GAC_AUXILIAIRE', TOBE.GetValue('E_AUXILIAIRE'));
          TobA.putValue('GAC_QUALIFPIECE', TOBE.GetValue('E_QUALIFPIECE'));

//          if MontantDispo > (TOBPiece.getValue('GP_TOTALTTCDEV') ) then
          if MontantDispo > MontantAregler then
            TobA.putValue('GAC_MONTANTDEV',MontantAregler)
          else // déja affecté
            TobA.putValue('GAC_MONTANTDEV', MontantDispo);

          TobA.putValue('MONTANTDISPO', MontantDispo);
          TobA.putValue('GAC_DATEECR', TOBE.GetValue('E_DATECOMPTABLE'));
          TobA.putValue('GAC_DATEECHEANCE', TOBE.GetValue('E_DATEECHEANCE'));
          TobA.putValue('GAC_MODEPAIE', TOBE.GetValue('E_MODEPAIE'));
          TobA.putValue('GAC_CBLIBELLE', '');
          TobA.putValue('GAC_CBINTERNET', '');
          TobA.putValue('GAC_DATEEXPIRE', '');
          TobA.putValue('GAC_TYPECARTE', '');
          TobA.putValue('GAC_NUMCHEQUE', '');
          // Modif BTP
          TobA.putValue('GAC_LIBELLE', TOBE.GetValue('E_LIBELLE'));
          TobA.putValue('GAC_NUMCHEQUE', TOBE.GetValue('E_NUMTRAITECHQ'));
          // --
          if copy(TOBE.GetValue('E_NATUREPIECE'), 1, 1) = 'R' then TobA.putValue('GAC_ISREGLEMENT', 'X') else TobA.putValue('GAC_ISREGLEMENT', '-');
        end;
      end else
      begin
        TOBA.addchampsupValeur('MONTANTDISPO', 0);
        TobA.addchampsupValeur('MONTANTINIT', MontantInit);
        TobA.addchampsupValeur('MONTANTDISPO', MontantDispo);
        TobA.AddChampSupValeur ('TRAITEVIACPTA','-');
        if (Bprorata <> nil) and (Bprorata.Checked) then
        begin
          Ratio := (TOBPiece.GetValue('AVANCSAISIE') - TOBPiece.GetValue('AVANCPREC')) / (TOBPiece.GetValue('GP_TOTALHTDEV')- GetmontantPorts(TOBPiece) +
            TOBPiece.GetValue('GP_TOTALREMISEDEV') + TOBPiece.GetValue('GP_TOTALESCDEV'));
          if Ratio = 0 then Ratio := 1;
          MontantProrata := TOBA.GetValue('MONTANTINIT') * Ratio;
          if MontantProrata > MontantDispo then MontantProrata := MontantDispo;
          TobA.putValue('GAC_MONTANTDEV', MontantProrata);
        end else
        begin
          if TOBA.GetValue('GAC_NATUREPIECEG') = '' then // non affecté
          begin

//            if MontantDispo > (TOBPiece.getValue('GP_TOTALTTCDEV') ) then
            if MontantDispo > MontantAregler then
              TobA.putValue('GAC_MONTANTDEV',MontantAregler)
            else // déja affecté
              TobA.putValue('GAC_MONTANTDEV', MontantDispo);
          end;
        end;
      end;
    end;
  end;
  // MODIF LS
  if (not TransformationPiece) or (TheOnglet.TOBAcomptes.detail.count = 0) then
  begin
    for Indice := 0 to TheOnglet.TOBAcomptes.detail.count -1 do
    begin
      if TheOnglet.TOBAcomptes.detail[Indice].GetValue('TRAITEVIACPTA')<>'X' then
      begin
        if not TheOnglet.TOBAcomptes.detail[Indice].FieldExists ('MONTANTINIT') then
        begin
          TheOnglet.TOBAcomptes.detail[Indice].addchampsupValeur('MONTANTINIT',0);
        end;
        if not TheOnglet.TOBAcomptes.detail[Indice].FieldExists ('MONTANTDISPO') then
        begin
          TheOnglet.TOBAcomptes.detail[Indice].addchampsupValeur('MONTANTDISPO',0);
        end;
        TheOnglet.TOBAcomptes.detail[Indice].PutValue('MONTANTINIT', TheOnglet.TOBAcomptes.detail[Indice].GetValue('GAC_MONTANTDEV'));
        TheOnglet.TOBAcomptes.detail[Indice].PutValue('MONTANTDISPO', theOnglet.TOBAcomptes.detail[Indice].GetValue('GAC_MONTANTDEV'));
      end;
    end;
  end;

  // Modif BTP
  if SansCompta then
  begin
    for i := 0 to TheOnglet.TOBAcomptes.detail.count - 1 do
    begin
      TOBA := TheOnglet.TOBAcomptes.detail[i];
      TobA.addchampsupValeur('MONTANTDISPO', 0, false);
      TobA.addchampsupValeur('MONTANTINIT', TOBA.GetValue('GAC_MONTANTDEV'));
    end;
  end;
end;

function TOF_GCACOMPTES.DispoSiTransfoPiece(MontantDispo: double; TobA: Tob): double;
var QQ: TQUERY;
begin
  // retrouve le montant dispo original si acompte issu d'une pièce précédente.
  Result := MontantDispo;
  if (TobA.cle1 <> '') and ((TobPiece.GetValue('GP_NATUREPIECEG') <> TobA.GetValue('GAC_NATUREPIECEG'))
    or (TobPiece.GetValue('GP_SOUCHE') <> TobA.GetValue('GAC_SOUCHE'))
    or (TobPiece.GetValue('GP_NUMERO') <> TobA.GetValue('GAC_NUMERO'))
    or (TobPiece.GetValue('GP_INDICEG') <> TobA.GetValue('GAC_INDICEG'))) then
  begin
    QQ := OpenSql('SELECT GAC_MONTANTDEV from ACOMPTES where ' + TobA.cle1, True);
    if not QQ.EOF then Result := Result + QQ.Fields[0].asFloat;
    Ferme(QQ);
  end;
end;

procedure TOF_GCACOMPTES.CalculelesMontants;
var ii : Integer;
		TT : TOnglet;
begin
  for ii := 0 to listeOnglet.count -1 do
  begin
    TT := listeOnglet.Items[II];
  	CalculeMontantTotal(TT);
  end;
end;


procedure TOF_GCACOMPTES.OnLoad;
var Sortir: boolean;
  Acol, Arow: integer;
  Cancel: boolean;
begin
  inherited;
  Cancel := false;
  EnErreur := False;
  OldMontant := 0;
  Sortir := False;
  if fCurrentOnglet.TobAcomptes.detail.count = 0 then
  begin
    PostMessage(TFVierge(Ecran).Handle, WM_CLOSE, 0, 0);
    exit;
  end;
  AffichelesGrids;
end;

procedure TOF_GCACOMPTES.PositionneLesSelections;
var II : Integer;
		TT : TOnglet;
begin
	for II := 0 to listeOnglet.Count -1 do
  begin
  	TT := listeOnglet.items[II];
    PositionneSelected (TT);
  end;
end;

procedure TOF_GCACOMPTES.PositionneSelected (TT : Tonglet);
var i : Integer;
begin
  for i := TT.GS.fixedrows to TT.GS.RowCount - 1 do
  begin
    if (TT.GS.cells[ColAffected, i] <> '') (*and (not TT.GS.isSelected(i))*) then
    begin
      TT.GS.FlipSelection(i);
    end;
  end;
end;

procedure TOF_GCACOMPTES.CalculeMontantTotal( TheOnglet : Tonglet);
var i: integer;
begin
  TheOnglet.MtTotal  := 0;
  for i := TheOnglet.GS.fixedrows to TheOnglet.GS.RowCount - 1 do
  begin
    if TheOnglet.GS.isSelected(i) then
    begin
    	TheOnglet.MtTotal := TheOnglet.MtTotal + valeur(TheOnglet.GS.cells[colMontantDev, i])
    end;
  end;
  TheOnglet.Total.Text := FloatToStrF(TheOnglet.MtTotal, ffNumber, 15, 2);
  //SetControlText('TOTAL', FloatToStrF(Total, ffNumber, 15, 2));
end;

procedure TOF_GCACOMPTES.OnUpdate;
var i: integer;
begin
  inherited;
  EnErreur := False;
  // if TobAcomptes.detail.count = 0 then exit;
  if not AcomptesSaisie then exit;

  if Copy(fCurrentOnglet.GS.ColFormats[fCurrentOnglet.GS.Col], 1, 3) = 'CB=' then
  begin
    if not (fCurrentOnglet.GS.valcombo = nil) then PostMessage(fCurrentOnglet.GS.ValCombo.Handle, WM_KEYDOWN, VK_TAB, 0);
  end;
  application.processMessages;

  // verif
  if not ControleLesSaisies then Exit;
  // --
  RecomposeAcomptes;
  //------
  for i := fCurrentOnglet.GS.fixedrows to fCurrentOnglet.GS.RowCount - 1 do
  begin
    if not fCurrentOnglet.GS.isselected(i) then TOB(fCurrentOnglet.GS.Objects[0, i]).free
    else
    begin
      TOB(fCurrentOnglet.GS.Objects[0, i]).GetLigneGrid(fCurrentOnglet.GS, i, LesColonnes);
      MajMontants(TOBPiece, TOB(fCurrentOnglet.GS.Objects[0, i]));
    end;
  end;
  TobAcomptes.SetAllModifie(false);
  if LaTob.FieldExists('VALIDEOK') then  LaTob.SetString('VALIDEOK','X'); // validation a terme
end;

procedure MajMontants(TobPiece, TobA: Tob);
var Taux, MontantDev: Double;
  CodeD: string;
  LaDEV: RDEVISE;
begin
  CodeD := TOBPiece.GetValue('GP_DEVISE');
  Taux := TOBPiece.GetValue('GP_TAUXDEV');
  MontantDev := TOBA.GetValue('GAC_MONTANTDEV');
  TOBA.PutValue('GAC_NATUREPIECEG',TOBPiece.getValue('GP_NATUREPIECEG'));
  if CodeD <> V_PGI.DevisePivot then
  begin
    LaDEV.Code := CodeD;
    GetInfosDevise(LaDEV);
    if VH^.TenueEuro then
    begin
      TobA.PutValue('GAC_MONTANT', DeviseToEuro(MontantDev, Taux, LaDEV.Quotite));
    end else
    begin
      TobA.PutValue('GAC_MONTANT', DeviseToFranc(MontantDev, Taux, LaDEV.Quotite));
    end;
  end else
  begin
    if VH^.TenueEuro then
    begin
      if not (TOBPiece.GetValue('GP_SAISIECONTRE') = 'X') then
      begin
        TobA.PutValue('GAC_MONTANT', MontantDev);
      end else
      begin
        TobA.PutValue('GAC_MONTANT', FrancToEuro(MontantDev));
      end;
    end else
    begin
      if not (TOBPiece.GetValue('GP_SAISIECONTRE') = 'X') then
      begin
        TobA.PutValue('GAC_MONTANT', MontantDev);
      end else
      begin
        TobA.PutValue('GAC_MONTANT', EuroToFranc(MontantDev));
      end;
    end;
  end;
end;

procedure TOF_GCACOMPTES.OnClose;
begin
  inherited;
  if EnErreur then
  begin
    LastError := (-1);
    LastErrorMsg := '';
    exit;
  end;
	listeOnglet.free;
  (*
  If TobAcomptes.detail.count<>0 then
     begin
     PostMessage(GS.Handle, WM_KEYDOWN, VK_TAB,  0) ;
     application.processMessages;
     end;
  *)
  LastError := 0;
  if TobCompta <> nil then TobCompta.free;
  TobAcomptesOrigine.free;
end;

procedure TOF_GCACOMPTES.BFermerClick;
begin
//  if Action <> taConsult then
  begin
    {$IFDEF BTP}
    if (ModeGestion <> ttaNormal) then
    begin
      EnErreur := true;
      Exit;
    end;
    {$ENDIF}
    if TobAcomptes.detail.count <> 0 then
    begin
      PostMessage(fCurrentOnglet.GS.Handle, WM_KEYDOWN, VK_TAB, 0);
      application.processMessages;
    end;
    LastError := 0;
    EnErreur := False;
    if TobAcomptes.detail.count > 0 then
    begin
      TobAcomptes.GetGridDetail(fCurrentOnglet.GS, fCurrentOnglet.GS.rowcount - 1, 'ACOMPTES', LesColonnes);
      if (TobAcomptes.IsOneModifie(True)) then
      begin
        if PGIAsk('Voulez-vous enregistrer les modifications ?', Ecran.Caption) = mrYes then
        begin
          LastError := 0;
          OnUpdate;
          if LastError <> 0 then exit;
        end;
      end;
    end;
    TobAcomptes.ClearDetail;
    TobAcomptes.Dupliquer(TobAcomptesOrigine, True, True, True);
    LastError := 0;
  end;
end;

procedure TOF_GCACOMPTES.OnCancel;
begin
  inherited;
  TobAcomptes.ClearDetail;
  TobAcomptes.Dupliquer(TobAcomptesOrigine, True, True, True);
  OnLoad;
end;

procedure TOF_GCACOMPTES.OnNew;
begin
  inherited;
  CreerAcompte;
//  OnLoad;
end;

procedure TOF_GCACOMPTES.CreerAcompte;

  Function LeNumMaxAcompte(ToBAcompte : TOB) : integer;
  var Indice : integer;
  begin
    result := 0;
    for Indice := 0 to TOBAcompte.detail.count -1 do
    begin
      if TOBAcompte.detail[Indice].getValue('GAC_NUMECR') > result then result := TOBAcompte.detail[Indice].getValue('GAC_NUMECR');
    end;
  end;

var TOBAcc: TOB;
  StRegle: string;
  LibelleTitre: string;
begin
  inherited;
  if isReglement then StRegle := ';ISREGLEMENT=X' else StRegle := '';
  if Titre <> '' then
  begin
    LibelleTitre := ';TITRE=' + Titre;
  end else LibelleTitre := '';
  TheTob := LaTob;
  if fCurrentOnglet.fournisseur = '' then
  begin
  	AGLLanceFiche('BTP', 'BTACOMPTE', '', '', 'ACTION=CREATION' + StRegle + LibelleTitre);
  end else
  begin
    StRegle := ';ISREGLEMENT=X;FOURNISSEUR='+fCurrentOnglet.Fournisseur;
  	AGLLanceFiche('BTP', 'BTREGLSOUSTRAIT', '', '', 'ACTION=CREATION' + StRegle + LibelleTitre);
  end;
  TobAcc := TheTob;
  TheTob := nil;
  if TobAcc <> nil then
  begin
    TobAcc.addchampsup('MONTANTDISPO', false);
    TobAcc.putValue('MONTANTDISPO', TobAcc.GetValue('GAC_MONTANTDEV'));
    if not PassCompta then TobAcc.putValue('GAC_NUMECR', LeNumMaxAcompte(fCurrentOnglet.ToBAcomptes)+1);
    TobAcc.ChangeParent(fCurrentOnglet.TobAcomptes, -1);
		AfficheGrid(fCurrentOnglet);
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(fCurrentOnglet.GS);
  	CalculeMontantTotal(fCurrentOnglet);
  end;
end;

procedure TOF_GCACOMPTES.OnDelete;
begin
  inherited;
  if fCurrentOnglet.GS.Row <= 0 then Exit;
  if fCurrentOnglet.GS.RowCount <= 2 then Exit;
  fCurrentOnglet.GS.DeleteRow(fCurrentOnglet.GS.row);
end;

{$IFDEF BTP}

function TOF_GCACOMPTES.ExisteAffectation(TOBL: TOB): boolean;
var QQ: Tquery;
  REq: string;
begin
  REq := 'SELECT GAC_NUMECR FROM ACOMPTES WHERE GAC_JALECR="' + TOBL.GetValue('GAC_JALECR') + '" AND ' +
    'GAC_NUMECR="' + inttostr(TOBL.GetValue('GAC_NUMECR')) + '" AND ' +
    'GAC_NATUREPIECEG<>"' + TobPiece.GetValue('GP_NATUREPIECEG') + '" AND ' +
    'GAC_SOUCHE<>"' + TobPiece.GetValue('GP_SOUCHE') + '" AND ' +
    'GAC_NUMERO<>"' + IntToStr(TobPiece.GetValue('GP_NUMERO')) + '" AND ' +
    'GAC_INDICEG <> "' + IntToStr(TobPiece.GetValue('GP_INDICEG')) + '"';
  QQ := OpenSql(req, true);
  result := not QQ.eof;
  ferme(QQ);
end;
{$ENDIF}

procedure TOF_GCACOMPTES.ConstituelaTOBlocale (fCurrentOnglet : Tonglet ; LaTOBLocale : TOB; Position : integer);
var TOBACs,TOBAC,TOBP : TOB;
begin
	TOB.Create('', LaTOBLocale, -1);
  LaTOBLocale.Detail[0].Dupliquer(TobTiers, False, TRUE, TRUE);
  TOBP := Tob.Create('', LaTOBLocale.Detail[0], -1);
  TOBP.Dupliquer(TobPiece, False, TRUE, TRUE);
  TOBACs := Tob.Create('', TOBP, -1); // les acomptes
  TOBAC := Tob.Create('ACOMPTES', TOBACS, -1); // l acompte
  TOBAC.dupliquer(fCurrentOnglet.TOBAcomptes.detail[Position-1],True,True);
  LaTOBLocale.data := TOBpieceRG;
  if TOBPieceTrait <> nil then TOBPieceRG.Data := TOBPIECETRAIT;
end;

procedure TOF_GCACOMPTES.RecupAcomptesFromTOBlocale(fCurrentOnglet : TOnglet ; LaTOBLocale : TOB; ligne : Integer);
begin
  fCurrentOnglet.TOBAcomptes.detail[Ligne-1].Dupliquer(LaTOBLocale,True,true);
end;


procedure TOF_GCACOMPTES.GSLigneDClick(Sender: TObject);
{$IFDEF BTP}
var LAction: string;
		LaTOBLocale : TOB;
    stRegle : string;
  {$ENDIF}
begin
  inherited;
  if (fCurrentOnglet.GS.Col = colMontantDev) and (Action <> taConsult) then
  begin
    GridSetEditing;
    exit;
  end;
  LaTOBLocale := TOB.Create ('UNE TOB',nil,-1);
  ConstituelaTOBlocale (fCurrentOnglet,LaTOBLocale,fCurrentOnglet.GS.Row);
  TRY
    TheTob := LaTOBLocale;
    if (ExisteAffectation(Tob(fCurrentOnglet.GS.Objects[0, fCurrentOnglet.GS.Row]))) or (Action = TaConsult) then
    begin
      LACTION := 'ACTION=CONSULTATION'; // une affectation de document est presente sur cet acompte
    end else
    begin
      LACTION := 'ACTION=MODIFICATION';
    end;
    if fCurrentOnglet.fournisseur = '' then
    begin
      AGLLanceFiche('BTP', 'BTACOMPTE', '', '',  LACTION + ';LATOB=' +
      							 intTostr(Tob(fCurrentOnglet.GS.Objects[0, fCurrentOnglet.GS.Row]).GetIndex));
    end else
    begin
      StRegle := ';FOURNISSEUR='+fCurrentOnglet.Fournisseur;
      AGLLanceFiche('BTP', 'BTREGLSOUSTRAIT', '',  '', LACTION + ';LATOB=' +
      							intTostr(Tob(fCurrentOnglet.GS.Objects[0, fCurrentOnglet.GS.Row]).GetIndex)+StRegle);
    end;

  FINALLY
    if TheTOB <> nil then
    begin
    	RecupAcomptesFromTOBlocale(fCurrentOnglet,LaTOBLocale,fCurrentOnglet.GS.row);
    end;
  	TheTOB := nil;
    LaTOBLocale.Free;
  	OnLoad;
  end;
end;

procedure TOF_GCACOMPTES.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  if not (fCurrentOnglet.GS.valCombo = nil) then fCurrentOnglet.GS.valCombo.visible := false;
end;

procedure TOF_GCACOMPTES.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var TOBA: Tob;
begin
  if Action = taConsult then Exit;
  if ACol = ColMontantDev then
  begin
  	if TOBPiece.fieldExists ('AVANCSAISIE') then
    begin
      if (TOBPiece.GetValue('AVANCSAISIE') <> 0) and (TOBPiece.GetValue('AVANCPREC') <> 0) then
      begin
        if TobPiece.GetValue('AVANCSAISIE') - TOBPiece.GetValue('AVANCPREC') < valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow]) then
        begin
          PGIBox('l''acompte depasse le montant du document', TFVierge(ECran).caption);
          Cancel := True;
          ACol := ColMontantDev;
          Exit;
        end;
      end;
    end;
    if valeur(fCurrentOnglet.GS.cells[colDispo, ARow]) < valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow]) then
    begin
      PGIBox('Vous depassez le montant disponible', TFVierge(ECran).caption);
      Cancel := True;
      ACol := ColMontantDev;
      Exit;
    end;

    if fCurrentOnglet.GS.isselected(ARow) then
    begin
      fCurrentOnglet.MtTotal := fCurrentOnglet.MtTotal - OldMontant + valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow]);
  		fCurrentOnglet.Total.Text := FloatToStrF(fCurrentOnglet.MtTotal, ffNumber, 15, 2);
//      SetControlText('TOTAL', FloatToStrF(Total, ffNumber, 15, 2));
      oldMontant := valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow]);
      TOBA := Tob(fCurrentOnglet.GS.Objects[0, Arow]);
      if TOBA = nil then Exit;
      TobA.putValue('GAC_MONTANTDEV', valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow]));
    end;
  end;
  fCurrentOnglet.GS.options := fCurrentOnglet.GS.Options - [goEditing];
  fCurrentOnglet.GS.editormode := false;
end;

procedure TOF_GCACOMPTES.GSOnBeforeFlip(Sender: TObject; ARow: Longint; var Cancel: Boolean);
var Journal : string;
		ModeSaisie : string;
    QQ  : TQuery;
begin
  if fCurrentOnglet.GS.isselected(Arow) then
  begin
  	fCurrentOnglet.MtTotal := fCurrentOnglet.mtTotal - valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow])
  end else
  begin
  	if PassCompta then
    begin
      Journal := fCurrentOnglet.GS.cells[ColJal, ARow];
      QQ := OpenSql('Select J_MODESAISIE from JOURNAL where J_JOURNAL="' + Journal + '"', TRUE);
      if not QQ.EOF then
      begin
        ModeSaisie := QQ.findField('J_MODESAISIE').asString;
        ferme (QQ);
        if ModeSaisie <> '-' then
        begin
          PgiBox ('Le journal doit être en mode pièce',ecran.caption);
          cancel := true;
          exit;
        end;
      end else Ferme (QQ);
    end;
  	fCurrentOnglet.mtTotal := fCurrentOnglet.mtTotal + valeur(fCurrentOnglet.GS.cells[colMontantDev, ARow]);
  end;
  fCurrentOnglet.Total.Text := FloatToStrF(fCurrentOnglet.mtTotal, ffNumber, 15, 2);
//  SetControlText('TOTAL', FloatToStrF(Total, ffNumber, 15, 2));
end;

procedure TOF_GCACOMPTES.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  oldMontant := Valeur(fCurrentOnglet.GS.cells[colMontantDev, ou]);
  fCurrentOnglet.GS.InvalidateRow(ou);
end;

procedure TOF_GCACOMPTES.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  fCurrentOnglet.GS.InvalidateRow(ou);
end;

procedure TOF_GCACOMPTES.DessineCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var Triangle: array[0..2] of TPoint;
    Arect: Trect;
    TOBA : TOB;
    TheText : string;
begin
  if Arow < fCurrentOnglet.GS.Fixedrows then exit;
  if (gdFixed in AState) and (ACol = 0) then
  begin
    Arect := fCurrentOnglet.GS.CellRect(Acol, Arow);
    Canvas.Brush.Color := fCurrentOnglet.GS.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = fCurrentOnglet.GS.row) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ARect.Right - 2;
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
  (*
  if (Acol = ColJal) and (Arow >= fCurrentOnglet.GS.FixedRows) then
  begin
    ARect := fCurrentOnglet.GS.CellRect(ACol, ARow);
    if Arow > fCurrentOnglet.TobAcomptes.detail.count then exit;
    TOBA := fCurrentOnglet.TobAcomptes.detail[Arow-1]; if TOBA = nil then exit;
    TheText := rechdom('TTJALBANQUEOD',TOBA.getValue('GAC_JALECR'),false);
    Canvas.FillRect(ARect);
    fCurrentOnglet.GS.Canvas.Brush.Style := bsSolid;
    fCurrentOnglet.GS.Canvas.TextOut (((Arect.left+Arect.Right)- canvas.TextWidth(TheText)- 4 ) div 2,Arect.Top +2,TheText);
  end
  else if (Acol = ColModR) and (Arow >= fCurrentOnglet.GS.FixedRows) then
  begin
    ARect := fCurrentOnglet.GS.CellRect(ACol, ARow);
    if Arow > fCurrentOnglet.TobAcomptes.detail.count then exit;
    TOBA := fCurrentOnglet.TobAcomptes.detail[Arow-1]; if TOBA = nil then exit;
    TheText := rechdom('TTMODEPAIE',TOBA.getValue('GAC_MODEPAIE'),false);
    Canvas.FillRect(ARect);
    fCurrentOnglet.GS.Canvas.Brush.Style := bsSolid;
    fCurrentOnglet.GS.Canvas.TextOut (((Arect.left+Arect.Right)- canvas.TextWidth(TheText)- 4 ) div 2,Arect.Top +2,TheText);
  end;
  *)
end;

procedure TOF_GCACOMPTES.GridSetEditing;
begin
  if goEditing in fCurrentOnglet.GS.Options then Exit;
  with fCurrentOnglet.GS do
  begin
    options := Options + [goEditing];
    editormode := True;
    col := ColMontantDev;
    if (not isSelected(row)) then FlipSelection(row);
  end;
end;

procedure AGLGridSetEditing(parms: array of variant; nb: integer);
var F: TForm;
  ToTof: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then ToTof := TFVierge(F).LaTof else exit;
  TOF_GCACOMPTES(totof).GridSetEditing;
end;

procedure AGLACOMPTES_BFermerClick(parms: array of variant; nb: integer);
var F: TForm;
  ToTof: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge) then ToTof := TFVierge(F).LaTof else exit;
  TOF_GCACOMPTES(totof).BFermerClick;
end;

procedure TOF_GCACOMPTES.GSEnter(Sender: TObject);
var Arow, Acol: integer;
  cancel: boolean;
begin
  cancel := false;
  Acol := 1;
  Arow := 1;
  TFVierge(ecran).HMTrad.ResizeGridColumns(fCurrentOnglet.GS);
  GSRowEnter(Sender, Arow, cancel, false);
  GSCellEnter(Sender, acol, arow, cancel);
  fCurrentOnglet.GS.row := Arow;
  fCurrentOnglet.GS.col := Acol;
end;


procedure TOF_GCACOMPTE.BdeleteClick(Sender: Tobject);
begin
	if PGIAsk('Confirmez-vous la suppression ?', Ecran.Caption) = mryes then
  begin
    TobAcomptes.detail[FIndexFille].free;
    ecran.ModalResult:=mrCancel;
  end;
end;


procedure TOF_GCACOMPTES.MultiChargeCompta;
var Indice : integer;
    TOBLesCptas : TOB;
begin
  TOBLesCptas := TOB.Create ('LES ECRITURES CPTA',nil,-1);
  //
  for Indice := 0 to TobAComptes.detail.count -1 do
  begin
    if not TobAComptes.detail[Indice].FieldExists ('TRAITEVIACPTA') then
    begin
      TobAComptes.detail[Indice].AddChampSup ('TRAITEVIACPTA',false);
    end;
    TobAComptes.detail[Indice].PutValue ('TRAITEVIACPTA','-');
  end;
  //
  TRY
    GetReglementsFromCompta;
    SetMtAcomptes;
    DefiniAcomptesGlobaux;
  FINALLY
    TOBLesCptas.free;
  END;
end;



procedure TOF_GCACOMPTES.GetReglementsFromCompta;
var
  SQL,QualifP  : string;
  Q: TQuery;
  TobTiersCpta : Tob;
  Achat: boolean;
  stDebit, stCredit, stCouverture, WhereNatEcr, NotWhereNatEcr: string;
  CalcDebCred,CalcDebCredDev : string;
  XX : RMVT ;
  RefCP : string;
begin
//-----------
  TobCompta.ClearDetail;
  Achat := (TOBPiece.GetValue('GP_VENTEACHAT') = 'ACH');
  // champs devise suivant type de saisie de la pièce
  if TobPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then
  begin
    stDebit := 'E_DEBITDEV';
    stCredit := 'E_CREDITDEV';
    stCouverture := 'E_COUVERTUREDEV';
  end
  else if TOBPiece.GetValue('GP_SAISIECONTRE') = '-' then
  begin
    stDebit := 'E_DEBITDEV';
    stCredit := 'E_CREDITDEV';
    stCouverture := 'E_COUVERTUREDEV';
  end;
  if Achat then
  begin
    WhereNatEcr := '(E_NATUREPIECE="OF" OR E_NATUREPIECE="RF")';
    NotWhereNatEcr := '(E_NATUREPIECE<>"OF" AND E_NATUREPIECE<>"RF")';
  end else
  begin
    WhereNatEcr := '(E_NATUREPIECE="OC" OR E_NATUREPIECE="RC")';
    NotWhereNatEcr := '(E_NATUREPIECE<>"OC" AND E_NATUREPIECE<>"RC")';
  end;
  //Modif JT 13/10/2003 - test vente ou achat
  if Achat then
  begin
    CalcDebCred := 'E_DEBIT-E_CREDIT';
    CalcDebCredDev := 'E_DEBITDEV-E_CREDITDEV';
  end else
  begin
    CalcDebCred := 'E_CREDIT-E_DEBIT';
    CalcDebCredDev := 'E_CREDITDEV-E_DEBITDEV';
  end;
//JT 29/10/2003 - Recherche pièce en simu
  if VH_GC.GCIfDefCEGID then
    QualifP:='S'
    else
    QualifP:='N';
  // Modif JT le 09/10/03 -- Gestion du tiers facturé
  TobTiersCpta := TOB.Create('TIERS',Nil,-1) ;
  TobTiersCpta.Dupliquer(TobTiers,False,True) ;
  RenseigneTiersFact(TOBPiece,TOBTiers,TOBTiersCpta) ;
  // Fin JT
  SQL := 'SELECT E_JOURNAL,E_NUMEROPIECE,E_AUXILIAIRE,E_NATUREPIECE,E_MODEPAIE,E_QUALIFPIECE,('+ CalcDebCred +') as SOLDECR '
    + ',('+ CalcDebCredDev +') as SOLDECRDEV,'
    + '(SELECT SUM(GAC_MONTANTDEV) FROM ACOMPTES WHERE GAC_JALECR=ECR.E_JOURNAL AND '
    + 'GAC_NUMECR=ECR.E_NUMEROPIECE AND GAC_AUXILIAIRE=ECR.E_AUXILIAIRE) AS MONTANTACOMPTE'
    + ',E_LIBELLE, E_NUMTRAITECHQ '
    + ' FROM ECRITURE ECR '
    + ' LEFT JOIN JOURNAL ON J_JOURNAL=E_JOURNAL ' // Ajout LS
    + ' WHERE E_AUXILIAIRE="' + TobTiersCpta.GetValue('T_AUXILIAIRE') + '" '
    + ' AND E_GENERAL="' + TobTiersCpta.GetValue('T_COLLECTIF') + '" '
    + ' AND J_MODESAISIE="-" '
// ATTENTION le jour ou il y aura des pièces en simulation (idem CEGID)  Ben voila c'est fait (JT/PCS)
    + ' AND E_QUALIFPIECE="' + QualifP + '" AND E_ECHE="X" AND ' + WhereNatEcr + ' AND (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") '
    + ' AND E_DEVISE="' + TOBPiece.GetValue('GP_DEVISE') + '" '
    + ' ORDER BY E_JOURNAL,E_NUMEROPIECE,E_AUXILIAIRE';

  // Modif BTP (consultation)
  if not SansCompta then
  begin
    Q := OpenSQL(SQL, True);
    if not Q.EOF then TOBCompta.LoadDetailDB('', '', '', Q, True, True);
    Ferme(Q);
  end;
  FreeAndNil(TobTiersCpta);
//-------
end;

procedure TOF_GCACOMPTES.SetMtAcomptes;
var Indice : integer;
    TOBC : TOB;
    TOBAcp : TOB;
begin
  // init
  for Indice := 0 to TOBCompta.detail.count -1 do
  begin
    TOBC := TOBCompta.detail[Indice];
    TOBC.AddChampSupValeur ('MONTANTACOMPTEPIECE',0);
    TOBC.AddChampSupValeur ('MTLETTRAGEPIECE',0);
  end;
  for Indice := 0 to TobAcomptes.detail.count -1 do
  begin
    TOBAcp := TOBAcomptes.detail[Indice];
    TOBC := TOBCompta.findFirst(['E_JOURNAL','E_NUMEROPIECE','E_AUXILIAIRE'],
                                [TOBACP.getValue('GAC_JALECR'),TOBACP.getValue('GAC_NUMECR'),TOBACP.getValue('GAC_AUXILIAIRE')],True);
    repeat
      if TOBC <> nil then
      begin
        TOBC.putValue('MONTANTACOMPTEPIECE',TOBC.getValue('MONTANTACOMPTEPIECE')+TOBACP.getValue('GAC_MONTANTDEV'));
        //
        TOBC := TOBCompta.findNext(['E_JOURNAL','E_NUMEROPIECE','E_AUXILIAIRE'],
                                   [TOBACP.getValue('GAC_JALECR'),TOBACP.getValue('GAC_NUMECR'),TOBACP.getValue('GAC_AUXILIAIRE')],True);
      end;
    until TOBC = nil;
  end;
end;


function TOF_GCACOMPTES.FindLaPiece (TOBA : TOB) : TOB;
begin
  result := TOBPieceDetail.findFirst(['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                     [TOBA.GetValue('GAC_NATUREPIECEG'),TOBA.GetValue('GAC_SOUCHE'),
                                     TOBA.GetValue('GAC_NUMERO'),TOBA.GetValue('GAC_INDICEG')],True);
end;

procedure TOF_GCACOMPTES.DefiniAcomptesGlobaux;
var i : integer;
    MtAcompte, MtLettre, MontantDispo, MontantInit,mtAcomptePiece,MtLettrePiece,MtCOnsomme: double;
    TOBE,TOBA,TOBP : TOB;
    MontantProrata, Ratio : double;
    Indice : integer;
begin
//
  if TOBCompta.Detail.Count > 0 then
  begin
    for i := 0 to TOBCompta.Detail.Count - 1 do
    begin
      TOBE := TOBCompta.Detail[i];
      // montant acompte deja affecté sur autre pièce
      mtAcompte := VarToDouble(TOBE.GetValue('MONTANTACOMPTE'));
      // montant acompte deja affecté par lettrage sans affectation a une pièce
      mtLettre := VarToDouble(TOBE.GetValue('MONTANTLETTRE'));
      // montant acompte affecte a la piece
      if MtAcompte > 0 then mtAcomptePiece := VarToDouble(TOBE.GetValue('MONTANTACOMPTEPIECE'))
      								 else mtAcomptePiece := 0;
      // mt de lettrage assosice a la piece
      MtLettrePiece := VarToDouble(TOBE.GetValue('MTLETTRAGEPIECE'));
      //
      TOBA := TOBAcomptes.FindFirst(['GAC_JALECR', 'GAC_NUMECR', 'GAC_AUXILIAIRE', 'GAC_MODEPAIE', 'GAC_QUALIFPIECE' ],
        [TOBE.GetValue('E_JOURNAL'), TOBE.GetValue('E_NUMEROPIECE'), TOBE.GetValue('E_AUXILIAIRE'), TOBE.GetValue('E_MODEPAIE'),TOBE.GetValue('E_QUALIFPIECE')],
        True);
      //
      TOBP := nil;
      if TOBA <> nil then
      begin
        TOBP := FindLaPiece (TOBA);
        if TOBP = nil then
        begin
          if (MtAcompte - MtAcomptePiece) > (Mtlettre - MtLettrePiece) then
          begin
            MontantInit := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece ;
            MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece;
          end else
          begin
            MontantInit := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece ;
            MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece;
          end;
        end else
        begin
          MontantDispo := GetMontantReliquatAcompte (TOBP);
          MontantInit := GetMontantAcompteInit (TOBP);
        end;
      end else
      begin
        if (MtAcompte - MtAcomptePiece) > (Mtlettre - MtLettrePiece) then
        begin
          MontantInit := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece ;
          MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtAcompte + mtAcomptePiece;
        end else
        begin
          MontantInit := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece ;
          MontantDispo := TOBE.GetValue('SOLDECRDEV') - mtLettre + MtLettrePiece;
        end;
      end;
      if MontantDispo < 0 then MontantDispo := 0;
      MontantDispo := Arrondi(Montantdispo, 6);
//
      if TOBA = nil then
      begin
        if (MontantDispo) > 0 then
        begin
          TOBA := Tob.create('ACOMPTES', TOBAcomptes, -1);

          TobA.AddChampSupValeur ('TRAITEVIACPTA','X');
          TobA.addchampsupValeur('MONTANTINIT', MontantInit);
          TOBA.addchampsupValeur('MONTANTDISPO', MontantDispo);
          TobA.putValue('GAC_JALECR', TOBE.GetValue('E_JOURNAL'));
          TobA.putValue('GAC_NUMECR', TOBE.GetValue('E_NUMEROPIECE'));
          TobA.putValue('GAC_AUXILIAIRE', TOBE.GetValue('E_AUXILIAIRE'));
          TobA.putValue('GAC_QUALIFPIECE', TOBE.GetValue('E_QUALIFPIECE'));
          if (TOBP <> nil) and (Bprorata <> nil) and (Bprorata.Checked) then
          begin
            Ratio := (TobP.GetValue('AVANCSAISIE') - TOBP.GetValue('AVANCPREC')) / (TOBP.GetValue('GP_TOTALHTDEV') +
              TOBP.GetValue('GP_TOTALREMISEDEV') + TOBP.GetValue('GP_TOTALESCDEV'));
            if Ratio = 0 then Ratio := 1;
            MontantProrata := TOBA.GetValue('MONTANTINIT') * Ratio;
            if MontantProrata > MontantDispo then MontantProrata := MontantDispo;
            TobA.putValue('GAC_MONTANTDEV', MontantProrata);
          end else
          if (TOBP<>nil) and (MontantDispo > TOBP.getValue('GP_TOTALTTCDEV')) then
          begin
            TobA.putValue('GAC_MONTANTDEV',TOBPiece.getValue('GP_TOTALTTCDEV'))
          end else // déja affecté
          begin
            TobA.putValue('GAC_MONTANTDEV', MontantDispo);
          end;

          TobA.putValue('MONTANTDISPO', MontantDispo);
          TobA.putValue('GAC_MODEPAIE', TOBE.GetValue('E_MODEPAIE'));
          TobA.putValue('GAC_CBLIBELLE', '');
          TobA.putValue('GAC_CBINTERNET', '');
          TobA.putValue('GAC_DATEEXPIRE', '');
          TobA.putValue('GAC_TYPECARTE', '');
          TobA.putValue('GAC_NUMCHEQUE', '');
          // Modif BTP
          TobA.putValue('GAC_LIBELLE', TOBE.GetValue('E_LIBELLE'));
          TobA.putValue('GAC_NUMCHEQUE', TOBE.GetValue('E_NUMTRAITECHQ'));
          // --
          if copy(TOBE.GetValue('E_NATUREPIECE'), 1, 1) = 'R' then TobA.putValue('GAC_ISREGLEMENT', 'X') else TobA.putValue('GAC_ISREGLEMENT', '-');
        end;
      end else
      begin
        TOBA.AddChampSupValeur('MONTANTDISPO', MontantDispo);
        TOBA.AddChampSupValeur('MONTANTINIT',MontantInit);
        TobA.AddChampSupValeur ('TRAITEVIACPTA','X');
        if (TOBP<>nil) and (Bprorata <> nil) and (Bprorata.Checked) then
        begin
          Ratio := (TOBP.GetValue('AVANCSAISIE') - TOBP.GetValue('AVANCPREC')) / (TOBP.GetValue('GP_TOTALHTDEV') +
            TOBP.GetValue('GP_TOTALREMISEDEV') + TOBP.GetValue('GP_TOTALESCDEV'));
          if Ratio = 0 then Ratio := 1;
          MontantProrata := TOBA.GetValue('MONTANTINIT') * Ratio;
          if MontantProrata > MontantDispo then MontantProrata := MontantDispo;
          TobA.putValue('GAC_MONTANTDEV', MontantProrata);
        end else
        begin
          if TOBA.GetValue('GAC_NATUREPIECEG') = '' then // non affecté
          begin
            if (TOBP <> nil) and (MontantDispo > TOBPiece.getValue('GP_TOTALTTCDEV') ) then
            begin
              TobA.putValue('GAC_MONTANTDEV',TOBPiece.getValue('GP_TOTALTTCDEV'));
            end else // déja affecté
            begin
              TobA.putValue('GAC_MONTANTDEV', MontantDispo);
            end;
          end;
        end;
      end;
    end;
  end;

  // Modif BTP
  if SansCompta then
  begin
    for i := 0 to TOBAcomptes.detail.count - 1 do
    begin
      TOBA := TobAcomptes.detail[i];
      TobA.addchampsupValeur('MONTANTDISPO', 0, false);
      TobA.addchampsupValeur('MONTANTINIT', TOBA.GetValue('GAC_MONTANTDEV'));
    end;
  end;
end;
function TOF_GCACOMPTES.AcomptesSaisie : boolean;
var II : integer;
		Nb : Integer;
begin
  Nb := 0;
  For II := 0 to listeOnglet.Count -1 do
  begin
		Nb := Nb + listeOnglet.items[II].TOBAcomptes.Detail.Count;
  end;
  Result := (Nb > 0);
end;

procedure TOF_GCACOMPTES.AffichelesGrids;
var ii : Integer;
		TheOnglet,lastOnglet : Tonglet;
    cancel : Boolean;
    Arow,Acol : Integer;
begin
  lastOnglet := fCurrentOnglet;
  for II := 0 to listeOnglet.Count -1 do
  begin
		TheOnglet := listeOnglet.Items[II];
		AfficheGrid(TheOnglet);
    TFVierge(Ecran).Hmtrad.ResizeGridColumns(TheOnglet.GS);
    fCurrentOnglet := TheOnglet;
  	GSEnter(TheOnglet.GS);
    Acol := COlJal;
    ARow := 1;
    GSRowEnter(TheOnglet.GS,Arow,cancel,False);
    GSCellEnter(TheOnglet.GS,Acol,Arow,cancel);
    TheOnglet.GS.ShowEditor;
  end;
  fCurrentOnglet := lastOnglet;
  
end;

procedure TOF_GCACOMPTES.AfficheGrid(TheOnglet : Tonglet);
var ii : Integer;
		nbrLig : integer;
begin
  NbrLig := TheOnglet.TobAcomptes.Detail.Count +1;
  if NbrLig < 2 then NbrLig := 2;
  TheOnglet.GS.RowCount := nbrLig;
  for II := 0 to TheOnglet.TOBAcomptes.Detail.count -1 do
  begin
		TheOnglet.TOBAcomptes.Detail[II].PutLigneGrid(TheOnglet.GS,II+1,false,False,Lescolonnes);
  end;
  //fCurrentOnglet.TobAcomptes.PutGridDetail(fCurrentOnglet.GS, True, True, LesColonnes, True);
end;

function TOF_GCACOMPTES.ControleLesSaisies : boolean;
var II,i : Integer;
		TheOnglet : TOnglet;
begin
  Result := True;
  for II := 0 to listeOnglet.Count -1 do
  begin
		TheOnglet := listeOnglet.Items[II];
    for i := TheOnglet.GS.fixedrows to TheOnglet.GS.RowCount - 1 do
    begin
      if TheOnglet.GS.isselected(i) and (valeur(TheOnglet.GS.cells[colDispo, i]) < valeur(TheOnglet.GS.cells[colMontantDev, i])) then
      begin
        LastError := (-1);
        LastErrorMsg := 'Vous depassez le montant disponible';
        PGACOMPTES.TabIndex := II;
        TheOnglet.GS.row := i;
        TheOnglet.GS.Col := colMontantDev;
        EnErreur := True;
        result := false;
        Break;
      end;
    end;
	end;
end;

procedure TOF_GCACOMPTES.GetAcomptesFromgene (TheOnglet : Tonglet);
var II : Integer;
begin
  if TobAcomptes.Detail.Count = 0 then exit;
  II := 0;
	repeat
    if TobAcomptes.detail[II].GetString('GAC_FOURNISSEUR')=TheOnglet.Fournisseur then
    begin
			TobAcomptes.detail[II].ChangeParent(TheOnglet.TOBAcomptes,-1)
    end else inc(II);
  until II>= TobAcomptes.Detail.Count;
end;

procedure TOF_GCACOMPTES.RecomposeAcomptes;
var ii,i : Integer;
		TheOnglet : TOnglet;
begin
  TobAcomptes.ClearDetail;
  for II := 0 to listeOnglet.Count -1 do
  begin
    TheOnglet := listeOnglet.items[II];
    if TheOnglet.GS.RowCount = 1 then Continue;
    i := TheOnglet.GS.fixedrows;
    for I := TheOnglet.GS.fixedrows to TheOnglet.GS.RowCount do
    begin
      if TheOnglet.GS.isselected(i) then
      begin
        if TOB(TheOnglet.GS.Objects[0, i]) <> nil then
        begin
          TOB(TheOnglet.GS.Objects[0, i]).GetLigneGrid(TheOnglet.GS, i, LesColonnes);
          MajMontants(TOBPiece, TOB(TheOnglet.GS.Objects[0, i]));
          TOB(TheOnglet.GS.Objects[0, i]).SetString('GAC_FOURNISSEUR',TheOnglet.fournisseur);
          TOB(TheOnglet.GS.Objects[0, i]).ChangeParent(TobAcomptes,-1);
        end;
      end;
    end;
  end;
end;

procedure TOF_GCACOMPTES.DefiniTOBsSaisie;
var II : Integer;
		TheOnglet : TOnglet;
begin
  for II := 0 to listeOnglet.Count -1 do
  begin
  	TheOnglet := listeOnglet.items[II];
    GetAcomptesFromgene (TheOnglet);
  end;
end;

procedure TOF_GCACOMPTES.ConfigureGrillesSaisie;
var st : string;
    depart,I,II : integer;
    Nam : string;
    TheOnglet : TOnglet;
begin
  for II := 0 to listeOnglet.Count -1 do
  begin
    TheOnglet := listeOnglet.Items [II];
    TheOnglet.GS.ColCount := NbCol;
    TheOnglet.GS.ColWidths[0] := 15;
    TheOnglet.GS.ColWidths[1] := 0;
    if not PassCompta then
    begin
      TheOnglet.GS.ColWidths[2] := 0;
    end;
    St := LesColonnes;
    colRgt := (-1);
    colDispo := (-1);
    ColJal := -1;
    ColModR := 1;
    ColLib := -1;
  	TheOnglet.GS.cells[0, 0] := '';
    for i := 0 to TheOnglet.GS.ColCount - 1 do
    begin
      depart := 1;
      if not PassCompta then
      begin
        depart := 2;
      end;
      if i > depart then TheOnglet.GS.ColWidths[i] := 100;
      Nam := ReadTokenSt(St);
      if Nam = 'GAC_MODEPAIE' then
      begin
        ColModR := i;
				TheOnglet.GS.cells[i, 0] := 'Mode de paiement';
				TheOnglet.GS.colaligns[i]:= tacenter;
        TheOnglet.GS.ColFormats[i] := 'CB=TTMODEPAIE'
      end else if Nam = 'GAC_NATUREPIECEG' then
      begin
        COlAffected := i;
      end else if Nam = 'GAC_JALECR' then
      begin
        COlJal := i;
				TheOnglet.GS.cells[i, 0] := 'Journal';
				TheOnglet.GS.colaligns[i]:= tacenter;
        TheOnglet.GS.ColFormats[i] := 'CB=TTJALBANQUEOD'
      end else if Nam = 'GAC_LIBELLE' then
      begin
        TheOnglet.GS.ColLengths[i] := -1;
        ColLib := i;
				TheOnglet.GS.cells[i, 0] := 'Libellé';
      end else if (Nam = 'GAC_MONTANTDEV') then
      begin
        TheOnglet.GS.ColWidths[i] := 80;
        TheOnglet.GS.ColFormats[i] := '#,##0.00';
        TheOnglet.GS.ColAligns[i] := taRightJustify;
        ColMontantdev := i;
				TheOnglet.GS.cells[i, 0] := 'Montant affecté';
      end else if (Nam = 'MONTANTDISPO') then
      begin
        TheOnglet.GS.ColWidths[i] := 80;
        TheOnglet.GS.ColFormats[i] := '#,##0.00';
        TheOnglet.GS.ColAligns[i] := taRightJustify;
        ColDispo := i;
				TheOnglet.GS.cells[i, 0] := 'Montant disponible';
			end else if (Nam = 'GAC_ISREGLEMENT') then
      begin
        TheOnglet.GS.ColTypes[i] := 'B';
        TheOnglet.GS.ColWidths[i] := 15;
        TheOnglet.GS.ColAligns[i] := taCenter;
        TheOnglet.GS.ColFormats[i] := intToStr(integer(csCoche));
        colRgt := i;
				TheOnglet.GS.cells[i, 0] := 'Rgt.';
      end
        ;
    end;
    TheOnglet.GS.options := TheOnglet.GS.Options - [goEditing];
    //
  end;
end;


procedure TOF_GCACOMPTES.BeforeChangeVue (Sender: TObject; var AllowChange: Boolean);
var isModified : Boolean;
begin
  if fCurrentOnglet.GS.InplaceEditor = nil then Exit;
  if fCurrentOnglet.GS.InplaceEditor.Modified then
  begin
		SendMessage(fCurrentOnglet.GS.Handle, WM_KeyDown, VK_TAB, 0);
  end;
end;

procedure TOF_GCACOMPTES.ChangeVue(Sender: Tobject);
begin
  SetEventsGrid (fCurrentOnglet.GS,false);
  //
  fCurrentOnglet := listeOnglet.items[PGACOMPTES.TabIndex];
  fCurrentOnglet.GS.Invalidate;
  //
  SetEventsGrid (fCurrentOnglet.GS,True);
end;

procedure TOF_GCACOMPTES.SetEventsGrid (GS : THgrid; Etat : boolean);
begin
  if Etat then
  begin
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    // modif BTP
    GS.OnEnter := GSEnter;
    // ---
    GS.OnDblClick := GSLigneDClick;
    GS.PostDrawCell := DessineCell;
    GS.OnBeforeFlip := GSOnBeforeFlip;
  end else
  begin
    GS.OnCellEnter := Nil;
    GS.OnCellExit := Nil;
    GS.OnRowEnter := Nil;
    GS.OnRowExit := Nil;
    GS.OnEnter := nil;
    GS.OnDblClick := nil;
    GS.PostDrawCell := nil;
    GS.OnBeforeFlip := nil;
  end;
end;


procedure TOF_GCACOMPTES.ConstitueTabSheets(TOBPIECETRAIT: TOB);
var ii,AI : Integer;
    FirstOnglet,OneOnglet : TOnglet;
    TS : TTabSheet;
    Grid : THGrid;
    PN : THpanel;
    TL,TTR : THLabel;
    TE,TR : THEdit;
begin
  // Onglet de base
  FirstOnglet := TOnglet.create;
  FirstOnglet.fournisseur := '';
  FirstOnglet.GS := THGrid(getControl ('G'));
  FirstOnglet.TabSheet := TTabSheet (getControl ('TSENTREP'));
  FirstOnglet.Total := THEdit(GetControl('TOTAL'));
  FirstOnglet.ARegler  := THEdit(GetControl('AREGLER'));
  FirstOnglet.TabSheet.caption := GetParamSocSecur('SO_LIBELLE','Notre entreprise');
	listeOnglet.Add(FirstOnglet);
  fCurrentOnglet := FirstOnglet;
  //
  if TOBPIECETRAIT = nil then Exit;
  //
  if (TOBPIECETRAIT.detail.count > 0) and (ExisteReglDirect(TOBpieceTrait)) then
  begin
  	AI := 0;
    for ii := 0 to TOBPIECETRAIT.detail.Count -1 do
    begin
      if (TOBPieceTrait.detail[ii].getValue('BPE_TYPEINTERV')='Y00') and
         (TOBPieceTrait.detail[ii].getValue('TYPEPAIE')='001') then
      begin
        //
        Inc(AI);
        TS := TTabSheet.Create( TpageControl (getControl ('PGACOMPTES')));
        TS.Parent :=  TpageControl (getControl ('PGACOMPTES'));
        TS.PageControl := TpageControl (getControl ('PGACOMPTES'));
        TS.ImageIndex := -1;
        TS.name := 'TS'+InttoStr(AI);
        TS.TabVisible := true;
        //
        PN := THPanel.Create(PGAcomptes);
        PN.Parent := TS;
        PN.Align := alBottom;
        PN.Height  := THPanel(GetControl('PBASENT')).Height;
        //
        Grid := THGrid.Create(TS); // la grille de saisie
        Grid.Parent := TS;
        Grid.Align := alClient;
        Grid.RowCount := THGrid(getControl('G')).RowCount;
        Grid.ColCount := THGrid(getControl('G')).ColCount;
        Grid.DefaultRowHeight := THGrid(getControl('G')).DefaultRowHeight;
        Grid.Fixedcols := THGrid(getControl('G')).Fixedcols;
        Grid.FixedRows := THGrid(getControl('G')).FixedRows;
        Grid.DefaultColWidth := THGrid(getControl('G')).DefaultColWidth;
        Grid.ParentColor := THGrid(getControl('G')).ParentColor;
        Grid.FixedColor := THGrid(getControl('G')).FixedColor;
        Grid.Color := THGrid(getControl('G')).Color;
        Grid.AlternateColor := THGrid(getControl('G')).AlternateColor;
        Grid.Ctl3D := THGrid(getControl('G')).Ctl3D;
        Grid.Options := THGrid(getControl('G')).Options;
        Grid.MultiSelect := THGrid(getControl('G')).MultiSelect;
        Grid.TitleBold := THGrid(getControl('G')).TitleBold;
        Grid.TitleCenter := THGrid(getControl('G')).TitleCenter;
        Grid.SortEnabled  := THGrid(getControl('G')).SortEnabled;
        Grid.RowHeights [0] := THGrid(getControl('G')).RowHeights [0];
        //
        TE := THEdit.Create(PN);
        TE.Parent := PN;
        TE.Top := THEdit(GetControl('TOTAL')).Top;
        TE.Left := THEdit(GetControl('TOTAL')).Left;
        TE.Height  := THEdit(GetControl('TOTAL')).Height;
        TE.Width  := THEdit(GetControl('TOTAL')).Width;
        TE.ReadOnly  := THEdit(GetControl('TOTAL')).ReadOnly;
        TE.OpeType   := THEdit(GetControl('TOTAL')).OpeType;
        TE.Text := '0.0';
        //
        TR := THEdit.Create(PN);
        TR.Parent := PN;
        TR.Top := THEdit(GetControl('AREGLER')).Top;
        TR.Left := THEdit(GetControl('AREGLER')).Left;
        TR.Height  := THEdit(GetControl('AREGLER')).Height;
        TR.Width  := THEdit(GetControl('AREGLER')).Width;
        TR.ReadOnly  := THEdit(GetControl('AREGLER')).ReadOnly;
        TR.OpeType   := THEdit(GetControl('AREGLER')).OpeType;
        TR.Text := '0.0';
        //
        TL := THLabel.Create(PN);
        TL.Parent := PN;
        TL.Top := THLabel(GetControl('TTOTAL')).Top;
        TL.Left := THLabel(GetControl('TTOTAL')).Left;
        TL.Height  := THLabel(GetControl('TTOTAL')).Height;
        TL.Width  := THLabel(GetControl('TTOTAL')).Width;
        TL.Caption  := THLabel(GetControl('TTOTAL')).Caption;
        //
        TTR := THLabel.Create(PN);
        TTR.Parent := PN;
        TTR.Top := THLabel(GetControl('TAREGLER')).Top;
        TTR.Left := THLabel(GetControl('TAREGLER')).Left;
        TTR.Height  := THLabel(GetControl('TAREGLER')).Height;
        TTR.Width  := THLabel(GetControl('TAREGLER')).Width;
        TTR.Caption  := THLabel(GetControl('TAREGLER')).Caption;
        //
      	OneOnglet := TOnglet.create;
        OneOnglet.TabSheet := TS;
        OneOnglet.GS := Grid;
        OneOnglet.Total := TE;
        OneOnglet.ARegler := TR;
        OneOnglet.fournisseur := TOBPieceTrait.detail[ii].getValue('BPE_FOURNISSEUR');
  			OneOnglet.TabSheet.caption := TOBPieceTrait.detail[ii].getValue('LIBELLE');
        listeOnglet.Add(OneOnglet);
      end;
    end;
  end;
end;

function TOF_GCACOMPTE.GetLibelleRegl : string;
begin
  if Action = taCreat then
  begin
    if THRadioGroup(GetControl('GAC_ISREGLEMENT')).Value <> 'X' then
    begin
      result := Copy('Acompte ' + GetControlText ('GAC_DATEECR') + ' Pce. '+ IntToStr(TOBPiece.GetValue('GP_NUMERO')), 1, 35);
    end else
    begin
      result := Copy('Règlement ' + GetControlText ('GAC_DATEECR') + ' Pce. '+ IntToStr(TOBPiece.GetValue('GP_NUMERO')), 1, 35);
    end;
  end;
end;

procedure TOF_GCACOMPTE.ChangeDateAcc(Sender: Tobject);
begin
	SetControlText('GAC_LIBELLE',GetLibelleRegl);
end;

procedure TOF_GCACOMPTE.ChangeTypeSaisie(Sender: TObject);
begin
	SetControlText('GAC_LIBELLE',GetLibelleRegl);
end;

{ TlistTabSheet }

function TlistTabSheet.Add(AObject: TOnglet): Integer;
begin
	Result := inherited ADD(Aobject);
end;

procedure TlistTabSheet.clear;
var indice : integer;
begin
  if count > 0 then
  begin
    for Indice := count -1 downto 0 do
    begin
      if TOnglet(Items [Indice])<> nil then
      begin
         if TOnglet(Items[Indice]).fournisseur <> '' then
         begin
           TOnglet (Items [Indice]).free;
           Items[Indice] := nil;
         end;
      end;
    end;
  end;
  inherited;
end;

constructor TlistTabSheet.create;
begin

end;

destructor TlistTabSheet.destroy;
begin
	clear;
  inherited;
end;

function TlistTabSheet.findOnglet(NomOnglet: string): TOnglet;
var Indice : integer;
begin
  result := nil;
  for Indice := 0 to Count -1 do
  begin
    if Items[Indice].Name = NomOnglet then
    begin
      result:=Items[Indice];
      break;
    end;
  end;
end;

function TlistTabSheet.GetItems(Indice: integer): TOnglet;
begin
  result := TOnglet (Inherited Items[Indice]);
end;

procedure TlistTabSheet.SetItems(Indice: integer; const Value: Tonglet);
begin
  Inherited Items[Indice]:= Value;
end;

{ TOnglet }

constructor TOnglet.create;
begin
  Name := '';
	TabSheet := nil;
  GS := nil;
  Fournisseur := '';
  TOBAcomptes := TOB.Create('LES ACOMPTES',nil,-1);
end;

destructor TOnglet.destroy;
begin
	if TOBAcomptes <> nil then FreeAndNil(TOBAcomptes);
  if GS <> nil then FreeAndNil(GS);
  inherited;
end;

function TOF_GCACOMPTES.GetmontantPorts(TOBpiece: TOB): Double;
var QQ : TQuery;
    cledoc : r_cledoc;
    TOBPiedP : TOB;
    ii : Integer;
begin
  Result := 0;
  cledoc := TOB2cledoc (TOBpiece);
  TOBPiedP := TOB.Create('LES PORRCS',nil,-1);
  QQ := OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(cledoc,ttdPorc,true),True,-1,'',true);
  TRY
    if not QQ.Eof then
    begin
      TOBPiedP.LoadDetailDB('PIEDPORT','','',QQ,false);
    end;
    for II := 0 to TOBPiedP.detail.count -1 do
    begin
      if (TOBPiedP.detail[II].getValue('GPT_FRAISREPARTIS')<>'X') and (not TOBPiedP.detail[II].getboolean('GPT_RETENUEDIVERSE')) then
      begin
        result := result + TOBPiedP.detail[II].getValue('GPT_TOTALHTDEV');
      end;
    end;
  finally
    ferme (QQ);
    TOBPiedP.free;
  end;
end;

initialization
  registerclasses([TOF_GCACOMPTES, TOF_GCACOMPTE]);

  RegisterAglProc('GridSetEditing', TRUE, 2, AGLGridSetEditing);
  RegisterAglProc('ACOMPTES_BFermerClick', TRUE, 0, AGLACOMPTES_BFermerClick);
end.
