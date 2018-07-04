unit Integecr;

// CA - 17/05/1999 - Gestion en direct du code Journal
//                 - Gestion correcte des couleurs
//                 - Affichage direct de la liste des pièces après intégration
// CA - 24/06/1999 - GetNewNumJal : ajout paramètre date pièce
// CA - 08/07/1999 - Type de journal d'échéances = journal d'achat
// CA - 08/07/1999 - Gestion du libellé unique sans case à cocher
// CA - 08/07/1999 - F9/F10 pour lancer le recalcul
// CA - 09/07/1999 - paramètres supplémentaires : intégration dans la compta et détail
// CA - 21/08/2000 - Affichage bouton ventilation suivant paramétrage.
// CA - 18/07/2003 - Suppression champs euro
// CA - 26/08/2004 - Mise à jour des totaux pour toutes les lignes
//MBO - 08/07/2005 - FQ 15095 - Remplacement check box ecritures de dotation simulation par
//                   radio bouton car ajout écritures de situation
//                   + correction pb disable du rb simu et situ en entrée ds la fenetre
//mbo - 26/07/2005 - FQ 16280 Gestion du nb de décimales suivant monnaie de tenue du dossier
//FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
//PGR - 27/01/2006 - Nouveaux paramètres société pour intégration des écritures
//mbo - 20/03/2006 - fq 17641 - pb en intégration d'écriture de situation
//XVI 20/03/2006 FD 3643 Initialisation de la date de lettrage à idate2099
//XVI 18/04/2006 FD 3978 Ajoute des l'analytiques aux Immos S1
//XVI 25/04/2006 FD 3978 ajoute de la mono ventilation 
//MBO - 03/04/2006 - fq 17410 - pouvoir intégrer sur des journaux d'OD et aussi de type REG
// BTY - 04/06 FQ 17775 Modif textes pour l'appel Intégration à partir de la liste immos
// MVG - 20/04/06 - pour perte des modifications du 03/04/06 et 05/04/06, conservation des modifs de L. Gendreau
// BTY - cf FQ 17775 06/06 Ouvrir F10 via le traitement d'intégration et fermer via la liste des immos
//XVI 07/07/2006 FQ(BL) 12731 Ne pas affichewr les journaux réservés à l'expert.
// MVG 12/07/2006 pour report des modifs de XVI de 04/2006 et correction conseil
// MVG 15/11/2006 FQ 19132
// MVG 15/11/2006 FQ 19133
// MVG 11/05/2007 La colonne E_DATELETTRAGE existe que pour SERIE1
// MVG 04/06/2007 Ajout modif pour CBL préfixé de CCNF
// MVG 20/08/2007 Annulation modif CCNF (FQ 14452) 

interface

uses
  Windows, Menus, HSysMenu, hmsgbox, Classes, HPanel, Forms, StdCtrls,
  Hctrls, Mask, ExtCtrls, ComCtrls, Grids, Controls, HTB97, ImOutGen,
  Graphics, ImGenEcr, uTob, uListByUser, uTXML, ParamSoc
{$IFDEF SERIE1}
  ,utLog
{$ELSE}
  , UtilPGI, ZCompte, ZTiers, uLibEcriture
{$ENDIF}
  ;

function IsValidDateEcriture(sDate: string): boolean;
function CompareSuivantCompteRef(Item1, Item2: Pointer): integer;

const
  COL_CPT = 2;

type
  TIntegration = class(TForm)
    HM: THMsgBox;
    HPanel1: THPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bExport: TToolbarButton97;
    BImprimer: TToolbarButton97;
    HPanel4: THPanel;
    HPanel3: THPanel;
    GridEcriture: THGrid;
    HPanel2: THPanel;
    Pages: TPageControl;
    PDotation: TTabSheet;
    Bevel4: TBevel;
    GroupBox2: TGroupBox;
    HLabelDateCalcul: THLabel;
    HLabel2: THLabel;
    DateCalcul: THCritMaskEdit;
    DateGeneration: THCritMaskEdit;
    GroupBox1: TGroupBox;
    LibelleUnique: TEdit;
    CodeJournal: THValComboBox;
    PEcheance: TTabSheet;
    Bevel1: TBevel;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    rbRecapEch: TRadioButton;
    rbSynthEch: TRadioButton;
    tJEcheance: THLabel;
    GroupBox6: TGroupBox;
    tCBanque: THLabel;
    cbEcrPaiement: TCheckBox;
    cbEcrEcheance: TCheckBox;
    tCTva: THLabel;
    tJBanque: THLabel;
    JEcheance: THValComboBox;
    HLabel6: THLabel;
    DateDebutEch: THCritMaskEdit;
    HLabel7: THLabel;
    DateFinEch: THCritMaskEdit;
    CTva: THCritMaskEdit;
    JBanque: THValComboBox;
    CBanque: THCritMaskEdit;
    HMTrad: THSystemMenu;
    GroupBox3: TGroupBox;
    rbDetail: TRadioButton;
    rbSynth: TRadioButton;
    Dock972: TDock97;
    PFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    bRecalculDot: TToolbarButton97;
    FFiltres: THValComboBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    CreditProg: THNumEdit;
    DebitProg: THNumEdit;
    CreditGen: THNumEdit;
    DebitGen: THNumEdit;
    HLabel1: THLabel;
    HLabel3: THLabel;
    BVentil: TToolbarButton97;
    LibelleCBanque: THLabel;
    HLabel4: THLabel;
    rbGroupe: TRadioButton;
    HLabel5: THLabel;
    bSimuEche: TCheckBox;
    cbGestionTva: TCheckBox;
    iCritGlyphModified: TImage;
    iCritGlyph: TImage;
    BDUPFILTRE: TMenuItem;
    // mbo 15095 ajout des 4 lignes ci_dessous
    GroupBox7: TGroupBox;
    rbnor: TRadioButton;
    Rbsimu: TRadioButton;
    rbsitu: TRadioButton;
    //
    procedure FormShow(Sender: TObject);
    procedure DateExit(Sender: TObject);
    procedure bExportClick(Sender: TObject);
    procedure GridEcritureClick(Sender: TObject);
    procedure BRecalculClick(Sender: TObject);
    procedure DateEnter(Sender: TObject);
    procedure CodeJournalEnter(Sender: TObject);
    procedure LibelleUniqueEnter(Sender: TObject);
    procedure LibelleUniqueExit(Sender: TObject);
    procedure OnDetailClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure cbEcrEcheanceClick(Sender: TObject);
    procedure cbEcrPaiementClick(Sender: TObject);
    procedure DateKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    //    procedure Button1Click(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure OnEnterZone(Sender: TObject);
    procedure OnExitZone(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure JBanqueChange(Sender: TObject);
    procedure OnChangeCritere(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BDUPFILTREClick(Sender: TObject);
  private
    { Déclarations privées }
    fTypeEcr: TypeIntegration;
    CurDate, CurLibelle, CurJournal: string;
    NivDetail: integer;
    fIndiceImmoEnCours: integer;
    fIndiceEcriture: integer;
    fListeEcriture: TList;
    fNomFiltre: string;
    fListeImmo: HTStrings;
    fbModifie: boolean;
    fCurVal: variant;
    fbCompta: boolean;
    fbDetail: boolean;
    fIntegrationDB: boolean;
    fFichierExport: string;
    PendantTraitment: boolean;
    FListeByUser:TListByUser;
{$IFDEF SERIE1}
{$ELSE}
    fCR: TList;
    procedure InitFromListe(OBEcr: TOB; ARecord: TLigneEcriture; ModeSaisie:
      string = '-');
    procedure InitFromListeAna(OBAna: TOB; ARecord: TAna; OBEcr: TOB);
    procedure MajSolde(NumPiece: longint; OBEcr: TOB);
    // modif mbo 15095
    //procedure InitCommun(OBEcr: TOB; bSimul: boolean);
    //procedure InitCommunAna(OBEcr: TOB; bSimul: boolean);
    procedure InitCommun(OBEcr: TOB; bType: integer);
    procedure InitCommunAna(OBAna: TOB; bType: integer);
    procedure ExporteImmoFormatPGI(OB: TOB);
{$ENDIF}
    procedure InitGrid;
    procedure AfficheGrid;
    function CalculCumul(nLigneFin: integer): TCumul;
    procedure VideGrid;
    procedure DessineCell(Acol, ARow: LongInt; Canvas: TCanvas; AState:
      TGridDrawState);
    procedure AjouteLigneDansGrid(ColDa, ColJ, ColCo, ColL, ColDe, ColCr:
      string);
    procedure ValeursParDefaut;
    procedure RecalculEcritures;
    procedure MajFiche;
{$IFDEF SERIE1}
    procedure IntegrationCompta;
{$ELSE}
    procedure IntegreDansLaCompta;
{$ENDIF}
    procedure MajTableEcheance;
    procedure BougeFocus;
    function ControleZoneEcheance: boolean;
    function GetCompteBanqueContrepartie: string;
    function ControleZoneDotation: boolean;
    procedure EclateLigneDC(L: TList);
    function ControleExistenceCompte: boolean;
    // Code du MUL pour la nouvelle Gestion des FILTRES
    procedure InitAddFiltre(T:TOB);
    procedure InitGetFiltre(T:TOB);
    procedure InitSelectFiltre(T:TOB);
    procedure ParseParamsFiltre(Params: HTStrings);
    procedure UpgradeFiltre(T:TOB);
  public
    { Déclarations publiques }
  end;

procedure IntegrationEcritures(Op: TypeIntegration; ListeImmo: HTStrings;
  bCompta, bDetail: boolean);

implementation

uses
  ImpExp, SiscoTrt, UiUTIL, SysUtils, Filtre, ImEnt,
  Hent1, Outils, HStatus, Paramdat
{$IFDEF eAGLCLient}
  , UtileAGL
{$ELSE}
  , PrintDBG
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF eAGLCLient}

{$IFDEF SERIE1}
  , ecr_edit
  , UtInput
  , S1Util
  , UtY_;
{$ELSE}
  , ImVisAna, IMMOCPTE_TOM, SaisUtil, ImRapInt, UtilSais //AboUtil

  ;
{$ENDIF}

(*, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, StdCtrls, HPanel, Buttons, HCtrls,ExtCtrls,HEnt1,DBTables, Db, Math,
  DBGrids, Spin, hmsgbox, ImEnt, MajTable, Hqry, DBCtrls,
  Mask,Outils, HTB97, HSpeller, ComCtrls,ParamDat, HSysMenu, UiUtil,
  Menus,UTob,HStatus,ParamSoc,ImOutGen*)

{$R *.DFM}

const
 //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
 HM: array[0..0] of string =
 {1} ('Veuillez définir les paramètres d''intégration d''écritures dans la commande paramètres.'
    );
{$IFDEF SERIE1}
function GetCollectif(CpteAux: string): string;
var
  Q: TQuery;
begin
  result := '';
  Q := OpenSql('SELECT T_COLLECTIF FROM TIERS WHERE T_AUXILIAIRE="' + CpteAux +
    '"', false);
  if not Q.Eof then
    result := Q.Fields[0].AsString;
  ferme(Q);
end;
{$ENDIF}

procedure IntegrationEcritures(Op: TypeIntegration; ListeImmo: HTStrings;
  bCompta, bDetail: boolean);
var
  Integration: TIntegration;
  PP: THpanel;
  j: Integer;
begin
{$IFDEF SERIE1}
{$ELSE}
  if GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Encours.Code then
    exit;

  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  if Op = toDotation then
  begin
    if (GetParamSocSecur('SO_IMMOJALDOTDEF','') = '') or (GetParamSocSecur('SO_IMMODOTCHOIXDET','') = '') or
       (GetParamSocSecur('SO_IMMODOTCHOIXTYP','') = '') then
    begin
      PGIBox(HM[0]);
      exit;
    end;
  end
  else
  begin
    if (GetParamSocSecur('SO_IMMOJALECHDEF','') = '') or (GetParamSocSecur('SO_IMMOECHCHOIXDET','') = '') or
       (GetParamSocSecur('SO_IMMOECHCHOIXTYP','') = '') then
    begin
      PGIBox(HM[0]);
      exit;
    end;
  end;
{$ENDIF}

  Integration := TIntegration.Create(Application);
  Integration.fTypeEcr := Op;
  Integration.fIndiceEcriture := 0;
  Integration.fListeImmo := ListeImmo;
  Integration.NivDetail := 2;
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  Integration.fbDetail := bDetail;

  Integration.rbDetail.Checked := bDetail;
  Integration.fbCompta := bCompta;
  // ajout mbo 26.07.05 FQ 16
  if V_PGI.OkDecV=0 then Integration.CreditProg.Masks.PositiveMask:='#,##0' else Integration.CreditProg.Masks.PositiveMask:='#,##0.' ;
  for j:=1 to V_PGI.OkDecV do Integration.CreditProg.Masks.PositiveMask:=Integration.CreditProg.Masks.PositiveMask+'0' ;
  Integration.DebitGen.Masks.PositiveMask := Integration.CreditProg.Masks.PositiveMask ;
  Integration.CreditGen.Masks.PositiveMask := Integration.CreditProg.Masks.PositiveMask ;
  Integration.DebitProg.Masks.PositiveMask := Integration.CreditProg.Masks.PositiveMask ;

  Integration.DebitGen.Masks.ZeroMask:=Integration.CreditProg.Masks.PositiveMask ;
  Integration.CreditGen.Masks.ZeroMask:=Integration.CreditProg.Masks.PositiveMask ;
  Integration.DebitProg.Masks.ZeroMask:=Integration.CreditProg.Masks.PositiveMask ;
  Integration.CreditProg.Masks.ZeroMask:=Integration.CreditProg.Masks.PositiveMask ;
  // fin ajout mbo

  PP := FindInsidePanel;
  if PP = nil then
  begin
    try
      Integration.ShowModal;
    finally
      Integration.Free;
    end;
  end
  else
  begin
    InitInside(Integration, PP);
    Integration.Show;
  end;
end;

procedure TIntegration.ValeursParDefaut;
begin
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  //DateCalcul.Text := DateToStr(VHImmo^.Encours.Fin);
  //DateGeneration.Text := DateToStr(VHImmo^.Encours.Fin);
  DateDebutEch.Text := DateToStr(VHImmo^.Encours.Deb);
  DateFinEch.Text := DateToStr(VHImmo^.Encours.Fin);
  JEcheance.Text := '';
  CTva.Text := '';
  JBanque.Text := '';
  CBanque.Text := '';
  //cbEcrEcheance.Checked := false;
  (*JEcheance.Enabled := false;tJEcheance.Enabled := false;
  CTva.Enabled := false;tCTva.Enabled := false;*)
  cbEcrPaiement.Checked := false;
  JBanque.Enabled := false;
  tJBanque.Enabled := false;
  CBanque.Enabled := false;
  tCBanque.Enabled := false;
{$IFDEF SERIE1}
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  DateCalcul.Text := DateToStr(VHImmo^.Encours.Fin);
  DateGeneration.Text := DateToStr(VHImmo^.Encours.Fin);

  cbGestionTVA.Visible := False;
  CTVA.Visible := True;
  tCTVA.Visible := True;
{$ELSE}
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  if FinDeMois(Date) < VHImmo^.Encours.Fin then
      DateCalcul.Text := DateToStr(FinDeMois(Date))
  else DateCalcul.Text := DateToStr(VHImmo^.Encours.Fin);

  //MVG 15/11/2006 FQ 19133
if FinDeMois(Date) < VHImmo^.Encours.Deb then
      DateCalcul.Text := DateToStr(VHImmo^.Encours.Fin);

  DateGeneration.Text := DateCalcul.Text;

  cbGestionTVA.Visible := True;
  CTVA.Visible := False;
  tCTVA.Visible := False;
{$ENDIF}
end;

procedure TIntegration.FormShow(Sender: TObject);

begin
{$IFDEF SERIE1}
  BVentil.Visible := false;
  bSimuEche.Visible := false;
  BExport.Visible := false;
  //mbo 15095 modif check en rb
  rbSimu.Visible := false;
  // ajout mbo n° 15095
  rbSitu.Visible := false;
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  LibelleUnique.Text := '';

{$ELSE}
  BVentil.Visible := (GetParamSocSecur('SO_CPPCLSANSANA',False) = False);

  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures

  if fTypeEcr = toDotation then
  begin
    CodeJournal.Value := GetParamSocSecur('SO_IMMOJALDOTDEF','');
    Libelleunique.Text := GetParamSocSecur('SO_IMMODOTCHOIXLIB','');
    rbSynth.Checked := (GetParamSocSecur('SO_IMMODOTCHOIXDET','SYN') = 'SYN');
    rbGroupe.Checked := (GetParamSocSecur('SO_IMMODOTCHOIXDET','CPT') = 'CPT');
    rbDetail.Checked := (GetParamSocSecur('SO_IMMODOTCHOIXDET','DET') = 'DET');
    rbnor.Checked := (GetParamSocSecur('SO_IMMODOTCHOIXTYP','N') = 'N');
    rbsimu.Checked := (GetParamSocSecur('SO_IMMODOTCHOIXTYP','S') = 'S');
    rbsitu.Checked := (GetParamSocSecur('SO_IMMODOTCHOIXTYP','U') = 'U');

//    rbDetail.Checked := fbDetail; MVG 15/11/2006 FQ 19132
if fbdetail then rbDetail.Checked:=true; //pour accès depuis la liste des immos

    if (rbDetail.Checked = false) and (rbSynth.Checked = false) and (rbGroupe.Checked = false) then
       rbGroupe.Checked := true;
  end
  else
  begin
    JEcheance.Value := GetParamSocSecur('SO_IMMOJALECHDEF','');
    cbGestionTva.Checked := (GetParamSocSecur('SO_IMMOGERETVA', True));
    rbSynthEch.Checked := (GetParamSocSecur('SO_IMMOECHCHOIXDET','SYN') = 'SYN');
    rbRecapEch.Checked := (GetParamSocSecur('SO_IMMOECHCHOIXDET','DET') = 'DET');
    bSimuEche.Checked := (GetParamSocSecur('SO_IMMOECHCHOIXTYP','S')= 'S');
  end;

  bSimuEche.Enabled := fbCompta;
  //PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
  if  bSimuEche.Enabled = false then
      bSimuEche.Checked := false;

  //mbo 15095
  rbSimu.Enabled := fbCompta;
  rbSitu.Enabled :=fbCompta;

{$ENDIF}

  // BTY 04/06 FQ 17775
  bValider.Visible := fbCompta;
  bValider.Enabled := fbCompta;

  // ajout mbo FQ 15561
 if (GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Encours.Code) then
  begin
     bvalider.Enabled := False;
  end;

 if fTypeEcr = toDotation then
  begin
{$IFDEF SERIE1}
    HelpContext := 511010;
{$ELSE}
    HelpContext := 2140100;
{$ENDIF}
    // BTY 04/06 FQ 17775
    //Caption := HM.Mess[3];
    if fbcompta then Caption := HM.Mess[3]
    else Caption := HM.Mess[25];
    Pages.ActivePage := PDotation;
    PEcheance.TabVisible := false;
  end
  else
  begin
{$IFDEF SERIE1}
    HelpContext := 511020;
{$ELSE}
    HelpContext := 2140200;
{$ENDIF}
    // BTY 04/06 FQ 17775
    //Caption := HM.Mess[4];
    if fbcompta then Caption := HM.Mess[4]
    else Caption := HM.Mess[26];
    Pages.ActivePage := PEcheance;
    PDotation.TabVisible := false;
  end;
  UpdateCaption(self);
  if fTypeEcr = toDotation then
    fNomFiltre := 'IMODOT'
  else
    fNomFiltre := 'IMOECH';
  // Création du TlisteByUser pour la nouvelle gestion des filtres
    FListeByUser := TListByUser.Create(FFiltres, BFiltre, toFiltre, False);
    with FListeByUser do
    begin
        OnSelect := InitSelectFiltre;
        OnInitGet := InitGetFiltre;
        OnInitAdd := InitAddFiltre;
        OnUpgrade := UpgradeFiltre;
        OnParams := ParseParamsFiltre;
    end;
    if FListeByUser <> nil then
      FListeByUser.LoadDB(FNomFiltre);
  GridEcriture.GetCellCanvas := DessineCell;
  if FFiltres.ItemIndex = -1 then
    ValeursParDefaut;
  InitGrid;
  if rbDetail.Checked then
    NivDetail := 2
  else if rbGroupe.Checked then
    NivDetail := 1
  else
    NivDetail := 0;
  RecalculEcritures;
  (*if fListeEcriture.Count = 0 then
  begin
    Action := caHide;
    FormClose (Sender,Action);
    exit;
  end
  else *)MajFiche;
  // ajout du test mbo 15095
  if fTypeEcr = toDotation then
     OnChangeCritere(CodeJournal)
  else
     OnChangeCritere(JEcheance);
  fbModifie := false;
  BRecalculDot.Glyph := iCritGlyph.Picture.BitMap;
end;

procedure TIntegration.FormCreate(Sender: TObject);
begin
  inherited;
  fListeEcriture := TList.Create;
{$IFDEF SERIE1}
  CodeJournal.Plus := '';
  CodeJournal.DataType := 'TTJOURNAL';
  // mbo fq 17410 CodeJournal.Plus := 'AND J_NATUREJAL="OD" ';
  CodeJournal.Plus := 'AND (J_NATUREJAL="OD" OR J_NATUREJAL="REG") '
  {$IFDEF SERIE1}
                    +'and J_RESERVEEXP<>"X"' //XVI 07/07/2006 FQ 12731
  {$ELSE}
  {$ENDIF !SERIE1}
      ;
  JEcheance.DataType := 'TTJOURNAL';
  JEcheance.Plus := 'AND J_NATUREJAL="ACH"'
  {$IFDEF SERIE1}
                    +'and J_RESERVEEXP<>"X"' //XVI 07/07/2006 FQ 12731
  {$ELSE}
  {$ENDIF !SERIE1}
      ;
  JBanque.DataType := 'TTJOURNAL';
  JBanque.Plus := 'AND J_NATUREJAL="TRE"'
  {$IFDEF SERIE1}
                    +'and J_RESERVEEXP<>"X"' //XVI 07/07/2006 FQ 12731
  {$ELSE}
  {$ENDIF !SERIE1}
      ;
  CTva.DataType := 'TTGENERAUX';
  CTva.Plus := 'AND G_NATUREGENE="TDE"';
  CTva.ElipsisButton := true;
  CBanque.DataType := 'TTGENERAUX';
  CBanque.Plus := 'AND G_NATUREGENE="BQE"';
  CBanque.ElipsisButton := true;
{$ELSE}
  if V_PGI.NumVersionBase > 516 then
  begin
    CodeJournal.Plus := '';
    CodeJournal.Datatype := 'TTJALCRIT';
    // mbo FQ 17410 CodeJournal.Plus := 'J_FERME="-" AND J_NATUREJAL="OD" AND J_JOURNAL<>"' +
        CodeJournal.Plus := 'J_FERME="-" AND (J_NATUREJAL="OD" OR J_NATUREJAL="REG") AND J_JOURNAL<>"' +
      VHImmo^.JALATP + '" AND J_JOURNAL<>"' + VHImmo^.JALVTP + '"';
    JEcheance.DataType := 'TTJOURNAL';
    JEcheance.Plus := 'AND J_NATUREJAL="ACH"'
  {$IFDEF SERIE1}
                    +'and J_RESERVEEXP<>"X"' //XVI 07/07/2006 FQ 12731
  {$ELSE}
  {$ENDIF !SERIE1}
         ;
  end
  else
    CodeJournal.Datatype := 'TTJALODSTP';
{$ENDIF}
  CodeJournal.Value := GetParamSocSecur('SO_IMMOJALDOTDEF','');
  JEcheance.Value := GetParamSocSecur('SO_IMMOJALECHDEF','');
{$IFDEF SERIE1}
{$ELSE}
  cbGestionTva.Checked := (GetParamSocSecur('SO_IMMOGERETVA', True));
{$ENDIF}
  PendantTraitment := false;
end;

procedure TIntegration.FormClose(Sender: TObject;
  var Action: TCloseAction);
var
    bCanClose : Boolean;  
begin
  if FListeByUser <> nil then
  begin
      if FListeByUser.IsModified then
          FListeByUser.CloseQuery(bCanClose);
      FreeAndNil(FListeByUser); // Libération du TListByUser
  end;
  GridEcriture.VidePile(True);
  (*VIdeListeCompte(fListeComptes); fListeComptes.Free ;*)
  VideListeEcritures(fListeEcriture);
  fListeEcriture.free;
  if isInside(Self) then
    Action := caFree;
end;

(*procedure TIntegration.Button1Click(Sender: TObject);
begin
    ExecuteVerificationEcarts ;
end;*)

procedure TIntegration.BCreerFiltreClick(Sender: TObject);
begin
  FListeByUser.Creer;
  FListeByUser.Save;
end;

procedure TIntegration.BSaveFiltreClick(Sender: TObject);
begin
  FListeByUser.Save;
end;

procedure TIntegration.BDelFiltreClick(Sender: TObject);
begin
  VideFiltre(FFiltres, Pages);
  FListeByUser.Delete;
  FListeByUser.Save;
end;

procedure TIntegration.BRenFiltreClick(Sender: TObject);
begin
  FListeByUser.Rename;
end;

procedure TIntegration.BNouvRechClick(Sender: TObject);
begin
  VideFiltre(FFiltres, Pages);
  FListeByUser.New;
  ValeursParDefaut;
end;

procedure TIntegration.FFiltresChange(Sender: TObject);
begin
  FListeByUser.Select(FFiltres.Value);
end;

procedure TIntegration.POPFPopup(Sender: TObject);
begin
  UpdatePopFiltre(BSaveFiltre, BDelFiltre, BRenFiltre, FFiltres);
end;

procedure TIntegration.InitGrid;
begin
  GridEcriture.ColAligns[0] := taCenter;
  GridEcriture.ColAligns[1] := taCenter;
  GridEcriture.ColAligns[2] := taCenter;
  GridEcriture.ColAligns[3] := taLeftJustify;
  GridEcriture.ColAligns[4] := taRightJustify;
  GridEcriture.ColAligns[5] := taRightJustify;
end;

function TIntegration.CalculCumul(nLigneFin: integer): TCumul;
var
  TotalDebit, TotalCredit: double;
  i: integer;
begin
  TotalDebit := 0.0;
  TotalCredit := 0.0;
  for i := 1 to nLigneFin do
  begin
    if IsFloat(GridEcriture.Cells[4, i]) then
      TotalDebit := TotalDebit + StrToFloat(GridEcriture.Cells[4, i]);
    if IsFloat(GridEcriture.Cells[5, i]) then
      TotalCredit := TotalCredit + StrToFloat(GridEcriture.Cells[5, i]);
  end;
  result.Debit := TotalDebit;
  result.Credit := TotalCredit;
end;

procedure TIntegration.DateExit(Sender: TObject);
begin
  if fCurVal <> THCritMaskEdit(Sender).Text then
    fbModifie := True;
  if IsValidDate(THCritMaskEdit(Sender).Text) then
  begin
    if (StrToDate(THCritMaskEdit(Sender).Text) < VHImmo^.Encours.Deb)
      or (StrToDate(THCritMaskEdit(Sender).Text) > VHImmo^.Encours.Fin) then
    begin
      THCritMaskEdit(Sender).Text := CurDate;
      //      FocusControl(THCritMaskEdit(Sender));
      HM.Execute(5, Caption, '');
    end;
  end
  else
  begin
    HM.Execute(14, Caption, '');
    //    FocusControl(THCritMaskEdit(Sender));
  end;
end;

procedure TIntegration.AjouteLigneDansGrid(ColDa, ColJ, ColCo, ColL, ColDe,
  ColCr: string);
var
  bTrouve: boolean;
  i: integer;
begin
  bTrouve := false;
  {  if NivDetail = 0 then   // déjà fait en amont.
    begin
      for i := 1 to GridEcriture.RowCount - 2 do
      begin
        if (GridEcriture.Cells[0,i]=ColDa) and (GridEcriture.Cells[1,i]=ColJ)
             and (GridEcriture.Cells[2,i]=ColCo) then
        begin
          GridEcriture.Cells[4,i] := MontantToStr(StrToFloat(GridEcriture.Cells[4,i]) + StrToFloat(ColDe));
          GridEcriture.Cells[5,i] := MontantToStr(StrToFloat(GridEcriture.Cells[5,i]) + StrToFloat(ColCr));
          bTrouve := True;
          break;
        end;
      end;
    end; }
  if (not bTrouve) then
  begin
    i := GridEcriture.RowCount - 1;
    GridEcriture.Cells[0, i] := ColDa;
    GridEcriture.Cells[1, i] := ColJ;
    GridEcriture.Cells[2, i] := ColCo;
    GridEcriture.Cells[3, i] := ColL;
    GridEcriture.Cells[4, i] := ColDe;
    GridEcriture.Cells[5, i] := ColCr;
    GridEcriture.RowCount := GridEcriture.RowCount + 1;
  end;
end;

function CompareSuivantCompteRef(Item1, Item2: Pointer): integer;
var
  Ecr1, Ecr2: TLigneEcriture;
begin
  Ecr1 := Item1;
  Ecr2 := Item2;
  if Ecr1.CompteRef > Ecr2.CompteRef then
    Result := 1
  else if Ecr1.CompteRef < Ecr2.CompteRef then
    Result := -1
  else
  begin
    if Ecr1.Compte > Ecr2.Compte then
      Result := 1
    else if Ecr1.Compte < Ecr2.Compte then
      Result := -1
    else
      Result := 0;
  end;
end;

procedure TIntegration.EclateLigneDC(L: TList);
var
  i: integer;
  ARecord, ANewRecord: TLigneEcriture;
begin
  if (L <> nil) and (L.Count > 0) then
  begin
    i := 0;
    while (i < (L.Count - 1)) do
    begin
      ARecord := L.Items[i];
      if ((ARecord.Debit <> 0) and (ARecord.Credit <> 0)) then
      begin
        ANewRecord := TLigneEcriture.Create;
        ANewRecord.Copy(ARecord);
        ARecord.Credit := 0;
        ANewRecord.Debit := 0;
        L.Insert(i + 1, ANewRecord);
      end;
      Inc(i, 1);
    end;
  end;
end;

procedure TIntegration.AfficheGrid;
var
  i: integer;
  ARecord: TLigneEcriture;
  TotalDebit, TotalCredit: double;
  Cumul: TCumul;
  Compte, Libelle, Cj: string;
begin
  TotalDebit := 0.0;
  TotalCredit := 0.0;
  if (fListeEcriture <> nil) and (fListeEcriture.Count > 0) then
  begin
    if NivDetail = 1 then
      fListeEcriture.Sort(CompareSuivantCompteRef);
    EclateLigneDC(fListeEcriture);
    GridEcriture.RowCount := 2;
    for i := 0 to fListeEcriture.Count - 1 do
    begin
      ARecord := fListeEcriture.Items[i];
      if fTypeEcr = toEcheance then
      begin
        if ARecord.Auxi = '' then
          Compte := ARecord.Compte
        else
          Compte := ARecord.Auxi;
        Cj := ARecord.CodeJournal;
        Libelle := ARecord.Libelle;
      end
      else
      begin
        if CodeJournal.Value = '' then
          Cj := ARecord.CodeJournal
        else
          Cj := CodeJournal.Value;
        if LibelleUnique.Text = '' then
          Libelle := ARecord.Libelle
        else
          Libelle := LibelleUnique.Text;
        Compte := ARecord.Compte;
      end;
      AjouteLigneDansGrid(DateToStr(ARecord.Date), Cj, Compte,
        Libelle, MontantToStr(ARecord.Debit),
        MontantToStr(ARecord.Credit));
      TotalDebit := TotalDebit + ARecord.Debit;
      TotalCredit := TotalCredit + ARecord.Credit;
    end;
    if GridEcriture.RowCount > 2 then
    begin
      GridEcriture.RowCount := GridEcriture.RowCount - 1;
      Cumul := CalculCumul(1);
    end
    else if GridEcriture.RowCount = 2 then
    begin
      for i := 0 to 5 do
        GridEcriture.Cells[i, 1] := '';
    end;
  end
  else
  begin
    VideGrid;
    Cumul := CalculCumul(0);
  end;
  DebitGen.Value := TotalDebit;
  CreditGen.Value := TotalCredit;
  DebitProg.Value := Cumul.Debit;
  CreditProg.Value := Cumul.Credit;
end;

procedure TIntegration.VideGrid;
var
  i: integer;
begin
  GridEcriture.RowCount := 2;
  for i := 0 to GridEcriture.ColCount - 1 do
    GridEcriture.Cells[i, 1] := '';
end;

procedure TIntegration.bExportClick(Sender: TObject);
var
  TypeFichier: integer;
begin
  if fListeEcriture = nil then
    exit;
  if (fTypeEcr <> toDotation) and (not ControleZoneEcheance) then
    exit;
  if (FtypeEcr = toDotation) and (not ControleZoneDotation) then
    exit;
  fIntegrationDB := false;
  TypeFichier := ChoixFichierExport(fFichierExport);
  if TypeFichier = 2 then
    GenereFichierMouvementSisco(nil, fListeEcriture, fFichierExport)
  else if (TypeFichier = 1) then
  begin
{$IFDEF SERIE1}
    if Blocage(['nrCloture', 'nrBatchImmo'], True, 'nrBatchImmo') then
      exit;
{$ELSE}
    if _Blocage(['nrCloture', 'nrBatchImmo'], True, 'nrBatchImmo') then
      exit;
{$ENDIF}
{$IFDEF SERIE1}
    if Transactions(IntegrationCompta, 3) <> oeOK then
      MessageAlerte(HM.Mess[9]);
{$ELSE}
    if Transactions(IntegreDansLaCompta, 3) <> oeOK then
      MessageAlerte(HM.Mess[9]);
{$ENDIF}
{$IFDEF SERIE1}
    Bloqueur('nrBatchImmo', False);
{$ELSE}
    _Bloqueur('nrBatchImmo', False);
{$ENDIF}
    AfficheCompteRenduExport(fFichierExport);
  end;
end;

procedure TIntegration.GridEcritureClick(Sender: TObject);
begin
  MajFiche;
end;

procedure TIntegration.MajFiche;
var
  Cumul: TCumul;
  ClasseCompte: string;
{$IFDEF SERIE1}
{$ELSE}
  i: integer;
{$ENDIF}
begin
  if (GridEcriture.RowCount < 1) or (GridEcriture.Row < 1) then
    exit;
  if fListeEcriture = nil then
    exit;
  if fListeEcriture.Count < 1 then
    exit;
  Cumul := CalculCumul(GridEcriture.Row);
  DebitProg.Value := Cumul.Debit;
  CreditProg.Value := Cumul.Credit;
  ClasseCompte := Copy(GridEcriture.Cells[COL_CPT, GridEcriture.Row], 1, 1);
{$IFDEF SERIE1}
  //A FAIRE
{$ELSE}
  for i := 0 to ImMaxAxe - 1 do
    if TLigneEcriture(fListeEcriture.Items[GridEcriture.Row - 1]).Axes[i] <> nil
      then
      break;
  BVentil.Enabled := not (i = ImMaxAxe);
{$ENDIF}
  BExport.Enabled := fListeEcriture.Count <> 0;
end;

procedure TIntegration.RecalculEcritures;
var
  i: integer;
  ParamEcr, ParamEch, ParamPai: TParamEcr;
  dtEchMin, dtEchMax: TDateTime;
begin
  Screen.Cursor := crHourglass;
  VideListeEcritures(fListeEcriture);
  ParamEcr := TParamEcr.Create;
  if fTypeEcr = toDotation then
  begin
    if rbDetail.Checked then
      NivDetail := 2
    else if rbGroupe.Checked then
      NivDetail := 1
    else
      NivDetail := 0;
    ParamEcr.Journal := CodeJournal.Value;
    ParamEcr.Libelle := LibelleUnique.Text;
    ParamEcr.Date := StrToDate(DateGeneration.Text);
    ParamEcr.DateCalcul := StrToDate(DateCalcul.Text);
    ParamEcr.bDotation := True;
    if fListeImmo = nil then
    begin
      InitMove(1, '');
      CalculEcrituresDotation(fListeEcriture, '', ParamEcr, NivDetail);
      FiniMove;
    end
    else
    begin
      InitMove(fListeImmo.Count, '');
      for i := 0 to fListeImmo.Count - 1 do
        CalculEcrituresDotation(fListeEcriture, fListeImmo[i], ParamEcr,
          NivDetail);
      FiniMove;
    end;
    AfficheGrid;
    ParamEcr.Free;
  end
  else if fTypeEcr = toEcheance then
  begin
    if rbRecapEch.Checked then
      NivDetail := 2
    else if rbSynthEch.Checked then
      NivDetail := 0;
    ParamPai := TParamEcr.Create;
    ParamEch := TParamEcr.Create;
    if cbEcrEcheance.Checked then
    begin
      ParamEch.Journal := JEcheance.Value;
{$IFDEF SERIE1}
      ParamEch.CompteRef := CTva.Text;
{$ELSE}
      if cbGestionTVA.Checked then
        ParamEch.CompteRef := 'X'
      else
        ParamEch.CompteRef := '-';
{$ENDIF}
    end
    else
    begin
      ParamEch.Journal := '';
      ParamEch.CompteRef := '';
    end;
    if cbEcrPaiement.Checked then
    begin
      ParamPai.Journal := JBanque.Value;
      ParamPai.CompteRef := CBanque.Text;
    end
    else
    begin
      ParamPai.Journal := '';
      ParamPai.CompteRef := '';
    end;
    ParamEch.bDotation := False;
    ParamPai.bDotation := False;
    dtEchMin := StrToDate(DateDebutEch.Text);
    dtEchMax := StrToDate(DateFinEch.Text);
    if fListeImmo = nil then
    begin
      InitMove(1, '');
      CalculEcrituresEcheances(fListeEcriture, '', ParamEch, ParamPai,
        NivDetail, dtEchMin, dtEchMax);
      FiniMove;
    end
    else
    begin
      InitMove(fListeImmo.Count, '');
      for i := 0 to fListeImmo.Count - 1 do
        CalculEcrituresEcheances(fListeEcriture, fListeImmo[i], ParamEch,
          ParamPai, NivDetail, dtEchMin, dtEchMax);
      FiniMove;
    end;
    AfficheGrid;
    ParamPai.Free;
    ParamEch.Free;
  end;
  Screen.Cursor := SyncrDefault;
end;

procedure TIntegration.BRecalculClick(Sender: TObject);
begin
  try
    hent1.EnableControls(self, false);
    PendantTraitment := true;
    BougeFocus;
    if (fTypeEcr = toEcheance) and (not ControleZoneEcheance) then
      exit;
    if (FtypeEcr = toDotation) and (not ControleZoneDotation) then
      exit;
    fIndiceImmoEnCours := 0;
    RecalculEcritures;
    fbModifie := false;
    MajFiche;
    BRecalculDot.Glyph := iCritGlyph.Picture.BitMap;
  finally
    hent1.EnableControls(self, true);
    PendantTraitment := false;
  end;
end;

procedure TIntegration.DateEnter(Sender: TObject);
begin
  CurDate := THCritMaskEdit(Sender).Text;
  fCurVal := CurDate;
end;

procedure TIntegration.CodeJournalEnter(Sender: TObject);
begin
  CurJournal := CodeJournal.Value;
  fCurVal := CurJournal;
end;

procedure TIntegration.LibelleUniqueEnter(Sender: TObject);
begin
  CurLibelle := TEdit(Sender).Text;
  fCurVal := CurLibelle;
end;

procedure TIntegration.LibelleUniqueExit(Sender: TObject);
begin
  if (fCurVal <> TEdit(Sender).Text) then
    fbModifie := True;
end;

procedure TIntegration.OnDetailClick(Sender: TObject);
begin
  if (rbSynth.Checked) or (rbSynthEch.Checked) then
    NivDetail := 0
  else if (rbGroupe.Checked) then
    NivDetail := 1
  else
    NivDetail := 0;
  fbModifie := True;
end;

procedure TIntegration.BImprimerClick(Sender: TObject);
begin
{$IFDEF eAGLCLient}
  PrintDBGrid(caption, GridEcriture.name, '');
{$ELSE}
  PrintDBGrid(GridEcriture, nil, Caption, '');
{$ENDIF eAGLCLient}
end;

procedure TIntegration.DessineCell(Acol, ARow: LongInt; Canvas: TCanvas; AState:
  TGridDrawState);
var
  CanvasPrec: TCanvas;
begin
  if PendantTraitment then
    exit;
  if (fListeEcriture = nil) or (fListeEcriture.Count <= 0) then
    exit;
  if ARow > 1 then
  begin
    CanvasPrec := TCanvas.Create;
    Canvas.Brush.Color := clWindow; //EPZ 30/10/00
    GridEcriture.GetCellCanvas(ACol, Arow - 1, CanvasPrec, AState);
    if TLigneEcriture(fListeEcriture.Items[ARow - 1]).Piece =
      TLigneEcriture(fListeEcriture.Items[ARow - 2]).Piece then
      Canvas.Brush.Color := CanvasPrec.Brush.Color
    else
    begin
      if (gdSelected in AState) then
        Canvas.Brush.Color := clHighLight
      else
      begin
        //EPZ 30/10/00
        //        if CanvasPrec.Brush.Color = clNone then Canvas.Brush.Color := AltColors[V_PGI.NumAltCol]
        //        else Canvas.Brush.Color := clNone;
        if CanvasPrec.Brush.Color = clWindow then
          Canvas.Brush.Color := AltColors[V_PGI.NumAltCol]
        else
          Canvas.Brush.Color := clWindow;
        //EPZ 30/10/00
      end;
    end;
    CanvasPrec.Free;
  end
  else if ARow = 1 then
  begin
    if (gdSelected in AState) then
      Canvas.Brush.Color := clHighLight
        //EPZ 30/10/00
    else
      Canvas.Brush.Color := clWindow;
    //    else Canvas.Brush.Color := clNone;
    //EPZ 30/10/00
  end;
end;

procedure TIntegration.cbEcrEcheanceClick(Sender: TObject);
begin
  JEcheance.Enabled := cbEcrEcheance.Checked;
  tJEcheance.Enabled := cbEcrEcheance.Checked;
  CTva.Enabled := cbEcrEcheance.Checked;
  tCTva.Enabled := cbEcrEcheance.Checked;
  fbModifie := True;
end;

procedure TIntegration.cbEcrPaiementClick(Sender: TObject);
begin
  JBanque.Enabled := cbEcrPaiement.Checked;
  tJBanque.Enabled := cbEcrPaiement.Checked;
  tCBanque.Enabled := cbEcrPaiement.Checked;
  fbModifie := True;
end;

procedure TIntegration.DateKeyPress(Sender: TObject;
  var Key: Char);
begin
  ParamDate(Self, Sender, Key);
end;

function IsValidDateEcriture(sDate: string): boolean;
begin
  Result := FALSE;
  if IsValidDate(sDate) then
    if (StrToDate(sDate) >= VHImmo^.Encours.Deb) and (StrToDate(sDate) <=
      VHImmo^.Encours.Fin) then
      Result := TRUE;
end;

procedure TIntegration.BVentilClick(Sender: TObject);
begin
{$IFDEF SERIE1}
  // A FAIRE
{$ELSE}
  ImVisuAnalytique(GridEcriture.Cells[2, GridEcriture.Row],
    TLigneEcriture(fListeEcriture.Items[GridEcriture.Row - 1]).Axes,
    StrToFloat(GridEcriture.Cells[4, GridEcriture.Row]),
    StrToFloat(GridEcriture.Cells[5, GridEcriture.Row]));
{$ENDIF}
end;

procedure TIntegration.BValiderClick(Sender: TObject);
var
  mr: integer;
{$IFDEF SERIE1}
{$ELSE}
  i: integer;
  ARecord: TCRPiece;
{$ENDIF}
begin
  fIntegrationDB := True;
  BougeFocus;
  if fbCompta = false then
  begin
    ModalResult := mrYes;
    exit;
  end;
  if fbModifie then
  begin
    HM.Execute(13, Caption, '');
    exit;
  end;
  // CA - 29/04/2002 - Contrôle d'existence des comptes
  if not ControleExistenceCompte then
    exit;
  if (fTypeEcr = toDotation) and (not ControleZoneDotation) then
    exit;
  if (fTypeEcr = toEcheance) and (not ControleZoneEcheance) then
    exit;
  if fListeEcriture.Count <= 0 then
    exit;
  mr := HM.Execute(8, Caption, '');
  if mr = mrYes then
  begin
{$IFDEF SERIE1}
{$ELSE}
    if (rbGroupe.Checked) then
      if
        (PGIAsk('En écritures groupées, les pièces sont générées sur l''établissement principal.#10#13Voulez-vous continuer ?', Caption) <> mrYes) then
        exit;
{$ENDIF}
{$IFDEF SERIE1}
    if Blocage(['nrCloture', 'nrBatchImmo'], True, 'nrBatchImmo') then
      exit;
{$ELSE}
    if _Blocage(['nrCloture', 'nrBatchImmo'], True, 'nrBatchImmo') then
      exit;
{$ENDIF}
{$IFDEF SERIE1}
    if Transactions(IntegrationCompta, 3) <> oeOK then
{$ELSE}
    if Transactions(IntegreDansLaCompta, 3) <> oeOK then
{$ENDIF}
    begin
      MessageAlerte(HM.Mess[9]);
      mr := mrNone;
    end
    else
    begin
{$IFDEF SERIE1}
      HM.Execute(23, Caption, '');
      mr := mrOk;
{$ELSE}
      if (fCR <> nil) and (fCR.Count > 0) then
      begin
        AfficheCRIntegration(fCR);
        for i := 0 to (fCr.Count - 1) do
        begin
          ARecord := fCR.Items[i];
          ARecord.Free;
        end;
        fCR.Free;
        fCR := nil;
      end
      else
      begin
        if (fCR <> nil) then
          fCR.Free;
        HM.Execute(11, Caption, '');
      end;
{$ENDIF}
//PGR 27/01/2006 Nouveaux paramètres société pour intégration des écritures
//      SetParamSoc('SO_IMMOJALDOTDEF', CodeJournal.Value);
//      SetParamSoc('SO_IMMOJALECHDEF', JEcheance.Value);
//{$IFDEF SERIE1}
//{$ELSE}
//      SetParamSoc('SO_IMMOGERETVA', cbGestionTva.Checked);
//{$ENDIF}
      {$IFDEF SERIE1}
        SetParamSoc('SO_IMMOJALDOTDEF', CodeJournal.Value);
        SetParamSoc('SO_IMMOJALECHDEF', JEcheance.Value);
      {$ENDIF}

    end;
{$IFDEF SERIE1}
    Bloqueur('nrBatchImmo', False);
{$ELSE}
    _Bloqueur('nrBatchImmo', False);
{$ENDIF}
  end
  else
    mr := mrNone;
  if mr = mrCancel then
    mr := mrNone;
  ModalResult := mr;
end;

procedure TIntegration.MajTableEcheance;
var
  i: integer;
  Q, QImmo: TQuery;
begin
  if fListeImmo <> nil then
  begin
    InitMove(fListeImmo.Count, '');
    for i := 0 to fListeImmo.Count - 1 do
    begin
      MoveCur(False);
      Q := OpenSQL('SELECT * FROM IMMOECHE WHERE IH_IMMO="' + fListeImmo[i] + '"'
        +
        ' AND IH_DATE>="' + USDate(DateDebutEch) + '" AND IH_DATE<="' +
          USDate(DateFinEch) + '"' +
        ' AND (IH_INTEGREECH="-" OR IH_INTEGREECH="-")', False);
      if not Q.Eof then
      begin
        Q.First;
        while not Q.Eof do
        begin
          Q.Edit;
          if cbEcrEcheance.Checked then
            Q.FindField('IH_INTEGREECH').AsString := 'X';
          if cbEcrPaiement.Checked then
            Q.FindField('IH_INTEGREPAI').AsString := 'X';
          Q.Post;
          Q.Next;
        end;
      end;
      Ferme(Q);
    end;
    FiniMove;
  end
  else
  begin
    QImmo :=
      OpenSQL('SELECT I_IMMO,I_NATUREIMMO,I_ETAT FROM IMMO WHERE (I_NATUREIMMO="LOC" OR I_NATUREIMMO="CB") AND (I_ETAT <> "FER")', True);
    if not QImmo.Eof then
    begin
      InitMove(QImmo.RecordCount, '');
      QImmo.First;
      while not QImmo.Eof do
      begin
        MoveCur(False);
        Q :=
          OpenSQL('SELECT IH_DATE,IH_INTEGREECH,IH_INTEGREPAI FROM IMMOECHE WHERE IH_IMMO="'
          + QImmo.FindField('I_IMMO').AsString +
          '" AND IH_DATE>="' + USDate(DateDebutEch) + '" AND IH_DATE<="' +
            USDate(DateFinEch) + '"' +
          ' AND (IH_INTEGREECH="-" OR IH_INTEGREPAI="-")', False);
        if not Q.Eof then
        begin
          Q.First;
          while not Q.Eof do
          begin
            Q.Edit;
            if cbEcrEcheance.Checked then
              Q.FindField('IH_INTEGREECH').AsString := 'X';
            if cbEcrPaiement.Checked then
              Q.FindField('IH_INTEGREPAI').AsString := 'X';
            Q.Post;
            Q.Next;
          end;
        end;
        QImmo.Next;
        Ferme(Q);
      end;
    end;
    Ferme(QImmo);
    FiniMove;
  end;
end;

procedure TIntegration.OnEnterZone(Sender: TObject);
begin
  fCurVal := TEdit(Sender).Text;
end;

procedure TIntegration.OnExitZone(Sender: TObject);
begin
  if fCurVal <> TEdit(Sender).Text then
    fbModifie := True;
end;

procedure TIntegration.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F9:
      begin
        Key := 0;
        BRecalculClick(nil);
      end;
// BTY 06/06 Ouvrir F10 via le traitement d'intégration
// BTY 04/06 FQ 17775
    VK_F10:
      begin
        if fbCompta then
          begin
          Key := 0;
          BValiderClick(nil);
        end;
      end;
    {^P}80: if Shift = [ssCtrl] then
      begin
        Key := 0;
        BImprimerClick(nil);
      end;
  end;
end;

procedure TIntegration.BougeFocus;
var
  WC: TWinControl;
begin
  WC := ActiveControl;
  if WC <> nil then
  begin
    if Pages.CanFocus then
      FocusControl(Pages);
    if WC.CanFocus then
      FocusControl(WC);
  end;
end;

procedure TIntegration.JBanqueChange(Sender: TObject);
begin
  fbModifie := True;
  if JBanque.Value <> '' then
    CBanque.Text := GetCompteBanqueContrepartie;
end;

function TIntegration.ControleZoneEcheance: boolean;
begin
  if not (IsValidDate(DateDebutEch.Text)) then
  begin
    HM.Execute(21, Caption, '');
    result := false;
    exit;
  end;
  if not (IsValidDate(DateFinEch.Text)) then
  begin
    HM.Execute(22, Caption, '');
    result := false;
    exit;
  end;
  if cbEcrEcheance.Checked then
  begin
    if (JEcheance.Value = '') then
    begin
      HM.Execute(16, Caption, '');
      result := false;
      exit;
    end;
{$IFDEF SERIE1}
    if (not (ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' +
      CTVA.Text + '"'))) then
    begin
      HM.Execute(17, Caption, '');
      result := false;
      exit;
    end;
{$ENDIF}
  end;
  if cbEcrPaiement.Checked then
  begin
    if (JBanque.Value = '') then
    begin
      HM.Execute(15, Caption, '');
      result := false;
      exit;
    end;
    if (not (ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' +
      CBanque.Text + '"'))) then
    begin
      HM.Execute(18, Caption, '');
      result := false;
      exit;
    end;
    if GetCompteBanqueContrepartie <> CBanque.Text then
    begin
      HM.Execute(18, Caption, '');
      result := false;
      exit;
    end;
  end;
  result := True;
end;

function TIntegration.ControleExistenceCompte: boolean;
var
  i: integer;
  ARecord: TLigneEcriture;
begin
  for i := 0 to fListeEcriture.Count - 1 do
  begin
    ARecord := fListeEcriture.Items[i];
    if (ARecord.Compte <> '') and (not Presence('GENERAUX', 'G_GENERAL',
      ARecord.Compte)) then
    begin
      PGIBox('Le compte ' + ARecord.Compte +
        ' n''existe pas. Veuillez le créer.', Self.Caption);
      Result := False;
      exit;
    end;
  end;
  result := True;
end;

function TIntegration.ControleZoneDotation: boolean;
begin
  if (CodeJournal.Value = '') then
  begin
    HM.Execute(10, Caption, '');
    result := false;
    exit;
  end;
  if not (IsValidDate(DateCalcul.Text)) then
  begin
    HM.Execute(19, Caption, '');
    result := false;
    exit;
  end;
  if not (IsValidDate(DateGeneration.Text)) then
  begin
    HM.Execute(20, Caption, '');
    result := false;
    exit;
  end;
  if ((StrToDate(DateCalcul.Text) > VHImmo^.Encours.Fin) or
    (StrToDate(DateCalcul.Text) < VHImmo^.Encours.Deb)
    or (StrToDate(DateGeneration.Text) > VHImmo^.Encours.Fin) or
      (StrToDate(DateGeneration.Text) < VHImmo^.Encours.Deb)) then
  begin
    HM.Execute(12, Caption, '');
    result := false;
    exit;
  end;
  result := True;
end;

function TIntegration.GetCompteBanqueContrepartie: string;
var
  Q: TQuery;
begin
  result := '';
  Q := OpenSQL('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="' +
    JBanque.Value + '"', True);
  if not Q.Eof then
    result := Q.FindField('J_CONTREPARTIE').AsString;
  Ferme(Q);
end;

procedure TIntegration.OnChangeCritere(Sender: TObject);
begin
  fbModifie := True;
  if Sender is THValComboBox then
  begin
    {$IFDEF SERIE1}
    rbSimu.Enabled :=  false ;
    rbSitu.Enabled :=  false ;
    {$ELSE}
    //mbo 15095 chgt check en rb
    rbSimu.Enabled :=  (GetColonneSQL('JOURNAL','J_MODESAISIE','J_JOURNAL="'+THValComboBox(Sender).Value+'"')='-');
    // mbo 15095 if not rbSimu.Enabled then bSimuDot.Checked := False;
    rbSitu.Enabled :=  (GetColonneSQL('JOURNAL','J_MODESAISIE','J_JOURNAL="'+THValComboBox(Sender).Value+'"')='-');
    if (rbSimu.Enabled = false) and (rbSitu.Enabled = false) then
        rbNor.checked := true;
    {$ENDIF SERIE1}
  end;
  BRecalculDot.Glyph := iCritGlyphModified.Picture.BitMap
end;

procedure TIntegration.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

{$IFDEF SERIE1}

procedure TIntegration.IntegrationCompta;
var
  i: integer;
  ARecord: TLigneEcriture; //ARecordAna : TAna;
  ToutesLesTobs, E_s, Ecr_{, Y_}: tob; //XVI 18/04/2006 FD 3978
  io: TIOS1;
  tDate: TDateTime;
  tJal: string;
//  CCNFCleAudit : string ;

  //Q: TQuery;  //XVI 18/04/2006 FD 3978
begin
  E_s := nil;
  EnableControls(Self, false);
  tDate := iDate1900;
  tJal := '';
  ToutesLesTobs := Tob.Create('UNKOWN', nil, -1);
//  CCNFCleAudit := CCNF_InscritPisteAudit('Intégration écritures provisoires - Immobilisations - Ecriture DAA ') ;
  for i := 0 to fListeEcriture.Count - 1 do
  begin
    ARecord := fListeEcriture.Items[i];
    if (tdate <> ARecord.Date) or (tJal <> ARecord.CodeJournal) then
      E_s := Tob.Create('ECRITURE_S', ToutesLesTobs, -1);
    tdate := ARecord.Date;
    tJal := ARecord.CodeJournal;
    Ecr_ := Tob.create('ECRITURE', E_s, -1);
    if ARecord.Auxi <> '' then
      ARecord.Compte := GetCollectif(ARecord.Auxi);
//    Ecr_.PutValue('E_GUIDSESSION', CCNFCleAudit);
    Ecr_.PutValue('E_GENERAL', ARecord.Compte);
    Ecr_.PutValue('E_AUXILIAIRE', ARecord.Auxi);
    Ecr_.PutValue('E_NUMEROPIECE', -1);
    Ecr_.PutValue('E_LIBELLE', ARecord.Libelle);
    Ecr_.PutValue('E_DATECOMPTABLE', ARecord.Date);
    Ecr_.PutValue('E_DATELETTRAGE', idate2099); //XVI 20/03/2006 FD 3643
    Ecr_.PutValue('E_JOURNAL', ARecord.CodeJournal);
    Ecr_.PutValue('E_DEBITDEV', ARecord.Debit);
    Ecr_.PutValue('E_CREDITDEV', ARecord.Credit);
    Ecr_.PutValue('E_TYPESAISIE', 8); //YCP 07/12/00 a voir type saisie => IMMO
    Ecr_.PutValue('E_DEVISE', V_PGI.DevisePivot);
    Ecr_.PutValue('E_MODEP', GetModeReglementTiers(ARecord.Auxi));
    //YCP 10/02/03
    //XVI 18/04/2006 FD 3978 début
    {Q := OpenSQL('SELECT I_TABLE9 FROM IMMO WHERE I_IMMO="' + ARecord.CodeImmo +
      '"', true);

    if ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' +
      ARecord.Compte + '" AND G_VENTILABLE="X"')
      and not Q.eof and (trim(Q.Fields[0].asstring) <> '') then
    begin
      Ecr_.PutValue('E_ANA', 'X');
      Ecr_.AddChampSupValeur('IMMO',ARecord.CodeImmo) ;
      Y_ := Tob.Create('ANALYTIQ', Ecr_, -1);
      Y_.PutValue('Y_GENERAL', ARecord.Compte);
      Y_.PutValue('Y_SECTION', trim(Q.Fields[0].asstring));
      Y_.PutValue('Y_DEBITDEV', ARecord.Debit);
      Y_.PutValue('Y_CREDITDEV', ARecord.Credit);
    end} //XVI 18/04/2006 FD 3978
    //XVI 25/04/2006  Mono et Multi ventilation.... début
    if VS1Cpta.Y_Saisie then
    begin
      Ecr_.AddChampSupValeur('IMMO',ARecord.CodeImmo) ;
      Y_Immo(Ecr_) ;
    end ;
    {if VS1Cpta.Y_analytiqueAdvanced and
       ExisteSql('select G_GENERAL from GENERAUX where G_GENERAL="' + ARecord.Compte + '" and G_VENTILABLE="X"') and
       ExisteSql('select I_IMMO from IMMO where I_IMMO="' + ARecord.CodeImmo+ '" and I_VENTILABLE="X"') then
    begin
      Ecr_.PutValue('E_ANA', 'X');
      Ecr_.AddChampSupValeur('IMMO',ARecord.CodeImmo) ;
      Y_ := Tob.Create('ANALYTIQ', Ecr_, -1);
      Y_.PutValue('Y_GENERAL', ARecord.Compte);
      Y_.PutValue('Y_SECTION', trim(Q.Fields[0].asstring));
      Y_.PutValue('Y_DEBITDEV', ARecord.Debit);
      Y_.PutValue('Y_CREDITDEV', ARecord.Credit);
    end
    else
    begin
      Ecr_.PutValue('E_ANA', '-');
    end;} //XVI 25/04/2006 FD 3978 mono et multiventilation...
    //ferme(Q);
    //XVI 18/04/206 FD 3978 fin
  end;

  io := TIOS1.create;
  io.IsImport := True;
  io.ArreterSiErreur := true;
  for i := 0 to ToutesLesTobs.Detail.Count - 1 do
  begin
    E_s := ToutesLesTobs.Detail[i];
    if (io.DispatchInputTob(E_S) = 0) then
    begin
      if (fTypeEcr = toEcheance) then
        MajTableEcheance;
    end
    else
    begin
      PgiBox(io.ErreurIO.Libelle, 'Erreur');
      break;
    end;
  end;
  io.Free;
  ToutesLesTobs.Free;
  EnableControls(Self, true);
end;
{$ELSE}

procedure TIntegration.IntegreDansLaCompta;

  procedure VentilAna(T: TOB; ARecord: TLigneEcriture);
  var
    TobAna: TOB;
    ARecordAna: TAna;
    j, k: integer;
    bType : integer;
  begin
    // CA - 25/11/2005 - FQ 15095 non validée - Pour répercuter la QUALIFPIECE sur l'ANALYTIQ
    if (rbsimu.checked) or (bSimuEche.Checked) then bType := 1
    else if rbsitu.checked then bType := 2
    else btype := 0;

    for j := 0 to ImMaxAxe - 1 do
    begin
      if ARecord.Axes[j] <> nil then
      begin
        for k := 0 to ARecord.Axes[j].Count - 1 do
        begin
          ARecordAna := ARecord.Axes[j].Items[k];
          TobAna := TOB.Create('ANALYTIQ', T, -1);
          InitFromListeAna(TobAna, ARecordAna, T);
          TobAna.PutValue('Y_ETABLISSEMENT', ARecord.Etablissement);
          //mbo 15095 InitCommunAna(TobAna, T.GetValue('E_QUALIFPIECE') = 'S');
          InitCommunAna(TobAna, bType);
          TobAna.PutValue('Y_NUMLIGNE', T.GetValue('E_NUMLIGNE'));
          TobAna.PutValue('Y_DATEMODIF', nowH);
          TobAna.PutValue('Y_NUMEROPIECE', T.GetValue('E_NUMEROPIECE'));
          TobAna.PutValue('Y_NUMVENTIL', k + 1);
          TobAna.PutValue('Y_AXE', 'A' + IntToStr(j + 1));
        end;
        T.PutValue('E_ANA', 'X');
      end;
    end;
  end;

  // MVG 27/03/2006 - FQ17641 - ajout du type de journal
  procedure EnregistrePiece(T: TOB; LGene: TZCompte; LAuxi: TZTiers; btype : integer);
  var
    stRegimeTVA, stTVA: string;
    St: string;
    j: integer;
    ALigneCR: TCRPiece;
  begin
    { On met à jour les informations de TVA pour la dernière pièce traitée }
    stRegimeTVA := '';
    stTVA := '';
    for j := 0 to T.Detail.Count - 1 do
    begin
      if (T.Detail[j].GetValue('E_AUXILIAIRE') <> '') and (stRegimeTVA = '')
        then
      begin
        St := T.Detail[j].GetValue('E_AUXILIAIRE');
        LAuxi.GetCompte(St);
        stRegimeTVA := LAuxi.GetValue('T_REGIMETVA', 0);
      end
      else if (T.Detail[j].GetValue('E_GENERAL') <> '') and (stTVA = '') then
        stTVA :=
          LGene.GetCodeTVAPourUnCompte(T.Detail[j].GetValue('E_GENERAL'));
    end;
    if (stTVA <> '') or (stRegimeTVA <> '') then
    begin
      for j := 0 to T.Detail.Count - 1 do
      begin
        T.Detail[j].PutValue('E_REGIMETVA', stRegimeTVA);
        T.Detail[j].PutValue('E_TVA', stTVA);
      end;
    end;
    { Enregistrement }
    if fIntegrationDB then
    begin
      T.InsertDB(nil);
      // modif mvg - fq 17641 - pb en intégration d'écriture de situation
      // btype = 0 = écritures de dotation - 1 = écritures de simu - 2 =
      if bType = 0 then MajSolde(T.Detail[0].GetValue('E_NUMEROPIECE'), T);
    end;
    if fTypeEcr = toEcheance then
      MajTableEcheance;
    { Mise à jour du compte-rendu }
    ALigneCR := TCRPiece.Create;
    ALigneCR.NumPiece := T.Detail[0].GetValue('E_NUMEROPIECE');
    ALigneCR.Journal := T.Detail[0].GetValue('E_JOURNAL');
    ALigneCR.Date := T.Detail[0].GetValue('E_DATECOMPTABLE');
    ALigneCR.QualifPiece := T.Detail[0].GetValue('E_QUALIFPIECE');
    ALigneCR.Debit := T.Somme('E_DEBIT', ['E_NUMEROPIECE'],
      [T.Detail[0].GetValue('E_NUMEROPIECE')], TRUE);
    ALigneCR.Credit := T.Somme('E_CREDIT', ['E_NUMEROPIECE'],
      [T.Detail[0].GetValue('E_NUMEROPIECE')], TRUE);
    fCR.Add(ALigneCR);
    { Suppression des lignes }
    if fIntegrationDB then
      T.ClearDetail;
  end;

  function TrouveModeReglement(TMr: TOB; stModeRegle: string): string;
  var
    T: TOB;
  begin
    T := TMr.FindFirst(['MR_MODEREGLE'], [stModeRegle], False);
    if T <> nil then
      Result := T.GetValue('MR_MP1')
    else
      Result := '';
  end;

var
  ARecord: TLigneEcriture;
  stJournal, stModeSaisie: string;
  LastDate: TDateTime;
  // ajout mbo 15095
  bType: integer;
  TMere, TEcr: TOB;
  TExportPGI: TOB;
  LesGeneraux: TZCompte;
  LesAuxiliaires: TZTiers;
  i: integer;
  LastPeriode: integer;
  iNumPiece, iNumLigne: integer;
  TModeRegle: TOB;
//  {$IFDEF SERIE1}
//  CCNFCleAudit : string ;
//  {$ENDIF}
begin
  if fListeEcriture.Count = 0 then
    exit;
  { Récupération des éléments communs à toutes les écritures }
  if not fIntegrationDB then
    TExportPGI := TOB.Create('', nil, -1)
  else
    TExportPGI := nil;
  fCR := TList.Create;

  //mbo 15095 modif check en rb
  //bSimulation := (rbSimu.Checked) or (bSimuEche.Checked);
  // ajout type mbo 15095
  if (rbsimu.checked) or (bSimuEche.Checked) then
     bType := 1
  else if rbsitu.checked = true then
     bType := 2
  else
     btype := 0;

  ARecord := fListeEcriture.Items[0];
  stJournal := ARecord.CodeJournal;
  {$IFDEF SERIE1}
  stModeSaisie := '-' ;
  {$ELSE}
  stModeSaisie := GetColonneSQL('JOURNAL', 'J_MODESAISIE', 'J_JOURNAL="' +
    stJournal + '"');
  {$ENDIF SERIE1}
  LastPeriode := 0;
  LastDate := 0;
  iNumLigne := 1;
  iNumPiece := 0;
  TModeRegle := TOB.Create('', nil, -1);
  TModeRegle.LoadDetailFromSQL('SELECT * FROM MODEREGL');
  TMere := TOB.Create('', nil, -1);
  LesGeneraux := TZCompte.Create;
  LesAuxiliaires := TZTiers.Create;
//  {$IFDEF SERIE1}
//  CCNFCleAudit := CCNF_InscritPisteAudit('Intégration écritures provisoires - Immobilisations - Ecriture DAA ') ;
//  {$ENDIF SERIE1}
  for i := 0 to fListeEcriture.Count - 1 do
  begin
    ARecord := fListeEcriture.Items[i];
    if (((stModeSaisie = '-') and (ARecord.Date <> LastDate))
      // Rupture en mode pièce
      or ((stModeSaisie <> '-') and (ImGetPeriode(ARecord.Date) <> LastPeriode)))
        then // Rupture en bordereau ou libre
    begin { C'est donc une nouvelle pièce }
      { On l'enregistre }
      if LastPeriode <> 0 then // Sauf la première fois
      begin
        // MVG 27/03/2006 - FQ17641 - ajout du type de journal
        EnregistrePiece(TMere, LesGeneraux, LesAuxiliaires, btype);
        if not fIntegrationDB then
        begin
          while TMere.Detail.Count <> 0 do
            TMere.Detail[0].ChangeParent(TExportPGI, -1);
        end;
      end;
      { On calcule les éléments de la nouvelle pièce }

      // modif mbo - fq 17641 - pb en intégration d'écriture de situation
      // btype = 0 = écritures de dotation - 1 = écritures de simu - 2 = écritures de situ
      if bType = 0 then
        iNumPiece := GetNewNumJal(stJournal, True, ARecord.Date, '', stModeSaisie)
      else
        iNumPiece := GetNewNumJal(stJournal, False, ARecord.Date, '', stModeSaisie);

      iNumLigne := 1;
      LastDate := ARecord.Date;
      LastPeriode := ImGetPeriode(ARecord.Date);
    end;
    TEcr := TOB.Create('ECRITURE', TMere, -1);

    InitFromListe(TEcr, ARecord, stModeSaisie);
    // modif mbo 15095  InitCommun(TEcr, bSimulation);
    InitCommun(TEcr, bType);
    CSupprimerInfoLettrage(TEcr);
    LesGeneraux.GetCompte(ARecord.Compte);
    if ARecord.Auxi <> '' then
    begin
      LesAuxiliaires.GetCompte(ARecord.Auxi);
      if LesAuxiliaires.GetValue('T_LETTRABLE', 0) = 'X' then
      begin
        CRemplirInfoLettrage(TEcr);
        TEcr.PutValue('E_MODEPAIE', TrouveModeReglement(TModeRegle,
          LesAuxiliaires.GetValue('T_MODEREGLE', 0)));
      end;
      TEcr.PutValue('E_GENERAL', LesAuxiliaires.GetValue('T_COLLECTIF', 0));
    end
    else if LesGeneraux.IsLettrable then
    begin
      CRemplirInfoLettrage(TEcr);
      TEcr.PutValue('E_MODEPAIE', TrouveModeReglement(TModeRegle,
        LesGeneraux.GetValue('G_MODEREGLE')));
    end;
    TEcr.PutValue('E_NUMLIGNE', iNumLigne);
    if (stModeSaisie <> '-') then
      TEcr.PutValue('E_NUMGROUPEECR', 1);
    TEcr.PutValue('E_DATEMODIF', nowH);
    TEcr.PutValue('E_NUMEROPIECE', iNumPiece);
    TEcr.PutValue('E_ETABLISSEMENT', ARecord.Etablissement);
    TEcr.PutValue('E_ENCAISSEMENT', SENSENC(ARecord.Debit, ARecord.Credit));
//    {$IFDEF SERIE1}
//    TEcr.PutValue('E_GUIDSESSION', CCNFCleAudit);
//    {$ENDIF}
    TEcr.PutValue('E_IO', 'X');
    Inc(iNumLigne);
    { Mise à jour de l'analytique }
    VentilAna(TEcr, ARecord);
  end;
  // MVG 27/03/2006 - FQ17641 - ajout du type de journal
  EnregistrePiece(TMere, LesGeneraux, LesAuxiliaires, btype);
  if not fIntegrationDB then
  begin
    while TMere.Detail.Count <> 0 do
      TMere.Detail[0].ChangeParent(TExportPGI, -1);
    ExporteImmoFormatPGI(TExportPGI);
    TExportPGI.Free;
  end;
  TModeRegle.Free;
  LesGeneraux.Free;
  LesAuxiliaires.Free;
  TMere.Free;
end;

procedure TIntegration.InitCommun(OBEcr: TOB; btype: integer);
begin
  OBEcr.PutValue('E_EXERCICE', VHImmo^.EnCours.Code);
  if fTypeEcr = toEcheance then
    OBEcr.PutValue('E_NATUREPIECE', 'FF')
  else
    OBEcr.PutValue('E_NATUREPIECE', 'OD');

  //modif mbo n°15095
  if btype = 1 then
    OBEcr.PutValue('E_QUALIFPIECE', 'S')
  else
  if btype = 2 then
     OBEcr.PutValue('E_QUALIFPIECE', 'U')
  else
    OBEcr.PutValue('E_QUALIFPIECE', 'N');

  OBEcr.PutValue('E_DATECREATION', Date);
  OBEcr.PutValue('E_DEVISE', V_PGI.DevisePivot);
  OBEcr.PutValue('E_TAUXDEV', 1);
  OBEcr.PutValue('E_COTATION', 1);
  OBEcr.PutValue('E_CREERPAR', 'GEN');
  OBEcr.PutValue('E_QUALIFORIGINE', 'IMO');
  OBEcr.PutValue('E_ECRANOUVEAU', 'N');
  OBEcr.PutValue('E_ETATLETTRAGE', 'RI');
  {$IFDEF SERIE1} // MVG 10/05/2007
  OBEcr.PutValue('E_DATELETTRAGE',idate2099) ; //XVI 20/03/2006 FD 3643
  {$ENDIF SERIE1}
  OBEcr.PutValue('E_UTILISATEUR', V_PGI.User);
  OBEcr.PutValue('E_CONFIDENTIEL', '0');
  OBEcr.PutValue('E_SOCIETE', V_PGI.CodeSociete);
  // FQ 16298 - Champs obligatoire à renseigner
  OBEcr.PutValue('E_CONTROLETVA','RIE');
end;

procedure TIntegration.InitCommunAna(OBAna: TOB; bType: integer);
begin
  OBAna.PutValue('Y_EXERCICE', VHImmo^.EnCours.Code);
  if fTypeEcr = toEcheance then
    OBAna.PutValue('Y_NATUREPIECE', 'FF')
  else
    OBAna.PutValue('Y_NATUREPIECE', 'OD');
  // modif  mbo  15095
  if bType = 1 then
    OBAna.PutValue('Y_QUALIFPIECE', 'S')
  else if bType = 2 then
     OBAna.PutValue('Y_QUALIFPIECE', 'U')
  else
    OBAna.PutValue('Y_QUALIFPIECE', 'N');
  OBAna.PutValue('Y_DATECREATION', Date);
  OBAna.PutValue('Y_DEVISE', V_PGI.DevisePivot);
  OBAna.PutValue('Y_TAUXDEV', 1);
  OBAna.PutValue('Y_CREERPAR', 'GEN');
  OBAna.PutValue('Y_ECRANOUVEAU', 'N');
  OBAna.PutValue('Y_UTILISATEUR', V_PGI.User);
  OBAna.PutValue('Y_CONFIDENTIEL', '0');
  OBAna.PutValue('Y_SOCIETE', V_PGI.CodeSociete);
end;

procedure TIntegration.InitFromListe(OBEcr: TOB; ARecord: TLigneEcriture;
  ModeSaisie: string = '-');
begin
  OBEcr.PutValue('E_REFINTERNE', ARecord.Piece);
  OBEcr.PutValue('E_GENERAL', ARecord.Compte);
  OBEcr.PutValue('E_AUXILIAIRE', ARecord.Auxi);
  OBEcr.PutValue('E_LIBELLE', ARecord.Libelle);
  OBEcr.PutValue('E_DATECOMPTABLE', ARecord.Date);
  OBEcr.PutValue('E_DATETAUXDEV', ARecord.Date);
  OBEcr.PutValue('E_JOURNAL', ARecord.CodeJournal);
  OBEcr.PutValue('E_DEBIT', ARecord.Debit);
  OBEcr.PutValue('E_CREDIT', ARecord.Credit);
  OBEcr.PutValue('E_DEBITDEV', ARecord.Debit);
  OBEcr.PutValue('E_CREDITDEV', ARecord.Credit);
  OBEcr.PutValue('E_MODESAISIE', ModeSaisie);
  OBEcr.PutValue('E_EQUILIBRE', '-');
  OBEcr.PutValue('E_AVOIRRBT', '-');
  OBEcr.PutValue('E_ETAT', '0000000000');
  OBEcr.PutValue('E_TYPEMVT', 'DIV');
  if VHImmo^.AttribRibAuto then
    OBEcr.PutValue('E_RIB', GetRIBPrincipal(ARecord.auxi));
{$IFNDEF SPEC302}
  //  OBEcr.PutValue('E_PERIODE',GetPeriode(V_PGI.DateEntree)) ;
  OBEcr.PutValue('E_PERIODE', ImGetPeriode(ARecord.Date));
  OBEcr.PutValue('E_SEMAINE', NumSemaine(ARecord.Date));
{$ENDIF}
end;

procedure TIntegration.InitFromListeAna(OBAna: TOB; ARecord: TAna; OBEcr: TOB);
begin
  OBAna.PutValue('Y_SECTION', ARecord.Section);
  OBAna.PutValue('Y_LIBELLE', ARecord.Libelle);
  OBAna.PutValue('Y_POURCENTAGE', ARecord.TauxMontant);
  OBAna.PutValue('Y_LIBELLE', ARecord.Libelle);
  if OBECR.GetValue('E_DEBIT') <> 0 then
  begin
    OBAna.PutValue('Y_DEBIT', ARecord.Montant);
    OBAna.PutValue('Y_DEBITDEV', ARecord.Montant);
  end
  else
  begin
    OBAna.PutValue('Y_CREDIT', ARecord.Montant);
    OBAna.PutValue('Y_CREDITDEV', ARecord.Montant);
  end;
  OBAna.PutValue('Y_GENERAL', OBECR.GetValue('E_GENERAL'));
  OBAna.PutValue('Y_DATECOMPTABLE', OBECR.GetValue('E_DATECOMPTABLE'));
  OBAna.PutValue('Y_DATETAUXDEV', OBECR.GetValue('E_DATETAUXDEV'));
  OBAna.PutValue('Y_JOURNAL', OBECR.GetValue('E_JOURNAL'));
  OBAna.PutValue('Y_TOTALECRITURE', OBECR.GetValue('E_DEBIT') +
    OBECR.GetValue('E_CREDIT'));
  OBAna.PutValue('Y_TOTALDEVISE', OBECR.GetValue('E_DEBITDEV') +
    OBECR.GetValue('E_CREDITDEV'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 02/04/2004
Modifié le ... :   /  /
Description .. : - LG - 02/03/2004 - suprresion de malsoldespiece de aboutl
Mots clefs ... :
*****************************************************************}

procedure TIntegration.MajSolde(NumPiece: longint; OBEcr: TOB);
var
  M: RMVT; {défini dans SAISUTIL}
begin
(* FQ 14334 - CA - 26/08/2004 - Inutile : il faut parcourir toute la pièce pour mettre tous les soldes à jour
  bTrouve := false;
  if OBEcr.Detail.Count <= 0 then
    exit;
  for i := 0 to OBEcr.Detail.Count - 1 do
    if OBEcr.Detail[i].GetValue('E_NUMEROPIECE') = NumPiece then
    begin
      bTrouve := True;
      break;
    end;
  if not bTrouve then
    exit;
  *)
  if OBEcr.Detail.Count <= 0 then exit;
  {Renseigner ici un record de type RMVT qui identifie chauqe écriture (pièce) pour envoyer une procédure standard de mise à jour
  des sodles des comptes, auxi, sections et journaux
  OBEcr est une TOB de type écriture représentative d'une pièce}
  FillChar(M, Sizeof(M), #0);
  M.Etabl := VHImmo^.EtablisDefaut;
  M.Jal := OBEcr.detail[0].GetValue('E_JOURNAL');
  M.Exo := OBEcr.detail[0].GetValue('E_EXERCICE');
  M.CodeD := OBEcr.detail[0].GetValue('E_DEVISE');
  M.Simul := 'N';
  M.Nature := OBEcr.detail[0].GetValue('E_NATUREPIECE');
  M.DateC := OBEcr.detail[0].GetValue('E_DATECOMPTABLE');
  M.DateTaux := M.DateC;
  M.Num := OBEcr.detail[0].GetValue('E_NUMEROPIECE');
  M.TauxD := 1;
  {Envoi de la procédure de mise à jour}
  MajSoldesEcritureTOB(OBEcr,true) ;
   // MajDesSoldesTOB(OBEcr.detail[i], True); // à mettre en standard

  {Envoi de la procédure de dévalidation éventuelle de la période et/ou du journal}
  ADevalider(M.Jal, M.DateC);
end;

procedure TIntegration.ExporteImmoFormatPGI(OB: TOB);
var
  i, j, k: integer;
  OBEcr, OBAna: TOB;
  FichierIE: TextFile;
  SCegid: TStructureCegid;
  SEntete: TEnteteCegid;
  LigneEntete, LigneCegid: string;
begin
  OuvrirFichierIE(FichierIE, fFichierExport, 1025);
  CreeEnregIE(SEntete, SCegid);
  InitEnteteCegid(SEntete, 'CEGID', 'CEG', 'FEC');
  FormateEnteteCegid(LigneEntete, SEntete);
  EcrireLigneIE(FichierIE, LigneEntete);
  for i := 0 to OB.Detail.Count - 1 do
  begin
    OBEcr := OB.detail[i];
    InitStructureCegid(SCegid);
    VentileStructureCegid(SCegid, OBEcr);
    FormateStructureCegid(LigneCegid, SCegid);
    EcrireLigneIE(FichierIE, LigneCegid);
    for j := 0 to OBEcr.Detail.Count - 1 do
    begin
      OBAna := OBEcr.Detail[j];
      for k := 0 to OBAna.Detail.Count - 1 do
      begin
        InitStructureCegid(SCegid);
        VentileStructureCegid(SCegid, OBAna.Detail[k]);
        FormateStructureCegid(LigneCegid, SCegid);
        EcrireLigneIE(FichierIE, LigneCegid);
      end;
    end;
  end;
  FermerFichierIE(FichierIE);
  DetruitEnregIE(SEntete, SCegid);
end;

{$ENDIF}

procedure TIntegration.InitAddFiltre(T: TOB);
var
    Lines: HTStrings;
begin
    Lines := HTStringList.Create;
    SauveCritMemoire(Lines, Pages);
    FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
    Lines.Free;
end;

procedure TIntegration.InitGetFiltre(T: TOB);
var
    Lines: HTStrings;
begin
    Lines := HTStringList.Create;
    SauveCritMemoire(Lines, Pages);
    FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
    Lines.Free;
end;

procedure TIntegration.InitSelectFiltre(T: TOB);
var
    Lines: HTStrings;
    i:integer;
    stChamp, stVal:string;
begin
    if T = nil then exit;
    Lines := HTStringList.Create;
    for i := 0 to T.Detail.Count - 1 do
    begin
        stChamp := T.Detail[i].GetValue('N');
        stVal := T.Detail[i].GetValue('V');
        Lines.Add(stChamp + ';' + stVal);
    end;
    VideFiltre(FFiltres, Pages, False);
    ChargeCritMemoire(Lines, Pages);
    Lines.Free;
end;

procedure TIntegration.ParseParamsFiltre(Params: HTStrings);
var
    T:TOB;
begin
    FListeByUser.AddVersion;
    T := FListeByUser.Add;
  //en position 0 de Params se trouve le nom du filtre
    T.PutValue('NAME', XMLDecodeSt(Params[0]));
    T.PutValue('USER', '---');
    Params.Delete(0);
    FListeByUser.AffecteTOBFiltreMemoire(T, Params);
end;

procedure TIntegration.UpgradeFiltre(T: TOB);
begin
//
end;

procedure TIntegration.BDUPFILTREClick(Sender: TObject);
begin
  FListeByUser.Duplicate;
end;

end.

