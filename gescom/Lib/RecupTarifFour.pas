unit RecupTarifFour;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  Hent1, UIUtil, HPanel, Mask, UTOB,Math,ParamSoc,
{$IFDEF YOYO}
  Rapport,
{$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE} // EAGLCLIENT
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF} // EAGLCLIENT
  HStatus, UtilArticle, Ed_Tools,
  TarifAutoFour, TarifUtil, UtilGC,UtilTarif, TntStdCtrls, TntComCtrls,
  TntExtCtrls;

procedure EntreeRecupTarifFour (StParFou : string);

type
  TFRecupTarifFour = class(TFAssist)
    TINTRO: THLabel;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    PTITRE: THPanel;
    HLabel1: THLabel;
    GBCreationFiches: TGroupBox;
    CBCreationCatalogue: TCheckBox;
    CBCreationArticle: TCheckBox;
    TCreationCatalogue: THLabel;
    TCreationArticle: THLabel;
    TNomChamp: THLabel;
    TLongueur: THLabel;
    TOffset: THLabel;
    TCritere: THLabel;
    TBorneSup: THLabel;
    TBorneInf: THLabel;
    HLabel2: THLabel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    ListRecap: TListBox;
    TRecap: THLabel;
    GBPremier: TGroupBox;
    HChamp1: THValComboBox;
    HLongueur1: THCritMaskEdit;
    HOffset1: THCritMaskEdit;
    HBorneInf1: THCritMaskEdit;
    HBorneSup1: THCritMaskEdit;
    GBDeuxieme: TGroupBox;
    HChamp2: THValComboBox;
    HLongueur2: THCritMaskEdit;
    HOffset2: THCritMaskEdit;
    HBorneInf2: THCritMaskEdit;
    HBorneSup2: THCritMaskEdit;
    GBTroisieme: TGroupBox;
    HChamp3: THValComboBox;
    HLongueur3: THCritMaskEdit;
    HOffset3: THCritMaskEdit;
    HBorneInf3: THCritMaskEdit;
    HBorneSup3: THCritMaskEdit;
    TRGF_PARFOU: THLabel;
    TGRF_TIERS: THLabel;
    GRF_PARFOU: THCritMaskEdit;
    GRF_TIERS: THCritMaskEdit;
    LGRF_TIERS: THLabel;
    HRecap: THMsgBox;
    OpenDialogButton: TOpenDialog;
    TGRF_FICHIER: THLabel;
    GRF_FICHIER: THCritMaskEdit;
    TDATESUP: THLabel;
    DATESUP: THCritMaskEdit;
    GRF_MULTIFOU: THCheckbox;
    CBRecInit: TCheckBox;
    chkFamSsFam: TCheckBox;
    TMAJFamSsFam: THLabel;
    procedure bFinClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure GRF_FICHIERElipsisClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure HChamp1Change(Sender: TObject);
    procedure HChamp2Change(Sender: TObject);
    procedure HChamp3Change(Sender: TObject);
    procedure DATESUPDblClick(Sender: TObject);
  private
    { Déclarations privées }
    TheFournisseur : string;
    TobArticle      : TOB;
    TobCatalogu     : TOB;
    TOBEnt          : TOB;
    TOBLig          : TOB;
    TOBArticleCompl : TOB;
    TOBCB           : TOB; //Code Barre
    TOBErreurs      : TOB;
{$IFDEF BTP}
    TOBTiers : TOB;
{$ENDIF}
{$IFDEF YOYO}
    TpsLect : double;
    TpsTrt : double;
    TpsRec : double;
    TpsTot : double;
{$ENDIF}
    StEnreg           : String;
    CodeArticle       : string;

    IndexParCatalogu  : Integer;
    IndexParArticle   : integer;
    //
    BSTop             : boolean;
    //
    TOBTablettes      : TOB;
    TOBP              : TOB;
    TOBChampsArt      : TOB;
    TOBChampsCatalogu : TOB;
    TOBChampsArtC     : TOB;
    TOBChampsCB       : TOB;
    //
    ListChampsArt     : String;
    ListChampsCat     : String;
    ListChampsArtc    : string;
    ListChampsCB      : String;

    // Initialisations
    procedure ActiveChamp (HCrit : THCritMaskEdit);
    procedure DepileTOBLigne ;
    procedure DesactiveChamp (HCrit : THCritMaskEdit);
    Procedure InitialiseFiche;
    procedure RempliCombo;

    // Controles
    function ControleChampOk (HChamp : THValComboBox; ValChamp : string;
                              stLongueur, stOffset, stBorneInf, stBorneSup : string) : boolean;
    function ControleEnreg : boolean;
    function ControleSaisie : boolean;
    function ControleSelection (stChamp, stBorneSup, stBorneInf, stLongueur,
                                stOffset : string; Message1, Message2 : integer) : boolean;
    // Traitements
    procedure CompleteChampArticle (Index : integer);
    procedure CompleteChampCatalogu (Index : integer; Code : string);
    procedure CreationElementTob (stTable : string; IndexPar : integer; TobTable : Tob;
                                  CBCreation : boolean; var Index : integer);
    procedure MiseajourTables;
    procedure ListeRecap;
    procedure RecapChamp(stChamp, stLongueur, stOffset, stBorneInf, stBorneSup : string);
    procedure RecupereTarif;
    procedure RecupereValChamp (TobPar : TOB; Index : integer; var ValChamp : string);
    procedure RecupValChampEnregVar (var ValChamp : string; IndexChamp : integer);
    procedure RenseigneJournalEvenement (ioerr : TIOErr ; Partiel : boolean = false);
    procedure TraiteChamp (IndexTobTable : integer; ValChamp, Champ : string; TobTable,TOBParam : TOB;IndexArt : integer = -1);
    function  TraiteEnregistrement (var CodeErreur : integer; var Texte : string) : boolean;
    // -- Modif BTP
    procedure SetDetailPrixPose (IndexArt : integer);
    function GenereLaValeur(ValChamp: string; TOBParam: TOB): double;
    function FindTablette(champ: string; TobParam: TOB): TOB;
    procedure AddTablettes(champ, valeur: string ; TOBParam ,TobTable: TOB; IndextobTable : integer);
    function RecupTablettes(champ : string;TOBParam: TOB;Var TheTable : string; Var TheCodeRec : string ; var TheType : string; Var LePrefixe: string) : tob;
    procedure CumulTablette(laTabl: TOB; champ, valeur: string;TobTable: TOB; IndextobTable : integer);
    procedure FindAjouteOnIT (ThisTablette : TOB; LaTable, LePrefixe, Letype, Valeur : string;TobTable: TOB; IndextobTable : integer);
    procedure MAJTablette(TheTablette: TOB);
    procedure GetInfoProfil(TobArt: TOB; Profil: string);
    procedure MajLibelletablette(TheTablette: TOB);
    function RecupereValChampTablette(TheEnreg: string; longueur, offset,index: integer): string;
    function RecupValChampEnregVarTablette(TheEnreg: string;IndexChamp: integer): string;
    procedure MajInfoLig(CodeArt, Libelle: string);
    function InFamille(recherche, ChaineControle: string): boolean;
    function EnregistreTables: TIOErr;
    procedure VideTobs;
{$IFDEF BTP}
    procedure DefinieValorisation(TOBRef : TOB);
    function FindCatalogue(Article, Fournisseur: string): TOB;
{$ENDIF}
    procedure MAjPRixArticle;
    procedure RemplitChampsTob;
  	procedure ElimineDoublonsCatalog (TobCatalogu : TOB);
    procedure ElimineDoublonsArticles(TobArticles: TOB);
    procedure VerifSousFamilleTarif;
    function GetFournisseurFromEnreg: string;
    function GetArticleFromEnreg: string;
    procedure AffichageErreurs;
    procedure EnregistreErreur(Erreur, LigneDep : integer; ligneFin: integer=-1; Texte:string='');
    procedure CreationErreur (stErreur : string; LigneDep : Integer; LigneFin : integer =-1);
    //FV1 : 21/01/2016 - FS#1854 - En import catalogue, les tables BTFAMILLETARIF et BTSOUSFAMILLETARIF ne sont pas alimentées
    procedure MiseAJourCodeBarre(TOBA: TOB);
    procedure MiseAJourFamilleTarif(TOBA: TOB);
    procedure MiseAJourSsFamilleTarif(TOBA: TOB);
    function RechercheLibelleTablette(Thecode: string; TheTablette: TOB): string;
    // Utils
//    function  ExtractLibelle ( St : string) : string;
  public
    { Déclarations publiques }
    TobParFou, TobParFouLig : TOB;
  end;

const MAXTRAITENREG = 1000;

implementation
uses TiersUtil,RECALCPIECE_RAP_TOF
   ,CbpMCD
   ,CbpEnumerator
;

{$R *.DFM}

procedure EntreeRecupTarifFour (StParFou : string);
var FF : TFRecupTarifFour;
    PPANEL  : THPanel ;
BEGIN
  SourisSablier;
  FF := TFRecupTarifFour.Create(Application) ;
  FF.GRF_PARFOU.Text := StParFou;

  PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
  if PPANEL = Nil then        // Le PANEL est le premier ecran affiché
  BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
    SourisNormale ;
  end else
  BEGIN
    InitInside (FF, PPANEL);
    FF.Show ;
  end ;
end ;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

procedure TFRecupTarifFour.ActiveChamp (HCrit : THCritMaskEdit);
begin
  HCrit.Enabled := True;
  HCrit.Color := clWindow;
end;

Procedure TFRecupTarifFour.DepileTOBLigne ;
var Index : integer;
BEGIN
  for Index := TobParFouLig.Detail.Count - 1 Downto 0 do
  BEGIN
  	TobParFouLig.Detail[Index].Free ;
  end;
end;

procedure TFRecupTarifFour.DesactiveChamp (HCrit : THCritMaskEdit);
begin
  HCrit.Enabled := False;
  HCrit.Color := clActiveBorder;
  HCrit.Text := '';
end;

Procedure TFRecupTarifFour.InitialiseFiche;
BEGIN
  TobParFou.InitValeurs;
  TobParFouLig.InitValeurs;

  LoadLesTobParFou (GRF_PARFOU.Text, TobParFou, TobParFouLig);
  RempliCombo;

  HChamp1.Text := 'Aucun';
  HChamp2.Enabled := False;
  HChamp3.Enabled := False;
  HChamp2.Text := 'Aucun';
  HChamp3.Text := 'Aucun';

  DesactiveChamp (HLongueur1);
  DesactiveChamp (HLongueur2);
  DesactiveChamp (HLongueur3);

  DesactiveChamp (HOffset1);
  DesactiveChamp (HOffset2);
  DesactiveChamp (HOffset3);

  DesactiveChamp (HBorneInf1);
  DesactiveChamp (HBorneInf2);
  DesactiveChamp (HBorneInf3);

  DesactiveChamp (HBorneSup1);
  DesactiveChamp (HBorneSup2);
  DesactiveChamp (HBorneSup3);

  CBCreationArticle.Checked := False;
  CBCreationCatalogue.Checked := False;

  GRF_TIERS.Text := TobParFou.GetValue ('GRF_TIERS');
  LGRF_TIERS.Caption := GetValChamp ('TIERS','T_LIBELLE','T_TIERS="' + GRF_TIERS.Text + '"');
  GRF_FICHIER.Text := TobParFou.GetValue ('GRF_FICHIER');
  GRF_MULTIFOU.Checked := TobParFou.GetBoolean ('GRF_MULTIFOU');
  if not GRF_MULTIFOU.Checked then GRF_MULTIFOU.Visible := false;

  chkFamSsFam.checked :=  GetParamSocSecur('SO_MODFAMSSFAM',true);

end;

procedure TFRecupTarifFour.RempliCombo;
var Index, IndexItem, iTable, iChamp : integer;
    LibelleChamp : string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  if TobParFou.getValue('GRF_TYPEENREG')<>'X' then
  begin
	  TobParFouLig.Detail.Sort('GFL_OFFSET');
  end;
  for Index := 0 to TobParFouLig.Detail.Count - 1 do
  begin
    if TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP') <> 'PASS' then
    begin
      if pos ('GCA_', TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP')) > 0 then
      begin
        Table := Mcd.getTable ('CATALOGU');
      end else
      begin
        Table := Mcd.getTable ('ARTICLE');
      end;
      FieldList := Table.Fields;
      FieldList.Reset();
      While FieldList.MoveNext do
      begin
        if (FieldList.Current as IFieldCOM).name = TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP') then
        LibelleChamp :=(FieldList.Current as IFieldCOM).Libelle;
      end;
      IndexItem := HChamp1.Items.Add (LibelleChamp);
      HChamp1.Items.objects[IndexItem] := TobParFouLig.Detail [Index];
      IndexItem := HChamp2.Items.Add (LibelleChamp);
      HChamp2.Items.Objects[IndexItem] := TobParFouLig.Detail [Index];
      IndexItem := HChamp3.Items.Add (LibelleChamp);
      HChamp3.Items.objects[IndexItem] := TobParFouLig.Detail [Index];
    end;
  end;
  HChamp1.Items.add ('Autre');
  HChamp2.Items.add ('Autre');
  HChamp3.Items.add ('Autre');

  HChamp1.Items.add ('Aucun');
  HChamp2.Items.add ('Aucun');
  HChamp3.Items.add ('Aucun');
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFRecupTarifFour.bFinClick(Sender: TObject);
begin
  inherited;
  ListeRecap;

  if ControleSaisie then
  begin
    if fileexists (GRF_FICHIER.Text) then
    begin
// --
      TOBErreurs        := TOB.Create ('LES ANOMALIES',nil,-1);
      TOBChampsArt      := TOB.Create('LES CHAMPS ART',nil,-1);
      TOBChampsCatalogu := TOB.Create ('LES CHAMPS CAT',nil,-1);
      TOBChampsArtC     := TOB.Create ('LES CHAMPS ARTICLES COMPL',nil,-1);
      TOBChampsCB       := TOB.Create ('LES CHAMPS CAB',nil,-1);
// --
      TRY
				//
  			if not GRF_MULTIFOU.Checked then
        begin
    			TheFournisseur := GRF_TIERS.Text;
        end;
        //
        RecupereTarif;
        if TOBErreurs.detail.count > 0 then AffichageErreurs;
        if BStop = False then
        begin
          Msg.Execute(9, Caption, '');
          Close;
        end;
      FINALLY
        if TOBP <> nil then TOBP.free;
        TOBChampsArt.free;
        TOBChampsArtC.free;
        TOBChampsCatalogu.free;
        TOBErreurs.Free;
        TOBChampsCB.Free;
      END;
    end
    else
      Msg.Execute (7, Caption, '');
  end
  else
    ModalResult := 0;

end;

procedure TFRecupTarifFour.bSuivantClick(Sender: TObject);
var i_NumEcran : integer;
begin
  i_NumEcran := strtoint(Copy(P.ActivePage.Name, length(P.ActivePage.Name), 1));
  case i_NumEcran of
  	2 : if not ControleSaisie then Exit;
  end;
  inherited;
  if bFin.Enabled then ListeRecap;
end;

procedure TFRecupTarifFour.DATESUPDblClick(Sender: TObject);
begin
  inherited;
	GetDateRecherche (TForm(Self), DATESUP) ;
end;

procedure TFRecupTarifFour.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  DepileTOBLigne;
  TobParFouLig.Free;
  TobParFou.Free;
  TOBTablettes.free;
end;

procedure TFRecupTarifFour.GRF_FICHIERElipsisClick(Sender: TObject);
begin
  inherited;
  if OpenDialogButton.Execute then
  begin
    if OpenDialogButton.FileName <> '' then
    begin
    	GRF_FICHIER.Text := OpenDialogButton.Filename;
    end;
  end;
end;

procedure TFRecupTarifFour.FormCreate(Sender: TObject);
begin
  inherited;
  TobParFou := TOB.Create ('PARFOU', Nil, -1);
  TobParFouLig := TOB.Create ('', Nil, -1);
  TOBTablettes := TOB.create ('LES TABLETTES',nil,-1);
  TheFournisseur := '';
end;

procedure TFRecupTarifFour.FormShow(Sender: TObject);
begin
  inherited;
	InitialiseFiche;
end;

procedure TFRecupTarifFour.HChamp1Change(Sender: TObject);
begin
  inherited;
  if HChamp1.Text = 'Aucun' then
  begin
    DesactiveChamp (HBorneInf1);
    DesactiveChamp (HBorneSup1);
    DesactiveChamp (HLongueur1);
    DesactiveChamp (HOffset1);
    HChamp2.Enabled := False;
    HChamp2.Text := 'Aucun';
    HChamp2Change (HChamp1);
  end else
  begin
    HChamp2.Enabled := True;
    if HChamp1.Text = 'Autre' then
    begin
      if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
      begin
      	ActiveChamp (HLongueur1);
      end;
      ActiveChamp (HOffset1);
      ActiveChamp (HBorneInf1);
      ActiveChamp (HBorneSup1);
    end else
    begin
      ActiveChamp (HBorneInf1);
      ActiveChamp (HBorneSup1);
      DesactiveChamp (HLongueur1);
      DesactiveChamp (HOffset1);
    end;
  end;
end;

procedure TFRecupTarifFour.HChamp2Change(Sender: TObject);
begin
  inherited;
  if HChamp2.Text = 'Aucun' then
  begin
    DesactiveChamp (HBorneInf2);
    DesactiveChamp (HBorneSup2);
    DesactiveChamp (HLongueur2);
    DesactiveChamp (HOffset2);
    HChamp3.Enabled := False;
    HChamp3.Text := 'Aucun';
    HChamp3Change (HChamp2);
  end else
  begin
    HChamp3.Enabled := True;
    if HChamp2.Text = 'Autre' then
    begin
      if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
      begin
      	ActiveChamp (HLongueur2);
      end;
      ActiveChamp (HOffset2);
      ActiveChamp (HBorneInf2);
      ActiveChamp (HBorneSup2);
    end else
    begin
      ActiveChamp (HBorneInf2);
      ActiveChamp (HBorneSup2);
      DesactiveChamp (HLongueur2);
      DesactiveChamp (HOffset2);
    end;
  end;
end;

procedure TFRecupTarifFour.HChamp3Change(Sender: TObject);
begin
  inherited;
  if HChamp3.Text = 'Aucun' then
  begin
    DesactiveChamp (HBorneInf3);
    DesactiveChamp (HBorneSup3);
    DesactiveChamp (HLongueur3);
    DesactiveChamp (HOffset3);
  end else
  begin
    if HChamp3.Text = 'Autre' then
    begin
      if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
      begin
        ActiveChamp (HLongueur3);
      end;
      ActiveChamp (HOffset3);
      ActiveChamp (HBorneInf3);
      ActiveChamp (HBorneSup3);
    end else
    begin
      ActiveChamp (HBorneInf3);
      ActiveChamp (HBorneSup3);
      DesactiveChamp (HLongueur3);
      DesactiveChamp (HOffset3);
    end;
  end;
end;

{==============================================================================================}
{=============================== Controles ====================================================}
{==============================================================================================}

function TFRecupTarifFour.ControleChampOk (HChamp : THValcomboBox; ValChamp : string;
                                           stLongueur, stOffset, stBorneInf, stBorneSup : string) : boolean;
var IndexItem : integer;
    TobPar : TOB;
begin
  Result := True;
  if HChamp.Text <> 'Autre' then
  begin
    IndexItem := HChamp.Items.IndexOf (HChamp.Text);
    TobPar := TOB(HChamp.Items.Objects[IndexItem]);
    RecupereValChamp (TobPar, TobPar.GetValue ('GFL_OFFSET'), ValChamp);
  end else
  begin
    if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
    begin
      ValChamp := Copy (stEnreg, StrToInt (stOffset), StrToInt (stLongueur));
    end else
    begin
      RecupereValChamp (Nil, StrToInt (stOffset) - 1, ValChamp);
    end;
  end;
  if (Copy(ValChamp, 0, Length(stBorneInf)) < stBorneInf) or
  	 (Copy(ValChamp, 0, Length(stBorneSup)) > stBorneSup) then Result := False;
end;

function TFRecupTarifFour.ControleEnreg : boolean;
var ValChamp : string;
begin
  Result := True;
  if HChamp1.Text <> 'Aucun' then
  begin
    if not ControleChampOk (HChamp1, ValChamp, HLongueur1.Text, HOffset1.Text,
                            HBorneInf1.Text, HBorneSup1.Text) then
    begin
      Result := False
    end else if HChamp2.Text <> 'Aucun' then
    begin
      if not ControleChampOk (HChamp2, ValChamp, HLongueur2.Text, HOffset2.Text,
      												HBorneInf2.Text, HBorneSup2.Text) then
      begin
      	Result := False
      end else if HChamp3.Text <> 'Aucun' then
      begin
        if not ControleChampOk (HChamp3, ValChamp, HLongueur3.Text, HOffset3.Text,
        HBorneInf3.Text, HBorneSup3.Text) then Result := False;
      end;
    end;
  end;
end;

function TFRecupTarifFour.ControleSaisie : boolean;
begin
  Result := True;
  if not ControleSelection (HChamp1.Text, HBorneSup1.Text, HBorneInf1.Text, HLongueur1.Text,
  													HOffset1.Text, 0, 7) then
  begin
  	Result := False
  end else if not ControleSelection (HChamp2.Text, HBorneSup2.Text, HBorneInf2.Text, HLongueur2.Text,
  																	 HOffset2.Text, 2, 3) then
  begin
  	Result := False
  end else if not ControleSelection (HChamp3.Text, HBorneSup3.Text, HBorneInf3.Text,
  																	HLongueur3.Text, HOffset3.Text, 4, 5) then
  begin
  	Result := False;
  end;
end;

function TFRecupTarifFour.ControleSelection (stChamp, stBorneSup, stBorneInf, stLongueur,
                                             stOffset : string; Message1, Message2 : integer) : boolean;
begin
  Result := True;
  if stChamp <> 'Aucun' then
  begin
    if stBorneSup < stBorneInf then
    begin
      Msg.Execute (Message1 + 1, Caption, '');
      Result := False;
    end else
    begin
      if stChamp = 'Autre' then
      begin
        if (((stLongueur = '') or (Not IsNumeric (stLongueur))) and
            (TobParFou.GetValue ('GRF_TYPEENREG') = 'X')) or
            ((stOffset = '') or (Not IsNumeric (stOffset))) then
        begin
          Msg.Execute (Message2 + 1, Caption, '');
          Result := False;
        end;
      end;
    end;
  end;
end;

{==============================================================================================}
{=============================== Traitements ==================================================}
{==============================================================================================}
procedure TFRecupTarifFour.GetInfoProfil (TobArt : TOB;Profil : string);
var i : integer;
begin
  if TOBP = nil then
  begin
    TOBP := TOB.create ('PROFILART',nil,-1);
    TOBP.putvalue('GPF_PROFILARTICLE',Profil);
    TOBP.LoadDB (true);
  end;
  if TobArt.GetValue('GA_FAMILLENIV1') = '' then
  begin
  	TobArt.putValue('GA_FAMILLENIV1', TOBP.getValue('GPF_FAMILLENIV1'));
  end;
  if TobArt.GetValue('GA_FAMILLENIV2') = '' then
  begin
  	TobArt.putValue('GA_FAMILLENIV2', TOBP.getValue('GPF_FAMILLENIV2'));
  end;
  if TobArt.GetValue('GA_FAMILLENIV3') = '' then
  begin
  	TobArt.putValue('GA_FAMILLENIV3', TOBP.getValue('GPF_FAMILLENIV3'));
  end;
  TobArt.putValue('GA_COMPTAARTICLE', TOBP.getValue('GPF_COMPTAARTICLE'));
  TobArt.putValue('GA_TENUESTOCK',    TOBP.getValue('GPF_TENUESTOCK'));
  TobArt.putValue('GA_CALCPRIXPR',    TOBP.getValue('GPF_CALCPRIXPR'));
  TobArt.putValue('GA_COEFFG',        TOBP.getValue('GPF_COEFCALCPR'));
  TobArt.putValue('GA_DPRAUTO',       TOBP.getValue('GPF_CALCAUTOPR'));

{$IFDEF CCS3}
  TobArt.putValue('GA_LOT','-');
  TobArt.putValue('GA_NUMEROSERIE','-');
{$ELSE}
  TobArt.putValue('GA_LOT',           TOBP.getValue('GPF_LOT'));
  TobArt.putValue('GA_NUMEROSERIE',   TOBP.getValue('GPF_NUMEROSERIE'));
{$ENDIF}
  TobArt.putValue('GA_CONTREMARQUE',  TOBP.getValue('GPF_CONTREMARQUE'));
  TobArt.putValue('GA_REMISEPIED',    TOBP.getValue('GPF_REMISEPIED'));
  TobArt.putValue('GA_REMISEPIED',    TOBP.getValue('GPF_REMISEPIED'));
  TobArt.putValue('GA_REMISELIGNE',   TOBP.getValue('GPF_REMISELIGNE'));
  TobArt.putValue('GA_ESCOMPTABLE',   TOBP.getValue('GPF_ESCOMPTABLE'));
  TobArt.putValue('GA_FAMILLETAXE1',  TOBP.getValue('GPF_CODETAXE'));
  TobArt.putValue('GA_COMMISSIONNABLE',TOBP.getValue('GPF_COMMISSIONNABL'));
  TobArt.putValue('GA_CALCPRIXHT',    TOBP.getValue('GPF_CALCPRIXHT'));
  TobArt.putValue('GA_CALCPRIXTTC',   TOBP.getValue('GPF_CALCPRIXTTC'));
  TobArt.putValue('GA_COEFCALCHT',    TOBP.getValue('GPF_COEFCALCHT'));
  TobArt.putValue('GA_COEFCALCTTC',   TOBP.getValue('GPF_COEFCALCTTC'));
  TobArt.putValue('GA_CALCAUTOHT',    TOBP.getValue('GPF_CALCAUTOHT'));
  TobArt.putValue('GA_CALCAUTOTTC',   TOBP.getValue('GPF_CALCAUTOTTC'));
  TobArt.putValue('GA_TARIFARTICLE',  TOBP.getValue('GPF_TARIFARTICLE'));
  TobArt.putValue('GA2_SOUSFAMTARART',TOBP.getValue('GPF_SOUSFAMTARART'));

  TobArt.putValue('GA_PAYSORIGINE',   TOBP.getValue('GPF_PAYSORIGINE'));

  if (CtxMode in V_PGI.PGIContexte) then
  begin
    TobArt.putValue('GA_COEFFG',      TOBP.GetValue('GPF_COEFCALCPR'));
    TobArt.putValue('GA_DPRAUTO',     TOBP.GetValue('GPF_CALCAUTOPR'));
    TobArt.putValue('GA_ARRONDIPRIX', TOBP.GetValue('GPF_ARRONDIPRIX'));
    TobArt.putValue('GA_ARRONDIPRIXTTC',TOBP.GetValue('GPF_ARRONDIPRIXTTC'));
    TobArt.putValue('GA_PRIXUNIQUE',  TOBP.GetValue('GPF_PRIXUNIQUE'));
    //JD
    TobArt.putValue('GA_COLLECTION',  TOBP.GetValue('GPF_COLLECTION'));
    TobArt.putValue('GA_FOURNPRINC',  TOBP.GetValue('GPF_FOURNPRINC'));
    //
    For i:=1 to $A  do TobArt.putValue('GA_LIBREART'+format('%x',[i]),  TOBP.GetValue('GPF_LIBREART' +format('%x',[i])));
    For i:=1 to 3   do TobArt.putValue('GA_VALLIBRE'+IntToStr(i),       TOBP.GetValue('GPF_VALLIBRE' +IntToStr(i)));
    For i:=1 to 3   do TobArt.putValue('GA_DATELIBRE'+IntToStr(i),      TOBP.GetValue('GPF_DATELIBRE'+IntToStr(i)));
    For i:=1 to 3   do TobArt.putValue('GA_CHARLIBRE'+IntToStr(i),      TOBP.GetValue('GPF_CHARLIBRE'+IntToStr(i)));
    For i:=1 to 3   do TobArt.putValue('GA_BOOLLIBRE'+IntToStr(i),      TOBP.GetValue('GPF_BOOLLIBRE'+IntToStr(i)));
  end;
  
end;


procedure TFRecupTarifFour.CompleteChampArticle (Index : integer);
begin
  TobArticle.Detail[Index].PutValue ('GA_CODEARTICLE', CodeArticle);
  TobArticle.Detail[Index].PutValue ('GA_ARTICLE', CodeArticleUnique2 (CodeArticle, ''));
  TobArticle.Detail[Index].PutValue ('GA_TYPEARTICLE', TobParFou.GetValue ('GRF_TYPEARTICLE'));
  TobArticle.Detail[Index].PutValue ('GA_LIBELLE', CodeArticle);
  TobArticle.Detail[Index].PutValue ('GA_COMPTAARTICLE', TobParFou.GetValue ('GRF_COMPTAARTICLE'));
  TobArticle.Detail[Index].PutValue ('GA_FERME', '-');
  TobArticle.Detail[Index].PutValue ('GA_DATECREATION', V_PGI.DateEntree);
  TobArticle.Detail[Index].PutValue ('GA_DATEMODIF', V_PGI.DateEntree);
  TobArticle.Detail[Index].PutValue ('GA_DATESUPPRESSION', StrToDate (DATESUP.Text));
  TobArticle.Detail[Index].PutValue ('GA_TENUESTOCK', TobParFou.GetValue ('GRF_TENUESTOCK'));
  TobArticle.Detail[Index].PutValue ('GA_FAMILLETAXE1', TobParFou.GetValue ('GRF_FAMILLETAXE1'));
  TobArticle.Detail[Index].PutValue ('GA_UTILISATEUR', V_PGI.User);
  TobArticle.Detail[Index].PutValue ('GA_SOCIETE', V_PGI.CodeSociete);
  TobArticle.Detail[Index].PutValue ('GA_CREATEUR', V_PGI.User);
  TobArticle.Detail[Index].PutValue ('GA_STATUTART', 'UNI');
  TobArticle.Detail[Index].PutValue ('GA_REMISEPIED', TobParFou.GetValue ('GRF_REMISEPIED'));
  TobArticle.Detail[Index].PutValue ('GA_ESCOMPTABLE', TobParFou.GetValue ('GRF_ESCOMPTABLE'));
  // Modif BTP
  if TOBParFou.GetValue('GRF_PROFILARTICLE') <> '' then
  begin
    GetInfoProfil (TobArticle.Detail[Index],TOBParFou.GetValue('GRF_PROFILARTICLE'));
  end;
  if TobParFou.GetValue('GRF_FOURPRINC')='X' then
  begin
//    TOBArticle.detail[Index].putValue('GA_FOURNPRINC',TobParFou.GetValue ('GRF_TIERS'));
    TOBArticle.detail[Index].putValue('GA_FOURNPRINC',TheFournisseur);
  end;
  // --
  if TobArticle.Detail[Index].GetValue ('GA_TYPEARTICLE')='ARP' Then
  begin
    SetDetailPrixPose (Index);
  end;
  //
  TobArticle.detail[Index].PutValue ('BCB_NATURECAB','MAR');
  TobArticle.detail[Index].PutValue ('BCB_IDENTIFCAB', TobArticle.Detail[Index].GetValue ('GA_ARTICLE'));
  TobArticle.detail[Index].PutValue ('BCB_CABPRINCIPAL', 'X');

end;

procedure TFRecupTarifFour.CompleteChampCatalogu (Index : integer; Code : string);
begin
  TobCatalogu.Detail[Index].PutValue ('GCA_REFERENCE', UpperCase (Code));
//  TobCatalogu.Detail[Index].PutValue ('GCA_TIERS', TobParFou.GetValue ('GRF_TIERS'));
  TobCatalogu.Detail[Index].PutValue ('GCA_TIERS', TheFournisseur);
  TobCatalogu.Detail[Index].PutValue ('GCA_LIBELLE', CodeArticle);
  TobCatalogu.Detail[Index].PutValue ('GCA_PRIXPOURQTEAC', 1);
  TobCatalogu.Detail[Index].PutValue ('GCA_DATEREFERENCE', V_PGI.DateEntree);
  TobCatalogu.Detail[Index].PutValue ('GCA_DATESUP', StrToDate (DATESUP.Text));
  if CBCreationArticle.Checked = True then
  	TobCatalogu.Detail[Index].PutValue ('GCA_ARTICLE', CodeArticleUnique2 (CodeArticle, ''));
end;

procedure TFRecupTarifFour.CreationElementTob (stTable : string; IndexPar : integer;
                                               TobTable : Tob; CBCreation : boolean;
                                               var Index : integer);
var //TOBL : TOB;
    ValChamp, Champ, stRequete : string;
    TSql : TQuery;
begin

  RecupereValChamp (TobParFouLig.Detail[IndexPar], IndexPar, Valchamp);

{$IFDEF BTP}
  FindEtReplace (ValChamp,'"','''',true);
{$ENDIF}

  Champ := TobParFouLig.Detail[IndexPar].GetValue ('GFL_CHAMP');

  if (stTable = 'ARTICLE') and (Champ = 'GCA_REFERENCE') then
  begin
    Champ := 'GA_CODEARTICLE';
  {$IFDEF BTP}
    CodeArticle := TobParFou.GetValue ('GRF_PREFIXE') + UpperCase (ValChamp) + TobParFou.GetValue ('GRF_SUFFIXE');
    CodeArticle := Copy(CodeArticle, 1, 18); // Modif brl 180308 : on limite à 18c
  {$ELSE}
    CodeArticle := UpperCase(ValChamp);
  {$ENDIF}
  end
  else if stTable = 'ARTICLE' then
  begin
    CodeArticle := TobParFou.GetValue ('GRF_PREFIXE') + UpperCase (ValChamp) + TobParFou.GetValue ('GRF_SUFFIXE');
  end;

  Index := -1;

  if ValChamp <> '' then
  begin
    if stTable = 'ARTICLE' then
    begin
      stRequete := 'SELECT '+ListChampsArt;
      if ListChampsArtC <> '' then stRequete := stRequete + ',' +ListChampsArtC;
      stRequete := stRequete + ' FROM ' + stTable + ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE ';
      StRequete := StRequete + ' LEFT JOIN BTCODEBARRE ON BCB_IDENTIFCAB=GA_ARTICLE ';
      StRequete := StRequete + 'WHERE ' + Champ + '="' + CodeArticle + '"';
    end
    else
    begin
      stRequete := 'SELECT ' + ListChampsCat + ' FROM ' + stTable + ' WHERE ' + Champ + '="' + ValChamp + '"' + ' AND GCA_TIERS="' + TheFournisseur + '"';
    end;

    if (CBCreation = True) then
    begin
      Index := TobTable.Detail.Count;
      TOB.create (stTable, TobTable, Index);
      TobTable.Detail[Index].initValeurs;
      TOBTable.detail[Index].addchampsupValeur ('GA2_SOUSFAMTARART','');

      TOBTable.detail[Index].addchampsupValeur ('BCB_NATURECAB','');
      TOBTable.detail[Index].addchampsupValeur ('BCB_IDENTIFCAB','');
      TOBTable.detail[Index].addchampsupValeur ('BCB_CABPRINCIPAL','');
      TOBTable.detail[Index].addchampsupValeur ('BCB_QUALIFCODEBARRE','');
      TOBTable.detail[Index].addchampsupValeur ('BCB_CODEBARRE','');

      // AJOUT LS Pour Optimisation
      TobTable.Detail[Index].addchampsupValeur('UPDATE','-');
      TobTable.Detail[Index].addchampsupValeur ('ISCHAMPSMODIFIE','-');
      // --
      if not CBRecInit.checked then
      begin
        Tsql := OpenSql (stRequete, True,-1,'',true);
        if not Tsql.eof then
        begin
          TobTable.Detail[Index].SelectDB ('', TSql);
          TOBTable.Detail[Index].addchampsupValeur('UPDATE','X');
        end
        else
        begin
          if stTable = 'ARTICLE' then CompleteChampArticle (Index)
                                 else CompleteChampCatalogu (Index, ValChamp);
        end;
        Ferme (TSql);
      end
      else
      begin
        if stTable = 'ARTICLE' then CompleteChampArticle (Index)
                               else CompleteChampCatalogu (Index, ValChamp);
      end;
    end;

    if (stTable = 'CATALOGU') and (Index >= 0) then
    Begin
      // Modif BRL 11/03/2010 : Mise à jour systématique de la date de suppression catalogue si celle-ci change
      if (TobCatalogu.Detail[Index].GetValue ('GCA_DATESUP') <> StrToDate (DATESUP.Text)) then
      Begin
        TobCatalogu.Detail[Index].PutValue ('GCA_DATESUP', StrToDate (DATESUP.Text));
        TobCatalogu.Detail[Index].Putvalue ('ISCHAMPSMODIFIE','X');
      End;
    End;

  // Modif BRL 05/03/2010 : Modification du fournisseur principal de la fiche article
    if stTable = 'ARTICLE' then
    begin
      if (Index >= 0) and (TobParFou.GetValue('GRF_FOURPRINC')='X') then
      begin
        //TOBArticle.detail[Index].putValue('GA_FOURNPRINC',TobParFou.GetValue ('GRF_TIERS'));
        TOBArticle.detail[Index].putValue('GA_FOURNPRINC',TheFournisseur);
        TOBArticle.Detail[Index].Putvalue ('ISCHAMPSMODIFIE','X');
      end;
      // Modif BRL 03/07/2012 : Mise à jour systématique de la date de suppression article si celle-ci change
      if (Index >= 0) and (TOBArticle.Detail[Index].GetValue ('GA_DATESUPPRESSION') <> StrToDate (DATESUP.Text)) then
      Begin
        TOBArticle.Detail[Index].PutValue ('GA_DATESUPPRESSION', StrToDate (DATESUP.Text));
        TOBArticle.Detail[Index].Putvalue ('ISCHAMPSMODIFIE','X');
      end;
    end;
  end;

end;

procedure TFRecupTarifFour.MAJTablette (TheTablette : TOB);
var Indice : integer;
    TheTabll : TOB;
begin

  if (TheTablette.GetValue('FICHIER LIE') <> '') and (TheTablette.detail.count > 0) then
  begin
    if FileExists(TheTablette.GetValue('FICHIER LIE')) then MajLibelletablette(TheTablette);
  end;

  for Indice := 0 to TheTablette.detail.count -1 do
  begin
    TheTabll := TheTablette.detail[Indice];
    TheTabll.SetAllModifie (true);
    if TheTabll.GetValue('EXIST')='-' then
    begin
      TheTabll.InsertDB (nil);
      TheTabll.PutValue('EXIST','X');
    end else
      TheTabll.UpdateDB;
  end;

  AvertirTable (TheTablette.getValue('DO_COMBO'));

end;

procedure TFRecupTarifFour.VerifSousFamilleTarif;
var indice : integer;
    TOBA : TOB;
begin
{$IFDEF BTP}
  for indice := 0 to TobArticle.detail.count -1 do
  begin
    TOBA := TobArticle.detail[Indice];
    if TOBA.GetValue ('GA_TARIFARTICLE')= '' then
    begin
      TOBA.putValue('GA2_SOUSFAMTARART','');
    end;
  end;
{$ENDIF}
end;

procedure TFRecupTarifFour.MAjPRixArticle;
var indice : integer;
    TOBA : TOB;
begin
{$IFDEF BTP}
  for indice := 0 to TobArticle.detail.count -1 do
  begin
    TOBA := TobArticle.detail[Indice];
    if TOBA.GetValue ('UPDATE')= 'X' then
    begin
      // si pas de modif --> on ne met pas a jour la table car inutile
      if TOBA.GetValue('ISCHAMPSMODIFIE') = 'X' then
      begin
        DefinieValorisation (TOBA);
      end;
    end else
    begin
      DefinieValorisation (TOBA);
    end;
  end;
{$ENDIF}
end;

procedure TFRecupTarifFour.MiseajourTables;
var Indice        : integer;
    TheTablette   : TOB;
    TOBA,TOBARTC  : TOB;
begin

  if CBCreationCatalogue.checked then
  begin
  	ElimineDoublonsCatalog (TobCatalogu);
    for indice := 0 to TobCatalogu.detail.count -1 do
    begin
      TOBA := TobCatalogu.detail[Indice];
//      TobCatalogu.InsertOrUpdateDB (true);
      if TOBA.GetValue ('UPDATE')= 'X' then
      begin
        if TOBA.GetValue('ISCHAMPSMODIFIE') = 'X' then TOBA.UpdateDB (false);
      end else
      begin
        TobA.SetAllModifie (True);
        TOBA.InsertDB (nil,false);
      end;
    end;
  end;

  if CBCreationArticle.checked then
  begin
    ElimineDoublonsArticles(TobArticle);
//    TobArticle.InsertOrUpdateDB(true);
    for indice := 0 to TobArticle.detail.count -1 do
    begin
      TOBA := TobArticle.detail[Indice];
      if TOBA.GetValue ('UPDATE')= 'X' then
      begin
        // si pas de modif --> on ne met pas a jour la table car inutile
        if TOBA.GetValue('ISCHAMPSMODIFIE') = 'X' then
        begin
          TOBA.UpdateDB (false);
        end;
      end else
      begin
        TobA.SetAllModifie (True);
        TOBA.InsertDB (nil,false);
      end;
      TOBARTC := TOB.create('ARTICLECOMPL',TOBArticleCompl,-1);
      TOBARTC.putValue('GA2_ARTICLE',       TOBA.GetValue('GA_ARTICLE'));
      TOBARTC.putValue('GA2_CODEARTICLE',   TOBA.GetValue('GA_CODEARTICLE'));
      TOBARTC.putValue('GA2_SOUSFAMTARART', TOBA.GetValue('GA2_SOUSFAMTARART'));

      //Mise à jour des Tarifs Article dans BTFAMILLETARIF et BTSOUSFAMILLETARIF
      if TOBA.GetValue('GA_TARIFARTICLE')<> '' then MiseAJourFamilleTarif(TOBA);

      //
      //Gestion de la table des codes barres uniquement si la zone code barre est présente dans la TOB Article.
      IF TOBA.FieldExists('BCB_CODEBARRE') then
      begin
        if TOBA.GetValue('BCB_CODEBARRE') <> '' then MiseAJourCodeBarre(TOBA);
      End;

    end;
    if TOBArticleCompl.detail.count > 0 then TOBArticleCompl.InsertorUpdateDB (true);
    if TOBEnt.detail.count > 0          then TobEnt.InsertorUpdateDB (true);
    if TOBLig.detail.count > 0          then TobLig.InsertorUpdateDB (true);
    if TOBCB.detail.count > 0           then TOBCB.InsertorUpdateDB (true);

    for Indice := 0 to TOBTablettes.detail.count -1 do
    begin
      TheTablette := TOBTablettes.detail[Indice];
      MAJTablette (TheTablette);
    end;
  end;


end;

//FV1 : 21/01/2016 - FS#1854 - En import catalogue, les tables BTFAMILLETARIF et BTSOUSFAMILLETARIF ne sont pas alimentées
Procedure TFRecupTarifFour.MiseAJourFamilleTarif(TOBA : TOB);
Var CodeTarif : string;
    TobTarif  : TOB;
    Req       : string;
begin

  if not chkFamSsFam.Checked then Exit;

  CodeTarif := TOBA.GetString('GA_TARIFARTICLE');

  if CodeTarif = '' then Exit;

  // Contrôle existence nouveau code tarif avant traitement dans tablette tarif article (GCTARIFARTICLE)
  Req := 'SELECT BFT_FAMILLETARIF FROM BTFAMILLETARIF WHERE BFT_FAMILLETARIF="'+ CodeTarif +'"';
  if Not ExisteSql(Req) then
  begin
    TobTarif := Tob.create('BTFAMILLETARIF', nil, -1);
    TobTarif.PutValue('BFT_FAMILLETARIF', CodeTarif);
    TobTarif.PutValue('BFT_LIBELLE', 'Famille : ' + CodeTarif);
    TOBtarif.SetAllModifie(true); 
    TobTarif.InsertOrUpdateDB;
    AvertirTable('GCTARIFARTICLE');
    FreeAndNil(TobTarif);
  end ;

  if TobA.GetString('GA2_SOUSFAMTARART') <> '' then MiseAJourSsFamilleTarif(TOBA);

end;

  //FV1 : 21/01/2016 - FS#1854 - En import catalogue, les tables BTFAMILLETARIF et BTSOUSFAMILLETARIF ne sont pas alimentées
Procedure TFRecupTarifFour.MiseAJourSsFamilleTarif(TOBA : TOB);
Var CodeSTarif: string;
    CodeTarif : string;
    TobSTarif : TOB;
    Req       : string;
begin

  if not chkFamSsFam.Checked then Exit;

  CodeTarif  := TOBA.GetString('GA_TARIFARTICLE');
  CodeSTarif := TOBA.GetString('GA2_SOUSFAMTARART');

  // Contrôle existence nouveau code tarif avant traitement dans tablette tarif article (GCTARIFARTICLE)
  Req := 'SELECT * FROM BTSOUSFAMILLETARIF WHERE BSF_FAMILLETARIF="'+ CodeTarif +'" AND BSF_SOUSFAMTARART="' + CodeStarif + '"';
  if Not ExisteSql(Req) then
  begin
    TobSTarif := Tob.create('BTSOUSFAMILLETARIF', nil, -1);
    TobSTarif.PutValue('BSF_FAMILLETARIF',      CodeTarif);
    TobSTarif.PutValue('BSF_SOUSFAMTARART',  CodeSTarif);
    TobSTarif.PutValue('BSF_LIBELLE', 'Sous-Famille : ' + CodeSTarif);
    TobSTarif.SetAllModifie(true); 
    TobSTarif.InsertOrUpdateDB;
    AvertirTable ('BTSOUSFAMTARART'); // rechargement tablette
    FreeAndNil(TobSTarif);
  end ;

end;

Procedure TFRecupTarifFour.MiseAJourCodeBarre(TOBA : TOB);
Var StSQL : String;
    QQ    : TQuery;
    TOBLCB: TOB;
    okCrea: boolean;
begin

  okcrea := False;

  StSQL := 'SELECT * FROM BTCODEBARRE WHERE BCB_NATURECAB="MAR" AND BCB_IDENTIFCAB="' + TOBA.GetValue('GA_ARTICLE') + '"';
  QQ := OpenSQL(StSQL, False);

  if QQ.EOF Then
  begin
    OkCrea  := True;
    TOBLCB  := TOB.create('BTCODEBARRE',TOBCB,-1);
    TOBLCB.putValue('BCB_NATURECAB',      'MAR');
    TOBLCB.putValue('BCB_CABPRINCIPAL',   'X');
    TOBLCB.putValue('BCB_IDENTIFCAB',   TOBA.GetValue('GA_ARTICLE'));
  end
  else
  begin
    TOBLCB  := TOB.create('BTCODEBARRE',TOBCB,-1);
    TOBLCB.SelectDB('BTCODEBARRE', QQ);
  end;

  Ferme(QQ);

  if TOBA.FieldExists('BCB_QUALIFCODEBARRE') then
  begin
    if TOBA.GetValue('BCB_QUALIFCODEBARRE') = '' then
      TOBLCB.putValue('BCB_QUALIFCODEBARRE', '128')
    else
    begin
      //contrôle si Qualifiant du code barre dans le combo yyybartype.
      StSQL := 'SELECT * FROM COMMUN WHERE CO_TYPE="YCB" AND CO_CODE="' + TOBA.GetValue('BCB_QUALIFCODEBARRE') + '"';
      if ExisteSQL(StSQL) then
        TOBLCB.putValue('BCB_QUALIFCODEBARRE', TOBA.GetValue('BCB_QUALIFCODEBARRE'))
      else
        TOBLCB.putValue('BCB_QUALIFCODEBARRE', '128')
    end;
  end
  else
    TOBLCB.putValue('BCB_QUALIFCODEBARRE','128');

  //
  if TOBA.FieldExists('BCB_CODEBARRE')       then
  begin
    if TOBLCB.GetValue('BCB_CODEBARRE') <> TOBA.GetValue('BCB_CODEBARRE') then
    begin
     if not okcrea then TOBLCB.putValue('BCB_CABPRINCIPAL',   '-');
     TOBLCB.putValue('BCB_CODEBARRE',     TOBA.GetValue('BCB_CODEBARRE'));
    end;
  end;


end;

procedure TFRecupTarifFour.ListeRecap;
BEGIN
  ListRecap.Items.Clear;
  ListRecap.Items.Add (PTITRE.Caption);
  ListRecap.Items.Add ('');
  ListRecap.Items.Add ('Fichier récupéré : ' + GRF_FICHIER.Text);
  ListRecap.Items.Add ('');
  ListRecap.Items.Add (TDateSup.Caption + ' : ' + DATESUP.Text);
  ListRecap.Items.Add ('');
  if CBCreationCatalogue.checked then
  BEGIN
  	ListRecap.Items.Add (ExtractLibelle (CBCreationCatalogue.Caption) + HRecap.Mess[0]);
  end else
  BEGIN
  	ListRecap.Items.Add (ExtractLibelle (CBCreationCatalogue.Caption) + HRecap.Mess[1]);
  end;
  ListRecap.Items.Add ('');
  if CBCreationArticle.checked then
  BEGIN
	  ListRecap.Items.Add (ExtractLibelle (CBCreationArticle.Caption) + HRecap.Mess[0]);
  end else
  BEGIN
	  ListRecap.Items.Add (ExtractLibelle (CBCreationArticle.Caption) + HRecap.Mess[1]);
  end;
  ListRecap.Items.Add ('');
  if HChamp1.Text <> 'Aucun' then
  begin
    ListRecap.Items.Add ('Sélection des enregistrements');
    RecapChamp (HChamp1.Text, HLongueur1.Text, HOffset1.Text, HBorneInf1.Text,
    HBorneSup1.Text);
    if HChamp2.Text <> 'Aucun' then
    begin
      ListRecap.Items.Add ('');
      RecapChamp (HChamp2.Text, HLongueur2.Text, HOffset2.Text, HBorneInf2.Text,
      HBorneSup2.Text);
      if HChamp3.Text <> 'Aucun' then
      begin
        ListRecap.Items.Add ('');
        RecapChamp (HChamp3.Text, HLongueur3.Text, HOffset3.Text, HBorneInf3.Text,
        HBorneSup3.Text);
      end;
    end;
  end else
  begin
  	ListRecap.Items.Add ('Aucun élément de sélection des enregistrements');
  end;
  ListRecap.Items.Add ('');
end;

procedure TFRecupTarifFour.RecapChamp (stChamp, stLongueur, stOffset, stBorneInf,
                                       stBorneSup : string);
var stChaine : string;
begin
  if stChamp <> 'Autre' then
  begin
  	ListRecap.Items.Add ('Sélection sur le champ : ' + stChamp);
  end else
  begin
    if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
    begin
      ListRecap.Items.Add ('Sélection à la position : ' + stOffset);
      stChaine := '    sur une longeur de : ' + stLongueur + ' caractère';
      if StrToInt(stLongueur) > 1 then stChaine := stChaine + 's';
      ListRecap.Items.Add (stChaine);
    end else
    begin
      ListRecap.Items.Add ('Sélection sur le champ n° : ' + stOffset);
    end;
  end;
  ListRecap.Items.Add ('    Borne inférieure : ' + stBorneInf);
  ListRecap.Items.Add ('    Borne supérieure : ' + stBorneSup);
end;

function TFRecupTarifFour.EnregistreTables : TIOErr;
begin
	VerifSousFamilleTarif;
  MajPrixArticle;
  result := Transactions (MiseajourTables, 0);
  if result = OeOk then RenseigneJournalEvenement (result,true);
end;

procedure TFRecupTarifFour.VideTobs;
begin
  TobArticle.ClearDetail;
  TobArticleCompl.ClearDetail;
  TobCatalogu.ClearDetail;
  TOBEnt.ClearDetail;
  TOBLig.ClearDetail;
  TOBTablettes.ClearDetail;
  TOBCB.ClearDetail;
end;

procedure TFRecupTarifFour.RemplitChampsTob;
var indice    : integer;
    TOBPFL    : TOB;
    TOBCC     : TOB;
    TOBCA     : TOB;
    TheChamps : String;
begin

  for Indice := 0 to TOBParFouLig.detail.count -1 do
  begin
    TOBPFL := TobParFouLig.Detail[Indice];
    TheChamps := TOBPFL.GetValue ('GFL_CHAMP');
    if Copy(theChamps,1,4) = 'GCA_' then
    begin
      if TheChamps = 'GCA_REFERENCE' then
      begin
        IndexParCatalogu := Indice;
      end;
      // Champs catalogue
      if TOBChampsCatalogu.FindFirst (['CHAMPS'],[TheChamps],true) = nil then
      begin
        TOBCC := TOB.Create('The CHAMPS',TOBChampsCatalogu,-1);
        TOBCC.AddChampSupValeur ('CHAMPS',TheChamps,false);
      end;
      //
      if (TheChamps = 'GCA_PRIXBASE') or (TheChamps = 'GCA_PRIXVENTE') then
      begin

        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_PAHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_PAHT',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_COEFFG'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_COEFFG',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCPRIXPR'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCPRIXPR',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_DPR'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_DPR',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_COEFCALCHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_COEFCALCHT',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCPRIXHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCPRIXHT',false);
        end;

        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_PVHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_PVHT',false);
        end;

        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_PVTTC'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_PVTTC',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_COEFCALCTTC'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_COEFCALCTTC',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCPRIXTTC'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCPRIXTTC',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_DPRAUTO'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_DPRAUTO',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCAUTOHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCAUTOHT',false);
        end;
      end;
    end
    else if Copy(theChamps,1,3) = 'GA_' then
    begin
      // champs Article
      if TheChamps = 'GA_CODEARTICLE' then
      begin
        IndexParArticle := Indice;
      end;

      if TOBChampsArt.FindFirst (['CHAMPS'],[TheChamps],true) = nil then
      begin
        TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
        TOBCA.AddChampSupValeur ('CHAMPS',TheChamps,false);
      end;
      if (TheChamps = 'GA_PAHT') or (TheChamps = 'GA_DPR') or (TheChamps = 'GA_PVHT') then
      begin

        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_PAHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_PAHT',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_COEFFG'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_COEFFG',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCPRIXPR'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCPRIXPR',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_DPR'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_DPR',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_COEFCALCHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_COEFCALCHT',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCPRIXHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCPRIXHT',false);
        end;

        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_PVHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_PVHT',false);
        end;

        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_PVTTC'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_PVTTC',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_COEFCALCTTC'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_COEFCALCTTC',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCPRIXTTC'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCPRIXTTC',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_DPRAUTO'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_DPRAUTO',false);
        end;
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CALCAUTOHT'],true) = nil then
        begin
          TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
          TOBCA.AddChampSupValeur ('CHAMPS','GA_CALCAUTOHT',false);
        end;
      end;
    end
    else if Copy(TheChamps,1,4)= 'GA2_' then
    begin
      if TOBChampsArtC.FindFirst (['CHAMPS'],[TheChamps],true) = nil then
      begin
        TOBCA := TOB.Create('The CHAMPS',TOBChampsArtC,-1);
        TOBCA.AddChampSupValeur ('CHAMPS',TheChamps,false);
      end;
    end
    else if Copy(theChamps,1,4) = 'BCB_' then
    begin
      if TOBChampsCB.FindFirst (['CHAMPS'],[TheChamps],true) = nil then
      begin
        TOBCA := TOB.Create('The CHAMPS',TOBChampsArt,-1);
        TOBCA.AddChampSupValeur ('CHAMPS',TheChamps,false);
      end;
    end;
  end;

  // --
  // AJOUT DES CLEFS et des champs obligatoire pour les traitements
  if TOBChampsArt.detail.count > 0 then
  begin
    // CODE ARTICLE
    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_ARTICLE'],true) = nil then
    begin
       ListChampsArt := 'GA_ARTICLE';
    end;

    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_CODEARTICLE'],true) = nil then
    begin
        if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_CODEARTICLE'
                               else ListChampsArt := 'GA_CODEARTICLE';
    end;

    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_TARIFARTICLE'],true) = nil then
    begin
        if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_TARIFARTICLE'
                               else ListChampsArt := 'GA_TARIFARTICLE';
    end;

    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_TYPEARTICLE'],true) = nil then
    begin
        if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_TYPEARTICLE'
                               else ListChampsArt := 'GA_TYPEARTICLE';
    end;

    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_FOURNPRINC'],true) = nil then
    begin
        if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_FOURNPRINC'
                               else ListChampsArt := 'GA_FOURNPRINC';
    end;

    if TOBChampsCatalogu.FindFirst (['CHAMPS'],['GCA_LIBELLE'],true) <> nil then
    begin
        if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_LIBELLE'
                               else ListChampsArt := 'GA_LIBELLE';
    end;

    // FAMILLE NIVEAU 3
    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_FAMILLENIV3'],true) <> nil then
    begin
      if TOBChampsArt.FindFirst (['CHAMPS'],['GA_FAMILLENIV2'],true) = nil then
      begin
        if TOBChampsArt.FindFirst (['CHAMPS'],['GA_FAMILLENIV1'],true) = nil then
        begin
            if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_FAMILLENIV1,GA_FAMILLENIV2'
                                   else ListChampsArt := 'GA_FAMILLENIV1,GA_FAMILLENIV2';
        end else
        begin
          if ListChampsArt <> '' then ListChampsArt := ListChampsCat + ',GA_FAMILLENIV2'
                                 else ListChampsArt := 'GA_FAMILLENIV2';
        end;
      end;
    end;

    // FAMILLE NIVEAU 2
    if TOBChampsArt.FindFirst (['CHAMPS'],['GA_FAMILLENIV2'],true) <> nil then
    begin
      if TOBChampsArt.FindFirst (['CHAMPS'],['GA_FAMILLENIV1'],true) = nil then
      begin
        if ListChampsArt <> '' then ListChampsArt := ListChampsArt + ',GA_FAMILLENIV1'
                               else ListChampsArt := 'GA_FAMILLENIV1';
      end;
    end;
  end;

  for Indice := 0 to TOBChampsArt.detail.count -1 do
  begin
    if ListChampsArt <> '' then ListChampsArt := ListChampsArt+ ','+TOBChampsArt.detail[Indice].getValue('CHAMPS')
                           else ListChampsArt := TOBChampsArt.detail[Indice].getValue('CHAMPS');
  end;

  // ARTICLE COMPL
  if TOBChampsArtC.detail.count > 0 then
  begin
    if TOBChampsArtC.FindFirst (['CHAMPS'],['GA2_SOUSFAMTARART'],true) = nil then
    begin
        if ListChampsArtC <> '' then ListChampsArtC := ListChampsArtC + ',GA2_SOUSFAMTARART'
                                else ListChampsArtC := 'GA2_SOUSFAMTARART';
    end;
  end;

  for Indice := 0 to TOBChampsArtC.detail.count -1 do
  begin
    if ListChampsArtC <> '' then ListChampsArtC := ListChampsArtC+ ','+TOBChampsArtC.detail[Indice].getValue('CHAMPS')
                            else ListChampsArtC := TOBChampsArtC.detail[Indice].getValue('CHAMPS');
  end;
  // -------

  if TOBChampsCatalogu.detail.count > 0 then
  begin
    if TOBChampsCatalogu.FindFirst (['CHAMPS'],['GCA_REFERENCE'],true) = nil then
        ListChampsCat := 'GCA_REFERENCE';

    if TOBChampsCatalogu.FindFirst (['CHAMPS'],['GCA_PRIXPOURQTEAC'],true) = nil then
    begin
      if ListChampsCat <> '' then ListChampsCat := ListChampsCat + ',GCA_PRIXPOURQTEAC'
                             else ListChampsCat := 'GCA_PRIXPOURQTEAC';
    end;

    if TOBChampsCatalogu.FindFirst (['CHAMPS'],['GCA_TIERS'],true) = nil then
    begin
      if ListChampsCat <> '' then ListChampsCat := ListChampsCat + ',GCA_TIERS'
                             else ListChampsCat := 'GCA_TIERS';
    end;

    if TOBChampsCatalogu.FindFirst (['CHAMPS'],['GCA_ARTICLE'],true) = nil then
    begin
      if ListChampsCat <> '' then ListChampsCat := ListChampsCat + ',GCA_ARTICLE'
                             else ListChampsCat := 'GCA_ARTICLE';
    end;

    if TOBChampsCatalogu.FindFirst (['CHAMPS'],['GCA_DATESUP'],true) = nil then
    begin
      if ListChampsCat <> '' then ListChampsCat := ListChampsCat + ',GCA_DATESUP'
                             else ListChampsCat := 'GCA_DATESUP';
    end;
  end;

  for Indice := 0 to TOBChampsCatalogu.detail.count -1 do
  begin
    if ListChampsCat <> '' then ListChampsCat := ListChampsCat+ ','+TOBChampsCatalogu.detail[Indice].getValue('CHAMPS')
                           else ListChampsCat := TOBChampsCatalogu.detail[Indice].getValue('CHAMPS');
  end;

  for Indice := 0 to TOBChampsCB.detail.count -1 do
  begin
    if ListChampsCB <> '' then
      ListChampsCB := ListChampsCB+ ','+TOBChampsCB.detail[Indice].getValue('CHAMPS')
    else
      ListChampsCB := TOBChampsCB.detail[Indice].getValue('CHAMPS');
  end;

end;

procedure TFRecupTarifFour.RecupereTarif;
var SizeFic : integer;
    Fichier : textfile;
    SearchRec: TSearchRec;
    ioerr : TIOErr;
    NbrEnregTrait : integer;
    Ligne,LigneDep,LigneFin : integer;
    Erreur : integer;
    texte : string;
begin

  NbrEnregTrait := 0;
  IndexParCatalogu := -1;
  IndexParArticle := -1;

  RemplitChampsTob;
  ioerr := oeOk;

  if (IndexParCatalogu <> -1) or (IndexParArticle <> -1) then
  begin
    TobArticle      := TOB.Create ('', Nil, -1);
    TobCatalogu     := TOB.Create ('', Nil, -1);
    // Modif BTP
    TOBEnt          := TOB.Create ('THE NOMENENT',nil,-1);
    TOBLig          := TOB.Create ('THE NOMENLIG',nil,-1);
    TOBArticleCompl := TOB.Create ('LES COMPLEMENTS ARTICLE',nil,-1);
    TOBTiers        := TOB.Create ('TIERS',nil,-1);
    TOBCB           := TOB.Create ('CODEBARRE',nil,-1);

    // --
    TobArticle.InitValeurs;
    TobCatalogu.InitValeurs;
    //
    Ligne := 0;
    LigneDep := 1;

    TRY
      AssignFile (Fichier, GRF_FICHIER.Text); // ouvre le pointeur sur le fichier
      Reset (Fichier);                        // equivalent open file
      // calcul nb enreg du fichier de données
      readln (Fichier, stEnreg);
      FindFirst(GRF_FICHIER.Text, faAnyFile, SearchRec);
      SizeFic := Trunc(SearchRec.Size / Length(stEnreg));
      FindClose (SearchRec);
      // ---
      Reset (Fichier);
      //-----------------
      InitMoveProgressForm (nil,'Traitement en cours...','Récupération du catalogue fournisseur', SizeFic, true, true) ;
      //
      BStop := False;
      //
      while not EOF(Fichier) and (BStop = False) do
      begin
        if NbrEnregTrait >= MAXTRAITENREG then
        begin
          LigneFin := Ligne;
          Ioerr := EnregistreTables;
          if Ioerr <> oeOk then
          begin
          	EnregistreErreur (99,LigneDep,LigneFin);
            break;
          end;
          LigneDep := Ligne+1;
          NbrEnregTrait := 0;
          VideTobs;
        end;
        Application.ProcessMessages;
        if not MoveCurProgressForm ('') then
        begin
          BStop := True;
        end
        else
        begin
          readln (Fichier,stEnreg);     // lecture des données
          if ControleEnreg then
          try
            inc(NbrEnregTrait);
            if not TraiteEnregistrement (Erreur,Texte) then
            begin
              EnregistreErreur (Erreur,Ligne,-1,Texte);
            end;
          except
            BStop := True;
          end;
          Inc(Ligne);
        end;
      end;

      FiniMoveProgressForm ;

      CloseFile (Fichier);

      if BStop = False then
      begin
        Ioerr := EnregistreTables;
        if Ioerr <> oeOk then
        begin
          LigneFin := Ligne;
          EnregistreErreur (99,LigneDep,LigneFin);
        end;
      end;

    FINALLY
      TobArticle.Free;
      TobCatalogu.Free;
      // MOdif BTP
      TOBEnt.free;
      TOBLig.free;
      TobArticleCompl.Free;
      TOBCB.Free;
      TOBTiers := TOB.Create ('TIERS',nil,-1);
      // --
    END;
  end;
  RenseigneJournalEvenement (ioerr);
end;

procedure TFRecupTarifFour.RecupereValChamp (TobPar : TOB; Index : integer; var ValChamp : string);
var Longueur, Offset : integer;
begin
  if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
  begin
    Longueur := TobPar.GetValue ('GFL_LONGUEUR');
    Offset := TobPar.GetValue ('GFL_OFFSET');
    ValChamp := Copy (stEnreg, Offset, Longueur);
  end else
  begin
  	RecupValChampEnregVar (ValChamp, Index);
  end;
  ValChamp := Trim (ValChamp);
end;

procedure TFRecupTarifFour.RecupValChampEnregVar (var ValChamp : string; IndexChamp : integer);
var stEnregTrav, stEnregTravBis, stSeparateur : string;
    Index, Position : integer;
begin
  if TobParFou.GetValue ('GRF_SEPARATEUR') = 'AUT' then stSeparateur := TobParFou.GetValue ('GRF_SEPTEXTE')
  else if TobParFou.GetValue ('GRF_SEPARATEUR') = 'TAB' then stSeparateur := #9
  else if TobParFou.GetValue ('GRF_SEPARATEUR') = 'PIP' then stSeparateur := '|'
  else if TobParFou.GetValue ('GRF_SEPARATEUR') = 'PTV' then stSeparateur := ';';
  Position := 1;
  stEnregTrav := stEnreg;
  stEnregTrav := stEnregTrav + stSeparateur;
  stEnregTravBis := stEnregTrav;
  for Index := 0 to IndexChamp - 1 do
  begin
    Position := Position + Pos (stSeparateur, stEnregTrav);
    stEnregTrav := Copy (stEnregTravBis, Position, Length (stEnregTravBis) - Position + 1);
  end;
  ValChamp := Copy (stEnregTravBis, Position, Pos (stSeparateur, stEnregTrav) - 1)
end;

function TFRecupTarifFour.GenereLaValeur (ValChamp : string;TOBParam : TOB) : double;
var PartEnt,PartDec : string;
    Position,lgEnt,LgDec : Integer;
    Separateur : string;
begin

  result := 0;

  if (TOBParam.GetValue('GFL_SEPDECIMALE') = '') and (TOBParam.GetValue('GFL_NBRDECIMALE')=0) then
  begin
    Result := valeur(ValChamp); // compatibilite avec l'existant
    exit;
  end;

  if TOBParam.getValue('GFL_SEPDECIMALE')<> '' then
  begin
    if TOBParam.getValue('GFL_SEPDECIMALE') = 'PT' then Separateur := '.' else
    if TOBParam.getValue('GFL_SEPDECIMALE') = 'VIR' then Separateur := ',' else Separateur := '';
    Position :=  Pos(Separateur,ValChamp);
    if Position = 0 then BEGIN result := valeur (ValChamp); Exit; END;
    PartEnt := copy(ValChamp,1,Position-1);
    PartDec := copy(ValChamp,Position+1,255);
    Result := Valeur(PartEnt)+(Valeur(PartDec)/power(10,Length(PartDec)));
  end
  else if TOBParam.GetValue('GFL_NBRDECIMALE') > 0 then
  begin
    LgEnt := Length(ValChamp)-TOBParam.GetValue('GFL_NBRDECIMALE');
    LgDec := TOBParam.GetValue('GFL_NBRDECIMALE');
    PartEnt := copy(ValChamp,1,LgEnt);
    PartDec := copy(ValChamp,LgEnt+1,LgDec);
    Result := Valeur(PartEnt)+(Valeur(PartDec)/power(10,Length(PartDec)));
  end;

end;

function TFRecupTarifFour.FindTablette (champ : string;TobParam : TOB) : TOB;
var Lesuffixe : string;
    TOBT : TOB;
    indice : integer;
begin
  LeSuffixe := copy (champ,pos('_',champ)+1,255);
  // Boucle pour rechercher le suffixe car dans DO_NOMCHAMP il peut y avoir plusieurs noms séparés par point-virgule
  result := Nil;
  For indice :=0 to TOBTablettes.detail.count -1 do
  begin
    TOBT := TOBTablettes.detail[Indice];
    if Pos (LeSuffixe,TOBT.GetValue('DO_NOMCHAMP')) > 0 then
    begin
      result := TOBT;
      Exit;
    end;
  end;
end;

function TFRecupTarifFour.RecupTablettes(champ : string;TOBParam: TOB;Var TheTable : string; Var TheCodeRec : string ; var TheType : string; Var LePrefixe : string) : tob;
var Lesuffixe,TheRequete: string;
    laTabl : TOB;
    TheQuery : TQuery;
begin
  result:= nil;
  LeSuffixe := copy (champ,pos('_',champ)+1,255);
  Therequete := 'SELECT * FROM DECOMBOS WHERE DO_NOMCHAMP LIKE "%'+LeSuffixe+'%"';

  TheQuery := OpenSql (TheRequete,true,-1,'',true);
  if TheQuery.eof then BEGIN ferme(TheQuery); Exit; END;
  laTabl := TOB.Create ('DECOMBOS',TOBTablettes,-1);
  laTabl.SelectDB ('',TheQuery);
  ferme (TheQuery);

  lePrefixe := laTabl.GetValue ('DO_PREFIXE');
  if (lePrefixe<> 'CC') and (lePrefixe <> 'CO') then BEGIN laTabl.free; Exit; END;

  result := laTabl;
  if TobParam.GetValue('GFL_FICHIER') <> '' then
  begin
    laTabl.addchampSupValeur ('FICHIER LIE',extractfilePath (GRF_FICHIER.text)+TobParam.GetValue('GFL_FICHIER'),false)
  end else
  begin
    laTabl.addchampSupValeur ('FICHIER LIE','',false)
  end;

  if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
  begin
    LaTabl.addchampsupValeur('LONGUEUR',TobParam.GetValue ('GFL_LONGUEUR'),false);
    LaTabl.addchampsupValeur('OFFSET',1,false);
    LaTabl.addchampsupValeur('OFFSETLIB',LaTabl.GetValue('LONGUEUR')+1,false);
    LaTabl.addchampsupValeur('SEPARATEUR','',false);
  end else
  begin
    LaTabl.addchampsupValeur('LONGUEUR',3,false);
    LaTabl.addchampsupValeur('OFFSET',TobParam.GetValue('GFL_OFFSET'),false);
    LaTabl.addchampsupValeur('OFFSETLIB',4,false);
    LaTabl.addchampsupValeur('SEPARATEUR',TobParFou.GetValue ('GRF_SEPARATEUR'),false);
  end;


  TheTable := PrefixeToTable (laTabl.getValue('DO_PREFIXE'));
  TheType := laTabl.GetValue('DO_TYPE');

  // Recup des lignes existantes
  TheRequete := 'SELECT "X" AS EXIST, * FROM '+TheTable+' WHERE '+lePrefixe+'_TYPE="'+laTabl.GetValue('DO_TYPE')+'"';
  TheQuery := openSql (TheRequete,True,-1,'',true);
  TRY
    if not TheQuery.eof then
    begin
      laTabl.LoadDetailDB (TheTable,'','',TheQuery,false);
      //for Indice := 0 to laTabl.detail.count -1 do
      //begin
      //  laTabl.detail[Indice].addchampsupValeur('EXIST','X',false);
      //end;
      end;
  FINALLY
  Ferme (theQuery);
  END;
end;

function TFRecupTarifFour.InFamille (recherche,ChaineControle : string) : boolean;
var SLoc : string;
    SCode : string;
begin
  result := false;
  SLoc := ChaineControle;
  repeat
    SCode := readtokenst (SLoc);
    if Scode <> '' then
    begin
      if SCode = recherche Then BEGIN result := true; break; END;
    end;
  until SCode = '';
end;

procedure TFRecupTarifFour.FindAjouteOnIT (ThisTablette : TOB; LaTable, LePrefixe, Letype, Valeur : string;TobTable: TOB; IndextobTable : integer);
var ThisLigne : TOB;
    TheFather : string;
    TheGdFather : String;
    NivPrec : integer;
    Code    : String;
    Libelle : String;
begin
  Code := Copy(Valeur,0,3);
  //
  ThisLigne := ThisTablette.findFirst ([LePrefixe+'_TYPE',LePrefixe+'_CODE'],[LeType, Code],true);
  //
  if ThisLigne = nil then
  Begin
    if Letype = 'YCB' then exit;
    ThisLigne := TOB.Create (LaTable,ThisTablette,-1);
    ThisLigne.PutValue (LePrefixe+'_TYPE',LeType);
    ThisLigne.putValue (LePrefixe+'_CODE',Code);
    //FV1 : 06/03/2017 - FS#2410 - SES ETANCH - Import tarif ne reprends pas les libellés des familles articles contenu ds le TXT
    Libelle := RechercheLibelleTablette(Code, ThisTablette);
    ThisLigne.putValue (LePrefixe+'_LIBELLE', Libelle);
    ThisLigne.putValue (LePrefixe+'_ABREGE',  Libelle);
    ThisLigne.putValue (LePrefixe+'_LIBRE',   '');
    ThisLigne.addchampsupValeur('EXIST','-',  false);
  end;
  //

  if (copy(LeType,1,2) = 'FN') and (TObTable.detail[IndexTobTable].NomTable = 'ARTICLE') and
     (strtoint(copy(leType,3,1))> 1) and
     (GetParamSoc('SO_GCFAMHIERARCHIQUE')=True) then // tablette famille d'article
  begin
    TheFather   := '';
    TheGdFather := '';
    NivPrec := strtoint(copy(LeType,3,1)) - 1;
    TheFather := TObTable.detail[IndexTobTable].GetValue('GA_FAMILLENIV'+inttostr(NivPrec));
    //Gestion Famille de niveau 3.... Etrangement pas géré ??? 
    if NivPrec  > 1 then TheGdFather := TObTable.detail[IndexTobTable].GetValue('GA_FAMILLENIV'+inttostr(NivPrec-1));
    //
    if ThisLigne.getValue(LePrefixe+'_LIBRE') = '' then
    begin
       if TheGdFather <> '' then
         ThisLigne.putValue (LePrefixe+'_LIBRE',TheGdFather + TheFather)
       else
       ThisLigne.putValue (LePrefixe+'_LIBRE',TheFather);
    end
    else
    begin
      If TheGdFather <> '' then
      begin
        if not (InFamille(TheGdFather+TheFather,ThisLigne.GetValue(LePrefixe+'_LIBRE'))) then
        begin
          ThisLigne.putValue (LePrefixe+'_LIBRE',ThisLigne.GetValue (LePrefixe+'_LIBRE')+';'+TheGdFather+TheFather);  //Je sais pas comment ça peut marcher par la suite ?????
        end;
      end
      else
      begin
      if not (InFamille(TheFather,ThisLigne.GetValue(LePrefixe+'_LIBRE'))) then
      begin
        ThisLigne.putValue (LePrefixe+'_LIBRE',ThisLigne.GetValue (LePrefixe+'_LIBRE')+';'+TheFather);
      end;
    end;
  end;
  //
  end;
  //
  if (LeType = 'BFT') and (TObTable.detail[IndexTobTable].NomTable = 'ARTICLE') then
  begin
    TheFather := TObTable.detail[IndexTobTable].GetValue('GA_TARIFARTICLE');
    ThisLigne.putValue (LePrefixe+'_LIBRE',TheFather);
  end;
end;

procedure TFRecupTarifFour.AddTablettes(champ, valeur: string ; TOBParam ,TobTable: TOB; IndextobTable : integer);
var ThisTablette : TOB;
    LaTable,LeCodeRec,LeType,LePrefixe : string;
begin
  ThisTablette := RecupTablettes (Champ,TOBParam,LaTable,LeCodeRec,LeType,LePrefixe);
  if ThisTablette <> nil then
  begin
    FindAjouteOnIT (ThisTablette ,LaTable, LePrefixe, Letype, Valeur,TOBTable,IndexTobTable);
  end;
end;

//FV1 : 06/03/2017 - FS#2410 - SES ETANCH - Import tarif ne reprends pas les libellés des familles articles contenu ds le TXT
Function TFRecupTarifFour.RechercheLibelleTablette(Thecode : string; TheTablette : TOB) : string;
var Fichier   : textFile;
    TheEnreg  : string;
    Code      : string;
    Longueur  : Integer;
    TOBL      : TOB;
    TheFIc : string;
begin

  Result := '';

  if TheTablette = nil then
  begin
    result := TheCode;
    Exit;
  end;

  if TheTablette.GetValue('FICHIER LIE') = '' then
  begin
    result := TheCode;
    Exit;
  end;

  if (TheTablette.GetValue('FICHIER LIE') <> '') and (TheTablette.detail.count > 0) then
  begin
    if not FileExists(TheTablette.GetValue('FICHIER LIE')) then
    begin
      result := TheCode;
      Exit;
    end;
  end;
  TheFIc := TheTablette.getValue('FICHIER LIE');
  AssignFile (Fichier, TheFIc); // ouvre le pointeur sur le fichier
  Reset (Fichier);  // equivalent open file

  Longueur := StrToInt(TheTablette.getValue('LONGUEUR'));

  TRY
    while not EOF(Fichier) do
    begin
      readln (Fichier, TheEnreg);
      Code := Trim (UpperCase(Copy(TheEnreg,0, Longueur)));
      IF Code <> '' then
      begin
        if TheCode = Code then
        begin
          TOBL := TheTablette.findFirst (['DO_CODE'],[Code],True);
          if TOBL = nil then
          begin
            Result := Copy(TheEnreg, Longueur+1, 255);
            if Result = '' then Result := Code;
          end
          else
            Result := TOBL.GetValue('DO_CHAMPLIB');
        end;
      end;
    end;
  FINALLY
    CloseFile (Fichier);
  end;

end;


procedure TFRecupTarifFour.CumulTablette(laTabl: TOB; champ, valeur: string;TobTable: TOB; IndextobTable : integer);
var LaTable,LePrefixe,LeType : string;
begin

  LaTable := PrefixeToTable (laTabl.getValue('DO_PREFIXE'));
  lePrefixe := laTabl.GetValue ('DO_PREFIXE');
  LeType := laTabl.GetValue('DO_TYPE');

  FindAjouteOnIT (laTabl ,LaTable, LePrefixe, Letype, Valeur, TOBTable, IndexTobTable);
end;

procedure TFRecupTarifFour.TraiteChamp (IndexTobTable : integer; ValChamp, Champ : string; TobTable,TOBParam : TOB;IndexArt : integer = -1);
var TypeChamp : string;
    iTable, iChamp : integer;
    TheValeur : double;
    laTabl : TOB;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded(); 

  if pos ('BCB_', Champ) > 0 then
  begin
  	Table := Mcd.getTable ('BTCODEBARRE');
  end else if pos ('GCA_', Champ) > 0 then
  begin
  	Table := Mcd.getTable ('CATALOGU');
  end else if pos ('GA2_', Champ) > 0 then
  begin
  	Table := Mcd.getTable ('ARTICLECOMPL');
  end else Table := Mcd.getTable ('ARTICLE');

  TypeChamp := Mcd.GetField(Champ).Tipe;

  if TypeChamp = 'DATE' then
  begin
    if IsValidDate(ValChamp) then
    begin
      if (TobTable.detail[IndexTobTable].Getvalue ('UPDATE')='X') and
         (TobTable.Detail[IndexTobTable].GetValue (Champ) <> strtodate (ValChamp) ) then
      begin
        TobTable.Detail[IndexTobTable].Putvalue ('ISCHAMPSMODIFIE','X');
      end;
      TobTable.Detail[IndexTobTable].PutValue (Champ, strtodate (ValChamp));
    end;
  end
  else if TypeChamp = 'INTEGER' then
  begin
      if (TobTable.detail[IndexTobTable].Getvalue ('UPDATE')='X') and
         (TobTable.Detail[IndexTobTable].GetValue (Champ) <> strtoint (ValChamp) ) then
      begin
        TobTable.Detail[IndexTobTable].Putvalue ('ISCHAMPSMODIFIE','X');
      end;
      TobTable.Detail[IndexTobTable].PutValue (Champ, strtoint (ValChamp));
  end
  else if (TypeChamp = 'DOUBLE') or (TypeChamp = 'EXTENDED') then
    begin
      // Modif BTP
        TheValeur := GenereLaValeur (ValChamp,TOBParam);
        // --
        if (TobTable.detail[IndexTobTable].Getvalue ('UPDATE')='X') and
           (TobTable.Detail[IndexTobTable].GetValue (Champ) <> TheValeur ) then
        begin
          TobTable.Detail[IndexTobTable].Putvalue ('ISCHAMPSMODIFIE','X');
          if ((champ = 'GCA_PRIXBASE') or (champ = 'GCA_PRIXVENTE')) then
          begin
            if IndexArt >= 0 then Tobarticle.Detail[IndexArt].Putvalue ('ISCHAMPSMODIFIE','X');
          end;
        end;
        TobTable.Detail[IndexTobTable].PutValue (Champ, TheValeur);
  end
  else if TypeChamp = 'COMBO' then // Gestion des tablettes
      begin
        valchamp:=uppercase(valchamp); // FQ11788 : on passe le code en majuscule pour les combos
    //
        if (TobTable.detail[IndexTobTable].Getvalue ('UPDATE')='X') and
           (TobTable.Detail[IndexTobTable].GetValue (Champ) <> valchamp ) then
        begin
          TobTable.Detail[IndexTobTable].Putvalue ('ISCHAMPSMODIFIE','X');
        end;
    //
    TobTable.Detail[IndexTobTable].PutValue (Champ, valchamp);
    if ValChamp <> '' then
    begin
      laTabl := FindTablette (champ,TobParam);
      //
      if laTabl = nil then
        AddTablettes (champ,ValChamp,TOBParam,TobTable,IndextobTable)
      else
        CumulTablette (laTabl,champ,ValChamp,TobTable,IndextobTable);
    end;
  end
  else
      begin
        if (TobTable.detail[IndexTobTable].Getvalue ('UPDATE')='X') and
           (TobTable.Detail[IndexTobTable].GetValue (Champ) <> valchamp ) and
       ((champ <> 'GA_CODEARTICLE') and
        (champ <> 'GA_ARTICLE')     and
        (champ <> 'GA_LIBELLE')     and
        (champ <> 'GCA_REFERENCE')  and
        (Champ <> 'GCA_ARTICLE'))  then
        begin
          TobTable.Detail[IndexTobTable].Putvalue ('ISCHAMPSMODIFIE','X');
        end;
        // Modif BRL 28/09/2012 : chargement du code article unique si celui-ci est présent dans le paramétrage
    if Champ = 'GCA_ARTICLE'       then ValChamp := CodeArticleUnique2 (Valchamp, '');
    //
        TobTable.Detail[IndexTobTable].PutValue (Champ, ValChamp);
      end;

end;

procedure TFRecupTarifFour.RenseigneJournalEvenement (ioerr : TIOErr ; Partiel : boolean = false);
var TOBJnal : TOB ;
    NumEvt : integer ;
    QQ : TQuery ;
begin
  NumEvt:=0 ;
  TOBJnal:=TOB.Create('JNALEVENT', Nil, -1) ;
  TOBJnal.PutValue('GEV_TYPEEVENT', 'MTF');
  TOBJnal.PutValue('GEV_LIBELLE', PTITRE.Caption);
  TOBJnal.PutValue('GEV_DATEEVENT', Date);
  TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
  QQ:=OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT',True,-1,'',true) ;
  if Not QQ.EOF then NumEvt:=QQ.Fields[0].AsInteger ;
  Ferme(QQ) ;
  Inc(NumEvt) ;
  TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
  ListRecap.Items.Add ('');
  if BStop = True then
  begin
    TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
    ListRecap.Items.add ('Interrompue par l''utilisateur');
  end else
  begin
    if ioerr = oeOk then
    begin
      TOBJnal.PutValue('GEV_ETATEVENT', 'OK');
      if Partiel then ListRecap.Items.Add('Traitement partiel terminé')
                 else ListRecap.Items.Add('Traitement terminé');
    end else
    begin
      TOBJnal.PutValue ('GEV_ETATEVENT', 'ERR');
      if Partiel then ListRecap.Items.Add('Traitement partiel non terminé')
                 else ListRecap.Items.Add('Traitement non terminé');
    end;
  end;
  TOBJnal.PutValue('GEV_BLOCNOTE', ListRecap.Items.Text);
  TOBJnal.InsertDB(Nil) ;
  TOBJnal.Free ;
end;

FUNCTION TFRecupTarifFour.GetFournisseurFromEnreg : string;
var index : Integer;
    ValChamp, Champ : string;
begin
	Result := '';
  Index := 0;
  while (Index < TobParFouLig.Detail.Count) do
  begin
    Champ := TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP');
    if Champ = 'GCA_TIERS' then
    begin
    	RecupereValChamp (TobParFouLig.Detail[Index], Index, Valchamp);
      Result := ValChamp;
      break;
    end;
    inc(Index);
  end;
end;

FUNCTION TFRecupTarifFour.GetArticleFromEnreg : string;
var index : Integer;
    ValChamp, Champ : string;
begin
	Result := '';
  Index := 0;
  while (Index < TobParFouLig.Detail.Count) do
  begin
    Champ := TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP');
    if Champ = 'GCA_ARTICLE' then
    begin
    	RecupereValChamp (TobParFouLig.Detail[Index], Index, Valchamp);
      Result := ValChamp;
      break;
    end;
    inc(Index);
  end;
end;

function TFRecupTarifFour.TraiteEnregistrement (var codeErreur : integer; var Texte : string) : boolean;
var IndexCatalogu, IndexArticle, Index : integer;
    ValChamp, Champ, Codeart : string;
begin
  Result := true;
  IndexArticle := -1;
  IndexCatalogu := -1;
  CodeErreur := 0;
  Texte := '';

  if GRF_MULTIFOU.Checked then
  begin
  	TheFournisseur := GetFournisseurFromEnreg;
  	Codeart := GetArticleFromEnreg;
    if (CBCreationCatalogue.Checked) and (TheFournisseur = '') then
    begin
      // on doit créer les catalogues et pas de fournisseur ?? --> erreur...
      codeErreur := 1; // --> Code fournisseur non renseigné
      Texte := 'Article : ' + Codeart;
    	Result := false;
      Exit;
    end;
    if not ExisteCodeTiers('FOU',TheFournisseur) then
    begin
      codeErreur := 2; // --> Code fournisseur inexistant
      Texte := TheFournisseur + '- Article : ' + Codeart;
      Result := false;
      exit;
    end;
  end;

  if IndexParArticle <> -1 then
  begin
    CreationElementTob ('ARTICLE', IndexParArticle, TobArticle, CBCreationArticle.Checked, IndexArticle);
  end
  else if IndexParCatalogu <> -1 then
  begin
    CreationElementTob ('ARTICLE', IndexParCatalogu, TobArticle, CBCreationArticle.Checked, IndexArticle);
  end;

  if IndexParCatalogu <> -1 then
  begin
    CreationElementTob ('CATALOGU', IndexParCatalogu, TobCatalogu, CBCreationCatalogue.Checked, IndexCatalogu);
  end;

  Index := 0;

  while (Index < TobParFouLig.Detail.Count) and (Result) do
  begin
    Champ := TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP');
    if Champ <> 'PASS' then
    begin
      if (Champ <> 'GCA_REFERENCE') and (((IndexArticle <> -1) and (Champ < 'GCA_')) or
         ((Champ > 'GCA_') and (IndexCatalogu <> -1))) then
      begin
        RecupereValChamp (TobParFouLig.Detail[Index], Index, Valchamp);
        if champ > 'GCA_' then
        begin
          // Traitement des champs du catalogue
          TraiteChamp (IndexCatalogu, ValChamp, Champ, TobCatalogu,TobParFouLig.Detail[Index],IndexArticle);
          // EXCEPTION DE GESTION
          if (Champ = 'GCA_LIBELLE') and (IndexArticle <> -1) then
          begin
            if TOBArticle.detail[indexArticle].GetValue('UPDATE') = '-' then
            begin
              // en création uniquement
              TOBArticle.detail[indexArticle].PutValue('GA_LIBELLE',ValChamp);
              if TOBArticle.detail[indexArticle].GetValue('GA_TYPEARTICLE') = 'ARP' then
              begin
                MajInfoLig (TobArticle.Detail[IndexArticle].GetValue ('GA_ARTICLE'),
                TobArticle.Detail[IndexArticle].GetValue ('GA_LIBELLE'));
                Tobarticle.Detail[IndexArticle].Putvalue ('ISCHAMPSMODIFIE','X');
              end;
            end;
          end;
        end else if Pos('GA2_',Champ) > 0 then
        begin
          if Champ = 'GA2_SOUSFAMTARART' then
          begin
            TraiteChamp (IndexArticle, ValChamp, Champ, TobArticle,TobParFouLig.Detail[Index]);
          end;
        end
        else if Pos('BCB_',Champ) > 0 then
        begin
          TraiteChamp (IndexArticle, ValChamp, Champ, TobArticle,TobParFouLig.Detail[Index]);
        end
        else
        begin
          if (champ <> 'GA_ARTICLE') and (champ <> 'GA_CODEARTICLE') and (IndexArticle <> -1) then
          begin
            // Traitement des champs de l'article
            TraiteChamp (IndexArticle, ValChamp, Champ, TobArticle,TobParFouLig.Detail[Index]);

            if (Champ = 'GA_LIBELLE') and (IndexArticle <> -1) then
            begin
              if TOBArticle.detail[indexArticle].GetValue('UPDATE') = '-' then
              begin
                // en création uniquement
                TOBArticle.detail[indexArticle].PutValue('GA_LIBELLE',ValChamp);
                if TOBArticle.detail[indexArticle].GetValue('GA_TYPEARTICLE') = 'ARP' then
                begin
                  MajInfoLig (TobArticle.Detail[IndexArticle].GetValue ('GA_ARTICLE'),
                  TobArticle.Detail[IndexArticle].GetValue ('GA_LIBELLE'));
                  Tobarticle.Detail[IndexArticle].Putvalue ('ISCHAMPSMODIFIE','X');
                end;
              end;
            end;
          end;
        end;
      end;
    end;
    inc(Index);
  end;

end;

{==============================================================================================}
{=============================== Utils ========================================================}
{==============================================================================================}

(* function TFRecupTarifFour.ExtractLibelle ( St : string) : string;
Var St_Chaine : string ;
    i_pos : integer ;
begin
Result := '';
i_pos := Pos ('&', St);
if i_pos > 0 then
    begin
    St_Chaine := Copy (St, 1, i_pos - 1) + Copy (St, i_pos + 1, Length(St));
    end else St_Chaine := St;
Result := St_Chaine + ' : ';
end; *)

procedure TFRecupTarifFour.SetDetailPrixPose(IndexArt: integer);
var TOBE,TOBD,TOBART,TOBA : TOB;
    Article,Libelle : string;
begin
   TOBART := TobArticle.detail[IndexArt];

   Article := TOBART.GetValue('GA_ARTICLE');
   Libelle := TOBART.GetValue('GA_LIBELLE');
   // Entete Nomenclature
   TOBE := TOB.Create ('NOMENENT',TOBENT,-1);
   TOBE.putValue ('GNE_NOMENCLATURE',Article);
   TOBE.putValue ('GNE_LIBELLE',Libelle);
   TOBE.putValue ('GNE_ARTICLE',Article);
   TOBE.putValue ('GNE_DATECREATION',V_PGI.DateEntree);
   TOBE.putValue ('GNE_QTEDUDETAIL',1);
   // Ligne Nomenclature
   TOBD := TOB.Create ('NOMENLIG',TOBLig,-1);
   TOBD.putvalue ('GNL_NOMENCLATURE',Article);
   TOBD.putvalue ('GNL_NUMLIGNE',1);
   TOBD.putValue ('GNL_LIBELLE',Libelle);
   TOBD.PutValue ('GNL_CODEARTICLE',Copy(Article,1,18));
   TOBD.PutValue ('GNL_QTE',1);
   TOBD.PutValue ('GNL_JOKER','N');
   TOBD.PutValue ('GNL_ARTICLE',Article);
   TOBD.PutValue ('GNL_DATECREATION',V_PGI.DateEntree);
   TOBD.PutValue ('GNL_CREATEUR',V_PGI.User);
   TOBD.PutValue ('GNL_UTILISATEUR',V_PGI.User);
   if TobParFou.getValue('GRF_PRESTATION') <> '' then
   begin
     TOBA := TOB.Create ('ARTICLE',nil,-1);
     TRY
       TOBA.putvalue('GA_ARTICLE',TobParFou.getValue('GRF_PRESTATION'));
       TOBA.loadDb (true);
     // Ligne MO
       TOBD := TOB.Create ('NOMENLIG',TOBLig,-1);
       TOBD.putvalue ('GNL_NOMENCLATURE',Article);
       TOBD.putvalue ('GNL_NUMLIGNE',2);
       TOBD.putValue ('GNL_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
       TOBD.PutValue ('GNL_CODEARTICLE',Copy(TobParFou.getValue('GRF_PRESTATION'),1,18));
       TOBD.PutValue ('GNL_QTE',0);
       TOBD.PutValue ('GNL_JOKER','N');
       TOBD.PutValue ('GNL_ARTICLE',TobParFou.getValue('GRF_PRESTATION'));
       TOBD.PutValue ('GNL_DATECREATION',V_PGI.DateEntree);
       TOBD.PutValue ('GNL_CREATEUR',V_PGI.User);
       TOBD.PutValue ('GNL_UTILISATEUR',V_PGI.User);
     FINALLY
       TOBA.free;
     END;
   end;
end;

(* Modif BTP *)

procedure TFRecupTarifFour.MajLibelletablette(TheTablette : TOB);
var Fichier : textFile;
    TheEnreg,Code,Libelle : string;
    TOBL : TOB;
begin

  AssignFile (Fichier, TheTablette.getValue('FICHIER LIE')); // ouvre le pointeur sur le fichier
  Reset (Fichier);  // equivalent open file

  TRY
//    readln (Fichier, TheEnreg);
    while not EOF(Fichier) do
    begin
      readln (Fichier, TheEnreg);
      Code := RecupereValChampTablette (TheEnreg,TheTablette.getValue('LONGUEUR'),TheTablette.getValue('OFFSET'),0);
      IF Code <> '' then
      begin
        TOBL := TheTablette.findFirst ([TheTablette.getValue('DO_CODE')],[Code],True);
        if TOBL <> nil then
        begin
          Libelle := RecupereValChampTablette (TheEnreg,255,TheTablette.getValue('OFFSETLIB'),1);
          if Libelle <> '' then TOBL.putValue(TheTablette.getValue('DO_CHAMPLIB'),Libelle);
        end;
      end;
//      readln (Fichier, TheEnreg);
    end;
  FINALLY
    CloseFile (Fichier);
  end;

end;

//
function TFRecupTarifFour.RecupereValChampTablette (TheEnreg : string; longueur,offset,index : integer) : string;
var valChamp : string;
    i_pos : integer;
begin
  if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
  begin
    ValChamp := Copy (TheEnreg, Offset, Longueur);
    // on considere que plus de 5 blanc définis que c'est la fin de la zone
    // pas terrible ...en attendant la definition du fichier associé a la tablette
    i_pos := pos ('     ',valchamp);
    if i_pos > 0 then ValChamp := trim(copy (ValChamp,1,i_pos));
  end else
  begin
  	valchamp := RecupValChampEnregVartablette (TheEnreg,Index);
  end;
  ValChamp := Trim (ValChamp);
  result := valchamp;
end;

function TFRecupTarifFour.RecupValChampEnregVarTablette (TheEnreg : string; IndexChamp : integer) : string;
var stEnregTrav, stEnregTravBis, stSeparateur : string;
    Index, Position : integer;
begin
  if TobParFou.GetValue ('GRF_SEPARATEUR') = 'AUT' then stSeparateur := TobParFou.GetValue ('GRF_SEPTEXTE')
  else if TobParFou.GetValue ('GRF_SEPARATEUR') = 'TAB' then stSeparateur := #9
  else if TobParFou.GetValue ('GRF_SEPARATEUR') = 'PIP' then stSeparateur := '|'
  else if TobParFou.GetValue ('GRF_SEPARATEUR') = 'PTV' then stSeparateur := ';';
  Position := 1;
  stEnregTrav := TheEnreg;
  stEnregTrav := stEnregTrav + stSeparateur;
  stEnregTravBis := stEnregTrav;
  for Index := 0 to IndexChamp - 1 do
  begin
    Position := Position + Pos (stSeparateur, stEnregTrav);
    stEnregTrav := Copy (stEnregTravBis, Position, Length (stEnregTravBis) - Position + 1);
  end;
  result := Copy (stEnregTravBis, Position, Pos (stSeparateur, stEnregTrav) - 1);
end;

procedure TFRecupTarifFour.MajInfoLig (CodeArt,Libelle : string);
var TOBE, TOBD : TOB;
begin
   TOBE := TOBENT.FindFirst (['GNE_ARTICLE'],[CodeArt],true);
   if TOBE <> nil then TOBE.putValue ('GNE_LIBELLE',Libelle);
   TOBD := TOBLig.FindFirst (['GNL_NOMENCLATURE','GNL_ARTICLE'],[CodeArt,CodeArt],true);
   if TOBD <> nil then TOBD.putValue ('GNL_LIBELLE',Libelle);
end;

{$IFDEF BTP}
function TFRecupTarifFour.FindCatalogue (Article,Fournisseur : string) : TOB;
begin
  result := TobCatalogu.findfirst(['GCA_ARTICLE', 'GCA_TIERS'],[Article,fournisseur],true);
end;

procedure TFRecupTarifFour.DefinieValorisation (TOBref: TOB);
var  TOBCatalog,TOBTarif: TOB;
     MTPAF : double;
     TypeArticle :string;
     Fournisseur : string;
     PrixPourQte : double;
     IsUniteAchat : boolean;
begin
  IsUniteAchat := true; // par défaut
  TypeArticle:= TOBRef.GetValue('GA_TYPEARTICLE');
  Fournisseur := TOBRef.GetValue('GA_FOURNPRINC');
  if (Fournisseur <> TOBTiers.getValue('T_TIERS')) and (Fournisseur <> '') then
  begin
    TOBTiers.InitValeurs;
    TOBTiers.putValue('T_TIERS',Fournisseur);
    TOBTiers.LoadDB (true);
  end;
  if (TypeArticle<> 'MAR') and (TypeArticle <> 'ARP') and (TypeArticle <> 'PRE') then exit;

  TOBTarif := TOB.Create ('TARIF',nil,-1);
  TOBCatalog := FindCatalogue (TOBRef.GetValue('GA_ARTICLE'),Fournisseur);
  TRY
    if TOBCatalog <> nil then
    begin
      if Fournisseur <> '' then
      begin
        GetTarifGlobal (TOBRef.getValue('GA_ARTICLE'),TOBRef.getValue('GA_TARIFARTICLE'),TOBRef.getValue('GA2_SOUSFAMTARART'),'ACH',TOBRef,TOBTiers,TOBTarif,true);
        if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
        begin
          MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
        end else
        begin
          PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
          if PrixPourQte = 0 then PrixPourQte := 1;
          if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte
                                                      else MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte;
        end;
        if MTPAF = 0 then
        begin
          MTPAF := TobRef.getValue('GA_PAHT');
          IsUniteAchat := false;
        end;
        MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
        // voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
        if IsUniteAchat then MTPAF := PassageUAUV (TOBRef,TOBCatalog,MTPAF);
        //
        if MTPAF <> TOBREF.GetValue('GA_PAHT') then
        begin
          TOBREF.putValue('GA_PAHT',MTPAF);
        end;
        RecalculPrPV (TOBRef,TOBCatalog);
      end;
    end else
    begin
      RecalculPrPV (TOBRef,nil);
    end;
  FINALLY
  	TOBTarif.free;
  end;
end;
{$ENDIF}

procedure TFRecupTarifFour.ElimineDoublonsCatalog (TobCatalogu : TOB);

	procedure ElimineCatalog (Code, Fournisseur : string;TOBCatalogu : TOB);
  var TOBC : TOB;
  begin
  	TOBC := TOBCatalogu.findFirst(['GCA_REFERENCE','GCA_TIERS'],[Code,Fournisseur],true);
    if TOBC <> nil then
    begin
      TOBC := TOBCatalogu.findNext(['GCA_REFERENCE','GCA_TIERS'],[Code,Fournisseur],true);
      repeat
        if TOBC <> nil then
        begin
          TOBC.Free;
      		TOBC := TOBCatalogu.findNext(['GCA_REFERENCE','GCA_TIERS'],[Code,Fournisseur],true);
        end;
      until TOBC = nil;
    end;
  end;

var indice : integer;
		TOBC : TOB;
    Code, Fournisseur : string;
    TOBLocale : TOB;
begin
	TOBLocale := TOB.Create ('CATALOGU',nil,-1);
	TOBLocale.Dupliquer (TOBCatalogu,true,true);
  TRY
    For indice :=0 to TOBLocale.detail.count -1 do
    begin
      TOBC := TOBlocale.detail[Indice];
      Code := TOBC.getValue('GCA_REFERENCE');
      Fournisseur := TOBC.getValue('GCA_TIERS');
      ElimineCatalog (Code, Fournisseur, TOBCatalogu);
    end;
  FINALLY
  	TOBLOCALE.Free;
  END;
end;

procedure TFRecupTarifFour.ElimineDoublonsArticles (TobArticles : TOB);

	procedure ElimineArticle (Code : string;TOBArticles : TOB);
  var TOBC : TOB;
  begin
  	TOBC := TOBArticles.findFirst(['GA_ARTICLE'],[Code],true);
    if TOBC <> nil then
    begin
      TOBC := TOBArticles.findNext(['GA_ARTICLE'],[Code],true);
      repeat
        if TOBC <> nil then
        begin
          TOBC.Free;
      		TOBC := TOBArticles.findNext(['GA_ARTICLE'],[Code],true);
        end;
      until TOBC = nil;
    end;
  end;

var indice : integer;
		TOBC : TOB;
    Code : string;
    TOBLocale : TOB;
begin
	TOBLocale := TOB.Create ('ARTICLE',nil,-1);
	TOBLocale.Dupliquer (TOBArticles,true,true);
  TRY
    For indice :=0 to TOBLocale.detail.count -1 do
    begin
      TOBC := TOBlocale.detail[Indice];
      Code := TOBC.getValue('GA_ARTICLE');
      ElimineArticle (Code,TOBArticles);
    end;
  FINALLY
  	TOBLOCALE.Free;
  END;
end;

procedure TFRecupTarifFour.EnregistreErreur(Erreur, LigneDep : integer; ligneFin: integer=-1; Texte:string='');
begin
	if Erreur = 1 then
  begin
    CreationErreur ('Fournisseur non renseigné. '+Texte,LigneDep);
  end else if Erreur = 2 then
  begin
    CreationErreur ('Fournisseur inconnu '+Texte,LigneDep);
  end else if Erreur = 99 then
  begin
    CreationErreur ('Erreur d''enregistrement des données entre la ligne',LigneDep,LigneFin);
  end;
end;

procedure TFRecupTarifFour.AffichageErreurs;
begin
	RapportsErreurs (TOBErreurs);
end;

procedure TFRecupTarifFour.CreationErreur(stErreur: string; LigneDep, LigneFin: integer);
var TOBL : TOB;
    stErr : string;
begin
	TOBL := TOB.Create('UNE ERREUR',TOBErreurs,-1);
  if LigneFin <> -1 then
  begin
  	stErr := stErreur +' '+IntToStr(LigneDep)+ ' et la ligne '+InttoStr(LigneFin);
  end else
  begin
  	stErr := stErreur + ' sur la ligne '+InttoStr(LigneDep);
  end;
  TOBL.AddChampSupValeur('RAPPORT',stErr);
end;

end.
