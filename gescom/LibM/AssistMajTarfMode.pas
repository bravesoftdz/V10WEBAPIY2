unit AssistMajTarfMode;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, StdCtrls, HTB97, ComCtrls, ExtCtrls, Hctrls,
  HPanel, HEnt1, UtilGC, EntGc, Ent1, SaisUtil,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  dbTables, DBGrids, db,
  {$ENDIF}
  UTob, Grids, UDimArticle, Math, TarifUtil, AssistRecopiTarf, UtilArticle;

procedure MiseAJourTarifMode(TOBA: TOB; STATUT, NATURETYPE: string);

type
  TFMajTarfMode = class(TFAssist)
    TabSheet1: TTabSheet;
    PTITRE: THPanel;
    TINTRO: THLabel;
    GBTARIF: TGroupBox;
    TTYPETARIF: THLabel;
    TPERIODE: THLabel;
    TYPETARIF: THValComboBox;
    PERIODE: THValComboBox;
    TabSheet2: TTabSheet;
    GLISTE: THGrid;
    GMAJ: THGrid;
    BFLECHEDROITE: TToolbarButton97;
    BFLECHEGAUCHE: TToolbarButton97;
    TLISTEET: THLabel;
    TMAJ: THLabel;
    TabSheet3: TTabSheet;
    HMsgErr: THMsgBox;
    Bevel1: TBevel;
    GBDe: TGroupBox;
    GBA: TGroupBox;
    GBPrix: TGroupBox;
    RBDPA: TRadioButton;
    RBPA: TRadioButton;
    RBPVTTC: TRadioButton;
    ETABBASE: THValComboBox;
    TETABBASE: THLabel;
    GBTARIFMAJ: TGroupBox;
    TTYPETARIFDEST: THLabel;
    TPERIODEDEST: THLabel;
    TYPETARIFDEST: THValComboBox;
    PERIODEDEST: THValComboBox;
    COEF: THNumEdit;
    TCOEF: THLabel;
    Bevel2: TBevel;
    Bevel3: TBevel;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    ListRecap: TListBox;
    HRecap: THMsgBox;
    BFinRecap: TButton;
    BStop: TButton;
    RBPOURC: TRadioButton;
    BFELCHETOUS: TToolbarButton97;
    BFLECHEAUCUN: TToolbarButton97;
    RBPVART: TRadioButton;
    ARRONDIC: THValComboBox;
    TARRONDI: THLabel;
    RBPXSAISI: TRadioButton;
    TPXSAISI: THLabel;
    PXSAISI: THNumEdit;
    RBAHT: TRadioButton;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);

    // Onglet 1
    procedure GBPrixOnClick(Sender: TObject);
    procedure ComboOnChange(Sender: TObject);

    // Onglet 2
    procedure ClickFlecheDroite(Sender: TObject);
    procedure ClickFlecheGauche(Sender: TObject);
    procedure ClickFlecheTous(Sender: TObject);
    procedure ClickFlecheAucun(Sender: TObject);

    // Onglet 3
    procedure ClickBFinRecap(Sender: TObject);

    // Arret traitement
    procedure ClickBStop(Sender: TObject);
    procedure FormDestroy(Sender: TObject);

  private
    TobListeDepot, TobListeDepotMaj, TOBArticle, TobTarif, TobPrix, TobArtTraite,
      TobAnomalie, TobReussite, TOBMode, TOBDev: TOB;
    Statut, InfoGen, ArticleSansTarif: string;
    StopT, terminer: Boolean;
    CoefType: Double;
    MaxTarif: Integer;
    NatureType: string; // Tarif ACH ou VTE
    TarifASup: string; // Supprimer dimension non MAJ

    // Onglet 1
    procedure Initcombo;

    // Onglet 2
    procedure RemplirGridEtablissement;
    procedure RefreshGrid(posListe, posMaj: integer);
    procedure RefreshBouton;

    //Onglet 3
    function SpaceStr(nb: integer): string;
    function ExtractLibelle(St: string): string;
    procedure ListeRecap;
    procedure ListeRecap100Art(NbArtSelect: Integer);
    procedure ListeDepotTrait(Depot: string);
    procedure ListeRecapTermine;
    procedure RemplirTobRecap(TOBTarif: TOB; Statut: string);
    procedure Anomalie(TOBD: TOB; Depot: string; i_ind: Integer);

    // Traitement
    procedure TraiterTarif;
    procedure CouperTobArt; // Traite NbArtParTraitement Articles/NbArtParTraitement Articles
    function TraiterArticleDim: Tob; // pour les prix recup Prix Dim
    function CreerTobPrix(Depot: string): boolean;
    function MAJDepot(Depot: string): Boolean;
    procedure CreerTobTarif(Depot: string);
    procedure CreerTarifGen;
    procedure NumTarif;
    procedure MajTarifArticle(Depot: string);
    procedure MajPeriode(TOBT: TOB);
    procedure NouveauTarif(TobArt: Tob; Depot: string);
    procedure CreerTarifDim(TobTarfGen: TOB);
    procedure TraiterNvDimension(CodeArticle, Depot: string; TarifMode: integer);
    function PrixDifferent(TobPxDim: Tob; PxGen, PxAchGen: Double): Boolean;
    procedure RecupInfoTypeEtPeriode;
    procedure RecupInfoPeriodeEtab(Depot: string);
    procedure LibereDepot;
    procedure LibereTout;
    procedure MAJDateModif;

  public
    { Déclarations publiques }
  end;

const NbArtParRequete: integer = 50;
  NbArtParTraitement: integer = 100;

var
  FMajTarfMode: TFMajTarfMode;
  i_NumEcran: Integer;

function RequeteTarif(Article, Depot: string; TarifMode: Integer; Dim: Boolean = False): string;
function RechLibArt(TOBArt: TOB): string;
procedure NoteEvenement(Mess: TStringList; Etat: string);
function CalculCoefTypeTarif(TypeTarifRef, TypeTarifDest, NatureType: string; TobMode: Tob): Double;
procedure SuppressionDimension(TarifASup: string);

implementation

{$R *.DFM}

procedure MiseAJourTarifMode;
var
  FMajTarfMode: TFMajTarfMode;
begin
  FMajTarfMode := TFMajTarfMode.Create(Application);
  FMajTarfMode.TOBArticle := TOBA;
  FMajTarfMode.Statut := STATUT;
  FMajTarfMode.NatureType := NATURETYPE;
  try
    FMajTarfMode.ShowModal;
  finally
    FMajTarfMode.free;
  end;
end;

{=========================================================================================}
{============================= Evenements de la forme ====================================}
{=========================================================================================}

procedure TFMajTarfMode.FormShow(Sender: TObject);
var Onglet: TTabSheet;
  St_NomOnglet: string;
begin
  inherited;
  bAnnuler.Visible := True;
  bSuivant.Enabled := False;
  bFin.Visible := True;
  bFin.Enabled := False;
  StopT := False;
  i_NumEcran := 0;
  Onglet := P.ActivePage;
  st_NomOnglet := Onglet.Name;
  i_NumEcran := strtoint(Copy(st_NomOnglet, length(st_NomOnglet), 1)) - 1;
  AglEmpileFiche(Self);
  // Affichage des informations du tarif ACH ou VTE
  if NatureType = 'VTE' then
  begin
    TYPETARIF.DataType := 'GCTARIFTYPE1VTE';
    TYPETARIFDEST.DataType := 'GCTARIFTYPE1VTE';
    PTITRE.Caption := 'Mise à jour des tarifs de vente';
  end else
  begin
    TYPETARIF.DataType := 'GCTARIFTYPE1ACH';
    TYPETARIFDEST.DataType := 'GCTARIFTYPE1ACH';
    PTITRE.Caption := 'Mise à jour des tarifs d''achat';
  end;
  UpdateCaption(Self);
end;

procedure TFMajTarfMode.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  LibereTout;
end;

procedure TFMajTarfMode.FormDestroy(Sender: TObject);
begin
  inherited;
  AglDepileFiche;
end;

procedure TFMajTarfMode.bSuivantClick(Sender: TObject);
var Onglet: TTabSheet;
  St_NomOnglet: string;
begin
  inherited;
  Onglet := P.ActivePage;
  st_NomOnglet := Onglet.Name;
  i_NumEcran := strtoint(Copy(st_NomOnglet, length(st_NomOnglet), 1)) - 1;
  if (i_NumEcran = 2) and (TobListeDepotMaj.Detail.Count = 0) and (TYPETARIFDEST.Value <> '...') then
  begin
    RestorePage;
    PGIBox(HMsgErr.Mess[0], Caption);
    Onglet := PreviousPage;
    if Onglet = nil then P.SelectNextPage(True) else
    begin
      P.ActivePage := Onglet;
      PChange(nil);
    end;
  end;
  if (i_NumEcran = 1) and (TYPETARIFDEST.Value = '...') then
  begin
    RestorePage;
    Onglet := NextPage;
    if Onglet = nil then P.SelectNextPage(True) else
    begin
      P.ActivePage := Onglet;
      PChange(nil);
    end;
  end;
  if (i_NumEcran = 1) then
  begin
    COEF.Value := StrToFloat(COEF.Text);
    PXSAISI.Value := StrToFloat(PXSAISI.Text);
  end;
  if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
  if bFin.Enabled then ListeRecap;
  terminer := False;
end;

procedure TFMajTarfMode.bPrecedentClick(Sender: TObject);
var Onglet: TTabSheet;
  St_NomOnglet: string;
begin
  inherited;
  Onglet := P.ActivePage;
  st_NomOnglet := Onglet.Name;
  i_NumEcran := strtoint(Copy(st_NomOnglet, length(st_NomOnglet), 1)) - 1;
  if (i_NumEcran = 1) and (TYPETARIFDEST.Value = '...') then
  begin
    RestorePage;
    Onglet := PreviousPage;
    if Onglet = nil then P.SelectNextPage(True) else
    begin
      P.ActivePage := Onglet;
      PChange(nil);
    end;
  end;
  if (bSuivant.Enabled) then bFin.Enabled := False else bFin.Enabled := True;
end;

procedure TFMajTarfMode.bFinClick(Sender: TObject);
begin
  inherited;
  if not terminer then
  begin
    bPrecedent.Enabled := False;
    bFin.Visible := False;
    TraiterTarif;
  end else
  begin
    LibereTout;
    close;
  end;
end;

{=========================================================================================}
{============================= Evenements Onglet 1 =======================================}
{=========================================================================================}

procedure TFMajTarfMode.GBPrixOnClick(Sender: TObject);
begin
  if (RBPVTTC.Checked) or (RBAHT.Checked) then // Prix de vente TTC tarif ou Prix de achat HT
  begin
    TTYPETARIF.Enabled := True;
    TYPETARIF.Enabled := True;
    TPERIODE.Enabled := True;
    PERIODE.Enabled := True;
    TETABBASE.Enabled := True;
    ETABBASE.Enabled := True;
    bSuivant.Enabled := False;
    TCOEF.Caption := TraduireMemoire('Coefficient à appliquer');
    TARRONDI.Enabled := True;
    ARRONDIC.Enabled := True;
    TPXSAISI.Enabled := False;
    PXSAISI.Enabled := False;
    COEF.Enabled := True;
    TCOEF.Enabled := True;
    InitCombo;
    if (RBPVTTC.Checked) then TYPETARIF.DataType := 'GCTARIFTYPE1VTE'
    else TYPETARIF.DataType := 'GCTARIFTYPE1ACH';
  end
  else
    if (RBDPA.Checked) then //  dernier prix d'achat
  begin
    TTYPETARIF.Enabled := False;
    TYPETARIF.Enabled := False;
    TPERIODE.Enabled := False;
    PERIODE.Enabled := False;
    TETABBASE.Enabled := True;
    ETABBASE.Enabled := True;
    bSuivant.Enabled := False;
    TCOEF.Caption := TraduireMemoire('Coefficient à appliquer');
    TARRONDI.Enabled := True;
    ARRONDIC.Enabled := True;
    TPXSAISI.Enabled := False;
    PXSAISI.Enabled := False;
    COEF.Enabled := True;
    TCOEF.Enabled := True;
    InitCombo;
  end else
    if (RBPA.Checked) or (RBPVART.Checked) then //Prix d'achat, Prix de vente TTC article
  begin
    TTYPETARIF.Enabled := False;
    TYPETARIF.Enabled := False;
    TPERIODE.Enabled := False;
    PERIODE.Enabled := False;
    TETABBASE.Enabled := False;
    ETABBASE.Enabled := False;
    TCOEF.Caption := TraduireMemoire('Coefficient à appliquer');
    TARRONDI.Enabled := True;
    ARRONDIC.Enabled := True;
    TPXSAISI.Enabled := False;
    PXSAISI.Enabled := False;
    COEF.Enabled := True;
    TCOEF.Enabled := True;
    InitCombo;
  end else
    if RBPOURC.checked then // tarif en %
  begin
    TTYPETARIF.Enabled := False;
    TYPETARIF.Enabled := False;
    TPERIODE.Enabled := False;
    PERIODE.Enabled := False;
    TETABBASE.Enabled := False;
    ETABBASE.Enabled := False;
    InitCombo;
    TCOEF.Caption := TraduireMemoire('Pourcentage de remise');
    TARRONDI.Enabled := False;
    ARRONDIC.Enabled := False;
    TPXSAISI.Enabled := False;
    PXSAISI.Enabled := False;
    COEF.Enabled := True;
    TCOEF.Enabled := True;
  end else // Prix saisi
  begin
    TTYPETARIF.Enabled := False;
    TYPETARIF.Enabled := False;
    TPERIODE.Enabled := False;
    PERIODE.Enabled := False;
    TETABBASE.Enabled := False;
    ETABBASE.Enabled := False;
    COEF.Enabled := False;
    TCOEF.Enabled := False;
    TARRONDI.Enabled := False;
    ARRONDIC.Enabled := False;
    TPXSAISI.Enabled := True;
    PXSAISI.Enabled := True;
    InitCombo;
  end;
end;

procedure TFMajTarfMode.ComboOnChange(Sender: TObject);
begin
  RemplirGridEtablissement;
  RefreshBouton;
  if TYPETARIF.Enabled then
  begin
    if TYPETARIF.Value = '...' then
    begin
      TETABBASE.Enabled := False;
      ETABBASE.Enabled := False;
    end else
    begin
      TETABBASE.Enabled := True;
      ETABBASE.Enabled := True;
      if RBPVTTC.Checked then ETABBASE.Plus := '(ET_TYPETARIF = "' + TYPETARIF.Value + '")'
      else
        if (RBAHT.Checked) then ETABBASE.Plus := '(ET_TYPETARIFACH = "' + TYPETARIF.Value + '")'
    end;
    bSuivant.Enabled := (TYPETARIF.Value <> '') and (PERIODE.Value <> '') and (TYPETARIFDEST.Value <> '') and (PERIODEDEST.Value <> '');
  end else
  begin
    bSuivant.Enabled := (TYPETARIFDEST.Value <> '') and (PERIODEDEST.Value <> '');
  end;
  if (TYPETARIFDEST.Value = '...') and (TobListeDepotMaj.Detail.count = 0) then
  begin
    TOB.Create('', TobListeDepotMaj, -1);
    TobListeDepotMaj.Detail[0].AddChampSup('ET_ETABLISSEMENT', False);
    TobListeDepotMaj.Detail[0].PutValue('ET_ETABLISSEMENT', '');
    TobListeDepotMaj.Detail[0].AddChampSup('ET_ABREGE', False);
    TobListeDepotMaj.Detail[0].PutValue('ET_ABREGE', TraduireMemoire('Aucun établissement'));
  end else
    if (TYPETARIFDEST.Value <> '...') and (TobListeDepotMaj.Detail.count > 0) then TobListeDepotMaj.ClearDetail;
  if (ExisteSQL('Select GFP_PROMO from TarifPer where GFP_CODEPERIODE="' + PERIODEDEST.Value + '" and GFP_PROMO="-"'))
    and (RBPOURC.Checked = true) then
  begin
    PGIBox('Pour une période hors promo, les tarifs sont exprimés en montant uniquement', Caption);
    RBPOURC.Checked := False;
    RBPXSAISI.Checked := True;
    GBPrixOnClick(nil);
  end;
end;

procedure TFMajTarfMode.Initcombo;
begin
  TYPETARIF.Value := '';
  PERIODE.Value := '';
  ETABBASE.Value := '';
end;

{=========================================================================================}
{============= Evenements de l'onglet 2 -consernant les établissement ====================}
{=========================================================================================}

procedure TFMajTarfMode.RemplirGridEtablissement;
var QQ: TQuery;
  j: integer;
begin
  if TobListeDepot = nil then TobListeDepot := TOB.CREATE('Liste établissements', nil, -1);
  if TobListeDepotMaj = nil then TobListeDepotMaj := TOB.CREATE('Etablissements maj', nil, -1);
  if NatureType = 'VTE' then
    QQ := OpenSQL('select ET_ETABLISSEMENT,ET_ABREGE from ETABLISS where ET_TYPETARIF="' + TYPETARIFDEST.VALUE + '" order by ET_ETABLISSEMENT', True)
  else QQ := OpenSQL('select ET_ETABLISSEMENT,ET_ABREGE from ETABLISS where ET_TYPETARIFACH="' + TYPETARIFDEST.VALUE + '" order by ET_ETABLISSEMENT', True);
  if not QQ.EOF then TobListeDepot.LoadDetailDB('ETABLISS', '', '', QQ, False);
  j := TobListeDepot.Detail.count;
  if TobListeDepot.FindFirst(['ET_ETABLISSEMENT'], [''], True) = nil then
  begin
    TOB.Create('ETABLISS', TobListeDepot, j);
    TobListeDepot.Detail[j].PutValue('ET_ETABLISSEMENT', '');
    TobListeDepot.Detail[j].PutValue('ET_ABREGE', TraduireMemoire('Toutes boutiques'));
  end;
  TobListeDepot.PutGridDetail(GLISTE, False, False, 'ET_ETABLISSEMENT;ET_ABREGE', True);
  GLISTE.ColWidths[0] := 30;
  GLISTE.ColAligns[0] := taCenter;
  GLISTE.ColWidths[1] := 195;
  GLISTE.ColAligns[1] := taLeftJustify;
  GMAJ.ColWidths[0] := 30;
  GMAJ.ColAligns[0] := taCenter;
  GMAJ.ColWidths[1] := 195;
  GMAJ.ColAligns[1] := taLeftJustify;
  GMAJ.RowCount := 0;
  Ferme(QQ);
end;

procedure TFMajTarfMode.ClickFlecheDroite;
var indiceFille: integer;
begin
  // Y a t il quelque chose de sélectionné ?
  if TobListeDepot.Detail.Count = 0 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if TobListeDepotMaj.Detail.Count > 0 then indiceFille := GMAJ.Row + 1 else indiceFille := 0;
  TobListeDepot.detail[GLISTE.Row].ChangeParent(TobListeDepotMaj, indiceFille);
  RefreshGrid(GLISTE.Row, GMAJ.Row + 1);
end;

procedure TFMajTarfMode.ClickFlecheTous;
var indiceFille, iGrd, posListe: integer;
begin
  if TobListeDepot.Detail.Count = 0 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if TobListeDepotMaj.Detail.Count > 0 then indiceFille := GMAJ.Row + 1 else indiceFille := 0;
  posListe := TobListeDepot.detail.count - 1;
  for iGrd := 0 to posListe do TobListeDepot.detail[0].ChangeParent(TobListeDepotMaj, indiceFille + iGrd);
  RefreshGrid(0, indiceFille + posListe);
end;

procedure TFMajTarfMode.ClickFlecheGauche;
var indiceFille: integer;
begin
  // Y a t il quelque chose de sélectionné ?
  if TobListeDepotMaj.Detail.Count = 0 then exit;
  // Changement du parent de l'élément des établissements affichés
  if TobListeDepot.Detail.Count > 0 then indiceFille := GLISTE.Row + 1 else indiceFille := 0;
  TobListeDepotMaj.detail[GMAJ.Row].ChangeParent(TobListeDepot, indiceFille);
  RefreshGrid(GLISTE.Row + 1, GMAJ.Row);
end;

procedure TFMajTarfMode.ClickFlecheAucun;
var indiceFille, iGrd, posMAJ: integer;
begin
  if TobListeDepotMaj.Detail.Count = 0 then exit;
  // Changement du parent de l'élément de la liste des établissements
  if TobListeDepot.Detail.Count > 0 then indiceFille := GLISTE.Row + 1 else indiceFille := 0;
  posMAJ := TobListeDepotMaj.detail.count - 1;
  for iGrd := 0 to posMAJ do TobListeDepotMaj.detail[0].ChangeParent(TobListeDepot, indiceFille + iGrd);
  RefreshGrid(indiceFille + posMAJ, 0);
end;

procedure TFMajTarfMode.RefreshGrid(posListe, posMaj: integer);
begin
  TobListeDepotMaj.PutGridDetail(GMAJ, False, False, 'ET_ETABLISSEMENT;ET_ABREGE', True);
  TobListeDepot.PutGridDetail(GLISTE, False, False, 'ET_ETABLISSEMENT;ET_ABREGE', True);
  GMAJ.Row := Min(posMaj, GMAJ.RowCount - 1);
  GLISTE.Row := Min(posListe, GLISTE.RowCount - 1);
  RefreshBouton;
end;

procedure TFMajTarfMode.RefreshBouton;
begin
  BFLECHEDROITE.Enabled := TobListeDepot.Detail.Count > 0;
  BFLECHEGAUCHE.Enabled := TobListeDepotMaj.Detail.Count > 0;
end;

{=========================================================================================}
{================================= Onglet 3 Récapitulatif ================================}
{=========================================================================================}

function TFMajTarfMode.SpaceStr(nb: integer): string;
var St_Chaine: string;
  i_ind: integer;
begin
  St_Chaine := '';
  for i_ind := 1 to nb do St_Chaine := St_chaine + ' ';
  Result := St_Chaine;
end;

function TFMajTarfMode.ExtractLibelle(St: string): string;
var St_Chaine: string;
  i_pos: integer;
begin
  Result := '';
  i_pos := Pos('&', St);
  if i_pos > 0 then
  begin
    St_Chaine := Copy(St, 1, i_pos - 1) + Copy(St, i_pos + 1, Length(St));
  end else St_Chaine := St;
  Result := St_Chaine + ' : ';
end;

procedure TFMajTarfMode.ListeRecap;
var st_chaine: string;
  i: Integer;
begin
  ListRecap.Items.Clear;
  ListRecap.Items.Add(PTITRE.Caption);
  ListRecap.Items.Add('');
  //Info concerant les tarifs articles à mettre à jour
  ListRecap.Items.Add('Nombres d''articles séléctionnés: ' + IntToStr(TobArticle.Detail.count) + '');
  //Info concerant le premier onglet
  ListRecap.Items.Add(ExtractLibelle(GBDe.Caption));
  if RBDPA.checked then St_Chaine := RBDPA.Caption else
    if RBPA.checked then st_chaine := RBPA.Caption else
    if RBPVTTC.checked then st_chaine := RBPVTTC.Caption else
    if RBAHT.checked then st_chaine := RBAHT.Caption else
    if RBPVART.checked then st_chaine := RBPVART.Caption else
    if RBPOURC.Checked then st_chaine := RBPOURC.Caption else
    if RBPXSAISI.Checked then st_chaine := RBPXSAISI.Caption;

  ListRecap.Items.Add(SpaceStr(4) + st_chaine);
  if TYPETARIF.Value <> '' then
    ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TTYPETARIF.Caption) + TYPETARIF.Text);
  if PERIODE.Value <> '' then
    ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TPERIODE.Caption) + PERIODE.Text);
  if ETABBASE.Value <> '' then
    ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TETABBASE.Caption) + ETABBASE.Text);
  ListRecap.Items.Add(ExtractLibelle(GBA.Caption));
  if TYPETARIFDEST.Value <> '' then
    ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TTYPETARIFDEST.Caption) + TYPETARIFDEST.Text);
  if PERIODEDEST.Value <> '' then
    ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TPERIODEDEST.Caption) + PERIODEDEST.Text);

  if RBPXSAISI.Checked then ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TPXSAISI.Caption) + FloatToStr(PXSAISI.Value)) else
    ListRecap.Items.Add(SpaceStr(4) + ExtractLibelle(TCOEF.Caption) + FloatToStr(COEF.Value));

  ListRecap.Items.Add('');
  //Info concerant le deuxième onglet
  if TobListeDepotMaj.Detail.Count > 0 then
  begin
    ListRecap.Items.Add(ExtractLibelle(TMAJ.Caption));
    for i := 0 to TobListeDepotMaj.Detail.count - 1 do
    begin
      ListRecap.Items.Add(SpaceStr(4) + TobListeDepotMaj.Detail[i].GetValue('ET_ABREGE'));
    end;
  end;
  ListRecap.Items.Add('');
end;

procedure TFMajTarfMode.ListeRecap100Art(NbArtSelect: Integer);
var CompteurDeb, CompteurFin: Integer;
begin
  ListRecap.Items.Clear;
  ListRecap.Items.Add(PTITRE.Caption);
  ListRecap.Items.Add('');
  //Info concerant les tarifs articles à mettre à jour
  CompteurDeb := NbArtSelect - TobArticle.Detail.count - 99;
  CompteurFin := CompteurDeb + TobArtTraite.Detail.count;
  if TobArticle.Detail.count = 0 then
  begin
    CompteurDeb := NbArtSelect - (TobArtTraite.Detail.count - 1);
    CompteurFin := NbArtSelect;
  end;
  ListRecap.ITems.Add('Traitement des articles de ' + IntToStr(compteurDeb) + ' à ' + IntToStr(CompteurFin) + ' sur ' + IntToStr(NbArtSelect) + '');
end;

procedure TFMajTarfMode.ListeDepotTrait(Depot: string);
var Libelle: string;
begin
  Libelle := RechDom('TTETABLISSEMENT', Depot, False);
  if Libelle = 'Error' then Libelle := TraduireMemoire('Toutes boutiques');
  ListRecap.Items.Add(SpaceStr(4) + 'Traitement de l''établissement ' + Libelle + '');
end;

procedure TFMajTarfMode.ListeRecapTermine;
var i: Integer;
begin
  ListRecap.Items.Clear;
  ListRecap.Items.Add(TraduireMemoire('La mise à jour des tarifs est terminée.'));
  ListRecap.Items.Add('');
  ListRecap.Items.Add('');
  // Tarif non mise à jour
  if TobAnomalie.Detail.count > 0 then
  begin
    ListRecap.Items.Add(HRecap.Mess[7]);
    ListRecap.Items.Add(SpaceStr(4) + 'DEPOT' + SpaceStr(20) + 'ARTICLE');
    for i := 0 to TobAnomalie.Detail.count - 1 do
    begin
      if TobAnomalie.Detail[i].Getvalue('DEPOT') = '' then
        ListRecap.Items.Add(SpaceStr(4) + '' + SpaceStr(20) + TobAnomalie.Detail[i].Getvalue('ARTICLE')) else
        ListRecap.Items.Add(SpaceStr(4) + RechDom('TTETABLISSEMENT', TobAnomalie.Detail[i].Getvalue('DEPOT'), False) + SpaceStr(10) +
          TobAnomalie.Detail[i].Getvalue('ARTICLE'));
    end;
  end;
  BFin.Visible := True;
  BFin.Caption := 'Quitter';
  terminer := True;
end;

function RechLibArt(TOBArt: TOB): string;
var k: Integer;
  Grille, CodeDim, LibDim, Libelle: string;
begin
  Result := '';
  if (TOBArt.GetValue('GA_STATUTART') = 'GEN') or (TOBArt.GetValue('GA_STATUTART') = 'UNI') then
    Result := RechDom('GCARTICLEGENERIQUE', TOBArt.GetValue('GA_CODEARTICLE'), False)
  else
    if TOBArt.GetValue('GA_STATUTART') = 'DIM' then
  begin
    for k := 1 to 5 do
    begin
      if VarIsNull(TOBArt.GetValue('GA_GRILLEDIM' + IntToStr(k))) then Grille := '' else
        Grille := TOBArt.GetValue('GA_GRILLEDIM' + IntToStr(k));
      if VarIsNull(TOBArt.GetValue('GA_CODEDIM' + IntToStr(k))) then CodeDim := '' else
        CodeDim := TOBArt.GetValue('GA_CODEDIM' + IntToStr(k));
      if ((Grille <> '') and (CodeDim <> '')) then
      begin
        LibDim := RechDom('GCGRILLEDIM' + IntToStr(k), Grille, True) + ' ' + GCGetCodeDim(Grille, CodeDim, k);
        Libelle := Libelle + '  ' + LibDim;
      end;
    end;
    Result := Libelle;
  end;
end;

procedure TFMajTarfMode.ClickBFinRecap(Sender: TObject);
begin
  LibereTout;
  close;
end;

procedure TFMajTarfMode.LibereTout;
begin
  TobListeDepot.Free;
  TobListeDepot := nil;
  TobListeDepotMaj.Free;
  TobListeDepotMaj := nil;
  TobArtTraite.Free;
  TobArtTraite := nil;
  TobDev.Free;
  TobDev := nil;
  TOBMode.Free;
  TOBMode := nil;
  TobAnomalie.Free;
  TobAnomalie := nil;
  TobReussite.Free;
  TobReussite := nil;
end;

procedure TFMajTarfMode.ClickBStop(Sender: TObject);
begin
  StopT := True;
end;

{=========================================================================================}
{================================== Traitement ===========================================}
{=========================================================================================}

procedure NoteEvenement(Mess: TStringList; Etat: string);
var TobJNL: TOB;
  Indice: integer;
  QIndice: TQuery;
begin
  Indice := 0;
  QIndice := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true);
  if not QIndice.Eof then
    Indice := QIndice.Fields[0].AsInteger + 1;
  Ferme(QIndice);

  TobJNL := TOB.Create('JNALEVENT', nil, -1);
  TobJNL.PutValue('GEV_NUMEVENT', Indice);
  TobJNL.PutValue('GEV_TYPEEVENT', 'TAR');
  TobJNL.PutValue('GEV_LIBELLE', 'Mise à jour des tarifs');
  TobJNL.PutValue('GEV_DATEEVENT', Date);
  TobJNL.PutValue('GEV_UTILISATEUR', V_PGI.USer);
  TobJNL.PutValue('GEV_ETATEVENT', Etat);
  TobJNL.PutValue('GEV_BLOCNOTE', Mess.Text);
  TobJNL.InsertOrUpdateDB(false);
  TobJNL.Free;
end;

procedure TFMajTarfMode.TraiterTarif;
var Erreur: boolean;
  Mess: TStringList;
  Depot, Evenement: string;
  i, NbArtSelec: Integer;
  Q: TQuery;
begin
  BStop.Visible := True;
  Erreur := false;
  TobAnomalie := TOB.Create('', nil, -1);
  TobReussite := TOB.Create('', nil, -1);
  TobDev := TOB.Create('', nil, -1);
  Q := OpenSql('Select * from devise', True);
  TobDev.LoadDetailDB('DEVISE', '', '', Q, False);
  ferme(Q);
  MaxTarif := -1;
  NbArtSelec := TOBArticle.Detail.count;
  Mess := TStringList.Create;
  RecupInfoTypeEtPeriode;
  while TOBArticle.Detail.count > 0 do
  begin
    CouperTobArt;
    ListeRecap100Art(NbArtSelec);
    for i := 0 to TobListeDepotMaj.Detail.Count - 1 do
    begin
      Depot := TobListeDepotMaj.Detail[i].GetValue('ET_ETABLISSEMENT');
      ListeDepotTrait(Depot);
      if not CreerTobPrix(Depot) then
      begin
        if TOBPrix.Detail.Count = 0 then
        begin
          Mess.Add('Erreur - Aucun tarif ne correspond à votre sélection, la Mise à jour des tarifs non effectué ');
          Evenement := 'ERR';
          Break;
        end else
        begin
          Mess.Add('Erreur - Certain tarif ne correspond à votre sélection, la Mise à jour des tarifs n''est pas complete ');
          Mess.Add('Liste article sans tarif - ' + ArticleSansTarif + '');
          Erreur := True;
          Evenement := 'ERR';
        end;
      end;
      try
        BeginTrans;
        if MAJDepot(Depot) then CommitTrans
        else raise ERangeError.create('');
      except
        Erreur := true;
        RollBack;
      end;
      if StopT then
      begin
        if Depot = '' then Mess.Add('Traitement interrompu par l''utilisateur à partir du tarif "Toutes boutiques"')
        else Mess.Add('Traitement interrompu par l''utilisateur à partir du depot ' + RechDom('TTETABLISSEMENT', Depot, False) + '');
        Evenement := 'INT';
        break;
      end else
        if Erreur then
      begin
        if Depot = '' then Mess.Add('Erreur - Mise à jour des tarifs "Toutes boutiques" non effectuée')
        else Mess.Add('Erreur - Mise à jour des tarifs du depot ' + RechDom('TTETABLISSEMENT', Depot, False) + ' non effectuée');
        Evenement := 'ERR';
      end else
      begin
        if Depot = '' then Mess.Add('La Mise à jour des tarifs "Toutes boutiques" s''est bien effectuée')
        else Mess.Add('La Mise à jour des tarifs du depot ' + RechDom('TTETABLISSEMENT', Depot, False) + ' s''est bien effectuée');
        Evenement := 'OK';
      end;
    end;
  end;
  NoteEvenement(Mess, Evenement);
  Mess.Free;
  ListeRecapTermine;
end;

// Traitement par tranche de NbArtParTraitement (Variable locale)

procedure TFMajTarfMode.CouperTobArt;
var i, Compteur: Integer;
begin
  if TOBArticle.Detail.count > 100 then Compteur := NbArtParTraitement - 1
  else Compteur := TOBArticle.Detail.count - 1;
  if TobArtTraite <> nil then
  begin
    TobArtTraite.Free;
    TobArtTraite := nil;
  end;
  TobArtTraite := TOB.Create('ARTICLE', nil, -1);
  for i := Compteur downto 0 do
  begin
    TOBArticle.Detail[i].ChangeParent(TobArtTraite, -1);
  end;
end;

function TFMajTarfMode.CreerTobPrix(Depot: string): boolean;
var i, y, TarifMode, CountStArt, NbRequete: Integer;
  QPrix: TQuery;
  SQLPrix, DepotApplic: string;
  TabWhere: array of string;
  TobArtPrix: Tob;
begin
  Result := True;
  CountStArt := 0;
  NbRequete := 0;
  QPrix := nil;
  SetLength(TabWhere, 1);
  if (ETABBASE.Value <> '') or (TYPETARIF.Value <> '') then DepotApplic := ETABBASE.Value else
    if TYPETARIF.Value = '...' then DepotApplic := '' else DepotApplic := Depot;
  TOBPrix := TOB.Create('LesPrix', nil, -1);
  //if Statut='GEN' then
  //  begin
  TobArtPrix := TraiterArticleDim;
  //  end ;
  // Création d'un tableau de where sur l'article
  if TobArtPrix <> nil then
  begin
    for i := 0 to TobArtPrix.Detail.Count - 1 do
    begin
      if CountStArt >= NbArtParRequete then
      begin
        NbRequete := NbRequete + 1;
        SetLength(TabWhere, NbRequete + 1);
        CountStArt := 0;
      end;
      CountStArt := CountStArt + 1;
      if TabWhere[NbRequete] = '' then TabWhere[NbRequete] := '"' + Trim(TobArtPrix.Detail[i].GetValue('GA_ARTICLE')) + '"'
      else TabWhere[NbRequete] := TabWhere[NbRequete] + ',"' + Trim(TobArtPrix.Detail[i].GetValue('GA_ARTICLE')) + '"';
    end;
  end;
  // Execute 1 requete par tableau ;
  if TabWhere[0] <> '' then
  begin
    for y := Low(TabWhere) to High(TabWhere) do
    begin
      if TabWhere[y] <> '' then
      begin
        if (RBPA.Checked) or (RBPVART.Checked) then
        begin
          SQLPrix := 'Select GA_ARTICLE, GA_PAHT,GA_PVTTC, GCA_PRIXBASE, GA_PRIXUNIQUE, GA_STATUTART From ARTICLE left outer join CATALOGU ON GA_ARTICLE=GCA_ARTICLE '
            +
            'where GA_ARTICLE IN (' + Trim(TabWhere[y]) + ') Order by GA_ARTICLE';
          QPrix := OpenSQL(SQLPrix, True);
          if not QPrix.Eof then TOBPrix.LoadDetailDB('_Prix', '', '', QPrix, True, False);
        end else
          if (RBDPA.Checked) or (RBPOURC.Checked) then
        begin
          SQLPrix := 'Select GA_ARTICLE, GA_PAHT, GQ_DPA, GA_PRIXUNIQUE, GA_STATUTART From ARTICLE left outer join DISPO ON GA_ARTICLE=GQ_ARTICLE ' +
            'and GQ_DEPOT="' + DepotApplic + '" ' +
            'where GA_ARTICLE IN (' + Trim(TabWhere[y]) + ') Order by GA_ARTICLE ';
          QPrix := OpenSQL(SQLPrix, True);
          if not QPrix.Eof then TOBPrix.LoadDetailDB('_Prix', '', '', QPrix, True, False);
        end else
          if (RBPVTTC.Checked) or (RBAHT.Checked) then
        begin
          if RBPVTTC.Checked then TarifMode := RecupCodeTarifMode(TYPETARIF.Value, PERIODE.Value, 'VTE')
          else TarifMode := RecupCodeTarifMode(TYPETARIF.Value, PERIODE.Value, 'ACH');
          SQLPrix :=
            'Select GA_ARTICLE,GA_PAHT, GF_PRIXUNITAIRE, GA_PRIXUNIQUE,GA_STATUTART,GF_CALCULREMISE,GF_REMISE from ARTICLE left outer join TARIF ON GA_ARTICLE=GF_ARTICLE ';
          SQLPrix := SQLPrix + 'where GA_ARTICLE IN (' + Trim(TabWhere[y]) + ') AND GF_DEPOT="' + DepotApplic + '"' +
            ' AND GF_TARFMODE="' + IntToStr(TarifMode) + '" AND GF_FERME="-"';
          SQLPrix := SQLPrix + ' ORDER BY GA_ARTICLE';
          QPrix := OpenSQL(SQLPrix, True);
          if not QPrix.Eof then TOBPrix.LoadDetailDB('_Prix', '', '', QPrix, True, False)
          else
          begin
            Result := False;
            ArticleSansTarif := ArticleSansTarif + Trim(TabWhere[y]);
          end; //PGIBox (HMsgErr.Mess[1],Caption) ;
        end;
        ferme(QPrix);
      end;
    end;
  end;
  TobArtPrix.Free;
end;

function TFMajTarfMode.TraiterArticleDim: Tob;
var i, j, Tarifmode: Integer;
  Q: TQuery;
  TobArt, TobTravail, TobDim: TOB;
  SQL, SQL2: string;
begin
  //if tobtravail<>nil then begin TobTravail.Free; end ;
  j := 0;
  TobTravail := TOB.Create('_ARTICLE', nil, -1);
  TobTravail.Dupliquer(TobArtTraite, True, False);
  for i := 0 to TobArtTraite.Detail.Count - 1 do
  begin
    TobArt := TobTravail.FindFirst(['GA_ARTICLE'], [TobArtTraite.Detail[i].GetValue('GA_ARTICLE')], False);
    if TobArt <> nil then j := TobArt.GetIndex + 1;
    if (TobArt.GetValue('GA_STATUTART') = 'GEN') and (TobArt.Getvalue('GA_PRIXUNIQUE') = '-') then
    begin
      SQL2 := 'Select GA_ARTICLE,GA_CODEARTICLE,GA_STATUTART,GA_PRIXUNIQUE,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5' +
        ',GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5 from ARTICLE where GA_CODEARTICLE="' + TobArt.GetValue('GA_CODEARTICLE') +
        '" and GA_STATUTART="DIM"';
      Q := OpenSQL(SQL2, True);
      while not Q.EOF do
      begin
        Tobdim := TobTravail.FindFirst(['GA_ARTICLE'], [Q.Findfield('GA_ARTICLE').AsString], False);
        if TobDim <> nil then break;
        //TarifMode:=RecupCodeTarifMode(TYPETARIFDEST.Value,PERIODEDEST.Value,NatureType) ;
        if RBPVTTC.Checked then TarifMode := RecupCodeTarifMode(TYPETARIF.Value, PERIODE.Value, 'VTE')
        else TarifMode := RecupCodeTarifMode(TYPETARIF.Value, PERIODE.Value, 'ACH');
        SQL := 'Select GF_TARIF from Tarif Where GF_ARTICLE="' + Q.Findfield('GA_ARTICLE').AsString + '" AND GF_TARFMODE="' + IntToStr(TarifMode) + '"';
        if ExisteSQL(SQL) or (RBPA.Checked) or (RBDPA.Checked) or (RBPVART.Checked) then //or (RBPOURC.Checked) or (RBPVART.Checked) then
        begin
          TOB.Create('', TobTravail, j);
          TobTravail.Detail[j].SelectDB('', Q);
          TobTravail.Detail[j].LoadDB;
          j := j + 1;
        end;
        Q.next;
      end;
      ferme(Q);
    end;
  end;
  Result := TobTravail;
end;

procedure TFMajTarfMode.CreerTobTarif(Depot: string);
var i, TarifMode: Integer;
  SQLTarif: string;
  QTarif: TQuery;
begin
  Application.ProcessMessages;
  TOBTarif := TOB.Create('', nil, -1);
  RecupInfoPeriodeEtab(Depot);
  TarifMode := TOBMode.Detail[0].GetValue('GFM_TARFMODE');
  for i := 0 to TobArtTraite.Detail.Count - 1 do
  begin
    if StopT then exit;
    if ExisteSQL('SELECT GF_TARIF FROM TARIF WHERE GF_ARTICLE="' + TobArtTraite.Detail[i].GetValue('GA_ARTICLE') + '" AND GF_DEPOT="' + Depot +
      '" AND GF_TARFMODE="' + IntToStr(TarifMode) + '" AND GF_FERME="-"') then
    begin
      if (Statut = 'DIM') or (TobArtTraite.Detail[i].GetValue('GA_PRIXUNIQUE') = '-') then
        SQLTarif := RequeteTarif(TobArtTraite.Detail[i].GetValue('GA_ARTICLE'), Depot, TarifMode, true)
      else SQLTarif := RequeteTarif(TobArtTraite.Detail[i].GetValue('GA_ARTICLE'), Depot, TarifMode);
      QTarif := OpenSql(SQLTarif, True);
      if not QTarif.EOF then TobTarif.LoadDetailDB('TARIF', '', '', QTarif, True, False);
      Ferme(QTarif);
      if TobArtTraite.Detail[i].GetValue('GA_PRIXUNIQUE') = '-' then TraiterNvDimension(TobArtTraite.Detail[i].GetValue('GA_ARTICLE'), Depot, TarifMode);
    end else
    begin
      NouveauTarif(TobArtTraite.Detail[i], Depot);
    end;
  end;
end;

procedure TFMajTarfMode.RecupInfoTypeEtPeriode;
var SQL: string;
  Q: TQuery;
  TobTarifMode: TOB;
begin
  // TarifMode
  if TobMode = nil then
  begin
    SQL := 'Select GFM_TARFMODE from TarifMode Where GFM_TYPETARIF="' + TYPETARIFDEST.Value + '" and GFM_PERTARIF="' + PERIODEDEST.Value +
      '" and GFM_NATURETYPE="' + NatureType + '"';
    Q := OpenSQL(SQL, True);
    if Q.EOF then
    begin
      TobTarifMode := TraiterTableTarifMode(TYPETARIFDEST.Value, PERIODEDEST.Value, NatureType, nil);
      TobTarifMode.InsertDB(nil);
      TobTarifMode.Free;
      Q := OpenSQL(SQL, True);
    end;
    TobMode := TOB.Create('', nil, -1);
    TobMode.LoadDetailDB('TARIFMODE', Q.Fields[0].AsString, '', nil, true);
    Ferme(Q);
  end;
  CoefType := CalculCoefTypeTarif(TYPETARIF.Value, TYPETARIFDEST.Value, NatureType, TobMode);
end;

procedure TFMajTarfMode.RecupInfoPeriodeEtab(Depot: string);
begin
  InfoGen := RecupInfoPeriode(PERIODEDEST.Value, Depot) ;
end;

procedure TFMajTarfMode.NouveauTarif(TobArt: Tob; Depot: string);
var i, CodeTarif, j: Integer;
  QMax: TQuery;
  Info: string;
begin
  Application.ProcessMessages;
  if StopT then exit;
  Info := InfoGen;
  CodeTarif := 0;
  if MaxTarif = -1 then
  begin
    QMax := OpenSQL('SELECT MAX(GF_TARIF) FROM TARIF', TRUE);
    if QMax.EOF then MaxTarif := 1 else MaxTarif := QMax.Fields[0].AsInteger + 1;
    Ferme(QMax);
  end else inc(MaxTarif);
  for j := 0 to TOBTarif.Detail.count - 1 do
  begin
    if CodeTarif < TOBTarif.Detail[j].GetValue('GF_TARIF') + 1 then
      CodeTarif := TOBTarif.Detail[j].GetValue('GF_TARIF') + 1;
  end;
  if MaxTarif > CodeTarif then CodeTarif := MaxTarif;
  i := TobTarif.Detail.count;
  TOB.Create('TARIF', TobTarif, i);
  TobTarif.Detail[i].PutValue('GF_TARIF', CodeTarif);
  TobTarif.Detail[i].PutValue('GF_ARTICLE', TobArt.GetValue('GA_ARTICLE'));
  TobTarif.Detail[i].PutValue('GF_TARIFARTICLE', '');
  TobTarif.Detail[i].PutValue('GF_BORNESUP', 999999);
  TobTarif.Detail[i].PutValue('GF_PRIXUNITAIRE', 0);
  TobTarif.Detail[i].PutValue('GF_REMISE', 0);
  TobTarif.Detail[i].PutValue('GF_CALCULREMISE', '');
  TobTarif.Detail[i].PutValue('GF_MODECREATION', 'MAN');
  TobTarif.Detail[i].PutValue('GF_DATEMODIF', NowH);
  //TobTarif.Detail[i].PutValue('GF_ARRONDI',Q.FindField('GFM_ARRONDI').AsString) ; // code arrondi article
  TobTarif.Detail[i].PutValue('GF_DEVISE', TOBMode.Detail[0].GetValue('GFM_DEVISE'));
  TobTarif.Detail[i].PutValue('GF_QUALIFPRIX', 'GRP');
  TobTarif.Detail[i].PutValue('GF_FERME', '-');
  TobTarif.Detail[i].PutValue('GF_TARFMODE', TOBMode.Detail[0].GetValue('GFM_TARFMODE'));
  TobTarif.Detail[i].PutValue('GF_LIBELLE', copy(TYPETARIFDEST.Text + ' - ' + PERIODEDEST.Text, 1, 35));
  if NatureType = 'VTE' then
  begin
    TobTarif.Detail[i].PutValue('GF_REGIMEPRIX', 'TTC');
    TobTarif.Detail[i].PutValue('GF_NATUREAUXI', 'CLI');
  end else
  begin
    TobTarif.Detail[i].PutValue('GF_REGIMEPRIX', 'HT');
    TobTarif.Detail[i].PutValue('GF_NATUREAUXI', 'FOU');
  end;
  TobTarif.Detail[i].PutValue('GF_BORNEINF', -999999);
  TobTarif.Detail[i].PutValue('GF_QUANTITATIF', '-');
  TobTarif.Detail[i].PutValue('GF_DEPOT', Depot);
  TobTarif.Detail[i].PutValue('GF_SOCIETE', Depot);
  //Info:=RecupInfoPeriode(PERIODEDEST.Value,Depot) ;
  TobTarif.Detail[i].PutValue('GF_DATEDEBUT', StrToDate(ReadTokenSt(Info)));
  TobTarif.Detail[i].PutValue('GF_DATEFIN', StrToDate(ReadTokenSt(Info)));
  TobTarif.Detail[i].PutValue('GF_ARRONDI', ReadTokenSt(Info));
  TobTarif.Detail[i].PutValue('GF_CASCADEREMISE', ReadTokenSt(Info));
  TobTarif.Detail[i].PutValue('GF_DEMARQUE', ReadTokenSt(Info));
  if (TobArt.GetValue('GA_STATUTART') = 'GEN') and (TobArt.GetValue('GA_PRIXUNIQUE') <> 'X') then CreerTarifDim(TobTarif.Detail[i]);
end;

procedure TFMajTarfMode.CreerTarifDim(TobTarfGen: TOB);
var MaxTarif, j: Integer;
  QArtDim: TQuery;
  CodeArticle: string;
  PrixGen, PrixAchGen: Double;
  TobPrixGen, TobPrixDim: Tob;
begin
  Application.ProcessMessages;
  PrixGen := 0;
  PrixAchGen := 0;
  if StopT then exit;
  if (RBPOURC.Checked) or (RBPXSAISI.Checked) then exit;
  TobPrixGen := TobPrix.FindFirst(['GA_ARTICLE'], [TobTarfGen.GetValue('GF_ARTICLE')], False);
  if TobPrixGen <> nil then
  begin
    if RBPA.Checked then
    begin
      if VarIsNull(TobPrixGen.GetValue('GCA_PRIXBASE')) then PrixGen := 0 else PrixGen := TobPrixGen.GetValue('GCA_PRIXBASE');
    end;
    if RBDPA.Checked then
    begin
      if VarIsNull(TobPrixGen.GetValue('GQ_DPA')) then PrixGen := 0 else PrixGen := (TobPrixGen.GetValue('GQ_DPA'));
    end;
    if RBPVART.Checked then PrixGen := TobPrixGen.GetValue('GA_PVTTC');
    if (RBPVTTC.Checked) or (RBAHT.Checked) then PrixGen := TobPrixGen.GetValue('GF_PRIXUNITAIRE');
    PrixAchGen := TobPrixGen.GetValue('GA_PAHT');
  end;
  J := TobTarif.Detail.count;
  MaxTarif := TobTarfGen.GetValue('GF_TARIF') + 1;
  codeArticle := TobTarfGen.GetValue('GF_ARTICLE');
  QArtDim := OpenSql('Select GA_ARTICLE from Article where GA_CODEARTICLE="' + TRIM(copy(TobTarfGen.GetValue('GF_ARTICLE'), 1, 18)) +
    '" And GA_STATUTART="DIM"', True);
  while not QArtDim.EOF do
  begin
    TobPrixDim := TobPrix.FindFirst(['GA_ARTICLE'], [QArtDim.FindField('GA_ARTICLE').AsString], False);
    if PrixDifferent(TobPrixDim, PrixGen, PrixAchGen) then
    begin
      TOB.Create('TARIF', TobTarif, j);
      TobTarif.Detail[j].Dupliquer(TobTarfGen, False, True);
      TobTarif.Detail[j].PutValue('GF_ARTICLE', QArtDim.FindField('GA_ARTICLE').AsString);
      TobTarif.Detail[j].PutValue('GF_TARIF', MaxTarif);
      QArtDim.next;
      MaxTarif := MaxTarif + 1;
      j := j + 1;
    end else QArtDim.next;
  end;
  ferme(QArtDim);
end;

procedure TFMajTarfMode.TraiterNvDimension(CodeArticle, Depot: string; TarifMode: integer);
var TobTarifGen: TOB;
begin
  TobTarifGen := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_TARFMODE'], [CodeArticle, Depot, TarifMode], False);
  CreerTarifDim(TobTarifGen);
end;

function TFMajTarfMode.PrixDifferent(TobPxDim: Tob; PxGen, PxAchGen: Double): Boolean;
var PrixDim: Double;
begin
  result := False;
  PrixDim := 0;
  if TobPxDim = nil then exit;
  if TobPxDim <> nil then
  begin
    if RBPA.Checked then
      if VarIsNull(TobPxDim.GetValue('GCA_PRIXBASE')) then PrixDim := 0 else PrixDim := TobPxDim.GetValue('GCA_PRIXBASE')
    else
      if RBDPA.Checked then
      if VarIsNull(TobPxDim.GetValue('GQ_DPA')) then PrixDim := 0 else PrixDim := TobPxDim.GetValue('GQ_DPA')
    else
      if RBPVART.Checked then PrixDim := TobPxDim.GetValue('GA_PVTTC')
    else
      if (RBPVTTC.Checked) or (RBAHT.Checked) then PrixDim := TobPxDim.GetValue('GF_PRIXUNITAIRE');
  end;
  if PrixDim = 0 then PrixDim := TobPxDim.GetValue('GA_PAHT');
  if (PxGen <> PrixDim) and (PxAchGen <> PrixDim) then result := true;
end;

function RequeteTarif(Article, Depot: string; TarifMode: Integer; Dim: Boolean = False): string;
var SQL: string;
begin
  SQL := 'Select * from tarif where ';
  SQL := SQL + ' GF_DEPOT="' + Depot + '"';
  if Dim then SQL := SQL + ' AND GF_ARTICLE like "' + TRIM(Copy(Article, 1, 18)) + '%"'
  else SQL := SQL + ' AND GF_ARTICLE ="' + Article + '"';
  SQL := SQL + ' AND GF_FERME="-"';
  SQL := SQL + ' AND GF_TARFMODE="' + IntToStr(TarifMode) + '"';
  SQL := SQL + ' ORDER BY GF_PRIORITE DESC';
  Result := SQL;
end;

function CalculCoefTypeTarif(TypeTarifRef, TypeTarifDest, NatureType: string; TobMode: Tob): Double;
var CoefTypeRef, CoefTypeDest: Double;
  TobM: TOB;
begin
  Result := 0;
  CoefTypeRef := 0;
  if TobMode = nil then
  begin
    if TypeTarifRef <> '' then CoefTypeRef := RecupCoefTypeTarif(TypeTarifRef, NatureType);
    CoefTypeDest := RecupCoefTypeTarif(TypeTarifDest, NatureType);
  end else
  begin
    if TypeTarifRef <> '' then
    begin
      TobM := TOBMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarifRef, NatureType], True);
      if TobM = nil then CoefTypeRef := RecupCoefTypeTarif(TypeTarifRef, NatureType) else CoefTypeRef := TOBM.GetValue('GFM_COEF');
    end;
    TobM := TOBMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarifDest, NatureType], True);
    if TobM = nil then CoefTypeDest := RecupCoefTypeTarif(TypeTarifDest, NatureType) else CoefTypeDest := TOBM.GetValue('GFM_COEF');
  end;
  if (CoefTypeDest = 0) then exit;
  if CoefTypeRef = CoefTypeDest then
  begin
    Result := 0;
    exit;
  end;
  if (CoefTypeRef = 0) then
  begin
    result := CoefTypeDest;
    exit;
  end;
  if (CoefTypeRef = 1) and (CoefTypeDest <> 1) then
  begin
    Result := CoefTypeDest;
    exit;
  end;
  if (CoefTypeRef <> 1) and (CoefTypeDest = 1) then
  begin
    Result := 1 / CoefTypeRef;
    exit;
  end;
  if (CoefTypeRef <> 1) and (CoefTypeDest <> 1) then
  begin
    result := (1 / CoefTypeRef) * CoefTypeDest;
    exit;
  end;
end;

function TFMajTarfMode.MAJDepot(Depot: string): Boolean;
begin
  Application.ProcessMessages;
  CreerTobTarif(Depot);
  //MajTarifArticle(Depot) ;
  if Statut = 'DIM' then CreerTarifGen;
  MajTarifArticle(Depot);
  MAJDateModif;
  Result := TOBTarif.InsertOrUpdateDB(True);
  if TarifASup <> '' then SuppressionDimension(TarifASup);
  LibereDepot;
  if Result = true and not StopT then RemplirTobRecap(TOBTarif, 'Reussite') else RemplirTobRecap(TOBTarif, 'Anomalie');
  if StopT then Result := False;
end;

procedure TFMajTarfMode.LibereDepot;
begin
  TobTarif.Free;
  TobTarif := nil;
  TobPrix.Free;
  TobPrix := nil;
end;

procedure TFMajTarfMode.MAJDateModif;
var i: Integer;
begin
  for i := 0 to TOBTarif.Detail.count - 1 do
  begin
    TOBTarif.Detail[i].PutValue('GF_DATEMODIF', NowH);
  end;
end;

procedure TFMajTarfMode.RemplirTobRecap(TOBTarif: TOB; Statut: string);
var i, k: Integer;
begin
  Application.ProcessMessages;
  if Statut = 'Anomalie' then
  begin
    k := TobAnomalie.Detail.count;
    for i := 0 to TOBTarif.Detail.Count - 1 do
    begin
      TOB.Create('_Anomalie', TobAnomalie, k);
      TobAnomalie.Detail[k].AddChampSup('ARTICLE', False);
      TobAnomalie.Detail[k].AddChampSup('DEPOT', False);
      TobAnomalie.Detail[k].PutValue('ARTICLE', TobTarif.Detail[i].GetValue('GF_ARTICLE'));
      TobAnomalie.Detail[k].PutValue('DEPOT', TobTarif.Detail[i].GetValue('GF_DEPOT'));
      k := k + 1;
    end;
  end;
end;

procedure TFMajTarfMode.MajTarifArticle(Depot: string);
var i_ind, i_gen: integer;
  TOBD, TobDevise: TOB;
  PrixInit: Double;
  CodeArtGen: string;
begin
  Application.ProcessMessages;
  for i_ind := TOBTarif.detail.count - 1 downto 0 do
  begin
    if StopT then exit;
    PrixInit := 0;
    if (Statut = 'DIM') and (TOBTarif.detail[i_ind].FieldExists('GENZERO')) then Break;
    if RBPXSAISI.Checked then
    begin
      PrixInit := PXSAISI.Value;
      if PrixInit <> 0 then
      begin
        MajPeriode(TOBTarif.detail[i_ind]);
        TOBTarif.detail[i_ind].PutValue('GF_PRIXANCIEN', TOBTarif.detail[i_ind].GetValue('GF_PRIXUNITAIRE'));
        TOBTarif.detail[i_ind].PutValue('GF_PRIXUNITAIRE', PrixInit);
        CalcPriorite(TOBTarif.detail[i_ind]);
      end else TOBTarif.detail[i_ind].free;
    end else
    begin
      TOBD := TOBPrix.FindFirst(['GA_ARTICLE'], [TOBTarif.Detail[i_ind].GetValue('GF_ARTICLE')], False);
      if TOBD <> nil then
      begin
        MajPeriode(TOBTarif.detail[i_ind]);
        if (TOBD.GetValue('GA_STATUTART') = 'DIM') and (TOBD.GetValue('GA_PRIXUNIQUE') = 'X') then
        begin
          Anomalie(TOBD, Depot, i_ind);
        end
        else
        begin
          if RBPA.Checked or RBDPA.Checked or RBPVTTC.Checked or RBPVART.Checked or RBAHT.Checked then
          begin
            if TOBTarif.detail[i_ind].GetValue('GF_REMISE') <> 0 then Anomalie(TOBD, Depot, i_ind) else
            begin
              if RBPA.Checked or RBDPA.Checked then
              begin
                if RBPA.Checked then
                  if VarIsNull(TOBD.GetValue('GCA_PRIXBASE')) then PrixInit := 0 else PrixInit := TOBD.GetValue('GCA_PRIXBASE');
                if RBDPA.Checked then
                  if VarIsNull(TOBD.GetValue('GQ_DPA')) then PrixInit := 0 else PrixInit := (TOBD.GetValue('GQ_DPA'));
                if (PrixInit = 0) and (not VarIsNull(TOBD.GetValue('GA_PAHT'))) then PrixInit := TOBD.GetValue('GA_PAHT');
              end
              else
                if (RBPVTTC.Checked) or (RBAHT.Checked) then PrixInit := TOBD.GetValue('GF_PRIXUNITAIRE')
              else
                if RBPVART.Checked then PrixInit := TOBD.GetValue('GA_PVTTC');
              if (CoefType <> 0) then
                PrixInit := PrixInit * COEF.Value * CoefType else PrixInit := PrixInit * COEF.Value;
              if PrixInit <> 0 then PrixInit := ArrondirPrix(ARRONDIC.Value, PrixInit);
              TobDevise := TOBDev.FindFirst(['D_DEVISE'], [TOBTarif.detail[i_ind].GetValue('GF_DEVISE')], False);
              PrixInit := Arrondi(PrixInit, TobDevise.GetValue('D_DECIMALE'));
              if PrixInit <> 0 then
              begin
                TOBTarif.detail[i_ind].PutValue('GF_PRIXANCIEN', TOBTarif.detail[i_ind].GetValue('GF_PRIXUNITAIRE'));
                TOBTarif.detail[i_ind].PutValue('GF_PRIXUNITAIRE', PrixInit);
                CalcPriorite(TOBTarif.detail[i_ind]);
              end else TOBTarif.detail[i_ind].free;
            end;
          end
          else
          begin
            if ((RBPOURC.Checked) or (TOBD.GetValue('GF_REMISE') <> 0)) then
            begin
              if TOBTarif.detail[i_ind].GetValue('GF_PRIXUNITAIRE') <> 0 then Anomalie(TOBD, Depot, i_ind) else
              begin
                TOBTarif.detail[i_ind].PutValue('GF_REMISEANCIEN', TOBTarif.detail[i_ind].GetValue('GF_REMISE'));
                if RBPOURC.Checked then TOBTarif.detail[i_ind].PutValue('GF_CALCULREMISE', FloatToStr(COEF.Value))
                else TOBTarif.detail[i_ind].PutValue('GF_CALCULREMISE', TOBD.GetValue('GF_CALCULREMISE'));
                TOBTarif.detail[i_ind].PutValue('GF_REMISE', TOBTarif.detail[i_ind].GetValue('GF_CALCULREMISE'));
                CalcPriorite(TOBTarif.detail[i_ind]);
              end;
            end;
          end;
        end;
      end else
      begin
        if TarifASup = '' then TarifASup := IntToStr(TOBTarif.detail[i_ind].GetValue('GF_TARIF'))
        else TarifASup := TarifASup + ',' + IntToStr(TOBTarif.detail[i_ind].GetValue('GF_TARIF'));
        TOBTarif.detail[i_ind].Free;
      end;
    end;
  end;
end;

procedure TFMajTarfMode.Anomalie(TOBD: TOB; Depot: string; i_ind: Integer);
var k: Integer;
begin
  k := TobAnomalie.Detail.count;
  TOB.Create('_Anomalie', TobAnomalie, k);
  TobAnomalie.Detail[k].AddChampSup('ARTICLE', False);
  TobAnomalie.Detail[k].AddChampSup('DEPOT', False);
  TobAnomalie.Detail[k].PutValue('ARTICLE', TobD.GetValue('GA_ARTICLE'));
  TobAnomalie.Detail[k].PutValue('DEPOT', Depot);
  TOBTarif.Detail[i_ind].free;
end;

procedure TFMajTarfMode.CreerTarifGen;
var TobGen, TobDim, TOBG: Tob;
  CodeArticle, CodeArtDim: string;
  i: Integer;
begin
  CodeArticle := '';
  for i := TOBTarif.detail.count - 1 downto 0 do
  begin
    TobDim := TobTarif.Detail[i];
    CodeArtDim := Trim(copy(TobDim.GetValue('gf_article'), 1, 18));
    if CodeArticle <> CodeArtDim then
    begin
      TOBG := TOBTarif.FindFirst(['GF_ARTICLE'], [CodeArticleUnique2(CodeArtDim, '')], False);
      if TOBG <> nil then
      begin
        TOBG.AddChampSup('GENZERO', False);
        continue;
      end;
      CodeArticle := Trim(copy(TobDim.GetValue('gf_article'), 1, 18));
      TobGen := TOB.Create('TARIF', TobTarif, i);
      TobGen.Dupliquer(TobDim, False, True, True);
      TobGen.PutValue('GF_ARTICLE', CodeArticleUnique2(CodeArticle, ''));
      TobGen.PutValue('GF_CALCULREMISE', '');
      TobGen.PutValue('GF_REMISE', 0);
      TobGen.PutValue('GF_PRIXUNITAIRE', 0);
      TobGen.AddChampSup('GENZERO', False);
      NumTarif;
    end;
  end;
  //NumTarif ;
end;

procedure TFMajTarfMode.NumTarif;
var i, MaxTarif: Integer;
  QMax: TQuery;
begin
  QMax := OpenSQL('SELECT MAX(GF_TARIF) FROM TARIF', TRUE);
  if QMax.EOF then MaxTarif := 1 else MaxTarif := QMax.Fields[0].AsInteger + 1;
  Ferme(QMax);
  for i := 0 to TobTarif.Detail.count - 1 do
  begin
    TobTarif.Detail[i].PutValue('GF_TARIF', MaxTarif);
    Inc(MaxTarif);
  end;
end;

procedure TFMajTarfMode.MajPeriode(TOBT: TOB);
var Info: string;
begin
  Info := InfoGen;
  TOBT.PutValue('GF_DATEDEBUT', StrToDate(ReadTokenSt(Info)));
  TOBT.PutValue('GF_DATEFIN', StrToDate(ReadTokenSt(Info)));
  TOBT.PutValue('GF_ARRONDI', ReadTokenSt(Info));
  TOBT.PutValue('GF_CASCADEREMISE', ReadTokenSt(Info));
  TOBT.PutValue('GF_DEMARQUE', ReadTokenSt(Info));
end;

procedure SuppressionDimension(TarifASup: string);
var SQLTarif: string;
begin
  SQLTarif := 'delete from tarif where GF_TARIF IN (' + TarifASup + ')';
  ExecuteSQL(SQLTarif);
end;

end.
