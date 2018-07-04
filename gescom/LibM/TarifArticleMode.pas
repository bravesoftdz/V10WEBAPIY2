{***********UNITE*************************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Saisie des tarifs article et catégories article
Mots clefs ... : TARIF;ARTICLE
*****************************************************************}
unit TarifArticleMode;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  TarifArticle, Menus, hmsgbox, HSysMenu, Grids, Hctrls, StdCtrls,
  ComCtrls, HRichEdt, HRichOLE, ExtCtrls, HTB97, Mask, HPanel, Utob, AGLInit,
  TarifUtil, Hent1, UIUtil, LookUp, SaisUtil, AglInitGc, UtilArticle, UtilDimArticle, entgc,
  ParamSoc, math, utilPGI, Ent1,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  dbctrls, dbTables, Fe_Main,
  {$ENDIF}
  M3FP;

function EntreeTarifArticleMode(Action: TActionFiche; TarifTTC: Boolean = False): boolean;
function SaisieTarifArticleMode(CodeArt: string; Action: TActionFiche; TarifTTC: Boolean = False): boolean;

type
  TFTarifArticleMode = class(TFTarifArticle)
    BAffDim: TToolbarButton97;
    Ptype: TPanel;
    PEtabli: TPanel;
    _TYPTARIFMODE: THValComboBox;
    TGF_TYPTARIF_MODE: THLabel;
    TGF_DEVISE_: THLabel;
    TGF_DEPOT: THLabel;
    TGF_PERTARIF_MODE: THLabel;
    _PERTARIFMODE: THValComboBox;
    TGF_DATEDEBUT: THLabel;
    TGF_DATEFIN: THLabel;
    GF_DEPOT: THValComboBox;
    HTarif: THMsgBox;
    GF_DEVISE_: TEdit;
    GF_DATEDEBUT: TEdit;
    GF_DATEFIN: TEdit;
    TGF_DEMARQUE: TLabel;
    GF_DEMARQUE: TEdit;
    BNVTYPE: TToolbarButton97;
    BNVPERIODE: TToolbarButton97;

    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);

    // Evènement des boutons
    procedure BAffDimClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BNvPeriodeClick(Sender: TObject);
    procedure BNvTypeClick(Sender: TObject);
    procedure BZoomArticle(Sender: TObject);

    procedure _TYPTARIFMODEExit(Sender: TObject);
    procedure _PERTARIFMODEExit(Sender: TOBject);
    procedure GF_DEPOTExit(Sender: TOBject);
    procedure _TYPTARIFMODEChange(Sender: TObject);
    procedure _PERTARIFMODEChange(Sender: TObject);
    procedure GF_DEPOTChange(Sender: TObject);

    procedure InitialiseEntete; override;

    // Evenement du grid
    procedure G_ARTCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_ARTCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_ARTElipsisClick(Sender: TObject);
    procedure G_ARTRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_ARTRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_ARTEnter(Sender: TObject);

  protected
    TobTarfArtDim, TobTarifMode, TobArticle, TobArtExist: Tob; //AC
    ValidationTarifOK, SaisiPrix, Promo: Boolean;
    HMTrad: THSystemMenu;
    NouvelleLigne, InitEntete, SavEntete, ModifMul: Boolean;
    CodeType, CodePeriode, CodeDepot, TarifArticle, CodeArrondi, CodeRemise, CodeDemarque: string;
    CodeTarifMode, IdTarif, RegimePrix, NatureType: string;
    CodeTarif: Longint;
    LigneASav: Integer;
    CoefType: Double;

    // Action liées au dimension
    procedure AfficheDim(Article: string);
    procedure GereTobDim(Arow: Integer);
    procedure MiseAjourDim(ARow: Integer);

    procedure VoirFicheArticle;

    //Validation
    procedure SupDimPrixGen(ARow: Integer);
    procedure ValideTarif; override;
    procedure VerifLesTOB; override;
    function LigneOk(ARow: Integer): Boolean;
    procedure MAJDateModif;

    // Initialisations
    procedure LoadLesTOB; override;

    procedure LibereTOBMode;

    //Init
    procedure InitTarif;

    // Actions liées au Grid
    procedure EtudieColsListe; override;
    procedure FormateZoneSaisie(ACol, ARow: Longint; ttd: T_TableTarif); override;
    function GrilleModifie: Boolean; override;
    function ChampModifie(ARow: Integer): Boolean;
    procedure FormateGrilleSaisie(Promo: Boolean);

    // Action liée à la ligne
    function LigneVide(ARow: integer; ttd: T_TableTarif): Boolean; override;
    procedure InitLaLigne(ARow: integer; ttd: T_TableTarif); override;

    procedure TraiterArticle(ACol, ARow: Integer);
    procedure TraiterTarifArticle(ACol, ARow: Integer);
    procedure TraiterEtablissement(ARow: integer);
    procedure TraiterRemise(ACol, ARow: integer; ttd: T_TableTarif); override;
    procedure TraiterPrix(ACol, ARow: integer; ttd: T_TableTarif); override;

    // Traiter la table Tarif Mode
    procedure MAJChampEntete; // MAJ Bouleen dernier utilise pour entete

    // Visualiser un tarif existant à partir de la consultation
    procedure ChargeTarifDepuisMul;

  public
    TarifExistant: Boolean;
    Etablissement: string;

  end;

var
  FTarifArticleMode: TFTarifArticleMode;
function TrouverArticle(Article: string; TOBArt: TOB): T_RechArt;

implementation

function EntreeTarifArticleMode(Action: TActionFiche; TarifTTC: Boolean = False): boolean;
begin
  result := SaisieTarifArticleMode('', Action, TarifTTC);
end;

function SaisieTarifArticleMode(CodeArt: string; Action: TActionFiche; TarifTTC: Boolean = False): boolean;
var FF: TFTarifArticleMode;
  PPANEL: THPanel;
begin
  Result := True;
  SourisSablier;
  FF := TFTarifArticleMode.Create(Application);
  FF.BVoirCond.Visible := False;
  FF.CBQUANTITAIF.Visible := False;
  FF.CBCATTIERS.Visible := False;
  FF.PEtabli.Visible := True;
  FF.Ptype.Visible := True;
  FF.Action := Action;
  FF.TarifTTC := TarifTTC;
  FF.CodeArticle := CodeArt;
  PPANEL := FindInsidePanel;
  if PPANEL = nil then
  begin
    try
      FF.BorderStyle := bsSingle;
      FF.ShowModal;
    finally
      FF.Free;
      Result := False;
    end;
    SourisNormale;
  end else
  begin
    InitInside(FF, PPANEL);
    FF.Show;
  end;
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFTarifArticleMode.FormCreate(Sender: TObject);
begin
  inherited;
  //TobTarifMode:=TOB.Create('TARIFMODE',NIL,-1) ;
end;

procedure TFTarifArticleMode.FormDestroy(Sender: TObject);
begin
  inherited;
  TobTarfArtDim.Free;
  AglDepileFiche;
end;

procedure TFTarifArticleMode.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  G_Art.VidePile(True);
  G_Qte.VidePile(True);
  G_QCa.VidePile(True);
  TOBTarif.Free;
  TOBTarif := nil;
  TOBArt.Free;
  TOBArt := nil;
  TobArtExist.Free;
  TobArtExist := nil;
  LibereTOBMode;
  if IsInside(Self) then Action := caFree;
  FClosing := True;
end;

procedure TFTarifArticleMode.FormShow(Sender: TObject);
begin
  G_ART.ListeParam := 'GCTARIFPRIXMODE';
  G_ART.PostDrawCell := PostDrawCell;
  G_Qte.ListeParam := 'GCTARIFQTEPRIX';
  G_Qte.PostDrawCell := PostDrawCell;
  G_QCa.ListeParam := 'GCTARIFCTQTEPRIX';
  G_QCa.PostDrawCell := PostDrawCell;
  G_Cat.ListeParam := 'GCTARIFCTPRIX';
  G_Cat.PostDrawCell := PostDrawCell;
  if TarifTTC then
  begin
    RegimePrix := 'TTC';
    NatureType := 'VTE';
  end else
  begin
    RegimePrix := 'HT';
    NatureType := 'ACH';
  end;
  EtudieColsListe;
  HMTrad.ResizeGridColumns(G_Art);
  AffecteGrid(G_Art, Action);
  PART.Visible := True;
  PQUANTITATIF.Visible := False;
  PQTECAT.Visible := False;
  PCATEGORIE.Visible := FAlse;
  PTITRE.Caption := HTitre.Mess[3];
  CBQUANTITAIF.Checked := False;
  CBCATTIERS.Checked := False;
  if TarifTTC then _TYPTARIFMODE.DataType := 'GCTARIFTYPE1VTE' else
    _TYPTARIFMODE.DataType := 'GCTARIFTYPE1ACH';
  _TYPTARIFMODE.Reload;
  if TheTob <> nil then
  begin
    ChargeTarifDepuisMul;
  end else
  begin
    if CodeArticle <> '' then PrepareEntete;
    InitialiseEntete;
  end;
  DEV.Code := CodeDevise;
  GetInfosDevise(DEV);
  if TarifTTC then TTYPETARIF.Caption := HTitre.Mess[6] + RechDom('TTDEVISE', CodeDevise, False)
  else TTYPETARIF.Caption := HTitre.Mess[5] + RechDom('TTDEVISE', CodeDevise, False);
  InitComboChamps;
  if Action = taConsult then //NA : 20/09/2002
  begin
    BNvPeriode.Visible := False;
    BNvType.Visible := False;
  end;
  AglEmpileFiche(Self);
end;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}

procedure TFTarifArticleMode.G_ARTCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  if not Cancel then
  begin
    if (G_Art.col = SA_Px) then
    begin
      G_Art.ElipsisButton := false;
      if (G_Art.Cells[SA_TarifArticle, G_Art.Row] = '') then G_Art.Col := SA_Article;
      if (G_Art.Cells[SA_TarifArticle, G_Art.Row] <> '') or (Valeur(G_Art.Cells[SA_Rem, G_Art.Row]) <> 0) then G_Art.Col := SA_Rem;
    end;
    if (G_Art.col = SA_Rem) then
    begin
      G_Art.ElipsisButton := false;
      if (Valeur(G_Art.Cells[SA_Px, G_Art.Row]) <> 0) then G_Art.Col := SA_Px;
    end;
    if (G_Art.col = SA_Article) then
      if (G_Art.Cells[SA_TarifArticle, G_Art.Row] <> '') then
      begin
        SaisiPrix := True;
        G_Art.Col := SA_Rem;
      end
      else if (G_Art.Cells[SA_Article, G_Art.Row] <> '') then
      begin
        ACol := G_Art.col;
        ARow := G_Art.Row;
        G_ArtCellExit(Sender, ACol, ARow, cancel);
        G_Art.Col := SA_Px;
      end;
    if (G_Art.col = SA_TarifArticle) then
      if (G_Art.Cells[SA_Article, G_Art.Row] <> '') then
      begin
        SaisiPrix := True;
        G_Art.Col := SA_Px;
      end
      else if (G_Art.Cells[SA_TarifArticle, G_Art.Row] <> '') then
      begin
        ACol := G_Art.col;
        ARow := G_Art.Row;
        G_ArtCellExit(Sender, ACol, ARow, cancel);
        G_Art.Col := SA_Rem;
      end;
    G_Art.ElipsisButton := ((G_Art.Col = SA_Article) or (G_Art.Col = SA_TarifArticle));
    StCellCur := G_Art.Cells[G_Art.Col, G_Art.Row];
    AffectMenuCondApplic(G_Art, ttdArt);
  end;
end;

procedure TFTarifArticleMode.G_ARTCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var Prix: Double;
begin
  FormateZoneSaisie(ACol, ARow, ttdArt);
  if ACol = SA_Article then
  begin
    if (G_Art.Cells[ACol, Arow] <> '') then
    begin
      if SaisiPrix then exit;
      if (Valeur(G_Art.Cells[Sa_Px, ARow]) = 0) and (Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0) then ValidationTarifOk := False;
      TraiterArticle(ACol, ARow);
    end else
    begin
      CodeArticle := '';
      exit;
    end;
  end else
    if ACol = SA_TarifArticle then
  begin
    if (G_Art.Cells[Sa_Article, Arow] <> '') then exit;
    if G_Art.Cells[ACol, Arow] <> '' then
    begin
      if Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0 then ValidationTarifOk := False;
      TraiterTarifArticle(ACol, ARow);
    end else
    begin
      TarifArticle := '';
      exit;
    end;
  end else
    if ACol = SA_Px then
  begin
    Prix := Valeur(G_Art.Cells[SA_Px, ARow]);
    if Valeur(G_Art.Cells[Sa_Px, ARow]) = 0 then ValidationTarifOK := False;
    G_Art.Cells[SA_Px, ARow] := FloatToStr(Prix);
    TraiterPrix(ACol, Arow, ttdArt);
    if Prix <> 0 then
    begin
      G_Art.Row := ARow + 1;
      if G_Art.Cells[Sa_Article, Arow] <> '' then G_Art.Col := Sa_Article;
      if G_Art.Cells[Sa_TarifArticle, Arow] <> '' then G_Art.Col := Sa_TarifArticle;
    end;
  end else
    if ACol = SA_Rem then
  begin
    if Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0 then ValidationTarifOK := False;
    TraiterRemise(ACol, Arow, ttdArt);
    if Valeur(G_Art.Cells[SA_Rem, ARow]) <> 0 then
    begin
      if (G_Art.Cells[Sa_Article, Arow] = '') and (G_Art.Cells[Sa_TarifArticle, Arow] = '') then G_Art.Row := ARow + 1;
      if G_Art.Cells[Sa_Article, Arow] <> '' then G_Art.Col := Sa_Article;
      if G_Art.Cells[Sa_TarifArticle, Arow] <> '' then G_Art.Col := Sa_TarifArticle;
    end else if (G_Art.Cells[Sa_Article, Arow] <> '') then G_Art.Col := Sa_Px else G_Art.Row := Sa_Rem;
  end;
  if not Cancel then
  begin
  end;
end;

procedure TFTarifArticleMode.G_ARTElipsisClick(Sender: TObject);
var TARIFARTICLE: THCritMaskEdit;
  Coord: TRect;
  Article, Categorie: string;
  Cancel: Boolean;
  ACol, ARow: Integer;
begin
  inherited;
  Cancel := False;
  ARow := G_Art.Row;
  if G_Art.Col = SA_Article then
  begin
    Coord := G_Art.CellRect(G_Art.Col, G_Art.Row);
    Article := DispatchArtMode(1, '', 'GA_CODEARTICLE=' + Trim(Copy(CodeArticle, 1, 18)), 'TARIF');
    if Article <> '' then
    begin
      G_Art.Cells[G_Art.Col, G_Art.Row] := Article;
      ACol := SA_Article;
      G_ARTCellExit(Sender, ACol, ARow, Cancel);
      if (G_Art.Cells[Sa_Article, G_Art.Row] <> '') then G_Art.Col := Sa_Px;
    end
    else CodeArticle := '';
  end;
  if G_Art.Col = SA_TarifArticle then
  begin
    Coord := G_Art.CellRect(G_Art.Col, G_Art.Row);
    TARIFARTICLE := THCritMaskEdit.Create(Self);
    TARIFARTICLE.Parent := G_Art;
    TARIFARTICLE.Top := Coord.Top;
    TARIFARTICLE.left := Coord.Left;
    TARIFARTICLE.Width := 3;
    TARIFARTICLE.Visible := False;
    TARIFARTICLE.DataType := 'GCTARIFARTICLE';
    LookUpCombo(TARIFARTICLE);
    if TARIFARTICLE.Text <> '' then
    begin
      categorie := TARIFARTICLE.Text;
      G_Art.Cells[G_Art.Col, G_Art.Row] := TARIFARTICLE.Text;
      ACol := SA_TarifArticle;
      G_ARTCellExit(Sender, ACol, ARow, Cancel);
      if (G_Art.Cells[Sa_TarifArticle, G_Art.Row] <> '') then G_Art.Col := SA_Rem;
    end;
    TARIFARTICLE.Destroy;
  end;
end;

procedure TFTarifArticleMode.G_ARTRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var Arow: Integer;
begin
  //inherited;
  if Ou >= G_Art.RowCount - 1 then G_Art.RowCount := G_Art.RowCount + NbRowsPlus;
  ;
  ARow := Min(Ou, TOBTarfArt.detail.count + 1);
  if (LigneVide(ARow, ttdArt)) and (not LigneVide(ARow - 1, ttdArt)) then
    PreAffecteLigne(ARow, ttdArt);

  AffichePied(ttdArt);
  AfficheCondTarf(ARow, ttdArt);
  CodeArticle := G_Art.Cells[SA_Article, G_Art.Row];
  //if (ModifMul) and (LigneVide(Ou,ttdArt)) then BAbandonClick(Nil);
  TraiterEtablissement(Ou);
  if (G_Art.Cells[SA_Article, G_Art.Row] <> '') then
  begin
    TOBArt.SelectDB('"' + G_Art.Cells[SA_Article, G_Art.Row] + '"', nil);
    if (TOBArt.GetValue('GA_STATUTART') = 'GEN') then
    begin
      BAffDim.Visible := True;
      if (TOBArt.GetValue('GA_PRIXUNIQUE') = 'X') then PTITRE.Caption := HTitre.Mess[9]
      else PTITRE.Caption := HTitre.Mess[8];
    end else PTITRE.Caption := HTitre.Mess[7];
  end else
  begin
    BaffDim.Visible := False;
    PTITRE.Caption := HTitre.Mess[3];
  end;
  if (G_Art.Cells[SA_TarifArticle, G_Art.Row] <> '') then PTITRE.Caption := HTitre.Mess[10];
end;

procedure TFTarifArticleMode.G_ARTRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
  if not LigneOk(Ou) then exit;
  NouvelleLigne := True;
  if LigneVide(G_ART.Row, ttdArt) and (not ValidationTarifOK) and ((Valeur(G_Art.Cells[Sa_Px, G_Art.Row - 1]) <> 0) or (Valeur(G_Art.Cells[SA_Rem, G_Art.Row -
    1]) <> 0)) or (GrilleModifie and ChampModifie(G_Art.Row - 1)) then //and not ligneVide(G_ART.Row,ttdArt)) then
    ValideTarif;
end;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}

procedure TFTarifArticleMode.EtudieColsListe;
var NomCol, LesCols: string;
  icol, ichamp, i_ind: integer;
begin
  SA_Depot := 10;
  SA_Lib := 11;
  SA_Px := 12;
  SA_Rem := 13;
  SA_Datedeb := 14;
  SA_Datefin := 15;
  SA_Article := 16;
  SA_TarifArticle := 17;
  G_Qte.ColWidths[0] := 0;
  G_QCa.ColWidths[0] := 0;
  G_Cat.ColWidths[0] := 0;
  for i_ind := Low(ColsInter) to High(ColsInter) do ColsInter[i_ind] := False;
  LesCols := G_Art.Titres[0];
  LesColArt := LesCols;
  icol := 0;
  repeat
    NomCol := uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if Pos('X', V_PGI.Dechamps[iTableLigne, ichamp].Control) > 0 then ColsInter[icol] := True;
        if NomCol = 'GF_DEPOT' then SA_Depot := icol else
          if NomCol = 'GF_ARTICLE' then SA_Article := icol else
          if NomCol = 'GF_TARIFARTICLE' then SA_TarifArticle := icol else
          if NomCol = 'GF_PRIXUNITAIRE' then SA_Px := icol else
          if NomCol = 'GF_CALCULREMISE' then SA_Rem := icol else
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));
  LesCols := G_Qte.Titres[0];
  LesColQtes := LesCols;
  icol := 0;
  repeat
    NomCol := uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if Pos('X', V_PGI.Dechamps[iTableLigne, ichamp].Control) > 0 then ColsInter[icol] := True;
        if NomCol = 'GF_DEPOT' then SG_Depot := icol else
          if NomCol = 'GF_LIBELLE' then SG_Lib := icol else
          if NomCol = 'GF_BORNEINF' then SG_QInf := icol else
          if NomCol = 'GF_BORNESUP' then SG_QSup := icol else
          if NomCol = 'GF_PRIXUNITAIRE' then SG_Px := icol else
          if NomCol = 'GF_CALCULREMISE' then SG_Rem := icol else
          if NomCol = 'GF_DATEDEBUT' then SG_Datedeb := icol else
          if NomCol = 'GF_DATEFIN' then SG_Datefin := icol;
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));

  LesCols := G_QCa.Titres[0];
  LesColQCa := LesCols;
  icol := 0;
  repeat
    NomCol := uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if NomCol = 'GF_DEPOT' then SC_Depot := icol else
          if NomCol = 'GF_TARIFTIERS' then SC_Cat := icol else
          if NomCol = 'GF_LIBELLE' then SC_Lib := icol else
          if NomCol = 'GF_BORNEINF' then SC_QInf := icol else
          if NomCol = 'GF_BORNESUP' then SC_QSup := icol else
          if NomCol = 'GF_PRIXUNITAIRE' then SC_Px := icol else
          if NomCol = 'GF_CALCULREMISE' then SC_Rem := icol else
          if NomCol = 'GF_DATEDEBUT' then SC_Datedeb := icol else
          if NomCol = 'GF_DATEFIN' then SC_Datefin := icol;
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));

  LesCols := G_Cat.Titres[0];
  LesColCat := LesCols;
  icol := 0;
  repeat
    NomCol := uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if NomCol = 'GF_DEPOT' then SC2_Depot := icol else
          if NomCol = 'GF_TARIFTIERS' then SC2_Cat := icol else
          if NomCol = 'GF_LIBELLE' then SC2_Lib := icol else
          if NomCol = 'GF_PRIXUNITAIRE' then SC2_Px := icol else
          if NomCol = 'GF_CALCULREMISE' then SC2_Rem := icol else
          if NomCol = 'GF_DATEDEBUT' then SC2_Datedeb := icol else
          if NomCol = 'GF_DATEFIN' then SC2_Datefin := icol;
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));
end;

procedure TFTarifArticleMode.FormateZoneSaisie(ACol, ARow: Longint; ttd: T_TableTarif);
var St, StC: string;
begin
  inherited;
  if ttd = ttdArt then
  begin
    St := G_Art.Cells[ACol, ARow];
    StC := St;
    if (ACol = SA_Article) or (ACol = SA_TarifArticle) then StC := uppercase(Trim(St));
    G_Art.Cells[ACol, ARow] := StC;
  end;
end;

function TFTarifArticleMode.GrilleModifie: Boolean;
begin
  Result := False;
  if Action = taConsult then Exit;
  Result := (TOBTarfArt.IsOneModifie) or (TOBTarfQte.IsOneModifie) or (TOBTarfQCa.IsOneModifie) or
    (TOBTarfCat.IsOneModifie) or (TOBTarifDel.IsOneModifie);
  if (LigneVide(1, ttdArt)) then Result := False;
end;

function TFTarifArticleMode.ChampModifie(ARow: Integer): Boolean;
var TOBT: Tob;
begin
  Result := False;
  if (Action = taConsult) or (ARow = 0) then Exit;
  TOBT := TOBTarfArt.Detail[ARow - 1];
  Result := (TOBT.IsFieldModified('GF_ARTICLE')) or (TOBT.IsFieldModified('GF_TARIFARTICLE')) or
    (TOBT.IsFieldModified('GF_PRIXUNITAIRE')) or (TOBT.IsFieldModified('GF_CALCULREMISE'));
  if (LigneVide(1, ttdArt)) then Result := False;
end;

procedure TFTarifArticleMode.FormateGrilleSaisie(Promo: Boolean);
begin
  if not Promo then
  begin
    G_Art.ColLengths[Sa_TarifArticle] := -1;
    G_Art.ColWidths[Sa_TarifArticle] := 0;
    G_Art.ColLengths[Sa_Rem] := -1;
    G_Art.ColWidths[Sa_Rem] := 0;
    Hmtrad.ResizeGridColumns(G_Art);
  end else
  begin
    G_Art.ColLengths[Sa_TarifArticle] := 0;
    G_Art.ColWidths[Sa_TarifArticle] := 300;
    G_Art.ColLengths[Sa_Rem] := 0;
    G_Art.ColWidths[Sa_Rem] := 300;
    Hmtrad.ResizeGridColumns(G_Art);
  end;
end;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

procedure TFTarifArticleMode.LoadLesTOB;
var Q: TQuery;
  WhereTarifMode, WhereTTC: string;
begin
  // Inherited ;
  if ModifMul then exit;
  if TarifTTC then WhereTTC := ' AND GF_REGIMEPRIX = "TTC" '
  else WhereTTC := ' AND GF_REGIMEPRIX <> "TTC" ';
  WhereTarifMode := ' AND GF_TARFMODE = "' + CodeTarifMode + '"';
  if TARIFARTICLE <> '' then
  begin
    {  for i_ind := TOBTarfArt.Detail.Count - 1 downto 0 do
          BEGIN
          TOBTarfArt.Detail[i_ind].Free;
          END; }
    Q := OpenSQL('SELECT * FROM TARIF WHERE GF_TARIFARTICLE="' + TarifArticle + '" And GF_DEVISE="' + CodeDevise + '"' +
      WhereTarifMode + 'AND GF_DEPOT="' + CodeDepot + '" AND GF_REGIMEPRIX = "' + RegimePrix + '"'
      + 'AND GF_TIERS="" AND GF_TARIFTIERS="" ORDER BY GF_DEPOT, GF_BORNEINF', True);
    TOBTarfArt.LoadDetailDB('TARIF', '', '', Q, True);
    Ferme(Q);
  end;
  if CodeArticle <> '' then
  begin
    {for i_ind := TOBTarfArt.Detail.Count - 1 downto 0 do
        BEGIN
        TOBTarfArt.Detail[i_ind].Free;
        END; }
    Q := OpenSQL('SELECT * FROM TARIF WHERE ' + WhereTarifArt(CodeArticle, CodeDevise, ttdArt, False) +
      WhereTTC + WhereTarifMode + ' AND GF_DEPOT="' + CodeDepot + '" ORDER BY GF_DEPOT, GF_BORNEINF', True);
    TOBTarfArt.LoadDetailDB('TARIF', '', '', Q, True);
    if LigneVide(1, ttdArt) then SupprimeLigne(1, ttdArt);
    Ferme(Q);
  end;
end;

procedure TFTarifArticleMode.TraiterArticle(ACol, ARow: Integer);
var RechArt: T_RechArt;
  OkArt: Boolean;
  i_Rep, i_ind: Integer;
  ioerr: TIOErr;
  SQL: string;
  Q: TQuery;
begin
  if G_Art.Cells[ACol, ARow] = '' then
  begin
    //InitialiseEntete ;
    //InitialiseGrille;
    //DepileTOBLigne;
    //LibereTobMode ;
    exit
  end;
  OkArt := False;
  // Interdire la saisie d'un article existant sur une nouvelle ligne
  if TobArtExist <> nil then
  begin
    if TobArtExist.FindFirst(['GA_ARTICLE'], [G_Art.Cells[ACol, ARow]], False) <> nil then
      if (Valeur(G_Art.Cells[Sa_PX, ARow]) = 0) and (Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0) then
      begin
        G_Art.Cells[ACol, ARow] := '';
        exit;
      end;
  end;
  RechArt := TrouverArticle(G_Art.Cells[ACol, ARow], TOBArt);
  case RechArt of
    traOk: OkArt := True;
    traAucun:
      begin
        // Recherche sur code via LookUp ou Recherche avancée
        G_Art.Cells[ACol, ARow] := DispatchRecherche(GF_CODEARTICLE, 1, '',
          'GA_CODEARTICLE=' + Trim(Copy(G_Art.Cells[ACol, ARow], 1, 18)), '');
        if G_Art.Cells[ACol, ARow] <> '' then
        begin
          Okart := TOBArt.SelectDB('"' + G_Art.Cells[ACol, ARow] + '"', nil);
        end;
      end;
    traGrille:
      begin
        // Forcement objet dimension avec saisie obligatoire
        Okart := True;
      end;
  end; // Case
  if (Okart) then
  begin
    if (CodeArticle <> TOBArt.GetValue('GA_ARTICLE')) or (ModifMul) then
    begin
      if (Valeur(G_Art.Cells[Sa_PX, ARow]) = 0) or (Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0) then i_Rep := 7
      else i_Rep := QuestionTarifEnCours;
      case i_Rep of
        mrYes:
          begin
            ioerr := Transactions(ValideTarif, 2);
            case ioerr of
              oeOk: ;
              oeUnknown:
                begin
                  MessageAlerte(HMess.Mess[1]);
                end;
              oeSaisie:
                begin
                  MessageAlerte(HMess.Mess[2]);
                end;
            end;
            CodeArticle := G_Art.Cells[ACol, ARow];
            Transactions(ChargeTarif, 1);
          end;
        mrNo:
          begin
            CodeArticle := TOBArt.GetValue('GA_ARTICLE');
            if (_PERTARIFMODE.Value <> '') then
              SQL := 'SELECT * FROM TARIF WHERE ' + WhereTarifArt(CodeArticle, CodeDevise, ttdArt, False) +
                'AND GF_DEPOT="' + CodeDepot + '" AND GF_REGIMEPRIX = "' + RegimePrix + '" AND GF_TARFMODE="' + CodeTarifMode +
                '" AND GF_FERME="-" ORDER BY GF_DEPOT, GF_BORNEINF'
            else SQL := 'SELECT * FROM TARIF WHERE ' + WhereTarifArt(CodeArticle, CodeDevise, ttdArt, False) +
              'AND GF_DEPOT="' + CodeDepot + '" AND GF_REGIMEPRIX = "' + RegimePrix + '" AND GF_FERME="-" ORDER BY GF_DEPOT, GF_BORNEINF';
            if ExisteSQL(SQL) then
            begin
              Q := OpenSQL(SQL, True);
              //CodeArticle := G_Art.Cells[ACol,ARow] ;
              G_Art.Cells[SA_Px, ARow] := Q.FindField('GF_PRIXUNITAIRE').AsString;
              G_Art.Cells[SA_Rem, Arow] := Q.FindField('GF_CALCULREMISE').AsString;
              CodeTarif := StrToInt(Q.FindField('GF_TARIF').AsString);
              if not ValidationTarifOk then Transactions(ChargeTarif, 1);
              ferme(Q);
              if TobTarfArt.Detail.count <> 0 then //AC TarifDim
              begin
                if (TOBArt.GetValue('GA_STATUTART') = 'GEN') then
                begin
                  TobTarfArtDim := TOB.Create('', nil, -1);
                  ; // AC ArtDim
                  BAffDim.Visible := True;
                  if (TOBArt.GetValue('GA_PRIXUNIQUE') = 'X') then PTITRE.Caption := HTitre.Mess[9]
                  else PTITRE.Caption := HTitre.Mess[8];
                  for i_ind := TobTarfArtDim.Detail.Count - 1 downto 0 do
                  begin
                    TobTarfArtDim.Detail[i_ind].Free;
                  end;
                end;
              end;
            end else
            begin
              ValidationTarifOk := False;
              CreerTOBligne(ARow, ttdArt);
              InitLaLigne(ARow, ttdArt);
              CodeTarif := 0;
            end;
          end;
        mrCancel:
          begin
            TOBArt.SelectDB('"' + CodeArticle + '"', nil);
          end;
      end; // Case
    end else
    begin
      //InitialiseGrille;
     // DepileTOBLigne;
     // LibereTobMode ;
      codeArticle := '';
    end;
  end else
  begin
    //InitialiseEntete ;
  end;
end;

procedure TFTarifArticleMode.TraiterTarifArticle(ACol, ARow: integer);
var TOBL: TOB;
  St, SQL: string;
  i_Rep: Integer;
  ioerr: TIOErr;
  Q: TQuery;
begin
  // Interdire la saisie d'une catégorie article existante sur une nouvelle ligne
  if TobArtExist <> nil then
  begin
    if TobArtExist.FindFirst(['GA_TARIFARTICLE'], [G_Art.Cells[ACol, ARow]], False) <> nil then
      if (Valeur(G_Art.Cells[Sa_PX, ARow]) = 0) and (Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0) then
      begin
        G_Art.Cells[ACol, ARow] := '';
        exit;
      end;
  end;
  if (Valeur(G_Art.Cells[Sa_PX, ARow]) = 0) or (Valeur(G_Art.Cells[Sa_Rem, ARow]) = 0) then i_Rep := 7
  else i_Rep := QuestionTarifEnCours;
  case i_Rep of
    mrYes:
      begin
        ioerr := Transactions(ValideTarif, 2);
        ;
        case ioerr of
          oeOk: ;
          oeUnknown:
            begin
              MessageAlerte(HMess.Mess[1]);
            end;
          oeSaisie:
            begin
              MessageAlerte(HMess.Mess[2]);
            end;
        end;
        TarifArticle := G_Art.Cells[ACol, ARow];
        Transactions(ChargeTarif, 1);
      end;
    mrNo:
      begin
        TarifArticle := G_Art.Cells[ACol, ARow];
        if (_PERTARIFMODE.Value <> '') then
          SQL := 'SELECT * FROM TARIF WHERE GF_TARIFARTICLE="' + G_Art.Cells[SA_TarifArticle, Arow] + '" And GF_DEVISE="' + CodeDevise +
            '" AND GF_DEPOT="' + CodeDepot + '" AND GF_REGIMEPRIX = "' + RegimePrix + '" AND GF_TARFMODE="' + CodeTarifMode +
            '" AND GF_TIERS="" AND GF_TARIFTIERS="" ORDER BY GF_DEPOT, GF_BORNEINF'
        else SQL := 'SELECT * FROM TARIF WHERE GF_TARIFARTICLE="' + G_Art.Cells[SA_TarifArticle, Arow] + '" And GF_DEVISE="' + CodeDevise + '"'
          + ' AND GF_DEPOT="' + CodeDepot + '" AND GF_REGIMEPRIX = "' + RegimePrix + '"'
            + 'AND GF_TIERS="" AND GF_TARIFTIERS="" ORDER BY GF_DEPOT, GF_BORNEINF';
        if ExisteSQL(SQL) then
        begin
          Q := OpenSQL(SQL, True);
          TarifArticle := G_Art.Cells[ACol, ARow];
          G_Art.Cells[SA_Px, ARow] := Q.FindField('GF_PRIXUNITAIRE').AsString;
          G_Art.Cells[SA_Rem, Arow] := Q.FindField('GF_CALCULREMISE').AsString;
          CodeTarif := StrToInt(Q.FindField('GF_TARIF').AsString);
          //ValidationTarifOk:=False ;
          if not ValidationTarifOk then Transactions(ChargeTarif, 1);
          Ferme(Q);
        end else
        begin
          PTITRE.Caption := HTitre.Mess[10];
          CreerTOBligne(ARow, ttdArt);
          InitLaLigne(ARow, ttdArt);
          TOBL := GetTOBLigne(ARow, ttdArt);
          if TOBL = nil then exit;
          St := G_Art.Cells[ACol, ARow];
          if ExisteTablette('GCTARIFARTICLE', St) then
          begin
            TOBL.PutValue('GF_TARIFARTICLE', St);
          end else
          begin
            // message TarifArticle inexistant
            MsgBox.Execute(5, Caption, '');
            G_Art.Cells[ACol, ARow] := TOBL.GetValue('GF_TARIFARTICLE');
          end;
          CodeArticle := '';
          //InitLaLigne (ARow, ttdArt) ;
          CodeTarif := 0;
          ValidationTarifOk := False;
        end;
      end;
    mrCancel:
      begin
      end;
  end;
end;

procedure TFTarifArticleMode.TraiterEtablissement(ARow: integer);
var TOBL: TOB;
  TempDepot: string;
begin
  TOBL := GetTOBLigne(ARow, ttdArt);
  if TOBL = nil then exit;
  TempDepot := GF_DEPOT.Value;
  TOBL.PutValue('GF_DEPOT', TempDepot);
  TOBL.PutValue('GF_SOCIETE', TempDepot);
end;

procedure TFTarifArticleMode.MAJChampEntete;
begin
  if SavEntete then exit;
  ExecuteSQL('Update TarifTypMode SET GFT_DERUTILISE="-" Where GFT_DERUTILISE="X"');
  ExecuteSQL('Update TarifTypMode SET GFT_DERUTILISE="X" Where GFT_CODETYPE="' + CodeType + '"');
  ExecuteSQL('Update TarifPer SET GFP_DERUTILISE="-" Where GFP_DERUTILISE="X"');
  ExecuteSQL('Update TarifPer SET GFP_DERUTILISE="X" Where GFP_CODEPERIODE="' + CodePeriode + '"');
  SavEntete := True;
end;

procedure TFTarifArticleMode.TraiterRemise(ACol, ARow: integer; ttd: T_TableTarif);
var TOBL: TOB;
  St: string;
begin
  if ttd <> ttdArt then exit;
  TOBL := GetTOBLigne(ARow, ttd);
  if TOBL = nil then exit;
  St := G_Art.Cells[ACol, ARow];
  TOBL.PutValue('GF_CALCULREMISE', St);
  TOBL.PutValue('GF_REMISE', RemiseResultante(St));
  AffichePied(ttd);
end;

procedure TFTarifArticleMode.TraiterPrix(ACol, ARow: integer; ttd: T_TableTarif);
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(ARow, ttd);
  if TOBL = nil then exit;
  case ttd of
    ttdArt: TOBL.PutValue('GF_PRIXUNITAIRE', Valeur(G_Art.Cells[ACol, ARow]));
    ttdArtQte: TOBL.PutValue('GF_PRIXUNITAIRE', Valeur(G_Qte.Cells[ACol, ARow]));
    ttdArtQCa: TOBL.PutValue('GF_PRIXUNITAIRE', Valeur(G_QCa.Cells[ACol, ARow]));
    ttdArtCat: TOBL.PutValue('GF_PRIXUNITAIRE', Valeur(G_Cat.Cells[ACol, ARow]));
  end;
end;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}

function TFTarifArticleMode.LigneOk(ARow: Integer): Boolean;
begin
  Result := True;
  if (G_Art.Cells[SA_Px, ARow] = '') and (G_Art.Cells[SA_Rem, ARow] = '') and not (LigneVide(ARow, ttdArt)) then
  begin
    Result := False;
  end;
end;

procedure TFTarifArticleMode.BValiderClick(Sender: TObject);
var ioerr: TIOErr;
begin
  if Action = taConsult then exit;
  NouvelleLigne := False;
  // validation
  ioerr := Transactions(ValideTarif, 2);
  case ioerr of
    oeOk: ;
    oeUnknown:
      begin
        MessageAlerte(HMess.Mess[1]);
      end;
    oeSaisie:
      begin
        MessageAlerte(HMess.Mess[2]);
      end;
  end;
end;

procedure TFTarifArticleMode.ValideTarif;
var Arow, i_ind, j: Integer;
begin
  if NouvelleLigne then Arow := G_Art.Row - 1 else ARow := G_Art.Row;
  if Arow = 0 then exit;
  if not SortDeLaLigne(ttdArt) then Exit;
  if not SortDeLaLigne(ttdArtQte) then Exit;
  if not SortDeLaLigne(ttdArtQCa) then Exit;
  if not SortDeLaLigne(ttdArtCat) then Exit;
  if TOBTarifDel.detail.count > 0 then
  begin
    TOBTarifDel.DeleteDB(False);
    for i_ind := TOBTarifDel.Detail.Count - 1 downto 0 do
    begin
      TOBTarifDel.Detail[i_ind].Free;
    end;
    InitTarif;
    exit;
  end;
  if (G_Art.Cells[SA_Article, ARow] <> '') then
  begin
    if (TOBArt.GetValue('GA_STATUTART') = 'GEN') then
    begin
      BAffDim.Visible := True;
      if TOBArt.GetValue('GA_PRIXUNIQUE') = 'X' then PTITRE.Caption := HTitre.Mess[9]
      else PTITRE.Caption := HTitre.Mess[8];
    end else PTITRE.Caption := HTitre.Mess[7];
  end;
  if (G_Art.Cells[SA_TarifArticle, ARow] <> '') then PTITRE.Caption := HTitre.Mess[10];
  MAJChampEntete;
  if StrToDate(GF_DATEDEBUT.Text) > StrToDate(GF_DATEFIN.Text) then
  begin
    MsgBox.Execute(3, Caption, '');
    Exit;
  end else
    if (G_ART.Cells[SA_px, Arow] = '') and (G_ART.Cells[SA_rem, Arow] = '') then
  begin
    MsgBox.Execute(7, Caption, '');
    Exit;
  end;
  if (G_Art.cells[SA_Article, ARow] = '') and (G_Art.cells[SA_TarifArticle, ARow] = '') or ((G_ART.Cells[SA_px, Arow] = '') and (G_Art.Cells[Sa_Rem, Arow] =
    '')) then exit;
  LigneASav := ARow;
  VerifLesTOB;
  // Mise à jour de la date de modification
  MAJDateModif;
  TOBTarif.InsertOrUpdateDB(False);
  TOBTarifMode.InsertOrUpdateDB(False);
  ValidationTarifOK := True;
  SaisiPrix := False;
  CodeTarif := 0;
  if TobArtExist = nil then TobArtExist := TOB.Create('', nil, 0);
  j := TOBArtExist.Detail.count;
  TOB.Create('ARTICLE', TOBArtExist, j);
  if G_Art.cells[SA_Article, ARow] <> '' then TOBArtExist.Detail[j].PutValue('GA_ARTICLE', G_Art.Cells[SA_Article, ARow])
  else TOBArtExist.Detail[j].PutValue('GA_TARIFARTICLE', G_Art.Cells[Sa_TarifArticle, ARow]);
  tobtarfArt.SetAllModifie(False);
end;

procedure TFTarifArticleMode.MAJDateModif;
var i: Integer;
begin
  for i := 0 to TOBtarfArt.Detail.count - 1 do
  begin
    TOBtarfArt.Detail[i].PutValue('GF_DATEMODIF', NowH);
  end;
end;

procedure TFTarifArticleMode.VerifLesTOB;
var i_ind: Integer;
begin
  inherited;
  if (CodeTarif <> 0) then
  begin
    TOBTarfArt.Detail[LigneASav - 1].PutValue('GF_TARIF', CodeTarif);
  end;
  for i_ind := TOBTarfArt.Detail.count - 1 downto 0 do
  begin
    if not LigneOk(i_ind + 1) then
    begin
      TOBTarfArt.Detail[i_ind].Free;
      G_Art.Objects[0, i_ind + 1] := nil;
    end;
  end;
end;

procedure TFTarifArticleMode.SupDimPrixGen(ARow: Integer);
var i_ind, MaxTarif, i_del: Integer;
  Prix: Double;
  Remise: string;
  QMax: TQuery;
  TobTarfArtDimDel, TobTarfDim: Tob;
begin
  TobTarfArtDimDel := TOB.Create('', nil, -1);
  i_del := 0;
  i_ind := 0;
  Prix := Valeur(G_Art.Cells[Sa_Px, Arow]);
  Remise := G_Art.Cells[Sa_Rem, Arow];
  if (TobArt.GetValue('GA_PRIXUNIQUE') = 'X') then TOBTarfArtDim.ClearDetail;
  while TOBTarfArtDim.Detail.count > i_ind do
  begin
    TobTarfDim := TOBTarfArtDim.Detail[i_ind];
    if (Prix <> 0) then
    begin
      if (TobTarfDim.GetValue('GF_PRIXUNITAIRE') = Prix) then
      begin
        if TobTarfDim.GetValue('GF_TARIF') <> 0 then
        begin
          TOB.Create('TARIF', TobTarfArtDimDel, i_del);
          TobTarfArtDimDel.Detail[i_del].Dupliquer(TobTarfDim, False, True);
          i_del := i_del + 1;
        end;
        TobTarfDim.Free;
      end
      else i_ind := i_ind + 1
    end else
      if (Remise <> '') then
    begin
      if (TobTarfDim.GetValue('GF_CALCULREMISE') = Remise) then
      begin
        if TobTarfDim.GetValue('GF_TARIF') <> 0 then
        begin
          TOB.Create('TARIF', TobTarfArtDimDel, i_del);
          TobTarfArtDimDel.Detail[i_del].Dupliquer(TobTarfDim, False, True);
          i_del := i_del + 1;
        end;
        TobTarfDim.Free;
      end
      else i_ind := i_ind + 1;
    end else
      if (Remise = '') and (Prix = 0) then // Cas du tarif à la dimension
    begin
      if (TobTarfDim.GetValue('GF_PRIXUNITAIRE') = 0) and (TobTarfDim.GetValue('GF_CALCULREMISE') = '') then
      begin
        if TobTarfDim.GetValue('GF_TARIF') <> 0 then
        begin
          TOB.Create('TARIF', TobTarfArtDimDel, i_del);
          TobTarfArtDimDel.Detail[i_del].Dupliquer(TobTarfDim, False, True);
          i_del := i_del + 1;
        end;
        TobTarfDim.Free;
      end
      else i_ind := i_ind + 1;
    end;
  end;
  if TobTarfArtDimDel.Detail.Count > 0 then TobTarfArtDimDel.DeleteDB(False);
  QMax := OpenSQL('SELECT MAX(GF_TARIF) FROM TARIF', TRUE);
  if QMax.EOF then MaxTarif := 1 else MaxTarif := QMax.Fields[0].AsInteger + 1;
  for i_ind := 0 to TOBTarfArtDim.Detail.Count - 1 do
  begin
    if TOBTarfArtDim.Detail[i_ind].GetValue('GF_TARIF') = 0 then
    begin
      TOBTarfArtDim.Detail[i_ind].PutValue('GF_TARIF', MaxTarif);
      TOBTarfArtDim.Detail[i_ind].PutValue('GF_REMISE', Valeur(TOBTarfArtDim.Detail[i_ind].GetValue('GF_CALCULREMISE')));
      MaxTarif := MaxTarif + 1;
    end else
    begin
      TOBTarfArtDim.Detail[i_ind].PutValue('GF_REMISE', Valeur(TOBTarfArtDim.Detail[i_ind].GetValue('GF_CALCULREMISE')));
    end;
  end;
  ferme(QMax);
  TobTarfArtDimDel.Free;
end;

{==============================================================================================}
{========================= Manipulation des LIGNES Quantitatif ================================}
{==============================================================================================}

function TFTarifArticleMode.LigneVide(ARow: integer; ttd: T_TableTarif): Boolean;
begin
  Result := True;
  case ttd of
    ttdArt:
      begin
        if (G_Art.Cells[SA_Article, ARow] <> '') or (G_Art.Cells[SA_TarifArticle, ARow] <> '') then
        begin
          Result := False;
          Exit;
        end;
      end;
    ttdArtQte:
      begin
        if G_Qte.Cells[SG_Lib, ARow] <> '' then
        begin
          Result := False;
          Exit;
        end;
      end;
    ttdArtQCa:
      begin
        if (G_QCa.Cells[SC_Lib, ARow] <> '') and (G_QCa.Cells[SC_Cat, ARow] <> '') then
        begin
          Result := False;
          Exit;
        end;
      end;
    ttdArtCat:
      begin
        if (G_Cat.Cells[SC2_Lib, ARow] <> '') and (G_Cat.Cells[SC2_Cat, ARow] <> '') then
        begin
          Result := False;
          Exit;
        end;
      end;
  end;
end;

procedure TFTarifArticleMode.InitLaLigne(ARow: integer; ttd: T_TableTarif);
var TOBL: TOB;
  PrixConver: Double;
begin
  if ModifMul then exit;
  TOBL := GetTOBLigne(ARow, ttd);
  if TOBL = nil then Exit;
  TOBL.PutValue('GF_ARTICLE', CodeArticle);
  TOBL.PutValue('GF_TARIFARTICLE', TarifArticle);
  TOBL.PutValue('GF_BORNESUP', 999999);
  if CodeArticle <> '' then
  begin
    if NatureType = 'VTE' then PrixConver := (TOBArt.GetValue('GA_PVTTC') * CoefType)
    else PrixConver := (TOBArt.GetValue('GA_PAHT') * CoefType);
    PrixConver := ArrondirPrix(CodeArrondi, PrixConver);
    PrixConver := Arrondi(PrixConver, DEV.DECIMALE);
    TOBL.PutValue('GF_PRIXUNITAIRE', PrixConver);
  end else
    TOBL.PutValue('GF_PRIXUNITAIRE', 0);
  TOBL.PutValue('GF_REMISE', 0);
  TOBL.PutValue('GF_CALCULREMISE', '');
  TOBL.PutValue('GF_MODECREATION', 'MAN');
  TOBL.PutValue('GF_ARRONDI', CodeArrondi);
  //if GF_CASCADEREMISE.Value <> '' then
  //    TOBL.PutValue ('GF_CASCADEREMISE', GF_CASCADEREMISE.Value)
  //    else TOBL.PutValue ('GF_CASCADEREMISE', 'MIE');
  TOBL.PutValue('GF_CASCADEREMISE', CodeRemise);
  TOBL.PutValue('GF_DEVISE', CodeDevise);
  TOBL.PutValue('GF_DATEDEBUT', StrToDate(GF_DATEDEBUT.Text));
  TOBL.PutValue('GF_DATEFIN', StrToDate(GF_DATEFIN.Text));
  TOBL.PutValue('GF_DEMARQUE', CodeDemarque);
  TOBL.PutValue('GF_QUALIFPRIX', 'GRP');
  TOBL.PutValue('GF_FERME', '-');
  TraiterEtablissement(ARow);
  TOBL.PutValue('GF_TARFMODE', CodeTarifMode);
  TOBL.PutValue('GF_LIBELLE', Copy(_TYPTARIFMODE.Text + ' - ' + _PERTARIFMODE.Text, 1, 35));
  if TarifTTC then
  begin
    TOBL.PutValue('GF_REGIMEPRIX', 'TTC');
    TOBL.PutValue('GF_NATUREAUXI', 'CLI');
  end
  else
  begin
    TOBL.PutValue('GF_REGIMEPRIX', 'HT');
    TOBL.PutValue('GF_NATUREAUXI', 'FOU');
  end;
  case ttd of
    ttdArt:
      begin
        TOBL.PutValue('GF_BORNEINF', -999999);
        TOBL.PutValue('GF_QUANTITATIF', '-');
      end;
    ttdArtQte:
      begin
        TOBL.PutValue('GF_BORNEINF', 1);
        TOBL.PutValue('GF_QUANTITATIF', 'X');
      end;
    ttdArtQCa:
      begin
        TOBL.PutValue('GF_BORNEINF', 1);
        TOBL.PutValue('GF_QUANTITATIF', 'X');
      end;
    ttdArtCat:
      begin
        TOBL.PutValue('GF_BORNEINF', -999999);
        TOBL.PutValue('GF_QUANTITATIF', '-');
      end;
  end;
  AfficheLaligne(ARow, ttd);
end;

{==============================================================================================}
{=============================== Actions liées aux Dimensions =================================}
{==============================================================================================}

procedure TFTarifArticleMode.GereTobDim(Arow: Integer); // AC
var MaxTarif, j: Integer;
  QArtDim: TQuery;
begin
  j := 0;
  TobTarfArtDim := TOB.Create('', nil, -1);
  MaxTarif := TOBTarfArt.Detail[Arow - 1].GetValue('GF_TARIF') + 1; // + TOBTarfArt.Detail.count;
  codeArticle := G_Art.Cells[SA_Article, ARow];
  begin
    QArtDim := OpenSql('Select GA_ARTICLE from Article where GA_CODEARTICLE="' + TRIM(copy(CodeArticle, 1, 18)) + '" And GA_STATUTART="DIM"', True);
    while not QArtDim.EOF do
    begin
      TOB.Create('TARIF', TobTarfArtDim, j);
      TobTarfArtDim.Detail[j].Dupliquer(TOBTarfArt.Detail[Arow - 1], False, True);
      TobTarfArtDim.Detail[j].PutValue('GF_ARTICLE', QArtDim.FindField('GA_ARTICLE').AsString);
      TobTarfArtDim.Detail[j].PutValue('GF_TARIF', MaxTarif);
      QArtDim.next;
      MaxTarif := MaxTarif + 1;
      j := j + 1;
    end;
  end;
  ferme(QArtDim);
end;

procedure TFTarifArticleMode.AfficheDim(Article: string); //AC
var CoordonneEcran: TPoint;
  Top, Left, Height, Width, compare, i: Integer;
  TobTarfArtDimSav: TOB;
  ChampSaisi: string;
begin
  if valeur(G_Art.Cells[SA_Rem, G_Art.Row]) <> 0 then ChampSaisi := 'GF_CALCULREMISE' else
    ChampSaisi := 'GF_PRIXUNITAIRE';
  // Cas Tarif à la dimension
  if (valeur(G_Art.Cells[SA_Rem, G_Art.Row]) = 0) and (valeur(G_Art.Cells[SA_Px, G_Art.Row]) = 0) then
    if (TobTarfArtDim <> nil) and (TobTarfArtDim.Detail[0].GetValue('GF_PRIXUNITAIRE') = 0) then
      ChampSaisi := 'GF_CALCULREMISE' else ChampSaisi := 'GF_PRIXUNITAIRE';
  //
  TobTarfArtDimSav := TOB.Create('TARIF', nil, -1);
  if TobTarfArtDim <> nil then TobTarfArtDimSav.Dupliquer(TobTarfArtDim, True, True);
  TheTob := TobTarfArtDim;
  CoordonneEcran := RetourneCoordonneeCellule(0, G_Art);
  Top := CoordonneEcran.y;
  Left := CoordonneEcran.x;
  Height := G_Art.Height - G_Art.RowHeights[0] - G_Art.RowHeights[1];
  Width := G_Art.Width - G_Art.ColWidths[0];
  V_PGI.FormCenter := False;
  AglLanceFiche('GC', 'GCMSELECTDIMDOC', '', '', 'GA_CODEARTICLE=' + Trim(copy(Article, 1, 18)) + ';ACTION=SAISIE;TOP=' + IntToStr(Top) + ';LEFT=' +
    IntToStr(Left) + ';HEIGTH=' + IntToStr(Height) + ';WIDTH=' + IntToStr(Width) + ';TYPEPARAM=TAF;CHAMP=' + ChampSaisi + '');
  TobTarfArtDim := TheTob;
  for i := 0 to TobTarfArtDim.Detail.count - 1 do
  begin
    Compare := CompareTOB(TobTarfArtDim.Detail[i], TobTarfArtDimSav.Detail[i], 'GF_PRIXUNITAIRE;GF_CALCULREMISE');
    if Compare <> 0 then
    begin
      SupDimPrixGen(G_Art.row);
      TobTarfArtDim.InsertOrUpdateDB;
      break;
    end;
  end;
  TheTob := nil;
  TobTarfArtDimSav.Free;
end;

procedure TFTarifArticleMode.MiseAJourDim(ARow: Integer);
var QArtDim, QTarifDim: TQuery;
  i_ind, j: Integer;
begin
  j := 0;
  if TobTarfArtDim = nil then TobTarfArtDim := TOB.Create('', nil, -1);
  for i_ind := TobTarfArtDim.Detail.Count - 1 downto 0 do
  begin
    TobTarfArtDim.Detail[i_ind].Free;
  end;
  QArtDim := OpenSql('Select GA_ARTICLE from Article where GA_CODEARTICLE="' + Trim(copy(G_Art.Cells[Sa_Article, G_Art.Row], 1, 18)) +
    '" And GA_STATUTART="DIM"', True);
  while not QartDim.EOF do
  begin
    QTarifDim := OpenSQL('SELECT * FROM TARIF WHERE GF_ARTICLE="' + QArtDim.FindField('GA_ARTICLE').AsString + '" AND GF_TARFMODE="' + CodeTarifMode + '"' +
      'AND GF_DEPOT="' + CodeDepot + '" AND GF_FERME="-"', True);
    if not QTarifDim.EOF then TobTarfArtDim.LoadDetailDB('TARIF', '', '', QTarifDim, True, False)
    else
    begin
      TOB.Create('TARIF', TobTarfArtDim, j);
      TobTarfArtDim.Detail[j].Dupliquer(TOBTarfArt.Detail[Arow - 1], False, True);
      TobTarfArtDim.Detail[j].PutValue('GF_ARTICLE', QArtDim.FindField('GA_ARTICLE').AsString);
      TobTarfArtDim.Detail[j].PutValue('GF_TARIF', 0);
      TobTarfArtDim.Detail[j].PutValue('GF_PRIXUNITAIRE', TOBTarfArt.Detail[ARow - 1].GetValue('GF_PRIXUNITAIRE'));
      TobTarfArtDim.Detail[j].PutValue('GF_CALCULREMISE', TOBTarfArt.Detail[ARow - 1].GetValue('GF_CALCULREMISE'));
      TobTarfArtDim.Detail[j].PutValue('GF_REMISE', TOBTarfArt.Detail[ARow - 1].GetValue('GF_REMISE'));
      TobTarfArtDim.Detail[j].PutValue('GF_ARRONDI', TOBTarfArt.Detail[ARow - 1].GetValue('GF_ARRONDI'));
    end;
    QArtDim.next;
    j := j + 1;
    ferme(QTarifDim);
  end;
  ferme(QArtDim);
end;

procedure TFTarifArticleMode.InitTarif;
begin
  if not SortDeLaLigne(ttdArt) then Exit;
  if not SortDeLaLigne(ttdArtQte) then Exit;
  if not SortDeLaLigne(ttdArtQCa) then Exit;
  if not SortDeLaLigne(ttdArtCat) then Exit;
  InitialiseGrille;
  DepileTOBLigne;
  LibereTOBMode;
  InitialiseEntete;
end;

procedure TFTarifArticleMode.LibereTOBMode;
begin
  TobTarifMode.Free;
  TobTarifMode := nil;
  if TobTarfArtDim <> nil then
  begin
    TobTarfArtDim.Free;
    TobTarfArtDim := nil;
  end;
end;

{$R *.DFM}
// Evènement bouton

procedure TFTarifArticleMode.BAffDimClick(Sender: TObject);
var i_ind: Integer;
begin
  MiseAjourDim(G_Art.Row);
  AfficheDim(G_Art.Cells[Sa_Article, G_Art.Row]);
  for i_ind := TobTarfArtDim.Detail.Count - 1 downto 0 do
  begin
    TobTarfArtDim.Detail[i_ind].Free;
  end;
end;

procedure TFTarifArticleMode.VoirFicheArticle;
begin
  V_PGI.DispatchTT(7, taConsult, CodeArticle, 'ACTION=CONSULTATION;TARIF=N', '');
  //AglLanceFiche ('MBO', 'ARTICLE', '', CodeArticle, 'ACTION=CONSULTATION;TARIF=N');
end;

procedure TFTarifArticleMode.BNvPeriodeClick(Sender: TObject);
begin
  if Action = taConsult then Exit; //NA : 20/09/2002
  AglLanceFiche('MBO', 'TARIFPER', '', '', 'ACTION=CREATION');
end;

procedure TFTarifArticleMode.BNvTypeClick(Sender: TObject);
begin
  if Action = taConsult then Exit; //NA : 20/09/2002
  AglLanceFiche('MBO', 'TARIFTYPE', '', '', 'TYPE=' + NatureType + ';ACTION=CREATION');
end;

procedure TFTarifArticleMode.BZoomArticle(Sender: TObject);
begin
  if (G_Art.Cells[Sa_Article, G_Art.Row] = '') or (CodeArticle = '') then Exit;
  VoirFicheArticle;
end;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}

procedure TFTarifArticleMode._TYPTARIFMODEExit(Sender: TObject);
var QQ: TQuery;
  Etablissement: string;
begin
  CodeType := _TYPTARIFMODE.Value;
  CoefType := 1;
  if CodeType = '' then
  begin
    CodeType := '...';
    _TYPTARIFMODE.VALUE := '...';
  end;
  if (CodeType = '...') then
  begin
    GF_DEPOT.Visible := False;
    TGF_DEPOT.Visible := False;
    GF_DEPOT.Value := '';
    CodeDepot := '';
  end else
  begin
    GF_DEPOT.Visible := True;
    TGF_DEPOT.Visible := True;
  end;
  QQ := OpenSql('Select GFT_DEVISE, GFT_ETABLISREF,GFT_COEF from TARIFTYPMODE where GFT_CODETYPE="' + CodeType + '" and GFT_NATURETYPE="' + NatureType + '"',
    True);
  if not QQ.EOF then
  begin
    CodeDevise := QQ.FindField('GFT_DEVISE').AsString;
    GF_DEVISE_.Text := RechDom('TTDEVISE', CodeDevise, False);
    DEV.Code := CodeDevise;
    GetInfosDevise(DEV);
    //GF_DEPOT.Value:=QQ.FindField('GFT_ETABLISREF').AsString ;
    Etablissement := QQ.FindField('GFT_ETABLISREF').AsString;
    CoefType := QQ.FindField('GFT_COEF').AsFloat;
  end else
  begin
    GF_DEVISE_.Text := RechDom('TTDEVISE', GetParamSoc('SO_DEVISEPRINC'), False);
    CodeDevise := GetParamSoc('SO_DEVISEPRINC');
  end;
  ferme(QQ);
  if NatureType = 'VTE' then GF_DEPOT.Plus := '(ET_TYPETARIF = "' + CodeType + '")'
  else GF_DEPOT.Plus := '(ET_TYPETARIFACH = "' + CodeType + '")';
  GF_DEPOT.Value := ReadTokenSt(Etablissement);
  TTYPETARIF.Caption := HTitre.Mess[6] + RechDom('TTDEVISE', CodeDevise, False);
  DEV.Code := CodeDevise;
  GetInfosDevise(DEV);
  G_Art.Enabled := True;
end;

procedure TFTarifArticleMode._PERTARIFMODEExit(Sender: TOBject);
var Info: string;
begin
  CodePeriode := _PERTARIFMODE.Value;
  Info := RecupInfoPeriode(_PERTARIFMODE.Value, GF_DEPOT.Value);
  GF_DATEDEBUT.Text := ReadTokenSt(Info);
  GF_DATEFIN.Text := ReadTokenSt(Info);
  CodeArrondi := ReadTokenSt(Info);
  CodeRemise := ReadTokenSt(Info);
  CodeDemarque := ReadTokenSt(Info);
  Promo := StringToCheck(Info);
  GF_DEMARQUE.Text := RechDom('GCTYPEREMISE', CodeDemarque, False);
  if GF_DEMARQUE.Text = 'Error' then GF_DEMARQUE.Text := '';
  FormateGrilleSaisie(Promo);
end;

procedure TFTarifArticleMode.GF_DEPOTExit(Sender: TOBject);
var Info: string;
begin
  CodeDepot := GF_DEPOT.Value;
  Info := RecupInfoPeriode(_PERTARIFMODE.Value, CodeDepot);
  GF_DATEDEBUT.Text := ReadTokenSt(Info);
  GF_DATEFIN.Text := ReadTokenSt(Info);
  CodeArrondi := ReadTokenSt(Info);
  CodeRemise := ReadTokenSt(Info);
  CodeDemarque := ReadTokenSt(Info);
  GF_DEMARQUE.Text := RechDom('GCTYPEREMISE', CodeDemarque, False);
  if GF_DEMARQUE.Text = 'Error' then GF_DEMARQUE.Text := '';
end;

procedure TFTarifArticleMode._TYPTARIFMODEChange(Sender: TObject);
var i_Rep: integer;
  ioerr: TIOErr;
begin
  if InitEntete then exit;
  if (CodeType = _TYPTARIFMODE.Value) or (CodeType = '') or (CodeType = '...') then Exit;
  i_Rep := QuestionTarifEnCours;
  case i_Rep of
    mrYes:
      begin
        ioerr := Transactions(ValideTarif, 2);
        ;
        case ioerr of
          oeOk: ;
          oeUnknown:
            begin
              MessageAlerte(HMess.Mess[1]);
            end;
          oeSaisie:
            begin
              MessageAlerte(HMess.Mess[2]);
            end;
        end;
        InitTarif;
        CodeType := _TYPTARIFMODE.Value;
        Transactions(ChargeTarif, 1);
      end;
    mrNo:
      begin
        InitialiseGrille;
        DepileTOBLigne;
        LibereTOBMode;
        TarifExistant := False;
        SaisiPrix := False;
        CodetarifMode := '';
        CodeArticle := '';
        TarifArticle := '';
        CodeType := _TYPTARIFMODE.Value;
        Transactions(ChargeTarif, 1);
      end;
    mrCancel:
      begin
        _TYPTARIFMODE.Value := CodeType;
      end;
  end;
  InitialisePied;
end;

procedure TFTarifArticleMode._PERTARIFMODEChange(Sender: TObject);
var i_Rep: integer;
  ioerr: TIOErr;
begin
  if InitEntete then exit;
  if (CodePeriode = _PERTARIFMODE.Value) or (CodePeriode = '') then Exit;
  i_Rep := QuestionTarifEnCours;
  case i_Rep of
    mrYes:
      begin
        ioerr := Transactions(ValideTarif, 2);
        ;
        case ioerr of
          oeOk: ;
          oeUnknown:
            begin
              MessageAlerte(HMess.Mess[1]);
            end;
          oeSaisie:
            begin
              MessageAlerte(HMess.Mess[2]);
            end;
        end;
        InitTarif;
        CodePeriode := _PERTARIFMODE.Value;
        Transactions(ChargeTarif, 1);
      end;
    mrNo:
      begin
        InitialiseGrille;
        DepileTOBLigne;
        LibereTOBMode;
        TarifExistant := False;
        SaisiPrix := False;
        CodetarifMode := '';
        CodeArticle := '';
        TarifArticle := '';
        CodePeriode := _PERTARIFMODE.Value;
        Transactions(ChargeTarif, 1);
      end;
    mrCancel: _PERTARIFMODE.Value := CodePeriode;
  end;
  InitialisePied;
end;

procedure TFTarifArticleMode.GF_DEPOTChange(Sender: TObject);
var i_Rep: integer;
  ioerr: TIOErr;
begin
  if InitEntete then exit;
  if (CodeDepot = GF_DEPOT.Value) then Exit;
  i_Rep := QuestionTarifEnCours;
  case i_Rep of
    mrYes:
      begin
        ioerr := Transactions(ValideTarif, 2);
        ;
        case ioerr of
          oeOk: ;
          oeUnknown:
            begin
              MessageAlerte(HMess.Mess[1]);
            end;
          oeSaisie:
            begin
              MessageAlerte(HMess.Mess[2]);
            end;
        end;
        InitTarif;
        CodeDepot := GF_DEPOT.Value;
        Transactions(ChargeTarif, 1);
      end;
    mrNo:
      begin
        InitialiseGrille;
        DepileTOBLigne;
        LibereTOBMode;
        CodeDepot := GF_DEPOT.Value;
        Transactions(ChargeTarif, 1);
      end;
    mrCancel: GF_DEPOT.Value := CodeDepot;
  end;
  InitialisePied;
end;

procedure TFTarifArticleMode.InitialiseEntete;
var QPer, QTyp: TQuery;
  info, Etablissement: string;
begin
  inherited;
  SavEntete := False;
  InitEntete := True;
  // modif le 13/08/02 on parle d'établissement et non de dépôt
  GF_DEPOT.Value := VH^.EtablisDefaut; //VH_GC.GCDepotDefaut ;
  CodeDepot := GF_DEPOT.Value;
  Etablissement := '';
  CoefType := 1;
  QTyp := OpenSQL('Select GFT_DEVISE,GFT_CODETYPE,GFT_ETABLISREF,GFT_COEF from TarifTypMode where GFT_DERUTILISE="X" AND GFT_NATURETYPE="' + NatureType + '"',
    True);
  if not QTyp.EOF then
  begin
    CodeType := QTyp.FindField('GFT_CODETYPE').AsString;
    if CodeType = '...' then
    begin
      GF_DEPOT.Visible := False;
      TGF_DEPOT.Visible := False;
      GF_DEPOT.Value := '';
      _TYPTARIFMODE.Value := '...';
    end;
    _TYPTARIFMODE.Value := CodeType;
    CodeDevise := QTyp.FindField('GFT_DEVISE').AsString;
    GF_DEVISE_.Text := RechDom('TTDEVISE', CodeDevise, False);
    TTYPETARIF.Caption := HTitre.Mess[6] + RechDom('TTDEVISE', CodeDevise, False);
    Etablissement := QTyp.FindField('GFT_ETABLISREF').AsString;
    CoefType := QTyp.FindField('GFT_COEF').AsFloat;
  end else
  begin
    GF_DEVISE_.Text := RechDom('TTDEVISE', GetParamSoc('SO_DEVISEPRINC'), False);
    CodeDevise := GetParamSoc('SO_DEVISEPRINC');
    GF_DEPOT.Visible := False;
    TGF_DEPOT.Visible := False;
    GF_DEPOT.Value := '';
    CodeDepot := GF_DEPOT.Value;
  end;
  QPer := OpenSQL('Select * from TarifPer where GFP_DERUTILISE="X"', True);
  if not QPer.EOF then
  begin
    CodePeriode := QPer.FindField('GFP_CODEPERIODE').AsString;
    _PERTARIFMODE.Value := CodePeriode;
    GF_DATEDEBUT.Text := QPer.FindField('GFP_DATEDEBUT').AsString;
    GF_DATEFIN.Text := QPer.FindField('GFP_DATEFIN').AsString;
    CodeArrondi := QPer.FindField('GFP_ARRONDI').AsString;
    CodeDemarque := QPer.FindField('GFP_DEMARQUE').AsString;
    if CodeDemarque <> '' then GF_DEMARQUE.Text := RechDom('GCTYPEREMISE', CodeDemarque, False);
    if GF_DEMARQUE.Text = 'Error' then GF_DEMARQUE.Text := '';
  end else
  begin
    CodePeriode := '';
    _PERTARIFMODE.Value := '';
    GF_DATEDEBUT.Text := '01/01/1900';
    GF_DATEFIN.Text := '31/12/2099';
  end;
  Ferme(QTyp);
  Ferme(QPer);
  if TarifTTC then GF_DEPOT.Plus := '(ET_TYPETARIF = "' + CodeType + '")'
  else GF_DEPOT.Plus := '(ET_TYPETARIFACH = "' + CodeType + '")';
  GF_DEPOT.Value := ReadTokenSt(Etablissement);
  CodeDepot := GF_DEPOT.Value;
  CodeArticle := '';
  SaisiPrix := False;
  TarifExistant := False;
  G_ART.Enabled := True;
  InitEntete := False;
  Info := RecupInfoPeriode(_PERTARIFMODE.Value, GF_DEPOT.Value);
  GF_DATEDEBUT.Text := ReadTokenSt(Info);
  GF_DATEFIN.Text := ReadTokenSt(Info);
  CodeArrondi := ReadTokenSt(Info);
  CodeRemise := ReadTokenSt(Info);
  CodeDemarque := ReadTokenSt(Info);
  Promo := StringToCheck(Info);
  GF_DEMARQUE.Text := RechDom('GCTYPEREMISE', CodeDemarque, False);
  if GF_DEMARQUE.Text = 'Error' then GF_DEMARQUE.Text := '';
  FormateGrilleSaisie(Promo);
end;

{==================================FONCTIONS ARTICLES =================================}

function TrouverArticle(Article: string; TOBArt: TOB): T_RechArt;
var Q: TQuery;
  Etat, CodeArticle: string;
begin
  Result := traAucun;
  if Article = '' then exit;
  CodeArticle := Trim(copy(Article, 1, 18));
  Q := OpenSQL('Select * from ARTICLE Where GA_CODEARTICLE = "' +
    CodeArticle + '" AND GA_STATUTART in ("GEN","UNI") ', True);
  if not Q.EOF then
  begin
    TOBArt.SelectDB('', Q);
    Etat := TOBArt.GetValue('GA_STATUTART');
    if Etat = 'UNI' then Result := traOk else
      if Etat = 'GEN' then Result := traGrille else Result := traOk;
  end;
  Ferme(Q);
end;

procedure TFTarifArticleMode.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  //inherited;
  if Action = taConsult then Exit;
  if not ValidationTarifOk and GrilleModifie then
  begin
    if MsgBox.Execute(6, Caption, '') <> mrYes then CanClose := False;
  end;
  if CanClose then InitTarif;
end;

procedure TFTarifArticleMode.G_ARTEnter(Sender: TObject);
var Cancel, Chg: Boolean;
  ACol, ARow: integer;
begin
  CodeTarif := 0;
  if _PERTARIFMODE.Value = '' then
  begin
    HTarif.Execute(1, Caption, '');
    _PERTARIFMODE.SetFocus;
    Exit;
  end;
  if _PERTARIFMODE.Value <> '' then
  begin
    TobTarifMode := TraiterTableTarifMode(_TYPTARIFMODE.Value, _PERTARIFMODE.Value, NatureType, TobTarifMode);
    CodeTarifMode := TobTarifMode.Getvalue('GFM_TARFMODE');
  end;
  Cancel := False;
  Chg := False;
  G_ArtRowEnter(Sender, G_Art.Row, Cancel, Chg);
  ACol := G_Art.Col;
  ARow := G_Art.Row;
  G_ARTCellEnter(Sender, ACol, ARow, Cancel);
end;

procedure TFTarifArticleMode.ChargeTarifDepuisMul;
var Q: TQuery;
  i_ind: integer;
  Select: string;
  TobInit: TOB;
begin
  TobTarifMode := TOB.Create('TARIFMODE', nil, -1);
  TobInit := TheTob;
  TheTob := nil;
  TOBArt.InitValeurs;
  IdTarif := TobInit.GetValue('_CodeTarif');
  CodeArticle := TobInit.GetValue('_CodeArticle');
  TarifArticle := TobInit.GetValue('_TarifArticle');
  if (CodeArticle <> '') then //And (TarifArticle='') then
  begin
    // Charge article
    TOBArt.SelectDB('"' + CodeArticle + '"', nil);

    // Lecture Quantitatif
    Select := 'SELECT * FROM TARIF WHERE GF_ARTICLE="' + CodeArticle + '" AND GF_TARIF="' + IdTarif + '" ' +
      ' ORDER BY GF_BORNEINF';
    Q := OpenSQL(Select, True);
    TOBTarfArt.LoadDetailDB('TARIF', '', '', Q, False);
    //TOBTarfArt.SetAllModifie(False);
    CodeTarifMode := Q.FindField('GF_TARFMODE').AsString;
    CodeDepot := Q.FindField('GF_DEPOT').AsString;
    Ferme(Q);
  end;
  if (TarifArticle <> '') and (TOBTarfArt.Detail.count = 0) then
  begin
    // Lecture Quantitatif
    Select := 'SELECT * FROM TARIF WHERE GF_TARIFARTICLE="' + TarifArticle + '" AND GF_TARIF="' + IdTarif + '" ' +
      ' ORDER BY GF_BORNEINF';
    Q := OpenSQL(Select, True);
    TOBTarfArt.LoadDetailDB('TARIF', '', '', Q, False);
    //TOBTarfArt.SetAllModifie(False);
    CodeTarifMode := Q.FindField('GF_TARFMODE').AsString;
    CodeDepot := Q.FindField('GF_DEPOT').AsString;
    Ferme(Q);
  end;
  // Affichage entête
  Q := OpenSql('Select * from TARIFMODE where GFM_TARFMODE="' + CodeTarifMode + '"', True);
  if not Q.EOF then
  begin
    TobTarifMode.SelectDB('', Q);
  end;
  ferme(Q);
  CodeType := TobTarifMode.GetValue('GFM_TYPETARIF');
  if CodeType = '...' then
  begin
    GF_DEPOT.Visible := False;
    TGF_DEPOT.Visible := False;
  end else
  begin
    GF_DEPOT.Visible := True;
    TGF_DEPOT.Visible := True;
  end;
  _TypTarifMode.Value := CodeType;
  CodePeriode := TobTarifMode.GetValue('GFM_PERTARIF');
  _PerTarifMode.Value := CodePeriode;
  GF_DEPOT.Value := CodeDepot;
  GF_DATEDEBUT.Text := TobTarifMode.GetValue('GFM_DATEDEBUT');
  GF_DATEFIN.Text := TobTarifMode.GetValue('GFM_DATEFIN');
  CodeDevise := TobTarifMode.GetValue('GFM_DEVISE');
  DEV.Code := CodeDevise;
  GetInfosDevise(DEV);
  GF_DEVISE_.Text := RechDom('TTDEVISE', CodeDevise, False);
  CodeArrondi := TobTarifMode.GetValue('GFM_ARRONDI');
  CodeDemarque := TobTarifMode.GetValue('GFM_DEMARQUE');
  if CodeDemarque <> '' then GF_DEMARQUE.Text := RechDom('GCTYPEREMISE', CodeDemarque, False);
  // Affichage
  for i_ind := 0 to TOBTarfArt.Detail.Count - 1 do
  begin
    AfficheLaLigne(i_ind + 1, ttdArt);
  end;
  if (TOBArt.GetValue('GA_STATUTART') = 'GEN') then
  begin
    BAffDim.Visible := True;
    if TOBArt.GetValue('GA_PRIXUNIQUE') = 'X' then PTITRE.Caption := HTitre.Mess[9]
    else PTITRE.Caption := HTitre.Mess[8];
  end else PTITRE.Caption := HTitre.Mess[7];
  if TarifArticle <> '' then PTITRE.Caption := HTitre.Mess[10];
  ModifMul := True;
  InitialisePied;
  G_ART.SetFocus;
  //Interdire la modification de l'entête
  _TYPTARIFMODE.Enabled := False;
  _PERTARIFMODE.Enabled := False;
  GF_DEPOT.Enabled := False;
  //
  // bouton inutile visible
  BNvPeriode.Visible := False;
  BNvType.Visible := False;
  //
  if G_Art.Cells[Sa_Px, G_Art.Row] = '' then G_Art.Col := Sa_Rem;
  if G_Art.Cells[Sa_Rem, G_Art.Row] = '' then G_Art.Col := Sa_Px;
  ValidationTarifOk := True;
end;

procedure AGLEntreeTarifArticleMode(Parms: array of variant; nb: integer);
begin
  EntreeTarifArticleMode(taModif, TRUE);
end;

initialization
  RegisterAglProc('EntreeTarifArticleMode', False, 0, AGLEntreeTarifArticleMode);

end.
