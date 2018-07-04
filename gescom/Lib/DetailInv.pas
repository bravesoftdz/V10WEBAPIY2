unit DetailInv;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, HSysMenu, Menus, HTB97, StdCtrls, Mask, ExtCtrls, HPanel,
  UTOB,EntGC;

type
  TFDetailInv = class(TForm)
    PENTETE: THPanel;
    TCodeListe: THLabel;
    TLibelle: THLabel;
    TDepot: THLabel;
    HLabel1: THLabel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BChercher: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BInfos: TToolbarButton97;
    POPZ: TPopupMenu;
    InfArticle: TMenuItem;
    HMTrad: THSystemMenu;
    FindLigne: TFindDialog;
    G_Dtl: THGrid;
    TLibelleArticle: THLabel;
    HLabel2: THLabel;
    TUStock: THLabel;
    TGIL_EMPLACEMENT: THLabel;
    TEmplacement: THLabel;
    PDim: THPanel;
    TGF_GRILLEDIM1: THLabel;
    GF_CODEDIM1: THCritMaskEdit;
    GF_CODEDIM2: THCritMaskEdit;
    TGF_GRILLEDIM2: THLabel;
    TGF_GRILLEDIM3: THLabel;
    GF_CODEDIM3: THCritMaskEdit;
    GF_CODEDIM4: THCritMaskEdit;
    TGF_GRILLEDIM4: THLabel;
    TGF_GRILLEDIM5: THLabel;
    GF_CODEDIM5: THCritMaskEdit;
    POPU: TPopupMenu;
    MenuItem3: TMenuItem;
    SetSaisi: TMenuItem;
    SetNonSaisi: TMenuItem;
    BZoom: TToolbarButton97;
    BDelLigne: TToolbarButton97;
    BAddLigne: TToolbarButton97;

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

    procedure G_DtlRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_DtlCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_DtlCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_DtlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

    procedure POPUPopup(Sender: TObject);
    procedure InfArticleClick(Sender: TObject);
    procedure Allerlapremirelignenonsaisie1Click(Sender: TObject);
    procedure SetSaisiClick(Sender: TObject);

    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure BAddLigneClick(Sender: TObject);
    procedure BDelLigneClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);

  private
    FFirstFind : boolean;
    TOBArtick,
    TOBLots, TOBDeleted,
    TOBGrid, TOBCurrentLigne : TOB;
    FColonneSaisie, FColonneEcart, FColonneSaisi : integer;

    procedure DessineCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);

    procedure EtudieColsListe;
    procedure InitEntete(CodListe, LibListe, LibDepot : String);
    function CopyTOB(T : TOB) : integer;
    procedure ExtractDetailDesLots;
    procedure DisplayGrid;

    procedure GetLigneTOB(ARow : integer);
    function GetCurrentLigneValue(Field : String) : String;

    procedure RefreshPied;
    procedure RefreshLigne(ARow : integer);
    procedure TOBizeLaSaisie(ARow : integer);

    procedure SupprimeLigne(ARow : integer);
    procedure AjouteLot(NumLot : String);

    procedure UpdateDetailDesLots;
    function QuestionSaisieEnCours : Integer;
    function ICanContinue : boolean;

  public
    procedure SetParams(TOBLin, TOBLot, TOBDel : TOB; CodListe, LibListe, LibDepot : String);

  end;


procedure CalculeSaisi(Prefix : String; Ligne : TOB);
procedure CalculeEcart(Prefix : String; Ligne : TOB);
procedure ChargeTob_LibCodeDim(Ligne : TOB; NbDim : Integer);
function CreateTOBLot(TOBMere : TOB; CodeListe, GaArticle, Depot, NumLot : String) : TOB;
//procedure ChangeKeyValue(TOBMere, TOBDel : TOB; var T : TOB; Field : String; Value : Variant);
procedure EntreeDetailInv(TOBLin, TOBLot, TOBDel : TOB; CodListe, LibListe, LibDepot : String);

const Cazakocher = 'þ'; // Symbole de coche en Windingz

implementation
Uses SaisUtil, HEnt1, UIUtil, HDimension,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
     HMsgBox,DimUtil, InvUtil;

{$R *.DFM}


// Place le symbole de coche (qui sera verte) dans le bon champ, si la ligne est saisie
procedure CalculeSaisi(Prefix : String; Ligne : TOB);
var c : Char;
begin
if Ligne = nil then exit;
if Ligne.GetValue(Prefix+'_SAISIINV') <> 'X' then c := ' '
                                         else c := Cazakocher;
Ligne.PutValue('(Saisi)', c);
end;

// Calcule l'écart (stock inv - stock ordi) pour une ligne
procedure CalculeEcart(Prefix : String; Ligne : TOB);
begin
if Ligne = nil then exit;
Ligne.PutValue('(Ecart)', Ligne.GetValue(Prefix+'_INVENTAIRE')-Ligne.GetValue(Prefix+'_QTEPHOTOINV'));
end;

// Chargement en Tob du libellé de chaque dimension d'un article pour une ligne
procedure ChargeTob_LibCodeDim(Ligne : TOB; NbDim : Integer);
var i_ind : integer;
    GrilleDim,CodeDim : String ;
    ChampDim : String ;
begin
if Ligne = nil then exit;
for i_ind := 1 to NbDim do
  begin
  GrilleDim:=VarToStr(Ligne.GetValue('GA_GRILLEDIM'+IntToStr(i_ind))) ;
  CodeDim:=VarToStr(Ligne.GetValue('GA_CODEDIM'+IntToStr(i_ind))) ;
  ChampDim := '(Dim'+IntToStr(i_ind)+')';
  if (GrilleDim='') or (CodeDim='') then Ligne.PutValue(ChampDim, '')
  else Ligne.PutValue(ChampDim, GCGetCodeDim(GrilleDim,CodeDim,i_ind)) ;
  end;
end;

// Crée une TOB lot et retourne cette TOB
function CreateTOBLot(TOBMere : TOB; CodeListe, GaArticle, Depot, NumLot : String) : TOB;
begin
result := TOB.Create('LISTEINVLOT', TOBMere, -1);
with result do
  begin
  AddChampSup('GLI_CODELISTE', false); PutValue('GLI_CODELISTE', CodeListe);
  AddChampSup('GLI_ARTICLE', false); PutValue('GLI_ARTICLE', GaArticle);
  AddChampSup('GLI_NUMEROLOT', false);
  AddChampSup('GLI_SAISIINV', false);
  AddChampSup('GLI_INVENTAIRE', false);
  AddChampSup('GLI_QTEPHOTOINV', false);

  PutValue('GLI_NUMEROLOT', NumLot);
  PutValue('GLI_INVENTAIRE', 0);
  PutValue('GLI_QTEPHOTOINV', GetStockOrdi(GaArticle, Depot, NumLot));
  PutValue('GLI_SAISIINV', '-');
  end;
end;


(*
// Permet de modifier un champ clé1 (effacage de ligne et recréation avec modif)
procedure ChangeKeyValue(TOBMere, TOBDel : TOB; var T : TOB; Field : String; Value : Variant);
var T2 : TOB;
begin
T2 := TOB.Create('LISTINVLIG', TOBMere, T.GetIndex);
T2.Dupliquer(T, false, true);
T2.PutValue(Field, Value);

if T.FieldExists('NewArticleFlag') then T.Free
                                   else begin T.ChangeParent(TOBDel, -1);
                                              T2.AddChampSup('NewArticleFlag', false); end;
T := T2;
end;
*)
// Procédure d'ouverture de la fiche de détail des lots
procedure EntreeDetailInv(TOBLin, TOBLot, TOBDel : TOB; CodListe, LibListe, LibDepot : String);
var FF : TFDetailInv;
    PPANEL : THPanel;
begin
SourisSablier;
FF := TFDetailInv.Create(Application);
FF.SetParams(TOBLin, TOBLot, TOBDel, CodListe, LibListe, LibDepot);
PPANEL := FindInsidePanel;
if PPANEL = nil then
  begin
  try
    FF.ShowModal;
  finally
    FF.Free;
  end;
  SourisNormale;
  end else
  begin
  InitInside(FF, PPANEL);
  FF.Show;
  end;
end;

// Initialisation de l'entête de la fiche
procedure TFDetailInv.SetParams(TOBLin, TOBLot, TOBDel : TOB; CodListe, LibListe, LibDepot : String);
begin
TOBLots := TOBLot;
TOBDeleted := TOBDel;
TOBArtick := TOBLin;

InitEntete(CodListe, LibListe, LibDepot);
end;

// Bidouille pour faire apparaitre les cases à cocher vertes dans la colonne "Ligne saisie"
procedure TFDetailInv.DessineCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);
begin
if (ACol = FColonneSaisi) and (ARow > 0) then
  begin
  Canvas.Font.Name := 'Wingdings';
  Canvas.Font.Size := 10;
  Canvas.Font.Color := $00FF00; // Vert pleine intensité
  Canvas.Font.Style := [fsBold];
  end;
end;

// Repère et mémorise les colonnes importantes sur le grid
procedure TFDetailInv.EtudieColsListe;
var NomCol, LesCols : String;
    col : integer;
begin
LesCols := G_Dtl.Titres[0];
col := 0;
repeat
  NomCol := Uppercase(Trim(ReadTokenSt(LesCols)));
  if NomCol = 'GLI_INVENTAIRE' then FColonneSaisie := col;
  if NomCol = '(ECART)' then FColonneEcart := col;
  if NomCol = '(SAISI)' then FColonneSaisi := col;
  inc(col);
until (LesCols = '') or (NomCol = '');
end;

// Affiche (dans l'entête de la fiche) les intitulés des dimensions utilisées pour l'article
Procedure TFDetailInv.InitEntete(CodListe, LibListe, LibDepot : String);
begin
TCodeListe.Caption := CodListe;
TLibelle.Caption := LibListe;
TLibelleArticle.Caption := TOBArtick.GetValue('GA_LIBELLE');

TDepot.Caption := LibDepot;
TEmplacement.Caption := MyRechDom('GCEMPLACEMENT', TOBArtick.GetValue('GIL_EMPLACEMENT'), '', false);
TUStock.Caption := MyRechDom('GCQUALUNITPIECE', TOBArtick.GetValue('GA_QUALIFUNITESTO'), '', false);

DisplayDimensions(TOBArtick, PDim, 'TGF_GRILLEDIM', 'GF_CODEDIM', 1, true);
end;

// Copie une tob T dans une tob fille de TOBGrid (et calcule le saisi et l'ecart)
function TFDetailInv.CopyTOB(T : TOB) : integer;
var NewTOB : TOB;
begin
result := 0;
NewTOB := TOB.Create('prout', TOBGrid, -1);
NewTOB.Dupliquer(T, false, true);

NewTOB.AddChampSup('(Saisi)', false);
CalculeSaisi('GLI', NewTOB);
NewTOB.AddChampSup('(Ecart)', false);
CalculeEcart('GLI', NewTOB);

NewTOB.Modifie := false;
end;

// Génère la TOBGrid (pour l'affichage) contenant juste les lignes du détail des
//  lots qu'on veut saisir
procedure TFDetailInv.ExtractDetailDesLots;
begin
TOBLots.ParcoursTraitement(['GLI_ARTICLE'], [TOBArtick.GetValue('GIL_ARTICLE')], false, CopyTOB);
end;

// Affiche le détail des lots dans le grid
procedure TFDetailInv.DisplayGrid;
begin
TOBGrid.PutGridDetail(G_Dtl, false, true, G_Dtl.Titres[0], true);
with G_Dtl do if Row >= RowCount then Row := RowCount;
end;

// Récupère la TOB associée à une ligne du grid
procedure TFDetailInv.GetLigneTOB(ARow : integer);
begin
TOBCurrentLigne := TOB(G_Dtl.Objects[0,ARow]); // Le PutGridDetail associe automatiquement
                //à la propriété Objects de la colonne 0, les TOBs correspondant aux lignes
end;

// Récupère une valeur de TOBCurrentLigne et renvoie vide si y'a pas de currentligne
function TFDetailInv.GetCurrentLigneValue(Field : String) : String;
begin
result := GetTOBValue(TOBCurrentLigne, Field);
end;

// Mise a jour des infos par ligne
procedure TFDetailInv.RefreshPied;
begin
GetLigneTOB(G_Dtl.Row);
//BDelLigne.Enabled := (GetCurrentLigneValue('GLI_NUMEROLOT') <> DefaultLot);
end;

// Recalcule l'écart et la case 'Saisi' d'une ligne, et réaffiche cette ligne
procedure TFDetailInv.RefreshLigne(ARow : integer);
begin
GetLigneTOB(ARow);
CalculeSaisi('GLI', TOBCurrentLigne);
CalculeEcart('GLI', TOBCurrentLigne);
G_Dtl.Cells[FColonneSaisie,ARow] := GetCurrentLigneValue('GLI_INVENTAIRE');
G_Dtl.Cells[FColonneSaisi,ARow] := GetCurrentLigneValue('(Saisi)');
G_Dtl.Cells[FColonneEcart,ARow] := GetCurrentLigneValue('(Ecart)');
end;

// Balance l'information saisie dans la TOB ligne correspondante
procedure TFDetailInv.TOBizeLaSaisie(ARow : integer);
var vale : double;
begin
GetLigneTOB(ARow);
vale := Valeur(G_Dtl.Cells[FColonneSaisie,ARow]);
if TOBCurrentLigne <> nil then
  begin
  TOBCurrentLigne.PutValue('GLI_INVENTAIRE',vale);
  if vale <> 0 then TOBCurrentLigne.PutValue('GLI_SAISIINV','X'); // Auto cochage
  end;

RefreshLigne(ARow);
end;

// Supprime une ligne de la liste après confirmation
procedure TFDetailInv.SupprimeLigne(ARow : integer);
begin
GetLigneTOB(ARow);
if (TOBCurrentLigne = nil) or {(GetCurrentLigneValue('GLI_NUMEROLOT') = DefaultLot) or}
   (PGIAsk('Etes-vous sûr de vouloir supprimer cette ligne?','') = mrNo) then exit;

TOBCurrentLigne.AddChampSup('DeletedFlag', false);
with G_Dtl do
  begin
  if RowCount = 2 then RowCount := 3;//AjouteLot(DefaultLot);
  CacheEdit; SynEnabled := false;
  DeleteRow(ARow);
  MontreEdit; SynEnabled := true;
  end;
RefreshPied;
end;

// Ajoute un lot a la liste
procedure TFDetailInv.AjouteLot(NumLot : String);
var NewTOB : TOB;
begin
NewTOB := TOBGrid.FindFirst(['GLI_ARTICLE','GLI_NUMEROLOT'], [TOBArtick.GetValue('GA_ARTICLE'), NumLot], false);
if (NewTOB <> nil) and (NewTOB.FieldExists('DeletedFlag')) then NewTOB.Free;

NewTOB := CreateTOBLot(TOBGrid, TOBArtick.GetValue('GIL_CODELISTE'), TOBArtick.GetValue('GIL_ARTICLE'), TOBArtick.GetValue('GIL_DEPOT'), NumLot);
with NewTOB do
  begin
  AddChampSup('LotAAjouter', false);

  AddChampSup('(Saisi)', false);
  AddChampSup('(Ecart)', false);
  CalculeSaisi('GLI', NewTOB);
  CalculeEcart('GLI', NewTOB);
  end;

with G_Dtl do
  begin
  CacheEdit; SynEnabled := false;
  if TOBGrid.Detail.Count-1 > 0 then InsertRow(RowCount);
  NewTOB.PutLigneGrid(G_Dtl, RowCount-1, false, false, Titres[0]);
  MontreEdit; SynEnabled := true;
  end;
end;

// Met à jour les valeurs saisies, dans la TOBLots (la même que dans SaisieInv)
procedure TFDetailInv.UpdateDetailDesLots;
var SrcTOB, UpdTOB : TOB;
    lig : integer;
    CodArtick, NumLot : String;
begin
for lig := 0 to TOBGrid.Detail.Count-1 do
  begin
  SrcTOB := TOBGrid.Detail[lig];

  CodArtick := SrcTOB.GetValue('GLI_ARTICLE');
  NumLot := SrcTOB.GetValue('GLI_NUMEROLOT');

  if not SrcTOB.FieldExists('LotAAjouter')
   then UpdTOB := TOBLots.FindFirst(['GLI_ARTICLE','GLI_NUMEROLOT'], [CodArtick,NumLot], false)
   else
    begin
    UpdTOB := TOB.Create('LISTEINVLOT', TOBLots, -1);
    UpdTOB.Dupliquer(SrcTOB, false, true);
    UpdTOB.DelChampSup('LotAAjouter', false);
    UpdTOB.AddChampSup('NewLotFlag', false);
    end;

  if SrcTOB.FieldExists('DeletedFlag') then
    begin
    if (not SrcTOB.FieldExists('LotAAjouter')) and (not SrcTOB.FieldExists('NewLotFlag'))
      then UpdTOB.ChangeParent(TOBDeleted, -1)
      else UpdTOB.Free;
    end else
    begin
    UpdTOB.PutValue('GLI_INVENTAIRE', SrcTOB.GetValue('GLI_INVENTAIRE'));
    UpdTOB.PutValue('GLI_SAISIINV', SrcTOB.GetValue('GLI_SAISIINV'));
    end;
  end;

G_Dtl.VidePile(false);
TOBGrid.ClearDetail; // Car le videpile ne libèrerait pas les lignes effacées
end;

// Si la saisie a été modifiée, demande confirmation pour enregistrer...
function TFDetailInv.QuestionSaisieEnCours : Integer;
begin
if (TOBGrid.IsOneModifie) then result := HShowMessage('0;?caption?;La saisie a été modifiée, voulez-vous enregistrer?;Q;YNC;Y;C;','','')
                          else result := mrNo;
end;

// Je peux continuer ? Seulement après avoir demandé confirmation, et être sur
//  que la sauvegarde s'est bien faite
function TFDetailInv.ICanContinue : boolean;
begin
result := false;
 case QuestionSaisieEnCours of
    mrYes : begin UpdateDetailDesLots; result := true; end;
    mrNo : result := true;
 end;
end;


// ------- EVENEMENTS ----------------------------------------------------------

procedure TFDetailInv.FormCreate(Sender: TObject);
begin
// Création de la TOB
TOBGrid := TOB.Create('Copie du detail des lots', nil, -1);
G_Dtl.GetCellCanvas := DessineCell; // Fameuse bidouille pour les coches vertes
end;

procedure TFDetailInv.FormDestroy(Sender: TObject);
begin
// Destruction de la TOB
TOBGrid.Free; TOBGrid := nil;
end;

procedure TFDetailInv.FormShow(Sender: TObject);
begin
// Initialisations diverses de toute la fiche
G_Dtl.ListeParam := 'GCLISTINVDTL';
EtudieColsListe;
HMTrad.ResizeGridColumns(G_Dtl);
AffecteGrid(G_Dtl, taModif);

ExtractDetailDesLots;
DisplayGrid;
RefreshPied;
G_Dtl.SetFocus;
G_Dtl.Col := FColonneSaisie;
end;

procedure TFDetailInv.FormClose(Sender: TObject; var Action: TCloseAction);
begin
G_Dtl.VidePile(true); // Le true libère les TOBs associées au grid
if IsInside(Self) then Action := caFree;
end;

procedure TFDetailInv.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
CanClose := ICanContinue; // Je ferme la fiche seulement si j'ai pas annulé et
                          //  si ca a bien enregistré
end;

procedure TFDetailInv.G_DtlRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if not Chg then RefreshPied;
end;

procedure TFDetailInv.G_DtlCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if G_Dtl.Col <> FColonneSaisie then G_Dtl.Col := FColonneSaisie; // Forcer la colonne active à la colonne de saisie
end;

procedure TFDetailInv.G_DtlCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then exit;
if (G_Dtl.Col = FColonneSaisie) and (ARow <= TOBGrid.Detail.Count)
  then TOBizeLaSaisie(ARow); // Mémorisation de la valeur saisie quand on se barre
end;

procedure TFDetailInv.G_DtlKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var a : integer;
begin
 case Key of
  vk_return : TOBizeLaSaisie(G_Dtl.Row); // Si on appuie sur Entrée ca saisit la valeur
  vk_tab : with G_Dtl do  // Gestion du Tab / Shift-Tab pour descendre ou remonter
      begin
      if ssShift in Shift then a := -1
                          else a := 1;
      if (Row+a) in[1..RowCount-1] then Row := Row + a
                                   else TOBizeLaSaisie(Row);
      Key := 0;
      end;
  vk_delete : if [ssCtrl] = Shift then
      begin
      BDelLigne.Click;
      Key := 0;
      end;
 end;
end;

// Gestion du grisage des items pour le menu info
procedure TFDetailInv.POPUPopup(Sender: TObject);
begin
GetLigneTOB(G_Dtl.Row);
if (GetCurrentLigneValue('GLI_SAISIINV') = 'X') then
  begin
  SetSaisi.Visible := false;
  SetNonSaisi.Enabled := (TOBCurrentLigne <> nil) and (GetCurrentLigneValue('GLI_INVENTAIRE') = '0');
  end else
  begin
  SetSaisi.Visible := true;
  SetSaisi.Enabled := (TOBCurrentLigne <> nil);
  end;

SetNonSaisi.Visible := not SetSaisi.Visible;
end;

// Voir l'article
procedure TFDetailInv.InfArticleClick(Sender: TObject);
var GaArticle : String;
begin
GaArticle := TOBArtick.GetValue('GA_ARTICLE');
{$IFNDEF GPAO}
  if GaArticle <> '' then AglLanceFiche ('GC', 'GCARTICLE', '', GaArticle, 'ACTION=CONSULTATION');
{$ELSE}
  if GaArticle <> '' then V_PGI.DispatchTT(7, taConsult, GaArticle, '', '');
{$ENDIF GPAO}
end;

// Jump to la 1ere ligne non saisie
procedure TFDetailInv.Allerlapremirelignenonsaisie1Click(Sender: TObject);
var lig, ac, ar : integer;
    C : boolean;
begin
for lig := 1 to G_Dtl.RowCount-1 do
  begin
  GetLigneTOB(lig);
  if GetCurrentLigneValue('GLI_SAISIINV') <> 'X' then
    begin
    G_Dtl.SetFocus;
    ac := G_Dtl.Col; ar := G_Dtl.Row;
    G_Dtl.Row := lig;
    C := false;
    G_DtlCellEnter(nil, AC, AR, C);
    break;
    end;
  end;
end;

// Marquer la ligne comme saisie / non saisie
procedure TFDetailInv.SetSaisiClick(Sender: TObject);
var C : Char;
begin
GetLigneTOB(G_Dtl.Row);
if TOBCurrentLigne = nil then exit;
if Sender = SetSaisi then C := 'X'
                     else C := '-';
TOBCurrentLigne.PutValue('GLI_SAISIINV',C);
RefreshLigne(G_Dtl.Row);
end;


// Recherche
procedure TFDetailInv.BChercherClick(Sender: TObject);
begin
if G_Dtl.RowCount < 3 then Exit;
FFirstFind := true;
FindLigne.Execute;
end;

procedure TFDetailInv.FindLigneFind(Sender: TObject);
begin
Rechercher(G_Dtl, FindLigne, FFirstFind);
end;

// Ajour lot
procedure TFDetailInv.BAddLigneClick(Sender: TObject);
var NumLot : String;
    Q : TQuery;
begin
NumLot := '';
if (not InputQuery('Numéro de lot','Entrez le n° de lot à insérer', NumLot)) or (NumLot = '') then exit;

Q := OpenSQL('SELECT GIE_CODELISTE FROM LISTEINVLIG LEFT JOIN LISTEINVENT ON GIL_CODELISTE=GIE_CODELISTE '+
                                                   'LEFT JOIN LISTEINVLOT ON GLI_CODELISTE=GIE_CODELISTE '+
             'WHERE GIE_VALIDATION<>"X" AND GLI_ARTICLE="'+TOBArtick.GetValue('GA_ARTICLE')+'" '+
                    'AND GIE_DEPOT="'+TOBArtick.GetValue('GIL_DEPOT')+'" AND GLI_NUMEROLOT="'+NumLot+'"', true,-1,'',true);
if (not Q.EOF)
   and (TOBDeleted.FindFirst(['GLI_ARTICLE','GLI_NUMEROLOT'], [TOBArtick.GetValue('GA_ARTICLE'), NumLot], false) = nil)
   and (not ((TOBGrid.FindFirst(['GLI_ARTICLE','GLI_NUMEROLOT'], [TOBArtick.GetValue('GA_ARTICLE'), NumLot], false) <> nil)
             and TOBGrid.FindFirst(['GLI_ARTICLE','GLI_NUMEROLOT'], [TOBArtick.GetValue('GA_ARTICLE'), NumLot], false).FieldExists('DeletedFlag'))) then
  begin
  PGIInfo('Le lot est déjà dans la liste '+Q.FindField('GIE_CODELISTE').AsString+' sur le même dépôt','');
  Ferme(Q);
  end else
  begin
  Ferme(Q);
  AjouteLot(NumLot);
  end;

end;

// Suppr. lot
procedure TFDetailInv.BDelLigneClick(Sender: TObject);
begin
SupprimeLigne(G_Dtl.Row);
end;

// Validation
procedure TFDetailInv.BValiderClick(Sender: TObject);
begin
TOBizeLaSaisie(G_Dtl.Row);
UpdateDetailDesLots;
end;

end.
