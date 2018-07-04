unit SaisieInv;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, HTB97, ExtCtrls, HPanel, Menus, StdCtrls, Hctrls, Mask, Grids,
  HEnt1, UIUtil, SaisUtil, UTOB, HMsgBox, EntGC,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
  UtilPGI,M3FP,HDimension,DetailInv,HStatus,UTOF,AGLInit,ComCtrls,UtilArticle,
  ParamSoc;


procedure EntreeSaisieInv(CodeListe : String);

type
  TFSaisieInv = class(TForm)
    PENTETE: THPanel;
    ToolWindow971: TToolWindow97;
    Dock971: TDock97;
    BChercher: TToolbarButton97;
    BInfos: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    PPIED: THPanel;
    G_Inv: THGrid;
    TGIL_EMPLACEMENT: THLabel;
    POPZ: TPopupMenu;
    FindLigne: TFindDialog;
    InfArticle: TMenuItem;
    TCodeListe: THLabel;
    TLibelle: THLabel;
    TDepot: THLabel;
    T_Depot: THLabel;
    HLabel2: THLabel;
    TUStock: THLabel;
    T_Emplacement: THLabel;
    TEmplacement: THLabel;
    TDetailPrix: TToolWindow97;
    BDetailPrix: TToolbarButton97;
    grp_depot: TGroupBox;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    grp_article: TGroupBox;
    HLabel10: THLabel;
    HLabel11: THLabel;
    HLabel12: THLabel;
    HLabel13: THLabel;
    grp_saisi: TGroupBox;
    HLabel14: THLabel;
    HLabel15: THLabel;
    HLabel16: THLabel;
    HLabel17: THLabel;
    GIL_EMPLACEMENT: THValComboBox;
    BZoom: TToolbarButton97;
    POPU: TPopupMenu;
    ModifPrix: TMenuItem;
    JumpNonSaisi: TMenuItem;
    SetSaisi: TMenuItem;
    SetNonSaisi: TMenuItem;
    BAddLigne: TToolbarButton97;
    BDelLigne: TToolbarButton97;
    GIL_DPA: THNumEdit;
    GIL_PMAP: THNumEdit;
    GIL_DPR: THNumEdit;
    GIL_PMRP: THNumEdit;
    GIL_PMAPART: THNumEdit;
    GIL_PMRPART: THNumEdit;
    GIL_DPRART: THNumEdit;
    GIL_DPAART: THNumEdit;
    GIL_PMAPSAIS: THNumEdit;
    GIL_PMRPSAIS: THNumEdit;
    GIL_DPRSAIS: THNumEdit;
    GIL_DPASAIS: THNumEdit;
    TDate: THLabel;
    PageScroller1: TPageScroller;
    TOTQTESTO: THNumEdit;
    TOTQTEINV: THNumEdit;
    TOTECART: THNumEdit;
    LTOTQTE: TLabel;
    bSelectAll: TToolbarButton97;
    LConsult: TLabel;
    SetAllSaisie: TMenuItem;
    BActualise: TToolbarButton97;

    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

    procedure G_InvRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_InvCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_InvCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_InvElipsisClick(Sender: TObject);
    procedure G_InvKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure G_InvSorted(Sender: TObject);

    procedure GIL_EMPLACEMENTChange(Sender: TObject);

    procedure InfArticleClick(Sender: TObject);
    procedure POPUPopup(Sender: TObject);
    procedure Voirlesprix1Click(Sender: TObject);
    procedure Allerlapremirelignenonsaisie1Click(Sender: TObject);
    procedure SetSaisiClick(Sender: TObject);

    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure BDetailPrixClick(Sender: TObject);
    procedure TDetailPrixClose(Sender: TObject);
    procedure BAddLigneClick(Sender: TObject);
    procedure BDelLigneClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure SetAllSaisieClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BActualiseClick(Sender: TObject);
    procedure G_InvDblClick(Sender: TObject);

  private
    FFirstFind : boolean;
    TOBLignes, TOBDeletedLines, TOBCurrentLigne,
    TOBLots, TOBDeletedLots,TOBListe : TOB;
    FColonneSaisie, FColonneEcart, FColonneSaisi, FColonneOrdi,FColonnePMAP,FColonneDate : integer;

    procedure DessineCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);

    procedure EtudieColsListe;
    procedure InitEntete;
    procedure LoadLesTOB;
    function CalcCumul(T : TOB) : integer;
    procedure ReCalculeCumuls;
    procedure ProcessTOBLignes;
    procedure ChargeListe;
    procedure VideLesTOB;

    procedure GetLigneTOB(ARow : integer);
    function GetCurrentLigneValue(Field : String) : String;
    function IsCurrentLigneCumul : Boolean;

    procedure RefreshPied;
    procedure RefreshLigne(ARow : integer);
    procedure UpdateElipsis;
    function TOBizeLaSaisie(ARow : integer) : boolean;

    procedure InsereLigne(GaArtick : String);
    function DelLot(T : TOB) : integer;
    procedure SupprimeLigne(ARow : integer);

    procedure ValiderSaisie;
    function TransactionValide : boolean;
    function QuestionSaisieEnCours : Integer;
    function ICanContinue : boolean;

//    procedure UnDelTOB(T : TOB);
//    procedure UnDeleteLigne(CodArtick : String);

  public
    FCodeEmpl : String;
    FCodeListe : String;
    FDepot : String;
    FDateInv : TDateTime;
    FStockClos, BSaisieAveugle : Boolean;
    Action : TActionFiche ;
  end;


  TOF_PrixInv = class(TOF)
  public
    procedure OnLoad; override;
    procedure OnUpdate; override;

  end;


implementation
uses UTOFListeInv, InvUtil, DimUtil, AGLInitGC, UtilDispGC,BTPlanning;

const
	// libellés des messages
	TexteMessage: array[1..4] of string 	= (
          {1}         'Confirmation'
          {2}        ,'Le stock inventorié va être forcé à la valeur du stock ordinateur, sur toutes les lignes non saisies !'
          {3}        ,'Le stock inventorié va être remis à zéro, sur toutes les lignes saisies !'
          {4}        ,'Si le stock ordinateur est négatif, le stock inventorié sera forcé à zéro.'
                     );
{$R *.DFM}

// Procédure d'ouverture de la fiche de saisie d'inventaire
procedure EntreeSaisieInv(CodeListe : String);
var FF : TFSaisieInv;
    PPANEL : THPanel;
begin
if CodeListe = '' then exit;
SourisSablier;
FF := TFSaisieInv.Create(Application);
FF.FCodeListe := CodeListe;
FF.Action:=taModif ;
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

// Procédure pour l'appel de la saisie d'inventaire depuis le script
procedure AGLEntreeSaisieInv(Parms : Array of Variant; Nb : Integer);
begin
EntreeSaisieInv(Parms[0]);
end;


// ------- METHODES ------------------------------------------------------------

// Bidouille pour faire apparaitre les cases à cocher vertes dans la colonne "Ligne saisie"
procedure TFSaisieInv.DessineCell(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState);
begin
if (ACol = FColonneSaisi) and (ARow > 0) then
  begin
  Canvas.Font.Name := 'Wingdings';
  Canvas.Font.Size := 10;
  Canvas.Font.Color := $00FF00; // Vert pleine intensité
  Canvas.Font.Style := [fsBold];
  end;
end;

// Repère et mémorise les colonnes importantes sur le grid ; interdit également
// certaines colonnes au tri par clic sur l'entête
procedure TFSaisieInv.EtudieColsListe;
var NomCol, LesCols : String;
    LibDim, FF,FFP : String;
    col, i : integer;

begin
LesCols := G_Inv.Titres[0]; // Contient les noms des champs séparés par des ;
col := 0;

if V_PGI.OkDecQ > 0 then
begin
  FF := '0.';
  for i := 1 to V_PGI.OkDecQ - 1 do
  begin
    FF := FF + '#';
  end;
  FF := FF + '0';
end;
if V_PGI.OkDecP > 0 then
begin
  FFP := '0.';
  for i := 1 to V_PGI.OkDecP - 1 do
  begin
    FFP := FFP + '#';
  end;
  FFP := FFP + '0';
end;

TOTQTESTO.Decimals := V_PGI.OkDecQ;
TOTQTEINV.Decimals := V_PGI.OkDecQ;
TOTECART.Decimals := V_PGI.OkDecQ;
TOTQTESTO.Masks.PositiveMask := FF;
TOTQTEINV.Masks.PositiveMask := FF;
TOTECART.Masks.PositiveMask := FF;

repeat
  NomCol := Uppercase(Trim(ReadTokenSt(LesCols)));
  if NomCol = '(SAISI)' then begin FColonneSaisi := col; G_Inv.ColTypes[col] := 'I'; end // I = pas pouvoir trier
  else if NomCol = '(DIM1)' then LibDim:=RechDom('GCCATEGORIEDIM','DI1',False)
  else if NomCol = '(DIM2)' then LibDim:=RechDom('GCCATEGORIEDIM','DI2',False)
  else if NomCol = '(DIM3)' then LibDim:=RechDom('GCCATEGORIEDIM','DI3',False)
  else if NomCol = '(DIM4)' then LibDim:=RechDom('GCCATEGORIEDIM','DI4',False)
  else if NomCol = '(DIM5)' then LibDim:=RechDom('GCCATEGORIEDIM','DI5',False)
  else if NomCol = 'GIL_QTEPHOTOINV' then
  begin
    FColonneOrdi := col; G_Inv.ColTypes[col] := 'I';
    if BSaisieAveugle then G_Inv.ColWidths[col]:=-1;
    G_Inv.ColFormats[col] := FF;
  end
  else if NomCol = 'GIL_INVENTAIRE' then
  begin
    FColonneSaisie := col;
    G_Inv.ColFormats[col] := FF + ';' + FF;
  end
  else if NomCol = '(ECART)' then
  begin
    FColonneEcart := col; G_Inv.ColTypes[col] := 'I';
    if BSaisieAveugle then G_Inv.ColWidths[col]:=-1;
    G_Inv.ColFormats[col] := FF;
  end else if NomCol = 'GIL_PMAP' then
  begin
    FColonnePMAP := col; G_Inv.ColTypes[col] := 'I';
    G_Inv.ColFormats[col] := FFP;
  end else if NomCol = 'GIL_DATESAISIE' then
  begin
    FColonneDate := col; 
  end;

  if (NomCol='(DIM1)') or (NomCol='(DIM2)') or (NomCol='(DIM3)') or (NomCol='(DIM4)') or (NomCol='(DIM5)') then
  begin
    if (LibDim='') or (LibDim='Error') then G_Inv.ColLengths[col] := 0
    else G_Inv.cells[col,0]:=LibDim ;
  end;

{$IFDEF BTP}
  if (Pos(NomCol,'GIL_STATUTDISPO;GIL_STATUTFLUX;GCQ_LIBELLE;GIL_LOTINTERNE;GIL_LOTEXTERNE;GIL_SERIEINTERNE;GIL_SERIEEXTERNE;GIL_REAFFECTATION')>0) then
  begin
  	G_Inv.ColWidths [col] := 0;
  	G_Inv.ColLengths [col] := -1;
  end;
{$ENDIF}
  inc(col);
until (LesCols = '') or (NomCol = '');
end;

// Initialise l'entête de la fiche (affichage des libellés du dépot et de la liste)
procedure TFSaisieInv.InitEntete;
var Q : TQuery;
begin
if not VH_GC.GCMultiDepots then T_DEPOT.Caption:='Etablissement :' ;

if (ctxMode in V_PGI.PGIContexte) then
  begin
  TGIL_EMPLACEMENT.Visible:=False;
  GIL_EMPLACEMENT.Visible:=False;
  T_Emplacement.Visible:=False;
  TEmplacement.Visible:=False;
  end;

Q := OpenSQL('SELECT GIE_LIBELLE, GIE_DEPOT, GIE_DATEINVENTAIRE, GIE_STOCKCLOS, GIE_VALIDATION '+
             'FROM LISTEINVENT WHERE GIE_CODELISTE="'+FCodeListe+'"', true,-1,'',true);
TCodeListe.Caption := FCodeListe;
if Q.EOF then
  begin
  TLibelle.Caption := '';
  TDate.Caption := '';
  FDepot := '';
  TDepot.Caption := '';
  FStockClos := False;
  end else
  begin
  TLibelle.Caption := Q.FindField('GIE_LIBELLE').AsString;
  TDate.Caption := 'Inventaire du '+Q.FindField('GIE_DATEINVENTAIRE').AsString;
  FDepot := Q.FindField('GIE_DEPOT').AsString;
  TDepot.Caption := RechDom('GCDEPOT', FDepot, false);
  FDateInv := Q.FindField('GIE_DATEINVENTAIRE').AsDateTime;
  FStockClos:=((Q.FindField('GIE_STOCKCLOS').AsString)='X');
  if (Q.FindField('GIE_VALIDATION').AsString)='X' then Action:=taConsult;
  TOBListe.SelectDB ('',Q);
  end;
Ferme(Q);
end;

// Chargement de la liste des lignes d'inventaire
procedure TFSaisieInv.LoadLesTOB;
var Q : TQuery;
    i_ind : integer ;
    iLot : integer;
    Emplacement : string;
begin
VideLesTOB;
if FCodeEmpl<>'' then Emplacement := 'AND GIL_EMPLACEMENT = "'+FCodeEmpl+'" '
else Emplacement := '';
// Lignes
Q := OpenSQL('SELECT LISTEINVLIG.*,GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_GRILLEDIM1,'+
        'GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,GA_CODEDIM2,'+
        'GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_LOT,GA_QUALIFUNITESTO,GA_QUALIFUNITEVTE,'+
        'GA_PRIXPOURQTE,GA_DPA,GA_PMAP,GA_DPR,GA_PMRP,GQ_PHYSIQUE '+
        'FROM LISTEINVLIG LEFT JOIN ARTICLE ON GIL_ARTICLE=GA_ARTICLE '+
        'LEFT JOIN DISPO ON GQ_ARTICLE=GIL_ARTICLE AND GQ_DEPOT=GIL_DEPOT AND GQ_CLOTURE="-" '+
        'WHERE GIL_CODELISTE="'+FCodeListe+'" '+ Emplacement +
        'ORDER BY GIL_EMPLACEMENT, GIL_ARTICLE', true,-1,'',true);
TOBLignes.LoadDetailDB('LISTEINVLIG','','',Q,false,true);
Ferme(Q);
// Lots
Q := OpenSQL('SELECT * FROM LISTEINVLOT WHERE GLI_CODELISTE="'+FCodeListe+'" ', true,-1,'',true);
TOBLots.LoadDetailDB('LISTEINVLOT','','',Q,false,true);
Ferme(Q);
for iLot := 0 to TOBLots.Detail.Count -1 do
  begin
  if TOBLots.Detail[iLot].GetValue('GLI_SAISIINV') = '-' then
     TOBLots.Detail[iLot].PutValue('GLI_INVENTAIRE',TOBLots.Detail[iLot].GetValue('GLI_QTEPHOTOINV'));
  TOBLots.Detail[iLot].PutValue('GLI_SAISIINV','X');
  end;
if VH_GC.GCIfDefCEGID then
  for i_ind:=0 to TOBLignes.Detail.Count-1 do
      begin
      if TOBLignes.Detail[i_ind].GetValue('GIL_QTEPHOTOINV')=0 then TOBLignes.Detail[i_ind].PutValue('GIL_SAISIINV','X');
      end;
end;

function TFSaisieInv.CalcCumul(T : TOB) : integer;
begin
result := 0;
T.PutValue('GIL_QTEPHOTOINV', TOBLots.Somme('GLI_QTEPHOTOINV', ['GLI_ARTICLE'], [T.GetValue('GIL_ARTICLE')], false));
T.PutValue('GIL_INVENTAIRE', TOBLots.Somme('GLI_INVENTAIRE', ['GLI_ARTICLE'], [T.GetValue('GIL_ARTICLE')], false));
if (TOBLots.FindFirst(['GLI_ARTICLE'], [T.GetValue('GIL_ARTICLE')], false) = nil) or
   (TOBLots.FindFirst(['GLI_ARTICLE', 'GLI_SAISIINV'], [T.GetValue('GIL_ARTICLE'),'-'], false) <> nil)
  then T.PutValue('GIL_SAISIINV','-')
  else T.PutValue('GIL_SAISIINV','X');
end;

// Recalcule les quantités cumulées pour les articles gérés par lots
procedure TFSaisieInv.ReCalculeCumuls;
begin
TOBLignes.ParcoursTraitement(['GIL_LOT'], ['X'], false, CalcCumul);
end;

// Parcourt la liste de cumuls (préalablement à son PutGridage) pour ajouter à
//  chaque ligne les champs Ecart et Saisi, et transformer les codes dimentions
//  en libellés.
procedure TFSaisieInv.ProcessTOBLignes;
var lig : integer;
    T : TOB;
begin
InitMove(TOBLignes.Detail.Count,'');
for lig := 0 to TOBLignes.Detail.Count-1 do
  begin
  T := TOBLignes.Detail[lig];
  T.AddChampSup('(Saisi)',false);
  CalculeSaisi('GIL', T);
  T.AddChampSup('(Ecart)',false);
  CalculeEcart('GIL', T);
  T.AddChampSup('(Dim1)',false);
  T.AddChampSup('(Dim2)',false);
  T.AddChampSup('(Dim3)',false);
  T.AddChampSup('(Dim4)',false);
  T.AddChampSup('(Dim5)',false);
  ChargeTob_LibCodeDim(T, MaxDimension);
  MoveCur(false);
  end;
FiniMove;
end;

// Procédure qui charge la liste d'inventaire, fait toutes les opérations nécessaires
//  et affiche le tout dans le grid.
procedure TFSaisieInv.ChargeListe;
begin
Transactions(LoadLesTOB, 1);
ReCalculeCumuls;
ProcessTOBLignes;
TOTQTEINV.Value := TOBLignes.Somme('GIL_INVENTAIRE',[''],[''],False);
TOTQTESTO.Value := TOBLignes.Somme('GIL_QTEPHOTOINV',[''],[''],False);
TOTECART.Value := TOTQTEINV.Value - TOTQTESTO.Value;
TOBLignes.PutGridDetail(G_Inv, false, true, G_Inv.Titres[0], true);
with G_Inv do if Row >= RowCount then Row := RowCount-1;
RefreshPied;
UpdateElipsis;
end;

procedure TFSaisieInv.VideLesTOB;
begin
TOBLignes.ClearDetail;
TOBDeletedLines.ClearDetail;
TOBLots.ClearDetail;
TOBDeletedLots.ClearDetail;
end;

// Récupère la TOB (cumul) associée à une ligne du grid (ainsi que la première
//  TOB correspondante dans la liste de détail)
procedure TFSaisieInv.GetLigneTOB(ARow : integer);
begin
TOBCurrentLigne := TOB(G_Inv.Objects[0,ARow]); // Le PutGridDetail associe automatiquement
                //à la propriété Objects de la colonne 0, les TOBs correspondant aux lignes
end;

// Récupère une valeur de TOBCurrentLigne et renvoie vide si y'a pas de currentligne
function TFSaisieInv.GetCurrentLigneValue(Field : String) : String;
begin
if TOBCurrentLigne = nil then result := ''
                         else result := TOBCurrentLigne.GetValue(Field);
end;

// Retourne true si la ligne courante est une ligne cumulée
function TFSaisieInv.IsCurrentLigneCumul : Boolean;
begin
result := (GetCurrentLigneValue('GIL_LOT') = 'X');
end;

// Rafraichit les informations du pied de la fiche en fonction de la ligne en cours
procedure TFSaisieInv.RefreshPied;
begin
GetLigneTOB(G_Inv.Row);
if TOBCurrentLigne = nil then exit;
TOBCurrentLigne.PutEcran(Self, TDetailPrix);
if VarIsNull(TOBCurrentLigne.getvalue('GA_QUALIFUNITESTO')) then TUStock.Caption := '1'
else TUStock.Caption := MyRechDom('GCQUALUNITPIECE', TOBCurrentLigne.getvalue('GA_QUALIFUNITESTO'), '', false);
if Not(ctxMode in V_PGI.PGIContexte) then
   TEmplacement.Caption := MyRechDom('GCEMPLACEMENT', GetCurrentLigneValue('GIL_EMPLACEMENT'), '', false);
end;

// Recalcule l'écart et la case 'Saisi' d'une ligne, et réaffiche cette ligne
procedure TFSaisieInv.RefreshLigne(ARow : integer);
begin
GetLigneTOB(ARow);
if TOBCurrentLigne <> nil then
   begin
   CalculeSaisi('GIL', TOBCurrentLigne);
   CalculeEcart('GIL', TOBCurrentLigne);
   {G_Inv.Cells[FColonneOrdi,ARow] := GetCurrentLigneValue('GIL_QTEPHOTOINV');
   G_Inv.Cells[FColonneSaisie,ARow] := GetCurrentLigneValue('GIL_INVENTAIRE');
   G_Inv.Cells[FColonneSaisi,ARow] := GetCurrentLigneValue('(Saisi)');
   G_Inv.Cells[FColonneEcart,ARow] := GetCurrentLigneValue('(Ecart)');}
   TOBCurrentLigne.PutLigneGrid(G_Inv, ARow, false, true, G_Inv.Titres[0]);
   end;
end;

// Affiche ou cache le bouton elipsis selon que la ligne est cumulée ou non
procedure TFSaisieInv.UpdateElipsis;
begin
GetLigneTOB(G_Inv.Row);
G_Inv.ElipsisButton := ((TOBCurrentLigne <> nil) and IsCurrentLigneCumul) or  ( not ((TOBCurrentLigne <> nil) and IsCurrentLigneCumul) and (G_inv.col = FColonneDate) ) ;
G_Inv.CacheEdit; G_Inv.MontreEdit; // Pour contourner le bug de l'elipsis qui n'apparait pas
                                   //  lorsqu'on change la Row courante dans le code
end;

// Balance l'information saisie dans les TOB lignes correspondantes (cumul ET détail)
function TFSaisieInv.TOBizeLaSaisie(ARow : integer) : boolean;
var vale  : double;
		DateS : TdateTime;
    ind   : Integer;
    TOBL  : TOB;
begin
  result := true;

  GetLigneTOB(ARow);

  //FV1 - 04/05/2017 -  FS#2514 - SERVAIS : En saisie d'inventaire, la modification d'une date sur une ligne ne fonctionne pas.
  if TOBCurrentLigne = nil then Exit;

  //contrôle de la quantité saisie
  vale := Abs(Valeur(G_Inv.Cells[FColonneSaisie,ARow]));
  if IsCurrentLigneCumul and (TOBLots.Somme('GLI_INVENTAIRE', ['GLI_ARTICLE'], [TOBCurrentLigne.GetValue('GIL_ARTICLE')], false) <> vale) then
  begin
    PGIInfo('Pour modifier le stock de cette ligne, cliquez sur le bouton adjacent','');
    // Warning et on renvoi faux pour dire que c'est pas bon
    result := false;
    exit;
  end
  else
  begin
    TOTQTEINV.Value := TOTQTEINV.Value + vale - TOBCurrentLigne.GetValue('GIL_INVENTAIRE');
    TOTECART.Value  := TOTQTEINV.Value - TOTQTESTO.Value;
    TOBCurrentLigne.PutValue('GIL_INVENTAIRE',vale); // hop
    // Mettre automatiquement la coche 'Saisi'
    if (vale <> 0) or (TOBCurrentLigne.GetValue('GIL_QTEPHOTOINV') = 0) then TOBCurrentLigne.PutValue('GIL_SAISIINV','X');
  end;

  //FV1 - 04/05/2017 -  FS#2514 - SERVAIS : En saisie d'inventaire, la modification d'une date sur une ligne ne fonctionne pas.
  Try
    Try
      DateS := StrToDate(G_Inv.Cells[FColonneDate,ARow]);
    Except
      DateS := StrToDate(DateToStr(now));
    end;
  Finally
    if StrToDate(TOBCurrentLigne.GetValue('GIL_DATESAISIE')) <> DateS then
    begin
      if PGIAsk('Désirez-vous appliquer cette modification à l''ensemble des lignes de cette liste d''inventaire ?', 'Liste Inventaire') = mrYes then
      begin
        For ind := 0 to TOBLignes.Detail.count - 1 do
        begin
          TOBL := TOBLignes.Detail[Ind];
          TOBL.SetDateTime('GIL_DATESAISIE',DateS); // hop
        end;
        TOBLignes.PutGridDetail(G_Inv, false, true, G_Inv.Titres[0], true);
      end;
    end;
    TOBCurrentLigne.SetDateTime('GIL_DATESAISIE',DateS); // hop
  end;

  RefreshLigne(ARow);

end;

// Insere un article dans la liste
procedure TFSaisieInv.InsereLigne(GaArtick : String);
var QA, QD,QDL : TQuery;
    FUS,FUV,FPPQ : Double;
    GereParLot : Boolean;
    NewTOB : TOB;
    SQLArt : String;
    CodArtick, Emplacement : String;
    i_fields : Integer;
    StocPhy, DPAD, PMAPD, DPRD, PMRPD, DPAA, PMAPA, DPRA, PMRPA : Double;
    QPR : TQtePrixRec;
begin
SQLArt:='SELECT GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,'+
        'GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,'+
        'GA_CODEDIM5,GA_LOT,GA_QUALIFUNITESTO,GA_QUALIFUNITEVTE,GA_PRIXPOURQTE,'+
        'GA_DPA,GA_PMAP,GA_DPR,GA_PMRP FROM ARTICLE WHERE GA_ARTICLE="'+GaArtick+'"';
QA := OpenSQL(SQLArt, true,-1,'',true);
QD := OpenSQL('SELECT * FROM DISPO WHERE GQ_ARTICLE="'+GaArtick+'" AND GQ_DEPOT="'+FDepot+'" AND '+
              'GQ_CLOTURE="-"',true,-1,'',true);
{MRNONGESTIONLOT}
QDL := OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="'+GaArtick+'" AND GQL_DEPOT="'+FDepot+'"',true,-1,'',true);
GereParLot := (QA.FindField('GA_LOT').AsString = 'X');

CodArtick := QA.FindField('GA_CODEARTICLE').AsString;
Emplacement := QD.FindField('GQ_EMPLACEMENT').AsString;

If FStockClos=True then
   begin
   QPR := GetQtePrixDateListe(GaArtick, FDepot, FDateInv);
   if not QPR.SomethingReturned then StocPhy := QD.FindField('GQ_PHYSIQUE').AsFloat
                                else StocPhy := QPR.Qte;
   end else
   begin
   if QD.EOF then StocPhy := 0
             else StocPhy := QD.FindField('GQ_PHYSIQUE').AsFloat;
   end;


FUS := RatioMesure('PIE', QA.FindField('GA_QUALIFUNITESTO').AsString);
FUV := RatioMesure('PIE', QA.FindField('GA_QUALIFUNITEVTE').AsString);
FPPQ := QA.FindField('GA_PRIXPOURQTE').AsFloat; if FPPQ = 0 then FPPQ := 1.0;

DPAA := Ratioize(QA.FindField('GA_DPA').AsFloat, FUS, FUV, FPPQ);
PMAPA := Ratioize(QA.FindField('GA_PMAP').AsFloat, FUS, FUV, FPPQ);
DPRA := Ratioize(QA.FindField('GA_DPR').AsFloat, FUS, FUV, FPPQ);
PMRPA := Ratioize(QA.FindField('GA_PMRP').AsFloat, FUS, FUV, FPPQ);
if QD.EOF then
  begin
  DPAD := DPAA;
  PMAPD := PMAPA;
  DPRD := DPRA;
  PMRPD := PMRPA;
  end else
  begin
  DPAD := Ratioize(QD.FindField('GQ_DPA').AsFloat, FUS, FUV, FPPQ);
  PMAPD := Ratioize(QD.FindField('GQ_PMAP').AsFloat, FUS, FUV, FPPQ);
  DPRD := Ratioize(QD.FindField('GQ_DPR').AsFloat, FUS, FUV, FPPQ);
  PMRPD := Ratioize(QD.FindField('GQ_PMRP').AsFloat, FUS, FUV, FPPQ);
  end;

// Ajout de TOB
NewTOB := TOB.Create('LISTEINVLIG', TOBLignes, -1);
with NewTOB do
  begin
  //
  AddChampSupValeur ('GQ_DISPO',StocPhy);
  //
  PutValue('GIL_CODELISTE', FCodeListe);
  PutValue('GIL_ARTICLE', GaArtick);
  PutValue('GIL_CODEARTICLE', CodArtick);
  if GereParLot then PutValue('GIL_LOT', 'X') else PutValue('GIL_LOT', '-');
  PutValue('GIL_DEPOT', FDepot);
  PutValue('GIL_EMPLACEMENT', Emplacement);
  PutValue('GIL_SAISIINV', '-');
  PutValue('GIL_INVENTAIRE', 0);
  PutValue('GIL_QTEPHOTOINV', StocPhy);
  PutValue('GIL_DPA', DPAD);
  PutValue('GIL_PMAP', PMAPD);
  PutValue('GIL_DPR', DPRD);
  PutValue('GIL_PMRP', PMRPD);
  PutValue('GIL_DPAART', DPAA);
  PutValue('GIL_PMAPART', PMAPA);
  PutValue('GIL_DPRART', DPRA);
  PutValue('GIL_PMRPART', PMRPA);
  PutValue('GIL_DPASAIS', 0.0);
  PutValue('GIL_PMAPSAIS', 0.0);
  PutValue('GIL_DPRSAIS', 0.0);
  PutValue('GIL_PMRPSAIS', 0.0);

  AddChampSup('(Saisi)', false);
  AddChampSup('(Ecart)', false);
  AddChampSup('(Dim1)', false);
  AddChampSup('(Dim2)', false);
  AddChampSup('(Dim3)', false);
  AddChampSup('(Dim4)', false);
  AddChampSup('(Dim5)', false);
  AddChampSup('NewArticleFlag', false);

  TOTQTESTO.Value := TOTQTESTO.Value + StocPhy;
  TOTECART.Value := TOTQTEINV.Value - TOTQTESTO.Value;

  // Recopie (en vrac) des champs utiles après dans le fonctionnement de la saisie
 with QA do
//   for i_fields := 0 to Fields.Count-1 do  SD pour e-AGL le 03/04/01
   for i_fields := 0 to FieldCount -1 do
     begin
     AddChampSup(Fields[i_fields].FieldName, false);
     PutValue(Fields[i_fields].FieldName, Fields[i_fields].AsVariant);
     end;
  end;
ChargeTob_LibCodeDim(NewTOB, MaxDimension);

{MRNONGESTIONLOT}
if GereParLot then
 while not QDL.EOF do
   begin
   CreateTOBLot(TOBLots, FCodeListe, GaArtick, FDepot,
                QDL.FindField('GQL_NUMEROLOT').AsString).AddChampSup('NewLotFlag', false);
   QDL.Next;
   end;
Ferme(QDL);

Ferme(QA);
Ferme(QD);

//RePutDetail;
with G_Inv do
  begin
  CacheEdit; SynEnabled := false;
  if TOBLignes.Detail.Count-1 > 0 then InsertRow(RowCount);
  NewTOB.PutLigneGrid(G_Inv, RowCount-1, false, false, Titres[0]);
  MontreEdit; SynEnabled := true;
  end;
RefreshPied;
UpdateElipsis;

end;

function TFSaisieInv.DelLot(T : TOB) : integer;
begin
result := 0;
if T.FieldExists('NewLotFlag') then T.Free
                               else T.ChangeParent(TOBDeletedLots, -1);
end;

// Supprime une ligne de la liste après confirmation
procedure TFSaisieInv.SupprimeLigne(ARow : integer);
begin
GetLigneTOB(ARow);
if (TOBCurrentLigne = nil)
   or (PGIAsk('Etes-vous sûr de vouloir supprimer cette ligne?','') = mrNo) then exit;

if TOBCurrentLigne.GetValue('GIL_LOT') = 'X' then
 TOBLots.ParcoursTraitement(['GLI_CODELISTE', 'GLI_ARTICLE'],
                            [FCodeListe, TOBCurrentLigne.GetValue('GIL_ARTICLE')],
                            false,
                            DelLot);

TOTQTEINV.Value := TOTQTEINV.Value - TOBCurrentLigne.GetValue('GIL_INVENTAIRE');
TOTQTESTO.Value := TOTQTESTO.Value - TOBCurrentLigne.GetValue('GIL_QTEPHOTOINV');
TOTECART.Value := TOTQTEINV.Value - TOTQTESTO.Value;
if TOBCurrentLigne.FieldExists('NewArticleFlag') then TOBCurrentLigne.Free
                                                 else TOBCurrentLigne.ChangeParent(TOBDeletedLines, -1);
with G_Inv do
  begin
  CacheEdit; SynEnabled := false;
  if RowCount = 2 then RowCount := 3;
  DeleteRow(ARow);
  MontreEdit; SynEnabled := true;
  end;
RefreshPied;
UpdateElipsis;
end;

// Balance les TOBs modifiées dans la table
procedure TFSaisieInv.ValiderSaisie;
begin
if Action<>taConsult then
   begin
   TOBizeLaSaisie(G_Inv.Row);
   TOBDeletedLines.DeleteDB(true);
   TOBLignes.InsertOrUpdateDB(true);
   TOBListe.UpdateDB;
   TOBDeletedLots.DeleteDB(true);
   TOBLots.InsertOrUpdateDB(true);
   // Mise à jour LastUser et DateModif
   ExecuteSQL('UPDATE LISTEINVENT SET GIE_DATEMODIF="'+USDateTime(NowH)+'", GIE_UTILISATEUR="'+V_PGI.USER+'" WHERE GIE_CODELISTE="'+FCodeListe+'"');
   end;
VideLesTOB;  // Pour pas qu'il dise que c'est modifié après avoir sauvé
end;

// Transactionne la methode précédente, avec messages d'erreur et tout
function TFSaisieInv.TransactionValide : boolean;
var IOErr : TIOErr;
begin
IOErr := Transactions(ValiderSaisie, 2);
result := false;
 case IOErr of
  oeOK : result := true;
  oeUnknown : PGIBox('Saisie non enregistrée !','');
  oeSaisie : PGIBox('Saisie non enregistrée (en cours de traitement par un autre utilisateur) !','');
 end;
end;

// Si la saisie a été modifiée, demande confirmation pour enregistrer...
function TFSaisieInv.QuestionSaisieEnCours : Integer;
begin
if (TOBLignes.IsOneModifie) or (TOBLots.IsOneModifie) or
   (TOBDeletedLines.Detail.Count > 0) or (TOBDeletedLots.Detail.Count > 0)
  then result := HShowMessage('0;?caption?;La saisie a été modifiée, voulez-vous enregistrer?;Q;YNC;Y;C;','','')
  else result := mrNo;
end;

// Je peux continuer ? Seulement après avoir demandé confirmation, et être sur
//  que la sauvegarde s'est bien faite
function TFSaisieInv.ICanContinue : boolean;
begin
result := false;
 case QuestionSaisieEnCours of
    mrYes : if TransactionValide then result := true;
    mrNo : result := true;
 end;
end;


(*
procedure TFSaisieInv.UnDelTOB(T : TOB);
var idx : integer;
begin
// Calcul de la position d'insertion
idx := 0;
while (idx < TOBDetail.Detail.Count)
 and (TOBDetail.Detail[idx].GetValue('GIL_EMPLACEMENT') <= T.GetValue('GIL_EMPLACEMENT'))
 and (TOBDetail.Detail[idx].GetValue('GIL_ARTICLE') <= T.GetValue('GIL_ARTICLE'))
  do Inc(idx);
T.ChangeParent(TOBDetail, idx); // hop
end;

// Récupère une ligne effacée
procedure TFSaisieInv.UnDeleteLigne(CodArtick : String);
begin
ParcourTOB(TOBDeletedLines, ['GIL_ARTICLE'], [CodArtick], false, UnDelTOB);

RePutDetail;
end;
*)

// ------- EVENEMENTS ----------------------------------------------------------

procedure TFSaisieInv.FormCreate(Sender: TObject);
begin
// Création des TOBs
TOBListe := TOB.Create ('LISTEINVENT',nil,-1);
TOBLignes := TOB.Create('Lignes d''inventaire', nil, -1);
TOBLots := TOB.Create('Détail des lots', nil, -1);
TOBDeletedLines := TOB.Create('Lignes effacees', nil, -1);
TOBDeletedLots := TOB.Create('Lots effacees', nil, -1);
TOBCurrentLigne := nil;
G_Inv.GetCellCanvas := DessineCell; // Fameuse bidouille pour les coches vertes
end;

procedure TFSaisieInv.FormDestroy(Sender: TObject);
begin
// Destruction des TOBs
TOBListe.free;
TOBLignes.Free; TOBLignes := nil;
TOBLots.Free; TOBLots := nil;
TOBDeletedLines.Free; TOBDeletedLines := nil;
TOBDeletedLots.Free; TOBDeletedLots := nil;
TOBCurrentLigne := nil;
end;

procedure TFSaisieInv.FormShow(Sender: TObject);
Var FF : string ;
    Dec,ind : integer ;
    Ch : ThNumEdit;
begin
// Initialisations diverses de toute la fiche
BSaisieAveugle := GetParamSoc('SO_GCSAISIEAVEUGLEINV'); //Cache la colonne Stock et la colonne Ecart si égal à True
if BSaisieAveugle then
  begin
  TOTQTESTO.Visible := False; TOTECART.Visible := False;
  LTOTQTE.Left := TOTQTEINV.Left; TOTQTEINV.Left := TOTECART.Left;
  end;
G_Inv.ListeParam := 'GCLISTINVLIGNE';
EtudieColsListe;
//HMTrad.ResizeGridColumns(G_Inv);
FCodeEmpl := '';
InitEntete;
FF:='#,##0.'; Dec:=V_PGI.OkDecV+2;
for ind:=1 to Dec do FF:=FF+'0';
GIL_DPA.Masks.PositiveMask:=FF;
GIL_DPR.Masks.PositiveMask:=FF;
GIL_PMAP.Masks.PositiveMask:=FF;
GIL_PMRP.Masks.PositiveMask:=FF;
GIL_DPAART.Masks.PositiveMask:=FF;
GIL_DPRART.Masks.PositiveMask:=FF;
GIL_PMAPART.Masks.PositiveMask:=FF;
GIL_PMRPART.Masks.PositiveMask:=FF;
GIL_DPASAIS.Masks.PositiveMask:=FF;
GIL_DPRSAIS.Masks.PositiveMask:=FF;
GIL_PMAPSAIS.Masks.PositiveMask:=FF;
GIL_PMRPSAIS.Masks.PositiveMask:=FF;
if Action=taConsult then
   begin
   BAddLigne.Enabled:=False;
   BDelLigne.Enabled:=False;
   bSelectAll.Enabled:=False;
   BInfos.Enabled:=False;
   LConsult.Visible:=True;
   AffecteGrid(G_Inv, Action);
   end
   else AffecteGrid(G_Inv, taModif);
ChargeListe;
with GIL_EMPLACEMENT do if Items.Count > -1 then ItemIndex := 0;
if Action<>taConsult then
   begin
   G_Inv.SetFocus;
   G_Inv.Col := FColonneSaisie;
   end;
HMTrad.ResizeGridColumns(G_Inv);
end;

procedure TFSaisieInv.FormClose(Sender: TObject; var Action: TCloseAction);
begin
VideLesTOB;
G_Inv.VidePile(false);
if IsInside(Self) then Action := caFree;
end;

procedure TFSaisieInv.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
CanClose := ICanContinue; // Je ferme la fiche seulement si j'ai pas annulé et
                          //  si ca a bien enregistré
end;

procedure TFSaisieInv.G_InvRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if not Chg then RefreshPied; // Rafraichissement du pied lorsqu'on change de ligne
// Il semblerait que le Chg soit True si on change aussi de colonne
end;

procedure TFSaisieInv.G_InvCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (G_Inv.Col = FColonneSaisie) or (G_Inv.Col = FColonneDate) then
  begin
    UpdateElipsis;
  end else
  begin
    G_Inv.Col := FColonneSaisie; // Forcer la colonne active à la colonne de saisie
  end;
end;

procedure TFSaisieInv.G_InvCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if csDestroying in ComponentState then exit;

  // Si la TOBization renvoi faux, refuser le changement de cellule
  if ((G_Inv.Col = FColonneSaisie) or (G_Inv.Col = FColonneDate)) and (ARow <= TOBLignes.Detail.Count) then Cancel := not TOBizeLaSaisie(ARow);

end;

procedure TFSaisieInv.G_InvElipsisClick(Sender: TObject);
var vale : double;
  HDATE: THCritMaskEdit;
  Coord: TRect;
begin
  GetLigneTOB(G_Inv.Row);
  if G_Inv.col = FColonneSaisi then
  begin
    vale := Valeur(G_Inv.Cells[FColonneSaisie,G_Inv.Row]);
    // Gestion des lignes cumulées
    EntreeDetailInv(TOBCurrentLigne, TOBLots, TOBDeletedLots,
                TCodeListe.Caption, TLibelle.Caption, TDepot.Caption);
    // Recalcul du cumul pour la ligne
    if (TOBLots.FindFirst(['GLI_ARTICLE'], [TOBCurrentLigne.GetValue('GIL_ARTICLE')], false) = nil) or
   (TOBLots.FindFirst(['GLI_ARTICLE', 'GLI_SAISIINV'], [TOBCurrentLigne.GetValue('GIL_ARTICLE'),'-'], false) <> nil)
  then TOBCurrentLigne.PutValue('GIL_SAISIINV','-')
  else TOBCurrentLigne.PutValue('GIL_SAISIINV','X');
    CalcCumul(TOBCurrentLigne);
    RefreshLigne(G_Inv.Row);
    TOTQTEINV.Value := TOBLignes.Somme('GIL_INVENTAIRE',[''],[''],False); //TOTQTEINV.Value + vale - TOBCurrentLigne.GetValue('GIL_INVENTAIRE');
    TOTECART.Value := TOTQTEINV.Value - TOTQTESTO.Value;
  end else if G_Inv.col = FColonneDate then
  begin
    // DEBUT MODIF CHR
    Coord := G_Inv.CellRect(G_Inv.Col, G_Inv.Row);
    HDATE := THCritMaskEdit.Create(G_Inv);
    HDATE.Parent := G_Inv;
    HDATE.Top := Coord.Top;
    HDATE.Left := Coord.Left;
    HDATE.Width := 3;
    HDATE.Visible := False;
    HDATE.OpeType := otDate;
    //GetDateRecherche(TForm(HDATE.Owner), HDATE);
    // DEBUT MODIF CHR
    if HDATE.Text <> '' then G_Inv.Cells[G_Inv.Col, G_Inv.Row] := HDATE.Text;
    // FIN MODIF CHR
    HDATE.Free;
  end;
end;

procedure TFSaisieInv.G_InvKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var a : integer;
begin
 case Key of
  vk_return : TOBizeLaSaisie(G_Inv.Row); // Si on appuie sur Entrée ca saisit la valeur
  //
  //vk_up     : FColonneSaisie := G_Inv.col;
  //vk_down   : FColonneSaisie := G_Inv.col;
  //vk_Left   : FColonneSaisie := G_Inv.col;
  //vk_right  : FColonneSaisie := G_Inv.col;
  //
  vk_tab : with G_Inv do  // Gestion du Tab / Shift-Tab pour descendre ou remonter
      begin
        if G_Inv.col = FColonneDate then
        begin
      if ssShift in Shift then a := -1
                          else a := 1;
        if (Row+a) > RowCount -1 then
        begin
        	Row := 1;
          Col := FColonneSaisie;
        end else if (Row+a) in [1..RowCount-1] then
        begin
        	Row := Row + a;
          Col := FColonneSaisie;
        end else TOBizeLaSaisie(Row);
      Key := 0;
      end;
      end;
  vk_delete : if [ssCtrl] = Shift then
      begin
      BDelLigne.Click;
      Key := 0;
      end;
{Alt+S}  83 : if Shift=[ssAlt]  then BEGIN Key:=0 ; SetSaisiClick(SetSaisi) ; END ;
 end;
end;

procedure TFSaisieInv.G_InvSorted(Sender: TObject);
var lig : integer;
begin
// Après un tri par clic sur entête de colonne,
for lig := 1 to G_Inv.RowCount-1 do
 if TOB(G_Inv.Objects[0,lig]) = TOBCurrentLigne then
   begin
   G_Inv.Row := lig; // il faut remettre la cellule active sur la bonne ligne
   break;
   end;
end;

procedure TFSaisieInv.GIL_EMPLACEMENTChange(Sender: TObject);
begin
if FCodeEmpl = GIL_EMPLACEMENT.Value then exit;

if ICanContinue then  // Confirmation avant de changer d'emplacement
  begin
  FCodeEmpl := GIL_EMPLACEMENT.Value;
  ChargeListe;
  end
  else GIL_EMPLACEMENT.Value := FCodeEmpl;
end;


// Voir l'article
procedure TFSaisieInv.InfArticleClick(Sender: TObject);
var GaArticle : String;
begin
GetLigneTOB(G_Inv.Row);
GaArticle := GetCurrentLigneValue('GA_ARTICLE');
V_PGI.DispatchTT(7, taConsult, GaArticle, '', 'MONOFICHE');
end;

// Gestion du grisage des items pour le menu info
procedure TFSaisieInv.POPUPopup(Sender: TObject);
begin
GetLigneTOB(G_Inv.Row);
if (GetCurrentLigneValue('GIL_SAISIINV') = 'X') then
  begin
  SetSaisi.Visible := false;
  SetNonSaisi.Enabled := (TOBCurrentLigne <> nil) and (not IsCurrentLigneCumul)
                         and (GetCurrentLigneValue('GIL_INVENTAIRE') = '0');
  end else
  begin
  SetSaisi.Visible := true;
  SetSaisi.Enabled := (TOBCurrentLigne <> nil) and (not IsCurrentLigneCumul);
  end;

SetNonSaisi.Visible := not SetSaisi.Visible;
end;

// Détail des prix
procedure TFSaisieInv.Voirlesprix1Click(Sender: TObject);
begin
GetLigneTOB(G_Inv.Row);
TheTOB := TOBCurrentLigne; // On passe la ligne courante dans TheTOB
if TheTOB <> nil then
  begin
  AGLLanceFiche('GC','GCSAISIEINV_PRIX','','','');
  RefreshPied;
  end;
end;

// Jump to la 1ere ligne non saisie
procedure TFSaisieInv.Allerlapremirelignenonsaisie1Click(Sender: TObject);
var lig, ac, ar : integer;
    C : boolean;
begin
for lig := 1 to G_Inv.RowCount-1 do
  begin
  GetLigneTOB(lig);
  if GetCurrentLigneValue('GIL_SAISIINV') <> 'X' then
    begin
    G_Inv.SetFocus;
    ac := G_Inv.Col; ar := G_Inv.Row;
    G_Inv.Row := lig;
    C := false;
    G_InvRowEnter(Sender, G_Inv.Row, C, false); // Déclanchement manuel des évènements (grr)
    G_InvCellEnter(nil, AC, AR, C);
    break;
    end;
  end;
end;

// Marquer la ligne comme saisie / non saisie
procedure TFSaisieInv.SetSaisiClick(Sender: TObject);
var C : Char;
begin
GetLigneTOB(G_Inv.Row);
if TOBCurrentLigne = nil then exit;
if Sender = SetSaisi then C := 'X'
                     else C := '-';
TOBCurrentLigne.PutValue('GIL_SAISIINV',C);
RefreshLigne(G_Inv.Row);
end;


procedure TFSaisieInv.SetAllSaisieClick(Sender: TObject);
Var ind : integer ;
    TOBLInv : TOB ;
begin
for ind:=0 to G_Inv.RowCount-1 do
    begin
    TOBLInv := TOB(G_Inv.Objects[0,ind]);
    if TOBLInv = nil then Continue;
    TOBLInv.PutValue('GIL_SAISIINV','X');
    RefreshLigne(ind);
    end;
end;

// Recherche
procedure TFSaisieInv.BChercherClick(Sender: TObject);
begin
if G_Inv.RowCount < 3 then Exit;
FFirstFind := true;
FindLigne.Execute;
end;

procedure TFSaisieInv.FindLigneFind(Sender: TObject);
begin
Rechercher(G_Inv, FindLigne, FFirstFind);
end;

// Fenêtre popup du détail des prix
procedure TFSaisieInv.BDetailPrixClick(Sender: TObject);
begin
TDetailPrix.Visible := BDetailPrix.Down;
end;

procedure TFSaisieInv.TDetailPrixClose(Sender: TObject);
begin
BDetailPrix.Down := False;
end;

// Ajout article
procedure TFSaisieInv.BAddLigneClick(Sender: TObject);
var CodArtick : String;
    Q : TQuery;
begin
if ctxMode in V_PGI.PGIContexte
   then CodArtick := DispatchArtMode(1,'','','XX_WHERE=GA_TENUESTOCK="X"')
   else CodArtick := AGLLanceFiche('GC','GCARTICLE_RECH','','','XX_WHERE=GA_TENUESTOCK="X"');
if CodArtick = '' then exit;

if ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE="'+CodArtick+'" AND GA_STATUTART="GEN"') then
  begin
  TheTOB := TOB.Create('',Nil,-1);
  AglLanceFiche ('GC','GCSELECTDIM','','', 'GA_ARTICLE='+CodArtick+';ACTION=SELECT;CHAMP= ');
  if TheTOB = nil then exit;
  CodArtick := TheTOB.Detail[0].GetValue('GA_ARTICLE');
  end;

// Si l'artick vient d'être effacé on le désefface, c'est plus simple
//   et sinon on se fait avoir par le "if" d'après
(*if TOBDeletedLines.FindFirst(['GA_ARTICLE'], [CodArtick], false) <> nil then
  begin
  UnDeleteLigne(CodArtick);
  exit;
  end;*)

Q := OpenSQL('SELECT GIE_CODELISTE FROM LISTEINVLIG LEFT JOIN LISTEINVENT ON GIL_CODELISTE=GIE_CODELISTE '+
             'WHERE GIE_VALIDATION<>"X" AND GIL_ARTICLE="'+CodArtick+'" AND GIE_DEPOT="'+FDepot+'"', true,-1,'',true);
if (not Q.EOF) and (TOBDeletedLines.FindFirst(['GA_ARTICLE'], [CodArtick], false) = nil) then
  begin
  PGIInfo('L''article est déjà dans la liste '+Q.FindField('GIE_CODELISTE').AsString+' sur le même dépôt','');
  Ferme(Q);
  end else
  begin
  Ferme(Q);
  InsereLigne(CodArtick);
  end;
end;

// Suppression Ligne
procedure TFSaisieInv.BDelLigneClick(Sender: TObject);
begin
SupprimeLigne(G_Inv.Row);
end;

// Validation
procedure TFSaisieInv.BValiderClick(Sender: TObject);
begin
ValiderSaisie;
end;

procedure TFSaisieInv.BAbandonClick(Sender: TObject);
begin
if Action<>taConsult then TOBizeLaSaisie(G_Inv.Row);
end;


// ----------- TOF pour la saisie des prix -------------------------------------

procedure TOF_PrixInv.OnLoad;
Var FF : string ;
    Dec,ind : integer ;
    Ch : ThNumEdit;
begin
inherited;
FF:='#,##0.'; Dec:=V_PGI.OkDecV+2;
for ind:=1 to Dec do FF:=FF+'0';
Ch:=ThNumEdit(GetControl('Gil_DPA'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_DPR'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_PMAP'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_PMRP'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_DPAART'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_DPRART'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_PMAPART'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_PMRPART'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_DPASAIS'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_DPRSAIS'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_PMAPSAIS'));
Ch.Masks.PositiveMask:=FF;
Ch:=ThNumEdit(GetControl('GIL_PMRPSAIS'));
Ch.Masks.PositiveMask:=FF;
LaTOB.PutEcran(Ecran); // LaTOB a pris la valeur de TheTOB, donc pointe sur TOBCurrentLigne
end;

procedure TOF_PrixInv.OnUpdate;
begin
LaTOB.GetEcran(Ecran); // Récupération de ce qui a été saisi
inherited;
end;

procedure TFSaisieInv.bSelectAllClick(Sender: TObject);
var i_row : integer;
    vale : double;
    AfficheInfo : boolean;
begin
AfficheInfo := false;
if bSelectAll.Down then
begin
  if PGIAsk(TraduireMemoire(TexteMessage[2]) + Chr(10) + TraduireMemoire(TexteMessage[4])
     ,TraduireMemoire(TexteMessage[1]))<>mrYes then exit ;
  for i_row := 1 to G_Inv.RowCount -1 do
  begin
    GetLigneTOB(i_row);
    if (TOBCurrentLigne <> nil) and (TOBCurrentLigne.GetValue('GIL_SAISIINV') <>'X') then
    begin
      vale := Valeur(G_Inv.Cells[FColonneOrdi,i_row]);
      if Vale > 0 then
      begin
        if (IsCurrentLigneCumul) and (TOBLots.Somme('GLI_INVENTAIRE', ['GLI_ARTICLE'], [TOBCurrentLigne.GetValue('GIL_ARTICLE')], false) <> vale)
        then AfficheInfo := true
        else
        begin // Si c'est pas une ligne cumulée
          TOTQTEINV.Value := TOTQTEINV.Value + vale - TOBCurrentLigne.GetValue('GIL_INVENTAIRE');
          TOBCurrentLigne.PutValue('GIL_INVENTAIRE',vale);
          TOBCurrentLigne.PutValue('GIL_SAISIINV','X');
        end ;
        RefreshLigne(i_row);
      end;
    end;
  end;
  if AfficheInfo then PGIInfo('Un ou plusieurs articles sont gérés par lot,#13vous devez mettre à jour les lignes correspondantes manuellement ','');
end
else
begin
  if PGIAsk(TraduireMemoire(TexteMessage[3]),TraduireMemoire(TexteMessage[1]))<>mrYes then exit ;
  for i_row := 1 to G_Inv.RowCount -1 do
  begin
    GetLigneTOB(i_row);
    if (TOBCurrentLigne <> nil) and (TOBCurrentLigne.GetValue('GIL_SAISIINV') = 'X') then
    begin
      TOTQTEINV.Value := TOTQTEINV.Value - TOBCurrentLigne.GetValue('GIL_INVENTAIRE');
      TOBCurrentLigne.PutValue('GIL_INVENTAIRE',0);
      TOBCurrentLigne.PutValue('GIL_SAISIINV','-');
      RefreshLigne(i_row);
    end;
  end;
end;
TOTECART.Value := TOTQTEINV.Value - TOTQTESTO.Value;
end;

procedure TFSaisieInv.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSaisieInv.BActualiseClick(Sender: TObject);
var Indice : integer;
		TOBL : TOB;
begin
//
	for Indice := 0 to TOBLignes.detail.count -1 do
  begin
  	TOBL := TOBLIgnes.detail[Indice];
    if VarIsNull(TOBL.getValue('GQ_DISPO')) then TOBL.PutValue('GIL_QTEPHOTOINV',0)
    																				else TOBL.PutValue('GIL_QTEPHOTOINV',TOBL.getValue('GQ_PHYSIQUE'));
		CalculeEcart ('GIL',TOBL);
	end;
	TOBLignes.PutGridDetail(G_Inv, false, true, G_Inv.Titres[0], true);
  TOBListe.putValue('GIE_DATEINVENTAIRE',V_PGI.DateEntree);
  TDate.Caption := 'Inventaire du '+DateToStr (V_PGI.DateEntree);
end;

procedure TFSaisieInv.G_InvDblClick(Sender: TObject);
var HDATE: THCritMaskEdit;
  	Coord: TRect;
begin
  if G_Inv.col = FColonneDate then
  begin
    // DEBUT MODIF CHR
    Coord := G_Inv.CellRect(G_Inv.Col, G_Inv.Row);
    HDATE := THCritMaskEdit.Create(G_Inv);
    HDATE.Parent := G_Inv;
    HDATE.Top := Coord.Top;
    HDATE.Left := Coord.Left;
    HDATE.Width := 3;
    HDATE.Visible := False;
    HDATE.OpeType := otDate;
    //GetDateRecherche(TForm(HDATE.Owner), HDATE);
    // DEBUT MODIF CHR
    if HDATE.Text <> '' then G_Inv.Cells[G_Inv.Col, G_Inv.Row] := HDATE.Text;
    // FIN MODIF CHR
    HDATE.Free;
  end;
end;

Initialization
RegisterClasses([TOF_PrixInv]);
RegisterAGLProc('EntreeSaisieInv',false,1,AGLEntreeSaisieInv);
end.
