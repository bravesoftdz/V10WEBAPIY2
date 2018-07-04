unit consultMouv;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, hmsgbox,  Hqry, StdCtrls, HTB97,
  ExtCtrls, HPanel, Grids,   Hent1, UIUtil, Mask,
  Hctrls, Math, SaisUtil, UTOB, LookUp,
{$IFDEF EAGLCLIENT}
  eMul,eFichList,
{$ELSE}
  Mul,FichList,Db,DBTables,DBCtrls,DBGrids,HDB,
{$ENDIF}
  Menus, HRichOLE, ComCtrls,UtilPGI,Ent1,
  HRichEdt, ColMemo, EntGC, AglInitGC, TarifUtil, Facture, FactComm, FactUtil,
  StockUtil, HFLabel;

procedure EntreeConsultMouv;

Const
     MOU_Fleche         : integer = 0;
     MOU_Depot          : integer = 0;
     MOU_date           : integer = 0;
     MOU_QualifMvt      : integer = 0;
     MOU_QteStock       : integer = 0;
     NbRowsInit = 50;
     NbRowsPlus = 20;

type
  TFConsultMouv = class(TFMul)
    G_ConsultMouv: THGrid;
    XX_WHERE_CONTROLE: TEdit;
    TGL_DEPOT: THLabel;
    GL_DEPOT: THValComboBox;
    GL_CODEARTICLE: THCritMaskEdit;
    TGL_CODEARTICLE: THLabel;
    GL_DATEPIECE: THCritMaskEdit;
    TGL_DATEPIECE: THLabel;
    GL_DATEPIECE_: THCritMaskEdit;
    TGL_DATEPIECE_: THLabel;
    GL_CODEARTICLE_: THCritMaskEdit;
    TGL_CODEARTICLE_: THLabel;
    TGA_FAMILLENIV1: TLabel;
    TGA_FAMILLENIV2: TLabel;
    TGA_FAMILLENIV3: TLabel;
    GA_FAMILLENIV1: THValComboBox;
    GA_FAMILLENIV2: THValComboBox;
    GA_FAMILLENIV3: THValComboBox;
    PComplement2: TTabSheet;
    Complement2: TPanel;
    TGL_LIBREART1: TFlashingLabel;
    GL_LIBREART1: THValComboBox;
    TGL_LIBREART2: TFlashingLabel;
    GL_LIBREART2: THValComboBox;
    TGL_LIBREART3: TFlashingLabel;
    GL_LIBREART3: THValComboBox;
    TGL_LIBREART4: TFlashingLabel;
    GL_LIBREART4: THValComboBox;
    TGL_LIBREART5: TFlashingLabel;
    GL_LIBREART5: THValComboBox;
    TGL_LIBREART6: TFlashingLabel;
    GL_LIBREART6: THValComboBox;
    TGL_LIBREART7: TFlashingLabel;
    GL_LIBREART7: THValComboBox;
    TGL_LIBREART8: TFlashingLabel;
    GL_LIBREART8: THValComboBox;
    TGL_LIBREART9: TFlashingLabel;
    GL_LIBREART9: THValComboBox;
    TGL_LIBREARTA: TFlashingLabel;
    GL_LIBREARTA: THValComboBox;
    GL_QUALIFMVT: THValComboBox;
    TGL_QUALIFMVT: THLabel;
    procedure BChercheClick (Sender: TObject); override;
    procedure FormClose (Sender: TObject; var Action: TCloseAction);
    procedure FormCreate (Sender: TObject); override;
    procedure FormShow (Sender: TObject);
    procedure G_ConsultMouvRowEnter(Sender: TObject; Ou: Integer;
                                    var Cancel: Boolean; Chg: Boolean);
    procedure G_ConsultMouvRowExit(Sender: TObject; Ou: Integer;
                                   var Cancel: Boolean; Chg: Boolean);
    procedure GL_CODEARTICLEElipsisClick(Sender: TObject);
    procedure GL_CODEARTICLE_ElipsisClick(Sender: TObject);
    procedure G_ConsultMouvDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure BParamListeClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas; AState: TGridDrawState);
  public
    { Déclarations publiques }
    LesColConsultMouv : String;
    DEV : RDEVISE;
    TobLigne, TobLigneAff : TOB;
    iTableLigne : integer;
    stWhereAffiche : string;
    stWhereNat : string;
// Actions liées au Grid
    procedure EtudieColsListe;
// Initialisation
    Procedure DepileTOBLigne;
    procedure InitialiseEntete;
    procedure InitialiseGrille;
// Gestion des données
    procedure AjoutSelect (var stSelect : string; stChamp : string);
    procedure ChargeMouvement;
    function CalcPrixRatio (dPrix, dUniteStock, dUnite, dPrixPourQte : Double) : Double;
    Function CalcRatioMesure (stCat, stMesure : String) : Double;
    procedure CopieLaLigne (index : integer);
    procedure DecodeRefLignePiece (stReferenceLignePiece : String; Var CleDoc : R_CleDoc );
    function DetermineNatureAutorisee : string;
    function Evaluedate (St : String) : TDateTime;
    procedure LoadLesTobLigne;
    procedure TraiteLesLignes;
    function TypePieceStockPhy (stNaturePiece : string) : boolean;
  end;

var
  FConsultMouv: TFConsultMouv;

implementation

{$R *.DFM}

procedure EntreeConsultMouv;
var FF : TFConsultMouv;
    PPANEL  : THPanel;
begin
SourisSablier;
FF := TFConsultMouv.Create(Application);
PPANEL := FindInsidePanel;
if PPANEL = Nil then
   begin
    try
      FF.ShowModal;
    finally
      FF.Free;
    end;
   SourisNormale;
   end else
   begin
   InitInside (FF, PPANEL);
   FF.Show;
   end;
end;

procedure TFConsultMouv.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
G_ConsultMouv.VidePile(True);
DepileTOBLigne;
TobLigne.Free;
TobLigne := Nil;
TobLigneAff.Free;
TobLigneAff := Nil;
if IsInside(Self) then Action := caFree;
  inherited;
end;

procedure TFConsultMouv.FormCreate(Sender: TObject);
begin
  inherited;
iTableLigne := PrefixeToNum ('GL');
TobLigne := TOB.Create ('', Nil, -1);
TobLigneAff := TOB.Create ('', Nil, -1);
stWhereNat := DetermineNatureAutorisee;
end;

procedure TFConsultMouv.FormShow(Sender: TObject);
var iInd : integer;
begin
for iInd := 1 to 3 do
    begin
    THLabel(FindComponent('TGA_FAMILLENIV' + IntToStr (iInd))).Caption :=
        RechDom ('GCLIBFAMILLE', 'LF' + IntToStr (iInd), False);
    end;

for iInd := 1 to 9 do
    begin
    THLabel(FindComponent ('TGL_LIBREART' + IntToStr(iInd))).Caption :=
        RechDom('GCZONELIBRE','AT'+IntToStr(iInd),True);
    end;
THLabel(FindComponent ('TGL_LIBREARTA')).Caption := RechDom('GCZONELIBRE','ATA',True);

Q.Liste := 'GCMULCONSUMOUV';
TypeAction := TActionFiche (Action);
DEV.Code := V_PGI.DevisePivot;
GetInfosDevise (DEV);
FAGL := True;
FiltreDisabled := True;
EtudieColsListe;
AffecteGrid (G_ConsultMouv, taConsult);
InitialiseEntete;
G_ConsultMouv.GetCellCanvas := DessineCell;
inherited;
end;

{========================================================================================}
{========================= Actions liées au Grid ========================================}
{========================================================================================}
procedure TFConsultMouv.EtudieColsListe;
Var NomCol, LesCols : String;
    icol, ichamp : Integer;
begin
LesCols := G_ConsultMouv.Titres[0];
LesColConsultMouv := G_ConsultMouv.Titres[0];
iCol := 0;
Repeat
    NomCol := uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol<>'' Then
        begin
        ichamp := ChampToNum(NomCol);
        if ichamp>=0 Then
           begin
           if Nomcol='GL_DEPOT' then MOU_Depot := icol else
           if NomCol='GL_DATEPIECE' Then MOU_Date := icol else
           if NomCol='GL_QUALIFMVT' Then MOU_QualifMvt := icol else
           if NomCol='GL_QTESTOCK' Then MOU_QteStock := icol;
           end;
        end;
    Inc(icol);
    Until ((LesCols='') OR (NomCol=''));
MOU_Fleche := 0;
G_ConsultMouv.ColWidths[MOU_Fleche] := 15;
G_ConsultMouv.ColLengths [MOU_Fleche] := -1;
G_ConsultMouv.ColColors [MOU_Fleche] := clActiveBorder;
G_ConsultMouv.ColFormats[MOU_Date] := '';
end;

{========================================================================================}
{===================      Initialisation    =============================================}
{========================================================================================}
Procedure TFConsultMouv.DepileTOBLigne;
var Index : integer;
begin
for Index := TobLigne.Detail.Count - 1 Downto 0 do
    begin
    TobLigne.Detail[Index].Free;
    end;
for Index := TobLigneAff.Detail.Count - 1 Downto 0 do
    begin
    TobLigneAff.Detail[Index].Free;
    end;
end;

procedure TFConsultMouv.DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
Var Coord : TRect;
    ST : string;
begin
if (ACol = MOU_Fleche) AND (ARow > 0) then
    begin
    Coord := G_ConsultMouv.CellRect (ACol, ARow);
    Canvas.Font.Name := 'Wingdings 3';
    Canvas.Font.Size := 10;
    Canvas.Font.Style := [fsBold];
    st := G_ConsultMouv.Cells [ACol, Arow];
    Canvas.TextOut ( (Coord.Left+Coord.Right) div 2 - Canvas.TextWidth(st) div 2,
                     (Coord.Top+ Coord.Bottom) div 2 - Canvas.TextHeight(st) div 2, st);
    end;
end;

procedure TFConsultMouv.InitialiseEntete;
begin
InitialiseGrille;
end;

procedure TFConsultMouv.InitialiseGrille;
begin
G_ConsultMouv.VidePile (False);
end;

{========================================================================================}
{================================= Gestion des Données ==================================}
{========================================================================================}
procedure TFConsultMouv.AjoutSelect (var stSelect : string; stChamp : string);
begin
if Pos (stChamp, stSelect) = 0 then stSelect := stSelect + ', ' + stChamp;
end;

function TFConsultMouv.CalcPrixRatio (dPrix, dUniteStock, dUnite, dPrixPourQte : Double) : Double;
begin
if dPrixPourQte = 0.0 then Result := dPrix
else result := (dPrix * dUniteStock)/(dPrixPourQte * dUnite);
end;

Function TFConsultMouv.CalcRatioMesure (stCat, stMesure : String) : Double;
var TobM : TOB;
    dRatio : Double;
begin
TobM := VH_GC.TOBMEA.FindFirst(['GME_QUALIFMESURE','GME_MESURE'], [stCat, stMesure], False);
dRatio := 0;
if TobM <> nil then dRatio := TobM.GetValue('GME_QUOTITE');
if dRatio = 0 then dRatio := 1.0;
result := dRatio;
end;

procedure TFConsultMouv.ChargeMouvement;
begin
DepileTobLigne;
InitialiseGrille;
LoadLesTobLigne;
TobLigneAff.PutGridDetail (G_ConsultMouv, True, True, LesColConsultMouv, False);
G_ConsultMouv.RowCount := TobLigneAff.Detail.Count + 1;
G_ConsultMouv.SetFocus;
G_ConsultMouv.Cells [MOU_Fleche, 1] := '„';
G_ConsultMouv.Cells [MOU_Fleche, 0] := '';
if MOU_QualifMvt <> 0 then G_ConsultMouv.Cells [MOU_QualifMvt, 0] := 'Mouvement';
if MOU_QteStock <> 0 then G_ConsultMouv.Cells [MOU_QteStock, 0] := 'Quantité';
Z_SQL.Text := stWhereAffiche;
end;

procedure TFConsultMouv.CopieLaLigne (index : integer);
var Tobl : TOB;
    dCoef : double;
begin
TobL := TobLigneAff.FindFirst (['GL_DATEPIECE', 'GL_NATUREPIECEG', 'GL_NUMERO', 'GL_DEPOT',
                                'GL_QUALIFMVT', 'GL_ARTICLE'],
                               [TobLigne.Detail[index].GetValue('GL_DATEPIECE'),
                                TobLigne.Detail[index].GetValue('GL_NATUREPIECEG'),
                                TobLigne.Detail[index].GetValue('GL_NUMERO'),
                                TobLigne.Detail[index].GetValue('GL_DEPOT'),
                                TobLigne.Detail[index].GetValue('GL_QUALIFMVT'),
                                TobLigne.Detail[index].GetValue('GL_ARTICLE')], True);
if Tobl = Nil then
    begin
    Tobl := Tob.Create ('LIGNE', TobLigneAff, -1);
    Tobl.Dupliquer (TobLigne.Detail[index], true, true);
    TobL.AddChampSup ('RATIO', False);
    TobL.AddChampSup ('VA', False);
    Tobl.PutValue('VA', GetInfoParPiece (TOBL.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT'));
    TobL.PutValue ('RATIO', GetRatio (Tobl, nil, trsStock));
    Tobl.PutValue('GL_QTESTOCK', Tobl.GetValue ('GL_QTESTOCK')/TobL.GetValue('RATIO'));
    if Tobl.GetValue ('VA') = 'ACH' then
        dCoef := CalcRatioMesure('PIE', Tobl.GetValue('GL_QUALIFQTEACH'))
    else dCoef := CalcRatioMesure('PIE', Tobl.GetValue('GL_QUALIFQTEVTE'));
    if pos ('GL_PUHT', LesColConsultMouv) > 0 then
        begin
        Tobl.PutValue('GL_PUHT',
                      CalcPrixRatio(Tobl.GetValue('GL_PUHT'),
                                    CalcRatioMesure('PIE', Tobl.GetValue('GL_QUALIFQTESTO')),
                                    dCoef, TobL.GetValue('GL_PRIXPOURQTE')));
        end;
    end else
    begin
    if TobL.GetValue ('VA') = 'ACH' then
        dCoef := CalcRatioMesure('PIE', TobLigne.Detail[index].GetValue('GL_QUALIFQTEACH'))
    else dCoef := CalcRatioMesure('PIE', TobLigne.Detail[index].GetValue('GL_QUALIFQTEVTE'));
    Tobl.PutValue ('GL_QTESTOCK',
                   TobL.GetValue ('GL_QTESTOCK') +
                       (TobLigne.Detail[index].GetValue('GL_QTESTOCK')/
                            Tobl.GetValue('RATIO')));
    if pos ('GL_PUHT', LesColConsultMouv) > 0 then
        begin
        Tobl.PutValue('GL_PUHT',
                      Tobl.GetValue ('GL_PUHT') +
                      CalcPrixRatio(TobLigne.Detail[index].GetValue('GL_PUHT'),
                                    CalcRatioMesure('PIE',
                                                    TobLigne.Detail[index].GetValue('GL_QUALIFQTESTO')),
                                    dCoef, TobLigne.Detail[index].GetValue('GL_PRIXPOURQTE')));
        end;
    end;
end;

Procedure TFConsultMouv.DecodeRefLignePiece (stReferenceLignePiece : String; Var CleDoc : R_CleDoc);
Var stCleDoc : String;
begin
FillChar (CleDoc, Sizeof(CleDoc), #0);
stCleDoc := stReferenceLignePiece;
CleDoc.DatePiece := EvalueDate(ReadTokenSt(stCleDoc));
CleDoc.NaturePiece := ReadTokenSt(stCleDoc);
CleDoc.Souche := ReadTokenSt(stCleDoc);
CleDoc.NumeroPiece := StrToInt(ReadTokenSt(stCleDoc));
CleDoc.Indice := StrToInt(ReadTokenSt(stCleDoc));
CleDoc.NumLigne := StrToInt(ReadTokenSt(stCleDoc));
end;

function TFConsultMouv.DetermineNatureAutorisee : string;
var Index : integer;
begin
Result := '';
for Index := 0 to VH_GC.TobParPiece.Detail.Count - 1 do
    begin
    if (pos ('PHY;', VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_QTEPLUS')) > 0) or
       (pos ('PHY;', VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_QTEMOINS')) > 0) then
        begin
        if Result <> '' then Result := Result + ') OR'
        else Result := '(';
        Result := Result + '(GL_NATUREPIECEG="' + VH_GC.TobParPiece.Detail[Index].GetValue ('GPP_NATUREPIECEG') + '"';
        end;
    end;
if Result <> '' then Result := Result + '))';
end;

Function TFConsultMouv.Evaluedate (St : String) : TDateTime;
Var dd,mm,yy : Word;
begin
Result := 0; if St='' then Exit;
dd := StrToInt(Copy(St,1,2)); mm := StrToInt(Copy(St,3,2)); yy := StrToInt(Copy(St,5,4));
Result := Encodedate(yy,mm,dd);
end;

procedure TFConsultMouv.LoadLesTobLigne;
var stWhere, stSelect, stOrder, stLesCols, stChamp : string;
    TSql : TQuery;
    iInd : integer;
begin
stWhere := RecupWhereCritere(Pages) ;
//TSql := TQuery.Create (Nil);
//TSql.SQL.Add ('');
//RecupWhereSQL (Q, TSql);
//stWhere := TSql.Text;
//TSql.Free;
stSelect := '';
stLesCols := LesColConsultMouv;
stChamp := ReadTokenSt (stLesCols);
while stChamp <> '' do
    begin
    if stSelect <> '' then stSelect := stSelect + ', ';
    stSelect := stSelect + stChamp;
    stChamp := ReadTokenSt (stLesCols);
    end;
stWhereAffiche := 'SELECT ' + stSelect +
                  ' FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG';
if (GA_FAMILLENIV1.Value <> '') or (GA_FAMILLENIV2.Value <> '') or
   (GA_FAMILLENIV3.Value <> '') then
    begin
    stWhereAffiche := stWhereAffiche +
                        ' LEFT JOIN ARTICLE ON GL_ARTICLE=GA_ARTICLE';
    end;
stWhereAffiche := stWhereAffiche +
                  ' WHERE (GPP_QTEMOINS LIKE "%PHY%" OR GPP_QTEPLUS LIKE "%PHY%")' +
                  ' AND GL_TENUESTOCK="X" AND GL_QUALIFMVT<>"ANN"';
stSelect := 'SELECT ' + stSelect;
AjoutSelect (stSelect, 'GL_QTERELIQUAT');
AjoutSelect (stSelect, 'GL_QTERESTE');
AjoutSelect (stSelect, 'GL_PIECEPRECEDENTE');
AjoutSelect (stSelect, 'GL_INDICEG');
AjoutSelect (stSelect, 'GL_NATUREPIECEG');
AjoutSelect (stSelect, 'GL_ARTICLE');
AjoutSelect (stSelect, 'GL_SOUCHE');
AjoutSelect (stSelect, 'GL_NUMERO');
AjoutSelect (stSelect, 'GL_DEPOT');
AjoutSelect (stSelect, 'GL_PRIXPOURQTE');
AjoutSelect (stSelect, 'GL_QUALIFQTEVTE');
AjoutSelect (stSelect, 'GL_QUALIFQTESTO');
AjoutSelect (stSelect, 'GL_QUALIFQTEACH');
stSelect := stSelect +
            ' FROM LIGNE LEFT JOIN PARPIECE ON GL_NATUREPIECEG=GPP_NATUREPIECEG';
if (GA_FAMILLENIV1.Value <> '') or (GA_FAMILLENIV2.Value <> '') or
   (GA_FAMILLENIV3.Value <> '') then
    begin
    stSelect := stSelect + ' LEFT JOIN ARTICLE ON GL_ARTICLE=GA_ARTICLE';
    end;
stSelect := stSelect +
            ' WHERE (GPP_QTEMOINS LIKE "%PHY%" OR GPP_QTEPLUS LIKE "%PHY%")' +
            ' AND GL_TENUESTOCK="X" AND GL_QUALIFMVT<>"ANN"';
stOrder := ' ORDER BY GL_DATEPIECE, GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO';

for iInd := 1 to 9 do
    begin
    if THValComboBox(FindComponent ('GL_LIBREART' + IntToStr(iInd))).Value <> '' then
        begin
        stSelect := stSelect + ' AND GL_LIBREART' + IntToStr(iInd) + '="' +
            THValComboBox(FindComponent ('GL_LIBREART' + IntToStr(iInd))).Value + '"';
        end;
    end;
if THValComboBox(FindComponent ('GL_LIBREARTA')).Value <> '' then
    begin
    stSelect := stSelect + ' AND GL_LIBREARTA="' +
        THValComboBox(FindComponent ('GL_LIBREARTA')).Value + '"';
    end;

//while pos (#$D, stWhere) > 0 do delete (stWhere, pos (#$D, stWhere), 1);
//while pos (#$A, stWhere) > 0 do delete (stWhere, pos (#$A, stWhere), 1);
if stWhere <> '' then
    begin
    delete (stWhere, pos ('WHERE ', stWhere), 6);
    while pos('''', stWhere) > 0 do stWhere[pos('''', stWhere)] := '"';
    stWhereAffiche := stWhereAffiche + ' AND ' + stWhere + stOrder;
    stWhere := stSelect + ' AND ' + stWhere + stOrder;
    end else
    begin
    stWhere := stSelect + stOrder;
    stWhereAffiche := stWhereAffiche + stOrder;
    end;


TSql := OpenSql (stWhere, True);
if not TSql.Eof then
    begin
    TobLigne.LoadDetailDB ('LIGNE', '', '', TSql, False);
    TraiteLesLignes;
    end else TOB.Create ('LIGNE', TobLigneAff, 0);
Ferme (TSql);

Z_SQL.Text := stWhere;
end;

procedure TFConsultMouv.TraiteLesLignes;
Var stPiecePrecedente, stWhere : string;
    TSql : TQuery;
    iIndex : integer;
    QteStockPre, QteMouvement : double;
    CleDocPrec : R_CleDoc;
begin
for iIndex := 0 to TobLigne.Detail.Count - 1 do
    begin
    QteStockPre := 0.0;
    if trim (TobLigne.Detail[iIndex].GetValue ('GL_CODEARTICLE')) <> '' then
        begin
        if TobLigne.Detail[iIndex].GetValue ('GL_INDICEG') = 0 then
            begin
            stPiecePrecedente := TobLigne.Detail[iIndex].GetValue ('GL_PIECEPRECEDENTE');
            if stPiecePrecedente <> '' then
                begin
                DecodeRefLignePiece (stPiecePrecedente, CleDocPrec);
                if TypePieceStockPhy (CleDocPrec.NaturePiece) then
                    begin
                    stWhere := 'SELECT GL_QTESTOCK, GL_QTERELIQUAT, GL_QTERESTE FROM LIGNE WHERE ' +
                                'GL_NATUREPIECEG="' + CleDocPrec.NaturePiece +
                                '" AND GL_DATEPIECE="' + UsDateTime (CleDocPrec.DatePiece) +
                                '" AND GL_SOUCHE="' + CleDocPrec.Souche +
                                '" AND GL_NUMERO=' + inttostr (CleDocPrec.NumeroPiece) +
                                ' AND GL_INDICEG=' + inttostr (CleDocPrec.Indice) +
                                ' AND GL_NUMLIGNE=' + inttostr (CleDocPrec.NumLigne);
                    TSql := OpenSql (stWhere, True);
                    if not TSql.Eof then
                        begin
                        QteStockPre := StrToFloat (TSql.FindField('GL_QTESTOCK').AsString);
                        end;
                    ferme (TSql);
                    end;
                QteMouvement := TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') - QteStockPre;
                end else
                begin
                QteMouvement := TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') -
                                    TobLigne.Detail[iIndex].GetValue ('GL_QTERESTE');
                end;
            TobLigne.Detail [iIndex].PutValue ('GL_QTESTOCK', QteMouvement);
            end;
        if (TobLigne.Detail[iIndex].GetValue ('GL_QTESTOCK') > 0) or
           (stPiecePrecedente = '') then
                CopieLaLigne (iIndex);
        end;
    end;

if (MOU_Depot <> 0) or (MOU_QualifMvt <> 0) then
   begin
    for iIndex := 0 to TobLigneAff.Detail.Count - 1 do
        begin
        if MOU_Depot <> 0 then
            begin
            stWhere := 'SELECT GDE_LIBELLE FROM DEPOTS WHERE GDE_DEPOT="' +
                       TobLigneAff.Detail[iIndex].GetValue ('GL_DEPOT') + '"';
            TSql := Opensql (stWhere, True);
            if not TSql.Eof then
               begin
               TobLigneAff.Detail[iIndex].PutValue (
                   'GL_DEPOT', TSql.FindField('GDE_LIBELLE').AsString);
               end;
               ferme (TSql);
            end;
        if MOU_QualifMvt <> 0 then
            begin
            if TobLigneAff.Detail[iIndex].GetValue ('GL_QUALIFMVT') <> '' then
               TobLigneAff.Detail[iIndex].PutValue (
                   'GL_QUALIFMVT',
                   RechDom ('GCQUALIFMVT',
                            TobLigneAff.Detail[iIndex].GetValue ('GL_QUALIFMVT'), False));
            end;
        end;
    end;
end;

function TFConsultMouv.TypePieceStockPhy (stNaturePiece : string) : boolean;
var TSql : TQuery;
    stSelect : string;
begin
stSelect := 'SELECT GPP_QTEMOINS FROM PARPIECE ' +
            'WHERE GPP_NATUREPIECEG="' + stNaturePiece + '"' +
            ' AND (GPP_QTEMOINS LIKE "%PHY%" OR GPP_QTEPLUS LIKE "%PHY%")';
TSql := OpenSql (stSelect, True);
if not TSql.Eof then
    Result := True
    else Result := False;
Ferme (TSql);
end;

{========================================================================================}
{======================================= Evenements de la Grid ==========================}
{========================================================================================}
procedure TFConsultMouv.BChercheClick(Sender: TObject);
begin
XX_WHERE_CONTROLE.text := 'GL_INDICEG=-1';
  inherited;
XX_WHERE_CONTROLE.text := 'GL_QUALIFMVT<>"ANN"';
ChargeMouvement;
HMTrad.ResizeGridColumns (G_ConsultMouv);
end;

procedure TFConsultMouv.G_ConsultMouvRowEnter(Sender: TObject; Ou: Integer;
                                              var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
G_ConsultMouv.Cells [MOU_Fleche, Ou] := '„';
end;

procedure TFConsultMouv.G_ConsultMouvRowExit(Sender: TObject; Ou: Integer;
                                             var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;
  inherited;
G_ConsultMouv.Cells [MOU_Fleche, Ou] := ' ';
end;

procedure TFConsultMouv.GL_CODEARTICLEElipsisClick(Sender: TObject);
begin
  inherited;
DispatchRecherche (GL_CODEARTICLE, 1, '', '', '');
GL_CODEARTICLE.Text := copy (GL_CODEARTICLE.Text, 1, 18);
GL_CODEARTICLE.Text := Trim (GL_CODEARTICLE.Text);
end;

procedure TFConsultMouv.GL_CODEARTICLE_ElipsisClick(Sender: TObject);
begin
  inherited;
DispatchRecherche (GL_CODEARTICLE_, 1, '', '', '');
GL_CODEARTICLE_.Text := copy (GL_CODEARTICLE_.Text, 1, 18);
GL_CODEARTICLE_.Text := Trim (GL_CODEARTICLE_.Text);
end;

procedure TFConsultMouv.G_ConsultMouvDblClick(Sender: TObject);
var CleDoc : R_CleDoc;
    stNature, stDate, stSouche, stNumero, stIndiceg : string;
begin
inherited;
stNature := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NATUREPIECEG');
stDate := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_DATEPIECE');
stSouche := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_SOUCHE');
stNumero := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NUMERO'));
stIndiceg := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_INDICEG'));
if (stNature <> '') and (stDate <> '') and (stSouche <> '') and (stNumero <> '') and
   (stIndiceg <> '') then
    begin
    StringToCleDoc (
        stNature + ';' + stDate + ';' + stSouche + ';' + stNumero + ';' + stIndiceg + ';',
        CleDoc);
    SaisiePiece (CleDoc, taConsult) ;
    end;
end;

procedure TFConsultMouv.BOuvrirClick(Sender: TObject);
var CleDoc : R_CleDoc;
    stNature, stDate, stSouche, stNumero, stIndiceg : string;
begin
inherited;
stNature := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NATUREPIECEG');
stDate := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_DATEPIECE');
stSouche := Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_SOUCHE');
stNumero := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_NUMERO'));
stIndiceg := IntToStr(Tob(G_ConsultMouv.Objects [0, G_ConsultMouv.Row]).GetValue ('GL_INDICEG'));
StringToCleDoc (
    stNature + ';' + stDate + ';' + stSouche + ';' + stNumero + ';' + stIndiceg + ';',
    CleDoc);
SaisiePiece (CleDoc, taConsult) ;
end;

procedure TFConsultMouv.BParamListeClick(Sender: TObject);
begin
Q.Liste := 'GCMULCONSUMOUV';
EtudieColsListe;
AffecteGrid (G_ConsultMouv, taConsult);
ChargeMouvement;
HMTrad.ResizeGridColumns (G_ConsultMouv);
inherited;
end;

end.
