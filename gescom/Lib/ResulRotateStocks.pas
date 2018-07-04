unit ResulRotateStocks;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  TeeProcs, TeEngine, Chart, HEnt1, EntGC, HTB97, Grids, Hctrls, Spin, StdCtrls, Mask,
  ExtCtrls, HPanel, DBGrids, HDB, UTob, Mul, DBTables, UtilArticle, AGLInitGC, Series, Math,
  HSysMenu;

type
  TResulRotateStock = class(TForm)
    HPanel1: THPanel;
    _RG_QTEDEP: THRadioGroup;
    _CB_STKNUL: TCheckBox;
    T_SE_CONSO: THLabel;
    _SE_CONSO: TSpinEdit;
    G_Resul: THGrid;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    bPrint: TToolbarButton97;
    BAbandon: TToolbarButton97;
    bCourbes: TToolbarButton97;
    FListe: THDBGrid;
    HMTrad: THSystemMenu;
    TW1: TToolWindow97;
    DG1: TChart;
    Series1: THorizBarSeries;
    bCalcul: TToolbarButton97;
    _SE_CALCUL: TSpinEdit;
    T_SE_CALCUL: THLabel;
    procedure FormShow(Sender: TObject);
    procedure bCourbesClick(Sender : TObject);
    procedure bCalculClick(Sender : TObject);
  private
    { Déclarations privées }
    procedure CalculRotate;
    procedure Calcul_Rotate_Article(TobBase : TOB);
    procedure Affiche_Resul(TobBase : TOB);
    procedure Affiche_Courbes(TobBase : TOB);
    procedure Calcul_Date_Fin_Stock(TobBase : TOB; BorneSup : integer);

  public
    { Déclarations publiques }
    NomList, Where : string;
  end;

var
  ResulRotateStock: TResulRotateStock;
  F : TFMul;
  FFieldList : string;
  AuMoinsUn : boolean;
  LesCols, FTitre : string;
  NbPeriode, PeriodeStock, PeriodeDemande, BaseStock, BaseConso : integer;
  Limite_Prec, PeriodeStock_Prec, PeriodeDemande_Prec, BaseStock_Prec, BaseConso_Prec : integer;
  TobBase : TOB;

implementation

{$R *.DFM}

procedure TResulRotateStock.FormShow(Sender: TObject);
var
   i_ind1, i_ind2, Dec : integer;
   St,stAl,stA,CH, FF, Perso : string ;
   FRecordSource,FLien,FSortBy,FLargeur,FAlignement,FParams,tt,NC : string ;
   Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul,OkTri,OkNumCol : boolean ;

begin
ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,Perso,OkTri,OkNumCol);
LesCols := FFieldList;
CH := FFieldList ;
StAl := FAlignement;
St := LesCols;
//G_Resul.RowCount := TobBase.Detail.Count + 1;
G_Resul.ColCount := 6;
G_Resul.Cells[0, 0] := 'Dépôt';
G_Resul.ColWidths[0] := 40;
G_Resul.Cells[1, 0] := 'Article';
G_Resul.ColWidths[1] := 100;
G_Resul.Cells[2, 0] := 'Libellé';
G_Resul.ColWidths[2] := 280;
G_Resul.Cells[3, 0] := 'Stock départ';
G_Resul.ColWidths[3] := 60;
G_Resul.ColAligns[3] := taRightJustify;
G_Resul.ColFormats[3] := '# ##0,000';
G_Resul.Cells[4, 0] := 'Conso. moy.';
G_Resul.ColWidths[4] := 60;
G_Resul.ColAligns[4] := taRightJustify;
G_Resul.ColFormats[4] := '# ##0,000';
G_Resul.Cells[5, 0] := 'Date épuisement';
G_Resul.ColWidths[5] := 60;
for i_ind1 := 0 to FListe.Columns.Count - 1 do
    begin
    StA:=ReadTokenSt(StAl);
    if ReadTokenSt(CH) = 'GQ_PHYSIQUE' then
        begin
        TransAlign(StA,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
        for i_ind2 := 3 to G_Resul.ColCount - 1 do
            begin
            G_Resul.ColAligns[i_ind2]:=FListe.Columns.Items[i_ind1].Field.Alignment;
            G_Resul.ColWidths[i_ind2]:=FListe.Columns.Items[i_ind1].Width;
            G_Resul.ColFormats[i_ind2]:=FF;
            end;
        end;
    end;
//G_Resul.Visible := False;
CalculRotate;
TobBase.PutGridDetail(G_Resul, False, False, LesCols, False);
if not _CB_STKNUL.Checked then
   begin
   for i_ind1 := G_Resul.RowCount - 1 downto 1 do
       if StrToFloat(G_Resul.Cells[3, i_ind1]) = 0.0 then
           G_Resul.DeleteRow(i_ind1);
   end;
G_Resul.Visible := True;
HMTrad.ResizeGridColumns(G_Resul);
Affiche_Courbes(TobBase);
end;

procedure TResulRotateStock.bCourbesClick(Sender : TObject);
begin
TW1.Visible := not TW1.Visible;
end;

procedure TResulRotateStock.bCalculClick(Sender : TObject);
var
   i_ind1 : integer;

begin
//G_Resul.Visible := False;
CalculRotate;
TobBase.PutGridDetail(G_Resul, False, False, LesCols, False);
if not _CB_STKNUL.Checked then
   begin
   for i_ind1 := G_Resul.RowCount - 1 downto 1 do
       if StrToFloat(G_Resul.Cells[3, i_ind1]) = 0.0 then
           G_Resul.DeleteRow(i_ind1);
   end;
G_Resul.Visible := True;
HMTrad.ResizeGridColumns(G_Resul);
Affiche_Courbes(TobBase);
end;

//
//===========================================================================
//
procedure TResulRotateStock.CalculRotate;
var
    QQ : TQuery;
    i_ind1, i_ind2 : integer;
    st1, st2 : string;
    TobTemp : TOB;

begin
{i_ind1 := StrToInt(_RG_PERIODE.Values[_RG_PERIODE.ItemIndex]);
//if PeriodeDemande <> i_ind1 then PeriodeDemande_Prec := PeriodeDemande;
PeriodeDemande_Prec := PeriodeDemande;
PeriodeDemande := StrToInt(_RG_PERIODE.Values[_RG_PERIODE.ItemIndex]); }
i_ind1 := StrToInt(_RG_QTEDEP.Values[_RG_QTEDEP.ItemIndex]);
if BaseStock <> i_ind1 then
    begin
//    G_Resul.Visible := False;
    BaseStock_Prec := BaseStock;
    end;
BaseStock := StrToInt(_RG_QTEDEP.Values[_RG_QTEDEP.ItemIndex]);
i_ind1 := _SE_CONSO.Value;
if BaseConso <> i_ind1 then
    begin
//    G_Resul.Visible := False;
    BaseConso_Prec := BaseConso;
    end;
BaseConso := _SE_CONSO.Value;
if VH_GC.GCPeriodeStock = 'SEM' then PeriodeStock := 1
else if VH_GC.GCPeriodeStock = 'QUI' then PeriodeStock := 2
else PeriodeStock := 4;
NbPeriode := PeriodeStock * _SE_CALCUL.Value;
LesCols := 'GQ_DEPOT;GA_CODEARTICLE;GA_LIBELLE;STOCK_BASE;CONSO_MOY;DATE_FIN;';
//for i_ind1 := 0 to Limite do LesCols := LesCols + 'PERIODE' + IntToStr(i_ind1) + ';';
//for i_ind1 := 5 to G_Resul.ColCount - 1 do G_Resul.Cols[i_ind1].Clear;

//if not G_Resul.Visible then
    begin
    G_Resul.VidePile(True);
    for i_ind1 := 0 to G_Resul.RowCount - 1 do G_Resul.Rows[i_ind1].Clear;
    if TobBase <> nil then TobBase.Free;
    TobBase := TOB.Create ('DISPO', Nil, -1) ;
    QQ := OpenSQL('SELECT * FROM GCROTATESTOCKS ' + Where + ' ORDER BY GQ_DEPOT,GA_CODEARTICLE', True) ;
    if QQ.EOF then
        begin
        Ferme(QQ);
        Exit;
        end;
    TobBase.LoadDetailDB('GCROTATESTOCKS', '', '', QQ, False);
    Ferme (QQ) ;
    AuMoinsUn := False;
    end;
for i_ind1 := TobBase.Detail.count - 1 downto 0 do
    begin
//    if not G_Resul.Visible then
        begin
        st1 := TobBase.Detail[i_ind1].GetValue('GQ_DEPOT');
        st2 := TobBase.Detail[i_ind1].GetValue('GA_CODEARTICLE');
        for i_ind2 := 0 to FListe.nbSelected - 1 do
            begin
            //FListe.DataSource.DataSet.GotoBookMark(pointer(FListe.SelectedRows.Items[i_ind2]));
            FListe.GotoLeBookMark(i_ind2);
            if (FListe.DataSource.DataSet.Fields[0].AsString = st1) and
               (FListe.DataSource.DataSet.Fields[1].AsString = st2) then Break;
            end;
        if i_ind2 = FListe.NbSelected then
            begin
            TobBase.Detail.Delete(i_ind1);
            Continue;
            end;
        end;
    TobTemp := TobBase.Detail[i_ind1];
    Calcul_Rotate_Article(TobTemp);
    end;
if AuMoinsUn then Affiche_Resul(TobBase);
end;

procedure TResulRotateStock.Calcul_Rotate_Article(TobBase : TOB);
var
   Select, Depot, Article, Dim1, Dim2, Dim3, Dim4, Dim5 : string;
   i_ind1 : integer;
   Q : TQuery;
   TobStk : TOB;
   StockDepart, Conso_Moyenne : double;

begin
//if not G_Resul.Visible then
    begin
    Depot := TobBase.GetValue('GQ_DEPOT');
    Article := TobBase.GetValue('GA_CODEARTICLE');
    Dim1 := TobBase.GetValue('GA_CODEDIM1');
    Dim2 := TobBase.GetValue('GA_CODEDIM2');
    Dim3 := TobBase.GetValue('GA_CODEDIM3');
    Dim4 := TobBase.GetValue('GA_CODEDIM4');
    Dim5 := TobBase.GetValue('GA_CODEDIM5');
    Select := 'Select * From DISPO Where GQ_DEPOT="' + Depot +
              '" and GQ_ARTICLE="' + CodeArticleUnique(Article, Dim1, Dim2, Dim3, Dim4, Dim5) + '" ' +
              'and GQ_CLOTURE="-"';
    Q := OpenSQL(Select, True) ;
    if Q.EOF then
        begin
        Ferme(Q);
        Exit;
        end;
    TobStk := TOB.Create('DISPO', nil, -1);
    TobStk.SelectDB('', Q, False);
    if BaseStock = 0 then
        StockDepart := TobStk.GetValue('GQ_PHYSIQUE')
        else
        StockDepart := TobStk.GetValue('GQ_PHYSIQUE') + TobStk.GetValue('GQ_RESERVEFOU');
    Ferme(Q);
    TobStk.Free;
    Select := 'Select * From DISPO Where GQ_DEPOT="' + Depot +
              '" and GQ_ARTICLE="' + CodeArticleUnique(Article, Dim1, Dim2, Dim3, Dim4, Dim5) + '" ' +
              'and GQ_CLOTURE="X" Order By GQ_DATECLOTURE Desc';
    Q := OpenSQL(Select, True) ;
    if Q.EOF then
        begin
        Ferme(Q);
        Exit;
        end;
    TobStk := TOB.Create('', nil, -1);
    TobStk.LoadDetailDB('DISPO', '', '', Q, False);
    Ferme(Q);
    //for i_ind1 := 0 to TobStk.Detail.Count - 1 do
    for i_ind1 := 0 to Min(BaseConso, TobStk.Detail.Count) - 1 do
        Conso_Moyenne := Conso_Moyenne + TobStk.Detail[i_ind1].GetValue('GQ_CUMULSORTIES');
    // Calcul autre : on ramene la conso moyenne à la semaine et on multiplie par
    // le nombre de semaines necessaire.
    Conso_Moyenne := (Conso_Moyenne / (BaseConso * 1.0834)) * (PeriodeStock * 1.0834);
    TobBase.AddChampSup('CONSO_MOY', False);
    TobBase.PutValue('CONSO_MOY', Conso_Moyenne);
    TobBase.AddChampSup('STOCK_BASE', False);
    TobBase.PutValue('STOCK_BASE', StockDepart);
    TobBase.AddChampSup('DATE_FIN', False);
{    end
    else
    begin
    Conso_Moyenne := TobBase.GetValue('CONSO_MOY');
    Conso_Moyenne := Conso_Moyenne * (PeriodeDemande / PeriodeDemande_Prec);
    TobBase.PutValue('CONSO_MOY', Conso_Moyenne); }
    end;
//for i_ind1 := 0 to Limite do TobBase.AddChampSup('PERIODE' + IntToStr(i_ind1), False);
Calcul_Date_Fin_Stock(TobBase, NbPeriode);
AuMoinsUn := True;
end;

procedure TResulRotateStock.Calcul_Date_Fin_Stock(TobBase : TOB; BorneSup : integer);
var
   i_ind1 : integer;
   StockDepart, Conso_Moyenne : double;

begin
Conso_Moyenne := TobBase.GetValue('CONSO_MOY');
StockDepart := TobBase.GetValue('STOCK_BASE');

if Conso_Moyenne <> 0 then
    begin
    Conso_Moyenne := Conso_Moyenne / 7;
    i_ind1 := Trunc(StockDepart / Conso_Moyenne);
    TobBase.PutValue('DATE_FIN', DateToStr(Date + i_ind1));
    end
    else
    begin
    TobBase.PutValue('DATE_FIN', DateToStr(Date + (BorneSup * 7)));
    end;
if StrToDate(TobBase.GetValue('DATE_FIN')) > (Date + (BorneSup * 7)) then
    TobBase.PutValue('DATE_FIN', DateToStr(Date + (BorneSup * 7)));
end;

procedure TResulRotateStock.Affiche_Resul(TobBase : TOB);
begin
;
end;

procedure TResulRotateStock.Affiche_Courbes(TobBase : TOB);
var
   i_ind1 : integer;
   TobLocale, TobTemp : TOB;
   Article : string;
   DateFin : TDateTime;

begin
TobLocale := TOB.Create('', nil, -1);
TobLocale.Dupliquer(TobBase, True, True);
DG1.ClipPoints := True;
DG1.Series[0].Clear;
DG1.BottomAxis.Maximum := Date;
for i_ind1 := 0 to TobLocale.Detail.Count - 1 do
    begin
    TobTemp := TobLocale.Detail[i_ind1];
    if (TobTemp.GetValue('STOCK_BASE') = 0.0) and (not _CB_STKNUL.Checked) then Continue;
    Article := TobTemp.GetValue('GA_CODEARTICLE');
    DateFin := StrToDate(TobTemp.GetValue('DATE_FIN'));
    if DG1.BottomAxis.Maximum < DateFin then DG1.BottomAxis.Maximum := DateFin;
    DG1.Series[0].Add(DateFin, Article, clTeeColor);
    DG1.Series[0].Identifier := Article;
    DG1.Series[0].Title := TobTemp.GetValue('GA_CODEARTICLE');
    end;
DG1.BottomAxis.Minimum := Date;
end;

end.
