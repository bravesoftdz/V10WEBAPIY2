unit TIDLot;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Mask, ExtCtrls, HTB97, HEnt1,
  UTOB, hmsgbox, HSysMenu, FactUtil;

Procedure EntreeTIDLot (var TOBOri, TOBDes : TOB; Infos : string);

type
  TFTIDLot = class(TForm)
    Dock971: TDock97;
    Toolbar972: TToolWindow97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    PANENTETE: TPanel;
    TGQ_ARTICLE: THLabel;
    TGQ_DEPOT: THLabel;
    GQ_ARTICLE: THCritMaskEdit;
    GQ_DEPOT: THCritMaskEdit;
    PANPIED: TPanel;
    TQTEAAFFECT: THLabel;
    TQTEAFFECT: THLabel;
    TQTERESTEAFFECT: THLabel;
    QTEAAFFECT: THCritMaskEdit;
    QTEAFFECT: THCritMaskEdit;
    QTERESTEAFFECT: THCritMaskEdit;
    G_Lig: THGrid;
    Titres: THMsgBox;
    HMTrad: THSystemMenu;
    BChercher: TToolbarButton97;
    FindLigne: TFindDialog;
    procedure FormShow(Sender: TObject);
    procedure G_LigRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_LigDblClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure G_LigCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_LigCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
  private
    { Déclarations privées }
    FindDebut : Boolean;
    procedure AffectePANPIED;
    procedure EtudieColsListe;
  public
    { Déclarations publiques }
    TOBLotOri, TOBLotDes : TOB;
    NUMEROSERIE : string;
  end;

var
  FTIDLot: TFTIDLot;
  LesCols : string;
  d_QteAAffect, d_QteAffect, d_QteResteAffect : double;
  Old_Cell : string;
  SG_NLig : integer;
  SG_Lot  : integer;
  SG_Seri : integer;
  SG_QOri : integer;
  SG_QDes : integer;
  ChDepo  : string;
  ChLot   : string;
  ChSerie : string;
  ChPhys  : string;
  ChTabNo : string;
  iNumTable : integer;

implementation

{$R *.DFM}

Procedure EntreeTIDLot (var TOBOri, TOBDes : TOB; Infos : string);
begin
FTIDLot := TFTIDLot.Create(Application);
FTIDLot.GQ_ARTICLE.Text := ReadTokenSt(Infos);
FTIDLot.GQ_DEPOT.Text := ReadTokenSt(Infos);
FTIDLot.QTEAAFFECT.Text := ReadTokenSt(Infos);
FTIDLot.NUMEROSERIE := ReadTokenSt(Infos);
FTIDLot.TOBLotOri := TOBOri;
FTIDLot.TOBLotDes := TOBDes;
FTIDLot.ShowModal;
end;

//=========================================================================
//  Préparation de la forme et des tobs passéees en paramètres
//=========================================================================
procedure TFTIDLot.FormShow(Sender: TObject);
var
    TOBTemp, TOBTemp2 : TOB;
    i_ind1, i_ind2 : integer;

begin
if TOBLotOri = nil then Close;
Old_Cell := '';
d_QteAAffect := 0.0;
d_QteAffect := 0.0;
{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
if NUMEROSERIE <> 'X' then
    begin
    LesCols := 'NUMERO;GQL_NUMEROLOT;GQL_PHYSIQUE;AFFECTE';
    ChDepo  := 'GQL_DEPOT';
    ChLot   := 'GQL_NUMEROLOT';
    ChSerie := '';
    ChPhys  := 'GQL_PHYSIQUE';
    ChTabNo := 'DISPOLOT';
    iNumTable := PrefixeToNum('GQL');
    end
    else
    begin
    LesCols := 'NUMERO;GQL_NUMEROSERIE;GQL_PHYSIQUE;AFFECTE';
    ChDepo  := 'GQL_DEPOT';
    ChLot   := '';
    ChSerie := 'GQL_NUMEROSERIE';
    ChPhys  := 'GQL_PHYSIQUE';
    ChTabNo := 'DISPOLOT';
    iNumTable := PrefixeToNum('GQL');
    end;
{$ENDIF}
EtudieColsListe;
G_Lig.Cells[SG_NLig, 0] := Titres.Mess[0];
if SG_Lot  <> -1 then G_Lig.Cells[SG_Lot , 0] := Titres.Mess[1];
if SG_Seri <> -1 then G_Lig.Cells[SG_Seri, 0] := Titres.Mess[2];
G_Lig.Cells[SG_QOri, 0] := Titres.Mess[3];
G_Lig.Cells[SG_QDes, 0] := Titres.Mess[4];
HMTrad.ResizeGridColumns (G_Lig) ;
G_Lig.ColLengths[SG_Lot] := -1;
G_Lig.ColLengths[SG_QOri] := -1;
//G_Lig.ColFormats[SG_QOri] := '#####0,000';
//G_Lig.ColFormats[SG_QDes] := '#####0,000';
G_Lig.ColAligns[SG_NLig] := taCenter;
G_Lig.ColAligns[SG_QOri] := taRightJustify;
G_Lig.ColAligns[SG_QDes] := taRightJustify;
G_Lig.RowCount := TOBLotOri.Detail.Count + 1;
TOBTemp := TOB.Create('', nil, -1);
TOBTemp.Dupliquer(TOBLotDes, True, True);
TOBLotDes.ClearDetail;
for i_ind1 := 0 to TOBLotOri.Detail.Count - 1 do
    begin
    TOBLotOri.Detail[i_ind1].AddChampSup('NUMERO', False);
    TOBLotOri.Detail[i_ind1].PutValue('NUMERO', i_ind1 + 1);
    TOBTemp2 := TOB.Create(ChTabNo, TOBLotDes, i_ind1);
    TOBTemp2.Dupliquer(TOBLotOri.Detail[i_ind1], True, True);
    end;
for i_ind1 := 0 to TOBLotOri.Detail.Count - 1 do
    begin
//    TOBLotDes.Detail[i_ind1].Dupliquer(TOBLotOri.Detail[i_ind1], True, True);
    TOBLotDes.Detail[i_ind1].PutValue(ChDepo, TOBLotDes.GetValue('GQ_DEPOT'));
    TOBLotDes.Detail[i_ind1].PutValue(ChPhys, 0.0);
    if TOBTemp.Detail.Count > 0 then
        begin
        TOBTemp2 := TOBTemp.FindFirst([ChLot],
                                      [TOBLotDes.Detail[i_ind1].GetValue(ChLot)],
                                      False);
        if TOBTemp2 <> nil then
            TOBLotDes.Detail[i_ind1].PutValue(ChPhys, TOBTemp2.GetValue(ChPhys));
        end;
    G_Lig.Cells[SG_NLig, i_ind1 + 1] := IntToStr(TOBLotOri.Detail[i_ind1].GetValue('NUMERO'));
    if SG_Lot  <> -1 then
        G_Lig.Cells[SG_Lot , i_ind1 + 1] := TOBLotOri.Detail[i_ind1].GetValue(ChLot);
    if SG_Seri <> -1 then
        G_Lig.Cells[SG_Seri, i_ind1 + 1] := TOBLotOri.Detail[i_ind1].GetValue(ChLot);
    G_Lig.Cells[SG_QOri, i_ind1 + 1] := FloatToStr(TOBLotOri.Detail[i_ind1].GetValue(ChPhys));
    G_Lig.Cells[SG_QDes, i_ind1 + 1] := FloatToStr(TOBLotDes.Detail[i_ind1].GetValue(ChPhys));
    end;
//TOBLotOri.PutGridDetail(G_Lig, False, False, LesCols);
G_Lig.Row := 1;
G_Lig.Col := 3;
d_QteAAffect := StrToFloat(QTEAAFFECT.Text);
d_QteResteAffect := d_QteAAffect;
AffectePANPIED;
TOBLotOri.PutValue('GQ_PHYSIQUE', d_QteAAffect);
TOBLotDes.PutValue('GQ_PHYSIQUE', d_QteAAffect);
end;

procedure TFTIDLot.G_LigCellEnter(Sender: TObject; var ACol, ARow: Integer;
  var Cancel: Boolean);
begin
Old_Cell := G_Lig.Cells[G_Lig.Col, G_Lig.Row];
end;

procedure TFTIDLot.G_LigCellExit(Sender: TObject; var ACol, ARow: Integer;
  var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;
if G_Lig.Col <> 3 then  Exit;
if Old_Cell = G_Lig.Cells[SG_QDes, ARow] then Exit;
if G_Lig.Cells[SG_QDes, ARow] = '' then Exit;
if StrToFloat(G_Lig.Cells[SG_QDes, ARow]) < 0.0 then G_Lig.Cells[SG_QDes, ARow] := '0.000';
if (NUMEROSERIE = 'X') and
   (StrToFloat(G_Lig.Cells[SG_QDes, ARow]) <> 1.0) then G_Lig.Cells[SG_QDes, ARow] := '1.000';
if StrToFloat(G_Lig.Cells[SG_QDes, ARow]) > TOBLotOri.Detail[ARow - 1].GetValue('QTEORIG') then
    G_Lig.Cells[SG_QDes, ARow] := FloatToStr(TOBLotOri.Detail[ARow - 1].GetValue('QTEORIG'));
end;

//=========================================================================
//  On verifie que la quantité saisie est > 0, <= à la qte dispo,
//  <= au reste à affecter
//=========================================================================
procedure TFTIDLot.G_LigRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;
if G_Lig.Cells[SG_QDes, Ou] = '' then Exit;
if Old_Cell = G_Lig.Cells[SG_QDes, Ou] then Exit;
if StrToFloat(G_Lig.Cells[SG_QDes, Ou]) <> 0.0 then
    begin
    TOBLotDes.Detail[Ou - 1].PutValue(ChPhys, StrToFloat(G_Lig.Cells[SG_QDes, Ou]));
    TOBLotOri.Detail[Ou - 1].PutValue(ChPhys, TOBLotOri.Detail[Ou - 1].GetValue('QTEORIG') -
                                                      StrToFloat(G_Lig.Cells[SG_QDes, Ou]));
    G_Lig.Cells[SG_QOri, Ou] := FloatToStr(TOBLotOri.Detail[Ou - 1].GetValue(ChPhys));
    end;
AffectePANPIED;
end;

//=========================================================================
//  double clic selectionne le lot entier; on verifie que la qte selectionnée
//  est <= au reste à affecter
//=========================================================================
procedure TFTIDLot.G_LigDblClick(Sender: TObject);
var
    i_ind1 : integer;

begin
i_ind1 := G_Lig.Row - 1;
if d_QteResteAffect = 0.0 then Exit;
if TOBLotOri.Detail[i_ind1].GetValue(ChPhys) <= d_QteResteAffect then
    TOBLotDes.Detail[i_ind1].PutValue(ChPhys, TOBLotDes.Detail[i_ind1].GetValue(ChPhys) +
                                                             TOBLotOri.Detail[i_ind1].GetValue(ChPhys))
    else
    TOBLotDes.Detail[i_ind1].PutValue(ChPhys, d_QteResteAffect);
TOBLotOri.Detail[i_ind1].PutValue(ChPhys, TOBLotOri.Detail[i_ind1].GetValue('QTEORIG') -
                                                         TOBLotDes.Detail[i_ind1].GetValue(ChPhys));
G_Lig.Cells[SG_QOri, G_Lig.Row] := FloatToStr(TOBLotOri.Detail[i_ind1].GetValue(ChPhys));
G_Lig.Cells[SG_QDes, G_Lig.Row] := FloatToStr(TOBLotDes.Detail[i_ind1].GetValue(ChPhys));
Old_Cell := G_Lig.Cells[SG_QDes, G_Lig.Row];
AffectePANPIED;
end;

//=========================================================================
//  Abandon : on purge la tobdes avant le retour à l'appelant
//=========================================================================
procedure TFTIDLot.BAbandonClick(Sender: TObject);
begin
TOBLotDes.ClearDetail;
Close;
end;

procedure TFTIDLot.BValiderClick(Sender: TObject);
begin
TOBLotOri.PutValue('GQ_PHYSIQUE', d_QteAffect);
TOBLotDes.PutValue('GQ_PHYSIQUE', d_QteAffect);
Close;
end;

procedure TFTIDLot.BChercherClick(Sender: TObject);
begin
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TFTIDLot.FindLigneFind(Sender: TObject);
begin
Rechercher (G_Lig, FindLigne, FindDebut);
end;

procedure TFTIDLot.AffectePANPIED;
var
    i_ind1 : integer;
begin
d_QteAffect := 0.0;
for i_ind1 := 0 to TOBLotDes.Detail.Count - 1 do
    d_QteAffect := d_QteAffect + TOBLotDes.Detail[i_ind1].GetValue(ChPhys);
d_QteResteAffect := d_QteAAffect - d_QteAffect;
QTEAFFECT.Text := FloatToStr(d_QteAffect);
QTERESTEAFFECT.Text := FloatToStr(d_QteResteAffect);
end;

procedure TFTIDLot.EtudieColsListe;
Var NomCol,LesColsLoc : String ;
    icol,ichamp : integer ;
BEGIN
SG_Lot  := -1;
SG_Seri := -1;
LesColsLoc := LesCols;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    BEGIN
    ichamp:=ChampToNum(NomCol) ;
    if ichamp>=0 then
       BEGIN
       if NomCol=ChLot     then SG_Lot :=icol else
       if NomCol=ChSerie   then SG_Seri:=icol else
       if NomCol=ChPhys    then SG_QOri:=icol else
       end
       else
       begin
       if NomCol='NUMERO'  then SG_NLig:=icol else
       if NomCol='AFFECTE' then SG_QDes:=icol;
       END ;
    END ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
G_Lig.ColCount := icol;
end;

end.
