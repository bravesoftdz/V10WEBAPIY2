unit ImVisAna;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  ComCtrls,
  Grids,
  Hctrls,
  StdCtrls,
  ExtCtrls,
  HTB97,
  ImEnt,
  HEnt1,
  Db,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  SaisUtil,
  HSysMenu,
  hmsgbox,
  CPSECTION_TOM,
  {$IFDEF EAGLCLIENT}
  utileagl,
  {$ELSE}
  PrintDBG,
  {$ENDIF}
  ImGenEcr, TntGrids, TntStdCtrls;

const COL_LIG = 0;
      COL_SEC = 1;
      COL_LIB = 2;
      COL_TAU = 3;
      COL_MON = 4;
type
  TFImVisAna = class(TForm)
    PFEN: TPanel;
    PEntete: TPanel;
    H_GENERAL: THLabel;
    G_GENERAL: THLabel;
    H_MONTANTECR: THLabel;
    E_MONTANT: THNumEdit;
    PGA: TPanel;
    GridAna: THGrid;
    TA: TTabControl;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BZoom: TToolbarButton97;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure TAChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BZoomClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Déclarations privées }
    fCompte        : string;
    fAxes          : Array[0..ImMaxAxe-1] of TList;
    fMontant : double;
    procedure DefautEntete ;
    procedure AfficheGrid;
  public
    { Déclarations publiques }
  end;

procedure ImVisuAnalytique( sCompte : string; Axes : Array of TList;MontantD,MontantC : double) ;


implementation

uses Outils;

{$R *.DFM}

procedure ImVisuAnalytique( sCompte : string; Axes : Array of TList;MontantD,MontantC : double) ;
Var X : TFImVisAna ;
    i :integer;
begin
  X:=TFImVisAna.Create(Application) ;
  try
    X.fCompte:=sCompte ;
    if MontantD = 0.0 then X.fMontant := (-1)*MontantC else
    X.fMontant := MontantD;
    for i:=0 to ImMaxAxe - 1 do X.fAxes[i] := Axes[i];
    X.Showmodal ;
  finally
    X.Free ;
  end;
  Screen.Cursor:=SyncrDefault ;
end ;

procedure TFImVisAna.DefautEntete ;
BEGIN
  G_GENERAL.Caption := fCompte;
  E_MONTANT.Value := abs(fMontant);
  E_MONTANT.Debit := (fMontant <> 0.0);
  ChangeFormatDevise(Self,2,'') ;
  // E_MONTANTD.Debit := True;
END ;

procedure TFImVisAna.FormShow(Sender: TObject);
var i : integer;
begin
  DefautEntete ;
  Ta.Tabs.Clear;
  for i:=0 to ImMaxAxe-1 do
  begin
    if fAxes[i] <> nil then
      Ta.Tabs.Add(HM.Mess[i]);
  end;

  GridAna.DefaultRowHeight:=16 ; GridAna.ListeParam:='SASAISIE1' ;
  GridAna.ColAligns[1] := taLeftJustify;
  GridAna.ColAligns[2] := taLeftJustify;
  GridAna.ColAligns[3] := taCenter;
  GridAna.ColAligns[4] := taRightJustify;
  AfficheGrid;
end;

procedure TFImVisAna.TAChange(Sender: TObject);
begin
  AfficheGrid;
end;

procedure TFImVisAna.AfficheGrid;
var L : TList;
    ARecord : TAna;
    i : integer;
    TotalVentil : double;
    stTitreTab : string;
begin
  TotalVentil := 0.0;
  GridAna.VidePile (True);
  stTitreTab := Ta.Tabs[Ta.TabIndex];
  L := fAxes[StrToInt(stTitreTab[Length(stTitreTab)])-1];
  if (L=nil) or (L.Count = 0) then GridAna.RowCount := 2 else
  begin
    GridAna.RowCount := L.Count + 1;
    for i := 0 to L.Count - 1 do
      TotalVentil := TotalVentil + TAna(L.Items[i]).Montant;
    for i := 0 to L.Count - 1 do
    begin
      ARecord := L.Items[i];
      GridAna.Cells[0,i+1] := IntToStr(i+1);
      GridAna.Cells[1,i+1] := ARecord.Section;
      GridAna.Cells[2,i+1] := ARecord.Libelle;
      GridAna.Cells[3,i+1] := StrFMontant(ARecord.TauxMontant,15,4,'',TRUE) ;
      GridAna.Cells[4,i+1] := MontantToStr(ARecord.Montant);
    end;
  end;
end;

procedure TFImVisAna.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  GridAna.VidePile (True);
end;

procedure TFImVisAna.BZoomClick(Sender: TObject);
var A  : TActionFiche ;
begin
//EPZ 15/11/00  if Not JaiLeDroitCpta(ccSecModif,False) then A:=taConsult else A:=taModif ;
  if Not ExJaiLeDroitConcept(TConcept(ImccSecModif),False) then A:=taConsult else A:=taModif ;
  FicheSection(Nil,'A'+Copy(TA.Tabs[TA.TabIndex],Length(TA.Tabs[TA.TabIndex]),1),GridAna.Cells[COL_SEC,GridAna.Row],A,0)
end;

procedure TFImVisAna.BImprimerClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  PrintDBGrid (caption,GridAna.name,'');
{$ELSE}
  PrintDBGrid (GridAna,Nil, Caption,'');
{$ENDIF}
end;

procedure TFImVisAna.BFermeClick(Sender: TObject);
begin
  Close;
end;

procedure TFImVisAna.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
