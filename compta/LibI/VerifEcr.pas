unit VerifEcr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, Grids, Hctrls, HTB97,UiUtil,HPanel,PrintDBG,dbTables,ImEnt;

procedure ExecuteVerificationEcarts ;
procedure CalculMontantFiches ( L : TList);
function CompareCompte (Item1,Item2:Pointer) : integer;

type
  TFVerifEcr = class(TForm)
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    GridEcart: THGrid;
    HMTrad: THSystemMenu;
    Msg: THMsgBox;
    BImprimer: TToolbarButton97;
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;


implementation

uses (*IntegEcr,*)Outils,ImGenEcr;

{$R *.DFM}

procedure ExecuteVerificationEcarts ;
var FVerifEcr: TFVerifEcr;
    PP:THPanel;
begin
  FVerifEcr:=TFVerifEcr.Create(Application) ;
  PP:=FindInsidePanel;
  if PP=nil then
  begin
    try
      FVerifEcr.ShowModal ;
    finally
      FVerifEcr.Free ;
    end ;
  end else
  begin
    InitInside(FVerifEcr,PP);
    FVerifEcr.Show;
  end;
end;


procedure TFVerifEcr.FormShow(Sender: TObject);
var ListeEcriture : TList;
    i : integer;
    ARecord : TLigneEcriture;
    ParamEcr : TParamEcr;
begin
  ParamEcr := TParamEcr.Create;
  GridEcart.ColAligns[0]:= taCenter;
  GridEcart.ColAligns[1]:= taRightJustify;
  GridEcart.ColAligns[2]:= taRightJustify;
  GridEcart.ColAligns[3]:= taRightJustify;
  ListeEcriture := nil;
  ListeEcriture := Tlist.Create;
  CalculMontantFiches (ListeEcriture);
  ParamEcr.DateCalcul := VHImmo^.Encours.fin;
  ParamEcr.Date := VHImmo^.Encours.Fin;
  CalculEcrituresDotation ( ListeEcriture , '',ParamEcr,0);
  ListeEcriture.Sort (CompareCompte);
  GridEcart.RowCount := ListeEcriture.Count + 1;
  for i := 1 to ListeEcriture.Count do
  begin
    ARecord := ListeEcriture.Items[i-1];
    GridEcart.Cells[0,i] := ARecord.Compte;
    GridEcart.Cells[1,i] := MontantToStr (ARecord.Debit - ARecord.Credit);
    GridEcart.Cells[2,i] := MontantToStr (0.0);
    GridEcart.Cells[3,i] := MontantToStr (StrToFloat(GridEcart.Cells[1,i])- StrToFLoat(GridEcart.Cells[2,i]));
  end;
  VideListeEcritures(ListeEcriture);
  ListeEcriture.Free ;
  ParamEcr.Free;
end;

procedure TFVerifEcr.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  GridEcart.VidePile(False);
  if Parent is THPanel then
  begin
    Action:=caFree ;
  end ;
end;

procedure TFVerifEcr.BImprimerClick(Sender: TObject);
begin
  PrintDBGrid (GridEcart,Nil, Caption,'');
end;

function CompareCompte (Item1,Item2:Pointer) : integer;
var Ecr1,Ecr2 : ^TLigneEcriture;
begin
  Ecr1 := Item1;Ecr2 := Item2;
  if Ecr1.Compte > Ecr2.Compte then Result := 1
  else if Ecr1.Compte < Ecr2.Compte then Result := -1
  else Result := 0;
end;

procedure CalculMontantFiches ( L : TList);
var Q : TQuery;
    ParamEcr : TParamEcr;
begin
  ParamEcr := TParamEcr.Create;
  ParamEcr.Date := VHImmo^.Encours.fin;
  ParamEcr.DateCalcul := VHImmo^.Encours.fin;
  Q := OpenSQL ('SELECT * FROM IMMO',TRUE);
  if not Q.Eof then
  begin
    Q.First;
    while not Q.Eof do
    begin
      MajLigneEcriture (L,0,'',ParamEcr,Q.FindField('I_COMPTEIMMO').AsString,
                        Q.FindField('I_MONTANTHT').AsFloat,True,0);
      Q.Next;
    end;
  end;
  Ferme ( Q );
  ParamEcr.Free;
end;

end.
