unit ZParams;

interface

uses Windows, Messages, SysUtils, Classes, Graphics,         { +- Delphi  -+ }
     Controls, Forms, Dialogs, StdCtrls, Buttons, ExtCtrls,  { +- Delphi  -+ }
     ComCtrls, Grids,                                        { +- Delphi  -+ }
     HEnt1, Hctrls,                                          { +- Bibli   -+ }
     TZ, ZTypes ;                                            { +- PFUGIER -+ }

function SaisieZParams(var Params : RPar; Ecrs: TZF) : Boolean ;

type
  TZFParams = class(TForm)
    panel: TPanel;
    BValider: TBitBtn;
    BFerme: TBitBtn;
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    FAnalytique: TCheckBox;
    FEcheance: TCheckBox;
    FLibre: TCheckBox;
    GroupBox2: TGroupBox;
    FPiece: TCheckBox;
    FSoldeProg: TCheckBox;
    FBackup: TCheckBox;
    FBlackout: TCheckBox;
    FTime: TCheckBox;
    FOptim: TCheckBox;
    Panel2: TPanel;
    TV: TTreeView;
    Panel3: TPanel;
    GT: THGrid;
    Splitter1: TSplitter;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
  public
    FParams : RPar ;
    FEcrs   : TZF ;
  end;

implementation

{$R *.DFM}

uses ZFolioU ;

//=======================================================
//======== Point d'entrée dans la saisie Params =========
//=======================================================
function SaisieZParams(var Params : RPar; Ecrs: TZF) : Boolean ;
var ZFParams : TZFParams ;
begin
ZFParams:=TZFParams.Create(Application) ;
  try
  ZFParams.FParams:=Params ;
  ZFParams.FEcrs:=Ecrs ;
  Result:=(ZFParams.ShowModal=mrOK) ;
  if Result then Params:=ZFParams.FParams ;
  finally
  ZFParams.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TZFParams.FormShow(Sender: TObject);
var i, j, k, c : Integer ; Ecr, Ana : TZF ; bFirst : Boolean ;
    Folio: TFolioU ;
begin
// TreeView
//TV.Items.Add(nil, 'Folio') ;
//if FEcrs.Detail.Count>0 then
//  for i:=0 to FEcrs.Detail.Count-1 do
//    FEcrs.Detail[i].PutTreeView(TV, TV.Items[0], 'E_DATECOMPTABLE|" "|E_GENERAL|" "|E_DEBIT|" "|E_CREDIT') ;
if FEcrs.Detail.Count>0 then
  begin
  Folio:=TFolioU.Create( FEcrs.Detail[0].GetValue('E_NUMEROPIECE'),
                         FEcrs.Detail[0].GetValue('E_JOURNAL'),
                         FEcrs.Detail[0].GetValue('E_EXERCICE'),
                         FindeMois(FEcrs.Detail[0].GetValue('E_DATECOMPTABLE')),
                         FEcrs.Detail[0].GetValue('E_MODESAISIE')='LIB') ;
  Folio.Read ;
//  if Folio.VerifFolio then Folio.Write ;
  Folio.VerifFolio ;
  Folio.Ecrs.PutGridDetail(GT, TRUE, TRUE, 'E_NUMGROUPEECR;E_DATECOMPTABLE;E_GENERAL;E_AUXILIAIRE;E_DEBIT;E_CREDIT;E_DEBITEURO;E_CREDITEURO;E_DEBITDEV;E_CREDITDEV') ;
  Folio.Write ;
  Folio.Free ;
  end ;
Exit ;  
// THGrid
bFirst:=TRUE ;
for i:=0 to FEcrs.Detail.Count-1 do
  begin
  Ecr:=TZF(FEcrs.Detail[i]) ; if Ecr.GetValue('BADROW')='X' then Continue ;
  for j:=0 to 4 do
    begin
//    Ecr.Detail[j].PutGridDetail(GT, TRUE, TRUE, '') ;
    for c:=0 to TZF(Ecr.Detail[j]).Detail.Count-1 do
      begin
      if bFirst then
        begin Ecr.Detail[j].Detail[c].PutLigneGrid(GT, 0, TRUE, TRUE, '') ; bFirst:=FALSE ; end ;
      Ecr.Detail[j].Detail[c].PutLigneGrid(GT, GT.RowCount, FALSE, FALSE, '') ;
      GT.RowCount:=GT.RowCount+1 ;
      end ;
    (*
    for k:=0 to c-1 do
      begin
      Ana:=TZF(TZF(Ecr.Detail[j]).Detail[k]) ;
      Ana.PutGridDetail(GT, TRUE, TRUE, '') ;
      end ;
    *)
    end ;
  end ;
//FEcrs.PutGridDetail(GT, TRUE, TRUE, '') ;
Exit ;
FAnalytique.Checked:=FParams.bAnalytique ;
FEcheance.Checked:=FParams.bEcheance ;
FPiece.Checked:=FParams.bPiece ;
FLibre.Checked:=FParams.bLibre ;
FSoldeProg.Checked:=FParams.bSoldeProg ;
FBackup.Checked:=FParams.bBackup ;
FBlackout.Checked:=False ; 
FTime.Checked:=FParams.bTime;
FOptim.Checked:=FParams.bOptim;
end;

procedure TZFParams.BValiderClick(Sender: TObject);
begin
ModalResult:=mrOk ;
Exit ;
FParams.bAnalytique:=FAnalytique.Checked ;
FParams.bEcheance:=FEcheance.Checked ;
FParams.bPiece:=FPiece.Checked ;
FParams.bSoldeProg:=FSoldeProg.Checked ;
FParams.bBackup:=FBackup.Checked ;
FParams.bTime:=FTime.Checked ;
FParams.bOptim:=FOptim.Checked ;
ModalResult:=mrOk ;
end;

procedure TZFParams.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
ModalResult:=mrOk ;
Exit ;
FParams.bAnalytique:=FAnalytique.Checked ;
FParams.bEcheance:=FEcheance.Checked ;
FParams.bPiece:=FPiece.Checked ;
FParams.bLibre:=FLibre.Checked ;
FParams.bSoldeProg:=FSoldeProg.Checked ;
FParams.bBackup:=FBackup.Checked ;
FParams.bTime:=FTime.Checked ;
FParams.bOptim:=FOptim.Checked ;
ModalResult:=mrOk ;
end;

end.
