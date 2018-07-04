unit OpenSaveDlg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids,
{$IFNDEF EAGLCLIENT}
  DBGrids, HDB, ExtCtrls,
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

  Hqry, HTB97,
  HPanel,
  StdCtrls, Hctrls, Mask, Buttons, TntDBGrids, TntButtons, TntExtCtrls,
  TntStdCtrls;

type
  TDelProc = procedure(Q : TQuery; Var CanDelete : boolean) of object;

  TOpenSaveDialog = class(TForm)
    TCode: THLabel;
    TLib: THLabel;
    ed_code: THCritMaskEdit;
    ed_lib: THCritMaskEdit;
    G_Lst: THDBGrid;
    HPanel1: THPanel;
    SQ: TDataSource;
    BDel: THBitBtn;
    BValid: TBitBtn;
    BAnnuler: THBitBtn;
    BAide: THBitBtn;

    procedure SQDataChange(Sender: TObject; Field: TField);
    procedure BDelClick(Sender: TObject);
    procedure BValidClick(Sender: TObject);
    procedure G_LstDblClick(Sender: TObject);
    function Check4Suffix : boolean;


  private
    Q : TQuery;
    FCCode, FCLib : String;
    FSuffix : String;
    FDelProc : TDelProc;

  public
    CodeRetour, LibRetour : String;
    destructor Destroy; override;
    procedure Init(Titre, NomTable, ChampCode, ChampLib, Where, SufX : String; DeleteProc : TDelProc; PeutSaisir, PeutEffacer : boolean);

  end;


function PGIOpenSave(Titre, NomTable, ChampCode, ChampLib, Where, SufX : String; Var RetCode, RetLib : String; DeleteProc : TDelProc; PeutSaisir : boolean = true; PeutEffacer : boolean = true) : boolean;


implementation
uses HEnt1, HMsgBox
   ,CbpMCD
   ,CbpEnumerator
;

{$R *.DFM}

Const SfxError = 'Le code doit commencer par ';


function PGIOpenSave(Titre, NomTable, ChampCode, ChampLib, Where, SufX : String; Var RetCode, RetLib : String; DeleteProc : TDelProc; PeutSaisir : boolean = true; PeutEffacer : boolean = true) : boolean;
begin
with TOpenSaveDialog.Create(Application) do
  begin
  CodeRetour := RetCode; LibRetour := RetLib;
  Init(Titre, NomTable, ChampCode, ChampLib, Where, SufX, DeleteProc, PeutSaisir, PeutEffacer);
  ShowModal;
  result := (CodeRetour <> '') and (ModalResult = mrOk);
  RetCode := CodeRetour; RetLib := LibRetour;
  Release;
  end;
end;

Procedure TOpenSaveDialog.Init(Titre, NomTable, ChampCode, ChampLib, Where, SufX : String; DeleteProc : TDelProc; PeutSaisir, PeutEffacer : boolean);
var iChamp, iTableLigne : integer;
    s : string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
Caption := Titre;
FDelProc := DeleteProc;
FCCode := ChampCode;
FCLib := ChampLib;
FSuffix := SufX;

S := 'SELECT '+FCCode+', '+FCLib+' FROM '+NomTable;
if Where <> '' then S := S+' WHERE '+Where;
Q := OpenSQL(S, false);
SQ.DataSet := Q;

BDel.Visible := PeutEffacer;

ed_code.Enabled := PeutSaisir;
ed_lib.Enabled := PeutSaisir;
if PeutSaisir then
  begin
  ed_code.text := CodeRetour;
  ed_lib.text := LibRetour;
  end;

// Reglage de la taille des champs (nb de caracteres max)
table := Mcd.getTable(mcd.PrefixetoTable('LI'));
FieldList := Table.Fields;
FieldList.Reset();
While FieldList.MoveNext do
  begin
  s := (FieldList.Current as IFieldCOM).Tipe;
  Delete(s, 1, 8); Delete(S, Length(s), 1);
  if (FieldList.Current as IFieldCOM).name = FCCode then ed_code.MaxLength := StrToIntDef(s,17);
  if (FieldList.Current as IFieldCOM).name = FCLib then ed_lib.MaxLength := StrToIntDef(s,17);
  end;
end;

destructor TOpenSaveDialog.Destroy;
begin
SQ.DataSet := nil;
if Q <> nil then Ferme(Q);
inherited;
end;

procedure TOpenSaveDialog.SQDataChange(Sender: TObject; Field: TField);
begin
if Field <> nil then exit;
ed_code.text := Q.FindField(FCCode).AsString;
ed_lib.text := Q.FindField(FCLib).AsString;
end;

procedure TOpenSaveDialog.BDelClick(Sender: TObject);
var CanDelete : boolean;
begin
CanDelete := true;
if Assigned(FDelProc) then FDelProc(Q, CanDelete);
if CanDelete then Q.Delete;
end;

procedure TOpenSaveDialog.BValidClick(Sender: TObject);
begin
if not Check4Suffix then modalresult := mrNone;
CodeRetour := ed_code.text;
LibRetour := ed_lib.text;
end;

procedure TOpenSaveDialog.G_LstDblClick(Sender: TObject);
begin
BValid.Click;
end;

function TOpenSaveDialog.Check4Suffix : boolean;
var i : integer;
    s : string;
begin
result := true;
S := ed_code.text;
if Copy(S,1,Length(FSuffix)) <> FSuffix then
  begin
  PGIInfo(SfxError+FSuffix, Caption);
  result := false;
  for i := 1 to Length(FSuffix) do
   if i > Length(S) then S := S + FSuffix[i] else
   if S[i] <> FSuffix[i] then Insert(FSuffix[i], S, i);
  ed_code.text := S;
  ed_code.SetFocus;
  end;
end;

end.
