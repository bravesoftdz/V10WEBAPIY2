unit RecupPlantype;

interface

uses
Windows, SysUtils, Classes, Graphics, Forms, Controls, StdCtrls,
Buttons, ExtCtrls, Dialogs, UTOB, HTB97, HStatus, HMsgBox, ComCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
  Db,
 {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_main,  // AGLLanceFiche
{$ENDIF}

  Paramsoc,
  Grids,
  HCtrls;

type
  TRECUPPLAN = class(TForm)
    OKBtn: TButton;
    CancelBtn: TButton;
    Bevel1: TBevel;
    OpenDialog1: TOpenDialog;
    Label1: TLabel;
    EditFich: TEdit;
    ToolbarButton971: TToolbarButton97;
    ProgressBar1: TProgressBar;
    procedure ToolbarButton971Click(Sender: TObject);
    procedure OKBtnClick(Sender: TObject);
  private
    { Private declarations }
    SizeFich :integer;

  public
    { Public declarations }
  end;

var
  RECUPPLAN: TRECUPPLAN;
  Procedure RecupPlansisco;

implementation

{$R *.DFM}
Procedure RecupPlansisco;
begin
    RECUPPLAN := TRECUPPLAN.Create (Application);
    try
       RECUPPLAN.ShowModal;
    finally
      RECUPPLAN.Free;
    end;

end;


procedure TRECUPPLAN.ToolbarButton971Click(Sender: TObject);
var
  fs      : file of Byte;
begin
  if OpenDialog1.Execute then          { Affichage de la boîte de dialogue d'ouverture }
  begin
    AssignFile(fs, OpenDialog1.FileName);   { Fichier sélectionné dans la boîte de dialogue }
    Reset(fs);
    SizeFich := FileSize(fs);
    EditFich.Text := OpenDialog1.FileName;                       { Introduction de la chaîne dans un contrôle TEdit }
    CloseFile(fs);
 end;
end;

procedure TRECUPPLAN.OKBtnClick(Sender: TObject);
var
  F          : TextFile;
  S,S1       : string;
  FTob       : TOB;
  Texte      : string;
  sizegen    : integer;
  ii         : integer;
  QDos       : TQuery;
begin
Texte := 'Confirmez-vous la récupération';
if HShowMessage('0;Récupération du plan ;'+Texte+';Q;YN;N;N;','','')<>mrYes then exit ;
QDos := OpenSQL('SELECT * FROM GENERAUX',True);
if not QDos.Eof then
      ExecuteSQL('Delete from GENERAUX');
Ferme (QDos);

if  OpenDialog1.FileName <> '' then
begin
    sizegen := GetParamSoc ('SO_LGCPTEGEN');
    AssignFile(F, OpenDialog1.FileName);   { Fichier sélectionné dans la boîte de dialogue }
    Reset(F);
//    InitMove(SizeFich,'Chargement du plan');
    ProgressBar1.min := 0;
    ProgressBar1.max := 100;
    ii := SizeFich div 100;
    ProgressBar1.step := 1;

    FTob:=TOB.Create('GENERAUX', Nil,-1) ;
    While Not EOF(F) do
    begin
        Readln(F, S);                          { Lecture de la première ligne du fichier }
        if (S[1] = 'C') and (S[2] <> '}') then
        begin
            S1 :=  copy (S, 2, sizegen);
            FTob.PutValue('G_GENERAL', S1);
            S1 := copy (S, 46, 25);
            FTob.PutValue('G_LIBELLE', S1);
            S1 := copy (S, 35, 1);
            if S[2] = '2' then
            FTob.PutValue('G_NATUREGENE', 'IMO')
            else
            if S[2] = '6' then
            FTob.PutValue('G_NATUREGENE', 'CHA')
            else
            if S[2] = '7' then
            FTob.PutValue('G_NATUREGENE', 'PRO')
            else
            if S1 = 'F' then
                    FTob.PutValue('G_NATUREGENE', 'COF')
            else if S1 = 'C' then
                    FTob.PutValue('G_NATUREGENE', 'COC')
            else FTob.PutValue('G_NATUREGENE', 'DIV');

            FTob.InsertOrUpdateDB(true);
        end;
//        MoveCur (False);
          ProgressBar1.StepIt;
    end;
    FTob.Free;
    CloseFile(F);
end;
//  FiniMove;
end;

end.
