unit DetailOp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, ExtCtrls, HPanel,dbtables,majtable,StdCtrls, Hctrls, HTB97,
  hmsgbox, HSysMenu,HEnt1, UiUtil;

procedure  AfficheDetailOperation (CodeImmo :string);

const NB_COL_CESSION=5;

type
  TFDetailOp = class(TForm)
    FListeOpe: THPanel;
    ListeOpe: TListView;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    CodeImmo : string;
  public
    { Déclarations publiques }
  end;


implementation

uses Outils;

{$R *.DFM}

procedure  AfficheDetailOperation (CodeImmo :string);
var
  FDetailOp: TFDetailOp;
begin
  FDetailOp := TFDetailOp.Create(Application);
  FDetailOp.CodeImmo := CodeImmo;
  try
  FDetailOp.ShowModal;
  finally
  FDetailOp.Free;
  end;
end;


procedure TFDetailOp.FormCreate(Sender: TObject);
var  Colonne: TListColumn;
     i :integer;
begin
  for i := 0 to NB_COL_CESSION - 1 do
  begin
    Colonne := ListeOpe.Columns.Add;
    ListeOpe.Columns[i].Width := 95;
    Colonne.Caption := HM.Mess[i];
  end;
end;

procedure TFDetailOp.FormShow(Sender: TObject);
var Q : TQuery; Ligne : TListItem; LeSql: string ;
begin
  LeSql:='SELECT IL_TYPEOP,IL_MOTIFCES,IL_DATEOP,IL_VOCEDE,IL_MONTANTCES '
        +' FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+'" AND IL_TYPEOP LIKE "CE%"' ;
  Q := OpenSQL (LeSql,TRUE);
  while not Q.EOF do
    begin
    Ligne := ListeOpe.Items.Add;
    Ligne.Caption := RechDom('TIMOTIFCESSION',Q.FindField('IL_MOTIFCES').AsString,false);
    Ligne.SubItems.Add(DateToStr(Q.FindField('IL_DATEOP').AsDateTime));
    Ligne.SubItems.Add(MontantToStr(Q.FindField('IL_VOCEDEE').AsFloat));
    Ligne.SubItems.Add(MontantToStr(Q.FindField('IL_MONTANTCES').AsFloat));
    Ligne.SubItems.Add('Non calculé');
    end;
  Ferme (Q);
end;

end.
