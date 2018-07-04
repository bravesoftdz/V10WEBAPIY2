unit FmtLiens;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Hqry, StdCtrls, Hctrls, Buttons, ExtCtrls, Ent1, Hent1;

function ChoisirLiensParam (Prefixe : String3 ; var Liens : String) : String ;

type
  TFLiensPar = class(TForm)
    PBouton: TPanel;
    BValider: TBitBtn;
    HelpBtn: TBitBtn;
    BFerme: TBitBtn;
    HLabel3: THLabel;
    FTable: TEdit;
    FListe: TListBox;
    Q: THQuery;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    Liens : String ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

function ChoisirLiensParam (Prefixe : String3 ; var Liens : String) : String ;
var F : TFLiensPar ;
begin
Result:='' ;
F:=TFLiensPar.Create(Application) ;
try
  F.FTable.Text:=PrefixeToTable(Prefixe) ;
  F.Liens:=Liens ;
  F.ShowModal ;
  finally
  Result:=F.Liens ;
  Liens:=F.Liens ;
  F.Free ;
  end ;
SourisNormale ;
end ;

procedure TFLiensPar.FormShow(Sender: TObject);
var i :  Integer ;
    CodeG : String ;
begin
Q.Params[0].AsString:=FTable.Text ;
Q.Open ;
While Not Q.EOF do
   BEGIN
   FListe.Items.Add(Q.Fields[0].AsString) ;
   Q.Next ;
   END ;
CodeG:=Liens ;
While (CodeG<>'') do
    BEGIN
    i:=FListe.Items.Indexof(ReadTokenSt(CodeG)) ;
    if i>-1 then FListe.Selected[i]:=True ;
    END ;
end;

procedure TFLiensPar.BValiderClick(Sender: TObject);
var i    : integer ;
    Code : String ;
begin
if FListe.MultiSelect then
  BEGIN
  Code:='' ;
  For i:=0 to FListe.Items.Count-1 do
    if FListe.Selected[i] then Code:=Code+FListe.Items[i]+';' ;
  END else Code:=Fliste.Items[Fliste.ItemIndex] ;
Liens:=Code ;
ModalResult:=mrOk ;
end;

end.
