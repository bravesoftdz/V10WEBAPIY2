unit SisEdBalC;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QRS1, Mask, Hctrls, Db, HMemTb, ExtCtrls, DBTables, Hqry, Menus,
  HSysMenu, hmsgbox, HPdfviewer, Grids, DBGrids, HDB, HQuickrp, StdCtrls,
  HTB97, ComCtrls, HPanel, UiUtil, HEnt1, Ent1, ParamDat;

procedure LanceEditionSISS57BalC(TypeEdition : string) ;

type
  TFSisEdBalC = class(TFQRS1)
    G_GENERAL: THCritMaskEdit;
    TG_GENERAL: THLabel;
    TG_GENERAL2: THLabel;
    G_GENERAL_: THCritMaskEdit;
    HLabel2: THLabel;
    ETABLISSEMENT: THValComboBox;
    HLabel3: THLabel;
    DEVISE: THValComboBox;
    HLabel1: THLabel;
    EXO1: THValComboBox;
    TE_DATECOMPTABLE: THLabel;
    DATE1: THCritMaskEdit;
    TE_DATECOMPTABLE2: THLabel;
    DATE2: THCritMaskEdit;
    HLabel4: THLabel;
    EXO2: THValComboBox;
    DATE3: THCritMaskEdit;
    DATE4: THCritMaskEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure EXO1Change(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
  private
    { Déclarations privées }
    fTypeEdition : string;
    procedure FillComboBds ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure LanceEditionSISS57BalC(TypeEdition : string) ;
var X : TFSisEdBalC ; PP : THPanel ;
begin
  X := TFSisEdBalC.Create(Application) ;
  PP:=FindInsidePanel ;
  X.fTypeEdition := TypeEdition;
  X.InitQR('E', 'UCO', TypeEdition, TRUE, TRUE) ;
  if PP=Nil then
  begin
    try
      X.ShowModal ;
     finally
      X.Free ;
     end ;
   END else
   BEGIN
     InitInside(X,PP) ;
     X.Show ;
   end;
   Screen.Cursor:=SyncrDefault ;
end;

procedure TFSisEdBalC.FormClose(Sender: TObject; var Action: TCloseAction);
begin
inherited ;
if isInside(Self) then Action:=caFree ;
end;

procedure TFSisEdBalC.FormShow(Sender: TObject);
begin
bEuro.Visible:=FALSE ;
EXO1.Value:=VH^.Entree.Code ;
DATE1.Text:=DateToStr(VH^.Entree.Deb) ;
DATE2.Text:=DateToStr(VH^.Entree.Fin) ;
ETABLISSEMENT.Value:=VH^.EtablisDefaut ;
DEVISE.Value:=V_PGI.DevisePivot ;
FillComboBds ;
inherited ;
UpdateCaption(self);
end;

procedure TFSisEdBalC.FillComboBds ;
var Q : TQuery ; Cols, Where : string ;
begin
EXO2.Clear ; EXO2.Values.Clear ;
Cols:='HB_EXERCICE, HB_DATE1, HB_DATE2' ;
Where:='HB_TYPEBAL="BDS"' ;
Q:=OpenSQL('SELECT DISTINCT '+Cols+' FROM HISTOBAL WHERE '+Where, TRUE) ;
while not Q.EOF do
  begin
  EXO2.Values.Add(Q.Fields[0].AsString+';'+Q.Fields[1].AsString+';'+Q.Fields[2].AsString+';') ;
  EXO2.Items.Add(TraduireMemoire('Balance du '+DateToStr(Q.Fields[1].AsDateTime)+' au '+DateToStr(FinDeMois(Q.Fields[2].AsDateTime)))) ;
  Q.Next ;
  end ;
Ferme(Q) ;
EXO2.ItemIndex:=0 ;
end ;

procedure TFSisEdBalC.EXO1Change(Sender: TObject);
begin
inherited ;
ExoToDates(EXO1.Value, DATE1, DATE2) ;
end;

procedure TFSisEdBalC.BValiderClick(Sender: TObject);
var Par, RToken, sSave : string ;
begin
sSave:=EXO2.Value ;
RToken:=EXO2.Value ;
Par:=ReadTokenST(RToken) ;
EXO2.Values[EXO2.ItemIndex]:=Par ;
Par:=ReadTokenST(RToken) ;
DATE3.Text:=Par ;
Par:=ReadTokenST(RToken) ;
DATE4.Text:=Par ;
inherited ;
EXO2.Values[EXO2.ItemIndex]:=sSave ;
end;

procedure TFSisEdBalC.E_DATECOMPTABLEKeyPress(Sender: TObject;
  var Key: Char);
begin
inherited ;
ParamDate(Self,Sender,Key) ;
end;

end.
