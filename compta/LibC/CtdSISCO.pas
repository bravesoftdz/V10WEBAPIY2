unit CtdSISCO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Ent1, HEnt1, HCtrls, HTB97, ExtCtrls, HPanel, UiUtil,uDbxDataSet ;

Procedure CtrlDivSISCO ;

type
  TFCtrlDiv = class(TForm)
    Panel1: THPanel;
    ToolbarButton976: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    ToolbarButton972: TToolbarButton97;
    ToolbarButton973: TToolbarButton97;
    ToolbarButton974: TToolbarButton97;
    CSouche: TToolbarButton97;
    CMdsGen: TToolbarButton97;
    CMdsAux: TToolbarButton97;
    Dock: TDock97;
    HPB: TToolWindow97;
    BFerme: TToolbarButton97;
    procedure ToolbarButton976Click(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure ToolbarButton972Click(Sender: TObject);
    procedure ToolbarButton973Click(Sender: TObject);
    procedure ToolbarButton974Click(Sender: TObject);
    procedure CMdsGenClick(Sender: TObject);
    procedure CMdsAuxClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CSoucheClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

Uses UtofMulParamGen, ImpCptSI, Choix ;

{$R *.DFM}

Procedure CtrlDivSISCO ;
var FFR: TFCtrlDiv ;
    PP       : THPanel ;
begin
FFR:=TFCtrlDiv.Create(Application) ;
PP:=FindInsidePanel ;
if PP=NIL then
   BEGIN
   try
    FFR.ShowModal ;
   finally
    FFR.Free ;
   end ;
   END else
   BEGIN
   InitInside(FFR,PP) ;
   FFR.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
end ;


procedure TFCtrlDiv.ToolbarButton976Click(Sender: TObject);
begin
CCLanceFiche_MULGeneraux('C;7112000');
end;

procedure TFCtrlDiv.ToolbarButton971Click(Sender: TObject);
begin
CPLanceFiche_MULTiers('C;7145000');
end;

procedure TFCtrlDiv.ToolbarButton972Click(Sender: TObject);
begin
CPLanceFiche_MULSection('C;7178000');
end;

procedure TFCtrlDiv.ToolbarButton973Click(Sender: TObject);
begin
CPLanceFiche_MULJournal('C;7205000');
end;

procedure TFCtrlDiv.ToolbarButton974Click(Sender: TObject);
begin
MAJCptBqe ;
end;

Procedure MAJDECHAMPS(Ajoute : Boolean) ;
Var St,St1 : String ;
    Q : TQuery ;
BEGIN
St:='Select DH_CONTROLE FROM DECHAMPS WHERE DH_PREFIXE="G" AND DH_NOMCHAMP="G_VENTILABLE1"' ;
Q:=OpenSQL(St,FALSE) ;
If Not Q.Eof Then
  BEGIN
  St1:=Q.FindField('DH_CONTROLE').AsString ;
  If Ajoute Then St1:=St1+'Z' Else St1:=Copy(St,1,Length(St1)-1) ;
  Q.Edit ; Q.FindField('DH_CONTROLE').AsString:=St1 ; Q.Post ;
  END ;
Ferme(Q) ;
END ;

procedure TFCtrlDiv.CMdsGenClick(Sender: TObject);
begin
MajDECHAMPS(TRUE) ;
CCLanceFiche_MULGeneraux('SERIE;7115000');
ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE="X" WHERE G_VENTILABLE1="X"') ;
ExecuteSQL('UPDATE GENERAUX SET G_VENTILABLE="-" WHERE G_VENTILABLE1="-"') ;
MajDECHAMPS(FALSE) ;
end;

procedure TFCtrlDiv.CMdsAuxClick(Sender: TObject);
begin
CPLanceFiche_MULTiers('SERIE;7148000');
end;

procedure TFCtrlDiv.FormShow(Sender: TObject);
Var ExisteMvt : Boolean ;
begin
ExisteMvt:=ExisteSQL('SELECT E_GENERAL FROM ECRITURE') ;
CSouche.Enabled:=Not ExisteMvt ;
CMdsGen.Enabled:=Not ExisteMvt ;
CMdsAux.Enabled:=Not ExisteMvt ;
end;

procedure TFCtrlDiv.CSoucheClick(Sender: TObject);
Var Q : TQuery ;
    Choisir : Boolean ;
    LaSouche : String ;
begin
LaSouche:=Choix.Choisir('Choix d''une souche','SOUCHE','SH_LIBELLE','SH_SOUCHE','SH_ANALYTIQUE="-" AND SH_SIMULATION="-"','SH_SOUCHE') ;
If LaSouche='' Then Exit ;
Q:=OpenSQL('SELECT J_COMPTEURNORMAL FROM JOURNAL WHERE J_COMPTEURNORMAL=""',FALSE) ;
While Not Q.Eof Do
  BEGIN
  Q.Edit ; Q.FindField('J_COMPTEURNORMAL').AsString:=LaSouche ;
  Q.Next ;
  END ;
Ferme(Q) ;
end;

procedure TFCtrlDiv.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

end.
