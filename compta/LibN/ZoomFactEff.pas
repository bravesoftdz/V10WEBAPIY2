unit ZoomFactEff;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, StdCtrls, Hcompte, Mask, Hctrls, hmsgbox, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry,
  Grids, DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, Saisie, SaisUtil,
  SaisComm, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  EncUtil, UTOB, FE_MAIN, UtilPGI, Choix, HStatus
  ,SaisBor, HRichOLE, ADODB
  ;

Procedure ZoomFactCliEff(LOBM : TList) ;

type
  TFZoomfactEff = class(TFMul)
    XX_WHERE: TEdit;
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    LOBM : TList ;
    procedure FaitXX_WHERE ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure ZoomFactCliEff(LOBM : TList) ;
var X : TFZoomfactEff ;
    PP : THPanel ;
begin
PP:=FindInsidePanel ;
X:=TFZoomfactEff.Create(Application) ;
X.FNomFiltre:='' ; X.Q.Liste:='CPMULFACTCLIEFFT' ;
X.LOBM:=LOBM ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

procedure TFZoomfactEff.FaitXX_WHERE ;
Var St,St1 : String ;
    O : TOBM ;
    i : Integer ;
BEGIN
St:='' ;
For i:=0 To LOBM.Count-1 Do
  BEGIN
  O:=TOBM(LOBM[i]) ;
  St1:='(E_JOURNAL="'+O.GetMvt('E_JOURNAL')+'"  AND E_EXERCICE="'+QuelExo(DateToStr(O.GetMvt('E_DATECOMPTABLE')))+'" AND '+
       ' E_DATECOMPTABLE="'+UsDateTime(O.GetMvt('E_DATECOMPTABLE'))+'" AND '+
       ' E_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE'))+' AND '+
       ' E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE'))+') ' ;
  If St='' Then St:=St1 Else St:=St+' OR '+St1 ;
  END ;
If St='' Then Exit ;
St:='('+St+')' ;
XX_WHERE.Text:=St ;
END ;

procedure TFZoomfactEff.FormShow(Sender: TObject);
begin
FaitXX_WHERE ;
  inherited;
UpdateCaption(Self) ;
BAgrandirClick(Nil) ;
BAgrandir.Visible:=FALSE ; BReduire.Visible:=FALSE ;
end;


end.
