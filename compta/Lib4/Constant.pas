unit Constant;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,

  HSysMenu,
  HMsgbox,
  StdCtrls,
  Buttons,
{$IFDEF EAGLCLIENT}
  eTablette,
{$ELSE}
  Db,
{$IFNDEF DBXPRESS}{$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},{$ELSE}uDbxDataSet,{$ENDIF}
  DBCtrls,
  DBGrids,
  HDB,
  Tablette,
{$ENDIF}
  Cumuls,
  ExtCtrls,
  Grids,
  Ent1,
  HEnt1,
  HCtrls,
  Hqry,
  HTB97,
  HPanel,
  UiUtil ;

procedure Constantes ;

type
  TFConstante = class(TFTablette)
    Msg: THMsgBox;
    BParamC: TToolbarButton97;
    procedure BParamCClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
  private { Déclarations privées }
    Function  SupprimeConstante : Boolean ;
  public  { Déclarations publiques }
    Function  Supprime : Boolean ; override ;
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

procedure Constantes ;
var X  : TFConstante ;
    PP : THPanel ;
begin
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
X:=TFConstante.Create(Application) ;
X.FQuoi:='TTCONSTANTE' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end ;

Function TFConstante.Supprime : Boolean ;
begin
Result:=SupprimeConstante ;
end ;

Function TFConstante.SupprimeConstante : Boolean ;
Var Q : TQuery ;
BEGIN
Result:=True ;
Q:=OpenSql('Select * from CUMULS Where CU_COMPTE1="'+TChoixCod.FindField(Fprefixe+'_CODE').AsString+'"',True) ;
if Not Q.Eof then
   BEGIN
   if Msg.Execute(0,caption,'')=mrYes then
      BEGIN
      ExecuteSql('Delete from CUMULS Where CU_COMPTE1="'+TChoixCod.FindField(Fprefixe+'_CODE').AsString+'"') ;
      END else Result:=False ;
   END ;
Ferme(Q) ; Screen.Cursor:=SyncrDefault ;
END ;


procedure TFConstante.BParamCClick(Sender: TObject);
Var Code,Lib : String ;
begin
//  inherited;
if ((TChoixCod.Eof) and (TChoixCod.Bof)) then Exit ;
Code:=TChoixCod.FindField(Fprefixe+'_CODE').AsString ;
Lib:=TChoixCod.FindField(Fprefixe+'_LIBELLE').AsString ;
if Code='' then Exit ;
ParametrageCumuls(True,Code,Lib) ;
end;

procedure TFConstante.BDeleteClick(Sender: TObject);
begin
if Not Supprime then Exit ;
  inherited;
end;

end.
