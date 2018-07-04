unit RegTVA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Tablette, HSysMenu, hmsgbox, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} StdCtrls, DBCtrls, Buttons,
  ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, HCtrls, Hqry, HTB97, HPanel, UiUtil;

procedure RegimesTVA ;

type
  TFRegimeTVA = class(TFTablette)
    Msg: THMsgBox;
    procedure TChoixCodBeforeDelete(DataSet: TDataSet);
  private { Déclarations privées }
  Function SupprimeRegimeTva : Boolean ;
  public  { Déclarations publiques }
  Function Supprime : Boolean ; override ;
  end;

implementation

{$R *.DFM}

procedure RegimesTVA ;
var X : TFRegimeTVA ;
    PP : THPanel ;
begin
if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
X:=TFRegimeTVA.Create(Application) ;
X.FQuoi:='TTREGIMETVA' ;
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

Function TFRegimeTVA.Supprime : Boolean ;
begin
Result:=SupprimeRegimeTva ;
end ;

Function TFRegimeTVA.SupprimeRegimeTva : Boolean ;
Var Q : TQuery ;
    St : String ;
    Trouver : Boolean ;
BEGIN
Result:=False ;
St:='' ;
Q:=OpenSql('Select T_REGIMETVA From TIERS Where T_REGIMETVA="'+TChoixCod.FindField(Fprefixe+'_CODE').AsString+'"',True,-1,'',true) ;
if Not Q.EOF then St:=Q.Fields[0].AsString else St:='' ;
Ferme(Q) ;
if St<>'' then BEGIN Msg.Execute(0,caption,'') ; Screen.Cursor:=SyncrDefault ; Exit ; END ;
St:='' ;
Q:=OpenSql('Select G_REGIMETVA From GENERAUX Where G_REGIMETVA="'+TChoixCod.FindField(Fprefixe+'_CODE').AsString+'"',True,-1,'',true) ;
if Not Q.EOF then St:=Q.Fields[0].AsString else St:='' ;
Ferme(Q) ;
if St<>'' then BEGIN Msg.Execute(1,caption,'') ; Screen.Cursor:=SyncrDefault ; Exit ; END ;
St:='' ;
Q:=OpenSql('Select TV_TAUXACH,TV_TAUXVTE,TV_CPTEACH,TV_CPTEVTE From TXCPTTVA '+
           'Where TV_REGIME="'+TChoixCod.FindField(Fprefixe+'_CODE').AsString+'"',True,-1,'',true) ;
Trouver:=False ;
While(Not Q.EOF)And(Not Trouver) do
     BEGIN
     if((Not IsFieldNull(Q,'TV_TAUXACH')) Or (Not IsFieldNull(Q,'TV_TAUXVTE')) Or
        (Q.Fields[2].AsString<>'') And (Not IsFieldNull(Q,'TV_CPTEACH')) Or
        (Q.Fields[3].AsString<>'') And (Not IsFieldNull(Q,'TV_CPTEVTE')))then Trouver:=True ;
     Q.Next ;
     END ;
Ferme(Q) ;
if Trouver then BEGIN Msg.Execute(2,caption,'') ; Screen.Cursor:=SyncrDefault ; Exit ; END ;
Result:=True ; Screen.Cursor:=SyncrDefault ;
END ;

procedure TFRegimeTVA.TChoixCodBeforeDelete(DataSet: TDataSet);
begin
inherited;
ExecuteSql('Delete From TXCPTTVA Where TV_REGIME="'+TChoixCod.FindField(Fprefixe+'_CODE').AsString+'"') ;
end;

end.
