unit PssToTli;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, hmsgbox, HSysMenu, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hctrls, Ent1,
  HEnt1, HStatus
  ,HPanel, UIUtil // MODIF PACK AVANCE pour gestion mode inside
  ;

Procedure CopiePlanSousSectionToTableLibre ;

type
  TFPssToTli = class(TForm)
    GBTabli: TGroupBox;
    TCbTlS: TLabel;
    CbPss: THValComboBox;
    GBAxe: TGroupBox;
    Q1: TQuery;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    RgEntite: TRadioGroup;
    TCbAxeD: TLabel;
    CbAxeS: THValComboBox;
    TCbTlD: TLabel;
    CbTlD: THValComboBox;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CbAxeSChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure RgEntiteClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Long : Integer ;
    Procedure ChargeTableLibre ;
    Procedure MajLongueur ;
    Procedure RunLaCopie ;
    Procedure MajDesChamps ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... :   /  /
Description .. : 
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion 
Suite ........ : mode inside
Mots clefs ... : 
*****************************************************************}
Procedure CopiePlanSousSectionToTableLibre ;
var FPssToTli : TFPssToTli ;
    PP : THPanel ;
BEGIN
  FPssToTli:=TFPssToTli.Create(Application) ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FPssToTli.ShowModal ;
      Finally
      FPssToTli.Free ;
      End ;
    end
  else
    begin
    InitInside(FPssToTli,PP) ;
    FPssToTli.Show ;
    end ;

  SourisNormale ;
END ;

procedure TFPssToTli.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFPssToTli.FormShow(Sender: TObject);
begin
RgEntiteClick(Nil) ;
If CbAxeS.Values.Count>0 Then CbAxeS.Value:=CbAxeS.Values[0] ;
If CbTlD.Values.Count>0 Then CbTlD.Value:=CbTlD.Values[0] ;
end;

Procedure TFPssToTli.ChargeTableLibre ;
Var Sql : String ;
    Pref : String ;
BEGIN
Case RgEntite.ItemIndex of
     0 : Pref:='S' ;
     1 : Pref:='D' ;
 End ;
CbTlD.Values.Clear ; CbTlD.Items.Clear ;
Sql:='Select CC_CODE,CC_LIBELLE From CHOIXCOD Where CC_TYPE="NAT" And CC_CODE Like "'+Pref+'%"' ;
Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(Sql) ; ChangeSql(Q1) ; Q1.Open ;
While Not Q1.Eof do
  BEGIN
  CbTlD.Values.Add(Q1.Fields[0].AsString) ; CbTlD.Items.Add(Q1.Fields[1].AsString) ;
  Q1.Next ;
  END ;
Q1.Close ; If CbTlD.Values.Count>0 Then CbTlD.Value:=CbTlD.Values[0] ;
END ;

procedure TFPssToTli.CbAxeSChange(Sender: TObject);
Var Sql : String ;
begin
CbPss.Values.Clear ; CbPss.Items.Clear ;
Sql:='Select SS_SOUSSECTION,SS_LIBELLE,SS_LONGUEUR From STRUCRSE Where SS_AXE="'+CbAxeS.Value+'"' ;
Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(Sql) ; ChangeSql(Q1) ; Q1.RequestLive:=False ; Q1.Open ;
While Not Q1.Eof do
  BEGIN
  CbPss.Values.Add(Q1.Fields[0].AsString) ; CbPss.Items.Add(Q1.Fields[1].AsString) ;
  Q1.Next ;
  END ;
if CbPss.Values.Count>0 then CbPss.Value:=CbPss.Values[0] ;
Q1.Close ;
end;

procedure TFPssToTli.BValiderClick(Sender: TObject);
begin
if CbPss.Value='' then Exit ;
if HM.Execute(0,'','')<>mrYes then Exit ;
RunLaCopie ;
end;

procedure TFPssToTli.MajLongueur ;
Var Sql : String ;
    Ind : Integer ;
BEGIN
Sql:='Update CHOIXCOD Set CC_LIBRE="'+IntToStr(Long)+'" Where CC_CODE="'+CbTlD.Value+'"' ;
Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(Sql) ; ChangeSql(Q1) ; Q1.RequestLive:=False ; Q1.ExecSql ;
Ind:=3 ;
Case CbTlD.Value[1] of
     'S' : Ind:=3 ;
     'D' : Ind:=5 ;
    End ;
VH^.LgTableLibre[Ind,StrToInt(Copy(CbTlD.Value,2,2))+1]:=Long ;
END ;

Procedure TFPssToTli.RunLaCopie ;
Var Sql : String ;
    QLoc : TQuery ;
BEGIN
Sql:='Delete From NATCPTE Where NT_TYPECPTE="'+CbTlD.Value+'"' ;
Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(Sql) ; ChangeSql(Q1) ; Q1.RequestLive:=False ; Q1.ExecSql ;

Sql:='Select PS_CODE,PS_LIBELLE From SSSTRUCR Where PS_AXE="'+CbAxeS.Value+'" And PS_SOUSSECTION="'+CbPss.Value+'"' ;
QLoc:=OpenSql(Sql,True) ;

Sql:='Select * From NATCPTE Where NT_TYPECPTE="'+W_W+'"' ;
Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(Sql) ; ChangeSql(Q1) ; Q1.RequestLive:=True ; Q1.Open ;
InitMove(RecordsCount(QLoc),'') ;
Long:=Length(QLoc.FindField('PS_CODE').AsString) ;
While Not QLoc.Eof do
  BEGIN
  MoveCur(False) ;
  Q1.Insert ; InitNew(Q1) ;
  Q1.FindField('NT_TYPECPTE').AsString:=CbTlD.Value ;
  Q1.FindField('NT_NATURE').AsString:=QLoc.FindField('PS_CODE').AsString ;
  Q1.FindField('NT_LIBELLE').AsString:=QLoc.FindField('PS_LIBELLE').AsString ;
  Q1.Post ; QLoc.Next ;
  END ;
Ferme(QLoc) ; Q1.Close ; FiniMove ; MajDesChamps ;
if Long<>0 then MajLongueur ;
END ;

Procedure TFPssToTli.MajDesChamps ;
Var Pref,Champ,Sql,WhereAx,Table : String ;
BEGIN
Pref:='' ; WhereAx:='' ;
Case CbTlD.Value[1] of
     'S' : BEGIN Pref:='S_' ; Table:='SECTION' ; WhereAx:=' Where S_AXE="'+CbAxeS.Value+'"' ; END ;
     'D' : BEGIN Pref:='BS_' ; Table:='BUDSECT' ; WhereAx:=' Where BS_AXE="'+CbAxeS.Value+'"' ; END ;
    End ;
Champ:=Pref+'TABLE'+IntToStr(StrToInt(Copy(CbTlD.Value,2,2))) ;
Sql:='Update '+Table+' Set '+Champ+'=""'+WhereAx ;
Q1.Close ; Q1.Sql.Clear ; Q1.Sql.Add(Sql) ; ChangeSql(Q1) ;
Q1.RequestLive:=False ; Q1.ExecSql ; Q1.Close ;
END ;

procedure TFPssToTli.RgEntiteClick(Sender: TObject);
begin ChargeTableLibre ; end;

procedure TFPssToTli.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
