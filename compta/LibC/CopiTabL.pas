unit CopiTabL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Ent1, Hent1, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  HSysMenu, hmsgbox, StdCtrls, Hctrls, Buttons,
  ExtCtrls, HStatus
  ,HPanel, UIUtil // MODIF PACK AVANCE pour gestion mode inside
  ;

Procedure CopiTableLibre ;

type
  TFCopiTabL = class(TForm)
    HPB: TPanel;
    Panel1: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    GBTabli: TGroupBox;
    TCbTlS: TLabel;
    CbEnS: THValComboBox;
    TCbTlD: TLabel;
    CbTlS: THValComboBox;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    CbEnD: THValComboBox;
    CbTlD: THValComboBox;
    TCbAxeD: TLabel;
    CbAxeD: THValComboBox;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure CbTlSChange(Sender: TObject);
    procedure CbEnSChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    Procedure RunLaCopie ;
    Procedure MajLongueur ;
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
Procedure CopiTableLibre ;
var FCopiTabL : TFCopiTabL ;
    PP : THPanel ;
begin
  FCopiTabL:=TFCopiTabL.Create(Application) ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FCopiTabL.ShowModal ;
      Finally
      FCopiTabL.Free ;
      End ;
    end
  else
    begin
    InitInside(FCopiTabL,PP) ;
    FCopiTabL.Show ;
    end ;

  SourisNormale ;
END ;

procedure TFCopiTabL.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFCopiTabL.FormShow(Sender: TObject);
begin
If CbEnS.Values.Count>0 Then CbEnS.Value:=CbEnS.Values[0] ;
If CbEnD.Values.Count>0 Then CbEnD.Value:=CbEnD.Values[0] ;
CbAxeD.ItemIndex:=0 ;
end;

procedure TFCopiTabL.CbTlSChange(Sender: TObject);
Var Pref : String ;
begin
Pref:=THValComboBox(Sender).Name[Length(THValComboBox(Sender).Name)] ;
if (THValComboBox(Sender).Value[1]='D') Or (THValComboBox(Sender).Value[1]='S') Or (THValComboBox(Sender).Value[1]='A') then
   THValComboBox(FindComponent('CbAxe'+Pref)).Enabled:=True else
   THValComboBox(FindComponent('CbAxe'+Pref)).Enabled:=False ;
end;

procedure TFCopiTabL.BValiderClick(Sender: TObject);
begin
HM.Execute(3,'', CbTlD.Items[CbTlD.ItemIndex]+' , elles vont être supprimées.'); // Attention : Si la liste des données de la table libre d'origine ne contient pas toutes les valeurs de la table libre de destination et si ces valeurs avaient été utilisées dans la table % 1 , elles vont être supprimées.
if HM.Execute(0,'','')<>mrYes then Exit ;
if CbTlS.Value=CbTlD.Value then BEGIN HM.Execute(1,'','') ; Exit ; END ;
BeginTrans ; RunLaCopie ; CommitTrans ;
end;

Procedure TFCopiTabL.RunLaCopie ;
Var Q,Q1 : TQuery ;
    Sql : String ;
    i : Integer ;
BEGIN
MajLongueur ;

Sql:='Delete From NATCPTE Where NT_TYPECPTE="'+CbTlD.Value+'"' ;
ExecuteSQL(Sql);

Q1:=OpenSql('Select * From NATCPTE Where NT_TYPECPTE="'+W_W+'"',False);
Q:=OpenSql('Select * from NATCPTE Where NT_TYPECPTE="'+CbTlS.Value+'"',True) ;
InitMove(RecordsCount(Q),'') ;
While Not Q.Eof do
  BEGIN
  MoveCur(False) ;
  Q1.Insert ; InitNew(Q1) ;
  for i:=0 to Q1.FieldCount-1 do
      BEGIN
      if Q1.Fields[i].FieldName='NT_TYPECPTE' then Q1.Fields[i].AsString:=CbTlD.Value
                                              else Q1.Fields[i].AsVariant:=Q.Fields[i].AsVariant ;
      END ;
  Q1.Post ; Q.Next ;
  END ;
Ferme(Q) ; Q1.Close ; FiniMove ;
MajDesChamps ;
HM.Execute(2,'','') ; // Recopie effectuée avec succès.
END ;

Procedure TFCopiTabL.MajLongueur ;
Var Ind1,Ind2,Long : Integer ;
BEGIN
Ind1:=1 ; Ind2:=1 ;
Case CbTlS.Value[1] of
     'G' : Ind1:=1 ; 'T' : Ind1:=2 ; 'S' : Ind1:=3 ; 'B' : Ind1:=4 ;
     'D' : Ind1:=5 ; 'E' : Ind1:=6 ; 'A' : Ind1:=7 ; 'U' : Ind1:=8 ;
    End ;
Case CbTlD.Value[1] of
     'G' : Ind2:=1 ; 'T' : Ind2:=2 ; 'S' : Ind2:=3 ; 'B' : Ind2:=4 ;
     'D' : Ind2:=5 ; 'E' : Ind2:=6 ; 'A' : Ind2:=7 ; 'U' : Ind2:=8 ;
    End ;
Long:=VH^.LgTableLibre[Ind1,StrToInt(Copy(CbTlS.Value,2,2))+1] ;
VH^.LgTableLibre[Ind2,StrToInt(Copy(CbTlD.Value,2,2))+1]:=Long ;
ExecuteSql('Update COMMUN Set CO_LIBRE="'+IntToStr(Long)+'" Where CO_CODE="'+CbTlD.Value+'"') ;
END ;

Procedure TFCopiTabL.MajDesChamps ;
Var Pref,Champ,Sql,WhereAx,Table : String ;
BEGIN
Pref:='' ; WhereAx:='' ;
Case CbTlD.Value[1] of
     'A' : BEGIN Pref:='Y_'  ; Table:='ANALYTIQ' ; WhereAx:=' AND Y_AXE="'+CbAxeD.Value+'"' ; END ;
     'B' : BEGIN Pref:='BG_' ; Table:='BUDGENE'  ; END ;
     'D' : BEGIN Pref:='BS_' ; Table:='BUDSECT'  ; WhereAx:=' AND BS_AXE="'+CbAxeD.Value+'"' ; END ;
     'E' : BEGIN Pref:='E_'  ; Table:='ECRITURE' ; END ;
     'G' : BEGIN Pref:='G_'  ; Table:='GENERAUX' ; END ;
     'S' : BEGIN Pref:='S_'  ; Table:='SECTION'  ; WhereAx:=' AND S_AXE="'+CbAxeD.Value+'"' ; END ;
     'T' : BEGIN Pref:='T_'  ; Table:='TIERS'    ; END ;
     'U' : BEGIN Pref:='BE_' ; Table:='BUDECR'   ; END ;
    End ;
Champ:=Pref+'TABLE'+IntToStr(StrToInt(Copy(CbTlD.Value,2,2))) ;
Sql:='Update '+Table+' Set '+Champ+'="" WHERE '+Champ+' NOT IN (SELECT NT_NATURE FROM NATCPTE WHERE NT_TYPECPTE="'+CbTlD.Value+'") '+WhereAx ;
ExecuteSQL(SQL);
END ;

procedure TFCopiTabL.CbEnSChange(Sender: TObject);
Var QLoc : TQuery ;
    Val : String ;
    C : THValComboBox ;
begin
C:=THValComboBox(FindComponent('CbTl'+THValComboBox(Sender).Name[Length(THValComboBox(Sender).Name)])) ;
Val:=THValComboBox(Sender).Value ;
QLoc:=OpenSql('Select CO_CODE,CO_LIBELLE From COMMUN Where CO_TYPE="NAT" And CO_CODE Like"'+Val+'%"',True) ;
C.Values.Clear ; C.Items.Clear ;
While Not QLoc.Eof do
   BEGIN
   C.Values.Add(QLoc.Fields[0].AsString) ; C.Items.Add(QLoc.Fields[1].AsString) ;
   QLoc.Next ;
   END ;
if C.Values.Count>0 then C.Value:=C.Values[0] ;
Ferme(QLoc) ;
end;

procedure TFCopiTabL.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFCopiTabL.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
