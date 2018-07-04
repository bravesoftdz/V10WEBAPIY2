unit CreRuCat;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB, {$IFNDEF DBXPRESS}dbtables, HSysMenu, hmsgbox, StdCtrls, Hctrls,
  Buttons, ExtCtrls{$ELSE}uDbxDataSet{$ENDIF}, StdCtrls, Hctrls, HSysMenu, hmsgbox, Buttons, ExtCtrls, HEnt1,
  Ent1, HStatus ;

Procedure CreationRubriqueParCategorie ;

type
  TFCrebuCat = class(TForm)
    PBouton: TPanel;
    Patience: TLabel;
    BAide: TBitBtn;
    BFerme: TBitBtn;
    BValider: TBitBtn;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    CbDelTout: TCheckBox;
    GbChoix: TGroupBox;
    Cb1: TCheckBox;
    Cb2: TCheckBox;
    Cb3: TCheckBox;
    Cb4: TCheckBox;
    TCbCat: TLabel;
    CbCat: THValComboBox;
    TCbAxe: TLabel;
    CbAxe: THValComboBox;
    QCpteG: TQuery;
    QCpteS: TQuery;
    QRub: TQuery;
    QRubRB_RUBRIQUE: TStringField;
    QRubRB_LIBELLE: TStringField;
    QRubRB_FAMILLES: TStringField;
    QRubRB_SIGNERUB: TStringField;
    QRubRB_TYPERUB: TStringField;
    QRubRB_COMPTE1: TStringField;
    QRubRB_EXCLUSION1: TStringField;
    QRubRB_COMPTE2: TStringField;
    QRubRB_EXCLUSION2: TStringField;
    QRubRB_AXE: TStringField;
    QRubRB_NATRUB: TStringField;
    QRubRB_TABLELIBRE: TStringField;
    QRubRB_BUDJAL: TStringField;
    QRubRB_CODEABREGE: TStringField;
    QRubRB_CATEGORIE: TStringField;
    procedure BFermeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure CbAxeChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    FaireQuoi : Array [1..4] of Boolean ;
    Combien : Integer ;
    ListG,ListS : TStringList ;
    Pref : String ;
    MsgPresence,CodeVide : Boolean ;
    SigneRub : String ;
    Procedure DetruitTout ;
    Function  OnFaitQuoi : Boolean ;
    Procedure DetruitParGroupe ;
    Procedure FabriqueRequeteCoiser ;
    Procedure RunLaCreation ;
    Procedure FaitListeCompte ;
    Function  PositionneCombien : Integer ;
    Procedure FabriqueRequeteCpte ;
    Procedure EcritLenreg(Q1,Q2 : TQuery ; Axe : String) ;
    Function  FabriqueCompte(St,St1 : String) : String ;
    Function  FabriqueCompteAbr(St,St1 : String) : String ;
    Function  FabriqueCompteRub(Cpte,Pre : String) : String ;
    Function  FabriqueCompteRubAtt(St : String) : String ;
    Procedure ChercheSigneRub ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure CreationRubriqueParCategorie ;
var FCrebuCat : TFCrebuCat ;
BEGIN
FCrebuCat:=TFCrebuCat.Create(Application) ;
 Try
 FCrebuCat.ShowModal ;
 Finally
 FCrebuCat.Free ;
 End ;
SourisNormale ;
END ;

procedure TFCrebuCat.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFCrebuCat.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCrebuCat.FormShow(Sender: TObject);
begin
CodeVide:=False ; MsgPresence:=False ; CbAxe.Value:=CbAxe.Values[0] ;
end;

procedure TFCrebuCat.CbAxeChange(Sender: TObject);
begin
Case CbAxe.Value[2] of
     '1' : CbCat.DataType:='ttCatJalBud1' ;
     '2' : CbCat.DataType:='ttCatJalBud2' ;
     '3' : CbCat.DataType:='ttCatJalBud3' ;
     '4' : CbCat.DataType:='ttCatJalBud4' ;
     '5' : CbCat.DataType:='ttCatJalBud5' ;
   End ;
CbCat.Reload ;
if CbCat.Items.Count>0 then CbCat.ItemIndex:=0 ;
end;

Function TFCrebuCat.OnFaitQuoi : Boolean ;
Var i : Integer ;
BEGIN
Result:=False ; Combien:=0 ;
FillChar(FaireQuoi,SizeOf( FaireQuoi),False) ;
for i:=1 to 4 do
   if (TCheckBox(FindComponent('Cb'+IntToStr(i))).State=cbChecked) then
      BEGIN
      FaireQuoi[i]:=True ; Inc(Combien) ; Result:=True ;
      END ;
END ;

Procedure TFCrebuCat.DetruitTout ;
BEGIN
if Not CbDelTout.Checked then Exit ;
BeginTrans ;
ExecuteSql('Delete From RUBRIQUE Where RB_CATEGORIE="X" And RB_NATRUB="BUD"') ;
CommitTrans ;
END ;

Procedure TFCrebuCat.DetruitParGroupe ;
Var i : Integer ;
BEGIN
if CbDelTout.Checked then Exit ;
BeginTrans ;
for i:=1 to 4 do
   BEGIN
   Case i of
        1 : if FaireQuoi[i] then ExecuteSql('Delete From RUBRIQUE Where RB_FAMILLES="CBG;" And RB_NATRUB="BUD" And RB_CATEGORIE="X"') ;
        2 : if FaireQuoi[i] then ExecuteSql('Delete From RUBRIQUE Where RB_FAMILLES="CBS;" And RB_NATRUB="BUD" And RB_CATEGORIE="X"') ;
        3 : if FaireQuoi[i] then ExecuteSql('Delete From RUBRIQUE Where RB_FAMILLES="G/S;" And RB_NATRUB="BUD" And RB_CATEGORIE="X"') ;
        4 : if FaireQuoi[i] then ExecuteSql('Delete From RUBRIQUE Where RB_FAMILLES="S/G;" And RB_NATRUB="BUD" And RB_CATEGORIE="X"') ;
      End ;
   END ;
CommitTrans ;
END ;

procedure TFCrebuCat.BValiderClick(Sender: TObject);
begin
if CbCat.Value='' then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if Not OnFaitQuoi then BEGIN HM.Execute(1,'','') ; Exit ; END ;
if HM.Execute(2,'','')<>mrYes then Exit ;
DetruitTout ; DetruitParGroupe ;
RunLaCreation ;
end;

Procedure TFCrebuCat.FabriqueRequeteCoiser ;
Var SqlG,SqlS : String ;
BEGIN
QCpteG.Close ; QCpteG.Sql.Clear ; QCpteS.Close ; QCpteS.Sql.Clear ;
SqlG:='Select Distinct CX_COMPTE From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+CbCat.Value+'" '+
      'And CX_CATEGORIE="'+CbCat.Value+'" ' ;
SqlS:='Select Distinct CX_SECTION From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+CbCat.Value+'" '+
      'And CX_CATEGORIE="'+CbCat.Value+'" ' ;
QCpteG.Sql.Add(SqlG) ; ChangeSql(QCpteG) ;
QCpteS.Sql.Add(SqlS) ; ChangeSql(QCpteS) ;
END ;

Procedure TFCrebuCat.FaitListeCompte ;
BEGIN
QCpteG.Open ; QCpteS.Open ;
While Not QCpteG.Eof do BEGIN ListG.Add(QCpteG.Fields[0].AsString) ; QCpteG.Next ; END ;
While Not QCpteS.Eof do BEGIN ListS.Add(QCpteS.Fields[0].AsString) ; QCpteS.Next ; END ;
QCpteG.Close ; QCpteS.Close ;
END ;

Procedure TFCrebuCat.FabriqueRequeteCpte ;
BEGIN
QCpteG.Close ; QCpteG.Sql.Clear ;
QCpteS.Close ; QCpteS.Sql.Clear ;
QCpteG.Sql.Add('Select * From BUDGENE Where BG_BUDGENE=:CpteG') ;
QCpteS.Sql.Add('Select * From BUDSECT Where BS_AXE=:Ax And BS_BUDSECT=:CpteS') ;
ChangeSql(QCpteG) ; QCpteG.Prepare ;
ChangeSql(QCpteS) ; QCpteS.Prepare ;
END ;

Procedure TFCrebuCat.RunLaCreation ;
Var i,j,k,l : Integer ;
BEGIN
ListG:=TStringList.Create ; ListS:=TStringList.Create ;
FabriqueRequeteCoiser ; FaitListeCompte ; FabriqueRequeteCpte ;
QRub.Close ; ChangeSql(QRub) ; QRub.Prepare ; QRub.Open ;
Combien:=PositionneCombien ; CodeVide:=False ; MsgPresence:=False ;
InitMove(Combien,'') ; ChercheSigneRub ;
for i:=1 to 4 do
   BEGIN
   Case i of
       1:BEGIN
         if FaireQuoi[i] then
            BEGIN
            Pref:='CBG' ;
            BeginTrans ;
            for j:=0 to ListG.Count-1 do
                BEGIN
                QCpteG.Close ; MoveCur(False) ;
                if ListG.Strings[j]<>'' then
                   BEGIN
                   QCpteG.ParamByName('CpteG').AsString:=ListG.Strings[j] ;
                   QCpteG.Open ; EcritLenreg(QCpteG,Nil,'') ;
                   END ;
                END ;
             CommitTrans ;
            END ;
         END ;
       2:BEGIN
         if FaireQuoi[i] then
            BEGIN
            Pref:='CBS' ;
            BeginTrans ;
            for j:=0 to ListS.Count-1 do
                BEGIN
                QCpteS.Close ; MoveCur(False) ;
                if ListS.Strings[j]<>'' then
                   BEGIN
                   QCpteS.ParamByName('CpteS').AsString:=ListS.Strings[j] ;
                   QCpteS.ParamByName('Ax').AsString:=CbAxe.Value ;
                   QCpteS.Open ; EcritLenreg(QCpteS,Nil,CbAxe.Value) ;
                   END ;
                END ;
             CommitTrans ;
            END ;
         END ;
       3:BEGIN
         Pref:='G/S' ;
         l:=0 ;
         BeginTrans ;
         for j:=0 to ListG.Count-1 do
            BEGIN
            MoveCur(False) ;
            for k:=0 to ListS.Count-1 do
                BEGIN
                if (j*k)-l>2000 then BEGIN l:=j*k ; CommitTrans ; BeginTrans ; END ;
                if (ListG.Strings[j]<>'') And (ListS.Strings[k]<>'') then
                    BEGIN
                    QCpteG.Close ; QCpteS.Close ;
                    QCpteG.ParamByName('CpteG').AsString:=ListG.Strings[j] ;
                    QCpteS.ParamByName('CpteS').AsString:=ListS.Strings[k] ;
                    QCpteS.ParamByName('Ax').AsString:=CbAxe.Value ;
                    QCpteG.Open ; QCpteS.Open ;
                    EcritLenreg(QCpteG,QCpteS,CbAxe.Value) ;
                    END ;
                END ;
            END ;
         CommitTrans ;
         END ;
       4:BEGIN
         Pref:='S/G' ;
         l:=0 ;
         BeginTrans ;
         for j:=0 to ListS.Count-1 do
            BEGIN
            MoveCur(False) ;
            for k:=0 to ListG.Count-1 do
                BEGIN
                if (j*k)-l>2000 then BEGIN l:=j*k ; CommitTrans ; BeginTrans ; END ;
                if (ListS.Strings[j]<>'') And (ListG.Strings[k]<>'') then
                    BEGIN
                    QCpteG.Close ; QCpteS.Close ;
                    QCpteG.ParamByName('CpteG').AsString:=ListG.Strings[k] ;
                    QCpteS.ParamByName('CpteS').AsString:=ListS.Strings[j] ;
                    QCpteS.ParamByName('Ax').AsString:=CbAxe.Value ;
                    QCpteG.Open ; QCpteS.Open ; EcritLenreg(QCpteS,QCpteG,CbAxe.Value) ;
                    END ;
                END ;
            END ;
         CommitTrans ;
         END ;
    End ;
   END ;
FiniMove ; ListG.Free ; ListS.Free ;
if MsgPresence then HM.Execute(3,'','') ;
if CodeVide then HM.Execute(4,'','') ;
END ;

Function TFCrebuCat.PositionneCombien : Integer ;
Var i :Integer ;
BEGIN
i:=0 ;
if FaireQuoi[1] then i:=i+ListG.Count ; if FaireQuoi[2] then i:=i+ListS.Count ;
if FaireQuoi[3] then i:=i+ListG.Count ; if FaireQuoi[4] then i:=i+ListS.Count ;
Result:=i ;
END ;

Procedure TFCrebuCat.EcritLenreg(Q1,Q2 : TQuery ; Axe : String) ;
Var CpteRu,TypRub,Pre1,Pre2,Ax,Champ1,Champ2,Cpt1,
    Cpt2,ChaineCpte,Cle1,Cle2,CPtAtt1,CPtAtt2,LeLib,CpteRuAbr : String ;
BEGIN
if (Q1=Nil) And (Q2=Nil) then Exit ;
if UpperCase(TQuery(Q1).Name)='QCPTEG' then
   BEGIN
   Pre1:='BG_' ; Champ1:='BG_RUB' ; TypRub:='GEN' ; Cpt1:='BG_COMPTERUB' ; Cle1:='BG_BUDGENE' ;
   END else
   BEGIN
   Pre1:='BS_' ; Champ1:='BS_RUB' ; TypRub:='ANA' ; Cpt1:='BS_SECTIONRUB' ; Cle1:='BS_BUDSECT' ;
   END ;
if Q2<>Nil then
   BEGIN
   if UpperCase(TQuery(Q2).Name)='QCPTEG' then
      BEGIN
      Pre2:='BG_' ; Champ2:='BG_RUB' ; TypRub:='A/G' ; Cpt2:='BG_COMPTERUB' ; Cle2:='BG_BUDGENE' ;
      END else
      BEGIN
      Pre2:='BS_' ; Champ2:='BS_RUB' ; TypRub:='G/A' ; Cpt2:='BS_SECTIONRUB' ; Cle2:='BS_BUDSECT' ;
      END ;
   END ;
if Q1.FindField(Cle1).AsString<>CptAtt1 then
  if Pos(';',Q1.FindField(CPt1).AsString)=1 then Exit ;

if Q2<>Nil then
   if Q2.FindField(Cle2).AsString<>CptAtt2 then
      if Pos(';',Q2.FindField(CPt2).AsString)=1 then Exit ;
if Q1.FindField(Champ1).AsString='' then BEGIN CodeVide:=True ; Exit ; END ;
if Q2<>Nil then if Q2.FindField(Champ2).AsString='' then BEGIN CodeVide:=True ; Exit ; END ;

if Q2<>Nil then
   BEGIN
   CpteRu:=Pref+CbCat.Value+FabriqueCompte(Q1.FindField(Champ1).AsString,Q2.FindField(Champ2).AsString) ;
   CpteRuAbr:=FabriqueCompteAbr(Q1.FindField(Champ1).AsString,Q2.FindField(Champ2).AsString) ;
   LeLib:=Q1.FindField(Pre1+'LIBELLE').AsString+Q2.FindField(Pre2+'LIBELLE').AsString ;
   END else
   BEGIN
   CpteRu:=Pref+CbCat.Value+FabriqueCompte(Q1.FindField(Champ1).AsString,'') ;
   CpteRuAbr:=FabriqueCompteAbr(Q1.FindField(Champ1).AsString,'') ;
   LeLib:=Q1.FindField(Pre1+'LIBELLE').AsString ;
   END ;
if Length(LeLib)>35 then LeLib:=Copy(LeLib,1,35) ;
if Presence('RUBRIQUE','RB_RUBRIQUE',CpteRu) then
   BEGIN MsgPresence:=True ; Exit ; END ;
if Q1.FindField(Cle1).AsString<>CptAtt1 then ChaineCpte:=FabriqueCompteRub(Q1.FindField(Cpt1).AsString,Pre1)
                                        else ChaineCpte:=FabriqueCompteRubAtt(Pre1) ;
QRub.Insert ; InitNew(QRub) ;
QRubRB_RUBRIQUE.AsString:=CpteRu ;
QRubRB_CODEABREGE.AsString:=CpteRuAbr ;
QRubRB_LIBELLE.AsString:=LeLib ;
QRubRB_FAMILLES.AsString:=Pref+';' ;
QRubRB_SIGNERUB.AsString:=SigneRub ;
QRubRB_TYPERUB.AsString:=TypRub ;
QRubRB_COMPTE1.AsString:=ChaineCpte ;
QRubRB_EXCLUSION1.AsString:=Q1.FindField(Pre1+'EXCLURUB').AsString ;
if Q2<>Nil then
   BEGIN
   if Q2.FindField(Cle2).AsString<>CptAtt2 then ChaineCpte:=FabriqueCompteRub(Q2.FindField(Cpt2).AsString,Pre2)
                                           else ChaineCpte:=FabriqueCompteRubAtt(Pre2) ;
   QRubRB_COMPTE2.AsString:=ChaineCpte ;
   QRubRB_EXCLUSION2.AsString:=Q2.FindField(Pre2+'EXCLURUB').AsString ;
   END ;
QRubRB_AXE.AsString:=Axe ;
QRubRB_NATRUB.AsString:='BUD' ;
QRubRB_BUDJAL.AsString:=CbCat.Value ;
QRubRB_CATEGORIE.AsString:='X' ;
QRub.Post ;
END ;

Function TFCrebuCat.FabriqueCompte(St,St1 : String) : String ;
BEGIN
Result:='' ;
if St1='' then Result:=St
          else Result:=Copy(St,1,5)+':'+Copy(St1,1,5) ;
END ;

Function TFCrebuCat.FabriqueCompteAbr(St,St1 : String) : String ;
BEGIN
Result:='' ;
if St1='' then Result:=St
          else Result:=Copy(St,1,3)+Copy(St1,1,3) ;
END ;

Function TFCrebuCat.FabriqueCompteRubAtt(St : String) : String ;
Const TypCalc = '(SM)' ;
BEGIN
if St='BG_' then Result:=VH^.Cpta[fbGene].Attente+TypCalc+';'
            else Result:=VH^.Cpta[AxeToFb(CbAxe.Value)].Attente+TypCalc+';'
END ;

Function TFCrebuCat.FabriqueCompteRub(Cpte,Pre : String) : String ;
Const TypCalc = '(SM)' ;
Var St,St1 : String ;
    TabSousPlan : TSousPlanCat ;
    LesComptes,Code : String ;
    i,j,Deb,Lon : Integer ;
    CodIns : String ;
BEGIN
Code:=Cpte ; LesComptes:='' ;
if Pre='BG_' then
   BEGIN
   While Code<>'' do
       BEGIN
       St:=ReadTokenSt(Code) ;
       if St<>'' then LesComptes:=LesComptes+St+TypCalc+';' ;
       END ;
   END else
   BEGIN
   FillChar(TabSousPlan,SizeOf(TabSousPlan),#0) ;
   TabSousPlan:=SousPlanCat(CbCat.Value,True) ;
   While Code<>'' do
      BEGIN
      St:=ReadTokenSt(Code) ;
      if St<>'' then
         BEGIN
         for i:=1 to MaxSousPlan do
            BEGIN
            if TabSousPlan[i].Code='' then Break ;
            Deb:=TabSousPlan[i].Debut ; Lon:=TabSousPlan[i].Longueur ;
            CodIns:='' ;
            for j:=1 to Lon do CodIns:=CodIns+'?' ;
            Insert(CodIns,St,Deb) ;
            END ;
         END ;
      if St<>'' then LesComptes:=LesComptes+St+TypCalc+';' ;
      END ;
   END ;
Result:=LesComptes+';' ;
END ;

Procedure TFCrebuCat.ChercheSigneRub ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select BJ_BUDJAL,BJ_NATJAL From BUDJAL Where BJ_CATEGORIE="'+CbCat.Value+'" Order by BJ_BUDJAL',True) ;
if (QLoc.Fields[1].AsString='CHA') or (QLoc.Fields[1].AsString='') then SigneRub:='POS' else SigneRub:='NEG' ;
Ferme(QLoc) ;
END ;

end.
