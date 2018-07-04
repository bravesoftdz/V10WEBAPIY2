{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CreRubBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Ent1, HEnt1, hmsgbox, Hctrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} 
{$ENDIF}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HSysMenu, HCompte, HPanel, UiUtil, HTB97,
  UTOB ;

Procedure CreationRubrique ;

Const TypeCreation = 'BUD' ;

type
  TFCreRubBu = class(TForm)
    HM: THMsgBox;
    GbChoix: TGroupBox;
    Cb1: TCheckBox;
    Cb2: TCheckBox;
    Cb3: TCheckBox;
    Cb4: TCheckBox;
    Label1: TLabel;
    CbJal: THValComboBox;
    CbRubG: THValComboBox;
    CbRubS: THValComboBox;
    HMTrad: THSystemMenu;
    RgChoix: TRadioGroup;
    CbDelTout: TCheckBox;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    PFen: TPanel;
    Patience: TLabel;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CbJalChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    FaireQuoi : Array [1..4] of Boolean ;
    Pref,Laxe,CptGenAtt,CptSecAtt,NatJal : String ;
    ListG,ListS : TStringList ;
    Combien : Integer ;
    MsgPresence,CodeVide,OnCategorie : Boolean ;
    TabSousPlan : TSousPlanCat ;
    QuelCodeSp,LaCat : String ;
    Function  OnFaitQuoi : Boolean ;
    Procedure RunLaCreation ;
//    Procedure FabriqueRequete ;
//    Procedure FabriqueRequeteCroiser ;
//    Procedure FabriqueRequeteCpte ;
    Procedure FabriqueListeCompte ;
    Procedure DeglingueLesRubriques(i : Integer) ;
    Procedure EffaceRub(St : String) ;
    Function  FabriqueCompte(St,St1 : String) : String ;
    Function  FabriqueCompteAbr(St,St1 : String) : String ;
    Procedure InsereLesEnreg(i : Integer ; TOBG,TOBS : TOB ; Var TOBRUB : TOB) ;
    Procedure EcritLenreg(TOBLG,TOBLS : TOB ; Var TOBRUB : TOB ; Axe : String) ;
    Function  FabriqueCompteRub(St,Prefi : String) : String ;
    Function  FabriqueCompteRubAtt(St : String) : String ;
    Function  TrouveEcrBud(StG,StS : String) : Boolean ;
    Function  ExisteCroise ( Bud,Sect : String ) : boolean ;
    Procedure RempliCbRub ;
    Function  PositionneCombien : Integer ;
    Procedure InitAxeCptAttReqDel ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus;


Procedure EcritTobRUB(Var TC : Tob ; Fin : Boolean = FALSE) ;
BEGIN
If TC=NIL Then Exit ;
If Fin Then TC.INSERTDB(NIL,TRUE) Else
  BEGIN
  If TC.Detail.Count>300 Then
    BEGIN
    TC.INSERTDB(NIL,TRUE) ;
    TC.Free ;
    CommitTrans ; Delay(1000) ; BeginTrans ;
    TC:=TOB.Create('',Nil,-1) ;
    END ;
  END ;
END ;

Procedure CreationRubrique ;
var FCreRubBu : TFCreRubBu ;
    PP : THPanel ;
BEGIN
FCreRubBu:=TFCreRubBu.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FCreRubBu.ShowModal ;
    Finally
     FCreRubBu.Free ;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FCreRubBu,PP) ;
   FCreRubBu.Show ;
   END ;
END ;

procedure TFCreRubBu.BFermeClick(Sender: TObject);
begin
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

procedure TFCreRubBu.FormShow(Sender: TObject);
begin
If CbJal.Values.Count>0 Then CbJal.Value:=CbJal.Values[0] ;
Pref:='' ; Laxe:='' ; MsgPresence:=False ; CodeVide:=False ;
end;

Function TFCreRubBu.OnFaitQuoi : Boolean ;
Var i : Integer ;
BEGIN
Result:=False ; Combien:=0 ;
FillChar(FaireQuoi,SizeOf(FaireQuoi),False) ;
for i:=1 to 4 do
   if (TCheckBox(FindComponent('Cb'+IntToStr(i))).State=cbChecked) then
      BEGIN
      FaireQuoi[i]:=True ; Inc(Combien) ; Result:=True ;
      END ;
END ;

procedure TFCreRubBu.BValiderClick(Sender: TObject);
begin
if CbJal.Text='' then BEGIN HM.Execute(0,'','') ; Exit ; END ;
if Not OnFaitQuoi then BEGIN HM.Execute(1,'','') ; Exit ; END ;
if CbDelTout.Checked then
   BEGIN
   if HM.Execute(2,'','')<>mrYes then Exit ;
   END else
   BEGIN
   if HM.Execute(6,'','')<>mrYes then Exit ;
   END ;
Patience.Visible:=TRUE ; Application.ProcessMessages ;
if CbDelTout.Checked then ExecuteSql('Delete From RUBRIQUE Where RB_RUBRIQUE Like "___'+CbJal.Value+'%"') ;
HPB.Enabled:=False ; RunLaCreation ;
Patience.Visible:=FALSE ; HPB.Enabled:=True ; Application.ProcessMessages ;
Close;
if IsInside(Self) then CloseInsidePanel(Self);
end;

Function TFCreRubBu.PositionneCombien : Integer ;
Var i :Integer ;
BEGIN
i:=0 ;
if FaireQuoi[1] then i:=i+ListG.Count ;
if FaireQuoi[2] then i:=i+ListS.Count ;
if FaireQuoi[3] then i:=i+ListG.Count ;
if FaireQuoi[4] then i:=i+ListS.Count ;
Result:=i ;
END ;

Procedure TFCreRubBu.RunLaCreation ;
Var i : Integer ;
    TOBRUB : TOB ;
    Q : TQuery ;
    TOBG,TOBS : TOB ;
BEGIN
ListG:=TStringList.Create ; ListS:=TStringList.Create ;
//if RgChoix.ItemIndex=0 then FabriqueRequete else FabriqueRequeteCroiser ;
FabriqueListeCompte ; //FabriqueRequeteCpte ;
(*
QRub.Close ; ChangeSql(QRub) ; //QRub.Prepare ;
PrepareSQLODBC(QRub) ;
QRub.Open ;
*)
RempliCbRub ;
(*
ChangeSql(QEcr) ; //QEcr.Prepare ;
PrepareSQLODBC(QEcr) ;
*)
if Not CbDelTout.Checked then Combien:=2*PositionneCombien else Combien:=PositionneCombien ;
InitMove(Combien,HM.Mess[5]) ;
TOBG:=TOB.Create('',Nil,-1) ; TOBS:=TOB.Create('',Nil,-1) ;
Q:=openSQL('Select * From BUDGENE',TRUE) ; TOBG.LoadDetailDB('BUDGENE','','',Q,FALSE,FALSE) ; Ferme(Q) ;
Q:=openSQL('Select * From BUDSECT',TRUE) ; TOBS.LoadDetailDB('BUDSECT','','',Q,FALSE,FALSE) ; Ferme(Q) ;
for i:=1 to 4 do
  BEGIN
  BEGINTRANS ;
  TOBRUB:=TOB.Create('',Nil,-1) ;
  Case i of
      1 : BEGIN
          if FaireQuoi[i] then
             BEGIN
             Pref:='CBG' ;
             if (ListG<>Nil) And (ListG.Count>0) then
                BEGIN DeglingueLesRubriques(i) ; InsereLesEnreg(i,TOBG,TOBS,TOBRUB) ; END ;
             END ;
          END ;
      2 : BEGIN
          if FaireQuoi[i] then
             BEGIN
             Pref:='CBS' ;
             if (ListS<>Nil) And (ListS.Count>0) then
                BEGIN DeglingueLesRubriques(i) ; InsereLesEnreg(i,TOBG,TOBS,TOBRUB) ; END ;
             END ;
          END ;
      3 : BEGIN
          if FaireQuoi[i] then
             BEGIN
             Pref:='G/S' ;
             if ((ListG<>Nil) And (ListG.Count>0)) And ((ListS<>Nil) And (ListS.Count>0)) then
                BEGIN DeglingueLesRubriques(i) ; InsereLesEnreg(i,TOBG,TOBS,TOBRUB) ; END ;
             END ;
          END ;
      4 : BEGIN
          if FaireQuoi[i] then
             BEGIN
             Pref:='S/G' ;
             if ((ListG<>Nil) And (ListG.Count>0)) And ((ListS<>Nil) And (ListS.Count>0)) then
                BEGIN DeglingueLesRubriques(i) ; InsereLesEnreg(i,TOBG,TOBS,TOBRUB) ; END ;
             END ;
          END ;
    End ;
//  TOBRUB.InsertDB(NIL,FALSE) ;
  EcritTobRUB(TOBRUB,TRUE) ;
  TOBRUB.Free ;
  COMMITTRANS ;
  END ;
ListG.Free ; ListS.Free ; FiniMove ;
TOBG.Free ; TOBS.Free ;
END ;

Function TFCreRubBu.FabriqueCompteAbr(St,St1 : String) : String ;
BEGIN
if St1='' then Result:=St
          else Result:=Copy(St,1,3)+Copy(St1,1,3) ;
END ;

Function TFCreRubBu.FabriqueCompte(St,St1 : String) : String ;
BEGIN
if St1='' then Result:=St
          else Result:=Copy(St,1,5)+':'+Copy(St1,1,5) ;
END ;

Procedure TFCreRubBu.EffaceRub(St : String) ;
Var Ax : String ;
    SQL : String ;
BEGIN
(*
BeginTrans ;
St:=Pref+CbJal.Value+St ;
if Pos('S',Pref)>0 then Ax:=Laxe else Ax:=W_W ;
QDel.ParamByName('LeCpte').AsString:=Copy(St,1,17) ;
QDel.ParamByName('Fam').AsString:=Pref+';' ;
QDel.ParamByName('Ax').AsString:=Ax ;
QDel.ExecSql ;
CommitTrans ;
*)
//  QDel.Sql.Add('Delete From RUBRIQUE Where RB_RUBRIQUE=:LeCpte And RB_FAMILLES=:Fam And RB_AXE=:Ax') ;
//BeginTrans ;
St:=Pref+CbJal.Value+St ;
if Pos('S',Pref)>0 then Ax:=Laxe else Ax:=W_W ;
SQL:='Delete From RUBRIQUE Where RB_RUBRIQUE="'+Copy(St,1,17)+'" And RB_FAMILLES="'+Pref+';'+'" And RB_AXE="'+Ax+'" ' ;
ExecuteSQL(SQL) ;
//CommitTrans ;
END ;

Procedure TFCreRubBu.RempliCbRub ;
Var QLoc : TQuery ;
BEGIN
CbRubG.Values.Clear ; CbRubG.Items.Clear ; CbRubS.Values.Clear ; CbRubS.Items.Clear ;
QLoc:=OpenSql('Select BG_BUDGENE,BG_RUB From BUDGENE Order By BG_BUDGENE',True) ;
While Not QLoc.Eof do
    BEGIN
    CbRubG.Values.Add(Trim(QLoc.FindField('BG_BUDGENE').AsString)) ;
    CbRubG.Items.Add(Trim(QLoc.FindField('BG_RUB').AsString)) ;
    QLoc.Next ;
    END ;
Ferme(QLoc) ;
QLoc:=OpenSql('Select BS_BUDSECT,BS_RUB From BUDSECT Order By BS_BUDSECT',True) ;
While Not QLoc.Eof do
    BEGIN
    CbRubS.Values.Add(Trim(QLoc.FindField('BS_BUDSECT').AsString)) ;
    CbRubS.Items.Add(Trim(QLoc.FindField('BS_RUB').AsString)) ;
    QLoc.Next ;
    END ;
Ferme(QLoc) ;
END ;

Procedure TFCreRubBu.DeglingueLesRubriques(i : Integer) ;
Var StG,StS,St : String ;
    j,k : Integer ;
BEGIN
if CbDelTout.Checked then Exit ;
//QDel.Close ;
StG:='' ; StS:='' ;
Case i of
     1 : BEGIN
         ExecuteSql('Update RUBRIQUE Set RB_AXE="'+W_W+'" Where RB_RUBRIQUE Like "CBG%" And RB_NATRUB="BUD"') ;
         for j:=0 to ListG.Count-1 do
             BEGIN
             //QDel.Close ;
             MoveCur(False) ;
             if ListG.Strings[j]<>'' then
                EffaceRub(CbRubG.Items[CbRubG.Values.IndexOf(ListG.Strings[j])]) ;
             END ;
         END ;
     2 : BEGIN
         for j:=0 to ListS.Count-1 do
             BEGIN
             //QDel.Close ;
             MoveCur(False) ;
             if ListS.Strings[j]<>'' then
                EffaceRub(CbRubS.Items[CbRubS.Values.IndexOf(ListS.Strings[j])]) ;
             END ;
         END ;
     3 : BEGIN
         for j:=0 to ListG.Count-1 do
            BEGIN
            MoveCur(False) ;
            for k:=0 to ListS.Count-1 do
                BEGIN
                //QDel.Close ;
                if (ListG.Strings[j]<>'') And (ListS.Strings[k]<>'') then
                    BEGIN
                    St:=FabriqueCompte(CbRubG.Items[CbRubG.Values.IndexOf(ListG.Strings[j])],CbRubS.Items[CbRubS.Values.IndexOf(ListS.Strings[k])]) ;
                    EffaceRub(St) ;
                    END ;
                END ;
            END ;
         END ;
     4 : BEGIN
         for j:=0 to ListS.Count-1 do
            BEGIN
            MoveCur(False) ;
            for k:=0 to ListG.Count-1 do
                BEGIN
                //QDel.Close ;
                if (ListS.Strings[j]<>'') And (ListG.Strings[k]<>'') then
                    BEGIN
                    St:=FabriqueCompte(CbRubS.Items[CbRubS.Values.IndexOf(ListS.Strings[j])],CbRubG.Items[CbRubG.Values.IndexOf(ListG.Strings[k])]) ;
                    EffaceRub(St) ;
                    END ;
                END ;
            END ;
         END ;
  End ;
END ;

Function TFCreRubBu.ExisteCroise ( Bud,Sect : String ) : boolean ;
Var Q : TQuery ;
BEGIN
Result:=False ;
if RGChoix.ItemIndex=1 then
   BEGIN
   if OnCategorie and (LaCat<>'') then
      Q:=OpenSQL('SELECT Count(*) from CROISCPT Where CX_TYPE="BUD" AND CX_COMPTE="'+Bud+'" AND CX_SECTION="'+Sect+'" AND CX_JAL="'+LaCat+'"',True)
   else
      Q:=OpenSQL('SELECT Count(*) from CROISCPT Where CX_TYPE="BUD" AND CX_COMPTE="'+Bud+'" AND CX_SECTION="'+Sect+'" AND CX_JAL="'+CBJal.Value+'"',True) ;
   if Not Q.EOF then if Q.Fields[0].AsInteger>0 then Result:=True ;
   Ferme(Q) ;
   END else
   BEGIN
   if TrouveEcrBud(Bud,Sect) then Result:=True ;
   END ;
END ;

Procedure TFCreRubBu.InsereLesEnreg(i : Integer ; TOBG,TOBS : TOB ; Var TOBRUB : TOB) ;
Var j,k : Integer ;
    TOBLG,TOBLS : TOB ;
BEGIN
(*
SqlG:='Select * From BUDGENE Where BG_BUDGENE=:CpteG' ;
SqlS:='Select * From BUDSECT Where BS_AXE=:Ax And BS_BUDSECT=:CpteS' ;
*)
MsgPresence:=False ; CodeVide:=False ;
Case i of
     1 : BEGIN
         for j:=0 to ListG.Count-1 do
            BEGIN
            //QCpteG.Close ;
            MoveCur(False) ;
            if ListG.Strings[j]<>'' then
               BEGIN
               (*
               QCpteG.ParamByName('CpteG').AsString:=ListG.Strings[j] ;
               QCpteG.Open ;
               *)

               TobLG:=TOBG.FindFirst(['BG_BUDGENE'],[ListG.Strings[j]],False) ;
               If TOBLG<>NIL THen EcritLenreg(TOBLG,Nil,TOBRUB,'') ;
               END ;
            END ;
         END ;
     2 : BEGIN
         for j:=0 to ListS.Count-1 do
            BEGIN
            //QCpteS.Close ;
            MoveCur(False) ;
            if ListS.Strings[j]<>'' then
               BEGIN
               (*
               QCpteS.ParamByName('CpteS').AsString:=ListS.Strings[j] ;
               QCpteS.ParamByName('Ax').AsString:=Laxe ;
               QCpteS.Open ;
               EcritLenreg(QCpteS,Nil,Laxe) ;
               *)
               TobLS:=TOBS.FindFirst(['BS_BUDSECT','BS_AXE'],[ListS.Strings[j],LAxe],False) ;
               If TOBLS<>NIL THen EcritLenreg(TOBLS,Nil,TOBRUB,Laxe) ;
               END ;
            END ;
         END ;
     3 : BEGIN
//         l:=0 ;
         for j:=0 to ListG.Count-1 do
            BEGIN
            MoveCur(False) ;
            for k:=0 to ListS.Count-1 do
                BEGIN
//                if (j*k)-l>2000 then BEGIN l:=j*k ; CommitTrans ; BeginTrans ; END ;
                if (ListG.Strings[j]<>'') And (ListS.Strings[k]<>'') then
                    BEGIN
                    if ExisteCroise(ListG.Strings[j],ListS.Strings[k]) then
                       BEGIN
                       (*
                       QCpteG.Close ; QCpteS.Close ;
                       QCpteG.ParamByName('CpteG').AsString:=ListG.Strings[j] ;
                       QCpteS.ParamByName('CpteS').AsString:=ListS.Strings[k] ;
                       QCpteS.ParamByName('Ax').AsString:=Laxe ;
                       QCpteG.Open ; QCpteS.Open ;
                       *)
                       TobLG:=TOBG.FindFirst(['BG_BUDGENE'],[ListG.Strings[j]],False) ;
                       TobLS:=TOBS.FindFirst(['BS_BUDSECT','BS_AXE'],[ListS.Strings[k],LAxe],False) ;
                       If (TOBLG<>NIL) And (TOBLS<>NIL) Then EcritLenreg(TOBLG,TOBLS,TOBRUB,Laxe) ;
                       END ;
                    END ;
                END ;
            END ;
         END ;
     4 : BEGIN
//         l:=0 ;
         for j:=0 to ListS.Count-1 do
            BEGIN
            MoveCur(False) ;
            for k:=0 to ListG.Count-1 do
                BEGIN
//                if (j*k)-l>2000 then BEGIN l:=j*k ; CommitTrans ; BeginTrans ; END ;
                if (ListS.Strings[j]<>'') And (ListG.Strings[k]<>'') then
                    BEGIN
                    if ExisteCroise(ListG.Strings[k],ListS.Strings[j]) then
                       BEGIN
                       (*
                       QCpteG.Close ; QCpteS.Close ;
                       QCpteG.ParamByName('CpteG').AsString:=ListG.Strings[k] ;
                       QCpteS.ParamByName('CpteS').AsString:=ListS.Strings[j] ;
                       QCpteS.ParamByName('Ax').AsString:=Laxe ;
                       QCpteG.Open ; QCpteS.Open ;
                       *)
                       TobLG:=TOBG.FindFirst(['BG_BUDGENE'],[ListG.Strings[k]],False) ;
                       TobLS:=TOBS.FindFirst(['BS_BUDSECT','BS_AXE'],[ListS.Strings[j],LAxe],False) ;
                       If (TOBLG<>NIL) And (TOBLS<>NIL) Then EcritLenreg(TOBLS,TOBLG,TOBRUB,Laxe) ;
                       END ;
                    END ;
                END ;
            END ;
         END ;
   End ;
if MsgPresence then HM.Execute(3,'','') ;
if CodeVide then HM.Execute(4,'','') ;
END ;

Function TFCreRubBu.TrouveEcrBud(StG,StS : String) : Boolean ;
Var Q1 : TQuery ;
BEGIN
(*
QEcr.Close ;
QEcr.ParamByName('EJAL').AsString:=CbJal.Value ;
QEcr.ParamByName('BGEN').AsString:=StG ;
QEcr.ParamByName('BSEC').AsString:=StS ;
QEcr.Open ;
Result:=Not QEcr.Eof ;
*)
Q1:=OPENSQL('Select BE_BUDJAL From BUDECR Where BE_BUDJAL="'+CbJal.Value+'" And BE_BUDGENE="'+StG+'" And BE_BUDSECT="'+StS+'" ',TRUE) ;
Result:=Not Q1.Eof ;
Ferme(Q1) ;
END ;

Procedure TFCreRubBu.EcritLenreg(TOBLG,TOBLS : TOB ; Var TOBRUB : TOB ; Axe : String ) ;
Var CpteRu,TypRub,Pre1,Pre2,Champ1,Champ2,Cpt1,
    Cpt2,ChaineCpte,Cle1,Cle2,CPtAtt1,CPtAtt2,LeLib,CpteRuAbr : String ;
    TOBLR : TOB ;
BEGIN
if (TOBLG=Nil) And (TOBLS=Nil) then Exit ;
If TOBLG<>NIL Then
  BEGIN
  if TOBLG.NomTable='BUDGENE' then
     BEGIN
     Pre1:='BG_' ; Champ1:='BG_RUB' ; TypRub:='GEN' ; Cpt1:='BG_COMPTERUB' ;
     Cle1:='BG_BUDGENE' ; CptAtt1:=CptGenAtt ;
     END else
     BEGIN
     Pre1:='BS_' ; Champ1:='BS_RUB' ; TypRub:='ANA' ; Cpt1:='BS_SECTIONRUB' ;
     Cle1:='BS_BUDSECT' ; CPtAtt1:=CptSecAtt ;
     END ;
  END ;
If TOBLS<>NIL Then
  BEGIN
  if TOBLS.NomTable='BUDGENE' then
    BEGIN
    Pre2:='BG_' ; Champ2:='BG_RUB' ; TypRub:='A/G' ; Cpt2:='BG_COMPTERUB' ;
    Cle2:='BG_BUDGENE' ; CptAtt2:=CptGenAtt ;
    END else
    BEGIN
    Pre2:='BS_' ; Champ2:='BS_RUB' ; TypRub:='G/A' ; Cpt2:='BS_SECTIONRUB' ;
    Cle2:='BS_BUDSECT' ; CptAtt2:=CptSecAtt ;
    END ;
  END ;
(*
if Q1.FindField(Cle1).AsString<>CptAtt1 then
  if Pos(';',Q1.FindField(CPt1).AsString)=1 then Exit ;
*)
if TOBLG.GetValue(Cle1)<>CptAtt1 then
  if Pos(';',TOBLG.GetValue(CPt1))=1 then Exit ;

if TOBLS<>Nil then
   if TOBLS.GetValue(Cle2)<>CptAtt2 then
      if Pos(';',TOBLS.GetValue(CPt2))=1 then Exit ;
if TOBLG.GetValue(Champ1)='' then BEGIN CodeVide:=True ; Exit ; END ;
if TOBLS<>Nil then if TOBLS.GetValue(Champ2)='' then BEGIN CodeVide:=True ; Exit ; END ;

if TOBLS<>Nil then
   BEGIN
   CpteRu:=Pref+CbJal.Value+FabriqueCompte(TOBLG.GetValue(Champ1),TOBLS.GetValue(Champ2)) ;
   CpteRuAbr:=FabriqueCompteAbr(TOBLG.GetValue(Champ1),TOBLS.GetValue(Champ2)) ;
   LeLib:=TOBLG.GetValue(Pre1+'LIBELLE')+TOBLS.GetValue(Pre2+'LIBELLE') ;
   END else
   BEGIN
   CpteRu:=Pref+CbJal.Value+FabriqueCompte(TOBLG.GetValue(Champ1),'') ;
   CpteRuAbr:=FabriqueCompteAbr(TOBLG.GetValue(Champ1),'') ;
   LeLib:=TOBLG.GetValue(Pre1+'LIBELLE') ;
   END ;
if Length(LeLib)>35 then LeLib:=Copy(LeLib,1,35) ;
if Presence('RUBRIQUE','RB_RUBRIQUE',CpteRu) then
   BEGIN MsgPresence:=True ; Exit ; END ;
if TOBLG.GetValue(Cle1)<>CptAtt1 then ChaineCpte:=FabriqueCompteRub(TOBLG.GetValue(Cpt1),Pre1)
                                 else ChaineCpte:=FabriqueCompteRubAtt(Pre1) ;

TobLR := TOB.Create('RUBRIQUE',TOBRUB,-1); TobLR.InitValeurs ;
TOBLR.PutValue('RB_RUBRIQUE',CpteRu) ; TOBLR.PutValue('RB_CODEABREGE',CpteRuAbr) ; TOBLR.PutValue('RB_LIBELLE',LELIB) ;
TOBLR.PutValue('RB_FAMILLES',Pref+';') ;
if NatJal='CHA' then TOBLR.PutValue('RB_SIGNERUB','POS') else TOBLR.PutValue('RB_SIGNERUB.AsString','NEG') ;
TOBLR.PutValue('RB_TYPERUB',TypRub) ; TOBLR.PutValue('RB_COMPTE1',ChaineCpte) ; TOBLR.PutValue('RB_EXCLUSION1',TOBLG.GetValue(Pre1+'EXCLURUB')) ;
if TOBLS<>Nil then
   BEGIN
   if TOBLS.GetValue(Cle2)<>CptAtt2 then ChaineCpte:=FabriqueCompteRub(TOBLS.GetValue(Cpt2),Pre2)
                                    else ChaineCpte:=FabriqueCompteRubAtt(Pre2) ;
   TOBLR.PutValue('RB_COMPTE2',ChaineCpte) ;
   TOBLR.PutValue('RB_EXCLUSION2',TOBLS.GetValue(Pre2+'EXCLURUB')) ;
   END ;
TOBLR.PutValue('RB_AXE',Axe) ;
TOBLR.PutValue('RB_NATRUB',TypeCreation) ;
TOBLR.PutValue('RB_BUDJAL',CbJal.Value) ;

TOBLR.PutValue('RB_PREDEFINI','DOS') ;
{JP 12/10/07 : FQ 21561 : il ne faut pas initialiser les rubriques à iDate1900}
TOBLR.PutValue('RB_DATEVALIDITE',iDate2099) ;
TOBLR.PutValue('RB_NODOSSIER',V_PGI.NoDossier) ;
EcritTobRUB(TOBRUB,FALSE) ;

(*
QRub.Insert ; InitNew(QRub) ;
QRubRB_RUBRIQUE.AsString:=CpteRu ;
QRubRB_CODEABREGE.AsString:=CpteRuAbr ;
QRubRB_LIBELLE.AsString:=LeLib ;
QRubRB_FAMILLES.AsString:=Pref+';' ;

// JLD + GP Pour XX_PREDEFINI
QRubRB_PREDEFINI.AsString:='DOS' ; QRubRB_NODOSSIER.AsString:='000000' ;
if V_PGI_Env<>Nil then if V_PGI_Env.ModeFonc<>'MONO' then QRubRB_NODOSSIER.AsString:=V_PGI_ENV.NoDossier ;

if NatJal='CHA' then QRubRB_SIGNERUB.AsString:='POS' else QRubRB_SIGNERUB.AsString:='NEG' ;
QRubRB_TYPERUB.AsString:=TypRub ;
QRubRB_COMPTE1.AsString:=ChaineCpte ;
QRubRB_EXCLUSION1.AsString:=TOBLG.GetValue(Pre1+'EXCLURUB') ;
if TOBLS<>Nil then
   BEGIN
   if TOBLS.GetValue(Cle2)<>CptAtt2 then ChaineCpte:=FabriqueCompteRub(TOBLS.GetValue(Cpt2),Pre2)
                                    else ChaineCpte:=FabriqueCompteRubAtt(Pre2) ;
   QRubRB_COMPTE2.AsString:=ChaineCpte ;
   QRubRB_EXCLUSION2.AsString:=TOBLS.GetValue(Pre2+'EXCLURUB') ;
   END ;
QRubRB_AXE.AsString:=Axe ;
QRubRB_NATRUB.AsString:=TypeCreation ;
QRubRB_BUDJAL.AsString:=CbJal.Value ;
QRub.Post ;
*)
END ;

Function TFCreRubBu.FabriqueCompteRub(St,Prefi : String) : String ;
Const TypCalc = '(SM)' ;
Var St1,St2,StC,StSp,StSp1 : String ;
    i,Deb,Lon : Integer ;
BEGIN
St1:=St ; StC:='' ;
if (Prefi='BG_') or (Not OnCategorie) or (QuelCodeSp='') then
   BEGIN
   While St1<>'' do
      BEGIN
      St2:=ReadTokenSt(St1) ;
      if St2<>'' then StC:=StC+St2+TypCalc+';'
      END ;
   END else
   BEGIN
   While St1<>'' do
      BEGIN
      St2:=ReadTokenSt(St1) ;
      if St2<>'' then
         BEGIN
         StSp:=QuelCodeSp ;
         for i:=1 to MaxSousPlan do
            BEGIN
            if TabSousPlan[i].Code='' then Break ;
            StSp1:=ReadTokenSt(StSp) ;
            Deb:=TabSousPlan[i].Debut ; Lon:=TabSousPlan[i].Longueur ;
            Insert(StSp1,St2,Deb) ;
            Delete(St2,Deb+Lon,Lon) ;
            END ;
         END ;
      if St2<>'' then StC:=StC+St2+TypCalc+';'
      END ;
   END ;
StC:=StC+';' ; Result:=StC ;
END ;

Function TFCreRubBu.FabriqueCompteRubAtt(St : String) : String ;
Const TypCalc = '(SM)' ;
BEGIN
if St='BG_' then Result:=VH^.Cpta[fbGene].Attente+TypCalc+';'
            else Result:=VH^.Cpta[AxeToFb(Laxe)].Attente+TypCalc+';'
END ;

Procedure TFCreRubBu.InitAxeCptAttReqDel ;
Var Q1 : TQuery ;
BEGIN
(*
//Vu avec GG le 7/01/98 BJ_NATJAL
QCpteG.Close ; QCpteG.Sql.Clear ;
QCpteG.Sql.Add('Select BJ_AXE,BJ_GENEATTENTE,BJ_SECTATTENTE,BJ_NATJAL From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"') ;
ChangeSql(QCpteG) ; QCpteG.Open ;
Laxe:=QCpteG.Fields[0].AsString ; CptGenAtt:=QCpteG.Fields[1].AsString ;
CptSecAtt:=QCpteG.Fields[2].AsString ; NatJal:=QCpteG.Fields[3].AsString ;
QDel.Close ; QDel.Sql.Clear ;
QDel.Sql.Add('Delete From RUBRIQUE Where RB_RUBRIQUE=:LeCpte And RB_FAMILLES=:Fam And RB_AXE=:Ax') ;
ChangeSql(QDel) ; //QDel.Prepare ;
PrepareSQLODBC(QDel) ;
*)
Q1:=OpenSQL('Select BJ_AXE,BJ_GENEATTENTE,BJ_SECTATTENTE,BJ_NATJAL From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"',TRUE) ;
If Not Q1.Eof Then
  BEGIN
  Laxe:=Q1.Fields[0].AsString ; CptGenAtt:=Q1.Fields[1].AsString ;
  CptSecAtt:=Q1.Fields[2].AsString ; NatJal:=Q1.Fields[3].AsString ;
  END ;
Ferme(Q1) ;
(*
  QDel.Close ; QDel.Sql.Clear ;
  QDel.Sql.Add('Delete From RUBRIQUE Where RB_RUBRIQUE=:LeCpte And RB_FAMILLES=:Fam And RB_AXE=:Ax') ;
  ChangeSql(QDel) ; //QDel.Prepare ;
  PrepareSQLODBC(QDel) ;
  END ;
*)
END ;

(*
Procedure TFCreRubBu.FabriqueRequeteCroiser ;
Var SqlG,SqlS : String ;
BEGIN
InitAxeCptAttReqDel ;
QCpteG.Close ; QCpteG.Sql.Clear ; QCpteS.Close ; QCpteS.Sql.Clear ;
if OnCategorie and (LaCat<>'') then
   BEGIN
   SqlG:='Select Distinct CX_COMPTE From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+LaCat+'" '+
         'And CX_CATEGORIE="'+LaCat+'" ' ;
   SqlS:='Select Distinct CX_SECTION From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+LaCat+'" '+
         'And CX_CATEGORIE="'+LaCat+'" ' ;
   END else
   BEGIN
   SqlG:='Select Distinct CX_COMPTE From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+CbJal.Value+'" '+
         'And CX_CATEGORIE="" ' ;
   SqlS:='Select Distinct CX_SECTION From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+CbJal.Value+'" '+
         'And CX_CATEGORIE="" ' ;
   END ;
QCpteG.Sql.Add(SqlG) ; ChangeSql(QCpteG) ;
QCpteS.Sql.Add(SqlS) ; ChangeSql(QCpteS) ;
END ;
*)
(*
Procedure TFCreRubBu.FabriqueRequete ;
Var SqlG,SqlS : String ;
BEGIN
InitAxeCptAttReqDel ;
QCpteG.Close ; QCpteG.Sql.Clear ; QCpteS.Close ; QCpteS.Sql.Clear ;
SqlG:='Select BJ_BUDGENES,BJ_BUDGENES2 From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"' ;
SqlS:='Select BJ_BUDSECTS,BJ_BUDSECTS2 From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"' ;
QCpteG.Sql.Add(SqlG) ; ChangeSql(QCpteG) ;
QCpteS.Sql.Add(SqlS) ; ChangeSql(QCpteS) ;
END ;
*)
{Procedure TFCreRubBu.FabriqueRequeteCpte ;
Var SqlG,SqlS : String ;
BEGIN
(*
QCpteG.Close ; QCpteG.Sql.Clear ; QCpteS.Close ; QCpteS.Sql.Clear ;
SqlG:='Select * From BUDGENE Where BG_BUDGENE=:CpteG' ;
SqlS:='Select * From BUDSECT Where BS_AXE=:Ax And BS_BUDSECT=:CpteS' ;
QCpteG.Sql.Add(SqlG) ; ChangeSql(QCpteG) ; //QCpteG.Prepare ;
PrepareSQLODBC(QCpteG) ;
QCpteS.Sql.Add(SqlS) ; ChangeSql(QCpteS) ; //QCpteS.Prepare ;
PrepareSQLODBC(QCpteS) ;
*)
END ;}

Procedure TFCreRubBu.FabriqueListeCompte ;
Var St,St1,CptG,CptS : String ;
    Q1,Q2 : TQuery ;
BEGIN
(*
QCpteG.Open ; QCpteS.Open ;
if RgChoix.ItemIndex=0 then
   BEGIN
   if QCpteG.Fields[1].AsString<>'' then CptG:=QCpteG.Fields[0].AsString+QCpteG.Fields[1].AsString
                                    else CptG:=QCpteG.Fields[0].AsString ;
   if QCpteS.Fields[1].AsString<>'' then CptS:=QCpteS.Fields[0].AsString+QCpteS.Fields[1].AsString
                                    else CptS:=QCpteS.Fields[0].AsString ;
   END ;
Case RgChoix.ItemIndex of
     0 : BEGIN
         if Not QCpteG.Eof then
            BEGIN
            St:=Trim(CptG) ;
            While St<>'' do BEGIN St1:=ReadTokenSt(St) ; if St1<>'' then ListG.Add(St1) ; END ;
            END ;
         if Not QCpteS.Eof then
            BEGIN
            St:=Trim(CptS) ;
            While St<>'' do BEGIN St1:=ReadTokenSt(St) ; if St1<>'' then ListS.Add(St1) ; END ;
            END ;
         QCpteG.Close ; QCpteS.Close ;
         ListG.Add(CptGenAtt) ; ListS.Add(CptSecAtt) ;
         END ;
     1 : BEGIN
         While Not QCpteG.Eof do BEGIN ListG.Add(QCpteG.Fields[0].AsString) ; QCpteG.Next ; END ;
         While Not QCpteS.Eof do BEGIN ListS.Add(QCpteS.Fields[0].AsString) ; QCpteS.Next ; END ;
         QCpteG.Close ; QCpteS.Close ;
         END ;
   END ;
*)
InitAxeCptAttReqDel ;
if RgChoix.ItemIndex=0 then
  BEGIN
  Q1:=OpenSQL('Select BJ_BUDGENES,BJ_BUDGENES2 From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"',TRUE) ;
  Q2:=OpenSQL('Select BJ_BUDSECTS,BJ_BUDSECTS2 From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"',TRUE) ;
  END Else
  BEGIN
  if OnCategorie and (LaCat<>'') then
     BEGIN
     Q1:=OpenSQL('Select Distinct CX_COMPTE From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+LaCat+'" '+
                 'And CX_CATEGORIE="'+LaCat+'" ',TRUE) ;
     Q2:=OpenSQL('Select Distinct CX_SECTION From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+LaCat+'" '+
                 'And CX_CATEGORIE="'+LaCat+'" ',TRUE) ;
     END else
     BEGIN
     Q1:=OpenSQL('Select Distinct CX_COMPTE From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+CbJal.Value+'" '+
                 'And CX_CATEGORIE="" ',TRUE) ;
     Q2:=OpenSQL('Select Distinct CX_SECTION From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+CbJal.Value+'" '+
               'And CX_CATEGORIE="" ',TRUE) ;
     END ;
  END ;
if RgChoix.ItemIndex=0 then
   BEGIN
   if Q1.Fields[1].AsString<>'' then CptG:=Q1.Fields[0].AsString+Q1.Fields[1].AsString
                                else CptG:=Q1.Fields[0].AsString ;
   if Q2.Fields[1].AsString<>'' then CptS:=Q2.Fields[0].AsString+Q2.Fields[1].AsString
                                    else CptS:=Q2.Fields[0].AsString ;
   END ;
Case RgChoix.ItemIndex of
     0 : BEGIN
         if Not Q1.Eof then
            BEGIN
            St:=Trim(CptG) ;
            While St<>'' do BEGIN St1:=ReadTokenSt(St) ; if St1<>'' then ListG.Add(St1) ; END ;
            END ;
         if Not Q2.Eof then
            BEGIN
            St:=Trim(CptS) ;
            While St<>'' do BEGIN St1:=ReadTokenSt(St) ; if St1<>'' then ListS.Add(St1) ; END ;
            END ;
//         Q1.Close ; Q2.Close ;
         Ferme(Q1) ; Ferme(Q2) ;
         ListG.Add(CptGenAtt) ; ListS.Add(CptSecAtt) ;
         END ;
     1 : BEGIN
         While Not Q1.Eof do BEGIN ListG.Add(Q1.Fields[0].AsString) ; Q1.Next ; END ;
         While Not Q2.Eof do BEGIN ListS.Add(Q2.Fields[0].AsString) ; Q2.Next ; END ;
//         Q1.Close ; Q2.Close ;
         Ferme(Q1) ; Ferme(Q2) ;
         END ;
   END ;
END ;

procedure TFCreRubBu.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCreRubBu.FormCreate(Sender: TObject);
begin PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; end;

procedure TFCreRubBu.CbJalChange(Sender: TObject);
Var QLoc : TQuery ;
begin
FillChar(TabSousPlan,SizeOf(TabSousPlan),#0) ; OnCategorie:=False ; QuelCodeSp:='' ;
QLoc:=OpenSql('Select BJ_CATEGORIE,BJ_SOUSPLAN From BUDJAL Where BJ_BUDJAL="'+CbJal.Value+'"',True) ;
if QLoc.Fields[0].AsString<>'' then
   BEGIN
   TabSousPlan:=SousPlanCat(QLoc.Fields[0].AsString,True) ; OnCategorie:=True ;
   LaCat:=QLoc.Fields[0].AsString ; QuelCodeSp:=QLoc.Fields[1].AsString ;
   END ;
Ferme(QLoc) ;
end;

procedure TFCreRubBu.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Parent is THPanel then Action:=caFree;
end;

end.
