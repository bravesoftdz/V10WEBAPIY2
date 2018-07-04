{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par SUPPRBUDS_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Suprbuds;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB, DBTables, Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  ExtCtrls, ComCtrls, Buttons, Mask, RapSuppr, Budsect, hmsgbox, HEnt1, Ent1,
  HRichEdt, Hcompte, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE ;

Procedure SuppressionCpteBudS ;

type
  TFSuprbuds = class(TFMul)
    TBS_BUDSECT: THLabel;
    TBS_LIBELLE: THLabel;
    BS_LIBELLE: TEdit;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    BS_BUDSECT: THCpteEdit;
    BZgene: TToolbarButton97;
    CbAtt: TComboBox;
    TBS_AXE: THLabel;
    BS_AXE: THValComboBox;
    TBS_DATECREATION: THLabel;
    BS_DATECREATION: THCritMaskEdit;
    BS_DATECREATION_: THCritMaskEdit;
    TBS_DATECREATION_: THLabel;
    TBS_DATEMODIF: THLabel;
    BS_DATEMODIF: THCritMaskEdit;
    TBS_DATEMODIF_: THLabel;
    BS_DATEMODIF_: THCritMaskEdit;
    BS_UTILISATEUR: THValComboBox;
    TBS_UTILISATEUR: THLabel;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject);
  private
    Nblig     : Integer ;
    TDelBudg  : TList ;
    TNotDel   : TList ;
    Effacer   : Boolean ;
    NotEffacer : Boolean ;
    BudgSect,BudgAxe  : String ;
    Function  Detruit(St,StAx : String) : Byte ;
    Procedure Degage ;
    Procedure RempliCbAtt ;
    Procedure MajListeCompte(St,Stax : String) ;
    Procedure MajCroisementCompte(St,Stax : String) ;
    Function  CoupeA2000(St : String) : String ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure SuppressionCpteBudS ;
var FSuprbuds : TFSuprbuds ;
    PP       : THPanel ;
BEGIN
if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
FSuprbuds:=TFSuprbuds.Create(Application) ;
FSuprbuds.FNomFiltre:='SUPPRBUDS' ; FSuprbuds.Q.Manuel:=True ; FSuprbuds.Q.Liste:='SUPPRBUDS' ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    FSuprbuds.ShowModal ;
   finally
    FSuprbuds.Free ;
    _DeblocageMonoPoste(False,'',TRUE) ;
   end ;
  SourisNormale ;
  END else
  BEGIN
  InitInside(FSuprbuds,PP) ;
  FSuprbuds.Show ;
  END ;
end ;


procedure TFSuprbuds.FormShow(Sender: TObject);
begin
BS_DATECREATION.Text:=StDate1900 ; BS_DATECREATION_.Text:=StDate2099 ;
BS_DATEMODIF.Text:=StDate1900 ; BS_DATEMODIF_.Text:=StDate2099 ; RempliCbAtt ;
BS_AXE.ItemIndex:=0 ; BS_UTILISATEUR.ItemIndex:=0 ;
Pages.ActivePage:=Pages.Pages[0] ; PComplement.TabVisible:=False ;
  inherited;
Q.Manuel:=FALSE ;
end;

procedure TFSuprbuds.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X : DelInfo ;
    Code,Ax,Lib : String ;
begin
//  inherited;
NbLig:=Fliste.NbSelected ; TDelBudg.Clear ; TNotDel.Clear ;
if NbLig<=0 then BEGIN HM.execute(2,'','') ; Exit ; END ;
if HM.Execute(0,'','')<>mrYes then Exit ;
Effacer:=False ; NotEffacer:=False ;
if NbLig>1 then
   BEGIN
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       Code:=Q.FindField('BS_BUDSECT').AsString ;
       Ax:=Q.FindField('BS_AXE').AsString ;
       Lib:=Q.FindField('BS_LIBELLE').AsString ;
       j:=Detruit(Code,Ax) ;
       if j<=0 then
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[5] ;
          TDelBudg.Add(X) ; Effacer:=True ;
          END else
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ;
          X.LeMess:=HM.Mess[j] ;
          TNotDel.Add(X) ;  NotEffacer:=True ;
          END
      END ;
   END else
   BEGIN
   Fliste.GotoLeBookMark(0) ;
   Code:=Q.FindField('BS_BUDSECT').AsString ;
   Ax:=Q.FindField('BS_AXE').AsString ;
   j:=Detruit(Code,Ax) ;
   if j=6 then MsgDel.Execute(1,'','') ;
   if j=10 then MsgDel.Execute(2,'','') ;
   END ;
if Effacer    then if HM.Execute(3,'','')=mrYes then RapportDeSuppression(TDelBudg,1) ;
if NotEffacer then if HM.Execute(4,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;

procedure TFSuprbuds.Degage ;
BEGIN
if ExecuteSQL('DELETE FROM BUDSECT WHERE BS_BUDSECT="'+BudgSect+'" And BS_AXE="'+BudgAxe+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

Function TFSuprbuds.Detruit(St,StAx : String):Byte ;
BEGIN
Result:=0 ;
if CbAtt.Items.IndexOf(St+';'+StAx)<>-1 then BEGIN Result:=10 ; Exit ; END ;
if EstMouvementeBudsect(St,StAx)then
   BEGIN
   if MsgDel.Execute(0,'','')<>mrYes then BEGIN Result:=1 ; Exit ; END ;
   END ;
BudgSect:=St ; BudgAxe:=StAx ;
if Transactions(Degage,5)<>oeOK then
   BEGIN MessageAlerte(HM.Mess[6]) ; Result:=6 ; Exit ; END ;
ExecuteSQL('DELETE FROM BUDECR WHERE BE_BUDSECT="'+BudgSect+'" And BE_AXE="'+BudgAxe+'"') ;
MajCroisementCompte(BudgSect,BudgAxe) ;
MajListeCompte(BudgSect,BudgAxe) ;
END ;

procedure TFSuprbuds.FormCreate(Sender: TObject);
begin
  inherited;
TDelBudg:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFSuprbuds.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TDelBudg.Clear ; TDelBudg.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(False,'',TRUE) ;
end;

procedure TFSuprbuds.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
FicheBudsect(Q,Q.FindField('BS_AXE').AsString,Q.FindField('BS_BUDSECT').AsString,taConsult,0) ;
end;

Procedure TFSuprbuds.RempliCbAtt ;
Var QLoc : TQuery ;
BEGIN
CbAtt.Items.Clear ;
QLoc:=OpenSql('Select BS_BUDSECT,BS_AXE From BUDSECT Where BS_ATTENTE="X"',True) ;
While Not QLoc.Eof Do
   BEGIN
   CbAtt.Items.Add(QLoc.Fields[0].AsString+';'+QLoc.Fields[1].AsString) ; QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

Procedure TFSuprbuds.MajCroisementCompte(St,Stax : String) ;
Var QLoc : TQuery ;
    StC,StTemp : String ;
BEGIN
QLoc:=OpenSql('Select BJ_BUDJAL,BJ_BUDSECTS From BUDJAL Where BJ_AXE="'+Stax+'"',True) ;
While Not QLoc.Eof do
   BEGIN
   StC:=QLoc.Fields[1].AsString ;
   While StC<>'' do
      BEGIN
      StTemp:=ReadTokenSt(StC) ;
      if StTemp=St then
         BEGIN
         ExecuteSql('Delete From CROISCPT Where CX_SECTION="'+St+'" And '+
                    'CX_JAL="'+QLoc.Fields[0].AsString+'" And CX_TYPE="BUD"') ;
         Break ;
         END ;
      END ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

Function TFSuprbuds.CoupeA2000(St : String) : String ;
Var St1,St2 : String ;
BEGIN
if Length(St)<=2000 then BEGIN Result:=St ; Exit ; END ;
St1:='' ; St2:='' ;
While St<>'' do
    BEGIN
    St1:=ReadTokenSt(St) ;
    if Length(St2+St1+';')>2000 then
       BEGIN Result:=St2+'<<>>'+St1+';'+St ; Exit ; END else
       BEGIN if St2<>'' then St2:=St2+St1+';' else St2:=St1+';' ; END ;
    END ;
END ;

Procedure TFSuprbuds.MajListeCompte(St,Stax : String) ;
Var QLoc : TQuery ;
    StC,St1,StTemp : String ;
    Trouver : Boolean ;
BEGIN
QLoc:=OpenSql('Select BJ_BUDSECTS,BJ_BUDSECTS2 From BUDJAL Where BJ_AXE="'+Stax+'"',False) ;
While Not QLoc.Eof do
   BEGIN
   Trouver:=False ; St1:='' ;
   if QLoc.Fields[1].AsString<>'' then StC:=QLoc.Fields[0].AsString+QLoc.Fields[1].AsString
                                  else StC:=QLoc.Fields[0].AsString ;
   While StC<>'' do
      BEGIN
      StTemp:=ReadTokenSt(StC) ;
      if StTemp<>' ' then //fb 28/09/2005 FQ 16046
        if StTemp=St then Trouver:=True else St1:=St1+StTemp+';' ;
      END ;
   if Trouver then
      BEGIN
      StTemp:=CoupeA2000(St1) ;
      QLoc.Edit ;
      if Pos('<<>>',StTemp)<=0 then QLoc.Fields[0].AsString:=StTemp else
         BEGIN
         QLoc.Fields[0].AsString:=Copy(StTemp,1,Pos('<<>>',StTemp)-1) ;
         QLoc.Fields[1].AsString:=Copy(StTemp,Pos('<<>>',StTemp)+1,Length(StTemp)) ;
         END ;
      QLoc.Post ;
      END ;
   QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

end.
