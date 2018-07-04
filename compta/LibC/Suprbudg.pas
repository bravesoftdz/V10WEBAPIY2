{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par SUPPRBUDG_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Suprbudg;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB, DBTables, Hqry, Grids, DBGrids, StdCtrls, Hctrls,
  ExtCtrls, ComCtrls, Buttons, Mask, RapSuppr, Budgene, hmsgbox, HEnt1, Ent1,
  HRichEdt, Hcompte, HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil,
  HRichOLE;

Procedure SuppressionCpteBudG ;

type
  TFSuprbudg = class(TFMul)
    TBG_BUDGENE: THLabel;
    TBG_LIBELLE: THLabel;
    BG_LIBELLE: TEdit;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    BG_BUDGENE: THCpteEdit;
    BZgene: TToolbarButton97;
    CbAtt: TComboBox;
    TBG_DATECREATION: THLabel;
    BG_DATECREATION: THCritMaskEdit;
    BG_DATECREATION_: THCritMaskEdit;
    TBG_DATEMODIF: THLabel;
    BG_DATEMODIF: THCritMaskEdit;
    BG_DATEMODIF_: THCritMaskEdit;
    TBG_DATEMODIF_: THLabel;
    TBG_DATECREATION_: THLabel;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject); override;
  private
    Nblig     : Integer ;
    TDelBudg  : TList ;
    TNotDel   : TList ;
    Effacer   : Boolean ;
    NotEffacer : Boolean ;
    BudgCode  : String ;
    Function  Detruit(St : String): Byte ;
    Procedure Degage ;
    Procedure RempliCbAtt ;
    Procedure MajListeCompte(St : String) ;
    Function  CoupeA2000(St : String) : String ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure SuppressionCpteBudG ;
var FSuprbudg : TFSuprbudg ;
    PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(False) then Exit ;
FSuprbudg:=TFSuprbudg.Create(Application) ;
FSuprbudg.FNomFiltre:='SUPPRBUDG' ;
FSuprbudg.Q.Manuel:=True ;
FSuprbudg.Q.Liste:='SUPPRBUDG' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
    FSuprbudg.ShowModal ;
   finally
    FSuprbudg.Free ;
    _DeblocageMonoPoste(False) ;
   end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FSuprbudg,PP) ;
   FSuprbudg.Show ;
   END ;
end ;


procedure TFSuprbudg.FormShow(Sender: TObject);
begin
BG_DATECREATION.Text:=StDate1900 ; BG_DATECREATION_.Text:=StDate2099 ;
BG_DATEMODIF.Text:=StDate1900 ; BG_DATEMODIF_.Text:=StDate2099 ; RempliCbAtt ;
Pages.ActivePage:=Pages.Pages[0] ; PComplement.TabVisible:=False ;
  inherited;
Q.Manuel:=FALSE ;
end;

procedure TFSuprbudg.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X : DelInfo ;
    Code,Lib : String ;
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
       Code:=Q.FindField('BG_BUDGENE').AsString ;
       Lib:=Q.FindField('BG_LIBELLE').AsString ;
       j:=Detruit(Code) ;
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
   Code:=Q.FindField('BG_BUDGENE').AsString ; j:=Detruit(Code) ;
   if j=6 then MsgDel.Execute(1,'','') ;
   if j=10 then MsgDel.Execute(2,'','') ;
   END ;
if Effacer    then if HM.Execute(3,'','')=mrYes then RapportDeSuppression(TDelBudg,1) ;
if NotEffacer then if HM.Execute(4,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;

procedure TFSuprbudg.Degage ;
BEGIN
if ExecuteSQL('DELETE FROM BUDGENE WHERE BG_BUDGENE="'+BudgCode+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

Function TFSuprbudg.Detruit(St : String):Byte ;
BEGIN
Result:=0 ;
if CbAtt.Items.IndexOf(St)<>-1 then BEGIN Result:=10 ; Exit ; END ;
if EstMouvementeBudgen(St)then
   BEGIN
   if MsgDel.Execute(0,'','')<>mrYes then BEGIN Result:=1 ; Exit ; END ;
   END ;
BudgCode:=St ;
if Transactions(Degage,5)<>oeOK then
   BEGIN MessageAlerte(HM.Mess[6]) ; Result:=6 ; Exit ; END ;
ExecuteSQL('DELETE FROM BUDECR WHERE BE_BUDGENE="'+BudgCode+'"') ;
ExecuteSql('Delete From CROISCPT Where CX_COMPTE="'+BudgCode+'" And CX_TYPE="BUD"') ;
MajListeCompte(BudgCode) ;
END ;

Function TFSuprbudg.CoupeA2000(St : String) : String ;
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

Procedure TFSuprbudg.MajListeCompte(St : String) ;
Var QLoc : TQuery ;
    StC,St1,StTemp : String ;
    Trouver : Boolean ;
BEGIN
QLoc:=OpenSql('Select BJ_BUDGENES,BJ_BUDGENES2 From BUDJAL',False) ;
While Not QLoc.Eof do
   BEGIN
   Trouver:=False ; St1:='' ;
   if QLoc.Fields[1].AsString<>'' then StC:=QLoc.Fields[0].AsString+QLoc.Fields[1].AsString
                                  else StC:=QLoc.Fields[0].AsString ;
   While StC<>'' do
      BEGIN
      StTemp:=ReadTokenSt(StC) ;
      if StTemp<>' ' then //fb 28/09/2005 FQ 16043
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

procedure TFSuprbudg.FormCreate(Sender: TObject);
begin
  inherited;
TDelBudg:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFSuprbudg.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TDelBudg.Clear ; TDelBudg.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(True) ;
end;

procedure TFSuprbudg.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
FicheBudgene(Q,'',Q.FindField('BG_BUDGENE').AsString,taConsult,0) ;
end;

Procedure TFSuprbudg.RempliCbAtt ;
Var QLoc : TQuery ;
BEGIN
CbAtt.Items.Clear ;
QLoc:=OpenSql('Select BG_BUDGENE From BUDGENE Where BG_ATTENTE="X"',True) ;
While Not QLoc.Eof Do
   BEGIN
   CbAtt.Items.Add(QLoc.Fields[0].AsString) ; QLoc.Next ;
   END ;
Ferme(QLoc) ;
END ;

end.
