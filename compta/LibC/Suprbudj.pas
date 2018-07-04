{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par SUPPRBUDJ_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Suprbudj;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, DBTables, Hqry, Grids,
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, Budjal, HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE ;

Procedure SuppressionJournauxBud ;

type
  TFSupjalbu = class(TFMul)
    TBJ_BUDJAL: THLabel;
    BJ_BUDJAL: THValComboBox;
    TBJ_AXE: THLabel;
    BJ_AXE: THValComboBox;
    TBJ_LIBELLE: THLabel;
    BJ_LIBELLE: TEdit;
    TBJ_DATECREATION: THLabel;
    BJ_DATECREATION: THCritMaskEdit;
    TJ_DATECREATION_: THLabel;
    BJ_DATECREATION_: THCritMaskEdit;
    TBJ_DATEMODIF: THLabel;
    BJ_DATEMODIF: THCritMaskEdit;
    TJ_DATEMODIFICATION2: THLabel;
    BJ_DATEMODIF_: THCritMaskEdit;
    BZjal: TToolbarButton97;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject);
  private    { Déclarations privées }
    Nblig     : Integer ;
    TDelJal   : TList ;
    TNotDel   : TList ;
    Effacer   : Boolean ;
    NotEffacer: Boolean ;
    JalCode   : String ;
    FAvertir : Boolean ;
    Function  Detruit(Stc : String): Byte ;
    Procedure Degage ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure SuppressionJournauxBud ;
var Fsupjalbu : TFsupjalbu;
    PP       : THPanel ;
begin
if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
Fsupjalbu:=TFsupjalbu.Create(Application) ;
Fsupjalbu.FNomFiltre:='SUPPRBUDJ' ; Fsupjalbu.Q.Manuel:=True ; Fsupjalbu.Q.Liste:='SUPPRBUDJ' ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    Fsupjalbu.ShowModal ;
   finally
    Fsupjalbu.Free ;
    _DeblocageMonoPoste(False,'',TRUE) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(Fsupjalbu,PP) ;
  Fsupjalbu.Show ;
  END ;
end ;

procedure TFsupjalbu.FormShow(Sender: TObject);
begin
BJ_DATECREATION.Text:=StDate1900 ; BJ_DATECREATION_.Text:=StDate2099 ;
BJ_DATEMODIF.Text:=StDate1900 ; BJ_DATEMODIF_.Text:=StDate2099 ;
BJ_AXE.ItemIndex:=0 ; BJ_BUDJAL.ItemIndex:=0 ;
  inherited;
Q.Manuel:=FALSE ;
end;

procedure TFsupjalbu.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X : DelInfo ;
    Code,Lib : String ;
begin
//  inherited;
NbLig:=Fliste.NbSelected ; TDelJal.Clear ; TNotDel.Clear ;
if NbLig<=0 then BEGIN MsgDel.Execute(0,'','') ; Exit ; END ;
if MsgDel.Execute(2,'','')<>mrYes then Exit ;
Effacer:=False ; NotEffacer:=False ;
if NbLig>1 then
   BEGIN
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       Code:=Q.FindField('BJ_BUDJAL').AsString ;
       Lib:=Q.FindField('BJ_LIBELLE').AsString ;
       j:=Detruit(Code) ;
       if j<=0 then
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[0] ;
          TDelJal.Add(X) ; Effacer:=True ;
          END else
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[j] ;
          TNotDel.Add(X) ;  NotEffacer:=True ;
          END
      END ;
   END else
   BEGIN
   Fliste.GotoLeBookMark(0) ;
   Code:=Q.FindField('BJ_BUDJAL').AsString ; j:=Detruit(Code) ;
   if j=2 then MsgDel.Execute(3,'','') ;
   END ;
if Effacer    then if MsgDel.Execute(4,'','')=mrYes then RapportDeSuppression(TDelJal,1) ;
if NotEffacer then if MsgDel.Execute(5,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;

Function TFsupjalbu.Detruit(Stc : String):Byte ;
Var Qloc : TQuery ;
BEGIN
Result:=0 ;
if EstMouvementeBudjal(Stc)then
   BEGIN
   if MsgDel.Execute(1,'','')<>mrYes then BEGIN Result:=1 ; Exit ; END ;
   END ;
JalCode:=Stc ;
if Transactions(Degage,5)<>oeOK then BEGIN MessageAlerte(HM.Mess[2]) ; Result:=2 ; Exit ; END ;
ExecuteSQL('DELETE FROM BUDECR WHERE BE_BUDJAL="'+JalCode+'"') ;
QLoc:=OpenSql('Select BJ_CATEGORIE From BUDJAL Where BJ_BUDJAL="'+JalCode+'"',True) ;
if QLoc.Fields[0].AsString<>'' then JalCode:=QLoc.Fields[0].AsString ;
Ferme(QLoc) ;
ExecuteSQL('DELETE FROM CROISCPT WHERE CX_JAL="'+JalCode+'" And CX_TYPE="BUD"') ;
FAvertir:=True ;
END ;

procedure TFsupjalbu.Degage ;
BEGIN
if ExecuteSQL('DELETE FROM BUDJAL WHERE BJ_BUDJAL="'+JalCode+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFsupjalbu.FormCreate(Sender: TObject);
begin
  inherited;
TDelJal:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFsupjalbu.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
AvertirMultiTable('ttBudJal') ;
TDelJal.Clear ; TDelJal.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(False,'',TRUE) ;
end;

procedure TFsupjalbu.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
FicheBudjal(Q,'',Q.FindField('BJ_BUDJAL').AsString,taConsult,0);
end;

end.
