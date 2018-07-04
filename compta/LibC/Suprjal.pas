{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 02/03/2004
Description .. :
Suite ........ : GCO - 02/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Mots clefs ... : 
*****************************************************************}
unit Suprjal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, Hqry, Grids,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1,
  CPJOURNAL_TOM,
  HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE,
  ADODB ;

Procedure SuppressionJournaux ;

type
  TFSuprjal = class(TFMul)
    TJ_JOURNAL: THLabel;
    J_JOURNAL: THValComboBox;
    TJ_AXE: THLabel;
    J_AXE: THValComboBox;
    TJ_LIBELLE: THLabel;
    J_LIBELLE: TEdit;
    TJ_NATUREJAL: THLabel;
    J_NATUREJAL: THValComboBox;
    TJ_DATECREATION: THLabel;
    J_DATECREATION: THCritMaskEdit;
    TJ_DATECREATION_: THLabel;
    J_DATECREATION_: THCritMaskEdit;
    J_DATEDERNMVT_: THCritMaskEdit;
    TJ_DATEDERNMVT2: THLabel;
    J_DATEDERNMVT: THCritMaskEdit;
    TJ_DATEDERNMVT: THLabel;
    TJ_DATEMODIFICATION: THLabel;
    J_DATEMODIF: THCritMaskEdit;
    TJ_DATEMODIFICATION2: THLabel;
    J_DATEMODIF_: THCritMaskEdit;
    BZjal: TToolbarButton97;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject); override;
    procedure QAfterDelete(DataSet: TDataSet);
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

Procedure SuppressionJournaux ;
var FSuprjal : TFSuprjal;
    PP       : THPanel ;
begin
if Not _BlocageMonoPoste(False) then Exit ;
FSuprjal:=TFSuprjal.Create(Application) ;
FSuprjal.FNomFiltre:='SUPPRJAL' ; FSuprjal.Q.Manuel:=True ; FSuprjal.Q.Liste:='SUPPRJAL' ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    FSuprjal.ShowModal ;
   finally
    FSuprjal.Free ;
    _DeblocageMonoPoste(False) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(FSuprjal,PP) ;
  FSuprjal.Show ;
  END ;
end ;

procedure TFSuprjal.FormShow(Sender: TObject);
begin
J_DATECREATION.Text:=StDate1900 ; J_DATECREATION_.Text:=StDate2099 ;
J_DATEMODIF.Text:=StDate1900 ; J_DATEMODIF_.Text:=StDate2099 ;
J_DATEDERNMVT.Text:=StDate1900 ; J_DATEDERNMVT_.Text:=StDate2099 ;
J_AXE.ItemIndex:=0 ; J_JOURNAL.ItemIndex:=0 ;
  inherited;
Q.Manuel:=FALSE ;
{$IFDEF CCS3}
J_AXE.Visible:=False ; TJ_AXE.Visible:=False ; 
{$ENDIF}
end;

procedure TFSuprjal.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X,Y : DelInfo ;
    Code,Lib : String ;
begin
//  inherited;
NbLig:=Fliste.NbSelected ; TDelJal.Clear ; TNotDel.Clear ;
if NbLig<=0 then BEGIN MsgDel.Execute(0,'','') ; Exit ; END ;
if MsgDel.Execute(7,'','')<>mrYes then Exit ;
Effacer:=False ; NotEffacer:=False ;
if NbLig>1 then
   BEGIN
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       Code:=Q.FindField('J_JOURNAL').AsString ;
       Lib:=Q.FindField('J_LIBELLE').AsString ;
       j:=Detruit(Code) ;
       if j<=0 then
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[0] ;
          TDelJal.Add(X) ; Effacer:=True ;
          END else
          BEGIN
          Y:=DelInfo.Create ; Y.LeCod:=Code ; Y.LeLib:=Lib ;
          if j in [11,12] then Y.LeMess:=MsgDel.Mess[j] else Y.LeMess:=HM.Mess[j] ;
          TNotDel.Add(Y) ;  NotEffacer:=True ;
          END
      END ;
   END else
   BEGIN
   Fliste.GotoLeBookMark(0) ;
   Code:=Q.FindField('J_JOURNAL').AsString ; j:=Detruit(Code) ;
   if j in [1..6,13..14] then MsgDel.Execute(j,'','') else
      if j=7 then MsgDel.Execute(10,'','') else
       if j=11 then MsgDel.Execute(11,'','') else
        if j=12 then MsgDel.Execute(12,'','') ;
   END ;
if Effacer    then if MsgDel.Execute(8,'','')=mrYes then RapportDeSuppression(TDelJal,1) ;
if NotEffacer then if MsgDel.Execute(9,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;

Function TFSuprjal.Detruit(Stc : String):Byte ;
BEGIN
Result:=0 ;
if EstDansEcriture(Stc)then BEGIN Result:=2 ; Exit ; END ;
if EstDansAnalytiq(Stc)then BEGIN Result:=1 ; Exit ; END ;
//if EstDansImmo(Stc)    then BEGIN Result:=3 ; Exit ; END ;
// Table IMMO Non encore Créée
if EstDansSociete(Stc) then BEGIN Result:=5 ; Exit ; END ;
if EstDansSouche(Stc)  then BEGIN Result:=6 ; Exit ; END ;
if ((Stc=VH^.JalATP) or (StC=VH^.JalVTP)) then BEGIN Result:=11 ; Exit ; END ;
if Stc=VH^.JalRepBalAN then BEGIN Result:=12 ; Exit ; END ;
{$IFDEF EAGLCLIENT}
  {$IFDEF AFAIRE}
  {$ENDIF}
{$ELSE}
// TEST présence en GESCOM ( Modif Fiche 10856 SBO )
if EstDansPiece(Stc) then BEGIN Result:=13 ; Exit ; END ;
if EstDansParamPiece(Stc) then BEGIN Result:=14 ; Exit ; END ;
if EstDansParamPieceCompl(Stc) then BEGIN Result:=14 ; Exit ; END ;
// Fin Modif Fiche 10856 SBO
{$ENDIF}
JalCode:=Stc ;
if Transactions(Degage,5)<>oeOK then
   BEGIN
   MessageAlerte(HM.Mess[7]) ;
   Result:=7 ;
   Exit ;
   END ;
END ;

procedure TFSuprjal.Degage ;
BEGIN
if ExecuteSQL('DELETE FROM JOURNAL WHERE J_JOURNAL="'+JalCode+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFSuprjal.FormCreate(Sender: TObject);
begin
  inherited;
TDelJal:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFSuprjal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
AvertirMultiTable('ttJournal') ;
TDelJal.Clear ; TDelJal.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(False) ;
end;

procedure TFSuprjal.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
FicheJournal(Q,'',Q.FindField('J_JOURNAL').AsString,taConsult,0);
end;

procedure TFSuprjal.QAfterDelete(DataSet: TDataSet);
begin
  inherited;
FAvertir:=True ;
end;

end.
