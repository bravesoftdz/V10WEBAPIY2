{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 20/02/2003
Modifié le ... :   /  /    
Description .. : Remplace en eAGL par UTOFMULPARAMGEN.PAS
Mots clefs ... :
*****************************************************************}
unit Suprsect;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, Mask, Hctrls, StdCtrls, Menus, DB, Hqry, Grids,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBGrids, ExtCtrls, ComCtrls, Buttons, HEnt1, Ent1, CPSection_TOM, HStatus, RapSuppr,
  HCompte, HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HPanel,UiUtil, HRichOLE,
  ADODB ;

Procedure SuppressionSection ;

type
  TFSuprsect = class(TFMul)
    TG_GENERAL: THLabel;
    TS_AXE: THLabel;
    S_AXE: THValComboBox;
    TG_ABREGE: THLabel;
    S_LIBELLE: TEdit;
    TG_SENS: THLabel;
    S_SENS: THValComboBox;
    HLabel2: THLabel;
    S_DATECREATION: THCritMaskEdit;
    HLabel3: THLabel;
    S_DATECREATION_: THCritMaskEdit;
    TG_DATEDERNMVT: THLabel;
    TG_DATEMODIFICATION: THLabel;
    S_DATEMODIF: THCritMaskEdit;
    TG_DATEMODIFICATION2: THLabel;
    S_DATEMODIF_: THCritMaskEdit;
    S_DATEDERNMVT_: THCritMaskEdit;
    TG_DATEDERNMVT2: THLabel;
    S_DATEDERNMVT: THCritMaskEdit;
    BZgene: TToolbarButton97;
    HM: THMsgBox;
    MsgDel: THMsgBox;
    S_SECTION: THCpteEdit;
    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeDblClick(Sender: TObject); override;
    procedure S_AXEChange(Sender: TObject);
  private    { Déclarations privées }
    Nblig     : Integer ;
    TDelSect  : TList ;
    TNotDel   : TList ;
    Effacer   : Boolean ;
    NotEffacer: Boolean ;
    SectCode  : String ;
    Codax     : String ;
    Function  Detruit(Stc,Sta : String): Byte ;
    Procedure Degage ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

Procedure SuppressionSection ;
var FSuprsect : TFSuprsect;
    PP       : THPanel ;
begin
if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
FSuprsect:=TFSuprsect.Create(Application) ;
FSuprSect.FNomFiltre:='SUPPRSECT' ; FSuprsect.Q.Manuel:=True ; FSuprSect.Q.Liste:='SUPPRSECT' ;
PP:=FindInsidePanel ;
if PP=Nil then
  BEGIN
   try
    FSuprsect.ShowModal ;
   finally
    FSuprsect.Free ;
    _DeblocageMonoPoste(False,'',TRUE) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(FSuprsect,PP) ;
  FSuprsect.Show ;
  END ;
end ;

procedure TFSuprsect.FormShow(Sender: TObject);
begin
S_DATECREATION.Text:=StDate1900 ; S_DATECREATION_.Text:=StDate2099 ;
S_DATEMODIF.Text:=StDate1900 ; S_DATEMODIF_.Text:=StDate2099 ;
S_DATEDERNMVT.Text:=StDate1900 ; S_DATEDERNMVT_.Text:=StDate2099 ;
S_AXE.ItemIndex:=0 ; S_SENS.ItemIndex:=0 ;
  inherited;
Q.Manuel:=FALSE ;
{$IFDEF CCS3}
S_AXE.Visible:=False ; TS_AXE.Visible:=False ; 
{$ENDIF}
end;

procedure TFSuprsect.BOuvrirClick(Sender: TObject);
Var i : Integer ;
    j : Byte ;
    X,Y : DelInfo ;
    Code,Lib : String ;
begin
//  inherited;
NbLig:=Fliste.NbSelected ; TDelSect.Clear ; TNotDel.Clear ;
if NbLig<=0 then BEGIN MsgDel.Execute(0,'','') ; Exit ; END ;
if MsgDel.Execute(6,'','')<>mrYes then Exit ;
Effacer:=False ; NotEffacer:=False ;
if NbLig>1 then
   BEGIN
   for i:=0 to NbLig-1 do
       BEGIN
       Fliste.GotoLeBookMark(i) ;
       Code:=Q.FindField('S_SECTION').AsString ;
       Lib:=Q.FindField('S_LIBELLE').AsString ;
       j:=Detruit(Code,S_AXE.Value) ;
       if j<=0 then
          BEGIN
          X:=DelInfo.Create ; X.LeCod:=Code ; X.LeLib:=Lib ; X.LeMess:=HM.Mess[0] ;
          TDelSect.Add(X) ; Effacer:=True ;
          END
       else
          BEGIN
          Y:=DelInfo.Create ; Y.LeCod:=Code ; Y.LeLib:=Lib ; Y.LeMess:=HM.Mess[j] ;
          TNotDel.Add(Y) ;  NotEffacer:=True ;
          END
      END ;
   END else
   BEGIN
   Fliste.GotoLeBookMark(0) ;
   Code:=Q.FindField('S_SECTION').AsString ; j:=Detruit(Code,S_AXE.Value) ;
   if j in [1..5] then MsgDel.Execute(j,'','') else
      if j=7 then MsgDel.Execute(9,'','') ;
   END ;
if Effacer    then if MsgDel.Execute(7,'','')=mrYes then RapportDeSuppression(TDelSect,1) ;
if NotEffacer then if MsgDel.Execute(8,'','')=mrYes then RapportDeSuppression(TNotDel,1) ;
BChercheClick(Nil) ;
end;

Function TFSuprsect.Detruit(Stc,Sta : String):Byte ;
BEGIN
Result:=0 ;
if EstDansAnalytiq(Stc,Sta)then BEGIN Result:=1 ; Exit ; END ;
if EstDansAxe(Stc,Sta)then BEGIN Result:=2 ; Exit ; END ;
if EstCorresp(Stc,Sta)then BEGIN Result:=3 ; Exit ; END ;
if EstDansVentil(Stc,Sta)then BEGIN Result:=4 ; Exit ; END ;
//if EstModele(Stc,Sta)then BEGIN Result:=5 ; Exit ; END ;
SectCode:=Stc ; Codax:=Sta ;
if Transactions(Degage,5)<>oeOK then
   BEGIN MessageAlerte(HM.Mess[6]) ; Result:=6 ; Exit ; END ;
END ;

procedure TFSuprsect.Degage ;
BEGIN
if ExecuteSQL('DELETE FROM SECTION WHERE S_SECTION="'+SectCode+'" AND S_AXE="'+Codax+'"')<>1 then V_PGI.IoError:=oeUnknown ;
END ;

procedure TFSuprsect.FormCreate(Sender: TObject);
begin
  inherited;
TDelSect:=TList.Create ; TNotDel:=TList.Create ;
end;

procedure TFSuprsect.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
TDelSect.Clear ; TDelSect.Free ; TNotDel.Clear ; TNotDel.Free ;
if IsInside(Self) then _DeblocageMonoPoste(False,'',TRUE) ;
end;

procedure TFSuprsect.FListeDblClick(Sender: TObject);
begin
  inherited;
if(Q.Eof) And (Q.Bof) then Exit ;
FicheSection(Q,S_AXE.Value,Q.FindField('S_SECTION').AsString,taConsult,0) ;
end;

procedure TFSuprsect.S_AXEChange(Sender: TObject);
Var St : String ;
begin
  inherited;
St:=S_AXE.Value ;
Case St[2] of
     '1': S_SECTION.ZoomTable:=tzSection  ;
     '2': S_SECTION.ZoomTable:=tzSection2 ;
     '3': S_SECTION.ZoomTable:=tzSection3 ;
     '4': S_SECTION.ZoomTable:=tzSection4 ;
     '5': S_SECTION.ZoomTable:=tzSection5 ;
 End ;
end;

end.
