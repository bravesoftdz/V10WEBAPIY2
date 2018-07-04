{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 26/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par TTCATJALBUD_TOT.PAS
Mots clefs ... : 
*****************************************************************}
unit StrucBud;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBCtrls, HSysMenu, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  hmsgbox, StdCtrls, Buttons, ExtCtrls,
  Grids, DBGrids, HDB, HCtrls, HEnt1, Ent1, Hqry
  ,HPanel, UIUtil, HTB97, ADODB // MODIF PACK AVANCE pour gestion mode inside
  ;

Procedure ParamStructureBudget ;

type
  TFStrucBud = class(TForm)
    FListe: THDBGrid;
    {b FP FQ15146}
    Dock: TDock97;
    HPB: TToolWindow97;
    BSousPlanBud: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BAide: TToolbarButton97;
    BFerme: TToolbarButton97;
    BValider: TToolbarButton97;
    BImprimer: TToolbarButton97;
    {e FP FQ15146}
    HM: THMsgBox;
    Ta: THTable;
    Sa: TDataSource;
    HMTrad: THSystemMenu;
    DBNav: TDBNavigator;
    TaCC_TYPE: TStringField;
    TaCC_CODE: TStringField;
    TaCC_LIBELLE: TStringField;
    TaCC_ABREGE: TStringField;
    TaCC_LIBRE: TStringField;
    FAutoSave: TCheckBox;
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure SaDataChange(Sender: TObject; Field: TField);
    procedure SaUpdateData(Sender: TObject);
    procedure TaAfterDelete(DataSet: TDataSet);
    procedure TaAfterPost(DataSet: TDataSet);
    procedure TaNewRecord(DataSet: TDataSet);
    procedure TaPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BSousPlanBudClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeRowEnter(Sender: TObject);
  private
    Favertir : Boolean ;
    Acreer : Boolean ;
    NbEnreg : Integer ;
    Procedure NewEnreg ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  VerifiSiExiste : Boolean ;
    Procedure ChargeEnreg ;
    Function  Supprime : Boolean ;
    Function  ChercheNewCode : String ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  {$ENDIF MODENT1}
  PrintDBG, ParaSpBu ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... :   /  /    
Description .. :
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion 
Suite ........ : mode inside
Mots clefs ... : 
*****************************************************************}
Procedure ParamStructureBudget ;
var FStrucBud : TFStrucBud ;
    PP : THPanel ;
BEGIN
  FStrucBud:=TFStrucBud.Create(Application) ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FStrucBud.ShowModal ;
      Finally
      FStrucBud.Free ;
      End ;
    end
  else
    begin
    InitInside(FStrucBud,PP) ;
    FStrucBud.Show ;
    end ;

  SourisNormale ;
END ;

procedure TFStrucBud.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFStrucBud.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFStrucBud.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFStrucBud.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFStrucBud.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFStrucBud.BInsertClick(Sender: TObject);
begin
if NbEnreg>=MaxCatBud then BEGIN HM.Execute(8,Caption,'') ; Ta.Cancel ; END
                      else BEGIN Bouge(nbInsert) ; FListe.SelectedIndex:=1 ; END ;
end;

procedure TFStrucBud.BImprimerClick(Sender: TObject);
begin PrintDBGrid(FListe,Nil,Caption,'') ; end;

procedure TFStrucBud.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFStrucBud.BFermeClick(Sender: TObject);
begin
  Close ;
  {b FP FQ 15146}
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
  {e FP FQ 15146}
end;

procedure TFStrucBud.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFStrucBud.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  {FP FQ 15146 BFerme.SetFocus ;}
  CanClose:=OnSauve ;
end;

procedure TFStrucBud.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

Function TFStrucBud.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(HM.Mess[5]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFStrucBud.EnregOK : boolean ;
BEGIN
Result:=False ;
if Ta.state in [dsEdit,dsInsert]=False then Exit ;
if Ta.state in [dsEdit,dsInsert] then
   BEGIN
   if TaCC_CODE.AsString='' then
      BEGIN HM.Execute(2,caption,'') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
   if TaCC_LIBELLE.AsString='' then
      BEGIN HM.Execute(3,caption,'') ; FListe.SelectedIndex:=1 ; FListe.SetFocus ; Exit ; END ;
   if Acreer then BEGIN Acreer:=False ; Inc(NbEnreg) ; BSousPlanBudClick(Nil) ; END ;
   if TaCC_ABREGE.AsString='' then
      BEGIN HM.Execute(9,caption,'') ; BSousPlanBud.Enabled:=True ;
      {FP FQ15146 BSousPlanBud.SetFocus ;}
      Exit ;
      END ;
   END ;
if Ta.state in [dsInsert] then
   if VerifiSiExiste then
      BEGIN HM.Execute(4,caption,'') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
Result:=True ;
END ;

Function TFStrucBud.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if Ta.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=HM.execute(0,caption,'') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFStrucBud.VerifiSiExiste : Boolean ;
BEGIN
Result:=PresenceComplexe('CHOIXCOD',['CC_CODE','CC_TYPE'],['=','='],[TaCC_CODE.AsString,'CJB'],['S','S']) ;
END ;

Function TFStrucBud.Supprime : Boolean ;
BEGIN
Result:=False ;
if HM.Execute(1,Caption,'')<>mrYes then Exit ;
if HM.Execute(7,Caption,'')<>mrYes then Exit ;
ExecuteSql('Update BUDJAL Set BJ_CATEGORIE="" Where BJ_CATEGORIE="'+TaCC_CODE.AsString+'"') ;
ExecuteSql('Delete From CROISCPT Where CX_TYPE="BUD" And CX_JAL="'+TaCC_CODE.AsString+'" And '+
           'CX_CATEGORIE="'+TaCC_CODE.AsString+'"') ;
Result:=True ;
END ;

Function TFStrucBud.ChercheNewCode : String ;
Var QLoc : TQuery ;
    i : Integer ;
BEGIN
if NbEnreg=0 then BEGIN Result:='S01' ; Exit ; END ;
QLoc:=OpenSql('Select CC_CODE From CHOIXCOD Where CC_TYPE="CJB" Order by CC_CODE',True) ;
i:=1 ;
While Not QLoc.Eof do
   BEGIN
   if Copy(QLoc.Fields[0].AsString,3,1)<>IntToStr(i) then Break ;
   Inc(i) ; QLoc.Next ;
   END ;
Ferme(QLoc) ;
if i<=9 then Result:='S0'+IntToStr(i) else Result:='S'+IntToStr(i); SourisNormale ;
END ;

Procedure TFStrucBud.NewEnreg ;
BEGIN
TaCC_TYPE.AsString:='CJB' ; Acreer:=True ;
TaCC_CODE.AsString:=ChercheNewCode ;
FListe.SetFocus ;
END ;

Procedure TFStrucBud.ChargeEnreg ;
BEGIN
InitCaption(Self,TaCC_CODE.AsString,TaCC_LIBELLE.AsString) ;
if Ta.State<>dsInsert then Acreer:=False ;
END ;

procedure TFStrucBud.SaDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not(Ta.Modified) And (Not (Ta.State in [dsInsert,dsEdit])) ;
BDelete.Enabled:=Not(Ta.Modified) And (Not (Ta.State in [dsInsert,dsEdit])) ;
BSousPlanBud.Enabled:=Not(Ta.Modified) And (Not (Ta.State in [dsInsert,dsEdit])) ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not Ta.BOF;
   DnEnable := Enabled and not Ta.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   END else
   BEGIN
   ChargeEnreg ;
   END ;
end;

procedure TFStrucBud.SaUpdateData(Sender: TObject);
begin
if Acreer then BEGIN Acreer:=False ; Inc(NbEnreg) ; BSousPlanBudClick(Nil) ; END ;
if(Trim(TaCC_CODE.AsString)='') And
  (Trim(TaCC_LIBELLE.AsString)='')And
  (Trim(TaCC_ABREGE.AsString)='') then
  BEGIN Ta.Cancel ; Exit ; END ;
if(Trim(TaCC_LIBELLE.AsString)<>'') And (Trim(TaCC_ABREGE.AsString)<>'') And
  (Trim(TaCC_CODE.AsString)='') then
  BEGIN HM.Execute(2,caption,'') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TaCC_CODE.AsString)<>'') And (Trim(TaCC_ABREGE.AsString)<>'') And
  (Trim(TaCC_LIBELLE.AsString)='')then
  BEGIN HM.Execute(3,caption,'') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TaCC_LIBELLE.AsString)<>'')And(Trim(TaCC_CODE.AsString)<>'') And
  (Trim(TaCC_ABREGE.AsString)='')then
  BEGIN
  HM.Execute(9,caption,'') ; BSousPlanBud.Enabled:=True ;
  {FP FQ15146 BSousPlanBud.SetFocus ;}
  SysUtils.Abort ;
  Exit ;
  END ;
end;

procedure TFStrucBud.TaAfterDelete(DataSet: TDataSet);
begin FAvertir:=TRUE ; Dec(NbEnreg) ; end;

procedure TFStrucBud.TaAfterPost(DataSet: TDataSet);
begin FAvertir:=TRUE ; end;

procedure TFStrucBud.TaNewRecord(DataSet: TDataSet);
begin
if NbEnreg>=MaxCatBud then BEGIN HM.Execute(8,Caption,'') ; Ta.Cancel ; END else NewEnreg ;
end;

procedure TFStrucBud.TaPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
begin
if Ta.State=dsInsert then BEGIN HM.Execute(6,caption,'') ; Fliste.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

procedure TFStrucBud.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Ta.State in [dsInsert] then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift) and (Key=VK_DELETE)) then
  BEGIN
  if Not Bouge(nbDelete) then BEGIN Key:=0 ; Exit ; END ;
  Key:=0 ;
  END ;
end;

procedure TFStrucBud.FormShow(Sender: TObject);
Var QLoc : TQuery ;
begin
QLoc:=OpenSql('Select Count(*) From CHOIXCOD Where CC_TYPE="CJB"',True) ;
NbEnreg:=QLoc.Fields[0].AsInteger ; Ferme(QLoc) ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ; Acreer:=False ; Ta.Open ;
Ta.SetRange(['CJB'],['CJB']) ;
if (Ta.Eof) And (Ta.Bof) then BInsertClick(Nil) ;
end;

procedure TFStrucBud.BSousPlanBudClick(Sender: TObject);
Var Axe,SousPlan : String ;
begin
if (TaCC_CODE.AsString='') or (TaCC_LIBELLE.AsString='') then Exit ;
Axe:=TaCC_ABREGE.AsString ; SousPlan:=TaCC_LIBRE.AsString ;
ParamSousPlanBudget(TaCC_CODE.AsString,TaCC_LIBELLE.AsString,Axe,SousPlan) ;
if (Axe<>TaCC_ABREGE.AsString) or (SousPlan<>TaCC_LIBRE.AsString) then
   BEGIN
   if Ta.State=dsBrowse then Ta.Edit ;
   TaCC_ABREGE.AsString:=Axe ; TaCC_LIBRE.AsString:=SousPlan ;
   END ;
end;

procedure TFStrucBud.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Favertir then BEGIN ChargeCatBud ; AvertirMultiTable('ttCatJalBud') ; END ;
end;

procedure TFStrucBud.FListeRowEnter(Sender: TObject);
begin
if Ta.State=dsInsert then FListe.SelectedIndex:=1 ;
end;

end.
