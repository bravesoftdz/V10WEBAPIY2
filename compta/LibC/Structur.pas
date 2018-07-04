{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 09/04/2003
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par CPSTRUCTURE_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit Structur;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, StdCtrls, Ent1, Grids, DBGrids, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hctrls, ComCtrls, SaisComm,
  ExtCtrls, Buttons, DBCtrls, HCompte, HEnt1, HDB, HSysMenu, Hqry, HTB97, HPanel, UiUtil, PrintDBG,
  ADODB, TntDBGrids ;

Procedure ParamPlanAnal(Faxe : String) ;

type
  TFStructure = class(TForm)
    TPlan: THTable;
    SPlan  : TDataSource;
    FListe1: THDBGrid;
    TPlanSS_AXE         : TStringField;
    TPlanSS_SOUSSECTION : TStringField;
    TPlanSS_LIBELLE     : TStringField;
    TPlanSS_CONTROLE    : TStringField;
    TPlanSS_DEBUT: TIntegerField;
    TPlanSS_LONGUEUR: TIntegerField;
    TSSPlan: THTable;
    SSSPlan : TDataSource;
    FListe2 : THDBGrid;
    TSSPlanPS_AXE         : TStringField;
    TSSPlanPS_SOUSSECTION : TStringField;
    TSSPlanPS_CODE        : TStringField;
    TSSPlanPS_LIBELLE     : TStringField;
    TSSPlanPS_ABREGE      : TStringField;
    PAxe     : TTabControl;
    Panel2   : TPanel;
    Panel3   : TPanel;
    HPB: TToolWindow97;
    Panel5   : TPanel;
    PPlan    : TPanel;
    BInsert2: TToolbarButton97;
    BDelete2: TToolbarButton97;
    BDefaire2: TToolbarButton97;
    BValider2: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide2: TToolbarButton97;
    FAutoSave : TCheckBox;
    MsgBox    : THMsgBox;
    MsgBox1   : THMsgBox;
    MsgBox2   : THMsgBox;
    BAutomate: TToolbarButton97;
    TSSPlanPS_SOCIETE: TStringField;
    TPlanSS_SOCIETE: TStringField;
    Lgaxe: TLabel;
    HMTrad: THSystemMenu;
    BMultiSect: TToolbarButton97;
    BSimpSect: TToolbarButton97;
    DBNav1: TDBNavigator;
    DBNav2: TDBNavigator;
    BDefaire1: TToolbarButton97;
    BInsert1: TToolbarButton97;
    BDelete1: TToolbarButton97;
    BValider1: TToolbarButton97;
    Dock: TDock97;
    BImprimer1: TToolbarButton97;
    BImprimer2: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure PAxeChange(Sender: TObject);
    procedure BInsert1Click(Sender: TObject);
    procedure BDelete1Click(Sender: TObject);
    procedure BDefaire1Click(Sender: TObject);
    procedure BValider1Click(Sender: TObject);
    procedure BFerme1Click(Sender: TObject);
    procedure BInsert2Click(Sender: TObject);
    procedure BDelete2Click(Sender: TObject);
    procedure BDefaire2Click(Sender: TObject);
    procedure BValider2Click(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure TPlanNewRecord(DataSet: TDataSet);
    procedure TSSPlanNewRecord(DataSet: TDataSet);
    procedure SPlanStateChange(Sender: TObject);
    procedure SPlanUpdateData(Sender: TObject);
    procedure SSSPlanStateChange(Sender: TObject);
    procedure SSSPlanUpdateData(Sender: TObject);
    procedure SPlanDataChange(Sender: TObject; Field: TField);
    procedure SSSPlanDataChange(Sender: TObject; Field: TField);
    procedure FListe1KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListe2KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TPlanBeforeDelete(DataSet: TDataSet);
    procedure BAutomateClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TPlanAfterDelete(DataSet: TDataSet);
    procedure TPlanAfterPost(DataSet: TDataSet);
    procedure FListe1RowEnter(Sender: TObject);
    procedure FListe2RowEnter(Sender: TObject);
    procedure TPlanPostError(DataSet: TDataSet; E: EDatabaseError;var Action: TDataAction);
    procedure TSSPlanPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure BAide2Click(Sender: TObject);
    procedure BMultiSectClick(Sender: TObject);
    procedure BSimpSectClick(Sender: TObject);
    procedure FListe1DblClick(Sender: TObject);
    procedure FListe1KeyPress(Sender: TObject; var Key: Char);
    procedure BImprimer1Click(Sender: TObject);
    procedure BImprimer2Click(Sender: TObject);
  private { Déclarations privées }
    FAxe : String ;
    Modifier1,Modifier2 : Boolean ;
    LgSect : Byte ;
    LgDeb : Byte ;
    MemoLg : Byte ;
    FAvertir : Boolean ;
    OnlyOne : Boolean ;
    Function  Bouge1(Button: TNavigateBtn) : boolean ;
    Function  Bouge2(Button: TNavigateBtn) : boolean ;
    Procedure NewEnreg1 ;
    Procedure NewEnreg2 ;
    Function  EnregOK1 : boolean ;
    Function  EnregOK2 : boolean ;
    Function  OnSauve1 : boolean ;
    Function  OnSauve2 : boolean ;
    Procedure LongueurSection ;
    Function  NewEnreg1Possible : Boolean;
    Procedure CodeEditMask ;
    Function  ChercheEnreg2 : Boolean ;
    Function  ValideLg : Boolean ;
    Function  SommeLongeur : Byte ;
    {$IFDEF CCS3}
    Function  NbSousPlan : Integer ;
    {$ENDIF}
    Procedure DetruitContenuPlan ;
    Function  TrouDansEnchainement(Var NumMess : Byte) : Boolean ;
    Function  VerifiSiExiste1 : Boolean ;
    Function  VerifiSiExiste2 : Boolean ;
    Function  MauvaiseLg : Boolean ;
    Procedure GenereCodeContenuPlan ;
    Procedure ActiveBouton ;
    Function  ChercheUnTrou : Byte ;
    Procedure SwapLaValeur ;
    Function  BourreLeCompte : Boolean ;
    procedure ErgoS3S5 ;
  public  { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses CPSection_TOM, HStatus,UtilPGI
{$IFNDEF CCMP}
{$IFNDEF GCGC}
     ,GeSecAna
{$ENDIF}
{$ENDIF}
   ;
Procedure ParamPlanAnal(Faxe : String) ;
var FStructure: TFStructure;
    PP : THPanel ;
BEGIN
  // GCO - 20/07/2005 - 15543
  if not Presence('AXE', 'X_STRUCTURE', 'X') then
  begin
    Pgiinfo('Vous devez au moins avoir un axe analytique structuré.', 'Structures analytiques');
    Exit;
  end;
  // FIN GCO

if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FStructure:=TFStructure.Create(Application) ;
FStructure.Faxe:=Faxe ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FStructure.ShowModal ;
    Finally
     FStructure.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FStructure,PP) ;
   FStructure.Show ;
   END ;
END ;

procedure TFStructure.ErgoS3S5 ;
BEGIN
DelTabsSerie(PAxe) ;
if EstSerie(S3) then
   BEGIN
   BMultiSect.Visible:=False ;
   BAutomate.Visible:=False ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 18/06/2002
Modifié le ... :   /  /    
Description .. : redimensionnement des grilles
Mots clefs ... : 
*****************************************************************}
procedure TFStructure.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
FAvertir:=False ;
TPlan.Open ;
if Faxe='' then FAxe:='A1' else PAxe.TabIndex:=StrToInt(Copy(FAxe,2,1))-1 ;
LongueurSection ;
TPlan.SetRange([FAxe],[FAxe]) ;
TSSPlan.Open ;
ErgoS3S5 ;
HMTrad.ResizeDBGridColumns(FListe1) ; HMTrad.ResizeDBGridColumns(FListe2) ;
end;

procedure TFStructure.PAxeChange(Sender: TObject);
Var OldAxe : Integer ;
    NewAxe : String ;
    i : Byte ;
begin
OldAxe:=StrToInt(Copy(FAxe,2,1)) ; NewAxe:='A'+IntToStr(PAxe.TabIndex+1) ;
if (TPlan.State in [dsEdit,dsInsert])Or(TSSPlan.State in [dsEdit,dsInsert]) then
   BEGIN PAxe.TabIndex:=OldAxe-1 ; Exit ; END ;
if TrouDansEnchainement(i) then
   BEGIN MsgBox1.Execute(i,'','') ; PAxe.TabIndex:=OldAxe-1 ; Exit ; END ;
if Not ValideLg then
    BEGIN Bouge1(nbCancel) ; PAxe.TabIndex:=OldAxe-1 ; Exit ; END ;
if VH^.Cpta[AxeToFb(NewAxe)].Structure then
   BEGIN Faxe:=NewAxe ; TPlan.SetRange([FAxe],[FAxe]) ; END
else
   BEGIN MsgBox.Execute(0,'','') ; PAxe.TabIndex:=OldAxe-1 ; END ;
LongueurSection ;
end;

Procedure TFStructure.LongueurSection ;
BEGIN
LgSect:=VH^.Cpta[AxeToFb(Faxe)].Lg ;
if LgSect<=0 then Lgaxe.Visible:=False
             else Lgaxe.Visible:=True ;
Case Faxe[2] of
   '1' : BEGIN Lgaxe.Left:=5   ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '2' : BEGIN Lgaxe.Left:=50  ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '3' : BEGIN Lgaxe.Left:=100 ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '4' : BEGIN Lgaxe.Left:=150 ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   '5' : BEGIN Lgaxe.Left:=200 ; Lgaxe.Caption:=IntToStr(LgSect)+' '+MsgBox.Mess[3] ; END ;
   End ;
END ;

procedure TFStructure.BInsert1Click(Sender: TObject);
begin Bouge1(nbInsert) ; end;

procedure TFStructure.BDelete1Click(Sender: TObject);
begin Bouge1(nbDelete) ; end;

procedure TFStructure.BDefaire1Click(Sender: TObject);
begin Bouge1(nbCancel) ; end;

procedure TFStructure.BValider1Click(Sender: TObject);
begin Modifier1:=False ; Bouge1(nbPost) ; end;

procedure TFStructure.BFerme1Click(Sender: TObject);
begin Close ; end;

procedure TFStructure.BInsert2Click(Sender: TObject);
begin Bouge2(nbInsert) ; end;

procedure TFStructure.BDelete2Click(Sender: TObject);
begin Bouge2(nbDelete) ; end;

procedure TFStructure.BDefaire2Click(Sender: TObject);
begin Bouge2(nbCancel) ; end;

procedure TFStructure.BValider2Click(Sender: TObject);
begin Modifier2:=False ; Bouge2(nbPost) ; end;

procedure TFStructure.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFStructure.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
Var i : Byte ;
begin
CanClose:=OnSauve1 ; CanClose:=OnSauve2 ;
if TrouDansEnchainement(i) then BEGIN Canclose:=False ; MsgBox1.Execute(i,'','') ; END ;
ChargeSousPlanAxe ;
end;

Function TFStructure.Bouge1(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve1 then Exit ;
   nbPost           : if Not EnregOK1 then Exit ;
   nbDelete         : if MsgBox1.execute(1,'','')<>mrYes then Exit ;
   end ;
if Button=NbInsert then BEGIN if Not NewEnreg1Possible then Exit ; END ;
if  not TransacNav(DBNav1.BtnClick,Button,10) then MessageAlerte(Msgbox.mess[1]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg1 ;
END ;

Function TFStructure.Bouge2(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve2 then Exit ;
   nbPost           : if Not EnregOK2 then Exit ;
   nbDelete         : if MsgBox2.execute(1,'','')<>mrYes then Exit ;
   end ;
if not TransacNav(DBNav2.BtnClick,Button,10) then MessageAlerte(Msgbox.mess[1]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg2 ;
END ;

Procedure TFStructure.NewEnreg1 ;
BEGIN
InitNew(TPlan) ;
TPlanSS_AXE.AsString:=FAxe ;
FListe1.Columns.Items[3].ReadOnly:=False ;
FListe1.Columns.Items[0].ReadOnly:=False ;
TPlanSS_DEBUT.AsInteger:=LgDeb ; Modifier1:=False ;
FListe1.SelectedIndex:=0 ; FListe1.SetFocus ;
END ;

Procedure TFStructure.NewEnreg2 ;
BEGIN
if (OnlyOne) And (Not (TSSPlan.BOF And TSSPlan.EOF)) then BEGIN TSSPlan.Cancel ; Exit ; END ;
InitNew(TSSPlan) ;
TSSPlanPS_AXE.AsString:=FAxe ;
TSSPlanPS_SOUSSECTION.AsString:=TPlanSS_SOUSSECTION.AsString ;
FListe2.Columns.Items[0].ReadOnly:=False ; Modifier2:=False ;
FListe2.SelectedIndex:=0 ; FListe2.SetFocus ;
END ;

Procedure TFStructure.CodeEditMask ;
Var i : Byte ;
    St : String ;
BEGIN
St:='' ;
for i:=1 to TPlanSS_LONGUEUR.AsInteger do
   if Not OnlyOne then St:=St+'a' else St:=St+'c' ;
TSSPlanPS_CODE.EditMask:='>'+St+';0; ' ;
END ;

Function TFStructure.EnregOK1 : boolean ;
BEGIN
result:=FALSE  ; Modifier1:=True ;
if TPlan.state in [dsinsert,dsedit]=False then Exit ;
if TPlan.state in [dsinsert,dsedit] then
   BEGIN
   if Not ValideLg then
      BEGIN
      FListe1.SetFocus ; Fliste1.SelectedIndex:=4 ;  Exit ;
      END ;
   if TPlanSS_SOUSSECTION.asString='' then
      BEGIN
      MsgBox1.execute(2,'','') ; FListe1.SetFocus ; Fliste1.SelectedIndex:=0 ; Exit ;
      END ;
   if TPlanSS_LIBELLE.asString='' then
      BEGIN
      MsgBox1.execute(3,'','') ;  FListe1.SetFocus ; Fliste1.SelectedIndex:=1 ; Exit ;
      END ;
   END ;
if TPlan.state in [dsinsert] then
   BEGIN
   if VerifiSiExiste1 then
      BEGIN
      MsgBox1.Execute(4,'','') ; FListe1.SetFocus ; Fliste1.SelectedIndex:=0 ; Exit ;
      END ;
   END ;
FListe1.Columns.Items[3].ReadOnly:=True ;
Result:=True  ; Modifier1:=False ;
END ;

Function TFStructure.EnregOK2 : boolean ;
BEGIN
result:=FALSE  ; Modifier2:=True ;
if TSSPlan.state in [dsinsert,dsedit]=False then Exit ;
if TSSPlan.state in [dsinsert,dsedit] then
   BEGIN
   if TSSPlanPS_CODE.asString='' then
      BEGIN
      MsgBox2.execute(2,'','') ; Fliste2.SelectedIndex:=0 ; FListe2.SetFocus ; Exit ;
      END ;
   if TSSPlanPS_LIBELLE.asString='' then
      BEGIN
      MsgBox2.execute(3,'','') ; Fliste2.SelectedIndex:=1 ; FListe2.SetFocus ; Exit ;
      END ;
   if Not BourreLeCompte then Exit ;
   END ;
if TSSPlan.state in [dsinsert] then
   BEGIN
   if VerifiSiExiste2 then
      BEGIN
      MsgBox2.Execute(4,'','') ; Fliste2.SelectedIndex:=0 ; FListe2.SetFocus ; Exit ;
      END ;
   END ;
result:=TRUE  ; Modifier2:=False ;
END ;

Function TFStructure.OnSauve1 : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ; Modifier1:=False ;
if TPlan.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox1.execute(0,'','') ;
   END else rep:=321 ;
Case rep of
  mrYes : if not Bouge1(nbPost)   then Exit ;
  mrNo  : if not Bouge1(nbCancel) then Exit ;
  mrCancel : BEGIN Abort ; Exit ; END ;
  end ;
result:=TRUE  ;
end ;

Function TFStructure.OnSauve2 : boolean ;
Var Rep : Integer ;
BEGIN
Result:=FALSE  ;
Modifier2:=True ;
if TSSPlan.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox2.execute(0,'','') ;
   END else rep:=321 ;
Case rep of
  mrYes : if not Bouge2(nbPost)   then Exit ;
  mrNo  : if not Bouge2(nbCancel) then Exit ;
  mrCancel : BEGIN Abort ; Exit ; END ;
  end ;
Modifier1:=False ;
result:=TRUE  ;
end ;

procedure TFStructure.TPlanNewRecord(DataSet: TDataSet);
begin
if Not NewEnreg1Possible then Exit ;
Newenreg1 ;
end;

procedure TFStructure.TSSPlanNewRecord(DataSet: TDataSet);
begin Newenreg2 ; end;

procedure TFStructure.SPlanStateChange(Sender: TObject);
begin Modifier1:=True ; end;

procedure TFStructure.SPlanUpdateData(Sender: TObject);
Var i : Byte ;
begin
//if Modifier1 then if Not OnSauve1 then BEGIN SysUtils.Abort ; Exit ; END ;
//if Not FListe1.Focused then Exit ;
if((Trim(TPlanSS_SOUSSECTION.AsString)='') And
   (Trim(TPlanSS_LIBELLE.AsString)='')) And
   ((TPlanSS_LONGUEUR.AsInteger)=0) then
   BEGIN TPlan.Cancel ; Exit ; END ;
if(Trim(TPlanSS_SOUSSECTION.AsString)<>'') then
  BEGIN
  if(TPlanSS_LONGUEUR.AsInteger)=0 then
     BEGIN
     MsgBox1.Execute(11,'','') ; FListe1.SetFocus ; Fliste1.SelectedIndex:=4 ;
     SysUtils.Abort ; Exit ;
     END ;
  if(Trim(TPlanSS_LIBELLE.AsString)='')then
     BEGIN
     MsgBox1.Execute(3,'','') ; FListe1.SetFocus ; Fliste1.SelectedIndex:=1 ;
     SysUtils.Abort ; Exit ;
     END ;
  END ;
if(Trim(TPlanSS_SOUSSECTION.AsString)='') then
   BEGIN
   if(Trim(TPlanSS_LIBELLE.AsString)<>'')Or
     ((TPlanSS_LONGUEUR.AsInteger)<>0) then
      BEGIN
      MsgBox1.Execute(2,'','') ; FListe1.SetFocus ; Fliste1.SelectedIndex:=0 ;
      SysUtils.Abort ; Exit ;
      END ;
   END ;
if TPlan.State=dsEdit then
   if TrouDansEnchainement(i) then BEGIN MsgBox1.Execute(i,'','') ; SysUtils.Abort ; Exit ; END ;
if Not ValideLg then
   BEGIN FListe1.SetFocus ; FListe1.SelectedIndex:=4 ; SysUtils.Abort ; Exit ; END ;
end;

procedure TFStructure.SSSPlanStateChange(Sender: TObject);
begin
Modifier2:=True ;
end;

procedure TFStructure.SSSPlanUpdateData(Sender: TObject);
begin
if(Trim(TSSPlanPS_CODE.AsString)='') And (Trim(TSSPlanPS_LIBELLE.AsString)='') then
   BEGIN TSSPlan.Cancel ; Exit ; END ;
if(Trim(TSSPlanPS_CODE.AsString)<>'')And(Trim(TSSPlanPS_LIBELLE.AsString)='') then
  BEGIN
  MSgBox2.Execute(3,'','') ; FListe2.SelectedIndex:=1 ; Fliste2.SetFocus ; SysUtils.Abort ; Exit ;
  END ;
if(Trim(TSSPlanPS_LIBELLE.AsString)<>'')And(Trim(TSSPlanPS_CODE.AsString)='')then
  BEGIN
  MsgBox2.Execute(2,'','') ;  FListe2.SelectedIndex:=0 ; Fliste2.SetFocus ; SysUtils.Abort ; Exit ;
  END ;
if Not BourreLeCompte then SysUtils.Abort ;
end;

Function TFStructure.NewEnreg1Possible : Boolean ;
Var Lg : Byte ;
    i : Byte ;
BEGIN
Result:=True ;
Lg:=SommeLongeur ;
if Lg>=LgSect then
   BEGIN
   MsgBox1.Execute(5,'','') ; Result:=False ; Bouge1(nbCancel) ;
   END else
   BEGIN
   i:=ChercheUnTrou ;
   if(i=0)then LgDeb:=Lg+1
          else LgDeb:=i ;
   END ;
{$IFDEF CCS3}
If Result And (NbSousPlan=3) Then
  BEGIN
  MsgBox1.Execute(12,'','') ; Result:=False ; Bouge1(nbCancel) ;
  END ;
{$ENDIF}
END ;

Procedure TFStructure.ActiveBouton ;
BEGIN
BInsert1.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert]))And
                  (Not(TSSPlan.State in [dsEdit,dsInsert])) ;
BDelete1.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert]))And
                  (Not(TSSPlan.State in [dsEdit,dsInsert])) ;
BDefaire1.Enabled:=(Not(TSSPlan.State in [dsEdit,dsInsert])) ;
BDefaire2.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert])) ;
BInsert2.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert]))And
                  (Not(TSSPlan.State in [dsEdit,dsInsert]))
                   And(Not((TPlan.Eof)And(TPlan.Bof))) ;
BDelete2.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert]))And
                  (Not(TSSPlan.State in [dsEdit,dsInsert])) ;
BFerme.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert])) ;
BMultiSect.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert]))And
                    (Not(TSSPlan.State in [dsEdit,dsInsert])) ;
BSimpSect.Enabled:=(Not(TPlan.State in [dsEdit,dsInsert]))And
                   (Not(TSSPlan.State in [dsEdit,dsInsert])) ;
//BFerme1.Enabled:=(Not(TSSPlan.State in [dsEdit,dsInsert])) ;
END ;

procedure TFStructure.SPlanDataChange(Sender: TObject; Field: TField);
begin
ActiveBouton ;
if TPlan.State=dsBrowse then FListe1.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then MemoLg:=TPlanSS_LONGUEUR.AsInteger ;
if(TPlan.Eof)And(TPlan.Bof) then BDelete1.Enabled:=False ;
PPlan.Caption:=Msgbox.Mess[2]+Fliste1.Fields[1].AsString ;
Panel5.Caption:=Msgbox.Mess[4]+Fliste1.Fields[0].AsString ;
if TPlanSS_CONTROLE.AsString='X' then OnlyOne:=True else OnlyOne:=False ;
CodeEditMask ;
end;

procedure TFStructure.SSSPlanDataChange(Sender: TObject; Field: TField);
begin
ActiveBouton ;
if(TSSPlan.Eof)And(TSSPlan.Bof)then BDelete2.Enabled:=False ;
if((TPlan.Eof)And(TPlan.Bof))Or(TPlanSS_SOUSSECTION.AsString='')
  Or(TPlan.State in [dsInsert,dsEdit])Or(TSSPlan.State in [dsInsert,dsEdit])then
  BEGIN
  BAutomate.Enabled:=False ; BMultiSect.Enabled:=False ; BSimpSect.Enabled:=False ;
  END else
  BEGIN
  BAutomate.Enabled:=True ; BMultiSect.Enabled:=True ; BSimpSect.Enabled:=True ;
  END ;
end;

Function TFStructure.ChercheEnreg2 : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select PS_CODE From SSSTRUCR Where PS_AXE="'+Faxe+'" And '+
           'PS_SOUSSECTION="'+TPlanSS_SOUSSECTION.AsString+'"',True,-1,'',true) ;
Result:=(Q.Eof) ; Ferme(Q) ;
END ;

Function TFStructure.ValideLg : Boolean ;
Var i : Byte ;
BEGIN
Result:=True ;
if(TPlan.Eof)And(TPlan.Bof) then Exit ;
if TPlanSS_LONGUEUR.AsInteger=0 then
   BEGIN Result:=False ; MsgBox1.Execute(11,'','') ; Exit ; END ;
if TPlanSS_LONGUEUR.AsInteger > MemoLg then
   BEGIN
   i:=SommeLongeur ;
   i:=i-MemoLg+TPlanSS_LONGUEUR.AsInteger ;
   if i > LgSect then
      BEGIN Result:=False ; MsgBox1.Execute(6,'','') ; Exit ; END ;
   END ;
if TPlanSS_LONGUEUR.AsInteger<>MemoLg then
   BEGIN
   if Not ChercheEnreg2 then
      BEGIN
       if MsgBox1.Execute(7,'','')=mrYes then DetruitContenuPlan
                                         else BEGIN TPlanSS_LONGUEUR.AsInteger:=MemoLg ; Result:=False ; END ;
      END ;
   END ;
END ;

Function TFStructure.SommeLongeur : Byte ;
Var Q : TQuery ;
BEGIN
Result:=0 ;
Q:=OpenSql('Select Sum(SS_LONGUEUR) from STRUCRSE where SS_AXE="'+Faxe+'"',True,-1,'',true) ;
If Not Q.Eof Then Result:=Q.Fields[0].AsInteger ; Ferme(Q) ;
END ;

{$IFDEF CCS3}
Function TFStructure.NbSousPlan : Integer ;
Var Q : TQuery ;
BEGIN
Result:=0 ;
Q:=OpenSql('Select Count(*) from STRUCRSE where SS_AXE="'+Faxe+'"',True,-1,'',true) ;
If Not Q.Eof Then Result:=Q.Fields[0].AsInteger ; Ferme(Q) ;
END ;
{$ENDIF}


Procedure TFStructure.DetruitContenuPlan ;
BEGIN
ExecuteSql('Delete from SSSTRUCR Where PS_AXE="'+Faxe+'" And '+
           'PS_SOUSSECTION="'+TPlanSS_SOUSSECTION.AsString+'"') ;
if V_PGI.Driver=dbMSACCESS then BEGIN TSSPlan.Close ; TSSPlan.Open ; END else TSSPlan.Refresh ;
END ;

Function TFStructure.TrouDansEnchainement(Var NumMess : Byte) : Boolean ;
Var Q : TQuery ;
    Ok : Boolean ;
    DL : Byte ;
BEGIN
if(TPlan.Eof)And(TPlan.Bof) then BEGIN Result:=False ; Exit ; END ;
if TPlanSS_LONGUEUR.AsInteger=0 then BEGIN NumMess:=11 ; Result:=True ; Exit ; END ;
if TPlanSS_LONGUEUR.AsInteger<0 then BEGIN NumMess:=9 ; Result:=True ; Exit ; END ;
Q:=OpenSql('Select SS_DEBUT,SS_LONGUEUR From STRUCRSE Where SS_AXE="'+Faxe+'" '+
           'Order by SS_DEBUT',True,-1,'',true) ;
Ok:=False ;
While Not Q.Eof And Not Ok do
   BEGIN
    DL:=Q.Fields[0].AsInteger+Q.Fields[1].AsInteger ;
    Q.Next ;
    if(DL<>Q.Fields[0].AsInteger)And(Not Q.Eof) then
       BEGIN Ok:=True ; NumMess:=8 ; FListe1.Columns.Items[3].ReadOnly:=False ; END ;
   END ;
Ferme(Q) ; Result:=Ok ;
END ;

procedure TFStructure.FListe1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if((ssCtrl in Shift)AND(Key=VK_DELETE)AND(TPlan.Eof)AND(TPlan.Bof)) then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)AND(Key=VK_DELETE))then BEGIN Bouge1(nbDelete) ; Key:=0 ; END ;
end;

procedure TFStructure.FListe1KeyPress(Sender: TObject; var Key: Char);
begin
if (FListe1.SelectedIndex=2) then if Key=#32 then SwapLaValeur ;
end;

procedure TFStructure.FListe2KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if(ssCtrl in Shift)AND(Key=VK_DELETE)AND(TSSPlan.EOF)AND(TSSPlan.BOF)then BEGIN Key:=0 ; Exit ; END ;
if((ssCtrl in Shift)AND(Key=VK_DELETE))then BEGIN Bouge2(nbDelete) ; Key:=0 ; END ;
end;

procedure TFStructure.TPlanBeforeDelete(DataSet: TDataSet);
begin DetruitContenuPlan ; end;

Function TFStructure.VerifiSiExiste1 : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select SS_SOUSSECTION From STRUCRSE Where SS_AXE="'+Faxe+'" And '+
           'SS_SOUSSECTION="'+TPlanSS_SOUSSECTION.AsString+'"',True,-1,'',true) ;
Result:=(Not Q.Eof) ; Ferme(Q) ;
END ;

Function TFStructure.VerifiSiExiste2 : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select PS_CODE From SSSTRUCR Where PS_AXE="'+Faxe+'" And '+
           'PS_SOUSSECTION="'+TPlanSS_SOUSSECTION.AsString+'" And '+
           'PS_CODE="'+TSSPlanPS_CODE.AsString+'"',True,-1,'',true) ;
Result:=(Not Q.Eof) ; Ferme(Q) ;
END ;

Function TFStructure.MauvaiseLg : Boolean ;
BEGIN Result:=(Length(TSSPlanPS_CODE.AsString)< TPlanSS_LONGUEUR.AsInteger) ; END ;

Procedure TFStructure.GenereCodeContenuPlan ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select S_SECTION From SECTION Where S_AXE="'+Faxe+'"',True,-1,'',true) ;
InitMove(RecordsCount(Q),MsgBox2.Mess[7]) ;
BeginTrans ;
While Not Q.Eof do
   BEGIN
   CodeStructure(Faxe,Q.Fields[0].AsString) ; Q.Next ; MoveCur(False) ;
   END ;
CommitTrans ; Ferme(Q) ; FiniMove ; 
END ;

procedure TFStructure.BAutomateClick(Sender: TObject);
begin
if MsgBox2.Execute(5,'','')=mrYes then
   BEGIN GenereCodeContenuPlan ; TSSPlan.Refresh ; TSSPlan.First ; END ;
end;

procedure TFStructure.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
   BEGIN
   ChargeStructureUnique ;
   AvertirTable('ttStrucrse1') ;
   AvertirTable('ttStrucrse2') ;
   AvertirTable('ttStrucrse3') ;
   AvertirTable('ttStrucrse4') ;
   AvertirTable('ttStrucrse5') ;
   END ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFStructure.TPlanAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFStructure.TPlanAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFStructure.FListe1RowEnter(Sender: TObject);
begin
if TPlan.State=dsInsert then FListe1.SelectedIndex:=0 ;
end;

procedure TFStructure.FListe2RowEnter(Sender: TObject);
begin
if TSSPlan.State=dsInsert then FListe2.SelectedIndex:=0 ;
end;

procedure TFStructure.TPlanPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
if TPlan.State=dsInsert then
   BEGIN MsgBox1.Execute(4,'','') ; Action:=daAbort ; END ;
end;

procedure TFStructure.TSSPlanPostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
begin
if TSSPlan.State=dsInsert then
   BEGIN MsgBox2.Execute(4,'','') ; Action:=daAbort ; END ;
end;

Function TFStructure.ChercheUnTrou : Byte ;
Var Trou : Byte ;
    Q : TQuery ;
    Premiertour : Boolean ;
BEGIN
if(TPlan.Eof)And(TPlan.Bof)then BEGIN Result:=0 ; Exit ; END ;
Trou:=0 ;
Q:=OpenSql('Select SS_DEBUT,SS_LONGUEUR From STRUCRSE Where SS_AXE="'+Faxe+'" '+
           'Order by SS_DEBUT',True,-1,'',true) ;
PremierTour:=True ;
While Not Q.Eof do
   BEGIN
   if PremierTour then
      BEGIN
      if Q.Fields[0].AsInteger<>1 then BEGIN Trou:=1 ; Break ; END else PremierTour:=False ;
      END ;
    Trou:=Q.Fields[0].AsInteger+Q.Fields[1].AsInteger ;
    Q.Next ;
    if(Trou<>Q.Fields[0].AsInteger)And(Not Q.Eof) then Break ;
   END ;
if Q.Eof then Trou:=0 ;
Ferme(Q) ; Result:=Trou ;
END ;

procedure TFStructure.BAide2Click(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFStructure.BMultiSectClick(Sender: TObject);
begin
{$IFNDEF CCMP}
{$IFNDEF GCGC}
if MsgBox2.Execute(8,'','')=mrOk then
GenereSectionsAna(Faxe,TPlanSS_SOUSSECTION.AsString,TSSPlanPS_CODE.AsString,TSSPlanPS_LIBELLE.AsString,False) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFStructure.BSimpSectClick(Sender: TObject);
begin
{$IFNDEF CCMP}
{$IFNDEF GCGC}
GenereSectionsAna(Faxe,'','','',False) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFStructure.FListe1DblClick(Sender: TObject);
begin
if FListe1.SelectedIndex<>2 then Exit ;
SwapLaValeur ;
end;

Procedure TFStructure.SwapLaValeur ;
Var MyBookMark : TBookMark ;
    St : String ;
BEGIN
if TPlan.State=dsBrowse then TPlan.Edit ;
if TPlanSS_CONTROLE.AsString='-' then TPlanSS_CONTROLE.AsString:='X'
                                 else TPlanSS_CONTROLE.AsString:='-' ;
MyBookMark:=TPlan.GetBookMark ; St:=TPlanSS_SOUSSECTION.AsString ;
TPlan.Post ; TPlan.Close ;
ExecuteSql('Update STRUCRSE Set SS_CONTROLE="-" Where SS_AXE="'+Faxe+'" And '+
           'SS_SOUSSECTION<>"'+St+'" ') ;
TPlan.Open ; TPlan.SetRange([Faxe],[Faxe]) ;
TPlan.GotoBookmark(MyBookmark) ; TPlan.FreeBookmark(MyBookmark);
TPlan.Edit ; Modifier1:=True ;
END ;

Function TFStructure.BourreLeCompte : Boolean ;
Var i : Integer ;
BEGIN
Result:=True ;
if MauvaiseLg then
   BEGIN
   if Not OnlyOne then
      BEGIN
      for i:=Length(TSSPlanPS_CODE.AsString)+1 to TPlanSS_LONGUEUR.AsInteger do
          TSSPlanPS_CODE.AsString:=TSSPlanPS_CODE.AsString+VH^.Cpta[AxeToFb(Faxe)].Cb ;
      END else
      BEGIN
      MsgBox2.execute(6,'','') ; Fliste2.SelectedIndex:=1 ; FListe2.SetFocus ; Result:=False ; Exit ;
      END ;
   END ;
END ;

procedure TFStructure.BImprimer1Click(Sender: TObject);
begin
PrintDBGrid (FListe1,Nil, Caption+' '+PAxe.Tabs[PAxe.TabIndex],'');
//if ((HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid) and (V_PGI.OutLook)) then HMTrad.ResizeDBGridColumns(FListe) ;
end;

procedure TFStructure.BImprimer2Click(Sender: TObject);
Var St : String ;
begin
If (Not TPlan.Eof) And Not (Not TPlan.Bof) Then St:=TPLANSS_LIBELLE.AsString Else St:='' ;
PrintDBGrid (FListe2,Nil, Caption+' '+PAxe.Tabs[PAxe.TabIndex]+' '+St,'');
end;

end.
