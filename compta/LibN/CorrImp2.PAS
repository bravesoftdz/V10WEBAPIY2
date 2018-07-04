unit  CorrImp2 ;

interface

uses SysUtils, WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons,Dialogs,
     StdCtrls, ExtCtrls, Hctrls, DB, Grids, DBGrids, Mask, DBCtrls,
     Ent1, Spin, HmsgBox, Hqry, HEnt1, ComCtrls, HDB, Messages, HSysMenu, HPanel, UiUtil,
     HTB97,Hcompte,HXLSPAS,Hstatus,Lookup, ADODB, udbxDataset ;

procedure ParamCorrespImp2 ( Quoi : String ) ;
Procedure ZoomCorrespImp2  (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);

type
  TFCorrImp2 = class(TForm)
    FListe: THDBGrid;
    TCORRESP: THTable;
    SCORRESP : TDataSource;
    DBNav    : TDBNavigator;
    HPB: TToolWindow97;
    BDefaire: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    MsgBox   : THMsgBox;
    TCORRESPCR_TYPE   : TStringField;
    TCORRESPCR_CORRESP: TStringField;
    TCORRESPCR_LIBELLE: TStringField;
    HMTrad: THSystemMenu;
    FAutoSave: TCheckBox;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock: TDock97;
    Cache: THCpteEdit;
    SD: TSaveDialog;
    TCORRESPCR_ABREGE: TStringField;
    ModeReg: TEdit;
    TCORRESPCR_LIBRETEXTE1: TStringField;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure TCorrespNewRecord(DataSet: TDataset);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BDefaireClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure SCORRESPDataChange(Sender: TObject; Field: TField);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure TCORRESPAfterDelete(DataSet: TDataSet);
    procedure TCORRESPAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TCORRESPPostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure SCORRESPUpdateData(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
  private
    FType : string ;
    FAvertir : boolean ;
    WMinX,WMinY : Integer ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure NewEnreg ;
    Procedure ChargeEnreg ;
    Function  EnregOK : boolean ;
    Function  OnSauve : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure FaitCaption ;
    Procedure Initialisation ;
    Function  Supprime : Boolean ;
    Function  VerifiSiExiste : Boolean ;
    procedure ClickZoom ;
  public
  end;

const
     CL_LIBELLE     = 0 ;
     CL_ABREGE      = 1 ;
     CL_LIBRETEXTE1 = 2 ;
     CL_CORRESP     = 3 ;
implementation

uses PrintDBG, HZoomSP,FichComm ;
{$R *.DFM}


{---Procédure d'appel de la fiche CORRESP---------------------------------}
Procedure ZoomCorrespImp2  (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
var FCorrImp: TFCorrImp2;
    PP : THPanel ;
begin
FCorrImp:=TFCorrImp2.Create(Application) ;
FCorrImp.FType:=Axe ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCorrImp.ShowModal ;
    finally
     FCorrImp.Free ;
    end ;
   END else
   BEGIN
   InitInside(FCorrImp,PP) ;
   FCorrImp.Show ;
   END ;                                          
Screen.Cursor:=SyncrDefault ;
END ;

procedure ParamCorrespImp2 ( Quoi : String ) ;
var FCorrImp: TFCorrImp2;
    PP : THPanel ;
begin
FCorrImp:=TFCorrImp2.Create(Application) ;
FCorrImp.FType:=Quoi ;
if Quoi='IRM' then FCorrImp.HelpContext := 0 ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FCorrImp.ShowModal ;
    finally
     FCorrImp.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FCorrImp,PP) ;
   FCorrImp.Show ;
   END ;
end;

Procedure TFCorrImp2.FaitCaption ;
BEGIN
if FType='IRM' then Caption:=Caption+' '+MsgBox.Mess[8] ;
UpdateCaption(Self) ;
END ;

Procedure TFCorrImp2.Initialisation ;
BEGIN
TCorresp.SetRange([FType],[FType]) ;
FAvertir:=False ;
if(TCorresp.EOF)And(TCorresp.BOF) then Bouge(nbInsert) ;
END ;

procedure TFCorrImp2.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TCorresp.Open ;
Initialisation ; FaitCaption ;
FListe.SetFocus ;
end;


procedure TFCorrImp2.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFCorrImp2.TCorrespNewRecord(DataSet: TDataset);
begin NewEnreg ; end;

procedure TFCorrImp2.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCorrImp2.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCorrImp2.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCorrImp2.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCorrImp2.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFCorrImp2.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFCorrImp2.BDefaireClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCorrImp2.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFCorrImp2.NewEnreg ;
BEGIN
InitNew(TCorresp) ;
TCORRESPCR_TYPE.AsString:=FType ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.SelectedIndex:=0 ; FListe.SetFocus ;
END ;

Procedure TFCorrImp2.ChargeEnreg ;
BEGIN
// code ...
END ;

Function TFCorrImp2.VerifiSiExiste : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select CR_TYPE,CR_CORRESP From CORRESP Where CR_TYPE="'+TCORRESPCR_TYPE.AsString+'" And '+
              'CR_CORRESP="'+TCORRESPCR_CORRESP.AsString+'"',True) ;
Result:=(Not QLoc.Eof) ; Ferme(QLoc) ;
END ;

Function TFCorrImp2.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if not (TCORRESP.state in [dsEdit,dsInsert]) then Exit ;
if trim(TCORRESPCR_LIBRETEXTE1.AsString)='' then
  begin MsgBox.Execute(2,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=CL_LIBRETEXTE1 ; Exit ; end ;
if trim(TCORRESPCR_LIBELLE.AsString)='' then
  begin MsgBox.Execute(3,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=CL_LIBELLE ;     Exit ; end ;
if length(TCORRESPCR_LIBRETEXTE1.AsString)>3 then
  begin MsgBox.Execute(25,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=CL_LIBRETEXTE1 ; Exit ; end ;
if length(TCORRESPCR_LIBELLE.AsString)>14 then
  begin MsgBox.Execute(26,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=CL_LIBELLE ;     Exit ; end ;
if TCORRESP.state in [dsInsert] then
   begin
   if VerifiSiExiste then
      begin MsgBox.Execute(4,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=CL_LIBRETEXTE1 ; Exit ; end ;
   end ;
Result:=TRUE  ;
END ;


procedure TFCorrImp2.SCORRESPUpdateData(Sender: TObject);
begin
if(Trim(TCORRESPCR_LIBRETEXTE1.AsString)='') And (Trim(TCORRESPCR_LIBELLE.AsString)='') then BEGIN TCORRESP.Cancel ; Exit ; END ;
if not EnregOK then begin SysUtils.Abort ; Exit ; end ;
FListe.Fields[CL_CORRESP].AsString:=FListe.Fields[CL_LIBELLE].AsString+';'+FListe.Fields[CL_ABREGE].AsString;
end ;

Function TFCorrImp2.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TCORRESP.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : if not Bouge(nbCancel) then exit ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFCorrImp2.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : if Not Supprime then Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[6]);
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

procedure TFCorrImp2.SCORRESPDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
BInsert.Enabled:=Not((TCORRESP.State in [dsInsert,dsEdit])) ;
BDelete.Enabled:=Not((TCORRESP.State in [dsInsert,dsEdit])) ;
if(TCORRESP.Eof)And(TCORRESP.Bof) then BDelete.Enabled:=False ;
//if TCORRESP.State=dsBrowse then Fliste.Columns.Items[0].ReadOnly:=True ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TCORRESP.BOF;
   DnEnable := Enabled and not TCORRESP.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END ;
end;

procedure TFCorrImp2.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFCorrImp2.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,'') ; end;

procedure TFCorrImp2.TCORRESPAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrImp2.TCORRESPAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrImp2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
  BEGIN
  END ;
if Parent is THPanel then
   BEGIN
   Action:=caFree ;
   END ;
end;

procedure TFCorrImp2.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var Vide : Boolean ;
begin
if(ssCtrl in Shift)And(Key=VK_DELETE)And(TCORRESP.EOF)And(TCORRESP.BOF) then
  BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)And(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
Vide:=(Shift=[]) ;
Case Key of
   VK_F5     : if (Vide) then BEGIN Key:=0 ; ClickZoom ; END ;
   end ;
end;

procedure TFCorrImp2.TCORRESPPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
if TCORRESP.State=dsInsert then
   BEGIN MsgBox.Execute(4,'','') ; Fliste.SelectedIndex:=CL_LIBRETEXTE1 ; Action:=daAbort ; END ;
end;

Function TFCorrImp2.Supprime : Boolean ;
Var St : String ;
    i : Byte ;
BEGIN
Result:=False ;
if (MsgBox.execute(1,'','')<>mrYes) then Exit ;
St:=TCORRESPCR_CORRESP.AsString ; i:=0 ;
if FType='IGE' then i:=1 else
if FType='IAU' then i:=2 else
if FType='IBU' then i:=3 else
if(FType='IA1')Or(FType='IA2')Or(FType='IA3')Or(FType='IA4')Or(FType='IA5') then i:=4 ;
(*
Case i of
     1:ExecuteSql('Update GENERAUX SET G_CORRESP'+zz+'="" Where G_CORRESP'+zz+'="'+St+'"') ;
     2:ExecuteSql('Update TIERS SET T_CORRESP'+zz+'="" Where T_CORRESP'+zz+'="'+St+'"') ;
//Simon     3:ExecuteSql('Update BUDGET SET B_CORRESP'+zz+'="" Where B_CORRESP'+zz+'="'+St+'"') ;
     4:ExecuteSql('Update SECTION SET S_CORRESP'+zz+'="" Where S_AXE="'+FType+'" AND S_CORRESP'+zz+'="'+St+'"') ;
     End ;
*)
Result:=True ; Screen.Cursor:=SyncrDefault ;

END ;

procedure TFCorrImp2.FListeRowEnter(Sender: TObject);
begin if TCORRESP.State=dsInsert then Fliste.SelectedIndex:=CL_LIBRETEXTE1 ; end;

procedure TFCorrImp2.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCorrImp2.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFCorrImp2.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFCorrImp2.ClickZoom ;
begin
If FListe.SelectedField.FieldName<>'CR_LIBRETEXTE1' Then Exit ;
ModeReg.Text:='' ;
if LookupList(ModeReg,TraduireMemoire('Mode de reglement'),'MODEREGL','MR_MODEREGLE','MR_LIBELLE','','MR_MODEREGLE',TRUE, 1) then
begin
  if TCorresp.State=dsBrowse then TCorresp.Edit ;
  FListe.Fields[CL_LIBRETEXTE1].AsString:=ModeReg.Text ;
  FListe.Fields[CL_CORRESP].AsString:=FListe.Fields[CL_LIBELLE].AsString+';'+FListe.Fields[CL_ABREGE].AsString+';';
end ;
end;

procedure TFCorrImp2.FListeDblClick(Sender: TObject);
begin
ClickZoom ;
end;

end.
