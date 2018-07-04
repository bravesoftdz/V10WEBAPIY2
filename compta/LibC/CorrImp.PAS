unit  CORRIMP ;

interface

uses SysUtils, WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons,Dialogs,
     StdCtrls, ExtCtrls, Hctrls, DBTables, Grids, DBGrids, Mask, DBCtrls,
     Ent1, Spin, HmsgBox, Hqry, HEnt1, ComCtrls, HDB, Messages, HSysMenu, HPanel, UiUtil,
     HTB97,Hcompte,HXLSPAS,Hstatus, DB, ADODB, udbxDataset, TntDBGrids ;

procedure ParamCorrespImp ( Quoi : String ) ;
Procedure ZoomCorrespImp  (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);

type
  TFCorrImp = class(TForm)
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
    bExport: TToolbarButton97;
    bImport: TToolbarButton97;
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
    procedure bExportClick(Sender: TObject);
    procedure bImportClick(Sender: TObject);
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
    procedure ClickZoomSousPlan ;
    Function  OkZoom (Var A : TActionFiche) : Boolean ;
  public
  end;


implementation

uses PrintDBG, HZoomSP,
{$IFDEF RECUPLTL}
     ImpFicU,
{$ELSE}
{$IFDEF RECUPSISCOS3}
     ImpFicU,
{$ENDIF}
{$ENDIF}
     FichComm ;
{$R *.DFM}


{---Procédure d'appel de la fiche CORRESP---------------------------------}
Procedure ZoomCorrespImp  (Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
var FCorrImp: TFCorrImp;
    PP : THPanel ;
begin
FCorrImp:=TFCorrImp.Create(Application) ;
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

procedure ParamCorrespImp ( Quoi : String ) ;
var FCorrImp: TFCorrImp;
    PP : THPanel ;
begin
FCorrImp:=TFCorrImp.Create(Application) ;
FCorrImp.FType:=Quoi ;
if Quoi='IGE' then FCorrImp.HelpContext := 0 ;
if Quoi='IAU' then FCorrImp.HelpContext := 0 ;
if (Quoi='IA1') or (Quoi='IA2') or (Quoi='IA3') or (Quoi='IA4') or (Quoi='IA5') then FCorrImp.HelpContext := 0 ;
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

Procedure TFCorrImp.FaitCaption ;
BEGIN
if FType='IGE' then Caption:=Caption+' '+MsgBox.Mess[8] else
   if FType='IAU' then Caption:=Caption+' '+MsgBox.Mess[9] else
      if FType='IA1' then Caption:=Caption+' '+MsgBox.Mess[10] else
         if FType='IA2' then Caption:=Caption+' '+MsgBox.Mess[11] else
            if FType='IA3' then Caption:=Caption+' '+MsgBox.Mess[12] else
               if FType='IA4' then Caption:=Caption+' '+MsgBox.Mess[13] else
                  if FType='IA5' then Caption:=Caption+' '+MsgBox.Mess[14] else
                     if FType='IBU' then Caption:=Caption+' '+MsgBox.Mess[15] ;
                        if FType='IPM' then Caption:=Caption+' '+MsgBox.Mess[16] ;
                           if FType='IJA' then Caption:=Caption+' '+MsgBox.Mess[20] ;
UpdateCaption(Self) ;
END ;

Procedure TFCorrImp.Initialisation ;
BEGIN
TCorresp.SetRange([FType],[FType]) ;
FAvertir:=False ;
if(TCorresp.EOF)And(TCorresp.BOF) then Bouge(nbInsert) ;
END ;

procedure TFCorrImp.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TCorresp.Open ;
Initialisation ; FaitCaption ;
If FType='IPM' Then
  BEGIN
  FListe.Columns[0].Title.Caption:=MsgBox.Mess[18] ;
  FListe.Columns[1].Title.Caption:=MsgBox.Mess[19] ;
  END ;
If FType='IJA' Then
  BEGIN
  FListe.Columns[0].Title.Caption:=MsgBox.Mess[21] ;
  FListe.Columns[1].Title.Caption:=MsgBox.Mess[22] ;
  END ;
FListe.SetFocus ;
end;


procedure TFCorrImp.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFCorrImp.TCorrespNewRecord(DataSet: TDataset);
begin NewEnreg ; end;

procedure TFCorrImp.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCorrImp.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCorrImp.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCorrImp.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCorrImp.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFCorrImp.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFCorrImp.BDefaireClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCorrImp.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFCorrImp.NewEnreg ;
BEGIN
InitNew(TCorresp) ;
TCORRESPCR_TYPE.AsString:=FType ;
FListe.Columns.Items[0].ReadOnly:=False ;
FListe.SelectedIndex:=0 ; FListe.SetFocus ;
END ;

Procedure TFCorrImp.ChargeEnreg ;
BEGIN
// code ...
END ;

Function TFCorrImp.VerifiSiExiste : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select CR_TYPE,CR_CORRESP From CORRESP Where CR_TYPE="'+TCORRESPCR_TYPE.AsString+'" And '+
              'CR_CORRESP="'+TCORRESPCR_CORRESP.AsString+'"',True) ;
Result:=(Not QLoc.Eof) ; Ferme(QLoc) ;
END ;

Function TFCorrImp.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TCORRESP.state in [dsEdit,dsInsert]=False then Exit ;
if TCORRESP.state in [dsEdit,dsInsert] then
   BEGIN
   if TCORRESPCR_CORRESP.AsString='' then
      BEGIN MsgBox.Execute(2,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   if TCORRESPCR_LIBELLE.AsString='' then
      BEGIN MsgBox.Execute(3,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=1 ; Exit ; END ;
   END else exit ;
if TCORRESP.state in [dsInsert] then
   BEGIN
   if VerifiSiExiste then
      BEGIN MsgBox.Execute(4,'','') ; FListe.SetFocus ; FListe.SelectedIndex:=0 ; Exit ; END ;
   END ;
Result:=TRUE  ;
END ;

Function TFCorrImp.OnSauve : boolean ;
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

Function TFCorrImp.Bouge(Button: TNavigateBtn) : boolean ;
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

procedure TFCorrImp.SCORRESPDataChange(Sender: TObject; Field: TField);
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

procedure TFCorrImp.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFCorrImp.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,'') ; end;

procedure TFCorrImp.TCORRESPAfterDelete(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrImp.TCORRESPAfterPost(DataSet: TDataSet);
begin FAvertir:=True ; end;

procedure TFCorrImp.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if FAvertir then
  BEGIN
  END ;
if Parent is THPanel then
   BEGIN
   Action:=caFree ;
   END ;
end;

procedure TFCorrImp.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var Vide : Boolean ;
begin
if(ssCtrl in Shift)And(Key=VK_DELETE)And(TCORRESP.EOF)And(TCORRESP.BOF) then
  BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)And(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
Vide:=(Shift=[]) ;
Case Key of
   VK_F5     : if (Vide) then BEGIN Key:=0 ; ClickZoom ; END Else
                  If (Shift=[ssCtrl]) then BEGIN Key:=0 ; ClickZoomSousPlan ; END ;
   END ;
end;

procedure TFCorrImp.TCORRESPPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
if TCORRESP.State=dsInsert then
   BEGIN MsgBox.Execute(4,'','') ; Fliste.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

Function TFCorrImp.Supprime : Boolean ;
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

procedure TFCorrImp.SCORRESPUpdateData(Sender: TObject);
begin
if(Trim(TCORRESPCR_CORRESP.AsString)='') And (Trim(TCORRESPCR_LIBELLE.AsString)='')
   then BEGIN TCORRESP.Cancel ; Exit ; END ;
if(Trim(TCORRESPCR_LIBELLE.AsString)<>'') And (Trim(TCORRESPCR_CORRESP.AsString)='') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TCORRESPCR_CORRESP.AsString)<>'') And (Trim(TCORRESPCR_LIBELLE.AsString)='') then
  BEGIN MsgBox.Execute(3,'','') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
end ;

procedure TFCorrImp.FListeRowEnter(Sender: TObject);
begin if TCORRESP.State=dsInsert then Fliste.SelectedIndex:=0 ; end;

procedure TFCorrImp.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCorrImp.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFCorrImp.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Function TFCorrImp.OkZoom (Var A : TActionFiche) : Boolean ;
Var CCo : TConcept ;
    OkTest : Boolean ;
BEGIN
Result:=FALSE ; OkTest:=TRUE ;
If FListe.SelectedField.FieldName<>'CR_LIBELLE' Then Exit ;
if FType='IGE' then BEGIN CCo:=TConcept(ccGenModif) ; Cache.ZoomTable:=tzGeneral ; END  else
   if FType='IAU' then BEGIN CCo:=TConcept(ccAuxModif) ; Cache.ZoomTable:=tzTiers ; END else
     if FType='IJA' then BEGIN CCo:=TConcept(ccJalModif) ; Cache.ZoomTable:=tzJournal ; END else
        if FType='IPM' then BEGIN OkTest:=FALSE ; END Else
           BEGIN CCo:=TConcept(ccSecModif) ; Cache.ZoomTable:=AxeToTz(Copy(FType,2,2)) ; END ;
if OkTest And (Not ExJaiLeDroitConcept(CCo,False)) then A:=taConsult else A:=taModif ;
Result:=TRUE ;
END ;

procedure TFCorrImp.ClickZoom ;
Var A   : TActionFiche ;
    St,St1 : String ;
begin
If Not OkZoom(A) Then Exit ;
If FType='IPM' Then
  BEGIN
  St:=Trim(FListe.SelectedField.Text) ;
  St1:=FicheModePaieZoom_AGL(St) ;
  If St1<>'' Then
    BEGIN
    If TCorresp.State=dsBrowse Then TCorresp.Edit ;
    FListe.Fields[FListe.SelectedIndex].AsString:=St1 ;
    END ;
  END Else
  BEGIN
  Cache.Text:=Trim(FListe.SelectedField.Text) ;
  if GChercheCompte(Cache,Nil) then
    BEGIN
    If TCorresp.State=dsBrowse Then TCorresp.Edit ;
    FListe.Fields[FListe.SelectedIndex].AsString:=Cache.Text ;
    END ;
  END ;
end;

procedure TFCorrImp.ClickZoomSousPlan ;
Var
    A   : TActionFiche ;
    fb  : TFichierBase ;
    StS : String ;
begin
If FType='IPM' Then Exit ;
If Not OkZoom(A) Then Exit ;
If Ftype[3] In ['1'..'5']=FALSE Then Exit ;
fb:=AxeTofb('A'+Ftype[3]) ;
If VH^.Cpta[fb].Structure Then
   BEGIN
   StS:=Trim(FListe.SelectedField.Text) ;
   If ChoisirSousPlan(fb,StS,TRUE,A) Then
      If (A in [taCreat,taModif]) Then
         BEGIN
         If TCorresp.State=dsBrowse Then TCorresp.Edit ;
         FListe.Fields[FListe.SelectedIndex].AsString:=Cache.Text ;
         END ;
   END ;
end;


procedure TFCorrImp.FListeDblClick(Sender: TObject);
begin
ClickZoom ;
end;


procedure TFCorrImp.bExportClick(Sender: TObject);
begin
if not ExJaiLeDroitConcept(ccExportListe,True) then exit ;
SD.Options:=SD.Options+[ofOverwritePrompt] ;
SD.Title:=TraduireMemoire('Export') ;
if SD.Execute then ExportGrid(Nil,FListe,SD.FileName,SD.FilterIndex,TRUE) ;
end;

Procedure ImportCorresp(StFichier,SType : String ; OkDel : Boolean) ;
Var F : TextFile ;
    Cod,Lib,St : String ;
    Q : TQuery ;
BEGIN
Assign(F,StFichier) ;
{$i-} Reset(F,StFichier) ; {$i+}
If Ioresult<>0 Then
  BEGIN
  Exit ;
  END ;
InitMove(100,'') ;
If OkDel Then ExecuteSQL('DELETE FROM CORRESP WHERE CR_TYPE="'+SType+'"') ;
While Not Eof(F) Do
  BEGIN
  Readln(F,St) ; St:=Trim(St) ;
  MoveCur(FALSE) ;
  If St<>'' Then
    BEGIN
    Cod:='' ; Lib:='' ;
    If St[Length(St)]<>';' Then St:=St+';' ;
    Cod:=ReadTokenSt(St) ;
{$IFDEF RECUPLTL}
    If VH^.RecupLTL Then
      BEGIN
      if SType='IGE' then Cod:=BourreOuTronque(Cod,fbGene) else
         if SType='IAU' then Cod:=BourreOuTronque(Cod,fbAux) ;
      END;
{$ELSE}
{$IFDEF RECUPSISCOS3}
    If VH^.RecupSISCOS3 Then
      BEGIN
      if SType='IGE' then Cod:=BourreOuTronque(Cod,fbGene) else
         if SType='IAU' then Cod:=BourreOuTronque(Cod,fbAux) ;
      END;
{$ENDIF}
{$ENDIF}
    If St<>'' Then Lib:=ReadTokenSt(St) ;
    If Lib<>'' Then
      BEGIN
      Q:=OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="'+SType+'" AND CR_CORRESP="'+Cod+'" ',FALSE) ;
      If Q.Eof Then
        BEGIN
        Q.Insert ; InitNew(Q) ;
        Q.FindField('CR_TYPE').AsString:=SType ;
        Q.FindField('CR_CORRESP').AsString:=Cod ;
        Q.FindField('CR_LIBELLE').AsString:=Lib ;
        Q.Post ;
        END Else
        BEGIN
        Q.Edit ;
        Q.FindField('CR_LIBELLE').AsString:=Lib ;
        Q.Post ;
        END ;
      Ferme(Q) ;
      END ;
    END ;
  END ;
CloseFile(F) ;
FiniMove ;
END ;


procedure TFCorrImp.bImportClick(Sender: TObject);
Var OkDel : Boolean ;
    mr : tmodalResult ;

begin
SD.Title:=TraduireMemoire('Import') ;
SD.Options:=SD.Options-[ofOverwritePrompt] ;
if SD.Execute then
  BEGIN
  mr:=MsgBox.Execute(23,'','') ;
  if mr=mrCancel Then Exit ;
  OkDel:=mr=mrYes ;
  ImportCorresp(SD.FileName,FType,OkDel) ;
  TCorresp.Close ; TCorresp.Open ;
  TCorresp.SetRange([FType],[FType]) ;
  END ;
end;

end.
