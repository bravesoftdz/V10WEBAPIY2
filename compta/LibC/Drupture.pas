unit DRUPTURE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBCtrls, Grids, DBGrids, HDB, StdCtrls, Buttons, ExtCtrls,
  hmsgbox, Ent1, ComCtrls, Spin, Hctrls, Hent1, HCompte, HSysMenu, Hqry,
  ADODB;

Procedure DetailPlanRupture(LaNature, LePlan : String ; Mode : TActionFiche) ;

type
  TFDrupture = class(TForm)
    HPB      : TPanel;
    FAutoSave: TCheckBox;
    BAnnuler: THBitBtn;
    BFirst: THBitBtn;
    BPrev: THBitBtn;
    BNext: THBitBtn;
    BLast: THBitBtn;
    BInsert: THBitBtn;
    BDelete: THBitBtn;
    BRuptanal: THBitBtn;
    SRupture : TDataSource;
    TRupture: THTable;
    FListe   : THDBGrid;
    DBNav    : TDBNavigator;
    MsgBox   : THMsgBox;
    BAutomate: THBitBtn;
    QRupture : TQuery;
    PAuto    : TPanel;
    GBAuto   : TGroupBox;
    TEClasse1: THLabel;
    TEClasse2: THLabel;
    TEClasse3: THLabel;
    TEClasse4: THLabel;
    TEClasse5: THLabel;
    TEClasse6: THLabel;
    EClasse1 : TEdit;
    EClasse2 : TEdit;
    EClasse3 : TEdit;
    EClasse4 : TEdit;
    EClasse5 : TEdit;
    EClasse6 : TEdit;
    GbAuto1  : TGroupBox;
    Classe1  : TCheckBox;
    Classe3  : TCheckBox;
    Classe5  : TCheckBox;
    Classe6  : TCheckBox;
    Classe4  : TCheckBox;
    Classe2  : TCheckBox;
    Panel1   : TPanel;
    LClasse  : TLabel;
    SEAuto   : TSpinEdit;
    TRuptureRU_NATURERUPT    : TStringField;
    TRuptureRU_CLASSE        : TStringField;
    TRuptureRU_LIBELLECLASSE : TStringField;
    TRuptureRU_SOCIETE       : TStringField;
    TRuptureRU_PLANRUPT      : TStringField;
    Panel2: TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BImprimer: THBitBtn;
    HMTrad: THSystemMenu;
    BZoomSp: THBitBtn;
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure SRuptureDataChange(Sender: TObject; Field: TField);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TRuptureNewRecord(DataSet: TDataSet);
    procedure BImprimerClick(Sender: TObject);
    procedure BAutomateClick(Sender: TObject);
    procedure SEAutoChange(Sender: TObject);
    procedure BRuptanalClick(Sender: TObject);
    procedure SRuptureUpdateData(Sender: TObject);
    procedure TRuptureBeforePost(DataSet: TDataSet);
    procedure TRupturePostError(DataSet: TDataSet; E: EDatabaseError; var Action: TDataAction);
    procedure FListeRowEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BZoomSpClick(Sender: TObject);
  private
   Nature   : String ;
   Plan     : String ;
   Axe      : String ;
   Mode     : TActionFiche ;
   WMinX,WMinY : Integer ;
   procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
   Function  Bouge(Button: TNavigateBtn) : boolean ;
   Function  OnSauve : boolean ;
   Procedure NewEnreg ;
   Procedure ChargeEnreg ;
   Function  Supprime : Boolean ;
   Function  EnregOK : boolean ;
   Function  VerifiSiExiste : Boolean ;
   Function  PresenceEnChainements(Var St : String) : Boolean ;
   Procedure FaitLeCaption ;
   Function  GenereRupAnal : Boolean ;
   Procedure GeleLesBoutons ;
   Procedure BrancheHelpContext ;
   Procedure PautoVisible(Avec : Boolean) ;
  public
    { Déclarations publiques }
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  printDBG, Rupanal, DRupanal,
  HZoomSp,
  DRupGene, UtilPgi ;

{$R *.DFM}


procedure TFDrupture.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
//if Pauto.Visible then BAFermeClick(Nil) ;
BFerme.SetFocus ; CanClose:=OnSauve ;
end;

procedure TFDrupture.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFDrupture.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFDrupture.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFDrupture.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFDrupture.BInsertClick(Sender: TObject);
begin Bouge(nbInsert) ; end;

procedure TFDrupture.BDeleteClick(Sender: TObject);
begin Bouge(nbDelete) ; end;

procedure TFDrupture.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFDrupture.BValiderClick(Sender: TObject);
Var RupClass : TabBool ;
    LibClass : TabLib ;
    i : Byte ;
begin
if Not Pauto.Visible then Bouge(nbPost) else
   BEGIN
   if SEAuto.Value<=0 then Exit ;
   for i:=1 to 6 do
       BEGIN
       if TCheckBox(FindComponent('Classe'+InttoStr(i))).Checked then RupClass[i]:=True
                                                                 else RupClass[i]:=False ;
       LibClass[i]:=TEdit(FindComponent('EClasse'+InttoStr(i))).Text ;
       END ;
   DetailPlanRuptureGene(RupClass,LibClass,Nature,Plan) ;
   PautoVisible(False) ;
   if V_PGI.Driver=dbMSACCESS then BEGIN TRupture.Close ; TRupture.Open ; TRupture.SetRange([Nature,Plan],[Nature,Plan]) ; END
                           else BEGIN TRupture.Refresh ; TRupture.First ; END ;
   FListe.SetFocus ;
   END ;
end;

procedure TFDrupture.BFermeClick(Sender: TObject);
begin
if Not Pauto.Visible then Close else
   BEGIN
   PautoVisible(False) ;
   if V_PGI.Driver=dbMSACCESS then BEGIN TRupture.Close ; TRupture.Open ; TRupture.SetRange([Nature,Plan],[Nature,Plan]) ; END
                           else BEGIN TRupture.Refresh ; TRupture.First ; END ;
   BrancheHelpContext ;
   END ;
end;

Function TFDrupture.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
     nblast,nbprior,nbnext,nbfirst,
     nbinsert : if Not OnSauve  then Exit ;
     nbPost   : if Not EnregOK  then Exit ;
     nbDelete : if Not Supprime then Exit ;
    end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[8]) ;
result:=TRUE ;
if Button=NbInsert then NewEnreg ;
END ;

Function TFDrupture.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if TRupture.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.Execute(0,'','')  ;
   END else Rep:=321 ;
Case rep of
 mrYes : if not Bouge(nbPost) then exit ;
 mrNo  : if not Bouge(nbCancel) then exit ;
 mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function  TFDrupture.Supprime : Boolean ;
BEGIN
Result:=False ;
if MsgBox.Execute(1,'','')=mrYES then
   BEGIN if NOT((TRupture.EOF)And(TRupture.BOF)) then Result:=True END ;
END ;

Procedure TFDrupture.ChargeEnreg ;
BEGIN
(******)
END ;

procedure TFDrupture.TRuptureNewRecord(DataSet: TDataSet);
begin NewEnreg ; end;

Procedure TFDrupture.NewEnreg ;
BEGIN
InitNew(TRupture) ;
TRuptureRU_NATURERUPT.AsString:=Nature ; TRuptureRU_PLANRUPT.AsString:=Plan ;
FListe.Columns.Items[0].ReadOnly:=False ;
Fliste.SelectedIndex:=0 ;
END ;

Function TFDrupture.EnregOK : boolean ;
BEGIN
result:=FALSE  ;
if TRupture.state in [dsEdit,dsInsert]=False then Exit ;
if TRupture.state in [dsEdit,dsInsert] then
   BEGIN
   if TRuptureRU_CLASSE.AsString='' then
      BEGIN MsgBox.Execute(2,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
   if TRuptureRU_LIBELLECLASSE.AsString='' then
      BEGIN MsgBox.Execute(3,'','') ; FListe.SelectedIndex:=1 ;FListe.SetFocus ; Exit ; END ;
   END ;
if TRupture.state in [dsInsert] then
   BEGIN
    if VerifiSiExiste then
       BEGIN MsgBox.Execute(4,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
   END ;
if (Copy(TRuptureRU_CLASSE.AsString,Length(TRuptureRU_CLASSE.AsString),1)<>'x') then
   TRuptureRU_CLASSE.AsString:=TRuptureRU_CLASSE.AsString+'x' ;
Result:=TRUE  ;
END ;

Function TFDRupture.VerifiSiExiste : Boolean ;
BEGIN
//Simon "
QRupture.Close ; QRupture.SQL.Clear ;
QRupture.SQL.Add('Select RU_NATURERUPT,RU_PLANRUPT,RU_CLASSE from RUPTURE ') ;
QRupture.SQL.Add('Where RU_NATURERUPT="'+Nature+'" AND RU_PLANRUPT="'+Plan+'" ') ;
QRupture.SQL.Add('AND RU_CLASSE="'+TRuptureRU_CLASSE.AsString+'"') ;
ChangeSQL(QRupture) ; QRupture.Open ;
Result:=(not QRupture.EOF) ;
QRupture.Close ;
END ;

procedure TFDrupture.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if(ssCtrl in Shift)AND(Key=VK_DELETE)AND(TRupture.Eof)AND(TRupture.Bof)then BEGIN Key:=0 ; Exit ; END ;
if(ssCtrl in Shift)AND(Key=VK_DELETE)then BEGIN Bouge(nbDelete) ; Key:=0 ; END ;
end;

Procedure TFDrupture.GeleLesBoutons ;
BEGIN
BInsert.Enabled:=Not(TRupture.State in [dsInsert,dsEdit]) ;
BDelete.Enabled:=Not (TRupture.State in [dsInsert,dsEdit]) ;
BAutomate.Enabled:=Not (TRupture.State in [dsInsert,dsEdit]) ;
if(TRupture.EOF)And(TRupture.BOF)then BDelete.Enabled:=False ;
if TRupture.State=dsBrowse then FListe.Columns.Items[0].ReadOnly:=True ;
END ;

procedure TFDrupture.SRuptureDataChange(Sender: TObject; Field: TField);
Var UpEnable, DnEnable: Boolean;
begin
GeleLesBoutons ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not TRupture.BOF;
   DnEnable := Enabled and not TRupture.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
//   if ((Field.FieldName='RU_LIBELLE') and (RU_ABREGE.Field.AsString='')) then
//   RU_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

Procedure TFDrupture.FaitLeCaption ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CC_LIBELLE From CHOIXCOD Where CC_TYPE="'+Nature+'" And CC_CODE="'+Plan+'"',True) ;
Case Nature[3] of
     'G':Caption:=Caption+MsgBox.Mess[9] +':'+Plan+' '+Q.Fields[0].AsString   ;
     'T':Caption:=Caption+MsgBox.Mess[10]+':'+Plan+' '+Q.Fields[0].AsString  ;
     '1':Caption:=Caption+MsgBox.Mess[11]+':'+Plan+' '+Q.Fields[0].AsString  ;
     '2':Caption:=Caption+MsgBox.Mess[12]+':'+Plan+' '+Q.Fields[0].AsString  ;
     '3':Caption:=Caption+MsgBox.Mess[13]+':'+Plan+' '+Q.Fields[0].AsString  ;
     '4':Caption:=Caption+MsgBox.Mess[14]+':'+Plan+' '+Q.Fields[0].AsString  ;
     '5':Caption:=Caption+MsgBox.Mess[15]+':'+Plan+' '+Q.Fields[0].AsString  ;
     'B':Caption:=Caption+MsgBox.Mess[17]+':'+Plan+' '+Q.Fields[0].AsString  ;
  End ;
Ferme(Q) ;
END ;

procedure TFDrupture.FormShow(Sender: TObject);
Var Lg,i : Byte ;
    St : String17 ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; BrancheHelpContext ;
TRupture.Open ; TRupture.SetRange([Nature,Plan],[Nature,Plan]) ;
PAuto.Visible:=False ;
BRuptanal.Enabled:=((Nature<>'RUT') AND (Nature<>'RUG')) ;
BZoomSp.Enabled:=((Nature<>'RUT') AND (Nature<>'RUG')) ;
if Nature='RUT'then Lg:=VH^.Cpta[fbAux].Lg else
   if Nature='RUG'then Lg:=VH^.Cpta[fbGene].Lg  else
      BEGIN Axe:='A'+Nature[3] ; Lg:= VH^.Cpta[AxeToFb(Axe)].Lg ; END ;
St:='' ;
for i:=1 to Lg do St:=St+'a' ;
TRuptureRU_CLASSE.EditMask:='>'+St+'<>a;0; ';
FaitLeCaption ;
if Mode=taCreat then BAutomateClick(Nil) ;
if TRupture.Eof And TRupture.Bof then Bouge(nbInsert) ;
end;

Procedure DetailPlanRupture(LaNature,LePlan : String ; Mode : TActionFiche) ;
var Frupture: TFDrupture;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
Frupture:=TFDrupture.Create(Application) ;
  Try
   Frupture.Nature:=LaNature ;
   Frupture.Plan:=LePlan ;
   Frupture.Mode:=Mode ;
   Frupture.ShowModal ;
  Finally
   Frupture.Free ;
  End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFDrupture.BImprimerClick(Sender: TObject);
begin PrintDBGrid (FListe,Nil,Caption,''); end;


//******------ Pour l'Automate ---------************

procedure TFDrupture.BAutomateClick(Sender: TObject);
Var St : String ;
begin
Width:=460 ; Height:=323 ;
if (Nature='RUT')OR(Nature='RUG') then
   BEGIN
   PautoVisible(True) ; SEAuto.Value:=0 ; SEAutoChange(Nil) ;
   Case Nature[3] of
        'G' : HelpContext:=1370200 ;
        'T' : HelpContext:=1380200 ;
      else HelpContext:=0 ;
     End ;
   END else
   BEGIN
   if VH^.Cpta[AxeToFb(Axe)].Structure then
      if PresenceEnChainements(St) then GenereRupAnal 
                                   else MsgBox.Execute(7,'','') ;
   END ;
end;

procedure TFDrupture.SEAutoChange(Sender: TObject);
Var i : Integer ;
    C : TComponent ;
begin
for i:=1 to 6 do
    BEGIN
    C:=FindComponent('Classe'+InttoStr(i))  ; TControl(C).Enabled:=(i<=SEAuto.Value) ;
    if Not TControl(C).Enabled then TCheckBox(C).Checked:=False ;
    C:=FindComponent('EClasse'+InttoStr(i)) ; TControl(C).Enabled:=(i<=SEAuto.Value) ;
    if Not TControl(C).Enabled then TEdit(C).Text:='' ;
    C:=FindComponent('TEClasse'+InttoStr(i)) ;TControl(C).Enabled:=(i<=SEAuto.Value) ;
    END ;
end;

Function TFDrupture.PresenceEnChainements(Var St : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CC_LIBRE from CHOIXCOD Where CC_TYPE="'+Nature+'" AND '+
           'CC_CODE="'+Plan+'"',True) ;
St:=Q.Fields[0].AsString ; Ferme(Q) ; Result:=(St<>'') ;
END ;

Function TFDrupture.GenereRupAnal : Boolean ;
BEGIN
Result:=False ;
if DetailPlanRuptureAnal(Nature,Plan,Axe,'') then
   BEGIN
   TRupture.Close ; TRupture.Open ;
   TRupture.SetRange([Nature,Plan],[Nature,Plan]) ;
   Result:=True ;
   END ;
END ;

procedure TFDrupture.BRuptanalClick(Sender: TObject);
Var CLibreMemo,CLibre : String ;
begin
PresenceEnChainements(CLibreMemo) ; CLibre:=CLibreMemo ;
RuptureAnalytique(Nature,Plan,Clibre) ;
if CLibre<>CLibreMemo then
   BEGIN
    if MsgBox.Execute(6,'','')=mrYes then
       BEGIN
       BeginTrans ;
       ExecuteSql('UPDATE CHOIXCOD SET CC_LIBRE="'+CLibre+'" Where CC_TYPE="'+Nature+'" AND '+
                  'CC_CODE="'+Plan+'"') ;
       CommitTrans ;
       if Not GenereRupAnal then
          BEGIN
          If CLibreMemo<>'' Then { GP le 23/10/98 }
             BEGIN
             MsgBox.Execute(16,'','') ;
             BeginTrans ;
             ExecuteSql('UPDATE CHOIXCOD SET CC_LIBRE="'+CLibreMemo+'" Where CC_TYPE="'+Nature+'" AND '+
                        'CC_CODE="'+Plan+'"') ;
             CommitTrans ;
             END ;
          END ;
       END ;
   END ;
end;

procedure TFDrupture.SRuptureUpdateData(Sender: TObject);
begin
if(Trim(TRuptureRU_CLASSE.AsString)='') And (Trim(TRuptureRU_LIBELLECLASSE.AsString)='') then
   BEGIN TRupture.Cancel ; Exit ; END ;
if(Trim(TRuptureRU_CLASSE.AsString)='') And (Trim(TRuptureRU_LIBELLECLASSE.AsString)<>'') then
  BEGIN MsgBox.Execute(2,'','') ; Fliste.SelectedIndex:=0 ; SysUtils.Abort ; Exit ; END ;
if(Trim(TRuptureRU_CLASSE.AsString)<>'') And (Trim(TRuptureRU_LIBELLECLASSE.AsString)='') then
  BEGIN MsgBox.Execute(3,'','') ; Fliste.SelectedIndex:=1 ; SysUtils.Abort ; Exit ; END ;
end;

procedure TFDrupture.TRuptureBeforePost(DataSet: TDataSet);
begin
if (Copy(TRuptureRU_CLASSE.AsString,Length(TRuptureRU_CLASSE.AsString),1)<>'x') then
   TRuptureRU_CLASSE.AsString:=TRuptureRU_CLASSE.AsString+'x' ;
end;

procedure TFDrupture.TRupturePostError(DataSet: TDataSet;
  E: EDatabaseError; var Action: TDataAction);
begin
if TRupture.State=dsInsert then
   BEGIN MsgBox.Execute(4,'','') ; FListe.SelectedIndex:=0 ; Action:=daAbort ; END ;
end;

procedure TFDrupture.FListeRowEnter(Sender: TObject);
begin if TRupture.State=dsInsert then FListe.SelectedIndex:=0 ; end;

procedure TFDrupture.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
if PAuto.Visible then
   with PMinMaxInfo(MSG.lparam)^.ptMaxTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFDrupture.FormCreate(Sender: TObject);
begin
WMinX:=Width ; WMinY:=Height ;
end;

procedure TFDrupture.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

Procedure TFDrupture.BrancheHelpContext ;
BEGIN
Case Nature[3] of
     'G'     : HelpContext:=1370100 ;
     'T'     : HelpContext:=1380100 ;
     '1'..'5': HelpContext:=1395100 ;
     'B'     : HelpContext:=0 ;
   else        HelpContext:=0 ;
  End ;
END ;

Procedure TFDrupture.PautoVisible(Avec : Boolean) ;
BEGIN
BFirst.Visible:=Not Avec ; BPrev.Visible:=Not Avec ; BNext.Visible:=Not Avec ;
BLast.Visible:=Not Avec ;  BAnnuler.Visible:=Not Avec ; BInsert.Visible:=Not Avec ;
BDelete.Visible:=Not Avec ; BRuptanal.Visible:=Not Avec ; BAutomate.Visible:=Not Avec ;
BImprimer.Visible:=Not Avec ; FListe.Visible:=Not Avec ; PAuto.Visible:=Avec ;
END ;

procedure TFDrupture.BZoomSpClick(Sender: TObject);
Var UnFb : TFichierBase ;
    St : String;
    HSt : hString;
    St1,StSp : String ;
begin
{$IFNDEF CCS3}
if Mode=taConsult then Exit ;
UnFb:=AxeToFb('A'+Nature[3]) ;
if TRupture.State=dsInsert then
  St:=''
else
  St:=TRuptureRU_CLASSE.AsString ;

  HSt := St;

if ChoisirSousPlan(UnFb,St,True,taModif) then
   BEGIN
   if Pos('?',St)<=0 then
      BEGIN
      if TRupture.State=dsBrowse then TRupture.Insert ;
      TRuptureRU_CLASSE.AsString:=St+'x' ;
      END else
      BEGIN
      StSp:='' ;
      While St<>'' do
          BEGIN
          St1:=ReadTokenInterro(HSt) ;
          if St1<>'' then StSp:=St1+';' ;
          END ;
      if DetailPlanRuptureAnal(Nature,Plan,Axe,StSp) then
         BEGIN
         TRupture.Close ; TRupture.Open ;
         TRupture.SetRange([Nature,Plan],[Nature,Plan]) ;
         END ;
      END ;
   END ;
{$ENDIF}
end;

end.
