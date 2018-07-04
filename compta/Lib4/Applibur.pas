unit AppliBur;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, hmsgbox, DB, {$IFNDEF DBXPRESS}{$IFNDEF DBXPRESS}dbtables,
  StdCtrls, Mask, DBCtrls, Hctrls, HSysMenu, Hqry, HTB97, ExtCtrls, HPanel,
  Grids, DBGrids, HDB{$ELSE}uDbxDataSet{$ENDIF}, StdCtrls, Mask,
  DBCtrls, Hctrls, HSysMenu, Hqry, HTB97, ExtCtrls, HPanel, Grids, DBGrids,
  HDB{$ELSE}uDbxDataSet{$ENDIF}, DBCtrls, StdCtrls, Buttons, ExtCtrls,Menus,
  Grids, DBGrids, HDB, Mask, Hctrls, HEnt1, ComCtrls, Inifiles,Ent1,
  HSysMenu, HTB97, Hqry, HPanel, UiUtil ;

type TAppShort = (asOpen,asSave,asSaveAs,asPrint,asReplace,asNew,asYes) ;
     TAppliParam   = RECORD
        Executable : String ;
        Fenetre    : String ;
        ShortName  : Boolean ;
        Touche     : Array [TAppShort] of String ;
        END ;

Function GetAppliParam ( App : String) : TAppliParam ;
Procedure ParamAppliBureautique ;

type
  TFAppliBur = class(TFFicheListe)
    TCC_CODE: THLabel;
    CC_CODE: TDBEdit;
    TCC_LIBELLE: THLabel;
    CC_LIBELLE: TDBEdit;
    TaCC_TYPE: TStringField;
    TaCC_CODE: TStringField;
    TaCC_LIBELLE: TStringField;
    TaCC_ABREGE: TStringField;
    TaCC_LIBRE: TStringField;
    FExecutable: TEdit;
    TExecutable: THLabel;
    BFile: TToolbarButton97;
    OpenDialog: TOpenDialog;
    TFenetre: THLabel;
    FFenetre: TEdit;
    BClass: TToolbarButton97;
    TOuvrir: THLabel;
    TEnregistrer: THLabel;
    TSous: THLabel;
    TImprimer: THLabel;
    TRemplacer: THLabel;
    TNouveau: THLabel;
    FOpen: TEdit;
    FSave: TEdit;
    FSaveAs: TEdit;
    FPrint: TEdit;
    FReplace: TEdit;
    FNew: TEdit;
    FShort: TCheckBox;
    HLabel1: THLabel;
    FYes: TEdit;
    procedure FormShow(Sender: TObject);
    procedure BFileClick(Sender: TObject);
    procedure BClassClick(Sender: TObject);
    procedure FExecutableKeyPress(Sender: TObject; var Key: Char);
    procedure FShortClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
  private { Déclarations privées }
    Updating : boolean ;
    Function EnregOK : boolean ; Override ;
    Procedure SauveINI ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
    Function VerifiSiExiste : Boolean ;
  public  { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function GetAppliParam ( App : String) : TAppliParam ;
Var IniFile : TIniFile ;
    T : TAppliParam ;
begin
IniFile :=TIniFile.Create('HALMACH.INI');
T.Executable:=IniFile.ReadString(App,'X','');
T.Fenetre:=IniFile.ReadString(App,'F','');
T.ShortName:=(IniFile.ReadString(App,'8','-')='X') ;
T.Touche[asOpen]   :=IniFile.ReadString(App,'O','') ;
T.Touche[asSave]   :=IniFile.ReadString(App,'E','') ;
T.Touche[asSaveAs] :=IniFile.ReadString(App,'S','') ;
T.Touche[asPrint]  :=IniFile.ReadString(App,'I','') ;
T.Touche[asReplace]:=IniFile.ReadString(App,'R','') ;
T.Touche[asNew]    :=IniFile.ReadString(App,'N','') ;
T.Touche[asYes]    :=IniFile.ReadString(App,'Y','') ;
IniFile.Free;
Result:=T ;
END ;

Procedure ParamAppliBureautique ;
var FAppliBur: TFAppliBur;
    PP : THPanel ;
BEGIN
FAppliBur:=TFAppliBur.Create(Application) ;
FAppliBur.InitFL('CC','','','',taModif,TRUE,FAppliBur.TaCC_CODE,FAppliBur.TaCC_LIBELLE,NIL,['ttApplication']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FAppliBur.ShowModal ;
    finally
     FAppliBur.Free ;
    end ;
   END else
   BEGIN
   InitInside(FAppliBur,PP) ;
   FAppliBur.Show ;
   END ;
END ;

procedure TFAppliBur.FormShow(Sender: TObject);
begin
  inherited;
BImprimer.Visible:=False ;
Ta.SetRange(['APP'],['APP']) ;
if(Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then BinsertClick(Nil) ;
end;

Procedure TFAppliBur.SauveINI ;
Var IniFile : TIniFile ;
BEGIN
IniFile :=TIniFile.Create('HALMACH.INI');
IniFile.WriteString(TaCC_CODE.AsString,'X',FExecutable.text);
if FShort.Checked then IniFile.WriteString(TaCC_CODE.AsString,'8','X')
                  else IniFile.WriteString(TaCC_CODE.AsString,'8','-') ;
IniFile.WriteString(TaCC_CODE.AsString,'F',FFenetre.Text);
IniFile.WriteString(TaCC_CODE.AsString,'O',FOpen.Text);
IniFile.WriteString(TaCC_CODE.AsString,'E',FSave.Text);
IniFile.WriteString(TaCC_CODE.AsString,'S',FSaveAs.Text);
IniFile.WriteString(TaCC_CODE.AsString,'I',FPrint.Text);
IniFile.WriteString(TaCC_CODE.AsString,'R',FReplace.Text);
IniFile.WriteString(TaCC_CODE.AsString,'N',FNew.Text);
IniFile.WriteString(TaCC_CODE.AsString,'Y',FYes.Text);
IniFile.Free;
END ;

Procedure TFAppliBur.ChargeEnreg ;
Var tt : TAppliParam ;
BEGIN
InHerited ;
Updating:=TRUE ;
if ta.State=dsInsert then
   BEGIN
   FExecutable.Text:='' ;
   FFenetre.Text:='';
   FShort.Checked:=FALSE ;
   FOpen.Text:='' ;
   FSave.Text:='' ;
   FSaveAs.Text:='' ;
   FPrint.Text:='' ;
   FReplace.Text:='' ;
   FNew.Text:='' ;
   FYes.Text:='' ;
   END else
   if ta.State=dsBrowse then
   BEGIN
   TT:=GetAppliParam(TaCC_CODE.AsString) ;
   FExecutable.Text:=TT.Executable ;
   FFenetre.Text:=TT.Fenetre ;
   FShort.Checked:=TT.ShortName ;
   FOpen.Text:=TT.Touche[asOpen] ;
   FSave.Text:=TT.Touche[asSave] ;
   FSaveAs.Text:=TT.Touche[asSaveAs] ;
   FPrint.Text:=TT.Touche[asPrint] ;
   FReplace.Text:=TT.Touche[asReplace] ;
   FNew.Text:=TT.Touche[asNew] ;
   FYes.Text:=TT.Touche[asYes] ;
   END ;
Updating:=FALSE ;
END ;

Function TFAppliBur.VerifiSiExiste : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CC_CODE From CHOIXCOD Where CC_TYPE="APP" And CC_CODE="'+TaCC_CODE.AsString+'"',True) ;
Result:=Not Q.Eof ; Ferme(Q) ;
END ;

Function TFAppliBur.EnregOK : boolean ;
BEGIN
result:=InHerited EnregOK ; if Not Result then Exit ;
Modifier:=True ;
if (Result) and (Ta.State in [dsEdit,dsInsert]) then
   BEGIN
   Result:=False ;
   if Ta.state in [dsInsert] then
     if VerifiSiExiste then
        BEGIN HM.Execute(4,'','') ; FListe.SelectedIndex:=0 ; FListe.SetFocus ; Exit ; END ;
   if FFenetre.text='' then BEGIN HM2.Execute(2,'','') ; FFenetre.SetFocus ; Exit ; END ;
   END ;
SauveINI ;
Result:=True ; Modifier:=False ;
END ;

Procedure TFAppliBur.NewEnreg ;
BEGIN
InHerited ;
TaCC_TYPE.AsString:='APP' ;
END ;

procedure TFAppliBur.BFileClick(Sender: TObject);
Var Key : Char ;
begin
  inherited;
OpenDialog.FileName:=FExecutable.text ;
if OpenDialog.Execute then FExecutable.text:=OpenDialog.FileName ;
FExecutableKeyPress(Nil,Key);
end;

procedure TFAppliBur.BClassClick(Sender: TObject);
Var Handle : hWnd ;
    St1,St2 : String ;
    Key : Char ;
begin
  inherited;
HM2.Execute(0,'','') ;
Delay(2000) ;
Handle:= GetForegroundWindow;
St1:=GetWindowTitleText(handle) ;
St2:=GetWindowClassName(handle) ;
if HM2.Execute(1,'',St2)=mrOK then FFenetre.Text:=St2 ;
FExecutableKeyPress(Nil,Key);
end;

procedure TFAppliBur.FExecutableKeyPress(Sender: TObject; var Key: Char);
Var Old : String ;
begin
  inherited;
if Updating then exit ;
if Ta.Modified then exit ;
Ta.Edit ;
Old:=TaCC_ABREGE.AsString ;
TaCC_ABREGE.AsString:=TaCC_ABREGE.AsString+'e' ;
TaCC_ABREGE.AsString:=Old ;
end;

procedure TFAppliBur.FShortClick(Sender: TObject);
Var Key: Char ;
begin
  inherited;
FExecutableKeyPress(Nil,Key);
end;

procedure TFAppliBur.BDeleteClick(Sender: TObject);
Var Q : TQuery ;
    AFaire : Boolean ;
begin
Q:=OpenSql('Select NM_APPLICATION From NOTEMOD Where NM_APPLICATION="'+TaCC_CODE.AsString+'"',True) ;
AFaire:=Q.Eof ; Ferme(Q) ;
if Not AFaire then BEGIN HM2.Execute(3,'','') ; Exit ; END ;
  inherited;
end;

end.
