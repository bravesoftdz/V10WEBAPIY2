unit EcheMono;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,HSysMenu,
  Forms, Dialogs, StdCtrls, Mask, Hctrls, Buttons, ExtCtrls, hmsgbox,Ent1,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  HEnt1;

Type T_MONOECH = RECORD
                 ModePaie,Cat : String3 ;
                 NumTraiteChq : String17 ;
                 DateMvt : TDateTime ;
                 DateValeur   : TDateTime ;
                 DateEche     : TDateTime ;
                 Treso,OkInit,OkVAL : boolean ;
                 Action             : TActionFiche ;
                 END ;

Function  SaisirMonoEcheance ( Var X : T_MONOECH ; Plus : String = '') : Boolean ;

type
  TFMonoEche = class(TForm)
    FModePaie: THValComboBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    FDateEche: TMaskEdit;
    BValider: THBitBtn;
    BAnnuler: THBitBtn;
    Baide: THBitBtn;
    HLabel3: THLabel;
    FDatevaleur: TMaskEdit;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    HNumTraiteCHQ: THLabel;
    FNumTraiteCHQ: TMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FDatevaleurExit(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BaideClick(Sender: TObject);
    procedure FModePaieChange(Sender: TObject);
    procedure FDateEcheExit(Sender: TObject);
  private
  public
    Treso,OkInit,OkVal : Boolean ;
    Action             : TActionFiche ;
    EcheInit,ValInit,DateMvt : TDateTime ;
  end;

implementation

{$R *.DFM}

Uses SaisUtil ;

procedure ChargeMode ( FC : THValComboBox ; Cat : String ) ;
Var Q : TQuery ;
BEGIN
FC.Values.Clear ; FC.Items.Clear ;
Q:=OpenSQL('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE Where MP_CATEGORIE="'+Cat+'"',True) ;
While Not Q.EOF do
   BEGIN
   FC.Values.Add(Q.FindField('MP_MODEPAIE').asString) ;
   FC.Items.Add(Q.FindField('MP_LIBELLE').asString) ;
   Q.Next ;
   END ;
Ferme(Q) ;
END ;

Function SaisirMonoEcheance ( Var X : T_MONOECH ; Plus : String = '') : Boolean ;
var FMonoEche : TFMonoEche ;
BEGIN
FMonoEche:=TFMonoEche.Create(Application) ;
try
  if X.Cat='' then FMonoEche.FModePaie.DataType:='ttModePaie' else ChargeMode(FMonoEche.FModePaie,X.Cat) ;
  FMonoEche.FmodePaie.Plus:=Plus ;
  FMonoEche.FModePaie.Value:=X.ModePaie ;
  FMonoEche.OkInit:=X.OkInit ; FMonoEche.OkVal:=X.OkVal ;
  FMonoEche.Action:=X.Action ;
  FMonoEche.FDateEche.Text:=DateToStr(X.DateEche) ;
  FMonoEche.FDateValeur.Text:=DateToStr(X.DateValeur) ;
  FMonoEche.EcheInit:=X.DateEche ;
  FMonoEche.ValInit:=X.DateValeur ;
  FMonoEche.DateMvt:=X.DateMvt ;
  FMonoEche.FNumTraiteCHQ.Text:=X.NumTraiteCHQ ;
  { GP le 08/07/97 N° 1834
  if X.Treso then FMonoEche.Height:=FMonoEche.Height-FMonoEche.BValider.Height-8 ;
  }
  if Not X.OkVal then FMonoEche.FDateValeur.Enabled:=False ;
  Result:=(FMonoEche.ShowModal=mrOK) ;
  if Result then
     BEGIN
     X.ModePaie:=FMonoEche.FModePaie.Value ;
     X.DateEche:=StrToDate(FMonoEche.FDateEche.Text) ;
     X.DateValeur:=StrToDate(FMonoEche.FDateValeur.Text) ;
     X.NumTraiteCHQ:=FMonoEche.FNumTraiteCHQ.Text ;
     END ;
  finally
  FMonoEche.Free ;
  end;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFMonoEche.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
if Action=taConsult then
   BEGIN
   FModePaie.Enabled:=False ;
   FDateEche.Enabled:=False ;
   FDateValeur.Enabled:=False ;
   FNumTraiteCHQ.Enabled:=False ;
   END ;
if Not Treso then Exit ;
BorderIcons:=[] ; BorderStyle:=bsNone ; Caption:='' ;
BValider.Visible:=False ; BAnnuler.Visible:=False ; BAide.Visible:=False ;
BValider.Enabled:=False ; BAnnuler.Enabled:=False ; BAide.Enabled:=False ;
end;

procedure TFMonoEche.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Shift<>[] then Exit ;
if Key=VK_F10 then BEGIN Key:=0 ; BValiderClick(Nil) ; END else
 if Key=VK_ESCAPE then BEGIN Key:=0 ; ModalResult:=mrCancel ; END ;
end;

procedure TFMonoEche.FDateValeurExit(Sender: TObject);
//Var DD : TDateTime ;
begin
//DD:=StrToDate(FDateValeur.Text) ;
if Treso then BValider.Click ;
end;

procedure TFMonoEche.BValiderClick(Sender: TObject);
Var DD : tDateTime ;
begin
if ((FModePaie.Value='') and (Action<>taConsult)) then BEGIN HM.Execute(0,'','') ; Exit ; END ;
If DateMvt<>0 Then
  BEGIN
  DD:=StrToDate(FDateEche.Text) ;
  if Not NbJoursOK(DateMvt,DD) then BEGIN HM.Execute(2,'','') ; Exit ; END ;
  END ;
if Not IsValidDate(FDateEche.Text) then FDateEche.Text:=DateToStr(EcheInit) ;
if Not IsValidDate(FDateValeur.Text) then FDateValeur.Text:=DateToStr(ValInit) ;
ModalResult:=mrOk ;
end;

procedure TFMonoEche.BAnnulerClick(Sender: TObject);
begin
if ((Not OkInit) or (Action=taConsult)) then ModalResult:=mrCancel else HM.Execute(1,'','') ;
end;

procedure TFMonoEche.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if ((OkInit) and (ModalResult<>mrOk) and (Action<>taConsult)) then BEGIN HM.Execute(1,'','') ; CanClose:=False ; END ;
end;

procedure TFMonoEche.BaideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFMonoEche.FModePaieChange(Sender: TObject);
Var sCat : String ;
begin
sCat:=MPToCategorie(FModePaie.Value) ;
if sCat='CHQ' then
   BEGIN
   HNumTraiteCHQ.Caption:=HM.Mess[3] ;
   FNumTraiteCHQ.Enabled:=True ;
   HNumTraiteCHQ.Enabled:=True ;
   END else if sCat='LCR' then
   BEGIN
   HNumTraiteCHQ.Caption:=HM.Mess[4] ;
   FNumTraiteCHQ.Enabled:=True ;
   HNumTraiteCHQ.Enabled:=True ;
   END else
   BEGIN
   HNumTraiteCHQ.Caption:=HM.Mess[4] ;
   FNumTraiteCHQ.Enabled:=False ;
   HNumTraiteCHQ.Enabled:=False ;
   END ;
end;

procedure TFMonoEche.FDateEcheExit(Sender: TObject);
//Var DD : TDateTime ;
begin
//DD:=StrToDate(FDateEche.Text) ;
end;

end.
