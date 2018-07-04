unit Moderegl;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, {$IFNDEF DBXPRESS}dbtables, Db, DBCtrls, ExtCtrls, Spin, HDB,
  StdCtrls, Hctrls, Mask, ComCtrls, HSysMenu, hmsgbox, Hqry, HTB97, HPanel,
  Grids, DBGrids{$ELSE}uDbxDataSet{$ENDIF}, DB, DBCtrls, ExtCtrls, Spin, HDB, StdCtrls, Hctrls,
  Mask, ComCtrls, hmsgbox, Buttons, Grids, DBGrids, HEnt1, Ent1, HSysMenu,
  Hqry, HTB97, UiUtil, HPanel ;

Function FicheRegle(Quel : String; OkGuide : Boolean ; Comment : TActionFiche): string;

type
  TFModeRegle = class(TFFicheListe)
    PageControl: TPageControl;
    PCar: TTabSheet;
    PEche: TTabSheet;
    HGBEcheances2: TGroupBox;
    TMR_MP1: THLabel;
    TMR_MP2: THLabel;
    TMR_MP3: THLabel;
    TMR_MP4: THLabel;
    TMR_MP5: THLabel;
    TMR_MP6: THLabel;
    TMR_MP7: THLabel;
    TMR_MP8: THLabel;
    TMR_MP9: THLabel;
    TMR_MP10: THLabel;
    TMR_MP11: THLabel;
    TMR_MP12: THLabel;
    Bevel1: TBevel;
    TT_MODEPAIEMENT1: THLabel;
    TT_MODEPAIEMENT2: THLabel;
    TT_POURCENT1: THLabel;
    TT_POURCENT2: THLabel;
    TT_ESC2: THLabel;
    TT_ESC: THLabel;
    MR_MP1: THDBValComboBox;
    MR_MP2: THDBValComboBox;
    MR_MP3: THDBValComboBox;
    MR_MP4: THDBValComboBox;
    MR_MP5: THDBValComboBox;
    MR_MP6: THDBValComboBox;
    MR_MP7: THDBValComboBox;
    MR_MP8: THDBValComboBox;
    MR_MP9: THDBValComboBox;
    MR_MP10: THDBValComboBox;
    MR_MP11: THDBValComboBox;
    MR_MP12: THDBValComboBox;
    MR_TAUX1: TDBEdit;
    MR_TAUX2: TDBEdit;
    MR_TAUX3: TDBEdit;
    MR_TAUX4: TDBEdit;
    MR_TAUX5: TDBEdit;
    MR_TAUX6: TDBEdit;
    MR_TAUX7: TDBEdit;
    MR_TAUX8: TDBEdit;
    MR_TAUX9: TDBEdit;
    MR_TAUX10: TDBEdit;
    MR_TAUX11: TDBEdit;
    MR_TAUX12: TDBEdit;
    MR_ESC1: TDBCheckBox;
    MR_ESC2: TDBCheckBox;
    MR_ESC3: TDBCheckBox;
    MR_ESC4: TDBCheckBox;
    MR_ESC5: TDBCheckBox;
    MR_ESC6: TDBCheckBox;
    MR_ESC7: TDBCheckBox;
    MR_ESC8: TDBCheckBox;
    MR_ESC9: TDBCheckBox;
    MR_ESC10: TDBCheckBox;
    MR_ESC11: TDBCheckBox;
    MR_ESC12: TDBCheckBox;
    TaMR_MODEREGLE: TStringField;
    TaMR_LIBELLE: TStringField;
    TaMR_ABREGE: TStringField;
    TaMR_APARTIRDE: TStringField;
    TaMR_PLUSJOUR: TIntegerField;
    TaMR_ARRONDIJOUR: TStringField;
    TaMR_NOMBREECHEANCE: TIntegerField;
    TaMR_SEPAREPAR: TStringField;
    TaMR_MONTANTMIN: TFloatField;
    TaMR_REMPLACEMIN: TStringField;
    TaMR_REPARTECHE: TStringField;
    TaMR_MP1: TStringField;
    TaMR_TAUX1: TFloatField;
    TaMR_ESC1: TStringField;
    TaMR_MP2: TStringField;
    TaMR_TAUX2: TFloatField;
    TaMR_ESC2: TStringField;
    TaMR_MP3: TStringField;
    TaMR_TAUX3: TFloatField;
    TaMR_ESC3: TStringField;
    TaMR_MP4: TStringField;
    TaMR_TAUX4: TFloatField;
    TaMR_ESC4: TStringField;
    TaMR_MP5: TStringField;
    TaMR_TAUX5: TFloatField;
    TaMR_ESC5: TStringField;
    TaMR_MP6: TStringField;
    TaMR_TAUX6: TFloatField;
    TaMR_ESC6: TStringField;
    TaMR_MP7: TStringField;
    TaMR_TAUX7: TFloatField;
    TaMR_ESC7: TStringField;
    TaMR_MP8: TStringField;
    TaMR_TAUX8: TFloatField;
    TaMR_ESC8: TStringField;
    TaMR_MP9: TStringField;
    TaMR_TAUX9: TFloatField;
    TaMR_ESC9: TStringField;
    TaMR_MP10: TStringField;
    TaMR_TAUX10: TFloatField;
    TaMR_ESC10: TStringField;
    TaMR_MP11: TStringField;
    TaMR_TAUX11: TFloatField;
    TaMR_ESC11: TStringField;
    TaMR_MP12: TStringField;
    TaMR_TAUX12: TFloatField;
    TaMR_ESC12: TStringField;
    TaMR_SOCIETE: TStringField;
    TaMR_MODEGUIDE: TStringField;
    TaMR_ECARTJOURS: TStringField;
    TMR_MODEREGLE: THLabel;
    MR_MODEREGLE: TDBEdit;
    MR_SOCIETE: TDBEdit;
    MR_MODEGUIDE: TDBEdit;
    TMR_ABREGE: THLabel;
    MR_ABREGE: TDBEdit;
    TMR_LIBELLE: THLabel;
    MR_LIBELLE: TDBEdit;
    GroupBox1: TGroupBox;
    TMR_APARTIRDE: THLabel;
    TMR_PLUSJOUR: THLabel;
    TMR_ARRONDIJOUR: THLabel;
    TMR_SEPAREPAR: THLabel;
    HLabel1: THLabel;
    TMR_REMPLACEMIN: THLabel;
    MR_APARTIRDE: THDBValComboBox;
    MR_PLUSJOUR: THDBSpinEdit;
    MR_ARRONDIJOUR: THDBValComboBox;
    MR_NOMBREECHEANCE: THDBSpinEdit;
    MR_SEPAREPAR: THDBValComboBox;
    MR_MONTANTMIN: TDBEdit;
    MR_REMPLACEMIN: THDBValComboBox;
    procedure MR_NOMBREECHEANCEChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
  private    { Déclarations privées }
    Guide   : boolean;
    Mode : TActionFiche ;
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    procedure SelectEche ;
    Function  ModePaieOk : Boolean ;
    Function  TotalTauxOK : Boolean ;
    Function  PresenceModeRegle : Boolean ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}


Uses SaisUtil, SaisComm, Echeance ; 

Function FicheRegle(Quel : String; OkGuide : Boolean ; Comment : TActionFiche): string;
var FModeRegle: TFModeRegle;
    PP : THPanel ;
begin
if Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
FicheRegle:=Quel;
FModeRegle:= TFModeRegle.Create(Application) ;
FModeRegle.InitFL('MR','PRT_MODEREGLE',Quel,'',Comment,(Comment<>taConsult),FModeRegle.TAMR_MODEREGLE,
                  FModeRegle.TAMR_LIBELLE,FModeRegle.TAMR_MODEREGLE,['ttModeRegle']) ;
FModeRegle.Guide:=OkGuide ;
FModeRegle.Mode:=Comment ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FModeRegle.ShowModal ;
     FicheRegle:=FModeRegle.FLequel;
    finally
     FModeRegle.Free ;
    end ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FModeRegle,PP) ;
   FModeRegle.Show ;
   END ;
end ;

procedure TFModeRegle.SelectEche ;
Var i : integer ;
    C : TComponent ;
begin
For i:=1 to 12 do
   BEGIN
   C:=FindComponent('MR_MP'+InttoStr(i))   ; TControl(C).Visible:=(i<=MR_NombreEcheance.Value) ; if not TControl(C).Visible then if (Ta.FindField('MR_MP'+IntToStr(i)).AsString<>'') then THDBValComboBox(C).Value:='' ;
   C:=FindComponent('MR_TAUX'+InttoStr(i)) ; TControl(C).Visible:=(i<=MR_NombreEcheance.Value) ; if not TControl(C).Visible then if (Ta.FindField('MR_TAUX'+IntToStr(i)).AsFloat<>0) then BEGIN Ta.Edit ; Ta.FindField('MR_TAUX'+IntToStr(i)).AsFloat:=0 ; END ;
//   C:=FindComponent('MR_ESC'+InttoStr(i))  ; TControl(C).Visible:=(i<=MR_NombreEcheance.Value) ; if not TControl(C).Visible then if TDBCheckBox(C).Checked then TDBCheckBox(C).Checked:=False ;
   C:=FindComponent('TMR_MP'+InttoStr(i))  ; TControl(C).Visible:=(i<=MR_NombreEcheance.Value) ;
   END ;
end;


Function TFModeRegle.TotalTauxOK : Boolean ;
Var i, j : Byte ;
    Total : Double ;
    C : TDBEdit ;
BEGIN
{ Calcul si le total des taux fait 100% et détruit les anciennes valeurs }
Total:=0 ;
for i:=1 to MR_NombreEcheance.Value do
    BEGIN
    C:=TDBEdit(FindComponent('MR_TAUX'+InttoStr(i))) ;
    if C.Field.AsFloat<=0 then
       BEGIN
       PageControl.ActivePage:=PEche ; C.SetFocus ; HM2.Execute(5,'','') ;
       Result:=False ; Exit ;
       END ;
    Total:=Total+C.Field.AsFloat ;
    END ;
if Total=100 then
   BEGIN
   for j:=MR_NombreEcheance.Value+1 to 12 do
       TDBEdit(FindComponent('MR_TAUX'+InttoStr(j))).Field.AsFloat:=0 ;
   TotalTauxOK:=True ;
   END else
   BEGIN
   PageControl.ActivePage:=PEche ; HM2.Execute(2,'','') ; TotalTauxOK:=False ;
   END ;
END ;

Function TFModeRegle.EnregOK : boolean ;
BEGIN
PageControl.ActivePage:=PCar ;
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if ((Result) and (ta.state in [dsEdit,dsInsert])) then
   BEGIN
   Result:=FALSE ;
   if TaMR_APARTIRDE.AsString='' then BEGIN PageControl.ActivePage:=PCar ; MR_APARTIRDE.SetFocus ; HM2.Execute(4,'','') ; Exit ; END ;
   if not ModePaieOk then BEGIN HM2.Execute(3,'','') ; exit ; END ;
   if not TotalTauxOK then Exit ;
   if ((MR_REMPLACEMIN.Value='') and (MR_REMPLACEMIN.Values.Count>0)) then MR_REMPLACEMIN.Value:=MR_REMPLACEMIN.Values[0] ;
   if TaMR_MONTANTMIN.AsFloat<0 then BEGIN PageControl.ActivePage:=PCar ; MR_MONTANTMIN.SetFocus ; HM2.Execute(1,'','') ; Exit ; END ;
   END ;
result:=TRUE  ; Modifier:=False ;
END ;

Procedure TFModeRegle.NewEnreg ;
var i : byte ;
BEGIN
PageControl.ActivePage:=PCar ;
Inherited ;
TaMR_NOMBREECHEANCE.AsInteger:=1 ;
TaMR_TAUX1.AsFloat:=0 ;
For i:=1 to 12 do
   BEGIN
   TDBEdit(FindComponent('MR_TAUX'+InttoStr(i))).Field.AsFloat:=0 ;
   TDBCheckBox(FindComponent('MR_ESC'+InttoStr(i))).Field.AsString:='-' ;
   END ;
if Guide then MR_MODEGUIDE.Text:='X' else MR_MODEGUIDE.Text:='-';
END ;

Function TFModeRegle.ModePaieOk : Boolean ;
var i : byte ;
    C : THDBValComboBox ;
BEGIN
Result:=True ;
For i:=1 to MR_NOMBREECHEANCE.Value do
   BEGIN
   C:=THDBValComboBox(FindComponent('MR_MP'+InttoStr(i))) ;
   if C.Field.AsString='' then
      BEGIN PageControl.ActivePage:=PEche ; C.SetFocus ; Result:=False ; Exit ; END ;
   END ;
END ;


procedure TFModeRegle.MR_NOMBREECHEANCEChange(Sender: TObject);
begin
  inherited;
SelectEche ;
end;

procedure TFModeRegle.FormShow(Sender: TObject);
begin
PageControl.ActivePage:=PageControl.Pages[0] ;
if Not Guide then Ta.Filtered:=TRUE ;
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
if FLeQuel<>'' then BValider.Enabled:=Guide;
SelectEche ;
end;

Function TFModeRegle.PresenceModeRegle : Boolean ;
Var Trouv : boolean ;
BEGIN
Trouv:=False ; Result:=False ;
Trouv:=Presence('TIERS','T_MODEREGLE',TaMR_MODEREGLE.AsString) ;
if Not Trouv then Trouv:=Presence('GENERAUX','G_MODEREGLE',TaMR_MODEREGLE.AsString) ;
if Not Trouv then Trouv:=Presence('ECRGUI','EG_MODEREGLE',TaMR_MODEREGLE.AsString) ;
if Trouv then BEGIN HM2.Execute(6,'','') ; Result:=True ; END ;
END ;

procedure TFModeRegle.BDeleteClick(Sender: TObject);
begin
if PresenceModeRegle then Exit ;
  inherited;
end;

end.
