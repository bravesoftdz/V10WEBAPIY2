unit PlanRef;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DB, StdCtrls, DBCtrls, Hctrls, Mask, Spin, hmsgbox,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, Hent1, Ent1, Menus, ComCtrls,
  HRichEdt, HSysMenu, HRichOLE, MajTable, Hqry, HTB97, HPanel, UiUtil, Paramsoc,
  ADODB ;

Procedure FichePlanRef(NumPlan : integer ; Compte : String ; Mode : TActionFiche);

type
  TFPlanRef = class(TFFicheListe)
    TPR_COMPTE: THLabel;
    PR_COMPTE: TDBEdit;
    tPR_ABREGE: THLabel;
    PR_ABREGE: TDBEdit;
    TPR_LIBELLE: THLabel;
    PR_LIBELLE: TDBEdit;
    TPR_SENS: THLabel;
    PR_SENS: THDBValComboBox;
    TPR_NATUREGENE: THLabel;
    PR_NATUREGENE: THDBValComboBox;
    PR_COLLECTIF: TDBCheckBox;
    PR_POINTABLE: TDBCheckBox;
    PR_LETTRABLE: TDBCheckBox;
    HGBOptionaxes: TGroupBox;
    PR_VENTILABLE1: TDBCheckBox;
    PR_VENTILABLE2: TDBCheckBox;
    PR_VENTILABLE3: TDBCheckBox;
    PR_VENTILABLE4: TDBCheckBox;
    PR_VENTILABLE5: TDBCheckBox;
    HGBOptionImpression: TGroupBox;
    PR_SOLDEPROGRESSIF: TDBCheckBox;
    PR_TOTAUXMENSUELS: TDBCheckBox;
    PR_SAUTPAGE: TDBCheckBox;
    PR_CENTRALISABLE: TDBCheckBox;
    TG_BLOCNOTE: TGroupBox ;
    PR_BLOCNOTE: THDBRichEditOLE;
    TaPR_NUMPLAN: TIntegerField;
    TaPR_COMPTE: TStringField;
    TaPR_LIBELLE: TStringField;
    TaPR_ABREGE: TStringField;
    TaPR_NATUREGENE: TStringField;
    TaPR_CENTRALISABLE: TStringField;
    TaPR_SOLDEPROGRESSIF: TStringField;
    TaPR_SAUTPAGE: TStringField;
    TaPR_TOTAUXMENSUELS: TStringField;
    TaPR_COLLECTIF: TStringField;
    TaPR_BLOCNOTE: TMemoField;
    TaPR_SENS: TStringField;
    TaPR_LETTRABLE: TStringField;
    TaPR_POINTABLE: TStringField;
    TaPR_VENTILABLE1: TStringField;
    TaPR_VENTILABLE2: TStringField;
    TaPR_VENTILABLE3: TStringField;
    TaPR_VENTILABLE4: TStringField;
    TaPR_VENTILABLE5: TStringField;
    TaPR_REPORTDETAIL: TStringField;
    Pnum: TPanel;
    HLabel1: THLabel;
    FNumPlan: TSpinEdit;
    XX_WHERE: TPanel;
    PR_PREDEFINI: THDBValComboBox;
    TaPR_PREDEFINI: TStringField;
    procedure FormShow(Sender: TObject);
    procedure FNumPlanChange(Sender: TObject);
    procedure PR_NATUREGENEChange(Sender: TObject);
    procedure PR_COMPTEExit(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
  private
    Nature,Compte : String ;
    Ventipoint : Boolean ;
    NumPlan : Integer ;
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
    Procedure PointageAutorise ;
    Procedure LettrageAutorise ;
    Procedure CollectifAutorise ;
    Procedure VentilationAutorise ;
    Procedure ChangeCheckBox(Sender : TDBCheckBox) ;
    Procedure ForceCollectif ;
    Function  CodeExiste : Boolean ;

  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  PrintDBG, UtilPgi ;


Procedure FichePlanRef(NumPlan : integer ; Compte : String ; Mode : TActionFiche) ;
var FPlanRef: TFPlanRef;
{$IFDEF SPEC302}
    QLoc : TQuery ;
{$ENDIF}
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FPlanRef:=TFPlanRef.Create(Application) ;
{$IFDEF SPEC302}
QLoc:=OpenSql('Select SO_NUMPLANREF From SOCIETE',True) ;
if Not QLoc.Eof then if QLoc.Fields[0].AsInteger<>0 then NumPlan:=QLoc.Fields[0].AsInteger ;
Ferme(QLoc) ;
{$ELSE}
NumPlan:=GetParamsoc('SO_NUMPLANREF') ;
{$ENDIF}
FPlanRef.Compte:=Compte ;
FPlanRef.NumPlan:=NumPlan ;
FPlanRef.InitFL('PR','PRT_PLANREF',Compte,'',Mode,TRUE,FPlanRef.TaPR_COMPTE,
             FPlanRef.TaPR_LIBELLE,Nil,['ttPlanReference']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FPlanRef.ShowModal ;
    Finally
     FPlanRef.Free ;
    End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FPlanRef,PP) ;
   FPlanRef.Show ;
   END ;
END ;

Procedure TFPlanRef.ChargeEnreg ;
BEGIN
{ FQ 20958 BVE 13.07.07 }
if PR_PREDEFINI.Value = 'CEG' then
   FTypeAction := taConsult
else                        
   FTypeAction := taModif;

Inherited ;
                     
if PR_PREDEFINI.Value = 'CEG' then
begin
   FNumPlan.Enabled := true;
   Binsert.Enabled := true;
   BDelete.Enabled := false;
end
else
begin
   FTypeAction := taModif;
   FNumPlan.Enabled := true;
   PR_ABREGE.Enabled := true;
   PR_LIBELLE.Enabled := true;
   PR_SENS.Enabled := true;
   PR_NATUREGENE.Enabled := true;
   PR_COLLECTIF.Enabled := true;
   PR_LETTRABLE.Enabled := true;
   PR_BLOCNOTE.Enabled := true;   
   PR_BLOCNOTE.ReadOnly := false;
   PR_SOLDEPROGRESSIF.Enabled := true;
   PR_TOTAUXMENSUELS.Enabled := true;
   PR_SAUTPAGE.Enabled := true;
   PR_CENTRALISABLE.Enabled := true;   
   Binsert.Enabled := true;
   BDelete.Enabled := true;  
   BValider.Enabled := true;
end;
{ END FQ 20958 }
Caption:=HM.Mess[6]+' '+IntToStr(FNumPlan.Value)+' : '+TaPR_COMPTE.AsString+' '+TaPR_LIBELLE.AsString ;
if FTypeAction<>taConsult then PR_NATUREGENEChange(Nil) ;
UpdateCaption(Self) ;
END ;

Procedure TFPlanRef.NewEnreg ;
BEGIN
Inherited ;
TaPR_NumPlan.AsInteger:=FNumPlan.Value ;
PR_NATUREGENE.Value:='DIV' ;
TaPR_SOLDEPROGRESSIF.AsString:='X' ;
{ FQ 20958 BVE 13.07.07 }
PR_PREDEFINI.Value := 'STD';
{ END FQ 20958 }
TaPR_SENS.AsString:='M' ;
END ;

procedure TFPlanRef.FormShow(Sender: TObject);
begin
ChangeSizeMemo(TaPR_BLOCNOTE) ;
Ventipoint:=True ;
  inherited;
if NumPlan=1 then FNumPlanChange(Nil)
             else FNumPlan.Value:=NumPlan ;
// JLD5Axes
if EstSerie(S3) then
   BEGIN
   PR_VENTILABLE3.Visible:=False ; PR_VENTILABLE4.Visible:=False ; PR_VENTILABLE5.Visible:=False ;
   PR_VENTILABLE2.Visible:=False ;
   END ;
end;

procedure TFPlanRef.FNumPlanChange(Sender: TObject);
begin
  inherited;
if Compte<>'' then Ta.SetRange([FNumPlan.Value,Compte],[FNumPlan.Value,Compte])
              else Ta.SetRange([FNumPlan.Value],[FNumPlan.Value]) ;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
end;

Function TFPlanRef.EnregOK : boolean ;
BEGIN
Result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ; Result:=False ;
if Ta.State in [dsInsert] then
   BEGIN
   if CodeExiste then BEGIN HM.Execute(4,'','') ; Exit ; END ;
   END ;
Modifier:=False ; Result:=True ; ForceCollectif ;
END ;

procedure TFPlanRef.PR_NATUREGENEChange(Sender: TObject);
begin
  inherited;
Nature:=PR_NATUREGENE.Value ;
PointageAutorise ; LettrageAutorise ; CollectifAutorise ; VentilationAutorise ;
end;

Procedure TFPlanRef.PointageAutorise ;
BEGIN
if(Nature='DIV')Or(Nature='EXT')Or
  (Nature='BQE')Or(Nature='CAI')then PR_POINTABLE.Enabled:=True
                                else PR_POINTABLE.Enabled:=False ;
ChangeCheckBox(PR_POINTABLE) ;
END ;

Procedure TFPlanRef.LettrageAutorise ;
BEGIN
if(Nature='TIC')Or(Nature='TID')then
  BEGIN
  PR_LETTRABLE.Enabled:=True ;
  if Ta.State in [dsEdit,dsInsert] then TaPR_LETTRABLE.AsString:='X' ;
  END else PR_LETTRABLE.Enabled:=False ;
ChangeCheckBox(PR_LETTRABLE) ;
END ;

Procedure TFPlanRef.CollectifAutorise ;
BEGIN
if(Nature='COC')Or(Nature='COD')Or
  (Nature='COF')Or(Nature='COS')then PR_COLLECTIF.Enabled:=True
                                else PR_COLLECTIF.Enabled:=False ;
ChangeCheckBox(PR_COLLECTIF) ;
END ;

Procedure TFPlanRef.VentilationAutorise ;
Var Avec : Boolean ;
    i : Byte ;
    C : TComponent ;
BEGIN
if(Nature='CHA')Or(Nature='DIV')Or(Nature='EXT')Or
  (Nature='IMO')Or(Nature='PRO')Or(((Nature='BQE')Or
  (Nature='CAI'))And Ventipoint) then Avec:=True
                                 else Avec:=False ;
for i:=1 to 5 do
    BEGIN
    C:=FindComponent('PR_VENTILABLE'+IntToStr(i)+'') ; TDBCheckBox(C).Enabled:=Avec ;
    if(Not Avec)And(Ta.State in [dsInsert,dsEdit]) then TDBCheckBox(C).Field.AsString:='-' ;
    END ;
END ;

Procedure TFPlanRef.ChangeCheckBox(Sender : TDBCheckBox) ;
Var F : TField ;
BEGIN
F:=Ta.FindField(Sender.Name) ;
if Ta.State in [dsEdit,dsInsert] then
   if Not TDBCheckBox(Sender).Enabled then F.AsString:='-' ;
END ;

Procedure TFPlanRef.ForceCollectif ;
BEGIN
if(Nature='COC')Or(Nature='COD')Or
  (Nature='COF')Or(Nature='COS')then TaPR_COLLECTIF.AsString:='X'
                                else TaPR_COLLECTIF.AsString:='-' ;
END ;

procedure TFPlanRef.PR_COMPTEExit(Sender: TObject);
begin
  inherited;
if Length(PR_COMPTE.Text)>VH^.Cpta[fbGene].Lg then
   TaPR_COMPTE.AsString:=Copy(TaPR_COMPTE.AsString,1,VH^.Cpta[fbGene].Lg) ;
end;

procedure TFPlanRef.BImprimerClick(Sender: TObject);
var MyBookmark: TBookmark;
begin
//  inherited;
MyBookmark :=Ta.GetBookmark ;
XX_WHERE.Hint:=' Where PR_NUMPLAN='+IntToStr(FNumPlan.Value)+'' ;
PrintDBGrid (Nil,XX_WHERE,Copy(Caption,1,Pos(':',Caption)-1),'PRT_PLANREF') ;
Ta.GotoBookmark(MyBookmark) ; Ta.FreeBookmark(MyBookmark);
end;

Function TFPlanRef.CodeExiste : Boolean ;
BEGIN
Result:=PresenceComplexe('PLANREF',['PR_COMPTE','PR_NUMPLAN'],['=','='],[TaPR_COMPTE.AsString,IntToStr(TaPR_NUMPLAN.AsInteger)],['S','I']) ;
END ;

end.
