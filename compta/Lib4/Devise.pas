unit Devise;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFNDEF EAGLCLIENT}
  FichList,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, StdCtrls, Hcompte, HDB, DBCtrls, Spin, Mask,
  Hctrls, hmsgbox, Buttons, ExtCtrls, Grids, DBGrids, HEnt1, Ent1, HSysMenu,
  Hqry, HRegCpte, ComCtrls, HTB97, HPanel, UiUtil, SaisUtil, Paramsoc,
  ADODB
{$ELSE}
	UtilEagl,eFichList,
{$ENDIF}
;
Procedure FicheDevise(Quel : String ; Comment : TActionFiche ; FromAss : boolean ) ;

type
  TFDevise = class(TFFicheListe)
    TD_DEVISE: THLabel;
    D_DEVISE: TDBEdit;
    TD_SYMBOLE: THLabel;
    D_SYMBOLE: TDBEdit;
    TD_LIBELLE: THLabel;
    D_LIBELLE: TDBEdit;
    TD_QUOTITE: THLabel;
    D_QUOTITE: TDBEdit;
    TD_DECIMALE: THLabel;
    D_DECIMALE: THDBSpinEdit;
    D_FERME: TDBCheckBox;
    TaD_DEVISE: TStringField;
    TaD_LIBELLE: TStringField;
    TaD_SYMBOLE: TStringField;
    TaD_DECIMALE: TIntegerField;
    TaD_QUOTITE: TFloatField;
    TaD_CPTLETTRDEBIT: TStringField;
    TaD_CPTLETTRCREDIT: TStringField;
    TaD_CPTPROVDEBIT: TStringField;
    TaD_CPTPROVCREDIT: TStringField;
    TaD_SOCIETE: TStringField;
    TaD_MAXDEBIT: TFloatField;
    TaD_MAXCREDIT: TFloatField;
    BChancel: TToolbarButton97;
    TaD_FERME: TStringField;
    Pages: TPageControl;
    TSREGUL: TTabSheet;
    TSEURO: TTabSheet;
    GB1: TGroupBox;
    TTD_CPTLETTRCREDIT: THLabel;
    TD_CPTLETTRDEBIT: THLabel;
    TD_CPTLETTRCREDIT: THLabel;
    TTD_CPTLETTRDEBIT: THLabel;
    D_CPTLETTRDEBIT: THDBCpteEdit;
    D_CPTLETTRCREDIT: THDBCpteEdit;
    GB2: TGroupBox;
    TTD_CPTPRODEBIT: THLabel;
    TTD_CPTPROCREDIT: THLabel;
    TD_CPTPRODEBIT: THLabel;
    TD_CPTPROCREDIT: THLabel;
    D_CPTPRODEBIT: THDBCpteEdit;
    D_CPTPROCREDIT: THDBCpteEdit;
    GroupBox3: TGroupBox;
    TD_MAXDEBIT: THLabel;
    TD_MAXCREDIT: THLabel;
    TaD_MONNAIEIN: TStringField;
    TaD_FONGIBLE: TStringField;
    D_MONNAIEIN: TDBCheckBox;
    D_FONGIBLE: TDBCheckBox;
    FPARITEEURO: THLabel;
    D_PARITEEURO: TDBEdit;
    Label1: TLabel;
    FDateDebutEuro: TLabel;
    DateDebutEuro: TLabel;
    Label3: TLabel;
    BChancelOut: TToolbarButton97;
    HConseil: THLabel;
    DEVPIV: TCheckBox;
    TaD_PARITEEURO: TFloatField;
    TaD_CODEISO: TStringField;
    TD_ISO: THLabel;
    D_CODEISO: TDBEdit;
    TSFIXING: TTabSheet;
    TaD_PARITEEUROFIXING: TFloatField;
    D_PARITEEUROFIXING: TDBEdit;
    TD_PARITEEUROFIXING: THLabel;

    procedure BChancelClick(Sender: TObject);
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure BDeleteClick(Sender: TObject);
    procedure TaAfterPost(DataSet: TDataSet);
    procedure TaAfterDelete(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject); Override;
    procedure TaBeforeDelete(DataSet: TDataSet);
    procedure D_MONNAIEINClick(Sender: TObject);
    procedure BChancelOutClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure D_PARITEEUROExit(Sender: TObject);
    procedure D_FONGIBLEClick(Sender: TObject);
    procedure D_PARITEEUROChange(Sender: TObject);
    procedure ChargeEnreg ; Override ;
    Function  EnregOK : boolean ; override ;
    Procedure NewEnreg ; override ;
  private    { Déclarations privées }
    OldDec,OldQuot : Integer ;
    OldParite      : double ;
    Favertir,AvecMvt : Boolean ;
    Mode : TActionFiche ;
    Acreer,FromAss : Boolean ;
    MemoDev  : String3 ;
    Procedure EstMouvementer(St : String) ;
    Function  MouvementDepuisEuro (St : String) : boolean ;
    function  VerifPariteEuro : boolean ;
    procedure DisaChange ;
    procedure PosDevPiv ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses Chancel,UtilPGI ;

Procedure FicheDevise(Quel : String ; Comment : TActionFiche ; FromAss : boolean ) ;
var FDevise: TFDevise;
    PP : THPanel ;
begin
if _Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;
FDevise:= TFDevise.Create(Application) ;
FDevise.Mode:=Comment ; FDevise.FromAss:=FromAss ;
FDevise.InitFL('D','PRT_DEVISE',quel,'',Comment,TRUE,FDevise.TaD_DEVISE,FDevise.TaD_LIBELLE,FDevise.TaD_DEVISE,['ttDevise','ttDeviseEtat']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FDevise.ShowModal ;
    finally
     FDevise.Free ;
     _Bloqueur('nrBatch',False) ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FDevise,PP) ;
   FDevise.Show ;
   END ;
end ;

Procedure TFDevise.ChargeEnreg  ;
BEGIN
Inherited ;
OldDec:=TaD_DECIMALE.AsInteger ; OldQuot:=TaD_QUOTITE.AsInteger ; OldParite:=TaD_PARITEEURO.AsFloat ;
D_CPTLETTRDEBIT.ExisteH ; D_CPTLETTRCREDIT.ExisteH ;
D_CPTPRODEBIT.ExisteH ; D_CPTPROCREDIT.ExisteH ;
if TaD_DEVISE.AsString=V_PGI.DevisePivot then
   BEGIN
   GB1.Enabled:=False ; D_CPTLETTRDEBIT.Color:=clBtnFace ; D_CPTLETTRCREDIT.Color:=clBtnFace ;
   GB2.Enabled:=False ; D_CPTPRODEBIT.Color:=clBtnFace ; D_CPTPROCREDIT.Color:=clBtnFace ;
   END else
   BEGIN
   GB1.Enabled:=True ; D_CPTLETTRDEBIT.Color:=clWindow ; D_CPTLETTRCREDIT.Color:=clWindow ;
   GB2.Enabled:=True ; D_CPTPRODEBIT.Color:=clWindow ; D_CPTPROCREDIT.Color:=clWindow ;
   BChancel.Enabled:=True ;
   END ;
D_MONNAIEINClick(nil) ;
DisaChange ;
if ((TAD_FONGIBLE.AsString='X') and (VH^.TenueEuro)) then
   BEGIN
   D_PARITEEURO.Enabled:=False ;
   D_FONGIBLE.Enabled:=False ; 
   END ;
HConseil.Visible:=D_PARITEEURO.Enabled ;
HConseil.Caption:=HM.Mess[16]+' '+D_SYMBOLE.Text+' '+D_DEVISE.Text+')' ;
BChancelOut.Visible:=Not (D_MONNAIEIN.Checked) ;
TSEURO.TabVisible:=not (VH^.TenueEuro and (D_DEVISE.Text=V_PGI.DevisePivot)) ;
if not (TSEURO.TabVisible) and (Pages.ActivePage<>TSREGUL) then Pages.ActivePage:=TSREGUL ;
PosDevPiv ;
END ;

Function TFDevise.EnregOK : boolean ;
//Var QLoc : TQuery ;
BEGIN
Result:=Inherited EnregOK ; if Not Result then Exit ;
if Ta.State in [dsInsert] then Acreer:=True ;
if Ta.State in [dsEdit,dsInsert] then
   if not VerifPariteEuro then
      BEGIN
      HM.Execute(14,'','') ; D_PARITEEURO.SetFocus ; Modifier:=True ;
      if Not V_PGI.SAV then BEGIN Result:=False ; Exit ; END ;
      END ;
if Ta.State in [dsEdit] then
   BEGIN
   if (OldDec<>TaD_DECIMALE.AsInteger) or (OldQuot<>TaD_QUOTITE.AsInteger) then
      BEGIN
      EstMouvementer(TaD_DEVISE.AsString) ;
      if AvecMvt then
         BEGIN
         Modifier:=True ; Result:=False ;
         if OldDec<>TaD_DECIMALE.AsInteger then BEGIN HM.Execute(11,'','') ; TaD_DECIMALE.AsInteger:=OldDec ; END ;
         if OldQuot<>TaD_QUOTITE.AsInteger then BEGIN HM.Execute(12,'','') ; TaD_QUOTITE.AsInteger:=OldQuot ; END ;
         END ;
      END ;
   if ((OldParite<>TaD_PARITEEURO.AsFloat) and (Not V_PGI.SAV)) then
      BEGIN
      if MouvementDepuisEuro(TaD_DEVISE.AsString) then
         BEGIN
         Modifier:=True ; Result:=False ;
         HM.Execute(15,'','') ; TaD_PARITEEURO.AsFloat:=OldParite ;
         END ;
      END ;
   END ;
{$IFDEF SPEC302}
if ((DEVPIV.Visible) and (DEVPIV.Checked) and (V_PGI.DevisePivot='') and (FromAss)) then
   BEGIN
   ExecuteSQL('UPDATE SOCIETE SET SO_DEVISEPRINC="'+TaD_DEVISE.AsString+'"') ;
   V_PGI.DevisePivot:=TaD_DEVISE.AsString ;
   END ;
if TaD_Devise.AsString=V_PGI.DevisePivot then
  BEGIN
  V_PGI.OkDecV:=StrToInt(D_DECIMALE.Text) ;
  ExecuteSql('Update SOCIETE SET SO_DECVALEUR='+IntToStr(V_PGI.OkDecV)+', '+
             'SO_TAUXEURO='+StrfPoint(TaD_PARITEEURO.AsFloat)+' '+
             'Where SO_SOCIETE="'+V_PGI.CodeSociete+'" AND SO_DEVISEPRINC="'+TaD_Devise.AsString+'"') ;
  END ;
{$ELSE}
if ((DEVPIV.Visible) and (DEVPIV.Checked) and (V_PGI.DevisePivot='') and (FromAss)) then
   BEGIN
   SetParamSoc('SO_DEVISEPRINC',TaD_DEVISE.AsString) ;
   V_PGI.DevisePivot:=TaD_DEVISE.AsString ;
   END ;
if TaD_Devise.AsString=V_PGI.DevisePivot then
  BEGIN
  V_PGI.OkDecV:=StrToInt(D_DECIMALE.Text) ;
  SetParamSoc('SO_DECVALEUR',V_PGI.OkDecV) ;
  SetParamSoc('SO_TAUXEURO',TaD_PARITEEURO.AsFloat) ;
  END ;
{$ENDIF}
END ;

Procedure TFDevise.NewEnreg ;
BEGIN
InHerited ;
TaD_DECIMALE.AsInteger:=2 ;
TaD_QUOTITE.AsInteger:=1 ;
TaD_FONGIBLE.AsString:='-' ;
TaD_MONNAIEIN.AsString:='-' ;
TaD_PARITEEURO.AsFloat:=1 ;
TaD_PARITEEUROFIXING.AsFloat:=1.00 ;
END ;

procedure TFDevise.BChancelClick(Sender: TObject);
begin
  inherited;
if FAvertir then AvertirTable('ttDevise') ;
FicheChancel(D_DEVISE.Text,False,0,Mode,False) ;
end;

procedure TFDevise.STaDataChange(Sender: TObject; Field: TField);
begin
  inherited;
BChancel.Enabled:=(Not (Ta.State in [dsEdit,dsInsert])) ;
BChancelOut.Enabled:=(Not (Ta.State in [dsEdit,dsInsert])) ;
if (TaD_DEVISE.AsString=V_PGI.DevisePivot) then BChancel.Enabled:=False ;
if (VH^.TenueEuro and (TaD_DEVISE.Text=V_PGI.DevisePivot)) then BChancelOut.Enabled:=False ;
end;

Function TFDevise.MouvementDepuisEuro (St : String) : boolean ;
Var QQ : TQuery ;
BEGIN
Result:=False ;
if Not EstMonnaieIn(St) then Exit ;
QQ:=OpenSQL('Select E_DEVISE from ECRITURE Where E_DATECOMPTABLE>="'+UsDateTime(V_PGI.DateDebutEuro)+'" AND E_EXERCICE="'+QuelExoDT(V_PGI.DateDebutEuro)+'" AND E_DEVISE="'+St+'"',True,-1,'',true);
if Not QQ.EOF then Result:=True ;
Ferme(QQ) ;
END ;

Procedure TFDevise.EstMouvementer(St : String) ;
BEGIN
AvecMvt:=ExisteSql('Select D_DEVISE from DEVISE Where D_DEVISE="'+St+'" '+
                   'AND (Exists(Select E_DEVISE from ECRITURE Where E_DEVISE="'+St+'"))') ;
if Not AvecMvt then
   BEGIN
   AvecMvt:=ExisteSql('Select D_DEVISE from DEVISE Where D_DEVISE="'+St+'" '+
                      'AND (Exists(Select GP_DEVISE from PIECE Where GP_DEVISE="'+St+'"))') ;
   if Not AvecMvt then AvecMvt:=ExisteSql('Select T_DEVISE from TIERS Where T_DEVISE="'+St+'"') ;
   END ;
END ;

procedure TFDevise.BDeleteClick(Sender: TObject);
begin
if ((VH^.TenueEuro) and (TAD_DEVISE.AsString=V_PGI.DeviseFongible)) then  BEGIN HM.Execute(17,'',''); Exit ; END ;
SourisSablier ; EstMouvementer(TaD_DEVISE.AsString) ;
if AvecMvt then BEGIN HM.Execute(13,'','') ; Exit ; END ;
  inherited;
SourisNormale ;
end;

procedure TFDevise.TaAfterPost(DataSet: TDataSet);
begin
  inherited;
Favertir:=True ;
end;

procedure TFDevise.TaAfterDelete(DataSet: TDataSet);
begin
  inherited;
Favertir:=True ;
end;

procedure TFDevise.FormShow(Sender: TObject);
begin
Favertir:=False ; Acreer:=False ;
  inherited;
//if(Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then BinsertClick(Nil) ;
DateDebutEuro.Caption:=FormatDateTime('dd mmmm yyyy',V_PGI.DateDebutEuro) ;
PosDevPiv ;
TSFIXING.TabVisible := (GetParamSoc('SO_PREFSYSTTARIF') and ( not (ctxCompta in V_PGI.PGIContexte)));
end;

procedure TFDevise.BValiderClick(Sender: TObject);
Var QLoc : TQuery ;
begin
  inherited;
if ACreer then
   BEGIN
   if (D_DEVISE.Text<>V_PGI.DevisePivot) and (not D_FONGIBLE.Checked) then
      BEGIN
      AvertirTable('ttDevise') ;
      if Not D_MONNAIEIN.Checked then
         BEGIN
         FicheChancel(D_DEVISE.Text,False,0,taCreat,True) ;
         QLoc:=OpenSql('Select H_TAUXREEL From CHANCELL Where H_DEVISE="'+D_DEVISE.Text+'"',True,-1,'',true);
         if QLoc.Eof then HM.Execute(9,'','') else
            if(QLoc.Fields[0].AsFloat=0)then HM.Execute(8,'','') ;
         Ferme(QLoc) ;
         END ;
      END ;
   MemoDev:=D_DEVISE.Text ; ACreer:=False ;
   END ;
end;

procedure TFDevise.TaBeforeDelete(DataSet: TDataSet);
begin
  inherited;
ExecuteSql('Delete From CHANCELL Where H_DEVISE="'+TaD_DEVISE.AsString+'"') ;
end;

function TFDevise.VerifPariteEuro : boolean ;
Var StParite : String ;
    Parite   : Integer ;
BEGIN
Result:=True ;
StParite:=Trim(D_PARITEEURO.Text) ;
if Pos(V_PGI.SepDecimal,StParite)<>0 then Delete(StParite,Pos(V_PGI.SepDecimal,StParite),1) ;
Parite:=StrToInt(StParite) ;
if Parite<=0 then BEGIN Result:=False ; Exit ; END ;
StParite:=IntToStr(Parite) ;
if Length(StParite)>6 then Result:=False ;
END ;

function ExistMonnaieFongible(St : String) : boolean ;
BEGIN
Result:=ExisteSql('Select D_DEVISE from DEVISE Where D_FONGIBLE="X" AND D_DEVISE<>"'+St+'"') ;
END ;

procedure TFDevise.D_MONNAIEINClick(Sender: TObject);
begin
  inherited;
D_FONGIBLE.Enabled:=(D_MONNAIEIN.Checked) and (not ExistMonnaieFongible(taD_DEVISE.AsString)) ;
if not (D_MONNAIEIN.Checked) then
  BEGIN
  if D_FONGIBLE.Checked then D_FONGIBLE.Checked:=False ;
  DateDebutEuro.Caption:=HM2.Mess[2] ;
  END else
  BEGIN
  if not D_FONGIBLE.Enabled and D_FONGIBLE.Checked then D_FONGIBLE.Checked:=False ;
  DateDebutEuro.Caption:=DateToStr(V_PGI.DateDebutEuro) ;
  END ;
D_PARITEEURO.Enabled:=((D_MONNAIEIN.Checked)) and (V_PGI.DateDebutEuro>=EncodeDate(1999,01,01))
                       and (V_PGI.DateDebutEuro<EncodeDate(1999,12,31)) ;
if V_PGI.SAV then D_PARITEEURO.Enabled:=True ;
FPARITEEURO.Enabled:=(D_MONNAIEIN.Checked) ;
DateDebutEuro.Enabled:=(D_MONNAIEIN.Checked) ;
FDateDebutEuro.Enabled:=(D_MONNAIEIN.Checked) ;
end;

procedure TFDevise.BChancelOutClick(Sender: TObject);
begin
  inherited;
if FAvertir then BEGIN AvertirTable('ttDevise') ; AvertirTable('ttDeviseEtat') ; END ;
FicheChancel(D_DEVISE.Text,False,0,Mode,True) ;
end;

procedure TFDevise.DisaChange ;
BEGIN
if ((D_MONNAIEIN.Checked) or (TAD_DEVISE.AsString=V_PGI.DevisePivot)) then
   BEGIN
   GB1.Enabled:=False ; GB2.Enabled:=False ;
   D_CPTLETTRDEBIT.Enabled:=False ; D_CPTLETTRCREDIT.Enabled:=False ;
   D_CPTPRODEBIT.Enabled:=False   ; D_CPTPROCREDIT.Enabled:=False ;
   TD_CPTLETTRDEBIT.Enabled:=False ; TD_CPTLETTRCREDIT.Enabled:=False ;
   TD_CPTPRODEBIT.Enabled:=False   ; TD_CPTPROCREDIT.Enabled:=False ;
   D_CPTLETTRDEBIT.Text:='' ; D_CPTLETTRCREDIT.Text:='' ;
   D_CPTPRODEBIT.Text:=''   ; D_CPTPROCREDIT.Text:='' ;
   TTD_CPTLETTRDEBIT.Caption:='' ; TTD_CPTLETTRCREDIT.Caption:='' ;
   TTD_CPTPRODEBIT.Caption:=''   ; TTD_CPTPROCREDIT.Caption:='' ;
   END else
   BEGIN
   GB1.Enabled:=True ; GB2.Enabled:=True ;
   D_CPTLETTRDEBIT.Enabled:=True ; D_CPTLETTRCREDIT.Enabled:=True ;
   D_CPTPRODEBIT.Enabled:=True   ; D_CPTPROCREDIT.Enabled:=True ;
   TD_CPTLETTRDEBIT.Enabled:=True ; TD_CPTLETTRCREDIT.Enabled:=True ;
   TD_CPTPRODEBIT.Enabled:=True   ; TD_CPTPROCREDIT.Enabled:=True ;
   END ;
END ;

procedure TFDevise.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then _Bloqueur('nrBatch',False) ;
if FAvertir then ChargeSocieteHalley ; 
  inherited;
end;

procedure TFDevise.D_PARITEEUROExit(Sender: TObject);
begin
  inherited;
if D_PARITEEURO.Text='' then taD_PARITEEURO.AsFloat:=0 ;
end;

procedure TFDevise.PosDevPiv ;
BEGIN
DEVPIV.Visible:=((FromAss) and (D_FONGIBLE.Checked) and (V_PGI.DevisePivot='')) ;
if Not DEVPIV.Visible then DEVPIV.Checked:=False ;
END ;

procedure TFDevise.D_FONGIBLEClick(Sender: TObject);
begin
  inherited;
PosDevPiv ;
end;

procedure TFDevise.D_PARITEEUROChange(Sender: TObject);
Var St : String ;
begin
  inherited;
St:=D_PARITEEURO.Text ;
if Valeur(St)>999999.9 then
   BEGIN
   HM.Execute(14,'','') ;
   While Valeur(St)>999999.9 do Delete(St,Length(St),1) ;
   D_PARITEEURO.Text:=St ; 
   END ;
end;

end.
