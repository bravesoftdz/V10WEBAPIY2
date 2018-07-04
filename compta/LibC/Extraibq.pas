unit Extraibq;

interface


uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, hmsgbox, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBCtrls, StdCtrls, Buttons, ExtCtrls,
  Grids, DBGrids, HDB, Hctrls, Mask, Hent1, Hcompte, Ent1, HSysMenu, Hqry,
  HRegCpte, HTB97, HPanel, ADODB;

Procedure ExtraitBanquaire ;

type
  TFExtraibq = class(TFFicheListe)
    TEE_GENERAL     : THLabel;
    EE_REFPOINTAGE  : TDBEdit;
    TEE_REFPOINTAGE : THLabel;
    EE_DATEPOINTAGE : TDBEdit;
    TEE_DATEPOINTAGE: THLabel;
    GbAncSol           : TGroupBox;
    EE_DATEOldSOLDE : TDBEdit;
    TEE_DATEOldSOLDE: THLabel;
    EE_OldSOLDECRE  : TDBEdit;
    TEE_OldSOLDECRE : THLabel;
    EE_OldSOLDEDEB  : TDBEdit;
    TEE_OldSOLDEDEB : THLabel;
    GbNouSol           : TGroupBox;
    EE_NEWSOLDECRE : TDBEdit;
    TEE_NEWSOLDECRE: THLabel;
    EE_NEWSOLDEDEB : TDBEdit;
    TEE_NEWSOLDEDEB: THLabel;
    TaEE_GENERAL         : TStringField;
    TaEE_REFPOINTAGE     : TStringField;
    TaEE_DATEOldSOLDE : TDateTimeField;
    TaEE_OldSOLDECRE  : TFloatField;
    TaEE_DATEPOINTAGE    : TDateTimeField;
    TaEE_NEWSOLDECRE : TFloatField;
    TaEE_SOCIETE         : TStringField;
    TaEE_OldSOLDEDEB  : TFloatField;
    TaEE_NEWSOLDEDEB : TFloatField;
    EE_GENERAL: THDBCpteEdit;
    TaEE_OLDSOLDEDEBEURO: TFloatField;
    TaEE_OLDSOLDECREEURO: TFloatField;
    TaEE_NEWSOLDEDEBEURO: TFloatField;
    TaEE_NEWSOLDECREEURO: TFloatField;
    GroupBox1: TGroupBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    EE_NEWSOLDECREEURO: TDBEdit;
    EE_NEWSOLDEDEBEURO: TDBEdit;
    iEuroEuro: TImage;
    TLibDevise: TLabel;
    TaEE_NUMERO: TIntegerField;
    procedure EE_GENERALExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
    Function  EnregOK : boolean ; Override ;
  private
    DateMemo : TDateTime ;
    RefMemo  : String ;
    Function  CodeExiste : Boolean ;
    Procedure DeviseDuCompte(St : String) ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Procedure ExtraitBanquaire ;
var FExtraibq : TFExtraibq ;
BEGIN
FExtraibq:=TFExtraibq.Create(Application) ;
 Try
  FExtraibq.InitFL('EE','PRT_EEXBQ','','',taModif,True,FExtraibq.TaEE_GENERAL,
                   FExtraibq.TaEE_REFPOINTAGE,Nil,['']) ;
  FExtraibq.ShowModal ;
 Finally
  FExtraibq.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFExtraibq.NewEnreg ;
BEGIN
 inherited;
EE_GENERAL.Enabled:=True ; EE_REFPOINTAGE.Enabled:=True ; EE_DATEPOINTAGE.Enabled:=True ;
TLibDevise.Caption:='' ; TaEE_NUMERO.AsInteger:=1 ;
END ;

Procedure TFExtraibq.DeviseDuCompte(St : String) ;
Var Q : TQuery ;
    St1 : String ;
BEGIN
If Trim(St)='' Then Exit ;
Q:=OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL="'+St
          +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ',TRUE) ; // 19/10/2006 YMO Multisociétés
If Not Q.Eof Then
  BEGIN
  St1:=RechDom('TTDEVISETOUTES',Q.Fields[0].AsString,False) ;
  TLibDevise.Caption:='('+St1+')' ;
  END Else
  BEGIN
  TLibDevise.Caption:='' ; HM.Execute(8,'','') ;
  END ;
Ferme(Q) ;
END ;

Procedure TFExtraibq.ChargeEnreg ;
BEGIN
 inherited;
EE_GENERAL.Enabled:=False ; EE_REFPOINTAGE.Enabled:=TRUE ; EE_DATEPOINTAGE.Enabled:=TRUE ;
DateMemo:=TaEE_DATEPOINTAGE.AsDateTime ;
RefMemo:=TaEE_REFPOINTAGE.AsString ;
DeviseDuCompte(TaEE_GENERAL.AsString) ;
END ;

Function TFExtraibq.EnregOK : boolean ;
Var St,St1,St2 : String ;
    ModRef,ModDate : Boolean ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if ((Result) and (Ta.state in [dsEdit,dsInsert])) then
   BEGIN
   Result:=FALSE ;
   if CodeExiste then
      BEGIN
      if Ta.state=dsInsert then EE_GENERAL.Setfocus
                           else EE_DATEPOINTAGE.SetFocus ;
      HM.Execute(4,'','') ; Exit ;
      END ;
   if Ta.State=dsInsert then
      if EE_GENERAL.ExisteH=0 then
         BEGIN HM.Execute(6,'','') ; EE_GENERAL.Setfocus ; Exit ; END ;
   END ;
If (Ta.state in [dsEdit]) then
  BEGIN
  ModRef:=TaEE_REFPOINTAGE.AsString<>RefMemo ;
  ModDate:=TaEE_DATEPOINTAGE.AsDateTime<>DateMemo ;
  If ModRef Or ModDate Then
    BEGIN
    St1:='' ; St2:='' ;
    If ModRef  Then St1:='E_REFPOINTAGE="'+TaEE_REFPOINTAGE.AsString+'",' ;
    If ModDate Then St2:='E_DATEPOINTAGE="'+USDATETIME(TaEE_DATEPOINTAGE.AsDateTime)+'",' ;
    St:='UPDATE ECRITURE SET '+St1+St2 ;
    Delete(St,Length(St),1) ;
    St:=St+' WHERE E_REFPOINTAGE="'+RefMemo+'" AND E_DATEPOINTAGE="'+USDATETIME(DateMemo)+'" ' ;
    ExecuteSQL(St) ;
    END ;
  END ;
Result:=TRUE  ; Modifier:=False ;
END ;

Function TFExtraibq.CodeExiste : Boolean ;
Var QLoc : TQuery ;
BEGIN
if(Ta.state in [dsEdit])And(DateMemo=TaEE_DATEPOINTAGE.AsDateTime) then
  BEGIN Result:=False ; Exit ; END ;
QLoc:=OpenSql('Select EE_GENERAL,EE_DATEPOINTAGE,EE_REFPOINTAGE From EEXBQ Where '+
              'EE_GENERAL="'+TaEE_GENERAL.AsString+'" And '+
              'EE_DATEPOINTAGE="'+UsDateTime(TaEE_DATEPOINTAGE.AsDateTime)+'" And '+
              'EE_REFPOINTAGE="'+TaEE_REFPOINTAGE.AsString+'"',True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

procedure TFExtraibq.EE_GENERALExit(Sender: TObject);
begin
  inherited;
If GChercheCompte(EE_GENERAL,Nil) Then DeviseDuCompte(EE_GENERAL.Text) ;
end;

procedure TFExtraibq.FormShow(Sender: TObject);
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then BEGIN If Ta.State=dsInsert Then NewEnreg Else BinsertClick(Nil) ; END ;
end;

procedure TFExtraibq.BDeleteClick(Sender: TObject);
Var Q : TQuery ;
    OkOk : Boolean ;
begin
OkOk:=TRUE ;
Q:=OpenSql('Select E_GENERAL,E_DATEPOINTAGE,E_REFPOINTAGE From ECRITURE Where '+
           'E_GENERAL="'+TaEE_GENERAL.AsString+'" And '+
           'E_DATEPOINTAGE="'+UsDateTime(TaEE_DATEPOINTAGE.AsDateTime)+'" And '+
           'E_REFPOINTAGE="'+TaEE_REFPOINTAGE.AsString+'"',True) ;
If Not Q.Eof Then OkOk:=FALSE ;
Ferme(Q) ;
//If Not OkOk Then HM.Execute(7,'','') Else inherited; //gv

//gv
If Not OkOk Then HM.Execute(7,'','') Else
  if ExisteSql('Select Cel_Refpointage from EEXBQLIG Where Cel_Refpointage="'+TaEE_RefPointage.AsString+'" and Cel_DatePointage="'+UsDateTime(Idate1900)+'"' )
     Then
       begin
       if HM.Execute(9,'','')=MrYes then Inherited else exit ;
       end
     else inherited ;
//+rajout lig 9 dams box de msg
//fin gv
end;

end.
