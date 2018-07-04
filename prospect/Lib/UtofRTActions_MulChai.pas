{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CHAINAGE_RESP ()
Mots clefs ... : TOF;CHAINAGE_RESP
*****************************************************************}
Unit UtofRTActions_MulChai ;

Interface

Uses
     Classes,
{$IFDEF EAGLCLIENT}
     Maineagl,eMul,
{$ELSE}
     FE_Main,Mul,
{$ENDIF}
     sysutils,
     HCtrls,
     wTof,
     UTOF,HTB97 ;

Type
  TOF_Actions_MulChai = Class (tWTOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad               ; override;
  Private
    stProduitpgi,stTiers : String;
    procedure BZoom_OnClick(Sender : tObject);
    procedure BInsert_OnClick(Sender : tObject);
    procedure FLISTE_OnDblClick(Sender: TObject);
  end ;

Function RTLanceFiche_Chainage_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

Uses
  UtilPGI
  ;

Function RTLanceFiche_Chainage_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_Actions_MulChai.OnArgument (S : String ) ;
begin
  Inherited ;
  stProduitpgi:= GetArgumentValue (S,'PRODUITPGI') ;
  stTiers:= GetArgumentValue (S,'TIERS') ;
  if Assigned(GetControl('FLISTE')) then
    THGrid(GetControl('FLISTE' )).OnDblClick := FLISTE_OnDblClick;
  if Assigned(GetControl('BZOOM')) then
    TToolBarButton97(GetControl( 'BZOOM' )).OnClick := BZoom_OnClick;
  if Assigned(GetControl('Binsert')) then
    TToolBarButton97(GetControl( 'Binsert' )).OnClick := BInsert_OnClick;
end;

procedure TOF_Actions_MulChai.OnLoad;
begin
  Inherited;
  if copy(ecran.name,1,17) <> 'RTCHAINAGES_TIERS' then
    SetControlText ('XX_WHERE', 'NOT EXISTS (SELECT RCH_NUMERO FROM WINTERVENTION WHERE WIV_NUMCHAINAGE=RCH_NUMERO)');
end;

procedure TOF_Actions_MulChai.FLISTE_OnDblClick(Sender: TObject);
begin
  if copy(ecran.name,1,17) <> 'RTCHAINAGES_TIERS' then
    begin
    if GetString('RCH_NUMERO') <> '' then
      TFMul(ecran).retour:= IntToStr(GetField('RCH_NUMERO'));
    if GetString('RCH_CHAINAGE') <> '' then
      TFMul(ecran).retour:= TFMul(ecran).retour+';'+GetField('RCH_CHAINAGE');
    TFMul(Ecran).Close;
    end
  else
    begin
    if GetString('RCH_NUMERO') = '' then exit;
    AGLLanceFiche('RT','RTCHAINAGES','','','GRC;ORIGINETIERS;RCH_NUMERO='+getString('RCH_NUMERO'));
    end;
end;

procedure TOF_Actions_MulChai.BZoom_OnClick(Sender : tObject);
begin
  Inherited ;
  if GetString('RCH_NUMERO') = '' then exit;
  AGLLanceFiche('RT','RTACTIONSCHAINE','',IntToStr(GetField('RCH_NUMERO'))+';'+stProduitpgi,
       'ACTION=MODIFICATION;MONOFICHE;LISTEMENU;PRODUITPGI='+stProduitpgi);
  RefreshDB;
end ;

procedure TOF_Actions_MulChai.Binsert_OnClick(Sender : tObject);
var NumChainage : String;
begin
  Inherited ;
  NumChainage := AGLLanceFiche('RT','RTACTIONSCHAINE','','','MONOFICHE;ACTION=CREATION;LISTEMENU;PRODUITPGI='+stProduitpgi+';RCH_TIERS='+stTiers);
  if copy(ecran.name,1,17) <> 'RTCHAINAGES_TIERS' then
    begin
    TFMul(ecran).retour:= NumChainage;
    TFMul(Ecran).Close;
    end
  else
    begin
    if NumChainage = '' then exit;
    AGLLanceFiche('RT','RTCHAINAGES','','','GRC;ORIGINETIERS;RCH_NUMERO='+NumChainage);
    end;

end ;

Initialization
  registerclasses ( [ TOF_Actions_MulChai ] ) ;
end.


