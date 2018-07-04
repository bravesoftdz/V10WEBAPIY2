{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ACTIONCONTREM ()
Mots clefs ... : TOF;ACTIONCONTREM
*****************************************************************}
Unit ActionContreM_Tof ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
     maineagl,
{$ELSE}
     db,dbtables,FE_main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,
     EntGC,Vierge,UtilConfid ;

function GCLanceFiche_ActionContreM(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_ACTIONCONTREM = Class (TOF)
  protected
    NaturePiece : string ;
    RBCONSULT : TRadioButton;
    RBMODIF : TRadioButton;
    RBGENPRE : TRadioButton;
    RBGENBLC : TRadioButton;
    RBGENBLF : TRadioButton;
    RBGENFAC : TRadioButton;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArg : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure RadioEnter (Sender: TObject);
  end ;

Implementation

function GCLanceFiche_ActionContreM(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_ACTIONCONTREM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ACTIONCONTREM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ACTIONCONTREM.OnUpdate ;
begin
  Inherited ;
TFVierge(Ecran).Retour:='' ;
if RBCONSULT.Checked then TFVierge(Ecran).Retour:='CON'
else if RBMODIF.Checked then TFVierge(Ecran).Retour:='MOD'
else if RBGENPRE.Checked then TFVierge(Ecran).Retour:='PRE'
else if RBGENBLF.Checked then TFVierge(Ecran).Retour:='BLF'
else if RBGENBLC.Checked then TFVierge(Ecran).Retour:='BLC'
else if RBGENFAC.Checked then TFVierge(Ecran).Retour:='FAC';
end ;

procedure TOF_ACTIONCONTREM.OnLoad ;
Var NatSuivante,VenteAchat,St : string;
begin
  Inherited ;
NatSuivante:=GetInfoParPiece(NaturePiece, 'GPP_NATURESUIVANTE') ;
VenteAchat:=GetInfoParPiece(NaturePiece, 'GPP_VENTEACHAT') ;
if VenteAchat='VEN' then
   begin
   St:=ReadTokenSt(NatSuivante);
   While st<>'' do
       begin
       if St<>Naturepiece then SetControlVisible('PGEN'+St,True);
       St:=ReadTokenSt(NatSuivante);
       end;
   end else if Naturepiece='CF' then
   begin
   St:=ReadTokenSt(NatSuivante);
   While st<>'' do
       begin
       if St<>Naturepiece then SetControlVisible('PGEN'+St,True);
       St:=ReadTokenSt(NatSuivante);
       end;
   end;
{$IFNDEF CCS3}
AppliquerConfidentialite(Ecran,'');
{$ENDIF}
if (GetControlVisible('PCONSULT')) and (GetControlEnabled('PCONSULT')) then SetControlChecked('XXRBCONSULT',True)
else if (GetControlVisible('PMODIF')) and (GetControlEnabled('PMODIF')) then SetControlChecked('XXRBMODIF',True)
else if (GetControlVisible('PGENPRE')) and (GetControlEnabled('PGENPRE')) then SetControlChecked('XXRBGENPRE',True)
else if (GetControlVisible('PGENBLC')) and (GetControlEnabled('PGENBLC')) then SetControlChecked('XXRBGENBLC',True)
else if (GetControlVisible('PGENBLF')) and (GetControlEnabled('PGENBLF')) then SetControlChecked('XXRBGENBLF',True)
else if (GetControlVisible('PGENFAC')) and (GetControlEnabled('PGENFAC')) then SetControlChecked('XXRBGENFAC',True) ;
SetControlEnabled('XXRBCONSULT',GetControlEnabled('PCONSULT')) ;
SetControlEnabled('XXRBMODIF',GetControlEnabled('PMODIF')) ;
SetControlEnabled('XXRBGENPRE',GetControlEnabled('PGENPRE')) ;
SetControlEnabled('XXRBGENBLC',GetControlEnabled('PGENBLC')) ;
SetControlEnabled('XXRBGENBLF',GetControlEnabled('PGENBLF')) ;
SetControlEnabled('XXRBGENFAC',GetControlEnabled('PGENFAC')) ;
SetControlProperty('PCONSULT','Align',alTop);
SetControlProperty('PMODIF','Align',alTop);
SetControlProperty('PGENPRE','Align',alTop);
SetControlProperty('PGENBLF','Align',alTop);
SetControlProperty('PGENBLC','Align',alTop);
SetControlProperty('PGENFAC','Align',alTop);
SetControlProperty('PCONSULT','TabOrder',0);
SetControlProperty('PMODIF','TabOrder',1);
SetControlProperty('PGENPRE','TabOrder',2);
SetControlProperty('PGENBLF','TabOrder',3);
SetControlProperty('PGENBLC','TabOrder',4);
SetControlProperty('PGENFAC','TabOrder',5);
SetControlVisible('GBACTION',True) ;
end ;

procedure TOF_ACTIONCONTREM.OnArgument (StArg : String ) ;
begin
  Inherited ;
NaturePiece:=StArg;
SetControlVisible('GBACTION',False) ;
SetControlProperty('PGENFAC','Align',alNone);
SetControlProperty('PGENBLC','Align',alNone);
SetControlProperty('PGENBLF','Align',alNone);
SetControlProperty('PGENPRE','Align',alNone);
SetControlProperty('PMODIF','Align',alNone);
SetControlProperty('PCONSULT','Align',alNone);
SetControlVisible('PCONSULT',True) ;
SetControlVisible('PMODIF',True) ;
SetControlVisible('PGENPRE',False) ;
SetControlVisible('PGENBLF',False) ;
SetControlVisible('PGENBLC',False) ;
SetControlVisible('PGENFAC',False) ;
RBCONSULT:=TRadioButton(GetControl('XXRBCONSULT'));
RBMODIF:=TRadioButton(GetControl('XXRBMODIF'));
RBGENPRE:=TRadioButton(GetControl('XXRBGENPRE'));
RBGENBLF:=TRadioButton(GetControl('XXRBGENBLF'));
RBGENBLC:=TRadioButton(GetControl('XXRBGENBLC'));
RBGENFAC:=TRadioButton(GetControl('XXRBGENFAC'));
RBCONSULT.OnEnter:=RadioEnter ;
RBMODIF.OnEnter:=RadioEnter ;
RBGENPRE.OnEnter:=RadioEnter ;
RBGENBLF.OnEnter:=RadioEnter ;
RBGENBLC.OnEnter:=RadioEnter ;
RBGENFAC.OnEnter:=RadioEnter ;
TFVierge(Ecran).Retour:='' ;
end ;

procedure TOF_ACTIONCONTREM.OnClose ;
begin
  Inherited ;
end ;

{==============================================================================================}
{============================= Evènements RadioButton =========================================}
{==============================================================================================}
procedure TOF_ACTIONCONTREM.RadioEnter(Sender: TObject);
begin
if (TRadioButton(Sender).Checked) then exit;
SetControlChecked('XXRBCONSULT',False);
SetControlChecked('XXRBMODIF',False);
SetControlChecked('XXRBGENPRE',False);
SetControlChecked('XXRBGENBLF',False);
SetControlChecked('XXRBGENBLC',False);
SetControlChecked('XXRBGENFAC',False);
SetControlChecked(TRadioButton(Sender).Name,True) ;
end;


Initialization
  registerclasses ( [ TOF_ACTIONCONTREM ] ) ;
end.
