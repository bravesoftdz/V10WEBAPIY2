{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MUL_EMPLOIINTERIM ()
Mots clefs ... : TOF;MUL_EMPLOIINTERIM
*****************************************************************}
Unit UTOFPG_MulEmploiInterim ;

Interface

Uses StdCtrls,Controls,Classes,HTB97,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,FE_Main,
{$ELSE}
     emul,UtileAGL,MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF ;

Type
  TOF_MUL_EMPLOIINTERIM = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    procedure IdentiteClick(Sender:TObject);
  end ;

Implementation

procedure TOF_MUL_EMPLOIINTERIM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MUL_EMPLOIINTERIM.OnArgument (S : String ) ;
var Interimaire,Precedent:String;
    BIdentite:TToolBarButton97;
begin
  Inherited ;
Interimaire :=Trim(ReadTokenSt(S));  //:= stArgument;
Precedent:=trim(ReadTokenPipe(S,';'));
SetControlText('PEI_INTERIMAIRE',Interimaire);
TFMul(Ecran).Caption:='Liste des emplois de : '+Interimaire+' '+RechDom('PGINTERIMAIRES',Interimaire,False);
BIdentite:=TToolBarButton97(GetControl('BIDENTITE'));
If Precedent='FICHE' Then SetControlVisible('BIDENTITE',False);
If BIdentite<>Nil Then BIdentite.OnClick:=IdentiteClick;
end ;

procedure TOF_MUL_EMPLOIINTERIM.IdentiteClick(Sender:TObject);
var Interim:String;
begin
Interim:=GetControlText('PEI_INTERIMAIRE');
If Interim<>'' Then AGLLanceFiche('PAY','INTERIMAIRES','PSI_INTERIMAIRE='+Interim,Interim,'ACTION=CONSULTATION');
end;

Initialization
  registerclasses ( [ TOF_MUL_EMPLOIINTERIM ] ) ;
end.
