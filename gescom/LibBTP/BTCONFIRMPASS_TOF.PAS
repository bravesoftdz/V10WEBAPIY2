{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/09/2009
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONFIRMPASS ()
Mots clefs ... : TOF;BTCONFIRMPASS
*****************************************************************}
Unit BTCONFIRMPASS_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     LicUtil,
     UTOF;

Type
  TOF_BTCONFIRMPASS = Class (TOF)
  private
  	TOBOMFG : TOB;
    IsDayPass : boolean;
    Irpt : integer;
    ePwd : THEdit;
    Complement : string;

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function IsOkDayPass(Complement : string='') : Boolean;

Implementation

uses
  wCommuns
  , UtilPGI
  ;

function IsOkDayPass(Complement : string='') : Boolean;
var
  OneTOB : TOB;
begin
	OneTOB := TOB.Create('UNE TOB',nil,-1);
  try
    OneTOB.AddChampSupValeur('RESULTAT','NOK');
    OneTOB.AddChampSupValeur('DAYPASS','YES');
    TheTOB := OneTOB;
    AGLLanceFiche('BTP','BTCONFIRMPASS','','',iif(Complement <>'', 'COMPLEMENT=' + Complement, ''));
    TheTOB := nil;
    Result := (OneTOB.GetString('RESULTAT')='OK'); 
  finally
    OneTOB.Free;
  end;
end;

procedure TOF_BTCONFIRMPASS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMPASS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMPASS.OnUpdate ;
begin
  Inherited ;
  if (IsDayPass) then
  begin
    if (THEdit(GetControl('EPASSWORD')).Text) <> (DayPass(Now) + Complement) then
    begin
    	THLabel(getCOntrol('TMESSAGEBAD')).visible := true;
    	THLabel(getCOntrol('TMESSAGEBAD')).Caption := 'Mot de passe invalide';
    	inc(IRpt);
      if IRpt < 4 then BEGIN ecran.ModalResult := 0; Exit; END;
    end else TOBOMFG.putValue('RESULTAT','OK');
  end else
  begin
  	TOBOMFG.putValue('RESULTAT','NOK');
  end;
  TheTOB := TOBOMFG;
end ;

procedure TOF_BTCONFIRMPASS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMPASS.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBOMFG := LaTOB;
  IsDayPass := TOBOMFG.fieldExists('DAYPASS');
  THLabel(getCOntrol('TMESSAGEBAD')).visible := false;
  Irpt :=0;
  ePwd       := THEdit(GetControl('EPASSWORD'));
  Complement := GetArgumentString(S, 'COMPLEMENT');
  ePwd.MaxLength := iif(Complement = '', ePwd.MaxLength, 0);
end ;

procedure TOF_BTCONFIRMPASS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMPASS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONFIRMPASS.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTCONFIRMPASS ] ) ; 
end.
