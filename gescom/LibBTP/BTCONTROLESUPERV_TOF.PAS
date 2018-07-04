{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/11/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTCONTROLESUPERV ()
Mots clefs ... : TOF;BTCONTROLESUPERV
*****************************************************************}
Unit BTCONTROLESUPERV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     AglInit,
     HCtrls,
     HEnt1,
     HMsgBox,
     vierge,
     LicUtil,
     UTOF ;

Type
  TOF_BTCONTROLESUPERV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function CheckSuperviseur (var StSuperviseur : string) : boolean;

Implementation

function CheckSuperviseur (var StSuperviseur : string) : boolean;
var TOBC : TOB;
begin
  result := false;
  TOBC := TOB.Create ('UN ECHANGE',nil,-1);
  TOBC.AddChampSupValeur('OKOK','-');
  TOBC.AddChampSupValeur('SUPERVISEUR','-');
  TheTOB := TOBC;
  TRY
    AGLLanceFiche('BTP','BTCONTROLESUPERV','','','');
  FINALLY
    TheTOB := nil;
    result :=  (TOBC.GEtString('OKOK')='X');
    if result then StSuperviseur := TOBC.GetString('SUPERVISEUR');
    TOBC.Free;
  end;
end;


procedure TOF_BTCONTROLESUPERV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLESUPERV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLESUPERV.OnUpdate ;
var Sql : string;
begin
  Inherited ;
  if (GetControlText('USER')='')  then
  begin
    PgiInfo('Merci de renseigner le compte et le mot de passe du superviseur');
    TFVierge (ecran).ModalResult := 0;
    exit;
  end;
  Sql := 'SELECT 1 FROM UTILISAT WHERE US_ABREGE="'+ GetControlText('USER')+'" AND ';
  if GetControlText('PASSW')='' then
  begin
    Sql := Sql + '((US_PASSWORD="'+CryptageSt('CGDPGIMDP')+'") OR (US_PASSWORD="")) AND ';
  end else
  begin
    Sql := Sql + 'US_PASSWORD="'+CryptageSt(GetControlText('PASSW'))+'" AND ';
  end;
  Sql := Sql + 'US_SUPERVISEUR="X"';
  If not ExisteSql (Sql) then
  begin
    PGIError('ERREUR : Compte ou mot de passe invalide');
    TFVierge (ecran).ModalResult := 0;
    exit;
  end;
  LaTOB.setString('OKOK','X');
  LaTOB.SetString('SUPERVISEUR',GetControlText('USER'));
  TFVierge (ecran).Close;
end ;

procedure TOF_BTCONTROLESUPERV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLESUPERV.OnArgument (S : String ) ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLESUPERV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLESUPERV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTCONTROLESUPERV.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTCONTROLESUPERV ] ) ;
end.
