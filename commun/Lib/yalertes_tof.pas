Unit yalertes_tof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
{$else}
     eMul,
     uTob,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,UtilAlertes,htb97,AGLInit,ParamSoc,yalertes_tom,yalertesMail_tof,yalertesMemo_tof ;

Type
  TOF_yalertes = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
    stArgument : string;
    procedure CreationClick(Sender: TObject);
  end ;

Procedure RTLanceFiche_Alertes_Mul (Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation

Uses
  EntPgi
  ,YAlertesConst
  ;

Procedure RTLanceFiche_Alertes_Mul(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
  if not ExisteAlerteActive then
    begin
    PgiBox('Aucune alerte n''est activée','Alertes');
    exit;
    end;
  AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_yalertes.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.OnArgument (S : String ) ;
begin
  Inherited ;
{$IFNDEF GPAO}
  { Alertes - Tables gérées seulement en GP : Wxx et GPO, GEM, GDE, YTS }
  if Assigned(GetControl('YAL_PREFIXE')) then
    THMultiValcomboBox(GetControl('YAL_PREFIXE')).Plus := ' AND ' + WhereAlertesNotGP;
{$ENDIF GPAO}
  TToolBarButton97(GetControl('Binsert')).OnClick:=CreationClick;
  stArgument:='ACTION=CREATION';
  if S <> 'ACTION=CREATION' then stArgument:=stArgument+';'+S;
end ;

procedure TOF_yalertes.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_yalertes.CreationClick(Sender: TObject);
begin
    AGLLancefiche('Y','YALERTES','','',stArgument) ;
    AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

Initialization
  registerclasses ( [ TOF_yalertes ] ) ;
end.

