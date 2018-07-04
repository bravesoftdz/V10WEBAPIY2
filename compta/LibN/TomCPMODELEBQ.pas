unit TomCPMODELEBQ;

interface
uses
    Classes, Dialogs, stdctrls, Sysutils,
{$IFDEF EAGLCLIENT}
{$ELSE}
    db, dbctrls, dbtables,FichList, HDB,
{$ENDIF}
    UTOM, UTOB, HmsgBox, Hctrls, Ent1, HEnt1;

type
  TOM_CPMODELEBQ = class(TOM)
    procedure OnUpdateRecord ; override ;
//    procedure OnLoadRecord ; override ;
    procedure OnChangeField (F : TField); override ;
    procedure OnNewRecord  ; override ;
    procedure OnDeleteRecord ; override ;
    procedure OnArgument(Arguments : String ); override ;
    procedure OnClose ; override ;

  private
    General: THDBEdit;
    procedure CliFouOnClick(Sender :TObject);

End;

implementation

procedure TOM_CPMODELEBQ.OnArgument(Arguments: String);
VAR CliFou : TDBCheckBox;
begin
inherited;
CliFou:=TDBCheckBox(GetControl('CMB_CLIFOU'));
General:=THDBEdit(GetControlText('CMB_GENERAL'));
if CliFou<>nil  then CliFou.OnClick:=CliFouOnClick;
end;

procedure TOM_CPMODELEBQ.OnChangeField(F: TField);
VAR
  OkBq,OkBqCli,OkLcr,OkBor : boolean;
  St:String;
Begin
OkBqCli:=False;OkBq:=False;OkLcr:=False;OkBor:=False;
if F.FieldName = 'CMB_GENERAL' then
begin
St:=copy(GetControlText('CMB_GENERAL'),1,2);
if St<>'' then
begin
  if (St='51') and (GetField('CMB_CLIFOU')='X') then OkBqCli:=True;
  if (St='51') and (GetField('CMB_CLIFOU')='-') then OkBq:=True;
  if St='41' then begin OkLcr:=True;SetField('CMB_CLIFOU','X');end;
  if St='40' then begin OkBor:=True;SetField('CMB_CLIFOU','-');end;
  SetControlEnabled('CMB_DCLILCR',OkLcr);
  SetControlEnabled('CMB_DCLIVIR',OkBqCli);
  SetControlEnabled('CMB_BCLICHQ',OkBqCli);
  SetControlEnabled('CMB_BCLILCR',OkLcr);
  SetControlEnabled('CMB_BCLIVIR',OkBqCli);
  SetControlEnabled('TCLICHQ',OkBqCli);
  SetControlEnabled('TCLILCR',OkLcr);
  SetControlEnabled('TCLIVIR',OkBqCli);
  SetControlEnabled('CMB_DFOUCHQ',OkBq);
  SetControlEnabled('CMB_DFOUBOR',OkBor);
  SetControlEnabled('CMB_DFOUPRE',OkBq);
  SetControlEnabled('CMB_BFOUBOR',OkBor);
  SetControlEnabled('CMB_BFOUPRE',OkBq);
  SetControlEnabled('TFOUCHQ',OkBq);
  SetControlEnabled('TFOUBOR',OkBor);
  SetControlEnabled('TFOUPRE',OkBq);
end;
end;
End;

procedure TOM_CPMODELEBQ.CliFouOnClick(Sender :TObject);
VAR OkCli,OkFou : boolean;
begin
inherited;
OkCli:=False;OkFou:=False;
if TCheckBox(GetControl('CMB_CLIFOU')).Checked then OkCli:=True else OkFou:=True;
SetControlEnabled('CMB_DCLILCR',OkCli);
SetControlEnabled('CMB_DCLIVIR',OkCli);
SetControlEnabled('CMB_BCLICHQ',OkCli);
SetControlEnabled('CMB_BCLILCR',OkCli);
SetControlEnabled('CMB_BCLIVIR',OkCli);
SetControlEnabled('TCLICHQ',OkCli);
SetControlEnabled('TCLILCR',OkCli);
SetControlEnabled('TCLIVIR',OkCli);
SetControlEnabled('CMB_DFOUCHQ',OkFou);
SetControlEnabled('CMB_DFOUBOR',OkFou);
SetControlEnabled('CMB_DFOUPRE',OkFou);
SetControlEnabled('CMB_BFOUBOR',OkFou);
SetControlEnabled('CMB_BFOUPRE',OkFou);
SetControlEnabled('TFOUCHQ',OkFou);
SetControlEnabled('TFOUBOR',OkFou);
SetControlEnabled('TFOUPRE',OkFou);
end;

procedure TOM_CPMODELEBQ.OnClose ;
begin
inherited ;
end ;

procedure TOM_CPMODELEBQ.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_CPMODELEBQ.OnNewRecord;
begin
  inherited;
end;

procedure TOM_CPMODELEBQ.OnUpdateRecord;
begin
  inherited;
if GetControlText('CMB_GENERAL') = '' then begin
  PGIBox('Vous devez renseigner un compte général', 'Paramétrage des Modèles de Banque');
  SetFocusControl('CMB_GENERAL');
  Abort;
end;
if GetControlText('CMB_DEVISE') = '' then begin
  PGIBox('Vous devez renseigner une devise', 'Paramétrage des Modèles de Banque');
  SetFocusControl('CMB_DEVISE');
  Abort;
end;

if not ExisteSql('Select G_GENERAL from GENERAUX Where G_GENERAL="'+GetControlText('CMB_GENERAL')+'" ') then
  begin
  LastError:=1;
  PGIBox('Le compte général n''existe pas.', 'Paramétrage des Modèles de Banque');
  Abort;
  end;
if PresenceComplexe('CPMODELEBQ',
                   ['CMB_GENERAL','CMB_DEVISE','CMB_CLIFOU'],
                   ['=','=','='],
                   [GetControlText('CMB_GENERAL'), GetControlText('CMB_DEVISE'),GetControlText('CMB_CLIFOU')],
                   ['S','S','S']) then begin
  LastError:=1;
  PGIBox('Cette combinaison Compte-Devise-Client existe déjà.', 'Paramétrage des Modèles de Banque');
  Abort;
end;
end;



Initialization
registerclasses([TOM_CPMODELEBQ]);


end.
