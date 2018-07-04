{***********UNITE*************************************************
Auteur  ...... : G.Harlez
Créé le ...... : 02/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CREAT_GRPTRAVAIL ()
Mots clefs ... : TOF;CREAT_GRPTRAVAIL
*****************************************************************}
Unit CREAT_GRPTRAVAIL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     Utob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     VIERGE,
     windows;

Type
  TOF_CREAT_GRPTRAVAIL = Class (TOF)
    Obj_BtValid  : TToolbarButton97;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure DoValider (sender : Tobject);
    procedure DoKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

Implementation
////////////////////////////////////////////////////////////////////////////////
procedure TOF_CREAT_GRPTRAVAIL.OnArgument (S : String ) ;
begin
  Inherited ;
  Obj_BtValid := TToolbarButton97(GetControl('BTVALID'));
  if Obj_BtValid <> nil then
    Obj_BtValid.OnClick := DoValider;

  TForm(ecran).OnKeyDown := DoKeyDown;

  if s = 'MENU' then
    Ecran.Caption := 'Ajouter un Groupe de travail'
  else if s = 'SOUS_MENU' then
    Ecran.Caption := 'Ajouter un Sous-Groupe de travail'
  else
  begin
    Ecran.Caption := 'Modification';
    THEdit(GetControl('EDT_CODE')).Enabled := FALSE;
    THEdit(GetControl('EDT_LIBELLE')).Text := Copy(S,1,pos('[',S)-1);
    THEdit(GetControl('EDT_CODE')).Text    := Copy(S,pos('[',S)+1,length(S)-pos('[',S)-1);
  end;
end ;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_CREAT_GRPTRAVAIL.DoValider(sender: Tobject);
begin
  if GetControlText('EDT_CODE') = '' then
  begin
    PGIBOX('Veuillez renseigner le code','Saisie des groupes de travail');
    SetFocusControl('EDT_CODE');
  end
  else if GetControlText('EDT_LIBELLE') = '' then
  begin
    PGIBOX('Veuillez renseigner le libellé','Saisie des groupes de travail');
    SetFocusControl('EDT_LIBELLE');
  end
  else
  begin
    TFVierge(ecran).retour := GetControlText('EDT_LIBELLE')+'['+GetControlText('EDT_CODE')+']';
    TForm(ecran).Close;

  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_CREAT_GRPTRAVAIL.DoKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case key of
    VK_RETURN : DoValider(sender);
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOF_CREAT_GRPTRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CREAT_GRPTRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CREAT_GRPTRAVAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CREAT_GRPTRAVAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CREAT_GRPTRAVAIL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CREAT_GRPTRAVAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CREAT_GRPTRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;



Initialization
  registerclasses ( [ TOF_CREAT_GRPTRAVAIL ] ) ; 
end.
