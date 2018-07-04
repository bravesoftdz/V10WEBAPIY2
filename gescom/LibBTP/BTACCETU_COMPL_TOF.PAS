{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/09/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTACCETU_COMPL ()
Mots clefs ... : TOF;BTACCETU_COMPL
*****************************************************************}
Unit BTACCETU_COMPL_TOF ;

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
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     vierge,
     utofAfBaseCodeAffaire ;

Type

  TOF_BTACCETU_COMPL = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  public
  	procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;

  end ;

Implementation
uses calcolegenericBTP;

procedure TOF_BTACCETU_COMPL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTACCETU_COMPL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTACCETU_COMPL.OnUpdate ;
begin
  Inherited ;
  if StrToDate(THedit(getControl('DATEACCEPT')).Text) > StrToDate(THedit(getControl('DATEDEBUT')).Text) then
  begin
  	PgiError ('La Date de début du contrat doit être ultérieure à la date d''acceptation');
    TFvierge(ecran).ModalResult := 0;
    exit;
  end;
  if THedit(getControl('CH_CHANTIER1')).Text = '' then
  begin
  	PgiError ('Vous devez indiquer une affaire !');
    TFvierge(ecran).ModalResult := 0;
    exit;
  end;
  laTOB.putValue('AFFAIRE',THedit(getControl('CH_CHANTIER')).Text);
  laTOB.putValue('DATE',StrToDate(THedit(getControl('DATEACCEPT')).Text));
  laTOB.putValue('DATECON',StrToDate(THedit(getControl('DATEDEBUT')).Text));
  laTOB.PutValue('OK','X');
end ;

procedure TOF_BTACCETU_COMPL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTACCETU_COMPL.OnArgument (S : String ) ;
var     Aff0,Aff1,Aff2,Aff3,Aff4 : string;

begin
  Inherited ;
  THedit(getControl('TIERS')).Text := laTOB.getValue('TIERS');
  THedit(getControl('LIBELLETIERS')).Text := laTOB.getValue('LIBELLETIERS');
  THedit(getControl('CH_CHANTIER')).Text := laTOB.getValue('AFFAIRE');
  BTPCodeAffaireDecoupe (laTOB.getValue('AFFAIRE'),Aff0,Aff1,Aff2,Aff3,Aff4,tacreat,false);
  THedit(getControl('CH_CHANTIER0')).Text := Aff0; // grmfff
  THedit(getControl('CH_CHANTIER1')).Text := Aff1;
  THedit(getControl('CH_CHANTIER2')).Text := Aff2;
  THedit(getControl('CH_CHANTIER3')).Text := Aff3;
  THedit(getControl('CH_AVENANT')).Text := Aff4;
  if Aff2 = '' then
  begin
  	THedit(getControl('CH_CHANTIER2')).Text := laTOB.getValue('CODEAFFAIRE2');
    THedit(getControl('CH_CHANTIER2')).enabled := false;
  end;

  ecran.caption := ecran.caption + ' N° '+IntToStr(laTOB.getValue('NUMETUDE'));
  UpdateCaption(ecran);
end ;

procedure TOF_BTACCETU_COMPL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTACCETU_COMPL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTACCETU_COMPL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTACCETU_COMPL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
Aff  := THEdit(GetControl('CH_CHANTIER'));
Aff0 := THEdit(GetControl('CH_CHANTIER0'));
Aff1 := THEdit(GetControl('CH_CHANTIER1'));
Aff2 := THEdit(GetControl('CH_CHANTIER2'));
Aff3 := THEdit(GetControl('CH_CHANTIER3'));
Aff4 := THEdit(GetControl('CH_AVENANT'));
Tiers := THEdit(GetControl('TIERS'));
end;

Initialization
  registerclasses ( [ TOF_BTACCETU_COMPL ] ) ; 
end.
