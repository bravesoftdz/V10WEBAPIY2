{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/01/2008
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREFCDEBATIPRIX ()
Mots clefs ... : TOF;BTREFCDEBATIPRIX
*****************************************************************}
Unit BTREFCDEBATIPRIX_TOF ;

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
     ParamSoc,
     UTOF ; 

Type
  TOF_BTREFCDEBATIPRIX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBSOC : TOB;
    procedure GetComposants;
    procedure ChargeInfos;
    procedure SaveInfos;
    procedure MessageErreur(CodeErreur: integer);
    procedure GetInfoSocParamSoc;
  end ;

Implementation

procedure TOF_BTREFCDEBATIPRIX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREFCDEBATIPRIX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREFCDEBATIPRIX.OnUpdate ;
begin
  Inherited ;
  if THEDit(GetControl('SO_LIBELLE')).Text        = '' then begin MessageErreur(1); Exit; end;
  if THEDit(GetControl('SO_ADRESSE1')).Text       = '' then BEGIN MessageErreur(2); Exit; end;
  if THEDit(GetControl('SO_CODEPOSTAL')).Text     = '' then BEGIN MessageErreur(3); Exit; end;
  if THEDit(GetControl('SO_VILLE')).Text          = '' then BEGIN MessageErreur(4); Exit; end;
  if THValComBoBox(GetControl('SO_PAYS')).Value   = '' then BEGIN MessageErreur(5); Exit; end;
  if THEdit(GetControl('SO_TELEPHONE')).Text      = '' then BEGIN MessageErreur(6); Exit; end;
  if THEdit(GetControl('SO_CONTACT')).text        = '' then BEGIN MessageErreur(7); Exit; end;
  if THEdit(GetControl('SO_PRENOMCONTACT')).text  = '' then BEGIN MessageErreur(8); Exit; end;
  SaveInfos;
end ;


procedure TOF_BTREFCDEBATIPRIX.MessageErreur (CodeErreur : integer);
var LibMessage : string;
begin
  if CodeErreur = 1 then
  begin
    LibMessage := 'Veuillez renseigner la désignation de la société';
    THEDit(GetControl('SO_LIBELLE')).SetFocus;
  end else if CodeErreur = 2 then
  begin
    LibMessage := 'Veuillez renseigner l''adresse de la société';
    THEDit(GetControl('SO_ADRESSE1')).SetFocus;
  end else if CodeErreur = 3 then
  begin
    LibMessage := 'Veuillez renseigner le code postal';
    THEDit(GetControl('SO_CODEPOSTAL')).SetFocus;
  end else if CodeErreur = 4 then
  begin
    LibMessage := 'Veuillez renseigner la ville';
    THEDit(GetControl('SO_VILLE')).SetFocus;
  end else if CodeErreur = 5 then
  begin
    LibMessage := 'Veuillez renseigner le pays';
    THEDit(GetControl('SO_PAYS')).SetFocus;
  end else if CodeErreur = 6 then
  begin
    LibMessage := 'Veuillez renseigner le téléphone';
    THEDit(GetControl('SO_TELEPHONE')).SetFocus;
  end else if CodeErreur = 7 then
  begin
    LibMessage := 'Veuillez renseigner le nom du contact principal';
    THEDit(GetControl('SO_CONTACT')).SetFocus;
  end else if CodeErreur = 8 then
  begin
    LibMessage := 'Veuillez renseigner le prénom du contact principal';
    THEDit(GetControl('SO_PRENOMCONTACT')).SetFocus;
  end;

  PgiInfo(Libmessage);
  Ecran.ModalResult := 0;
end;


procedure TOF_BTREFCDEBATIPRIX.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREFCDEBATIPRIX.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBSOC := TOB.Create ('SOCIETE',nil,-1);
  ChargeInfos;
end ;

procedure TOF_BTREFCDEBATIPRIX.OnClose ;
begin
  TOBSOC.free;
  Inherited ;
end ;

procedure TOF_BTREFCDEBATIPRIX.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREFCDEBATIPRIX.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREFCDEBATIPRIX.ChargeInfos;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT * FROM SOCIETE',true);
  if not QQ.eof then
  begin
    TOBSOC.SelectDb ('',QQ);
  end;
  Ferme (QQ);
  if TOBSOc.getValue('SO_LIBELLE')='' then
  begin
    GetInfoSocParamSoc;
  end;

  TOBSoc.PutEcran (ecran);
end;

procedure TOF_BTREFCDEBATIPRIX.SaveInfos;
begin
  TOBSoc.GetEcran (ecran);
  TOBSOC.UpdateDB;
end;

procedure TOF_BTREFCDEBATIPRIX.GetComposants;
begin
//
end;

procedure TOF_BTREFCDEBATIPRIX.GetInfoSocParamSoc;
var Valeur : string;
    DValeur : double;
begin
  Valeur := GetParamSoc('SO_SOCIETE');
  TOBSOC.putValue('SO_SOCIETE',Valeur);
  //
  Valeur := GetParamSoc('SO_LIBELLE');
  TOBSOC.putValue('SO_LIBELLE',Valeur);
  //
  Valeur := GetParamSoc('SO_ADRESSE1');
  TOBSOC.putValue('SO_ADRESSE1',Valeur);
  //
  Valeur := GetParamSoc('SO_ADRESSE2');
  TOBSOC.putValue('SO_ADRESSE2',Valeur);
  //
  Valeur := GetParamSoc('SO_ADRESSE3');
  TOBSOC.putValue('SO_ADRESSE3',Valeur);
  //
  Valeur := GetParamSoc('SO_DIVTERRIT');
  TOBSOC.putValue('SO_DIVTERRIT',Valeur);
  //
  Valeur := GetParamSoc('SO_CODEPOSTAL');
  TOBSOC.putValue('SO_CODEPOSTAL',Valeur);
  //
  Valeur := GetParamSoc('SO_VILLE');
  TOBSOC.putValue('SO_VILLE',Valeur);
  //
  Valeur := GetParamSoc('SO_PAYS');
  TOBSOC.putValue('SO_PAYS',Valeur);
  //
  Valeur := GetParamSoc('SO_TELEPHONE');
  TOBSOC.putValue('SO_TELEPHONE',valeur);
  //
  Valeur := GetParamSoc('SO_FAX');
  TOBSOC.putValue('SO_FAX',Valeur);
  //
  Valeur := GetParamSoc('SO_TELEX');
  TOBSOC.putValue('SO_TELEX',Valeur);
  //
  Valeur := GetParamSoc('SO_MAIL');
  TOBSOC.putValue('SO_MAIL',Valeur);
  //
  Valeur := GetParamSoc('SO_RVA');
  TOBSOC.putValue('SO_RVA',Valeur);
  //
  Valeur := GetParamSoc('SO_CONTACT');
  TOBSOC.putValue('SO_CONTACT',Valeur);
  //
  //FV1 : 16/03/2017 -  FS#2391 - SES ETANCH - Batiprix ne conserve pas le prénom dans informations société
  Valeur := GetParamSoc('SO_PRENOMCONTACT');
  TOBSOC.putValue('SO_PRENOMCONTACT',Valeur);
  //
  Valeur := GetParamSoc('SO_NIF');
  TOBSOC.putValue('SO_NIF',Valeur);
  //
  Valeur := GetParamSoc('SO_APE');
  TOBSOC.putValue('SO_APE',Valeur);
  //
  Valeur := GetParamSoc('SO_SIRET');
  TOBSOC.putValue('SO_SIRET',Valeur);
  //
  DValeur := GetParamSoc('SO_CAPITAL');
  TOBSOC.putValue('SO_CAPITAL',DValeur);
  //
  Valeur := GetParamSoc('SO_RC');
  TOBSOC.putValue('SO_RC',Valeur);
  //
  Valeur := GetParamSoc('SO_TXTJURIDIQUE');
  TOBSOC.putValue('SO_TXTJURIDIQUE',Valeur);
  //
end;

Initialization
  registerclasses ( [ TOF_BTREFCDEBATIPRIX ] ) ;
end.
