{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTVENTILCPTAS1 ()
Mots clefs ... : TOF;BTVENTILCPTAS1
*****************************************************************}
Unit BTVENTILCPTAS1_TOF ;

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
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UtilSoc,
     BTPARAMSOCGEN_TOF,
     UtilLine,
     UTOF ;

Type
  TOF_BTVENTILCPTAS1 = Class (TOF_BTPARAMSOCGEN)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	function VerifInfos : boolean;
  end ;

Implementation

procedure TOF_BTVENTILCPTAS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCPTAS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCPTAS1.OnUpdate ;
begin
  Inherited ;

  if verifInfos then
     StockeInfos(Ecran, LaTob)
  else
  	TForm(Ecran).ModalResult:=0;

end ;

procedure TOF_BTVENTILCPTAS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCPTAS1.OnArgument (S : String ) ;
begin
  Inherited ;
  ChargeEcran(ecran, LaTob);
end ;

procedure TOF_BTVENTILCPTAS1.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCPTAS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTVENTILCPTAS1.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_BTVENTILCPTAS1.VerifInfos: boolean;
var FF : TForm;
begin
	result := false;
  FF := ecran;

  if Not GenExiste('SO_GCCPTEESCACH',FF,True,'le Compte escompte Achat n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEREMACH',FF,True,'le Compte Remise Achat n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEHTACH',FF,True,'le Compte H.T. Achat n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEPORTACH',FF,True,'le Compte Port Achat n''existe pas') then Exit ;
  if (GetControlText('GPP_JOURNALCPTAA')='') then
  BEGIN
  	PGiError('Vous devez renseigner un journal de ventilation des Achats',ecran.caption);
    exit;
  END;

  if Not GenExiste('SO_GCCPTEESCVTE',FF,True,'le Compte escompte Vente n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEREMVTE',FF,True,'le Compte Remise Vente n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEHTVTE',FF,True,'le Compte H.T. Vente n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEPORTVTE',FF,True,'le Compte Port Vente n''existe pas') then Exit ;
  if GetControltext('SO_GCECARTDEBIT') <> '' then
  	if Not GenExiste('SO_GCECARTDEBIT',FF,True,'le Compte d''écart débiteur n''existe pas') then Exit ;
  if GetControltext('SO_GCECARTCREDIT') <> '' then
  	if Not GenExiste('SO_GCECARTCREDIT',FF,True,'le Compte d''écart débiteur n''existe pas') then Exit ;
  if (GetControlText('GPP_JOURNALCPTAV')='') then
  BEGIN
  	PGiError('Vous devez renseigner un journal de ventilation des ventes',ecran.caption);
    exit;
  END;
  result := true;
end;

Initialization
  registerclasses ( [ TOF_BTVENTILCPTAS1 ] ) ; 
end.

