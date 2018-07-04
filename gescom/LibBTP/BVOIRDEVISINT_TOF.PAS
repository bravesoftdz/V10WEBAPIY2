{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/02/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BVOIRDEVISINT ()
Mots clefs ... : TOF;BVOIRDEVISINT
*****************************************************************}
Unit BVOIRDEVISINT_TOF ;

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
     HRichEdt,
     HRichOLE,
     Vierge,
     UTOF ;

Type
  TOF_BVOIRDEVISINT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_BVOIRDEVISINT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRDEVISINT.OnDelete ;
var Sql : string;
begin
  Inherited ;
  if PgiAsk('Etes-vous sur de considérer le devis comme traité ?')=mryes then
  begin
    SQL := 'UPDATE LIENSOLE SET LO_QUALIFIANTBLOB="REA" WHERE '+
           'LO_TABLEBLOB ="APP" AND '+
           'LO_RANGBLOB=20 AND '+
           'LO_IDENTIFIANT="'+LaTOB.GetString('Lo_IDENTIFIANT')+'"';
    LaTOB.setString('LO_QUALIFIANTBLOB','REA');
    ExecuteSql (SQL);
    SetControlVisible('BDELETE',false);
  end;
end ;

procedure TOF_BVOIRDEVISINT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRDEVISINT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRDEVISINT.OnArgument (S : String ) ;
begin
  Inherited ;
  StringToRich(THRichEditOLE (GetControl('DEVIS')), LaTOB.Getvalue('LO_OBJET'));
  if laTOB.getString('LO_QUALIFIANTBLOB')<>'REA' then SetControlVisible('BDELETE',true);
  THRichEditOLE(GetControl('DEVIS')).ReadOnly := True;
end ;

procedure TOF_BVOIRDEVISINT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRDEVISINT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BVOIRDEVISINT.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BVOIRDEVISINT ] ) ; 
end.
