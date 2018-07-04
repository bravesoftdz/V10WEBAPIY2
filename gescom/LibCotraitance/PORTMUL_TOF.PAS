{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 01/04/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PORTMUL ()
Mots clefs ... : TOF;PORTMUL
*****************************************************************}
Unit PORTMUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     HDB,
     fe_main,
{$else}
     eMul,
     uTob,
     MailEagl,
{$ENDIF}
		 HTB97,
     forms,
     uTOFComm,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type

  TOF_PORTMUL = Class (tTOFComm)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	fTypegestion : string;
    procedure FListeDblClick( Sender : Tobject);
    procedure BinsertClick(Sender : Tobject);
  end ;

Implementation

procedure TOF_PORTMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.OnArgument (S : String ) ;
begin
	fMulDeTraitement := true;
  Inherited ;
  if pos('COTRAITANCE',S) > 0 then
  begin
  	fTypegestion := ';COTRAITANCE';
    SetControlText('GPO_TYPEFRAIS','B00');
    ecran.Caption := 'Frais de gestion de cotraitance';
    updatecaption (ecran);
  end else
  begin
  	fTypegestion := '';
    THValComboBox(getControl('GPO_TYPEFRAIS')).Plus := ' AND (CO_CODE <> "B00")';
    SetControlText('GPO_TYPEFRAIS','501');
  end;
  THDBGrid(getControl('Fliste')).OnDblClick := FListeDblClick;
  TToolbarButton97(getControl('Binsert')).OnClick := BinsertClick;

end ;

procedure TOF_PORTMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PORTMUL.FListeDblClick(Sender: Tobject);
begin
  AglLanceFiche('GC','GCPORT','',tfmul(ecran).Q.FindField('GPO_CODEPORT').asstring,'ACTION=MODIFICATION'+fTypegestion);
  RefreshDB;
end;

procedure TOF_PORTMUL.BinsertClick(Sender: Tobject);
begin
  AglLanceFiche('GC','GCPORT','','','ACTION=CREATION'+fTypegestion);
  RefreshDB;
end;

Initialization
  registerclasses ( [ TOF_PORTMUL ] ) ;
end.
