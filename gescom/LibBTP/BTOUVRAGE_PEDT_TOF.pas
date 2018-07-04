{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTOUVRAGE_PEDT ()
Mots clefs ... : TOF;BTOUVRAGE_PEDT
*****************************************************************}
Unit BTOUVRAGE_PEDT_TOF ;

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
		 Aglinit,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     UTOF ;

Type
  TOF_BTOUVRAGE_PEDT = Class (TOF)
  private
    GSOUSDET : TGroupBox;
    SCODE,SLIBELLE,SQTE,SUNITE : TCheckBox;
    VALPA,VALPR,VALPV,OPTSOUSDET : TCheckBox;
    TheTOBParam : TOB;
    procedure ClickSousDet (Sender : TObject);
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

Implementation

procedure TOF_BTOUVRAGE_PEDT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_PEDT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_PEDT.OnUpdate ;
begin
  Inherited ;
  if ValPA.Checked then  TheTOBParam.PutValue ('VALPA','X');
  if ValPR.Checked then  TheTOBParam.PutValue ('VALPR','X');
  if ValPR.Checked then TheTOBParam.PutValue ('VALPV','X');
  if OPTSOUSDET.Checked then TheTOBParam.PutValue ('OPTSOUSDET','X');
  if SCODE.Checked then TheTOBParam.PutValue ('SCODE','X');
  if SLIBELLE.Checked then TheTOBParam.PutValue ('SLIBELLE','X');
  if SQTE.Checked then TheTOBParam.PutValue ('SQTE','X');
  if SUNITE.Checked then TheTOBParam.PutValue ('SUNITE','X');
  TheTOBParam.PutValue ('OK','X');
  TheTOB := TheTOBParam;
end ;

procedure TOF_BTOUVRAGE_PEDT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_PEDT.OnArgument (S : String ) ;
begin
  Inherited ;
  TheTOBParam := LaTOB;
	GSOUSDET := TGroupBox(getControl('GSOUSDET'));
  GSOUSDET.enabled := false;
  OPTSOUSDET := TCheckBox(GetControl('OPTSOUSDET'));
  OPTSOUSDET.OnClick := ClickSousDet;
  //
  SCODE := TCheckBox(GetControl('SCODE'));
  SLIBELLE := TCheckBox(GetControl('SLIBELLE'));
  SQTE  := TCheckBox(GetControl('SQTE'));
  SUNITE := TCheckBox(GetControl('SUNITE'));
  VALPA  := TCheckBox(GetControl('VALPA'));
  VALPR  := TCheckBox(GetControl('VALPR'));
  VALPV  := TCheckBox(GetControl('VALPV'));
end ;

procedure TOF_BTOUVRAGE_PEDT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_PEDT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_PEDT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTOUVRAGE_PEDT.ClickSousDet(Sender: TObject);
begin
	if TCheckBox(Sender).Checked then
  begin
    GSOUSDET.Enabled := True;
  end else
  begin
    SCODE.Checked := false;
    SLIBELLE.Checked := false;
    SQTE.Checked := false;
    SUNITE.Checked := false;
    GSOUSDET.Enabled := false;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTOUVRAGE_PEDT ] ) ; 
end.

