{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTRECUPDOSS1 ()
Mots clefs ... : TOF;BTRECUPDOSS1
*****************************************************************}
Unit BTRECUPDOSS1_TOF ;

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
     HTB97,
     UtilActionComSx,
     UTOF ;

Type
  TOF_BTRECUPDOSS1 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	REPERDESTINATION : THCritMaskEdit;
  	UnEchangeComSx : TEchangeComSx;
  	BVAlide : TToolbarButton97;
    procedure BValideClick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTRECUPDOSS1.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPDOSS1.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPDOSS1.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPDOSS1.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPDOSS1.OnArgument (S : String ) ;
begin
  Inherited ;
  BValide := TToolBarButton97(GetControl('BVALIDE'));
  BValide.onClick := BValideClick;
  REPERDESTINATION := THCritMaskEdit(GetControl('REPERTDESTINATION'));
  UnEchangeComSx := TEchangeComSx.create;
end ;

procedure TOF_BTRECUPDOSS1.OnClose ;
begin
  Inherited ;
  UnEchangeComSx.free;
end ;

procedure TOF_BTRECUPDOSS1.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPDOSS1.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTRECUPDOSS1.BValideClick(Sender: TObject);
begin
  if REPERDESTINATION.Text = '' then
  begin
  	PgiBox(TraduireMemoire('Vous devez sélectionner un fichier '),Ecran.caption);
  	TForm(Ecran).ModalResult:=0;
    exit;
  end;
  If PgiAsk(TraduireMemoire ('Validez-vous le traitement d''intégration ?'),ecran.caption) = MrYes then
  begin
  	UnEchangeComSx.Init;
  	UnEchangeComSx.ExecuteAction (TmaSxImport,REPERDESTINATION.Text);
    PgiBox('Traitement Terminé',ecran.caption);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTRECUPDOSS1 ] ) ; 
end.

