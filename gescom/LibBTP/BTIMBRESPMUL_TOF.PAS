{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/08/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTIMBRESPMUL ()
Mots clefs ... : TOF;BTIMBRESPMUL
*****************************************************************}
Unit BTIMBRESPMUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
     DBGrids,
{$else}
     eMul,
     Main_eagl,
{$ENDIF}
		 HTB97,
     HDB,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTIMBRESPMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	procedure BnewClick (Sender : TObject);
  	procedure fListeDblClick (Sender : TObject);
  end ;

Implementation

procedure TOF_BTIMBRESPMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolbarButton97 (GetControl('Binsert')).OnClick := BnewClick;
  THDbgrid (GetControl('Fliste')).OnDblClick  := fListeDblClick;
end ;

procedure TOF_BTIMBRESPMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMBRESPMUL.BnewClick(Sender: TObject);
begin
	AGLLanceFiche('BTP','BTIMBREPARAM','','','ACTION=CREATION');
  TToolbarButton97(GetControl('BCherche')).Click;
end;

procedure TOF_BTIMBRESPMUL.fListeDblClick(Sender: TObject);
var Code : string;
begin
  Code:=THDbgrid (GetControl('Fliste')).datasource.dataset.FindField('BTP_CODE').AsString;
  AGLLanceFiche('BTP','BTIMBREPARAM','',Code,'ACTION=MODIFICATION');
  TToolbarButton97(GetControl('BCherche')).Click;
end;

Initialization
  registerclasses ( [ TOF_BTIMBRESPMUL ] ) ;
end.
