{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPREVMULPARAMCYCLE ()
Mots clefs ... : TOF;CPREVMULPARAMCYCLE
*****************************************************************}
Unit CPREVMULPARAMCYCL_TOF ;

Interface

Uses StdCtrls,
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl,      // AGLLanceFiche
{$ELSE}
     Hdb,
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_main,       // AGLLanceFiche
{$ENDIF}
     Htb97,         // TToolBarButton97
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_CPREVMULPARAMCYCL( const vStArg : string ) : string;
////////////////////////////////////////////////////////////////////////////////

Type
  TOF_CPREVMULPARAMCYCL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

  public

  {$IFDEF EAGLCLIENT}
    FListe : THGrid;
  {$ELSE}
    FListe : THDBGrid;
  {$ENDIF}

    procedure OnDblClickFListe     ( Sender : TObject );
    procedure OnClickBInsert       ( Sender : TObject );

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ELSE}
  Ent1,               // EstSpecif
  {$ENDIF MODENT1}
  CRevParamCycle_Tom; // CPLanceFiche_CRevParamCycle

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_CPREVMULPARAMCYCL( const vStArg : string ) : string;
begin
  Result := AGLLanceFiche('CP', 'CPREVMULPARAMCYCL', '', '', vStArg );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVMULPARAMCYCL.OnArgument (S : String ) ;
begin
  Inherited ;

{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE', True));
{$ELSE}
  FListe := THDBGrid(GetControl('FLISTE', True));;
{$ENDIF}
  FListe.OnDblClick := OnDblClickFListe;

  TToolBarButton97(GetControl('BINSERT')).OnClick := OnClickBInsert;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVMULPARAMCYCL.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVMULPARAMCYCL.OnNew ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVMULPARAMCYCL.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_CPREVMULPARAMCYCL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPREVMULPARAMCYCL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPREVMULPARAMCYCL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPREVMULPARAMCYCL.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVMULPARAMCYCL.OnDblClickFListe(Sender: TObject);
begin
  if (GetField('CPC_PREDEFINI') = 'CEG') and (not EstSpecif('51502')) then
    CPLanceFiche_CRevParamCycle( GetField('CPC_CODECYCLE') + ';' +  GetField('CPC_PREDEFINI'), TaConsult )
  else
    CPLanceFiche_CRevParamCycle( GetField('CPC_CODECYCLE') + ';' +  GetField('CPC_PREDEFINI'), TaModif );
  TToolBarButton97(GetControl('BCHERCHE')).Click; 
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVMULPARAMCYCL.OnClickBInsert(Sender: TObject);
begin
  CPLanceFiche_CRevParamCycle( '', TaCreat );
  TToolBarButton97(GetControl('BCHERCHE')).Click;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPREVMULPARAMCYCL ] ) ; 
end.
