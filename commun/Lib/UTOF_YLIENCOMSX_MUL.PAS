{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 18/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YLIENCOMSX_MUL ()
Mots clefs ... : TOF;YLIENCOMSX_MUL
*****************************************************************}
Unit UTOF_YLIENCOMSX_MUL;

Interface

Uses
     Classes,
{$IFNDEF EAGLCLIENT}
     fe_main,
{$else}
     maineagl,
{$ENDIF}
      UTOF
     ;

Type
  TOF_YLIENCOMSX_MUL = Class( TOF)

  private
    procedure BInsertClick(Sender : TObject);
    procedure BOuvrirClick(Sender : TObject);
    Procedure BJournal_OnClick(Sender: TObject);
  public
    procedure OnArgument (S : String ) ; override ;
  end ;

function LanceFicheYLienComsx_Mul( Range, Lequel, Argument : string ) : string;

Implementation

uses
  ParamSoc,
  HTB97,
  UTOF_YLIENCOMSX,
  AGLInit ;


procedure TOF_YLIENCOMSX_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
  TToolBarButton97(GetControl('BInsert')).OnClick := BInsertClick ;
  TToolBarButton97(GetControl('BOuvrir')).OnClick := BOuvrirClick ;
  TToolBarButton97(GetControl('BJOurnal')).OnClick := BJournal_OnClick ;
  //attente process serveur.. zones cachées dans un 1er temps
  SetcontrolVisible ('LPlanification',false);
  SetcontrolVisible ('YLO_NUMJOB',false);
  SetcontrolVisible ('TYLO_NUMJOB',false);
end ;

Procedure TOF_YLIENCOMSX_MUL.BJournal_OnClick(Sender: TObject);
begin
  // Affichage des rapports de génération du journal d'évènements
  AGLLanceFiche('YY','YYJNALEVENT','GEV_TYPEEVENT=COM','','TYPEEVEN=COM');
end;

procedure TOF_YLIENCOMSX_MUL.BInsertClick(Sender : TObject);
begin
  LanceYLienComsx( '','CREATION')  ;
  AGLRefreshDB( [ LongInt( Ecran ), 'FListe' ], 2 );
end ;

procedure TOF_YLIENCOMSX_MUL.BOuvrirClick(Sender : TObject);
begin
  LanceYLienComsx( GetField('YLO_CODECPTA'), 'MODIFICATION')  ;
  AGLRefreshDB( [ LongInt( Ecran ), 'FListe' ], 2 );
end ;


//*****************************************************************
// ACTIONS FICHES
//*****************************************************************

function LanceFicheYLienComsx_Mul( Range, Lequel, Argument : string ) : string;
begin
  Result := AGLLanceFiche( 'YY', 'YLIENCOMSX_MUL', Range, Lequel, Argument );
end;

Initialization
  registerclasses ( [ TOF_YLIENCOMSX_MUL ] ) ;
end.
