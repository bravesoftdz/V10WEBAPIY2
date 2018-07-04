{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPREVHISTOCYCLE ()
Mots clefs ... : TOF;CPREVHISTOCYCLE
*****************************************************************}
Unit CPREVHISTOCYCLE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl,      // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_main,       // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     UTOF ;

Type
  TOF_CPREVHISTOCYCLE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    FStArgument : string;

  public
    FListe : THGrid;
  end ;

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_CPREVHISTOCYCLE( vStCycle, vStExercice : string ) : string;
////////////////////////////////////////////////////////////////////////////////

Implementation

uses  HSysMenu, // THSystemMenu
      uTob;    // TOB

const
      cColModifPar  = 1;
      cColDateModif = 2;
      cColEtatCycle = 3;
      cColEtatPar   = 4;
      cColEtatLe    = 5;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_CPREVHISTOCYCLE( vStCycle, vStExercice : string ) : string;
begin
  Result := AGLLanceFiche('CP', 'CPREVHISTOCYCLE', '', '', vStCycle + ';' + vStExercice );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVHISTOCYCLE.OnArgument (S : String ) ;
begin
  Inherited ;

  FStArgument := S;

  FListe := THGrid(GetControl('FLISTE', True));

  FListe.ColALigns[cColModifPar]  := TaCenter;
  FListe.ColALigns[cColDateModif] := TaCenter;
  FListe.ColALigns[cColEtatCycle] := TaCenter;
  FListe.ColALigns[cColEtatPar]   := TaCenter;
  FListe.ColALigns[cColEtatLe]    := TaCenter;

  FListe.ColFormats[cColModifPar]  := 'CB=TTUTILISATEUR||';
  FListe.ColFormats[cCOlEtatCycle] := 'CB=CREVETATCYCLE||';
  FListe.ColFormats[cCOlEtatPar]   := 'CB=TTUTILISATEUR||';
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVHISTOCYCLE.OnLoad ;
var lStCycle    : string;
    lStLibelle  : string;
    lStExercice : string;
    lTob : Tob;

begin
  Inherited ;

  lStCycle    := ReadTokenSt( FStArgument );
  lStExercice := ReadTokenSt( FStArgument );
  lStLibelle  := GetColonneSQL('CREVCYCLE', 'CCY_LIBELLECYCLE', 'CCY_CODECYCLE = "' + lStCycle + '"');

  lTob := Tob.Create('HISTO', nil, -1 );
  try
    lTob.LoadDetailFromSQL('SELECT CHC_CREATEUR,CHC_DATECREATION, CHC_ETATCYCLE, ' +
                           'CHC_ETATCYCLEPAR, CHC_ETATCYCLELE ' +
                           'FROM CREVHISTOCYCLE WHERE ' +
                           'CHC_CODECYCLE = "' + lStCycle + '" AND ' +
                           'CHC_EXERCICE = "' + lStExercice + '" ' +
                           'ORDER BY CHC_DATECREATION DESC', False);

    if lTob.Detail.Count > 0 then
      lTob.PutGridDetail(FListe, False, False, '');
  finally
    FreeAndNil( lTob );
    THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
    Ecran.Caption := Ecran.Caption + lStcycle + ' - ' + lStLibelle;
    UpdateCaption( Ecran ); 
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVHISTOCYCLE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPREVHISTOCYCLE.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_CPREVHISTOCYCLE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPREVHISTOCYCLE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPREVHISTOCYCLE ] ) ; 
end.
