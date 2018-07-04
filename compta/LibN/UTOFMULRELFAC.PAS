{***********UNITE*************************************************
Auteur  ...... : BOUSSERT Stéphane
Créé le ...... : 29/04/2003
Modifié le ... : 29/04/2003
Description .. : Source TOF de la FICHE : CPMULRELFAC ()
Mots clefs ... : TOF;CPMULRELFAC
*****************************************************************}
Unit UTOFMULRELFAC ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     CPTIERS_TOM,        // pour Zoom sur fiche tiers
     UTofMulRelDetail,   // pour GenereReleves
{$IFDEF EAGLCLIENT}
  UTob ,
  maineAGL,
  UtileAGL,
  eMul,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  fe_main,
  EdtRDoc,
  Mul,
  hdb,
  saisutil,    // pour EditDocumentS5S7
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HTB97,
     HEnt1,
     Ent1,
     Ed_Tools,
     HMsgBox,
     HQry,
     LettUtil,        // pour InitTablesLibresTiers
     uRecupSQLModele, // pour RecupSQLModele
     ParamSoc,		// GetParamSocSecur YMO
     UTOF ;


Procedure ReleveFacture ;
Procedure CPLanceFiche_MulRelFac ( vStRange, vStLequel, vStArgs : String ) ;


Type
  TOF_CPMULRELFACT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  Private
  	Q                   : THQuery ;
{$IFDEF EAGLCLIENT}
    FListe              : THGrid ;
{$ELSE}
    FListe              : THDBGrid ;
{$ENDIF}
    // Initialisations
    procedure InitComposants ;
    procedure InitEvenements ;
    // Boutons
    procedure FListeDblClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BParamModeleClick(Sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  end ;

Implementation

uses edtdoc // EditDocument
     ,UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

// =============================================================================

Procedure ReleveFacture ;
Begin
  CPLanceFiche_MulRelFac ('', '', '') ;
End ;

// =============================================================================

Procedure CPLanceFiche_MulRelFac ( vStRange, vStLequel, vStArgs : String ) ;
begin
  AGLLanceFiche('CP', 'CPMULRELFACT', vStRange, vStLequel, vStArgs ) ;
end ;

// =============================================================================

procedure TOF_CPMULRELFACT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.OnArgument (S : String ) ;
begin
  Inherited ;

  InitComposants ;
  InitEvenements ;

  InitTablesLibresTiers( TTabSheet(GetControl('PComplement', True )) ) ;
  SetControlProperty('T_SECTEUR','ItemIndex',0) ;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('REL_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('REL_AUXILIAIRE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

  {JP 30/10/07 : FQ 21432 : Glyph du bouton valider "VVert"}
  TFMul(Ecran).BOuvrir.GlobalIndexImage := 'Z0003_S16G2';
end ;

procedure TOF_CPMULRELFACT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULRELFACT.BParamModeleClick(Sender: TObject);
begin
{ FQ 19667 BVE 23.04.07 }
  EditDocument('L','RLV','RLW',True) ;
(*
{$IFDEF EAGLCLIENT}
  // Non compatible eAGL
{$ELSE}
  EditDocumentS5S7('L','RLV','RLV',True) ;
{$ENDIF}
*)
{ END FQ 19667 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULRELFACT.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPMULRELFACT.BZoomClick(Sender: TObject);
Var Cpte : String ;
begin
  inherited;
{$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then Exit ;
  Q.TQ.Seek( FListe.row - 1) ;
  Cpte := Q.FindField('REL_AUXILIAIRE').AsString ;
  FicheTiers(nil, '', Cpte, taConsult, 1) ;
{$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  Cpte := Q.FindField('REL_AUXILIAIRE').AsString ;
  FicheTiers(Q,'',Cpte,taConsult,1) ;
{$ENDIF}

end;

procedure TOF_CPMULRELFACT.FListeDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  if (Q.TQ.Detail.Count = 0) then Exit ;
  GenereReleves( Q.TQ , FListe.row - 1 ) ;
{$ELSE}
  if ((Q.EOF) and (Q.BOF)) then Exit ;
  GenereReleves( Q ) ;
{$ENDIF}

end;

procedure TOF_CPMULRELFACT.InitComposants;
begin
 	Q      := THQuery(GetControl('Q', True)) ;
{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FListe',True)) ;
{$ELSE}
  FListe := THDBGrid(GetControl('FListe',True)) ;
{$ENDIF}
  // Boutons
{$IFDEF EAGLCLIENT}   
{ FQ 19667 BVE 23.04.07
  TButton(GetControl('BParamModele',True)).Visible := False ;
  END FQ 19667 }
{$ENDIF}
end;

procedure TOF_CPMULRELFACT.InitEvenements;
begin
  // Grille
{$IFDEF EAGLCLIENT}
  FListe.OnDblClick := FListeDblClick ;
{$ELSE}
  FListe.OnDblClick := FListeDblClick ;
{$ENDIF}
  // Boutons
  TButton(GetControl('BParamModele',True)).OnClick := BParamModeleClick ;
  // menuZoom
  TButton(GetControl('BZOOM',True)).OnClick        := BZoomClick ;
end;


Initialization
  registerclasses ( [ TOF_CPMULRELFACT ] ) ;
end.

