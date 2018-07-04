{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/06/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYYRESSOURCE_MUL ()
Mots clefs ... : TOF;YYYRESSOURCE_MUL
*****************************************************************}
Unit YYYRESSOURCE_MUL_TOF ;

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
     UTOF,
     AglInit;

Type
  TOF_YYYRESSOURCE_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    Private
      BRecalculer:    TToolBarButton97;
      Procedure BRecalculerDoClick ( Sender: TObject );

  End ;

Implementation
Uses
  YRessource;       //Pour la fonction RecalculYRS;

procedure TOF_YYYRESSOURCE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYYRESSOURCE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYYRESSOURCE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYYRESSOURCE_MUL.OnLoad ;
begin
  Inherited ;
end ;

Procedure TOF_YYYRESSOURCE_MUL.OnArgument( S : String );
Begin
  Inherited ;
  BRecalculer := TToolBarButton97( GetControl( 'BRECALCULER' ) );
  BRecalculer.OnClick := BRecalculerDoClick;
End ;

Procedure TOF_YYYRESSOURCE_MUL.BRecalculerDoClick ( Sender: TObject );
Begin
  If PGIAsk( 'Etes-vous sûr de vouloir recalculer les ressources unifiées ?', Ecran.Caption ) <> mrYes Then Exit;
  SourisSablier;
  RecalculYRS;
  SourisNormale;
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
End;

procedure TOF_YYYRESSOURCE_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_YYYRESSOURCE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYYRESSOURCE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_YYYRESSOURCE_MUL ] ) ; 
end.
