{***********UNITE*************************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 07/12/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TYPEFLUX (TYPEFLUX)
Mots clefs ... : TOM;TYPEFLUX
*****************************************************************}
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
                 07/12/01   BT    Création de l'unité
1.0.1.001.001    04/03/04   JP    Mise en place d'un contrôle dans le OnUpdateRecord et
                                  sur le OnDeleteRecord sur le TypeFlux pour interdire
                                  les modifications sur les 3 types EQD, EQR, REI fournis
                                  dans la "maquette"
--------------------------------------------------------------------------------------}
unit TomTypeFlux ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL, UTob,
  {$ELSE}
  FE_Main, db,
  {$ENDIF}
  Controls, Classes, Forms, HEnt1, UTOM;

type
  TOM_TYPEFLUX = class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  protected
    procedure CodeTypeFluxOnKeyPress(Sender: TObject; var Key: Char);
  end ;

procedure TRLanceFiche_TypeFlux(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

Implementation

uses
  Commun, HDB, Constantes, HMsgBox;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_TypeFlux(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TYPEFLUX.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  Ecran.HelpContext := 150;
  THDBEdit(GetControl('TTL_TYPEFLUX')).OnKeyPress := CodeTypeFluxOnKeyPress;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_TYPEFLUX.CodeTypeFluxOnKeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  Inherited;
  ValidCodeOnKeyPress(Key);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TYPEFLUX.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  {JP 04/03/04 : On interdit la modification des enregistrements fournis dans la maquette}
  if (GetField('TTL_TYPEFLUX') = CODEREGULARIS) or
     (GetField('TTL_TYPEFLUX') = CODETRANSACDEP) or
     (GetField('TTL_TYPEFLUX') = CODETRANSACREC) then begin
    HShowMessage('0;' + Ecran.Caption + ';Vous n''avez pas le droit de modifier ce type de flux;W;O;O;O', '', '');
    LastError := 1;
  end
  {03/06/04 : on force la saisie d'une sens}
  else if VarToStr(GetField('TTL_SENS')) = '' then begin
    LastErrorMsg := TraduireMemoire('Vous devez saisir un sens');
    LastError := 1;
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOM_TYPEFLUX.OnDeleteRecord ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  {JP 04/03/04 : On interdit la modification des enregistrements fournis dans la maquette}
  if (GetField('TTL_TYPEFLUX') = CODEREGULARIS) or
     (GetField('TTL_TYPEFLUX') = 'ER') or
     (GetField('TTL_TYPEFLUX') = 'ED') then begin
    HShowMessage('0;' + Ecran.Caption + ';Vous n''avez pas le droit de supprimer ce type de flux;W;O;O;O', '', '');
    LastError := 1;
  end;
end ;

procedure TOM_TYPEFLUX.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPEFLUX.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_TYPEFLUX.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPEFLUX.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_TYPEFLUX.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_TYPEFLUX.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_TYPEFLUX ] ) ;
end.

