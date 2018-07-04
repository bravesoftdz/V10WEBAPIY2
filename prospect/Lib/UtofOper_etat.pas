{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OPERETAT ()
Mots clefs ... : TOF;OPERETAT
*****************************************************************}
Unit UtofOper_etat ;

Interface

Uses Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     FE_Main,
{$ENDIF}
     forms, 
     UTOF,
{$ifdef AFFAIRE}
    UTOFAFTRADUCCHAMPLIBRE,
{$endif}
     ParamSoc,
     UtilSelection ; 

Type
{$ifdef AFFAIRE}
                //PL le 18/05/07 pour gérer les champs libres si paramétrés
     TOF_OPERETAT = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_OPERETAT = Class (TOF)
{$endif}
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Procedure RTLanceFiche_OperEtat (Nat,Cod : String ; Range,Lequel,Argument : string) ;

Implementation

Procedure RTLanceFiche_OperEtat(Nat,Cod : String ; Range,Lequel,Argument : string) ;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_OPERETAT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_OPERETAT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_OPERETAT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_OPERETAT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_OPERETAT.OnArgument (S : String ) ;
var F : TForm;
begin
inherited ;
F := TForm (Ecran);

if GetParamSocSecur('SO_RTGESTINFOS002',False) = True then
    MulCreerPagesCL(F,'NOMFIC=RTOPERATIONS');
end;

procedure TOF_OPERETAT.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_OPERETAT ] ) ; 
end.

 
