{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TarifsSurQuoi ()
Mots clefs ... : TOF;TarifsSurQuoi
*****************************************************************}
Unit TarifsSurQuoi_TOF ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    dbtables,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  wTOF,
  SaisieList,
  uTableFiltre
  ;

Type
  TOF_TarifsSurQuoi = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
      TF : TTAbleFiltre;
  end ;

Implementation

uses
  wCommuns;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsSurQuoi.OnNew ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsSurQuoi.OnDelete ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsSurQuoi.OnUpdate ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsSurQuoi.OnLoad ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsSurQuoi.OnArgument (S : String ) ;
var
  sAppel, sFonctionnalite : string;
begin
  Inherited ;
  sAppel          := GetArgumentValue(S, 'APPEL');
  sFonctionnalite := GetArgumentValue(S, 'YTA_FONCTIONNALITE');

  if (Ecran<>nil) then
  begin
    TF := TFSaisieList(Ecran).LeFiltre;
    if (sAppel='ARTICLE') then
    begin
      TF.WhereTable := TF.WhereTable + ' and ((YTQ_ARTICLE="'+GetArgumentValue(S,'YTQ_ARTICLE')+'") or (YTQ_TARIFARTICLE="'+GetArgumentValue(S,'YTQ_TARIFARTICLE')+'"))';
    end;

    //Fixer la liste par rapport à la fonctionnalité
    TFSaisieList(ecran).Dbliste               := 'YTARIFSURQUOI'+sFonctionnalite;
    TFsaisieList(ecran).Lefiltre.LaGridListe  := 'YTARIFSURQUOI'+sFonctionnalite;
  end;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsSurQuoi.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_TarifsSurQuoi ] ) ;
end.
