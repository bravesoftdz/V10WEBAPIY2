{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/03/2015
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BTCODEBARRE (BTCODEBARRE)
Mots clefs ... : TOM;BTCODEBARRE
*****************************************************************}
Unit UtomCodeBarrre ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob ;

Type
  TOM_BTCODEBARRE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
    //
  Private
    //Définition des variables utilisées dans le Uses
    Action        : TActionFiche;
    TypeAction    : String; //C=Création - M=Modification - S=Suppression
    //
    NatureCAB     : THEdit;
    IdentifCAB    : THEdit;
    CodeBarre     : THEdit;
    //
    CabPrincipal  : THCheckbox;
    QualifCAB     : THValComboBox;
    //
    procedure Controlechamp(Champ, Valeur: String);
    procedure GetObjects;
    procedure SetScreenEvents;

    end ;

Implementation

procedure TOM_BTCODEBARRE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnArgument ( S: String ) ;
var i       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
    Indice  : Integer;
begin

  Inherited ;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  Critere := S;

  While (Critere <> '') do
  BEGIN
    i:=pos(':',Critere);
    if i = 0 then i:=pos('=',Critere);
    if i <> 0 then
       begin
       Champ:=copy(Critere,1,i-1);
       Valeur:=Copy (Critere,i+1,length(Critere)-i);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  //Gestion des évènement des zones écran
  SetScreenEvents;

end ;

procedure TOM_BTCODEBARRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.OnAfterRollBack;
begin
  Inherited ;
end ;

procedure TOM_BTCODEBARRE.GetObjects;
begin

  NatureCAB     := THEdit(GetControl('BCB_NATURECAB'));
  IdentifCAB    := THEdit(GetControl('BCB_NATURECAB'));
  CodeBarre     := THEdit(GetControl('BCB_NATURECAB'));
  //
  CabPrincipal  := THCheckbox(GetControl('BCB_CABPRINCIPAL'));
  //
  QualifCAB     := THValComboBox(GetControl('BCB_QUALIFCODEBARRE'));
  //
end;

Procedure TOM_BTCODEBARRE.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if Valeur='CREATION'          then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

end;

procedure TOM_BTCODEBARRE.SetScreenEvents;
begin

end;

Initialization
  registerclasses ( [ TOM_BTCODEBARRE ] ) ; 
end.

