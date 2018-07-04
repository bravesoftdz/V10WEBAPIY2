{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 25/04/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : HIERARCHIE (HIERARCHIE)
Mots clefs ... : TOM;HIERARCHIE
*****************************************************************
PT1 21/06/2002 V582 SB En création test si niveau existe déjà}
Unit UTomHierarchie ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,FichList,
{$ELSE}
     eFiche,eFichList,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob ;

Type
  TOM_HIERARCHIE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    end ;

Implementation

procedure TOM_HIERARCHIE.OnNewRecord ;
var Q:TQuery;
    NiveauMax:Integer;
begin
  Inherited ;
Q:=OpenSQL('SELECT MAX(PHO_NIVEAUH) AS MAXIMUM FROM HIERARCHIE',True);
If Q<>Nil then  NiveauMax:=Q.FindField('MAXIMUM').AsInteger
Else NiveauMax:=0;
Ferme(Q);
NiveauMax:=NiveauMax+1;
If NiveauMax>20 Then NiveauMax:=0;
SetField('PHO_NIVEAUH',NiveauMax);
end ;

procedure TOM_HIERARCHIE.OnUpdateRecord ;
Var St : String;
begin
  Inherited ;
if DS.State in [dsInsert] then //PT1 SB 21/06/2002 En création test si niveau existe déjà
  Begin
  St:='SELECT PHO_HIERARCHIE FROM HIERARCHIE '+
  'WHERE PHO_NIVEAUH="'+IntToStr(GetField('PHO_NIVEAUH'))+'" '+
  'AND PHO_HIERARCHIE<>"'+GetField('PHO_HIERARCHIE')+'"';   //PT1 SB 21/06/2002 Exclure la hiérarchie en cours de création
  If ExisteSQL(St) Then
    begin
    LastError:=1;
    PgiBox('Le niveau '+IntToStr(GetField('PHO_NIVEAUH'))+' existe déjà, veuillez changer le niveau','Saisie d''un niveau hiérarchique');
    SetFocusControl('PHO_NIVEAUH');
    end;
  End;
end ;

procedure TOM_HIERARCHIE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_HIERARCHIE.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_HIERARCHIE ] ) ;
end.
