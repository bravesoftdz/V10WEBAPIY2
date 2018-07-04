{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/02/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARAMEXTPAIE (PARAMEXTPAIE)
Mots clefs ... : TOM;PARAMEXTPAIE
*****************************************************************}
{
PT1 14/12/2007 GGU V_8 FQ14065 paramétrage des extractions : pouvoir choisir le TAUX d'une rémunération
}
Unit PARAMEXTPAIE_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
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
  TOM_PARAMEXTPAIE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

procedure TOM_PARAMEXTPAIE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnChangeField ( F: TField ) ;
var
  num            : Integer ;
  st, LaTablette,Nom : String ;
  ok : Boolean;
begin
  Inherited ;
  Nom :=  F.FieldName ;
  Nom :=  Copy (Nom, 1,  10);
  if Nom = 'PPE_ETALIM' then
  begin
    num := StrToInt (Copy (F.FieldName, 11,  1)) ;
    St := GetControlText (F.FieldName);
    if St = '001' then LaTablette := 'PGCUMULPAIE'
    else
    if (St = '005') OR (St = '006') OR (St = '310') then  LaTablette := 'PGREMUNERATION'  // PT1
    else LaTablette := 'PGCOTISATION';
    if st <> '' then SetControlProperty ('PPE_ETCOL'+IntToStr (Num), 'Datatype', LaTablette);
  end;
  if F.FieldName = 'PPE_PGPRESENTATION' then
  begin
    st := GetControlText ('PPE_PGPRESENTATION');
    if st = 'MEN' then ok := FALSE
    else Ok := TRUE ;
    SetControlVisible ('GRPBX2', ok);
    SetControlVisible ('GRPBX3', ok);
    SetControlVisible ('TBSHTPAR2', ok);
    SetControlVisible ('TBSHTPAR3', ok);
    SetControlVisible ('TBSHTPAR4', ok);
  end;
end ;

procedure TOM_PARAMEXTPAIE.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PARAMEXTPAIE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_PARAMEXTPAIE ] ) ; 
end.
