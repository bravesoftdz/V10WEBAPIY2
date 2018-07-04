
{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 01/06/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SALAIREMOYEN (SALAIREMOYEN)
Mots clefs ... : TOM;SALAIREMOYEN
*****************************************************************
pt1    :  19/04/2007   NA   Affichage des salaires de la période et salaire des 3 derniers mois}

Unit SALAIREMOYEN_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     {$else}
     eFiche, 
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls, 
     P5Def,
     UTOM,
     UTob ;

Type
  TOM_SALAIREMOYEN = Class (TOM)
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
    end ;

Implementation

procedure TOM_SALAIREMOYEN.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_SALAIREMOYEN.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_SALAIREMOYEN.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_SALAIREMOYEN.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_SALAIREMOYEN.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/09/2006
Modifié le ... :   /  /    
Description .. : Chargement des enregistrements
Mots clefs ... : 
*****************************************************************}
procedure TOM_SALAIREMOYEN.OnLoadRecord ;
begin
  Inherited ;
  TFFiche(Ecran).Caption := 'Salarié :' + ' ' + Getfield('PSY_SALARIE') + ' ' + Getfield('PSY_NOM') + ' ' + Getfield('PSY_PRENOM');
  UpdateCaption(Ecran);
  // pt1
  if (getfield('PSY_SALAIREMOYPER') = 0) and (getfield('PSY_SALAIREMOYDERM') = 0) then
  begin
    setcontrolvisible('PSY_SALAIREMOYPER', false);
    setcontrolvisible('PSY_SALAIREMOYDERM', false);
    setcontrolvisible('TPSY_SALAIREMOYPER', false);
    setcontrolvisible('TPSY_SALAIREMOYDERM', false);
  end
  else
  begin
    setcontrolvisible('PSY_SALAIREMOYPER', true);
    setcontrolvisible('PSY_SALAIREMOYDERM', true);
    setcontrolvisible('TPSY_SALAIREMOYPER', true);
    setcontrolvisible('TPSY_SALAIREMOYDERM', true);
  end;
  // fin pt1
end ;

procedure TOM_SALAIREMOYEN.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 01/06/2006
Modifié le ... :   /  /    
Description .. : On Argument
Mots clefs ... : 
*****************************************************************}
procedure TOM_SALAIREMOYEN.OnArgument ( S: String ) ;
var
num : integer;
begin
  Inherited ;
  // Recherche des libellés des zones TRAVAIL, CODESTAT, LIBREPCMB
  for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PSY_TRAVAILN' + IntToStr(Num)), GetControl('TPSY_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSY_CODESTAT'), GetControl('TPSY_CODESTAT'));
   
  for Num := 1 to 4 do
  begin
  VisibiliteChamplibresal(Inttostr(Num), GetControl('PSY_LIBREPCMB'+ Inttostr(Num)),GetControl('TPSY_LIBREPCMB'+ IntToStr(Num)));
  end;

end ;

procedure TOM_SALAIREMOYEN.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_SALAIREMOYEN.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_SALAIREMOYEN ] ) ;
end.
