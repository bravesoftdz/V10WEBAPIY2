{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/05/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : EVOLUTIONSAL 
Suite ........ : (EVOLUTIONSAL)
Mots clefs ... : TOM;EVOLUTIONSAL;PAIE;PGMODIFCOEFFSAL
*****************************************************************}
Unit UTOMEVOLUTIONSAL ;

Interface

Uses Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
{$else}
     UTob,
{$ENDIF}
     UTOM;

Type
  TOM_EVOLUTIONSAL = Class (TOM)
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

procedure TOM_EVOLUTIONSAL.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_EVOLUTIONSAL.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_EVOLUTIONSAL.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_EVOLUTIONSAL.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_EVOLUTIONSAL.OnLoadRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/05/2004
Modifié le ... :   /  /    
Description .. : OnChangeFiled
Mots clefs ... : PAIE;PGMODIFCOEFFSAL
*****************************************************************}
procedure TOM_EVOLUTIONSAL.OnChangeField ( F: TField ) ;
begin
Inherited ;

if ((F.FieldName <> 'PVL_ANCIENPOSTEP') and
   (GetField('PVL_ANCIENPOSTEP') <> 'X')) then
   SetField('PVL_ANCIENPOSTEP', 'X');
   
if (F.FieldName = 'PVL_EVOLUANT') then
   begin
   if (GetField('PVL_EVOLUANT')='001') then
      begin
      SetField('PVL_COEFFICIENTP', '-');
      SetControlEnabled('PVL_COEFFICIENTP', False);
      SetControlEnabled('PVL_QUALIFP', True);
      SetControlEnabled('PVL_NIVEAUP', True);
      SetControlEnabled('PVL_INDICEP', True);
      end
   else
   if (GetField('PVL_EVOLUANT')='002') then
      begin
      SetField('PVL_QUALIFP', '-');
      SetControlEnabled('PVL_COEFFICIENTP', True);
      SetControlEnabled('PVL_QUALIFP', False);
      SetControlEnabled('PVL_NIVEAUP', True);
      SetControlEnabled('PVL_INDICEP', True);
      end
   else
   if (GetField('PVL_EVOLUANT')='003') then
      begin
      SetField('PVL_NIVEAUP', '-');
      SetControlEnabled('PVL_COEFFICIENTP', True);
      SetControlEnabled('PVL_QUALIFP', True);
      SetControlEnabled('PVL_NIVEAUP', False);
      SetControlEnabled('PVL_INDICEP', True);
      end
   else
   if (GetField('PVL_EVOLUANT')='004') then
      begin
      SetField('PVL_INDICEP', '-');
      SetControlEnabled('PVL_COEFFICIENTP', True);
      SetControlEnabled('PVL_QUALIFP', True);
      SetControlEnabled('PVL_NIVEAUP', True);
      SetControlEnabled('PVL_INDICEP', False);
      end;
   end;

if (ds.State in [dsEdit]) then   
   SetControlEnabled('bDefaire', True);
end ;

procedure TOM_EVOLUTIONSAL.OnArgument ( S: String ) ;
begin
  Inherited ;
SetControlEnabled('PVL_ANCIENPOSTEP', False);
SetControlEnabled('bDefaire', False);
end ;

procedure TOM_EVOLUTIONSAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_EVOLUTIONSAL.OnCancelRecord ;
begin
  Inherited ;
if (GetField('PVL_EVOLUANT')='001') then
   begin
   SetField('PVL_COEFFICIENTP', '-');
   SetControlEnabled('PVL_COEFFICIENTP', False);
   SetControlEnabled('PVL_QUALIFP', True);
   SetControlEnabled('PVL_NIVEAUP', True);
   SetControlEnabled('PVL_INDICEP', True);
   end
else
if (GetField('PVL_EVOLUANT')='002') then
   begin
   SetField('PVL_QUALIFP', '-');
   SetControlEnabled('PVL_COEFFICIENTP', True);
   SetControlEnabled('PVL_QUALIFP', False);
   SetControlEnabled('PVL_NIVEAUP', True);
   SetControlEnabled('PVL_INDICEP', True);
   end
else
if (GetField('PVL_EVOLUANT')='003') then
   begin
   SetField('PVL_NIVEAUP', '-');
   SetControlEnabled('PVL_COEFFICIENTP', True);
   SetControlEnabled('PVL_QUALIFP', True);
   SetControlEnabled('PVL_NIVEAUP', False);
   SetControlEnabled('PVL_INDICEP', True);
   end
else
if (GetField('PVL_EVOLUANT')='004') then
   begin
   SetField('PVL_INDICEP', '-');
   SetControlEnabled('PVL_COEFFICIENTP', True);
   SetControlEnabled('PVL_QUALIFP', True);
   SetControlEnabled('PVL_NIVEAUP', True);
   SetControlEnabled('PVL_INDICEP', False);
   end;

SetControlEnabled('bDefaire', False);
end ;

Initialization
  registerclasses ( [ TOM_EVOLUTIONSAL ] ) ;
end.
