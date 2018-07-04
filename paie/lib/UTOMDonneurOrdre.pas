{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 31/07/2001
Modifié le ... :   /  /    
Description .. : Gestion des donneurs d'ordre
Mots clefs ... : PAIE;DONNEURORDRE
*****************************************************************}
{
PT1    : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
PT2    : 14/09/2001 SB V547 Champ PDO_PGMODEREGLE DataType non renseigné
PT3    : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
}
unit UTOMDonneurOrdre;

interface
uses  StdCtrls,Classes,
{$IFNDEF EAGLCLIENT}
      db,
{$ELSE}
     UTOB,
{$ENDIF}
      UTOM,
      HCtrls,
      Commun;

Type
      TOM_DonneurOrdre = Class (TOM)
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoadRecord ; override ;
     END ;
implementation

{ TOM_DonneurOrdre }

procedure TOM_DonneurOrdre.OnArgument(stArgument: String);
begin
Inherited ;
SetControlProperty('PDO_PGMODEREGLE','Datatype','PGMODEREGLE'); //PT2
SetPlusBanqueCP (GetControl ('PDO_RIBASSOCIE'));                //PT3
end;

procedure TOM_DonneurOrdre.OnLoadRecord;
var
{$IFNDEF EAGLCLIENT}
     Etab,Profil,Mode : THDBValComboBox;
{$ELSE}
     Etab,Profil,Mode : THValComboBox;
{$ENDIf}
     okok : boolean ;
begin
{$IFNDEF EAGLCLIENT}
Etab :=THDBValComboBox(GetControl('PDO_ETABLISSEMENT')) ;
Profil :=THDBValComboBox(GetControl('PDO_PROFIL')) ;
Mode :=THDBValComboBox(GetControl('PDO_PGMODEREGLE')) ; {PT1}
{$ELSE}
Etab :=THValComboBox(GetControl('PDO_ETABLISSEMENT')) ;
Profil :=THValComboBox(GetControl('PDO_PROFIL')) ;
Mode :=THValComboBox(GetControl('PDO_PGMODEREGLE')) ; {PT1}
{$ENDIF}
okok:=(DS<>NIL) and (DS.State=dsInsert) ;
if Etab <> NIL then Etab.Enabled :=okok;
if Profil <> NIL then Profil.Enabled :=okok;
if Mode <> NIL then Mode.Enabled :=okok;

if Etab<>nil then
  if Etab.Value='' then Etab.ItemIndex:=0;
if Profil<>nil then
  if Profil.Value='' then Profil.ItemIndex:=0;
end;


Initialization
registerclasses([TOM_DonneurOrdre]) ;
end.
 