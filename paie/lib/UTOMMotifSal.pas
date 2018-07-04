{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 20/06/2001
Modifié le ... :   /  /    
Description .. : TOM gestion des motifs d'entrée et des motifs de
Suite ........ : sortie
Suite ........ : Utilisé pour les Attestations et la DADS
Mots clefs ... : PAIE
*****************************************************************}
unit UTOMMotifSal;

interface
uses StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,HCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FichList,
{$ELSE}
     eFichList,UtileAgl,UTOB,
{$ENDIF}
     PgOutils2,Utom;
Type
      TOM_MotifEntreeSal = Class (TOM)
        procedure OnUpdateRecord ; override ;
        procedure OnDeleteRecord ; override;
     END ;

Type
      TOM_MotifSortieSal = Class (TOM)
        procedure OnUpdateRecord ; override ;
        procedure OnDeleteRecord ; override;
     END ;
implementation

{ TOM_MotifEntreeSal }
//============================================================================================//
//                               MOTIF ENTREE
//============================================================================================//
procedure TOM_MotifEntreeSal.OnDeleteRecord;
begin
  inherited;
ChargementTablette(TFFicheListe(Ecran).TableName,'');  //Rechargement des tablettes
end;

procedure TOM_MotifEntreeSal.OnUpdateRecord;
begin
  inherited;
   //Rechargement des tablettes
If (LastError=0) and (Getfield('PME_CODE')<>'') and (Getfield('PME_LIBELLE')<>'') then
   ChargementTablette(TFFicheListe(Ecran).TableName,'');
end;

{ TOM_MotifSortieSal }
//============================================================================================//
//                               MOTIF SORTIE
//============================================================================================//
procedure TOM_MotifSortieSal.OnDeleteRecord;
begin
  inherited;
ChargementTablette(TFFicheListe(Ecran).TableName,'');  //Rechargement des tablettes
end;

procedure TOM_MotifSortieSal.OnUpdateRecord;
begin
  inherited;
   //Rechargement des tablettes
If (LastError=0) and (Getfield('PMS_CODE')<>'') and (Getfield('PMS_LIBELLE')<>'') then
   ChargementTablette(TFFicheListe(Ecran).TableName,'');
end;

Initialization
registerclasses([TOM_MotifEntreeSal,TOM_MotifSortieSal]) ;
end.
 
