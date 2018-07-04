{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/04/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CORRESP (CORRESP)
Mots clefs ... : TOM;CORRESP
*****************************************************************}
Unit REGROUP_TOM ;

Interface

uses Windows, SysUtils, Classes, Controls, Forms, StdCtrls,
{$IFDEF EAGLCLIENT}
	{Utob, }eFiche, MaineAGL, eMul, eFichList, eTablette, eFichGrid, HPdfPrev,
{$ELSE}
	db, dbTables, Fiche, Fe_main, Mul, FichList, Tablette, FichGrid, EdtEtat,
{$ENDIF}
	HCtrls, HEnt1, HMsgBox, UTOM, UTob, UtilPgi, Grids, Ent1, EntGC, HDB;

procedure CCLanceFiche_Regroupement(Quoi : String);

Type
  TOM_REGROUP = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
    Quoi : String ;
    FListe : THDBGrid ;
    Procedure ZoomGene(Sender : tobject) ;
end ;


Implementation

Uses Choix ;

{=======================================================================================}
procedure CCLanceFiche_Regroupement(Quoi : String);
begin
AGLLanceFiche('CP','REGROUP','','',Quoi);
end;

{=======================================================================================}
procedure TOM_REGROUP.OnNewRecord ;
begin
  Inherited ;
  // Renseigne le champ CR_TYPE automatiquement
	DS.Fields.Fields[0].Value := 'ZZG';
end ;

{=======================================================================================}
procedure TOM_REGROUP.OnArgument (S: String);
begin
  Inherited ;
	Quoi:=S ;;

  // Contexte d'aide

  DS.Filtered := True;
  DS.Filter := 'CR_TYPE='''+Quoi+'';
  FListe := THDBGrid(GetControl('FListe'));	if (FListe <> nil) then FListe.OnDblClick := ZoomGene;

end ;

Procedure TOM_REGROUP.ZoomGene(Sender : tobject) ;
Var St,St1 : String ;
BEGIN
St:=Trim(FListe.SelectedField.Text) ;
St1:=Choisir('Comptes généraux','GENERAUX','G_GENERAL||"  "||G_LIBELLE','G_GENERAL','','G_GENERAL') ;
If St1<>'' Then
  BEGIN
//  If TCorresp.State=dsBrowse Then TCorresp.Edit ;
  FListe.Fields[FListe.SelectedIndex].AsString:=St1 ;
  END ;

END ;


{=======================================================================================}
Initialization
  registerclasses ( [ TOM_REGROUP ] ) ;
end.
