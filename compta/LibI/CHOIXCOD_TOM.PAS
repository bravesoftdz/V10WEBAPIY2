{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 13/03/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CHOIXCOD (CHOIXCOD)
               : Appel 1 fiche de saisie sur la table Choixcod
               : code IGI pour les regroupements
Mots clefs ... : TOM;CHOIXCOD
*****************************************************************}
Unit CHOIXCOD_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     Fiche,
     FichList,
     Fe_main,
     FichGrid,
{$else}
     eFiche,
     eFichList,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     HTB97,
     Windows,
     Messages,
     Dialogs,
     Grids,
     UTob ;

procedure LanceFiche_SaisieTableCommune(pszType : String);

Type
  TOM_CHOIXCOD = Class (TOM)
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
    private
    //Suppression : TNotifyEvent;
    gszFType  : string ;
    FListe : THGrid ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    //procedure SupprimerClick (Sender : TObject);
    end ;

Implementation

procedure LanceFiche_SaisieTableCommune(pszType : String);
begin
  Begin
    IF pszType = 'IGI' Then
      AGLLanceFiche('AM','AMSREGRO','','',pszType);
  End;
end;


procedure TOM_CHOIXCOD.OnNewRecord ;
begin
  Inherited ;
     IF gszFType='IGI' Then
      // Renseigne le champ CC_TYPE automatiquement
      SetField('CC_TYPE', 'IGI');
end ;

procedure TOM_CHOIXCOD.OnDeleteRecord ;
begin
  Inherited ;
end ;

//procedure TOM_CHOIXCOD.SupprimerClick(Sender : TObject);
//begin
//  Suppression(Sender);
//end;

procedure TOM_CHOIXCOD.OnUpdateRecord ;
begin
  Inherited ;
   Begin

     if (Trim(GetField('CC_CODE')) = '') then
       begin
         PGIBox('Vous devez renseigner un code regroupement.', Ecran.Caption ) ;
         LastError := 1;
         exit;
       end

     { Essai de gérer si flèche Haute sur ligne non saisie suppression de cette ligne
     if (Trim(GetField('CC_CODE')) = '') and (Trim(GetField('CC_LIBELLE'))<>'')  then
       begin
         PGIBox('Vous devez renseigner un code regroupement.', Ecran.Caption ) ;
         LastError := 1;
         exit;
       end
      else if (Trim(GetField('CC_CODE')) = '') then
       begin
         // Test pour forcer la suppression d'1 ligne blanche
         // => ne supprime pas
         SupprimerClick(TToolbarButton97(GetControl('BDelete')));

         // Essai pour gérer manuellement la suppression si ligne blanche
         //=> provoque une erreur indice hors limite  ??
         executeSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="IGI" AND CC_CODE="'+s1+'"') ;
         Nbl:=Fliste.Row;
         FListe.Objects[0,Row] := nil ;
         FListe.DeleteRow(FListe.Row) ;
       end;
      }
   end;
end ;

procedure TOM_CHOIXCOD.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CHOIXCOD.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CHOIXCOD.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CHOIXCOD.OnChangeField ( F: TField ) ;
var
  SL,SA : String;
begin
  Inherited ;
   If F.FieldName='CC_LIBELLE' then
     // Libellé abrégé = 17 premiers caractères du libellé
     Begin
       SL := GetField('CC_LIBELLE');
       SA := GetField('CC_ABREGE');
       If Trim(SA)='' then
          SetField('CC_ABREGE',Copy(SL,1,17));
     End;
end ;

Procedure TOM_CHOIXCOD.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
  If (ARow<>0) AND (Acol=1) then
    // Mise en gras de la première colonne
    Fliste.Canvas.Font.Style:=Fliste.Canvas.Font.Style+[fsBold];
END ;


procedure TOM_CHOIXCOD.OnArgument ( S: String ) ;
begin
  Inherited ;
   gszFType:=S ;
   FListe := THGrid(GetControl('FListe'));
   FListe.GetCellCanvas:=GetCellCanvas;
   // Sauvegarde de l'action supprimer
   //Suppression := TToolbarButton97(GetControl('BDelete')).OnClick;
   //TToolbarButton97(GetControl('BDelete')).OnClick := SupprimerClick;
end ;

procedure TOM_CHOIXCOD.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CHOIXCOD.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_CHOIXCOD ] ) ;
end.
