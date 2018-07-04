{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMSREGRO()
Mots clefs ... : TOF;AMSREGRO
*****************************************************************}
Unit AMSREGRO_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     graphics,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Grids,
     UTOF,
     uTOB,
     vierge,
     HTB97;

Type
  TOF_AMSREGRO = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  Private
    FListe : THGrid ;
    Faction: Boolean;
    OldCloseQuery : TCloseQueryEvent ;
    Ffermeture : TNotifyEvent;
    Fvalidation: TNotifyEvent;
    Procedure DoCloseQuery( Sender : TObject ; var CanClose : Boolean ) ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    procedure OnExitCC_LIBELLE(Sender:TObject);
    procedure OnExitCC_CODE(Sender:TObject);
    procedure OnbFermeClick(Sender: TObject);
    procedure OnbpostClick(Sender: TObject);
end ;

procedure AMLanceFiche_SaisieRegroupement(Lequel: string);

Implementation

const
  HM: array[0..0] of string =
   ('Voulez-vous enregistrer cette opération ?');


{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Appel de la saisie
Mots clefs ... :
*****************************************************************}
procedure AMLanceFiche_SaisieRegroupement(Lequel: string);
begin
  AGLLanceFiche('AM', 'AMSREGRO', '', '', Lequel);
  // Pour alimenter en sortie les combo utilisant les regroupements
  AvertirTable('AMREGROUPEMENT');
end;


{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Mise à jour de la saisie
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSREGRO.OnUpdate ;
begin
  Inherited ;
end ;


{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Préparation de la saisie
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSREGRO.OnArgument (S : String ) ;
begin
   Inherited ;
   FListe := THGrid(GetControl('FListe'));
   FListe.GetCellCanvas:=GetCellCanvas;
   ThEdit(GetControl ('CC_LIBELLE')).OnExit := OnExitCC_LIBELLE;
   ThEdit(GetControl ('CC_CODE')).OnExit := OnExitCC_CODE;

   // Tri du grid
   FListe.SortEnabled:=True;
   Fliste.SortGrid(1,true);

   // Sauvegarde de l'action fermer
   Ffermeture := TToolbarButton97(GetControl('BFerme')).OnClick;
   TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;

   // Sauvegarde de l'action valider
   Fvalidation := TToolbarButton97(GetControl('Bpost')).OnClick;
   TToolbarButton97(GetControl('Bpost')).Onclick := OnbpostClick;

   OldCloseQuery := ecran.OnCloseQuery; //Sauve la gestion de la fiche
   ecran.OnCloseQuery := DoCloseQuery; //assigne ta proper gestion
end ;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Affichage en cours de saisie du libellé abrégé
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSREGRO.OnExitCC_LIBELLE(Sender:TObject);
var
  SL,SA : String;
begin
  Inherited ;
     SL := GetControlText('CC_LIBELLE');
     SA := GetControlText('CC_ABREGE');
     If Trim(SA)='' then
          SetControlText('CC_ABREGE',Copy(SL,1,17));
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Contrôle de la saisie d'un code
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSREGRO.OnExitCC_CODE(Sender:TObject);
begin
  Inherited ;
  if (Trim(GetControlText('CC_CODE')) = '') then
    begin
       IF Faction = False Then
          PGIBox('Vous devez renseigner un code regroupement.', Ecran.Caption ) ;
       ecran.ModalResult :=mrNone;
       SetFocusControl('CC_CODE');
    end;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 04/11/2005
Modifié le ... :   /  /
Description .. : Gestion de la police en gras de la première colonne
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMSREGRO.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
  If (ARow<>0) AND (Acol=1) then
    // Mise en gras de la première colonne
    Fliste.Canvas.Font.Style:=Fliste.Canvas.Font.Style+[fsBold];
END ;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 08/06/2006
Modifié le ... :   /  /
Description .. : Sortie par la X en bas
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSREGRO.OnbFermeClick(Sender: TObject);
begin
  Inherited ;
  if (Trim(GetControlText('CC_CODE')) = '') then
    begin
      Faction := True;
      PGIBox('Vous devez renseigner un code regroupement.', Ecran.Caption ) ;
      ecran.ModalResult :=mrNone;
      SetFocusControl('CC_CODE');
      Faction := False;
    end
  else
    begin
      // Pour valider la saisie du code
      SetFocusControl('CC_LIBELLE');
      Ffermeture(Sender);
    end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 08/06/2006
Modifié le ... :   /  /
Description .. : Sortie par la X en haut à droite
Mots clefs ... :
*****************************************************************}
Procedure TOF_AMSREGRO.DoCloseQuery( Sender : TObject ; var CanClose : Boolean ) ;
Begin
  if (Trim(GetControlText('CC_CODE')) = '') then
    begin
       Faction := True;
       PGIBox('Vous devez renseigner un code regroupement.', Ecran.Caption ) ;
       SetFocusControl('CC_CODE');
       Canclose := false;
       Faction := False;
    end
  ELSE
    Begin
       // Pour valider la saisie du code
       SetFocusControl('CC_LIBELLE');
       OldCloseQuery( Sender, CanClose );
    end;
End;



{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 08/06/2006
Modifié le ... :   /  /
Description .. : Validation de la saisie
Mots clefs ... :
*****************************************************************}
procedure TOF_AMSREGRO.OnbPostClick(Sender: TObject);
begin
  Inherited ;
  if (Trim(GetControlText('CC_CODE')) = '') then
    begin
      Faction := True;
      PGIBox('Vous devez renseigner un code regroupement.', Ecran.Caption ) ;
      SetFocusControl('CC_CODE');
      Faction := False;
    end
  else
    begin
      // Pour valider la saisie du code
      SetFocusControl('CC_LIBELLE');
      FValidation(Sender);
    end;
end;





Initialization
  registerclasses ( [ TOF_AMSREGRO ] ) ;
end.
