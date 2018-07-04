{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 13/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit dpTOMAnnuLien;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   menus, SysUtils, Classes,
{$IFDEF EAGLCLIENT}
   eFiche,
{$ELSE}
   Fiche,
{$ENDIF}
   UTOMYYAnnulien, ClassMenuFonction;

type
   TOM_ANNULIEN = class(TOM_YYANNULIEN)
         procedure OnArgument(stArgument : String); override;
         procedure OnLoadRecord; override;
         procedure OnClose; override;
      private
         FMenuFonction_c : TMenuFonction;
      public
         procedure OnClick_BDELETE(Sender: TObject);

    { Déclarations publiques }
   end;

/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////
uses
   CLASSAnnuLien, UtilLiensAnnuaire;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 13/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_ANNULIEN.OnArgument(stArgument: String);
begin
   inherited;
   SetControlVisible('FE_ENTETEJUR', false);
   SetControlVisible('FE_ENTETEDP', true);

   if GetControl('POPMENUJUR') <> nil then
   begin
      FMenuFonction_c := TMenuFonction.Create;
      FMenuFonction_c.Init(TPopUpMenu(GetControl('POPMENUJUR')), FLien_c.nNoOrdre_c,
                           FLien_c.sGuidPerDos_c, FLien_c.sGuidPer_c,
                           FLien_c.sTypeDos_c, FLien_c.sForme_c );
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 05/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_ANNULIEN.OnLoadRecord;
begin
   inherited;
   SetControlEnabled('ANL_FONCTION', false);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_ANNULIEN.OnClose;
begin
   if (GetControl('POPMENUJUR') <> nil) and (FMenuFonction_c <> nil) then
   begin
     FMenuFonction_c.DetruitMenu(0);
     FreeAndNil( FMenuFonction_c );
   end;
   inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/03/2003
Modifié le ... :   /  /
Description .. : lors de la suppression d'un lien dans une fiche d'intervention,
                 permet de supprimer le dernier lien INT lié à la loi NRE
Mots clefs ... :
*****************************************************************}
procedure TOM_ANNULIEN.OnClick_BDELETE(Sender: TObject);
var
   sGuidPerDos_l, sGuidPer_l : string;
   sTypeDos_l, sFonction_l : string;

begin
   sGuidPerDos_l := GetField('ANL_GUIDPERDOS');
   sTypeDos_l := GetField('ANL_TYPEDOS');
   sFonction_l := GetField('ANL_FONCTION');
   sGuidPer_l := GetField('ANL_GUIDPER');

   if (sGuidPerDos_l = '') or (sGuidPer_l = '') then
      exit;

   // suppression enreg en cours
   TFFiche(Ecran).BDeleteClick(Nil);

   // suppression avec param "intseulmt" à True
   // => ne traite que l'enreg INT
   // => et si on a refusé la suppression lors du bdeleteclick,
   //    l'enreg en cours INT ne sera pas supprimé
   SupprimeLienAnnuaire(sGuidPerDos_l, sGuidPer_l, sTypeDos_l, sFonction_l, true);

end;


Initialization
registerclasses([TOM_ANNULIEN]) ;
end.
