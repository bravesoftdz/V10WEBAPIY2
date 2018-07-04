unit UCLASSKeys;

interface

uses
   {$IFDEF EAGLCLIENT}
   emul,
   {$ELSE}
   mul,
   {$ENDIF}
   Vierge,
   Classes, HTB97, controls;

const vk_Imp_c = 1;
const vk_Ins_c = 2;
const vk_Ouv_c = 3;
const vk_Val_c = 4;
const vk_Sup_c = 5;

Type
   TKeysPlus = Class
      procedure AddBouton(aoiType_p : array of integer; aobBouton_p : array of TToolbarButton97); overload;
      procedure AddBouton(iType_p : integer; bBouton_p : TToolbarButton97); overload;
   public
   protected
      BImprimer_c  : TToolbarButton97;
      Binsert_c    : TToolbarButton97;
      BOuvrir_c    : TToolbarButton97;
      BValide_c    : TToolbarButton97;
      BSupprimer_c : TToolbarButton97;

      function OnFormKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState) : boolean;
   private
   end;

Type
   TKeysPlusMul = Class(TKeysPlus)
   public
      constructor Create(MyMul_p : TFMul;
                         aoiType_p : array of integer;
                         aobBouton_p : array of TToolbarButton97); reintroduce; overload;
      procedure OnMulKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
   protected
   private
      MyMul_c : TFMul;
   end;

Type
   TKeysPlusVie = Class(TKeysPlus)
   public
      constructor Create(MyVie_p : TFVierge;
                         aoiType_p : array of integer;
                         aobBouton_p : array of TToolbarButton97); reintroduce; overload;
      procedure OnVieKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
   protected
   private
      MyVie_c : TFVierge;
   end;

{Type
   TKeysPlusFic = Class(TKeysPlus)
   public
      constructor Create(MyFic_p : TFFiche;
                         aoiType_p : array of integer;
                         aobBouton_p : array of TToolbarButton97); reintroduce; overload;
      procedure OnFicKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
   protected
   private
      MyFic_c : TFFiche;
   end;}
//////////////////////////////////////////////////////////////////
implementation
//////////////////////////////////////////////////////////////////
uses
   HEnt1, Windows;
//////////////////////////////////////////////////////////////////
// TKEYSPLUS
//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TKeysPlus.AddBouton(aoiType_p : array of integer; aobBouton_p : array of TToolbarButton97);
var
   iInd_l : integer;
begin
   for iInd_l := 0 to length(aobBouton_p) - 1 do
   begin
      AddBouton(aoiType_p[iInd_l], aobBouton_p[iInd_l]);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 26/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TKeysPlus.AddBouton(iType_p : integer; bBouton_p : TToolbarButton97);
begin
   if iType_p = vk_imp_c then
      BImprimer_c := bBouton_p
   else if iType_p = vk_ins_c then
      Binsert_c := bBouton_p
   else if iType_p = vk_ouv_c then
      BOuvrir_c := bBouton_p
   else if iType_p = vk_val_c then
      BValide_c := bBouton_p
   else if iType_p = vk_Sup_c then
      BSupprimer_c := bBouton_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TKeysPlus.OnFormKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState) : boolean;
var
   Point_l : TPoint;
begin
   result := true;
   if (MyKey_p = VK_DELETE) and (MyShift_p = [ssCtrl]) then  // Ctrl + suppr : supprimer
   begin
      MyKey_p := 0;
      if (BSupprimer_c <> nil) and (BSupprimer_c.Visible) and (BSupprimer_c.Enabled) then
         BSUPPRIMER_c.OnClick(nil);
   end
   else if (MyKey_p = vk_nouveau)  and (MyShift_p = [ssCtrl]) then  // Ctrl + N : nouveau
   begin
      MyKey_p := 0;
      if (Binsert_c <> nil) and (Binsert_c.Visible)  and (Binsert_c.Enabled)then
      begin
         if Binsert_c.DropdownMenu <> nil then
         begin
            Point_l.x := Binsert_c.Left;
            Point_l.y := Binsert_c.Top + Binsert_c.Height;
            Point_l := Binsert_c.Parent.ClientToScreen(Point_l);
            Binsert_c.DropdownMenu.Popup(Point_l.X, Point_l.Y);
         end
         else
            Binsert_c.OnClick(nil);
      end
   end
   else if (MyKey_p = vk_imprime)  and (MyShift_p = [ssCtrl]) then  // Ctrl + P : imprimer
   begin
      MyKey_p := 0;
      if (BImprimer_c <> nil) and (BImprimer_c.Visible) and (BImprimer_c.Enabled) then
         BImprimer_c.OnClick(nil);
   end
   else if (MyKey_p = VK_F10) then  // Enregistrement ou validation
   begin
      MyKey_p := 0;
      if (BValide_c <> nil) and (BValide_c.Visible) and (BValide_c.Enabled) then
         BValide_c.OnClick(nil);
   end
   else
      result := false;
end;
//////////////////////////////////////////////////////////////////
// TKEYSPLUSMUL
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TKeysPlusMul.Create(MyMul_p : TFMul;
                                aoiType_p : array of integer;
                                aobBouton_p : array of TToolbarButton97);
begin
   MyMul_c := MyMul_p;
   Binsert_c := MyMul_p.Binsert;
   BImprimer_c := MyMul_p.BImprimer;
   BOuvrir_c := MyMul_p.BOuvrir;
   AddBouton(aoiType_p, aobBouton_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TKeysPlusMul.OnMulKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
var
   Point_l : TPoint;
begin
   if (MyKey_p = VK_F5) then  // Sélection ouvrir
   begin
      MyKey_p := 0;
      if MyMul_c.FListe.Focused then
      begin
         if (BOuvrir_c <> nil) and (BOuvrir_c.Visible) and (BOuvrir_c.Enabled) then
            BOuvrir_c.OnClick(nil)
      end;
   end
   else if (MyKey_p = VK_F11) then  // Sélection ouvrir
   begin
      MyKey_p := 0;
      if (MyMul_c.FListe.Focused) and (MyMul_c.FListe.PopupMenu <> nil) then
      begin
         Point_l.x := round(MyMul_c.FListe.Width / 3);
         Point_l.y := MyMul_c.FListe.Top +
                    (MyMul_c.FListe.Row + 1) *
                    (MyMul_c.FListe.RowHeights[1]);

         Point_l := MyMul_c.FListe.Parent.ClientToScreen(Point_l);
         MyMul_c.FListe.PopupMenu.Popup(Point_l.x, Point_l.y);
      end;
   end
   else if not OnFormKeyDown(MySender_p, MyKey_p, MyShift_p) then
      MyMul_c.FormKeyDown(MySender_p, MyKey_p, MyShift_p);
end;
//////////////////////////////////////////////////////////////////
// TKEYSPLUSVIE
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TKeysPlusVie.Create(MyVie_p : TFVierge;
                                aoiType_p : array of integer;
                                aobBouton_p : array of TToolbarButton97);
begin
   MyVie_c := MyVie_p;
   Binsert_c := MyVie_p.Binsert;
   BImprimer_c := MyVie_p.BImprimer;
   BSupprimer_c := MyVie_p.BDelete;
   BValide_c := MyVie_p.BValider;
   AddBouton(aoiType_p, aobBouton_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TKeysPlusVie.OnVieKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
begin
   if (MyKey_p = VK_F5) then  // Sélection ouvrir
   begin
      MyKey_p := 0;
      if (BOuvrir_c <> nil) and (BOuvrir_c.Visible) and (BOuvrir_c.Enabled) then
         BOuvrir_c.OnClick(nil)
   end
   else if not OnFormKeyDown(MySender_p, MyKey_p, MyShift_p) then
      MyVie_c.FormKeyDown(MySender_p, MyKey_p, MyShift_p);
end;

//////////////////////////////////////////////////////////////////
// TKEYSPLUSFIC
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{constructor TKeysPlusFic.Create(MyFic_p : TFFiche;
                                aoiType_p : array of integer;
                                aobBouton_p : array of TToolbarButton97);
begin
   MyFic_c := MyFic_p;
   Binsert_c := MyFic_p.Binsert;
   BImprimer_c := MyFic_p.BImprimer;
   BSupprimer_c := MyFic_p.BDelete;
   BValide_c := MyFic_p.BValider;

   AddBouton(aoiType_p, aobBouton_p);
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure TKeysPlusFic.OnFicKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
begin
   MyFic_c.FormKeyDown(MySender_p, MyKey_p, MyShift_p);
end;}
//////////////////////////////////////////////////////////////////
end.

