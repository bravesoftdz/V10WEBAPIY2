{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 28/04/2003
Modifié le ... : 12/08/2004
Description .. : 
Mots clefs ... : DP;JURI
*****************************************************************}
unit CLASSMenuNew;

interface

uses
{$IFDEF EAGLCLIENT}
   MaineAGL,
{$ELSE}
   //dbTables,
{$ENDIF}
   GenMenu, CLASSMenuElement, Menus, HTB97, hctrls, UTOB;

/////////////////////////////////////////////////////////////////
type
   TMenuNew = class(TMenuElement)
   private
      btNouveau_c : TToolbarButton97;
      procedure IniTMenuNew(MyPopUpMenu_p : TPopUpMenu;
                            cmbMyTablette_p : THValComboBox); reintroduce; overload;

      procedure IniTMenuNew(MyPopUpMenu_p : TPopUpMenu;
                            sListeCaption_p, sListeHint_p : string); reintroduce; overload;

      procedure IniTMenuNew(MyPopUpMenu_p : TPopUpMenu;
                            OBElements_p : TOB;
                            sChampCaption_p, sChampHint_p : string); reintroduce; overload;
   public
      sValue_c : string;
      Procedure Init(MyPopUpMenu_p : TPopUpMenu;
                     btNouveau_p : TToolbarButton97;
                     cmbMyTablette_p : THValComboBox); reintroduce; overload;

      Procedure Init(MyPopUpMenu_p : TPopUpMenu;
                     btNouveau_p : TToolbarButton97;
                     sListeCaption_p, sListeHint_p : string); reintroduce; overload;

      Procedure Init(MyPopUpMenu_p : TPopUpMenu;
                     btNouveau_p : TToolbarButton97;
                     OBElements_p : TOB;
                     sChampCaption_p, sChampHint_p : string); reintroduce; overload;

      procedure OnClickCmbNew(Sender : TObject);
   end;
/////////////////////////////////////////////////////////////////

implementation

/////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Procedure TMenuNew.Init(MyPopUpMenu_p : TPopUpMenu; btNouveau_p : TToolbarButton97;
                        cmbMyTablette_p : THValComboBox);
begin
   InitMenuElement( MyPopUpMenu_p );
   btNouveau_c := btNouveau_p;
   IniTMenuNew( MyPopUpMenu_p, cmbMyTablette_p);
end;

Procedure TMenuNew.Init(MyPopUpMenu_p : TPopUpMenu; btNouveau_p : TToolbarButton97;
                        sListeCaption_p, sListeHint_p : string);
begin
   InitMenuElement( MyPopUpMenu_p );
   btNouveau_c := btNouveau_p;
   IniTMenuNew( MyPopUpMenu_p, sListeCaption_p, sListeHint_p);
end;

Procedure TMenuNew.Init(MyPopUpMenu_p : TPopUpMenu;
                        btNouveau_p : TToolbarButton97;
                        OBElements_p : TOB;
                        sChampCaption_p, sChampHint_p : string);
begin
   InitMenuElement( MyPopUpMenu_p );
   btNouveau_c := btNouveau_p;
   IniTMenuNew( MyPopUpMenu_p, OBElements_p, sChampCaption_p, sChampHint_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMenuNew.IniTMenuNew(MyPopUpMenu_p : TPopUpMenu;
                               cmbMyTablette_p : THValComboBox);
begin
   InitMenuDepuisCombo(MyPopUpMenu_p.Items, cmbMyTablette_p, OnClickCmbNew );
end;

procedure TMenuNew.IniTMenuNew(MyPopUpMenu_p : TPopUpMenu;
                               sListeCaption_p, sListeHint_p : string);
begin
   InitMenuDepuisListe(MyPopUpMenu_p.Items, sListeCaption_p, sListeHint_p, OnClickCmbNew );
end;

procedure TMenuNew.IniTMenuNew(MyPopUpMenu_p : TPopUpMenu; OBElements_p : TOB;
                               sChampCaption_p, sChampHint_p : string);
begin
   InitMenuDepuisTob(MyPopUpMenu_p.Items, OBElements_p, sChampCaption_p, sChampHint_p, OnClickCmbNew );
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TMenuNew.OnClickCmbNew(Sender : TObject);
begin
   sValue_c := OnClickElementMenu( Sender );
   btNouveau_c.Click;
   sValue_c := '';
end;

/////////////////////////////////////////////////////////////////

end.
