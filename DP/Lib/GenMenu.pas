{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Unité ........ : GenMenu
Description .. : Gestion des menus
*****************************************************************}

unit GenMenu;

interface
uses
   UTOB,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}

   Menus, classes, hctrls;

function  InitPopUpSubMenu( MyPopUpMenu_p : TPopUpMenu; sListeCaptions_p : string ) : integer;
procedure InitPopUpMenu( MyPopUpMenu_p : TPopUpMenu; sListeCaption_p, sListeHint_p : string;
                  neOnClick_p : TNotifyEvent; sListeChecked_p : string);
function  InitSubMenu( MySubMenu_p : TMenuItem; sCaption_p : string ) : integer;
procedure InitItem( MySubMenu_p : TMenuItem; sCaption_p : string; sHint_p : string = '';
                  neOnClick_p : TNotifyEvent = nil; bChecked_p : boolean = false );

function  InitMenuDepuisTob( MyMenu_p: TMenuItem; TOBElement_p : TOB; sChampCaption_p, sChampHint_p : string;
                  neOnClick_p : TNotifyEvent ) : integer;
function  InitMenuDepuisListe( MyMenu_p: TMenuItem; sListeCaption_p, sListeHint_p : string;
                  neOnClick_p : TNotifyEvent; sListeChecked_p : string   = '') : integer;
function InitMenuDepuisCombo( MyMenu_p: TMenuItem; cmbMyTablette_p : THValComboBox;
                  neOnClick_p : TNotifyEvent ) : integer;

procedure CheckItem( MySubMenu_p : TMenuItem; sValeur_p : string );

procedure FiltrePopUp( MyPopUpMenu_p : TPopUpMenu; nIndex_p : integer;
                       sRestriction_p : string );
procedure FiltreSubMenu( MyPopUpMenu_p : TPopUpMenu; nIndex_p : integer;
                         sRestriction_p : string );
procedure FiltreItem( MySubMenu_p: TMenuItem; nIndex_p : integer;
                      sRestriction_p : string );
procedure DeletePopUp( MyPopUpMenu_p : TPopUpMenu; nIndex_p : integer );
procedure DeleteItem( MySubMenu_p: TMenuItem; nIndex_p : integer );
function InListeOfString(sValeur_p, sListe_p : string) : boolean;

implementation


{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Fonction ..... : InitPopUpSubMenu
Description .. : Initialise le popup menu avec des sous menus
Paramètres ... : Le menu popup
                 La liste des libellés des sous menus
Renvoie ...... : Le nombre d'éléments
*****************************************************************}

function InitPopUpSubMenu( MyPopUpMenu_p : TPopUpMenu; sListeCaptions_p : string ) : integer;
var
   sCaption_l : string;
begin
   sCaption_l := ReadTokenSt( sListeCaptions_p );
   while sCaption_l <> '' do
   begin
      InitSubMenu( MyPopUpMenu_p.Items, sCaption_l );
      sCaption_l := ReadTokenSt( sListeCaptions_p );
   end;
   result := MyPopUpMenu_p.Items.Count - 1;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Procédure .... : InitPopUpMenu
Description .. : Initialise le popup menu
Paramètres ... : Le menu popup
                 La liste des libellés des sous menus
*****************************************************************}

procedure InitPopUpMenu( MyPopUpMenu_p : TPopUpMenu; sListeCaption_p, sListeHint_p : string;
                  neOnClick_p : TNotifyEvent; sListeChecked_p : string );
begin
   InitMenuDepuisListe( MyPopUpMenu_p.Items, sListeCaption_p, sListeHint_p, neOnClick_p, sListeChecked_p );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Fonction ..... : InitSubMenu
Description .. : Initialise le sous menu
Paramètres ... : Le menu popup
                 Le libellé de l'item
Renvoie ...... : Le nombre d'éléments
*****************************************************************}

function InitSubMenu( MySubMenu_p : TMenuItem; sCaption_p : string ) : integer;
begin
   InitItem( MySubMenu_p, sCaption_p, '', nil );
   result := MySubMenu_p.Count;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Procédure .... : InitItem
Description .. : Initialise l'item
Paramètres ... : Le menu popup
                 Le libellé de l'item
                 Le code correspondant de l'item
                 L'évenement à appliquer
*****************************************************************}

procedure InitItem( MySubMenu_p : TMenuItem; sCaption_p : string; sHint_p : string = '';
                    neOnClick_p : TNotifyEvent = nil; bChecked_p : boolean = false );
var
   MyItems_l: TMenuItem;
begin
   MyItems_l := TMenuItem.Create( MySubMenu_p );
   MyItems_l.Caption := sCaption_p;
   MyItems_l.Hint := sHint_p;
   MyItems_l.OnClick := neOnClick_p;
   MyItems_l.Checked := bChecked_p;
   MySubMenu_p.Add( MyItems_l );

end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Fonction ..... : InitMenuDepuisTob
Description .. : Initialise les items du menu avec une TOB
Paramètres ... : Le menu
                 La TOB
                 Le champ libellé de la TOB
                 Le champ code de la TOB
                 L'évenement à appliquer
Renvoie ...... : Le nombre d'éléments
*****************************************************************}

function InitMenuDepuisTob( MyMenu_p: TMenuItem; TOBElement_p : TOB; sChampCaption_p, sChampHint_p : string;
                  neOnClick_p : TNotifyEvent ) : integer;
var
   nItem_l : integer;
begin
   for nItem_l := 0 to TOBElement_p.Detail.count-1 do
   begin
      InitItem( MyMenu_p, VarToStr( TOBElement_p.Detail[nItem_l].GetValue(sChampCaption_p) ),
              VarToStr( TOBElement_p.Detail[nItem_l].GetValue(sChampHint_p) ), neOnClick_p );
   end;
   TOBElement_p.Free;
   result := MyMenu_p.Count;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function InitMenuDepuisCombo( MyMenu_p: TMenuItem; cmbMyTablette_p : THValComboBox;
                  neOnClick_p : TNotifyEvent ) : integer;
var
   nItem_l : integer;
begin
   for nItem_l := 0 to cmbMyTablette_p.Items.count - 1 do
   begin
      if (cmbMyTablette_p.Items[nItem_l] <> '<<Tous>>') and
         (cmbMyTablette_p.Values[nItem_l] <> '') then
         InitItem(MyMenu_p, cmbMyTablette_p.Items[nItem_l],
                  cmbMyTablette_p.Values[nItem_l], neOnClick_p );
   end;
   result := MyMenu_p.Count;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Fonction ..... : InitMenuDepuisListe
Description .. : Initialise les items du menu avec une liste
Paramètres ... : Le menu
                 La TOB
                 La liste de libellés
                 La liste de codes
                 L'évenement à appliquer
Renvoie ...... : Le nombre d'éléments                 
*****************************************************************}

function InitMenuDepuisListe( MyMenu_p: TMenuItem; sListeCaption_p, sListeHint_p : string;
                  neOnClick_p : TNotifyEvent; sListeChecked_p : string = '') : integer;
var
   sCaption_l, sHint_l, sChecked_l : string;
   bChecked_l : boolean;
begin
   sCaption_l := ReadTokenSt( sListeCaption_p );
   while sCaption_l <> '' do
   begin
      sHint_l := ReadTokenSt( sListeHint_p );
      bChecked_l := false;
      sChecked_l := ReadTokenSt( sListeChecked_p );

      if ( sChecked_l <> '' ) and ( sChecked_l[1] in ['1','t','T'] ) then
         bChecked_l := true;

      InitItem( MyMenu_p, sCaption_l, sHint_l, neOnClick_p, bChecked_l );
      sCaption_l := ReadTokenSt( sListeCaption_p );
   end;

   result := MyMenu_p.Count;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure FiltreSubMenu( MyPopUpMenu_p : TPopUpMenu; nIndex_p : integer;
                       sRestriction_p : string );
var
   nItem_l : integer;
begin
   for nItem_l := MyPopUpMenu_p.Items.Count - 1 downto nIndex_p do
   begin
      FiltreItem( MyPopUpMenu_p.Items[nItem_l], 0, sRestriction_p );
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure FiltrePopUp( MyPopUpMenu_p : TPopUpMenu; nIndex_p : integer;
                              sRestriction_p : string );
var
   nItem_l : integer;
   sElement_l : string;
begin
   for nItem_l := MyPopUpMenu_p.Items.Count - 1 downto nIndex_p do
   begin
    sElement_l := MyPopUpMenu_p.Items[nItem_l].Hint;
       if not InListeOfString(sElement_l, sRestriction_p) then
          MyPopUpMenu_p.Items[nItem_l].Free;
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure  FiltreItem( MySubMenu_p: TMenuItem; nIndex_p : integer;
                       sRestriction_p : string );
var
   nItem_l : integer;
   sElement_l : string;
begin
   for nItem_l := nIndex_p to MySubMenu_p.Count - 1 do
   begin
      sElement_l := MySubMenu_p.Items[nItem_l].Hint;
       if not InListeOfString(sElement_l, sRestriction_p) then
          MySubMenu_p.Items[nItem_l].Free;
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/11/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function InListeOfString(sValeur_p, sListe_p : string) : boolean;
var
   sString_l : string;
   bTrouve_l : boolean;
begin
   bTrouve_l := false;
   sString_l := READTOKENST( sListe_p);
   while (not bTrouve_l) and (sString_l <> '') do
   begin
      if sValeur_p = sString_l then
         bTrouve_l := true
      else
         sString_l := READTOKENST( sListe_p);      
   end;
   result := bTrouve_l;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Procédure .... : DeletePopUp
Description .. : Détruit les éléments du popup menu depuis nIndex_p
                 jusqu'à MyPopUpMenu_p.Items.Count
Paramètres ... : Le menu
                 L'index
*****************************************************************}

procedure DeletePopUp( MyPopUpMenu_p : TPopUpMenu; nIndex_p : integer );
var
   nItem_l : integer;
begin
   for nItem_l := MyPopUpMenu_p.Items.Count - 1 downto nIndex_p do
   begin
      DeleteItem( MyPopUpMenu_p.Items[nItem_l], 0 );
      MyPopUpMenu_p.Items[nItem_l].Free;
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Procedure .... : DeleteItem
Description .. : Détruit les éléments du sous menu depuis nIndex_p
Paramètres ... : Le menu
                 L'index
*****************************************************************}

procedure  DeleteItem( MySubMenu_p: TMenuItem; nIndex_p : integer );
var
   nItem_l : integer;
begin
   for nItem_l := MySubMenu_p.Count - 1 downto nIndex_p do
   begin
      MySubMenu_p.Items[nItem_l].Free;
   end;
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 28/08/02
Procédure .... : CheckItem
Description .. : Initialise le sous menu
Paramètres ... : Le menu popup
                 Le libellé de l'item
*****************************************************************}

procedure CheckItem( MySubMenu_p : TMenuItem; sValeur_p : string );
var
   nItem_l : integer;
begin
   for nItem_l := 0 to MySubMenu_p.Count - 1 do
   begin
      if MySubMenu_p.Items[nItem_l].Hint = sValeur_p then
         MySubMenu_p.Items[nItem_l].Checked := true
      else
         MySubMenu_p.Items[nItem_l].Checked := false;
   end;
end;

end.
