{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 04/09/2002
Modifié le ... : 04/11/2004
Description .. :
Mots clefs ... : 
*****************************************************************}
unit UCLASSMyMenu;

interface

uses
{$IFDEF EAGLCLIENT}

{$ELSE}

{$ENDIF}
   Classes, Menus, Hctrls, UTOB, UCLASSMyMenuItem;

type
   TMyMenu = class(TMyMenuItem)

      function GetHint : string;
      function GetIndex : integer;
      function GetFirst : integer;
      
      //procedure OnMenuDetruit;
      procedure OnMenuFiltre(sRestriction_p : string);

   protected

   public
      constructor Create(MyPopUp_p : TPopUpMenu;
                         cmbMyTablette_p : THValComboBox;
                         AfterMenuClick_p : TNotifyEvent = nil); reintroduce; overload;
      constructor Create(MyPopUp_p : TPopUpMenu;
                         OBElements_p : TOB; sCaption_p, sHint_p : string;
                         AfterMenuClick_p : TNotifyEvent = nil); reintroduce; overload;
      constructor Create(MyPopUp_p : TPopUpMenu;
                         sRequete_p, sCaption_p, sHint_p : string;
                         AfterMenuClick_p : TNotifyEvent = nil); reintroduce; overload;
      constructor Create(MyPopUp_p : TPopUpMenu;
                         sCaption_p, sHint_p : string;
                         AfterMenuClick_p : TNotifyEvent = nil{; sListeChecked_p : string = ''}); reintroduce; overload;

      destructor Destroy; override;
   private

      AfterMenuClick_c : TNotifyEvent;

      procedure OnMenuClick(Sender : TObject); override;
      
   end;


implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... : 08/11/2004
Description .. : 
Mots clefs ... : 
*****************************************************************}
//////////////////////////////////////////////////////////////////
constructor TMyMenu.Create(MyPopUp_p : TPopUpMenu; cmbMyTablette_p : THValComboBox;
                           AfterMenuClick_p : TNotifyEvent = nil);
begin
   MenuInit(MyPopUp_p);
//   iFirstInd_c := MyPopUp_p.Items.Count;
   AfterMenuClick_c := AfterMenuClick_p;
   AddItems(MyPopUp_p.Items, cmbMyTablette_p, OnMenuClick);
end;
//////////////////////////////////////////////////////////////////
constructor TMyMenu.Create(MyPopUp_p : TPopUpMenu; OBElements_p : TOB; sCaption_p, sHint_p : string;
                           AfterMenuClick_p : TNotifyEvent = nil);
begin
   MenuInit(MyPopUp_p);
//   iFirstInd_c := MyPopUp_p.Items.Count;
   AfterMenuClick_c := AfterMenuClick_p;
   AddItems(MyPopUp_p.Items, OBElements_p, sCaption_p, sHint_p, OnMenuClick);
end;
//////////////////////////////////////////////////////////////////
constructor TMyMenu.Create(MyPopUp_p : TPopUpMenu; sRequete_p, sCaption_p, sHint_p : string;
                           AfterMenuClick_p : TNotifyEvent = nil);
begin
   MenuInit(MyPopUp_p);
//   iFirstInd_c := MyPopUp_p.Items.Count;
   AfterMenuClick_c := AfterMenuClick_p;
   AddItems(MyPopUp_p.Items, sRequete_p, sCaption_p, sHint_p, OnMenuClick);
end;
//////////////////////////////////////////////////////////////////
constructor TMyMenu.Create(MyPopUp_p : TPopUpMenu; sCaption_p, sHint_p : string;
                           AfterMenuClick_p : TNotifyEvent = nil{; sListeChecked_p : string = ''});
begin
   MenuInit(MyPopUp_p);
//   iFirstInd_c := MyPopUp_p.Items.Count;
   AfterMenuClick_c := AfterMenuClick_p;
   AddItems(MyPopUp_p.Items, sCaption_p, sHint_p, OnMenuClick);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure TMyMenu.OnSubAdd(iIndice_p : integer;
                           sRequete_p, sCaption_p, sHint_p : string);
begin
    AddItems(GetPopUp.Items[iIndice_p],
             sRequete_p, sCaption_p, sHint_p,
             OnMenuClick);
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyMenu.OnMenuClick(Sender : TObject);
begin
   inherited;
   if (@AfterMenuClick_c <> nil) then
   begin
      AfterMenuClick_c(Sender);
      sCurHint_c := '';
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TMyMenu.Destroy;
begin
   MenuDetruit(iFirstInd_c);
   inherited Destroy;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyMenu.OnMenuFiltre(sRestriction_p : string);
begin
   MenuFiltre(GetFirst, sRestriction_p);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenu.GetHint : string;
begin
   result := sCurHint_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenu.GetIndex : integer;
begin
   result := iCurInd_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenu.GetFirst : integer;
begin
   result := iFirstInd_c;
end;


end.
