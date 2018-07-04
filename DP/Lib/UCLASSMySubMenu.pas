{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 04/09/2002
Modifié le ... : 04/11/2004
Description .. :
Mots clefs ... : 
*****************************************************************}
unit UCLASSMySubMenu;

interface

uses
   {$IFDEF EAGLCLIENT}
   eMul,
   {$ELSE}
   Mul,
   {$ENDIF}
   Classes, Menus, Hctrls, UTOB, UCLASSMyMenuItem;


type
   TMySubMenu = class(TMyMenuItem)

      function GetHint : string;
      function GetSubCaption : string;
      function GetSubLeHint : string;
      function GetSubIndice : integer;
      function IsSubHint(sHint_p : string) : integer;

      //procedure OnMenuDetruit;
      procedure OnMenuFiltre(sRestriction_p : string);
   public
      procedure SubMenuAdd(iIndice_p : integer;
                           sRequete_p, sCaption_p, sHint_p : string;
                           AfterMenuClick_p : TNotifyEvent); reintroduce; overload;

      procedure SubMenuAdd(aosRequete_p, aosCaption_p, aosHint_p : array of string;
                           aoAfterMenuClick_p : array of TNotifyEvent); reintroduce; overload;

      procedure SubMenuAdd(iIndice_p : integer;
                           cmbMyTablette_p : THValComboBox;
                           AfterMenuClick_p : TNotifyEvent); reintroduce; overload;

      procedure SubMenuAdd(iIndice_p : integer;
                           cmbMyMultiTablette_p : THMultiValComboBox;
                           AfterMenuClick_p : TNotifyEvent); reintroduce; overload;

      procedure SubMenuAdd(iIndice_p : integer;
                           OBElement_p : TOB; sCaption_p, sHint_p : string;
                           AfterMenuClick_p : TNotifyEvent); reintroduce; overload;

      procedure SubMenuAdd(OBElements_p : TOB; sCaption_p, sHint_p, sMaitre_p : string;
                           AfterMenuClick_p : TNotifyEvent); reintroduce; overload;

      Constructor Create(MyPopUp_p : TPopUpMenu;
                         cmbMyTablette_p : THValComboBox;
                         bSep_p : boolean = false); reintroduce; overload;

      Constructor Create(MyPopUp_p : TPopUpMenu;
                         OBElements_p : TOB; sCaption_p : string;
                         bSep_p : boolean = false); reintroduce; overload;

      Constructor Create(MyPopUp_p : TPopUpMenu;
                         sRequete_p, sCaption_p : string;
                         bSep_p : boolean = false); reintroduce; overload;

      Constructor Create(MyPopUp_p : TPopUpMenu;
                         sCaption_p : string;
                         bSep_p : boolean = false); reintroduce; overload;

      Constructor Create(MyPopUp_p : TPopUpMenu;
                         frmMul_p : TFMul; sChamps_p : string;
                         bSep_p : boolean = false); reintroduce; overload;

      destructor Destroy; override;
      
   private
      iSep_c : integer;
      iSousMenu_c : integer;
      sSubCaption_c, sSubHint_c : string;
      AfterMenuClick_c : array of TNotifyEvent;
         
      Procedure SubMenuInit(MyPopUp_p : TPopUpMenu; bSep_p : boolean);
      procedure OnSubMenuClick(Sender : TObject);
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
Constructor TMySubMenu.Create(MyPopUp_p : TPopUpMenu;
                              cmbMyTablette_p : THValComboBox;
                              bSep_p : boolean = false);
begin
   MenuInit(MyPopUp_p);
   SubMenuInit(MyPopUp_p, bSep_p);
   AddItems(MyPopUp_p.Items, cmbMyTablette_p, OnSubMenuClick);
end;
//////////////////////////////////////////////////////////////////
Constructor TMySubMenu.Create(MyPopUp_p : TPopUpMenu;
                              OBElements_p : TOB; sCaption_p : string;
                              bSep_p : boolean = false);
begin
   MenuInit(MyPopUp_p);
   SubMenuInit(MyPopUp_p, bSep_p);
   AddItems(MyPopUp_p.Items, OBElements_p, sCaption_p, '', OnSubMenuClick);
end;

//////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////
Constructor TMySubMenu.Create(MyPopUp_p : TPopUpMenu;
                              sRequete_p, sCaption_p : string;
                              bSep_p : boolean = false);
begin
   MenuInit(MyPopUp_p);
   SubMenuInit(MyPopUp_p, bSep_p);
   AddItems(MyPopUp_p.Items, sRequete_p, sCaption_p, '', OnSubMenuClick);
end;
//////////////////////////////////////////////////////////////////
Constructor TMySubMenu.Create(MyPopUp_p : TPopUpMenu;
                              sCaption_p : string;
                              bSep_p : boolean = false);
begin
   MenuInit(MyPopUp_p);
   SubMenuInit(MyPopUp_p, bSep_p);
   AddItems(MyPopUp_p.Items, sCaption_p, '', OnSubMenuClick);
end;
//////////////////////////////////////////////////////////////////
Constructor TMySubMenu.Create(MyPopUp_p : TPopUpMenu;
                         frmMul_p : TFMul; sChamps_p : string;
                         bSep_p : boolean = false);
var
   sCaption_l, sTmp_l : string;
begin
   MenuInit(MyPopUp_p);
   SubMenuInit(MyPopUp_p, bSep_p);
   while sChamps_p <> '' do
   begin
      sTmp_l := READTOKENST(sChamps_p);
      if sCaption_l <> '' then
         sCaption_l := sCaption_l + ';';
      sCaption_l := sCaption_l + ChoixUnChamp(frmMul_p, sTmp_l);
   end;
   AddItems(MyPopUp_p.Items, sCaption_l, '', OnSubMenuClick);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TMySubMenu.SubMenuInit(MyPopUp_p : TPopUpMenu; bSep_p : boolean);
begin
   iSep_c := 0;
   if bSep_p then
   begin
      AddSep(MyPopUp_p.Items);
      Inc(iSep_c);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMySubMenu.SubMenuAdd(aosRequete_p, aosCaption_p, aosHint_p : array of string;
                                aoAfterMenuClick_p : array of TNotifyEvent);
var
   iInd_l : integer;
begin
   for iInd_l := 0 to Length(aosRequete_p) - 1 do
   begin
      SubMenuAdd(iInd_l, aosRequete_p[iInd_l], aosCaption_p[iInd_l], 
                         aosHint_p[iInd_l], aoAfterMenuClick_p[iInd_l])
   end;
end;
//////////////////////////////////////////////////////////////////
procedure TMySubMenu.SubMenuAdd(iIndice_p : integer;
                                sRequete_p, sCaption_p, sHint_p : string;
                                AfterMenuClick_p : TNotifyEvent);
var
   iInd_l : integer;
begin
   iInd_l := iIndice_p - iSep_c - iFirstInd_c;
   SetLength(AfterMenuClick_c, iInd_l + 1);
   AfterMenuClick_c[iInd_l] := AfterMenuClick_p;
   AddItems(GetPopUp.Items[iIndice_p], sRequete_p, sCaption_p, sHint_p, OnMenuClick);
end;
//////////////////////////////////////////////////////////////////
procedure TMySubMenu.SubMenuAdd(iIndice_p : integer;
                                OBElement_p : TOB; sCaption_p, sHint_p : string;
                                AfterMenuClick_p : TNotifyEvent);
var
   iInd_l, iOBInd_l : integer;
begin
   iInd_l := iIndice_p - iSep_c - iFirstInd_c;
   SetLength(AfterMenuClick_c, iInd_l + 1);
   AfterMenuClick_c[iInd_l] := AfterMenuClick_p;
   AddItems(GetPopUp.Items[iIndice_p], OBElement_p, sCaption_p, sHint_p, OnMenuClick);
end;
//////////////////////////////////////////////////////////////////
procedure TMySubMenu.SubMenuAdd(OBElements_p : TOB; sCaption_p, sHint_p, sMaitre_p : string;
                                AfterMenuClick_p : TNotifyEvent);
var
   iOBInd_l, iSubInd_l : integer;
   sMaitre_l, sCaption_l, sHint_l : string;
begin
   iSubInd_l := -1;
   SetLength(AfterMenuClick_c, OBElements_p.Detail.Count);
   for iOBInd_l := 0 to OBElements_p.Detail.Count -1 do
   begin
      sCaption_l := OBElements_p.Detail[iOBInd_l].GetString(sCaption_p);
      sHint_l := OBElements_p.Detail[iOBInd_l].GetString(sHint_p);
      sMaitre_l := OBElements_p.Detail[iOBInd_l].GetString(sMaitre_p);
      AfterMenuClick_c[iOBInd_l] := AfterMenuClick_p;
      iSubInd_l := IsHint(GetPopUp.Items, sMaitre_l);
      if iSubInd_l = -1 then
         AddItems(GetPopUp.Items, sCaption_l, sHint_l, OnMenuClick)
      else
         AddItems(GetPopUp.Items[iSubInd_l], sCaption_l, sHint_l, OnMenuClick);
   end;
end;
//////////////////////////////////////////////////////////////////
procedure TMySubMenu.SubMenuAdd(iIndice_p : integer;
                                cmbMyTablette_p : THValComboBox;
                                AfterMenuClick_p : TNotifyEvent);
var
   iInd_l : integer;
begin
   iInd_l := iIndice_p - iSep_c - iFirstInd_c;
   SetLength(AfterMenuClick_c, iInd_l + 1);
   AfterMenuClick_c[iInd_l] := AfterMenuClick_p;
   AddItems(GetPopUp.Items[iIndice_p], cmbMyTablette_p, OnMenuClick);
end;

//////////////////////////////////////////////////////////////////
procedure TMySubMenu.SubMenuAdd(iIndice_p : integer;
                     cmbMyMultiTablette_p : THMultiValComboBox;
                     AfterMenuClick_p : TNotifyEvent);
var
   iInd_l : integer;
begin
   iInd_l := iIndice_p - iSep_c - iFirstInd_c;
   SetLength(AfterMenuClick_c, iInd_l + 1);
   AfterMenuClick_c[iInd_l] := AfterMenuClick_p;
   AddItems(GetPopUp.Items[iIndice_p], cmbMyMultiTablette_p, OnMenuClick);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMySubMenu.OnSubMenuClick(Sender : TObject);
begin
   iSousMenu_c := (Sender as TMenuItem).MenuIndex;
   sSubCaption_c := GetPopup.Items[iSousMenu_c].Caption;
   sSubHint_c := GetPopup.Items[iSousMenu_c].Hint;   
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMySubMenu.OnMenuClick(Sender : TObject);
var
   iInd_l : integer;
begin
   inherited;
   iInd_l := iSousMenu_c - iSep_c - iFirstInd_c;
   if (@AfterMenuClick_c[iInd_l] <> nil) then
   begin
      AfterMenuClick_c[iInd_l](Sender);
      sCurHint_c := '';
      iCurInd_c := -1;
      iSousMenu_c := -1;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMySubMenu.OnMenuFiltre(sRestriction_p : string);
begin
   MenuFiltre(iFirstInd_c, sRestriction_p);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
destructor TMySubMenu.Destroy;
begin
   MenuDetruit(iFirstInd_c);
   inherited Destroy;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMySubMenu.GetHint : string;
begin
   result := GetSubHint(iSousMenu_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMySubMenu.IsSubHint(sHint_p : string) : integer;
begin
   result := IsHint(GetPopUp.Items, sHint_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/11/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function TMySubMenu.GetSubCaption : string;
begin
   result := sSubCaption_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMySubMenu.GetSubLeHint : string;
begin
   result := sSubHint_c;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/11/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
function TMySubMenu.GetSubIndice : integer;
begin
   result := iSep_c + iFirstInd_c;
end;

end.
