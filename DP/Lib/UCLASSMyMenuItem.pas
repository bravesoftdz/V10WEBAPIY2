{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 04/09/2002
Modifié le ... : 04/11/2004
Description .. :
Mots clefs ... : 
*****************************************************************}
unit UCLASSMyMenuItem;

interface

uses
   {$IFDEF EAGLCLIENT}
   emul,
   {$ELSE}
   mul,
   {$ENDIF}
   Classes,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}
   Menus, Hctrls, UTOB, galoutil, SysUtils;

type
   TMyMenuItem = class(TMenuItem)

   protected
      iCurInd_c : integer;
      sCurHint_c : string;
      iFirstInd_c : integer;

      function GetPopUp : TPopUpMenu;
      function GetSubHint(iInd_p : integer) : string;
      function IsHint(MyMenuItem_p : TMenuItem; sHint_p : string) : integer;

      function  ChoixUnChamp(frmMul_p : TFMul; sChamp_p : string) : string;
      
      procedure OnMenuClick(Sender : TObject); virtual;

      procedure AddItems(MyMenuItem_p : TMenuItem;
                         cmbMyTablette_p : THValComboBox;
                         MenuClick_p : TNotifyEvent = nil); reintroduce; overload;

      procedure AddItems(MyMenuItem_p : TMenuItem;
                         cmbMyMultiTablette_p : THMultiValComboBox;
                         MenuClick_p : TNotifyEvent = nil); reintroduce; overload;

      procedure AddItems(MyMenuItem_p : TMenuItem;
                         OBElements_p : TOB; sCaption_p, sHint_p : string;
                         MenuClick_p : TNotifyEvent = nil); reintroduce; overload;

      procedure AddItems(MyMenuItem_p : TMenuItem;
                         sRequete_p, sCaption_p, sHint_p : string;
                         MenuClick_p : TNotifyEvent = nil); reintroduce; overload;

      procedure AddItems(MyMenuItem_p : TMenuItem;
                         sCaption_p, sHint_p : string;
                         MenuClick_p : TNotifyEvent = nil{; sListeChecked_p : string = ''}); reintroduce; overload;

      procedure AddSep(MyMenuItem_p : TMenuItem); 

      procedure MenuInit(MyPopUp_p : TPopUpMenu);
      procedure MenuFiltre(iToInd_p : integer; sRestriction_p : string);
      procedure MenuDetruit(iToInd_p : integer);

   private
      MyPopUp_c : TPopUpMenu;

      function  AddItem(sCaption_p : string;
                        sHint_p : string = '';
                        neOnClick_p : TNotifyEvent = nil;
                        bChecked_p : boolean = false) : TMenuItem;
      function  FiltreItem(sValeur_p, sListe_p : string) : boolean;

      procedure FreeItem(MyMenuItem_p : TMenuItem; iToInd_p : integer);

   end;


implementation


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyMenuItem.MenuInit(MyPopUp_p : TPopUpMenu);
begin
   MyPopUp_c := MyPopUp_p;
   sCurHint_c := '';
   iCurInd_c := -1;
   iFirstInd_c := MyPopUp_c.Items.Count;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyMenuItem.AddItems(MyMenuItem_p : TMenuItem;
                               cmbMyTablette_p : THValComboBox;
                               MenuClick_p : TNotifyEvent = nil);
var
   iItem_l : integer;
begin
   for iItem_l := 0 to cmbMyTablette_p.Items.count - 1 do
   begin
      if (cmbMyTablette_p.Items[iItem_l] <> '<<Tous>>') and
         (cmbMyTablette_p.Values[iItem_l] <> '') then
         MyMenuItem_p.Add(AddItem(cmbMyTablette_p.Items[iItem_l],
                                  cmbMyTablette_p.Values[iItem_l],
                                  MenuClick_p));
   end;
end;
//////////////////////////////////////////////////////////////////
procedure TMyMenuItem.AddItems(MyMenuItem_p : TMenuItem;
                               cmbMyMultiTablette_p : THMultiValComboBox;
                               MenuClick_p : TNotifyEvent = nil);
var
   iItem_l : integer;
begin
   for iItem_l := 0 to cmbMyMultiTablette_p.Items.count - 1 do
   begin
      if (cmbMyMultiTablette_p.Items[iItem_l] <> '<<Tous>>') and
         (cmbMyMultiTablette_p.Values[iItem_l] <> '') then
         MyMenuItem_p.Add(AddItem(cmbMyMultiTablette_p.Items[iItem_l],
                                  cmbMyMultiTablette_p.Values[iItem_l],
                                  MenuClick_p));
   end;
end;
//////////////////////////////////////////////////////////////////
procedure TMyMenuItem.AddItems(MyMenuItem_p : TMenuItem;
                               OBElements_p : TOB; sCaption_p, sHint_p : string;
                               MenuClick_p : TNotifyEvent = nil);
var
   iItem_l : integer;
begin
   for iItem_l := 0 to OBElements_p.Detail.count-1 do
   begin
      MyMenuItem_p.Add(AddItem(VarToStr(OBElements_p.Detail[iItem_l].GetValue(sCaption_p)),
                               VarToStr(OBElements_p.Detail[iItem_l].GetValue(sHint_p)),
                               MenuClick_p));
   end;
end;
//////////////////////////////////////////////////////////////////
procedure TMyMenuItem.AddItems(MyMenuItem_p : TMenuItem;
                               sRequete_p, sCaption_p, sHint_p : string;
                               MenuClick_p : TNotifyEvent = nil);
var
   OBElements_l : TOB;
   iItem_l : integer;
begin
   OBElements_l := TOB.Create('', nil, -1);
   OBElements_l.LoadDetailFromSQL(sRequete_p);
   for iItem_l := 0 to OBElements_l.Detail.count-1 do
   begin
      MyMenuItem_p.Add(AddItem(VarToStr(OBElements_l.Detail[iItem_l].GetValue(sCaption_p)),
                               VarToStr(OBElements_l.Detail[iItem_l].GetValue(sHint_p)),
                               MenuClick_p));
   end;

   OBElements_l.Free;
end;
//////////////////////////////////////////////////////////////////
procedure TMyMenuItem.AddItems(MyMenuItem_p : TMenuItem;
                               sCaption_p, sHint_p : string;
                               MenuClick_p : TNotifyEvent = nil);{; sListeChecked_p : string = ''}
var
   sCaption_l, sHint_l{, sChecked_l} : string;
   bChecked_l : boolean;
begin
   while sCaption_p <> '' do
   begin
      sCaption_l := ReadTokenSt(sCaption_p);
      sHint_l := ReadTokenSt(sHint_p);
      bChecked_l := false;

//      sChecked_l := ReadTokenSt(sListeChecked_p);
//      if (sChecked_l <> '') and (sChecked_l[1] in ['1','t','T']) then
//         bChecked_l := true;

      MyMenuItem_p.Add(AddItem(sCaption_l, sHint_l, MenuClick_p, bChecked_l));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyMenuItem.AddSep(MyMenuItem_p : TMenuItem);
begin
   MyMenuItem_p.Add(AddItem('-'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenuItem.AddItem(sCaption_p : string; sHint_p : string = '';
                             neOnClick_p : TNotifyEvent = nil;
                             bChecked_p : boolean = false ): TMenuItem;
var
   MyItem_l: TMenuItem;
begin
   MyItem_l := TMenuItem.Create(Self);
   MyItem_l.Caption := sCaption_p;
   MyItem_l.Hint := sHint_p;
   MyItem_l.OnClick := neOnClick_p;
   MyItem_l.Checked := bChecked_p;
   result := MyItem_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenuItem.FiltreItem(sValeur_p, sListe_p : string) : boolean;
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

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenuItem.GetSubHint(iInd_p : integer) : string;
begin
   result := '';
   if iInd_p >= MyPopUp_c.Items.Count then exit;
   result := MyPopUp_c.Items[iInd_p].Items[iCurInd_c].Hint;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenuItem.IsHint(MyMenuItem_p : TMenuItem; sHint_p : string) : integer;
var
   iInd_l : integer;
   bOK_l : boolean;
begin
   bOK_l := false;
   iInd_l := 0;
   while not BOK_l and (iInd_l < MyMenuItem_p.Count) do
   begin
      bOK_l := MyMenuItem_p[iInd_l].Hint = sHint_p;
      Inc(iInd_l)
   end;
   Dec(iInd_l);
   result := iInd_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{function TMyMenuItem.GetIndex : integer;
begin
   result := iCurInd_c;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{function TMyMenuItem.GetFirst : integer;
begin
   result := iFirstInd_c;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenuItem.GetPopUp : TPopUpMenu;
begin
   result := MyPopUp_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TMyMenuItem.OnMenuClick(Sender : TObject);
begin
   sCurHint_c := (Sender as TMenuItem).Hint;
   iCurInd_c := (Sender as TMenuItem).MenuIndex;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyMenuItem.MenuFiltre(iToInd_p : integer; sRestriction_p : string);
var
   iPopInd_l : integer;
   sElement_l : string;
begin
   for iPopInd_l := MyPopUp_c.Items.Count - 1 downto iToInd_p do
   begin
      sElement_l := MyPopUp_c.Items[iPopInd_l].Hint;
      if not FiltreItem(sElement_l, sRestriction_p) then
         MyPopUp_c.Items[iPopInd_l].Free;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 15/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyMenuItem.ChoixUnChamp(frmMul_p : TFMul; sChamp_p : string) : string;
var
   sLesTitres_l, sLesChamps_l, sTitre_l, sChamp_l : string;
   nIDField_l, nPos_l : integer;
begin
   sTitre_l := '';

   if not ChampEstDansQuery( sChamp_p, frmMul_p.Q) then
   begin
      sTitre_l := ChampToLibelle(sChamp_p);
      nPos_l := Pos ( ':', sTitre_l );
      if nPos_l > 0 then
         sTitre_l := Copy ( sTitre_l, 1, nPos_l - 1 ) ;
   end
   else
   begin
      sLesChamps_l := frmMul_p.q.Champs;
		sLesTitres_l := frmMul_p.q.Titres;

      while (sLesChamps_l <> '') and (sChamp_l <> sChamp_p) do
      begin
      	sChamp_l := Trim(READTOKENPipe(sLesChamps_l, ','));
      	sTitre_l := Trim(READTOKENPipe(sLesTitres_l, ';'));
      end;

      if sChamp_l <> sChamp_p then
      	sTitre_l := '';

      {$IFDEF EAGLCLIENT}
//      nIDField_l := frmMul_p.Q.TQ.GetNumChamp(sChamp_p) - 1;
//      sTitre_l := frmMul_p.FListe.CellValues[0, nIDField_l];
      {$ELSE}
//      nIDField_l := frmMul_p.Q.gFieldByName(sChamp_p).Index - 1;
//      sTitre_l := frmMul_p.FListe.Columns.Items[nIDField_l].Title.Caption;
      {$ENDIF}
   end;
   result := sTitre_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 04/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyMenuItem.MenuDetruit(iToInd_p : integer);
var
   iPopInd_l : integer;
begin
   if MyPopUp_c = nil then exit;
   for iPopInd_l := MyPopUp_c.Items.Count - 1 downto iToInd_p do
   begin
      FreeItem(MyPopUp_c.Items[iPopInd_l], 0);
      MyPopUp_c.Items[iPopInd_l].Free;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 10/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyMenuItem.FreeItem(MyMenuItem_p : TMenuItem; iToInd_p : integer);
var
   iInd_l : integer;
begin
   for iInd_l := MyMenuItem_p.Count - 1 downto iToInd_p do
   begin
      MyMenuItem_p.Items[iInd_l].Free;
   end;
end;

end.
