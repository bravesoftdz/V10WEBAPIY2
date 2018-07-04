{***********UNITE*************************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/05/2004
Modifié le ... : 08/10/2004
Description .. : 
Mots clefs ... : 
*****************************************************************}
unit UJurOutilsMvtRem;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   {$IFDEF EAGLCLIENT}
   {$ELSE}
   DB,
   {$ENDIF}
   sysutils, HCTRLS, stdctrls, UTOB, controls;

/////////////////////////////////////////////////////////////////
type
   TMyTablette = class

      constructor Create(sTablette_p : string); reintroduce; overload;
      destructor Destroy(bAll_p : boolean); reintroduce; overload;

  protected
      procedure ChargeTablette;
      function  IsPaire(iValeur_p : integer) : boolean;
      function  IsChecked(sCode_p, sValeurs_p : string) : boolean;

   private
      sTablette_c, sCode_c, sChampLib_c : string;
      OBValeurs_c : TOB;
   end;
/////////////////////////////////////////////////////////////////
type
   TMyCheckGroup = class(TMyTablette)

      constructor Create(MyGroupBox_p : TGroupBox;
                         sTablette_p, sField_p : string;
                         iNbCol_p : integer = 1;
                         bEnabled_p : boolean = true;
                         bVisible_p : boolean = true); reintroduce; overload;

      destructor Destroy(bAll_p : boolean); reintroduce; overload;

      procedure OnLoadFromDS(DS_p : TDataSet);
      procedure OnLoadFromTOB(OB_p : TOB);
      procedure RazValeurs;

   private
      MyGroupBox_c : TGroupBox;
      DS_c : TDataSet;
      OB_c : TOB;
      sField_c : string;
      bEnabled_c, bVisible_c, bCharge_c : boolean;
      iNbCol_c : integer;
      aocbCheck_c : array of TCheckBox;

      procedure CreerCheckBox;
      procedure ChargeValeurs(sValeurs_p : string);
      procedure OnClick_CB(Sender : TObject);
   end;
/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TMyTablette.Create(sTablette_p : string);
begin
   sTablette_c  := sTablette_p;
   ChargeTablette;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TMyTablette.Destroy(bAll_p : boolean);
begin
   OBValeurs_c.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyTablette.ChargeTablette;
var
   OBTablette_l : TOB;
   sReq_l, sTable_l, sWhere_l : string;
begin
   OBTablette_l := TOB.Create('DECOMBOS', nil, -1);
   OBTablette_l.LoadDetailDBFromSQL('DECOMBOS', 'SELECT * FROM DECOMBOS ' +
                                    'WHERE DO_COMBO = "' + sTablette_c + '"');

   if OBTablette_l.Detail.Count > 0 then
   begin
      sTable_l    := PrefixeToTable(OBTablette_l.Detail[0].GetValue('DO_PREFIXE'));
      sCode_c     := OBTablette_l.Detail[0].GetValue('DO_CODE');
      sChampLib_c := OBTablette_l.Detail[0].GetValue('DO_CHAMPLIB');
      sWhere_l    := OBTablette_l.Detail[0].GetValue('DO_WHERE');
      sWhere_l    := StringReplace(sWhere_l, '&#@', '', [rfReplaceAll, rfIgnoreCase]);
   end;
   OBTablette_l.Free;

   OBValeurs_c := TOB.Create(sTable_l, nil, -1);
   sReq_l := 'SELECT ' + sCode_c + ', ' + sChampLib_c + ' FROM ' + sTable_l + ' WHERE ' + sWhere_l;
   OBValeurs_c.LoadDetailDBFromSQL(sTable_l, sReq_l);
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TMyTablette.IsPaire(iValeur_p : integer) : boolean;
begin
   if iValeur_p = 0 then
      result := true
   else if (iValeur_p / 2 = iValeur_p div 2) then
      result := true
   else
      result := false;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TMyTablette.IsChecked(sCode_p, sValeurs_p : string) : boolean;
var
   bOK_l : boolean;
   sCode_l : string;
begin
   bOK_l := false;
   while not bOK_l and (sValeurs_p <> '') do
   begin
      sCode_l := READTOKENST(sValeurs_p);
      bOK_l := (sCode_l = sCode_p);
   end;
   result := bOK_l;
end;

/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TMyCheckGroup.Create(MyGroupBox_p : TGroupBox;
                                sTablette_p, sField_p : string;
                                iNbCol_p : integer = 1;
                                bEnabled_p : boolean = true;
                                bVisible_p : boolean = true);
begin
   bCharge_c := false;
   inherited Create(sTablette_p);

   MyGroupBox_c := MyGroupBox_p;
   sField_c     := sField_p;
   iNbCol_c     := iNbCol_p;
   bEnabled_c   := bEnabled_p;
   bVisible_c   := bVisible_p;

   CreerCheckBox;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyCheckGroup.OnLoadFromDS(DS_p : TDataSet);
begin
   DS_c := DS_p;
   {$IFDEF EAGLCLIENT}
   ChargeValeurs(DS_c.FindField(sField_c).AsString);
   {$ELSE}
   ChargeValeurs(DS_c.FieldValues[sField_c]);
   {$ENDIF}
   bCharge_c := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyCheckGroup.OnLoadFromTOB(OB_p : TOB);
begin
   OB_c := OB_p;
   ChargeValeurs(OB_c.GetString(sField_c));
   bCharge_c := true;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TMyCheckGroup.Destroy(bAll_p : boolean);
var
   iInd_l : integer;
begin
   inherited; // Destroy(true);
   For iInd_l := 0 to Length(aocbCheck_c) - 1 do
      aocbCheck_c[iInd_l].Free;
   Setlength(aocbCheck_c, 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyCheckGroup.CreerCheckBox;
var
   iInd_l, iTop_l, iLeft_l, ciWidth_l, ciNb_l, ciEsp_l, iCptCol_l : integer;
begin
   if iNbCol_c > 1 then
      ciWidth_l := Round((MyGroupBox_c.Width - (10 + iNbCol_c))/ iNbCol_c);

{   if (iNbCol_c = 1) then
   begin
      ciNb_l  := OBValeurs_c.Detail.Count;
      ciEsp_l := MyGroupBox_c.Height - (ciNb_l * 17) - 15;
      ciEsp_l := Round(ciEsp_l / (ciNb_l - 1));
   end
   else if IsPaire(OBValeurs_c.Detail.Count) then
   begin
      ciNb_l  := Round(OBValeurs_c.Detail.Count / iNbCol_c);
      ciEsp_l := MyGroupBox_c.Height - (ciNb_l * 17) - 15;
      ciEsp_l := Round(ciEsp_l / (ciNb_l - 1));
   end
   else}
   begin
      ciNb_l  := Round((OBValeurs_c.Detail.Count + 1 ) / iNbCol_c);
      ciEsp_l := MyGroupBox_c.Height - 20;// - (ciNb_l * 17);
      ciEsp_l := Round(ciEsp_l / ciNb_l);
   end;
//   ciEsp_l := ciEsp_l + 17;
//   ciEsp_l := 26;

   SetLength(aocbCheck_c, OBValeurs_c.Detail.Count);
   iTop_l := 18;
   iLeft_l := 10;
   iCptCol_l := 1;

   For iInd_l := 0 to OBValeurs_c.Detail.Count - 1 do
   begin
      aocbCheck_c[iInd_l] := TCheckBox.Create(TWinControl(MyGroupBox_c));
      aocbCheck_c[iInd_l].Caption := OBValeurs_c.Detail[iInd_l].GetValue(sChampLib_c);
      aocbCheck_c[iInd_l].Width   := ciWidth_l;
      aocbCheck_c[iInd_l].Enabled := bEnabled_c;
      aocbCheck_c[iInd_l].Hint    := OBValeurs_c.Detail[iInd_l].GetValue(sCode_c);
      aocbCheck_c[iInd_l].Left    := iLeft_l;
      aocbCheck_c[iInd_l].Name    := sField_c + IntToStr(iInd_l + 1);
      aocbCheck_c[iInd_l].Parent  := MyGroupBox_c;
      aocbCheck_c[iInd_l].ParentColor := false;
      aocbCheck_c[iInd_l].Top     := iTop_l;
      aocbCheck_c[iInd_l].Visible := bVisible_c;
      aocbCheck_c[iInd_l].OnClick := OnClick_CB;

      if iCptCol_l = iNbCol_c then
      begin
         iLeft_l := 10;
         iTop_l := iTop_l + ciEsp_l;
         iCptCol_l := 1;
      end
      else
      begin
         iLeft_l := iLeft_l + ciWidth_l;
         Inc(iCptCol_l);
      end;

   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyCheckGroup.ChargeValeurs(sValeurs_p : string);
var
   iInd_l: integer;
begin
   For iInd_l := 0 to Length(aocbCheck_c) - 1 do
   begin
      aocbCheck_c[iInd_l].Checked := IsChecked(aocbCheck_c[iInd_l].Hint, sValeurs_p);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TMyCheckGroup.RazValeurs;
var
   iInd_l: integer;
begin
   For iInd_l := 0 to Length(aocbCheck_c) - 1 do
   begin
      aocbCheck_c[iInd_l].Checked := false;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TMyCheckGroup.OnClick_CB(Sender : TObject);
var
   iInd_l: integer;
   sValeurs_l : string;
begin
   if not bCharge_c then exit;
   if DS_c = nil then exit;
   For iInd_l := 0 to Length(aocbCheck_c) - 1 do
   begin
      if aocbCheck_c[iInd_l].Checked then
         sValeurs_l := sValeurs_l + aocbCheck_c[iInd_l].Hint + ';';
   end;
   {$IFDEF EAGLCLIENT}
   if DS_c.FindField(sField_c).AsString <> sValeurs_l then
   {$ELSE}
   if DS_c.FieldValues[sField_c] <> sValeurs_l then
   {$ENDIF}
   begin
      DS_c.Edit;
      {$IFDEF EAGLCLIENT}
      DS_c.FindField(sField_c).AsString := sValeurs_l;
      {$ELSE}
      DS_c.FieldValues[sField_c] := sValeurs_l;
      {$ENDIF}
   end;
end;
/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
end.
