unit UTOFLISTEMAQ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, HTB97, Grids, ExtCtrls,DbTables,Math,
  HEnt1,HPanel,UiUtil, hmsgbox, UTOB,FE_Main, ComCtrls, Menus, HPop97,
  UTOF,HDB,Mul, AssistPL,Paramsoc,
M3FP,DB,Fiche,UTOM;


type
  TOF_LISTEMAQ = Class (TOF)
     procedure OnLoad ; override ;
     private
     function ImportFromListeFichier (stPath, stType : string) : integer;
     procedure ClickMaquette(Sender: TObject);
     function RetourMAq : variant;
  end;

implementation

procedure TOF_LISTEMAQ.Onload;
begin
ImportFromListeFichier ('c:\pgi00\std\', 'CR');
//TListBox(GetControl('LSTMAQ')).OnDblClick := ClickMaquette;
SetControlProperty('BVALIDER','ModalResult',mrOK);
end;

function TOF_LISTEMAQ.ImportFromListeFichier (stPath, stType : string) : integer;
var SearchRec: TSearchRec;
    r_search  : integer;
    NumMaquette : integer;
    St          : string;
begin
  stPath := UpperCase (stPath);
  r_search := FindFirst(stPath+stType+'*.txt', faAnyFile, SearchRec);
  while (r_search = 0) do
  begin
    if SearchRec.Attr <> faDirectory then
    begin
       // SearchRec.Name;
       St := copy (SearchRec.Name, 3, 2);
       NumMaquette := StrToInt (St);
       if NumMaquette <= 20 then
       St := 'Maquette numéro : ' + St + ' type CEGID'
       else
       St := 'Maquette numéro : ' + St + ' type CABINET';
       TListBox(GetControl('LSTMAQ')).Items.Add(St);
    end;
   r_search := FindNext(SearchRec);
   end;
  FindClose(SearchRec);
end;

procedure TOF_LISTEMAQ.ClickMaquette(Sender: TObject);
var
Text      : string;
Index     : integer;
begin
     Index := TListBox(GetControl('LSTMAQ')).ItemIndex;
     if Index >= 0 then
     Text := TListBox(GetControl('LSTMAQ')).Items[Index];
     Text := copy (Text, 19, 2);
     TFMul(Ecran).ModalResult := 1;
end;
function TOF_LISTEMAQ.RetourMAq : variant;
var
  Text      : string;
  Index     : integer;
begin
     Index := TListBox(GetControl('LSTMAQ')).ItemIndex;
     if Index >= 0 then
     result := TListBox(GetControl('LSTMAQ')).Items[Index];
end;

function AGLValiderNumMaq(parms: array of variant; nb: integer ): variant;
var
  F        : TForm;
  Text     : string;
  OM       : TOM;
begin
     F:=TForm(Longint(Parms[0])) ;
     if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
     Text := TOF_LISTEMAQ(OM).RetourMAq;
     TFFiche(F).ModalResult := 1;
     Result := copy (Text, 19, 2);
end;

Initialization
RegisterClasses([TOF_LISTEMAQ]);
RegisterAglFunc( 'ValiderNumMaq', TRUE , 1, AGLValiderNumMaq);
end.
