{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 27/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : LISTEINVPRECON ()
Mots clefs ... : TOF;LISTEINVPRECON
*****************************************************************}
Unit ListeInvPreCon_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db, dbtables,FE_Main,
{$ENDIF}
     forms,Graphics,sysutils,ComCtrls,
     HCtrls,HEnt1,HMsgBox,UTOF ;

function GCLanceFiche_ListeInvPreCon(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_LISTEINVPRECON = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    procedure OnComboChange(Sender : TObject);
    procedure OnComboPopUp(Sender : TObject);
  end ;

const
    NbRupture = 4;

Implementation

function GCLanceFiche_ListeInvPreCon(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_LISTEINVPRECON.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_LISTEINVPRECON.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_LISTEINVPRECON.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_LISTEINVPRECON.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_LISTEINVPRECON.OnArgument (S : String ) ;
var iInd : integer;
begin
  Inherited ;
for iInd := 1 to NbRupture do
  begin
  THValComboBox(GetControl('RUPT'+inttostr(iInd))).OnChange := OnComboChange;
  THValComboBox(GetControl('RUPT'+inttostr(iInd))).OnDropDown := OnComboPopup;
  end;
THValcomboBox(GetControl('GQC_DEPOT')).ItemIndex := 0;
end ;

procedure TOF_LISTEINVPRECON.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_LISTEINVPRECON.OnComboChange(Sender : TObject);
var RuptNb,ComboName,NextComboName,ValCombo : string;
begin
ComboName := THValComboBox(Sender).Name;
RuptNb := Copy(ComboName,5,1);
ValCombo := GetControlText(ComboName);
if ValCombo = '' then
  begin
  SetControlText('XX_RUPTURE'+RuptNb,'');
  SetControlText('XX_VARIABLE'+RuptNb,'');
  end else
  begin
  SetControlText('XX_RUPTURE'+RuptNb,RechDom('GCGROUPINVDIS',ValCombo,true));
  SetControlText('XX_VARIABLE'+RuptNb,RechDom('GCGROUPINVDIS',ValCombo,false));
  end;

NextComboName := ComboName; NextComboName[5] := Chr(Ord(NextComboName[5])+1);
if (GetControl(NextComboName) <> nil) and (ValCombo <> '') then
  begin
  SetControlEnabled(NextComboName,true);
  SetControlProperty(NextComboName,'Color',clWindow);
  end else
  while GetControl(NextComboName) <> nil do
    begin
    SetControlEnabled(NextComboName,true);
    SetControlProperty(NextComboName,'Color',clBtnFace);
    SetControlText(NextComboName,'');
    NextComboName[5] := Chr(Ord(NextComboName[5])+1);
    end;
end;

procedure TOF_LISTEINVPRECON.OnComboPopUp(Sender : TObject);
var St_Plus, St_Value, St_Text, St : String;
    FCombo : THValComboBox;
    Indice, i_ind : Integer;
begin
FCombo := (Sender as THValComboBox);
Indice := StrToInt(Copy(FCombo.Name,5,1));
St_Plus := '';
St_Value := FCombo.Value;
St_Text := FCombo.Text;
for i_ind := 1 to NbRupture do
  begin
  if i_ind = Indice then Continue;
  St := String(THValComboBox(GetControl('RUPT'+IntToStr(i_ind))).Value);
  if St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"'+St+'"';
  end;
FCombo.Plus := St_Plus;
FCombo.Value := St_Value;
FCombo.Text := St_Text;
end;


Initialization
  registerclasses ( [ TOF_LISTEINVPRECON ] ) ; 
end.

