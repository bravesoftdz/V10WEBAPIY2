{***********UNITE*************************************************
Auteur  ...... : SANTUCCI Lionel
Créé le ...... : 18/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CHANGECODEART (Changement de codification des articles)
Mots clefs ... : TOF;CHANGECODEART
*****************************************************************}
Unit UtofBTChangeCodeArt ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,Fe_Main,
{$ENDIF}
     AglInit,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HTB97,
     UTOB,
     ParamSoc,
     EntGC;

Type
  TOF_CHANGECODEART = Class (TOF)
  private
    AncCode, NewCode : THEdit;
    Bvalide : TToolbarButton97 ;
    procedure BValideClick(sender: Tobject);
    procedure AncCodeExit(sender: Tobject);
    procedure NewCodeExit(sender: Tobject);
    Function ControleExiste : boolean;
//    Function ControleNonExiste : boolean;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

function SaisieNewCodeArt (prefixe,TypeArticle,CodeArticle,Libelle : string ; var NewCode : string): boolean;
function SaisieNewCodeTar (var AncCode, NewCode, TraiteCodeArt : string): boolean;

Implementation

function SaisieNewCodeTar (var AncCode, NewCode, TraiteCodeArt : string): boolean;
var TOBParam : TOB;
begin
result := false;
TOBParam := TOB.Create ('LATTTOB',nil,-1);
TOBParam.AddChampSupValeur ('TYPEARTICLE','TAR');
TOBParam.AddChampSupValeur ('ANCCODE','');
TOBParam.AddChampSupValeur ('NEWCODE','');
TOBParam.AddChampSupValeur ('TRAITECODEART','-');
TOBParam.AddChampSupValeur ('RESULT','-');
TheTob := TOBParam;
AGLLanceFiche('BTP','BTCHANGECODEART','','','ACTION=CREATION') ;
if TheTob <> nil then
   begin
   if TheTob.GetValue('RESULT') = 'X' then
      begin
      AncCode := TheTOB.GetValue('ANCCODE');
      NewCode := TheTOB.GetValue('NEWCODE');
      TraiteCodeArt := TheTOB.GetValue('TRAITECODEART');
      Result := True;
      end;
   end;
TOBParam.free;
TheTob := nil;
end;

function SaisieNewCodeArt (prefixe,TypeArticle,CodeArticle,Libelle : string ; var NewCode : string): boolean;
var TOBParam : TOB;
    lng : integer;
begin
result := false;
TOBParam := TOB.Create ('LATTTOB',nil,-1);
TOBParam.AddChampSupValeur ('PREFIXE',prefixe);
TOBParam.AddChampSupValeur ('TYPEARTICLE',TypeArticle);
TOBParam.AddChampSupValeur ('ANCCODE',CodeArticle);
TOBParam.AddChampSupValeur ('LIBELLE',Libelle);
TOBParam.AddChampSupValeur ('NEWCODE','');
TOBParam.AddChampSupValeur ('RESULT','-');
TheTob := TOBParam;
AGLLanceFiche('BTP','BTCHANGECODEART','','','ACTION=MODIFICATION') ;
if TheTob <> nil then
   begin
   if TheTob.GetValue('RESULT') = 'X' then
      begin
      NewCode := TheTOB.GetValue('NEWCODE');
      if (length (prefixe) > 0) and (length(newcode) > length(prefixe)) and (copy (NewCode,1,length(prefixe)) <> prefixe) then
         begin
         lng := GetParamsoc('SO_GCLGNUMART');
         newCode := copy (Prefixe+newCode,1,lng);
         end;
      if (NewCode <> '') and (length(newcode) > length (prefixe)) then result := true;
      end;
   end;
TOBParam.free;
TheTob := nil;
end;

procedure TOF_CHANGECODEART.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CHANGECODEART.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CHANGECODEART.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CHANGECODEART.OnLoad ;
var Libelle : string;
begin
  Inherited ;
  if LaTob.getValue('TYPEARTICLE') <> 'TAR' Then
  begin
    Libelle := RechDom ('GCTYPEARTICLE',laTob.getValue('TYPEARTICLE'),false);
    if LaTob.getValue('TYPEARTICLE') = 'MAR' Then Libelle := 'Article ';
    ecran.Caption := Libelle + LaTOB.GetValue('LIBELLE');
    SetControlCaption ('ANCCODEARTICLE',laTob.GetValue('ANCCODE'));
    if GetParamsoc('SO_GCNUMARTAUTO') then SetControlText ('NEWCODE',LaTOB.GetValue('PREFIXE'));
  end;
end ;

procedure TOF_CHANGECODEART.OnArgument (S : String ) ;
begin
  Inherited ;
  Bvalide := TToolbarButton97 (GetControl('BValider'));
  BValide.OnClick := BValideClick;
  if LaTob.getValue('TYPEARTICLE') = 'TAR' Then
  begin
    ecran.Caption := 'Changement Code Tarif';
    SetControlVisible('ANCCODEARTICLE', False);
    SetControlVisible('ANCCODETARIF', True);
    SetControlVisible('TRAITECODEART', True);
    AncCode := THEdit(GetControl('ANCCODETARIF'));
    AncCode.OnExit  := AncCodeExit;
    NewCode := THEdit(GetControl('NEWCODE'));
    NewCode.OnExit  := NewCodeExit;
  end;
end ;

procedure TOF_CHANGECODEART.BValideClick (sender : Tobject);
begin
if LaTob.getValue('TYPEARTICLE') = 'TAR' Then
begin
  laTob.putValue('ANCCODE', THEdit(GetControl('ANCCODETARIF')).Text );
  if THCheckBox(GetControl('TRAITECODEART')).Checked then LaTob.PutValue('TRAITECODEART','X');
  if Not ControleExiste then begin TForm(Ecran).ModalResult:=0; Exit; end;
// inutile car le nouveau code peut déjà exister : if Not ControleNonExiste then begin TForm(Ecran).ModalResult:=0; Exit; end;
end;
laTob.putValue('NEWCODE', THEdit(GetControl('NEWCODE')).Text );
LaTob.PutValue('RESULT','X');
if (LaTOB <> Nil) then TheTob := LaTOB ;
end;

procedure TOF_CHANGECODEART.AncCodeExit (sender : Tobject);
begin
  if Not ControleExiste then Exit;
end;

procedure TOF_CHANGECODEART.NewCodeExit (sender : Tobject);
begin
// inutile car le nouveau code peut déjà exister : if Not ControleNonExiste then Exit;
end;

Function TOF_CHANGECODEART.ControleExiste : boolean ;
var Res : string;
begin
  Result := True;
  Res:=RechDom ('GCTARIFARTICLE',THEdit(GetControl('ANCCODETARIF')).Text,False);
  if res = '' then
  begin
    PGIBox ('Code tarif inexistant');
    SetFocusControl('ANCCODETARIF');
    Result := False;
  end
end;

{*Function TOF_CHANGECODEART.ControleNonExiste : boolean;
var Res : string;
begin
  Result := True;
  Res:=RechDom ('GCTARIFARTICLE',THEdit(GetControl('NEWCODE')).Text,False);
  if res <> '' then
  begin
    PGIBox ('Code tarif déjà existant');
    SetFocusControl('NEWCODE');
    Result := False;
  end;
end;}

procedure TOF_CHANGECODEART.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CHANGECODEART ] ) ;
end.

