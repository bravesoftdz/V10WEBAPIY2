{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMEMORISATION ()
Mots clefs ... : TOF;BTMEMORISATION
*****************************************************************}
Unit UtofBTMemorisation ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     forms,
     UtilPgi,
     HMsgBox,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HTB97,
     UTOF,
     Aglinit,
     UTOB ;

Type
  TOF_BTMEMORISATION = Class (TOF)
    BselCode : TToolbarButton97;
    BValide : TToolBarButton97;
    Bannule : TToolBarButton97;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure BSelCodeClick(Sender: Tobject);
    procedure BAnnuleClick(Sender: Tobject);
    procedure BValideClick(Sender: Tobject);
  end ;
const
	TexteMsgCompl: array[1..2] of string 	= (
          {1}        'Vous devez renseigner un Code',
          {2}        'Vous devez renseigner une désignation'
            );

function DemandeNomEnreg (TypeMemo : string; var Code,Libelle : string): boolean;

Implementation

function DemandeNomEnreg (TypeMemo : string; var Code,Libelle : string): boolean;
var TOBMemo : TOB;
begin
result := false;
TOBMemo := TOB.create ('BMEMORISATION',nil,-1);
TOBMEMO.PutValue ('BMO_TYPEMEMO',TypeMemo);
TheTob := TOBMemo;
AGLLanceFiche('BTP','BTMEMORISATION','','','ACTION=CREATION');
if TheTob <> nil then
   begin
   TOBMemo := Thetob;
   Code := TOBMemo.GetValue('BMO_CODEMEMO');
   Libelle := TOBMemo.GetValue('BMO_LIBMEMO');
   TheTob := nil;
   if Code <> '' then result := true;
   end;
TOBMemo.free;
end;

procedure TOF_BTMEMORISATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORISATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORISATION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORISATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMEMORISATION.BSelCodeClick (Sender : Tobject);
var Retour : string;
begin
Retour := AGLLanceFiche ('BTP','BTMEMORIS_MUL','BMO_TYPEMEMO='+LATob.getValue('BMO_TYPEMEMO'),'','ACTION=CONSULTATION');
if Retour <> '' then SetControlText ('BMO_CODEMEMO',Retour);
end;

procedure TOF_BTMEMORISATION.BValideClick (Sender : Tobject);
begin
if GetControlText ('BMO_CODEMEMO') = '' then
   begin
   PGIBox(TexteMsgCompl[1],Ecran.Caption);
   SetFocusControl('BMO_CODEMEMO') ; TForm(Ecran).ModalResult:=0;
   Exit;
   end;
if GetControlText ('BMO_LIBMEMO') = '' then
   begin
   PGIBox(TexteMsgCompl[2],Ecran.Caption);
   SetFocusControl('BMO_LIBMEMO') ; TForm(Ecran).ModalResult:=0;
   Exit;
   end;
LaTob.putvalue ('BMO_CODEMEMO',GetControlText('BMO_CODEMEMO'));
LaTob.putvalue ('BMO_LIBMEMO',GetControlText('BMO_LIBMEMO'));
if laTob.ExistDB then
   begin
   if PGIAsk ('Cette mémorisation existe déjà#10#13Voulez-vous la remplacer ?','Question') <> Mryes then
      begin
      SetFocusControl('BMO_CODEMEMO') ; TForm(Ecran).ModalResult:=0;
      Exit;
      end;
   end;
close;
end;

procedure TOF_BTMEMORISATION.BAnnuleClick (Sender : Tobject);
begin
Latob.putvalue ('BMO_CODEMEMO','');
Latob.putvalue ('BMO_LIBMEMO','');
close;
end;

procedure TOF_BTMEMORISATION.OnArgument (S : String ) ;
begin
  Inherited ;
  BSelCode := TToolBarButton97(ecran.FindComponent('BselCode'));
  BSelCode.OnClick := BSelCodeClick;
  BValide := TToolBarButton97(ecran.FindComponent('BValider'));
  BValide.OnClick := BValideClick;
  BAnnule := TToolBarButton97(ecran.FindComponent('BFerme'));
  BAnnule.OnClick := BAnnuleClick;
end ;

procedure TOF_BTMEMORISATION.OnClose ;
begin
  Inherited ;
  TheTob := Latob;
end ;

Initialization
  registerclasses ( [ TOF_BTMEMORISATION ] ) ; 
end.

