{***********UNITE*************************************************
Auteur  ...... :  SANTUCCI
Créé le ...... : 12/04/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CHOIXDUPNOMEN ()
Mots clefs ... : TOF;CHOIXDUPNOMEN
*****************************************************************}
Unit ChoixDupNomen ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Htb97,
     HMsgBox,
     UTOF,
     UTOB ;

Type
  TOF_CHOIXDUPNOMEN = Class (TOF)
    TOBParam : TOB;
    G_UNCODE : TGroupBox;
    CODENOMENAUTO : TRadioButton;
    CODEMANUEL : TRadioButton;
    CODENOMEN : THedit;
    Valide : TToolbarButton97;
    Annule : TToolbarButton97;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure CodeNomenAutoClick(Sender: Tobject);
    procedure CodeManuelClick(Sender: Tobject);
    procedure CodeNomenExit(Sender: Tobject);
    procedure ValideClick(Sender: Tobject);
    procedure AnnuleClick(Sender: Tobject);
    function ControleNOmen : boolean;
    procedure AddLesChampsRetour;
  end ;

function ChoixDuplication (TOBRetour:TOB): boolean;

const
	TexteMessage: array[1..3] of string 	= (
          {1}        'La référence de la nomenclature existe déjà'
          {2}        ,'La référence de l''ouvrage existe déjà'
          {3}        ,'Veuillez renseigner le code');

Implementation

function ChoixDuplication (TOBRetour:TOB) : boolean;
begin
result := false;
TheTOB := TOBRetour;
AGLLanceFiche('BTP','BTCHOIXDUPNOM','','','') ;
if TheTob <> nil then
   begin
   if (TOBRetour.GetValue('AUTOMATIQUE')<>'-') or (TOBRetour.GetValue('MANUEL')<>'-') then
      begin
      result := true;
      end;
   end;
TheTob := nil;
end;

procedure TOF_CHOIXDUPNOMEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXDUPNOMEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXDUPNOMEN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXDUPNOMEN.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CHOIXDUPNOMEN.CodeNomenAutoClick (Sender : Tobject);
begin
G_UNCODE.Visible := false;
end;

procedure TOF_CHOIXDUPNOMEN.CodeManuelClick (Sender : Tobject);
begin
G_UNCODE.Visible := true;
SetFocusControl ('CODENOMEN');
end;

procedure TOF_CHOIXDUPNOMEN.CodeNomenExit (Sender : Tobject);
begin
end;

function TOF_CHOIXDUPNOMEN.ControleNOmen : boolean;
var Q : Tquery;
    req : string;
begin
result := false;
req := 'SELECT GNE_NOMENCLATURE FROM NOMENENT WHERE ';
req := req + 'GNE_NOMENCLATURE="'+TobParam.GetValue ('CODENOMEN')+'"';
Q := opensql (req,true,-1,'',true);
if Q.eof then result := true;
ferme (Q);
end;

procedure TOF_CHOIXDUPNOMEN.ValideClick (Sender : Tobject);
begin
if CODENOMENAUTO.Checked then
   BEGIN
   TOBParam.putvalue('AUTOMATIQUE','X');
   TOBParam.putvalue('MANUEL','-');
   TOBParam.Putvalue('CODENOMEN','');
   END ELSE
   BEGIN
   TOBParam.putvalue('AUTOMATIQUE','-');
   TOBParam.putvalue('MANUEL','X');
   TOBParam.Putvalue('CODENOMEN',GetControlText('CODENOMEN') );
   if not ControleNOmen then
      BEGIn
{$IFDEF BTP}
      LastError:=2 ; LastErrorMsg:=TexteMessage[LastError] ;
{$ELSE}
      LastError:=1 ; LastErrorMsg:=TexteMessage[LastError] ;
{$ENDIF}
      PgiInfo (LastErrorMsg, Ecran.Caption);
      TForm(Ecran).ModalResult:=0;
      SetFocusControl ('CODENOMEN');
      exit;
      END;
   if TOBParam.Getvalue('CODENOMEN') = '' then
      Begin
      LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ;
      PgiInfo (LastErrorMsg, Ecran.Caption);
      TForm(Ecran).ModalResult:=0;
      SetFocusControl ('CODENOMEN');
      exit;
      End;
   END;
TheTob := TOBPARAM;
close;
end;

procedure TOF_CHOIXDUPNOMEN.AnnuleClick (Sender : Tobject);
begin
TOBParam.putvalue('AUTOMATIQUE','-');
TOBParam.putvalue('MANUEL','-');
TOBParam.Putvalue('CODENOMEN','');
TheTob := TOBPARAM;
close;
end;

procedure TOF_CHOIXDUPNOMEN.AddLesChampsRetour;
begin
TOBParaM.addchampsupValeur ('AUTOMATIQUE','X',false);
TOBParaM.addchampsupValeur ('MANUEL','-',false);
TOBParaM.addchampsupValeur ('CODENOMEN','',false);
end;

procedure TOF_CHOIXDUPNOMEN.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBParam:= Latob;
  AddLesChampsRetour;
  G_UNCODE := TGroupBox (GetControl('G_UNCODE'));
{$IFDEF BTP}
  G_UNCODE.Caption := TraduireMemoire('Code Ouvrage');
  THRadioGroup(GetControl('G_LECHOIX')).Caption := TraduireMemoire('Mode d''affectation du code ouvrage');
{$ENDIF}
  CODENOMENAUTO := TRadioButton (GetControl('CODENOMENAUTO'));
  CODENOMENAUTO.OnClick := CodeNomenAutoClick;

  CODEMANUEL := TRadioButton (GetControl('CODEMANUEL'));
  CODEMANUEL.OnClick := CodeManuelClick;

  CODENOMEN := THEdit(GEtCOntrol('CODENOMEN'));
  CODENOMEN.OnExit := CodeNomenExit;

  Valide:= TToolbarButton97(GetControl('BValider'));
  Valide.OnClick := ValideClick;
  Annule := TToolbarButton97 (GetControl('BFerme'));
  Annule.OnClick := AnnuleClick;
end ;

procedure TOF_CHOIXDUPNOMEN.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CHOIXDUPNOMEN ] ) ;
end.

