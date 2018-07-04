{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/03/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEAVANC ()
Mots clefs ... : TOF;SAISIEAVANC
*****************************************************************}
Unit UTofSaisieAvanc ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
{$ELSE}
     MainEAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AglInit,
     UTOF,UTOB,HTB97 ;

Type
  TOF_SAISIEAVANC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure BokClick(Sender: Tobject);
    procedure BAnnulClick(Sender: Tobject);
  public
    Bok : TToolbarButton97;
    BAnnul : TToolbarButton97;
    ModeApplic : boolean;
    pourcent : double;
    ParagOk : boolean;
    SaisieValide : boolean;
  end ;

function AssistSaisieAvanc (var ModeApplic:integer;var PourcentAvanc:double;ParagOk : boolean;TousOk : boolean=false):boolean;
Implementation

procedure CreeLesChampsSup (TOBI : TOB);
begin
TOBI.AddChampSup ('APPLIC',false);
TOBI.AddChampSup ('POURCENT',false);
TOBI.AddChampSup ('SAISOK',false);
TOBI.AddChampSup ('PARAG',false);
TOBI.AddChampSup ('TOUS',false);
end;

function AssistSaisieAvanc (var ModeApplic:integer;var PourcentAvanc:double;ParagOk : boolean;TousOk : boolean=false):boolean;
var TOBINFO : TOB;
begin
result := false;
TOBINFO := TOB.create ('',nil,-1);

CreeLesChampsSup (TOBINfo);
TOBInfo.putvalue ('APPLIC',ModeApplic);
TOBInfo.putvalue ('SAISOK','-');
TOBInfo.putvalue ('POURCENT',0);
if ParagOk then TOBInfo.putvalue ('PARAG','X')
           else TOBInfo.putvalue ('PARAG','-');
if TousOk then TOBInfo.putvalue ('TOUS','X')
          else TOBInfo.putvalue ('TOUS','-');
TheTob := TOBInfo;
TRY
  AGLLanceFiche('BTP','BTSAISAVANC','','','ACTION=CREATION') ;
  if TheTOB = nil then Exit;
  if TOBInfo.GetValue ('SAISOK') = 'X' then
   begin
    ModeApplic := TOBInfo.GetValue('APPLIC');
    PourcentAvanc := TOBInfo.GetValue ('POURCENT');
   result := true;
   end;
finally
  TOBInfo.free;
  TheTob := nil;
end;
end;

procedure TOF_SAISIEAVANC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEAVANC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEAVANC.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEAVANC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEAVANC.OnArgument (S : String ) ;
begin
  Inherited ;
(*
  SetControlChecked ('RDOCUMENT',False);
  if LaTOB.GetValue ('PARAG') = 'X' then SetControlEnabled ('RPARAG',true)
                                   else SetControlEnabled ('RPARAG',False);
*)
  SetControlEnabled ('RPARAG',true);
  SetControlChecked ('RPARAG',true);
  Bok := TToolbarButton97 (GetControl ('BValider'));
  Bok.OnClick := BokClick;
  Bannul := TToolbarButton97 (GetControl ('BFerme'));
  Bannul.OnClick := BAnnulClick;
end ;

procedure TOF_SAISIEAVANC.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEAVANC.BokClick (Sender : Tobject);
begin

NextPrevControl (tform(ecran));

LaTob.putvalue ('POURCENT',THNumEdit(GetControl('POURCENTAVANC')).Value );
If LaTob.getvalue('POURCENT') = 0 then
  begin
	if PGIAsk (TraduireMemoire('Le pourcentage est à zéro. Confirmez-vous ?'))=MrNo then
    begin
    SetFocusControl('POURCENTAVANC');
    Ecran.ModalResult :=  0;
    Exit;
    end;
  end;

if (TRadioButton (GetControl('RDOCUMENT')).Checked) then
begin
  LaTob.putvalue('APPLIC', 0);
end else if (TRadioButton (GetControl('RPARAG')).Checked) then
begin
  LaTob.putvalue('APPLIC', 1);
end else
begin
  LaTob.putvalue('APPLIC', 2);
end;
LaTob.putValue('SAISOK','X');
TheTob := Latob;
close;
end;

procedure TOF_SAISIEAVANC.BAnnulClick (Sender : Tobject);
begin
LaTob.putValue('SAISOK','-');
TheTob := Latob;
close;
end;

Initialization
  registerclasses ( [ TOF_SAISIEAVANC ] ) ;
end.

