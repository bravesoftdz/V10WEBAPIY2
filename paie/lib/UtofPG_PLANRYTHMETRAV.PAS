{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PLANRYTHMETRAV ()
Mots clefs ... : TOF;PLANRYTHMETRAV
*****************************************************************}
Unit UtofPG_PLANRYTHMETRAV ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_PLANRYTHMETRAV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure OnChangeCBTYPE(Sender : TObject);
    procedure OnClickBValider(Sender : TObject);
  end ;

Implementation

Uses
  PGPlanningPresenceCycle, HTB97, PgPlanningOutils;

procedure TOF_PLANRYTHMETRAV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnArgument (S : String ) ;
begin
  Inherited ;
  (GetControl('CBTYPE') as THValComboBox).OnChange := OnChangeCBTYPE;
  (GetControl('BValider') as TToolBarButton97).OnClick := OnClickBValider;
end ;

procedure TOF_PLANRYTHMETRAV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PLANRYTHMETRAV.OnChangeCBTYPE(Sender: TObject);
var
  StCBTYPE : String;
begin
  StCBTYPE := GetControlText('CBTYPE');
  if StCBTYPE = 'CYC' then
  begin
    (GetControl('MCBLISTECYCLES') as THMultiValComboBox).DataType := 'PGCYCLE';
  end else if StCBTYPE = 'MOD' then
  begin
    (GetControl('MCBLISTECYCLES') as THMultiValComboBox).DataType := 'PGMODELECYCLE';
  end;
end;

procedure TOF_PLANRYTHMETRAV.OnClickBValider(Sender: TObject);
var
  stMCBLISTECYCLES, stTemp, stListeCycles : String;
begin
  { Ajout des guillemets autour des codes cycle pour la liste des cycles à afficher }
  stMCBLISTECYCLES := GetControlText('MCBLISTECYCLES');
  if stMCBLISTECYCLES <> '<<Tous>>' then
  begin
    stTemp := READTOKENST(stMCBLISTECYCLES);
    While stTemp <> '' do
    begin
      stListeCycles := stListeCycles + ',"'+stTemp+'"';
      stTemp := READTOKENST(stMCBLISTECYCLES);
    end;
    stListeCycles := copy(stListeCycles,2,Length(stListeCycles)-1);
  end else stListeCycles := '';
  if (StrToDate(GetControlText('EDDATE')) = iDate1900)
  or (StrToDate(GetControlText('EDDATE_'))-StrToDate(GetControlText('EDDATE')) > 366) then
    PGIError('Vous devez choisir une période d''affichage d''un an maximum', Ecran.Caption)
  else
    PGPlanningPresenceCycleOpen(StrToDate(GetControlText('EDDATE')),StrToDate(GetControlText('EDDATE_')),'',stListeCycles, [piaSemaine, piaMois, piaHeure, piaOutlook], GetControlText('CBTYPE'));  //
end;

Initialization
  registerclasses ( [ TOF_PLANRYTHMETRAV ] ) ;
end.
