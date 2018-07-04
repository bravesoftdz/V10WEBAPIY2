{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 10/04/2014
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREVALPMAP ()
Mots clefs ... : TOF;BTREVALPMAP
*****************************************************************}
Unit BTREVALPMAP_TOF ;

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
     uEntCommun,
     HMsgBox,
     HTB97,
     UTOB,
     UTOF;

Type
  TOF_BTREVALPMAP = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBEchange : TOB;
    fTraitementGlobal : Boolean;
  end ;

Implementation
uses  UtilSaisieConso,
      FactCalc,
      PiecesRecalculs,
      utilTOBpiece,
      UtilPMAPCalcul
      ;

procedure TOF_BTREVALPMAP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREVALPMAP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREVALPMAP.OnUpdate ;
begin
  Inherited ;
  if not fTraitementGlobal then
  begin
    if Assigned (TOBEchange) then
    begin
      if THCheckbox (GetControl('CHKLIV')).Checked then TOBEchange.SetString('LES LIV','X');
      if THCheckbox (GetControl('CHKCONSO')).Checked then TOBEchange.SetString('LES CONSOS','X');
      if THValComboBox (getControl('DEPOT')).Enabled then TOBEchange.SetString('DEPOT',THValComboBox (getControl('DEPOT')).Value);
      TOBEchange.SetString('OKGO','X');
    end;
  end else
  begin
    if THValComboBox (getControl('DEPOT')).Value = '' then exit;
    If PgiAsk ('ATTENTION : Ce traitement risque d''être long#13#10 Confirmez-vous ?')<>mryes then exit;
    LanceTraitementRecalcDepot (THValComboBox (getControl('DEPOT')).Value,
    														THCheckbox (GetControl('CHKLIV')).Checked,
                                THCheckbox (GetControl('CHKCONSO')).Checked );
  end;
end ;

procedure TOF_BTREVALPMAP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREVALPMAP.OnArgument (S : String ) ;
var Critere : String;
		x : Integer;
    ChampMul,ValMul : string;
begin
  Inherited ;
  TOBEchange := LaTOB;
  fTraitementGlobal := false;
  SetControlText('LARTICLE','');
  //
  repeat
    Critere := Trim(ReadTokenSt(S));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'ARTICLE' then SetControlText('LARTICLE','ARTICLE : '+ValMul);
        if ChampMul = 'DEPOT' then
        begin
        	SetControlText('DEPOT',ValMul);
          SetControlEnabled('DEPOT',False);
        end;
      end else
      begin
        if Critere = 'TRAITEDEPOT' then fTraitementGlobal := true;
      end;
    end;
  until Critere = '';

end ;

procedure TOF_BTREVALPMAP.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTREVALPMAP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREVALPMAP.OnCancel () ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_BTREVALPMAP ] ) ;
end.

