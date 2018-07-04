{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 22/07/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARECHREXEL ()
Mots clefs ... : TOF;BTPARECHREXEL
*****************************************************************}
Unit BTPARECHREXEL_TOF ;

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
{$ENDIF}
     uTob, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ; 

Type
  TOF_BTPARECHREXEL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBPARAM : TOB;
    TheAuxiliaire : string;
  end ;

Implementation

procedure TOF_BTPARECHREXEL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPARECHREXEL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARECHREXEL.OnUpdate ;
begin
  Inherited ;
  if (GetControlText('BTE_NUMCLI')='') or (GetControlText('BTE_CODE')='') or (GetControlText('BTE_AGENCETIERS')='') then
  begin
    PgiInfo ('Vous devez saisir vos identifiants de société ainsi que l''agence REXEL de destination');
    ecran.ModalResult := 0;
    exit;
  end;
  TOBParam.GetEcran(ecran);
  TOBParam.SetAllModifie(true);
  TOBParam.InsertOrUpdateDB(false);    
end ;

procedure TOF_BTPARECHREXEL.OnLoad ;
var QQ : TQuery;
begin
  Inherited ;
  QQ := OpenSql ('SELECT * FROM BTIERSECHG WHERE BTE_AUXILIAIRE="'+GetControlText('BTE_AUXILIAIRE')+'"',true,1,'',true);
  if not QQ.eof then
  begin
    TOBPARAM.SelectDB('',QQ);
    TOBPAram.PutEcran(ecran); 
  end;
  ferme (QQ);
end ;

procedure TOF_BTPARECHREXEL.OnArgument (S : String ) ;
var Critere,ChampMul,ValMul : string;
    x : integer;
begin
  Inherited ;
  TOBPARAM := TOB.Create ('BTIERSECHG',nil,-1);
  repeat
    Critere := Trim(ReadTokenSt(S));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'AUXILIAIRE' then TheAuxiliaire := ValMul;
      end;
    end;
  until Critere = '';
  if TheAuxiliaire <> '' then
  begin
    SetControlText('BTE_AUXILIAIRE',TheAuxiliaire);
    SeTControltext('TBTE_AUXILIAIRE',rechDom('TZTTOUS',TheAuxiliaire,false));
  end;
end ;

procedure TOF_BTPARECHREXEL.OnClose ;
begin
  TOBPARAM.free;
  Inherited ;
end ;

procedure TOF_BTPARECHREXEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARECHREXEL.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTPARECHREXEL ] ) ; 
end.
