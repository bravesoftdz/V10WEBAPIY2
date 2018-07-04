{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 24/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGDUPLIEMETTEUR ()
SUITE          : Duplication d'un émetteur social
Mots clefs ... : TOF;UTOFPGDUPLIEMETTEUR
*****************************************************************}
Unit UTOFPGDUPLIEMETTEUR ;

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
     UTOF,
     PGOutils,
     HTB97;

Type
  TOF_UTOFPGDUPLIEMETTEUR = Class (TOF)
    public
      procedure OnArgument (S : String ) ; override ;
      procedure OnClose                  ; override ;
      procedure Validation(Sender: TObject);
  end ;

Implementation

procedure TOF_UTOFPGDUPLIEMETTEUR.OnArgument (S : String ) ;
var
  Btn: TToolBarButton97;
begin
  Inherited ;
  DonneCodeDupliquerDos('', '', '', '');

  Btn := TToolBarButton97(GetControl('BValider'));
  if Btn <> nil then Btn.OnClick := Validation;
end ;

procedure TOF_UTOFPGDUPLIEMETTEUR.OnClose ;
var
  ECode: THEdit;
begin
  inherited;
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  if (ECode <> nil) then
  begin
    if Ecode.text = '' then
      DonneCodeDupliquerDos('', '', '', '');
  end;
end ;

procedure TOF_UTOFPGDUPLIEMETTEUR.Validation(Sender: TObject);
var
  ECode: THEdit;
  Btn: TToolBarButton97;
begin
  Btn := TToolBarButton97(GetControl('BValider'));
  Ecode := THEdit(GetControl('CODEDUPLIQUER'));
  if (ECode <> nil) and (Btn <> nil) then
  begin
    DonneCodeDupliquerDos('', '', '', '');
    if (ECode.Text = '') then
    begin
      PGIBox('Vous devez renseigner le nouveau code émetteur social!', 'Duplication');
      exit;
    end;

    PGIBox('Code émetteur social correct.#13#10#13#10Cliquer sur la croix pour terminer le traitement.', 'Duplication');
    SetControlEnabled('BValider',False);    
    DonneCodeDupliquerDos(Ecode.text,'', '', '');
  end;
end;

Initialization
  registerclasses ( [ TOF_UTOFPGDUPLIEMETTEUR ] ) ; 
end.
