{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/11/2011
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BTTYPEAFFAIRE (BTTYPEAFFAIRE)
Mots clefs ... : TOM;BTTYPEAFFAIRE
*****************************************************************}
Unit UTOMBTTYPEAFFAIRE ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob ;

Type
  TOM_BTTYPEAFFAIRE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
  private

    TypeAffaire : string;

    Affaire0    : THEdit;

    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Critere: String);
    end ;

const
  // libellés des messages de la TOM Affaire
  TexteMessage : array [1..1] of string = (
    {1} 'Suppression impossible ce type d''affaire est utilisée sur des contrats et/ou Appels'
    ) ;
Implementation

procedure TOM_BTTYPEAFFAIRE.OnNewRecord ;
begin
  Inherited ;

  Affaire0.Text := TypeAffaire;

end ;

procedure TOM_BTTYPEAFFAIRE.OnDeleteRecord ;
var Req : String;
begin
  Inherited ;
  TypeAffaire := GetControlText('BTY_TYPEAFFAIRE');
  //Contrôle si le type d'affaire non présent sur un contrat ou un appel !!!
  Req := 'SELECT AFF_TYPEAFFAIRE FROM AFFAIRE WHERE AFF_TYPEAFFAIRE = "' + TypeAffaire + '"';

  if Existesql(Req) then
  begin
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    exit;
  end;
end ;

procedure TOM_BTTYPEAFFAIRE.OnUpdateRecord ;
begin
  Inherited ;

end ;

procedure TOM_BTTYPEAFFAIRE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnLoadRecord ;
begin
  Inherited ;

end ;

procedure TOM_BTTYPEAFFAIRE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnArgument ( S: String ) ;
var Critere : String;
    X       : Integer;
    Champ   : string;
    Valeur  : string;
begin
  Inherited;

  //Récupération valeur de argument
  Critere:=(Trim(ReadTokenSt(S)));

  while (Critere <> '') do
  begin
    if Critere <> '' then
      begin
      X := pos (':', Critere) ;
      if x = 0 then
        X := pos ('=', Critere) ;
      if x <> 0 then
        begin
        Champ := copy (Critere, 1, X - 1) ;
        Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        ControleChamp(champ, valeur);
        end
      end;
    ControleCritere(Critere);
    Critere := (Trim(ReadTokenSt(S)));
  end;

  if Assigned(GetControl('BTY_AFFAIRE0')) then
  begin
    Affaire0 := THEdit(ecran.FindComponent('BTY_AFFAIRE0'));
    Affaire0.Text := TypeAffaire;
  end;

  TFFicheListe (Ecran).SetNewRange (TFFicheListe (Ecran).FRange, '(BTY_AFFAIRE0="' + TypeAffaire + '")');
  if TypeAffaire = 'I' then
  begin
    TGroupBox(ecran.FindComponent('GPBINFOS')).visible := false;
    ecran.caption := 'Type Contrat :';
  end else
  begin
    ecran.caption := 'Type Intervention :';
  end;
end ;

Procedure TOM_BTTYPEAFFAIRE.ControleChamp(Champ : String;Valeur : String);
Begin

  //Chargement du code affaire
  if Champ ='TYPEAFFAIRE' then TypeAffaire := valeur;

end;

Procedure TOM_BTTYPEAFFAIRE.ControleCritere(Critere : String);
Begin

end;


procedure TOM_BTTYPEAFFAIRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BTTYPEAFFAIRE.OnAfterRollBack;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_BTTYPEAFFAIRE ] ) ; 
end.

