{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 08/09/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFFECTATION ()
Mots clefs ... : TOF;UTOFAFFECTATION
*****************************************************************}
Unit UTofAffectation;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     Lookup,
     uTob,
     HTB97,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_AFFECTATION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

    Responsable : THEdit;

    procedure AfficheResponsable(TobRessource: tob);
    procedure LectureResponsable;

    procedure AppelResponsable(sender: Tobject);


  end ;

Implementation

procedure TOF_AFFECTATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnUpdate ;
begin
  Inherited ;

  if GetControlText('ARS_RESSOURCE') <> '' then
     Begin
	   LaTob.PutValue('RETOUR','X');
	   LaTob.PutValue('RESSOURCE',GetControlText('ARS_RESSOURCE'));
     exit;
     end;

end ;

procedure TOF_AFFECTATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnArgument (S : String ) ;
begin
  Inherited ;

  Responsable :=THEdit(GetControl('AFF_RESPONSABLE'));
  Responsable.OnElipsisClick := AppelResponsable;


  SetControlText('ARS_RESSOURCE','');

end ;

procedure TOF_AFFECTATION.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AFFECTATION.AppelResponsable(sender: Tobject);
Var Req        : String;
begin

 	Req := '';
  Req := ' ARS_TYPERESSOURCE="SAL" OR ARS_TYPERESSOURCE="ST"';

  SetControlProperty('AFF_RESPONSABLE', 'Plus', req);

  LookupCombo(Responsable);

  if getcontroltext('AFF_RESPONSABLE')='' then exit;

  SetControlText('ARS_RESSOURCE', GetControlText('AFF_RESPONSABLE'));
  SetControlText('AFF_RESPONSABLE', '');

  //Lecture des Ressources sélectionné et affichage des informations
  LectureResponsable;

end;

procedure TOF_AFFECTATION.LectureResponsable;
Var TobRessource : TOB;
    Req          : String;
Begin

  Req := '';
  Req := 'SELECT * FROM RESSOURCE ';
  Req := Req + 'WHERE ARS_RESSOURCE ="' + GetControlText('ARS_RESSOURCE') + '"';

  TobRessource := Tob.Create('LesRessources',Nil, -1);
  TobRessource.LoadDetailDBFromSQL('RESSOURCE',req,false);

  if TobRessource.Detail.Count <> 0 then AfficheResponsable(TobRessource);

  TobRessource.free;

end;

procedure TOF_AFFECTATION.AfficheResponsable(TobRessource: tob);
Var Nom : String;
begin

	Nom := TobRessource.Detail[0].GetValue('ARS_LIBELLE') + ' ' + TobRessource.Detail[0].GetValue('ARS_LIBELLE2');

	SetControlText('AFF_RESPONSABLE',Nom);

end;

Initialization
  registerclasses ( [ TOF_AFFECTATION ] ) ;
end.

