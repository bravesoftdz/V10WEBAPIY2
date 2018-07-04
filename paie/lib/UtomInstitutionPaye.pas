{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... :
Modifié le ... :   /  /
Description .. : Gestion de la table INSTITUTIONPAYE
Mots clefs ... : PAIE
*****************************************************************}
{
 PT1 : 06/01/2003 : V591 MF
                             1- Correction des avertissements de compile
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****

 PT2 : 07/02/2006 : V_65 MF     FQ 11720  : En création correction contrôle institution déjà
                                existante.

}
unit UtomInstitutionPaye;

interface
uses
  {$IFDEF EAGLCLIENT}
  eFiche,
//unused  MaineAGL,
//unused  UtileAGL,
  UTOB,
  {$ELSE}
  db, Fiche,
  {$ENDIF$}
  Classes, UTOM, HMsgBox,
  PGOutils,PgOutils2;
type
  TOM_InstitutionPaye = class(TOM)

    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnChangeField(F: TField); override;

  private
    LectureSeule, CEG, STD, DOS: Boolean;
    ///PT1-1      mode : string;
  end;

implementation

//unused uses P5Def;
//============================================================================================//
//                                 PROCEDURE OnArgument
//============================================================================================//
procedure TOM_InstitutionPaye.OnArgument;
begin
  AccesPredefini('TOUS', CEG, STD, DOS);
  inherited
end;

//============================================================================================//
//                                 PROCEDURE On Load Record
//============================================================================================//
procedure TOM_InstitutionPaye.OnLoadRecord;
begin
  LectureSeule := FALSE;

  if (Getfield('PIP_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;

  if (Getfield('PIP_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  if (Getfield('PIP_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('BDelete', DOS);
  end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PIP_PREDEFINI', False);
  SetControlEnabled('PIP_INSTITUTION', False);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PIP_PREDEFINI', True);
    SetControlEnabled('PIP_INSTITUTION', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;

end;
//============================================================================================//
//                                 PROCEDURE On New Record
//============================================================================================//
procedure TOM_InstitutionPaye.OnNewRecord;
begin
  if (CEG = TRUE) then
    SetField('PIP_PREDEFINI', 'CEG')
  else
    SetField('PIP_PREDEFINI', 'DOS');
end;

//============================================================================================//
//                                 PROCEDURE On Delete Record
//============================================================================================//
procedure TOM_InstitutionPaye.OnDeleteRecord;
var
  Champ: array[1..1] of Hstring;
  Valeur: array[1..1] of variant;
  ExisteCod: Boolean;

begin

  Champ[1] := 'POG_INSTITUTION';
  Valeur[1] := GetField('PIP_INSTITUTION');
  ;
  ExisteCod := RechEnrAssocier('ORGANISMEPAIE', Champ, Valeur);
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Cette institution est utilisée.#13#10' +
      'Vous ne pouvez pas la supprimer!';
  end
  else
  begin
    Champ[1] := 'PDD_INSTITUTION';
    Valeur[1] := GetField('PIP_INSTITUTION');
    ;
    ExisteCod := RechEnrAssocier('DUCSDETAIL', Champ, Valeur);
    if ExisteCod = TRUE then
    begin
      LastError := 1;
      LastErrorMsg := 'Attention! Cette institution est utilisée.#13#10' +
        'Vous ne pouvez pas la supprimer!';
    end;
  end;
end;
//============================================================================================//
//                                 PROCEDURE On Change Field
//============================================================================================//
procedure TOM_InstitutionPaye.OnChangeField(F: TField);
var
  //PT1-1   Champ :array[1..1] of string;
  //PT1-1   Valeur :array[1..1] of variant ;
  //PT1-1   ExisteCod : Boolean;
  Pred: string;
begin

  if (F.FieldName = 'PIP_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PIP_PREDEFINI');
    if Pred = '' then exit;
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer d''institution prédéfini CEGID', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PIP_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer d''institution prédéfini Standard', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PIP_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PIP_PREDEFINI') then SetField('PIP_PREDEFINI', pred);
  end;
end;

//============================================================================================//
//                                 PROCEDURE On Update Record
//============================================================================================//
procedure TOM_InstitutionPaye.OnUpdateRecord;
var
{PT2
  Champ: array[1..1] of string;
  Valeur: array[1..1] of variant;}
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;

  ExisteCod: Boolean;

begin
  if (DS.State = dsinsert) then
  begin
    if (GetField('PIP_PREDEFINI') <> 'DOS') then
      SetField('PIP_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PIP_NODOSSIER', PgRendNoDossier());

    // Test si champ institution renseigné
    if (GetField('PIP_INSTITUTION') = '') then
    begin
      LastError := 1;
      LastErrorMsg := 'Vous devez renseigner une institution';
      SetFocusControl('PIP_INSTITUTION');
    end;

    // test si institution existe déjà
// d PT2
    Champ[1] := 'PIP_PREDEFINI';
    Valeur[1] := GetField('PIP_PREDEFINI');
    Champ[2] := 'PIP_NODOSSIER';

    if GetField('PIP_PREDEFINI') = 'DOS' then
       Valeur[2] := PgRendNoDossier()
    else
        Valeur[2] := '000000';

    Champ[3] := 'PIP_INSTITUTION';
    Valeur[3] := GetField('PIP_INSTITUTION');

//    Champ[1] := 'PIP_INSTITUTION';
//    Valeur[1] := GetField('PIP_INSTITUTION');
// f PT2
    ;
    ExisteCod := RechEnrAssocier('INSTITUTIONPAYE', Champ, Valeur);
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      LastErrorMsg := 'Cette institution existe déjà';
      SetFocusControl('PIP_INSTITUTION');
    end
  end;
end;

initialization
  registerclasses([TOM_InstitutionPaye]);

end.

