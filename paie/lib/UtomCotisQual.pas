{***********UNITE*************************************************
Auteur  ...... : PAIE
Créé le ...... : 02/04/2002
Modifié le ... :   /  /
Description .. : Gestion des Qualifiant de cotisations. Utilisés dans le
Suite ........ : traitement de la Ducs-EDI
Mots clefs ... : PAIE, PGDUCS, PGDUCSEDI
*****************************************************************}
unit UtomCotisQual;

{
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT1 : 11/04/2006 : V_650 MF  FQ 12058 : correction traitement prédéfini.
}

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  eFiche, UTOB,
  {$ELSE}
  Fiche, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  UTOM,
  HEnt1,
  PgOutils,
  Classes;

type
  TOM_CotisQual = class(TOM)

    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
  private
    LectureSeule, CEG, STD, DOS: Boolean;

  end;
implementation
uses P5Def,PgOutils2;
procedure TOM_CotisQual.OnArgument;
begin
  AccesPredefini('TOUS', CEG, STD, DOS);
  inherited
end;

procedure TOM_CotisQual.OnLoadRecord;
begin

  LectureSeule := FALSE;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PQC_PREDEFINI', False);
  SetControlEnabled('PQC_QUALIFIANT', False);


  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PQC_PREDEFINI', True);
    SetControlEnabled('PQC_QUALIFIANT', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
   if (Getfield('PQC_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;

  if (Getfield('PQC_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  if (Getfield('PQC_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', DOS);
  end;
  PaieConceptPlanPaie(Ecran);
end;

procedure TOM_CotisQual.OnNewRecord;
begin

  if (CEG = TRUE) then
    SetField('PQC_PREDEFINI', 'CEG')
  else
    SetField('PQC_PREDEFINI', 'DOS');
end;

procedure TOM_CotisQual.OnDeleteRecord;
var
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  ExisteCod: Boolean;
begin

// d PT1
  Champ[1] := 'PDP_COTISQUAL';
  Valeur[1] := GetField('PQC_QUALIFIANT');
  Champ[2] := '';
  Valeur[2] := '';
  Champ[3] := '';
  Valeur[3] := '';
  ExisteCod := RechEnrAssocier('DUCSPARAM', Champ, Valeur);
  if ExisteCod = TRUE then
  begin
    if (GetField('PQC_PREDEFINI') <> 'CEG') then
    begin
      Champ[1] := 'PQC_QUALIFIANT';
      Valeur[1] := GetField('PQC_QUALIFIANT');
      Champ[2] := 'PQC_PREDEFINI';
      Valeur[2] := 'CEG';
      Champ[3] := 'PQC_NODOSSIER';
      Valeur[3] := '000000';
      ExisteCod := RechEnrAssocier('COTISQUAL', Champ, Valeur);
      if (ExisteCod = FALSE) and (GetField('PQC_PREDEFINI') <> 'STD') then
      begin
        Champ[2] := 'PQC_PREDEFINI';
        Valeur[2] := 'STD';
        ExisteCod := RechEnrAssocier('COTISQUAL', Champ, Valeur);
      end;
      if ExisteCod = FALSE  then
      begin
        LastError := 1;
        LastErrorMsg := 'Attention! Ce qualifiant a été utilisée.#13#10' +
                        'Vous ne pouvez pas le supprimer!';
      end;
    end
    else
    begin
      Champ[1] := 'PQC_QUALIFIANT';
      Valeur[1] := GetField('PQC_QUALIFIANT');
      Champ[2] := 'PQC_PREDEFINI';
      Valeur[2] := 'STD';
      Champ[3] := 'PQC_NODOSSIER';
      Valeur[3] := '000000';
      ExisteCod := RechEnrAssocier('COTISQUAL', Champ, Valeur);
      if (ExisteCod = FALSE) then
      begin
        Valeur[2] := 'DOS';
        Champ[3] := 'PQC_NODOSSIER';
        Valeur[3] := PgRendNoDossier();
        ExisteCod := RechEnrAssocier('COTISQUAL', Champ, Valeur);
      end;
      if (ExisteCod = FALSE) then
      begin
        LastError := 1;
        LastErrorMsg := 'Attention! Ce qualifiant a été utilisée.#13#10' +
                        'Vous ne pouvez pas le supprimer!';
      end;
    end;
  end
// f PT1  
end;

procedure TOM_CotisQual.OnUpdateRecord;
var

  Champ: array[1..3] of Hstring;        // PT1
  Valeur: array[1..3] of variant;      // PT1
  ExisteCod: Boolean;

begin
  if (DS.State = dsinsert) then
  begin
    if (GetField('PQC_PREDEFINI') <> 'DOS') then
      SetField('PQC_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PQC_NODOSSIER', PgRendNoDossier());

    // Test si champ codification renseigné
    if (GetField('PQC_QUALIFIANT') = '') then
    begin
      LastError := 1;
      LastErrorMsg := 'Vous devez renseigner un qualifiant';
      SetFocusControl('PQC_QUALIFIANT');
    end;

    // test si codification existe déjà
    Champ[1] := 'PQC_QUALIFIANT';
    Valeur[1] := GetField('PQC_QUALIFIANT');
// d PT1
    Champ[2] := 'PQC_PREDEFINI';
    Valeur[2] := GetField('PQC_PREDEFINI');
    Champ[3] := 'PQC_NODOSSIER';
    Valeur[3] := GetField('PQC_NODOSSIER');
// f PT1

    ExisteCod := RechEnrAssocier('COTISQUAL', Champ, Valeur);
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      LastErrorMsg := 'Ce qualifiant existe déjà';
      SetFocusControl('PQC_QUALIFIANT');
    end
  end;
end;

initialization
  registerclasses([TOM_CotisQual]);

end.

