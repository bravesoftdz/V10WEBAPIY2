{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion des guides d'écritures pour la génération
Suite ........ : comptablre
Mots clefs ... : PAIE;
*****************************************************************}
unit UTOMGUIDEECRPAIE;
{
  PT1  : 23/09/02   : V585  PH Acces au type alimentation pour les comptes de la classe 5
  PT2  : 21/08/03   : V_421 PH FQ 10146 Le modèle CEGID n'est plus modifiable
  PT3  : 21/08/03   : V_421 PH Correction requête non fermée
  PT4  : 31/03/04   : V_500 PH Prise en compte du cas d'un journal provisions CP ou RTT
  PT5  : 02/06/04   : V_500 PH FQ 11161 Controle acces champ de la clé
  PT6  : 04/05/06   : V_650 PH FQ 13088 Non acces zone folio et numéro de ligne si lecture seule

}

interface
uses Windows, StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fiche,
{$ELSE}
  eFiche,
{$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOM, UTOB, Vierge, P5Def, Pgoutils,
  Pgoutils2, AGLInit;
type
  TOM_GUIDEECRPAIE = class(TOM)
  private
    PGPredef, PGNodossier, PGJeu, TypeProv: string;
    Dernum: Integer;
    procedure AfficheTypAlim;
    procedure ControlAffAlim;
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnChangeField(F: TField); override;
  end;

implementation

procedure TOM_GUIDEECRPAIE.AfficheTypAlim;
var
  zone: TControl;
begin
  Zone := getcontrol('PGC_ALIM421');
  InitialiseCombo(zone);
  SetControlProperty('TPGC_ALIM421', 'Visible', TRUE);
  SetControlProperty('PGC_ALIM421', 'Visible', TRUE);
end;

// PT4 Prise en compte du cas d'un journal provisions CP ou RTT

procedure TOM_GUIDEECRPAIE.ControlAffAlim;
var
  LeDebut: string;
begin
  LeDebut := Copy(GetField('PGC_GENERAL'), 1, 1);
  if TypeProv = '' then
  begin
    if (Copy(GetField('PGC_GENERAL'), 1, 3) = '421') or (Copy(GetField('PGC_GENERAL'), 1, 3) = '641')
      or (LeDebut = '4') or (LeDebut = '5') then AfficheTypAlim;
  end
  else AfficheTypAlim;
end;
// FIN PT4

procedure TOM_GUIDEECRPAIE.OnArgument(Arguments: string);
var
  st: string;
begin
  inherited;
  st := Trim(Arguments);
  PGPredef := '';
  PGNodossier := '';
  PGJeu := '';
  PGPredef := ReadTokenSt(st); // Recup Predefini
  PGNodossier := ReadTokenSt(st); // Recup Nodossier
  PGJeu := ReadTokenSt(st); // Recup modele
  //   PT4  : 31/03/04   : V_500 PH Prise en compte du cas d'un journal provisions CP ou RTT
  TypeProv := ReadTokenSt(st); // Recup type de journal=type provision ou ODS paies
end;


procedure TOM_GUIDEECRPAIE.OnChangeField(F: TField);
var
  Q: TQuery;
  NumLigne: Integer;
  chaine, LeDebut: string;
begin
  inherited;
  if (F.FieldName = 'PGC_NUMFOLIO') and (DS.State in [dsInsert]) then
  begin
    Dernum := 0;
    Chaine := ' SELECT MAX (PGC_NUMECRIT) FROM GUIDEECRPAIE WHERE PGC_NODOSSIER="' + PGNodossier + '" AND PGC_PREDEFINI="' +
      PGPredef + '" AND PGC_JEUECR=' + PGJeu + ' AND PGC_NUMFOLIO="' + // DB2
      GetField('PGC_NUMFOLIO') + '"';
    Q := OpenSql(Chaine, TRUE);
    if not Q.EOF then
    begin
      NumLigne := Q.Fields[0].AsInteger;
      SetField('PGC_NUMECRIT', NumLigne + 1);
    end
    else SetField('PGC_NUMECRIT', 1);
    Ferme(Q); //   PT3  : 21/08/03   : V_421 PH Correction requête non fermée
  end;
  if (F.FieldName = 'PGC_GENERAL') then
  begin
    LeDebut := Copy(GetField('PGC_GENERAL'), 1, 1);
    // PT4 Prise en compte du cas d'un journal provisions CP ou RTT
    SetControlProperty('TPGC_ALIM421', 'Visible', FALSE);
    SetControlProperty('PGC_ALIM421', 'Visible', FALSE);
    ControlAffAlim;
    // FIN PT4
  end;
end;

procedure TOM_GUIDEECRPAIE.OnLoadRecord;
var
  LectureSeule, CEG, STD, DOS, OnFerme: boolean;
begin
  inherited;
  //   PT2  : 21/08/03   : V_421  PH Acces au type alimentation pour les comptes de la classe 5
  AccesPredefini('TOUS', CEG, STD, DOS);
  SetControlEnabled('BInsert', TRUE);
  LectureSeule := FALSE;
  if (Getfield('PGC_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
    SetControlEnabled('BInsert', CEG);
  end
  else
    if (Getfield('PGC_PREDEFINI') = 'STD') then
    begin
      LectureSeule := (STD = False);
      PaieLectureSeule(TFFiche(Ecran), (STD = False));
      SetControlEnabled('BDelete', STD);
      SetControlEnabled('BInsert', STD);
    end
    else
      if (Getfield('PGC_PREDEFINI') = 'DOS') then
      begin
        LectureSeule := False;
        PaieLectureSeule(TFFiche(Ecran), False);
        SetControlEnabled('BDelete', DOS);
        SetControlEnabled('BInsert', DOS);
      end;
  // FIN PT2

  SetControlEnabled('BDUPLIQUER', True);

  if DS.State in [dsInsert] then
  begin
    PaieLectureSeule(TFFiche(Ecran), False);
    // PT5  : 02/06/04   : V_500 PH FQ 11161 Controle acces champ de la clé
        //    SetControlEnabled('PGC_PREDEFINI', not LectureSeule);
        //    SetControlEnabled('PGC_JEUECR', not LectureSeule);
    SetControlEnabled('PGC_PREDEFINI', FALSE);
    SetControlEnabled('PGC_JEUECR', FALSE);
    SetControlEnabled('PGC_NUMFOLIO', True);
    SetControlEnabled('PGC_NUMECRIT', true);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BDelete', False);
    SetControlEnabled('BValider', not LectureSeule);
  end
  else
  begin
    SetControlEnabled('PGC_PREDEFINI', FALSE);
    SetControlEnabled('PGC_JEUECR', FALSE);
    // DEB PT6
    if LectureSeule then
    begin
      SetControlEnabled('PGC_NUMFOLIO', FALSE);
      SetControlEnabled('PGC_NUMECRIT', FALSE);
    end
    else
    begin
      SetControlEnabled('PGC_NUMFOLIO', TRUE);
      SetControlEnabled('PGC_NUMECRIT', TRUE);
    end;
    // FIN PT6
  end;
  // PT4 Prise en compte du cas d'un journal provisions CP ou Rtt
// FIN PT5

  ControlAffAlim;
end;

procedure TOM_GUIDEECRPAIE.OnNewRecord;
begin
  inherited;
  if PGPredef <> '' then SetField('PGC_PREDEFINI', PGPredef);
  if PGNodossier <> '' then SetField('PGC_NODOSSIER', PGNodossier);
  if PGJeu <> '' then SetField('PGC_JEUECR', PGJeu);
  Dernum := Dernum + 1;
  SetField('PGC_NUMECRIT', Dernum);
// FIN PT5
  SetControlEnabled('PGC_PREDEFINI', FALSE);
  SetControlEnabled('PGC_JEUECR', FALSE);
end;

procedure TOM_GUIDEECRPAIE.OnUpdateRecord;
begin
  inherited;
  if GetField('PGC_NUMFOLIO') = '' then
  begin
    LastError := 1;
    PGIBox('Vous devez renseigner le folio!', 'Guide des écritures comptable');
    SetFocusControl('PGC_NUMFOLIO');
  end;

end;



initialization
  registerclasses([TOM_GUIDEECRPAIE]);
end.

