{***********UNITE*************************************************
Auteur  ...... : FCO
Créé le ...... : 29/12/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITLIENRUBRIQUE ()
Mots clefs ... : TOF;PGEDITLIENRUBRIQUE
*****************************************************************
PT1  18/04/2007 V_72 : FC Suppression du LanceEtatTob
PT2  31/08/2007 V_80 : FLO Ne pas éditer les liens vers les variables de paie
}
unit UTofPgEditLienRubrique;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ELSE}
  eQRS1,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, P5Def, EntPaie,
  PGoutils,PGoutils2, PGEditOutils,PgEditOutils2, HQry,
{$IFDEF EAGLCLIENT}
  UtileAGL,
{$ELSE}
  EdtREtat,
{$ENDIF}
  UTob,
  Paramsoc,
  HStatus,ed_tools;

type
  TOF_PGEDITLIENRUBRIQUE = class(TOF)
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure Change(Sender: TObject);
    procedure OnClose; override ; //PT2
  private
    TobEtat, T: Tob;
    procedure ExitEdit(Sender: TObject);
    procedure ConstruireTob();
  end;

implementation

procedure TOF_PGEDITLIENRUBRIQUE.OnLoad;
begin
  inherited;
end;

procedure TOF_PGEDITLIENRUBRIQUE.Change(Sender: TObject);
var
  Min, Max: string;
begin
  if GetControlText('PCT_NATURERUB') = '' then
  begin
    SetControlProperty('PCT_RUBRIQUE', 'DataType', '');
    SetControlProperty('PCT_RUBRIQUE_', 'DataType', '');
    SetControlProperty('PCT_RUBRIQUE', 'Text', '');
    SetControlProperty('PCT_RUBRIQUE_', 'Text', '');
    SetControlEnabled('PCT_RUBRIQUE', False);
    SetControlEnabled('PCT_RUBRIQUE_', False);
  end
  else
    if GetControlText('PCT_NATURERUB') = 'AAA' then
  begin
    SetControlEnabled('PCT_RUBRIQUE', True);
    SetControlEnabled('PCT_RUBRIQUE_', True);
    SetControlProperty('PCT_RUBRIQUE', 'DataType', 'PGREMUNERATION');
    SetControlProperty('PCT_RUBRIQUE_', 'DataType', 'PGREMUNERATION');
    RecupMinMaxTablette('PG', 'REMUNERATION', 'PRM_RUBRIQUE', Min, Max);
    SetControlProperty('PCT_RUBRIQUE', 'Text', Min);
    SetControlProperty('PCT_RUBRIQUE_', 'Text', Max);
  end
  else
    if GetControlText('PCT_NATURERUB') = 'BAS' then
  begin
    SetControlEnabled('PCT_RUBRIQUE', True);
    SetControlEnabled('PCT_RUBRIQUE_', True);
    SetControlProperty('PCT_RUBRIQUE', 'DataType', 'PGBASECOTISATION');
    SetControlProperty('PCT_RUBRIQUE_', 'DataType', 'PGBASECOTISATION');
    RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
    SetControlProperty('PCT_RUBRIQUE', 'Text', Min);
    SetControlProperty('PCT_RUBRIQUE_', 'Text', Max);
  end
  else
    if GetControlText('PCT_NATURERUB') = 'COT' then
  begin
    SetControlEnabled('PCT_RUBRIQUE', True);
    SetControlEnabled('PCT_RUBRIQUE_', True);
    SetControlProperty('PCT_RUBRIQUE', 'DataType', 'PGCOTISATION');
    SetControlProperty('PCT_RUBRIQUE_', 'DataType', 'PGCOTISATION');
    RecupMinMaxTablette('PG', 'COTISATION', 'PCT_RUBRIQUE', Min, Max);
    SetControlProperty('PCT_RUBRIQUE', 'Text', Min);
    SetControlProperty('PCT_RUBRIQUE_', 'Text', Max);
  end;
end;

procedure TOF_PGEDITLIENRUBRIQUE.OnArgument(S: string);
var
  Defaut: THEdit;
  combo: ThValComboBox;
begin
  inherited;
  Defaut := THEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    Defaut.text := GetParamSoc('SO_LIBELLE');

  combo := ThValComboBox(getcontrol('PCT_NATURERUB'));
  if combo <> nil then Combo.OnChange := Change;

  Defaut := ThEdit(getcontrol('PCT_RUBRIQUE'));
  if Defaut <> nil then Defaut.OnExit := ExitEdit;

  Defaut := ThEdit(getcontrol('PCT_RUBRIQUE_'));
  if Defaut <> nil then Defaut.OnExit := ExitEdit;

  TFQRS1(Ecran).Caption := 'Liens des rubriques entre elles';
  UpdateCaption(TFQRS1(Ecran));
end;

procedure TOF_PGEDITLIENRUBRIQUE.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 5) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 4);
end;

// Procédure de lancement de l'état avec TOB
procedure TOF_PGEDITLIENRUBRIQUE.OnUpdate;
var
  Pages               : TPageControl;
  Q, QLiens           : TQuery;
  TobRubrique         : Tob;
  TobLesLiens         : Tob;
  TobEtab             : Tob;
  Where, Tri, SQL     : string;
  Requete             : string;
  LibRubrique         : string;
  Rubrique            : string;
  NatureRubrique      : string;
  BaseCotisation      : string;
  TauxSal             : string;
  TauxPat             : string;
  BaseRem             : string;
  TauxRem             : string;
  CoeffRem            : string;
  LibelleRubrique     : string;
  i, c                : Integer;
  CritereNature       : string;
  CritereRubDu        : string;
  CritereRubAu        : string;
  TypeBase,TypeTaux,TypeCoeff,TypeTauxSal,TypeTauxPat : String; //PT2
  {$IFDEF EAGLCLIENT}
  StPages : String;
  {$ENDIF}

begin
  inherited;

  FreeAndNil(TobEtat);
  Pages := TPageControl(GetControl('Pages'));
  TobEtat := Tob.Create('Les salariés pour édition', nil, -1);

  if (GetControlText('PCT_NATURERUB') = 'BAS') or (GetControlText('PCT_NATURERUB') = 'COT') or (GetControlText('PCT_NATURERUB') = '') then
  begin
    Where := RecupWhereCritere(Pages);
    Tri := ' ORDER BY PCT_RUBRIQUE';
    SQL := 'SELECT PCT_RUBRIQUE,PCT_LIBELLE,PCT_BASECOTISATION,PCT_TAUXSAL,PCT_TAUXPAT,PCT_NATURERUB,PCT_TYPEBASE,PCT_TYPETAUXSAL,PCT_TYPETAUXPAT' + //PT2
      ' FROM COTISATION ';
    Requete := SQL + Where + Tri;

    Q := OpenSQL(Requete, True);
    TobRubrique := Tob.create('Les rubriques', nil, -1);
    TobRubrique.LoadDetailDB('COTISATION', '', '', Q, false);
    Ferme(Q);
    InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', TobRubrique.Detail.Count, False, True);
    InitMove(TobRubrique.Detail.Count, '');
    For i := 0 to TobRubrique.Detail.Count - 1 do
    begin
      Rubrique        := TobRubrique.Detail[i].GetValue('PCT_RUBRIQUE');
      LibRubrique     := TobRubrique.Detail[i].GetValue('PCT_LIBELLE');
      if (TobRubrique.Detail[i].GetValue('PCT_NATURERUB') = 'BAS') then
        NatureRubrique  := 'Base de cotisation'
      else
        NatureRubrique  := 'Cotisation';

      BaseCotisation  := TobRubrique.Detail[i].GetValue('PCT_BASECOTISATION');
      TauxSal         := TobRubrique.Detail[i].GetValue('PCT_TAUXSAL');
      TauxPat         := TobRubrique.Detail[i].GetValue('PCT_TAUXPAT');
      //PT2 - Début
      TypeBase        := TobRubrique.Detail[i].GetValue('PCT_TYPEBASE');
      TypeTauxSal     := TobRubrique.Detail[i].GetValue('PCT_TYPETAUXSAL');
      TypeTauxPat     := TobRubrique.Detail[i].GetValue('PCT_TYPETAUXPAT');
      //PT2 - Fin

      //Recherche des liens avec d'autres rubriques
      if (Rubrique <> BaseCotisation) and (BaseCotisation <> '') And (TypeBase <> 'VAR') then //PT2
      begin
        QLiens := OpenSQL('SELECT PCT_RUBRIQUE,PCT_LIBELLE FROM COTISATION ' +
        ' WHERE PCT_RUBRIQUE="' + BaseCotisation + '"', True);

        TobLesLiens := Tob.Create('Les liens', nil, -1);
        TobLesLiens.LoadDetailDB('Les liens', '', '', QLiens, False);
        Ferme(QLiens);
        If TobLesLiens.Detail.Count > 0 then
        begin
          For c := 0 to TobLesLiens.Detail.Count -1 do
          begin
            ConstruireTob;
            LibelleRubrique := TobLesLiens.Detail[c].GetValue('PCT_LIBELLE');
            T.PutValue('PCT_RUBRIQUE', Rubrique);
            T.PutValue('PCT_RUBRIQUE_LIB', Rubrique + ' ' + LibRubrique);
            T.PutValue('PCT_NATURERUB', NatureRubrique);
            T.PutValue('LIEN_RUBRIQUE_LIB', BaseCotisation + ' ' + LibelleRubrique);
          end;
        end;
        FreeAndNil(TobLesLiens);
      end;

      if (Rubrique <> TauxSal) and (TauxSal <> '') And (TypeTauxSal <> 'VAR') then //PT2
      begin
        QLiens := OpenSQL('SELECT PCT_RUBRIQUE,PCT_LIBELLE FROM COTISATION ' +
        'WHERE PCT_RUBRIQUE="' + TauxSal + '"', True);

        TobLesLiens := Tob.Create('Les liens', nil, -1);
        TobLesLiens.LoadDetailDB('Les liens', '', '', QLiens, False);
        Ferme(QLiens);
        If TobLesLiens.Detail.Count > 0 then
        begin
          For c := 0 to TobLesLiens.Detail.Count -1 do
          begin
            ConstruireTob;
            LibelleRubrique := TobLesLiens.Detail[c].GetValue('PCT_LIBELLE');
            T.PutValue('PCT_RUBRIQUE', Rubrique);
            T.PutValue('PCT_RUBRIQUE_LIB', Rubrique + ' ' + LibRubrique);
            T.PutValue('PCT_NATURERUB', NatureRubrique);
            T.PutValue('LIEN_RUBRIQUE_LIB', TauxSal + ' ' + LibelleRubrique);
          end;
        end;
        FreeAndNil(TobLesLiens);
      end;

      if (Rubrique <> TauxPat) and (TauxPat <> '')  And (TypeTauxPat <> 'VAR') then //PT2
      begin
        QLiens := OpenSQL('SELECT PCT_RUBRIQUE,PCT_LIBELLE FROM COTISATION ' +
        'WHERE PCT_RUBRIQUE="' + TauxPat + '"', True);

        TobLesLiens := Tob.Create('Les liens', nil, -1);
        TobLesLiens.LoadDetailDB('Les liens', '', '', QLiens, False);
        Ferme(QLiens);
        If TobLesLiens.Detail.Count > 0 then
        begin
          For c := 0 to TobLesLiens.Detail.Count -1 do
          begin
            ConstruireTob;
            LibelleRubrique := TobLesLiens.Detail[c].GetValue('PCT_LIBELLE');
            T.PutValue('PCT_RUBRIQUE', Rubrique);
            T.PutValue('PCT_RUBRIQUE_LIB', Rubrique + ' ' + LibRubrique);
            T.PutValue('PCT_NATURERUB', NatureRubrique);
            T.PutValue('LIEN_RUBRIQUE_LIB', TauxPat + ' ' + LibelleRubrique);
          end;
        end;
        FreeAndNil(TobLesLiens);
      end;
    end;
    FiniMoveProgressForm;
  end;

  if (GetControlText('PCT_NATURERUB') = 'AAA') or (GetControlText('PCT_NATURERUB') = '') then
  begin
    Tri := ' ORDER BY PRM_RUBRIQUE';
    SQL := 'SELECT PRM_RUBRIQUE,PRM_LIBELLE,PRM_BASEREM,PRM_TAUXREM,PRM_COEFFREM,PRM_NATURERUB,PRM_TYPEBASE,PRM_TYPETAUX,PRM_TYPECOEFF' + //PT2
      ' FROM REMUNERATION ';
    CritereNature := GetcontrolText('PCT_NATURERUB');
    CritereRubDu  := GetcontrolText('PCT_RUBRIQUE');
    CritereRubAu  := GetcontrolText('PCT_RUBRIQUE_');
    if (CritereNature <> '') and (CritereNature <> '<<Toutes>>') then
    begin
      Where := ' WHERE PRM_NATURERUB = "' + CritereNature + '"';
      if CritereRubDu <> '' then
        Where := Where + ' AND PRM_RUBRIQUE >= "' + CritereRubDu + '"';
      if CritereRubAu <> '' then
        Where := Where + ' AND PRM_RUBRIQUE <= "' + CritereRubAu + '"';
    end;

    Requete := SQL + Where + Tri;

    FreeAndNil(TobRubrique);
    Q := OpenSQL(Requete, True);
    TobRubrique := Tob.create('Les rubriques', nil, -1);
    TobRubrique.LoadDetailDB('REMUNERATION', '', '', Q, false);
    Ferme(Q);
    InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', TobRubrique.Detail.Count, False, True);
    InitMove(TobRubrique.Detail.Count, '');
    For i := 0 to TobRubrique.Detail.Count - 1 do
    begin
      Rubrique        := TobRubrique.Detail[i].GetValue('PRM_RUBRIQUE');
      LibRubrique     := TobRubrique.Detail[i].GetValue('PRM_LIBELLE');
      NatureRubrique  := 'Rémunération';
      BaseRem         := TobRubrique.Detail[i].GetValue('PRM_BASEREM');
      TauxRem         := TobRubrique.Detail[i].GetValue('PRM_TAUXREM');
      CoeffRem        := TobRubrique.Detail[i].GetValue('PRM_COEFFREM');
      //PT2 - Début
      TypeBase        := TobRubrique.Detail[i].GetValue('PRM_TYPEBASE');
      TypeTaux        := TobRubrique.Detail[i].GetValue('PRM_TYPETAUX');
      TypeCoeff       := TobRubrique.Detail[i].GetValue('PRM_TYPECOEFF');
      //PT2 - Fin

      //Recherche des liens avec d'autres rubriques
      if (Rubrique <> BaseRem) and (BaseRem <> '') And (TypeBase <> '03') then  //PT2
      begin
        QLiens := OpenSQL('SELECT PRM_RUBRIQUE,PRM_LIBELLE FROM REMUNERATION ' +
        ' WHERE PRM_RUBRIQUE="' + BaseRem + '"', True);

        TobLesLiens := Tob.Create('Les liens', nil, -1);
        TobLesLiens.LoadDetailDB('Les liens', '', '', QLiens, False);
        Ferme(QLiens);
        If TobLesLiens.Detail.Count > 0 then
        begin
          For c := 0 to TobLesLiens.Detail.Count -1 do
          begin
            ConstruireTob;
            LibelleRubrique := TobLesLiens.Detail[c].GetValue('PRM_LIBELLE');
            T.PutValue('PCT_RUBRIQUE', Rubrique);
            T.PutValue('PCT_RUBRIQUE_LIB', Rubrique + ' ' + LibRubrique);
            T.PutValue('PCT_NATURERUB', NatureRubrique);
            T.PutValue('LIEN_RUBRIQUE_LIB', BaseRem + ' ' + LibelleRubrique);
          end;
        end;
        FreeAndNil(TobLesLiens);
      end;

      if (Rubrique <> TauxRem) and (TauxRem <> '') and (TauxRem <> BaseRem) And (TypeTaux <> '03') then  //PT2
      begin
        QLiens := OpenSQL('SELECT PRM_RUBRIQUE,PRM_LIBELLE FROM REMUNERATION ' +
        'WHERE PRM_RUBRIQUE="' + TauxRem + '"', True);

        TobLesLiens := Tob.Create('Les liens', nil, -1);
        TobLesLiens.LoadDetailDB('Les liens', '', '', QLiens, False);
        Ferme(QLiens);
        If TobLesLiens.Detail.Count > 0 then
        begin
          For c := 0 to TobLesLiens.Detail.Count -1 do
          begin
            ConstruireTob;
            LibelleRubrique := TobLesLiens.Detail[c].GetValue('PRM_LIBELLE');
            T.PutValue('PCT_RUBRIQUE', Rubrique);
            T.PutValue('PCT_RUBRIQUE_LIB', Rubrique + ' ' + LibRubrique);
            T.PutValue('PCT_NATURERUB', NatureRubrique);
            T.PutValue('LIEN_RUBRIQUE_LIB', TauxRem + ' ' + LibelleRubrique);
          end;
        end;
        FreeAndNil(TobLesLiens);
      end;

      if (Rubrique <> CoeffRem) and (CoeffRem <> '') and (CoeffRem <> BaseRem) and (CoeffRem <> TauxRem) And (TypeCoeff <> '03') then  //PT2
      begin
        QLiens := OpenSQL('SELECT PRM_RUBRIQUE,PRM_LIBELLE FROM REMUNERATION ' +
        'WHERE PRM_RUBRIQUE="' + CoeffRem + '"', True);

        TobLesLiens := Tob.Create('Les liens', nil, -1);
        TobLesLiens.LoadDetailDB('Les liens', '', '', QLiens, False);
        Ferme(QLiens);
        If TobLesLiens.Detail.Count > 0 then
        begin
          For c := 0 to TobLesLiens.Detail.Count -1 do
          begin
            ConstruireTob;
            LibelleRubrique := TobLesLiens.Detail[c].GetValue('PRM_LIBELLE');
            T.PutValue('PCT_RUBRIQUE', Rubrique);
            T.PutValue('PCT_RUBRIQUE_LIB', Rubrique + ' ' + LibRubrique);
            T.PutValue('PCT_NATURERUB', NatureRubrique);
            T.PutValue('LIEN_RUBRIQUE_LIB', CoeffRem + ' ' + LibelleRubrique);
          end;
        end;
        FreeAndNil(TobLesLiens);
      end;
    end;
    FiniMoveProgressForm;
  end;

  TobEtat.Detail.Sort('PCT_NATURERUB;PCT_RUBRIQUE');

  {$IFDEF EAGLCLIENT}
    StPages := AglGetCriteres(Pages, FALSE);     //PT1
  {$ENDIF}
  TFQRS1(Ecran).LaTob:= TobEtat;                 //PT1

  FreeAndNil(TobRubrique); //PT1
end;

procedure TOF_PGEDITLIENRUBRIQUE.ConstruireTob();
begin
  T := Tob.Create('Les Liens rubriques', TobEtat, -1);
  T.AddChampSup('PCT_RUBRIQUE', False);
  T.AddChampSup('PCT_RUBRIQUE_LIB', False);
  T.AddChampSup('PCT_NATURERUB', False);
  T.AddChampSup('LIEN_RUBRIQUE_LIB', False);
end;

procedure TOF_PGEDITLIENRUBRIQUE.OnClose;
begin
  inherited;
  FreeAndNil (TobEtat); //PT1
end;

initialization
  registerclasses([TOF_PGEDITLIENRUBRIQUE]);
end.

