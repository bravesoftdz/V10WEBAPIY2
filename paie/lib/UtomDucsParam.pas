{
PT1 SB 29/11/2001 V563 On force la validation avant la Duplication
PT2 SB 13/12/2001 V570 Fiche de bug n° 279
                       Test code existant ne test pas bon numéro de dossier
PT3 MF 19/03/2002 V571
                       1-Correction création d'une nouvelle codification,
                         modification du contrôle si existe déjà
                       2-Correction du contrôle sur DUCSAFFECT avant suppression
                         Si suppression d'une codification CEGI D ou STD pas de
                         contrôle sur la n° de dossier dans DUCSAFFECT, sinon
                         contrôle du n° de dossier.
PT4 MF 30/04/2002 V571
                       1-Ajout du contrôle "le type de cotisation doit être
                         renseigné" sur OnUpdate
PT5 MF 23/10/2002 V585 1-Fiche qualité 10215 - il est possible de dupliquer
                         une codif CEG en STD en DOS
                       2-Quand suppression pas de contrôle sur DUCSDETAIL
PT6 MF 06/01/2003 V591
                       1- Correction des avertissements de compile
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT7  MF 22/07/2003 V_421 FQ 10765 : suppression de codification après duplication
PT8  MF 26/07/2005 V_604 FQ 12220 : contrôle qualifiant de cot. pour IRC
PT9  MF 02/02/2006 V_65  FQ 12664 : correction codif CEGID supprimable à tort
PT10 MF 13/02/2006 V_65  FQ 11793 : (Codifs. IRC)  Ajout champ Condition spéciale  de cotisation
PT11 MF 13/02/2006 V_702 FQ 13070 : Traitement des codifications ALSACE MOSELLE
PT22 PH 07/05/2007 V_72 Concept Paie

}


unit UtomDucsParam;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  eFiche, MaineAGL, UtileAGL,
  UTOB, //unused
//PT11 (compile)  HCtrls, //unused
  {$ELSE}
  db,
//unused  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche, FE_Main, DBCtrls, HDB,
//unused  Hent1,
  {$ENDIF$}
  Classes, HTB97,   HCtrls,UTOM, PgOutils, PgOutils2,    // PT11 (compile)
  HMsgBox;
//unused  HCtrls,
//unused  UTOB,
//unused  Controls;
type
  TOM_DucsParam = class(TOM)

    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnChangeField(F: TField); override;

  private
    LectureSeule, CEG, STD, DOS: Boolean;
    mode: string;
    procedure DupliquerCodification(Sender: TObject);
  end;

implementation
//unused
uses P5Def;
//============================================================================================//
//                                 PROCEDURE OnArgument
//============================================================================================//
procedure TOM_DucsParam.OnArgument;
var
  Btn: TToolBarButton97;
begin
  inherited;

  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerCodification;

  AccesPredefini('TOUS', CEG, STD, DOS);
end;

//============================================================================================//
//                                 PROCEDURE On Load Record
//============================================================================================//
procedure TOM_DucsParam.OnLoadRecord;
begin
  inherited;

  if mode = 'DUPLICATION' then exit;

  LectureSeule := FALSE;

// d PT11
  if ((GetField('PDP_CODIFICATION') > '1ZZZZZZ') or
      (GetField('PDP_CODIFICATION') < '1000000'))then
  begin
    SetControlVisible ('TPDP_CODIFALSACE',false);
    SetControlVisible ('PDP_CODIFALSACE',false);
    SetControlEnabled ('PDP_CODIFALSACE',false);
  end;
// f @@PT11

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PDP_PREDEFINI', False);
  SetControlEnabled('PDP_CODIFICATION', False);
  SetControlEnabled('BDUPLIQUER', True);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PDP_PREDEFINI', True);
    SetControlEnabled('PDP_CODIFICATION', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BDelete', False);
  end;
  SetControlVisible('BDUPLIQUER', True);
// d PT10 Condition spéciale de cotisation
  if (((GetField('PDP_CODIFICATION') >= '3') and
       (GetField('PDP_CODIFICATION') <= '5ZZZZZZ')) or
      ((GetField('PDP_CODIFICATION') >= '8') and
       (GetField('PDP_CODIFICATION') <= '9ZZZZZZ'))) then
  begin
       SetControlEnabled('PDP_CONDITION', True);
       SetControlVisible('PDP_CONDITION', True);
       SetControlVisible('TPDP_CONDITION', True);
  end
  else
  begin
       SetControlEnabled('PDP_CONDITION', False);
       SetControlVisible('PDP_CONDITION', False);
       SetControlVisible('TPDP_CONDITION', False);
  end;
// f PT10
  if (Getfield('PDP_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;

  if (Getfield('PDP_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  if (Getfield('PDP_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False);
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;

   PaieConceptPlanPaie(Ecran); // PT22
end;

//============================================================================================//
//                                 PROCEDURE On New Record
//============================================================================================//
procedure TOM_DucsParam.OnNewRecord;
begin
  inherited;

  if mode = 'DUPLICATION' then exit;
  if (CEG = TRUE) then
    SetField('PDP_PREDEFINI', 'CEG')
  else
    SetField('PDP_PREDEFINI', 'DOS');
end;

//============================================================================================//
//                                 PROCEDURE On Delete Record
//============================================================================================//
procedure TOM_DucsParam.OnDeleteRecord;
var
  //PT3-2   Champ :array[1..1] of string;
  //PT3-2  Valeur :array[1..1] of variant ;
  Champ: array[1..2] of Hstring;
  Valeur: array[1..2] of variant;

  ExisteCod: Boolean;

begin
  inherited;

  mode := 'SUPPRESSION'; // PT7
  Champ[1] := 'PDD_CODIFICATION';
  Valeur[1] := GetField('PDP_CODIFICATION');
  Champ[2] := '';
  Valeur[2] := GetField(''); //PT3-2
  ExisteCod := RechEnrAssocier('DUCSDETAIL', Champ, Valeur);
  if ExisteCod = TRUE then
  begin
    //     LastError:=1;
    //     LastErrorMsg:='Attention! Cette codification a été utilisée.#13#10'+
    //                   'Vous ne pouvez pas la supprimer!';
    PgiBox('Attention! Cette codification a été utilisée.', 'Suppression');
  end
  else
  begin
    // PT3-2
    Champ[1] := 'PDF_CODIFICATION';
    Valeur[1] := GetField('PDP_CODIFICATION');
    if ((GetField('PDP_PREDEFINI') = 'CEG') or
      (GetField('PDP_PREDEFINI') = 'STD')) then
    begin
      Champ[2] := '';
      Valeur[2] := GetField('');
    end
    else
      Champ[2] := 'PDF_NODOSSIER';
    Valeur[2] := GetField('PDP_NODOSSIER');
    // PT3-2
    ExisteCod := RechEnrAssocier('DUCSAFFECT', Champ, Valeur);
    if ExisteCod = TRUE then
    begin
      //       LastError:=1;
      //       LastErrorMsg:='Attention! Certaines rubriques de cotisation ont été affectées à cette codification.#13#10'+
      //                     'Vous ne pouvez pas la supprimer!';
      PgiBox('Attention! Certaines rubriques de cotisation ont été affectées à cette codification.', 'Suppression');

    end;
  end;
end;

//============================================================================================//
//                                 PROCEDURE On Change Field
//============================================================================================//
procedure TOM_DucsParam.OnChangeField(F: TField);
var
  //PT6-1   Champ :array[1..1] of string;
  //PT6-1   Valeur :array[1..1] of variant ;
  //PT6-1   ExisteCod : Boolean;
  //PT6-1   Pred, Codif, StrDelAffect, StrDelCodif : string;
  Pred: string;
  //PT6-1   TOB_Affect, TOB_Codif, TCodif,Affect : TOB;
  //PT6-1   QL : TQuery;
begin
  inherited;

  if mode = 'DUPLICATION' then exit;

  if (F.FieldName = 'PDP_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PDP_PREDEFINI');
    if Pred = '' then exit;
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de codification prédéfini CEGID', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PDP_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de codification prédéfini Standard', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PDP_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PDP_PREDEFINI') then SetField('PDP_PREDEFINI', pred);
  end;
// d PT10 Condition spéciale de cotisation
  if (F.FieldName = 'PDP_CODIFICATION') and (DS.State = dsinsert) then
  begin
// d PT11
    if (((GetField('PDP_CODIFICATION') >= '2') and
         (GetField('PDP_CODIFICATION') <= '5ZZZZZZ')) or
        ((GetField('PDP_CODIFICATION') >= '8') and
         (GetField('PDP_CODIFICATION') <= '9ZZZZZZ'))) then
    begin
      SetControlEnabled('PDP_CODIFALSACE', False);
      SetControlVisible('PDP_CODIFALSACE', False);
      SetControlVisible('TPDP_CODIFALSACE', False);
      SetField('PDP_CODIFALSACE', GetField('PDP_CODIFICATION'));
    end
    else
    begin
      SetControlEnabled('PDP_CODIFALSACE', true);
      SetControlVisible('PDP_CODIFALSACE', true);
      SetControlVisible('TPDP_CODIFALSACE', true);
      SetField('PDP_CODIFALSACE', GetField('PDP_CODIFICATION'));
    end;
// f PT11
    if (((GetField('PDP_CODIFICATION') >= '3') and
        (GetField('PDP_CODIFICATION') <= '5ZZZZZZ')) or
       ((GetField('PDP_CODIFICATION') >= '8') and
        (GetField('PDP_CODIFICATION') <= '9ZZZZZZ'))) then
    begin
      SetControlEnabled('PDP_CONDITION', True);
      SetControlVisible('PDP_CONDITION', True);
      SetControlVisible('TPDP_CONDITION', True);
    end
    else
    begin
      SetControlEnabled('PDP_CONDITION', False);
      SetControlVisible('PDP_CONDITION', False);
      SetControlVisible('TPDP_CONDITION',False);
    end;
  end;
// f PT10
  if (Getfield('PDP_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end;

  if (Getfield('PDP_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  if (Getfield('PDP_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False);
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;

   PaieConceptPlanPaie(Ecran); // PT22

end;

//============================================================================================//
//                                 PROCEDURE On Update Record
//============================================================================================//
procedure TOM_DucsParam.OnUpdateRecord;
var
  // PT3-1   Champ :array[1..1] of string;
  // PT3-1   Valeur :array[1..1] of variant ;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  ExisteCod: Boolean;

begin
  inherited;

  if (DS.State = dsinsert) then
  begin
    if (GetField('PDP_PREDEFINI') <> 'DOS') then
      SetField('PDP_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PDP_NODOSSIER', PgRendNoDossier());

    // Test si champ codification renseigné
    if (GetField('PDP_CODIFICATION') = '') then
    begin
      LastError := 1;
      LastErrorMsg := 'Vous devez renseigner une codification';
      SetFocusControl('PDP_CODIFICATION');
    end;

    // début PT4
    // Test si champ codification renseigné
    if (GetField('PDP_TYPECOTISATION') = '') then
    begin
      LastError := 1;
      LastErrorMsg := 'Vous devez renseigner le type de cotisation';
      SetFocusControl('PDP_TYPECOTISATION');
    end;
    // fin PT4

    // test si codification existe déjà
    Champ[1] := 'PDP_CODIFICATION';
    Valeur[1] := GetField('PDP_CODIFICATION');
    // début PT3-1
    if (GetField('PDP_PREDEFINI') <> 'DOS') then
    begin
      Champ[2] := 'PDP_NODOSSIER';
      Valeur[2] := '000000';
      Champ[3] := 'PDP_PREDEFINI';
      Valeur[3] := GetField('PDP_PREDEFINI');
    end
    else
    begin
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      Champ[2] := 'PDP_NODOSSIER';
      Valeur[2] := PgRendNoDossier();
      Champ[3] := 'PDP_PREDEFINI';
      Valeur[3] := GetField('PDP_PREDEFINI');
      //PT5-1        Champ[3]:= '';                Valeur[3]:= '';
    end;
    // fin PT3-1
    ExisteCod := RechEnrAssocier('DUCSPARAM', Champ, Valeur);
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      LastErrorMsg := 'Cette codification existe déjà';
      SetFocusControl('PDP_CODIFICATION');
    end
  end;
// d  PT8  FQ 12220
  if (DS.State = dsEdit) or (DS.State = dsinsert) then
  begin
    // IRC Seult.
    if (((GetField('PDP_CODIFICATION') >= '3') and
         (GetField('PDP_CODIFICATION') <= '5ZZZZZZ')) or
        ((GetField('PDP_CODIFICATION') >= '8') and
         (GetField('PDP_CODIFICATION') <= '9ZZZZZZ'))) then
    begin
      if (GetField('PDP_COTISQUAL') = '') or
         (GetField('PDP_COTISQUAL') = NULL) or
         (GetField('PDP_COTISQUAL') = '   ') then
      begin
        LastError := 1;
        LastErrorMsg := 'Il manque un qualifiant de cotisation';
        SetFocusControl('PDP_COTISQUAL');
      end;
    end;
  end;
// f PT8
// d PT11
 if ((IsFieldModified('PDP_CODIFALSACE')) and  (GetField('PDP_PREDEFINI')= 'CEG')) then 
    ExecuteSQL('UPDATE DUCSAFFECT SET PDF_CODIFALSACE="' +
               GetField('PDP_CODIFALSACE') +
               '" WHERE PDF_CODIFICATION = "'+
               GetField('PDP_CODIFICATION') + '"');
// f PT11
end;

//============================================================================================//
//                                 PROCEDURE Dupliquer Codification
//============================================================================================//
procedure TOM_DucsParam.DupliquerCodification(Sender: TObject);
var
  Ok: Boolean;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  NoDossier, Champlib, ChamplibS, ChampArrB, ChampArrM, ChampTyp, ChampAlsace: string; //PT11
  {$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
  {$ELSE}
  Code: THEdit;
  {$ENDIF}

begin
  TFFiche(Ecran).BValider.Click; //PT1
  ChampLib := GetField('PDP_LIBELLE');
  ChampLibS := GetField('PDP_LIBELLESUITE');
  ChampArrB := GetField('PDP_BASETYPARR');
  ChampArrM := GetField('PDP_MTTYPARR');
  ChampTyp := GetField('PDP_TYPECOTISATION');
  ChampAlsace := GetField('PDP_CODIFALSACE'); //PT11
  mode := 'DUPLICATION';
  AglLanceFiche('PAY', 'CODE', '', '', 'DUC; ; ;7');
  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PDP_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PDP_NODOSSIER';
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier() //PT2
    else Valeur[2] := '000000'; //PT2
    Champ[3] := 'PDP_CODIFICATION';
    Valeur[3] := PGCodeDupliquer;
    Ok := RechEnrAssocier('DUCSPARAM', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
      {$IFNDEF EAGLCLIENT}
      Code := THDBEdit(GetControl('PDP_CODIFICATION'));
      {$ELSE}
      Code := THEdit(GetControl('PDP_CODIFICATION'));
      {$ENDIF}
      if (code <> nil) then DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PDP_CODIFICATION', PGCodeDupliquer);
      SetField('PDP_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PDP_NODOSSIER', NoDossier);
      SetControlEnabled('PDP_CODIFICATION', False);
      SetControlEnabled('PDP_PREDEFINI', False);
      SetField('PDP_LIBELLE', ChampLib);
      SetField('PDP_LIBELLESUITE', ChampLibS);
      SetField('PDP_BASETYPARR', ChampArrB);
      SetField('PDP_MTTYPARR', ChampArrM);
      SetField('PDP_TYPECOTISATION', ChampTyp);
      SetField('PDP_CODIFALSACE', ChampAlsace);   // PT11
// d PT11
  if (((PGCodeDupliquer >= '2') and
       (PGCodeDupliquer <= '5ZZZZZZ')) or
      ((PGCodeDupliquer >= '8') and
       (PGCodeDupliquer <= '9ZZZZZZ'))) then
  begin
       SetControlEnabled('PDP_CODIFALSACE', False);
       SetControlVisible('PDP_CODIFALSACE', False);
       SetControlVisible('TPDP_CODIFALSACE', False);
       SetField('PDP_CODIFALSACE', PGCodeDupliquer);
  end
  else
  begin
       SetControlEnabled('PDP_CODIFALSACE', True);
       SetControlVisible('PDP_CODIFALSACE', True);
       SetControlVisible('TPDP_CODIFALSACE', True);
  end;
// f PT11

// d PT10 condition spéciale de cotisation
  if (((PGCodeDupliquer >= '3') and
       (PGCodeDupliquer <= '5ZZZZZZ')) or
      ((PGCodeDupliquer >= '8') and
       (PGCodeDupliquer <= '9ZZZZZZ'))) then
  begin
       SetControlEnabled('PDP_CONDITION', True);
       SetControlVisible('PDP_CONDITION', True);
       SetControlVisible('TPDP_CONDITION', True);
  end
  else
  begin
       SetControlEnabled('PDP_CONDITION', False);
       SetControlVisible('PDP_CONDITION', False);
       SetControlVisible('TPDP_CONDITION', False);
  end;
// f PT10
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
      //      ChargementTablette(TFFiche(Ecran).TableName,'');  //recharge tablette
    end
    else
      HShowMessage('5;Codification :;La duplication est impossible, la codification existe déjà.;W;O;O;O;;;', '', '');

  end;
    mode := ''; // PT9 FQ 12664
end;

initialization
  registerclasses([TOM_DucsParam]);

end.

