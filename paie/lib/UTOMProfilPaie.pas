{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... :   /  /
Description .. : TOM Gestion des profils de paie
Suite ........ : Attention, si on rajoute un profil dans la table salaries, il
Suite ........ : faut vérifier avant la destruction s'il n'est pas utiliser
Mots clefs ... : PAIE;PROFIL
*****************************************************************
  PT1 : 01/10/2001 V562 PH Nouveau champ PROFILREM dans EtABCOMPL et SALARIE
  PT2 SB 14/11/2001 V562 Duplication affectation des predefinis et no dossier
  PT3 SB 29/11/2001 V563 On force la validation avant la Duplication
  PT4 SB 30/11/2001 V563 Dysfonctionnement test code existant sur Duplication
  PT5 SB 13/12/2001 V570 Fiche de bug n° 279
                         Test code existant ne test pas bon numéro de dossier
  PT6 SB 08/01/2002 V571 le TypeProfil n'est plus géré ds la clé
  PT7 SB 30/01/2002 V571 Fiche de bug n°451 Chargement des tablettes associés à la table PROFILRUB
  PT8 SB 19/03/2002 V571 Fiche de bug n°10054 Champ tob incorrect en duplication
  PT9 SB 19/04/2002 V571 Fiche de bug n°369 Duplication d'un profil. Dossier en Standard :
                         L'affectation des prédéfini est erronée..
 PT10 SB 21/02/2003 V595 Duplication d'un profil : controle rubrique DOS erroné
 PT11 SB 25/04/2003 V_42 Intégration de la gestion des activités
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
 PT12 SB 19/06/2003 V595 Force validation pour rechargement contexte paie
 PT13 PH 04/11/2003 V_500 CWAS gestion de l'appel rubriqueprofil par exe au lieu du script
 PT14 SB 23/02/2003 V_50 Mise à jour date modif en duplication
 PT15 PH 02/02/2005 V_602 FQ 11336 Code profil en alphanumérique alors PREDEFINI=STD obligatoirement
 PT16 SB 10/06/2005 V_65 FQ 12246 : Rechargement des tablettes en CWAS
 PT17 SB 10/06/2005 V_65 Ergonomie CWAS
 PT18 GGS 24/01/2007 V_80 FQ 12694 Journal événements
 PT18-2 GGS 26/04/2007 V_80 Revue gestion trace
 PT19 PH 07/05/2007 V_72 Concept Paie
 PT20 FC 20/06/2007 V_72 FQ 14353 les lignes affichées doivent être filtrées en fonction de l'activité associée au profil
 PT21 FC 06/07/2007 V_72 FQ 14475 Concepts
 PT22 FC 12/09/2007 V_80 FQ 14767 Concepts
 PT23 FC 06/11/2007 V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
}
unit UTOMProfilPaie;

interface
uses StdCtrls, Controls, Classes, forms, sysutils,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fiche, Fe_Main,
{$ELSE}
  eFiche, MaineAgl, UtileAgl,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB,
   HTB97,
   P5Def,
   PAIETOM;       //PT18
type
  TOM_ProfilPaie = class(PGTOM)     //PT18 class TOM devient PGTOM
  public
    procedure OnChangeField(F: TField); override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnLoadRecord; override;
    procedure OnArgument(stArgument: string); override;
    procedure OnAfterUpdateRecord; override;
    procedure OnClose; override;

  private
    ExisteCod, LectureSeule, CEG, STD, DOS, OnFerme: Boolean;
    mode, DerniereCreate: string;
    NomChamp: array[1..2] of string;
    ValChamp: array[1..2] of variant;
    Trace: TStringList;                   //PT18
    LeStatut:TDataSetState; //PT23
    procedure DupliquerProfil(Sender: TObject);
    procedure CBAccesFicheClick(Sender: TObject);
//{$IFDEF EAGLCLIENT} //PT20
    procedure BGESTIONASSClick(Sender: TObject);
//{$ENDIF}            //PT20
    end;
implementation
uses
//{$IFDEF EAGLCLIENT} //PT20
RubriqueProfil,
//{$ENDIF}            //PT20
Pgoutils, PgOutils2;

procedure TOM_ProfilPaie.OnDeleteRecord;
var
  QQuery: TQuery;
  j: Integer;
  profil, Salarie, Etab: string;
begin
  inherited;
  {Regle : Avant de supprimer un profil on test si ce dernier est utilisé dans la table
  SALARIES et ETABCOMPL, si non on supprime aussi la gestion associée}
  profil := GetField('PPI_PROFIL');

  //  PT1 : 01/10/2001 V562 PH Nouveau champ PROFILREM dans SALARIE
  //test table SALARIES
  QQuery := OpenSql('SELECT PSA_SALARIE,PSA_PROFIL,PSA_PERIODBUL,PSA_PROFILRBS,PSA_PROFILTPS,PSA_PROFILAFP, ' +
    'PSA_PROFILAPP,PSA_PROFILMUT,PSA_PROFILPRE,PSA_PROFILTSS,PSA_PROFILCGE,PSA_PROFILCDD, ' +
    'PSA_PROFILMUL,PSA_PROFILANCIEN,PSA_PROFILFNAL,PSA_PROFILTRANS,PSA_PROFILREM FROM SALARIES', True);
  if not QQuery.EOF then QQuery.First; // PORTAGECWAS
  while not QQuery.Eof do
  begin
    for j := 1 to QQuery.FieldCount - 1 do //la valeur 0 correspond au code Salarie : à ne pas tester com profil
    begin
      if QQuery.Fields[j].asstring = profil then
      begin
        ExisteCod := True;
        Salarie := QQuery.Fields[0].asstring;
      end;
      if ExisteCod = True then Break;
    end;
    if ExisteCod = True then Break;
    QQuery.Next;
  end;
  Ferme(QQuery);
  if ExisteCod = TRUE then
  begin
    LastError := 1;
    LastErrorMsg := 'Attention! Ce profil est utilisé pour le salarié : ' + Salarie + ',' +
      '#13#10#13#10Vous ne pouvez le supprimer!';
    ExisteCod := FALSE;
  end;
  //test table ETABCOMPL SI AUCUN SALARIE AFFECTER AU PROFIL
  if LastError <> 1 then
  begin
    //  PT1 : 01/10/2001 V562 PH Nouveau champ PROFILREM dans EtABCOMPL
    QQuery := OpenSql('SELECT ETB_ETABLISSEMENT,ETB_PROFIL,ETB_PERIODBUL,ETB_PROFILRBS,ETB_PROFILAFP, ' +
      'ETB_PROFILAPP,ETB_PROFILMUT,ETB_PROFILPRE,ETB_PROFILTSS,ETB_PROFILCGE, ' +
      'ETB_PROFILANCIEN,ETB_PROFILTRANS,ETB_PROFILREM FROM ETABCOMPL', True);
    if not QQuery.EOF then QQuery.First; // PORTAGECWAS
    while not QQuery.Eof do
    begin
      for j := 1 to QQuery.FieldCount - 1 do //la valeur 0 correspond au code Etab : à ne pas tester com profil
      begin
        if QQuery.Fields[j].asstring = profil then
        begin
          ExisteCod := True;
          Etab := QQuery.Fields[0].asstring;
        end;
        if ExisteCod = True then Break;
      end;
      if ExisteCod = True then Break;
      QQuery.Next;
    end;
    Ferme(QQuery);
    if ExisteCod = TRUE then
    begin
      LastError := 1;
      LastErrorMsg := 'Attention! Ce profil est utilisé pour l''établissement : ' + Etab + ' ' +
        '#13#10#13#10Vous ne pouvez le supprimer!';
      ExisteCod := FALSE;
    end;
  end;
  //SUPPRESSION de la gestion associée
  if LastError <> 1 then
  begin
    NomChamp[1] := 'PPM_PROFIL';
    ValChamp[1] := GetField('PPI_PROFIL');
    ValChamp[2] := GetField('PPI_TYPEPROFIL');
    {ExisteCod:=RechEnrAssocier('PROFILRUB',NomChamp,ValChamp); PT2
    if ExisteCod=TRUE then }
    ExecuteSQL('DELETE FROM PROFILRUB WHERE ##PPM_PREDEFINI## PPM_PROFIL="' + ValChamp[1] + '" ');
  end;
  {$IFNDEF EAGLCLIENT}
  ChargementTablette(TFFiche(Ecran).TableName, '');
  {$ELSE}
  ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); // PT16
  {$ENDIF}
  ChargementTablette('PPM', ''); //PT7
//PT18
    Trace := TStringList.Create ;         //PT18-2
    Trace.Add('SUPPRESSION PROFIL '+GetField('PPI_PROFIL')+' '+ GetField('PPI_TYPEPROFIL'));
    CreeJnalEvt('003','082','OK',nil,nil,Trace);
    FreeAndNil (Trace);    //PT18-2  Trace.free;
end;

procedure TOM_ProfilPaie.OnChangeField(F: TField);
var
  icode: integer;
  Rubrique, TempRub: string;
  Nat, vide, Mes, Pred: string;
  OKRub: boolean;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  vide := '';
  // DEB  PT15 Rajout aussi controle sur le nbre de caractères saisis.
  if (F.FieldName = 'PPI_PROFIL') then
  begin
    Rubrique := Getfield('PPI_PROFIL');
    if Rubrique = '' then exit;
    if length(Rubrique) < 3 then
    begin
      SetField('PPI_PROFIL', '   ');
      PGIbox('Le code profil doit comporter 3 caractères', Ecran.Caption);
      SetFocusControl('PPI_PROFIL');
      exit;
    end;
    if (not isnumeric(Rubrique)) then
    begin // Prise en compte du code alphanumérique
      TempRub := Copy(Rubrique, 1, 1);
      if isnumeric(TempRub) then
      begin
        SetField('PPI_PROFIL', '   ');
        PGIbox('Votre code profil doit commencer par une lettre ', Ecran.Caption);
        SetFocusControl('PPI_PROFIL');
        exit;
      end;
      if (DS.State = dsinsert) then
      begin
        if (STD = True) then  //PT22
        begin
          SetField('PPI_PREDEFINI', 'STD');
          SetControlEnabled('PPI_PREDEFINI', FALSE);
        end
        else
        //DEB PT22
        begin
          SetField('PPI_PREDEFINI', 'DOS');
          SetField('PPI_PROFIL', '');
          PGIbox('La codification alphanumérique est réservée aux profils prédéfinis Standard', Ecran.Caption);
          SetFocusControl('PPI_PREDEFINI');
          SetControlEnabled('PPI_PREDEFINI', TRUE);
          exit;
        end;
        //FIN PT22
      end;
    end
    else
    begin
      SetControlEnabled('PPI_PREDEFINI', TRUE);
      iCode := strtoint(trim(Rubrique));
      TempRub := ColleZeroDevant(iCode, 3);
      if (DS.State = dsinsert) and (TempRub <> '') and (GetField('PPI_PREDEFINI') <> '') then
      begin
        OKRub := TestRubrique(GetField('PPI_PREDEFINI'), TempRub, 100);
        if (OkRub = False) or (rubrique = '000') then
        begin
          Mes := MesTestRubrique('PRO', GetField('PPI_PREDEFINI'), 100);
          HShowMessage('2;Code Erroné: ' + TempRub + ' ;' + Mes + ';W;O;O;;;', '', '');
          TempRub := '';
        end;
      end;
      if TempRub <> Rubrique then
      begin
        SetField('PPI_PROFIL', TempRub);
        SetFocusControl('PPI_PROFIL');
      end;
    end;
  end;
  // FIN  PT15
  if (F.FieldName = 'PPI_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Rubrique := GetField('PPI_PROFIL');
        Pred := GetField('PPI_PREDEFINI');
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de profil prédéfini CEGID', 'Accès refusé');
      Pred := ''; //PT21 'DOS'
      SetControlProperty('PPI_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de profil prédéfini Standard', 'Accès refusé');
      Pred := ''; //PT21 'DOS'
      SetControlProperty('PPI_PREDEFINI', 'Value', Pred);
    end;
    //DEB PT21
    if (Pred = 'DOS') and (DOS = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de profil prédéfini Dossier', 'Accès refusé');
      Pred := '';
      SetControlProperty('PPI_PREDEFINI', 'Value', Pred);
    end;
    //FIN PT21
    if (Trim(rubrique) <> '') and (Pred <> '') and (IsNumeric(Rubrique)) then // PT15
    begin
      OKRub := TestRubrique(Pred, rubrique, 100);
      if (OkRub = False) or (rubrique = '000') then
      begin
        Mes := MesTestRubrique('PRO', Pred, 100);
        HShowMessage('2;Code Erroné: ' + Rubrique + ' ;' + mes + ';W;O;O;;;', '', '');
        SetField('PPI_PROFIL', vide);
        if Pred <> GetField('PPI_PREDEFINI') then SetField('PPI_PREDEFINI', pred);
        SetFocusControl('PPI_PROFIL');
        exit;
      end;
    end;
    if Pred <> GetField('PPI_PREDEFINI') then SetField('PPI_PREDEFINI', pred);
  end;

  if (F.FieldName = 'PPI_TYPEPROFIL') then
  begin
    Nat := GetField('PPI_TYPEPROFIL');
    SetControlVisible('PPI_THEMEPROFIL', (Nat = 'PRT'));
    SetControlVisible('TPPI_THEMEPROFIL', (Nat = 'PRT'));
    if (Nat <> 'PRT') AND  (GetField('PPI_THEMEPROFIL')<>'') then SetField('PPI_THEMEPROFIL', ''); //PT17
    { DEB PT11 }
    SetControlEnabled('PPI_ACTIVITE', (RechDom('PGTYPEPROFIL', Nat, True) = 'ACT'));
    if (RechDom('PGTYPEPROFIL', Nat, True) <> 'ACT') and (GetField('PPI_ACTIVITE') <> '') then
      SetField('PPI_ACTIVITE', '');
    { FIN PT11 }
  end;

  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PPI_PREDEFINI', False);
    SetControlEnabled('PPI_PROFIL', False);
    SetControlEnabled('BDUPLIQUER', True);
    PaieConceptPlanPaie(Ecran); // PT19
  end;
end;


procedure TOM_ProfilPaie.OnUpdateRecord;
var
  Predef: string;
  Q: TQuery;
begin
  inherited;
  OnFerme := False;
  LeStatut := DS.State; //PT23
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PPI_PROFIL')
  else
    if (DerniereCreate = GetField('PPI_PROFIL')) then OnFerme := True; // le bug arrive on se casse !!!
  if mode = 'DUPLICATION' then exit;
  if (DS.State = dsinsert) then
  begin
    if (GetField('PPI_PREDEFINI') <> 'DOS') then
      SetField('PPI_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PPI_NODOSSIER', PgRendNodossier());
  end;

  if (GetField('PPI_TYPEPROFIL') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le type du profil';
  end;
  if (GetField('PPI_PROFIL') = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le code du profil';
  end;

  Predef := GetField('PPI_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PPI_PREDEFINI');
  end;

  //Rechargement des tablettes
  if (LastError = 0) and (Getfield('PPI_PROFIL') <> '') and (Getfield('PPI_LIBELLE') <> '') then
  begin
  {$IFNDEF EAGLCLIENT}
  ChargementTablette(TFFiche(Ecran).TableName, '');
  {$ELSE}
  ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); // PT16
  {$ENDIF}
  ChargementTablette('PPM', ''); //PT7
  end;

  if (LastError = 0) then
  begin
    Q := OpenSQL('SELECT * FROM PROFILPAIE WHERE ##PPI_PREDEFINI## PPI_PROFIL="' + GetField('PPI_PROFIL') + '"', FALSE);
    if (Q.EOF) then
    begin
      SetControlEnabled('BDUPLIQUER', True);
      SetControlEnabled('BGESTIONASS', True);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('BDelete', True);
      SetControlEnabled('PPI_PREDEFINI', False);
      SetControlEnabled('PPI_PROFIL', False);
    end;
    Ferme(Q);
  end;

  { DEB PT12 Validation inutile si pas acces modification gestion associée }
  if (CEG = False) and (Predef = 'CEG') then SetControlChecked('CBACCESFICHE', False);
  if (STD = False) and (Predef = 'STD') then SetControlChecked('CBACCESFICHE', False);
  if (DOS = False) and (Predef = 'DOS') then SetControlChecked('CBACCESFICHE', False);
  if LastError = 0 then
  begin
    SetControlChecked('CBACCESFICHE', False);
    ExecuteSql('UPDATE PROFILRUB SET PPM_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE ##PPM_PREDEFINI## PPM_PROFIL="' + GetField('PPI_PROFIL') + '"');
  end;
  { FIN PT12 }
end;

procedure TOM_ProfilPaie.OnNewRecord;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  //SetField ('PPI_NODOSSIER', V_PGI_env.nodossier);
  AccesPredefini('TOUS', CEG, STD, DOS);   //PT22
  if (STD = True) then                 //PT21
    SetField('PPI_PREDEFINI', 'STD')
  else if (DOS = True) then             //PT22
    SetField('PPI_PREDEFINI', 'DOS');
end;

procedure TOM_ProfilPaie.DupliquerProfil(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
  Nature: THDBValComboBox;
{$ELSE}
  Code: THEdit;
  Nature: THValComboBox;
{$ENDIF}
  T_Cumul, TOB_GestAssoc, T: TOB;
  i: integer;
  AncValNat, AncValRub, ChampBug: string;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  Ok: Boolean;
  St, NoDossier, Rubriquepaie, OriginePred: string;
  Q: TQuery;
begin
  TFFiche(Ecran).BValider.Click; //PT3
  mode := 'DUPLICATION';
  ChampBug := GetField('PPI_TYPEPROFIL');
  AncValNat := GetField('PPI_TYPEPROFIL');
  AncValRub := GetField('PPI_PROFIL');
  AglLanceFiche('PAY', 'CODE', '', '', 'PRO;' + AncValRub + '; ;3');
  // DEB PT19
  if (PGCodePredefini = 'DOS') AND (NOT DOS) then
  begin
    PgiInfo ('Vous n''êtes pas autorisé à créer un profil de type dossier.', Ecran.Caption);
    exit;
  end;
  // FIN PT19
  // DEB PT21
  if (PGCodePredefini = 'STD') AND (NOT STD) then
  begin
    PgiInfo ('Vous n''êtes pas autorisé à créer un profil de type standard.', Ecran.Caption);
    exit;
  end;
  // FIN PT21
  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PPI_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PPI_NODOSSIER';
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier() //PT5
    else Valeur[2] := '000000'; //PT5
    Champ[3] := 'PPI_PROFIL';
    Valeur[3] := PGCodeDupliquer; //PT4 StrToInt
    Ok := RechEnrAssocier('PROFILPAIE', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
      TOB_GestAssoc := TOB.Create('Le profil original', nil, -1);
      st := 'SELECT * FROM PROFILRUB WHERE ##PPM_PREDEFINI## PPM_PROFIL="' + AncValRub + '"'; //PT6 PPM_TYPEPROFIL ="'+ AncValNat +'" AND
      Q := OpenSql(st, TRUE);
      TOB_GestAssoc.LoadDetailDB('PROFILRUB', '', '', Q, FALSE);
      Ferme(Q);
      T_Cumul := TOB.Create('Le profil dupliqué', nil, -1);
{$IFNDEF EAGLCLIENT}
      Code := THDBEdit(GetControl('PPI_PROFIL'));
      Nature := THDBValComboBox(GetControl('PPI_TYPEPROFIL'));
{$ELSE}
      Code := THEdit(GetControl('PPI_PROFIL'));
      Nature := THValComboBox(GetControl('PPI_TYPEPROFIL'));
{$ENDIF}
      if (code <> nil) and (nature <> nil) then
        DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PPI_PROFIL', PGCodeDupliquer);
      for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
        if (TOB_GestAssoc.Detail[i].GetValue('PPM_PROFIL') = AncValRub) then //PT6 (TOB_GestAssoc.Detail[i].GetValue('PPM_TYPEPROFIL'))= AncValNat) and
        begin
          T := TOB_GestAssoc.Detail[i];
          if T <> nil then
          begin //PT2
            T.PutValue('PPM_PROFIL', PGCodeDupliquer);
            T.PutValue('PPM_PREDEFINI', PGCodePredefini);
            Rubriquepaie := T.GetValue('PPM_RUBRIQUE'); //DEB PT9
            OriginePred := '';
            if (Rubriquepaie <> '') and (Length(Rubriquepaie) = 4) then
              if (StrToInt(Rubriquepaie) > 1000) and ((Rubriquepaie[4] = '5') or (Rubriquepaie[4] = '7') or (Rubriquepaie[4] = '9')) then //PT10 pos de 4 au lieu de 2
                OriginePred := 'DOS';
            if OriginePred = 'DOS' then
            begin
              T.PutValue('PPM_PREDEFINI', 'DOS');
              // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
              T.PutValue('PPM_NODOSSIER', PgRendNoDossier());
            end
            else
              T.PutValue('PPM_NODOSSIER', '000000');
            if PGCodePredefini = 'DOS' then
              T.PutValue('PPM_NODOSSIER', PgRendNoDossier()); //FIN PT9
            { T.PutValue('PPM_NODOSSIER', '000000'); //PT8
             if T.GetValue('PPM_PREDEFINI')='DOS' then
               T.PutValue('PPM_NODOSSIER', V_PGI_env.nodossier);
             if PGCodePredefini='DOS' then
               Begin
               T.PutValue('PPM_PREDEFINI', PGCodePredefini);
               T.PutValue('PPM_NODOSSIER', V_PGI_env.nodossier);
               End;    }
          end;
        end;
      T_Cumul.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
      T_Cumul.InsertDB(nil, False);
      TOB_GestAssoc.free;
      T_Cumul.free;
      SetField('PPI_TYPEPROFIL', ChampBug);
      SetField('PPI_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PPI_NODOSSIER', NoDossier);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PPI_PREDEFINI', False);
      SetControlEnabled('PPI_PROFIL', False);
      SetControlEnabled('BDUPLIQUER', True);
      { DEB PT14 }
      SetControlChecked('CBACCESFICHE', False);
      ExecuteSql('UPDATE PROFILRUB SET PPM_DATEMODIF="' + UsTime(Now) + '" ' +
        'WHERE ##PPM_PREDEFINI## PPM_PROFIL="' + GetField('PPI_PROFIL') + '"');
      { FIN PT14 }
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
      {$IFNDEF EAGLCLIENT}
      ChargementTablette(TFFiche(Ecran).TableName, '');
      {$ELSE}
      ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); // PT16
      {$ENDIF}
      ChargementTablette('PPM', ''); //PT7
    end
    else
      HShowMessage('5;Profil :;La duplication est impossible, le profil existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;

procedure TOM_ProfilPaie.OnArgument(stArgument: string);
var
  Btn : TToolBarButton97;
//{$IFDEF EAGLCLIENT} //PT20
  BtnRP : TToolBarButton97;
//{$ENDIF}            //PT20
begin
  inherited;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerProfil;
  { DEB PT12 }
  if TCheckBox(GetControl('CBACCESFICHE')) <> nil then
    TCheckBox(GetControl('CBACCESFICHE')).OnClick := CBAccesFicheClick;
  { FIN PT12 }
//{$IFDEF EAGLCLIENT} //PT20
  BtnRP := TToolBarButton97(GetControl('BGESTIONASS'));
  if BtnRP <> nil then BtnRP.onClick := BGESTIONASSClick;
//{$ENDIF}            //PT20
//PT18-2  Trace := TStringList.Create ;         //PT18
  PaieConceptPlanPaie(Ecran); // PT19
end;

procedure TOM_ProfilPaie.OnLoadRecord;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  if mode = 'DUPLICATION' then exit;
  SetControlProperty('ACCES', 'Text', 'FALSE'); //PT12
  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PPI_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
    if CEG = True then SetControlProperty('ACCES', 'Text', 'TRUE') //PT12
    else SetControlProperty('ACCES', 'Text', 'FALSE');
  end
  else
    if (Getfield('PPI_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
    if STD = True then SetControlProperty('ACCES', 'Text', 'TRUE') //PT12
    else SetControlProperty('ACCES', 'Text', 'FALSE');
  end
  else
    if (Getfield('PPI_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False);
    PaieLectureSeule(TFFiche(Ecran), LectureSeule);
    SetControlEnabled('BDelete', True);
    if DOS = True then SetControlProperty('ACCES', 'Text', 'TRUE') //PT12
    else SetControlProperty('ACCES', 'Text', 'FALSE');
  end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PPI_PREDEFINI', False);
  SetControlEnabled('PPI_PROFIL', False);
  SetControlEnabled('BDUPLIQUER', True);

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PPI_PREDEFINI', True);
    SetControlEnabled('PPI_PROFIL', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BGESTIONASS', False);
    SetControlEnabled('BDelete', False);
  end;
  SetControlVisible('BDUPLIQUER', True);
  PaieConceptPlanPaie(Ecran); // PT19
end;

procedure TOM_ProfilPaie.OnAfterUpdateRecord;
var
    even: boolean;          //PT18

begin
  inherited;
  Trace := TStringList.Create ;         //PT18-2
  even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),TFFiche(Ecran).CodeName,TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut);  //PT18  //PT23
  FreeAndNil (Trace);  //PT18-2   Trace.free;                                                               //PT18
  if OnFerme then Ecran.Close;
end;



procedure TOM_ProfilPaie.OnClose;
begin
  inherited;
  { DEB PT12 }
  if GetControlText('CBACCESFICHE') = 'X' then
  begin
    LastError := 1;
    LastErrorMsg := '';
    PgiBox('Suite à la modification de la gestion associée, vous devez valider votre fiche.', Ecran.caption);
  end;
  { FIN PT12 }
end;

procedure TOM_ProfilPaie.CBAccesFicheClick(Sender: TObject);
begin
  { DEB PT12 disabled si acces fiche gestion associée }
  SetControlEnabled('BFirst', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BPrev', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BNext', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BLast', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('bInsert', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('bDefaire', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BDUPLIQUER', (GetControlText('CBACCESFICHE') = '-'));
  { FIN PT12 }
end;

//  PT13 PH 04/11/2003 V_500 CWAS gestion de l'appel rubriqueprofil par exe au lieu du script
procedure TOM_ProfilPaie.BGESTIONASSClick(Sender: TObject);
var
  ret: Boolean;
begin
{$IFDEF EAGLCLIENT}
  ret := ProfilPaieDetail(GetField('PPI_PROFIL'), GetField('PPI_LIBELLE'), GetField('PPI_PREDEFINI'),GetField('PPI_ACTIVITE')); // and (GetControlText ('CBACCESFICHE') = 'X'); //PT20
  if ret then
  begin
    SetField('PPI_LIBELLE', GetField('PPI_LIBELLE') + ' ');
    SetControlChecked('CBACCESFICHE', TRUE);
  end;
{$ELSE}
  //DEB PT20
  ret := ProfilPaieDetail(GetField('PPI_PROFIL'), GetField('PPI_LIBELLE'), GetField('PPI_PREDEFINI'),GetField('PPI_ACTIVITE')); // and (GetControlText ('CBACCESFICHE') = 'X');
  if ret and (GetControlText ('CBACCESFICHE') = 'X') then
  begin
    SetField('PPI_LIBELLE',GetField('PPI_LIBELLE'));
    SetControlChecked('CBACCESFICHE', TRUE);
  end;
  //FIN PT20
{$ENDIF}
end;

initialization
  registerclasses([TOM_PROFILPAIE]);
end.

