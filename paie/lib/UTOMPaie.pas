{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM regroupant la gestion de différentes tables
Mots clefs ... : PAIE
*****************************************************************
PT1 SB 29/11/2001 V563 On force la validation avant la Duplication
PT2 SB 30/11/2001 V563 Dysfonctionnement test code existant sur Duplication
PT3 JL 07/12/2001 V563 Vérification date naissance enfants
PT4 SB 13/12/2001 V570 Fiche de bug n° 279
                       Test code existant ne test pas bon numéro de dossier
PT5 Ph 08/08/2002 V582 Creation TOM_DROITACCES  FQ10140
PT6 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
PT7 17/12/2002 PH V591 Affiche en entete du nom + prenom du salarie pour enfant et multi employeur
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT8 21/08/2003 PH V_421 FQ 10146 Le modèle CEGID n'est plus modifiable
PT9 11/05/2004 PH V_50  FQ 11161 Proctection des champs de la clé
PT10 04/10/2004 PH V_50 Suppression rechargement tablette
PT11 30/05/2007 FC V_72 FQ 12925 Gestion du type de lien de parenté
PT12 13/09/2007 FC V_80 FQ 14557 Concepts
}
unit UTOMPaie;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, DBGrids, Fiche, Fe_Main, FichList,
  {$ELSE}
  eFiche, MaineAgl, eFichList, UtileAGL,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97,
  ParamSoc, ParamDat,
  PgOutils,PgOutils2;

type
  TOM_CompetRessource = class(TOM)
    procedure OnUpdateRecord; override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
  end;

  TOM_EnfantSALARIE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override; // PT3
    procedure DateElipsisclick(Sender: TObject);
  end;

  TOM_MultiEmploySal = class(TOM)
    procedure OnNewRecord; override;
    procedure OnArgument(stArgument: string); override;
    procedure DateElipsisclick(Sender: TObject);
  end;

  TOM_TranchePlafond = class(TOM)
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
  end;

  TOM_GroupePaie = class(TOM)
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
  end;

  TOM_JeuEcrPaie = class(TOM)
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
  private
    mode: string;
    LectureSeule, CEG, STD, DOS: boolean;
    procedure DupliquerJeuEcrPaie(Sender: TObject);
  end;

  TOM_DROITACCES = class(TOM)
    procedure OnNewRecord; override;
  end;

implementation
uses P5Def;
{ TOM_EnfantSALARIE }

procedure TOM_EnfantSALARIE.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOM_EnfantSALARIE.OnChangeField(F: TField); //PT 3
var
  QNaissance: TQuery;
  DateNaissance: TDateTime;
begin
  inherited;

  if F.FieldName = 'PEF_DATENAISSANCE' then
  begin
    if GetField('PEF_DATENAISSANCE') <> NULL then
    begin
      if (IsValidDate(DateToStr(GetField('PEF_DATENAISSANCE')))) and (GetField('PEF_DATENAISSANCE') <> idate1900) then
      begin
        DateNaissance := IDate1900; // PORTAGE CWAS
        QNaissance := OpenSQL('SELECT PSA_DATENAISSANCE FROM SALARIES WHERE PSA_SALARIE="' + GetField('PEF_SALARIE') + '"', True);
        if not QNaissance.EOF then DateNaissance := QNaissance.FindField('PSA_DATENAISSANCE').AsDateTime; // PORTAGECWAS
        Ferme(QNaissance);
        if (DateNaissance > GetField('PEF_DATENAISSANCE')) and (GetField('PEF_TYPEPARENTAL') = '001') then //PT11
        begin
          PGIBox('La date de naissance de l''enfant est antérieure à la date de naissance du parent', 'Saisie date de naissance erronnée');
          SetFocusControl('PEF_DATENAISSANCE');
        end;
        //DEB PT11
        if (DateNaissance < GetField('PEF_DATENAISSANCE')) and (GetField('PEF_TYPEPARENTAL') <> '001') then
        begin
          PGIBox('La date de naissance de l''ascendant est postérieure à la date de naissance du parent', 'Saisie date de naissance erronnée');
          SetFocusControl('PEF_DATENAISSANCE');
        end;
        //FIN PT11
      end;
    end;
  end;
end;

procedure TOM_EnfantSALARIE.OnArgument(stArgument: string);
var
  {$IFNDEF EAGLCLIENT}
  Date: THDBEdit;
  {$ELSE}
  Date: THEdit;
  {$ENDIF}
begin
  inherited;
  {$IFNDEF EAGLCLIENT}
  Date := THDBEdit(GetControl('PEF_DATENAISSANCE'));
  {$ELSE}
  Date := THEdit(GetControl('PEF_DATENAISSANCE'));
  {$ENDIF}
  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;
  // PT7 17/12/2002 PH V591 Affiche en entete du nom + prenom du salarie pour enfant et multi employeur
  if StArgument <> '' then
  begin
    Ecran.caption := 'Personnes à charge de ' + Trim(StArgument) + ' : ';
    UpdateCaption(Ecran);
  end;
end;

procedure TOM_EnfantSALARIE.OnNewRecord;
var
  QQ: TQuery;
  IMax: integer;
begin
  inherited;
  QQ := OpenSQL('SELECT MAX(PEF_ENFANT) FROM ENFANTSALARIE WHERE PEF_SALARIE="' + GetField('PEF_SALARIE') + '"', TRUE);
  if not QQ.EOF then imax := QQ.Fields[0].AsInteger + 1 else iMax := 1;
  Ferme(QQ);
  SetField('PEF_ENFANT', IMax);
  // PT6 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
  SetField('PEF_DATENAISSANCE', IDate1900);
  SetField('PEF_TYPEPARENTAL', '001');  //PT11
end;


{ TOM_CompetRessource }
procedure TOM_CompetRessource.OnChangeField(F: TField);
var
  Compet: string;
  P: TPageControl;
  T1, T2, T3: TTabSheet;
  QQ: TQuery;
begin
  inherited;
  if (F.FieldName = 'ACR_COMPETENCE') then
  begin
    Compet := GetField('ACR_COMPETENCE');
    if Compet <> '' then
    begin
      QQ := OpenSQL('SELECT ACO_TYPECOMPETENCE FROM COMPETENCE WHERE ACO_COMPETENCE="' + Compet + '"', TRUE);
      if QQ.EOF then Compet := '' else Compet := QQ.Fields[0].AsString;
      Ferme(QQ);
    end;
    P := TPageControl(GetControl('ONGLET'));
    if P <> nil then
    begin
      P.ActivePage := P.Pages[0];
      T1 := TTabSheet(GetControl('GLANGUE'));
      if T1 <> nil then T1.TabVisible := (Compet = 'LGU');
      T2 := TTabSheet(GetControl('GPERMIS'));
      if T2 <> nil then T2.TabVisible := (Compet = 'PER');
      T3 := TTabSheet(GetControl('GDIPLOME'));
      if T3 <> nil then T3.TabVisible := (Compet = 'DIP');
      P.ActivePage := P.Pages[0];
    end;
  end;
end;

procedure TOM_CompetRessource.OnLoadRecord;
begin
  inherited;
  Ecran.Caption := 'Salarié ' + RechDom('PGSALARIE', GetField('ACR_SALARIE'), FALSE);
  UpdateCaption(Ecran);
end;

procedure TOM_CompetRessource.OnNewRecord;
var
  QQ: TQuery;
  IMax: integer;
begin
  inherited;
  QQ := OpenSQL('SELECT MAX(ACR_RANG) FROM COMPETRESSOURCE WHERE ACR_RESSOURCE="' + GetField('ACR_RESSOURCE') + '" AND ACR_SALARIE="' + GetField('ACR_SALARIE') + '"', TRUE);
  if not QQ.EOF then imax := QQ.Fields[0].AsInteger + 1 else iMax := 1;
  Ferme(QQ);
  SetField('ACR_RANG', IMax);
  SetField('ACR_DATEDEBUT', Date);
  SetField('ACR_DATEFIN', Date);
end;

procedure TOM_CompetRessource.OnUpdateRecord;
var
  Compet: string;
begin
  inherited;
  Compet := GetField('ACR_COMPETENCE');
  if (Compet = '') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner une compétence';
  end;
end;

{ TOM_MultiEmploySal }
procedure TOM_MultiEmploySal.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOM_MultiEmploySal.OnArgument(stArgument: string);
var
  {$IFNDEF EAGLCLIENT}
  Date: THDBEdit;
  {$ELSE}
  Date: THEdit;
  {$ENDIF}
begin
  inherited;
  {$IFNDEF EAGLCLIENT}
  Date := THDBEdit(GetControl('PML_DATEDEBUT'));
  {$ELSE}
  Date := THEdit(GetControl('PML_DATEDEBUT'));
  {$ENDIF}
  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;
  {$IFNDEF EAGLCLIENT}
  Date := THDBEdit(GetControl('PML_DATEFIN'));
  {$ELSE}
  Date := THEdit(GetControl('PML_DATEFIN'));
  {$ENDIF}
  if Date <> nil then Date.OnElipsisClick := DateElipsisclick;
  // PT7 17/12/2002 PH V591 Affiche en entete du nom + prenom du salarie pour enfant et multi employeur
  if StArgument <> '' then
  begin
    Ecran.caption := 'Employeurs de ' + Trim(StArgument) + ' : ';
    UpdateCaption(Ecran);
  end;
end;

procedure TOM_MultiEmploySal.OnNewRecord;
var
  QQ: TQuery;
  IMax: integer;
begin
  inherited;
  QQ := OpenSQL('SELECT MAX(PML_ORDRE) FROM MULTIEMPLOYSAL WHERE PML_SALARIE="' + GetField('PML_SALARIE') + '"', TRUE);
  if not QQ.EOF then imax := QQ.Fields[0].AsInteger + 1 else iMax := 1;
  Ferme(QQ);
  SetField('PML_ORDRE', IMax);
  //PT6 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
  SetField('PML_DATEDEBUT', IDate1900);
  SetField('PML_DATEFIN', IDate1900);
end;

{ TOM_TranchePlafond }

procedure TOM_TranchePlafond.OnChangeField(F: TField);
var
  {$IFNDEF EAGLCLIENT}
  BASEINF, BASESUP: THDBEdit;
  {$ELSE}
  BASEINF, BASESUP: THEdit;
  {$ENDIF}
begin
  {$IFNDEF EAGLCLIENT}
  BASEINF := THDBEdit(GetControl('PTA_TYPEBASEINF'));
  BASESUP := THDBEdit(GetControl('PTA_TYPEBASESUP'));
  {$ELSE}
  BASEINF := THEdit(GetControl('PTA_TYPEBASEINF'));
  BASESUP := THEdit(GetControl('PTA_TYPEBASESUP'));
  {$ENDIF}
  if (BASEINF = nil) or (BASESUP = nil) then exit;
  if (F.FieldName = 'PTA_TYPEBASEINF') then
  begin
    //  BASEINF.DisplayFormat := '';
    if (GetField('PTA_TYPEBASEINF') = 'ELT') then
    begin
      SetControlProperty('PTA_BASEINF', 'DataType', 'PGELEMENTNAT');
      SetControlProperty('PTA_BASEINF', 'ElipsisButton', TRUE);
      //   BASEINF.DisplayFormat := '';
    end;
    if (GetField('PTA_TYPEBASEINF') = 'MON') then
    begin
      SetControlProperty('PTA_BASEINF', 'DataType', '');
      SetControlProperty('PTA_BASEINF', 'ElipsisButton', FALSE);
      //   BASEINF.DisplayFormat := '#,##0.00';
    end;
  end;

  if (F.FieldName = 'PTA_TYPEBASESUP') then
  begin
    //  BASESUP.DisplayFormat := '';
    if (GetField('PTA_TYPEBASESUP') = 'ELT') then
    begin
      SetControlProperty('PTA_BASESUP', 'DataType', 'PGELEMENTNAT');
      SetControlProperty('PTA_BASESUP', 'ElipsisButton', TRUE);
      //   BASESUP.DisplayFormat := '';
    end;
    if (GetField('PTA_TYPEBASESUP') = 'MON') then
    begin
      SetControlProperty('PTA_BASESUP', 'DataType', '');
      SetControlProperty('PTA_BASESUP', 'ElipsisButton', FALSE);
      //   BASESUP.DisplayFormat := '#,##0.00';
    end;
  end;

  if (F.FieldName = 'PTA_COEFFSUP') and (GetField('PTA_COEFFSUP') = 0.00) then
  begin
    LastError := 1;
    LastErrorMsg := 'Le coefficient ne peux pas être nul !';
  end;

end;

procedure TOM_TranchePlafond.OnNewRecord;
var
  valeur: double;
begin
  valeur := 1.00;
  SetField('PTA_COEFFINF', valeur);
  SetField('PTA_COEFFSUP', valeur);
  SetField('PTA_DEPLAFONNE', '-');
end;


{ TOM_GroupePaie }//24/11/2000

procedure TOM_GroupePaie.OnDeleteRecord;
begin
  inherited;
  ChargementTablette(TFFiche(Ecran).TableName, '');
end;

procedure TOM_GroupePaie.OnUpdateRecord;
begin
  inherited;
  //Rechargement des tablettes
  if (LastError = 0) and (Getfield('PGR_GROUPEPAIE') <> '') and (Getfield('PGR_LIBELLE') <> '') then
    ChargementTablette(TFFiche(Ecran).TableName, '');
end;

{ TOM_JeuEcrPaie }//24/11/2000


procedure TOM_JeuEcrPaie.OnArgument(stArgument: string);
var
  Btn: TToolBarButton97;
begin
  inherited;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerJeuEcrPaie;
end;

procedure TOM_JeuEcrPaie.OnDeleteRecord;
var
  predefini, dossier, jeu: string;
begin
  inherited;
// PT10  ChargementTablette(TFFiche(Ecran).TableName, '');
  predefini := Getfield('PJP_PREDEFINI');
  dossier := Getfield('PJP_NODOSSIER');
  jeu := IntToStr(Getfield('PJP_NOJEU'));
  if (predefini <> '') and (dossier <> '') and (jeu <> '') then
    ExecuteSQL('DELETE FROM GUIDEECRPAIE WHERE PGC_PREDEFINI="' + predefini + '" ' +
      'AND PGC_NODOSSIER="' + dossier + '" AND PGC_JEUECR=' + jeu + '');
end;

procedure TOM_JeuEcrPaie.OnUpdateRecord;
var
  Predef: string;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  if (DS.State = dsinsert) then
  begin
    if (GetField('PJP_PREDEFINI') <> 'DOS') then
      SetField('PJP_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PJP_NODOSSIER', PgRendNodossier());
  end;
  Predef := GetField('PJP_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PJP_PREDEFINI');
  end;
  if (GetField('PJP_NOJEU') <= 0) then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner un jeu d''écriture superieur à zéro!';
    SetFocusControl('PJP_NOJEU');
  end;
  if (LastError = 0) and (Getfield('PJP_NOJEU') > 0) and (Getfield('PJP_LIBELLE') <> '') then
// PT10    ChargementTablette(TFFiche(Ecran).TableName, '');
end;

procedure TOM_JeuEcrPaie.DupliquerJeuEcrPaie(Sender: TObject);
var
  T_Cumul, TOB_GestAssoc, T: TOB;
  i: integer;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  Q: TQuery;
  st, NoDossier, jeu: string;
begin
  TFFicheListe(Ecran).BValider.Click; //PT1
  mode := 'DUPLICATION';
  jeu := IntToStr(GetField('PJP_NOJEU'));
  AglLanceFiche('PAY', 'CODE', '', '', 'JEU;' + jeu + '; ;0');
  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PJP_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PJP_NODOSSIER';
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier() //PT4
    else Valeur[2] := '000000'; //PT4
    Champ[3] := 'PJP_NOJEU';
    Valeur[3] := PGCodeDupliquer; //PT2 StrToInt
    if RechEnrAssocier('JEUECRPAIE', Champ, Valeur) = False then //Test si code existe ou non
    begin
      TOB_GestAssoc := TOB.Create('Le jeu d''écriture originale', nil, -1);
      st := 'SELECT * FROM GUIDEECRPAIE WHERE ##PGC_PREDEFINI## PGC_JEUECR=' + jeu + ''; //**////DB2
      Q := OpenSql(st, TRUE);
      TOB_GestAssoc.LoadDetailDB('GUIDEECRPAIE', '', '', Q, FALSE);
      FERME(Q);
      T_Cumul := TOB.Create('Le jeu dupliqué', nil, -1);
      DupliquerPaie(TFFicheListe(Ecran).TableName, Ecran);
      SetField('PJP_NOJEU', PGCodeDupliquer);
      for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
      begin
        if ((TOB_GestAssoc.Detail[i].GetValue('PGC_JEUECR')) = jeu) then
        begin
          T := TOB_GestAssoc.Detail[i];
          if T <> nil then
          begin
            T.PutValue('PGC_JEUECR', PGCodeDupliquer);
            T.PutValue('PGC_PREDEFINI', PGCodePredefini);
            if (PGCodePredefini <> 'DOS') then
              T.PutValue('PGC_NODOSSIER', '000000')
            else
              // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
              T.PutValue('PGC_NODOSSIER', PgRendNoDossier());
          end;
        end;
      end;
      T_Cumul.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
      T_Cumul.InsertDB(nil, False);
      TOB_GestAssoc.free;
      T_Cumul.free;
      SetField('PJP_PREDEFINI', PGCodePredefini);
      AccesFicheDupliquer(TFFicheListe(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PJP_NODOSSIER', NoDossier);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PJP_PREDEFINI', False);
      SetControlEnabled('PJP_NOJEU', False);
      SetControlEnabled('BDUPLIQUER', True);
      TFFicheListe(Ecran).Bouge(nbPost); //Force enregistrement
// PT10      ChargementTablette(TFFiche(Ecran).TableName, ''); //Recharge les tablettes
    end
    else
      HShowMessage('5;Jeu d''écriture :;La duplication est impossible, le jeu d''écriture existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;


procedure TOM_JeuEcrPaie.OnChangeField(F: TField);
var
  jeu, TempRub, Mes, Pred: string;
  OKRub, CEG, STD, DOS: Boolean;

begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  if (F.FieldName = 'PJP_NOJEU') then
  begin
    Jeu := IntToStr(Getfield('PJP_NOJEU'));
    if Jeu = '0' then exit;
    TempRub := Jeu;
    if ((isnumeric(Jeu)) and (Jeu <> '')) and (DS.State = dsinsert) and (GetField('PJP_PREDEFINI') <> '') then
    begin
      OKRub := TestRubrique(GetField('PJP_PREDEFINI'), Jeu, 0);
      if (OkRub = False) then
      begin
        Mes := MesTestRubrique('JEU', GetField('PJP_PREDEFINI'), 0);
        HShowMessage('2;Code Erroné: ' + Jeu + ' ;' + Mes + ';W;O;O;;;', '', '');
        TempRub := '0';
      end;
      if TempRub <> Jeu then
      begin
        SetControlText('PJP_NOJEU', TempRub);
        SetFocusControl('PJP_NOJEU');
      end;
    end;
  end;

  if (F.FieldName = 'PJP_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PJP_PREDEFINI');
    Jeu := (GetField('PJP_NOJEU'));
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de jeu prédéfini CEGID', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PJP_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de jeu prédéfini Standard', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PJP_PREDEFINI', 'Value', Pred);
    end;
    //DEB PT12
    if (Pred = 'DOS') and (DOS = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de jeu prédéfini Dossier', 'Accès refusé');
      if (STD = FALSE) then
        Pred := ''
      else
        Pred := 'STD';
      SetControlProperty('PJP_PREDEFINI', 'Value', Pred);
    end;
    //FIN PT12
    if (jeu <> '0') and (Pred <> '') then
    begin
      OKRub := TestRubrique(pred, Jeu, 0);
      if (OkRub = False) then
      begin
        Mes := MesTestRubrique('JEU', pred, 0);
        HShowMessage('2;Code Erroné: ' + Jeu + ' ;' + Mes + ';W;O;O;;;', '', '');
        SetField('PJP_NOJEU', '0');
        if Pred <> GetField('PJP_PREDEFINI') then SetField('PJP_PREDEFINI', pred);
        SetFocusControl('PJP_RUBRIQUE');
        exit;
      end;
    end;
    if Pred <> GetField('PJP_PREDEFINI') then SetField('PJP_PREDEFINI', pred);
  end;

end;

procedure TOM_JeuEcrPaie.OnNewRecord;
begin
  inherited;
  SetField('PJP_PREDEFINI', 'DOS');
end;

procedure TOM_JeuEcrPaie.OnLoadRecord;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  SetControlEnabled('BInsert', True);
  SetControlEnabled('BTNGUIDE', True);
  SetControlEnabled('BDUPLIQUER', True);
  SetControlEnabled('BDelete', True);
  AccesPredefini('TOUS', CEG, STD, DOS);
  if (Getfield('PJP_PREDEFINI') = 'CEG') then SetControlEnabled('BDelete', CEG);
  if (Getfield('PJP_PREDEFINI') = 'STD') then SetControlEnabled('BDelete', STD);
  if (Getfield('PJP_PREDEFINI') = 'DOS') then SetControlEnabled('BDelete', DOS);
  if DS.State in [dsInsert] then
  begin
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BTNGUIDE', False);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BDelete', False);
  end;
// PT9 11/05/2004 PH V_50  FQ 11161 Proctection des champs de la clé
  if NOT (DS.state in [dsInsert]) then
  begin
    SetControlEnabled('PJP_PREDEFINI', FALSE);
    SetControlEnabled('PJP_NOJEU', FALSE);
  end
  else
  begin
    SetControlEnabled('PJP_PREDEFINI', True);
    SetControlEnabled('PJP_NOJEU', True);
  end;
// FIN PT9
  // PT8 21/08/2003 PH V_421 FQ 10146 Le modèle CEGID n'est plus modifiable
  if (Getfield('PJP_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end
  else
    if (Getfield('PJP_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end
  else
    if (Getfield('PJP_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False); //PT12
    PaieLectureSeule(TFFiche(Ecran), (DOS = False));
    SetControlEnabled('BDelete', DOS);
  end;
  // FIN PT8
   PaieConceptPlanPaie(Ecran);
end;

{ TOM_DROITACCES }
//PT5 Ph 08/08/2002 V582 Creation TOM_DROITACCES  FQ10140
procedure TOM_DROITACCES.OnNewRecord;
begin
  inherited;
  SetField('YDA_DROITUTIL', 'X');

end;

initialization
  registerclasses([TOM_TranchePlafond, TOM_MultiEmploySal, TOM_CompetRessource, TOM_EnfantSALARIE, TOM_GroupePaie, TOM_JeuEcrPaie, TOM_DROITACCES]);
end.

