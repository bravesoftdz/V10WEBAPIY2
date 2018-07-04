{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 16/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SYNELTNAT_MUL ()
Mots clefs ... : TOF;UTOFSYNELTNAT_MUL
*****************************************************************}
{
PT1 06/06/2007 FC V_72 Récupération des les critères du MUL éléments nationaux pour ne charger que le nécessaire
PT2 15/06/2007 FC V_72 Ne pas proposer le niveau Population si elles ne sont pas gérées
}
Unit UTOFSYNELTNAT_MUL;

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
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     ed_Tools, // Barre de progression
     P5Def,StrUtils ;

Type
  TOF_UTOFSYNELTNAT_MUL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnClose                   ; override ;
  private
    DernierTypeNiveau:String;
    MaxCpt_S : String;
    CodeElt,ThemeElt,DatVal, DatVal_ : String;
    procedure ChangeXXWhere2;
    procedure ChangeValeurType(Sender : TObject);
    procedure ChangeValeurSaisie(Sender: TObject);
    procedure ChargementTablePEY;
  end ;

Implementation

procedure TOF_UTOFSYNELTNAT_MUL.OnLoad ;
var
  StWhere : String;
  ValeurSaisie : THEdit;
  TypNiveau,Convention : THValComboBox;
begin
  Inherited ;

  SetControlText('XX_WHERE', '');
  StWhere := '';                                       

  // Type d'enregistrement dans le fichier
  StWhere := StWhere + ' PEY_TYPUTI = "2" AND PEY_NODOSSIER = "' + MaxCpt_S + '" ';

  // Convention
  Convention := THValComboBox(GetControl('CONVENTION'));
  if (Convention <> nil) and (Convention.Value <> '') then
    // Tenir compte de ceux qui correspondent au critère + les éléments dynamiques qui n'ont pas de convention
    StWhere := StWhere + ' AND (PEY_CONVENTION = "' + Convention.Value + '" )';

  // Type et valeur de niveau
  TypNiveau := THValComboBox(GetControl('TYPNIV'));
  if (TypNiveau.Value = 'CEG') or (TypNiveau.Value = 'STD') or (TypNiveau.Value = 'DOS') then
    StWhere := StWhere + ' AND (PEY_TYPNIV = "' + TypNiveau.Value + '")';

  if (TypNiveau.Value = 'ETB') then
  begin
    ValeurSaisie := THEdit(getcontrol('ETABLISSEMENT'));
    StWhere := StWhere + ' AND (PEY_TYPNIV = "' + TypNiveau.Value + '")';
    if (ValeurSaisie <> nil) and (ValeurSaisie.Text <> '') then
      StWhere := StWhere + ' AND (PEY_ETABLISSEMENT = "' + ValeurSaisie.Text + '")';
  end;
  if (TypNiveau.Value = 'POP') then
  begin
    ValeurSaisie := THEdit(getcontrol('CODEPOP'));
    StWhere := StWhere + ' AND (PEY_TYPNIV = "' + TypNiveau.Value + '")';
    if (ValeurSaisie <> nil) and (ValeurSaisie.Text <> '') then
      StWhere := StWhere + ' AND (PEY_CODEPOP = "' + ValeurSaisie.Text + '")';
  end;
  if (TypNiveau.Value = 'SAL') then
  begin
    ValeurSaisie := THEdit(getcontrol('SALARIE'));
    StWhere := StWhere + ' AND (PEY_TYPNIV = "' + TypNiveau.Value + '")';
    if (ValeurSaisie <> nil) and (ValeurSaisie.Text <> '') then
      StWhere := StWhere + ' AND (PEY_SALARIE = "' + ValeurSaisie.Text + '")';
  end;

  SetControlText('XX_WHERE', StWhere);

  ChangeXXWhere2;
end ;

procedure TOF_UTOFSYNELTNAT_MUL.OnArgument (Argument : String ) ;
var
  CmbType: THValComboBox;
  ValeurSaisie:THEdit;
  LibValeur:THLabel;
  Arg : String;
begin
  Inherited ;
  //DEB PT1
  Arg := Argument;
  if (Arg <> '') then
  begin
    CodeElt := READTOKENST(Arg);
    ThemeElt := READTOKENST(Arg);
    DatVal := READTOKENST(Arg);
    DatVal_ := READTOKENST(Arg);
    SetControlText('PEY_CODEELT',CodeElt);
    SetControlText('PEY_THEMEELT',ThemeElt);
    SetControlText('PEY_DATVAL',DatVal);
    SetControlText('PEY_DATVAL_',DatVal_);
    if (CodeElt <> '') then
      SetControlEnabled('PEY_CODEELT',False);
    if (ThemeElt <> '') then
      SetControlEnabled('PEY_THEMEELT',False);
    if (DatVal <> '01/01/1900') then
      SetControlEnabled('PEY_DATVAL',False);
    if (DatVal_ <> '31/12/2099') then
      SetControlEnabled('PEY_DATVAL_',False);
  end
  else
  begin
    CodeElt := '';
    ThemeElt := '';
    DatVal := '';
    DatVal_ := '';
  end;
  //FIN PT1

  //DEB PT2
{$IFNDEF CPS1}
  //S'il existe un parametre population
  if not ExisteSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION') then
  begin
    CmbType := ThValComboBox(getcontrol('TYPNIV'));
    if CmbType <> nil then
      CmbType.Plus := ' AND CO_CODE <> "POP"';
  end;
{$ENDIF}
  //FIN PT2

  CmbType := THValComboBox(getcontrol('TYPNIV'));
  if (CmbType <> nil) then
    CmbType.OnChange := ChangeValeurType;

  ValeurSaisie := THEdit(getcontrol('VALEUR'));
  if (ValeurSaisie<>nil) then ValeurSaisie.OnChange := ChangeValeurSaisie;

  // Chargement par défaut du niveau établissement
  DernierTypeNiveau := '';
  CmbType.value := '';
  ValeurSaisie.Name := 'ETABLISSEMENT';
  ValeurSaisie.DataType := 'TTETABLISSEMENT';
  ValeurSaisie.Text := '';
  ValeurSaisie.Visible := False;
  CmbType.Enabled := True;
  ValeurSaisie.Enabled := True;
  LibValeur := THLabel(getcontrol('TVALEUR'));
  LibValeur.Caption := 'Etablissement';
  LibValeur.Visible := False;

  // Chargement de la table temporaire
  ChargementTablePEY;
end ;

procedure TOF_UTOFSYNELTNAT_MUL.ChangeValeurSaisie(Sender: TObject);
var
  ValeurSaisie:THEdit;
  LibValeur:THLabel;
  Q:TQuery;
begin
  // Afficher le libellé de la valeur qui vient d'être récupérée
  if (DernierTypeNiveau = 'ETB') then
  begin
    ValeurSaisie := THEdit(getcontrol('ETABLISSEMENT'));
    Q := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + ValeurSaisie.Text + '"', True);
  end;
  if (DernierTypeNiveau = 'SAL') then
  begin
    ValeurSaisie := THEdit(getcontrol('SALARIE'));
    Q := OpenSQL('SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="' + ValeurSaisie.Text + '"', True);
  end;
  if (DernierTypeNiveau = 'POP') then
  begin
    ValeurSaisie := THEdit(getcontrol('CODEPOP'));
    Q := OpenSQL('SELECT PPC_LIBELLE FROM ORDREPOPULATION WHERE PPC_POPULATION="' + ValeurSaisie.Text + '"', True);
  end;
  if (ValeurSaisie<>nil) then
  begin
    LibValeur := THLabel(getcontrol('LBL_VALEURNIVEAU'));
    if (LibValeur<>nil) then
      if not Q.Eof then
        LibValeur.Caption := Q.Fields[0].AsString
      else
        LibValeur.Caption := '';
  end;
end;

procedure TOF_UTOFSYNELTNAT_MUL.ChangeXXWhere2;
var
  Where2:THEdit;
  Habilitation:String;
  St:String;
  Etab:String;
  j,Longueur:integer;
  CritEtab,CritSal:String;
begin
  Habilitation := '';
  CritEtab := '';
  CritSal := '';
  Where2 := THEdit(GetControl('XX_WHERE2'));

  if Assigned(MonHabilitation) then
  begin
    St := MonHabilitation.LesEtab;
    Etab := ReadTokenSt(St);
    j := 0;
    while Etab <> '' do
    begin
      j := j + 1;
      if Etab <> '' then
      begin
        if j > 1 then CritEtab := CritEtab + ',';
        CritEtab := CritEtab + '"' + Etab + '"';
      end;
      Etab := ReadTokenSt(St);
    end;
    if j > 0 then
      CritEtab := ' OR (PEY_ETABLISSEMENT IN (' + CritEtab + ') AND PEY_TYPNIV = "ETB")';

    if (MonHabilitation.LeSQL <> '') then
      if (copy(MonHabilitation.LeSQL, 1, 3) <> 'PEY') then
      begin
        Longueur := Length(MonHabilitation.LeSQL);
        Longueur := Longueur - 2;
        CritSal := MidStr(MonHabilitation.LeSQL, 4, Longueur);
        CritSal := ' OR (PEY' + CritSal + ' AND PEY_TYPNIV = "SAL")';
      end
      else
        CritSal := 'OR (' + MonHabilitation.LeSQL + ' AND PEY_TYPNIV = "SAL")';
  end
  else
  begin
    CritEtab := ' OR PEY_TYPNIV = "ETB"';
    CritSal := ' OR PEY_TYPNIV = "SAL"';
  end;

  Habilitation := '(PEY_TYPNIV = "CEG" OR PEY_TYPNIV = "STD" OR PEY_TYPNIV = "DOS" ' + CritEtab + CritSal + ' OR PEY_TYPNIV = "POP")';

  if (Where2 <> nil) then SetControlText('XX_WHERE2', Habilitation);
end;

procedure TOF_UTOFSYNELTNAT_MUL.ChangeValeurType(Sender : TObject);
var
  TypNiveau:THValComboBox;
  Valeur:THEdit;
begin
  TypNiveau := THValComboBox(GetControl('TYPNIV'));

  if (DernierTypeNiveau <> TypNiveau.value) then
  begin
    // Changer la liste si besoin est
    if (TypNiveau.Value = 'ETB') then
    begin
      TFMul(Ecran).SetDBListe('PGSYNELTNATETB');
      TFMul(Ecran).BCherche.Click;
    end;
    if (TypNiveau.Value = 'POP') then
    begin
      TFMul(Ecran).SetDBListe('PGSYNELTNATPOP');
      TFMul(Ecran).BCherche.Click;
    end;
    if (TypNiveau.Value = 'SAL') then
    begin
      TFMul(Ecran).SetDBListe('PGSYNELTNATSAL');
      TFMul(Ecran).BCherche.Click;
    end;
    if (TypNiveau.Value = 'CEG') or (TypNiveau.Value = 'STD') or (TypNiveau.Value = 'DOS') or (TypNiveau.Value = '') then
    begin
      TFMul(Ecran).SetDBListe('PGSYNELTNAT');
      TFMul(Ecran).BCherche.Click;
    end;

    if (DernierTypeNiveau = '') then
    begin
      Valeur := THEdit(getcontrol('ETABLISSEMENT'));
      if (TypNiveau.value = 'ETB') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := 'ETB';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Etablissement');
      end;
      if (TypNiveau.value = 'SAL') then
      begin
        Valeur.Name := 'SALARIE';
        Valeur.DataType := 'PGSALARIE';
        DernierTypeNiveau := 'SAL';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Salarié');
      end;
      if (TypNiveau.value = 'POP') then
      begin
        Valeur.Name := 'CODEPOP';
        Valeur.DataType := 'PGPOPULATION';
        DernierTypeNiveau := 'POP';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Population');
      end;
      if (TypNiveau.value <> 'ETB') and (TypNiveau.value <> 'POP') and (TypNiveau.value <> 'SAL') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := '';
        Valeur.Visible := False;
        SetControlVisible('TVALEUR',False);
        SetControlText('LBL_VALEURNIVEAU','');
      end;
    end;
    if (DernierTypeNiveau = 'ETB') then
    begin
      Valeur := THEdit(getcontrol('ETABLISSEMENT'));
      if (TypNiveau.value = 'SAL') then
      begin
        Valeur.Name := 'SALARIE';
        Valeur.DataType := 'PGSALARIE';
        DernierTypeNiveau := 'SAL';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Salarié');
      end;
      if (TypNiveau.value = 'POP') then
      begin
        Valeur.Name := 'CODEPOP';
        Valeur.DataType := 'PGPOPULATION';
        DernierTypeNiveau := 'POP';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Population');
      end;
      if (TypNiveau.value <> 'ETB') and (TypNiveau.value <> 'POP') and (TypNiveau.value <> 'SAL') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := '';
        Valeur.Visible := False;
        SetControlVisible('TVALEUR',False);
        SetControlText('LBL_VALEURNIVEAU','');
      end;
    end;
    if (DernierTypeNiveau = 'SAL') then
    begin
      Valeur := THEdit(getcontrol('SALARIE'));
      if (TypNiveau.value = 'ETB') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := 'ETB';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Etablissement');
      end;
      if (TypNiveau.value = 'POP') then
      begin
        Valeur.Name := 'CODEPOP';
        Valeur.DataType := 'PGPOPULATION';
        DernierTypeNiveau := 'POP';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Population');
      end;
      if (TypNiveau.value <> 'ETB') and (TypNiveau.value <> 'POP') and (TypNiveau.value <> 'SAL') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := '';
        Valeur.Visible := False;
        SetControlVisible('TVALEUR',False);
        SetControlText('LBL_VALEURNIVEAU','');
      end;
    end;
    if (DernierTypeNiveau = 'POP') then
    begin
      Valeur := THEdit(getcontrol('CODEPOP'));
      if (TypNiveau.value = 'ETB') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := 'ETB';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Etablissement');
      end;
      if (TypNiveau.value = 'SAL') then
      begin
        Valeur.Name := 'SALARIE';
        Valeur.DataType := 'PGSALARIE';
        DernierTypeNiveau := 'SAL';
        Valeur.Visible := True;
        SetControlVisible('TVALEUR',True);
        SetControlCaption('TVALEUR','Salarié');
      end;
      if (TypNiveau.value <> 'ETB') and (TypNiveau.value <> 'POP') and (TypNiveau.value <> 'SAL') then
      begin
        Valeur.Name := 'ETABLISSEMENT';
        Valeur.DataType := 'TTETABLISSEMENT';
        DernierTypeNiveau := '';
        Valeur.Visible := False;
        SetControlVisible('TVALEUR',False);
        SetControlText('LBL_VALEURNIVEAU','');
      end;
    end;
    if (Valeur <> nil) then Valeur.Text := '';

    // Positionner le champ XX_WHERE2 qui contient l'habilitation en fonction du type de niveau sélectionné
    ChangeXXWhere2;
  end
  else
  begin
    if (TypNiveau.value = '')  then
      TFMul(Ecran).BCherche.Click;
  end;
end;

procedure TOF_UTOFSYNELTNAT_MUL.ChargementTablePEY;
var
  Q,Qt : TQuery;
  TobElt,T : Tob;
  Montant,MontantEuro : Double;
  DecimalSeparatorOld : Char;
  LibValNiveau,LibEtabliss : String;
  MaxCpt_I : int64;
  StWhere : String;
begin
  // On n'est pas au niveau du traitement d'un salarié. Gérer un code qui sera incrémenté de 1 en 1
  // et créer un nouveau type "0" qui ne contiendra que le n° comme ça on se le réserve dès le début
  Q := OpenSql('SELECT MAX(PEY_NODOSSIER) AS MAXCPT FROM PGSYNELTNAT WHERE PEY_TYPUTI = "0"', True);
  if not Q.eof then
    MaxCpt_S := Q.FindField('MAXCPT').AsString;
  Ferme(Q);
  if (MaxCpt_S = '') then
    MaxCpt_S := '0';
  MaxCpt_I := StrToInt64(MaxCpt_S);
  MaxCpt_I := MaxCpt_I + 1;
  MaxCpt_S := IntToStr(MaxCpt_I);
  ExecuteSQL('INSERT INTO PGSYNELTNAT (PEY_TYPUTI, PEY_CODEELT, PEY_DATVAL, PEY_PREDEFINI, PEY_NODOSSIER,' +
    ' PEY_CONVENTION, PEY_TYPNIV, PEY_VALNIV, PEY_LIBELLE, PEY_THEMEELT, PEY_MONTANT, PEY_MONTANTEURO, ' +
    ' PEY_MONETAIRE, PEY_ABREGE, PEY_BLOCNOTE, PEY_DECALMOIS, PEY_REGIMEALSACE, PEY_ETABLISSEMENT, ' +
    ' PEY_SALARIE, PEY_CODEPOP, PEY_LIBVALNIV)' +
    ' VALUES ("0","","' + USDATETIME(iDate1900) + '","","' + MaxCpt_S + '",' +
    ' "","","","","",0,0,' +
    ' "","","","","-","",' +
    ' "","","")');

{$IFNDEF EAGLSERVER}
  InitMoveProgressForm(nil, 'Chargement des données ...', 'Veuillez patienter SVP ...', 2, FALSE, TRUE);
  MoveCurProgressForm('1/2');
{$ENDIF}

  DecimalSeparatorOld := DecimalSeparator;
  // Forcer le séparateur de décimal à '.' car si ',' on a des pb lors de l'Insert si on a des décimales
  DecimalSeparator := '.';

  // Chargement des éléments nationaux
  //DEB PT1
  StWhere := '';
  if (CodeElt <>'') then
    StWhere := ' AND PEL_CODEELT = "' + CodeElt + '"';
  if (ThemeElt <> '') then
    StWhere := StWhere + ' AND PEL_THEMEELT = "' + ThemeElt + '"';
  if (DatVal <> '') then
    StWhere := StWhere + ' AND PEL_DATEVALIDITE >= "' + USDATETIME(StrToDateTime(DatVal)) + '"';
  if (DatVal_ <> '') then
    StWhere := StWhere + ' AND PEL_DATEVALIDITE <= "' + USDATETIME(StrToDateTime(DatVal_)) + '"';
  //FIN PT1

  Q := OpenSQL('SELECT PEL_CODEELT,PEL_DATEVALIDITE,PEL_LIBELLE,PEL_THEMEELT,PEL_MONTANT,PEL_MONTANTEURO' +
    ' ,PEL_MONETAIRE,PEL_ABREGE,PEL_BLOCNOTE,PEL_PREDEFINI,PEL_NODOSSIER,PEL_DECALMOIS,PEL_REGIMEALSACE' +
    ' ,PEL_CONVENTION FROM ELTNATIONAUX WHERE ##PEL_PREDEFINI## ' + StWhere, True); //PT1
  TobElt := Tob.Create('Les élts nationaux', nil, -1);
  TobElt.LoadDetailDB('ELTNATIONAUX', '', '', Q, False);
  Ferme(Q);
  T := TobElt.FindFirst([''], [''], False);
  while T <> nil do
  begin
    Montant := StrToFloat(T.GetValue('PEL_MONTANT'));
    MontantEuro := StrToFloat(T.GetValue('PEL_MONTANTEURO'));
    executeSQL('INSERT INTO PGSYNELTNAT ' +
      '(PEY_TYPUTI,PEY_CODEELT,PEY_DATVAL,PEY_PREDEFINI,PEY_NODOSSIER,PEY_CONVENTION ' +
      ' ,PEY_TYPNIV,PEY_VALNIV,PEY_LIBELLE,PEY_THEMEELT,PEY_MONTANT,PEY_MONTANTEURO,PEY_MONETAIRE,PEY_ABREGE ' +
      ' ,PEY_BLOCNOTE,PEY_DECALMOIS,PEY_REGIMEALSACE, PEY_ETABLISSEMENT,PEY_SALARIE,PEY_CODEPOP,PEY_LIBVALNIV)' +
      ' VALUES ' +
      '("2",'
      + '"' + T.GetValue('PEL_CODEELT') + '",'
      + '"' + UsDateTime(T.GetValue('PEL_DATEVALIDITE')) + '",'
      + '"' + T.GetValue('PEL_PREDEFINI') + '",'
      + '"' + MaxCpt_S + '",'
      + '"' + T.GetValue('PEL_CONVENTION') + '",'
      + '"' + T.GetValue('PEL_PREDEFINI') + '",'
      + '"",'
      + '"' + T.GetValue('PEL_LIBELLE')  + '",'
      + '"' + T.GetValue('PEL_THEMEELT') + '",'
      + FloatToStr(Montant) + ','
      + FloatToStr(MontantEuro) + ','
      + '"' + T.GetValue('PEL_MONETAIRE') + '",'
      + '"' + T.GetValue('PEL_ABREGE') + '",'
      + '"",'
      + '"' + T.GetValue('PEL_DECALMOIS') + '",'
      + '"' + T.GetValue('PEL_REGIMEALSACE') + '",'
      + '"",'
      + '"",'
      + '"",'
      + '"")');
    T := TobElt.FindNext([''], [''], False);
  end;
  FreeAndNil(TobElt);

{$IFNDEF EAGLSERVER}
  MoveCurProgressForm('2/2');
{$ENDIF}

  // Chargement des éléments dynamiques dossier
  //DEB PT1
  StWhere := '';
  if (CodeElt <>'') then
  begin
    if (StWhere = '') then StWhere := ' WHERE '
    else StWhere := StWhere + ' AND ';
    StWhere := StWhere + ' PED_CODEELT = "' + CodeElt + '"';
  end;
  if (ThemeElt <> '') then
  begin
    if (StWhere = '') then StWhere := ' WHERE '
    else StWhere := StWhere + ' AND ';
    StWhere := StWhere + ' PED_THEMEELT = "' + ThemeElt + '"';
  end;
  if (DatVal <> '') then
  begin
    if (StWhere = '') then StWhere := ' WHERE '
    else StWhere := StWhere + ' AND ';
    StWhere := StWhere + ' PED_DATEVALIDITE >= "' + USDATETIME(StrToDateTime(DatVal)) + '"';
  end;
  if (DatVal_ <> '') then
  begin
    if (StWhere = '') then StWhere := ' WHERE '
    else StWhere := StWhere + ' AND ';
    StWhere := StWhere + ' PED_DATEVALIDITE <= "' + USDATETIME(StrToDateTime(DatVal_)) + '"';
  end;
  //FIN PT1

  Q := OpenSQL('SELECT PED_TYPENIVEAU,PED_VALEURNIVEAU,PED_CODEELT,PED_DATEVALIDITE,PED_LIBELLE,PED_THEMEELT,PED_MONTANT,PED_MONTANTEURO' +
    ' ,PED_MONETAIRE,PED_ABREGE,PED_BLOCNOTE,PED_DECALMOIS,PED_REGIMEALSACE,PED_ETABLISSEMENT,PED_SALARIE,PED_CODEPOP' +
    ' FROM ELTNATIONDOS ' + StWhere, True); //PT1
  TobElt := Tob.Create('Les élts dyn', nil, -1);
  TobElt.LoadDetailDB('ELTNATIONDOS', '', '', Q, False);
  Ferme(Q);
  T := TobElt.FindFirst([''], [''], False);
  while T <> nil do
  begin
    MontantEuro := StrToFloat(T.GetValue('PED_MONTANTEURO'));
    Montant := StrToFloat(T.GetValue('PED_MONTANT'));
    LibEtabliss := T.GetValue('PED_ETABLISSEMENT');
    // Récupérer le libellé de la valeur du niveau
    If (T.GetValue('PED_TYPENIVEAU') = 'ETB') then
    begin
      Qt := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT = "' + T.GetValue('PED_ETABLISSEMENT') + '"'  ,True);
      if not Qt.Eof then
        LibValNiveau := Qt.FindField('ET_LIBELLE').AsString;
      Ferme(Qt);
    end;
    If (T.GetValue('PED_TYPENIVEAU') = 'POP') then
    begin
      Qt := OpenSQL('SELECT PPC_LIBELLE FROM ORDREPOPULATION WHERE PPC_POPULATION = "' + T.GetValue('PED_CODEPOP') + '"'  ,True);
      if not Qt.Eof then
        LibValNiveau := Qt.FindField('PPC_LIBELLE').AsString;
      ferme(Qt);
    end;
    If (T.GetValue('PED_TYPENIVEAU') = 'SAL') then
    begin
      Qt := OpenSQL('SELECT PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE = "' + T.GetValue('PED_SALARIE') + '"'  ,True);
      if not Qt.Eof then
        LibValNiveau := Qt.FindField('PSA_LIBELLE').AsString + ' ' + Qt.FindField('PSA_PRENOM').AsString;
        LibEtabliss := Qt.FindField('PSA_ETABLISSEMENT').AsString;
      Ferme(Qt);
    end;

    if (LibValNiveau <> '') and (Length(LibValNiveau) > 35) then
      LibValNiveau := Copy(LibValNiveau,1,35);

    executeSQL('INSERT INTO PGSYNELTNAT ' +
      '(PEY_TYPUTI,PEY_CODEELT,PEY_DATVAL,PEY_PREDEFINI,PEY_NODOSSIER,PEY_CONVENTION ' +
      ' ,PEY_TYPNIV,PEY_VALNIV,PEY_LIBELLE,PEY_THEMEELT,PEY_MONTANT,PEY_MONTANTEURO,PEY_MONETAIRE,PEY_ABREGE ' +
      ' ,PEY_BLOCNOTE,PEY_DECALMOIS,PEY_REGIMEALSACE, PEY_ETABLISSEMENT,PEY_SALARIE,PEY_CODEPOP,PEY_LIBVALNIV)' +
      ' VALUES ' +
      '("2","'
      + T.GetValue('PED_CODEELT') + '",'
      + '"' + UsDateTime(T.GetValue('PED_DATEVALIDITE')) + '",'
      + '"",'
      + '"' + MaxCpt_S + '",'
      + '"",'
      + '"' + T.GetValue('PED_TYPENIVEAU') + '",'
      + '"' + T.GetValue('PED_VALEURNIVEAU') + '",'
      + '"' + T.GetValue('PED_LIBELLE')  + '",'
      + '"' + T.GetValue('PED_THEMEELT') + '",'
      + floatToStr(Montant) + ','
      + floatToStr(MontantEuro) + ','
      + '"' + T.GetValue('PED_MONETAIRE') + '",'
      + '"' + T.GetValue('PED_ABREGE') + '",'
      + '"",'
      + '"' + T.GetValue('PED_DECALMOIS') + '",'
      + '"' + T.GetValue('PED_REGIMEALSACE') + '",'
      + '"' + LibEtabliss + '",'
      + '"' + T.GetValue('PED_SALARIE') + '",'
      + '"' + T.GetValue('PED_CODEPOP') + '",'
      + '"' + LibValNiveau + '")');
    T := TobElt.FindNext([''], [''], False);
  end;
  FreeAndNil(TobElt);

  // Remettre le séparateur de décimal de départ
  DecimalSeparator := DecimalSeparatorOld;

{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm;
{$ENDIF}
end;

procedure TOF_UTOFSYNELTNAT_MUL.OnClose;
begin
  inherited;
  ExecuteSQL('DELETE FROM PGSYNELTNAT WHERE (PEY_TYPUTI = "0" OR PEY_TYPUTI = "2")' +
    ' AND PEY_NODOSSIER = "' + MaxCpt_S + '"');
end;

Initialization
  registerclasses ( [ TOF_UTOFSYNELTNAT_MUL ] ) ;
end.
