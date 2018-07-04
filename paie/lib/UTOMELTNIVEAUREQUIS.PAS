{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ELTNIVEAUREQUIS (ELTNIVEAUREQUIS)
Mots clefs ... : TOM;ELTNIVEAUREQUIS
*****************************************************************
PT1   31/05/2007 FC V_72 : FQ 14301 Le niveau préconisé ne doit pas être supérieur au niveau maxi paramétré
PT2   14/06/2007 FC V_72 : Ne pas proposer le niveau Population si elles ne sont pas gérées
PT3   28/06/2007 FC V_72 : FQ 14495 Gestion des niveaux / messsage d'alerte
PT4   04/07/2007 FC V_72 : FQ 14475 mise en place des concepts
PT5   13/07/2007 FC V_72 : FQ 14559 suppression de messages mis en PT3
PT6   25/01/2007 FC V_80 : FQ 14722 Permettre la gestion CEG pour les niveaux préconisés
PT7   05/11/2007 FC V_80 : FQ 14908 Journal événement
}
Unit UTOMELTNIVEAUREQUIS ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,HDB,
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
     UTob,
     Dialogs,
     ParamSoc,
     PgOutils,PgOutils2,P5Def,  //PT4
     PAIETOM; //PT7

Type
  TOM_ELTNIVEAUREQUIS = Class (PGTOM) //PT7 class TOM devient PGTOM
    public
      procedure OnNewRecord; override;
      procedure OnDeleteRecord             ; override ;
      procedure OnUpdateRecord             ; override ;
      procedure OnLoadRecord               ; override ;
      procedure OnChangeField(F: TField); override;
      procedure OnArgument ( S: String )   ; override ;
      procedure OnAfterUpdateRecord; override;    //PT7
    private
      CodeElt,StPop : String;
      LectureSeule, CEG, STD, DOS : Boolean;
      Trace: TStringList;                //PT7
      DerniereCreate: string;            //PT7
      LeStatut:TDataSetState;            //PT7
    end ;

Implementation

procedure TOM_ELTNIVEAUREQUIS.OnUpdateRecord ;
var
  St : String;
  Predefini : String;
{PT5  Q : TQuery;
  Avertiss : Boolean;
  valeurPEL, valeurPED : String;}
 begin
  Inherited ;
  LastErrorMsg:='';
  
  // Le niveau maxi de perso est obligatoire
  if GetField('PNR_NIVMAXPERSO') = '' then
  begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Le niveau maximum de personnalisation est obligatoire');
    SetFocusControl('PNR_NIVMAXPERSO');
    exit;
  end;

  // Vérifier que le niveau maxi n'est pas plus générique que le niveau préconisé
  if   ((GetField('PNR_TYPENIVEAU') = 'STD') and (GetField('PNR_NIVMAXPERSO') = 'CEG'))
    or ((GetField('PNR_TYPENIVEAU') = 'DOS') and ((GetField('PNR_NIVMAXPERSO') = 'CEG') or (GetField('PNR_NIVMAXPERSO') = 'STD')))
    or ((GetField('PNR_TYPENIVEAU') = 'ETB') and ((GetField('PNR_NIVMAXPERSO') = 'CEG') or (GetField('PNR_NIVMAXPERSO') = 'STD') or (GetField('PNR_NIVMAXPERSO') = 'DOS')))
    or ((GetField('PNR_TYPENIVEAU') = 'POP') and (GetField('PNR_NIVMAXPERSO') <> 'POP') and (GetField('PNR_NIVMAXPERSO') <> 'SAL'))
    or ((GetField('PNR_TYPENIVEAU') = 'SAL') and (GetField('PNR_NIVMAXPERSO') <> 'SAL')) then
  begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Le niveau maximum de personnalisation ne peut pas être plus générique que le niveau préconisé');
    SetFocusControl('PNR_NIVMAXPERSO');
    exit;
  end;

  //DEB PT1
  if (GetField('PNR_TYPENIVEAU') = 'CEG') or (GetField('PNR_TYPENIVEAU') = 'STD') or (GetField('PNR_TYPENIVEAU') = 'DOS') then
  begin
    if (GetField('PNR_TYPENIVEAU') = 'CEG') then
      Predefini := '"CEG"';
    if (GetField('PNR_TYPENIVEAU') = 'STD') then
      Predefini := '"CEG","STD"';
    if (GetField('PNR_TYPENIVEAU') = 'DOS') then
      Predefini := '"CEG","STD","DOS"';

    St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT="' + GetField('PNR_CODEELT') + '"'
      + ' AND PEL_DATEVALIDITE<="' + UsDateTime(Date) + '"'
      + ' AND PEL_PREDEFINI IN (' + Predefini + ')';
    if ExisteSQL(St) = False then
    begin
      LastError := 2;
      LastErrorMsg := TraduireMemoire('Le niveau préconisé ne peut pas être supérieur au niveau prédéfini');
      SetFocusControl('PNR_TYPENIVEAU');
      exit;
    end;
  end;
  //FIN PT1

{PT5  //DEB PT3
  valeurPEL := '';
  valeurPED := '';
  if (GetField('PNR_TYPENIVEAU') = 'CEG') then
  begin
    valeurPEL := '"CEG"';
    if (GetField('PNR_NIVMAXPERSO') = 'STD') then
      valeurPEL := valeurPEL + ',"STD"';
    if (GetField('PNR_NIVMAXPERSO') = 'DOS') then
      valeurPEL := valeurPEL + ',"STD","DOS"';
    if (GetField('PNR_NIVMAXPERSO') = 'ETB') then
      valeurPED := '"ETB"';
    if (GetField('PNR_NIVMAXPERSO') = 'POP') then
      valeurPED := '"ETB","POP"';
    if (GetField('PNR_NIVMAXPERSO') = 'SAL') then
      valeurPED := '"ETB","POP","SAL"';
  end;

  if (GetField('PNR_TYPENIVEAU') = 'STD') then
  begin
    valeurPEL := '"STD"';
    if (GetField('PNR_NIVMAXPERSO') = 'DOS') then
      valeurPEL := valeurPEL + ',"DOS"';
    if (GetField('PNR_NIVMAXPERSO') = 'ETB') then
      valeurPED := '"ETB"';
    if (GetField('PNR_NIVMAXPERSO') = 'POP') then
      valeurPED := '"ETB","POP"';
    if (GetField('PNR_NIVMAXPERSO') = 'SAL') then
      valeurPED := '"ETB","POP","SAL"';
  end;

  if (GetField('PNR_TYPENIVEAU') = 'DOS') then
  begin
    valeurPEL := '"DOS"';
    if (GetField('PNR_NIVMAXPERSO') = 'ETB') then
      valeurPED := '"ETB"';
    if (GetField('PNR_NIVMAXPERSO') = 'POP') then
      valeurPED := '"ETB","POP"';
    if (GetField('PNR_NIVMAXPERSO') = 'SAL') then
      valeurPED := '"ETB","POP","SAL"';
  end;

  if (GetField('PNR_TYPENIVEAU') = 'ETB') then
  begin
    valeurPED := '"ETB"';
    if (GetField('PNR_NIVMAXPERSO') = 'POP') then
      valeurPED := valeurPED + ',"POP"';
    if (GetField('PNR_NIVMAXPERSO') = 'SAL') then
      valeurPED := valeurPED + ',"POP","SAL"';
  end;

  if (GetField('PNR_TYPENIVEAU') = 'POP') then
  begin
    valeurPED := '"POP"';
    if (GetField('PNR_NIVMAXPERSO') = 'SAL') then
      valeurPED := valeurPED + ',"SAL"';
  end;

  if (GetField('PNR_TYPENIVEAU') = 'SAL') then
  begin
    valeurPED := '"SAL"';
  end;

  Avertiss := False;
  if (valeurPEL <> '') then
    if ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetField('PNR_CODEELT') + '"'
      + ' AND PEL_PREDEFINI IN (' + valeurPEL + ')'
      + ' AND PEL_DATEVALIDITE <= "' + UsDateTime(Date) + '"') then
      Avertiss := True;
  if (valeurPED <> '') then
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT = "' + GetField('PNR_CODEELT') + '"'
      + ' AND PED_TYPENIVEAU IN (' + valeurPED + ')'
      + ' AND PED_DATEVALIDITE <= "' + UsDateTime(Date) + '"') then
      Avertiss := True;
  if Avertiss then
    ShowMessage(TraduireMemoire('Il existe déjà des éléments valorisés au niveau ' + GetControlText('PNR_TYPENIVEAU') + ' ou à un niveau inférieur'));


  valeurPEL := '';
  valeurPED := '';
  if (GetField('PNR_NIVMAXPERSO') = 'CEG') then
  begin
    valeurPEL := '"STD","DOS"';
    valeurPED := '"ETB","POP","SAL"';
  end;
  if (GetField('PNR_NIVMAXPERSO') = 'STD') then
  begin
    valeurPEL := '"DOS"';
    valeurPED := '"ETB","POP","SAL"';
  end;
  if (GetField('PNR_NIVMAXPERSO') = 'DOS') then
    valeurPED := '"ETB","POP","SAL"';
  if (GetField('PNR_NIVMAXPERSO') = 'ETB') then
    valeurPED := '"POP","SAL"';
  if (GetField('PNR_NIVMAXPERSO') = 'POP') then
    valeurPED := '"SAL"';

  Avertiss := False;
  if (valeurPEL <> '') then
    if ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetField('PNR_CODEELT') + '"'
      + ' AND PEL_PREDEFINI IN (' + valeurPEL + ')'
      + ' AND PEL_DATEVALIDITE <= "' + UsDateTime(Date) + '"') then
      Avertiss := True;
  if (valeurPED <> '') then
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT = "' + GetField('PNR_CODEELT') + '"'
      + ' AND PED_TYPENIVEAU IN (' + valeurPED + ')'
      + ' AND PED_DATEVALIDITE <= "' + UsDateTime(Date) + '"') then
      Avertiss := True;
  if Avertiss then
    ShowMessage(TraduireMemoire('Il existe des éléments dossier valorisés à un niveau inférieur au niveau maximum de personnalisation'));


  valeurPEL := '';
  valeurPED := '';
  if (GetField('PNR_TYPENIVEAU') = 'STD') then
    valeurPEL := '"CEG"';
  if (GetField('PNR_TYPENIVEAU') = 'DOS') then
    valeurPEL := '"CEG","STD"';
  if (GetField('PNR_TYPENIVEAU') = 'ETB') then
    valeurPEL := '"CEG","STD","DOS"';
  if (GetField('PNR_TYPENIVEAU') = 'POP') then
  begin
    valeurPEL := '"CEG","STD","DOS"';
    valeurPED := '"ETB"';
  end;
  if (GetField('PNR_TYPENIVEAU') = 'SAL') then
  begin
    valeurPEL := '"CEG","STD","DOS"';
    valeurPED := '"ETB","POP"';
  end;

  Avertiss := False;
  if (valeurPEL <> '') then
    if ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetField('PNR_CODEELT') + '"'
      + ' AND PEL_PREDEFINI IN (' + valeurPEL + ')'
      + ' AND PEL_DATEVALIDITE <= "' + UsDateTime(Date) + '"') then
      Avertiss := True;
  if (valeurPED <> '') then
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT = "' + GetField('PNR_CODEELT') + '"'
      + ' AND PED_TYPENIVEAU IN (' + valeurPED + ')'
      + ' AND PED_DATEVALIDITE <= "' + UsDateTime(Date) + '"') then
      Avertiss := True;
  if Avertiss then
    ShowMessage(TraduireMemoire('Il existe des éléments dossier valorisés à un niveau supérieur au niveau préconisé'));
  //FIN PT3
PT5}
  if (LastErrorMsg = '') then
  begin
    if (GetField('PNR_PREDEFINI') = 'DOS') then
      SetField('PNR_NODOSSIER', PgRendNoDossier())
    else
      SetField('PNR_NODOSSIER', '000000');
    setField('PNR_PREDEFINI',GetField('PNR_PREDEFINI'));
    setField('PNR_CODEELT',GetField('PNR_CODEELT'));
    setField('PNR_TYPENIVEAU',GetField('PNR_TYPENIVEAU'));
  end;

  //DEB PT7
  if (DS.State = dsinsert) then
    DerniereCreate := GetField('PNR_CODEELT');
  LeStatut := DS.State;
  //FIN PT7
end ;

procedure TOM_ELTNIVEAUREQUIS.OnLoadRecord ;
begin
  Inherited ;
{  if (DS.State in [dsInsert])then
  begin
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end;
}

  //DEB PT4
  AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1}
   STD := FALSE;
{$ENDIF}

  LectureSeule := FALSE;
  if (Getfield('PNR_PREDEFINI') = 'STD') and (STD = False) then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;
  SetControlEnabled('PNR_PREDEFINI', False);

  SetControlEnabled('BInsert', True);
  if (DS.State in [dsInsert]) then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
  end
  else DerniereCreate := ''; //PT7
  PaieConceptPlanPaie(Ecran);
  //FIN PT4
end ;

procedure TOM_ELTNIVEAUREQUIS.OnArgument ( S: String ) ;
var
{$IFNDEF EAGLCLIENT}
  Niveau: THDBValComboBox;
  NiveauMaxi: THDBValComboBox;
{$ELSE}
  Niveau: THValComboBox;
  NiveauMaxi: THValComboBox;
{$ENDIF}
  Plus:String;
  Pred : String;
  St : String;
begin
  Inherited ;
  if Pos('ACTION',S)=0 then
    CodeElt := S;

  //DEB PT2
{$IFNDEF CPS1}
  //S'il existe un parametre population
  if ExisteSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION') then
    StPop := '"POP",'
  else
    StPop := '';
{$ENDIF}
  //FIN PT2

  if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    if (CodeElt <> '') then
    begin
      Plus := ' AND CO_CODE <> "ETB" AND CO_CODE <> "POP" AND CO_CODE <> "SAL"';
      {$IFNDEF EAGLCLIENT}
      Niveau := THDBValComboBox(getcontrol('PNR_TYPENIVEAU'));
      NiveauMaxi := THDBValComboBox(getcontrol('PNR_NIVMAXPERSO'));
      {$ELSE}
      Niveau := THValComboBox(getcontrol('PNR_TYPENIVEAU'));
      NiveauMaxi := THValComboBox(getcontrol('PNR_NIVMAXPERSO'));
      {$ENDIF}

      // Personnaliser les tablettes des niveaux en fonction de ce qui est déjà saisi dans les paramètres
      St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + CodeElt + '" ' +
        ' AND PEL_PREDEFINI = "CEG"';
      Pred := '"CEG","STD","DOS"';
      if ExisteSQL(St) = False then
      begin
        St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + CodeElt + '" ' +
          ' AND PEL_PREDEFINI = "STD"';
        Pred := '"STD","DOS"';
        if ExisteSQL(St) = False then
          Pred := '"DOS"';
      end;

      if (Pred <> '') then
      begin
        if Niveau <> nil then
          Niveau.Plus := ' AND CO_CODE IN (' + Pred + ')';
        if NiveauMaxi <> nil then
          NiveauMaxi.Plus := ' AND CO_CODE IN (' + Pred + ')';
      end
      else
      begin
        if Niveau <> nil then
          Niveau.Plus := Plus;
        if NiveauMaxi <> nil then
          NiveauMaxi.Plus := Plus;
      end;
    end;
  end
  else
  begin
    // Personnaliser les tablettes des niveaux en fonction de ce qui est déjà saisi dans les paramètres
    Pred := '"CEG","STD","DOS","ETB",' + StPop + '"SAL"';
    if (CodeElt <> '') then
    begin
      St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + CodeElt + '" ' +
        ' AND PEL_PREDEFINI = "CEG"';
      if ExisteSQL(St) = False then
      begin
        St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + CodeElt + '" ' +
          ' AND PEL_PREDEFINI = "STD"';
        Pred := '"STD","DOS","ETB",' + StPop + '"SAL"';
        if ExisteSQL(St) = False then
          Pred := '"DOS","ETB",' + StPop + '"SAL"';
      end;
    end;
    if (Pred <> '') then
    begin
      {$IFNDEF EAGLCLIENT}
      Niveau := THDBValComboBox(getcontrol('PNR_TYPENIVEAU'));
      NiveauMaxi := THDBValComboBox(getcontrol('PNR_NIVMAXPERSO'));
      {$ELSE}
      Niveau := THValComboBox(getcontrol('PNR_TYPENIVEAU'));
      NiveauMaxi := THValComboBox(getcontrol('PNR_NIVMAXPERSO'));
      {$ENDIF}
      if Niveau <> nil then
        Niveau.Plus := ' AND CO_CODE IN (' + Pred + ')';
      if NiveauMaxi <> nil then
        NiveauMaxi.Plus := ' AND CO_CODE IN (' + Pred + ')';
    end;
  end;

  SetControlEnabled('PNR_PREDEFINI',False);
end ;

procedure TOM_ELTNIVEAUREQUIS.OnChangeField(F: TField);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
  Predefini:THDBValComboBox;
  Niveau: THDBValComboBox;
  NiveauMaxi: THDBValComboBox;
{$ELSE}
  Code: THEdit;
  Predefini:THValComboBox;
  Niveau: THValComboBox;
  NiveauMaxi: THValComboBox;
{$ENDIF}
  Libelle:THLabel;
  Q:TQuery;
  Pred,Plus,St : String;
begin
  inherited;
  // Récupérer le libellé de l'élément national s'il existe pour le prédéfini saisi
  if (F.FieldName = 'PNR_CODEELT') or (F.FieldName = 'PNR_PREDEFINI') then
  begin
    {$IFNDEF EAGLCLIENT}
    Code := THDBEdit(GetControl('PNR_CODEELT'));
    Predefini := THDBValComboBox(GetControl('PNR_PREDEFINI'));
    {$ELSE}
    Code := THEdit(GetControl('PNR_CODEELT'));
    Predefini := THValComboBox(GetControl('PNR_PREDEFINI'));
    {$ENDIF}
    Libelle := THLabel(GetControl('LIBELLE_ELT'));
    if (Code <> nil) and (Predefini<>nil) then
      if (Code.Text <> '') and (Predefini.value <> '') then
      begin
        Q := OpenSQL('SELECT PEL_LIBELLE FROM ELTNATIONAUX WHERE PEL_CODEELT="' + GetField('PNR_CODEELT') + '"'
        + ' AND PEL_PREDEFINI="' + GetField('PNR_PREDEFINI') + '"', True);
        if not Q.Eof then
          Libelle.Caption := Q.Fields[0].AsString
        else
          Libelle.Caption := '';
        ferme(Q);
      end
      else
        Libelle.Caption := '';

    // Personnaliser les tablettes des niveaux en fonction de ce qui est déjà saisi dans les paramètres
    if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
    begin
      if GetControlText('PNR_CODEELT') <> '' then
      begin
        Plus := ' AND CO_CODE <> "ETB" AND CO_CODE <> "POP" AND CO_CODE <> "SAL"';
        {$IFNDEF EAGLCLIENT}
        Niveau := THDBValComboBox(getcontrol('PNR_TYPENIVEAU'));
        NiveauMaxi := THDBValComboBox(getcontrol('PNR_NIVMAXPERSO'));
        {$ELSE}
        Niveau := THValComboBox(getcontrol('PNR_TYPENIVEAU'));
        NiveauMaxi := THValComboBox(getcontrol('PNR_NIVMAXPERSO'));
        {$ENDIF}

        St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetControlText('PNR_CODEELT') + '" ' +
          ' AND PEL_PREDEFINI = "CEG"';
        Pred := '"CEG","STD","DOS"';
        if ExisteSQL(St) = False then
        begin
          St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetControlText('PNR_CODEELT') + '" ' +
            ' AND PEL_PREDEFINI = "STD"';
          Pred := '"STD","DOS"';
          if ExisteSQL(St) = False then
            Pred := '"DOS"';
        end;

        if (Pred <> '') then
        begin
          if Niveau <> nil then
            Niveau.Plus := ' AND CO_CODE IN (' + Pred + ')';
          if NiveauMaxi <> nil then
            NiveauMaxi.Plus := ' AND CO_CODE IN (' + Pred + ')';
        end
        else
        begin
          if Niveau <> nil then
            Niveau.Plus := Plus;
          if NiveauMaxi <> nil then
            NiveauMaxi.Plus := Plus;
        end;
      end;
    end
    else
    begin
      Pred := '"CEG","STD","DOS","ETB",'+StPop+'"SAL"';
      if GetControlText('PNR_CODEELT') <> '' then
      begin
        // Personnaliser les tablettes des niveaux en fonction de ce qui est déjà saisi dans les paramètres
        St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetControlText('PNR_CODEELT') + '" ' +
          ' AND PEL_PREDEFINI = "CEG"';
        if ExisteSQL(St) = False then
        begin
          St := 'SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + GetControlText('PNR_CODEELT') + '" ' +
            ' AND PEL_PREDEFINI = "STD"';
          Pred := '"STD","DOS","ETB",'+StPop+'"SAL"';
          if ExisteSQL(St) = False then
            Pred := '"DOS","ETB",'+StPop+'"SAL"';
        end;
      end;
      if (Pred <> '') then
      begin
        {$IFNDEF EAGLCLIENT}
        Niveau := THDBValComboBox(getcontrol('PNR_TYPENIVEAU'));
        NiveauMaxi := THDBValComboBox(getcontrol('PNR_NIVMAXPERSO'));
        {$ELSE}
        Niveau := THValComboBox(getcontrol('PNR_TYPENIVEAU'));
        NiveauMaxi := THValComboBox(getcontrol('PNR_NIVMAXPERSO'));
        {$ENDIF}
        if Niveau <> nil then
          Niveau.Plus := ' AND CO_CODE IN (' + Pred + ')';
        if NiveauMaxi <> nil then
          NiveauMaxi.Plus := ' AND CO_CODE IN (' + Pred + ')';
      end;
    end;
  end;

  //DEB PT4
  if (F.FieldName = 'PNR_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PNR_PREDEFINI');
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1}
   STD := FALSE;
{$ENDIF}
    //DEB PT6
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de niveau préconisé prédéfini CEGID', 'Accès refusé');
      SetControlText('PNR_PREDEFINI','STD');
      SetFocusControl('PNR_PREDEFINI');
    end;
    //FIN PT6
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de niveau préconisé prédéfini Standard', 'Accès refusé');
      SetControlText('PNR_PREDEFINI',''); //PT6
      SetFocusControl('PNR_PREDEFINI');   //PT6
    end;
  end;

  if (ds.state in [dsBrowse]) then
  begin
    AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1}
   STD := FALSE;
{$ENDIF}

    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
  end;
  //FIN PT4
end;

procedure TOM_ELTNIVEAUREQUIS.OnNewRecord;
begin
  inherited;
  SetField('PNR_NODOSSIER', '000000');
  SetField('PNR_PREDEFINI','STD');
//  SetControlEnabled('PNR_PREDEFINI',False);

  if (CodeElt <> '') then
  begin
    SetField('PNR_CODEELT',CodeElt);
    SetControlEnabled('PNR_CODEELT',False);
  end;
end;

//DEB PT3
procedure TOM_ELTNIVEAUREQUIS.OnDeleteRecord;
var
  Avertiss : Boolean;
begin
  inherited;
  Avertiss := False;
  if (GetControlText('PNR_TYPENIVEAU') = 'CEG') then
  begin
    if ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PEL_PREDEFINI="CEG" OR PEL_PREDEFINI="STD" OR PEL_PREDEFINI="DOS")') then
      Avertiss := True
    else if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PED_TYPENIVEAU="ETB" OR PED_TYPENIVEAU="POP" OR PED_TYPENIVEAU="SAL")') then
      Avertiss := True;
  end;
  if (GetControlText('PNR_TYPENIVEAU') = 'STD') then
  begin
    if ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PEL_PREDEFINI="STD" OR PEL_PREDEFINI="DOS")') then
      Avertiss := True
    else if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PED_TYPENIVEAU="ETB" OR PED_TYPENIVEAU="POP" OR PED_TYPENIVEAU="SAL")') then
      Avertiss := True;
  end;
  if (GetControlText('PNR_TYPENIVEAU') = 'DOS') then
  begin
    if ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND PEL_PREDEFINI="DOS"') then
      Avertiss := True
    else if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PED_TYPENIVEAU="ETB" OR PED_TYPENIVEAU="POP" OR PED_TYPENIVEAU="SAL")') then
      Avertiss := True;
  end;
  if (GetControlText('PNR_TYPENIVEAU') = 'ETB') then
  begin
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PED_TYPENIVEAU="ETB" OR PED_TYPENIVEAU="POP" OR PED_TYPENIVEAU="SAL")') then
      Avertiss := True;
  end;
  if (GetControlText('PNR_TYPENIVEAU') = 'POP') then
  begin
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND (PED_TYPENIVEAU="POP" OR PED_TYPENIVEAU="SAL")') then
      Avertiss := True;
  end;
  if (GetControlText('PNR_TYPENIVEAU') = 'SAL') then
  begin
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_CODEELT="' + GetControlText('PNR_CODEELT') + '" AND PED_TYPENIVEAU="SAL"') then
      Avertiss := True;
  end;

  if Avertiss then
    PGIBox(TraduireMemoire('Attention, il existe des valeurs affectées au niveau préconisé ou à un niveau inférieur.'));

  //DEB PT7
  Trace := TStringList.Create ;
  Trace.Add('SUPPRESSION NIVEAU PRECONISE POUR ELEMENT '+GetField('PNR_CODEELT')+' NIVEAU '+ GetField('PNR_TYPENIVEAU'));
  CreeJnalEvt('003','095','OK',nil,nil,Trace);
  FreeAndNil (Trace);
  //FIN PT7
end;
//FIN PT3

//DEB PT7
procedure TOM_ELTNIVEAUREQUIS.OnAfterUpdateRecord;
var
  even: boolean;
  LaTable,lecode, LeLibelle : String;
begin
  LaTable := 'ELTNIVEAUREQUIS';
  LeCode := 'PNR_CODEELT';
  LeLibelle := 'PNR_TYPENIVEAU';
  Trace := TStringList.Create ;
  even := IsDifferent(dernierecreate,Latable,LeCode,LeLibelle,Trace,TFFicheListe(Ecran),LeStatut);
  FreeAndNil (Trace);
end;
//FIN PT7

Initialization
  registerclasses ( [ TOM_ELTNIVEAUREQUIS ] ) ;
end.
