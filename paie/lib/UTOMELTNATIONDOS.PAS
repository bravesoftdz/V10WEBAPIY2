{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 09/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : ELTNATIONDOS (ELTNATIONDOS)
Mots clefs ... : TOM;ELTNATIONDOS
*****************************************************************
PT1 GGS 26/04/2007 V_80 gestion Trace
PT2 GGS 29/05/2007 V_80 gestion Trace Duplication
PT3 FC  06/06/2007 V_72 Appel pour la duplication depuis un élément national
PT4 FC  14/06/2007 V_72 : Ne pas proposer le niveau Population si elles ne sont pas gérées
PT5 FC  04/07/2007 V_72 : FQ 14475 Revue des concepts
PT6 FC  03/09/2007 V_80 : FQ 14722 prédéfini CEG dans le niveau préconisé
PT7 PH  18/10/2007 V_80 : Pas de order by ds la requete SQL en CWAS sur lookuplist
PT8 FC  15/11/2007 V_80 : FQ 14947 ne pas contrôler la valeur niveau si on vient de la fiche salarié
}
Unit UTOMELTNATIONDOS ;

Interface

uses Controls, Classes, forms, sysutils,
{$IFNDEF EAGLCLIENT}
  db, HDB, DBCtrls, Fiche, Fe_Main,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
  eFiche, MaineAgl, UtileAGL, Utob,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, HTB97,
  HDebug, PgOutils, PgOutils2,P5Def,PAIETOM,
{$IFNDEF CPS1}
  PGPOPULOUTILS,
{$ENDIF}
  LookUp,Entpaie,SaisieList,StrUtils;

Type
  TOM_ELTNATIONDOS = Class (PGTOM)
  private
    Code,CodeElt,StPop, mode, AppelSpecifique: string; // Code Element
    Libelle: string; // Libelle
    ThemeElt, DerniereCreate: string; // Theme
    Monetaire, LectureSeule, OnFerme: Boolean;
    Trace: TStringList;
    ParamTyp:String;
    ParamVal:String;
    ParamElt:String;
    ParamDat:TDateTime;
    Date1900:TDateTime;
    DupliDirecte : Boolean;
    Binsert: TToolbarButton97;
    CEG, STD, DOS : Boolean;//PT5
    procedure DupliquerEltNationaux(Sender: TObject);
    procedure ValNiveauElipsisClick(Sender : TObject);
    procedure ChargerLibelleValeurNiveau() ;
    procedure EltnatElipsisClick(Sender : TObject);
    procedure MajFiltreNiveau(CodeElt : String);
  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( Argument: String )   ; override ;
    procedure OnChangeField(F: TField); override;
  end ;

Implementation

procedure TOM_ELTNATIONDOS.OnNewRecord ;
var
  Q : TQuery;
begin
  Inherited ;
  //DEB PT3
  if DupliDirecte then
  begin
    // Préchargement des zones depuis table des éléments nationaux
    SetField('PED_CODEELT',CodeElt);
    SetField('PED_TYPENIVEAU','');
    SetField('PED_VALEURNIVEAU','');
    SetField('PED_DATECREATION',Now);
    SetField('PED_DATEMODIF',Now);
    SetField('PED_UTILISATEUR',V_PGI.User);

    Q := OpenSQL('SELECT PEL_LIBELLE,PEL_DATEVALIDITE,PEL_ABREGE,PEL_MONTANTEURO,PEL_MONTANT,PEL_THEMEELT,'
      + ' PEL_MONETAIRE,PEL_DECALMOIS,PEL_REGIMEALSACE FROM ELTNATIONAUX WHERE PEL_CODEELT = "' + CodeElt + '"'
      + ' ORDER BY PEL_DATEVALIDITE DESC',True);
    if not Q.Eof then
    begin
      SetField('PED_LIBELLE',Q.FindField('PEL_LIBELLE').AsString);
      SetField('PED_DATEVALIDITE',Q.FindField('PEL_DATEVALIDITE').AsDateTime);
      SetField('PED_ABREGE',Q.FindField('PEL_ABREGE').AsString);
      SetField('PED_MONTANTEURO',Q.FindField('PEL_MONTANTEURO').AsFloat);
      SetField('PED_MONTANT',Q.FindField('PEL_MONTANT').AsFloat);
      SetField('PED_THEMEELT',Q.FindField('PEL_THEMEELT').AsString);
      SetField('PED_MONETAIRE',Q.FindField('PEL_MONETAIRE').AsString);
      SetField('PED_DECALMOIS',Q.FindField('PEL_DECALMOIS').AsString);
      SetField('PED_REGIMEALSACE',Q.FindField('PEL_REGIMEALSACE').AsString);
    end;
    Ferme(Q);
    SetControlEnabled('PED_CODEELT',False);
  end
  //FIN PT3
  else
  begin
    SetField('PED_DATEVALIDITE', Date);
    setField('PED_DECALMOIS', 'X');
    setField('PED_MONETAIRE', 'X');
    if (AppelSpecifique = 'CREATION_SPEC') then
    begin
        if ParamTyp <> '' then
        begin
          if ParamTyp = 'ELT' then
          begin
            SetField('PED_CODEELT', ParamVal);
            SetControlEnabled('PED_CODEELT',False);
          end
          else
          begin
            SetField('PED_TYPENIVEAU', ParamTyp);
            SetField('PED_VALEURNIVEAU', ParamVal);
            SetControlEnabled('PED_TYPENIVEAU',False);
            SetControlEnabled('PED_VALEURNIVEAU',False);
              if ParamTyp = 'ETB' then
                SetControlcaption('TPED_VALEURNIVEAU','Etablissement');
              if ParamTyp = 'SAL' then
                SetControlcaption('TPED_VALEURNIVEAU','Salarié');
              if ParamTyp = 'POP' then
                SetControlcaption('TPED_VALEURNIVEAU','Population');
          end;
        end;
        if (ParamElt <> '') then
        begin
            SetField('PED_CODEELT', ParamElt);
            SetControlEnabled('PED_CODEELT',false);
        end;
        Date1900 := StrToDate('01/01/1900');
        if ParamDat <> Date1900 then
        begin
            SetField('PED_DATEVALIDITE', ParamDat);
        end;
    end;
  end;
end ;

procedure TOM_ELTNATIONDOS.OnDeleteRecord ;
begin
  Inherited ;
  Trace := TStringList.Create ;   //PT1
  Trace.Add('SUPPRESSION ELEMENT NATIONAL DOSSIER '+GetField('PED_CODEELT')+' '+ GetField('PED_LIBELLE')+' '+ GetField('PED_TYPENIVEAU')+' '+ GetField('PED_VALEURNIVEAU'));
  CreeJnalEvt('003','092','OK',nil,nil,Trace);
  FreeAndNil (Trace);  //PT1 Trace.free;
end ;

procedure TOM_ELTNATIONDOS.OnUpdateRecord ;
var
  Q : TQuery;
  sNiveau:String;
begin
  OnFerme := False;
  LastErrorMsg := '';
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PED_CODEELT')
  else
    if (DerniereCreate = GetField('PED_CODEELT')) then OnFerme := True;

  if (GetField('PED_CODEELT') = '') then
  begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Vous devez renseigner le code de l''élément dossier');
    SetFocusControl('PED_CODEELT');
    exit;
  end
  else
  begin
    Q := OpenSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE PEL_CODEELT="' + GetField('PED_CODEELT') + '"', True);
    if Q.Eof then
    begin
      LastError := 2;
      LastErrorMsg := TraduireMemoire('Le code élément saisi n''existe pas');
      SetFocusControl('PED_CODEELT');
      exit;
    end;
    ferme(Q);
  end;

  if (GetField('PED_TYPENIVEAU') = '') then
  begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Vous devez renseigner le niveau concerné de l''élément dossier');
    SetFocusControl('PED_TYPENIVEAU');
    exit;
  end;

  if (GetField('PED_VALEURNIVEAU') = '') then
  begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Vous devez renseigner la valeur niveau concerné de l''élément dossier');
    SetFocusControl('PED_VALEURNIVEAU');
    exit;
  end;

  Q := OpenSQL('SELECT PNR_TYPENIVEAU,PNR_NIVMAXPERSO FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT="' + GetField('PED_CODEELT') + '" AND ##PNR_PREDEFINI## ORDER BY PNR_PREDEFINI DESC', True);
  if not Q.Eof then
  begin
    // On ne peut valoriser qu'à un niveau inférieur ou égal au niveau préconisé
    if   ((Q.Fields[0].AsString = 'STD') and (GetField('PED_TYPENIVEAU') = 'CEG'))
      or ((Q.Fields[0].AsString = 'DOS') and ((GetField('PED_TYPENIVEAU') = 'CEG') OR (GetField('PED_TYPENIVEAU') = 'STD')))
      or ((Q.Fields[0].AsString = 'ETB') and ((GetField('PED_TYPENIVEAU') = 'CEG') OR (GetField('PED_TYPENIVEAU') = 'STD') OR (GetField('PED_TYPENIVEAU') = 'DOS')))
      or ((Q.Fields[0].AsString = 'POP') and ((GetField('PED_TYPENIVEAU') = 'CEG') OR (GetField('PED_TYPENIVEAU') = 'STD') OR (GetField('PED_TYPENIVEAU') = 'DOS') OR (GetField('PED_TYPENIVEAU') = 'ETB')))
      or ((Q.Fields[0].AsString = 'SAL') and (GetField('PED_TYPENIVEAU') <> 'SAL')) then
    begin
      LastError := 2;
      if (Q.Fields[0].AsString = 'CEG') then sNiveau := 'CEGID';
      if (Q.Fields[0].AsString = 'STD') then sNiveau := 'Standard';
      if (Q.Fields[0].AsString = 'DOS') then sNiveau := 'Dossier';
      if (Q.Fields[0].AsString = 'ETB') then sNiveau := 'Etablissement';
      if (Q.Fields[0].AsString = 'POP') then sNiveau := 'Population';
      if (Q.Fields[0].AsString = 'SAL') then sNiveau := 'Salarié';
      LastErrorMsg := TraduireMemoire('Vous ne pouvez pas choisir un niveau plus générique que le niveau préconisé ' + sNiveau);
      SetFocusControl('PED_TYPENIVEAU');
      exit;
    end;

    // On ne peut valoriser qu'à un niveau supérieur ou égal au niveau maxi de personnalisation
    if   ((Q.Fields[1].AsString = 'CEG') and (GetField('PED_TYPENIVEAU') <> 'CEG'))
      or ((Q.Fields[1].AsString = 'STD') and ((GetField('PED_TYPENIVEAU') = 'DOS') OR (GetField('PED_TYPENIVEAU') = 'ETB') OR (GetField('PED_TYPENIVEAU') = 'POP') OR (GetField('PED_TYPENIVEAU') = 'SAL')))
      or ((Q.Fields[1].AsString = 'DOS') and ((GetField('PED_TYPENIVEAU') = 'ETB') OR (GetField('PED_TYPENIVEAU') = 'POP') OR (GetField('PED_TYPENIVEAU') = 'SAL')))
      or ((Q.Fields[1].AsString = 'ETB') and ((GetField('PED_TYPENIVEAU') = 'POP') OR (GetField('PED_TYPENIVEAU') = 'SAL')))
      or ((Q.Fields[1].AsString = 'POP') and (GetField('PED_TYPENIVEAU') = 'SAL')) then
    begin
      LastError := 2;
      if (Q.Fields[1].AsString = 'CEG') then sNiveau := 'CEGID';
      if (Q.Fields[1].AsString = 'STD') then sNiveau := 'Standard';
      if (Q.Fields[1].AsString = 'DOS') then sNiveau := 'Dossier';
      if (Q.Fields[1].AsString = 'ETB') then sNiveau := 'Etablissement';
      if (Q.Fields[1].AsString = 'POP') then sNiveau := 'Population';
      if (Q.Fields[1].AsString = 'SAL') then sNiveau := 'Salarié';
      LastErrorMsg := TraduireMemoire('Vous ne pouvez pas choisir un niveau plus fin que le niveau maximum de personnalisation ' + sNiveau);
      SetFocusControl('PED_TYPENIVEAU');
      exit;
    end;
  end;
  Ferme(Q);

  // Vérifier que la valeur du niveau correspond au type du niveau : ETB, SAL ou POP
  if GetControlEnabled('PED_VALEURNIVEAU') and (GetControlText('PED_VALEURNIVEAU') <> '') then   //PT8
  begin
    if GetField('PED_TYPENIVEAU') = 'ETB' then
      Q := OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS WHERE ET_ETABLISSEMENT="' + GetField('PED_VALEURNIVEAU') + '"', True);
    if GetField('PED_TYPENIVEAU') = 'SAL' then
      Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="' + GetField('PED_VALEURNIVEAU') + '"', True);
    if GetField('PED_TYPENIVEAU') = 'POP' then
      Q := OpenSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION WHERE PPC_POPULATION="' + GetField('PED_VALEURNIVEAU') + '"', True);

    if (GetField('PED_TYPENIVEAU') = 'ETB') or (GetField('PED_TYPENIVEAU') = 'SAL') or (GetField('PED_TYPENIVEAU') = 'POP') then
    begin
      if Q.Eof then
      begin
        LastError := 2;
        LastErrorMsg := TraduireMemoire('La valeur niveau concerné ne correspond pas au niveau de l''élément dossier');
        SetFocusControl('PED_VALEURNIVEAU');
        exit;
      end;
    end;
    ferme(Q);
  end;

  Date1900 := StrToDate('01/01/1900');
  if (ParamDat <> Date1900) and (GetField('PED_DATEVALIDITE') > ParamDat) then
  begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('La date de validité doit être inférieure ou égale au ' + DateToStr(ParamDat));
    SetFocusControl('PED_DATEVALIDITE');
    exit;
  end;

  if (LastErrorMsg = '') then
  begin
    setField('PED_BLOCNOTE',' ');

    // En fonction du type de niveau, mettre à jour la zone concernée
    setField('PED_ETABLISSEMENT','');
    setField('PED_SALARIE','');
    setField('PED_CODEPOP','');
    if (GetField('PED_TYPENIVEAU') = 'ETB') then
      setField('PED_ETABLISSEMENT',GetField('PED_VALEURNIVEAU'));
    if (GetField('PED_TYPENIVEAU') = 'SAL') then
      setField('PED_SALARIE',GetField('PED_VALEURNIVEAU'));
    if (GetField('PED_TYPENIVEAU') = 'POP') then
      setField('PED_CODEPOP',GetField('PED_VALEURNIVEAU'));
  end;

  // Si la duplication depuis un élément national a aboutit, le signaler en retour
  if DupliDirecte then
  begin
    if ExisteSQL('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_TYPENIVEAU = "' + GetField('PED_TYPENIVEAU') + '"' +
      ' AND PED_VALEURNIVEAU = "' + GetField('PED_VALEURNIVEAU') + '"' +
      ' AND PED_CODEELT = "' + GetField('PED_CODEELT') + '"' +
      ' AND PED_DATEVALIDITE = "' + USDATETIME(GetField('PED_DATEVALIDITE')) + '"') then
      TFFiche(Ecran).Retour := 'KO'
    else
      TFFiche(Ecran).Retour := 'OK';
  end;
end;

procedure TOM_ELTNATIONDOS.OnAfterUpdateRecord ;
var
  even: boolean;
begin
  inherited;
  Trace := TStringList.Create ;   //PT1
  if not (Ecran is TFSaisieList) then
    even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),'PED_CODEELT',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran))
  else
    even := IsDifferent(dernierecreate,'PED','PED_CODEELT','PED_LIBELLE',Trace,TFSaisieList(Ecran));
  FreeAndNil(Trace);              //PT1  Trace.Free;
  if OnFerme then Ecran.Close;
end ;

procedure TOM_ELTNATIONDOS.OnLoadRecord ;
var
  okok: boolean;
begin
  if DupliDirecte then exit;

  //DEB PT5
  Code := string(GetField('PED_CODEELT'));
  LectureSeule := False;
  if Code <> '' then
  begin
    AccesPredefini('TOUS', CEG, STD, DOS);
  {$IFDEF CPS1}
     STD := FALSE;
  {$ENDIF}
    if (DOS = FALSE) then // or (not ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT = "' + Code + '" AND PNR_NIVMAXPERSO IN ("ETB","POP","SAL") AND ##PNR_PREDEFINI##')) then
      LectureSeule := True;
  end;
  //FIN PT5

  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  if LectureSeule = False then
    MajFiltreNiveau(Code);
  Libelle := string(GetField('PED_LIBELLE'));
  ThemeElt := GetField('PED_THEMEELT');
  Monetaire := (GetField('PED_MONETAIRE') = 'X');
  okok := (DS <> nil) and (DS.State = dsInsert);
  if GetField('PED_CODEELT') = '' then
    SetControlEnabled('PED_CODEELT',okok);
  SetControlEnabled('PED_DATEVALIDITE',okok);

  SetControlEnabled('PED_DATEVALIDITE', False);
  SetControlEnabled('PED_CODEELT', False);

  if (DS.State in [dsInsert]) and (Mode <> 'DUPLICATION') then
  begin
    LectureSeule := FALSE;
    if GetField('PED_CODEELT') = '' then
      SetControlEnabled('PED_CODEELT', True);
    SetControlEnabled('PED_DATEVALIDITE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BDelete', False);
    if (AppelSpecifique = 'CREATION_SPEC') then
      ChargerLibelleValeurNiveau;
  end
  else
  begin
    if (DOS = True) and (LectureSeule = False) then
    begin
      SetControlEnabled('BDUPLIQUER', True);
      SetControlVisible('BDUPLIQUER', True);
    end;
    ChargerLibelleValeurNiveau;
  end;

  //DEB PT5
  if (Code <> '') and (LectureSeule = True) then
  begin
    PaieLectureSeule(TFFiche(Ecran), True);
    SetControlEnabled('BDUPLIQUER',False);
  end;
  //FIN PT5
end;

procedure TOM_ELTNATIONDOS.ChargerLibelleValeurNiveau() ;
var
  Q : TQuery;
  LibelleValeur:THLabel;
  LibLibValeur:THLabel;
begin
  // Charger le libellé de la valeur
  LibelleValeur := THLabel(GetControl('LBL_VALEURNIVEAU'));
  LibLibValeur := THLabel(GetControl('TPED_VALEURNIVEAU'));
  if GetField('PED_TYPENIVEAU') = 'ETB' then
  begin
    Q := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + GetField('PED_VALEURNIVEAU') + '"', True);
    LibLibValeur.Caption := 'Etablissement';
  end;
  if GetField('PED_TYPENIVEAU') = 'SAL' then
  begin
    Q := OpenSQL('SELECT PSA_LIBELLE||" "||PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="' + GetField('PED_VALEURNIVEAU') + '"', True);
    LibLibValeur.Caption := 'Salarié';
  end;
  if GetField('PED_TYPENIVEAU') = 'POP' then
  begin
    Q := OpenSQL('SELECT PPC_LIBELLE FROM ORDREPOPULATION WHERE PPC_POPULATION="' + GetField('PED_VALEURNIVEAU') + '"', True);
    LibLibValeur.Caption := 'Population';
  end;
  if (GetField('PED_TYPENIVEAU') = 'SAL') or (GetField('PED_TYPENIVEAU') ='ETB') or (GetField('PED_TYPENIVEAU') = 'POP') then
  begin
    if not Q.Eof then
      LibelleValeur.Caption := Q.Fields[0].AsString
    else
      LibelleValeur.Caption := '';
      ferme(Q);
  end
  else
  LibelleValeur.Caption := '';
end;

procedure TOM_ELTNATIONDOS.OnArgument ( Argument: String ) ;
var
  Arg:String;
  Action:String;
  Btn: TToolBarButton97;
{$IFNDEF EAGLCLIENT}
  ValNiveau,Element: THDBEdit;
{$ELSE}
  ValNiveau,Element: THEdit;
{$ENDIF}
begin
  Inherited ;
  Arg := Argument;
  Action := READTOKENST(Arg);
  //DEB PT3
  if Pos('DUPLI',Argument)>0 then
  begin
    CodeElt := READTOKENST(Arg);
    ParamDat := StrToDate('01/01/1900');
  end
  //FIN PT3
  else
  begin
    ParamTyp := '';
    ParamVal := '';
    ParamElt := '';
    ParamDat := StrToDate('01/01/1900');

    AppelSpecifique := '';
    if Arg <> '' then
    begin
      // Quand l'appel vient d'un établissement, d'un salarié ou d'une population sont renseignés : ParamTyp, ParamVal
      // Quand l'appel vient du calcul d'un bulletin, sont renseignés : ParamTyp, ParamVal, ParamElt, ParamDat
      ParamTyp := READTOKENST(Arg);
      ParamVal := READTOKENST(Arg);
      ParamElt := READTOKENST(Arg);
      if Arg <> '' then
        ParamDat := StrToDate(READTOKENST(Arg));
      AppelSpecifique := 'CREATION_SPEC';
    end;
  end;

  //DEB PT4
{$IFNDEF CPS1}
  //S'il existe un parametre population
  if ExisteSQL('SELECT PPC_POPULATION FROM ORDREPOPULATION') then
    StPop := '"POP"'
  else
    StPop := '';
{$ENDIF}
  //FIN PT4

  //PT1   Trace := TStringList.Create ;
  if not (Ecran is TFSaisieList) then
  begin
    Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
    if btn <> nil then Btn.OnClick := DupliquerEltNationaux;
   {$IFNDEF EAGLCLIENT}
    ValNiveau := THDBEdit(GetControl('PED_VALEURNIVEAU'));
  {$ELSE}
    ValNiveau := THEdit(GetControl('PED_VALEURNIVEAU'));
  {$ENDIF}
  If ValNiveau <> Nil then
     ValNiveau.OnElipsisClick:= ValNiveauElipsisClick;
        {$IFNDEF EAGLCLIENT}
    Element := THDBEdit(GetControl('PED_CODEELT'));
  {$ELSE}
    Element := THEdit(GetControl('PED_CODEELT'));
  {$ENDIF}
  If Element <> Nil then
     Element.OnElipsisClick:= EltnatElipsisClick;
   end;

  //DEB PT3
  DupliDirecte := False;
  if Pos('DUPLI',Argument)>0 then
  begin
    DupliDirecte := True;
    SetControlEnabled('BDUPLIQUER',False);
    BInsert := TToolbarButton97(GetControl('BINSERT'));
    OnNewRecord;
//    Binsert.Click;
  end;
  //FIN PT3

end ;

Procedure TOM_ELTNATIONDOS.ValNiveauElipsisClick(Sender : TObject);
var
{$IFNDEF EAGLCLIENT}
  ValNiveau: THDBEdit;
{$ELSE}
  ValNiveau: THEdit;
{$ENDIF}
  Titre, Habilitation, Table, Ordre, St, Etab : String;
  j, Longueur : integer;
  Predefini:String;
begin
  {$IFNDEF EAGLCLIENT}
    ValNiveau := THDBEdit(GetControl('PED_VALEURNIVEAU'));
  {$ELSE}
    ValNiveau := THEdit(GetControl('PED_VALEURNIVEAU'));
  {$ENDIF}
  ValNiveau.DataType := '';

  if GetField('PED_TYPENIVEAU') = 'ETB' then
  begin
    // Rajouter la gestion des habilitations
    Titre := 'Liste des établissements';
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
          if j > 1 then Habilitation := Habilitation + ',';
          Habilitation := Habilitation + '"' + Etab + '"';
        end;
        Etab := ReadTokenSt(St);
      end;
      if j > 0 then
        Habilitation := '((ET_ETABLISSEMENT IN (' + Habilitation + ')))';
    end;
    Table := 'ETABLISS';
    Ordre := 'ET_ETABLISSEMENT';
    LookupList(THEdit(Sender),Titre,Table,'ET_ETABLISSEMENT','ET_LIBELLE',Habilitation,Ordre, True,-1);
  end;
  if GetField('PED_TYPENIVEAU') = 'SAL' then
  begin
    Titre := 'Liste des salariés';
    // Rajouter la gestion des habilitations
    if Assigned(MonHabilitation) then      
    begin
      if (MonHabilitation.LeSQL <> '') then
        if (copy(MonHabilitation.LeSQL, 1, 3) = 'PED') then
        begin
          Longueur := Length(MonHabilitation.LeSQL);
          Longueur := Longueur - 2;
          Habilitation := MidStr(MonHabilitation.LeSQL, 4, Longueur);
          Habilitation := 'PSA' + Habilitation;
        end
        else
          Habilitation := '(' + MonHabilitation.LeSQL + ')';
    end;
    Table := 'SALARIES';
    Ordre := 'PSA_SALARIE';
    LookupList(THEdit(Sender),Titre,Table,'PSA_SALARIE','PSA_LIBELLE||" "||PSA_PRENOM NOM',Habilitation,Ordre, True,-1);
  end;
  if GetField('PED_TYPENIVEAU') = 'POP' then
  begin
    ValNiveau.DataType := 'PGPOPULATIONSAL';
{$IFNDEF CPS1}
    Predefini := GetPredefiniPopulation('PAI');
    ValNiveau.Plus := ' AND PPC_PREDEFINI = "' + Predefini + '" ';
{$ENDIF}
    LookUpCombo(ValNiveau);
  end;
  ChargerLibelleValeurNiveau;
end;

procedure TOM_ELTNATIONDOS.DupliquerEltNationaux(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
{$ELSE}
  Code: THEdit;
{$ENDIF}
  AncValDate, AncValCode, AncValNiv, AncValValNiv: string;
  Champ: array[1..4] of Hstring;
  Valeur: array[1..4] of variant;
  PgLibel, PgAbreg, PgThem: string;
  Ok: Boolean;
  st: string;   //PT2
begin
  TFFiche(Ecran).BValider.Click; //PT1
  AncValDate := GetField('PED_DATEVALIDITE');
  AncValCode := GetField('PED_CODEELT');
  AncValNiv := GetField('PED_TYPENIVEAU');
  AncValValNiv := GetField('PED_VALEURNIVEAU');
  // PT9 PH 06/10/2003 V_421 Duplication en CWAS BUG
  PgLibel := GetField('PED_LIBELLE');
  PgAbreg := GetField('PED_ABREGE');
  PgThem := GetField('PED_THEMEELT');
  mode := 'DUPLICATION';
  if (AppelSpecifique <> 'CREATION_SPEC') then
    AglLanceFiche('PAY', 'CODEDOS', '', '', AncValCode + ';' + AncValNiv + ';' + AncValValNiv + ';4;TOUS')
  else
  begin
    if ParamTyp = 'ETB' then
      AglLanceFiche('PAY', 'CODEDOS', '', '', AncValCode + ';' + AncValNiv + ';' + AncValValNiv + ';4;ETAB')
    else if ParamTyp = 'SAL' then
      AglLanceFiche('PAY', 'CODEDOS', '', '', AncValCode + ';' + AncValNiv + ';' + AncValValNiv + ';4;SALARIE')
    else if ParamTyp = 'POP' then
      AglLanceFiche('PAY', 'CODEDOS', '', '', AncValCode + ';' + AncValNiv + ';' + AncValValNiv + ';4;POP')
    else
      AglLanceFiche('PAY', 'CODEDOS', '', '', AncValCode + ';' + AncValNiv + ';' + AncValValNiv + ';4;TOUS');
  end;

  if PGDateDupliquer <> '' then
  begin
    Champ[1] := 'PED_TYPENIVEAU';
    Valeur[1] := PGNiveauDupliquer;
    Champ[2] := 'PED_VALEURNIVEAU';
    Valeur[2] := PGValNiveauDupliquer;
    Champ[3] := 'PED_CODEELT';
    Valeur[3] := PGCodeDupliquer;
    Champ[4] := 'PED_DATEVALIDITE';
    Valeur[4] := UsDateTime(StrToDate(PGDateDupliquer));
    Ok := RechEnrAssocier('ELTNATIONDOS', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
      {$IFNDEF EAGLCLIENT}
        Code := THDBEdit(GetControl('PED_CODEELT'));
      {$ELSE}
        Code := THEdit(GetControl('PED_CODEELT'));
      {$ENDIF}
      if (code <> nil) then
        DupliquerPaie(TFFiche(Ecran).TableName, Ecran);

      SetField('PED_CODEELT', PGCodeDupliquer);
      SetField('PED_DATEVALIDITE', StrToDate(PGDateDupliquer));
      SetField('PED_TYPENIVEAU', PGNiveauDupliquer);
      SetField('PED_VALEURNIVEAU', PGValNiveauDupliquer);

      // PT9 PH 06/10/2003 V_421 Duplication en CWAS BUG
      // DEB PT16 on recupère toutes les valeurs initiales par defaut
      SetField('PED_LIBELLE', PgLibel);
      SetField('PED_ABREGE', PgAbreg);
      SetField('PED_THEMEELT', PgThem);

      SetControlEnabled('PED_CODEELT', False);
      SetControlEnabled('PED_DATEVALIDITE', False);
      SetControlEnabled('PED_TYPENIVEAU', False);
      SetControlEnabled('PED_VALEURNIVEAU', False);
//PT2
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create;
      st := 'Duplication de la rubrique '+AncValCode;
      Trace.add (st);
      st := 'Création de la rubrique '+ GetField('PED_CODEELT');
      Trace.add (st);
      EnDupl := 'OUI';
      if not (Ecran is TFSaisieList) then
        IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran))
      else
        IsDifferent(dernierecreate, 'PED', '', 'PED_LIBELLE', Trace, TFSaisieList(Ecran));
      FreeAndNil(Trace);
      EnDupl := 'NON';
//Fin PT2
      TFFiche(Ecran).Bouge(nbPost);
    end
    else
      HShowMessage('5;Elément National Dossier :;La duplication est impossible, l''élément existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;

procedure TOM_ELTNATIONDOS.OnChangeField(F: TField);
var
  Q : TQuery;
  Salarie,Etab : String;
  StWhere, Niveau, St : String;
begin
  inherited;
  if (F.FieldName = 'PED_CODEELT') and (DS.State in [dsInsert]) then
  begin
    StWhere := '';
    if ExisteSQL('SELECT PNR_TYPENIVEAU FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI##') then
    begin
      Niveau := GetField('PED_TYPENIVEAU');
      If Niveau = 'SAL' then StWhere := ' AND (PNR_TYPENIVEAU="SAL" OR PNR_NIVMAXPERSO="SAL")'
      else If Niveau = 'POP' then StWhere := ' AND (PNR_TYPENIVEAU="POP" OR (PNR_TYPENIVEAU<>"SAL" AND (PNR_NIVMAXPERSO="POP" OR PNR_NIVMAXPERSO="SAL") ))'
      else If Niveau = 'ETB' then StWhere := ' AND (PNR_TYPENIVEAU="ETB" OR ((PNR_TYPENIVEAU<>"POP" AND PNR_TYPENIVEAU<>"SAL") AND (PNR_NIVMAXPERSO="ETB" OR PNR_NIVMAXPERSO="POP" OR PNR_NIVMAXPERSO="SAL") ))';
    end;

    St := ' SELECT PEL_LIBELLE,PEL_ABREGE,PEL_THEMEELT,PEL_MONETAIRE,PEL_REGIMEALSACE,PEL_DECALMOIS FROM ELTNATIONAUX '+
          ' LEFT JOIN ELTNIVEAUREQUIS ON PNR_CODEELT=PEL_CODEELT' +
          ' WHERE ##PEL_PREDEFINI## PEL_CODEELT="' + Getfield('PED_CODEELT') + '"' + StWhere +
          ' ORDER BY PEL_DATEVALIDITE DESC';

    if LectureSeule = False then
      MajFiltreNiveau(Getfield('PED_CODEELT'));
    Q := OpenSQL(St, True);
    if not Q.Eof then
    begin
      SetControlText('TPEL_LIBELLE',Q.Fields[0].AsString);
      SetField('PED_LIBELLE', Q.Fields[0].AsString);
      SetField('PED_ABREGE', Q.Fields[1].AsString);
      SetField('PED_THEMEELT', Q.Fields[2].AsString);
      SetField('PED_MONETAIRE', Q.Fields[3].AsString);
      SetField('PED_REGIMEALSACE', Q.Fields[4].AsString);
      SetField('PED_DECALMOIS', Q.Fields[5].AsString);
    end
    else
    begin
      SetControlText('TPEL_LIBELLE','');
      SetField('PED_LIBELLE', '');
      SetField('PED_ABREGE', '');
      SetField('PED_THEMEELT', '');
      SetField('PED_MONETAIRE', '-');
      SetField('PED_REGIMEALSACE', '-');
      SetField('PED_DECALMOIS', '-');
      if Getfield('PED_CODEELT') <> '' then
      begin
        PGIBox('Le code élément ' + Getfield('PED_CODEELT') + ' n''existe pas');
        SetFocusControl('PED_CODEELT');
      end;
    end;
    Ferme(Q);
  end;

  if (F.FieldName = 'PED_VALEURNIVEAU') then
  begin
    if (GetField('PED_VALEURNIVEAU') = '') then
      SetcontrolCaption('LBL_VALEURNIVEAU','')
    else if GetField('PED_TYPENIVEAU') = 'SAL' then
    begin
      Salarie := Trim(GetField('PED_VALEURNIVEAU'));
      if (VH_PAIE.PgTypeNumSal = 'NUM') and (isnumeric(Salarie)) and (Salarie <> '') then
      begin
         Salarie := ColleZeroDevant(StrToInt(trim(Salarie)), 10);
         if Salarie <> (GetField('PED_VALEURNIVEAU')) then
            SetField('PED_VALEURNIVEAU', Salarie);
         SetControlCaption('LBL_VALEURNIVEAU',RechDom('PGSALARIE',Salarie,False));
      end;
    end
    else if GetField('PED_TYPENIVEAU') = 'ETB' then
    begin
      Etab := Trim(GetField('PED_VALEURNIVEAU'));
    end;
  end;
  if (F.FieldName = 'PED_TYPENIVEAU') then
    ChargerLibelleValeurNiveau;
end;

procedure TOM_ELTNATIONDOS.EltnatElipsisClick(Sender : TObject);
var
  StSQL, Cond, Niveau : String;
begin
  StSQL := 'SELECT DISTINCT PEL_CODEELT, PEL_LIBELLE FROM ELTNATIONAUX';
  if ExisteSQL('SELECT PNR_TYPENIVEAU FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI##') then
  begin
    StSQL := StSQL + ' LEFT JOIN ELTNIVEAUREQUIS ON PNR_CODEELT=PEL_CODEELT Where ##PEL_PREDEFINI## ';

    Niveau := GetField('PED_TYPENIVEAU');
    If Niveau = 'SAL' then Cond := ' (PNR_TYPENIVEAU="SAL" OR PNR_NIVMAXPERSO="SAL")'
    else If Niveau = 'POP' then Cond := ' (PNR_TYPENIVEAU="POP" OR (PNR_TYPENIVEAU<>"SAL" AND (PNR_NIVMAXPERSO="POP" OR PNR_NIVMAXPERSO="SAL") ))'
    else If Niveau = 'ETB' then Cond := ' (PNR_TYPENIVEAU="ETB" OR ((PNR_TYPENIVEAU<>"POP" AND PNR_TYPENIVEAU<>"SAL") AND (PNR_NIVMAXPERSO="ETB" OR PNR_NIVMAXPERSO="POP" OR PNR_NIVMAXPERSO="SAL") ))';

    //DEB PT6
    StSQL := StSQL + ' AND PEL_CODEELT IN (SELECT DISTINCT PNR_CODEELT FROM ELTNIVEAUREQUIS ' +
      ' WHERE ((PNR_PREDEFINI = "CEG" AND ' + Cond +
      ' AND PNR_CODEELT NOT IN (SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_PREDEFINI="STD"))' +
      ' OR (PNR_PREDEFINI = "STD" AND ' + Cond + ')))';
    //FIN PT6
  end;
  {$IFNDEF EAGLCLIENT}
  StSQL := StSQL + ' ORDER BY PEL_CODEELT'; // PT7
  {$ENDIF}

  {$IFNDEF EAGLCLIENT}
//  LookupList(THDBEdit(Sender),'Eléments nationaux','ELTNATIONAUX LEFT JOIN ELTNIVEAUREQUIS ON PNR_CODEELT=PEL_CODEELT','DISTINCT PEL_CODEELT','PEL_LIBELLE',StWhere,StOrder, True,-1);
  LookupList(THDBEdit(Sender),'Eléments nationaux','','DISTINCT PEL_CODEELT','','','', True,-1,StSQL);  //PT6
  {$ELSE}
//  LookupList(THEdit(Sender),'Eléments nationaux','ELTNATIONAUX LEFT JOIN ELTNIVEAUREQUIS ON PNR_CODEELT=PEL_CODEELT','DISTINCT PEL_CODEELT','PEL_LIBELLE',StWhere,StOrder, True,-1);
  LookupList(THEdit(Sender),'Eléments nationaux','','DISTINCT PEL_CODEELT','','','', True,-1,StSQL);    //PT6
  {$ENDIF}
end;

procedure TOM_ELTNATIONDOS.MajFiltreNiveau(CodeElt : String);
var
  Q : TQuery;
  {$IFNDEF EAGLCLIENT}
  TypeNiveau : THDBValComboBox;
  {$ELSE}
  TypeNiveau : THValComboBox;
  {$ENDIF}
  NiveauOK : String;
begin
  if CodeElt <> '' then
  begin
    {$IFNDEF EAGLCLIENT}
    TypeNiveau := THDBValComboBox(GetControl('PED_TYPENIVEAU'));
    {$ELSE}
    TypeNiveau := THValComboBox(GetControl('PED_TYPENIVEAU'));
    {$ENDIF}
    if TypeNiveau <> nil then
    begin
      if StPop <> '' then
        NiveauOk := '"ETB","POP","SAL"'
      else
        NiveauOk := '"ETB","SAL"';
      Q := OpenSQL('SELECT PNR_TYPENIVEAU,PNR_NIVMAXPERSO FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT="' + CodeElt + '" AND ##PNR_PREDEFINI##', True);
      if not Q.Eof then
      begin
        NiveauOK := '""';
        if (Q.Fields[1].AsString <> 'CEG') and (Q.Fields[1].AsString <> 'STD') and (Q.Fields[1].AsString <> 'DOS') then
        begin
          if (Q.Fields[0].AsString = 'ETB') or(Q.Fields[0].AsString = 'CEG') or (Q.Fields[0].AsString = 'STD') or (Q.Fields[0].AsString = 'DOS') then
          begin
            NiveauOK := '"ETB"';
            if (Q.Fields[1].AsString = 'POP') then
            begin
              if StPop <> '' then
                NiveauOk := NiveauOK + ',' + StPop;
            end;
            if (Q.Fields[1].AsString = 'SAL') then
            begin
              if StPop <> '' then
                NiveauOk := NiveauOK + ',' + StPop + ',"SAL"'
              else
                NiveauOk := NiveauOK + ',"SAL"';
            end;
          end;
          if (Q.Fields[0].AsString = 'POP') then
          begin
            if StPop <> '' then
              NiveauOK := StPop
            else
              NiveauOK := '';
            if (Q.Fields[1].AsString = 'SAL') then
            begin
              if NiveauOk <> '' then
                NiveauOk := NiveauOK + ',';
              NiveauOk := NiveauOK + '"SAL"';
            end;
          end;
          if (Q.Fields[0].AsString = 'SAL') then
            NiveauOK := '"SAL"';
        end;
      end;
      TypeNiveau.Plus := ' AND CO_CODE IN (' + NiveauOK + ')';
      Ferme(Q);
    end;
  end;
end;

Initialization
  registerclasses ( [ TOM_ELTNATIONDOS ] ) ;
end.
