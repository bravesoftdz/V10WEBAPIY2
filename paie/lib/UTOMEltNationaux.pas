{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : unit de gestion des éléments nationnaux
Mots clefs ... : PAIE
*****************************************************************
PT1 SB 29/11/2001 V563 On force la validation avant la Duplication
PT2 SB 30/11/2001 V563 Dysfonctionnement test code existant sur Duplication
PT3 SB 13/12/2001 V570 Fiche de bug n° 279
                       Test code existant ne test pas bon numéro de dossier
PT4 PH 23/01/2002 V571 Dysfonctionnement test code existant sur Duplication à cause
                       de la date de validité non convertie en UsDateTime
PT5 PH 08/10/2002 V585 Possibilité de créer des elts nationaux std et dos à partir de
                       ceux de CEGID. Plus de gestion pair, 1 ou 3, 5/7/9
PT6 PH 07/11/2002 V591 Indication si on applique le décalage dans le cas d'un exercice décalé
PT7 SB 20/02/2003 V_42C FQ 10526 Accès à la duplication au utilisateur non Cegid no réviseur
PT8 PH 05/06/2003 V_421 FQ 10630 Contrôle du contenu du champ predefini
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT10 PH 05/04/2004 V_50 FQ 11231 Prise en compte Alsace LOrraine
PT11 PH 23/09/2004 V_50 FQ 11458 On ne peut plus détruire un élément national CEGID
PT12 PH 04/10/2004 V_50 FQ 11590 Modif chargement de tablette
PT13 PH 09/08/2005 V_60 FQ 12279 Controle du code de l'élément national
PT14 PH 09/08/2005 V_60 FQ 12026 Impossibilité de dupliquer un elt national STD en CEG
PT15 PH 01/09/2005 V_60 FQ 12543 Controle du thème obligatoire
PT16 PH 03/10/2005 V_60 FQ 12627 Duplication Element nationaux reprend les valeurs par defaut
PT17 EPI 25/04/2006 V_65 FQ 12026 Impossibilité de dupliquer un elt national DOS en CEG
PT18 GGS 31/01/2007 V_72 FQ 12694 Journal événements
PT19 FC 08/02/2007 V_72 Mise en place de la convention collective dans les éléments nationaux
PT20 FC 05/03/2007 V_70 Mise en place de paramétres en entrée dans le cas où la saisie est lancée depuis
                        le calcul du bulletin
PT21 PH 08/03/2007 V_70 FQ 13960 Création elt nationaux STD
PT22 FC 12/03/2007 V_70 Renseigner un flag + btn pour dire s'il existe des redéfinitions d'élément au niveau dossier
PT23 FC 13/04/2007 V_72 FQ 14059 Rendre accessible les zones montants quand duplication d'un élément CEGID
PT24 GGS 26/04/2007 V_72 Revue gestion de trace
PT25 MF 03/05/2007 V_72 Correction acces vio en gestion des traces
                        (cas : suppression d'un élément CEGID)
PT24-2 GGS 29/05/2007 V_72 Revue gestion de trace  en duplication
PT26 FC 31/05/2007 V_72 FQ 14304 Accéder directement au niveau préconisé du dit élément
PT27 FC 04/06/2007 V_72 Duplication d'un élément national en élément dossier
PT28 FC 12/06/2007 V_72 Gestion du niveau préconisé pour filter la combo Predefini
PT29 FC 12/06/2007 V_72 Popupmenu à la place de 2 boutons duplication
PT30 FC 22/06/2007 V_72 FQ 14440 Les éléments dossiers héritent du libellé défini au niveau supérieur
PT31 FC 04/07/2007 V_72 FQ 14475 Concepts
PT32 FC 05/11/2007 V_80 FQ 14903 Duplication / contrôle unicité incohérent si convention collective différente
PT33 FC 06/11/2007 V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
PT34 FC 20/11/2007 V_80 FQ 14956 Duplication élément national, initialiser la convention collective
PT35 FC 02/01/2008 V_81 FQ 15031 Avoir un contrôle lors de la création d'un élément national en STD
                        si le code existe déjà
}

unit UTOMEltNationaux;

interface
uses Controls, Classes, forms, sysutils,

{$IFNDEF EAGLCLIENT}
  db, HDB, DBCtrls, Fiche, Fe_Main,
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}     //PT22
{$ELSE}
  eFiche, MaineAgl, UtileAGL, Utob,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, HTB97,
  HDebug, PgOutils, PgOutils2,P5Def,PAIETOM,      //PT18 rajout PAIETOM
  ParamSoc,Menus;//PT22
type

  TOM_EltNationaux = class(PGTOM)           //PT18 class TOM devient PGTOM
  private
    Code, mode: string; // Code Element
    Libelle, predefini: string; // Libelle
    ThemeElt, DerniereCreate: string; // Theme
    Monetaire, LectureSeule, CEG, STD, DOS, OnFerme: Boolean;
    Trace: TStringList;
    Action,ParamElt:string; //PT20
    ParamDat,Date1900 :TDateTime;  //PT20
    LeStatut:TDataSetState; //PT33
    procedure DupliquerEltNationaux(Sender: TObject);
    procedure DupliquerEltDossier(Sender: TObject); //PT27
    procedure AppelEltDynDos(Sender: TObject); //PT22
    procedure AppelNivPreconise(Sender : TObject); //PT26
    procedure MajFiltreNiveau(CodeElt : String); //PT28
  public
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnAfterUpdateRecord; override;
  end;

implementation

{ TOM_EltNationaux }

procedure TOM_EltNationaux.DupliquerEltNationaux(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
{$ELSE}
  Code: THEdit;
{$ENDIF}
  AncValDate, AncValCode, AncValPred: string;
  Champ: array[1..5] of Hstring;      //PT32
  Valeur: array[1..5] of variant;     //PT32
  PgLibel, PgAbreg, PgThem, PgConv: string;
  Ok: Boolean;
  TypeP: string; // PT14
  CEG, STD, DOS: boolean;// PT14
  st: string;
begin
  TFFiche(Ecran).BValider.Click; //PT1
  AncValDate := GetField('PEL_DATEVALIDITE');
  AncValCode := GetField('PEL_CODEELT');
  AncValPred := GetField('PEL_PREDEFINI');
  // PT9 PH 06/10/2003 V_421 Duplication en CWAS BUG
  PgLibel := GetField('PEL_LIBELLE');
  PgAbreg := GetField('PEL_ABREGE');
  PgThem := GetField('PEL_THEMEELT');
  PgConv := GetField('PEL_CONVENTION');  //PT34
  mode := 'DUPLICATION';
  AglLanceFiche('PAY', 'CODE', '', '', 'ELT;' + AncValCode + ';' + AncValPred + ';4;' + PgConv); //PT34
  // DEB PT14
  TypeP := PGCodePredefini;
  AccesPredefini(TypeP, CEG, STD, DOS);
{$IFDEF CPS1} // PT21
   STD := FALSE;
{$ENDIF}

  if (TypeP = 'CEG') AND (NOT CEG) AND (AncValPred='STD') then
  begin
    PGiBox ('Vous ne pouvez pas dupliquer un élément national standard en CEGID.', Ecran.Caption);
    exit;
  end;
  // FIN PT14

  // DEBUT PT17
  if (TypeP = 'CEG') AND (NOT CEG) AND (AncValPred='DOS') then
  begin
    PGiBox ('Vous ne pouvez pas dupliquer un élément national dossier en CEGID.', Ecran.Caption);
    exit;
  end;
  // FIN PT17

  if PGDateDupliquer <> '' then
  begin
    Champ[1] := 'PEL_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PEL_NODOSSIER';
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier() //PT3
    else Valeur[2] := '000000'; //PT3
    Champ[3] := 'PEL_CODEELT';
    Valeur[3] := PGCodeDupliquer; //PT2 StrToInt
    // PT4 PH 23/01/2002 V571 Dysfonctionnement test code existant sur Duplication sur date de validité
    Champ[4] := 'PEL_DATEVALIDITE';
    Valeur[4] := UsDateTime(StrToDate(PGDateDupliquer));
    //DEB PT32
    Champ[5] := 'PEL_CONVENTION';
    Valeur[5] := PGConvDupliquer;
    //FIN PT32
    Ok := RechEnrAssocier('ELTNATIONAUX', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
{$IFNDEF EAGLCLIENT}
      Code := THDBEdit(GetControl('PEL_CODEELT'));
{$ELSE}
      Code := THEdit(GetControl('PEL_CODEELT'));
{$ENDIF}
      if (code <> nil) then
        DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PEL_CODEELT', PGCodeDupliquer);
      SetField('PEL_DATEVALIDITE', StrToDate(PGDateDupliquer));
      SetField('PEL_PREDEFINI', PGCodePredefini);
      SetField('PEL_CONVENTION', PGConvDupliquer);

      // PT9 PH 06/10/2003 V_421 Duplication en CWAS BUG
      // DEB PT16 on recupère toutes les valeurs initiales par defaut
      SetField('PEL_LIBELLE', PgLibel);
      SetField('PEL_ABREGE', PgAbreg);
      SetField('PEL_THEMEELT', PgThem);
      // FIN PT16
      AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1} // PT21
   STD := FALSE;
{$ENDIF}

      if PGCodePredefini = 'CEG' then
      begin
        LectureSeule := (CEG = False);
        PaieLectureSeule(TFFiche(Ecran), (CEG = False));
        if LectureSeule = True then
        begin
          if (STD = True) then //DEB PT23
          begin
            SetControlEnabled('PEL_MONTANT', STD);
            SetControlEnabled('PEL_MONTANTEURO', STD);
            SetControlEnabled('BValider', STD);
          end                  //FIN PT23
          else
          begin
            SetControlEnabled('PEL_MONTANT', CEG);
            SetControlEnabled('PEL_MONTANTEURO', CEG);
          end;
        end;
      end;
      if PGCodePredefini = 'STD' then
      begin
        LectureSeule := (STD = False);
        PaieLectureSeule(TFFiche(Ecran), (STD = False));
        if LectureSeule = True then
          if STD = True then
          begin
            SetControlEnabled('PEL_MONTANT', STD);
            SetControlEnabled('PEL_MONTANTEURO', STD);
          end;
      end;
      SetField('PEL_NODOSSIER', '000000');
      SetControlEnabled('PEL_CODEELT', False);
      SetControlEnabled('PEL_DATEVALIDITE', False);
      SetControlEnabled('PEL_PREDEFINI', False);
      SetControlEnabled('PEL_CONVENTION', False);
//PT24-2
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create; //PT44
      st := 'Duplication de la rubrique '+AncValCode;
      Trace.add (st);
      st := 'Création de la rubrique '+ GetField('PEL_CODEELT');
      Trace.add (st);
      EnDupl := 'OUI';
      IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran)); //PT44
      FreeAndNil(Trace);
      EnDupl := 'NON';
//Fin PT24-2
      Debug('Avant bouge');
      TFFiche(Ecran).Bouge(nbPost);
      Debug('Après bouge');
// PT12     ChargementTablette(TFFiche(Ecran).TableName, ''); //recharge tablette
    end
    else
      HShowMessage('5;Elément National :;La duplication est impossible, l''élément existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';

end;

procedure TOM_EltNationaux.OnAfterUpdateRecord;
var
  even: boolean;      //PT18
begin
  inherited;
// Exceptionnellement  TFFiche(Ecran).CodeName est remplace par le nom du champ de la table
  Trace := TStringList.Create ;         //PT24
  even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),'PEL_CODEELT',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut);  //PT18 //PT33
  FreeAndNil (Trace);   //PT24  Trace.Free;
  if OnFerme then Ecran.Close;
end;

procedure TOM_EltNationaux.OnArgument(stArgument: string);
var
  Btn: TToolBarButton97;
  MenuDupli : TPopupMenu; //PT29
begin
  Inherited ;
  //DEB PT20
  Action := '';
  ParamElt := '';
  ParamDat := StrToDate('01/01/1900');
  if stArgument <> '' then
  begin
    Action := READTOKENST(stArgument);
    ParamElt := READTOKENST(stArgument);
    if stArgument <> '' then
      ParamDat := StrToDate(READTOKENST(stArgument));
  end;
  //FIN PT20

  //DEB PT29
  MenuDupli := TPopupMenu(GetControl('POPUPDUPLI'));
  if MenuDupli <> nil then
  begin
    if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
    begin
      MenuDupli.OnPopup := DupliquerEltNationaux;
      MenuDupli.Items[0].Visible := False;
      MenuDupli.Items[1].Visible := False;
    end
    else
    begin
      MenuDupli.Items[0].OnClick := DupliquerEltNationaux;
      MenuDupli.Items[1].OnClick := DupliquerEltDossier;
    end;
  end;

  //DEB PT22
  Btn := TToolBarButton97(GetControl('BELTNATIONDOS'));
  if Btn <> nil then Btn.Visible := False;
  if GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    Btn.Visible := True;
    if Btn <> nil then
      Btn.OnClick := AppelEltDynDos;
  end;
  //FIN PT22

  //DEB PT26
  Btn := TToolBarButton97(GetControl('BNIVEAUREQUIS'));
  if Btn <> nil then Btn.Visible := False;
  if GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    Btn.Visible := True;
    if Btn <> nil then
      Btn.OnClick := AppelNivPreconise;
  end;
  //FIN PT26

  if not JaiLeDroitTag(41104) then
    SetControlEnabled('BNIVEAUREQUIS',False);
end;

procedure TOM_EltNationaux.OnChangeField(F: TField);
var
  EltNat, TempEltNat, pred: string;
  Icode: integer;
  OKRub: boolean;
  Mes, vide: string;
begin
  inherited;
  if mode = 'DUPLICATION' then exit;
  vide := '';
  if (F.FieldName = 'PEL_CODEELT') then
  begin
    MajFiltreNiveau(Getfield('PEL_CODEELT'));
    EltNat := Getfield('PEL_CODEELT');
    if EltNat = '' then exit;
    if ((isnumeric(EltNat)) and (EltNat <> '    ')) then
    begin
      iCode := strtoint(trim(EltNat));
      TempEltNat := ColleZeroDevant(iCode, 4);
      if (DS.State = dsinsert) and (TempEltNat <> '') and (GetField('PEL_PREDEFINI') <> '') then
      begin
        OKRub := TestRubriqueCegStd(GetField('PEL_PREDEFINI'), TempEltNat, 1000);
        // PT5 PH 08/10/2002 V585 Possibilité de créer des elts nationaux std et dos à partir de
        {           Pair:= CodePair(EltNat);
                   if (GetField('PEL_PREDEFINI') = 'CEG') then OKRub := FALSE;
                   if (GetField('PEL_PREDEFINI') = 'CEG') AND (Pair) AND (V_PGI.User <> 'WZX') then OKRub := FALSE;
                   if StrToInt (EltNat) < 1000 then OKRub := FALSE;}
        if (EltNat = '0000') or (OkRub = False) then
        begin
          Mes := MesTestRubrique('ELT', GetField('PEL_PREDEFINI'), 1000);
          //              Mes:='Vous n''êtes pas autorisé à créer un élément national CEGID';
          HShowMessage('2;Code Erroné: ' + TempEltNat + ' ;' + Mes + ';W;O;O;;;', '', '');
          TempEltNat := '';
        end;
      end;
      if TempEltNat <> EltNat then
      begin
        SetField('PEL_CODEELT', TempEltNat);
        SetFocusControl('PEL_CODEELT');
      end;
    end;
  end;

  if (F.FieldName = 'PEL_PREDEFINI') and (DS.State = dsinsert) then
  begin
    EltNat := (GetField('PEL_CODEELT'));
    Pred := GetField('PEL_PREDEFINI');
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1} // PT21
   STD := FALSE;
{$ELSE}
    if (Pred = 'DOS') then
    begin
      PGIBox('Vous ne pouvez créer d''élément prédéfini dossier, #13#10' +
        ' Vous ne pouvez que dupliquer un élément national ' {PT7 standard ou cegid} + 'en dossier.', 'Accès refusé');
      Pred := '';
      SetControlProperty('PEL_PREDEFINI', 'Value', Pred);
    end;
{$ENDIF}

    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer d''élément prédéfini CEGID', 'Accès refusé');
      Pred := '';
      SetControlProperty('PEL_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer d''élément prédéfini Standard', 'Accès refusé');
      Pred := '';
      SetControlProperty('PEL_PREDEFINI', 'Value', Pred);
    end;
    if (EltNat <> '') and (pred <> '') then
    begin
      // PT5 PH 08/10/2002 V585 Possibilté de créer des elts nationaux std et dos
      OKRub := TestRubriqueCegStd(pred, EltNat, 1000);
      if (OkRub = False) or (EltNat = '0000') then
      begin
        Mes := MesTestRubrique('ELT', pred, 1000);
        HShowMessage('2;Code Erroné: ' + EltNat + ' ;' + Mes + ';W;O;O;;;', '', '');
        SetField('PEL_CODEELT', vide);
        if Pred <> GetField('PEL_PREDEFINI') then SetField('PEL_PREDEFINI', pred);
        SetFocusControl('PEL_CODEELT');
        exit;
      end;
    end;
    if Pred <> GetField('PEL_PREDEFINI') then SetField('PEL_PREDEFINI', pred);
  end;

  if (ds.state in [dsBrowse]) then
  begin
    AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1} // PT21
   STD := FALSE;
{$ENDIF}

    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
      if (Getfield('PEL_PREDEFINI') = 'CEG') and (STD = True) then
      begin
        SetControlEnabled('PEL_MONTANT', STD);
        SetControlEnabled('PEL_MONTANTEURO', STD);
        SetControlEnabled('BInsert', STD);
        SetControlEnabled('BValider', STD); //PT23
      end;
    end;
    SetControlEnabled('PEL_PREDEFINI', False);
    SetControlEnabled('PEL_CODEELT', False);
    SetControlEnabled('PEL_DATEVALIDITE', False);
    SetControlEnabled('PEL_CONVENTION', False);
  end;
  // PT10 PH 05/04/2004 V_50 FQ 11231 Prise en compte Alsace LOrraine
{  if F.FieldName = ('PEL_REGIMEALSACE') then
  begin
    if (GetField('PEL_MONETAIRE') = 'X') then
    begin
      SetControlEnabled('PEL_MONTANT', TRUE)
    end
    else
    begin
      SetControlEnabled('PEL_MONTANT', FALSE);
    end;
  end;}
  // FIN PT10
end;

procedure TOM_EltNationaux.OnDeleteRecord;
begin
  inherited;
// DEB PT11
  if (GetField('PEL_PREDEFINI') = 'CEG') then
  begin
    PGIBox('Vous ne pouvez pas détruire un élément national CEGID', Ecran.Caption);
    LastError := 1;
  end
  else
// PT12  ChargementTablette(TFFiche(Ecran).TableName, '');
// FIN PT11
//PT18
    begin // PT25
    Trace := TStringList.Create ;         //PT24
    Trace.Add('SUPPRESSION ELEMENT NATIONAL '+GetField('PEL_CODEELT')+' '+ GetField('PEL_LIBELLE'));
    CreeJnalEvt('003','085','OK',nil,nil,Trace);
    FreeAndNil (Trace);  //PT24  Trace.free;
    end; // PT25
end;

procedure TOM_EltNationaux.OnLoadRecord;
var
{$IFNDEF EAGLCLIENT}
  CodeElt, DateElt: THDBEdit;
{$ELSE}
  CodeElt, DateElt: THEdit;
{$ENDIF}
  okok: boolean;
  Q:TQuery;      //PT22
  Flag:THLabel;  //PT22
  MenuDupli : TPopupMenu; //PT31
begin
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  //If MODE='REFUSE' then exit;
  Code := string(GetField('PEL_CODEELT'));
  Libelle := string(GetField('PEL_LIBELLE'));
{$IFNDEF EAGLCLIENT}
  CodeElt := THDBEdit(GetControl('PEL_CODEELT'));
  DateElt := THDBEdit(GetControl('PEL_DATEVALIDITE'));
{$ELSE}
  CodeElt := THEdit(GetControl('PEL_CODEELT'));
  DateElt := THEdit(GetControl('PEL_DATEVALIDITE'));
{$ENDIF}
  predefini := string(GetField('PEL_PREDEFINI'));
  ThemeElt := GetField('PEL_THEMEELT');
  Monetaire := (GetField('PEL_MONETAIRE') = 'X');
  okok := (DS <> nil) and (DS.State = dsInsert);
  if CodeElt <> nil then CodeElt.Enabled := okok;
  if DateElt <> nil then DateElt.Enabled := okok;
  AccesPredefini('TOUS', CEG, STD, DOS);
{$IFDEF CPS1} // PT21
   STD := FALSE;
{$ENDIF}

  LectureSeule := FALSE;
  if (Getfield('PEL_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
    if LectureSeule = True then
    begin
      if STD = True then
      begin
        SetControlEnabled('PEL_MONTANT', STD);
        SetControlEnabled('PEL_MONTANTEURO', STD);
        SetControlEnabled('BInsert', STD);
        SetControlEnabled('BValider', STD); //PT23
      end;
    end;
  end;
  if (Getfield('PEL_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  if (Getfield('PEL_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := (DOS = False);
    PaieLectureSeule(TFFiche(Ecran), LectureSeule);
    SetControlEnabled('BDelete', DOS);
    SetControlEnabled('PEL_MONTANT',DOS);
    SetControlEnabled('PEL_MONTANTEURO',DOS);
  end;

  SetControlEnabled('PEL_PREDEFINI', False);
  SetControlEnabled('PEL_DATEVALIDITE', False);
  SetControlEnabled('PEL_CONVENTION', False);
  SetControlEnabled('PEL_CODEELT', False);

  if (DS.State in [dsInsert]) and (Mode <> 'DUPLICATION') then
  begin
    LectureSeule := FALSE;
    SetControlEnabled('PEL_PREDEFINI', True);
    if Trim(CodeElt.text) = '' then
      SetControlEnabled('PEL_CODEELT', True);
    SetControlEnabled('PEL_DATEVALIDITE', True);
    SetControlEnabled('PEL_CONVENTION', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDUPLIQUER', False);
    SetControlEnabled('BDelete', False);
  end;

  //DEB PT22
  // Si les éléments dynamiques dossier sont gérés, vérifier s'il existe des
  // redéfinitions pour l'élément national en cours
  Flag := THLabel(GetControl('LBLFLAGELTNATDOS'));
  if GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    if (Flag <> nil) then
    begin
      Flag.Caption := '';
      Q := OpenSql('SELECT PED_CODEELT FROM ELTNATIONDOS ' +
      ' WHERE PED_CODEELT = "' + GetField('PEL_CODEELT') + '"',True);
      if not Q.eof then //PORTAGECWAS
        Flag.Caption := 'Attention, cet élément national est personnalisé';
      Ferme(Q);
    end;
  end
  else
  begin
    if (Flag <> nil) then
      Flag.Caption := '';
  end;
  //FIN PT22

  MenuDupli := TPopupMenu(GetControl('POPUPDUPLI'));
  //DEB PT31
  if (Code <> '') then
  begin
    if ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT = "' + Code + '" AND ##PNR_PREDEFINI##') then
    begin
      if ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT = "' + Code + '" AND PNR_NIVMAXPERSO IN ("ETB","POP","SAL") AND ##PNR_PREDEFINI##') then
      begin
        if MenuDupli <> nil then
          MenuDupli.Items[1].Visible := True;
      end
      else
      begin
        if MenuDupli <> nil then
          MenuDupli.Items[1].Visible := False;
      end;
    end;
  end;

  if (CEG = False) and (STD = False) and (DOS = False) and (not (DS.State in [dsInsert])) then
  begin
    SetControlEnabled('BDUPLIQUER',False);
  end
  else
  begin
    SetControlVisible('BDUPLIQUER', True);
    if (DOS = False) then
      MenuDupli.Items[1].Visible := False
//    else
//      MenuDupli.Items[1].Visible := True;
  end;
  if (Code <> '') then
    if (STD = False) and (not ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT="' + Code + '" AND ##PNR_PREDEFINI##')) then
      SetControlEnabled('BNIVEAUREQUIS',False);
  //FIN PT31
  MajFiltreNiveau(Code);
end;

//DEB PT22
procedure TOM_EltNationaux.AppelEltDynDos(Sender: TObject);
begin
  AglLanceFiche('PAY', 'ELTNATDOS_MUL', '', '', 'ELT;' + GetField('PEL_CODEELT'));
end;
//FIN PT22

procedure TOM_EltNationaux.OnNewRecord;
var
{$IFNDEF EAGLCLIENT}
  CodElt:THDBEdit;
  DatVal:THDBEdit;
{$ELSE}
  CodElt:THEdit;
  DatVal:THEdit;
{$ENDIF}
begin
  //DEB PT20
  if (ParamElt <> '') then
  begin
    {$IFNDEF EAGLCLIENT}
      CodElt := THDBEdit(GetControl('PEL_CODEELT'));
    {$ELSE}
      CodElt := THEdit(GetControl('PEL_CODEELT'));
    {$ENDIF}
    if CodElt <> nil then
    begin
      CodElt.Text := ParamElt;
      SetField('PEL_CODEELT', ParamElt);
      CodElt.Enabled := False;
    end;
  end;
  Date1900 := StrToDate('01/01/1900');
  if ParamDat <> Date1900 then
  begin
    {$IFNDEF EAGLCLIENT}
      DatVal := THDBEdit(GetControl('PEL_DATEVALIDITE'));
    {$ELSE}
      DatVal := THEdit(GetControl('PEL_DATEVALIDITE'));
    {$ENDIF}
    if DatVal <> nil then
    begin
      DatVal.Text := DateToStr(ParamDat);
      SetField('PEL_DATEVALIDITE', ParamDat);
    end;
  end
  else
    SetField('PEL_DATEVALIDITE', Date);
  //FIN PT20

  {SetField('PEL_CODEELT', Code);
  SetField('PEL_LIBELLE', Libelle);   }
//PT20  SetField('PEL_DATEVALIDITE', Date);
  {SetField ('PEL_THEMEELT', ThemeElt);
  SetField ('PEL_MONETAIRE', Monetaire);   }
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  SetField('PEL_NODOSSIER', PgRendNoDossier());
  // PT6 PH 07/11/2002 V591 Indication si on applique le décalage dans le cas d'un exercice décalé
  setField('PEL_DECALMOIS', 'X');
  setField('PEL_MONETAIRE', 'X');
  //SetField ('PEL_PREDEFINI', predefini);
  //PT10 PH 05/04/2004 V_50 FQ 11231 Prise en compte Alsace LOrraine
  setField('PEL_REGIMEALSACE', '-');
  setField('PEL_CONVENTION', '000'); //PT19

  SetControlEnabled('BELTNATIONDOS',False);
  SetControlEnabled('BNIVEAUREQUIS',False);
end;

procedure TOM_EltNationaux.OnUpdateRecord;
var
  Predef: string;
begin
  OnFerme := False;
  LeStatut := DS.State; //PT33
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PEL_CODEELT')
  else
    if (DerniereCreate = GetField('PEL_CODEELT')) then OnFerme := True; // le bug arrive on se casse !!!

  if (DS.State = dsinsert) then
  begin
    if (GetField('PEL_PREDEFINI') <> 'DOS') then
      SetField('PEL_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PEL_NODOSSIER', PgRendNoDossier());
  end;

  Predef := GetField('PEL_PREDEFINI');
  // PT5 PH 08/10/2002 V585 Possibilté de créer des elts nationaux std et dos
  // PT8 PH 05/06/2003 V_421 FQ 10630 Contrôle du contenu du champ predefini
  if (Predef <> 'CEG') and (Predef <> 'STD') and (Predef <> 'DOS') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PEL_PREDEFINI');
  end;
  // FIN PT8
  // DEB PT13
  if (GetField('PEL_CODEELT') = '') then
  begin
    LastError := 2;
    LastErrorMsg := 'Vous devez renseigner le code de l''élément national';
    SetFocusControl('PEL_CODEELT');
  end;
  // DEB PT15
  if (GetField('PEL_THEMEELT') = '') then
  begin
    LastError := 2;
    LastErrorMsg := 'Vous devez renseigner le thème de l''élément national';
    SetFocusControl('PEL_THEMEELT');
  end;
  // FIN PT15
  // FIN PT13

  //DEB PT20
  Date1900 := StrToDate('01/01/1900');
  if (ParamDat <> Date1900) and (GetField('PEL_DATEVALIDITE') > ParamDat) then
  begin
    LastError := 2;
    LastErrorMsg := 'La date de validité doit être inférieure ou égale au ' + DateToStr(ParamDat);
    SetFocusControl('PEL_DATEVALIDITE');
  end;
  //FIN PT20

  //DEB PT30
  if GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    if Libelle <> GetField('PEL_LIBELLE') then
    begin
      ExecuteSQL('UPDATE ELTNATIONDOS SET PED_LIBELLE = "' + GetField('PEL_LIBELLE') + '"' +
        ' WHERE PED_CODEELT = "' + GetField('PEL_CODEELT') + '"');
      Libelle := GetField('PEL_LIBELLE');
    end;
  end;
  //FIN PT30

     //Rechargement des tablettes
{ PT12
  if (LastError = 0) and (Getfield('PEL_CODEELT') <> '') and (Getfield('PEL_LIBELLE') <> '') then
    ChargementTablette(TFFiche(Ecran).TableName, '');
    }

  //DEB PT35
  if (DS.State in [dsInsert]) and (ExisteSQL('SELECT PEL_CODEELT FROM ELTNATIONAUX WHERE ##PEL_PREDEFINI## PEL_CODEELT="' + Getfield('PEL_CODEELT') + '" AND PEL_PREDEFINI="STD"')) then
  begin
    if PGIAsk('Un élément STD existe déjà avec ce code. Confirmez-vous la création ?','Le code existe déjà') <> mrYes then
    begin
      LastError := 2;
      LastErrorMsg := 'Le code de l''élément national existe déjà';
    end;
  end;
  //FIN PT35
end;

//DEB PT26
procedure TOM_EltNationaux.AppelNivPreconise(Sender: TObject);
var
  MenuDupli : TPopupMenu; //PT31
begin
  AglLanceFiche('PAY', 'ELEMENTNIVREQUIS', '', 'STD;000000;' + GetField('PEL_CODEELT'), GetField('PEL_CODEELT'));
  MenuDupli := TPopupMenu(GetControl('POPUPDUPLI'));
  if MenuDupli <> nil then
  begin
    if (GetField('PEL_CODEELT') <> '') then
    begin
      if ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT = "' + GetField('PEL_CODEELT') + '" AND ##PNR_PREDEFINI##') then
      begin
        if ExisteSQL('SELECT PNR_CODEELT FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT = "' + GetField('PEL_CODEELT') + '" AND PNR_NIVMAXPERSO IN ("ETB","POP","SAL") AND ##PNR_PREDEFINI##') then
        begin
          if (DOS = False) then
            MenuDupli.Items[1].Visible := False
          else
            MenuDupli.Items[1].Visible := True;
        end
        else
          MenuDupli.Items[1].Visible := False;
      end;
    end;
  end;
end;
//FIN PT26

//DEB PT27
procedure TOM_EltNationaux.DupliquerEltDossier(Sender: TObject);
var
  Retour : String;
begin
  Retour := AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION;' + GetField('PEL_CODEELT') + ';DUPLI');
  // Chainer automatiquement sur la saisie des éléments dossier
  if (Retour = 'OK') then
    AglLanceFiche('PAY', 'ELTNATDOS_MUL', '', '', 'ELT;' + GetField('PEL_CODEELT'));
end;
//FIn PT27

//DEB PT28
procedure TOM_EltNationaux.MajFiltreNiveau(CodeElt : String);
var
  Q : TQuery;
  {$IFNDEF EAGLCLIENT}
  Predefini : THDBValComboBox;
  {$ELSE}
  Predefini : THValComboBox;
  {$ENDIF}
  NiveauOK : String;
begin
  if CodeElt <> '' then
  begin
    NiveauOK := '"CEG","STD","DOS"';
    {$IFNDEF EAGLCLIENT}
    Predefini := THDBValComboBox(GetControl('PEL_PREDEFINI'));
    {$ELSE}
    Predefini := THValComboBox(GetControl('PEL_PREDEFINI'));
    {$ENDIF}
    if Predefini <> nil then
    begin
      Q := OpenSQL('SELECT PNR_TYPENIVEAU,PNR_NIVMAXPERSO FROM ELTNIVEAUREQUIS WHERE PNR_CODEELT="' + CodeElt + '" AND ##PNR_PREDEFINI## ORDER BY PNR_PREDEFINI DESC', True);
      if not Q.Eof then
      begin
        if (Q.Fields[1].AsString = 'CEG') then
          NiveauOK := '"CEG"';
        if (Q.Fields[1].AsString = 'STD') then
          NiveauOK := '"CEG","STD"';
      end;
      Predefini.Plus := ' AND CO_CODE IN (' + NiveauOK + ')';
    end;
    Ferme(Q);
  end;
end;
//FIN PT28

initialization
  registerclasses([TOM_EltNationaux]);


end.

