{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion des variables
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 17/12/2002 Ph V591 Affectation tablette cotisation si variable renvoie
                           Taux pat ou sal
PT2   : 06/01/2003 SB V591 FQ 10314 Affectation tablette base cotisation
                           manquante aprés affectation type de base de variable
PT3   : 31/01/2003 PH V591 suppression des bouton insertion/supp et zoom specif
                           à la paie et non utilisés
PT4   : 07/02/2003 PH V_42 Traitement des tables divers DIW
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT5   : 22/07/2003 SB V_42 FQ 10376 Intégration de la duplication des variables
PT5-1 : 11/09/2003 SB V_42 FQ 10376 Restriction acces bouton dupliquer
PT6   : 16/12/2003 PH V_50 FQ 11018  Imputation de la tablette sur le champs
                           typebase au lieu de base
PT7   : 25/03/2004 PH V_50 PVA_MTARRONDI à 7 pour avoir un maximum de précision
PT8   : 04/10/2004 PH V_50 Suppression rechargement tablette
PT9   : 08/10/2004 PH V_50 FQ 11446 Ellipsis si type base est de nature Date
PT10  : 08/11/2004 PH V_60 FQ 11757 Controle des dates de validité
PT11  : 20/01/2005 PH V_60 FQ 11941 Revu test de cohérence sur variable de test
                           sur des dates
PT12  : 30/05/2005 PH V_60 FQ 12318 CWAS controles différents
PT13  : 25/04/2006 NA V_65 FQ 11911 les conditions SI, ALORS, SINON doivent se
                           terminer par FIN
PT14  : 27/04/2006 NA V_65 Le type de base et la base ne sont pas contrôlés dans
                           la condition SI
PT15  : 06/04/2007 GGU V_72 Gestion des variables liées aux tables dynamiques
PT16  : 26/04/2007 GGS V_72 Gestion de Trace
PT17  : 07/05/2007 PH V_72 Concept Paie
PT18  : 22/05/2007 PH V_72 FQ 14265 prise en compte des Combo et Edit
PT19  : 29/05/2007 GGS V_72 Gestion de Trace Duplication
PT20  : 30/05/2007 VG V_72 Correction requête avec left join where or qui plante
                           sous Oracle
PT21  : 21/06/2007 FC V_72 FQ 14434 Tester l'existence de l'élément dynamique avant d'utiliser la table dynamique
PT22  : 26/06/2007 GGU V_72 FQ 14452 le libellé de la table dynamique ne s'affiche pas
PT23  : 11/07/2007 FC V_72 FQ 14148 Concepts
PT24  : 16/07/2007 GGU V_8 Gestion des variables de présence
PT25  : 04/09/2007 GGU V_8 Présence : gestion de la période de calcul "Semaines précédentes du mois"
PT26  : 17/09/2007 FC V_80 FQ 14757 Message erreur en validation variable de paie cumul suite duplication
PT27  : 20/09/2007 NA V_80 Accés au ompteur de présence associée à la variable
PT28  : 21/09/2007 NA V_80 FQ 14804 Les variables de type CALCUL doit se terminer par FIN
PT29  : 03/10/2007 GGU V_80 FQ 14313 Gestion des éléments dynamiques et des zones libres salarié
PT30  : 06/11/2007 FC V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
PT31  : 04/12/2007 GGU V_80 FQ 15010 Afficher les tables dynamiques, quelque soit leurs niveau de saisie
                   établissement ou convention collective si on saisie une variable prédéfini cegid ou
                   standard. Ne faire le filter que pour les prédéfinis dossier.
PT32  : 11/12/2007 GGU V_80 FQ 14006 impossible de valider une variable nature test contenant 4 conditions
PT33  : 18/04/2008 GGU V_81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
PT37  : 03/07/2008 PH V_81 FQ 15577 CWAS non RAZ nbre de mois glissant dans le cas de variable de type cumul
}
unit UTOMVariablePaie;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, DBCtrls, Fe_Main, Fiche,
{$ELSE}
  MaineAgl, eFiche,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97,
  ParamSoc, HPanel, ParamDat, Dialogs, PgOutils, Menus, Windows, P5Util, PgOutils2,
  PAIETOM;

type
  PVarOnglet = class
    LBASE: string;
    BDD_TYPEBASE: string;
    BDD_BASE: string;
    BDD_OPERATMATH: string;
    max: integer;
  end;

type

  TOM_VariablePaie = class(PGTOM) //PT16 TOM devient PGTOM
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure ControleAncienneTablette(Champ, Tablette: string; modif: boolean);
    procedure ControleOuvreLigne(MaCombo: THDBValComboBox; TypeZ: string; top: integer);
    procedure AfficheTypeBase(TypeBase, libbase, Base: string);
    procedure EnableChampsOngletCumul(Nat: string);
    procedure ControleZonesOngletValeur;
    procedure ControleZonesOngletCumul;
    procedure ControleZonesOngletTest;
    procedure ControleZonesOngletCalcul;
    procedure InitialiseChampsTable;
    function IsOperateurInegalite(Operateur: string): boolean;
    function IsOperateurLogique(Operateur: string): boolean;
    function ControleSaisieBaseOk(RgErr: string; TypeZ: string): boolean;
    procedure OnChangeOnglet_PCALCUL(F: TField);
    procedure OnChangeOnglet_PTEST(F: TField);
    procedure AlimenteNomsZones(var VarOnglet: PVarOnglet; Typez: string);
    function RechercheTablette(Base: TControl): string;
    procedure AfficheLibelle(Base, libBase: string);
    //       function ColleZeroDevant(Nombre , LongChaine : integer) : string;
    procedure DateElipsisclick(Sender: TObject);
    function IsVariabledate(code: string): boolean;
    function IsOperateurDate(Operateur: string): boolean;
    procedure DupliquerVariable(Sender: TObject); { PT5 }
    procedure OnElipsisClickBase(Sender: TObject); //PT15
    procedure OnElipsisClickBaseEltDyna(Sender: TObject); //PT29
    procedure OnElipsisClickBaseZoneLibreSal(Sender: TObject); //PT29
    procedure OnChangePeriodicite(Sender : TObject);
  private
    iNouveau: integer;
    BoolInit: boolean; // si true, initialisation champs après modif opérateur
    iMaxEpva_Base: integer; //nbre de champs TPVA_BASE de type calcul
    iMaxTpva_Base: integer; //nbre de champs TPVA_BASE
    iMaxTpva_resthen: integer; //nbre de champs TPVA_RESTHEN
    iMaxTPVA_RESELSE: integer; //nbre de champs TPVA_RESELSE
    iMaxTpva_operatmath: integer; //nbre de champs TPVA_OPERATMATH
    iMaxTpva_Operattest: integer; //nbre de champs TPVA_OPERATTEST
    LectureSeule, CEG, STD, DOS, OnFerme: boolean;
    LaNature, Mode: string; { PT5 }
    DerniereCreate: string;
    ListeEtab, ListeConv: string; //PT15
    HadSearchEtabConv: Boolean; //PT15
    Trace: TStringList; //PT16
    GblTypeVariable: string; //PT15
    LeStatut:TDataSetState; //PT30
    procedure RechercheEtabConv; //PT15
    procedure InitModePresence; //PT24
    procedure compteurassocie(Sender : TObject); // PT27
  end;

  function RechercheNomTablette(TypeBase, NatureVariable: String; Presence : Boolean): string;

implementation
uses P5Def, Lookup, StrUtils, PGPresence
  , PGTablesDyna //PT33
  ;

procedure TOM_VariablePaie.OnArgument(stArgument: string);
var
  Btn: TToolBarButton97;
begin
  inherited;
  LaNature := readtokenst(stArgument);
  Mode := readtokenst(stArgument);
  if pos('ACTION', Mode) = 0 then //Le parametre mode n'est pas présent quand on ouvre une variable existante
  begin
    GblTypeVariable := Mode;
    Mode := '';
  end;
  //PT24 : Gestion des variables de présence
  //L'attribut Type variable est toujours le dernier. Il y a parfois un attribut MONOFICHE avant
  //GblTypeVariable := readtokenst(stArgument); //PT15
  While stArgument <> '' do
    GblTypeVariable := readtokenst(stArgument);
  //Seul les valeurs PRE et PAI sont admises ('PAI' par défaut)
  if GblTypeVariable <> 'PRE' then GblTypeVariable := 'PAI';
  SetControlEnabled('PVA_NATUREVAR', FALSE);
  iMaxEpva_Base := 10;
  iMaxTpva_Base := 8;
  iMaxTpva_resthen := 4;
  iMaxTpva_reselse := 4;
  iMaxTpva_operatmath := 8;  //PT32 on passe de 7 à 8
  iMaxTpva_Operattest := 3;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER')); { PT5 }
  if Btn <> nil then Btn.OnClick := DupliquerVariable; { PT5 }
  HadSearchEtabConv := False; // PT15
  PaieConceptPlanPaie(Ecran); // PT17
  //PT24 : Gestion des variables de présence
  if GblTypeVariable = 'PRE' then
  begin
    InitModePresence;
    Btn := TToolBarButton97(GetControl('BCOMPTEURPRES')); { PT27 }
    if Btn <> nil then Btn.OnClick := CompteurAssocie; { PT27 }
  end else begin
    SetControlProperty('PVA_VARIABLE', 'EditMask', '9999');
    SetControlProperty('PVA_THEMEVAR', 'DataType', 'PGTHEMEVAR');
  end;
end;

procedure TOM_VariablePaie.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  DS.edit;
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

procedure TOM_VariablePaie.OnNewRecord;
begin
  inherited;
  if mode = 'DUPLICATION' then exit; { PT5 }
  iNouveau := 9999;
  SetField('PVA_NATUREVAR', LaNature);
  SetControlEnabled('PVA_NATUREVAR', FALSE);
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  SetField('PVA_NODOSSIER', PgRendNoDossier());
  SetField('PVA_PREDEFINI', 'DOS');
  // PT7 25/03/2004 V_50 PH PVA_MTARRONDI à 7 pour avoir un maximum de précision
  SetField('PVA_MTARRONDI', 7);
  //PT24 : Gestion des variables de présence
  SetField('PVA_TYPEVARIABLE', GblTypeVariable); //PT15
end;

procedure TOM_VariablePaie.OnLoadRecord;
var
  Bool: boolean;
begin
  inherited;

  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  if mode = 'DUPLICATION' then exit; { PT5 }

  if iNouveau = 9999 then Bool := True else Bool := False;
  SetControlEnabled('PVA_VARIABLE', Bool);
  //  SetControlEnabled ('TPVA_VARIABLE', Bool) ;
  BoolInit := true;

  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;
  if (Getfield('PVA_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('BDelete', CEG);
  end
  else
    if (Getfield('PVA_PREDEFINI') = 'STD') then
    begin
      LectureSeule := (STD = False);
      PaieLectureSeule(TFFiche(Ecran), (STD = False));
      SetControlEnabled('BDelete', STD);
    end
    else
      if (Getfield('PVA_PREDEFINI') = 'DOS') then
      begin
        LectureSeule := (DOS = False);     //PT23
        PaieLectureSeule(TFFiche(Ecran), (DOS = False));  //PT23
        SetControlEnabled('BDelete', DOS);
      end;

  SetControlEnabled('BInsert', True);
  SetControlEnabled('PVA_PREDEFINI', False);
  SetControlEnabled('PVA_VARIABLE', False);
  SetControlEnabled('BDUPLIQUER', not ((GetField('PVA_NATUREVAR') = 'VAL') and (GetField('PVA_PREDEFINI') = 'CEG'))); { PT5-1 }

  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PVA_PREDEFINI', True);
    SetControlEnabled('PVA_VARIABLE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
    SetControlEnabled('BDUPLIQUER', False); { PT5 }
  end;
  if Mode = 'CONSULTATION' then
    PaieLectureSeule(TFFiche(Ecran), True);
  //PT24 Gestion des variables de présence
  OnChangePeriodicite(Self);
end;

procedure TOM_VariablePaie.OnUpdateRecord;
var
  Nat: string;
  iCode: integer;
  Predef: string;
  Retour: Boolean;
  Q : TQuery;
  DD, DF : TDatetime; // pt24
begin
  inherited;
  OnFerme := False;
  LeStatut := DS.State; //PT30
  if (DS.State in [dsInsert]) then DerniereCreate := GetField('PVA_VARIABLE')
  else if (DerniereCreate = GetField('PVA_VARIABLE')) then OnFerme := True;
  if mode = 'DUPLICATION' then exit; { PT5 }
  if (DS.State = dsinsert) then
  begin
    if (GetField('PVA_PREDEFINI') <> 'DOS') then SetField('PVA_NODOSSIER', '000000')
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    else SetField('PVA_NODOSSIER', PgRendNoDossier());
  end;

  //PT24 : Gestion des variables de présence
  if GblTypeVariable <> 'PRE' then
  begin
    // complète par des zeros devant si nécessaire
    if ((isnumeric(GetControlText('PVA_VARIABLE'))) and (GetControlText('PVA_VARIABLE') <> '    ')) then
    begin
      iCode := strtoint(trim(GetControlText('PVA_VARIABLE')));
      SetField('PVA_VARIABLE', ColleZeroDevant(iCode, length(GetControlText('PVA_VARIABLE'))));
    end;
  end;

  Nat := GetField('PVA_NATUREVAR');
  if Nat = 'VAL' then ControleZonesOngletValeur
  else if LaNature = 'CAL' then ControleZonesOngletCalcul
  else if ((Nat = 'CUM') or (Nat = 'REM') or (Nat = 'COT') or (Nat = 'CUP')) then
    ControleZonesOngletCumul  //PT24
  else if LaNature = 'TES' then ControleZonesOngletTest;

  Predef := GetField('PVA_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PVA_PREDEFINI');
  end;
  if ((GblTypeVariable = 'PRE') and (GetControlText('PVA_VARPERIODICITE') = '')) then
  begin
    LastError := 1;
    LastErrorMsg := 'Vous devez renseigner la périodicité';
    SetFocusControl('PVA_VARPERIODICITE');
  end;               
  // DEB PT10
  if (Nat = 'VAL') or (Nat = 'CUM') or (Nat = 'COT') or (Nat = 'REM') or (Nat = 'CUP') then  //PT24
  begin
    Retour := TRUE; // PT12
    if Nat = 'VAL' then
    begin
      if GetControlText('PVA_DATEDEBUT') <> '    ' then
{$IFNDEF EAGLCLIENT}
        Retour := ControleDateCourte(THDbEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', true);
{$ELSE}
        Retour := ControleDateCourte(THEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', true);
{$ENDIF}
    end
    else
    begin
      if (GetControlText('PVA_DATEDEBUT') <> '') then
{$IFNDEF EAGLCLIENT}
        Retour := ControleDateCourte(THDbEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', true);
{$ELSE}
        Retour := ControleDateCourte(THEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', true);
{$ENDIF}
      if (GetControlText('PVA_DATEFIN') <> '') and (Retour) and (GetControlText('PVA_DATEDEBUT') <> '    ') then  //PT26
{$IFNDEF EAGLCLIENT}
        Retour := ControleDateCourte(THDbEdit(getcontrol('PVA_DATEFIN')), 'PVA_DATEFIN', false);
{$ELSE}
        Retour := ControleDateCourte(THEdit(getcontrol('PVA_DATEFIN')), 'PVA_DATEDEBUT', false);
{$ENDIF}
    end;
    if not Retour then
    begin
      LastError := 2;
      LastErrorMsg := 'Vos périodes de calcul ou de validité ne sont pas valides';
      SetFocusControl('PVA_DATEDEBUT');
    end;
  end;
  // FIN PT10

  //DEB PT21
  if GetControlText('PVA_TYPEBASE0') = '225' then
  begin
    Q:= OpenSQL('SELECT PTE_NATURETABLE FROM TABLEDIMENT WHERE ##PTE_PREDEFINI## AND PTE_CODTABL = "' + GetControlText('PVA_BASE0') + '"', true);
    if (Q.findfield('PTE_NATURETABLE').AsString = 'COD') then
    begin
      if not ExisteSQL('SELECT PPP_CODTABL FROM PARAMSALARIE WHERE PPP_CODTABL ="' + GetControlText('PVA_BASE0') + '"') then
      begin
        LastError := 3;
        LastErrorMsg := 'Il n''existe pas d''élément dynamique correspondant à ce code table dynamique';
        SetFocusControl('PVA_BASE0');
      end;
    end;
  end;
  //FIN PT21

    //Rechargement des tablettes
  //if (LastError = 0) and (Getfield('PVA_VARIABLE') <> '') and (Getfield('PVA_LIBELLE') <> '') then
    // PT8    ChargementTablette(TFFiche(Ecran).TableName, '');

  If (GblTypeVariable = 'PRE') and (LastError = 0) and (DS.State = dsEdit) then
  begin
    PresenceDonneMoisCalculActuel (DD,DF);   //PT24
    CompteursARecalculer(DD); //PT24
  end;
end;

procedure TOM_VariablePaie.OnChangeField(F: TField);
var
  Mpredefini: THDBValComboBox;
  i: integer;
  j: string;
  iCode: integer;
  sCode, variable, mes, Pred: string;
  numCode : Integer;
//  Bool: Boolean;
begin
  inherited;
  //------------------------------------------------------------------------------
  // détermination de l'onglet à afficher + affichages des libellés dynamiques
  if mode = 'DUPLICATION' then exit; { PT5 }
  numCode := 0;
  if (F.FieldName = 'PVA_VARIABLE') then
  begin
    //PT24 : Gestion des variables de présence
    Variable := Getfield('PVA_VARIABLE');
    Pred := GetField('PVA_PREDEFINI');
    if Variable = '' then exit;
    if (GblTypeVariable = 'PRE') and (Trim(variable) <> '') then
    begin
      if RightStr(Variable, 3) <> '   ' then
      begin
        try
          numCode := strToInt(RightStr(Variable, 3));
        except
          PGIBox('Le code doit être de la forme Pxxx, avec xxx un nombre à 3 chiffres.', 'Code erroné');
          SetField('PVA_VARIABLE', '');
          SetFocusControl('PVA_VARIABLE');
          exit;
        end;
        if not TestRubrique(Pred, intToStr(numCode), 200) then
        begin
          Mes := MesTestRubrique('PVA', Pred, 200);
          HShowMessage('2;Code Erroné: ' + intToStr(numCode) + ' ;' + Mes + ';W;O;O;;;', '', '');
          SetField('PVA_VARIABLE', '');
          if Pred <> GetField('PVA_PREDEFINI') then SetField('PVA_PREDEFINI', pred);
          SetFocusControl('PVA_VARIABLE');
          exit;
        end;
      end;
    end else begin
      if ((isnumeric(Variable)) and (Variable <> '    ')) then
      begin
        iCode := strtoint(trim(Variable));
        sCode := ColleZeroDevant(iCode, 4);
        if (DS.State = dsinsert) and (sCode <> '') and (GetField('PVA_PREDEFINI') <> '') then
        begin
          if (TestRubrique(GetField('PVA_PREDEFINI'), sCode, 1000) = False) or (sCode = '0000') then
          begin
            Mes := MesTestRubrique('PVA', GetField('PVA_PREDEFINI'), 1000);
            HShowMessage('2;Code Erroné: ' + sCode + ' ;' + Mes + ';W;O;O;;;', '', '');
            SCode := '';
          end;
        end;
        if sCode <> Variable then
        begin
          SetField('PVA_VARIABLE', scode);
          SetFocusControl('PVA_VARIABLE');
        end;
      end;
    end;
  end;

  if (F.FieldName = 'PVA_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PVA_PREDEFINI');
    variable := GetField('PVA_VARIABLE');
    if Pred = '' then exit;

    AccesPredefini('TOUS', CEG, STD, DOS);

    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de variable prédéfini CEGID', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PVA_PREDEFINI', 'Value', Pred);
    end;

    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de variable prédéfini Standard', 'Accès refusé');
      Pred := 'DOS';
      SetControlProperty('PVA_PREDEFINI', 'Value', Pred);
    end;

    if (GblTypeVariable = 'PRE') and (Trim(variable) <> '') and (Pred <> '') then
    begin
      try
        numCode := strToInt(RightStr(Variable, 3));
      except
        PGIBox('Le code doit être de la forme Pxxx, avec xxx un nombre à 3 chiffres.', 'Code erroné');
      end;
      if not TestRubrique(Pred, intToStr(numCode), 200) then
      begin
        Mes := MesTestRubrique('PVA', Pred, 200);
        HShowMessage('2;Code Erroné: ' + intToStr(numCode) + ' ;' + Mes + ';W;O;O;;;', '', '');
        SetField('PVA_VARIABLE', '');
        if Pred <> GetField('PVA_PREDEFINI') then SetField('PVA_PREDEFINI', pred);
        SetFocusControl('PVA_VARIABLE');
        exit;
      end;
    end else begin
      if (Trim(variable) <> '') and (Pred <> '') then
      begin
        if (TestRubrique(Pred, variable, 1000) = False) or (variable = '0000') then
        begin
          Mes := MesTestRubrique('PVA', Pred, 1000);
          HShowMessage('2;Code Erroné: ' + variable + ' ;' + Mes + ';W;O;O;;;', '', '');
          SetField('PVA_VARIABLE', '');
          if Pred <> GetField('PVA_PREDEFINI') then SetField('PVA_PREDEFINI', pred);
          SetFocusControl('PVA_VARIABLE');
          exit;
        end;
      end;
    end;
    if Pred <> GetField('PVA_PREDEFINI') then SetField('PVA_PREDEFINI', pred);
  end;

  if (ds.state in [dsBrowse]) and (F.FieldName = 'PVA_PREDEFINI') then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PVA_PREDEFINI', False);
    SetControlEnabled('PVA_VARIABLE', False);
    SetControlEnabled('BDUPLIQUER', not ((GetField('PVA_NATUREVAR') = 'VAL') and (GetField('PVA_PREDEFINI') = 'CEG'))); { PT5-1 }
  end;

  if LaNature = 'VAL' then
  begin
    if F.FieldName = 'PVA_TYPEBASE0' then AfficheTypeBase(F.FieldName, 'TPVA_BASE0', 'PVA_BASE0');
    //ajout Label associé au choix ellipsis
    if F.FieldName = 'PVA_BASE0' then AfficheLibelle('PVA_BASE0', 'TPVA_BASE0');
    // contrôle date jj/mm
    if F.FieldName = 'PVA_DATEDEBUT' then
      if GetControlText('PVA_DATEDEBUT') <> '    ' then
{$IFNDEF EAGLCLIENT}
        ControleDateCourte(THDbEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', true);
{$ELSE}
        ControleDateCourte(THEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', true);
{$ENDIF}
  end
  else
    if ((LaNature = 'CUM') or (LaNature = 'COT') or (LaNature = 'REM') or (LaNature = 'CUP')) then //PT24
    begin
      if F.FieldName = 'PVA_TYPEBASE0' then AfficheTypeBase(F.FieldName, 'TPVA_BASE0', 'PVA_BASE0');
      if F.FieldName = 'PVA_BASE0' then AfficheLibelle('PVA_BASE0', 'TPVA_BASE0');
      if F.FieldName = 'PVA_PERIODECALCUL' then EnableChampsOngletCumul(LaNature);
    // Controle de cohérence date à date format jj/mm
      if (F.FieldName = 'PVA_DATEDEBUT') and (GetControlText('PVA_DATEDEBUT') <> '') then
{$IFNDEF EAGLCLIENT}
        ControleDateCourte(THDbEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', false);
{$ELSE}
        ControleDateCourte(THEdit(getcontrol('PVA_DATEDEBUT')), 'PVA_DATEDEBUT', false);
{$ENDIF}
      if (F.FieldName = 'PVA_DATEFIN') and (GetControlText('PVA_DATEFIN') <> '') then
{$IFNDEF EAGLCLIENT}
        ControleDateCourte(THDbEdit(getcontrol('PVA_DATEFIN')), 'PVA_DATEFIN', false);
{$ELSE}
        ControleDateCourte(THEdit(getcontrol('PVA_DATEFIN')), 'PVA_DATEDEBUT', false);
{$ENDIF}
    end
    else
      if LaNature = 'CAL' then OnChangeOnglet_PCALCUL(F)
      else
        if LaNature = 'TES' then OnChangeOnglet_PTEST(F);

  if (F.FieldName = 'PVA_NATUREVAR') then
  begin //     InitialiseChampsTable;
    if LaNature = 'ALE' then SetControlProperty('PGENERAL', 'TabVisible', TRUE)
    else SetControlProperty('PGENERAL', 'TabVisible', FALSE);

    Mpredefini := THDBValComboBox(getcontrol('PVA_PREDEFINI'));

    if ((LaNature = 'VAL') and (MPredefini <> nil)) then
    begin
      if MPredefini.Value = 'CEG' then
      begin
        if (GblTypeVariable = 'PRE') and (CEG) and (GetField('PVA_VARIABLE') > 'P200') then
        begin
          SetControlProperty('PPREDEFINI', 'TabVisible', FALSE);
          SetControlProperty('PVALEUR', 'TabVisible', TRUE);
        end else begin
          SetControlProperty('PPREDEFINI', 'TabVisible', TRUE);
          SetControlProperty('PVALEUR', 'TabVisible', FALSE);
        end;
      end else begin
        SetControlProperty('PPREDEFINI', 'TabVisible', FALSE);
        SetControlProperty('PVALEUR', 'TabVisible', TRUE);
      end;
//      Bool := ((LaNature = 'VAL') and (MPredefini.Value <> 'CEG'));
//      SetControlProperty('PVALEUR', 'TabVisible', Bool);
      SetControlProperty('PVA_BASE0', 'Datatype', RechercheTablette(GetControl('PVA_TYPEBASE0')));
      AfficheLibelle('PVA_BASE0', 'TPVA_BASE0');
    end;

    if ((LaNature = 'CUM') or (LaNature = 'COT') or (LaNature = 'REM') or (LaNature = 'CUP')) then  //PT24
    begin
      SetControlText('PCUMUL', RechDom('PGNATUREVARGBL', GetControlText('PVA_NATUREVAR'), FALSE));
      EnableChampsOngletCumul(LaNature);
      if getcontrolText('PVA_TYPEBASE0') = '' then SetField('PVA_TYPEBASE0', '');
      // PT6 16/12/2003 V_50 PH FQ 11018  Imputation de la tablette sur le champs typebase au lieu de base
      SetControlProperty('PVA_BASE0', 'Datatype', RechercheTablette(getcontrol('PVA_TYPEBASE0')));
      if LaNature = 'CUM' then SetControlVisible('PVA_TYPEBASE0', FALSE)
      else SetControlVisible('PVA_TYPEBASE0', TRUE);
      AfficheLibelle('PVA_BASE0', 'TPVA_BASE0');
      SetControlVisible('PVA_BASE0', TRUE);
    end;

    if (LaNature = 'CAL') then
    begin
      for i := 0 to (iMaxEpva_Base - 1) do
      begin
        j := inttostr(i);
        SetControlProperty('PVA_BASE' + j, 'Datatype', RechercheTablette(GetControl('PVA_TYPEBASE' + j)));
        AfficheLibelle('PVA_BASE' + j, 'TPVA_BASE' + j);
      end;
    end;

    if (LaNature = 'TES') then
    begin
      for i := 0 to (iMaxTpva_Base - 1) do
      begin
        j := inttostr(i);
        SetControlProperty('PVA_BASE' + j, 'Datatype', RechercheTablette(GetControl('PVA_TYPEBASE' + j)));
        AfficheLibelle('PVA_BASE' + j, 'TPVA_BASE' + j);
      end;
      for i := 0 to (iMaxTpva_resthen - 1) do
      begin
        j := inttostr(i);
        SetControlProperty('PVA_RESTHEN' + j, 'Datatype', RechercheTablette(GetControl('PVA_TYPERESTHEN' + j)));
        AfficheLibelle('PVA_RESTHEN' + j, 'TPVA_RESTHEN' + j);
      end;
      for i := 0 to (iMaxTpva_reselse - 1) do
      begin
        j := inttostr(i);
        SetControlProperty('PVA_RESELSE' + j, 'Datatype', RechercheTablette(GetControl('PVA_TYPERESELSE' + j)));
        AfficheLibelle('PVA_RESELSE' + j, 'TPVA_RESELSE' + j);
      end;
    end;
  end;
  SetControlEnabled('PVA_NATUREVAR', False);
end;

//------------------------------------------------------------------------------
// Onglet Calcul : traitement des OnChangeField
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.OnChangeOnglet_PCALCUL(F: TField);
var
  i: char;
begin
  if Pos('PVA_TYPEBASE', F.FieldName) > 0 then
  begin
    i := F.FieldName[length(F.FieldName)];
    AfficheTypeBase(F.FieldName, 'TPVA_BASE' + i, 'PVA_BASE' + i);
  end
  else
    //ajout Label associé au choix ellipsis
    if Pos('PVA_BASE', F.FieldName) > 0 then
    begin
      i := F.FieldName[length(F.FieldName)];
      AfficheLibelle('PVA_BASE' + i, 'TPVA_BASE' + i);
    end
    else
      if Pos('PVA_OPERATMATH', F.FieldName) > 0 then
      begin
        i := F.FieldName[length(F.FieldName)];
    // dans le cas où on est en insertion de ligne, ne pas tout remettre à blanc
    // pendant le décalalge des lignes.
        ControleOuvreLigne(THDBValComboBox(GetControl('PVA_OPERATMATH' + i)), 'CALC', strtoint(i) + 1);
      end;
end;

//------------------------------------------------------------------------------
// Onglet Test : traitement des OnChangeField
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.OnChangeOnglet_PTEST(F: TField);
var
  i: char;
begin
  //partie Si
  if Pos('PVA_TYPEBASE', F.FieldName) > 0 then
  begin
    i := F.FieldName[length(F.FieldName)];
    AfficheTypeBase(F.FieldName, 'TPVA_TYPEBASE' + i, 'PVA_BASE' + i);
  end
  else
    //ajout Label associé au choix ellipsis
    if Pos('PVA_BASE', F.FieldName) > 0 then
    begin
      i := F.FieldName[length(F.FieldName)];
      AfficheLibelle('PVA_BASE' + i, 'TPVA_BASE' + i);
    end
    else
    // partie Alors
      if Pos('PVA_TYPERESTHEN', F.FieldName) > 0 then
      begin
        i := F.FieldName[length(F.FieldName)];
        AfficheTypeBase(F.FieldName, 'TPVA_TYPERESTHEN' + i, 'PVA_RESTHEN' + i);
      end
      else
    //ajout Label associé au choix ellipsis
        if Pos('PVA_RESTHEN', F.FieldName) > 0 then
        begin
          i := F.FieldName[length(F.FieldName)];
          AfficheLibelle('PVA_RESTHEN' + i, 'TPVA_RESTHEN' + i);
        end
        else
    // partie Sinon
          if Pos('PVA_TYPERESELSE', F.FieldName) > 0 then
          begin
            i := F.FieldName[length(F.FieldName)];
            AfficheTypeBase(F.FieldName, 'TPVA_TYPERESELSE' + i, 'PVA_RESELSE' + i);
          end
          else
    //ajout Label associé au choix ellipsis
            if Pos('PVA_RESELSE', F.FieldName) > 0 then
            begin
              i := F.FieldName[length(F.FieldName)];
              AfficheLibelle('PVA_RESELSE' + i, 'TPVA_RESELSE' + i);
            end
            else
              if Pos('PVA_OPERATMATH', F.FieldName) > 0 then
              begin
                i := F.FieldName[length(F.FieldName)];
                ControleOuvreLigne(THDBValComboBox(GetControl('PVA_OPERATMATH' + i)), 'TESTSI', strtoint(i) + 1);
              end
              else
    // opérateur pavé Alors
                if Pos('PVA_OPERATRESTHEN', F.FieldName) > 0 then
                begin
                  i := F.FieldName[length(F.FieldName)];
                  ControleOuvreLigne(THDBValComboBox(GetControl('PVA_OPERATRESTHEN' + i)), 'TESTALORS', strtoint(i) + 1);
                end
                else
    // opérateur pavé Sinon
                  if Pos('PVA_OPERATRESELSE', F.FieldName) > 0 then
                  begin
                    i := F.FieldName[length(F.FieldName)];
                    ControleOuvreLigne(THDBValComboBox(GetControl('PVA_OPERATRESELSE' + i)), 'TESTSINON', strtoint(i) + 1);
                  end;
end;

//------------------------------------------------------------------------------
//- détermine, pour l'onglet cumul, la visibilité et les options de saisie pour
//- champs de cet onglet
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.EnableChampsOngletCumul(Nat: string);
var
  MVPVA_TYPEBASE: THDBValComboBox;
  MCPVA_PERIODECALCUL: THDBValComboBox;
  MCPVA_ANNEE: THDBValComboBox;
  MCPVA_TRIMESTRE: THDBValComboBox;
  MCPVA_MOISCALCUL: THDBValComboBox;
  MCPVA_DATEDEBUT: THDBValComboBox;
  MCPVA_DATEFIN: THDBValComboBox;
  Bool: boolean;
  Lib: string;
begin
  MVPVA_TYPEBASE := THDBValComboBox(GetControl('PVA_TYPEBASE0'));
  if ((LaNature = 'CUM') or (LaNature = 'CUP')) then Lib := 'Cumul' //PT24
  else if LaNature = 'COT' then Lib := 'Type de base';
  SetControlText('TPVA_TYPEBASE0', Lib);
  if MVPVA_TYPEBASE <> nil then
  begin
    SetControlVisible('PVA_TYPEBASE0', TRUE);
    if LaNature = 'COT' then MVPVA_TYPEBASE.DataType := 'PGTYPECUMCOT'
    else if LaNature = 'REM' then MVPVA_TYPEBASE.DataType := 'PGTYPECUMREM'
    else if (LaNature = 'CUM') then
    begin
      SetControlVisible('PVA_TYPEBASE0', FALSE);
      SetControlProperty('PVA_BASE0', 'datatype', 'PGCUMULPAIE');
    end else if (LaNature = 'CUP') then  //PT24 : Gestion des variables de présence
    begin
      SetControlVisible('PVA_TYPEBASE0', True);
    end;
  end;
  // ouvrir la / les combo(s) correspondant(es) à la saisie de la période
  MCPVA_PERIODECALCUL := THDBValComboBox(GetControl('PVA_PERIODECALCUL'));
  MCPVA_DATEDEBUT := THDBValComboBox(GetControl('PVA_DATEDEBUT'));
  MCPVA_DATEFIN := THDBValComboBox(GetControl('PVA_DATEFIN'));
  MCPVA_ANNEE := THDBValComboBox(GetControl('PVA_ANNEE'));
  MCPVA_TRIMESTRE := THDBValComboBox(GetControl('PVA_TRIMESTRE'));
  MCPVA_MOISCALCUL := THDBValComboBox(GetControl('PVA_MOISCALCUL'));
  if MCPVA_PERIODECALCUL <> nil then
  begin
    Bool := (   (MCPVA_PERIODECALCUL.value = '006')
             or (MCPVA_PERIODECALCUL.value = '010')
             or (MCPVA_PERIODECALCUL.value = '011'));  //PT24
    SetControlEnabled('PVA_NBREMOISGLISS', Bool);
    if (Assigned(MCPVA_ANNEE) and not MCPVA_ANNEE.enabled) then
      if (Not Bool) then  SetField('PVA_NBREMOISGLISS', 0);   // PT37

    Bool := (MCPVA_PERIODECALCUL.value = '010');   //PT24
    if Bool then SetControlCaption('TPVA_NBREMOISGLISS', 'Nombre de mois glissants');
    Bool := (MCPVA_PERIODECALCUL.value = '011');   //PT24
    if Bool then SetControlCaption('TPVA_NBREMOISGLISS', 'Nombre de semaines glissantes');

    Bool := ((MCPVA_PERIODECALCUL.value = '004') or (MCPVA_PERIODECALCUL.value = '012'));   //PT24
    SetControlEnabled('PVA_DATEDEBUT', Bool);
    if not MCPVA_DATEDEBUT.enabled then SetField('PVA_DATEDEBUT', '');
    Bool := (MCPVA_PERIODECALCUL.value = '004');
    SetControlEnabled('PVA_DATEFIN', Bool);
    if not MCPVA_DATEFIN.enabled then SetField('PVA_DATEFIN', '');

    Bool := (MCPVA_PERIODECALCUL.value = '003');
    SetControlEnabled('PVA_ANNEE', Bool);
    SetControlEnabled('TPVA_ANNEE', Bool);
    if Assigned(MCPVA_ANNEE) and not MCPVA_ANNEE.enabled then SetField('PVA_ANNEE', '');

    Bool := (MCPVA_PERIODECALCUL.value = '002');
    SetControlEnabled('PVA_TRIMESTRE', Bool);
    SetControlEnabled('TPVA_TRIMESTRE', Bool);
    if Assigned(MCPVA_TRIMESTRE) and not MCPVA_TRIMESTRE.enabled then SetField('PVA_TRIMESTRE', '');

    Bool := (MCPVA_PERIODECALCUL.value = '001');
    SetControlEnabled('PVA_MOISCALCUL', Bool);
    SetControlEnabled('TPVA_MOISCALCUL', Bool);
    if Assigned(MCPVA_MOISCALCUL) and not MCPVA_MOISCALCUL.enabled then SetField('PVA_MOISCALCUL', '');

  end;
end;
//------------------------------------------------------------------------------
// Controles de cohérence des zones de l'onglet Valeur
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.ControleZonesOngletValeur;
begin
  if (GetField('PVA_BASE0') = '') then
  begin
    LastError := 102;
    LastErrorMsg := 'Vous devez renseigner une base';
    SetFocusControl('PVA_BASE0');
  end;
  if (GetField('PVA_TYPEBASE0') = '') then
  begin
    LastError := 101;
    LastErrorMsg := 'Vous devez renseigner un Type de base';
    SetFocusControl('PVA_TYPEBASE0');
  end;
end;

//------------------------------------------------------------------------------
// Controles de cohérence des zones de l'onglet Cumul
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.ControleZonesOngletCumul;
var
  Periode: string;
begin
  if (GetField('PVA_BASE0') = '') then
  begin
    LastError := 201;
    LastErrorMsg := 'Vous devez renseigner une base';
    SetFocusControl('PVA_BASE0')
  end;
  if ((LaNature = 'COT') or (LaNature = 'REM') or (LaNature = 'CUP')) then
  begin
    if (GetField('PVA_TYPEBASE0') = '') then
    begin
      LastError := 202;
      LastErrorMsg := 'Vous devez renseigner un Type de base';
      SetFocusControl('PVA_TYPEBASE0');
    end;
  end;
  Periode := GetField('PVA_PERIODECALCUL');
  if (Periode = '') then
  begin
    LastError := 203;
    LastErrorMsg := 'Vous devez renseigner une période de calcul';
    SetFocusControl('PVA_PERIODECALCUL');
  end
  else
  begin
    if Periode <> '' then
    begin
      if Periode = '003' then
      begin
        if (GetField('PVA_ANNEE') = '') then
        begin
          LastError := 204;
          LastErrorMsg := 'Vous devez définir l''année de calcul';
          SetFocusControl('PVA_ANNEE');
        end;
      end else if Periode = '002' then
      begin
        if (GetField('PVA_TRIMESTRE') = '') then
        begin
          LastError := 205;
          LastErrorMsg := 'Vous devez définir le trimestre de calcul';
          SetFocusControl('PVA_TRIMESTRE');
        end;
      end else if Periode = '001' then
      begin
        if (GetField('PVA_MOISCALCUL') = '') then
        begin
          LastError := 206;
          LastErrorMsg := 'Vous devez définir le mois de calcul';
          SetFocusControl('PVA_MOISCALCUL');
        end;
      end else if Periode = '004' then
      begin
        if (GetField('PVA_DATEFIN') = '') then
        begin
          LastError := 208;
          LastErrorMsg := 'Vous devez définir la date de fin';
          SetFocusControl('PVA_DATEFIN');
        end;
        if (GetField('PVA_DATEDEBUT') = '') then
        begin
          LastError := 207;
          LastErrorMsg := 'Vous devez définir la date de début';
          SetFocusControl('PVA_DATEDEBUT');
        end;
      end else if Periode = '012' then
      begin
        if (GetField('PVA_DATEDEBUT') = '') then
        begin
          LastError := 207;
          LastErrorMsg := 'Vous devez définir la date de début';
          SetFocusControl('PVA_DATEDEBUT');
        end;
      end else if Periode = '010' then
      begin
        if (GetControlText('PVA_NBREMOISGLISS') = '') then
        begin
          LastError := 207;
          LastErrorMsg := 'Vous devez définir le nombre de mois glissants';
          SetFocusControl('PVA_NBREMOISGLISS');
        end;
      end else if Periode = '011' then
      begin
        if (GetControlText('PVA_NBREMOISGLISS') = '') then
        begin
          LastError := 207;
          LastErrorMsg := 'Vous devez définir le nombre de semaines glissantes';
          SetFocusControl('PVA_NBREMOISGLISS');
        end;
      end; // fin periode = '011'
    end;
  end; // fin si periode non nulle
end;

//------------------------------------------------------------------------------
// Controles de cohérence des zones de l'onglet Test
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.ControleZonesOngletTest;
var
  i, j: integer;
  RgErr: string;
  Operateur, Operateurj: string;
  OperateurValid: boolean;
  MTypeBase, MTypeBaseP, MTypeBaseS: THDBValComboBox;
{$IFNDEF EAGLCLIENT}
  MBaseP, MBaseS: THDbEdit;
  MBase: THDbEdit;
{$ELSE}
  MBaseP, MBaseS: THEdit;
  MBase: THEdit;
{$ENDIF}
begin
  RgErr := '410';
  //----1 : controle de validite des opérateurs

  for i := 0 to (iMaxTpva_operatmath - 1) do //for 1
  begin
    Operateur := GetField('PVA_OPERATMATH' + inttostr(i));
    // Est ce un opérateur logique (et/ou)
    if IsOperateurLogique(Operateur) then
    begin
      // il doit y avoir au moins un opérateur d'inégalité avant
      OperateurValid := false;
      for j := 0 to (i - 1) //for 2
        do
      begin
        Operateurj := GetField('PVA_OPERATMATH' + inttostr(j));
        if IsOperateurInegalite(Operateurj) then
        begin
          OperateurValid := true;
          break;
        end;
      end; // for 2
      if not OperateurValid then
      begin
        LastError := 401;
        LastErrorMsg := 'Impossible de mettre un opérateur logique sans exprimer préalablement un opérateur d''inégalité';
        SetFocusControl('PVA_OPERATMATH' + inttostr(i));
        exit;
      end;
      // il doit y avoir au moins un opérateur d'inégalité après
      OperateurValid := false;
      for j := (i + 1) to (iMaxTpva_operatmath - 1) //for 3
        do
      begin
        Operateurj := GetField('PVA_OPERATMATH' + inttostr(j));
        if IsOperateurInegalite(Operateurj) then
        begin
          OperateurValid := true;
          break;
        end;
      end; // for 2
      if not OperateurValid then
      begin
        LastError := 402;
        LastErrorMsg := 'Impossible de mettre un opérateur logique sans exprimer ultérieurement un opérateur d''inégalité';
        SetFocusControl('PVA_OPERATMATH' + inttostr(i));
        exit;
      end;
      continue;
    end;
    if IsOperateurInegalite(Operateur) then
    begin
      // on controle que l'opérateur qui le précède immédiatement n'est pas un opérateur d'inégalité
      j := i - 1;
      if j >= 0 then
      begin
        Operateurj := GetField('PVA_OPERATMATH' + inttostr(j));
        if IsOperateurInegalite(Operateurj) then
        begin
          LastError := 403;
          LastErrorMsg := '2 opérateurs d''inégalité ne peuvent se succéder';
          SetFocusControl('PVA_OPERATMATH' + inttostr(i));
          exit;
        end;
      end;
    end;
    if Operateur = 'FIN' then
    begin
      if i = 0 then
      begin
        LastError := 404;
        LastErrorMsg := 'Choisir un autre opérateur à cette position';
        SetFocusControl('PVA_OPERATMATH' + inttostr(i));
        exit;
      end;
    end;
  end; // do begin du for 1

  //----2 : controle de validite des opérations sur date
  for i := 0 to (iMaxTpva_Base - 1) do //for 2
  begin
    MTypeBase := THDBValComboBox(getcontrol('PVA_TYPEBASE' + inttostr(i)));
{$IFNDEF EAGLCLIENT}
    MBase := THDbEdit(getcontrol('PVA_BASE' + inttostr(i)));
{$ELSE}
    MBase := THEdit(getcontrol('PVA_BASE' + inttostr(i)));
{$ENDIF}
    if ((MTypeBase <> nil) and (MBase <> nil)) then
    begin
      if ((MTypeBase.Value <> '23') or
        ((MTypeBase.Value = '03') and (not IsVariabledate(MBase.Text))))
        then continue; // type date
      // On est sur une date
      // Il doit y avoir une autre date avant ou après
      if i > 0 then
      begin
        MTypeBaseP := THDBValComboBox(getcontrol('PVA_TYPEBASE' + inttostr(i - 1)));
{$IFNDEF EAGLCLIENT}
        MBaseP := THDbEdit(getcontrol('PVA_BASE' + inttostr(i - 1)));
{$ELSE}
        MBaseP := THEdit(getcontrol('PVA_BASE' + inttostr(i - 1)));
{$ENDIF}
        if ((MTypeBaseP.Value = '23') or (MTypeBaseP.value = '04') or
          ((MTypeBaseP.Value = '03') and (IsVariabledate(MBaseP.Text))))
          then
          if IsOperateurDate('PVA_OPERATMATH' + inttostr(i - 1)) then continue
          else
          begin //6
            LastError := 405;
            LastErrorMsg := 'L''opérateur entre 2 dates doit être de type inégalité ou + ou -';
            SetFocusControl('PVA_OPERATMATH' + inttostr(i - 1));
            exit;
          end; //6
      end;
      if i <= (iMaxTpva_Base - 2) then
      begin //8
        MTypeBaseS := THDBValComboBox(getcontrol('PVA_TYPEBASE' + inttostr(i + 1)));
{$IFNDEF EAGLCLIENT}
        MBaseS := THDbEdit(getcontrol('PVA_BASE' + inttostr(i + 1)));
{$ELSE}
        MBaseS := THEdit(getcontrol('PVA_BASE' + inttostr(i + 1)));
{$ENDIF}
        if (MTypeBaseS <> nil) and (MBaseS <> nil) then
        begin
          if ((MTypeBaseS.Value = '23') or (MTypeBaseS.value = '04') or
            ((MTypeBaseS.Value = '03') and (IsVariabledate(MBaseS.Text))))
            then
          begin //5
            if IsOperateurDate('PVA_OPERATMATH' + inttostr(i)) then continue
            else
            begin //6
              LastError := 405;
              LastErrorMsg := 'L''opérateur entre 2 dates doit être de type inégalité ou + ou -';
              SetFocusControl('PVA_OPERATMATH' + inttostr(i));
              exit;
            end; //6
          end //5
          else
          begin //6
            if ((MTypeBaseS.Value <> '') or (MBaseS.Text <> '')) then // PT11
            begin
              LastError := 406;
              LastErrorMsg := 'Un type date doit être précédé ou suivi d''une autre date ou d''une valeur';
              SetFocusControl('PVA_TYPEBASE' + inttostr(i));
              exit;
            end;
          end; //6
        end;
      end; //8
    end;
  end; // for 2
  //----3 controle de saisie des zones
//  if not ControleSaisieBaseOk(RgErr, 'TEST') then exit;   // PT14
  if not ControleSaisieBaseOk(RgErr, 'TESTSI') then exit; // PT14

  // PT13 contrôle que la condition SI se termine par FIN
  for i := 0 to (ImaxTpva_operatmath - 1) do
  begin
    operateur := Getfield('PVA_OPERATMATH' + intTostr(i));
    if operateur = '' then break;
  end;
  if i > 1 then j := i - 1 else j := 0;
  if i > (ImaxTpva_operatmath - 1) then i := (ImaxTpva_operatmath - 1);
  operateur := Getfield('PVA_OPERATMATH' + intTostr(j));
  if operateur <> 'FIN' then begin
    lasterror := 407;
    lasterrormsg := 'La condition SI doit se terminer par FIN';
    setfocuscontrol('PVA_OPERATMATH' + inttostr(i));
    exit;
  end;
  // fin PT13

  if not ControleSaisieBaseOk(RgErr, 'TESTALORS') then exit;

  // PT13 contrôle que la condition ALORS se termine par FIN
  for i := 0 to (ImaxTpva_operattest - 1) do
  begin
    operateur := Getfield('PVA_OPERATRESTHEN' + intTostr(i));
    if operateur = '' then break;
  end;
  if i > 1 then j := i - 1 else j := 0;
  if i > (ImaxTpva_operattest - 1) then i := ImaxTpva_operattest - 1;
  operateur := Getfield('PVA_OPERATRESTHEN' + intTostr(j));
  if operateur <> 'FIN' then begin
    lasterror := 407;
    lasterrormsg := 'La condition ALORS doit se terminer par FIN';
    setfocuscontrol('PVA_OPERATRESTHEN' + inttostr(i));
    exit;
  end;
  // fin PT13

  if not ControleSaisieBaseOk(RgErr, 'TESTSINON') then exit;

   // PT13 contrôle que la condition SINON se termine par FIN
  for i := 0 to (ImaxTpva_operattest - 1) do
  begin
    operateur := Getfield('PVA_OPERATRESELSE' + intTostr(i));
    if operateur = '' then break;
  end;
  if i > 1 then j := i - 1 else j := 0;
  if i > (ImaxTpva_operattest - 1) then i := ImaxTpva_operattest - 1;
  operateur := Getfield('PVA_OPERATRESELSE' + intTostr(j));
  if operateur <> 'FIN' then begin
    lasterror := 407;
    lasterrormsg := 'La condition SINON doit se terminer par FIN';
    setfocuscontrol('PVA_OPERATRESELSE' + inttostr(i));
    exit;
  end;
  // fin PT13

end;

//------------------------------------------------------------------------------
// Controles de cohérence des zones de l'onglet Test
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.ControleZonesOngletCalcul;
var
  C: TControl;
  i: integer;
  err : boolean; // PT28
begin    
  err := false; // PT28
  for i := 0 to (iMaxEpva_Base - 1) do
  begin
    // si ligne non visible, inutile de la contrôler
    C := GetControl('PVA_TYPEBASE' + inttostr(i));
    if ((C = nil) or (C.visible = false)) then break;
    // Contrôle champ type de base
    if (GetField('PVA_TYPEBASE' + inttostr(i)) = '') then
    begin
      LastError := (300 + i + 1);
      LastErrorMsg := 'Vous devez renseigner un Type de base';
      SetFocusControl('PVA_TYPEBASE' + inttostr(i));
      err := true; // PT28
      break;
    end;
    // Contrôle champ base
    if (GetField('PVA_BASE' + inttostr(i)) = '') then
    begin
      LastError := (300 + i + 2);
      LastErrorMsg := 'Vous devez renseigner une base';
      SetFocusControl('PVA_BASE' + inttostr(i));
      err := true; // PT28
      break;
    end;
  end; // fin du for
  // deb PT28 : controle que la formule se termine par FIN
  if not err then
  begin
     i := i-1;
     if getfield('PVA_OPERATMATH' + inttostr(i)) <> 'FIN' then
     begin
        LastError := (300 + i + 3);
        LastErrorMsg := 'Le calcul doit se terminer par FIN';
        SetFocusControl('PVA_OPERATMATH' + inttostr(i));
     end;
  end;
  // FIN pt28
end;

//------------------------------------------------------------------------------
//- En cas de changement de nature de la variable, tous les champs à renseigner
//- sont initialisés
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.InitialiseChampsTable;
var
  i: integer;
begin
  for i := 0 to (iMaxEpva_Base - 1) do
  begin
    SetField('PVA_BASE' + inttostr(i), '');
    SetControlVisible('PVA_BASE' + inttostr(i), TRUE);
    SetField('PVA_TYPEBASE' + inttostr(i), '');
    SetControlVisible('PVA_TYPEBASE' + inttostr(i), TRUE);
    SetField('PVA_OPERATMATH' + inttostr(i), '');
    SetControlVisible('PVA_OPERATMATH' + inttostr(i), TRUE);
  end;

  for i := 0 to (iMaxTpva_resthen - 1) do
  begin
    SetField('PVA_RESTHEN' + inttostr(i), '');
    SetControlVisible('PVA_RESTHEN' + inttostr(i), TRUE);
    SetField('PVA_TYPERESTHEN' + inttostr(i), '');
    SetControlVisible('PVA_TYPERESTHEN' + inttostr(i), TRUE);
    SetField('PVA_RESELSE' + inttostr(i), '');
    SetControlVisible('PVA_RESELSE' + inttostr(i), TRUE);
    SetField('PVA_TYPERESELSE' + inttostr(i), '');
    SetControlVisible('PVA_TYPERESELSE' + inttostr(i), TRUE);
  end;

  for i := 0 to (iMaxTpva_Operattest - 1) do
  begin
    SetField('PVA_OPERATTEST' + inttostr(i), '');
    SetControlVisible('PVA_OPERATTEST' + inttostr(i), TRUE);
    SetField('PVA_OPERATRESTHEN' + inttostr(i), '');
    SetControlVisible('PVA_OPERATRESTHEN' + inttostr(i), TRUE);
    SetField('PVA_OPERATRESELSE' + inttostr(i), '');
    SetControlVisible('PVA_OPERATRESELSE' + inttostr(i), TRUE);
  end;
end;

//------------------------------------------------------------------------------
// fonction qui ouvre les lignes en fonction de la valeur de l'opérateur saisi
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.ControleOuvreLigne(MaCombo: THDBValComboBox; TypeZ: string; top: integer);
var
  i: integer; //max : indice maxi (Penser qu'indice commence à 0!)
  VarOnglet: PVarOnglet;
  M_TYPE: THDBValComboBox;
{$IFNDEF EAGLCLIENT}
  Medit: THDbEdit;
{$ELSE}
  Medit: THEdit;
{$ENDIF}
begin
  VarOnglet := PVarOnglet.create;
  //On donne les bons libelles de zones aux variables en fct du type
  AlimenteNomsZones(VarOnglet, Typez); 
  with VarOnglet do
  begin
    if MaCombo = nil then
    begin
      VarOnglet.Free;
      exit;
    end;
    M_TYPE := THDBValComboBox(GetControl(BDD_TYPEBASE + inttostr(Top)));
    if ((MaCombo.value <> '') and (MaCombo.value <> 'FIN')) then
    begin // rendre la ligne visible
      SetControlVisible(BDD_TYPEBASE + inttostr(Top), True);
      SetControlVisible(BDD_BASE + inttostr(Top), True);
      if ((M_TYPE.value = '04') or (M_type.value = '23')) then
      begin
        SetControlVisible(LBASE + inttostr(Top), False);
{$IFNDEF EAGLCLIENT}
        Medit := THDbEdit(getcontrol(BDD_BASE + inttostr(Top)));
{$ELSE}
        Medit := THEdit(getcontrol(BDD_BASE + inttostr(Top)));
{$ENDIF}
        if M_type.value = '23' then
          Medit.ElipsisButton := true else Medit.ElipsisButton := false;
      end
      else SetControlVisible(LBASE + inttostr(Top), True);
      SetControlVisible(BDD_OPERATMATH + inttostr(Top), True);
    end
    else // cacher toutes les lignes en dessous de la ligne courante,
      // et initialiser les zones correspondantes.
      if BoolInit then
      begin
        for i := Top to max do
        begin
          SetControlVisible(BDD_TYPEBASE + inttostr(i), False);
          SetControlVisible(LBASE + inttostr(i), False);
          SetControlVisible(BDD_OPERATMATH + inttostr(i), False);
          SetField(BDD_TYPEBASE + inttostr(i), '');
          SetField(BDD_OPERATMATH + inttostr(i), '');
          SetField(BDD_BASE + inttostr(i), '');
        end; // for do begin}
      end; // if Init
  end; //with do begin
  VarOnglet.Free;
end;

//------------------------------------------------------------------------------
// fonction qui alimente un ensemble de variable récurrentes aux fonctions de cette
// TOM (en fonction de l'endroit d'où on vient : typez)
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.AlimenteNomsZones(var VarOnglet: PVarOnglet; Typez: string);
begin
  with VarOnglet do
  begin
    if typeZ = 'CALC' then
    begin
      LBASE := 'TPVA_BASE';
      BDD_TYPEBASE := 'PVA_TYPEBASE';
      BDD_BASE := 'PVA_BASE';
      BDD_OPERATMATH := 'PVA_OPERATMATH';
      max := 9;
    end
    else
      if typeZ = 'TESTSI' then
      begin
        LBASE := 'TPVA_BASE';
        BDD_TYPEBASE := 'PVA_TYPEBASE';
        BDD_BASE := 'PVA_BASE';
        BDD_OPERATMATH := 'PVA_OPERATMATH';
        max := 9;
      end
      else
        if typeZ = 'TESTALORS' then
        begin
          LBASE := 'TPVA_RESTHEN';
          BDD_TYPEBASE := 'PVA_TYPERESTHEN';
          BDD_BASE := 'PVA_RESTHEN';
          BDD_OPERATMATH := 'PVA_OPERATRESTHEN';
          max := 3;
        end
        else
          if typeZ = 'TESTSINON' then
          begin
            LBASE := 'TPVA_RESELSE';
            BDD_TYPEBASE := 'PVA_TYPERESELSE';
            BDD_BASE := 'PVA_RESELSE';
            BDD_OPERATMATH := 'PVA_OPERATRESELSE';
            max := 3;
          end;
  end;
end;

//------------------------------------------------------------------------------
// Fonction qui vérifie si l'opérateur passé en paramètre est de type logique
//------------------------------------------------------------------------------

function TOM_VariablePaie.IsOperateurLogique(Operateur: string): boolean;
begin
  result := ((Operateur = 'AND') or (Operateur = 'OR'));
end;

//------------------------------------------------------------------------------
// Fonction qui vérifie si l'opérateur passé en paramètre est de type Inégalité
//------------------------------------------------------------------------------

function TOM_VariablePaie.IsOperateurInegalite(Operateur: string): boolean;
begin
  result := ((Operateur = '<') or (Operateur = '<=') or
    (Operateur = '>') or (Operateur = '>=') or
    (Operateur = '=') or (Operateur = '<>'));
end;

//------------------------------------------------------------------------------
// Fonction qui vérifie la saisie d'une base correspondant à un type de base saisi
//------------------------------------------------------------------------------

function TOM_VariablePaie.ControleSaisieBaseOk(RgErr: string; typez: string): boolean;
var
  C: TControl;
  i: integer;
  mTypeBase: string;
  mBase: string;
  VarOnglet: PVarOnglet;
begin
  VarOnglet := PVarOnglet.create;
  //On donne les bons libelles de zones aux variables, en fct du type passé
  AlimenteNomsZones(VarOnglet, Typez);

  with VarOnglet do
  begin
    for i := 0 to max do
    begin
      // si ligne non visible, inutile de la contrôler
      C := GetControl(BDD_TYPEBASE + inttostr(i));
      if ((C = nil) or (C.visible = false)) then break;
      // Contrôle champ type de base
      mTypeBase := GetField(BDD_TYPEBASE + inttostr(i));
      if (mTypeBase = '') then
      begin
        LastError := (strtoint(RgErr)) + i + 1;
        LastErrorMsg := 'Vous devez renseigner un Type de base';
        SetFocusControl(BDD_TYPEBASE + inttostr(i));
        VarOnglet.free;
        result := false;
        exit;
      end;

      // Contrôle champ base
      mBase := GetField(BDD_BASE + inttostr(i));
      if (mBase = '') then
      begin
        LastError := (strtoint(RgErr)) + i + 2;
        LastErrorMsg := 'Vous devez renseigner une base';
        SetFocusControl(BDD_BASE + inttostr(i));
        VarOnglet.free;
        result := false;
        exit;
      end;
    end; // end du for
  end; // with do begin

  VarOnglet.free;
  result := true;
end;

procedure TOM_VariablePaie.AfficheTypeBase(TypeBase, libbase, Base: string);
var
  MTYPEBASE: THDBValComboBox;
{$IFNDEF EAGLCLIENT}
  MBASEEdit: THDbEdit;
{$ELSE}
  MBASEEdit: THEdit;
{$ENDIF}
  tablette: string;
begin
  MTYPEBASE := THDBValComboBox(GetControl(TypeBase));
{$IFNDEF EAGLCLIENT}
  MBASEEdit := THDbEdit(GetControl(Base));
{$ELSE}
  MBASEEdit := THEdit(GetControl(Base));
{$ENDIF}
  if (MTYPEBASE <> nil) and (MBASEEdit <> nil) then
  begin
    // Type de base = 'valeur' ou 'aucun'
    if ((MTYPEBASE.Value = '04') or (Mtypebase.value = '23')) then
    begin
      if MBaseEdit.Datatype <> '' then
      begin
        MBaseEdit.Datatype := '';
        SetField(Base, '');
      end;
      if ((MTYPEBASE.Value = '23') and (MBaseedit <> nil)) then
      begin // DEB PT9
        MBaseEdit.ElipsisButton := true;
        MBAseEdit.OnElipsisClick := DateElipsisclick;
        MBAseEdit.Width := 100;
      end
      else
      begin
        MBAseEdit.Width := 50;
        MBaseEdit.ElipsisButton := false;
      end; // FIN PT9
    end // initialiation champ BDD
    else if MTYPEBASE.Value = '' then
    begin
      if (Base[length(Base)] = '0') then
      begin
        if GetField('PVA_NATUREVAR') <> 'CUM' then
          SetControlVisible(TypeBase, TRUE)
        else
          SetControlVisible(TypeBase, FALSE);
        SetControlVisible(Base, TRUE);
      end
      else
      begin
        SetControlVisible(TypeBase, False);
        SetControlVisible(Base, False);
      end;
      if (MBaseEdit.Datatype <> '') and (GetField('PVA_NATUREVAR') <> 'CUM') then
      begin
        MBaseEdit.Datatype := '';
        SetField(Base, '');
      end;
    end
    else
    begin // DEB PT9
      if    (MTYPEBASE.Value >= '500') and (Mtypebase.value <= '530') then
      begin
//        MBAseEdit.Width := 69;
        MBaseEdit.ElipsisButton := true;
        MBAseEdit.OnElipsisClick := nil;
      end else begin
        MBAseEdit.Width := 50;
        MBAseEdit.OnElipsisClick := nil;
      end;
    end; // FIN PT9
    // Ai je une tablette qui corresponde à ce type de base ?
    // si oui, je l'affecte au datatype de la base correspondante .
    Tablette := RechercheTablette(MTYPEBASE);
    if Tablette <> '' then
    begin
      ControleAncienneTablette(Base, tablette, FALSE);
      if GetField('PVA_NATUREVAR') <> 'CUM' then SetControlVisible(TypeBase, TRUE);
    end;
  end;
end;

function TOM_VariablePaie.IsVariabledate(code: string): boolean;
begin
  if ((code = '0015') or (code = '0016') or (code = '0017')) then result := true
  else result := false;
end;

function TOM_VariablePaie.IsOperateurDate(Operateur: string): boolean;
var
  Moperateur: THDBValComboBox;
begin
  result := false;
  MOperateur := THDBValComboBox(getcontrol(Operateur));
  if MOperateur <> nil then
  begin
    if ((Moperateur.value = '+') or (Moperateur.value = '-') or
      IsOperateurInegalite(Moperateur.value)) then
      result := true;
  end;
end;

//------------------------------------------------------------------------------
// Affectation, dans la variable Tablette de la tablette correspondant au type de base
//------------------------------------------------------------------------------

function RechercheNomTablette(TypeBase, NatureVariable: String; Presence : Boolean): string;
begin
  if      TypeBase = '02' then
    result := 'PGELEMENTNAT'
  else if TypeBase = '30' then
    result := 'PGZONEELTDYN'
  else if TypeBase = '31' then
    result := 'PGZONELIBRESAL'
  else if (TypeBase = '12') or (TypeBase = '13') or (TypeBase = '14') or (TypeBase = '25') or (TypeBase = '26') then
    result := 'PGCOTISATION'
  else if (TypeBase = '16') or (TypeBase = '17') or (TypeBase = '18') or (TypeBase = '19') then
    result := 'PGBASECOTISATION'
  else if (TypeBase = '05') or (TypeBase = '06') or (TypeBase = '07') or (TypeBase = '08') or (TypeBase = '09') or (TypeBase = '10') then
    result := 'PGREMUNERATION'
  else if TypeBase = '03' then
    if not Presence then
      result := 'PGVARIABLE'
    else
      result := 'PGVARIABLECEGSALVAL'
  else if TypeBase = '15' then
    result := 'PGCHAMPSALARIE'
  else if TypeBase = '20' then
    result := 'PGTABINTAGE'
  else if TypeBase = '21' then
    result := 'PGTABINTANC'
  else if TypeBase = '22' then
    result := 'PGTABINTDIV'
  else if TypeBase = '220' then
    result := 'PGTABINTDIW'
  else if TypeBase = '225' then
    result := 'PGTABLEDIMDSAL'
  else if TypeBase = '500' then
    result := 'PGCOMPTEURPRESJ'
  else if TypeBase = '501' then
    result := 'PGVARIABLEPREJ'
  else if TypeBase = '502' then
    result := 'PGCOMPTEURPRESH'
  else if TypeBase = '503' then
    result := 'PGVARIABLEPREH'
  else if TypeBase = '504' then
    result := 'PGCOMPTEURPRESM'
  else if TypeBase = '505' then
    result := 'PGVARIABLEPREM'
  else if (TypeBase = '506') or (TypeBase = '507') or (TypeBase = '508') then
    result := 'PGDROITJOURNEETYPE'
  else if (TypeBase = '510') or (TypeBase = '512') or (TypeBase = '513') or (TypeBase = '514') or
          (TypeBase = '515') or (TypeBase = '516') or (TypeBase = '517') or (TypeBase = '518') then
    result := 'PGMOTIFEVENEMENT'
  else if TypeBase = '520' then
    result := 'PGCOMPTEURPRESC'
  else if TypeBase = '521' then
    result := 'PGVARIABLEPREC'
  else if TypeBase = '530' then
    result := 'PGVARIABLEPRECEG'
  else
    if NatureVariable = 'CUM' then
      result := 'PGCUMULPAIE'
    else
      result := '';
end;



//------------------------------------------------------------------------------
// Affectation, dans la variable Tablette de la tablette correspondant au type de base
//------------------------------------------------------------------------------

function TOM_VariablePaie.RechercheTablette(Base: TControl): string;
var
  //Debut PT15
{$IFNDEF EAGLCLIENT}
  ControlBASE: THDbEdit;
{$ELSE}
  ControlBASE: THEdit;
{$ENDIF}
  TypeBase, ControlName: string;
   //Fin PT15
begin
  if Base is THDBValComboBox then
    TypeBase := (Base as THDBValComboBox).Value
  else
    TypeBase := GetControlText(Base.Name);
  ControlBASE := nil;
  ControlName := StringReplace(Base.Name, 'TYPE', '', [rfIgnoreCase]);
{$IFNDEF EAGLCLIENT}
  if GetControl(ControlName) is THDbEdit then
    ControlBASE := THDbEdit(GetControl(ControlName));
{$ELSE}
  if GetControl(ControlName) is THEdit then
    ControlBASE := THEdit(GetControl(ControlName));
{$ENDIF}
  if Assigned(ControlBASE) then ControlBASE.OnElipsisClick := nil; //PT15
  result := RechercheNomTablette( TYPEBASE, GetField('PVA_NATUREVAR'), (GblTypeVariable = 'PRE'));
  if Assigned(ControlBASE) then
  begin
    if TYPEBASE = '30' then
      ControlBASE.OnElipsisClick := OnElipsisClickBaseEltDyna
    else if TYPEBASE = '31' then
      ControlBASE.OnElipsisClick := OnElipsisClickBaseZoneLibreSal
    else if TYPEBASE = '225' then
    begin
      RechercheEtabConv;
      ControlBASE.OnElipsisClick := OnElipsisClickBase;
    end else if (TYPEBASE = '500') or (TYPEBASE = '502') or (TYPEBASE = '504') or (TYPEBASE = '520') then
      ControlBASE.OnElipsisClick := OnElipsisClickBase;
  end;
end;

    //Debut PT15 Tables Dynamiques

procedure TOM_VariablePaie.RechercheEtabConv;
var
  Qry: TQuery;
  Requete: string;
begin
if not HadSearchEtabConv then
   begin
   Requete:= 'SELECT ET_ETABLISSEMENT, ET_NODOSSIER, ETB_CONVENTION,'+
             ' ETB_CONVENTION1, ETB_CONVENTION2'+
             ' FROM ETABLISS'+
             ' LEFT JOIN ETABCOMPL ON'+
             ' ET_ETABLISSEMENT = ETB_ETABLISSEMENT WHERE'+
             ' (ET_NODOSSIER="000000" OR'+
             ' ET_NODOSSIER="'+V_PGI.NoDossier+'")';
   Qry:= opensql (Requete, true);
   while not Qry.EOF do
         begin
         if pos (Qry.findfield ('ET_ETABLISSEMENT').asstring, ListeEtab)=0 then
            ListeEtab:= ListeEtab+'"'+Qry.findfield ('ET_ETABLISSEMENT').asstring+'", ';
         if pos (Qry.findfield ('ETB_CONVENTION').asstring, ListeConv)=0 then
            ListeConv:= ListeConv+'"'+Qry.findfield ('ETB_CONVENTION').asstring+'", ';
         if pos (Qry.findfield ('ETB_CONVENTION1').asstring, ListeConv)=0 then
            ListeConv:= ListeConv+'"'+Qry.findfield ('ETB_CONVENTION1').asstring+'", ';
         if pos (Qry.findfield ('ETB_CONVENTION2').asstring, ListeConv)=0 then
            ListeConv:= ListeConv+'"'+Qry.findfield ('ETB_CONVENTION2').asstring+'", ';
         Qry.next;
         end;
   ferme(Qry);
   HadSearchEtabConv := True;
   end;
end;
//Fin PT15




//------------------------------------------------------------------------------
// affecte la bonne table au datatype de la base
//------------------------------------------------------------------------------

procedure TOM_VariablePaie.ControleAncienneTablette(Champ, Tablette: string; modif: boolean);
var
{$IFNDEF EAGLCLIENT}
  LaCombo: THDBValComboBox;
  LEdit: THDBEdit;
{$ELSE}
  LaCombo: THValComboBox;
  LEdit: THEdit;
{$ENDIF}
  Affect : Boolean;
begin
// DEB PT18
  if Champ = '' then exit;
{$IFNDEF EAGLCLIENT}
  if ((GetControl(Champ) is THDBValComboBox)) then LaCombo := THDBValComboBox(GetControl(Champ));
{$ELSE}
  if ((GetControl(Champ) is THValComboBox)) then LaCombo := THValComboBox(GetControl(Champ));
{$ENDIF}
  if LaCombo = nil then
  begin
{$IFNDEF EAGLCLIENT}
    if ((GetControl(Champ) is THDBEdit)) then LEdit := THDBEdit(GetControl(Champ));
{$ELSE}
    if ((GetControl(Champ) is THEdit)) then LEdit := THEdit(GetControl(Champ));
{$ENDIF}
  end;
  if (LaCombo <> nil) or (LEdit <> nil) then
  begin
    if (LaCombo <> nil) then
    begin
      if LaCombo.Datatype <> Tablette then Affect := TRUE
      else Affect := FALSE;
    end
    else
    begin
      if LEdit.Datatype <> Tablette then Affect := TRUE
      else Affect := FALSE;
    end;
    if Affect then
    begin
      if Modif then SetField(Champ, '');
      SetControlProperty(Champ, 'datatype', tablette);
    end;
  end;
// FIN PT18
end;

procedure TOM_VariablePaie.AfficheLibelle(Base, libBase: string);
var
  Libelle: thlabel;
  Qry: TQuery;
  Requete, MBaseDatatype, perio: string;
{$IFNDEF EAGLCLIENT}
  MBase: THDBEdit;
{$ELSE}
  MBase: THEdit;
{$ENDIF}
begin
  Libelle := THLabel(GetControl(libBase));
{$IFNDEF EAGLCLIENT}
  MBase := THDBEdit(GetControl(Base));
{$ELSE}
  MBase := THEdit(GetControl(Base));
{$ENDIF}
  if ((MBase <> nil) and (MBase.text <> '')) then
  begin
    if (Libelle <> nil) then
    begin
      { Début PT22 }
      MBaseDatatype := MBase.datatype;
      if (MBaseDatatype = 'PGTABLEDIMDSAL') or (pos('PGCOMPTEURPRES', MBaseDatatype) > 0) then
      begin
        if MBaseDatatype = 'PGTABLEDIMDSAL' then
          Requete := 'Select PTE_LIBELLE from TABLEDIMENT TD where ##PTE_PREDEFINI## '
                   + 'AND PTE_CODTABL = "'+GetField(Base)+ '" '
                   + 'AND PTE_DTVALID = (SELECT MAX(Z.PTE_DTVALID) '
                   + '                        FROM TABLEDIMENT Z '
//PT33                   + '                       WHERE Z.PTE_CODTABL=TD.PTE_CODTABL '
                                         + ' WHERE ##Z.PTE_PREDEFINI## ' //PT6
                                         + ' AND TD.PTE_CODTABL = Z.PTE_CODTABL '
                                         + ' AND TD.PTE_PREDEFINI = Z.PTE_PREDEFINI '
                                         + ' AND TD.PTE_NODOSSIER = Z.PTE_NODOSSIER '
                                         + ' AND TD.PTE_NIVSAIS = Z.PTE_NIVSAIS '
                                         + ' AND TD.PTE_DTVALID<="'+UsDateTime(Date())+'")';
//PT33                   + '                         AND TD.PTE_DTVALID<="' + UsDateTime(Date()) + '")';
          if Getfield('PVA_PREDEFINI') = 'DOS' then  //PT31
          begin
            Requete := Requete
              + ' AND (   PTE_NIVSAIS = "GEN" '
              + '      or PTE_PREDEFINI = "DOS" '
              + '      or (  (PTE_NIVSAIS = "CON" and PTE_VALNIV in (' + ListeConv + '"000")) '
              + '         or (PTE_NIVSAIS = "ETB" and PTE_VALNIV in (' + ListeEtab + '"...")) '
              + '         ) '
              + '      ) ';
          end;
        if pos('PGCOMPTEURPRES', MBaseDatatype) > 0 then
        begin
          perio := '';
          if MBaseDatatype[Length(MBaseDatatype)] = 'C' then perio := periodiciteFinDeCycle;
          if MBaseDatatype[Length(MBaseDatatype)] = 'J' then perio := periodiciteJournaliere;
          if MBaseDatatype[Length(MBaseDatatype)] = 'H' then perio := periodiciteHebdomadaire;
          if MBaseDatatype[Length(MBaseDatatype)] = 'M' then perio := periodiciteMensuelle;
          Requete := 'Select PYR_LIBELLE from COMPTEURPRESENCE TD '
                   + ' where ##PYR_PREDEFINI## AND PYR_COMPTEURPRES = "'+MBase.text+'" '
                   + ' AND PYR_PERIODICITEPRE = "'+perio+'" '
                   + ' AND PYR_DATEVALIDITE = (SELECT MAX(Z.PYR_DATEVALIDITE) '
                   + '                        FROM COMPTEURPRESENCE Z '
                   + '                       WHERE Z.PYR_COMPTEURPRES=TD.PYR_COMPTEURPRES '
                   + '                         AND TD.PYR_DATEVALIDITE <= "' + UsDateTime(Date()) + '")'
        end;
        Qry:= opensql (Requete, true);
        if not Qry.EOF then
          Libelle.Caption := Qry.Fields[0].AsString
        else
          Libelle.Caption := '';
        ferme(Qry);
      end else  { Fin PT22 } 
        if MBase.datatype <> '' then
          Libelle.Caption := RechDom(MBase.datatype, GetField(Base), FALSE)
        else
          Libelle.Caption := GetField(Base);
      SetControlVisible(libBase, TRUE);
    end;
    SetControlVisible(Base, TRUE);
  end
  else
  begin
    if (Libelle <> nil) then
    begin
      if ((MBase <> nil) and (MBase.Text = '')) then Libelle.Caption := '';
      SetControlVisible(libBase, FALSE);
      if (Base[length(Base)] = '0') then SetControlVisible(Base, TRUE);
    end;
  end;
end;

procedure TOM_VariablePaie.OnDeleteRecord;
begin
  inherited;
  // PT8  ChargementTablette(TFFiche(Ecran).TableName, '');
  //PT16
  Trace := TStringList.Create;
  Trace.Add('SUPPRESSION VARIABLE ' + GetField('PVA_VARIABLE') + ' ' +
    GetField('PVA_LIBELLE'));
  CreeJnalEvt('003', '083', 'OK', nil, nil, Trace);
  FreeAndNil(Trace);
  //FIN PT16
  If (GblTypeVariable = 'PRE') then
    CompteursARecalculer(Date); //PT24
end;

procedure TOM_VariablePaie.OnAfterUpdateRecord;
var
  even: boolean;
begin
  inherited;
//PT16
  Trace := TStringList.Create;
  even := IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran),LeStatut); //PT37  //PT30
  FreeAndNil(Trace);
//FIN PT16
  if OnFerme then Ecran.Close;
  SetControlEnabled('BDUPLIQUER', not ((GetField('PVA_NATUREVAR') = 'VAL') and (GetField('PVA_PREDEFINI') = 'CEG'))); { PT5-1 }

end;

{ DEB PT5 Méthode de duplication de la fiche }

procedure TOM_VariablePaie.DupliquerVariable(Sender: TObject);
var
  NoDossier: string;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  st, ancvalcod: string;
begin
  TFFiche(Ecran).BValider.Click;
  mode := 'DUPLICATION';
  ancvalcod := GetField('PVA_VARIABLE');
  AglLanceFiche('PAY', 'CODE', '', '', 'VAR;' + GetField('PVA_VARIABLE') + '; ;4');
      // DEB PT17
  if (PGCodePredefini = 'DOS') and (not DOS) then
  begin
    PgiInfo('Vous n''êtes pas autorisé à créer une variable de type dossier.', Ecran.Caption);
    exit;
  end;
  // FIN PT17
  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PVA_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PVA_NODOSSIER';
    if PGCodePredefini = 'DOS' then Valeur[2] := PgRendNoDossier()
    else Valeur[2] := '000000';
    Champ[3] := 'PVA_VARIABLE';
    Valeur[3] := PGCodeDupliquer;
    if RechEnrAssocier('VARIABLEPAIE', Champ, Valeur) = False then //Test si code existe ou non
    begin
      DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PVA_VARIABLE', PGCodeDupliquer);
      SetField('PVA_PREDEFINI', PGCodePredefini);
      SetField('PVA_TYPEVARIABLE', GblTypeVariable); //PT15
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PVA_NODOSSIER', NoDossier);
//PT19
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create;
      st := 'Duplication de la rubrique '+AncValCod;
      Trace.add (st);
      st := 'Création de la rubrique '+ GetField('PVA_VARIABLE');
      Trace.add (st);
      EnDupl := 'OUI';
      IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran)); //PT44
      FreeAndNil(Trace);
      EnDupl := 'NON';
//Fin PT19
{$IFNDEF EAGLCLIENT}
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
{$ENDIF}
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PVA_PREDEFINI', False);
      SetControlEnabled('PVA_VARIABLE', False);
      SetControlEnabled('BDUPLIQUER', not ((GetField('PVA_NATUREVAR') = 'VAL') and (GetField('PVA_PREDEFINI') = 'CEG'))); { PT5-1 }
      // PT8      ChargementTablette(TFFiche(Ecran).TableName, ''); //Recharge les tablettes
    end
    else
      PgiBox('La duplication est impossible, la variable de paie existe déjà.', Ecran.Caption);
  end;
  mode := '';
end;
{ FIN PT5 }

  //Debut PT15

procedure TOM_VariablePaie.OnElipsisClickBase(Sender: TObject);
var
  Tablette, perio : String;
  SQLWhere : String; // PT31
begin
  if GblTypeVariable <> 'PRE' then
  begin
    SQLWhere := ' ##PTE_PREDEFINI## '
      + ' AND PTE_DTVALID = (SELECT MAX(Z.PTE_DTVALID) '
      + '                        FROM TABLEDIMENT Z '
//PT33+ '                       WHERE Z.PTE_CODTABL=TD.PTE_CODTABL '
                            + ' WHERE ##Z.PTE_PREDEFINI## ' //PT6
                            + ' AND TD.PTE_CODTABL = Z.PTE_CODTABL '
                            + ' AND TD.PTE_PREDEFINI = Z.PTE_PREDEFINI '
                            + ' AND TD.PTE_NODOSSIER = Z.PTE_NODOSSIER '
                            + ' AND TD.PTE_NIVSAIS = Z.PTE_NIVSAIS '
                            + ' AND TD.PTE_DTVALID<="'+UsDateTime(Date())+'")';
//PT33+ '                       AND TD.PTE_DTVALID<="' + UsDateTime(Date()) + '")';
    if Getfield('PVA_PREDEFINI') = 'DOS' then  //PT31
    begin
      SQLWhere := SQLWhere
        + ' AND (   PTE_NIVSAIS = "GEN" '
        + '      or PTE_PREDEFINI = "DOS" '
        + '      or (  (PTE_NIVSAIS = "CON" and PTE_VALNIV in (' + ListeConv + '"000")) '
        + '         or (PTE_NIVSAIS = "ETB" and PTE_VALNIV in (' + ListeEtab + '"...")) '
        + '         ) '
        + '      ) ';
    end;
    LookupList(Sender as TControl
      , 'Tables dynamiques'
      , 'TABLEDIMENT TD'
      , 'DISTINCT PTE_CODTABL'
      , 'PTE_LIBELLE' // , TD.PTE_DTVALID, TD.PTE_NIVSAIS, TD.PTE_VALNIV
      , SQLWhere
      , 'TD.PTE_CODTABL', True, -1);
  end;
  if GblTypeVariable = 'PRE' then
  begin
{$IFDEF EAGLCLIENT}
    if (Sender is THEdit) then
      Tablette := (Sender as THEdit).DataType;
{$ELSE}
    if (Sender is THDBEdit) then
      Tablette := (Sender as THDBEdit).DataType;
{$ENDIF}
    if length(Tablette) > 0 then
    begin
      if Tablette[Length(Tablette)] = 'C' then perio := periodiciteFinDeCycle;
      if Tablette[Length(Tablette)] = 'J' then perio := periodiciteJournaliere;
      if Tablette[Length(Tablette)] = 'H' then perio := periodiciteHebdomadaire;
      if Tablette[Length(Tablette)] = 'M' then perio := periodiciteMensuelle;
      if pos('PGCOMPTEURPRES', Tablette) > 0 then
        LookupList(Sender as TControl
          , 'Compteurs de présence'
          , 'COMPTEURPRESENCE TD'
          , 'PYR_COMPTEURPRES'
          , 'PYR_LIBELLE'
          , ' ##PYR_PREDEFINI## AND PYR_PERIODICITEPRE = "'+perio+'" '
          + ' AND PYR_DATEVALIDITE = (SELECT MAX(Z.PYR_DATEVALIDITE) '
          + '                        FROM COMPTEURPRESENCE Z '
          + '                       WHERE Z.PYR_COMPTEURPRES=TD.PYR_COMPTEURPRES '
          + '                         AND TD.PYR_DATEVALIDITE <= "' + UsDateTime(Date()) + '")'
          , 'PYR_COMPTEURPRES,PYR_DATEVALIDITE DESC', True, -1);
    end;
  end;
end;
  //Fin PT15


//PT29
procedure TOM_VariablePaie.OnElipsisClickBaseEltDyna(Sender: TObject); //PT29
begin
  LookupList(Sender as TControl
    , 'Eléments dynamiques'
    , 'paramsalarie'
    , 'ppp_pginfosmodif'
    , 'ppp_libelle'
    , 'ppp_pgtypeinfols = "ZLS" and ppp_pgtypedonne = "B" '
    , 'ppp_libelle', True, -1);
end;

//PT29
procedure TOM_VariablePaie.OnElipsisClickBaseZoneLibreSal(Sender: TObject); //PT29
begin
  LookupList(Sender as TControl
    , 'Zone libre salarié'
    , 'paramsoc'
    , 'soc_nom'
    , 'SOC_DATA'
    , '    (soc_nom = "SO_PGLIBCOCHE4" and (select SOC_DATA from paramsoc where SOC_NOM = "SO_PGNBCOCHE") >= 4 ) '
    + ' or (soc_nom = "SO_PGLIBCOCHE3" and (select SOC_DATA from paramsoc where SOC_NOM = "SO_PGNBCOCHE") >= 3 ) '
    + ' or (soc_nom = "SO_PGLIBCOCHE2" and (select SOC_DATA from paramsoc where SOC_NOM = "SO_PGNBCOCHE") >= 2 ) '
    + ' or (soc_nom = "SO_PGLIBCOCHE1" and (select SOC_DATA from paramsoc where SOC_NOM = "SO_PGNBCOCHE") >= 1 ) '
    , 'soc_nom', True, -1);
end;

//PT24 Gestion des variables de présence
procedure TOM_VariablePaie.InitModePresence;
var
  i : Integer;
begin
  if GblTypeVariable <> 'PRE' then exit;
  SetControlVisible('TPVA_VARPERIODICITE', True);
  SetControlVisible('PVA_VARPERIODICITE', True);
  setcontrolVisible('BCOMPTEURPRES', True); // PT27
{$IFNDEF EAGLCLIENT}
  if GetControl('PVA_VARPERIODICITE') is THDBValComboBox then
    (GetControl('PVA_VARPERIODICITE') as THDBValComboBox).OnChange := OnChangePeriodicite;
{$ELSE}
  if GetControl('PVA_VARPERIODICITE') is THValComboBox then
    (GetControl('PVA_VARPERIODICITE') as THValComboBox).OnChange := OnChangePeriodicite;
{$ENDIF}
  SetControlProperty('PVA_VARIABLE', 'EditMask', 'P000;1;_');
  SetControlProperty('PVA_THEMEVAR', 'DataType', 'PGTHEMECOMPTEURPRES');
  if (LaNature = 'CAL') then
  begin
    for i := 0 to (iMaxEpva_Base - 1) do
      SetControlProperty('PVA_TYPEBASE' + inttostr(i), 'Datatype', 'PGTYPECHAMPVARPRE');
  end;
  if (LaNature = 'CUP') then
  begin
    SetControlProperty('PVA_TYPEBASE0', 'Datatype', 'PGTYPECHAMPVARPRE');
  end;
  if (LaNature = 'TES') then
  begin
    for i := 0 to (iMaxTpva_Base - 1) do
      SetControlProperty('PVA_TYPEBASE' + inttostr(i), 'Datatype', 'PGTYPECHAMPVARPRE');
    for i := 0 to (iMaxTpva_resthen - 1) do
      SetControlProperty('PVA_TYPERESTHEN' + inttostr(i), 'Datatype', 'PGTYPECHAMPVARPRE');
    for i := 0 to (iMaxTPVA_RESELSE - 1) do
      SetControlProperty('PVA_TYPERESELSE' + inttostr(i), 'Datatype', 'PGTYPECHAMPVARPRE');
  end;
  if (LaNature = 'VAL') then
  begin
    SetControlProperty('PVA_TYPEBASE0', 'Datatype', 'PGTYPECHAMPVARPRE');
  end;

end;

//PT24 Gestion des variables de présence
procedure TOM_VariablePaie.OnChangePeriodicite(Sender: TObject);
var
  perio, stperio : String;
  i : Integer;
  procedure SetPlusPropertyPeriodicite(ChampsBase : String);
  begin
    SetControlProperty(ChampsBase, 'Plus',
      ' AND (CO_LIBRE LIKE "%'+stperio+LaNature+'%" OR CO_LIBRE LIKE "%'+stperio+'CCTV%")');
  end;
begin
  perio := GetControlText('PVA_VARPERIODICITE');
  if perio = periodiciteMensuelle    then stperio := 'm';
  if perio = periodiciteHebdomadaire then stperio := 'h';
  if perio = periodiciteJournaliere  then stperio := 'j';
  if perio = periodiciteFinDeCycle   then stperio := 'c';
  if (LaNature = 'CAL') then
  begin
    for i := 0 to (iMaxEpva_Base - 1) do
      SetPlusPropertyPeriodicite('PVA_TYPEBASE' + inttostr(i));
  end;
  if (LaNature = 'CUP') then
    SetPlusPropertyPeriodicite('PVA_TYPEBASE0');
  if (LaNature = 'TES') then
  begin
    for i := 0 to (iMaxTpva_Base - 1) do
      SetPlusPropertyPeriodicite('PVA_TYPEBASE' + inttostr(i));
    for i := 0 to (iMaxTpva_resthen - 1) do
      SetPlusPropertyPeriodicite('PVA_TYPERESTHEN' + inttostr(i));
    for i := 0 to (iMaxTPVA_RESELSE - 1) do
      SetPlusPropertyPeriodicite('PVA_TYPERESELSE' + inttostr(i));
  end;
  if (LaNature = 'VAL') then
    SetPlusPropertyPeriodicite('PVA_TYPEBASE0');

  //PT25 gestion de la période de calcul "Semaines précédentes du mois"
  if perio = '001' then
    SetControlProperty('PVA_PERIODECALCUL', 'Plus', ' AND CO_CODE <> "009" AND CO_CODE <> "011" AND CO_CODE <> "013" ')
  else if perio = '002' then
    SetControlProperty('PVA_PERIODECALCUL', 'Plus', ' AND CO_CODE <> "009" AND CO_CODE <> "011" AND CO_CODE <> "013" ')
  else if perio = '003' then
    SetControlProperty('PVA_PERIODECALCUL', 'Plus', ' AND CO_CODE <> "008" ')
  else
    SetControlProperty('PVA_PERIODECALCUL', 'Plus', ' AND CO_CODE <> "013" ');

end;

// PT27 Accès au compteur associé à cette variable
procedure TOM_VariablePaie.CompteurAssocie(Sender: TObject);
var
Q : Tquery;
varpres, paramcompteur: string;

Begin

  varpres:= getfield('PVA_VARIABLE');
  Q := opensql('SELECT PYR_PREDEFINI, PYR_NODOSSIER, PYR_COMPTEURPRES, PYR_DATEVALIDITE FROM COMPTEURPRESENCE '+
  ' WHERE PYR_VARIABLEPRES = "'+varpres+'"', true);
  if not Q.EOF then
  begin
    paramcompteur := getfield('PVA_VARPERIODICITE');
    AgllanceFiche('PAY', 'COMPTEURPRES', '', Q.Findfield('PYR_COMPTEURPRES').asstring + ';' + Q.Findfield('PYR_DATEVALIDITE').asstring + ';' +
    Q.findfield('PYR_PREDEFINI').asstring + ';' +  Q.findfield('PYR_NODOSSIER').asstring, 'ACTION=MODIFICATION' + ';' + paramcompteur);
  end
  else
  begin
    paramcompteur := getfield('PVA_PREDEFINI') + ';' + getfield('PVA_LIBELLE') + ';' + getfield('PVA_THEMEVAR') + ';' +
    getfield('PVA_VARPERIODICITE') + ';' + getfield('PVA_VARIABLE');
    AglLanceFiche('PAY', 'COMPTEURPRES', '','','ACTION=CREATION' + ';' + paramcompteur);
  end;

  ferme(Q);
end;

initialization
  registerclasses([TOM_VariablePaie]);
end.

