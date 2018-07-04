{***********UNITE*************************************************
Auteur  ...... : Paie PGI
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : unit de saisie des cotisations et des bases de cotisations
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 19/09/2001 VG V547 Le groupebox DADS doit être enabled pour une base
                           mais pas pour une cotisation
PT2   : 15/10/2001 VG V562 Gestion du enable/disable au niveau de la checkbox et
                           non au niveau du groupebox
PT3   : 14/11/2001 SB V562 Duplication affectation des predefinis et no dossier
PT4   : 14/11/2001 SB V562 Test si enr existe ds PROFILRUB
                           Ajout des champs PREDEFINI Et NODOSSIER
PT5   : 28/11/2001 VG V562 Gestion du enable/disable des nouveaux champs
                           Exonération
PT6   : 29/11/2001 SB V563 On force la validation avant la Duplication
PT7   : 30/11/2001 SB V563 Dysfonctionnement test code existant sur Duplication
PT8   : 13/12/2001 SB V570 Fiche de bug n° 279
                           Test code existant ne test pas bon numéro de dossier
PT9   : 04/04/2002 Ph V571 On saisit un element variable en base,taux et
                           forfaits
PT10  : 05/04/2002 Ph V571 Rubrique de cotisation en forfait, controle sur les
                           taux et non les forfaits
PT11  : 19/04/2002 SB V571 Fiche de bug n°369 Duplication d'une rub. Dossier en
                           Standard : L'affectation des prédéfini est erronée..
PT12  : 16/05/2002 JL V582 Fiche de bug n° 10116 suppression du message pour les
                           taux si élément variable
PT13  : 18/02/2003 SB V595 FQ 10531 Anomalie sur controle existence rubrique
                           dans table commune
PT14  : 19/02/2003 SB V595 FQ 10486 Affichage des zones selon COT ou BAS
PT15  : 04/06/2003 PH V421 FQ 10689 gestion du champ presence du salarié en fin
                           de mois pour le calcul de la cotisation
// **** Refonte accès V_PGI_env ***** PgRendNoDossier () remplacé par PgRendNoDossier() *****
PT16  : 19/06/2003 SB V595 Force validation pour rechargement contexte paie
PT17-1: 17/07/2003 SB V_42 FQ 10649 Duplication des éléments associés
PT17-2: 17/07/2003 SB V_42 FQ 10649 Intégration d'un lien vers l'affectation
                           DUCS
PT17-3: 17/07/2003 SB V_42 FQ 10649 Suppression des éléments associés
PT18-1: 16/09/2003 SB V_42 FQ 10806 Ajout contrôle message info
PT18-2: 16/09/2003 SB V_42 Déactivation de l'affichage du premier onglet en
                           validation
PT18-3: 16/09/2003 SB V_42 FQ 10820 Message duplication effectuée
PT19  : 14/10/2003    V_42 Portage CWAS : Impression fiche
PT20  : 12/01/2004 PH V_50 FQ 11029 Contrôle des bornes de validité
PT21  : 23/02/2004 SB V_50 Mise à jour date modif en duplication
PT22  : 08/06/2004 VG V_50 Ajout champ Epargne-Retraite pour la DADS-U
PT23  : 02/09/2004 PH V_50 FQ 11574 Faute orthographe
PT24  : 16/09/2004 PH V_50 FQ 11510 Cohérence champs de validité Du et AU
PT25  : 04/10/2004 PH V_50 Suppression rechargement tablette
PT26  : 11/05/2005 PH V_60 FQ 12274 Cohérence champs de validité Du et AU
PT27  : 24/05/2005 SB V_60 Rplct LastErrorMsg par pgibox.
PT28  : 26/05/2005 SB V_60 FQ 12307 Modification de l'ordre de présentation des
                           rubriques
PT29  : 10/06/2005 SB V_60 Ergonomie CWAS
PT30  : 21/06/2005 PH V_60 FQ 12297 Controle des champs onglet Etats en fonction
                           de la nature de la rubrique
PT31  : 12/07/2005 SB V_60 FQ 12308 Modification de l'ordre de présentation des
                           bulletins
PT32  : 02/08/2005 PH V_60 FQ Ergonomie
PT33  : 08/08/2005 PH V_60 FQ 12488 Blocage de validation si anomalie détectée
PT30-1: 09/08/2005 PH V_60 FQ 12297 Suite tests idem mais avec certains champs
PT34  : 11/08/2005 PH V_60 FQ 12245 CWAS Accès aux champ PCT_TYPEBASE_ et
                           PCT_BASECOTISATION_ ne fonctionne pas (AGL)
PT35  : 11/08/2005 PH V_60 FQ 12502 Duplication de rubrique aveccumul alpha
                           associé
PT36  : 18/08/2005 SB V_65 FQ 11969 Mise à jour du libellé de la table PROFILRUB
PT37  : 01/09/2005 SB V_60 FQ 12488 Refonte contrôles
PT38  : 12/09/2005 PH V_60 Affichage memo liste des champs personnalisés
PT39  : 20/10/2005 SB V_65 FQ 12628 Dysfonctionnement CWAS, sur base de
                           cotisation
PT40  : 09/02/2006 EP V_65 FQ 12781 Utilisateur non réviseur : interdire la
                           personnalisation
PT41  :	24/04/2006 SB V_65 FQ 12613 Refonte contrôle ondeleterecord pour test
                           parametrage multidossier
PT42  : 25/04/2006 SB V_65 FQ 12917 Erreur SQL : Refonte execute sql sur libellé
                           suite FQ 11969
PT43  : 12/10/2006 PH V_70 FQ 12274 Suite affichage pour la base de cotisation
PT44  : 11/12/2006 GGS V_80 FQ 12694 Journal d'événements
PT45  : 16/04/2007 VG V_72 Ajout bouton gestion spécifique
PT46  : 07/05/2007 PH V_72 Concept Paie
PT47  : 28/06/2007 MF V_72 FQ 14488 accès à la fiche Affectation Ducs
PT48  : 06/09/2007 VG V_80 Gestion spécifique grisée en prédéfini dossier
                           FQ N°14601
PT49  : 06/11/2007 FC V_80 Correction bug journal des évènement quand création (lib Modification au lieu de Création)
PT52  : 20/08/2008 JPP V_80 FQ15450 On empeche la duplication de la personnalisation lors de la duplication d'une rubrique CEGID en dossier
PT53  : 30/09/2008 JS FQ n°15702 le type des champs Taux salarial, Taux patronal ou Forfait salarial, Forfait patronal ou Minimum salarial, Minimum patroanl ou Maximum salarial, Maximum patronal est renseigné
}
unit UTOMCotisation;

interface
uses StdCtrls,
  Controls,
  Classes,
  forms,
  sysutils,
  ComCtrls,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  HDB,
  DBCtrls,
  Fiche,
  FE_Main,
  EdtREtat,
{$ELSE}
  eFiche,
  MaineAgl,
  eTablette,
  UtilEAgl,
{$ENDIF}
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTOB,
  HTB97,
  Ed_tools,
  P5Def,
  PAIETOM; //PT44


type
  TOM_Cotisation = class(PGTOM) //PT44 class TOM devien PGTOM
    procedure OnArgument(stArgument: string); override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnClose; override;
  private
    LectureSeule, CEG, STD, DOS : boolean;
    ExisteCod, OnFerme: Boolean;
    mode, DerniereCreate: string;
    NomChamp: array[1..4] of Hstring;
    ValChamp: array[1..4] of variant;
    OrdreEtat: Integer; { PT28 }
    Trace: TStringList; //PT44
    LeStatut:TDataSetState; //PT49
    procedure DupliquerCotisation(Sender: TObject);
    procedure AffecteOnglet();
    procedure AffecteLookupChamp(Champ, TypeChamp, libelle: string);
    function ControlChampValeur(nature: string): Boolean; // PT33
    procedure CBAccesFicheClick(Sender: TObject);
    procedure ImprimerClick(Sender: TObject); //PT19
    procedure PgUpdateDateAgl(Rub: string); //PT21
    procedure GestionSpec_OnClick(Sender: TObject);
    procedure AffectDucs_OnClick(Sender: TObject); // PT47

  end;


implementation
uses Pgoutils, PgOutils2;
//============================================================================================//
//                                 PROCEDURE On Delete Record
//============================================================================================//
{ PT41 Refonte de la proc. }

procedure TOM_Cotisation.OnDeleteRecord;
begin
  inherited;

  NomChamp[3] := '';
  ValChamp[3] := '';
  NomChamp[4] := '';
  ValChamp[4] := '';

//Recherche sur l'historique de paie
  NomChamp[1] := 'PHB_NATURERUB';
  ValChamp[1] := GetField('PCT_NATURERUB');
  NomChamp[2] := 'PHB_RUBRIQUE';
  ValChamp[2] := GetField('PCT_RUBRIQUE');
  ExisteCod := RechEnrAssocier('HISTOBULLETIN', NomChamp, ValChamp);
  if (ExisteCod = TRUE) then
  begin
    LastError := 1;
    PgiBox('Suppression impossible. Des bulletins de paie utilisent cette' +
      ' rubrique.', Ecran.caption);
    Exit;
  end;

//Recherche sur les profils
  NomChamp[1] := 'PPM_NATURERUB';
  NomChamp[2] := '##PPM_PREDEFINI## PPM_RUBRIQUE';
  ExisteCod := RechEnrAssocier('PROFILRUB', NomChamp, Valchamp);
  if (ExisteCod = TRUE) then
  begin
    LastError := 1;
    PgiBox('Suppression impossible. Des profils de paie utilisent cette' +
      ' rubrique.', Ecran.caption);
    Exit;
  end;

//Recherche enregistrement associé multidossier   PT41
  if (V_PGI.ModePCL = '1') and (GetField('PCT_PREDEFINI') <> 'DOS') then
  begin
    NomChamp[2] := 'PPM_RUBRIQUE';
    ExisteCod := RechEnrAssocier('PROFILRUB', NomChamp, Valchamp);
    if (ExisteCod = TRUE) then
    begin
      LastError := 1;
      PgiBox('Suppression impossible. Des profils de paie d''autres dossiers' +
        ' utilisent cette rubrique.', Ecran.caption);
      Exit;
    end;
  end;

  if (ExisteCod = False) then
  begin
    ExecuteSQL('DELETE FROM CUMULRUBRIQUE WHERE' +
      ' PCR_NATURERUB="' + ValChamp[1] + '" AND' +
      ' ##PCR_PREDEFINI## PCR_RUBRIQUE="' + ValChamp[2] + '"');
//PT45
    ExecuteSQL('DELETE FROM CUMULRUBDOSSIER WHERE' +
      ' PKC_NATURERUB="' + ValChamp[1] + '" AND' +
      ' PKC_RUBRIQUE="' + ValChamp[2] + '"');
//FIN PT45


    ExecuteSQL('DELETE FROM VENTICOTPAIE WHERE' +
      ' ##PVT_PREDEFINI## PVT_RUBRIQUE="' + ValChamp[2] + '"');
    ExecuteSQL('DELETE FROM DUCSAFFECT WHERE' +
      ' ##PDF_PREDEFINI## PDF_RUBRIQUE="' + ValChamp[2] + '"');
//Rechargement des tablettes
//A AJOUTER L'ANALYTIQUE COMPTABILITE   DELETE FROM VENTIL WHERE V_NATURE="RC" AND V_IDENTIFIANT="Rubrique"
//  RR pour remuneration
//PT44
    if Assigned(Trace) then FreeAndNil(Trace);
    Trace := TStringList.Create; //PT44
    Trace.Add('SUPPRESSION COTISATION ' + GetField('PCT_RUBRIQUE') + ' ' + GetField('PCT_LIBELLE'));
    CreeJnalEvt('003', '081', 'OK', nil, nil, Trace);
    FreeAndNil(Trace);
//FIN PT44
  end;
end;

//============================================================================================//
//                                 PROCEDURE On Update Record
//============================================================================================//

procedure TOM_Cotisation.OnUpdateRecord;
var
  Onglet: TPageControl;
{$IFNDEF EAGLCLIENT}
  Combo: THDBValComboBox;
  imprim: TDBCheckBox;
{$ELSE}
  Combo: THValComboBox;
  imprim: TCheckBox;
{$ENDIF}
  ThemeCot, rubrique, Nature, Organisme, CodeTranche, TypeBase, mes, mes1, mes2, libelle: string;
  vide, Predef: string;
  ExistCod, OkOk: Boolean; // PT33
//  ordre: integer;
begin
  inherited;

  LeStatut := DS.State; //PT46
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PCT_RUBRIQUE')
  else
    if (DerniereCreate = GetField('PCT_RUBRIQUE')) then OnFerme := True; // le bug arrive on se casse !!!
  { DEB PT28 }
  if (GetField('PCT_THEMECOT') = 'RDC') and (GetField('PCT_ORDREETAT') <> 4) then
    SetField('PCT_ORDREETAT', 4)
  else
    if (GetField('PCT_THEMECOT') <> 'RDC') and (GetField('PCT_ORDREETAT') <> 3) then
      SetField('PCT_ORDREETAT', 3);
  { FIN PT28 }
  if mode = 'DUPLICATION' then exit;
  vide := '';
  // Gestion de la nature de la cotisation
  // Nature.onclick affichage de l'onglet PREGUL ou GBASE selon nature choisie
  // Dans chaque onglets,Utilisation des mêmes Datafield  PCT_TYPEBASE et PCT_BASECOTISATION
  // Distinction des Datafields sur la propriété Name du control
  // Pour ne pas,Lors de la validation,écraser les valeurs des datafields dans les 2 onglets
  // on récupere la valeur par les TControl puis on les réaffecte après réinitialisation des autres champs
  if Ecran is TFFiche then TFFiche(Ecran).Retour := '0';

  {ValTypBas:=GetField('PCT_TYPEBASE');
  ValBasCot:=GetField('PCT_BASECOTISATION');
  ValTra:=GetField('PCT_CODETRANCHE');

  if Nature='COT' then ControlParent:=TTabSheet(GetControl('PREGUL'));
  if Nature='BAS' then ControlParent:=TGroupBox(GetControl('GBASE'));

  // Fonction qui réinit à Null les champs de l'onglet non affiché
  If ControlParent<>nil then RecupControl(ControlParent);

  //Réaffectation des valeurs
  if ValTypBas<>'' then SetField('PCT_TYPEBASE',ValTypBas);
  if Nature='COT' then ;//SetField('PCT_CODETRANCHE',ValTra);
  if ValBasCot<>'' then SetField('PCT_BASECOTISATION',ValBasCot);   }
  Nature := GetField('PCT_NATURERUB');
  if (Nature = 'BAS') then Setfield('PCT_ORGANISME', vide);
//  ordre := 3; { PT28 Mise en commentaire }
//  Setfield('PCT_ORDREETAT', ordre);

  if (DS.State = dsinsert) then
  begin
    if (GetField('PCT_PREDEFINI') <> 'DOS') then
      SetField('PCT_NODOSSIER', '000000')
    else
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
      SetField('PCT_NODOSSIER', PgRendNoDossier());
  end;
  // DEB PT33
  OkOk := ControlChampValeur(Nature); //Test les valeurs des types renseigné
  if OkOk then LastError := 1;
  // FIN PT33
  //////********************************** TEST DE VALIDATION
    // Recherche si Gestion Associe renseigné
  NomChamp[1] := 'PCR_NATURERUB';
  ValChamp[1] := GetField('PCT_NATURERUB');
  NomChamp[2] := '##PCR_PREDEFINI## PCR_RUBRIQUE';
  ValChamp[2] := GetField('PCT_RUBRIQUE'); //PT13
  NomChamp[3] := '';
  ValChamp[3] := ''; //PT4
  NomChamp[4] := '';
  ValChamp[4] := '';
  ExistCod := RechEnrAssocier('CUMULRUBRIQUE', NomChamp, ValChamp);

  Onglet := TPageControl(GetControl('Pages'));
  ThemeCot := Getfield('PCT_THEMECOT');
  rubrique := Getfield('PCT_RUBRIQUE');
  libelle := Getfield('PCT_LIBELLE');
{$IFNDEF EAGLCLIENT}
  Combo := THDBValComboBox(GetControl('PCT_ORGANISME'));
{$ELSE}
  Combo := THValComboBox(GetControl('PCT_ORGANISME'));
{$ENDIF}
  if combo <> nil then Organisme := Combo.Value else Organisme := '';
  //Getfield('PCT_ORGANISME');
  CodeTranche := GetField('PCT_CODETRANCHE');
  if Nature = 'BAS' then CodeTranche := 'Base affiché';
  TypeBase := GetField('PCT_TYPEBASE');
// DEB PT34
{$IFDEF EAGLCLIENT}
  if (TypeBase = '') and (Nature = 'BAS') then
    TypeBase := GetField('PCT_TYPEBASE_');
{$ENDIF}
// FIN PT34
  Mes := '';

  if (rubrique <> '') and (Libelle <> '') then
    if (rubrique = '0000') then
    begin
      LastError := 1;
//      LastErrorMsg := 'Vous devez renseigner une valeur superieure à 0'; PT27
      PgiBox('Vous devez renseigner un ordre de rubrique supérieur à 0.', Ecran.caption);
      SetFocusControl('PCT_NATURERUB');
    end;

  if ((rubrique <> '') or (rubrique <> '0000')) and (Libelle <> '') and (nature = '') then
  begin
    LastError := 1;
//    LastErrorMsg := 'Vous devez renseigner la nature de la cotisation'; PT27
    PgiBox('Vous devez renseigner la nature de la cotisation.', Ecran.caption);
    SetFocusControl('PCT_RUBRIQUE');
  end;

  if ((rubrique <> '') or (rubrique <> '0000')) and (Libelle <> '') and (nature <> '') then
  begin
    if ThemeCot = '' then
    begin
      mes1 := '#13#10#13#10    - le thème de la cotisation';
      Mes := Mes1;
      Onglet.ActivePage := Onglet.Pages[0];
    end;
    if (Nature = 'COT') then
    begin
      if Organisme = '' then
      begin
        mes2 := '#13#10#13#10    - l''organisme de la cotisation';
        Mes := Mes + Mes2;
        Onglet.ActivePage := Onglet.Pages[0];
      end;
      if Organisme = '' then
      begin
        SetFocusControl('PCT_ORGANISME');
      end;
    end;
    if ThemeCot = '' then
    begin
      SetFocusControl('PCT_THEMECOT');
    end;


    if (mes = '') and (Nature = 'COT') and ((GetField('PCT_TYPETAUXSAL') <> '') or (GetField('PCT_TYPETAUXPAT') <> '')) then
    begin
      if (TypeBase = '') then
      begin
        mes1 := '#13#10#13#10    - la base de cotisation';
        Mes := Mes1;
        Onglet.ActivePage := Onglet.Pages[1];
      end;

      if (CodeTranche = '') then
      begin
        mes2 := '#13#10#13#10    - la tranche Plafond';
        Mes := Mes + Mes2;
        Onglet.ActivePage := Onglet.Pages[1];
      end;
      if (CodeTranche = '') then SetFocusControl('PCT_CODETRANCHE');
      if (TypeBase = '') then SetFocusControl('PCT_TYPEBASE');
    end;

    if (mes = '') and (Nature = 'BAS') then
    begin
      if (TypeBase = '') then
      begin
        mes1 := '#13#10#13#10    - la base de cotisation';
        Mes := Mes1;
        Onglet.ActivePage := Onglet.Pages[2];
      end;
      if (TypeBase = '') then SetFocusControl('PCT_TYPEBASE_');
    end;

    ///*******Affichage de la boite à message si champs non renseigné
    if (mes <> '') then
    begin
      LastError := 2;
///      LastErrorMsg := 'Vous devez renseigner ' + mes + ''; PT27
      PgiBox('Vous devez renseigner ' + Mes, Ecran.caption);
    end;
  end;
{$IFNDEF EAGLCLIENT}
  imprim := TDBCheckBox(GetControl('PCT_IMPRIMABLE'));
{$ELSE}
  imprim := TCheckBox(GetControl('PCT_IMPRIMABLE'));
{$ENDIF}
  if (mes = '') and ((rubrique <> '0000') or (rubrique <> '')) and (libelle <> '') and (imprim <> nil) then
  begin
    {Onglet.ActivePage:=Onglet.Pages[0]; PT18-2 Mise en commentaire }
{On change d'avis ==> possiblilté d'avoir une remuneration non imprimable qui alimente des cumuls 07/09/00
if (imprim.Checked=True) and (ExistCod=False) then
           begin   LastError:=4;    SetFocusControl('PCT_IMPRIMABLE');
                   LastErrorMsg:='Veuillez #13#10    alimenter vos cumuls,#13#10 ou #13#10    décocher l''option " rubrique imprimable sur le bulletin "';
           end;}
    if (imprim.Checked = False) and (ExistCod = False) then
      HShowMessage('5;Cotisation :;Vous ne gérez aucun cumul !;E;O;O;O;;;', '', '');
    //if (ExistCod=True) and (imprim.checked=False) then  SetField('PCT_IMPRIMABLE','X');
  end;

  Predef := GetField('PCT_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 1;
//    LastErrorMsg := 'Vous devez renseigner le champ prédéfini'; PT27
    PgiBox('Vous devez renseigner le champ prédéfini.', Ecran.caption);
    SetFocusControl('PCT_PREDEFINI');
  end;
{$IFDEF EAGLCLIENT}
  if GetField('PCT_NATURERUB') = 'BAS' then
  begin
    SetField('PCT_BASECOTISATION', GetControlText('PCT_BASECOTISATION_'));
    SetField('PCT_TYPEBASE', GetControlText('PCT_TYPEBASE_')); // PT34
  end;
{$ENDIF}

  { DEB PT16 Validation inutile si pas acces modification gestion associée }
  if (CEG = False) and (Predef = 'CEG') then SetControlChecked('CBACCESFICHE', False);
  if (STD = False) and (Predef = 'STD') then SetControlChecked('CBACCESFICHE', False);
  if (DOS = False) and (Predef = 'DOS') then SetControlChecked('CBACCESFICHE', False);
  if LastError = 0 then
  begin
    PgUpdateDateAgl(GetField('PCT_RUBRIQUE')); //PT21 code déplacé dans procédure
  end;
  { FIN PT16 }
// PT20 FQ 11029 Contrôle des bornes de validité
  if ((GetField('PCT_DU') <> '') and (GetField('PCT_DU') <> '00')) and (GetField('PCT_AU') = '') then
  begin
    LastError := 1;
    //LastErrorMsg := 'Vous devez renseigner la borne de fin de validité'; PT27
    PgiBox('Vous devez renseigner la borne de fin de validité.', Ecran.caption);
    SetFocusControl('PCT_AU');
  end;
  // FIN PT20
end;


{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 19/02/2003
Modifié le ... :   /  /
Description .. : Affect onglet PCALCUL ou PREGUL selon Nature de la rubrique
Suite ........ : Rend enabled ou non des zones utilisées pour diverses
Suite ........ : traitements
Mots clefs ... : PAIE;COTISATION
*****************************************************************}

procedure TOM_Cotisation.AffecteOnglet();
(* PT14 Déclaration ancienne gestion trop lourde supprimée *)
begin
  { DEB PT14 }
  { Onglet }
  SetControlProperty('PCALCUL', 'TabVisible', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlProperty('PREGUL', 'TabVisible', (GetField('PCT_NATURERUB') = 'BAS'));
  { GroupBox Etat des réductions }
  SetControlEnabled('PCT_REDUCBASSAL', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_REDUCREPAS', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_ALLEGEMENTA2', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_MAJORATA2', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_MINORATA2', (GetField('PCT_NATURERUB') = 'COT'));
  { GroupBox DUCS }
  SetControlEnabled('PCT_ORDREAT', (GetField('PCT_NATURERUB') = 'COT'));
  { GroupBox Rubrique de Cotisation }
  SetControlEnabled('PCT_BASECSGCRDS', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_BASECRDS', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_DADSMONTTSS', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_DADSEXOCOT', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('LDADSEXOCOT', (GetField('PCT_NATURERUB') = 'COT'));
  SetControlEnabled('PCT_DADSEPARGNE', (GetField('PCT_NATURERUB') = 'COT')); //PT22
  { GroupBox Bases de Cotisation }
  SetControlEnabled('PCT_BRUTSS', (GetField('PCT_NATURERUB') = 'BAS'));
  SetControlEnabled('PCT_PLAFONDSS', (GetField('PCT_NATURERUB') = 'BAS'));
  SetControlEnabled('PCT_DADSTOTIMPTSS', (GetField('PCT_NATURERUB') = 'BAS'));
  SetControlEnabled('PCT_DADSEXOBASE', (GetField('PCT_NATURERUB') = 'BAS'));
  SetControlEnabled('LDADSEXOBASE', (GetField('PCT_NATURERUB') = 'BAS'));
  { FIN PT14 }

  (* PT14 Ancien gestion trop lourde , lignes de code supprimées *)
end;
//============================================================================================//
//                                 PROCEDURE On Change Field
//============================================================================================//

procedure TOM_Cotisation.OnChangeField(F: TField);
var
  TypeChamp, Champ, ValChamp, libelle, vide: string;
{$IFNDEF EAGLCLIENT}
  Mois1, Mois2, Mois3, Mois4: THDBValComboBox;
  ControlEdit: THDBEdit;
{$ELSE}
  Mois1, Mois2, Mois3, Mois4: THValComboBox;
  ControlEdit: THEdit;
{$ENDIF}
  onglet: TPageControl;
  Rubrique, TempRub, MsgSaisie, CodeNature, mes, pred: string;
  icode: integer;
  OKRub: boolean;
  ValAu, ValDu, tempo, code: integer;
begin
  inherited;

  if mode = 'DUPLICATION' then exit;
  vide := '';
  Champ := '';
  TypeChamp := '';
  MsgSaisie := 'Veuiller saisir une valeur existante';
  Onglet := TpageControl(GetControl('Pages'));
  //Test de conformité du code de la rubrique   4 positions affectés

  if (F.FieldName = 'PCT_RUBRIQUE') then
  begin
    Rubrique := Getfield('PCT_RUBRIQUE');
    if (Rubrique = '') then exit;
    if ((isnumeric(Rubrique)) and (Rubrique <> '    ')) then
    begin
      iCode := strtoint(trim(Rubrique));
      TempRub := ColleZeroDevant(iCode, 4);
      if (DS.State = dsinsert) and (TempRub <> '') and (GetField('PCT_PREDEFINI') <> '') then
      begin
        OKRub := TestRubrique(GetField('PCT_PREDEFINI'), TempRub, 0);
        if (OkRub = False) or (Rubrique = '0000') then
        begin
          Mes := MesTestRubrique('COT', GetField('PCT_PREDEFINI'), 0);
          HShowMessage('2;Code Erroné: ' + TempRub + ' ;' + mes + ';W;O;O;;;', '', '');
          TempRub := '';
        end;
      end;
      if TempRub <> Rubrique then
      begin
        if GetField('PCT_RUBRIQUE') <> TempRub then SetField('PCT_RUBRIQUE', TempRub); //PT29
        SetFocusControl('PCT_RUBRIQUE');
      end;
    end;
  end;


  if (F.FieldName = 'PCT_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PCT_PREDEFINI');
    Rubrique := (GetField('PCT_RUBRIQUE'));
    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie CEGID', 'Accès refusé'); // PT32
      Pred := 'DOS';
      SetControlProperty('PCT_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Standard', 'Accès refusé'); // PT32
      Pred := 'DOS';
      SetControlProperty('PCT_PREDEFINI', 'Value', Pred);
    end;
    if (rubrique <> '') and (Pred <> '') then
    begin
      OKRub := TestRubrique(Pred, rubrique, 0);
      if (OkRub = False) or (Rubrique = '0000') then
      begin
        Mes := MesTestRubrique('COT', Pred, 0);
        HShowMessage('2;Code Erroné: ' + Rubrique + ' ;' + mes + ';W;O;O;;;', '', '');
        if GetField('PCT_RUBRIQUE') <> Vide then SetField('PCT_RUBRIQUE', vide); // PT29
        if Pred <> GetField('PCT_PREDEFINI') then SetField('PCT_PREDEFINI', pred);
        SetFocusControl('PCT_RUBRIQUE');
        exit;
      end;
    end;
    if Pred <> GetField('PCT_PREDEFINI') then SetField('PCT_PREDEFINI', pred);
  end;

  if (F.FieldName = 'PCT_NATURERUB') then
  begin
    AffecteOnglet(); //Affiche l'ONGLET CALCUL OU REGUL
    if GetField('PCT_NATURERUB') = 'COT' then
    begin
      SetControlEnabled('PCT_ORGANISME', True);
      SetControlEnabled('PCT_SOUMISMALAD', False);
      SetControlChecked('PCT_SOUMISMALAD', False);
      SetControlEnabled('PCT_ASSIETTEBRUT', False);
      SetControlChecked('PCT_ASSIETTEBRUT', False);
      SetControlEnabled('PCT_PRECOMPTEASS', True);
      SetControlEnabled('PCT_ETUDEDROIT', True);
      SetControlVisible('PCT_PRESFINMOIS', TRUE); // PT15
    end;
    if GetField('PCT_NATURERUB') = 'BAS' then
    begin
      SetControlEnabled('PCT_ORGANISME', False);
      if GetField('PCT_ORGANISME') <> vide then SetField('PCT_ORGANISME', vide); //PT29
      SetControlEnabled('PCT_SOUMISMALAD', True);
      SetControlEnabled('PCT_ASSIETTEBRUT', True);
      SetControlEnabled('PCT_PRECOMPTEASS', False);
      SetControlChecked('PCT_PRECOMPTEASS', False);
      SetControlEnabled('PCT_ETUDEDROIT', False);
      SetControlChecked('PCT_ETUDEDROIT', False);
      SetControlVisible('PCT_PRESFINMOIS', FALSE); // PT15
    end;

  end;
 { DEB PT28 }
  if (F.FieldName = 'PCT_THEMECOT') and (GetField('PCT_THEMECOT') <> '') then
  begin
    if (GetField('PCT_THEMECOT') = 'RDC') and (GetField('PCT_ORDREETAT') <> 4) then
      SetField('PCT_ORDREETAT', 4)
    else
      if (GetField('PCT_THEMECOT') <> 'RDC') and (GetField('PCT_ORDREETAT') <> 3) then
        SetField('PCT_ORDREETAT', 3);
  end;
  { FIN PT28 }

  if (F.FieldName = 'PCT_TYPEBASE') or (F.FieldName = 'PCT_TYPEBASE_') then // PT34
  begin
    // FieldName PCT_TYPEBASE de l'onglet PCalcul
    CodeNature := GetField('PCT_NATURERUB');

    //if (Onglet.ActivePage=Onglet.Pages[1]) then
    if (CodeNature = 'COT') or (Onglet.ActivePage = Onglet.Pages[1]) then
    begin
      begin
        TypeChamp := GetField('PCT_TYPEBASE'); { PT39 Refonte d'une partie du code }
        champ := 'PCT_BASECOTISATION';
        ValChamp := GetField('PCT_BASECOTISATION');
        libelle := 'LBLBASECOT';
        AffecteLookupChamp(Champ, TypeChamp, libelle);
        if (GetField('PCT_TYPEBASE') = 'BDC') then SetControlText('PCT_BASECOTISATION_', vide);
        //SetField('PCT_BASECOTISATION',vide);
        if (TypeChamp = '') then
        begin
          if GetField(champ) <> vide then Setfield(champ, vide); //PT29
          SetControlEnabled(Champ, False);
          SetControlChecked('PCT_BASEIMP', False);
          SetControlEnabled('PCT_BASEIMP', False);
          SetControlEnabled('PCT_DECBASE', False);
          if GetField('PCT_DECBASE') <> '2' then SetField('PCT_DECBASE', '2'); //PT29
          SetControlEnabled('PCT_CODETRANCHE', False); // SetField('PCT_CODETRANCHE',vide);
        end
        else if (TypeChamp <> '') then
        begin
          SetControlEnabled(Champ, True);
          SetControlEnabled('PCT_BASEIMP', True); //SetControlChecked('PCT_BASEIMP',True);
          if GetField('PCT_BASEIMP') <> 'X' then SetField('PCT_BASEIMP', 'X'); //PT29
          SetControlEnabled('PCT_DECBASE', True);
          if (GetField('PCT_TYPEBASE') = 'BDC') then
          begin
            SetControlEnabled('PCT_CODETRANCHE', True); //SetField('PCT_CODETRANCHE',vide);
          end;
          if (GetField('PCT_TYPEBASE') <> 'BDC') and (GetField('PCT_TYPEBASE') <> '') then
          begin
            if GetField('PCT_CODETRANCHE') <> 'TOT' then SetField('PCT_CODETRANCHE', 'TOT'); //PT29
            SetControlEnabled('PCT_CODETRANCHE', False);
          end;
{$IFNDEF EAGLCLIENT}
          ControlEdit := THDBEdit(GetControl('PCT_BASECOTISATION'));
{$ELSE}
          ControlEdit := THEdit(GetControl('PCT_BASECOTISATION'));
{$ENDIF}
          if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
          if LastError = 1 then SetField('PCT_BASECOTISATION', vide);
        end;
      end;
    end;
    // FieldName PCT_TYPEBASE de l'onglet PRegul
//if (Onglet.ActivePage=Onglet.Pages[2]) then
    if (CodeNature = 'BAS') or (Onglet.ActivePage = Onglet.Pages[2]) then
    begin
      TypeChamp := GetControlText('PCT_TYPEBASE_'); { PT39 Refonte d'une partie du code }
      champ := 'PCT_BASECOTISATION_';
      ValChamp := GetControlText('PCT_BASECOTISATION_');
      libelle := 'LBLBASECOTT';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
      if (GetControlText('PCT_TYPEBASE_') = '') then SetControlEnabled(Champ, False)
      else
        if (GetControlText('PCT_TYPEBASE_') <> '') then
        begin
          SetControlEnabled(Champ, True);
{$IFNDEF EAGLCLIENT}
          ControlEdit := THDBEdit(GetControl('PCT_BASECOTISATION_'));
{$ELSE}
          ControlEdit := THEdit(GetControl('PCT_BASECOTISATION_'));
{$ENDIF}
          if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
          if LastError = 1 then SetField('PCT_BASECOTISATION', vide);
        end;
    end;
  end;

  if (F.FieldName = 'PCT_BASECOTISATION') then
  begin
    // FieldName PCT_BASECOTISATION de l'onglet PCalcul
    if (Onglet.ActivePage = Onglet.Pages[1]) then
    begin
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_BASECOTISATION'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_BASECOTISATION'));
{$ENDIF}
      if (ControlEdit <> nil) and (GetField('PCT_BASECOTISATION') <> '') then
      begin
        LastError := TestValLookupChamp(ControlEdit, GetField('PCT_BASECOTISATION'));
        if LastError = 1 then
        begin
          LastErrorMsg := '"' + GetField('PCT_BASECOTISATION') + '" :#13#10 ' + MsgSaisie + '';
          if GetField('PCT_BASECOTISATION') <> vide then SetField('PCT_BASECOTISATION', vide); //PT29
          SetFocusControl('PCT_TYPEBASE');
          LastError := 1;
        end;
        TypeChamp := GetField('PCT_TYPEBASE'); // PT43  Mauvais nom du champ type mis le type au lieu du champ
        champ := 'PCT_BASECOTISATION';
        ValChamp := GetField('PCT_BASECOTISATION');
        libelle := 'LBLBASECOT';
        AffecteLookupChamp(Champ, TypeChamp, libelle);
      end;
    end;
    if (Onglet.ActivePage = Onglet.Pages[2]) then // FieldName PCT_BASECOTISATION de l'onglet PREGUL
    begin
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_BASECOTISATION_'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_BASECOTISATION_'));
{$ENDIF}
      if (ControlEdit <> nil) and (GetField('PCT_BASECOTISATION') <> '') then
      begin
        LastError := TestValLookupChamp(ControlEdit, ControlEdit.Text);
        if LastError = 1 then
        begin
          LastErrorMsg := '"' + ControlEdit.Text + '" :#13#10  ' + MsgSaisie + '';
          if GetField('PCT_BASECOTISATION') <> vide then SetField('PCT_BASECOTISATION', vide); //PT29
          SetFocusControl('PCT_TYPEBASE_');
          LastError := 1;
        end;
      end;
    end;

  end;

  //***********************Onglet Calcul ( RUBRIQUE DE COTISATION) *********************************/

    // Champ qui génère ElipsisButton + tablettes associées
  if (F.FieldName = 'PCT_TYPETAUXSAL') then
  begin
    TypeChamp := GetField('PCT_TYPETAUXSAL');
    Champ := 'PCT_TAUXSAL';
    ValChamp := GetField('PCT_TAUXSAL');
    libelle := 'LBLPTXSAL';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin // Controles associés no enabled  et réinit // Forfait Salarial Enabled
      if GetField(champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_TAUXSAL', False);
      SetControlChecked('PCT_TXSALIMP', False);
      SetControlEnabled('PCT_TXSALIMP', False);
      SetControlEnabled('PCT_DECTXSAL', False);
      if GetField('PCT_DECTXSAL') <> '2' then SetField('PCT_DECTXSAL', '2'); //PT29
      SetControlEnabled('PCT_TYPEFFSAL', True);
    end
    else
      if (TypeChamp <> '') then
      begin // Controles associés enabled
        SetControlEnabled('PCT_TAUXSAL', True);
      //SetControlChecked('PCT_TXSALIMP',True);
        SetControlEnabled('PCT_TXSALIMP', True);
        if GetField('PCT_TXSALIMP') <> 'X' then SetField('PCT_TXSALIMP', 'X'); //PT29
        SetControlEnabled('PCT_DECTXSAL', True);
        SetControlEnabled('PCT_TYPEFFSAL', False); // Forfait Salarial et Champs Associés no Enabled
        SetControlEnabled('PCT_FFSAL', False);
        SetControlChecked('PCT_FFSALIMP', False);
        SetControlEnabled('PCT_FFSALIMP', False);
{$IFNDEF EAGLCLIENT}
        ControlEdit := THDBEdit(GetControl('PCT_TAUXSAL'));
{$ELSE}
        ControlEdit := THEdit(GetControl('PCT_TAUXSAL'));
{$ENDIF}
        if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
        if (LastError = 1) and (GetField('PCT_TAUXSAL') <> vide) then //PT29
          SetField('PCT_TAUXSAL', vide);
      end;
  end;

  if (F.FieldName = 'PCT_TYPETAUXPAT') then
  begin
    TypeChamp := GetField('PCT_TYPETAUXPAT');
    Champ := 'PCT_TAUXPAT';
    ValChamp := GetField('PCT_TAUXPAT');
    libelle := 'LBLTXPAT';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin // Controles associés no enabled  et réinit
      if GetField(champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_TAUXPAT', False);
      SetControlChecked('PCT_TXPATIMP', False);
      SetControlEnabled('PCT_TXPATIMP', False);
      SetControlEnabled('PCT_DECTXPAT', False);
      if GetField('PCT_DECTXPAT') <> '2' then SetField('PCT_DECTXPAT', '2'); //PT29
      SetControlEnabled('PCT_TYPEFFPAT', True); // Forfait Patronal Enabled
    end
    else
    begin // Controles associés enabled
      SetControlEnabled('PCT_TAUXPAT', True);
      //SetControlChecked('PCT_TXPATIMP',True);
      SetControlEnabled('PCT_TXPATIMP', True);
      if GetField('PCT_TXPATIMP') <> 'X' then SetField('PCT_TXPATIMP', 'X'); //PT29
      SetControlEnabled('PCT_DECTXPAT', True);
      SetControlEnabled('PCT_TYPEFFPAT', False);
      SetControlEnabled('PCT_FFPAT', False); // Forfait Patronal et Champs Associés no Enabled
      SetControlChecked('PCT_FFPATIMP', False);
      SetControlEnabled('PCT_FFPATIMP', False);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_TAUXPAT'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_TAUXPAT'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
      if (LastError = 1) and (GetField('PCT_TAUXPAT') <> vide) then //PT29
        SetField('PCT_TAUXPAT', vide);
    end;
  end;

  if (F.FieldName = 'PCT_TYPEFFSAL') then
  begin
    TypeChamp := GetField('PCT_TYPEFFSAL');
    Champ := 'PCT_FFSAL';
    ValChamp := GetField('PCT_FFSAL');
    libelle := 'LBLFFSAL';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      if GetField(champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_FFSAL', False); //Controles associés no enabled  et réinit
      SetControlChecked('PCT_FFSALIMP', False);
      SetControlEnabled('PCT_FFSALIMP', False);
      SetControlEnabled('PCT_DECMTSAL', False);
      if GetField('PCT_DECMTSAL') <> '2' then SetField('PCT_DECMTSAL', '2'); //PT29
      SetControlEnabled('PCT_TYPETAUXSAL', True); // Taux Salarial Enabled
    end
    else
    begin
      SetControlEnabled('PCT_FFSAL', True); // Controles associés enabled
      //SetControlChecked('PCT_FFSALIMP',True);
      if GetField('PCT_FFSALIMP') <> 'X' then SetField('PCT_FFSALIMP', 'X'); //PT29
      SetControlEnabled('PCT_FFSALIMP', True);
      SetControlEnabled('PCT_DECMTSAL', True);
      SetControlEnabled('PCT_TYPETAUXSAL', False); // Taux Salarial et Champs Associés no Enabled
      SetControlEnabled('PCT_TAUXSAL', False);
      SetControlChecked('PCT_TXSALIMP', False);
      SetControlEnabled('PCT_TXSALIMP', False);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_FFSAL'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_FFSAL'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
      if LastError = 1 then SetField('PCT_FFSAL', vide);
    end;
  end;

  if (F.FieldName = 'PCT_TYPEFFPAT') then
  begin
    TypeChamp := GetField('PCT_TYPEFFPAT');
    Champ := 'PCT_FFPAT';
    ValChamp := GetField('PCT_FFPAT');
    libelle := 'LBLFFPAT';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      if GetField(champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_FFPAT', False); //Controles associés no enabled  et réinit
      SetControlChecked('PCT_FFPATIMP', False);
      SetControlEnabled('PCT_FFPATIMP', False);
      SetControlEnabled('PCT_DECMTPAT', False);
      if GetField('PCT_DECMTPAT') <> '2' then SetField('PCT_DECMTPAT', '2'); //PT29
      SetControlEnabled('PCT_TYPETAUXPAT', True); // Taux Patronal Enabled
    end
    else
    begin
      SetControlEnabled('PCT_FFPAT', True); // Controles associés enabled
      //SetControlChecked('PCT_FFPATIMP',True);
      SetControlEnabled('PCT_FFPATIMP', True);
      if GetField('PCT_FFPATIMP') <> 'X' then SetField('PCT_FFPATIMP', 'X'); //PT29
      SetControlEnabled('PCT_DECMTPAT', True);
      SetControlEnabled('PCT_TYPETAUXPAT', False); // Taux Patronal et Champs Associés no Enabled
      SetControlChecked('PCT_TXPATIMP', False);
      SetControlEnabled('PCT_TXPATIMP', False);
      SetControlEnabled('PCT_TAUXPAT', False);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_FFPAT'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_FFPAT'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
      if LastError = 1 then SetField('PCT_FFPAT', vide);
    end;
  end;

  if (F.FieldName = 'PCT_TYPEMINISAL') then
  begin
    TypeChamp := GetField('PCT_TYPEMINISAL');
    Champ := 'PCT_VALEURMINISAL';
    Valchamp := GetField('PCT_VALEURMINISAL');
    libelle := 'LBLMINSAL';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      if GetField(Champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_VALEURMINISAL', False);
    end
    else
      if (TypeChamp <> '') then
      begin
        SetControlEnabled('PCT_VALEURMINISAL', True);
{$IFNDEF EAGLCLIENT}
        ControlEdit := THDBEdit(GetControl('PCT_VALEURMINISAL'));
{$ELSE}
        ControlEdit := THEdit(GetControl('PCT_VALEURMINISAL'));
{$ENDIF}
        if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
        if LastError = 1 then
          SetField('PCT_VALEURMINISAL', vide);
      end;
  end;

  if (F.FieldName = 'PCT_TYPEMAXISAL') then
  begin
    TypeChamp := GetField('PCT_TYPEMAXISAL');
    Champ := 'PCT_VALEURMAXISAL';
    ValChamp := GetField('PCT_VALEURMAXISAL');
    libelle := 'LBLMAXSAL';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      if GetField(Champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_VALEURMAXISAL', False);
    end
    else
    begin
      SetControlEnabled('PCT_VALEURMAXISAL', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_VALEURMAXISAL'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_VALEURMAXISAL'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
      if LastError = 1 then
        SetField('PCT_VALEURMAXISAL', vide);
    end;
  end;

  if (F.FieldName = 'PCT_TYPEMAXIPAT') then
  begin
    TypeChamp := GetField('PCT_TYPEMAXIPAT');
    Champ := 'PCT_VALEURMAXIPAT';
    ValChamp := GetField('PCT_VALEURMAXIPAT');
    libelle := 'LBLMAXPAT';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      if GetField(champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_VALEURMAXIPAT', False);
    end
    else
    begin
      SetControlEnabled('PCT_VALEURMAXIPAT', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_VALEURMAXIPAT'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_VALEURMAXIPAT'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
      if LastError = 1 then SetField('PCT_VALEURMAXIPAT', vide);
    end;
  end;

  if (F.FieldName = 'PCT_TYPEMINIPAT') then
  begin
    TypeChamp := GetField('PCT_TYPEMINIPAT');
    Champ := 'PCT_VALEURMINIPAT';
    ValChamp := GetField('PCT_VALEURMINIPAT');
    Libelle := 'LBLMINPAT';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      if GetField(Champ) <> vide then SetField(champ, vide); //PT29
      SetControlEnabled('PCT_VALEURMINIPAT', False);
    end
    else
    begin
      SetControlEnabled('PCT_VALEURMINIPAT', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_VALEURMINIPAT'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_VALEURMINIPAT'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, ValChamp);
      if LastError = 1 then SetField('PCT_VALEURMINIPAT', vide);
    end;
  end;

  //******Test sur Valeur Des THDBEdit par rapport au Contenu des tablettes associées *********

  if (F.FieldName = 'PCT_TAUXSAL') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_TAUXSAL'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_TAUXSAL'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_TAUXSAL') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TAUXSAL'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_TAUXSAL') + '" :#13#10  ' + MsgSaisie + '';
        SetField('PCT_TYPETAUXSAL', GetField('PCT_TYPETAUXSAL'));
        SetFocusControl('PCT_TYPETAUXSAL');
        SetField('PCT_TAUXSAL', vide);
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPETAUXSAL');
      Champ := 'PCT_TAUXSAL';
      ValChamp := GetField('PCT_TAUXSAL');
      libelle := 'LBLPTXSAL';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_TAUXPAT') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_TAUXPAT'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_TAUXPAT'));
{$ENDIF}

    if (ControlEdit <> nil) and (GetField('PCT_TAUXPAT') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TAUXPAT'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_TAUXPAT') + '" :#13#10  ' + MsgSaisie + '';
        SetField('PCT_TYPETAUXPAT', GetField('PCT_TYPETAUXPAT'));
        SetFocusControl('PCT_TYPETAUXSAL');
        SetField('PCT_TAUXPAT', vide);
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPETAUXPAT');
      Champ := 'PCT_TAUXPAT';
      ValChamp := GetField('PCT_TAUXPAT');
      libelle := 'LBLTXPAT';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_FFSAL') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_FFSAL'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_FFSAL'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_FFSAL') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_FFSAL'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_FFSAL') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPEFFSAL', GetField('PCT_TYPEFFSAL'));
        SetFocusControl('PCT_TYPEFFSAL');
        SetField('PCT_FFSAL', vide);
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPEFFSAL');
      Champ := 'PCT_FFSAL';
      ValChamp := GetField('PCT_FFSAL');
      libelle := 'LBLFFSAL';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_FFPAT') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_FFPAT'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_FFPAT'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_FFPAT') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_FFPAT'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_FFPAT') + '" :#13#10  ' + MsgSaisie + '';
        SetField('PCT_TYPEFFPAT', GetField('PCT_TYPEFFPAT'));
        SetField('PCT_FFPAT', vide);
        SetFocusControl('PCT_TYPEFFPAT');
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPEFFPAT');
      Champ := 'PCT_FFPAT';
      ValChamp := GetField('PCT_FFPAT');
      libelle := 'LBLFFPAT';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_VALEURMINISAL') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_VALEURMINISAL'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_VALEURMINISAL'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_VALEURMINISAL') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_VALEURMINISAL'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_VALEURMINISAL') + '" :#13#10  ' + MsgSaisie + '';
        SetField('PCT_TYPEMINISAL', GetField('PCT_TYPEMINISAL'));
        SetFocusControl('PCT_TYPEMINISAL');
        SetField('PCT_VALEURMINISAL', vide);
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPEMINISAL');
      Champ := 'PCT_VALEURMINISAL';
      Valchamp := GetField('PCT_VALEURMINISAL');
      libelle := 'LBLMINSAL';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_VALEURMAXISAL') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_VALEURMAXISAL'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_VALEURMAXISAL'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_VALEURMAXISAL') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_VALEURMAXISAL'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_VALEURMAXISAL') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPEMAXISAL', GetField('PCT_TYPEMAXISAL'));
        SetField('PCT_VALEURMAXISAL', vide);
        SetFocusControl('PCT_TYPEMAXISAL');
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPEMAXISAL');
      Champ := 'PCT_VALEURMAXISAL';
      ValChamp := GetField('PCT_VALEURMAXISAL');
      libelle := 'LBLMAXSAL';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_VALEURMINIPAT') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_VALEURMINIPAT'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_VALEURMINIPAT'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_VALEURMINIPAT') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_VALEURMINIPAT'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_VALEURMINIPAT') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPEMINIPAT', GetField('PCT_TYPEMINIPAT'));
        SetField('PCT_VALEURMINIPAT', vide);
        SetFocusControl('PCT_TYPEMINIPAT');
        LastError := 1;
      end;

      TypeChamp := GetField('PCT_TYPEMINIPAT');
      Champ := 'PCT_VALEURMINIPAT';
      ValChamp := GetField('PCT_VALEURMINIPAT');
      libelle := 'LBLMINPAT';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  if (F.FieldName = 'PCT_VALEURMAXIPAT') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_VALEURMAXIPAT'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_VALEURMAXIPAT'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_VALEURMAXIPAT') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_VALEURMAXIPAT'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_VALEURMAXIPAT') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPEMAXIPAT', GetField('PCT_TYPEMAXIPAT'));
        SetField('PCT_VALEURMAXIPAT', vide);
        SetFocusControl('PCT_TYPEMAXIPAT');
        LastError := 1;
      end;
      TypeChamp := GetField('PCT_TYPEMAXIPAT');
      Champ := 'PCT_VALEURMAXIPAT';
      ValChamp := GetField('PCT_VALEURMAXIPAT');
      libelle := 'LBLMAXPAT';
      AffecteLookupChamp(Champ, TypeChamp, libelle);
    end;
  end;

  //***********************Onglet TRANCHE (RUBRIQUE DE BASE) *********************************/
   // Champ qui génère ElipsisButton + tablettes associées
  if (F.FieldName = 'PCT_TYPEPLAFOND') then
  begin
    TypeChamp := GetField('PCT_TYPEPLAFOND');
    Champ := 'PCT_PLAFOND';
    libelle := 'LBLPLAF';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if TypeChamp = '' then
    begin
      SetControlEnabled('PCT_PLAFOND', False);
      if GetField('PCT_PLAFOND') <> vide then SetField('PCT_PLAFOND', vide); //PT29
    end
    else
    begin
      SetControlEnabled('PCT_PLAFOND', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_PLAFOND'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_PLAFOND'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, GetField('PCT_PLAFOND'));
      if LastError = 1 then SetField('PCT_PLAFOND', vide);
    end;

  end;

  if (F.FieldName = 'PCT_TYPETRANCHE1') then
  begin
    TypeChamp := GetField('PCT_TYPETRANCHE1');
    Champ := 'PCT_TRANCHE1';
    libelle := 'LBLTRA1';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if TypeChamp = '' then
    begin
      SetControlEnabled('PCT_TRANCHE1', False);
      if GetField('PCT_TRANCHE1') <> vide then SetField('PCT_TRANCHE1', vide); //PT29
    end
    else
    begin
      SetControlEnabled('PCT_TRANCHE1', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_TRANCHE1'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_TRANCHE1'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TRANCHE1'));
      if LastError = 1 then SetField('PCT_TRANCHE1', vide);
    end;
  end;

  if (F.FieldName = 'PCT_TYPETRANCHE2') then
  begin
    TypeChamp := GetField('PCT_TYPETRANCHE2');
    Champ := 'PCT_TRANCHE2';
    libelle := 'LBLTRA2';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if (TypeChamp = '') then
    begin
      SetControlEnabled('PCT_TRANCHE2', False);
      if GetField('PCT_TRANCHE2') <> vide then SetField('PCT_TRANCHE2', vide); //PT29
    end
    else
    begin
      SetControlEnabled('PCT_TRANCHE2', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_TRANCHE2'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_TRANCHE2'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TRANCHE2'));
      if LastError = 1 then SetField('PCT_TRANCHE2', vide);
    end;

  end;

  if (F.FieldName = 'PCT_TYPETRANCHE3') then
  begin
    TypeChamp := GetField('PCT_TYPETRANCHE3');
    Champ := 'PCT_TRANCHE3';
    libelle := 'LBLTRA3';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
    if TypeChamp = '' then
    begin
      if GetField('PCT_TRANCHE3') <> vide then SetField('PCT_TRANCHE3', vide); //PT29
      SetControlEnabled('PCT_TRANCHE3', False);
    end
    else
    begin
      SetControlEnabled('PCT_TRANCHE3', True);
{$IFNDEF EAGLCLIENT}
      ControlEdit := THDBEdit(GetControl('PCT_TRANCHE3'));
{$ELSE}
      ControlEdit := THEdit(GetControl('PCT_TRANCHE3'));
{$ENDIF}
      if ControlEdit <> nil then LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TRANCHE3'));
      if LastError = 1 then SetField('PCT_TRANCHE3', vide);
    end;
  end;

  //Test de vraisemblance sur Saisie du champ par rapport au contenu tablettes

  if (F.FieldName = 'PCT_PLAFOND') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_PLAFOND'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_PLAFOND'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_PLAFOND') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_PLAFOND'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_PLAFOND') + '" :#13#10  ' + MsgSaisie + '';
        SetField('PCT_TYPEPLAFOND', GetField('PCT_TYPEPLAFOND'));
        SetFocusControl('PCT_TYPEPLAFOND');
        SetField('PCT_PLAFOND', vide);
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PCT_TRANCHE1') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_TRANCHE1'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_TRANCHE1'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_TRANCHE1') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TRANCHE1'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_TRANCHE1') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPETRANCHE1', GetField('PCT_TYPETRANCHE1'));
        SetField('PCT_TRANCHE1', vide);
        SetFocusControl('PCT_TYPETRANCHE1');
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PCT_TRANCHE2') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_TRANCHE2'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_TRANCHE2'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_TRANCHE2') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TRANCHE2'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_TRANCHE2') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPETRANCHE2', GetField('PCT_TYPETRANCHE2'));
        SetField('PCT_TRANCHE2', vide);
        SetFocusControl('PCT_TYPETRANCHE2');
        LastError := 1;
      end;
    end;
  end;

  if (F.FieldName = 'PCT_TRANCHE3') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlEdit := THDBEdit(GetControl('PCT_TRANCHE3'));
{$ELSE}
    ControlEdit := THEdit(GetControl('PCT_TRANCHE3'));
{$ENDIF}
    if (ControlEdit <> nil) and (GetField('PCT_TRANCHE3') <> '') then
    begin
      LastError := TestValLookupChamp(ControlEdit, GetField('PCT_TRANCHE3'));
      if LastError = 1 then
      begin
        LastErrorMsg := '"' + GetField('PCT_TRANCHE3') + '" :#13#10   ' + MsgSaisie + '';
        SetField('PCT_TYPETRANCHE3', GetField('PCT_TYPETRANCHE3'));
        SetField('PCT_TRANCHE3', vide);
        SetFocusControl('PCT_TYPETRANCHE3');
        LastError := 1;
      end;
    end;
  end;


  //****************************Onglet Validité*******************
  //Gestion des Mois ou Periode de Validité :
  //   -par saisie d'une periode Du ? Au ?
                  // A FAIRE : LE CAS D'UNE PAIE DECALEE
                  // AUTORISE SAISIE DU MOIS DE DECEMBRE POUR COMBO "DU"
  //   -par Saisie de 4 Mois au choix
  { DEB PT39 Refonte d'une partie du code }
  if (F.Fieldname = 'PCT_DU') then
  begin
    begin
      if GetField('PCT_MOIS1') <> vide then SetField('PCT_MOIS1', Vide); //PT29
      if GetField('PCT_MOIS2') <> vide then SetField('PCT_MOIS2', Vide); //PT29
      if GetField('PCT_MOIS3') <> vide then SetField('PCT_MOIS3', Vide); //PT29
      if GetField('PCT_MOIS4') <> vide then SetField('PCT_MOIS4', Vide); //PT29
      SetControlEnabled('PCT_MOIS1', False);
      SetControlEnabled('PCT_MOIS2', False);
      SetControlEnabled('PCT_MOIS3', False);
      SetControlEnabled('PCT_MOIS4', False);
      if (GetField('PCT_DU') = '00') then
      begin
        SetControlProperty('PCT_AU', 'DataType', 'PGVALIDITERUB');
        if GetField('PCT_AU') <> GetField('PCT_DU') then SetField('PCT_AU', GetField('PCT_DU')); //*** ValidAu.value := ValidDu.value;
        SetControlEnabled('PCT_AU', False);
      end;
      if (GetField('PCT_DU') <> '00') then
      begin
        SetControlProperty('PCT_AU', 'DataType', 'PGVALIDITEMOISRUB');
        SetControlEnabled('PCT_AU', True);
      end;
      if (GetField('PCT_DU') = '') then
      begin
        SetControlEnabled('PCT_AU', False); // PT24
        if GetField('PCT_AU') <> '' then SetField('PCT_AU', ''); //PT29
        SetControlEnabled('PCT_MOIS1', True);
        SetControlEnabled('PCT_MOIS2', True);
        SetControlEnabled('PCT_MOIS3', True);
        SetControlEnabled('PCT_MOIS4', True);
      end;
      val(GetField('PCT_DU'), tempo, code);
      ValDu := tempo;
      val(GetField('PCT_AU'), tempo, code);
      ValAu := tempo;

      if (ValAu <> 0) and (ValDu <> 0) then
        if ValAu < ValDu then
        begin
          LastError := 2;
          LastErrorMsg := 'Saisie Erronée: "' + GetField('PCT_DU') + '" #13#10#13#10' +
            'Vous devez saisir un mois antérieur au mois de ' + GetField('PCT_AU') + ' ';
          SetField('PCT_DU', '');
        end;
    end;
  end;

  if (F.Fieldname = 'PCT_AU') then
  begin
    begin
      val(GetField('PCT_DU'), tempo, code);
      ValDu := tempo;
      val(GetField('PCT_AU'), tempo, code);
      ValAu := tempo;
      if (ValAu <> 0) and (ValDu <> 0) then
        if ValAu < ValDu then
        begin
          LastError := 2;
          LastErrorMsg := 'Saisie Erronée: "' + GetField('PCT_AU') + '" #13#10#13#10' +
            'Vous devez saisir un mois postérieur au mois de ' + GetField('PCT_DU') + ' ';
          SetField('PCT_AU', '');
        end;
    end;
  end;

  if (F.Fieldname = 'PCT_MOIS1') then
  begin
{$IFNDEF EAGLCLIENT}
    Mois2 := THDBValComboBox(GetControl('PCT_MOIS2'));
    Mois3 := THDBValComboBox(GetControl('PCT_MOIS3'));
    Mois4 := THDBValComboBox(GetControl('PCT_MOIS4'));
{$ELSE}
    Mois2 := THValComboBox(GetControl('PCT_MOIS2'));
    Mois3 := THValComboBox(GetControl('PCT_MOIS3'));
    Mois4 := THValComboBox(GetControl('PCT_MOIS4'));
{$ENDIF}
    if (Mois2 <> nil) and (Mois3 <> nil) and (Mois4 <> nil) then
    begin
      if (Getfield('PCT_MOIS1') = '') and (Getfield('PCT_MOIS2') = '') and (Getfield('PCT_MOIS3') = '') and (Getfield('PCT_MOIS4') = '') then
        SetControlEnabled('PCT_DU', True);
      if (GetField('PCT_MOIS1') <> '') then
        SetControlEnabled('PCT_DU', False);
      //if (Getfield('PCT_MOIS1')='') then  ReinitMoisValidite(Mois2,Mois3,Mois4);
    end;
  end;


  if (F.Fieldname = 'PCT_MOIS2') then
  begin
    if (Getfield('PCT_MOIS1') = '') and (Getfield('PCT_MOIS2') = '') and (Getfield('PCT_MOIS3') = '') and (Getfield('PCT_MOIS4') = '') then
      SetControlEnabled('PCT_DU', True);
    if (GetField('PCT_MOIS2') <> '') then
      SetControlEnabled('PCT_DU', False);
  //if (Getfield('PCT_MOIS2')='')  then  ReinitMoisValidite(Mois1,Mois3,Mois4);
  end;

  if (F.Fieldname = 'PCT_MOIS3') then
  begin
    if (Getfield('PCT_MOIS1') = '') and (Getfield('PCT_MOIS2') = '') and (Getfield('PCT_MOIS3') = '') and (Getfield('PCT_MOIS4') = '') then
      SetControlEnabled('PCT_DU', True);
    if (GetField('PCT_MOIS3') <> '') then
      SetControlEnabled('PCT_DU', False);
      //if (Getfield('PCT_MOIS3')='') then  ReinitMoisValidite(Mois1,Mois2,Mois4)  ;
  end;

  if (F.Fieldname = 'PCT_MOIS4') then
  begin
    if (Getfield('PCT_MOIS1') = '') and (Getfield('PCT_MOIS2') = '') and (Getfield('PCT_MOIS3') = '') and (Getfield('PCT_MOIS4') = '') then
      SetControlEnabled('PCT_DU', True);
    if (GetField('PCT_MOIS4') <> '') then
      SetControlEnabled('PCT_DU', False);
      //if (Getfield('PCT_MOIS4')='') then  ReinitMoisValidite(Mois1,Mois2,Mois3)  ;
  end;
  { FIN PT39 }
  //PT5
  if (F.Fieldname = 'PCT_DADSEXOBASE') then
  begin
    TypeChamp := 'EXB';
    champ := 'PCT_DADSEXOBASE';
    ValChamp := GetField('PCT_DADSEXOBASE');
    libelle := 'LDADSEXOBAS';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
  end;

  if (F.Fieldname = 'PCT_DADSEXOCOT') then
  begin
    TypeChamp := 'EXC';
    champ := 'PCT_DADSEXOCOT';
    ValChamp := GetField('PCT_DADSEXOCOT');
    libelle := 'LDADSEXOCOT';
    AffecteLookupChamp(Champ, TypeChamp, libelle);
  end;
  //FIN PT5

              // Affichage de l'Elipsis Button et tablette associées
  //if (Champ<>'') and (TypeChamp<>'') then AffecteLookupChamp (Champ, TypeChamp, libelle);

  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
      SetControlEnabled('MEMOPERSO', True); // PT38
    end;
    SetControlEnabled('BInsert', True);
    SetControlEnabled('PCT_PREDEFINI', False);
    SetControlEnabled('PCT_RUBRIQUE', False);
    SetControlEnabled('BGESTIONASS', True);
    SetControlEnabled('BDUCS', (GetField('PCT_NATURERUB') = 'COT')); //PT17-2

    SetControlEnabled('BDUPLIQUER', True);
    SetControlEnabled('BVENTIL', True);
  end;

end;


//============================================================================================//
//                                 PROCEDURE On Load Record
//============================================================================================//
procedure TOM_Cotisation.OnLoadRecord;
var
Q: TQuery;
MonMemo: TMemo;
LabPers: ThLabel;
Personnalise : Boolean;
begin
inherited;
Personnalise:= False; //PT45
if not (DS.State in [dsInsert]) then
   DerniereCreate:= '';
if (mode='DUPLICATION') then
   exit;

OrdreEtat:= GetField ('PCT_ORDREETAT');

SetControlProperty ('ACCES', 'Text', 'FALSE');
AccesPredefini ('TOUS', CEG, STD, DOS);
LectureSeule:= FALSE;
if (Getfield ('PCT_PREDEFINI')='CEG') then
   begin
   LectureSeule:= (CEG=False);
   PaieLectureSeule (TFFiche (Ecran), (CEG=False));
   SetControlEnabled ('BDelete', CEG);
   if (CEG=True) then
      SetControlProperty ('ACCES', 'Text', 'TRUE')
   else
      SetControlProperty ('ACCES', 'Text', 'FALSE');
   SetControlEnabled ('BGESTIONSPEC', True);        //PT48
   end
else
if (Getfield ('PCT_PREDEFINI')='STD') then
   begin
   LectureSeule:= (STD=False);
   PaieLectureSeule (TFFiche (Ecran), (STD=False));
   SetControlEnabled ('BDelete', STD);
   if (STD=True) then
      SetControlProperty ('ACCES', 'Text', 'TRUE')
   else
      SetControlProperty ('ACCES', 'Text', 'FALSE');
   SetControlEnabled ('BGESTIONSPEC', True);        //PT48
   end
else
if (Getfield ('PCT_PREDEFINI')='DOS') then
   begin
   LectureSeule:= (DOS=False);
   PaieLectureSeule (TFFiche (Ecran), False);
   SetControlEnabled ('BDelete', DOS);
   if (DOS=True) then
      SetControlProperty ('ACCES', 'Text', 'TRUE')
   else
      SetControlProperty ('ACCES', 'Text', 'FALSE');
   SetControlEnabled ('BGESTIONSPEC', False);       //PT48
   end;

SetControlEnabled ('BInsert', True);
SetControlEnabled ('PCT_PREDEFINI', False);
SetControlEnabled ('PCT_RUBRIQUE', False);
SetControlEnabled ('BGESTIONASS', True);
SetControlEnabled ('BDUCS', (GetField ('PCT_NATURERUB')='COT'));

if (DS.State in [dsInsert]) then
   begin
   LectureSeule:= FALSE;
   PaieLectureSeule (TFFiche (Ecran), False);
   SetControlEnabled ('PCT_PREDEFINI', True);
   SetControlEnabled ('PCT_RUBRIQUE', True);
   SetControlEnabled ('BVENTIL', False);
   SetControlEnabled ('BDUPLIQUER', False);
   SetControlEnabled ('BGESTIONASS', False);
   SetControlEnabled ('BDUCS', False);
   SetControlEnabled ('BInsert', False);
   SetControlEnabled ('BDelete', False);
   OrdreEtat:= -1;
   end;

AffecteOnglet ();
SetControlVisible ('BDUPLIQUER', True);
PaieConceptPlanPaie (Ecran); // PT46

{$IFDEF EAGLCLIENT}
if (GetField ('PCT_NATURERUB')='BAS') then
   begin
   if (GetControlText ('PCT_BASECOTISATION_')<>GetField ('PCT_BASECOTISATION')) then
      SetControlText ('PCT_BASECOTISATION_', GetField ('PCT_BASECOTISATION'));
   if (GetControlText ('PCT_TYPEBASE_')<>GetField ('PCT_TYPEBASE')) then
      SetControlText ('PCT_TYPEBASE_', GetField ('PCT_TYPEBASE'));
   end;
{$ENDIF}
//Personnalisation Message avertissement
Q:= OpenSql ('SELECT *'+
             ' FROM PGEXCEPTIONS WHERE'+
             ' PEN_RUBRIQUE="'+GetField ('PCT_RUBRIQUE')+'" AND'+
             ' ##PEN_PREDEFINI## AND PEN_NATURERUB="COT"'+
             ' ORDER BY PEN_PREDEFINI', false);
if (not Q.EOF) then
   begin
   Q.FIRST;
   LabPers:= THLabel (GetControl ('LBLPERSO'));
   if (LabPers<>nil) then
      begin
      LabPers.Caption:= 'Attention, cette cotisation est personnalisée';
      if (Q.FindField ('PEN_PREDEFINI').AsString='DOS') then
         LabPers.Caption:= LabPers.Caption+' dossier'
      else
         LabPers.Caption:= LabPers.Caption+' standard';
      end;
   SetControlVisible ('LBLPERSO', TRUE);
   MonMemo:= TMemo (GetControl ('MEMOPERSO'));
   if (MonMemo<>nil) then
      begin
      SetControlVisible ('MEMOPERSO', TRUE);
      MonMemo.Enabled:= TRUE;
      MonMemo.Clear;
      MonMemo.Lines.Add ('Les champs suivants ont été modifiés :');
      if (GetField ('PCT_ORGANISME')<>Q.FindField ('PEN_ORGANISME').AsString) then
         MonMemo.Lines.Add ('-Organisme');
      if (GetField ('PCT_LIBELLE')<>Q.FindField ('PEN_LIBELLE').AsString) then
         MonMemo.Lines.Add ('-Libellé');
      if (GetField ('PCT_IMPRIMABLE')<>Q.FindField ('PEN_IMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Rubrique imprimable');
      if (GetField ('PCT_BASEIMPRIMABLE')<>Q.FindField ('PEN_BASEIMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Base imprimable');
      if (GetField ('PCT_TXSALIMP')<>Q.FindField ('PEN_TXSALIMP').AsString) then
         MonMemo.Lines.Add ('-Taux salarial imprimable');
      if (GetField ('PCT_TXPATIMP')<>Q.FindField ('PEN_TXPATIMP').AsString) then
         MonMemo.Lines.Add ('-Taux patronal imprimable');
      if (GetField ('PCT_FFSALIMP')<>Q.FindField ('PEN_FFSALIMP').AsString) then
         MonMemo.Lines.Add ('-Forfait salarial imprimable');
      if (GetField ('PCT_FFPATIMP')<>Q.FindField ('PEN_FFPATIMP').AsString) then
         MonMemo.Lines.Add ('-Forfait patronal imprimable');
      if (GetField ('PCT_DECBASE')<>Q.FindField ('PEN_DECBASE').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales de la base');
      if (GetField ('PCT_DECTXSAL')<>Q.FindField ('PEN_DECTXSAL').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du taux salarial');
      if (GetField ('PCT_DECTXPAT')<>Q.FindField ('PEN_DECTXPAT').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du taux patronal');
      if (GetField ('PCT_DECMTSAL')<>Q.FindField ('PEN_DECMTSAL').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du montant salarial');
      if (GetField ('PCT_DECMTPAT')<>Q.FindField ('PEN_DECMTPAT').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du montant patronal');
      end
   else
      SetControlVisible ('MEMOPERSO', FALSE);
   end
else
   begin
   SetControlVisible ('LBLPERSO', FALSE);
   SetControlVisible ('MEMOPERSO', FALSE);
   end;
Ferme (Q);

//PT45
Q:= OpenSql ('SELECT *'+
             ' FROM PGEXCEPTIONS WHERE'+
             ' PEN_RUBRIQUE="'+GetField ('PCT_RUBRIQUE')+'" AND'+
             ' ##PEN_PREDEFINI## AND'+
             ' PEN_NATURERUB="COT"'+
             ' ORDER BY PEN_PREDEFINI', false);
if (not Q.EOF) then
   begin
   Personnalise:= True;
   Q.FIRST;
   LabPers:= THLabel (GetControl ('LBLPERSO'));
   if (LabPers <> nil) then
      begin
      LabPers.Caption:= 'Attention, cette cotisation est personnalisée';
      if (Q.FindField ('PEN_PREDEFINI').AsString='DOS') then
         LabPers.Caption:= LabPers.Caption+' dossier'
      else
         LabPers.Caption:= LabPers.Caption+' standard';
      end;
   SetControlVisible ('LBLPERSO', TRUE);
   MonMemo:= TMemo (GetControl ('MEMOPERSO'));
   if (MonMemo <> nil) then
      begin
      SetControlVisible ('MEMOPERSO', TRUE);
      MonMemo.Enabled:= TRUE;
      MonMemo.Clear;
      MonMemo.Lines.Add ('Les champs suivants ont été modifiés :');
      if (GetField ('PCT_ORGANISME')<>Q.FindField ('PEN_ORGANISME').AsString) then
         MonMemo.Lines.Add ('-Organisme');
      if (GetField ('PCT_LIBELLE')<>Q.FindField ('PEN_LIBELLE').AsString) then
         MonMemo.Lines.Add ('-Libellé');
      if (GetField ('PCT_IMPRIMABLE')<>Q.FindField ('PEN_IMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Rubrique imprimable');
      if (GetField ('PCT_BASEIMPRIMABLE')<>Q.FindField ('PEN_BASEIMPRIMABLE').AsString) then
         MonMemo.Lines.Add ('-Base imprimable');
      if (GetField ('PCT_TXSALIMP')<>Q.FindField ('PEN_TXSALIMP').AsString) then
         MonMemo.Lines.Add ('-Taux salarial imprimable');
      if (GetField ('PCT_TXPATIMP')<>Q.FindField ('PEN_TXPATIMP').AsString) then
         MonMemo.Lines.Add ('-Taux patronal imprimable');
      if (GetField ('PCT_FFSALIMP')<>Q.FindField ('PEN_FFSALIMP').AsString) then
         MonMemo.Lines.Add ('-Forfait salarial imprimable');
      if (GetField ('PCT_FFPATIMP')<>Q.FindField ('PEN_FFPATIMP').AsString) then
         MonMemo.Lines.Add ('-Forfait patronal imprimable');
      if (GetField ('PCT_DECBASE')<>Q.FindField ('PEN_DECBASE').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales de la base');
      if (GetField ('PCT_DECTXSAL')<>Q.FindField ('PEN_DECTXSAL').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du taux salarial');
      if (GetField ('PCT_DECTXPAT')<>Q.FindField ('PEN_DECTXPAT').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du taux patronal');
      if (GetField ('PCT_DECMTSAL')<>Q.FindField ('PEN_DECMTSAL').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du montant salarial');
      if (GetField ('PCT_DECMTPAT')<>Q.FindField ('PEN_DECMTPAT').AsInteger) then
         MonMemo.Lines.Add ('-Nombre décimales du montant patronal');
      end
   else
      SetControlVisible ('MEMOPERSO', FALSE);
   end
else
   begin
   SetControlVisible ('LBLPERSO', FALSE);
   SetControlVisible ('MEMOPERSO', FALSE);
   end;
Ferme (Q);
//FIN PT45
end;

//============================================================================================//
//                                 PROCEDURE Affecte Lookup Champ
//============================================================================================//

procedure TOM_Cotisation.AffecteLookupChamp(Champ, TypeChamp, libelle: string);
var
  vide, st: string;
  Q: TQuery;
{$IFNDEF EAGLCLIENT}
  ControlChamp: THDBEdit;
{$ELSE}
  ControlChamp: THEdit;
{$ENDIF}
  ControlLib: THLabel;
begin
  vide := '';
  if (TypeChamp = 'NBR') or (TypeChamp = 'ELP') or (TypeChamp = 'VAL') or (TypeChamp = 'NOM') or (TypeChamp = '') then //aucune tablette associé  (TypeChamp = 'BDC') or
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := vide;
      ControlChamp.ElipsisButton := False;
      ControlLib := THLabel(GetControl(libelle));
      if ControlLib <> nil then ControlLib.caption := '';
    end;
  end;

  if (TypeChamp = 'ELN') then //02
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      //  RendTabEltNationaux(ControlChamp);
      ControlChamp.DataType := 'PGELEMENTNAT';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
// DEB PT26
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
      begin
        St := 'SELECT PEL_LIBELLE,PEL_PREDEFINI FROM ELTNATIONAUX WHERE PEL_CODEELT="' +
          ControlChamp.Text + '" AND (PEL_PREDEFINI ="STD" OR (PEL_PREDEFINI ="DOS" AND PEL_NODOSSIER="' +
          PGRendNoDossier() + '")) ORDER BY PEL_PREDEFINI';
        Q := OpenSQL(st, TRUE);
        if not Q.EOF then
          ControlLib.caption := Q.FindField('PEL_PREDEFINI').asstring + ' ' + Q.FindField('PEL_LIBELLE').asstring
        else ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
        Ferme(Q);
// FIN ¨PT26
      end;
    end;
  end;

  if (TypeChamp = 'VAR') then //   03
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := 'PGVARIABLE';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
        ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
    end;
  end;

  if (TypeChamp = 'REM') then // 03 08
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := 'PGREMUNERATION';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
        ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
    end;
  end;

  if (TypeChamp = 'BDC') then //  08
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := 'PGBASECOTISATION';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
        ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
    end;
  end;

  //PT5
  if (TypeChamp = 'EXB') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := 'PGEXONERATIONBAS';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
        ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
    end;
  end;

  if (TypeChamp = 'EXC') then
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := 'PGEXONERATIONCOT';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
        ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
    end;
  end;
  //FIN PT5

  if (TypeChamp = 'ZCU') then // Recherche montant de cumul paie en cours
  begin
{$IFNDEF EAGLCLIENT}
    ControlChamp := THDBEdit(GetControl(champ));
{$ELSE}
    ControlChamp := THEdit(GetControl(champ));
{$ENDIF}
    if ControlChamp <> nil then
    begin
      ControlChamp.DataType := 'PGCUMULPAIE';
      ControlChamp.ElipsisButton := TRUE;
      ControlLib := THLabel(GetControl(libelle));
      if (ControlLib <> nil) and (ControlChamp.text <> '') then
        ControlLib.caption := RechDom(ControlChamp.datatype, ControlChamp.Text, FALSE);
    end;
  end;


end;
//============================================================================================//
//                                 PROCEDURE On New Record
//============================================================================================//

procedure TOM_Cotisation.OnNewRecord;
var
  vide: string;
begin
  inherited;

  if mode = 'DUPLICATION' then exit;
  vide := '';
  SetControlChecked('PCT_SOUMISREGUL', True);
  AffecteOnglet();
  SetField('PCT_DECBASE', '2');
  SetField('PCT_PREDEFINI', 'DOS');
  SetField('PCT_DECTXSAL', '2');
  SetField('PCT_DECTXPAT', '2');
  SetField('PCT_DECMTSAL', '2');
  SetField('PCT_DECMTPAT', '2');
  SetField('PCT_DECBASECOT', '2');
  // PT15 PH 04/06/2003 V_421 FQ 10689 gestion du champ presence du salarié en fin de mois
  SetField('PCT_PRESFINMOIS', '-');
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  SetField('PCT_NODOSSIER', PgRendNoDossier());
  SetControlEnabled('PCT_BASECOTISATION', False);
  SetControlEnabled('PCT_CODETRANCHE', False);
  SetControlEnabled('PCT_TAUXSAL', False);
  SetControlEnabled('PCT_FFSAL', False);
  SetControlChecked('PCT_FFSALIMP', False);
  SetControlEnabled('PCT_FFSALIMP', False);

  SetControlEnabled('PCT_TAUXPAT', False);
  SetControlEnabled('PCT_FFPAT', False);
  SetControlChecked('PCT_FFPATIMP', False);
  SetControlEnabled('PCT_FFPATIMP', False);
  SetControlEnabled('PCT_VALEURMAXI', False);
  SetControlEnabled('PCT_VALEURMINI', False);
  SetControlEnabled('PCT_AU', False);

  SetControlEnabled('PCT_BASECOTISATION_', False);
  SetControlEnabled('PCT_PLAFOND', False);
  SetControlEnabled('PCT_TRANCHE1', False);
  SetControlEnabled('PCT_TRANCHE2', False);
  SetControlEnabled('PCT_TRANCHE3', False);

end;

//============================================================================================//
//                                 PROCEDURE OnArgument
//============================================================================================//

procedure TOM_Cotisation.OnArgument(stArgument: string);
var
  Btn: TToolBarButton97;
  Q: TQuery;
  Perso: Boolean;
begin
  inherited;
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if (btn <> nil) then
    Btn.OnClick := DupliquerCotisation;
  if TCheckBox(GetControl('CBACCESFICHE')) <> nil then
    TCheckBox(GetControl('CBACCESFICHE')).OnClick := CBAccesFicheClick;
  Btn := TToolBarButton97(GetControl('BIMPRIMER'));
  if (btn <> nil) then
  begin
    SetControlvisible('BIMPRIMER', TRUE);
    Btn.OnClick := ImprimerClick;
  end;

  Perso := False;
  Q := OpenSql('SELECT US_CONTROLEUR' +
    ' FROM UTILISAT WHERE' +
    ' US_UTILISATEUR="' + V_PGI.FUser + '"', True);
  if (not Q.EOF) then
  begin
    if Q.FindField('US_CONTROLEUR').AsString = 'X' then
      Perso := True;
  end;
  Ferme(Q);
  SetControlEnabled('BTNPERSO', Perso);

//PT45
  Btn := TToolBarButton97(GetControl('BGESTIONSPEC'));
  if (Btn <> nil) then
    Btn.OnClick := GestionSpec_OnClick;
//FIN PT45
// d PT47
  Btn := TToolBarButton97(GetControl('BDUCS'));
  if (Btn <> nil) then
    Btn.OnClick := AffectDucs_OnClick;
// f PT47
  PaieConceptPlanPaie(Ecran); // PT46
end;
//============================================================================================//
//                                 PROCEDURE DupliquerCotisation
//============================================================================================//

procedure TOM_Cotisation.DupliquerCotisation(Sender: TObject);
var
  T, T_Cumul, TCumulSpec, T_Ducs, T_Ventil, TOB_GestAssoc, TOB_GestionSpec: TOB;
  i: integer;
  {AncValNat, AncValRub , ChampBug, ChampBug1, ChampBug2, ChampBug3: string;}//PT53
  Cumulpaie, NoDossier, OriginePred, PredExist, ST, temp, vide: string;
  Champ: array[1..3] of Hstring;
  Valeur: array[1..3] of variant;
  Ok: Boolean;
  Q: TQuery;
  TDuplic : TOB;
begin
  TFFiche(Ecran).BValider.Click;
  mode := 'DUPLICATION';
  vide := '';
{debut PT53
  ChampBug := GetField('PCT_THEMECOT');
  ChampBug1 := GetField('PCT_TYPEBASE');
  ChampBug2 := GetField('PCT_CODETRANCHE');
  ChampBug3 := GetField('PCT_IMPRIMABLE');
  AncValNat := GetField('PCT_NATURERUB');
  AncValRub := GetField('PCT_RUBRIQUE');}
  
  TDuplic:= TOB.Create('MaTob initiale', nil, -1);
  TDuplic.AddChampSupValeur ('PCT_THEMECOT', GetField('PCT_THEMECOT'));
  TDuplic.AddChampSupValeur ('PCT_TYPEBASE', GetField('PCT_TYPEBASE'));
  TDuplic.AddChampSupValeur ('PCT_CODETRANCHE', GetField('PCT_CODETRANCHE'));
  TDuplic.AddChampSupValeur ('PCT_IMPRIMABLE', GetField('PCT_IMPRIMABLE'));
  TDuplic.AddChampSupValeur ('PCT_NATURERUB', GetField('PCT_NATURERUB'));
  TDuplic.AddChampSupValeur ('PCT_RUBRIQUE', GetField('PCT_RUBRIQUE'));
  TDuplic.AddChampSupValeur ('PCT_TYPETAUXSAL', GetField('PCT_TYPETAUXSAL'));
  TDuplic.AddChampSupValeur ('PCT_TYPETAUXPAT', GetField('PCT_TYPETAUXPAT'));
  TDuplic.AddChampSupValeur ('PCT_TYPEFFSAL', GetField('PCT_TYPEFFSAL'));
  TDuplic.AddChampSupValeur ('PCT_TYPEFFPAT', GetField('PCT_TYPEFFPAT'));
  TDuplic.AddChampSupValeur ('PCT_TYPEMINISAL', GetField('PCT_TYPEMINISAL'));
  TDuplic.AddChampSupValeur ('PCT_TYPEMAXISAL', GetField('PCT_TYPEMAXISAL'));
  TDuplic.AddChampSupValeur ('PCT_TYPEMINIPAT', GetField('PCT_TYPEMINIPAT'));
  TDuplic.AddChampSupValeur ('PCT_TYPEMAXIPAT', GetField('PCT_TYPEMAXIPAT'));
//FIN PT53
  temp := GetField('PCT_BLOCNOTE');
  AglLanceFiche('PAY', 'CODE', '', '', 'COT;' + TDuplic.GetValue ('PCT_RUBRIQUE') + '; ;4');
  // DEB PT46
  if (PGCodePredefini = 'DOS') and (not DOS) then
  begin
    PgiInfo('Vous n''êtes pas autorisé à créer une cotisation de type dossier.', Ecran.Caption);
    exit;
  end;
  // FIN PT46
  if (PGCodeDupliquer <> '') then
  begin
    Champ[1] := 'PCT_PREDEFINI';
    Valeur[1] := PGCodePredefini;
    Champ[2] := 'PCT_NODOSSIER';
    if (PGCodePredefini = 'DOS') then
      Valeur[2] := PgRendNoDossier()
    else
      Valeur[2] := '000000';
    Champ[3] := 'PCT_RUBRIQUE';
    Valeur[3] := PGCodeDupliquer;
    Ok := RechEnrAssocier('COTISATION', Champ, Valeur);
    if (Ok = False) then //Test si code existe ou non
    begin
      TOB_GestAssoc := TOB.Create('La Cotisation originale', nil, -1);
      st := 'SELECT *' +
        ' FROM CUMULRUBRIQUE WHERE' +
        ' PCR_NATURERUB="' + TDuplic.GetValue ('PCT_NATURERUB') + '" AND' +          //PT53
        ' ##PCR_PREDEFINI## PCR_RUBRIQUE="' + TDuplic.GetValue ('PCT_RUBRIQUE') + '"';//PT53
      Q := OpenSql(st, TRUE);
      TOB_GestAssoc.LoadDetailDB('CUMULRUBRIQUE', '', '', Q, FALSE);
      ferme(Q);
      T_Cumul := TOB.Create('La Cotisation dupliquée', nil, -1);
//PT45
      TOB_GestionSpec := TOB.Create('La gestion spéc orig', nil, -1);
      st := 'SELECT *' +
        ' FROM CUMULRUBDOSSIER WHERE' +
        ' PKC_NATURERUB="' + TDuplic.GetValue ('PCT_NATURERUB') + '" AND' +   //PT53
        ' PKC_RUBRIQUE="' + TDuplic.GetValue ('PCT_RUBRIQUE') + '"'; //PT53
      Q := OpenSql(st, TRUE);
      TOB_GestionSpec.LoadDetailDB('CUMULRUBDOSSIER', '', '', Q, FALSE);
      FERME(Q);
      TCumulSpec := TOB.Create('La gestion spéc dest', nil, -1);
//FIN PT45
      DupliquerPaie(TFFiche(Ecran).TableName, TFFiche(Ecran));
      SetField('PCT_RUBRIQUE', PGCodeDupliquer);
//Refonte pour intégration de la duplication des éléments associés
//Duplication de la gestion associée
      try
        BeginTrans;
        for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
          if ((TOB_GestAssoc.Detail[i].GetValue('PCR_NATURERUB')) = TDuplic.GetValue ('PCT_NATURERUB')) and //PT53
            ((TOB_GestAssoc.Detail[i].GetValue('PCR_RUBRIQUE')) = TDuplic.GetValue ('PCT_RUBRIQUE')) then  //PT53
          begin
            T := TOB_GestAssoc.Detail[i];
            if (T <> nil) then
            begin
              T.PutValue('PCR_RUBRIQUE', PGCodeDupliquer);
              T.PutValue('PCR_PREDEFINI', PGCodePredefini);
              Cumulpaie := T.GetValue('PCR_CUMULPAIE');
              OriginePred := '';
              if (Cumulpaie <> '') and (Length(Cumulpaie) = 2) then
              begin
                if (IsNumeric(CumulPaie)) then
                  if (StrToInt(CumulPaie) > 50) and ((Cumulpaie[2] = '5') or
                    (Cumulpaie[2] = '7') or (Cumulpaie[2] = '9')) then
                    OriginePred := 'DOS';
              end;
              if OriginePred = 'DOS' then
              begin
                T.PutValue('PCR_PREDEFINI', 'DOS');
                T.PutValue('PCR_NODOSSIER', PgRendNoDossier());
              end
              else
                T.PutValue('PCR_NODOSSIER', '000000');
              if (PGCodePredefini = 'DOS') then
                T.PutValue('PCR_NODOSSIER', PgRendNoDossier());
            end;
          end;
        T_Cumul.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
        T_Cumul.InsertDB(nil, False);
        TOB_GestAssoc.free;
        T_Cumul.free;
//PT45
//DEB PT52
        if (PGCodePredefini <> 'DOS') then
        begin
            for i := 0 to TOB_GestionSpec.Detail.Count - 1 do
            begin
              if ((TOB_GestionSpec.Detail[i].GetValue('PKC_NATURERUB')) = TDuplic.GetValue ('PCT_NATURERUB')) and //PT53
                ((TOB_GestionSpec.Detail[i].GetValue('PKC_RUBRIQUE')) = TDuplic.GetValue ('PCT_RUBRIQUE')) then //PT53
              begin
                T := TOB_GestionSpec.Detail[i];
                if (T <> nil) then
                begin
                  T.PutValue('PKC_RUBRIQUE', PGCodeDupliquer);
                  Cumulpaie := T.GetValue('PKC_CUMULPAIE');
                end;
              end;
            end;
            TCumulSpec.Dupliquer(TOB_GestionSpec, TRUE, TRUE, FALSE);
            TCumulSpec.InsertDB(nil, False);
        end;
        //else
        //   PGIBox('Personnalisation non dupliquée');
//FIN PT52
        FreeAndNil(TOB_GestionSpec);
        FreeAndNil(TCumulSpec);
//FIN PT45

        if (TDuplic.GetValue ('PCT_NATURERUB') = 'COT') then //PT53
        begin
//Duplication de la Ventilation analytique
          PredExist := '';
          st := 'SELECT *' +
            ' FROM VENTICOTPAIE WHERE' +
            ' ##PVT_PREDEFINI## PVT_RUBRIQUE="' + TDuplic.GetValue ('PCT_RUBRIQUE') + '"'; //PT53
          Q := OpenSql(st, TRUE);
          if (not Q.eof) then
          begin
            TOB_GestAssoc := TOB.Create('La ventilation comptable', nil, -1);
            TOB_GestAssoc.LoadDetailDB('VENTICOTPAIE', '', '', Q, FALSE);
            ferme(Q);
//Suppression des élements comptables non conservés pour ne pas créer de doublon
            if (TOB_GestAssoc.FindFirst(['PVT_PREDEFINI'], ['CEG'], False) <> nil) then
              PredExist := 'CEG';
            if (TOB_GestAssoc.FindFirst(['PVT_PREDEFINI'], ['STD'], False) <> nil) then
              PredExist := PredExist + ';STD';
            if (TOB_GestAssoc.FindFirst(['PVT_PREDEFINI'], ['DOS'], False) <> nil) then
              PredExist := PredExist + ';DOS';
            T := TOB_GestAssoc.FindFirst([''], [''], False);
            while T <> nil do
            begin
              if (PGCodePredefini <> 'CEG') and
                (T.GetValue('PVT_PREDEFINI') <> PGCodePredefini) and
                (Pos(PGCodePredefini, PredExist) > 0) then
                T.Free;
              if (Pos('DOS', PredExist) = 0) and
                (PGCodePredefini = 'DOS') and
                (Pos('CEG', PredExist) > 0) and
                (Pos('STD', PredExist) > 0) then
                if (T.GetValue('PVT_PREDEFINI') = 'CEG') then
                  T.Free;
              T := TOB_GestAssoc.FindNext([''], [''], False);
            end;
//Duplication des éléments restants
            T_Ventil := TOB.Create('VENTICOTPAIE', nil, -1);
            for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
            begin
              T := TOB_GestAssoc.Detail[i];
              T.PutValue('PVT_RUBRIQUE', PGCodeDupliquer);
              T.PutValue('PVT_PREDEFINI', PGCodePredefini);
              if (PGCodePredefini = 'DOS') then
                T.PutValue('PVT_NODOSSIER', PgRendNoDossier)
              else
                T.PutValue('PVT_NODOSSIER', '000000');
            end;
            T_Ventil.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
            if (T_Ventil.detail.count > 0) then
              T_Ventil.InsertDB(nil, False);
            TOB_GestAssoc.free;
            T_Ventil.free;
          end
          else
            Ferme(Q);
//Duplication de la codification DUCS
          PredExist := '';
          st := 'SELECT *' +
            ' FROM DUCSAFFECT WHERE' +
            ' ##PDF_PREDEFINI## PDF_RUBRIQUE="' + TDuplic.GetValue ('PCT_RUBRIQUE') + '"'; //PT53
          Q := OpenSql(st, TRUE);
          if (not Q.eof) then
          begin
            TOB_GestAssoc := TOB.Create('La DUCS', nil, -1);
            TOB_GestAssoc.LoadDetailDB('DUCSAFFECT', '', '', Q, FALSE);
            ferme(Q);
//Suppression des élements ducs non conservés pour ne pas créer de doublon
            if (TOB_GestAssoc.FindFirst(['PDF_PREDEFINI'], ['CEG'], False) <> nil) then
              PredExist := 'CEG';
            if (TOB_GestAssoc.FindFirst(['PDF_PREDEFINI'], ['STD'], False) <> nil) then
              PredExist := PredExist + ';STD';
            if (TOB_GestAssoc.FindFirst(['PDF_PREDEFINI'], ['DOS'], False) <> nil) then
              PredExist := PredExist + ';DOS';
            T := TOB_GestAssoc.FindFirst([''], [''], False);
            while (T <> nil) do
            begin
              if (PGCodePredefini <> 'CEG') and
                (T.GetValue('PDF_PREDEFINI') <> PGCodePredefini) and
                (Pos(PGCodePredefini, PredExist) > 0) then
                T.Free;
              if (Pos('DOS', PredExist) = 0) and
                (PGCodePredefini = 'DOS') and
                (Pos('CEG', PredExist) > 0) and
                (Pos('STD', PredExist) > 0) then
                if (T.GetValue('PDF_PREDEFINI') = 'CEG') then
                  T.Free;
              T := TOB_GestAssoc.FindNext([''], [''], False);
            end;
//Duplication des éléments restants
            T_DUCS := TOB.Create('La DUCS dupliquée', nil, -1);
            for i := 0 to TOB_GestAssoc.Detail.Count - 1 do
            begin
              T := TOB_GestAssoc.Detail[i];
              T.PutValue('PDF_RUBRIQUE', PGCodeDupliquer);
              T.PutValue('PDF_PREDEFINI', PGCodePredefini);
              if (PGCodePredefini = 'DOS') then
                T.PutValue('PDF_NODOSSIER', PgRendNoDossier)
              else
                T.PutValue('PDF_NODOSSIER', '000000');
            end;
            T_DUCS.Dupliquer(Tob_GestAssoc, TRUE, TRUE, FALSE);
            T_DUCS.InsertDB(nil, False);
            TOB_GestAssoc.free;
            T_DUCS.free;
          end
          else
            Ferme(Q);
        end;
        CommitTrans;
      except
        Rollback;
        PGIBox('Une erreur est survenue lors de la duplication des éléments' +
          ' associés à la rubrique.', Ecran.caption);
      end;

{debut PT53
      SetField('PCT_THEMECOT', ChampBug);
      SetField('PCT_TYPEBASE', ChampBug1);
      SetField('PCT_CODETRANCHE', ChampBug2);
      SetField('PCT_IMPRIMABLE', ChampBug3);
      SetField('PCT_NATURERUB', AncValNat);
      SetField('PCT_PREDEFINI', PGCodePredefini);
}
      SetField('PCT_PREDEFINI', PGCodePredefini);
      SetField('PCT_THEMECOT', TDuplic.GetValue ('PCT_THEMECOT'));
      SetField('PCT_TYPEBASE', TDuplic.GetValue ('PCT_TYPEBASE'));
      SetField('PCT_CODETRANCHE', TDuplic.GetValue ('PCT_CODETRANCHE'));
      SetField('PCT_IMPRIMABLE', TDuplic.GetValue ('PCT_IMPRIMABLE'));
      SetField('PCT_NATURERUB', TDuplic.GetValue ('PCT_NATURERUB'));
      SetField('PCT_TYPETAUXSAL', TDuplic.GetValue ('PCT_TYPETAUXSAL'));
      SetField('PCT_TYPETAUXPAT', TDuplic.GetValue ('PCT_TYPETAUXPAT'));
      SetField('PCT_TYPEFFSAL', TDuplic.GetValue ('PCT_TYPEFFSAL'));
      SetField('PCT_TYPEFFPAT', TDuplic.GetValue ('PCT_TYPEFFPAT'));
      SetField('PCT_TYPEMINISAL', TDuplic.GetValue ('PCT_TYPEMINISAL'));
      SetField('PCT_TYPEMAXISAL', TDuplic.GetValue ('PCT_TYPEMAXISAL'));
      SetField('PCT_TYPEMINIPAT', TDuplic.GetValue ('PCT_TYPEMINIPAT'));
      SetField('PCT_TYPEMAXIPAT', TDuplic.GetValue ('PCT_TYPEMAXIPAT'));
//Fin PT53
      SetControlEnabled('PCT_RUBRIQUE', False);
      AccesFicheDupliquer(TFFiche(Ecran), PGCodePredefini, NoDossier, LectureSeule);
      SetField('PCT_NODOSSIER', NoDossier);
      SetControlEnabled('BInsert', True);
      SetControlEnabled('PCT_PREDEFINI', False);
      SetControlEnabled('PCT_RUBRIQUE', False);
      SetControlEnabled('BDUPLIQUER', True);
      PgUpdateDateAgl(GetField('PCT_RUBRIQUE'));
      if Assigned(Trace) then FreeAndNil(Trace);
      Trace := TStringList.Create; //PT44
      st := 'Duplication de la rubrique '+TDuplic.GetValue ('PCT_RUBRIQUE');
      Trace.add (st);
      st := 'Création de la rubrique '+ GetField('PCT_RUBRIQUE');
      Trace.add (st);
      EnDupl := 'OUI';
      IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran)); //PT44
      FreeAndNil(Trace); //PT44
      EnDupl := 'NON';
      TFFiche(Ecran).Bouge(nbPost); //Force enregistrement
      PgiInfo('Duplication effectuée. Seule la ventilation analytique n''est' +
        ' pas reprise.', Ecran.caption);
      FreeandNil(TDuplic); //PT53
      EnDupl := '';
      AffecteOnglet();
      if (GetField('PCT_NATURERUB') = 'COT') then  
      begin
        SetControlEnabled('PCT_ORGANISME', True);
        SetControlEnabled('PCT_SOUMISMALAD', False);
        SetControlChecked('PCT_SOUMISMALAD', False);
        SetControlEnabled('PCT_ASSIETTEBRUT', False);
        SetControlChecked('PCT_ASSIETTEBRUT', False);
        SetControlEnabled('PCT_PRECOMPTEASS', True);
        SetControlEnabled('PCT_ETUDEDROIT', True);
        SetControlVisible('PCT_PRESFINMOIS', TRUE);
      end;
      if GetField('PCT_NATURERUB') = 'BAS' then
      begin
        SetControlEnabled('PCT_ORGANISME', False);
        if GetField('PCT_ORGANISME') <> vide then SetField('PCT_ORGANISME', vide); //PT29
        SetControlEnabled('PCT_SOUMISMALAD', True);
        SetControlEnabled('PCT_ASSIETTEBRUT', True);
        SetControlEnabled('PCT_PRECOMPTEASS', False);
        SetControlChecked('PCT_PRECOMPTEASS', False);
        SetControlEnabled('PCT_ETUDEDROIT', False);
        SetControlChecked('PCT_ETUDEDROIT', False);
        SetControlVisible('PCT_PRESFINMOIS', FALSE);
      end;
    end
    else
      HShowMessage('5;Cotisation :;La duplication est impossible, la rubrique' +
        ' existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;


function TOM_Cotisation.ControlChampValeur(nature: string): Boolean; // PT33 Procedure en function
var
  Mes: string;
  OkError: Boolean;
begin
  OkError := False;
  Mes := 'Vous n''avez pas renseigné : '; { PT37 }
  if GetField('PCT_NATURERUB') = 'BAS' then { DEB PT39 }
  begin
    if (GetControlText('PCT_TYPEBASE_') <> '') and (GetControlText('PCT_BASECOTISATION_') = '') and (GetControlText('PCT_TYPEBASE_') <> 'ELV') then //PT12
    begin
      Mes := mes + '#13#10- la base de cotisation';
      OkError := true;
    end;
  end { FIN PT39 }
  else
    if (GetField('PCT_TYPEBASE') <> '') and (GetField('PCT_BASECOTISATION') = '') and (GetField('PCT_TYPEBASE') <> 'ELV') then //PT12
    begin
      Mes := mes + '#13#10- la base de cotisation';
      OkError := true;
    end;

  if Nature = 'COT' then
  begin
    // PT9 Ph 04/04/2002 V571 On saisit un element variable en base,taux et forfaits
    if (GetField('PCT_TYPETAUXSAL') <> '') and (GetField('PCT_TYPETAUXSAL') <> 'ELV') { PT37 }
      and (GetField('PCT_TAUXSAL') = '') and (GetField('PCT_TYPEFFSAL') <> 'ELV') then //PT12
      if (GetField('PCT_TYPEFFSAL') = '') and (GetField('PCT_TYPEFFPAT') = '') then { PT18-1 }
      begin
        Mes := mes + '#13#10- le taux salarial';
        OkError := true;
      end;
    if (GetField('PCT_TYPETAUXPAT') <> '') and (GetField('PCT_TYPETAUXPAT') <> 'ELV') { PT37 }
      and (GetField('PCT_TAUXPAT') = '') and (GetField('PCT_TYPEFFPAT') <> 'ELV') then //PT12
      if (GetField('PCT_TYPEFFSAL') = '') and (GetField('PCT_TYPEFFPAT') = '') then { PT18-1 }
      begin
        Mes := mes + '#13#10- le taux patronal';
        OkError := true;
      end;
    if (GetField('PCT_TYPEFFSAL') <> '') and (GetField('PCT_TYPEFFSAL') <> 'ELV') { PT37 }
      and (GetField('PCT_FFSAL') = '') and (GetField('PCT_TYPETAUXSAL') <> 'ELV') then //PT12
      if (GetField('PCT_TYPETAUXSAL') = '') and (GetField('PCT_TYPETAUXPAT') = '') then { PT18-1 }
      begin
        Mes := mes + '#13#10- le forfait salarial';
        OkError := true;
      end;
    if (GetField('PCT_TYPEFFPAT') <> '') and (GetField('PCT_TYPEFFPAT') <> 'ELV') and { PT37 }
      (GetField('PCT_FFPAT') = '') and (GetField('PCT_TYPETAUXPAT') <> 'ELV') then //PT12
      if (GetField('PCT_TYPETAUXSAL') = '') and (GetField('PCT_TYPETAUXPAT') = '') then { PT18-1 }
      begin
        Mes := mes + '#13#10- le forfait patronal';
        OkError := true;
      end;
    // FIN PT9
    if (GetField('PCT_TYPEMINISAL') <> '') and (GetField('PCT_VALEURMINISAL') = '') then
    begin
      Mes := mes + '#13#10- le minimum salarial';
      OkError := true;
    end;
    if (GetField('PCT_TYPEMAXISAL') <> '') and (GetField('PCT_VALEURMAXISAL') = '') then
    begin
      Mes := mes + '#13#10- le maximum salarial';
      OkError := true;
    end;
    if (GetField('PCT_TYPEMINIPAT') <> '') and (GetField('PCT_VALEURMINIPAT') = '') then
    begin
      Mes := mes + '#13#10- le minimum patronal';
      OkError := true;
    end;
    if (GetField('PCT_TYPEMAXIPAT') <> '') and (GetField('PCT_VALEURMAXIPAT') = '') then
    begin
      Mes := mes + '#13#10- le maximum patronal';
      OkError := true;
    end;
  end
  else
    if Nature = 'BAS' then
    begin
      if (GetField('PCT_TYPEPLAFOND') <> '') and (GetField('PCT_PLAFOND') = '') then
      begin
        Mes := mes + '#13#10- le plafond';
        OkError := true;
      end;
      if (GetField('PCT_TYPETRANCHE1') <> '') and (GetField('PCT_TRANCHE1') = '') then
      begin
        Mes := mes + '#13#10- la tranche T1';
        OkError := true;
      end;
      if (GetField('PCT_TYPETRANCHE2') <> '') and (GetField('PCT_TRANCHE2') = '') then
      begin
        Mes := mes + '#13#10- la tranche T2';
        OkError := true;
      end;
      if (GetField('PCT_TYPETRANCHE3') <> '') and (GetField('PCT_TRANCHE3') = '') then
      begin
        Mes := mes + '#13#10- la tranche T3';
        OkError := true;
      end;
    end;
  if OkError = True then
    PgiBox(Mes + '.#13#10 Attention, le calcul de la rubrique sur le bulletin de paie peut-être erroné!', Ecran.caption); { PT37 }
  Result := OKError; // PT33
end;

procedure TOM_Cotisation.OnAfterUpdateRecord;
var StSql: string;
  Tob_Prof: Tob;
  i: integer;
  Q: TQuery;
  Even: Boolean; //PT44
begin
  inherited;
  //Rechargement des tablettes
//PT25  ChargementTablette(TFFiche(Ecran).TableName, '');
  SetControlEnabled('BInsert', True);
  SetControlEnabled('BDelete', True);
  SetControlEnabled('PCT_PREDEFINI', False);
  SetControlEnabled('PCT_RUBRIQUE', False);
  SetControlEnabled('BDUPLIQUER', True);
  SetControlEnabled('BGESTIONASS', True);
  SetControlEnabled('BDUCS', (GetField('PCT_NATURERUB') = 'COT')); //PT17-2
  SetControlEnabled('BVENTIL', True);
  { DEB PT28 }
  if (OrdreEtat <> GetField('PCT_ORDREETAT')) and (OrdreEtat <> -1) then
  begin
    InitMoveProgressForm(nil, 'Mise à jour en cours.', 'Veuillez patienter..', 2, False, True);
    MoveCurProgressForm;
    ExecuteSql('UPDATE HISTOBULLETIN SET PHB_OMTSALARIAL=' + IntToStr(GetField('PCT_ORDREETAT')) + ' ' + { PT31 }
      'WHERE PHB_NATURERUB="' + GetField('PCT_NATURERUB') + '" ' +
      'AND (PHB_RUBRIQUE="' + GetField('PCT_RUBRIQUE') + '" OR ' +
      'SUBSTRING(PHB_RUBRIQUE,1,4)="' + GetField('PCT_RUBRIQUE') + '" )'); { SB 20/10/2005 }
    MoveCurProgressForm;
    ExecuteSql('UPDATE HISTOANALPAIE SET PHA_OMTSALARIAL=' + IntToStr(GetField('PCT_ORDREETAT')) + ' ' + { PT31 }
      'WHERE PHA_NATURERUB="' + GetField('PCT_NATURERUB') + '" ' +
      'AND PHA_RUBRIQUE="' + GetField('PCT_RUBRIQUE') + '" ');
    FiniMoveProgressForm;
    OrdreEtat := GetField('PCT_ORDREETAT');
  end;
  { FIN PT28 }

  { DEB PT42 Refonte du executeSQL }
  { DEB PT36 }
  StSql := 'SELECT * FROM PROFILRUB ' +
    'WHERE ##PPM_PREDEFINI## PPM_NATURERUB="' + GetField('PCT_NATURERUB') + '" ' +
    'AND PPM_RUBRIQUE="' + GetField('PCT_RUBRIQUE') + '" ';
  try
    Q := OpenSql(StSql, True);
    if not Q.eof then
    begin
      Tob_Prof := TOB.Create('Les profils', nil, -1);
      Tob_Prof.LoadDetailDB('PROFILRUB', '', '', Q, FAlse);
    end;
    Ferme(Q);
    if Assigned(Tob_Prof) and (Tob_Prof.detail.count > 0) then
    begin
      for i := 0 to Tob_Prof.detail.count - 1 do
        Tob_Prof.detail[i].PutValue('PPM_LIBELLE', GetField('PCT_LIBELLE'));
      Tob_Prof.SetAllModifie(True);
      Tob_Prof.UpdateDB;
    end;
  finally
    FreeAndNil(Tob_Prof);
  end;
  { FIN PT36 }
  { FIN PT42 }
  if Assigned(Trace) then FreeAndNil(Trace);
  Trace := TStringList.Create; //PT44
  even := IsDifferent(dernierecreate, PrefixeToTable(TFFiche(Ecran).TableName), TFFiche(Ecran).CodeName, TFFiche(Ecran).LibelleName, Trace, TFFiche(Ecran), LeStatut); //PT44  //PT49
  FreeAndNil(Trace); //PT44
  if OnFerme then Ecran.Close;
end;

procedure TOM_Cotisation.OnClose;
begin
  inherited;
  { DEB PT16 }
  if GetControlText('CBACCESFICHE') = 'X' then
  begin
    LastError := 1;
    LastErrorMsg := '';
    PgiBox('Suite à la modification de la gestion associée, vous devez valider votre fiche.', Ecran.caption);
  end;
  { FIN PT16 }
end;

procedure TOM_Cotisation.CBAccesFicheClick(Sender: TObject);
begin
  { DEB PT16 disabled si acces fiche gestion associée }
  SetControlEnabled('BFirst', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BPrev', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BNext', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BLast', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('bInsert', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('bDefaire', (GetControlText('CBACCESFICHE') = '-'));
  SetControlEnabled('BDUPLIQUER', (GetControlText('CBACCESFICHE') = '-'));
  { FIN PT16 }
{$IFDEF EAGLCLIENT}
  //SetField ('PCT_ABREGE', GetField ('PCT_ABREGE')+ ' ');
  DS.Edit;
{$ENDIF}
end;


{ DEB PT19 }

procedure TOM_Cotisation.ImprimerClick(Sender: TObject);
begin
  LanceEtat('E', 'PAY', 'PCT', True, False, False, nil, 'AND PCT_RUBRIQUE="' + GetField('PCT_RUBRIQUE') + '"', '', False);
end;
{ FIN PT19 }

{ PT21 Mise à jour date modif table associée }

procedure TOM_Cotisation.PgUpdateDateAgl(Rub: string);
begin
  SetControlChecked('CBACCESFICHE', False);
  ExecuteSql('UPDATE CUMULRUBRIQUE SET' +
    ' PCR_DATEMODIF="' + UsTime(Now) + '" WHERE' +
    ' ##PCR_PREDEFINI## PCR_RUBRIQUE="' + Rub + '"');
//PT45
  ExecuteSql('UPDATE CUMULRUBDOSSIER SET' +
    ' PKC_DATEMODIF="' + UsTime(Now) + '" WHERE' +
    ' PKC_RUBRIQUE="' + Rub + '"');
//FIN PT45
end;

//PT45
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 16/04/2007
Modifié le ... :   /  /
Description .. : Click sur le bouton de la gestion spécifique
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_Cotisation.GestionSpec_OnClick(Sender: TObject);
var
  Quoi: string;
begin
  if (GetField('PCT_PREDEFINI') = 'CEG') then
    Quoi := 'C'
  else
    if (GetField('PCT_PREDEFINI') = 'STD') then
      Quoi := 'S'
    else
      Quoi := 'D';
  AglLanceFiche('PAY', 'CUMULGESTIONSPEC', '', '', GetField('PCT_NATURERUB') +
    ';' + GetField('PCT_RUBRIQUE') + ';' + GetField('PCT_LIBELLE') + ';' +
    Quoi);
end;
//FIN PT45

// d PT47
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Monique FAUDEL
Créé le ...... : 16/04/2007
Modifié le ... : 28/06/2007
Description .. : Click sur le bouton des affectation aux codes DUCS
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_Cotisation.AffectDucs_OnClick(Sender: TObject);
var
  Q                                 : TQuery;
  Predefini, Nodossier, Rubrique,st : string;
begin

  Predefini :=  GetField('PCT_PREDEFINI');
  Nodossier :=  GetField('PCT_NODOSSIER');
  Rubrique  :=  GetField('PCT_RUBRIQUE');
  st := 'ACTION=CONSULTATION';

  Q := OpenSql('SELECT PDF_PREDEFINI,PDF_NODOSSIER,PDF_RUBRIQUE' +
                 ' FROM DUCSAFFECT ' +
                 'WHERE PDF_PREDEFINI = "'+Predefini+'" '+
                 'AND PDF_NODOSSIER = "'+Nodossier+'" '+
                 'AND PDF_RUBRIQUE = "'+Rubrique+ '"', True);
  if  Q.eof then
  begin
    st := 'ACTION=CREATION';
    ferme(Q);
    Q := OpenSql('SELECT PDF_PREDEFINI,PDF_NODOSSIER,PDF_RUBRIQUE' +
                 ' FROM DUCSAFFECT ' +
                 'WHERE ##PDF_PREDEFINI## '+
                 'AND PDF_RUBRIQUE = "'+Rubrique+ '"', True);
    if  not Q.eof then
    begin
      Predefini :=  Q.FindField('PDF_PREDEFINI').AsString;
      Nodossier :=  Q.FindField('PDF_NODOSSIER').AsString;
      Rubrique  :=  Q.FindField('PDF_RUBRIQUE').AsString;
      st := 'ACTION=CONSULTATION';
    end;
  end;
  Ferme(Q);
  AglLanceFiche('PAY','DUCSAFFECT','',Predefini +';'+Nodossier+';'+Rubrique,
                                      Predefini +';'+Nodossier+';'+Rubrique+';'+st);
end;
// f PT47

initialization
  registerclasses([TOM_COTISATION]);

end.

