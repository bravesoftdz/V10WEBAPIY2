{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 18/02/2000
Modifié le ... : 30/05/2001
Description .. : Multi critère de lancement, de sélection des salariés
Suite ........ : en vue de faire une liste des salariés ayant des paies
Suite ........ :
Suite ........ : Confectionne une clause XX_WHERE
Suite ........ : Jointure sur la table des Paies pour faire apparaitre les
Suite ........ : paies faites et celle à faire dans le mois sélectioné
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1   : 10/09/2001 PH V547 Controle des dates avant lancement de la paie
PT2   : 22/10/2001 PH V562 Suppression des tests entree,sortie,suspension de
                           paie sur les salaries pour la liste des paies faites
PT3   : 22/10/2001 PH V562 Rajout un paramètre à l'appel de la saisie du
                           bulletin
PT4   : 26/03/2002 PH V571 Modif message alerte
PT5   : 15/05/2002 PH V575 Message alerte si liste vide et dble click
PT6   : 15/05/2002 VG V582 IFDEF CCS3, limitation à 30 bulletins par période
                          (date à date)
PT7   : 01/07/2002 VG V585 PT6, finalement non !!!
PT8   : 22/10/2002 VG V585 Version S3 - FQ N° 10290
PT9   : 05/11/2002 PH V591 Changement de la liste de la query du mul si besoin
PT10  : 26/11/2002 PH V591 Acces à la fiche salarie dans le cas de la liste des
                           paies à effectuer
PT11  : 13/05/2003 PH V_42 Optimisation chargement des tob de la paie non fait
                           systématiquement
PT12  : 03/07/2003 PH V_42 Suppression automatisée de bulletin
PT13  : 16/10/2003 FQ V_42 10874 PH Test si bulletin cloturé avt suppression
PT14  : 03/11/2003 FQ V_42 10839 Code salarié sur 10 caractères
PT15  : 23/04/2004 VG V_50 Version S3 - FQ N° 11129
PT16  : 08/06/2004 PH V_50 FQ N° 11068 Controle bulletin postérieur avant
                           suppression
PT17  : 06/07/2004 PH V_50 FQ N° 11117 Message de modification de bulletin si
                           ecriture générée paie non clôturée
PT18  : 19/08/2004 PH V_50 FQ N° 11490 Message de suppression impossible de
                           bulletin
PT19  : 02/09/2004 PH V_50 FQ N° 11560 Visibilité des champs onglets compléménts
PT20  : 03/09/2004 PH V_50 FQ N° 10949 Message d'alerte de suppression de
                           bulletin comptablilisé
PT21  : 13/10/2004 VG V_50 Message lors de la suppression si aucun élément n'est
                           sélectionné - FQ N°11609
PT22  : 03/06/2005 PH V_60 FQ N° 11437 Cloture decloture individuelle
PT23  : 14/04/2006 EPI V_65 Gestion des processus ajout sélection salarié
                            et arret du processus
PT24  : 15/05/2007 FC V_72 S'il existe un élément national dossier niveau population,
                            appeler la fonction de test de validité des populations
PT25  : 23/08/2007 FC V_72 FQ 14651 Application automatique des critères
PT26  : 03/01/2008 GGU V_80 FQ 13889 permettre de choisir : Paie à effectuer/Paies effectuées
PT27  : 03/01/2008 FLO V_802 FQ 14743 - Ajout d'une coche pour les salariés sortis en suppression de bulletins
PT28  : 30/05/2008 FC V_810 FQ 15462 msg "le mois et l'année ne correspondent pas à un exercice social actif. "  lorsqu'on utilise un filtre
PT29  : 07/10/2008 MF FQ 15816 : correction : salariés ne s'affichent pas sur la liste des clotures individuelles
}

unit UTofPG_MulSaisieBul;

interface
uses StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
  HStatus,
{$IFNDEF EAGLCLIENT}
  HDB, DBGrids, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, mul,
{$ELSE}
  MaineAGL, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, vierge, EntPaie, HMsgBox, Hqry, UTOF, UTOB, UTOM,
{$IFNDEF CPS1}                                 
  PGPOPULOUTILS, //PT24
{$ENDIF}
  AGLInit;
type
  TOF_PGMULSAISIEBUL = class(TOF)
  private
    WW: THEdit;
    Q_Mul: THQuery;
    EnCours: TCheckBox;
    vcbxMois, vcbxAnnee: THValComboBox;
    AnneeOk, Cloture: string;
    LAnnee: THEdit;
    ETAB, TN1, TN2, TN3, TN4, CST, CBXBUL: THValComboBox;
    BtnCherche: TToolbarButton97;
    Trait: string; // PT22 Prise en compte cloture et decloture
    DebutBulletin, FinBulletin: TDateTime;
    procedure ActiveWhere(Okok: Boolean);
    procedure GrilleDblClick(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure AccesSal(Sender: TObject);
    // PT12  : 03/07/2003 V_42 PH Suppression automatisée de bulletin
    procedure SuppressionBullettin(Sender: TObject);
    procedure CloturBullettin(Sender: TObject); // PT22
    //procedure ChangeListeP(Sender: TObject); //PT27
    // PT23 ajout fonction de sortie
    procedure ClickSortie (Sender: TObject);
    procedure ClickPaieEnCours (Sender: TObject); //PT25
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
    procedure OnLoad; override;
  end;

implementation

uses P5Util, P5Def, PgOutils
  , PGoutils2
  , paramsoc //PT26
  , SaisBul;

procedure TOF_PGMULSAISIEBUL.ActiveWhere(Okok: Boolean);
var
  Annee, Mois: WORD;
  LaDate: TDateTime;
  St: string;
begin
  //PT27
  annee := 0;
  mois  := 0;

  WW.Text := '';
  st := '';
  if (GetControl ('CLOTDECLOT') <> NIL) then Cloture := GetControlText('CLOTDECLOT') // PT22
  else Cloture := '';
  if (GetCheckBoxState('CHBXPAIESENCOURS') = cbChecked) and (Trait <> 'C') then // PT22
    SetControlVisible('DELETEPAIE', TRUE)
  else SetControlVisible('DELETEPAIE', FALSE);

  if Okok = FALSE then
  begin
    WW.Text := ' PPU_ETABLISSEMENT = "ZZZ" AND PPU_ETABLISSEMENT <> "ZZZ"'; // Pour avoir aucun salarie //PT28 remplacement PSA par PPU car erreur colonne non trouvée
    exit;
  end;
  if vcbxMois.Value <> '' then Mois := Trunc(StrToInt(vcbxMois.Value));
  if vcbxAnnee.Value <> '' then Annee := Trunc(StrToInt(LAnnee.Text));
  if vcbxMois.Value <> '' then // cas où mois renseigné alors année obligatoire
  begin
    LaDate := EncodeDate(annee, mois, 1);
    DebutBulletin := LaDate;
    LaDate := FINDEMOIS(LaDate);
    FinBulletin := LaDate;
  end;
  // PT2 22/10/01 V562 PH Suppression des tests entree,sortie,suspension de paie
  st := '';
  if EnCours <> nil then
  begin
    if not EnCours.Checked then
    begin
      if (vcbxMois <> nil) and (WW <> nil) and (vcbxAnnee <> nil) then
      begin
        St := '(PSA_DATEENTREE <="' + UsDateTime(FinBulletin) + '") AND ((PSA_DATESORTIE >="' + UsDateTime(DebutBulletin) +
          '") OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '")) AND (PSA_DATEENTREE <="' + UsDateTime(FinBulletin) + '")';
      end;
    end;
  end;
  // FIN PT2
  
  //PT27 - Début
  if (GetCheckBoxState('CKSORTIS') = cbChecked) then
  begin
    St := ' AND NOT (PSA_DATESORTIE >="' + UsDateTime(DebutBulletin) + '" AND PSA_DATESORTIE <= "' + UsDateTime(FinBulletin) + '")';
  end;
  //PT27 - Fin
  
  if St <> '' then WW.Text := st;
  if EnCours <> nil then
  begin
    if EnCours.Checked = TRUE then
    begin
      // PT10  : 26/11/2002 V591 PH Acces à la fiche salarie dans le cas de la liste des paies à effectuer
      setControlVisible('BTNSAL', FALSE);
      // PT9   : 05/11/2002 V591 PH Changement de la liste de la query du mul si besoin est
      if (Q_Mul <> nil) and (Q_Mul.Liste <> 'PGLANCEBULLAVEC') then
      begin
//        Q_Mul.Liste := ;
{$IFNDEF EAGLCLIENT}
        TFMul(Ecran).SetDBListe('PGLANCEBULLAVEC');
{$ELSE}
//        TFMul(Ecran).SetDBListe('PGLANCEBULLAVEC');
        TFMul(Ecran).Q.Liste := 'PGLANCEBULLAVEC';
        TFMul(Ecran).DBListe := 'PGLANCEBULLAVEC';
{$ENDIF}
      end;
      WW.Text := WW.Text + ' AND ((PPU_DATEFIN >="' + UsDateTime(DebutBulletin) + '") AND (PPU_DATEFIN <="' + UsDateTime(FinBulletin) + '"))';
      if (ETAB <> nil) and (ETAB.Value <> '') then WW.Text := WW.Text + ' AND (PPU_ETABLISSEMENT="' + ETAB.Value + '")';
      if (TN1 <> nil) and (TN1.Value <> '') then WW.Text := WW.Text + ' AND (PPU_TRAVAILN1="' + TN1.Value + '")';
      if (TN2 <> nil) and (TN2.Value <> '') then WW.Text := WW.Text + ' AND (PPU_TRAVAILN2="' + TN2.Value + '")';
      if (TN3 <> nil) and (TN3.Value <> '') then WW.Text := WW.Text + ' AND (PPU_TRAVAILN3="' + TN3.Value + '")';
      if (TN4 <> nil) and (TN4.Value <> '') then WW.Text := WW.Text + ' AND (PPU_TRAVAILN4="' + TN4.Value + '")';
      if (CST <> nil) and (CST.Value <> '') then WW.Text := WW.Text + ' AND (PPU_CODESTAT="' + CST.Value + '")';
    end
    else
    begin
      // PT10  : 26/11/2002 V591 PH Acces à la fiche salarie dans le cas de la liste des paies à effectuer
      setControlVisible('BTNSAL', TRUE);
      // PT9   : 05/11/2002 V591 PH Changement de la liste de la query du mul si besoin est
      if (Q_Mul <> nil) and (Q_Mul.Liste <> 'PGLANCEBULLSANS') then
      begin
//        Q_Mul.Liste := ;
{$IFNDEF EAGLCLIENT}
        TFMul(Ecran).SetDBListe('PGLANCEBULLSANS');
{$ELSE}
//        TFMul(Ecran).SetDBListe('PGLANCEBULLSANS');
        TFMul(Ecran).Q.Liste := 'PGLANCEBULLSANS';
        TFMul(Ecran).DBListe := 'PGLANCEBULLSANS';
{$ENDIF}
      end;
      if (ETAB <> nil) and (ETAB.Value <> '') then WW.Text := WW.Text + ' AND (PSA_ETABLISSEMENT="' + ETAB.Value + '")';
      if (TN1 <> nil) and (TN1.Value <> '') then WW.Text := WW.Text + ' AND (PSA_TRAVAILN1="' + TN1.Value + '")';
      if (TN2 <> nil) and (TN2.Value <> '') then WW.Text := WW.Text + ' AND (PSA_TRAVAILN2="' + TN2.Value + '")';
      if (TN3 <> nil) and (TN3.Value <> '') then WW.Text := WW.Text + ' AND (PSA_TRAVAILN3="' + TN3.Value + '")';
      if (TN4 <> nil) and (TN4.Value <> '') then WW.Text := WW.Text + ' AND (PSA_TRAVAILN4="' + TN4.Value + '")';
      if (CST <> nil) and (CST.Value <> '') then WW.Text := WW.Text + ' AND (PSA_CODESTAT="' + CST.Value + '")';
      St := 'AND (PSA_SUSPENSIONPAIE <> "X") AND (NOT Exists(Select PPU_SALARIE FROM PAIEENCOURS WHERE PPU_SALARIE=PGLANCEBULLSANSPAIE.PSA_SALARIE';
      St := St + ' AND PPU_DATEDEBUT >="' + UsDateTime(DebutBulletin) + '" AND PPU_DATEFIN <= "' + UsDateTime(FinBulletin) + '" AND (PPU_BULCOMPL <> "X")))';
      WW.Text := WW.Text + St;
      {   if V_PGI <> NIL then
            begin  // gestion des confidentiels
            if V_PGI.Confidentiel = '0' then
               WW.Text:=WW.Text+ ' AND (PSA_CONFIDENTIEL = "0") ';
            end;}
    end;
  end;
  // DEB PT22
  if Trait = 'C' then
  begin
    if Cloture = 'C' then WW.text := WW.text + ' AND PPU_TOPCLOTURE="-"'
    else WW.text := WW.text + ' AND PPU_TOPCLOTURE="X"';
  end;
  // FIN PT22
  if GetControlText('CBXBULLCOMPL') = '001' then exit;
  if (CBXBUL <> nil) and EnCours.Checked then
  begin
    if GetControlText('CBXBULLCOMPL') = '002' then WW.Text := WW.Text +
      ' AND PPU_BULCOMPL ="-"';
    if GetControlText('CBXBULLCOMPL') = '003' then WW.Text := WW.Text +
      ' AND PPU_BULCOMPL ="X"';
  end;
end;


procedure TOF_PGMULSAISIEBUL.OnArgument(Arguments: string);
var
  Num: Integer;
  MoisE, AnneeE, ComboExer: string;
{$IFDEF EAGLCLIENT}
  Grille: THGrid;
{$ELSE}
  Grille: THDBGrid;
{$ENDIF}
  DebExer, FinExer: TDateTime;
  Defaut: ThEdit;
  // PT23 ajout bouton STOP
  BTNSAL, Supp, Clotur, BTNSTOP : TToolbarButton97;
  BP: TCheckBox;
  // PT23
  LeSalarie : String;
  Q : TQuery;
begin
  inherited;
  // PT12  : 03/07/2003 V_42 PH Suppression automatisée de bulletin
  Trait := Trim(ReadTokenSt(Arguments));
  // DEB PT22
  if Trait = 'C' then
  begin
    SetControlVisible('DELETEPAIE', FALSE);
    SetControlVisible('CLOTUR', TRUE);
    SetControlVisible('CLOTDECLOT', TRUE);
    SetControlVisible('LBLCLOTDECLOT', TRUE);
  end
  else
  begin
    if Trait = 'S' then
    begin
      SetControlVisible('CLOTUR', FALSE);
      SetControlVisible('DELETEPAIE', TRUE);
      SetControlVisible('CLOTDECLOT', FALSE);
      SetControlVisible('LBLCLOTDECLOT', FALSE);
      SetControlVisible('CKSORTIS', True); //PT27
    end
    else
    begin   // PT23 Gestion des processus
    if Trait = 'P' then
    begin
     LeSalarie := ReadTokenSt(Arguments);
     if LeSalarie <> '' then SetControlText ('PSA_SALARIE', LeSalarie);
     SetControlVisible('BTNSTOP', TRUE);
     BTNSTOP := TToolbarButton97 (GetControl ('BTNSTOP'));
     if BTNSTOP <> NIL then BTNSTOP.Onclick := ClickSortie;
    end;
    // PT23 fin
   end;
  end;
  if (GetControl ('CLOTDECLOT') <> NIL) then SetControlText('CLOTDECLOT', 'C');
  // PT22
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
{$IFDEF EAGLCLIENT}
  Grille := THGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THDBGrid(GetControl('Fliste'));
{$ENDIF}
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  EnCours := TCheckBox(GetControl('CHBXPAIESENCOURS'));
  //if EnCours <> NIL then EnCours.OnClick := ClickEnCours;
  WW := THEdit(GetControl('XX_WHERE'));
  vcbxMois := THValComboBox(GetControl('CBXMOIS'));
  vcbxAnnee := THValComboBox(GetControl('CBXANNEE'));
  LAnnee := THedit(GetControl('LANNEE'));
  ETAB := THValComboBox(GetControl('ETABLISSEMENT'));
  TN1 := THValComboBox(GetControl('TRAVAILN1'));
  TN2 := THValComboBox(GetControl('TRAVAILN2'));
  TN3 := THValComboBox(GetControl('TRAVAILN3'));
  TN4 := THValComboBox(GetControl('TRAVAILN4'));
  CST := THValComboBox(GetControl('CODESTAT'));
  CBXBUL := THValComboBox(GetControl('CBXBULLCOMPL'));
  //PT8
{$IFDEF CCS3}
  if CBXBUL <> nil then
  begin
    CBXBUL.value := '002';
    CBXBUL.Enabled := False;
  end;
  //PT15
  SetControlVisible('LBLBULCOMPL', False);
  SetControlVisible('CBXBULLCOMPL', False);
  //FIN PT15
{$ELSE}
  if CBXBUL <> nil then
    CBXBUL.value := '001';
{$ENDIF}
  //FIN PT8
  if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
  begin
    if vcbxMois <> nil then vcbxMois.value := MoisE;
    if vcbxAnnee <> nil then vcbxAnnee.value := ComboExer;
  end;

  //if vcbxMois <> NIL then vcbxMois.OnChange := MoisChange;
  // if vcbxAnnee <> NIL then vcbxAnnee.OnChange := AnneeChange;

  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  for Num := 1 to 4 do
  begin // PT22
    if (Trait <> 'S') and (Trait <> 'C') then // PT12  : 03/07/2003 V_42 PH Suppression automatisée de bulletin
      VisibiliteChampSalarie(IntToStr(Num), GetControl('TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)))
    else // PT19 traitement des bons noms des labels dans onglet compléménts
      VisibiliteChampSalarie(IntToStr(Num), GetControl('PPU_TRAVAILN' + IntToStr(Num)), GetControl('TPPU_TRAVAILN' + IntToStr(Num)));
  end;
  // PT12  : 03/07/2003 V_42 PH Suppression automatisée de bulletin
  if (Trait <> 'S') and (Trait <> 'C') then // PT22
  begin
    VisibiliteStat(GetControl('CODESTAT'), GetControl('TPSA_CODESTAT'));
    InitLesTOBPaie;
    ChargeLesTOBPaie;
  end
  else
    VisibiliteStat(GetControl('PPU_CODESTAT'), GetControl('TPPU_CODESTAT'));

  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then Defaut.OnExit := ExitEdit;

  //PT10  : 26/11/2002 V591 PH Acces à la fiche salarie dans le cas de la liste des paies à effectuer
  BTNSAL := TToolbarButton97(GetControl('BTNSAL'));
  if BTNSAL <> nil then BTNSAL.OnClick := AccesSal;
  // PT12  : 03/07/2003 V_42 PH Suppression automatisée de bulletin
  Supp := TToolbarButton97(GetControl('DELETEPAIE'));
  if Supp <> nil then Supp.onClick := SuppressionBullettin;
  // DEB PT22
  Clotur := TToolbarButton97(GetControl('CLOTUR'));
  if Clotur <> nil then Clotur.onClick := CloturBullettin;
  // FIN PT22

  BP := TCheckBox(GetControl('CHBXPAIESENCOURS'));
  if BP <> nil then BP.OnClick := ClickPaieEnCours; //PT25

  //DEB PT24
{$IFNDEF CPS1}
  //S'il existe un élément national dossier niveau population, appeler la fonction de test de validité des populations
  Q := Opensql ('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_TYPENIVEAU = "POP"',true);
  if not Q.Eof then
  begin
    if not CanUsePopulation('PAI') then
      PGIBox(TraduireMemoire('La population "PAI" utilisée dans les éléments nationaux dossier n''est pas valide'));
  end;
  Ferme(Q);
{$ENDIF}
  //FIN PT24
// d PT29
  if Trait <> 'C' then
  begin
  //Début PT26
    if not GetParamSocSecur('SO_PGDEFAULTBULLIST', True) then
      if BP <> nil then
      begin
        BP.Checked := False;
        ClickPaieEnCours(Self);
      end;
  //Fin PT26
  end;
// f PT29  
end;

procedure TOF_PGMULSAISIEBUL.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //PT14 AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGMULSAISIEBUL.GrilleDblClick(Sender: TObject);
var
  CodeSalarie, st, st1, stq, CP, Etabl, BulCompl, ProfilPart: string;
  DateDebut, DateFin, DTClot, DateSortie: TDateTime;
  ActionBul: TActionBulletin;
  Annee, Mois, aa, mm, jj: WORD;
  testB: Boolean;
  Q: tquery;
  GestionCPEtab, AutorisationBulletin, AutorisationSaisie: boolean;
  rep: Integer;
begin
  if (Trait = 'S') or (Trait = 'C') then exit; // PT22

  //PT27
  DateDebut := 0;
  DateFin   := 0;
  DTCLOT    := 0;
  Annee     := 0;
  Mois      := 0;

  testB := FALSE;
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
  AutorisationBulletin := True;
  ActionBul := taModification;
  if (vcbxMois <> nil) and (vcbxMois.Value <> '') then Mois := Trunc(StrToInt(vcbxMois.Value));
  if (vcbxAnnee <> nil) and (LAnnee <> nil) then Annee := Trunc(StrToInt(LAnnee.Text));
  if vcbxMois.Value <> '' then // cas où mois renseigné alors année obligatoire
  begin
    if AnneeOk = '' then
    begin
      if vcbxAnnee <> nil then AnneeOk := IntToStr(Annee);
    end;
    DateDebut := EncodeDate(StrToInt(anneeok), mois, 1);
    DateFin := FINDEMOIS(DateDebut);
    DateSortie := Q_Mul.FindField('PSA_DATESORTIE').AsDatetime;
    if (dateSortie <> 0) and (datesortie > 10) then
    begin
      if Datesortie < Datefin then
        Datefin := Datesortie;
    end;
  end;

  AutorisationSaisie := TRUE;

  CodeSalarie := Q_Mul.FindField('PSA_SALARIE').AsString;
  //  PT5 15/05/02 V575 PH  Message alerte si liste vide et dble click
  if CodeSalarie = '' then
  begin
    if EnCours.Checked = FALSE then st := 'Aucun salarié sélectionné'
    else st := 'Aucune paie sélectionnée';
    PGiBox(st, 'Saisie des bulletins impossible');
    exit;
  end;
  // FIN PT5
  if EnCours.Checked = FALSE then
  begin
    ActionBul := taCreation;
    try
      Etabl := Q_Mul.FindField('PSA_ETABLISSEMENT').Asstring;
      if Etabl = '' then
      begin
        PGiBox('Le salarié doit appartenir à un établissement', 'Saisie des bulletins');
        AutorisationSaisie := FALSE;
      end;
    except
      PGiBox('Le salarié doit appartenir à un établissement', 'Saisie des bulletins');
      AutorisationSaisie := FALSE;
    end;
  end
  else
  begin
    st := Q_Mul.FindField('PPU_TOPCLOTURE').AsString;
    st1 := Q_Mul.FindField('PPU_TOPGENECR').AsString;
    Etabl := Q_Mul.FindField('PPU_ETABLISSEMENT').Asstring;
    if (st1 = 'X') and (St <> 'X') then // PT17
    begin // On force la modification du bulletin uniquement dans le cas où la paie n'est pas cloturée
      //  PT4 26/03/02 V571 PH  Modif message alerte
      Rep := PGIAsk('Confirmez-vous la modification de votre bulletin', 'Attention, les écritures comptables ont été générées');
      if rep = mrYes then st1 := '';
    end;
    if (st = 'X') or (st1 = 'X') then testB := TRUE;
    if testB = TRUE then ActionBul := taConsultation;
  end;
  GestionCPEtab := false;
  if VH_Paie.PGCongesPayes then
  begin
    stq := 'SELECT ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM ETABCOMPL ' +
      'WHERE ETB_ETABLISSEMENT="' + Etabl + '"';
    Q := Opensql(stq, true);
    if not Q.eof then
    begin
      try
        CP := Q.findfield('ETB_CONGESPAYES').AsVariant;
      except
        AutorisationSaisie := FALSE;
        PGiBox('Vous devez indiquer si vous gérez les congés payés ?', 'Saisie des bulletins');
      end;
      GestionCPEtab := (CP = 'X');
      if (GestionCPEtab) and (AutorisationSaisie) then
      begin
        try
          DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
          if (DTClot <= Idate1900) then
          begin
            AutorisationSaisie := FALSE;
            PGiBox('Vous devez saisir la date de clôture des congés payés ?', 'Saisie des bulletins');
          end;
        except
          AutorisationSaisie := FALSE;
          PGiBox('Vous devez saisir la date de clôture des congés payés ?', 'Saisie des bulletins');
        end;
        decodedate(DTCLOT, aa, mm, jj);
        if ((aa < Annee) or ((aa = Annee) and (mm < Mois))) then
          AutorisationBulletin := false;
      end;
    end;
    ferme(Q);
  end;

  // PT6
{$IFDEF CCS3}
  {PT7
   Stq := 'SELECT COUNT(PPU_SALARIE) AS NBPAYE'+
          ' FROM PAIEENCOURS WHERE'+
          ' PPU_DATEDEBUT="'+UsDateTime(DateDebut)+'" AND'+
          ' PPU_DATEFIN="'+UsDateTime(DateFin)+'"';
   Q := Opensql (stq,true);
   if Q.FindField('NBPAYE').AsString <>'' then
      Compteur := Q.FindField('NBPAYE').AsInteger
   else
      Compteur := 0;
   if Compteur > 29 then
    begin
    PgiBox('Bulletin impossible car le nombre de bulletin est limité à 30',Ecran.Caption);
    AutorisationSaisie := FALSE;
    end;
   Ferme (Q);
  FIN PT7}
{$ENDIF}
  // FIN PT6
  if (AutorisationBulletin) then
  begin
    if AutorisationSaisie then
    begin
      if ActionBul = taCreation then
      begin
        // PT1 : 10/09/2001 V547 PH Controle de la date de la paie en création de bulletin
        if not ControlPaieCloture(DateDebut, DateFin) then
        begin
          PgiBox('Vous ne pouvez pas saisir de paie sur une période close', Ecran.Caption);
          exit;
        end;
        // FIN PT1
      end;
      // PT3 22/10/01 V562 PH  Rajout un paramètre à l'appel de la saisie du bulletin
      if ActionBul = taCreation then
        begin
{        st := 'SELECT PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PSA_SALARIE="'+CodeSalarie+'"'+
              ' AND PPU_DATEDEBUT="'+UsDateTime (DateDebut)+'" AND PPU_DATEFIN="'+UsDateTime (DateFin)+'"';
        if ExisteSQL (st) then
          PGIBOX ('Il existe déjà un bulletin pour ce salarié aux mêmes dates,#13#10Vous devez annuler le}
        SaisieBulletin(CodeSalarie, Etabl, '-', '', DateDebut, DateFin, ActionBul, Q_Mul, EnCours.Checked, GestionCPEtab)
        end
      else
      begin
        BulCompl := Q_Mul.FindField('PPU_BULCOMPL').AsString;
        ProfilPart := Q_Mul.FindField('PPU_PROFILPART').AsString;
        SaisieBulletin(CodeSalarie, Etabl, BulCompl, ProfilPart, DateDebut, DateFin, ActionBul, Q_Mul, EnCours.Checked, GestionCPEtab);
      end;

      BtnCherche.click;
    end;
  end
  else
    HShowMessage('1;;La clôture des congés payés de l''établissement n''a pas été effectuée#13#10 Bulletin impossible;N;C;C;', '', '');
end;

procedure TOF_PGMULSAISIEBUL.OnClose;
begin
  if (Trait = 'S') or (Trait = 'C') then exit; // PT22
  VideLesTOBPaie(FALSE);
  //if TOB_Salarie <> NIL Then TOB_Salarie.Free;
  // TOB_Salarie := NIL;
  if TOB_ExerSocial <> nil then TOB_ExerSocial.Free;
  TOB_ExerSocial := nil;
end;

procedure TOF_PGMULSAISIEBUL.OnLoad;
var
  TT: TFMul;
  Okok: Boolean;
  St: string;
begin

  if not (Ecran is TFMul) then exit;
  if (EnCours <> nil) and (Trait <> 'S') and (Trait <> 'C') then // PT22
  begin
    if EnCours.Checked = FALSE then TFMul(Ecran).Caption := TraduireMemoire ('Liste des paies à préparer')
    else TFMul(Ecran).Caption := TraduireMemoire ('Liste des paies effectuées');
    TT := TFMul(Ecran);
    if TT <> nil then UpdateCaption(TT);
  end;
  // DEB PT22
  if (Trait = 'C') then
  begin
    TFMul(Ecran).Caption := TraduireMemoire ('Cloture,décloture individuelle des paies');
    TT := TFMul(Ecran);
    if TT <> nil then UpdateCaption(TT);
  end;
  // FIN PT22
  Okok := TRUE;
  if (vcbxMois <> nil) and (vcbxAnnee <> nil) then
  begin
    SetControlEnabled('BOuvrir', FALSE);
    SetControlEnabled('FListe', FALSE);
    SetControlEnabled('BInsert', FALSE);
    st := GetControlText('CBXANNEE'); //vcbxAnnee.Value
    st := RechDom('PGANNEESOCIALE', St, FALSE);
    //PT28 Tester que le mois est renseigné car après sélection du filtre, on passe dans le OnLoad mais
    //sans que les critères soient renseignés ?? du coup message.
    if vcbxMois.value <> '' then
    begin
      if not ControlMoisAnneeExer(vcbxMois.value, st, AnneeOk)
        then
      begin
        SetFocusControl('CBXMOIS');
        Okok := FALSE;
      end
      else
      begin
        SetControlEnabled('BOuvrir', TRUE);
        SetControlEnabled('FListe', TRUE);
        SetControlEnabled('BInsert', TRUE);
        SetControlProperty('LANNEE', 'Text', AnneeOk);
      end;
    end
    else
      Okok := FALSE;                    //PT28
  end;
  ActiveWhere(Okok);

  inherited;
end;
//PT10  : 26/11/2002 V591 PH Acces à la fiche salarie dans le cas de la liste des paies à effectuer

procedure TOF_PGMULSAISIEBUL.AccesSal(Sender: TObject);
var
  sal: string;
begin
  if (Trait = 'S') or (Trait = 'C') then exit; // PT22

{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
  if Q_Mul <> nil then
  begin
    Sal := TFMUL(Ecran).Q.findfield('PSA_SALARIE').Asstring;
    if Sal <> '' then AGLLanceFiche('PAY', 'SALARIE', '', Sal, '');
    BtnCherche.click;
  end;
end;
// PT12  : 03/07/2003 V_42 PH Suppression automatisée de bulletin

procedure TOF_PGMULSAISIEBUL.SuppressionBullettin(Sender: TObject);
var
  sal, Etabl, BullCompl: string;
  D1, D2: TDateTime;
  St, St1, Clot, SEtab, LeMess: string;
  i, OkOk, rep: Integer;
begin
  if Trait <> 'S' then exit;
  // DEB PT16
  st := GetControlText('PPU_ETABLISSEMENT');
  if st <> '' then
  begin
    i := 0; //PT27
    st1 := readtokenst(st);
    SEtab := ' AND (';
    while st1 <> '' do
    begin
      if i > 0 then SEtab := SEtab + ' OR ';
      SEtab := SEtab + 'PPU_ETABLISSEMENT="' + st1 + '"';
      i := i + 1;
      st1 := readtokenst(st);
    end;
    SEtab := SEtab + ')';
  end;
  St1 := '';
  if GetControlText('PSA_SALARIE') <> '' then St1 := ' AND (PPU_SALARIE="' + GetControlText('PSA_SALARIE') + '")';
  st := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEFIN > "' + UsDAteTime(FinBulletin) + '"' + SEtab + St1;
  if ExisteSQL(St) then
  begin // PT18
    LeMess := 'Des bulletins postérieurs au ' + DateToStr(FinBulletin) + ' ont été validés dans le dossier, suppression impossible#13#10 ' +
      'Vous devez d''abord effectuer le traitement de suppression de tous les bulletins de tous les salariés des périodes postérieures#13#10 ' +
      'Ou bien effectuez la suppression du bulletin de paie en saisie de bulletin';
    PGIError(LeMess, Ecran.caption);
    exit;
  end;
  St := '';
  St1 := '';
  //i := 0;  //PT27
  // FIN PT16
  OkOk := PGIAsk('Attention, voulez-vous vraiment supprimer les bulletins sélectionnés ?', Ecran.caption);
  if OkOk <> MrYes then exit;

  with TFMul(Ecran) do
  begin
//PT21
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;
//FIN PT21

    if FListe.AllSelected then
    begin
      InitMove(Q.RecordCount, '');
      Q.First;
      while not Q.EOF do
      begin
        MoveCur(False);
        Sal := Q.findfield('PSA_SALARIE').Asstring;
        BullCompl := Q.findfield('PPU_BULCOMPL').Asstring;
        D1 := Q.findfield('PPU_DATEDEBUT').AsDateTime;
        D2 := Q.findfield('PPU_DATEFIN').AsDateTime;
        //        st := Q.FindField('PPU_TOPCLOTURE').AsString;
        st1 := Q.FindField('PPU_TOPGENECR').AsString;
        // PT13 :  16/10/2003 V_42 FQ 10874 PH Test si bulletin cloturé avt suppression
        Clot := Q.FindField('PPU_TOPCLOTURE').AsString;
        Etabl := Q.FindField('PPU_ETABLISSEMENT').Asstring;
        if (Clot <> 'X') or (st1 <> 'X') then
        begin // DEB PT20
          if Clot <> 'X' then
          begin
            if St1 = 'X' then rep := PgiAsk('Attention, ce bulletin a été pris en compte pour la génération des ODs de paie.#13#10Voulez vous quand même le supprimer ?', Ecran.Caption)
            else rep := mrYes;
            if rep = mrYes then SuppressionBull(D1, D2, Sal, Etabl, BullCompl, TRUE)
          end
          else PgiBox('La suppression est impossible car la paie est close', Ecran.Caption);
        end;
        Q.NEXT;
      end; // FIN PT20
      FListe.AllSelected := False;
    end
    else
    begin
      InitMove(FListe.NbSelected, '');
      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row - 1);
{$ENDIF}
        MoveCur(False);
        Sal := Q.findfield('PSA_SALARIE').Asstring;
        BullCompl := Q.findfield('PPU_BULCOMPL').Asstring;
        D1 := Q.findfield('PPU_DATEDEBUT').AsDateTime;
        D2 := Q.findfield('PPU_DATEFIN').AsDateTime;
        //        st := Q.FindField('PPU_TOPCLOTURE').AsString;
        st1 := Q.FindField('PPU_TOPGENECR').AsString;
        Etabl := Q.FindField('PPU_ETABLISSEMENT').Asstring;
        // PT13 :  16/10/2003 V_42 FQ 10874 PH Test si bulletin cloturé avt suppression
        Clot := Q.FindField('PPU_TOPCLOTURE').AsString;
        if (Clot <> 'X') or (st1 <> 'X') then
        begin // DEB PT20
          if Clot <> 'X' then
          begin
            if St1 = 'X' then rep := PgiAsk('Attention, ce bulletin a été pris en compte pour la génération des ODs de paie.#13#10Voulez vous quand même le supprimer ?', Ecran.Caption)
            else rep := mrYes;
            if rep = mrYes then SuppressionBull(D1, D2, Sal, Etabl, BullCompl, TRUE)
          end
          else PgiBox('La suppression est impossible car la paie est close', Ecran.Caption);
        end; // FIN PT20
      end;
      FListe.ClearSelected;
    end;
    FiniMove;
  end;

  TFMul(Ecran).BChercheClick(nil);

end;
// FIN PT12
// DEB PT22

procedure TOF_PGMULSAISIEBUL.CloturBullettin(Sender: TObject);
var Okok, i: Integer;
  Sal, Etabl, Quoi, st: string;
  D1, D2: TDateTime;
begin
  if Cloture = 'C' then
  begin
    st := 'cloturer';
    Quoi := 'X';
  end
  else
  begin
    st := 'décloturer';
    Quoi := '-';
  end;
  OkOk := PGIAsk('Attention, voulez-vous vraiment ' + st + ' les bulletins sélectionnés ?', Ecran.caption);
  if OkOk <> MrYes then exit;
  if (TFMul(Ecran).FListe.AllSelected) then
  begin
    OkOk := PGIAsk('Attention, vous avez sélectionné un ensemble de bulletins, #13#10 Vous devrez utiliser la cloture/décloture mensuelle quand tous les bulletins du mois seront faits.' +
      '#13#10 Voulez-vous continuer ?', Ecran.caption);
    if OkOk <> MrYes then exit;
  end;
  with TFMul(Ecran) do
  begin
//PT21
    if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;
//FIN PT21

    if FListe.AllSelected then
    begin
      InitMove(Q.RecordCount, '');
      Q.First;
      while not Q.EOF do
      begin
        MoveCur(False);
        Sal := Q.findfield('PSA_SALARIE').Asstring;
        D1 := Q.findfield('PPU_DATEDEBUT').AsDateTime;
        D2 := Q.findfield('PPU_DATEFIN').AsDateTime;
        Etabl := Q.FindField('PPU_ETABLISSEMENT').Asstring;
        ExecuteSQL('UPDATE PAIEENCOURS SET PPU_TOPCLOTURE="' + Quoi + '" WHERE PPU_SALARIE="' + Sal + '" AND PPU_DATEDEBUT="' +
          USDATETIME(D1) + '" AND PPU_DATEFIN="' + USDATETIME(D2) + '" AND PPU_ETABLISSEMENT="' + Etabl + '"');
        Q.NEXT;
      end; // FIN PT20
      FListe.AllSelected := False;
    end
    else
    begin
      InitMove(FListe.NbSelected, '');
      for i := 0 to FListe.NbSelected - 1 do
      begin
        FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        Q.TQ.Seek(FListe.Row - 1);
{$ENDIF}
        MoveCur(False);
        Sal := Q.findfield('PSA_SALARIE').Asstring;
        D1 := Q.findfield('PPU_DATEDEBUT').AsDateTime;
        D2 := Q.findfield('PPU_DATEFIN').AsDateTime;
        Etabl := Q.FindField('PPU_ETABLISSEMENT').Asstring;
        ExecuteSQL('UPDATE PAIEENCOURS SET PPU_TOPCLOTURE="' + Quoi + '" WHERE PPU_SALARIE="' + Sal + '" AND PPU_DATEDEBUT="' +
          USDATETIME(D1) + '" AND PPU_DATEFIN="' + USDATETIME(D2) + '" AND PPU_ETABLISSEMENT="' + Etabl + '"');
      end;
      FListe.ClearSelected;
    end;
    FiniMove;
  end;
  TFMul(Ecran).BChercheClick(nil);
end;
// FIN PT22

//PT27
(*
procedure TOF_PGMULSAISIEBUL.ChangeListeP(Sender: TObject);
begin
  if (EnCours <> nil) then
  begin
    if EnCours.Checked = TRUE then
    begin
      if (Q_Mul <> nil) and (Q_Mul.Liste <> 'PGLANCEBULLAVEC') then
      begin
        TFMul(Ecran).SetDBListe('PGLANCEBULLAVEC');
      end;
    end
    else if (Q_Mul.Liste <> 'PGLANCEBULLSANS') then
    begin

      TFMul(Ecran).SetDBListe('PGLANCEBULLSANS');
{
      TFMul(Ecran).Q.Liste := 'PGLANCEBULLSANS';
      TFMul(Ecran).DBListe := 'PGLANCEBULLSANS';
}
    end;
  end;
end;
*)

// PT23 gestion arret du processus
procedure TOF_PGMULSAISIEBUL.ClickSortie(Sender: TObject);
var BTNAnn :  TToolbarButton97;
begin
  BTNAnn := TToolbarButton97 (GetControl ('BAnnuler'));
  if BTNAnn <> NIL then
  begin
    TFMul(Ecran).Retour := 'STOP';
    BTNAnn.Click;
  end;
end;

//DEB PT25
procedure TOF_PGMULSAISIEBUL.ClickPaieEnCours(Sender: TObject);
begin
  TFMul(Ecran).BChercheClick(nil);
end;
//FIN PT25

initialization
  registerclasses([TOF_PGMULSAISIEBUL]);
end.

