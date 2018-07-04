{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 18/02/2000
Modifié le ... : 18/02/2000
Description .. : Multi critère de lancement, de sélection des salariés
Suite ........ : en vue de faire une préparation automation
Suite ........ :
Suite ........ : Confectionne une clause XX_WHERE
Mots clefs ... : PG_PREPAUTO
*****************************************************************}
{
PT1  : 11/10/2001 PH      Raffraichissement de la liste avant lancement du
                          traitement
PT2  : 20/11/2001 PH      Bornes Etablissements
PT3  : 26/11/2001 SB V563 Fiche de bug n°372 Message d'alerte incomplet
PT4  : 20/09/2002 PH V585 Fiche de bug n°10217 complète à gauche avec des 0 si
                          code salarié numérique
PT5  : 23/09/2002 PH V585 Date de la preparation auto en fonction des paramsoc
PT6  : 26/11/2002 PH V591 Remplacement des bornes etablissements et salariés par
                          un multiValComboBox
PT7  : 16/01/2003 PH V591 Controles des etabs 1 par 1 modif en complement du PT6
                          FQ 10455
PT8  : 31/01/2003 PH V591 Confection du multiCombo salarié avec que les salariés
                          des etabs concernés
PT9  : 10/03/2003 PH V_42 Rajout nouvelles options Salariés non sortis FQ10435
                          et Salariés ayant eu une saisie par rubrique dans la
                          période.
PT9-1: 30/03/2003 PH V_42 Test si au moins une zone non nulle dans histosaisrub
PT9-2: 17/09/2003 PH V_421 On exclut les lignes d'acompte FQ 10435 suite
PT10 : 29/09/2003 VG V_42 Suppression des GetControlText sur les anciens champs
                          des bornes établissements
PT11 : 30/03/2003 PH V_42 FQ 10423 Dates preparation en paie décalée
PT12 : 03/10/2003 PH V_42 FQ 10835 Sélection des salariés en fonction des établissements
PT13 : 08/12/2003 PH V_50 Prise en compte preparation automatique de bulletins complémentaires
PT14 : 22/03/2004 PH V_50 Controle et tests des exercices CP si bulletins classiques
PT15 : 15/11/2004 PH V_60 FQ 11769 Ergonomie
PT16 : 17/01/2007 GGU V_72 FQ 13404 Préparation automatique de la saisie par rubrique
                           des bulletins complémentaires
PT17 : 19/01/2007 GGU V_72 Gestion des paies au contrat
PT18 : 24/01/2007 V_72 FC Mise en place filtrage des habilitations/poupulations
PT19 : 02/02/2007 PH V_70 Paies au contrat
PT20 : 05/02/2007 FC V_72 Ajout sélection salarié par le processus
PT21 : 18/04/2007 GGU V_72 FQ13728 impossible de passer à la préparation de la
                           période suivante sans sortir de la commande préparation.
PT22 : 15/05/2007 FC V_72 S'il existe un élément national dossier niveau population,
                            appeler la fonction de test de validité des populations
PT23 : 23/10/2007 GGU V_80 FQ 14348 Ajout d'une option de sélection calculé/non calculé
PT24 : 26/11/2007 GGU V_80 Planificateur de tâches multidossier
PT25 : 03/01/2008 GGU V_80 FQ 11470 Ne pas prendre en compte les établissements fictifs Idem DADS-U
PT26 : 03/01/2008 GGU V_80 FQ 13629 Ajouter dans l'onglet complément les tables libres salariés
PT28 : 17/03/2008 FC V_90 FQ 15295 Ajout code statistique dans l'onglet complément
}
unit UTofPG_MulPrepAuto;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97, Hqry,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, DBCtrls, Mul, Fe_Main,
{$ELSE}
  MaineAgl, eMul,
{$ENDIF}
  Grids, HCtrls, HEnt1, EntPaie, HMsgBox, UTOF, UTOB, UTOM, Vierge, P5Util, P5Def,
  AGLInit, PgOutils,
{$IFNDEF CPS1}
  PGPOPULOUTILS, //PT22
{$ENDIF}
  PGoutils2;
type
  TOF_PGMULPREPAUTO = class(TOF)
  private
    Date1, Date2: THEdit;
    // PT16 : La gestion se fait dans le script    LaCheck: TCheckBox;
    DebExer, FinExer: TDateTime;
    SavCheckClickAction: TNotifyEvent;
    BoPaiesAuContrat : string; //PT17
    boMultiDossier : Boolean;
    function RendCloture: Boolean;
    procedure ActiveWhere(Sender: TObject);
    procedure LanceFichePrep(Sender: TObject);
    procedure DateDebutExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    //  PT8 31/01/2003 PH V591 Confection du multiCombo salarié avec que les salariés des etabs concernés
    procedure ChangeEtab(Sender: TObject);
    // PT16 : La gestion se fait dans le script    procedure ExitCheck(Sender: TObject);
    procedure CheckBulCompl(Sender: TObject); // PT16
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  end;

implementation

uses ed_tools, ParamSoc //PT16
    , PGFicheJobMultiDossier //PT24
    ; 

function TOF_PGMULPREPAUTO.RendCloture: Boolean;
var
  Q: TQuery;
  Dat2: TDateTime;
  Annee, Mois, Jour: WORD;
  Indice: Integer;
  IndiceClot: string;
  DF: THEdit;
begin
  result := FALSE;
  DF := THEdit(GetControl('DATEFIN'));
  if DF = nil then exit;
  Dat2 := StrToDate(DF.Text);
  DecodeDate(Dat2, Annee, Mois, Jour);
  if VH_Paie.PGDecalage = TRUE then
  begin
    if Mois = 12 then
    begin
      Indice := 1;
      Annee := Annee + 1;
    end
    else Indice := Mois + 1;
  end
  else Indice := Mois;
  IndiceClot := '';
  Q := OpenSQL('SELECT PEX_CLOTURE FROM EXERSOCIAL WHERE PEX_ANNEEREFER="' + IntToStr(Annee) + '"', TRUE);
  while not Q.EOF do
  begin
    IndiceClot := Q.Fields[0].AsString;
    Q.Next;
  end;
  Ferme(Q);
  if IndiceClot = '' then exit;
  if IndiceClot[Indice] = '-' then result := TRUE;
end;


procedure TOF_PGMULPREPAUTO.ActiveWhere(Sender: TObject);
var
  DD, DF, WW, NomD, NomF: THEdit;
  St, Dat1, Dat2: string;
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  WW := THEdit(GetControl('XX_WHERE'));
  NomD := THEdit(GetControl('NOM1'));
  NomF := THEdit(GetControl('NOM2'));
  // PT13 : 08/12/2003 PH V_50 Prise en compte preparation automatique de bulletins complémentaires
  if (GetControl('CHBXSAISPRIM') = nil) then
  begin
    Dat1 := UsDateTime(StrToDate(DD.Text));
    Dat2 := UsDateTime(StrToDate(DF.Text));
  end
  else
  begin
    // On borne par le debut et la fin de mois si on fait un bulletin complémentaire sur 1 jour
    Dat1 := UsDateTime(DEBUTDEMOIS(StrToDate(DD.Text)));
    Dat2 := UsDateTime(FINDEMOIS(StrToDate(DF.Text)));
  end;
  // FIN PT13
  if (DD <> nil) and (WW <> nil) then
  begin
    if (GetControl('CHBXSAISPRIM') = nil) then // PT13 Test supplémentaire
    begin
      //PT17  Gestion des paies au contrat  and (BoUnBulletinParContrat = 'X')
      // DEB PT19
      if (BoPaiesAuContrat = 'X') then
      begin
        if (GetParamSocSecur ('SO_PGMETHODCONTRAT', '001') = '002') then
        WW.Text := '(PSA_SUSPENSIONPAIE <> "X") AND ((PCI_FINCONTRAT >="' + Dat1 +
          '") AND (PCI_FINCONTRAT <="' + Dat2 +'") AND (PCI_FINCONTRAT IS NOT NULL) AND (PCI_FINCONTRAT > "' + UsDateTime(iDate1900) + '"))' +
          ' AND (PCI_DEBUTCONTRAT <="' + Dat2 + '")'
          else
        WW.Text := '(PSA_SUSPENSIONPAIE <> "X") AND EXISTS (SELECT 1 FROM CONTRATTRAVAIL WHERE PCI_SALARIE = PSA_SALARIE AND ((PCI_FINCONTRAT >="' + Dat1 +
          '") AND (PCI_FINCONTRAT <="' + Dat2 +'") AND (PCI_FINCONTRAT IS NOT NULL) AND (PCI_FINCONTRAT > "' + UsDateTime(iDate1900) + '"))' +
          ' AND (PCI_DEBUTCONTRAT <="' + Dat2 + '"))';

      end else begin
      //Fin PT17
        WW.Text := '(PSA_SUSPENSIONPAIE <> "X") AND ((PSA_DATESORTIE >="' + Dat1 +
          '") OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '"))' +
          ' AND (PSA_DATEENTREE <="' + Dat2 + '")'
      end;       // FIN PT19
    end
    else
      WW.Text := '';
    if (NomD <> nil) and (NomF <> nil) then
    begin
      if (NomD.Text <> '') then
      begin
        if (NomF.Text <> '') then St := ' AND PSA_SALARIE >="' + NomD.Text + '" AND PSA_SALARIE <="' + NomF.Text + '"'
        else St := ' AND PSA_SALARIE ="' + NomD.Text + '"';
        WW.Text := WW.Text + St;
      end;
    end;
  end;
  // PT2 20/11/01 PH Bornes Etablissements
  {PT10
  if (GetControlText ('ETABLISSEMENT') <> '') OR (GetControlText ('ETABLISSEMENT_') <> '')then
     begin
     if GetControlText ('ETABLISSEMENT') <> '' then
        WW.Text := WW.Text+' AND PSA_ETABLISSEMENT >= "'+GetControlText ('ETABLISSEMENT')+'"';
     if GetControlText ('ETABLISSEMENT_') <> '' then
        WW.Text := WW.Text+' AND PSA_ETABLISSEMENT <= "'+GetControlText ('ETABLISSEMENT_')+'"';
     end;
  }
  //  PT9 10/03/2003 PH V_42 options Salariés non sortis et Salariés ayant eu une saisie par rubrique
  if (GetCheckBoxState('CHBXSAISRUB') = cbChecked) then
  begin
    st := ' AND PSA_SALARIE IN (SELECT DISTINCT(PSD_SALARIE) FROM HISTOSAISRUB WHERE PSD_SALARIE=PSA_SALARIE';
    // PT16 17/01/2007 GGU G4estion de la saisie par rubrique des bulletins complémentaires
    if (GetCheckBoxState('CHBXBULCOMPL') = cbChecked) then
      st := st + ' AND PSD_ORIGINEMVT = "BCP" '
    else
      st := st + ' AND PSD_ORIGINEMVT <> "ACP" AND PSD_ORIGINEMVT <> "BCP" '; //  PT9-2 17/09/2003 PH V_421 On exclut les lignes d'acompte FQ 10435 suite
    st := st + ' AND PSD_DATEDEBUT >= "' + Dat1 + '" AND PSD_DATEFIN <= "' + Dat2 + '"' +
      ' AND NOT (PSD_BASE = 0 AND PSD_TAUX = 0 AND PSD_COEFF = 0 AND PSD_MONTANT = 0) ) ';
    WW.Text := WW.Text + St;
  end;
  // PT13 : 08/12/2003 PH V620 Prise en compte preparation automatique de bulletins complémentaires
  if (GetCheckBoxState('CHBXSAISPRIM') = cbChecked) then
  begin
    st := ' AND PSA_SALARIE IN (SELECT DISTINCT(PSP_SALARIE) FROM HISTOSAISPRIM WHERE PSP_SALARIE=PSA_SALARIE' +
      ' AND PSP_DATEDEBUT >= "' + Dat1 + '" AND PSP_DATEFIN <= "' + Dat2 + '"' +
      ' AND NOT (PSP_BASE = 0 AND PSP_TAUX = 0 AND PSP_COEFF = 0 AND PSP_MONTANT = 0) ) ';
    WW.Text := WW.Text + St;
  end;
  if (GetCheckBoxState('CHBXSALSORTIS') = cbChecked) then
  begin
    st := ' AND NOT (PSA_DATESORTIE >="' + Dat1 + '" AND PSA_DATESORTIE <= "' + Dat2 + '")';
    WW.Text := WW.Text + St;
  end;
  // fin PT9
  //Debut PT23
  if (GetCheckBoxState('CHBXNONCALCULE') = cbChecked) then
  begin
    st := ' AND ( NOT Exists( '
        +    ' Select PPU_SALARIE FROM PAIEENCOURS '
        +     ' WHERE PPU_SALARIE=salaries.PSA_SALARIE '
        +       ' AND PPU_DATEDEBUT >= "'+Dat1+'" '
        +       ' AND PPU_DATEFIN <= "'+Dat2+'" ';
    if (GetControl('CHBXBULCOMPL') <> nil) and (GetCheckBoxState('CHBXBULCOMPL') = cbChecked) then
      st := st + ' AND (PPU_BULCOMPL = "X") '
    else
      st := st + ' AND ((PPU_BULCOMPL IS NULL) OR (PPU_BULCOMPL <> "X")) ';
      st := st + ' ) ) ';
    WW.Text := WW.Text + St;
  end;
  //Fin PT23
  //PT25
  WW.Text := WW.Text + ' AND (PSA_ETABLISSEMENT in (Select ET_ETABLISSEMENT from etabliss WHERE (ET_FICTIF <> "X") or (ET_FICTIF is null))) '
end;

procedure TOF_PGMULPREPAUTO.DateDebutExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
  MoisE, AnneeE, ComboExer: string; //PT21
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', 'Préparation automatique');
      SetFocusControl('DATEDEBUT');
    end;
    RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);  //PT21
    if StrToDate(DD.Text) < DebExer then
    begin
      PGIBox('Attention, la date de début est inférieure à la date de début d''exercice', 'Préparation automatique');
      SetFocusControl('DATEFIN');
    end;
  end;
end;

procedure TOF_PGMULPREPAUTO.DateFinExit(Sender: TObject);
var
  DD, DF: THEdit;
  DateDeb, DateFin: TDateTime;
  MoisE, AnneeE, ComboExer: string; //PT21
begin
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de fin est inférieure à la date de début', 'Préparation automatique');
      SetFocusControl('DATEDEBUT');
    end;
    RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);  //PT21
    if StrToDate(DF.Text) > FinExer then
    begin // PT15
      PGIBox('Attention, la date de fin est supérieure à la date de fin d''exercice', 'Préparation automatique');
      SetFocusControl('DATEFIN');
    end;
  end;
end;

procedure TOF_PGMULPREPAUTO.LanceFichePrep(Sender: TObject);
var
  st, CpGere, TheEtab, Etablis: string;
  DD, DF: THEdit;
  LadateD, LadateF, DateDeb, DateFin: TDateTime;
  rep: Integer;
  Q: TQuery;
  Etab: THValComboBox;
  OkOk: Boolean;
  BtnCherche: TToolbarButton97;
  NbreEtab: Integer;


  TobParam, T1: TOB; //PT24
  LeSQL : string;       //PT24

begin
  // PT1 11/10/01 PH Raffraichissement de la liste avant lancement du traitement
  BtnCherche := TToolbarButton97(GetControl('BCherche'));
  if BtnCherche <> nil then BtnCherche.Click;

  // Controle des clotures pour ne pas preparer/recalculer des paies cloturees
  OkOk := RendCloture;
  if not OkOk then
  begin
    PGIBox('Attention, le mois sélectionné est clos #13#10 Vous ne pouvez pas faire de préparation automatique', 'Préparation automatique');
    exit;
  end;
  DD := THEdit(GetControl('DATEDEBUT'));
  DF := THEdit(GetControl('DATEFIN'));
  if (DD <> nil) and (DF <> nil) then
  begin
    DateDeb := StrToDate(DD.Text);
    DateFin := StrToDate(DF.Text);
    if DateFin < DateDeb then
    begin
      PGIBox('Attention, la date de début est supérieure à la date de fin', 'Préparation automatique');
      SetFocusControl('DATEDEBUT');
      exit;
    end;
    LaDateD := DEBUTDEMOIS(DateDeb);
    LaDateF := FINDEMOIS(DateFin);
    rep := mrYes;
    if (FINDEMOIS(LaDateD) <> LaDateF) then
      rep := PGIAsk('Attention, la période de paie est supérieure à 1 mois ou à cheval sur 2 mois', 'Préparation automatique');
    if rep = mrYes then
    begin
      st := 'Select PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEDEBUT="' + USDateTime(StrToDate(DD.Text)) + '" AND PPU_DATEFIN="' + USDateTime(StrToDate(DF.Text)) + '"';
      Q := OpenSql(st, TRUE);
      if not Q.Eof then
        rep := PGIAsk('Attention, il existe déjà des paies sur la même période. Voulez-vous continuer ?', 'Préparation automatique'); //PT3
      Ferme(Q);
    end;
    // PT14 : 22/03/2004 PH V_50 Controle et tests des exercices CP si bulletins classiques
    if GetControl('CHBXSAISPRIM') = nil then
    begin // Contrôle des exerices CP uniquement dans le cas de bulletins non complémentaires
      // Controle des dates de cloture des exerices congés payés par rapport à la période
      //  PT6 26/11/2002 PH V591 Remplacement des bornes etablissements par un multiValComboBox
      Etab := THValComboBox(GetControl('PSA_ETABLISSEMENT'));
      if Etab <> nil then
      begin
        St := 'SELECT ETB_ETABLISSEMENT,ETB_LIBELLE,ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM ETABCOMPL';
        if GetControlText('PSA_ETABLISSEMENT') <> '' then
        begin
          // PT7 16/01/2003 PH V591 Controles des etabs 1 par 1 modif en complement du PT6
          st := st + ' WHERE ';
          NbreEtab := 0;
          TheEtab := GetControlText('PSA_ETABLISSEMENT');
          while TheEtab <> '' do
          begin
            Etablis := ReadTokenSt(TheEtab);
            if NbreEtab > 0 then st := st + ' OR';
            st := st + ' ETB_ETABLISSEMENT = "' + Etablis + '"';
            NbreEtab := NbreEtab + 1;
          end;
        end else
        //DEBUT PT25
        begin
          st := st + ' WHERE ';
          NbreEtab := 0;
          Q := OpenSql('Select ET_ETABLISSEMENT from etabliss WHERE (ET_FICTIF <> "X") or (ET_FICTIF is null)', TRUE);
          while not Q.Eof do
          begin
            Etablis := Q.Fields[0].AsString;
            if NbreEtab > 0 then st := st + ' OR';
            st := st + ' ETB_ETABLISSEMENT = "' + Etablis + '"';
            NbreEtab := NbreEtab + 1;
            Q.next;
          end;
          Ferme(Q);
        end;
        //Fin PT25
        // FIN PT7
        St := St + ' ORDER BY ETB_ETABLISSEMENT';
        Q := OpenSql(st, TRUE);
        while not Q.Eof do
        begin
          LaDateF := Q.FindField('ETB_DATECLOTURECPN').AsFloat;
          CpGere := Q.FindField('ETB_CONGESPAYES').AsString;
          if (DateFin > LaDateF) and (LaDateF > 10) and (CpGere = 'X') then // Erreur la date de la preparation automatique est > à la date de cloture des cp
          begin
            rep := mrCancel;
            PGIBox('La date de la paie est postérieure à la date de clôture des congés payés pour l''établissement ' + Q.FindField('ETB_LIBELLE').AsString,
              'Préparation automatique');
            break;
          end;
          Q.next;
        end;
        Ferme(Q);
      end
      else
      begin
        rep := mrCancel; // Pas d'etablissement donc erreur
      end
    end
    else rep := MrYes; // Bulletin complémentaires donc OK
    // FIN PT14
    if rep = mrYes then
    begin
      st := DD.Text + ';' + DF.Text;
{$IFDEF EAGLCLIENT}
      if TFMul(Ecran).Fetchlestous then
        TheMulQ := TOB(Ecran.FindComponent('Q'))
      else
      begin
        PgiBox('Vous n''avez pas de ligne total dans votre liste, #13#10 Traitement impossible ', Ecran.Caption);
        exit;
      end;
{$ELSE}
      TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}

      // PT14 : 22/03/2004 PH V_50 Controle et tests des exercices CP si bulletins classiques
      if GetControl('CHBXSAISPRIM') <> nil then
      begin // Passage en paramètre du profil et bulcompl
        st := st + ';X;' + GetControlText('LEPROFILPART');
      end;
      // FIN PT14
      // PT16    // Debut PT19
      if (GetControl('CHBXBULCOMPL') <> nil) and (GetCheckBoxState('CHBXBULCOMPL') = cbChecked) then
        st := st + ';X;' + GetControlText('LEPROFILPART')
      else st := st + ';-;';

      if BoPaiesAuContrat = 'X' then st := st + ';X'
      else st := st + ';-';
      // Fin PT16   // Fin PT19

      //PT24
      if boMultiDossier then
      begin
        st := st + ';MULTI';
        LeSQL := 'SELECT PSA_LIBELLE, PSA_PRENOM, PSA_SALARIE, PSA_ETABLISSEMENT, PSA_DATESORTIE, PSA_DATEENTREE, PSA_SUSPENSIONPAIE '
               + ' FROM SALARIES '
               + ' WHERE (((PSA_SUSPENSIONPAIE <> "X") AND ((PSA_DATESORTIE >="'+USDATETIME(DateDeb)+'") OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "'+USDATETIME(iDate1900)+'")) AND (PSA_DATEENTREE <="'+USDATETIME(DateFin)+'"))) '
               + ' ORDER BY PSA_ETABLISSEMENT ,PSA_LIBELLE';
        TobParam := TOB.create('Ma Tob de Param', nil, -1);
        T1 := TOB.Create('XXX', TobParam, -1);
        T1.AddChampSupValeur('DD', DD.Text);
        T1.AddChampSupValeur('DF', DF.Text);
        T1.AddChampSupValeur('LEMODE', '2TIERS'); //Mode 2/3 pour que la requete soit traduite
        T1.AddChampSupValeur('LESQL', LeSQL);
        FicheJobMultiDossier(0, taCreat, 'cgiPaieS5', 'PREPAUTO', TobParam, '', '', '', '', True, 2);
        FreeAndNil(TOBParam); 
      end else begin
        st := st + ';-';
        AglLanceFiche('PAY', 'PREPARATION_AUTO', '', '', st);
      end;
      TheMulQ := nil;
    end;
  end;
end;

procedure TOF_PGMULPREPAUTO.OnArgument(Arguments: string);
var
  Num: Integer;
  BtnValidMul: TToolbarButton97;
  Nom1, Nom2: ThEdit;
  Zdate: TDateTime;
  ZEtab: ThMultiValComboBox;
  MoisE, AnneeE, ComboExer: string;
  st: string; //PT17
  LeSalarie : String;  // PT20
  Q : TQuery; //PT22
  NewTW : TToolWindow97; 
begin
  inherited;
  //PT17
  st := Trim(Arguments);
  //PT24
  if st = 'MULTI' then
  begin
    boMultiDossier := True;
    SetControlEnabled('CHBXBULCOMPL', False);
    SetControlEnabled('PSA_ETABLISSEMENT', False);
    SetControlEnabled('PSA_SALARIE', False);
    SetControlEnabled('TPSA_ETABLISSEMENT', False);
    SetControlEnabled('LBLDU', False);
    SetControlEnabled('CHBXSAISRUB', False);
    SetControlEnabled('CHBXSALSORTIS', False);
    SetControlEnabled('CHBXNONCALCULE', False);
    (GetControl('PComplement') as THTabSheet).TabVisible := False;
    (GetControl('PAvance') as THTabSheet).TabVisible := False;
    SetControlVisible('DOCK971', False);
    SetControlVisible('FLISTE', False);
    NewTW := TToolWindow97.Create(self.Ecran);
    NewTW.DockedTo := GetControl('DOCK') as TDock97;
    NewTW.FullSize := True;
    SetControlVisible('PANELBOUTON', False);
    (GetControl('BOuvrir') as TToolbarButton97).Parent := NewTW;
    (GetControl('BAnnuler') as TToolbarButton97).Parent := NewTW;
    (GetControl('BAide') as TToolbarButton97).Parent := NewTW;  
  end;
  // DEB PT19
  BoPaiesAuContrat := ReadTokenSt(st);
  if (BoPaiesAuContrat = 'X') then // and (BoUnBulletinParContrat = 'X')
  begin
    if (GetParamSocSecur ('SO_PGMETHODCONTRAT', '001') = '002') then
      TFMul(Ecran).SetDBListe('PGMULBULLETINCONT');
    Ecran.Caption := 'Préparation automatique des paies aux contrats';
    UpdateCaption(Ecran);
  end;
  // FIN PT19
  
  //DEB PT20
  LeSalarie := ReadTokenSt(st);
  If LeSalarie <> '' then
    SetControlText('PSA_SALARIE',LeSalarie);
  //FIN PT20

  //fin PT17
  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then Break;
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  
  //Debut PT26
  for Num := 1 to 4 do
  begin
    VisibiliteChampLibreSal(IntToStr(Num),GetControl ('PSA_LIBREPCMB'+IntToStr(Num)),
                            GetControl ('TPSA_LIBREPCMB'+IntToStr(Num)));
    VisibiliteBoolLibreSal (IntToStr(Num),
                            GetControl ('PSA_BOOLLIBRE'+IntToStr(Num)));
  end;
  //Fin PT26
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT')); //PT28

  // PT16 On sauvegarde l'action (le script) associé au click du CheckBox par l'AGL
  if (GetControl('CHBXBULCOMPL') <> nil) then
  begin
    if Assigned((GetControl('CHBXBULCOMPL') as THCheckbox).OnClick) then
      SavCheckClickAction := (GetControl('CHBXBULCOMPL') as THCheckbox).OnClick;
    (GetControl('CHBXBULCOMPL') as THCheckbox).OnClick := CheckBulCompl;
  end;
  //Fin PT16
  SetControlVisible('LBLANNEE', FALSE);
  SetControlVisible('COMBOEXER', FALSE);
  BtnValidMul := TToolbarButton97(GetControl('BOuvrir'));
  if BtnValidMul <> nil then BtnValidMul.OnClick := LanceFichePrep;
  Date1 := THEdit(GetControl('DATEDEBUT'));
  if Date1 <> nil then Date1.OnClick := DateDebutExit;
  Date2 := THEdit(GetControl('DATEFIN'));
  if Date2 <> nil then Date2.OnClick := DateFinExit;
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  //  PT5 23/09/2002 PH V585 Date de la preparation auto en fonction des paramsoc
  if VH_Paie.PgDatePrepAuto = 'MOI' then
  begin
    if StrToDate(Date2.Text) > FinExer then
    begin
      Date2.Text := DateToStr(FinExer);
      Date1.Text := DateToSTr(DEBUTDEMOIS(FinExer));
    end;
  end
  else
  begin //Recuperation de la periode en cours de la saisie par rubrique
    // PT11 : 30/03/2003 PH V_42 FQ 10423 Dates preparation en paie décalée
    if not VH_Paie.PGDecalage then ZDate := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01)
    else
    begin
      if MoisE = '12' then ZDate := EncodeDate(StrToInt(AnneeE) - 1, StrToInt(MoisE), 01)
      else ZDate := EncodeDate(StrToInt(AnneeE), StrToInt(MoisE), 01);
    end;
    // FIN PT11
    Date1.Text := DateToSTr(Zdate);
    Date2.Text := DateToSTr(FINDEMOIS(Zdate));
  end;
  // FIN PT5
  //  PT4 20/09/2002 complète à gauche avec des 0 si code salarié numérique
  Nom1 := ThEdit(getcontrol('NOM1'));
  if Nom1 <> nil then Nom1.OnExit := ExitEdit;
  Nom2 := ThEdit(getcontrol('NOM2'));
  if Nom2 <> nil then Nom2.OnExit := ExitEdit;
{$IFDEF EAGLCLIENT}
  SetControlVisible('BParamListe', TRUE);
{$ENDIF}
  //  PT8 31/01/2003 PH V591 Confection du multiCombo salarié avec que les salariés des etabs concernés
  ZEtab := THMultiValComboBox(GetControl('PSA_ETABLISSEMENT'));
  if ZEtab <> nil then
    ZEtab.OnChange := ChangeEtab;
// PT16 17/01/2007 : La gestion se fait dans le script
//  // PT13 : 08/12/2003 PH V620 Prise en compte preparation automatique de bulletins complémentaires
//  LaCheck := TCheckBox(GetControl('CHBXBULCOMPL'));
//  if LaCheck <> nil then LaCheck.OnClick := ExitCheck;

  //DEB PT22
{$IFNDEF CPS1}
  //S'il existe un élément national dossier niveau population, appeler la fonction de test de validité des populations
  Q := Opensql ('SELECT PED_CODEELT FROM ELTNATIONDOS WHERE PED_TYPENIVEAU = "POP"',true);
  if not Q.Eof then
  begin
    if not CanUsePopulation('PAI') then
      PGIBox('Le paramétrage population n''est pas valide');
  end;
  Ferme(Q);
{$ENDIF}
  //FIN PT22
end;

procedure TOF_PGMULPREPAUTO.OnLoad;
begin
  inherited;
  ActiveWhere(nil);
end;
//  PT4 20/09/2002 complète à gauche avec des 0 si code salarié numérique

procedure TOF_PGMULPREPAUTO.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then //AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;
//  PT8 31/01/2003 PH V591 Confection du multiCombo salarié avec que les salariés des etabs concernés

procedure TOF_PGMULPREPAUTO.ChangeEtab(Sender: TObject);
var
  ZSal: THMultiValComboBox;
  TheEtab, Etablis: string;
  NbreEtab: Integer;
  St: string;
begin
  St := '';
  if not MonHabilitation.Active then //PT18
  begin
    ZSal := THMultiValComboBox(GetControl('PSA_SALARIE'));
    if ZSal <> nil then
    begin
      NbreEtab := 0;
    // PT12 : 03/10/2003 PH V_42 FQ 10835 Sélection des salariés en fonction des établissements
      TheEtab := GetControlText('PSA_ETABLISSEMENT');
      while (TheEtab <> '') and (TheEtab[1] <> '<') do
      begin
        Etablis := ReadTokenSt(TheEtab);
        if NbreEtab > 0 then st := st + ' OR';
        st := st + ' PSA_ETABLISSEMENT = "' + Etablis + '"';
        NbreEtab := NbreEtab + 1;
      end;
      if St <> '' then ZSal.Plus := ' ' + St + ' ORDER BY PSA_LIBELLE'
      else ZSal.Plus := ' ORDER BY PSA_LIBELLE';
    end;
  end;
  // FIN PT12
end;
// FIN PT8

// PT16 : La gestion se fait dans le script
// PT13 : 08/12/2003 PH V620 Prise en compte preparation automatique de bulletins complémentaires
//procedure TOF_PGMULPREPAUTO.ExitCheck(Sender: TObject);
//begin
//  if LaCheck <> nil then
//  begin
//    if LaCheck.Checked then SetControlEnabled('LEPROFILPART', TRUE)
//    else SetControlEnabled('LEPROFILPART', FALSE);
//  end;
//end;
// FIN PT13
// Fin PT16

procedure TOF_PGMULPREPAUTO.CheckBulCompl(Sender: TObject);
begin
  //On lance le script d'action défini dans la fiche si il existe.
  if Assigned(SavCheckClickAction) then SavCheckClickAction(Sender);

  if GetCheckBoxState('CHBXBULCOMPL') = cbChecked then begin
    SetControlText('DATEFIN', GetControlText('DATEDEBUT'));
    SetFocusControl('DATEFIN');
  end else begin
    SetControlText('DATEFIN', AGLdatetostr(FinDeMois(AGLstrtodate(GetControlText('DATEDEBUT')))));
    SetFocusControl('DATEFIN');
  end;
end;

initialization
  registerclasses([TOF_PGMULPREPAUTO]);
end.

