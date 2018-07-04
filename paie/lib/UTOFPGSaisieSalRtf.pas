{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 08/10/2001
Modifié le ... :   /  /
Description .. : Multicritère de selection salarie pour la fusion WORD
Mots clefs ... : PAIE;FUSION
*****************************************************************
PT1 08/10/2001 V562 SB Copie des modèles DOT proc. CopyDotInModele
PT2 29/01/2002 V571 SB Idate1900 au lieu de 01/01/1900
PT3 29/03/2002 V571 JL Ajout médecine du travail
PT4 06/11/2002 V595 JL Modification pour médecine du travail, gestion intervalle de dates pour els visites,
                       nouvelle fiche : CONVOCMEDECINE utilisée au lieu de SALARIE_RTF
PT5 08/11/2002 V595 JL Ajout convocation pour formation
                       Modification des arguments pour affichage formation
PT6 05/12/2002 V591 SB Copy des .dot systématique si existe sous StdPath
PT7 16/09/2003 V_42 JL FQ 10342 remplacement de l'établissement salarié par établissement du contrat
PT8 03/10/2003 V_42 SB Portage CWAS, gestion des .dot et .doc en eagl
PT9 16/10/2003 V_42 JL Plantage liste contratravail : PSA_SALARIE renommé en PCI_SALARIE
PT10 15/12/2003 V_50 JL Correction nom fichier convocation formation
PT11 19/01/2003 V_50 JL FQ 10987
PT12 08/09/2004 V_50 PH Optimisation SQL pour éviter de charger l'ensemble des salariés pour traiter 1 salarié
PT13 04/11/2004 V_60 PH Ergonomie rajout bouton de lancement double mouette verte au lieu de Bouvrir
PT14 04/02/2005 V_65 SB Ajout champ tob etablissement
PT15 18/02/2005 V_60 JL Ajout dévlarant médecine travail (nouvelle fiche SALARIE_RTFMED)
PT16 21/03/2005 V_60 SB FQ 11423 Ajout fn renvoie donnée organisme
PT17 20/09/2005 V_65 SB FQ 12421 Ajout info registre, table annuaire
PT18 03/02/2006 V_65 SB FQ 12876 Optimisation du chargement des tobs
PT19 31/03/2006 V_65 EPI FQ 12791 Ajout gestion des processus, passage salarié en paramètre
PT20 18/04/2006 V_65 EPI Gestion des processus ajout arret du processus
PT21 16/11/2006 V_70 JL Gestion des fichiers en base (avec possibilité d'utiliser fichier sur disque)
PT22 26/04/2007 V_72 FC FQ 13712 Pb clé pour annuaire
PT23 04/07/2007 V_72 FC FQ 14516 Si coche "Afficher tous les salariés" avec des habilitations paramétrées, alors ligne grise.
}
unit UTOFPGSaisieSalRtf;

interface
uses Controls, Classes,  sysutils,  HTB97, Windows,StdCtrls,

{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,  Fe_main, mul,  AGLUtilOLE,FileCtrl,
{$ELSE}
  MaineAgl, emul, AglUtilOle,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, HQry,UyFileSTD,
  HStatus, ed_tools, UTOB, ParamDat,P5Def;

var
  PGPathDocContrat: string;
type
  TOF_PGSAISIESALRTF = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
  private
    Fichier: string;
    ToutSelectioner: Boolean;
    T_Sal, T_Net, T_Cal, T_Contrat, T_Soc, T_Conv, T_Org, T_VisiteMed, T_AdrMed, T_ConvocForm : tob; //PT3
    T_Annuaire : Tob; { PT17 }
    Q_Mul: THQuery;
    THSalContrat: THEdit;
    THVEtabContrat: THValComboBox;
    // PT19 début appel processus
    AppelProc : boolean;
    SelectSal : String;
    // PT19 fin
    procedure Change(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure CopyDotInModele(StNomModele: string);
    procedure RecupDotInDirectoryTemp; //PT8
    procedure InitTob_Fusion(ListSal: TStringList);
    function GetESexe(Sexe: string): string;
    procedure DateElipsisclick(Sender: TObject); //PT3
    procedure AffichDeclarant(Sender:TObject); //PT15
    // PT20 ajout fonction de sortie
    procedure ClickSortie (Sender: TObject);
    procedure voirDocDefaut (Sender: TObject);
    procedure VoirImport (Sender: TObject);


  end;

implementation
uses PGFusionWord, EntPaie, Pgoutils, PgOutils2;

procedure TOF_PGSAISIESALRTF.OnArgument(Arguments: string);
var
{$IFNDEF EAGLCLIENT}
  Grille: THDBGrid;
{$ELSE}
  Grille: THGrid;
{$ENDIF}
  Defaut,Edit : ThEdit;
  Arg, CodeStage, Ordre,St : string; //PT5
  TrtProc : String;                 // PT20
  // PT20 BTN: TToolbarbutton97;
  // PT20 ajout bouton STOP
  BTN, BTNSTOP : TToolbarButton97;
  Q : TQuery;
  Nature,StPcl,Crit2,Mes : String;
  Combo : THValComboBox;
  Rep : Integer;
begin
  //DEBUT PT5
  Arg := ReadTokenPipe(Arguments, ';');
  //  DEBUT PT21 Vérification gestion des fichiers en base
  //Recherche de modèle par défaut
  If ARg = 'CONTRAT' then
  begin
    Crit2 := 'CNT';
    If V_PGI.ModePcl = '1' then StPcl := ' AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
    else StPcl := '';
    If not existeSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_CRIT1="MAW" AND YFS_CRIT2="'+Crit2+'" '+
    ' AND ((YFS_PREDEFINI="CEG" AND YFS_BCRIT1="X") OR (YFS_PREDEFINI<>"CEG"))'+StPcl) then //Recherche fichier <> CEG ou modèle par défaut pour savoir si client utilise fichier en base
    begin
{         PGIBox('Cette version apporte un nouveau fonctionnement dans '+
         '#13#10la gestion du stockage des maquettes word.'+
         '#13#10#13#10Par défaut l''utilisation reste identique.'+
         '#13#10#13#10Toutefois nous vous invitons à mettre en place cette nouvelle fonctionnalitée','Gestion des fichiers en base');}
         Mes := 'Cette version apporte un nouveau fonctionnement dans '+
         '#13#10la gestion du stockage des maquettes word.'+
         '#13#10#13#10Par défaut l''utilisation reste identique.'+
         '#13#10#13#10Toutefois nous vous invitons à mettre en place cette nouvelle fonctionnalité';
         SetControlChecked('SANSFICHIERBASE',true);
         HShowMessage('1;Gestion des fichiers en base ;' + Mes + ';E;HO;;;;;41910025', '', '');

  {       PGIAsk('Vous n''avez importé aucun fichier en base, voulez-vous effectuer l''import maintenant','Gestion des fichiers en base'); //si non utlisé on propose l'import
         If Rep = MrYes then
         begin
              AglLanceFiche('PAY', 'IMPORTRTF', '', '', '');
              If not existeSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_CRIT1="MAW" AND YFS_CRIT2="'+Crit2+'" '+
              ' AND YFS_BCRIT1="X"'+StPcl) then  //Recherche de doc par défaut, si aucun on porpose saisie
              begin
                rep := PGIAsk('Vous n''avez aucun fichier par défaut, voulez-vous en affecter un maintenant','Gestion des fichiers en base');
                If Rep = MrYes then AglLanceFiche('PAY', 'RTFDEFAUT', '', '', Crit2);
              end
         end
         else
         begin
              SetControlChecked('SANSFICHIERBASE',true);
         end;
    end
    else
    begin
         If not existeSQL('SELECT YFS_NOM FROM YFILESTD WHERE YFS_CODEPRODUIT="PAIE" AND YFS_CRIT1="MAW" AND YFS_CRIT2="'+Crit2+'" '+
         ' AND YFS_BCRIT1="X"'+StPcl) then  //Recherche de doc par défaut, si aucun on porpose saisie
         begin
                rep := PGIAsk('Vous n''avez aucun fichier par défaut, voulez-vous en affecter un maintenant','Gestion des fichiers en base');
                If Rep = MrYes then AglLanceFiche('PAY', 'RTFDEFAUT', '', '', Crit2);
         end}
    end;
  end
  else SetControlVisible('SANSFICHIERBASE',False);
  BTN := TToolBarButton97(GetControl('BVOIRIMPORT'));
  If BTN <> Nil then
  begin
       BTN.ShowHint := True;
       BTN.OnClick := VoirImport;
  end;
  BTN := TToolBarButton97(GetControl('BVOIRDEFAUT'));
  If BTN <> Nil then
  begin
       BTN.ShowHint := True;
       BTN.OnClick := voirDocDefaut;
  end;
  //FIN PT21
  // PT19 CodeStage := ReadTokenPipe(Arguments, ';');
  // PT19 Ordre := ReadTokenPipe(Arguments, ';');
  // PT19 début
  SelectSal := '';
  AppelProc := False;
  If (Arg = 'SOLDE') Or (Arg = 'CERTIFICAT') Or (Arg = 'CONTRAT') then
  begin
    // PT20 début
    TrtProc := ReadTokenPipe(Arguments, ';');
    If TrtProc = 'P' then
      AppelProc := True;
    // PT20 fin
  	SelectSal := ReadTokenPipe(Arguments, ';');
    // PT20 If SelectSal <> '' then AppelProc := True;
  end
  else
    begin
	  CodeStage := ReadTokenPipe(Arguments, ';');
	  Ordre := ReadTokenPipe(Arguments, ';');
  end;
  // PT19 fin
  //FIN PT5
  // PT1  restructuration du onargument
  RecupDotInDirectoryTemp;
  //CopyDotInModele('DicoPGI.dot');  PT8 Mise en commentaire regroupé dans RecupDotInDirectoryTemp
  SetControlText('OBJET', Arg);
  Nature := 'DIV';
  If Arg = 'CERTIFICAT' then Nature := 'CER';
  if Arg = 'FORMATION' then //PT5
  begin
    Nature := 'COF';
    SetControlVisible('Pages', false);
    // SetControlProperty('PSA_SALARIE','Name','SALARIE');
    // SetControlProperty('PSA_LIBELLE','Name','LIBELLE');
    // SetControlProperty('PSA_NUMEROSS','Name','NUMEROSS');
    // SetControlProperty('PSA_ETABLISSEMENT','Name','ETABLISSEMENT');
  end;
  if Arg = 'MEDECINE' then //DEBUT PT3 et PT4
  begin
    Nature := 'MED';
    SetControlVisible('DATEVISITE', True);
    SetControlVisible('TDATEVISITE', True);
    SetControlVisible('DATEVISITE_', True);
    SetControlVisible('TDATEVISITE_', True);
    Defaut := THedit(GetControl('DATEVISITE'));
    if Defaut <> nil then Defaut.OnElipsisClick := dateElipsisClick;
    Defaut := THedit(GetControl('DATEVISITE_'));
    if Defaut <> nil then Defaut.OnElipsisClick := dateElipsisClick;
    SetControlText('DATEVISITE', DateToStr(Date));
    SetControlText('DATEVISITE_', DateToStr(Date));
  end
  else
  begin
    SetControlVisible('DATEVISITE', False);
    SetControlVisible('TDATEVISITE', False);
    SetControlVisible('DATEVISITE_', False);
    SetControlVisible('TDATEVISITE_', False);
    SetControlText('DATEVISITE', DateToStr(IDate1900));
    SetControlText('DATEVISITE_', DateToStr(IDate1900));
  end; //FIN PT3 et PT4
  if Arg = 'SOLDE' then
  begin
    Nature := 'SOL';
    TFMul(Ecran).Caption := 'Solde de tout compte'; //PT2
    // PT19 SetControlText('XX_WHERE', 'PSA_DATESORTIE IS NOT NULL AND PSA_DATESORTIE<>"' + UsdateTime(Idate1900) + '" ' +
    // PT19  'AND PSA_SALARIE IN (SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS)');
    // PT19 Début
    // PT20 If AppelProc = False then
    If (AppelProc = False) or (SelectSal = '') then
      begin
      SetControlText('XX_WHERE', 'PSA_DATESORTIE IS NOT NULL AND PSA_DATESORTIE<>"' + UsdateTime(Idate1900) + '" ' +
        'AND PSA_SALARIE IN (SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS)');
      end
		Else
      begin
			SetControlText('XX_WHERE',SelectSal);
   End;
   // PT19 fin
    //  CopyDotInModele('Soldedetoutcompte.dot'); PT8 Mise en commentaire regroupé dans RecupDotInDirectoryTemp
    SetControlVisible('CKVOIRTSSAL', True);
  end
  else
    if Arg = 'CERTIFICAT' then
  begin
    Nature := 'CER';
    SetControlText('XX_WHERE', '');
    // PT19 Début
    // PT20 If AppelProc = True then
    If (AppelProc = True) and (SelectSal <> '') then
      begin
			SetControlText('XX_WHERE',SelectSal);
      end;
    // PT19 fin
    TFMul(Ecran).caption := 'Certificat de travail';
    //    CopyDotInModele('CertificatTravail.dot'); PT8 Mise en commentaire regroupé dans RecupDotInDirectoryTemp
  end
  else
    if Arg = 'CONTRAT' then
  begin
    Nature := 'CNT';
    SetControlText('XX_WHERE', '');
    // PT19 Début
    // PT20 If AppelProc = True then
    If (AppelProc = True) and (SelectSal <> '') then
      begin
			SetControlText('XX_WHERE',SelectSal);
    end;
    // PT19 fin

    TFMul(Ecran).caption := 'Contrat de travail';
    TFMul(Ecran).SetDBListe('PGMULCONTRAT');
    Q_Mul := THQuery(Ecran.FindComponent('Q'));
    TFMul(Ecran).SetDBListe('PGMULCONTRAT');
    //      SetControlProperty('PSA_ETABLISSEMENT','Name','PCI_ETABLISSEMENT'); // PT7
    //      SetControlProperty('PSA_SALARIE','Name','PCI_SALARIE'); // PT9
    THSalContrat := THEdit(GetControl('PSA_SALARIE')); // PT11
    THVEtabContrat := THValComboBox(GetControl('PSA_ETABLISSEMENT'));
    //      CopyDotInModele('ContratTravail.dot'); PT8 Mise en commentaire regroupé dans RecupDotInDirectoryTemp
    SetControlVisible('CKVOIRTSSAL', True);
  end
  else
    if Arg = 'MEDECINE' then //DEBUT PT3
  begin
    SetControlText('XX_WHERE', '');
    TFMul(Ecran).caption := 'Médecine du travail';
    TFMul(Ecran).SetDBListe('PGVISITEMEDICAL');
    Q_Mul := THQuery(Ecran.FindComponent('Q'));
//    TFMul(Ecran).SetDBListe('PGVISITEMEDICAL');
    St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('PSA_ETABLISSEMENT')+'%") '+
        ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MED%" )  ' ;
        SetControlProperty('DECLARANT', 'Plus' ,St );
    Q:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+   //PT15
                   'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+GetControlText('PSA_ETABLISSEMENT')+'%") '+
                   'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%MED%" )  '+
                   'ORDER BY PDA_ETABLISSEMENT DESC',True);
        If Not Q.eof then
          Begin
          SetControlText('DECLARANT',Q.FindField('PDA_DECLARANTATTES').AsString);
          AffichDeclarant(nil);
          End;
        Ferme(Q);
        Edit := THEdit(GetControl('DECLARANT'));
        If Edit<>Nil Then Edit.OnExit := AffichDeclarant;

            //      CopyDotInModele('MedecineTravail.dot'); PT8 Mise en commentaire regroupé dans RecupDotInDirectoryTemp
  end //FIN PT3
  else
    if Arg = 'FORMATION' then //PT5
  begin
    if (Ordre <> '') and (CodeStage <> '') then //DB2
      SetControlText('XX_WHERE', 'PFO_CODESTAGE="' + CodeStage + '" AND PFO_ORDRE=' + Ordre + ''); //DB2
    SetControltext('CODESTAGE', CodeStage);
    SetControlText('ORDRE', Ordre);
    TFMul(Ecran).caption := 'Convocation formation';
    TFMul(Ecran).SetDBListe('PGCONVOCFORMATION');
    Q_Mul := THQuery(Ecran.FindComponent('Q'));
//    if Q_Mul <> nil then Q_Mul.Liste := 'PGCONVOCFORMATION';
    //          CopyDotInModele('ConvocFormation.dot'); PT8 Mise en commentaire regroupé dans RecupDotInDirectoryTemp
  end;
  UpdateCaption(TFMul(Ecran));
  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  if Defaut <> nil then Defaut.OnExit := ExitEdit;

  Fichier := Arg;
{$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('FLISTE'));
{$ELSE}
  Grille := THGrid(GetControl('FLISTE'));
{$ENDIF}
  if Grille <> nil then
    Grille.OnDblClick := Change;
  // DEB PT13
  BTN := TToolbarbutton97(GetControl('BTNLANCE'));
  if BTN <> nil then BTN.OnClick := Change;
  // FIN PT13
  //DEBUT PT21
  Q := OpenSQL('SELECT YFS_NOM,YFS_PREDEFINI FROM YFILESTD WHERE YFS_BCRIT1="X" AND YFS_CRIT1="MAW" AND YFS_CRIT2="'+Nature+'" AND YFS_CODEPRODUIT="PAIE"',True);
  If Not Q.Eof then
  begin
       SetcontrolText('DOCUMENT',Q.FindField('YFS_PREDEFINI').AsString+Q.FindField('YFS_NOM').AsString);
  end;
  Ferme(Q);
  If V_PGI.ModePcl = '1' then SetControlProperty('DOCUMENT','Plus',' AND YFS_CRIT2="'+Nature+'" AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")')
  else SetControlProperty('DOCUMENT','Plus',' AND YFS_CRIT2="'+Nature+'"');
  //FIN PT21
  // PT20 début
  If AppelProc = True then
  begin
     SetControlVisible('BTNSTOP', TRUE);
     BTNSTOP := TToolbarButton97 (GetControl ('BTNSTOP'));
     if BTNSTOP <> NIL then BTNSTOP.Onclick := ClickSortie;
  end;
  // PT20 fin

end;

procedure TOF_PGSAISIESALRTF.Change(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Grille: THDBGrid;
{$ELSE}
  Grille: THGrid;
{$ENDIF}
  Salarie: TStringList;
  St: string;
  i: integer;
  //TProgress: TQRProgressForm ; //PORTAGECWAS
  doc: string;
  CodeRetour,Long : Integer;
  Crit1,Crit2,NomF : String;
  Q : TQuery;
  Predef,StCombo : String;
  Combo : THvalComboBox;
begin
  Crit1 := 'MAW';
  Crit2 := 'DIV';
  If Fichier = 'CERTIFICAT' then
  begin
       doc := 'CertificatTravail.doc';
       Crit2 := 'CER';
  end;
  if Fichier = 'SOLDE' then
  begin
       doc := 'Soldedetoutcompte.doc';
       Crit2 := 'SOL';
  end;
  if Fichier = 'CONTRAT' then
  begin
       doc := 'ContratTravail.doc';
       Crit2 := 'CNT';
  end;
  if Fichier = 'MEDECINE' then
  begin
       doc := 'TESTMED.doc'; //PT3
       Crit2 := 'MED';
  end;
  if Fichier = 'FORMATION' then
  begin
       doc := 'ConvocForm.doc'; //PT5 et PT10
       Crit2 := 'COF';
  end;
  NomF := GetControlText('DOCUMENT');
  Long := Length(NomF);
  nomF := Copy(NomF,4,Long);
  If GetCheckBoxState('SANSFICHIERBASE') = cbChecked then    //Si coché alors on recherche fichier sur disque
  begin
    fILEp := '';
    if Fichier = 'CERTIFICAT' then doc := 'CertificatTravail.doc';
    if Fichier = 'SOLDE' then doc := 'Soldedetoutcompte.doc';
    if Fichier = 'CONTRAT' then doc := 'ContratTravail.doc';
    if Fichier = 'MEDECINE' then doc := 'MedecineTravail.doc'; //PT3
    if Fichier = 'FORMATION' then doc := 'ConvocForm.doc'; //PT5 et PT10
    { Pour certaines lignes de menu, les doc générés sont définis pas l'appli. }
    if (fichier <> '') and (fichier <> 'CONTRAT') and (fichier <> 'DOCLIBRE') then
    begin
      if VH_Paie.PGCheminRech = '' then
      begin
           PgiBox('Vous devez renseigner le chemin de recherche dans les paramètres société.','Chemin de recherche');
           Exit;
      end;
      if VH_Paie.PGCheminSav = '' then
      begin
           PgiBox('Vous devez renseigner le chemin de sauvegarde dans les paramètres société.','Chemin de sauvegarde');
           Exit;
      end;
      if not FileExists(VH_Paie.PGCheminRech + '\' + doc + '') then
      begin
           PgiBox('Le fichier ' + doc + ' n''existe pas sous le répertoire spécifié.','Fichier Introuvable');
           Exit;
      end;
    end;
  end
  else
  //Si non coché alors utilisation fichier en base PT21
  begin
       Combo := THvalComboBox(GetControl('DOCUMENT'));
       StCombo := Combo.Text;
       Predef := Copy (StCombo,0,3);
       CodeRetour := AGL_YFILESTD_EXTRACT (FileP, 'PAIE', NomF , Crit1, Crit2,'','','',False,'FRA',Predef);
       if fichier <> '' then
       begin
       if not FileExists(fILEp) then
          begin
               PgiBox('Le fichier ' + doc + ' n''existe pas sous le répertoire spécifié.','Fichier Introuvable');
               Exit;
          end;
       end;
  end;
  ToutSelectioner := False;
  Salarie := TStringList.Create;
{$IFNDEF EAGLCLIENT}
  Grille := THDBGrid(GetControl('FLISTE'));
{$ELSE}
  Grille := THGrid(GetControl('FLISTE'));
{$ENDIF}
  if (Grille <> nil) then
    if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

  if (Grille <> nil) then
    if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
    begin
      InitMoveProgressForm(nil, 'Chargement des données en cours...', 'Veuillez patienter SVP...', Grille.NbSelected, FALSE, TRUE);
      InitMove(Grille.NbSelected, '');
      for i := 0 to Grille.NbSelected - 1 do
      begin
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        Grille.GotoLeBOOKMARK(i);
        if (Fichier = 'CONTRAT') and (GetControlText('CKVOIRTSSAL') = '-') then st := TFmul(Ecran).Q.FindField('PCI_SALARIE').asstring
        else
          if Fichier = 'MEDECINE' then st := TFmul(Ecran).Q.FindField('PVM_SALARIE').asstring //PT3
        else
          if Fichier = 'FORMATION' then St := TFmul(Ecran).Q.FindField('PFO_SALARIE').asstring //PT5
        else St := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
        if st <> '' then Salarie.add(st);
        MoveCur(False);
        MoveCurProgressForm(st);
      end;
      Grille.ClearSelected;
      FiniMove;
      FiniMoveProgressForm;
    end;

  if (Grille.AllSelected = TRUE) then
  begin
    ToutSelectioner := True;
    InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
{$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
    TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.EOF do
    begin
{$IFDEF EAGLCLIENT}
      //    TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
      MoveCur(False);
      if (Fichier = 'CONTRAT') and (GetControlText('CKVOIRTSSAL') = '-') then st := TFmul(Ecran).Q.FindField('PCI_SALARIE').asstring
      else
      begin
        if Fichier = 'MEDECINE' then //PT3
        begin
          st := TFmul(Ecran).Q.FindField('PVM_SALARIE').AsString;
          // MAJ DATE CONVOCATION
          ExecuteSQL('UPDATE VISITEMEDTRAV SET PVM_EDITCONVOC="X",PVM_DATECONVOC="' + UsDateTime(Date) + '" WHERE' +
            ' PVM_SALARIE="' + St + '"');
        end
        else if Fichier = 'FORMATION' then St := TFmul(Ecran).Q.FindField('PFO_SALARIE').asstring //PT5
        else St := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
      end;
      if st <> '' then Salarie.add(st);
      MoveCurProgressForm(st);
      TFmul(Ecran).Q.Next;
    end;
    Grille.AllSelected := False;
    FiniMove;
    FiniMoveProgressForm;
  end;

  PGPathDocContrat := '';
  if Salarie.Count > 0 then
  begin
     If GetCheckBoxState('SANSFICHIERBASE') = CbChecked then
     begin
         if (fichier = 'CONTRAT') or (Fichier = 'DOCLIBRE') then
         begin
              AglLanceFiche('PAY', 'LISTEDOC', '', '', doc);
              if PGPathDocContrat = '' then exit;
         end;
     end;
    InitTob_Fusion(Salarie);
    LanceFusionWord(GetControlText('OBJET'), PGPathDocContrat, Salarie, T_Sal, T_Net, T_Cal, T_Contrat, T_Soc, T_Conv, T_Org, T_VisiteMed, T_AdrMed, T_ConvocForm,T_Annuaire); { PT17 }
  end;
  if Salarie <> nil then Salarie.Free;
  // PT20
  If AppelProc = True then
  Ecran.Close;
end;


procedure TOF_PGSAISIESALRTF.OnLoad;
var
  DateVisiteD, DateVisiteF: TDateTime;
  LePrefixe : String;
  Where2 : THEdit;
begin
  inherited;
  if Fichier = 'MEDECINE' then //PT3 et PT4 mis en commentaire
  begin
    DateVisiteD := StrToDate(GetControlText('DATEVISITE'));
    DateVisiteF := StrToDate(GetControlText('DATEVISITE_'));
    SetControlText('XX_WHERE', ' PVM_DATEVISITE>="' + UsDateTime(DateVisiteD) + '" AND PVM_DATEVISITE<="' + UsDateTime(DateVisiteF) + '"');
  end
  else
    if Fichier = 'SOLDE' then
  begin
    if GetControlText('CKVOIRTSSAL') = '-' then
      // PT19 SetControlText('XX_WHERE', 'PSA_DATESORTIE IS NOT NULL AND PSA_DATESORTIE<>"' + UsdateTime(Idate1900) + '" ' +
      // PT19  'AND PSA_SALARIE IN (SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS)')
      // PT19 début
      begin
      SetControlText('XX_WHERE', 'PSA_DATESORTIE IS NOT NULL AND PSA_DATESORTIE<>"' + UsdateTime(Idate1900) + '" ' +
        'AND PSA_SALARIE IN (SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS)');
      // PT20 If AppelProc = True then
      If (AppelProc = True) and (SelectSal <> '') then
			SetControlText('XX_WHERE',SelectSal)
      end
      // PT19 fin
    else
      begin   // PT19
      SetControlText('XX_WHERE', '');
      // PT19 Début
      // PT20 If AppelProc = True then
      If (AppelProc = True) and (SelectSal <> '') then
        begin
			  SetControlText('XX_WHERE',SelectSal);
        end;
      // PT19 fin
    end  // PT19
  end
  else
    if Fichier = 'CONTRAT' then
    begin
      //DEB PT23
      if GetControlText('CKVOIRTSSAL') = '-' then
      begin
        LePrefixe := 'PCI';
        THSalContrat.Name := 'PCI_SALARIE'; //  PT11
        THVEtabContrat.Name := 'PCI_ETABLISSEMENT';
      end
      else
      begin
        LePrefixe := 'PSA';
        THSalContrat.Name := 'PSA_SALARIE';
        THVEtabContrat.Name := 'PSA_ETABLISSEMENT';
      end;
      Where2 := THEdit(GetControl('XX_WHERE2'));
      if Assigned(MonHabilitation) then
        if (LePrefixe <> copy(MonHabilitation.LeSQL, 1, 3)) and (MonHabilitation.LeSQL <> '') then
        begin
          SetControlText('LEPREFIXE', LePrefixe);
          MonHabilitation.LeSQL := '';
          PGAffecteEtabByUser(TFMul(Ecran));
        end;

      if Where2 <> nil then SetControlText('XX_WHERE2', MonHabilitation.LeSQL);

      if GetControlText('CKVOIRTSSAL') = '-' then
        TFMul(Ecran).SetDBListe('PGMULCONTRAT')
      else
        TFMul(Ecran).SetDBListe('PGMULSALARIE');

      MonHabilitation.AvecCreat := 'N';
      //FIN PT23
    end;
  // PT19 début
  // PT20 If AppelProc = True then
   If (AppelProc = True) and (SelectSal <> '') then
   begin
    TFMul(Ecran).FListe.AllSelected := True;
  end;
  // PT19 fin
end;

procedure TOF_PGSAISIESALRTF.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

//PT1 Nouvelle procédure

procedure TOF_PGSAISIESALRTF.CopyDotInModele(StNomModele: string);
var
  stRepDest, stRepOrig: string;
begin
  stRepDest := GetWordTemplateDirectory;
{$IFNDEF EAGLCLIENT}
  //il s'agit des répertoire d'installation des doc et dot du kit!!
  stRepOrig := V_PGI.StdPath + '\Modèles';
  if not DirectoryExists(stRepOrig) then
    if not CreateDir(stRepOrig) then
    begin
      PGIBox('Impossible de créer le répertoire modèles ' + stRepOrig + '.', 'Abandon de la copie');
      Exit;
    end;
  if stRepDest = '' then
  begin
    PGIBox('Le répertoire modèles WORD n''est pas défini sur votre machine!', 'Abandon de la copie');
    exit;
  end;
  //Copie du fichier modele Kit de l'ancien répertoire sur le nouveau défini
  if {PT6 Mise en commentaire(FileExists(stRepOrig+'\'+stNomModele)=False) and}(FileExists(V_Pgi.StdPath + '\' + stNomModele) = True) then
  begin
    CopyFile(PChar(V_Pgi.StdPath + '\' + stNomModele), PChar(stRepOrig + '\' + stNomModele), TRUE);
    DeleteFile(PChar(V_Pgi.StdPath + '\' + stNomModele));
  end
  else
    if (FileExists(stRepOrig + '\' + stNomModele) = False) and (FileExists(V_Pgi.StdPath + '\' + stNomModele) = False) then
  begin
    PgiBox('Attention le modèle ' + stNomModele + ' n''est pas installé sous ' + stRepOrig + '.', 'Abandon de la copie');
    Exit;
  end;
{$ELSE}
  { DEB PT8 }
  stRepOrig := ChangeStdDatPath('$STD\');
  stRepOrig := ExtractFileDir(stRepOrig);
  { FIN PT8 }
{$ENDIF}

  if FileExists(stRepOrig + '\' + stNomModele) then
    CopyFile(PChar(stRepOrig + '\' + stNomModele), PChar(stRepDest + '\' + stNomModele), TRUE);
end;

procedure TOF_PGSAISIESALRTF.InitTob_Fusion(ListSal: TStringList);
var
  QSalaries, QPaie, Q: TQuery;
  i: integer;
  St, StSal, StWhere: string;
  WhereAuto : Boolean;  { PT18 }
begin
  WhereAuto := (ToutSelectioner = False) Or (ListSal.Count < 31);  { PT18 }
  if WhereAuto then  { PT18 }
    StSal := RendStringListSal('PSA', ListSal)
  else
    StSal := '';
  // DEB PT12
  if StSal <> '' then StWhere := 'WHERE ' + Copy(StSal, Pos('AND', StSal) + 3, Length(StSal))
  else stWhere := '';

  St := 'SELECT SALARIES.*,ET_ETABLISSEMENT,ET_LIBELLE,ET_ADRESSE1,' +
    'ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,ET_DIVTERRIT,ET_PAYS,' +
    'ET_LANGUE,ET_TELEPHONE,ET_FAX,ET_TELEX,ET_ETABLIE,ET_EMAIL,ET_SIRET,' +
    'ET_APE,ET_ACTIVITE,ET_ABREGE,ET_JURIDIQUE,ETB_HORAIREETABL '+    { PT14 }
    'FROM SALARIES ' +
    'LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    StWhere + ' ORDER BY PSA_SALARIE'; //StWhere+' '+
  // FIN PT12
  QSalaries := OpenSql(St, TRUE);
  T_Sal := TOB.Create('Salariè concerné', nil, -1);
  T_Sal.LoadDetailDB('SALARIES', '', '', QSalaries, False);
  Ferme(QSalaries);
  //initialisation du champ psa_e pour les accords au féminin
  for i := 0 to T_Sal.detail.count - 1 do
  begin
    T_Sal.detail[i].AddChampSup('PSA_E', TRUE);
    T_Sal.detail[i].PutValue('PSA_E', GetESexe(T_Sal.detail[i].GetValue('PSA_SEXE')));
  end;

  //Init Tob convention collective
  St := 'SELECT PSA_SALARIE,PSA_CONVENTION,PCV_CONVENTION,' +
    'PCV_LIBELLE,PCV_LIBELLE1,PCV_LIBELLE2,PCV_LIBELLE3 ' +
    'FROM SALARIES ' +
    'LEFT JOIN CONVENTIONCOLL ON PSA_CONVENTION=PCV_CONVENTION ' +
    'WHERE ##PCV_PREDEFINI## PSA_CONVENTION<>"" ' +
    StSal + ' ' +
    'ORDER BY PSA_SALARIE';
  Q := OpenSql(St, TRUE);
  T_Conv := TOB.Create('Convention collective', nil, -1);
  T_Conv.LoadDetailDB('CONVENTIONCOLL', '', '', Q, False);
  Ferme(Q);
  //Init Tob contrat de travail
  if WhereAuto then   { PT18 }
    StSal := RendStringListSal('PCI', ListSal);
  St := 'SELECT * FROM CONTRATTRAVAIL CT ' +
    'WHERE PCI_ORDRE=(SELECT MAX(PCI_ORDRE) FROM CONTRATTRAVAIL ' +
    'WHERE PCI_SALARIE=CT.PCI_SALARIE GROUP BY PCI_SALARIE) ' + StSal;
  Q := OpenSql(St, TRUE);
  T_Contrat := TOB.Create('Contrat de travail', nil, -1);
  T_Contrat.LoadDetailDB('CONTRATTRAVAIL', '', '', Q, False);
  Ferme(Q);
  //Init Tob Organismes et caisses
  St := 'SELECT * FROM ORGANISMEPAIE '; {WHERE POG_ORGANISME="001" ' +  PT16 Suppression clause
  'ORDER BY POG_ETABLISSEMENT,POG_ORGANISME'};
  Q := OpenSql(St, TRUE);
  T_Org := TOB.Create('Chgmt Organisme Urssaf', nil, -1);
  T_Org.LoadDetailDB('ORGANISMEPAIE', '', '', Q, False);
  Ferme(Q);

  {St := 'SELECT * FROM ORGANISMEPAIE ';
  Q := OpenSql(St, TRUE);
  T_AutreOrg := TOB.Create('Chgmt Organisme', nil, -1);
  T_AutreOrg.LoadDetailDB('ORGANISMEPAIE', '', '', Q, False);
  Ferme(Q); }

  //Init Tob Société                                     //DEB PT2-2
  St := 'SELECT SOC_NOM,SOC_DATA FROM PARAMSOC ' +
    'WHERE (SOC_NOM="SO_SOCIETE" OR SOC_NOM="SO_LIBELLE" ' +
    'OR SOC_NOM="SO_ADRESSE1" OR SOC_NOM="SO_ADRESSE2" ' +
    'OR SOC_NOM="SO_ADRESSE3" OR SOC_NOM="SO_CODEPOSTAL" ' +
    'OR SOC_NOM="SO_VILLE" OR SOC_NOM="SO_DIVTERRIT" OR SOC_NOM="SO_PAYS" ' +
    'OR SOC_NOM="SO_TELEPHONE" OR SOC_NOM="SO_FAX" OR SOC_NOM="SO_TELEX" ' +
    'OR SOC_NOM="SO_NIF" OR SOC_NOM="SO_APE" OR SOC_NOM="SO_SIRET" ' +
    'OR SOC_NOM="SO_CONTACT" OR SOC_NOM="SO_CAPITAL" OR SOC_NOM="SO_RC")';
  Q := OpenSql(St, True); //FIN PT2-2
  T_Soc := TOB.Create('Société', nil, -1);
  T_Soc.LoadDetailDB('SOCIETE', '', '', Q, False);
  Ferme(Q);
  //initialisation d'une tob pour le Calendrier
  QPaie := OpenSql('SELECT * FROM CALENDRIER ', true);
  T_Cal := TOB.Create('Calendrier', nil, -1);
  T_Cal.LoadDetailDB('CALENDRIER', '', '', QPaie, False);
  Ferme(QPaie);
  //initialisation d'une tob pour le solde du net à payer : Dernier net à payer des bulletins
  if WhereAuto then  { PT18 }
  begin
    StSal := RendStringListSal('PPU', ListSal);
    if StSal <> '' then StSal := 'WHERE ' + Copy(StSal, Pos('AND', StSal) + 3, Length(StSal));
  end;
  St := 'SELECT DISTINCT PPU_SALARIE,MAX(PPU_DATEDEBUT) AS DATEDEBUT,' +
    'MAX(PPU_DATEFIN) AS DATEFIN,PPU_CNETAPAYER,PPU_PGMODEREGLE ' + {PT1}
  'FROM PAIEENCOURS ' +
    StSal + ' ' +
    'GROUP BY PPU_SALARIE,PPU_PGMODEREGLE,PPU_CNETAPAYER ' + {PT1}
  'ORDER BY PPU_SALARIE,DATEDEBUT,DATEFIN ';
  QPaie := OpenSql(st, True);
  T_Net := TOB.Create('Paie en cours', nil, -1);
  T_Net.LoadDetailDB('PAIEENCOURS', '', '', QPaie, False);
  Ferme(QPaie);
  //Init Tob Médecine du tarvail        //PT3
  StSal := RendStringListSal('PVM', ListSal);
  St := 'SELECT PVM_SALARIE,PVM_DATEVISITE,PVM_TYPEVISITMED,PVM_MEDTRAVGU,PVM_HEUREVISITE FROM VISITEMEDTRAV WHERE ' + //PT22
    ' PVM_DATEVISITE>="' + UsDateTime(StrToDate(GetControltext('DATEVISITE'))) + '"' +
    ' AND  PVM_DATEVISITE<="' + UsDateTime(StrToDate(GetControltext('DATEVISITE_'))) + '" ' + StSal; //PT4

  Q := OpenSql(St, TRUE);
  T_VisiteMed := TOB.Create('Visite médecine', nil, -1);
  T_VisiteMed.LoadDetailDB('VISITEMEDTRAV', '', '', Q, False);
  Ferme(Q);
  If GetControlText('OBJET') = 'MEDECINE' then  //PT15
  begin
       For i := 0 to T_VisiteMed.Detail.Count - 1 do
       begin
            T_VisiteMed.Detail[i].AddChampSupValeur('PVM_LIEU',GetControlText('LIEUMEDECINE'));
       end;
  end;
  // Tob annuaire          //PT3
  St := 'SELECT ANN_GUIDPER,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE FROM ANNUAIRE' +
    ' WHERE ANN_TYPEPER="MED"';
  //' WHERE ANN_GUIDPER="'+TFMul(Ecran).Q.FindField('PVM_MEDTRAV').AsString+'"';
  Q := OpenSql(St, TRUE);
  T_AdrMed := TOB.Create('Adresse médecine', nil, -1);
  T_AdrMed.LoadDetailDB('ANNUAIRE', '', '', Q, False);
  Ferme(Q);
  // Tob Formation PT5
  StSal := RendStringListSal('PFO', ListSal);
  // DEBUT DB2
  if (GetControlText('ORDRE') <> '') and (GetControlText('CODESTAGE') <> '') then St := 'SELECT PFO_CODESTAGE,PFO_DATEDEBUT,' +
    'PFO_DATEFIN,PFO_SALARIE,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_NATUREFORM,PFO_LIEUFORM,' +
      'PFO_FORMATION1,PFO_FORMATION2,PFO_FORMATION3,PFO_FORMATION4,PFO_FORMATION5,PFO_FORMATION6,PFO_FORMATION7,PFO_FORMATION8' +
      ' FROM FORMATIONS WHERE ' +
      ' PFO_CODESTAGE="' + GetControltext('CODESTAGE') + '"' +
      ' AND  PFO_ORDRE=' + GetControltext('ORDRE') + ' ' + StSal
  else
  begin
    if StSal <> '' then StSal := ' WHERE ' + Copy(StSal, Pos('AND', StSal) + 3, Length(StSal));
    St := 'SELECT PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_SALARIE,PFO_HEUREDEBUT,PFO_HEUREFIN,PFO_NATUREFORM,PFO_LIEUFORM,' +
      'PFO_FORMATION1,PFO_FORMATION2,PFO_FORMATION3,PFO_FORMATION4,PFO_FORMATION5,PFO_FORMATION6,PFO_FORMATION7,PFO_FORMATION8 ' +
      'FROM FORMATIONS' + StSal;
  end;
  //FIN DB2
  Q := OpenSql(St, TRUE);
  T_ConvocForm := TOB.Create('Les formations', nil, -1);
  T_ConvocForm.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);

  { DEB PT17 }
  St := 'SELECT ANNUAIRE.* FROM ANNUAIRE, DOSSIER '+
  'WHERE ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="'+V_PGI.NoDossier+'"';
  Q := OpenSql(St, TRUE);
  T_Annuaire := TOB.Create('Annuaire', nil, -1);
  T_Annuaire.LoadDetailDB('ANNUAIRE', '', '', Q, False);
  Ferme(Q);
  { FIN PT17 }


end;

function TOF_PGSAISIESALRTF.GetESexe(Sexe: string): string;
begin
  if Sexe = 'M' then result := '' else result := 'e';
end;

procedure TOF_PGSAISIESALRTF.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;
{ DEB PT8 }
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 03/10/2003
Modifié le ... : 03/10/2003
Description .. : Fonction compatible AGL / EAGL
Suite ........ : Kit installe les fichiers .doc & .dot sur PGI00\STD\
Suite ........ : Arborescence Agl :
Suite ........ : - Copier/Coller des .dot dans répertoire Word Modeles
Suite ........ : - Couper/Coller des .dot dans répertoire STD\Modeles
Suite ........ : Arborescence EAgl:
Suite ........ : - Copier/Coller des .doc dans répertoire Temp\STD
Suite ........ : - Copier/Coller des .dot dans répertoire Word Modeles
Suite ........ : - Copier/Coller des .doc dans répertoire de recherche défini
Suite ........ : dans paramsoc
Mots clefs ... : PAIE;FUSIONWORD
*****************************************************************}

procedure TOF_PGSAISIESALRTF.RecupDotInDirectoryTemp;
var
//  sr: TSearchRec;
//  FileAttrs: Integer;
  stPath: string;
  docList: TStringList;
  i: Integer;
  {$IFDEF EAGLCLIENT}
  stRepOrig : String;
  {$ENDIF}
const
  ChaineEntier = ['0'..'9'];
begin
//  FileAttrs := faAnyFile;
//  docList := TStringList.Create;
  docList := ListStdDatDirectory('$STD', '*.DOT');
  for i := 0 to docList.Count - 1 do
  begin
    stPath := ChangeStdDatPath('$STD\' + docList.Strings[i], true);
    CopyDotInModele(docList.Strings[i]);
  end;
  docList := ListStdDatDirectory('$STD', '*.DOC');
  for i := 0 to docList.Count - 1 do
  begin
    stPath := ChangeStdDatPath('$STD\' + docList.Strings[i], true);
{$IFDEF EAGLCLIENT}
    stRepOrig := ExtractFileDir(stPath);
    if (FileExists(stRepOrig + '\' + docList.Strings[i])) and (stRepOrig <> VH_Paie.PGCheminRech) then
      CopyFile(PChar(stRepOrig + '\' + docList.Strings[i]), PChar(VH_Paie.PGCheminRech + '\' + docList.Strings[i]), TRUE);
{$ENDIF}
  end;
  docList.free;
end;
{ FIN PT8 }

procedure TOF_PGSAISIESALRTF.AffichDeclarant(Sender: TObject);  //PT15
begin
if GetControlText('DECLARANT')='' then exit;
SetControlText('LIEUMEDECINE'       ,RechDom('PGDECLARANTVILLE',GetControlText('DECLARANT'),False));
end;

// PT20 bouton gestion arret du processus
procedure TOF_PGSAISIESALRTF.ClickSortie(Sender: TObject);
var BTNAnn :  TToolbarButton97;
begin
  BTNAnn := TToolbarButton97 (GetControl ('BAnnuler'));
  if BTNAnn <> NIL then
  begin
    TFMul(Ecran).Retour := 'STOP';
    BTNAnn.Click;
  end;
end;

procedure TOF_PGSAISIESALRTF.voirDocDefaut (Sender: TObject);
var Crit2,DocDefaut : String;
begin
  If Fichier = 'CERTIFICAT' then Crit2 := 'CER'
  else if Fichier = 'SOLDE' then Crit2 := 'SOL'
  else if Fichier = 'CONTRAT' then Crit2 := 'CNT'
  else if Fichier = 'MEDECINE' then Crit2 := 'MED'
  else if Fichier = 'FORMATION' then Crit2 := 'COF';
  DocDefaut := AglLanceFiche('PAY', 'RTFDEFAUT', '', '', Crit2);
  If DocDefaut <> '' then SetControlText('DOCUMENT',DocDefaut);

end;

procedure TOF_PGSAISIESALRTF.VoirImport (Sender: TObject);
begin
     AglLanceFiche('PAY', 'IMPORTRTF', '', '', '');
end;

initialization
  registerclasses([TOF_PGSAISIESALRTF]);
end.

