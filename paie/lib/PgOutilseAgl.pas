{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/06/2001
Modifié le ... :   /  /
Description .. : Comprend des fonctions utilitaires à l'import/export des
Suite ........ : données pour une saisie décentalisée
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************}
{PT1 14/06/01 PH Rajout export des 2 libellés complémentaires
             Mise à jour exportok quand les absences sont déportées
PT2 09/11/01 Rajout paramètre pour savoir qui utilise la fonction
             EAbsences ou NETSERVICE
PT3 26/11/01 Rajout barre de patience et message de traitement
PT4 26/06/2002 SB V582 On récupère les mvts d'acquis posterieur au dernier SLD
                       Pour calculer le récapitulatif qu'à partir du dernier contrat
PT5 11/07/2002 V582 SB Retrait des salariés sorties pour le calcul du recap
PT6 19/07/2002 V582 SB Violation d'acces export table dû à une requête non fermé
PT7 16/09/2002 SB V585 FQ n° 10192 Intégration des cumuls RTT : contrôle de vraisemblance
PT8 04/11/2002 SB V585 Intégration du calcul des compteurs Abs en cours
PT9-1 28/11/2002 SB V591 Anomalie traitement dans l'utilisation de la date système,
                         utilisation d'une date d'intégration saisissable par l'utilisateur
PT9-2 28/11/2002 SB V591 Traitement par salarié du calcul du récap
PT9-3 28/11/2002 SB V591 Modification de l'etat des absences à exporter dans la base de production
PT10  08/01/2003 SB V585 Calcul de la date de RAZ des cumuls érroné si mois défini>mois édité
PT11  15/01/2004 SB V591 FQ10447 Gestion multicombo des paramètres soc RTT Acquis Pris
PT12  05/02/2003 SB V591 Reprise des cumuls au debut du mois donné
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT13     V_42 08/10/2003 SB Econges Spéc. CEGID Gestion des abences
PT14  03/05/2004 SB V_50 FQ 11280 Modif champ TYPERTT => TYPEABS
PT15  08/10/2004 PH V_50 Suppression variable globale inutile d'autant qu'il en existe une autre avec le *
                         même nom
PT16 02/06/2005 SB V_60 FQ 12327 Econges : Calcul récap pour un salarié
PT17 07/07/2005 SB V_65 FQ 12434 Erreur SQL Oracle
PT18 25/07/2005 SB V_65 FQ 12399 Gestion Econges des assistances
PT19 23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT20 19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT21 19/06/2006 SB V_71 FQ 13690 Reprise erronée des mvts clôturés
PT22 31/01/2007 FC V_80 Prise en compte de l'habilitation dans le lookuplist salarié
}
unit PgOutilseAgl;

interface
uses SysUtils, HCtrls, HEnt1,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Db,
  {$ENDIF}  
  Controls, HMsgBox,  UTOB, LookUp,
  ed_tools,
  PgOutils //PT22
  ,ULibEditionPaie
  ;


// fonction de chargement et de calcul dans une tob de la table RECAPSAL
procedure ChargeTobSalRecap(LeType: string; DateIntegration: TDateTime = 2; Salarie: string = '');
// Extraction Base eAgl des absences déportées saisies et validées et écriture
procedure RecupAbsFromeAgl(LaDate: string);
// Calcul cumul CP et RTT sur différentes périodes
procedure trouveCumRtt(var Cum45, cum46: double; salarie, etab: string; DateIntegration: TDateTime; DateDebutExerSoc: tdatetime);
procedure AfficheTabSalResp(Sender: Tobject; Matricule: string; Quoi: string = '');
// Recalcul des compteurs en cours CP et RTT
procedure CalculRecapAbsEnCours(Salarie: string);
function IfMotifabsenceSaisissable(TypeConge, FicPrec: string): Boolean; //PT13




implementation
uses PgCongespayes,PgOutils2,EntPaie;

var
  TobSalRecap, tobetab: tob;
//  DateDebutExerSoc: tdatetime;

  // PgOutilseAgl-1 Déchargement ABSENCESALARIE ds fichier RECAPSAL
  // PT2 09/11/01 Rajout paramètre pour savoir qui utilise la fonction
procedure ChargeTobSalRecap(LeType: string; DateIntegration: TDateTime = 2; Salarie: string = '');
var
  TS, T, TobSal: tob;
  st, Periode, Etab: string;
  Q: TQuery;
  DSolde, FinExer: TdateTime;
  PN, AN, RN, PN1, AN1, RN1, Base, Mois, RTTA, RTTP: double;
  //TProgress                   : TQRProgressForm ; //PORTAGECWAS
  i: Integer;
  MoisE, AnneeE, ComboExer: string;
  DateDebutExerSoc: tdatetime; // PT15
begin
  if DateIntegration = 2 then DateIntegration := Date; //PT9-1
  //DEB PT7 Contrôle de vraisemblance
  if (VH_Paie.PgRttAcquis = '') or (VH_Paie.PgRttPris = '') then
  begin
    PgiBox('Abandon du traitement.#13#10Veuiller renseigner le cumul Rtt acquis et pris dans les paramètres sociètes.', 'Gestion des absences');
    Exit;
  end;
  //FIN PT7
  //RAZ = Date de debut d'exercice social
  (*DEB PT9-1 Mise en commentaire : utilisation de la function RendExerSocialEnCours
    QExer:=OpenSql('SELECT PEX_DATEDEBUT FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="'+USDateTime(Date)+'" '+
    'AND PEX_DATEFIN>="'+USDateTime(Date)+'"' ,True);
    if Not QExer.Eof then  //PORTAGECWAS
      DateDebutExerSoc:=QExer.FindField('PEX_DATEDEBUT').AsDateTime
    else
      DateDebutExerSoc:=idate1900;
    Ferme(QExer);*)
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DateDebutExerSoc, FinExer);
  (*FIN PT9-1*)
  tobetab := nil;
  TobEtab := TOB.Create('ETABCOMPL', nil, -1);
  TobEtab.LoadDetailDB('ETABCOMPL', '', '', nil, FALSE);

  TobSalRecap := nil;
  TobSalRecap := Tob.Create('La mère RECAP', nil, -1);

  if TobSalRecap = nil then
  begin
    PGIBOX('Echec de la récupération des cumuls CP', 'récupération des cumuls CP');
    exit;
  end;
  TobSalRecap.addchampsup('PRS_SALARIE', false);
  TobSalRecap.addchampsup('PRS_LIBELLE', false);
  TobSalRecap.addchampsup('PRS_PRENOM', false);
  TobSalRecap.addchampsup('PRS_ACQUISN1', false);
  TobSalRecap.addchampsup('PRS_PRISN1', false);
  TobSalRecap.addchampsup('PRS_RESTN1', false);
  TobSalRecap.addchampsup('PRS_ACQUISN', false);
  TobSalRecap.addchampsup('PRS_PRISN', false);
  TobSalRecap.addchampsup('PRS_RESTN', false);
  TobSalRecap.addchampsup('PRS_CUMRTTACQUIS', false);
  TobSalRecap.addchampsup('PRS_CUMRTTPRIS', false);
  TobSalRecap.addchampsup('PRS_CUMRTTREST', false); //PT8
  TobSalRecap.addchampsup('PRS_DATEMODIF', false);
  TobSal := Tob.create('SALARIES', nil, -1);
  // rajouter controle de la date de sortie par rapport à la période de paie
  // PT2 09/11/01 traitement Netservice ou Eabsences
  if Salarie <> '' then St := ' AND PSA_SALARIE="' + Salarie + '" ' else st := ''; //PT9-2 Traitement par salarié PT16
  if LeType <> 'N' then
    st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM FROM SALARIES'
      + ' WHERE EXISTS (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE=PSA_SALARIE) ' + St //PT9-2 PT17
    { + ' WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+DateToStr(idate1900)+'" ' //PT5
     + ' OR PSA_DATESORTIE>=(SELECT MAX(PPU_DATEFIN) FROM PAIEENCOURS))'+ }
    + ' ORDER BY PSA_SALARIE '
  else
    st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM FROM SALARIES'
      {      + ' WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+DateToStr(idate1900)+'" '  //PT5
          + ' OR PSA_DATESORTIE>=(SELECT MAX(PPU_DATEFIN) FROM PAIEENCOURS)) '}
    + ' ORDER BY PSA_SALARIE ';

  Q := OpenSQL(St, true);
  if Q.eof then
  begin
    Ferme(Q);
    exit;
  end;
  i := Q.RecordCount;
  InitMoveProgressForm(nil, 'Calcul des récapitulatifs', 'Veuillez patienter SVP ...', i, FALSE, TRUE);

  TobSal.Loaddetaildb('SALARIES', '', '', Q, False);
  Ferme(Q); //PT6 suite motif PT4 on ferme la première requête avant la boucle
  TS := TobSal.findfirst([''], [''], true);

  while TS <> nil do
  begin
    Periode := '0';
    MoveCurProgressForm('Salarié : ' + TS.getvalue('PSA_SALARIE'));
    {DEB PT4 On récupère les mvts d'acquis posterieur au dernier SLD
    Pour calculer le récapitulatif qu'à partir du dernier contrat}
    Q := OpenSql('SELECT PCN_DATEVALIDITE FROM ABSENCESALARIE ' +
      'WHERE PCN_SALARIE="' + TS.getValue('PSA_SALARIE') + '" ' + //PT9-1 Anomalie Tob mère utilisée au lieu de la fille
      'AND PCN_TYPECONGE="SLD" AND PCN_TYPEMVT="CPA" ORDER BY PCN_DATEVALIDITE DESC', TRUE);
    if (not Q.eof) then
      DSolde := Q.FindField('PCN_DATEVALIDITE').AsDateTime
    else
      DSolde := iDate1900;
    Ferme(Q);
    {FIN PT4}//DEB PT9-1 utilisation dateintegration
    AffichelibelleAcqPri(Periode, TS.getvalue('PSA_SALARIE'), DSolde, DateIntegration, PN, AN, RN, Base, Mois, False, False);
    Periode := '1';
    AffichelibelleAcqPri(Periode, TS.getvalue('PSA_SALARIE'), DSolde, DateIntegration, PN1, AN1, RN1, Base, Mois, False, False);
    Etab := TS.getvalue('PSA_ETABLISSEMENT');
    trouveCumRtt(RTTA, RTTP, TS.getvalue('PSA_SALARIE'), Etab, DateIntegration, DateDebutExerSoc); //FIN PT9-1 // PT15
    T := Tob.Create('RECAPSALARIES', TobSalRecap, -1);
    T.PutValue('PRS_SALARIE', TS.getvalue('PSA_SALARIE'));
    T.PutValue('PRS_LIBELLE', TS.getvalue('PSA_LIBELLE'));
    T.PutValue('PRS_PRENOM', TS.getvalue('PSA_PRENOM'));
    T.PutValue('PRS_ACQUISN1', AN1);
    T.PutValue('PRS_PRISN1', PN1);
    T.PutValue('PRS_RESTN1', RN1);
    T.PutValue('PRS_ACQUISN', AN);
    T.PutValue('PRS_PRISN', PN);
    T.PutValue('PRS_RESTN', RN);
    T.PutValue('PRS_CUMRTTACQUIS', RTTA);
    T.PutValue('PRS_CUMRTTPRIS', RTTP);
    T.PutValue('PRS_CUMRTTREST', Arrondi(RTTA - RTTP, 2)); //PT8
    //   T.PutValue('PRS_DATEMODIF'   , DateIntegration); //PT9-1 maj par l'agl
    TS := TobSal.findNext([''], [''], true);
  end; // fin de la boucle salariés
  MoveCurProgressForm('Sauvegarde en cours des récapitulatifs');
  if (VH_PAIE.PGEcabMonoBase) or (not VH_PAIE.PGEcabBaseDeporte) then
  begin //pour calcul monobase ou base production
    try
      beginTrans;
      if Salarie <> '' then St := ' WHERE PRS_SALARIE="' + Salarie + '" ' else st := '';
      ExecuteSql('DELETE FROM RECAPSALARIES' + St);
      for i := 0 to TobSalRecap.detail.count - 1 do
        TobSalRecap.detail[i].InsertOrUpdateDB;
      CommitTrans;
    except
      Rollback;
    end;
  end;
  if (not VH_PAIE.PGEcabBaseDeporte) then
    TobSalRecap.SaveToFile(VH_Paie.PGCHEMINEAGL + '\recapsal.txt', False, True, True, '');
  FiniMoveProgressForm;

  if Tobsal <> nil then Tobsal.Free;
  TobSalRecap.free;
  TobEtab.free;
end;

{ Souad ! Couper la procedure ici. Le fichier ici crée est temporaire. Utiliser un
 fichier permanent.

 Bid   := TBidon.Create;
 Bid.ExporteRecapSalaries;
 Bid.free;

 PGIBOX('Récupération des cumuls CP réalisée', 'récupération des cumuls CP');
 end;
}

////////////////////////////////////////////////////////////////////////////////
// Extraction Base eAgl des absences déportées saisies et validées et écriture
// ds fichier export avec codeenr = MAB
procedure RecupAbsFromeAgl(LaDate: string);
var
  TobABS, TA: tob;
  Fichier: Textfile;
  TempFName, st, NoDoss: string;
  Q: TQuery;
  MoisE, AnneeE, ComboExer: string;
  DebExer, FinExer: tdatetime;
  jjd, mmd, aad, jjf, mmf, aaf: word;
  reponse, Nbre: Integer;
  //TProgress                   : TQRProgressForm ;//PORTAGECWAS
begin
  st := 'Attention, l''export des absences validées est irréversible.' +
    '#13#10 Vous devez avoir fait au préalable une sauvegarde de la base.' +
    '#13#10 Voulez-vous continuer?';
  reponse := PGIAsk(st, 'Export des absences');
  if reponse <> 6 then exit;
  // PT3 26/11/01 Rajout barre de patience et message de traitement
  Q := OpenSQL('SELECT count (*) FROM ABSENCESALARIE WHERE PCN_SAISIEDEPORTEE = "X" ' +
    ' AND PCN_VALIDRESP = "VAL" AND PCN_EXPORTOK <> "X" AND PCN_DATEFINABS <="' + USDATETIME(StrToDAte(LaDate)) + '"', true);
  if Q.eof then
  begin
    PgiBox('Vous n''avez aucune absence à exporter.', 'Exportation des absences');
    Ferme(Q);
    exit;
  end;
  Nbre := Q.Fields[0].AsInteger;
  Ferme(Q);
  InitMoveProgressForm(nil, 'Exportation des absences dans le fichier ' + 'AbsencesValid.txt', 'Veuillez patienter SVP ...', Nbre + 3, FALSE, TRUE);

  TobABS := Tob.create('ABSENCESALARIE', nil, -1);
  Q := OpenSQL('SELECT * FROM ABSENCESALARIE WHERE PCN_SAISIEDEPORTEE = "X" ' +
    ' AND PCN_VALIDRESP = "VAL" AND PCN_EXPORTOK <> "X" AND PCN_DATEFINABS <="' +
    USDATETIME(StrToDAte(LaDate)) + '"' +
    ' ORDER BY PCN_SALARIE', true);
  if Q.eof then
  begin
    Ferme(Q);
    exit;
  end;
  TobAbs.Loaddetaildb('ABSENCESALARIE', '', '', Q, False);
  Ferme(Q);

  TempFName := VH_Paie.PGCHEMINEAGL + '\AbsencesValid.txt';
  if FileExists(TempFName) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer le fichier extraction  ' + ExtractFileName(TempFName) + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DeleteFile(PChar(TempFName))
    else
    begin
      TobAbs.Free;
      FiniMoveProgressForm;
      exit;
    end;
  end;
  AssignFile(fichier, TempFName);
  {$I-}ReWrite(fichier);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + TempFName, 'Abandon du traitement');
    Exit;
  end;

  MoveCurProgressForm('Fin d''extraction des absences, écriture dans le fichier');

  st := '***DEBUT***';
  Writeln(Fichier, st);
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  noDoss := PgRendNoDossier();

  MoisE := '';
  AnneeE := '';
  ComboExer := '';
  DebExer := 0;
  FinExer := 0;
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  decodedate(DebExer, aad, mmd, jjd);
  decodedate(FinExer, aaf, mmf, jjf);
  st := '000;' + nodoss + ';' + inttostr(jjd) + '/' + inttostr(mmd) + '/' + inttostr(aad)
    + ';' + inttostr(jjf) + '/' + inttostr(mmf) + '/' + inttostr(aaf) + ';';
  Writeln(Fichier, st);
  TA := TobAbs.findfirst([''], [''], true);

  while TA <> nil do
  begin
    st := 'MAB;' + Ta.getValue('PCN_SALARIE') + ';';
    st := st + datetostr(Ta.getValue('PCN_DATEDEBUTABS')) + ';';
    st := st + datetostr(Ta.getValue('PCN_DATEFINABS')) + ';';
    st := st + floattostr(Ta.getValue('PCN_JOURS')) + ';';
    st := st + floattostr(Ta.getValue('PCN_HEURES')) + ';';
    st := st + Ta.getValue('PCN_TYPECONGE') + ';';
    st := st + Ta.getValue('PCN_LIBELLE') + ';';
    //  PT1 14/06/01 PH
    st := st + Ta.getValue('PCN_LIBCOMPL1') + ';';
    st := st + Ta.getValue('PCN_LIBCOMPL2') + ';';
    st := st + Ta.getValue('PCN_DEBUTDJ') + ';';
    st := st + Ta.getValue('PCN_FINDJ');

    Writeln(Fichier, st);
    MoveCurProgressForm('Absences du salarié : ' + Ta.getvalue('PCN_SALARIE'));
    TA := TobAbs.findnext([''], [''], true);
  end;
  st := '***FIN***';
  Writeln(Fichier, st);
  Closefile(fichier);
  MoveCurProgressForm('Fin d''écriture du fichier, Mise à jour en cours');

  if Tobabs <> nil then Tobabs.Free;
  //  PT1 14/06/01 PH
  // les absences sont récupérées dans le fichier, il convient de les toper comme exportées
  // Si import non Ok pour quelconque raison, alors il faut les reintéger
  //PT9-3 Etat en cours d'intégration
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_EXPORTOK = "ENC" WHERE PCN_SAISIEDEPORTEE = "X" AND' +
    ' PCN_VALIDRESP = "VAL" AND PCN_EXPORTOK <> "X" AND PCN_DATEFINABS <="' +
    USDATETIME(StrToDAte(LaDate)) + '"');
  // on met à jour toutes les absences qui n'ont pas été exportées et qui sont de saisies déportées
  // et validées par le salariés et par le responsables des absences
  FiniMoveProgressForm;
end;

// fonction qui calcule les cumuls cp et rtt (45 et 46) pour les exporter dans le récapsal
procedure trouveCumRtt(var Cum45, cum46: double; salarie, etab: string; DateIntegration: TDateTime; DateDebutExerSoc: tdatetime);
var
  raz, st, PCLCumul, StCumul: string;
  CumulDD, {DateDebutExerSoc,PT9-1} DateJamais, TDate, DateCP, DTcloturePi, DateBid: tdatetime;
  QPrem, QMontEx, QRechCumul: TQuery;
  TE: tob;
  EnAnnee, EnMois, EnJour, RMois, CpAnnee, Cpmois, Cpjour: word;
  T_RechCumul, TRech: TOB;
begin
  raz := '';
  CumulDD := Dateintegration; //PT9-1 Utilisation de DateIntegration
  Cum45 := 0;
  cum46 := 0;

  {DEB PT11 Ajout d'un traitement multi cumul}
  StCumul := '';
  St := VH_Paie.PgRttAcquis;
  while St <> '' do
    StCumul := StCumul + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
  St := VH_Paie.PgRttPris;
  while St <> '' do
    StCumul := StCumul + ' PHC_CUMULPAIE="' + ReadTokenSt(St) + '" OR';
  if StCumul <> '' then StCumul := ' AND (' + Copy(StCumul, 1, Length(StCumul) - 2) + ') ';
  {FIN PT11}
  T_RechCumul := TOB.Create('Lescumuls 45et46', nil, -1);
  QRechCumul := OpenSql('SELECT DISTINCT PHC_CUMULPAIE,PCL_RAZCUMUL FROM HISTOCUMSAL  ' +
    'LEFT JOIN CUMULPAIE ON PCL_CUMULPAIE=PHC_CUMULPAIE ' +
    'WHERE ##PCL_PREDEFINI## PHC_SALARIE="' + salarie + '" ' + StCumul +
    // PT11 'AND (PHC_CUMULPAIE="'+VH_Paie.PgRttAcquis+'" OR PHC_CUMULPAIE="'+VH_Paie.PgRttPris+'" ) '+ //PT7 Modif requête
    'ORDER BY PHC_CUMULPAIE', True);
  T_RechCumul.loadDetailDb('HISTOCUMSAL', '', '', QRechCumul, FALSE, FALSE);
  Ferme(QRechCumul);
  //RAZ = Jamais : Date de debut premier bulletin
  (*PT9-1 Recherhce de la date de fin de calcul des compteurs*)
  QPrem := OpenSQl('SELECT MIN(PHC_DATEDEBUT) AS DATEDEB ' +
    'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '"', True);
  if not QPrem.eof then //PORTAGECWAS
    DateJamais := QPrem.FindField('DATEDEB').AsDateTime
  else
    DateJamais := idate1900;
  Ferme(QPrem);

  //RAZ = Date de debut d'exercice CP
  TE := tobetab.findfirst(['ETB_ETABLISSEMENT'], [Etab], True);
  TDate := 0;
  if TE <> nil then
    TDate := TE.getvalue('ETB_DATECLOTURECPN');
  if TDate > 0 then
  begin
    RMois := 1;
    RendPeriode(DTcloturePi, DateBid, TDate, DateIntegration); //PT9-1 Utilisation de DateIntegration
    DecodeDate(DateIntegration, EnAnnee, EnMois, EnJour); //Date de fin bulletin
    DecodeDate(DTcloturePi, CpAnnee, Cpmois, Cpjour);
    //DateCP Posterieur à DateBulletin alors on remonte à l'année précédente
    if EnMois <= CpMois then
    begin
      EnAnnee := EnAnnee - 1;
      RMois := Cpmois + 1;
    end;
    //DateCP antérieur à DateBulletin alors on remonte au debut mois suivant
    if EnMois > Cpmois then RMois := Cpmois + 1;
    //si DateCp = Decembre alors on remonte en Janvier
    if CpMois = 12 then RMois := 1;
    DateCP := EncodeDate(EnAnnee, RMois, 1);
  end
  else DateCP := 2;
  //Récupère les sommes des cumuls en fonction de sa periode de remise à zero
  TRech := T_RechCumul.Findfirst([''], [''], FALSE);
  while TRech <> nil do
  begin
    PclCumul := TRech.GetValue('PHC_CUMULPAIE');
    ;
    Raz := TRech.GetValue('PCL_RAZCUMUL');
    if Raz = '00' then
    begin
      CumulDD := DateDebutExerSoc;
    end;
    if (raz = '01') or (raz = '02') or (raz = '03') or (raz = '04') or (raz = '05') or (raz = '06') or (raz = '07') or (raz = '08') or (raz = '09') or (raz = '10') or (raz = '11')
      or (raz = '12') then
    begin //PT9-1 Utilisation de DateIntegration
      DecodeDate(DateIntegration, CpAnnee, CpMois, CpJour);
      if CpMois < StrtoInt(raz) then Cpannee := Cpannee - 1;
      {PT10 mise en commentaire on doit tenir compte de RAZ dans l'encodedate
      //If CpMois>=StrToInt(raz) then CpMois:=StrToInt(raz);
      //CumulDD:=EncodeDate(CpAnnee,CpMois,Cpjour);}
      CumulDD := EncodeDate(CpAnnee, StrToInt(raz), 1); //PT12
    end;
    if raz = '99' then cumulDD := DateJamais;
    if raz = '20' then cumulDD := DateCP;

    if (DateIntegration > 2) and (CumulDD > 0) and (Salarie <> '') then
    begin //PT9-1 Utilisation de DateIntegration
      st := 'SELECT SUM(PHC_MONTANT) AS MONTANT,PHC_CUMULPAIE ' +
        'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '" AND ' +
        'PHC_CUMULPAIE="' + PclCumul + '" ' +
        'AND PHC_DATEDEBUT>="' + UsDateTime(CumulDD) + '" ' +
        'AND PHC_DATEFIN<="' + UsDateTime(DateIntegration) + '" ' +
        'GROUP BY PHC_CUMULPAIE';
      QMontEx := OpenSql(st, TRUE);
      if not QMontEx.EOF then //PORTAGECWAS
      begin
        {PT11 Mise en commentaire, remplacé par un traitement multi cumul
        if PclCumul = VH_Paie.PgRttAcquis then Cum45 := QMontEx.FindField('MONTANT').Asfloat //PT7 Modif clause
        else Cum46 := QMontEx.FindField('MONTANT').Asfloat;}
        if Pos(PclCumul, VH_Paie.PgRttAcquis) > 0 then
          Cum45 := Cum45 + QMontEx.FindField('MONTANT').Asfloat;
        if Pos(PclCumul, VH_Paie.PgRttPris) > 0 then
          Cum46 := Cum46 + QMontEx.FindField('MONTANT').Asfloat;
      end;
      Ferme(QMontEx);
    end;
    TRech := T_RechCumul.FindNext([''], [''], FALSE);
  end;
  if T_RechCumul <> nil then T_RechCumul.Free;
end;


procedure AfficheTabSalResp(Sender: Tobject; Matricule: string; Quoi: string = '');
var
  StWhere: string;
begin
  if sender = nil then exit;
  if Matricule <> '' then
  begin
    StWhere := 'PSE_SALARIE=PSA_SALARIE AND ';
    if Quoi = 'P' then StWhere := StWhere + ' PSE_RESPONSVAR="' + Matricule + '" '
    else if Quoi = 'ASS' then StWhere := StWhere + ' PSE_ASSISTABS="' + Matricule + '" '  { PT18 }
    else StWhere := StWhere + ' PSE_RESPONSABS="' + Matricule + '" ';
  end
  else
    StWhere := 'PSE_SALARIE=PSA_SALARIE ';

  StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT22
  LookupList(TControl(Sender), 'Salarié', 'DEPORTSAL,SALARIES', 'DISTINCT PSE_SALARIE', 'PSA_LIBELLE', StWhere, 'PSE_SALARIE', True, -1);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 31/10/2002
Modifié le ... :   /  /
Description .. : Calcul des compteur en cours de la table RECAPSALARIES
Mots clefs ... : PAIE;ABSENCE;ECONGES
*****************************************************************}
{DEB PT8}
procedure CalculRecapAbsEnCours(Salarie: string);
var
  St: string;
  Q: TQuery;
  Tob_MvtEncours, Tob_RecapSalarie, Tmvt, TRecap: TOB;
  i: integer;
begin
  Tob_RecapSalarie := nil;
  Tob_MvtEncours := nil;
  try
    beginTrans;
    InitMoveProgressForm(nil, 'Calcul des compteurs en cours', 'Veuillez patienter SVP ...', 1, FALSE, TRUE);
    //Remise à zero des compteurs salariés
    if salarie <> '' then St := ' WHERE PRS_SALARIE="' + salarie + '"' else St := '';
    ExecuteSql('UPDATE RECAPSALARIES SET PRS_PRIVALIDE=0,PRS_PRIATTENTE=0,' +
      'PRS_RTTVALIDE=0,PRS_RTTATTENTE=0 ' + St);
    if salarie <> '' then St := ' AND PCN_SALARIE="' + salarie + '"' else St := '';
    St := 'SELECT PCN_TYPEMVT,PCN_SALARIE,PCN_TYPECONGE,PCN_JOURS,PCN_VALIDRESP FROM ABSENCESALARIE ' +
      'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE ' + { PT 19 }
      'WHERE ((PCN_TYPEMVT="ABS" AND PMA_TYPEABS="RTT") '+
      'OR (PCN_TYPECONGE="PRI" AND PCN_PERIODECP<2)) ' + { PT14 } { PT21 } 
      'AND PCN_ETATPOSTPAIE <> "NAN" '+  { PT20 }
      'AND PMA_MOTIFEAGL="X" AND (PCN_EXPORTOK<>"X" OR PCN_CODETAPE="...") ' + St +
      'ORDER BY PCN_SALARIE,PCN_TYPEMVT';
    Q := OpenSql(st, True);
    if not Q.Eof then
    begin
      //Charger la tob des mvts en cours
      Tob_MvtEncours := TOb.Create('Les mvts en cours', nil, -1);
      Tob_MvtEncours.LoadDetailDB('Les mvts en cours', '', '', Q, False);
      Ferme(Q);
      //Charge la tob des recapsalaries
      Tob_RecapSalarie := TOb.Create('RECAPSALARIES', nil, -1);
      if Salarie <> '' then
        Tob_RecapSalarie.LoadDetailDB('RECAPSALARIES', '"' + Salarie + '"', '', nil, False)
      else
        Tob_RecapSalarie.LoadDetailDB('RECAPSALARIES', '', '', nil, False);
      //Boucle sur mvts en cours et alimente la table RECAPSALARIE en fonction de l'état de validation
      for i := 0 to Tob_MvtEncours.detail.count - 1 do
      begin
        Tmvt := Tob_MvtEncours.detail[i];
        MoveCurProgressForm('Salarié : ' + Tmvt.GetValue('PCN_SALARIE'));
        TRecap := Tob_RecapSalarie.FindFirst(['PRS_SALARIE'], [Tmvt.GetValue('PCN_SALARIE')], False);
        if TRecap <> nil then
        begin
          if (Tmvt.GetValue('PCN_TYPECONGE') = 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'VAL') then
            TRecap.PutValue('PRS_PRIVALIDE', TRecap.GetValue('PRS_PRIVALIDE') + Tmvt.GetValue('PCN_JOURS'))
          else
            if (Tmvt.GetValue('PCN_TYPECONGE') = 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'ATT') then
            TRecap.PutValue('PRS_PRIATTENTE', TRecap.GetValue('PRS_PRIATTENTE') + Tmvt.GetValue('PCN_JOURS'))
          else
            if (Tmvt.GetValue('PCN_TYPECONGE') <> 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'VAL') then
            TRecap.PutValue('PRS_RTTVALIDE', TRecap.GetValue('PRS_RTTVALIDE') + Tmvt.GetValue('PCN_JOURS'))
          else
            if (Tmvt.GetValue('PCN_TYPECONGE') <> 'PRI') and (Tmvt.GetValue('PCN_VALIDRESP') = 'ATT') then
            TRecap.PutValue('PRS_RTTATTENTE', TRecap.GetValue('PRS_RTTATTENTE') + Tmvt.GetValue('PCN_JOURS'));
        end;
      end;
    end
    else Ferme(Q);
    if Tob_RecapSalarie <> nil then
      if Tob_RecapSalarie.IsOneModifie then
        for i := 0 to Tob_RecapSalarie.detail.count - 1 do
        begin
          if Tob_RecapSalarie.Detail[i].Modifie then
            Tob_RecapSalarie.Detail[i].UpdateDB;
        end;
    CommitTrans;
  except
    Rollback;
  end;
  FiniMoveProgressForm;

  if Tob_MvtEncours <> nil then Tob_MvtEncours.free;
  if Tob_RecapSalarie <> nil then Tob_RecapSalarie.free;
end;
{FIN PT8}
{ DEB PT13 }
function IfMotifabsenceSaisissable(TypeConge, FicPrec: string): Boolean;
begin
  Result := (RechDom('PGMOTIFABSENCERESP',TypeConge,False)<>'');
  if (FicPrec <> 'SAL') AND (FicPrec <> 'ASS') and (FicPrec <> 'RESP') then Result := True; { PT18 }
 (* {$IFDEF ETEMPS}
  Result := True;
  TypeConge := Uppercase(TypeConge);
  if (TypeConge = 'MAL') or (TypeConge = 'MAT') or (TypeConge = 'PAT')
    or (TypeConge = 'ACT') or (TypeConge = 'AM1') or (TypeConge = 'AMS') then
  begin
    Result := False;
    if (FicPrec <> 'SAL') and (FicPrec <> 'RESP') then
      Result := True;
  end;
  {$ENDIF}    *)
end;
{ FIN PT13 }

end.

