{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 13/06/2001
Modifié le ... :   /  /
Description .. : Gestion, Import,Export des éléments nécessaires à une
Suite ........ : saisie déportée des Absences.
Suite ........ : Comporte les motifs pour ne prendre en compte que les
Suite ........ : motifs autorisés.
Suite ........ : Attention,on génére un fichier contenant les infos
Suite ........ : nécessaires
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************}
{
PT1 : 31/01/2002 V571 PH réaffectation des numéros d'ordre et des codes de regroupement
                      en récupération des des absences dans la base epaie
PT2 : 08/04/2002 V571 PH Controle des noms de fichier
PT3 : 08/04/2002 V571 PH Pas de gestion des confidentiels dans la base eConges
PT4 : 08/04/2002 V571 PH Analyse des donnees contenus dans le fichier pour eliminer les enrg
                         à blanc venant des jointures multiples
PT5 : 27/05/2002 V575 PH Filtrage des données de choixcod uniquement sur les tablettes de la paie
PT6 : 11/07/2002 V582 SB Retrait des salariés sorties pour l'export des salariés
PT7 : 06/09/2002 V582 SB Ajout des tables HIERARCHIE,SERVICES ET SERVICEORDRE pour l'import export des données
PT8 : 17/10/2002 V585 SB Génération d'évènements pour traçage
PT9   04/11/2002 V585 SB Intégration du calcul des compteurs Abs en cours
PT10  28/11/2002 V591 SB Modification de l'etat des absences à supprimer dans la base Econges avant inport
PT11-1  05/12/2002 V591 SB Retrait du PT6 : la gestion des temps nécessite l'utilisation des salariés sortis
PT11-2  05/12/2002 V591 SB Ajout d'une date d'intégration saisissable
PT12    04/03/2003 V591 SB Ajout d'une périodicité pour l'export des absences non CP
PT13    18/03/2003 V591 SB Filtrage des données communes récupérées
                           Récupération de la date d'intégration sur coche récap
PT14    07/11/2003 V_42 SB Exclusion des motifs spécifiques CEGID
PT15    09/04/2004 V_50 SB FQ 11136 Ajout Gestion des congés payés niveau salarié
PT16    08/10/2004 V_50 PH Rechargement du cache des paramsoc
PT17    24/06/2005 V_60 PH Traitement de la table INTERIMAIRES avec Salariés
PT18    09/09/2005 V_60 PH On ne prend que les salariés présents dans les 2 tables
PT19    23/01/2006 V_65 SB FQ 10866 Ajout clause predefini motif d'absence
PT20    19/06/2006 V_65 SB FQ 13231 Retrait des mvt absences annulées
PT21    06/12/2006 V_65 SB FQ 13735 Parmaétrage de la periode cp de récupération
PT22  : 14/03/2008 VG V_80 Export du champ PSA_SORTIEDEFINIT
}

unit UtofElisteTables;

interface
uses SysUtils, HCtrls, HEnt1, Classes,
  Controls,  HMsgBox,  UTOB, UTOF, HTB97,
{$IFDEF EAGLCLIENT}
  eFiche,
{$ELSE}
  Fiche, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Paramsoc;
type
  TOF_ElisteTables = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnClose {Ajout proc. PT8}; override;
    procedure ValidExport(Sender: Tobject);
    procedure ValidImport(Sender: Tobject);
    function ValideFic(Nom, st, LaTable: string): boolean;
    function RecupFic(Nom, st, LaTable: string): boolean;
  private
    TypTrait: string;
    TEvent: TStringList; //PT8
    procedure AnalyseNomFic(var NomFic: string);
  public
  end;
  {
  NoErreur :
  5 : Manque marque début fichier

  }

implementation

uses P5Def, entpaie, PgOutilseAgl;

procedure TOF_ElisteTables.OnArgument(Arguments: string);
var
  BTN: TToolbarbutton97;
begin
  TEvent := TStringList.create; //PT8 Ajout TStringlist
  TypTrait := Arguments;
  if TypTrait = '' then TypTrait := 'E';
  if TypTrait = '' then Close; // pas d'argument, on ne sait pas si on traite un import ou un export
  BTN := TToolbarbutton97(GetControl('BVALIDER'));
  if BTN <> nil then
  begin
    if TypTrait = 'E' then
    begin
      TFFiche(Ecran).caption := TFFiche(Ecran).caption + ' exporter';
      BTN.OnClick := ValidExport;
    end
    else
    begin
      TFFiche(Ecran).caption := TFFiche(Ecran).caption + ' importer';
      BTN.OnClick := ValidImport;
      SetControlVisible('DATEINTEGRATION', False);
      SetControlVisible('LDATEINTEGRATION', False);
      SetControlVisible('LBPERRTT', False); //PT12
      SetControlVisible('LBPERRTT_', False);
      SetControlVisible('DATEDEBABS', False);
      SetControlVisible('DATEFINABS', False);
    end;
  end;
  UpdateCaption(TFFiche(Ecran));
end;
// Fonction de lancement de l'import des données dans la base eAGL

procedure TOF_ElisteTables.ValidImport(Sender: Tobject);
var
  st, Chemin: string;
  Okok: Boolean;
begin
  Chemin := VH_Paie.PGCHEMINEAGL;
  if Chemin = '' then
  begin
    PGIBox('Chemin de localisation des fichiers inexistant', 'Export Fichier');
    exit;
  end;
  try
    BeginTrans;
    TEvent.Clear; //PT8
    if GetControlText('CBSAL') = 'X' then
    begin
      RecupFic('\Salaries.txt', '', 'SALARIES');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'SALARIES', False)); //PT8
      //DEB PT7 Import des nouvelles tables
      RecupFic('\Hierarchie.txt', '', 'HIERARCHIE');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'HIERARCHIE', False)); //PT8
      RecupFic('\Services.txt', '', 'SERVICES');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'SERVICES', False)); //PT8
      RecupFic('\ServiceOrdre.txt', '', 'SERVICEORDRE');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'SERVICEORDRE', False)); //PT8
      //FIN PT7
      // DEB PT17
      RecupFic('\Interimaires.txt', '', 'INTERIMAIRES');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'INTERIMAIRES', False));
      //FIN PT17
    end;

    if GetControlText('CBDROITSAL') = 'X' then
    begin
      RecupFic('\DeportSal.txt', '', 'DEPORTSAL');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'DEPORTSAL', False)); //PT8
    end;

    if GetControlText('CBCALENDRIER') = 'X' then
    begin
      RecupFic('\Calendrier.txt', '', 'CALENDRIER');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'CALENDRIER', False)); //PT8
      RecupFic('\JourFerie.txt', '', 'JOURFERIE');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'JOURFERIE', False)); //PT8
    end;
    if GetControlText('CBMAB') = 'X' then
    begin
      RecupFic('\Motif.txt', '', 'MOTIFABSENCE');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'MOTIFABSENCE', False)); //PT8
    end;


    if GetControlText('CBPARAMSOC') = 'X' then // Voir si base commune avec autre  appli
    begin
      //PT13 Suppression des paramsoc Paie avant l'import
      st := 'DELETE FROM PARAMSOC WHERE SOC_NOM LIKE "SO_PG%"';
      RecupFic('\ParamSoc.txt', st, 'PARAMSOC');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'PARAMSOC', False)); //PT8
{$IFDEF EAGLCLIENT} // PT16
      AvertirCacheServer('PARAMSOC');
{$ENDIF}
    end;
    if GetControlText('CBTABLES') = 'X' then
    begin
      //PT13 Suppression des tablettes libres Paie
   (* PT5 : 27/05/2002 V575 PH Filtrage des données de choixcod uniquement sur les tablettes de la paie
      st := 'select choixcod.* from choixcod left join decombos on do_type=cc_type and do_domaine="P"  where CC_TYPE like "P%" ';*)
      st := 'DELETE FROM CHOIXCOD WHERE CC_TYPE IN ' +
        '(SELECT DO_TYPE FROM DECOMBOS WHERE DO_DOMAINE="P" AND DO_PREFIXE="CC")';
      RecupFic('\Tables.txt', st, 'CHOIXCOD');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'CHOIXCOD', False)); //PT8
    end;

    if GetControlText('CBETAB') = 'X' then
    begin
      okok := RecupFic('\Etab.txt', '', 'ETABLISS');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'ETABLISS', False)); //PT8
      if not okok then exit;
      okok := RecupFic('\EtabCompl.txt', '', 'ETABCOMPL');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'ETABCOMPL', False)); //PT8
      if not okok then exit;
    end;


    // Recup des absences salaries PeriodCp = 0 ou 1 ou -1
    if GetControlText('CBABSENCES') = 'X' then
    begin
      // destruction des absences exportees sauf celles qui sont annulées pour conserver une trace
      // d'un congé qui a été posée mais réfusé ou annulé par le responsable ou Tjrs en attente
      //PT10 On supprime les mouvements intégrés et en cours d'intégration
      st := 'DELETE FROM ABSENCESALARIE WHERE (PCN_EXPORTOK = "X" OR PCN_EXPORTOK="ENC" )';
{$IFDEF ETEMPS}
      st := St + ' AND PCN_TYPECONGE NOT IN ("MAL","MAT","PAT","ACT","AM1","AMS")'; //PT14
{$ENDIF}
      //         '(PCN_VALIDRESP <> "NAN" AND PCN_VALIDRESP <> "REF" AND PCN_VALIDRESP <> "ATT" AND PCN_SAISIEDEPORTEE ="X")';
      RecupFic('\Absences.txt', st, 'ABSENCESALARIE');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'ABSENCESALARIE', False)); //PT8
    end;
    //on importe d'abord les absences ensuite le récap puis on calcul les compteurs en cours
    // Recap Salaries CP et RTT sur N-1 et N
    if GetControlText('CBCAL') = 'X' then
    begin
      RecupFic('\recapsal.txt', '', 'RECAPSALARIES');
      TEvent.Add('Import de la table ' + RechDom('TTTABLES', 'RECAPSALARIES', False)); //PT8
      //Calcul des mvts En cours validé et en attente pour tous les salariés
      CalculRecapAbsEnCours(''); //PT9
      //PT13 Importation du paramsoc dateintegration
      st := 'DELETE FROM PARAMSOC WHERE SOC_NOM = "SO_PGECABDATEINTEGRATION"';
      RecupFic('\DateIntegre.txt', st, 'PARAMSOC');
      TEvent.Add('Import du paramètre société date d''intégration.');
    end;

    CommitTrans;
    PGIBox('Le traitement d''import est terminé', Ecran.caption);
    CreeJnalEvt('004', '129', 'OK', nil, nil, TEvent); //PT8
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de l''écriture dans la base', 'Traitement fichier import');
    TEvent.add('Une erreur est survenue lors de l''écriture dans la base');
    CreeJnalEvt('004', '129', 'ERR', nil, nil, TEvent); //PT8
  end;

end;

procedure TOF_ElisteTables.ValidExport(Sender: Tobject);
var
  Chemin, st, StWhere, StWhereCP : string;
  okok: Boolean;
begin
st:= '';
Chemin:= VH_Paie.PGCheminRech;
if (Chemin='') then
   begin
   PGIBox ('Chemin de localisation des fichiers inexistant', 'Export Fichier');
   exit;
   end;
TEvent.Clear;
if (GetControlText ('CBSAL')='X') then
   begin
   st:= 'SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_ETABLISSEMENT,'+
        ' PSA_BASANCCP, PSA_LIBELLEEMPLOI, PSA_CONGESPAYES, PSA_CALENDRIER,'+
        ' PSA_CPACQUISANC, PSA_CPACQUISMOIS, PSA_CPACQUISSUPP,'+
        ' PSA_CPTYPEMETHOD, PSA_CPTYPERELIQ, PSA_DATANC, PSA_DATEACQCPANC,'+
        ' PSA_DATEANCIENNETE, PSA_DATEENTREE, PSA_DATESORTIE,'+
        ' PSA_SORTIEDEFINIT, PSA_NBREACQUISCP, PSA_NBRECPSUPP, PSA_PROFILCGE,'+    //PT22
        ' PSA_RELIQUAT, PSA_RIBACPSOC, PSA_STANDCALEND, PSA_TYPACPSOC,'+
        ' PSA_TYPDATANC, PSA_TYPPROFILCGE, PSA_VALANCCP, PSA_VALORINDEMCP,'+
        ' PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN3, PSA_TRAVAILN4,'+
        ' PSA_CODESTAT, PSA_LIBREPCMB1, PSA_LIBREPCMB2, PSA_LIBREPCMB3,'+
        ' PSA_LIBREPCMB4'+
        ' FROM SALARIES WHERE'+
        ' EXISTS (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE=PSA_SALARIE)'+
        ' ORDER BY PSA_SALARIE';
   okok:= ValideFic ('\Salaries.txt', st, 'SALARIES');
   TEvent.Add ('Export de la table '+RechDom ('TTTABLES', 'SALARIES', False));
   if not okok then
      exit;

    //DEB PT7 Export des nouvelles tables
    okok := ValideFic('\Hierarchie.txt', '', 'HIERARCHIE');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'HIERARCHIE', False)); //PT8
    if not okok then exit;

    okok := ValideFic('\Services.txt', '', 'SERVICES');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'SERVICES', False)); //PT8
    if not okok then exit;

    okok := ValideFic('\ServiceOrdre.txt', '', 'SERVICEORDRE');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'SERVICEORDRE', False)); //PT8
    if not okok then exit;
    //FIN PT7
    //DEB PT17
    st := 'SELECT * FROM INTERIMAIRES  WHERE (EXISTS (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSI_INTERIMAIRE = PSE_SALARIE) AND PSI_TYPEINTERIM="SAL") OR (PSI_TYPEINTERIM<>"SAL")';
    okok := ValideFic('\Interimaires.txt', st, 'INTERIMAIRES');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'INTERIMAIRES', False)); //PT8
    if not okok then exit;
    //FIN PT17
  end;

  if GetControlText('CBDROITSAL') = 'X' then
  begin //PT6 18/09/2002
    { PT11-1 Mise en commentaire
    st:='SELECT * FROM DEPORTSAL WHERE PSE_SALARIE IN '+
       '(SELECT PSA_SALARIE FROM SALARIES WHERE (PSA_DATESORTIE IS NULL '+
       'OR PSA_DATESORTIE="'+DateToStr(idate1900)+'"  '+
       'OR PSA_DATESORTIE>=(SELECT MAX(PPU_DATEFIN) FROM PAIEENCOURS)))';}
    okok := ValideFic('\DeportSal.txt', '', 'DEPORTSAL');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'DEPORTSAL', False)); //PT8
    if not okok then exit;
  end;

  if GetControlText('CBCALENDRIER') = 'X' then
  begin
    okok := ValideFic('\Calendrier.txt', '', 'CALENDRIER');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'CALENDRIER', False)); //PT8
    if not okok then exit;
    okok := ValideFic('\JourFerie.txt', '', 'JOURFERIE');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'JOURFERIE', False)); //PT8
    if not okok then exit;
  end;

  if GetControlText('CBMAB') = 'X' then
  begin
    st := 'SELECT * FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFEAGL = "X"';  { PT19 } 
    okok := ValideFic('\Motif.txt', st, 'MOTIFABSENCE');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'MOTIFABSENCE', False)); //PT8
    if not okok then exit;
  end;

  if GetControlText('CBPARAMSOC') = 'X' then
  begin
    //PT13 Récupération des paramsoc Paie
    st := 'SELECT * FROM PARAMSOC WHERE SOC_NOM LIKE "SO_PG%"';
    okok := ValideFic('\ParamSoc.txt', st, 'PARAMSOC');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'PARAMSOC', False)); //PT8
    if not okok then exit;
  end;

  if GetControlText('CBTABLES') = 'X' then
  begin
    //PT13 selection des tablettes libres Paie
    st := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE IN ' +
      '(SELECT DO_TYPE FROM DECOMBOS WHERE DO_DOMAINE="P" AND DO_PREFIXE="CC")';
    okok := ValideFic('\Tables.txt', St, 'CHOIXCOD');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'CHOIXCOD', False)); //PT8
    if not okok then exit;
  end;

  if GetControlText('CBETAB') = 'X' then
  begin
    okok := ValideFic('\Etab.txt', '', 'ETABLISS');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'ETABLISS', False)); //PT8
    if not okok then exit;
    okok := ValideFic('\EtabCompl.txt', '', 'ETABCOMPL');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'ETABCOMPL', False)); //PT8
    if not okok then exit;
  end;

  // Recap Salaries CP et RTT sur N-1 et N
  if GetControlText('CBCAL') = 'X' then
  begin
    //DEB PT11-2 On calcul le recap à partir de la date saisie, date maj dans les paramètres sociétés
    if (StrToDate(GetControlText('DATEINTEGRATION')) <> VH_Paie.PGECabDateIntegration) or (VH_Paie.PGECabDateIntegration <= idate1900) then
    begin
      SetParamSoc('SO_PGECABDATEINTEGRATION', StrToDate(GetControlText('DATEINTEGRATION')));
      VH_Paie.PGECabDateIntegration := StrToDate(GetControlText('DATEINTEGRATION'));
{$IFDEF EAGLCLIENT} // PT16
      AvertirCacheServer('PARAMSOC');
{$ENDIF}
      ChargeParamsPaie;
    end;
    ChargeTobSalRecap('A', StrToDate(GetControlText('DATEINTEGRATION')));
    //FIN PT11-2
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'RECAPSALARIES', False)); //PT8
    //PT13 Exportation du paramsoc dateintegration
    st := 'SELECT * FROM PARAMSOC WHERE SOC_NOM = "SO_PGECABDATEINTEGRATION"';
    ValideFic('\DateIntegre.txt', st, 'PARAMSOC');
  end;
  { Recup des absences salaries
  PeriodCp = 0 = Periode En Cours
             1 Periode Antérieure
            -1 Période suivante ou par anticipation
  Soit exercice EnCours + N-1 A définir
  }
  if GetControlText('CBABSENCES') = 'X' then
  begin
    {GetLocalTime(SystemTime0);
    LaDate := SystemTimeToDateTime(SystemTime0);
    LaDate := DebutDeMois (LaDate);   // On se positionne au debut du mois courant
    LaDate := PlusMois (LaDate, -12); // On recherche l'antériorité d'un an   }
    { DEB PT12 Ajout périodicité}
    if (GetControl('DATEDEBABS') <> nil) and (GetControl('DATEFINABS') <> nil) then
      StWhere := 'AND PCN_DATEDEBUTABS>="' + USDateTime(StrToDate(GetControlText('DATEDEBABS'))) + '" ' +
        'AND PCN_DATEFINABS<="' + USDateTime(StrToDate(GetControlText('DATEFINABS'))) + '"'
    else
      StWhere := '';
    { FIN PT12 }

    { DEB PT21 }
    If GetControlText('PERIODECP') <> '' then
      Begin
      St := GetControlText('PERIODECP');
      If Pos('Tous',St) > 0 then
        StWhereCP := ' AND (PCN_PERIODECP >= -1 AND PCN_PERIODECP <= 2)'
      else
        Begin
        While St <> '' do
          StWhereCP := StWhereCP +' PCN_PERIODECP='+ReadTokenSt(St)+' OR';
        StWhereCP := 'AND (PCN_PERIODECP = -1 OR '+Copy(StWhereCP,1,length(StWhereCP)-3)+')';
        End;
      End
    else
       StWhereCP := ' AND (PCN_PERIODECP >= -1 AND PCN_PERIODECP <= 2)';
    { FIN PT21 }
    St := 'SELECT * FROM ABSENCESALARIE RIGHT OUTER JOIN DEPORTSAL ON PSE_SALARIE=PCN_SALARIE ' +
      ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE ' + { PT19 }
      ' LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE ' +
      ' WHERE ##PMA_PREDEFINI## PMA_MOTIFEAGL = "X" ' +
      ' AND ((PCN_MVTDUPLIQUE<>"X" AND PCN_TYPEMVT = "CPA" AND PCN_TYPECONGE="PRI" '+StWhereCP+')' + { PT21 }
      ' OR (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-" ' + StWhere + '))' + { PT12 }
      ' AND PCN_ETATPOSTPAIE <> "NAN" '+  { PT20 }
    {PT11-1 Mise en commentaire
     ' AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+DateToStr(idate1900)+'" '+  //PT6
     ' OR PSA_DATESORTIE>=(SELECT MAX(PPU_DATEFIN) FROM PAIEENCOURS)) '+  //PT6 retrait des salariés sortis}
    ' ORDER BY PCN_SALARIE ';
    okok := ValideFic('\Absences.txt', st, 'ABSENCESALARIE');
    TEvent.Add('Export de la table ' + RechDom('TTTABLES', 'ABSENCESALARIE', False)); //PT8
    if not okok then exit;
  end;

  CreeJnalEvt('004', '128', 'OK', nil, nil, TEvent); //PT8
  PGIBox('Le traitement d''export est terminé.', Ecran.caption);
end;

// fonction qui ecrit dans un fichier texte mais qui prend soin de verifier si le fichier existe et si on peut le creer
// Nom = Nom du fichier avec son extension, st = chaine comprenant la requete SQL pour la Query

function TOF_ElisteTables.ValideFic(Nom, st, LaTable: string): boolean;
var
  F: TextFile;
  Q: TQuery;
  T: TOB;
  NomFic: string;
begin
  NomFic := VH_Paie.PGCHEMINEAGL + Nom;
  result := FALSE;
  // PT2 : 08/04/2002 V571 PH Controle des noms de fichier
  AnalyseNomFic(NomFic);
  if FileExists(NomFic) then DeleteFile(PChar(NomFic));
  AssignFile(F, NomFic);
{$I-}ReWrite(F);
{$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + NomFic, 'Abandon du traitement');
    Exit;
  end;
  CloseFile(F);
  T := tob.create('', nil, -1);
  if St <> '' then
  begin
    Q := OpensQl(st, True);
    T.loaddetaildb(LaTable, '', '', Q, false);
  end
  else T.loaddetaildb(LaTable, '', '', nil, false);
  if T.detail.count >= 1 then T.SaveToFile(NomFic, False, True, True, '');
  if st <> '' then Ferme(Q);
  T.free;
  Result := TRUE;
end;
{ Fonction de recupération des données extraite de la base de production pour insertin dans la base ePaie
Nom = Nom du fichier avec son extension,
st = chaine comprenant la requete SQL pour une suppression specifique sur une table
}

function TOF_ElisteTables.RecupFic(Nom, st, LaTable: string): boolean;
var
  T, TA: TOB;
  Q: TQuery;
  NomFic, MSal, LeSal, SQL, Prefixe: string;
  i, LeMAx: Integer;
begin
  NomFic := VH_Paie.PGCHEMINEAGL + Nom;
  result := FALSE;
  // PT2 : 08/04/2002 V571 PH Controle des noms de fichier
  AnalyseNomFic(NomFic);
  Prefixe := '';
  if not FileExists(NomFic) then
  begin
    PGIBox('Fichier inaccessible : ' + NomFic, 'Abandon du traitement');
    Exit;
  end;
  T := tob.create('', nil, -1);
  if (LaTable <> 'ABSENCESALARIE') and (LaTable <> 'PARAMSOC') and (LaTable <> 'CHOIXCOD') then
  begin // annulation  des tables
    ExecuteSQl('DELETE FROM ' + LaTable);
  end
  else
    ExecuteSQL(St);
  // Chargement de la TOB à partir du fichier
  TobLoadfromfile(NomFic, nil, T);
  // PT1 : 31/01/2002 V571 PH réaffectation des numéros d'ordre et des codes de regroupement
  if LaTable = 'ABSENCESALARIE' then
  begin
    LeSal := '';
    Msal := '';
    LeMax := 0;
    for i := 0 to T.Detail.count - 1 do
    begin
      TA := T.detail[i];
      if TA <> nil then
      begin
        LeSal := TA.getvalue('PCN_SALARIE');
        if LeSal <> Msal then
        begin
          SQL := 'SELECT MAX(PCN_ORDRE) MONMAX FROM ABSENCESALARIE WHERE PCN_SALARIE="' + LeSal + '"';
          Q := OpenSql(SQL, TRUE);
          if not Q.eOF then LeMax := Q.FindField('MONMAX').AsInteger + 1
          else LeMax := 1;
          Ferme(Q);
          MSal := LeSal;
        end;
        TA.putvalue('PCN_ORDRE', LeMax);
        TA.putvalue('PCN_CODERGRPT', LeMax);
        // PT3 : 08/04/2002 V571 PH Pas de gestion des confidentiels dans la base eConges
        TA.putValue('PCN_CONFIDENTIEL', '0');
        LeMax := LeMax + 1;
      end;
    end;
  end;
  // Fin PT1 : 31/01/2002 V571 PH
  // PT4 : 08/04/2002 V571 PH Analyse des donnees contenus dans le fichier pour eliminer les enrg à blanc
  if LaTable = 'SALARIES' then Prefixe := 'PSA'
  else if LaTable = 'DEPORTSAL' then Prefixe := 'PSE'
  else if LaTable = 'RECAPSAL' then Prefixe := 'PRS';
  if Prefixe <> '' then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      TA := T.detail[i];
      if TA <> nil then
      begin
        LeSal := TA.getvalue(Prefixe + '_SALARIE');
        if LeSal = '' then TA.Free;
      end;
    end;
  end;
  // Transaction ecriture dans la base
  try
    BeginTrans;
    if (T <> nil) and (T.Detail.count >= 1) then T.InsertDB(nil, false);
    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue sur le fichier ' + NomFic, 'Traitement du fichier import');
    TEvent.add('Une erreur est survenue sur le fichier ' + NomFic); //PT8
  end;
  T.free;
  Result := TRUE;
end;
// PT2 : 08/04/2002 V571 PH Controle du nom de fichier visant à remplacer \\ par \ dans le no du fichier
// si \\ ne se trouve pas au debut de la chaine

procedure TOF_ElisteTables.AnalyseNomFic(var NomFic: string);
var
  i, j: Integer;
  Chaine: string;
begin
  Chaine := '';
  j := Length(NomFic) - 2;
  for i := 3 to j do
  begin
    if Copy(NomFic, i, 2) = '\\' then
    begin
      Chaine := Copy(NomFic, 1, i - 1) + Copy(NomFic, i + 1, j + 2 - i);
      break;
    end;
  end;
  if Chaine <> '' then NomFic := Chaine;
end;
//DEB PT8 Ajout Onclose

procedure TOF_ElisteTables.OnClose;
begin
  inherited;
  if TEvent <> nil then TEvent.Free;
end;
//FIN PT8

initialization
  registerclasses([TOF_ElisteTables]);
end.

