{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 04/09/2001

Modifié le ... :   /  /
Description .. : Unit de lancement de la préparation automatique des
Suite ........ : bulletins
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************
PT1   : 10/09/2001 V547 SB Fiche de bug n°266
                           Gestion des msg d'erreur CP en prep auto
PT2   : 04/10/2001 V562 SB Modification du chemin de sauvegarde
PT3   : 23/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et
                           Dates edition
PT4   : 25/10/2001 V562 PH test de presence d'un bulletin complémentaire sur les
                           mêmes critères
PT5   : 18/12/2001 V571 PH controle dates edition par rapport à l'etablissement
PT6   : 31/01/2002 V571 SB Ouverture du fichier .log en fin de préparation
PT7   : 26/03/2002 V571 PH Raz structure entete ChpEntete et traitement
                           reglement
PT8   : 15/05/2002 V582 VG IFDEF CCS3, limitation à 30 bulletins par période
                          (date à date)
PT9   : 03/03/2002 V582 PH Gestion historique des évènements
PT10  : 01/07/2002 V585 VG PT8, finalement non !!!
PT11    04/09/2002 V582 SB Vidage des tobs CP : Correction bug mvt topé S pour salarié non sortie
PT12    17/10/2002 SB V585 FQ n°10235 Intégration du calcul des variables 28,29 heures ouvrées ouvrables
PT13    19/12/2002 SB V591 FQ 10391 Nombre de salariés traités erroné
PT14    19/12/2002 PH V591 Affectation de la date de paiement à la date de fin de paie si non renseignée
PT15    31/01/2003 PH V591 Barre de patience pdt la prépa
PT16    13/05/2003 PH V_42 Optimisation chargement,Vidage tobs de la paie non fait systématiquement
PT17    24/06/2003 SB V_42 FQ 10454 Calcul au maintien sur premier bulletin implique un recalcul des CP
PT18    26/08/2003 PH V_42 Mise en place saisie arret
PT19    05/12/2003 PH V620 Automatisme de calcul de la paie à l'envers en fonction de la saisie par rubrique
                           Détection des rubriques de rémunérations dont le thème est NetAPayer
PT20    08/12/2003 PH V620 Préparation automatique de bulletins complémentaires
PT21    28/06/2004 MF V_50 Ajout FreeAndNil sur Tob_Abs
PT22    23/08/2004 PH V_50 FQ 11459 Limitation des trentièmes si 2 bulletins dans le même mois
PT23    10/09/2004 MF V_50 Intégration IJSS et maintien
PT24    23/09/2004 MF V_50 Intégration IJSS et maintien
PT25    12/10/2004 MF v_50 Intégration IJSS et maintien : FQ 11673  IJSS uniquement
PT26    26/10/2004 PH V_60 RAZ Tob des absences du salarie
PT27    30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
PT28    06/12/2004 PH V_60 FQ 11831 Erreur SQL spécif DB2
PT29    05/01/2005 PH V_60 Traitement des paies aux extras payées à la semaine en NET avec récup des commentaires
PT30    28/01/2005 MF V_60 IJSS & maintien: Ajout de 2 paramètres à l'appel de RecupereRegltIJSS
PT31    17/05/2005 PH V_60 FQ 11801 gestion arrondi paie à l'envers
PT32    31/05/2005 PH V_60 FQ 12337 Gestion des erreurs lors de la validation du bulletin
PT33    05/07/2005 PH V_60 FQ 12416 Réécriture date intégration dans la paie des lignes provenant de la saisie par rubrique
PT34    01/08/2005 PH V_60 FQ 12467 Cas récup du Net A payer si le salarié rentre dans la mois
PT35    23/02/2006 MF V_65 Correction memcheck
PT35    10/04/2006 PH V_65 FQ 13037 Controle existance buletin sur un autre étab aux mêmes dates
PT36    24/04/2006 MF V_65 FQ 13077 : modification du message d'anomalie qd rub. de garantie absente
PT37	04/05/2006 SB V_65 FQ 12786 Anomalie calcul variable 0027 en prep. auto
PT38    29/12/2006 PH V_70 Integration automatique des rubriques
PT39    02/02/2007 PH V_70 Paies au contrat et Renommage ActionBul en ActionBulletin
PT40    14/03/2007 FC V_70 Gestion de la table PGSYNELTNAT pour stocker les éléments nationaux manquants détectés
                           Si SO_PGGESTELTDYNDOS à Vrai
                           Gestion des éléments dynamiques dossier (chargement des tob)
PT41   27/03/2007 PH Rajout paramètre fonction RecupTobSalarie pour prise en compte Historique date
PT42   21/06/2007 PH FQ 14256, on ne recalcule pas les bulletins clos
PT43   18/07/2007 MF V_72 FQ 14588 : on encadre le nom du fichier (Chemin + Fichier) par des guillemets
PT44   20/08/2007 PH FQ 14623  automatisme paie envers sur rubrique de net
PT45   29/10/2007 NA Si Gestion de la présence : mise à jour de la table PRESENCESALARIE si intégration des rubriques dans la paie
PT46 : 26/11/2007 GGU V_80 Planificateur de tâches multidossier
PT49  : 08/09/2008 JS Optimisation GBL
PT50 : 12/09/2008 Ph  FQ 15727 Fin de paie = fin de mois si pas de date de fin de contrat
PT51 : 02/10/2008 MF FQ 15781 : pas de calcul du maintien qd ancienneté <= 01/01/1900
}
unit UTofPG_PrepAuto;

interface
uses Windows,
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
  Dialogs,
  ed_Tools,
{$IFNDEF EAGLSERVER}
  UTOF,
  utobdebug,
  Vierge,
  AGLInit,
{$IFDEF EAGLCLIENT}
  UtileAGL,
  eFiche,
  MainEAGL, //PT40 (AglLanceFiche)
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  Qre,
  FE_Main,  //PT40 (AglLanceFiche)
{$ENDIF}
{$ENDIF}
{$IFDEF EAGLSERVER}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Grids,
  HCtrls,
  HEnt1,
  EntPaie,
  HMsgBox,
  UTOB,
  P5Util,
  P5Def,
  SaisBul,
  PgCongesPayes,
  uPaieEltDynSal, //PT40
  uPaieEltDynEtab,//PT40
  uPaieEltDynPop, //PT40
{$IFNDEF CPS1}
  PGPOPULOUTILS,  //PT40
{$ENDIF}
{$IFNDEF CCS3}
  PGIJSSMaintien, // PT24
{$ENDIF}
  uPaieEtabCompl,
  HQry,
  PgOutils,
  PgOutils2,
  ShellAPI;
{$IFNDEF EAGLSERVER}
type
  TOF_PG_PrepAuto = class(TOF)
  private
    BtnLance, BTNPROCESS: TToolbarButton97;
    procedure LancePrepAuto(Sender: TObject);
    procedure ImprimeClick(Sender: TObject);
    procedure ParamPrepAuto(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnBPROGMULTIDOSSIERClick(Sender: TObject);//PT46
  end;
{$ENDIF}
function Process_PREPAUTO(D1, D2, LeMode: string; LeSQL: string): TOB;
procedure Lancetraitement(Ecran: TForm);
procedure RecupCasPaie(var ActionBulletin: TActionBulletin);

implementation

uses
  UTofPG_PaieEnvers,
  Paramsoc,
{$IFNDEF EAGLSERVER}
  uFicheJob,
{$ENDIF}
  P5RecupInfos, DB
{$IFNDEF EAGLSERVER}
  , PGFicheJobMultiDossier     //PT46
{$ENDIF}
  ;


var
  QMul: TQUERY; // Query recuperee du mul
{$IFDEF EAGLSERVER}
  ZD, ZF: TDateTime;
{$ENDIF}
  LblDD, LblDF: THLabel;
  Numerateur, Denominateur: Integer; // Du Trentieme
  Ss : String;

function Process_PREPAUTO(D1, D2, LeMode : string; LeSQL: string): TOB;
{$IFDEF EAGLSERVER}
var TobResult: TOB;
    StSQL : String;
{$ENDIF}
begin
{$IFDEF EAGLSERVER}
  DateD := StrToDate(D1);
  DateF := StrToDate(D2);
  ZD := DateD;
  ZF := DateF;
  TOBResult := TOB.Create('PRERAUTOPROCESS', nil, -1);
  if LeMode = 'CWAS' then StSQL := LeSQL
  else StSQL := '@@' + LeSQL;
  QMul := OpenSQL(StSQL, TRUE);
  if not QMUL.EOF then
  begin
    LanceTraitement(nil); ;
    TOBResult.AddChampSupValeur('OKTRAITE', TRUE);
  end
  else TOBResult.AddChampSupValeur('OKTRAITE', FALSE);
  result := TOBResult;
  FERME(QMUL);
{$ENDIF}
end;


{$IFNDEF EAGLSERVER}

procedure TOF_PG_PrepAuto.LancePrepAuto(Sender: TObject);
begin
  LanceTraitement(Ecran);
end;
{$ENDIF}

procedure Lancetraitement(Ecran: TForm);
var
  ActionBulletin: TActionBulletin; // PT39
  TOB_Rub, TOBAcquis, TOBPris, T_Sal, MonEtab, T: TOB;
  Salarie, Info, St, StMsgErr, FileN, stq, LaRubriq, LeThemRem : string; // PT44
  NetAPAYER, PlusApprochant, LeNet : Double; // PT44
{$IFNDEF EAGLSERVER}
  Pan: TPageControl;
  Tbsht: TTabSheet;
  Trace: TListBox;
{$ELSE}
  Trace, TraceE: TStringList;
{$ENDIF}
  DateE, DateS: TDateTime;
  SystemTime0: TSystemTime;
  CP, ErreurAnal: boolean;
  Q: TQuery;
  Compteur, i, zz: Integer;
  F: TextFile;
  TheErrBulCompl: Boolean;
  Tot: Integer;
  TLig: TOB;
  RetEnvers, okok: Boolean;
  GenereMaintien, IJSS: Boolean; //PT23
  Tob_IJSS, Tob_Maintien: TOB; // PT23
  TAbLibel: array[1..9] of string; // Tableau de commentaires liés PT29
  kk: Integer; // PT29
  LeLib: string;
{$IFNDEF CCS3}
  MaintienCount: Boolean;
  mode, ChampCateg: string; //PT23
  anomalie: integer; //PT23
  TR: TOB; // PT23
{$ENDIF}
  EntreePer: Boolean; //PT34
  LaDateD: TDateTime; //PT34
  NbEltManquant: Integer; //PT40
begin

  //  LeMode := 'A';
  GenereMaintien := False; //PT23

  if (VH_Paie.PGGestIJSS = true) then
    IJSS := True //PT23
  else
    IJSS := false;
{$IFDEF EAGLSERVER}
  V_PGI.eAGLDebugMode := TRUE;
{$ENDIF}

{$IFNDEF CCS3}
  MaintienCount := False;
{$ENDIF}
  Salarie := '';
  Etab := '';
  zz := 0;
{$IFNDEF EAGLSERVER}
  Pan := TPageControl(PGGetControl('PANELPREP', Ecran));
  Tbsht := TTabSheet(PGGetControl('TBSHTTRACE', Ecran));
  Trace := TListBox(PGGetControl('LSTBXTRACE', Ecran));
  TraceErr := TListBox(PGGetControl('LSTBXERROR', Ecran));
  Trace.Clear;
  TraceErr.Clear;
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La préparation automatique ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
    exit;
  end;
//  BtnLance.Enabled := FALSE;
  if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
  DateD := StrToDate(LblDD.Caption);
  DateF := StrToDate(LblDF.Caption);
  if QMul = nil then
  begin
    PGIBox('Erreur sur la liste des salariés', 'Abandon du traitement');
    exit;
  end;
{$ENDIF}
  InitLesTOBPaie; // Creation des Tob de la paie pour effectuer les calculs
{$IFNDEF EAGLSERVER}
  Trace.Items.Add('Préparation automatique des paies du ' + LblDD.Caption + ' au ' + LblDF.Caption);
  Trace.Items.Add('Chargement des informations nécessaires aux calculs des paies');
{$ELSE}
  Trace := TStringList.Create;
  TraceE := TStringList.Create;
  Trace.Add('Préparation automatique lancée par process serveur détaché');
  Trace.Add('Préparation automatique des paies du ' + DateToStr(ZD) + ' au ' + DateToStr(ZF));
  Trace.Add('Chargement des informations nécessaires aux calculs des paies');
{$ENDIF}
  ChargeLesTOBPaie; // Chargement des TOB de la paie
  ChargeLesExercPaie(DateD, DateF); // Chargement des Exercices de La PAIE
{$IFNDEF EAGLSERVER}
  Trace.Items.Add('Fin de Chargement des informations nécessaires aux calculs des paies');
{$ELSE}
  Trace.Add('Fin de Chargement des informations nécessaires aux calculs des paies');
{$ENDIF}
  TOB_Rub := nil;
{$IFNDEF CCS3}
  // récupération du champ critère suplémentaire IJSS et Maintien
  if (VH_Paie.PGCritMaintien <> '') then
  begin // PT28 Suppression des "" pour un integer en DB2
    Q := OpenSQL('SELECT PAI_PREFIX, PAI_SUFFIX FROM PAIEPARIM WHERE PAI_IDENT = ' +
      VH_Paie.PGCritMaintien, TRUE);
    if not Q.EOF then
    begin
      ChampCateg := Q.Fields[0].AsString + '_' + Q.Fields[1].AsString;
    end
    else
      ChampCateg := '';
    ferme(Q);
  end;
{$ENDIF}

  QMul.First;
  GetLocalTime(SystemTime0);
{$IFNDEF EAGLSERVER}
  Trace.Items.Add('Début du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
{$ELSE}
  Trace.Add('Début du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
{$ENDIF}

  // PT15    31/01/2003 PH V591 Barre de patience pdt la prépa
{$IFNDEF EAGLSERVER}
  Tot := QMUL.RecordCount;
  InitMoveProgressForm(nil, 'Calcul des paies en cours ...', 'Veuillez patienter SVP ...', Tot, FALSE, TRUE);
{$ENDIF}

  while not QMul.EOF do
  begin
    // Réinitialisation d'un top qui permet de savoir si le bulletin doit être validé ou non
    ValidationOK := True;
    if TOB_HistoBasesCot <> nil then
    begin
      TOB_HistoBasesCot.Free;
      TOB_HistoBasesCot := nil;
    end;

    // d PT23
{$IFNDEF EAGLSERVER}
{$IFNDEF CCS3}
    if (PGGetControl('MAINTIEN', Ecran) <> nil) and TCheckBox(PGGetControl('MAINTIEN', Ecran)).Checked then
    begin
      GenereMaintien := True;
      IJSS := False;
    end;
{$ENDIF}
{$ENDIF}
    // f PT23

        // PT15    31/01/2003 PH V591 Barre de patience pdt la prépa
{$IFNDEF EAGLSERVER}
    zz := ZZ + 1;
    MoveCurProgressForm(IntToStr(zz) + '/' + IntToStr(Tot));
{$ENDIF}
    Salarie := QMul.FindField('PSA_SALARIE').AsString;
    if (ClePGSynEltNAt = '') then  // PT40 Une seule valeur dans la clé PGSYNELTNAT, le 1er salarié rencontré
      ClePGSynEltNAt := Salarie;
    CodeSalarie := Salarie; // Memorisation du code salarie en variable globale du moteur de la paie
    Etab := QMul.FindField('PSA_ETABLISSEMENT').AsString;
    // DEB PT39
    if (BullContrat <> 'X') OR
    ((BullContrat = 'X') AND (GetParamSocSecur ('SO_PGMETHODCONTRAT', '001') <> '002')) then
    begin // Cas Plusieurs contrats dans le même mois ou dates de contrat à cheval sur plusieurs paie
{$IFNDEF EAGLSERVER}
      DateD := StrToDate(LblDD.Caption);
      DateF := StrToDate(LblDF.Caption);
{$ELSE}
      DateD := ZD;
      DateF := ZF;
{$ENDIF}
      DateE := QMul.FindField('PSA_DATEENTREE').AsDateTime;
    // DEB PT34
      if DateE > DateD then
      begin
        DateD := DateE; // cas du salarié entré en cours de période
        EntreePer := TRUE;
      end
      else EntreePer := FALSE;
    // FIN PT34
      DateS := QMul.FindField('PSA_DATESORTIE').AsDateTime;
      if (DateS < DateF) and (DateS > 100) then DateF := DateS; // cas du salarié sorti en cours de période
    end
    else
    begin
      DateD := QMul.FindField('PCI_DEBUTCONTRAT').AsDateTime;
      DateF := QMul.FindField('PCI_FINCONTRAT').AsDateTime;
    end;
    // FIN PT39
    if (DateF <= iDate1900) OR (DateF < DateD) then
      DateF := FINDEMOIS (DateD);  // PT50
    if Tob_Rub <> nil then
    begin
      Tob_Rub.Free;
      Tob_Rub := nil;
    end;
    if TOBAna <> nil then
    begin
      TOBAna.Free;
      TOBAna := nil;
    end;
    if TOB_DuSalarie <> nil then
    begin
      TOB_DuSalarie.free;
      TOB_DuSalarie := nil;
    end;
    // DEB PT26
    if TOB_Abs <> nil then
    begin
      TOB_Abs.Free;
      TOB_Abs := nil;
    end;
    // FIN PT26
    if not TheErrBulCompl then
    begin
      RecupTobSalarie(Salarie, DateD, DateF); // PT41 Chargement de la TOB salarie

    // PT4 25/10/2001 V562  PH test de presence d'un bulletin complémentaire sur les mêmes critères
      TheErrBulCompl := FALSE;
      Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND ' +
        'PPU_SALARIE="' + Salarie + '" AND AND PPU_DATEDEBUT="' + UsDateTime(DateD) +
        '" AND PPU_DATEFIN="' + UsDateTime(DateF) + '" AND PPU_BULCOMPL="X"';
      Q := Opensql(stq, true);
      if not Q.eof then
      begin
{$IFNDEF EAGLSERVER}
        TraceErr.Items.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car il possède déjà un bulletin complémentaire sur la même période');
{$ELSE}
      ddwriteln ('trace Bulletin impossible');
      readln (ss);
      TraceE.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car il possède déjà un bulletin complémentaire sur la même période');
{$ENDIF}
        TheErrBulCompl := TRUE;
      end;
      Ferme(Q);
    end;
    // DEB PT35
    if not TheErrBulCompl then
    begin
      Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT<>"' + Etab + '" AND ' +
        'PPU_SALARIE="' + Salarie + '" AND AND PPU_DATEDEBUT="' + UsDateTime(DateD) +
        '" AND PPU_DATEFIN="' + UsDateTime(DateF) + '" AND PPU_BULCOMPL="-"';
      Q := Opensql(stq, true);
      if not Q.eof then
      begin
{$IFNDEF EAGLSERVER}
        TraceErr.Items.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car il possède déjà un bulletin pour un autre établissement sur la même période');
{$ELSE}
        TraceE.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car il possède déjà un bulletin pour un autre établissement sur la même période');
{$ENDIF}
        TheErrBulCompl := TRUE;
      end;
      Ferme(Q);
    end;
    // FIN PT35
    // DEB PT42
    if not TheErrBulCompl then
    begin
      Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT = "' + Etab + '" AND ' +
        'PPU_SALARIE="' + Salarie + '" AND AND PPU_DATEDEBUT="' + UsDateTime(DateD) +
        '" AND PPU_DATEFIN="' + UsDateTime(DateF) + '" AND PPU_TOPCLOTURE="X"';
      Q := Opensql(stq, true);
      if not Q.eof then
      begin
{$IFNDEF EAGLSERVER}
        TraceErr.Items.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car le bulletin est clos');
{$ELSE}
        TraceE.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car  le bulletin est clos');
{$ENDIF}
        TheErrBulCompl := TRUE;
      end;
      Ferme(Q);
    end;
    // FIN PT42
    // PT8
{$IFDEF CCS3}
    { PT10
    Stq := 'SELECT COUNT(PPU_SALARIE) AS NBPAYE'+
            ' FROM PAIEENCOURS WHERE'+
            ' PPU_DATEDEBUT="'+UsDateTime(DateD)+'" AND'+
            ' PPU_DATEFIN="'+UsDateTime(DateF)+'"';
    Q := Opensql (stq,true);
    if Q.FindField('NBPAYE').AsString <>'' then
        Compteur := Q.FindField('NBPAYE').AsInteger
    else
        Compteur := 0;
    if Compteur > 29 then
        begin
        TraceErr.Items.Add('Bulletin impossible pour le salarié : '+ Salarie + ' car le nombre de bulletin est limité à 30');
        TheErrBulCompl := TRUE;
        end;
    Ferme (Q);
    FIN PT10}
{$ENDIF}
    // FIN PT8
    if not TheErrBulCompl then
    begin
      InitialiseVariableStatCP;

      //DEB PT40
      if GetParamSocSecur('SO_PGGESTELTDYNDOS', False) then
      begin
        Nettoyage_EltDynSal();
        Nettoyage_EltDynPop();
        Nettoyage_EltDynEtab();
        initTOB_EltDynSal(Salarie, DateF);
        initTOB_EltDynEtab(Etab, DateF);
        initTOB_EltDynPop(Salarie, DateF);
      end;
      //FIN PT40

      // chargement des la liste des rubriques en foction des profils
      CreationTOBCumSal; // Creation de la TOB des Cumuls Salariés
      RecupCasPaie(ActionBulletin); // Chargement des Infos entete bulletin PT39
      Date1 := DateD; // Memorisation des Dates pour le cas creation idem saisie de bulletin
      Date2 := DateF;
      ErreurAnal := FALSE;
      if (VH_Paie.PGAnalytique = TRUE) then
      begin
        St := 'SELECT V_NATURE FROM VENTIL WHERE V_NATURE LIKE "SA%" AND V_COMPTE="' + Salarie + '"';
        Q := OpenSql(st, TRUE);
        if Q.EOF then ErreurAnal := TRUE;
        Ferme(Q);
      end;

      if not ErreurAnal then
      begin
        // PT18    26/08/2003 PH V_42 Mise en place saisie arret
        FreeAndNil(TOB_SaisieArret);
        // PT3 23/10/2001 V562  PH Gestion cas particulier du bulletin complémentaire et Dates edition
        Tob_Rub := ChargeRubriqueSalarie(Tob_Salarie, DateD, DateF, ActionBulletin, 0, ProfilSpec); //PT39
        if assigned(tob_rub) then
        begin
          ActionBulCP := ActionBulletin; //PT39
          ActionBul := ActionBulletin;   //PT39
          ChargeBasesSal(Salarie, Etab, DateD, DateF);
          // Gestion des Modules Complémentaires : CP, Gestion des Absences, Saisie par rubrique ==> integration automatique des éléments de la paie
          MonEtab := Paie_RechercheOptimise(Tob_etablissement, 'ETB_ETABLISSEMENT', Tob_Salarie.getvalue('PSA_ETABLISSEMENT'));
          if MonEtab <> nil then
            CP := (MonEtab.getvalue('ETB_CONGESPAYES') = 'X');
          TopRecalculCPBull := False;
          NetAPAYER := 0;
          LeThemRem := ''; // PT44
          // PT20    08/12/2003 PH V620 Préparation automatique de bulletins complémentaires donc pas de calcul
          //                            de paie à l'envers en fct du net à payer de la saisie par rubrique
          if BullCompl <> 'X' then
          begin
            //PT19    05/12/2003 PH V620 Automatisme de calcul de la paie à l'envers en focntion de la saisie par rubrique
            // Cas du net à payer automatique Attention nouveau themerem="ZZZ"
            // DEB PT34
            if not EntreePer then LaDateD := DateD
            else LaDateD := DebutdeMois(DateD);
            // FIN PT34
            St := 'SELECT PSD_MONTANT,PSD_RUBRIQUE,PSD_LIBELLE,PRM_THEMEREM FROM HISTOSAISRUB LEFT JOIN REMUNERATION ON PSD_RUBRIQUE=PRM_RUBRIQUE' +
              ' WHERE PSD_SALARIE = "' + Salarie + '" AND (PSD_DATEDEBUT >="' + usdatetime(LaDateD) +
              '" AND PSD_DATEFIN <= "' + usdatetime(Datef) + '") AND PSD_MONTANT > 0 AND (PRM_THEMEREM = "ZZZ" OR PRM_THEMEREM = "ZZY") ORDER BY PSD_DATEDEBUT'; // PT44
            Q := OpenSql(st, TRUE);
            // DEB PT29
            for kk := 1 to 9 do
              TabLibel[kk] := '';
            kk := 1;
            while not Q.EOF do
            begin
              NetAPAYER := NetAPAYER + Q.FindField('PSD_MONTANT').AsFloat;
              LaRubriq := Q.FindField('PSD_RUBRIQUE').AsString;
              LeThemrem := Q.FindField('PRM_THEMEREM').AsString; // PT44
              if (Q.FindField('PSD_LIBELLE').AsString <> '') AND (LeThemrem = 'ZZZ') then // PT44
              begin
                TabLibel[kk] := Q.FindField('PSD_LIBELLE').AsString;
                kk := kk + 1;
              end;
              Q.Next;
            end;
            // FIN PT29
            Ferme(Q);
            // Fin PT19
          end;
          // FIN PT20
// d PT23 Intégration IJSS & maintien
{$IFNDEF CCS3}
          // vérif. si maintien à intégrer
          if (GenereMaintien) then
            // On gère le maintien
          begin
            if ActionBulletin = TaCreation then Mode := 'C' else Mode := 'M'; //PT39
            //PT27            TR :=  (RecupereRegltIJSS(Datef, CodeSalarie, Mode));
            TR := (RecupereRegltIJSS(Datef, CodeSalarie, Mode, ChampCateg, nil, nil)); // PT30
            if not RecupereAbsencesIJSS(Datef, CodeSalarie, Mode) and
              (TR.Detail.Count <= 0) then
              // Pas d'absence ni de règlement d'IJSS pour ce salarié donc pas
              // de maintien à intégrer
            begin
              GenereMaintien := False;
            end;
            FreeAndNil(TR); //PT35 memcheck
          end;
{$ENDIF}
          // f PT23
          //d PT23  PT25
          //        RecupInfoModules*(TOB_Rub, Tob_Salarie, ActionBul, true, CP, StMsgErr); //DEB1 PT1 Ajout StMsgErr
          //        RecupInfoModules(TOB_Rub, Tob_Salarie, ActionBul, true, CP, StMsgErr,GenereMaintien,IJSS);
          RecupInfoModules(TOB_Rub, Tob_Salarie, ActionBulletin, true, CP, StMsgErr, GenereMaintien, IJSS, Tob_IJSS); //PT39
          // f PT23 PT25
                    //PT19    05/12/2003 PH V620 Automatisme de calcul de la paie à l'envers en focntion de la saisie par rubrique
          if NetAPayer > 0 then
          begin
            if LeThemrem = 'ZZZ' then // PT44
              LaRubriq := RechDom('PGREMNETAPAYER', LaRubriq, FALSE); // On recupère la rubrique associée uniquement si Net A payer global
            if (LaRubriq <> '') and (LaRubriq <> 'Error') then
            begin
              Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
              if TLig = nil then
              begin // on insere la rubrique dans le bulletin
                ChargeProfilSPR(Tob_salarie, TOB_Rub, LaRubriq, 'AAA');
                Tlig := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', LaRubriq], FALSE);
              end;
              // PT23              if TLig <> nil then TLig.putValue('PHB_MONTANT', 0); // Pour remettre à 0 la valeur de la ligne pour le calcul de la paie à l'envers
              if TLig <> nil then TLig.putValue('PHB_MTREM', 0); // Pour remettre à 0 la valeur de la ligne pour le calcul de la paie à l'envers
              // DEB PT29
              for kk := 1 to 9 do
              begin
                if TabLibel[kk] <> '' then
                  IntegreRubCommentaire(TOB_Rub, nil, Tob_Salarie, 'SRB', LaRubriq, CodeSalarie, Etab, DateD, DateF, TabLibel[kk]);
              end;
              // FIN PT29
            end;
          end
          else TLig := nil;
          // FIN PT19
          if StMsgErr <> '' then
{$IFNDEF EAGLSERVER}
            TraceErr.Items.Add(StMsgErr); //FIN1 PT1
{$ELSE}
            TraceE.Add(StMsgErr); //FIN1 PT1
{$ENDIF}
// DEB PT38
          IntegreAuto := FALSE;
          CalculBulletin(TOB_Rub);
          if IntegreAuto then CalculBulletin(TOB_Rub);
// FIN PT38
          { DEB PT17 Recalcul CP si Calcul maintien CP=0 et cumul12 <> 0 }
          if (TopRecalculCPBull) and (RendCumulSalSess('12') <> 0) then
          begin
            AnnuleCongesPris(Tob_salarie.GetValeur(iPSA_SALARIE), Tob_salarie.GetValeur(iPSA_ETABLISSEMENT), DateD, DateF);
            Tob_Pris := SalIntegreCP(Tob_salarie, TOB_Rub, T_MvtAcquis, DateD, DateF, False, StMsgErr);
            CalculBulletin(TOB_Rub);
          end;
          { FIN PT17 }
// PT19    05/12/2003 PH V620 Automatisme de calcul de la paie à l'envers en fonction de la saisie par rubrique
          RetEnvers := TRUE;
// DEB PT44
          if LeThemrem = 'ZZY' then
          begin
            if VH_Paie.PGEnversNet = 'NET' then Lenet := RendCumulSalSess('10')
            else if VH_Paie.PGEnversNet = 'BRU' then lenet := RendCumulSalSess('01')
            else if VH_Paie.PGEnversNet = 'BRH' then lenet := RendCumulSalSess('05')
            else lenet := RendCumulSalSess('07') + RendCumulSalSess('08') + RendCumulSalSess('10');
            NetAPayer := NetAPayer + LeNet;
          end;
// FIN PT44
          if (NetAPAYER > 0) and (LaRubriq <> '') then
          begin
            if TLig <> nil then
            begin
              PlusApprochant := 0;
              RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, NetAPayer, LaRubriq, TRUE);
// DEB PT31 On relance le calcul avec arrondi si echec
              if (not RetEnvers) and (VH_PAIE.PGArrondiEnvers > 0) then
                RetEnvers := CalculPEnvers(PlusApprochant, TOB_Rub, TLig, NetAPayer, LaRubriq, TRUE, TRUE);
// FIN PT31
              if not RetEnvers then
{$IFNDEF EAGLSERVER}
                if TraceErr <> nil then
                begin
                  TraceErr.Items.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car le calcul de la paie à l''envers a échoué');
                  TraceErr.Items.Add('La valeur la plus approchante trouvée est : ' + DoubleToCell(PlusApprochant, 2));
                end;
{$ELSE}
                if TraceE <> nil then
                begin
                  TraceE.Add('Bulletin impossible pour le salarié : ' + Salarie + ' car le calcul de la paie à l''envers a échoué');
                  TraceE.Add('La valeur la plus approchante trouvée est : ' + DoubleToCell(PlusApprochant, 2));
                end;
{$ENDIF}
            end;
          end;
          // FIN PT19
// d PT23
{$IFNDEF CCS3}
          if GenereMaintien then
            // maintien à calculer et intégrer
          begin
            CalculBulletin(TOB_Rub);

            MaintienCount := True;
            Anomalie := 0;

            // calcul et maj Maintien   (paie envers...)
            MaintienDuSalarie(Tob_salarie, TOB_Rub, Tob_etablissement, Tob_Abs, DateToStr(Datef), DateToStr(dated), Mode, Anomalie, ChampCateg, TOB_IJSS, TOB_Maintien);

            if (anomalie <> 0) then
            begin
              if (Anomalie = 1) then
{$IFNDEF EAGLSERVER}
                TraceErr.Items.Add(CodeSalarie + ' : Pas de règle de Maintien')
{$ELSE}
                TraceE.Add(CodeSalarie + ' : Pas de règle de Maintien')
{$ENDIF}
              else
                if (Anomalie = 2) then
{$IFNDEF EAGLSERVER}
                  TraceErr.Items.Add('Calcul du montant du maintien : Bulletin impossible pour le salarié : '
                    + Salarie +
                    ' car le calcul de la paie à l''envers a échoué')
{$ELSE}
                  TraceE.Add('Calcul du montant du maintien : Bulletin impossible pour le salarié : '
                    + Salarie +
                    ' car le calcul de la paie à l''envers a échoué')
{$ENDIF}
// d PT36
                else
                  if (Anomalie = 3) then
{$IFNDEF EAGLSERVER}
                    TraceErr.Items.Add(CodeSalarie + ' : Vérifiez le maintien de salaire et les ijss, la rubrique de garantie de salarie n''a pu être calculée ou est absente')
{$ELSE}
                    TraceE.Add(CodeSalarie + ' : Vérifiez le maintien de salaire et les ijss, la rubrique de garantie de salarie n''a pu être calculée ou est absente')
// d PT51
{$ENDIF}          else
                    if (Anomalie = 4) then
{$IFNDEF EAGLSERVER}
                      TraceErr.Items.Add(CodeSalarie + ' La date d''ancienneté du salarié n''est pas renseignée : le maintien ne peut pas être calculé');
{$ELSE}
                      TraceE.Add(CodeSalarie + ' La date d''ancienneté du salarié n''est pas renseignée : le maintien ne peut pas être calculé');
{$ENDIF}
// f PT51
// f PT36
            end;

            // intégration abs IJSS, IJSS, Maintien
            CalculBulletin(TOB_Rub);
          end;
{$ENDIF}
          // f PT23
                    // Traitement de l'analytique
          if (VH_Paie.PGAnalytique = TRUE) then
          begin
            if ActionBulletin = taModification then //PT39
            begin
              TOBAna := PreChargeVentileLignePaie(CodeSalarie, DateD, DateF); // Prechargement des pre-ventilations analytiques
              ControlAffecteAnal(CodeSalarie, DateD, DateF, TOB_VenRem, TOB_VenCot, TOB_Rub, TOBAna);
            end
            else TOBAna := PreVentileLignePaie(TOB_VenRem, TOB_VenCot, TOB_Rub, Salarie, 'PRE', DateD, DateF);
            // Reaffectation des ventilations analytiques des cotisations en focntion de la valorisation des remunerations
            PGReaffSectionAnal(TOB_Rub, TOBAna);
          end;
{$IFNDEF CCS3}
          if MaintienCount then
          begin
            LeLib := 'Salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
              Tob_Salarie.GetValue('PSA_PRENOM') + ' Etablissement : ' +
              RechDom('TTETABLISSEMENT', Tob_Salarie.GetValue('PSA_ETABLISSEMENT'), FALSE) +
              ' Traitement du maintien ';
{$IFNDEF EAGLSERVER}
            Trace.Items.Add(LeLib);
{$ELSE}
            Trace.Add(LeLib);
{$ENDIF}
            MaintienCount := False;
          end
          else
{$ENDIF}
{$IFNDEF EAGLSERVER}
            Trace.Items.Add('Salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
              Tob_Salarie.GetValue('PSA_PRENOM') + ' Etablissement : ' + RechDom('TTETABLISSEMENT', Tob_Salarie.GetValue('PSA_ETABLISSEMENT'), FALSE) + ' ');
{$ELSE}
            Trace.Add('Salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
              Tob_Salarie.GetValue('PSA_PRENOM') + ' Etablissement : ' + RechDom('TTETABLISSEMENT', Tob_Salarie.GetValue('PSA_ETABLISSEMENT'), FALSE) + ' ');
{$ENDIF}
          // Ecriture du bulletin et des Mvts Cp acquis ....
//PT19    05/12/2003 PH V620 Automatisme de calcul de la paie à l'envers en focntion de la saisie par rubrique
// d PT24
//          if RetEnvers then SauvegardeBul(Salarie, Etab, DateD, DateF, DateS, Tob_Rub, Numerateur, Denominateur, FALSE);
// DEB PT32
          if RetEnvers and ValidationOK then
          begin
            okok := SauvegardeBul(Salarie, Etab, DateD, DateF, DateS, Tob_Rub, Numerateur, Denominateur, FALSE, TOB_IJSS, TOB_Maintien);
            // DEB PT33
            if okok then
            begin
              if Assigned(Tob_PSD) then
              begin
                if VH_Paie.PGMODULEPRESENCE then
                  MiseAjourPresence(DateD, DateF, Codesalarie, 'MAJ'); // PT45

                TOB_PSD.UpdateDB(false, false);
                FreeAndNil(Tob_PSD);
              end;
              if Assigned(Tob_PSP) then
              begin
                TOB_PSP.UpdateDB(false, false);
                FreeAndNil(Tob_PSP);
              end;
            end;
          end;
          // FIN PT33
          if not OkOk then
{$IFNDEF EAGLSERVER}
            TraceErr.Items.Add('Erreur lors de la validation du bulletin pour le salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
              Tob_Salarie.GetValue('PSA_PRENOM'));
{$ELSE}
            TraceE.Add('Erreur lors de la validation du bulletin pour le salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
              Tob_Salarie.GetValue('PSA_PRENOM'));
{$ENDIF}
// FIN PT32
          // f PT24
                    // FIN PT19
          LibereTobCP(T); //PT11
          LibereTobCalendrier; { PT37 }
        end;
      end
      else
{$IFNDEF EAGLSERVER}
        TraceErr.Items.Add('Il n''y a pas de ventilation analytique pour le salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
          Tob_Salarie.GetValue('PSA_PRENOM'));
{$ELSE}

        TraceE.Add('Il n''y a pas de ventilation analytique pour le salarié ' + Salarie + ' ' + Tob_Salarie.GetValue('PSA_LIBELLE') + ' ' +
          Tob_Salarie.GetValue('PSA_PRENOM'));
{$ENDIF}
    end; // Fin Si pas de doublon dû à un bulletin complémentaire
    // debut PT49
    FreeAndNil(TobAcSaisieArret);
    // fin PT49
    QMul.NEXT;
  end; // Fin boucle sur la Query du mul
  // PT15    31/01/2003 PH V591 Barre de patience pdt la prépa
{$IFNDEF EAGLSERVER}
  FiniMoveProgressForm;
{$ENDIF}
  // PT16    13/05/2003 PH V_42 Optimisation chargement des tob de la paie non fait systématiquement
  VideLesTOBPaie(FALSE); // Désallocation des TOB necessaires aux calculs de la paie
  // FIN PT16
  if TOB_Salarie <> nil then
  begin
    TOB_Salarie.free;
    TOB_Salarie := nil
  end;
  if TOB_DUSALARIE <> nil then
  begin
    TOB_DUSALARIE.Free;
    TOB_DUSALARIE := nil;
  end;
  if TOB_HistoBasesCot <> nil then
  begin
    TOB_HistoBasesCot.Free;
    TOB_HistoBasesCot := nil;
  end;
  FreeAndNil(Tob_Rub);
  if TOBAna <> nil then
  begin
    TOBAna.Free;
    TOBAna := nil;
  end;
  // d PT21
  if Assigned(Tob_Abs) then
    FreeAndNil(Tob_Abs);
  // f PT21

  VideLaTobExer;
  GetLocalTime(SystemTime0);
{$IFNDEF EAGLSERVER}
  Trace.Items.Add('Fin du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  i := Trace.Items.Count - 5; //PT13 5 au lieu de 4
  Trace.Items.Add('Nombre de Salariés traités : ' + IntToStr(i));
  // PT9   : 03/03/2002 V582 PH Gestion historique des évènements
  CreeJnalEvt('001', '001', 'OK', Trace, TraceErr);
  PGIInfo('La préparation automatique est terminée', Info);
  if TraceErr.Items.Count > 0 then
  begin
    Tbsht := TTabSheet(PGGetControl('TBSHTERROR', Ecran));
    if Tbsht <> nil then Pan.ActivePage := Tbsht;
    //Génération d'un fichier de log
    {DEB2 PT1}
    if MessageDlg('Voulez-vous générez le fichier Prepa.log sous le répertoire ' + VH_Paie.PGCheminEagl + '\.', mtConfirmation, [mbYes, mbNo], 0) = mrYes then //PT2
    begin
      if VH_Paie.PGCheminEagl <> '' then FileN := VH_Paie.PGCheminEagl + '\Prepa.log' //PT2
{$IFDEF EAGLCLIENT}
      else FileN := 'C:\Prepa.log'; //PT6
{$ELSE}
      else FileN := V_PGI.StdPath + '\Prepa.log'; //PT6
{$ENDIF}
      if SupprimeFichier(FileN) = False then exit;
      AssignFile(F, FileN);
{$I-}ReWrite(F);
{$I+}
      if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
        Exit;
      end;
      writeln(F, 'Préparation automatique : Gestion des messages d''erreur.');
      for i := 0 to TraceErr.Items.Count - 1 do
      begin
        St := TraceErr.Items.Strings[i];
        writeln(F, St);
      end;
      CloseFile(F);
      PGIInfo('La génération est terminée', Info);

      if VH_Paie.PGCheminEagl <> '' then FileN := '"'+VH_Paie.PGCheminEagl + '\Prepa.log"'; // PT43
      ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(FileN), nil, SW_RESTORE); //PT6
    end;
    {FIN2 PT1}
  end;
{$ELSE}
  Trace.Add('Fin du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  i := Trace.Count - 5; //PT13 5 au lieu de 4
  Trace.Add('Nombre de Salariés traités : ' + IntToStr(i));
  // PT9   : 03/03/2002 V582 PH Gestion historique des évènements
  CreeJnalEvt('001', '001', 'OK', nil, nil, Trace, TraceE);
  Trace.Free;
  TraceE.Free;
{$ENDIF}
  // d PT25
  if Assigned(Tob_IJSS) then
    FreeAndNil(Tob_IJSS);
  if Assigned(Tob_Maintien) then
    FreeAndNil(Tob_Maintien);
  // f PT25
//  LeMode := '';
  IntegreAuto := FALSE; // PT38

  //DEB PT40
{$IFNDEF EAGLSERVER}
  Q  := OpenSql('SELECT COUNT(*) AS COMPTEUR FROM PGSYNELTNAT WHERE (PEY_TYPUTI = "1" OR PEY_TYPUTI = "3") AND PEY_NODOSSIER = "' + ClePGSynEltNAt + '"',True) ;
  if not Q.eof then
    NbEltManquant := Q.FindField('COMPTEUR').AsInteger;
  Ferme(Q);

  if NbEltManquant <> 0 then
  begin
    if PgiAsk(TraduireMemoire('Des éléments ont été détectés manquants. Voulez-vous les saisir maintenant ?'), TraduireMemoire('Eléments nationaux manquants')) = mrYes then
      AglLanceFiche('PAY', 'SAISIESYNELTNAT', '', '', ClePGSynEltNAt);
    ExecuteSQL('DELETE FROM PGSYNELTNAT WHERE (PEY_TYPUTI = "1" OR PEY_TYPUTI = "3") AND PEY_NODOSSIER = "' + ClePGSynEltNAt + '"');
  end;

{$ENDIF}
  //FIN PT40
end;

{$IFNDEF EAGLSERVER}

procedure TOF_PG_PrepAuto.ImprimeClick(Sender: TObject);
var
  MPages: tpagecontrol;
begin
{$IFNDEF EAGLCLIENT}
  MPages := TPageControl(getcontrol('PANELPREP'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
{$ENDIF}
end;

procedure TOF_PG_PrepAuto.OnArgument(Arguments: string);
var
  F: TFVierge;
  st: string;
  TheBulCompl: string;
  BImprime: TToolbarButton97;
begin
  inherited;
  GrilleBull := nil;
  TypeTraitement := 'PREPA'; // pour indiquer le type de traitement saisie ou prepa auto au moteur de calcul
  st := Trim(Arguments);
  LblDD := THLabel(GetControl('DATEDEBUT'));
  LblDF := THLabel(GetControl('DATEFIN'));
  if LblDD <> nil then LblDD.Caption := ReadTokenSt(st);
  if LblDF <> nil then LblDF.Caption := ReadTokenSt(st);
  // PT20    08/12/2003 PH V620 Préparation automatique de bulletins complémentaires
  TheBulCompl := ReadTokenSt(st);
  ProfilSpec := ReadTokenSt(st);
  if TheBulCompl = 'X' then
  begin
    BullCompl := 'X';
    Ecran.Caption := 'Génération automatique de bulletins complémentaires';
    UpdateCaption(Ecran);
  end
  else
  begin
    // PT3 23/10/2001 V562  PH Gestion cas particulier du bulletin complémentaire et Dates edition
    BullCompl := '-';
  end;
  // FIN PT20
  BullContrat := readtokenst (st);   //PT39
  // d PT23

{$IFNDEF EAGLSERVER}
  //PT46
  if st = 'MULTI' then
  begin
    SetControlVisible('BPROGMULTIDOSSIER', True);
    (GetControl('BPROGMULTIDOSSIER') as TToolbarButton97).OnClick := OnBPROGMULTIDOSSIERClick;
  end;
{$ENDIF}


{$IFNDEF CCS3}
  if not VH_Paie.PgMaintien then
  begin
    SetControlEnabled('MAINTIEN', False);
    SetControlVisible('MAINTIEN', False);
    SetControlChecked('MAINTIEN', False);
  end;
{$ELSE}
  SetControlEnabled('MAINTIEN', False);
  SetControlVisible('MAINTIEN', False);
  SetControlChecked('MAINTIEN', False);
{$ENDIF}
  // f PT23

  BtnLance := TToolbarButton97(GetControl('BTNLANCE'));
  if BtnLance <> nil then BtnLance.OnClick := LancePrepAuto;
  if V_Pgi.SAV then SetControlChecked('BPROG', true) // FQ 13181
  else SetControlChecked('BPROG', False);
{$IFDEF DLL}
{$IFDEF EAGLCLIENT}
  SetControlVisible ('BPROG', True);
{$ENDIF}
{$ENDIF}

  BTNPROCESS := TToolbarButton97(GetControl('BPROG'));
  if BTNPROCESS <> nil then BTNPROCESS.OnClick := ParamPrepAuto;


  BImprime := ttoolbarbutton97(getcontrol('BIMPRIMER'));
  if Bimprime <> nil then
    Bimprime.Onclick := ImprimeClick;
  if not (Ecran is TFVierge) then exit;
  F := TFVierge(Ecran);
  if F <> nil then
  begin
{$IFDEF EAGLCLIENT}
    QMUL := THQuery(F.FMULQ).TQ;
{$ELSE}
    QMUL := F.FMULQ;
{$ENDIF}
  end;
end;
{$ENDIF}


{ Fonction qui recherche et alimente la structure chpEntente contenant toutes les
variables nécessaires au calcul de la paie.
Exemple : Bases Forcees ....
          Trentieme
          Date de paiement des Salaires
          Mode de règlement
}

procedure RecupCasPaie(var ActionBulletin: TActionBulletin);
var
  TPR, T2, T_Etab: TOB;
  st: string;
  Q: TQuery;
  LD, LF: TDateTime;
begin
  ActionBulletin := taModification; //PT39
  Numerateur := CalculTrentieme(DateD, DateF, TRUE); // PT22
  Denominateur := 30;
  // PT5 18/12/2001 V571  PH controle dates edition par rapport à l'etablissement
  T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValue('PSA_ETABLISSEMENT')], True);
  LD := DateD;
  LF := DateF;
  if T_ETab <> nil then
  begin
    if (T_ETAB.GetValue('ETB_JEDTDU') <> 0) and (T_ETAB.GetValue('ETB_JEDTAU') <> 0) then
    begin // les dates sont calculées et accessibles
      CalculeDateEdtPaie(LD, LF, T_Etab);
    end;
  end;
  // FIN PT5
  // PT7 26/03/2002 V571  PH Raz structure entete ChpEntete
  // PT14    19/12/2002 PH V591 Affectation de la date de paiement = Date de fin de paie si non renseignée
  RazChptEntete(DateF); // RAZ structure entete
  with ChpEntete do
  begin
    Reglt := RendModeRegle(TOB_Salarie);
    DTrent := Denominateur;
    NTrent := Numerateur;
    Ouvres := 0;
    HOuvres := 0;
    Ouvrables := 0;
    HOuvrables := 0;
    if T_Etab <> nil then
    begin
      {DEB PT12 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables
      Ouvres:=NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,TRUE);
      Ouvrables:=NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,FALSE);}
      Tob_CalendrierSalarie := ChargeCalendrierSalarie(TOB_Salarie.getvalue('PSA_ETABLISSEMENT'),
        TOB_Salarie.getvalue('PSA_SALARIE'), TOB_Salarie.getvalue('PSA_CALENDRIER'),
        TOB_Salarie.getvalue('PSA_STANDCALEND'), T_Etab.getvalue('ETB_STANDCALEND'));
      CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, Tob_CalendrierSalarie, DateD, DateF, TRUE, Ouvres, HOuvres);
      CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, Tob_CalendrierSalarie, DateD, DateF, FALSE, Ouvrables, HOuvrables);
    end; //FIN PT12
    HeuresTrav := 0;
    HorMod := FALSE;
    BasesMod := FALSE;
    TranchesMod := FALSE;
    TrentMod := FALSE;
    DatePai := CalculDatePaie(TOB_Salarie, DateF);
    DateVal := NOW;
    // PT5 18/12/2001 V571  PH controle dates edition par rapport à l'etablissement
    Edtdu := LD;
    Edtau := LF;
    // FIN PT5
  end;
  MemoriseTrentieme(Numerateur / Denominateur);
  TPR := TOB.Create('Le Bulletin de la Paie En Cours', TOB_Paie, -1);
  st := 'SELECT * FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + TOB_Salarie.GetValue('PSA_ETABLISSEMENT') + '"' + ' AND PPU_SALARIE="' +
    TOB_Salarie.GetValue('PSA_SALARIE') + '" AND PPU_DATEDEBUT="' +
    USDateTime(DateD) + '" AND PPU_DATEFIN="' + USDateTime(DateF) + '"';
  Q := OpenSql(st, TRUE);
  TPR.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
  T2 := TPR.FindFirst(['PPU_ETABLISSEMENT', 'PPU_SALARIE'], [TOB_Salarie.GetValue('PSA_ETABLISSEMENT'), TOB_Salarie.GetValue('PSA_SALARIE')], TRUE);
  if (T2 <> nil) then
  begin
    with ChpEntete do
    begin
      if T2.GetValue('PPU_BASESMOD') = 'X' then BasesMod := TRUE;
      if T2.GetValue('PPU_TRANCHESMOD') = 'X' then TranchesMod := TRUE;
      if T2.GetValue('PPU_HORAIREMOD') = 'X' then HorMod := TRUE;
      // PT7 26/03/2002 V571  PH Prise en compte des règlements
      if T2.GetValue('PPU_REGLTMOD') = 'X' then RegltMod := TRUE;

      if T2.GetValue('PPU_TRENTIEMEMOD') = 'X' then
      begin
        TrentMod := TRUE;
        DTrent := T2.GetValue('PPU_DENOMINTRENT');
        Ntrent := T2.GetValue('PPU_NUMERATTRENT');
      end;
      if HorMod = FALSE then
      begin
        HeuresTrav := T2.GetValue('PPU_HEURESREELLES');
        Ouvres := T2.GetValue('PPU_JOURSOUVRES');
        Ouvrables := T2.GetValue('PPU_JOURSOUVRABLE');
      end;
      // PT7 26/03/2002 V571  PH Prise en compte des règlements
      if RegltMod then
        ChpEntete.Reglt := T2.getvalue('PPU_PGMODEREGLE');

      DatePai := T2.GetValue('PPU_PAYELE');
      DateVal := T2.GetValue('PPU_VALIDELE');
      // PT3 23/10/2001 V562  PH Gestion cas particulier du bulletin complémentaire et Dates edition
      Edtdu := T2.GetValue('PPU_EDTDEBUT');
      Edtau := T2.GetValue('PPU_EDTFIN');
    end;
  end
  else
  begin
    ActionBulletin := taCreation; //PT39
  end;
  Ferme(Q);
  TPR.Free;
end;
{$IFNDEF EAGLSERVER}

procedure TOF_PG_PrepAuto.ParamPrepAuto(Sender: TObject);
var TobParam, T1: TOB;
  LeSQL, LeMode : string;
  DD, DF: TDateTime;
begin
  LeSQL := QMUL.SQL[0];
  DD := StrToDate(LblDD.Caption);
  DF := StrToDate(LblDF.Caption);
  TobParam := TOB.create('Ma Tob de Param', nil, -1);
  T1 := TOB.Create('XXX', TobParam, -1);
  T1.AddChampSupValeur('DD', LblDD.Caption);
  T1.AddChampSupValeur('DF', LblDF.Caption);
{$IFDEF EAGLCLIENT}
  LeMode := 'CWAS';
{$ELSE}
  LeMode := '2TIERS';
{$ENDIF}
  T1.AddChampSupValeur('LEMODE', LeMode);
  T1.AddChampSupValeur('LESQL', LeSQL);
//  LancePRocess ....
  AGLFicheJob(0, taCreat, 'cgiPaieS5', 'PREPAUTO', TobParam);
  FreeAndNil(TOBParam);
end;
{$ENDIF}


{$IFNDEF EAGLSERVER}

procedure TOF_PG_PrepAuto.OnBPROGMULTIDOSSIERClick(Sender: TObject);
var TobParam, T1: TOB;
  LeSQL : string;
  DD, DF: TDateTime;
begin
  LeSQL := QMUL.SQL[0];
  DD := StrToDate(LblDD.Caption);
  DF := StrToDate(LblDF.Caption);
  TobParam := TOB.create('Ma Tob de Param', nil, -1);
  T1 := TOB.Create('XXX', TobParam, -1);
  T1.AddChampSupValeur('DD', LblDD.Caption);
  T1.AddChampSupValeur('DF', LblDF.Caption);
  T1.AddChampSupValeur('LEMODE', {$IFDEF EAGLCLIENT}'CWAS'{$ELSE}'2TIERS'{$ENDIF});
  T1.AddChampSupValeur('LESQL', LeSQL);
//  LancePRocess ....
  FicheJobMultiDossier(0, taCreat, 'cgiPaieS5', 'PREPAUTO', TobParam, '', '', '', '', True, 2);
  FreeAndNil(TOBParam);
end;
{$ENDIF}

initialization
{$IFNDEF EAGLSERVER}
  registerclasses([TOF_PG_PrepAuto]);
{$ENDIF}
end.

