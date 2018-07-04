{***********UNITE*************************************************
Auteur  ...... : MF
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : TOF de la fiche vierge CALCULACF
Suite ........ : Calcul des jours de fractionnement
Mots clefs ... : FRACTIONNEMENT
*****************************************************************}
{
  PT1 04/09/2008  MF  correction liée à l'utilisation de listes différentes
                      + StrToFloat remplacé par Valeur afin de récupérer une
                      valeur réelle correcte
}
unit PGCalculACF;
interface
uses
    Classes,
    Controls,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
    EntPaie,
    Dialogs,
    HCtrls,
    HEnt1,
    HQry,
    HMsgBox,
    HTB97,
    P5Def,
    ShellAPI,
    StdCtrls,
    sysutils,
    UTOB,
    UTOF,
    Vierge,
    Windows;

Type TParamCalcul = record
  ACQUISDROITFRAC     : string;        // Types d'acquis donnant droit au fractionnement
  CONGEPRINCIPALA     : string;        // Date fin congé principal
  CONGEPRINCIPALDE    : string;        // Date début congé principal
  DATECALCULFRAC      : string;        // date du calcul des jours de congés pour fractionnement
  NBPRISCONGEPRINC    : double;        // nombre de jours du congé principal
  TBJoursAcquis       : array of Real; // table du nbre d'ACF possibles (en fct du nbre de jours restant)
  TBJoursRestant      : array of Real; // table du du nbre de jours restant.
  BornePris           : double;        // nbre de pris mini nécessaires pour lancer le calcul
end;

type
  TOF_PGCALCULACF= class(TOF)
  private
    datedebut, datefin                       : TDatetime;
    NomListe                                 : string;  
    procedure  LanceCalculFractionnement(Sender: TObject);
    procedure  CreerLesTob(var TOB_ACFN1,TOB_Acquis,TOB_PrisPayes : TOB);
    procedure  VerifACFExisteDeja(var T_MvtAcquis : TOB ; const TOB_ACFN1 : TOB);
    procedure  ViderLesTob(var TOB_ACFN1,TOB_Acquis,TOB_PrisPayes : TOB);
    procedure  CreerTobSalPopul(const T_MvtAcquis : TOB; var TOB_SalPopul : TOB);
    procedure  CreerTobSal(const T_MvtAcquis : TOB; var TOB_Sal : TOB);
    procedure  RecupParamCalcul(var Paramcalcul : TParamCalcul; const PGParametreAssoc : string);
    procedure  CalculJoursFractionnement(const ParamCalcul : TParamCalcul;const TOB_Sal_fille,TOB_PrisPayes,TOB_Acquis : TOB; var T_ACF : TOB; var Trace : TListBox);

  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;

implementation
uses
  PGCalendrier,
  PgCommun,
  PgOutils2;

var
  QMul                                : TQUERY; // Query recupérée du mul
  Trace                               : TListBox;
  IndTrace                            : boolean;
  NbreSalaries                        : integer;


procedure TOF_PGCALCULACF.OnArgument(Arguments: string);
var
  BtnLance                 : TToolBarButton97;
  F                        : TFVierge;
  st                       : string;

begin
  inherited;
  St := Trim (Arguments);
  NomListe := ReadTokenSt(st);              // Récup nom de la liste utilisée
  datedebut := StrToDate(ReadTokenSt(st));  // Recup date de debut
  datefin := StrToDate (ReadTokenSt(st));   // Recup date fin
//  annee := copy(DateToStr(datedebut),7,4);
  BtnLance := TToolbarButton97(GetControl('BTNLANCE'));
  if BtnLance <> nil then
    BtnLance.OnClick := LanceCalculFractionnement;

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
procedure TOF_PGCALCULACF.OnClose;
begin
   inherited;
   Trace.Free;
end;
procedure  TOF_PGCALCULACF.LanceCalculFractionnement(Sender: TObject);
var
  TOB_ACFN1,TOB_Acquis,TOB_PrisPayes          : TOB;
  T_MvtAcquis, T_ACF                          : TOB;
  TOB_SalPopul, TOB_SalPopul_fille            : TOB;
  TOB_Sal,TOB_Sal_fille                       : TOB;
  No_Ordre,i                                  : integer;
  SystemTime0                                 : TSystemTime;
  population, PGParametreAssoc,LibPopul,FileN : string;
  Q                                           : TQuery;
  Paramcalcul                                 : TParamcalcul;
  JnlTrace, JnlTraceE                         : TStringList; //  Pour trace dans jnal des événements
  F                                           : TextFile;
  St                                          : string;

//                                                    DTClot := TobEtablissement.getvalue('ETB_DATECLOTURECPN');
begin
  Trace := TListBox(GetControl('LSTBXTRACE'));
  NbreSalaries := 0;

  JnlTrace := TStringList.Create;
  JnlTraceE := TStringList.Create;

  if (PGIAsk('Voulez- vous afficher le détail du calcul par salarié ?',Ecran.Caption)=mrYes)  then
    IndTrace := true
  else
    IndTrace := false;

  Trace.Clear;
  GetLocalTime(SystemTime0);
  Trace.Items.Add('Début du traitement de calcul des jours de fractionnement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  Trace.Items.Add('');
  Trace.Items.Add('Jours de fractionnment acquis le : '+DateToStr(DebutdeMois(Date)));
  Trace.Items.Add('');
  JnlTrace.Add ('Début du traitement de calcul des jours de fractionnement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  JnlTrace.Add ('Jours de fractionnment acquis le : '+DateToStr(DebutdeMois(Date)));

  // création TOBs utiles au calcul
  CreerLesTob(TOB_ACFN1,TOB_Acquis,TOB_PrisPayes);

  // Création de la TOB des Acquis de type ACF à calculer (à partir de la liste)
  T_MvtAcquis := TOB.Create('CONGES PAYE SALARIE mère', nil, -1);

  QMul.First;
  while not QMul.EOF do
  begin
    T_ACF := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
    InitialiseTobAbsenceSalarie(T_ACF);
// d PT1
    if ( NomListe = 'PGAFFECTSALPOPUL') then
      No_Ordre := IncrementeSeqNoOrdre('CPA',QMul.findfield('PNA_SALARIE').AsString)
    else
      No_Ordre := IncrementeSeqNoOrdre('CPA',QMul.findfield('PSA_SALARIE').AsString);
// f PT1
    if ( NomListe = 'PGAFFECTSALPOPUL') then
      T_ACF.PutValue('PCN_SALARIE', QMul.findfield('PNA_SALARIE').AsString)
    else
      T_ACF.PutValue('PCN_SALARIE', QMul.findfield('PSA_SALARIE').AsString);

    T_ACF.PutValue('PCN_ORDRE', No_Ordre);
    T_ACF.PutValue('PCN_CODERGRPT', No_Ordre);
    T_ACF.PutValue('PCN_DATEDEBUT', PlusDate(DebutdeMois(Date),-1,'M'));
    T_ACF.PutValue('PCN_DATEFIN', FindeMois(PlusDate(DebutdeMois(Date),-1,'M')));
    T_ACF.PutValue('PCN_TYPECONGE', 'ACF');
    T_ACF.PutValue('PCN_SENSABS', '+');
    T_ACF.PutValue('PCN_LIBELLE', 'Acquis de fractionnement au ' + DateToStr(DebutdeMois(Date)));
    T_ACF.PutValue('PCN_DATEMODIF', Date);
    T_ACF.PutValue('PCN_DATEVALIDITE', DebutdeMois(Date));
    T_ACF.PutValue('PCN_PERIODECP', 1);
    T_ACF.PutValue('PCN_JOURS', 0);
    T_ACF.PutValue('PCN_BASE', 0);
    T_ACF.PutValue('PCN_NBREMOIS', 0);
    T_ACF.PutValue('PCN_CODETAPE', '...');

    QMul.Next;
  end;
  NbreSalaries := T_MvtAcquis.detail.Count;

  // Pour chaque éléments (salarié) de la TOB des acquis ACF à calculer vérification
  // Mvt ACF pas déjà existant
  if (TOB_ACFN1 <> nil) then
    VerifACFExisteDeja(T_MvtAcquis, TOB_ACFN1);

  if (NomListe = 'PGAFFECTSALPOPUL') then
    // créer la tob des salariés par population avec ses paramètres CP
    CreerTobSalPopul(T_MvtAcquis,TOB_SalPopul)
  else
  if ( NomListe = 'PGSALARIES') then
    // créer la tob des salariés par population avec ses paramètres CP
    CreerTobSal(T_MvtAcquis,TOB_Sal);

  population := '';
  PGParametreAssoc := '';
  LibPopul := '';
  
  if ( NomListe = 'PGAFFECTSALPOPUL') then
  begin
    // calcul des jours : recherche des paramètres de calcul + calcul
    for i:= 0 to TOB_SalPopul.detail.count-1 do
    begin
      TOB_SalPopul_fille := TOB_SalPopul.detail[i];
      if TOB_SalPopul_fille.getvalue('POPULATION') <> population  then
      // changement de population
      begin
        // traitement d'une population
        population := TOB_SalPopul_fille.getvalue('POPULATION');

        //récupération du code paramètre associé
        Q := OpenSql('SELECT PGO_PGPARAMETRE, PPC_LIBELLE FROM PGPARAMETRESASSOC '+
                     'LEFT JOIN ORDREPOPULATION ON PPC_POPULATION = PGO_CODEASSOCIE '+
                     'WHERE '+
                     'PGO_CODEASSOCIE="'+population+'" ', TRUE);
        if not Q.eof then
        begin
          PGParametreAssoc := Q.findField('PGO_PGPARAMETRE').AsString;
          LibPopul := Q.findField('PPC_LIBELLE').AsString;
        end;
        Ferme (Q);

        //récupération des paramètres de calcul
        RecupParamCalcul(Paramcalcul,PGParametreAssoc);
        if (StrToDate(Paramcalcul.DATECALCULFRAC) > datedebut) then
        begin
          Trace.Items.Add('Population '+LibPopul);
          Trace.Items.Add('           La date d''application du fractionnement est normalement prévue au '+
                          Paramcalcul.DATECALCULFRAC);
          Trace.Items.Add(' ');

          JnlTrace.Add('Population '+LibPopul);
          JnlTrace.Add('           La date d''application du fractionnement est normalement prévue au '+
                          Paramcalcul.DATECALCULFRAC);
        end;
      end;
      T_ACF := T_MvtAcquis.findfirst(['PCN_SALARIE'], [TOB_SalPopul_fille.GetValue('SALARIE')], False);

      // Calcul des Jours de fractionnement pour le salarié en cours
      CalculJoursFractionnement(Paramcalcul, TOB_SalPopul_fille,TOB_PrisPayes,TOB_Acquis,T_ACF, Trace);
      if (T_ACF.GetValue('PCN_JOURS') = 0) then
      begin
        T_ACF.Free;
        NbreSalaries := NbreSalaries - 1;
      end
      else
      begin
        T_ACF.PutValue('PCN_GUID',AglGetGuid());
        T_ACF.PutValue('PCN_ETABLISSEMENT', TOB_SalPopul_fille.getvalue('ETABLISSEMENT'));
        T_ACF.Putvalue('PCN_TRAVAILN1', TOB_SalPopul_fille.getvalue('TRAVAILN1'));
        T_ACF.Putvalue('PCN_TRAVAILN2', TOB_SalPopul_fille.getvalue('TRAVAILN2'));
        T_ACF.Putvalue('PCN_TRAVAILN3', TOB_SalPopul_fille.getvalue('TRAVAILN3'));
        T_ACF.Putvalue('PCN_TRAVAILN4', TOB_SalPopul_fille.getvalue('TRAVAILN4'));
        T_ACF.Putvalue('PCN_CODESTAT', TOB_SalPopul_fille.getvalue('CODESTAT'));
        T_ACF.Putvalue('PCN_CONFIDENTIEL', TOB_SalPopul_fille.getvalue('CONFIDENTIEL'));

      end;
    end;
  end
  else
  // (NomListe = 'PGSALARIES')
  begin
    Q := OpenSql('SELECT PGP_PGPARAMETRE FROM PGPARAMETRES WHERE '+
                 'PGP_MODULECHAMP="ACF"', TRUE);
    if not Q.eof then
    begin
      PGParametreAssoc := Q.findField('PGP_PGPARAMETRE').AsString;
    end;
    Ferme (Q);

    //récupération des paramètres de calcul
    RecupParamCalcul(Paramcalcul,PGParametreAssoc);
    if (StrToDate(Paramcalcul.DATECALCULFRAC) > datedebut) then
    begin
      Trace.Items.Add('           La date d''application du fractionnement est normalement prévue au '+
                      Paramcalcul.DATECALCULFRAC);
      Trace.Items.Add(' ');
    end;

    for i:= 0 to TOB_Sal.detail.count-1 do
    begin
      TOB_Sal_fille := TOB_Sal.detail[i];

      T_ACF := T_MvtAcquis.findfirst(['PCN_SALARIE'], [TOB_Sal_Fille.GetValue('SALARIE')], False);
     // Calcul des Jours de fractionnement pour le salarié en cours
      CalculJoursFractionnement(Paramcalcul, TOB_Sal_Fille,TOB_PrisPayes,TOB_Acquis,T_ACF, Trace);
      if (T_ACF.GetValue('PCN_JOURS') = 0) then
      begin
        T_ACF.Free;
        NbreSalaries := NbreSalaries - 1;
      end
      else
      begin
        T_ACF.PutValue('PCN_GUID',AglGetGuid());
        T_ACF.PutValue('PCN_ETABLISSEMENT', TOB_Sal_Fille.getvalue('ETABLISSEMENT'));
        T_ACF.Putvalue('PCN_TRAVAILN1', TOB_Sal_Fille.getvalue('TRAVAILN1'));
        T_ACF.Putvalue('PCN_TRAVAILN2', TOB_Sal_Fille.getvalue('TRAVAILN2'));
        T_ACF.Putvalue('PCN_TRAVAILN3', TOB_Sal_Fille.getvalue('TRAVAILN3'));
        T_ACF.Putvalue('PCN_TRAVAILN4', TOB_Sal_Fille.getvalue('TRAVAILN4'));
        T_ACF.Putvalue('PCN_CODESTAT', TOB_Sal_Fille.getvalue('CODESTAT'));
        T_ACF.Putvalue('PCN_CONFIDENTIEL', TOB_Sal_Fille.getvalue('CONFIDENTIEL'));

      end;
    end;
  end;

  // mise à jour de la table ABSENCESALARIE
  BeginTrans;
  Try
  begin
  T_MvtAcquis.InsertDB(nil, FALSE);
  CommitTrans;
  end;
  except
    RollBack;
  end;

  FreeAndNil(TOB_Sal);
  FreeAndNil(TOB_SalPopul);
  FreeAndNil(T_MvtAcquis);
  ViderLesTob (TOB_ACFN1,TOB_Acquis,TOB_PrisPayes);

  GetLocalTime(SystemTime0);

  Trace.Items.Add(' ');
  if (NbreSalaries <= 1) then
  begin
    Trace.Items.Add(IntToSTr(NbreSalaries) + ' salarié traité');
    JnlTrace.Add (IntToSTr(NbreSalaries) + ' salarié traité');
  end
  else
  begin
    Trace.Items.Add(IntToSTr(NbreSalaries) + ' salariés traités');
    JnlTrace.Add (IntToSTr(NbreSalaries) + ' salariés traités');
  end;
  Trace.Items.Add(' ');

  JnlTrace.Add ('Fin du traitement de calcul des jours de fractionnement');
  CreeJnalEvt('002', '132', 'OK', nil, nil, JnlTrace, JnlTraceE);
  JnlTrace.Free;
  JnlTraceE.Free;

  Trace.Items.Add('Fin du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  PGIBox('Traitement teminé.','Calcul des jours de fractionnement');

  if MessageDlg('Voulez-vous générez le fichier Fract.log sous le répertoire ' + VH_Paie.PGCheminEagl + '\.', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    if VH_Paie.PGCheminEagl <> '' then FileN := VH_Paie.PGCheminEagl + '\Fract.log'
{$IFDEF EAGLCLIENT}
    else
    begin
      FileN := 'C:\Fract.log';
{$ELSE}
    else
    begin
      FileN := V_PGI.StdPath + '\Fract.log';
{$ENDIF}
      PgiInfo('Attention, le répertoire de stockage n''existe pas.',
      FILEN);
    end;

    if SupprimeFichier(FileN) = False then exit;
      AssignFile(F, FileN);
{$I-}
    ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
      Exit;
    end;
    for i := 0 to Trace.Items.Count - 1 do
    begin
      St := Trace.Items.Strings[i];
      writeln(F, St);
    end;
    CloseFile(F);

    if VH_Paie.PGCheminEagl <> '' then
      FileN := '"'+VH_Paie.PGCheminEagl + '\Fract.log"';
    ShellExecute(0, PCHAR('open'), PChar('WordPad'), PChar(FileN), nil, SW_RESTORE);
  end;


end;
procedure  TOF_PGCALCULACF.CreerLesTob(var TOB_ACFN1,TOB_Acquis,TOB_PrisPayes : TOB);
var
  Q                                 : TQuery;
begin
  // créer la tob des ACF sur N-1 pour tous les salaries de type congés payés
  Q := OpenSql('SELECT PCN_SALARIE, PCN_TYPECONGE, PCN_JOURS, PCN_HEURES '+
               'FROM ABSENCESALARIE WHERE '+
               'PCN_TYPEMVT="CPA" AND '+
               'PCN_TYPECONGE="ACF" AND '+
               'PCN_PERIODECP=1', TRUE);
  if not Q.eof then
  begin
    TOB_ACFN1 := TOB.Create('Les ACF', nil, -1);
    TOB_ACFN1.LoadDetailDB('ACFALARIE', '', '', Q, FALSE);
  end;
  Ferme (Q);

  // créer la tob des acquis N-1 (ts les acquis) pour tous les salaries (PCN_PERIODECP = 1)
  Q := OpenSql('SELECT PCN_SALARIE, PCN_TYPECONGE, PCN_JOURS, PCN_HEURES '+
               'FROM ABSENCESALARIE WHERE '+
               'PCN_TYPEMVT="CPA" AND '+
               '(PCN_TYPECONGE="ACQ" '+
               'OR PCN_TYPECONGE="REP" '+
               'OR PCN_TYPECONGE="ACS" '+
               'OR PCN_TYPECONGE="ACA" '+
               'OR PCN_TYPECONGE="REL" '+
               'OR PCN_TYPECONGE="ARR" '+
               'OR PCN_TYPECONGE="AJU" '+
               'OR PCN_TYPECONGE="AJP") '+
               'AND PCN_SENSABS="+" '+
               'AND PCN_PERIODECP=1', TRUE);
  if not Q.eof then
  begin
    TOB_Acquis := TOB.Create('Les Acquis', nil, -1);
    TOB_Acquis.LoadDetailDB('ACQUISSALARIE', '', '', Q, FALSE);
  end;
  Ferme (Q);

  // créer la tob des pris payés sur N-1  (PCN_PERIODECP = 1) avec dates de période
  // on ne tient pas compte des pris parents (PCN_CODERGRPT <> -1)
  Q := OpenSql('SELECT PCN_SALARIE, PCN_TYPECONGE, PCN_DATEDEBUTABS,'+
               'PCN_DATEFINABS,PCN_JOURS, PCN_HEURES '+
               'FROM ABSENCESALARIE WHERE '+
               'PCN_TYPEMVT="CPA" AND '+
               'PCN_TYPECONGE="PRI" '+
               'AND PCN_PERIODECP=1 '+
               'AND PCN_CODERGRPT<>-1 '+
               'AND PCN_CODETAPE="P"', TRUE);
  if not Q.eof then
  begin
    TOB_PrisPayes := TOB.Create('Les pris payés', nil, -1);
    TOB_PrisPayes.LoadDetailDB('PRISSALARIE', '', '', Q, FALSE);
  end;
  Ferme (Q);
end;

procedure  TOF_PGCALCULACF.VerifACFExisteDeja(var T_MvtAcquis : TOB ; const TOB_ACFN1 : TOB);
var
  T_ACF       : TOB ;
begin
  T_ACF := T_MvtAcquis.findfirst([''], [''], False);
  while T_ACF <> nil do
  begin
    if TOB_ACFN1.FindFirst(['PCN_SALARIE', 'PCN_TYPECONGE', 'PCN_PERIODCP'],
                           [T_ACF.GetValue('PCN_SALARIE'), 'ACF','1'],
                           FALSE) <> nil then
    begin
      Trace.Items.Add('Salarié :'+T_ACF.GetValue('PCN_SALARIE')+' : Les jours de fractionnement ont déjà été calculés.');
      T_ACF.free;
      NbreSalaries :=   NbreSalaries - 1;
    end;
    T_ACF := T_MvtAcquis.FindNext([''], [''], False);
  end;
end;
procedure TOF_PGCALCULACF.CreerTobSalPopul(const T_MvtAcquis : TOB; var TOB_SalPopul : TOB);
var
    i                           : integer;
    TOB_SalPopul_fille, TACF    : TOB;
    Q                          : TQuery;
begin
    TOB_SalPopul := TOB.Create('SALARIE par POULATION mere', nil, -1);

    for i:= 0 to T_MvtAcquis.detail.count-1 do
    begin
      TACF := T_MvtAcquis.detail[i];
      TOB_SalPopul_fille := TOB.Create('salpopul', TOB_SalPopul, -1);

      TOB_SalPopul_fille.AddChampSupValeur('SALARIE',TACF.getvalue('PCN_SALARIE'));

      Q := OpenSql('SELECT PNA_POPULATION,PSA_ETABLISSEMENT, PSA_NBREACQUISCP,'+
                   'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,'+
                   'PSA_CODESTAT,PSA_CONFIDENTIEL,ETB_DATECLOTURECPN '+
                   'FROM SALARIEPOPUL '+
                   'LEFT JOIN SALARIES ON PSA_SALARIE=PNA_SALARIE '+
                   'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT = PSA_ETABLISSEMENT '+
                   'WHERE '+
                   'PNA_TYPEPOP="ACF" AND '+
                   'PNA_SALARIE="'+TACF.getvalue('PCN_SALARIE')+'" ', TRUE);
      if not Q.eof then
      begin
        TOB_SalPopul_fille.AddChampSupValeur ('POPULATION', Q.findField('PNA_POPULATION').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('ETABLISSEMENT', Q.findField('PSA_ETABLISSEMENT').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('NBREACQUISCP', Q.findField('PSA_NBREACQUISCP').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('TRAVAILN1', Q.findField('PSA_TRAVAILN1').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('TRAVAILN2', Q.findField('PSA_TRAVAILN2').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('TRAVAILN3', Q.findField('PSA_TRAVAILN3').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('TRAVAILN4', Q.findField('PSA_TRAVAILN4').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('CODESTAT', Q.findField('PSA_CODESTAT').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('CONFIDENTIEL', Q.findField('PSA_CONFIDENTIEL').AsString);
        TOB_SalPopul_fille.AddChampSupValeur ('DATECLOTURECPN', Q.findField('ETB_DATECLOTURECPN').AsString);
      end;
      Ferme (Q);
    end;

    TOB_SalPopul.detail.sort('POPULATION;SALARIE');
end;
procedure TOF_PGCALCULACF.CreerTobSal(const T_MvtAcquis : TOB; var TOB_Sal : TOB);
var
    i                           : integer;
    TOB_Sal_fille, TACF         : TOB;
    Q                           : TQuery;
begin
    TOB_Sal := TOB.Create('SALARIE  mere', nil, -1);

    for i:= 0 to T_MvtAcquis.detail.count-1 do
    begin
      TACF := T_MvtAcquis.detail[i];
      TOB_Sal_fille := TOB.Create('sal', TOB_Sal, -1);

      TOB_Sal_fille.AddChampSupValeur('SALARIE',TACF.getvalue('PCN_SALARIE'));

      Q := OpenSql('SELECT PSA_ETABLISSEMENT, PSA_NBREACQUISCP,'+
                   'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,'+
                   'PSA_CODESTAT,PSA_CONFIDENTIEL,ETB_DATECLOTURECPN '+
                   'FROM SALARIES '+
                   'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT = PSA_ETABLISSEMENT '+
                   'WHERE '+
                   'PSA_SALARIE="'+TACF.getvalue('PCN_SALARIE')+'" ', TRUE);
      if not Q.eof then
      begin
        TOB_Sal_fille.AddChampSupValeur ('ETABLISSEMENT', Q.findField('PSA_ETABLISSEMENT').AsString);
        TOB_Sal_fille.AddChampSupValeur ('NBREACQUISCP', Q.findField('PSA_NBREACQUISCP').AsString);
        TOB_Sal_fille.AddChampSupValeur ('TRAVAILN1', Q.findField('PSA_TRAVAILN1').AsString);
        TOB_Sal_fille.AddChampSupValeur ('TRAVAILN2', Q.findField('PSA_TRAVAILN2').AsString);
        TOB_Sal_fille.AddChampSupValeur ('TRAVAILN3', Q.findField('PSA_TRAVAILN3').AsString);
        TOB_Sal_fille.AddChampSupValeur ('TRAVAILN4', Q.findField('PSA_TRAVAILN4').AsString);
        TOB_Sal_fille.AddChampSupValeur ('CODESTAT', Q.findField('PSA_CODESTAT').AsString);
        TOB_Sal_fille.AddChampSupValeur ('CONFIDENTIEL', Q.findField('PSA_CONFIDENTIEL').AsString);
        TOB_Sal_fille.AddChampSupValeur ('DATECLOTURECPN', Q.findField('ETB_DATECLOTURECPN').AsString);
      end;
      Ferme (Q);
    end;

    TOB_Sal.detail.sort('SALARIE');
end;

procedure  TOF_PGCALCULACF.ViderLesTob (var TOB_ACFN1,TOB_Acquis,TOB_PrisPayes : TOB);
begin
  FreeAndNil (TOB_ACFN1);
  FreeAndNil (TOB_Acquis);
  FreeAndNil (TOB_PrisPayes);
end;

procedure  TOF_PGCALCULACF.RecupParamCalcul(var Paramcalcul : TParamCalcul; const PGParametreAssoc : string);
var
  Q                                   : TQuery;
  ind                                 : integer;

begin
      if (PGParametreAssoc <> '') then
      begin
        Q := OpenSql('SELECT PGP_PGNOMCHAMP,PGP_PGVALCHAMP  FROM PGPARAMETRES WHERE '+
                     'PGP_PGPARAMETRE="'+PGParametreAssoc+'" '+
                     'AND PGP_PGNOMCHAMP not LIKE "GridJoursAcquis%" '+
                     'AND PGP_PGNOMCHAMP not LIKE "GridJoursRestant%"', TRUE);
        while not Q.eof do
        begin

          // Types d'acquis donnant droit au fractionnement
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'ACQUISDROITFRAC') then
            Paramcalcul.ACQUISDROITFRAC := Q.findField('PGP_PGVALCHAMP').AsString
          else
          // Date fin congé principal
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'CONGEPRINCIPALA') then
            Paramcalcul.CONGEPRINCIPALA  := Q.findField('PGP_PGVALCHAMP').AsString
          else
          // Date début congé principal
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'CONGEPRINCIPALDE') then
            Paramcalcul.CONGEPRINCIPALDE := Q.findField('PGP_PGVALCHAMP').AsString
          else
          // date du calcul des jours de congés pour fractionnement
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'DATECALCULFRAC') then
            Paramcalcul.DATECALCULFRAC := Q.findField('PGP_PGVALCHAMP').AsString
          else
          // nombre de jours du congé principal
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'NBPRISCONGEPRINC') then
// PT1      Paramcalcul.NBPRISCONGEPRINC := StrToFloat(Q.findField('PGP_PGVALCHAMP').AsString);
            Paramcalcul.NBPRISCONGEPRINC := Valeur(Q.findField('PGP_PGVALCHAMP').AsString);

          Q.Next;
        end;
        Ferme (Q);

        // table de correcpondance Jours Acquis / Jours Restants
        Q := OpenSql('SELECT PGP_PGNOMCHAMP,PGP_PGVALCHAMP  FROM PGPARAMETRES WHERE '+
                     'PGP_PGPARAMETRE="'+PGParametreAssoc+'" '+
                     'AND PGP_PGNOMCHAMP  LIKE "GridJoursAcquis%" ', TRUE);
        if (not Q.eof) then
          SetLength (Paramcalcul.TBJoursAcquis, Q.RecordCount);
        ind := 0;
        while not Q.eof do
        begin

// PT1    Paramcalcul.TBJoursAcquis[ind] := StrToFloat(Q.findField('PGP_PGVALCHAMP').AsString);
          Paramcalcul.TBJoursAcquis[ind] := Valeur(Q.findField('PGP_PGVALCHAMP').AsString);
          ind := ind+1;
          Q.Next;
        end;
        Ferme (Q);

{*       Paramcalcul.BornePris := 0;*}
        Q := OpenSql('SELECT PGP_PGNOMCHAMP,PGP_PGVALCHAMP  FROM PGPARAMETRES WHERE '+
                     'PGP_PGPARAMETRE="'+PGParametreAssoc+'" '+
                     'AND PGP_PGNOMCHAMP  LIKE "GridJoursRestant%" ', TRUE);
        if (not Q.eof) then
          SetLength (Paramcalcul.TBJoursRestant, Q.RecordCount);
        ind := 0;
        while not Q.eof do
        begin
//PT1    Paramcalcul.TBJoursRestant[ind] := StrToFloat(Q.findField('PGP_PGVALCHAMP').AsString);
          Paramcalcul.TBJoursRestant[ind] := Valeur(Q.findField('PGP_PGVALCHAMP').AsString);
  {*        if (Paramcalcul.BornePris > Paramcalcul.TBJoursRestant[ind]) or
             (Paramcalcul.BornePris = 0) then
            Paramcalcul.BornePris := Paramcalcul.TBJoursRestant[ind];*}
          ind := ind+1;
          Q.Next;
        end;
        Ferme (Q);
        // Nbre de jours pris en continu minimum pour bénéficier des jours de fractionnement
        Paramcalcul.BornePris := Paramcalcul.NBPRISCONGEPRINC/2;
      end
      else
      // (PGParametreAssoc = '') Pour traiter le cas où les populations ne sont pas gérées
      begin
        Q := OpenSql('SELECT PGP_PGNOMCHAMP,PGP_PGVALCHAMP  FROM PGPARAMETRES WHERE '+
                     'PGP_MODULECHAMP = "ACF" '+
                     'AND PGP_PGNOMCHAMP not LIKE "GridJoursAcquis%" '+
                     'AND PGP_PGNOMCHAMP not LIKE "GridJoursRestant%"', TRUE);
        while not Q.eof do
        begin

          // Types d'acquis donnant droit au fractionnement
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'ACQUISDROITFRAC') then
            Paramcalcul.ACQUISDROITFRAC := Q.findField('PGP_PGVALCHAMP').AsString ;

          // Date fin congé principal
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'CONGEPRINCIPALA') then
            Paramcalcul.CONGEPRINCIPALA  := Q.findField('PGP_PGVALCHAMP').AsString ;

          // Date début congé principal
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'CONGEPRINCIPALDE') then
            Paramcalcul.CONGEPRINCIPALDE := Q.findField('PGP_PGVALCHAMP').AsString;

          // date du calcul des jours de congés pour fractionnement
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'DATECALCULFRAC') then
            Paramcalcul.DATECALCULFRAC := Q.findField('PGP_PGVALCHAMP').AsString;

          // nombre de jours du congé principal
          if (Q.findField('PGP_PGNOMCHAMP').AsString = 'NBPRISCONGEPRINC') then
//PT1      Paramcalcul.NBPRISCONGEPRINC := StrToFloat(Q.findField('PGP_PGVALCHAMP').AsString);
            Paramcalcul.NBPRISCONGEPRINC := Valeur(Q.findField('PGP_PGVALCHAMP').AsString);


          Q.Next;
        end;
        Ferme (Q);

        // table de correcpondance Jours Acquis / Jours Restants
        Q := OpenSql('SELECT PGP_PGNOMCHAMP,PGP_PGVALCHAMP  FROM PGPARAMETRES WHERE '+
                     'PGP_MODULECHAMP = "ACF" '+
                     'AND PGP_PGNOMCHAMP  LIKE "GridJoursAcquis%" ', TRUE);
        if (not Q.eof) then
          SetLength (Paramcalcul.TBJoursAcquis, Q.RecordCount);
        ind := 0;
        while not Q.eof do
        begin
//PT1     Paramcalcul.TBJoursAcquis[ind] := StrToFloat(Q.findField('PGP_PGVALCHAMP').AsString);
          Paramcalcul.TBJoursAcquis[ind] := Valeur(Q.findField('PGP_PGVALCHAMP').AsString);
          ind := ind+1;
          Q.Next;
        end;
        Ferme (Q);


        Q := OpenSql('SELECT PGP_PGNOMCHAMP,PGP_PGVALCHAMP  FROM PGPARAMETRES WHERE '+
                     'PGP_MODULECHAMP = "ACF" '+
                     'AND PGP_PGNOMCHAMP  LIKE "GridJoursRestant%" ', TRUE);
        if (not Q.eof) then
          SetLength (Paramcalcul.TBJoursRestant, Q.RecordCount);
        ind := 0;
        while not Q.eof do
        begin
//PT1     Paramcalcul.TBJoursRestant[ind] := StrToFloat(Q.findField('PGP_PGVALCHAMP').AsString);
          Paramcalcul.TBJoursRestant[ind] := Valeur(Q.findField('PGP_PGVALCHAMP').AsString);
{*          if (Paramcalcul.BornePris < Paramcalcul.TBJoursRestant[ind]) then
            Paramcalcul.BornePris := Paramcalcul.TBJoursRestant[ind];*}
          ind := ind+1;
          Q.Next;
        end;
        Ferme (Q);
        
        // Nbre de jours pris en continu minimum pour bénéficier des jours de fractionnment
        Paramcalcul.BornePris := Paramcalcul.NBPRISCONGEPRINC/2;
      end;
end;


procedure  TOF_PGCALCULACF.CalculJoursFractionnement(const ParamCalcul : TParamCalcul;const TOB_Sal_fille,TOB_PrisPayes,TOB_Acquis : TOB; var T_ACF : TOB; var Trace : TListBox);
var
  wsal:string; //??
  wdate                                 : TDatetime;
  TotPRI                                : double;
  CumTotPRI                             : double;
  TotACQN1                              : double;
  SldCPPrincipal                        : double;
  DroitFrac                             : double;
  PRIContinu                            : double ;
  NbACF                                 : double;
  TOB_PrisPayes_fille                   : TOB;
  i                                     : integer;
  annee                                 : integer;
  CongePrincipalde, CongePrincipalA     : TDatetime;
  WdateFin                              : TDatetime;
  AssezPRIContinu                       : boolean;

begin
  wdate := StrToDate(TOB_Sal_fille.GetValue('DATECLOTURECPN'));
  Wdate := plusdate(Wdate, -1, 'A');
  NbACF := 0;
  TotACQN1 := 0;
  
  if (Date < Wdate) then
  begin
    Trace.Items.Add('Salarié : ' + TOB_Sal_fille.GetValue('SALARIE')+'. Pas de calcul du fractionnement.');
    Trace.Items.Add('                       La clôture des CP n''a pas été réalisée.');
    exit;
  end;
  // Formatage des dates du congé principal passées en paramètre
  annee := StrToInt(copy(TOB_Sal_fille.GetValue('DATECLOTURECPN'),7,4));
  CongePrincipalde := StrToDate(ParamCalcul.CONGEPRINCIPALDE+'/'+IntToStr(annee-1));
  CongePrincipala := StrToDate(ParamCalcul.CONGEPRINCIPALA+'/'+IntToStr(annee-1));

  // somme des pris sur congé principal  (TOB_PrisPayes)
  // & vérification qu'il y a assez de pris en continu
  TotPRI := 0;
  PRIContinu := 0;
  AssezPRIContinu := false;
  WDateFin := IDate1900;

  for i := 0 to TOB_PrisPayes.detail.count-1 do
  begin
    TOB_PrisPayes_fille := TOB_PrisPayes.detail[i];
    if ((TOB_PrisPayes_fille.GetValue('PCN_SALARIE') = TOB_Sal_fille.GetValue('SALARIE')) AND
        (TOB_PrisPayes_fille.GetValue('PCN_DATEDEBUTABS') >=  CongePrincipalde) and
        (TOB_PrisPayes_fille.GetValue('PCN_DATEFINABS') <=  CongePrincipala)) then
    begin
      TotPRI := TotPRI + TOB_PrisPayes_fille.GetValue('PCN_JOURS');

      if (not AssezPRIContinu) then
      begin
        PRIContinu := PRIContinu + TOB_PrisPayes_fille.GetValue('PCN_JOURS');
        if ((TOB_PrisPayes_fille.GetValue('PCN_DATEDEBUTABS') <> WDateFin + 1) and
            ( WDateFin <> IDate1900)) then
          // Pris non consécutifs
          PRIContinu := 0
        else
          // Pris consécutifs
          if (PRIContinu >= ParamCalcul.BornePris) then
            AssezPRIContinu := true;
        WDateFin := TOB_PrisPayes_fille.GetValue('PCN_DATEFINABS');
      end;
    end;
  end;

  // somme des pris jusqu'à fin du congé principal  (TOB_PrisPayes)
  // (des acquis N1 ont pu être consommés avant la clôture des congés)
  CumTotPRI := 0;
  for i := 0 to TOB_PrisPayes.detail.count-1 do
  begin
    TOB_PrisPayes_fille := TOB_PrisPayes.detail[i];
    if ((TOB_PrisPayes_fille.GetValue('PCN_SALARIE') = TOB_Sal_fille.GetValue('SALARIE')) AND
        (TOB_PrisPayes_fille.GetValue('PCN_DATEFINABS') <=  CongePrincipala)) then
      CumTotPRI := CumTotPRI + TOB_PrisPayes_fille.GetValue('PCN_JOURS');
  end;

  // somme des acquis N-1  (TOB_Acquis)
  TotACQN1 :=  arrondi(TOB_Acquis.Somme('PCN_JOURS',
                                ['PCN_SALARIE'],
                                [TOB_Sal_fille.GetValue('SALARIE')],
                                false, False),0);
  wsal := TOB_Sal_fille.GetValue('SALARIE');//??

  // solde des CP à la fin du congé principal = somme des acquis N-1 - somme des pris jusqu'à fin du congé principal
  SldCPPrincipal := TotACQN1 -  CumTotPRI;

  // nombre d'acquis donnant droit au fractionnement (25)
  DroitFrac := arrondi((TOB_Sal_fille.GetValue('NBREACQUISCP')*12),0);

  if (IndTrace) then
  begin
    Trace.Items.Add(wsal);
    Trace.Items.Add('Total pris avant fin du congés principal                 '+FloatToStr(CumTotPRI));
    Trace.Items.Add('Total pris pendant le congés principal                   '+FloatToStr(TotPRI));
    Trace.Items.Add('Nombre de pris en continu pendant le congés principal    '+FloatToStr(TotPRI));
    Trace.Items.Add('Somme des acquis N-1                                     '+FloatToStr(TotACQN1));
    Trace.Items.Add('Solde des CP à la fin du congé principal                 '+FloatToStr(SldCPPrincipal));
    Trace.Items.Add('Nombre d''acquis pour droit au fractionnement             '+FloatToStr(DroitFrac));
  end;

  if (TotACQN1 >= DroitFrac) then
  // Calcul des ACF uniquement si Somme des acquis CP  période n-1
  //                              >=
  //                              nbre d'acquis donnant droit au fractionnement
  begin
    if (TotPRI < ParamCalcul.NBPRISCONGEPRINC) then
    // Calcul des ACF uniquement si Total des Pris dans la période du congé principal
    //                              <
    //                              paramètre nbre de jours pris dans la période
    begin
      if (PRIContinu >= ParamCalcul.BornePris) then
      // Calcul des ACF uniquement si Total des Pris en continu
      //                              >=
      //                              nbre de pris mini nécessaires (la moitié du congé principal)
      begin
        for i:= 0 to (length(ParamCalcul.TBJoursRestant)-1) do
        begin
          if (SldCPPrincipal  >= ParamCalcul.TBJoursRestant[i]) then
            NbACF := ParamCalcul.TBJoursAcquis[i];
        end;
      end;
    end;
  end;
  T_ACF.PutValue('PCN_JOURS', NbACF);

  if (IndTrace) then
  begin
    Trace.Items.Add('Nombre d''acquis pour fractionnement (ACF)-->             '+ FloatToStr(NbACF));
    Trace.Items.Add('');
  end;

end;

initialization
  registerclasses([TOF_PGCALCULACF]);

end.
