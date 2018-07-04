{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... :   /  /
Description .. : Cloture des périodes de paie
Mots clefs ... : PAIE;CLOTURE
*****************************************************************
PT1  | 07/09/2001 | V547  | PH | Fonction RendMois publiée donc mise dans P5DEF
PT2  | 03/06/2002 | V582  | PH | Utilisation HMtrad pour resize de la grille
PT3  | 03/06/2002 | V582  | PH | Gestion historique des évènements
PT4  | 26/06/2003 | V_421 | PH | Mise à jour DP Social si multi et PCL alors obligation
     |            |       | de | clôturer mois par mois
PT5  | 13/10/2004 | V_50  | PH | Mise en place Info DPSOCIAL si PCL et nettoyage journal des
     |            |       |    | évènements.
PT6  | 14/10/2004 | V_50  | PH | FQ 11684 Suppression contrôle blocage monoposte
PT7  | 28/09/2005 | V_60  | PH | FQ 11970 Externalisation fonction mise à jour du DP pour être
     |            |       |    | aussi utilisée à chaque sortie de pgm
PT8  | 28/09/2005 | V_60  | PH | FQ 11970 La clôture sur plusieurs mois lance x clôtures mensuelles
PT9  | 10/04/2006 | V_650 | PH | Prise en compte du GUID de la table DPSOCIAL
PT10 | 15/11/2007 | 8.01  | FL | Clôture multi-dossiers
PT11 | 23/11/2007 | V_810 | FC | Ajout de compteurs - Calcul des compteurs après la clôture car besoin d'être à jour
PT12 | 03/01/2008 | V_802 | FL | FQ 14854 - Proposer l'ouverture de l'exercice social suivant si clôture du dernier mois
}
unit UTofPG_CLOTURE;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  {$IFNDEF EAGLCLIENT}
  fe_main,
  {$ELSE}
  MainEagl,
  {$ENDIF}
  Grids, HCtrls, HEnt1, HMsgBox, HSysMenu, UTOF, UTOB, Vierge,
  ParamSoc,
  P5Util,
  P5Def,
  AGLInit,
  EntPaie,
  PgOutils,
  ed_tools,
  majtable,
  PGAnomaliesTraitement,
  {$IFDEF STATDIR}
  DPMajPaieOutils,        // BTY 23/06/08 FQ 15526
  {$ENDIF}
  PgOutils2;

type
  TOF_PG_CLOTURE = class(TOF)
    procedure OnClose; override;
    procedure OnArgument(Arguments: string); override;
  private
    Grille          : THGrid;
    VCbxLAnnee      : THValComboBox;
    VCbxLMois       : THValComboBox;
    IndiceClot      : string;
    Trou            : Integer;
    ListeAnomalies  : TAnomalies;   //PT10
    Dossier         : String;       //PT10
    DecalagePaie    : Boolean;      //PT10
    Exercice        : String;       //PT10
    ModePCL         : Boolean;      //PT10
    TTMulti         : Boolean;      //PT10
    DBPrefixe       : String;       //PT10

    procedure RecupPeriode;
    Function  RendCloture : String;
    procedure LanceCloture          (DD, DF: TDateTime; MoisClos: String);
    procedure AnneeChange           (Sender: TObject);
    //PT10 - Début
    //procedure ClotureEnter(Sender: TObject);
    procedure Valide                (Sender: TObject);
    procedure ClotureMois           (Annee, MoisACloturer : Integer; EtatCloture : String);
    Procedure DonnePeriodeACloturer (Annee, MoisACloturer : Integer; var EtatCloture : String; var DD, DF : TDateTime);
    Procedure ChangeRegroupement    (Sender : TObject);
    Procedure AfficheListeDossiers  (Sender : TObject);
    Procedure ActiveMultiDossier    (Sender : TObject);
    Procedure ClotureMultiple;
    Procedure OuvrePeriodeSuivante  (DateFinExercice, DebutPeriode, FinPeriode : TDateTime);
    //PT10 - Fin
  end;

  procedure PgAlimDPsocial(DatDeb,DatFin : TDateTime);  // PT7

implementation
{$IFDEF PAIETOX}
uses uToxConf,
  uToxConst,
  uToxFiche,
  PaieToxWork,
  PaieToxMoteur,
  PaieToxIntegre;
{$ENDIF}

type
  //PT10 - Début
  EPaieError     = class(EAbort);       // Exception levée suite à une erreur de cohérence métier
  //PT10 - Fin

const
  Tableau: array[1..12] of string = (
    '01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12');

//DEB PT7
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Alimentation du tableau de bord
Mots clefs ... :
*****************************************************************}
procedure PgAlimDPsocial(DatDeb,DatFin : TDateTime);
var T, T1, TC, TC1: TOB;
  Q, QEtab, QCaisse,QExer: TQuery;
  Ordre: Integer;
  St, EtabDef, GuidPer : String; // PT9
  AA, MM, JJ : WORD;
  DateDebutExerSoc, DateFinExerSoc : TDateTime;
  TobSalaries,TobPaieEnCours : Tob;
  i,CompteurAncInf,CompteurAncSup,CompteurPeriode : integer;
  LaDate : TDateTime;
  PremMois,PremAnnee,NbMois : word;
begin
    T := TOB.create('', nil, -1);
    T1 := TOB.Create('DPTABGENPAIE', T, -1);

    // Nombre de bulletins et total brut des paies
    St := 'SELECT Count (*) N1, SUM (PPU_CBRUT) N2 FROM PAIEENCOURS WHERE PPU_DATEDEBUT>="' + UsDateTime(DatDeb) +
          '" AND PPU_DATEFIN <="' + UsDateTime(DatFin) + '"';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_NBBULLETINS', Q.FindField('N1').AsInteger);
        T1.PutValue('DT1_TOTALBRUT', Q.FindField('N2').AsFloat);
    End;
    Ferme(Q);

    // Nombre de salariés sortis
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATESORTIE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATESORTIE <="' + UsDateTime(DatFin) + '"';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_NBSORTIES', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);

    // Nombre de salariés entrés
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATEENTREE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATEENTREE <="' + UsDateTime(DatFin) + '"';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_NBENTREES', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);

    // Nombre de salariés actuels
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATESORTIE>="' + UsDateTime(DatFin) +
          '" OR PSA_DATESORTIE = "' + UsDateTime(Idate1900) + '" OR PSA_DATESORTIE IS NULL';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_NBPRESENTS', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);

    //DEB PT11
    //Nombre de mutations inter-établissement
    St := 'SELECT Count (*) N1 FROM SALARIES WHERE PSA_DATEENTREE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATEENTREE <="' + UsDateTime(DatFin) + '"' +
          ' AND PSA_MOTIFENTREE="019"';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
      T1.PutValue('DT1_LIBREMONTANT1', Q.FindField('N1').AsInteger);
    Ferme(Q);

    //Nombre de bulletins complémentaires
    St := 'SELECT Count (*) N1 FROM PAIEENCOURS ' +
          ' WHERE PPU_DATEDEBUT>="' + UsDateTime(DatDeb) +  '"' +
          ' AND PPU_DATEFIN <="' + UsDateTime(DatFin) + '"' +
          ' AND PPU_BULCOMPL="X"';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
      T1.PutValue('DT1_LIBREMONTANT2', Q.FindField('N1').AsInteger);
    Ferme(Q);

    //Nombre de bulletins validés
    St := 'SELECT Count (*) N1 FROM PAIEENCOURS ' +
          ' WHERE PPU_DATEDEBUT>="' + UsDateTime(DatDeb) +  '"' +
          ' AND PPU_DATEFIN <="' + UsDateTime(DatFin) + '"' +
          ' AND PPU_TOPCLOTURE="X"';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
      T1.PutValue('DT1_LIBREMONTANT3', Q.FindField('N1').AsInteger);
    Ferme(Q);

    //Nombre de sorties moins de 6 mois d'ancienneté et 6 mois ou plus d'ancienneté
    St := 'SELECT PSA_SALARIE,PSA_DATEANCIENNETE,PSA_REGULANCIEN,PSA_DATESORTIE FROM SALARIES' +
      ' WHERE PSA_DATESORTIE>="' + UsDateTime(DatDeb) + '"' +
      ' AND PSA_DATESORTIE<="' + UsDateTime(DatFin) + '"';
    TobSalaries := Tob.Create('LesSalaries',Nil,-1);
    TobSalaries.LoadDetailDBFromSQL ('LesSalaries', St);
    CompteurAncInf := 0;
    CompteurAncSup := 0;
    For i := 0 to TobSalaries.Detail.Count - 1 do
    begin
      PremMois := 0;
      PremAnnee := 0;
      Ladate := DebutDeMois(PlusMois(TobSalaries.Detail[i].GetValue('PSA_DATEANCIENNETE'),TobSalaries.Detail[i].GetValue('PSA_REGULANCIEN')));
      NOMBREMOIS(Ladate, TobSalaries.Detail[i].GetValue('PSA_DATESORTIE'), PremMois, PremAnnee, NbMois);
      if NbMois < 6 then
        CompteurAncInf := CompteurAncInf + 1
      else
        CompteurAncSup := CompteurAncSup + 1;
    end;
    T1.PutValue('DT1_LIBREMONTANT4', CompteurAncInf);
    T1.PutValue('DT1_LIBREMONTANT5', CompteurAncSup);
    FreeAndNil(TobSalaries);

    //Nombre de bulletins trimestriels, semestriels et annuels : Périodicité > 1 mois
    St := 'SELECT PPU_DATEFIN,PPU_DATEDEBUT FROM PAIEENCOURS ' +
          ' WHERE PPU_DATEFIN>="' + UsDateTime(DatDeb) +  '"' +
          ' AND PPU_DATEFIN <="' + UsDateTime(DatFin) + '"';
    TobPaieEnCours := Tob.Create('LesPaiesEnCours',Nil,-1);
    TobPaieEnCours.LoadDetailDBFromSQL ('LesPaiesEnCours', St);
    CompteurPeriode := 0;
    For i := 0 to TobPaieEnCours.Detail.Count - 1 do
    begin
      PremMois := 0;
      PremAnnee := 0;
      NOMBREMOIS(TobPaieEnCours.Detail[i].GetValue('PPU_DATEDEBUT'), TobPaieEnCours.Detail[i].GetValue('PPU_DATEFIN'), PremMois, PremAnnee, NbMois);
      if NbMois > 1 then
        CompteurPeriode := CompteurPeriode + 1;
    end;
    T1.PutValue('DT1_LIBREMONTANT6', CompteurPeriode);
    FreeAndNil(TobPaieEnCours);

    //Nombre de bulletins de salariés non permanents
    St := 'SELECT Count (*) N1 FROM PAIEENCOURS ' +
          ' WHERE PPU_DATEDEBUT>="' + UsDateTime(DatDeb) +  '"' +
          ' AND PPU_DATEFIN <="' + UsDateTime(DatFin) + '"' +
          ' AND PPU_SALARIE IN (SELECT DISTINCT PHD_SALARIE FROM PGHISTODETAIL P1' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_NEWVALEUR="-"' +
          ' AND PHD_DATEAPPLIC IN (SELECT MAX(PHD_DATEAPPLIC) FROM PGHISTODETAIL P2' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_DATEAPPLIC<="' + UsDateTime(DatDeb) + '"' +
          ' AND P1.PHD_SALARIE=P2.PHD_SALARIE))';
      Q := OpenSql(St, TRUE);
    if not Q.EOF then
      T1.PutValue('DT1_LIBREMONTANT7', Q.FindField('N1').AsInteger);
    Ferme(Q);

    // Nombre d'entrée de salariés non permanents
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATEENTREE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATEENTREE <="' + UsDateTime(DatFin) + '"' +
          ' AND PSA_SALARIE IN (SELECT DISTINCT PHD_SALARIE FROM PGHISTODETAIL P1' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_NEWVALEUR="-"' +
          ' AND PHD_DATEAPPLIC IN (SELECT MAX(PHD_DATEAPPLIC) FROM PGHISTODETAIL P2' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_DATEAPPLIC<="' + UsDateTime(DatDeb) + '"' +
          ' AND P1.PHD_SALARIE=P2.PHD_SALARIE))';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_LIBREMONTANT8', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);

    // Nombre d'entrée de salariés permanents
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATEENTREE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATEENTREE <="' + UsDateTime(DatFin) + '"' +
          ' AND PSA_SALARIE NOT IN (SELECT DISTINCT PHD_SALARIE FROM PGHISTODETAIL P1' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_NEWVALEUR="-"' +
          ' AND PHD_DATEAPPLIC IN (SELECT MAX(PHD_DATEAPPLIC) FROM PGHISTODETAIL P2' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_DATEAPPLIC<="' + UsDateTime(DatDeb) + '"' +
          ' AND P1.PHD_SALARIE=P2.PHD_SALARIE))';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_LIBREMONTANT9', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);

    // Nombre de sorties de salariés non permanents
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATESORTIE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATESORTIE <="' + UsDateTime(DatFin) + '"' +
          ' AND PSA_SALARIE IN (SELECT DISTINCT PHD_SALARIE FROM PGHISTODETAIL P1' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_NEWVALEUR="-"' +
          ' AND PHD_DATEAPPLIC IN (SELECT MAX(PHD_DATEAPPLIC) FROM PGHISTODETAIL P2' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_DATEAPPLIC<="' + UsDateTime(DatDeb) + '"' +
          ' AND P1.PHD_SALARIE=P2.PHD_SALARIE))';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_LIBREMONTANT10', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);

    // Nombre de sorties de salariés permanents
    St := 'SELECT Count (*) FROM SALARIES WHERE PSA_DATESORTIE>="' + UsDateTime(DatDeb) +
          '" AND PSA_DATESORTIE <="' + UsDateTime(DatFin) + '"' +
          ' AND PSA_SALARIE NOT IN (SELECT DISTINCT PHD_SALARIE FROM PGHISTODETAIL P1' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_NEWVALEUR="-"' +
          ' AND PHD_DATEAPPLIC IN (SELECT MAX(PHD_DATEAPPLIC) FROM PGHISTODETAIL P2' +
          ' WHERE PHD_PGINFOSMODIF="C01" AND PHD_DATEAPPLIC<="' + UsDateTime(DatDeb) + '"' +
          ' AND P1.PHD_SALARIE=P2.PHD_SALARIE))';
    Q := OpenSql(St, TRUE);
    if not Q.EOF then
    Begin
        T1.PutValue('DT1_LIBREMONTANT11', Q.Fields[0].AsInteger);
    End;
    Ferme(Q);
    //FIN PT11

    T1.PutValue('DT1_NODOSSIER', PgRendNoDossier);

    //MD 27/04/06 - mis 0 en carac de bourrage, pour pb tri sur mois dans tableau de bord DP
    //T1.PutValue('DT1_MOIS', FloaTToStr(MM));
    DecodeDate(DatFin, AA, MM, JJ);
    T1.PutValue('DT1_MOIS', Format ('%2.2d',[MM]) );
    T1.PutValue('DT1_ANNEE', FloaTToStr(AA));
    T.InsertOrUpdateDB(FALSE);
    FreeAndNil(T);

    // DEB PT5
    // DEB PT9
    // Bilan social
    Q := OpenSql('SELECT DOS_GUIDPER FROM DOSSIER WHERE DOS_NODOSSIER="' + PgRendNoDossier + '"', TRUE);
    if not Q.EOF then
    begin
        GuidPer := Q.FindField('DOS_GUIDPER').AsString;
        FERME(Q);

        T := TOB.create('LE_DP_SOCIAL', nil, -1);
        Q := OpenSQL('SELECT * FROM DPSOCIAL WHERE DSO_GUIDPER="' + GuidPer+'"', TRUE);
        if not q.eof then
        begin
            T.LoadDetailDB('DPSOCIAL', '', '', Q, False);
            T1 := T.FindFirst(['DSO_GUIDPER'], [GuidPer], FALSE);
        end
        else
        begin
            T1 := TOB.Create('DPSOCIAL', T, -1);
        end;
        Ferme(Q);

        if T1 <> nil then
        begin
            T1.PutValue('DSO_GUIDPER', GuidPer);
            if VH_Paie.PGDecalage then T1.PutValue('DSO_PAIEDECALEE', 'X') //PT10
            else T1.PutValue('DSO_PAIEDECALEE', '-');
            if VH_Paie.PGLibCodeStat <> '' then T1.PutValue('DSO_PAIESTATS', 'X')
            else T1.PutValue('DSO_PAIESTATS', '-');
            if VH_Paie.PGAnalytique then T1.PutValue('DSO_PAIEANALYTIQUE', 'X')
            else T1.PutValue('DSO_PAIEANALYTIQUE', '-');
            if VH_Paie.PGJournalPaie <> '' then T1.PutValue('DSO_PAIEGENECRIT', 'X')
            else T1.PutValue('DSO_PAIEGENECRIT', '-');
            T1.PutValue('DSO_PAIECAB', 'X');
            QExer := OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="' + USDateTime(DebutDeMois(DatFin)) + '" ' +
            'AND PEX_DATEFIN>="' + USDateTime(DatFin) + '"', True);
            if not QExer.eof then
            begin
                DateDebutExerSoc := QExer.FindField('PEX_DATEDEBUT').AsDateTime;
                DateFinExerSoc := QExer.FindField('PEX_DATEFIN').AsDateTime;
            end
            else
            begin
                DateDebutExerSoc := idate1900;
                DateFinExerSoc := IDate1900;
            end;
            Ferme(QExer);
            T1.PutValue('DSO_DATEDEBEX', DateDebutExerSoc);
            T1.PutValue('DSO_DATEFINEX', DateFINExerSoc);
            T1.PutValue('DSO_DATEDERPAIE', DatFin);
            if VH_Paie.PGCongesPayes then T1.PutValue('DSO_GESTCONGES', 'X')
            else T1.PutValue('DSO_GESTCONGES', '-');
            if VH_Paie.PGMsa then T1.PutValue('DSO_MUTSOCAGR', 'X')
            else T1.PutValue('DSO_MUTSOCAGR', '-');
            if VH_Paie.PGIntermittents then T1.PutValue('DSO_INTERMSPEC', 'X')
            else T1.PutValue('DSO_INTERMSPEC', '-');
            if VH_Paie.PGTicketRestau then T1.PutValue('DSO_TICKETREST', 'X')
            else T1.PutValue('DSO_TICKETREST', '-');
            if VH_Paie.PGBTP then T1.PutValue('DSO_BTP', 'X')
            else T1.PutValue('DSO_BTP', '-');
            EtabDef := GetParamSoc('SO_ETABLISDEFAUT');
            if EtabDef = '' then
            begin // cas où il n'ya pas d'etablissement par defaut dans la compta
                QEtab := OpenSql('SELECT ETB_ETABLISSEMENT FROM ETABCOMPL ORDER BY ETB_ETABLISSEMENT"', TRUE);
                if not QEtab.eof then EtabDef := QEtab.FindField('ETB_ETABLISSEMENT').AsString;
                Ferme (QEtab);
            end;
            if EtabDef <> '' then
            begin
                QEtab := OpenSql('SELECT * FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' + EtabDef + '"', TRUE);
                if not QEtab.eof then
                begin
                    if (QEtab.FindField('ETB_ACTIVITE').AsString <> '') then T1.PutValue('DSO_PLANPAIEACT', 'X')
                    else T1.PutValue('DSO_PLANPAIEACT', '-');
                    T1.PutValue('DSO_CONVENCOLLEC', QEtab.FindField('ETB_CONVENTION').AsString);
                end;
                Ferme(QEtab);
                QEtab := OpenSql('SELECT PAT_TAUXAT FROM TAUXAT WHERE PAT_ETABLISSEMENT="' + EtabDef + '" ORDER BY PAT_ORDREAT DESC', TRUE);
                if not QEtab.eof then T1.PutValue('DSO_TAUXACCTRAV', QEtab.FindField('PAT_TAUXAT').AsFloat)
                else T1.PutValue('DSO_TAUXACCTRAV', 0);
                Ferme(QEtab);
                QCaisse := OpenSql('SELECT POG_ORGANISME,POG_LIBELLE,POG_NATUREORG,POG_PERIODICITDUCS FROM ORGANISMEPAIE WHERE POG_ETABLISSEMENT="' + EtabDef + '" ORDER BY POG_ORGANISME', TRUE);
                if not QCaisse.eof then
                begin
                    TC := TOB.create('LES_CAISSES', nil, -1);
                    Ordre := 0;
                    while not QCaisse.EOF do
                    begin
                        Ordre := Ordre + 1;
                        TC1 := TOB.Create('DPSOCIALCAISSE', TC, -1);
                        TC1.PutValue('DSC_NODP', 0);
                        TC1.PutValue('DSC_GUIDPER', GuidPer);
                        TC1.PutValue('DSC_ORDRE', Ordre);
                        TC1.PutValue('DSC_CODECAISSE', QCaisse.FindField('POG_ORGANISME').AsString);
                        TC1.PutValue('DSC_NOMCAISSE', QCaisse.FindField('POG_LIBELLE').AsString);
                        TC1.PutValue('DSC_NATUREDUCS', RECHDOM('PGTYPORGDUCS', QCaisse.FindField('POG_NATUREORG').AsString, FALSE));
                        TC1.PutValue('DSC_PERIODICITE', RECHDOM('PGPERIODICITEDUCS', QCaisse.FindField('POG_PERIODICITDUCS').AsString, FALSE));
                        QCaisse.Next;
                    end;
                    ExecuteSQL('DELETE FROM DPSOCIALCAISSE WHERE DSC_GUIDPER="' + GuidPer+'"');
                    TC.InsertOrUpdateDB(FALSE);
                    FreeAndNil(TC);
                end;
                Ferme(QCaisse);
            end;
            T.InsertOrUpdateDB(FALSE);
            FreeAndNil(T);
        end;
    end
    else FERME(Q);
    // FIN PT9
end;
// FIN PT7

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Récupère la clôture actuel de l'exercice et rafraîchit le tableau
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_CLOTURE.AnneeChange(Sender: TObject);
var
  i, j, max: Integer;
  stClos,stOuvert : String;
begin
    if Grille = nil then exit;

    stClos   := TraduireMemoire('Clos');
    stOuvert := TraduireMemoire('A Cloturer');

    Exercice := vcbxLAnnee.value; //PT10

    // Récupère l'état de clôture existant
    IndiceClot := RendCloture;
    if IndiceClot = '' then IndiceClot := '------------';

    // Affiche la liste des mois
    RecupPeriode;

    // Affiche l'état de chaque mois
    max := 12;
    if VH_Paie.PGDecalage = FALSE then i := 1 else i := 2;
    if i = 2 then
    begin
        if IndiceClot[1] = 'X' then Grille.Cells[1, 1] := stClos
        else Grille.Cells[1, 1] := stOuvert;
    end;
    for j := i to max do
    begin
        if IndiceClot[j] = 'X' then
        begin
            Grille.Cells[1, j] := stClos;
            Trou := i;
        end
        else Grille.Cells[1, j] := stOuvert;
    end;

    //PT10 - Début
    // Préselection du mois à clôturer dans la grille
    If Pos('-',IndiceClot) > 0 Then
    Begin
        Grille.Row := Pos('-',IndiceClot) ;
        // Préselection du mois à clôturer dans la combo
        VCbxLMois.Value := Format('%.2d',[RendMois(Grille.Row)]);
    End
    Else
    Begin
        Grille.Row := 1;
        VCbxLMois.Value := '';
    End;
    //PT10 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/11/2007
Modifié le ... :   /  /
Description .. : Gestion du clic sur le bouton Valider
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_CLOTURE.Valide(Sender: TObject);
begin
    if (VCBXLMois = Nil) Or (VCBXLAnnee = Nil) then
        Exit;

    If (VCBXLMois.Text =  '') Or (VCBXLAnnee.Text = '') Then
        Exit;

    //PT10 - Début
    If GetControlText ('CKCLOTUREMULTI') = '-' Then
    Begin
        TTMulti := False;
        SetControlVisible('DETAIL', False);
        ClotureMois(StrToInt(vcbxLAnnee.Text), StrToInt(VCBXLMois.Value), IndiceClot)
    End
    Else
    Begin
        TTMulti := True;
        SetControlVisible('DETAIL', True);
        ClotureMultiple;
    End;
    //PT10 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/11/2007
Modifié le ... :   /  /
Description .. : Détermine la période à clôturer pour un mois et une année donnée
Mots clefs ... :
*****************************************************************}
Procedure TOF_PG_CLOTURE.DonnePeriodeACloturer (Annee, MoisACloturer : Integer; var EtatCloture : String; var DD, DF : TDateTime);
var
  Min, Max, i, Mois: WORD;
  Lig: Integer;
  LibMessage : String;
begin
    DD := iDate1900;
    DD := iDate1900;

    //PT10 - Début
    //if VCBXLMois = nil then exit;
    //Mois := StrToInt(VCBXLMois.Value);
    Mois := MoisACloturer;
    //Decal := VH_Paie.PGDecalage;
    //VH_Paie.PGDecalage := Decalage;
    //PT10 - Fin

    //if VH_Paie.PGDecalage = TRUE then
    If DecalagePaie Then
    begin
        if Mois = 12 then Lig := 1
        else Lig := Mois + 1;
    end
    else Lig := Mois;

    if Lig = 0 then exit; // titre non accessible

    //if Grille.Cells[1, Lig] = TraduireMemoire('Clos') then //PT10
    If EtatCloture[Lig] = 'X' Then
    begin
        LibMessage := TraduireMemoire('Vous ne pouvez pas clôturer une période déjà close.');
        If ListeAnomalies = Nil Then
            PGIBox(LibMessage, Ecran.Caption)
        Else
            Raise EPaieError.Create(LibMessage); //PT10
        exit; // Periode close on ne va pas la recloturer
    end
    else
    begin
        Max := Lig;
        Min := Max;
        for i := 1 to max do
        begin
            //if IndiceClot[i] <> 'X' then //PT10
            If EtatCloture[i] <> 'X' Then
            begin
                Min := i;
                break;
            end;
        end;
        //MoisClos := IndiceClot; //PT10
        for i := Min to max do
        begin
            //MoisClos[i] := 'X'; //PT10
            EtatCloture[i] := 'X';
        end;
        Mois := RendMois(Min);
        //if Not Decalage then DD := EncodeDate(StrToInt(vcbxLAnnee.Text), Mois, 1) // Recup debut 1er mois A clore //PT10
        if Not DecalagePaie then DD := EncodeDate(Annee, Mois, 1) // Recup debut 1er mois A clore
        else
        begin
            //if lig = 1 then DD := EncodeDate(StrToInt(vcbxLAnnee.Text) - 1, Mois, 1) // Annee - 1 car paie decalée //PT10
            if lig = 1 then DD := EncodeDate(Annee - 1, Mois, 1) // Annee - 1 car paie decalée
            else DD := EncodeDate(Annee, Mois, 1);
        end;

        Mois := RendMois(Max);
        //DF := EncodeDate(StrToInt(vcbxLAnnee.Text), Mois, 1); //PT10
        DF := EncodeDate(Annee, Mois, 1);
        //if Decalage and (Mois = 12) then DF := EncodeDate(StrToInt(vcbxLAnnee.Text) - 1, Mois, 1); //PT10
        if DecalagePaie and (Mois = 12) then DF := EncodeDate(Annee - 1, Mois, 1);
        DF := FINDEMOIS(DF); // Recup dernier Mois à clôturer
    End;

    If (V_PGI.SAV) And (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  Période à clôturer du ['+DateToStr(DD)+'] au ['+DateToStr(DF)+']');
End;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Clôture un mois de paie
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_CLOTURE.ClotureMois (Annee, MoisACloturer : Integer; EtatCloture : String);
var
  DD, DF, D1, D2: TDateTime;
  St: string;
  ReponseOuv : Integer;
  Q : TQuery;
  PerDebut,PerFin,DateFinExe : TDateTime;
begin
(*  //PT10
    if VCBXLMois = nil then exit;
    Mois := StrToInt(VCBXLMois.Value);

  if VH_Paie.PGDecalage = TRUE then
  begin
    if Mois = 12 then Lig := 1
    else Lig := Mois + 1;
  end
  else Lig := Mois;

  if Lig = 0 then exit; // titre non accessible

  if Grille.Cells[1, Lig] = TraduireMemoire('Clos') then
  begin
    PGIBox('Vous ne pouvez pas clôturer une période déjà close', 'Clôture des Paies');
    exit; // Periode close on ne va pas la recloturer
  end
  else
  begin
    Max := Lig;
    Min := Max;
    for i := 1 to max do
    begin
      if IndiceClot[i] <> 'X' then
      begin
        Min := i;
        break;
      end;
    end;
    MoisClos := IndiceClot;
    for i := Min to max do
    begin
      MoisClos[i] := 'X';
    end;
    Mois := RendMois(Min);
    if Not Decalage then DD := EncodeDate(StrToInt(vcbxLAnnee.Text), Mois, 1) // Recup debut 1er mois A clore
    else
    begin
      if lig = 1 then DD := EncodeDate(StrToInt(vcbxLAnnee.Text) - 1, Mois, 1) // Annee - 1 car paie decalée
      else DD := EncodeDate(Annee, Mois, 1);
    end;

    Mois := RendMois(Max);
    DF := EncodeDate(StrToInt(vcbxLAnnee.Text), Mois, 1);
    if Decalage and (Mois = 12) then DF := EncodeDate(StrToInt(vcbxLAnnee.Text) - 1, Mois, 1);
    DF := FINDEMOIS(DF); // Recup dernier Mois à clôturer
*)
    DonnePeriodeACloturer (Annee, MoisACloturer, EtatCloture, DD, DF);  //PT10

    if DD > DF then
    begin
        PGiBox(TraduireMemoire('Vous devez clôturer vos paies mois par mois.'), Ecran.caption);
        exit;
    end;
    //St := 'Voulez-vous clôturer les paies du ' + DateToStr(DD) + ' au ' + DateToStr(DF);  //PT10
    //if PGIAsk(St, 'Clôture des Paies') = mrYes then
    St := TraduireMemoire('Voulez-vous clôturer les paies du %s au %s ?');
    if PGIAsk(Format(St,[DateToStr(DD),DateToStr(DF)]), Ecran.Caption) = mrYes Then
    begin
        // DEB PT8
        if (DD <> DebutDeMois (DF)) then
        begin // Plusieurs mois
            D1 := DD;
            D2 := FinDeMois (D1);
            While D2 <= DF do
            begin
                //LanceCloture(D1, D2, MoisClos); //PT10
                LanceCloture(D1, D2, EtatCloture);
                D1 := PLUSMOIS (D1, 1);
                D2 := FindEmois (D1);
            end;
        end
        //else LanceCloture(DD, DF, MoisClos); // Clôture d'un seul mois //PT10
        else LanceCloture(DD, DF, EtatCloture); // Clôture d'un seul mois
        // FIN PT8

        //PT10 - Début
        St := TraduireMemoire('Voulez-vous ouvrir automatiquement la période suivante?');
        ReponseOuv := PGIAsk(St, Ecran.Caption);
        If ReponseOuv = mrYes Then
        Begin
            DateFinExe  :=iDate1900;
            PerDebut    :=iDate1900;
            PerFin      :=iDate1900;
            Q := OpenSQL('SELECT PEX_DATEFIN,PEX_DEBUTPERIODE,PEX_FINPERIODE FROM EXERSOCIAL WHERE PEX_EXERCICE="'+Exercice+'"', True);
            If Not Q.EOF Then
            Begin
                PerDebut     := Q.FindField('PEX_DEBUTPERIODE').AsDateTime;
                PerFin       := Q.FindField('PEX_FINPERIODE').AsDateTime;
                DateFinExe   := Q.FindField('PEX_DATEFIN').AsDateTime;
            End;
            Ferme(Q);
            OuvrePeriodeSuivante (DateFinExe,PerDebut,PerFin);
        End;
        //PT10 - Fin

        AnneeChange(nil); // Pour reafficher les clotures effectuees
    end;
(*  end; *)
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Exécute la clôture de la période spécifiée
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_CLOTURE.LanceCloture(DD, DF: TDateTime; MoisClos: string);
var
  st, St1, S1, S2   : string;
  Trace             : TStringList;
  Err               : Boolean;
  LaDate, LaDateEjvt: TDateTime;
  OkOk              : Boolean;
  Code              : String; //PT10
begin
    {$IFDEF PAIETOX}
    // PT3 : 03/06/2002 V582 PH Gestion historique des évènements
    //  if GetParamSoc ('SO_PGLANCETOX', TRUE) then
    if 2 = 1 then
    begin // Cas spécifique lancement automatisé du traitement des échanges de la TOX
        st := 'UPDATE STOXEVENTS SET SEV_ACTIF="X" WHERE SEV_ACTIF="-" AND SEV_TYPETRT="001" ' +
              'AND SEV_CODEEVENT LIKE "PGS5%"';
        ExecuteSql(St); //  aceStart         , PaieAvantArchivageTOX
        AglToxConf(aceConfigure, 'PAIES5', PaieNotConfirmeTox, PaieTraitementTox, AglToxFormError);
        Delay(1000);
        exit;
    end;
    {$ENDIF}
    Trace := TStringList.Create;
    Err := FALSE;
    LaDate := DebutDemois(DF);

    If (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  '+Format(TraduireMemoire('Clôture des paies du %s au %s'), [DateToStr(DD), DateToStr(DF)]));

    // PT4 : 26/06/2003 V_421 PH Mise à jour DP Social
    OkOk := True;
    //if V_PGI.ModePCL = '1' then //PT10
    If ModePCL Then
    begin
        if (LaDate <> DD) then
        begin
            PgiBox('Vous devez clôturer les paies mois par mois !', Ecran.Caption);
            okok := FALSE;
        end;
    end
    else OkOk := FALSE;

    // PT4
    // Epuration du journal des évènements
    LaDateEjvt := Date;
    LaDateEjvt := DebutDeMois(LaDateEjvt);
    LaDateEjvt := PLUSMOIS(LaDateEjvt, -6);
    If (V_PGI.SAV) And (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  Epuration des évènements du journal en date du ['+DateToStr(LaDateEjvt)+']');
    st := 'DELETE FROM '+DBPrefixe+'PAIEJEVT WHERE PJE_DATEEVENT < "' + UsDateTime(LaDateEjvt) + '"'; //PT10
    ExecuteSql(st);

    Try
        BeginTrans;

        // Clôture des paies en cours pour chaque salarié
        st := 'UPDATE '+DBPrefixe+'PAIEENCOURS SET PPU_TOPCLOTURE="X" WHERE PPU_DATEDEBUT>="' + UsDateTime(DD) + //PT10
              '" AND PPU_DATEFIN <="' + UsDateTime(DF) + '"';
        If (V_PGI.SAV) And (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  Clôture des paies en cours...');
        ExecuteSql(st);

        // Clôture de l'exercice social
        If (V_PGI.SAV) And (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  Clôture de l''exercice social tel que ['+MoisClos+'] pour l''année ['+vcbxLAnnee.value+']');
        st := 'UPDATE '+DBPrefixe+'EXERSOCIAL SET PEX_CLOTURE="' + MoisClos + '" WHERE PEX_EXERCICE="' + Exercice + '"'; //PT10
        ExecuteSql(st);

        CommitTrans;

    Except
        Rollback;

        St := TraduireMemoire('Une erreur est survenue lors de la clôture.');
        If (ListeAnomalies = Nil) Then
            PGIBox(St, Ecran.Caption);
        err := TRUE;
    end;

    {$IFDEF STATDIR}
    // BTY 23/06/08 FQ 15526
    if ModePCL and not TTMulti then DPMajDP_Paie (DD, DF, PGRendNoDossier,'CLOTURE');
    {$ENDIF}

    //DEB PT11
    //Code déplacé après la màj du topcloture car besoin pour les compteurs
    if OkOk then
    //if (V_PGI.ModePCL = '1') then PgAlimDPsocial(DD, DF) // PT7 //PT10
    If ModePcl And not TTMulti {//tAnd (DBDos=Nil)} Then PGAlimDPSocial(DD,DF)
    else exit; // Pas de lancement de la clôture des  paies
    //FIN PT11
    
    // PT3 : 03/06/2002 V582 PH Gestion historique des évènements
    S1 := DateToStr(DD);
    S2 := DateToStr(DF);
    St1 := '';

    if Err then st1 := 'Erreur de ';
    //PT10 - Début
    If Dossier<>'' Then 
    Begin
    	St1 := 'Dossier : '+Dossier+' | ' + St1; 
    	Code := '300';
    End
    Else Code := '004';
    //PT10 - Fin
    St := 'Clôture des paies';
    Trace.add(St1 + St + ' du ' + S1 + ' au ' + S2);
    if Err then
    Begin
        CreeJnalEvt('001', Code, 'ERR', nil, nil, Trace);
        Raise EPaieError.Create(TraduireMemoire('Une erreur est survenue lors de la clôture.'));
    End
    else
    begin
        CreeJnalEvt('001', Code, 'OK', nil, nil, Trace);
        If ListeAnomalies <> Nil Then ListeAnomalies.Add(INFO1, '  '+TraduireMemoire('> Clôture OK')); //PT10
    end;
    Trace.free;
    // FIN PT3
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Lancement de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_CLOTURE.OnArgument(Arguments: string);
var
  MoisE, AnneeE, ComboExer: string;
  BtnVal: TToolbarButton97;
  DebExer, FinExer: TDateTime;
  CB        : THValComboBox;
  NbRegr    : Integer;
  Edit      : THEdit;
  Ck        : TCheckBox;
  MultiDos  : Boolean;
begin
  inherited;

    BtnVal := TToolbarButton97(GetControl('BValider'));
    if BtnVal <> nil then BtnVal.OnClick := Valide; //PT10

    { // PT6
    if not BlocageMonoPoste(TRUE) then
    rep := PgiAsk('Attention, des utilisateurs sont déjà connectés.#13#10Voulez-vous quand même continuer ?', Ecran.caption)
    else rep := mrYes;

    if rep <> mrYes then
    begin
    SetControlEnabled('BVALIDER', FALSE);
    exit;
    end;}

    Grille := THGrid(GetControl('GRILLECLOTURE'));
    VCbxLAnnee := THValComboBox(GetControl('VCBXANNEE'));
    VCbxLMois := THValComboBox(GetControl('VCBXMOIS'));

    if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
    begin
        if vcbxLAnnee <> nil then
        begin
            Trou := 1;
            vcbxLAnnee.OnChange := AnneeChange;
            vcbxLAnnee.value := ComboExer;
            //if Trou < 12 then Grille.Row := Trou + 1 else Grille.Row := 1; //PT10
        end;
    end;
    if VCbxLMois <> nil then VCbxLMois.Value := MoisE;
    //PT2 : 03/06/2002 V582 PH Utilisation HMtrad pour resize de la grille
    TFVierge(Ecran).HMTrad.ResizeGridColumns(Grille);

    //PT10 - Début
    // Gestion du multi-dossier
	MultiDos        := (Arguments='MULTI');
    ListeAnomalies  := Nil;
    Dossier         := '';
    DecalagePaie    := VH_PAIE.PGDecalage;
    ModePCL         := (V_PGI.ModePCL='1');
    DBPrefixe       := '';

    If MultiDos Then
    Begin
        SetControlVisible('GBMULTIDOS', True);
        CB := THValComboBox(GetControl('MULTIDOSSIER'));
        NbRegr := 0;
        If CB <> Nil Then
        Begin
            NbRegr := CB.Items.Count;
            CB.OnChange := ChangeRegroupement;
        End;

        // Si aucun regroupement de paramétré, pas de sélection possible
        If NbRegr = 0 Then
        Begin
            SetControlEnabled('CKCLOTUREMULTI', False);
        End
        Else
        Begin
            Edit := THEdit(GetControl('DOSSIERS'));
            If Edit <> Nil Then Edit.OnElipsisClick := AfficheListeDossiers;

            Ck := TCheckBox(GetControl('CKCLOTUREMULTI'));
            If Ck <> Nil Then Ck.OnClick := ActiveMultiDossier;
        End;
    End;
    //PT10 - Fin
end;

procedure TOF_PG_CLOTURE.OnClose;
begin
// PT6  DeblocageMonoPoste(TRUE);
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Fonction de remplissage du libellé des mois dans la 1ere colonne en fonction du
Suite ........ : décalage de paie
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_CLOTURE.RecupPeriode;
var
  i, j, max, val: Integer;
begin
    if Grille = nil then exit;
    max := 12;
    // On va remplir la 1ere colonne qui represente les mois de l'exercice social
    if VH_Paie.PGDecalage = FALSE then
    begin
        i := 1;
    end
    else
    begin
        i := 2;
    end;
    if i = 2 then
    begin
        Grille.Cells[0, 1] := RechDOM('PGMOIS', TABLEAU[12], FALSE);
        Grille.Cells[0, 2] := RechDOM('PGMOIS', TABLEAU[1], FALSE);
        i := i + 1;
    end;
    for j := i to max do
    begin
        if VH_Paie.PGDecalage = FALSE then val := j
        else val := j - 1;
        Grille.Cells[0, j] := RechDOM('PGMOIS', TABLEAU[val], FALSE);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Récupération de l'état de clôture des paies actuel
Mots clefs ... :
*****************************************************************}
Function TOF_PG_CLOTURE.RendCloture : String;
var
  Q: TQuery;
  StSQL : String;
  Indice : String; //PT10
begin
    //PT10 - Début
    Indice := '';

    StSQL := 'SELECT PEX_CLOTURE FROM '+DBPrefixe+'EXERSOCIAL WHERE PEX_EXERCICE="' + Exercice + '"';
    Q := OpenSQL(StSQL, TRUE);
    //PT10 - Fin

    if not Q.EOF then
    begin
        Indice := Q.FindField('PEX_CLOTURE').AsString; //PT10
    end;
    Ferme(Q);

    If (V_PGI.SAV) And (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  Indice de clôture = ['+Indice+']');
    Result := Indice; //PT10
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/11/2007 / PT10
Modifié le ... :   /  /
Description .. : Gestion du changement de regroupement par l'utilisateur
Mots clefs ... :
*****************************************************************}
Procedure TOF_PG_CLOTURE.ChangeRegroupement (Sender : TObject);
Begin
    SetControlText('DOSSIERS', '');
    SetControlText('NUM_DOSSIERS', '');
    If (Sender <> Nil) And (THValComboBox(Sender).Text <> '') Then AfficheListeDossiers(Nil);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/11/2007  / PT10
Modifié le ... :   /  /
Description .. : Gestion du clic sur l'elipsis de la liste des dossiers
Mots clefs ... :
*****************************************************************}
Procedure TOF_PG_CLOTURE.AfficheListeDossiers (Sender : TObject);
var Retour : String;
Begin
    If GetControlText('MULTIDOSSIER') <> '' Then
    Begin
        Retour := AGLLanceFiche ('PAY', 'MULTIDOSSIERS', '', '', 'EXERSOCIAL|'+vcbxLAnnee.Text+';'+GetControlText('MULTIDOSSIER')+';'+GetControlText('NUM_DOSSIERS'));
        If Retour <> '' Then
        Begin
            SetControlText('NUM_DOSSIERS',ReadTokenPipe(Retour, '|'));
            SetControlText('DOSSIERS',    Retour);
        End;
    End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/11/2007
Modifié le ... :   /  /
Description .. : Affiche ou cache le group box dédié au multi-dossier
Mots clefs ... :
*****************************************************************}
Procedure TOF_PG_CLOTURE.ActiveMultiDossier (Sender : TObject);
Var Etat : Boolean;
Begin
    Etat := False;
    If Sender <> Nil Then Etat := TCheckBox(Sender).Checked;

    SetControlEnabled('MULTIDOSSIER',   Etat);
    SetControlEnabled('TMULTIDOSSIER',  Etat);
    SetControlEnabled('DOSSIERS',       Etat);
    SetControlEnabled('TDOSSIERS',      Etat);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 21/11/2007 / PT10
Modifié le ... :   /  /
Description .. : Ouverture de la période de paie suivante
Mots clefs ... :
*****************************************************************}
Procedure TOF_PG_CLOTURE.OuvrePeriodeSuivante (DateFinExercice, DebutPeriode, FinPeriode : TDateTime);
var
  Annee, Mois, Jour: Word;
  StSQL,LibMessage : String;
Begin
    LibMessage := '';

    If (DebutPeriode > iDate1900) and (FinPeriode > iDate1900) then
    Begin
        DecodeDate(FinPeriode, Annee, Mois, Jour);

        If ((Not DecalagePaie) And (Mois = 12)) Or (DecalagePaie And (Mois = 11)) Or (DateFinExercice = FinPeriode) Then
        Begin
        	//PT12 - Début           	
        	{ S'il existe une liste des anomalies, on est en multidossier. Dans ce cas, arrivé en fin d'exercice
        	  on ne fait rien. Sinon, on propose à l'utilisateur de créer un nouvel exercice. }
            If ListeAnomalies = Nil Then
            Begin
				DecodeDate(DateFinExercice, Annee, Mois, Jour);
            	Annee := Annee + 1;
      			If ExisteSQL('SELECT 1 FROM EXERSOCIAL WHERE PEX_ANNEEREFER="'+IntToStr(Annee)+'"') Then
      				LibMessage := TraduireMemoire('La fin de l''exercice social en cours est atteinte.') + ' ' + TraduireMemoire('La période de paie suivante ne peut être ouverte.')
      			Else
      			Begin
	            	If PGIAsk(TraduireMemoire('Voulez-vous ouvrir l''exercice social suivant?')) = mrYes Then
         				AglLanceFiche('PAY', 'EXERSOCIAL', '', '', 'ACTION=CREATION');
            	End;
            End
            Else
				LibMessage := TraduireMemoire('La période de paie suivante ne peut être ouverte.');
            //PT12 - Fin
        End
        Else
        Begin
            StSQL := 'UPDATE '+DBPrefixe+'EXERSOCIAL SET PEX_DEBUTPERIODE="'+UsDateTime(PlusMois(DebutPeriode, 1))+'",'+ //PT10
                     'PEX_FINPERIODE="'+UsDateTime(FindeMois(PlusMois(FinPeriode, 1)))+'" '+
                     'WHERE PEX_EXERCICE="'+Exercice+'"';
            ExecuteSQL(StSQL);

            If (ListeAnomalies<>Nil) Then ListeAnomalies.Add(INFO1, '  '+Format(TraduireMemoire('> La période du %s au %s a été ouverte.'),[DateToStr(PlusMois(DebutPeriode, 1)),DateToStr(FindeMois(PlusMois(FinPeriode, 1)))]));
        End;
    End
    Else
    Begin
        LibMessage := TraduireMemoire('La période de paie suivante ne peut être ouverte.');
    End;

    If LibMessage <> '' Then
    Begin
        If (ListeAnomalies<>Nil) Then
            Raise EPaieError.Create(LibMessage)
        Else
            PGIBox(LibMessage, Ecran.Caption);
    End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 19/11/2007 / PT10
Modifié le ... :   /  /
Description .. : Clôture des paies sur plusieurs bases
Mots clefs ... :
*****************************************************************}
Procedure TOF_PG_CLOTURE.ClotureMultiple;
var NumDossier                  : String;
    i,NbDos                     : Integer;
    NomsDossiers,NumDossiers    : TStringList;
    AnneeRef,MoisRef            : Integer;
    Actif                       : Boolean;
    IndiceClotTemp              : String;
    Q                           : TQuery;
    DD,DF,DateFinExe            : TDateTime;
    PerDebut, PerFin            : TDateTime;
    Reponse,ReponseOuv          : Integer;
    LibMessage                  : String;
    Year,Month,Day              : WORD;
    ListeBox                    : TListBox;
Begin
    If (GetControlText('MULTIDOSSIER') = '') Or (GetControlText('NUM_DOSSIERS') = '') Then
    Begin
        PGIBox(TraduireMemoire('Veuillez sélectionner les sociétés à traiter.'), Ecran.Caption);
        Exit;
    End
    Else
    Begin
        // Récupération de la sélection qui servira de référence pour le traitement
        AnneeRef := StrToInt(VCBXLAnnee.Text);
        MoisRef  := StrToInt(VCBXLMois.Value);

        ListeBox := TListBox(GetControl('LBANO'));
        If ListeBox <> Nil Then ListeBox.Clear;

        // Création de la liste des anomalies pour suivre le déroulement du traitement
        ListeAnomalies := TAnomalies.Create;
        ListeAnomalies.SetDuplicateSymbols := True;
        ListeAnomalies.SetUnderlinedTitles := True;
        ListeAnomalies.ChangeLibAno(INFO1, TraduireMemoire('Déroulement du traitement :'));
        ListeAnomalies.ChangeLibAno(ERR1,  TraduireMemoire('Dossiers non traités (en erreur) :'));

        // Création de la liste des dossiers à traiter
        NomsDossiers := TStringList.Create;
        NomsDossiers.Delimiter := ';';
        NomsDossiers.DelimitedText := GetControlText('DOSSIERS');

        NumDossiers := TStringList.Create;
        NumDossiers.Delimiter := ';';
        NumDossiers.DelimitedText := GetControlText('NUM_DOSSIERS');

        NbDos := NumDossiers.Count;

        // Démarrage de l'écran d'attente
        InitMoveProgressForm(Nil, TraduireMemoire('Traitement en cours. Veuillez patienter...'), TraduireMemoire('Clôture du mois de ')+RechDOM('PGMOIS', IntToStr(MoisRef), False), NbDos,  True, True);

        // Affichage du détail du traitement
        SetFocusControl ('LBANO');

        Reponse    := -1;
        ReponseOuv := -1;

        Actif        := False;
        PerDebut     := iDate1900;
        PerFin       := iDate1900;
        DateFinExe   := iDate1900;

        // On lance le traitement sur chaque dossier
        For i := 0 To NbDos-1 Do
        Begin
            // Réinitialisation des paramètres
            DecalagePaie := False;
            ModePCL      := False;
            Exercice     := '';
            Dossier      := NomsDossiers[i];
            NumDossier   := NumDossiers[i];
            DBPrefixe    := Dossier+'.DBO.';
            If ModePCL Then DBPrefixe := 'DB'+DBPrefixe;

            LibMessage := Format(TraduireMemoire('Traitement du dossier %s / %s : %s'), [IntToStr(i+1),IntToStr(NbDos),Dossier]);
            ListeAnomalies.Add(INFO1, LibMessage);

            If Not MoveCurProgressForm(LibMessage) Then
            Begin
                // Clic sur abandon
                If Q <> Nil Then Ferme(Q);
                ListeAnomalies.Add(INFO1, '');
                ListeAnomalies.Add(INFO1, TraduireMemoire('>>> Arrêt du traitement par l''utilisateur <<<'));
                Break;
            End;

            Try
                // Récupération de l'exercice correspondant à l'année sélectionnée
                Q := OpenSQL('SELECT PEX_EXERCICE,PEX_DATEFIN,PEX_ACTIF,PEX_DEBUTPERIODE,PEX_FINPERIODE FROM '+DBPrefixe+'EXERSOCIAL WHERE PEX_ANNEEREFER="'+IntToStr(AnneeRef)+'"', True);

                If Not Q.EOF Then
                Begin
                    Exercice     := Q.FindField('PEX_EXERCICE').AsString;
                    Actif        := (Q.FindField('PEX_ACTIF').AsString = 'X');
                    PerDebut     := Q.FindField('PEX_DEBUTPERIODE').AsDateTime;
                    PerFin       := Q.FindField('PEX_FINPERIODE').AsDateTime;
                    DateFinExe   := Q.FindField('PEX_DATEFIN').AsDateTime;
                End
                Else
                    Raise EPaieError.Create(Format(TraduireMemoire('L''exercice %s n''existe pas.'), [IntToStr(AnneeRef)]));
                Ferme(Q);
                If V_PGI.SAV Then ListeAnomalies.Add(INFO1, '  Exercice trouvé ['+Exercice+'] pour l''année ['+IntToStr(AnneeRef)+'] : '+BoolToStr(Actif));

                // Si l'exercice n'est pas actif, on sort
                If Not Actif Then
                Begin
                    Raise EPaieError.Create(Format(TraduireMemoire('L''exercice %s n''est pas actif.'),[IntToStr(AnneeRef)]));
                End;

                // Récupération de l'état de clôture actuel
                IndiceClotTemp := RendCloture;

                // Récupération du paramétrage du dossier
                //ModePCL  := ????;
                Q := OpenSQL('SELECT SOC_DATA FROM '+DBPrefixe+'PARAMSOC WHERE SOC_NOM="SO_PGDECALAGE"', True);
                If Not Q.EOF Then
                    DecalagePaie := (Q.FindField('SOC_DATA').AsString = 'X');
                Ferme(Q);
                If V_PGI.SAV Then ListeAnomalies.Add(INFO1, '  Paie décalée : '+BoolToStr(DecalagePaie));

                // Récupération de la période à clôturer
                DonnePeriodeACloturer (AnneeRef, MoisRef, IndiceClotTemp, DD, DF);

                // Le mois précédent doit être clôturé
                If DD > DF then
                Begin
                    DecodeDate(DD,Year,Month,Day);
                    Raise EPaieError.Create(TraduireMemoire('Paie(s) précédente(s) non clôturées. Mois à traiter : ')+RechDOM('PGMOIS', IntToStr(Month), False));
                End;

                // Confirmation du traitement par l'utilisateur uniquement la première fois
                If Reponse = -1 Then
                Begin
                    //St := 'Voulez-vous clôturer les paies du ') + DateToStr(DD) + ' au ' + DateToStr(DF);  //PT10
                    LibMessage := Format(TraduireMemoire('Voulez-vous clôturer les paies du %s au %s ?'),[DateToStr(DD),DateToStr(DF)]);
                    Reponse    := PGIAsk(LibMessage, Ecran.Caption);

                    If Reponse=mrYes Then 
                    Begin
                    	LibMessage := TraduireMemoire('Voulez-vous ouvrir automatiquement la période suivante?');
                    	ReponseOuv := PGIAsk(LibMessage, Ecran.Caption);
                    End;
                End;

                // Lancement de la clôture
                If Reponse = mrYes then
                begin
                    // DEB PT8
                    if (DD <> DebutDeMois (DF)) then
                    begin
                        // Plusieurs mois à traiter : On interdit ce cas en traitement multi-dossier
                        Raise EPaieError.Create(TraduireMemoire('Les paies doivent être clôturées mois par mois.'));
                    end
                    else
                    LanceCloture(DD, DF, IndiceClotTemp); // Clôture d'un seul mois

                    // Ouverture du mois suivant
                    If ReponseOuv = mrYes Then OuvrePeriodeSuivante (DateFinExe, PerDebut, PerFin);
                  // FIN PT8
                End
                Else
                Begin
                    If Q <> Nil Then Ferme(Q);
                    ListeAnomalies.Add(INFO1, '');
                    ListeAnomalies.Add(INFO1, TraduireMemoire('>>> Arrêt du traitement par l''utilisateur <<<'));
                    Break;
                End;

                If Q <> Nil Then Ferme(Q);
            Except
                // Problème de cohérence métier
                On E : EPaieError Do
                Begin
                    ListeAnomalies.Add(INFO1, '  '+TraduireMemoire('Une erreur est survenue durant le traitement.'));
                    ListeAnomalies.Add(ERR1, '- '+Dossier + ' : ' + E.Message);
                    If Q <> Nil Then Ferme(Q);
                End;
            Else
                // Erreur SQL ou de traitement quelconque
                Begin
                    ListeAnomalies.Add(ERR1, '- '+Dossier + ' : ' + TraduireMemoire('Une erreur inconnue est survenue durant le traitement.'));
                    ListeAnomalies.Add(INFO1, TraduireMemoire('  Erreur'));
                    If Q <> Nil Then Ferme(Q);
                End;
            End;
            ListeAnomalies.Add(INFO1, '');
        End;

        // Fermeture de l'écran d'attente
        FiniMoveProgressForm;

        ListeAnomalies.Add(INFO1, '');
        ListeAnomalies.Add(INFO1, TraduireMemoire('>> Traitement terminé <<'));

        If Not ListeAnomalies.IsErrorOccurred Then ListeAnomalies.Add(ERR1,TraduireMemoire('Aucune anomalie n''est survenue.'));
        ListeAnomalies.PutInList(ListeBox);
        FreeAndNil(ListeAnomalies);

        AnneeChange(nil); // Pour reafficher les clôtures éffectuées

        If Assigned (NumDossiers)  Then FreeAndNil(NumDossiers);
        If Assigned (NomsDossiers) Then FreeAndNil(NomsDossiers);
    End;
End;

initialization
  registerclasses([TOF_PG_CLOTURE]);
end.
