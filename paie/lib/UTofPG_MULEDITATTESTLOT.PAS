{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 29/05/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULEDITATTESTLOT ()
Mots clefs ... : TOF;PGMULEDITATTESTLOT
*****************************************************************}
{
PT1   : 10/09/2007 VG V_80 Corrections suite à passage par LanceEtatTob
                           FQ N°14746+FQ N°14754
PT2   : 18/01/2008 GGU V_81 Correction violation d'accès en CWAS
PT4   : 05/08/2008 JPP V_80 FQ 15520 Les deux parenthèses permettent à la requête de fonctionner sous Oracle
PT5   : 10/10/2008 JPP V_81 FQ N°15741 Ajout indemnité spécifique de rupture conventionnelle
}
Unit UTofPG_MULEDITATTESTLOT ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, Hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, 
     EdtREtat,
{$else}
     eMul,
     UtileAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF, utob ;

Type
  TOF_PGMULEDITATTESTLOT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    WW: THEdit; // Clause XX_WHERE
{$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
{$ELSE}
    Liste : THGrid;
{$ENDIF}
    Procedure OnOuvrirClick (Sender : TObject);
    procedure ActiveWhere(Sender: TObject);
  end ;

  Procedure EditionAttestASSEDIC(ListSal : String);

Implementation

uses
  HTB97, StrUtils, HQry, ed_Tools, HStatus, Variants;

procedure TOF_PGMULEDITATTESTLOT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEDITATTESTLOT.ActiveWhere(Sender: TObject);
var
  Annee: string;
begin
  Annee := THValComboBox(GetControl('ANNEESORTIE')).Text;
  if (Annee <> '') and (IsNumeric(Annee)) then
  begin // liste des salaries sortis
    WW.Text:= ' AND PAS_TYPEATTEST = "ASS" '
            + ' AND PSA_DATESORTIE >= "'+UsDateTime(StrToDate('01/01/'+Annee))+'" '
            + ' AND PSA_DATESORTIE <= "'+UsDateTime(StrToDate('31/12/'+Annee))+'" ';
  end else begin
    WW.Text:= ' AND PAS_TYPEATTEST = "ASS" ';
  end;
end;

procedure TOF_PGMULEDITATTESTLOT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEDITATTESTLOT.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEDITATTESTLOT.OnLoad ;
begin
  Inherited ;
  ActiveWhere(Self);
end ;

procedure TOF_PGMULEDITATTESTLOT.OnArgument (S : String ) ;
begin
  Inherited ;
  SetControlText('PSA_SALARIE',Trim(ReadTokenPipe(S, ';')));
  SetControlText('PSA_LIBELLE',Trim(ReadTokenPipe(S, ';')));
  SetControlText('PSA_ETABLISSEMENT',Trim(ReadTokenPipe(S, ';')));
  SetControlText('ANNEESORTIE',Trim(ReadTokenPipe(S, ';')));
{$IFNDEF EAGLCLIENT}
  Liste := (GetControl('FListe') as THDBGrid);
{$ELSE}
  Liste := (GetControl('FListe') as THGrid);
{$ENDIF}
  if Liste <> nil then
  begin
    Liste.OnDblClick:= OnOuvrirClick;
  end;
  (GetControl('BOuvrir') as TToolbarButton97).OnClick := OnOuvrirClick;
//  SetControlText('PAS_TYPEATTEST','ASS');
  WW:= THEdit(GetControl('XX_WHERE'));
  ActiveWhere(Self);
end ;

procedure TOF_PGMULEDITATTESTLOT.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEDITATTESTLOT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEDITATTESTLOT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULEDITATTESTLOT.OnOuvrirClick(Sender: TObject);
var
  Q_Mul: THQuery;
  ListSal : String;
  i : Integer;

  Procedure AddSal(Q : THQuery; var Liste : String);
  var
    Salarie, Ordre : String;
  begin
    Salarie := Q.FindField('PAS_SALARIE').asstring;
    Ordre := Q.FindField('PAS_ORDRE').asstring;
    Liste := Liste + ' OR ((PAS_SALARIE = "'+Salarie+'") AND (PAS_ORDRE = "'+Ordre+'"))'
  end;

begin
//Construction de la liste des salariés/numéro d'ordre sélectionnés
  ListSal := '';
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));
  if Q_Mul = nil then Exit;
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      AddSal(Q_Mul,ListSal);
      MoveCurProgressForm();
    end;
    FiniMoveProgressForm;
  end else if liste.AllSelected then
  begin
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous; //PT3
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
      AddSal(Q_Mul,ListSal);
      Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  end;
  if Length(ListSal) > 0 then
    ListSal := RightStr(ListSal,Length(ListSal) - 4 );
  if ListSal <> '' then //Aucune sélection
    EditionAttestASSEDIC(ListSal)
  else
    PGIError(TraduireMemoire('Vous devez sélectionner au moins une attestation.'),Ecran.Caption);
end;

procedure EditionAttestASSEDIC(ListSal : String);
var
  StSal, StEtab, PGAttesEtab : String;
  QRechSal, QRechEtab : TQuery;
  Salarie, Ordre, Motif : String;
  Rech, SQLTob : String;
  StSQL : String;
  Q : TQuery;
  TobDonnees : Tob;
  i,j : Integer;
  Libelle, Adresse : String;
  Total, Leg, CDD, CNE, FinMiss, Retraite, Sinistre, Spec, Spcfiq, Journal, Client, Avion, Apprenti, Autre : Extended;

  Procedure GetInfoCaisse(Etab, Organisme : String; var Libelle, Adresse : String);
  var
    StCaisse : String;
    QRechCaisse : TQuery; 
  begin
    Libelle := ''; Adresse := '';
    StCaisse:= 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_ADRESSE1, '
             + ' POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, POG_VILLE'
             + ' FROM ORGANISMEPAIE WHERE'
             + ' POG_ETABLISSEMENT = "'+Etab+'" AND POG_ORGANISME="'+Organisme+'"';
    QRechCaisse:=OpenSql(StCaisse,TRUE);
    if (NOT QRechCaisse.EOF) then
    begin //PT2
      Libelle := QRechCaisse.FindField ('POG_LIBELLE').AsString;
      Adresse := Copy(QRechCaisse.FindField('POG_ADRESSE1').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_ADRESSE1').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_ADRESSE2').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_ADRESSE2').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_ADRESSE3').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_ADRESSE3').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_CODEPOSTAL').Asstring))+' '+
                 Copy(QRechCaisse.FindField('POG_VILLE').Asstring, 1,
                      Length(QRechCaisse.FindField('POG_VILLE').Asstring));
    end; //PT2
    Ferme(QRechCaisse);      
  end;

  Function GetIndemnite(TobSearch : Tob; Indemnite : String) : Extended;
  begin
    if (TobSearch.GetString('PAS_IND'+Indemnite+'B') = 'X')  and
       (TobSearch.GetValue('PAS_IND'+Indemnite+'M') <> null) then
      result := TobSearch.GetDouble('PAS_IND'+Indemnite+'M')
    else
      result := 0;
  end;
begin
  //Récupération des données saisies dans la table attestations
  SQLTob := 'SELECT  '
          + 'PAS_SALARIE, PAS_TYPEATTEST, PAS_ORDRE, PAS_TRAVAILTEMP, '
          + 'PAS_DATEATTEST, PAS_ASSEDICCAISSE, PAS_EFFECTIF, PAS_ASSEDICNUM, '
          + 'PAS_PUBLIC, PAS_PUBLICTYPE, PAS_CONVNUMERO, PAS_ADHESION, '
          + 'PAS_CONVCODEANA, PAS_DATEADHESION, PAS_REGIMESPECSS, PAS_CAISSESS, '
          + 'PAS_MATRICULE, PAS_DERNIEREMPLOI, PAS_ALSACEMOSEL, PAS_LIEUTRAVAIL, '
          + 'PAS_DEPART, PAS_QUALIF, PAS_CADRE, PAS_LIENPARENTE, '
          + 'PAS_LIBLIENP, PAS_QUALITEEMPLOI, PAS_LQUALITE, PAS_CONTRATNAT, '
          + 'PAS_CONTRATPARTIC, PAS_LCONTRATPART, PAS_COLLTERRITOR, PAS_FONCTIONSAL, '
          + 'PAS_LIBFONCTION, PAS_AUTREVICTIME, PAS_DATEACCIDENT, PAS_DERNIERJOUR, '
          + 'PAS_DATEARRET, PAS_REPRISEARRET, PAS_MOTIFARRET, PAS_PREAVISMOTIF, '
          + 'PAS_SITUATION, PAS_NONREPRIS, PAS_MAINTIEN, PAS_TYPMAINTIEN, '
          + 'PAS_REPRISEPARTIEL, PAS_CASGEN, PAS_MONTANT, PAS_PLUSDE, '
          + 'PAS_PERIODEDEBUT, PAS_PERIODEFIN, PAS_AUTRERUPT, PAS_MOTRUPCONT, '
          + 'PAS_PLANSOC, PAS_DATEPLANSOC, PAS_RECLASS, PAS_DATEPREAVISD2, '
          + 'PAS_DATEPREAVISF2, PAS_MOTIFPREAVIS2, PAS_DATEPREAVISD3, PAS_DATEPREAVISF3, '
          + 'PAS_MOTIFPREAVIS3, PAS_MOTIFPREAVIS1, PAS_AGESALARIE, PAS_REFUSFNE, '
          + 'PAS_CHOMTOTAL, PAS_DDTEFP, PAS_DEBCHOMAGE, PAS_FINCHOMAGE, '
          + 'PAS_REPRISECHOM, PAS_HORHEBENT, PAS_HORANNENT, PAS_HORHEBSAL, '
          + 'PAS_HORANNSAL, PAS_MOTIFDIFF, PAS_AUTREDIFF, PAS_ORGARRCO, '
          + 'PAS_ORGAGIRC, PAS_ORGAUTRE, PAS_INDPREAVIS, PAS_INDCPMONTANT, '
          + 'PAS_INDCPJOURS, PAS_INDCAISSECP, PAS_INDINHER, PAS_INDLEGB, '
          + 'PAS_INDLEGM, PAS_INDCDDB, PAS_INDCDDM, PAS_INDCNEB, '
          + 'PAS_INDCNEM, PAS_INDAUTREB, PAS_INDAUTREM, PAS_INDFINMISSB, '
          + 'PAS_INDFINMISSM, PAS_INDRETRAITEB, PAS_INDRETRAITEM, PAS_INDSINISTREB, '
          + 'PAS_INDSINISTREM, PAS_INDSPECB, PAS_INDSPECM, PAS_INDSPCFIQB, '
          + 'PAS_INDSPCFIQM, PAS_INDJOURNALB, PAS_INDJOURNALM, PAS_INDCPEB, '
          + 'PAS_INDCPEM, PAS_INDCLIENTB, PAS_INDCLIENTM, PAS_INDAPPRENTIB, '
          + 'PAS_INDAPPRENTIM, PAS_INDAVIONB, PAS_INDAVIONM, PAS_INDCONV2, '
          + 'PAS_INDTRANS2, PAS_INDTRANSC, PAS_FNGSRED, PAS_FNGS, '
          + 'PAS_FNGSCRE1, PAS_FNGSMOT1, PAS_FNGS2, PAS_FNGSCRE2, '
          + 'PAS_FNGSMOT2, PAS_DECLARNOM, PAS_DECLARPRENOM, PAS_DECLARQUAL, '
          + 'PAS_DECLARQAUTRE, PAS_DECLARRENS1, PAS_DECLARDATE, PAS_DECLARPERS, '
          + 'PAS_DECLARTEL, PAS_DECLARLIEU, PAS_SUBDEBUT, PAS_SUBFIN, '
          + 'PAS_SUBINTEGRAL, PAS_SUBCOMPTE, PAS_SUBCPTEINT, PAS_SUBMONNAIE, '
          + 'PAS_TYPEABS, '
           + ' PSA_SALARIE, PSA_CIVILITE, PSA_NOMJF, PSA_ADRESSE1,'
           + ' PSA_ADRESSE2, PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE,'
           + ' PSA_DATENAISSANCE, '
           + ' 0 AS INDSPECONVENTIONM ' //PT5
          + 'FROM ATTESTATIONS '
           + ' LEFT JOIN SALARIES ON PAS_SALARIE=PSA_SALARIE '
          + 'WHERE ('+ListSal+')'; // PT4
  TobDonnees := Tob.Create('Les données',nil,-1);
  TobDonnees.LoadDetailFromSQL(SQLTob);
  //Récupération des données de salaires, primes et soldes depuis la table  ATTSALAIRES
  for i := 0 to TobDonnees.FillesCount(0)-1 do
  begin
    //On ajouter les lignes de titre
    With TobDonnees.Detail[i] do
    begin
      AddChampSupValeur('GRID_SALAIRES_0_0_', 'Période du');
      AddChampSupValeur('GRID_SALAIRES_1_0_', 'Au');
      AddChampSupValeur('GRID_SALAIRES_2_0_', 'Date paie');
      AddChampSupValeur('GRID_SALAIRES_3_0_', 'Heures trav.');
      AddChampSupValeur('GRID_SALAIRES_4_0_', 'J. non payés');
      AddChampSupValeur('GRID_SALAIRES_5_0_', 'Salaire brut');
      AddChampSupValeur('GRID_SALAIRES_6_0_', 'Précompte');
      AddChampSupValeur('GRID_SALAIRES_7_0_', 'Observations');
      AddChampSupValeur('GRID_PRIMES_0_0_', 'Période du');
      AddChampSupValeur('GRID_PRIMES_1_0_', 'Au');
      AddChampSupValeur('GRID_PRIMES_2_0_', 'Nature Prime');
      AddChampSupValeur('GRID_PRIMES_3_0_', 'Date Paiement');
      AddChampSupValeur('GRID_PRIMES_4_0_', 'Montant');
      AddChampSupValeur('GRID_SOLDE_0_0_', 'Période du');
      AddChampSupValeur('GRID_SOLDE_1_0_', 'Au');
      AddChampSupValeur('GRID_SOLDE_2_0_', 'Date Paiement');
      AddChampSupValeur('GRID_SOLDE_3_0_', 'Nb heures trav.');
      AddChampSupValeur('GRID_SOLDE_4_0_', 'Salaire brut');
      AddChampSupValeur('GRID_SOLDE_5_0_', 'Précompte');
      Salarie := GetString('PAS_SALARIE');
      Ordre   := GetString('PAS_ORDRE');
    end;
    StSQL:= 'SELECT PAL_DATEDEBUT, PAL_DATEFIN, PAL_PAYELE, PAL_NBHEURES,'
          + ' PAL_JNONPAYES, PAL_SALAIRE, PAL_PRECOMPTE, PAL_MONNAIE,'
          + ' PAL_OBSERVATIONS, PAL_MOIS'
          + ' FROM ATTSALAIRES '
          + ' WHERE PAL_ORDRE='+Ordre
          + ' AND PAL_SALARIE="'+Salarie+'" '
          + ' AND PAL_TYPEATTEST="ASS"'
          + ' ORDER BY PAL_MOIS';
    Q:=OpenSql(StSQL,TRUE);
    if not Q.Eof then
    begin
      for j := 1 to 12 do
      begin
        With TobDonnees.Detail[i] do
        begin
          AddChampSupValeur('GRID_SALAIRES_0_'+IntToStr(j)+'_', Q.FindField('PAL_DATEDEBUT').Asstring           );
          AddChampSupValeur('GRID_SALAIRES_1_'+IntToStr(j)+'_', Q.FindField('PAL_DATEFIN').Asstring             );
          AddChampSupValeur('GRID_SALAIRES_2_'+IntToStr(j)+'_', Q.FindField('PAL_PAYELE').Asstring              );
          AddChampSupValeur('GRID_SALAIRES_3_'+IntToStr(j)+'_', FloatToStr(Q.FindField('PAL_NBHEURES').AsFloat) );
          AddChampSupValeur('GRID_SALAIRES_4_'+IntToStr(j)+'_', FloatToStr(Q.FindField('PAL_JNONPAYES').AsFloat));
          AddChampSupValeur('GRID_SALAIRES_5_'+IntToStr(j)+'_', FloatToStr(Q.FindField('PAL_SALAIRE').AsFloat)  );
          AddChampSupValeur('GRID_SALAIRES_6_'+IntToStr(j)+'_', FloatToStr(Q.FindField('PAL_PRECOMPTE').AsFloat));
          AddChampSupValeur('GRID_SALAIRES_7_'+IntToStr(j)+'_', Q.FindField('PAL_OBSERVATIONS').Asstring        );
        end;
        Q.Next;
      end;
      for j := 13 to 15 do
      begin
        With TobDonnees.Detail[i] do
        begin
          AddChampSupValeur('GRID_PRIMES_0_'+IntToStr(j-12)+'_', Q.FindField('PAL_DATEDEBUT').Asstring          );
          AddChampSupValeur('GRID_PRIMES_1_'+IntToStr(j-12)+'_', Q.FindField('PAL_DATEFIN').Asstring            );
          AddChampSupValeur('GRID_PRIMES_2_'+IntToStr(j-12)+'_', Q.FindField('PAL_OBSERVATIONS').Asstring       );
          AddChampSupValeur('GRID_PRIMES_3_'+IntToStr(j-12)+'_', Q.FindField('PAL_PAYELE').Asstring             );
          AddChampSupValeur('GRID_PRIMES_4_'+IntToStr(j-12)+'_', FloatToStr(Q.FindField('PAL_SALAIRE').AsFloat) );
        end;
        Q.Next;
      end;
      for j := 16 to 17 do
      begin
        With TobDonnees.Detail[i] do
        begin
          AddChampSupValeur('GRID_SOLDE_0_'+IntToStr(j-15)+'_', Q.FindField('PAL_DATEDEBUT').Asstring           );
          AddChampSupValeur('GRID_SOLDE_1_'+IntToStr(j-15)+'_', Q.FindField('PAL_DATEFIN').Asstring             );
          AddChampSupValeur('GRID_SOLDE_2_'+IntToStr(j-15)+'_', Q.FindField('PAL_PAYELE').Asstring              );
          AddChampSupValeur('GRID_SOLDE_3_'+IntToStr(j-15)+'_', FloatToStr(Q.FindField('PAL_NBHEURES').AsFloat) );
          AddChampSupValeur('GRID_SOLDE_4_'+IntToStr(j-15)+'_', FloatToStr(Q.FindField('PAL_SALAIRE').AsFloat)  );
          AddChampSupValeur('GRID_SOLDE_5_'+IntToStr(j-15)+'_', FloatToStr(Q.FindField('PAL_PRECOMPTE').AsFloat));
        end;
        Q.Next;
      end;
    end;
    Ferme(Q);
    //On renseigne les champs complémentaires
    With TobDonnees.Detail[i] do
    begin
      AddChampSupValeur('LABEL_QUALIF', RechDom ('PGQUALIF',GetString('PAS_QUALIF'),FALSE,''));
      Motif := GetString('PAS_MOTRUPCONT');
      if (Motif='403') then Motif:= '';
      if (Motif='90') then Motif:= '60';
      AddChampSupValeur('LABEL_MOTIFLIB', RechDom ('PGMOTIFSORTIE',Motif,FALSE,''));
    end;
    //Infos Salarié
    StSal:='SELECT PSA_LIBELLE, PSA_PRENOM, PSA_ETABLISSEMENT, PSA_NUMEROSS,'+
           ' PSA_LIBELLEEMPLOI, PSA_STATUT, PSA_DATEENTREE, PSA_DATESORTIE,'+
           ' PSA_HORHEBDO, PSA_HORANNUEL, PSA_MOTIFSORTIE, PSA_CATDADS'+
           ' FROM SALARIES WHERE'+
           ' PSA_SALARIE="'+Salarie+'"';
    QRechSal:=OpenSql(StSal,TRUE);
    if not QRechSal.Eof then
    begin
      With TobDonnees.Detail[i] do
      begin
        AddChampSupValeur('EDIT_NOM',    QRechSal.FindField('PSA_LIBELLE').Asstring );
        AddChampSupValeur('EDIT_PRENOM', QRechSal.FindField('PSA_PRENOM').Asstring  );
        AddChampSupValeur('EDIT_NUMSS',  QRechSal.FindField('PSA_NUMEROSS').Asstring );
      end;
    end;
    //Infos établissement
    PGAttesEtab := QRechSal.FindField('PSA_ETABLISSEMENT').Asstring;
    StEtab:='SELECT ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3,'+
            ' ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_JURIDIQUE, ET_SIRET, ET_APE'+
            ' FROM ETABLISS WHERE'+
            ' ET_ETABLISSEMENT = "'+PGAttesEtab+'"';
    QRechEtab:=OpenSql(StEtab,TRUE);
    if not QRechEtab.EOF then
    begin
      With TobDonnees.Detail[i] do
      begin
        AddChampSupValeur('EDIT_ETABNOM',  QRechEtab.FindField('ET_LIBELLE').Asstring);
        AddChampSupValeur('EDIT_ETABADR1', QRechEtab.FindField('ET_ADRESSE1').Asstring);
        AddChampSupValeur('EDIT_ETABADR2', QRechEtab.FindField('ET_ADRESSE2').Asstring);
        AddChampSupValeur('EDIT_ETABADR3', QRechEtab.FindField('ET_ADRESSE3').Asstring);
        AddChampSupValeur('EDIT_ETABCP',   QRechEtab.FindField('ET_CODEPOSTAL').Asstring);
        AddChampSupValeur('EDIT_ETABVILLE',QRechEtab.FindField('ET_VILLE').Asstring);
        AddChampSupValeur('EDIT_ETABSTAT', RechDom('TTFORMEJURIDIQUE', QRechEtab.FindField('ET_JURIDIQUE').Asstring, False));
        AddChampSupValeur('EDIT_ETABTEL',  QRechEtab.FindField('ET_TELEPHONE').Asstring);
        AddChampSupValeur('EDIT_ETABSIRET',QRechEtab.FindField('ET_SIRET').Asstring);
        AddChampSupValeur('EDIT_ETABAPE',  QRechEtab.FindField('ET_APE').Asstring);
      end;
    end;
    Ferme(QRechEtab);
    //ARRCO
    GetInfoCaisse(PGAttesEtab, TobDonnees.Detail[i].GetString('PAS_ORGARRCO'), Libelle, Adresse);
    TobDonnees.Detail[i].AddChampSupValeur('LABEL_ARRCO', Libelle);
    TobDonnees.Detail[i].AddChampSupValeur('LABEL_COADR1', Adresse);
    //AGIRC
    GetInfoCaisse(PGAttesEtab, TobDonnees.Detail[i].GetString('PAS_ORGAGIRC'), Libelle, Adresse);
    TobDonnees.Detail[i].AddChampSupValeur('LABEL_AGIRC', Libelle);
    TobDonnees.Detail[i].AddChampSupValeur('LABEL_RCADR1', Adresse);
    //Autre caisse de retraite complémentaire
    GetInfoCaisse(PGAttesEtab, TobDonnees.Detail[i].GetString('PAS_ORGAUTRE'), Libelle, Adresse);
    TobDonnees.Detail[i].AddChampSupValeur('LABEL_AUTRES', Libelle);
    TobDonnees.Detail[i].AddChampSupValeur('LABEL_AUADR1', Adresse);
    //Total des indémnités
    Leg      := GetIndemnite(TobDonnees.Detail[i], 'LEG');
    CDD      := GetIndemnite(TobDonnees.Detail[i], 'CDD');
    CNE      := GetIndemnite(TobDonnees.Detail[i], 'CNE');
    FinMiss  := GetIndemnite(TobDonnees.Detail[i], 'FINMISS');
    Retraite := GetIndemnite(TobDonnees.Detail[i], 'RETRAITE');
    Sinistre := GetIndemnite(TobDonnees.Detail[i], 'SINISTRE');
    Spec     := GetIndemnite(TobDonnees.Detail[i], 'SPEC');
    Spcfiq   := GetIndemnite(TobDonnees.Detail[i], 'SPCFIQ');
    Journal  := GetIndemnite(TobDonnees.Detail[i], 'JOURNAL');
    Client   := GetIndemnite(TobDonnees.Detail[i], 'CLIENT');
    Avion    := GetIndemnite(TobDonnees.Detail[i], 'AVION');
    Apprenti := GetIndemnite(TobDonnees.Detail[i], 'APPRENTI');
    Autre    := GetIndemnite(TobDonnees.Detail[i], 'AUTRE');
    Total:= Leg+CDD+CNE+FinMiss+Retraite+Sinistre+Spec+Spcfiq+Journal+Client+
            Avion+Apprenti+Autre;
    TobDonnees.Detail[i].AddChampSupValeur('EDIT_INDTOT', Total);

    //On remplace la valeur par le libellé
    TobDonnees.detail[i].SetString ('PAS_DERNIEREMPLOI', RechDom ('PGLIBEMPLOI',
                                                                  TobDonnees.detail[i].GetString ('PAS_DERNIEREMPLOI'),
                                                                  False));
    TobDonnees.detail[i].SetString ('PAS_PUBLICTYPE', RechDom ('PGPUBLICTYPE',
                                                               TobDonnees.detail[i].GetString ('PAS_PUBLICTYPE'),
                                                               False));
    TobDonnees.detail[i].SetString ('PAS_COLLTERRITOR', RechDom ('PGCOLLTERR',
                                                                 TobDonnees.detail[i].GetString ('PAS_COLLTERRITOR'),
                                                                 False));
//PT1
    TobDonnees.detail[i].SetString ('PAS_DEPART', RechDom ('JUDEPART',
                                                           TobDonnees.detail[i].GetString ('PAS_DEPART'),
                                                           False));
    TobDonnees.detail[i].SetString ('PAS_QUALITEEMPLOI',
                                    RechDom ('PGQUALITEEMPLOI',
                                             TobDonnees.detail[i].GetString ('PAS_QUALITEEMPLOI'),
                                             False));
    TobDonnees.detail[i].SetString ('PAS_MOTIFDIFF', RechDom ('PGMOTIFDIFF',
                                                              TobDonnees.detail[i].GetString ('PAS_MOTIFDIFF'),
                                                              False));
    TobDonnees.detail[i].SetString ('PAS_CONTRATNAT',
                                    RechDom ('PGCONTRATNATURE',
                                             TobDonnees.detail[i].GetString ('PAS_CONTRATNAT'),
                                             False));
    TobDonnees.detail[i].SetString ('PAS_CONTRATPARTIC',
                                    RechDom ('PGCONTRATPART',
                                             TobDonnees.detail[i].GetString ('PAS_CONTRATPARTIC'),
                                             False));
    TobDonnees.detail[i].SetString ('PAS_FONCTIONSAL', RechDom ('PGSTATUT',
                                                                TobDonnees.detail[i].GetString ('PAS_FONCTIONSAL'),
                                                                False));
//FIN PT1

  end;
  Rech := 'SELECT PAS_SALARIE, PSA_CIVILITE, PSA_NOMJF, PSA_ADRESSE1,'+
          ' PSA_ADRESSE2, PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE,'+
          ' PSA_DATENAISSANCE'+
          ' FROM ATTESTATIONS'+
          ' LEFT JOIN SALARIES ON'+
          ' PAS_SALARIE=PSA_SALARIE WHERE '+
          ListSal + ' AND PAS_TYPEATTEST="ASS"';

  LanceEtatTob('E','PAT','ASS',TobDonnees,True,False,False,nil,Rech,'',False);
  FreeAndNil(TobDonnees);
end;

Initialization
  registerclasses ( [ TOF_PGMULEDITATTESTLOT ] ) ;
end.
