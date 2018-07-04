{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CPREVENTILANA ()
Mots clefs ... : TOF;CPREVENTILANA
*****************************************************************}
unit Reventilana;

interface

uses StdCtrls, Controls, Classes, db, forms, sysutils,
  ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB, HTB97, HDB, ULibAnalytique {SG6 17.02.05},
  ParamSoc,
  {$IFDEF EAGLCLIENT}
  eMul,
  MaineAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Mul,
  FE_Main,
  {$ENDIF}
  AGLInit, AGLInitPlus,
  HQry, Ent1, SaisBor, SaisUtil;

procedure CPLanceFiche_ReventilAna;

type
  TOF_CPREVENTILANA = class(TOF)
    BOuvrir: TToolBarButton97;
    BCherche: TToolBarButton97;
    Exercice: THValComboBox;
    Journal: THValComboBox;
    General: THValComboBox;
    General_: THValComboBox;
    DateComptable: THvalComboBox;
    DateComptable_: THvalComboBox;
    ChercheVentil: TToolBarButton97;
    LAxe: THValComboBox;
    LaVentil: THValComboBox;
    ListeEcr: THDBGrid;
    FListe: THDbGrid;
    QlisteEcr: TQuery;
    TobVentil: TOB;
    TobAna: TOB;
    TobDev: TOB;
    TobSection: TOB;
{b fb 16/06/2006 FQ17912}
    ZC1 : THValComboBox;
    ZO1 : THValComboBox;
    ZV1 : THEdit;
    ZC2 : THValComboBox;
    ZO2 : THValComboBox;
    ZV2 : THEdit;
{e fb 16/06/2006 FQ17912}
    procedure OnArgument(S: string); override;
    procedure AfterShow;
  private
    //SG6 17.02.05 mode ana croisaxe
    iAxe: integer;

    function VeriAxeSection(ValAxe: string): Boolean;
    function VerifModelRestriction(i: Integer): Boolean;          {FP 29/12/2005}
    function RechDecDev(TobDetail: TOB): Integer;
    procedure BOuvrirOnClick(Sender: TObject);
    procedure PasChoixSaisi;
    procedure RecupAna;
    procedure RecupEcritAna;
    procedure MarqueEcrAna;
    procedure VireInutiles;
    procedure ChargeTobDevise;
    procedure ChargeTobSection;
    procedure DetruitLigneAna;
    procedure MAJTobSection(TobAnaDet: TOB; Signe: Integer);
    procedure ModifLigne(TobAnaDet: TOB);
    procedure MAJZoneMontant(var TobAjout: TOB; M, MD: Double; Debit: Boolean);
    procedure LigneFinal(var TobAnaAjout: TOB; S, SD, Taux: Double; NumVentil: Integer; Debit: Boolean; DecDev: Integer);
    procedure MAJSection;
    procedure LibereTobs;
    procedure FListeDblClick(Sender: TObject);
    procedure ExerciceOnChange(Sender: TObject);
    procedure ChercheVentilOnClick(Sender: TObject);
    function ChargeTobVentil(i: integer): Boolean;
    procedure VentilParDefautOnClick(Sender: TObject);
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  Saisie;

var
  Titre: string;
  Dec, Dec2: Integer;

procedure CPLanceFiche_ReventilAna;
begin
  AGLLanceFiche('CP', 'CPREVENTILANA', '', '', '');
end;

procedure TOF_CPREVENTILANA.PasChoixSaisi;
begin
  PGIBox(TraduireMemoire('La ventilation sélectionnée n''appartient pas à cet axe.'), TraduireMemoire(Titre));
  //ActiveBouton
  Exit;
end;

function TOF_CPREVENTILANA.VeriAxeSection(ValAxe: string): Boolean;
var
  Q: TQuery;
begin
  Result := False;
  if GetCheckBoxState('VentilParDefaut') = CbChecked then exit;
  Result := True;
  Q := OpenSql('SELECT V_SOUSPLAN1, V_SOUSPLAN2, V_SOUSPLAN3, V_SOUSPLAN4, V_SOUSPLAN5, V_NATURE,V_COMPTE,V_NUMEROVENTIL, V_MONTANT, '
    + 'V_SECTION,V_TAUXMONTANT,V_TAUXQTE1,V_TAUXQTE2 FROM VENTIL WHERE '
    + 'V_NATURE ="' + ValAxe + '" AND V_COMPTE="' + LaVentil.Value + '"', True);
  if Q.EOF then
  begin
    Result := False;
    Exit;
  end;
  TobVentil := Tob.Create('_VENTIL', nil, -1);
  TobVentil.LoadDetailDb('VENTIL', '', '', Q, False, False);
  Ferme(Q);
end;

procedure TOF_CPREVENTILANA.RecupAna;
var
  i: Integer;
  St1, St2, RSql: string;
  Q: TQuery;
begin // Modif du recupwherecritere portant sur la table ECRITURE pour l'adapter à la table Analytique
  St1 := AnsiUpperCase(RecupWhereCritere(TPageControl(GetControl('PAGES'))));
  St2 := 'E_ANA="X"';
  i := Pos(St2, St1);
  if i > 0 then
  begin
    System.Delete(St1, i, Length(St2));
    System.Insert('Y_TYPEANALYTIQUE="-"', St1, i);
  end;
  St2 := 'E_';
  i := Pos(St2, St1);
  while i > 0 do
  begin
    System.Delete(St1, i, Length(St2));
    System.Insert('Y_', St1, i);
    i := Pos(St2, St1);
  end;
  RSql := 'SELECT Y_JOURNAL, Y_GENERAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_DEVISE, '
    + 'Y_NUMLIGNE, Y_AXE, Y_SECTION, Y_NUMVENTIL, Y_QUALIFPIECE, Y_CREDIT, Y_CREDITDEV, Y_TOTALDEVISE, '
    + 'Y_DEBIT, Y_DEBITDEV, Y_NATUREPIECE, Y_POURCENTAGE, Y_TOTALECRITURE, '
    + 'Y_LIBELLE, Y_REFINTERNE, Y_REFEXTERNE, Y_AFFAIRE, Y_PERIODE, Y_SEMAINE, '
    + 'Y_ETAT, Y_SOCIETE, Y_ETABLISSEMENT, Y_TAUXDEV, Y_DATETAUXDEV, Y_QUALIFQTE1, Y_QUALIFQTE2, Y_TYPEMVT, '
    + 'Y_ECRANOUVEAU, Y_CONFIDENTIEL, Y_CREERPAR, Y_EXPORTE, Y_CONTREPARTIEGEN, Y_CONTREPARTIEAUX  '
    // Gestion QTE // MODIF FICHE 12570
  + ', Y_QTE1, Y_QTE2, Y_POURCENTQTE1, Y_POURCENTQTE2' // données qtés lignes analytiques
  + ', Y_QUALIFECRQTE1, Y_QUALIFECRQTE2, Y_TOTALQTE1, Y_TOTALQTE2' // données qtés ligne générale
  + ' FROM ANALYTIQ '
    // FIN MODIF FICHE 12570
  + St1 + ' AND Y_AXE="A' + IntToStr(iAxe) + '" ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_QUALIFPIECE';
  Q := OpenSql(Rsql, True);
  TobAna := TOB.Create('_ANALYTIQ', nil, -1);
  if not Q.Eof then
  begin
    TobAna.LoadDetailDB('ANALYTIQ', '', '', Q, True, True);
    if TobAna.Detail.Count > 0 then TobAna.Detail[0].AddChampSup('MARQUE', True);
  end;
  Ferme(Q);
end;

procedure TOF_CPREVENTILANA.MarqueEcrAna;
var
  TobAnaDetail: TOB;
begin
  TobAnaDetail := TobAna.FindFirst(['Y_JOURNAL', 'Y_EXERCICE', 'Y_DATECOMPTABLE', 'Y_NUMEROPIECE', 'Y_NUMLIGNE', 'Y_QUALIFPIECE'],
    [QListeEcr.FindField('E_JOURNAL').AsString,
    QListeEcr.FindField('E_EXERCICE').AsString,
      QListeEcr.FindField('E_DATECOMPTABLE').AsDateTime,
      QListeEcr.FindField('E_NUMEROPIECE').AsInteger,
      QListeEcr.FindField('E_NUMLIGNE').AsInteger,
      QListeEcr.FindField('E_QUALIFPIECE').AsString], False);
  while TobAnaDetail <> nil do
  begin
    TobAnaDetail.PutValue('MARQUE', 'X');
    TobAnaDetail := TobAna.FindNext(['Y_JOURNAL', 'Y_EXERCICE', 'Y_DATECOMPTABLE', 'Y_NUMEROPIECE', 'Y_NUMLIGNE', 'Y_QUALIFPIECE'],
      [QListeEcr.FindField('E_JOURNAL').AsString,
      QListeEcr.FindField('E_EXERCICE').AsString,
        QListeEcr.FindField('E_DATECOMPTABLE').AsDateTime,
        QListeEcr.FindField('E_NUMEROPIECE').AsInteger,
        QListeEcr.FindField('E_NUMLIGNE').AsInteger,
        QListeEcr.FindField('E_QUALIFPIECE').AsString], False);
  end;
end;


procedure TOF_CPREVENTILANA.VireInutiles;
var
  i: integer;
begin
  for i := TobAna.Detail.Count - 1 downto 0 do
  begin
    if TobAna.Detail[i].GetValue('MARQUE') <> 'X' then TobAna.Detail[i].Free;
  end;
end;

procedure TOF_CPREVENTILANA.RecupEcritAna;
var
  i: integer;
  Fiche: TFMul;
begin
  Fiche := TFMul(Ecran);
  ListeEcr := Fiche.FListe;
  {$IFDEF EAGLCLIENT}
  QListeEcr := Fiche.Q.TQ;
  {$ELSE}
  QListeEcr := Fiche.Q;
  {$ENDIF}
  RecupAna; // Recupère les écritures analytiques correspondantes
  if not ListeEcr.AllSelected then
  begin
    for i := 0 to ListeEcr.NbSelected - 1 do
    begin
      ListeEcr.GotoLeBookmark(i);
      MarqueEcrAna; // Marque les lignes selectionnées
    end;
    VireInutiles; // Enlève les lignes non selectionnées
  end;
end;

procedure TOF_CPREVENTILANA.ChargeTobSection;
var
  Q: TQuery;
begin
  TobSection := Tob.Create('_SECTION', nil, -1);
  Q := OpenSql('SELECT S_AXE,S_SECTION,S_TOTALCREDIT,S_TOTALDEBIT,S_TOTCREP,S_TOTCREE,S_TOTCRES,S_TOTDEBE,S_TOTDEBP,S_TOTDEBS FROM SECTION', true);
  if not Q.EOF then TobSection.LoadDetailDb('SECTION', '', '', Q, True, True);
  if TobSection.Detail.Count > 0 then TobSection.Detail[0].AddChampSup('MOUVEMENTE', True); //Mouvemente X=oui pour MàJ
  Ferme(Q);
end;

procedure TOF_CPREVENTILANA.ChargeTobDevise;
var
  Q: TQuery;
begin
  TobDev := Tob.Create('_DEVISE', nil, -1);
  Q := OpenSql('SELECT D_DEVISE, D_DECIMALE FROM DEVISE', true);
  TobDev.LoadDetailDB('DEVISE', '', '', Q, True, True);
  Ferme(Q);
end;

procedure TOF_CPREVENTILANA.DetruitLigneAna;
var
  i: integer;
begin
  for i := TobAna.Detail.Count - 1 downto 0 do
  begin
    MAJTobSection(TobAna.Detail[i], -1);
    if TobAna.Detail[i].GetValue('Y_NUMVENTIL') > 1 then
    begin
      TobAna.Detail[i].DeleteDB;
      TobAna.Detail[i].Free;
    end else
    begin
      TobAna.Detail[i].DeleteDB;
    end;
  end;
end;

procedure TOF_CPREVENTILANA.MAJTobSection(TobAnaDet: TOB; Signe: Integer);
var
  TobSectionDet: TOB;
begin
  TobSectionDet := TobSection.Findfirst(['S_AXE', 'S_SECTION'],
    [TobAnaDet.GetValue('Y_AXE'), TobAnaDet.GetValue('Y_SECTION')], False);
  if TobSectionDet <> nil then
  begin
    TobSectionDet.PutValue('S_TOTALDEBIT', Arrondi(TobSectionDet.GetValue('S_TOTALDEBIT') + (Signe * TobAnaDet.GetValue('Y_DEBIT')), dec));
    TobSectionDet.PutValue('S_TOTALCREDIT', Arrondi(TobSectionDet.GetValue('S_TOTALCREDIT') + (Signe * TobAnaDet.GetValue('Y_CREDIT')), dec));
    TobSectionDet.PutValue('MOUVEMENTE', 'X');
    if TobAnaDet.GetValue('Y_EXERCICE') = VH^.EnCours.Code then
    begin
      TobSectionDet.PutValue('S_TOTDEBE', Arrondi(TobSectionDet.GetValue('S_TOTDEBE') + (Signe * TobAnaDet.GetValue('Y_DEBIT')), dec));
      TobSectionDet.PutValue('S_TOTCREE', Arrondi(TobSectionDet.GetValue('S_TOTCREE') + (Signe * TobAnaDet.GetValue('Y_CREDIT')), dec));
    end else
    begin
      TobSectionDet.PutValue('S_TOTDEBS', Arrondi(TobSectionDet.GetValue('S_TOTDEBS') + (Signe * TobAnaDet.GetValue('Y_DEBIT')), dec));
      TobSectionDet.PutValue('S_TOTCRES', Arrondi(TobSectionDet.GetValue('S_TOTCRES') + (Signe * TobAnaDet.GetValue('Y_CREDIT')), dec));
    end;
  end;
end;

function TOF_CPREVENTILANA.RechDecDev(TobDetail: TOB): Integer;
var
  TobDecDev: TOB;
begin
  TobDecDev := TobDev.Findfirst(['D_DEVISE'], [TobDetail.GetValue('Y_DEVISE')], False);
  if TobDecDev <> nil then Result := TobDecDev.getValue('D_DECIMALE')
  else Result := 2;
end;

procedure TOF_CPREVENTILANA.MAJZoneMontant(var TobAjout: TOB; M, MD: Double; Debit: Boolean);
var
  Zone: string;
begin
  if Debit then Zone := 'DEBIT' else Zone := 'CREDIT';
  TobAjout.PutValue('Y_' + Zone, M);
  TobAjout.PutValue('Y_' + Zone + 'DEV', MD);
end;

procedure TOF_CPREVENTILANA.LigneFinal(var TobAnaAjout: TOB; S, SD, Taux: Double; NumVentil: Integer; Debit: Boolean; DecDev: Integer);
var
  M, MD: Double;
  i : integer;
begin
  TobAnaAjout.PutValue('Y_NUMVENTIL', NumVentil);
  TobAnaAjout.PutValue('Y_AXE', 'A' + IntToStr(iAxe));
  TobAnaAjout.PutValue('Y_SECTION', VH^.Cpta[AxeToFb('A' + IntToStr(iAxe))].Attente);

  //SG6 17.02.05 mode ana croisaxe
  if VH^.AnaCroisaxe then
  begin
    for i := 1 to 5 do
    begin
      if GetParamSocSecur('SO_VENTILA' + IntToStr(i),False) then
         TobAnaAjout.PutValue('Y_SOUSPLAN' + IntToStr(i), VH^.Cpta[AxeToFb('A' + IntToStr(i))].Attente);
    end;
  end;

  TobAnaAjout.PutValue('Y_POURCENTAGE', Taux);
  M := Arrondi(TobAnaAjout.GetValue('Y_TOTALECRITURE') - S, Dec);
  MD := Arrondi(TobAnaAjout.GetValue('Y_TOTALDEVISE') - SD, DecDev);
  MAJZoneMontant(TobAnaAjout, M, MD, Debit);
  // Gestion QTE // MODIF FICHE 12570
  TobAnaAjout.PutValue('Y_QUALIFQTE1', TobAnaAjout.GetValue('Y_QUALIFECRQTE1'));
  TobAnaAjout.PutValue('Y_POURCENTQTE1', 0);
  TobAnaAjout.PutValue('Y_QTE1', 0);
  TobAnaAjout.PutValue('Y_QUALIFQTE2', TobAnaAjout.GetValue('Y_QUALIFECRQTE2'));
  TobAnaAjout.PutValue('Y_POURCENTQTE2', 0);
  TobAnaAjout.PutValue('Y_QTE2', 0);
  // FIN MODIF FICHE 12570
end;

procedure TOF_CPREVENTILANA.ModifLigne(TobAnaDet: TOB); // Creation des nouvelles lignes de ventilation analytique
var
  i, j, z: Integer;
  TobAnaAjout: TOB;
  Zone, SWhere: string;
  Q : TQuery;
  Debit, ModifDerniereLigne: Boolean;
  Solde, SoldeDev, Montant, MontantDev, TotTaux: Double;
  DecDev: Integer;
  TotalQte1, TotalQte2, qte1, qte2: Double; // Gestion QTE // MODIF FICHE 12570
  TotalTauxQte1, TotalTauxQte2: Double; // Gestion QTE // MODIF FICHE 12570
  GereResteQte1, GereQte1: Boolean; // Gestion QTE // MODIF FICHE 12570
  GereResteQte2, GereQte2: Boolean; // Gestion QTE // MODIF FICHE 12570
  //    x,y : integer ;//gv
begin
  Solde := 0;
  SoldeDev := 0;
  TotTaux := 0;
  DecDev := RechDecDev(TobAnaDet);
  j := 0;

  SWhere:='E_JOURNAL="'+TobAnaDet.GetString('Y_JOURNAL')
  + '" AND E_EXERCICE="'+TobAnaDet.GetString('Y_EXERCICE')
  + '" AND E_DATECOMPTABLE="'+USDateTime(TobAnaDet.GetDateTime('Y_DATECOMPTABLE'))
  + '" AND E_NUMEROPIECE='+TobAnaDet.GetString('Y_NUMEROPIECE')
  + '  AND E_QUALIFPIECE="'+TobAnaDet.GetString('Y_QUALIFPIECE')
  + '" AND E_NUMLIGNE='+TobAnaDet.GetString('Y_NUMLIGNE');

  {FQ22200 23.01.08 YMO Test [Debit := Abs(TobAnaDet.GetValue('Y_DEBIT')) > 0;] erroné}
  Q := OpenSQL('SELECT E_DEBIT FROM ECRITURE WHERE '+SWhere, FALSE);
  if not Q.EOF then
    Debit := Q.FindField('E_DEBIT').AsFloat > 0;   

  Ferme(Q);

  TobAnaAjout := Tob.Create('Analytiq', nil, -1);

  // Gestion QTE // MODIF FICHE 12570
  TotalTauxQte1 := 0;
  TotalTauxQte2 := 0;
  for i := 0 to TobVentil.Detail.Count - 1 do
  begin
    TotalTauxQte1 := TotalTauxQte1 + TobVentil.Detail[i].GetValue('V_TAUXQTE1');
    TotalTauxQte2 := TotalTauxQte2 + TobVentil.Detail[i].GetValue('V_TAUXQTE2');
  end;
  GereQte1 := (TotalTauxQte1 > 0) and (TobAnaDet.GetValue('Y_TOTALQTE1') <> 0);
  GereQte2 := (TotalTauxQte2 > 0) and (TobAnaDet.GetValue('Y_TOTALQTE2') <> 0);
  GereResteQte1 := Arrondi(TotalTauxQte1 - 100.0, 3) = 0;
  GereResteQte2 := Arrondi(TotalTauxQte2 - 100.0, 3) = 0;
  TotalQte1 := 0;
  TotalQte2 := 0;
  // FIN MODIF FICHE 12570

  for i := 0 to TobVentil.Detail.Count - 1 do
  begin
    if TobVentil.Detail[i].GetValue('V_TAUXMONTANT') = 0 then continue;

    j := j + 1;
    TobAnaAjout.ClearDetail;
    TobAnaAjout.Dupliquer(TobAnaDet, True, True);
    //    TobAnaDet.SaveTofile('c:\amatob\anadet', false, true, true);
    TobAnaAjout.PutValue('Y_NUMVENTIL', j);
    TobAnaAjout.PutValue('Y_AXE', 'A' + IntToStr(iAxe));
    TobAnaAjout.PutValue('Y_SECTION', TobVentil.Detail[i].GetValue('V_SECTION'));

    //SG6 17.02.05 Gestion ana croisaxe
    if VH^.AnaCroisaxe then
    begin
      for z := 1 to 5 do
      begin
        TobAnaAjout.PutValue('Y_SOUSPLAN' + IntToStr(z), TobVentil.Detail[i].GetValue('V_SOUSPLAN' + IntToStr(z)));
      end;
    end;

    TobAnaAjout.PutValue('Y_POURCENTAGE', TobVentil.Detail[i].GetValue('V_TAUXMONTANT'));

    Montant := Arrondi(TobVentil.Detail[i].GetValue('V_TAUXMONTANT') / 100 * TobAnaAjout.GetValue('Y_TOTALECRITURE'), Dec);
    MontantDev := Arrondi(TobVentil.Detail[i].GetValue('V_TAUXMONTANT') / 100 * TobAnaAjout.GetValue('Y_TOTALDEVISE'), DecDev);

    {FQ22200  14.02.08  YMO On s'assure de ne pas alimenter Y_DEBITCREDITDEV si écart de change}
    If TobAnaDet.GetString('Y_NATUREPIECE')='ECC' then MontantDev:=0;

    MAJZoneMontant(TobAnaAjout, Montant, MontantDev, Debit);

    Solde := Arrondi(Solde + Montant, Dec);
    SoldeDev := Arrondi(SoldeDev + MontantDev, Dec);
    TotTaux := Arrondi(TotTaux + TobVentil.Detail[i].GetValue('V_TAUXMONTANT'), Dec2);

    // Gestion QTE // MODIF FICHE 12570
      // ----- Qte1 ----
    if GereQte1 then
    begin
      TobAnaAjout.PutValue('Y_POURCENTQTE1', TobVentil.Detail[i].GetValue('V_TAUXQTE1'));
      TobAnaAjout.PutValue('Y_QUALIFQTE1', TobAnaAjout.GetValue('Y_QUALIFECRQTE1'));
      if GereResteQte1 and (i = (TobVentil.Detail.Count - 1)) // dernière ligne ?
      // Dans le cas du 100%, dernière ligne calculé en différence pour tombé sur la valeur totale juste :
      then Qte1 := TobAnaAjout.GetValue('Y_TOTALQTE1') - TotalQte1
        // Sinon calculé d'après le taux et arrondi à 3 décimales :
      else Qte1 := Arrondi((TobVentil.Detail[i].GetValue('V_TAUXQTE1') * TobAnaAjout.GetValue('Y_TOTALQTE1')) / 100, 3);
      TobAnaAjout.PutValue('Y_QTE1', Qte1);
      TotalQte1 := TotalQte1 + Qte1;
    end
    else
    begin
      TobAnaAjout.PutValue('Y_QUALIFQTE1', TobAnaAjout.GetValue('Y_QUALIFECRQTE1'));
      TobAnaAjout.PutValue('Y_POURCENTQTE1', 0);
      TobAnaAjout.PutValue('Y_QTE1', 0);
    end;
    // ----- Qte2 ----
    if GereQte2 then
    begin
      TobAnaAjout.PutValue('Y_POURCENTQTE2', TobVentil.Detail[i].GetValue('V_TAUXQTE2'));
      TobAnaAjout.PutValue('Y_QUALIFQTE2', TobAnaAjout.GetValue('Y_QUALIFECRQTE2'));
      if GereResteQte2 and (i = (TobVentil.Detail.Count - 1)) // dernière ligne ?
      // Dans le cas du 100%, dernière ligne calculé en différence pour tombé sur la valeur totale juste :
      then Qte2 := TobAnaAjout.GetValue('Y_TOTALQTE2') - TotalQte2
        // Sinon calculé d'après le taux et arrondi à 3 décimales :
      else Qte2 := Arrondi((TobVentil.Detail[i].GetValue('V_TAUXQTE2') * TobAnaAjout.GetValue('Y_TOTALQTE2')) / 100, 3);
      TobAnaAjout.PutValue('Y_QTE2', Qte2);
      TotalQte2 := TotalQte2 + Qte2;
    end
    else
    begin
      TobAnaAjout.PutValue('Y_QUALIFQTE2', TobAnaAjout.GetValue('Y_QUALIFECRQTE2'));
      TobAnaAjout.PutValue('Y_POURCENTQTE2', 0);
      TobAnaAjout.PutValue('Y_QTE2', 0);
    end;
    // FIN MODIF FICHE 12570

  //  ModifDerniereLigne:=False ;
    ModifDerniereLigne := (i = TobVentil.Detail.Count - 1) and (TotTaux = 100.0000) and
      ((Solde <> TobAnaAjout.GetValue('Y_TOTALECRITURE')) or
      (SoldeDev <> TobAnaAjout.GetValue('Y_TOTALDEVISE')));

    if ModifDerniereLigne then
    begin
      if Debit then Zone := 'DEBIT' else Zone := 'CREDIT';
      Montant := Arrondi(TobAnaAjout.GetValue('Y_' + Zone) + TobAnaAjout.GetValue('Y_TOTALECRITURE') - Solde, Dec);
      MontantDev := Arrondi(TobAnaAjout.GetValue('Y_' + Zone + 'DEV') + TobAnaAjout.GetValue('Y_TOTALDEVISE') - SoldeDev, DecDev);

      {FQ22200  14.02.08  YMO On s'assure de ne pas alimenter Y_DEBITCREDITDEV si écart de change}
      If TobAnaDet.GetString('Y_NATUREPIECE')='ECC' then MontantDev:=0;
      MAJZoneMontant(TobAnaAjout, Montant, MontantDev, Debit);
    end;
    TobAnaAjout.InsertDB(nil);
    MAJTobSection(TobAnaAjout, 1);
  end;

  if TotTaux <> 100.0000 then // Ajout ligne pour equilibre entre somme Compta gene et lignes analytiques si total ventil <> 100%
  begin
    TobAnaAjout.ClearDetail;
    TobAnaAjout.Dupliquer(TobAnaDet, True, True);
    LigneFinal(TobAnaAjout, Solde, SoldeDev, Arrondi(100 - TotTaux, dec2), j + 1, Debit, DecDev);
    TobAnaAjout.InsertDb(nil);
    MAJTobSection(TobAnaAjout, 1);
  end;

  TobAnaAjout.Free;
  TobAnaAjout := nil;
end;

procedure TOF_CPREVENTILANA.MAJSection;
var
  i: integer;
begin
  for i := 0 to TobSection.Detail.Count - 1 do
  begin
    if TobSection.Detail[i].GetValue('MOUVEMENTE') = 'X' then TobSection.Detail[i].UpdateDB;
  end;
end;

procedure TOF_CPREVENTILANA.FListeDblClick(Sender: TObject);
var
  sMode: string;
  Fiche: TFMul;
begin
  inherited;
  Fiche := TFMul(Ecran);
  ListeEcr := Fiche.FListe;
  {$IFDEF EAGLCLIENT}
  QListeEcr := Fiche.Q.TQ;
  {$ELSE}
  QListeEcr := Fiche.Q;
  {$ENDIF}
  if ((QListeEcr.EOF) and (QListeEcr.BOF)) then Exit;
  sMode := QListeEcr.FindField('E_MODESAISIE').AsString;
  if ((sMode <> '') and (sMode <> '-')) then LanceSaisieFolio(QListeEcr, TaConsult)
  else TrouveEtLanceSaisie(QListeEcr, TaConsult, 'N');
end;

procedure TOF_CPREVENTILANA.LibereTobs;
begin
  if TobAna <> nil then
  begin
    TobAna.Free;
    TobAna := nil;
  end;
  if TobSection <> nil then
  begin
    TobSection.Free;
    TobSection := nil;
  end;
  if TobVentil <> nil then
  begin
    TobVentil.Free;
    TobVentil := nil;
  end;
  if TobDev <> nil then
  begin
    TobDev.Free;
    TobDev := nil;
  end;
end;

procedure TOF_CPREVENTILANA.BOuvrirOnClick(Sender: TObject);
var
  ValAxe: string;
  i: Integer;
  Fiche: TFMul;
  ParDefaut: boolean;
begin
  Fiche := TFMul(Ecran);
  ListeEcr := Fiche.FListe;
  {$IFDEF EAGLCLIENT}
  QListeEcr := Fiche.Q.TQ;
  {$ELSE}
  QListeEcr := Fiche.Q;
  {$ENDIF}

  ParDefaut := GetCheckBoxState('VentilParDefaut') = CbChecked;

  if (not VH^.AnaCroisaxe and not ParDefaut) then
    if (LAxe.Text = '') or (LaVentil.Text = '') then
    begin
      PGIBox(TraduireMemoire('Vous devez sélectionner un axe et une ventilation type'),
        TraduireMemoire(Titre));
      Exit;
    end;
  if (VH^.AnaCroisaxe) and (not ParDefaut) then
    if (LaVentil.Text = '') then
    begin
      PGIBox(TraduireMemoire('Vous devez sélectionner une ventilation type'), TraduireMemoire(Titre));
      Exit;
    end;

  if VH^.AnaCroisaxe then
  begin
    iAxe := RecherchePremDerAxeVentil.premier_axe;
  end
  else
  begin
    iAxe := Ord(Laxe.Value[2]) - 48;
  end;


  if PGIAsk(TraduireMemoire('Confirmez vous le traitement ?'), TraduireMemoire(Titre)) = MrNo then Exit;
  // La section selectionnée doit être ventilable sur l'axe choisi
  if not ParDefaut then
  begin

    //SG6 17.02.05 Gestion croisaxe
    ValAxe := 'TY' + IntToStr(iAxe);

    if not VeriAxeSection(ValAxe) then
    begin
      Paschoixsaisi;
      exit;
    end;
  end;

  if (ListeEcr.NbSelected = 0) and (not ListeEcr.AllSelected) then // vérification qu'une ligne au moins a été selectionnée
  begin
    PgiBox(TraduireMemoire('Vous devez sélectionner au moins une écriture.'), TraduireMemoire(Titre));
    Exit;
  end;

  EnableControls(Fiche, False);
  RecupEcritAna; // récupération des écritures analytiques
  if TobAna.Detail.Count = 0 then
  begin
    PGIBox(TraduireMemoire('La nouvelle ventilation est identique à l''ancienne, aucun traitement à effectuer.'), TraduireMemoire(Titre));
    EnableControls(Fiche, True);
    exit;
  end;
  ChargeTobSection;
  ChargeTobDevise; // Chargement des sections et des devises

  // S'assure que toutes les lignes ont une ventilations par défaut // FICHE 12439
  if ParDefaut then
  begin
    for i := 0 to TobAna.Detail.Count - 1 do
    begin
      if not ChargeTobVentil(i) then
      begin
        PGIBox(TraduireMemoire('Le compte ' + TobAna.Detail[i].GetValue('Y_GENERAL') + ' n''a pas de ventilation par défaut.'), TraduireMemoire(Titre));
        EnableControls(Fiche, True);
        Exit;
      end;
    end;
  end
  {b FP 29/12/2005: Vérifie la cohérence du modèle de restriction avec le modèle de ventilation}
  else
  begin
    for i := 0 to TobAna.Detail.Count - 1 do
    begin
      if not VerifModelRestriction(i) then
      begin
        PGIBox(TraduireMemoire('Il existe des incompatibilités entre le modèle de ventilations à appliquer et les modèles de restrictions analytiques attachés au comptes généraux sélectionnés.'), TraduireMemoire(Titre));
        EnableControls(Fiche, True);
        Exit;
      end;
    end;
  end;
  {e FP 29/12/2005}

  try
    BeginTrans;
    DetruitLigneAna; // destruction des lignes analytiques des écritures selectionnées
    for i := 0 to TobAna.Detail.Count - 1 do
    begin
      if ParDefaut then if not ChargeTobVentil(i) then Continue;
      ModifLigne(TobAna.Detail[i]); // Modif et/ou création des lignes
    end;
    //TobAna.DeleteDB ;
    MAJSection; // maj des sections analytiques modifiées
    CommitTrans;
  except
    PGiBox(TraduireMemoire('Une erreur est survenue pendant le traitement.'), TraduireMemoire(Titre));
    RollBack;
    LibereTobs;
    EnableControls(Fiche, True);
    exit;
  end; // fin du try
  LibereTobs;
  PGiBox(TraduireMemoire('Traitement terminé'), TraduireMemoire(Titre));
  TToolBarButton97(GetControl('BCHERCHE')).Click;
  EnableControls(Fiche, True);
end;

function TOF_CPREVENTILANA.ChargeTobVentil(i: integer): boolean;
var
  Q: TQuery;
begin
  Result := False;
  if ExisteSql('SELECT V_NATURE,V_COMPTE,V_NUMEROVENTIL FROM VENTIL WHERE V_NATURE ="GE' + IntToStr(iAxe) + '" AND V_COMPTE="' +
    TobAna.Detail[i].GetValue('Y_GENERAL') + '"')
    then
    Q := OpenSql('SELECT V_SOUSPLAN1, V_SOUSPLAN2, V_SOUSPLAN3, V_SOUSPLAN4, V_SOUSPLAN5, V_NATURE,V_COMPTE,V_NUMEROVENTIL, V_MONTANT, '
      + 'V_SECTION,V_TAUXMONTANT,V_TAUXQTE1,V_TAUXQTE2 FROM VENTIL WHERE '
      + 'V_NATURE ="GE' + IntToStr(iAxe) + '" AND V_COMPTE="' + TobAna.Detail[i].GetValue('Y_GENERAL') + '"', true)
  else
    Q := OpenSql('SELECT V_SOUSPLAN1, V_SOUSPLAN2, V_SOUSPLAN3, V_SOUSPLAN4, V_SOUSPLAN5, V_NATURE,V_COMPTE,V_NUMEROVENTIL, V_MONTANT, '
      + 'V_SECTION,V_TAUXMONTANT,V_TAUXQTE1,V_TAUXQTE2 FROM VENTIL WHERE '
      + 'V_NATURE ="GE' + IntToStr(iAxe) + '" AND V_COMPTE="' + VH^.Cpta[AxeToFb('A' + IntToStr(iAxe))].Attente + '"', True);
  if Q.EOF then {begin Result := False ; } Exit; // end ;
  Result := True;
  if TobVentil = nil then TobVentil := Tob.Create('_VENTIL', nil, -1) else TobVentil.ClearDetail;
  TobVentil.LoadDetailDb('VENTIL', '', '', Q, False, False);
  Ferme(Q);
end;

procedure TOF_CPREVENTILANA.ExerciceOnChange(Sender: TObject);
var
  Q: TQuery;
begin
  Q := OpenSql('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="' + Exercice.Value + '"', True);
  SetControlText('E_DATECOMPTABLE', DateToStr(Q.Fields[0].AsDateTime));
  SetControlText('E_DATECOMPTABLE_', DateToStr(Q.Fields[1].AsDateTime));
  Ferme(Q);
end;

procedure TOF_CPREVENTILANA.ChercheVentilOnClick(Sender: TObject);
begin
  LaVentil.Value := AGLLanceFiche('CP', 'CPAFFVENTILANA', '', '', '');
  LaVentil.ReLoad;
end;

procedure TOF_CPREVENTILANA.VentilParDefautOnClick(Sender: TObject);
begin
  if GetCheckBoxState('VENTILPARDEFAUT') = CbChecked then
  begin
    SetControlEnabled('TLAVENTIL', False);
    SetControlEnabled('LAVENTIL', False);
  end else
  begin
    SetControlEnabled('TLAVENTIL', True);
    SetControlEnabled('LAVENTIL', True);
  end;
end;

procedure TOF_CPREVENTILANA.OnArgument(S: string);
{b fb 14/02/2006}
var
  lSt : string;
{e fb 14/02/2006}
begin
  Titre := 'Re-ventilation analytique';
  Dec := 2;
  Dec2 := 4;
  Ecran.HelpContext := 999999105;
  BOuvrir := TToolBarButton97(GetControl('BOUVRIR'));
  BCherche := TToolBarButton97(GetControl('BCHERCHE'));
  ChercheVentil := TToolBarButton97(GetControl('chercheventil'));
  Exercice := THValComboBox(GetControl('E_EXERCICE'));
  Journal := THValComboBox(GetControl('E_JOURNAL'));
  General := THValComboBox(GetControl('E_GENERAL'));
  General_ := THValComboBox(GetControl('E_GENERAL_'));
  DateComptable := THvalComboBox(GetControl('E_DATECOMPTABLE'));
  DateComptable_ := THvalComboBox(GetControl('E_DATECOMPTABLE_'));
  LAxe := THValComboBox(GetControl('LAXE'));
  LaVentil := THValComboBox(GetControl('LAVENTIL'));
  BOuvrir.OnClick := BOuvrirOnClick;
  Exercice.OnChange := ExerciceOnChange;
  SetControlText('E_DATECOMPTABLE', DateToStr(VH^.EnCours.Deb));
  SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.EnCours.Fin));
  LAxe.Itemindex := 0;
  THDbgrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
  SetControlText('E_EXERCICE', VH^.EnCours.Code);
  ChercheVentil.OnClick := ChercheVentilOnClick;
  Exercice.Value := VH^.EnCours.Code;
  TCheckBox(GetControl('VENTILPARDEFAUT')).OnClick := VentilParDefautOnClick;
  Journal.ItemIndex := 1;
{b fb 16/06/2006 FQ17912}
  ZC1 := THValComBoBox(GetControl('Z_C1'));
  ZO1 := THValComBoBox(GetControl('ZO1'));
  ZV1 := THEdit(GetControl('ZV1'));
  ZC2 := THValComBoBox(GetControl('Z_C2'));
  ZO2 := THValComBoBox(GetControl('ZO2'));
  ZV2 := THEdit(GetControl('ZV2'));
{e fb 16/06/2006 FQ17912}
  // 10469
  if VH^.CPExoRef.Code <> '' then
  begin
    Exercice.Value := VH^.CPExoRef.Code;
    //  DateComptable.Text := DateToStr(VH^.CPExoRef.Deb);
    //  DateComptable_.Text := DateToStr(VH^.CPExoRef.Fin);
  end
  else      
  begin
    Exercice.Value := VH^.Entree.Code;
  end;
  DateComptable.Text := DateToStr(V_PGI.DateEntree);
  DateComptable_.Text := DateToStr(V_PGI.DateEntree);

  //SG6 19.02.05 Gestion mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    LAxe.Text := '';
    LAxe.Visible := False;
    SetControlVisible('TLAxe', False);
    SetControlProperty('FE__GROUPBOX1', 'Left', 28);
  end;

{b fb 14/02/2006}
  lSt := ReadTokenSt(S);
  if lSt <> '' then
    Journal.Value  := lSt;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    Exercice.Value := lSt;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('E_QualifPiece', lSt);

  lSt := ReadTokenSt(S);
  if lSt <> '' then begin
    General.Text  := lSt;
    General_.Text  := lSt;
    end;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    DateComptable.Text  := lSt;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    DateComptable_.Text  := lSt;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    LAxe.Text  := lSt;
{e fb 14/02/2006}
{b fb 16/06/2006 FQ17912}
  lSt := ReadTokenSt(S);
  if lSt <> '' then begin
    SetControlText('ZV1',lSt);

    lSt := ReadTokenSt(S);
    if lSt <> '' then
      SetControlText('ZV2',lSt);

    TFMul(Ecran).OnAfterFormShow := AfterShow;
    end;
{e fb 16/06/2006 FQ17912}

  inherited;
end;

function TOF_CPREVENTILANA.VerifModelRestriction(i: Integer): Boolean;
var
  RestrictAna: TRestrictionAnalytique;
begin
  {b FP 29/12/2005}
  RestrictAna := TRestrictionAnalytique.Create;
  try
    Result := RestrictAna.IsCompteGeneAutorise(TobAna.Detail[i].GetValue('Y_GENERAL'),
                   'A'+IntToStr(iAxe), 'TY'+IntToStr(iAxe), LaVentil.Value);
  finally
    RestrictAna.Free;
  end;
  {e FP 29/12/2005}
end;

{b fb 16/06/2006 FQ17912}
procedure TOF_CPREVENTILANA.AfterShow;
begin
  ZC1.Value:='E_NUMEROPIECE';
  ZO1.ItemIndex:=3;
  ZC2.Value:='E_NUMLIGNE';
  ZO2.ItemIndex:=ZO1.ItemIndex;
  TToolBarButton97(GetControl('BCHERCHE')).Click;
end;
{e fb 16/06/2006 FQ17912}
initialization
  registerclasses([TOF_CPREVENTILANA]);
end.

