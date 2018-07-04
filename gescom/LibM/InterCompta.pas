{***********UNITE*************************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 19/03/2002
Modifié le ... : 25/03/2002
Description .. : Chargement de la table InterCompta pour l'exportation d'un
Suite ........ : fichier compta.
Mots clefs ... : COMPTA;MODE
*****************************************************************}
unit InterCompta;

interface

uses UTOB, UtilInterCompta,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  DBTables,
  {$ENDIF}
  HCtrls, SysUtils, HEnt1, UtilPGI, Ent1, EntGC, ParamSoc, HMsgBox, HStatus, MajTable, ed_tools,
  SaisUtil, UtilUtilitaires;

function RecupParametre(var LastError: Integer; Nature, CriEtab, CriDatedeb, Cridatefin, EtatCompta: string; TobPieceEqui: TOB = nil): Boolean;
function TraitementNature(): boolean;
// Chargement des informations pour le regroupement des requetes
procedure AucunRegroupement(NatCou: string);
procedure RegroupementParJour(NatCou: string);
procedure RegroupementParSemaine(NatCou: string);
procedure RegroupementParMois(NatCou: string);
procedure ArrondirLesMontants(TobL: TOB; NomChpMt: string);
//
// Chargement des ligne et pièce de vente
function ChargeLesLignes(NatCou, Champ, GroupBy, Trie: string): Boolean;
procedure ChargeLesPieces(NatCou, Champ, GroupBy: string; RDate: Boolean);
//
// Traitement pour les pièces gérant les acomptes (GPP_ACOMPTE="X")
procedure GereAcomptesEcheances(NatCou, Champ, GroupBy, Trie: string; Rdate: boolean; Grp: string);
function ChargeLesEcheances(NatCou, Champ, GroupBy, Trie: string; RDate: Boolean; Grp: string): TOB;
function ChargeLesAcomptes(NatCou, Champ, GroupBy: string; Rdate: boolean; Grp: string): TOB;
function RegroupeAcomptesEcheance(TobAcomptes, TobEcheances: TOB): TOB;
procedure FusionIntercomptaAcpteEche(TobRegroupeAcEc: TOB; Grp: string);
//
// Traitement pour les pièces ne gérant pas les acomptes (GPP_ACOMPTE="-")
procedure ChargeLesLigneArticleEtOP(NatCou, Champ, GroupBy, Trie: string);
procedure ChargeLesModesPaie(NatCou, Champ, GroupBy, Trie, Regroupement: string; RDate: Boolean);
procedure RegroupementModePaie(TOBInterTemp: TOB; Regroupement: string);
//
procedure Initialisation;
function FindTOB(TobRech, TOBOrig: Tob; RDate, First: Boolean): Tob;
procedure CreerChamps(TOBInter, TOBLig, TobEch: TOB; Prix, PrixDev: Double; Virtuelle: Boolean; Regroupement: string);
function Rupture(TOBL, TOBSuiv: TOB; indice: Integer; regroupement: string): boolean;
function LesTobEgales(TOBLig, TOBEch: TOB): Boolean;
function ComparePrix(TOBLig, TOBEch: TOB): Integer;
procedure ChargeNatureDansTable;
procedure TraitementPremierJourSemaine(TOBCourante: TOB; Piece: Boolean);
procedure TraitementPremierJourMois(TOBCourante: TOB; Piece: Boolean);
procedure LibererTob;
// Fonction pour les axes analytiques paramettre
function IsLigneAna(TobL: TOB): boolean;
procedure TraitementDesAxesAnalytiques(NatCou, Champ, GroupBy, Trie: string; IsOperationCaisse: Boolean = False);
procedure ConstruireTobAna(VenteAchat: string);
procedure ConstruireGroupByAna;
procedure ConstruireLibelleAna;
procedure ChargeLesAxeAnalytique(NatCou, Champ, GroupBy, Trie, CodeAxe, Etab: string);
procedure ChargeLesAxeAnalytiqueOperationCaisse(NatCou, Champ, GroupBy, Trie, CodeAxe, Etab: string);
function GetChampSecAxeAnalytique(Etablissement, CodeAxe: string): string;
function GetChampLibAxeAnalytique(Etablissement, CodeAxe: string): string;
procedure PutAxeAnalytique(CodeAxe: string; ModePaie: Boolean);
function ChercheAxeAnalytiqueDesOP(CodeAxe, Etab: string): Boolean;
//
// Gestion Port et Frais
procedure EcrireLignePortEtFrais;
//
procedure AjoutChampsTob(quoi: string);

implementation

var TOBInterCompta, TOBPiece, TOBAna, TOBLigne, TobPieceEqui_: TOB;
  NaturePiece, Etab, EtabOr, DateDeb, DateFin, Comptabilise,
    ChampSecAna, ChampSecAnaReq, ChampLibAna, ChampLibAnaReq: string;
  iCount, NumError: Integer;
  AxeParam, GestionAcompte, AvecEcheance, EstAvoir: Boolean;

function RecupParametre(var LastError: Integer; Nature, CriEtab, CriDatedeb, Cridatefin, EtatCompta: string; TobPieceEqui: TOB = nil): Boolean;
begin
  LastError := 0;
  Result := False;
  Etab := criEtab;
  EtabOr := CriEtab;
  if Etab = '' then Etab := ' ';
  DateDeb := criDateDeb;
  DateFin := criDateFin;
  Comptabilise := EtatCompta;
  TobPieceEqui_ := TobPieceEqui;
  Initialisation;
  try
    NaturePiece := Nature;
    if not TraitementNature() then LastError := NumError
    else Result := True;
  finally
    LibererTOB;
  end;
end;

procedure Initialisation;
begin
  iCount := 0;
  ExecuteSQL('DELETE FROM INTERCOMPTA WHERE GIC_USER="' + V_PGI.User + '"');
  TOBInterCompta := TOB.Create('_INTERCOMPTA', nil, -1);
  TOBPiece := TOB.Create('_PIECE', nil, -1);
  TOBLigne := TOB.Create('_LIGNE', nil, -1);
  //TOBEcheance := TOB.Create('_ECHEANCE', nil, -1);
  AxeParam := GetParamSoc('SO_GCAXEANALYTIQUE');
end;

function TraitementNature(): boolean;
var NatCou, Regroupement, Natures, Temp: string;
begin
  Natures := NaturePiece;
  // Création d'un chaine utilisable dans une requete
  while Etab <> '' do
  begin
    if Temp = '' then Temp := '"' + ReadTokenSt(Etab) + '"'
    else Temp := Temp + ',' + '"' + ReadTokenSt(Etab) + '"';
  end;
  Etab := Temp;
  repeat
    NatCou := ReadTokenSt(Natures);
    MoveCurProgressForm('Traitement: "' + RechDom('GCNATUREPIECEG', NatCou, False) + '"');
    if NatCou <> '' then
    begin
      TobInterCompta.ClearDetail;
      ChampSecAna := '';
      ChampSecAnaReq := '';
      Regroupement := GetInfoParPiece(NatCou, 'GPP_REGROUPCPTA');
      GestionAcompte := GetInfoParPiece(NatCou, 'GPP_ACOMPTE') = 'X';
      EstAvoir := GetInfoParPiece(NatCou, 'GPP_ESTAVOIR') = 'X';
      if Regroupement = 'AUC' then AucunRegroupement(NatCou)
      else if Regroupement = 'JOU' then RegroupementParJour(NatCou)
      else if Regroupement = 'SEM' then RegroupementParSemaine(NatCou)
      else if Regroupement = 'MOI' then RegroupementParMois(NatCou);
    end;
    EcrireLignePortEtFrais;
    ChargeNatureDansTable;
  until Natures = '';
  if NumError = 0 then Result := True else Result := False;
end;

procedure ChargeNatureDansTable;
var i: integer;
  InsertOk: Boolean;
  TobTmp: TOB;
begin
  if NumError = 0 then
  begin
    for i := 0 to TobInterCompta.Detail.count - 1 do
    begin
      TobTmp := TobIntercompta.Detail[i];
      TobTmp.PutValue('GIC_COMPTEUR', iCount);
      TobTmp.PutValue('GIC_USER', V_PGI.User);
      if not TobTmp.FieldExists('GIC_ISOPCAISSE') then
        TobTmp.AddChampSupValeur('GIC_ISOPCAISSE', '-');
      if IsNatCaisse(TobTmp.GetValue('GIC_NATUREPIECEG')) then
      begin
        if ((TobTmp.GetValue('GIC_MODEPAIE') = '') or (GetInfoParPiece(TobTmp.GetValue('GIC_NATUREPIECEG'), 'GPP_ACOMPTE') = 'X')
          and ((TobTmp.GetValue('GIC_NUMECHE')) > 0)) then
          TobTmp.PutValue('GIC_TYPEJOURNAL', 'VEN')
        else
          TobTmp.PutValue('GIC_TYPEJOURNAL', 'BTQ');
      end else
        if (TobTmp.GetValue('GIC_TYPEJOURNAL') = 'VEN') and (TobTmp.GetValue('GIC_NUMECHE') > 1)
        and (TobTmp.GetValue('GIC_MODEPAIE') = '') then
        TobTmp.PutValue('GIC_MONTANTTTC', 0);
      inc(icount);
    end;
    TobInterCompta.SetAllModifie(True);
    InsertOk := TobInterCompta.InsertDB(nil);
  end else InsertOk := False;
  if InsertOk then numError := 0 else NumError := 1 // Chargement non effectué
end;

procedure LibererTOB;
begin
  FreeAndNil(TobIntercompta);
  FreeAndNil(TobPiece);
  //if TOBEcheance <> nil then FreeAndNil(TOBEcheance);
  FreeAndNil(TOBLigne);
  if TOBAna <> nil then FreeAndNil(TOBAna);
end;

// Traitement à la pièce

procedure AucunRegroupement(NatCou: string);
var ChampL, GroupByL, TrieL, ChampP, GroupByP,
  ChampMP, GroupByMP, TrieMP: string;
  LigneOk: Boolean;
begin
  AjoutChampsTob('');
  ChampL := 'gl_datepiece as gic_date,gl_souche as gic_souche, gl_numero as gic_numero,gl_indiceg as gic_indice, ';
  GroupByL := 'gl_datepiece,gl_souche,gl_numero,gl_indiceg, ';
  TrieL := 'Gic_Date;gic_souche;gic_numero;gic_indice; ';
  ChampP := 'gp_datepiece as gic_date,gp_souche as gic_souche, gp_numero as gic_numero,gp_indiceg as gic_indice, ';
  GroupByP := 'gp_datepiece,gp_souche, gp_numero,gp_indiceg, ';
  LigneOK := ChargeLesLignes(NatCou, ChampL, GroupByL, TrieL);
  if LigneOK then
    ChargeLesPieces(NatCou, ChampP, GroupByP, False);
  if AxeParam then
    TraitementDesAxesAnalytiques(NatCou, ChampL, GroupByL, TrieL);
  ChampMP := 'gpe_datepiece as gic_date,gpe_souche as gic_souche, gpe_numero as gic_numero,gpe_indiceg as gic_indice, ';
  GroupByMP := 'gpe_datepiece,gpe_souche,gpe_numero,gpe_indiceg, ';
  TrieMP := 'Gic_Date;gic_souche;gic_numero;gic_indice; ';
  if IsNatCaisse(NatCou) then
  begin
    ChargeLesLigneArticleEtOP(NatCou, ChampL, GroupByL, TrieL);
    ChargeLesModesPaie(NatCou, ChampP, GroupByP, TrieMP, 'AUC', False);
    if AxeParam then
      TraitementDesAxesAnalytiques(NatCou, ChampMP, GroupByMP, TrieMP, True);
  end else
  begin
    GereAcomptesEcheances(NatCou, ChampP, GroupByP, TrieMP, False, 'AUC');
  end;
end;

// Traitement à la journée

procedure RegroupementParJour(NatCou: string);
var ChampL, GroupByL, TrieL, ChampP, GroupByP,
  ChampMP, GroupByMP, TrieMP: string;
  LigneOk: Boolean;
begin
  AjoutChampsTob('');
  ChampL := 'gl_datepiece as gic_date,';
  GroupByL := 'gl_datepiece, ';
  TrieL := 'Gic_Date; ';
  ChampP := 'gp_datepiece as gic_date, ';
  GroupByP := 'gp_datepiece, ';
  LigneOK := ChargeLesLignes(NatCou, ChampL, GroupByL, TrieL);
  if LigneOK then
    ChargeLesPieces(NatCou, ChampP, GroupByP, True);
  if AxeParam then
    TraitementDesAxesAnalytiques(NatCou, ChampL, GroupByL, TrieL);
  ChampMP := 'gpe_datepiece as gic_date,';
  GroupByMP := 'gpe_datepiece, ';
  TrieMP := 'Gic_Date; ';
  if IsNatCaisse(NatCou) then
  begin
    ChargeLesLigneArticleEtOP(NatCou, ChampL, GroupByL, TrieL);
    ChargeLesModesPaie(NatCou, ChampMP, GroupByMP, TrieMP, 'JOU', True);
    if AxeParam then
      TraitementDesAxesAnalytiques(NatCou, ChampMP, GroupByMP, TrieMP, True);
  end else
  begin
    GereAcomptesEcheances(NatCou, ChampP, GroupByP, TrieMP, False, 'JOU');
  end;
end;

// Traitement à la semaine

procedure RegroupementParSemaine(NatCou: string);
var ChampL, GroupByL, TrieL, ChampP, GroupByP,
  ChampMP, GroupByMP, TrieMP: string;
  LigneOk: Boolean;
begin
  AjoutChampsTob('SEM');
  ChampL := '' + DB_Year('gl_datepiece') + ' as annee, ' + DB_Week('gl_datepiece') + ' as semaine,';
  GroupByL := '' + DB_Year('gl_datepiece') + ', ' + DB_Week('gl_datepiece') + ', ';
  TrieL := 'Gic_Date; ';
  ChampP := '' + DB_Year('gp_datepiece') + ' as annee, ' + DB_Week('gp_datepiece') + ' as semaine, ';
  GroupByP := '' + DB_Year('gp_datepiece') + ', ' + DB_Week('gp_datepiece') + ', ';
  LigneOK := ChargeLesLignes(NatCou, ChampL, GroupByL, TrieL);
  if LigneOK then
    ChargeLesPieces(NatCou, ChampP, GroupByP, True);
  if AxeParam then
    TraitementDesAxesAnalytiques(NatCou, ChampL, GroupByL, TrieL);
  ChampMP := '' + DB_Year('gpe_datepiece') + ' as annee, ' + DB_Week('gpe_datepiece') + ' as semaine,';
  GroupByMP := '' + DB_Year('gpe_datepiece') + ', ' + DB_Week('gpe_datepiece') + ', ';
  TrieMP := 'Gic_Date; ';
  if IsNatCaisse(NatCou) then
  begin
    ChargeLesLigneArticleEtOP(NatCou, ChampL, GroupByL, TrieL);
    ChargeLesModesPaie(NatCou, ChampMP, GroupByMP, TrieMP, 'SEM', True);
    if AxeParam then
      TraitementDesAxesAnalytiques(NatCou, ChampMP, GroupByMP, TrieMP, True);
  end else
  begin
    GereAcomptesEcheances(NatCou, ChampP, GroupByP, TrieMP, False, 'SEM');
  end;
end;

// Traitement au mois

procedure RegroupementParMois(NatCou: string);
var ChampL, GroupByL, TrieL, ChampP, GroupByP,
  ChampMP, GroupByMP, TrieMP: string;
  LigneOk: Boolean;
begin
  AjoutChampsTob('MOI');
  ChampL := '' + DB_Year('gl_datepiece') + ' as annee, ' + DB_Month('gl_datepiece') + ' as mois,';
  GroupByL := '' + DB_Year('gl_datepiece') + ', ' + DB_Month('gl_datepiece') + ', ';
  TrieL := 'Gic_Date; ';
  ChampP := '' + DB_Year('gp_datepiece') + ' as annee, ' + DB_Month('gp_datepiece') + ' as mois, ';
  GroupByP := '' + DB_Year('gp_datepiece') + ', ' + DB_Month('gp_datepiece') + ', ';
  LigneOK := ChargeLesLignes(NatCou, ChampL, GroupByL, TrieL);
  if LigneOK then
    ChargeLesPieces(NatCou, ChampP, GroupByP, True);
  if AxeParam then
    TraitementDesAxesAnalytiques(NatCou, ChampL, GroupByL, TrieL);
  ChampMP := '' + DB_Year('gpe_datepiece') + ' as annee, ' + DB_Month('gpe_datepiece') + ' as mois,';
  GroupByMP := '' + DB_Year('gpe_datepiece') + ', ' + DB_Month('gpe_datepiece') + ', ';
  TrieMP := 'Gic_Date; ';
  if IsNatCaisse(NatCou) then
  begin
    ChargeLesLigneArticleEtOP(NatCou, ChampL, GroupByL, TrieL);
    ChargeLesModesPaie(NatCou, ChampMP, GroupByMP, TrieMP, 'MOI', True);
    if AxeParam then
      TraitementDesAxesAnalytiques(NatCou, ChampMP, GroupByMP, TrieMP, True);
  end else
  begin
    GereAcomptesEcheances(NatCou, ChampP, GroupByP, TrieMP, False, 'MOI');
  end;
end;

procedure ArrondirLesMontants(TobL: TOB; NomChpMt: string);
var Montant: double;
begin
  Montant := Arrondi(Valeur(NullToVide(TobL.GetValue(NomChpMt))), V_PGI.OkDecV);
  TobL.PutValue(NomChpMt, Montant);
end;

// Chargement des informations concernant les lignes et les articles

function ChargeLesLignes(NatCou, Champ, GroupBy, Trie: string): Boolean;
var SQLLigne, StSelect, PlusWhere, Regroup, stCoeff, NomTaxe1, NomTaxe2: string;
  QLigne: TQuery;
  Cpt, Numero, Coeff: integer;
  TobTmp: TOB;
begin
  StSelect := 'select ' + Champ + 'gl_etablissement as gic_etablissement, gl_tiers as gic_tiers,gl_naturepieceg as gic_naturepieceg,' +
    'gl_regimetaxe as gic_regimetaxe, gl_familletaxe1 as gic_familletaxe1,gl_familletaxe2 as gic_familletaxe2,gl_familletaxe3 as gic_familletaxe3,' +
    'gl_familletaxe4 as gic_familletaxe4,gl_familletaxe5 as gic_familletaxe5,ga_comptaarticle as gic_comptaarticle, ' +
    'gl_devise as gic_devise, ' +
    'sum(gl_qtefact) as GIC_qte1, ' +
    'sum(round(gl_totalTaxe1,'+IntToStr(V_PGI.OkDecV)+')) as Gic_montantTaxe1,' +
    'sum(round(gl_totalTaxe2,'+IntToStr(V_PGI.OkDecV)+')) as Gic_montantTaxe2,' +
    'sum(round(gl_totalTaxe3,'+IntToStr(V_PGI.OkDecV)+')) as Gic_montantTaxe3,' +
    'sum(round(gl_totalTaxe4,'+IntToStr(V_PGI.OkDecV)+')) as Gic_montantTaxe4,' +
    'sum(round(gl_totalTaxe5,'+IntToStr(V_PGI.OkDecV)+')) as Gic_montantTaxe5,' +
    'sum(round(gl_montantht,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantht, '+
    'sum(round(gl_montanthtdev,'+IntToStr(V_PGI.OkDecV)+')) as gic_montanthtdev, ' +
    'sum(round(gl_montantttc,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttc, '+
    'sum(round(gl_totalttc,'+IntToStr(V_PGI.OkDecV)+')) as gic_totalttc, '+
    'sum(round(gl_montantttcdev,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttcdev ' +
    'from piece left join ligne on gl_naturepieceg=gp_naturepieceg and gl_souche=gp_souche and gl_numero=gp_numero ' +
    'and gl_indiceg=gp_indiceg ' +
    'left join article on ga_article=gl_article ' +
    'where gp_naturepieceg in ("' + NatCou + '") and gl_datepiece>="' + UsDateTime(StrToDate(DateDeb)) + '" and gl_datepiece<="' +
    UsDateTime(StrToDate(DateFin)) + '" and (gl_montantht<>0 and gl_montantttc<>0) and gp_etablissement in (' + Etab + ') ' +
    'and gp_etatcompta in (' + Comptabilise + ') and gl_typeligne<>"COM" ';
  PlusWhere := ' and ga_typearticle<>"FI" ';
  Regroup := 'group by ' + GroupBy + 'gl_etablissement,gl_tiers,gl_naturepieceg,gl_regimetaxe,gl_devise,' +
    'gl_familletaxe1,gl_familletaxe2 ,gl_familletaxe3 ,gl_familletaxe4 ,gl_familletaxe5,' +
    'ga_comptaarticle ';
  SQLLigne := StSelect + PlusWhere + Regroup;
  QLigne := OpenSQL(SQLLigne, True);
  if not QLigne.Eof then
  begin
    Cpt := V_PGI.OkDecV;
    stCoeff := '1';
    while Cpt > 0 do
    begin
      stCoeff := stCoeff + '0';
      Cpt := Cpt - 1;
    end;
    Coeff := StrToInt(stCoeff);
    TOBInterCompta.LoadDetailDB('INTERCOMPTA', '', '', QLigne, False);
    if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBInterCompta, False)
    else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBInterCompta, False);
    // AC ETAB
    TOBInterCompta.Detail.Sort('' + Trie + 'gic_tiers;gic_naturepieceg;gic_etablissement;gic_regimetaxe');
    Result := True;
  end
  else Result := False;
  if Result then NumError := 0 else NumError := 1;
  Ferme(QLigne);
end;
//

// Chargement des informations concernant les pièces,les tiers,les ports et les etablissements

procedure ChargeLesPieces(NatCou, Champ, GroupBy: string; RDate: Boolean);
var SQLPiece, NomTaxe: string;
  QPiece: TQuery;
  i, Cpt: Integer;
  TOBL, TOBP: TOB;
  ChargementOk: Boolean;
  NumEche: Integer;
begin
  ChargementOk := False;
  // AC ETAB
  SQLPiece := 'select ' + Champ + ' gp_tiers as gic_tiers,gp_etablissement as gic_etablissement' +
    ',gp_naturepieceg as gic_naturepieceg, ' +
    'gp_regimetaxe as gic_regimetaxe, gp_devise as gic_devise,' +
    'gp_tauxdev as gic_tauxdevise, gp_tvaencaissement as gic_tvaencaissement,' +
    'gp_venteachat as gic_venteachat, t_comptatiers as gic_comptatiers, ' +
    'sum(gp_totalesc) as gic_montantesc,sum(gp_totalescdev) as gic_montantescdev,' +
    'sum(gp_totalremise) as gic_montantrem,sum(gp_totalremisedev) as gic_montantremdev, ' +
    'sum(gpt_totalht) as gic_Montantpf,sum(gpt_totalhtdev) as gic_Montantpfdev,  ' +
    'sum(gpt_totaltaxe1) as TAXEPORT1 ,sum(gpt_totaltaxe2) as TAXEPORT2, ' +
    'sum(gpt_totaltaxedev1) as TAXEPORTDEV1 ,sum(gpt_totaltaxedev2) as TAXEPORTDEV2, ' +
    'sum(gpt_totalttc) as TAXEPORTTTC ' +
    'from piece left join tiers on gp_tiers=t_tiers ' +
    'left join piedport on gp_naturepieceg=gpt_naturepieceg and gp_souche=gpt_souche and gp_numero=gpt_numero ' +
    'and gp_indiceg=gpt_indiceg ' +
    'where gp_naturepieceg in ("' + NatCou + '") and gp_datepiece>="' + USDateTime(StrToDate(DateDeb)) + '" and gp_datepiece<="' +
    USDateTime(StrToDate(DateFin)) + '" ' +
    'and gp_etatcompta in (' + Comptabilise + ') and gp_etablissement in (' + Etab + ')' +
    'group by ' + GroupBy + ' gp_etablissement,gp_tiers,gp_naturepieceg,gp_venteachat,gp_devise,gp_tauxdev,gp_tvaencaissement,gp_regimetaxe,t_comptatiers  ';
  QPiece := OpenSQL(SQLPiece, True);
  if not QPiece.Eof then
  begin
    TOBPiece.LoadDetailDB('_LAPIECE', '', '', QPiece, False);
    if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBPiece, True)
    else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBPiece, True);
    InitMove(TOBPiece.Detail.count, '');
    for i := 0 to TOBPiece.Detail.count - 1 do
    begin
      MoveCur(False);
      TOBP := TOBPiece.Detail[i];
      TOBL := FindTOB(TOBInterCompta, TOBP, RDate, True);
      NumEche := 1;
      if TOBL <> nil then
      begin
        TOBL.PutValue('GIC_MONTANTESC', TOBP.GetValue('GIC_MONTANTESC'));
        TOBL.PutValue('GIC_MONTANTREM', TOBP.GetValue('GIC_MONTANTREM'));
        TOBL.PutValue('GIC_MONTANTPF', TOBP.GetValue('GIC_MONTANTPF'));
        TOBL.PutValue('GIC_MONTANTPFDEV', TOBP.GetValue('GIC_MONTANTPFDEV'));
        TOBL.PutValue('GIC_TOTALTTC', TOBL.GetValue('GIC_TOTALTTC') + TOBP.GetValue('TAXEPORTTTC'));
        for Cpt := 1 to 2 do
        begin
          NomTaxe := 'TAXEPORT' + IntToStr(Cpt);
          if TOBP.FieldExists(NomTaxe) then
          begin
            ArrondirLesMontants(TobP, NomTaxe);
            TOBL.AddChampSupValeur(NomTaxe, TOBP.GetValue(NomTaxe), False)
          end else
          begin
            TOBL.AddChampSupValeur(NomTaxe, 0, False);
          end;
          NomTaxe := 'TAXEPORTDEV' + IntToStr(Cpt);
          if TOBP.FieldExists(NomTaxe) then
            TOBL.AddChampSupValeur(NomTaxe, TOBP.GetValue(NomTaxe), False)
          else
            TOBL.AddChampSupValeur(NomTaxe, 0, False);
        end;
        repeat
          if not TOBL.FieldExists('GIC_NUMECHE') then
            TOBL.AddChampSup('GIC_NUMECHE', false);
          TOBL.PutValue('GIC_NUMECHE', NumEche);
          TOBL.PutValue('GIC_DEVISE', TOBP.GetValue('GIC_DEVISE'));
          TOBL.PutValue('GIC_TAUXDEVISE', TOBP.GetValue('GIC_TAUXDEVISE'));
          TOBL.PutValue('GIC_TVAENCAISSEMENT', TOBP.GetValue('GIC_TVAENCAISSEMENT'));
          TOBL.PutValue('GIC_COMPTATIERS', TOBP.GetValue('GIC_COMPTATIERS'));
          TOBL.PutValue('GIC_VENTEACHAT', TOBP.GetValue('GIC_VENTEACHAT'));
          TOBL.PutValue('GIC_DATEECHE', TOBP.GetValue('GIC_DATE'));
          TOBL := FindTOB(TOBInterCompta, TOBP, RDate, False);
          NumEche := NumEche + 1;
        until TOBL = nil;
        ChargementOk := True;
      end;
    end;
  end;
  if ChargementOk then NumError := 0 else NumError := 2;
  FiniMove;
  Ferme(Qpiece);
end;
//

//*********************************** Avec Acompte ***************************//

procedure GereAcomptesEcheances(NatCou, Champ, GroupBy, Trie: string; Rdate: boolean; Grp: string);
var TobAcomptes, TobEcheances, TobRegroupeAcEc: TOB;
begin
  TobAcomptes := ChargeLesAcomptes(NatCou, Champ, GroupBy, RDate, Grp);
  TobEcheances := ChargeLesEcheances(NatCou, Champ, GroupBy, Trie, RDate, Grp);
  TobRegroupeAcEc := RegroupeAcomptesEcheance(TobAcomptes, TobEcheances);
  FusionIntercomptaAcpteEche(TobRegroupeAcEc, Grp);
  if TobAcomptes <> nil then FreeAndNil(TobAcomptes);
  if TobEcheances <> nil then FreeAndNil(TobEcheances);
  if TobRegroupeAcEc <> nil then FreeAndNil(TobRegroupeAcEc);
end;

function ChargeLesAcomptes(NatCou, Champ, GroupBy: string; Rdate: boolean; Grp: string): TOB;
var SQLAcomptes: string;
  QAcomptes: TQuery;
  TOBAcompte, TobTmp: TOB;
  Cpt: Integer;
  Date: string;
begin
  TOBAcompte := nil;
  SQLAcomptes := 'SELECT ' + Champ + ' ' +
    'GP_ETABLISSEMENT AS GIC_ETABLISSEMENT, ' +
    'GP_NATUREPIECEG AS GIC_NATUREPIECEG, ' +
    'GP_REGIMETAXE AS GIC_REGIMETAXE, ' +
    'GP_DEVISE AS GIC_DEVISE, ' +
    'GP_TAUXDEV AS GIC_TAUXDEVISE, ' +
    'GP_TVAENCAISSEMENT AS GIC_TVAENCAISSEMENT,' +
    'GP_VENTEACHAT AS GIC_VENTEACHAT, ' +
    'GAC_MODEPAIE AS GIC_MODEPAIE, ' +
    'GP_TIERS AS GIC_TIERS, ' +
    'SUM(GAC_MONTANT) AS GIC_MONTANTTTC, ' +
    'SUM(GAC_MONTANTDEV) AS GIC_MONTANTTTCDEV, ' +
    'GAC_ISREGLEMENT AS GIC_ISREGLEMENT, ' +
    'GP_DATEPIECE AS GIC_DATE ' +
    'FROM ACOMPTES ' +
    'LEFT JOIN PIECE ON ' +
    'GP_NATUREPIECEG=GAC_NATUREPIECEG AND GP_SOUCHE=GAC_SOUCHE AND ' +
    'GP_NUMERO=GAC_NUMERO AND GP_INDICEG=GAC_INDICEG ' +
    'WHERE ' +
    'GP_NATUREPIECEG IN ("' + NatCou + '") AND ' +
    'GP_DATEPIECE>="' + USDateTime(StrToDate(DateDeb)) + '" AND ' +
    'GP_DATEPIECE<="' + USDateTime(StrToDate(DateFin)) + '" AND ' +
    'GP_ETATCOMPTA IN (' + Comptabilise + ') AND ' +
    'GP_ETABLISSEMENT IN (' + Etab + ')' +
    'GROUP BY ' + GroupBy + ' ' +
    'GP_ETABLISSEMENT,GP_NATUREPIECEG,GP_VENTEACHAT,' +
    'GP_DEVISE,GP_TAUXDEV,GP_TVAENCAISSEMENT,GP_REGIMETAXE,' +
    'GAC_MODEPAIE,GP_TIERS,GAC_ISREGLEMENT,GP_DATEPIECE';
  QAcomptes := OpenSQL(SQLAcomptes, True);
  if not QAcomptes.Eof then
  begin
    {Pour traitement des pieces}
    TOBAcompte := TOB.Create('_ACOMPTES', nil, -1);
    TOBAcompte.LoadDetailDB('_ACOMPTES', '', '', QAcomptes, False);
    for Cpt := 0 to TobAcompte.detail.count - 1 do
    begin
      TobTmp := TobAcompte.Detail[Cpt];
      if (Grp = 'MOI') or (Grp = 'SEM') then
      begin
        Date := DateToStr(TobTmp.GetValue('GIC_DATE'));
        TobTmp.AddChampSupValeur('ANNEE', Copy(Date, 7, 4));
        TobTmp.AddChampSupValeur('MOIS', Copy(Date, 4, 2), False);
        TobTmp.AddChampSupValeur('SEMAINE', NumSemaine(TobTmp.GetValue('GIC_DATE')), False);
      end;
      TobTmp.AddChampSupValeur('GIC_DATEECHE', '', false);
      ArrondirLesMontants(TobTmp, 'GIC_MONTANTTTC');
    end;
    if Grp = 'SEM' then
      TraitementPremierJourSemaine(TOBAcompte, True)
    else if Grp = 'MOI' then
      TraitementPremierJourMois(TOBAcompte, True);
    {Pour traitement des reglements (journal de banque)}
    TOBInterCompta.LoadDetailDB('INTERCOMPTA', '', '', QAcomptes, True);
    if TOBInterCompta.FieldExists('SEMAINE') then
      TraitementPremierJourSemaine(TOBInterCompta, True)
    else
      if TOBInterCompta.FieldExists('MOIS') then
      TraitementPremierJourMois(TOBInterCompta, True);
  end;
  Ferme(QAcomptes);
  Result := TOBAcompte;
end;

function ChargeLesEcheances(NatCou, Champ, GroupBy, Trie: string; Rdate: boolean; Grp: string): TOB;
var SQLEcheance: string;
  QEcheance: TQuery;
  TOBEcheance, TobTmp: TOB;
  Cpt: Integer;
  Date: string;
begin
  Result := nil;
  AvecEcheance := False;
  if (TOBInterCompta.Detail.count = 0) or (TOBInterCompta.FindFirst(['GIC_NATUREPIECEG'], [NatCou], True) = nil) then exit;
  SQLEcheance := 'SELECT ' + Champ + ' ' +
    'GPE_TIERS AS GIC_TIERS , ' +
    'GPE_NATUREPIECEG AS GIC_NATUREPIECEG, ' +
    'GPE_MODEPAIE AS GIC_MODEPAIE, ' +
    'GPE_DEVISE AS GIC_DEVISE, ' +
    'GPE_TAUXDEV AS GIC_TAUXDEVISE, ' +
    'SUM(GPE_MONTANTECHE) AS GIC_MONTANTTTC, ' +
    'SUM(GPE_MONTANTDEV) AS GIC_MONTANTTTCDEV, ' +
    'GP_ETABLISSEMENT AS GIC_ETABLISSEMENT, ' +
    'GPE_DATEECHE AS GIC_DATEECHE, ' +
    'GP_REGIMETAXE AS GIC_REGIMETAXE, ' +
    'GP_TVAENCAISSEMENT AS GIC_TVAENCAISSEMENT, ' +
    'GP_VENTEACHAT AS GIC_VENTEACHAT, ' +
    'GP_DATEPIECE AS GIC_DATE ' +
    'FROM PIEDECHE ' +
    'LEFT JOIN PIECE ON ' +
    'GP_NATUREPIECEG=GPE_NATUREPIECEG AND GP_SOUCHE=GPE_SOUCHE AND ' +
    'GP_NUMERO=GPE_NUMERO AND GP_INDICEG=GPE_INDICEG ' +
    'WHERE GPE_MONTANTECHE > 0 AND GP_NATUREPIECEG IN ("' + NatCou + '") AND ' +
    //  'WHERE GP_NATUREPIECEG IN ("' + NatCou + '") AND ' +
  'GP_DATEPIECE>="' + USDateTime(StrToDate(DateDeb)) + '" AND ' +
    'GP_DATEPIECE<="' + USDateTime(StrToDate(DateFin)) + '" AND ' +
    'GP_ETATCOMPTA IN (' + Comptabilise + ') AND ' +
    'GP_ETABLISSEMENT IN (' + Etab + ') ' +
    'GROUP BY ' + GroupBy + ' ' +
    'GPE_TIERS,GPE_NATUREPIECEG,GPE_MODEPAIE,GPE_DEVISE,GPE_TAUXDEV, ' +
    'GP_ETABLISSEMENT,GPE_DATEECHE,GP_REGIMETAXE,GP_TVAENCAISSEMENT,GP_VENTEACHAT,GP_DATEPIECE';
  QEcheance := OpenSQL(SQLEcheance, True);
  if not QEcheance.Eof then
  begin
    TOBEcheance := TOB.Create('_ECHEANCE', nil, -1);
    TOBEcheance.LoadDetailDB('_ECHEANCE', '', '', QEcheance, False);
    for Cpt := 0 to TobEcheance.detail.count - 1 do
    begin
      TobTmp := TobEcheance.Detail[Cpt];
      if (Grp = 'MOI') or (Grp = 'SEM') then
      begin
        Date := DateToStr(TobTmp.GetValue('GIC_DATE'));
        TobTmp.AddChampSupValeur('ANNEE', Copy(Date, 7, 4));
        TobTmp.AddChampSupValeur('MOIS', Copy(Date, 4, 2), False);
        TobTmp.AddChampSupValeur('SEMAINE', NumSemaine(TobTmp.GetValue('GIC_DATE')), False);
      end;
      TobTmp.AddChampSupValeur('GIC_SECTANA1', '', False);
      ArrondirLesMontants(TobTmp, 'GIC_MONTANTTTC');
    end;
    if Grp = 'SEM' then
      TraitementPremierJourSemaine(TOBEcheance, True)
    else if Grp = 'MOI' then
      TraitementPremierJourMois(TOBEcheance, True);
    //AC ETAB
    TOBEcheance.Detail.Sort('' + Trie + 'GIC_TIERS;GIC_NATUREPIECEG;;GIC_ETABLISSEMENT;GIC_REGIMETAXE;GIC_DATEECHE');
    {Regroupement du PARPIECE}
    if TOBInterCompta.FieldExists('SEMAINE') then
      TraitementPremierJourSemaine(TOBEcheance, True)
    else
      if TOBInterCompta.FieldExists('MOIS') then
      TraitementPremierJourMois(TOBEcheance, True);
    Result := TOBEcheance;
  end;
  Ferme(QEcheance);
end;

function RegroupeAcomptesEcheance(TobAcomptes, TobEcheances: TOB): TOB;
var TobRegroupement: TOB;
  Cpt: integer;
  {Qte: double;}

  procedure ChercheEtAjoute(TobAtraiter: TOB; Titre: string; Acompte: boolean);
  var TobTmp, TobTmp1: TOB;
  begin
    TobTmp := FindTOB(TobAtraiter, TOBIntercompta.Detail[Cpt], False, True);
    while TobTmp <> nil do
    begin
      TobTmp1 := TOB.Create(Titre, TobRegroupement, -1);
      TobTmp1.AddChampSupValeur('GIC_NATUREPIECEG', TobTmp.GetValue('GIC_NATUREPIECEG'), False);
      TobTmp1.AddChampSupValeur('GIC_REGIMETAXE', TobTmp.GetValue('GIC_REGIMETAXE'), False);
      TobTmp1.AddChampSupValeur('GIC_DEVISE', TobTmp.GetValue('GIC_DEVISE'), False);
      TobTmp1.AddChampSupValeur('GIC_TAUXDEVISE', TobTmp.GetValue('GIC_TAUXDEVISE'), False);
      TobTmp1.AddChampSupValeur('GIC_TVAENCAISSEMENT', TobTmp.GetValue('GIC_TVAENCAISSEMENT'), False);
      TobTmp1.AddChampSupValeur('GIC_VENTEACHAT', TobTmp.GetValue('GIC_VENTEACHAT'), False);
      TobTmp1.AddChampSupValeur('GIC_MODEPAIE', TobTmp.GetValue('GIC_MODEPAIE'), False);
      TobTmp1.AddChampSupValeur('GIC_DATEECHE', TobTmp.GetValue('GIC_DATEECHE'), False);
      TobTmp1.AddChampSupValeur('GIC_TIERS', TobTmp.GetValue('GIC_TIERS'), False);
      TobTmp1.AddChampSupValeur('GIC_MONTANTTTC', TobTmp.GetValue('GIC_MONTANTTTC'), False);
      TobTmp1.AddChampSupValeur('GIC_MONTANTTTCDEV', TobTmp.GetValue('GIC_MONTANTTTCDEV'), False);
      TobTmp1.AddChampSupValeur('GIC_ETABLISSEMENT', TobTmp.GetValue('GIC_ETABLISSEMENT'), False);
      if TobTmp.FieldExists('GIC_ISEREGLEMENT') then
        TobTmp1.AddChampSupValeur('GIC_ISREGLEMENT', TobTmp.GetValue('GIC_ISREGLEMENT'), False)
      else
        TobTmp1.AddChampSupValeur('GIC_ISREGLEMENT', '', False);
      if Acompte then
        TobTmp1.AddChampSupValeur('GIC_NUMECHE', 1, False)
      else
        TobTmp1.AddChampSupValeur('GIC_NUMECHE', 0, False);
      TobTmp1.AddChampSupValeur('NUMREGROUPE', Cpt, False);
      TobTmp1.AddChampSupValeur('TYPE', Titre, False);
      TobTmp := FindTOB(TobAtraiter, TOBIntercompta.Detail[Cpt], False, False);
    end;
  end;

begin
  TobRegroupement := TOB.Create('_RegroupeAcEc', nil, -1);
  {Test depuis TObInterCompta qui contient déjà les pièces}
  for Cpt := 0 to TobInterCompta.Detail.count - 1 do
  begin
    {Doit rechercher une seule fois par pièce et pas pour les lignes analytiques}
    if (TOBIntercompta.Detail[Cpt].GetValue('GIC_NUMECHE') <= 1) and (not IsLigneAna(TOBIntercompta.Detail[Cpt])) then
    begin
      TOBIntercompta.Detail[Cpt].AddChampSupValeur('ADUPLIQUER', 'X', False);
      if TobAcomptes <> nil then
        ChercheEtAjoute(TobAcomptes, 'ligne acompte', True);
      if TobEcheances <> nil then
        ChercheEtAjoute(TobEcheances, 'ligne echeance', False);
    end else
    begin
      TOBIntercompta.Detail[Cpt].AddChampSupValeur('ADUPLIQUER', '-', False);
    end;
    {Affecte un n° unique pour retrouver}
    TOBIntercompta.Detail[Cpt].AddChampSupValeur('NUMREGROUPE', Cpt, False);
  end;
  Result := TobRegroupement;
end;

procedure FusionIntercomptaAcpteEche(TobRegroupeAcEc: TOB; Grp: string);
var Cpt, NumEche, QteATraiter, NumRegroupe: integer;
  TobTmp, TobRecherche, TobDupli: TOB;
  CalcEche: boolean;
  Arr_Chps: array of string;
  Arr_Valeur: array of variant;

  function RchEcheance(Champs: array of string; Valeurs: array of variant): integer;
  var TobEcheMax: TOB;
  begin
    Result := 0;
    TobEcheMax := TOBIntercompta.FindFirst(Champs, Valeurs, True);
    while TobEcheMax <> nil do
    begin
      if TobEcheMax.GetValue('GIC_NUMECHE') > Result then
        Result := TobEcheMax.GetValue('GIC_NUMECHE');
      TobEcheMax := TOBIntercompta.FindNext(Champs, Valeurs, True);
    end;
  end;

  procedure RecalculMontantTTC(Champ1, Champ2: string; Valeur1, Valeur2: variant);
  var MtTotalGrp, MtEcheance, MtEscompte, MtRemise, QteEcheance: double;
    MtHT, MtTaxe, MtDecTaxe, TaxeEtDec: double;
    Cpt: integer;
    TobTmp1: TOB;
    LigneIdentique: boolean;
  begin
    {Cumuls Mt échéances et qtés d'échéances}
    MtEcheance := 0;
    QteEcheance := 0;
    for Cpt := 0 to TobRegroupeAcEc.detail.count - 1 do
    begin
      TobTmp1 := TobRegroupeAcEc.detail[Cpt];
      if (TobTmp1.GetValue('NUMREGROUPE') = TobTmp.GetValue('NUMREGROUPE')) and (TobTmp1.GetValue('TYPE') = 'ligne echeance') then
      begin
        MtEcheance := MtEcheance + TobTmp1.GetValue('GIC_MONTANTTTC');
        QteEcheance := QteEcheance + 1;
      end;
    end;
    {Cumuls Montants divers}
    MtEscompte := 0;
    MtRemise := 0;
    MtHT := 0;
    MtTaxe := 0;
    for Cpt := 0 to TobIntercompta.detail.count - 1 do
    begin
      TobTmp1 := TobIntercompta.detail[Cpt];
      if not IsLigneAna(TobTmp1) then
      begin
        LigneIdentique := False;
        if Champ2 <> '' then
        begin
          if (TobTmp1.GetValue(Champ1) = Valeur1) and (TobTmp1.GetValue(Champ2) = Valeur2) then
            LigneIdentique := True;
        end else
        begin
          if (TobTmp1.GetValue(Champ1) = Valeur1) then
            LigneIdentique := True;
        end;
        if LigneIdentique then
        begin
          MtEscompte := MtEscompte + TobTmp1.GetValue('GIC_MONTANTESC');
          MtRemise := MtRemise + TobTmp1.GetValue('GIC_MONTANTREM');
          MtHT := MtHT + TobTmp1.GetValue('GIC_MONTANTHT');
          MtDecTaxe := MtDecTaxe + (TobTmp1.GetValue('GIC_MONTANTTAXEDEV1') + TobTmp1.GetValue('GIC_MONTANTTAXEDEV2') + TobTmp1.GetValue('GIC_MONTANTTAXEDEV3')
            + TobTmp1.GetValue('GIC_MONTANTTAXEDEV4') + TobTmp1.GetValue('GIC_MONTANTTAXEDEV5'));
          MtTaxe := MtTaxe + (TobTmp1.GetValue('GIC_MONTANTTAXE1') + TobTmp1.GetValue('GIC_MONTANTTAXE2') + TobTmp1.GetValue('GIC_MONTANTTAXE3') +
            TobTmp1.GetValue('GIC_MONTANTTAXE4') + TobTmp1.GetValue('GIC_MONTANTTAXE5'));
        end;
      end;
    end;
    //    MtTaxe := Arrondi(MtTaxe + MtDecTaxe,V_PGI.OkDecV);
    MtTotalGrp := MtHT + MtTaxe - MtEscompte - MtRemise;
    MtTotalGrp := Arrondi(MtTotalGrp, V_PGI.OkDecV);
    MtEcheance := Arrondi(MtEcheance, V_PGI.OkDecV);
    if QteEcheance <= 1 then
    begin
      if MtEcheance = MtTotalGrp then
      begin
        //TobTmp.PutValue('GIC_MONTANTTTC', MtTotalGrp);
        //CalcEche := False;
        TobTmp.PutValue('GIC_MONTANTTTC', 0);
        CalcEche := True;
      end else
      begin
        TobTmp.PutValue('GIC_MONTANTTTC', MtTotalGrp - MtEcheance);
        CalcEche := True;
      end;
    end else
    begin
      TobTmp.PutValue('GIC_MONTANTTTC', MtTotalGrp - MtEcheance);
      CalcEche := True;
    end;
  end;

  procedure CreerNvelleLigne(LeJournal, LeReglement: string; LEcheance: integer; DateEche: boolean);
  var Cpt1: integer;
  begin
    TobDupli := TOB.Create('INTERCOMPTA', TOBInterCompta, TOBIntercompta.Detail.count);
    TobDupli.Dupliquer(TOBIntercompta.Detail[Cpt], False, True, True);
    TobDupli.PutValue('GIC_TYPEJOURNAL', LeJournal);
    TobDupli.PutValue('GIC_ISREGLEMENT', LeReglement);
    TobDupli.PutValue('GIC_NUMECHE', LEcheance);
    if DateEche then
      TobDupli.PutValue('GIC_DATEECHE', TobRecherche.GetValue('GIC_DATEECHE'));
    TobDupli.PutValue('GIC_MODEPAIE', TobRecherche.GetValue('GIC_MODEPAIE'));
    TobDupli.PutValue('GIC_MONTANTTTC', TobRecherche.GetValue('GIC_MONTANTTTC'));
    TobDupli.PutValue('GIC_MONTANTHT', 0);
    for Cpt1 := 1 to 5 do
    begin
      TobDupli.PutValue('GIC_MONTANTTAXE' + IntToStr(Cpt1), 0);
      TobDupli.PutValue('GIC_MONTANTTAXEDEV' + IntToStr(Cpt1), 0);
    end;
    TobDupli.PutValue('GIC_MONTANTESC', 0);
    TobDupli.PutValue('GIC_MONTANTPF', 0);
    TobDupli.PutValue('GIC_MONTANTREM', 0);
  end;

begin
  {Nbre de ligne à traiter sauf les analytiques (des lignes sont ajoutées)}
  QteATraiter := TobInterCompta.Detail.Count - 1;
  for Cpt := 0 to QteATraiter do
  begin
    TOBIntercompta.Detail[Cpt].PutValue('GIC_TYPEJOURNAL', 'VEN');
    if not IsLigneAna(TobIntercompta.detail[Cpt]) then
    begin
      TobTmp := TOBIntercompta.Detail[Cpt];
      NumRegroupe := StrToInt(TobTmp.GetValue('NUMREGROUPE'));
      if (TobTmp.GetValue('ADUPLIQUER') = 'X') then
      begin
        if Grp = 'AUC' then
        begin
          RecalculMontantTTC('GIC_NUMERO', '', TobTmp.GetValue('GIC_NUMERO'), '');
          SetLength(Arr_Chps, 1);
          Arr_Chps[0] := 'GIC_NUMERO';
          SetLength(Arr_Valeur, 1);
          Arr_Valeur[0] := TobTmp.GetValue('GIC_NUMERO');
        end else
          if Grp = 'JOU' then
        begin
          RecalculMontantTTC('GIC_DATE', '', TobTmp.GetValue('GIC_DATE'), '');
          SetLength(Arr_Chps, 1);
          Arr_Chps[0] := 'GIC_DATE';
          SetLength(Arr_Valeur, 1);
          Arr_Valeur[0] := TobTmp.GetValue('GIC_DATE');
        end else
          if Grp = 'SEM' then
        begin
          RecalculMontantTTC('ANNEE', 'SEMAINE', TobTmp.GetValue('ANNEE'), TobTmp.GetValue('SEMAINE'));
          SetLength(Arr_Chps, 2);
          Arr_Chps[0] := 'ANNEE';
          Arr_Chps[1] := 'SEMAINE';
          SetLength(Arr_Valeur, 2);
          Arr_Valeur[0] := TobTmp.GetValue('ANNEE');
          Arr_Valeur[1] := TobTmp.GetValue('SEMAINE');
        end else
          if Grp = 'MOI' then
        begin
          RecalculMontantTTC('ANNEE', 'MOIS', TobTmp.GetValue('ANNEE'), TobTmp.GetValue('MOIS'));
          SetLength(Arr_Chps, 2);
          Arr_Chps[0] := 'ANNEE';
          Arr_Chps[1] := 'MOIS';
          SetLength(Arr_Valeur, 2);
          Arr_Valeur[0] := TobTmp.GetValue('ANNEE');
          Arr_Valeur[1] := TobTmp.GetValue('MOIS');
        end;
        NumEche := RchEcheance(Arr_Chps, Arr_Valeur);
        {Recherche les acomptes et échéances liés à la ligne en cours}
        TobRecherche := TobRegroupeAcEc.FindFirst(['NUMREGROUPE'], [NumRegroupe], True);
        while TobRecherche <> nil do
        begin
          {Acompte ou règlement}
          if TobRecherche.GetValue('TYPE') = 'ligne acompte' then
          begin
            CreerNvelleLigne('BTQ', TobRecherche.GetValue('GIC_ISREGLEMENT'), 1, False);
          end else
            {Echéance}
            if (TobRecherche.GetValue('TYPE') = 'ligne echeance') and (CalcEche) then
          begin
            NumEche := NumEche + 1;
            CreerNvelleLigne('VEN', '', NumEche, True);
          end;
          TobRecherche := TobRegroupeAcEc.FindNext(['NUMREGROUPE'], [TobTmp.GetValue('NUMREGROUPE')], True);
        end;
      end else
      begin
        TobTmp.PutValue('GIC_MONTANTTTC', 0);
        TobTmp.PutValue('GIC_MONTANTTTCDEV', 0);
      end;
    end;
  end;
end;

function IsLigneAna(TobL: TOB): boolean;
begin
  Result := False;
  if (TobL.GetValue('GIC_SECTANA1') = '') and (TobL.GetValue('GIC_SECTANA2') = '') and
    (TobL.GetValue('GIC_SECTANA3') = '') and (TobL.GetValue('GIC_SECTANA4') = '') and
    (TobL.GetValue('GIC_SECTANA5') = '') then
    Result := False
  else
    Result := True;
end;

//****************************************************************************//

//********************************** Sans Acompte ****************************//
// Chargement des lignes articles et OP

procedure ChargeLesLigneArticleEtOP(NatCou, Champ, GroupBy, Trie: string);
var SQLLigne: string;
  QLigne: TQuery;
begin
  // Pas de regroupement au niveau famille article et article pour la gestion des opération de caisses
  SQLLigne := 'select gl_datepiece as gic_date,gl_souche as gic_souche, gl_numero as gic_numero, ' +
    'gl_indiceg as gic_indice, gl_tiers as gic_tiers ,gl_naturepieceg as gic_naturepieceg, ' +
    'gl_devise as gic_devise,gl_tauxdev as gic_tauxdevise, ' +
    'sum(round(gl_montantttc,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttc, '+
    'sum(round(gl_montantttcdev,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttcdev, ' +
    'gl_etablissement as gic_etablissement ,gl_regimetaxe as gic_regimetaxe, ga_comptaarticle as gic_comptaarticle, ' +
    'ga_article as gic_article,ga_Typearticle from ligne ' +
    'left join piece on gp_naturepieceg=gl_naturepieceg and gp_souche=gl_souche ' +
    'and gp_numero=gl_numero and gp_indiceg=gl_indiceg ' +
    'left join article on ga_article=gl_article ' +
    'where gl_naturepieceg in ("' + NatCou + '") and ' +
    'gl_datepiece>="' + USDateTime(StrToDate(DateDeb)) + '" and gl_datepiece<="' + USDateTime(StrToDate(DateFin)) + '" ' +
    'and gp_etatcompta in (' + Comptabilise + ') and gl_typeligne<>"COM" and gl_etablissement in (' + Etab + ') ' +
    'group by gl_datepiece,gl_souche,gl_numero,gl_indiceg,gl_tiers,gl_naturepieceg,gl_devise,gl_tauxdev,gl_etablissement,gl_regimetaxe,ga_comptaarticle,ga_article,ga_typearticle';
  QLigne := OpenSQL(SQLLigne, True);
  if not QLigne.Eof then
  begin
    TOBLigne.LoadDetailDB('_LIGNE', '', '', QLigne, False);
    //if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBLigne,True)
    //   else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBLigne,True) ;
    // AC ETAB
    TOBLigne.Detail.Sort('Gic_Date;gic_souche;gic_numero;gic_indice;gic_tiers;gic_naturepieceg;gic_etablissement;gic_regimetaxe;gic_montantttc');
  end;
  ferme(QLigne);
end;

procedure ChargeLesModesPaie(NatCou, Champ, GroupBy, Trie, Regroupement: string; Rdate: boolean);
Const Egal=0 ; MontantArticleSup=1; MontantEcheanceSup=2 ;
var SQLEcheance: string;
  QEcheance: TQuery;
  i, j, CompPrix, NumChampLigne, NumChampEcheance: integer;
  TOBEcheance, TOBL, TOBE, TOBInterTemp, TOBI: TOB;
  DeltaPrix, DeltaPrixDev, DeltaEch, DeltaEchDev,
  PrixLigne,PrixLigneDev, PrixEch,PrixEchDev: Double;
  IncI, IncJ, Rupture, TraitePrix: Boolean;
begin
  TOBInterTemp := nil;
  TOBEcheance := TOB.Create('_ECHEANCE', nil, -1);
  if (TOBInterCompta.Detail.count = 0) or (TOBInterCompta.FindFirst(['GIC_NATUREPIECEG'], [NatCou], True) = nil) then exit;
  // Traitement des articles et articles financier.
  SQLEcheance := 'select gpe_datepiece as gic_date,gpe_souche as gic_souche, gpe_numero as gic_numero,gpe_indiceg as gic_indice, gpe_tiers as gic_tiers ,' +
    'gpe_naturepieceg as gic_naturepieceg,gpe_ModePaie as Gic_ModePaie,' +
    'gpe_devise as gic_devise,gpe_tauxdev as gic_tauxdevise,gpe_dateeche as gic_dateeche, ' +
    'sum(round(gpe_montanteche,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttc, '+
    'sum(round(gpe_montantdev,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttcdev, ' +
    'gp_etablissement as gic_etablissement  ' +
    'from PiedEche ' +
    'left join piece on gp_naturepieceg=gpe_naturepieceg and gp_souche=gpe_souche ' +
    'and gpe_montanteche <> 0 and gp_numero=gpe_numero and gp_indiceg=gpe_indiceg ' +
    'where gp_naturepieceg in ("' + NatCou + '") and ' +
    'gp_datepiece>="' + USDateTime(StrToDate(DateDeb)) + '" and gp_datepiece<="' + USDateTime(StrToDate(DateFin)) + '" ' +
    'and gp_etatcompta in (' + Comptabilise + ') and gp_etablissement in (' + Etab + ') ' +
    'group by gpe_datepiece,gpe_souche,gpe_numero,gpe_indiceg, gpe_tiers,gpe_naturepieceg,gpe_ModePaie,gpe_devise,gpe_tauxdev,gpe_dateeche,gp_etablissement';
  QEcheance := OpenSQL(SQLEcheance, True);
  if not QEcheance.EOF then
  begin
    TOBEcheance.LoadDetailDB('_ECHEANCE', '', '', QEcheance, False);
    //if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBEcheance,True)
    //  else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBEcheance,True) ;
    // AC ETAB
    TOBEcheance.Detail.Sort('Gic_Date;gic_souche;gic_numero;gic_indice;gic_tiers;gic_naturepieceg;gic_etablissement;gic_regimetaxe;gic_montantttc');
    InitMove(TOBLigne.Detail.count, '');
    TOBInterTemp := TOB.Create('_INTERTEMP', nil, -1);
    i := 0;
    j := 0;
    // Pour utiliser le GetValeur
    NumChampLigne := TOBLigne.detail[0].GetNumChamp('GIC_MONTANTTTC');
    NumChampEcheance := TOBEcheance.detail[0].GetNumChamp('GIC_MONTANTTTC');
    while (i <= TOBLigne.Detail.count - 1) and (TOBEcheance.Detail.count>0) do
    begin
      TOBL := TOBLigne.Detail[i];
      TOBE := TOBEcheance.Detail[0];
      CompPrix := ComparePrix(TOBL, TOBE);
      case CompPrix of
      Egal :
        begin
          TOBI := TOB.Create('_TEMP', TOBInterTemp, TOBInterTemp.Detail.count);
          CreerChamps(TOBI, TOBL, TOBE, TOBL.GetValeur(NumChampLigne), TOBL.GetValue('GIC_MONTANTTTCDEV'), True, Regroupement);
          TOBEcheance.Detail[0].Free ;
        end ;
      MontantEcheanceSup :
        begin
          TOBI := TOB.Create('_TEMP', TOBInterTemp, TOBInterTemp.Detail.count);
          PrixLigne := Arrondi(TOBL.GetValeur(NumChampLigne), V_PGI.OkDecV);
          PrixEch := Arrondi(TOBE.GetValeur(NumChampEcheance), V_PGI.OkDecV);
          PrixLigneDev := Arrondi(TOBL.GetValue('GIC_MONTANTTTCDEV'), V_PGI.OkDecV);
          PrixEchDev := Arrondi(TOBE.GetValue('GIC_MONTANTTTCDEV'), V_PGI.OkDecV);
          CreerChamps(TOBI, TOBL, TOBE, PrixLigne, PrixLigneDev, True, Regroupement);
          TOBE.PutValue('GIC_MONTANTTTC',PrixEch-PrixLigne) ;
          TOBE.PutValue('GIC_MONTANTTTCDEV',PrixEchDev-PrixLigneDev) ;
        end ;
      MontantArticleSup :
        begin
          TOBI := TOB.Create('_TEMP', TOBInterTemp, TOBInterTemp.Detail.count);
          PrixLigne := Arrondi(TOBL.GetValeur(NumChampLigne), V_PGI.OkDecV);
          PrixEch := Arrondi(TOBE.GetValeur(NumChampEcheance), V_PGI.OkDecV);
          PrixLigneDev := Arrondi(TOBL.GetValue('GIC_MONTANTTTCDEV'), V_PGI.OkDecV);
          PrixEchDev := Arrondi(TOBE.GetValue('GIC_MONTANTTTCDEV'), V_PGI.OkDecV);
          DeltaPrix := PrixLigne - PrixEch ;
          DeltaPrixDev := PrixLigneDev - PrixEchDev ;
          CreerChamps(TOBI, TOBL, TOBE, PrixEch, PrixEchDev, True, Regroupement);
          While (DeltaPrix > 0) and (TOBEcheance.Detail.count>1) do
          begin
            TOBEcheance.Detail[0].Free ;
            TOBE := TOBEcheance.Detail[0] ;
            PrixEch := Arrondi(TOBE.GetValeur(NumChampEcheance), V_PGI.OkDecV) ;
            PrixEchDev := Arrondi(TOBE.GetValue('GIC_MONTANTTTCDEV'), V_PGI.OkDecV);
            TOBI := TOB.Create('_TEMP', TOBInterTemp, TOBInterTemp.Detail.count);
            if DeltaPrix < PrixEch then CreerChamps(TOBI, TOBL, TOBE, DeltaPrix, DeltaPrixDev, True, Regroupement)
            else CreerChamps(TOBI, TOBL, TOBE, PrixEch, PrixEchDev, True, Regroupement);
            DeltaPrix := DeltaPrix - PrixEch ;
            DeltaPrixDev := DeltaPrixDev - PrixEchDev ;
          end ;
          If DeltaPrix < 0 then TOBE.PutValue('GIC_MONTANTTTC',ABS(DeltaPrix)) ;
          If DeltaPrix = 0 then TOBEcheance.Detail[0].Free ;
        end ;
      end ;
      i := i + 1;
    end;
  end;
  FiniMove;
  Ferme(QEcheance);
  TOBEcheance.Free;
  if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBInterTemp, True)
  else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBInterTemp, True);
  //TOBInterTemp.Detail.Sort('-ga_typearticle;gic_modepaie;gic_article'+Trie+'gic_tiers;gic_naturepieceg;gic_regimetaxe') ;
  // AC ETAB
  TOBInterTemp.Detail.Sort('gic_date;-ga_typearticle;gic_modepaie;gic_article;gic_tiers;gic_naturepieceg;gic_etablissement;gic_regimetaxe');
 RegroupementModePaie(TOBInterTemp, Regroupement);
end;
//****************************************************************************//

procedure RegroupementModePaie(TOBInterTemp: TOB; Regroupement: string);
var i: integer;
  TOBL, NewTOB: TOB;
  MontantTTC, MontantTTCDEV: Double;
begin
  i := 0;
  if Regroupement = 'AUC' then
  begin
    while (i <= TOBInterTemp.Detail.count - 1) do
    begin
      MontantTTC := 0.0;
      MontantTTCDEV := 0.0;
      TOBL := TOBInterTemp.Detail[i];
      MontantTTC := TOBL.GetValue('GIC_MONTANTTTC');
      MontantTTCDEV := TOBL.GetValue('GIC_MONTANTTTCDEV');
      while (not Rupture(TOBL, TOBInterTemp, i, 'AUC'))
        and ((TOBL.GetValue('GA_TYPEARTICLE') <> 'FI')
        or (TOBL.GetValue('GA_TYPEARTICLE') <> 'MAR'))
        and (i < TOBInterTemp.Detail.count - 1) do
      begin
        i := i + 1;
        TOBL := TOBInterTemp.Detail[i];
        MontantTTC := MontantTTC + TOBL.GetValue('GIC_MONTANTTTC');
        MontantTTCDEV := MontantTTCDEV + TOBL.GetValue('GIC_MONTANTTTCDEV');
      end;
      NewTOB := TOB.Create('INTERCOMPTA', TOBInterCompta, TOBInterCompta.Detail.count);
      CreerChamps(NewTob, TOBL, nil, MontantTTC, MontantTTCDEV, False, regroupement);
      i := i + 1;
    end;
  end
  else if Regroupement = 'JOU' then
  begin
    while (i <= TOBInterTemp.Detail.count - 1) do
    begin
      MontantTTC := 0.0;
      MontantTTCDEV := 0.0;
      TOBL := TOBInterTemp.Detail[i];
      MontantTTC := TOBL.GetValue('GIC_MONTANTTTC');
      MontantTTCDEV := TOBL.GetValue('GIC_MONTANTTTCDEV');
      while (not Rupture(TOBL, TOBInterTemp, i, 'JOU'))
        and ((TOBL.GetValue('GA_TYPEARTICLE') <> 'FI')
        or (TOBL.GetValue('GA_TYPEARTICLE') <> 'MAR'))
        and (i < TOBInterTemp.Detail.count - 1) do
      begin
        i := i + 1;
        TOBL := TOBInterTemp.Detail[i];
        MontantTTC := MontantTTC + TOBL.GetValue('GIC_MONTANTTTC');
        MontantTTCDEV := MontantTTCDEV + TOBL.GetValue('GIC_MONTANTTTCDEV');
      end;
      NewTOB := TOB.Create('INTERCOMPTA', TOBInterCompta, TOBInterCompta.Detail.count);
      CreerChamps(NewTob, TOBL, nil, MontantTTC, MontantTTCDEV, False, Regroupement);
      i := i + 1;
    end;
  end
  else if Regroupement = 'SEM' then
  begin
    while (i <= TOBInterTemp.Detail.count - 1) do
    begin
      MontantTTC := 0.0;
      MontantTTCDEV := 0.0;
      TOBL := TOBInterTemp.Detail[i];
      MontantTTC := TOBL.GetValue('GIC_MONTANTTTC');
      MontantTTCDEV := TOBL.GetValue('GIC_MONTANTTTCDEV');
      while (not Rupture(TOBL, TOBInterTemp, i, 'SEM'))
        and ((TOBL.GetValue('GA_TYPEARTICLE') <> 'FI')
        or (TOBL.GetValue('GA_TYPEARTICLE') <> 'MAR'))
        and (i < TOBInterTemp.Detail.count - 1) do
      begin
        i := i + 1;
        TOBL := TOBInterTemp.Detail[i];
        MontantTTC := MontantTTC + TOBL.GetValue('GIC_MONTANTTTC');
        MontantTTCDEV := MontantTTCDEV + TOBL.GetValue('GIC_MONTANTTTCDEV');
      end;
      NewTOB := TOB.Create('INTERCOMPTA', TOBInterCompta, TOBInterCompta.Detail.count);
      CreerChamps(NewTob, TOBL, nil, MontantTTC, MontantTTCDEV, False, regroupement);
      i := i + 1;
    end;
  end
  else if Regroupement = 'MOI' then
  begin
    while (i <= TOBInterTemp.Detail.count - 1) do
    begin
      MontantTTC := 0.0;
      MontantTTCDEV := 0.0;
      TOBL := TOBInterTemp.Detail[i];
      MontantTTC := TOBL.GetValue('GIC_MONTANTTTC');
      MontantTTCDEV := TOBL.GetValue('GIC_MONTANTTTCDEV');
      while (not Rupture(TOBL, TOBInterTemp, i, 'MOI'))
        and ((TOBL.GetValue('GA_TYPEARTICLE') <> 'FI')
        or (TOBL.GetValue('GA_TYPEARTICLE') <> 'MAR'))
        and (i < TOBInterTemp.Detail.count - 1) do
      begin
        i := i + 1;
        TOBL := TOBInterTemp.Detail[i];
        MontantTTC := MontantTTC + TOBL.GetValue('GIC_MONTANTTTC');
        MontantTTCDEV := MontantTTCDEV + TOBL.GetValue('GIC_MONTANTTTCDEV');
      end;
      NewTOB := TOB.Create('INTERCOMPTA', TOBInterCompta, TOBInterCompta.Detail.count);
      CreerChamps(NewTob, TOBL, nil, MontantTTC, MontantTTCDEV, False, regroupement);
      i := i + 1;
    end;
  end;
end;

function Rupture(TOBL, TOBSuiv: TOB; indice: Integer; regroupement: string): boolean;
var IsRupture: Boolean;
begin
  Result := False;
  IsRupture := False;
  if indice + 1 > TOBSuiv.Detail.count - 1 then exit;
  // Rupture dû au MP, devise, taxe.....
  if (TOBL.GetValue('GIC_MODEPAIE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_MODEPAIE'))
    or (TOBL.GetValue('GIC_REGIMETAXE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_REGIMETAXE'))
    or (TOBL.GetValue('GIC_DEVISE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_DEVISE'))
    or (TOBL.GetValue('GIC_DATEECHE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_DATEECHE'))
    or (TOBL.GetValue('GIC_TAUXDEVISE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_TAUXDEVISE'))
    or (TOBL.GetValue('GIC_ETABLISSEMENT') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_ETABLISSEMENT'))
    then IsRupture := True;
  // Rupture dû au regroupement de date
  if Regroupement = 'AUC' then
  begin
    if (TOBL.GetValue('GIC_NATUREPIECEG') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_NATUREPIECEG'))
      or (TOBL.GetValue('GIC_SOUCHE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_SOUCHE'))
      or (TOBL.GetValue('GIC_NUMERO') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_NUMERO'))
      or (TOBL.GetValue('GIC_INDICE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_INDICE'))
      or (TOBL.GetValue('GIC_DATE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_DATE'))
      or (TOBL.GetValue('GIC_TIERS') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_TIERS'))
      then IsRupture := True;
  end
  else if (TOBL.GetValue('GIC_NATUREPIECEG') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_NATUREPIECEG'))
    or (TOBL.GetValue('GIC_DATE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_DATE'))
    or (TOBL.GetValue('GIC_TIERS') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_TIERS'))
    then IsRupture := True;
  if TOBL.GetValue('GIC_COMPTAARTICLE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_COMPTAARTICLE')
    then IsRupture := True;
  if (TOBL.GetValue('GA_TYPEARTICLE') = 'FI') and (TOBL.GetValue('GIC_ARTICLE') <> TOBSuiv.Detail[indice + 1].GetValue('GIC_ARTICLE'))
    then IsRupture := True;
  Result := IsRupture;
end;

procedure CreerChamps(TOBInter, TOBLig, TobEch: TOB; Prix, PrixDev: Double; Virtuelle: boolean; Regroupement: string);
var Date: string;
begin
  if Virtuelle then
  begin
    TOBInter.Dupliquer(TOBLig, False, True, True);
    TOBInter.PutValue('GIC_MONTANTTTC', Prix);
    TOBInter.PutValue('GIC_MONTANTTTCDEV', PrixDev);
    TOBInter.AddChampSupValeur('GIC_MODEPAIE', TOBEch.GetValue('GIC_MODEPAIE'));
    TOBInter.AddChampSupValeur('GIC_DATEECHE', TOBEch.GetValue('GIC_DATEECHE'));
    if Regroupement = 'SEM' then
    begin
      Date := DateToStr(TOBLig.GetValue('GIC_DATE'));
      TOBInter.AddChampSupValeur('SEMAINE', NumSemaine(TOBLig.GetValue('GIC_DATE')));
      TOBInter.AddChampSupValeur('ANNEE', Copy(Date, 7, 4));
    end
    else if Regroupement = 'MOI' then
    begin
      Date := DateToStr(TOBLig.GetValue('GIC_DATE'));
      TOBInter.AddChampSupValeur('MOIS', Copy(Date, 4, 2));
      TOBInter.AddChampSupValeur('ANNEE', Copy(Date, 7, 4));
    end;
  end else
  begin
    TOBInter.AddChampSupValeur('GIC_DATE', TOBLig.GetValue('GIC_DATE'));
    TOBInter.AddChampSupValeur('GIC_TIERS', TOBLig.GetValue('GIC_TIERS'));
    TOBInter.AddChampSupValeur('GIC_REGIMETAXE', TOBLig.GetValue('GIC_REGIMETAXE'));
    TOBInter.AddChampSupValeur('GIC_ETABLISSEMENT', TOBLig.GetValue('GIC_ETABLISSEMENT'));
    TOBInter.AddChampSupValeur('GIC_NATUREPIECEG', TOBLig.GetValue('GIC_NATUREPIECEG'));
    TOBInter.AddChampSupValeur('GIC_MONTANTTTC', Prix);
    TOBInter.AddChampSupValeur('GIC_DEVISE', TOBLig.GetValue('GIC_DEVISE'));
    TOBInter.AddChampSupValeur('GIC_TAUXDEVISE', TOBLig.GetValue('GIC_TAUXDEVISE'));
    TOBInter.AddChampSupValeur('GIC_MONTANTTTCDEV', PrixDev);
    TOBInter.AddChampSupValeur('GIC_MODEPAIE', TOBLig.GetValue('GIC_MODEPAIE'));
    TOBInter.AddChampSupValeur('GIC_DATEECHE', TOBLig.GetValue('GIC_DATEECHE'));
    if TOBLig.GetValue('ga_typearticle') = 'FI' then
    begin
      TOBInter.AddChampSupValeur('GIC_ARTICLE', TOBLig.GetValue('GIC_ARTICLE'));
      TOBInter.AddChampSupValeur('GIC_COMPTAARTICLE', TOBLig.GetValue('GIC_COMPTAARTICLE'));
      TOBInter.AddChampSupValeur('GIC_ISOPCAISSE', 'X');
    end;
    if Regroupement = 'SEM' then
    begin
      Date := DateToStr(TOBLig.GetValue('GIC_DATE'));
      TOBInter.AddChampSupValeur('SEMAINE', NumSemaine(TOBLig.GetValue('GIC_DATE')));
      TOBInter.AddChampSupValeur('ANNEE', Copy(Date, 7, 4));
    end
    else if Regroupement = 'MOI' then
    begin
      Date := DateToStr(TOBLig.GetValue('GIC_DATE'));
      TOBInter.AddChampSupValeur('MOIS', Copy(Date, 4, 2));
      TOBInter.AddChampSupValeur('ANNEE', Copy(Date, 7, 4));
    end;
  end;
end;

function LesTobEgales(TOBLig, TOBEch: TOB): Boolean;
begin
  Result := False;
  if (TOBLig.GetValue('GIC_NATUREPIECEG') = TOBEch.GetValue('GIC_NATUREPIECEG'))
    and (TOBLig.GetValue('GIC_SOUCHE') = TOBEch.GetValue('GIC_SOUCHE'))
    and (TOBLig.GetValue('GIC_NUMERO') = TOBEch.GetValue('GIC_NUMERO'))
    and (TOBLig.GetValue('GIC_INDICE') = TOBEch.GetValue('GIC_INDICE'))
    and (TOBLig.GetValue('GIC_DATE') = TOBEch.GetValue('GIC_DATE'))
    and (TOBLig.GetValue('GIC_TIERS') = TOBEch.GetValue('GIC_TIERS')) then
    Result := True;
end;

function ComparePrix(TOBLig, TOBEch: TOB): Integer;
var Res: integer;
begin
  Res := -1;
  // Une echeance pour un article
  if Abs(TOBLig.GetValue('GIC_MONTANTTTC')) = Abs(TOBEch.GetValue('GIC_MONTANTTTC'))
    then Res := 0
    // Plusieurs echeance pour un article
  else if Abs(TOBLig.GetValue('GIC_MONTANTTTC')) > Abs(TOBEch.GetValue('GIC_MONTANTTTC'))
    then Res := 1
    // Une echeance pour plusieurs articles
  else if Abs(TOBLig.GetValue('GIC_MONTANTTTC')) < Abs(TOBEch.GetValue('GIC_MONTANTTTC'))
    then Res := 2;
  result := Res;
end;

// Recherche d'un TOB suivant des critères (ligne sans analytique)

function FindTOB(TobRech, TOBOrig: Tob; RDate, First: Boolean): Tob;
var TOBL: Tob;
  Souche, Tiers, Naturepieceg, RegimeTaxe, Etablissement: string;
  Date: TDateTime;
  Numero, Indice: Integer;
begin
  Numero := 0;
  Indice := 0;
  if TOBOrig.FieldExists('GIC_SOUCHE') then Souche := TOBOrig.GetValue('GIC_SOUCHE');
  if TOBOrig.FieldExists('GIC_NUMERO') then Numero := TOBOrig.GetValue('GIC_NUMERO');
  if TOBOrig.FieldExists('GIC_INDICE') then Indice := TOBOrig.GetValue('GIC_INDICE');
  Date := TOBOrig.GetValue('GIC_DATE');
  Tiers := TOBOrig.GetValue('GIC_TIERS');
  NaturePieceg := TOBOrig.GetValue('GIC_NATUREPIECEG');
  RegimeTaxe := TOBOrig.GetValue('GIC_REGIMETAXE');
  // AC ETAB
  Etablissement := TOBOrig.GetValue('GIC_ETABLISSEMENT');
  if RDate then
  begin
    if First then
      //      TOBL := TobRech.FindFirst(['GIC_Date', 'Gic_tiers', 'Gic_nature', 'Gic_Etablissement', 'gic_regimetaxe', 'gic_sectana1'], [Date, Tiers, NaturePieceg, Etablissement, RegimeTaxe, ''], False)
      TOBL := TobRech.FindFirst(['GIC_Date', 'Gic_tiers', 'Gic_naturepieceg', 'Gic_Etablissement', 'gic_regimetaxe', 'gic_sectana1'], [Date, Tiers,
        NaturePieceg, Etablissement, RegimeTaxe, ''], False)
    else
      //    TOBL := TobRech.FindNext(['GIC_Date', 'Gic_tiers', 'Gic_nature', 'Gic_Etablissement', 'gic_regimetaxe', 'gic_sectana1'], [Date, Tiers, NaturePieceg, Etablissement, RegimeTaxe, ''], False);
      TOBL := TobRech.FindNext(['GIC_Date', 'Gic_tiers', 'Gic_naturepieceg', 'Gic_Etablissement', 'gic_regimetaxe', 'gic_sectana1'], [Date, Tiers,
        NaturePieceg, Etablissement, RegimeTaxe, ''], False);
  end else
  begin
    if First then
      //      TOBL := TobRech.FindFirst(['GIC_Date', 'gic_souche', 'gic_numero', 'gic_indice', 'Gic_tiers', 'Gic_nature', 'Gic_Etablissement', 'gic_regimetaxe', 'gic_sectana1'], [Date,
      TOBL := TobRech.FindFirst(['GIC_Date', 'gic_souche', 'gic_numero', 'gic_indice', 'Gic_tiers', 'Gic_naturepieceg', 'Gic_Etablissement', 'gic_regimetaxe',
        'gic_sectana1'], [Date,
        Souche, Numero, Indice, Tiers, NaturePieceg, Etablissement, RegimeTaxe, ''], False)
    else
      //      TOBL := TobRech.FindNext(['GIC_Date', 'gic_souche', 'gic_numero', 'gic_indice', 'Gic_tiers', 'Gic_nature', 'Gic_Etablissement', 'gic_regimetaxe', 'gic_sectana1'], [Date,
      TOBL := TobRech.FindNext(['GIC_Date', 'gic_souche', 'gic_numero', 'gic_indice', 'Gic_tiers', 'Gic_naturepieceg', 'Gic_Etablissement', 'gic_regimetaxe',
        'gic_sectana1'], [Date,
        Souche, Numero, Indice, Tiers, NaturePieceg, Etablissement, RegimeTaxe, ''], False);
  end;
  Result := TOBL;
end;

// Remplis GP_DATEPIECE par le premier jour d'une semaine

procedure TraitementPremierJourSemaine(TOBCourante: TOB; Piece: Boolean);
var i: Integer;
  DateCalc: TDateTime;
  TOBL: TOB;
begin
  for i := 0 to TOBCourante.Detail.Count - 1 do
  begin
    TOBL := TOBCourante.Detail[i];
    DateCalc := PremierJourSemaine(StrToInt(TOBL.GetValue('SEMAINE')), StrToInt(TOBL.GetValue('ANNEE')));
    if not TOBL.FieldExists('GIC_DATE') then
      TOBL.AddChampSup('GIC_DATE', False);
    TOBL.PutValue('GIC_DATE', DateCalc);
  end;
end;

// Remplis GP_DATEPIECE par le premier jour d'un mois

procedure TraitementPremierJourMois(TOBCourante: TOB; Piece: Boolean);
var i: Integer;
  Date: TDateTime;
  TOBL: TOB;
  StrDate: string;
begin
  for i := 0 to TOBCourante.Detail.Count - 1 do
  begin
    TOBL := TOBCourante.Detail[i];
    StrDate := '01/' + IntToStr(TOBL.GetValue('MOIS')) + '/' + IntToStr(TOBL.GetValue('ANNEE'));
    Date := StrToDate(StrDate);
    // Dernière date du mois dans la date comptable
    Date := FinDeMois(Date);
    if Piece then
    begin
      if not TOBL.FieldExists('GIC_DATE') then
        TOBL.AddChampSup('GIC_DATE', False);
      TOBL.PutValue('GIC_DATE', Date);
    end
    else TOBL.PutValue('GIC_DATE', Date);
  end;
end;

procedure AjoutChampsTob(quoi: string);
begin
  TOBPiece.AddChampSupValeur('TAXEPORT1', 0, True);
  TOBPiece.AddChampSupValeur('TAXEPORT2', 0, True);
  if (quoi = 'SEM') or (quoi = 'MOI') then
  begin
    TOBInterCompta.AddChampSupValeur('ANNEE', '', True);
    TOBPiece.AddChampSupValeur('ANNEE', '', True);
  end;
  if quoi = 'SEM' then
  begin
    TOBInterCompta.AddChampSupValeur('SEMAINE', '', True);
    {IFDEF V545
    TOBInterCompta.AddChampSupValeur('MOIS','',True) ;
    TOBPiece.AddChampSupValeur('MOIS','',True) ;
    ENDIF}
    TOBPiece.AddChampSupValeur('SEMAINE', '', True);
  end;
  if quoi = 'MOI' then
  begin
    TOBInterCompta.AddChampSupValeur('MOIS', '', True);
    TOBPiece.AddChampSupValeur('MOIS', '', True);
  end;
end;

{=========================================================================================}
{============================= Port et frais    ==========================================}
{=========================================================================================}

procedure EcrireLignePortEtFrais;
var i: Integer;
  TOBL: TOB;
  Port, TTC, HT, Taxe1, Taxe2, TaxeP1, TaxeP2: Double;
  PortDEV, TTCDEV, HTDEV, Taxe1DEV, Taxe2DEV, TaxeP1DEV, TaxeP2DEV: Double;
begin
  for i := 0 to TobInterCompta.Detail.Count - 1 do
  begin
    TOBL := TobInterCompta.Detail[i];
    if TOBL.GetValue('GIC_MONTANTPF') > 0 then
    begin
      TTC := TOBL.GetValue('GIC_MONTANTTTC');
      TTCDev := TOBL.GetValue('GIC_MONTANTTTCDEV');
      HT := TOBL.GetValue('GIC_MONTANTHT'); // Montant HT piece sans frais et port
      HTDEV := TOBL.GetValue('GIC_MONTANTHTDEV'); // Montant HT piece sans frais et port
      TaxeP1 := TOBL.GetValue('GIC_MONTANTTAXE1'); // Montant taxe piece sans frais et port
      //      TaxeP1DEV := TOBL.GetValue('GIC_MONTANTTAXEDEV1'); // Montant taxe piece sans frais et port
      TaxeP2 := TOBL.GetValue('GIC_MONTANTTAXE2'); // Montant taxe piece sans frais et port
      //      TaxeP2DEV := TOBL.GetValue('GIC_MONTANTTAXEDEV2'); // Montant taxe piece sans frais et port
      Port := TOBL.GetValue('GIC_MONTANTPF'); // Valeur Port en HT
      PortDEV := TOBL.GetValue('GIC_MONTANTPFDEV'); // Valeur Port en HT
      Taxe1 := TOBL.GetValue('TAXEPORT1'); // Valeur taxe 1 Port
      Taxe1Dev := TOBL.GetValue('TAXEPORTDEV1'); // Valeur taxe 1 Port
      Taxe2 := TOBL.GetValue('TAXEPORT2'); // Valeur taxe 2 Port
      Taxe2Dev := TOBL.GetValue('TAXEPORTDEV2'); // Valeur taxe 2 Port
      TTC := TTC + Port + Taxe1 + Taxe2;
      TTCDEV := TTCDEV + PortDEV + Taxe1DEV + Taxe2DEV;
      TOBL.PutValue('GIC_MONTANTTTC', TTC);
      TOBL.PutValue('GIC_MONTANTTTCDEV', TTCDEV);
      TaxeP1 := Taxe1 + TaxeP1;
      TaxeP2 := Taxe2 + TaxeP2;
      TaxeP1DEV := Taxe1DEV + TaxeP1DEV;
      TaxeP2DEV := Taxe2DEV + TaxeP2DEV;
      TOBL.PutValue('GIC_MONTANTTAXE1', TaxeP1); // Montant taxe piece avec port
      TOBL.PutValue('GIC_MONTANTTAXE2', TaxeP2); // Montant taxe piece avec port
      //      TOBL.PutValue('GIC_MONTANTTAXEDEV1', TaxeP1Dev); // Montant taxe piece avec port
      //      TOBL.PutValue('GIC_MONTANTTAXEDEV2', TaxeP2Dev); // Montant taxe piece avec port
    end;
  end;
end;

{=========================================================================================}
{============================= AXES ANALYTIQUES ==========================================}
{=========================================================================================}
// Traitement des chargements des axes analytiques

procedure TraitementDesAxesAnalytiques(NatCou, Champ, GroupBy, Trie: string; IsOperationCaisse: Boolean = False);
var VenteAchat, CodeAxe, EtabTot, EtabTrait: string;
  i: Integer;
  TOBL: TOB;
  AxeOk: Boolean;
begin
  VenteAchat := GetInfoParPiece(NatCou, 'GPP_VENTEACHAT');
  if IsOperationCaisse then VenteAchat := 'ACH';
  ConstruireTOBAna(VenteAchat);
  i := 1;
  while i < 5 do
  begin
    EtabTot := EtabOr;
    while EtabTot <> '' do
    begin
      AxeOk := False;
      EtabTrait := ReadTokenSt(EtabTot);
      CodeAxe := 'A' + IntToStr(i);
      // Chercher Axe analytique défini dans les opération de caisse
      if IsOperationCaisse and ExisteSQL('SELECT GFC_ARTICLE FROM ARTFINANCIERCOMPL WHERE GFC_ETABLISSEMENT="'
        + EtabTrait + '" AND GFC_SECTION' + IntToStr(i) + '<>""') then
      begin
        ChargeLesAxeAnalytiqueOperationCaisse(NatCou, Champ, GroupBy, Trie, CodeAxe, EtabTrait);
        AxeOk := ChercheAxeAnalytiqueDesOP(CodeAxe, Etab);
      end;
      if AxeOk then continue;
      TOBL := TOBAna.FindFirst(['GDA_AXE'], [CodeAxe], False);
      if TOBL = nil then break;
      ChampSecAna := GetChampSecAxeAnalytique(EtabTrait, CodeAxe);
      //if ChampSecAna='' then exit ;
      ChampLibAna := GetChampLibAxeAnalytique(EtabTrait, CodeAxe);
      ConstruireGroupByAna;
      ConstruireLibelleAna;
      if IsOperationCaisse then
      begin
        ChargeLesAxeAnalytiqueOperationCaisse(NatCou, Champ, GroupBy, Trie, CodeAxe, EtabTrait);
        // Envoie TOBAna avec valeur util pour construction des axes
        PutAxeAnalytique(CodeAxe, True);
      end
      else ChargeLesAxeAnalytique(NatCou, Champ, GroupBy, Trie, CodeAxe, EtabTrait);
    end;
    Inc(i)
  end;
end;

procedure ConstruireTOBAna(VenteAchat: string);
var SQLAna: string;
  i: Integer;
  QAna: TQuery;
begin
  if TOBAna <> nil then
  begin
    TOBAna.Free;
    TOBAna := nil;
  end;
  TOBAna := TOB.Create('_TOBANA', nil, -1);
  SQLAna := 'select GDA_AXE,GDA_TYPECOMPTE,GDA_ETABLISSEMENT,GDA_ISLIBELLE,GDA_LONGUEUR,GDA_LIBCHAMP,GDA_TYPESTRUCTANA from decoupeana where GDA_TYPECOMPTE="'
    + VenteAchat + '" order by GDA_AXE,GDA_TYPECOMPTE,GDA_ETABLISSEMENT,GDA_TYPESTRUCTANA';
  QAna := OpenSQL(SQLAna, True);
  if not QAna.EOF then TOBAna.LoadDetailDB('_Ana', '', '', QAna, True);
  Ferme(QAna);
  for i := 0 to TOBAna.Detail.count - 1 do
  begin
    if TOBAna.Detail[i].GetValue('GDA_ETABLISSEMENT') = NULL then TOBAna.Detail[i].PutValue('GDA_ETABLISSEMENT', '');
  end;
end;

function GetChampSecAxeAnalytique(Etablissement, CodeAxe: string): string;
var ChampSecAna: string;
  TOBL: Tob;
begin
  ChampSecAna := '';
  TOBL := TOBAna.FindFirst(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'SEC', Etablissement], False);
  if TOBL <> nil then
    while TOBL <> nil do
    begin
      if Etablissement = TOBL.GetValue('GDA_ETABLISSEMENT') then
      begin
        if (Pos(TOBL.GetValue('GDA_LIBCHAMP'), ChampSecAna) = 0) and (Pos('Constante', TOBL.GetValue('GDA_LIBCHAMP')) = 0) then
        begin
          if ChampSecAna = '' then ChampSecAna := +TOBL.GetValue('GDA_LIBCHAMP') + ''
          else ChampSecAna := ChampSecAna + ';' + TOBL.GetValue('GDA_LIBCHAMP') + '';
        end;
      end;
      TOBL := TOBAna.FindNext(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'SEC', Etablissement], False);
    end
  else
  begin
    TOBL := TOBAna.FindFirst(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'SEC', ''], False);
    if TOBL <> nil then
      while TOBL <> nil do
      begin
        if (Pos(TOBL.GetValue('GDA_LIBCHAMP'), ChampSecAna) = 0) and (Pos('Constante', TOBL.GetValue('GDA_LIBCHAMP')) = 0) then
        begin
          if ChampSecAna = '' then ChampSecAna := +TOBL.GetValue('GDA_LIBCHAMP') + ''
          else ChampSecAna := ChampSecAna + ';' + TOBL.GetValue('GDA_LIBCHAMP') + '';
        end;
        TOBL := TOBAna.FindNext(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'SEC', ''], False);
      end;
  end;
  Result := ChampSecAna;
end;

function GetChampLibAxeAnalytique(Etablissement, CodeAxe: string): string;
var ChampLibAna: string;
  TOBL: Tob;
begin
  ChampLibAna := '';
  TOBL := TOBAna.FindFirst(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'LSE', Etablissement], False);
  if TOBL <> nil then
    while TOBL <> nil do
    begin
      if Etablissement = TOBL.GetValue('GDA_ETABLISSEMENT') then
      begin
        if (Pos(TOBL.GetValue('GDA_LIBCHAMP'), ChampLibAna) = 0) and (Pos(TOBL.GetValue('GDA_LIBCHAMP'), ChampSecAna) = 0) and (Pos('Constante',
          TOBL.GetValue('GDA_LIBCHAMP')) = 0) then
        begin
          if ChampLibAna = '' then ChampLibAna := +TOBL.GetValue('GDA_LIBCHAMP') + ''
          else ChampLibAna := ChampLibAna + ';' + TOBL.GetValue('GDA_LIBCHAMP') + '';
        end;
      end;
      TOBL := TOBAna.FindNext(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'LSE', Etablissement], False);
    end
  else
  begin
    TOBL := TOBAna.FindFirst(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'LSE', ''], False);
    if TOBL <> nil then
      while TOBL <> nil do
      begin
        if (Pos(TOBL.GetValue('GDA_LIBCHAMP'), ChampSecAna) = 0) and (Pos('Constante', TOBL.GetValue('GDA_LIBCHAMP')) = 0) then
        begin
          if ChampLibAna = '' then ChampLibAna := +TOBL.GetValue('GDA_LIBCHAMP') + ''
          else ChampLibAna := ChampLibAna + ';' + TOBL.GetValue('GDA_LIBCHAMP') + '';
        end;
        TOBL := TOBAna.FindNext(['GDA_AXE', 'GDA_TYPESTRUCTANA', 'GDA_ETABLISSEMENT'], [CodeAxe, 'LSE', ''], False);
      end;
  end;
  Result := ChampLibAna;
end;

// Recupère les info pour faire le regroupement pour les axes analytiques

procedure ConstruireGroupByAna;
var ChampLoc, Champ: string;
begin
  ChampSecAnaReq := '';
  ChampLoc := ChampSecAna;
  while ChampLoc <> '' do
  begin
    Champ := ReadTokenSt(ChampLoc);
    if ChampSecAnaReq = '' then ChampSecAnaReq := ',' + Champ + ' '
    else ChampSecAnaReq := ChampSecAnaReq + ',' + '' + Champ + ' ';
  end;
end;

// Recupère les info pour récuperer les informations pour les libelles des axes analytiques

procedure ConstruireLibelleAna;
var ChampLoc, Champ: string;
begin
  ChampLibAnaReq := '';
  ChampLoc := ChampLibAna;
  while ChampLoc <> '' do
  begin
    Champ := ReadTokenSt(ChampLoc);
    if ChampLibAnaReq = '' then ChampLibAnaReq := ',MAX(' + Champ + ') As ' + Champ + ' '
    else ChampLibAnaReq := ChampLibAnaReq + ',MAX(' + '' + Champ + ') As ' + Champ + ' ';
  end;
end;

// Chargement des axes analytiques

procedure ChargeLesAxeAnalytique(NatCou, Champ, GroupBy, Trie, CodeAxe, Etab: string);
var SQLAxeAna, ChampAxe, Temp: string;
  QAxeAna: TQuery;
  i: Integer;
  Ok: Boolean;
begin
  SQLAxeAna :=
    'SELECT ' + Champ +
    'GL_ETABLISSEMENT AS GIC_ETABLISSEMENT, GL_TIERS AS GIC_TIERS, ' +
    'GL_NATUREPIECEG AS GIC_NATUREPIECEG, GL_REGIMETAXE AS GIC_REGIMETAXE, ' +
    'GL_FAMILLETAXE1 AS GIC_FAMILLETAXE1, GL_FAMILLETAXE2 AS GIC_FAMILLETAXE2, ' +
    'GL_FAMILLETAXE3 AS GIC_FAMILLETAXE3, GL_FAMILLETAXE4 AS GIC_FAMILLETAXE4, ' +
    'GL_FAMILLETAXE5 AS GIC_FAMILLETAXE5, ' +
    'GA_COMPTAARTICLE AS GIC_COMPTAARTICLE, T_COMPTATIERS AS GIC_COMPTATIERS, ' +
    'SUM(GL_QTEFACT) AS GIC_QTE1, GL_DEVISE AS GIC_DEVISE, ' +
    'SUM(ROUND(GL_TOTALTAXE1,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXE1, '+
    'SUM(ROUND(GL_TOTALTAXEDEV1,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXEDEV1,' +
    'SUM(ROUND(GL_TOTALTAXE2,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXE2, '+
    'SUM(ROUND(GL_TOTALTAXEDEV2,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXEDEV2, ' +
    'SUM(ROUND(GL_TOTALTAXE3,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXE3, '+
    'SUM(ROUND(GL_TOTALTAXEDEV3,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXEDEV3, ' +
    'SUM(ROUND(GL_TOTALTAXE4,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXE4, '+
    'SUM(ROUND(GL_TOTALTAXEDEV4,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXEDEV4, ' +
    'SUM(ROUND(GL_TOTALTAXE5,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXE5, '+
    'SUM(ROUND(GL_TOTALTAXEDEV5,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTAXEDEV5,' +
    'SUM(ROUND(GL_MONTANTHT,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTHT, '+
    'SUM(ROUND(GL_MONTANTHTDEV,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTHTDEV, ' +
    'SUM(ROUND(GL_MONTANTTTC,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTTC, '+
    'SUM(ROUND(GL_MONTANTTTCDEV,'+IntToStr(V_PGI.OkDecV)+')) AS GIC_MONTANTTTCDEV ' +
    ChampSecAnaReq + '' + ChampLibAnaReq + '' +
    'FROM PIECE LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG AND GL_SOUCHE=GP_SOUCHE ' +
    'AND GL_NUMERO=GP_NUMERO AND GL_INDICEG=GP_INDICEG ' +
    'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE ' +
    'LEFT JOIN TIERS ON T_TIERS=GL_TIERS ' +
    'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=GL_ETABLISSEMENT '+
    'WHERE GP_NATUREPIECEG IN ("' + NatCou + '") ' +
    'AND GP_DATEPIECE>="' + UsDateTime(StrToDate(DateDeb)) + '" ' +
    'AND GP_DATEPIECE<="' + UsDateTime(StrToDate(DateFin)) + '" ' +
    'AND GP_ETABLISSEMENT="' + Etab + '" ' +
    'AND GP_ETATCOMPTA IN (' + Comptabilise + ') ' +
    'AND GL_TYPELIGNE<>"COM" ' +
    'AND GL_TYPEARTICLE<>"FI" ' +
    'GROUP BY ' + GroupBy + 'GL_ETABLISSEMENT, GL_TIERS, GL_NATUREPIECEG, GL_REGIMETAXE,' +
    'GL_FAMILLETAXE1, GL_FAMILLETAXE2, GL_FAMILLETAXE3, GL_FAMILLETAXE4, GL_FAMILLETAXE5,' +
    'GA_COMPTAARTICLE, GL_DEVISE ' + ChampSecAnaReq + ', T_COMPTATIERS ';
  QAxeAna := OpenSQL(SQLAxeAna, True);
  if not QAxeAna.Eof then
  begin
    TOBInterCompta.LoadDetailDB('INTERCOMPTA', '', '', QAxeAna, True);
    if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBInterCompta, False)
    else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBInterCompta, False);
    // AC ETAB
    TOBInterCompta.Detail.Sort('' + Trie + 'gic_tiers;gic_naturepieceg;gic_etablissement;gic_regimetaxe');
  end else
  begin
    Ferme(QAxeAna);
    exit;
  end;
  Ferme(QAxeAna);
  for i := 0 to TOBInterCompta.Detail.count - 1 do
  begin
    Temp := ChampSecAna;
    Ok := True;
    while Temp <> '' do
    begin
      ChampAxe := ReadTokenSt(Temp);
      if (ChampAxe = '') or (not TOBInterCompta.Detail[i].FieldExists(ChampAxe)) then Ok := False;
    end;
    if Ok then
    begin
      TOBInterCompta.Detail[i].PutValue('GIC_SECTAN' + CodeAxe, CodeAxe);
      TOBInterCompta.Detail[i].PutValue('GIC_LIBAN' + CodeAxe, CodeAxe);
    end;
  end;
  // Envoie TOBAna avec valeur util pour construction des axes
  PutAxeAnalytique(CodeAxe, False);
end;

// Chargement des axes analytiques pour les opération de caisse

procedure ChargeLesAxeAnalytiqueOperationCaisse(NatCou, Champ, GroupBy, Trie, CodeAxe, Etab: string);
var SQLAxeAna, ChampAxe, Temp: string;
  QAxeAna: TQuery;
  i: Integer;
  Ok: Boolean;
begin
  SQLAxeAna := 'select ' + Champ + ' gpe_tiers as gic_tiers ,gpe_naturepieceg as gic_naturepieceg,gpe_ModePaie as Gic_ModePaie,' +
    'gpe_devise as gic_devise,gpe_tauxdev as gic_tauxdevise, ' +
    //'sum(gpe_montanteche) as gic_montantttc, sum(gpe_montantdev) as gic_montantttcdev, '+
    'sum(round(gl_montantttc,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttc, '+
    'sum(round(gl_montantttcdev,'+IntToStr(V_PGI.OkDecV)+')) as gic_montantttcdev,' +
    'gp_etablissement as gic_etablissement ,gl_regimetaxe as gic_regimetaxe, ga_comptaarticle as gic_comptaarticle, ' +
    'ga_article as gic_article ' + ChampSecAnaReq + '' + ChampLibAnaReq + '' +
    'from piece left join ligne on gl_naturepieceg=gp_naturepieceg and gl_souche=gp_souche ' +
    'and gl_numero=gp_numero and gl_indiceg=gp_indiceg ' +
    'left join article on ga_article=gl_article ' +
    'left join etabliss on et_etablissement=gl_etablissement ' +
    'left join piedeche on gpe_naturepieceg=gp_naturepieceg and gpe_souche=gp_souche ' +
    'and gpe_numero=gp_numero and gpe_indiceg=gp_indiceg and ' +
    'gl_naturepieceg=gp_naturepieceg and gl_souche=gp_souche and ' +
    'gl_numero=gp_numero and gl_indiceg=gp_indiceg where gpe_montanteche > 0 and gp_naturepieceg in ("' + NatCou + '") and ' +
    'gp_datepiece>="' + USDateTime(StrToDate(DateDeb)) + '" and gp_datepiece<="' + USDateTime(StrToDate(DateFin)) + '" ' +
    'and gp_etatcompta in (' + Comptabilise + ') and gp_etablissement = "' + Etab + '" and ga_typearticle="FI"' +
    'group by ' + GroupBy + ' gpe_tiers,gpe_naturepieceg,gpe_ModePaie,gpe_devise,gpe_tauxdev,gp_etablissement , ' +
    'gl_regimetaxe,ga_comptaarticle,ga_article ' + ChampSecAnaReq + ' ';
  QAxeAna := OpenSQL(SQLAxeAna, True);
  if not QAxeAna.Eof then
  begin
    TOBInterCompta.LoadDetailDB('INTERCOMPTA', '', '', QAxeAna, True);
    if TOBInterCompta.FieldExists('SEMAINE') then TraitementPremierJourSemaine(TOBInterCompta, False)
    else if TOBInterCompta.FieldExists('MOIS') then TraitementPremierJourMois(TOBInterCompta, False);
    //TOBInterCompta.Detail.Sort(''+Trie+'gic_tiers;gic_naturepieceg;gic_regimetaxe') ;
  end else
  begin
    Ferme(QAxeAna);
    exit;
  end;
  Ferme(QAxeAna);
  for i := 0 to TOBInterCompta.Detail.count - 1 do
  begin
    Temp := ChampSecAna;
    Ok := True;
    while Temp <> '' do
    begin
      ChampAxe := ReadTokenSt(Temp);
      if (ChampAxe = '') or (not TOBInterCompta.Detail[i].FieldExists(ChampAxe)) or (TOBInterCompta.Detail[i].GetValue('GIC_MODEPAIE') = '') then Ok := False;
    end;
    if Ok then
    begin
      TOBInterCompta.Detail[i].PutValue('GIC_SECTAN' + CodeAxe, CodeAxe);
      TOBInterCompta.Detail[i].PutValue('GIC_LIBAN' + CodeAxe, CodeAxe);
    end;
  end;
end;

// Chargement des Axes analytique dans la table InterCompta

procedure ChargerValeurAxeAna(TOBTempAna: Tob; CodeAxe: string; ModePaie: Boolean);
var j: Integer;
  TOBL, TOBA: TOB;
begin
  if ModePaie then // Axe analytique que sur les opérations de caisse avec montant négatif
  begin
    for j := TOBInterCompta.Detail.count - 1 downto 0 do
    begin
      TOBL := TOBInterCompta.Detail[j];
      if (TOBL.GetValue('GIC_SECTAN' + CodeAxe) = CodeAxe) and (TOBL.GetValue('GIC_MODEPAIE') = '') then
        TOBL.Free
      else if (TOBL.GetValue('GIC_SECTAN' + CodeAxe) = CodeAxe) and (TOBL.GetValue('GIC_MODEPAIE') <> '')
        and ((TOBL.GetValue('GIC_MONTANTTTC') > 0) or (TOBL.GetValue('GIC_MONTANTTTC') = 0)) then TOBL.Free;
    end;
  end;
  for j := 0 to TOBTempAna.Detail.count - 1 do
  begin
    TOBA := TOBTempAna.Detail[j];
    TOBL := TOBInterCompta.FindFirst(['GIC_SECTAN' + CodeAxe], [CodeAxe], False);
    if TOBL <> nil then
    begin
      if (TOBL.GetValue('GIC_ETABLISSEMENT') = TOBA.GetValue('GDA_ETABLISSEMENT')) or (TOBA.GetValue('GDA_ETABLISSEMENT') = '') then
      begin
        if TOBL.GetValue('GIC_SECTAN' + CodeAxe) = CodeAxe then TOBL.PutValue('GIC_SECTAN' + CodeAxe, TOBA.GetValue('GDA_VALEUR1'));
        if TOBL.GetValue('GIC_LIBAN' + CodeAxe) = CodeAxe then TOBL.PutValue('GIC_LIBAN' + CodeAxe, TOBA.GetValue('GDA_VALEUR2'));
      end;
    end;
  end;
end;

// Traitement des axes analytiques

procedure PutAxeAnalytique(CodeAxe: string; ModePaie: Boolean);
var i, Ind: integer;
  Champ: string;
  TOBL, TOBA, TOBTempAna: Tob;
begin
  TOBTempAna := TOB.Create('_TOBTEMP', nil, -1);
  Ind := 0;
  for i := 0 to TOBInterCompta.Detail.count - 1 do
  begin
    TOBL := TOBInterCompta.Detail[i];
    {Recherche avec un établissement spécifique}
    TOBA := TOBAna.FindFirst(['GDA_AXE', 'GDA_ETABLISSEMENT'], [CodeAxe, TOBL.GetValue('GIC_ETABLISSEMENT')], False);
    if TOBA <> nil then
    begin
      while TOBA <> nil do
      begin
        Champ := TOBA.GetValue('GDA_LIBCHAMP');
        if (Pos('Constante', Champ) <> 0) then
        begin
          TOBA := TOBAna.FindNext(['GDA_AXE', 'GDA_ETABLISSEMENT'], [CodeAxe, ''], False);
          Continue;
        end;

        TOB.Create('_', TOBTempAna, Ind);
        TOBTempAna.Detail[Ind].Dupliquer(TOBA, False, True);
        TOBTempAna.Detail[Ind].AddChampSup('VALEUR', False);
        if (TOBL.GetValue('GIC_SECTAN' + CodeAxe) = CodeAxe) and (TOBL.GetValue('GIC_LIBAN' + CodeAxe) = CodeAxe) then
        begin
          TOBTempAna.Detail[Ind].PutValue('VALEUR', TOBL.GetValue(Champ));
          Inc(Ind);
        end
        else TOBTempAna.Detail[Ind].Free;
        TOBA := TOBAna.FindNext(['GDA_AXE', 'GDA_ETABLISSEMENT'], [CodeAxe, TOBL.GetValue('GIC_ETABLISSEMENT')], False);
      end;
    end else
      {Recherche avec tous les établissements}
    begin
      TOBA := TOBAna.FindFirst(['GDA_AXE', 'GDA_ETABLISSEMENT'], [CodeAxe, ''], False);
      if TOBA <> nil then
      begin
        while TOBA <> nil do
        begin
          Champ := TOBA.GetValue('GDA_LIBCHAMP');
          if (Pos('Constante', Champ) <> 0) then
          begin
            TOBA := TOBAna.FindNext(['GDA_AXE', 'GDA_ETABLISSEMENT'], [CodeAxe, ''], False);
            Continue;
          end;
          TOB.Create('_', TOBTempAna, Ind);
          TOBTempAna.Detail[Ind].Dupliquer(TOBA, False, True);
          TOBTempAna.Detail[Ind].AddChampSup('VALEUR', False);
          if (TOBL.GetValue('GIC_SECTAN' + CodeAxe) = CodeAxe) and (TOBL.GetValue('GIC_LIBAN' + CodeAxe) = CodeAxe) then
          begin
            TOBTempAna.Detail[Ind].PutValue('VALEUR', TOBL.GetValue(Champ));
            Inc(Ind);
          end
          else TOBTempAna.Detail[Ind].Free;
          TOBA := TOBAna.FindNext(['GDA_AXE', 'GDA_ETABLISSEMENT'], [CodeAxe, ''], False);
        end;
      end;
    end;
  end;
  TOBTempAna := ValeurAxeAna(TOBTempAna);
  ChargerValeurAxeAna(TOBTempAna, CodeAxe, ModePaie);
  TOBTempAna.Free;
end;

// Recherche des section analytiques définie dans les complément des opération de caisse

function ChercheAxeAnalytiqueDesOP(CodeAxe, Etab: string): Boolean;
var QOPCompl: TQuery;
  SQLCompl, LibSection: string;
  TOBOPCompl, TOBL, TOBC: TOB;
  i: Integer;
begin
  Result := False;
  SQLCompl := 'SELECT * FROM ARTFINANCIERCOMPL WHERE GFC_ETABLISSEMENT IN (' + Etab + ') AND GFC_SECTION' + Copy(CodeAxe, 2, 1) + '<>""';
  QOPCompl := OpenSQL(SQLCompl, True);
  if not QOPCompl.EOF then
  begin
    TOBOPCompl := TOB.Create('_ARTFINANCIERCOMPL', nil, -1);
    TOBOPCompl.LoadDetailDB('ARTFINANCIERCOMPL', '', '', QOPCompl, True);
  end else
  begin
    Ferme(QOPCompl);
    exit;
  end;
  for i := TOBInterCompta.Detail.count - 1 downto 0 do
  begin
    TOBL := TOBInterCompta.Detail[i];
    if (TOBL.GetValue('GIC_ARTICLE') <> '') and (TOBL.GetValue('GIC_MODEPAIE') <> '') and (TOBL.GetValue('GIC_SECTAN' + CodeAxe) = CodeAxe) then
    begin
      TOBC := TOBOPCompl.FindFirst(['GFC_ARTICLE'], [TOBL.GetValue('GIC_ARTICLE')], True);
      if TOBC <> nil then
      begin
        if (TOBL.GetValue('GIC_MONTANTTTC') > 0) or (TOBL.GetValue('GIC_MONTANTTTC') = 0) then
        begin
          TOBL.Free;
          continue;
        end;
        LibSection := RechDom('TZSECTION', TOBC.GetValue('GFC_SECTION' + Copy(CodeAxe, 2, 1)), False);
        TOBL.PutValue('GIC_SECTAN' + CodeAxe, TOBC.GetValue('GFC_SECTION' + Copy(CodeAxe, 2, 1)));
        TOBL.PutValue('GIC_LIBAN' + CodeAxe, LibSection);
        Result := True;
      end else
      begin
        TOBL.Free;
        //Result:=False ;
      end;
    end;
  end;
  TOBOPCompl.Free;
end;

end.
