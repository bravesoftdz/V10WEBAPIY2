{***********UNITE*************************************************
Auteur  ...... : Julien D
Créé le ...... : 22/03/2002
Modifié le ... : 24/05/2002
Description .. : Export de la table InterCompta vers un fichier ASCII au
Suite ........ : format CEGID
Mots clefs ... : EXPORT;COMPTA
*****************************************************************}
unit MBOExportCompta;

interface

uses StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  Db, DbTables,
  {$ENDIF}
  Forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM, UTob, UtilPGI,
  M3FP, Ent1, FactCpta, FactUtil, EntGC, EXPORTASCII_ENTETE_TOM, ed_tools, ParamSoc,
  SaisUtil, HStatus, UtilIntercompta, UlibEcriture,FactTob;

var TOBTiers, TOBCpta, TOBPiece, TOBRegroupeParNature, TOBModePaie, TOBAxe: TOB;
  TOBTVAEtabl, TOBModePaieEtabl, TOBOperationCaisseEtab, TOBOperationCaisse, TOBLignesAnalytiques, TOBFinale, TOBEcartCaisse: TOB;
  TobEcriture, TOBAnalytique, TOBDesGeneraux, TobJournaux, TobExportTiers: TOB;
  NumErreur: Integer;
  ModePaiementDefaut, ArticleFondCaisse, ArticleEcartCaisse, ArticleVersementArrhe, TypeEcriture: string;
  LeCptAuxiliaire, LeCptGeneral, LaNatureTiers, LeTypeTiers, ContrePartieGen: string;
  AuxiCentralisation, GestionOpCaisse, GestionAcompte: boolean;
  DetEntreprise, RegrEntreprise, CentEntreprise: boolean;
  DetParticulier, RegrParticulier, CentParticulier: boolean;
  PaiemDiffere, RgtBanque, RgtTiers: boolean;
  Rapport: boolean;
  RapportLst_: TListBox;
  CptEcritureVen, CptEcritureRgt, QteEcr: integer;
  TobRegRgt, TobRegrParPiece_, TobPieceEqui_: TOB;

procedure InitLesTOB;
procedure DetruitLesTOB;
function ExportCompta(var LastError: integer; MonEcran: TForm; OptionsRegr: string = ''; RapportFinTrt: boolean = False; RapportLst: TListBox = nil;
  TobRegrParPiece: TOB = nil; TobPieceEqui: TOB = nil): Boolean;
function TraiteLignesInterCompta(QLigneCompta: TQuery; Nature: string; IsRegl: boolean): T_RetCpta;
procedure InitTOBLigneExportCompta(TOBLignes: TOB);
function RecupValeursLibEcr(QLC: TQuery; ValeurEcr: string): string;
procedure AlimRefInterneEtLibelle(TOBL: TOB; QLC: TQuery; LibDefaut, Nature: string; TypeEcr: T_TypeEcr; LibEcr: Boolean);
procedure ChargeTOBLigneCommun(TOBLignes: TOB; QLC: TQuery; NatureG: string; IsRegl: boolean);
function ChargeLignesTiers(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne, NumEche: integer): T_RetCpta;
function ChargeLignesHT(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
function ChargeLignesTaxes(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
function ChargeLigneRemise(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
function ChargeLigneEscompte(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
function ChargeLignesModePaie(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
function ChargeLignesFraisPorts(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
function ChargeLignesAnalytiques(TOBLignes: TOB; QLC: TQuery; Nature: string; ModePaie: Boolean; IsRegl: boolean): T_RetCpta;
procedure ChargeTiersEnTOB(Q: TQuery);
function ChargeTOBTiers(T_Tiers: string): TOB;
function ChargeTOBPiece(CleDoc: R_CleDoc): TOB;
function ChargeCompteVente(QLC: TQuery): TOB;
function ChargeModePaie(ModePaie: string): TOB;
procedure RegrouperEtEcrireLeFichier(NaturePiece: string);
function ExportTOBLignes(OptionsRegr: string): boolean;
procedure EcrireUneLigne(TOBL: TOB);
procedure EcrireLigneTiers(Q: TQuery);
procedure EcrireLigneSectionAna(Code, Libelle, Axe: string);
procedure ChargeModePaiementDefaut;
procedure ChargeArticlesFinancier;
function IsArtFinancier(Article, TypeArt: string): Boolean;
procedure ChargeTOBEcartCaisse(Etab: string);
function EquilibrageDesComptes(QLC: TQuery): Double; //Taxes1, Taxes2, Taxes3, Taxes4, Taxes5, HT, TTC, Remise, Escompte, PortEtFrais : double): Double;
procedure GestionEcartDansLesJournaux(TypeJournal: string);
procedure LigneEcartJournaux(TotalCredit, TotalDebit, TotalCreditDev, TotalDebitDev: Double; TOBL: Tob; TypeJournal: string);
// Charge les axes anatytique de charge en fonction du ParamSoc SO_GCAXEANALYTIQUE=FALSE
procedure ChargeAxeAnalytique(Srang, Nature: string; QLC: TQuery);
//Procédures permettant d'écrire dans les tables ECRITURE et ANALYTIQ
procedure TestRegroupements(OptionsRegr: string);
procedure AddTobEcriture(TobL: TOB; NvellePce: boolean; TobAna: TOB);
procedure AddTobAnalytique(TobL: TOB; NvellePce: boolean; NumLigne: integer);
procedure CalcDebCred(TobL, TobAAffecter: TOB; Prefixe: string; TobEcr: TOB; InfosDevise: RDevise);
procedure InfosCompl(IsRegl: boolean; TobAAffecter: TOB; LeCptAuxi, LeCptGene, TypeLigne, LaNatTiers: string; NumLigne: integer; Rib, LeTypTiers: string);
procedure MajRapport;
procedure RecalculNumLignes(TobATraiter: TOB);
function RegroupmentEcriture(TobL, TobEcr, TobAna: TOB): boolean;

type
  MiseAJour = class
    procedure MajFiles;
  end;

implementation

procedure InitLesTOB;
var TobTmp: TOB;
begin
  TobExportTiers := TOB.Create('', nil, -1);
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBPiece := TOB.Create('PIECE', nil, -1);
  TOBCpta := CreerTOBCpta;
  TOBFinale := TOB.Create('Toutes les lignes', nil, -1);
  TOBRegroupeParNature := TOB.Create('Lignes pour une nature', nil, -1);
  TOBModePaie := TOB.Create('MODEPAIE', nil, -1);
  TOBTVAEtabl := TOB.Create('TXCPTTVACOMPL', nil, -1);
  TOBModePaieEtabl := TOB.Create('MODEPAIECOMPL', nil, -1);
  TOBOperationCaisseEtab := TOB.Create('ARTFINANCIERCOMPL', nil, -1);
  TOBOperationCaisse := TOB.Create('_ARTFINANCIER', nil, -1);
  TOBLignesAnalytiques := TOB.Create('Lignes Analytiques', nil, -1);
  TOBEcartCaisse := TOB.Create('Ecart caisse', nil, -1);
  // Axes analytiques de compte
  TOBAxe := TOB.Create('VENTIL', nil, -1);
  //Charge la table TXCPTTVACOMPL en TOB
  TOBTVAEtabl.LoadDetailDB('TXCPTTVACOMPL', '', '', nil, True);
  //Charge la table MODEPAIECOMPL en TOB
  TOBModePaieEtabl.LoadDetailDB('MODEPAIECOMPL', '', '', nil, True);
  //Charge la table ARTFINANCIERCOMPL en TOB
  TOBOperationCaisseEtab.LoadDetailDB('ARTFINANCIERCOMPL', '', '', nil, True);
  TobEcriture := TOB.Create('Maj ecritures', nil, -1);
  TobTmp := TOB.Create('TABLE ECRITURE', TobEcriture, -1);
  TobTmp := TOB.Create('TABLE ANALYTIQUE', TobEcriture, -1);
  TOBAnalytique := TOB.Create('Maj analytique', nil, -1);
  TOBDesGeneraux := TOB.Create('GENERAUX', nil, -1);
  TOBDesGeneraux.LoadDetailDB('GENERAUX', '', '', nil, True);
  TobJournaux := TOB.Create('JOURNAL', nil, -1);
  TobJournaux.LoadDetailDB('JOURNAL', '', '', nil, True);
  TobRegRgt := TOB.Create('Regroupement reglement', nil, -1);
end;

procedure DetruitLesTOB;
begin
  TOBFinale.Free;
  TOBTiers.Free;
  TobExportTiers.Free;
  TOBPiece.Free;
  TOBCpta.Free;
  TOBRegroupeParNature.Free;
  TOBModePaie.Free;
  TOBTVAEtabl.Free;
  TOBModePaieEtabl.Free;
  TOBLignesAnalytiques.Free;
  TOBAxe.Free;
  TOBOperationCaisseEtab.Free;
  TOBOperationCaisse.Free;
  TOBEcartCaisse.Free;
  TobEcriture.Free;
  TOBAnalytique.Free;
  TOBDesGeneraux.Free;
  TobJournaux.Free;
  TobRegRgt.Free;
end;

procedure TraiteLignesTiers;
var Q_Tiers: TQuery;
  StSQL, StSQL1, StSQL2: string;
begin
  StSQL1 := 'SELECT DISTINCT T_TIERS,T_CONSO,T_AUXILIAIRE,T_LIBELLE,T_NATUREAUXI,' +
    'T_COLLECTIF,T_EAN,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL,T_VILLE,' +
    'R_DOMICILIATION,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB,' +
    'T_PAYS,T_ABREGE,T_LANGUE,T_MULTIDEVISE,T_DEVISE,T_TELEPHONE,' +
    'T_FAX,T_REGIMETVA,T_MODEREGLE,T_COMMENTAIRE,T_NIF,T_SIRET,T_APE,' +
    'C_NOM,C_SERVICE,C_FONCTION,C_TELEPHONE,C_FAX,C_TELEX,C_RVA,C_CIVILITE,C_PRINCIPAL,' +
    'T_JURIDIQUE,R_PRINCIPAL ';
  StSQL2 := 'FROM INTERCOMPTA ' +
    'LEFT JOIN TIERS ON GIC_TIERS=T_TIERS ' +
    'LEFT JOIN RIB ON T_AUXILIAIRE=R_AUXILIAIRE ' +
    'LEFT JOIN CONTACT ON T_AUXILIAIRE=C_AUXILIAIRE ' +
    'WHERE GIC_USER="' + V_PGI.User + '"';
  StSQL := StSQL1 + StSQL2;
  Q_Tiers := OpenSQL(StSQL, True);
  Q_Tiers.First;
  while not Q_Tiers.Eof do
  begin
    ChargeTiersEnTOB(Q_Tiers);
    EcrireLigneTiers(Q_Tiers);
    Q_Tiers.Next;
  end;
  Ferme(Q_Tiers);
end;

function RetourneNumChampAna(Q: TQuery): integer;
var i: integer;
  Section: string;
begin
  Result := 0;
  for i := 1 to 5 do
  begin
    Section := Q.FindField('GIC_SECTANA' + IntToStr(i)).AsString;
    if (Section <> '') and (Section <> 'A' + IntToStr(i)) then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TraiteLignesSectionAna;
var Q_SectionAna: TQuery;
  StSQL, Code, Libelle: string;
  NumChamp: integer;
begin
  StSQL := 'SELECT DISTINCT GIC_SECTANA1,GIC_SECTANA2,GIC_SECTANA3,GIC_SECTANA4,GIC_SECTANA5,' +
    'GIC_LIBANA1,GIC_LIBANA2,GIC_LIBANA3,GIC_LIBANA4,GIC_LIBANA5 ' +
    'FROM INTERCOMPTA WHERE ' +
    'GIC_USER="' + V_PGI.User + '" AND ' +
    '(GIC_SECTANA1<>"" OR GIC_SECTANA2<>"" OR GIC_SECTANA3<>"" OR ' +
    'GIC_SECTANA4<>"" OR GIC_SECTANA5<>"")';
  Q_SectionAna := OpenSQL(StSQL, True);
  Q_SectionAna.First;
  while not Q_SectionAna.Eof do
  begin
    NumChamp := RetourneNumChampAna(Q_SectionAna);
    if NumChamp > 0 then
    begin
      Code := Q_SectionAna.FindField('GIC_SECTANA' + IntToStr(NumChamp)).AsString;
      Libelle := Q_SectionAna.FindField('GIC_LIBANA' + IntToStr(NumChamp)).AsString;
      EcrireLigneSectionAna(Code, Libelle, 'A' + IntToStr(NumChamp));
    end;
    Q_SectionAna.Next;
  end;
  Ferme(Q_SectionAna);
end;

procedure ChargeModePaiementDefaut;
var QQ: TQuery;
begin
  QQ := OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="' + VH_GC.GCModeRegleDefaut + '"', True);
  if not QQ.Eof then ModePaiementDefaut := QQ.FindField('MR_MP1').AsString else ModePaiementDefaut := 'CHQ';
  Ferme(QQ);
end;

procedure ChargeArticlesFinancier;
var QArt: TQuery;
  TOBO: TOB;
begin
  QArt := OpenSQL('SELECT GA_ARTICLE, GA_TYPEARTFINAN, GA_DESIGNATION1, GA_DESIGNATION2 FROM ARTICLE WHERE GA_TYPEARTICLE="FI"', True);
  TOBOperationCaisse.LoadDetailDB('_ARTFINANCIER', '', '', QArt, False);
  TOBO := TOBOperationCaisse.FindFirst(['GA_TYPEARTFINAN'], ['FCA'], true);
  if TOBO <> nil then ArticleFondCaisse := TOBO.GetValue('GA_ARTICLE');
end;

function IsArtFinancier(Article, TypeArt: string): Boolean;
var TOBA: Tob;
begin
  Result := False;
  TOBA := TOBOperationCaisse.FindFirst(['GA_ARTICLE', 'GA_TYPEARTFINAN'], [Article, TypeArt], True);
  Result := (TOBA <> nil);
end;

procedure ChargeTOBEcartCaisse(Etab: string);
var Ind: integer;
begin
  Ind := TOBEcartCaisse.Detail.count;
  TOB.Create('_', TOBEcartCaisse, Ind);
  TOBEcartCaisse.Detail[Ind].AddChampSupValeur('_ETAB', Etab, False);
  TOBEcartCaisse.Detail[Ind].AddChampSupValeur('_ECARTCAISSE', TRUE, False);
end;

//function  ExportCompta(var LastError:integer; StCheminFichier,TypeEcriture:string; MonEcran : TForm; OptionsRegr : string=''): Boolean;

function ExportCompta(var LastError: integer; MonEcran: TForm; OptionsRegr: string = ''; RapportFinTrt: boolean = False; RapportLst:
  TListBox = nil; TobRegrParPiece: TOB = nil; TobPieceEqui: TOB = nil): Boolean;
var Q_InterCompta, Q_Count: TQuery;
  StSQL, Nature, NatureTemp: string;
  BReglement, BReglTemp: boolean;
begin
  Result := False;
  NumErreur := 0;
  Rapport := RapportFinTrt;
  RapportLst_ := RapportLst; //Liste du rapport
  TobRegrParPiece_ := TobRegrParPiece; //Tob contenant le type de regroupement du ParPiece pour chaque type de pièce
  TobPieceEqui_ := TobPieceEqui;
  CptEcritureVen := 0;
  CptEcritureRgt := 0;

  {StSQL := ' FROM INTERCOMPTA WHERE GIC_USER="' + V_PGI.User + '" ' +
     'AND GIC_SECTANA1="" AND GIC_SECTANA2="" AND GIC_SECTANA3="" ' +
     'AND GIC_SECTANA4="" AND GIC_SECTANA5="" ';
   Q_Count := OpenSQL('SELECT COUNT(*) AS Compteur' + StSQL, True);
   StSQL := StSQL + ' ORDER BY GIC_NUMERO, GIC_TIERS, GIC_TYPEJOURNAL, GIC_COMPTEUR';
   Ferme(Q_Count); }

  stSQL := ' FROM INTERCOMPTA WHERE GIC_USER="' + V_PGI.User + '" ' +
    'AND GIC_SECTANA1="" AND GIC_SECTANA2="" AND GIC_SECTANA3="" ' +
    'AND GIC_SECTANA4="" AND GIC_SECTANA5="" ';
  Q_Count := OpenSQL('SELECT COUNT(*) AS Compteur' + StSQL, True);
  StSQL := StSQL + ' ORDER BY GIC_NUMERO, GIC_TIERS, GIC_TYPEJOURNAL, GIC_DATE,GIC_MODEPAIE , GIC_COMPTEUR ';
  Ferme(Q_Count); 

  Q_InterCompta := OpenSQL('SELECT *' + StSQL, True);
  if Q_InterCompta.Eof then
  begin
    LastError := 1;
    Ferme(Q_InterCompta);
    Exit;
  end;

  InitLesTOB;
  GestionOpCaisse := GetParamSoc('SO_GESTIONOPCAISSE');
  ChargeModePaiementDefaut;
  ChargeArticlesFinancier;
  TestRegroupements(OptionsRegr);
  try
    //Creer et charge en TOB les lignes de Vente et de Règlement
    Q_InterCompta.First;
    NatureTemp := Q_InterCompta.FindField('GIC_NATUREPIECEG').AsString;
    BReglTemp := (Q_InterCompta.FindField('GIC_TYPEJOURNAL').AsString = 'BTQ');
    while not Q_InterCompta.Eof do
    begin
      if not MoveCurProgressForm('Chargement des lignes d''écritures') then
      begin
        LastError := 10;
        Break;
      end;
      Nature := Q_InterCompta.FindField('GIC_NATUREPIECEG').AsString;
      BReglement := (Q_InterCompta.FindField('GIC_TYPEJOURNAL').AsString = 'BTQ'); //VEN = vente
      if (Nature <> NatureTemp) or (BReglement <> BReglTemp) then
        RegrouperEtEcrireLeFichier(Nature);
      if TraiteLignesInterCompta(Q_InterCompta, Nature, BReglement) <> rcOk then
      begin
        LastError := NumErreur;
        Break;
      end;
      NatureTemp := Nature;
      BReglTemp := BReglement;
      Q_InterCompta.Next;
    end;
    if (Q_InterCompta.Eof) then
    begin
      if TOBRegroupeParNature.Detail.Count > 0 then
        RegrouperEtEcrireLeFichier(Nature);
      Result := True;
    end;
    // Equilibrer les comptes de vente pour intégration si ecart Debit-Crédit (Ecart conversion...)
    GestionEcartDansLesJournaux('FC');
    // Equilibrer les comptes de banques jusqu'à la gestion des multi op caisse par ticket
    GestionEcartDansLesJournaux('OD');
    if MoveCurProgressForm('Génération des lignes ...') then
      if not ExportTOBLignes(OptionsRegr) then
      begin
        LastError := 11;
        Result := False
      end;
  finally
    Ferme(Q_InterCompta);
    DetruitLesTOB;
  end;
end;

function TraiteLignesInterCompta(QLigneCompta: TQuery; Nature: string; IsRegl: boolean): T_RetCpta;
var TOBLignes: TOB;
  NumLigne, NumEche: integer;
  LigVente, OpCaisse: boolean;
begin
  GestionAcompte := (GetInfoParPiece(Nature, 'GPP_ACOMPTE') = 'X');
  if GetParamSoc('SO_COMPTAEXTERNE') then
  begin
    TypeEcriture := copy(GetInfoParPiece(Nature, 'GPP_TYPEECRCPTA'), 1, 1);
    {    if GetInfoParPiece(Nature, 'GPP_TYPEECRCPTA') = 'NOR' then
          TypeEcriture := 'N'
        else if GetInfoParPiece(Nature, 'GPP_TYPEECRCPTA') = 'SIM' then
          TypeEcriture := 'S'; }
  end else
    TypeEcriture := 'S';
  NumLigne := 0;
  TOBLignes := TOB.Create('Lignes export compta', nil, -1);
  try
    InitTOBLigneExportCompta(TOBLignes);
    ChargeTOBLigneCommun(TOBLignes, QLigneCompta, Nature, IsRegl);
    {Les lignes de Ventes}
    if not IsRegl then
    begin
      NumLigne := NumLigne + 1;
      {Test si nouvelle pièce}
      NumEche := TOBLignes.Detail[0].GetValue('NUMECHE');
      if NumEche <= 1 then
        CptEcritureVen := CptEcritureVen + 1;
      {Test si il faut calculer le HT et TVA
       (il ne faut pas le calculer pour la 2ème ligne d'échéances)}
      LigVente := (TOBLignes.Detail[0].GetValue('NATURECOMPTABLE') = 'VEN') and (TOBLignes.Detail[0].GetValue('MODEPAIE') = '');
      {Lignes de Tiers}
      Result := ChargeLignesTiers(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne, NumEche);
      {Lignes HT}
      if Result = rcOk then
      begin
        if LigVente then
        begin
          NumLigne := NumLigne + 1;
          Result := ChargeLignesHT(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne);
          {Affectation CONTREPARTIEGEN à la ligne du tiers (=cpt vente HT, forcément la ligne précédente)}
          ContrePartieGen := TOBLignes.Detail[TOBLignes.Detail.count - 1].GetValue('GENERAL');
          TOBLignes.Detail[TOBLignes.Detail.count - 2].PutValue('CONTREPARTIEGEN', ContrePartieGen);
          {Lignes Analytique}
          if Result = rcOk then
            Result := ChargeLignesAnalytiques(TOBLignes, QLigneCompta, Nature, False, IsRegl);
        end else
        begin
          {Affectation CONTREPARTIEGEN à la ligne du tiers (=cpt vente HT)}
          TOBLignes.Detail[TOBLignes.Detail.count - 1].PutValue('CONTREPARTIEGEN', ContrePartieGen);
        end;
      end;
      if Result = rcOk then
      begin
        if LigVente then
        begin
          {Taxes}
          NumLigne := NumLigne + 1;
          Result := ChargeLignesTaxes(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne);
          {Escompte et remises}
          if Result = rcOk then
          begin
            NumLigne := NumLigne + 1;
            Result := ChargeLigneRemise(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne);
          end;
          if Result = rcOk then
          begin
            NumLigne := NumLigne + 1;
            Result := ChargeLigneEscompte(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne);
          end;
          {Ports et Frais}
          if Result = rcOk then
          begin
            NumLigne := NumLigne + 1;
            Result := ChargeLignesFraisPorts(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne);
          end;
        end;
      end;
    end else
      {Les lignes de Règlement}
    begin
      OpCaisse := (QLigneCompta.FindField('GIC_ISOPCAISSE').AsString = 'X');
      if not OpCaisse then
        CptEcritureRgt := CptEcritureRgt + 1;
      NumLigne := NumLigne + 1;
      {Lignes de Tiers}
      Result := ChargeLignesTiers(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne, 0);
      {Lignes de Mode de Règlement}
      if Result = rcOk then
      begin
        if not OpCaisse then
          NumLigne := NumLigne + 1;
        Result := ChargeLignesModePaie(TOBLignes, QLigneCompta, Nature, IsRegl, NumLigne);
      end;
      {Lignes Analytique pour les mode de Règlement}
      if Result = rcOk then
        Result := ChargeLignesAnalytiques(TOBLignes, QLigneCompta, Nature, True, IsRegl);
    end;
    {Supprimer Ligne modele}
    TOBLignes.Detail[0].Free;
    while TOBLignes.Detail.Count > 0 do
      TOBLignes.Detail[0].ChangeParent(TOBRegroupeParNature, -1);
  finally
    TOBLignes.Free;
  end;
end;

procedure InitTOBLigneExportCompta(TOBLignes: TOB);
var TOBL: TOB;
  Cpt: integer;
begin
  TOBL := TOB.Create('Premiere ligne modele', TOBLignes, -1);
  with TOBL do
  begin
    AddChampSupValeur('JOURNAL', '', False);
    AddChampSupValeur('DATECOMPTABLE', Date, False);
    AddChampSupValeur('TYPEPIECE', '', False);
    AddChampSupValeur('GENERAL', '', False);
    AddChampSupValeur('TYPECPTE', '', False);
    AddChampSupValeur('AUXILIAIRE', '', False);
    AddChampSupValeur('REFINTERNE', '', False);
    AddChampSupValeur('LIBELLE', '', False);
    AddChampSupValeur('MODEPAIE', '', False);
    AddChampSupValeur('ECHEANCE', iDate1900, False);
    AddChampSupValeur('SENS', '', False);
    AddChampSupValeur('MONTANT1', 0, False);
    AddChampSupValeur('MONTANTORIGINE', 0, False);
    AddChampSupValeur('TYPEECRITURE', '', False);
    AddChampSupValeur('NUMEROPIECE', 0, False);
    AddChampSupValeur('DEVISE', '', False);
    AddChampSupValeur('TAUXDEV', 1, False);
    AddChampSupValeur('CODEMONTANT', 'F--', False);
    AddChampSupValeur('MONTANT2', '', False);
    AddChampSupValeur('MONTANT3', '', False);
    AddChampSupValeur('ETABLISSEMENT', '', False);
    AddChampSupValeur('AXE', '', False);
    AddChampSupValeur('NUMECHE', 0, False);
    AddChampSupValeur('REFEXTERNE', '', False);
    AddChampSupValeur('DATEREFEXTERNE', iDate1900, False);
    AddChampSupValeur('DATECREATION', Date, False);
    AddChampSupValeur('SOCIETE', '', False);
    AddChampSupValeur('AFFAIRE', '', False);
    AddChampSupValeur('DATETAUXDEV', iDate1900, False);
    AddChampSupValeur('ECRANOUVEAU', '', False);
    AddChampSupValeur('QTE1', Arrondi(0, 4), False);
    AddChampSupValeur('QTE2', Arrondi(0, 4), False);
    AddChampSupValeur('QUALIFQTE1', '', False);
    AddChampSupValeur('QUALIFQTE2', '', False);
    AddChampSupValeur('REFLIBRE', '', False);
    AddChampSupValeur('TVAENCAISSEMENT', '-', False);
    AddChampSupValeur('REGIMETVA', '', False);
    AddChampSupValeur('TVA', '', False);
    AddChampSupValeur('TPF', '', False);
    AddChampSupValeur('CONTREPARTIEGEN', '', False);
    AddChampSupValeur('CONTREPARTIEAUX', '', False);
    AddChampSupValeur('REFPOINTAGE', '', False);
    AddChampSupValeur('DATEPOINTAGE', iDate1900, False);
    AddChampSupValeur('DATERELANCE', iDate1900, False);
    AddChampSupValeur('DATEVALEUR', iDate1900, False);
    AddChampSupValeur('RIB', '', False);
    AddChampSupValeur('REFRELEVE', '', False);
    AddChampSupValeur('NUMEROIMMO', '', False);
    for Cpt := 0 to 9 do
      AddChampSupValeur('LIBRETEXTE' + IntToStr(Cpt), '', False);
    for Cpt := 0 to 3 do
      AddChampSupValeur('TABLE' + IntToStr(Cpt), '', False);
    for Cpt := 0 to 3 do
      AddChampSupValeur('LIBREMONTANT' + IntToStr(Cpt), '', False);
    for Cpt := 0 to 1 do
      AddChampSupValeur('LIBREBOOL' + IntToStr(Cpt), '', False);
    AddChampSupValeur('LIBREDATE', iDate1900, False);
    AddChampSupValeur('CONSO', '', False);
    AddChampSupValeur('COUVERTURE', Arrondi(0, 4), False);
    AddChampSupValeur('COUVERTUREDEV', Arrondi(0, 4), False);
    //AddChampSupValeur('COUVERTUREEURO', Arrondi(0, 4), False);
    AddChampSupValeur('DATEPAQUETMAX', iDate1900, False);
    AddChampSupValeur('DATEPAQUETMIN', iDate1900, False);
    AddChampSupValeur('LETTRAGE', False);
    PutValue('LETTRAGE', '');
    AddChampSupValeur('LETTRAGEDEV', '', False);
    //AddChampSupValeur('LETTRAGEEURO', '-', False);
    AddChampSupValeur('ETATLETTRAGE', '', False);
    //Ajout pour génération dans ECRITURE/ANALYTIQ
    AddChampSupValeur('NATUREPIECEG', '', False);
    AddChampSupValeur('NUMPIECEG', '', False);
    AddChampSupValeur('INDICEG', 0, False);
    AddChampSupValeur('SOUCHE', '', False);
    AddChampSupValeur('COTATION', 0, False);
    AddChampSupValeur('REFGESCOM', '', False);
    AddChampSupValeur('LECPTAUXI', '', False);
    AddChampSupValeur('LECPTGENE', '', False);
    AddChampSupValeur('TYPELIGNE', '', False);
    AddChampSupValeur('MODESAISIE', '', False);
    AddChampSupValeur('NATUREAUXI', '', False);
    AddChampSupValeur('TYPETAXE', '', False);
    AddChampSupValeur('NUMEROLIGNE', 0, False);
    AddChampSupValeur('COMPTEURECRITURE', 0, False);
    AddChampSupValeur('NATURECOMPTABLE', '', False);
    AddChampSupValeur('TYPETIERS', '', False);
    AddChampSupValeur('TYPEFLUX', '', False);
    //Fin ajout
  end;
end;

function RecupValeursLibEcr(QLC: TQuery; ValeurEcr: string): string;
var Tmp, Champ, ChampInterCompta, Prefixe, Valeur, T_Auxi,Depot : string;
  CleDoc: R_CleDoc;
  TOBP, TOBT: TOB;
begin
  Result := '';
  if ValeurEcr = '' then Exit;
  Tmp := ReadTokenSt(ValeurEcr);
  while (tmp <> '') do
  begin
    Champ := RechDom('GCLIBECRITURECPTA', Tmp, True);
    if Champ <> '' then
    begin
      Valeur := '';
      Prefixe := Copy(Champ, 1, Pos('_', Champ));
      ChampInterCompta := 'GIC_' + Copy(Champ, Pos('_', Champ) + 1, (Length(Champ) + 1) - Pos('_', Champ));
      if (Prefixe = 'T_') or (Prefixe = 'ET_') or (QLC.FindField(ChampInterCompta) = nil) then
      begin
        if Prefixe = 'GP_' then
        begin
          CleDoc.NaturePiece := QLC.FindField('GIC_NATUREPIECEG').AsString;
          CleDoc.Souche := QLC.FindField('GIC_SOUCHE').AsString;
          CleDoc.NumeroPiece := QLC.FindField('GIC_NUMERO').AsInteger;
          CleDoc.Indice := QLC.FindField('GIC_INDICE').AsInteger;
          if CleDoc.NumeroPiece > 0 then
          begin
            TOBP := TOBPiece.FindFirst(['GP_NATUREPIECEG', 'GP_SOUCHE', 'GP_NUMERO', 'GP_INDICEG'],
              [CleDoc.NaturePiece, CleDoc.Souche, IntToStr(CleDoc.NumeroPiece), IntToStr(CleDoc.Indice)], False);
            if TOBP = nil then TOBP := ChargeTOBPiece(CleDoc);
            if TOBP <> nil then
              if TOBP.FieldExists(Champ) then Valeur := Trim(TOBP.GetValue(Champ));
          end;
        end
        else if Prefixe = 'T_' then
        begin
          T_Auxi := QLC.FindField('GIC_TIERS').AsString;
          TOBT := TOBTiers.FindFirst(['T_TIERS'], [T_Auxi], False);
          if TOBT = nil then TOBT := ChargeTOBTiers(T_Auxi);
          if TOBT <> nil then
            if TOBT.FieldExists(Champ) then Valeur := Trim(TOBT.GetValue(Champ));
        end
        else if Prefixe = 'ET_' then
        begin
          Depot := QLC.FindField('GIC_ETABLISSEMENT').AsString;
          Valeur :=  RechDom('TTETABLISSEMENT',Depot,FALSE) ;
        end ;
      end
      else Valeur := QLC.FindField(ChampInterCompta).AsString;
      if Result = '' then Result := Valeur else Result := Result + ' ' + Valeur;
    end;
    Tmp := ReadTokenSt(ValeurEcr);
  end;
end;

procedure AlimRefInterneEtLibelle(TOBL: TOB; QLC: TQuery; LibDefaut, Nature: string; TypeEcr: T_TypeEcr; LibEcr: Boolean);
var Fixe1, Fixe2, ValeurEcr1, ValeurEcr2, st1, st2, stTot: string;
  LibFixe1, LibFixe2, LibValeurEcr1, LibValeurEcr2, Libst1, Libst2, LibstTot: string;
begin
  if TOBL = nil then exit;
  TOBL.PutValue('LIBELLE', LibDefaut);
  TOBL.PutValue('REFINTERNE', LibDefaut);
  if not (LibEcrATraiter(TypeEcr, Nature, LibEcr)) then Exit;
  // Libelle
  LibFixe1 := Trim(GetInfoParPiece(Nature, 'GPP_RACINELIBECR1'));
  LibFixe2 := Trim(GetInfoParPiece(Nature, 'GPP_RACINELIBECR2'));
  LibValeurEcr1 := GetInfoParPiece(Nature, 'GPP_VALEURLIBECR1');
  LibValeurEcr2 := GetInfoParPiece(Nature, 'GPP_VALEURLIBECR2');
  // Ref interne
  Fixe1 := Trim(GetInfoParPiece(Nature, 'GPP_RACINEREFINT1'));
  Fixe2 := Trim(GetInfoParPiece(Nature, 'GPP_RACINEREFINT2'));
  ValeurEcr1 := GetInfoParPiece(Nature, 'GPP_VALEURREFINT1');
  ValeurEcr2 := GetInfoParPiece(Nature, 'GPP_VALEURREFINT2');

  // Ref interne
  if ((Fixe1 = '') and (Fixe2 = '') and (ValeurEcr1 = '') and (ValeurEcr2 = '')) and ((LibFixe1 = '') and (LibFixe2 = '') and (LibValeurEcr1 = '') and
    (LibValeurEcr2 = '')) then Exit;
  st1 := RecupValeursLibEcr(QLC, ValeurEcr1);
  st2 := RecupValeursLibEcr(QLC, ValeurEcr2);
  if Fixe1 <> '' then stTot := Fixe1 + ' ' + st1 else stTot := st1;
  if Fixe2 <> '' then stTot := stTot + ' ' + Fixe2 + ' ' + st2 else stTot := stTot + ' ' + st2;
  Trim(stTot);
  if stTot <> '' then TOBL.PutValue('REFINTERNE', Copy(stTot, 1, 35));

  // Libelle
  Libst1 := RecupValeursLibEcr(QLC, LibValeurEcr1);
  Libst2 := RecupValeursLibEcr(QLC, LibValeurEcr2);
  if LibFixe1 <> '' then LibstTot := LibFixe1 + ' ' + Libst1 else LibstTot := Libst1;
  if LibFixe2 <> '' then LibstTot := LibstTot + ' ' + LibFixe2 + ' ' + Libst2 else LibstTot := LibstTot + ' ' + Libst2;
  Trim(LibstTot);
  if LibstTot <> '' then TOBL.PutValue('LIBELLE', Copy(LibstTot, 1, 35));
end;

function ChargeModePaie(ModePaie: string): TOB;
var StSelect: string;
begin
  StSelect := 'SELECT MP_MODEPAIE,MP_CPTEREGLE,MP_JALREGLE,MP_GENERAL ' +
    'FROM MODEPAIE WHERE MP_MODEPAIE="' + ModePaie + '"';
  TOBModePaie.LoadDetailFromSQL(StSelect);
  Result := TOBModePaie.FindFirst(['MP_MODEPAIE'], [ModePaie], False);
end;

procedure ChargeTOBLigneCommun(TOBLignes: TOB; QLC: TQuery; NatureG: string; IsRegl: boolean);
var FinRefInterne, CodeMontant, Dev, StModePaie, StEtab: string;
  TOBModePaiement, TobDeLaPiece: TOB;
  CleDoc: R_CleDoc;
begin
  with TOBLignes.Detail[0] do
  begin
    StEtab := QLC.FindField('GIC_ETABLISSEMENT').AsString;
    if IsRegl then
    begin
      StModePaie := QLC.FindField('GIC_MODEPAIE').AsString;
      TOBModePaiement := TOBModePaieEtabl.FindFirst(['MPC_MODEPAIE', 'MPC_ETABLISSEMENT'], [StModePaie, StEtab], False);
      if TOBModePaiement = nil then
      begin
        TOBModePaiement := TOBModePaie.FindFirst(['MP_MODEPAIE'], [StModePaie], False);
        if TOBModePaiement = nil then
          TOBModePaiement := ChargeModePaie(StModePaie);
        if TOBModePaiement <> nil then
          PutValue('JOURNAL', VarToStr(TOBModePaiement.GetValue('MP_JALREGLE')));
      end else
        if TOBModePaiement <> nil then
        PutValue('JOURNAL', VarToStr(TOBModePaiement.GetValue('MPC_JALREGLE')));
    end else
    begin
      PutValue('JOURNAL', GetInfoParPiece(NatureG, 'GPP_JOURNALCPTA'));
      //JT le 17/06/2003
      PutValue('MODEPAIE', QLC.FindField('GIC_MODEPAIE').AsString);
    end;
    PutValue('DATECOMPTABLE', QLC.FindField('GIC_DATE').AsDateTime);
    PutValue('TYPEPIECE', GetInfoParPiece(NatureG, 'GPP_NATURECPTA'));
    PutValue('TYPEFLUX', GetInfoParPiece(NatureG, 'GPP_VENTEACHAT'));
    if QLC.FindField('GIC_NUMERO').AsInteger > 0 then
      FinRefInterne := ' N°' + QLC.FindField('GIC_NUMERO').AsString
    else
      FinRefInterne := '';
    AlimRefInterneEtLibelle(TOBLignes.Detail[0], QLC, Copy(GetInfoParPiece(NatureG, 'GPP_LIBELLE') + FinRefInterne, 1, 35), NatureG, tecTiers, False);
    PutValue('ECHEANCE', QLC.FindField('GIC_DATEECHE').AsDateTime);
    PutValue('TYPEECRITURE', TypeEcriture);
    PutValue('NUMEROPIECE', QLC.FindField('GIC_NUMERO').AsInteger);
    Dev := QLC.FindField('GIC_DEVISE').AsString;
    PutValue('DEVISE', Dev);
    PutValue('TAUXDEV', QLC.FindField('GIC_TAUXDEVISE').AsFloat);
    if Dev = 'EUR' then
      CodeMontant := 'E--'
    else
      CodeMontant := 'ED-';
    //if V_PGI.DeviseFongible=Dev then CodeMontant:=CodeMontant+'F' else CodeMontant:=CodeMontant+'-';
    PutValue('CODEMONTANT', CodeMontant);
    PutValue('ETABLISSEMENT', StEtab);
    PutValue('DATECREATION', Date);
    PutValue('SOCIETE', V_PGI.CodeSociete); //V_PGI.NomSociete
    PutValue('DATETAUXDEV', Date);
    PutValue('ECRANOUVEAU', 'N');
    if QLC.FindField('GIC_TVAENCAISSEMENT').AsString = 'TE' then
      PutValue('TVAENCAISSEMENT', 'X');
    PutValue('REGIMETVA', QLC.FindField('GIC_REGIMETAXE').AsString);
    //Ajout pour génération dans ECRITURE/ANALYTIQ
    PutValue('NATUREPIECEG', QLC.FindField('GIC_NATUREPIECEG').AsString);
    PutValue('NUMPIECEG', QLC.FindField('GIC_NUMERO').AsString);
    PutValue('NUMECHE', QLC.FindField('GIC_NUMECHE').AsInteger);
    PutValue('INDICEG', QLC.FindField('GIC_INDICE').AsInteger);
    PutValue('SOUCHE', QLC.FindField('GIC_SOUCHE').AsString);
    PutValue('NATURECOMPTABLE', QLC.FindField('GIC_TYPEJOURNAL').AsString);
    if GetValue('NUMPIECEG') > 0 then
    begin
      CleDoc.NaturePiece := GetValue('NATUREPIECEG');
      CleDoc.Souche := GetValue('SOUCHE');
      CleDoc.NumeroPiece := GetValue('NUMPIECEG');
      CleDoc.Indice := GetValue('INDICEG');
      TobDeLaPiece := ChargeTOBPiece(CleDoc);
      PutValue('REFGESCOM', EncodeRefCPGescom(TobDeLaPiece));
    end else
    begin
      PutValue('REFGESCOM', '');
    end;
    //Fin ajout
  end;
end;

procedure ChargeTiersEnTOB(Q: TQuery);
var TOBT: TOB;
begin
  if TOBTiers.FindFirst(['T_TIERS'], [Q.FindField('T_TIERS').AsString], False) = nil then
  begin
    TOBT := TOB.Create('TIERS', TOBTiers, -1);
    TOBT.PutValue('T_TIERS', Q.FindField('T_TIERS').AsString);
    TOBT.PutValue('T_AUXILIAIRE', Q.FindField('T_AUXILIAIRE').AsString);
    TOBT.PutValue('T_COLLECTIF', Q.FindField('T_COLLECTIF').AsString);
    TOBT.PutValue('T_CONSO', Q.FindField('T_CONSO').AsString);
    TOBT.PutValue('T_LIBELLE', Q.FindField('T_LIBELLE').AsString);
    TOBT.PutValue('T_NATUREAUXI', Q.FindField('T_NATUREAUXI').AsString);
  end;
end;

function ChargeTOBTiers(T_Tiers: string): TOB;
var StSelect: string;
begin
  StSelect := 'SELECT T_AUXILIAIRE,T_TIERS,T_COLLECTIF,T_CONSO,T_LIBELLE,T_NATUREAUXI,T_PARTICULIER,' +
    'R_ETABBQ, R_GUICHET, R_NUMEROCOMPTE, R_CLERIB, R_DOMICILIATION ' +
    'FROM TIERS ' +
    'LEFT JOIN RIB ON T_AUXILIAIRE = R_AUXILIAIRE AND R_PRINCIPAL="X" ' +
    'WHERE T_TIERS="' + T_Tiers + '"';
  TOBTiers.LoadDetailFromSQL(StSelect);
  Result := TOBTiers.FindFirst(['T_TIERS'], [T_Tiers], False);
end;

function ChargeTOBPiece(CleDoc: R_CleDoc): TOB;
var StSelect: string;
begin
  StSelect := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_DATEPIECE,' +
    'GP_REFINTERNE,GP_REFEXTERNE,GP_AFFAIRE,GP_ETABLISSEMENT,GP_TIERS ' +
    'FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False);
  TOBPiece.LoadDetailFromSQL(StSelect);
  Result := TOBPiece.FindFirst(['GP_NATUREPIECEG', 'GP_SOUCHE', 'GP_NUMERO', 'GP_INDICEG'],
    [CleDoc.NaturePiece, CleDoc.Souche, IntToStr(CleDoc.NumeroPiece), IntToStr(CleDoc.Indice)], False);
end;

function ChargeLignesTiers(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne, NumEche: integer): T_RetCpta;
var TOBLigneTiers, TOBT, TOBModePaiement, TOBO: TOB;
  T_Tiers, DebitCredit, Rib: string;
  MontantTTC: Double;
  BVente: boolean;
  QCaisse: TQuery;
begin
  Result := rcOk;
  //  if not IsRegl then
  //    MontantTTC := QLC.FindField('GIC_TOTALTTC').AsFloat
  //  else
  MontantTTC := QLC.FindField('GIC_MONTANTTTC').AsFloat;
  if (MontantTTC = 0) then exit;

  TOBLigneTiers := TOB.Create('TOB Ligne Tiers', TOBLignes, -1);
  TOBLigneTiers.Dupliquer(TOBLignes.Detail[0], False, True, True);

  if GetInfoParPiece(Nature, 'GPP_CPTCENTRAL') = 'X' then
  begin
    QCaisse := OpenSQL('Select GPK_TIERS from PARCAISSE where GPK_ETABLISSEMENT="' + QLC.FindField('GIC_ETABLISSEMENT').AsString + '"', True);
    if not QCaisse.EOF then
      T_Tiers := QCaisse.Fields[0].AsString;
    ferme(QCaisse);
  end;
  if T_Tiers = '' then
    T_Tiers := QLC.FindField('GIC_TIERS').AsString;

  TOBT := TOBTiers.FindFirst(['T_TIERS'], [T_Tiers], False);
  if TOBT = nil then
  begin
    TOBT := ChargeTOBTiers(T_Tiers);
    if TOBT = nil then
    begin
      Result := rcPar;
      NumErreur := 3;
      Exit;
    end;
  end;

  LaNatureTiers := TOBT.GetValue('T_NATUREAUXI');
  LeTypeTiers := TOBT.GetValue('T_PARTICULIER');
  //Test si regroupement sur auxi de centralisation
  if not AuxiCentralisation then
  begin
    LeCptAuxiliaire := TOBT.GetValue('T_AUXILIAIRE');
  end else
    if GetInfoParPiece(Nature, 'GPP_CPTCENTRAL') = 'X' then
  begin
    LeCptAuxiliaire := TOBT.GetValue('T_AUXILIAIRE');
  end else
  begin
    //Particulier
    if LeTypeTiers = 'X' then
    begin
      if not DetParticulier then
      begin
        if TobLigneTiers.GetValue('TYPEFLUX') = 'VEN' then
          LeCptAuxiliaire := GetParamSoc('SO_GCCLICPTADIFFPART')
        else
          LeCptAuxiliaire := GetParamSoc('SO_GCFOUCPTADIFFPART');
      end else
      begin
        LeCptAuxiliaire := TOBT.GetValue('T_AUXILIAIRE');
      end;
    end else
      //Entreprise
    begin
      if not DetEntreprise then
      begin
        if TobLigneTiers.GetValue('TYPEFLUX') = 'VEN' then
          LeCptAuxiliaire := GetParamSoc('SO_GCCLICPTADIFF')
        else
          LeCptAuxiliaire := GetParamSoc('SO_GCFOUCPTADIFF');
      end else
      begin
        LeCptAuxiliaire := TOBT.GetValue('T_AUXILIAIRE');
      end;
    end;
    if LeCptAuxiliaire = '' then
      LeCptAuxiliaire := TOBT.GetValue('T_AUXILIAIRE');
  end;

  with TOBLigneTiers do
  begin
    // Pour les opérations de caisse la contre partie est équivalent au compte de caisse de l'etab ou de banque
    if (IsRegl) and (QLC.FindField('GIC_COMPTAARTICLE').AsString <> '') and (not GestionAcompte) then
    begin
      // Si ligne de fond de caisse ne pas traiter la ligne et gérer contrepartie avec caisse
      if IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'FCA') then
      begin
        // Tenir compte de l'ecart de caisse que si il est négatif
        if (MontantTTC > 0) and (GestionOpCaisse) then TOBLigneTiers.free;
        ChargeTOBEcartCaisse(QLC.FindField('GIC_ETABLISSEMENT').AsString);
        if (MontantTTC > 0) and (GestionOpCaisse) then Exit;
      end;
      // Si ecart de caisse: Pour les opérations de caisse la contre partie est équivalent au compte de caisse de l'etab
      if TOBEcartCaisse.FindFirst(['_ETAB'], [QLC.FindField('GIC_ETABLISSEMENT').AsString], False) <> nil then
      begin
        TOBO := TOBOperationCaisseEtab.FindFirst(['GFC_ETABLISSEMENT', 'GFC_ARTICLE'], [QLC.FindField('GIC_ETABLISSEMENT').AsString, ArticleFondCaisse],
          False);
        //2- Cherche compte dans Opération caisse (ARTICLE FI)
        if TOBO = nil then
        begin
          TOBO := TOBOperationCaisse.FindFirst(['GA_ARTICLE'], [ArticleFondCaisse], False);
          if TOBO <> nil then PutValue('GENERAL', TOBO.GetValue('GA_DESIGNATION1'));
        end
        else PutValue('GENERAL', TOBO.GetValue('GFC_COMPTE'));
        if TOBO <> nil then
        begin
          PutValue('MONTANT1', Abs(MontantTTC));
          // MODIF AC 30/10/03
          // Pour les ecarts de caisse inversion et fond de caisse si montant négatif
          if (MontantTTC < 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString,'FCA')) then
            DebitCredit := 'D'
              // Pour les versements d'arrhe inversion si montant positif
          else if (MontantTTC > 0) and IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'VAR') then
            DebitCredit := 'D'
              // Pour les fonds de caisse positif
          else if (MontantTTC > 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'FCA')) then
            DebitCredit := 'C'
          // MODIF AC 30/10/03
          // Pour les ecarts de caisse le compte de banque est au D si Mt>0 et au C si Mt<0
          else if (MontantTTC < 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'ECA')) then
            DebitCredit := 'C'
          else if (MontantTTC > 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'ECA')) then
            DebitCredit := 'D'
          // FIN MODIF AC 30/10/03
          else
            DebitCredit := 'C';
        end;
      end else
        // Sinon contre partie sur les comptes de banques (ModePaieCompl, ModePaie)
      begin
        TOBModePaiement := TOBModePaieEtabl.FindFirst(['MPC_MODEPAIE', 'MPC_ETABLISSEMENT'], [QLC.FindField('GIC_MODEPAIE').AsString,
          QLC.FindField('GIC_ETABLISSEMENT').AsString], False);
        if TOBModePaiement = nil then
        begin
          TOBModePaiement := TOBModePaie.FindFirst(['MP_MODEPAIE'], [QLC.FindField('GIC_MODEPAIE').AsString], False);
          if TOBModePaiement = nil then
            TOBModePaiement := ChargeModePaie(QLC.FindField('GIC_MODEPAIE').AsString);
          if TOBModePaiement <> nil then
          begin
            LeCptGeneral := VarToStr(TOBModePaiement.GetValue('MP_CPTEREGLE'));
            PutValue('GENERAL', VarToStr(TOBModePaiement.GetValue('MP_CPTEREGLE')));
            PutValue('JOURNAL', TOBModePaiement.GetValue('MP_JALREGLE'));
          end;
        end else
          if TOBModePaiement <> nil then
        begin
          PutValue('GENERAL', VarToStr(TOBModePaiement.GetValue('MPC_CPTEREGLE')));
          PutValue('JOURNAL', TOBModePaiement.GetValue('MPC_JALREGLE'));
        end;
        PutValue('MONTANT1', Abs(MontantTTC));
        // Pour les ecarts de caisse inversion si montant négatif
        if (MontantTTC < 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'ECA')) then
          DebitCredit := 'D'
            // Pour les versements d'arrhe inversion si montant positif
        else if (MontantTTC > 0) and IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'VAR') then
          DebitCredit := 'D'
        // MODIF AC 30/10/03
        // Pour les ecarts de caisse le compte de banque est au D si Mt>0 et au C si Mt<0
        else if (MontantTTC < 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'ECA')) then
          DebitCredit := 'C'
        else if (MontantTTC > 0) and (IsArtFinancier(QLC.FindField('GIC_ARTICLE').AsString, 'ECA')) then
          DebitCredit := 'D'
        // FIN MODIF AC 30/10/03
        else
          DebitCredit := 'C';
      end;
    end else
    begin
      PutValue('GENERAL', TOBT.GetValue('T_COLLECTIF'));
      PutValue('TYPECPTE', 'X');
      //            PutValue('AUXILIAIRE',TOBT.GetValue('T_AUXILIAIRE'));
      PutValue('AUXILIAIRE', LeCptAuxiliaire);
      PutValue('CONSO', NullToVide(TOBT.GetValue('T_CONSO')));
      PutValue('NUMECHE', NumEche);
      PutValue('MONTANT1', MontantTTC);
      PutValue('MONTANT2', QLC.FindField('GIC_MONTANTTTCDEV').AsFloat);
    end;

    BVente := GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN';
    if IsRegl then
    begin
      // MODIF AC 27/10/2003
      //PutValue('LIBELLE', TOBT.GetValue('T_LIBELLE'));
      PutValue('MODEPAIE', QLC.FindField('GIC_MODEPAIE').AsString);
      // Type de piece=OD si ticket
      if IsNatCaisse(Nature) then
        PutValue('TYPEPIECE', 'OD');
      if DebitCredit = '' then
        if BVente then
        begin
          //          DebitCredit := 'C'
          if TOBLignes.Detail[0].GetValue('TYPEPIECE') <> 'AC' then
            DebitCredit := 'C'
          else
            DebitCredit := 'D';
        end else
        begin
          DebitCredit := 'D';
        end;
      PutValue('SENS', DebitCredit);
    end else
    begin
      //JT le 17/06/2003
      if QLC.FindField('GIC_MODEPAIE').AsString = '' then
        PutValue('MODEPAIE', ModePaiementDefaut)
      else
        PutValue('MODEPAIE', QLC.FindField('GIC_MODEPAIE').AsString);
      if BVente then
        PutValue('SENS', 'D')
      else
        PutValue('SENS', 'C');
    end;
    PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
    // MODIF AC 27/10/03
    // Si montant négatif inverse SENS sauf pour les OPCaisse et mettre le montant1 en valeur absolue
    if (QLC.FindField('GIC_COMPTAARTICLE').AsString = '') and (GetValue('MONTANT1')<0) then
    begin
      PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
      PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
      if GetValue('SENS')='D' then PutValue('SENS','C')
      else PutValue('SENS','D') ;
    end ;
    LeCptGeneral := GetValue('GENERAL');
    TOBT.PutValue('R_ETABBQ', NullToVide(TOBT.GetValue('R_ETABBQ')));
    TOBT.PutValue('R_GUICHET', NullToVide(TOBT.GetValue('R_GUICHET')));
    TOBT.PutValue('R_NUMEROCOMPTE', NullToVide(TOBT.GetValue('R_NUMEROCOMPTE')));
    TOBT.PutValue('R_CLERIB', NullToVide(TOBT.GetValue('R_CLERIB')));
    TOBT.PutValue('R_DOMICILIATION', NullToVide(TOBT.GetValue('R_DOMICILIATION')));
    if (TOBLigneTiers.GetValue('NUMEROPIECE') > 0) and (TOBT.GetValue('R_ETABBQ') <> '') then
    begin
      Rib := EncodeRIB(TOBT.GetValue('R_ETABBQ'), TOBT.GetValue('R_GUICHET'),
        TOBT.GetValue('R_NUMEROCOMPTE'), TOBT.GetValue('R_CLERIB'),
        TOBT.GetValue('R_DOMICILIATION'))
    end else
    begin
      Rib := '';
    end;
    InfosCompl(IsRegl, TOBLigneTiers, LeCptAuxiliaire, LeCptGeneral, 'TIERS', LaNatureTiers, NumLigne, Rib, LeTypeTiers);
  end;
end;

function ChargeCompteVente(QLC: TQuery): TOB;
var FamArt, FamTiers, FamAff, Etab, Regime,
  Nature, VenteAchat, NatV, SQL, LastSQL: string;
  TOBC: TOB;
  Q: TQuery;
begin
  FamArt := QLC.FindField('GIC_COMPTAARTICLE').AsString;
  FamTiers := QLC.FindField('GIC_COMPTATIERS').AsString;
  FamAff := '';
  Etab := QLC.FindField('GIC_ETABLISSEMENT').AsString;
  Regime := QLC.FindField('GIC_REGIMETAXE').AsString;

  Nature := QLC.FindField('GIC_NATUREPIECEG').AsString;
  VenteAchat := GetInfoParPiece(Nature, 'GPP_VENTEACHAT');
  if VenteAchat = 'VEN' then NatV := 'HV' else NatV := 'HA';

  TOBC := FindTOBCode(TOBCpta, FamArt, FamTiers, FamAff, Etab, Regime);
  LastSQL := '';
  SQL := '';
  if TOBC = nil then
  begin
    SQL := FabricSQLCompta(FamArt, FamTiers, FamAff, Etab, Regime);
    if TOBCpta.FieldExists('LASTSQL') then
    begin
      LastSQL := TOBCpta.GetValue('LASTSQL');
      TOBCpta.PutValue('LASTSQL', SQL);
    end;
  end;
  if ((TOBC = nil) and (SQL <> LastSQL)) then
  begin
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      TOBC := CreerTOBCodeCpta(TOBCpta);
      TOBC.SelectDB('', Q);
      TOBCpta.Detail.Sort('GCP_REGIMETAXE;GCP_ETABLISSEMENT');
    end;
    Ferme(Q);
  end;
  Result := TOBC;
end;

function ChargeLignesHT(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
var TOBLigneHT, TOBCode: TOB;
  NatureCpta, CptHT: string;
  MontantHT: Double;
begin
  Result := rcOk;
  MontantHT := QLC.FindField('GIC_MONTANTHT').AsFloat;
  if MontantHT = 0 then exit;

  TOBLigneHT := TOB.Create('TOB Ligne HT', TOBLignes, -1);
  TOBLigneHT.Dupliquer(TOBLignes.Detail[0], False, True, True);

  NatureCpta := TOBLignes.Detail[0].GetValue('TYPEPIECE');
  TOBCode := ChargeCompteVente(QLC);
  // Charge les axes analytiques de charge en fonction du ParamSoc
  if (TOBCode <> nil) and (not GetParamSoc('SO_GCAXEANALYTIQUE')) then
    ChargeAxeAnalytique(IntToStr(TOBCode.GetValue('GCP_RANG')), Nature, QLC);
  //
  if TOBCode <> nil then
  begin
    if ((NatureCpta = 'FC') or (NatureCpta = 'AC')) then
      CptHT := TOBCode.GetValue('GCP_CPTEGENEVTE')
    else
      CptHT := TOBCode.GetValue('GCP_CPTEGENEACH');
  end;
  if (TOBCode = nil) or (CptHT = '') then
  begin
    if ((NatureCpta = 'FC') or (NatureCpta = 'AC')) then
      CptHT := VH_GC.GCCpteHTVTE
    else
      CptHT := VH_GC.GCCpteHTACH;
  end;
  // Erreur sur le compte HT article
  if CptHT = '' then
  begin
    Result := rcPar;
    NumErreur := 4;
    exit;
  end;
  with TOBLigneHT do
  begin
    PutValue('GENERAL', CptHT);
    PutValue('MONTANT1', MontantHT);
    PutValue('MONTANT2', QLC.FindField('GIC_MONTANTHTDEV').AsFloat);
    PutValue('QTE1', Arrondi(QLC.FindField('GIC_QTE1').AsFloat, 4));
    PutValue('TVA', QLC.FindField('GIC_FAMILLETAXE1').AsString);
    PutValue('TPF', QLC.FindField('GIC_FAMILLETAXE2').AsString);
    if GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN' then
      PutValue('SENS', 'C')
    else
      PutValue('SENS', 'D');
    PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
    InfosCompl(IsRegl, TOBLigneHT, LeCptAuxiliaire, LeCptGeneral, 'HT', LaNatureTiers, NumLigne, '', LeTypeTiers);
    // MODIF AC 27/10/03
    // Si montant négatif inverse SENS et on met Montant>0
    if GetValue('MONTANT1')<0 then
    begin
      PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
      PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
      if GetValue('SENS')='D' then PutValue('SENS','C')
      else PutValue('SENS','D') ;
    end ;
  end;
  //AlimRefInterneEtLibelle(TOBLigneHT,QLC,'HT',tecHT,True);
end;

function EquilibrageDesComptes(QLC: TQuery): Double;
var Taxes, HT, TTC, Remise, Escompte, PortEtFrais: Double;
  Cpt: integer;
begin
  //Normalement le résultat doit être égal à zéro si les montants sont exacts.
  Result := 0;
  Taxes := 0;
  HT := 0;
  TTC := 0;
  Remise := 0;
  Escompte := 0;
  PortEtFrais := 0;
  for Cpt := 1 to 5 do
    Taxes := Taxes + QLC.FindField('GIC_MONTANTTAXE' + IntToStr(Cpt)).AsFloat;
  HT := QLC.FindField('GIC_MONTANTHT').AsFloat;
  //  TTC := QLC.FindField('GIC_TOTALTTC').AsFloat;
  TTC := HT + Taxes;
  Remise := QLC.FindField('GIC_MONTANTREM').AsFloat;
  Escompte := QLC.FindField('GIC_MONTANTESC').AsFloat;
  PortEtFrais := QLC.FindField('GIC_MONTANTPF').AsFloat;
  Result := TTC + Remise + Escompte - HT - Taxes - PortEtFrais;
  if (Result < 0.0000001) and (Result > -0.0000001) then
    Result := 0.00;
end;

function TVAETAB2(CatTaxe, ModeTVA, Tva, Etablissement: String3; Achat, TvaEnc: Boolean): string;
var TOBT: TOB;
begin
  Result := '';
  //TOBT:=TOBTVAEtabl.FindFirst(['TVC_TVAOUTPF','TVC_REGIME','TVC_CODETAUX','TVC_ETABLISSEMENT'],[VH^.DefCatTVA,ModeTVA,TVA,Etablissement],False) ;
  TOBT := TOBTVAEtabl.FindFirst(['TVC_TVAOUTPF', 'TVC_REGIME', 'TVC_CODETAUX', 'TVC_ETABLISSEMENT'],
    [CatTaxe, ModeTVA, TVA, Etablissement], False);
  if TOBT <> nil then
  begin
    if Achat then
      Result := TOBT.GetValue('TVC_CPTEACH')
    else
      Result := TOBT.GetValue('TVC_CPTEVTE');
  end else
  begin
    if TvaEnc then
      Result := TVA2ENCAIS(ModeTVA, Tva, Achat);
    if Result = '' then
      Result := TVA2CPTE(ModeTVA, Tva, Achat);
  end;
end;

function ChargeLignesTaxes(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
var TOBLigneTaxe: TOB;
  i: integer;
  Regime, FamTaxe, NatureCpta, CptTaxe, Etab, TypeTaxe: string;
  Achat, TvaEnc: boolean;
  MontantTaxe: Double;
begin
  Result := rcOk;
  Etab := QLC.FindField('GIC_ETABLISSEMENT').AsString;
  Regime := QLC.FindField('GIC_REGIMETAXE').AsString;
  NatureCpta := TOBLignes.Detail[0].GetValue('TYPEPIECE');
  Achat := ((NatureCpta = 'AF') or (NatureCpta = 'FF'));
  TvaEnc := (QLC.FindField('GIC_TVAENCAISSEMENT').AsString = 'TE');
  for i := 1 to 5 do
  begin
    MontantTaxe := QLC.FindField('GIC_MONTANTTAXE' + IntToStr(i)).AsFloat;
    //    MontantTaxe := Arrondi(QLC.FindField('GIC_MONTANTTAXE' + IntToStr(i)).AsFloat + QLC.FindField('GIC_MONTANTTAXEDEV' + IntToStr(i)).AsFloat,V_PGI.OkDecV);
    FamTaxe := QLC.FindField('GIC_FAMILLETAXE' + IntToStr(i)).AsString;
    if (MontantTaxe <> 0) and (FamTaxe <> '') then
    begin
      //On met l'eccart des comptes sur le compte de taxe de la TVA
//      MontantTaxe := MontantTaxe + EquilibrageDesComptes(QLC);
      TOBLigneTaxe := TOB.Create('TOB Ligne HT', TOBLignes, -1);
      TOBLigneTaxe.Dupliquer(TOBLignes.Detail[0], False, True, True);
      CptTaxe := '';
      case i of
        1:
          begin
            CptTaxe := TVAETAB2(VH^.DefCatTVA, Regime, FamTaxe, Etab, Achat, TvaEnc);
            TypeTaxe := 'TVA';
          end;
        2:
          begin
            CptTaxe := TVAETAB2(VH^.DefCatTPF, Regime, FamTaxe, Etab, Achat, TvaEnc);
            TypeTaxe := 'TPF';
          end;
      end;
      // Erreur sur Compte Taxe
      if CptTaxe = '' then
      begin
        Result := rcPar;
        NumErreur := 5;
        Break;
      end;
      with TOBLigneTaxe do
      begin
        PutValue('GENERAL', CptTaxe);
        PutValue('TYPETAXE', TypeTaxe);
        PutValue('MONTANT1', MontantTaxe);
        PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
        PutValue('MONTANT2', QLC.FindField('GIC_MONTANTTAXEDEV' + IntToStr(i)).AsFloat);
        if GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN' then PutValue('SENS', 'C') else PutValue('SENS', 'D');
        InfosCompl(IsRegl, TOBLigneTaxe, LeCptAuxiliaire, LeCptGeneral, 'TAXE', LaNatureTiers, NumLigne, '', LeTypeTiers);
        // MODIF AC 27/10/03
        // Si montant négatif inverse SENS et MONTANT1=ABS(MONTANT1)
        if GetValue('MONTANT1')<0 then
        begin
          PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
          PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
          if GetValue('SENS')='D' then PutValue('SENS','C')
          else PutValue('SENS','D') ;
        end ;
      end;
      //AlimRefInterneEtLibelle(TOBLigneTaxe,QLC,'Taxe',tecTaxe,True);
    end;
  end;
end;

function ChargeLigneRemise(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
var TOBLigneRemise: TOB;
  NatureCpta, CpteRem: string;
  MontantRemise: Double;
begin
  Result := rcOk;
  MontantRemise := QLC.FindField('GIC_MONTANTREM').AsFloat;
  if MontantRemise = 0 then exit;

  TOBLigneRemise := TOB.Create('TOB Ligne HT', TOBLignes, -1);
  TOBLigneRemise.Dupliquer(TOBLignes.Detail[0], False, True, True);

  NatureCpta := TOBLignes.Detail[0].GetValue('TYPEPIECE');
  if ((NatureCpta = 'FC') or (NatureCpta = 'AC')) then
    CpteRem := VH_GC.GCCpteRemVTE else CpteRem := VH_GC.GCCpteRemACH;
  // Erreur sur Compte Remise
  if CpteRem = '' then
  begin
    Result := rcPar;
    NumErreur := 6;
    Exit;
  end;
  with TOBLigneRemise do
  begin
    PutValue('GENERAL', CpteRem);
    PutValue('MONTANT1', MontantRemise);
    PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
    PutValue('MONTANT2', QLC.FindField('GIC_MONTANTREMDEV').AsFloat);
    if GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN' then PutValue('SENS', 'D') else PutValue('SENS', 'C');
    InfosCompl(IsRegl, TOBLigneRemise, LeCptAuxiliaire, LeCptGeneral, 'REMISE', LaNatureTiers, NumLigne, '', LeTypeTiers);
    // MODIF AC 27/10/03
    // Si montant négatif inverse SENS + Abs montant1
    if GetValue('MONTANT1')<0 then
    begin
      PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
      PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
      if GetValue('SENS')='D' then PutValue('SENS','C')
      else PutValue('SENS','D') ;
    end ;
  end;
  //AlimRefInterneEtLibelle(TOBLigneRemise,QLC,'Remise pied',tecRemise,True);
end;

function ChargeLigneEscompte(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
var TOBLigneEscompte: TOB;
  NatureCpta, CpteEsc: string;
  MontantEscompte: Double;
begin
  Result := rcOk;
  MontantEscompte := QLC.FindField('GIC_MONTANTESC').AsFloat;
  if MontantEscompte = 0 then exit;

  TOBLigneEscompte := TOB.Create('TOB Ligne HT', TOBLignes, -1);
  TOBLigneEscompte.Dupliquer(TOBLignes.Detail[0], False, True, True);

  NatureCpta := TOBLignes.Detail[0].GetValue('TYPEPIECE');
  if ((NatureCpta = 'FC') or (NatureCpta = 'AC')) then
    CpteEsc := VH_GC.GCCpteEscVTE else CpteEsc := VH_GC.GCCpteEscACH;
  // Erreur sur Compte Escompte
  if CpteEsc = '' then
  begin
    Result := rcPar;
    NumErreur := 7;
    Exit;
  end;
  with TOBLigneEscompte do
  begin
    PutValue('GENERAL', CpteEsc);
    PutValue('MONTANT1', MontantEscompte);
    PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
    PutValue('MONTANT2', QLC.FindField('GIC_MONTANTESCDEV').AsFloat);
    if GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN' then PutValue('SENS', 'D') else PutValue('SENS', 'C');
    InfosCompl(IsRegl, TOBLigneEscompte, LeCptAuxiliaire, LeCptGeneral, 'ESCOMPTE', LaNatureTiers, NumLigne, '', LeTypeTiers);
     // MODIF AC 27/10/03
    // Si montant négatif inverse SENS
    if GetValue('MONTANT1')<0 then
    begin
      PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
      PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
      if GetValue('SENS')='D' then PutValue('SENS','C')
      else PutValue('SENS','D') ;
    end ;
  end;
  //AlimRefInterneEtLibelle(TOBLigneEscompte,QLC,'Escompte',tecRemise,True);
end;

function ChargeLignesModePaie(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
var TOBLigneModePaie, TOBModePaiement, TOBCode, TOBOP: TOB;
  StModePaie, CpteGeneral, StEtab, StComptaArticle, NatureCpta, DebitCredit, StArticle, Journal: string;
  MontantReglement: Double;
begin
  Result := rcOk;
  DebitCredit := '';
  MontantReglement := QLC.FindField('GIC_MONTANTTTC').AsFloat;
  if MontantReglement = 0 then exit;

  TOBLigneModePaie := TOB.Create('TOB Ligne Mode Paiement', TOBLignes, -1);
  TOBLigneModePaie.Dupliquer(TOBLignes.Detail[0], False, True, True);

  StModePaie := QLC.FindField('GIC_MODEPAIE').AsString;
  StEtab := QLC.FindField('GIC_ETABLISSEMENT').AsString;
  StComptaArticle := QLC.FindField('GIC_COMPTAARTICLE').AsString;
  StArticle := QLC.FindField('GIC_ARTICLE').AsString;
  // Pas de traitement pour les fond de caisse modifiée et est positif
  if (StArticle <> '') and (IsArtFinancier(StArticle, 'FCA')) and (MontantReglement > 0) and (GestionOpCaisse) then
  begin
    TOBLigneModePaie.Free;
    Exit;
  end;
  //Traitement différent si opération de caisse (GIC_COMPTAARTICLE<>"")
  //2- Cherche compte dans la ventilation comptable, s'il ne trouve pas dans MODEPAIECOMPL puis MODEPAIE
  if (StComptaArticle <> '') and (QLC.FindField('GIC_ISOPCAISSE').AsString = 'X') then
  begin
    //1- Cherche compte dans ARTFINANCIERCOMPL
    TOBOP := TOBOperationCaisseEtab.FindFirst(['GFC_ETABLISSEMENT', 'GFC_ARTICLE'], [StEtab, StArticle], False);
    //2- Cherche compte dans ARTICLE FI
    if TOBOP = nil then
    begin
      //      TOBOP := TOBOperationCaisse.FindFirst(['GA_ARTICLE'], [ArticleFondCaisse],False) ;
      TOBOP := TOBOperationCaisse.FindFirst(['GA_ARTICLE'], [StArticle], False);
      if TOBOP <> nil then
      begin
        CpteGeneral := TOBOP.GetValue('GA_DESIGNATION1');
        Journal := TOBOP.GetValue('GA_DESIGNATION2');
      end;
    end else
    begin
      CpteGeneral := TOBOP.GetValue('GFC_COMPTE');
      Journal := TOBOP.GetValue('GFC_JOURNAL');
    end;
    if TOBOP <> nil then
    begin
      // Opération de caisse positif = Crédit
      if MontantReglement > 0 then
      begin
        // MODFI AC 30/10/03
        //if (IsArtFinancier(StArticle, 'ECA')) or IsArtFinancier(StArticle, 'FCA') then DebitCredit := 'D' else DebitCredit := 'C';
        if IsArtFinancier(StArticle, 'FCA') then DebitCredit := 'D' else DebitCredit := 'C';
      end else
      begin
        // Si opération caisse négative Montant devient positif et Sens Debit
        MontantReglement := abs(MontantReglement);
        // MODFI AC 30/10/03
        // Inversé pour les fonds de caisse //et les ecarts de caisse//
        //if IsArtFinancier(StArticle, 'ECA') or IsArtFinancier(StArticle, 'FCA') then DebitCredit := 'C' else DebitCredit := 'D';
        if IsArtFinancier(StArticle, 'FCA') then DebitCredit := 'C' else DebitCredit := 'D';
      end;
    end;
    //2- Cherche compte dans la ventilation comptable, s'il ne trouve pas dans OPERATIONCAISSECOMPL puis OPERATIONCAISSE
    if CpteGeneral = '' then
    begin
      NatureCpta := TOBLignes.Detail[0].GetValue('TYPEPIECE');
      TOBCode := ChargeCompteVente(QLC);
      if TOBCode <> nil then
      begin //Si montant<0 pour les opérations de caisse récupérer le compte d'achat
        if (((NatureCpta = 'FC') or (NatureCpta = 'AC')) and (MontantReglement > 0)) then
        begin
          CpteGeneral := TOBCode.GetValue('GCP_CPTEGENEVTE');
          // Opération de caisse positif = Crédit
          DebitCredit := 'C';
        end else
        begin
          CpteGeneral := TOBCode.GetValue('GCP_CPTEGENEACH');
          // Si opération caisse négative Montant devient positif et Sens Debit
          MontantReglement := abs(MontantReglement);
          DebitCredit := 'D';
        end;
      end;
    end;
    // Pour les fonds de caisse récupérer le compte et journal du mode de paie
    if IsArtFinancier(StArticle, 'FCA') then
    begin
      CpteGeneral := '';
      Journal := '';
    end;
  end;
  if CpteGeneral = '' then
  begin
    //3- Cherche le compte dans la Table MODEPAIECOMPL et s'il ne trouve pas cherche dans MODEPAIE
    TOBModePaiement := TOBModePaieEtabl.FindFirst(['MPC_MODEPAIE', 'MPC_ETABLISSEMENT'], [StModePaie, StEtab], False);
    if TOBModePaiement = nil then
    begin
      TOBModePaiement := TOBModePaie.FindFirst(['MP_MODEPAIE'], [StModePaie], False);
      if TOBModePaiement = nil then TOBModePaiement := ChargeModePaie(StModePaie);
      if TOBModePaiement <> nil then CpteGeneral := VarToStr(TOBModePaiement.GetValue('MP_CPTEREGLE'));
    end
    else if TOBModePaiement <> nil then CpteGeneral := VarToStr(TOBModePaiement.GetValue('MPC_CPTEREGLE'));
  end;
  if CpteGeneral = '' then
  begin
    Result := rcPar;
    NumErreur := 8;
    Exit;
  end;

  with TOBLigneModePaie do
  begin
    PutValue('GENERAL', CpteGeneral);
    if Journal <> '' then PutValue('JOURNAL', Journal);
    // Type de piece=OD si ticket
    if IsNatCaisse(Nature) then
      PutValue('TYPEPIECE', 'OD');
    PutValue('MONTANT1', MontantReglement);
    PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
    PutValue('MONTANT2', QLC.FindField('GIC_MONTANTTTCDEV').AsFloat);
    if DebitCredit = '' then
      if GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN' then
      begin
        if TOBLignes.Detail[0].GetValue('TYPEPIECE') = 'FC' then
          DebitCredit := 'D'
        else
          DebitCredit := 'C';
      end else
      begin
        DebitCredit := 'C';
      end;
    PutValue('SENS', DebitCredit);
    PutValue('MODEPAIE', TOBLignes.Detail[1].GetValue('MODEPAIE'));
    InfosCompl(IsRegl, TOBLigneModePaie, LeCptAuxiliaire, LeCptGeneral, 'MODEPAIE', LaNatureTiers, NumLigne, '', LeTypeTiers);
    // MODIF AC 27/10/03
    // Si montant négatif inverse SENS (déja traité pour les opérations de caisse
    if (StComptaArticle = '') and (GetValue('MONTANT1')<0) then
    begin
      PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
      PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
      if GetValue('SENS')='D' then PutValue('SENS','C')
      else PutValue('SENS','D') ;
    end ;
  end;
end;

function ChargeLignesFraisPorts(TOBLignes: TOB; QLC: TQuery; Nature: string; IsRegl: boolean; NumLigne: integer): T_RetCpta;
var TOBLigneFraisPorts {,TOBCode}: TOB;
  NatureCpta, CptPortFrais: string;
  MontantPF: Double;
begin
  Result := rcOk;
  MontantPF := QLC.FindField('GIC_MONTANTPF').AsFloat;
  if MontantPF = 0 then exit;
  TOBLigneFraisPorts := TOB.Create('TOB Ligne HT', TOBLignes, -1);
  TOBLigneFraisPorts.Dupliquer(TOBLignes.Detail[0], False, True, True);
  NatureCpta := TOBLignes.Detail[0].GetValue('TYPEPIECE');
  if ((NatureCpta = 'FC') or (NatureCpta = 'AC')) then CptPortFrais := VH_GC.GCCptePortVTE
  else CptPortFrais := VH_GC.GCCptePortACH;
  // Erreur sur le compte Ports et Frais
  if CptPortFrais = '' then
  begin
    Result := rcPar;
    NumErreur := 9;
    exit;
  end;
  with TOBLigneFraisPorts do
  begin
    PutValue('GENERAL', CptPortFrais);
    PutValue('MONTANT1', MontantPF);
    PutValue('MONTANT2', QLC.FindField('GIC_MONTANTPFDEV').AsFloat);
    PutValue('QTE1', Arrondi(QLC.FindField('GIC_QTE1').AsFloat, 4));
    PutValue('TVA', QLC.FindField('GIC_FAMILLETAXE1').AsString);
    PutValue('TPF', QLC.FindField('GIC_FAMILLETAXE2').AsString);
    if GetInfoParPiece(Nature, 'GPP_VENTEACHAT') = 'VEN' then PutValue('SENS', 'C') else PutValue('SENS', 'D');
    PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
    InfosCompl(IsRegl, TOBLigneFraisPorts, LeCptAuxiliaire, LeCptGeneral, 'FRAISPORT', LaNatureTiers, NumLigne, '', LeTypeTiers);
    // MODIF AC 27/10/03
    // Si montant négatif inverse SENS (déja traité pour les opérations de caisse
    if GetValue('MONTANT1')<0 then
    begin
      PutValue('MONTANT1',Abs(GetValue('MONTANT1'))) ;
      PutValue('MONTANT2',Abs(GetValue('MONTANT2'))) ;
      if GetValue('SENS')='D' then PutValue('SENS','C')
      else PutValue('SENS','D') ;
    end ;
  end;
  //AlimRefInterneEtLibelle(TOBLigneFraisPorts,QLC,'Ports et Frais',tecHT,True);   // ???
end;

function ChargeLignesAnalytiques(TOBLignes: TOB; QLC: TQuery; Nature: string; ModePaie: Boolean; IsRegl: boolean): T_RetCpta;
var Q: TQuery;
  StSQL: string;
  MontantAna: Double;
  TOBLigneAna, TobTmp: TOB;
  NumChamp, NumLigne: integer;
begin
  NumLigne := 0;
  Result := rcOk;
  StSQL := 'SELECT * FROM INTERCOMPTA WHERE ' +
    'GIC_USER="' + V_PGI.User + '" AND ' +
    '(GIC_SECTANA1<>"" OR GIC_SECTANA2<>"" OR GIC_SECTANA3<>"" OR ' +
    'GIC_SECTANA4<>"" OR GIC_SECTANA5<>"") AND ' +
    'GIC_DATE="' + USDATETIME(QLC.FindField('GIC_DATE').AsDateTime) + '" AND ' +
    'GIC_TIERS="' + QLC.FindField('GIC_TIERS').AsString + '" AND ' +
    'GIC_ETABLISSEMENT="' + QLC.FindField('GIC_ETABLISSEMENT').AsString + '" AND ' +
    'GIC_NATUREPIECEG="' + QLC.FindField('GIC_NATUREPIECEG').AsString + '" AND ' +
    'GIC_SOUCHE="' + QLC.FindField('GIC_SOUCHE').AsString + '" AND ' +
    'GIC_NUMERO=' + QLC.FindField('GIC_NUMERO').AsString + ' AND ' +
    'GIC_REGIMETAXE="' + QLC.FindField('GIC_REGIMETAXE').AsString + '" AND ' +
    'GIC_COMPTATIERS="' + QLC.FindField('GIC_COMPTATIERS').AsString + '" AND ' +
    'GIC_COMPTAARTICLE="' + QLC.FindField('GIC_COMPTAARTICLE').AsString + '" AND ' +
    'GIC_FAMILLETAXE1="' + QLC.FindField('GIC_FAMILLETAXE1').AsString + '" AND ' +
    'GIC_FAMILLETAXE2="' + QLC.FindField('GIC_FAMILLETAXE2').AsString + '" AND ' +
    'GIC_FAMILLETAXE3="' + QLC.FindField('GIC_FAMILLETAXE3').AsString + '" AND ' +
    'GIC_FAMILLETAXE4="' + QLC.FindField('GIC_FAMILLETAXE4').AsString + '" AND ' +
    'GIC_FAMILLETAXE5="' + QLC.FindField('GIC_FAMILLETAXE5').AsString + '"';
  Q := OpenSQL(StSQL, True);
  Q.First;

  while not Q.EOF do
  begin
    if (ModePaie) and ((TOBLignes.Detail.count < 2) or (Copy(TOBLignes.Detail[2].GetValue('GENERAL'), 1, 1) <> '6')) then Break;
    if ModePaie then
      MontantAna := abs(Q.FindField('GIC_MONTANTTTC').AsFloat)
    else
      MontantAna := Q.FindField('GIC_MONTANTHT').AsFloat;
    NumChamp := RetourneNumChampAna(Q);
    if (MontantAna <> 0) and (NumChamp > 0) and (Q.FindField('GIC_ARTICLE').AsString = QLC.FindField('GIC_ARTICLE').AsString) then
    begin
      //Cherche et duplique la ligne HT pour création de la ligne analytique
      TobTmp := TOBLignes.FindFirst(['TYPELIGNE'], ['HT_VT'], True);
      if TobTmp <> nil then
      begin
        TOBLigneAna := TOB.Create('TOB Ligne Analytique', TOBLignesAnalytiques, -1); //Modif 03/06/2002
        TOBLigneAna.Dupliquer(TobTmp, False, True, True);
        // Pour les opération de caisse analytique sur les comptes 6
        with TOBLigneAna do
        begin
          PutValue('MONTANT1', MontantAna);
          PutValue('MONTANTORIGINE',GetValue('MONTANT1')) ;
          PutValue('MONTANT2', Q.FindField('GIC_MONTANTHTDEV').AsFloat);
          PutValue('QTE1', Arrondi(Q.FindField('GIC_QTE1').AsFloat, 4));
          PutValue('TYPECPTE', 'A');
          PutValue('AUXILIAIRE', Q.FindField('GIC_SECTANA' + IntToStr(NumChamp)).AsString);
          PutValue('AXE', 'A' + IntToStr(NumChamp)); //24/05/2002
          NumLigne := NumLigne + 1;
          InfosCompl(IsRegl, TOBLigneAna, LeCptAuxiliaire, LeCptGeneral, 'ANALYTIQUE', LaNatureTiers, NumLigne, '', LeTypeTiers);
        end;
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
end;

procedure RegroupeParCompte(AucunRegroupement: boolean);
var TOBParCompte, TOBSommeEtArrondi, TOBUnCompte: TOB;
  i: integer;
begin
  TOBSommeEtArrondi := TOB.Create('Grouper et Arrondi', nil, -1);
  TOBParCompte := TOB.Create('Lignes par compte', nil, -1);
  while TOBRegroupeParNature.Detail.Count > 0 do
  begin
    TOBRegroupeParNature.Detail[0].ChangeParent(TOBParCompte, -1);
    with TOBParCompte.Detail[0] do
    begin
      if AucunRegroupement then
        TOBUnCompte := TOBRegroupeParNature.FindFirst(['DATECOMPTABLE', 'GENERAL', 'AUXILIAIRE', 'NUMEROPIECE', 'MODEPAIE', 'ECHEANCE', 'ETABLISSEMENT',
          'SENS'],
            [GetValue('DATECOMPTABLE'), GetValue('GENERAL'), GetValue('AUXILIAIRE'), GetValue('NUMEROPIECE'), GetValue('MODEPAIE'), GetValue('ECHEANCE'),
          GetValue('ETABLISSEMENT'), GetValue('SENS')], False)
      else
        TOBUnCompte := TOBRegroupeParNature.FindFirst(['DATECOMPTABLE', 'GENERAL', 'AUXILIAIRE', 'MODEPAIE', 'ECHEANCE', 'ETABLISSEMENT', 'SENS'],
          [GetValue('DATECOMPTABLE'), GetValue('GENERAL'), GetValue('AUXILIAIRE'), GetValue('MODEPAIE'), GetValue('ECHEANCE'), GetValue('ETABLISSEMENT'),
          GetValue('SENS')], False);
      while TOBUnCompte <> nil do
      begin
        {Somme les montants et les quantités}
        for i := 1 to 3 do
          PutValue('MONTANT' + IntToStr(i), Valeur(GetValue('MONTANT' + IntToStr(i))) + Valeur(TOBUnCompte.GetValue('MONTANT' + IntToStr(i))));
        for i := 1 to 2 do
          PutValue('QTE' + IntToStr(i), Valeur(GetValue('QTE' + IntToStr(i))) + Valeur(TOBUnCompte.GetValue('QTE' + IntToStr(i))));
        for i := 0 to 3 do
          PutValue('LIBREMONTANT' + IntToStr(i), Valeur(GetValue('LIBREMONTANT' + IntToStr(i))) + Valeur(TOBUnCompte.GetValue('LIBREMONTANT' + IntToStr(i))));

        TOBUnCompte.Free;
        if AucunRegroupement then
          TOBUnCompte := TOBRegroupeParNature.FindNext(['DATECOMPTABLE', 'GENERAL', 'AUXILIAIRE', 'NUMEROPIECE', 'MODEPAIE', 'ECHEANCE', 'ETABLISSEMENT',
            'SENS'],
              [GetValue('DATECOMPTABLE'), GetValue('GENERAL'), GetValue('AUXILIAIRE'), GetValue('NUMEROPIECE'), GetValue('MODEPAIE'), GetValue('ECHEANCE'),
            GetValue('ETABLISSEMENT'), GetValue('SENS')], False)
        else
          TOBUnCompte := TOBRegroupeParNature.FindNext(['DATECOMPTABLE', 'GENERAL', 'AUXILIAIRE', 'MODEPAIE', 'ECHEANCE', 'ETABLISSEMENT', 'SENS'],
            [GetValue('DATECOMPTABLE'), GetValue('GENERAL'), GetValue('AUXILIAIRE'), GetValue('MODEPAIE'), GetValue('ECHEANCE'), GetValue('ETABLISSEMENT'),
            GetValue('SENS')], False);
      end;
      {Arrondi les montants}
      for i := 1 to 3 do
        if Valeur(GetValue('MONTANT' + IntToStr(i))) <> 0 then
          PutValue('MONTANT' + IntToStr(i), Arrondi(GetValue('MONTANT' + IntToStr(i)), 2));
      for i := 0 to 3 do
        if Valeur(GetValue('LIBREMONTANT' + IntToStr(i))) <> 0 then
          PutValue('LIBREMONTANT' + IntToStr(i), Arrondi(GetValue('LIBREMONTANT' + IntToStr(i)), 2));
      ChangeParent(TOBSommeEtArrondi, -1);
    end;
  end;

  while TOBSommeEtArrondi.Detail.Count > 0 do
    TOBSommeEtArrondi.Detail[0].ChangeParent(TOBRegroupeParNature, -1);

  TOBSommeEtArrondi.Free;
  TOBParCompte.Free;
end;

//Regroupe les lignes mais n'écrit plus le fichier pour pouvoir insérer les sections
//analytiques en début de fichier. Donc chargement en TOB.

procedure RegrouperEtEcrireLeFichier(NaturePiece: string);
var RegroupementNature, CompteGeneral: string;
  TOBTemp: TOB;
  i: integer;
begin
  RegroupementNature := '';
  RegroupementNature := GetInfoParPiece(NaturePiece, 'GPP_REGROUPCPTA');

  {Regroupe par compte}
  RegroupeParCompte(RegroupementNature = 'AUC');
  TOBRegroupeParNature.Detail.Sort('GENERAL;JOURNAL');

  {Traite les lignes Analytiques}
  if TOBLignesAnalytiques.Detail.Count > 0 then
  begin
    TOBTemp := TOB.Create('TOB Temporaire', nil, -1);
    while TOBRegroupeParNature.Detail.Count > 0 do TOBRegroupeParNature.Detail[0].ChangeParent(TOBTemp, -1);

    TOBLignesAnalytiques.Detail.Sort('GENERAL;AXE');
    while TOBLignesAnalytiques.Detail.Count > 0 do
      TOBLignesAnalytiques.Detail[0].ChangeParent(TOBRegroupeParNature, -1);
    RegroupeParCompte(RegroupementNature = 'AUC');
    i := 0;
    CompteGeneral := TOBRegroupeParNature.Detail[0].GetValue('GENERAL');
    while (i < TOBTemp.detail.count) do
    begin
      if TOBTemp.Detail[i].GetValue('GENERAL') = CompteGeneral then
      begin
        while (TOBRegroupeParNature.Detail.Count > 0) and (TOBRegroupeParNature.Detail[0].GetValue('GENERAL') = CompteGeneral) do
        begin
          i := i + 1;
          TOBRegroupeParNature.Detail[0].ChangeParent(TOBTemp, i);
        end;
        if TOBRegroupeParNature.Detail.Count > 0 then CompteGeneral := TOBRegroupeParNature.Detail[0].GetValue('GENERAL');
      end;
      i := i + 1;
    end;

    while TOBTemp.Detail.Count > 0 do TOBTemp.Detail[0].ChangeParent(TOBRegroupeParNature, -1);
    TOBTemp.Free;
  end;
  {Export vers TOB Finale}
  for i := TOBRegroupeParNature.Detail.Count - 1 downto 0 do
    TOBRegroupeParNature.Detail[i].ChangeParent(TOBFinale, -1);
  {Supprime les lignes qui ont été insérées dans le fichier ASCII}
  TOBRegroupeParNature.ClearDetail;
end;

procedure GestionEcartDansLesJournaux(TypeJournal: string);
var IndiceFille, Cpt: integer;
  TotalCredit, TotalDebit, TotalCreditDev, TotalDebitDev: Double;
  Journal: string;
  DateComptable: TDateTime;
  TOBL, TobTmp: TOB;
begin
  //TOBFinale.detail.Sort('DATECOMPTABLE;NUMEROPIECE;-JOURNAL;MODEPAIE') ;
  TOBFinale.detail.Sort('TYPEPIECE;SENS;-JOURNAL;TYPECPTE;DATECOMPTABLE');
  TOBL := TOBFinale.FindFirst(['TYPEPIECE'], [TypeJournal], False);
  if TOBL = nil then exit;
  Journal := TOBL.GetValue('JOURNAL');
  DateComptable := TOBL.GetValue('DATECOMPTABLE');
  IndiceFille := TOBL.GetIndex;
  while TOBL <> nil do
  begin
    TotalCredit := 0;
    TotalDebit := 0;
    TotalCreditDev := 0;
    TotalDebitDev := 0;
    for Cpt := 0 to TobFinale.detail.count - 1 do
    begin
      TobTmp := TobFinale.detail[Cpt];
      if (TobTmp.GetValue('TYPEPIECE') = TypeJournal) and (TobTmp.GetValue('JOURNAL') = Journal) and
        (TobTmp.GetValue('DATECOMPTABLE') = DateComptable) and (TobTmp.GetValue('TYPECPTE') <> 'A') then
      begin
        if TobTmp.GetValue('SENS') = 'C' then
          TotalCredit := TotalCredit + TobTmp.GetValue('MONTANT1')
        else
          Totaldebit := TotalDebit + TobTmp.GetValue('MONTANT1');
      end;
    end;
    if (TotalCredit <> TotalDebit) and (TypeJournal='OD') then
    begin
       LigneEcartJournaux(TotalCredit, TotalDebit, TotalCreditDev, TotalDebitDev, TOBL, TypeJournal);
    end ;
    while (TOBL.GetValue('JOURNAL') = Journal) and (TOBL.GetValue('DATECOMPTABLE') = DateComptable) do
    begin
      TOBL := TOBFinale.Detail[IndiceFille];
      IndiceFille := IndiceFille + 1;
      if (TOBL = nil) or (IndiceFille > TOBFinale.Detail.count - 1) then Exit;
    end;
    IndiceFille := TOBL.GetIndex;
    Journal := TOBL.GetValue('JOURNAL');
    DateComptable := TOBL.GetValue('DATECOMPTABLE');
  end;
end;

procedure LigneEcartJournaux(TotalCredit, TotalDebit, TotalCreditDev, TotalDebitDev: Double; TOBL: Tob; TypeJournal: string);
var NvTob: TOB;
NvMontant : Double ;
begin
  if TypeJournal = 'OD' then
  begin
    NvTob := TOB.Create('une lignes', TOBFinale, -1);
    NvTob.Dupliquer(TOBL, False, True);
    NvTob.PutValue('GENERAL', '471000'); //Provisoire paramètre dans Back
    NvTob.PutValue('AUXILIAIRE', '');
    NvTob.PutValue('LIBELLE', 'Ecart de caisse');
    NvTob.PutValue('MONTANT1', Abs(TotalCredit - TotalDebit));
    NvTob.PutValue('MONTANT2', Abs(TotalCreditDev - TotalDebitDev));
    NvTob.PutValue('MODEPAIE', 'ESP');
    if TotalCredit > TotalDebit then
      NvTob.PutValue('SENS', 'D')
    else
      NvTob.PutValue('SENS', 'C');
  end else
  begin
    {NvTob := TOB.Create('une lignes', TOBFinale, -1);
    NvTob.Dupliquer(TOBL, False, True);
    NvTob.PutValue('AUXILIAIRE', '');
    NvTob.PutValue('REFINTERNE', '');
    NvTob.PutValue('LIBELLE', 'Ecart de banque');
    NvTob.PutValue('MONTANT1', Abs(TotalCredit - TotalDebit));
    NvTob.PutValue('MONTANT2', Abs(TotalCreditDev - TotalDebitDev));
    NvTob.PutValue('MODEPAIE', 'CHQ'); }
    if TotalCredit > TotalDebit then
    begin
      NvMontant := TOBL.GetValue('MONTANT1') - Abs(TotalCredit - TotalDebit) ;
      TOBL.PutValue('MONTANT1',NvMontant) ;
      //NvTob.PutValue('SENS', 'D');
      //NvTob.PutValue('GENERAL','668800'); //Provisoire paramètre dans Back
      //NvTob.PutValue('GENERAL', GetParamSoc('SO_GCECARTDEBIT'));
    end else
    begin
      NvMontant := TOBL.GetValue('MONTANT1') + Abs(TotalCredit - TotalDebit) ;
      TOBL.PutValue('MONTANT1',NvMontant) ;
      TOBL.AddChampSupValeur('ECARTJRL',(TotalCredit - TotalDebit)) ;
      //NvTob.PutValue('SENS', 'C');
      //NvTob.PutValue('GENERAL','768800'); //Provisoire paramètre dans Back
      //NvTob.PutValue('GENERAL', GetParamSoc('SO_GCECARTCREDIT'));
    end;
  end;
end;

//procedure ExportTOBLignes(OptionsRegr : string);

function ExportTOBLignes(OptionsRegr: string): boolean;
var i {, NumEcriture, NewNumLigne}: integer;
  NumPiece: LongInt;
  Maj: MiseAjour;
  NvellePce {, RecalcNumLigne}: boolean;
  TobOrigine, TobEcr, TobAna: TOB;

  function TestNvPce(TobATester: TOB): boolean;
  begin
    Result := False;
    if NumPiece <> TobATester.GetValue('NUMEROPIECE') then
    begin
      NumPiece := TobATester.GetValue('NUMEROPIECE');
      Result := True;

    end;
  end;

begin
  Result := True;
  if TOBFinale.detail.count <= 0 then exit;
  //Eclatement en 2 tob (écritures et analytique)
  TobOrigine := TOB.Create('Tob generale', nil, -1);
  TobOrigine.Dupliquer(TOBFinale, True, True, True);
  TobEcr := TOB.Create('Les ecritures', nil, -1);
  TobAna := TOB.Create('L analytique', nil, -1);
  for i := TobOrigine.Detail.count - 1 downto 0 do
  begin
    if Valeur(TobOrigine.Detail[i].GetValue('NUMEROPIECE')) = 0 then
      TobOrigine.Detail[i].PutValue('NUMEROPIECE', TobOrigine.Detail[i].GetValue('COMPTEURECRITURE'));
    if TobOrigine.detail[i].GetValue('TYPECPTE') = 'A' then
    begin
      TobOrigine.Detail[i].ChangeParent(TobAna, -1);
    end else
    begin
      TobOrigine.Detail[i].ChangeParent(TobEcr, -1);
    end;
  end;
  TobEcr.detail.sort('DATECOMPTABLE;NATURECOMPTABLE;NATUREPIECEG;NUMEROPIECE;NUMEROLIGNE;NUMECHE;ECHEANCE;-JOURNAL;MODEPAIE;GENERAL;AXE');
  TobAna.detail.sort('DATECOMPTABLE;NATURECOMPTABLE;NATUREPIECEG;NUMEROPIECE;NUMEROLIGNE;NUMECHE;ECHEANCE;-JOURNAL;MODEPAIE;GENERAL;AXE');

  for i := 0 to TobEcr.Detail.count - 1 do
  begin
    NvellePce := TestNvPce(TobEcr.detail[i]);
    AddTobEcriture(TobEcr.detail[i], NvellePce, TobAna);
  end;
  for i := 0 to TobAna.Detail.count - 1 do
  begin
    NvellePce := TestNvPce(TobAna.detail[i]);
    AddTobAnalytique(TobAna.detail[i], NvellePce, i + 1);
  end;

  if Transactions(Maj.MajFiles, 0) <> oeOk then
    Result := False
  else
    if Rapport then MajRapport;
  TobEcr.free;
  TobAna.free;
  TobOrigine.free;
end;

//*****************************************************************
// DEBUT ALIMENTATION TABLES ECRITURE ET ANALYTIQ
//*****************************************************************

procedure TestRegroupements(OptionsRegr: string);
begin
  AuxiCentralisation := (ReadTokenSt(OptionsRegr) = 'X');
  DetEntreprise := (ReadTokenSt(OptionsRegr) = 'X');
  DetParticulier := (ReadTokenSt(OptionsRegr) = 'X');
  PaiemDiffere := (ReadTokenSt(OptionsRegr) = 'X');
  RgtBanque := (ReadTokenSt(OptionsRegr) = 'X');
  RgtTiers := (ReadTokenSt(OptionsRegr) = 'X');
  RegrEntreprise := False;
  RegrParticulier := False;
  CentEntreprise := False;
  CentParticulier := False;
  if AuxiCentralisation then
    CentEntreprise := True;
  if not DetParticulier then
    CentParticulier := True;

  { if AuxiCentralisation then
    begin
      if not DetEntreprise then
        CentEntreprise := True;
      if not DetParticulier then
        CentParticulier := True;
    end else
    begin
      if not DetEntreprise then
        RegrEntreprise := True;
    end;}

end;

procedure AddTobEcriture(TobL: TOB; NvellePce: boolean; TobAna: TOB);
var TobEcr: TOBM;
  TobTmp: TOB;
  NumeroPiece, Cpt: integer;
  Exercice, Journal, NaturePiece, General, Auxiliaire, TypeLigne, CritereRegr: string;
  DatePiece: TDateTime;
  InfosDevise: RDEVISE;
  MtDebit, MtCredit: double;

  procedure CalcAutresChps(CAux, CGen, ETva, ELet, TMvt: string; Dte: TDateTime; NEche: integer);
  begin
    TobEcr.PutValue('E_CONTREPARTIEAUX', CAux);
    TobEcr.PutValue('E_CONTREPARTIEGEN', CGen);
    TobEcr.PutValue('E_EMETTEURTVA', ETva);
    TobEcr.PutValue('E_ETATLETTRAGE', ELet);
    TobEcr.PutValue('E_TYPEMVT', TMvt);
    TobEcr.PutValue('E_DATEPAQUETMAX', Dte);
    TobEcr.PutValue('E_DATEPAQUETMIN', Dte);
    TobEcr.PutValue('E_NUMECHE', NEche);
  end;

begin
  TypeLigne := TobL.GetValue('TYPELIGNE');
  DatePiece := TobL.GetValue('DATECOMPTABLE');
  NaturePiece := TobL.GetValue('TYPEPIECE');
  General := TobL.GetValue('GENERAL');
  Auxiliaire := TobL.GetValue('AUXILIAIRE');
  Journal := TobL.GetValue('JOURNAL');
  NumeroPiece := TobL.GetValue('NUMEROPIECE');
  Exercice := QuelExoDt(DatePiece);
  InfosDevise.Code := TobL.GetValue('DEVISE');
  GetInfosDevise(InfosDevise);
  InfosDevise.Taux := GetTaux(InfosDevise.Code, InfosDevise.DateTaux, DatePiece);

  TobEcr := TOBM.Create(EcrGen, '', True);
  TobEcr.PutValue('E_ETABLISSEMENT', TobL.GetValue('ETABLISSEMENT'));
  TobEcr.PutValue('E_EXERCICE', Exercice);
  TobEcr.PutValue('E_PERIODE', GetPeriode(DatePiece));
  TobEcr.PutValue('E_SEMAINE', NumSemaine(DatePiece));
  TobEcr.AddChampSupValeur('MONTANT1', TobL.GetValue('MONTANTORIGINE'), False);
  CalcDebCred(TobL, TobEcr, 'E', nil, InfosDevise);
  if TobEcr.GetValue('E_DEBIT') + TobEcr.GetValue('E_CREDIT') > 0 then
  begin
    MtDebit := Valeur(TobEcr.GetValue('E_DEBIT'));
    MtCredit := Valeur(TobEcr.GetValue('E_CREDIT'));
    {Si même pièce, cumul montant débit et crédit pour calcul E_ENCAISSEMENT}
    if not NvellePce then
    begin
      TobTmp := TobEcriture.FindFirst(['E_NUMEROPIECE'], [NumeroPiece], True);
      while TobTmp <> nil do
      begin
        MtDebit := MtDebit + Valeur(TobTmp.GetValue('E_DEBIT'));
        MtCredit := MtCredit + Valeur(TobTmp.GetValue('E_CREDIT'));
        TobTmp := TobEcriture.FindNext(['E_NUMEROPIECE'], [NumeroPiece], True);
      end;
    end;
    TobEcr.PutValue('E_ENCAISSEMENT', SensEnc(MtDebit, MtCredit));
    TobEcr.PutValue('E_VALIDE', '-');
    TobEcr.PutValue('E_JOURNAL', Journal);
    TobEcr.PutValue('E_DATECOMPTABLE', DatePiece);
    TobEcr.PutValue('E_DATEREFEXTERNE', DatePiece);
    TobEcr.PutValue('E_NATUREPIECE', NaturePiece);
    TobEcr.PutValue('E_GENERAL', General);
    TobEcr.PutValue('E_AUXILIAIRE', Auxiliaire);
    TobEcr.PutValue('E_REFINTERNE', TobL.GetValue('REFINTERNE'));
    if not AuxiCentralisation then
      TobEcr.PutValue('E_LIBELLE', TobL.GetValue('LIBELLE'))
    else
      TobEcr.PutValue('E_LIBELLE', '');
    TobEcr.PutValue('E_MODEPAIE', TobL.GetValue('MODEPAIE'));
    TobEcr.PutValue('E_DATEECHEANCE', TobL.GetValue('ECHEANCE'));
    TobEcr.PutValue('E_QUALIFPIECE', TypeEcriture);
    TobEcr.PutValue('E_NUMEROPIECE', NumeroPiece);
    TobEcr.PutValue('E_DATECREATION', TobL.GetValue('DATECREATION'));
    TobEcr.PutValue('E_SOCIETE', TobL.GetValue('SOCIETE'));
    TobEcr.PutValue('E_AFFAIRE', TobL.GetValue('AFFAIRE'));
    TobEcr.PutValue('E_DATETAUXDEV', TobL.GetValue('DATETAUXDEV'));
    TobEcr.PutValue('E_ECRANOUVEAU', TobL.GetValue('ECRANOUVEAU'));
    TobEcr.PutValue('E_QTE1', TobL.GetValue('QTE1'));
    TobEcr.PutValue('E_QTE2', TobL.GetValue('QTE2'));
    TobEcr.PutValue('E_QUALIFQTE1', TobL.GetValue('QUALIFQTE1'));
    TobEcr.PutValue('E_QUALIFQTE2', TobL.GetValue('QUALIFQTE2'));
    TobEcr.PutValue('E_REFLIBRE', TobL.GetValue('REFLIBRE'));
    TobEcr.PutValue('E_TVAENCAISSEMENT', TobL.GetValue('TVAENCAISSEMENT'));
    TobEcr.PutValue('E_REGIMETVA', TobL.GetValue('REGIMETVA'));
    TobEcr.PutValue('E_TVA', TobL.GetValue('TVA'));
    TobEcr.PutValue('E_TPF', TobL.GetValue('TPF'));
    TobEcr.PutValue('E_REFPOINTAGE', TobL.GetValue('REFPOINTAGE'));
    TobEcr.PutValue('E_DATEPOINTAGE', TobL.GetValue('DATEPOINTAGE'));
    TobEcr.PutValue('E_DATERELANCE', TobL.GetValue('DATERELANCE'));
    TobEcr.PutValue('E_DATEVALEUR', TobL.GetValue('DATEVALEUR'));
    TobEcr.PutValue('E_REFRELEVE', TobL.GetValue('REFRELEVE'));
    TobEcr.PutValue('E_NUMEROIMMO', Valeur(TobL.GetValue('NUMEROIMMO')));
    for Cpt := 0 to 9 do
      TobEcr.PutValue('E_LIBRETEXTE' + IntToStr(Cpt), TobL.GetValue('LIBRETEXTE' + IntToStr(Cpt)));
    for Cpt := 0 to 3 do
      TobEcr.PutValue('E_TABLE' + IntToStr(Cpt), TobL.GetValue('TABLE' + IntToStr(Cpt)));
    for Cpt := 0 to 3 do
      TobEcr.PutValue('E_LIBREMONTANT' + IntToStr(Cpt), Valeur(TobL.GetValue('LIBREMONTANT' + IntToStr(Cpt))));
    for Cpt := 0 to 1 do
    begin
      if TobL.GetValue('LIBREBOOL' + IntToStr(Cpt)) = '' then
        TobEcr.PutValue('E_LIBREBOOL' + IntToStr(Cpt), '-')
      else
        TobEcr.PutValue('E_LIBREBOOL' + IntToStr(Cpt), TobL.GetValue('LIBREBOOL' + IntToStr(Cpt)));
    end;
    TobEcr.PutValue('E_LIBREDATE', TobL.GetValue('LIBREDATE'));
    TobEcr.PutValue('E_CONSO', TobL.GetValue('CONSO'));
    TobEcr.PutValue('E_COUVERTURE', Valeur(TobL.GetValue('COUVERTURE')));
    TobEcr.PutValue('E_COUVERTUREDEV', Valeur(TobL.GetValue('COUVERTUREDEV')));
    //TobEcr.PutValue('E_COUVERTUREEURO', Valeur(TobL.GetValue('COUVERTUREEURO')));
    TobEcr.PutValue('E_LETTRAGE', TobL.GetValue('LETTRAGE'));
    if TobL.GetValue('LETTRAGEDEV') = '' then
      TobEcr.PutValue('E_LETTRAGEDEV', '-')
    else
      TobEcr.PutValue('E_LETTRAGEDEV', TobL.GetValue('LETTRAGEDEV'));
    //TobEcr.PutValue('E_LETTRAGEEURO', TobL.GetValue('LETTRAGEEURO'));
    TobEcr.PutValue('E_REFGESCOM', TobL.GetValue('REFGESCOM'));
    TobEcr.PutValue('E_COTATION', InfosDevise.Taux);
    TobEcr.PutValue('E_DEVISE', TobL.GetValue('DEVISE'));
    TobEcr.PutValue('E_TAUXDEV', TobL.GetValue('TAUXDEV'));
    TobTmp := TobDesGeneraux.FindFirst(['G_GENERAL', 'G_VENTILABLE'], [General, 'X'], True);
    if TobTmp <> nil then
      TobEcr.PutValue('E_ANA', 'X')
    else
      TobEcr.PutValue('E_ANA', '-');
    TobTmp := TobJournaux.FindFirst(['J_JOURNAL'], [Journal], True);
    if TobTmp <> nil then
      TobEcr.PutValue('E_MODESAISIE', TobTmp.GetValue('J_MODESAISIE'))
    else
      TobEcr.PutValue('E_MODESAISIE', '-');
    TobEcr.PutValue('E_NUMLIGNE', TobL.GetValue('NUMEROLIGNE'));
    if (TypeLigne = 'HT_VT') or (TypeLigne = 'REMISE_VT') or (TypeLigne = 'ESCOMPTE_VT') then
      CalcAutresChps(TobL.GetValue('LECPTAUXI'), TobL.GetValue('LECPTGENE'), '-', 'RI', 'HT', TobL.GetValue('DATEPAQUETMAX'), 0)
    else if TypeLigne = 'TAXE_VT' then
      CalcAutresChps(TobL.GetValue('LECPTAUXI'), TobL.GetValue('LECPTGENE'), '-', 'RI', TobL.GetValue('TYPETAXE'), TobL.GetValue('DATEPAQUETMAX'), 0)
    else if TypeLigne = 'TIERS_VT' then
      CalcAutresChps(TobL.GetValue('CONTREPARTIEAUX'), TobL.GetValue('CONTREPARTIEGEN'), 'X', 'AL', 'TTC', DatePiece, TobL.GetValue('NUMECHE'));
    if (Pos('_RG', TypeLigne) > 0) and (TobL.GetValue('NUMECHE') <= 1) and (TobL.GetValue('NATURECOMPTABLE') = 'BTQ') then
      TobEcr.PutValue('E_NUMECHE', 1);
    if TobEcr.GetValue('E_NUMECHE') > 0 then
      TobEcr.PutValue('E_ECHE', 'X')
    else
      TobEcr.PutValue('E_ECHE', '-');
    TobEcr.PutValue('E_CREERPAR', 'EXP');
    TobEcr.PutValue('E_QUALIFORIGINE', 'GC');
    TobEcr.PutValue('E_CONTROLETVA', 'RIE');
    TobEcr.PutValue('E_VISION', 'DEM');
    TobEcr.PutValue('E_RIB', TobL.GetValue('RIB'));
    CritereRegr := TobEcr.GetValue('E_ETABLISSEMENT') + ';' +
      TobEcr.GetValue('E_JOURNAL') + ';' +
      DateToStr(TobEcr.GetValue('E_DATECOMPTABLE')) + ';' +
      TobEcr.GetValue('E_GENERAL') + ';' +
      TobL.GetValue('LECPTAUXI') + ';' +
      TobL.GetValue('NATUREPIECEG') + ';' +
      TobEcr.GetValue('E_QUALIFPIECE') + ';' +
      TobEcr.GetValue('E_DEVISE') + ';' +
      FloatToStr(TobEcr.GetValue('E_TAUXDEV')) + ';' +
      TobEcr.GetValue('E_REGIMETVA') + ';' +
      TobEcr.GetValue('E_TYPEMVT');
    if TobL.GetValue('TYPELIGNE') = 'MODEPAIE_RG' then
      CritereRegr := CritereRegr + ';' + TobEcr.GetValue('E_GENERAL') + ';'
    else if TobL.GetValue('TYPELIGNE') = 'TIERS_RG' then
      CritereRegr := CritereRegr + ';;' + TobL.GetValue('LECPTAUXI')
    else
      CritereRegr := CritereRegr + ';;';
    TobEcr.AddChampSupValeur('CRITERESREGR', CritereRegr, False);
    TobEcr.AddChampSupValeur('TYPEMVT', TobL.GetValue('NATURECOMPTABLE'), False);
    TobEcr.AddChampSupValeur('TYPELIGNE', TobL.GetValue('TYPELIGNE'), False);
    TobEcr.AddChampSupValeur('SENS', TobL.GetValue('SENS'), False);
    //Momentané (test de la clé primaire dans le champs AFFAIRE)
    TobEcr.PutValue('E_AFFAIRE', TobEcr.GetValue('E_JOURNAL') + ';' + TobEcr.GetValue('E_EXERCICE') + ';' + DateToStr(TobEcr.GetValue('E_DATECOMPTABLE')) + ';'
      + IntToStr(TobEcr.GetValue('E_NUMEROPIECE')) + ';' + IntToStr(TobEcr.GetValue('E_NUMLIGNE')) + ';' + IntToStr(TobEcr.GetValue('E_NUMECHE')) + ';' +
      TobEcr.GetValue('E_QUALIFPIECE'));
    if not RegroupmentEcriture(TobL, TobEcr, TobAna) then
    begin
      TobTmp := TOB.Create('', TobEcriture.Detail[0], -1);
      TobTmp.Dupliquer(TobEcr, True, True, True);
    end;
  end;
  TobEcr.Free;
end;

procedure AddTobAnalytique(TobL: TOB; NvellePce: boolean; NumLigne: integer);
var TobAna, TobEcr, TobTmp: TOB;
  NumeroPiece, Cpt: integer;
  Exercice, Journal, NaturePiece, General, Auxiliaire, TypeLigne: string;
  DatePiece: TDateTime;
  MtPourc: double;
  InfosDevise: RDevise;
begin
  TypeLigne := TobL.GetValue('TYPELIGNE');
  DatePiece := TobL.GetValue('DATECOMPTABLE');
  NaturePiece := TobL.GetValue('TYPEPIECE');
  General := TobL.GetValue('GENERAL');
  Auxiliaire := TobL.GetValue('AUXILIAIRE');
  Journal := TobL.GetValue('JOURNAL');
  NumeroPiece := TobL.GetValue('NUMEROPIECE');
  Exercice := QuelExoDt(DatePiece);
  //Recherche la ligne correspondante dans la TobEcriture
  TobEcr := TobEcriture.FindFirst(['E_NUMEROPIECE', 'E_GENERAL', 'E_ANA', 'E_DATECOMPTABLE', 'E_ETABLISSEMENT'],
    [NumeroPiece, TobL.GetValue('GENERAL'), 'X', DatePiece, TobL.GetValue('ETABLISSEMENT')],
    True);
  if TobEcr = nil then exit;
  InfosDevise.Code := TobEcr.GetValue('E_DEVISE');
  GetInfosDevise(InfosDevise);
  InfosDevise.Taux := GetTaux(InfosDevise.Code, InfosDevise.DateTaux, DatePiece);
  TobAna := TOBM.Create(EcrAna, TobL.GetValue('AXE'), True);
  TobAna.PutValue('Y_GENERAL', TobL.GetValue('GENERAL'));
  TobAna.PutValue('Y_AXE', TobL.GetValue('AXE'));
  TobAna.PutValue('Y_DATECOMPTABLE', DatePiece);
  TobAna.PutValue('Y_NUMEROPIECE', NumeroPiece);
  TobAna.PutValue('Y_NUMLIGNE', Valeur(TobEcr.GetValue('E_NUMLIGNE')));
  TobAna.PutValue('Y_NUMVENTIL', NumLigne);
  TobAna.PutValue('Y_SECTION', Auxiliaire);
  TobAna.PutValue('Y_EXERCICE', Exercice);
  CalcDebCred(TobL, TobAna, 'Y', TobEcr, InfosDevise);
  TobAna.PutValue('Y_REFINTERNE', TobL.GetValue('REFINTERNE'));
  TobAna.PutValue('Y_LIBELLE', TobL.GetValue('LIBELLE'));
  TobAna.PutValue('Y_NATUREPIECE', TobL.GetValue('TYPEPIECE'));
  TobAna.PutValue('Y_QUALIFPIECE', TobEcr.GetValue('E_QUALIFPIECE'));
  TobAna.PutValue('Y_VALIDE', TobEcr.GetValue('E_VALIDE'));
  TobAna.PutValue('Y_ETAT', '0000000000');
  TobAna.PutValue('Y_REFEXTERNE', TobL.GetValue('REFEXTERNE'));
  TobAna.PutValue('Y_DATEREFEXTERNE', DatePiece);
  TobAna.PutValue('Y_DATECREATION', TobL.GetValue('DATECREATION'));
  TobAna.PutValue('Y_SOCIETE', TobL.GetValue('SOCIETE'));
  TobAna.PutValue('Y_ETABLISSEMENT', TobL.GetValue('ETABLISSEMENT'));
  TobAna.PutValue('Y_VISION', '-');
  TobAna.PutValue('Y_AFFAIRE', TobL.GetValue('AFFAIRE'));
  TobAna.PutValue('Y_REFLIBRE', TobL.GetValue('REFLIBRE'));
  TobAna.PutValue('Y_DEVISE', TobL.GetValue('DEVISE'));
  TobAna.PutValue('Y_TAUXDEV', TobL.GetValue('TAUXDEV'));
  TobAna.PutValue('Y_DATETAUXDEV', TobL.GetValue('DATETAUXDEV'));
  TobAna.PutValue('Y_TOTALECRITURE', TobEcr.GetValue('E_DEBIT') + TobEcr.GetValue('E_CREDIT'));
  if TobAna.GetValue('Y_TOTALECRITURE') > 0 then
  begin
    if TobAna.GetValue('Y_CREDIT') > 0 then
      MtPourc := TobAna.GetValue('Y_CREDIT')
    else
      MtPourc := TobAna.GetValue('Y_DEBIT');
    TobAna.PutValue('Y_POURCENTAGE', Arrondi(100.0 * MtPourc / TobAna.GetValue('Y_TOTALECRITURE'), 4));
    TobAna.PutValue('Y_POURCENTQTE1', Arrondi(TobAna.GetValue('Y_POURCENTAGE'), 0));
  end else
  begin
    TobAna.PutValue('Y_POURCENTAGE', 0);
    TobAna.PutValue('Y_POURCENTQTE1', 0);
  end;
  TobAna.PutValue('Y_TOTALDEVISE', TobEcr.GetValue('E_DEBITDEV') + TobEcr.GetValue('E_CREDITDEV'));
  //TobAna.PutValue('Y_TOTALEURO', TobEcr.GetValue('E_DEBITEURO') + TobEcr.GetValue('E_CREDITEURO'));
  TobAna.PutValue('Y_CONTROLE', '-');
  for Cpt := 1 to 2 do
    TobAna.PutValue('Y_QTE' + IntToStr(Cpt), 0);
  for Cpt := 1 to 2 do
    TobAna.PutValue('Y_QUALIFQTE' + IntToStr(Cpt), '');
  TobAna.PutValue('Y_JOURNAL', TobL.GetValue('JOURNAL'));
  TobAna.PutValue('Y_TYPEMVT', TobEcr.GetValue('E_TYPEMVT'));
  TobAna.PutValue('Y_ECRANOUVEAU', TobL.GetValue('ECRANOUVEAU'));
  TobAna.PutValue('Y_CONFIDENTIEL', '0');
  TobAna.PutValue('Y_CREERPAR', 'EXP');
  //TobAna.PutValue('Y_SAISIEEURO', '-');
  for Cpt := 0 to 9 do
    TobAna.PutValue('Y_LIBRETEXTE' + IntToStr(Cpt), TobL.GetValue('LIBRETEXTE' + IntToStr(Cpt)));
  for Cpt := 0 to 3 do
    TobAna.PutValue('Y_TABLE' + IntToStr(Cpt), TobL.GetValue('TABLE' + IntToStr(Cpt)));
  TobAna.PutValue('Y_LIBREDATE', TobL.GetValue('LIBREDATE'));
  for Cpt := 0 to 1 do
  begin
    if TobL.GetValue('LIBREBOOL' + IntToStr(Cpt)) = '' then
      TobAna.PutValue('Y_LIBREBOOL' + IntToStr(Cpt), '-')
    else
      TobAna.PutValue('Y_LIBREBOOL' + IntToStr(Cpt), TobL.GetValue('LIBREBOOL' + IntToStr(Cpt)));
  end;
  TobAna.PutValue('Y_CONSO', TobL.GetValue('CONSO'));
  for Cpt := 0 to 3 do
    TobAna.PutValue('Y_LIBREMONTANT' + IntToStr(Cpt), Valeur(TobL.GetValue('LIBREMONTANT' + IntToStr(Cpt))));
  TobAna.PutValue('Y_CONTREPARTIEGEN', TobEcr.GetValue('E_CONTREPARTIEGEN'));
  TobAna.PutValue('Y_CONTREPARTIEAUX', TobEcr.GetValue('E_CONTREPARTIEAUX'));
  TobAna.PutValue('Y_AUXILIAIRE', '');
  TobAna.PutValue('Y_PERIODE', GetPeriode(DatePiece));
  TobAna.PutValue('Y_SEMAINE', NumSemaine(DatePiece));
  TobAna.PutValue('Y_QTE1', TobEcr.GetValue('E_QTE1'));
  TobAna.PutValue('Y_QTE2', TobEcr.GetValue('E_QTE2'));
  TobAna.PutValue('Y_TOTALQTE1', TobEcr.GetValue('E_QTE1'));
  TobAna.PutValue('Y_TOTALQTE2', TobEcr.GetValue('E_QTE2'));
  TobAna.PutValue('Y_QUALIFECRQTE1', '');
  TobAna.PutValue('Y_QUALIFECRQTE2', '');
  //  if not RegroupementAnalytique(TobL, TobAna) then
  //  begin
  TobTmp := TOB.Create('', TobEcriture.Detail[1], -1);
  TobTmp.Dupliquer(TobAna, True, True, True);
  //  end;
  TobAna.free;
end;

procedure CalcDebCred(TobL, TobAAffecter: TOB; Prefixe: string; TobEcr: TOB; InfosDevise: RDevise);
var NatureCpta, TypeLigne: string;
  Mt1, Mt2, Credit, Debit: double;

  procedure CalcAna;
  begin
    if TobEcr.GetValue('E_CREDIT') > 0 then //Doit être écrit dans le même sens que la ligne liée
    begin
      Credit := Mt1;
      Debit := 0;
    end else
    begin
      Credit := 0;
      Debit := Mt1;
    end;
  end;

begin
  Mt1 := TobL.GetValue('MONTANT1');
  Credit := 0;
  Debit := 0;
  if Mt1 < 0 then
  begin
    if TobL.GetValue('SENS') = 'D' then
      Credit := abs(Mt1)
    else
      Debit := abs(Mt1);
  end else
  begin
    if TobL.GetValue('SENS') = 'D' then
      Debit := abs(Mt1)
    else
      Credit := abs(Mt1);
  end;
  CSetMontants(TobAAffecter, Debit, Credit, InfosDevise, False);
  exit;

  NatureCpta := TobL.GetValue('TYPEPIECE');
  TypeLigne := TobL.getValue('TYPELIGNE');
  //Facture client
  if NatureCpta = 'FC' then
  begin
    if (copy(TypeLigne, 1, 5) = 'TIERS') or (copy(TypeLigne, 1, 6) = 'REMISE') or (copy(TypeLigne, 1, 8) = 'ESCOMPTE') then
    begin
      Credit := 0;
      Debit := Mt1;
    end else
      if (copy(TypeLigne, 1, 2) = 'HT') or (copy(TypeLigne, 1, 4) = 'TAXE') or (copy(TypeLigne, 1, 9) = 'FRAISPORT') then
    begin
      Credit := Mt1;
      Debit := 0;
    end else
      if copy(TypeLigne, 1, 8) = 'MODEPAIE' then
    begin
      Credit := 0;
      Debit := Mt1;
    end else
      if copy(TypeLigne, 1, 10) = 'ANALYTIQUE' then
    begin
      CalcAna;
    end;
  end else
    //Avoir client, acompte, règlement, facture fournisseur
    if (NatureCpta = 'AC') or (NatureCpta = 'OC') or (NatureCpta = 'RC') or (NatureCpta = 'FF') then
  begin
    if (copy(TypeLigne, 1, 5) = 'TIERS') or (copy(TypeLigne, 1, 6) = 'REMISE') or (copy(TypeLigne, 1, 8) = 'ESCOMPTE') then
    begin
      Credit := Mt1;
      Debit := 0;
    end else
      if (copy(TypeLigne, 1, 2) = 'HT') or (copy(TypeLigne, 1, 4) = 'TAXE') or (copy(TypeLigne, 1, 9) = 'FRAISPORT') then
    begin
      Credit := 0;
      Debit := Mt1;
    end else
      if copy(TypeLigne, 1, 8) = 'MODEPAIE' then
    begin
      Credit := 0;
      Debit := Mt1;
    end else
      if copy(TypeLigne, 1, 10) = 'ANALYTIQUE' then
    begin
      CalcAna;
    end;
  end;
  CSetMontants(TobAAffecter, Debit, Credit, InfosDevise, False);
end;

procedure InfosCompl(IsRegl: boolean; TobAAffecter: TOB; LeCptAuxi, LeCptGene, TypeLigne, LaNatTiers: string; NumLigne: integer; Rib, LeTypTiers: string);
begin
  TobAAffecter.PutValue('LECPTGENE', LeCptGene);
  TobAAffecter.PutValue('LECPTAUXI', LeCptAuxi);
  TobAAffecter.PutValue('NATUREAUXI', LaNatTiers);
  if IsRegl then
  begin
    TobAAffecter.PutValue('TYPELIGNE', TypeLigne + '_RG');
    TobAAffecter.PutValue('COMPTEURECRITURE', CptEcritureRgt);
    TobAAffecter.PutValue('TYPEPIECE', 'OC'); //OC=Acompte client, RC=Rgt client
    TobAAffecter.PutValue('NUMEROPIECE', CptEcritureRgt);
  end else
  begin
    TobAAffecter.PutValue('TYPELIGNE', TypeLigne + '_VT');
    TobAAffecter.PutValue('COMPTEURECRITURE', CptEcritureVen);
    TobAAffecter.PutValue('NUMEROPIECE', CptEcritureVen);
  end;
  TobAAffecter.PutValue('NUMEROLIGNE', NumLigne);
  TobAAffecter.PutValue('RIB', Rib);
  if LeTypTiers = 'X' then
    TobAAffecter.PutValue('TYPETIERS', 'PARTICULIER')
  else
    TobAAffecter.PutValue('TYPETIERS', 'ENTREPRISE');
end;

function RegroupmentEcriture(TobL, TobEcr, TobAna: TOB): boolean;
var TobTmp, TobTmp1: TOB;

  procedure Cumul;
  var TobTmp2: TOB;
    OldNum, NewNum: integer;

    procedure ModifNumAna;
    var TobTmp3: TOB;
    begin
      {Si ligne avec analytique, modif aussi du n° pièce des lignes analytique}
      if TobEcr.GetValue('E_ANA') = 'X' then
      begin
        TobTmp3 := TobAna.FindFirst(['JOURNAL', 'NUMEROPIECE'], [TobEcr.GetValue('E_JOURNAL'), OldNum], True);
        while TobTmp3 <> nil do
        begin
          TobTmp3.PutValue('NUMEROPIECE', NewNum);
          TobTmp3.PutValue('NUMECRITURE', NewNum);
          TobTmp3 := TobAna.FindNext(['JOURNAL', 'NUMEROPIECE'], [TobEcr.GetValue('E_JOURNAL'), OldNum], True);
        end;
      end;
    end;

  begin
    Result := True;
    OldNum := TobEcr.GetValue('E_NUMEROPIECE');
    NewNum := TobTmp1.GetValue('E_NUMEROPIECE');
    TobTmp1.PutValue('E_DEBIT', Valeur(TobTmp1.GetValue('E_DEBIT')) + Valeur(TobEcr.GetValue('E_DEBIT')));
    TobTmp1.PutValue('E_CREDIT', Valeur(TobTmp1.GetValue('E_CREDIT')) + Valeur(TobEcr.GetValue('E_CREDIT')));
    //TobTmp1.PutValue('E_DEBITEURO', Valeur(TobTmp1.GetValue('E_DEBITEURO')) + Valeur(TobEcr.GetValue('E_DEBITEURO')));
    //TobTmp1.PutValue('E_CREDITEURO', Valeur(TobTmp1.GetValue('E_CREDITEURO')) + Valeur(TobEcr.GetValue('E_CREDITEURO')));
    TobTmp1.PutValue('E_DEBITDEV', Valeur(TobTmp1.GetValue('E_DEBITDEV')) + Valeur(TobEcr.GetValue('E_DEBITDEV')));
    TobTmp1.PutValue('E_CREDITDEV', Valeur(TobTmp1.GetValue('E_CREDITDEV')) + Valeur(TobEcr.GetValue('E_CREDITDEV')));
    TobTmp1.PutValue('E_REFINTERNE', '');
    if TobL.GetValue('TYPELIGNE') <> 'TIERS_RG' then
      TobTmp1.PutValue('E_LIBELLE', '');
    TobTmp1.PutValue('E_RIB', '');
    TobTmp1.PutValue('E_REFGESCOM', '');
    {Si on cumul mode paiement différent et ce n'est pas le même, on vide}
    if (PaiemDiffere) and (TobTmp1.GetValue('E_MODEPAIE') <> '') and (TobTmp1.GetValue('E_MODEPAIE') <> TobEcr.GetValue('E_MODEPAIE')) then
      TobTmp1.PutValue('E_MODEPAIE', '');
    ModifNumAna;
    {Recherche dans les écritures pour changer le n° les éventuelles lignes de la même pièce (pour écriture et analytique)}
    TobTmp2 := TobEcriture.Detail[0].FindFirst(['E_NUMEROPIECE'], [OldNum], True);
    while TobTmp2 <> nil do
    begin
      TobTmp2.PutValue('E_NUMEROPIECE', NewNum);
      ModifNumAna;
      TobTmp2 := TobEcriture.Detail[0].FindNext(['E_NUMEROPIECE'], [OldNum], True);
    end;
  end;

begin
  Result := False;
  Exit ; // MODIF AC en attendant traitement des regroupement
  {VENTE}
  if pos('_VT', TobL.GetValue('TYPELIGNE')) > 0 then
  begin
    //Test s'il y a un regroupement défini dans le ParPiece
    TobTmp := TobRegrParPiece_.FindFirst(['LANATURE', 'LEREGROUPEMENT'], [TobL.GetValue('NATUREPIECEG'), 'NON'], True);
    if TobTmp = nil then exit;
    //Test si on veut garder le détail par type de tiers
    if ((TobL.GetValue('TYPETIERS') = 'ENTREPRISE') and (not DetEntreprise)) or
      ((TobL.GetValue('TYPETIERS') = 'PARTICULIER') and (not DetParticulier)) then
      TobTmp1 := TobEcriture.Detail[0].FindFirst(['CRITERESREGR'], [TobEcr.GetValue('CRITERESREGR')], True);
    if TobTmp1 <> nil then
      cumul;
  end else
    {REGLEMENT}
    if pos('_RG', TobL.GetValue('TYPELIGNE')) > 0 then
  begin
    {Regroupement sur le compte de banque (prioritaire sur regroupement sur le compte de tiers)}
    if RgtBanque then
    begin
      if TobL.GetValue('TYPELIGNE') = 'MODEPAIE_RG' then
      begin
        {Ne pas regrouper les modes de paiements différents}
        if not PaiemDiffere then
          TobTmp1 := TobEcriture.Detail[0].FindFirst(['E_DATECOMPTABLE', 'E_JOURNAL', 'E_NATUREPIECE', 'E_GENERAL', 'E_MODEPAIE'],
            [TobEcr.GetValue('E_DATECOMPTABLE'), TobEcr.GetValue('E_JOURNAL'), TobEcr.GetValue('E_NATUREPIECE'), TobEcr.GetValue('E_GENERAL'),
              TobEcr.GetValue('E_MODEPAIE')], True)
        else
          {Regrouper les modes de paiements différents}
          TobTmp1 := TobEcriture.Detail[0].FindFirst(['E_DATECOMPTABLE', 'E_JOURNAL', 'E_NATUREPIECE', 'E_GENERAL'],
            [TobEcr.GetValue('E_DATECOMPTABLE'), TobEcr.GetValue('E_JOURNAL'), TobEcr.GetValue('E_NATUREPIECE'), TobEcr.GetValue('E_GENERAL')], True);
        if TobTmp1 <> nil then
          cumul;
      end;
      {Regroupement sur tiers après dans un regroupement sur banque}
      if (TobL.GetValue('TYPELIGNE') = 'TIERS_RG') and (RgtTiers) then
      begin
        {Ne pas regrouper les modes de paiements différents}
        if not PaiemDiffere then
          TobTmp1 := TobEcriture.Detail[0].FindFirst(['E_DATECOMPTABLE', 'E_JOURNAL', 'E_NATUREPIECE', 'E_GENERAL', 'E_AUXILIAIRE', 'E_MODEPAIE'],
            [TobEcr.GetValue('E_DATECOMPTABLE'), TobEcr.GetValue('E_JOURNAL'), TobEcr.GetValue('E_NATUREPIECE'), TobEcr.GetValue('E_GENERAL'),
              TobEcr.GetValue('E_AUXILIAIRE'), TobEcr.GetValue('E_MODEPAIE')], True)
        else
          {Regrouper les modes de paiements différents}
          TobTmp1 := TobEcriture.Detail[0].FindFirst(['E_DATECOMPTABLE', 'E_JOURNAL', 'E_NATUREPIECE', 'E_GENERAL', 'E_AUXILIAIRE'],
            [TobEcr.GetValue('E_DATECOMPTABLE'), TobEcr.GetValue('E_JOURNAL'), TobEcr.GetValue('E_NATUREPIECE'), TobEcr.GetValue('E_GENERAL'),
              TobEcr.GetValue('E_AUXILIAIRE')], True);
        if TobTmp1 <> nil then
          cumul;
      end;

    end else
      {Regroupement sur le compte de tiers}
      if RgtTiers then
    begin
      TobTmp1 := TobEcriture.FindFirst(['AUXILIAIRE'], [TobEcr.GetValue('AUXILIAIRE')], True);
      if TobTmp1 <> nil then
        cumul;
    end;
  end;
end;

procedure MiseAJour.MajFiles;
var NumEcriture, NewNumPiece: LongInt;
  Cpt, NewNumLigne, i: integer;
  TobTmp, TobTmp1, TobTmp2, TobTmp3, TobTstEquilibre, TobEcart: TOB;
  Jnal, CpteGene, CpteAux, PathFile, ExpJrnx: string;
  ExportTiers, NewEcriture: Boolean;

  procedure TestEquilibre;
  var TotCredit, TotDebit, Diff: double;
    TobLigEqu, TobLigEcr, TobLigAna: TOB;
    Cpt: integer;
  begin
    TotCredit := 0;
    TotDebit := 0;
    for Cpt := 0 to TobTstEquilibre.detail.count - 1 do
    begin
      if TobTstEquilibre.detail[cpt].FieldExists('ASUPPRIMER') then continue ;
      TotCredit := TotCredit + TobTstEquilibre.detail[cpt].GetValue('E_CREDIT');
      TotDebit := TotDebit + TobTstEquilibre.detail[cpt].GetValue('E_DEBIT');
    end;
    if TotCredit <> TotDebit then
    begin
      Diff := abs(TotCredit - TotDebit);
      TobLigEqu := TobTstEquilibre.FindFirst(['TYPELIGNE'], ['HT_VT'], true);
      TobLigEcr := TobEcriture.Detail[0].FindFirst(['E_DATECOMPTABLE','E_JOURNAL','E_NUMEROPIECE','E_NUMLIGNE','TYPELIGNE'],
                [TobLigEqu.GetValue('E_DATECOMPTABLE'),TobLigEqu.GetValue('E_JOURNAL'),TobLigEqu.GetValue('E_NUMEROPIECE'),TobLigEqu.GetValue('E_NUMLIGNE'),TobLigEqu.GetValue('TYPELIGNE')],True) ;
      TobLigAna := TobEcriture.Detail[1].FindFirst(['Y_DATECOMPTABLE','Y_JOURNAL','Y_NUMEROPIECE','Y_NUMLIGNE'],
                [TobLigEqu.GetValue('E_DATECOMPTABLE'),TobLigEqu.GetValue('E_JOURNAL'),TobLigEqu.GetValue('E_NUMEROPIECE'),TobLigEqu.GetValue('E_NUMLIGNE')],True) ;
      if TobLigEcr <> nil then
      begin
        if TotCredit > TotDebit then
           if TobLigEcr.GetValue('E_CREDIT') > 0 then
           begin
              TobLigEcr.PutValue('E_CREDIT', TobLigEcr.GetValue('E_CREDIT') - Diff) ;
              if TobLigAna<>nil then TobLigAna.PutValue('Y_CREDIT', TobLigAna.GetValue('Y_CREDIT') - Diff) ;
           end
           else
           begin
              TobLigEcr.PutValue('E_DEBIT', TobLigEcr.GetValue('E_DEBIT') + Diff) ;
              if TobLigAna<>nil then TobLigAna.PutValue('Y_DEBIT', TobLigAna.GetValue('Y_DEBIT') + Diff) ;
           end 
        else
          if TobLigEcr.GetValue('E_CREDIT') > 0 then
          begin
             TobLigEcr.PutValue('E_CREDIT', TobLigEcr.GetValue('E_CREDIT') + Diff) ;
             if TobLigAna<>nil then TobLigAna.PutValue('Y_CREDIT', TobLigAna.GetValue('Y_CREDIT') + Diff) ;
          end
          else
          begin
             TobLigEcr.PutValue('E_DEBIT', TobLigEcr.GetValue('E_DEBIT') - Diff) ;
             if TobLigAna<>nil then TobLigAna.PutValue('Y_DEBIT', TobLigAna.GetValue('Y_DEBIT') - Diff) ;
          end ;
      end ;
    end;
    TobTstEquilibre.ClearDetail;
  end;

begin
  //Calcul n° définitifs des pièces et des lignes
  TobTstEquilibre := TOB.Create('Test equilibre', nil, -1);
  TobEcart := TOB.Create('Test ecart',nil,-1);
  //ExportTiers := GetParamSoc('SO_ACTIVECOMSX');
  ExportTiers := GetParamSoc('SO_COMPTAEXTERNE');
  PathFile := GetParamSoc('SO_MBOCHEMINCOMPTA');
  if Copy(PathFile, Length(PathFile), 1) <> '\' then
    PathFile := PathFile + '\Tiers.txt'
  else
    PathFile := PathFile + 'Tiers.txt';
  TobEcriture.Detail[0].Detail.sort('E_DATECOMPTABLE;E_JOURNAL;E_NUMEROPIECE;E_NUMLIGNE');
  QteEcr := 0;
  NumEcriture := 0;
  NewNumPiece := 0;
  NewNumLigne := 0;
  Jnal := '';
  for Cpt := 0 to TobEcriture.Detail[0].Detail.count - 1 do
  begin
    TobTmp := TobEcriture.Detail[0].Detail[Cpt];
    CpteGene := TobTmp.GetValue('E_GENERAL');
    CpteAux := TobTmp.GetValue('E_AUXILIAIRE');
    NewEcriture := False;
    // MODIF AC PB NUMECHE
    TobTmp.PutValue('E_NUMECHE',0) ;
    {Préparation de la Tob d'export des tiers et SO_EXPJRNX}
    if (ExportTiers) and (CpteAux <> '') then
    begin
      if pos(TobTmp.GetValue('E_JOURNAL'), ExpJrnx) = 0 then
        ExpJrnx := ExpJrnx + ';' + TobTmp.GetValue('E_JOURNAL');
      TobTmp2 := TobExportTiers.FindFirst(['T_AUXILIAIRE'], [CpteAux], True);
      if TobTmp2 = nil then
      begin
        TobTmp2 := TOB.Create('TIERS', TobExportTiers, -1);
        TobTmp2.SelectDb('"' + CpteAux + '"', nil);
      end;
    end;
    if (NumEcriture <> TobTmp.GetValue('E_NUMEROPIECE')) or (Jnal <> TobTmp.GetValue('E_JOURNAL')) then
    begin
      //Nvelle pièce, test équilibre
      if TobTstEquilibre.Detail.count > 0 then
        TestEquilibre;
      Jnal := TobTmp.GetValue('E_JOURNAL');
      NumEcriture := TobTmp.GetValue('E_NUMEROPIECE');
      NewNumPiece := GetNewNumJal(TobTmp.GetValue('E_JOURNAL'), True, TobTmp.GetValue('E_DATECOMPTABLE'));
      NewNumLigne := 1;
      QteEcr := QteEcr + 1;
      NewEcriture := True;
      TobEcart.ClearDetail ;
      //Gestion écart
      if Tobtmp.GetValue('TYPEMVT') = 'VEN' then
      begin
        TobTmp3 := TOB.Create('',TobEcart,-1);
        TobTmp3.AddChampSupValeur('CLIENT1','',false);
        TobTmp3.AddChampSupValeur('SENS',TobTmp.GetValue('SENS'),false);
        if TobTmp.GetValue('SENS') = 'D' then
          TobTmp3.AddChampSupValeur('MONTANT',TobTmp.GetValue('E_DEBIT'),false)
          else
          TobTmp3.AddChampSupValeur('MONTANT',TobTmp.GetValue('E_CREDIT'),false);
        TobTmp3.AddChampSupValeur('NUMEROLIGNE',1,false);
      end;
    end else
    begin
      if TobTmp.GetValue('TYPEMVT') = 'BTQ' then
        NewNumLigne := NewNumLigne + 1
      else if TobTmp.GetValue('E_ECHE') <> 'X' then
        NewNumLigne := NewNumLigne + 1;
      if TobTmp.GetValue('TYPEMVT') = 'VEN' then
      begin
        if pos('TIERS',TobTmp.GetValue('TYPELIGNE')) > 0 then
        begin
          //if TobTmp.GetValue('SENS') <> (TobEcart.Detail[0].GetValue('SENS')) then
          if ((TobTmp.GetValue('MONTANT1')<0) and (TobEcart.Detail[0].GetValue('MONTANT')>0))
          or ((TobTmp.GetValue('MONTANT1')>0) and (TobEcart.Detail[0].GetValue('MONTANT')<0)) 
          or (Int(Abs(TobTmp.GetValue('MONTANT1'))*10) = 0) or (Int(Abs(TobTmp.GetValue('MONTANT1'))) = 0) then
          begin
            TobTmp.AddChampSupValeur('ASUPPRIMER','X',false);
            TobTmp3 := TOB.Create('',TobEcart,-1);
            TobTmp3.AddChampSupValeur('CLIENT2','',false);
            TobTmp3.AddChampSupValeur('SENS',TobTmp.GetValue('SENS'),false);
            //if TobTmp.GetValue('SENS') = 'D' then
              TobTmp3.AddChampSupValeur('MONTANT',TobTmp.GetValue('MONTANT1'),false) ;
             // else
              //TobTmp3.AddChampSupValeur('MONTANT',TobTmp.GetValue('MONTANT1'),false);
            TobTmp3.AddChampSupValeur('NUMEROLIGNE',NewNumLigne+1,false);
          end;
        end else
        if (pos('HT_VT',TobTmp.GetValue('TYPELIGNE')) > 0) and (TobEcart.Detail.count = 2) then
        begin
          {NewNumLigne := TobEcart.Detail[1].GetValue('NUMEROLIGNE');
          if TobTmp.GetValue('SENS') = 'D' then
          begin
            if TobEcart.Detail[1].GetValue('MONTANT') < 0 then
               TobTmp.PutValue('E_DEBIT',TobTmp.GetValue('E_DEBIT')+ Abs(TobEcart.Detail[1].GetValue('MONTANT')))
            else
               TobTmp.PutValue('E_DEBIT',TobTmp.GetValue('E_DEBIT')- Abs(TobEcart.Detail[1].GetValue('MONTANT'))) ;
          end else
          begin
            if TobEcart.Detail[1].GetValue('MONTANT') < 0 then
              TobTmp.PutValue('E_CREDIT',TobTmp.GetValue('E_CREDIT')+ Abs(TobEcart.Detail[1].GetValue('MONTANT')))
            else TobTmp.PutValue('E_CREDIT',TobTmp.GetValue('E_CREDIT')- Abs(TobEcart.Detail[1].GetValue('MONTANT'))) ;
          end ; }
          TobEcart.ClearDetail;
        end;
      end;
    end;
    TobTmp.PutValue('E_NUMEROPIECE', NewNumPiece);
    TobTmp.PutValue('E_NUMLIGNE', NewNumLigne);
    //    TobTmp.PutValue('E_AFFAIRE', TobTmp.GetValue('E_JOURNAL') + ';' + TobTmp.GetValue('E_EXERCICE') + ';' + DateToStr(TobTmp.GetValue('E_DATECOMPTABLE')) + ';' + IntToStr(TobTmp.GetValue('E_NUMEROPIECE')) + ';' + IntToStr(TobTmp.GetValue('E_NUMLIGNE')) + ';' + IntToStr(TobTmp.GetValue('E_NUMECHE')) + ';' + TobTmp.GetValue('E_QUALIFPIECE'));
    if TobTmp.GetValue('E_ANA') = 'X' then
    begin
      TobTmp1 := TobEcriture.Detail[1].FindFirst(['Y_NUMEROPIECE', 'Y_GENERAL'], [NumEcriture, CpteGene], True);
      while TobTmp1 <> nil do
      begin
        TobTmp1.PutValue('Y_NUMEROPIECE', NewNumPiece);
        TobTmp1.PutValue('Y_NUMLIGNE', NewNumLigne);
        //        TobTmp1.PutValue('Y_AFFAIRE', TobTmp1.GetValue('Y_JOURNAL') + ';' + TobTmp1.GetValue('Y_EXERCICE') + ';' + DateToStr(TobTmp1.GetValue('Y_DATECOMPTABLE')) + ';' + IntToStr(TobTmp1.GetValue('Y_NUMEROPIECE')) + ';' + IntToStr(TobTmp1.GetValue('Y_NUMLIGNE')) + TobTmp1.GetValue('Y_AXE') + ';' + IntToStr(TobTmp1.GetValue('Y_NUMVENTIL')) + ';' + TobTmp1.GetValue('Y_QUALIFPIECE'));
        TobTmp1 := TobEcriture.Detail[1].FindNext(['Y_NUMEROPIECE', 'Y_GENERAL'], [NumEcriture, CpteGene], True);
      end;
    end;
    TobTmp1 := TOB.Create('', TobTstEquilibre, -1);
    TobTmp1.Dupliquer(TobTmp, true, true);
  end;
  // Pour tester la dernière piece
  if TobTstEquilibre.Detail.count > 0 then
        TestEquilibre;
  if TobTstEquilibre <> nil then
    FreeAndNil(TobTstEquilibre);
  if TobEcart <> nil then
    FreeAndNil(TobEcart);
  //****************************
  //  gcvoirtob(TobEcriture,'TobEcriture');
  //****************************
  for i:=TobEcriture.Detail[0].Detail.Count - 1 DownTo 0 do
  begin
     if TobEcriture.Detail[0].Detail[i].FieldExists('ASUPPRIMER') then TobEcriture.Detail[0].Detail[i].Free ;
  end ;
  if TobEcriture.InsertDB(nil) then
  begin
    if TobExportTiers.Detail.count > 0 then
    begin
      ExpJrnx := copy(ExpJrnx, 2, length(ExpJrnx));
      SetParamSoc('SO_EXPJRNX', ExpJrnx);
      TobExportTiers.SaveToFile(PathFile, False, False, True);
    end;
  end;
end;

procedure MajRapport;
var Cpt, NumPiece: integer;
  TobL, TobAna, TobTmp: TOB;
  Texte, Valeur, Montant, Sens: string;

  procedure CreatTexte(Ets, Jna, Dte, Nat, Gen, Typ, Aux, Mpa, Dev, Mt, DC: string);
  begin
    Texte := Format('%-4s', [Ets]) +
      Format('%-4s', [Jna]) +
      Dte + ' ' +
      Format('%-4s', [Nat]) +
      Format('%-15s', [Gen]) +
      Format('%-2s', [Typ]) +
      Format('%-15s', [Aux]) +
      Format('%-4s', [Mpa]) +
      Format('%-4s', [Dev]) +
      Format('%10s', [Mt]) + ' ' + DC;
  end;

begin
  NumPiece := 0;
  RapportLst_.Items.Add('');
  RapportLst_.Items.Add('***** ' + TraduireMemoire('PIECES GENEREES ') + '*****************************************************');
  RapportLst_.Items.Add(IntToStr(QteEcr) + TraduireMemoire(' pièces générée(s) :'));
  RapportLst_.Items.Add('- ' + IntToStr(TobEcriture.Detail[0].Detail.count) + TraduireMemoire(' lignes écriture'));
  RapportLst_.Items.Add('- ' + IntToStr(TobEcriture.Detail[1].Detail.count) + TraduireMemoire(' lignes analytique'));
  TobAna := TOB.Create('Donnees ana', nil, -1);
  TobAna.Dupliquer(TobEcriture, True, True, True);
  TobEcriture.Detail[0].Detail.Sort('E_ETABLISSEMENT;E_DATECOMPTABLE;E_NATUREPIECE;E_JOURNAL;E_NUMEROPIECE;E_NUMLIGNE;E_MODEPAIE;E_GENERAL');
  for Cpt := 0 to TobEcriture.Detail[0].Detail.count - 1 do
  begin
    TobL := TobEcriture.Detail[0].Detail[Cpt];
    if TobL.NomTable = 'ECRITURE' then
    begin
      if NumPiece <> TobL.GetValue('E_NUMEROPIECE') then
      begin
        RapportLst_.Items.Add('---------------------------------------------------------------------------');
        NumPiece := TobL.GetValue('E_NUMEROPIECE');
        Texte := TraduireMemoire('Pièce numéro ') + IntToStr(NumPiece);
        RapportLst_.Items.Add(Texte);
      end;
      if TobL.GetValue('E_AUXILIAIRE') <> '' then
        Valeur := 'X'
      else
        Valeur := '';
      if TobL.getValue('E_CREDIT') <> 0 then
      begin
        Montant := FloatToStrF(TobL.GetValue('E_CREDIT'), ffNumber, 8, 2);
        Sens := 'C';
      end else
      begin
        Montant := FloatToStrF(TobL.GetValue('E_DEBIT'), ffNumber, 8, 2);
        Sens := 'D';
      end;
      CreatTexte(TobL.GetValue('E_ETABLISSEMENT'), TobL.GetValue('E_JOURNAL'), DateToStr(TobL.GetValue('E_DATECOMPTABLE')),
        TobL.GetValue('E_NATUREPIECE'), TobL.GetValue('E_GENERAL'),
        Valeur, TobL.GetValue('E_AUXILIAIRE'), TobL.GetValue('E_MODEPAIE'),
        TobL.GetValue('E_DEVISE'), Montant, Sens);
      RapportLst_.Items.Add(Texte);
      //Ligne avec analytique, recherche la ou les lignes correspondantes
      if TobL.GetValue('E_ANA') = 'X' then
      begin
        TobTmp := TobAna.FindFirst(['Y_NUMEROPIECE', 'Y_GENERAL', 'Y_DATECOMPTABLE', 'Y_ETABLISSEMENT'],
          [TobL.GetValue('E_NUMEROPIECE'), TobL.GetValue('E_GENERAL'),
          TobL.GetValue('E_DATECOMPTABLE'), TobL.GetValue('E_ETABLISSEMENT')],
            True);
        while TobTmp <> nil do
        begin
          if TobTmp.getValue('Y_CREDIT') <> 0 then
          begin
            Montant := FloatToStrF(TobTmp.GetValue('Y_CREDIT'), ffNumber, 8, 2);
            Sens := 'C';
          end else
          begin
            Montant := FloatToStrF(TobTmp.GetValue('Y_DEBIT'), ffNumber, 8, 2);
            Sens := 'D';
          end;
          CreatTexte(TobTmp.GetValue('Y_ETABLISSEMENT'), TobTmp.GetValue('Y_JOURNAL'), DateToStr(TobTmp.GetValue('Y_DATECOMPTABLE')),
            TobTmp.GetValue('Y_NATUREPIECE'), TobTmp.GetValue('Y_GENERAL'), 'A',
            TobTmp.GetValue('Y_SECTION'), TobL.GetValue('E_MODEPAIE'), TobL.GetValue('E_DEVISE'),
            Montant, Sens);
          RapportLst_.Items.Add(Texte);
          TobTmp := TobAna.FindNext(['Y_NUMEROPIECE', 'Y_GENERAL', 'Y_DATECOMPTABLE', 'Y_ETABLISSEMENT'],
            [TobL.GetValue('E_NUMEROPIECE'), TobL.GetValue('E_GENERAL'),
            TobL.GetValue('E_DATECOMPTABLE'), TobL.GetValue('E_ETABLISSEMENT')],
              True);
        end;
      end;
    end;
  end;
  TobAna.Free;
end;

procedure RecalculNumLignes(TobATraiter: TOB);
var Cpt, NumPiece, NumLigne: integer;
  TobL: TOB;
begin
  TobATraiter.detail.sort('DATECOMPTABLE;TYPEPIECE;JOURNAL;NUMEROPIECE;NUMEROLIGNE');
  NumPiece := 0;
  NumLigne := 0;
  for Cpt := 0 to TOBFinale.Detail.count - 1 do
  begin
    TobL := TobATraiter.Detail[Cpt];
    if NumPiece <> TobL.GetValue('NUMEROPIECE') then
    begin
      NumPiece := TobL.GetValue('NUMEROPIECE');
      NumLigne := 1;
    end else
    begin
      if TobL.GetValue('NUMECHE') <= 1 then
        NumLigne := NumLigne + 1;
    end;
    TobL.PutValue('NUMEROLIGNE', NumLigne);
  end;
end;
//*****************************************************************
// FIN ALIMENTATION TABLES ECRITURES ET ANALYTIQ
//*****************************************************************

function GoodDate(StDate: string): string;
begin
  Result := StringReplace(StDate, '/', '', [rfReplaceAll]);
end;

procedure EcrireUneLigne(TOBL: TOB);
var Ligne: string;
begin
  {Format CEGID + CEGID étendu pour les Ecritures de ventes et de Règlements}
  Ligne := '';
  with TOBL do
  begin
    Ligne := ASCIIFormatChaine(GetValue('JOURNAL'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATECOMPTABLE')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TYPEPIECE'), '', '', ' ', ' ', 2);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('GENERAL'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TYPECPTE'), '', '', ' ', ' ', 1);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('AUXILIAIRE'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('REFINTERNE'), '', '', ' ', ' ', 35);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBELLE'), '', '', ' ', ' ', 35);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('MODEPAIE'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('ECHEANCE')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('SENS'), '', '', ' ', ' ', 1);
    //Ligne:=Ligne+ASCIIFormatChaine(GetValue('MONTANT1'),'','',' ',' ',20,True);
    Ligne := Ligne + ASCIIFormatChaine(StrFmontant(GetValue('MONTANT1'), 10, 2, '', False), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TYPEECRITURE'), '', '', ' ', ' ', 1);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('NUMEROPIECE'), '', '', ' ', ' ', 8, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('DEVISE'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TAUXDEV'), '', '', ' ', ' ', 10, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('CODEMONTANT'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('MONTANT2'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('MONTANT3'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('ETABLISSEMENT'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('AXE'), '', '', ' ', ' ', 2, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('NUMECHE'), '', '', ' ', ' ', 2, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('REFEXTERNE'), '', '', ' ', ' ', 35);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATEREFEXTERNE')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATECREATION')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('SOCIETE'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('AFFAIRE'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATETAUXDEV')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('ECRANOUVEAU'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('QTE1'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('QTE2'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('QUALIFQTE1'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('QUALIFQTE2'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('REFLIBRE'), '', '', ' ', ' ', 35);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TVAENCAISSEMENT'), '', '', ' ', ' ', 1);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('REGIMETVA'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TVA'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TPF'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('CONTREPARTIEGEN'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('CONTREPARTIEAUX'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('REFPOINTAGE'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATEPOINTAGE')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATERELANCE')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATEVALEUR')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('RIB'), '', '', ' ', ' ', 35);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('REFRELEVE'), '', '', ' ', ' ', 10);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('NUMEROIMMO'), '', '', ' ', ' ', 17);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE0'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE1'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE2'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE3'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE4'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE5'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE6'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE7'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE8'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBRETEXTE9'), '', '', ' ', ' ', 30);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TABLE0'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TABLE1'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TABLE2'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('TABLE3'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBREMONTANT0'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBREMONTANT1'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBREMONTANT2'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBREMONTANT3'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('LIBREDATE')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBREBOOL0'), '', '', ' ', ' ', 1);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LIBREBOOL1'), '', '', ' ', ' ', 1);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('CONSO'), '', '', ' ', ' ', 3);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('COUVERTURE'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('COUVERTUREDEV'), '', '', ' ', ' ', 20, True);
    //Ligne := Ligne + ASCIIFormatChaine(GetValue('COUVERTUREEURO'), '', '', ' ', ' ', 20, True);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATEPAQUETMAX')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GoodDate(GetValue('DATEPAQUETMIN')), '', '', ' ', ' ', 8);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LETTRAGE'), '', '', ' ', ' ', 5);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('LETTRAGEDEV'), '', '', ' ', ' ', 1);
    //Ligne := Ligne + ASCIIFormatChaine(GetValue('LETTRAGEEURO'), '', '', ' ', ' ', 1);
    Ligne := Ligne + ASCIIFormatChaine(GetValue('ETATLETTRAGE'), '', '', ' ', ' ', 3);
  end;
end;

procedure EcrireLigneTiers(Q: TQuery);
var St: string;
begin
  {Format CEGID + CEGID étendu pour les Comptes Tiers}
  St := '***CAE'; //FIXE + IDENTIFIANT
  St := St + ASCIIFormatChaine(Q.FindField('T_AUXILIAIRE').AsString, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine(Q.FindField('T_LIBELLE').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('T_NATUREAUXI').AsString, '', '', ' ', ' ', 3);
  St := St + 'X'; //LETTRABLE
  St := St + ASCIIFormatChaine(Q.FindField('T_COLLECTIF').AsString, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine(Q.FindField('T_EAN').AsString, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 1
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 2
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 3
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 4
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 5
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 6
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 7
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 8
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 9
  St := St + ASCIIFormatChaine('', '', '', ' ', ' ', 17); //TABLE 10
  St := St + ASCIIFormatChaine(Q.FindField('T_ADRESSE1').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('T_ADRESSE2').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('T_ADRESSE3').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('T_CODEPOSTAL').AsString, '', '', ' ', ' ', 9);
  St := St + ASCIIFormatChaine(Q.FindField('T_VILLE').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('R_DOMICILIATION').AsString, '', '', ' ', ' ', 24);
  St := St + ASCIIFormatChaine(Q.FindField('R_ETABBQ').AsString, '', '', ' ', ' ', 5);
  St := St + ASCIIFormatChaine(Q.FindField('R_GUICHET').AsString, '', '', ' ', ' ', 5);
  St := St + ASCIIFormatChaine(Q.FindField('R_NUMEROCOMPTE').AsString, '', '', ' ', ' ', 11);
  St := St + ASCIIFormatChaine(Q.FindField('R_CLERIB').AsString, '', '', ' ', ' ', 2);
  St := St + ASCIIFormatChaine(Q.FindField('T_PAYS').AsString, '', '', ' ', ' ', 3);
  St := St + ASCIIFormatChaine(Q.FindField('T_ABREGE').AsString, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine(Q.FindField('T_LANGUE').AsString, '', '', ' ', ' ', 3);
  St := St + Q.FindField('T_MULTIDEVISE').AsString;
  St := St + ASCIIFormatChaine(Q.FindField('T_DEVISE').AsString, '', '', ' ', ' ', 3);
  St := St + ASCIIFormatChaine(Q.FindField('T_TELEPHONE').AsString, '', '', ' ', ' ', 25);
  St := St + ASCIIFormatChaine(Q.FindField('T_FAX').AsString, '', '', ' ', ' ', 25);
  St := St + ASCIIFormatChaine(Q.FindField('T_REGIMETVA').AsString, '', '', ' ', ' ', 3);
  St := St + ASCIIFormatChaine(Q.FindField('T_MODEREGLE').AsString, '', '', ' ', ' ', 3);
  St := St + ASCIIFormatChaine(Q.FindField('T_COMMENTAIRE').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('T_NIF').AsString, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine(Q.FindField('T_SIRET').AsString, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine(Q.FindField('T_APE').AsString, '', '', ' ', ' ', 5);
  St := St + ASCIIFormatChaine(Q.FindField('C_NOM').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('C_SERVICE').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('C_FONCTION').AsString, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Q.FindField('C_TELEPHONE').AsString, '', '', ' ', ' ', 25);
  St := St + ASCIIFormatChaine(Q.FindField('C_FAX').AsString, '', '', ' ', ' ', 25);
  St := St + ASCIIFormatChaine(Q.FindField('C_TELEX').AsString, '', '', ' ', ' ', 25);
  St := St + ASCIIFormatChaine(Q.FindField('C_RVA').AsString, '', '', ' ', ' ', 50);
  St := St + ASCIIFormatChaine(Q.FindField('C_CIVILITE').AsString, '', '', ' ', ' ', 3);
  St := St + Q.FindField('C_PRINCIPAL').AsString;
  St := St + ASCIIFormatChaine(Q.FindField('T_JURIDIQUE').AsString, '', '', ' ', ' ', 3);
  St := St + Q.FindField('R_PRINCIPAL').AsString;
end;

procedure EcrireLigneSectionAna(Code, Libelle, Axe: string);
var St: string;
begin
  {Format CEGID + CEGID étendu pour les Comptes Tiers}
  St := '***SAT'; //FIXE + IDENTIFIANT
  St := St + ASCIIFormatChaine(Code, '', '', ' ', ' ', 17);
  St := St + ASCIIFormatChaine(Libelle, '', '', ' ', ' ', 35);
  St := St + ASCIIFormatChaine(Axe, '', '', ' ', ' ', 3);
end;

// Charge les axes anatytique de charge en fonction du ParamSoc SO_GCAXEANALYTIQUE=FALSE

procedure ChargeAxeAnalytique(Srang, Nature: string; QLC: TQuery);
var SQLAnal, VenteAchat, NatV, Ax, Section: string;
  Q, QCompta: TQuery;
  TOBA, TOBCompta, TOBC: Tob;
  Pourc, MontantHT, MontantHTDEV, MontantTTC, MontantTTCDEV,
    EcartHTDEV, EcartHT, EcartTTCDEV, EcartTTC: Double;
  i, j, Compteur: Integer;
begin
  j := 0;
  VenteAchat := GetInfoParPiece(Nature, 'GPP_VENTEACHAT');
  if VenteAchat = 'VEN' then NatV := 'HV' else NatV := 'HA';
  QCompta := OpenSQL('SELECT MAX(GIC_COMPTEUR) FROM INTERCOMPTA', True);
  Compteur := QCompta.Fields[0].AsInteger;
  Ferme(QCompta);
  // Charge les axes dans la tob
  SQLAnal := 'SELECT * FROM VENTIL WHERE (V_NATURE like "' + NatV + '%" OR V_NATURE LIKE "ST%") AND V_COMPTE="' + sRang + '"';
  Q := OpenSQL(SQLAnal, True);
  if not Q.EOF then
  begin
    TOBAxe.LoadDetailDB('VENTIL', '', '', Q, False, True);
    TOBAxe.Detail.Sort('V_NUMEROVENTIL');
    TOBCompta := TOB.Create('_INTERCOMPTA', nil, -1);
  end else
  begin
    Ferme(Q);
    Exit;
  end;
  Ferme(Q);
  // Charge les lignes analytique de charge dans la table intercompta
  for i := 0 to TOBAxe.Detail.Count - 1 do
  begin
    Inc(Compteur);
    TOBA := TOBAxe.Detail[i];
    Ax := 'A' + Copy(TOBA.GetValue('V_NATURE'), 3, 1);
    Pourc := TOBA.GetValue('V_TAUXMONTANT');
    Section := TOBA.GetValue('V_SECTION');
    MontantHTDEV := Arrondi(Pourc * QLC.FindField('GIC_MONTANTHTDEV').AsFloat / 100, 2);
    MontantHT := Arrondi(Pourc * QLC.FindField('GIC_MONTANTHT').AsFloat / 100, 2);
    MontantTTCDEV := Arrondi(Pourc * QLC.FindField('GIC_MONTANTTTCDEV').AsFloat / 100, 2);
    MontantTTC := Arrondi(Pourc * QLC.FindField('GIC_MONTANTTTC').AsFloat / 100, 2);
    TOBC := TOB.Create('INTERCOMPTA', TOBCompta, j);
    TOBC.PutValue('GIC_USER', QLC.FindField('GIC_USER').AsString);
    TOBC.PutValue('GIC_DATE', QLC.FindField('GIC_DATE').AsDateTime);
    TOBC.PutValue('GIC_TIERS', QLC.FindField('GIC_TIERS').AsString);
    TOBC.PutValue('GIC_COMPTATIERS', QLC.FindField('GIC_COMPTATIERS').AsString);
    TOBC.PutValue('GIC_COMPTAARTICLE', QLC.FindField('GIC_COMPTAARTICLE').AsString);
    TOBC.PutValue('GIC_REGIMETAXE', QLC.FindField('GIC_REGIMETAXE').AsString);
    TOBC.PutValue('GIC_FAMILLETAXE1', QLC.FindField('GIC_FAMILLETAXE1').AsString);
    TOBC.PutValue('GIC_FAMILLETAXE2', QLC.FindField('GIC_FAMILLETAXE2').AsString);
    TOBC.PutValue('GIC_FAMILLETAXE3', QLC.FindField('GIC_FAMILLETAXE3').AsString);
    TOBC.PutValue('GIC_FAMILLETAXE4', QLC.FindField('GIC_FAMILLETAXE4').AsString);
    TOBC.PutValue('GIC_FAMILLETAXE5', QLC.FindField('GIC_FAMILLETAXE5').AsString);
    TOBC.PutValue('GIC_ETABLISSEMENT', QLC.FindField('GIC_ETABLISSEMENT').AsString);
    TOBC.PutValue('GIC_NATUREPIECEG', QLC.FindField('GIC_NATUREPIECEG').AsString);
    TOBC.PutValue('GIC_SOUCHE', QLC.FindField('GIC_SOUCHE').AsString);
    TOBC.PutValue('GIC_NUMERO', QLC.FindField('GIC_NUMERO').AsInteger);
    TOBC.PutValue('GIC_INDICE', QLC.FindField('GIC_INDICE').AsInteger);
    TOBC.PutValue('GIC_NUMLIGNE', QLC.FindField('GIC_NUMLIGNE').AsInteger);
    TOBC.PutValue('GIC_VENTEACHAT', QLC.FindField('GIC_VENTEACHAT').AsString);
    TOBC.PutValue('GIC_TVAENCAISSEMENT', QLC.FindField('GIC_TVAENCAISSEMENT').AsString);
    TOBC.PutValue('GIC_DEVISE', QLC.FindField('GIC_DEVISE').AsString);
    TOBC.PutValue('GIC_QTE1', QLC.FindField('GIC_QTE1').AsInteger);
    TOBC.PutValue('GIC_TAUXDEVISE', QLC.FindField('GIC_TAUXDEVISE').AsFloat);
    TOBC.PutValue('GIC_MONTANTHTDEV', MontantHTDEV);
    TOBC.PutValue('GIC_MONTANTHT', MontantHT);
    TOBC.PutValue('GIC_MONTANTTTCDEV', MontantTTCDEV);
    TOBC.PutValue('GIC_MONTANTTTC', MontantTTC);
    TOBC.PutValue('GIC_SECTAN' + Ax, Section);
    Q := OpenSQL('SELECT S_LIBELLE FROM SECTION WHERE S_SECTION="' + Section + '"', True);
    TOBC.PutValue('GIC_LIBAN' + Ax, Q.Fields[0].AsString);
    Ferme(Q);
    TOBC.PutValue('GIC_COMPTEUR', Compteur);
    inc(j);
  end;
  // Equilibrer les lignes des axes analytiques
  MontantHTDEV := 0;
  MontantHT := 0;
  MontantTTCDEV := 0;
  MontantTTC := 0;
  for i := 0 to TOBCompta.Detail.Count - 1 do
  begin
    MontantHTDEV := MontantHTDEV + TOBCompta.Detail[i].GetValue('GIC_MONTANTHTDEV');
    MontantHT := MontantHT + TOBCompta.Detail[i].GetValue('GIC_MONTANTHT');
    MontantTTCDEV := MontantTTCDEV + TOBCompta.Detail[i].GetValue('GIC_MONTANTTTCDEV');
    MontantTTC := MontantTTC + TOBCompta.Detail[i].GetValue('GIC_MONTANTTTC');
  end;
  EcartHTDEV := QLC.FindField('GIC_MONTANTHTDEV').AsFloat - MontantHTDev;
  EcartHT := QLC.FindField('GIC_MONTANTHT').AsFloat - MontantHT;
  EcartTTCDEV := QLC.FindField('GIC_MONTANTTTCDEV').AsFloat - MontantTTCDev;
  EcartTTC := QLC.FindField('GIC_MONTANTTTC').AsFloat - MontantTTC;
  if EcartHTDev <> 0 then TOBCompta.Detail[0].PutValue('GIC_MONTANTHTDEV', TOBCompta.Detail[0].GetValue('GIC_MONTANTHTDEV') + EcartHTDev);
  if EcartHT <> 0 then TOBCompta.Detail[0].PutValue('GIC_MONTANTHT', TOBCompta.Detail[0].GetValue('GIC_MONTANTHT') + EcartHT);
  if EcartTTCDev <> 0 then TOBCompta.Detail[0].PutValue('GIC_MONTANTTTCDEV', TOBCompta.Detail[0].GetValue('GIC_MONTANTTTCDEV') + EcartTTCDev);
  if EcartTTC <> 0 then TOBCompta.Detail[0].PutValue('GIC_MONTANTTTC', TOBCompta.Detail[0].GetValue('GIC_MONTANTTTC') + EcartTTC);
  TOBCompta.InsertDB(nil);
  TOBCompta.Free;
end;

end.
