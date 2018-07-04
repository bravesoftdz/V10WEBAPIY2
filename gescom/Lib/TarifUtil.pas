unit TarifUtil;

interface

uses HEnt1, HCtrls, UTOB, Ent1, EntGC, LookUp, Controls, ComCtrls, StdCtrls,
  ExtCtrls, windows, ParamSoc,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB, Fe_Main,
  {$ENDIF}
  SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, Forms,
  Paramdat, AGLInit, TiersUtil, UtilArticle;
const NbRowsInit = 50;
  NbRowsPlus = 20;

type T_TableTarif = (ttdArt, ttdArtQte, ttdArtQCa, ttdArtCat,
    ttdCliQte, ttdCliArt, ttdCliFam,
    ttdTiersFam, ttdTiersArt, ttdTiersQArt,
    ttdCatA, ttdCatC);
  T_SaisieTarif = (tstRapide, tstNormale, tstCliArt, tstTiers, tstArticle,
    tstAutoFour);
  T_RechDepot = (trdAucun, trdErreur, trdOk);

  // Opérations sur les champs ou les zones
procedure CalcPriorite(TOBL: TOB);
function WhereTarifArt(CodeArticle, CodeDevise: string; ttd: T_TableTarif; Totale: boolean): string;
function WhereTarifCli(CodeTiers, CodeArticle, CodeDevise: string; ttd: T_TableTarif; Totale: boolean): string;
{$IFDEF MONCODE}
function NbreDecimalSuppPrix(NatTiers: string; TOBArt: TOB): Integer;
{$ENDIF}
// Articles
function ChoisirDimension(St_CodeArt: string; var TOBArt: tob): Boolean;
//  Tiers
function TrouverTiersSQL(G_CodeTiers: THCritMaskEdit; TOBTiers: TOB): T_RechArt;
function TrouverFouSQL(G_CodeTiers: THCritMaskEdit; TOBTiers: TOB): T_RechArt;
// Dépôt
procedure GetDepotRecherche(DEPOT: THCritMaskEdit);
function ExisteDepot(Tablette, Valeur: string): boolean;
// Catégorie tiers
procedure GetCategorieRecherche(CATEGORIE: THCritMaskEdit);
function ExisteCategorie(Tablette, Valeur: string): boolean;
// Famille Tarif Article
procedure GetFamilleRecherche(FAMILLE: THCritMaskEdit);
function ExisteFamille(Tablette, Valeur: string): boolean;
function ExisteSousFamille (Famille,SousFamille : string) : boolean;
// Date
procedure GetDateRecherche(F: TFORM; DATE: THCritMaskEdit);
// Calcul Remise
function ModifFormat(St: string): string;
function RemiseResultante(St: string): extended;
function Numeric(const St: string): Extended;
function ReadTokenPlus(var St: string): Extended;
// Tablette
function ExisteTablette(Tablette, Valeur: string): boolean;
// Convertion contrevaleur
function ConvertPrix(Prix: Double): Double;
// calcul arrondi
function ArrondirPrix(CodeArr: string; Prix: double; RecupTobArr: Boolean = False): Double;
function CalculArrondi(TOBA: Tob; Prix: Double): Double;
// Specifique mode
procedure FreeTobArr;
function PrixContreTarif(Prix: Double; Devise, TypeTarif, NatureType: string; TobDev: Tob): Double;
function TraiterTableTarifMode(CodeType, CodePeriode, CodeNature: string; TobTarifMode: TOB): TOB;
function RecupInfoPeriode(CodePeriode, Etablissement: string): string;
function RecupCoefTypeTarif(TypeTarif, NatureType: string): Double;
function SQLTarifGen(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot: string; LaDate: TDateTime; Qte: double; AvecDevise: Boolean = True):
  string;
function SQLTarifDim(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot: string; LaDate: TDateTime; Qte: double; AvecDevise: Boolean = True):
  string;
function ControlTypeTarifEtab(TobMode: Tob; TypeTarif, TarfMode: string): boolean;
// Prix
function RechTarifSpec(TOBTarif, TobMode: Tob; CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): string;
function RechPrixTarifBase(TOBTarif, TobMode: Tob; CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): Double;
function CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat: string; LaDate: TDateTime; EnHT: Boolean;
  AvecDevise: Boolean = True): TOB;
function CreerTobTarifArtDim(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat: string; LaDate: TDateTime; EnHT: Boolean;
  AvecDevise: Boolean = True): TOB;
procedure FindTarifArt(TOBArt: TOB; Etablissement: string; UnArticle: Boolean = False);
function RecupCodeTarifMode(TypeTarif, Periode, NatureType: string): Integer;
function CalculPrixVente(TOBTarif: TOB; CodeArticle, TarifArticle, CodeDepot, CodeDevise, TypeTarfEtab: string; PxVenteArt: Double; AvecInfo: Boolean = False):
  Double;
// Remises
function ChercheMieRem(TOBTarif, TobMode: Tob; CodeArt, TarfArt, CodeTiers, TarfTiers, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string): string;
function CalcRemArt(TOBTarif, TOBMode: TOB; CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): string;
function CalcRemCatArt(TOBTarif, TOBMode: TOB; TarfArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): string;
function CalcRemCli(TOBTarif, TOBMode: TOB; TarfArt, Codetiers, TarfTiers, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string): Double;
// Gestion des tarifs MODE
function ToxConvToDev(Montant: Double; DevOrg, DevDest: string; DatePiece: TDateTime; TOBDev: TOB): Double;
procedure ToxConvertir(MontantOrigine: double; var MontantDev, MontantContreValeur: double; DevOrg, DevDest: string; Date: TDateTime; TobDev: TOB);
procedure ToxConvToDev2(MontantOrg: Double; var MontantDestNA, MontantDest: Double; DevOrg, DevDest: string; DatePiece: TDateTime; TOBDev: TOB);
// RecupInfo
procedure RenvoieInfoTarif(TarfMode, CodeTarif: Integer; IsDepot: Boolean; IsCascade: Boolean = False);

var TobArr: Tob;

implementation

procedure CalcPriorite(TOBL: TOB);
var Priorite: Integer;
begin
  Priorite := 0;
  if TOBL.GetValue('GF_ARTICLE') <> '' then Priorite := Priorite + 40;
  if TOBL.GetValue('GF_TARIFARTICLE') <> '' then Priorite := Priorite + 10;
  if TOBL.GetValue('GF_SOUSFAMTARART') <> '' then Priorite := Priorite + 10;
  if TOBL.GetValue('GF_TIERS') <> '' then Priorite := Priorite + 40;
  if TOBL.GetValue('GF_TARIFTIERS') <> '' then Priorite := Priorite + 5;
  if TOBL.GetValue('GF_DEVISE') <> '' then Priorite := Priorite + 2;
  if TOBL.GetValue('GF_DEPOT') <> '' then Priorite := Priorite + 1;
  if TOBL.GetValue('GF_CASCADEREMISE') = 'FOR' then Priorite := Priorite + 200;
  if TOBL.GetValue('GF_CASCADEREMISE') = 'CAS' then Priorite := Priorite + 100;
  if TOBL.GetValue('GF_FERME') = 'X' then Priorite := 0;
  TOBL.PutValue('GF_PRIORITE', Priorite);
end;

function WhereTarifArt(CodeArticle, CodeDevise: string; ttd: T_TableTarif; Totale: boolean): string;
var St: string;
begin
  St := '';
  case ttd of
    ttdArt: St := ' GF_ARTICLE="' + CodeArticle + '" AND GF_TIERS="" AND ' +
      ' GF_TARIFTIERS="" AND GF_TARIFARTICLE="" ' +
        ' AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
    ttdArtQte: St := ' GF_ARTICLE="' + CodeArticle + '" AND GF_TIERS="" AND ' +
      ' GF_TARIFTIERS="" AND GF_TARIFARTICLE="" ' +
        ' AND GF_QUANTITATIF="X" AND GF_DEVISE="' + CodeDevise + '" ';
    ttdArtQCa: St := ' GF_ARTICLE="' + CodeArticle + '" AND GF_TIERS="" AND ' +
      ' GF_TARIFTIERS<>"" AND GF_TARIFARTICLE="" ' +
        ' AND GF_QUANTITATIF="X" AND GF_DEVISE="' + CodeDevise + '" ';
    ttdArtCat: St := ' GF_ARTICLE="' + CodeArticle + '" AND GF_TIERS="" AND ' +
      ' GF_TARIFTIERS<>"" AND GF_TARIFARTICLE="" ' +
        ' AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
  end;
  Result := St;
end;

function WhereTarifCli(CodeTiers, CodeArticle, CodeDevise: string;
  ttd: T_TableTarif; Totale: boolean): string;
var St: string;
begin
  St := '';
  case ttd of
    ttdCliQte: St := ' GF_TIERS="' + CodeTiers + '" AND GF_ARTICLE="' + CodeArticle +
      '" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="" ' +
        ' AND GF_QUANTITATIF="X" AND GF_DEVISE="' + CodeDevise + '" ';
    ttdCliArt: St := ' GF_TIERS="' + CodeTiers + '" AND GF_ARTICLE="' + CodeArticle +
      '" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="" ' +
        ' AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
    ttdCliFam: St := ' GF_TIERS="' + CodeTiers + '" AND GF_ARTICLE="" ' +
      ' AND GF_TARIFTIERS="" AND GF_TARIFARTICLE<>"" ' +
        ' AND GF_QUANTITATIF="-" AND GF_DEVISE="' + CodeDevise + '" ';
  end;
  Result := St;
end;

{$IFDEF MONCODE}

function NbreDecimalSuppPrix(NatTiers: string; TOBArt: TOB): Integer;
var TOBL: TOB;
  PPQ: Double;
  NbDec: integer;
begin
  Result := 0;
  NbDec := 0;
  if TOBArt = nil then exit;
  if TOBArt.GetValue('GA_DECIMALPRIX') = 'X' then
  begin
    if NatTiers = 'CLI' then PPQ := TOBArt.GetValue('GA_PRIXPOURQTE')
    else PPQ := TOBArt.GetValue('GA_PRIXPOURQTEAC');
    if PPQ <= 0 then PPQ := 1;
    while PPQ > 1 do
    begin
      PPQ := arrondi((PPQ / 10.0) - 0.499, 0);
      inc(NbDec);
    end;
    Result := NbDec;
  end;
end;
{$ENDIF}


{================================== ARTICLES =================================}

function ChoisirDimension(St_CodeArt: string; var TOBArt: tob): Boolean;
var QQ: TQuery;
begin
  TheTOB := TOB.Create('', nil, -1);
  AglLanceFiche('GC', 'GCSELECTDIM', '', '', 'GA_ARTICLE=' + St_CodeArt +
    ';ACTION=SELECT;CHAMP= ');
  if TheTOB = nil then
  begin
    Result := False;
  end else
  begin
    Result := True;
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                   TheTOB.Detail[0].GetValue('GA_ARTICLE')+'"',true,-1,'',true);
    if not QQ.EOF then
    begin
      TOBArt.SelectDB('', QQ);
      InitChampsSupArticle (TOBART);
    end;
    Ferme(QQ);
    TheTOB := nil;
  end;
end;

function TrouverTiersSQL(G_CodeTiers: THCritMaskEdit; TOBTiers: TOB): T_RechArt;
var Q: TQuery;
  Etat: string;
begin
  Result := traAucun;
  if G_CodeTiers.Text = '' then exit;
  Q := OpenSQL('Select * from Tiers Where T_TIERS="' +
    G_CodeTiers.Text + '" AND T_NATUREAUXI = "CLI" ', True,-1,'',true);
  if not Q.EOF then
  begin
    TOBTiers.SelectDB('', Q);
    Etat := TOBTiers.GetValue('T_NATUREAUXI');
    if Etat = 'CLI' then Result := traOk;
  end;
  Ferme(Q);
end;

function TrouverFouSQL(G_CodeTiers: THCritMaskEdit; TOBTiers: TOB): T_RechArt;
var Q: TQuery;
  Etat: string;
begin
  Result := traAucun;
  if G_CodeTiers.Text = '' then exit;
  Q := OpenSQL('Select * from Tiers Where T_TIERS="' +
    G_CodeTiers.Text + '" AND T_NATUREAUXI = "FOU" ', True,-1,'',true);
  if not Q.EOF then
  begin
    TOBTiers.SelectDB('', Q);
    Etat := TOBTiers.GetValue('T_NATUREAUXI');
    if Etat = 'FOU' then Result := traOk;
  end;
  Ferme(Q);
end;

{================================== DEPOT ===================================}

procedure GetDepotRecherche(DEPOT: THCritMaskEdit);
begin
  LookupCombo(DEPOT);
end;

function ExisteDepot(Tablette, Valeur: string): boolean;
begin
  Result := ExisteTablette(Tablette, Valeur);
end;

{============================= CATEGORIE TIERS =============================}

procedure GetCategorieRecherche(CATEGORIE: THCritMaskEdit);
begin
  LookupCombo(CATEGORIE);
end;

function ExisteCategorie(Tablette, Valeur: string): boolean;
begin
  Result := ExisteTablette(Tablette, Valeur);
end;

{=========================== FAMILLE TARIF ARTICLE ==========================}

procedure GetFamilleRecherche(FAMILLE: THCritMaskEdit);
begin
  LookupCombo(FAMILLE);
end;

function ExisteFamille(Tablette, Valeur: string): boolean;
begin
  Result := ExisteTablette(Tablette, Valeur);
end;

function ExisteSousFamille (Famille,SousFamille : string) : boolean;
var QQ : TQuery;
begin
  //QQ := OPenSql ('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="BFT" AND CC_CODE="'+SousFamille+'" AND CC_LIBRE="'+Famille+'"',true,-1,'',true);
  QQ := OPenSql ('SELECT BSF_SOUSFAMTARART FROM BTSOUSFAMILLETARIF WHERE BSF_SOUSFAMTARART="'+SousFamille+'" AND BSF_FAMILLETARIF="'+Famille+'"',true,-1,'',true);
  result := not QQ.eof;
  ferme (QQ);
end;

{================================== DATE ====================================}

procedure GetDateRecherche(F: TFORM; DATE: THCritMaskEdit);
var Key: Char;
begin
  Key := '*';
  Paramdate(F, DATE, Key);
end;

{================================== Calcul Remise ===========================}

function RemiseResultante(St: string): extended;
var RemPartiel, Remise: Extended;
begin
  Remise := 1;
  while St <> '' do
  begin
    RemPartiel := ReadTokenPlus(St);
    Remise := Remise * (1 - (RemPartiel / 100));
  end;
  Remise := (1 - Remise) * 100;
  Result := Remise;
end;

function ModifFormat(St: string): string;
var i_ind: integer;
begin
  i_ind := 1;
  while i_ind <= length(St) do
  begin
    if St[i_ind] = '.' then St[i_ind] := V_PGI.SepDecimal;
    i_ind := i_ind + 1;
  end;
  Result := St;
end;

function Numeric(const St: string): Extended;
const Num = ['0'..'9', '.', ',', ' ', '-'];
var i_ind: Integer;
  St_Modif, Rem: string;
begin
  Result := 0;
  rem := St;
  St_Modif := '';
  for i_ind := 1 to Length(Rem) do
  begin
    //if  Rem[i_ind] = '.' then  Rem[i_ind] := V_PGI.SepDecimal;
    if (Rem[i_ind] in Num) and (Ord(Rem[i_ind]) <> 160) then St_Modif := St_Modif + Rem[i_ind];
  end;
  if St_Modif <> '' then Result := StrToFloat(St_Modif);
end;

function ReadTokenPlus(var St: string): Extended;
var i_pos: Integer;
  St_Temp: string;
begin
  i_pos := Pos('+', St);
  if i_pos <= 0 then i_pos := Length(St) + 1;
  St_Temp := Trim(Copy(St, 1, i_pos - 1));
  Result := Numeric(St_Temp);
  Delete(St, 1, i_pos);
end;

{================================== Tablette ===========================}

function ExisteTablette(Tablette, Valeur: string): boolean;
var St: string;
begin
  Result := False;
  if Valeur = '' then
  begin
    Result := True;
    Exit;
  end;
  St := RechDom(Tablette, Valeur, False);
  if (St <> '') and (St <> 'Error') then Result := True;
end;

{=======================Conversion prix contrevaleur ====================}

function ConvertPrix(Prix: Double): Double;
begin
  if VH^.TenueEuro then Result := EuroToFranc(Prix) else Result := FrancToEuro(Prix);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Michel RICHAUD
Créé le ...... : 12/05/2000
Modifié le ... : 29/11/2001 par AC
Description .. : retourne le prix arrondi en fonction des règles d'arrondis
Mots clefs ... : ARRONDI;PRIX
*****************************************************************}

function ArrondirPrix(CodeArr: string; Prix: double; RecupTobArr: Boolean = False): double;
var TOBA: TOB;
  QQ: TQuery;
  i_ind: integer;
begin
  Result := Prix;
  TOBA := nil;
  if CodeArr = '' then exit;
  if (not RecupTobArr) and (TobArr <> nil) then
  begin
    TobArr := nil;
  end;
  if (TobArr <> nil) then TOBA := TobArr.FindFirst(['GAR_CODEARRONDI'], [CodeArr], False);
  if (TobArr = nil) or (TOBA = nil) then
  begin
    QQ := OpenSQL('SELECT * FROM ARRONDI Where GAR_CODEARRONDI="' + CodeArr +
      '" ORDER BY GAR_VALEURSEUIL DESC', TRUE,-1,'',true);
    if QQ.EOF then
    begin
      Ferme(QQ);
      exit
    end else
    begin
      if TOBArr = nil then TOBArr := TOB.Create('', nil, -1);
      TOBArr.LoadDetailDB('ARRONDI', '', '', QQ, False);
    end;
    Ferme(QQ);
  end;
  //TOBA := nil ;
  for i_ind := 0 to TOBArr.detail.count - 1 do
  begin
    if Prix <= TOBArr.Detail[i_ind].GetValue('GAR_VALEURSEUIL') then
      TOBA := TOBArr.Detail[i_ind] else Break;
  end;
  if TOBA = nil then
  begin
    FreeTobArr;
    exit;
  end;
  Prix := CalculArrondi(TOBA, Prix); //AC
  Result := Prix;
  if not RecupTobArr then FreeTobArr; // NA
end;

{***********A.G.L.***********************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 30/11/2001
Modifié le ... :
Description .. : calcul le prix arrondi en fonction des règles d'arrondis
Mots clefs ... : ARRONDI;PRIX
*****************************************************************}

function CalculArrondi(TOBA: Tob; Prix: Double): Double; //AC
var Poids, Con, Px: double;
  Methode: string;
begin
  Poids := TOBA.GetValue('GAR_POIDSARRONDI');
  Con := TOBA.GetValue('GAR_CONSTANTE');
  Methode := TOBA.GetValue('GAR_METHODE');
  Px := Prix;
  if Poids > 0 then
  begin
    if Methode = 'I' then
    begin
      Prix := Prix / Poids;
      if (Prix >= 0) then Prix := Prix + 0.00001 else Prix := Prix - 0.00001;
      Px := Trunc(Prix);
      Px := Px * Poids;
    end else if Methode = 'P' then
    begin
      Px := Arrondi((Prix / Poids), 0);
      Px := Px * Poids;
    end else if Methode = 'S' then
    begin
      if (Prix >= 0) then Px := HEnt1.ARRONDI((Prix / Poids) + 0.4989, 0)
      else Px := HEnt1.ARRONDI((Prix / Poids) - 0.4989, 0);
      Px := Px * Poids;
    end;
  end;
  Prix := Px - Con;
  Result := Prix;
end;

procedure FreeTobArr;
begin
  if TobArr <> nil then
  begin
    TobArr.Free;
    TobArr := nil;
  end;
end;

{=========================================================================================}
{================================== MODE =================================================}
{=========================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 09/11/2001
Modifié le ... :   /  /
Description .. : Retourne le prix contre valeur en fonction des coefficients
Suite ........ : des types de tarif
Mots clefs ... : TARIF;CONTREVALEUR;PRIX
*****************************************************************}

function PrixContreTarif(Prix: Double; Devise, TypeTarif, NatureType: string; TobDev: Tob): Double;
var PrixCon, PrixConArr, CoefTypeDest, PrixPivot: Double;
begin
  if VH^.TenueEuro then ToxConvToDev2(Prix, PrixCon, PrixConArr, Devise, V_PGI.DeviseFongible, Date, TobDev)
  else
  begin
    if Devise <> V_PGI.DevisePivot then
    begin
      CoefTypeDest := RecupCoefTypeTarif(TypeTarif, NatureType);
      PrixPivot := Prix / CoefTypeDest;
    end else PrixPivot := Prix;
    ToxConvToDev2(PrixPivot, PrixCon, PrixConArr, 'FRF', 'EUR', Date, TobDev);
    //DEV.Code := 'EUR' ; DEV.Decimale := 2 ;//GetInfosDevise (DEV) ;
    //PrixCon:=Arrondi(PrixCon,DEV.Decimale) ;
  end;
  result := PrixConArr;
end;

/////MAJ Table tarifMode
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Mise à jour de la table tarif mode
Mots clefs ... : TARIF;TYPE;PERIODE
*****************************************************************}

function TraiterTableTarifMode(CodeType, CodePeriode, CodeNature: string; TobTarifMode: TOB): TOB;
var Q, QMax: TQuery;
  CodeMaxInt: Integer;
  TobType, TobPer: TOB;
  CodeMax, CodeTarifMode: string;
begin
  CodeMaxInt := 1;
  if CodeType = '' then CodeType := '...';
  if TobTarifMode = nil then TobTarifMode := TOB.Create('TARIFMODE', nil, -1) else
  begin
    Result := TobTarifMode;
    Exit;
  end;
  TobType := TOB.Create('TARIFTYPMODE', nil, -1);
  TobPer := TOB.Create('TARIFPER', nil, -1);
  TOBType.SelectDB('"' + CodeType + '";"' + CodeNature + '"', nil);
  TobPer.SelectDB('"...";"' + CodePeriode + '"', nil);
  Q := OpenSql('Select * from TARIFMODE where GFM_TYPETARIF="' + CodeType + '" and GFM_NATURETYPE="' + CodeNature + '" and GFM_PERTARIF="' + CodePeriode + '"',
    True,-1,'',true);
  if not Q.EOF then
  begin
    TobTarifMode.SelectDB('', Q);
    CodeTarifMode := TobTarifMode.GetValue('GFM_TARFMODE');
  end else
  begin
    QMax := OpenSQL('SELECT MAX(GFM_TARFMODE) FROM TARIFMODE', True,-1,'',true);
    if not QMax.EOF then
    begin
      CodeMax := QMax.Fields[0].AsString;
      if CodeMax <> '' then CodeMaxInt := StrToInt(CodeMax) + 1;
    end;
    ferme(QMax);
    CodeTarifMode := IntToStr(CodeMaxInt);
    TobTarifMode.PutValue('GFM_TARFMODE', CodeTarifMode);
    TobTarifMode.PutValue('GFM_PERTARIF', TobPer.GetValue('GFP_CODEPERIODE'));
    TobTarifMode.PutValue('GFM_LIBELLE', copy(TobType.GetValue('GFT_LIBELLE') + '-' + TobPer.GetValue('GFP_LIBELLE'), 1, 35));
    TobTarifMode.PutValue('GFM_DATEDEBUT', TobPer.GetValue('GFP_DATEDEBUT'));
    TobTarifMode.PutValue('GFM_DATEFIN', TobPer.GetValue('GFP_DATEFIN'));
    TobTarifMode.PutValue('GFM_DEMARQUE', TobPer.GetValue('GFP_DEMARQUE'));
    TobTarifMode.PutValue('GFM_ARRONDI', TobPer.GetValue('GFP_ARRONDI'));
    TobTarifMode.PutValue('GFM_PROMO', TobPer.GetValue('GFP_PROMO'));
    TobTarifMode.PutValue('GFM_CASCADE', TobPer.GetValue('GFP_CASCADE'));
    TobTarifMode.PutValue('GFM_COEF', TobType.GetValue('GFT_COEF'));
    TobTarifMode.PutValue('GFM_DEVISE', TobType.GetValue('GFT_DEVISE'));
    TobTarifMode.PutValue('GFM_ETABLISREF', TobType.GetValue('GFT_ETABLISREF'));
    TobTarifMode.PutValue('GFM_TYPETARIF', TobType.GetValue('GFT_CODETYPE'));
    TobTarifMode.PutValue('GFM_NATURETYPE', TobType.GetValue('GFT_NATURETYPE'));
    TobTarifMode.PutValue('GFM_DEVISE', TobType.GetValue('GFT_DEVISE'));
    TobTarifMode.PutValue('GFM_ETABLISREF', TobType.GetValue('GFT_ETABLISREF'));
    if Tobtype.GetValue('GFT_CODETYPE') = '...' then
    begin
      TobTarifMode.PutValue('GFM_TYPETARIF', '...');
      TobTarifMode.PutValue('GFM_COEF', 1);
      TobTarifMode.PutValue('GFM_DEVISE', GetParamSoc('SO_DEVISEPRINC'));
    end;
  end;
  ferme(Q);
  Result := TobTarifMode;
  TobType.Free;
  TobPer.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Récupération des données d'une période pour un
Suite ........ : établissement
Mots clefs ... : TARIF;PERIODE
*****************************************************************}

function RecupInfoPeriode(CodePeriode, Etablissement: string): string;
var Debut, Fin, Arrondi, TypeRemise, Cascade, Demarque, Promo: string;
  Q: TQuery;
begin
  Q := OpenSql('Select * from TARIFPER where GFP_CODEPERIODE="' + CodePeriode + '" and GFP_ETABLISSEMENT="' + Etablissement + '"', True,-1,'',true);
  if not Q.EOF then
  begin
    Debut := Q.FindField('GFP_DATEDEBUT').AsString;
    Fin := Q.FindField('GFP_DATEFIN').AsString;
    Arrondi := Q.FindField('GFP_ARRONDI').AsString;
    Cascade := Q.FindField('GFP_CASCADE').AsString;
    Demarque := Q.FindField('GFP_DEMARQUE').AsString;
    Promo := Q.FindField('GFP_PROMO').AsString;
  end else
  begin
    ferme(Q);
    Q := OpenSql('Select * from TARIFPER where GFP_CODEPERIODE="' + CodePeriode + '" and GFP_ETABLISSEMENT="..."', True,-1,'',true);
    if not Q.EOF then
    begin
      Debut := Q.FindField('GFP_DATEDEBUT').AsString;
      Fin := Q.FindField('GFP_DATEFIN').AsString;
      Arrondi := Q.FindField('GFP_ARRONDI').AsString;
      Cascade := Q.FindField('GFP_CASCADE').AsString;
      Demarque := Q.FindField('GFP_DEMARQUE').AsString;
      Promo := Q.FindField('GFP_PROMO').AsString;
    end else
    begin
      Debut := DateToStr(iDate1900);
      Fin := DateToStr(iDate2099);
      Arrondi := '';
      Cascade := '-';
      Demarque := '';
      Promo := '-'
    end;
  end;
  if Cascade = 'X' then TypeRemise := 'CAS' else TypeRemise := 'MIE';
  Result := Debut + ';' + Fin + ';' + Arrondi + ';' + TypeRemise + ';' + Demarque + ';' + Promo;
  ferme(Q);
end;

function RecupCoefTypeTarif(TypeTarif, NatureType: string): Double;
var Q: TQuery;
begin
  Result := 0;
  Q := OpenSQL('Select GFT_COEF from TarifTypMode Where GFT_CODETYPE="' + TypeTarif + '" AND GFT_NATURETYPE="' + NatureType + '"', True,-1,'',true);
  if not Q.EOF then Result := Q.Fields[0].AsFloat;
  Ferme(Q);
end;

{=========================================================================================}
{================================== Fonction pour les tarifs mode ========================}
{=========================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Recherche un prix hors promotion pour un article
Mots clefs ... : TARIF;PRIX;BASE
*****************************************************************}

function RechPrixTarifBase(TOBTarif, TobMode: Tob; CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): Double;
var i, index: Integer;
  TobTarifBase, TobM: Tob;
  Prix: Double;
  CodeGen: string;
  TypeDefaut, IsDepot: Boolean;
begin
  TobTarifBase := nil;
  TypeDefaut := False;
  IsDepot := False; // Si tarif générique ou spec dépot
  Result := 0;
  Prix := 0;
  CodeGen := CodeArticleUnique(Copy(CodeArt, 1, 18), '', '', '', '', '');
  // Recherche tarif dimensionné et générique avec type de tarif
  TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarfEtab, NatureType], false);
  if TobM = nil then exit;
  Index := TobM.GetIndex;
  for i := Index to TobMode.detail.count - 1 do
  begin
    TobM := TobMode.Detail[i];
    if TobM.GetValue('GFM_TYPETARIF') = '...' then break;
    if TobM.GetValue('GFM_PROMO') = 'X' then continue;
    // Recherche du tarif dimensionné avec depot
    TobTarifBase := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeArt, CodeDepot, CodeDevise,
      TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobTarifBase <> nil) and (TobTarifBase.GetValue('GF_PRIXUNITAIRE') <> 0) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      IsDepot := true;
      Break;
    end else
      // Recherche du tarif générique avec depot
      TobTarifBase := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeGen, CodeDepot, CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobTarifBase <> nil) and (TobTarifBase.GetValue('GF_PRIXUNITAIRE') <> 0) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      IsDepot := true;
      Break;
    end else
      // Recherche du tarif dimensionné sans depot
      TobTarifBase := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeArt, '', CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobTarifBase <> nil) and (TobTarifBase.GetValue('GF_PRIXUNITAIRE') <> 0) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      Break;
    end else
      // Recherche du tarif générique sans depot
      TobTarifBase := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeGen, '', CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobTarifBase <> nil) and (TobTarifBase.GetValue('GF_PRIXUNITAIRE') <> 0) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      Break;
    end;
  end;
  if (TobTarifBase = nil) or (TobTarifBase.GetValue('GF_PRIXUNITAIRE') = 0) then
  begin
    // Recherche tarif dimensionné et générique avec type tarif par défaut
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], ['...', NatureType], false);
    if TobM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') <> '...' then break;
      if TobM.GetValue('GFM_PROMO') = 'X' then continue;
      // Recherche du tarif dimensionné
      TobTarifBase := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_TARFMODE', 'GF_REMISE'], [CodeArt, TobM.GetValue('GFM_TARFMODE'), 0], False);
      if (TobTarifBase <> nil) then
      begin
        TypeDefaut := True;
        Break;
      end
      else
        // Recherche du tarif générique
        TobTarifBase := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_TARFMODE', 'GF_REMISE'], [CodeGen, TobM.GetValue('GFM_TARFMODE'), 0], False);
      if (TobTarifBase <> nil) then
      begin
        TypeDefaut := True;
        Break;
      end;
    end;
  end;
  if TobTarifBase <> nil then
  begin
    Prix := TobTarifBase.GetValue('GF_PRIXUNITAIRE');
  end;
  if TypeDefaut then
  begin
    TOBM := TOBMode.FindFirst(['GFM_DEVISE'], [CodeDevise], False);
    if TOBM <> nil then Prix := Prix * TOBM.GetValue('GFM_COEF');
  end;
  // Utiliser pour justif prix
  if (TobTarifBase <> nil) and (AvecInfo) then RenvoieInfoTarif(TobTarifBase.GetValue('GF_TARFMODE'), TobTarifBase.GetValue('GF_TARIF'), IsDepot);
  Result := Prix;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 28/08/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Calcul d'un prix de vente détail
Mots clefs ... : TARIF;PRIX
*****************************************************************}

function CalculPrixVente(TOBTarif: TOB; CodeArticle, TarifArticle, CodeDepot, CodeDevise, TypeTarfEtab: string; PxVenteArt: Double; AvecInfo: Boolean = False):
  Double;
var Prix, Rem: Double;
  InfoRem, Arrondi: string;
  Q: TQuery;
  TOBMode, TOBM: TOB;
begin
  InfoRem := '';
  Arrondi := '';
  TOBMode := TOB.Create('', nil, -1);
  Q := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO,GFM_COEF,GFM_DEVISE from TARIFMODE where gfm_typetarif in ("' +
    TypeTarfEtab + '","...") order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True,-1,'',true);
  TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
  Ferme(Q);
  InfoRem := RechTarifSpec(TOBTarif, TobMode, CodeArticle, CodeDepot, CodeDevise, TypeTarfEtab, 'VTE');
  Prix := Valeur(ReadTokenSt(InfoRem));
  if Prix = 0 then
  begin
    infoRem := '';
    Prix := RechPrixTarifBase(TOBTarif, TOBMode, CodeArticle, CodeDepot, CodeDevise, TypeTarfEtab, 'VTE');
    InfoRem := ChercheMieRem(TOBTarif, TobMode, CodeArticle, TarifArticle, 'VIDE', 'VIDE', CodeDepot, CodeDevise, TypeTarfEtab, 'VTE');
    Rem := StrToFloat(ReadTokenSt(InfoRem));
    Arrondi := ReadTokenSt(InfoRem);
    if Prix = 0 then
    begin
      Prix := PxVenteArt;
      TOBM := TOBMode.FindFirst(['GFM_DEVISE'], [CodeDevise], False);
      if TOBM <> nil then Prix := Prix * TOBM.GetValue('GFM_COEF');
    end;
    if Rem <> 0 then
    begin
      Prix := Prix * (1 - (Rem / 100));
      Prix := ArrondirPrix(Arrondi, Prix);
    end;
  end;
  if (TheTob = nil) and (AvecInfo) then RenvoieInfoTarif(0, 0, False);

  Result := Prix;
  TOBMode.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Recherche un prix en période promo pour un article
Mots clefs ... : TARIF;PRIX
*****************************************************************}

function RechTarifSpec(TOBTarif, TobMode: Tob; CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): string;
var i, index: Integer;
  TobRechTarifDim, TobM: Tob;
  Prix: Double;
  CodeGen, Demarque: string;
  TypeDefaut, IsDepot: Boolean;
begin
  TobRechTarifDim := nil;
  TypeDefaut := False;
  IsDepot := False;
  Result := '';
  Prix := 0;
  CodeGen := CodeArticleUnique(Copy(CodeArt, 1, 18), '', '', '', '', '');
  // Recherche du tarif dimensionné et générique avec type de tarif
  TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarfEtab, NatureType], false);
  if TobM = nil then exit;
  Index := TobM.GetIndex;
  for i := Index to TobMode.detail.count - 1 do
  begin
    TobM := TobMode.Detail[i];
    if TobM.GetValue('GFM_TYPETARIF') = '...' then break;
    if TobM.GetValue('GFM_PROMO') = '-' then continue;
    // Recherche du tarif dimensionné avec dépot et type de tarif
    TobRechTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeArt, CodeDepot, CodeDevise,
      TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobRechTarifDim <> nil) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      IsDepot := True;
      Break;
    end else
      // Recherche du tarif générique avec dépot et type de tarif
      TobRechTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeGen, CodeDepot, CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobRechTarifDim <> nil) and (TobRechTarifDim.GetValue('GF_PRIXUNITAIRE') <> 0) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      IsDepot := True;
      Break;
    end else
      // Recherche du tarif dimensionné sans dépot et type de tarif
      TobRechTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeArt, '', CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobRechTarifDim <> nil) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      Break;
    end else
      // Recherche du tarif générique sans dépot et type de tarif
      TobRechTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_REMISE'], [CodeGen, '', CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], False);
    if (TobRechTarifDim <> nil) and (TobRechTarifDim.GetValue('GF_PRIXUNITAIRE') <> 0) then
    begin
      TOBTarif.AddChampSup('TYPETARIF', False);
      Break;
    end;
  end;
  if (TobRechTarifDim = nil) or (TobRechTarifDim.GetValue('GF_PRIXUNITAIRE') = 0) then
  begin
    // Recherche tarif dimensionné et générique avec type tarif par défaut
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], ['...', NatureType], false);
    if TOBM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') <> '...' then break;
      if TobM.GetValue('GFM_PROMO') = '-' then continue;
      // Recherche du tarif dimensionné
      TobRechTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_TARFMODE', 'GF_REMISE'], [CodeArt, TobM.GetValue('GFM_TARFMODE'), 0], False);
      if (TobRechTarifDim <> nil) then
      begin
        TypeDefaut := True;
        Break;
      end
      else
        // Recherche du tarif générique
        TobRechTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_TARFMODE', 'GF_REMISE'], [CodeGen, TobM.GetValue('GFM_TARFMODE'), 0], False);
      if (TobRechTarifDim <> nil) then
      begin
        TypeDefaut := True;
        Break;
      end;
    end;
  end;
  if TobRechTarifDim <> nil then
  begin
    Prix := TobRechTarifDim.GetValue('GF_PRIXUNITAIRE');
    Demarque := TobRechTarifDim.GetValue('GF_DEMARQUE');
  end;
  if TypeDefaut then
  begin
    TOBM := TOBMode.FindFirst(['GFM_DEVISE'], [CodeDevise], False);
    if TOBM <> nil then Prix := Prix * TOBM.GetValue('GFM_COEF');
  end;
  // Utiliser pour justif prix
  if (TobRechTarifDim <> nil) and (AvecInfo) then RenvoieInfoTarif(TobRechTarifDim.GetValue('GF_TARFMODE'), TobRechTarifDim.GetValue('GF_TARIF'), IsDepot);
  result := FloatToStr(Prix) + ';' + Demarque;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Retourne une tob tarif toutes dim
Mots clefs ... : TARIF;TOB
*****************************************************************}

function CreerTobTarifArt(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat: string; LaDate: TDateTime; EnHT: Boolean;
  AvecDevise: Boolean = True): TOB;
var //VenteAchat,
  SQL: string;
  QTarif: TQuery;
  //EnHT: Boolean ;
  TOBTarif: Tob;
begin
  TOBTarif := TOB.Create('', nil, -1);
  TOBTarif.AddChampSup('_TARFART', False);
  TOBTarif.PutValue('_TARFART', 'TarfArt');
  //VenteAchat:='VEN';
  //EnHT:=False ;
  SQL := SQLTarifGen(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, 0, AvecDevise);
  SQL := SQL + CompleteSQLTarif(VenteAchat, False, EnHT);
  QTarif := OpenSQL(SQL, True,-1,'',true);
  if not QTarif.EOF then TOBTarif.LoadDetailDB('TARIF', '', '', QTarif, False);
  Ferme(QTarif);
  Result := TOBTarif;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Retourne une tob tarif Dim
Mots clefs ... : TARIF;TOB
*****************************************************************}

function CreerTobTarifArtDim(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat: string; LaDate: TDateTime; EnHT: Boolean;
  AvecDevise: Boolean = True): TOB;
var SQL: string;
  QTarif: TQuery;
  TOBTarif: Tob;
  CodeDimGen: string;
begin
  TOBTarif := TOB.Create('', nil, -1);
  TOBTarif.AddChampSup('_TARFART', False);
  TOBTarif.PutValue('_TARFART', 'TarfArt');
  CodeDimGen := '"' + CodeArt + '" ,"' + CodeArticleUnique2(Copy(CodeArt, 1, 18), '') + '"';
  SQL := SQLTarifDim(CodeDimGen, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, 0, AvecDevise);
  SQL := SQL + CompleteSQLTarif(VenteAchat, False, EnHT);
  QTarif := OpenSQL(SQL, True,-1,'',true);
  if not QTarif.EOF then TOBTarif.LoadDetailDB('TARIF', '', '', QTarif, False);
  Ferme(QTarif);
  Result := TOBTarif;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Construit la requête pour remplir la tob tarif (toutes dim)
Mots clefs ... : TARIF;REQUETE
*****************************************************************}

function SQLTarifGen(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot: string; LaDate: TDateTime; Qte: double; AvecDevise: Boolean = True):
  string;
var SQL: string;
begin
  SQL := 'Select * From TARIF Where ' +
    '((GF_ARTICLE like "' + Copy(CodeArt, 1, 18) + '%" AND GF_TIERS="' + CodeTiers + '") OR' +
    ' (GF_ARTICLE like "' + Copy(CodeArt, 1, 18) + '%" AND GF_TIERS="" AND GF_TARIFTIERS="") OR';
  SQL := SQL +
    ' (GF_ARTICLE like "' + Copy(CodeArt, 1, 18) + '%" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '") OR';
  SQL := SQL +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="' + TarifArticle + '") OR';
  SQL := SQL +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '" AND GF_TARIFARTICLE="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE=""))';
  SQL := SQL + ' AND GF_DATEDEBUT<="' + USDatetime(LaDate) + '" AND GF_DATEFIN>="' + USDatetime(LaDate) + '"' +
    ' AND GF_BORNEINF<=' + StrFPoint(Qte) + ' AND GF_BORNESUP>=' + StrFPoint(Qte) +
    ' AND GF_FERME="-"';
  if AvecDevise then SQL := SQL + ' AND GF_DEVISE="' + CodeDevise + '"';
  Result := SQL;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Construit la requête pour remplir la tob tarif(Dim + Gen)
Mots clefs ... : TARIF;REQUETE
*****************************************************************}

function SQLTarifDim(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot: string; LaDate: TDateTime; Qte: double; AvecDevise: Boolean = True):
  string;
var SQL: string;
begin
  SQL := 'Select * From TARIF Where ' +
    '((GF_ARTICLE in (' + CodeArt + ') AND GF_TIERS="' + CodeTiers + '") OR' +
    ' (GF_ARTICLE in (' + CodeArt + ') AND GF_TIERS="" AND GF_TARIFTIERS="") OR';
  SQL := SQL +
    ' (GF_ARTICLE in (' + CodeArt + ') AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '") OR';
  SQL := SQL +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="' + TarifArticle + '") OR';
  SQL := SQL +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '" AND GF_TARIFARTICLE="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE=""))';
  SQL := SQL + ' AND GF_DATEDEBUT<="' + USDatetime(LaDate) + '" AND GF_DATEFIN>="' + USDatetime(LaDate) + '"' +
    ' AND GF_BORNEINF<=' + StrFPoint(Qte) + ' AND GF_BORNESUP>=' + StrFPoint(Qte) +
    ' AND GF_FERME="-"';
  if AvecDevise then SQL := SQL + ' AND GF_DEVISE="' + CodeDevise + '"';
  Result := SQL;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Retourne la meilleur remise entre l'article et le tiers
Mots clefs ... : TARIF;REMISE
*****************************************************************}

function ChercheMieRem(TOBTarif, TobMode: Tob; CodeArt, TarfArt, CodeTiers, TarfTiers, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string): string;
var RemiseTiers, RemiseCatArt, RemiseArt, Rem: Double;
  InfoArt, ArrondiArt, InfoCat, ArrondiCat, Arrondi, Demarque: string;
begin
  RemiseTiers := CalcRemCli(TOBTarif, TOBMode, TarfArt, CodeTiers, TarfTiers, CodeDepot, CodeDevise, TypeTarfEtab, NatureType);
  InfoCat := CalcRemCatArt(TOBTarif, TOBMode, TarfArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType);
  RemiseCatArt := Valeur(ReadTokenSt(InfoCat));
  ArrondiCat := ReadTokenSt(InfoCat);
  InfoArt := CalcRemArt(TOBTarif, TOBMode, CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType);
  RemiseArt := Valeur(ReadTokenSt(InfoArt));
  ArrondiArt := ReadTokenSt(InfoArt);
  Demarque := ReadTokenSt(InfoArt);
  if TobTarif.FieldExists('TYPETARIF') then TobTarif.DelChampSup('TYPETARIF', False);
  if RemiseTiers > RemiseCatArt then
  begin
    Rem := RemiseTiers;
    Demarque := '';
  end else
  begin
    Rem := RemiseCatArt;
    Arrondi := ArrondiCat;
  end;
  if RemiseArt > Rem then
  begin
    Rem := RemiseArt;
    Arrondi := ArrondiArt;
  end;
  Result := FloatToStr(Rem) + ';' + Arrondi + ';' + Demarque;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Calcul une remise article .
Suite ........ : Renvoi l'arrondi à appliquer sur le montant
Mots clefs ... : TARIF;REMISE;ARTICLE
*****************************************************************}

function CalcRemArt(TOBTarif, TOBMode: TOB; CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): string;
var TobTarifDim, TobM: Tob;
  QuelleCascade, CodeGen, Arrondi, Demarque: string;
  Remise: Double;
  i, index: Integer;
  TypeEtab, CodeDt: string;
  TypeTarif, IsDepot, IsCascade: Boolean;
begin
  TypeTarif := False;
  TobTarifDim := nil;
  IsDepot := False;
  IsCascade := False;
  Remise := 0;
  Arrondi := '';
  Demarque := '';
  Result := FloatToStr(Remise) + ';' + Arrondi + ';' + Demarque;
  CodeDt := CodeDepot;
  TypeEtab := TypeTarfEtab;
  CodeGen := CodeArticleUnique(Copy(CodeArt, 1, 18), '', '', '', '', '');
  TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarfEtab, NatureType], false);
  if TobM = nil then exit;
  Index := TobM.GetIndex;
  // Recherche tarif dimensionné et générique type de tarif
  for i := Index to TobMode.detail.count - 1 do
  begin
    TobM := TobMode.Detail[i];
    if TobM.GetValue('GFM_TYPETARIF') = '...' then break;
    if TobM.GetValue('GFM_PROMO') = '-' then continue;
    // Recherche d'une remise dimensionné avec depot
    TobTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [CodeArt, CodeDepot, CodeDevise,
      TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then
    begin
      TypeTarif := True;
      IsDepot := True;
      break;
    end else
      // Recherche d'une remise generique avec depot
      TobTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [CodeGen, CodeDepot, CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], false);
    if TobTarifDim <> nil then
    begin
      TypeTarif := True;
      IsDepot := True;
      break;
    end else
      // Recherche d'une remise dimensionné sans depot
      TobTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [CodeArt, '', CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then
    begin
      TypeTarif := True;
      break;
    end else
      // Recherche d'une remise generique sans depot
      TobTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [CodeGen, '', CodeDevise,
        TobM.GetValue('GFM_TARFMODE'), 0], false);
    if TobTarifDim <> nil then
    begin
      TypeTarif := True;
      break;
    end;
  end;
  // Recherche tarif dimensionné et générique sans dépot et type de tarif par défaut
  if (TobTarifDim = nil) and (not TobTarif.FieldExists('TYPETARIF')) then
  begin
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], ['...', NatureType], false);
    if TOBM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      CodeDt := '';
      TypeEtab := '...';
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') <> '...' then break;
      if TobM.GetValue('GFM_PROMO') = '-' then continue;
      // Tarif dimensionné
      TobTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [CodeArt, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if TobTarifDim <> nil then Break
      else
        // tarif Générique
        TobTarifDim := TOBTarif.FindFirst(['GF_ARTICLE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [CodeGen, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if TobTarifDim <> nil then break;
    end;
  end;
  if TobTarifDim <> nil then
  begin
    CodeArt := TobTarifDim.GetValue('GF_ARTICLE');
    QuelleCascade := TobTarifDim.GetValue('GF_CASCADEREMISE');
    Remise := TobTarifDim.GetValue('GF_REMISE');
    Arrondi := TobTarifDim.GetValue('GF_ARRONDI');
    Demarque := TobTarifDim.GetValue('GF_DEMARQUE');
    CodeDt := TobTarifDim.GetValue('GF_DEPOT');
    i := TobTarifDim.GetIndex;
    if (TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE') = 'CAS') then
    begin
      repeat
        i := i + 1;
        if i = TobTarif.Detail.Count then Break;
        if TypeTarif then
        begin
          if (CodeArt <> TOBTarif.Detail[i].GetValue('GF_ARTICLE')) or (CodeDt <> TOBTarif.Detail[i].GetValue('GF_DEPOT')) or (not
            ControlTypeTarifEtab(TobMode, TypeEtab, IntToStr(TOBTarif.Detail[i].GetValue('GF_TARFMODE')))) then Continue
        end
        else
          if (CodeArt <> TOBTarif.Detail[i].GetValue('GF_ARTICLE')) or (CodeDt <> TOBTarif.Detail[i].GetValue('GF_DEPOT')) then Continue;
        QuelleCascade := TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE');
        if QuelleCascade = 'CAS' then IsCascade := True;
        Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - TOBTarif.Detail[i].GetValue('GF_REMISE') / 100));
      until (i > TobTarif.Detail.Count) or ((TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE') <> 'CAS') and ControlTypeTarifEtab(TobMode, TypeEtab,
        IntToStr(TOBTarif.Detail[i].GetValue('GF_TARFMODE'))));
    end;
  end;
  // Utiliser pour justif prix
  if (TobTarifDim <> nil) and (AvecInfo) then RenvoieInfoTarif(TobTarifDim.GetValue('GF_TARFMODE'), TobTarifDim.GetValue('GF_TARIF'), IsDepot, IsCascade);
  Result := FloatToStr(Remise) + ';' + Arrondi + ';' + Demarque;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Calcul une remise catégorie article.
Suite ........ :
Mots clefs ... : TARIF;REMISE;ARTICLE;CATEGORIE
*****************************************************************}

function CalcRemCatArt(TOBTarif, TOBMode: TOB; TarfArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string; AvecInfo: Boolean = False): string;
var TobTarifDim, TobM: Tob;
  QuelleCascade, Arrondi: string;
  Remise: Double;
  i, index: Integer;
  TypeEtab, CodeDt: string;
  TypeTarif, IsDepot, IsCascade: Boolean;
begin
  TypeTarif := False;
  TobTarifDim := nil;
  IsDepot := False;
  IsCascade := False;
  Remise := 0;
  Arrondi := '';
  Result := FloatToStr(Remise) + ';' + Arrondi;
  CodeDt := CodeDepot;
  TypeEtab := TypeTarfEtab;
  TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarfEtab, NatureType], false);
  if TobM = nil then exit;
  Index := TobM.GetIndex;
  // Recherche tarif catégorie article type de tarif
  for i := Index to TobMode.detail.count - 1 do
  begin
    TobM := TobMode.Detail[i];
    if TobM.GetValue('GFM_TYPETARIF') = '...' then break;
    if TobM.GetValue('GFM_PROMO') = '-' then continue;
    // Recherche tarif catégorie article avec depot
    TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt,
      '',
        '', CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then
    begin
      TypeTarif := True;
      IsDepot := True;
      break;
    end else
      // Recherche tarif catégorie article sans depot
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt,
        '', '', '', CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then
    begin
      TypeTarif := True;
      break;
    end;
  end;
  // Recherche tarif catégorie article sans dépot et type de tarif par défaut
  if (TobTarifDim = nil) and (not TobTarif.FieldExists('TYPETARIF')) then
  begin
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], ['...', NatureType], false);
    if TOBM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      CodeDt := '';
      TypeEtab := '...';
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') <> '...' then break;
      if TobM.GetValue('GFM_PROMO') = '-' then continue;
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARIFTIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, '', '',
        TobM.GetValue('GFM_TARFMODE'), 0], false);
      if TobTarifDim <> nil then break;
    end;
  end;
  if TobTarifDim <> nil then
  begin
    TarfArt := TobTarifDim.GetValue('GF_TARIFARTICLE');
    QuelleCascade := TobTarifDim.GetValue('GF_CASCADEREMISE');
    Remise := TobTarifDim.GetValue('GF_REMISE');
    Arrondi := TobTarifDim.GetValue('GF_ARRONDI');
    i := TobTarifDim.GetIndex;
    if (TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE') = 'CAS') then
    begin
      repeat
        i := i + 1;
        if i = TobTarif.Detail.Count then Break;
        if (TOBTarif.Detail[i].GetValue('GF_TIERS') <> '') or (TOBTarif.Detail[i].GetValue('GF_TARIFTIERS') <> '') then continue;
        if TypeTarif then
        begin
          if (TarfArt <> TOBTarif.Detail[i].GetValue('GF_TARIFARTICLE')) or (CodeDt <> TOBTarif.Detail[i].GetValue('GF_DEPOT')) or (not
            ControlTypeTarifEtab(TobMode, TypeEtab, IntToStr(TOBTarif.Detail[i].GetValue('GF_TARFMODE')))) then Continue
        end
        else
          if (TarfArt <> TOBTarif.Detail[i].GetValue('GF_TARIFARTICLE')) or (CodeDt <> TOBTarif.Detail[i].GetValue('GF_DEPOT')) then Continue;
        QuelleCascade := TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE');
        if QuelleCascade = 'CAS' then IsCascade := True;
        Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - TOBTarif.Detail[i].GetValue('GF_REMISE') / 100));
      until (i > TobTarif.Detail.Count) or ((TOBTarif.Detail[i].GetValue('GF_CASCADEREMISE') <> 'CAS') and ControlTypeTarifEtab(TobMode, TypeEtab,
        IntToStr(TOBTarif.Detail[i].GetValue('GF_TARFMODE'))) and ((TOBTarif.Detail[i].GetValue('GF_TIERS') = '') and
        (TOBTarif.Detail[i].GetValue('GF_TARIFTIERS') = '')));
    end;
  end;
  if (TobTarifDim <> nil) and (AvecInfo) then RenvoieInfoTarif(TobTarifDim.GetValue('GF_TARFMODE'), TobTarifDim.GetValue('GF_TARIF'), IsDepot, IsCascade);
  Result := FloatToStr(Remise) + ';' + Arrondi;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Spécifique mode: Tarif détail
Suite ........ : Calcul une remise client.
Suite ........ :
Mots clefs ... : TARIF;REMISE;CLIENT
*****************************************************************}

function CalcRemCli(TOBTarif, TOBMode: TOB; TarfArt, Codetiers, TarfTiers, CodeDepot, CodeDevise, TypeTarfEtab, NatureType: string): Double;
var TobTarifDim, TobM: Tob;
  Remise1, Remise2, RemCliCatArt, RemCli, RemCatCliCatArt, RemCatCli: Double;
  i, index: Integer;
  TypeEtab, CodeDt: string;
begin
  Result := 0;
  RemCliCatArt := 0;
  RemCli := 0;
  RemCatCliCatArt := 0;
  RemCatCli := 0;
  CodeDt := CodeDepot;
  TypeEtab := TypeTarfEtab;
  TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarfEtab, NatureType], false);
  if TobM = nil then exit;
  Index := TobM.GetIndex;
  // Recherche tarif du client type de tarif ds période promotion
  for i := Index to TobMode.detail.count - 1 do
  begin
    TobM := TobMode.Detail[i];
    if TobM.GetValue('GFM_TYPETARIF') = '...' then break;
    //if TobM.GetValue('GFM_PROMO')='-' then continue ;
    // AVEC DEPOT
    // Recherche du tarif client + catégorie article
    TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, Codetiers,
      CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then RemCliCatArt := TobTarifDim.GetValue('GF_REMISE');
    // Recherche du tarif client
    TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', Codetiers, CodeDepot,
      CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then RemCli := TobTarifDim.GetValue('GF_REMISE');
    // Recherche de catégorie client + catégorie article
    TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, TarfTiers,
      CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then RemCatCliCatArt := TobTarifDim.GetValue('GF_REMISE');
    // Recherche de catégorie client
    TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', TarfTiers,
      CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
    if (TobTarifDim <> nil) then RemCatCli := TobTarifDim.GetValue('GF_REMISE');
    // SANS DEPOT
    if (RemCliCatArt = 0) and (RemCli = 0) and (RemCatCliCatArt = 0) and (RemCatCli = 0) then
    begin
      // Recherche du tarif client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, Codetiers, '',
        CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche du tarif client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', Codetiers, '',
        CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCli := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, TarfTiers,
        '', CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCatCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', TarfTiers, '',
        CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCatCli := TobTarifDim.GetValue('GF_REMISE');
    end;
  end;
  // Recherche tarif catégorie article sans dépot et type de tarif par défaut ds période promotion
  if (RemCliCatArt = 0) and (RemCli = 0) and (RemCatCliCatArt = 0) and (RemCatCli = 0) then
  begin
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], ['...', NatureType], false);
    if TOBM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      CodeDt := '';
      TypeEtab := '...';
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') <> '...' then break;
      if TobM.GetValue('GFM_PROMO') = '-' then continue;
      // Recherche du tarif client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, Codetiers, TobM.GetValue('GFM_TARFMODE'),
        0], false);
      if (TobTarifDim <> nil) then RemCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche du tarif client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', Codetiers, TobM.GetValue('GFM_TARFMODE'), 0],
        false);
      if (TobTarifDim <> nil) then RemCli := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, TarfTiers,
        TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCatCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', TarfTiers, TobM.GetValue('GFM_TARFMODE'),
        0], false);
      if (TobTarifDim <> nil) then RemCatCli := TobTarifDim.GetValue('GF_REMISE');
    end;
  end;
  // Recherche tarif du client avec dépot et type de tarif ds période hors promotion
  if (RemCliCatArt = 0) and (RemCli = 0) and (RemCatCliCatArt = 0) and (RemCatCli = 0) then
  begin
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], [TypeTarfEtab, NatureType], false);
    if TobM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') = '...' then break;
      if TobM.GetValue('GFM_PROMO') = '-' then continue;
      // Recherche du tarif client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, Codetiers,
        CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche du tarif client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', Codetiers, CodeDepot,
        CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCli := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, TarfTiers,
        CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCatCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_DEPOT', 'GF_DEVISE', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', TarfTiers,
        CodeDepot, CodeDevise, TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCatCli := TobTarifDim.GetValue('GF_REMISE');
    end;
  end;
  // Recherche tarif catégorie article sans dépot et type de tarif par défaut hors promotion
  if (RemCliCatArt = 0) and (RemCli = 0) and (RemCatCliCatArt = 0) and (RemCatCli = 0) then
  begin
    TobM := TobMode.FindFirst(['GFM_TYPETARIF', 'GFM_NATURETYPE'], ['...', NatureType], false);
    if TOBM = nil then exit;
    Index := TobM.GetIndex;
    for i := Index to TobMode.detail.count - 1 do
    begin
      CodeDt := '';
      TypeEtab := '...';
      TobM := TobMode.Detail[i];
      if TobM.GetValue('GFM_TYPETARIF') <> '...' then break;
      if TobM.GetValue('GFM_PROMO') <> '-' then continue;
      // Recherche du tarif client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, Codetiers, TobM.GetValue('GFM_TARFMODE'),
        0], false);
      if (TobTarifDim <> nil) then RemCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche du tarif client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', Codetiers, TobM.GetValue('GFM_TARFMODE'), 0],
        false);
      if (TobTarifDim <> nil) then RemCli := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client + catégorie article
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], [TarfArt, TarfTiers,
        TobM.GetValue('GFM_TARFMODE'), 0], false);
      if (TobTarifDim <> nil) then RemCatCliCatArt := TobTarifDim.GetValue('GF_REMISE');
      // Recherche de catégorie client
      TobTarifDim := TOBTarif.FindFirst(['GF_TARIFARTICLE', 'GF_TARIFTIERS', 'GF_TARFMODE', 'GF_PRIXUNITAIRE'], ['', TarfTiers, TobM.GetValue('GFM_TARFMODE'),
        0], false);
      if (TobTarifDim <> nil) then RemCatCli := TobTarifDim.GetValue('GF_REMISE');
    end;
  end;
  if RemCliCatArt > RemCli then Remise1 := RemCliCatArt else remise1 := RemCli;
  if RemCatCliCatArt > RemCatCli then Remise2 := RemCatCliCatArt else remise2 := RemCatCli;
  if Remise1 > Remise2 then result := Remise1 else Result := Remise2;
end;

function ControlTypeTarifEtab(TobMode: Tob; TypeTarif, TarfMode: string): boolean;
var TOBM: TOB;
begin
  Result := False;
  TobM := TobMode.FindFirst(['GFM_TARFMODE'], [TarfMode], False);
  if TobM <> nil then
    if (TobM.GetValue('GFM_TYPETARIF') = TypeTarif) and (TobM.GetValue('GFM_PROMO') = 'X') then result := True;
end;

procedure RenvoieInfoTarif(TarfMode, CodeTarif: Integer; IsDepot: Boolean; IsCascade: Boolean = False);
var TOBInfo: TOB;
begin
  TOBInfo := TOB.Create('info', nil, -1);
  TOBInfo.AddChampSup('TARFMODE', False);
  TOBInfo.AddChampSup('CODETARIF', False);
  TOBInfo.AddChampSup('ISDEPOT', False);
  TOBInfo.AddChampSup('ISCASCADE', False);
  TOBInfo.PutValue('TARFMODE', TarfMode);
  TOBInfo.PutValue('CODETARIF', CodeTarif);
  TOBInfo.PutValue('ISDEPOT', IsDepot);
  TOBInfo.PutValue('ISCASCADE', IsCascade);
  TheTob := TOBInfo;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 17/05/2002
Modifié le ... : 22/05/2002
Description .. : Fonction qui retourne le prix de base, le prix de solde ainsi
Suite ........ : que les dates de debut et de fin à partir du codearticle et de
Suite ........ : l'établisssement
Mots clefs ... : TARIF;ARTICLE;ETABLISSEMENT
*****************************************************************}

procedure FindTarifUnArticle(TOBA, TOBMode: TOB; Etablissement, TypeTarif, Devise: string; var CodeArticle: string; var TOBTarif: TOB);
var InfoRem, DateDeb, DateFin, Arrondi: string;
  PxBase, PxSolde, Rem: Double;
begin
  if TOBA.GetValue('GA_TYPEARTICLE') <> 'MAR' then Exit;
  if (CodeArticle = '') or (CodeArticle <> TOBA.GetValue('GA_CODEARTICLE')) then
  begin
    CodeArticle := TOBA.GetValue('GA_CODEARTICLE');
    TOBTarif := CreerTobTarifArt(TOBA.GetValue('GA_CODEARTICLE'), '', '', '', Devise, Etablissement, 'VEN', Date, False);
  end;
  if CodeArticle = TOBA.GetValue('GA_CODEARTICLE') then
  begin
    PxBase := RechPrixTarifBase(TOBTarif, TOBMode, TOBA.GetValue('GA_ARTICLE'), Etablissement, Devise, TypeTarif, 'VTE');
    if PxBase = 0 then PxBase := TOBA.GetValue('GA_PVTTC');
    InfoRem := RechTarifSpec(TOBTarif, TOBMode, TOBA.GetValue('GA_ARTICLE'), Etablissement, Devise, TypeTarif, 'VTE');
    PxSolde := Valeur(ReadTokenSt(InfoRem));
    if PxSolde = 0 then
    begin
      infoRem := '';
      InfoRem := ChercheMieRem(TOBTarif, TobMode, TOBA.GetValue('GA_ARTICLE'), '', 'VIDE', 'VIDE', Etablissement, Devise, TypeTarif, 'VTE');
      Rem := StrToFloat(ReadTokenSt(InfoRem));
      Arrondi := ReadTokenSt(InfoRem);
      if Rem <> 0 then
      begin
        PxSolde := PxBase * (1 - (Rem / 100));
        PxSolde := ArrondirPrix(Arrondi, PxSolde);
        ReadTokenSt(InfoRem);
        DateDeb := ReadTokenSt(InfoRem);
        DateFin := ReadTokenSt(InfoRem);
      end else
      begin
        PxSolde := PxBase;
        DateDeb := DateToStr(iDate1900);
        DateFin := DateToStr(iDate2099);
      end;
    end
    else
    begin
      ReadTokenSt(InfoRem);
      DateDeb := ReadTokenSt(InfoRem);
      DateFin := ReadTokenSt(InfoRem);
    end;
    TOBA.AddChampSup('PXBASE', False);
    TOBA.AddChampSup('PXSOLDE', False);
    TOBA.AddChampSup('DATEDEB', False);
    TOBA.AddChampSup('DATEFIN', False);
    TOBA.PutValue('PXBASE', PxBase);
    TOBA.PutValue('PXSOLDE', PxSolde);
    TOBA.PutValue('DATEDEB', StrToDate(DateDeb));
    TOBA.PutValue('DATEFIN', StrToDate(DateFin));
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 17/05/2002
Modifié le ... :   /  /
Description .. : Fonction qui retourne le prix de base, le prix de solde ainsi
Suite ........ : que les dates de debut et de fin à partir du codearticle et de
Suite ........ : l'établisssement
Mots clefs ... : TARIF;ARTICLE;ETABLISSEMENT
*****************************************************************}

procedure FindTarifArt(TOBArt: TOB; Etablissement: string; UnArticle: Boolean = False);
var TypeTarif, Devise, CodeArticle: string;
  i: integer;
  TOBA, TOBMode, TOBTarif: TOB;
  Q: TQuery;
begin
  TOBTarif := nil;
  TOBMode := TOB.Create('', nil, -1);
  // Info Etablissement
  Q := OpenSQL('Select ET_TYPETARIF,ET_DEVISE from ETABLISS Where ET_ETABLISSEMENT="' + Etablissement + '"', True,-1,'',true);
  if not Q.EOF then
  begin
    TypeTarif := Q.FindField('ET_TYPETARIF').AsString;
    Devise := Q.FindField('ET_DEVISE').AsString;
  end;
  if TypeTarif = '' then TypeTarif := '...';
  if Devise = '' then Devise := 'EUR';
  Ferme(Q);
  // Info TarifMode
  Q := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO,GFM_COEF,GFM_DEVISE from TARIFMODE where gfm_typetarif in ("' +
    TypeTarif + '","...") and gfm_naturetype="VTE" order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True,-1,'',true);
  TOBMode.LoadDetailDB('TARIFMODE', '', '', Q, False);
  Ferme(Q);
  CodeArticle := '';
  if UnArticle then
  begin
    FindTarifUnArticle(TOBArt, TOBMode, Etablissement, TypeTarif, Devise, CodeArticle, TOBTarif);
  end else
  begin
    TOBArt.Detail.Sort('GA_CODEARTICLE');
    for i := 0 to TOBArt.Detail.count - 1 do
    begin
      TOBA := TOBArt.Detail[i];
      FindTarifUnArticle(TOBA, TOBMode, Etablissement, TypeTarif, Devise, CodeArticle, TOBTarif);
    end;
  end;
end;

function RecupCodeTarifMode(TypeTarif, Periode, NatureType: string): Integer;
var Q: TQuery;
begin
  Result := -1;
  Q := OpenSQL('Select GFM_TARFMODE from TarifMode Where GFM_TYPETARIF="' + TypeTarif + '" and GFM_PERTARIF="' + Periode + '" and GFM_NATURETYPE="' + NatureType
    + '"', True,-1,'',true);
  if not Q.EOF then Result := Q.Fields[0].AsInteger;
  Ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : La devise est elle l'EURO
Mots clefs ... :
*****************************************************************}

function ToxTestCodeEuro(CodeDev: string): Boolean;
begin
  Result := (CodeDev = 'EUR');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : convertit un montant en une autre devise
Mots clefs ... : TOX
*****************************************************************}

function ToxConvToDev(Montant: Double; DevOrg, DevDest: string; DatePiece: TDateTime; TOBDev: TOB): Double;
var Taux: Double;
  DateTaux: TDateTime;
  TOrg, TDest: TOB;
begin
  Result := Montant;
  TOrg := nil;
  TDest := nil;
  if DevOrg <> DevDest then
  begin
    if TOBDev = nil then
    begin
      TOrg := TOB.Create('DEVISE', nil, -1);
      TOrg.SelectDB('"' + DevOrg + '"', nil, False);
    end else
      TOrg := TOBDev.FindFirst(['D_DEVISE'], [DevOrg], False);

    if TOrg = nil then Exit;

    if TOBDev = nil then
    begin
      TDest := TOB.Create('DEVISE', nil, -1);
      TDest.SelectDB('"' + DevDest + '"', nil, False);
    end else
      TDest := TOBDev.FindFirst(['D_DEVISE'], [DevDest], False);

    if TDest = nil then Exit;

    if (ToxTestCodeEuro(DevOrg)) or (ToxTestCodeEuro(DevDest)) then
    begin
      if not ToxTestCodeEuro(DevOrg) then
      begin
        // Devise d'origine en devise pivot
        Taux := GetTaux(DevOrg, DateTaux, DatePiece);
        Result := DeviseToEuro(Result, Taux, TOrg.GetValue('D_QUOTITE'));
      end;
      if not ToxTestCodeEuro(DevDest) then
      begin
        // Devise pivot en devise destination
        Taux := GetTaux(DevDest, DateTaux, DatePiece);
        Result := EuroToDevise(Result, Taux, TDest.GetValue('D_QUOTITE'), TDest.GetValue('D_DECIMALE'));
      end;
    end else
    begin
      if DevOrg <> V_PGI.DevisePivot then
      begin
        // Devise d'origine en devise pivot
        Taux := GetTaux(DevOrg, DateTaux, DatePiece);
        Result := DeviseToPivot(Result, Taux, TOrg.GetValue('D_QUOTITE'));
      end;
      if DevDest <> V_PGI.DevisePivot then
      begin
        // Devise pivot en devise destination
        Taux := GetTaux(DevDest, DateTaux, DatePiece);
        Result := PivotToDevise(Result, Taux, TDest.GetValue('D_QUOTITE'), TDest.GetValue('D_DECIMALE'));
      end;
    end;
  end;
  // Libération des TOB créées.
  if TOBDev = nil then
  begin
    if TOrg <> nil then TOrg.Free;
    if TDEst <> nil then TDest.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : convertit un montant en une autre devise
Mots clefs ... : TOX
*****************************************************************}

procedure ToxConvToDev2(MontantOrg: Double; var MontantDestNA, MontantDest: Double; DevOrg, DevDest: string; DatePiece: TDateTime; TOBDev: TOB);
var Taux: Double;
  DateTaux: TDateTime;
  TOrg, TDest: TOB;
begin
  //
  // initialisation des montants
  //
  MontantDestNA := MontantOrg;
  MontantDest := MontantOrg;

  TOrg := nil;
  TDest := nil;

  if DevOrg <> DevDest then
  begin
    if TOBDev = nil then
    begin
      TOrg := TOB.Create('DEVISE', nil, -1);
      TOrg.SelectDB('"' + DevOrg + '"', nil, False);
    end else
    begin
      TOrg := TOBDev.FindFirst(['D_DEVISE'], [DevOrg], False);
    end;

    if TOrg = nil then Exit;

    if TOBDev = nil then
    begin
      TDest := TOB.Create('DEVISE', nil, -1);
      TDest.SelectDB('"' + DevDest + '"', nil, False);
    end else
    begin
      TDest := TOBDev.FindFirst(['D_DEVISE'], [DevDest], False);
    end;

    if TDest = nil then Exit;

    if (ToxTestCodeEuro(DevOrg)) or (ToxTestCodeEuro(DevDest)) then
    begin
      if not ToxTestCodeEuro(DevOrg) then
      begin
        // Devise d'origine en devise pivot
        Taux := GetTaux(DevOrg, DateTaux, DatePiece);
        MontantDestNA := DeviseToEuroNA(MontantOrg, Taux, TOrg.GetValue('D_QUOTITE'));
        MontantDest := DeviseToEuro(MontantOrg, Taux, TOrg.GetValue('D_QUOTITE'));
      end;
      if not ToxTestCodeEuro(DevDest) then
      begin
        // Devise pivot en devise destination
        Taux := GetTaux(DevDest, DateTaux, DatePiece);
        MontantDestNA := EuroToDeviseNA(MontantOrg, Taux, TDest.GetValue('D_QUOTITE'));
        MontantDest := EuroToDevise(MontantOrg, Taux, TDest.GetValue('D_QUOTITE'), TDest.GetValue('D_DECIMALE'));
      end;
    end else
    begin
      if DevOrg <> V_PGI.DevisePivot then
      begin
        // Devise d'origine en devise pivot
        Taux := GetTaux(DevOrg, DateTaux, DatePiece);
        MontantDestNA := DeviseToPivotNA(MontantOrg, Taux, TOrg.GetValue('D_QUOTITE'));
        MontantDest := DeviseToPivot(MontantOrg, Taux, TOrg.GetValue('D_QUOTITE'));
      end;
      if DevDest <> V_PGI.DevisePivot then
      begin
        // Devise pivot en devise destination
        Taux := GetTaux(DevDest, DateTaux, DatePiece);
        MontantDestNA := PivotToDeviseNA(MontantOrg, Taux, TDest.GetValue('D_QUOTITE'));
        MontantDest := PivotToDevise(MontantOrg, Taux, TDest.GetValue('D_QUOTITE'), TDest.GetValue('D_DECIMALE'));
      end;
    end;
  end;
  // Libération des TOB créées.
  if TOBDev = nil then
  begin
    if TOrg <> nil then TOrg.Free;
    if TDEst <> nil then TDest.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. :  Convertit un montant dans la devise dossier et en contre
Suite ........ : valeur
Mots clefs ... : TOX
*****************************************************************}

procedure ToxConvertir(MontantOrigine: double; var MontantDev, MontantContreValeur: double; DevOrg, DevDest: string; Date: TDateTime; TobDev: TOB);
var MontantDevNA: double;
begin
  MontantDev := 0;
  MontantDevNa := 0;
  MontantContreValeur := 0;
  //
  // Conversion du montant dans la devise de tenue du dossier
  //
  ToxConvToDev2(MontantOrigine, MontantDevNA, MontantDev, DevOrg, DevDest, Date, TobDev);
  //
  // Conversion du montant calculé en contre valeur :
  //            1- Si Dossier en EURO, il faut calculer la contre valeur dans l'ancienne devise de tenue du dossier = DeviseFongible
  //            2- Si Dossier pas encore en Euro, conversion du montant en EURO
  //
  if VH^.TenueEuro then
  begin
    if DevDest = V_PGI.DevisePivot then MontantContreValeur := EuroToPivot(MontantDevNA)
    else MontantContreValeur := ToxConvToDev(MontantDevNA, DevDest, V_PGI.DeviseFongible, Date, TobDev);
  end else
  begin
    if DevDest = V_PGI.DevisePivot then MontantContreValeur := PivotToEuro(MontantDevNA)
    else MontantContreValeur := ToxConvToDev(MontantDevNA, DevDest, 'EUR', Date, TobDev);
  end;
end;

end.
