// modif mbo - 16/11/2005 : modification GetExceptionnel pour avoir cumul except+dépréciation actif
// modif mbo - 30/11/2005 : fq 17103 - ajout fonction getexcepFisc pour ne pas prendre en compte
//                          la dépréciation sur exceptionnel fiscal
// TGA 08/12/2005 Ajout GetAncienTaux pour liste des changements
// TGA 21/12/2005 - FQ 17215 -  GetParamSoc => GetParamSocSecur
// TGA 15/02/2006 Ajout Getcumulantuo - Getdotationuo pour Suivi des amortissements varaiables
// BTY 02/06 Anciennes éditions à ouvrir à Web Access
// BTY 03/06 FQ 17583 + Editions ne sortent pas en Web Access ou avec des montants faux
// BTY 03/06 Optimisation en Web Access par usage de la class d'édition TAmortissementContext
// (au lieu de dupliquer les fonctions pour ajouter le paramètre d'appel PlanInfo)
// BTY 03/06 2e optimisation en Web Access par usage de la class TCPContexte car
// chargée au départ par la compta
// TGA 05/07/2006 Ajout GetDPI  pour liste des DPI (IDD)
// TGA 25/07/2006 Ajout GetImmoDPI  pour Suivi des amorts non déductible (ITN)
// MBO 22/08/2006 Init variable Montant pour avertissement compil
// BTY 10/06      Prime d'équipement
// BTY 10/06      Subventions
// MBO 11/06 - modif pour les éditions à date < date fin exercice
// MBO 03/07 - FQ 17512 gestion réintégration extracomptable
// MBO 11/05/2007 - ajout de la fonction GetVReco (calcul de la valeur résiduelle)
// MBO 31/05/2007 - en cas de cession les réint/déduc sont calculées à date de cession
// MBO 26/06/2007 - ajout d'une fonction getbaseprime = getmontantprime/2
// MBO 14/09/2007 - FQ 21422 - ajout f° getDotFiscReel (dotation fiscale réellement pratiquée et non éco+dérog)
// BTY 09/07 FQ 21346 - Amortissements variables optimisé : suppression des appels SQL

// MBO 19/09/2007 - FQ 19913 - ajout de f° pour éditer les dotations simulées à une date <> fin exo en cours
//                           . function ChargeCbAtDate (=chargecontrat avec prise en cpte de la date d'arrêté
//                           . function GetHorsExo qui renvoie string VRAI si date arrêté < date deb exo en cours
//     20/09/2007            . function GetAmortFini qui renvoie string VRAI
//                                                   si l'immo est amortie totalement avant le début exo de la date arrêté
// MBO 19/09/2007 - FQ 21424 - ajout f° GetAntFiscReel = antérieurs fiscaux réellement pratiqués
//                             (et non éco + dérogatoire)
//                           - ajout f° GetPreviFiscReel = prévisionnel fiscal réellement pratiqué
// MBO 05/10/2007 - fq 19172 - nvelle f° GetQuantité (récupère la qté cédée si impres° avt date de cession - ILV)
// MBO 29/10/2007 - fq 21754 - ajout paramètres ds appel du planInfo.Calcul pour calcul cession 
// BTY 13/11/07 Fonctions d'édition de IFR immos issues du paggae forfait à réel (Agricoles)
unit ImEdCalc;

interface

uses SysUtils
      , HEnt1
      //, ImPlan      // BTY 02/06
      , ImPlanInfo
      , HCtrls
      , ImContra
      , ImOuPlan
      , ParamSoc,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
      UTOB
      ;

{$IFDEF EAGLSERVER}
function CalcOLEEtatImmo(PlanInfo: TPlanInfo; sf, sp: string): variant;
{$ELSE}
function CalcOLEEtatImmo(sf, sp: string): variant;
{$ENDIF}
function CalculPValue(PValue: double; DateOpe: TDateTime; MethodeEco: string;
  CumulCessEco, DotCessEco: double; DatePieceA, DateAmort: TDateTime; Compte:
  string): string;
function GetPValueCT(StCalculPValue: string): double;
function GetPValueLT(StCalculPValue: string): double;
function GetMValueCT(StCalculPValue: string): double;
function GetMValueLT(StCalculPValue: string): double;
function GetPreviEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; Indice: integer):
  double;
function GetPrevifisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; Indice: integer):
  double;
function ChargeContrat(CodeContrat: string): string;
//fq 19913 - mbo - charge contrat pour édition à date différente de fin d'exercice
function ChargeCbAtDate(CodeContrat: string; DateCalcul:TDateTime): string;


// BTY 03/06 1e Optimisation
// BTY 02/06 Anciennes éditions à ouvrir à CWAS
//function ChargeContrat2(PlanInfo:TPlanInfo; CodeContrat:string): string;
function GetResteAPayer(StContrat: string): double;
function GetSommeRedevances(StContrat: string; Indice: integer): double;
function GetMontantHT(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
                      TDateTime; EstImmoCedee: string): double;

// Fq 19172 - mbo -05.10.07
function GetQuantite(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
                      TDateTime; EstImmoCedee: string): double;

function GetMontantAchat(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
function GetMontantAchatAnc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
function GetDotationEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
// BTY 10/06 Prime d'équipement
function GetMontantPrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetBasePrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetDureePrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetTauxPrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetMethodePrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): string;
function GetAnterieurPrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
function GetDotationPrime(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
// BTY 10/06 Subventions
function GetMontantSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetDureeSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetTauxSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetMethodeSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): string;
function GetAnterieurSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
function GetDotationSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
function GetDateDebSbv(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): string;
// TGA 08/12/2005
function GetAncienTaux(s1:String;s2:TdateTime;s3:TdateTime;s4:Integer):double;
function GetDepreciation(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime):double;
function GetDotationFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;

  // Ajout mbo - fq 21422 - 14.09.07
function GetDotFiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;

function GetCumulAntEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetCumulAntFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetCumulEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetCumulFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetCumulDerog(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetCumulAntDerog(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetVnc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime;
  EstImmoCedee: string):double;
function GetVncEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime;
  EstImmoCedee:string): double;
function GetVncFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;
function GetCessionEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetCessionFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetCumulCessionEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetCumulCessionFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetValResiduelle(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetExceptionnel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
// ajout mbo FQ 17013 - pour exceptionnel fiscal qui ne prend pas en compte la dépréciation actif
function GetExceptFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetNonDeductible(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetReintegration(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetQuotePart(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
function GetBaseEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
// BTY 11/05 Affichage de la base fiscale
function GetBaseFisc(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetDateDebutEx(indice: integer = 0): TDateTime;
function GetDateFinEx(indice: integer = 0): TDateTime;
function GetMonnaieTenue: string;
function GetBaseTaxePro(indice: integer; BaseTaxePro: double; CodeImmo: string):
  double;
function GetNombreDeMois(DateDebut, DateFin: TDateTime): integer;
function GetCession(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): string;
function GetPVCT(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime):
  double;
function GetPVLT(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime):
  double;
function GetDotationDerog ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime ) : double;
function GetRepriseDerog ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime ) : double;
function GetReintegrationFEC ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime;EstImmoCedee:string ) : double;
function GetDeductionFEC ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime;EstImmoCedee:string ) : double;
function GetTypeReintegration( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul:TDateTime) : string;
function GetCumulAntFEC(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime; EstImmoCedee:string): double;
function GetVReco(PlanInfo:TplanInfo; CodeImmo:string; DateCalcul:TDateTime; EstImmoCedee:string):double;
{$IFDEF EAGLSERVER}
// BTY 02/06 Anciennes éditions à ouvrir à CWAS
//function GetDateDebutExWA(PlanInfo:TPlanInfo): TDateTime;
//function GetDateFinExWA(PlanInfo:TPlanInfo): TDateTime;
{$ELSE}
procedure UpdatePlanInfo(var PlanInfo : TPlanInfo; CodeImmo: string; DateRef: TDateTime);
{$ENDIF}

// TGA 15/02/2006
// FQ 21346
function GetDotationUO(PlanInfo:TPlanInfo; CodeImmo: string; DateCalcul:TDateTime; EstImmoCedee:string): double;
function GetCumulAntUO(PlanInfo:TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
// TGA 06/07/2006
function GetDPI(MtDPI:Double;DateDPI:TDateTime): String;
// TGA 25/07/2006
function GetImmoDPI(CodeImmo:String): Double;
// MBO 13/11/2006
function GetDureeEco(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetTauxEco(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetMethodeEco(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): string;
function GetDureeFisc(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;

// GetTauxFisc : si pas de plan fiscal on prend l'éco
function GetTauxFisc(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;

// TauxFiscReel - le taux fiscal réel
//                (sert pour imprimer mention 'plan fiscal' sur édition avec info complémentaires)
function TauxFiscReel(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): double;
function GetMethodeFisc(PlanInfo:TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): string;

// fq 21424 - mbo - 19/09/07 - calcul des antérieurs fiscaux réellement pratiqués (et non Eco + dérog)
function GetAntFiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:string): double;

// suite FQ 21424 - mbo - 20/09/07 - prévisions fiscales = fiscal réellement pratiqué
function GetPrevifiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; Indice: integer): double;

// fq 19913 - mbo - 19/09/07 - renvoie false si la date d'édition < date debut exercice en cours
function GetHorsExo(DateCalcul:TDateTime): string;
function GetAmortFini(PlanInfo: TPlanInfo; CodeImmo:string; DateCalcul:TDateTime): string;
// BTY 13/11/07 pour IFR
function GetDatePassageIFR(PlanInfo: TPlanInfo; DateCalcul:TDateTime): string;
function GetVNCIFR(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime; EstImmoCedee:string): double;
function GetDureeIFR(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;


implementation


uses

// BTY 03/06 2e optimastion
{$IFDEF EAGLSERVER}
{$IFDEF NOVH}
  ULibCpContexte ,
{$ENDIF}
{$ENDIF}
// BTY 03/06 Class d'édition remontée dans ImEnt
// BTY 03/06 Pour pouvoir récupérer le contexte d'édition sans avoir le paramètre PlanInfo
//{$IFDEF eAGLServer}
//  AmedServer,
//{$ENDIF}
  ImEnt;




{$IFDEF EAGLSERVER}
{$ELSE}
procedure UpdatePlanInfo(var PlanInfo : TPlanInfo; CodeImmo: string; DateRef: TDateTime);
begin
  if (PlanInfo = nil) then
    PlanInfo := TPlanInfo.Create('');
  if (PlanInfo.Plan.CodeImmo <> CodeImmo) or (PlanInfo.DateRef <> DateRef) then
  begin
    PlanInfo.ChargeImmo(CodeImmo);
    PlanInfo.Calcul(DateRef, true, false, '');   //fq 21754 - mbo -
  end;
  VHImmo^.PlanInfo := PlanInfo;
end;
{$ENDIF}

function ReadParam(var St: string): string;
var
  i: Integer;
begin
  i := Pos(';', St);
  if i <= 0 then
    i := Length(St) + 1;
  result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;

function ReadParamIndice(St: string; ind: integer): string;
var
  i: integer;
begin
  for i := 1 to ind do
    result := ReadParam(St);
  if result = '' then
    result := StrfMontant(0.00, 15, V_PGI.OkDecV, '', False);
end;

// Fonction d'interface pour le générateur d'état
{$IFDEF EAGLSERVER}

function CalcOLEEtatImmo(PlanInfo: TPlanInfo; sf, sp: string): variant;
var
{$ELSE}

function CalcOLEEtatImmo(sf, sp: string): variant;
var
  PlanInfo: TPlanInfo;
{$ENDIF}
  s1, s2, s3, s4, s5, s6, s7, s8: string;
  i1, n, n2: Integer;
begin
{$IFDEF EAGLSERVER}
{$ELSE}
  PlanInfo := VHImmo^.PlanInfo;
{$ENDIF}
  if sf = 'CALCULPVALUE' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    if s3 = '' then
      s3 := 'NAM'; //EPZ 06/11/00 cas des immo FI
    s4 := ReadParam(sp);
    s5 := ReadParam(sp);
    s6 := ReadParam(sp);
    s7 := ReadParam(sp);
    s8 := ReadParam(sp);
    result := CalculPValue(StrToFloat(s1), StrToFloat(s2), s3, StrToFloat(s4),
      StrToFloat(s5), StrToFloat(s6), StrToFloat(s7), s8);
  end
  else if sf = 'GETPVALUECT' then
  begin
    result := GetPValueCT(sp);
  end
  else if sf = 'GETPVALUELT' then
  begin
    result := GetPValueLT(sp);
  end
  else if sf = 'GETMVALUECT' then
  begin
    result := GetMValueCT(sp);
  end
  else if sf = 'GETMVALUELT' then
  begin
    result := GetMValueLT(sp);
  end
  else if sf = 'GETPREVIECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    Result := GetPreviEco(PlanInfo, s1, StrToFloat(s2), StrToInt(s3));
  end
  else if sf = 'GETPREVIFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    Result := GetPreviFisc(PlanInfo, s1, StrToFloat(s2), StrToInt(s3));
  end
  // fq 21424
  else if sf = 'GETPREVIFISCREEL' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    Result := GetPreviFiscReel(PlanInfo, s1, StrToFloat(s2), StrToInt(s3));
  end
  else if sf = 'CHARGECONTRAT' then
  begin
    s1 := ReadParam(sp);
    Result := ChargeContrat(s1);
  end
  //fq 19913 - mbo
  else if sf = 'CHARGECBATDATE' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := ChargeCbAtDate(s1, StrToFloat(s2));
  end

  // BTY 03/06
  // BTY 02/06 <=> ChargeContrat avec compatibilité CWAS
  //else if sf = 'CHARGECONTRAT2' then
  //begin
  //  s1 := ReadParam(sp);
  //  s2 := ReadParam(sp);
  //  Result := ChargeContrat2(PlanInfo, s1);
  //end
  else if sf = 'GETRESTEAPAYER' then
  begin
    s1 := ReadParam(sp);
    result := GetResteAPayer(s1);
    exit;
  end
  else if sf = 'GETSOMMEREDEVANCES' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    s4 := ReadParam(sp);
    s5 := ReadParam(sp);
    s6 := ReadParam(sp);
    s7 := ReadParam(sp);
    s8 := ReadParam(sp);
    result := GetSommeRedevances(s2 + ';' + s3 + ';' + s4 + ';' + s5 + ';' + s6
      +
      ';' + s7 + ';', StrToInt(s8));
  end
  else if sf = 'GETMONTANTHT' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getMontantht(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETQUANTITE' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getQuantite(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETMONTANTACHAT' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getMontantachat(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETMONTANTACHATANC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getMontantachatAnc(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDOTATIONECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getDotationEco(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDEPRECIATION' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := getDepreciation(PlanInfo, s1,StrToFloat(S2) );
  end
  else if sf = 'GETDOTATIONFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetDotationFisc(PlanInfo, s1, StrToFloat(S2), s3);
  end
  // fq 21422
  else if sf = 'GETDOTFISCREEL' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetDotFiscReel(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETCUMULANTECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulAntEco(PlanInfo, s1, StrToFloat(S2), s3);
  end
  // MBO FQ 17512 Chantier fiscal
  else if sf = 'GETCUMULANTFEC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulAntFEC(PlanInfo, s1, StrToFloat(S2), s3);
  end
  // MBO Remplacement de composant
  else if sf = 'GETVRECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetVReco(PlanInfo, s1, StrToFloat(S2), s3);
  end
  // TGA 15/02/2006
  else if sf = 'GETDOTATIONUO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getDotationUO(PlanInfo, s1, StrToFloat(S2), s3); // FQ 21346
  end
  else if sf = 'GETCUMULANTUO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetCumulAntUO(PlanInfo, s1, StrToFloat(S2)); // FQ 21346
  end
  // TGA 06/07/2006
  else if sf = 'GETDPI' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetDPI(StrToFloat(s1), StrToFloat(S2));
  end
  // TGA 25/07/2006
  else if sf = 'GETIMMODPI' then
  begin
    s1 := ReadParam(sp);
    result := GetImmoDPI(s1);
  end
  else if sf = 'GETCUMULANTFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulAntFisc(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETANTFISCREEL' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetAntFiscReel(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETCUMULECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulEco(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETCUMULFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulFisc(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETCUMULDEROG' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulDerog(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETCUMULANTDEROG' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetCumulAntDerog(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDOTATIONDEROG' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetDotationDerog(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETREPRISEDEROG' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetRepriseDerog(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETREINTEGRATIONFEC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetReintegrationFec(PlanInfo, s1, StrToFloat(S2),s3);
  end
  else if sf = 'GETDEDUCTIONFEC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetDeductionFEC(PlanInfo, s1, StrToFloat(S2),s3);
  end
  else if sf = 'GETTYPEREINTEGRATION' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetTypeReintegration(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETVNC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetVnc(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETVNCECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetVncEco(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETVNCFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetVncFisc(PlanInfo, s1, StrToFloat(S2), S3);
  end
  else if sf = 'GETCESSION' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetCession(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETCESSIONECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetCessionEco(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETCESSIONFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetCessionFisc(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETCUMULCESSIONECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetCumulCessionEco(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETCUMULCESSIONFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetCumulCessionFisc(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETVALRESIDUELLE' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetValResiduelle(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETEXCEPTIONNEL' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetExceptionnel(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETEXCEPTFISC' then    // fq 17013 ajout mbo 30/11/2005
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetExceptFisc(PlanInfo, s1, StrToFloat(S2));
  end
  // TGA 08/12/2005 pour liste des changements
  else if sf = 'GETANCIENTAUX' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    s4 := ReadParam(sp);
    result := GetAncienTaux(s1,StrToFloat(s2),StrToFloat(s3),StrToInt(s4));
  end
  else if sf = 'GETNONDEDUCTIBLE' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetNonDeductible(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETREINTEGRATION' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetReintegration(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETQUOTEPART' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetQuotePart(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETBASEECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetBaseEco(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETBASEFISC' then   // BTY 11/05 Base fiscale à prévoir
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetBaseFisc(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETPVCT' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetPVCT(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETPVLT' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetPVLT(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETDATEDEBUTEX' then
  begin
    try
      i1 := StrToInt(ReadParam(sp));
    except i1 := 0;
    end;
    result := GetDateDebutEx(i1);
  end
  else if sf = 'GETDATEFINEX' then
  begin
    try
      i1 := StrToInt(ReadParam(sp));
    except i1 := 0;
    end;
    result := GetDateFinEx(i1);
  end
  else if sf = 'GETMONNAIETENUE' then
  begin
    result := GetMonnaieTenue;
  end
  else if sf = 'GETBASETAXEPRO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetBaseTaxePro(StrToInt(s1), StrToFloat(s2), s3);
  end
  else if sf = 'GETNOMBREDEMOIS' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetNombreDeMois(StrToFloat(s1), StrToFloat(s2));
  end
  // BTY 10/06 Prime d'équipement
  else if sf = 'GETMONTANTPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetMontantPrime(PlanInfo, s1, StrToFloat(S2));
  end
  // MBO 26.06.07 la base prime = montant prime /2
  else if sf = 'GETBASEPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetBasePrime(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETDUREEPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetDureePrime(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETTAUXPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetTauxPrime(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETMETHODEPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetMethodePrime(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETANTERIEURPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetAnterieurPrime(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDOTATIONPRIME' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetDotationPrime(PlanInfo, s1, StrToFloat(S2), s3);
  end
  // BTY 10/06 Subvention
  else if sf = 'GETMONTANTSBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetMontantSbv(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETDUREESBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetDureeSbv(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETTAUXSBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetTauxSbv(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETMETHODESBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetMethodeSbv(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETANTERIEURSBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetAnterieurSbv(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDOTATIONSBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := GetDotationSbv(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDATEDEBSBV' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetDateDebSbv(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETDUREEECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetDureeEco(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETTAUXECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetTauxEco(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETMETHODEECO' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetMethodeEco(PlanInfo, s1, StrToFloat(S2));
  end
   else if sf = 'GETDUREEFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetDureeFisc(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETTAUXFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetTauxFisc(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'TAUXFISCREEL' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := TauxFiscReel(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'GETMETHODEFISC' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetMethodeFisc(PlanInfo, s1, StrToFloat(S2));
  end
  //fq 19913 -
  else if sf = 'GETHORSEXO' then
  begin
    s1 := ReadParam(sp);
    result := GetHorsExo(StrToFloat(S1));
  end
  else if sf = 'GETAMORTFINI' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := GetAmortFini(PlanInfo,s1, StrToFloat(S2));
  end
  else if sf = 'GETDATEPASSAGE' then
  begin
    s1 := ReadParam(sp);
    result := GetDatePassageIFR(PlanInfo, StrToFloat(S1));
  end
  else if sf = 'GETVNCIFR' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    result := getVNCIFR(PlanInfo, s1, StrToFloat(S2), s3);
  end
  else if sf = 'GETDUREEIFR' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    result := getDureeIFR(PlanInfo, s1, StrToFloat(S2));
  end
  else if sf = 'MODULO' then
  begin
    result := 0;
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    try
    n := StrToInt(s1)except n := 0;
    end;
    try
    n2 := StrToInt(s2)except n2 := 0;
    end;
    if (n <> 0) and (n2 <> 0) then
      result := IntToStr(n mod n2);
{$IFNDEF EAGLCLIENT}
  end
  else if sf = 'INFODOSSIERPIED' then
  begin
{$IFNDEF EAGLSERVER}
    if (ctxPCL in V_PGI.PGIContexte) then
      Result := V_PGI.DefaultSectionName + ' / ' + TraduireMemoire('Dossier n°') + V_PGI.NoDossier
    else Result := GetParamSocSecur ('SO_LIBELLE','');
{$ELSE}
      Result := '';
{$ENDIF}
{$ENDIF}
  end;
end;

function CalculPValue(PValue: double; DateOpe: TDateTime; MethodeEco: string;
  CumulCessEco, DotCessEco: double; DatePieceA, DateAmort: TDateTime; Compte:
  string): string;
var
  CumDot, PValueCT, PValueLT, MValueCT, MValueLT: double;
  bCourtTerme, bOkIS: boolean;
begin
  CumDot := CumulCessEco + DotCessEco;
  PValueCT := 0;
  PValueLT := 0;
  MValueCT := 0;
  MValueLT := 0;
  bOkIS := (copy(Compte, 1, 3) = '205') or (copy(Compte, 1, 3) = '261');
  // Impot société
  if MethodeEco = 'NAM' then
  begin
    bCourtTerme := (DateOpe < PlusMois(DatePieceA, 24)) or not bOkIs;
    if bCourtTerme then
    begin
      if PValue > 0 then
        PValueCT := PValue
      else
        MValueCT := (-1) * PValue;
    end
    else
    begin
      if PValue > 0 then
        PValueLT := PValue
      else
        MValueLT := (-1) * PValue;
    end;
  end
  else
  begin
    if MethodeEco = 'LIN' then
      bCourtTerme := (DateOpe < PlusMois(DateAmort, 24)) or not bOkIs
    else
      bCourtTerme := (DateOpe < PlusMois(DatePieceA, 24)) or not bOkIs;
    if bCourtTerme then
    begin
      if PValue > 0 then
        PValueCT := PValue
      else
        MValueCT := (-1) * PValue;
    end
    else
    begin
      if PValue < 0 then
        MValueCT := (-1) * PValue
      else
      begin
        if PValue > CumDot then
        begin
          PValueCT := CumDot;
          PValueLT := PValue - CumDot;
        end
        else
          PValueCT := PValue;
      end;
    end;
  end;
  result := StrfMontant(PValueCT, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(PValueLT, 15, V_PGI.OkDecV, '', False) +
    ';' + StrfMontant(MValueCT, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(MValueLT, 15, V_PGI.OkDecV, '', False);
end;

function GetPValueCT(StCalculPValue: string): double;
begin
  result := StrToFloat(ReadParamIndice(StCalculPValue, 1));
end;

function GetPValueLT(StCalculPValue: string): double;
begin
  result := StrToFloat(ReadParamIndice(StCalculPValue, 2));
end;

function GetMValueCT(StCalculPValue: string): double;
begin
  result := StrToFloat(ReadParamIndice(StCalculPValue, 3));
end;

function GetMValueLT(StCalculPValue: string): double;
begin
  result := StrToFloat(ReadParamIndice(StCalculPValue, 4));
end;

function GetPreviEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; Indice: integer): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.GetPreviEco(Indice);

end;

//------------------------------------------------------------------------------

function GetPrevifisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; Indice: integer):
  double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.GetPreviFisc(Indice);
end;

//------------------------------------------------------------------------------
//fq 21424 - 20/09/2007 - mbo
function GetPrevifiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; Indice: integer):
  double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.GetPreviFiscalReel(Indice);
end;
//------------------------------------------------------------------------------

function ChargeContrat(CodeContrat: string): string;
var
  Contrat: TImContrat;
  D1, D2, D3, D4: TDateTime;
  dbResteAPayer, dbEngN, dbEngCumuls, dbEngTotal, dbEngN1, dbEngN1a5,
    dbEngNSup5: double;
begin
  Contrat := TImContrat.Create;
  Contrat.ChargeFromCode(CodeContrat);
  Contrat.ChargeTableEcheance;
  dbResteAPayer := Contrat.GetResteAPayer(GetDateDebutEx);
  dbEngN := Contrat.GetSommeDesLoyers(GetDateDebutEx, GetDateFinEx);
  dbEngCumuls := Contrat.GetSommeDesLoyers(iDate1900, GetDateFinEx);
  dbEngTotal := Contrat.GetSommeDesLoyers(iDate1900, iDate2099);
  CalculDateDebutFinExByPos(1, D1, D2); //calcul dates exe suivant
  dbEngN1 := Contrat.GetSommeDesLoyers(D1, D2);
  CalculDateDebutFinExByPos(5, D3, D4); //calcul dates exe +5ans
  dbEngN1a5 := Contrat.GetSommeDesLoyers(D2 + 1, D4 - 1);
  dbEngNSup5 := Contrat.GetSommeDesLoyers(D4 + 1, iDate2099);
  result := StrfMontant(dbResteAPayer, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngN, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngCumuls, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngN1, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngN1a5, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngNSup5, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngTotal, 15, V_PGI.OkDecV, '', False);
end;

//fq 19913 - mbo
function ChargeCbAtDate(CodeContrat: string; DateCalcul:TDateTime): string;
var
  Contrat: TImContrat;
  D1, D2, D3, D4: TDateTime;
  dbResteAPayer, dbEngN, dbEngCumuls, dbEngTotal, dbEngN1, dbEngN1a5,
  dbEngNSup5: double;
  DebExoCalcul, FinExoCalcul : TDateTime;
begin
  DebExoCalcul := iDate1900;
  FinExoCalcul := iDate1900;
  Contrat := TImContrat.Create;
  Contrat.ChargeFromCode(CodeContrat);
  Contrat.ChargeEchSuivantDate(DateCalcul,('I_TYPELOYERCB'='AVA'));
  GetDatesExercice(DateCalcul,DebExoCalcul,FinExoCalcul);
  dbResteAPayer := Contrat.GetResteAPayer(DebExoCalcul);
  dbEngN := Contrat.GetSommeDesLoyers(DebExoCalcul, FinExoCalcul);
  dbEngCumuls := Contrat.GetSommeDesLoyers(iDate1900, FinExoCalcul);
  dbEngTotal := Contrat.GetSommeDesLoyers(iDate1900, iDate2099);
  CalculDateDebutFinExByPos(1, D1, D2); //calcul dates exe suivant
  dbEngN1 := Contrat.GetSommeDesLoyers(D1, D2);
  CalculDateDebutFinExByPos(5, D3, D4); //calcul dates exe +5ans
  dbEngN1a5 := Contrat.GetSommeDesLoyers(D2 + 1, D4 - 1);
  dbEngNSup5 := Contrat.GetSommeDesLoyers(D4 + 1, iDate2099);
  result := StrfMontant(dbResteAPayer, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngN, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngCumuls, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngN1, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngN1a5, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngNSup5, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngTotal, 15, V_PGI.OkDecV, '', False);
end;

//--------------------------------------------------------------------------------
//fq 19913 - renvoie VRAI si la date d'édition n'est pas sur l'exercice en cours
function GetHorsExo(DateCalcul:TDateTime):string;
begin
{$IFDEF EAGLSERVER}
  if DateCalcul < TCPContexte.GetCurrent.Exercice.Encours.Deb then
     result := 'VRAI'
  else
     result := 'FAUX';
{$ELSE}
  if DateCalcul < VHImmo^.EnCours.Deb then
     result := 'VRAI'
  else
     result := 'FAUX';
{$ENDIF}
end;
//--------------------------------------------------------------------------------
//fq 19913 - renvoie VRAI si fini d'amortir avant date deb de l'exo auquel appartient la date d'arrêté
function GetAmortFini(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime):string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}

  if PlanInfo.Plan.AmortEco.DateFinAmort < PlanInfo.DateDebutDotation(DateCalcul) then
     result:= 'VRAI'
  else
     result:= 'FAUX';

end;

//--------------------------------------------------------------------------------

// BTY 13/11/07 Fonctions pour l'édition IFR des immos issues du forfait
function GetDatePassageIFR(PlanInfo:TPlanInfo; DateCalcul:TDateTime):string;
begin
  result := GetParamSocSecur('SO_DATEREGREEL','');
end;

// Récup i_baseeco, 0 si cédée
function GetVNCIFR(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.GetVNCIFR(DateCalcul, EstImmoCedee = 'CES');
end;

// Récup i_dureeeco, 0 si cédée
function GetDureeIFR(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.GetDureeIFR(DateCalcul);
end;

// -------------------------------------------------------------------------------



// BTY 02/06 <=> ChargeContrat compatible CWAS
{function ChargeContrat2(PlanInfo: TPlanInfo; CodeContrat: string): string;
var
  Contrat: TImContrat;
  D1, D2, D3, D4: TDateTime;
  dbResteAPayer, dbEngN, dbEngCumuls, dbEngTotal, dbEngN1, dbEngN1a5,
    dbEngNSup5: double;
begin
  Contrat := TImContrat.Create;
//{$IFDEF EAGLSERVER}
//  Contrat.ChargeFromCodeWA(PlanInfo, CodeContrat); }
//{$ELSE}
//  Contrat.ChargeFromCode(CodeContrat);
//{$ENDIF}

//  Contrat.ChargeTableEcheance;

//{$IFDEF EAGLSERVER}
{  dbResteAPayer := Contrat.GetResteAPayer(GetDateDebutExWA(PlanInfo));
  dbEngN := Contrat.GetSommeDesLoyers(GetDateDebutExWA(PlanInfo),
                                      GetDateFinExWA(PlanInfo));
  dbEngCumuls := Contrat.GetSommeDesLoyers(iDate1900, GetDateFinExWA(PlanInfo));
  dbEngTotal := Contrat.GetSommeDesLoyers(iDate1900, iDate2099);
  // BTY 03/06
  //  CalculDateDebutFinExByPosWA(PlanInfo, 1, D1, D2, GetDateDebutExWA(PlanInfo)); //calcul dates exe suivant
  CalculDateDebutFinExByPos(1, D1, D2); //calcul dates exe suivant  dbEngN1 := Contrat.GetSommeDesLoyers(D1, D2);
  dbEngN1 := Contrat.GetSommeDesLoyers(D1, D2);
  //  CalculDateDebutFinExByPosWA(PlanInfo, 5, D3, D4, GetDateDebutExWA(PlanInfo)); //calcul dates exe +5ans
  CalculDateDebutFinExByPos(5, D3, D4); //calcul dates exe +5ans }
//{$ELSE}
{  dbResteAPayer := Contrat.GetResteAPayer(GetDateDebutEx);
  dbEngN := Contrat.GetSommeDesLoyers(GetDateDebutEx,GetDateFinEx);
  dbEngCumuls := Contrat.GetSommeDesLoyers(iDate1900, GetDateFinEx);
  dbEngTotal := Contrat.GetSommeDesLoyers(iDate1900, iDate2099);
  CalculDateDebutFinExByPos(1, D1, D2); //calcul dates exe suivant
  dbEngN1 := Contrat.GetSommeDesLoyers(D1, D2);
  CalculDateDebutFinExByPos(5, D3, D4); //calcul dates exe +5ans }
//{$ENDIF}
{  dbEngN1a5 := Contrat.GetSommeDesLoyers(D2 + 1, D4 - 1);
  dbEngNSup5 := Contrat.GetSommeDesLoyers(D4 + 1, iDate2099);
  result := StrfMontant(dbResteAPayer, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngN, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngCumuls, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngN1, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngN1a5, 15, V_PGI.OkDecV, '', False) + ';' +
    StrfMontant(dbEngNSup5, 15, V_PGI.OkDecV, '', False)
    + ';' + StrfMontant(dbEngTotal, 15, V_PGI.OkDecV, '', False);
end; }

function GetResteAPayer(StContrat: string): double;
begin
  result := StrToFLoat(ReadParam(StContrat));
end;

function GetSommeRedevances(StContrat: string; Indice: integer): double;
begin
  result := StrToFLoat(ReadParamIndice(StContrat, Indice));
end;

function GetMontantHT(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  //modif mbo 02.07.07 si l'immo remplaçante est elle-même remplacée il faut en tenir cpte
  // car ValeurHT = 0 si l'immo est remplacée (mais pas la valeur d'achat)
  //if PlanInfo.Plan.Remplace then
  //   result := PlanInfo.Plan.ValeurAchat
  //else
     result := PlanInfo.Plan.ValeurHT; //+VHImmo^.PlanInfo.CessionEco ;

  // fq 19172 if EstImmoCedee = 'CES' then
  if PlanInfo.Plan.CedeLe <> idate1900 then
     Result := PlanInfo.GetVoCedee(iDate1900);
end;

//-----------------------------------------------------------------------------

// nvelle f° pour FQ 19172 - liste des valorisations
// le but = récupérer la quantité si impression avant date de cession
function GetQuantite(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.Plan.Qte;

  if PlanInfo.Plan.CedeLe <> idate1900 then
     Result := PlanInfo.GetQteCedee(iDate1900);
end;

//--------------------------------------------------------------------------------

function GetMontantAchat(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.GetValeurAchat(DateCalcul, EstImmoCedee = 'CES');

end;

// BTY 10/06 Traitement prime d'équipement
function GetMontantPrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.MNTPrime;
end;

// mbo 26.06.07 la base prime = montant prime /2
function GetBasePrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := (PlanInfo.Plan.MNTPrime)/2;
end;

// Récup durée d'amortissement ou d'inaliénabilité
function GetDureePrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.PRIDuree;;
end;

// Récup taux d'amortissement
function GetTauxPrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.PRITaux;;
end;

// Récup méthode d'amortissement
function GetMethodePrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.PRIMethode;;
end;

// Récup suramortissement pratiqué au début de l'exercice de la date d'arrêté
function GetAnterieurPrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result:=PlanInfo.CumulAntPri ;
end;

// Récup suramortissement de l'exercice de la date d'arrêté
function GetDotationPrime(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
 Result := PlanInfo.DotationPri;
 if EstImmoCedee = 'CES' then
    Result := PlanInfo.CessionPri;
end;
//
// BTY 10/06 Traitement subventions
function GetMontantSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.MNTSbv;
end;

// Récup durée d'échelonnement ou inaliénabilité
function GetDureeSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.SBVDuree;;
end;

// Récup taux d'amortissement
function GetTauxSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.SBVTaux;;
end;

// Récup méthode d'amortissement
function GetMethodeSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.SBVMethode;;
end;

// Récup reprise de la subvention pratiquée au début de l'exercice de la date d'arrêté
function GetAnterieurSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result:=PlanInfo.CumulAntSbv ;
end;

// Récup reprise de la subvention de l'exercice de la date d'arrêté
function GetDotationSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime;
         EstImmoCedee:string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
 Result := PlanInfo.DotationSbv;
 if EstImmoCedee = 'CES' then
    Result := PlanInfo.CessionSbv;
end;

// Récup date de début d'amortissement
function GetDateDebSbv(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.SBVDateDeb  <> iDate1900 then
     Result := DateToStr (PlanInfo.Plan.SBVDateDeb);
end;
//

function GetDotationEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
//var F:TextFile;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.DotationEco;
  if EstImmoCedee = 'CES' then
//    result := GetCessionEco(PlanInfo, CodeImmo, DateCalcul) + PlanInfo.ExcepEco;
  result := GetCessionEco(PlanInfo, CodeImmo, DateCalcul);  // FQ 12370

// BTY***  Trace dans un fichier
{AssignFile (F,   'c:\imedcalc.log');
if   FileExists ('c:\imedcalc.log') then  Append (F)
else Rewrite (F);
Writeln (F, TimeToStr(time) + ' GetDotationEco:immo= ' + CodeImmo + 'estcede= ' + EstImmoCedee +
 ' datecalcul= ' + datetostr(datecalcul) + ' planinfo.doteco= ' + floattostr(PlanInfo.DotationEco) );
CloseFile (F); }

end;

function GetAncienTaux(s1:String;s2:TdateTime;s3:TdateTime;s4:Integer):double;
begin
  result := GetTaux(s1,s2,s3,s4);
end;

function GetMontantAchatAnc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee: string): double;
var Q:TQuery;
    debex : TdateTime;
begin

   // Ancien montant si l'immo est issue d'un éclatement (I_IMMOORIGINEECL<>"")
   // et enreg immolog type ACQ dans l'exercice avec IL_MONTANTECL<>"0"

   if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+codeimmo+ '"'+
                 'AND I_IMMOORIGINEECL<>"" ') then
    Begin
      debex := GetDateDebutEx(0);
      Q := OpenSQL('SELECT IL_MONTANTECL FROM IMMOLOG WHERE IL_IMMO="' +
                   codeimmo + '" AND IL_MONTANTECL<>"0" AND IL_TYPEOP = "ACQ" '+
                   'AND IL_DATEOPREELLE>= "'+USDateTime(debex)+'"',false);

      if not Q.Eof then
        Begin
          Result:=Q.FindField('IL_MONTANTECL').AsFloat ;
          ferme(Q);
          exit;
        End;
      ferme(Q);
    End;

    {$IFDEF EAGLSERVER}
    PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
    {$ELSE}
    UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
    {$ENDIF}
    Result := PlanInfo.GetValeurAchat(DateCalcul, EstImmoCedee = 'CES');
end;

function GetDepreciation(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.GetDotDeprec(DateCalcul);
end;


function GetDotationFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
  begin
    // result := PlanInfo.DotationFisc;
    result := PlanInfo.DotationPratiquee;
    if EstImmoCedee = 'CES' then
      result := GetCessionFisc(PlanInfo, CodeImmo, DateCalcul);
  end
  else
  begin
    result := PlanInfo.DotationEco;
    if EstImmoCedee = 'CES' then
      result := GetCessionEco(PlanInfo, CodeImmo, DateCalcul);
  end;
end;


// FQ 21422 - mbo - 14.09.07
function GetDotFiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
  begin
    result := PlanInfo.DotationFisc;
    if EstImmoCedee = 'CES' then
      result := PlanInfo.CessionFisc;
  end
  else
  begin
    result := PlanInfo.DotationEco;
    if EstImmoCedee = 'CES' then
      result := PlanInfo.CessionEco;
  end;
end;


function GetCumulAntEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
//var F:TextFile;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result:=PlanInfo.CumulAntEco ;
  if EstImmoCedee='CES' then
    // 03/06 BTY Edition liste simplifiée avec immos cédées ne sort pas en Web Access
    //Result := VHImmo^.PlanInfo.GetCumulAntEco(DateCalcul);
    Result := PlanInfo.GetCumulAntEco(DateCalcul);

// BTY***  Trace dans un fichier
{AssignFile (F,   'c:\imedcalc.log');
if   FileExists ('c:\imedcalc.log') then  Append (F)
else Rewrite (F);
Writeln (F, TimeToStr(time) + ' GetCUMULANTEco:immo= ' + CodeImmo + 'estcede= ' + EstImmoCedee +
 ' datecalcul= ' + datetostr(datecalcul) + ' planinfo.cumulanteco= ' + floattostr(PlanInfo.CumulAntEco) +
 ' result= ' + floattostr(result) );
CloseFile (F); }

end;

// TGA 06/07/2006
function GetDPI(MtDPI:Double;DateDPI:TDateTime): String;
var Q:TQuery ;
    retour:String;
begin
  retour := 'Fraction';

  Q := OpenSQL('SELECT IZ_MONTANT FROM IMMOMVTD WHERE (IZ_DATE="'+USDateTime(DateDPI)+'" AND IZ_NATURED="DPI")',True);
  IF (not Q.Eof) AND (Q.FindField('IZ_MONTANT').AsFloat = MtDPI) THEN
     retour := 'Total';

  Ferme(Q);
  result:=Retour;
end;

// TGA 26/07/2006
function GetImmoDPI(CodeImmo:String): Double;
var
  Q : TQuery ;
  Montant : Double;
begin
  Montant := 0;  // avertissement en compil mbo 22.08.06

  Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_IMMO="'+CodeImmo+'"', FALSE) ;
  try
    While Not Q.Eof do
      begin
        Montant := Montant + Q.FindField('IZ_MONTANT').AsFloat;
        Q.Next ;
      End;
  finally
    Ferme(Q);
  End;
  result:=Montant;
End;


// TGA 15/02/2006
function GetCumulAntUO(PlanInfo: TPlanInfo; CodeImmo:string;DateCalcul:TDateTime): double;
var //Q:TQuery ;
    Cumuluo: double;
    Debex:TdateTime;
    i: integer;
begin
  Cumuluo:=0;
  debex := GetDateDebutEx(0);

{ FQ 21346 Eviter l'appel SQL tout est déjà dans PlanInfo.Plan.fPlanUO
  Q := OpenSQL('SELECT IUO_UNITEOEUVRE FROM IMMOUO WHERE (IUO_IMMO="'
       + CodeImmo +'" AND IUO_DATE<"' + USDateTime(DebEx)+'")',True);

  while not Q.Eof do
     begin
        Cumuluo:=  Cumuluo + Q.FindField('IUO_UNITEOEUVRE').AsFloat;
        Q.Next;
     end;
  Ferme(Q); }

{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}

  for i := 0 to PlanInfo.Plan.PlanUO.Detail.Count - 1 do
  begin
    if (PlanInfo.Plan.PlanUO.Detail[i].GetValue('IUO_DATE') < debex) then
       CumulUo := cumulUO + PlanInfo.Plan.PlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE');
  end;

  result:=Cumuluo;
end;


function GetDotationUO(PlanInfo: TPlanInfo; CodeImmo:string;DateCalcul:TDateTime;EstImmoCedee:string):double;
var //Q:TQuery ;
    Nbuo: double;
    PremMois,PremAnnee,NbMois1,NbMois2 : Word;
    DateCession,Finex,Debex:TdateTime;
    i: integer;
    Acquise : boolean;
begin

  Nbuo:=0;
  debex := GetDateDebutEx(0);
  finex := GetDateFInEx(0);

  { FQ 21346 Aucun appel SQL tout est déjà dans PlanInfo.Plan et PlanInfo.Plan.PlanUO

  // Lecture du Nb d'unités d'oeuvre de l'exercice
  Q := OpenSQL('SELECT IUO_UNITEOEUVRE FROM IMMOUO WHERE (IUO_IMMO="'
      + CodeImmo +'" AND IUO_DATE="' + USDateTime(FinEx)+'")',True);

  IF not Q.Eof Then
     Begin

       // nombre unités d'oeuvre de l'exercice
       NbUo:=Q.FindField('IUO_UNITEOEUVRE').AsFloat;
       Ferme(Q);

       // Recherche si acquisition dans l'exercice
       Q := OpenSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE (IL_IMMO="'
            + CodeImmo +'" AND IL_DATEOP>="' + USDateTime(DebEx)
            +'" AND IL_TYPEOP="ACQ" )',True);

       IF not Q.Eof Then
          Begin
             Ferme(Q);
            // C'est une acquisition recherche de la date de début d'amortissement
            Q := OpenSQL('SELECT I_DATEDEBECO FROM IMMO WHERE I_IMMO="'+ CodeImmo +'"',True);
            IF not Q.Eof Then
               Debex:=Q.FindField('I_DATEDEBECO').AsFloat;
          End;
       Ferme(Q);

       // Nombre de mois de la date acquisition ou date début ex à la date fin exercice
       NOMBREMOIS(Debex,FinEx,PremMois,PremAnnee,NbMois1);

       // Test si Immo cédée
       IF EstImmoCedee<>'' Then

         Begin

           // Recherche Date de Cession
           Q := OpenSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE (IL_IMMO="'
             + CodeImmo +'" AND IL_DATEOP>="' + USDateTime(DebEx)
             +'" AND IL_TYPEOP="CES" )',True);

           IF not Q.Eof Then
              Begin
                DateCession :=Q.FindField('IL_DATEOP').AsFloat;

                // Test si date de cession < Date de calcul
                IF USDateTime(DateCession)< USDateTime(DateCalcul) Then
                  // Nombre de mois de la date acquisition/début ex à date de cession
                  NOMBREMOIS(Debex,DateCession,PremMois,PremAnnee,NbMois2)
                Else
                  // Nombre de mois de la date acquisition/début ex à date de calcul
                  NOMBREMOIS(Debex,DateCalcul,PremMois,PremAnnee,NbMois2);

                // Prorata à date de cession ou calcul
                NbUo := Arrondi((NbUo*NbMois2)/Nbmois1, V_PGI.OkDecV);

              End;
           Ferme(Q);
         End
       Else
         Begin
           // TGA Fq 18357
           // ne faire un prorata que si la date d'arrêté <> date de fin d'exercice
           // car si date d'acquisition <> date début exo le nombre uo est déjà proratisé

           IF (USDateTime(Datecalcul)<USDateTime(FinEx)) Then
             Begin
               // Nombre de mois de la date acquisition à la date d'arrété
               NOMBREMOIS(Debex,DateCalcul,PremMois,PremAnnee,NbMois2);
               // Prorata à la date d'arrêté
               NbUo := Arrondi((NbUo*NbMois2)/Nbmois1, V_PGI.OkDecV);
             End;
         End;
     End; }


{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}

  // nombre unités d'oeuvre de l'exercice
  for i := 0 to PlanInfo.Plan.PlanUO.Detail.Count - 1 do
  begin
    if (PlanInfo.Plan.PlanUO.Detail[i].GetValue('IUO_DATE') = FinEx) then
       NbUO := NbUO + PlanInfo.Plan.PlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE');
  end;

  Acquise := (PlanInfo.Plan.DateAchat >= DebEx);
  if Acquise then
     debex := PlanInfo.Plan.DateDebEco;

  // Nombre de mois de la date acquisition ou date début ex à la date fin exercice
  NOMBREMOIS(Debex,FinEx,PremMois,PremAnnee,NbMois1);

  // Test si Immo cédée
  IF EstImmoCedee<>'' Then
  begin
     DateCession := PlanInfo.Plan.CedeLe;

     // Test si date de cession < Date de calcul
     IF USDateTime(DateCession)< USDateTime(DateCalcul) Then
        // Nombre de mois de la date acquisition/début ex à date de cession
        NOMBREMOIS(Debex,DateCession,PremMois,PremAnnee,NbMois2)
     Else
        // Nombre de mois de la date acquisition/début ex à date de calcul
        NOMBREMOIS(Debex,DateCalcul,PremMois,PremAnnee,NbMois2);

     // Prorata à date de cession ou calcul
     NbUo := Arrondi((NbUo*NbMois2)/Nbmois1, V_PGI.OkDecV);
  end else
  begin
           // TGA Fq 18357
           // ne faire un prorata que si la date d'arrêté <> date de fin d'exercice
           // car si date d'acquisition <> date début exo le nombre uo est déjà proratisé

           IF (USDateTime(Datecalcul)<USDateTime(FinEx)) Then
             Begin
               // Nombre de mois de la date acquisition à la date d'arrété
               NOMBREMOIS(Debex,DateCalcul,PremMois,PremAnnee,NbMois2);
               // Prorata à la date d'arrêté
               NbUo := Arrondi((NbUo*NbMois2)/Nbmois1, V_PGI.OkDecV);
             End;
  end;

  result:=Nbuo;
end;



function GetCumulAntFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
  begin
    // result:=VHImmo^.PlanInfo.CumulAntFisc ;
    // 03/06 BTY Edition Prévisionnelle Fiscale avec immos cédées ne sort pas en Web Access
    // result := VHImmo^.PlanInfo.CumulAntPratique ;
    result := PlanInfo.CumulAntPratique ;
    if EstImmoCedee='CES' then Result := PlanInfo.GetCumulAntFisc(DateCalcul);

  end
  else Result := GetCumulAntEco(PlanInfo,CodeImmo, DateCalcul, EstImmoCedee);
end;
//------------------------------------------------------------------------
// fq 21424 - 19/09/2007 - mbo
function GetAntFiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
     Result := PlanInfo.GetCumulAntFisc(DateCalcul)
 else
     Result := GetCumulAntEco(PlanInfo,CodeImmo, DateCalcul, EstImmoCedee);
end;
//----------------------------------------------------------------------------------


function GetCumulEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := GetCumulAntEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee) +
    GetDotationEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee);
end;

function GetCumulFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
    result := GetCumulAntFisc(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee) +
      GetDotationFisc(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee)
  else
    result := GetCumulAntEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee) +
      GetDotationEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee);
end;

function GetCumulAntDerog(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.AmortFisc.Methode <> '' then
  begin
//    result := GetCumulAntFisc(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee) -
//      GetCumulAntEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee)
      result := PlanInfo.GetCumulAntDerogatoire(DateCalcul);
  end
  else
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. : Recupere antérieurs Dérogatoire + Dotation dérog de
Suite ........ : l'exercice
Mots clefs ... :
*****************************************************************}
function GetCumulDerog(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.AmortFisc.Methode <> '' then
  begin
//    result := GetCumulFisc(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee) -
//      GetCumulEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee)
      result := PlanInfo.GetCumulAntDerogatoire(DateCalcul)
            + GetDotationDerog(PlanInfo,CodeImmo,DateCalcul)
            - GetRepriseDerog(PlanInfo,CodeImmo,DateCalcul);
  end
  else
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. : Récupération de la dotation dérogatoire
Mots clefs ... :
*****************************************************************}
function GetDotationDerog ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime) : double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if (PlanInfo.Plan.AmortFisc.Methode <> '')  then
  begin
    if PlanInfo.Derogatoire > 0 then
      result := PlanInfo.Derogatoire
    else result := 0;
  end
  else
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 06/01/2005
Modifié le ... :   /  /
Description .. : Récupération de la reprise dérogatoire
Mots clefs ... :
*****************************************************************}
function GetRepriseDerog ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime ) : double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if (PlanInfo.Plan.AmortFisc.Methode <> '')  then
  begin
    if PlanInfo.Derogatoire < 0 then
      result := (-1)*PlanInfo.Derogatoire
    else result := 0;
  end
  else
    result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 12/03/2007
Modifié le ... :   /  /
Description .. : Renvoie la réintégration fiscale calculée
Suite ........ :
Mots clefs ... :
*****************************************************************}
function GetReintegrationFEC ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime;EstImmoCedee:string) : double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if (PlanInfo.Plan.AmortFisc.Methode <> '')  then
  begin
     if (EstImmoCedee='CES') then
     begin
        { mbo 10.05.07 if PlanInfo.GetCumulAntFec(DateCalcul) >= 0 then
           result := PlanInfo.GetcumulAntFec(DateCalcul)
        else
           result := 0;}

       // mbo 31.05.07 on prend cumul antérieur + réintégration de l'exercice
        if ((PlanInfo.GetCumulAntFec(DateCalcul)) + PlanInfo.ReintFisc) < 0 then
           result := ((PlanInfo.GetcumulAntFec(DateCalcul)) + PlanInfo.ReintFisc)* (-1)
        else
           result := 0;

     end else
     begin
        if PlanInfo.ReintFisc >= 0 then
           result := PlanInfo.ReintFisc
        else
           result := 0;
     end;
  end else
     result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 13/03/2007
Modifié le ... :   /  /
Description .. : Renvoie le montant de la déduction fiscale non signé
Mots clefs ... :
*****************************************************************}
function GetDeductionFEC ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime;EstImmoCedee:string ) : double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if (PlanInfo.Plan.AmortFisc.Methode <> '')  then
  begin
     if (EstImmoCedee='CES') then
     begin
        { MODIF MBO 10.05.07 if PlanInfo.GetCumulAntFec(DateCalcul) < 0 then
           result := PlanInfo.GetCumulAntFec(DateCalcul)*-1
        else
           result := 0;}

        // mbo 31.05.07 on prend le cumul antérieur + déduction de l'exercice
        if ((PlanInfo.GetCumulAntFec(DateCalcul))+PlanInfo.ReintFisc) >= 0 then
           result := PlanInfo.GetCumulAntFec(DateCalcul) + PlanInfo.ReintFisc
        else
           result := 0;

     end else
     begin
        if PlanInfo.ReintFisc < 0 then
           result := (PlanInfo.ReintFisc * -1)
        else
           result := 0;
     end;
  end else
     result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 13/03/2007
Modifié le ... :   /  /
Description .. : renvoie un string : PRI si prime , DPI si dpi, GF si gestion fiscale
Mots clefs ... :
*****************************************************************}
function GetTypeReintegration( PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime) : string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if (GetImmoDPI(CodeImmo) <> 0.00)  then
     result := 'Déduction pour investissement'
  else if (GetMontantPrime(PlanInfo, CodeImmo, DateCalcul) <> 0.00) then
     result := 'Prime d''équipement'
  else if PlanInfo.plan.GestionFiscale = true then
     result := 'Gestion fiscale'
  else
     result := '';

end;

//================================================================

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 13/03/2007
Modifié le ... :   /  /    
Description .. : Renvoie le total réintégration fiscale à fin d'exercice n-1
Mots clefs ... :
*****************************************************************}
function GetCumulAntFEC(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee: string): double;
//var F:TextFile;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result:=PlanInfo.CumulAntFEC ;
  if EstImmoCedee='CES' then
     Result := PlanInfo.GetCumulAntFEC(DateCalcul);

// BTY***  Trace dans un fichier
{AssignFile (F,   'c:\imedcalc.log');
if   FileExists ('c:\imedcalc.log') then  Append (F)
else Rewrite (F);
Writeln (F, TimeToStr(time) + ' GetCUMULANTEco:immo= ' + CodeImmo + 'estcede= ' + EstImmoCedee +
 ' datecalcul= ' + datetostr(datecalcul) + ' planinfo.cumulanteco= ' + floattostr(PlanInfo.CumulAntEco) +
 ' result= ' + floattostr(result) );
CloseFile (F); }

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 11/05/2007
Modifié le ... :   /  /
Description .. : renvoie la valeur résiduelle comptable
Mots clefs ... :
*****************************************************************}
function GetVReco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime;
  EstImmoCedee: string):
  double;
begin

{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.Plan.GetValResiduelle(CodeImmo,PlanInfo.Plan.AmortEco ,DateCalcul, (EstImmoCedee= 'CES'))
end;



function GetVnc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime;
  EstImmoCedee: string):
  double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.AmortFisc.Methode <> '' then
    result := GetVncFisc(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee)
  else
    result := GetVncEco(PlanInfo, CodeImmo, DateCalcul, EstImmoCedee);
  if EstImmocedee = 'CES' then
    result := 0
  else if result < 0 then
    result := 0;
end;

function GetVncEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime;
  EstImmoCedee:
  string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  (*  CA - 16/10/2003 - Inutile car bien calculé par objet PlanInfo
      wMeth:=VHImmo^.PlanInfo.Plan.AmortEco.Methode ;
    if (wMeth='NAM') or (wMeth='') then
      result:=VHImmo^.PlanInfo.Plan.ValeurAchat
    else*)
  result := PlanInfo.VNCEco;
  //  if result<0 then result:=0 ;
    //Pour les etats de dotations et liste des sorties par type
    //if EstImmocedee='CES' then result:=0 ;
end;

function GetVncFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime; EstImmoCedee: string): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
//    result := PlanInfo.VNCFisc
  begin
   if EstImmoCedee = 'CES' then
    result := PlanInfo.GetValeurAchat(DateCalcul, EstImmoCedee = 'CES') - PlanInfo.GetCumulAntFisc(DateCalcul) - PlanInfo.DotationPratiquee
    else result := PlanInfo.GetValeurAchat(DateCalcul, EstImmoCedee = 'CES') - PlanInfo.CumulAntPratique - PlanInfo.DotationPratiquee;
  end
  else
    result := PlanInfo.VNCEco;
  if result < 0 then
    result := 0;
  //Pour les etats de dotations et liste des sorties par type
  //if EstImmocedee='CES' then result:=0 ;
end;

function GetCession(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): string;
//var F:TextFile;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.GetTypeSortie(DateCalcul);

// BTY***  Trace dans un fichier
{AssignFile (F,   'c:\imedcalc.log');
if   FileExists ('c:\imedcalc.log') then  Append (F)
else Rewrite (F);
Writeln (F, TimeToStr(time) + ' GetCession:immo= ' + CodeImmo +
 ' datecalcul= ' + datetostr(datecalcul) + ' planinfo.gettypesortie= ' + result );
CloseFile (F); }

end;



function GetCessionEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.CessionEco;
end;

function GetCessionFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
//    result := PlanInfo.CessionFisc
    result := PlanInfo.CessionFiscPratiquee
  else
    result := PlanInfo.CessionEco
end;

function GetCumulCessionEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.CumulCessionEco;
end;

function GetCumulCessionFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
    result := PlanInfo.CumulCessionFisc
  else
    result := PlanInfo.CumulCessionEco;
end;

function GetValResiduelle(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.ValeurResiduelle;
end;

function GetExceptionnel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  // mbo 16.11.05 result := PlanInfo.ExcepEco; on ajoute la dépréciation d'actif
  // modif mbo 30/11/05 non prise en compte de la dépréciation dans le plan fiscal
  result := PlanInfo.dTotExcEco;
end;

// nouvelle fonction pour exceptionnel sur plan fiscal - fq 17013
function GetExceptFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  // mbo 16.11.05 result := PlanInfo.ExcepEco; on ajoute la dépréciation d'actif
  // modif mbo 30/11/05 non prise en compte de la dépréciation dans le plan fiscal
  result := PlanInfo.dTotExcFisc;
end;


function GetNonDeductible(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.NonDeductible;
end;

function GetReintegration(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.Reintegration;
end;



function GetQuotePart(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.QuotePart;
end;

function GetBaseEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:
  TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.BaseEco;
end;

// BTY 11/05
function GetBaseFisc(PlanInfo:TPlanInfo; CodeImmo:string ;DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.BaseFisc;
end;


function GetPVCT(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime):
  double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.GetPVCT;
end;

function GetPVLT(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul: TDateTime):
  double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  result := PlanInfo.GetPVLT;
end;

function GetDateDebutEx(indice: integer = 0): TDateTime;
var
  DateDebut, DateFin: TDateTime;
begin
// BTY 03/06 Pb en édition Web Access avec VHIMMO
{$IFDEF EAGLSERVER}
// BTY 03/06 2e optimisation : passer par la class initiale TCPContexte
// DateDebut := TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.DateDebutEnCours;
  DateDebut := TCPContexte.GetCurrent.Exercice.Encours.Deb;
{$ELSE}
  DateDebut := VHImmo^.EnCours.Deb;
{$ENDIF}

  //Q:=OpenSql('SELECT * FROM EXERCICE WHERE EX_ETATCPTA="OUV" ORDER BY EX_EXERCICE',False) ;
  //try if not Q.Eof then DateDebut:=Q.FindField('EX_DATEDEBUT').AsDateTime ; except DateDebut:=iDate1900; end ;
  //Ferme(Q) ;
  if indice > 0 then
     CalculDateDebutFinExByPos(indice, DateDebut, DateFin);
  result := DateDebut;
end;

// BTY 02/06 Nvelles fonctions pour CWAS
//{$IFDEF EAGLSERVER}
//function GetDateDebutExWA(PlanInfo:TPlanInfo): TDateTime;
//var
//  DateDebut: TDateTime;
//begin
//  DateDebut := PlanInfo.Plan.DateDebutEnCours;
//  result := DateDebut;
//end;
//{$ENDIF}


function GetDateFinEx(indice: integer = 0): TDateTime;
var
  DateDebut, DateFin: TDateTime;
begin
// BTY 03/06 Pb en édition Web Access avec VHIMMO
{$IFDEF EAGLSERVER}
  // BTY 03/06 2e optimisation : passer par la class initiale TCPContexte
  // DateFin := TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo.Plan.DateFinEnCours;
  DateFin := TCPContexte.GetCurrent.Exercice.EnCours.Fin;
{$ELSE}
  DateFin := VHImmo^.EnCours.Fin;
{$ENDIF}

  //Q:=OpenSql('SELECT * FROM EXERCICE WHERE EX_ETATCPTA="OUV" ORDER BY EX_EXERCICE',False) ;
  //try if not Q.Eof then DateFin:=Q.FindField('EX_DATEFIN').AsDateTime ; except DateFin:=iDate2099; end ;
  //Ferme(Q) ;
  if indice > 0 then
    CalculDateDebutFinExByPos(indice, DateDebut, DateFin);

  result := DateFin;
end;

// BTY 02/06 Nvelles fonctions pour CWAS
//{$IFDEF EAGLSERVER}
//function GetDateFinExWA(PlanInfo:TPlanInfo): TDateTime;
//var
//  DateFin: TDateTime;
//begin
//  DateFin := PlanInfo.Plan.DateFinEnCours;
//  result := DateFin;
//end;
//{$ENDIF}

function GetMonnaieTenue: string;
begin
  result := RechDom('TTDEVISE', V_PGI.DevisePivot, False);
end;

function GetBaseTaxePro(indice: integer; BaseTaxePro: double; CodeImmo: string):
  double;

  function TP217(CodeImmo: string): integer;
  begin
    if ((copy(CodeImmo, 1, 4) = '2171') or //2171 Véhicules moteursmarchandises
      (copy(CodeImmo, 1, 4) = '2172') or //2172 Remarques et semi remorques
      (copy(CodeImmo, 1, 4) = '2173') or //2173 Conteneurs
      (copy(CodeImmo, 1, 4) = '2174') or //2174 Matériels de levage
      (copy(CodeImmo, 1, 4) = '2175') or //2175 Matériel de transport voyageurs
      (copy(CodeImmo, 1, 4) = '2177') or  //2177 Voitures commerciales de tourisme
      (copy(CodeImmo, 1, 3) = '217')) then // On prend aussi les autres ( FQ 13118 - 17/09/2004 )
    begin
      if GetParamSocSecur('SO_IMMOTP217','') = 'MTR' then
        Result := 2
      else
        Result := 0;
    end
    else
      Result := -1;
  end;

var
  Okok: boolean;
begin
  Okok := false;
  case indice of
    0: Okok := ((copy(CodeImmo, 1, 3) = '215') or (TP217(CodeImmo) = 0));
    // 215000<=x<=215999 installations techniques
    1: Okok := (copy(CodeImmo, 1, 4) = '2135') or (copy(CodeImmo, 1, 4) =
        '2181');
          // 213500<=x<=213599 ou 218100<=x<=218199 installations générales
    2: Okok := ((copy(CodeImmo, 1, 4) = '2182') or (TP217(CodeImmo) = 2));
    // 218200<=x<=218299 matériel transport
    3: Okok := (copy(CodeImmo, 1, 4) >= '2183') and (copy(CodeImmo, 1, 4) <
        '2185'); // 218300<=x<=218499 matériel bureau
    4: Okok := (copy(CodeImmo, 1, 4) = '2186'); // 218600<=x<=218699 emballage
    5: Okok := (copy(CodeImmo, 1, 4) = '2185') or (copy(CodeImmo, 1, 2) = '22');
    // 218500<=x<=218599 ou 220000<=x<=229999 Cheptel et ...
  end;
  if Okok then
    result := BaseTaxePro
  else
    result := 0;
end;

function GetNombreDeMois(DateDebut, DateFin: TDateTime): integer;
var
  wDate: TDateTime;
begin
  result := 0;
  if DateDebut > DateFin then
    exit;
  wDate := DateFin - DateDebut;
  result := trunc(wDate / 30);
end;

// ajout mbo 13/11/2006 - pour justesse des infos dans le cadre d'opérations touchant méthode - durée - taux
// et impression à une date antérieure aux opérations (ex changement de méthode)
// Récup durée d'amortissement éco
function GetDureeEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.AmortEco.Duree;;
end;

// Récup taux d'amortissement éco
function GetTauxEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.AmortEco.Taux;
end;

// Récup méthode d'amortissement éco
function GetMethodeEco(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  Result := PlanInfo.Plan.AmortEco.Methode;
end;

// Récup durée d'amortissement fiscale
function GetDureeFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
     Result := PlanInfo.Plan.AmortFisc.Duree
  else
     Result := PlanInfo.Plan.AmortEco.Duree;
end;

// Récup taux d'amortissement fiscal - si pas fiscal on prend l'éco
function GetTauxFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
     Result := PlanInfo.Plan.AmortFisc.Taux
  else
     Result := PlanInfo.Plan.AmortEco.Taux;
end;

function TauxFiscReel(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): double;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
     Result := PlanInfo.Plan.AmortFisc.Taux
  else
     Result := 0;   
end;

// Récup méthode d'amortissement fiscal
function GetMethodeFisc(PlanInfo: TPlanInfo; CodeImmo: string; DateCalcul:TDateTime): string;
begin
{$IFDEF EAGLSERVER}
  PlanInfo.UpdateInfo(CodeImmo, DateCalcul);
{$ELSE}
  UpdatePlanInfo(PlanInfo,CodeImmo, DateCalcul);
{$ENDIF}
  if PlanInfo.Plan.Fiscal then
     Result := PlanInfo.Plan.AmortFisc.Methode
  else
    Result := PlanInfo.Plan.AmortEco.Methode;
end;



end.


