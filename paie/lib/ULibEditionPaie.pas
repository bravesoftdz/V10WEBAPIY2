unit ULibEditionPaie;
{
  PT1 : 05/10/2007 V8_0 MF traitement des jour de fractionnment
}

interface

uses
  Utob,SysUtils,HEnt1,HCtrls
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
  {$ENDIF}
  ;

const
  DCP: integer = 2; // Nbre de décimales pour le calcul des congés payés

//PgCongesPayes
function RendPeriode(var DTcloturePi, DTDebutPi: tdatetime; DTCloture, datevalidite: tdatetime): integer;
procedure RechercheCumulCP(DD, DF: tdatetime; Salarie, Etab: string; var PN, AN, RN, BN, PN1, AN1, RN1, BN1: double; var DDP, DFP, DDPN1, DFPN1: tdatetime; var EditbulCP: string);
procedure RechercheDate(var DDebut, Dfin, Datecloture: tdatetime; Periode: string);
procedure AffichelibelleAcqPri(Periode, Salarie: string; DateDeb, Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase: double; edit, SansCloture: boolean);
procedure CompteCp(TCP: tob; Periode, Salarie: string; Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase, valo: double);
function IsAcquis(TypeConge, Sens: string): boolean;
function IsPris(TypeConge: string): boolean;
//PgCalendrier
function  PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
//P5Util
function DiffMoisJour(const DateDebut, DateFin: TDateTime; var NbreMois, NbreJour: Integer): Boolean; // Calcule difference de date à date en mois et jours restants
function EstDebutMois(const LaDate: TDateTime): boolean; // Indique si la date est le début d'un mois
function EstFinMois(const LaDate: TDateTime): boolean; //  Indique si la date est la fin d'un mois
//PgCommun
function  CalculPeriode(DTClot, DTValidite: TDatetime): integer;


implementation

function RendPeriode(var DTcloturePi, DTDebutPi: tdatetime; DTCloture, datevalidite: tdatetime): integer;
var
  i: integer;
  aa, mm, jj: word;
begin
  result := -1;
  DTcloturePi := DTCloture;
  decodedate(DTCloture, aa, mm, jj);
  DTDebutPi := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1;
  i := 0;
  while i < 10 do { au delà de p-10 , il y a sûrement une erreur ! }
  begin
    if ((datevalidite >= DTDebutPi) and (datevalidite <= DTcloturePi)) then
    begin
      result := i;
      exit;
    end;
    i := i + 1;
    DTcloturePi := DTDebutPi - 1;
    decodedate(DTcloturePi, aa, mm, jj);
    DTDebutPi := PGEncodeDateBissextile(aa - 1, MM, JJ) + 1;
  end;
  if result = -1 then
    { on essaie dans l'autre sens pour essayer de récupérer les bons intervalles }
  begin
    DTcloturePi := DTCloture;
    decodedate(DTCloture, aa, mm, jj);
    DTDebutPi := PGEncodeDateBissextile(aa, mm, jj) + 1;
    DTcloturePi := PGEncodeDateBissextile(aa + 1, mm, jj);
    i := 0;
    while i < 10 do { au delà de p+10 , ca va , on arrête ! }
    begin
      if (datevalidite >= DTDebutPi) and (datevalidite <= DTcloturePi) then exit;
      i := i + 1;
      DTdebutPi := DTcloturePi + 1;
      decodedate(DTcloturePi, aa, mm, jj);
      DTcloturePi := PGEncodeDateBissextile(aa + 1, mm, jj);
    end;
  end;
end;

function PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
begin
  if (MM = 2) and (JJ = 29) then //Année bissextile
  begin
    Result := encodedate(AA, MM, 1);
    Result := FindeMois(Result);
  end
  else
    Result := encodedate(AA, MM, JJ);
end;

{ Fonction qui calcule le nombre de mois entre 2 dates et le nombre de jours restants
NbreMois contient le nombre de mois entier entre les 2 dates et NbreJour contient le nombre de
jours entre la date de fin et le debut du mois concernant la date de fin
}
function DiffMoisJour(const DateDebut, DateFin: TDateTime; var NbreMois, NbreJour: Integer): Boolean;
var
  PremMois, PremAnnee, NMois: WORD;
  DateCal: TDateTime;
  Calcul: double;
begin
  result := FALSE;
  PremMois := 0;
  PremAnnee := 0;
  NMois := 0;
  Calcul := 0; // PT31
  if DateDebut = DateFin then
  begin
    NbreJour := 1;
    NbreMois := 0;
    result := TRUE;
    exit;
  end;
  NOMBREMOIS(DateDebut, DateFin, PremMois, PremAnnee, NMois);
  //if (EstFinMois(DateFin)) then NMois:=NMois-1; // Car prend en compte le mois non complet
  if NMois > 0 then NMois := NMois - 1;
  NbreMois := NMois;
  if NMois = 0 then
  begin
    Calcul := DateFin - DateDebut;
    NbreJour := StrToInt(FloatToStr(Calcul));
    exit;
  end;
  {DateCal:=FindeMois (DateDebut);
  Calcul:=DateCal-DateDebut;}
  // 3 cas à gérer
  if (EstDebutMois(DateDebut)) and (not EstFinMois(DateFin)) then
  begin
    DateCal := DebutdeMois(DateFin);
    Calcul := DateFin - DateCal;
  end;
  if (not EstDebutMois(DateDebut)) and (EstFinMois(DateFin)) then
    Calcul := (FinDeMois(DateDebut)) - DateDebut;
  if Calcul >= 31 then Calcul := 30;
  if (not EstDebutMois(DateDebut)) and (not EstFinMois(DateFin)) then
  begin // On donne simplement le nombre de jours entre les 2 dates
    Calcul := DateFin - Datedebut;
    NbreMois := 0;
    result := TRUE;
  end;
  if Calcul = 0 then
  begin
    result := TRUE;
    Calcul := 1;
  end; // La date de fin correspond à un début de mois donc 1 trentieme
  // calcul ne peut contenir qu'une valeur entiere <= 31 jours ou > selon les 3 cas si on est sur plusieurs mois
  NbreJour := StrToInt(FloatToStr(Calcul));
end;

procedure RechercheCumulCP(DD, DF: tdatetime; Salarie, Etab: string; var PN, AN, RN, BN, PN1, AN1, RN1, BN1: double; var DDP, DFP, DDPN1, DFPN1: tdatetime; var EditbulCP: string);
var
  DTClot, DDebut, Dfin, DSolde: TDateTime;
  Q: TQuery;
  Periodei: Integer;
  Periode, PeriodeN1, st: string;
  MB: Double;
begin
  EditbulCP := '';
  DTClot := Idate1900;
  Q := Opensql('SELECT ETB_CONGESPAYES,ETB_DATECLOTURECPN,ETB_EDITBULCP,' +
    'PSA_TYPEDITBULCP,PSA_EDITBULCP,PSA_CONGESPAYES ' +
    'FROM SALARIES ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    'WHERE PSA_SALARIE="' + Salarie + '"', True);
  if not Q.eof then
  begin
    if Q.findfield('ETB_CONGESPAYES').AsString = '-' then
    begin
      Ferme(Q);
      exit;
    end;
    if Q.findfield('PSA_CONGESPAYES').AsString = '-' then
    begin
      Ferme(Q);
      exit;
    end;
    DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
    if Q.FindField('PSA_TYPEDITBULCP').AsString = 'ETB' then EditbulCP := Q.FindField('ETB_EDITBULCP').AsString
    else EditbulCP := Q.FindField('PSA_EDITBULCP').AsString;

  end;
  Ferme(Q);
  st := 'SELECT PCN_DATEVALIDITE FROM ABSENCESALARIE WHERE PCN_SALARIE="' + Salarie + '" ' +
    'AND PCN_DATEVALIDITE<"' + Usdatetime(DF) + '" AND PCN_TYPECONGE="SLD" AND PCN_TYPEMVT="CPA" ' +
    'AND PCN_GENERECLOTURE="-" ' +
    'ORDER BY PCN_DATEVALIDITE DESC';
  Q := OpenSql(st, TRUE);
  if (not Q.eof) then
    DSolde := Q.FindField('PCN_DATEVALIDITE').AsDateTime
  else
    DSolde := iDate1900;
  Ferme(Q);
  DDebut := 0;
  Dfin := 0;
  Periodei := CalculPeriode(DtClot, DF);
  Periode := inttostr(Periodei);
  PeriodeN1 := Inttostr(Periodei + 1);
  RechercheDate(DDebut, Dfin, DtClot, Periode);
  AffichelibelleAcqPri(Periode, Salarie, DSolde, DF, PN, AN, RN, BN, MB, False, True);
  DDP := DDebut;
  DFP := Dfin;
  DDebut := 0;
  Dfin := 0;
  RechercheDate(DDebut, Dfin, DtClot, PeriodeN1);
  AffichelibelleAcqPri(PeriodeN1, Salarie, DSolde, DF, PN1, AN1, RN1, BN1, MB, True, False);
  DDPN1 := DDebut;
  DFPN1 := Dfin;
end;

function CalculPeriode(DTClot, DTValidite: TDatetime): integer;
var
  Dtdeb, DtFin, DtFinS: TDATETIME;
  aa, mm, jj: word;
  i: integer;
begin
  result := -1;
  if DTClot <= idate1900 then exit;
  Decodedate(DTclot, aa, mm, jj);
  DtDeb := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1;
  DtFin := DtClot;
  DtFinS := PGEncodeDateBissextile(AA + 1, MM, JJ);
  if Dtvalidite > Dtfins then
  begin
    result := -9;
    exit;
  end;
  if DtValidite > DtClot then exit;
  result := 0;
  i := 0;
  while not ((DTValidite >= DtDeb) and (DTValidite <= DtFin)) do
  begin
    i := i + 1;
    if i > 50 then exit; // pour ne pas boucler au cas où....
    result := result + 1;
    DtFin := DtDeb - 1;
    Decodedate(DTFin, aa, mm, jj);
    DtDeb := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1;
  end;
end;

procedure RechercheDate(var DDebut, Dfin, Datecloture: tdatetime; Periode: string);
var
  Noperiode, i: integer;
  aa, mm, jj: word;
begin
  DDebut := 0;
  DFin := 0;
  if Isnumeric(Periode) then NoPeriode := strtoint(Periode)
  else exit;
  DFin := Datecloture;
  Decodedate(DFin, aa, mm, jj);
  DDebut := PGEncodeDateBissextile(aa - 1, mm, jj);
  DDebut := DDebut + 1;
  i := 0;
  while i < NoPeriode do
  begin
    DFin := DDebut - 1;
    Decodedate(DFin, aa, mm, jj);
    DDebut := PGEncodeDateBissextile(aa - 1, mm, jj);
    DDebut := DDebut + 1;
    i := i + 1;
  end;
end;

procedure AffichelibelleAcqPri(Periode, Salarie: string; DateDeb, Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase: double; edit, SansCloture: boolean);
var
st: string;
valo: Double;
TAbs: Tob;
begin
Pris:= 0;
Acquis:= 0;
Restants:= 0;
Base:= 0;
Moisbase:= 0;
if (Salarie = '') then
   exit;
{affichage du libellé correspondant à la période sélectionnée}
Tabs:= Tob.create ('ABSENCESALARIE', nil, -1);
st:= 'SELECT *'+
     ' FROM ABSENCESALARIE WHERE'+
     ' PCN_SALARIE="'+Salarie+'" AND'+
     ' PCN_TYPEMVT="CPA" AND'+
     ' PCN_PERIODECP='+Periode+' AND'+
     ' PCN_CODERGRPT<>-1 AND'+
     ' PCN_DATEVALIDITE>"'+USDateTime (DateDeb)+'"';
if (edit) and (StrToInt(periode) > 0) then
   begin
   St:= 'SELECT *'+
        ' FROM ABSENCESALARIE WHERE'+
        ' PCN_SALARIE="'+Salarie+'" AND'+
        ' PCN_DATEVALIDITE<="'+USDateTime (DateFin)+'" AND'+
        ' PCN_DATEVALIDITE>"'+USDateTime (DateDeb)+'" AND'+
        ' PCN_TYPEMVT="CPA" AND'+
        ' PCN_PERIODECP='+Periode+' AND'+
        ' PCN_CODERGRPT<>-1 AND'+
        ' ((PCN_GENERECLOTURE="-") OR'+
        ' (PCN_GENERECLOTURE="X" AND'+
        ' ((PCN_SENSABS="+" AND'+
        ' PCN_TYPECONGE<>"SLD") OR'+
        ' (PCN_TYPECONGE="SLD"))))';
{J'exclue les mvts générés par la cloture de période précédente}
   end;

if SansCloture = True then
   begin
   st:= 'SELECT *'+
        ' FROM ABSENCESALARIE WHERE'+
        ' PCN_SALARIE="'+Salarie+'" AND'+
        ' PCN_DATEVALIDITE<="'+USDateTime (DateFin)+'" AND'+
        ' PCN_DATEVALIDITE>"'+USDateTime (DateDeb)+'" AND'+
        ' PCN_TYPEMVT="CPA" AND'+
        ' PCN_PERIODECP='+Periode+' AND'+
        ' PCN_CODERGRPT<>-1 AND'+
        ' (PCN_GENERECLOTURE="-")';
   end;
{Flux optimisé
  Q := Opensql(st, True);
  if not Q.eof then
    Tabs.loaddetaildb('ABSENCESALARIE', '', '', Q, false, false);
  Ferme(Q);
  if Tabs = nil then exit;
}
Tabs.loaddetaildbFromSQL ('ABSENCESALARIE', st);
if (Tabs.detail.count<=0) then
   begin
   FreeAndNil (Tabs);
   exit;
   end;
Valo:= 0;
CompteCP (TAbs, Periode, Salarie, Datefin, Pris, Acquis, Restants, Base,
          Moisbase, Valo);
FreeAndNil (Tabs);
Pris:= Arrondi (Pris, DCP);
Acquis:= Arrondi (Acquis, DCP);
Base:= Arrondi (Base, DCP);
Restants:= Arrondi (Restants, DCP);
end;

procedure CompteCp(TCP: tob; Periode, Salarie: string; Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase, valo: double);
var
  Typeconge, Sens: string;
  Jours, MBase: double;
  T: Tob;
begin
  { datefin = 0 signifie qu'on n'est pas dans un bulletin }
  if datefin = 0 then datefin := 99999;
  Acquis := 0;
  Pris := 0;
  Base := 0;
  Restants := 0;
  MoisBase := 0;
  T := TCP.findfirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
  while T <> nil do
  begin { PCN_CODERGRPT <> -1  }
    if T.Getvalue('PCN_CODERGRPT') <> '-1' then
    begin
      Typeconge := T.getvalue('PCN_TYPECONGE');
      Sens := T.getvalue('PCN_SENSABS');
      Jours := T.getvalue('PCN_JOURS');
      MBase := T.getvalue('PCN_BASE');
      if IsAcquis(TypeConge, Sens) then
        if Sens = '+' then
        begin
          Acquis := ACquis + Jours;
          MoisBase := MoisBase + T.getvalue('PCN_NBREMOIS');
          Restants := Restants + Jours;
          Base := Base + MBase;
          if (T.getvalue('PCN_NBREMOIS')) = 0 then
            Valo := Valo + 0
          else
            Valo := Valo + (arrondi((12 * MBase) / (T.getvalue('PCN_NBREMOIS') * 10), DCP) * Jours);
        end
        else
        begin
          MoisBase := MoisBase - T.getvalue('PCN_NBREMOIS');
          Base := Base - MBase;
          Acquis := ACquis - Jours;
          Restants := Restants - Jours;
        end; { on prend les CPA de la période car ils st forcément payés ! }
      if ((IsPris(TypeConge)) and ((T.getvalue('PCN_DATEFIN') > 10) and
        (T.getvalue('PCN_DATEFIN') <= datefin) and
        (T.getvalue('PCN_DATEPAIEMENT') > 10)) or (Typeconge = 'CPA')) then
        if Sens = '-' then
        begin
          Pris := Pris + Jours;
          Restants := Restants - Jours;
        end
        else
        begin
          Pris := Pris - Jours;
          Restants := Restants + Jours;
        end;
      if ((TypeConge = 'REL') and (sens = '-')) then
        Restants := Restants - Jours;
    end;
    T := TCP.findnext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
  end; //while
end;

function IsAcquis(TypeConge, Sens: string): boolean;
begin
  result := false;
  if ((typeconge = 'ACQ') or
    (typeconge = 'ACS') or
    (typeconge = 'ACA') or
    (typeconge = 'ACF') or // PT1
    (typeconge = 'AJU') or
    (typeconge = 'ARR') or
    ((typeconge = 'REL') and (Sens = '+')) or
    (typeconge = 'REP')) then result := true;
end;

function IsPris(TypeConge: string): boolean;
begin
  result := false;
  if ((typeconge = 'PRI') or (typeconge = 'SLD') or
    (typeconge = 'AJP') or (typeconge = 'CPA')) then result := true;
end;

// Fonction qui indique si la date est un début de mois
function EstDebutMois(const LaDate: TDateTime): boolean;
var
  UneDate: TDateTime;
begin
  result := FALSE;
  UneDate := LaDate;
  if UneDate = DebutdeMois(LaDate) then result := TRUE;
end;

// Fonction qui indique si la date est une fin de mois
function EstFinMois(const LaDate: TDateTime): boolean;
var
  UneDate: TDateTime;
begin
  result := FALSE;
  UneDate := LaDate;
  if UneDate = FindeMois(LaDate) then result := TRUE;
end;

end.

