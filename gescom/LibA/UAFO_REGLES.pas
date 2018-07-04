unit UAFO_REGLES;

interface

uses classes,
  hctrls,
  utob,
  sysUtils,
  HEnt1,

  {$IFDEF EAGLCLIENT}

  {$ELSE}
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF}

  afPlanningCst,
  paramSoc;
type

		TInternalDate = record
    	TheDate : TDateTime;
    end;
    PInternalDate = ^ TinternalDate;
    //
		TListDate = class (Tlist)
    private
      function GetItems(Indice: integer): TDateTime;
      procedure SetItems(Indice: integer; const Value: TDateTime);
    	function GetPDate(Indice: integer): PInternalDate;
    public
    	property Items [Indice : integer] : TDateTime read GetItems write SetItems;
      function IndexOf(Item: TDateTime): Integer;
      function Add(Item: TDateTime): Integer;
      procedure Delete(Index: Integer);
    end;

    TAFRegles = class

    private

      fInFrequence        : Integer;
      fBoGeneAuPlusTot    : Boolean;
      //fBoGestionDateFin   : Boolean;
      fBoJourFeriesWork   : Boolean;
      fStUniteTemps       : String;
      fRdQteInterv        : Double;
    fDtLastDateGene: TDateTime;
      fDtDateDeb          : TDateTime;
      fDtDateFin          : TDateTime;
      fDtDateAnnuelle     : TDateTime;
    fInAnneeNb: Integer; // nombre de jours entre chques intervention
      fInJourInterval     : Integer; // nombre de jours entre chques intervention
      fInSemaineInterv    : Integer; // nombre de semaines entre chaque intervention
      fInMoisMethode      : Integer; // Methode mensuelle
      fInMoisJourFixe     : Integer; // numero du jour d'intervention dans le mois
      fInMoisFixe         : Integer; // nombre de mois entre chaque intervention
      fStMoisSemaine      : String;  // code tablette du numero de semaine d'intervention dans un mois
      fInNbIntervention   : Integer; // nombre d'interventions prévues
      fMoisJourLib        : TJour;   // libele du jour d'intervention dans la semaine
      fListeJours         : TJour;   // liste des jours
    //      fInStatutRes        : Integer; // Statut de la ressource
      fInQteRessource     : Integer; // Quantité initiale de la ressource
      fInNbJoursDecal     : Integer; // Nombre de jour de décalage maximum
    fStMethodeDecal: string; // méthode de décalage des jours

    procedure GetData(pQr: TQuery; pStPrefixe: string; Modegene: integer; pDateDeb, pDateFin: Tdatetime);
    function GenereListeAnnuelle(pInNbJour: Integer): TListDate;
      function GenereListeMensuelle(pInNbJour : Integer): TListDate;
      function GenereListeQuotidienne(pInNbJour : Integer): TListDate;
      function GenereListeHebdomadaire(pInNbJour : Integer): TListDate;
      function GenereListeNbintervention: TListDate;

      public
    function LoadDBReglesRes(pStAffaire, pStNumeroTache, pStCodeRessource: string; var pRdQte: Double; Compteur: boolean): Boolean;
    function LoadDBReglesTaches(pStAffaire, pStNumeroTache: string; ModeGene: integer): Boolean;
    function LoadDBReglesModele(pStModele: string; ModeGene: Integer; Datedeb, Datefin: tdatetime): Boolean;
    function GenereListeJours(var pRdDuree: Double; var pInNbJoursDecal: Integer; var pStMethodeDecal: string; pInNbJour: Integer): TListDate;
    function QuantiteRessource: Integer;
    property DtDateDeb: TDateTime read fDtDateDeb;
    property DtDateFin: TDateTime read fDtDateFin;
  end;                                                                        

implementation
uses
  UtilRessource
;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... :
Modifié le ... :
Description .. : result = false s'il n'existe pas de règles particulières
                 retourne la quantité initiale de la ressource si on planifie une quantité limité de jours

                 Modif MCD : si on n'est pas en RAP, on ressort de la fct. inutile !!!.
                 surppime ausis la jointure sur tache (?? inutile) et ne rend que la qté initiale (? le seul géré ???)
                 Ajout aussi de la lecture du planning existant ... seul moyen de connaitre le vrai RAP

                 mcd 27/01/04 pas de règles par ressource..
                 à voir éventuellement plus tard ..(après modif table)
Mots clefs ... :
*****************************************************************}
function TAFRegles.LoadDBReglesRes(pStAffaire, pStNumeroTache, pStCodeRessource: string; var pRdQte: Double; Compteur: boolean): boolean;
var
  vSt: string;
  vQr: TQuery;
  QteInit: Double;
  QtePla: Double;

begin

  if not Compteur then
    result := false
  else
  begin
    Qteinit := 0;
    QtePla := 0;
    pRdQte := 0;

    vSt := 'SELECT ATR_QTEINITPLA FROM TACHERESSOURCE WHERE ATR_AFFAIRE = "' + pStAffaire + '"';
    if pStNumeroTache <> '' then
      vSt := vSt + ' AND ATR_NUMEROTACHE = "' + pStNumeroTache + '"';
    vSt := vSt + ' AND ATR_RESSOURCE = "' + pStCodeRessource + '"';
    vQr := nil;
    try
      vQR := OpenSql(vSt, True,-1,'',true);
      if not vQR.Eof then
      begin
        // C.B 27/09/2005
        // ne generait rien
        result := true;
        QteInit := vQR.findField('ATR_QTEINITPLA').AsFloat;
      end
      else
        result := false;
    finally
      ferme(vQr);
    end;

    if result then
    begin
      // on est en RAP, il faut donc aller lire le planning existant
      vSt := 'SELECT sum(APL_QTEPLANIFIEE) as QTEPLA FROM AFPLANNING WHERE APL_AFFAIRE = "' + pStAffaire + '"';
      if pStNumeroTache <> '' then
        vSt := vSt + ' AND APL_NUMEROTACHE = "' + pStNumeroTache + '"';
      vSt := vSt + ' AND APL_RESSOURCE = "' + pStCodeRessource + '"';
      vQr := nil;
      try
        vQR := OpenSql(vSt, True,-1,'',true);
        if not vQR.Eof then
        begin
          result := false;
          QtePla := vQR.findField('QTEPLA').AsFloat;
        end
        else
          result := false;
      finally
        ferme(vQr);
      end;

      // calcul de la qté réelle restant à planifier
      pRdQte := Qteinit - Qtepla;
    end;
  end;
end;

function TAFRegles.LoadDBReglesTaches(pStAffaire, pStNumeroTache: string; ModeGene: Integer): Boolean;
var
  vSt: string;
  vQr: TQuery;
  vStPrefixe: string;

begin

  result := true;
  vSt := 'SELECT * FROM TACHE WHERE ATA_AFFAIRE = "' + pStAffaire + '"';
  if pStNumeroTache <> '' then
    vSt := vSt + ' AND ATA_NUMEROTACHE = "' + pStNumeroTache + '"';
  vStPrefixe := 'ATA_';

  vQr := nil;
  try
    vQR := OpenSql(vSt, True,-1,'',true);
    if not vQR.Eof then
      GetData(vQr, vStPrefixe, ModeGene, iDate1900, iDate2099) //mcd 12/05/05
    else
      result := false;
  finally
    ferme(vQr);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : actuellement(03/2004) appelé qu'avec le préfixe ATA. pas
                 de règles au niveau de tache ressource
                 mcd 12/05/2005 appelé à ce jour pour AFM, pourmodèle tache. utilisé dans DP
Mots clefs ... :
*****************************************************************}

procedure TAFRegles.GetData(pQr: TQuery; pStPrefixe: string; Modegene: integer;
  pDateDeb, pDateFin: TDatetime);
var
  Qq: Tquery;
  vTobJour: Tob;
  vTobDet: Tob;
  vStSql: string;
  ii: Integer;

begin

  //C.B FQ 11457
  {if pQr.findField(pStPrefixe + 'PLANNINGPERIOD').AsString = 'A' then
    fInFrequence := cInAnnuelle
  else
    if pQr.findField(pStPrefixe + 'PLANNINGPERIOD').AsString = 'M' then
    fInFrequence := cInMensuelle
  else
    if pQr.findField(pStPrefixe + 'PLANNINGPERIOD').AsString = 'Q' then
    fInFrequence := cInQuotidienne
  else
    if pQr.findField(pStPrefixe + 'PLANNINGPERIOD').AsString = 'S' then
    fInFrequence := cInHebdomadaire
  else
    if pQr.findField(pStPrefixe + 'PLANNINGPERIOD').AsString = 'NBI' then
    fInFrequence := cInNbInterv;}

  if pQr.findField('ATA_PERIODICITE').AsString = 'A' then
    fInFrequence := cInAnnuelle
  else
    if pQr.findField('ATA_PERIODICITE').AsString = 'M' then
    fInFrequence := cInMensuelle
  else
    if pQr.findField('ATA_PERIODICITE').AsString = 'Q' then
    fInFrequence := cInQuotidienne
  else
    if pQr.findField('ATA_PERIODICITE').AsString = 'S' then
    fInFrequence := cInHebdomadaire
  else
    if pQr.findField('ATA_PERIODICITE').AsString = 'NBI' then
    fInFrequence := cInNbInterv;

  fBoGeneAuPlusTot := pQr.findField(pStPrefixe + 'MODEGENE').AsInteger = 1;
  //fBoGestionDateFin   := pQr.findField(pStPrefixe + 'GESTIONDATEFIN').AsString = 'X';
  fBoJourFeriesWork := pQr.findField(pStPrefixe + 'JOURFERIE').AsString = 'X';
  fRdQteInterv := pQr.findField(pStPrefixe + 'QTEINTERVENT').AsFloat;
  fStUniteTemps := pQr.findField(pStPrefixe + 'UNITETEMPS').AsString;
  fDtLastDateGene := pQr.findField(pStPrefixe + 'LASTDATEGENE').AsDateTime;

  //mcd 12/05/05 pour DP
  if pStPRefixe = 'AFM_' then
  begin
    fDtDateDeb := pQr.findField('AFM_DATEDEBPERIOD').AsDateTime;
    if fDtDateDeb = idate1900 then
      fDtDateDeb := pDateDeb;
  end
    //fin mcd 12/05/05
  else
  begin
    fDtDateDeb := pQr.findField(pStPrefixe + 'DATEDEBPERIOD').AsDateTime;
    // si ajout, on ne part pas de la date de debut tache, mais de la date de dernière génération +1j
    if (modegene = cinAjout) and (fDtLastDateGene > fdTDateDeb) then
       fDtDateDeb := PlusDate(fDtLastDateGene, 1, 'J'); // si planning déjà générer on part de la derniere date de génération, et non pas de la date début
  end;

  //mcd 12/05/05
  // si la date de fin est supérieure à la date butoire, on ne génère que jusqu'à la
  // date butoire et si on ne tiens pas compte de la date de fin, on génère jusqu'à la date butoire
  if pStPRefixe = 'AFM_' then //mcd 12/05/05 pour DP
  begin
    fDtDateFin := pQr.findField('AFM_DATEFINPERIOD').AsDateTime;
    if fDtDateFin = idate2099 then
      fDtDateFin := pDateFin;
  end
  else
  begin
    fDtDateFin := pQr.findField(pStPrefixe + 'DATEFINPERIOD').AsDateTime;
//Modiffv if (fDtDateFin > GetParamSocSecur('SO_AFDATEBUTOIRE', iDate1900)) then
//Modiffv    fDtDateFin := GetParamSocSecur('SO_AFDATEBUTOIRE', iDate1900);
  end;

  fDtDateAnnuelle := pQr.findField(pStPrefixe + 'DATEANNUELLE').AsDateTime;
  fInAnneeNb := pQr.findField(pStPrefixe + 'ANNEENB').AsInteger;
  fInJourInterval := pQr.findField(pStPrefixe + 'JOURINTERVAL').AsInteger;
  fInSemaineInterv := pQr.findField(pStPrefixe + 'SEMAINEINTERV').AsInteger;
  fInMoisMethode := pQr.findField(pStPrefixe + 'MOISMETHODE').AsInteger;
  fInMoisJourFixe := pQr.findField(pStPrefixe + 'MOISJOURFIXE').AsInteger;

  // mcd 25/08/04 .. en cas ,evite de boucler ,jour 0 incorrect
  if fInMoisJourFixe = 0 then
    fInMoisJourFixe := 1;

  fInMoisFixe := pQr.findField(pStPrefixe + 'MOISFIXE').AsInteger;
  fStMoisSemaine := pQr.findField(pStPrefixe + 'MOISSEMAINE').AsString;
  fInNbIntervention := pQr.findField(pStPrefixe + 'NBINTERVENTION').AsInteger;

  if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J1' then
    fMoisJourLib := [LUN]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J2' then
    fMoisJourLib := [MAR]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J3' then
    fMoisJourLib := [MER]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J4' then
    fMoisJourLib := [JEU]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J5' then
    fMoisJourLib := [VEN]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J6' then
    fMoisJourLib := [SAM]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'J7' then
    fMoisJourLib := [DIM]
  else
    if pQr.findField(pStPrefixe + 'MOISJOURLIB').AsString = 'WE' then
    fMoisJourLib := [WE];

  fListeJours := [];

  // mcd 27/01/04 revu pour traiter à partir de la table annexe
  // on prend tout,car tout est utilisé .. peu d"enrgt et de champ

  //mcd 13/05/05
  if pStPrefixe = 'AFM_' then
    vStSql := 'SELECT * FROM AfTacheJour WHERE ATJ_TYPEJOUR="MOD"  AND ATJ_MODELETACHE="'
      + pQr.findField(pStPrefixe + 'MODELETACHE').AsString + '"'

  else
    vStSql := 'SELECT * FROM AfTacheJour WHERE ATJ_TYPEJOUR="AFF" AND ATJ_AFFAIRE="'
      + pQr.findField(pStPrefixe + 'AFFAIRE').AsString + '" AND'
      + ' ATJ_NUMEROTACHE=' + IntToStr(pQr.findField(pStPrefixe + 'NUMEROTACHE').AsInteger);

  QQ := OpenSql(vStSql, True,-1,'',true);
  vTobJour := Tob.create('#AFTACHEJOUR', nil, -1);
  try
    if not QQ.EOF then
      vTobJour.LoadDetailDb('#AFTACHEJOUR', '', '', QQ, False);

    for ii := 0 to vTobJour.detail.count - 1 do
    begin
      vTobDet := vTobJour.detail[ii];
      if vTobDet.getvalue('ATj_JOURAPLANIF') = 1 then
        fListeJours := fListeJours + [LUN]
      else
        if vTobDet.getvalue('ATj_JOURAPLANIF') = 2 then
        fListeJours := fListeJours + [MAR]
      else
        if vTobDet.getvalue('ATj_JOURAPLANIF') = 3 then
        fListeJours := fListeJours + [MER]
      else
        if vTobDet.getvalue('ATj_JOURAPLANIF') = 4 then
        fListeJours := fListeJours + [JEU]
      else
        if vTobDet.getvalue('ATj_JOURAPLANIF') = 5 then
        fListeJours := fListeJours + [VEN]
      else
        if vTobDet.getvalue('ATj_JOURAPLANIF') = 6 then
        fListeJours := fListeJours + [SAM]
      else
        if vTobDet.getvalue('ATj_JOURAPLANIF') = 7 then
        fListeJours := fListeJours + [DIM];
      // voir pour le jour 0 si à traiter
      // voir aussi pour traiter ici les durée et heure début
    end;

    // quantité initiale de la ressource
    // utilisé lorsqu'on utilise un compteur de jours
    if pStPrefixe <> 'AFM_' then
      fInQteRessource := pQr.findField(pStPrefixe + 'QTEINITPLA').AsInteger;
    fInNbJoursDecal := pQr.findField(pStPrefixe + 'NBJOURSDECAL').AsInteger;
    fStMethodeDecal := pQr.findField(pStPrefixe + 'METHODEDECAL').AsString;

  finally
    Ferme(QQ);
    FreeAndNil(vTobJour);
  end;
end;

//mcd 12/05/05 fct qui permet de récupérer les règles pour les tache modèle
//fait pour les obligations du dp

function TAFRegles.LoadDBReglesModele(pStModele: string; ModeGene: Integer; DateDeb, Datefin: tdatetime): Boolean;
var
  vSt: string;
  vQr: TQuery;
  vStPrefixe: string;

begin
  result := true;
  //une seule tache, on peut se permettre de tot prendre.
  //dans ce cas, il est dommage de refaire la requête qui a été faite
  //dans la fct appelant, mais seul moyen, si on veut utiliser GetData de cette classe
  vSt := 'SELECT * FROM AFMODELETACHE WHERE AFM_MODELETACHE = "' + pStModele + '"';
  vStPrefixe := 'AFM_';
  vQr := nil;
  try
    vQR := OpenSql(vSt, True,-1,'',true);
    if not vQR.Eof then
    begin
      GetData(vQr, vStPrefixe, ModeGene, DateDeb, DateFin);
    end
    else
      result := false;
  finally
    ferme(vQr);
  end;
end;

// fct qui calcule les dates annuelle. celles ci se font tous les ans à la date saisie
// (année renseignée dans la date, jamais utilisé), en tenant compte de tous les combien (anneenb)

function TAFRegles.GenereListeAnnuelle(pInNbJour: integer): TListDate;
var
  jj, mm, aa: Word;
  jjd, aad, mmd: word;
  vInNbJour: Integer;
  ii: Integer;
  vDtCurDate: TDateTime;
  vBoCompteur: Boolean;
  HeureDeb : double;

begin

  result := TListDate.create;
  vInNbJour := pInNbJour;
  vBoCompteur := not (vInNbJour = cInNoCompteur);
  Decodedate(fDtDateAnnuelle, aa, mm, jj); // on récupère le mois et l'année d'intervention

  // cas au plut tôt
  if fBoGeneAuPlusTot then
  begin
    Decodedate(fDtDateDeb, aad, mmd, jjd); // on récupère l'année de la date de début
    vDtCurDate := EncodeDate(aad, mm, jj); // fabrication date
    ii := 0;
    while ii = 0 do
    begin // si date comprise dans les période, on ajoute la date + gestion du RAP si géré
      if (vDtCurDate <= fDtDateFin) and (vDtCurDate >= fDtDateDeb) and ((not vBoCompteur) or (vInNbJour > 0)) then
      begin
      	// Ajout de l'heure de départ
        HeureDeb    := Getparamsoc('SO_HEUREDEB');
     		vDtCurDate := Trunc(vDtCurDate) + (HeureDeb);
        // -----
        result.Add(vDtCurDate);
        if vBoCompteur then
          vInNbJour := vInNbJour - 1;
      end
      else
        if (vDtCurDate >= fDtDateDeb) then ii := 1; // pour sortir de la boucle  ,mais uniquemen si faux sur autres chose que date début
      vDtCurDate := PlusDate(vDtCurDate, fInAnneeNb, 'A'); // tache annuelle, on ajoute le nombre d'année géré sur cette tache annuelle
    end;
  end

    // cas génération au plus tard
  else
  begin
    Decodedate(fDtDateFIN, aad, mmd, jjd); // on récupère l'année de la date de fin
    vDtCurDate := EncodeDate(aad, mm, jj); // fabrication date
    ii := 0;
    while ii = 0 do
    begin // date calculée compris dans les dates début et fin
      if (vDtCurDate >= fDtDateDeb) and (vDtCurDate <= fDtDateFin) and ((not vBoCompteur) or (vInNbJour > 0)) then
      begin
      	// Ajout de l'heure de départ
        HeureDeb    := Getparamsoc('SO_HEUREDEB');
     		vDtCurDate := Trunc(vDtCurDate) + (HeureDeb);
        // -----
        result.Add(vDtCurDate);
        if vBoCompteur then
          vInNbJour := vInNbJour - 1;
      end
      else
        if (vDtCurDate <= fDtDateFin) then
        ii := 1; //on sort de la boucle, sauf si date fin incorrecte
      vDtCurDate := PlusDate(vDtCurDate, -fInAnneeNb, 'A'); // tache annuelle, on retranche le nb d'année de  la tache
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :
Description .. : On ne tiend pas compte du nombre de jours a planifier
                 si pInNbJour = cInNoCompteur
Mots clefs ... :
*****************************************************************}

function TAFRegles.GenereListeMensuelle(pInNbJour: Integer): TListDate;
var
  vDtCurDate, vdtTrav: TDateTime;
  vInNbJour, ii: Integer;
  vBoCompteur: Boolean;

  // recherche un jour dans la semaine
  function RechJourSemaine(pDtDate: TDateTime): TDateTime;
  var
    Year, Month, Day: Word;
    vInJour: Integer;
    vInNumSem: Integer;
    vInYear: Integer;

  begin
    decodeDate(pDtDate, Year, Month, Day);
    if LUN in fMoisJourLib then
      vInJour := 0
    else
      if MAR in fMoisJourLib then
      vInJour := 1
    else
      if MER in fMoisJourLib then
      vInJour := 2
    else
      if JEU in fMoisJourLib then
      vInJour := 3
    else
      if VEN in fMoisJourLib then
      vInJour := 4
    else
      if SAM in fMoisJourLib then
      vInJour := 5
    else
      if DIM in fMoisJourLib then
      vInJour := 6
    else
      if WE in fMoisJourLib then
    begin
      vInJour := 5;
      fRdQteInterv := AFConversionUnite('J', fStUniteTemps, 2);
    end
    else
      vInJour := 0;

    // C.B 14/10/2004
    // séparer les 2 fonctions
    vInNumSem := NumSemaine(pDtDate, vInYear);
    result := PremierJourSemaine(vInNumSem, vInYear) + vInJour;
    if result < pDtDate then
      result := result + 7;
  end;

  // ajout de l'ensemble des jours planifiés
  procedure AddJoursSemaine(pDtDate: TDateTime; pListeJours: TListDate; pBoCompteur: boolean; var pInNbJourAPlanifier: Integer);

    procedure AddJour(var pDtDate: TDateTime; pListeJours: TListDate; pInNbJour: Integer);
    var
      vDtAdd: TDateTime;
      Year, Month, Day: Word;
      vInNumSem: Integer;
      vInYear: Integer;

    begin
      decodeDate(pDtDate, Year, Month, Day);
      // C.B 14/10/2004
      // séparer les 2 fonctions
      vInNumSem := NumSemaine(pDtDate, vInYear);
      vDtAdd := PremierJourSemaine(vInNumSem, vInYear) + pInNbJour;
      if vDtAdd < pDtDate then
        vDtAdd := vDtAdd + 7;
      vDtAdd := PlusDate(vDtAdd, StrToInt(fStMoisSemaine) - 1, 'S');
      // ajout d'un controle si le jour trouvé est dans la periode
      if (vDtAdd >= fDtDateDeb) and (vDtAdd <= fDtDateFin) then
      begin
        pListeJours.Add(vDtAdd);
        pDtDate := vDtAdd; // il faut changer la date passé, sinon décompte nb jour planifié faux
      end;
    end;

  begin
    if pBoCompteur then
    begin
      if (LUN in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 0);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
      if (MAR in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 1);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
      if (MER in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 2);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
      if (JEU in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 3);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
      if (VEN in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 4);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
      if (SAM in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 5);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
      if (DIM in fListeJours) and (pInNbJourAPlanifier > 0) then
      begin
        AddJour(pDtDate, pListeJours, 6);
        if pDtdate >= fDtDateDeb then
          pInNbJourAPlanifier := pInNbJourAPlanifier - 1;
      end;
    end
    else
    begin
      if (LUN in fListeJours) then
        AddJour(pDtDate, pListeJours, 0);
      if (MAR in fListeJours) then
        AddJour(pDtDate, pListeJours, 1);
      if (MER in fListeJours) then
        AddJour(pDtDate, pListeJours, 2);
      if (JEU in fListeJours) then
        AddJour(pDtDate, pListeJours, 3);
      if (VEN in fListeJours) then
        AddJour(pDtDate, pListeJours, 4);
      if (SAM in fListeJours) then
        AddJour(pDtDate, pListeJours, 5);
      if (DIM in fListeJours) then
        AddJour(pDtDate, pListeJours, 6);
    end;
  end;

begin
  result := TListDate.create;
  vInNbJour := pInNbJour;
  vBoCompteur := not (vInNbJour = cInNoCompteur);

  case fInMoisMethode of

    // numéro du jour
    cInMoisMethode1:
      begin
        if fBoGeneAuPlusTot then
        begin
          // positionnement début du mois
          vDtCurDate := DebutDeMois(fDtDateDeb);
          vDtTrav := vDtCurDAte; //mcd 01/07/2005
          // positionnement sur le jour
          vDtCurDate := PlusDate(vDtCurDate, fInMoisJourFixe - 1, 'J');
          if vDtCurDate > FinDeMois(vDtTrav) then
            vDtCurDate := FinDeMois(vDtTrav); //mcd 01/07/2005 si jour fixe le 31? ilf aut rester au 30 du mois et pas passer au 01 du mois suiv
          if (vDtCurDate <= fDtDateFin) and (vDtCurDate >= fDtDateDeb) and ((not vBoCompteur) or (vInNbJour > 0)) then
          begin
            result.Add(vDtCurDate);
            if vBoCompteur then
              vInNbJour := vInNbJour - 1;
          end;
          ii := 0;
          while ii = 0 do
          begin
            vDtCurDate := PlusDate(vDtCurDate, fInMoisFixe, 'M');
            vDtCurDate := DebutDeMois(vDtCurDate);
            vDtTrav := vDtCurDAte; //mcd 01/07/2005
            vDtCurDate := PlusDate(vDtCurDate, fInMoisJourFixe - 1, 'J');
            if vDtCurDate > FinDeMois(vDtTrav) then
              vDtCurDate := FinDeMois(vDtTrav); //mcd 01/07/2005 si jour fixe le 31? ilf aut rester au 30 du mois et pas passer au 01 du mois suiv
            if (vDtCurDate <= fDtDateFin) and ((not vBoCompteur) or (vInNbJour > 0)) then
            begin
              result.Add(vDtCurDate);
              if vBoCompteur then
                vInNbJour := vInNbJour - 1;
            end
            else
              ii := 1; // pour sortir de la boucle
          end;
        end
        else
        begin
          // positionnement début du mois
          vDtCurDate := DebutDeMois(fDtDateFin);
          // positionnement sur le jour
          vDtCurDate := PlusDate(vDtCurDate, fInMoisJourFixe - 1, 'J');
          if (vDtCurDate >= fDtDateDeb) and (vDtCurDate <= fDtDateFin) and ((not vBoCompteur) or (vInNbJour > 0)) then
          begin
            result.Add(vDtCurDate);
            if vBoCompteur then
              vInNbJour := vInNbJour - 1;
          end;
          ii := 0;
          while ii = 0 do
          begin
            vDtCurDate := PlusDate(vDtCurDate, -fInMoisFixe, 'M');
            vDtCurDate := DebutDeMois(vDtCurDate);
            vDtCurDate := PlusDate(vDtCurDate, fInMoisJourFixe - 1, 'J');
            if (vDtCurDate >= fDtDateDeb) and ((not vBoCompteur) or (vInNbJour > 0)) then
            begin
              result.Add(vDtCurDate);
              if vBoCompteur then
                vInNbJour := vInNbJour - 1;
            end
            else
              ii := 1;
          end;
        end;
      end;

    // libellé du jour
    cInMoisMethode2:
      begin
        if fBoGeneAuPlusTot then
        begin
          vDtCurDate := fDtDateDeb;
          while (vDtCurDate <= fDtDateFin) and ((not vBoCompteur) or (vInNbJour > 0)) do
          begin
            // positionnement sur le mois
            vDtCurDate := DebutDeMois(vDtCurDate);
            // positionnement du jour dans la semaine
            vDtCurDate := RechJourSemaine(vDtCurDate);
            // positionnement sur la semaine
            vDtCurDate := PlusDate(vDtCurDate, StrToInt(fStMoisSemaine) - 1, 'S');

            // ajout d'un controle si le jour trouvé est dans la periode
            if (vDtCurDate >= fDtDateDeb) and (vDtCurDate <= fDtDateFin) and
              ((not vBoCompteur) or (vInNbJour > 0)) then
            begin
              result.Add(vDtCurDate);
              // test sur date deb déjà fait au dessu. inutile dans ce cas
              if vBoCompteur then
                vInNbJour := vInNbJour - 1;
            end;

            // deplacement de x mois
            vDtCurDate := PlusDate(vDtCurDate, fInMoisFixe, 'M');
          end;
        end
        else
        begin
          vDtCurDate := fDtDateFin;
          while (vDtCurDate >= fDtDateDeb) and ((not vBoCompteur) or (vInNbJour > 0)) do
          begin
            // positionnement sur le mois
            vDtCurDate := DebutDeMois(vDtCurDate);
            // positionnement du jour dans la semaine
            vDtCurDate := RechJourSemaine(vDtCurDate);
            // positionnement sur la semaine
            vDtCurDate := PlusDate(vDtCurDate, StrToInt(fStMoisSemaine) - 1, 'S');
            // ajout d'un controle si le jour trouvé est dans la periode
            if (vDtCurDate >= fDtDateDeb) and (vDtCurDate <= fDtDateFin) and
              ((not vBoCompteur) or (vInNbJour > 0)) then
            begin
              result.Add(vDtCurDate);
              if vBoCompteur then
                vInNbJour := vInNbJour - 1;
            end;
            // deplacement de x mois
            vDtCurDate := PlusDate(vDtCurDate, -fInMoisFixe, 'M');
          end;
        end;
      end;

    // semaine
    cInMoisMethode3:
      begin
        if fBoGeneAuPlusTot then
        begin
          vDtCurDate := fDtDateDeb;
          while (vDtCurDate <= fDtDateFin) and (vDtCurDate >= fDtDateDEb) and ((not vBoCompteur) or (vInNbJour > 0)) do
          begin
            // positionnement sur le mois
            vDtCurDate := DebutDeMois(vDtCurDate);
            AddJoursSemaine(vDtCurDate, result, vBoCompteur, vInNbJour);
            // deplacement de x mois
            vDtCurDate := PlusDate(vDtCurDate, fInMoisFixe, 'M');
          end;
        end
        else
        begin
          vDtCurDate := fDtDateFin;
          while (vDtCurDate >= fDtDateDeb) and (vDtCurDate <= fDtDateFin) and ((not vBoCompteur) or (vInNbJour > 0)) do
          begin
            // positionnement sur le mois
            vDtCurDate := DebutDeMois(vDtCurDate);
            AddJoursSemaine(vDtCurDate, result, vBoCompteur, vInNbJour);
            // deplacement de x mois
            vDtCurDate := PlusDate(vDtCurDate, -fInMoisFixe, 'M');
          end;
        end;
      end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/10/2002
Modifié le ... :
Description .. : Genération à interval de jours réguliers
Mots clefs ... :
*****************************************************************}

function TAFRegles.GenereListeQuotidienne(pInNbJour: Integer): TListDate;
var
  i: Integer;
  vBoCompteur: Boolean;

begin
  i := 0;
  result := TListDate.create;
  vBoCompteur := not (pInNbJour = cInNoCompteur);

  if fBoGeneAuPlusTot then
    while ((fDtDateDeb + (i * fInJourInterval)) <= fDtDateFin) and ((fDtDateDeb + (i * fInJourInterval)) >= fDtDateDeb) and ((not vBoCompteur) or (pInNbJour > 0)) do
    begin
      result.Add(fDtDateDeb + (i * fInJourInterval));
      if vBoCompteur then
        pInNbJour := pInNbJour - 1;
      i := i + 1;
    end
  else
    while ((fDtDateFin - (i * fInJourInterval)) >= fDtDateDeb) and ((fDtDateFin - (i * fInJourInterval)) <= fDtDateFin) and ((not vBoCompteur) or (pInNbJour > 0)) do
    begin
      result.Add(fDtDateFin - (i * fInJourInterval));
      if vBoCompteur then
        pInNbJour := pInNbJour - 1;
      i := i + 1;
    end
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/10/2002
Modifié le ... :
Description .. : toutes les x semaines les jours cochés
Mots clefs ... :
*****************************************************************}

function TAFRegles.GenereListeHebdomadaire(pInNbJour: Integer): TListDate;
var
  vDtCurDate: TDateTime;
  Year, Month, Day: Word;
  vBoCompteur: Boolean;
  vInNumSem: Integer;
  vInYear: Integer;

begin
  if FinSemaineInterv = 0 then
    fInSemaineInterv := 1; // mcd 13/05/05 sinon, si semaine et 0, boucle !!!!
  result := TListDate.create;
  vBoCompteur := not (pInNbJour = cInNoCompteur);
  if fBoGeneAuPlusTot then
  begin
    decodeDate(fDtDateDeb, Year, Month, Day);
    // C.B 14/10/2004
    // séparer les 2 fonctions
    vInNumSem := NumSemaine(fDtDateDeb, vInYear);
    vDtCurDate := PremierJourSemaine(vInNumSem, vInYear);
    if vBoCompteur then
      while (vDtCurDate <= fDtDateFin) and (pInNbJour > 0) do
      begin
        if (LUN in fListeJours) and ((vDtCurDate) >= fDtDateDeb) and ((vDtCurDate) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (MAR in fListeJours) and ((vDtCurDate + 1) >= fDtDateDeb) and ((vDtCurDate + 1) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 1);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (MER in fListeJours) and ((vDtCurDate + 2) >= fDtDateDeb) and ((vDtCurDate + 2) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 2);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (JEU in fListeJours) and ((vDtCurDate + 3) >= fDtDateDeb) and ((vDtCurDate + 3) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 3);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (VEN in fListeJours) and ((vDtCurDate + 4) >= fDtDateDeb) and ((vDtCurDate + 4) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 4);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (SAM in fListeJours) and ((vDtCurDate + 5) >= fDtDateDeb) and ((vDtCurDate + 5) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 5);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (DIM in fListeJours) and ((vDtCurDate + 6) >= fDtDateDeb) and ((vDtCurDate + 6) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 6);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        vDtCurDate := PlusDate(vDtCurDate, fInSemaineInterv, 'S');
      end
    else
      while (vDtCurDate <= fDtDateFin) do
      begin
        if (LUN in fListeJours) and ((vDtCurDate) >= fDtDateDeb) and ((vDtCurDate) <= fDtDateFin) then
          result.Add(vDtCurDate);
        if (MAR in fListeJours) and ((vDtCurDate + 1) >= fDtDateDeb) and ((vDtCurDate + 1) <= fDtDateFin) then
          result.Add(vDtCurDate + 1);
        if (MER in fListeJours) and ((vDtCurDate + 2) >= fDtDateDeb) and ((vDtCurDate + 2) <= fDtDateFin) then
          result.Add(vDtCurDate + 2);
        if (JEU in fListeJours) and ((vDtCurDate + 3) >= fDtDateDeb) and ((vDtCurDate + 3) <= fDtDateFin) then
          result.Add(vDtCurDate + 3);
        if (VEN in fListeJours) and ((vDtCurDate + 4) >= fDtDateDeb) and ((vDtCurDate + 4) <= fDtDateFin) then
          result.Add(vDtCurDate + 4);
        if (SAM in fListeJours) and ((vDtCurDate + 5) >= fDtDateDeb) and ((vDtCurDate + 5) <= fDtDateFin) then
          result.Add(vDtCurDate + 5);
        if (DIM in fListeJours) and ((vDtCurDate + 6) >= fDtDateDeb) and ((vDtCurDate + 6) <= fDtDateFin) then
          result.Add(vDtCurDate + 6);
        vDtCurDate := PlusDate(vDtCurDate, fInSemaineInterv, 'S');
      end
  end
  else
  begin
    // onse positionne sur la semaine precedente
    vDtCurDate := PlusDate(fDtDateFin, -1, 'S');
    decodeDate(vDtCurDate, Year, Month, Day);
    // C.B 14/10/2004
    // séparer les 2 fonctions
    vInNumSem := NumSemaine(vDtCurDate, vInYear);
    vDtCurDate := PremierJourSemaine(vInNumSem, vInYear);

    if vBoCompteur then
      while (vDtCurDate >= fDtDateDeb) do
      begin
        if (LUN in fListeJours) and ((vDtCurDate) >= fDtDateDeb) and ((vDtCurDate) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (MAR in fListeJours) and ((vDtCurDate + 1) >= fDtDateDeb) and ((vDtCurDate + 1) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 1);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (MER in fListeJours) and ((vDtCurDate + 2) >= fDtDateDeb) and ((vDtCurDate + 2) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 2);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (JEU in fListeJours) and ((vDtCurDate + 3) >= fDtDateDeb) and ((vDtCurDate + 3) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 3);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (VEN in fListeJours) and ((vDtCurDate + 4) >= fDtDateDeb) and ((vDtCurDate + 4) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 4);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (SAM in fListeJours) and ((vDtCurDate + 5) >= fDtDateDeb) and ((vDtCurDate + 5) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 5);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        if (DIM in fListeJours) and ((vDtCurDate + 6) >= fDtDateDeb) and ((vDtCurDate + 6) <= fDtDateFin) and (pInNbJour > 0) then
        begin
          result.Add(vDtCurDate + 6);
          if (vDtCurDate >= fDtDateDEb) then
            pInNbJour := pInNbJour - 1;
        end;
        vDtCurDate := PlusDate(vDtCurDate, -fInSemaineInterv, 'S');
      end
    else
      while (vDtCurDate >= fDtDateDeb) do
      begin
        if (LUN in fListeJours) and ((vDtCurDate) >= fDtDateDeb) and ((vDtCurDate) <= fDtDateFin) then
          result.Add(vDtCurDate);
        if (MAR in fListeJours) and ((vDtCurDate + 1) >= fDtDateDeb) and ((vDtCurDate + 1) <= fDtDateFin) then
          result.Add(vDtCurDate + 1);
        if (MER in fListeJours) and ((vDtCurDate + 2) >= fDtDateDeb) and ((vDtCurDate + 2) <= fDtDateFin) then
          result.Add(vDtCurDate + 2);
        if (JEU in fListeJours) and ((vDtCurDate + 3) >= fDtDateDeb) and ((vDtCurDate + 3) <= fDtDateFin) then
          result.Add(vDtCurDate + 3);
        if (VEN in fListeJours) and ((vDtCurDate + 4) >= fDtDateDeb) and ((vDtCurDate + 4) <= fDtDateFin) then
          result.Add(vDtCurDate + 4);
        if (SAM in fListeJours) and ((vDtCurDate + 5) >= fDtDateDeb) and ((vDtCurDate + 5) <= fDtDateFin) then
          result.Add(vDtCurDate + 5);
        if (DIM in fListeJours) and ((vDtCurDate + 6) >= fDtDateDeb) and ((vDtCurDate + 6) <= fDtDateFin) then
          result.Add(vDtCurDate + 6);
        vDtCurDate := PlusDate(vDtCurDate, -fInSemaineInterv, 'S');
      end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : calcul des dates d'interventions en fonction d'un
                 nombre d'intervention pour une période donnée
Mots clefs ... :
*****************************************************************}

function TAFRegles.GenereListeNbIntervention: TListDate;
var
  vInNbJours: Integer;
  vInInterval: Integer;
  i: Integer;

begin
  result := TListDate.create;
  vInNbJours := trunc(fDtDateFin - fDtDateDeb);
  vInInterval := vInNbJours div fInNbIntervention;
  if fBoGeneAuPlusTot then
    for i := 0 to fInNbIntervention - 1 do
      result.Add(fDtDateDeb + (i * vInInterval))
  else
    for i := 0 to fInNbIntervention - 1 do
      result.Add(fDtDateFin - (i * vInInterval));
end;

function TAFRegles.GenereListeJours(var pRdDuree: Double; var pInNbJoursDecal: Integer; var pStMethodeDecal: string; pInNbJour: Integer): TListDate;
begin

  case fInFrequence of
    cInAnnuelle: result := GenereListeAnnuelle(pInNbJour);
    cInMensuelle: result := GenereListeMensuelle(pInNbJour);
    cInQuotidienne: result := GenereListeQuotidienne(pInNbJour);
    cInHebdomadaire: result := GenereListeHebdomadaire(pInNbJour);
    cInNbInterv: result := GenereListeNbIntervention;
  else
    result := nil;
  end;
  pRdDuree := fRdQteInterv;
  pInNbJoursDecal := fInNbJoursDecal;
  pStMethodeDecal := fStMethodeDecal;
end;

function TAFRegles.QuantiteRessource: Integer;
begin
  result := fInQteRessource;
end;

{ TListDate }

function TListDate.Add(Item: TDateTime): Integer;
var OneRecDate : PInternalDate;
begin
	New (OneRecDate);
  OneRecDate^.TheDate := Item;
  Inherited Add(OneRecDate);
end;

procedure TListDate.Delete(Index: Integer);
var OneRecDate : PInternalDate;
begin
  OneRecDate := GetPdate (Index);
  Dispose(OneRecDate);
  Inherited Delete(Index);
end;

function TListDate.GetPDate(Indice: integer): PInternalDate;
begin
 Result := PInternalDate (Inherited Items[Indice]);
end;

function TListDate.GetItems(Indice: integer): TDateTime;
var OneRecDate : PInternalDate;
begin
 OneRecDate := PInternalDate (Inherited Items[Indice]);
 Result := OneRecDate^.TheDate;
end;

function TListDate.IndexOf(Item: TDateTime): Integer;
begin
  Result := 0;
  while (Result < Count) and ((PInternalDate(List^[Result])^.TheDate) <> Item) do
    Inc(Result);
  if Result = Count then
    Result := -1;
end;

procedure TListDate.SetItems(Indice: integer; const Value: TDateTime);
var OneRecDate : PInternalDate;
begin
 	OneRecDate := PInternalDate (Inherited Items[Indice]);
  OneRecDate^.TheDate:= Value;
end;

end.
