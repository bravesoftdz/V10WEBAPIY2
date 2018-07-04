{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 10/07/2003
Modifié le ... :   /  /    
Description .. : Gestion de la fidélité client
Mots clefs ... : FIDELITE
*****************************************************************}
unit UFidelite;

interface

uses
  {$IFNDEF EAGLCLIENT}
  DBTables,
  {$ENDIF}
  UTOB, HCtrls, HMsgBox, HEnt1, EntPGI,
  ParamSoc, SysUtils, TarifUtil, UTilPGI;

type Fidelite = class
  private
    Etab: string;
    Tiers: string;
    TiersParDefaut: string;
    Prog: string;
    FNumCarteInterne: string;
    NumCarteExterne: string;
    Cadeau: string; // Fidelité client
    StMethodeFidelite: string;
    NouvelleCarte: Boolean;
    ChargeArrondi: Boolean;
    FProgrammeOk: Boolean;
    InverseValeur: Boolean;
    OkPourValidation: Boolean;
    Simulation : Boolean; //Pour ne pas incrementer les compteurs
    FCumulAnterieur: Double;
    LastCumulPiece: Double;
    DateMaxFidelite : TDateTime;
    DateMaxLigneFid : TDateTime;
    TOBFID: TOB; // Carte de fidelité
    TOBFIDLIG: TOB; // Nouvelles lignes de fidelité
    TOBProg: TOB; // Programme de fidelité
    TOBPiece: TOB; // Piece avec ses lignes
    TOBArrondi: TOB; // Table des ARRONDI
    TOBRegle: TOB; // Les regles valides du programme
    procedure InitTOB;
    procedure LibereTOB;
    procedure ChargeVariableAvecTOBFID;
    function DemarrageFidelite: Boolean;

    function GetALaDemande: Boolean;
    function GetCumulFidelite: Double;
    function GetFidelite: Boolean;
    function GetMessageFidelite: string;
    function GetFideliteDansPiece: Boolean;

    function VerifClient: Boolean;
    procedure ChargeDocument(TOB_Piece: TOB);
    function ChargeProgramme: Boolean;
    function ProgrammeValide: Boolean;
    procedure ChargeRegleFidelite;
    procedure ChargeInfoCompl;

    procedure ArrondiValeur(var LaValeur: Double; StArrondi: string);
    function CalculChampTOBPiece(Champ: string): Double;
    function VerifCondition(TOBRegleFidelite: TOB): Boolean;
    procedure Methode(CodeMethode: string);
    procedure CreationLigneFidelite(TOBRegleFidelite: TOB; CodeMethode: string);
    function TraiteVente(TOBLigne: TOB): Boolean;
    function CalculValeurRetour(TOBRegleFidelite: TOB; LaValeur: Double;
      BCadeau: Boolean = False; SeuilMini: Boolean = False): string;
    function LimiteDate(DateCreation: TDateTime): string;
    function CalculCumulAnterieur: Double;
    function FidInPiece: integer;
    procedure DetailCadeau(Var Choix: integer; Var DMaxi: Double; Var StCodeArticle: String);

    procedure Insert_TOBFIDLIG(TOBRegleFidelite: TOB; LaValeur: Double; Ordre: Integer);
  public
    constructor Create; virtual;
    destructor Destroy; override;

    property DemarreALaDemande: Boolean read GetALaDemande;
    property CumulFidelite: Double read GetCumulFidelite write FCumulAnterieur;
    property Fidelite: Boolean read GetFidelite;
    property MessageFidelite: string read GetMessageFidelite;
    property FideliteDansPiece: boolean read GetFideliteDansPiece;
    property NumeroCarteInterne: string read FNumCarteInterne;
    property ProgrammeOk: boolean read FProgrammeOk;

    function CumulPiece(TOB_Piece: TOB; InverseSens: Boolean = False): Double;
    function CumulFideliteEnMoins(TOB_Piece: TOB): Double;
    function PrepareArgumentsInfoCarte(TOBTiers, TOBPiece : TOB): String;

    function LoadProgramme(ClientDefaut, Etablissement: string): Boolean;
    procedure FermeProgramme;
    function ChargeCarte(Client, ClientDefaut, Etablissement,
      NumCarteInt, NumCarteExt: string; InfoCompl: Boolean): Boolean;
    function NewCarte: Boolean;

    function FideliteCorrecte(TOB_Piece: TOB): Integer;
    procedure ValideFidelite(InverseSens: Boolean = False);
    procedure AnnuleFidelite(Old_Tiers: string);
    function AutoriseRemiseGlobale(Pourcentage,MontantRemise : Double; Var Msg : string; TypeRem : String) : Boolean;
  end;

function CreationNumeroCarteInterne(Etablissement : string): string;
function RecupChronoParamSoc(NomParamSoc: string): Integer;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Création de l'objet Fidelite
Mots clefs ... :
*****************************************************************}

constructor Fidelite.Create;
begin
  inherited Create;
  NouvelleCarte := False;
  FProgrammeOk := False;
  OkPourValidation := False;
  StMethodeFidelite := 'ATT';
  InitTOB;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Libération de l'objet Fidelite
Mots clefs ... :
*****************************************************************}

destructor Fidelite.Destroy;
begin
  LibereTOB;
  inherited Destroy;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Initialisation des TOB
Mots clefs ... :
*****************************************************************}

procedure Fidelite.InitTOB;
begin
  TOBFID := TOB.Create('FIDELITEENT', nil, -1);
  TOBFIDLIG := TOB.Create('New FIDELITELIG', nil, -1);
  TOBProg := TOB.Create('PARFIDELITE', nil, -1);
  TOBArrondi := TOB.Create('Table ARRONDI', nil, -1);
  TOBRegle := TOB.Create('Les Regles', nil, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Libération des TOB
Mots clefs ... :
*****************************************************************}

procedure Fidelite.LibereTOB;
begin
  if TOBFID <> nil then TOBFID.Free;
  if TOBFIDLIG <> nil then TOBFIDLIG.Free;
  if TOBProg <> nil then TOBProg.Free;
  if TOBArrondi <> nil then TOBArrondi.Free;
  if TOBRegle <> nil then TOBRegle.Free;
  TOBPiece := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 15/05/2003
Modifié le ... :   /  /
Description .. : Charge le programme de fidelité et vérifie qu'il est valide
Mots clefs ... :
*****************************************************************}

function Fidelite.LoadProgramme(ClientDefaut, Etablissement: string): Boolean;
var Continue: boolean;
  StSQL: string;
  Q_ETAB: TQuery;
begin
  Result := True;
  if (FProgrammeOk) and (ClientDefaut = TiersParDefaut) and (Etab = Etablissement) then Exit;

  TiersParDefaut := ClientDefaut;
  Etab := Etablissement; //Mettre l'établissement de la boutique

  //if (Etab = '') then PGIError(TraduireMemoire('Pas d''établissement renseigné'),'FIDELITE');
  StSQL := 'Select ET_PROGRAMME From ETABLISS Where ET_ETABLISSEMENT = "' + Etab + '"';
  Q_ETAB := OpenSQL(StSQL, True);
  if not Q_ETAB.EOF then Prog := Trim(Q_ETAB.FindField('ET_PROGRAMME').AsString);
  Ferme(Q_ETAB);
  //if (Prog = '') then PGIError(TraduireMemoire('Pas de programme associé à l''établissement'),'FIDELITE');
  Continue := Prog <> '';
  if Continue then Continue := ChargeProgramme;
  if Continue then Continue := ProgrammeValide;
  if Continue then ChargeRegleFidelite;
  FProgrammeOk := Continue;
  Result := Continue;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 06/05/2003
Modifié le ... :   /  /
Description .. : Chargement du programme
Mots clefs ... :
*****************************************************************}

function Fidelite.ChargeProgramme: Boolean;
begin
  Result := True;
  if TOBProg.GetValue('GFO_CODEFIDELITE') <> Prog then
    Result := TOBProg.SelectDB('"' + Prog + '"', nil, True);
end;

function Fidelite.VerifClient: Boolean;
begin
  if Tiers = '' then Result := False
  else if Tiers = TiersParDefaut then Result := False
  else Result := True;
end;

procedure Fidelite.ChargeDocument(TOB_Piece: TOB);
begin
  if TOB_Piece <> Nil then TOBPiece := TOB_Piece;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 06/05/2003
Modifié le ... :   /  /
Description .. : Vérifie que le programme est toujours valide
Mots clefs ... :
*****************************************************************}

function Fidelite.ProgrammeValide: Boolean;
begin
  if TOBProg.GetValue('GFO_FERME') = 'X' then
  begin
    Result := False;
    Exit;
  end;

  Result := True;
  // '001' Illimité
  if TOBProg.GetValue('GFO_TYPEDUREEFID') = '002' then // Date de fin
  begin
    if TOBProg.GetValue('GFO_DATEFIN') <= V_PGI.DateEntree then
    begin
      FermeProgramme;
      Result := False;
      Exit;
    end;
  end
  else if TOBProg.GetValue('GFO_TYPEDUREEFID') = '003' then // Nombre de jours
  begin
    if TOBProg.GetValue('GFO_DATECREATION') + TOBProg.GetValue('GFO_NBJOUR') <= V_PGI.DateEntree then
    begin
      FermeProgramme;
      Result := False;
      Exit;
    end;
  end;

end;

procedure Fidelite.ChargeRegleFidelite;
var StSQL: string;
  Q_Regle: TQuery;
begin
  //Charge toutes les règles valides du programme en TOB
  StSQL := 'Select GFR_CHAMP1,GFR_CHAMP2,GFR_CHAMP3,GFR_CHAMPSEUIL,GFR_LIEN1,' +
    'GFR_LIEN2,GFR_NBJOUR,GFR_OPER1,GFR_OPER2,GFR_OPER3,GFR_REGLEFIDELITE,' +
    'GFR_TYPEREGLEFID,GFR_VAL1,GFR_VAL2,GFR_VAL3 From PARREGLEFID ' +
    'Where GFR_CODEFIDELITE = "' + Prog + '" AND ' +
    'GFR_FERME = "-" AND GFR_DATEDEBPERIOD <= "' + UsDateTime(NowH) + '" AND ' +
    'GFR_DATEFINPERIOD >= "' + UsDateTime(NowH) + '" ' +
    'ORDER BY GFR_TYPEREGLEFID, GFR_DATEDEBPERIOD';
  Q_Regle := OpenSQL(StSQL, True);
  if not Q_Regle.EOF then
    TOBRegle.LoadDetailDB('PARREGLEFID Virtuel', '', '', Q_Regle, False);
  Ferme(Q_Regle);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 06/05/2003
Modifié le ... :   /  /
Description .. : Fermeture du programme
Mots clefs ... :
*****************************************************************}

procedure Fidelite.FermeProgramme;
begin
  if not FProgrammeOk then Exit;
  with TOBProg do
  begin
    PutValue('GFO_UTILISATEUR', V_PGI.User);
    PutValue('GFO_DATEMODIF', NowH);
    PutValue('GFO_FERME', 'X');
    UpdateDB;
  end;
end;

procedure Fidelite.ChargeInfoCompl;
var QueCetEtablissement, StSQL : String;
    QQ : TQuery;
begin
  if TOBProg.GetValue('GFO_BONSERVEUR') = '-' then QueCetEtablissement := ' AND GFI_ETABLISSEMENT = "' + Etab + '"'
  else QueCetEtablissement := '';

  StSQL := 'SELECT GFI_TYPELIGNEFIDEL, MAX(GFI_DATECREATION) AS MAXIDATE ' +
           'FROM FIDELITELIG WHERE GFI_NUMCARTEINT = "' + FNumCarteInterne + '"' +
           QueCetEtablissement + ' GROUP BY GFI_TYPELIGNEFIDEL';
  QQ := OpenSQL(StSQL, True);
  if not QQ.EOF then
  begin
    DateMaxLigneFid := QQ.FindField('MAXIDATE').AsDateTime;
    if QQ.FindField('GFI_TYPELIGNEFIDEL').AsString = '006' then
      DateMaxFidelite := DateMaxLigneFid;
  end
  else DateMaxLigneFid := iDate1900;
  Ferme(QQ);

  if DateMaxFidelite = iDate1900 then
  begin
    StSQL := 'SELECT MAX(GFI_DATECREATION) AS MAXIDATE FROM FIDELITELIG ' +
             'WHERE GFI_NUMCARTEINT = "' + FNumCarteInterne + '"' +
             QueCetEtablissement + ' AND GFI_TYPELIGNEFIDEL = "006"';
    QQ := OpenSQL(StSQL, True);
    if not QQ.EOF then DateMaxFidelite := QQ.FindField('MAXIDATE').AsDateTime;
    if DateMaxFidelite = 0 then DateMaxFidelite := iDate1900;
    Ferme(QQ);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Chargement des informations de la carte du client en TOB
Mots clefs ... :
*****************************************************************}

function Fidelite.ChargeCarte(Client, ClientDefaut, Etablissement,
  NumCarteInt, NumCarteExt: string; InfoCompl: Boolean): Boolean;
var StSQL: string;
  Q_FIDELENTETE: TQuery;
begin
  Result := False;

  if not LoadProgramme(ClientDefaut, Etablissement) then Exit;

  StSQL := '';
  NouvelleCarte := True;
  if (Client = Tiers) and (NumCarteInt='') then
  begin
    Result := True;
    Exit;
  end;
  Tiers := Client;
  if (not VerifClient) and (NumCarteInt = '') and (NumCarteExt = '') then Exit;

  LastCumulPiece := 0;
  DateMaxFidelite := iDate1900;
  DateMaxLigneFid := iDate1900;
  FCumulAnterieur := 0;
  FNumCarteInterne := NumCarteInt;
  NumCarteExterne := NumCarteExt;

  if (Tiers <> '') then
    StSQL := 'SELECT * FROM FIDELITEENT WHERE ' +
      'GFE_TIERS="' + Tiers + '" ' + //AND GFE_ETABLISSEMENT="'+Etab+'" '+
    'AND GFE_FERME="-" AND GFE_PROGRAMME="' + Prog + '"'
  else if FNumCarteInterne <> '' then
    StSQL := 'SELECT * FROM FIDELITEENT WHERE ' +
      'GFE_NUMCARTEINT="' + FNumCarteInterne + '" AND GFE_FERME="-"'
  else if NumCarteExterne <> '' then
    StSQL := 'SELECT * FROM FIDELITEENT WHERE ' +
      'GFE_NUMCARTEEXT="' + NumCarteExterne + '" AND GFE_FERME="-"';

  if StSQL <> '' then
  begin
    Q_FIDELENTETE := OpenSQL(StSQL, FALSE);
    if not Q_FIDELENTETE.EOF then
    begin
      TOBFID.SelectDB('', Q_FIDELENTETE);
      NouvelleCarte := False;
      ChargeVariableAvecTOBFID;
      if InfoCompl then ChargeInfoCompl;
    end;
    Ferme(Q_FIDELENTETE);
  end;

  Result := not NouvelleCarte;
end;

function Fidelite.GetALaDemande: Boolean;
begin
  Result := TOBProg.GetValue('GFO_DEMARDEMANDE') = 'X';
end;

function Fidelite.GetCumulFidelite: Double;
begin
  if FCumulAnterieur <> 0 then Result := FCumulAnterieur
  else Result := CalculCumulAnterieur;
end;

function Fidelite.GetFidelite: Boolean;
var TOBRegleFidelite: TOB;
    CumulTotal : Double;
begin
  Cadeau := '0';
  TOBRegleFidelite := TOBRegle.FindFirst(['GFR_TYPEREGLEFID'], ['006'], False);
  if TOBRegleFidelite <> nil then
  begin
    LastCumulPiece := CumulPiece(Nil);
    CumulTotal := GetCumulFidelite + LastCumulPiece;
    if (CumulTotal > 0) and VerifCondition(TOBRegleFidelite) then
      Cadeau := CalculValeurRetour(TOBRegleFidelite, CumulTotal, True);
  end;
  if Cadeau = '0' then Cadeau := '';
  Result := Cadeau <> '';
end;

function Fidelite.GetFideliteDansPiece: Boolean;
begin
  Result := StMethodeFidelite = 'OUI';
end;

function Fidelite.GetMessageFidelite: string;
var StMsg, Critere, Champ, Val: string;
  PosEgal: integer;
begin
  if Cadeau = '' then Result := TraduireMemoire('Le client n''a pas atteint le seuil de fidelité.')
  else
  begin
    StMsg := Cadeau;
    repeat
      Critere := Trim(ReadTokenSt(StMsg));
      if Critere <> '' then
      begin
        PosEgal := pos('=', Critere);
        if PosEgal <> 0 then
        begin
          Champ := copy(Critere, 1, PosEgal - 1);
          Val := copy(Critere, PosEgal + 1, length(Critere));
          if Champ = 'CODEARTICLE' then Result := Val
          else if Champ = 'QTEFACT' then
            Result := Val + TraduireMemoire(' article(s) avec le code article ') + Result
          else if Champ = 'CONSTANTE' then
            Result := TraduireMemoire('une valeur de ') + Val + TraduireMemoire(' euros')
            //+ RechDom('TTDEVISETOUTES', V_PGI.DevisePivot, False)
          else if Champ = 'COEFF' then
            Result := Val + '%' ;
        end;
      end;
    until Critere = '';
    Result := TraduireMemoire('Le client peut bénéficier de sa fidelité : ') + Result;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Vérifie les conditions de démarrage d'une carte de fidelité
Mots clefs ... :
*****************************************************************}

function Fidelite.DemarrageFidelite: Boolean;
begin
  Result := False;
  if TOBProg.GetValue('GFO_DEMARCLIENT') = 'X' then Result := True;
  //else TOBProg.GetValue('GFO_DEMARDEMANDE') = 'X'
  //else TOBProg.GetValue('GFO_DEMARREGLE') = 'X'

end;

function RecupChronoParamSoc(NomParamSoc: string): Integer;
var iRecupChrono, iChrono: integer;
  stSQL: string;
  bOkChrono: Boolean;
begin
  iRecupChrono := 0;
  repeat
    inc(iRecupChrono);
    iChrono := GetParamSoc(NomParamSoc, True);
    inc(iChrono);
    stSQL := 'update PARAMSOC set SOC_DATA="' + IntToStr(iChrono) + '" ' +
      'WHERE SOC_NOM="' + NomParamSoc + '" and SOC_DATA="' + IntToStr(iChrono - 1) + '"';
    bOkChrono := (ExecuteSQL(stSQL) = 1);
  until ((bOkChrono) or (iRecupChrono > 10));

  if bOkChrono then Result := iChrono else Result := -1;
end;

function CreationNumeroCarteInterne(Etablissement : string): string;
var Etab3, Chrono14: string;
  Chrono: integer;
begin
  Chrono := RecupChronoParamSoc('SO_COMPTEURCARTEFID');

  Etab3 := '';
  if Length(Etablissement) < 3 then Etab3 := StringOfChar('0', 3 - Length(Etablissement));
  Etab3 := Etablissement + Etab3;

  Chrono14 := '';
  if Length(IntToStr(Chrono)) < 14 then Chrono14 := StringOfChar('0', 14 - Length(IntToStr(Chrono)));
  Chrono14 := Chrono14 + IntToStr(Chrono);

  if Chrono <> -1 then Result := Etab3 + Chrono14
  else Result := Etab3 + 'ERREUR';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Création d'une carte de fidelité
Mots clefs ... :
*****************************************************************}

function Fidelite.NewCarte: Boolean;
begin
  Result := False;

  if not VerifClient then Exit;
  if not FProgrammeOk then Exit;
  if not DemarrageFidelite then Exit;

  with TOBFID do
  begin
    PutValue('GFE_TIERS', Tiers);
    PutValue('GFE_ETABLISSEMENT', Etab);
    PutValue('GFE_FERME', '-');
    PutValue('GFE_DATEFERMETURE', iDate1900);
    PutValue('GFE_NUMCARTEINT', CreationNumeroCarteInterne(Etab));
    PutValue('GFE_NUMCARTEEXT', '');
    PutValue('GFE_PROGRAMME', Prog);
    PutValue('GFE_DATECREATION', NowH);
    PutValue('GFE_DATEMODIF', NowH);
    PutValue('GFE_DATEINTEGR', iDate1900);
    PutValue('GFE_CREATEUR', V_PGI.User);
    PutValue('GFE_UTILISATEUR', V_PGI.User);
    PutValue('GFE_DATEOUVERTURE', NowH);
    Result := InsertDB(nil);
  end;

  if Result then ChargeVariableAvecTOBFID;
end;

procedure Fidelite.ArrondiValeur(var LaValeur: Double; StArrondi: string);
var QQ: TQuery;
  TOBArr: TOB;
begin
  if StArrondi = '' then exit;
  // Charge la TOB des Arrondi
  if not ChargeArrondi then
  begin
    QQ := OpenSQL('SELECT * FROM ARRONDI Where GAR_RANG=1 ORDER BY GAR_CODEARRONDI,GAR_VALEURSEUIL DESC', TRUE);
    if not QQ.EOF then
    begin
      TOBArrondi.LoadDetailDB('ARRONDI', '', '', QQ, False);
      ChargeArrondi := True;
    end;
    Ferme(QQ);
  end;

  TOBArr := TOBArrondi.FindFirst(['GAR_CODEARRONDI'], [StArrondi], False);
  if TOBArr = nil then Exit;
  //Supprimé car il y a déjà une notion de valeur maximum dans les seuils de fidelité
  //if LaValeur > TOBArr.GetValue('GAR_VALEURSEUIL') then Exit;
  LaValeur := CalculArrondi(TOBArr, LaValeur);

  {  if (StArrondi = '') or (StArrondi = '0') then Exit
    else if StArrondi = '01' then LaValeur := ARRONDI(LaValeur, 1)
    else if StArrondi = '1' then LaValeur := Round(LaValeur)
    else if StArrondi = '10' then LaValeur := Round(LaValeur / 10) * 10
    else if StArrondi = '100' then LaValeur := Round(LaValeur / 100) * 100
    else if StArrondi = 'K' then LaValeur := Round(LaValeur / 1000) * 1000
    else if StArrondi = 'M' then LaValeur := Round(LaValeur / 1000000) * 1000000;}
end;

function Fidelite.CalculValeurRetour(TOBRegleFidelite: TOB; LaValeur: Double;
  BCadeau: Boolean = False; SeuilMini: Boolean = False): string;
var StSQL, CodeArt: string;
  Q_Seuil: TQuery;
  Qte, Coeff, Constante, ValeurMaxi, LaValeurABS, NewValeur: Double;
begin
  //Chargement des seuils de déclenchement
  Result := '0';
  LaValeurABS := ABS(LaValeur);
  StSQL := 'SELECT GFS_ARRONDI,GFS_CODEARTICLE,GFS_COEF,GFS_CONSTANTE,' +
    'GFS_QTEFACT,GFS_VALEURMAXIFID,GFS_SEUILMINI FROM PARSEUILFID ' +
    'Where GFS_CODEFIDELITE = "' + Prog + '" AND ' +
    'GFS_REGLEFIDELITE = "' + TOBRegleFidelite.GetValue('GFR_REGLEFIDELITE') + '" ' +
    'AND GFS_SEUILMINI <= ' + StrFPoint(LaValeurABS) + ' AND GFS_SEUILMAXI > ' + StrFPoint(LaValeurABS);
  Q_Seuil := OpenSQL(StSQL, True);
  if not Q_Seuil.EOF then
  begin
    if SeuilMini then Result := FloatToStr(-Q_Seuil.FindField('GFS_SEUILMINI').AsFloat)
    else if BCadeau then
    begin
      CodeArt := Q_Seuil.FindField('GFS_CODEARTICLE').AsString;
      Qte := Q_Seuil.FindField('GFS_QTEFACT').AsFloat;
      Coeff := Q_Seuil.FindField('GFS_COEF').AsFloat;
      Constante := Q_Seuil.FindField('GFS_CONSTANTE').AsFloat;
      if (CodeArt <> '') and (Qte <> 0) then
        Result := 'CODEARTICLE=' + CodeArt + ';QTEFACT=' + FloatToStr(Qte)
      else if Constante <> 0 then
        Result := 'CONSTANTE=' + FloatToStr(Constante)
      else if Coeff <> 0 then
        Result := 'COEFF=' + FloatToStr(Coeff);
    end
    else
    begin
      Coeff := Q_Seuil.FindField('GFS_COEF').AsFloat;
      Constante := Q_Seuil.FindField('GFS_CONSTANTE').AsFloat;
      ValeurMaxi := Q_Seuil.FindField('GFS_VALEURMAXIFID').AsFloat;

      NewValeur := Constante + (LaValeurABS * Coeff);

      if (ValeurMaxi <> 0) and (NewValeur > ValeurMaxi) then NewValeur := ValeurMaxi
      else ArrondiValeur(NewValeur, Q_Seuil.FindField('GFS_ARRONDI').AsString);

      if LaValeur < 0 then NewValeur := -NewValeur;

      Result := FloatToStr(NewValeur);
    end;
  end;
  Ferme(Q_Seuil);
end;

function Fidelite.LimiteDate(DateCreation: TDateTime): string;
var DateDebut: TDateTime;
  TOBRegleFidelite: TOB;
  NbJour : Integer;
begin
  Result := '';
  DateDebut := iDate1900;
  //Sélectionne la règle correspondante à la méthode (Bénéfice )
  TOBRegleFidelite := TOBRegle.FindFirst(['GFR_TYPEREGLEFID'], ['006'], False);
  if TOBRegleFidelite <> nil then
  begin
    NbJour := TOBRegleFidelite.GetValue('GFR_NBJOUR');
    if NbJour = 0 then Exit;
    DateDebut := NowH - NbJour;
  end;
  if DateCreation > DateDebut then DateDebut := DateCreation;
  if DateDebut <> iDate1900 then Result := ' AND GFI_DATECREATION > "' + UsTime(DateDebut) + '"';
end;

function Fidelite.CalculCumulAnterieur: Double;
var StSQL, QueCetEtablissement: string;
  Q_Cumul: TQuery;
  DateCreation: TDateTime;
begin
  Result := 0;
  if not VerifClient then Exit;
  if not FProgrammeOk then Exit;

  DateCreation := iDate1900;
  if TOBProg.GetValue('GFO_BONSERVEUR') = '-' then QueCetEtablissement := ' AND GFI_ETABLISSEMENT = "' + Etab + '"'
  else QueCetEtablissement := '';
  if TOBProg.GetValue('GFO_RAZCUMULFID') = 'X' then //Avec remise à zéro des cumuls
  begin
    //Recherche de la derniere remise à zéro
    StSQL := 'SELECT MAX(GFI_DATECREATION) AS MAXIDATE FROM FIDELITELIG ' +
      'WHERE GFI_NUMCARTEINT = "' + FNumCarteInterne + '"' +
      QueCetEtablissement + ' AND GFI_TYPELIGNEFIDEL = "006"';
    Q_Cumul := OpenSQL(StSQL, True);
    if not Q_Cumul.EOF then DateCreation := Q_Cumul.FindField('MAXIDATE').AsDateTime;
    Ferme(Q_Cumul);
  end;
  StSQL := 'SELECT SUM(GFI_VALEUR) AS CUMUL FROM FIDELITELIG ' +
    'WHERE GFI_NUMCARTEINT = "' + FNumCarteInterne + '"' +
    QueCetEtablissement + LimiteDate(DateCreation);
  Q_Cumul := OpenSQL(StSQL, True);
  if not Q_Cumul.EOF then FCumulAnterieur := Q_Cumul.FindField('CUMUL').AsFloat
  else FCumulAnterieur := 0;
  Ferme(Q_Cumul);

  Result := FCumulAnterieur;
end;

procedure Fidelite.Insert_TOBFIDLIG(TOBRegleFidelite: TOB; LaValeur: Double; Ordre: Integer);
var ValRetour: Double;
begin
  if LaValeur = 0 then Exit;
  if TOBRegleFidelite.GetValue('GFR_TYPEREGLEFID') = '006' then
    ValRetour := StrToFloat(CalculValeurRetour(TOBRegleFidelite, LaValeur, False, True))
  else ValRetour := StrToFloat(CalculValeurRetour(TOBRegleFidelite, LaValeur));
  if ValRetour = 0 then Exit;
  if InverseValeur then ValRetour := -ValRetour;
  with TOB.Create('FIDELITELIG', TOBFIDLIG, -1) do
  begin
    if Not Simulation then
    begin
      PutValue('GFI_CAISSE', TOBPiece.GetValue('GP_CAISSE'));
      PutValue('GFI_CREATEUR', V_PGI.User);
      PutValue('GFI_DATECREATION', NowH);
      PutValue('GFI_DATEINTEGR', iDate1900);
      PutValue('GFI_DATEMODIF', NowH);
      PutValue('GFI_ETABLISSEMENT', Etab);
      PutValue('GFI_LIGNE', RecupChronoParamSoc('SO_CHRONOFIDELLIG'));
      PutValue('GFI_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
      PutValue('GFI_NUMCARTEINT', FNumCarteInterne);
      PutValue('GFI_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
      PutValue('GFI_NUMORDRE', Ordre);
      PutValue('GFI_PROGRAMME', Prog);
      PutValue('GFI_REPRESENTANT', TOBPiece.GetValue('GP_REPRESENTANT'));
      PutValue('GFI_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
      PutValue('GFI_TYPECUMULFID', TOBProg.GetValue('GFO_TYPECUMULFID'));
      PutValue('GFI_TYPELIGNEFIDEL', TOBRegleFidelite.GetValue('GFR_TYPEREGLEFID'));
      PutValue('GFI_UTILISATEUR', V_PGI.User);
    end;
    PutValue('GFI_VALEUR', ValRetour);
  end;
end;

function Fidelite.TraiteVente(TOBLigne: TOB): Boolean;
var TypeRemise: string;
begin
  Result := True;
  if Pos(TOBLigne.GetValue('GL_TYPEARTICLE'), TOBProg.GetValue('GFO_TYPEARTICLEFID')) = 0 then
    Result := False //Le type d'article n'est pas accepté
  else if (TOBProg.GetValue('GFO_DEDUCRETOURFID') = '-') and (TOBLigne.GetValue('GL_TOTALTTC') < 0) then
    Result := False //Les retours ne sont pas acceptés
  else
  begin
    TypeRemise := TOBLigne.GetValue('GL_TYPEREMISE');
    if TypeRemise <> '' then
    begin
      if TypeRemise = TOBProg.GetValue('GFO_TYPEREMISE') then //Remise Fidelite
      begin
        if TOBProg.GetValue('GFO_FIDLIGREMFID') = '-' then Result := False; //La remise fidelité n'est pas prise en compte
      end
      else if TOBProg.GetValue('GFO_FIDLIGREMISE') = '-' then Result := False; //Les remises ne sont pas prise en compte
    end;
  end;
end;

procedure Fidelite.CreationLigneFidelite(TOBRegleFidelite: TOB; CodeMethode: string);
var iLigne, iNumChamp, iNumOrdre: integer;
  QteOuMontant, Total: Double;
  Champ: string;
begin
  Total := 0;
  if CodeMethode = '006' then Insert_TOBFIDLIG(TOBRegleFidelite, GetCumulFidelite + LastCumulPiece, 0)
  else
  begin
    Champ := RechDom('GCCHAMPFIDELITE', TOBRegleFidelite.GetValue('GFR_CHAMPSEUIL'), True);
    iNumChamp := 0;
    iNumOrdre := 0;
    if TOBPiece.Detail.Count > 0 then
    begin
      iNumChamp := TOBPiece.Detail[0].GetNumChamp(Champ);
      iNumOrdre := TOBPiece.Detail[0].GetNumChamp('GL_NUMORDRE');
    end;
    for iLigne := 0 to TOBPiece.Detail.count - 1 do
    begin
      if not TraiteVente(TOBPiece.Detail[iLigne]) then Continue;
      QteOuMontant := TOBPiece.Detail[iLigne].GetValeur(iNumChamp);
      if CodeMethode = '005' then //Si règle "Vente"
        Insert_TOBFIDLIG(TOBRegleFidelite, QteOuMontant, TOBPiece.Detail[iLigne].GetValeur(iNumOrdre))
      else Total := Total + QteOuMontant;
    end;
    Insert_TOBFIDLIG(TOBRegleFidelite, Total, 0); // Total = 0 si la règle est "Vente"
  end;
end;

function Fidelite.CalculChampTOBPiece(Champ: string): Double;
var iLigne: integer;
begin
  Result := 0;
  for iLigne := 0 to TOBPiece.Detail.count - 1 do
  begin
    if not TraiteVente(TOBPiece.Detail[iLigne]) then Continue;
    Result := Result + TOBPiece.Detail[iLigne].GetValue(Champ);
  end;
end;

function GoodCondition(Val, Ref: Double; Operateur: string): Boolean;
begin
  Result := True;
  if Operateur = 'EGA' then Result := Val = Ref
  else if Operateur = 'IEG' then Result := Val <= Ref
  else if Operateur = 'INF' then Result := Val < Ref
  else if Operateur = 'SEG' then Result := Val >= Ref
  else if Operateur = 'SUP' then Result := Val > Ref
    ;
end;

function GoodLesTrois(Ok1, Ok2, Ok3: Boolean; Lien1, Lien2: string): Boolean;
begin
  if (Lien1 = '') or (Lien1 = '3') then Result := Ok1
  else if (Lien2 = '') or (Lien2 = '3') then
  begin
    if Lien1 = '1' then Result := Ok1 and Ok2
    else Result := Ok1 or Ok2;
  end
  else
  begin
    if (Lien1 = '1') then
    begin
      if (Lien2 = '1') then Result := Ok1 and Ok2 and Ok3
      else Result := Ok1 and Ok2 or Ok3;
    end
    else
    begin
      if (Lien2 = '1') then Result := Ok1 or Ok2 and Ok3
      else Result := Ok1 or Ok2 or Ok3;
    end;
  end;
end;

function Fidelite.VerifCondition(TOBRegleFidelite: TOB): Boolean;
var CHAMP1, CHAMP2, CHAMP3: string;
  LaValeur: Double;
  Ok1, Ok2, Ok3: Boolean;
begin
  Result := True;
  Ok2 := False;
  Ok3 := False;
  CHAMP1 := NullToVide(TOBRegleFidelite.GetValue('GFR_CHAMP1'));
  if CHAMP1 = '' then Exit;

  LaValeur := CalculChampTOBPiece(RechDom('GCCHAMPFIDELITE', CHAMP1, True));
  Ok1 := GoodCondition(ABS(LaValeur),
    TOBRegleFidelite.GetValue('GFR_VAL1'),
    NullToVide(TOBRegleFidelite.GetValue('GFR_OPER1')));

  CHAMP2 := NullToVide(TOBRegleFidelite.GetValue('GFR_CHAMP2'));
  if CHAMP2 <> '' then
  begin
    if CHAMP2 <> CHAMP1 then LaValeur := CalculChampTOBPiece(RechDom('GCCHAMPFIDELITE', CHAMP2, True));
    Ok2 := GoodCondition(ABS(LaValeur),
      TOBRegleFidelite.GetValue('GFR_VAL2'),
      NullToVide(TOBRegleFidelite.GetValue('GFR_OPER2')));

    CHAMP3 := NullToVide(TOBRegleFidelite.GetValue('GFR_CHAMP3'));
    if CHAMP3 <> '' then
    begin
      if CHAMP3 <> CHAMP2 then LaValeur := CalculChampTOBPiece(RechDom('GCCHAMPFIDELITE', CHAMP3, True));
      Ok3 := GoodCondition(ABS(LaValeur),
        TOBRegleFidelite.GetValue('GFR_VAL3'),
        NullToVide(TOBRegleFidelite.GetValue('GFR_OPER3')));
    end;
  end;

  Result := GoodLesTrois(Ok1, Ok2, Ok3,
    NullToVide(TOBRegleFidelite.GetValue('GFR_OPER1')),
    NullToVide(TOBRegleFidelite.GetValue('GFR_OPER2')));
end;

procedure Fidelite.Methode(CodeMethode: string);
var TOBRegleFidelite: TOB;
begin
  TOBRegleFidelite := TOBRegle.FindFirst(['GFR_TYPEREGLEFID'], [CodeMethode], False);
  if (TOBRegleFidelite <> nil) then
    if VerifCondition(TOBRegleFidelite) then CreationLigneFidelite(TOBRegleFidelite, CodeMethode);
end;

function Fidelite.AutoriseRemiseGlobale(Pourcentage,MontantRemise : Double; Var Msg : string; TypeRem : String) : Boolean;
Var Choix: integer;
    DMaxi: Double;
    StCodeArticle: String;
begin
  Result := True;

  if TypeRem <> TOBProg.GetValue('GFO_TYPEREMISE') then Exit; //Remise Fidelite

  DetailCadeau(Choix, DMaxi, StCodeArticle);
  if Choix = 1 then
  begin
    Result := False;
    Msg := TraduireMemoire('Impossible car la remise correspond à un article');
  end
  else if ((Choix = 2) or (Choix = 3)) and (MontantRemise > DMaxi) then
  begin
    Result := False;
    Msg := TraduireMemoire('Le client n''a pas atteint le seuil de fidelité ');
    Case Choix of
      2 : Msg := Msg + '(' + FloatToStr(DMaxi) + ' Euros)';
      3 : Msg := Msg + '(' + FloatToStr(DMaxi) + ' %)';
    end;
  end;
end;

procedure Fidelite.DetailCadeau(Var Choix: integer; Var DMaxi: Double; Var StCodeArticle: String);
var StMsg, Champ, Val, Critere : String;
    PosEgal : Integer;
begin
  StMsg := Cadeau;
  Choix := 0;
  DMaxi := 0;
  StCodeArticle := '';
  repeat
    Critere := Trim(ReadTokenSt(StMsg));
    if Critere <> '' then
    begin
      PosEgal := pos('=', Critere);
      if PosEgal <> 0 then
      begin
        Champ := copy(Critere, 1, PosEgal - 1);
        Val := copy(Critere, PosEgal + 1, length(Critere));
        if Champ = 'CODEARTICLE' then
        begin
          StCodeArticle := Val;
          Choix := 1;
        end
        else if Champ = 'QTEFACT' then DMaxi := StrToFloat(Val)
        else if Champ = 'CONSTANTE' then
        begin
          DMaxi := StrToFloat(Val);
          Choix := 2;
        end
        else if Champ = 'COEFF' then
        begin
          DMaxi := StrToFloat(Val);
          Choix := 3;
        end;
      end;
    end;
  until Critere = '';
end;

function Fidelite.FidInPiece: integer;
var StCodeArticle, TypeRemise: string;
  Choix, iLigne: Integer;
  DMaxi, DTemp, TotalTTC, TotalRemise: Double;
  TOBLigne: TOB;
begin
  Result := 0; //Pas de fidelité dans la pièce

  DetailCadeau(Choix, DMaxi, StCodeArticle);

  DTemp := 0;
  TotalTTC := 0;
  TotalRemise := 0;
  TypeRemise := TOBProg.GetValue('GFO_TYPEREMISE'); //Remise Fidelite

  for iLigne := 0 to TOBPiece.Detail.count - 1 do
  begin
    TOBLigne := TOBPiece.Detail[iLigne];
    if TOBLigne.GetValue('GL_TYPEREMISE') = TypeRemise then
    begin
      Result := 1; //Fidelité dans la pièce
      case Choix of
        1: begin // 1 ou plusieurs articles offerts
            if TOBLigne.GetValue('GL_CODEARTICLE') <> StCodeArticle then
            begin
              Result := 2;
              Break;
            end //Fidélité dépassée
            else DTemp := DTemp + TOBLigne.GetValue('GL_QTEFACT');
           end;
        2: DTemp := DTemp + TOBLigne.GetValue('GL_TOTREMLIGNEDEV'); // Un montant offert
        3: begin
            TotalTTC := TotalTTC + (TOBLigne.GetValue('GL_PUTTCDEV') * TOBLigne.GetValue('GL_QTEFACT'));
            TotalRemise := TotalRemise + TOBLigne.GetValue('GL_TOTREMLIGNEDEV');
           end;
      end;
      if DTemp > DMaxi then
      begin
        Result := 2;
        Break;
      end //Fidélité dépassée
    end;
  end;

  if Choix = 3 then
  begin
    if ((TotalRemise / TotalTTC * 100) > DMaxi) then Result := 2; //Fidélité dépassée
  end;
end;

function Fidelite.FideliteCorrecte(TOB_Piece: TOB): Integer;
const erFidDepasse = 31;
var IntError: integer;
    AvecFidelite : boolean;
begin
  Result := 0;
  OkPourValidation := False;
  StMethodeFidelite := 'ATT';
  if not VerifClient then Exit;
  if not FProgrammeOk then Exit;

  ChargeDocument(TOB_Piece);
  //Vérifie que ce type de pièce est géré dans le programme de fidelité
  if Pos(TOBPiece.GetValue('GP_NATUREPIECEG'), TOBProg.GetValue('GFO_NATUREPIECEFID')) = 0 then Exit;

  // Reste à vérifier que la fidelité est utilisée dans la Piece
  AvecFidelite := GetFidelite;
  IntError := FidInPiece;
  if IntError = 0 then StMethodeFidelite := 'NON'
  else if (IntError = 1) and AvecFidelite then StMethodeFidelite := 'OUI'
  else //if IntError = 2 then
  begin
    StMethodeFidelite := 'TRO';
    Result := erFidDepasse;
    Exit;
  end;

  OkPourValidation := True;
end;

procedure Fidelite.ValideFidelite(InverseSens: Boolean = False);
begin
  if not OkPourValidation then Exit;

  InverseValeur := InverseSens;
  TOBFIDLIG.ClearDetail;

  Simulation := False;
  //Traite les règles de type Ticket
  Methode('004');
  //Traite les règles de type Vente
  Methode('005');
  //Traite les règles de type Fidelité
  if StMethodeFidelite = 'OUI' then Methode('006');

  if not TOBFIDLIG.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;

function Fidelite.CumulPiece(TOB_Piece: TOB; InverseSens: Boolean = False): Double;
begin
  Result := 0;
  if not VerifClient then Exit;
  if not FProgrammeOk then Exit;
  TOBFIDLIG.ClearDetail;
  ChargeDocument(TOB_Piece);
  if TOBPiece = Nil then Exit;

  InverseValeur := InverseSens;

  Simulation := True;
  //Traite les règles de type Ticket
  Methode('004');
  //Traite les règles de type Vente
  Methode('005');

  if TOBFIDLIG.Detail.Count = 0 then Result := 0
  else Result := TOBFIDLIG.Somme('GFI_VALEUR', [''], [''], TRUE);
end;

function Fidelite.CumulFideliteEnMoins(TOB_Piece: TOB): Double;
begin
  Result := 0;
  if not VerifClient then Exit;
  if not FProgrammeOk then Exit;
  TOBFIDLIG.ClearDetail;
  ChargeDocument(TOB_Piece);
  if TOBPiece = Nil then Exit;

  Simulation := True;
  //Traite les règles de type Fidelité
  if (FidInPiece = 1) then Methode('006');

  if TOBFIDLIG.Detail.Count = 0 then Result := 0
  else Result := TOBFIDLIG.Somme('GFI_VALEUR', [''], [''], TRUE);
end;

procedure Fidelite.AnnuleFidelite(Old_Tiers: string);
var New_Tiers: string;
begin
  if not OkPourValidation then Exit;

  New_Tiers := Tiers;
  if ChargeCarte(Old_Tiers, TiersParDefaut, Etab, '', '', False) then //Charge la carte de l'ancien client
    ValideFidelite(True);

  ChargeCarte(New_Tiers, TiersParDefaut, Etab, '', '', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Chargement des variables
Mots clefs ... :
*****************************************************************}

procedure Fidelite.ChargeVariableAvecTOBFID;
begin
  with TOBFID do
  begin
    Tiers := GetValue('GFE_TIERS');
    FNumCarteInterne := GetValue('GFE_NUMCARTEINT');
    NumCarteExterne := GetValue('GFE_NUMCARTEEXT');
  end;
end;

function Fidelite.PrepareArgumentsInfoCarte(TOBTiers, TOBPiece : TOB): String;
var CumulOld, CumulNew, CumulFidEnMoins : Double;
    DateFerm : TDateTime;
begin
  CumulOld := CumulFidelite;
  LastCumulPiece := CumulPiece(TOBPiece, False);
  CumulFidEnMoins := CumulFideliteEnMoins(TOBPiece);
  CumulNew := CumulOld + CumulFidEnMoins + LastCumulPiece;
  Fidelite;

  if TOBFID.GetValue('GFE_DATEFERMETURE') <> iDate1900 then DateFerm := TOBFID.GetValue('GFE_DATEFERMETURE')
  else if TOBProg.GetValue('GFO_TYPEDUREEFID') = '001' then DateFerm := iDate2099 // Illimite
  else if TOBProg.GetValue('GFO_TYPEDUREEFID') = '002' then DateFerm := TOBProg.GetValue('GFO_DATEFIN')// Date de fin
  else if TOBProg.GetValue('GFO_TYPEDUREEFID') = '003' then DateFerm := TOBProg.GetValue('GFO_DATECREATION') + TOBProg.GetValue('GFO_NBJOUR') // Nombre de jours
  else DateFerm := iDate2099;

  Result := 'NOM=' + TOBTiers.GetValue('T_LIBELLE')
  + ';PRENOM=' + TOBTiers.GetValue('T_PRENOM')
  + ';NUMINTERNE=' + FNumCarteInterne
  + ';NUMEXTERNE=' + NumCarteExterne
  + ';DATEDEMARRAGE=' + DateToStr(TOBFID.GetValue('GFE_DATEOUVERTURE'))
  + ';DATEFIN=' + DateToStr(DateFerm)
  + ';DATEVISITECLIENT=' + DateToStr(DateMaxLigneFid)
  + ';DATEREMISE=' + DateToStr(DateMaxFidelite)
  + ';CUMULOLD=' + FloatToStr(CumulOld)
  + ';CUMULPIECE=' + FloatToStr(LastCumulPiece)
  + ';CUMULFIDELITE=' + FloatToStr(CumulFidEnMoins)
  + ';CUMULNEW=' + FloatToStr(CumulNew)
  + ';CADEAU1=' + MessageFidelite
  + ';REPRESENTANT=' + TOBPiece.GetValue('GP_REPRESENTANT')
  ;
end;


end.
