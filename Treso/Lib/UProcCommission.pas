unit UProcCommission;

interface

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls , Hmsgbox, Hent1, Constantes, UTOB, Classes, UObjGen;

type
  {06/12/04 : TObjCommission a été créé avec la structure TrEcritures par soucis de compatibilité avec
              l'existant et surtout parce que je suis un gros paresseux qui n'a jamais eu le courage de
              repasser dans le code pour remplacer la structure par une TOB}
  TObjCommissionTob = class(TObject)
    {Tob contenant les ecritures générées pas l'objet}
    TobEcriture : TOB;
  private
    FGeneral   : string;
    FCodeFlux  : string;
    FDateOpe   : TDateTime;
    FCotation  : Double;
    FDevise    : string;

    {Read de la property ACommission}
    function IsSoumis : Boolean;
    {Reprise des valeurs communes à l'opératin de tréosrerie et à sa commission}
    procedure ReprendreChampsCommuns(var Source, Dest : TOB; Lib : string);
    {Cacul d'une commissions par pourcentage}
    function CalcParCom(Base, Com, Mini, Maxi : Double; DevF : string) : Double;
    {Calcul d'une commission par tranche}
    function CalcParTra(Base : Double; Tranche : string) : Double;
    {Calcul d'une commission par forfait}
    function CalcParFor(Forfait, Unite : Double; DevF : string) : Double;
    {Mise à jours des propriété FCotation et FDevise}
    procedure SetCotation(Devise : string);
  protected
    TobTransfo : TOB;

    {Vide la TobCommission des flux superflus}
    procedure PurgerTobCommision;
    {Calcul de la date de valeur de la commission : inspiré de l'unité échéance}
    procedure   CalculerDateValeur(T : TOB; var E : TOB);
    {Calcul du montant de la commission}
    procedure   CalculerMntCommission(T : TOB; var E : TOB);
  public
    NbDecimal : Integer; {25/05/05 : Gestion du format des montants}

    constructor Create(Cpte, Flux : string; dtOperation : TDateTime);
    destructor  Destroy; override;
    {Génération de l'écriture de trésorerie de la commission}
    procedure   GenererCommissions(var tEcritures : TOB); virtual;

    property ACommission : Boolean read IsSoumis;
    property General  : string     read FGeneral    write FGeneral;
    property CodeFlux : string     read FCodeFlux   write FCodeFlux;
    property DateOpe  : TDateTime  read FDateOpe    write FDateOpe;
  end;

  {29/11/04 : Pour les OPCVM, le traitement est un peu spécifique}
  TObjComOPCVM = class(TObjCommissionTob)
  private
    FAvecEcriture : Boolean;
    FMntTva       : Double;
    TobFC         : TOB;
    TobTva        : TOB;
    OkTVA         : Boolean;

    function  GetCommission : Double;
    function  GetFrais      : Double;
    procedure SetMntTva(var Value : Double; CodTva : string);
    procedure MajTobFC (tEcr : TOB);
  public
    EnVente : Boolean;

    constructor Create(Cpte, Transac : string; dtOperation : TDateTime; VenteOk, ReelOk : Boolean);
    destructor  Destroy; override;
    {Génération de l'écriture de trésorerie de la commission}
    procedure   GenererCommissions(var tEcr : TOB); override;

    property AvecEcriture  : Boolean read FAvecEcriture  write FAvecEcriture;
    property MntCommission : Double  read GetCommission;
    property MntFrais      : Double  read GetFrais;
    property MntTVA        : Double  read FMntTVA;
  end;

  TObjContrat = class
    lContrat : TStringList;

    {Remplit la liste lContrat en fonction des données de la Query}
    procedure RemplirListe(Q : TQuery);
  public
    constructor Create(DtDeb, DtFin : TDateTime; Gene : string = '');
    destructor  Destroy; override;
    {Renvoie le taux de la commission de mouvement pour un compte et une date donnés}
    function    GetComFromGene(Gene : string; Dt : TDateTime) : Double;
  end;


function SupprimeEcriture  (Titre, NumTransac, NumPiece, NumLigne, NoDossier : string) : Boolean;
function SupprimePiece     (Dos, NumTransac, NumPiece, ClefOpe : string) : Boolean;
function UpdateEcritureStr (Dos, NumTransac, NumPiece, NumLigne, ClefOpe , Chp, Val : string) : Boolean;
function UpdatePieceStr    (Dos, NumTransac, NumPiece, ClefOpe , Chp, Val : string) : string;

implementation

uses
  Commun, UProcGen, UProcEcriture, Controls {pour mrYes ...}
  {$IFNDEF EAGLSERVER}
  , AglInit {TheData}
  {$ENDIF EAGLSERVER}
  ;

{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}
{                                OBJET TOBJCONTRAT                                      }
{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TObjContrat.Create(DtDeb, DtFin : TDateTime; Gene : string = '');
{---------------------------------------------------------------------------------------}
var
  Requete : string;
  Q       : TQuery;
begin
  lContrat := TStringList.Create;
  lContrat.Duplicates := dupAccept;
  lContrat.Sorted     := True;

  Requete := 'SELECT TRC_COMMVT, TRC_DEBCONTRAT, TRC_FINCONTRAT, BQ_CODE FROM TRCONTRAT, BANQUECP WHERE ' +
             'BQ_AGENCE = TRC_AGENCE AND ' +
             'TRC_DEBCONTRAT <= "' + UsDateTime(DtFin) + '" AND TRC_FINCONTRAT >= "' + UsDateTime(DtDeb) +  '" ';
  if Gene <> '' then Requete := Requete + 'AND BQ_CODE = "' + Gene + '"';

  Q := OpenSQL(Requete, True);
  try
    if not Q.EOF then RemplirListe(Q);
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
destructor TObjContrat.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(lContrat) then LibereListe(lContrat, True);
  inherited Destroy;
end;

{Remplit la liste lContrat en fonction des données de la Query
{---------------------------------------------------------------------------------------}
procedure TObjContrat.RemplirListe(Q : TQuery);
{---------------------------------------------------------------------------------------}
var
  Obj : TObjDetContrat;
begin
  while not Q.EOF do begin
    Obj := TObjDetContrat.Create;
    try
      Obj.aDtDeb  := Q.FindField('TRC_DEBCONTRAT').AsDateTime;
      Obj.aDtFin  := Q.FindField('TRC_FINCONTRAT').AsDateTime;
      Obj.aComMvt := Q.FindField('TRC_COMMVT'    ).AsFloat;
      lContrat.AddObject(Q.FindField('BQ_CODE').AsString, Obj);
    except
      on E : Exception do FreeAndNil(Obj);
    end;
    Q.Next;
  end;
end;

{Renvoie le taux de la commission de mouvement pour un compte et une date donnés
{---------------------------------------------------------------------------------------}
function TObjContrat.GetComFromGene(Gene : string; Dt : TDateTime) : Double;
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  Obj : TObjDetContrat;
begin
  {Valeur commune, en général, pour les commissions de mouvement}
  Result := 0.025;
  for n := 0 to lContrat.Count - 1 do begin
    Obj := TObjDetContrat(lContrat.Objects[n]);
    if not Assigned(Obj) then Continue;
    if (lContrat[n] = Gene) and (Obj.aDtDeb <= Dt) and (Obj.aDtFin >= Dt) then begin
      Result := Obj.aComMvt;
      Break;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}
{                                FONCTIONS GLOBALES                                     }
{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}


{Suppression d'une écriture en tenant compte des commissions
{---------------------------------------------------------------------------------------}
function SupprimeEcriture(Titre, NumTransac, NumPiece, NumLigne, NoDossier : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;

    {------------------------------------------------------------------------}
    function GererComm : Boolean;
    {------------------------------------------------------------------------}
    var
      w : Word;
      s : string;
      o : Boolean;
      l : string;
    begin
      Result := True;
      s := 'L''écriture (numéro de transaction ' + NumTransac + ','#13 +
           '             numéro de pièce ' + NumPiece + ','#13 +
           '             numéro de ligne ' + NumLigne + ') ' +
           'est une écriture de commission.'#13;
      w := HShowMessage('4;' + Titre + ';' + s + 'Souhaitez-vous :'#13 +
                                                ' - Supprimer la commission (OUI) ?'#13 +
                                                ' - Supprimer la piéce entière, écriture et commissions (NON) ?'#13 +
                                                ' - Interrompre le traitement (ANNULER)?;Q;YNC;C;C;', '', '');

      case w of
        mrCancel : Result := False;
        mrYes    : begin
                     {On regarde si l'opération de la pièce a d'autres commissions}
                     o := False;
                     l := '1';
                     Q.First;
                     while not Q.EOF do begin
                       if Q.FindField('TE_COMMISSION').AsString = suc_AvecCom then
                         {Récupération du numéro de ligne de l'opération financière soumise à commission(théoriquement '1')}
                         l := Q.FindField('TE_NUMLIGNE').AsString
                       else if not o then
                         o := (Q.FindField('TE_NUMLIGNE').AsString <> NumLigne) and
                              (Q.FindField('TE_COMMISSION').AsString = suc_Commission);
                       Q.Next;
                     end;

                     {Suppression de la ligne}
                     ExecuteSQL('DELETE FROM TRECRITURE WHERE TE_NODOSSIER = "' + NoDossier + '" AND TE_NUMTRANSAC = "' +
                                NumTransac + '" AND TE_NUMEROPIECE = ' + NumPiece + ' AND TE_NUMLIGNE = ' + NumLigne);

                     {Aucune autre commission n'a été trouvée ...}
                     if not o then
                       {... On met à jour le champ TE_COMMISSION}
                       UpdateEcritureStr(NoDossier, NumTransac, NumPiece, l, '', 'TE_COMMISSION', suc_SansCom)
                   end;
                  {On supprimme la pièce, écriture en cours plus les éventuelles commissions afférentes}
        mrNo     : ExecuteSQL('DELETE FROM TRECRITURE WHERE TE_NODOSSIER = "' + NoDossier + '" AND TE_NUMTRANSAC = "' +
                              NumTransac + '"' + ' AND TE_NUMEROPIECE = ' + NumPiece);
      end;
    end;

begin
  Result := True;
  if NoDossier = '' then NoDossier := V_PGI.NoDossier;

  Q := OpenSQL('SELECT * FROM TRECRITURE WHERE TE_NODOSSIER = "' + NoDossier + '" AND TE_NUMTRANSAC = "' +
               NumTransac + '" AND TE_NUMEROPIECE = ' + NumPiece, True);
  try
    while not Q.EOF do begin
      if Q.FindField('TE_NUMLIGNE').AsString = NumLigne then begin
        if Q.FindField('TE_COMMISSION').AsString = suc_Commission then
          {s'il s'agit d'une écriture de commission}
          Result := GererComm
        else
          {Il ne s'agit pas d'une écriture de commission, on supprimme la pièce, écriture en cours plus les éventuelles
           commissions afférentes}
          ExecuteSQL('DELETE FROM TRECRITURE WHERE TE_NODOSSIER = "' + NoDossier + '" AND TE_NUMEROPIECE = ' +
                     NumPiece + ' AND TE_NUMTRANSAC = "' + NumTransac + '"');
        Break;
      end;
      Q.Next;
    end;

  finally
    Ferme(Q);
  end;
end;

{A la différence de SupprimeEcriture, on supprime ici un virement ou une transaction
 financière : toutes les écritures de la pièce/transaction sont supprimées
{---------------------------------------------------------------------------------------}
function SupprimePiece(Dos, NumTransac, NumPiece, ClefOpe : string) :  Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  s : string;
begin
  Result := True;
  try
    {On est dans le cas d'un virement}
    if ClefOpe <> '' then begin
      {On récupère le numéro de transaction.
       RAPPEL : pour un virement, on génère 2 écritures qui ont le même numéro de transaction, mais un
                numéro de pièce différent. Les commissions ont le même numéro de transaction et de pièce
                que l'écriture à laquelle elles se rapportent, mais un numéro de ligne différent
                25/09/06 : les comptes courants ont la même clef mais pas le même numéro de transaction}
      Q := OpenSQL('SELECT TE_NUMTRANSAC FROM TRECRITURE WHERE TE_CLEOPERATION = "' + ClefOpe + '"', True);
      try
        if Q.EOF then
          Result := False
        else
          {25/09/06 : Avec la gestion des comptes courants, on peut avoir plusieurs écritures avce la même clef}
          while not Q.EOF do begin
            s := 'DELETE FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + Q.FindField('TE_NUMTRANSAC').AsString + '"';
            if NumPiece <> '' then
              {On se limite à une pièce de la transaction}
              s := s + ' AND TE_NUMEROPIECE = ' + NumPiece;
            if Dos <> '' then
              s := s + ' AND TE_NODOSSIER = "' + Dos + '"';
            ExecuteSQL(s);
            Q.Next;
          end;
      finally
        Ferme(Q);
      end;
    end
    else if NumTransac <> '' then begin
      {Suppression de la transaction ?}
      s := 'DELETE FROM TRECRITURE WHERE TE_NUMTRANSAC = "' + NumTransac + '"';
      if NumPiece <> '' then
        {Non, on se limite à une pièce de la transaction}
        s := s + ' AND TE_NUMEROPIECE = ' + NumPiece;
      if Dos <> '' then
        s := s + ' AND TE_NODOSSIER = "' + Dos + '"';
      ExecuteSQL(s);
    end;
  except
    on E : Exception do Result := False;
  end;
end;

{---------------------------------------------------------------------------------------}
function UpdateEcritureStr(Dos, NumTransac, NumPiece, NumLigne, ClefOpe, Chp, Val : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := 'UPDATE TRECRITURE SET ' + Chp + ' = "' + Val + '" WHERE ';
  if ClefOpe <> '' then
    s := s + 'TE_CLEOPERATION = "' + ClefOpe + '" '
  else
    s := s + 'TE_NODOSSIER = "'+ Dos + '" AND TE_NUMTRANSAC = "' + NumTransac +
         '" AND TE_NUMEROPIECE = ' + NumPiece + 'AND TE_NUMLIGNE = ' + NumLigne;
  Result := ExecuteSQL(s) > 0;
end;

{---------------------------------------------------------------------------------------}
function UpdatePieceStr(Dos, NumTransac, NumPiece, ClefOpe, Chp, Val : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  s : string;
  T : string;
begin
  Result := NumTransac;
  T      := NumTransac;
  {On est dans le cas d'un virement}
  if ClefOpe <> '' then begin
    {On récupère le numéro de transaction.
     RAPPEL : pour un virement, on génère 2 écritures qui ont le même numéro de transaction, mais un
              numéro de pièce différent. Les commissions ont le même numéro de transaction et de pièce
              que l'écriture à laquelle elles se rapportent, mais un numéro de ligne différent}
    Q := OpenSQL('SELECT TE_NUMTRANSAC FROM TRECRITURE WHERE TE_CLEOPERATION = "' + ClefOpe + '" ' +
    {18/10/06 : exclusion des écritures courantes qui ne peuvent être modifiées que par le suivi des dates
                de valeur, et alors le traitement est fait dans TRCONTROLVALEUR_TOF}
                 'AND NOT (TE_NUMTRANSAC LIKE "' + CODEMODULECOU + '%")', True);
    if not Q.EOF then T := Q.FindField('TE_NUMTRANSAC').AsString;
    Ferme(Q);
  end;

  if T = '' then begin
    PGIError(TraduireMemoire('La mise à jour de la pièce à échouée (Cle opération : ' + ClefOpe + ').'#13 +
                             'Impossible de récupérer le numéro de transaction.'));
    Exit;
  end;

  Result := T;
  {Suppression de la transaction ?}
  s := 'UPDATE TRECRITURE SET ' + Chp + ' = "' + Val + '", TE_USERMODIF = "' + V_PGI.User +
       '", TE_DATEMODIF = "' + UsDateTime(Now) + '" WHERE TE_NUMTRANSAC = "' + T + '"';
  if NumPiece <> '' then
    {Non, on se limite à une pièce de la transaction}
    s := s + ' AND TE_NUMEROPIECE = ' + NumPiece;
  if Dos <> '' then
    s := s + ' AND TE_NODOSSIER = "' + Dos + '"';
  ExecuteSQL(s);
end;


{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}
{                                OBJET TObjComOPCVM                                     }
{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}

{FQ 10200 : L'application de la TVA aux commissions en plus des Frais implique, vu le mode
            de fonctionnement des propriétés GetCommissions et GetFrais que l'affectation
            des contrôles et des champs suivent un ordre bien précis :
            1/ Calcul des frais (c'est là que le cumul de la Tva est réinitialisé)
            2/ Calcul des Commissions
            3/ Récupération de la TVA qui maintenant est à jour
}

{---------------------------------------------------------------------------------------}
constructor TObjComOPCVM.Create(Cpte, Transac : string; dtOperation : TDateTime; VenteOk, ReelOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  Requete : string;
  QQ      : TQuery;
begin
  {On commence par récupérer le Code Flux}
  Requete := 'SELECT TTR_' + VERSEMENT + ', TTR_' + TOMBEE + ' FROM TRANSAC WHERE TTR_TRANSAC = "' + Transac + '"';
  QQ := OpenSQL(Requete, True);
  if not QQ.EOF then begin
    if VenteOk then
      CodeFlux := QQ.FindField('TTR_' + TOMBEE).AsString
    else
      CodeFlux := QQ.FindField('TTR_' + VERSEMENT).AsString;
  end;
  Ferme(QQ);{23/02/05 : Oubli de ma part}

  {Mémorisation des paramètres de création}
  General      := Cpte;
  DateOpe      := dtOperation;
  EnVente      := VenteOk;
  AvecEcriture := ReelOk;

  Requete := 'SELECT BANQUECP.BQ_AGENCE, BANQUECP.BQ_CODE, TRCONTRAT.TRC_CODECONTRAT, ' +
             'FLUXTRESO.TFT_CODETVA, FLUXTRESO.TFT_ASSUJETTITVA, ' +
             'TRCONTRAT.TRC_DEBCONTRAT, TRCONTRAT.TRC_FINCONTRAT, FRAIS.*, ' +
             'TRANCHEFRAIS.*, TRTRANCHEDETAIL.* FROM BANQUECP ' +
             'LEFT JOIN TRCONTRAT ON TRCONTRAT.TRC_AGENCE = BANQUECP.BQ_AGENCE ' +
             'LEFT JOIN FRAIS ON TFR_CONTRAT = TRC_CODECONTRAT ' +
             'LEFT JOIN FLUXTRESO ON TFT_FLUX = TFR_CODEFLUX ' +
             'LEFT JOIN TRANCHEFRAIS ON TTF_CODETRANCHE = TFR_TRANCHEFRAIS ' +
             'LEFT JOIN TRTRANCHEDETAIL ON TTD_TRANCHEFRAIS = TTF_CODETRANCHE ';
  Requete := Requete + 'WHERE BQ_CODE = "' + Cpte + '" AND TRC_DEBCONTRAT <= "' +
             UsDateTime(dtOperation) + '" AND TRC_FINCONTRAT >= "' + UsDateTime(dtOperation) +
             '" AND (TFR_TYPEFRAIS = "' + CODEFRAISOPCVM + '" OR TFR_TYPEFRAIS = "' + CODEFCOMOPCVM + '")';
  TobTransfo := TOB.Create('$$$', nil, -1);
  {Chargement de la requête}
  TobTransfo.LoadDetailFromSQL(Requete);
  {Purge des flux indésirables }
  PurgerTobCommision;
  {Si on n'est pas en mode écriture, création de la Tob qui contiendra les frais / commissions}
  TobFC  := TOB.Create('ùùù', nil, -1);
  TobTva := TOB.Create('$$$', nil, -1);
  OkTVA  := False;
end;

{---------------------------------------------------------------------------------------}
destructor TObjComOPCVM.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobFC ) then FreeAndNil(TobFC);
  if Assigned(TobTVA) then FreeAndNil(TobTVA);
  inherited Destroy;
end;

{Génération de l'écriture de trésorerie de la commission : TrEcrit est l'opération de
 trésorerie qui est soumise à commission}
{---------------------------------------------------------------------------------------}
procedure TObjComOPCVM.GenererCommissions(var tEcr : TOB);
{---------------------------------------------------------------------------------------}
var
  EcrCom : TOB;
  T      : TOB;
  n      : Integer;
  OldFra : string;
begin
  {Si l'on crée les écritures (c'est à dire que l'on est en validation BO), on exécute
   le traitement de l'ancêtre}
  if AvecEcriture then
    inherited GenererCommissions(tEcr)

  {Sinon on exécute seulement le calcul pour renseigner la table TROPCVM}
  else begin
    {Si une commission est par tranche, TobTransfo aura plusieurs ligne pour un même frais :
     on va donc tester que l'on change bien de frais avant de changer de ligne}
    OldFra := '';

    TobFC.ClearDetail;

    {Gestion des devises : on gère tout en devise pivot pour éviter les problèmes car on peut se retrouver
     avec trois devises : celle du frais, celle de l'OPCVM et celle du compte !!!}
    if tEcr.GetString('TE_DEVISE') <> V_PGI.DevisePivot then begin
      SetCotation(tEcr.GetString('TE_DEVISE'));
      tEcr.SetDouble('TE_COTATION', 1 / FCotation);
      if tEcr.GetDouble('TE_COTATION') = 0 then tEcr.SetDouble('TE_COTATION', 1);
      tEcr.SetDouble('TE_MONTANT', Arrondi(tEcr.GetDouble('TE_MONTANTDEV') * tEcr.GetDouble('TE_COTATION'), NbDecimal));
    end
    else begin
      tEcr.SetDouble('TE_COTATION', 1);
      tEcr.SetDouble('TE_MONTANT',  tEcr.GetDouble('TE_MONTANTDEV'));
    end;

    EcrCom := TOB.Create('TRECRITURE', nil, -1);
    try
      {On fait une boucle car une opération peut être soumise à plusieurs commission}
      for n := 0 to TobTransfo.Detail.Count - 1 do begin
        T := TobTransfo.Detail[n];
        {Si on est toujours sur le même frais, on continue}
        if OldFra = T.GetString('TFR_CODEFRAIS') then Continue;
        OldFra := T.GetString('TFR_CODEFRAIS');

        EcrCom.SetString('TE_CODECIB', T.GetString('TFR_TYPEFRAIS'));
        if T.GetString('TFT_ASSUJETTITVA') = 'X' then begin
          EcrCom.SetString('TE_CODEFLUX', T.GetString('TFT_CODETVA'));
          if not OkTVA then begin
            TobTva.LoadDetailFromSQL('SELECT TV_CODETAUX,TV_TAUXACH FROM TXCPTTVA WHERE TV_TVAOUTPF = "TX1" AND TV_REGIME IN ' +
                                     '(SELECT TRA_REGIMETVA FROM AGENCE WHERE TRA_AGENCE = "' + T.GetString('BQ_AGENCE') + '")');
            OkTva := True;
          end;
        end;

        EcrCom.SetDouble('TE_MONTANT', tEcr.GetDouble('TE_MONTANT'));
        {Calcul de la commission à partir du paramétrage}
        CalculerMntCommission(T, EcrCom);
        MajTobFC(EcrCom);
      end;
    finally
      if Assigned(EcrCom) then FreeAndNil(EcrCom);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjComOPCVM.MajTobFC(tEcr : TOB);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  T := TOB.Create('µµµ', TobFC, -1);
  T.AddChampSupValeur('NUM', TobFC.Detail.Count);
  T.AddChampSupValeur('ISF', tEcr.GetString('TE_CODECIB'));
  T.AddChampSupValeur('EUR', tEcr.GetDouble('TE_MONTANT'));
  T.AddChampSupValeur('TVA', tEcr.GetString('TE_CODEFLUX'));
end;

{---------------------------------------------------------------------------------------}
function TObjComOPCVM.GetCommission : Double;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  m : Double;
begin
  Result := 0;
  for n := 0 to TobFc.Detail.Count - 1 do begin
    {10/02/05 : FQ 10200 : application de la TVA au commission}
    if TobFc.Detail[n].GetString('ISF') = CODEFCOMOPCVM then begin
//      Result := Result + Abs(TobFc.Detail[n].GetDouble('EUR'));
      m := Abs(TobFC.Detail[n].GetDouble('EUR'));
      {Cumul de la Tva sur FMntTva et m est retourné amputé de la TVA}
      SetMntTva(m, TobFC.Detail[n].GetString('TVA'));
      Result := Result + m;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjComOPCVM.GetFrais : Double;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  m : Double;
begin
  {Initailisation des montants, la Tva n'etant calculée que sur les frais (pas sur les commissions)}
  Result  := 0;
  FMntTva := 0;

  for n := 0 to TobFc.Detail.Count - 1 do begin
    if TobFC.Detail[n].GetString('ISF') = CODEFRAISOPCVM then begin
      m := Abs(TobFC.Detail[n].GetDouble('EUR'));
      {Cumul de la Tva sur FMntTva et m est retourné amputé de la TVA}
      SetMntTva(m, TobFC.Detail[n].GetString('TVA'));
      Result := Result + m;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjComOPCVM.SetMntTva(var Value : Double; CodTva : string);
{---------------------------------------------------------------------------------------}
var
  TxTva : Double;
  T     : TOB;
  Mnt   : Double;
begin
  {Récupération du Taux de TVA}
  T := TobTVA.FindFirst(['TV_CODETAUX'], [CodTva], True);
  if Assigned(T) then TxTva := T.GetDouble('TV_TAUXACH') / 100
                 else Exit;

  {15/03/05 : FQ 10224 : La tva se calcul sur le HT et Non le TTC !!!}
  Mnt := GetMntTVAFromTTC(Abs(Value), TxTva);
  {On cumule la Tva}
  FMntTva := FMntTva + Mnt;
  {On ôte la Tva au montant des frais}
  Value := Value - Mnt;
end;

{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}
{                               OBJET TOBJCOMMISSIONTOB                                 }
{---------------------------------------------------------------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TObjCommissionTob.Create(Cpte, Flux : string; dtOperation : TDateTime);
{---------------------------------------------------------------------------------------}
var
  Requete : string;
begin
  {Mémorisation des paramètres de création}
  FGeneral  := Cpte;
  FCodeFlux := Flux;
  FDateOpe  := dtOperation;
  FDevise   := '';
  NbDecimal := V_PGI.OkDecV; {25/05/05 : Gestion du format des montants}

  Requete := 'SELECT BANQUECP.BQ_AGENCE, BANQUECP.BQ_CODE, TRCONTRAT.TRC_CODECONTRAT, ' +

             'TRCONTRAT.TRC_DEBCONTRAT, TRCONTRAT.TRC_FINCONTRAT, FRAIS.*, ' +
             'TRANCHEFRAIS.*, TRTRANCHEDETAIL.* FROM BANQUECP ' +
             'LEFT JOIN TRCONTRAT ON TRCONTRAT.TRC_AGENCE = BANQUECP.BQ_AGENCE ' +
             'LEFT JOIN FRAIS ON TFR_CONTRAT = TRC_CODECONTRAT ' +
             'LEFT JOIN TRANCHEFRAIS ON TTF_CODETRANCHE = TFR_TRANCHEFRAIS ' +
             'LEFT JOIN TRTRANCHEDETAIL ON TTD_TRANCHEFRAIS = TTF_CODETRANCHE ';

  Requete := Requete + 'WHERE BQ_CODE = "' + Cpte + '" AND TRC_DEBCONTRAT <= "' +
             UsDateTime(dtOperation) + '" AND TRC_FINCONTRAT >= "' + UsDateTime(dtOperation) +
             '" AND TFR_FLUXASSOCIE LIKE "%' + Flux + ';%"';
  TobTransfo := TOB.Create('$$$', nil, -1);
  {Chargement de la requête}
  TobTransfo.LoadDetailFromSQL(Requete);
  {Purge des flux indésirables (récupérés à cause du LIKE)}
  PurgerTobCommision;
  TobEcriture := TOB.Create('Ecr', nil, -1);
end;

{---------------------------------------------------------------------------------------}
destructor TObjCommissionTob.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobEcriture) then FreeAndNil(TobEcriture);
  if Assigned(TobTransfo)  then FreeAndNil(TobTransfo);
  inherited Destroy;
end;

{Read de la property ACommission
{---------------------------------------------------------------------------------------}
function TObjCommissionTob.IsSoumis : Boolean;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(TobTransfo) then Result := False
                              else Result := TobTransfo.Detail.Count > 0;
end;

{Vide la TobCommission des flux superflus : en effet les flux sont un VARCHAR(3) et non
 un CHAR(3) => si le Flux passé en paramètre du create de l'objet est CD, on a pu récupérer
 avec le LIKE des flux comme ACD, BCD ...
{---------------------------------------------------------------------------------------}
procedure TObjCommissionTob.PurgerTobCommision;
{---------------------------------------------------------------------------------------}
var
  n    : Integer;
  s, f : string;
  Ok   : Boolean;
begin
  for n := TobTransfo.Detail.Count - 1 downto 0 do begin
    s := TobTransfo.Detail[n].GetString('TFR_FLUXASSOCIE');
    f := ReadTokenSt(s);
    Ok := False;
    {On recherche si le flux appartient à la liste des flux associés ...}
    while (f <> '') and not Ok do begin
      Ok := FCodeFlux = f;
      f := ReadTokenSt(s);
    end;
    {... si ce n'est pas le cas, on supprime le frais}
    if not Ok then TobTransfo.Detail[n].Free;
  end;
end;

{Génération de l'écriture de trésorerie de la commission : TrEcrit est l'opération de
 trésorerie qui est soumise à commission}
{---------------------------------------------------------------------------------------}
procedure TObjCommissionTob.GenererCommissions(var tEcritures : TOB);
{---------------------------------------------------------------------------------------}
var
  tfEcr  : TOB;
  T      : TOB;
  n      : Integer;
  Q      : TQuery;
  NumLig : Integer;
  CodUnq : string;
  OldFra : string;
begin
  if not ACommission then Exit;

  {Le numéro de ligne des opérations financières est un. On par donc de deux pour les commissions}
  NumLig := 2;
  {Si une commission est par tranche, TobTransfo aura plusieurs ligne pour un même frais :
   on va donc tester que l'on change bien de frais avant de changer de ligne}
  OldFra := '';

  {On fait une boucle car une opération peut être soumise à plusieurs commission}
  for n := 0 to TobTransfo.Detail.Count - 1 do begin
    T := TobTransfo.Detail[n];

    {Si on est toujours sur le même frais, on continue}
    if OldFra = T.GetString('TFR_CODEFRAIS') then Continue;
    OldFra := T.GetString('TFR_CODEFRAIS');

    {JP 26/10/05 : Il est mieux de créer la Tob après le continue, cela evite d'insérer
                   en base des enregistrements vides et donc des duplications de clés
     Création de la tob fille qui va contenir les écritures}
    tfEcr := TOB.Create('TRECRITURE', TobEcriture, -1);

    {Reprise des valeurs communes à l'opération et à la commission}
    ReprendreChampsCommuns(tEcritures, tfEcr, T.GetString('TFR_LIBELLE'));
    {Récupération du Flux de génération}
    Q := OpenSQL('SELECT TFT_CODECIB,TFT_FLUX,TFT_GENERAL FROM FLUXTRESO WHERE TFT_FLUX = "' +
                 T.GetString('TFR_CODEFLUX') + '"', True);
    try
      tfEcr.SetString('TE_CODECIB'       , Q.FindField('TFT_CODECIB').AsString);
      tfEcr.SetString('TE_CONTREPARTIETR', Q.FindField('TFT_GENERAL').AsString);
      tfEcr.SetString('TE_CODEFLUX'      , Q.FindField('TFT_FLUX').AsString);
      tfEcr.SetInteger('TE_NUMLIGNE'     , NumLig);
      Inc(NumLig);
    finally
      Ferme(Q);
    end;

    {Calcul de la date de valeur}
    CalculerDateValeur(T, tfEcr);

    {Calcul des deux clefs}
    CodUnq := Commun.GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
    tfEcr.SetString('TE_CLEVALEUR'   , RetourneCleEcriture(tfEcr.GetDateTime('TE_DATEVALEUR'), StrToInt(CodUnq)));
    tfEcr.SetString('TE_CLEOPERATION', RetourneCleEcriture(tfEcr.GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodUnq)));
    Commun.SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodUnq);

    {Calcul de la commission à partir du paramétrage}
    CalculerMntCommission(T, tfEcr);
  end;
  {On modifie l'opération pour dire qu'il y a des commission}
  tEcritures.SetString('TE_COMMISSION', suc_AvecCom);
end;

{Reprise des valeurs communes à l'opération de trésorerie et à sa commission
{---------------------------------------------------------------------------------------}
procedure TObjCommissionTob.ReprendreChampsCommuns(var Source, Dest : TOB; Lib : string);
{---------------------------------------------------------------------------------------}
begin
  Dest.Dupliquer(Source, False, True);
  Dest.SetInteger('TE_CPNUMLIGNE', 1);
  Dest.SetInteger('TE_NUMECHE'   , 1);
  Dest.SetString('TE_LIBELLE'   , Lib + ' (' + Source.GetString('TE_LIBELLE') + ')');
  Dest.SetString('TE_COMMISSION', suc_Commission);
end;

{Calcul de la date de valeur de la commission : inspiré de l'unité échéance
{---------------------------------------------------------------------------------------}
procedure TObjCommissionTob.CalculerDateValeur(T : TOB; var E : TOB);
{---------------------------------------------------------------------------------------}
var
  APartir    : string;
  ArrondirAu : string;
  PlusJour   : Integer;
  D          : TDateTime;
  dd, mm, yy : Word ;
begin
  APartir    := T.GetString ('TFR_APARTIRDE');
  ArrondirAu := T.GetString ('TFR_ARRONDIJOUR');
  PlusJour   := T.GetInteger('TFR_PLUSJOUR');

  {S'il n'y de date de départ, la date de valeur est la même que la date d'opération}
  if (Trim(APartir) = '') then begin
    E.SetDateTime('TE_DATEVALEUR', E.GetDateTime('TE_DATECOMPTABLE'));
    Exit;
  end;

  D := E.GetDateTime('TE_DATECOMPTABLE');

  if APartir = 'FIN' then
    D := FinDeMois(E.GetDateTime('TE_DATECOMPTABLE'))
  else if APartir = 'DEB' then
    D := DebutDeMois(E.GetDateTime('TE_DATECOMPTABLE'));
  {else if APartir = 'ECR' then
    D := E.GetDateTime('TE_DATECOMPTABLE');}

  D := D + PlusJour;

  DecodeDate(D, yy, mm, dd) ;
  if ArrondirAu = 'FIN' then
    D := FinDeMois(D)
  else if ArrondirAu = 'DEB' then begin
    if dd > 1 then D := DebutDeMois(PlusMois(D, 1));
  end else if ArrondirAu = '05M' then begin
    if dd <= 05 then D := EncodeDate(yy, mm, 05)
                else D := DebutDeMois(PlusMois(D, 1)) + 4;
  end else if ArrondirAu = '10M' then begin
    if dd <= 10 then D := EncodeDate(yy, mm, 10)
                else D := DebutDeMois(PlusMois(D, 1)) + 9;
  end else if ArrondirAu = '15M' then begin
    if dd <= 15 then D := EncodeDate(yy, mm, 15)
                else D := DebutDeMois(PlusMois(D, 1)) + 14;
  end else if ArrondirAu = '20M' then begin
    if dd <= 20 then D := EncodeDate(yy, mm, 20)
                else D := DebutDeMois(PlusMois(D, 1)) + 19;
  end else if ArrondirAu = '25M' then begin
    if dd <= 25 then D := EncodeDate(yy, mm, 25)
                else D := DebutDeMois(PlusMois(D, 1)) + 24;
  end else if ArrondirAu='KED' then begin{Décade}
    if dd <= 10 then D := EncodeDate(yy, mm, 10)
                else if dd <= 20 then D := EncodeDate(yy, mm, 20)
                                 else D := FinDeMois(D);
  end else if ArrondirAu='KIN' then begin {Quinzaine}
    if dd <= 15 then D := EncodeDate(yy, mm, 15)
                else D := FinDeMois(D);
  end;
  E.SetDateTime('TE_DATEVALEUR', D);
end;

{Calcul du montant de la commission
{---------------------------------------------------------------------------------------}
procedure TObjCommissionTob.CalculerMntCommission(T : TOB; var E : TOB);
{---------------------------------------------------------------------------------------}
var
  Mnt : Double;
  Com : Double;
  Cot : Double;
begin
  Com := 0;
  Mnt := E.GetDouble('TE_MONTANT');
  {Si le mode de frais est la commission}
  if T.GetString('TFR_MODEFRAIS') = 'COM' then
    Com := CalcParCom(Mnt, T.GetDouble('TFR_COMMISSION'), T.GetDouble('TFR_COUTMIN'),
                           T.GetDouble('TFR_COUTMAX'), T.GetString('TFR_DEVISE'))
  {Sinon, si le mode de frais est par tranche}
  else if T.GetString('TFR_MODEFRAIS') = 'TRA' then
    Com := CalcParTra(Mnt, T.GetString('TFR_TRANCHEFRAIS'))
  {Sinon, si le mode de frais est le forfait}
  else if T.GetString('TFR_MODEFRAIS') = 'FOR' then
    Com := CalcParFor(T.GetDouble('TFR_COUTFORFAIT'), T.GetDouble('TFR_COUTUNITAIRE'), T.GetString('TFR_DEVISE'));
  {Calcul du montant en Devise}
  Cot := E.GetDouble('TE_COTATION');
  if Cot <> 0 then Mnt := Com * Cot
              else Mnt := Com;
  E.SetDouble('TE_MONTANT', Arrondi(Com, V_PGI.OkDecV));
  if V_PGI.DevisePivot <> T.GetString('TFR_DEVISE') then
    E.SetDouble('TE_MONTANTDEV', Arrondi(Mnt, NbDecimal))
  else
    E.SetDouble('TE_MONTANTDEV', Arrondi(Mnt, V_PGI.OkDecV));
end;

{Cacul d'une commissions par pourcentage}
{---------------------------------------------------------------------------------------}
function TObjCommissionTob.CalcParCom(Base, Com, Mini, Maxi : Double; DevF : string) : Double;
{---------------------------------------------------------------------------------------}
var
  Minimum  : Double;
  Maximum  : Double;
begin
  {Récupération du taux de change entre la devise du frais et celle de l'opération}
  SetCotation(DeVF);

  {Convertion des bornes dans la devise de l'opération}
  Minimum  := Mini * FCotation;
  Maximum  := Maxi * FCotation;

  Result := Base * Com /100;
  if (Abs(Result) > Maximum) and (Maximum <> 0) then Result := Maximum
  else if (Abs(Result) < Minimum) and (Minimum <>0) then Result := Minimum;

  {Pour être sûr que la commission sera négative}
  Result := - Abs(Result);
end;

{Calcul d'une commission par tranche.
 Remarque sur le signe des montant :
 1/ les tranches sont considérées comme toujours positives
 2/ le résultat (c-a-d, la commission) est toujours négatives
 3/ la base peut être des deux signes, c'est pourquoi on travaille en valeur absolue
 Remarque sur la devise des montants:
 1/ Le résultat doit être dans la devise de l'écriture, c'est à dire de la base
 2/ Par contre, le montant et la fourchette des tranches sont dans la devise des frais,
    ce qui nécessite une conversion des montants au préfixe "TTD_"
{---------------------------------------------------------------------------------------}
function TObjCommissionTob.CalcParTra(Base : Double; Tranche : string) : Double;
{---------------------------------------------------------------------------------------}
var
  Minimum  : Double;
  Maximum  : Double;
  T        : Tob;
  DevF     : string;
begin
  Result  := 0;

  {Récupération de la première tranche}
  TobTransfo.Detail.Sort('TRF_CONTRAT;TFR_CODEFRAIS;TTD_TRANCHEFRAIS;TTD_NUMTRANCHE');
  T := TobTransfo.FindFirst(['TFR_TRANCHEFRAIS'], [Tranche], True);
  if not Assigned(T) then Exit;

  {Récupération du taux de change entre la devise du frais et celle de l'opération}
  DevF := T.GetString('TFR_DEVISE');
  SetCotation(DevF);

  {On boucle sur les tranches pour connaître le montant ou le pourcentage qui s'applique}
  while T <> nil do begin
    Minimum := T.GetDouble('TTD_BORNEINF') * FCotation;
    Maximum := T.GetDouble('TTD_BORNESUP') * FCotation;

    {PAR NIVEAU}
    if T.GetString('TTF_TYPECALCFRAIS') = tcf_Niveau then begin
      {Si le montant se situe dans la tranche courante ...
       21/10/04 : Si Maximum = 0 et Minimum > 0, c'est a priori que l'on est sur la dernier Niveau}
      if (Abs(Base) > Minimum) and ((Abs(Base) <= Maximum) or ((Maximum = 0) and (Minimum > 0.001))) then begin
        {... on récupère le montant ou bien on applique le pourcentage à la base}
        if T.GetDouble('TTD_MONTANT') <> 0 then Result := - Abs(T.GetDouble('TTD_MONTANT') * FCotation)
                                           else Result := - Abs(T.GetDouble('TTD_POURCENTAGE') * Base / 100);
        Exit;
      end;
    end

    {PAR TRANCHE
     Les tranches peuvent se présenter de deux façons différentes
     a/ 0    -> 1000   ou   1/ 0    -> 1000
     b/ 1000 -> 2000        2/ 1000 -> 2000
     c/ 2000 -> 0           3/ 2000 -> 1000000}
    else begin
      {Si on est dans le cas 3/ ou bien c/}
      if ((Abs(Base) <= Maximum) or ((Maximum = 0) and (Minimum > 0.001))) then begin
        if T.GetDouble('TTD_MONTANT') <> 0 then
          {S'il s'agit d'un montant forfaitaire, on le cumule au résultat déjà calculé pour les autres tranches}
          Result := Result - Abs(T.GetDouble('TTD_MONTANT') * FCotation)
        else
          {S'il s'agit d'un pourcentage, on retranche à la base le minimum de la tranche courante(converti dans
           la devise de la base); on aplique le taux au montant obtenu et on cumule au résultat}
          Result := Result - Abs(T.GetDouble('TTD_POURCENTAGE') * (Base - (Minimum * FCotation)) / 100);

        {Pour être sûr que la commission sera négative}
        Result := - Abs(Result);
        Exit;
      end
      else if Abs(Base) > Maximum then begin
        if T.GetDouble('TTD_MONTANT') <> 0 then
          {S'il s'agit d'un montant forfaitaire, on le cumule au résultat déjà calculé pour les autres tranches}
          Result := Result - Abs(T.GetDouble('TTD_MONTANT') * FCotation)
        else
          {S'il s'agit d'un pourcentage, on aplique le taux sur le  montant de la tranche obtenu et on
           cumule au résultat des tranches précédentes}
          Result := Result - Abs(T.GetDouble('TTD_POURCENTAGE') * (Maximum - Minimum) * FCotation / 100);

        {Pour être sûr que la commission sera négative}
        Result := - Abs(Result);
      end;
    end;

    T := TobTransfo.FindNext(['TFR_TRANCHEFRAIS'], [Tranche], True);
  end;
end;

{Calcul d'une commission par forfait
 16/07/04 : POUR LE MOMENT LA NOTION DE FORFAIT N'EST PAS GÉRÉE !!!!! On se contente du
            coût unitaire. la gestion du forfait serait plutôt à envisager lors du
            traitement générant plusieurs écritures : dans ce cas, on générerait une
            écriture de commission lors de la génération du fichier des virements ....
{---------------------------------------------------------------------------------------}
function TObjCommissionTob.CalcParFor(Forfait, Unite : Double; DevF : string) : Double;
{---------------------------------------------------------------------------------------}
var
  Unitaire  : Double;
begin
  {Récupération du taux de change entre la devise du frais et celle de l'opération}
  SetCotation(DevF);

  {Récupération du coût unitaire}
  Unitaire  := Unite * FCotation;
  {Pour le moment (!?), le résultat est le coût unitaire}
  Result := - Abs(Unitaire);
end;


{---------------------------------------------------------------------------------------}
procedure TObjCommissionTob.SetCotation(Devise : string);
{---------------------------------------------------------------------------------------}
var
  ObjD : TObjDetailDevise;
begin
  if Devise <> FDevise then begin
    FDevise   := Devise;
    ObjD := TObjDetailDevise.Create;
    try
      {$IFNDEF EAGLSERVER}
      TheData := ObjD;
      {$ENDIF EAGLSERVER}
      FCotation := RetPariteEuro(Devise, FDateOpe);
      if FCotation = 0 then FCotation := 1;
      NbDecimal := ObjD.NbDecimal;
    finally
      {$IFNDEF EAGLSERVER}
      TheData := nil;
      {$ENDIF EAGLSERVER}
      If Assigned(ObjD) then FreeAndNil(ObjD);
    end;
  end;
end;

end.

