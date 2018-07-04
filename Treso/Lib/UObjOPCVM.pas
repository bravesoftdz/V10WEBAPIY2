unit UObjOPCVM;
{-------------------------------------------------------------------------------------
    Version  |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.20.001.001  03/01/04  JP    Création de l'unité
 6.20.001.003  10/02/05  JP    FQ 10200 : application de la TVA aux commissions en plus des Frais
 6.20.001.004  11/02/05  JP    Gestion des états
 7.09.001.001  28/11/06  JP    Nouveau mode de calcul du numéro de vente
 7.09.001.002  21/12/06  JP    Diverses corrections sur les ventes
 7.09.001.001  19/12/06  JP    FQ 10389 : Recalcul des ventes si des nouveaux cours ont été saisis
 7.09.001.004  11/01/07  JP    FQ 10397 : La plus value nette se calcul par déduction des frais
                               et commissions HT (=> on ne déduit pas la TVA de la plus value Brut)
 7.06.001.002  11/01/07  JP    FQ 10398 : On cherchait le contrat des frais à partir de la date d'achat
 8.01.001.001  15/12/06  JP    FQ 10373 : les nombres de part deviennent des doubles et non plus des Integers
--------------------------------------------------------------------------------------}

interface

uses
  SysUtils, HCtrls , Hmsgbox, Hent1, Constantes,
  UTOB, Classes, UObjGen, UProcCommission;

type
  TRecVente = record
    FraisVenEur : Double;
    FraisVenDev : Double;
    FraisAchEur : Double;
    FraisAchDev : Double;
  end;

  TObjOPCVMCalcul = class(TObject)
    ObjCom : TObjComOPCVM;

  private
    FNbPart      : Double; {15/12/06 : FQ 10373 : cela devient un Double. Nb d'OPCVM}
    FPrix        : Double;  {Prix unitaire de l'OPCVM}
    FMontant     : Double;  {NbPart * FPrix}
    FFrais       : Double;  {Frais HT sur FMontant}
    FTvaFrais    : Double;  {TVA sur FFrais}
    FCommission  : Double;  {Commissions sur FMontant}
    FTransaction : string;  {Code de la transaction}
    FMntTotal    : Double;  {Montant total de l'opération saisie par l'utilisateur}

    function  GetMontant    : Double;
    function  GetFrais      : Double;
    function  GetTvaFrais   : Double;
    function  GetCommission : Double;
    procedure SetNbPart     (Value : Double); {15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    procedure SetMontant    (Value : Double);
    procedure SetFrais      (Value : Double);
    procedure SetTvaFrais   (Value : Double);
    procedure SetCommission (Value : Double);
  protected
    FMntIntermediaire : Double; {Montant total de l'opération intermédiaire à comparer avec FMntTotal}
    FNbPartInter      : Double;{15/12/06 : FQ 10373 : On passe d'un Integer à un Double. Nb d'OPCVM intermédiaire à comparer avec FNbTotal}
    {15/12/06 : FQ 10373 : Puisque l'on passe au millième, on va affiner le calcul le calcul va être lancé
                quatre fois avec une pércision toujours plus grande : 1 ; 0,1 ; 0,001 ; 0,0001}
    procedure CalculRecursif(Unite : Double);
  public
    constructor Create;
    destructor  Destroy; override;
    procedure   Reinitialiser (Cpte, Transac : string; dtOperation : TDateTime);
    procedure   CalculerNbPart(var tEcr : TOB);

    property NbPart     : Double read FNbPart       write SetNbPart; {15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    property Montant    : Double read GetMontant    write SetMontant;
    property Frais      : Double read GetFrais      write SetFrais;
    property TvaFrais   : Double read GetTvaFrais   write SetTvaFrais;
    property Commission : Double read GetCommission write SetCommission;
  end;

  TObjOPCVMVente = class(TObject)
    tOPCVM  : TOB;
    tVente  : TOB;
    RecVen  : TRecVente;
  protected
    FOldDate   : TDateTime;
    FDateVente : TDateTime;
    FNbAVendre : Double; {15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    FCoursDev  : Double;
    FCoursOVM  : Double;
    FNumVente  : Integer;
    FNumLigne  : Integer;

    {Gestion des nombres de parts et de la date de fin}
    procedure CalculNbParts(var aOPCVM : TOB);
    {Affecte les champs de base de la tob de vente}
    procedure RemplitTobVente(var aOPCVM, aVente : TOB);
    {Effectue les calculs et génére les Tob de ventes}
    procedure LigneVente     (var aOPCVM, aVente : TOB);
    {Calcul des plus-values pour la Tob des ventes}
    procedure CalculPlusValue(var aOPCVM, aVente : TOB; mCamp : Double = 0);
    {Calcul des Rendements pour la Tob des ventes}
    procedure CalculRendement(var aOPCVM, aVente : TOB; mCamp : Double = 0);
    {Calcul des frais d'achat}
    procedure CalculFraisAchat(aOPCVM : TOB);
    {Récupération du numéro de vente dans la table COMPTEURTRESO}
    function  GetNumVente    : Integer;
    {Initialisation des diférentes variables de l'objet}
    procedure Initialiser    (VideTob : Boolean; aDate : TDateTime; Cpte, Opcvm : string);
    {Réinitialise le RecVente}
    procedure InitRecVente   ;
  public
    {En suppression d'un enregistrement de TRVENTEOPCVM, on se moque d'un certain nombre
     de traitements, en particulier la mise à jour de tVente : l'objectif est seulement de
     recalculer ("à l'envers") la vente pour pouvoir remettre à jour les champs concernant
     la vente de la table TROPCVM}
    EnSuppression : Boolean;
    EnAttente     : Boolean;
    TypeVente     : string;

    constructor Create(TypVente : string);
    destructor  Destroy; override;
    {Effectue les calculs et génére la Tob des ventes
     15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
    function    GenererVente(O : TOB; VideTob : Boolean; aDate : TDateTime; var NbParts : Double) : Boolean;
    {Ecrit les ventes dans la table TRVENTEOPCVM}
    procedure   PutVenteInBase;
    procedure   PmvEtRendementCamp;
  end;

  {20/12/06 : FQ 10389 : Je recrée un objet cr je ne suis pas sûr de mes précédents calculs}
  TObjCalculPMVR = class
  private
    TobVente  : TOB;
    TobAchat  : TOB;
    PrixMoyen : Double;
    TypeCal   : string;
    FNumVente : Integer;
    DateVente : TDateTime;
    CompteBq  : string;
    CodeOpcvm : string;
    {A True si on utilise CreateAvecTob, False avec CreateAvecVente}
    AvecTobs  : Boolean;

    procedure InitVariables;
    procedure CalculPxMoyen;
    {FQ 10389 : si on fait un recalcul car les cours on été modifiés}
    procedure MajCalculOPCVM;
    {19/12/06 : FQ 10389 : Mise à jour d'une ligne de vente}
    procedure MajLigneVente(var aVente : TOB);
    procedure CalculPMV (var V : TOB; A : TOB);
    procedure CalculRend(var V : TOB; A : TOB);
  public
    {on est en cours de création d'une vente}
    constructor CreateAvecTob(var tVente : TOB; tAchat : TOB; Gene, Opcvm : string);
    {FQ 10389 : on est en train de recalculer la vente}
    constructor CreateAvecVente(NumVente : Integer);
    destructor  Destroy; override;
    {si NewVente, on crée une vente, sinon FQ 10389 et on commence par MajCalculOPCVM}
    function    CalculVente(NewVente : Boolean) : Boolean;
    {Pour un changement de Parent}
    procedure   RecupVentes(var tVente : TOB);
  end;

  procedure SupprimerOPCVM(NumOPCVM : string);
  procedure SupprimerVente(NumVente : string);
  function  CreerEcritureVenteOPCVM(var lSolde : TStringList; FNumVente : string; ObjDev : TObjDevise) : Boolean;
  procedure SupprEcritureVenteOPCVM(FNumVente : string);


implementation

uses
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF EAGLCLIENT}
  UProcGen, Commun, UProcEcriture, UProcSolde, ParamSoc, Math, uTobDebug;


{---------------------------------------------------------------------------------------}
{---------------------            Méthodes publiques          --------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
procedure SupprimerOPCVM(NumOPCVM : string);
{---------------------------------------------------------------------------------------}
begin
  {Suppression de L'opcvm}
  ExecuteSQL('DELETE FROM TROPCVM WHERE TOP_NUMOPCVM = "' + NumOPCVM + '"');
  {Suppression des éventuelles écritures d'achat}
  SupprimePiece('', NumOPCVM, '', '');
end;

{---------------------------------------------------------------------------------------}
procedure SupprimerVente(NumVente : string);
{---------------------------------------------------------------------------------------}
var
  aTob : Tob;
  ObjV : TObjOPCVMVente;

    {Mise à jours dans la table TROPCVM des montants concernant la vente pour
     les enregistrements dont on va supprimer la vente
    {-----------------------------------------------------------------------}
    procedure MajOPCVM;
    {-----------------------------------------------------------------------}
    var
      n : Integer;
      T : TOB;
      F : TOB;
      P : Double;
    begin
      for n := 0 to aTob.Detail.Count - 1 do begin
        T := aTob.Detail[n];
        F := TOB.Create('TROPCVM', ObjV.tOPCVM, -1);
        F.SetString('TOP_NUMOPCVM', T.GetString('TVE_NUMTRANSAC'));
        F.LoadDB;
        {La mise du nombre de parts en une valeur négative permet d'avoir des montants négatifs
         et donc de retrancher les montants
         15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
        p := T.GetDouble('TVE_NBVENDUE') * -1;
        F.AddChampSupValeur('NBPARTS', p);

        ObjV.GenererVente(F, True, T.GetDateTime('TVE_DATEVENTE'), p);
        {MAj du statut}
        F.SetString('TOP_STATUT', '-');
        {Mise à jour de la table TROPCVM}
        F.UpdateDB;
      end;
    end;

begin
  ObjV  := TObjOPCVMVente.Create(vop_Supprime);
  aTob  := TOB.Create('!!!!', nil, - 1);
  try
    {REMARQUE : je ne fais volontairement pas de test sur  l'existance d'un enregistrement dans
                Q, car il doit toujours y avoir une vente : si tel n'était pas le cas, je préfère
                qu'il y ait une erreur pour interrompre le traitement}
    BeginTrans;
    try
      {Chargement des lignes de la vente en cours}
      ChargeDetailVente(aTob, NumVente);
      {Mise à jours des OPCVM concernées}
      MajOPCVM;
      aTob.ClearDetail;
      {Suppression de la vente ...}
      ExecuteSQL('DELETE FROM TRVENTEOPCVM WHERE TVE_NUMVENTE = ' + NumVente);
      {Suppression des éventuelles écritures}
      SupprimePiece('', GetNumTransacVente(NumVente, True), '', '');
      CommitTrans;
    except
      on E : Exception do begin
        RollBack;
        PGIError(E.Message);
        {On sort pour ne pas exécuter FinirTraitements;}
        Exit;
      end;
    end;
  finally
    if Assigned(ObjV) then FreeAndNil(ObjV);
    if Assigned(aTob) then FreeAndNil(aTob);
  end;
end;

{---------------------------------------------------------------------------------------}
function CreerEcritureVenteOPCVM(var lSolde : TStringList; FNumVente : string; ObjDev : TObjDevise) : Boolean;
{---------------------------------------------------------------------------------------}
var
  O : TOB;
  D : TOB;
  T : TOB;
  n : Integer;
begin
  Result := False;
  O := TOB.Create('$$$', nil, -1);
  T := TOB.Create('TRECRITURE', nil, -1);
  try
    ChargeDetailVente(O, FNumVente, True);
    if O.Detail.Count = 0 then Exit;
    try
      for n := 0 to O.Detail.Count - 1 do begin
        D := O.Detail[n];
        {Initialisation de l'écriture}
        InitNlleEcritureTob(T, D.GetString('TVE_GENERAL'), V_PGI.CodeSociete);
        {Reprise des champs de la transaction}
        T.SetDateTime('TE_DATECOMPTABLE', D.GetDateTime('TVE_DATEVENTE'));
        T.SetString('TE_LIBELLE', 'Vente n° ' + D.GetString('TVE_NUMVENTE') + ' du ' + D.GetString('TVE_NUMVENTE'));
        if GetParamSocSecur('SO_TRFIFO', True) then T.SetString('TE_REFINTERNE', T.GetString('TE_LIBELLE') + ' (FIFO)')
                                               else T.SetString('TE_REFINTERNE', T.GetString('TE_LIBELLE') + ' (CAMP)');

        T.SetString('TE_NUMTRANSAC', GetNumTransacVente(D.GetString('TVE_NUMVENTE'), True));

        {Récupération des informations depuis la table FluxTreso : Le flux est celui paramétré comme
         étant celui de versement dans la table Transac}
        ChargeChpFluxTreso(T, D.GetString('TVE_NUMTRANSAC'), TOMBEE, True);

        {Les montants sont négatifs puisqu'il s'agit d'un achat}
        T.SetDouble('TE_MONTANT', Arrondi(Abs(D.GetDouble('TVE_COURSEUR') * D.GetDouble('TVE_NBVENDUE')), V_PGI.OkDecV));

        {Termine l'écriture, gère les commissions et l'insère dans la rable}
        Result := TermineEcritureTob(T, ObjDev, True);
        if not Result then
          Exit;

        ExecuteSQL('UPDATE TRVENTEOPCVM SET TVE_VALBO = "X", TVE_USERVBO = "' + V_PGI.User + '", TVE_DATEVBO = "' +
                   UsDateTime(V_PGI.DateEntree) + '", TVE_USERMODIF = "' + V_PGI.User + '", TVE_DATEMODIF = "' +
                   UsDateTime(Now) + '" WHERE TVE_NUMVENTE = ' + FNumVente);
        AddGestionSoldes(lSolde, D.GetString('TVE_GENERAL'), D.GetDateTime('TVE_DATEVENTE'));
        {"Fermeture" des OPCVM si toutes les parts ont été vendues. }
        ExecuteSQL('UPDATE TROPCVM SET TOP_STATUT = "X", TOP_USERMODIF = "' + V_PGI.User + '", TOP_DATEMODIF = "' +
                   UsDateTime(Now) + '" WHERE TOP_GENERAL = "' + D.GetString('TVE_GENERAL') + '" AND ' + 'TOP_CODEOPCVM = "' +
                   D.GetString('TVE_CODEOPCVM') + '" AND TOP_STATUT <> "X" AND TOP_NBPARTVENDU = TOP_NBPARTACHETE');
      end; {For n := }
    except
      Result := False;
      raise;
    end;
  finally
    if Assigned(O) then FreeAndNil(O);
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure SupprEcritureVenteOPCVM(FNumVente : string);
{---------------------------------------------------------------------------------------}
var
  O : TOB;
  n : Integer;
begin
  O := TOB.Create('$$$', nil, -1);
  try
    ChargeDetailVente(O, FNumVente);
    if O.Detail.Count = 0 then Exit;

    {Suppression des écritures liées à la vente : opcvm, frais, commissions}
    if SupprimePiece('', GetNumTransacVente(FNumVente, True), '', '') then
      {Mise à jour de la table TROPCVM}
      ExecuteSQL('UPDATE TRVENTEOPCVM SET TVE_VALBO = "-", TVE_USERVBO = "", TVE_DATEVBO = "' + UsDateTime(iDate1900) +
                 '", TVE_USERMODIF = "' + V_PGI.User + '", TVE_DATEMODIF = "' + UsDateTime(Now) +
                 '" WHERE TVE_NUMVENTE = ' + FNumVente);
      {On Boucle sur les OPCVM liés à cette vente}
      for n := 0 to O.Detail.Count - 1 do
        ExecuteSQL('UPDATE TROPCVM SET TOP_STATUT = "-", TOP_USERMODIF = "' + V_PGI.User + '", TOP_DATEMODIF = "' +
                   UsDateTime(Now) + '" WHERE TOP_NUMOPCVM = "' + O.Detail[n].GetString('TVE_NUMTRANSAC') + '"');
  finally
    if Assigned(O) then FreeAndNil(O);
  end;
end;

{---------------------------------------------------------------------------------------}
{---------------------             TObjOPCVMCalcul            --------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TObjOPCVMCalcul.Create;
{---------------------------------------------------------------------------------------}
begin

end;

{---------------------------------------------------------------------------------------}
destructor TObjOPCVMCalcul.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjCom) then FreeAndNil(ObjCom);
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.Reinitialiser(Cpte, Transac : string; dtOperation : TDateTime);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjCom) then begin
    {Si l'objet est assigné mais que les des paramêtres de création est différent on le
    détruit avant de le recréer}
    if (ObjCom.General <> Cpte) or (ObjCom.DateOpe <> dtOperation) or (FTransaction <> Transac) then
      FreeAndNil(ObjCom)
    else
      Exit;
  end;
  {Création de l'objet qui va calculer les commissions, les frais et leur TVA}
  ObjCom := TObjComOPCVM.Create(Cpte, Transac, dtOperation, False, False);
  {On mémorise le code transaction pour une réinitialisation  ultérieure}
  FTransaction := Transac;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.CalculerNbPart(var tEcr : TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  U : Double;
begin
  Reinitialiser(tEcr.GetString('TE_GENERAL'), tEcr.GetString('TE_CODEFLUX'), tEcr.GetDateTime('TE_DATECOMPTABLE'));
  {Calcul du prix unitaire de l'OPCVM en devise Pivot}
  FPrix := tEcr.GetDouble('TE_SOLDEDEV') / tEcr.GetDouble('TE_COTATION');
  {Montant total de la transaction saisie par l'utilisateur}
  FMntTotal := tEcr.GetDouble('TE_MONTANTDEV');
  {Initialisation du Montant total intermédiaire}
  FMntIntermediaire := 0;
  {On lance le calcul récursif sur 90% du montant => on part par défaut du principe
   que les taxes, frais et commissions se montent à 10% du montant total}
  {Calcul du Nb de parts}
  NbPart := Round(FMntTotal * 0.9 / FPrix);

  {15/12/06 : FQ 10373 : on affine la précision du calcul à l'envers au millième}
  U := 10;
  for n := 1 to 4 do begin
    U := U / 10;
    {Initialisation du Montant total intermédiaire}
    FMntIntermediaire := 0;
    {Lancement du calcul récursif par précision}
    CalculRecursif(U);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.CalculRecursif(Unite : Double);
{---------------------------------------------------------------------------------------}
var
  tEcr : TOB;
  Mnt  : Double;

    {----------------------------------------------------------}
    procedure CalculComm;
    {----------------------------------------------------------}
    begin
      Montant := NbPart * FPrix;
      {Les montants sont déjà en devise pivot, donc on met la devise en devise pivot}
      tEcr.SetDouble('TE_MONTANTDEV', Montant);
      tEcr.SetString('TE_DEVISE', V_PGI.DevisePivot);
      {Calcul des frais, TVA et commissions}
      ObjCom.GenererCommissions(tEcr);
    end;
begin
  tEcr := TOB.Create('TRECRITURE', nil, -1);
  try
    {Calcul des commissions/frais/TVA}
    CalculComm;
    {Calcul du montant total de la transaction recalculé}
    Mnt := Montant + Frais + TvaFrais + Commission;
    if (Mnt = FMntTotal + Unite) or (Mnt = FMntTotal - Unite) then Exit;
    {Le précédent montant est situé de l'autre "côté" du montant saisi par l'utilisateur : il
     n'est donc pas possible d'être plus proche du montant saisi ...}
    if Between(FMntTotal, Mnt, FMntIntermediaire) or
       (Between(FMntIntermediaire, 0.1, FMntTotal) and (Mnt > FMntTotal)) then begin
       {... si le montant intermédiaire est plus proche du montant saisi que le Mnt, on recupère
        le nombre le nombre de part précédant et on recalcul les taxes, frais et commissions.
        Dans le cas contraire, on considère que les dernières valeurs calculées sont les bonnes.}
       if Abs(FMntIntermediaire - FMntTotal) < Abs(Mnt - FMntTotal) then begin
         NbPart := FNbPartInter;
        {Calcul des commissions/frais/TVA}
        CalculComm;
       end;
       Exit;
    end
    else if FMntIntermediaire < Mnt then begin
      {On mémorise le nombre de parts pour la dernière itération, si jamais il faut
       reprendre les valeurs de l'avant dernier passage}
      FNbPartInter := NbPart;
      {On mémorise le montant total pour comparaison à la prochaine itération}
      FMntIntermediaire := Mnt;
      if Mnt > FMntTotal then
        {I s'agit du cas où les commissions représentent plus de 10% du montant total}
        NbPart := NbPart - Unite
      else
        {Le montant précédent est plus petit, on augmente le nombre de part}
        NbPart := NbPart + Unite;
      CalculRecursif(Unite);
    end
    else if FMntIntermediaire > Mnt then begin
      {On mémorise le nombre de parts pour la dernière itération, si jamais il faut
       reprendre les valeurs de l'avant dernier passage}
      FNbPartInter := NbPart;
      {On mémorise le montant total pour comparaison à la prochaine itération}
      FMntIntermediaire := Mnt;
      {Le montant précédent est plus grand, on diminue le nombre de part}
      NbPart := NbPart - Unite;
      CalculRecursif(Unite);
    end;
  finally
    FreeAndNil(tEcr);
  end;
end;

{---------------------------------------------------------------------------------------}
function  TObjOPCVMCalcul.GetMontant : Double;
{---------------------------------------------------------------------------------------}
begin
  Result := FMontant;
end;

{---------------------------------------------------------------------------------------}
function  TObjOPCVMCalcul.GetFrais : Double;
{---------------------------------------------------------------------------------------}
begin
  Result := 0;
  if Assigned(ObjCom) then
    Result := ObjCom.MntFrais;
end;

{---------------------------------------------------------------------------------------}
function  TObjOPCVMCalcul.GetTvaFrais : Double;
{---------------------------------------------------------------------------------------}
begin
  Result := 0;
  if Assigned(ObjCom) then
    Result := ObjCom.MntTVA;
end;

{---------------------------------------------------------------------------------------}
function  TObjOPCVMCalcul.GetCommission : Double;
{---------------------------------------------------------------------------------------}
begin
  Result := 0;
  if Assigned(ObjCom) then
    Result := ObjCom.MntCommission;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.SetNbPart(Value : Double);
{---------------------------------------------------------------------------------------}
begin
  if Value <> FNbPart then begin
    FNbPart := Value;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.SetMontant(Value : Double);
{---------------------------------------------------------------------------------------}
begin
  if Value <> FMontant then begin
    FMontant := Value;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.SetFrais(Value : Double);
{---------------------------------------------------------------------------------------}
begin
  if Value <> FFrais then begin
    FFrais := Value;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.SetTvaFrais(Value : Double);
{---------------------------------------------------------------------------------------}
begin
  if Value <> FTvaFrais then begin
    FTvaFrais := Value;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMCalcul.SetCommission(Value : Double);
{---------------------------------------------------------------------------------------}
begin
  if Value <> FCommission then begin
    FCommission := Value;
  end;
end;


{---------------------------------------------------------------------------------------}
{---------------------             TObjOPCVMVente             --------------------------}
{---------------------------------------------------------------------------------------}


{---------------------------------------------------------------------------------------}
constructor TObjOPCVMVente.Create(TypVente : string);
{---------------------------------------------------------------------------------------}
begin
  TypeVente := TypVente;
  tOPCVM := TOB.Create('$TTROPOCVM', nil, -1);
  tVente := TOB.Create('$TRVENTEOPCVM', nil, -1);
  FOldDate   := iDate1900;
  FDateVente := iDate1900;
  FNumVente  := 0;
  EnSuppression := TypVente = vop_Supprime;
end;

{---------------------------------------------------------------------------------------}
destructor TObjOPCVMVente.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tOPCVM) then FreeAndNil(tOPCVM);
  if Assigned(tVente) then FreeAndNil(tVente);
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.Initialiser(VideTob : Boolean; aDate : TDateTime; Cpte, Opcvm : string);
{---------------------------------------------------------------------------------------}
var
  NewVente : Boolean;
  F        : TOB;
  n        : Integer;

    {-----------------------------------------------------------------}
    function _GetNumVente : Integer;
    {-----------------------------------------------------------------}
    var
      p : Integer;
    begin
      Result := 0;
      for P := tVente.Detail.Count - 1 downto 0 do 
        Result := Max(tVente.Detail[p].GetInteger('TVE_NUMVENTE'), Result);
      if Result = 0 then Result := GetNumVente
                    else Inc(Result);
    end;

begin
  {On mémorise la précédente date de vente pour limiter le nombre de requêtes}
  FOldDate   := FDateVente;
  FDateVente := aDate;
  NewVente   := False;

  if not EnSuppression then begin
    {Il faut un nouveau numéro de vente en vente partiel si la vente n'est pas déjà
     chargée et en vente multiple si on change d'OPCVM}
    NewVente := ((tVente.Detail.Count = 0) and (TypeVente = vop_Partiel)) or
                (TypeVente = vop_Simple) or
                (((TypeVente = vop_CAMP) or (TypeVente = vop_FIFO)) and (VideTob or (tVente.Detail.Count = 0)));
  end;

  {En vente multiple la tob est vidée dans la fiche TRVENTEOPCVM}
  if VideTob then begin
    for n := tVente.Detail.Count - 1 downto 0 do begin
      F := tVente.Detail[n];
      {Si on est sur le même général et la même OPCVM, on vide}
      if (F.GetString('TVE_GENERAL') = Cpte) and (F.GetString('TVE_CODEOPCVM') = Opcvm) then
        FreeAndNil(F);
    end;
  end;

  {Si l'objet n'est pas chargé, on récupère le numéro de vente dans COMPTEURS TRESO}
  if not EnSuppression and NewVente then begin
    FNumVente := _GetNumVente;
    FNumLigne := 1;
  end;
  EnAttente := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.InitRecVente;
{---------------------------------------------------------------------------------------}
begin
  with RecVen do begin
    FraisVenEur := 0;
    FraisVenDev := 0;
    FraisAchEur := 0;
    FraisAchDev := 0;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjOPCVMVente.GetNumVente : Integer;
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
begin
  Result := 1;
  {28/11/06 : Après la fusion Multi société, il peut y avoir eu un un changement de code société ...}
  {... On prend quelques précautions pour éviter une violation de clef : plutôt que d'utiliser
   les compteurs, on va directement dans la table}
  Q := OpenSQL('SELECT MAX(TVE_NUMVENTE) FROM TRVENTEOPCVM', True);
  if not Q.EOF then Result := Q.Fields[0].AsInteger + 1;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function TObjOPCVMVente.GenererVente(O : TOB; VideTob : Boolean; aDate : TDateTime; var NbParts : Double) : Boolean;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := False;

  Initialiser(VideTob, aDate, O.GetString('TOP_GENERAL'), O.GetString('TOP_CODEOPCVM'));
  {Gestion du nombre de parts}
  if (TypeVente = vop_FIFO) or (TypeVente = vop_CAMP) then
    FNbAVendre := NbParts
  else
    FNbAVendre := O.GetDouble('NBPARTS');

  CalculNbParts(O);


  if (TypeVente <> vop_Autre) and (TypeVente = vop_Supprime) then begin
    {FNbAVendre a été mis à jour dans CalculNbParts en fonction des parts restant à vendre
     sur la transaction chargée dans la tob O => pour la prochaine ligne de vente le nombre
     de parts restant à vendre est NbParts - FNbAVendre}
    NbParts := NbParts - FNbAVendre;
    {Mise à jour du champ calculé NBPARTS utile notamment en CAMP}
    O.SetDouble('NBPARTS', FNbAVendre);
  end;

  {Théoriquement les tests ont déjà été opérés, mais ...}
  if FNbAVendre = 0 then Exit;

  {Création de la ligne de vente}
  T := TOB.Create('TRVENTEOPCVM', tVente, -1);
  {Initialisation des champs de base}
  RemplitTobVente(O, T);
  {Initialise RecVen sauf en Camp ou il va s'agir d'un cumul}
  if TypeVente <> vop_CAMP then
    InitRecVente;
  {Calcul de la vente}
  LigneVente(O, T);

  {Il est inutile de faire les calculs sur les ventes lors d'une suppression de vente}
  if EnSuppression then begin
    EnAttente := True;
    Exit;
  end;
  (*
  {Initialise RecVen avec les frais de ventes et d'achat pour le calcul des
   plus-values et des rendements}
  CalculFraisAchat(O);
  if TypeVente <> vop_CAMP then begin
    {Calcul des plus-values}
    CalculPlusValue(O, T);
    {Calcul des Rendements}
    CalculRendement(O, T);
  end;
  *)

  {Pour les lignes des mêmes portefeuille, général et OPCVM, le numéro de vente est unique
   et c'est le numéro de ligne qui les distinguent}
  if (TypeVente = vop_FIFO) or (TypeVente = vop_CAMP) then
    {On incrémente le numéro ligne}
    Inc(FNumLigne);

  Result := True;
  EnAttente := Result;
//  TobDebug(tVente);
end;

{Gestion des nombres de parts et de la date de fin
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.CalculNbParts(var aOPCVM : TOB);
{---------------------------------------------------------------------------------------}
begin
  {Si toutes les parts achetées ont été vendues, on "clôture" la transaction d'achat
  15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
  if (aOPCVM.GetDouble('TOP_NBPARTVENDU') + FNbAVendre) >= aOPCVM.GetDouble('TOP_NBPARTACHETE') then begin
    {Si par extraordinaire le nombre de parts vendues et à vendre est supérieur à celui acheté, on limite
     le nombre de parts à vendre à la différence entre le nombre de parts achetées et vendues : sauf en
     cas de vente multiple le cas ne devrait pas se produire !!!}
    if (aOPCVM.GetDouble('TOP_NBPARTVENDU') + FNbAVendre) > aOPCVM.GetDouble('TOP_NBPARTACHETE') then
      FNbAVendre := aOPCVM.GetDouble('TOP_NBPARTACHETE') - aOPCVM.GetDouble('TOP_NBPARTVENDU');

    aOPCVM.SetDouble('TOP_NBPARTVENDU', aOPCVM.GetDouble('TOP_NBPARTACHETE'));
    {09/05/07 : FQ 10451 : Correction du contrôle des dates de vente}
    if FDateVente < aOPCVM.GetDateTime('TOP_DATEACHAT') then begin
      FNbAVendre := 0;
      Exit;
    end
    else
      aOPCVM.SetDateTime('TOP_DATEFIN', FDateVente);
  end
  {Sinon on cumule le nombre de parts vendues et à vendre}
  else
    aOPCVM.SetDouble('TOP_NBPARTVENDU', aOPCVM.GetDouble('TOP_NBPARTVENDU') + FNbAVendre);

  if EnSuppression then begin
    aOPCVM.SetString('TOP_STATUT', '-');
    aOPCVM.SetDateTime('TOP_DATEFIN', iDate1900);
    if aOPCVM.GetDouble('TOP_NBPARTVENDU') = 0 then
      aOPCVM.SetString('TOP_TYPEVENTE', '');
  end
  else
    {Gestion du type de vente}
    aOPCVM.SetString('TOP_TYPEVENTE', GetTypeVente(aOPCVM.GetString('TOP_TYPEVENTE'), TypeVente));
end;

{Affecte les champs de base de la tob de vente
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.RemplitTobVente(var aOPCVM, aVente : TOB);
{---------------------------------------------------------------------------------------}
begin
  if EnSuppression then Exit;

  aVente.SetInteger ('TVE_NUMVENTE'    , FNumVente);
  aVente.SetInteger ('TVE_NUMLIGNE'    , FNumLigne);
  aVente.SetDouble  ('TVE_NBVENDUE'    , FNbAVendre);
  aVente.SetString  ('TVE_LIBELLE'     , 'Vente n° ' + IntToStr(FNumVente) + ' - ' + IntToStr(FNumLigne) +
                                         ' du ' +  DateToStr(FDateVente));
  aVente.SetString  ('TVE_NUMTRANSAC'  , aOPCVM.GetString('TOP_NUMOPCVM'));
  aVente.SetString  ('TVE_PORTEFEUILLE', aOPCVM.GetString('TOP_PORTEFEUILLE'));
  aVente.SetString  ('TVE_GENERAL'     , aOPCVM.GetString('TOP_GENERAL'));
  aVente.SetString  ('TVE_CODEOPCVM'   , aOPCVM.GetString('TOP_CODEOPCVM'));
  aVente.SetString  ('TVE_DEVISE'      , aOPCVM.GetString('TOP_DEVISE'));
  aVente.SetString  ('TVE_SOCIETE'     , aOPCVM.GetString('TOP_SOCIETE'));
  aVente.SetString  ('TVE_NODOSSIER'   , aOPCVM.GetString('TOP_NODOSSIER'));
  aVente.SetString  ('TVE_USERCREATION', V_PGI.User);
  aVente.SetString  ('TVE_COMPTA'      , '-');
  aVente.SetString  ('TVE_VALBO'       , '-');
  aVente.SetString  ('TVE_TYPEVENTE'   , TypeVente);
  aVente.SetDateTime('TVE_DATECREATION', V_PGI.DateEntree);
  aVente.SetDateTime('TVE_DATEACHAT'   , aOPCVM.GetDateTime('TOP_DATEACHAT'));
  aVente.SetDateTime('TVE_DATEVENTE'   , FDateVente);
end;

{Effectue les calculs et génére les Tob de ventes
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.LigneVente(var aOPCVM, aVente : TOB);
{---------------------------------------------------------------------------------------}
var
  MntTemp  : Double;
  MntFrais : Double;
  MntTVA   : Double;
  MntEur   : Double;
  MntDev   : Double;
  Obj      : TObjComOPCVM;
  tEcr     : TOB;
begin
  MntDev := 0;
  MntEur := 0;
  try
    {Récupération des cours les plus récents de l'OPCVM et de la devise si la date a changé}
    if (FOldDate <> FDateVente) or (FCoursDev = 0) {or avec requêtes si pas partiel} then begin
      FCoursOVM := GetCoursOpcvm(aOPCVM.GetString('TOP_CODEOPCVM'), FDateVente);
      FCoursDev := RetPariteEuro(aOPCVM.GetString('TOP_DEVISE'), FDateVente, True);
    end;

    {Calcul du montant horstaxes et commissions}
    MntTemp := FNbAVendre * FCoursOVM;
    aOPCVM.SetDouble('TOP_MONTANTVENDEV', aOPCVM.GetDouble('TOP_MONTANTVENDEV') + Valeur(FormateMontant(MntTemp, 5)));
    MntDev := MntDev + MntTemp;

    {Calcul des frais d'achat au prorata de la vente}

    {Gestion de la devise}
    if FCoursDev = 0 then FCoursDev := 1;
    {On Stocke le cours en devise pivot}
    if not EnSuppression then
      aVente.SetDouble  ('TVE_COURSEUR', FCoursOVM / FCoursDev);

    MntTemp  := MntTemp / FCoursDev;
    MntEur   := MntEur + MntTemp;
    aVente.SetDouble('TVE_COTATION'   , Valeur(FormateMontant(FCoursDev, 5)));
    aOPCVM.SetDouble('TOP_MONTANTVEN' , aOPCVM.GetDouble('TOP_MONTANTVEN') + Valeur(FormateMontant(MntTemp , 5)));

    {La cotation de vente de la devise est une moyenne car la vente peut se faire en plusieurs fois}
    if aOPCVM.GetDouble('TOP_MONTANTVEN') = 0 then
      MntTva := 1
    else
      MntTva := aOPCVM.GetDouble('TOP_MONTANTVENDEV') / aOPCVM.GetDouble('TOP_MONTANTVEN');
    aOPCVM.SetDouble('TOP_COTATIONVEN', Valeur(FormateMontant(MntTVA , 5)));

    {Création de l'objet gérant le calcul des commissions et des frais
     11/01/07 : FQ 10398 : il faut partir de la date de vente pour récupérer le bon contrat}
    Obj := TObjComOPCVM.Create(aOPCVM.GetString('TOP_GENERAL'), aOPCVM.GetString('TOP_TRANSACTION'),
                               FDateVente, True, False); //aOPCVM.GetDateTime('TOP_DATEACHAT')
    tEcr := TOB.Create('TRECRITURE', nil, -1);
    try
      {Gestion des devises : on gère tout en devise pivot pour éviter les problèmes car on peut se retrouver
       avec trois devises : celle du frais, celle de l'OPCVM et celle du compte !!!}
      tEcr.SetDouble('TE_COTATION', FCoursDev);
      tEcr.SetDouble('TE_MONTANTDEV', Abs(MntTemp)); {En suppression, le montant peut être négatif}
      tEcr.SetString('TE_DEVISE', V_PGI.DevisePivot);

      {Mise en place des moyens nécessaires au calcul des frais et commissions}
      Obj.GenererCommissions(tEcr);
      {Calcul des frais et de leur TVA}
      MntFrais := Obj.MntFrais;
      {MntTVA  := Obj.MntTVA; FQ 10200 : la Tva est gérée en totalité avec les commissions}
      MntDev  := MntDev - (MntFrais * FCoursDev);
      MntEur  := MntEur - MntFrais;
      RecVen.FraisVenEur := MntFrais;
      RecVen.FraisVenDev := MntFrais * FCoursDev;

      {Pour être sûr que les montants sont bien négatifs}
      if EnSuppression then
        MntFrais := -1 * Abs(MntFrais);

      aOPCVM.SetDouble('TOP_FRAISVENDEV', aOPCVM.GetDouble('TOP_FRAISVENDEV') + Valeur(FormateMontant(MntFrais * FCoursDev, 5)));
      aOPCVM.SetDouble('TOP_FRAISVEN'   , aOPCVM.GetDouble('TOP_FRAISVEN'   ) + Valeur(FormateMontant(MntFrais, 5)));

      {Calcul des Commissions}
      MntTemp := Obj.MntCommission;
      {10/02/05 : FQ 10200 : application de la TVA au commission}
      MntTVA  := Obj.MntTVA;

      MntDev  := MntDev - (MntTemp * FCoursDev);
      MntEur  := MntEur - MntTemp;
      RecVen.FraisVenDev := RecVen.FraisVenDev + (MntTemp + MntTVA) * FCoursDev;
      RecVen.FraisVenEur := RecVen.FraisVenEur + MntTemp + MntTVA;

      {Pour être sûr que le montant est bien négatif}
      if EnSuppression then begin
        MntTemp := -1 * Abs(MntTemp);
        MntTVA  := -1 * Abs(MntTVA);
      end;

      aOPCVM.SetDouble('TOP_COMVENTE'   , aOPCVM.GetDouble('TOP_COMVENTE'   ) + Valeur(FormateMontant(MntTemp, 5)));
      aOPCVM.SetDouble('TOP_TVAVENTE'   , aOPCVM.GetDouble('TOP_TVAVENTE'   ) + Valeur(FormateMontant(MntTVA , 5)));
      aOPCVM.SetDouble('TOP_COMVENTEDEV', aOPCVM.GetDouble('TOP_COMVENTEDEV') + Valeur(FormateMontant(MntTemp * FCoursDev, 5)));
      aOPCVM.SetDouble('TOP_TVAVENTEDEV', aOPCVM.GetDouble('TOP_TVAVENTEDEV') + Valeur(FormateMontant(MntTVA  * FCoursDev, 5)));

      if not EnSuppression then begin
        {Dans l'historique, on ne stocke que le montant total}
        aVente.SetDouble('TVE_MONTANT'      , Valeur(FormateMontant(MntEur, 5)));
        aVente.SetDouble('TVE_MONTANTDEV'   , Valeur(FormateMontant(MntDev, 5)));
        {10/03/05 : Ajout des champs Frais commissions dans la base pour l'ordre de paiement
         21/12/06 : les bonnes variables pour les bons champs et les montants seront plus justes}
        aVente.SetDouble('TVE_TVAVENTEEUR'  , Valeur(FormateMontant(Abs(MntTVA  ), 5)));
        aVente.SetDouble('TVE_FRAISEUR'     , Valeur(FormateMontant(Abs(MntFrais), 5)));
        aVente.SetDouble('TVE_COMMISSIONEUR', Valeur(FormateMontant(Abs(MntTemp ), 5)));
      end;
    finally
      if Assigned(Obj) then FreeAndNil(Obj);
      if Assigned(tEcr) then FreeAndNil(tEcr);
    end;
  except
    raise;
  end;
end;

{Calcul des frais d'achat
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.CalculFraisAchat(aOPCVM : TOB);
{---------------------------------------------------------------------------------------}
var
  Prorata    : Double;
  fDev, fEur : Double;
begin
  {11/01/07 : FQ 10397 : La plus value nette se calcul par déduction des frais et commissions
              HT (=> on ne déduit pas la TVA de la plus value Brut)}
  Prorata := FNbAVendre / aOPCVM.GetDouble('TOP_NBPARTACHETE');
  fEur := aOPCVM.GetDouble('TOP_FRAISACH') + {aOPCVM.GetDouble('TOP_TVAACHAT') +}
          aOPCVM.GetDouble('TOP_COMACHAT');
  fDev := aOPCVM.GetDouble('TOP_FRAISACHDEV') + {aOPCVM.GetDouble('TOP_TVAACHATDEV') +}
          aOPCVM.GetDouble('TOP_COMACHATDEV');
  {En CAMP, le calcul des plus values et des rendement se fait en fin de traitement,
   donc on fait donc un cumul}
  if (TypeVente = vop_CAMP) then begin
    RecVen.FraisAchEur := RecVen.FraisAchEur + fEur * Prorata;
    RecVen.FraisAchDev := RecVen.FraisAchDev + fDev * Prorata;
  end else begin
    RecVen.FraisAchEur := fEur * Prorata;
    RecVen.FraisAchDev := fDev * Prorata;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.PmvEtRendementCamp;
{---------------------------------------------------------------------------------------}
var
  MntTmp : Double;
  Camp   : Double;
  aOpcvm : TOB;
  aVente : TOB;
  n      : Integer;
begin
  {1/ CALCUL DU COEFFICIENT CAMP}
  MntTmp := 0;
  Camp   := 0;

  for n := 0 to tOPCVM.Detail.Count - 1 do begin
    aOpcvm := tOPCVM.Detail[n];
    MntTmp := MntTmp + (aOpcvm.GetDouble('TOP_PRIXUNITAIRE') * aOpcvm.GetDouble('NBPARTS'));
    Camp   := Camp + aOpcvm.GetDouble('NBPARTS');
  end;

  {Camp est coût d'achat pondéré des OPCVM}
  if Camp = 0 then Camp := 1
              else Camp   := MntTmp / Camp;


  {2/ CALCUL DES PLUS OU MOINS VALUES}
  for n := 0 to tOPCVM.Detail.Count - 1 do begin
    aOpcvm := tOPCVM.Detail[n];
    {Recherche de la ligne de vente correspondant à la transaction courante}
    aVente := tVente.FindFirst(['TVE_NUMTRANSAC'], [aOpcvm.GetString('TOP_NUMOPCVM')], True);

    if aVente = nil then Continue;
    CalculPlusValue(aOpcvm, aVente, Camp);
  end;

  {3/ CALCUL DES RENDEMENTS}
  for n := 0 to tOPCVM.Detail.Count - 1 do begin
    aOpcvm := tOPCVM.Detail[n];
    {Recherche de la ligne de vente correspondant à la transaction courante}
    aVente := tVente.FindFirst(['TVE_NUMTRANSAC'], [aOpcvm.GetString('TOP_NUMOPCVM')], True);

    if aVente = nil then Continue;

    CalculRendement(aOpcvm, aVente, Camp);
  end;
end;

{Calcul des plus-values pour la Tob des ventes
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.CalculPlusValue(var aOPCVM, aVente : TOB; mCamp : Double = 0);
{---------------------------------------------------------------------------------------}
var
  MntTmp : Double;
begin
  if mCamp = 0 then
    mCamp := FCoursOVM - aOPCVM.GetDouble('TOP_PRIXUNITAIRE')
  else
    mCamp := FCoursOVM - mCamp;

  {Plus/Moins value brute : Nombre de part en vente * (Cours de vente - Cours d'achat)}
  MntTmp := FNbAVendre * (mCamp);
  aVente.SetDouble('TVE_PMVALUEBRUTDEV', MntTmp);
  aVente.SetDouble('TVE_PMVALUEBRUT'   , MntTmp / FCoursDev);
  {Plus/Moins value Nette : Plus / Moins values brute - Frais et commission d'achat
                                                      - frais et commissions de vente}
  MntTmp := MntTmp - RecVen.FraisAchDev - RecVen.FraisVenDev;
  aVente.SetDouble('TVE_PMVALUENETDEV', MntTmp);
  {21/12/06 : Avec TTE en fin de champ, cela va mieux}
  aVente.SetDouble('TVE_PMVALUENETTE' , MntTmp / FCoursDev);
end;

{Calcul des Rendements pour la Tob des ventes
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.CalculRendement(var aOPCVM, aVente : TOB; mCamp : Double = 0);
{---------------------------------------------------------------------------------------}
var
  MntTmp : Double;
  MntB1  : Double;
  MntB2  : Double;
  BaseJ  : Double;
begin
  if mCamp = 0 then
    mCamp := FNbAVendre * aOPCVM.GetDouble('TOP_PRIXUNITAIRE');

  {Base sans commissions et frais}
  MntB1 := (FDateVente - aOPCVM.GetDateTime('TOP_DATEACHAT')) * mCamp;
  {Base avec commissions et frais}
  MntB2 := (FDateVente - aOPCVM.GetDateTime('TOP_DATEACHAT')) *
           (mCamp + RecVen.FraisAchDev);
  {Sur la base, il y a un débat : quelle fonction utiliser ?
     -> CalcNbJourParAnBase() : qui renvoie le nombre de jours de l'année de la date passée en paramètre
                                en fonction de la base de cacul
     -> CalcNbJourReelBase() : qui renvoie le nombre de jour entre deux dates en fonction de la base
   Personnellement, la deuxième me parait plus pertinente !!!}
  BaseJ := CalcNbJourReelBase(aOPCVM.GetDateTime('TOP_DATEACHAT'), FDateVente, aOPCVM.GetInteger('TOP_BASECALCUL'));

  {Calcul du rendement brut :
      Plus/Moins value brute * Base de calcul * 100
      ---------------------------------------------
      Prix d'achat * nb Parts vendues * Nb de Jours}
  if MntB1 <> 0 then
    MntTmp := aVente.GetDouble('TVE_PMVALUEBRUTDEV') * 100 *  BaseJ/ MntB1
  else
    {Il n'est pas possible de caculer le remdement}
    MntTmp := 0;
  aVente.SetDouble('TVE_RENDEMENT', MntTmp);

  {Pour le rendement net, il y a discussion sur la base :
     - soit l'on prend la même base que pour le rendement brut (MntB1)
     - soit on tient compte dans la base des frais et commissions d'achat (MntB2)
   Pour le moment j'opte pour la seconde possibilité ...!!

   Calcul du rendement net :
            Plus/Moins value nette * Base de calcul * 100
            ---------------------------------------------
      Prix d'achat (com. incluses) * nb Parts vendues * Nb de Jours}

  if MntB2 <> 0 then
    MntTmp := aVente.GetDouble('TVE_PMVALUENETDEV') * 100 * BaseJ / MntB2
  else
    {Il n'est pas possible de caculer le remdement}
    MntTmp := 0;
  aVente.SetDouble('TVE_RENDEMENTNET', MntTmp);
end;

{Ecrit les ventes dans la table TRVENTEOPCVM
{---------------------------------------------------------------------------------------}
procedure TObjOPCVMVente.PutVenteInBase;
{---------------------------------------------------------------------------------------}
begin
  try
    tVente.InsertDb(nil);
  except
    raise;
  end;
end;


{---------------------------------------------------------------------------------------}
{---------------------------------   TObjCalculPMVR  -----------------------------------}
{---------------------------------------------------------------------------------------}

{On est en cours de création d'une vente}
{---------------------------------------------------------------------------------------}
constructor TObjCalculPMVR.CreateAvecTob(var tVente : TOB; tAchat: TOB; Gene, Opcvm : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  AvecTobs := True;
  TobAchat := tAchat;
  TobVente := TOB.Create('!!!!', nil, - 1);
  for n := tVente.Detail.Count - 1 downto 0 do begin
    F := tVente.Detail[n];
    if (F.GetString('TVE_GENERAL') = Gene) and (F.GetString('TVE_CODEOPCVM') = Opcvm) then
      F.ChangeParent(TobVente, 0);
  end;

  InitVariables;
end;

{FQ 10389 : on est en train de recalculer la vente}
{---------------------------------------------------------------------------------------}
constructor TObjCalculPMVR.CreateAvecVente(NumVente : Integer);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  AvecTobs := False;
  TobAchat := TOB.Create('!!!!', nil, - 1);
  TobVente := TOB.Create('!!!!', nil, - 1);
  {Chargement des lignes de la vente en cours}
  ChargeDetailVente(TobVente, IntToStr(NumVente));
  {Récupération de l'OPCVM}
  for n := 0 to TobVente.Detail.Count - 1 do begin
    F := TOB.Create('TROPCVM', TobAchat, -1);
    F.SetString('TOP_NUMOPCVM', TobVente.Detail[n].GetString('TVE_NUMTRANSAC'));
    F.LoadDB;
  end;
  InitVariables;
end;

{---------------------------------------------------------------------------------------}
destructor TObjCalculPMVR.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if not AvecTobs then 
    if Assigned(TobAchat) then FreeAndNil(TobAchat);
  if Assigned(TobVente) then FreeAndNil(TobVente);
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.InitVariables;
{---------------------------------------------------------------------------------------}
begin
  if GetParamSocSecur('SO_TRFIFO', True) then TypeCal := vop_FIFO
                                         else TypeCal := vop_CAMP;

  if TobVente.Detail.Count > 0 then begin           
    CompteBq  := TobVente.Detail[0].GetString('TVE_GENERAL');
    DateVente := TobVente.Detail[0].GetDateTime('TVE_DATEVENTE');
    CodeOpcvm := TobVente.Detail[0].GetString('TVE_CODEOPCVM');
    FNumVente := TobVente.Detail[0].GetInteger('TVE_NUMVENTE');
    {Calcul du prix unitaire moyen en devise pivot}
    CalculPxMoyen;
  end;
end;

{Retourne le prix unitaire moyen en devise pivot
{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.CalculPxMoyen;
{---------------------------------------------------------------------------------------}
begin
  PrixMoyen := 0;
  if TypeCal = vop_camp then
    PrixMoyen := GetCAMP(CompteBq, CodeOpcvm, DateVente);
end;

{---------------------------------------------------------------------------------------}
function TObjCalculPMVR.CalculVente(NewVente : Boolean) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  V : TOB;
  A : TOB;
begin
  Result := False;
  {FQ 10389 : si on recalcule la vente suite à une modification des cours, on commence par remettre à jour
              les montants d'achat, de vente et de frais et commissions}
  if not NewVente then MajCalculOPCVM;
  if not Assigned(TobAchat) or not Assigned(TobVente) or (TobAchat.Detail.Count = 0) or
     (TobVente.Detail.Count = 0) then Exit;

  Result := True;
  for n := 0 to TobVente.Detail.Count - 1 do begin
    V := TobVente.Detail[n];
    A := TobAchat.FindFirst(['TOP_NUMOPCVM'], [V.GetString('TVE_NUMTRANSAC')], True);
    if not Assigned(A) then begin
      Result := False;
      Continue;
    end;
    CalculPMV (V, A);
    CalculRend(V, A);
  end;

  if not AvecTobs then begin
    TobAchat.UpdateDB(True);
    TobVente.UpdateDB(True);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.CalculPMV(var V : TOB; A : TOB);
{---------------------------------------------------------------------------------------}
var
  MntTmp : Double;
  VenTmp : Double;
  AchTmp : Double;
begin
  {Plus/Moins value brute : Nombre de part en vente * (Cours de vente - Cours d'achat)}
  if TypeCal = vop_camp then
    MntTmp := (V.GetDouble('TVE_COURSEUR') - PrixMoyen) * V.GetDouble('TVE_NBVENDUE')
  else
    MntTmp := (V.GetDouble('TVE_COURSEUR') - A.GetDouble('TOP_PRIXUNITAIRE') / A.GetDouble('TOP_COTATIONACH')) * V.GetDouble('TVE_NBVENDUE');

  V.SetDouble('TVE_PMVALUEBRUT'   , MntTmp);
  V.SetDouble('TVE_PMVALUEBRUTDEV', MntTmp * V.GetDouble('TVE_COTATION'));

  {Plus/Moins value Nette : Plus / Moins values brute - Frais et commission d'achat
                                                      - frais et commissions de vente
   11/01/07 : FQ 10397 : La plus value nette se calcul par déduction des frais et commissions
              HT (=> on ne déduit pas la TVA de la plus value Brut)}
  VenTmp := V.GetDouble('TVE_FRAISEUR') + V.GetDouble('TVE_COMMISSIONEUR'){ + V.GetDouble('TVE_TVAVENTEEUR')};
  AchTmp := A.GetDouble('TOP_FRAISACH') + A.GetDouble('TOP_COMACHAT') { + A.GetDouble('TOP_TVAACHAT')} ;
  AchTmp := AchTmp * V.GetDouble('TVE_NBVENDUE') / A.GetDouble('TOP_NBPARTACHETE');
  MntTmp := MntTmp - VenTmp - AchTmp;

  V.SetDouble('TVE_PMVALUENETTE' , MntTmp);
  V.SetDouble('TVE_PMVALUENETDEV', MntTmp * V.GetDouble('TVE_COTATION'));
end;

{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.CalculRend(var V : TOB; A : TOB);
{---------------------------------------------------------------------------------------}
var
  MntTmp : Double;
  ComACh : Double;
  MntB1  : Double;
  MntB2  : Double;
  BaseJ  : Double;
  mCamp  : Double;
begin
  if TypeCal = vop_camp then
    mCamp := V.GetDouble('TVE_NBVENDUE') * PrixMoyen
  else
    mCamp := V.GetDouble('TVE_NBVENDUE') * A.GetDouble('TOP_PRIXUNITAIRE') / A.GetDouble('TOP_COTATIONACH');

  {Sur la base, il y a un débat : quelle fonction utiliser ?
     -> CalcNbJourParAnBase() : qui renvoie le nombre de jours de l'année de la date passée en paramètre
                                en fonction de la base de cacul
     -> CalcNbJourReelBase() : qui renvoie le nombre de jour entre deux dates en fonction de la base
   Personnellement, la deuxième me parait plus pertinente !!!}
  BaseJ := CalcNbJourReelBase(A.GetDateTime('TOP_DATEACHAT'), DateVente, A.GetInteger('TOP_BASECALCUL'));

  {Base sans commissions et frais}
  MntB1 := (DateVente - A.GetDateTime('TOP_DATEACHAT')) * mCamp;
  {Base avec commissions et frais
   11/01/07 : FQ 10397 : La plus value nette se calcul par déduction des frais et commissions
              HT (=> on ne déduit pas la TVA de la plus value Brut)}
  ComAch := A.GetDouble('TOP_FRAISACH') + {A.GetDouble('TOP_TVAACHAT') + }A.GetDouble('TOP_COMACHAT');
  ComAch := ComAch * V.GetDouble('TVE_NBVENDUE') / A.GetDouble('TOP_NBPARTACHETE');
  MntB2  := (DateVente - A.GetDateTime('TOP_DATEACHAT')) *
            (mCamp + ComAch);

  {Calcul du rendement brut :
      Plus/Moins value brute * Base de calcul * 100
      ---------------------------------------------
      Prix d'achat * nb Parts vendues * Nb de Jours}
  if MntB1 <> 0 then
    MntTmp := V.GetDouble('TVE_PMVALUEBRUT') * 100 *  BaseJ/ MntB1
  else
    {Il n'est pas possible de caculer le remdement}
    MntTmp := 0;
  V.SetDouble('TVE_RENDEMENT', MntTmp);

  {Pour le rendement net, il y a discussion sur la base :
     - soit l'on prend la même base que pour le rendement brut (MntB1)
     - soit on tient compte dans la base des frais et commissions d'achat (MntB2)
   Pour le moment j'opte pour la seconde possibilité ...!!

   Calcul du rendement net :
            Plus/Moins value nette * Base de calcul * 100
            ---------------------------------------------
      Prix d'achat (com. achats incluses) * nb Parts vendues * Nb de Jours}

  if MntB2 <> 0 then
    MntTmp := V.GetDouble('TVE_PMVALUENETTE') * 100 * BaseJ / MntB2
  else
    {Il n'est pas possible de caculer le remdement}
    MntTmp := 0;
  V.SetDouble('TVE_RENDEMENTNET', MntTmp);
end;

{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.MajCalculOPCVM;
{---------------------------------------------------------------------------------------}
var
  aOPCVM : TOB;

    {Mise à jours dans la table TROPCVM des montants concernant la vente pour
     les enregistrements dont on va supprimer la vente
    {-----------------------------------------------------------------------}
    procedure MajOPCVM(tVente : TOB; DeduitOk : Boolean);
    {-----------------------------------------------------------------------}
    begin
      if not DeduitOk then begin
        {On ote la précédente vente}
        aOPCVM.SetDouble('TOP_MONTANTVENDEV', aOPCVM.GetDouble('TOP_MONTANTVENDEV') - tVente.GetDouble('TVE_MONTANTDEV'));
        aOPCVM.SetDouble('TOP_MONTANTVEN'   , aOPCVM.GetDouble('TOP_MONTANTVEN'   ) - tVente.GetDouble('TVE_MONTANT'));
        aOPCVM.SetDouble('TOP_FRAISVEN'     , aOPCVM.GetDouble('TOP_FRAISVEN'     ) - tVente.GetDouble('TVE_FRAISEUR'));
        aOPCVM.SetDouble('TOP_COMVENTE'     , aOPCVM.GetDouble('TOP_COMVENTE'     ) - tVente.GetDouble('TVE_COMMISSIONEUR'));
        aOPCVM.SetDouble('TOP_TVAVENTE'     , aOPCVM.GetDouble('TOP_TVAVENTE'     ) - tVente.GetDouble('TVE_TVAVENTEEUR'));
        aOPCVM.SetDouble('TOP_COMVENTEDEV'  , aOPCVM.GetDouble('TOP_COMVENTEDEV'  ) - tVente.GetDouble('TVE_COMMISSIONEUR') * tVente.GetDouble('TVE_COTATION'));
        aOPCVM.SetDouble('TOP_TVAVENTEDEV'  , aOPCVM.GetDouble('TOP_TVAVENTEDEV'  ) - tVente.GetDouble('TVE_TVAVENTEEUR'  ) * tVente.GetDouble('TVE_COTATION'));
        aOPCVM.SetDouble('TOP_FRAISVENDEV'  , aOPCVM.GetDouble('TOP_FRAISVENDEV'  ) - tVente.GetDouble('TVE_FRAISEUR'     ) * tVente.GetDouble('TVE_COTATION'));
      end

      else begin
        {on ajoute la nouvelle vente}
        aOPCVM.SetDouble('TOP_MONTANTVENDEV', aOPCVM.GetDouble('TOP_MONTANTVENDEV') + tVente.GetDouble('TVE_MONTANTDEV'));
        aOPCVM.SetDouble('TOP_MONTANTVEN'   , aOPCVM.GetDouble('TOP_MONTANTVEN'   ) + tVente.GetDouble('TVE_MONTANT'));
        aOPCVM.SetDouble('TOP_FRAISVEN'     , aOPCVM.GetDouble('TOP_FRAISVEN'     ) + tVente.GetDouble('TVE_FRAISEUR'));
        aOPCVM.SetDouble('TOP_COMVENTE'     , aOPCVM.GetDouble('TOP_COMVENTE'     ) + tVente.GetDouble('TVE_COMMISSIONEUR'));
        aOPCVM.SetDouble('TOP_TVAVENTE'     , aOPCVM.GetDouble('TOP_TVAVENTE'     ) + tVente.GetDouble('TVE_TVAVENTEEUR'));
        aOPCVM.SetDouble('TOP_COMVENTEDEV'  , aOPCVM.GetDouble('TOP_COMVENTEDEV'  ) + tVente.GetDouble('TVE_COMMISSIONEUR') * tVente.GetDouble('TVE_COTATION'));
        aOPCVM.SetDouble('TOP_TVAVENTEDEV'  , aOPCVM.GetDouble('TOP_TVAVENTEDEV'  ) + tVente.GetDouble('TVE_TVAVENTEEUR'  ) * tVente.GetDouble('TVE_COTATION'));
        aOPCVM.SetDouble('TOP_FRAISVENDEV'  , aOPCVM.GetDouble('TOP_FRAISVENDEV'  ) + tVente.GetDouble('TVE_FRAISEUR'     ) * tVente.GetDouble('TVE_COTATION'));
        if aOPCVM.GetDouble('TOP_MONTANTVEN') <> 0 then
          aOPCVM.SetDouble('TOP_COTATIONVEN', aOPCVM.GetDouble('TOP_MONTANTVENDEV') / aOPCVM.GetDouble('TOP_MONTANTVEN'));
      end;
    end;

var
  n : Integer;
  F : TOB;
  T : TOB;
begin
  T := TOB.Create('mmmm', nil, -1);
  try
    for n := TobVente.Detail.Count - 1 downto 0 do begin
      F := TobVente.Detail[n];
      aOPCVM := TobAchat.FindFirst(['TOP_NUMOPCVM'], [F.GetString('TVE_NUMTRANSAC')], True);
      if not Assigned(aOPCVM) then begin
        F.ChangeParent(T, -1);
        Continue;
      end;
      {Mise à jours des OPCVM par suppression de la vente}
      MajOPCVM(F, False);
      {Nouvelle vente}
      MajLigneVente(F);
      {Mise à jours des OPCVM par ajout de la vente mise à jour et enregistrement dans la table}
      MajOPCVM(F, True);
    end;
  finally
    FreeAndNil(T);
  end;
end;

{19/12/06 : FQ 10389 : Mise à jour d'une ligne de vente}
{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.MajLigneVente(var aVente : TOB);
{---------------------------------------------------------------------------------------}
var
  MntTemp  : Double;
  MntFrais : Double;
  MntTVA   : Double;
  MntEur   : Double;
  MntDev   : Double;
  Obj      : TObjComOPCVM;
  tEcr     : TOB;
  FCoursOVM: Double;
  FCoursDev: Double;
begin
  MntDev := 0;
  MntEur := 0;
  try
    FCoursOVM := GetCoursOpcvm(aVente.GetString('TVE_CODEOPCVM'), DateVente);
    FCoursDev := RetPariteEuro(aVente.GetString('TVE_DEVISE'), DateVente, True);

    {Calcul du montant horstaxes et commissions}
    MntTemp := aVente.GetDouble('TVE_NBVENDUE') * FCoursOVM;
    MntDev := MntDev + MntTemp;

    {Calcul des frais d'achat au prorata de la vente}

    {Gestion de la devise}
    if FCoursDev = 0 then FCoursDev := 1;
    {On Stocke le cours en devise pivot}
    aVente.SetDouble  ('TVE_COURSEUR', FCoursOVM / FCoursDev);

    MntTemp  := MntTemp / FCoursDev;
    MntEur   := MntEur + MntTemp;
    aVente.SetDouble('TVE_COTATION', Valeur(FormateMontant(FCoursDev, 5)));

    {Création de l'objet gérant le calcul des commissions et des frais}
    Obj := TObjComOPCVM.Create(CompteBq, TobAchat.Detail[0].GetString('TOP_TRANSACTION'), DateVente, True, False);
    tEcr := TOB.Create('TRECRITURE', nil, -1);
    try
      {Gestion des devises : on gère tout en devise pivot pour éviter les problèmes car on peut se retrouver
       avec trois devises : celle du frais, celle de l'OPCVM et celle du compte !!!}
      tEcr.SetDouble('TE_COTATION', FCoursDev);
      tEcr.SetDouble('TE_MONTANTDEV', Abs(MntTemp)); {En suppression, le montant peut être négatif}
      tEcr.SetString('TE_DEVISE', V_PGI.DevisePivot);

      {Mise en place des moyens nécessaires au calcul des frais et commissions}
      Obj.GenererCommissions(tEcr);
      {Calcul des frais et de leur TVA}
      MntFrais := Obj.MntFrais;
      {MntTVA  := Obj.MntTVA; FQ 10200 : la Tva est gérée en totalité avec les commissions}
      MntDev  := MntDev - (MntFrais * FCoursDev);
      MntEur  := MntEur - MntFrais;

      {Calcul des Commissions}
      MntTemp := Obj.MntCommission;
      {10/02/05 : FQ 10200 : application de la TVA au commission}
      MntTVA  := Obj.MntTVA;

      MntDev  := MntDev - (MntTemp * FCoursDev);
      MntEur  := MntEur - MntTemp;

      {Dans l'historique, on ne stocke que le montant total}
      aVente.SetDouble('TVE_MONTANT'      , Valeur(FormateMontant(MntEur, 5)));
      aVente.SetDouble('TVE_MONTANTDEV'   , Valeur(FormateMontant(MntDev, 5)));
      {10/03/05 : Ajout des champs Frais commissions dans la base pour l'ordre de paiement}
      aVente.SetDouble('TVE_TVAVENTEEUR'  , Valeur(FormateMontant(Abs(MntTVA  ), 5)));
      aVente.SetDouble('TVE_FRAISEUR'     , Valeur(FormateMontant(Abs(MntFrais), 5)));
      aVente.SetDouble('TVE_COMMISSIONEUR', Valeur(FormateMontant(Abs(MntTemp ), 5)));
    finally
      if Assigned(Obj) then FreeAndNil(Obj);
      if Assigned(tEcr) then FreeAndNil(tEcr);
    end;
  except
    raise;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCalculPMVR.RecupVentes(var tVente: TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  for n := TobVente.Detail.Count - 1 downto 0 do begin
    F := TobVente.Detail[n];
    F.ChangeParent(tVente, 0);
  end;
end;

end.


