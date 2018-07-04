unit uLibCFONB;

interface

uses
  {$IFDEF EAGLCLIENT}
  UtileAGL, {LanceEtat}
  {$ELSE}
  EdtREtat, {LanceEtat}
  {$ENDIF}
  uTob, SysUtils, Classes, HCtrls, Vierge;

const
  {Valeurs de la tablettes CPTYPECFONB}
  cfonb_Transfert   = 'TRI';
  cfonb_Prelevement = 'PRE';
  cfonb_Virement    = 'VIR';
  cfonb_LCR         = 'LCR';
  cfonb_CISX        = 'CIS';

  const_Annuler     = 'ANNULER';

type
  TObjDevise = class
    Decimales : Integer;
    Quotite   : Integer;
  end;

  TObjCFONB = class(TOB)
  private
    {Liste contenant les informations des devises pour la fonction GetInfosDevise}
    ListeDev : TStringList;

    procedure MajLigne ; virtual;
    procedure MajEntete; virtual;
    procedure MajPied  ; virtual;
    function  LancerDialog : Boolean;
  protected
    {Code erreur lors du traitement}
    CodeErreur : Integer;
    {Pour les transfert internationaux, si mono-devise FQ 18603}
    CodeDevise : string;
    {Liste des factures en cas de cumul des tiers sur une ligne}
    ListeFac   : TStringList;

    function PrepareMontant(var Deb, Cred : Double; var T : TOB) : Integer; virtual;
    {Fonction de contrôle générique avant le lancement des traitements}
    function ControleOk : Boolean; virtual;
    {Mise à jour des champs E_CFONBOK, E_NUMCFONB, E_RIB}
    procedure FlagEcriture;
    {Gestion du refus}
    procedure GereRefus(MsgErr : string; T : TOB);
  public
    TypeExport     : string;
    {Tob contenant les données envoyées au générateur de fichiers}
    TobTraitement  : TOB;
    {Tob contenant les écritures rejetées car incomplètes}
    TobEcrRejetees : TOB;
    {Interuption du traitement à la suite d'une erreur}
    StopperTraitement : Boolean;
    {Si on ne veut pas ouvrir la boite de dialogue de LancerDialog,
     dans ce cas il faut renseigner le champs 'REPERTOIRE'}
    ModeSilent : Boolean;
    {Fait-on un encaissement ou un décaissement}
    IsEncaissement : Boolean;
    { Pays pour l'export CISX }
    Pays        : string;
    { Script CISX }
    Script      : string;
    { Banque Export CISX }
    Banque      : string;

    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    destructor  Destroy; override;
    {Récupération des informations generiques du fichier CFONB par appel de la fiche CPCRITERECFONB}
    procedure   GetInfosSaisies;
    {Complétion de la tob de traitement et génération du fichier}
    procedure   LanceTraitement; virtual;
    {Réinitialisation de l'objet}
    procedure   Nettoyer;
    {Lancement éventuel d'un bordereau}
    procedure   LanceBordereau;
    {Lancement du générateur de fichier}
    procedure   LancerFichier;
    {Si on demande un cumul par tier, concentre les tiers sur une tob fille}
    procedure   SommeTiers;
    {Retourne le message d'erreur en fonction de CodeErreur}
    function    GetMessageErreur : string; virtual;
    {Affichage du message d'erreur si l'on n'est pas en ModeSilent}
    procedure   AfficheMsgErreur;
    {Gestion en un seul champ des TIC/TID (table GENERAUX) et de collectifs (table TIERS)}
    function    GetInfosTiers(var T : TOB) : Boolean;
    {Gestion du RIB / IBAN sur les écriture / tiers / TICTID}
    function    GetInfosRib  (var T : TOB; IbanOk : Boolean; F : TOB) : Boolean;
    {Retourne les infos devises (Nb decimales, Quotité ...}
    function    GetInfosDevise(Dev : string)  : TObjDevise;
    {Récupération du ParamSoc SO_CPNUMCFONB}
    function    GetNumCFONB(Cat : string) : string;
  end;

  TObjTransfert = class(TObjCFONB)
  private
    procedure MajLigne ; override;
    procedure MajEntete; override;
    procedure MajTobFille(var T : TOB; F : TOB);
    function  GetImputation : string;
    function  GetRemise : Char;
  protected
    function PrepareMontant(var Deb, Cred : Double; var T : TOB) : Integer; override;
  public
    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    function    GetMessageErreur : string; override;
  end;

  TObjVirement = class(TObjCFONB)

  public
    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    function    GetMessageErreur : string; override;
  end;

  TObjPrelevement = class(TObjCFONB)

  public
    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    function    GetMessageErreur : string; override;
  end;

  TObjLCR = class(TObjCFONB)

  public
    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    function    GetMessageErreur : string; override;
  end;

{$IFDEF CISXPGI}
  TObjCISX = class(TObjCFONB)

  public
    constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
    function    GetMessageErreur : string; override;
    procedure   LanceTraitement ;override;
  private
    Variables : TOB;
    FSaisieVar    : TFVierge;
    procedure GetInfosSup;
    procedure BValideClickVar(Sender: TObject);
    function  SaisieVariable (TV : TOB): Boolean;
  end;

{$ENDIF CISXPGI}
  procedure ChargeTHValComboBox ( SQL : string ;var Combo : THValComboBox);

  { GetCurrentProcessEnvVar retourne la valeur de la variable d'environnement }
  { VariableName du processus courant ou une chaîne vide si cette variable    }
  { n'existe pas.                                                             }
  function GetCurrentProcessEnvVar(const VariableName: string): string;
implementation

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  mc_erreur, LP_Base, LP_View, LP_PARAM, CPCRITERECFONB_TOF, AglInit, Dialogs,
  HMsgBox, Forms, HEnt1, ULibWindows, ParamSoc, UtilPGI, UProcGen, Math,
  ULibEcriture,
{$IFDEF CISXPGI}
  // Pour l'export CISX :
{$IFNDEF EAGLCLIENT}
  uScript,
{$ELSE}
  uScriptTob,
{$ENDIF}
  echg_code,uexport,FileCtrl,Windows,
{$ENDIF}
  Controls, uTobDebug, CBPPath; //SDA le 14/12/2007 ajout CBPPath

                             { ------------------- }
                             {      TObjCFONB      }
                             { ------------------- }

{---------------------------------------------------------------------------------------}
constructor TObjCFONB.Create(LeNomTable: string; LeParent: TOB; IndiceFils: Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
  TobTraitement := TOB.Create('pppp', nil, -1);
  TobEcrRejetees := TOB.Create('rrrr', nil, -1);;

  AddChampSup('DOCUMENT', False);
  AddChampSup('COMPTEBQ', False);
  AddChampSup('MODELEFICHIER', False);
  AddChampSup('REPERTOIRE', False);
  CodeErreur := 0;
  ListeDev := TStringList.Create;
  ListeFac := TStringList.Create;
end;

{---------------------------------------------------------------------------------------}
destructor TObjCFONB.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobTraitement ) then FreeAndNil(TobTraitement);
  if Assigned(TobEcrRejetees) then FreeAndNil(TobEcrRejetees);
  if Assigned(ListeFac)       then FreeAndNil(ListeFac);
  if Assigned(ListeDev)       then LibereListe(ListeDev, True);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.LanceTraitement;
{---------------------------------------------------------------------------------------}
begin
  {Cumul éventuel des tiers sur une ligne}
  SommeTiers;
  {Exécute les contrôles de base}
  if not StopperTraitement then ControleOk;
  {Exécution d'un éventuel traitement spécifique en début de fichier}
  if not StopperTraitement then MajEntete;
  {Complétion de la tob de traitement avec les données nécessaires au fichier}
  if not StopperTraitement then MajLigne;
  {Exécution d'un éventuel traitement spécifique en fin de fichier}
  if not StopperTraitement then MajPied;
  {Génération du fichier proprement dite}
  if not StopperTraitement then LancerFichier;
  {Mise à jour de la table écriture}
  if not StopperTraitement then FlagEcriture;
  AfficheMsgErreur;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.MajEntete;
{---------------------------------------------------------------------------------------}
begin
{Virtuelle}
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.MajLigne;
{---------------------------------------------------------------------------------------}
begin
{Virtuelle}
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.MajPied;
{---------------------------------------------------------------------------------------}
begin
{Virtuelle}
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.Nettoyer;
{---------------------------------------------------------------------------------------}
begin
  ClearDetail;
  TobTraitement .ClearDetail;
  TobEcrRejetees.ClearDetail;
  ListeFac.Clear;
  StopperTraitement := False;
  CodeErreur := 0;
end;

{Récupération des informations à saisir par appel de la fiche CPCRITERECFONB
{---------------------------------------------------------------------------------------}
procedure TObjCFONB.GetInfosSaisies;
{---------------------------------------------------------------------------------------}
begin
  TheTob := TOB(self);
  StopperTraitement := False;
  StopperTraitement := LanceCritereCfonb(TypeExport) = const_Annuler;
end;

{Lancement éventuel du bordereau d'accompagnement
{---------------------------------------------------------------------------------------}
procedure TObjCFONB.LanceBordereau;
{---------------------------------------------------------------------------------------}
begin
  if not StopperTraitement and (GetString('DOCUMENT') <> '') then
    LanceEtatTob('E', 'BOR', GetString('DOCUMENT'), TobTraitement, True, False, False, nil, '', '', False);
end;

{Génération du fichier
{---------------------------------------------------------------------------------------}
procedure TObjCFONB.LancerFichier;
{---------------------------------------------------------------------------------------}
var
  Err : TMC_ERR;
begin
  if V_PGI.SAV and not ModeSilent then
    TobDebug(TobTraitement);
  if ModeSilent or LancerDialog  then begin
    {... Lancement de la génération du fichier}
    ExporteLP('F', TypeExport, GetString('MODELEFICHIER'), '', GetString('REPERTOIRE'), not ModeSilent, False, TobTraitement, Err);
    {Gestion des éventuelles erreurs}
    if Err.Code > 0 then begin
      CodeErreur := Err.Code * -1;
      StopperTraitement := True;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.GereRefus(MsgErr : string; T : TOB);
{---------------------------------------------------------------------------------------}
var
  s : string;
  n : Integer;
begin
  T.AddChampSupValeur('ERREUR', MsgErr);
  T.ChangeParent(TobEcrRejetees, -1);
  {Gestion du cas où les lignes de tiers sont cumulés : on retire toutes les lignes attachées}
  if (T.GetString('G_NATUREGENE') = 'TIC') or (T.GetString('G_NATUREGENE') = 'TID') then
    s := T.GetString('E_DEVISE') + ';' + T.GetString('G_NATUREGENE') + ';' + T.GetString('E_GENERAL') + ';'
  else
    s := T.GetString('E_DEVISE') + ';' + T.GetString('G_NATUREGENE') + ';' + T.GetString('E_AUXILIAIRE') + ';';

  for n := ListeFac.Count - 1 downto 0 do begin
    if s = Copy(ListeFac[n], 1, Length(s)) then
      ListeFac.Delete(n);
  end;
end;

{Boite de dialogue et génération du fichier ASCII
{---------------------------------------------------------------------------------------}
function TObjCFONB.LancerDialog : Boolean;
{---------------------------------------------------------------------------------------}
var
  SD  : TSaveDialog;
begin
  Result := True;

  if ModeSilent then begin
    if GetString('REPERTOIRE') = '' then begin
      StopperTraitement := True;
      CodeErreur := 1000;
      Result := False;
    end;
  end

  else begin
    {Création et paramétrage de la boite de dialogue}
    SD := TSaveDialog.Create(Application);
    try
      SD.DefaultExt := 'TXT';
      SD.Filter := 'Fichier Texte (*.txt)|*.txt';
      SD.FilterIndex := 1;
      SD.Title := 'Export';
      Result := SD.Execute;
      if not Result then begin
        StopperTraitement := True;
        CodeErreur := 0;
      end
      else
        SetString('REPERTOIRE', SD.FileName);
    finally
      if Assigned(SD) then FreeAndNil(SD);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.AfficheMsgErreur;
{---------------------------------------------------------------------------------------}
begin
  if not ModeSilent and (CodeErreur <> 0) then
    PGIError(GetMessageErreur);
end;

{---------------------------------------------------------------------------------------}
function TObjCFONB.GetInfosTiers(var T : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  k : Integer;
  p : Char;

      {--------------------------------------------------------------------}
      procedure MajLibelle(Chp : string);
      {--------------------------------------------------------------------}
      begin
        {Création du nouveau champ ...}
        T.AddChampSupValeur('C' + Chp, FullMajuscule(T.GetString(p + Chp)));
      end;

      {--------------------------------------------------------------------}
      procedure SuppChpSup(Chp : string);
      {--------------------------------------------------------------------}
      begin
        {... et suppression des anciens}
        T.DelChampSup('T' + Chp, False);
        T.DelChampSup('G' + Chp, False);
      end;

begin
  Result := True;

  if (T.GetNumChamp('G_NATUREGENE') < 0) then begin
    CodeErreur := 1001;
    Result := False;
    Exit;
  end;

  if (T.GetString('G_NATUREGENE') = 'TID') or (T.GetString('G_NATUREGENE') = 'TIC') then
    p := 'G'
  else
    p := 'T';

  k := T.GetChampCount(ttcSup) + 999; {ttcAll, ttcReal}
  for n := 1000 to k do begin
    if T.GetNomChamp(n) = '' then Break;
    if T.GetNomChamp(n)[1] <> p then Continue;
    {En attendant, dautres demandes : APE, division territoriale ...}
         if T.GetNomChamp(n) = (p + '_LIBELLE'         ) then MajLibelle('_LIBELLE'         )
    else if T.GetNomChamp(n) = (p + '_NATUREECONOMIQUE') then MajLibelle('_NATUREECONOMIQUE')
    else if T.GetNomChamp(n) = (p + '_PAYS'            ) then MajLibelle('_PAYS'            )
    else if T.GetNomChamp(n) = (p + '_ADRESSE1'        ) then MajLibelle('_ADRESSE1'        )
    else if T.GetNomChamp(n) = (p + '_ADRESSE2'        ) then MajLibelle('_ADRESSE2'        )
    else if T.GetNomChamp(n) = (p + '_ADRESSE3'        ) then MajLibelle('_ADRESSE3'        )
    else if T.GetNomChamp(n) = (p + '_VILLE'           ) then MajLibelle('_VILLE'           )
    else if T.GetNomChamp(n) = (p + '_CODEPOSTAL'      ) then MajLibelle('_CODEPOSTAL'      )
    else if T.GetNomChamp(n) = (p + '_FAX'             ) then MajLibelle('_FAX'             )
    else if T.GetNomChamp(n) = (p + '_TELEPHONE'       ) then MajLibelle('_TELEPHONE'       )
    else if T.GetNomChamp(n) = (p + '_SIRET'           ) then MajLibelle('_SIRET'           )
    else if T.GetNomChamp(n) = (p + '_TELEX'           ) then MajLibelle('_TELEX'           )
    else if T.GetNomChamp(n) = (p + '_NIF'             ) then MajLibelle('_NIF'             );
  end;

  {k := T.GetChampCount(ttcSup) + 999; {ttcAll, ttcReal}
  {On supprime les champs des tables Généraux et Tiers}
  SuppChpSup('_LIBELLE'         );
  SuppChpSup('_NATUREECONOMIQUE');
  SuppChpSup('_PAYS'            );
  SuppChpSup('_ADRESSE1'        );
  SuppChpSup('_ADRESSE2'        );
  SuppChpSup('_ADRESSE3'        );
  SuppChpSup('_VILLE'           );
  SuppChpSup('_CODEPOSTAL'      );
  SuppChpSup('_FAX'             );
  SuppChpSup('_TELEPHONE'       );
  SuppChpSup('_SIRET'           );
  SuppChpSup('_TELEX'           );
  SuppChpSup('_NIF'             );
end;

{Récupération des informations nécessaires à l'identification bancaire du tiers
{---------------------------------------------------------------------------------------}
function TObjCFONB.GetInfosRib(var T : TOB; IbanOk : Boolean; F : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  s   : string;
  SQL : string;
  Q   : TQuery;
begin
  Result := True;

  s := FullMajuscule(T.GetString('E_RIB'));

  if (s = '') or ((s[1] <> '*') and IbanOk) then begin
    {Rib / Iban incorrect et on ne demande pas de récupérer les RIB principal}
    if GetString('RIBPRINC') <> 'X' then begin
      GereRefus(TraduireMemoire('Impossible de récupérer le RIB / IBAN'), T);
      Result := False;
      Exit;
    end

    {On reprend les infos sur la précédente écriture du Tiers}
    else if Assigned(F) then begin
      s := FullMajuscule(F.GetString('E_RIB'));
      if s = '' then begin
        GereRefus(TraduireMemoire('Impossible de récupérer le RIB / IBAN'), T);
        Result := False;
        Exit;
      end;

      T.SetString('E_RIB', s);
      T.SetString('R_CODEIBAN', F.GetString('R_CODEIBAN'));
      T.SetString('NATECO', F.GetString('NATECO'));
      T.SetString('R_CODEBIC', F.GetString('R_CODEBIC'));
      T.DelChampSup('C_NATUREECONOMIQUE', False);
      T.AddChampSupValeur('PY_CODEISO2', F.GetString('PY_CODEISO2'));
      T.AddChampSupValeur('PY_LIBELLE' , F.GetString('PY_LIBELLE' ));

    end

    else begin
      {Si l'on travaille sur le code Iban}
      if IbanOk then begin
        SQL := '';
        if (T.GetString('G_NATUREGENE') = 'TID') or (T.GetString('G_NATUREGENE') = 'TIC') then
          SQL := 'SELECT R_CODEIBAN, R_NATECO, R_CODEBIC, R_PAYS, R_VILLE, PY_CODEISO2, PY_LIBELLE FROM RIB ' +
                 'LEFT JOIN PAYS ON PY_PAYS = R_PAYS WHERE R_AUXILIAIRE = "' + T.GetString('E_GENERAL') + '" AND R_PRINCIPAL = "X"'
        else
          SQL := 'SELECT R_CODEIBAN, R_NATECO, R_CODEBIC, R_PAYS, R_VILLE, PY_CODEISO2, PY_LIBELLE FROM RIB ' +
                 'LEFT JOIN PAYS ON PY_PAYS = R_PAYS WHERE R_AUXILIAIRE = "' + T.GetString('E_AUXILIAIRE') + '" AND R_PRINCIPAL = "X"';

        Q := OpenSelect(SQL, T.GetString('SYSDOSSIER'), True);
        {Impossible de récupérer le code Iban}
        if Q.FindField('R_CODEIBAN').AsString = '' then begin
          GereRefus(TraduireMemoire('Impossible de récupérer le RIB / IBAN'), T);
          Result := False;
          Ferme(Q);
          Exit;
        end
        else begin
          T.SetString('R_CODEIBAN', Q.FindField('R_CODEIBAN').AsString);
          T.SetString('E_RIB', '*' + Q.FindField('R_CODEIBAN').AsString);

          {Nature eco : 1/ La combo, 2/ celui du rib, 3/ celui du tiers}
          if Trim(GetString('NATECO')) <> '' then
            T.SetString('NATECO', GetString('NATECO'))
          else if Trim(Q.FindField('R_NATECO').AsString) <> '' then
            T.SetString('NATECO', Q.FindField('R_NATECO').AsString)
          else
            T.SetString('NATECO', T.GetString('C_NATUREECONOMIQUE'));

          {Impossible de récupérer la nature économique}
          if T.GetString('NATECO') = '' then begin
            GereRefus(TraduireMemoire('Impossible de récupérer la nature économique'), T);
            Result := False;
            Ferme(Q);
            Exit;
          end;
        end;
        T.DelChampSup('C_NATUREECONOMIQUE', False);
        T.SetString('R_CODEBIC', Q.FindField('R_CODEBIC').AsString);
        T.AddChampSupValeur('PY_CODEISO2', Q.FindField('PY_CODEISO2').AsString);
        T.AddChampSupValeur('PY_LIBELLE' , Q.FindField('PY_LIBELLE' ).AsString);
        Ferme(Q);

      end

      {Si l'on travaille sur le RIB}
      else begin
        {A FAIRE SI BESOIN}
        {A FAIRE SI BESOIN}
        {A FAIRE SI BESOIN}
        {A FAIRE SI BESOIN}
        {A FAIRE SI BESOIN}
      end;
    end;

  end;

  {Gestion du code pays}
  if T.GetString('C_PAYS') = '' then begin
    GereRefus(TraduireMemoire('Le tiers n''a pas de pays renseigné'), T);
    Result := False;
    Ferme(Q);
    Exit;
  end;

  {JP 25/09/03 : Il faut reprendre le libellé de l'auxiliaire et non celui de la banque
                qui est renseigné dans l'enregistrement 05}
  if Assigned(F) then begin
    {Reprose des informations sur la précédente écriture du Tiers}
    T.AddChampSupValeur('C_LIBPAYS', F.GetString('C_LIBPAYS'));
    T.AddChampSupValeur('C_PAYSISO2', F.GetString('C_PAYSISO2'));
  end
  else begin
    Q := OpenSelect('SELECT PY_CODEISO2, PY_LIBELLE FROM PAYS WHERE PY_PAYS = "' + T.GetString('C_PAYS') + '"', T.GetString('SYSDOSSIER'), True);
    if not Q.EOF then begin

      T.AddChampSupValeur('C_LIBPAYS', FullMajuscule(Q.FindField('PY_LIBELLE').AsString));
      T.AddChampSupValeur('C_PAYSISO2', Q.FindField('PY_CODEISO2').AsString);
    end;
    Ferme(Q);
  end;
end;

{Formatage des montants en fonction des fichiers : retourne ne nombre de décimales de la
 devise du montant. A surcharger.
{---------------------------------------------------------------------------------------}
function TObjCFONB.PrepareMontant(var Deb, Cred : Double; var T : TOB) : Integer;
{---------------------------------------------------------------------------------------}
begin
  Result := V_PGI.OkDecV;
{Virtuelle : Formatage des montants}
end;

{Regroupement des tiers sur une ligne
{---------------------------------------------------------------------------------------}
procedure TObjCFONB.SommeTiers;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  n : Integer;
  T : TOB;
  F : TOB;
  OldTiers  : string;
  OldDevise : string;
  OldTicTid : string;
  MntDeb    : Double;
  MntCred   : Double;
  MntDebDev : Double;
  MntCreDev : Double;
  CodeAccept: Boolean;
begin
  if GetString('CUMULE') <> 'X' then Exit;

  MntDeb    := 0;
  MntCred   := 0;
  MntDebDev := 0;
  MntCreDev := 0;
  CodeAccept:= False;

  TobTraitement.Detail.Sort('E_AUXILIAIRE;E_GENERAL;E_DEVISE;');
  T := nil;

  for n := TobTraitement.Detail.Count - 1 downto 0 do begin
    F := TobTraitement.Detail[n];

    if ((OldTiers = F.GetString('E_AUXILIAIRE')) or
        ((OldTicTid = F.GetString('G_NATUREGENE')) and (OldTiers = F.GetString('E_GENERAL')) and
         ((F.GetString('G_NATUREGENE') = 'TIC') or (F.GetString('G_NATUREGENE') = 'TID')))) and
       (OldDevise = F.GetString('E_DEVISE')) then begin
      MntDeb    := F.GetDouble('E_DEBIT'    ) + MntDeb   ;
      MntCred   := F.GetDouble('E_CREDIT'   ) + MntCred  ;
      MntDebDev := F.GetDouble('E_DEBITDEV' ) + MntDebDev;
      MntCreDev := F.GetDouble('E_CREDITDEV') + MntCreDev;

      {Pour les LCR : Si le code accpetation est différent du précédent : Récupère celui du mode de paiement}
      if TypeExport = cfonb_LCR then begin
        if T.GetString('E_DATEECHEANCE') < F.GetString('E_DATEECHEANCE') then
          T.SetString('E_DATEECHEANCE', F.GetString('E_DATEECHEANCE'));

        if not (T.GetString('E_CODEACCEPT') = F.GetString('E_CODEACCEPT')) and not CodeAccept then begin
          CodeAccept := True;
          Q := OpenSelect('SELECT MP_CODEACCEPT FROM MODEPAIE WHERE MP_MODEPAIE = "' + T.GetString('E_MODEPAIE') + '"', T.GetString('SYSDOSSIER'), True);
          if not Q.EOF then T.SetString('E_CODEACCEPT', Q.Fields[0].AsString);
          Ferme(Q);
        end;
      end;
      {On mémorise la clef de la facture pour la mise à jour dans FlagEcriture}
      ListeFac.Add(T.GetString('E_DEVISE') + ';' + OldTicTid + ';' + OldTiers + ';' +
                   T.GetString('E_JOURNAL') + ';' + T.GetString('E_EXERCICE') + ';' + T.GetString('E_DATECOMPTABLE') + ';' +
                   T.GetString('E_NUMEROPIECE') + ';' + T.GetString('E_QUALIFPIECE') + ';' + T.GetString('E_NUMLIGNE') + ';' +
                   T.GetString('E_NUMECHE') + ';');
      FreeAndNil(F);
    end
    else begin
      if Assigned(T) then begin
        T.SetDouble('E_DEBIT'    , MntDeb   );
        T.SetDouble('E_CREDIT'   , MntCred  );
        T.SetDouble('E_DEBITDEV' , MntDebDev);
        T.SetDouble('E_CREDITDEV', MntCreDev);
      end;

      OldDevise := F.GetString('E_DEVISE');
      OldTicTid := F.GetString('G_NATUREGENE');
      if (OldTicTid = 'TID') or (OldTicTid = 'TIC') then OldTiers := F.GetString('E_GENERAL')
                                                    else OldTiers := F.GetString('E_AUXILIAIRE');
      T := F;
      MntDeb    := F.GetDouble('E_DEBIT'    );
      MntCred   := F.GetDouble('E_CREDIT'   );
      MntDebDev := F.GetDouble('E_DEBITDEV' );
      MntCreDev := F.GetDouble('E_CREDITDEV');

      if TypeExport = cfonb_LCR then begin
        if F.GetDateTime('E_DATECOMPTABLE') > GetDateTime('DATEREMISE') then
          F.SetDateTime('E_DATECOMPTABLE', GetDateTime('DATEREMISE'));
      end;
    end;
  end;

  if Assigned(T) then begin
    T.SetDouble('E_DEBIT'    , MntDeb   );
    T.SetDouble('E_CREDIT'   , MntCred  );
    T.SetDouble('E_DEBITDEV' , MntDebDev);
    T.SetDouble('E_CREDITDEV', MntCreDev);
  end;
end;

{Retourne un objet contenant les informations de la devise (à compléter selon les besoins).
 Si la devise n'est pas stockée dans la liste "ListeDev", on la cherche dans la table
{---------------------------------------------------------------------------------------}
function TObjCFONB.GetInfosDevise(Dev : string) : TObjDevise;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  n : Integer;
begin
  n := ListeDev.IndexOf(Dev);
  if n > -1 then
    Result := TObjDevise(ListeDev.Objects[n])
  else begin
    Q := OpenSQL('SELECT D_DECIMALE, D_QUOTITE FROM DEVISE WHERE D_DEVISE = "' + Dev + '"', True);
    try
      if not Q.EOF then begin
        Result := TObjDevise.Create;
        Result.Decimales := Q.FindField('D_DECIMALE').AsInteger;
        Result.Quotite := Q.FindField('D_QUOTITE').AsInteger;
        ListeDev.AddObject(Dev, Result);
      end
      else
        Result := nil;
    finally
      Ferme(Q);
    end;
  end;
end;

{Contrôles de base : fonction surchargeable
{---------------------------------------------------------------------------------------}
function TObjCFONB.ControleOk : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
  Q : TQuery ;
  DiffMP  : Boolean;
  DiffBQE : Boolean;
  OkD     : Boolean;
  OkC     : Boolean;
  OldMP   : string;
  OldBQE  : string;
  CptB    : string;
begin
  Result := False;

  if TypeExport = '' then begin
    CodeErreur := 1002;
    StopperTraitement := True;
    Exit;
  end;

  if GetString('COMPTEBQ') = '' then begin
    CodeErreur := 1003;
    StopperTraitement := True;
    Exit;
  end;

  if ModeSilent and (GetString('REPERTOIRE') = '') then begin
    CodeErreur := 1004;
    StopperTraitement := True;
    Exit;
  end;


  OkD:=False ; OkC:=False ; OldMP:='' ; OldBQE:='' ;
  DiffMP:=False ; DiffBQE:=False ;
  F := nil;

  if TobTraitement.Detail.Count > 0 then
    CodeDevise := TobTraitement.Detail[0].GetString('E_DEVISE');

  for n := 0 to TobTraitement.Detail.Count - 1 dobegin
    F := TobTraitement.Detail[n];

    if (F.GetString('E_DEVISE') <> CodeDevise) and (CodeDevise <> '') then
      CodeDevise := '';

    if F.GetDouble('E_DEBIT' ) <> 0 then OkD := True;
    if F.GetDouble('E_CREDIT') <> 0 then OkC := True;

    if n = 0 then begin
      OldMP  := F.GetString('E_MODEPAIE');
      OldBQE := F.GetString('E_JOURNAL' );
    end
    else begin
      if F.GetString('E_MODEPAIE') <> OldMP  then DiffMP  := True;
      if F.GetString('E_JOURNAL' ) <> OldBQE then DiffBQE := True;
    end;
  end;

  if (not DiffBQE) and (not IsEncaissement) then begin
    Q := OpenSelect('SELECT J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL = "' + OldBQE + '"', F.GetString('SYSDOSSIER'), True);
    if not Q.EOF then begin
      CptB := Q.Fields[0].AsString;
      if (CptB <> '') and (CptB <> GetString('COMPTEBQ')) then DiffBQE := True;
    end;
    Ferme(Q);
  end;

   if not ModeSilent then begin
     if DiffMP then
       if HShowMessage('0;Export;Les échéances sélectionnées ont des modes de paiement différents.'#13 +
                    'Confirmez-vous le traitement ?;Q;YN;Y;Y;', '', '') <> mrYes then begin
         StopperTraitement := True;
         Exit;
       end;

     if ((OkD) and (IsEncaissement)) or ((OkC) and (not IsEncaissement)) then
       if HShowMessage('0;Export;Les échéances sélectionnées comportent des mouvements débiteurs et créditeurs.'#13 +
                       'Confirmez-vous le traitement ?;Q;YN;Y;Y;', '', '') <> mrYes then begin
         StopperTraitement := True;
         Exit;
       end;

     if ((DiffBQE) and (not IsEncaissement)) then
       if HShowMessage('0;Export;Les échéances sélectionnées ont des banques différentes.'#13 +
                    'Confirmez-vous le traitement ?;Q;YN;Y;Y;', '', '') <> mrYes then begin
         StopperTraitement := True;
         Exit;
       end;
   end;

  Result:=True ;
end;

{---------------------------------------------------------------------------------------}
function TObjCFONB.GetNumCFONB(Cat : string) : string;
{---------------------------------------------------------------------------------------}
var
  St  : string;
  Num : Integer;
begin
  St := Cat + Format_String(V_PGI.User,3) + FormatDateTime('YYMMDD', Date);
  Num := GetParamSocSecur('SO_CPNUMCFONB', 0) + 1;
  if Num >= 100000 then Num := 1;
  SetParamSoc('SO_CPNUMCFONB', Num);
  Result := St + FormatFloat('00000', Num);
end;

{---------------------------------------------------------------------------------------}
procedure TObjCFONB.FlagEcriture;
{---------------------------------------------------------------------------------------}
var
  n  : Integer;
  F  : TOB;
  Nb : Integer;
  SWhere     : string;
  RIB        : string;
  SQL        : string;
  NumCFONB   : string;
  NowFutur   : TDateTime;
  lStDossier : string;

    {---------------------------------------------------------------------------------------}
    procedure MajWhereParListe;
    {---------------------------------------------------------------------------------------}
    var
      e : string;
      s : string;
      c : string;
      D : string;
      T : string;
    begin
      e := ListeFac[n];
      D := ReadTokenSt(e);
      s := ReadTokenSt(e);
      if (s = 'TIC') or (s = 'TID') then c := 'E_GENERAL'
                                    else c := 'E_AUXILIAIRE';
      T := ReadTokenSt(e);

      SWhere := 'E_JOURNAL = "'         + ReadTokenSt(e) + '"' +
             ' AND E_EXERCICE = "'      + ReadTokenSt(e) + '"' +
             ' AND E_DATECOMPTABLE = "' + UsDateTime(StrToDate(ReadTokenSt(e))) + '"' +
             ' AND E_NUMEROPIECE ='     + ReadTokenSt(e) +
             ' AND E_QUALIFPIECE = "'   + ReadTokenSt(e) + '"' +
             ' AND E_NUMLIGNE ='        + ReadTokenSt(e) +
             ' AND E_NUMECHE ='         + ReadTokenSt(e);
      {Recherche du Rib dans l'enregistrement cumulé sur l'auxiliaire}
      //F := TobTraitement.FindFirst([c, 'E_DEVISE'], [D, T], True);
      F := TobTraitement.FindFirst(['E_DEVISE', c], [D, T], True);
      if Assigned(F) then RIB := FullMajuscule(F.GetValue('E_RIB'))
                     else RIB := '';
    end;

    {---------------------------------------------------------------------------------------}
    function MajEcriture : Boolean;
    {---------------------------------------------------------------------------------------}
    begin
      Result := True;

      if EstMultiSoc then begin
        if F.GetNumChamp('SYSDOSSIER') > 0 then
          lStDossier := F.GetString('SYSDOSSIER');
        if (F.GetString('E_SOCIETE') <> GetParamSocSecur('SO_SOCIETE', '')) then
          lStDossier := GetSchemaName(F.GetString('E_SOCIETE'));
        if lStDossier = '' then Exit;
      end
      else
        lStDossier := V_PGI.SchemaName;

      if Trim(RIB)<>'' then begin
        SQL := 'UPDATE ' + GetTableDossier(lStDossier, 'ECRITURE') +
               ' SET E_CFONBOK = "X", E_NUMCFONB="' + NumCFONB + '", E_RIB ="' + RIB +
               '", E_DATEMODIF = "' + UsTime(NowFutur) + '" WHERE ' + SWhere;
        Nb  := ExecuteSQL(SQL) ;
        if Nb <> 1 then begin
          CodeErreur := 1005;
          StopperTraitement := True;
          Result := False;
        end;
      end;
    end;

begin

  if TobTraitement.Detail.Count <= 0 then Exit;

  NowFutur := NowH ;
  NumCFONB := GetNumCFONB(TypeExport) ;

  BeginTrans;
  try
    for n := 0 to TobTraitement.Detail.Count-1 do begin
      F := TobTraitement.Detail[n];
      SWhere := whereEcritureTob(tsGene, F, True, F.GetString('J_MODESAISIE') = 'X', 'E');
      RIB    := FullMajuscule(F.GetValue('E_RIB'));

      if not MajEcriture then Break;
    end;

    for n := 0 to ListeFac.Count - 1 do begin
      {Constitution de la clause Where et initialisation du RIB}
      MajWhereParListe;
      if not MajEcriture then Break;
    end;
  finally
    if StopperTraitement then RollBack
                         else CommitTrans;
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjCFONB.GetMessageErreur : string;
{---------------------------------------------------------------------------------------}
begin
  case CodeErreur of
    1000 : Result := TraduireMemoire('Le fichier de sortie n''est pas renseigné');
    1001 : Result := TraduireMemoire('La nature du compte général est absente de la liste des champs');
    1002 : Result := TraduireMemoire('Veuillez choisir un type d''export.');
    1003 : Result := TraduireMemoire('Veuillez choisir un compte bancaire.');
    1004 : Result := TraduireMemoire('Veuillez choisir un fichier de sortie.');
    1005 : Result := TraduireMemoire('Mise à jour de la table écriture impossible.');

    else if CodeErreur < 0 then begin
      Result := MC_MsgErrDefaut(CodeErreur * -1);
    end;
  end;
end;

                             { ------------------- }
                             {    TOBJTRANSFERT    }
                             { ------------------- }

{---------------------------------------------------------------------------------------}
constructor TObjTransfert.Create(LeNomTable: string; LeParent: TOB; IndiceFils: Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  IsEncaissement := False;
  AddChampSup('IMPUTATION', False);
  AddChampSup('REMISE'    , False);
  AddChampSup('NATECO'    , False);
  AddChampSup('REFERENCE' , False);
  AddChampSup('DATEREMISE', False, CSTDate);
  AddChampSup('CUMULE'    , False, CSTBoolean);
  AddChampSup('RIBPRINC'  , False, CSTBoolean);
  AddChampSup('REFTIRELIB', False, CSTBoolean);
end;

{---------------------------------------------------------------------------------------}
function TObjTransfert.GetImputation : string;
{---------------------------------------------------------------------------------------}
begin
       if GetString('IMPUTATION') = 'BEN' then Result := '13'
  else if GetString('IMPUTATION') = 'EME' then Result := '15'
  else if GetString('IMPUTATION') = 'ABE' then Result := '14'
                                          else Result := '15';
end;

{---------------------------------------------------------------------------------------}
function TObjTransfert.GetRemise : Char;
{---------------------------------------------------------------------------------------}
begin
       if GetString('REMISE') = 'DEV' then Result := '3'
  else if GetString('REMISE') = 'GLO' then Result := '1'
  else if GetString('REMISE') = 'UNI' then Result := '2'
                                      else Result := '2';
end;

{---------------------------------------------------------------------------------------}
function TObjTransfert.GetMessageErreur : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetMessageErreur;
  if Result <> '' then Exit;

  case CodeErreur of
    0  : Result := '';
    1  : Result := TraduireMemoire('Le code BIC de la banque sélectionnée n''est pas renseigné.');
    2  : Result := TraduireMemoire('Le code IBAN de la banque sélectionnée n''est pas renseigné.');
    3  : Result := TraduireMemoire('Le code devise de la banque sélectionnée n''est pas renseigné.');
    4  : Result := TraduireMemoire('La référence interne de la pièce est absente de l''ensemble de données.');
    5  : Result := TraduireMemoire('Impossible de récupérer le code IBAN de certains tiers.');
    6  : Result := TraduireMemoire('');

    else
      Result := TraduireMemoire('Traitement interrompu');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjTransfert.MajEntete;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited;
  Q := OpenSelect('SELECT * FROM BANQUECP WHERE BQ_GENERAL = "' + GetString('COMPTEBQ')
                 +'" AND BQ_NODOSSIER = "'+V_PGI.NoDossier+'"', V_PGI.SchemaName, True); // 19/10/2006 YMO Multisociétés
  if not Q.EOF then begin
    AddChampSupValeur('BQ_CODEBIC' , Q.FindField('BQ_CODEBIC' ).AsString);
    AddChampSupValeur('BQ_CODEIBAN', Q.FindField('BQ_CODEIBAN').AsString);
    AddChampSupValeur('BQ_DEVISE'  , Q.FindField('BQ_DEVISE'  ).AsString);
    AddChampSupValeur('BQ_PAYS'    , Q.FindField('BQ_PAYS'    ).AsString);
    AddChampSupValeur('BQ_LIBELLE' , Q.FindField('BQ_LIBELLE' ).AsString);
  end;
  Ferme(Q);

       if GetString('BQ_CODEBIC' ) = '' then CodeErreur := 1
  else if GetString('BQ_CODEIBAN') = '' then CodeErreur := 2
  else if GetString('BQ_DEVISE'  ) = '' then CodeErreur := 3;
  if GetString('BQ_PAYS') = '' then SetString('BQ_PAYS', 'FRA');

  if CodeErreur > 0 then begin
    StopperTraitement := True;
    Exit;
  end;

  if CodeDevise = '' then AddChampSupValeur('TYPEFICHIER' , '2')
                     else AddChampSupValeur('TYPEFICHIER' , '1');
  AddChampSupValeur('DATECREATION', Date);
  AddChampSupValeur('CODEDEVISE' , CodeDevise);
  AddChampSupValeur('SO_LIBELLE' , FullMajuscule(GetParamsocSecur('SO_LIBELLE'   , '')));
  AddChampSupValeur('SO_ADRESSE1', FullMajuscule(GetParamsocSecur('SO_ADRESSE1'  , '')));
  AddChampSupValeur('CODE_VILLE' , FullMajuscule(GetParamsocSecur('SO_CODEPOSTAL', '')) + ' ' +
                                   FullMajuscule(GetParamsocSecur('SO_VILLE'     , '')));
  AddChampSupValeur('SO_PAYS'    , FullMajuscule(GetParamsocSecur('SO_PAYS'      , '')));
  AddChampSupValeur('SO_SIRET'   , StringReplace(GetParamsocSecur('SO_SIRET'     , ''), ' ', '', [rfReplaceAll]));
  AddChampSupValeur('SO_CPIDCLIENT', GetParamsocDossierSecur('SO_CPIDCLIENT'     , 0 ));
end;

{---------------------------------------------------------------------------------------}
procedure TObjTransfert.MajLigne;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  F : TOB;
  n : Integer;
  A : string;
  G : string;
  D : string;
begin
  inherited;
  F := nil;

  TobTraitement.Detail.Sort('E_AUXILIAIRE;E_GENERAL;E_DEVISE;');
  for n := TobTraitement.Detail.Count - 1 downto 0 do begin
    T := TobTraitement.Detail[n];
    if (T.GetString('E_AUXILIAIRE') <> A) or (D <> T.GetString('E_DEVISE')) or
       ((T.GetString('E_GENERAL') <> G) and ((T.GetString('G_NATUREGENE') = 'TIC') or (T.GetString('G_NATUREGENE') = 'TID'))) then begin
      A := T.GetString('E_AUXILIAIRE');
      G := T.GetString('E_GENERAL');
      D := T.GetString('E_DEVISE');
      F := T;
      MajTobFille(T, nil);
    end
    else
      MajTobFille(T, F);
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjTransfert.PrepareMontant(var Deb, Cred : Double; var T : TOB) : Integer;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjDevise;
begin
  Result := inherited PrepareMontant(Deb, Cred, T);
  {On ne travaille pas sur des mouvements créditeurs pour l'émetteur}
  if (Cred > 0) or (Deb < 0) then begin
    GereRefus(TraduireMemoire('L''écriture est créditrice pour l''émetteur'), T);
  end
  else begin
    {On va travailler sur le debit qui récupère le montant négatif du crédit}
    if Cred < 0 then Deb := Cred * -1;
    Obj := GetInfosDevise(T.GetString('E_DEVISE'));
    if Assigned(Obj) then
      Result := Obj.Decimales;

    Deb := Round(Deb * IntPower(10, Result));
    T.AddChampSupValeur('MONTANTDEV', Deb);
    T.AddChampSupValeur('NBDECIMALES', Result);
  end;
end;

{Vérifie et complète la ligne avec les champs nécessaires
{---------------------------------------------------------------------------------------}
procedure TObjTransfert.MajTobFille(var T : TOB; F : TOB);
{---------------------------------------------------------------------------------------}
var
  D : Double;
  C : Double;
begin
  D := T.GetDouble('E_DEBITDEV');
  C := T.GetDouble('E_CREDITDEV');
  PrepareMontant(D, C, T);
  
  T.AddChampSupValeur('BQ_CODEBIC'   , GetString('BQ_CODEBIC'   ));
  T.AddChampSupValeur('BQ_CODEIBAN'  , GetString('BQ_CODEIBAN'  ));
  T.AddChampSupValeur('BQ_DEVISE'    , GetString('BQ_DEVISE'    ));
  T.AddChampSupValeur('SO_LIBELLE'   , GetString('SO_LIBELLE'   ));
  T.AddChampSupValeur('SO_ADRESSE1'  , GetString('SO_ADRESSE1'  ));
  T.AddChampSupValeur('CODE_VILLE'   , GetString('CODE_VILLE'   ));
  T.AddChampSupValeur('SO_PAYS'      , GetString('SO_PAYS'      ));
  T.AddChampSupValeur('SO_SIRET'     , GetString('SO_SIRET'     ));
  T.AddChampSupValeur('SO_CPIDCLIENT', GetString('SO_CPIDCLIENT'));
  T.AddChampSupValeur('NATECO'       , GetString('NATECO'       ));
  T.AddChampSupValeur('REFERENCE'    , GetString('REFERENCE'    ));
  T.AddChampSupValeur('TYPEFICHIER'  , GetString('TYPEFICHIER'  ));
  T.AddChampSupValeur('CODEDEVISE'   , GetString('CODEDEVISE'   ));
  T.AddChampSupValeur('CUMULE'       , GetString('CUMULE'       ));
  T.AddChampSupValeur('DATECREATION', GetDateTime('DATECREATION'));
  T.AddChampSupValeur('DATEREMISE'  , GetDateTime('DATEREMISE'  ));
  T.AddChampSupValeur('IMPUTATION'  , GetImputation);
  T.AddChampSupValeur('REMISE'      , GetRemise);
  if not GetInfosTiers(T) then Exit;

  if not GetInfosRib(T, True, F) then Exit;

  if T.GetNumChamp('E_REFINTERNE') > -1 then
    T.SetString('E_REFINTERNE', FullMajuscule(T.GetString('E_REFINTERNE')))
  else
    GereRefus(TraduireMemoire('Référence interne non renseignée'), T);
end;


                             { ------------------- }
                             {   TOBJPRELEVEMENT   }
                             { ------------------- }

{---------------------------------------------------------------------------------------}
constructor TObjPrelevement.Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
  {Bien penser à initialiser cette variable}
  IsEncaissement := False;
  {Création des champs figurants dans la fiche CPCRITERECFONB
  AddChampSup('DATE'  , False);
  AddChampSup('REMISE', False);
  ...}
end;

{---------------------------------------------------------------------------------------}
function TObjPrelevement.GetMessageErreur : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetMessageErreur;
  if Result <> '' then Exit;

  case CodeErreur of
    0  : Result := '';
    1  : Result := TraduireMemoire('');
    2  : Result := TraduireMemoire('');
    3  : Result := TraduireMemoire('');
    else
      Result := TraduireMemoire('Traitement interrompu');
  end;

end;

                             { ------------------- }
                             {    TOBJVIREMENT     }
                             { ------------------- }


{---------------------------------------------------------------------------------------}
constructor TObjVirement.Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
  {Bien penser à initialiser cette variable}
  IsEncaissement := False;
  {Création des champs figurants dans la fiche CPCRITERECFONB
  AddChampSup('DATE'  , False);
  AddChampSup('REMISE', False);
  ...}
end;

{---------------------------------------------------------------------------------------}
function TObjVirement.GetMessageErreur : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetMessageErreur;
  if Result <> '' then Exit;

  case CodeErreur of
    0  : Result := '';
    1  : Result := TraduireMemoire('');
    2  : Result := TraduireMemoire('');
    else
      Result := TraduireMemoire('Traitement interrompu');
  end;

end;

                             { ------------------- }
                             {       TObjLCR       }
                             { ------------------- }

{---------------------------------------------------------------------------------------}
constructor TObjLCR.Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
  {Bien penser à initialiser cette variable}
  IsEncaissement := True;
  {Création des champs figurants dans la fiche CPCRITERECFONB
  AddChampSup('DATE'  , False);
  AddChampSup('REMISE', False);
  ...}
end;

{---------------------------------------------------------------------------------------}
function TObjLCR.GetMessageErreur : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetMessageErreur;
  if Result <> '' then Exit;

  case CodeErreur of
    0  : Result := '';
    1  : Result := TraduireMemoire('');
    2  : Result := TraduireMemoire('');
    else
      Result := TraduireMemoire('Traitement interrompu');
  end;

end;


{$IFDEF CISXPGI}
                             { ------------------- }
                             {    TObjCISX         }
                             { ------------------- }
{---------------------------------------------------------------------------------------}
constructor TObjCISX.Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer);
{---------------------------------------------------------------------------------------}
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
  {Bien penser à initialiser cette variable}
  IsEncaissement := False;
  {Création des champs figurants dans la fiche CPCRITERECFONB
  AddChampSup('DATE'  , False);
  AddChampSup('REMISE', False);
  ...}
end;

{---------------------------------------------------------------------------------------}
function TObjCISX.GetMessageErreur : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetMessageErreur;
  if Result <> '' then Exit;

  case CodeErreur of
    0  : Result := '';
    1  : Result := TraduireMemoire('');
    2  : Result := TraduireMemoire('');
    else
      Result := TraduireMemoire('Traitement interrompu');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCISX.GetInfosSup;
{---------------------------------------------------------------------------------------}
var
  i,j       : integer;
  Compte    : string;
  SQL       : string;
  Q         : TQuery;
  numemet   : string;
  etabbq    : string;
  guichet   : string;
  numcpte   : string;
  clerib    : string;
  codeiban  : string;
  devise    : string;
  Reference : string;
  PaysISO   : string;
  Ref       : extended;
  Cle       : integer;
  cumul     : extended;
begin
  // Infos Banques :
  SQL := ' SELECT BQ_NUMEMETVIR,BQ_ETABBQ,BQ_GUICHET,BQ_NUMEROCOMPTE,BQ_CLERIB,BQ_CODEIBAN,BQ_DEVISE ' +
         ' FROM BANQUECP ' +
         ' WHERE BQ_GENERAL = "' + Banque + '"';
  Q := OpenSQL(SQL,true);
  try
     if not(Q.eof) then
     begin
       numemet   := Q.FindField('BQ_NUMEMETVIR').AsString;
       etabbq    := Q.FindField('BQ_ETABBQ').AsString;
       guichet   := Q.FindField('BQ_GUICHET').AsString;
       numcpte   := Q.FindField('BQ_NUMEROCOMPTE').AsString;
       clerib    := Q.FindField('BQ_CLERIB').AsString;
       codeiban  := Q.FindField('BQ_CODEIBAN').AsString;
       devise    := Q.FindField('BQ_DEVISE').AsString;
     end;
  finally
     ferme(Q);
  end;
  // Rajout du champ
  cumul := 0;
  TobTraitement.Detail[0].AddChampSup('NB_COMPTE',true);
  TobTraitement.Detail[0].AddChampSup('CODE_NATURE',true);
  TobTraitement.Detail[0].AddChampSup('TYPE_COMPTE',true);
  TobTraitement.Detail[0].AddChampSup('T_CODEISOPAYS',true);
  for i := 0 to TobTraitement.Detail.Count - 1 do
  begin

     // Champs banque
     TobTraitement.Detail[i].PutValue('BQ_NUMEMETVIR',numemet);
     TobTraitement.Detail[i].PutValue('BQ_ETABBQ',etabbq);
     TobTraitement.Detail[i].PutValue('BQ_GUICHET',guichet);
     TobTraitement.Detail[i].PutValue('BQ_NUMEROCOMPTE',numcpte);
     TobTraitement.Detail[i].PutValue('BQ_CLERIB',clerib);
     TobTraitement.Detail[i].PutValue('BQ_CODEIBAN',codeiban);
     TobTraitement.Detail[i].PutValue('BQ_DEVISE',devise);

     // On corrige les montants
     TobTraitement.Detail[i].PutValue('E_DEBIT',GetValueMontant(StrToFloat(TobTraitement.Detail[i].GetValue('E_DEBIT'))));

     // Type Compte :
     Compte := TobTraitement.Detail[i].GetValue('E_RIB');
     if (Pos('*',Compte) <> 1) or (Script = 'VIRBELGE') then  
     begin
        // RIB
        TobTraitement.Detail[i].PutValue('TYPE_COMPTE','2');
        if (Pos('*',Compte) <> 1) then
          Compte := Trim(copy(Compte,1,12))
        else
          Compte := Trim(copy(Compte,2,12));
        for j := 1 to length(Compte) do
            if Compte[j] = ' ' then Compte[j] := '0';
        if (Pays = 'BEL') and isNumeric(Compte) then
        begin
           // Cumul de compte (Fait dans CISX finallement... et non ...)
           cumul := cumul + StrToFloat(Compte);
           if cumul > 999999999999999 then
           begin
              cumul := cumul - int(cumul/1000000000000000) * 1000000000000000;
           end;
        end;
     end
     else
     begin
        System.Delete(Compte,1,length(Compte)-1);
        if IsIBAN(Compte) then
           // IBAN
           TobTraitement.Detail[i].PutValue('TYPE_COMPTE','1')
        else
           // BBAN ou autre   
           TobTraitement.Detail[i].PutValue('TYPE_COMPTE','3');
     end;

     // Pays ISO :
     PaysISO := TobTraitement.Detail[i].GetValue('T_PAYS');
     TobTraitement.Detail[i].PutValue('T_CODEISOPAYS',CodeIsoDuPays(PaysISO));

     // Nature :
     Reference := TobTraitement.Detail[i].GetValue('E_REFINTERNE');
     if not(isNumeric(Reference)) or (Length(Reference) <> 10) then
        // Communication non structurée
        TobTraitement.Detail[i].PutValue('CODE_NATURE','3')
     else
     begin
        // Communication structurée
        TobTraitement.Detail[i].PutValue('CODE_NATURE','8');
        Ref := StrToFloat(Reference);
        Cle := Round(Ref - (Round(Int(Ref / 97)) * 97));
        if Cle < 10 then
           TobTraitement.Detail[i].PutValue('E_REFINTERNE',Reference + '0' + IntToStr(Cle))
        else
           TobTraitement.Detail[i].PutValue('E_REFINTERNE',Reference + IntToStr(Cle));
     end;
  end;
  // Gestion des Nb Compte Belge
  //SDA le 27/12/2007
  if cumul <> 0 then
  //Fin SDA le 27/12/2007
    TobTraitement.Detail[0].PutValue('NB_COMPTE',FloatToStr(cumul))
  //SDA le 27/12/2007
  else
    TobTraitement.Detail[0].PutValue('NB_COMPTE','000000000000000');
  //Fin SDA le 27/12/2007
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 02/02/2007
Modifié le ... :   /  /
Description .. : Procédure permettant de réaliser l'export en créant son
Suite ........ : environnement.
Mots clefs ... :
*****************************************************************}
procedure TObjCISX.LanceTraitement ;
Var
  ListVar   : TStringList;
  FicINI    : String;
  FicTOB    : String;
  i         : integer;
begin
  { On recupere les informations supplementaires }
  GetInfosSup;
  { Exécute les contrôles de base}
  if not StopperTraitement then ControleOk;
  { Création du répertoire }
  {SDA le 14/12/2007 if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
     gDirEchDuDos := V_PGI.DosPath + '\Echanges'
  else
     gDirEchDuDos := V_PGI.DosPath + 'Echanges'; }

  //SDA le 14/12/2007
  if TcbpPath.GetCegidUserTempPath[length (TcbpPath.GetCegidUserTempPath)] <> '\' then
     gDirEchDuDos := TcbpPath.GetCegidUserTempPath + '\Echanges'
  else
     gDirEchDuDos := TcbpPath.GetCegidUserTempPath + 'Echanges';
  //Fin SDA le 14/12/2007

  if Not DirectoryExists (gDirEchDuDos) then
     CreateDir (gDirEchDuDos);
  { Sauvegarde de la TOB }
  if not StopperTraitement then
  begin
     FicTOB := gDirEchDuDos + '\VIREMENT_CISX.txt';
     TobTraitement.SaveToFile(FicTOB,false,true,true,'');
  end;

  if not StopperTraitement then
  begin
     FicINI := gDirEchDuDos + ConstFicIniDemande;
     SysUtils.DeleteFile (FicIni);
  end;

  { Gestion Des Variables }
  if not StopperTraitement then
  begin
     Variables := TOB.Create('Mon Export',nil,-1);
     try ListVar := TStringList.Create ;
        // Création du fichier ini avec les variables.
        InitScriptAuto(Script);
        //SDA le 14/12/2007 if TobLoadFromFile (gDirEchDuDos + ConstFicIniRetour, nil, Variables) then
        if TobLoadFromFile (TcbpPath.GetCegidUserTempPath + 'Echanges\' + ConstFicIniRetour, nil, Variables) then // SDA le 14/12/2007
        begin
          // Saisie des variables demandables
          SaisieVariable(Variables);
          // Récupération des variables demandables
          for i := 0 to Variables.detail.Count -1 do
          begin
            if (Variables.Detail[i].FieldExists ('DEMANDABLE')) and
               (Variables.Detail[i].GetValue ('DEMANDABLE') = '-1') OR
               (Variables.Detail[i].GetValue ('DEMANDABLE') = 'True') then  //SDA le 14/12/2007
            begin
               ListVar.Add(Variables.Detail[i].GetValue ('NAME') + '=' + Variables.Detail[i].GetValue ('TEXT'));
            end;
          end;
        end;
        LancerDialog ;
        { Création du fichier INI }
        if not StopperTraitement then CreeFicIniPourCISX('TRA', 'X' +  Script, FicTOB, GetString('REPERTOIRE'), ListVar);
        { Lancement de l'export }
        if not StopperTraitement then LanceExportCisx ('/INI='+FicINI, 'EXPORT', Pays);
     finally
        FreeAndNil(ListVar);
        FreeAndNil(Variables);
        SysUtils.DeleteFile(FicTOB);
     end;
  end;
  {Mise à jour de la table écriture}
  if not StopperTraitement then FlagEcriture;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/02/2007
Modifié le ... :   /  /
Description .. : Gestion des variables demandables
Mots clefs ... : 
*****************************************************************}
function TObjCISX.SaisieVariable (TV : TOB): Boolean;
var
Ed1           : THCritMaskEdit;
Lb            : THLabel;
St            : string;
Lasth,Lastlh  : integer;
id            : integer;
Demandable    : Boolean;
begin
  Result := TRUE;
  FSaisieVar := TFVierge.Create(nil);
  FSaisieVar.BValider.Onclick := BValideClickVar;
  with FSaisieVar do
  begin
     Caption := TraduireMemoire('Saisie des variables :');
     BorderStyle := bsDialog;
     Position := poScreenCenter;
     ClientWidth := 324;
     ClientHeight := 200+(TV.detail.count*3);
     Lasth := 21;
     Lastlh := 25;
     left := length(TV.detail[0].getvalue('Libelle'));
     For id:=0 To TV.detail.count-1 do
     begin
        if left > length(TV.detail[id].getvalue('Libelle')) then
           left := length(TV.detail[id].getvalue('Libelle'));
     end;
     Demandable := FALSE;
     For id:=0 To TV.detail.count-1 do
     begin
        if (not TV.detail[id].getvalue('demandable')) then continue;
        Demandable := TRUE;
        Lb :=  THLabel.create(nil);
        Lb.Height := 13;
        Lb.Left := 12;
        Lb.Top := Lastlh;
        Lb.Width := length(TV.detail[id].getvalue('Libelle'));
        St := TV.detail[id].getvalue('Name');
        St := ReadTokenpipe (St,'=');

        Lb.caption := TV.detail[id].getvalue('Libelle');
        Lb.Parent := FSaisieVar;

        Ed1 :=  THCritMaskEdit.create(nil);
        Ed1.Name := St;
        Ed1.tag := 100;
        Ed1.Height := 21;
        Ed1.Left := 130+ left;
        Ed1.Top := Lasth;
        Ed1.text := TV.detail[id].getvalue('Text');
        Ed1.Width := length(TV.detail[id].getvalue('Text'))+Ed1.Left;
        Ed1.Parent := FSaisieVar;
        Lasth:=Ed1.Top+Ed1.Height+2;
        Lastlh:=Ed1.Top+Ed1.Height+5;

        // Gestion des variables spécifique
        if (Ed1.Name = 'DATECREATE') then
           Ed1.Text := FormatDateTime('DDMMYY',now);
     end;
     try
         if (Demandable) and (ShowModal <> mrOk) then
              Result      := FALSE;
     finally
         Free;
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/02/2007
Modifié le ... :   /  /
Description .. : Bouton valider de la fiche des variables
Mots clefs ... : 
*****************************************************************}
procedure TObjCISX.BValideClickVar(Sender: TObject);
var
id,idv    : integer;
TL        : TControl ;
begin
  inherited;
  for id:=0 To Variables.detail.count-1 do
  begin
     for idv:=0 to FSaisieVar.ControlCount-1 do
     begin
        TL:=FSaisieVar.Controls[idv] ;
        if (TL is THCritMaskEdit) and (TL.Name = Variables.detail[id].GetValue('Name')) and (TL.tag = 100) Then
        begin
           Variables.detail[id].PutValue('Text',THCritMaskEdit(TL).Text);
           break;
        end ;
     end;
  end;
end;     
{$ENDIF CISXPGI}


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/02/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de charger un THValCombobox en lui 
Suite ........ : passant en parametre la requete SQL.
Suite ........ : Le premier champs de la requete doit être la valeur
Suite ........ : retournée
Suite ........ : Le second champs la valeur affichée.
Mots clefs ... : 
*****************************************************************}
procedure ChargeTHValComboBox ( SQL : string ; var Combo : THValComboBox);
Var
  Q : TQuery;
  Values : hTStringList;
  Items : hTStringList;
begin
  try Q := OpenSQL(SQL,true);
     Values := hTStringList.Create;
     Items  := hTStringList.Create;
     while not(Q.eof) do
     begin
       Values.Add(Q.Fields[0].AsString);
       Items.Add( Q.Fields[1].AsString);
       Q.Next;
     end;
  finally
     Ferme(Q);
  end;
  Combo.Values := Values;
  Combo.Items  := Items;
  Combo.Refresh;
  Combo.ItemIndex := 0;  
  FreeAndNil(Values);
  FreeAndNil(Items);
end;

function GetCurrentProcessEnvVar(const VariableName: string): string;
{var
  nSize: dWord;
begin
  nSize:= 0;
  nSize:= GetEnvironmentVariable(PChar(VariableName), nil, nSize);
  if nSize = 0 then
    result:= ''
  else
  begin
    SetLength(result, nSize - 1);
    if GetEnvironmentVariable(PChar(VariableName), PChar(result), nSize) <> nSize - 1 then
      raise Exception.Create(SysErrorMessage(GetlastError))
  end;}
begin
  Result := GetEnvironmentVariable(PChar(VariableName));
end;
end.
