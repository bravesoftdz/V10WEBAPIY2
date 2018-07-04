{ Unité : Source TOF de la FICHE : TRQRJALPORTEFEUIL : Journal des portefeuilles
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 8.10.001.006  23/08/07  JP   Création du journal des portefeuilles
--------------------------------------------------------------------------------------}
unit TRQRJALPORTEFEUIL_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eQRS1, MaineAGL,
  {$ELSE}
  QRS1, FE_Main,
  {$ENDIF}
  UTob, SysUtils, HCtrls, uAncetreEtat;

type
  TOF_TRQRJALPORTEFEUIL = class (TRANCETREETAT)
    procedure OnArgument(S : string); override;
    procedure OnLoad                ; override;
    procedure OnClose               ; override;
    procedure OnNew                 ; override;
  private
    FBoEtatAuto : Boolean;
    TobTaux     : TOB;
    TobTmp      : TOB;
    UniqueKey   : Integer;

    procedure InitFiche(Arg : string);
    procedure RemplitTableTempo;
    procedure VideTableTempo   ;
    procedure CalculVolume     ;
    procedure CalculCumul      ;
    function  GetSQLRequetePrinc : string;
    function  GetSQLRequeteCumul : string;
    function  GetSQLRequeteTaux  : string;
    function  GetTaux(aOPCVM : string) : Double;
  public
    edPortefeuille  : THEdit;
    edPortefeuille_ : THEdit;
    edOPCVM         : THEdit;
    edOPCVM_        : THEdit;
    edOPE           : THEdit;
    edOPE_          : THEdit;

    procedure PortefChange(Sender : TObject);
  end;

procedure TRLanceFiche_JalPortefeuille(Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  HMsgBox, uTobDebug, Constantes, HEnt1, Commun, UProcGen;

const
  CHP_PORTEFEUIL = 'CED_NATURE';
  CHP_OPCVM      = 'CED_COMPTE2';
  CHP_PMVNETTE   = 'CED_DEBIT1';
  CHP_MNTVENTE   = 'CED_DEBIT2';
  CHP_MNTACHAT   = 'CED_DEBIT3';
  CHP_RENDJOUR   = 'CED_DEBIT4';
  CHP_BASEANNUEL = 'CED_DEBIT5';
  CHP_NBJOURSDT  = 'CED_DEBIT6';

  CHP_PARTSFIN   = 'CED_CREDIT1';
  CHP_COURSFIN   = 'CED_CREDIT2';
  CHP_PARTSCOU   = 'CED_CREDIT3';
  CHP_COURSCOU   = 'CED_CREDIT4';
  CHP_CUMACHAT   = 'CED_CREDIT5';
  //CHP_MNTACHFIN  = 'CED_CREDIT6';


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_JalPortefeuille(Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('TR', 'TRQRJALPORTEFEUIL', Range, Lequel, Arguments);
end;

{Les arguments : DateDeb ; DateFin ; Portefeuille ; OPCVM ;
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := '';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000161;
  {Initialisation des controls et de leur évènements}
  InitFiche(S);
  {On vide la table CEDTBALANCE}
  VideTableTempo;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.OnNew;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if FBoEtatAuto then begin
    TFQRS1(Ecran).bAgrandir.Click;
    TFQRS1(Ecran).BValider.Click;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  RemplitTableTempo;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {On vide la table CEDTBALANCE}
  VideTableTempo;
end;

{Gestion du filtre sur les portefeuilles
{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.PortefChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ch  : string;
  Ch_ : string;
begin
  Ch  := edPortefeuille .Text;
  Ch_ := edPortefeuille_.Text;
       if (Ch <> '') and (ch_ = '' ) then Ch := 'TOF_PORTEFEUILLE >= "' + Ch + '"'
  else if (Ch = '' ) and (ch_ <> '') then Ch := 'TOF_PORTEFEUILLE <= "' + Ch_ + '"'
  else if (Ch = '' ) and (ch_ = '' ) then Ch := ''
  else if (Ch <> '') and (ch_ <> '') then Ch := 'TOF_PORTEFEUILLE BETWEEN "' + Ch + '" AND "' + Ch_ + '"';

  edOPCVM .Plus := Ch;
  edOPCVM_.Plus := Ch;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.InitFiche(Arg : string);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  edPortefeuille  := (GetControl('PORTEFEUILLE' ) as THEdit);
  edPortefeuille_ := (GetControl('PORTEFEUILLE_') as THEdit);
  edOPCVM         := (GetControl('OPCVM'        ) as THEdit);
  edOPCVM_        := (GetControl('OPCVM_'       ) as THEdit);
  edOPE           := (GetControl('DTOPE'        ) as THEdit);
  edOPE_          := (GetControl('DTOPE_'       ) as THEdit);

  {Filtre des OPCVM en fonction du portefeuille}
  edPortefeuille .OnChange := PortefChange;
  edPortefeuille_.OnChange := PortefChange;

  FBoEtatAuto := False;
  if Trim(Arg) <> '' then begin
    FBoEtatAuto := True;
    {Les arguments : DateDeb ; DateFin ; Portefeuille ; OPCVM ;}
    s := ReadTokenSt(Arg);
    if (s <> '') and IsValidDate(s) then edOPE.Text := s;
    s := ReadTokenSt(Arg);
    if (s <> '') and IsValidDate(s) then edOPE_.Text := s;
    s := ReadTokenSt(Arg);
    if (s <> '') then begin
      edPortefeuille.Text := s;
      edPortefeuille_.Text := s;
    end;
    s := ReadTokenSt(Arg);
    if (s <> '') then begin
      edOPCVM.Text := s;
      edOPCVM_.Text := s;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.VideTableTempo;
{---------------------------------------------------------------------------------------}
begin
  ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"');
end;

{Récupération de tous les mouvements sur la période qu'il s'agisse de ventes ou d'achats
 afin d'en afficher le détail par date dans l'état : d'où la requête UNION à la différence
 de ce qui est fait dans GetSQLRequeteCumul
{---------------------------------------------------------------------------------------}
function TOF_TRQRJALPORTEFEUIL.GetSQLRequetePrinc : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'SELECT "A" AS CED_TYPE, TOP_PORTEFEUILLE AS CED_PORTEFEUILLE, TOP_CODEOPCVM AS CED_CODEOPCVM, TOP_NUMOPCVM AS CED_NUMTRANSAC, ' +
            'TOP_DATEACHAT AS CED_DATEOPERATION, TOP_NBPARTACHETE AS CED_NBACHAT, (TOP_PRIXUNITAIRE / TOP_COTATIONACH) AS CED_COURSACH, ';
  Result := Result + '"' + UsDateTime(iDate1900) + '" AS CED_DATECOMPL, ';
  Result := Result + '0 AS CED_COURSVEN, 0 AS CED_NBVENTE, 0 AS CED_PMV, 0 AS CED_RENDEMENT, TOP_BASECALCUL AS CED_BASE, ';
  Result := Result + 'TROPCVMREF.TOF_LIBELLE AS CED_LIBOPCVM, CHOIXCOD.CC_LIBELLE AS CED_LIBPORTEFEUILLE ';
  Result := Result + ' FROM TROPCVM LEFT JOIN TROPCVMREF ON TOF_CODEOPCVM = TOP_CODEOPCVM ' +
                      'LEFT JOIN CHOIXCOD ON CC_TYPE = "TRP" AND CC_CODE = TOP_PORTEFEUILLE ' +
                      'LEFT JOIN BANQUECP ON BQ_CODE = TOP_GENERAL ';

  {On récupère les OPCVM qui ne sont pas clôturés avant le début de la période ...}
  Result := Result + 'WHERE (TOP_DATEFIN IS NULL OR TOP_DATEFIN = "' + UsDateTime(iDate1900) +
            '" OR TOP_DATEFIN >= "' + UsDateTime(StrToDate(edOPE.Text)) + '") '  +
            {... et achetés avant la fin de la période}
            'AND TOP_DATEACHAT BETWEEN "' + UsDateTime(StrToDate(edOPE.Text)) +
                                '" AND "' + UsDateTime(StrToDate(edOPE_.Text)) + '" ';

  if (edPortefeuille.Text <> '') or (edPortefeuille_.Text <> '') then
    Result := Result + 'AND (TOP_PORTEFEUILLE BETWEEN "' + edPortefeuille.Text + '" AND "' + edPortefeuille_.Text + '") ';
  if (edOPCVM.Text <> '') or (edOPCVM_.Text <> '') then
    Result := Result + 'AND (TOP_CODEOPCVM BETWEEN "' + edOPCVM.Text + '" AND "' + edOPCVM_.Text + '") ';
  {Gestion des filtres sur BANQUECP avec éventuelle gestion des confidentialités}
  Result := Result + FiltreBanqueCp(tcp_Bancaire, '', '');

  Result := Result + ' ##UNION## ';

  Result := Result + 'SELECT "V" AS CED_TYPE, TOP_PORTEFEUILLE AS CED_PORTEFEUILLE, TOP_CODEOPCVM AS CED_CODEOPCVM, ' +
                     'TOP_NUMOPCVM AS CED_NUMTRANSAC, TVE_DATEVENTE AS CED_DATEOPERATION, TOP_NBPARTACHETE AS CED_NBACHAT, ' +
                     '(TOP_PRIXUNITAIRE / TOP_COTATIONACH) AS CED_COURSACH, TOP_DATEACHAT AS CED_DATECOMPL, ';

  Result := Result + 'TVE_COURSEUR AS CED_COURSVEN, TVE_NBVENDUE AS CED_NBVENTE, TVE_PMVALUENETTE AS CED_PMV, ' +
                     'TVE_RENDEMENTNET AS CED_RENDEMENT, TOP_BASECALCUL AS CED_BASE, ';
  Result := Result + 'TROPCVMREF.TOF_LIBELLE AS CED_LIBOPCVM, CHOIXCOD.CC_LIBELLE AS CED_LIBPORTEFEUILLE ';
  Result := Result + ' FROM TRVENTEOPCVM LEFT JOIN TROPCVMREF ON TOF_CODEOPCVM = TVE_CODEOPCVM ' +
                      'LEFT JOIN TROPCVM ON TOP_NUMOPCVM = TVE_NUMTRANSAC ' +
                      'LEFT JOIN CHOIXCOD ON CC_TYPE = "TRP" AND CC_CODE = TVE_PORTEFEUILLE ' +
                      'LEFT JOIN BANQUECP ON BQ_CODE = TVE_GENERAL ';

  {On récupère les OPCVM qui ne sont pas clôturés avant le début de la période ...}
  Result := Result + 'WHERE (TOP_DATEFIN IS NULL OR TOP_DATEFIN = "' + UsDateTime(iDate1900) +
            '" OR TOP_DATEFIN >= "' + UsDateTime(StrToDate(edOPE.Text)) + '") '  +
            {... et achetés avant la fin de la période}
            'AND TVE_DATEVENTE BETWEEN "' + UsDateTime(StrToDate(edOPE.Text)) +
                                '" AND "' + UsDateTime(StrToDate(edOPE_.Text)) + '" ';

  if (edPortefeuille.Text <> '') or (edPortefeuille_.Text <> '') then
    Result := Result + 'AND (TVE_PORTEFEUILLE BETWEEN "' + edPortefeuille.Text + '" AND "' + edPortefeuille_.Text + '") ';
  if (edOPCVM.Text <> '') or (edOPCVM_.Text <> '') then
    Result := Result + 'AND (TVE_CODEOPCVM BETWEEN "' + edOPCVM.Text + '" AND "' + edOPCVM_.Text + '") ';
  {Gestion des filtres sur BANQUECP avec éventuelle gestion des confidentialités}
  Result := Result + FiltreBanqueCp(tcp_Bancaire, '', '');

  Result := Result + ' ORDER BY CED_PORTEFEUILLE, CED_CODEOPCVM, CED_DATEOPERATION, CED_DATECOMPL';
end;

{Récupération des opérations faites avant la période de traitement mais bouclé avant la
 date de départ. Ici, on fait un LEFT JOIN et non pas une UNION, car le résultat de la
 requête sert à calculer des cumuls et non pas à afficher le détaile des opérations
{---------------------------------------------------------------------------------------}
function TOF_TRQRJALPORTEFEUIL.GetSQLRequeteCumul : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'SELECT TOP_PORTEFEUILLE, TOP_CODEOPCVM, TOP_DATEACHAT, TVE_DATEVENTE, ' +
            '(TOP_PRIXUNITAIRE / TOP_COTATIONACH) AS TOP_COURSACH, TOP_NBPARTACHETE, ';
  Result := Result + 'TROPCVMREF.TOF_LIBELLE AS CED_LIBOPCVM, CHOIXCOD.CC_LIBELLE AS CED_LIBPORTEFEUILLE, ' +
            'TVE_NBVENDUE, TOP_NUMOPCVM, TOP_MONTANTACH, TVE_COURSEUR FROM TROPCVM ';
  Result := Result + 'LEFT JOIN TRVENTEOPCVM ON TVE_NUMTRANSAC = TOP_NUMOPCVM ' +
                     'LEFT JOIN TROPCVMREF ON TOF_CODEOPCVM = TOP_CODEOPCVM ' +
                     'LEFT JOIN CHOIXCOD ON CC_TYPE = "TRP" AND CC_CODE = TOP_PORTEFEUILLE ';
  {On récupère les OPCVM qui ne sont pas clôturés avant le début de la période ...}
  Result := Result + 'WHERE (TOP_DATEFIN IS NULL OR TOP_DATEFIN = "' + UsDateTime(iDate1900) +
            '" OR TOP_DATEFIN >= "' + UsDateTime(StrToDate(edOPE.Text)) + '") '  +
            {... et achetés avant la fin de la période}
            'AND TOP_DATEACHAT < "' + UsDateTime(StrToDate(edOPE.Text)) + '" ';

  if (edPortefeuille.Text <> '') or (edPortefeuille_.Text <> '') then
    Result := Result + 'AND (TOP_PORTEFEUILLE BETWEEN "' + edPortefeuille.Text + '" AND "' + edPortefeuille_.Text + '") ';
  if (edOPCVM.Text <> '') or (edOPCVM_.Text <> '') then
    Result := Result + 'AND (TOP_CODEOPCVM BETWEEN "' + edOPCVM.Text + '" AND "' + edOPCVM_.Text + '") ';

  {Gestion des filtres sur BANQUECP avec éventuelle gestion des confidentialités}
  Result := Result + FiltreBanqueCp(tcp_Bancaire, '', '');
  Result := Result + ' ORDER BY TOP_PORTEFEUILLE, TOP_CODEOPCVM, TOP_DATEACHAT, TVE_DATEVENTE ';
end;

{Récupération des cours des OPCVM à la fin de la période
{---------------------------------------------------------------------------------------}
function TOF_TRQRJALPORTEFEUIL.GetSQLRequeteTaux : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'SELECT T1.*, TOF_DEVISE FROM TRCOURSCHANCEL T1 ' +
            'LEFT JOIN TROPCVMREF ON TTO_CODEOPCVM = TOF_CODEOPCVM ';
  Result := Result + 'WHERE T1.TTO_DATE = (SELECT MAX(T2.TTO_DATE) FROM TRCOURSCHANCEL T2 WHERE ' +
                     'T2.TTO_DATE <= "' + UsDateTime(StrToDate(edOPE_.Text)) +
                      '" AND T2.TTO_CODEOPCVM = T1.TTO_CODEOPCVM) ';

  if (edOPCVM.Text <> '') or (edOPCVM_.Text <> '') then
    Result := Result + 'AND (T1.TTO_CODEOPCVM BETWEEN "' + edOPCVM.Text + '" AND "' + edOPCVM_.Text + '") ';

  Result := Result + ' ORDER BY TTO_CODEOPCVM, TTO_DATE';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.RemplitTableTempo;
{---------------------------------------------------------------------------------------}
begin
  {On vide la table CEDTBALANCE}
  VideTableTempo;
  UniqueKey := 0;

  {Remplissage de la table CEDTBALANCE}
  TobTmp  := TOB.Create('mmmmm', nil, -1);
  TobTaux := TOB.Create('_TAUX', nil, -1);
  try
    TobTaux.LoadDetailFromSQL(GetSQLRequeteTaux);
    {Calcul des stocks et des PMValues latentes}
    CalculVolume;
    {Calcul du cumul}
    CalculCumul;

    {Ecriture dans la table temporaire}
    TobTmp.InsertDB(nil, True);
  finally
    FreeAndNil(TobTmp);
    FreeAndNil(TobTaux);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRQRJALPORTEFEUIL.GetTaux(aOPCVM : string) : Double;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  c : Double;
begin
  Result := 1;
  T := TobTaux.FindFirst(['TTO_CODEOPCVM'], [aOPCVM], True);

  if Assigned(T) then begin
    c := DoubleNotNul(T.GetDouble('H_COTATION'));
    Result := T.GetDouble('TTO_COTATION') / c;
  end
  else if V_PGI.SAV then
    PGIInfo(TraduireMemoire('Impossible de récupérer le cours de l''OPCVM : ') + aOPCVM + '.', Ecran.Caption);

end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.CalculVolume;
{---------------------------------------------------------------------------------------}
var
  TobSQL   : TOB;
  n        : Integer;
  F        : TOB;
  Enreg    : TOB;
  BaseJour : Double;
  PMVNette : Double;
  NbParts  : Double;
  CumulCou : Double;
  NbJourAn : Integer;
  OPCVM    : string;
  Portef   : string;

      {Mise à jour du cumul du montant d'achat des parts restantes.
       REMARQUE : n'est traité ici que le cumul sur la période.
                  L'antériorité sera traitée dans CalculCumul
      {-------------------------------------------------------------}
      procedure _MajCumul;
      {-------------------------------------------------------------}
      var
        T : TOB;
      begin
        T := TobTmp.FindFirst([CHP_PORTEFEUIL, CHP_OPCVM], [Portef, OPCVM], True);
        while  Assigned(T) do begin
          T.SetDouble(CHP_CUMACHAT, CumulCou);
          T.SetDouble(CHP_PARTSFIN, NbParts);
          T := TobTmp.FindNext([CHP_PORTEFEUIL, CHP_OPCVM], [Portef, OPCVM], True);
        end;
      end;

begin
  NbParts := 0;
  TobSQL  := TOB.Create('_VOLUME', nil, -1);
  try
    TobSQL.LoadDetailFromSQL(GetSQLRequetePrinc);
    //TobDebug(TobSQL);
    for n := 0 to TobSQL.Detail.Count - 1 do begin
      F := TobSQL.Detail[n];
      if (OPCVM <> F.GetString('CED_CODEOPCVM')) or (Portef <> F.GetString('CED_PORTEFEUILLE')) then begin
        {On change de titre, on met à jour le cumul du montant d'achat des parts restantes}
        if (OPCVM <> '') and (Portef <> '') then _MajCumul;

        {Initialisation du cumul du montant d'achat des parts restantes sur la période}
        if F.GetString('CED_TYPE') = 'A' then CumulCou := F.GetDouble('CED_NBACHAT') * F.GetDouble('CED_COURSACH')
                                         else CumulCou := 0 - F.GetDouble('CED_NBVENTE') * F.GetDouble('CED_COURSACH');
        {Initialisation du cumul des parts restantes sur la période}
        if F.GetString('CED_TYPE') = 'A' then NbParts := F.GetDouble('CED_NBACHAT')
                                         else NbParts := 0 - F.GetDouble('CED_NBVENTE');

        OPCVM  := F.GetString('CED_CODEOPCVM');
        Portef := F.GetString('CED_PORTEFEUILLE');
      end
      else begin
        {Cumul du montant d'achat des parts restantes sur la période}
        if F.GetString('CED_TYPE') = 'A' then CumulCou := CumulCou + F.GetDouble('CED_NBACHAT') * F.GetDouble('CED_COURSACH')
                                         else CumulCou := CumulCou - F.GetDouble('CED_NBVENTE') * F.GetDouble('CED_COURSACH');
        {Cumul des parts restantes sur la période}
        if F.GetString('CED_TYPE') = 'A' then NbParts := NbParts + F.GetDouble('CED_NBACHAT')
                                         else NbParts := NbParts - F.GetDouble('CED_NBVENTE');
      end;

      Enreg := Tob.Create('CEDTBALANCE', TobTmp, -1);
      Enreg.SetString('CED_USER'    , V_PGI.User);
      Enreg.SetString('CED_COMPTE'  , IntToStr(n));
      Enreg.SetString(CHP_OPCVM     , F.GetString('CED_CODEOPCVM'));
      Enreg.SetString('CED_LIBELLE2', F.GetString('CED_LIBOPCVM'));
      Enreg.SetString('CED_LIBELLE' , F.GetString('CED_LIBPORTEFEUILLE'));
      Enreg.SetString(CHP_PORTEFEUIL, F.GetString('CED_PORTEFEUILLE'));
      Enreg.SetString('CED_RUPTURE' , F.GetString('CED_DATEOPERATION'));
      Enreg.SetDouble(CHP_PMVNETTE  , F.GetDouble('CED_PMV'));
      Enreg.SetDouble(CHP_MNTVENTE  , F.GetDouble('CED_NBVENTE') * F.GetDouble('CED_COURSVEN'));
      Enreg.SetDouble(CHP_COURSFIN  , GetTaux(F.GetString('CED_CODEOPCVM')));
      Enreg.SetDouble(CHP_PARTSCOU  , NbParts);

      if F.GetString('CED_TYPE') = 'A' then begin
        Enreg.SetDouble(CHP_NBJOURSDT , 0);
        Enreg.SetDouble(CHP_RENDJOUR  , 0);
        Enreg.SetDouble(CHP_MNTACHAT  , F.GetDouble('CED_NBACHAT') * F.GetDouble('CED_COURSACH'));
        Enreg.SetDouble(CHP_BASEANNUEL, 365);
        Enreg.SetDouble(CHP_COURSCOU  , F.GetDouble('CED_COURSACH'));
      end else begin
        Enreg.SetDouble(CHP_MNTACHAT  , F.GetDouble('CED_NBVENTE') * F.GetDouble('CED_COURSVEN'));
        Enreg.SetDouble(CHP_NBJOURSDT , Round(F.GetDateTime('CED_DATEOPERATION') - F.GetDateTime('CED_DATECOMPL')));
        Enreg.SetDouble(CHP_COURSCOU  , F.GetDouble('CED_COURSVEN'));

        {Calcul du rendement journalier}
        NbJourAn := CalcNbJourParAnBase(F.GetDateTime('CED_DATEOPERATION'), F.GetInteger('CED_BASE'));
        if NbJourAn = 0 then NbJourAn := 1;
        BaseJour := (F.GetDateTime('CED_DATEOPERATION') - F.GetDateTime('CED_DATECOMPL')) / NbJourAn;
        PMVNette := F.GetDouble('CED_RENDEMENT') * BaseJour;
        Enreg.SetDouble(CHP_RENDJOUR  , PMVNette);
        Enreg.SetDouble(CHP_BASEANNUEL, NbJourAn);
      end;
      UniqueKey := n + 1;
    end; {For n := 0}

    {On met à jour le cumul du montant d'achat des parts restantes du dernier titre traité}
    if (OPCVM <> '') and (Portef <> '') then _MajCumul;

  finally
    FreeAndNil(TobSQL);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALPORTEFEUIL.CalculCumul;
{---------------------------------------------------------------------------------------}
var
  TobSQL  : TOB;
  n       : Integer;
  F       : TOB;
  OPCVM   : string;
  Transac : string;
  Portef  : string;
  NbParts : Double;
  cuAchat : Double;

    {-------------------------------------------------------------------}
    procedure _MajCumuls;
    {-------------------------------------------------------------------}
    var
      T : TOB;
    begin
      T := TobTmp.FindFirst([CHP_PORTEFEUIL, CHP_OPCVM], [Portef, OPCVM], True);

      {Si l'opcvm n'a pas été mouvementé sur la période, on crée un enregistrement
       pour afficher l'encours}
      if not Assigned(T) then begin
        T := Tob.Create('CEDTBALANCE', TobTmp, -1);
        T.SetString('CED_USER'    , V_PGI.User);
        T.SetString('CED_COMPTE'  , IntToStr(UniqueKey));
        T.SetString(CHP_OPCVM     , OPCVM);
        T.SetString('CED_LIBELLE2', F.GetString('CED_LIBOPCVM'));
        T.SetString('CED_LIBELLE' , F.GetString('CED_LIBPORTEFEUILLE'));
        T.SetString(CHP_PORTEFEUIL, Portef);
        T.SetString('CED_RUPTURE' , edOPE_.Text);
        T.SetDouble(CHP_PMVNETTE  , 0);
        T.SetDouble(CHP_MNTVENTE  , 0);
        T.SetDouble(CHP_NBJOURSDT , 0);
        T.SetDouble(CHP_RENDJOUR  , 0);
        T.SetDouble(CHP_MNTACHAT  , 0);
        T.SetDouble(CHP_BASEANNUEL, 365);
        T.SetDouble(CHP_PARTSFIN  , NbParts);
        T.SetDouble(CHP_COURSFIN  , GetTaux(F.GetString('TOP_CODEOPCVM')));
        T.SetDouble(CHP_COURSCOU  , GetTaux(F.GetString('TOP_CODEOPCVM')));
        T.SetDouble(CHP_PARTSCOU  , NbParts);
        T.SetDouble(CHP_CUMACHAT  , cuAchat);
        Inc(UniqueKey);
      end

      {Sinon mise à jour des lignes de l'opcvm et du portefeuille}
      else
        while Assigned(T) do begin
          T.SetDouble(CHP_PARTSFIN, T.GetDouble(CHP_PARTSFIN) + NbParts);
          T.SetDouble(CHP_CUMACHAT, T.GetDouble(CHP_CUMACHAT) + cuAchat);
          T := TobTmp.FindNext([CHP_PORTEFEUIL, CHP_OPCVM], [Portef, OPCVM], True);
        end;
    end;

begin
  NbParts := 0;
  TobSQL  := TOB.Create('_VOLUME', nil, -1);
  try
    TobSQL.LoadDetailFromSQL(GetSQLRequeteCumul);
   // TobDebug(TobSQL);
    for n := 0 to TobSQL.Detail.Count - 1 do begin
      F := TobSQL.Detail[n];
      {Nouvel OPCVM ou nouveau portefeuille}
      if (OPCVM <> F.GetString('TOP_CODEOPCVM')) or (Portef <> F.GetString('TOP_PORTEFEUILLE')) then begin
        {Ecritures des cumuls}
        if (OPCVM <> '') and (Portef <> '') then _MajCumuls;

        NbParts := F.GetDouble('TOP_NBPARTACHETE');
        cuAchat := F.GetDouble('TOP_MONTANTACH');
        {Si la date de vente est antérieure à la période, on déduit la vente.
         On ne fait pas le test sur la date d'achat, car c'est fait dans la requête}
        if F.GetDateTime('TVE_DATEVENTE') < StrToDate(edOpe.Text) then begin
          cuAchat := cuAchat - (F.GetDouble('TVE_NBVENDUE') * F.GetDouble('TOP_COURSACH'));
          NbParts := NbParts - F.GetDouble('TVE_NBVENDUE');
        end;

        OPCVM   := F.GetString('TOP_CODEOPCVM');
        Portef  := F.GetString('TOP_PORTEFEUILLE');
      end

      else begin
        if (Transac <> F.GetString('TOP_NUMOPCVM')) then begin
          {Il s'agit d'une ligne d'achat avec une éventuelle vente}
          NbParts := NbParts + F.GetDouble('TOP_NBPARTACHETE');
          cuAchat := cuAchat + F.GetDouble('TOP_MONTANTACH');
          {Si la date de vente est antérieure à la période, on déduit la vente.}
          if F.GetDateTime('TVE_DATEVENTE') < StrToDate(edOpe.Text) then begin
            cuAchat := cuAchat - (F.GetDouble('TVE_NBVENDUE') * F.GetDouble('TOP_COURSACH'));
            NbParts := NbParts - F.GetDouble('TVE_NBVENDUE');
          end;
        end
        {Il s'agit d'un vente}
        else begin
          {Si la date de vente est antérieure à la période, on déduit la vente.}
          if F.GetDateTime('TVE_DATEVENTE') < StrToDate(edOpe.Text) then begin
            cuAchat := cuAchat - (F.GetDouble('TVE_NBVENDUE') * F.GetDouble('TOP_COURSACH'));
            NbParts := NbParts - F.GetDouble('TVE_NBVENDUE');
          end;
        end;
      end;
      Transac := F.GetString('TOP_NUMOPCVM');
    end; {For n := 0}

    {Ecritures des cumuls}
    if (OPCVM <> '') and (Portef <> '') then _MajCumuls;

  finally
    FreeAndNil(TobSQL);
  end;
end;


initialization
  RegisterClasses([TOF_TRQRJALPORTEFEUIL]);

end.

