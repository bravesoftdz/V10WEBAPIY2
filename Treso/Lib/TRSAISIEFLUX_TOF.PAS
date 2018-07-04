{ Unité : Source TOF de la FICHE : TRSAISIEFLUX ()
--------------------------------------------------------------------------------------
    Version  |  Date  | Qui | Commentaires
--------------------------------------------------------------------------------------
 0.91         16/09/03  JP   Création de l'unité
1.05.001.002  11/03/04  JP   Mise en place d'initialisation pour les écritures
                             créées depuis le rapprochement cf Rappro
6.30.001.002  08/03/05  JP   FQ 10217 : gestion de la date d'opération pour toutes les fiches de suivi
6.50.001.001  24/05/05  JP   Gestion des formats des montants et nouvelle gestion des images
6.50.001.001  24/05/05  JP   FQ 10223 : Nouvelle gestion des erreurs
6.50.001.013  08/08/05  JP   FQ 10088 : En accès depuis le rappro, on filtre les flux en fonction du CIB
6.50.001.013  08/08/05  JP   FQ 10089 : Gestion du code ET du libéllé des flux / rubriques
6.60.001.001  28/12/05  JP   FQ 10321 : Les écritures de simulations ne passaient plus faute de contrepartie
7.09.001.001  07/08/06  JP   Gestion du multi sociétés
7.09.001.001  12/10/06  JP   Gestion des ICC
8.01.001.010  04/04/07  JP   Nouvelle fonction de création des flux TrSaisieFluxTob
-------------------------------------------------------------------------------------}
unit TRSAISIEFLUX_TOF ;

interface

uses StdCtrls, Controls, Classes, Vierge,
    {$IFNDEF EAGLCLIENT}
    FE_Main, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    {$ELSE}
    MaineAGL,
    {$ENDIF}
    Forms, SysUtils, HCtrls, HEnt1, HMsgBox, UTOF, Commun, Constantes,
    Ent1, HTB97, Messages, Windows, UTob, UProcGen;

type
  TOF_TRSAISIEFLUX = class (TOF)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  private
    BudgetMensuel   : Boolean;
    BudgetLibre     : Boolean;
    Rappro          : Boolean;
    IccOk           : Boolean; {12/10/06}
    NomBase         : string;

    Nature          : string;
    edGeneral       : THEdit;
    cbFlux          : THValComboBox;
    bValider        : TToolbarButton97;
    ValidationClick : TNotifyEvent;
    DateReference   : TDateTime;
    DateReferenceF  : TDateTime;
    FCLoseQuery     : TCloseQueryEvent;
    FKeyDown        : TKeyEvent;
    tResultCalcIcc  : TOB;

    function  VerifSensFlux  : Boolean;
    procedure InitTobCalculIcc;

    procedure GeneralOnChange(Sender : TObject);
    procedure ValiderOnClick (Sender : TObject);
    procedure NatureOnChange (Sender : TObject);
    procedure DateExit       (Sender : TObject);
    procedure CbFluxOnChange (Sender : TObject);
    procedure CbFlux1OnChange(Sender : TObject);
    procedure CbFluxOnExit   (Sender : TObject);
    procedure BCalculOnClick (Sender : TObject);
    procedure EcranCloseQuery(Sender : TObject; var CanClose : Boolean);
    procedure EcranKeyDown   (Sender: TObject; var Key: Word; Shift: TShiftState);
  end;

function TRLanceFiche_SaisieFlux(Dom, Fiche, Range, Lequel, Arguments : string) : string;
function TrSaisieFluxTob(Arg : string) : TOB;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UProcCommission, UProcEcriture, UProcSolde, AglInit, UtilPgi, TRCALCULICC_TOF;

var
  TobEcr : TOB;

{---------------------------------------------------------------------------------------}
function TRLanceFiche_SaisieFlux(Dom, Fiche, Range, Lequel, Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  TobEcr := nil;
  Result := AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
function TrSaisieFluxTob(Arg : string) : TOB;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := nil;
  T := TOB.Create('_ECR', nil, -1);
  TobEcr := TOB.Create('TRECRITURE', T, -1);
  if AGLLanceFiche('TR', 'TRSAISIEFLUX', '', '', Arg) <> '' then Result := T
                                                            else FreeAndNil(T);
end;

{Création de l'écriture : 08/08/06 Remplacement du TrEcritures par une Tob
{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  Num      : string;
  TypeFlux : string;
  Sens     : Shortint;
  CodeUnq  : string;
  ObjCom   : TObjCommissionTob;
  ObjDev   : TObjDetailDevise;
  Q        : TQuery;
  g, c     : string;
  MajDiff  : Boolean;
begin
  inherited;
  LastError := 0;
  if IccOk then begin
    {On va stocker dans la date de validation la date de début de période de calcul des iCC :
     1/ On n'a pas touché aux données issues de TRCALCULICC_TOF => on reprend "DEBUT"
     2/ Sinon, dans TRCALCULICC_TOF, on stocke la date de début d'exercice dans TE_DATEVALID
     3/ Si TE_DATEVALID = iDate1900, c'est qu'on n'est pas passé par TRCALCULICC_TOF ... }
    if (tResultCalcIcc.GetDouble('TE_MONTANT') <> Valeur(GetControlText('MONTANT'))) or
       (tResultCalcIcc.GetString('TE_DATECOMPTABLE') <> GetControlText('DATECOMPTABLE')) then begin
      if tResultCalcIcc.GetDouble('TE_DATEVALID') = iDate1900 then begin
        Q := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE "' + UsDateTime(StrToDate(GetControlText('DATECOMPTABLE'))) +
                     '" BETWEEN EX_DATEDEBUT AND EX_DATEFIN', True, -1, '', False, GetInfosSocFromBQ(edGeneral.Text, False).RT);
        if not Q.EOF then
          {On mémorise la date de début d'exercice pour la création du flux}
          tResultCalcIcc.SetDateTime('TE_DATEVALID', Q.FindField('EX_DATEDEBUT').AsDateTime)
        else
          tResultCalcIcc.SetDouble('TE_DATEVALID', StrToDate(GetControlText('DATECOMPTABLE')));
        Ferme(Q);
      end;
    end
    else
      tResultCalcIcc.SetDouble('TE_DATEVALID', tResultCalcIcc.GetDouble('DEBUT'));

    tResultCalcIcc.AddChampSupValeur('TE_DATECOMPTABLE', GetControlText('DATECOMPTABLE'));
    tResultCalcIcc.AddChampSupValeur('TE_CODEFLUX'     , GetControlText('CODEFLUX'));
    tResultCalcIcc.AddChampSupValeur('TE_CODEFLUX1'    , GetControlText('CODEFLUX1'));
    tResultCalcIcc.AddChampSupValeur('TE_GENERAL'      , edGeneral.Text);
    tResultCalcIcc.AddChampSupValeur('TE_LIBELLE'      , GetControlText('LIBELLE'));
    tResultCalcIcc.AddChampSupValeur('TE_MONTANT', Abs(Valeur(GetControlText('MONTANT'))));

    if not VerifSensFlux then begin
      LastError := 1;
      LastErrorMsg := TraduireMemoire('Les deux codes flux doivent être de sens opposé.');
      SetFocusControl('CODEFLUX1');
    end else if (GetControlText('CODEFLUX1') = '') or (GetControlText('CODEFLUX') = '') then begin
      LastError := 1;
      LastErrorMsg := TraduireMemoire('Les deux codes flux doivent être renseignés et être de sens opposé.');
      SetFocusControl('CODEFLUX');
    end
    else if not CreeTrEcritureIcc(tResultCalcIcc) then
      //LastError := 1;
      //LastErrorMsg := TraduireMemoire('Les écritures d''ICC ne peuvent pas être créées.');
    AfficheMessageErreur(Ecran.Caption, 'L''écriture ou sa commission n''a pu être créée');

    Exit;
  end;

  InitGestionErreur;
  {27/05/05 : on précise que l'on travaille sur une éciture de trésorerie}
  CategorieCurrent := CatErr_TRE;

  MajDiff := False;
  {La mise à jour est différée si l'appel s'est fait depuis TrSaisieFluxTob}
  if Assigned(TobEcr) then MajDiff := True
                      else TobEcr  := TOB.Create('TRECRITURE', nil, -1);
  try
    TobEcr.SetString('TE_GENERAL', edGeneral.Text);

    if IsTresoMultiSoc then begin
      Q := OpenSQL('SELECT BQ_CODE, BQ_NODOSSIER, DOS_SOCIETE FROM BANQUECP ' +
                   'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                   'WHERE BQ_CODE = "' + TobEcr.GetString('TE_GENERAL') + '"', True);
      if not Q.EOF then begin
        TobEcr.SetString('TE_SOCIETE', Q.FindField('DOS_SOCIETE').AsString);
        TobEcr.SetString('TE_NODOSSIER', Q.FindField('BQ_NODOSSIER').AsString);
      end
      else
        SetErreurTreso(NatErr_Gen);
      Ferme(Q);
    end
    else begin
      TobEcr.SetString('TE_SOCIETE', V_PGI.CodeSociete);
      TobEcr.SetString('TE_NODOSSIER', V_PGI.NoDossier);
    end;

    {Initialisation du record}
    InitNlleEcritureTob(TobEcr, edGeneral.Text, TobEcr.GetString('TE_NODOSSIER'));

    TobEcr.SetString('TE_CODEFLUX', GetControlText('CODEFLUX'));
    TobEcr.SetString('TE_NATURE',   GetControlText('NATURE'));

    {S'il s'agit d'une écriture prévisionnelle, donc gérée avec un code flux ...}
    if TobEcr.GetString('TE_NATURE') = na_Prevision then begin
      {Récupération du compte de contrepartie et du code cib dans la table FLUXTRESO}
      GetCibSensGeneral(g, TypeFlux, c, TobEcr.GetString('TE_CODEFLUX'));
      TobEcr.SetString('TE_CONTREPARTIETR', g);
      TobEcr.SetString('TE_CODECIB', c);

      {Maj du sens}
      if GetSensFromTypeFlux(TypeFlux) = 'D' then Sens := -1
                                             else Sens := 1;
    end

    {... sinon, il s'agit d'une écriture de simulation (S) gérée avec un code rubrique}
    else begin
      {On récupère le sens à partir des rubriques}
      if GetSensFromRubrique(TobEcr.GetString('TE_CODEFLUX')) = 'D' then Sens := -1
                                                                    else Sens := 1;
      {... En attendant la gestion des cib !!}
      TobEcr.SetString('TE_CODECIB', CODECIBSIMUL);
      {28/12/05 : FQ 10321 créée sur le sujet.
       21/12/05 : Maintenant avec la nouvel gestion des erreurs, la contrepartie est obligatoire,
                  en attendant quelque chose de mieux, car il n'est pas facile de trouver la
                  contrepartie dans la fiche des rubriques}
      TobEcr.SetString('TE_CONTREPARTIETR', 'SIMUL');
    end;

    {Récupération du RIB}
    GetRibTob(TobEcr);

    {27/05/05 : nouvelle gestion des erreurs}
    if TobEcr.GetString('TE_CODEBANQUE') = '' then SetErreurTreso(NatErr_Bqe);

    {Génération du numéro de transaction}
    Num := GetNum(CODEMODULE, TobEcr.GetString('TE_SOCIETE'), CODESAISIE);
    TobEcr.SetString('TE_NUMTRANSAC', CODEMODULE + TobEcr.GetString('TE_SOCIETE') + TRANSACSAISIE + Num);

    TobEcr.SetString('TE_LIBELLE', GetControlText('LIBELLE'));
    TobEcr.SetString('TE_USERCREATION', V_PGI.User);
    TobEcr.SetDateTime('TE_DATECREATION', V_PGI.DateEntree);
    TobEcr.SetDateTime('TE_DATECOMPTABLE', StrToDate(GetControlText('DATECOMPTABLE')));
    TobEcr.SetString('TE_REFINTERNE', 'Saisie le ' + GetControlText('DATECOMPTABLE') + ' par ' + V_PGI.User);
    TobEcr.SetString('TE_EXERCICE', VG_ObjetExo.GetExoNodos(StrToDate(GetControlText('DATECOMPTABLE')),
                                                            TobEcr.GetString('TE_NODOSSIER')));

    {Conversion du montant en Euros}
    TobEcr.SetString('TE_DEVISE', GetControlText('DEV'));

    {24/05/05 : Gestion du format des montants}
    ObjDev := TObjDetailDevise.Create;
    try
      TheData := ObjDev;

      TobEcr.SetDouble('TE_COTATION', RetPariteEuro(TobEcr.GetString('TE_DEVISE'), TobEcr.GetDateTime('TE_DATECREATION')));
      TobEcr.SetDouble('TE_MONTANTDEV', Arrondi(Valeur(GetControlText('MONTANT')) * Sens, ObjDev.NbDecimal));
      {RetPariteEuro renvoie le taux (1 Dev = x.xx €) => on transforme en cotation (1€ = x.xx Dev)}
      if TobEcr.GetDouble('TE_COTATION') = 0 then
        TobEcr.SetDouble('TE_COTATION', 1)
      else
        TobEcr.SetDouble('TE_COTATION', Arrondi(1 / TobEcr.GetDouble('TE_COTATION'), NBDECIMALTAUX));
      TobEcr.SetDouble('TE_MONTANT', Arrondi(TobEcr.GetDouble('TE_MONTANTDEV') / TobEcr.GetDouble('TE_COTATION'), V_PGI.OkDecV));
    finally
      if Assigned(ObjDev) then FreeAndNil(ObjDev);
      TheData := nil;
    end;

    {Si l'on bien récupérer le code CIB et le flux ..}
    if (TobEcr.GetString('TE_CODECIB') > '') and (TobEcr.GetString('TE_CODEFLUX') > '') then begin
      {... Calcul de la date de valeur}
      if Rappro then
        {On récupère la date fournie en paramètre du OnArgument}
        TobEcr.SetDateTime('TE_DATEVALEUR', DateReference)
      else if (TobEcr.GetString('TE_NATURE') = na_Simulation) then
        {Dans le cas d'une écriture de simulation, les dates de valeur et comptable sont équivalentes}
        TobEcr.SetDateTime('TE_DATEVALEUR', TobEcr.GetDateTime('TE_DATECOMPTABLE'))
      else
        {Sinon, on la calcul à partir du code CIB}
        TobEcr.SetDateTime('TE_DATEVALEUR', CalcDateValeur(TobEcr.GetString('TE_CODECIB'),
                                                           TobEcr.GetString('TE_GENERAL'),
                                                           TobEcr.GetDateTime('TE_DATECOMPTABLE')));
      {20/11/06 : Avertissement sur les dates de réinitialisation}
      if TestDateEtMillesime(TobEcr.GetDateTime('TE_DATEVALEUR'), Date) or
         TestDateEtMillesime(TobEcr.GetDateTime('TE_DATECOMPTABLE'), Date) then begin
        if HShowmessage('0;' + Ecran.Caption + ';ATTENTION !'#13'La nouveau flux n''appartient pas au millésime courant.'#13 +
           'Le montant saisi peut avoir un impact sur les montants de réinitialisation.'#13#13 +
           'Souhaitez-vous poursuivre ?;W;YN;N;N', '', '') = mrNo then begin
          PgiInfo(TraduireMemoire('Le Flux n''a pas été créé.'));
          Exit;
        end;
      end;
      {... Génération des clés}
      CodeUnq := GetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE);
      TobEcr.SetString('TE_CLEVALEUR',    RetourneCleEcriture(TobEcr.GetDateTime('TE_DATEVALEUR'), StrToInt(CodeUnq)));
      TobEcr.SetString('TE_CLEOPERATION', RetourneCleEcriture(TobEcr.GetDateTime('TE_DATECOMPTABLE'), StrToInt(CodeUnq)));
      SetNum(CODEUNIQUE, CODEUNIQUE, CODEUNIQUE, CodeUnq);

      ObjCom := TObjCommissionTob.Create(TobEcr.GetString('TE_GENERAL'), TobEcr.GetString('TE_CODEFLUX'), TobEcr.GetDateTime('TE_DATECOMPTABLE'));
      try
        if not MajDiff then BeginTrans;
        try
          {16/07/04 : gestion des commissions avant l'écriture de l'opération financière car le numéro
                      de transaction peut être modifié dans ObjCom.GenererCommissions}
          ObjCom.GenererCommissions(TobEcr);
          ObjCom.TobEcriture.ChangeParent(TobEcr, -1);
          {... écriture du flux}
          CategorieCurrent := CatErr_TRE;
          if not TesteEcrTresorerie(TobEcr) then begin
            {17/03/05 : FQ 10223 : Refonte de la gestion des erreurs}
            AfficheMessageErreur(Ecran.Caption, 'L''écriture ou sa commission n''a pu être créée');
            if not MajDiff then RollBackDiscret;
            Exit;
          end
          else
            if not MajDiff then TobEcr.InsertDb(nil);

          {... recalcul des soldes}
          if not MajDiff then
            RecalculSolde(TobEcr.GetString('TE_GENERAL'), TobEcr.GetString('TE_DATECOMPTABLE'), 0, True);
          {... mise à jour du compteur}
          SetNum(CODEMODULE, TobEcr.GetString('TE_SOCIETE'), CODESAISIE, Num);
          if not MajDiff then CommitTrans;
        except
          on E : Exception do begin
            if not MajDiff then RollBack;
            HShowMessage('1;' + Ecran.Caption + ';Impossible de créer l''écriture.;E;O;O;O;', '', '');
          end;
        end;
      finally
        if Assigned(ObjCom) then FreeAndNil(ObjCom);
      end;

      {Dans le cas du Rapprochement, on récupère la clé de l'écriture pour éventuellement
       réaliser l'écriture et créer l'écriture comptable}
      if Rappro then TFVierge(Ecran).Retour := TobEcr.GetString('TE_CLEOPERATION')
                else TFVierge(Ecran).Retour := 'N';
    end;
  finally
    if Assigned(TobEcr) and not MajDiff then FreeAndNil(TobEcr);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  CompteBqe : string;
  ch        : string;
begin
  inherited;
  edGeneral := THEdit(GetControl('GENERAL'));
  edGeneral.OnChange := GeneralOnChange;

  cbFlux := THValComboBox(GetControl('CODEFLUX'));
  THValComboBox(GetControl('NATURE')).OnChange := NatureOnChange;

  {08/08/05 : FQ 10089 Gestion des libellés ET des codes Flux}
  SetControlCaption('LBCODEFLUX', '');
  cbFlux.OnChange := CbFluxOnChange;

  {12/10/06 : Gestion des ICC si on est en multi sociétés}
  IccOk := False;
  if IsTresoMultiSoc then begin
    if Pos('ICCMUL', S) > 0 then IccOk := True;
  end;
  if IccOk then Ecran.HelpContext := 5000151
           else Ecran.HelpContext := 150;

  {Récupération du compte général éventuellement passé en paramètre ...}
  CompteBqe := ReadTokenSt(S);
  {... Dans ce cas, il n'est pas possible de choisir le compte général ...}
  SetControlEnabled('GENERAL', CompteBqe = '');
  {... Sauf si le compte n'est pas trouvable dans la tablette}
  if CompteBqe <> '' then begin
    if not ExisteSQL('SELECT BQ_CODE FROM BANQUECP WHERE BQ_CODE = "' + CompteBqe + '"') then
      SetControlEnabled('GENERAL', True)
    else
      edGeneral.Text := CompteBqe;
  end
  else
    {Si on est en saisie libre, on considère qu'il s'agit d'écritures de simulation}
    Nature := na_Simulation;

  {Maj du drapeau et du libellé}
  GeneralOnChange(edGeneral);

  ch := ReadTokenSt(S);

  {Si on est dans le cas d'une gestion de budget mensuel}
  BudgetMensuel := ch = 'MENS';
  BudgetLibre   := ch = 'LIBRE';
  {Si l'on vient de l'écran de rapprochement bancaire}
  Rappro := ch = 'RAPPRO';

  {On reprend en main la gestion du bouton valider, pour pouvoir tester la validité de la saisie}
  bValider := TToolbarButton97(GetControl('BVALIDER'));
  ValidationClick := bValider.OnClick;
  bValider.OnClick := ValiderOnClick;
  bValider.ModalResult := mrNone;

  {Si en budget mensuel}
  if BudgetMensuel then begin
    DateReference := StrToDate(ReadTokenSt(S));
    SetControlText('DATECOMPTABLE', DateToStr(DateReference));
    THEdit(GetControl('DATECOMPTABLE')).OnExit := DateExit;
  end
  {Si budget par période}
  else if BudgetLibre then begin
    DateReference  := StrToDate(ReadTokenSt(S));
    DateReferenceF := StrToDate(ReadTokenSt(S));
    SetControlText('DATECOMPTABLE', DateToStr(DateReference));
    THEdit(GetControl('DATECOMPTABLE')).OnExit := DateExit;
  end
  else if Rappro then begin
    SetControlText('DATECOMPTABLE', ReadTokenSt(S));
    SetControlText('LIBELLE',       ReadTokenSt(S));
    SetControlText('MONTANT',       ReadTokenSt(S));
    {On récupère ici la date de valeur qui sera forcée lors de la création de l'écriture}
    DateReference := StrToDate(ReadTokenSt(S));
    THValComboBox(GetControl('NATURE')).Value := na_Prevision;
    NatureOnChange(GetControl('NATURE'));
    SetControlEnabled('NATURE', False);
    {08/08/05 : FQ 10088 : récupération du code cib pour filtrer les flux}
    Ch := ReadTokenSt(S);
    if Ch <> '' then
      cbFlux.Plus := cbFlux.Plus + ' AND TFT_CODECIB = "' + Ch + '"';
  end

  {12/10/06 : Pour les Icc, il y a deux appels :
              1/ Depuis le calcul des Icc (uTOFIccGeneraux), là on ne peut choisir que le code Flux
              2/ Depuis le Mul (TRMulIcc_TOF), et là, seule la nature est bloquée}
  else if IccOk then begin
    THValComboBox(GetControl('NATURE')).Value := na_Prevision;
    NatureOnChange(GetControl('NATURE'));
    SetControlEnabled('NATURE', False);
    SetControlVisible('CODEFLUX1', True);
    SetControlVisible('LBCODEFLUX1', True);
    SetControlVisible('LBFLUX1', True);

    cbFlux.OnExit := CbFluxOnExit;
    cbFlux.Plus := 'TFT_CLASSEFLUX = "' + cla_ICC + '"';
    THValComboBox(GetControl('CODEFLUX1')).Plus := 'TFT_CLASSEFLUX = "' + cla_ICC + '"';
    THValComboBox(GetControl('CODEFLUX1')).OnChange := CbFlux1OnChange;
    THValComboBox(GetControl('CODEFLUX1')).OnExit := CbFluxOnExit;

    edGeneral.DataType := tcp_Tous;
    edGeneral.Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');

    FCLoseQuery := Ecran.OnCloseQuery;
    Ecran.OnCloseQuery := EcranCloseQuery;

    FKeyDown := Ecran.OnKeyDown;
    Ecran.OnKeyDown := EcranKeyDown;
    {Tob contenant le résultat du calcul des ICC}
    InitTobCalculIcc;

    SetControlText('DATECOMPTABLE', DateToStr(Date));
    SetControlVisible('BCALCUL'   , True);
    TToolbarButton97(GetControl('BCALCUL')).OnClick := BCalculOnClick;
  end

  else begin
    {08/03/05 : FQ 10217 : Pour le détail de suivi, le suivi par solde et le suivi du trésorerier,
                on se positionne sur la date d'opération, si on est en date d'opération}
    Ch := ReadTokenSt(S);
    if Ch = '' then Ch := DateToStr(V_PGI.DateEntree);
    SetControlText('DATECOMPTABLE', Ch);
  end;

  if not IccOk then
    edGeneral.Plus := FiltreBanqueCp(edGeneral.DataType, tcb_Bancaire, '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if IccOk and Assigned(tResultCalcIcc) then FreeAndNil(tResultCalcIcc);
  inherited;
end;

{Mise à jour du code devise et du drapeau du compte sélectionné
{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.GeneralOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {24/05/05 : Centralisation de la gestion de l'affichage des devises dans une fonction}
  MajAffichageDevise(GetControl('IDEV'), GetControl('DEV'), edGeneral.Text, sd_Compte);

  {12/10/06 : Récupération du nom de la base pour les requêtes sur la table exercice pour les ICC}
  if IccOk and (edGeneral.Text <> '') then
    Nombase := GetInfosSocFromBQ(edGeneral.Text, False).RT;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.ValiderOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Err     : Byte;
  MsgErr  : string;
begin
  {Tests de validation de la saisie}
  Err := 0;

       if GetControlText('CODEFLUX'     ) = '' then Err := 1
  else if GetControlText('DATECOMPTABLE') = '' then Err := 2
  else if GetControlText('LIBELLE'      ) = '' then Err := 3
  else if GetControlText('MONTANT'      ) = '' then Err := 4
  else if edGeneral.Text                  = '' then Err := 5;

  case Err of
    1 : MsgErr := '1;Saisie de flux financier; Veuillez choisir un flux !;W;O;O;O;';
    2 : MsgErr := '2;Saisie de flux financier; Veuillez saisir une date d''opération !;W;O;O;O;';
    3 : MsgErr := '3;Saisie de flux financier; Veuillez saisir un libellé !;W;O;O;O;';
    4 : MsgErr := '4;Saisie de flux financier; Veuillez saisir un montant !;W;O;O;O;';
    5 : MsgErr := '5;Saisie de flux financier; Veuillez choisir un compte !;W;O;O;O;';
  end;

  if Err > 0 then begin
    HShowMessage(MsgErr, '', '');
    Exit;
  end
  else begin
    bValider.ModalResult := mrOk;
    ValidationClick(Sender);
    SendMessage(Ecran.Handle, WM_CLOSE, 0, 0);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.NatureOnChange (Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Pour les écritures prévisionnelles, on travaille sur les flux treso,
   pour les écritures de simulation, on travaille sur les rubriques}
  if GetControlText('NATURE') = na_Prevision then begin
    {23/01/06 : en eAGL, lorsque l'on change le DataType, il réexécute la requête immédiatement
                avec la mauvaise condition  Plus}
    cbFlux.Plus := '';
    cbFlux.DataType := 'TRCODEFLUX';
    cbFlux.Plus := 'TFT_CLASSEFLUX = "' + cla_Previ + '"';
    cbFlux.DataTypeParametrable := True;
  end else begin
    {23/01/06 : en eAGL, lorsque l'on change le DataType, il réexécute la requête immédiatement
                avec la mauvaise condition  Plus}
    cbFlux.Plus := '';
    cbFlux.DataType := 'TZRUBRIQUE';
    cbFlux.Plus := 'RB_NATRUB = "TRE" AND RB_CLASSERUB = "TRE"';
    cbFlux.DataTypeParametrable := False;
  end;

  cbFlux.Refresh;
  cbFlux.ItemIndex := 0;

  SetControlCaption('LBCODEFLUX', ''); {08/08/05 : FQ 10089}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.DateExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  d, m, f : TDateTime;
begin
  m := StrToDate(GetControlText('DATECOMPTABLE'));
  if BudgetMensuel then begin
    d := DebutDeMois(DateReference);
    f := FinDeMois(DateReference);
  end
  else begin
    d := DateReference;
    f := DateReferenceF;
  end;

  {Dans le cas de budget on s'assure que la date correspond à la colonne d'où l'on vient}
  if not Between(m, d, f) then begin
    HShowMessage('0;Saisie de flux;La date doit être comprise entre le ' + DateToStr(d) + ' et le '  +
                    DateToStr(f) + '.;Q;O;O;O;', '', '');
    SetFocusControl('DATECOMPTABLE');
  end;
end;

{08/08/05 : FQ 10089 Gestion des libellés ET des codes Flux
{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.CbFluxOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlCaption('LBCODEFLUX', cbFlux.Value);
end;

{13/10/06 : Pour les flux ICC
{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.CbFlux1OnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  SetControlCaption('LBCODEFLUX1', GetControlText('CODEFLUX1'));
end;

{13/10/06 : Pour les flux ICC : les deux codes doivent être de sens inverses
{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.CbFluxOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
{ Pur éviter la répétition de la question, mise en commentaires
  if not VerifSensFlux then begin
    PGIError(TraduireMemoire('Les deux codes flux doivent être de sens opposé.'));
    THValComboBox(Sender).SetFocus;
  end;
}
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSAISIEFLUX.VerifSensFlux : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := True;
  if (GetControlText('CODEFLUX1') <> '') and (GetControlText('CODEFLUX') <> '') then begin
    Q := OpenSQL('SELECT DISTINCT(TTL_SENS) FROM TYPEFLUX LEFT JOIN FLUXTRESO  ON TFT_TYPEFLUX = TTL_TYPEFLUX ' +
                 'WHERE TFT_FLUX IN ("' + GetControlText('CODEFLUX1') + '", "' + GetControlText('CODEFLUX') + '")', True);
    Result := Q.RecordCount = 2;
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.EcranCloseQuery(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := LastError = 0;
  if CanClose then FCLoseQuery(Sender, CanClose)
              else PGIError(LastErrorMsg);
  LastError := 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.BCalculOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  s : string;
  Q : TQuery;
begin
  Q := OpenSQL('SELECT TTL_SENS FROM TYPEFLUX LEFT JOIN FLUXTRESO ON TFT_TYPEFLUX = TTL_TYPEFLUX ' +
                 'WHERE TFT_FLUX = "' + GetControlText('CODEFLUX') + '"', True);
  if not Q.EOF then s := Q.FindField('TTL_SENS').AsString;
  Ferme(Q);
  if s = '' then
    HShowMessage('0;'+ Ecran.Caption + ';Le sens du flux est irrécupérable.'#13'Le calcul ne peut être lancé.;W;O;O;O;', '', '')
  else begin
    InitTobCalculIcc;
    TheTob := tResultCalcIcc;
    s := TRLanceFiche_CalculIcc(edGeneral.Text + ';' + GetControlText('DATECOMPTABLE') + ';' + s + ';');
    if s <> '' then begin
      SetControlText('MONTANT', tResultCalcIcc.GetString('MONTANT'));
      SetControlText('DATECOMPTABLE', tResultCalcIcc.GetString('FIN'));
      tResultCalcIcc.PutValue('TE_DATEVALID', tResultCalcIcc.GetValue('DEBUT'));
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.EcranKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  if (Key = Ord('C')) and (ssCtrl in Shift) then
    BCalculOnClick(GetControl('BCALCUL'));
  FKeyDown(Sender, Key, Shift);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSAISIEFLUX.InitTobCalculIcc;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(tResultCalcIcc) then begin
    tResultCalcIcc  := TOB.Create('µCALCULICC', nil, -1);
    tResultCalcIcc.AddChampSupValeur('MONTANT'     , 0.0);
    tResultCalcIcc.AddChampSupValeur('DEBUT'       , iDate1900);
    tResultCalcIcc.AddChampSupValeur('FIN'         , iDate1900);
    tResultCalcIcc.AddChampSupValeur('TE_DATEVALID', iDate1900);
  end
  else begin
    tResultCalcIcc.SetDouble('MONTANT', 0.0);
    tResultCalcIcc.SetDateTime('DEBUT'       , iDate1900);
    tResultCalcIcc.SetDateTime('FIN'         , iDate1900);
    tResultCalcIcc.SetDateTime('TE_DATEVALID', iDate1900);
  end;
end;

initialization
  RegisterClasses([TOF_TRSAISIEFLUX]);

end.
