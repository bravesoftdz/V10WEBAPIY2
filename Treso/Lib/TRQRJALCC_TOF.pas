{ Unité : Source TOF de la FICHE : TRQRJALCC : Journal des comptes courants
--------------------------------------------------------------------------------------
    Version   |   Date | Qui |   Commentaires
--------------------------------------------------------------------------------------
 7.05.001.001  17/10/06  JP   Création de l'unité
 8.00.002.002  30/07/07  JP   Nouvelle gestion des soldes
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.001.010  17/09/07  JP   FQ 10519 : Gestion des taux qui couvrent toute la période
--------------------------------------------------------------------------------------}
unit TRQRJALCC_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eQRS1, MaineAGL,
  {$ELSE}
  QRS1, FE_Main, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, uTob, UTOF, HEnt1, uAncetreEtat;

type
  TOF_TRQRJALCC = class(TRANCETREETAT)
    procedure OnArgument(S : string); override;
    procedure OnUpdate              ; override;
    procedure OnClose               ; override;
    procedure OnLoad                ; override;
  private
    Iccok : Boolean;
    TobTaux : TOB;
    TobCalc : TOB;
    Ligne   : Integer;

    procedure NoDossierChange(Sender : TObject);
    procedure RemplitTableTemp;
    function  GenereInterets(DateD, DateF : TDateTime; Solde : Double; IccOk : Boolean = False) : Boolean;
    function  PrepareImpression : TOB;
  end;


procedure TRLance_JournalCC(Arguments : string);


implementation


uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Commun, Constantes, Math, ULibExercice, UProcSolde, HMsgBox, HDebug;

{---------------------------------------------------------------------------------------}
procedure TRLance_JournalCC(Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('TR', 'TRQRJALCC', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCC.OnArgument(S : string );
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  IccOk := ReadTokenSt(s) = 'ICC';
  if Iccok then begin
    Ecran.HelpContext := 50000153;
    TFQRS1(Ecran).NatureEtat := 'TJI';
    TFQRS1(Ecran).CodeEtat := 'TJI';
    TobTaux := Tob.Create('µTAUX', nil, -1);
  end else begin
    Ecran.HelpContext := 150;
    TFQRS1(Ecran).NatureEtat := 'TJC';
    TFQRS1(Ecran).CodeEtat := 'TJC';
  end;

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');
  THEdit(GetControl('TE_GENERAL_')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, '');
  THEdit(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
  THEdit(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
  SetControlText('XX_WHERE', 'TE_CODECIB = "' + CODECIBCOURANT + '"');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCC.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if IccOk and Assigned(TobTaux) then FreeAndNil(TobTaux);
  {On vide la table temporaire en sortant}
//  ExecuteSQL('DELETE FROM ICCEDTTEMP WHERE ICZ_UTILISATEUR = "' + V_PGI.User + '"');
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCC.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if GetControlText('TE_GENERAL' ) = '' then SetControlText('TE_GENERAL' , '00000000000000000');
  if GetControlText('TE_GENERAL_') = '' then SetControlText('TE_GENERAL_', 'ZZZZZZZZZZZZZZZZZ');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCC.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if Iccok then
    RemplitTableTemp
  else
    SetControlText('DEBANNEE', DateToStr(DebutAnnee(StrToDate(GetControlText('TE_DATEVALEUR')))));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCC.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL', '');
  THEdit(GetControl('TE_GENERAL_')).Plus := FiltreBanqueCp(tcp_Tous, tcb_Courant, GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL_', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRQRJALCC.RemplitTableTemp;
{---------------------------------------------------------------------------------------}
var
  TobEcr    : TOB;
  n         : Integer;
  OldDt     : TDateTime;
  OldDos    : string;
  OldCpt    : string;
  Solde     : Double;
  Exo       : TZExercice;
  Coherence : Boolean;

    {------------------------------------------------------------------}
    procedure _MajTable;
    {------------------------------------------------------------------}
    begin
      {30/07/07 : s'il s'agit d'une écriture d'ICC}
      if TobEcr.Detail[n].GetString('TFT_CLASSEFLUX') = cla_ICC then begin
        TobCalc := TobEcr.Detail[n];
        if not GenereInterets(iDate1900, TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR'), 0, True) then Coherence := False;
        {Le champ Te_DATEVALID contient la date de départ et TE_DATEVALEUR la date de fin de l'ICC}
        if not GenereInterets(Max(TobCalc.GetDateTime('TE_DATEVALID') - 1, OldDt), TobCalc.GetDateTime('TE_DATEVALEUR'), Solde) then Coherence := False;

        {30/07/07 : On récupère plus les soldes dans la base, mais on les calcule à la volée}
        if IsNewSoldes then Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT')
                       else Solde := TobEcr.Detail[n].GetDouble('TE_SOLDEDEVVALEUR');
        {On mémorise la nouvelle date}
        OldDt := TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR');
      end

      {On change de date de valeur, on calcule les intérêts sur le solde précédent}
      else begin
        {Calcul des intérêts}
        if not GenereInterets(OldDt, TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR'), Solde) then Coherence := False;
        {On mémorise la nouvelle date}
        OldDt := TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR');
        {30/07/07 : On récupère plus les soldes dans la base, mais on les calcule à la volée}
        if IsNewSoldes then Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT')
                       else Solde := TobEcr.Detail[n].GetDouble('TE_SOLDEDEVVALEUR');
      end;
    end;

begin
  Ligne := 1;
  Coherence := True;
  {Préparation de la table d'impression et des tobs}
  TobEcr := PrepareImpression;
  if TobEcr = nil then begin
    HShowMessage('0;' + Ecran.Caption + ';Impossible de charger les écritures;E;O;O;O;', '', '');
    Exit;
  end;

  if TobTaux.Detail.Count = 0 then begin
    HShowMessage('0;' + Ecran.Caption + ';Impossible de récupérer les taux.;E;O;O;O;', '', '');
    Exit;
  end;

  try
    if TobEcr.Detail.Count > 0 then begin
      OldDos   := TobEcr.Detail[0].GetString('DOS_NOMBASE');
      OldCpt   := TobEcr.Detail[0].GetString('TE_GENERAL');
      Exo      := TZExercice.Create(True, OldDos);
      OldDt    := StrToDateTime(GetControlText('TE_DATEVALEUR'));
      Solde    := GetSoldeInit(OldCpt, DateToStr(OldDt), '', True, True);
      TobCalc  := TobEcr.Detail[0];

      for n := 0 to TobEcr.Detail.Count - 1 do begin
        if (OldDos = TobEcr.Detail[n].GetString('DOS_NOMBASE')) then begin
          if (OldCpt = TobEcr.Detail[n].GetString('TE_GENERAL')) then begin
            if (OldDt = TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR')) then begin
              {30/07/07 : On récupère plus les soldes dans la base, mais on les calcule à la volée}
              if IsNewSoldes then Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT')
                             else Solde := TobEcr.Detail[n].GetDouble('TE_SOLDEDEVVALEUR');
            end
            else
              _MajTable;
          end
          {On change de compte}
          else begin
            {Calcul des intérêts}
            if not GenereInterets(OldDt, StrToDateTime(GetControlText('TE_DATEVALEUR_')), Solde) then Coherence := False;
            {On charge les infos du nouveau compte}
            OldCpt := TobEcr.Detail[n].GetString('TE_GENERAL');
            OldDt  := StrToDateTime(GetControlText('TE_DATEVALEUR'));//Exo.QuelExoDate(TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR')).Deb;
            Solde  := GetSoldeInit(OldCpt, DateToStr(Exo.QuelExoDate(OldDt).Deb), '', True, True);
            Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT');
            TobCalc := TobEcr.Detail[n];
            _MajTable;
          end;
        end
        {On change de dossier, on recharge tout}
        else begin
          {Calcul des intérêts}
          if not GenereInterets(OldDt, StrToDateTime(GetControlText('TE_DATEVALEUR_')), Solde) then Coherence := False;
          {On charge les infos du nouveau dossier}
          OldDos := TobEcr.Detail[n].GetString('DOS_NOMBASE');
          OldCpt := TobEcr.Detail[n].GetString('TE_GENERAL');
          OldDt  := StrToDateTime(GetControlText('TE_DATEVALEUR'));
          if Assigned(Exo) then FreeAndNil(Exo);
          Exo    := TZExercice.Create(True, OldDos);
          Solde  := GetSoldeInit(OldCpt, DateToStr(Exo.QuelExoDate(OldDt).Deb), '', True, True);
          //Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT');
          TobCalc := TobEcr.Detail[n];
          _MajTable;
        end;
        TobCalc := TobEcr.Detail[n];
      end; {for n := 0 to }

      {Calcul des intérêts, pour le dernier solde}
      TobCalc := TobEcr.Detail[TobEcr.Detail.Count - 1];
      {Si la dernière écriture est une écriture d'ICC, on prend le dernier solde avant l'ICC
      if TobCalc.GetString('TFT_CLASSEFLUX') = cla_Icc then
        if TobEcr.Detail.Count > 2 then
          Solde := TobEcr.Detail[TobEcr.Detail.Count - 2].GetDouble('TE_SOLDEDEVVALEUR');}

      if not GenereInterets(OldDt, StrToDateTime(GetControlText('TE_DATEVALEUR_')), Solde) then Coherence := False;

    end;{if TobEcr.Detail.Count}
  finally
    if Assigned(TobEcr) then FreeAndNil(TobEcr);
    if Assigned(Exo   ) then FreeAndNil(Exo);
    if not Coherence then
      HShowMessage('0;' + Ecran.Caption + ';La saisie des taux n''est pas cohérente.'#13'Les intérêts calculés ne sont qu''une approximation.;E;O;O;O;', '', '');
    TFQRS1(Ecran).WhereSQL := '';
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRQRJALCC.PrepareImpression : TOB;
{---------------------------------------------------------------------------------------}
begin
  TobTaux.ClearDetail;
  TobTaux.LoadDetailFromSQL('SELECT * FROM ICCTAUXCOMPTE WHERE (ICD_DATEDU BETWEEN "' +
                   UsDateTime(StrToDateTime(GetControlText('TE_DATEVALEUR'))) + '" AND "' +
                   UsDateTime(StrToDateTime(GetControlText('TE_DATEVALEUR_'))) + '") ' +
                   'OR (ICD_DATEAU BETWEEN "' +
                   UsDateTime(StrToDateTime(GetControlText('TE_DATEVALEUR'))) + '" AND "' +
                   UsDateTime(StrToDateTime(GetControlText('TE_DATEVALEUR_'))) + '") ' +
                   {17/09/07 : FQ 10519 : Gestion des taux qui couvent toute la période}
                   'OR (ICD_DATEAU <= "' +
                   UsDateTime(StrToDateTime(GetControlText('TE_DATEVALEUR'))) + '" AND ICD_DATEDU >= "' +
                   UsDateTime(StrToDateTime(GetControlText('TE_DATEVALEUR'))) + '") ' +
                   'AND (ICD_GENERAL BETWEEN "' + GetControlText('TE_GENERAL') + '" AND "' + GetControlText('TE_GENERAL_') + '") ' +
                   'ORDER BY ICD_GENERAL, ICD_DATEAU');
  ExecuteSQL('DELETE FROM ICCEDTTEMP WHERE ICZ_UTILISATEUR = "' + V_PGI.User + '"');
  Result := TOB.Create('µTRECRITURE', nil, -1);
  Result.LoadDetailFromSQL('SELECT TE_CLEVALEUR, TFT_CLASSEFLUX, TE_DATEVALEUR, TE_GENERAL, TE_MONTANT,'+
                           'BQ_LIBELLE, DOS_NOMBASE, TE_SOLDEDEVVALEUR, TE_DATEVALID, ' +
                           'IIF(TFT_CLASSEFLUX = "' + cla_ICC + '", "Z", "A") AS CHPICC FROM TRECRITURE ' +
                           'LEFT JOIN DOSSIER ON DOS_NODOSSIER = TE_NODOSSIER ' +
                           'LEFT JOIN FLUXTRESO ON TFT_FLUX = TE_CODEFLUX ' +
                           'LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ' +
                           'WHERE ' + TFQRS1(Ecran).WhereSQL + ' ' +
                           'ORDER BY TE_NODOSSIER, TE_GENERAL, TE_DATEVALEUR, CHPICC, TE_CLEVALEUR');
end;

{---------------------------------------------------------------------------------------}
function TOF_TRQRJALCC.GenereInterets(DateD, DateF : TDateTime; Solde : Double; IccOk : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Interets  : Double;

      {--------------------------------------------------------------------}
      procedure CalculTaux;
      {--------------------------------------------------------------------}
      var
        Taux  : Double;
        T     : TOB;
        n     : Integer;
        Jours : Integer;
        Test  : TDateTime;
      begin
        Jours := Trunc(DateF - DateD);
        {Si l'opération n'est pas la dernière du jour, on sort car ce que le veut
         récupérer et traiter, c'est le solde à la fin de la journée}
        if Jours = 0 then Exit;
        {Test sert à vérifier qu'il y ait des taux sur toute la période et aux calculs par niveaux}
        Test := DateD;

        for n := 0 to TobTaux.Detail.Count - 1 do begin
          T := TobTaux.Detail[n];
          if T.GetString('ICD_GENERAL') <> TobCalc.GetString('TE_GENERAL') then Continue;
          {Si la période de validité du solde courant ne correspond pas à
           la validité du taux courant, on passe au taux suivant}
          if T.GetDateTime('ICD_DATEAU') < DateD then Continue;
          {Si la période de validité du solde courant est antérieure
           à la validité du taux courant, on quitte la boucle}
          if T.GetDateTime('ICD_DATEDU') > DateF then Break;

          {Il y a des trous dans la saisie des taux}
          if (Test + 1) < T.GetDateTime('ICD_DATEDU') then
            Result := False;
          Taux  := T.GetDouble('ICD_TAUX');
          Jours := Trunc(Min(T.GetDateTime('ICD_DATEAU'), DateF) - Test);
          if Test = StrToDateTime(GetControlText('TE_DATEVALEUR')) then Inc(Jours);
          Interets := Interets + (Jours * Solde * Taux) / (100 * 365);

          Test := T.GetDateTime('ICD_DATEAU');
          if Test > DateF then Break;
        end;
        {Les taux ne couvrent pas toutes la période de validité du solde courant}
        if DateF > Test then Result := False;
      end;
var
  Lib  : string;
  tIcc : TOB;
begin
  Interets := 0;
  Result   := True;
  if (Trunc(DateF - DateD) = 0) and not IccOk then Exit;
  Lib := '';

//  Debug('Début : ' + DateToStr(DateD));
  //Debug('Fin : ' + DateToStr(DateF));

  if not IccOk then begin
    CalculTaux;
    Arrondi(Interets, V_PGI.OkDecV);
  end;

  tIcc := TOB.Create('ICCEDTTEMP', nil, -1);
  try
    tIcc.SetString('ICZ_UTILISATEUR', V_PGI.User);
    tIcc.SetString('ICZ_GENERAL'    , TobCalc.GetString('TE_GENERAL'));
    tIcc.SetString('ICZ_LIBELLE'    , TobCalc.GetString('BQ_LIBELLE'));
    tIcc.SetDouble('ICZ_SOLDE'      , Solde);
    tIcc.SetDateTime('ICZ_DATE'     , DateF);
    tIcc.SetInteger('ICZ_NBJOURS'   , Trunc(DateF - DateD));
    tIcc.SetInteger('ICZ_LIGNE'     , Ligne);
    Inc(Ligne);
    if not IccOk then begin
      if DateD = StrToDateTime(GetControlText('TE_DATEVALEUR')) then
        tIcc.SetInteger('ICZ_NBJOURS'   , Trunc(DateF - DateD) + 1)
      else
        tIcc.SetInteger('ICZ_NBJOURS'   , Trunc(DateF - DateD));

      tIcc.SetInteger('ICZ_SURCAPITAL', 0);
      if Solde = 0 then
        tIcc.SetDouble('ICZ_TAUX'     , 0)
      else
        tIcc.SetDouble('ICZ_TAUX'     , Arrondi((Interets / Solde / Trunc(DateF - DateD)) * 100 * 365, V_PGI.OkDecP));
      tIcc.SetDouble('ICZ_INTERET'    , Interets);
    end
    else begin
      tIcc.SetDouble('ICZ_INTERET'    , TobCalc.GetDouble('TE_MONTANT'));
      tIcc.SetInteger('ICZ_NBJOURS'   , Trunc(TobCalc.GetDateTime('TE_DATEVALEUR') - TobCalc.GetDateTime('TE_DATEVALID') + 1));
      tIcc.SetInteger('ICZ_SURCAPITAL', 100);
    end;
    tIcc.InsertDb(nil);
  finally
    FreeAndNil(tIcc);
  end
end;

initialization
  RegisterClasses([TOF_TRQRJALCC]);

(*
      for n := 0 to TobEcr.Detail.Count - 1 do begin
        if (OldDos = TobEcr.Detail[n].GetString('DOS_NOMBASE')) then begin
          if (OldCpt = TobEcr.Detail[n].GetString('TE_GENERAL')) then begin
            if (OldDt = TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR')) then begin
              {30/07/07 : On récupère plus les soldes dans la base, mais on les calcule à la volée}
              if IsNewSoldes then Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT')
                             else Solde := TobEcr.Detail[n].GetDouble('TE_SOLDEDEVVALEUR');
            end
            {On change de date de valeur, on calcule les intérêts sur le solde précédent}
            else begin
              {Calcul des intérêts}
              if not GenereInterets(OldDt, TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR'), Solde) then Coherence := False;
              {On mémorise la nouvelle date}
              OldDt := TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR');
              {30/07/07 : On récupère plus les soldes dans la base, mais on les calcule à la volée}
              if IsNewSoldes then Solde := Solde + TobEcr.Detail[n].GetDouble('TE_MONTANT')
                             else Solde := TobEcr.Detail[n].GetDouble('TE_SOLDEDEVVALEUR');
            end;
          end
          {On change de compte}
          else begin
            {Calcul des intérêts}
            if not GenereInterets(OldDt, StrToDateTime(GetControlText('TE_DATEVALEUR_')), Solde) then Coherence := False;
            {On charge les infos du nouveau compte}
            OldCpt := TobEcr.Detail[n].GetString('TE_GENERAL');
            OldDt  := Exo.QuelExoDate(TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR')).Deb;
            Solde  := GetSoldeInit(OldCpt, DateToStr(OldDt), '', True, True);
          end;
        end
        {On change de dossier, on recharge tout}
        else begin
          {Calcul des intérêts}
          if not GenereInterets(OldDt, StrToDateTime(GetControlText('TE_DATEVALEUR_')), Solde) then Coherence := False;
          {On charge les infos du nouveau dossier}
          OldDos := TobEcr.Detail[n].GetString('DOS_NOMBASE');
          OldCpt := TobEcr.Detail[n].GetString('TE_GENERAL');
          OldDt  := TobEcr.Detail[n].GetDateTime('TE_DATEVALEUR');
          if Assigned(Exo) then FreeAndNil(Exo);
          Exo    := TZExercice.Create(True, OldDos);
          Solde  := GetSoldeInit(OldCpt, DateToStr(Exo.QuelExoDate(OldDt).Deb), '', True, True);
        end;
        TobCalc := TobEcr.Detail[n];
      end; {for n := 0 to }

      {Calcul des intérêts, pour le dernier solde}
      TobCalc := TobEcr.Detail[TobEcr.Detail.Count - 1];

      if not GenereInterets(OldDt, StrToDateTime(GetControlText('TE_DATEVALEUR_')), Solde) then Coherence := False;

*)

end.

