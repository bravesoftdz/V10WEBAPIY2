{-------------------------------------------------------------------------------------
    Version    |   Date | Qui |   Commentaires
--------------------------------------------------------------------------------------
               11/03/04  BPY  Création de l'unité
               21/09/04  BPY  modification diverse et variée sur  "l'etat recapitulatif si pointage
                              complet" a la demande de la FFF
               19/11/04  SG6  FQ 14969 (pb sous DB2) : Modification des stDate1900 par usdatetime(istdate1900)
               28/07/05  JP   FQ 13731 : Gestion des comptes en devises
08.00.001.006  22/03/07  JP   Gestion du nouveau pointage
08.00.001.013  03/05/07  JP   Gestion des mouvements bancaires saisis manuellement dans le sens bancaire
08.00.001.018  01/06/07  JP   FQ TRESO 10470 : Aliassage des champs calculés des requêtes SQl pour Oracle
08.00.001.025  16/07/07  JP   FQ 21056 : gestion des mvts dont la date de pointage est postérieure à la session en cours
08.10.001.005  21/08/07  JP   FQ 21256 : récupération de la référence de pointage lorsque l'on vient de la
                              consultation des comptes généraux
08.10.001.010  18/09/07  JP   FQ 21331 : possibilité de choisir une date d'arrêté qui ne correspond à aucune session.
                              Dans ce cas, on traite toutes les écritures jusqu'à la date et les mouvements jusqu'à
                              la dernière référence de pointage. Pas de possibilité de faire un récapitulatif.
08.10.001.010  19/09/07  JP   FQ 21392 : On rebranche la gestion automatique du Récapitulatif, uniquement en PCL
08.10.001.012  02/10/07  JP   Ajout d'un critère pour savoir de quel produit on vient dans l'état
08.10.005.001  14/11/07  JP   Gestion des comptes pointables qui ne sont pas bancaires
--------------------------------------------------------------------------------------}

unit CPRAPPRODET_TOF;

//================================================================================
// Interface
//================================================================================
interface

uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    UTOB,
    eQRS1,
    MaineAGL,
{$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    QRS1,
    FE_Main,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    LookUp,
    Ent1,
    uLibWindows,
    Saisutil,
    TofMeth,
    Htb97,      //  TToolBarButton97
    CpteSav     // CTrouveContrePartie
    ;

//==================================================
// Externe
//==================================================
procedure CC_LanceFicheEtatRapproDet( vStArgument : string = '');

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPRAPPRODET = Class(TOF_METH)
        procedure OnNew                  ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnClose                ; override ;

    private
        glbl_Date : THLabel;
        gtxt_DatePointage : THEdit;
        ggbx_Sens : THRadioGroup;
        gtxt_General : THEdit;
        FStCompteUtiliser : string;
        gtxt_Devise : THEdit;
        gtxt_DevPivot : THEdit;
        ggbx_Devise : THRadioGroup;
        gcbx_Recap : TCheckBox;

        XX_Rupture : THEdit;
        TE_General : THLabel;

        StType : string3;
        StDevise : string3;
        FStGeneral : string;

        FStArgument  : string;
        NomBase      : string;
        FLanceAutoOk : Boolean; {JP 21/08/07 : Pour le F9 Automatique}
        FDateDebExo  : TDateTime; {JP 22/03/07 : Date des anouveaux de références}
        FCritGeneral : string;{JP 21/08/07 : Pour mémoriser le critère du OnArgument}
        FCritDatePtg : string;{JP 21/08/07 : Pour mémoriser le critère du OnArgument}

        procedure ggbx_DeviseClick(Sender : TObject);
        procedure OnExitGeneral(Sender : TOBject);
        procedure OnExitDatePointage(Sender : TOBject);
        procedure OnElipsisClickE_General(Sender : TObject);
        procedure ChargeRefPointage;
        procedure DatePointageOnClick (Sender : TObject);

        {$IFDEF TRESO}
        procedure BqCodeOnElipsisClick(Sender : TObject);
        procedure BqCodeOnChange      (Sender : TObject);
        {$ENDIF TRESO}
        {JP 28/07/05 : FQ 13731 : Pour éxecuter OnExitGeneral après l'éventuel chargement du filtre}
        procedure FormAfterShow;
    end;

//================================================================================
// Implementation
//================================================================================
implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UtilPgi, Commun, Constantes, ULibExercice, ULibPointage, UProcSolde, ParamSoc;

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CC_LanceFicheEtatRapproDet( vStArgument : string = '');
begin
    AGLLanceFiche('CP', 'CPRAPPRODET', '', '', vStArgument );
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... : 26/12/2005 (GCO)
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRAPPRODET.OnNew;
var
  lSt : string;
  Q   : TQuery;
begin
  inherited;
  FCritGeneral := '';
  FCritDatePtg := '';
  
  if FStArgument <> '' then
  begin
    lSt := ReadTokenSt(FStArgument);
    if EstPointageSurTreso then begin
      SetFocusControl('BQCODE');
      SetControlText('BQCODE', lSt);
    end
    else begin
      // GCO - 26/11/2004 - FQ 13180
      gtxt_General.SetFocus;
      gtxt_General.Text := lSt;
    end;
    FStGeneral   := lSt;
    FCritGeneral := lSt;

    // Date de Pointage
    lSt := ReadTokenSt(FStArgument);
    
    {JP 21/08/07 : FQ 21256 : récupération de la date par défaut}
    if lSt = '' then begin
      Q := OpenSQL('SELECT MAX(EE_DATEPOINTAGE) ADATE FROM ' + GetTableDossier(NomBase, 'EEXBQ') + ' WHERE EE_GENERAL = "' +
                   gtxt_General.Text + '" AND EE_DATEPOINTAGE <= "' + UsDateTime(Date) + '"', True);
      if not Q.EOF then
        lSt := Q.FindField('ADATE').AsString;
      Ferme(Q);
    end;

    if lSt <> '' then
      gtxt_DatePointage.Text := lSt;
    FCritDatePtg := lSt;

    {JP 21/08/07 : FQ 21256 : gestion automatique de l'impression s'il y a des critères passés en argument}
    FLanceAutoOk := (FCritDatePtg <> '') and (FCritGeneral <> '');

    {JP 22/03/07 : pour forcer le chargement de la référence de pointage lorsque l'on appelle depuis le pointage}
    ChargeRefPointage;

    // F9 Auto
    lSt := ReadtokenSt(FStArgument);
    {JP 21/08/07 : FQ 21256 : gestion automatique de l'impression s'il y a des critères passés en argument}
    if (lSt = 'X') or FLanceAutoOk then
      TToolBarButton97(GetControl('BVALIDER', True)).Click
    else
      FLanceAutoOk := False;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... : 22/03/2007
Description .. : JP 22/03/07 : gestion du nouveau pointage (EstSpecif)
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRAPPRODET.OnUpdate;
var
    StV8 : String;
    Query, Query1, Query2 : string;
    lStControleDevise : string;
begin
    Inherited;
    {02/10/07 : Pour dire que l'état de rappro est sur TRECRITURE}
    if (ctxTreso in V_PGI.PGIContexte) and GetParamSocSecur('SO_TRPOINTAGETRESO', False) then
      SetControlText('PRODUIT', 'TRESO')
    else
      SetControlText('PRODUIT', 'COMPTA');
                                     
    // Remplissage du THEdit LECOMPTE
    SetControlText('LECOMPTE', FStCompteUtiliser);

    if EstPointageSurTreso then begin
      Query1 := 'SELECT 1 IDN, TE_JOURNAL JAL, TE_DATECOMPTABLE DCP, TE_REFINTERNE REF, TE_DEVISE DEV, TE_GENERAL GEN, TE_DATECOMPTABLE DVA,' +
                'TE_NUMEROPIECE NUM, TE_LIBELLE LIB, ';
      Query1 := Query1 +
                'IIF((TE_MONTANT >= 0), 0, (-1) * TE_MONTANT) DEN, ' +
                'IIF((TE_MONTANT < 0 ), 0, TE_MONTANT) CRN, ' +
                'IIF((TE_MONTANTDEV >= 0), 0, (-1) * TE_MONTANTDEV) DED, ' +
                'IIF((TE_MONTANTDEV < 0 ), 0, TE_MONTANTDEV) CRD, ' +
                'TE_NUMTRANSAC RPA, BQ_LIBELLE G_LIBELLE';

      Query1 := Query1 +
            ' FROM BANQUECP, TRECRITURE WHERE BQ_CODE = TE_GENERAL' +
            ' AND TE_NATURE = "' + na_Realise + '" AND TE_CODECIB < "' + CODECIBTITRE + '" AND TE_NUMTRANSAC NOT LIKE "' +
            CODEMODULECOU  + '%" AND TE_CODEFLUX <> "' + CODEREGULARIS + '" AND TE_GENERAL="' + GetControlText('BQCODE') + '"';

      {Sélection uniquement des lignes d'ecriture plus vieille}
      Query1 := Query1 + ' AND TE_DATECOMPTABLE <= "' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';

      {État recapitulatif ....}
      if (gcbx_Recap.Checked) then
        Query1 := Query1 + ' AND TE_REFPOINTAGE="' + glbl_Date.Caption + '" AND (NOT TE_DATERAPPRO = "' + USDateTime(iDate1900) + '")'
      else
        Query1 := Query1 + ' AND (((TE_REFPOINTAGE <> "" OR TE_REFPOINTAGE IS NOT NULL) AND TE_DATERAPPRO > "' +
                 USDateTime(StrToDate(gtxt_DatePointage.Text)) + '") OR (TE_REFPOINTAGE="" OR TE_REFPOINTAGE IS NULL))';


      {---- Requête sur EEXBQLIG ----}
      Query2 := 'SELECT 2 IDN, CEL_CODEAFB JAL, CEL_DATEOPERATION DCP, CEL_REFPIECE REF, CEL_DEVISE DEV, CEL_GENERAL GEN, CEL_DATEVALEUR DVA, CEL_NUMRELEVE' +
                ' NUM,CEL_LIBELLE LIB, ';
      {03/05/07 : Dans le sens bancaire, les mouvements saisis manuellement viennent en correction du
                  solde bancaire, mais ne l'impacte pas, il faut donc inverser le sens pour que le solde
                  bancaire final soit impacté dans le bon sens. Dans le sens comptable le problème ne se
                  pose pas, car les mouvements concernés impactent le solde comptable pour aboutir à un
                  solde théorique}
      if ggbx_Sens.value = '0' then begin
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_CREDITEURO, CEL_DEBITEURO) DEN, ';
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_DEBITEURO, CEL_CREDITEURO) CRN, ';
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_CREDITDEV, CEL_DEBITDEV) DEV, ';
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_DEBITDEV, CEL_CREDITDEV) CRD, ';
        Query2 := Query2 + 'STR(CEL_NUMRELEVE) RPA, BQ_LIBELLE G_LIBELLE';
      end
      else begin
        Query2 := Query2 + ' CEL_DEBITEURO DEN, CEL_CREDITEURO CRN, CEL_DEBITDEV DEV, CEL_CREDITDEV CRD, STR(CEL_NUMRELEVE) RPA, BQ_LIBELLE G_LIBELLE';
      end;

      Query2 := Query2 + ' FROM BANQUECP, EEXBQLIG WHERE BQ_CODE = CEL_GENERAL ';
      {Sélection uniquement des lignes de releve plus vieille}
      Query2 := Query2 + ' AND CEL_DATEOPERATION <= "' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';
      {État recapitulatif ....}
      if (gcbx_Recap.Checked) then
        Query2 := Query2 + ' AND (NOT CEL_DATEPOINTAGE ="' + usdatetime(iDate1900) + '") AND CEL_GENERAL="' + GetControlText('BQCODE') + '"'
      else
        {Élements non pointés ...}
//        Query2 := Query2 + ' AND CEL_DATEPOINTAGE="' + usdatetime(iDate1900) + '" AND CEL_GENERAL="' + GetControlText('BQCODE') + '"';
        {16/07/07 : FQ 21056 : Gestion des dates de pointage postérieures}
        Query2 := Query2 + ' AND (CEL_DATEPOINTAGE="' + usdatetime(iDate1900) + '" OR CEL_DATEPOINTAGE > "' +
                  USDateTime(StrToDate(gtxt_DatePointage.Text)) + '") AND CEL_GENERAL="' + GetControlText('BQCODE') + '"';

      {---- Vérification des devises}
      if (not (StDevise = '')) then begin
        {On réutilise la requete sur ecriture pour vérifier la présence d'écriture avec devise différente}
        lStControleDevise := Query1 + ' AND TE_DEVISE <> "' + StDevise + '"';

        {Test d'existence des ecritures sur d'autres DEVISES que celle du compte de banque}
        if ExisteSql(lStControleDevise) then begin
          if (VH^.PointageJal) then PgiInfo('Attention  : certaines écritures non pointées du journal ' + gtxt_general.Text + ' ont ' +
                                            'une monnaie de saisie différente de ' + StDevise , Ecran.Caption)
          else PgiInfo('Attention  : certaines écritures non pointées du compte ' + gtxt_general.Text + ' ont ' +
                       'une monnaie de saisie différente de ' + StDevise , Ecran.Caption);
        end;

        if (not (StDevise = V_PGI.DevisePivot)) then begin
          Query1 := Query1 + 'AND TE_DEVISE ="' + StDevise + '"';
          Query2 := Query2 + 'AND CEL_DEVISE ="' + StDevise + '"';
        end;
      end;

      {===== UNION ALL et ORDER BY =====}
      Query := Query1 + ' UNION ALL ' + Query2 + ' ORDER BY IDN, GEN, DCP, NUM';
    end {if EstPointageSurTreso}

    else begin

      {---- Requête sur ECRITURE ----}
      Query1 := 'SELECT 1 IDN,E_JOURNAL JAL,E_DATECOMPTABLE DCP,E_REFINTERNE REF,E_DEVISE DEV,E_GENERAL GEN,E_DATECOMPTABLE DVA,' +
                'E_NUMEROPIECE NUM,E_LIBELLE LIB,E_DEBIT DEN,E_CREDIT CRN,E_DEBITDEV DED,E_CREDITDEV CRD,E_AUXILIAIRE RPA,G_GENERAL,G_LIBELLE';

      if (VH^.PointageJal) then
        Query1 := Query1 +
              ' FROM GENERAUX, ' + GetTableDossier(NomBase, 'ECRITURE') + ', JOURNAL WHERE J_JOURNAL=E_JOURNAL AND G_GENERAL=J_CONTREPARTIE'
            + ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")'
            + ' AND E_JOURNAL="' + gtxt_general.Text + '" AND E_GENERAL<>J_CONTREPARTIE'
      else
        Query1 := Query1 +
              ' FROM GENERAUX, ' + GetTableDossier(NomBase, 'ECRITURE') + ', JOURNAL WHERE J_JOURNAL=E_JOURNAL AND G_GENERAL=E_GENERAL'
            + ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")'
            + ' AND E_GENERAL="' + gtxt_general.Text + '"';

      // GCO - 07/11/2006 - FQ 17322
      Query1 := Query1 + ' AND E_CREERPAR <> "DET"';

      {Sélection uniquement des lignes d'ecriture plus vieille}
      Query1 := Query1 + ' AND E_DATECOMPTABLE<="' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';

      {État recapitulatif ....}
      if (gcbx_Recap.Checked) then
        Query1 := Query1 + ' AND E_REFPOINTAGE="' + glbl_Date.Caption + '" AND (NOT E_DATEPOINTAGE="' + usdatetime(iDate1900) + '")'
      else
        Query1 := Query1 + ' AND (((E_REFPOINTAGE<>"" OR E_REFPOINTAGE IS NOT NULL) AND E_DATEPOINTAGE>"' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '") OR (E_REFPOINTAGE="" OR E_REFPOINTAGE IS NULL))';

      {27/04/07 : FQ 20194 : exclusion des écritures antérieures à l'exo V8}
      StV8 := LWhereV8;
      if (StV8 <> '') then Query1 := Query1 + ' AND ' +StV8;

      {---- Requête sur EEXBQLIG ----}
      {03/05/07 : Dans le sens bancaire, les mouvements saisis manuellement viennent en correction du
                  solde bancaire, mais ne l'impacte pas, il faut donc inverser le sens pour que le solde
                  bancaire final soit impacté dans le bon sens. Dans le sens comptable le problème ne se
                  pose pas, car les mouvements concernés impactent le solde comptable pour aboutir à un
                  solde théorique}
      Query2 := 'SELECT 2 IDN,CEL_CODEAFB JAL,CEL_DATEOPERATION DCP,CEL_REFPIECE REF,CEL_DEVISE DEV,CEL_GENERAL GEN,CEL_DATEVALEUR DVA,CEL_NUMRELEVE' +
                ' NUM,CEL_LIBELLE LIB, ';
      if ggbx_Sens.value = '0' then begin
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_CREDITEURO, CEL_DEBITEURO) DEN, ';
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_DEBITEURO, CEL_CREDITEURO) CRN, ';
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_CREDITDEV, CEL_DEBITDEV) DEV, ';
        Query2 := Query2 + 'IIF(CEL_VALIDE = "X", CEL_DEBITDEV, CEL_CREDITDEV) CRD, ';
        Query2 := Query2 + 'STR(CEL_NUMRELEVE) RPA,G_GENERAL,G_LIBELLE';
      end
      else begin
        Query2 := Query2 + ' CEL_DEBITEURO DEN, CEL_CREDITEURO CRN, CEL_DEBITDEV DEV, CEL_CREDITDEV CRD, STR(CEL_NUMRELEVE) RPA, G_GENERAL, G_LIBELLE';
      end;

      if VH^.PointageJal then
        Query2 := Query2 + ' FROM GENERAUX, JOURNAL, ' + GetTableDossier(NomBase, 'EEXBQLIG') +
                ' WHERE AND J_JOURNAL = CEL_GENERAL AND G_GENERAL=J_CONTREPARTIE'
      else
        Query2 := Query2 + ' FROM GENERAUX, ' + GetTableDossier(NomBase, 'EEXBQLIG') +
                ' WHERE G_GENERAL = CEL_GENERAL';

      {Sélection uniquement des lignes de releve plus vieille}
      Query2 := Query2 + ' AND CEL_DATEOPERATION<="' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';

      {État recapitulatif ....}
      if (gcbx_Recap.Checked) then
        Query2 := Query2 + ' AND (NOT CEL_DATEPOINTAGE ="' + usdatetime(iDate1900) + '") AND CEL_GENERAL="' + gtxt_general.Text + '"'
      else
        {Élements non pointés ...}
//        Query2 := Query2 + ' AND CEL_DATEPOINTAGE="' + usdatetime(iDate1900) + '" AND CEL_GENERAL="' + gtxt_general.Text + '"';
        {16/07/07 : FQ 21056 : Gestion des dates de pointage postérieures}
        Query2 := Query2 + ' AND (CEL_DATEPOINTAGE="' + usdatetime(iDate1900) + '" OR CEL_DATEPOINTAGE > "' +
                  USDateTime(StrToDate(gtxt_DatePointage.Text)) + '") AND CEL_GENERAL="' + gtxt_general.Text + '"';

      {---- Vérification des devises}
      if (not (StDevise = '')) then begin
        {On réutilise la requete sur ecriture pour vérifier la présence d'écriture avec devise différente}
        lStControleDevise := Query1 + ' AND E_DEVISE <> "' + StDevise + '"';

        {Test d'existence des ecritures sur d'autres DEVISES que celle du compte de banque}
        if ExisteSql(lStControleDevise) then begin
          if (VH^.PointageJal) then PgiInfo('Attention  : certaines écritures non pointées du journal ' + gtxt_general.Text + ' ont ' +
                                            'une monnaie de saisie différente de ' + StDevise , Ecran.Caption)
          else PgiInfo('Attention  : certaines écritures non pointées du compte ' + gtxt_general.Text + ' ont ' +
                       'une monnaie de saisie différente de ' + StDevise , Ecran.Caption);
        end;

        if (not (StDevise = V_PGI.DevisePivot)) then begin
          Query1 := Query1 + 'AND E_DEVISE ="' + StDevise + '"';
          Query2 := Query2 + 'AND CEL_DEVISE ="' + StDevise + '"';
        end;
      end;
      //===== UNION ALL et ORDER BY =====
      // set de la requete dans le whereSQL de la fiche
      Query := Query1 + ' UNION ALL ' + Query2 + ' ORDER BY IDN, G_GENERAL, DCP, NUM';
    end; {if EstPointageSurTreso else if EstSpecif}

    TFQRS1(Ecran).WhereSQL := Query;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... : 01/06/2007
Description .. : JP 01/06/07 : FQ TRESO 10470 : gestion des alias des
Suite ........ : champs calculés des requêtes pour ORACLE (cf
Suite ........ : SOLDEEUR, SOLDEDEV)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.OnLoad;
var Q : TQuery;
    DatePointage : TDateTime;
    Solde, SoldeDev : Double;
    SoldeEuroB, SoldeDevB : Double;
    Exo : TExoDate;
    lStSql : string;
    NotOk  : Boolean;
    Mok    : Boolean;
    lGene  : string;
begin
  Inherited;
  {JP 21/08/07 : FQ 21256 : Les arguments ont priorités sur le filtre}
  if FCritGeneral <> '' then gtxt_General.Text := FCritGeneral;
  if FCritDatePtg <> '' then gtxt_DatePointage.Text := FCritDatePtg;

  if EstPointageSurTreso then lGene := GetControlText('BQCODE')
                         else lGene := gtxt_General.Text;

  // Met une date par défaut si la date est incorrecte
  if (gtxt_DatePointage.Text = '') or not IsValidDate(gtxt_DatePointage.Text) then
  begin
      Q := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM ' + GetTableDossier(NomBase, 'EEXBQ')
                   + ' WHERE EE_GENERAL="' + lGene + '" ORDER BY EE_DATEPOINTAGE DESC',True);
      if (not Q.EOF) then
      begin
          gtxt_DatePointage.Text := Q.Fields[0].AsString;
          glbl_Date.Caption      := Q.Fields[1].AsString;
      end
      else
      begin
          gtxt_DatePointage.Text := stDate1900;
          glbl_Date.Caption      := 'Aucune donnée';
      end;
      Ferme(Q);
  end;

  // Pour affichage dans l'état
  DatePointage := StrToDate(gtxt_DatePointage.Text);
  QuelDateDeExo(QUELEXODT(DatePointage),exo);

  {JP 22/03/07 : Gestion de la date des anouveaux de référence pour le nouveaux pointage}
 {On commence par rechercher la Date des a-nouveaux}
  FDateDebExo := VH^.ExoV8.Deb;
  if DatePointage >= GetEnCours.Deb then
    FDateDebExo := GetEnCours.Deb
  else begin
    QuelExoDate(StrToDate(gtxt_DatePointage.Text), StrToDate(gtxt_DatePointage.Text), Mok, Exo);
    if Exo.Deb > FDateDebExo then FDateDebExo := Exo.Deb;
  end;

  SoldeEuroB := 0;
  SoldeDevB  := 0;

  // Solde Banquaire
  Q := OpenSQL('SELECT EE_NEWSOLDEDEBEURO - EE_NEWSOLDECREEURO AS SOLDEUR, EE_NEWSOLDEDEB - EE_NEWSOLDECRE AS SOLDEDEV FROM ' +
               GetTableDossier(NomBase, 'EEXBQ') + ' WHERE EE_GENERAL="' + lGene + '" AND EE_REFPOINTAGE="' +
               glbl_Date.Caption + {'" AND EE_DATEPOINTAGE="' + USDateTime(DatePointage) +} '"',True);
  if not Q.EOF then begin
    SoldeEuroB := Q.FindField('SOLDEUR').AsFloat;
    SoldeDevB  := Q.FindField('SOLDEDEV').AsFloat;
  end
  else
    PgiError(TraduireMemoire('Le calcul des soldes est impossible : impossible de trouver la session de pointage'), Ecran.Caption);
    
  Ferme(Q);

  // SET DES SOLDE SUR LA FICHE
  SetControlText('SOLDEEUROB',FloatToStr(SoldeEuroB));
  SetControlText('SOLDEDEVB',FloatToStr(SoldeDevB));

  // Calcul du solde comptable réel
  if not EstPointageSurTreso then
    lStSql := 'SELECT SUM(E_CREDIT)-SUM(E_DEBIT) AS SOLDEEUR, SUM(E_CREDITDEV)-SUM(E_DEBITDEV) AS SOLDEDEV ' +
              'FROM ' + GetTableDossier(NomBase, 'ECRITURE') +
              ' WHERE E_GENERAL = "' + FStCompteUtiliser + '" AND ' +
              'E_QUALIFPIECE = "N" AND ((E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") OR ' +
              '(E_ECRANOUVEAU="OAN" AND E_DATECOMPTABLE="' + UsDateTime(FDateDebExo) + '"))  AND ' +
              'E_DATECOMPTABLE BETWEEN "' + USDateTime(FDateDebExo) + '" AND "'  + USDateTime(DatePointage) + '" ';

  // BPY le 31/07/2004 => Fiche 13731 :  Les écritures ne sont pas restituées,
  // par contre le solde comptable en tient compte ce qui a pour effet de déséquilibrer l'état de rapprochement.
  if (not ((StDevise = '') or (StDevise = V_PGI.DevisePivot))) then
    lStSql := lStSql + 'AND E_DEVISE ="' + StDevise + '"';

  if EstPointageSurTreso then begin
    Solde    := GetSoldeBancaire(lGene, ' AND TE_NATURE = "' + na_Realise + '"', DatePointage, False, True);
    SoldeDev := GetSoldeBancaire(lGene, ' AND TE_NATURE = "' + na_Realise + '"', DatePointage, False, False);
  end
  else begin
    Q := OpenSQL(lStSql,True);
    Solde     := Q.Fields[0].AsFloat;
    SoldeDev  := Q.Fields[1].AsFloat;
    Ferme(Q);
  end;

  // SET DES SOLDE SUR LA FICHE
  SetControlText('SOLDE',FloatToStr(Solde));
  SetControlText('SOLDEDEV',FloatToStr(SoldeDev));

  // Sens ...
  if (ggbx_Sens.visible) then
  begin
      if (ggbx_Sens.value = '0') then SetControlText('SENS','1')  // sens Bancaire
      else SetControlText('SENS','2');                            // sans comptable
  end
  else SetControlText('SENS','0');

  {JP 19/09/07 : FQ 21392 : on réactive la proposition automatique d'impression de récapitulatif.
                 Uniquement si en PCL (pour des raisons de performances sur la Table ECRITURE) et
                 si l'on vient des consultations ou du pointage}
  if FLanceAutoOk and (ctxPcl in V_PGI.PGIContexte) then begin
    NotOk := False;
    {S'il existe des lignes non pointées dans les lignes d'extrait bancaire alors pas de recap !
     FQ 18478 : 28/08/2006 : modifications du test sur les relevé auto : on recherche la présence
                             de lignes non pointées aussi sur les relevés précedents celui sélectionnés}
    lStSql := 'SELECT CEL_GENERAL FROM ' + GetTableDossier(NomBase, 'EEXBQLIG')
                  + ' WHERE CEL_GENERAL="' + lGene + '" AND '
                  + '(CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '" OR CEL_DATEPOINTAGE > "' + UsDateTime(DatePointage) + '")'
                  + ' AND ( CEL_NUMRELEVE <= (SELECT MAX(EE_NUMERO) FROM EEXBQ WHERE EE_GENERAL = "' + lGene + '" '
                                              + ' AND EE_REFPOINTAGE="' + glbl_Date.Caption + '")) ' ;

    {On regarde s'il existe des mouvements bancaires non pointés ...}
    if ExisteSQL(lStSQL) then
      NotOk := True
    {... Sinon on regarde s'il existe des écritures non pointées}
    else begin
      if EstPointageSurTreso then
        lStSql := 'SELECT TE_DATERAPPRO FROM TRECRITURE WHERE TE_DATECOMPTABLE<="' + USDateTime(DatePointage) + '"' +
                  ' AND ((TE_DATERAPPRO = "' + usdatetime(iDate1900) + '" AND TE_REFPOINTAGE = "") OR ' +
                        '(TE_DATERAPPRO > "' + USDateTime(DatePointage) + '" AND TE_REFPOINTAGE IS NOT NULL))' +
                  ' AND TE_NATURE = "' + na_Realise + '" AND TE_CODEFLUX <> "' + CODEREGULARIS + '"' +
                  ' AND TE_GENERAL = "' + lGene + '"'
      else begin
        lStSql := 'SELECT E_DATEPOINTAGE FROM ' + GetTableDossier(NomBase, 'ECRITURE') + ' WHERE E_DATECOMPTABLE<="' + USDateTime(DatePointage) + '"' +
                  ' AND ((E_DATEPOINTAGE = "' + usdatetime(iDate1900) + '" AND E_REFPOINTAGE = "") OR ' +
                        '(E_DATEPOINTAGE > "' + USDateTime(DatePointage) + '" AND E_REFPOINTAGE IS NOT NULL))' +
                  ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")';

        if VH^.PointageJal then
          lStSql := lStSql + ' AND E_JOURNAL = "' + lGene + '" AND E_GENERAL <> "' + FStCompteUtiliser + '"'
        else
          lStsql := lStSql + ' AND E_GENERAL = "' + lGene + '"';

        if VH^.ExoV8.Code <> '' then
          lStSql := lStSql + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '"';
      end;

      if (ExisteSQL(lstSql)) then NotOk := True;
    end;

    gcbx_Recap.Checked := not NotOk;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.OnArgument(S : String );
begin
    Inherited;
    FStArgument := S;

    FLanceAutoOk := FStArgument <> ''; {JP 21/08/07 : FQ 21256 : }

    Ecran.HelpContext := 7619000 ; //RR FQ 16068 : 999999427;
    {JP 28/07/05 : FQ 13731 : Pour éxecuter OnExitGeneral après l'éventuel chargement du filtre}
    TFQRS1(Ecran).OnAfterFormShow := FormAfterShow;

    // Compte general
    gtxt_General := THEdit(GetControl('EGENERAL',True));
    gtxt_General.OnExit := OnExitGeneral;
    gtxt_General.OnElipsisClick := OnElipsisClickE_General;
    // libelle du compte general
    TE_General := THLabel(GetControl('TE_GENERAL', True));

    // date de la reference de pointage
    gtxt_DatePointage := THEdit(GetControl('DATEPOINTAGE', True));
    gtxt_DatePointage.EditMask := '!99/99/0000;1;_';
    gtxt_DatePointage.OnExit := OnExitDatePointage;
    // libelle de la reference pointage
    glbl_Date := THLabel(GetControl('LIBDATE',true));

    // Pour l'affichage en fonction de la devise choisie
    ggbx_Devise := THRadioGroup(GetControl('GBXDEVISE',True));
    ggbx_Devise.OnClick := ggbx_DeviseClick;

    {JP 04/05/05 : FQ 15482 : il ne faut mettre 'EUR' en dur}
    ggbx_Devise.Items[0] := VH^.LibDevisePivot;
    SetControlText('DEVISE', V_PGI.DevisePivot);
    SetControlText('DEVPIVOT', VH^.LibDevisePivot);
    {FIN FQ 15482 : il ne faut mettre 'EUR' en dur}

    // Sens si on a un relevé manuel
    ggbx_Sens := THRadioGroup(GetControl('RBTSENSETAT',true));

    // check box pour l'etat recap !
    gcbx_Recap := TCheckBox(GetControl('AFFRECAPITULATIF',true));

    // Critere de rupture
    XX_Rupture := THEdit(GetControl('XX_RUPTURE', True));

    // modif si pointage sur journal ou pas!
    if VH^.PointageJal then
    begin
        TE_General.Caption     := traduirememoire('Journal') ;
        gtxt_General.MaxLength := 3;
        XX_RUPTURE.Text        := 'E_JOURNAL';
    end
    else
    begin
        if EstPointageSurTreso then begin
          TE_General.Caption     := TraduireMemoire('Compte bancaire') ;
          gtxt_General.MaxLength := 17;
          XX_RUPTURE.Text        := 'TE_GENERAL';
        end
        else begin
          TE_General.Caption     := traduirememoire('Compte général') ;
          gtxt_General.MaxLength := VH^.CPta[fbGene].Lg;
          XX_RUPTURE.Text        := 'E_GENERAL';
        end;
    end;

    // init des variables
    SetControlText('LIBDATE','');
    gtxt_Devise := THEdit(GetControl('DEVISE',True));
    gtxt_DevPivot := THEdit(GetControl('DEVPIVOT',True));
    StType := 'RIN';
    StDevise := '';
    FStGeneral := '';
    THEdit(GetControl('DATEPOINTAGE', True)).OnElipsisClick := DatePointageOnClick;

    {$IFDEF TRESO}
    if EstPointageSurTreso then begin
      SetControlVisible('EGENERAL', False);
      SetControlVisible('BQCODE', True);
      THEdit(GetControl('BQCODE', True)).OnElipsisClick := BqCodeOnElipsisClick;
      THEdit(GetControl('BQCODE', True)).OnChange := BqCodeOnChange;
    end
    else
      NomBase := V_PGI.SchemaName;
    {$ELSE}
    NomBase := '';
    {$ENDIF TRESO}


end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.OnClose;
begin
    if (not (IsValidDate(GetControlText('DATEPOINTAGE')))) then SetControlText('DATEPOINTAGE',stDate1900);

    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.ggbx_DeviseClick(Sender: TObject);
begin
    // Pour l'impression dans la monnaie sélectionnée
    if (ggbx_Devise.Value = '0') then
    begin
        {JP 04/05/05 : FQ 15482 : il ne faut mettre 'EUR' en dur}
        gtxt_Devise.Text := V_PGI.DevisePivot;
        gtxt_DevPivot.Text := VH^.LibDevisePivot;
    end
    else
    begin
        gtxt_Devise.Text := 'DEV';
        gtxt_DevPivot.Text := 'Devise';
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.OnExitGeneral(Sender: TOBject);
var
    lQuery : TQuery;
    lGene  : string;
begin
    if Trim(gtxt_General.Text) = '' then Exit;

    if (not VH^.PointageJal) and not EstPointageSurTreso then
    begin
        if (length(gtxt_General.Text) < VH^.Cpta[fbGene].Lg) then
          OnElipsisClickE_General(sender);
    end;

    if EstPointageSurTreso then lGene := GetControlText('BQCODE')
                           else lGene := gtxt_General.Text;

    // changement de la clause plus de la date
    // FQ 19329 - CA 22/12/2006 - Suppression AND car inutile et pose problème en DB2
    //gtxt_DatePointage.Plus := 'AND EE_GENERAL="' + gtxt_General.Text + '"';
    gtxt_DatePointage.Plus := ' EE_GENERAL="' + lGene + '"';

    // Récupération du compte de contrepartie si on est en pointage sur JAL
    if VH^.PointageJal then
      FStCompteUtiliser := CTrouveContrePartie( gtxt_general.Text )
    else
      FStCompteUtiliser := gtxt_General.Text;

    lQuery := nil;
    try
        try
            // Recherche la devise du compte Général ou du Compte de Contrepartie
          if EstPointageSurTreso then
            lQuery := OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_CODE = "' + GetControlText('BQCODE') + '"', True)
          else
            lQuery := OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL = "' + FStCompteUtiliser +
                              '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"', True);

            if (not lQuery.EOF) then
            begin
                StDevise := lQuery.FindField('BQ_DEVISE').AsString;

                if (StDevise = V_PGI.DevisePivot) then {JP 04/05/05 : FQ 15482}
                begin
                    ggbx_Devise.ItemIndex := 0;
                    ggbx_Devise.Enabled := true;
                end
                else
                begin
                    ggbx_Devise.ItemIndex := 1;
                    ggbx_Devise.Enabled := false;
                end;
            end;
        except
          on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message , 'Attention' );
        end;
    finally
      Ferme( lQuery );
    end;

    // Cherche la référence de pointage la plus récente et met rien si ya pas !
    // Si et seulement si on change de compte general ! => Fiche n° 14016
    if (not (FStGeneral = gtxt_General.Text)) then
    begin
      try
        try
          lQuery := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE, EE_ORIGINERELEVE FROM ' +
                    GetTableDossier(NomBase, 'EEXBQ') + ' WHERE EE_GENERAL="' + lGene +
                    '" ORDER BY EE_DATEPOINTAGE DESC',true);

          if (not lQuery.Eof) then
          begin
            gtxt_DatePointage.Text := lQuery.FindField('EE_DATEPOINTAGE').AsString;
            glbl_Date.Caption := lQuery.FindField('EE_REFPOINTAGE').AsString;
            StType := lQuery.FindField('EE_ORIGINERELEVE').AsString;
          end
          else
          begin
            gtxt_DatePointage.Text := '01/01/1900';
            glbl_Date.Caption := '';
            StType := 'RIN';
          end;

          {JP17/04/07 : Nouvel gestion des sessions de pointage}
          ggbx_Sens.Visible := (not IsReleveAuto(StType));
          {16/04/07 : FQ 19610 : je pense qu'il n'y a pas de raison de le désactiver pour une
                      session de pointage manuel, à voir ...
          if not IsReleveAuto(StType) then gcbx_Recap.Checked := False;
          gcbx_Recap.Enabled := IsReleveAuto(StType);}

        except
          on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message, 'Attention' );
        end;
      finally
        Ferme(lQuery);
      end;
    end;

    FStGeneral := gtxt_General.Text;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 16/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.OnExitDatePointage(Sender: TOBject);
begin
  ChargeRefPointage;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODET.OnElipsisClickE_General(Sender: TObject);
var lSt : String;
begin
  if VH^.PointageJal then
  begin
    lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X"';
    LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0 , lSt)
  end
  else
  begin
    if (ctxTreso in V_PGI.PGIContexte) then
      LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_NATUREGENE ="BQE" AND G_POINTABLE="X"', 'G_GENERAL', True, 0 )
    else
      LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_POINTABLE="X"', 'G_GENERAL', True, 0 );
  end;

  SetControlText('DATEPOINTAGE', DateToStr(iDate1900));
end;

{JP 28/07/05 : FQ 13731 : Pour éxecuter OnExitGeneral après l'éventuel chargement du filtre
{---------------------------------------------------------------------------------------}
procedure TOF_CPRAPPRODET.FormAfterShow;
{---------------------------------------------------------------------------------------}
begin
  if (GetControlText('FFILTRES') <> '') and not FLanceAutoOk then begin
    OnExitGeneral(gtxt_General);
    ChargeRefPointage;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRAPPRODET.ChargeRefPointage;
{---------------------------------------------------------------------------------------}
var
  lQuery : TQuery;
  lGene  : string;
begin
  {JP 21/08/07 : FQ 21256 : Pour éviter des messages d'erreur, on prend quelques précautions}
  if (gtxt_General.Text = '') or (gtxt_DatePointage.Text = '') then Exit;

  if EstPointageSurTreso then lGene := GetControlText('BQCODE')
                         else lGene := gtxt_General.Text;

  lQuery := OpenSQL('SELECT EE_ORIGINERELEVE, EE_REFPOINTAGE, EE_DATEPOINTAGE FROM ' + GetTableDossier(NomBase, 'EEXBQ') +
           ' WHERE EE_GENERAL = "' + lGene + '" AND EE_DATEPOINTAGE <= "' +
           USDateTime(StrToDate(gtxt_DatePointage.Text)) + '" ORDER BY EE_DATEPOINTAGE DESC',true);

  if (not lQuery.Eof) then begin
    StType := lQuery.FindField('EE_ORIGINERELEVE').AsString;
    glbl_Date.Caption := lQuery.FindField('EE_REFPOINTAGE').AsString;
  end
  else begin
    StType := 'RIN';
    glbl_Date.Caption := '';
  end;

  {JP17/04/07 : Nouvel gestion des sessions de pointage}
  ggbx_Sens.Visible := (not IsReleveAuto(StType));
  {16/04/07 : FQ 19610 : je pense qu'il n'y a pas de raison de le désactiver pour une
              session de pointage manuel, à voir ...
  if not IsReleveAuto(StType) then gcbx_Recap.Checked := False;
  gcbx_Recap.Enabled := IsReleveAuto(StType);}
  {18/09/07 : FQ 21331 : pas de gestion du récap !!!!}
  if (not lQuery.Eof) then
    gcbx_Recap.Enabled := lQuery.FindField('EE_DATEPOINTAGE').AsDateTime = StrToDate(gtxt_DatePointage.Text);
  if not gcbx_Recap.Enabled then gcbx_Recap.Checked := False; 

  Ferme(lQuery);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRAPPRODET.DatePointageOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  lStSql : string;
begin
  gtxt_DatePointage.Libelle := nil;

  {14/11/07 : Ajout du EE_STATUTRELEVE pour bien distinguer les relevés}
  if EstPointageSurTreso then
    lStSql := 'SELECT DISTINCT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ ' +
              'WHERE EE_GENERAL="' + GetControlText('BQCODE') + '" AND EE_STATUTRELEVE = "' + SESSIONTRESO + '"'
  else
    lStSql := 'SELECT DISTINCT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ ' +
              'WHERE EE_GENERAL="' + GetControlText('EGENERAL') + '" AND (EE_STATUTRELEVE <> "' +
              SESSIONTRESO + '" OR EE_STATUTRELEVE IS NULL)';
  {$IFDEF EAGLCLIENT}
  LookUpList(THEdit(Sender), 'Date de pointage', GetTableDossier(NomBase, 'EEXBQ'), 'EE_DATEPOINTAGE', '',
             '', 'EE_DATEPOINTAGE DESC', True, 0, lStSql);
  {$ELSE}
  lStSql := lStSql + ' ORDER BY EE_DATEPOINTAGE DESC';
  LookUpList(THEdit(Sender), 'Date de pointage', GetTableDossier(NomBase, 'EEXBQ'), 'EE_DATEPOINTAGE', '',
             '', '', True, 0, lStSql);
  {$ENDIF EAGLCLIENT}


  (*
  if EstPointageSurTreso then
    LookUpList(THEdit(Sender),'Date de pointage', GetTableDossier(NomBase, 'EEXBQ'), 'EE_DATEPOINTAGE', 'EE_REFPOINTAGE',
               'EE_GENERAL="' + GetControlText('BQCODE') + '"','EE_DATEPOINTAGE DESC', True, 0)
  else
    LookUpList(THEdit(Sender),'Date de pointage', GetTableDossier(NomBase, 'EEXBQ'), 'EE_DATEPOINTAGE', 'EE_REFPOINTAGE',
               'EE_GENERAL="' + GetControlText('EGENERAL') + '"','EE_DATEPOINTAGE DESC',True,0);
  *)
  ChargeRefPointage;
end;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOF_CPRAPPRODET.BqCodeOnElipsisClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LookupList(THEdit(Sender), 'Comptes bancaires', 'BANQUECP', 'BQ_CODE', 'BQ_LIBELLE', FiltreBanqueCp('', tcb_Bancaire, ''), 'BQ_CODE', True, 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPRAPPRODET.BqCodeOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT BQ_GENERAL, DOS_NOMBASE FROM BANQUECP ' +
               'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
               'WHERE BQ_CODE = "' + GetControlText('BQCODE') + '"', True);
  if not Q.EOF then begin
    NomBase := Q.FindField('DOS_NOMBASE').AsString;
    SetControlText('EGENERAL', Q.FindField('BQ_GENERAL').AsString);
  end;
  Ferme(Q);
  OnExitGeneral(THEdit(GetControl('EGENERAL')));
end;
{$ENDIF TRESO}

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPRAPPRODET]);

end.

