unit uLibRevision;

interface

uses

{$IFDEF EAGLCLIENT}
  MainEagl,
{$ELSE}
  DB,
  Fe_Main,       // AGLLanceFiche
  {$IFDEF DBXPRESS}uDbxDataSet,{$ELSE}dbtables,{$ENDIF}
{$ENDIF}

{$IFDEF EAGLSERVER}
{$ELSE}
  Ed_Tools,     // MoveCurProgressForm
  HMsgBox,      // PgiAsk
  Forms,        // Application
{$ENDIF}
  Printers,     // Printer
  Controls,     // MrYes
  SysUtils,     // Except
  HEnt1,        // BeginTrans
  HCtrls,       // OpensSQL
  Ent1,         // VH^.
  uTob,         // TOB
  UentCommun,  
  ParamSoc;     // GetParamSocSecur

const cATraiter  = 'ATR';
      cEnCours   = 'ENC';
      cEnSuspens = 'SUS';
      cSupervise = 'SUP';
      cValide    = 'VAL';

      cActif       = 'ACT';
      cNonActif    = 'NAC';
      cEtatinitial = 'INI';

      // Attestation
      cEnAttente   = 'ATT';
      cAvecObs     = 'AVC';
      cSansObs     = 'SNS';
      cRefus       = 'RFS';

      cImgBlocNote = 77;

type TCycleInterval =
     record
       Borne1Deb : string;
       Borne1Fin : string;
       Borne2Deb : string;
       Borne2Fin : string;
       AvecExclusion : Boolean;
     end ;

     TRic = class
     private
       FStPlanRevision : String;
       FLastErrorMsg   : String;

       FBoProgressForm : Boolean;
       FTobCycle       : Tob;
       FLastError      : Integer;

       procedure _LiensOLEVersCRevBlocNote; //
       procedure _ChargeLesCycles;          // Chargement des enreg de CREVPARAMCYCLE
       procedure _RecopieLesCycles;         // Copie de CREVPARAMCYCLE dans CREVCYCLE
       procedure _PrepareCRevGeneraux;      // Création des enregistrements du millésime
       procedure _AffecteLesCycles;         // Mise à jour du Champ G_CYCLEREVISION
       procedure _MajCRevGeneraux;
       procedure _CREVGenerauxVersGeneraux; // Recopie CRG_CYCLEREVISION,CRG_VISAREVISON dans GENERAUX
       procedure _AjouteCRevHistoCycle  ( const vStCycle, vStExercice, vStEtatCycle : string; vDtEtatCycleLe : TDateTime );

       function  _AutoriseActivation    ( vStActivation : string ) : Boolean;
       function  _MiseAJourVisaAuto     ( const vStCycle, vStExercice : string; vBoOk : Boolean ) : Boolean;


     public
       constructor Create;
       destructor  Destroy ; override;
       function    Affecte  ( const vStPlanRevision : String ) : Boolean;
       function    Supprime : Boolean;

       procedure ChangeActivationCycle ( const vStCycle, vStExercice, vStModeActivation : string );
       function  ChangeEtatCycle       ( const vStCycle, vStExercice, vStEtatCycle, vStUser : string ; const vDate : TDateTime ) : Boolean;
       function  ControleUnicite       ( const vStCycle, vStPredefini, vStListeCompte, vStListeExclusion, vStPlan : string; var vStCycleDoublon : string ): Boolean;

       property LastError    : integer read FLastError;
       property LastErrorMsg : string  read FLastErrorMsg;
       property ProgressForm : Boolean read FBoProgressForm write FBoprogressForm;

       // ------------ Fonctions d'impression des éditions de la révision ------
       // GCO - 02/08/2007 - Fonctions d'éditions du module Révision pour ne
       // corriger qu'une seule fois dans les sources
       class procedure ImpressionDossierTravail;
       class procedure CPLanceEtat_PageDeGarde;
       class function  BalanceAvecVariation         ( const vStCycle : string; vNumPage : integer ) : integer;
       class function  LanceEdition                 ( const vStNatureEtat, vStCodeEtat, vStWhere, vStRangeCritere : string; vNumPage: integer ) : integer;
       class function  CPLanceEtat_NoteSupervision  ( const vNumPage : integer = 0 ) : integer;
       class function  CPLanceEtat_ProgrammeTravail ( const vNumPage : integer = 0 ) : integer;
       class function  CPLanceEtat_NoteCtrlCompte   ( const vStGeneral : HString; vNumPage : integer = 0 ) : integer;
       class function  CPLanceEtat_NoteCtrlCycle    ( const vStCycle   : HString; vNumPage : integer = 0 ) : integer;
       class function  CPLanceEtat_NoteTravail      ( const vStGeneral, vStExercice : HString; vNumPage : integer = 0 ) : integer;
       class function  CPLanceEtat_TabVariation     ( const vStGeneral, vStExercice : HString; vNumPage : integer = 0 ) : integer;
       // ----------------------------------------------------------------------

     end;

////////////////////////////////////////////////////////////////////////////////
{$IFDEF EAGLSERVER}
{$ELSE}
  procedure AccesBStatutRevision;
  procedure ControlePlanRevision; // Point d'entrée de la RIC dans MDISP
  function  AlignementPlanRevision( const vStPlanRevision : string; vBoProgressForm : Boolean = False; vBoRazTotal : Boolean = False ) : Boolean;
  function  SupprimeRevision( vBoProgressForm : Boolean = False ) : Boolean;
  function  ControleConditionActivation : HString;
  procedure SynchroRICAvecExercice;
{$ENDIF}

// Fonctions de mise à jour des informations de la révision
function AutoriseSuppresionVisaRevision ( vStGeneral : string ) : Boolean;
function MiseAJourG_VisaRevision        ( vStGeneral : string; vBoVisaRevision : Boolean ) : Boolean;
function MiseAJourCREVGeneraux          ( vStGeneral : string; vBoVisaAuto : Boolean = False ) : Boolean;

procedure ReCalculInfoRevision;
procedure TraiteToutLeDossier( vStTraitement : string );

// Fonctions de recheche dans les cycles
function  TrouveCycleRevisionDuGene      ( vStGeneral : string ) : string;
function  TrouveGeneralSurCRevParamCycle ( vStGeneral, vStFamille : string ) : string;
////////////////////////////////////////////////////////////////////////////////

//procedure FQ14038;

implementation

uses
{$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
{$ENDIF MODENT1}

{$IFDEF EAGLCLIENT}
     UtileAGL,     // LanceEtat
{$ELSE}
     EdtREtat,     // LanceEtat
     uPdfBatch,    // StartPdfBatch
{$ENDIF}
     HPdfPrev,     // PreviewPDFFile
     uPrintF1Book, // PDFBatch
     uLibWindows;  // IIF

const cErrPlanRevision = 1;
      cMsgPlanRevision = 'Le plan de révision du dossier n''est pas renseigné.';

      cErrChargeLesCycles  = 2;
      CMsgChargeLesCycles  = 'Impossible de trouver les cycles du plan de révision.';

      cErrRecopieLesCycles = 3;
      cMsgRecopieLesCycles = 'Erreur de recopie dans la table CREVCYCLE';

      cErrAffecteLesCycles = 4;
      cMsgAffecteLesCycles = 'Erreur lors de l''affection des cycles dans la table GENERAUX';

      cErrMajCRevGeneraux  = 5;
      cMsgMajCRevGeneraux  = 'Erreur lors de l''affection des cycles dans la table CREVGENERAUX';

      cErrPrepareCRevGeneraux = 6;
      cMajPrepareCRevGeneraux = 'Erreur lors de l''insertion du millésimé dans la table CREVGENERAUX';

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure AccesBStatutRevision;
begin
  VH^.BStatutRevision.Visible := VH^.OkModRic and (VH^.Revision.Plan <> '');
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/09/2005
Modifié le ... : 23/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
procedure ControlePlanRevision;
var lQuery    : TQuery;
    lStResult : HString;
    lStAffiche: HString;
    lRic      : TRic;
begin
  if VH^.Revision.Plan = '' then Exit;

  try
    lStResult  := '';
    lStAffiche := '';

    lQuery := OpenSQL('SELECT CPR_DATEMODIF FROM CREVPLAN WHERE ' +
                      'CPR_PLANREVISION = "' + VH^.Revision.Plan + '"', True);

    if not lQuery.Eof then
    begin
      if IsValidDateHeure(lQuery.FindField('CPR_DATEMODIF').AsString) then
      begin
        if StrToDateTime(GetParamSocSecur('SO_CPDATEMAJPLANREVISION', iDate1900, True)) <
           StrToDateTime(lQuery.FindField('CPR_DATEMODIF').AsString) then
        begin
          if PgiAsk('Le paramétrage du plan de révision ' + VH^.Revision.Plan +
                    ' a été modifié. L''alignement du dossier entrainera ' + #13#10 +
                    'la perte des informations de révision. Voulez-vous effectuer ' +
                    'le traitement ?', 'Mise à jour de la révision') = MrYes then
          begin
            AlignementPlanRevision( VH^.Revision.Plan, True );
          end;
        end
        else
        begin
          lStResult := ControleConditionActivation;
          if (lStResult <> '') then
          begin
            lStAffiche := FindEtReplace(lStResult, ';', ', ', True);
            lStAffiche := '(' + Copy(lStAffiche, 0, Length(lStAffiche)-2) + ')';
            if PgiAsk('L''activation des cycles ' + lStAffiche + ' est remise en cause ' +
                      'par leurs conditions d''activation.' + #13#10 +
                      ' Souhaitez-vous appliquer ces changements sur le dossier ?',
                      'Mise à jour de la révision') = MrYes then
            begin
              lRic := TRic.Create;
              while Pos(';', lStResult) > 0 do
              begin
                lRic.ChangeActivationCycle(ReadTokenSt(lStResult),
                                           GetEnCours.Code, cEtatInitial );
              end;
              lRic.Free;
            end;
          end;
        end;
      end;
    end
    else
    begin // Le Plan de révision du PARAMSOC n'existe plus dans la table
      if PgiAsk('Le plan de révision associé au dossier n''existe plus. ' + #13#10 +
                ' Voulez-vous effacer les informations de révision ?', 'Mise à jour de la révision') = MrYes then
      begin
        SupprimeRevision(True);
      end;
    end;
  finally
    Ferme(lQuery);
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
function AlignementPlanRevision( const vStPlanRevision : string; vBoProgressForm : Boolean = False; vBoRazTotal : Boolean = False ) : Boolean;
var lRic : TRic;
begin
  Result := True;
  lRic := nil;
  try
    lRic := TRic.Create;
    lRic.ProgressForm := vBoProgressForm;

    // GCO - 13/07/2007 - FQ 21080 - Suppression totale de la révision si
    // on remplace le plan de révision actuel par un autre
    if vBoRazTotal then
    begin
      if not lRic.Supprime then
      begin
        PgiInfo( 'Traitement annulé. ' + lRic.LastErrorMsg, 'Mise à jour de la révision');
        Result := False;
        Exit;
      end;
    end;

    if not lRic.Affecte( vStPlanRevision ) then
    begin
      PgiInfo( 'Traitement annulé. ' + lRic.LastErrorMsg, 'Mise à jour de la révision');
      Result := False;
    end;
  finally
    lRic.Free;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{$IFNDEF EAGLSERVER}
function SupprimeRevision( vBoProgressForm : Boolean = False ) : Boolean;
var lRic : TRic;
begin
  Result := True;
  lRic := TRic.Create;
  try
    lRic.ProgressForm := vBoProgressForm;
    if not lRic.Supprime then
    begin
      Result := False;
      PgiInfo( 'Traiment annulé. ' + lRic.FLastErrorMsg, 'Suppresion des informations de la révision');
    end;
  finally
    lRic.Free;
  end;
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function MiseAJourCREVGeneraux( vStGeneral : string; vBoVisaAuto : Boolean = False ) : Boolean;
var lQuery : TQuery;
    lBoVisaRevision : Boolean;
begin
  Result := False;
  lQuery := nil;
  try
    lQuery := OpenSql('SELECT G_VISAREVISION, G_CYCLEREVISION FROM GENERAUX ' +
                      'WHERE G_GENERAL = "' + vStGeneral + '"', True);
    if not lQuery.Eof then
    begin
      lBoVisaRevision := lQuery.FindField('G_VISAREVISION').AsString = 'X';

      if ExisteSQL('SELECT CRG_GENERAL FROM CREVGENERAUX WHERE ' +
                   'CRG_GENERAL = "' + vStGeneral + '" AND ' +
                   'CRG_EXERCICE = "' + VH^.Encours.Code + '"') then
      begin
        if ExecuteSQL('UPDATE CREVGENERAUX SET ' +
                      'CRG_CYCLEREVISION = "' + lQuery.FindField('G_CYCLEREVISION').AsString + '", ' +
                      'CRG_VISAREVISION = "' + IIF(lBoVisaRevision, 'X', '-') + '",' +
                      'CRG_VISAPAR = "' + IIF(lBoVisaRevision and (not vBoVisaAuto), V_Pgi.User, '') + '",' +
                      'CRG_VISALE = "' + UsDateTime(IIF(lBoVisaRevision and (not vBoVisaAuto), Now, idate1900)) + '" WHERE ' +
                      'CRG_GENERAL = "' + vStGeneral + '" AND ' +
                      'CRG_EXERCICE = "' + VH^.EnCours.Code + '"') = 1 then
          Result := True;
      end
      else
      begin
        if ExecuteSQL('INSERT INTO CREVGENERAUX (' +
                      'CRG_GENERAL, CRG_EXERCICE, CRG_VISAREVISION, CRG_VISAPAR, ' +
                      'CRG_VISALE, CRG_CYCLEREVISION, CRG_VISAAUTO) VALUES (' +
                      '"' + vStGeneral + '", ' +
                      '"' + VH^.Encours.Code + '", ' +
                      '"' + IIF(lBoVisaRevision, 'X', '-') + '", ' +
                      '"' + IIF(lBoVisaRevision and (not vBoVisaAuto), V_Pgi.User, '') + '",' +
                      '"' + UsDateTime(IIF(lBoVisaRevision and (not vBoVisaAuto), Now, idate1900)) + '", ' +
                      '"' + lQuery.FindField('G_CYCLEREVISION').AsString + '",'+
                      '"' + IIF(lBoVisaRevision and vBoVisaAuto, 'X', '-') + '")') = 1 then
          Result := True;
      end;
    end
    else
      Result := False;

  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2007
Modifié le ... :   /  /    
Description .. : Pour l'instant dans IFDEF COMPTA car je ne sais pas comment
Suite ........ : ça se passe si affichage de l'écran depuis le BUREAU PGI
Mots clefs ... : 
*****************************************************************}
procedure ReCalculInfoRevision;
{$IFDEF COMPTA}
var lQuery : TQuery;
    lTotalATraiter  : integer;
    lTotalEnSuspens : integer;
    lTotalValide    : integer;
    lTotalSupervise : integer;
{$ENDIF}
begin
{$IFDEF COMPTA}
  lQuery          := nil;
  lTotalATraiter  := 0;
  lTotalEnSuspens := 0;
  lTotalValide    := 0;
  lTotalSupervise := 0;

  try
    lQuery := OpenSQL('SELECT COUNT(*) TOTAL, CCY_ETATCYCLE FROM CREVCYCLE WHERE ' +
                      'CCY_ACTIVECYCLE = "X" AND ' +
                      'CCY_EXERCICE = "' + VH^.EnCours.Code + '" GROUP BY CCY_ETATCYCLE', True);

    while not lQuery.Eof do
    begin
      if ((lQuery.Findfield('CCY_ETATCYCLE').AsString = cATraiter) or
          (lQuery.Findfield('CCY_ETATCYCLE').AsString = cEnCours)) then
        // GCO - 06/06/2007 - FQ 20556
        lTotalATraiter := lTotalATraiter + lQuery.FindField('TOTAL').AsInteger
      else
        if lQuery.Findfield('CCY_ETATCYCLE').AsString = cEnSuspens then
          lTotalEnSuspens := lQuery.FindField('TOTAL').AsInteger
        else
          if lQuery.Findfield('CCY_ETATCYCLE').AsString = cValide then
            lTotalValide := lQuery.FindField('TOTAL').AsInteger
          else
            if lQuery.Findfield('CCY_ETATCYCLE').AsString = cSupervise then
              lTotalSupervise := lQuery.FindField('TOTAL').AsInteger;
      lQuery.Next;
    end;

    if ExisteSQL('SELECT CIR_NODOSSIER FROM CREVINFODOSSIER WHERE ' +
                 'CIR_NODOSSIER = "' + V_Pgi.NoDossier + '" AND ' +
                 'CIR_EXERCICE = "' + VH^.EnCours.Code + '"') then
    begin
      ExecuteSQL('UPDATE CREVINFODOSSIER SET ' +
                 'CIR_NBATRAITER = ' + IntToStr(lTotalATraiter) + ', ' +
                 'CIR_NBENSUSPENS = ' + IntToStr(lTotalEnSuspens) + ', ' +
                 'CIR_NBVALIDE = ' + IntToStr(lTotalValide) + ', ' +
                 'CIR_NBSUPERVISE =' + IntToStr(lTotalSupervise) + ' WHERE ' +
                 'CIR_NODOSSIER = "' + V_Pgi.NoDossier + '" AND ' +
                 'CIR_EXERCICE = "' + VH^.EnCours.Code + '"');
    end;

  finally
    Ferme( lQuery );
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TraiteToutLeDossier( vStTraitement : string );
{$IFDEF COMPTA}
var lQuery : TQuery;
    lStWhere : string;
    lRic : TRic;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if (vStTraitement <> cATraiter) and (vStTraitement <> cEnSuspens) and
     (vStTraitement <> cSupervise) and (vStTraitement <> cValide) then Exit;

  lStWhere := '';
  lRic := TRic.Create;
  try
    // Validation de tous les cycles sauf les déja validés, ou supervisés
    if vStTraitement = cValide then
      lStWhere := ' AND CCY_ETATCYCLE <> "' + cValide + '" AND ' +
                  'CCY_ETATCYCLE <> "' + cSupervise + '" '
    else
      if vStTraitement = cSupervise then
        lStWhere := ' AND CCY_ETATCYCLE <> "' + cSupervise + '" ';

    lQuery := OpenSQL('SELECT CCY_CODECYCLE FROM CREVCYCLE WHERE ' +
                      'CCY_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
                      'CCY_ACTIVECYCLE = "X" ' + lStWhere + ' ORDER BY CCY_CODECYCLE', True);

    while not lQuery.Eof do
    begin
      lRic.ChangeEtatCycle( lQuery.FindField('CCY_CODECYCLE').AsString,
                                       VH^.EnCours.Code, vStTraitement, V_Pgi.User, Date );
      lQuery.Next;
    end;

  finally
    lRic.Free;
    Ferme( lQuery );
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function AutoriseSuppresionVisaRevision( vStGeneral : string ) : Boolean;
var lStCycle  : string;
    lStEtatCycle : string;
begin
  Result := False;

  if not VH^.OkModRIC then
  begin // Si pas de séria RIC, le visa peut être enlevé à tout moment.
    Result := True;
    Exit;
  end;

  lStCycle := GetColonneSQL('GENERAUX', 'G_CYCLEREVISION', 'G_GENERAL = "' + vStGeneral + '"');
  if lStCycle = '' then
  begin
    Result := True;
    Exit;
  end;


  lStEtatCycle := GetColonneSql('CREVCYCLE', 'CCY_ETATCYCLE',
    'CCY_CODECYCLE = "' + lStCycle + '" AND CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

  if (lStEtatCycle <> 'VAL') and (lStEtatCycle <> 'SUP') then
  begin
    Result := True;
    Exit;
  end;

  // GCO - 03/05/2007 - FQ 20172
  // Plus de passe droit pour N3,N4 pour enlever le visa si cValide ou cSupervise
  //if JaileRoleCompta(rcSuperviseur) then
  //  Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/01/2007
Modifié le ... : 08/02/2007
Description .. :
Mots clefs ... :
*****************************************************************}
function MiseAJourG_VisaRevision( vStGeneral : string; vBoVisaRevision : Boolean) : Boolean;
var lStVisaRevision : string;
begin
  Result := False;
  BeginTrans;
  try
    if vBoVisaRevision then
      lStvisaRevision := 'X'
    else
      lStvisaRevision := '-';

    ExecuteSql('UPDATE GENERAUX SET G_VISAREVISION = "' + lStVisaRevision + '" ' +
               'WHERE G_GENERAL = "' + vStGeneral + '"' );

    if not MiseAJourCREVGeneraux( vStGeneral ) then
      raise Exception.Create('');

    CommitTrans;
    Result := True;
  except
    on E : Exception do RollBack;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/02/2005
Modifié le ... : 23/01/2007
Description .. :
Suite ........ : GCO - 24/08/2005 - FQ 16423 Ajout de condition sur RB_NODOSSIER
Mots clefs ... :
*****************************************************************}
function TrouveCycleRevisionDuGene( vStGeneral : string) : string;
begin
  Result := TrouveGeneralSurCRevParamcycle( vStGeneral, VH^.Revision.Plan);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/03/2005
Modifié le ... : 29/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
function TrouveGeneralSurCRevParamCycle( vStGeneral, vStFamille : string ) : string;
var lQuery : TQuery;
    lStSql             : string;
    lStWhereCompte1    : string;
    lStWhereExclusion1 : string;
begin
  Result := '';

  if vStFamille = '' then Exit;
  if vStGeneral = '' then Exit;

  try
    try
      lStSql := 'SELECT CPC_CODECYCLE, CPC_LISTECOMPTE, CPC_LISTEEXCLUSION ' +
                'FROM CREVPARAMCYCLE ' +
                'LEFT JOIN CREVCYCLE ON CCY_CODECYCLE = CPC_CODECYCLE AND ' +
                'CCY_PREDEFINI = CPC_PREDEFINI WHERE ' +
                'CPC_PLANASSOCIE LIKE "%' + vStFamille + ';%" AND ' +
                'CCY_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
                'CCY_ACTIVECYCLE = "X" ORDER BY CPC_CODECYCLE';

      lQuery := OpenSql( lStSql, True );
      while not lQuery.Eof do
      begin
        lStWhereCompte1    := lQuery.FindField('CPC_LISTECOMPTE').AsString;
        lStWhereExclusion1 := lQuery.FindField('CPC_LISTEEXCLUSION').AsString;

        lStWhereCompte1    := AnalyseCompte( lStWhereCompte1 , fbGene, False, False) ;
        lStWhereExclusion1 := AnalyseCompte( lStWhereExclusion1, fbGene, True, False);

        // Cycle avec fourchettes de comptes vides, pas besoin de rechercher dedans
        if (lStWhereCompte1 <> '') then
        begin
          lStSql := 'SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + vStGeneral + '"';

          if lStWhereCompte1 <> '' then
            lStSql := lStSql + ' AND ' + lStWhereCompte1;

          if lStWhereExclusion1 <> '' then
            lStSql := lStSql + ' AND ' + lStWhereExclusion1;

          if ExisteSQL( lStSql ) then
          begin
            Result := lQuery.FindField('CPC_CODECYCLE').AsString;
            Break;
          end;
        end;
        lQuery.Next;
      end;

    except
      on E: Exception do
      begin
        Result := '';
      end;
    end;

  finally
    Ferme( lQuery );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{ TRic }
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... : 29/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TRic.Create;
begin
  FTobCycle    := Tob.Create('TOBCYCLE', nil, -1);
  FLastError      := 0;
  FLastErrorMsg   := '';
  FBoProgressForm := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... : 29/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TRic.Destroy;
begin
  FreeAndNil(FTobCycle);

  // Recalcul des cumuls des différents états des cycles de CREVINFODOSSIER
  ReCalculInfoRevision;

  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... : 05/02/2007
Description .. :
Mots clefs ... :
*****************************************************************}
function TRic.Affecte( const vStPlanRevision : String ) : Boolean;
begin
  Result := True;
  try
    if vStPlanRevision <> '' then
      FStPlanRevision := vStPlanRevision
    else
      FStPlanRevision := VH^.Revision.Plan;

    if FStPlanRevision = '' then
    begin
      FLastError := cErrPlanRevision;
      Exit;
    end;

    // Traitement de pré-deversement pour PGE
    // PREDEFINI = DOS deviennent STD
    if not (CtxPcl in V_Pgi.PgiContexte) then
    begin
      ExecuteSQL('UPDATE CREVPLAN SET CPR_PREDEFINI = "STD" WHERE CPR_PREDEFINI = "DOS"');
      ExecuteSQL('UPDATE CREVPARAMCYCLE SET CPC_PREDEFINI = "STD" WHERE CPC_PREDEFINI = "DOS"');
    end;

    // Si présence d'ancien commentaire millésimé, on les transfert dans CREVBLOCNOTE
    if ExisteSQL('SELECT LO_IDENTIFIANT FROM LIENSOLE WHERE LO_TABLEBLOB = "G"') then
      _LiensOLEVersCRevBlocNote;

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then
      InitMoveProgressForm(nil, 'Alignement du plan de révision ' + FStPlanRevision,
      'Mise à jour des informations de révision en cours...', 3 , True, True);
  {$ENDIF}

    BeginTrans;
    // Sélection des cycles de CREVPARAMCYCLE
    {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Chargement des cycles...');
    {$ENDIF}
    // Sélection des cycles du PLAN dans DB000000.DBO.CREVPARAMCYCLE
    _ChargeLesCycles;
    if FLastError <> 0 then Exit;

    {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Copie des cycles...');
    {$ENDIF}
    // Insertion des cycles dans DBXXXXXX.DBO.CREVCYCLE
    _RecopieLesCycles;
    if FLastError <> 0 then Exit;

    // Création des enregistrements du Millésime dans CREVGENERAUX
    _PrepareCRevGeneraux;
    if FLastError <> 0 then Exit;

    {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Mise à jour des comptes généraux...');
    {$ENDIF}
    // Affectation du champ G_CYCLEREVISON dans GENERAUX 
    _AffecteLesCycles;
    if FLastError <> 0 then Exit;

    // Mise à jour de GENERAUX vers CREVGENERAUX
    // G_CYCLEREVISION -> CRG_CYCLEREVISION
    // G_VISAREVISION  -> CRG_VISAREVISION
    _MAJCRevGeneraux;
    if FLastError <> 0 then Exit;

    // GCO - 04/10/2007 - Suppression des blocnotes n'ayant pas de cycles existant
    // dans CREVCYCLE
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE NOT EXISTS (SELECT CCY_CODECYCLE ' +
               'FROM CREVCYCLE WHERE CCY_CODECYCLE = CBN_CODE AND ' +
               'CCY_EXERCICE = CBN_EXERCICE)');

  finally
    if FLastError = 0 then
    begin
      Result := True;
      // Enregistrement du plan dans le PARAMSOC
      SetParamSoc('SO_CPPLANREVISION', FStPlanRevision);
      VH^.Revision.Plan := FStPlanRevision;

      // Mise à jour de la date de déversement dans le dossier
      SetParamSoc('SO_CPDATEMAJPLANREVISION', Now);

      CommitTrans;
    end
    else
    begin
      Result := False;

      case FLastError of
        cErrPlanRevision     : FLastErrorMsg := CMsgPlanRevision;
        cErrChargeLesCycles  : FLastErrorMsg := cMsgChargeLesCycles;
        cErrRecopieLesCycles : FLastErrorMsg := cMsgRecopieLesCycles;
        cErrAffecteLesCycles : FLastErrorMsg := cMsgAffecteLesCycles;
        cErrMajCRevGeneraux  : FLastErrorMsg := cMsgMajCRevGeneraux;
      else
      end;

      RollBack;
    end;

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then FiniMoveProgressForm;
  {$ENDIF}
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2007
Modifié le ... :   /  /
Description .. : Recopie des enregitrements de type "G" qui représentent les
Suite ........ : commentaires millésimés vers la table CREVBLOCNOTE.
Mots clefs ... : 
*****************************************************************}
procedure TRic._LiensOLEVersCRevBlocNote;
begin
  Exit;
  ExecuteSQL('INSERT INTO CREVBLOCNOTE ' +
             '(CBN_NATURE, CBN_CODE, CBN_EXERCICE, CBN_BLOCNOTE, ' +
             'CBN_DATECREATION, CBN_DATEMODIF, CBN_CREATEUR, CBN_UTILISATEUR) ' +
             '(SELECT "GEN", LO_IDENTIFIANT, LO_EMPLOIBLOB, LO_OBJET, ' +
             'LO_DATECREATION, LO_DATEMODIF, LO_CREATEUR, LO_UTILISATEUR ' +
             'FROM LIENSOLE WHERE LO_TABLEBLOB = "G")');

  ExecuteSQL('DELETE FROM LIENSOLE WHERE LO_TABLEBLOB = "G"');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TRic._ChargeLesCycles;
var lStSql : string;
begin
  lStSql := 'SELECT CPC_CODECYCLE, CPC_LIBELLECYCLE, CPC_LISTECOMPTE, ' +
            'CPC_LISTEEXCLUSION, CPC_ACTIVATIONSQL, CPC_PREDEFINI ' +
            'FROM CREVPARAMCYCLE WHERE ' +
            'CPC_PLANASSOCIE LIKE "%' + FStPlanRevision + ';%" AND ' +
            '((CPC_PREDEFINI = "CEG") OR (CPC_PREDEFINI = "STD")) ' +
            'ORDER BY CPC_CODECYCLE';

  FTobCycle.LoadDetailFromSql(lStSql, False, True);

  if FTobCycle.Detail.Count = 0 then
    FLastError := cErrChargeLesCycles;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... : 29/01/2007
Description .. : Recopie des enreg dans CREVCYCLE
Mots clefs ... :
*****************************************************************}
procedure TRic._RecopieLesCycles;
var i             : integer;
    lTob          : Tob;
    lBoActivation : Boolean;
begin
  try
    ExecuteSQL('DELETE FROM CREVCYCLE WHERE CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

    for i := 0 to FTobCycle.Detail.Count - 1 do
    begin
      lTob := FTobCycle.Detail[i];

      lBoActivation := _AutoriseActivation(lTob.GetString('CPC_ACTIVATIONSQL'));

      // Insertion de l'enregistrement dans CREVCYCLE
      ExecuteSQL('INSERT INTO CREVCYCLE (CCY_CODECYCLE, CCY_EXERCICE, ' +
                 'CCY_LIBELLECYCLE, CCY_ETATCYCLE, CCY_ETATCYCLEPAR, ' +
                 'CCY_ETATCYCLELE, CCY_ACTIVECYCLE, CCY_PREDEFINI) VALUES (' +
                 '"' + lTob.GetString('CPC_CODECYCLE') + '", ' +
                 '"' + VH^.EnCours.Code + '",' +
                 '"' + lTob.GetString('CPC_LIBELLECYCLE') + '", ' +
                 '"' + IIF(lBoActivation, cATraiter, cNonActif) + '", ' +
                 '"", "01/01/1900", ' +
                 IIF(lBoActivation, '"X"', '"-"')+ ', ' +
                 '"' + lTob.GetString('CPC_PREDEFINI') + '")');

      // Changemennt de son état qui automatise tout les traitements à effectuer
      ChangeEtatCycle( lTob.GetString('CPC_CODECYCLE'), VH^.EnCours.Code,
                       IIF( lBoActivation, cATraiter, cNonActif) , V_Pgi.User, Date);

    end;
  except
    on E : Exception do FLastError := cErrRecopieLesCycles;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/07/2007
Modifié le ... :   /  /    
Description .. : Préparation de la table CREVGENERAUX, on insère les
Suite ........ : enregistrements qui n'existent pas pour l'exercice
Mots clefs ... :
*****************************************************************}
procedure TRic._PrepareCRevGeneraux;
var lQuery : TQuery;
begin
  lQuery := nil;
  try
    try
      lQuery := OpenSQL('SELECT G_GENERAL, G_CYCLEREVISION FROM GENERAUX ' +
                      'LEFT JOIN CREVGENERAUX ON CRG_GENERAL = G_GENERAL AND ' +
                      'CRG_EXERCICE = "' + VH^.EnCours.Code + '" WHERE ' +
                      'CRG_GENERAL IS NULL ORDER BY CRG_GENERAL', True);

      while not lQuery.Eof do
      begin
        // GCO - 31/07/2007 - FQ 21198 (ajout du CRG_VISAREVISION à l'insertion)
        ExecuteSQL('INSERT INTO CREVGENERAUX (CRG_GENERAL, CRG_EXERCICE, ' +
                   'CRG_VISAREVISION,' +
                   'CRG_VISAPAR, CRG_VISALE, CRG_VISAAUTO) VALUES ' +
                   '("' + lQuery.FindField('G_GENERAL').AsString + '", ' +
                   '"' + VH^.EnCours.Code + '", "-", ' +
                   '"", "'+ UsDateTime(idate1900)  + '", "-")');
        lQuery.Next;
      end;
    except
      on E : Exception do FLastError := cErrPrepareCRevGeneraux;
    end;
  finally
    Ferme(lQuery );
  end;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/01/2007
Modifié le ... : 29/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TRic._AffecteLesCycles;
var i      : integer;
    lTob   : Tob;
    lTob2  : Tob;
    lTobCycleActif : Tob;
    lStSql : string;
    lStWhereCompte : string;
    lStWhereExclusion : string;
begin
  // GCO - 23/07/2007 - FQ 21040
  ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = ""');

  lStSql := 'SELECT CCY_CODECYCLE FROM CREVCYCLE WHERE CCY_ACTIVECYCLE = "X" AND ' +
            'CCY_EXERCICE = "' + VH^.EnCours.Code + '" ORDER BY CCY_CODECYCLE';

  lTobCycleActif := Tob.Create('CYCLEACTIF', nil, -1);
  lTobCycleActif.LoadDetailFromSql(lStSql, False, True);

  for i := 0 to FTobCycle.Detail.Count - 1 do
  begin
    lTob := FTobCycle.Detail[i];

    lTob2 := lTobCycleActif.FindFirst(['CCY_CODECYCLE'], [lTob.GetString('CPC_CODECYCLE')], False);
    // Tob2 = nil si le cycle n'est pas présent dans la liste des cycles actifs
    // donc on passe au suivant
    if lTob2 = nil then Continue;

    lStSql := '';
    lStWhereCompte    := AnalyseCompte( lTob.GetString('CPC_LISTECOMPTE'), fbGene, False, False);
    lStWhereExclusion := AnalyseCompte( lTob.GetString('CPC_LISTEEXCLUSION'), fbGene, True, False);

    if lStWhereCompte <> '' then
    begin
      lStSql := 'UPDATE GENERAUX SET G_CYCLEREVISION = "' + lTob.GetString('CPC_CODECYCLE') + '" ' +
                'WHERE ' + lStWhereCompte;

      if lStWhereExclusion <> '' then
        lStSql := lStSql + ' AND ' + lStWhereExclusion ;

      ExecuteSql( lStSql );
    end;
  end;

  FreeAndNil(lTobCycleActif);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/01/2007
Modifié le ... : 31/01/2007
Description .. : Recopie G_CYCLEREVISON DANS CRG_CYCLEREVISION
Suite ........ : GCO - 13/06/2007 - FQ 20671
Suite ........ : GCO - 24/07/2007 - FQ 21040
Mots clefs ... :
*****************************************************************}
procedure TRic._MajCRevGeneraux;
begin // GCO - 13/06/2007 - FQ 20671 - PB traducteur AGL vers ORACLE
  ExecuteSQL('UPDATE CREVGENERAUX SET ' +
             'CRG_VISAAUTO = "-", ' +
             'CRG_VISAREVISION =  (SELECT G_VISAREVISION FROM GENERAUX ' +
             'WHERE G_GENERAL = CRG_GENERAL), ' +
             'CRG_CYCLEREVISION = (SELECT G_CYCLEREVISION FROM GENERAUX ' +
             'WHERE G_GENERAL = CRG_GENERAL) WHERE ' +
             'CRG_EXERCICE = "' + VH^.EnCours.Code + '"');
end;
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2007
Modifié le ... :   /  /
Description .. : Ajoute d'une ligne d'historique du CYCLE dans CREVHISTOCYCLE
Mots clefs ... :
*****************************************************************}
procedure TRic._AjouteCRevHistoCycle(const vStCycle, vStExercice, vStEtatCycle : string; vDtEtatCycleLe : TDateTime );
begin
  ExecuteSQL('INSERT INTO CREVHISTOCYCLE (' +
             'CHC_CODECYCLE, CHC_EXERCICE, CHC_DATECREATION, CHC_ETATCYCLE, ' +
             'CHC_ETATCYCLEPAR, CHC_ETATCYCLELE, CHC_CREATEUR) VALUES (' +
             '"' + vStCycle + '", ' +
             '"' + vStExercice + '", ' +
             '"' + UsDateTime( Now ) + '", ' +
             '"' + vStEtatCycle + '", ' +
             '"' + V_Pgi.User + '", ' +
             '"' + UsDateTime( vDtEtatCycleLE ) + '", ' +
             '"' + V_Pgi.User + '")');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/03/2007
Modifié le ... : 27/03/2007
Description .. :
Mots clefs ... :
*****************************************************************}
function TRic._AutoriseActivation( vStActivation : string ) : Boolean;
begin
  Result := True;
  if vStActivation = '' then Exit;

  vStActivation := FindEtReplace( vStActivation , '|', '', True);

  Result := ExisteSQL('SELECT DOS_NODOSSIER FROM DOSSIER ' +
                      'LEFT JOIN DPORGA ON DOR_GUIDPER = DOS_GUIDPER ' +
                      'LEFT JOIN DPSOCIAL ON DSO_GUIDPER = DOS_GUIDPER ' +
                      'LEFT JOIN DPFISCAL ON DFI_GUIDPER = DOS_GUIDPER ' +
                      'LEFT JOIN ANNUBIS ON ANB_GUIDPER = DOS_GUIDPER ' +
                      'LEFT JOIN ANNUAIRE ON ANN_GUIDPER = DOS_GUIDPER ' +
                      'LEFT JOIN TIERS ON T_TIERS = ANN_TIERS ' +
                      'LEFT JOIN TIERSCOMPL ON YTC_TIERS = T_TIERS ' +
                      'LEFT JOIN JURIDIQUE ON JUR_GUIDPERDOS = DOS_GUIDPER WHERE ' +
                      'DOS_NODOSSIER = "' + V_Pgi.NoDossier + '" AND ' +
                      vStActivation);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TRic._MiseAJourVisaAuto( const vStCycle, vStExercice: string; vBoOk: Boolean): Boolean;
begin
  Result := True;

  if vBoOk then
  begin // On pause un visa Automatique
    ExecuteSQL('UPDATE CREVGENERAUX SET ' +
               'CRG_VISAAUTO = "X", CRG_VISAREVISION = "X" WHERE ' +
               'CRG_CYCLEREVISION = "' + vStCycle + '" AND ' +
               'CRG_EXERCICE = "' + vStExercice + '" AND ' +
               'CRG_VISAREVISION = "-"');

    ExecuteSQL('UPDATE GENERAUX SET G_VISAREVISION = "X" WHERE ' +
               'G_CYCLEREVISION = "' + vStCycle + '" AND ' +
               'G_VISAREVISION = "-"');

    (*ExecuteSQL('UPDATE CREVGENERAUX SET ' +
               'CRG_VISAREVISION = (SELECT G_VISAREVISION FROM GENERAUX ' +
               'WHERE G_GENERAL = CRG_GENERAL) WHERE ' +
               'CRG_EXERCICE = "' + vStExercice + '" AND ' +
               'CRG_CYCLEREVISION = "' + vStCycle + '"');*)
  end
  else
  begin // Suppression du visa automatique
    ExecuteSQL('UPDATE CREVGENERAUX SET ' +
               'CRG_VISAAUTO = "-", ' +
               'CRG_VISAREVISION = "-" WHERE ' +
               'CRG_CYCLEREVISION = "' + vStCycle + '" AND ' +
               'CRG_EXERCICE = "' + vStExercice + '" AND ' +
               'CRG_VISAAUTO = "X"');

    ExecuteSQL('UPDATE GENERAUX SET ' +
              'G_VISAREVISION = (SELECT CRG_VISAREVISION FROM CREVGENERAUX WHERE ' +
              'CRG_GENERAL = G_GENERAL AND ' +
              'CRG_EXERCICE = "' + vStExercice + '") WHERE ' +
              'G_CYCLEREVISION = "' + vStCycle + '"');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2007
Modifié le ... :   /  /    
Description .. : Recopie des champs CRG_VISAREVISION VERS
Mots clefs ... : 
*****************************************************************}
procedure TRic._CREVGenerauxVersGeneraux;
begin
  ExecuteSQL('UPDATE GENERAUX SET '+
    'G_VISAREVISION = (SELECT CRG_VISAREVISION FROM CREVGENERAUX WHERE ' +
    'CRG_GENERAL = G_GENERAL AND CRG_EXERCICE = "' + VH^.EnCours.Code + '"), ' +
    'G_CYCLEREVISION = (SELECT CRG_CYCLEREVISION FROM CREVGENERAUX WHERE ' +
    'CRG_GENERAL = G_GENERAL AND CRG_EXERCICE = "' + VH^.EnCours.Code + '") ' +
    'WHERE EXISTS (SELECT CRG_GENERAL FROM CREVGENERAUX WHERE ' +
    'CRG_GENERAL = G_GENERAL AND CRG_EXERCICE = "' + VH^.Encours.Code + '")');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TRic.Supprime : Boolean;
begin
  Result := True;
  try
  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then
      InitMoveProgressForm(nil, 'Suppression du plan de révision ' + FStPlanRevision,
      'Effacement des informations de révision en cours...', 5 , True, True);
  {$ENDIF}

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Suppression des cycles dans les comptes généraux...');
  {$ENDIF}

    // GCO - 13/07/2007 - FQ 21040
    // Suppression de tous les visas automatiques
    ExecuteSQL('UPDATE CREVGENERAUX SET ' +
               'CRG_VISAAUTO = "-", ' +
               'CRG_VISAREVISION = "-" WHERE ' +
               'CRG_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
               'CRG_VISAAUTO = "X"');

    // Suppression de l'affectation des cycles dans CREVGENERAUX
    ExecuteSQL('UPDATE CREVGENERAUX SET CRG_CYCLEREVISION = "" WHERE ' +
               'CRG_EXERCICE = "' + VH^.EnCours.Code + '"');

    ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = "", ' +
              'G_VISAREVISION = (SELECT CRG_VISAREVISION FROM CREVGENERAUX WHERE ' +
              'CRG_GENERAL = G_GENERAL AND ' +
              'CRG_EXERCICE = "' + VH^.EnCours.Code + '")');

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Suppression des cycles de révision...');
  {$ENDIF}
    ExecuteSQL('DELETE FROM CREVCYCLE WHERE CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Suppression des commentaires de révision...');
  {$ENDIF}
    // Efface tous les bloc-notes sauf les commentaires millésimés...
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE ' +
               'CBN_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
               'CBN_NATURE <> "GEN"');

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Suppression de l''historique...');
  {$ENDIF}
    ExecuteSQL('DELETE FROM CREVHISTOCYCLE WHERE ' +
               'CHC_EXERCICE = "' + VH^.EnCours.Code + '"');

  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then MoveCurProgressForm('Mise à jour des paramétres société...');
  {$ENDIF}
    SetParamSoc('SO_CPPLANREVISION', '');
    VH^.Revision.Plan := '';

    SetParamSoc('SO_CPREVISBLOQUEGENE', '-');
    SetParamSoc('SO_CPDATEMAJPLANREVISION', idate1900);
    
  finally
  {$IFNDEF EAGLSERVER}
    if FBoProgressForm then FiniMoveProgressForm;
  {$ENDIF}
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : GCO - 11/07/2007 - FQ 20954
Mots clefs ... :
*****************************************************************}
procedure TRic.ChangeActivationCycle( const vStCycle, vStExercice, vStModeActivation: string );
var lBoOkActive       : Boolean;
    lStPredefini      : string;
    lStActivation     : string;
    lStSql            : string;
    lStWhereCompte    : string;
    lStWhereExclusion : string;
    lTobParamCycle    : Tob;
    lQuery            : TQuery;
begin

  lStPredefini  := '';
  lStActivation := '';

  lQuery := OpenSQL('SELECT CCY_ACTIVECYCLE, CCY_PREDEFINI FROM CREVCYCLE WHERE ' +
                    'CCY_CODECYCLE = "' + vStCycle + '" AND ' +
                    'CCY_EXERCICE = "' + vStExercice + '"', True);

  if not lQuery.Eof then
  begin
    lStPredefini  := lQuery.FindField('CCY_PREDEFINI').AsString;
    lStActivation := lQuery.FindField('CCY_ACTIVECYCLE').AsString;
  end;

  Ferme( lQuery );

  if ((vStModeActivation = cActif) and (lStActivation = 'X')) or
     ((vStModeActivation = cNonActif) and (lStActivation = '-')) then
    Exit; // Pas besoin de traitement, on redemande de remettre dans le même état

  lTobParamCycle := Tob.Create('CREVPARAMCYCLE', nil, -1 );
  lTobParamCycle.SetString('CPC_CODECYCLE', vStCycle);
  lTobParamCycle.SetString('CPC_PREDEFINI', lStPredefini);
  lTobParamCycle.LoadDB(True);

  if vStModeActivation = cEtatinitial then
    lBoOkActive := _AutoriseActivation( lTobParamCycle.GetString('CPC_ACTIVATIONSQL'))
  else
    lBoOkActive := (vStModeActivation = cActif);

  // Si cycle désactivé, et qu'après évaluation de la Cdt d'activation, il sera
  // toujours désactivé, on sort, pas besoin de traitment
  if (lStActivation = '-') and (not lBoOkActive) then Exit;

  // Traitement spécifique si on remet en état initial
  if (vStModeActivation = cEtatInitial) then
  begin
    // Supprime VisaAuto si il y en a un, tout en conservant les visas Manuels
    _MiseAJourVisaAuto( vStCycle, vStExercice, False );
  end;

  if lBoOkActive then
  begin
    lStSql := '';
    lStWhereCompte    := AnalyseCompte( lTobParamCycle.GetString('CPC_LISTECOMPTE'), fbGene, False, False);
    lStWhereExclusion := AnalyseCompte( lTobParamCycle.GetString('CPC_LISTEEXCLUSION'), fbGene, True, False);
    if lStWhereCompte <> '' then
    begin
      lStSql := 'UPDATE GENERAUX SET G_CYCLEREVISION = "' + vStCycle + '" ' +
                'WHERE ' + lStWhereCompte;

      if lStWhereExclusion <> '' then
        lStSql := lStSql + ' AND ' + lStWhereExclusion ;

      ExecuteSql( lStSql );

      // GCO - 24/07/2007 - FQ 21170
      (*lStSql := 'UPDATE CREVGENERAUX SET ' +
                'CRG_CYCLEREVISION = (SELECT G_CYCLEREVISION FROM GENERAUX ' +
                'WHERE G_GENERAL = CRG_GENERAL ' +
                'AND G_CYCLEREVISION = "' + vStCycle + '") WHERE ' +
                'CRG_EXERCICE = "' + vStExercice + '"';*)

      lStSql := 'UPDATE CREVGENERAUX SET ' +
                'CRG_CYCLEREVISION = (SELECT G_CYCLEREVISION FROM GENERAUX ' +
                'WHERE G_GENERAL = CRG_GENERAL) WHERE ' +
                'CRG_EXERCICE = "' + vStExercice + '"';
      ExecuteSQL( lStSql );
    end;

    ExecuteSQL('UPDATE CREVCYCLE SET CCY_ACTIVECYCLE = "X", ' +
               'CCY_ETATCYCLEPAR = "", ' +
               'CCY_ETATCYCLELE = "' + UsDateTime(iDate1900) + '", ' +
               'CCY_ETATCYCLE = "' + cATraiter + '" WHERE ' +
               'CCY_CODECYCLE = "' + vStCycle + '" AND ' +
               'CCY_EXERCICE = "' + vStExercice + '"');
  end
  else
  begin
    // Supprime VisaAuto si il y en a un, tout en conservant les visas Manuels
    _MiseAJourVisaAuto( vStCycle, vStExercice, False );

    // Désactivation du cycle
    ExecuteSQL('UPDATE CREVCYCLE SET CCY_ACTIVECYCLE = "-", ' +
               'CCY_ETATCYCLE = "' + cNonActif + '", ' +
               'CCY_ETATCYCLEPAR = "", ' +
               'CCY_ETATCYCLELE = "' + UsDateTime(idate1900) + '" WHERE ' +
               'CCY_CODECYCLE = "' + vStCycle + '" AND ' +
               'CCY_EXERCICE = "' + vStExercice + '"');

    // Supprime la notion de cycle dans CREVGENERAUX
    lStSql := 'UPDATE CREVGENERAUX SET CRG_CYCLEREVISION = "" WHERE ' +
              'CRG_CYCLEREVISION = "' + vStCycle + '" AND '+
              'CRG_EXERCICE = "' + vStExercice + '"';
    ExecuteSQL( lStSql );

    // Supprime la notion de cycle dans GENERAUX
    lStSql := 'UPDATE GENERAUX SET G_CYCLEREVISION = "" WHERE ' +
              'G_CYCLEREVISION = "' + vStCycle + '"';

    ExecuteSQL( lStSql );
  end;

  // Ajout d'une ligne dans l'historique du cycle lors du changement d'état
  _AjouteCRevHistoCycle(vStCycle, vStExercice, IIF(lBoOkActive, cATraiter, cNonActif), Date);

  // Passage du Statut Révision en mode "à calculer"
  VH^.BStatutRevision.ImageIndex := 3;
  FreeAndNil(lTobParamCycle);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TRic.ChangeEtatCycle(const vStCycle, vStExercice, vStEtatCycle, vStUser: string; const vDate: TDateTime): Boolean;
begin
  Result := True;

  try
    if ExecuteSQL('UPDATE CREVCYCLE SET ' +
        'CCY_ETATCYCLE = "' + vStEtatCycle + '", ' +
        'CCY_ETATCYCLEPAR = "' + IIF((vStEtatCycle = cATraiter) or (vStEtatCycle = '') , '', vStUser) + '", ' +
        'CCY_ETATCYCLELE = "' + USDateTime(IIF((vStEtatCycle = cATraiter) or (vStEtatCycle = ''), iDate1900, vDate)) + '" ' +
        'WHERE ' +
        'CCY_CODECYCLE = "' + vStCycle + '" AND ' +
        'CCY_EXERCICE = "' + vStExercice + '"') <> 1 then
      raise
        Exception.Create('Erreur de MAJ CREVCYCLE');

    // Ajout d'une ligne dans l'historique du cycle lors du changement d'état
    _AjouteCRevHistoCycle(vStCycle, vStExercice, vStEtatCycle, vDate);

    // Validation du cycle -> VISAAUTO passe à X
    _MiseAJourVisaAuto( vStCycle, vStExercice, (vStEtatCycle = cValide) or (vStEtatCycle = cSupervise));

  except
    on E: Exception do
    begin
      Result := False;
      PgiError('Erreur de requête SQL : ' + E.Message, 'ChangeEtatCycle');
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/06/2007
Modifié le ... :   /  /
Description .. : Test l'unicité des fourchettes des comptés généraux d'un cycle
Mots clefs ... :
*****************************************************************}
function TRic.ControleUnicite(const vStCycle, vStPredefini, vStListeCompte, vStListeExclusion, vStPlan : string; var vStCycleDoublon : string ): Boolean;
var lRubInterval1     : TCycleInterval;
    lRubinterval2     : TCycleInterval;
    lStListeCompte    : string;
    lStListeExclusion : string;
    lStCompte         : string;
    lStExclusion      : string;
    lQuery            : TQuery;
    lTobListeLigne1   : Tob;
    lTobListeLigne2   : Tob;
    lTob              : Tob;
    i,j               : integer;

    lStPlan           : string; // GCO - 19/09/2007 - FQ 21460
    lStWherePlan      : string;

    procedure _RemplitRubInterval( var vInterval : TCycleInterval; vStCompte, vStExclusion : string);
    var lBoIntervalEntier : Boolean;
        lSt : string;
        lInPos1 : integer;
        lInPos2 : integer;
    begin
      FillChar( vInterval, sizeOf( vInterval), #0);

      // Remplissage des bornes extérieures
      lBoIntervalEntier := Pos(':', vStCompte) > 0;
      if lBoIntervalEntier then
      begin
        lInPos1 := Pos(':', vStCompte);
        lSt     := Copy( vStCompte, 1, lInPos1 - 1 );
        vInterval.Borne1Deb := BourreEtLess( lSt, fbGene); // OK

        lSt := Copy( vStCompte, lInPos1 + 1, Length(vStCompte));
        vInterval.Borne2Fin := BourreLaDoncSurLesComptes( lSt, '9');
      end
      else
      begin
        lSt := vStCompte;
        vInterval.Borne1Deb := BourreEtLess( lSt, fbGene);
        vInterval.Borne2Fin := BourreLaDoncSurLesComptes( lSt, '9');
      end;

      // Remplissage des bornes intérieures
      if vStExclusion = '' then
      begin
        vInterval.AvecExclusion := False;
        vInterval.Borne1Fin := vInterval.Borne2Fin;
        vInterval.Borne2Deb := vInterval.Borne1Deb;
      end
      else
      begin
        // Rajout d'un ; pour marquer la fin de la chaine
        vStExclusion := vStExclusion + ';';
        vInterval.AvecExclusion := True;
        lBoIntervalEntier := Pos(':', vStExclusion) > 0;
        if lBoIntervalEntier then
        begin
          lInPos1 := Pos(':', vStExclusion);
          lInPos2 := Pos(';', vStExclusion);

          lSt := Copy( vStExclusion, 1, lInPos1 - 1);
          vInterval.Borne1Fin := IntToStr( StrToInt( lst ) - 1);
          vInterval.Borne1Fin := BourreLaDoncSurLesComptes( vInterval.Borne1Fin, '9');

          lSt := Copy( vStExclusion, lInPos1 + 1, (lInPos2) - (lInPos1 + 1));
          if Length(lSt) = VH^.Cpta[fbGene].lg then
            vInterval.Borne2Deb := lSt
          else
          begin
            vInterval.Borne2Deb := IntToStr( StrToInt ( lSt ) + 1);
            vInterval.Borne2Deb := BourreEtLess( vInterval.Borne2Deb, fbGene);
          end;
        end
        else
        begin
          lSt := Copy( vStExclusion, 1, Pos(';', vStExclusion)-1);
          vInterval.Borne1Fin := IntToStr( StrToInt( lSt ));//IntToStr( StrToInt( lSt )- 1);
          vInterval.Borne1Fin := BourreLaDoncSurLesComptes( vInterval.Borne1Fin, '9');

          vInterval.Borne2Deb := IntToStr( StrToInt( lSt ));//IntToStr( StrToInt( lSt ) + 1);
          vInterval.Borne2Deb := BourreEtLess( vInterval.Borne2Deb, fbGene);
        end;
      end;
    end;

    function _ChevauchementRubInterval : Boolean;
    begin
      Result := False;

      // GCO - 28/04/2005 - Pour les if, y avait moyen d'en faire un seul mais
      // en cas de bug c'est plus simple à trouver qui renvoie TRUE
      if ((lRubInterval2.Borne1Deb >= lRubInterval1.Borne1Deb) and
          (lRubInterval2.Borne1Deb <= lRubInterval1.Borne1Fin)) or
          ((lRubInterval1.Borne1Deb >= lRubInterval2.Borne1Deb) and
          (lRubInterval1.Borne1Deb <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne1Fin >= lRubInterval1.Borne1Deb) and
          (lRubInterval2.Borne1Fin <= lRubInterval1.Borne1Fin)) or
          ((lRubInterval1.Borne1Fin >= lRubInterval2.Borne1Deb) and
          (lRubInterval1.Borne1Fin <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne1Deb >= lRubInterval1.Borne2Deb) and
          (lRubInterval2.Borne1Deb <= lRubInterval1.Borne2Fin)) or
          ((lRubInterval1.Borne1Deb >= lRubInterval2.Borne2Deb) and
          (lRubInterval1.Borne1Deb <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne1Fin >= lRubInterval1.Borne2Deb) and
          (lRubInterval2.Borne1Fin <= lRubInterval1.Borne2Fin)) or
          ((lRubInterval1.Borne1Fin >= lRubInterval2.Borne2Deb) and
          (lRubInterval1.Borne1Fin <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Deb >= lRubInterval1.Borne1Deb) and
         (lRubInterval2.Borne2Deb <= lRubInterval1.Borne1Fin)) or
         ((lRubInterval1.Borne2Deb >= lRubInterval2.Borne1Deb) and
         (lRubInterval1.Borne2Deb <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Fin >= lRubInterval1.Borne1Deb) and
          (lRubInterval2.Borne2Fin <= lRubInterval1.Borne1Fin)) or
          ((lRubInterval1.Borne2Fin >= lRubInterval2.Borne1Deb) and
          (lRubInterval1.Borne2Fin <= lRubInterval2.Borne1Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Deb >= lRubInterval1.Borne2Deb) and
          (lRubInterval2.Borne2Deb <= lRubInterval1.Borne2Fin)) or
          ((lRubInterval1.Borne2Deb >= lRubInterval2.Borne2Deb) and
          (lRubInterval1.Borne2Deb <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;

      if ((lRubInterval2.Borne2Fin >= lRubInterval1.Borne2Deb) and
         (lRubInterval2.Borne2Fin <= lRubInterval1.Borne2Fin)) or
         ((lRubInterval1.Borne2Fin >= lRubInterval2.Borne2Deb) and
         (lRubInterval1.Borne2Fin <= lRubInterval2.Borne2Fin)) then
      begin
        Result := True;
        Exit;
      end;
    end;

begin
  Result := True;
  lQuery := nil;
  try
    try
      // Contient les lignes du Cycle révision que l'on sauve
      lTobListeLigne1   := Tob.Create('', nil, -1);
      lStListeCompte    := vStListeCompte;
      lStListeExclusion := vStListeExclusion;

      while lStListeCompte <> '' do
      begin
        lStCompte    := ReadTokenSt( lStlisteCompte );
        lStExclusion := ReadTokenSt( lStlisteExclusion );

        if lStCompte <> '' then
        begin
          lTob := Tob.Create( '', lTobListeLigne1, -1);
          lTob.AddChampSupValeur('LISTECOMPTE', lStCompte, False );
          lTob.AddChampSupValeur('LISTEEXCLUSION', lStExclusion, False );
        end;
      end;

      // Contient les lignes des Cycles de révision dans la base
      lTobListeLigne2 := Tob.Create('', nil, -1);

      // GCO - 19/09/2007 - FQ 21460
      lStPlan := vStPlan;
      lStWherePlan := '(';
      while Pos(';', lStPlan) > 0 do
      begin
        lStWherePlan := lStWherePlan + ' CPC_PLANASSOCIE LIKE "%' + ReadTokenSt( lStPlan ) + ';%" OR ';
      end;
      // Suppression du ' OR '
      lStWherePlan := Copy(lStWherePlan, 0, Length(lStWherePlan) - 3) + ')';

      // Charge tout les intervalles de CREVPARAMCYCLE sauf celle concernée
      lQuery := OpenSql('SELECT CPC_CODECYCLE, CPC_PLANASSOCIE, CPC_LISTECOMPTE, ' +
                      'CPC_LISTEEXCLUSION FROM CREVPARAMCYCLE WHERE ' +
                      'CPC_CODECYCLE <> "' + vStCycle + '" AND ' +
                      lStWherePlan + ' ORDER BY CPC_CODECYCLE', True);
      // FIN GCO - 19/09/20007

      while not lQuery.Eof do
      begin
        lStListeCompte    := lQuery.FindField('CPC_LISTECOMPTE').AsString;
        lStListeExclusion := lQuery.FindField('CPC_LISTEEXCLUSION').AsString;

        while lStListeCompte <> '' do
        begin
          lStCompte    := ReadTokenSt( lStlisteCompte );
          lStExclusion := ReadTokenSt( lStlisteExclusion );

          if lStCompte <> '' then
          begin
            lTob := Tob.Create( '', lTobListeLigne2, -1);
            lTob.AddChampSupValeur('CODECYCLE', lQuery.FindField('CPC_CODECYCLE').AsString);
            lTob.AddChampSupValeur('LISTECOMPTE', lStCompte, False );
            lTob.AddChampSupValeur('LISTEEXCLUSION', lStExclusion, False );
          end;
        end;
        lQuery.Next;
      end;

      for i := 0 to lTobListeLigne1.Detail.Count -1 do
      begin
        // On remplit la ligne[i] dy cycle de révision que l'on sauve
        _RemplitRubInterval( lRubInterval1,
                             lTobListeLigne1.Detail[i].GetString('LISTECOMPTE'),
                             lTobListeLigne1.Detail[i].GetString('LISTEEXCLUSION'));

        // Parcours des lignes de chaque ycle de révision que l'on doit controller
        for j := 0 to lTobListeLigne2.Detail.Count -1 do
        begin
          _RemplitRubInterval( lRubInterval2,
                               lTobListeLigne2.Detail[j].GetString('LISTECOMPTE'),
                               lTobListeLigne2.Detail[j].GetString('LISTEEXCLUSION'));

          if _ChevauchementRubInterval then
          begin
            vStCycleDoublon := lTobListeLigne2.Detail[j].GetString('CODECYCLE');
            Result := False;
            Break;
          end;
        end;
      end;

    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : ComprareInterval');
    end;

  finally
    Ferme( lQuery ) ;
    FreeAndNil( lTobListeLigne1 );
    FreeAndNil( lTobListeLigne2 );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/06/2007
Modifié le ... : 04/07/2007
Description .. : Impression du dossier de travail
Suite ........ : Page de Garde
Suite ........ : Note de supervision du dossier
Suite ........ : Programme de travail
Suite ........ : Pour tout les cycles activés
Suite ........ : | - Note de contrôle cycle
Suite ........ : | - Balance de révision d'un cycle
Suite ........ : | - Pour tout les comptes d'un cycle
Suite ........ : |   | - Note de contrôle Compte
Suite ........ : |   | - Note de travail
Suite ........ : |   | - Tableau de variation
Suite ........ : - - -
Mots clefs ... :
*****************************************************************}
class procedure TRic.ImpressionDossierTravail;
var i : integer;
    lTobCycle   : Tob;
    lTobLeCycle : Tob;
    lTobGene    : Tob;
    lTobLeGene  : Tob;
    lStSQL      : HString;
    lStCycle    : string;
    lStGene     : string;
    lStFileName : string;
    lTotalPage  : integer;
    lBoOldNoPrintDialog : Boolean;
    lBoOldQRPdf         : Boolean;
begin

  if PgiAsk('Attention. Ce traîtement peut être relativement long.' + #13 + #10 +
            'Voulez-vous continuer ?', 'Impression du dossier de travail') = MrNo then Exit;

  lTobCycle := Tob.Create('LESCYCLES', nil, -1 );
  lTobGene  := Tob.Create('LESGENERAUX', nil, -1);

  lBoOldNoPrintDialog := V_Pgi.NoPrintDialog;
  lBoOldQRPdf         := V_Pgi.QRPdf;

  try
    try
    {$IFNDEF EAGLSERVER}
      InitMoveProgressForm(nil,
      TraduireMemoire('Préparation du dossier de travail'),
      TraduireMemoire('Traitement en cours. Veuillez patienter...'), 10, True, True);
    {$ENDIF}

      V_Pgi.SilentMode    := True;
      V_Pgi.NoPrintDialog := True;
      V_Pgi.QRPdf         := True;
      VH^.StopEdition     := False;
      lTotalPage          := 0;
      lStFileName         := TempFileName(); // Fichier de sortie
      V_PGI.QrPdfQueue    := TempFileName(); // Fichier temporaire de Queue

      // Démarrage de l'impression Batch
      THPrintBatch.StartPdfBatch(lStFileName);  //StartPdfBatch(lStFileName);

      // Page de garde
      CPLanceEtat_PageDeGarde;
    {$IFDEF EAGLCLIENT}
      THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
    {$ENDIF}

      // Note de supervision
      lTotalPage := lTotalPage + CPLanceEtat_NoteSupervision( lTotalPage );
    {$IFDEF EAGLCLIENT}
      THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
    {$ENDIF}

      // Programme de travail
      lTotalPage := lTotalPage + CPLanceEtat_ProgrammeTravail( lTotalPage );
    {$IFDEF EAGLCLIENT}
      THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
    {$ENDIF}

      lTobCycle.LoadDetailFromSQL('SELECT CCY_CODECYCLE, ' +
        'IIF((CBN_CODE IS NOT NULL), "X", "-") MEMOCYCLE FROM CREVCYCLE ' +
        'LEFT JOIN CREVBLOCNOTE ON CBN_CODE = CCY_CODECYCLE AND ' +
        'CBN_EXERCICE = "' + VH^.EnCours.Code + '" WHERE ' +
        'CCY_EXERCICE = "' + VH^.EnCours.Code + '" AND CCY_ACTIVECYCLE = "X" ' +
        'ORDER BY CCY_CODECYCLE');

      for i := 0 to lTobCycle.Detail.Count - 1 do
      begin
        lTobLeCycle := lTobCycle.Detail[i];
        lStCycle := lTobLeCycle.GetString('CCY_CODECYCLE');

      {$IFNDEF EAGLSERVER}
        if not MoveCurProgressForm('Cycle de révision ' + lStCycle ) then
          VH^.StopEdition := True;

        Application.ProcessMessages;
        if VH^.StopEdition then Break;
      {$ENDIF}

        // Note de contrôle CYCLE
        if lTobLeCycle.GetString('MEMOCYCLE') = 'X' then
          lTotalPage := lTotalPage + CPLanceEtat_NoteCtrlCycle( lStCycle, lTotalPage );

        // Balance de Cycle
        lTotalPage := lTotalPage + BalanceAvecVariation( lStCycle, lTotalPage );
      {$IFDEF EAGLCLIENT}
        THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
      {$ENDIF}

      lStSQL := 'SELECT DISTINCT CCY_CODECYCLE, G_GENERAL, G_LIBELLE, ' +
                'IIF((G_TOTDEBE <> 0 OR G_TOTCREE <> 0), "X", "-") MVT, ' +
                'IIF((CNO_GENERAL IS NOT NULL), "X", "-") NOTETRAVAIL, ' +                        'IIF((CTV_GENERAL IS NOT NULL), "X", "-") TABLEAUVARIATION FROM CREVCYCLE ' +
                'LEFT JOIN GENERAUX ON G_CYCLEREVISION = CCY_CODECYCLE ';

      lStSQL := lStSQL + 'LEFT JOIN CPNOTETRAVAIL ON CNO_GENERAL = G_GENERAL AND ' +
                'CNO_EXERCICE = "' + VH^.EnCours.Code + '" ' +
                'LEFT JOIN CPTABLEAUVAR  ON CTV_GENERAL = G_GENERAL AND ' +
                'CTV_EXERCICE = "' + VH^.EnCours.Code + '" WHERE (' +
                '(CNO_GENERAL IS NOT NULL) OR ' +
                '(CTV_GENERAL IS NOT NULL) OR ' +
                '(G_TOTDEBE <> 0 OR G_TOTCREE <> 0)) AND ' +
                'CCY_EXERCICE = "' + VH^.Encours.Code + '" AND CCY_ACTIVECYCLE = "X" '+
                'ORDER BY CCY_CODECYCLE, G_GENERAL';

        lTobGene.LoadDetailFromSql( lStSql );

        lTobLeGene := lTobGene.FindFirst(['CCY_CODECYCLE'],[lStCycle], False);
        while lTobLeGene <> nil do
        begin
          lStGene := lTobLeGene.GetString('G_GENERAL');
        {$IFNDEF EAGLSERVER}
          if not MoveCurProgressForm('Cycle de révision ' + lStCycle + ' - ' + 'Compte général ' + lStGene) then
            VH^.StopEdition := True;

          Application.ProcessMessages;
          if VH^.StopEdition then Break;
        {$ENDIF}

          // Note de Contrôle Compte
          if lTobLeGene.GetString('MVT') = 'X' then
          begin
            lTotalPage := lTotalPage + CPLanceEtat_NoteCtrlCompte( lStGene, lTotalPage );
          {$IFDEF EAGLCLIENT}
            THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
          {$ENDIF}
          end;

          // Note de Travail
          if lTobLeGene.GetString('NOTETRAVAIL') = 'X' then
          begin
            lTotalPage := lTotalPage + CPLanceEtat_NoteTravail( lStGene, VH^.EnCours.Code, lTotalPage );
          {$IFDEF EAGLCLIENT}
            THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
          {$ENDIF}
          end;

          // Tableau de Variation
          if lTobGene.GetString('TABLEAUVARIATION') = 'X' then
          begin
            lTotalPage := lTotalPage + CPLanceEtat_TabVariation( lStGene, VH^.EnCours.Code, lTotalPage );
          {$IFDEF EAGLCLIENT}
            THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
          {$ENDIF}
          end;

          FreeAndNil(lTobLeGene);
          lTobLeGene := lTobGene.FindFirst(['CCY_CODECYCLE'],[lStCycle], False);
        end;
      end;

    except
      on E: Exception do PgiError('Erreur : ' + E.Message, 'Class Fonction : ImpressionDossierTravail');
    end;

  finally
    THPrintBatch.StopPdfBatch; //CancelPDFBatch();

  {$IFNDEF EAGLSERVER}
     FiniMoveProgressForm;
  {$ENDIF}

    PreviewPDFFile('Dossier de travail : ' + V_Pgi.NoDossier + ' - ' + V_Pgi.NomSociete, lStFileName);

    // Remise en place du Contexte
    V_Pgi.NoPrintDialog := lBoOldNoPrintDialog;
    V_Pgi.QRPdf         := lBoOldQRPdf;
    V_Pgi.SilentMode    := False;
    VH^.StopEdition     := False;

    FreeAndNil( lTobCycle );
    FreeAndNil( lTobGene );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
class procedure TRic.CPLanceEtat_PageDeGarde;
begin
  LanceEdition('REV', 'PDG', '', 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/07/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
class function TRic.BalanceAvecVariation(const vStCycle: string;  vNumPage: integer): integer;
var lStEnCours   : HString;
    lStPrecedent : HString;
    lStSql       : HString;
    lStCritere   : HString;
    lStTitre1    : HString;
    lStTitre2    : HString;
begin
  Result := 0;

  ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_Pgi.User + '"');

  // Insertion pour Précédent
  lStPrecedent := 'INSERT INTO CEDTBALANCE (CED_USER, CED_COMPTE, CED_NATURE, ' +
    'CED_LIBELLE,CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2,CED_CREDIT3, ' +
    'CED_DEBIT3, CED_CREDIT4, CED_DEBIT4,CED_CREDIT5, CED_DEBIT5, CED_CREDIT6, ' +
    'CED_DEBIT6, CED_RUPTURE, CED_COLLECTIF, CED_COMPTE2, CED_LIBELLE2) ';

  lStPrecedent := lStPrecedent +
    ' SELECT "' + V_Pgi.User + '", G_GENERAL, G_NATUREGENE, G_LIBELLE, ' +
    '0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "-", CCY_CODECYCLE, ' +
    // GCO - 04/10/2007 - FQ 21227
    'SUBSTRING(CCY_LIBELLECYCLE, 1,35) FROM GENERAUX ' +
    'LEFT JOIN CREVCYCLE ON CCY_CODECYCLE = G_CYCLEREVISION AND ' +
    'CCY_EXERCICE = "' + VH^.EnCours.Code + '" ' +
    'LEFT JOIN ECRITURE ON E_GENERAL=G_GENERAL WHERE ' +
    //'G_GENERAL BETWEEN "101079" AND "999999"   AND
    'G_CYCLEREVISION = "' + vStCycle + '" AND E_QUALIFPIECE = "N" AND ' +
    'E_EXERCICE = "' + VH^.Precedent.Code + '"  AND ' +
    'E_DATECOMPTABLE >= "' + UsDateTime(VH^.Precedent.Deb) + '" AND ' +
    'E_DATECOMPTABLE <= "' + UsDateTime(VH^.Precedent.Fin) + '" AND ' +
    '(((G_CONFIDENTIEL = "-") OR (G_CONFIDENTIEL = "X")) OR (G_CONFIDENTIEL <= "1")) ' +
    'GROUP BY G_GENERAL, G_NATUREGENE, G_LIBELLE,CCY_CODECYCLE, CCY_LIBELLECYCLE';

  ExecuteSQL( lStPrecedent );

  // Insertion pour EnCours
  lStEnCours := 'INSERT INTO CEDTBALANCE (CED_USER, CED_COMPTE, CED_NATURE, ' +
    'CED_LIBELLE, CED_CREDIT1, CED_DEBIT1, CED_CREDIT2, CED_DEBIT2, ' +
    'CED_CREDIT3, CED_DEBIT3, CED_CREDIT4, CED_DEBIT4,CED_CREDIT5, ' +
    'CED_DEBIT5, CED_CREDIT6, CED_DEBIT6, CED_RUPTURE, CED_COLLECTIF, ' +
    'CED_COMPTE2, CED_LIBELLE2) ';

  lStEnCours := lStEnCours + 'SELECT "'+ V_PGI.User + '", G_GENERAL, ' +
    'G_NATUREGENE, G_LIBELLE, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", "-", ' +
    // GCO - 04/10/2007 - FQ 21227
    'CCY_CODECYCLE, SUBSTRING(CCY_LIBELLECYCLE, 1,35) FROM GENERAUX ' +
    'LEFT JOIN CREVCYCLE ON CCY_CODECYCLE = G_CYCLEREVISION AND ' +
    'CCY_EXERCICE = "' + VH^.EnCours.Code + '" ' +
    'LEFT JOIN ECRITURE ON E_GENERAL=G_GENERAL WHERE ' +
    //'G_GENERAL BETWEEN "101079" AND "999999"   AND ' +
    'G_CYCLEREVISION = "' + vStCycle + '" AND E_QUALIFPIECE = "N" AND ' +
    'E_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
    'E_DATECOMPTABLE >= "' + UsDateTime(VH^.EnCours.Deb) + '" AND ' +
    'E_DATECOMPTABLE <= "' + UsDateTime(VH^.EnCours.Fin) + '" AND ' +
    '(G_GENERAL NOT IN (SELECT CED_COMPTE FROM CEDTBALANCE WHERE ' +
    'CED_USER = "'+ V_Pgi.User + '")) AND ' +
    '(((G_CONFIDENTIEL = "-") OR (G_CONFIDENTIEL = "X")) OR (G_CONFIDENTIEL <= "1")) ' +
    'GROUP BY G_GENERAL, G_NATUREGENE, G_LIBELLE,CCY_CODECYCLE, CCY_LIBELLECYCLE';

  ExecuteSql( lStEnCours );

  if not ExisteSql('SELECT CED_COMPTE FROM CEDTBALANCE WHERE CED_USER = "' + V_Pgi.User + '"') then
    Exit;

  // DEBIT 1
  ExecuteSQL('UPDATE CEDTBALANCE SET CED_DEBIT1 = ( ' +
    'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_DEBIT <> 0), E_DEBIT - ' +
    'E_COUVERTURE, E_DEBIT)) FROM ECRITURE WHERE ' +
    '(E_QUALIFPIECE = "N") AND E_EXERCICE ="' + VH^.EnCours.Code + '" AND ' +
    'E_DATECOMPTABLE BETWEEN "' + UsDateTime(VH^.EnCours.Deb) + '" AND ' +
    '"' + UsDateTime(VH^.EnCours.Fin) + '" AND E_GENERAL = CED_COMPTE ) WHERE ' +
    'CED_USER = "' + V_Pgi.User + '"');

  ExecuteSQL('UPDATE CEDTBALANCE SET CED_DEBIT1 = 0 WHERE CED_DEBIT1 IS NULL AND ' +
             'CED_USER = "' + V_Pgi.User + '"');

  // CREDIT 1
  ExecuteSQL('UPDATE CEDTBALANCE SET CED_CREDIT1 = ( ' +
    'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_CREDIT <> 0),E_CREDIT - ' +
    'E_COUVERTURE, E_CREDIT)) FROM ECRITURE WHERE ' +
    '(E_QUALIFPIECE = "N") AND E_EXERCICE = "' + VH^.Encours.Code + '" AND ' +
    'E_DATECOMPTABLE BETWEEN "' + UsDateTime(VH^.EnCours.Deb) + '" AND ' +
    '"' + UsDateTime(VH^.EnCours.Fin) + '" AND E_GENERAL = CED_COMPTE  ) WHERE ' +
    'CED_USER = "' + V_Pgi.User + '"');

  ExecuteSQL('UPDATE CEDTBALANCE SET CED_CREDIT1 = 0 WHERE CED_CREDIT1 IS NULL AND ' +
             'CED_USER = "' + V_Pgi.User + '"');

  // DEBIT 2
  ExecuteSQL('UPDATE CEDTBALANCE SET CED_DEBIT2 = ( ' +
    'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_DEBIT <> 0), E_DEBIT - ' +
    'E_COUVERTURE, E_DEBIT)) FROM ECRITURE WHERE ' +
    '(E_QUALIFPIECE="N" ) AND E_EXERCICE = "' + VH^.Precedent.Code + '" AND ' +
    'E_DATECOMPTABLE BETWEEN "' + UsDateTime(VH^.Precedent.Deb )+ '" AND ' +
    '"' + UsDateTime(VH^.Precedent.Fin) + '" AND E_GENERAL = CED_COMPTE ) WHERE ' +
    'CED_USER = "' + V_Pgi.User + '"');

  ExecuteSQL('UPDATE CEDTBALANCE SET CED_DEBIT2 = 0 WHERE CED_DEBIT2 IS NULL AND ' +
             'CED_USER = "' + V_Pgi.User + '"');

  // CREDIT 2
  ExecuteSQL('UPDATE CEDTBALANCE SET CED_CREDIT2 = ( ' +
    'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_CREDIT <> 0), E_CREDIT - ' +
    'E_COUVERTURE, E_CREDIT)) FROM ECRITURE WHERE ' +
    '(E_QUALIFPIECE ="N") AND E_EXERCICE = "' + VH^.Precedent.Code + '" AND ' +
    'E_DATECOMPTABLE BETWEEN "' + UsDateTime(VH^.Precedent.Deb ) + '" AND ' +
    '"' + UsDateTime(VH^.Precedent.Fin) + '" AND E_GENERAL = CED_COMPTE ) WHERE ' +
    'CED_USER = "' + V_Pgi.User + '"');

  ExecuteSQL('UPDATE CEDTBALANCE SET CED_CREDIT2 = 0 WHERE CED_CREDIT2 IS NULL AND ' +
             'CED_USER = "' + V_Pgi.User + '"');

  lStSql := 'SELECT CED_COMPTE CED_GENERAL, CED_LIBELLE, CED_COMPTE2 CED_CYCLE, ' +
            'CED_LIBELLE2 CED_CYCLELIB, CED_RUPTURE, "" RUPTURELIB, CED_DEBIT1, ' +
            'CED_CREDIT1, CED_DEBIT2, CED_CREDIT2, CED_COLLECTIF, CED_NATURE, G_SENS, ' +
            'IIF(((G_SENS="D") OR (G_SENS="M")) , CED_DEBIT1-CED_CREDIT1, ' +
            'CED_CREDIT1-CED_DEBIT1) SOLDEE, ' +
            'IIF(((G_SENS="D") OR (G_SENS="M")) , ';

  lStSql := lStSql + 'CED_DEBIT2-CED_CREDIT2, CED_CREDIT2-CED_DEBIT2) SOLDEP ' +
            'FROM CEDTBALANCE LEFT JOIN GENERAUX ON G_GENERAL=CED_COMPTE ' +
            'WHERE CED_USER = "' + V_Pgi.User + '" ' +
            'ORDER BY CED_CYCLE, CED_RUPTURE, CED_COMPTE';

  lStTitre1  := TraduireMemoire('En cours') + ' : ' + FormatDateTime('yyyy', VH^.EnCours.Fin);
  lStTitre2  := IIF( VH^.Precedent.Code <> '', TraduireMemoire('Précédent') + ': ' +
                FormatDateTime('yyyy', VH^.Precedent.Fin), '');

  // GCO - 02/082007 - FQ 21229
  lStCritere := 'RUPCYCLEREVISION=X`' +
                'NATURECPT=<<' + TraduireMemoire('Tous') + '>>' + '`' +
                'ETABLISSEMENT=<<' + TraduireMemoire('Tous') + '>>' + '`' +
                'EXERCICE=' + TraduireMemoire('En cours') + '`' +
                'AVECQUALIFPIECE='+ TraduireMemoire('Normal') + '`' +
                'DATECOMPTABLE=' + DateToStr(VH^.EnCours.Deb) + '`' +
                'DATECOMPTABLE_=' + DateToStr(VH^.EnCours.Fin) + '`' +
                'TITRECOL1=' + lStTitre1 + '`' +
                'TITRECOL2=' + lStTitre2 + '';

  Result := LanceEdition('BAR', 'BAR', lStSql, lStCritere, vNumPage);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
class function TRic.LanceEdition(const vStNatureEtat, vStCodeEtat, vStWhere, vStRangeCritere: string; vNumPage: integer): integer;
begin
{$IFDEF EAGLCLIENT}
  Result := LanceEtat('E', vStNatureEtat, vStCodeEtat, True, False, False, nil,
                      vStWhere, '', False, 0, vStRangeCritere, vNumPage);
{$ELSE}
  Result := LanceEtat('E', vStNatureEtat, vStCodeEtat, True, False, False, nil,
                      vStWhere, '', False, 0, vStRangeCritere, nil, vNumPage);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
class function TRic.CPLanceEtat_NoteSupervision( const vNumPage : integer) : integer;
var lStWhere : HString;
begin
  lStWhere := 'AND CBN_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
              'CIR_NODOSSIER = "' + V_Pgi.NoDossier + '"';

  Result := LanceEdition('CCS', 'CCS', lStWhere, '', vNumPage);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2007
Modifié le ... : 02/08/2007
Description .. : 
Mots clefs ... : 
*****************************************************************}
class function TRic.CPLanceEtat_ProgrammeTravail( const vNumPage : integer ) : integer;
var lStWhere : HString;
begin
  lStWhere := 'AND CCY_EXERCICE = "' + VH^.EnCours.Code + '"';

  Result := LanceEdition('CCP', 'CCP', lStWhere , '', vNumPage );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2007 suite à la FQ 21228
Modifié le ... :   /  /
Description .. : Edition de révision - Note de contrôle Compte
Mots clefs ... :
*****************************************************************}
class function TRic.CPLanceEtat_NoteCtrlCompte( const vStGeneral : HString; vNumPage : integer = 0 ) : integer;
var lStWhere : HString;
begin
  lStWhere := '((E_BLOCNOTE IS NOT NULL) OR (CBN_BLOCNOTE IS NOT NULL)) AND ' +
              'E_GENERAL = "' + vStGeneral + '" AND ' +
              'E_EXERCICE = "' + VH^.Encours.Code + '"';

  Result := LanceEdition('CCT', 'CCT', lStWhere, '', vNumPage);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
class function TRic.CPLanceEtat_NoteCtrlCycle( const vStCycle : HString; vNumPage : integer = 0 ) : integer;
var lStWhere : HString;
begin
  lStWhere := 'CCY_CODECYCLE = "' + vStCycle + '" AND ' +
              'CBN_EXERCICE = "' + VH^.EnCours.Code + '"';

  Result := LanceEdition('CCY', 'CCY', lStWhere, '', vNumPage);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
class function TRic.CPLanceEtat_NoteTravail( const vStGeneral, vStExercice : HString; vNumPage : integer = 0 ) : integer;
var lStWhere : HString;
begin
  lStWhere := 'CNO_EXERCICE = "'+ vStExercice + '" AND ' +
              'CNO_GENERAL = "' + vStGeneral + '"';

  Result := LanceEdition('REV', 'NT1', lStWhere, '', vNumPage);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
class function TRic.CPLanceEtat_TabVariation( const vStGeneral, vStExercice : HString; vNumPage : integer = 0 ) : integer;
var lStWhere : HString;
begin
  lStWhere := 'CTV_EXERCICE = "'+ vStExercice + '" AND ' +
              'CTV_GENERAL = "' + vStGeneral + '"';

  Result := LanceEdition('REV', 'TV1', lStWhere,  '', vNumPage );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
(*
procedure FQ14038;
var lNBPage : integer;
    lTotal  : integer;
begin
  V_Pgi.SilentMode    := True;
  V_Pgi.NoPrintDialog := True;
  V_Pgi.QRPdf         := True;
  VH^.StopEdition     := False;
  lTotal              := 0;

  V_Pgi.QrPdfQueue := 'C:\Truc.pdf';

  // Démarrage de l'impression Batch
  THPrintBatch.StartPdfBatch('C:\KPMG.PDF');

  // 1-
{$IFDEF EAGLCLIENT}
  lNbPage := LanceEtat('E', 'REV', 'PDG', True, False, False, nil, '', '',
                       False, 0, 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', lTotal);

  THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);

{$ELSE}
  lNbPage := LanceEtat('E', 'REV', 'PDG', True, False, False, nil, '', '',
                       False, 0, 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', nil, lTotal);
{$ENDIF}
  lTotal := lTotal + lNbPage;

  // 2-
{$IFDEF EAGLCLIENT}
  lNbPage := LanceEtat('E', 'REV', 'PDG', True, False, False, nil, '', '',
                       False, 0, 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', lTotal);
  THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
{$ELSE}
  lNbPage := LanceEtat('E', 'REV', 'PDG', True, False, False, nil, '', '',
                       False, 0, 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', nil, lTotal);
{$ENDIF}
  lTotal := lTotal + lNbPage;

  // 3-
{$IFDEF EAGLCLIENT}
  lNbPage := LanceEtat('E', 'REV', 'PDG', True, False, False, nil, '', '',
                       False, 0, 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', lTotal);
  THPrintBatch.AjoutPDF(V_Pgi.QRPdfQueue, True);
{$ELSE}
  lNbPage := LanceEtat('E', 'REV', 'PDG', True, False, False, nil, '', '',
                       False, 0, 'EXDATEFIN='+ DateToStr(VH^.Encours.Fin) + '`', nil, lTotal);
{$ENDIF}
  lTotal := lTotal + lNbPage;

  PgiInfo(IntToStr(lTotal), '');

  //CancelPDFBatch();
  THPrintBatch.StopPdfBatch();

  // Remise en place du Contexte
  V_Pgi.NoPrintDialog := False;
  V_Pgi.SilentMode    := False;
  VH^.StopEdition     := False;
end;*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/10/2007
Modifié le ... :   /  /    
Description .. : Fonction de test des conditions d'activation dans CREVPARAMCYCLE
Suite ........ : avec les cycles présents dans CREVCYCLE sur l'exo ENCOURS
Suite ........ : Si différence, le code cycle est présent dans le RESULT
Mots clefs ... :
*****************************************************************}
function ControleConditionActivation : HString;
var i : integer;
    lBoActivation : Boolean;
    FTobCRevCycle, lTob : Tob;
    lRic : TRic;
begin
  Result := '';

  lRic := nil;
  FTobCRevCycle := nil;
  try
    lRic := TRic.Create;

    FTobCRevCycle := Tob.Create('', nil, -1);
    FTobCRevCycle.LoadDetailFromSQL('SELECT CCY_CODECYCLE, CCY_ACTIVECYCLE, ' +
                  'CPC_ACTIVATIONSQL FROM CREVCYCLE ' +
                  'LEFT JOIN CREVPARAMCYCLE ON CCY_CODECYCLE = CPC_CODECYCLE AND ' +
                  'CCY_PREDEFINI = CPC_PREDEFINI WHERE ' +
                  'CCY_EXERCICE = "' + GetEnCours.Code + '" ORDER BY CCY_CODECYCLE');

    for i := 0 to FTobCRevCycle.Detail.count -1 do
    begin
      lTob := FTobCRevCycle.Detail[i];
      lBoActivation := lRic._AutoriseActivation(lTob.GetString('CPC_ACTIVATIONSQL'));

      if (lBoActivation and (lTob.GetString('CCY_ACTIVECYCLE') = '-')) or
         (not lBoActivation and (lTob.GetString('CCY_ACTIVECYCLE') = 'X')) then
      begin
        Result := Result + lTob.GetString('CCY_CODECYCLE') + ';';
      end;
    end;

  finally
    FreeAndNil( FTobCRevCycle );
    lRic.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2007
Modifié le ... :   /  /    
Description .. : Nouveau Millésime pour la révision, suite à une cloture donc
Suite ........ : on recopie les cycles de CREVPARAMCYCLE vers CREVCYCLE, et on
Suite ........ : les affecte aux comptes généraux
*****************************************************************}
procedure SynchroRICAvecExercice;
var lRic : TRic;
begin
  ExecuteSQL('UPDATE GENERAUX SET G_CYCLEREVISION = "", G_VISAREVISION = "-"');

  lRic := nil;
  try
    lRic := TRic.Create;
    if not ExisteSQL('SELECT CCY_CODECYCLE FROM CREVCYCLE WHERE CCY_EXERCICE = "' + VH^.EnCours.Code + '"') then
    begin
      lRic.ProgressForm := True;
      if not lRic.Affecte(VH^.Revision.Plan) then
        PgiInfo( 'Traitement annulé. ' + lRic.LastErrorMsg, 'Mise à jour de la révision');
    end
    else
      lRic._CREVGenerauxVersGeneraux;

  finally
   lRic.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////


end.
