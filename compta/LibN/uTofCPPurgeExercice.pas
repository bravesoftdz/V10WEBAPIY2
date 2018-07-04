 {***********UNITE*************************************************
Auteur  ...... :  Gilles COSTE
Créé le ...... : 31/05/2005
Modifié le ... : 23/08/2005 - GCO - FQ 16436
Description .. : Source TOF de la FICHE : CPPURGEEXERCICE ()
Mots clefs ... : TOF;CPPURGEEXERCICE
*****************************************************************}
Unit uTofCPPurgeExercice ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl,       // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,        // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     uTob,           // TOB
     ParamSoc,       // GetParamSocSecur
     ED_TOOLS,       // InitMoveProgressForm
     Ent1,           // VH^
     SAISUTIL,       // RDevise
     uTobDebug,      // TobDebug
     uLibEcriture,   // TInfoEcriture, CPutDefautEcr, CRemplirDateComptable
     Constantes,     // ets_Pointe
     uLibAnalytique, // ChargeAnalytique
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     UTOF ;

Type

  TPurgeCompta = Class

  private
    FStJournalAno       : string;    // Nom du journal d'A-Nouveaux (SO_JALOUVRE)
    FStJournalOD        : string;
    FStModeSaisieJalOD  : string;
    FStModeSaisieJalAno : string;
    FStExoV8            : string;    //
    FStCptBilanOuv      : string;    // Compte de bilan d'ouverture
    FStCptBilanFer      : string;    // Compte de bilan de fermeture
    FBoSupprCptBilanOuv : Boolean;
    FBoSupprCptBilanFer : Boolean;

    FLastErrorMsg       : string;    // Dernier message d'eereur de l'Objet
    FCodeExo            : TExoDate;  // Code Exercice sur lequel on ramène les ECR
    FDatepurge          : TDateTime; // Date à laquelle on effectue la purge

    FTobEcrInsert           : Tob;
    FTobCptLettrageMultiple : Tob; // Liste des comptes ayant du lettrage multi établissement

    FBoProgressForm : Boolean;   // Purge avec une progressForm ??
    FBoPgiInfo      : Boolean;   // Affichage des Exceptions
    FBoBalSitAuto   : Boolean;   // Génération auto des balances de situation ??
    FBoModifLibelle : Boolean;   // Reprise du code journal et date comptable dans le libellé
    {$IFNDEF CERTIFNF}
    FBoSauvegardeTra: Boolean;   // Sauvegarde du dossier en Fichier TRA
    {$ENDIF}

    FMessageCompta  : TMessageCompta;
    FInfoEcriture   : TInfoEcriture;

    function      VerifieExercices : Boolean;
    function      VerifieLettrage  : Boolean;
    function      VerifieTreso     : Boolean;

    procedure     ControleCptBilan;
    procedure     LancementPurge;
    {$IFNDEF CERTIFNF}
    procedure     SauvegardeDossier;
    {$ENDIF}
    procedure     GenereBalanceSituation;
    procedure     TraitementPieceANO(vTobPiece : Tob);

    // Comptes lettrables
    procedure     GenereDetailAno         ( vTobEcrAno : Tob);
    function      _GenereDetailAnoSimple  ( vTobEcrAno : Tob ) : Double;
    function      _GenereDetailAnoMultiple( vTobEcrAno : Tob; vStCodeLettrage : string ) : Double;

    // Comptes pointables
    procedure     GenerePointage              ( vTobEcrAno : Tob );
    function      _GenereDetailNonPointe      ( vTobEcrAno : Tob ) : Double;
    function      _GenereDetailPointeSupPurge ( vTobEcrAno : Tob ) : Double;
    function      _GenereSoldePointe          ( vTobEcrAno : Tob ) : Double;

    // Comptes ni pointables, ni Lettrables
    procedure     GenereAnoEnSolde( vTobEcrAno : Tob );

    function      SupprimeAnoPgi         ( vTobEcrPiece : Tob ) : Boolean;
    function      SupprimeAnaSurAno      ( vTobEcrPiece : Tob ) : Boolean;

    function      CalculSoldePiecePurge : Boolean;
    function      TraitementPiecePurge  ( vNumeroPiece : integer) : Boolean;
    function      EnregistrePiecePurge  : Boolean;

    function      InitEcrDefautPurge : Tob;

    procedure     AvanceProgressForm( vMsg : string );
    function      PreTraitementJournalAno : Boolean;
  public
    constructor   Create( vDatePurge : TdateTime; vAvecProgressForm : Boolean; vAvecPgiInfo : Boolean );

    procedure     Execute;
    destructor    Destroy; override;

    property      GenereAutoBalSit : Boolean read FBoBalSitAuto    write FBoBalsitAuto;
    property      ModifieLibelle   : Boolean read FBoModifLibelle  write FBoModifLibelle;
    {$IFNDEF CERTIFNF}
    property      SauveGardeTra    : Boolean read FBoSauvegardeTra write FBoSauvegardeTra;
    {$ENDIF}
  end;

////////////////////////////////////////////////////////////////////////////////

  TOF_CPPURGEEXERCICE = Class (TOF)

    FListe          : THGrid;
    DatePurge       : THEdit;
    CBBalSitAuto    : TCheckBox;
    CBModifLibelle  : TCheckBox;

    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnClose                  ; override ;

    procedure OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure     OnClickArchive(Sender: TObject);
    {$IFDEF CERTIFNF}
    procedure     SauvegardeDossier;   // Ajout ME
    {$ENDIF}
  private
    FTobExercice : Tob;
    procedure RemplitFListe;
  end ;

procedure CPLanceFiche_CPPurgeExercice( vStParam : string = '');

Implementation

uses
{$IFDEF MODENT1}
     CPVersion,
     CPProcGen,
{$ENDIF MODENT1}
     Balsit,         // RBalSitInfo
     CritEdt,        // Reel
{$IFDEF NETEXPERT}
     uNetListe,      // GenereIniExport
     uAssistComSx,   // ExportDonnees
{$ENDIF}
     CpteSav,        // RecalculTotPointeNew1
     LettUtil,       // TL_Rappro
     uObjJournalAno, // TObjJournalAno
     uLibWindows,    // IIF
     cbpPath,        // TcbpPath.GetCegidUser
     uLibExercice;   // QuelExoDate

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_CPPurgeExercice( vStParam : string = '');
begin
  AGLLanceFiche('CP', 'CPPURGEEXERCICE', '', '', vStParam );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPURGEEXERCICE.OnArgument (S : String ) ;
begin
  Inherited ;
  FTobExercice := TOB.Create('EXERCICE', nil, -1);

  FListe := THGrid(GetControl('FLISTE', True));
  FListe.OnRowEnter := OnRowEnterFListe;

  DatePurge := THEdit(GetControl('DATEPURGE', True));

  CBBalSitAuto    := TCheckBox(GetControl('CBBALSITAUTO', True));
  CBModifLibelle  := TCheckBox(GetControl('CBMODIFLIBELLE', True));

  FListe.ColAligns[1] := TaCenter;
  FListe.ColAligns[2] := TaLeftJustify;
  FListe.ColAligns[3] := TaCenter;
  FListe.ColAligns[4] := TaCenter;
  FListe.ColAligns[5] := TaCenter;

  if FListe.CanFocus then
    FListe.SetFocus;
{$IFDEF CERTIFNF}
  //Ajout me pour les archivages
  THCheckBox(GetControl('CBARCHIVE')).OnClick := OnClickArchive;
{$ENDIF}

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPURGEEXERCICE.OnLoad ;
begin
  Inherited ;
  RemplitFListe;

  if FTobExercice.Detail.Count <> 0 then
  begin
    DatePurge.Text := FTobExercice.Detail[0].GetString('EX_DATEFIN');
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPPURGEEXERCICE.RemplitFListe;
begin
  FTobExercice.LoadDetailFromSQL('SELECT EX_EXERCICE, EX_LIBELLE, EX_DATEDEBUT, ' +
                                 'EX_DATEFIN, EX_ETATCPTA FROM EXERCICE ORDER BY EX_DATEDEBUT');

  FListe.RowCount := FTobExercice.Detail.Count + 1;

  FTobExercice.PutGridDetail(FListe, False, False, 'EX_EXERCICE;EX_LIBELLE;EX_DATEDEBUT;EX_DATEFIN;EX_ETATCPTA');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPURGEEXERCICE.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPURGEEXERCICE.OnRowEnterFListe(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  DatePurge.Text := FTobExercice.Detail[Ou-1].GetString('EX_DATEFIN');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPPURGEEXERCICE.OnUpdate ;
var lPurgeCompta : TPurgeCompta;
begin
  Inherited ;
  if FTobExercice.Detail[FListe.Row-1].GetString('EX_ETATCPTA') <> 'CDE' then
  begin
    PgiInfo('Traitement impossible sur un exercice non cloturé de manière définitive.','Purge comptable');
    Exit;
  end;

  if FTobExercice.Detail[FListe.Row-1].GetString('EX_EXERCICE') = VH^.Precedent.Code then
  begin
    PgiInfo('Traitement impossible sur l''exercice précédent.','Purge comptable');
    Exit;
  end;

  //06/12/2006 YMO Norme NF
  if PgiAsk('La purge des exercices va effectuer automatiquement une sauvegarde #10#13'
            +'du dossier. Auparavant vous devez avoir édité vos états déclaratifs #10#13'
            +'pour archiver l''ensemble de ces éléments sur un support annexe. #10#13'
            +'Voulez-vous continuer ?', Ecran.Caption) = MrNo then Exit;

  lPurgeCompta := TPurgeCompta.Create( StrToDate(DatePurge.Text), True, True );
  lPurgeCompta.FBoBalSitAuto    := CBBalSitAuto.Checked;
  lPurgeCompta.FBoModifLibelle  := CBModifLibelle.Checked;
{$IFDEF CERTIFNF}
  if (THCheckBox(GetControl('CBARCHIVE')).Checked) then
      SauvegardeDossier;
{$ELSE}
  lPurGeCompta.FBoSauvegardeTra := not (EstSpecif('51207'));
{$ENDIF}
  lPurgeCompta.Execute;
  lPurgeCompta.Destroy;

  RemplitFListe;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/05/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPPURGEEXERCICE.OnClose ;
begin
  FreeAndNil(FTobExercice);
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
{ TPurgeExercice }
////////////////////////////////////////////////////////////////////////////////
constructor TPurgeCompta.Create( vDatePurge : TdateTime; vAvecProgressForm : Boolean; vAvecPgiInfo : Boolean );
var lZExercice : TZExercice;
begin
  FDatePurge      := vDatePurge;
  FBoProgressForm := vAvecProgressForm;
  FBoPgiInfo      := vAvecPgiInfo;
  FBoBalSitAuto   := False;
  FBoModifLibelle := False;
  {$IFNDEF CERTIFNF}
  FBoSauvegardeTra:= False;
  {$ENDIF}
  lZExercice := TZExercice.Create( False );
  FCodeExo   := lZExercice.QuelExoDate( FDatePurge+1 );
  lZExercice.Free;

  FStJournalANO       := GetParamSocSecur('SO_JALOUVRE', '');
  FStJournalOD        := GetParamSocSecur('SO_LETCHOIXJAL', '');
  FStModeSaisieJalOD  := GetColonneSql('JOURNAL', 'J_MODESAISIE', 'J_JOURNAL = "'+ FStJournalOD + '"');
  FStModeSaisieJalAno := GetColonneSql('JOURNAL', 'J_MODESAISIE', 'J_JOURNAL = "'+ FStJournalANO + '"');

  FStCptBilanOuv      := GetParamSocSecur('SO_OUVREBIL', '');
  FStCptBilanFer      := GetParamSocSecur('SO_FERMEBIL', '');

  // Contrôle de l'équilibre du compte de BILAN
  ControleCptBilan;

  FStExoV8 := '';
  if VH^.ExoV8.Code <> '' then
    FStExoV8 := ' E_DATECOMPTABLE >= "' + USDateTime(VH^.ExoV8.Deb) + '" AND ';

  V_PGi.IoError := OeOk;
  FLastErrorMsg := '';

  // Init de l'écriture type générée par la PURGE
  FTobEcrInsert   := Tob.Create('', nil, -1);
  FTobCptLettrageMultiple := Tob.Create('', nil, -1);

  // GCO - 12/01/2007
  FMessageCompta := TMessageCompta.Create('');
  FInfoEcriture  := TInfoEcriture.Create;
  FInfoEcriture.AjouteErrIgnoree([RC_DATEINCORRECTE,
                                  RC_NOGRPMONTANT,
                                  RC_MODEPAIEINCORRECT,
                                  RC_REGIMETVAINCORRECT,
                                  RC_DATEVALEURINCORRECT]);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.Execute;
begin
  if not VerifieExercices then
  begin
    if FBoPgiInfo then
      Pgiinfo('Traitement impossible. ' + FLastErrorMsg, 'Purge Comptable');
    Exit;
  end;

  if not VerifieLEttrage then
  begin
    if FBoPgiInfo then
      Pgiinfo('Traitement impossible. ' + FLastErrorMsg, 'Purge Comptable');
    Exit;
  end;

  if GetParamSocSecur('SO_GENATTEND', '') = '' then
  begin
    if FBoPgiInfo then
      Pgiinfo('Traitement impossible. Vous devez renseigner le compte d''attente ' +
              'dans les paramètres société.', 'Purge Comptable');
    Exit;
  end;

  if not VerifieTreso then
  begin
    if FBoPgiInfo then
      PgiInfo('Traitement impossible. Vous devez d''abord synchroniser la trésorerie.', 'Purge Comptable');
    Exit;
  end;

  if (VH^.ExoV8.Code = '') or (FDatePurge >= VH^.ExoV8.Deb) then
  begin
    if VH^.PointageJal then
    begin
      if FBoPgiinfo then
        PgiInfo('Traitement incompatible avec du pointage sur journal', 'Purge Comptable');
      Exit;
    end;
  end;

  if Transactions(LancementPurge, 0) <> OeOk then
  begin
    if FBoPgiInfo then
      PgiInfo( FLastErrorMsg, 'Purge Comptable');
  end
  else
  begin
    // GCO - 04/11/2005 - FQ 16988 et 16990
    if (VH^.ExoV8.Code = '') or (FCodeExo.Deb > VH^.ExoV8.Deb) then
    begin
      SetParamSoc('SO_EXOV8', FCodeExo.Code);

      // GCO - 23/08/2005 - FQ 16436
      VH^.Exov8.Code              := FCodeExo.Code;
      VH^.Exov8.Deb               := FCodeExo.Deb;
      VH^.Exov8.Fin               := FCodeExo.Fin;
      VH^.Exov8.DateButoir        := FCodeExo.DateButoir;
      VH^.Exov8.DateButoirRub     := FCodeExo.DateButoirRub;
      VH^.Exov8.DateButoirBud     := FCodeExo.DateButoirBud;
      VH^.Exov8.DateButoirBudgete := FCodeExo.DateButoirBudgete;
      VH^.Exov8.NombrePeriode     := FCodeExo.NombrePeriode;
      VH^.Exov8.EtatCpta          := FCodeExo.EtatCpta;
    end;
    SetParamSoc('SO_DATEREFPURGE', FDatePurge + 1);

    // Mise à jour des tablettes sur les exercices
    AvertirMultiTable('TTEXERCICE');
    ChargeMagExo(False);
    
    //06/12/2006 YMO Norme NF  
    { BVE 29.08.07 : Mise en place d'un nouveau tracage }
{$IFNDEF CERTIFNF}
    CPEnregistreLog('PURGE JUSQU''AU '+DateToStr(FDatePurge));
{$ELSE}
    CPEnregistreJalEvent('CPU','Purge des écritures','PURGE JUSQU''AU '+DateToStr(FDatePurge));
{$ENDIF}

    if FBoPgiInfo then
      PgiInfo('Traitement terminé. Aucune anomalie détectée.', 'Purge Comptable');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2006
Modifié le ... : 21/11/2006    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TPurgeCompta.PreTraitementJournalAno : Boolean;
var lObjJournalAno : TObjJournalAno;
begin
  lObjJournalAno := TObjJournalAno.Create;
  try
    lObjJournalAno.StCodeExo := FCodeExo.Code;
    lObjJournalAno.Execute;
    Result := lObjJournalAno.LastError = 0;
  finally
    lObjJournalAno.Free;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TPurgeCompta.ControleCptBilan;
var lQuery : TQuery;
begin
  try
    // Test de l'équilibre du compte du bilan d'ouverture
    lQuery := OpenSql('SELECT SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL FROM ECRITURE WHERE ' +
                      'E_GENERAL = "' + FStCptBilanOuv + '" AND ' +
                      'E_EXERCICE = "' + FCodeExo.Code + '" AND ' +
                      'E_JOURNAL = "' + FStJournalAno + '"', True);

    FBoSupprCptBilanOuv := Arrondi(lQuery.FindField('TOTAL').AsFloat, V_Pgi.OkDecV) = 0;

    if FStCptBilanFer <> FStCptBilanOuv then
    begin
      // Test de l'équilibre du compte du bilan d'ouverture
      lQuery := OpenSql('SELECT SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL FROM ECRITURE WHERE ' +
                        'E_GENERAL = "' + FStCptBilanFer + '" AND ' +
                        'E_EXERCICE = "' + FCodeExo.Code + '" AND ' +
                        'E_JOURNAL = "' + FStJournalAno + '"', True);

      FBoSupprCptBilanFer := Arrondi(lQuery.FindField('TOTAL').AsFloat, V_Pgi.OkDecV) = 0;
    end
    else
      FBoSupprCptBilanFer := FBoSupprCptBilanOuv;

  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/06/2005
Modifié le ... : 23/09/2005 - Ajout de la purge rapide
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.LancementPurge;
var lBoPurgeRapide : Boolean;
    lTobListePiece : Tob;
    lTobPiece      : Tob;
    lStSql         : String;
    i              : integer;
begin
  try
    // GCO - 16/11/2006 - FQ 19088
    if not PreTraitementJournalAno then
    begin
      V_Pgi.IoError := OeUnknown;
      Exit;
    end;

    lTobListePiece := Tob.Create('', nil, -1);
    lTobPiece      := Tob.Create('', nil, -1);

    // Test si purge avant le VH^.Exov8, si oui, on effectue une purge simplifie
    lBoPurgeRapide := (VH^.Exov8.Code <> '') and (FDatePurge <= VH^.ExoV8.Deb);

    if not lBoPurgeRapide then
    begin
      lStSql := 'SELECT DISTINCT E_NUMEROPIECE, E_DEVISE, E_ETABLISSEMENT, ' +
                'E_JOURNAL, E_EXERCICE FROM ECRITURE WHERE ' +
                'E_JOURNAL = "' + FStJournalAno + '" AND ' +
                'E_EXERCICE = "' + FCodeExo.Code + '" AND ' +
                'E_ECRANOUVEAU = "OAN" AND E_CREERPAR <> "DET"';
      lTobListePiece.ClearDetail;
      lTobListePiece.LoadDetailFromSQL(lStSql, False);
    end;

    if FBoProgressForm then
    begin
      if lBoPurgeRapide then
        InitMoveProgressForm(nil, 'Purge Comptable', 'Traitement en cours...', 3, True, True)
      else
        InitMoveProgressForm(nil, 'Purge Comptable', 'Traitement en cours...', lTobListePiece.Detail.Count + 5, True, True);
    end;

    // Création du fichier de Sauvegarde au Format TRA
    {$IFNDEF CERTIFNF}
    if FBoSauvegardeTra then
    begin
      if FBoProgressForm then AvanceProgressForm('Sauvegarde du dossier au format TRA');
      SauvegardeDossier;
      CPEnregistreLog('ARCHIVAGE AVANT PURGE '+DateToStr(FDatePurge)); //06/12/2006 YMO Norme NF
    end;
    {$ENDIF}
    if not lBoPurgeRapide then
    begin
      // Chargement en mémoire des comptes ayant des paquets lettrés Multi ETAB
      // et Multi DEVISE ( un traitement particulier à faire )
      lStSql := 'SELECT E_GENERAL, E_AUXILIAIRE FROM ECRITURE WHERE ' +
        'E_ECRANOUVEAU <> "OAN" ' +
        'AND (E_DATEPAQUETMAX <= "'+ UsDateTime(FDatePurge + 1) + '" AND E_ETATLETTRAGE = "TL") ' +
        'AND E_QUALIFPIECE = "N" ' +
        'AND E_DATECOMPTABLE < "' + UsDateTime(FDatePurge + 1) + '" ' +
        'GROUP BY E_GENERAL, E_AUXILIAIRE ' +
        'HAVING MIN(E_ETABLISSEMENT) <> MAX(E_ETABLISSEMENT) OR ' +
        'MIN(E_DEVISE) <> MAX(E_DEVISE)';

      FTobCptLettrageMultiple.LoadDetailDBFromSQL('', lStSql, False);
      for i := 0 to FTobCptLettrageMultiple.Detail.Count -1 do
      begin
        FTobCptLettrageMultiple.Detail[i].AddChampSupValeur('CODELETTRAGE', GetSetCodeLettre( FTobCptLettrageMultiple.Detail[i].GetString('E_GENERAL'), FTobCptLettrageMultiple.Detail[i].GetString('E_AUXILIAIRE') , ''), False);
      end;

      // Création des Balances de Situation
      if FBoBalSitAuto then
      begin
        GenereBalanceSituation;
        if V_PGI.IoError <> OeOk then Exit;
      end;

      if lTobListePiece.Detail.Count <> 0 then
      begin
        for i := 0 to lTobListePiece.Detail.Count -1 do
        begin
          if V_PGI.IoError <> OeOk then Exit;

          lStSql := 'SELECT E_NUMEROPIECE, E_NUMLIGNE, E_DEVISE, E_ETABLISSEMENT, ' +
                    'E_GENERAL, E_AUXILIAIRE, E_DEBIT, E_CREDIT, ' +
                    'E_DEBITDEV, E_CREDITDEV, E_DEBIT - E_CREDIT TOTAL, ' +
                    'E_DEBITDEV - E_CREDITDEV TOTALDEV, E_REFPOINTAGE, ' +
                    'E_DATEPOINTAGE, G_POINTABLE, E_ETATLETTRAGE, G_LETTRABLE, ' +
                    'T_LETTRABLE FROM ECRITURE ' +
                    'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
                    'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE WHERE ' +
                    'E_ECRANOUVEAU = "OAN" AND ' +
                    'E_NUMEROPIECE = ' + IntToStr(lTobListePiece.Detail[i].GetInteger('E_NUMEROPIECE')) + ' AND ' +
                    'E_JOURNAL = "' + lTobListePiece.Detail[i].GetString('E_JOURNAL') + '" AND ' +
                    'E_EXERCICE = "' + lTobListePiece.Detail[i].GetString('E_EXERCICE') + '" ' +
                    'ORDER BY E_GENERAL, E_AUXILIAIRE, E_ETABLISSEMENT, E_DEVISE';

          lTobPiece.LoadDetailFromSql(lStSql, False);
          TraitementPieceANO( lTobPiece );
        end;
      end;
    end;

    // Retour aux traitements Communs entre les deux types de purge
    if V_PGI.IoError <> OeOk then Exit;
    if FBoProgressForm then AvanceProgressForm('Suppression des écritures antérieures au ' + DateToStr(FDatePurge));
    ExecuteSQL('DELETE FROM ECRITURE WHERE E_DATECOMPTABLE <= "' + UsDateTime(FDatePurge) + '"');
    ExecuteSQL('DELETE FROM ANALYTIQ WHERE Y_DATECOMPTABLE <= "' + UsDateTime(FDatePurge) + '"');

    // GCO - 22/11/2006 - Delete des tables créer par TJA
    ExecuteSQL('DELETE FROM CPTABLEAUVAR WHERE CTV_EXERCICE IN (SELECT EX_EXERCICE FROM EXERCICE WHERE EX_DATEFIN <="'+UsDateTime(FDatePurge)+'")');
    ExecuteSQL('DELETE FROM CPNOTETRAVAIL WHERE CNO_EXERCICE IN (SELECT EX_EXERCICE FROM EXERCICE WHERE EX_DATEFIN <="'+UsDateTime(FDatePurge)+'")');
    // FIN GCO

    ExecuteSQL('DELETE FROM EXERCICE WHERE EX_DATEFIN <= "' + UsDateTime(FDatePurge) + '"');

    if V_PGI.IoError <> OeOk then Exit;
    if FBoProgressForm then AvanceProgressForm('Suppression des anciennes références de pointage');
    ExecuteSql('DELETE FROM EEXBQ WHERE EE_DATEPOINTAGE < "' + UsDateTime(FDatePurge + 1) + '"');
    ExecuteSql('DELETE FROM EEXBQLIG WHERE CEL_DATEPOINTAGE < "' + UsDateTime(FDatePurge + 1) + '"');

  finally
    FreeAndNil(lTobPiece);
    FreeAndNil(lTobListePiece);
    if FBoProgressForm then FiniMoveProgressForm;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2005
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.TraitementPieceANO(vTobPiece: Tob);
var i : integer;
    lStGeneral,lStAuxiliaire, lStDevise, lStEtab : string;
    lBoLettrable : Boolean;
    lBoPointable : Boolean;
    lTobEcrAno : Tob;
    lNumeroPiece : integer;
    lStMessage : string;
    lStSql : string;
    vTobCptSolde : Tob;
begin
  try
    lNumeroPiece := vTobPiece.Detail[0].GetInteger('E_NUMEROPIECE');
    lStDevise    := vTobPiece.Detail[0].GetString('E_DEVISE');
    lStEtab      := vTobPiece.Detail[0].GetString('E_ETABLISSEMENT');
    lStMessage   := IntToStr(lNumeroPiece) + ' - ' + lStDevise + ' - ' + lStEtab;


    // Recherche si l'exercice où l'on génère les écritures, à des comptes lettrables
    // soldés mais avec des paquets lettrés à cheval sur la date de purge.
    // Si oui, il faut généré des A-Nouveaux fictifs pour que la purge, récupère
    // l'antérieurité afin que le lettrage ne se retrouve pas déséquilibré.
    lStSql := 'SELECT E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_ETABLISSEMENT, ' +
              'SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL, SUM(E_DEBITDEV)-SUM(E_CREDITDEV) TOTALDEV ' +
              'FROM ECRITURE LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL WHERE ' +
              '((G_COLLECTIF = "-" AND G_LETTRABLE = "X") OR (G_COLLECTIF = "X")) AND ' +
              'E_DATECOMPTABLE < "' + UsDateTime(FDatePurge + 1) + '" AND ' +
              'E_ETABLISSEMENT = "' + lStEtab + '" AND ' +
              'E_QUALIFPIECE = "N" AND ' +
              'E_DEVISE = "' + lStDevise + '" AND ' +
              '(E_ETATLETTRAGE = "AL" OR ' +
              '(E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "' + UsDateTime(FDatePurge + 1) + '")) AND ' +
              '(E_GENERAL||E_AUXILIAIRE NOT IN ' +
              '(SELECT E_GENERAL||E_AUXILIAIRE FROM ECRITURE WHERE E_JOURNAL = "'+ FStJournalAno + '" AND ' +
              'E_DATECOMPTABLE = "' + UsDateTime(FDatePurge + 1) +'")) ' +
              'GROUP BY E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_ETABLISSEMENT';

    vTobCptSolde := Tob.Create('', nil, -1);
    vTobCptSolde.LoadDetailFromSQL(lStSql, True, False);

    if vTobCptSolde.Detail.Count <> 0 then
    begin
      vTobCptSolde.Detail[0].AddchampSupValeur('E_ETATLETTRAGE', 'AL', True);
      vTobCptSolde.Detail[0].AddchampSupValeur('G_LETTRABLE', 'X', True);
      vTobCptSolde.Detail[0].AddChampSupValeur('E_NUMEROPIECE', IntToStr(lNumeroPiece), True);
    end;

    while vTobCptSolde.Detail.Count <> 0 do
    begin
      vTobCptSolde.Detail[0].ChangeParent(vTobPiece, -1);
    end;
    FreeAndNil(vTobCptSolde);

    vTobPiece.Detail.sort('E_GENERAL;E_AUXILIAIRE');
    //TobDebug(vTobPiece);
    for i := 0 to vTobPiece.Detail.Count -1 do
    begin
      if V_Pgi.IoError = OeUnknown then Exit;

      lTobEcrAno := vTobPiece.Detail[i];

      lStGeneral    := lTobEcrAno.GetString('E_GENERAL');
      lStAuxiliaire := lTobEcrAno.GetString('E_AUXILIAIRE');

      lBoLettrable  := (lTobEcrAno.GetString('E_ETATLETTRAGE') <> 'RI') and
                       ((lTobEcrAno.GetString('G_LETTRABLE') = 'X') or
                       (lTobEcrAno.GetString('T_LETTRABLE') = 'X'));

      lBoPointable  := lTobEcrAno.GetString('G_POINTABLE') = 'X';

      if FBoProgressForm then
        AvanceProgressForm('Traitement de l''A-Nouveaux : ' +
                          IntToStr(lTobEcrAno.GetInteger('E_NUMEROPIECE')) + ' - ' +
                          lStGeneral + ' ' +
                          lStAuxiliaire + ' - ' +
                          lStDevise + ' - ' +
                          lStEtab);

      if (lStGeneral = FStCptBilanOuv) and (FBoSupprCptBilanOuv) then
        Continue;

      if (lStGeneral = FStCptBilanFer) and (FBoSupprCptBilanFer) then
        Continue;

      if lBoLettrable then
      begin
        GenereDetailAno( lTobEcrAno  );
      end
      else
      begin
        if lBoPointable then
          GenerePointage( lTobEcrAno )
        else
          GenereAnoEnSolde( lTobEcrAno );
      end;
    end;
    // Fin de génération de la pièce

    //FTobEcrInsert.Detail.Sort('E_GENERAL;E_AUXILIAIRE');
    //TobDebug( FTobEcrInsert );

    if FBoProgressForm then AvanceProgressForm('Contrôle de l''équilibre de la pièce généréé');

    if not CalculSoldePiecePurge then
    begin
      raise Exception.Create(FLastErrorMsg);
    end;  

    if FBoProgressForm then AvanceProgressForm('Traitement sur la pièce générée par la purge : ' + lStMessage);
    if not TraitementPiecePurge( lNumeroPiece ) then
    begin
      raise Exception.Create(FLastErrorMsg);
    end;

    if FBoProgressForm then AvanceProgressForm('Suppression des A-Nouveaux : ' + lStMessage);
    if not SupprimeAnoPgi( vTobPiece ) then
    begin
      raise Exception.Create(FLastErrorMsg);
    end;

    if FBoProgressForm then AvanceProgressForm('Insertion des A-Nouveaux de la purge : ' + lStMessage);
    if not EnregistrePiecePurge then
    begin
      raise Exception.Create(FLastErrorMsg);
    end;

  except
    on E: Exception do
    begin
      V_Pgi.IoError := OeUnknown;
      FLastErrorMsg := 'Procedure TraitementPieceANO - Erreur ' + E.Message;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/09/2005
Modifié le ... :   /  /
Description .. : Vérifie que les exercices à purger sont équilibrés
Suite ........ : Utilisation du principe de la CLOTURE
Mots clefs ... :
*****************************************************************}
function TPurgeCompta.VerifieExercices : Boolean;
var lQueryExercice : TQuery;
    lQuery : TQuery;
    lNbDecimale : integer;
begin
  Result := True;
  lNbDecimale := GetParamSocSecur('SO_DECVALEUR', 2);
  try
    lQueryExercice := OpenSql('SELECT EX_EXERCICE, EX_DATEDEBUT, EX_LIBELLE FROM EXERCICE WHERE ' +
                               'EX_DATEFIN < "' + DateToStr(FDatePurge+1) + '"', True);

    if not lQueryExercice.Eof then
    begin
      if Result = True then
      begin
        lQuery := OpenSql('SELECT SUM(E_DEBIT), SUM(E_CREDIT) FROM ECRITURE WHERE ' +
                  'E_EXERCICE = "' + lQueryExercice.FindField('EX_EXERCICE').AsString + '" AND ' +
                  'E_QUALIFPIECE = "N"', True);

        if Arrondi( lQuery.Fields[0].AsFloat - lQuery.Fields[1].AsFloat, lNbDecimale ) <> 0 then
        begin
          if VH^.ExoV8.Code = '' then
          begin
            FLastErrorMsg := 'L''exercice ' + lQueryExercice.FindField('EX_EXERCICE').AsString + ' - ' +
                           lQueryExercice.FindField('EX_LIBELLE').AsString + ' n''est pas équilibré.';
            Result := False;
          end
          else
          begin
            // Si problème d'équilibre montant, on autorise quand la même la purge
            // si la date de début de l'exercice est inférieure à l'EXOV8
            if (lQueryExercice.FindField('EX_DATEDEBUT').AsDateTime) >= VH^.ExoV8.Deb then
              Result := False;
          end;
        end;
        Ferme( lQuery );
      end;
      lQueryExercice.Next;
    end;
  finally
    Ferme(lQueryExercice);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/09/2005
Modifié le ... : 09/01/2007
Description .. :
Suite ........ : 1) Ajout du controle si on purge avant EXOV8
Suite ........ : 2) Si purge après ExoV8, le controle des paquets lettrés
Suite ........ : ne se fait qu'à partir de l'exoV8 ( antérieurité inutile )
Mots clefs ... : 
*****************************************************************}
function TPurgeCompta.VerifieLettrage: Boolean;
var lStExoV8 : string;
begin
  Result := True;

  // Si la date de purge est inférieure à l'ExoV8, on considère que le lettrage est
  // bon, étant donné qu'on se fou de l'antérieurité.
  lStExoV8 := '';
  if Vh^.ExoV8.Code <> '' then
  begin
    lStExoV8 := ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '"';
    if FDatePurge < VH^.ExoV8.Deb then Exit;
  end;

  // Sinon contrôle de l'équilibre des paquets lettrés
  if ExisteSQL('SELECT E_AUXILIAIRE, E_GENERAL FROM ECRITURE WHERE ' +
               'E_ETATLETTRAGE = "TL" ' + IIF(VH^.ExoV8.Code <> '', lStExoV8, '') +
               'GROUP BY E_AUXILIAIRE, E_GENERAL ' +
               'HAVING SUM(E_DEBIT) <> SUM(E_CREDIT)') then
  begin
    Result := False;
    FLastErrorMsg := 'Le lettrage du dossier est déséquilibré.';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TPurgeCompta.VerifieTreso: Boolean;
begin
  Result := True;
  if EstComptaTreso then
  begin
    if ExisteSQL('SELECT E_TRESOSYNCHRO FROM ECRITURE WHERE ' +
                 '(E_TRESOSYNCHRO = "' + ets_Pointe + '"  OR ' +
                  'E_TRESOSYNCHRO = "' + ets_Lettre + '"  OR ' +
                  'E_TRESOSYNCHRO = "' + ets_Nouveau + '" OR ' +
                  'E_TRESOSYNCHRO = "' + ets_BqPrevi + '") AND ' +
                  'E_DATECOMPTABLE <= "' + UsDateTime(FDatePurge+1) + '"') then
      Result := False;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF CERTIFNF}
procedure TPurgeCompta.SauvegardeDossier;
{$ELSE}
procedure TOF_CPPURGEEXERCICE.SauvegardeDossier;
{$ENDIF}
{$IFDEF NETEXPERT}
{$IFNDEF CERTIFNF}
var
LigneCommande : string;
{$ENDIF}
{$ENDIF}
begin
{$IFDEF NETEXPERT}
{$IFNDEF CERTIFNF}
  // Sauvegarde du Dossier avant la Purge  fiche 17046
  LigneCommande := '/TRF=EXPORT;'+V_PGI.Nodossier+';;;;X;S5;DOS;'+ TcbpPath.GetCegidUserDocument +'\PGI'+V_PGI.Nodossier+'.TRA';
  ExportDonnees(LigneCommande+';EXPORT;Minimized',TRUE) ;
{$ELSE}
   EnvoiExportParDate (FTobExercice.Detail[0].GetString('EX_DATEDEBUT'), FTobExercice.Detail[0].GetString('EX_DATEFIN'), FTobExercice.Detail[0].GetString('EX_EXERCICE'),  'PUR');
{$ENDIF}
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.GenereBalanceSituation;
var lBalSit : TBalSit;
    lBalSitInfo : RBalSitInfo;
    lTobExercice : Tob;
    i : integer;
    lJour1, lMois1, lAnnee1 : Word;
    lJour2, lMois2, lAnnee2 : Word;
    lStCodeBalSit : string;
begin
  lTobExercice := Tob.Create('', nil, -1);
  try
    try
      lTobExercice.LoadDetailDBFromSQL('EXERCICE', 'SELECT EX_EXERCICE, ' +
        'EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_DATEFIN <= "' +
        UsDateTime(FDatePurge) + '" ORDER BY EX_DATEFIN', False);

      for i := 0 to lTobExercice.Detail.Count-1 do
      begin
        DecodeDate( lTobExercice.Detail[i].GetDateTime('EX_DATEDEBUT'), lAnnee1, lMois1, lJour1);
        DecodeDate( lTobExercice.Detail[i].GetDateTime('EX_DATEFIN'),   lAnnee2, lMois2, lJour2);
        if lAnnee1 = lAnnee2 then
          lStCodeBalSit := IntToStr(lAnnee1)
        else
          lStCodeBalSit := IntToStr(lAnnee1) + '/' + intToStr(lAnnee2);


        FillChar(lBalSitInfo, SizeOf(lBalSitInfo), #0);
        with lBalSitInfo do
        begin
          Libelle       := TraduireMemoire('Année ' + lStCodeBalSit);
          Abrege        := Libelle;
          Plan          := 'GEN';
          DateInf       := lTobExercice.Detail[i].GetDateTime('EX_DATEDEBUT');
          DateSup       := lTobExercice.Detail[i].GetDateTime('EX_DATEFIN');
          CumExo        := False;
          Exo           := lTobExercice.Detail[i].GetString('EX_EXERCICE');
          Ano           := True;
          TypeEcr[Reel] := True;
          DebCreColl    := True;
        end;

        lBalSit := TBalSit.Create( lStCodeBalSit );
        if not lBalSit.Existe then
          lBalSit.GenereAuto( lBalSitInfo );
        lBalSit.Free;
      end;

    except
      on E: Exception do
      begin
        V_Pgi.IoError := OeUnknown;
        FLastErrorMsg := 'Procedure GenereBalanceSituation - Erreur ' + E.Message;
      end;
    end;
  finally
    if Assigned( lTobExercice ) then FreeAndNil( lTobExercice );
    //if Assigned( lBalSit ) then FreeAndNil( lBalSit );
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.AvanceProgressForm( vMsg : string );
begin
  if not MoveCurProgressForm( vMsg ) then
  begin
    V_Pgi.IOError := OeUnknown;
    FLastErrorMsg := 'Arrêt du traitement par l''utilisateur';
  end;
  Application.ProcessMessages;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/06/2005
Modifié le ... :   /  /
Description .. : Transforme les écritures OAN SOLDE du journal d'ANO en écriture
Suite ........ : typées H des comptes non pointables
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.GenereAnoEnSolde( vTobEcrAno : Tob );
var lTob : Tob;
begin
  try
    try
      lTob := Tob.Create('', nil,-1);
      lTob.LoadDetailDBFromSQL('ECRITURE', 'SELECT * FROM ECRITURE WHERE ' +
        'E_EXERCICE = "' + FCodeExo.Code + '" AND ' +
        'E_JOURNAL = "' + FStJournalAno + '" AND ' +
        'E_NUMEROPIECE = ' + vTobEcrAno.GetString('E_NUMEROPIECE') + ' AND ' +
        'E_NUMLIGNE = ' + vTobEcrAno.GetString('E_NUMLIGNE') + ' AND ' +
        'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ' +
        'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILIAIRE') + '" AND ' +
        'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
        'E_DEVISE = "' + vTobEcrAno.GetString('E_DEVISE') + '" AND ' +
        'E_ECRANOUVEAU = "OAN" AND ' +
        'E_DATECOMPTABLE = "' + UsDateTime(FDatePurge+1) + '" ' +
        'ORDER BY E_GENERAL', False);

      while lTob.Detail.Count <> 0 do
      begin
        lTob.Detail[0].SetString('E_ECRANOUVEAU', 'H');
        lTob.Detail[0].ChangeParent( FTobEcrInsert , -1);
      end

    except
      on E: Exception do
      begin
        V_Pgi.IoError := OeUnknown;
        FLastErrorMsg := 'Procedure GenereAnoSolde - ' + E.Message;
      end;
    end;
  finally
    if Assigned(lTob) then FreeAndNil(lTob);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TPurgeCompta.InitEcrDefautPurge : Tob;
begin
  Result := Tob.Create('ECRITURE', nil, -1);
  // Champs par défaut du Noyau
  CPutDefautEcr(Result);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/09/2005
Modifié le ... :   /  /
Description .. : Fonction qui génère le détail A-Nouveaux des comptes lettrables
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.GenereDetailAno( vTobEcrAno : Tob );
var lTobTemp : Tob;
    lSoldeAnoPGI : Double;
    lSoldeAnoPurge : Double;
begin
  try
    if vTobEcrAno.GetString('E_DEVISE') <> V_PGI.DevisePivot then
      lSoldeAnoPgi := vTobEcrAno.GetDouble('TOTALDEV')
    else
      lSoldeAnoPgi := vTobEcrAno.GetDouble('TOTAL');

    lSoldeAnoPurge := _GenereDetailAnoSimple( vTobEcrAno );

    // Traitement si Lettrage MultiEtablissement, MultiDevise
    if FTobCptLettrageMultiple.Detail.Count > 0 then
    begin
      lTobTemp := FTobCptLettrageMultiple.FindFirst(['E_GENERAL','E_AUXILIAIRE'],
                  [vTobEcrAno.GetString('E_GENERAL'),
                   vTobEcrAno.GetString('E_AUXILIAIRE')], False);

      if lTobTemp <> nil then
        lSoldeAnoPurge := lSoldeAnoPurge +_GenereDetailAnoMultiple( vTobEcrAno, lTobTemp.GetString('CODELETTRAGE'));
    end;

    if Arrondi(lSoldeAnoPurge-lSoldeAnoPGI, V_PGI.OkDecV) <> 0 then
    begin
      raise Exception.Create('Ecart sur le compte ' +
            vTobEcrAno.GetString('E_GENERAL') + ' - ' +
            vTobEcrAno.GetString('E_AUXILIAIRE') + #13#10 +
            'Etablissement : ' + vTobEcrAno.GetString('E_ETABLISSEMENT') + ' - ' +
            'Devise : ' + vTobEcrAno.GetString('E_DEVISE') + #13#10 +
            'Solde ANO : ' + StrFMontant(lSoldeAnoPgi, 13, V_Pgi.OkDecV, '', True) + ' - ' +
            'Detail A-Nouveaux : ' + StrFMontant(lSoldeAnoPurge, 13, V_Pgi.OkDecV, '', True));
    end;

  except
    on E: Exception do
    begin
      V_Pgi.IOError := oeUnknown;
      FLastErrorMsg := 'Procedure GenereDetailAno - Erreur ' + E.Message;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/10/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
function TPurgeCompta._GenereDetailAnoSimple(vTobEcrAno : Tob) : Double;
var lStGeneral    : string;
    lStAuxiliaire : string;
    lStDevise     : string;
    lStWhere      : string;
    lStLibelle    : string;
    lTob          : Tob;
    lTobTemp      : Tob;
begin
  Result := 0;
  try
    lStGeneral := '';
    if vTobEcrAno.GetString('E_GENERAL') <> '' then
      lStGeneral := 'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ';

    lStAuxiliaire := '';
    if vTobEcrAno.GetString('E_AUXILIAIRE') <> '' then
      lStAuxiliaire := 'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILIAIRE') + '" AND ';

    lStDevise := vTobEcrAno.GetString('E_DEVISE');

    lStWhere := 'E_ECRANOUVEAU <> "OAN" AND ' +
      'E_QUALIFPIECE = "N" AND ' + FStExoV8 +
      'E_DATECOMPTABLE < "' + USDateTime(FDatePurge+1) + '" AND ' +
      'E_DEVISE = "'+ lStDevise + '" AND ' +
      'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
      lStGeneral + lStAuxiliaire +
      '((E_DATEPAQUETMAX >= "' + USDateTime(FDatePurge+1) + '" AND E_ETATLETTRAGE = "TL") ' +
      'OR E_ETATLETTRAGE = "AL" OR E_ETATLETTRAGE = "PL")';

    // Chargement des écritures qui représentent le détail des ANO
    lTob := Tob.Create('', nil, -1);
    lTob.LoadDetailDBFromSQL('ECRITURE', 'SELECT * FROM ECRITURE WHERE ' + lStWhere, True);

    while lTob.Detail.Count <> 0 do
    begin
      lTobTemp := lTob.Detail[0];

      if FBoModifLibelle then
      begin
        lStLibelle := lTobTemp.GetString('E_JOURNAL') + ' ' +
                      DateToStr(lTobTemp.GetDateTime('E_DATECOMPTABLE')) + ' ' +
                      lTobTemp.GetString('E_LIBELLE');

        lStLibelle := Copy(lStLibelle, 1, 35);
        lTobTemp.SetString('E_LIBELLE', lStLibelle);
      end;

      if lStDevise <> V_PGI.DevisePivot then
        Result := Result + lTobTemp.GetDouble('E_DEBITDEV') -
                           lTobTemp.GetDouble('E_CREDITDEV')
      else
        Result := Result + lTobTemp.GetDouble('E_DEBIT') -
                           lTobTemp.GetDouble('E_CREDIT');

      lTobTemp.ChangeParent( FTobEcrInsert , -1);
    end;

    if Assigned(lTob) then FreeAndNil(lTob);
  except
    on E: Exception do
    begin
      V_Pgi.IOError := oeUnknown;
      FLastErrorMsg := 'Function _GenereDetailAnoSimple - ' + E.Message;
    end;
  end;
end;

//////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TPurgeCompta._GenereDetailAnoMultiple( vTobEcrAno : Tob; vStCodeLettrage : string ) : Double;
var lQuery  : TQuery;
    lTobEcr : Tob;
    lDevise : RDevise;
    lSolde  : Double;
    lCouverture    : Double;
    lCouvertureDev : Double;
begin
  try
    Result := 0;

    lQuery := OpenSQL('SELECT SUM(E_DEBIT) - SUM(E_CREDIT) TOTAL, ' +
              'SUM(E_DEBITDEV) - SUM(E_CREDITDEV) TOTALDEV ' +
              'FROM ECRITURE WHERE E_ECRANOUVEAU <> "OAN" AND ' +
              'E_QUALIFPIECE = "N" AND ' +
              'E_DATECOMPTABLE < "' + UsDateTime(FDatePurge+1) + '" AND ' +
              'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ' +
              'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILAIRE') + '" AND ' +
              'E_DEVISE = "' + vTobEcrAno.GetString('E_DEVISE') + '" AND ' +
              'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
              '(E_DATEPAQUETMAX < "' + UsDateTime(FDatePurge + 1) + '" AND ' +
              'E_ETATLETTRAGE = "TL")', True);

    lDevise.Code := vTobEcrAno.GetString('E_DEVISE');
    GetInfosDevise(lDevise);
    lDevise.Taux := GetTaux(lDevise.Code , lDevise.DateTaux, V_PGI.DateEntree) ;

    if lDevise.Code <> V_PGI.DevisePivot then
      lSolde := lQuery.FindField('TOTALDEV').AsFloat
    else
      lSolde := lQuery.FindField('TOTAL').AsFloat;

    if lSolde <> 0 then
    begin
      lTobEcr := InitEcrDefautPurge;
      lTobEcr.SetString('E_GENERAL',         vTobEcrAno.GetString('E_GENERAL'));
      lTobEcr.SetString('E_AUXILIAIRE',      vTobEcrAno.GetString('E_AUXILIAIRE'));
      lTobEcr.SetString('E_EXERCICE',        FCodeExo.Code);
      lTobEcr.SetDateTime('E_DATECOMPTABLE', FDatePurge + 1);
      lTobEcr.SetDateTime('E_DATEMODIF',     FDatePurge + 1);
      lTobEcr.SetDateTime('E_DATEPAQUETMIN', FDatePurge + 1);
      lTobEcr.SetDateTime('E_DATEPAQUETMAX', FDatePurge + 1);
      lTobEcr.SetString('E_LETTRAGE',        vStCodeLettrage);
      lTobEcr.SetString('E_ETABLISSEMENT',   vTobEcrAno.GetString('E_ETABLISSEMENT'));
      lTobEcr.SetString('E_DEVISE',          vTobEcrAno.GetString('E_DEVISE'));
      lTobEcr.SetString('E_ETATLETTRAGE',    'TL');
      lTobEcr.SetString('E_TRESOSYNCHRO',    'LET');
      lTobEcr.SetString('E_MODESAISIE',      FStModeSaisieJalAno);
      lTobEcr.SetString('E_LIBELLE',         'Lettrage Multi établissement/devise');
      lTobEcr.SetInteger('E_NUMEROPIECE',    vTobEcrAno.GetInteger('E_NUMEROPIECE'));
      lTobEcr.SetInteger('E_NUMECHE',        1);

      if lSolde > 0 then
        CSetMontants( lTobEcr, lSolde, 0, lDevise, True)
      else
        CSetMontants( lTobEcr, 0, Abs(lSolde), lDevise, True);

      // Fait après le CSetMontants qui remplit comme il faut E_DEBIT, E_CREDIT
      lCouverture    := Abs(lTobEcr.GetDouble('E_DEBIT') - lTobEcr.GetDouble('E_CREDIT'));
      lCouvertureDev := Abs(lTobEcr.GetDouble('E_DEBITDEV') - lTobEcr.GetDouble('E_CREDITDEV'));

      lTobEcr.SetDouble('E_COUVERTURE', Arrondi(lCouverture, V_PGI.OkDecV)) ;
      lTobEcr.SetDouble('E_COUVERTUREDEV', Arrondi(lCouvertureDev, V_PGI.OkDecV)) ;

      lTobEcr.ChangeParent( FTobEcrInsert, -1);

      // Solde de l'écriture généré qui représente les écritures lettrés en
      // Multi-établissement, Multi Devise
      Result := lSolde;
    end;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TPurgeCompta.SupprimeAnoPgi(vTobEcrPiece : Tob): Boolean;
var lStsql : string;
begin
  Result := False;
  lStSql := 'DELETE FROM ECRITURE WHERE ' +
              'E_EXERCICE = "' + FCodeExo.Code + '" AND ' +
              'E_JOURNAL = "' + FStJournalAno + '" AND ' +
              'E_DEVISE = "' + vTobEcrPiece.Detail[0].GetString('E_DEVISE') + '" AND ' +
              'E_ETABLISSEMENT = "' + vTobEcrPiece.Detail[0].GetString('E_ETABLISSEMENT') + '" AND ' +
              'E_NUMEROPIECE = ' + IntToStr(vTobEcrPiece.Detail[0].GetInteger('E_NUMEROPIECE')) + ' AND ' +
              'E_DATECOMPTABLE = "' + UsDateTime(FDatePurge+1) + '"';

  // Suppression des écritures du Journal d'A-Nouveaux pour la pièce
  if ExecuteSQL( lStSql ) <> 0 then
    Result := True
  else
  begin
    FLastErrorMsg := 'Suppression impossible des écritures du journal d''A-Nouveaux' +#13#10 +
                     'Pièce : ' + IntToStr(vTobEcrPiece.Detail[0].GetInteger('E_NUMEROPIECE')) + ' - ' +
                     'Etablissement : ' + vTobEcrPiece.Detail[0].GetString('E_ETABLISSEMENT') + ' - ' +
                     'Devise : ' + vTobEcrPiece.Detail[0].GetString('E_DEVISE');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :                         
*****************************************************************}
function TPurgeCompta.SupprimeAnaSurAno(vTobEcrPiece: Tob): Boolean;
var lStsql : string;
begin
  Result := False;
  lStSql := 'DELETE FROM ANALYTIQ WHERE ' +
              'Y_EXERCICE = "' + FCodeExo.Code + '" AND ' +
              'Y_JOURNAL = "' + FStJournalAno + '" AND ' +
              'Y_DEVISE = "' + vTobEcrPiece.Detail[0].GetString('E_DEVISE') + '" AND ' +
              'Y_ETABLISSEMENT = "' + vTobEcrPiece.Detail[0].GetString('E_ETABLISSEMENT') + '" AND ' +
              'Y_NUMEROPIECE = ' + IntToStr(vTobEcrPiece.Detail[0].GetInteger('E_NUMEROPIECE')) + ' AND ' +
              'Y_DATECOMPTABLE = "' + UsDateTime(FDatePurge+1) + '"';

  // Suppression des écritures du Journal d'A-Nouveaux pour la pièce
  if ExecuteSQL( lStSql ) >= 0 then
    Result := True
  else
  begin
    FLastErrorMsg := 'Suppression impossible des écritures analytiques du journal d''A-Nouveaux' +#13#10 +
                     'Pièce : ' + IntToStr(vTobEcrPiece.Detail[0].GetInteger('E_NUMEROPIECE')) + ' - ' +
                     'Etablissement : ' + vTobEcrPiece.Detail[0].GetString('E_ETABLISSEMENT') + ' - ' +
                     'Devise : ' + vTobEcrPiece.Detail[0].GetString('E_DEVISE');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function TPurgeCompta.CalculSoldePiecePurge : Boolean;
var lTDC : RecCalcul; // structure de retour de la fonction de calcul du solde
    i : integer;
    lTobEcr : Tob;
begin
  Result  := True;
  lTDC.D  := 0;
  lTDC.C  := 0;
  lTDC.DD := 0;
  lTDC.CD := 0;

  for i := 0 to FTobEcrInsert.Detail.Count -1 do
  begin
    lTobEcr := FTobEcrInsert.Detail[i];
    lTDC.D  := lTDC.D  + lTobEcr.GetValue( 'E_DEBIT' );
    lTDC.C  := lTDC.C  + lTobEcr.GetValue( 'E_CREDIT' );
    lTDC.DD := lTDC.DD + lTobEcr.GetValue( 'E_DEBITDEV' );
    lTDC.CD := lTDC.CD + lTobEcr.GetValue( 'E_CREDITDEV' );
  end;

  if not ( Arrondi ( ( lTDC.DD  - lTDC.CD )  , V_PGI.OkDecV ) = 0 ) then
  begin
    FLastErrorMsg := 'Function CalculSoldePiecePurge - la pièce n''est pas soldée';
    V_PGI.IOError := oESaisie ;
    Result := False;
  end
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/10/2005
Modifié le ... : 15/09/2006
Description .. : LG - 15/09/2006 - FB 18633 - affectation du champs
Suite ........ : E_ENCAISSEMENT
Mots clefs ... :
*****************************************************************}
function TPurgeCompta.TraitementPiecePurge( vNumeroPiece : integer ) : Boolean;
var lRecError       : TRecError;
    i               : integer;
    lTobTemp        : Tob;
begin
  Result := True;
  try
    //Pgiinfo('Recordcount : ' + IntToStr(vTobEcrInsert.Detail.Count), '');
    for i:= 0 to FTobEcrInsert.Detail.Count - 1 do
    begin
      lTobTemp := FTobEcrInsert.Detail[i];

      if lTobTemp.GetString('E_ANA') = 'X' then
      begin // Chargement de l'analytique de la pièce
        CChargeAna( lTobTemp );
      end;

      lTobTemp.SetString('E_NATUREPIECE', 'OD');
      lTobTemp.SetDateTime('E_DATECREATION', Date);
      lTobTemp.SetString('E_QUALIFORIGINE', 'PUR');
      lTobTemp.SetString('E_IO', 'X');
      lTobTemp.SetString('E_TRESOSYNCHRO', 'RIE');
      lTobTemp.SetInteger('E_NUMGROUPEECR', 0);
      lTobTemp.SetInteger('E_NUMEROPIECE', vNumeroPiece);

      {$IFDEF GIL}
      lTobTemp.SetString('E_ENCAISSEMENT', 'RIE');
      {$ENDIF}

      // E_DATECOMPTABLE, E_EXERCICE, E_PERIODE, E_SEMAINE
      CRemplirDateComptable( lTobTemp, FDatePurge + 1);
      // E_ECHE
      lTobTemp.PutValue('E_ENCAISSEMENT',SENSENC(lTobTemp.GetValue('E_DEBIT') , lTobTemp.GetValue('E_CREDIT') ) ) ;
      //CGetRegimeTVA( lTobTemp, FInfoEcriture );
      // E_RIB
      lTobTemp.SetString('E_RIB', cGetRib( lTobTemp));

      lTobTemp.SetDateTime('E_DATEMODIF', FDatePurge + 1);
      lTobTemp.SetString('E_JOURNAL', FStJournalAno);
      lTobTemp.SetString('E_ECRANOUVEAU', 'H');
      lTobTemp.SetString('E_MODESAISIE', FStModeSaisieJalAno);

      if lTobTemp.Detail.Count > 0 then
        CSynchroVentil(lTobTemp);
      //else
      //  CGetAnalytique(lTobTemp, nil);

      lRecError := CIsValidLigneSaisie(lTobTemp, FInfoEcriture);
      if (lRecError.RC_Error <> RC_PASERREUR) then
      begin
        FLastErrorMsg := 'Function TraitementPiecePurge - CIsValidLigneSaisie' + #13#10 +
                         FMessageCompta.GetMessage(lRecError.RC_Error) + #13#10 + lRecError.RC_Message;
        V_PGI.IOError := oESaisie ;
        Result := False;
        Break ;
      end;

    end;

    if not SupprimeAnaSurAno( FTobEcrInsert ) then
    begin
      FLastErrorMsg := 'Function TraitementPiecePurge - SupprimeAnaSurAno';
      Result := False;
      Exit;
    end;

  finally

  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TPurgeCompta.EnregistrePiecePurge: Boolean;
begin
  Result := True;
  if not CEnregistreSaisie( FTobEcrInsert, False, False , True , FInfoEcriture ) then
  begin
    FLastErrorMsg := V_Pgi.LastSQLError;
    Result := False;
  end
  else
    FTobEcrInsert.ClearDetail;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/07/2005
Modifié le ... : 07/10/2005
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TPurgeCompta.Destroy;
begin
  FreeAndNil(FTobEcrInsert);
  FreeAndNil(FTobCptLettrageMultiple);
  FreeAndNil(FMessageCompta);
  FreeAndNil(FInfoEcriture);
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/10/2005
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TPurgeCompta.GenerePointage(vTobEcrAno: Tob);
var lDbTotA : double;
    lDbTotB : double;
    lDbTotC : double;
    lSoldeAnoPgi : Double;
    lSoldeAnoPurge : Double;
begin
  lDbTotA := _GenereDetailNonPointe( vTobEcrAno );

  lDbTotB := _GenereDetailPointeSupPurge( vTobEcrAno );

  lDbTotC := _GenereSoldePointe( vTobEcrAno );

  lSoldeAnoPurge := lDbTotA + lDbTotB + lDbTotC;

  if vTobEcrAno.GetString('E_DEVISE') <> V_PGI.DevisePivot then
    lSoldeAnoPgi := vTobEcrAno.GetDouble('TOTALDEV')
  else
    lSoldeAnoPgi := vTobEcrAno.GetDouble('TOTAL');

  if Arrondi(lSoldeAnoPgi,2) <> Arrondi(lSoldeAnoPurge, 2) then
  begin
    raise Exception.Create('Ecart sur le compte ' +
          vTobEcrAno.GetString('E_GENERAL') + #13#10 +
          'Etablissement : ' + vTobEcrAno.GetString('E_ETABLISSEMENT') + ' - ' +
          'Devise : ' + vTobEcrAno.GetString('E_DEVISE') + #13#10 +
          'Solde ANO : ' + FloatToStr( lSoldeAnoPgi) + ' - ' +
          'Detail A-Nouveaux : ' + FloatToStr(lSoldeAnoPurge));
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/11/2006
Modifié le ... : 21/11/2006
Description .. : Génère le détail des écritures non pointés dont la date est
Suite ........ : antérieure ou égale à la date de purge.
Suite ........ : Etape 1 de la FQ 18427
Mots clefs ... :
*****************************************************************}
function TPurgeCompta._GenereDetailNonPointe( vTobEcrAno : Tob ) : Double;
var lTob : Tob;
    lStLibelle : string;
    lStSql : string;
begin
  Result := 0;
  try
    try
      lTob := Tob.Create('', nil, -1);
      lStSql := 'SELECT * FROM ECRITURE WHERE ' +
                'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ' +
                'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILIAIRE') + '" AND ' + FStExoV8 +
                'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
                'E_DEVISE = "' + vTobEcrAno.GetString('E_DEVISE') + '" AND ' +
                'E_DATECOMPTABLE < "' + UsDateTime(FDatePurge+1) + '" AND ' +
//                'E_JOURNAL <> "' + FStJournalAno + '" AND E_QUALIFPIECE = "N" AND ' +
                'E_ECRANOUVEAU <> "OAN" AND E_QUALIFPIECE = "N" AND ' +
                '(E_REFPOINTAGE = "" AND E_DATEPOINTAGE = "' + UsDateTime(idate1900) + '") ' +
                'ORDER BY E_GENERAL';

      lTob.LoadDetailDBFromSQL('ECRITURE', lStSql);
      while (lTob.Detail.Count <> 0) do
      begin
        if vTobEcrAno.GetString('E_DEVISE') <> V_PGI.DevisePivot then
          Result := Result + lTob.Detail[0].GetDouble('E_DEBITDEV') -
                             lTob.Detail[0].GetDouble('E_CREDITDEV')
        else
          Result := Result + lTob.Detail[0].GetDouble('E_DEBIT') -
                             lTob.Detail[0].GetDouble('E_CREDIT');

        if FBoModifLibelle then
        begin
          lStLibelle := lTob.Detail[0].GetString('E_JOURNAL') + ' ' +
                        DateToStr(lTob.Detail[0].GetDateTime('E_DATECOMPTABLE')) + ' ' +
                        lTob.Detail[0].GetString('E_LIBELLE');

          lStLibelle := Copy(lStLibelle, 1, 35);
          lTob.Detail[0].SetString('E_LIBELLE', lStLibelle);
        end;
        lTob.Detail[0].ChangeParent( FTobEcrInsert, -1 );
      end;

    except
      on E: Exception do
      begin
        V_Pgi.IoError := OeUnknown;
        FLastErrorMsg := 'Procedure _GenereDetailNonPointe : ' + E.Message;
      end;
    end;
  finally
    if Assigned(lTob) then FreeAndNil(lTob);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/11/2006
Modifié le ... : 21/11/2006
Description .. : Génère le détail des écritures pointés dont la date est
Suite ........ : antérieure ou égale à la date de purge et dont la date de
Suite ........ : pointage est supérieure à la date de purge.
Suite ........ : Etape 2 de la FQ 18427
//                'E_JOURNAL <> "' + FStJournalAno + '" AND E_QUALIFPIECE = "N" AND ' +
Mots clefs ... :
*****************************************************************}
function TPurgeCompta._GenereDetailPointeSupPurge( vTobEcrAno : Tob ) : Double;
var lTob : Tob;
    lStSql : string;
    lStLibelle : string;
begin
  Result := 0;
  try
    try
      lTob := Tob.Create('', nil, -1);
      lStSql := 'SELECT * FROM ECRITURE WHERE ' +
                'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ' +
                'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILIAIRE') + '" AND ' +
                'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
                'E_DEVISE = "' + vTobEcrAno.GetString('E_DEVISE') + '" AND ' +
                'E_DATECOMPTABLE < "' + UsDateTime(FDatePurge+1) + '" AND ' +
                'E_ECRANOUVEAU <> "OAN" AND E_QUALIFPIECE = "N" AND ' +
                FStExoV8 +
                '(E_REFPOINTAGE <> "" AND E_DATEPOINTAGE >= "' + UsDateTime(FDatePurge+1) + '") '+
                'ORDER BY E_GENERAL';

      lTob.LoadDetailDBFromSQL('ECRITURE', lStSql);
      while (lTob.Detail.Count <> 0) do
      begin
        if vTobEcrAno.GetString('E_DEVISE') <> V_PGI.DevisePivot then
          Result := Result + lTob.Detail[0].GetDouble('E_DEBITDEV') -
                             lTob.Detail[0].GetDouble('E_CREDITDEV')
        else
          Result := Result + lTob.Detail[0].GetDouble('E_DEBIT') -
                             lTob.Detail[0].GetDouble('E_CREDIT');

        if FBoModifLibelle then
        begin
          lStLibelle := lTob.Detail[0].GetString('E_JOURNAL') + ' ' +
                        DateToStr(lTob.Detail[0].GetDateTime('E_DATECOMPTABLE')) + ' ' +
                        lTob.Detail[0].GetString('E_LIBELLE');

          lStLibelle := Copy(lStLibelle, 1, 35);
          lTob.Detail[0].SetString('E_LIBELLE', lStLibelle);
        end;
        lTob.Detail[0].ChangeParent( FTobEcrInsert, -1 );
      end;

    except
      on E: Exception do
      begin
        V_Pgi.IoError := OeUnknown;
        FLastErrorMsg := 'Procedure _GenereDetailPointeSupPurge : ' + E.Message;
      end;
    end;
  finally
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/11/2006
Modifié le ... : 21/11/2006
Description .. : Solde des écritures pointées dont la date de pointage et la
Suite ........ : date comptable est antérieure ou égale à la date de purge.
Suite ........ : Etape 3 de la FQ 18427
Mots clefs ... :
*****************************************************************}
function TPurgeCompta._GenereSoldePointe( vTobEcrAno : Tob ) : Double;
var lQuery  : TQuery;
    lStSql  : string;
    lTob    : Tob;
    lDevise : RDevise;
begin
  Result := 0;
  try
    try
      lQuery := OpenSql('SELECT SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL, ' +
                        'SUM(E_DEBITDEV)-SUM(E_CREDITDEV) TOTALDEV ' +
                        'FROM ECRITURE WHERE ' +
                        'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ' +
                        'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILIAIRE') + '" AND ' +
                        'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
                        'E_DEVISE = "' + vTobEcrAno.GetString('E_DEVISE') + '" AND ' +
                        'E_DATECOMPTABLE < "' + UsDateTime(FDatePurge+1) + '" AND ' +
                        'E_ECRANOUVEAU <> "OAN" AND E_QUALIFPIECE = "N" AND ' +
                         FStExoV8 +
                        // GCO - 29/11/2007 - FQ 21794
                        //'(E_REFPOINTAGE <> "" AND E_DATEPOINTAGE <= "' + UsDateTime(FDatePurge+1) + '")', True);
                        '(E_REFPOINTAGE <> "" AND E_DATEPOINTAGE < "' + UsDateTime(FDatePurge+1) + '")', True);

      if vTobEcrAno.GetString('E_DEVISE') <> V_PGI.DevisePivot then
        Result := lQuery.FindField('TOTALDEV').AsFloat
      else
        Result := lQuery.FindField('TOTAL').AsFloat;

      if Arrondi(Result, 2) <> 0 then
      begin // Si des écritures ont été pointées

        // Sélection en base de l'écriture d'A-Nouveaux
        lStSql := 'SELECT ##TOP 1## * FROM ECRITURE WHERE ' +
                  'E_GENERAL = "' + vTobEcrAno.GetString('E_GENERAL') + '" AND ' +
                  'E_AUXILIAIRE = "' + vTobEcrAno.GetString('E_AUXILIAIRE') + '" AND ' +
                  'E_DEVISE = "' + vTobEcrAno.GetString('E_DEVISE') + '" AND ' +
                  'E_ETABLISSEMENT = "' + vTobEcrAno.GetString('E_ETABLISSEMENT') + '" AND ' +
                  'E_EXERCICE = "' + FCodeExo.Code + '" AND ' +
                  'E_JOURNAL = "' + FStJournalAno + '" AND ' +
                  'E_NUMEROPIECE = ' + IntToStr(vTobEcrAno.GetInteger('E_NUMEROPIECE')) + ' AND ' +
                  'E_DATECOMPTABLE = "' + UsDateTime(FDatePurge+1) + '"';

        lTob := Tob.Create('', nil, -1);
        lTob.LoadDetailDBFromSQL('ECRITURE', lStSql , False);

        if lTob.Detail.Count = 1 then
        begin
          lDevise.Code := lTob.Detail[0].GetString('E_DEVISE');
          GetInfosDevise(lDevise);
          lDevise.Taux := GetTaux(lDevise.Code , lDevise.DateTaux, V_PGI.DateEntree) ;

          lTob.Detail[0].SetDateTime('E_DATEPOINTAGE', FDatePurge);
          lTob.Detail[0].PutValue('E_REFPOINTAGE', TraduireMemoire('POINTAGE'));

          lTob.Detail[0].SetString('E_LIBELLE', TraduireMemoire('Total des écritures pointées'));
          if lQuery.FindField('TOTAL').AsFloat > 0 then
            CSetMontants( lTob.Detail[0], lQuery.FindField('TOTAL').AsFloat, 0, lDevise, True)
          else
            CSetMontants( lTob.Detail[0], 0, Abs(lQuery.FindField('TOTAL').AsFloat), lDevise, True);
          lTob.Detail[0].ChangeParent( FTobEcrInsert, -1);
        end
        else
          raise Exception.Create('Ecriture d''A-Nouveaux du compte ' +
                                 vTobEcrAno.GetString('E_GENERAL') + ' non trouvé.');
      end;
    except
      on E: Exception do
      begin
        V_Pgi.IoError := OeUnknown;
        FLastErrorMsg := 'Procedure _GenereSoldePointe : ' + E.Message;
      end;
    end;
  finally
    Ferme( lQuery );
  end;
end;

procedure TOF_CPPURGEEXERCICE.OnClickArchive(Sender: TObject);
begin
 if not THCheckBox(GetControl('CBARCHIVE')).Checked then
 begin
       If PGIAsk ('En référence au BOI 13 L-1-06 N° 12 du 24 Janvier 2006 Paragraphe 95,'+#10#13+
       ' il est préconisé d''archiver l''ensemble des informations dont la conservation est obligatoire.'+#10#13+
       ' Vous avez choisi de ne pas effectuer un archivage automatique de ces données, la responsabilité '+#10#13+
       ' de le réaliser manuellement vous incombe. Confirmez-vous ce choix ?')  <> mrYes then
               THCheckBox(GetControl('CBARCHIVE')).Checked := TRUE;
 end;
end;


////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPPURGEEXERCICE ] ) ;
end.

