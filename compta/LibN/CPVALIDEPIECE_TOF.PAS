{***********UNITE*************************************************
Auteur  ...... : SBO
Créé le ...... : 25/11/2005
Modifié le ... : 25/11/2005
Description .. : Source TOF de la FICHE : VALSIMU
Mots clefs ... : TOF;VALSIMU
*****************************************************************}
Unit CPVALIDEPIECE_TOF;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     FE_Main, db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul, HDB,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, HTB97,
     utilPGI, SAISUTIL, Ent1, ED_Tools, SaisComm, Saisie, TiersPayeur, LettUtil, HStatus, DelVisuE,
     uTob,
     uLibAnalytique,
     uLibEcriture,
     UtilSais,          //    MajSoldesEcritureTOB
     ParamSoc           //    GetParamSocSecur
     ;

procedure CPLanceFiche_ValidePiece(vStRange, vStLequel, vStArgs : string);

Type
  TOF_CPVALIDEPIECE = Class (TOF)
    procedure OnNew                   ; override;
    procedure OnDelete                ; override;
    procedure OnUpdate                ; override;
    procedure OnLoad                  ; override;
    procedure OnArgument (S : String ); override;
    procedure OnDisplay               ; override;
    procedure OnClose                 ; override;
    procedure OnCancel                ; override;
  private

    // Indicateurs
    FModeGener       : String ;
    NowFutur         : TDateTime;       // Date/Heure de modif pour les pièces
    FPiecesGener     : TList;           // Liste des pièces générées
    FPiecesErr       : TList;           // Liste des pièces refusées
    Messages         : TStringList;     // Liste des messages
    FBoChargEnCours  : boolean;         // FQ 18696 : Indicateur de chargement de l'interface en cours (pour éviter 1ère recherche...)
    fbControleRefGescom : boolean;
    // Composants
    E_EXERCICE       : THValComboBox;
    E_DATECOMPTABLE  : THEdit;
    E_DATECOMPTABLE_ : THEdit;
    E_QUALIFPIECE    : THValComboBox;
    E_JOURNAL        : THValComboBox;
    E_NATUREPIECE    : THValComboBox;

    QualifGener      : THValComboBox;
    CChoixDate       : TCheckBox;
    DateGeneration   : THEdit;

    {$IFDEF EAGLCLIENT}
      FListe         : THGrid;
    {$ELSE}
      FListe         : THDBGrid;
    {$ENDIF}

    // initialisation
    procedure InitComposants ;
    procedure InitEvenements ;
    procedure InitMessages;
    procedure InitVariables;
    procedure InitPresentation;

    // Fonctions de tests
    function  DateGeneOk : boolean;
    function  AfficheMessage ( vInErr : Integer ) : word ;

    // Traitements
    function  ValideSelection : boolean;
    procedure ValideLaPiece;

    procedure TraiteTPVal;

    // Evènements de composants
    procedure FListeDblClick     ( Sender : TObject ) ;
    procedure BOuvrirClick       ( Sender : TObject ) ;
    procedure BZoomPieceClick    ( Sender : TObject ) ;
    procedure CChoixDateClick    ( Sender : TObject ) ;
    procedure ExerciceChange     ( Sender : TObject ) ;
    procedure QualifPieceChange  ( Sender : TObject ) ;
    procedure BChercheClick      ( Sender : TObject ) ; // FQ 18696
    procedure OnAfterShow;                              // FQ 18696
    procedure AuxiElipsisClick   ( Sender : TObject ) ;

  end;


function CTransformerPiece( vStOrig, vStDest : String ; vTobEcr : TOB ; vBoMajTob : Boolean ; vBoDateForce : Boolean = False ; vDtDateC : TDateTime = 0 ) : Boolean ;

Implementation

uses
    {$IFDEF MODENT1}
    CPVersion,
    CPProcGen,
    ULibExercice,
    {$ENDIF MODENT1}
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    , UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }



/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_ValidePiece(vStRange, vStLequel, vStArgs : string);
begin

  if _Blocage(['nrCloture','nrBatch','nrSaisieModif'],False,'nrBatch') then Exit ;

  AGLLanceFiche ('CP','CPVALIDEPIECE', vStRange, vStLequel, vStArgs );

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
function CTransformerPiece( vStOrig, vStDest : String ; vTobEcr : TOB ; vBoMajTob : Boolean ; vBoDateForce : Boolean ; vDtDateC : TDateTime ) : Boolean ;
Var lTOBPiece  : TOB ;          // Tob de la pièce
    lTOBEcr    : TOB ;          // Tob sur ligne d'écriture
    lTOBVentil : TOB ;          // Tob sur ligne d'analytique
    lQEcr      : TQuery ;       // req de chargement de la pièce
    lStReq     : String ;       // requêtes
    lNumPiece  : integer ;      // Numéro de pièce
    lNumL      : integer ;      // Numéro de ligne
    lNumV      : integer ;      // Numéro d'axe
    lNumA      : integer ;      // Numéro de ventilation
    lInNbEnreg : LongInt ;      // Nb d'enregistrement
    lStTrSyn   : String ;       // TR synchro
    lDtModif   : TDateTime ;    // Date / Heure de modif
    lInErr     : Integer ;
    lStRef     : String ;
    lInfo      : TInfoEcriture ;
    lTobAna    : Tob ;
    iAxe       : Integer ;
    iVentil    : Integer ;
    lStExo     : string ;
    lInPer     : Integer ;
    lInSem     : Integer ;
begin

  Result := False ;

  if vTobEcr = nil then Exit ;

  // ================================
  // ==== Chargement de la pièce ====
  // ================================
  lTOBPiece := TOB.Create('',Nil,-1) ;
  lQEcr     := OpenSQL('SELECT * FROM ECRITURE WHERE ' + WhereEcritureTob( tsGene, vTobEcr, False ) , True ) ;
  lTOBPiece.LoadDetailDB('ECRITURE','','', lQEcr, False ) ;
  Ferme(lQEcr) ;
  if lTOBPiece.Detail.Count <= 0 then
    begin
    lTOBPiece.Free ;
    Exit ;
    end ;
  ChargeAnalytique( lTOBPiece ) ;

  // Tests préalables
  lInfo := TInfoEcriture.Create ;
  lInErr := RC_PASERREUR ;
  for lNumL := 0 to ( lTOBPiece.Detail.Count - 1 ) do
    begin
    lTobEcr := lTobPiece.Detail[lNumL] ;
    // Tiers
    if ( lInErr = RC_PASERREUR ) and lInfo.LoadAux( lTobEcr.GetString('E_AUXILIAIRE') )
                                 and ( lInfo.Aux_GetValue('T_FERME')='X' ) then
      lInErr := RC_AUXFERME ;
    // Général
    if ( lInErr = RC_PASERREUR ) and lInfo.LoadCompte( lTobEcr.GetString('E_GENERAL') )
                                 and ( lInfo.Compte_GetValue('G_FERME')='X' ) then
      lInErr := RC_CFERME ;
    // Sections
    if ( lInErr = RC_PASERREUR ) and ( lTobEcr.GetString('E_ANA') = 'X' ) then
      for iAxe := 0 to lTobEcr.Detail.Count-1 do
        begin
        for iVentil := 0 to lTobEcr.Detail[iAxe].Detail.Count-1 do
          begin
          lTobAna := lTobEcr.Detail[iAxe].Detail[iVentil] ;
          if lInfo.LoadSection( lTobAna.GetString('Y_SECTION'), lTobAna.GetString('Y_AXE') )
              and ( lInfo.Section_GetValue('S_FERME')='X' ) then
              begin
              lInErr := RC_YSECTIONFERMEE ;
              break ;
              end ;
          end ;
        if ( lInErr <> RC_PASERREUR ) then Break ;
        end ;

    if lInErr <> RC_PASERREUR then
      begin
      Case lInErr of
        RC_CFERME             : lStRef    := 'ATTENTION ! Présence d''un compte fermé' ;
        RC_AUXFERME           : lStRef    := 'ATTENTION ! Présence d''un auxiliaire fermé' ;
        RC_YSECTIONFERMEE     : lStRef    := 'ATTENTION ! Présence d''une section fermée' ;
        end ;
      vTobEcr.PutValue( 'E_REFINTERNE', lStRef ) ;
      FreeAndNil (lInfo); // Ajout CA - 21/07/2006
      FreeAndNil( lTOBPiece ) ;
      Exit;
      end ;
    end ;
  FreeAndNil (lInfo); // Ajout CA - 21/07/2006
  // ====================================================
  // ==== Détermination des paramètres de génération ====
  // ====================================================
  // >> Date comptable
  if not vBoDateForce
    then vDtDateC := lTOBPiece.Detail[0].GetDateTime('E_DATECOMPTABLE') ;
  // Test de la date
  lInErr := DateCorrecte( vDtdateC );
  if lInErr > 0 then
    begin
    Case lInErr of
      1 : vTobEcr.PutValue( 'E_REFINTERNE', Copy( 'Date sur exercice non ouvert', 1, 35 ) ) ;
      2 : vTobEcr.PutValue( 'E_REFINTERNE', Copy( 'Date sur exercice non ouvert', 1, 35 ) ) ;
      3 : vTobEcr.PutValue( 'E_REFINTERNE', Copy( 'Date antérieure à clôture provisoire', 1, 35 ) ) ;
      4 : vTobEcr.PutValue( 'E_REFINTERNE', Copy( 'Date antérieure à clôture définitive', 1, 35 ) ) ;
      end ;
    lTOBPiece.Free ;
    Exit;
    end ;
  // >> Date / Heure de modif
  lDtModif := NowH ;
  // >> détermination du numéro de pièce pour le passage en écritures normales
  if vStDest = 'N' then
    begin
    lNumPiece := GetNewNumJal( lTOBPiece.Detail[0].GetValue('E_JOURNAL'), True, vDtDateC ) ;
    if lNumPiece <= 0 then
      begin
      lTOBPiece.Free ;
      Exit ;
      end ;
    end
  else
    lNumPiece := lTOBPiece.Detail[0].GetValue('E_NUMEROPIECE') ;

  // Détermination de la référence libre
  if vStOrig = 'S'
    then lStRef := 'Pièce de simulation N° '
  else if vStOrig = 'R'
    then lStRef := 'Pièce de révision N° '
  else if vStOrig = 'U'
    then lStRef := 'Pièce de situation N° '
  else lStRef := 'Pièce ' ;


  lStRef := lStRef + lTOBPiece.Detail[0].GetString('E_NUMEROPIECE') ;

  // paramètres liés à la date
  lStExo := QuelExoDt(  vDtDateC ) ;
  lInPer := GetPeriode( vDtDateC ) ;
  lInSem := NumSemaine( vDtDateC ) ;

  // ========================================
  // ==== PARCOURS DES LIGNES D'ECRITURE ====
  // ========================================
  for lNumL := 0 to ( lTOBPiece.Detail.Count - 1 ) do
    begin
    lTobEcr := lTobPiece.Detail[lNumL] ;

    // >> Détermination du TRSYNCHRO pour chaque ligne...
    lStTrSyn := 'RIE';
    if EstComptaTreso and EstCptSyn( lTobEcr.GetString('E_GENERAL'),
                                     lTobEcr.GetString('E_JOURNAL'),
                                     lTobEcr.GetString('E_NATUREPIECE') <> 'ECC')
      then lStTrSyn := 'CRE';

    // ========================================
    // ==== MAJ base des lignes d'écriture ====
    // ========================================
    // >> Ecritures

    lStReq := 'UPDATE ECRITURE SET ' + 'E_QUALIFPIECE="'    + vStDest                         + '", '
                                     + 'E_DATEMODIF="'      + UsTime( lDtModif )              + '", '
                                     + 'E_DATECOMPTABLE="'  + UsDatetime( vDtDateC )          + '", '
                                     + 'E_EXERCICE="'       + lStExo                          + '", ' // GCO - FQ 19059
                                     + 'E_PERIODE='         + IntToStr( lInPer )              + ', '
                                     + 'E_SEMAINE='         + IntToStr( lInSem )              + ', '
                                     + 'E_QUALIFORIGINE="'  + vStOrig                         + '" ' ;
    if vStDest = 'N' then
      lStReq := lStReq               + ', E_NUMEROPIECE='   + IntToStr( lNumPiece )
                                     + ', E_TRESOSYNCHRO="' + lStTrSyn                        + '"' ;
    lStReq := lStReq + ' WHERE ' + WhereEcritureTob( tsGene, lTobEcr ,True ) ;

    lInNbEnreg := ExecuteSQL ( lStReq ) ;
    Result := lInNbEnreg > 0 ;
    // >> Analytique
    if Result then
      begin
      lStReq := 'UPDATE ANALYTIQ SET ' + 'Y_QUALIFPIECE="'    + vStDest                         + '", '
                                       + 'Y_DATEMODIF="'      + UsTime( lDtModif )              + '", '
                                       + 'Y_DATECOMPTABLE="'  + UsDatetime( vDtDateC )          + '", '
                                       + 'Y_EXERCICE="'       + lStExo                          + '", ' // SBO - FQ 21629
                                       + 'Y_PERIODE='         + IntToStr( lInPer )              + ', '
                                       + 'Y_SEMAINE='         + IntToStr( lInSem ) ;
      if vStDest = 'N' then
        lStReq := lStReq + ', Y_NUMEROPIECE='   + IntToStr( lNumPiece ) ;
      lStReq := lStReq + ' WHERE ' + WhereEcritureTob( tsAnal, lTobEcr ,True ) ;

      ExecuteSQL ( lStReq ) ;
      end ;

    // ======================
    // ==== MAJ des Tobs ====
    // ======================
    lTobEcr.PutValue('E_QUALIFPIECE',   vStDest   ) ;
    lTobEcr.PutValue('E_DATECOMPTABLE', vDtDateC  ) ;
    lTobEcr.PutValue('E_EXERCICE',      lStExo    ) ; // GCO - FQ 19059
    lTobEcr.PutValue('E_PERIODE',       lInPer    ) ;
    lTobEcr.PutValue('E_SEMAINE',       lInSem    ) ;
    lTobEcr.PutValue('E_QUALIFORIGINE', vStOrig   ) ;
    lTobEcr.Putvalue('E_DATEMODIF',     lDtModif  ) ;
    lTobEcr.PutValue('E_REFLIBRE',      lStRef    ) ;
    if vStDest = 'N' then
      begin
      lTobEcr.PutValue('E_NUMEROPIECE',   lNumPiece ) ;
      lTobEcr.PutValue('E_TRESOSYNCHRO',     lStTrSyn  ) ;
      end ;

    // Parcours des lignes d'analytiques
    for lNumA := 0 to lTobEcr.Detail.Count - 1 do
      for lNumV := 0 to lTobEcr.Detail[ lNumA ].Detail.Count - 1 do
        begin
        lTobVentil := lTobEcr.Detail[ lNumA ].Detail[ lNumV ] ;
        lTobVentil.PutValue('Y_QUALIFPIECE',   vStDest   ) ;
        lTobVentil.PutValue('Y_DATECOMPTABLE', vDtDateC  ) ;
        lTobVentil.PutValue('Y_EXERCICE',      QuelExoDt( vDtDateC ) ); // SBO - FQ 21629
        lTobVentil.PutValue('Y_PERIODE',       lInPer    ) ;
        lTobVentil.PutValue('Y_SEMAINE',       lInSem    ) ;
        lTobVentil.Putvalue('Y_DATEMODIF',     lDtModif  ) ;
        if vStDest = 'N' then
          lTobVentil.PutValue('Y_NUMEROPIECE', lNumPiece ) ;
        end ;

    end ;


  // ====================================
  // ==== MAJ des soldes des comptes ====
  // ====================================
  if vStDest = 'N' then
    MajSoldesEcritureTOB( lTOBPiece, True ) ;

  // ======================================
  // ==== MAJ TOB paramètre si demandé ====
  // ======================================
  if vBoMajTOB then
    begin
    vTobEcr.PutValue('E_QUALIFPIECE',   vStDest   ) ;
    vTobEcr.PutValue('E_DATECOMPTABLE', vDtDateC  ) ;
    vTobEcr.PutValue('E_EXERCICE', QuelExoDt( vDtDateC )); // GCO - FQ 19059
    vTobEcr.PutValue('E_QUALIFORIGINE', vStOrig   ) ;
    vTobEcr.Putvalue('E_DATEMODIF',     lDtModif  ) ;
    vTobEcr.PutValue('E_REFLIBRE',      lStRef    ) ;
    if vStDest = 'N' then
      begin
      vTobEcr.PutValue('E_NUMEROPIECE',   lNumPiece ) ;
      vTobEcr.PutValue('E_TRESOSYNCHRO',     lStTrSyn  ) ;
      end ;
    end ;

  // Libération
  lTobPiece.ClearDetail ;
  lTobPiece.Free ;

end ;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnNew;
begin
  Inherited;
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnDelete;
begin
  Inherited;
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnUpdate;
begin
  Inherited;
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnLoad;
begin
  Inherited;
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnArgument (S : String );
begin

  FBoChargEnCours := true; // FQ 18696

  FModeGener := S ;
  if FModeGener = '' then
    FModeGener := 'S' ;

  Inherited;

  // composants
  InitComposants ;
  InitEvenements ;
  InitMessages;
  InitVariables;
  InitPresentation ;

  // Recherche au lancement
//  TFMul(Ecran).BChercheClick(nil);

  {JP 18/11/04 : Dans ValidSim c'est inférieur et ici égal, ce qui fait que le mul ne me renvoie
                 aucune requête}
  THEdit(GetControl('E_NUMECHE')).Operateur := Inferieur;
  { CA - 19/07/2006 - FQ 18607 }
  SetControlVisible('BSELECTALL',True);

  { CA - 29/11/2006 - On autorise la validation malgré la présence de E_REFGESCOM en BTP - cf B. Locatelli }
  (* if (GetParamSocSecur('SO_DISTRIBUTION','')='014') or (EstSpecif('51210')) then
    SetControlText('XX_WHEREGC',''); *)

  { CA - 25/04/2007 - On fait le contrôle désormais à la validation }
  SetControlText('XX_WHEREGC','');
  fbControleRefGescom := EstSpecif('51210');

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnClose;
begin
  Inherited;
  _Bloqueur('nrBatch',False);

  // Libération mémoire
  VideListe(  FPiecesGener );
  FreeAndNil( FPiecesGener ) ;

  VideListe(  FPiecesErr );
  FreeAndNil( FPiecesErr ) ;

  FreeAndNil( Messages ) ;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnDisplay ();
begin
  Inherited;
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.OnCancel ();
begin
  Inherited;
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.ExerciceChange(Sender: TObject);
begin
  ExoToDates( E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_ );
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.FListeDblClick(Sender: TObject);
begin
  BZoomPieceClick(Nil);
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.BOuvrirClick(Sender: TObject);
begin

  if Not DateGeneOk then Exit ;

  if ValideSelection then
    TFMul(Ecran).BChercheClick(Sender);
end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.InitMessages;
begin
  Messages := TStringList.Create ;
  with Messages do
    begin
    Add('0;?caption?;Confirmez-vous la transformation en écritures courantes ?;Q;YN;N;N;');
    Add('1;?caption?;Vous n''avez sélectionné aucune écriture.;E;O;O;O;');
    Add('2;LIBRE;LIBRE');
    Add('3;LIBRE;LIBRE');
    Add('4;LIBRE;LIBRE;');
    Add('5;LIBRE;LIBRE;');
    Add('6;?caption?;Voulez-vous voir la liste des écritures courantes générées ?;Q;YN;N;N;');
    Add('7;LIBRE;LIBRE;LIBRE;LIBRE;');
    Add('Pièce de simulation N°');
    Add('Date sur exercice non ouvert');
    Add('Date sur exercice non ouvert');
    Add('Date antérieure à clôture provisoire');
    Add('Date antérieure à clôture définitive');
    Add('Validation impossible. Certaines écritures ont été modifiées par un autre utilisateur.');
    Add('Validation non effectuée.');
    Add('Pièce de révision N°');
    Add('16;?caption?;La date que vous avez renseignée n''est pas valide.;W;O;O;O;');
    Add('17;?caption?;La date que vous avez renseignée n''est pas dans un exercice ouvert.;W;O;O;O;');
    Add('18;?caption?;La date que vous avez renseignée est antérieure à une clôture.;W;O;O;O;');
    Add('19;?caption?;La date que vous avez renseignée est antérieure à une clôture.;W;O;O;O;');
    Add('20;?caption?;La date que vous avez renseignée est en dehors des limites autorisées.;W;O;O;O;');
    Add('Validation des écritures de révision en écritures courantes');
    Add('22;?caption?;Confirmez-vous la transformation en écritures de simulation ?;Q;YN;N;N;');
    Add('23;');
    end;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.BZoomPieceClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(FListe.Row-1);
  TrouveEtLanceSaisie( TFMul(Ecran).Q.TQ, taConsult, GetControlText('E_QUALIFPIECE') ) ;
{$ELSE}
  TrouveEtLanceSaisie( TFMul(Ecran).Q,    taConsult, GetControlText('E_QUALIFPIECE') ) ;
{$ENDIF}
end;


/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.ValideLaPiece;
var  lQEcr    : TQuery ;
     lStReq   : String ;
     lDtDateC : TDateTime ;
     lTobEcr  : TOB ;
     lstRefGescom : string;
begin

  {23/03/07 YMO FQ19613 Compatibilité eAGL}
  lStReq :=  'SELECT * FROM ECRITURE '
           + 'WHERE E_JOURNAL="'       + TFMul(Ecran).Q.FindField('E_JOURNAL').AsString
           + '" AND E_EXERCICE="'      + QuelExoDT(TFMul(Ecran).Q.FindField('E_DATECOMPTABLE').AsDateTime)
           + '" AND E_DATECOMPTABLE="' + USDATETIME(TFMul(Ecran).Q.FindField('E_DATECOMPTABLE').AsDateTime)
           + '" AND E_NUMEROPIECE='    + TFMul(Ecran).Q.FindField('E_NUMEROPIECE').AsString
           + '  AND E_QUALIFPIECE="'   + GetControlText('E_QUALIFPIECE')
           + '" AND E_NUMLIGNE=1 AND E_NUMECHE<=1' ;

  lQEcr   := OpenSQL ( lStReq, True ) ;

  if lQEcr.Eof then
  begin
    Ferme( lQEcr ) ;
    Exit ;
  end ;

  { CA - 25/04/2007 - Si présence de refgescom, on regarde si on peut valider }
  lstRefGescom := lQEcr.FindField('E_REFGESCOM').AsString;
  if  (( lstRefGescom <> '') and (not CanValidPieceFromGC(lstRefGescom )) and fbControleRefGescom) then
  begin
    Ferme( lQEcr ) ;
    exit;
  end;

  lTobEcr := Tob.Create('ECRITURE', nil, -1) ;
  lTobEcr.SelectDB( '', lQEcr ) ;
  Ferme( lQEcr ) ;

  if CChoixDate.Checked
    then lDtDateC := StrToDate( DATEGENERATION.Text )
    else lDtDateC := iDate1900 ;

  if CTransformerPiece( E_QUALIFPIECE.Value, QualifGener.Value, lTobEcr, True, CChoixDate.Checked, lDtDateC )
  then
  begin
    // Mise à jour de la GESCOM le cas échéant
    // GP le 30/06/2008 : le long roman de la fiche 21906....
//    if (lstRefGescom <> '') and fbControleRefGescom and  (not ((Copy(lstRefGescom, 1, 3) = 'AA;') or (Copy(lstRefGescom, 1, 3) = 'AD;') or
    if (lstRefGescom <> '') and  (not ((Copy(lstRefGescom, 1, 3) = 'AA;') or (Copy(lstRefGescom, 1, 3) = 'AD;') or
    (Copy(lstRefGescom, 1, 3) = 'AR;'))) then
      MajRefCptaOnValidPce (lTobEcr);
    FPiecesGener.Add( lTobEcr );
  end
    else FPiecesErr.Add( lTobEcr ) ;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.TraiteTPVal;
var lTobEcr  : TOB ;
    i        : Integer ;
    M        : RMVT ;
begin

  for i:=0 to FPiecesGener.Count-1 do
    begin
    // Recup TOB
    lTobEcr  := TOB( FPiecesGener[i] );

    FillChar( M, Sizeof(M), #0 ) ;
    M.Etabl         := lTobEcr.GetString('E_ETABLISSEMENT') ;
    M.Jal           := lTobEcr.GetString('E_JOURNAL')         ;
    M.Exo           := lTobEcr.GetString('E_EXERCICE') ;
    M.DateC         := lTobEcr.GetDateTime('E_DATECOMPTABLE') ;
    M.Nature        := lTobEcr.GetString('E_NATUREPIECE') ;
    M.Num           := lTobEcr.GetInteger('E_NUMEROPIECE')     ;
    M.Simul         := lTobEcr.GetString('E_QUALIFPIECE') ;
    M.CodeD         := lTobEcr.GetString('E_DEVISE')        ;
    M.DateTaux      := lTobEcr.GetDateTime('E_DATETAUXDEV') ;
    M.TauxD         := lTobEcr.GetDouble('E_TAUXDEV')       ;
    M.Valide        := lTobEcr.GetString('E_VALIDE') = 'X' ;
    M.ModeSaisieJal := lTobEcr.GetString('E_MODESAISIE') ;

    // Génération des Pièces sur les TP
    GenerePiecesPayeur( M ) ;

    end;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
function TOF_CPVALIDEPIECE.ValideSelection: boolean;
var i      : integer ;
    NbD    : integer ;
    io     : TIOErr ;
    Okok   : boolean ;
    Q      : TQuery;
    lInIdx : integer ;
begin

  VideListe(  FPiecesGener );
  VideListe(  FPiecesErr );

  NowFutur := NowH ;
  Result   := False ;

{$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
{$ELSE}
  Q := TFMul(Ecran).Q;
{$ENDIF}

  if FListe.AllSelected then
    begin
    if ((Q.EOF) and (Q.BOF)) then
      begin
      AfficheMessage(1) ;
      Exit;
      end;

    NbD := 100;
    end
  else
    begin
    NbD := FListe.NbSelected;
    if NbD<=0 then
      begin
      AfficheMessage(1);
      Exit;
      end;
    end;

  // FQ 17334 : SBO 26/04/2006 : message confirmation dependant du module
  if FModeGener = 'U'
    then lInIdx := 22 // Situation >> Simulation
    else lInIdx := 0 ;

  if AfficheMessage(lInIdx) <> mrYes then
    begin
    FListe.ClearSelected;
    FListe.AllSelected := False;
    Exit;
    end;

  Okok := True;
  InitMove(NbD,'');

  if FListe.AllSelected then
    begin
    Q.First;
    while Not Q.EOF do
      begin
      MoveCur(False) ;
      if RevisionActive(Q.FindField('E_DATECOMPTABLE').AsDateTime) then
        begin
        Okok := False;
        Break;
        end;
      Q.Next ;
      end;
    end
  else
    begin
    for i:=0 to NbD-1 do
      begin
      FListe.GotoLeBookMark(i);
      MoveCur(FALSE);
      if RevisionActive( Q.FindField('E_DATECOMPTABLE').AsDateTime ) then
        begin
        Okok := False;
        Break;
        end;
      end;
    end;

  FiniMove;

  if not Okok then Exit;

  InitMove(NbD,'');

  if FListe.AllSelected then
    begin

    Q.First ;
    While Not Q.EOF do
      begin
      MoveCur(FALSE) ;
      io:=Transactions(ValideLaPiece,3) ;
      Case io of
        oeOk : ;
        oeSaisie :  begin MessageAlerte(Messages[13]); FiniMove; Exit; end;
        oeUnknown : begin MessageAlerte(Messages[14]); FiniMove; Exit; end;
        end;
      Q.Next ;
      end;
    end
  else
    begin
    for i:=0 to NbD-1 do
      begin
      Fliste.GotoLeBookMark(i);
    {$IFDEF EAGLCLIENT} {23/03/07 YMO FQ19613 Compatibilité eAGL}
      TFMul(Ecran).Q.TQ.Seek(FListe.Row -1);
    {$ENDIF}
      MoveCur(FALSE);
      io := Transactions(ValideLaPiece,5);
      Case io of
        oeOk : ;
        oeSaisie  : begin MessageAlerte(Messages[13]); FiniMove; Exit; end;
        oeUnknown : begin MessageAlerte(Messages[14]); FiniMove; Exit; end;
        end;
      end;
    end;

  FListe.ClearSelected;
  FListe.AllSelected := False;
  FiniMove;

  Result:=True;

  if FPiecesGener.Count>0 then
    begin
    if VH^.OuiTP then
      TraiteTPVal;
    VisuPiecesGenere( FPiecesGener, EcrGen, 1 ) ;
    end;

  if FPiecesErr.Count>0 then
    VisuPiecesGenere( FPiecesErr, EcrGen, 2 ) ;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
function TOF_CPVALIDEPIECE.DateGeneOk: boolean;
var Err : integer;
begin

  Result         := True;

  if ( ( DATEGENERATION.Text=StDate1900 ) or (DATEGENERATION.Text='') or (Not CChoixDate.Checked)) then Exit;

  Result         := False;

  // Contrôle la date saisie
  Err            := ControleDate( DATEGENERATION.Text );
  if Err>0 then
    begin
    AfficheMessage( 15 + Err ) ;
    // FQ 17989 SBO 26/04/2006 : désactivé date si pb
    CChoixDate.Checked := False ;
    CChoixDateClick(nil);
    Exit;
    end;

  // Test si mode de révisin est actif pour la date saisie
  if RevisionActive( StrToDate( DATEGENERATION.Text ) ) then
    begin
    // FQ 17989 SBO 26/04/2006 : désactivé date si pb
    CChoixDate.Checked := False ;
    CChoixDateClick(nil);
    Exit;
    end;

  Result         := True;

end;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.CChoixDateClick(Sender: TObject);
begin
  if CChoixDate.Checked then
    begin
    DateGeneration.Enabled := True ;
    if ( DateGeneration.Text = StDate1900 ) then
      DateGeneration.Text := DateToStr( V_PGI.DateEntree );
    end
  else
    begin
    DateGeneration.Text    := StDate1900;
    DateGeneration.Enabled := False;
    end ;
end;


procedure TOF_CPVALIDEPIECE.InitComposants ;
begin
  // Récupération des contrôles
  E_QUALIFPIECE      := THValComboBox( GetControl('E_QUALIFPIECE', True)    );
  E_EXERCICE         := THValComboBox( GetControl('E_EXERCICE', True)       );
  E_DATECOMPTABLE    := THEdit(        GetControl('E_DATECOMPTABLE', True)  );
  E_DATECOMPTABLE_   := THEdit(        GetControl('E_DATECOMPTABLE_', True) );
  DATEGENERATION     := THEdit(        GetControl('DATEGENERATION', True)   );

  CChoixDate         := TCheckBox(     GetControl('CCHOIXDATE', True)       );
  QualifGener        := THValComboBox( GetControl('QUALIFGENER', True)      );

  FListe             := TFMul(Ecran).FListe;

  E_JOURNAL          := THValComboBox( GetControl('E_JOURNAL', True)    );
  E_NATUREPIECE      := THValComboBox( GetControl('E_NATUREPIECE', True)    );

end ;

/////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPVALIDEPIECE.InitEvenements;
begin
  // Evénements sur les boutons
  TToolBarButton97( GetControl('BZOOMPIECE', True) ).OnClick := BZoomPieceClick ;
  TCheckBox(        GetControl('CCHOIXDATE', True) ).OnClick := CChoixDateClick ;
  TToolbarButton97( GetControl('BOUVRIR', true)    ).OnClick := BOuvrirClick    ;
  TToolbarButton97( GetControl('BCHERCHE', True)   ).OnClick := BChercheClick   ; // FQ 18696

  // Evénements sur les critères de recherche
  E_EXERCICE.OnChange           := EXERCICEChange    ;
  E_QUALIFPIECE.OnChange        := QualifPieceChange ;

  // Evènements sur la liste
  FListe.OnDblClick             := FListeDblClick     ;

  TFMul(Ecran).OnAfterFormShow := OnAfterShow;

end;

procedure TOF_CPVALIDEPIECE.InitPresentation;
begin

  // Type de pièce accessible
  if V_PGI.Controleur
    then E_QUALIFPIECE.Plus := ' AND CO_CODE IN ( "S", "U", "R") '
    else E_QUALIFPIECE.Plus := ' AND CO_CODE IN ( "S", "U" ) ' ;
  SetControlText('E_QUALIFPIECE', 'S');
  QualifPieceChange   ( nil ) ;

  // 1er Journal
  E_JOURNAL.ItemIndex := 0 ;

  // Dates comptables
  E_DATECOMPTABLE.Text  := StDate1900;
  E_DATECOMPTABLE_.Text := StDate2099;

  // Exercice (entraîne normalement la maj des dates...)
  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' )
    then E_EXERCICE.Value := VH^.CPExoRef.Code
    else E_EXERCICE.Value := VH^.Entree.Code ;
{  if VH^.Suivant.Code<>'' ????
    then E_EXERCICE.Value := VH^.Suivant.Code
    else E_EXERCICE.Value := VH^.Encours.Code;
}

  DATEGENERATION.Text := stDate1900;
  CChoixDate.Checked  := False ;
  CChoixDateClick( nil ) ;

  // Dates d'échéances
  SetControlText('E_DATEECHEANCE',  StDate1900 );
  SetControlText('E_DATEECHEANCE_', StDate2099 );

  // Etablissement // FQ 17338 : SBO 12/04/2006
  if VH^.EtablisCpta then
    PositionneEtabUser(GetControl('E_ETABLISSEMENT', True))
  else
    begin
    if VH^.EtablisDefaut<> ''
      then SetControlText( 'E_ETABLISSEMENT', VH^.EtablisDefaut )
      else THValComboBox( GetControl('E_ETABLISSEMENT', True)).ItemIndex := 1 ;
    SetControlEnabled('E_ETABLISSEMENT',False ) ;
    end ;

  // 1ére nature de pièce
  E_NATUREPIECE.ItemIndex := 0 ;

  // Gestion des différents mode de validation
  E_QualifPiece.Value   := FModeGener ;
  E_QualifPiece.Enabled := False ;
  QualifGener.Enabled   := False ;
  if FModeGener = 'R' then // Révision >> Normal
    begin
    Ecran.Caption     := 'Validation des écritures de révision en écritures courantes';// titre
    Ecran.HelpContext := 7662000 ;   // Aide
    QualifGener.Value := 'N' ;       // Type de pièce générée bloquer à "Normal"
    SetControlProperty('BOUVRIR', 'HINT', 'Valider' ) ; // FQ 17332 SBO 26/04/2006 : hint ouvrir -> Valider ou convertir
    end
  else if FModeGener = 'U' then // Situation >> Simulation
    begin
    Ecran.Caption      := 'Conversion des écritures de situation en écritures de simulation'; // titre
    Ecran.HelpContext  := 7303000;  // Aide // FQ 17717 SBO 26/04/2006
    QualifGener.Value  := 'S' ; // Type de pièce générée
    SetControlProperty('BOUVRIR', 'HINT', 'Convertir' ) ; // FQ 17332 SBO 26/04/2006 : hint ouvrir -> Valider ou convertir
    end
  else // Simulation >> Normal
    begin
    Ecran.Caption     := 'Validation des écritures de simulation en écritures courantes'; // titre
    Ecran.HelpContext := 7281000 ; // Aide
    QualifGener.Value := 'N' ; // Type de pièce générée bloquer à "Normal"
    SetControlProperty('BOUVRIR', 'HINT', 'Valider' ) ;  // FQ 17332 SBO 26/04/2006 : hint ouvrir -> Valider ou convertir
    end ;

  UpdateCaption(Ecran);

end;

procedure TOF_CPVALIDEPIECE.InitVariables;
begin
  FPiecesGener  := TList.Create;
  FPiecesErr    := TList.Create;
end;

function TOF_CPVALIDEPIECE.AfficheMessage(vInErr: Integer) : Word ;
begin
  result := HShowMessage( Messages[ vInErr ], Ecran.Caption, '' ) ;
end;

procedure TOF_CPVALIDEPIECE.QualifPieceChange(Sender: TObject);
begin

  Exit ;

{ plus nécessaire
  if (E_QualifPiece.Value = 'S') or ( E_QualifPiece.Value = 'R') then
    begin
    // Type de pièce générée bloquer à "Normal"
    QualifGener.Enabled := False ;
    QualifGener.Value := 'N'
    end
  else
    begin
    // Type de pièce générée
    QualifGener.Enabled := False ;
    QualifGener.Value   := 'S'
    end ;
}
end;

procedure TOF_CPVALIDEPIECE.BChercheClick(Sender: TObject);
begin
  // FQ 18696 : Au chargement, on ne fait rien !
  if FBoChargEnCours then Exit;
  TFMul(Ecran).BChercheClick( nil );
end;

procedure TOF_CPVALIDEPIECE.OnAfterShow;
begin
  // FQ 18696
  FBoChargEnCours := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVALIDEPIECE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


Initialization
  registerclasses ( [ TOF_CPVALIDEPIECE ] );
end.
