unit ULibIAS14;

interface

uses      SysUtils,
          HEnt1,
          HMsgBox,
          hctrls,
          {$IFDEF EAGLCLIENT}
          {$ELSE}
            {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
          {$ENDIF}
          uTob,
          Ent1,        // VH^
          saisUtil,    // pour RDevise , ADecimP
          UtilSais,    // pour MajSoldeSectionTOB
          ULibEcriture,   // pour WhereEcritureTob
          {$IFDEF MODENT1}
          CPTypeCons,
          {$ENDIF MODENT1}
          UlibAnalytique // pour CGetNewTobAna, InitCommunObjAnalNew
           ;

// ====================================================
// ==== Codes d'erreur utilisés lors du traitement ====
// ====================================================

Const    IAS_PASERREUR             = -1 ;
         // Erreur
         IAS_ERRTRAITEMENT         =  0 ;
         IAS_ERRINCONNU            =  1 ;
         IAS_ERRNOAXE              =  2 ;  // Axe obligatoire
         IAS_ERRNOEXO              =  3 ;  // Exercice obligatoire
         IAS_ERRNOTOLMONTANT       =  4 ;  // Montant obligatoire si gestin tolérance
         IAS_ERRNOPERDANSEXO       =  5 ;  // Période non comprise dans l'exercice
         IAS_ERRDATEDEBUTFIN       =  6 ;  // Date début < date fin
         // Warning
         IAS_WNGEXOCLO             = 100 ;
         IAS_WNGZEROLIGNE          = 101 ;

// ============================================
// ==== Processus de traitement de l'IAS14 ====
// ============================================

Type TTraitementIAS14 = Class
        public
          // Paramètres du traitement
          AvecExoClo        : Boolean ;
          Exercice          : String ;
          DateDebut         : TDateTime ;
          DateFin           : TDateTime ;
          Axe               : String ;
          NatureGene        : String ;
          NaturePiece       : String ;
          Tolerance         : Boolean ;
          ToleranceMontant  : Double ;

          // Traitement de l'IAS 14
          Function  TraitementIAS14 : Integer ;
          // Vérification des parmaètres avant traitement
          Function  VerifieParamIAS14 : Integer ;
          // Affichage des messages d'erreurs
          Procedure AfficheMessage ( vInErr : Integer ; vStTitre : String ) ;
          // Update des paramSoc de dernières périodes traitées
          Procedure UpdateInfosIAS14 ;
          // Vérifie que la période est dans l'exercice
          Function VerifPeriode : Boolean ;
          // Connaître le nombre de ligne traitée
          Function GetNbLignesTraitees : Integer ;

        private

          TobTTC    : TOB ;     // Contient LA ligne de TTC à reventiler
          TobHT     : TOB ;     // Contient les lignes de HT servant de base pour les calculs
          TobVentil : TOB ;     // Contient le résultat du calcul de reventilation
          TobSuppr  : TOB ;     // Contient la ligne de sectin d'attente supprimée

          RDev : RDevise ;      // Devise de l'ecriture en cours de traitement

          NbLignes  : Integer ;

          //SG6 26.01.05 Gestion croisaxe
          Axes : array[1..MaxAxe] of boolean;

          // Fonctions de chargement
          Function  GenererRequeteLignesCibles : String ;
          Procedure ChargePiece( vQEcr : TQuery ) ;

          // Fonctions de vérifications
          Function EstPieceCompatible : Boolean ;

          // Fonctions de Traitement
          Procedure CalculReventilation ;
          Procedure ReventileLigne( vTobEcr : TOB ) ;
          Procedure SauvePiece ;
          Procedure RecopieInfos( vTobAna, vTobAtt : TOB ) ;
          Procedure SetVentilLigne( vTobAna, vTobVent : TOB ; vBoDebit : Boolean );

          // Fonctions diverses
          Function EstLigneTTC ( vTobLigne : TOB ) : Boolean ;
          Function EstLigneHT  ( vTobLigne : TOB ) : Boolean ;
          Function GetCoefTVA  ( vTobLigne : TOB ) : Double ;

        End ;


implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  ParamSoc;
  
{ TTraitementIAS14 }

procedure TTraitementIAS14.AfficheMessage (vInErr: Integer; vStTitre: String ) ;
begin

  // ==========================================
  // ==== TRAITEMENT CORRECTEMENT EFFECTUE ====
  // ==========================================
  if vInErr = IAS_PASERREUR then
    if NbLignes > 1
      then PGIInfo('Le traitement s''est correctement effectué. ' + IntToStr( NbLignes ) + ' lignes d''écriture ont été traitées.', vStTitre)
      else  PGIInfo('Le traitement s''est correctement effectué. ' + IntToStr( NbLignes ) + ' ligne d''écriture a été traitée.', vStTitre)

  // =================
  // ==== WARNING ====
  // =================
  // Warning Traitement sur exo clo
  else if vInErr = IAS_WNGEXOCLO then
    PGIInfo('Attention ! La prise en compte des modifications d''imputation analytique'
           + #10#13 + 'sur les reports à nouveau necéssitera une dé-clôture', vStTitre)

  // Warning Aucune ligne à traiter
  else if vInErr = IAS_WNGZEROLIGNE then
    PGIInfo('Aucune ligne à traiter avec les paramètres saisis !', vStTitre)

  // =================
  // ==== ERREURS ====
  // =================

  // Erreur pendant traitement
  else if vInErr = IAS_ERRINCONNU then
    PGIBox('Attention ! Une erreur inconnue s''est produite durant le traitement !' , vStTitre)

  // Axe : paramètre obligatoire
  else if vInErr = IAS_ERRNOAXE then
    PGIBox('Aucun axe de sélection n''a été saisi. Ce paramètre est obligatoire.' , vStTitre)

  // Exercice : paramètre obligatoire
  else if vInErr = IAS_ERRNOEXO then
    PGIBox('Aucun exercice de sélection n''a été saisi. Ce paramètre est obligatoire.' , vStTitre)

  // Si tolérance : montant obligatoire
  else if vInErr = IAS_ERRNOTOLMONTANT then
    PGIBox('Vous devez saisir un montant de tolérance en cas de gestion de la tolérance.' , vStTitre)

  // date début > Date fin
  else if vInErr = IAS_ERRDATEDEBUTFIN then
    PGIBox('La date de début de période doit être inférieure ou égale à la date de fin de période.' , vStTitre)

  // Période dans exercice
  else if vInErr = IAS_ERRNOPERDANSEXO then
    PGIBox('La période saisie n''est pas comprise dans l''exercice sélectionnée. Veuillez revoir votre saisie.' , vStTitre)

    ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 18/05/2004
Modifié le ... : 18/05/2004
Description .. : Retourne une TOB contenant les lignes d'écritures de la
Suite ........ : pièces + les ventilations en 2ème niveau
Mots clefs ... :
*****************************************************************}
Procedure TTraitementIAS14.ChargePiece( vQEcr: TQuery ) ;
var lStReq    : String ;
    lQEcr     : TQuery ;
    lTobPiece : TOB ;
    lTobLigne : TOB ;
    lInEcr    : Integer ;
    lStSQL    : String; //SG6 26.01305 Mode Croisaxe
begin

  if TobHT.Detail.Count > 0 then TobHt.ClearDetail ;
  if TobTTC.Detail.Count > 0 then TobTTC.ClearDetail ;

  // Chargement des écritures ventilées sur l'axe à traiter
  lTobPiece := TOB.Create('PIECE_TMP', nil, -1) ;
  //SG6 26.01.05 Gestion mode croisaxe
  if not(VH^.AnaCroisaxe) then lStSQL :=  Axe[2] + '="X"'
  else lStSQL := '="X"';

  lStReq := 'SELECT ECRITURE.*, G_NATUREGENE FROM ECRITURE '
             + ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL '
             + ' WHERE E_JOURNAL="'        + vQEcr.FindField('E_JOURNAL').AsString                     + '" AND '
             + 'E_EXERCICE="'       + vQEcr.FindField('E_EXERCICE').AsString                    + '" AND '
             + 'E_DATECOMPTABLE="'  + UsDateTime(vQEcr.FindField('E_DATECOMPTABLE').AsDateTime) + '" AND '
             + 'E_NUMEROPIECE='     + IntToStr(vQEcr.FindField('E_NUMEROPIECE').AsInteger)      + ' AND ' + 'E_ANA="X" AND G_VENTILABLE' + lStSQL
             ;
  lQEcr := OpenSQL( lStReq , True ) ;
  lTobPiece.LoadDetailDB( 'ECRITURE', '', '', lQEcr, False, False) ;
  Ferme( lQEcr ) ;

  // Chargement de toutes les lignes d'analytique de l'axe à traiter
  //SG6 26.01.05 Gestion mode croisaxe
  if not(VH^.AnaCroisaxe) then lStSQL := ' AND Y_AXE="' + Axe + '"'
  else lStSQL := '';

  for lInEcr := 0 to ( lTobPiece.Detail.Count - 1 ) do
    begin
    lTobLigne := lTobPiece.Detail[ lInEcr ] ;
    if lTobLigne.GetValue('E_ANA')='X' then
      begin
      lStReq  := 'SELECT * FROM ANALYTIQ '
            + 'WHERE Y_JOURNAL="'        + lTobLigne.GetValue('E_JOURNAL')
            + '" AND Y_EXERCICE="'       + lTobLigne.GetValue('E_EXERCICE')
            + '" AND Y_DATECOMPTABLE="'  + UsDateTime(lTobLigne.GetValue('E_DATECOMPTABLE'))
            + '" AND Y_NUMEROPIECE='     + IntToStr(lTobLigne.GetValue('E_NUMEROPIECE'))
            +  ' AND Y_NUMLIGNE='        + IntToStr(lTobLigne.GetValue('E_NUMLIGNE'))
            +  ' AND Y_QUALIFPIECE="N"'
            +  lStSQL
            ;
      lQEcr := OpenSQL( lStReq , True ) ;
      lTobLigne.LoadDetailDB( 'ANALYTIQ', '', '', lQEcr, False, False) ;
      Ferme( lQEcr ) ;
      end ;
    end ;

  // Répartition HT / TTC
  For lInEcr := lTobPiece.Detail.Count - 1 downto 0 do
    begin
    lTobLigne := lTobPiece.Detail[ lInEcr ] ;
    // Ligne de TTC
    if EstLigneTTC( lTobLigne ) then
      lTobLigne.ChangeParent( TobTTC, -1)
    // Ligne de HT
    else if EstLigneHT( lTobLigne ) then
      lTobLigne.ChangeParent( TobHT, -1) ;
    end ;

  // Libération reste de la pièce
  lTobPiece.ClearDetail ;
  FreeAndNil( lTobPiece ) ;

end;

function TTraitementIAS14.EstLigneHT(vTobLigne: TOB): Boolean;
begin
  Result := ( vTobLigne.GetValue('E_TYPEMVT') = 'HT' ) or
          ( ( vTobLigne.GetValue('E_TYPEMVT') = 'DIV' ) and
            ( pos( vTobLigne.GetValue('G_NATUREGENE') , 'IMO;CHA;PRO') > 0 ) ) ;
end;

function TTraitementIAS14.EstLigneTTC(vTobLigne: TOB): Boolean;
begin
  Result := ( vTobLigne.GetValue('E_TYPEMVT') = 'TTC' ) or
          ( ( vTobLigne.GetValue('E_TYPEMVT') = 'DIV' ) and
            ( pos( vTobLigne.GetValue('G_NATUREGENE') , 'COF;COC;COS;COD;DIV;TIC;TID') > 0 ) ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 18/05/2004
Modifié le ... : 18/05/2004
Description .. : Retourne la requête de la boucle principale ciblant les
Suite ........ : lignes potentiellement traitables
Mots clefs ... :
*****************************************************************}
function TTraitementIAS14.GenererRequeteLignesCibles: String;
begin

   // Les champs + Table
  Result := 'SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE FROM ECRITURE ' ;
  Result := Result  + ' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;

  // Les conditions fixes
  Result := Result + ' WHERE E_MODESAISIE="-" AND E_QUALIFPIECE="N" AND E_ECRANOUVEAU<>"OAN" ' ;
  Result := Result + ' AND E_ANA="X" AND E_ETATLETTRAGE IN ("AL","PL") ' ;
  Result := Result + ' AND G_IAS14="X" ' ;

  // Condition sur type de mvt
  if Pos('OD',NaturePiece) > 0
    then Result := Result + ' AND ( ( E_TYPEMVT = "TTC" ) OR ( E_TYPEMVT = "DIV" AND E_NATUREPIECE="OD" ) ) '
    else Result := Result + ' AND E_TYPEMVT = "TTC" ' ;

  // Les conditions de paramétrage
  Result := Result + ' AND E_EXERCICE="' + Exercice + '"'
                   + ' AND E_DATECOMPTABLE>="' + USDateTime( DateDebut ) + '"'
                   + ' AND E_DATECOMPTABLE<="' + USDateTime( DateFin ) + '"' ;

  // Conditions sur Axe
  //SG6 26.01.05 Gestion du mode analytique croisaxe
  if not(VH^.AnaCroisaxe) then Result := Result + ' AND G_VENTILABLE' + Axe[2] + '="X"' ;

  // Nature de pièces
  Result := Result + ' AND E_NATUREPIECE IN ("' + FindEtReplace(NaturePiece,';','","', True) +'")' ;

  // Nature de compte
  Result := Result + ' AND G_NATUREGENE IN ("' + FindEtReplace(NatureGene,';','","', True) +'")' ;

  // Tri
  Result := Result + ' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE' ;

end;

function TTraitementIAS14.TraitementIAS14 : Integer;
var lQCibles      : TQuery ;
    lBoTransacOk  : Boolean ;
    lInCpt        : Integer ;
begin
  Result       := IAS_PASERREUR ;
  lBoTransacOk := False ;
  NbLignes     := 0 ;

  //SG6 26.01.05 Init du tableau Axes si mode croisaxe
  if VH^.AnaCroisaxe then
  begin
     Axe := '';

     for lInCpt := 1 to MaxAxe do
     begin
       Axes[lInCpt] := GetParamSocSecur('SO_VENTILA' + IntToStr( lInCpt ),False) ;
       if Axes[lInCpt] then
       begin
         if Axe = '' then Axe := 'A' + IntToStr( lInCpt );
       end;
     end;

  end;


  // Init des Tob
  TobHT     := TOB.Create('ECR_HT', nil, -1) ;
  TobTTC    := TOB.Create('ECR_HT', nil, -1) ;
  TobVentil := TOB.Create('NEW_VENTIL', nil, -1 ) ;
  TobSuppr  := TOB.Create('OLD_VENTIL', nil, -1 ) ;

  lQCibles := OpenSQL( GenererRequeteLignesCibles , True ) ;

  try
    BeginTrans ;

    While not ( lQCibles.Eof ) do
      begin

      // Chargement de la pièce
      ChargePiece( lQCibles ) ;

      // Traitement uniquement si pièce compatible
      if EstPieceCompatible then
        begin

        Inc( NbLignes ) ;

        // Recherche de la devise
        RDev.Code := TobHT.Detail[0].GetValue('E_DEVISE') ;
        GetInfosDevise( RDev ) ;
        RDev.Taux  := GetTaux( RDev.Code, RDev.DateTaux, TobHT.Detail[0].GetValue('E_DATECOMPTABLE') ) ;

        // Calcul ventilation
        CalculReventilation ;

        // Application à la ligne
        For lInCpt := 0 to TobTTC.Detail.Count - 1 do
          ReventileLigne( TobTTC.Detail[ lInCpt ] ) ;

        // Enregistrement
        if Transactions( SauvePiece, 2 ) <> oeOK then
          raise EAbort.Create('Erreur lors de l''enregistrement en base. Traitement annulé.') ;

        end ;

      // Au suivant
      lQCibles.Next ;
      end ;

    CommitTrans ;

    // Traitement ok uniquement si au moins 1 ligne traitée
    if NbLignes > 0 then
      begin
      UpdateInfosIAS14 ; // MAJ PARAMSOC
      lBoTransacOk := True ;
      end ;

    finally
      if not lBoTransacOk then
        if NbLignes > 0 then
          begin
          Rollback ;
          result := IAS_ERRTRAITEMENT ;
          end
        else
          result := IAS_WNGZEROLIGNE ;

  end ;


  // Libération mémoire
  Ferme( lQCibles ) ;
  TobHT.ClearDetail ;
  FreeAndNil( TobHt ) ;
  TobTTC.ClearDetail ;
  FreeAndNil( TobTTC ) ;
  TobVentil.ClearDetail ;
  FreeAndNil( TobVentil ) ;
  TobSuppr.ClearDetail ;
  FreeAndNil( TobSuppr ) ;

end;

procedure TTraitementIAS14.UpdateInfosIAS14 ;
begin
  SetParamSoc('SO_CPIAS14DEBUT', DateDebut ) ;
  SetParamSoc('SO_CPIAS14FIN',   DateFin ) ;
end;

function TTraitementIAS14.VerifieParamIAS14 : Integer;
begin
  Result := IAS_PASERREUR ;

  // Axe : paramètre obligatoire
  //SG6 26.01.05 Gestion croisaxe
  if Axe = '' then
    if not(VH^.AnaCroisaxe) then Result := IAS_ERRNOAXE
  // Exercice : paramètre obligatoire
  else if Exercice = '' then
    Result := IAS_ERRNOEXO
  // Date debut > Date fin ?
  else if Datedebut > DateFin then
    Result := IAS_ERRDATEDEBUTFIN
  // Période dans exercice
  else if not VerifPeriode then
    Result := IAS_ERRNOPERDANSEXO ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 18/05/2004
Modifié le ... :   /  /
Description .. : Retourne VRAI si les lignes d'écritures répondent aux
Suite ........ : critères de traitement de l'IAS14 :
Suite ........ :   Au moins une ligne de HT
Suite ........ :   Au moins une ligne de TTC avec même général / auxiliaire
Mots clefs ... :
*****************************************************************}
function TTraitementIAS14.EstPieceCompatible : Boolean;
var lBoTTCOk    : Boolean ;
    lBoHTOk     : Boolean ;
    lInCpt      : Integer ;
begin

  // Test ligne de TTC : 1 et 1 seule ligne TTC ventilée sur l'axe à traiter à 100% sur section d'attente
  lBoTTCOk :=  ( TobTTC.Detail.Count = 1 )                                         // 1 ligne de TTC
           and ( TobTTC.Detail[0].Detail.Count = 1 )                               // 1 ligne de ventilation analytique
           and ( TobTTC.Detail[0].Detail[0].GetValue('Y_POURCENTAGE') = 100.0 ) ;  // Test pourcentage 100%;

  //SG6 26.01.05 Gestion mode croisaxe
  if VH^.AnaCroisaxe then
  begin
    for lInCpt := 1 to MaxAxe do
    begin
      // Test section d'attente
      if Axes[lInCpt] then lBoTTCok := lBoTTCok and
                      ( TobTTC.Detail[0].Detail[0].GetValue('Y_SOUSPLAN' + IntToStr( lInCpt) ) = VH^.Cpta[ AxeToFb( 'A' + IntToStr( lInCpt) ) ].Attente );
    end;
  end
  else
  begin
    lBoTTCok := lBoTTCok and ( TobTTC.Detail[0].Detail[0].GetValue('Y_SECTION') = VH^.Cpta[ AxeToFb( Axe ) ].Attente )  // Test section d'attente
  end;



  // Au moins une ligne de HT ventilée sur axe
  lBoHTOk := ( TobHT.Detail.Count > 0 )                 // Au moins 1 ligne de TTC
         and ( TobHT.Detail[0].Detail.Count > 0 ) ;     // Au moins 1 ligne de ventilation analytique

  Result := lBoTTCOk and lBoHTOk ;

end;

function TTraitementIAS14.VerifPeriode: Boolean;
var lDtDebutExo : TDateTime ;
    lDtFinExo   : TDateTime ;
    lQExo       : TQuery ;
begin

  result := False ;

  lQExo := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="' + Exercice + '"' ,TRUE) ;
  if Not lQExo.EOF then
    begin
    lDtDebutExo := lQExo.FindField('EX_DATEDEBUT').asDateTime ;
    lDtFinExo   := lQExo.FindField('EX_DATEFIN').asDateTime ;
    Result := ( DateDebut <= Datefin ) and not ( ( DateDebut < lDtDebutExo ) or ( DateFin > lDtFinExo ) ) ;
    end;
  Ferme(lQExo) ;

end;

procedure TTraitementIAS14.CalculReventilation ;
var lTobAna     : TOB ;
    lTobLigneV  : TOB ;
    lTobEcr     : TOB ;
    lInEcr      : Integer ;
    lInCpt      : Integer ;
    lInCpt2     : Integer ;
    lCoefTva    : Double ;
    lMontant    : Double ;
    lDelta      : Double ;
    lSection    : String ;
    lTotal      : Double ;
    lInNum      : Integer ;
    lStAttente  : String ;
    lPourcent   : Double ;
    lTotalP     : Double ;
begin

  TobVentil.ClearDetail ;

  lStAttente   := VH^.Cpta[ AxeToFb( Axe ) ].Attente ;
  lTotal       := 0.0 ;
  lInNum       := 0 ;

  // ---------------------------------------------
  // -- Calcul des montants pour chaque section --
  // ---------------------------------------------
  for lInEcr := 0 to TobHT.Detail.Count - 1 do
    begin
    lTobEcr  := TobHT.Detail[ lInEcr ] ;
    lCoefTva := GetCoefTva( lTobEcr ) ;

    for lInCpt := 0 to lTobEcr.Detail.Count - 1 do
      begin
      lTobAna    := lTobEcr.Detail[ lInCpt ] ;
      lSection   := lTobAna.GetValue('Y_SECTION') ;
      // recherche de la ligne concernant la section courante
      lTobLigneV := TobVentil.FindFirst(['V_SECTION'], [lSection], True) ;
      if lTobLigneV = nil then
        begin
        Inc( lInNum ) ;

        lTobLigneV := Tob.Create('VENTIL', TobVentil, -1) ;

        lTobLigneV.PutValue('V_SECTION',       lSection ) ;
        lTobLigneV.PutValue('V_MONTANT',       0.0 ) ;
        lTobLigneV.PutValue('V_TAUXMONTANT',   0.0 ) ;
        lTobLigneV.PutValue('V_NUMEROVENTIL',  lInNum ) ;

        //SG6 26.01.05 Mode Croisaxe Ajout de champs supplémentaire SOUSPLAN
        for lInCpt2 := 1 to MaxAxe do
        begin
          lTobLigneV.PutValue('V_SOUSPLAN' + IntToStr(lInCpt2), lTobAna.GetValue('Y_SOUSPLAN' + IntToStr( lInCpt2 ) ));
        end;


        end ;

      // Calcul du montant TTC
      lMontant := lTobAna.GetValue('Y_DEBITDEV') + lTobAna.GetValue('Y_CREDITDEV') ;
      lMontant := Arrondi( lMontant * lCoefTva , RDev.Decimale ) ;

      // MAJ montant ligne
      lTobLigneV.PutValue('V_MONTANT', lTobLigneV.GetValue('V_MONTANT') + lMontant ) ;

      // MAJ TOTAL
      lTotal := lTotal + lMontant ;

      end ;
    end ;

  // -----------------------------
  // -- Gestion de la tolérance --
  // -----------------------------
  lTobEcr     := TobTTC.Detail[0] ;
  lDelta      := ( lTobEcr.GetValue('E_DEBITDEV') + lTobEcr.GetValue('E_CREDITDEV') ) - lTotal ;
  if (lDelta <> 0) then
    begin

    if Tolerance and ( abs(lDelta) >= ToleranceMontant ) then
      // ajout d'une ligne sur compte d'attente
      begin
      lTobLigneV := TobVentil.FindFirst(['V_SECTION'], [lStAttente], True) ;
      if lTobLigneV = nil then
        begin
        Inc( lInNum ) ;
        lTobLigneV := Tob.Create('VENTIL', TobVentil, -1) ;
        lTobLigneV.PutValue('V_SECTION',      lStAttente ) ;
        lTobLigneV.PutValue('V_MONTANT',      0.0 ) ;
        lTobLigneV.PutValue('V_TAUXMONTANT',  0.0 ) ;
        lTobLigneV.PutValue('V_NUMEROVENTIL', lInNum ) ;

        //SG6 26.01.05 Gestion mode croisaxe ajout champs supplémentaires SOUSPLAN
        for lInCpt := 1 to MaxAxe do
        begin
          if Axes[lInCpt] then lTobLigneV.PutValue('V_SOUSPLAN' + IntToStr( lInCpt), VH^.Cpta[ AxeToFb( 'A' + IntToStr( lInCpt) ) ].Attente);
        end;

        end ;
      end
    else
      // Absorption delta sur dernière ligne
      lTobLigneV := TobVentil.Detail[ lInNum - 1 ] ;

    // MAJ montant
    lTobLigneV.PutValue('V_MONTANT', lTobLigneV.GetValue('V_MONTANT') + lDelta ) ;
    lTotal := ( lTobEcr.GetValue('E_DEBITDEV') + lTobEcr.GetValue('E_CREDITDEV') ) ;

    end ;

  // -----------------------------------------
  // -- Calcul des taux pour chaque section --
  // -----------------------------------------
  lTotalP := 0 ;
  for lInCpt := 0 to TobVentil.Detail.Count - 1 do
    begin
    lTobLigneV := TobVentil.Detail[ lInCpt ];
    if lInCpt = (TobVentil.Detail.Count - 1)
      then lPourcent := 100.0 - lTotalP
      else lPourcent := Arrondi( lTobLigneV.GetValue('V_MONTANT') / lTotal * 100, ADecimP ) ;

    lTobLigneV.PutValue('V_TAUXMONTANT', lPourcent) ;
    lTotalP := lTotalP + lPourcent ;
    end ;

end;

Procedure TTraitementIAS14.SauvePiece ;
var lStReq  : String ;
    lTobEcr : TOB ;
    lTobAna : TOB ;
    lInEcr  : Integer ;
    lInAna  : Integer ;
    lNowH   : TDateTime ;
    //SG6 26.01.05 Gestion mode Croisaxe
    lStSQL  : String ;
begin
  //SG6 26.01.05 Gestion mode croisaxe
  lStSQL := '';
  if not (VH^.AnaCroisaxe) then lStSQL := 'Y_AXE="' + Axe + '" AND ';

  For lInEcr := 0 to ( TobTTC.Detail.Count - 1 ) do
    begin

    lTobEcr := TobTTC.Detail[ lInEcr ] ;
    lNowH   := NowH ;

    // Suppression de l'analytique de la pièce concernée
    For lInAna := 0 to ( TobSuppr.Detail.Count - 1 ) do
      begin
      lTobAna := TobSuppr.Detail[ lInAna ] ;
      lStReq := 'DELETE FROM ANALYTIQ WHERE ' + lStSQL
                        + ' Y_NUMLIGNE=' + IntToStr( lTobAna.GetValue('Y_NUMLIGNE') )
                        + ' AND Y_NUMVENTIL=' + IntToStr( lTobAna.GetValue('Y_NUMVENTIL') )
                        + ' AND ' + WhereEcritureTob( tsAnal, lTobAna, False ) ;
      ExecuteSQL( lStReq ) ;
      MajSoldeSectionTOB ( lTOBAna, False ) ;
      end ;

    // Insertion de la nouvelle ventilation
    For lInAna := 0 to ( lTobEcr.Detail.Count - 1 ) do
      begin
      lTobAna := lTobEcr.Detail[ lInAna ] ;
      lTobAna.InsertDB(nil);
      MajSoldeSectionTOB ( lTOBAna, True ) ;
      end ;

    // Update date modif de l'écriture générale ( toutes lignes )
    lStReq := 'UPDATE ECRITURE SET E_DATEMODIF = "' + UsTime(lNowH)+ '" WHERE ' + WhereEcritureTob( tsGene, lTobEcr, False ) ;
    ExecuteSQL( lStReq ) ;

    end ;

end;

function TTraitementIAS14.GetCoefTVA ( vTobLigne : TOB ) : Double ;
var lQTva : TQuery ;
begin

  Result := 1.0 ;
  if vTobLigne.GetValue('E_TYPEMVT')<>'HT' then Exit ;

  lQTva := OpenSQL( 'SELECT TV_TAUXACH, TV_TAUXVTE FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1"' +
                      ' AND TV_REGIME="'   + vTobLigne.GetValue('E_REGIMETVA') + '"' +
                      ' AND TV_CODETAUX="' + vTobLigne.GetValue('E_TVA') + '"', True ) ;

  if not lQTva.Eof then
    if vTobLigne.GetValue('G_NATUREGENE') = 'PRO'
       then Result := 1.0 + (lQTva.FindField('TV_TAUXVTE').AsFloat / 100.0)
       else Result := 1.0 + (lQTva.FindField('TV_TAUXACH').AsFloat / 100.0) ;

  Ferme( lQTva ) ;

end;

procedure TTraitementIAS14.ReventileLigne( vTobEcr : TOB );
Var lTotalQte1       : Double ;
    lTotalQte2       : Double ;
    lEcrQte1         : Double ;
    lEcrQte2         : Double ;
    lTobAna  			   : TOB ;
    lTobLigneV		   : TOB ;
    lInCpt           : Integer ;
//    lTotalDev        : Double ;
//    lTotal           : Double ;
    lPourcent			   : Double ;
    lTotalPourcent   : Double ;
begin

  lEcrQte1        := vTobEcr.Detail[0].GetValue('Y_QTE1') ;
  lEcrQte2        := vTobEcr.Detail[0].GetValue('Y_QTE2') ;
//  lTotalDev       := 0 ;
//  lTotal          := 0 ;
  lTotalQte1      := 0 ;
  lTotalQte2      := 0 ;
  lTotalPourcent  := 0 ;

  // Récupération de la ventilation initiale sur la section d'attente
  TobSuppr.ClearDetail ;
  vTobEcr.Detail[0].ChangeParent( TobSuppr, -1 ) ;
  vTobEcr.ClearDetail ; //  au cas où ?

  // Affectation de la ventilation
	for lInCpt := 0 to TobVentil.Detail.Count - 1 do
    begin
    lTobLigneV  := TobVentil.Detail[ lInCpt ] ;

    // Init Tob Ana
    lTobAna := CGetNewTOBAna( StrToInt( Axe[2] ) , vTobEcr ) ;
		InitCommunObjAnalNew( vTobEcr, lTobAna ) ;
    RecopieInfos( lTobAna, TobSuppr.Detail[ (TobSuppr.Detail.Count - 1) ] ) ;

    // Affectation des montants / pourcentages
    SetVentilLigne( lTobAna, lTobLigneV, vTOBEcr.GetValue('E_DEBIT')<>0 ) ;

    // Pour éviter pb d'arrondi, on recalcul automatiquement le pourcentage de la dernière ligne par différentiel
    if lInCpt = ( TobVentil.Detail.Count - 1 ) then
      lTobAna.PutValue( 'Y_POURCENTAGE' , 100.0 - lTotalPourcent ) ;

  (* ANCIENNE VERSION AVEC CALCUL SUR POURCENT
    // Mise en place des données reventilations
    VentilLigneTOB( lTOBAna,
                    lTobLigneV.GetValue('V_SECTION'),
                    lTobLigneV.GetValue('V_NUMEROVENTIL'),
                    RDev.Decimale,
                    lPourcent,
                    vTOBEcr.GetValue('E_DEBIT')<>0
                    ) ;

    // Pour éviter pb d'arrondi, on recalcul automatiquement les montants de la dernière ligen par différence de la somme total
    if lInCpt = ( TobVentil.Detail.Count - 1 ) then
      if vTOBEcr.GetValue('E_DEBIT')<>0 then
        begin
        lTobAna.PutValue('Y_DEBITDEV', lTobAna.GetValue('Y_TOTALDEVISE')   - lTotalDev ) ;
        lTobAna.PutValue('Y_DEBIT',    lTobAna.GetValue('Y_TOTALECRITURE') - lTotal ) ;
        end
      else
        begin
        lTobAna.PutValue('Y_CREDITDEV', lTobAna.GetValue('Y_TOTALDEVISE')   - lTotalDev ) ;
        lTobAna.PutValue('Y_CREDIT',    lTobAna.GetValue('Y_TOTALECRITURE') - lTotal ) ;
        end ;
    *)

    lPourcent   := lTobAna.GetValue('Y_POURCENTAGE') ;

    // Gestion des Qtés
    if lEcrQte1 <> 0 then
      if lInCpt = (TobVentil.Detail.Count - 1)
        then lTobAna.PutValue('Y_QTE1', lEcrQte1 - lTotalQte1)
        else lTobAna.PutValue('Y_QTE1', Arrondi( lEcrQte1 * lPourcent / 100 , 3 ) ) ;

    if lEcrQte2 <> 0 then
      if lInCpt = (TobVentil.Detail.Count - 1)
        then lTobAna.PutValue('Y_QTE2', lEcrQte2 - lTotalQte2)
        else lTobAna.PutValue('Y_QTE2', Arrondi( lEcrQte2 * lPourcent / 100 , 3 ) ) ;

    // MAJ totaux qte
    lTotalQte1       := lTotalQte1     + lTobAna.GetValue('Y_QTE1') ;
    lTotalQte2       := lTotalQte2     + lTobAna.GetValue('Y_QTE2') ;

    // MAJ Total Pourcentage
    lTotalPourcent   := lTotalPourcent + lPourcent ;

    // MAJ totaux montant
//    lTotalDev := lTotalDev + lTobAna.GetValue('Y_DEBITDEV') + lTobAna.GetValue('Y_CREDITDEV') ;
//    lTotal    := lTotal    + lTobAna.GetValue('Y_DEBIT')    + lTobAna.GetValue('Y_CREDIT') ;

    end ;

end ;

procedure TTraitementIAS14.RecopieInfos(vTobAna, vTobAtt: TOB);
begin

  vTobAna.PutValue('Y_AUXILIAIRE',       vTobAtt.GetValue('Y_AUXILIAIRE') ) ;

  vTobAna.PutValue('Y_LIBRETEXTE0',       vTobAtt.GetValue('Y_LIBRETEXTE0') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE1',       vTobAtt.GetValue('Y_LIBRETEXTE1') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE2',       vTobAtt.GetValue('Y_LIBRETEXTE2') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE3',       vTobAtt.GetValue('Y_LIBRETEXTE3') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE4',       vTobAtt.GetValue('Y_LIBRETEXTE4') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE5',       vTobAtt.GetValue('Y_LIBRETEXTE5') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE6',       vTobAtt.GetValue('Y_LIBRETEXTE6') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE7',       vTobAtt.GetValue('Y_LIBRETEXTE7') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE8',       vTobAtt.GetValue('Y_LIBRETEXTE8') ) ;
  vTobAna.PutValue('Y_LIBRETEXTE9',       vTobAtt.GetValue('Y_LIBRETEXTE9') ) ;

  vTobAna.PutValue('Y_TABLE0',            vTobAtt.GetValue('Y_TABLE0') ) ;
  vTobAna.PutValue('Y_TABLE1',            vTobAtt.GetValue('Y_TABLE1') ) ;
  vTobAna.PutValue('Y_TABLE2',            vTobAtt.GetValue('Y_TABLE2') ) ;
  vTobAna.PutValue('Y_TABLE3',            vTobAtt.GetValue('Y_TABLE3') ) ;

  vTobAna.PutValue('Y_LIBREMONTANT0',     vTobAtt.GetValue('Y_LIBREMONTANT0') ) ;
  vTobAna.PutValue('Y_LIBREMONTANT1',     vTobAtt.GetValue('Y_LIBREMONTANT1') ) ;
  vTobAna.PutValue('Y_LIBREMONTANT2',     vTobAtt.GetValue('Y_LIBREMONTANT2') ) ;
  vTobAna.PutValue('Y_LIBREMONTANT3',     vTobAtt.GetValue('Y_LIBREMONTANT3') ) ;

  vTobAna.PutValue('Y_LIBREDATE',         vTobAtt.GetValue('Y_LIBREDATE') ) ;

  vTobAna.PutValue('Y_LIBREBOOL0',        vTobAtt.GetValue('Y_LIBREBOOL0') ) ;
  vTobAna.PutValue('Y_LIBREBOOL1',        vTobAtt.GetValue('Y_LIBREBOOL1') ) ;

  vTobAna.PutValue('Y_CONFIDENTIEL',      vTobAtt.GetValue('Y_CONFIDENTIEL') ) ;

end;

function TTraitementIAS14.GetNbLignesTraitees: Integer;
begin
  Result := NbLignes ;
end;

procedure TTraitementIAS14.SetVentilLigne( vTobAna, vTobVent : TOB ; vBoDebit : Boolean );
var lSection  : String ;
    lNumVent  : Integer ;
    lPourcent : Double ;
    lMontant  : Double ;
    lInCpt    : Integer ;
begin

  lSection  := vTobVent.GetValue('V_SECTION') ;
  lNumVent  := vTobVent.GetValue('V_NUMEROVENTIL') ;
//  lPourcent := vTobVent.GetValue('V_TAUXMONTANT') ;
  lMontant  := vTobVent.GetValue('V_MONTANT') ;


  vTobAna.PutValue( 'Y_SECTION',    lSection) ;
  vTobAna.PutValue( 'Y_NUMVENTIL',  lNumVent) ;

  if vBoDebit                // Débit     Crédit
    then CSetMontants( vTobAna, lMontant, 0,        RDev, True )
    else CSetMontants( vTobAna, 0,        lMontant, RDev, False ) ;

  lPourcent := Arrondi( 100.0 * lMontant / vTobAna.GetValue('Y_TOTALDEVISE') , ADecimP ) ;

  vTobAna.PutValue( 'Y_POURCENTAGE', lPourcent ) ;

  //SG6 26.01.05 Gestion mode croisaxe
  for lInCpt := 1 to MaxAxe do
  begin
    vTobAna.PutValue( 'Y_SOUSPLAN' + IntToStr(lInCpt), vTobVent.GetValue('V_SOUSPLAN' + IntToStr(lInCpt)) );
  end;

end;

end.
