unit CPProcBanqueAuxiliaire ;

interface

uses
  {$IFNDEF EAGLCLIENT}
    MajTable,
    uDbxDataSet,
  {$ENDIF EAGLCLIENT}
  SysUtils,
  ParamSoc,
  {$IFNDEF EAGLSERVER}
  {$IFDEF COMPTA}
  {$IFNDEF ENTREPRISE}
  galOutil,
  {$ENDIF ENTREPRISE}
  {$ENDIF}
  {$ENDIF EAGLSERVER}
  {$IFDEF NOVH}
  ULibCpContexte,
  {$ENDIF NOVH}
  HMsgBox ,
  Classes,
  uTob,
  HEnt1,
  HCtrls,
  StdCtrls
  ;

  Function IsCoherenceComptesBancairesSurListeCollectifs(theTOBERREURList : TOB) : Boolean ;
  Function IsCoherenceComptesBancairesPourGeneral(theTOBERREURList : TOB ; theGENERAL : String) : Boolean ;

  Function IsAuxiliaireCompteBancaire(theAUXILIAIRE : String ; TOBTIERS : TOB = Nil ; TOBGENERAUX : TOB = Nil) : Boolean ;

  Function IsCoherenceCompteGeneralCodeBanque(theGENERAL : String ; theCODEBANQUE : String) : Boolean ;

  Function IsTiersBanquePourGeneral(theAUXILIAIRE : String ; theTOBTIERS : TOB) : Boolean ; overload ;
  Function IsTiersBanquePourGeneral(theGENERAL, theAUXILIAIRE : String) : Boolean ;  overload ;

{$IFNDEF EAGLCLIENT}
  Function IsCompteBancaireAuxiliarise(theGENERAL : String ; theTQRYGENERAUX : TQuery) : Boolean ; overload ;
  Function IsCompteBancaireAuxiliarise(theGENERAL : String ; TOBGENERAUX : TOB = Nil) : Boolean ;   overload ;
{$ELSE}
  Function IsCompteBancaireAuxiliarise(theGENERAL : String ; TOBGENERAUX : TOB = Nil) : Boolean ;
{$ENDIF}

  function IsCompteBancaireAuxiliairePointable(theGENERAL : string ; TOBGENERAUX : TOB = Nil): Boolean;
  Function IsGestionCompteBancaireAuxiliarisee() : Boolean ;
  Function IsPresenceRIBBancaire(var theGENERAL: String ; theAUXILIAIRE : String ; boAfficheMessageFg : Boolean = True) : Boolean ;
  Function IsPossibleCreateComtptesBancaires(bMultSocFg : Boolean) : Boolean ;

  function PresenceInTable(Fichier: string ; ChampList : array of string; Valeurs: array of hstring): Boolean;

  Function RechercherCpteCollectifPourAuxiliaireBanque(theAUXILIAIRE : string): String ;
  Function RechercherCodeBanquePourGeneral(theGENERAL : string): String ;

  Procedure RechercherContrepartiePourUnJournal(theJNL : string ; var theGENERAL, theAUXILIAIRE : String) ;

implementation

uses
{$IFNDEF PGIIMMO}    //* 3P8 V9IMMO
  {$IFNDEF IMP2}
  Commun
  {$ENDIF IMP2}
{$ENDIF}
//  UtilCBP
  ;

function ChampToType_ (TT : String) : string;
begin
  	Result := ChampToType(TT);
end;
{-------------------------------------------------------------------------------}
Function RechercherTexteMessage(theErrorCode : Integer) : String ;
BEGIN
  case theErrorCode of
    1 : result := TraduireMemoire('Il manque des comptes bancaires pour les tiers associés à ce compte banque (*V001).') ;
    2 : result := TraduireMemoire('Il manque un compte bancaire pour le tiers (*V002) du compte banque (*V001).') ;
    3 : result := TraduireMemoire('Vous devez avoir le même établissement bancaire entre tous le tiers d''un compte (*V001).')
                + TraduireMemoire(' Références comptes bancaires (*V002-*V003)') ;
    4 : result := TraduireMemoire('Vous devez renseigner soit le compte général soit le compte auxiliaire soit les deux.') ;
    5 : result := TraduireMemoire('Le compte général *V001 n''est pas un compte de type banque.') ;
    6 : result := TraduireMemoire('Le compte général *V001 ne possède pas de RIB bancaire.') ;
  end ;
END ;

{-------------------------------------------------------------------------------}
Function RechercherMessage(theErrorCode : Integer ; theValueList : Array of String) : String ;
var value : String ;
    indice, index  : Integer ;
BEGIN
 if (theErrorCode <> 0) then
 begin

  result := RechercherTexteMessage(theErrorCode) ;
  if (result = '') then
  begin
   result := TraduireMemoire('Code erreur non trouvé ! - "' + InttoStr(theErrorCode) + '".') ;
   Exit ;
  end ;

  result := TraduireMemoire(result) ;

  // Insertion des paramètres dans le message

  indice := 0  ;
  for index := Low(theValueList) to high(theValueList) do
  begin
   value := theValueList[index] ;
   if (value = '') then Exit ;
   Inc(indice)  ;
   result := FindEtReplace(result, '*V' + FormatFloat('000', indice), value, True) ;
  end ;

 end
 else result := TraduireMemoire('Code Erreur ' + IntToStr(theErrorCode) +  ' non trouvé !!!') ;
END ;

{------------------------------------------------------------------------------}
function EstPointageSurTreso : Boolean;
begin
{$IFNDEF PGIIMMO}    //* 3P8 V9IMMO
{$IFNDEF IMP2}
  Result := IsTresoMultiSoc or
            ((ctxTreso in V_PGI.PGIContexte) and
            GetParamSocSecur('SO_TRPOINTAGETRESO', False) and
            not GetParamSocSecur('SO_POINTAGEJAL', False));
  if Result and not GetParamSocSecur('SO_TRPOINTAGETRESO', False) then
    SetParamSoc('SO_TRPOINTAGETRESO', True);
{$ELSE}
  result := False ;
{$ENDIF IMP2}
{$ELSE}
  result := False ;
{$ENDIF}
end;


{-------------------------------------------------------------------------------
Fonction ouvrant à la gestion des comptes auxiliaires sur les comptes bancaires

Gestion qui ne fonctionne qu'en mode entreprise

}
Function IsGestionCompteBancaireAuxiliarisee() : Boolean ;
BEGIN
  result := False ;

{$IFDEF CBP900}
  if (ctxTreso in V_PGI.PGIContexte)
   or (CtxPCL  in V_PGI.PGIContexte)
   or EstPointageSurTreso then
  begin
    Exit ;
  end ;

  result := GetParamSocSecur('SO_CP_BANQUEAUXILIARISEE', False) ;
{$ENDIF}
END ;

{------------------------------------------------------------------------------
  Contrôle si le compte passé en paramètre est un compte auxiliarisable
}
{$IFNDEF EAGLCLIENT}
Function IsCompteBancaireAuxiliarise(theGENERAL : String ; theTQRYGENERAUX : TQuery) : Boolean ;
var
 TOBGENERAUX : TOB ;
begin
  result := false ;
  if Not IsGestionCompteBancaireAuxiliarisee() then Exit ;

  TOBGENERAUX := TOB.Create('', Nil, -1) ;
  TOBGENERAUX.LoadDetailDB('GENERAUX', '', '', theTQRYGENERAUX, True) ;
  result := IsCompteBancaireAuxiliarise(theGENERAL, TOBGENERAUX) ;
  FreeAndNil(TOBGENERAUX) ;
end ;
{$ENDIF}
Function IsCompteBancaireAuxiliarise(theGENERAL : String ; TOBGENERAUX : TOB = Nil) : Boolean ;
var
  WTOBLIST      : TOB ;
  WTOB          : TOB ;
  SQL           : String ;
  QGENERAUX     : TQuery ;
  G_NATUREGENE  : String ;
  G_GENERAL     : String ;
  G_COLLECTIF   : String3 ;
BEGIN
  result    := False ;
  WTOBLIST  := Nil ;
  WTOB      := Nil ;

  try
    if Not IsGestionCompteBancaireAuxiliarisee() then Exit ;

    //* Si on lui passe la table des generaux, on controle si compte de banque
    //* sinon on va rechercher dans la base de données.

    if (TOBGENERAUX <> Nil) then
    begin
      if (TOBGENERAUX.Detail.Count <> 0) then
        WTOB := TOBGENERAUX.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [theGENERAL, 'BQE'], True)
      else
        if (TOBGENERAUX.GetValue('G_GENERAL') = theGENERAL)
         and (TOBGENERAUX.GetValue('G_NATUREGENE') = 'BQE')  then WTOB := TOBGENERAUX ;
    end ;

    if (WTOB = Nil) then
    begin
      SQL := 'SELECT * FROM GENERAUX WHERE G_GENERAL = "' + theGENERAL + '" AND G_NATUREGENE = "BQE"' ;
      QGENERAUX := OpenSQL(SQL, True) ;
      If Not QGENERAUX.EOF then
      begin
        WTOBLIST := TOB.Create('', Nil, -1);
        WTOBLIST.LoadDetailDB('GENERAUX', '', '', QGENERAUX, True) ;
        WTOB := WTOBLIST.Detail[WTOBLIST.Detail.Count -1] ;
      end ;
      Ferme(QGENERAUX) ;
    end ;

    if (WTOB = Nil) then Exit ;

    G_GENERAL     := WTOB.GetValue('G_GENERAL') ;
    G_NATUREGENE  := WTOB.GetValue('G_NATUREGENE') ;
    G_COLLECTIF   := WTOB.GetValue('G_COLLECTIF') ;

    if (G_GENERAL <> theGENERAL) or (G_NATUREGENE <> 'BQE') or (G_COLLECTIF <> 'X') then Exit ;

    result    := True ;
  finally
    if (WTOBLIST <> Nil) then FreeAndNil(WTOBLIST)  ;
  end ;
END ;

{------------------------------------------------------------------------------
  Contrôle si le tiers passé en paramètre est un compte de banque
}
Function IsAuxiliaireCompteBancaire(theAUXILIAIRE : String ; TOBTIERS : TOB = Nil ; TOBGENERAUX : TOB = Nil) : Boolean ;
var
  WTOBLIST      : TOB ;
  WTOB          : TOB ;
  SQL           : String ;
  QTIERS        : TQuery ;
  T_NATUREAUXI  : String ;
  T_AUXILIAIRE  : String ;
  T_COLLECTIF   : String ;
BEGIN
  result    := False ;
  WTOBLIST  := Nil ;
  WTOB      := Nil ;

  try
    if Not IsGestionCompteBancaireAuxiliarisee() then Exit ;

    //* Si on lui passe la table des tiers, on controle si compte de banque
    //* sinon on va rechercher dans la base de données.

    if (TOBTIERS <> Nil) then
    begin
      if (TOBTIERS.Detail.Count <> 0) then
        WTOB := TOBTIERS.FindFirst(['T_AUXILIAIRE', 'T_NATUREAUXI'], [theAUXILIAIRE, 'DIV'], True)
      else
        if (TOBTIERS.GetValue('T_AUXILIAIRE') = theAUXILIAIRE) then WTOB := TOBTIERS ;
    end ;

    if (WTOB = Nil) then
    begin
      SQL := 'SELECT * FROM TIERS WHERE T_AUXILIAIRE = "' + theAUXILIAIRE + '" AND T_NATUREAUXI = "DIV"' ;
      QTIERS := OpenSQL(SQL, True) ;
      If Not QTIERS.EOF then
      begin
        WTOBLIST := TOB.Create('', Nil, -1);
        WTOBLIST.LoadDetailDB('', '', '', QTIERS, True) ;
        WTOB := WTOBLIST.Detail[WTOBLIST.Detail.Count -1] ;
      end ;
      Ferme(QTIERS) ;
    end ;

    if (WTOB = Nil) then Exit ;

    T_AUXILIAIRE  := WTOB.GetValue('T_AUXILIAIRE') ;
    T_NATUREAUXI  := WTOB.GetValue('T_NATUREAUXI') ;
    T_COLLECTIF   := WTOB.GetValue('T_COLLECTIF') ;

    if (T_AUXILIAIRE <> theAUXILIAIRE)
      or (T_NATUREAUXI <> 'DIV')
      or ((TOBGENERAUX <> Nil) and (TOBGENERAUX.GetValue('G_GENERAL') <> T_COLLECTIF))
      or Not IsCompteBancaireAuxiliarise(T_COLLECTIF, TOBGENERAUX) then Exit ;
    result    := True ;
  finally
    if (WTOBLIST <> Nil) then FreeAndNil(WTOBLIST)  ;
  end ;
END ;
{------------------------------------------------------------------------------
  Contrôle la présence ou non des BANQUECP en fonction des paramètres passés
}
//Function IsPresenceRIBBancaire(theGENERAL, theAUXILIAIRE : String ; boAfficheMessageFg : Boolean = True) : Boolean ;
Function IsPresenceRIBBancaire(var theGENERAL: String ; theAUXILIAIRE : String ; boAfficheMessageFg : Boolean = True) : Boolean ;
VAR
  SQL     : String ;
  QTIERS  : TQuery ;

  index         : Integer ;
  TIERSLIST     : TOB ;
  TIERS         : TOB ;
  T_COLLECTIF   : String ;
  T_AUXILIAIRE  : String ;
  ENTETE        : String ;
  MSG           : String ;
BEGIN
  result  := False ;

  if (theGENERAL = '') and  (theAUXILIAIRE = '') then
  begin
    ENTETE := TraduireMemoire('Attention - Function IsPresenceRIBBancaire()') ;
    MSG    := RechercherMessage(4, ['']) ;
    if boAfficheMessageFg then PGIError(MSG, ENTETE);
    Exit ;
  end ;

//* Si CBP7 -  IsGestionCompteBancaireAuxiliarisee() retourne toujours False
//* Si CBP9 -  IsGestionCompteBancaireAuxiliarisee() retourne en fonction du pamètre société

  if Not IsGestionCompteBancaireAuxiliarisee() then
  begin
    SQL := 'SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + theGENERAL + '" AND G_NATUREGENE="BQE"' ;
    If Not ExisteSQL(SQL) then
    begin
      ENTETE := TraduireMemoire('Attention') ;
      MSG    := RechercherMessage(5, [theGENERAL]) ;
      if boAfficheMessageFg then PGIError(MSG, ENTETE);
      Exit ;
    end ;

    SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE BQ_GENERAL = "' + theGENERAL + '" AND BQ_NODOSSIER="'+ V_PGI.NoDossier + '"' ;
    If Not ExisteSQL(SQL) then
    begin
      ENTETE := TraduireMemoire('Attention') ;
      MSG    := RechercherMessage(6, [theGENERAL]) ;
      if boAfficheMessageFg then PGIError(MSG, ENTETE);
      Exit ;
    end ;

    result := True ;
  end
  else
  begin
    SQL := 'SELECT T_AUXILIAIRE, T_COLLECTIF FROM TIERS WHERE T_NATUREAUXI = "DIV" '  ;

    if (theGENERAL <> '') then  SQL := SQL + ' AND T_COLLECTIF = "' + theGENERAL + '"' ;
    if (theAUXILIAIRE <> '') then SQL := SQL + ' AND T_AUXILIAIRE = "' + theAUXILIAIRE + '"' ;

    QTIERS := OpenSQL(SQL, True) ;
    if Not QTIERS.EOF then
    begin
      TIERSLIST := TOB.Create('', Nil, -1) ;
      TIERSLIST.LoadDetailDB('TIERS', '', '', QTIERS, True) ;
      for index := 0 to TIERSLIST.Detail.Count -1 do
      begin
        TIERS := TIERSLIST.Detail[index] ;
        if (TIERS <> Nil) then
        begin
          T_COLLECTIF  := TIERS.GetValue('T_COLLECTIF') ;
          T_AUXILIAIRE := TIERS.GetValue('T_AUXILIAIRE') ;

          if (theGENERAL = '') then theGENERAL :=  T_COLLECTIF ;

          SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE '
            +  ' BQ_GENERAL = "'    + T_COLLECTIF + '"'
            +  ' AND BQ_AUXILIAIRE = "' + T_AUXILIAIRE + '" AND BQ_NODOSSIER="'+ V_PGI.NoDossier + '"' ;

          if Not ExisteSQL(SQL) then
          begin
            FreeAndNil(TIERSLIST) ;
            Ferme(QTIERS) ;
            Exit ;
          end
        end ;
      end ;
      FreeAndNil(TIERSLIST) ;
      result := True ;
    end;
    Ferme(QTIERS) ;
  end ;
END ;

{------------------------------------------------------------------------------}
Procedure ChargerMessageErreur(theTOBERREURList : TOB; theErrorCode : Integer ; theValueList : Array of String) ;
var
  TOBERREUR : TOB ;
begin
  TOBERREUR := TOB.Create('ERREUR',         theTOBERREURList, -1 );
  TOBERREUR.AddChampSupValeur('MESSAGE',    RechercherMessage(theErrorCode,  theValueList) , False);
end ;


{------------------------------------------------------------------------------}
Function IsTiersBanquePourGeneral(theAUXILIAIRE : String ; theTOBTIERS : TOB) : Boolean ;
var
  SQL : String ;
  Qry : TQuery;
BEGIN
  result := False ;

  If Not IsGestionCompteBancaireAuxiliarisee() then Exit ;

  if (theTOBTIERS.GetString('T_AUXILIAIRE') = theAUXILIAIRE) and
     (theTOBTIERS.GetString('T_NATUREAUXI') = 'DIV') then
  begin
    if not theTOBTIERS.FieldExists('NATUREGENEDEFAUT') then
    begin
      theTOBTIERS.AddChampSup('NATUREGENEDEFAUT', False);
      SQL := 'SELECT G_NATUREGENE FROM GENERAUX'
          +  ' WHERE G_GENERAL = "' + theTOBTIERS.GetValue('T_COLLECTIF') + '"';
      Qry := OpenSQL(SQL, True);
      try
        if not Qry.Eof then
          theTOBTIERS.PutValue('NATUREGENEDEFAUT', Qry.FindField('G_NATUREGENE').AsString);
      finally
        Ferme(Qry);
      end;
    end;
    result := theTOBTIERS.GetString('NATUREGENEDEFAUT') = 'BQE'
  end ;
END ;
{------------------------------------------------------------------------------}
Function IsTiersBanquePourGeneral(theGENERAL, theAUXILIAIRE : String) : Boolean ;
var
  SQL : String ;
BEGIN
  SQL := 'SELECT  T_AUXILIAIRE FROM TIERS '
      +  ' LEFT JOIN GENERAUX ON G_GENERAL = T_COLLECTIF '
      +  ' WHERE '
      +  ' T_AUXILIAIRE    = "' + theAUXILIAIRE + '"'
      +  ' AND T_COLLECTIF = "' + theGENERAL + '"'
      +  ' AND T_NATUREAUXI  = "DIV" '
      +  ' AND G_COLLECTIF = "X"  AND G_NATUREGENE = "BQE" ' ;
  result := ExisteSQL(SQL) ;
END ;
{------------------------------------------------------------------------------}
Function IsPossibleCreateComtptesBancaires(bMultSocFg : Boolean) : Boolean ;
var
  SQL       : String ;
  qry       : TQuery ;
  wTOBLIST  : TOB ;
  wTOB      : TOB ;
  indexTob  : Integer ;
  T_COLLECTIF : String ;
  T_AUXILIAIRE : String ;
BEGIN
  result := False;

  //* Cas des non collectifs

  if bMultSocFg then
  begin
    SQL := 'SELECT G_GENERAL FROM GENERAUX WHERE '
         + ' (G_COLLECTIF = "-") '
         + ' AND '
         + ' ( ' 
         + ' (G_NATUREGENE = "BQE"  AND  (G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + V_PGI.NoDossier + '" AND (BQ_AUXILIAIRE = "" OR BQ_AUXILIAIRE IS NULL)))) '
         + ' OR '
         + ' (G_NATUREGENE = "DIV"  AND G_GENERAL IN (SELECT CLS_GENERAL FROM CLIENSSOC))'
         + ' )' ;
  end
  else
  begin
    SQL := 'SELECT G_GENERAL FROM GENERAUX WHERE '
        + 'G_NATUREGENE = "BQE"'
        + ' AND '
        + '(G_COLLECTIF = "-") '
        + ' AND '
        + '(G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + V_PGI.NoDossier + '" AND (BQ_AUXILIAIRE = "" OR BQ_AUXILIAIRE IS NULL)))' ;
  end;

  If ExisteSQL(SQL) then begin result := True ; Exit ; end ;  // Y a encore des création sur les banque non auxiliarisé

  //* Cas des collectifs

  if IsGestionCompteBancaireAuxiliarisee() then
  begin
    SQL := 'SELECT T_AUXILIAIRE, T_COLLECTIF FROM TIERS WHERE T_NATUREAUXI = "DIV"'
        + ' AND '
        + ' T_COLLECTIF IN (SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTIF = "X" AND G_NATUREGENE="BQE")' ;

    qry := OpenSQL(SQL, True) ;
    if Not qry.EOF then
    begin
      wTOBLIST  := TOB.Create('', Nil, -1) ;
      wTOBLIST.LoadDetailDB('', '', '', qry, True) ;
      for indexTob := 0 to wTOBLIST.Detail.Count -1 do
      begin
        wTOB := wTOBLIST.Detail[indexTob] ;
        if ( wTOB <> Nil) then
        begin
          T_COLLECTIF   :=  wTOB.GetValue('T_COLLECTIF') ;
          T_AUXILIAIRE  :=  wTOB.GetValue('T_AUXILIAIRE') ;
          SQL := 'SELECT BQ_CODE FROM BANQUECP WHERE '
              + 'BQ_GENERAL ="' +  T_COLLECTIF + '"'
              + ' AND '
              + 'BQ_AUXILIAIRE = "' + T_AUXILIAIRE + '"' ;
          SQL := SQL + ' AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"'  ;

          if Not ExisteSQL(SQL) then begin result := True ; break ; end ;
          
        end ;
      end ;
      FreeAndNil(wTOBLIST) ;
    end ;
    Ferme(Qry) ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function IsCoherenceComptesBancairesPourGeneral(theTOBERREURList : TOB ; theGENERAL : String) : Boolean ;
var
  SQL       : String ;

  indexTiers    : Integer ;
  QTIERS        : TQuery ;
  TOBTIERSLIST  : TOB ;
  TOBTIERS      : TOB ;
  T_AUXILIAIRE  : String ;

  indexBanqueCp   : Integer ;
  QBANQUECP       : TQuery ;
  TOBBANQUECPLIST : TOB ;
  TOBBANQUECP     : TOB ;
  BQ_BANQUE       : String ;
  savBQ_BANQUE    : String ;
  savBQ_CODE      : String ;
  BQ_CODE         : String ;

  BQ_GENERAL  : String ;
  BQ_AUXILIAIRE  : String ;
  BQ_DEVISE : String ;
  couple          : TStringlist ;
BEGIN

  if IsCompteBancaireAuxiliarise(theGENERAL) then
  begin
     couple := TStringlist.Create() ;

      //*** Voir si coherence dans les données

    SQL := 'SELECT T_AUXILIAIRE FROM TIERS WHERE '
        +   ' T_NATUREAUXI = "DIV" '
        +   ' AND '
        +   ' T_COLLECTIF = "' + theGENERAL + '"' ;

    QTIERS := OpenSQL(SQL, True) ;
    if Not QTIERS.EOF then
    begin
      TOBTIERSLIST := TOB.Create('', Nil,  -1) ;
      TOBTIERSLIST.LoadDetailDB('TIERS', '', '', QTIERS, True) ;

      //*** Voir si Nombre de tiers = nombre de banqueCp

      SQL := 'SELECT DISTINCT BQ_GENERAL, BQ_AUXILIAIRE FROM BANQUECP WHERE '
          +  ' BQ_GENERAL = "' + theGENERAL + '" AND BQ_NODOSSIER = "' + V_PGI.NoDossier +'"' ;
          
      if (ExecuteSQL(SQL) < TOBTIERSLIST.Detail.Count) then  ChargerMessageErreur(theTOBERREURList, 1, [theGENERAL]) ;

      savBQ_BANQUE  := '' ;
      savBQ_CODE    := '' ;

      for indexTiers := 0 to TOBTIERSLIST.Detail.Count -1 do
      begin
        TOBTIERS := TOBTIERSLIST.Detail[indexTiers] ;
        if (TOBTIERS <> Nil) then
        begin
          T_AUXILIAIRE := TOBTIERS.GetValue('T_AUXILIAIRE') ;
          if (T_AUXILIAIRE <> '') then
          begin

            SQL := 'SELECT BQ_CODE, BQ_BANQUE FROM BANQUECP WHERE '
                + ' BQ_GENERAL = "' + theGENERAL + '"'
                + ' AND '
                + ' BQ_AUXILIAIRE = "' + T_AUXILIAIRE + '"  AND BQ_NODOSSIER = "' + V_PGI.NoDossier +'"'  ;

            QBANQUECP := OpenSQL(SQL, True) ;
            if Not QBANQUECP.EOF then
            begin

              TOBBANQUECPLIST := TOB.Create('', Nil, -1) ;
              TOBBANQUECPLIST.LoadDetailDB('BANQUECP', '', '', QBANQUECP, True) ;

              for indexBanqueCp := 0 to TOBBANQUECPLIST.Detail.Count -1 do
              begin
                TOBBANQUECP :=  TOBBANQUECPLIST.Detail[indexBanqueCp];
                if (TOBBANQUECP <> Nil) then
                begin
                  BQ_BANQUE := TOBBANQUECP.GetValue('BQ_BANQUE') ;
                  BQ_CODE   := TOBBANQUECP.GetValue('BQ_CODE') ;

                  BQ_GENERAL := TOBBANQUECP.GetValue('BQ_GENERAL') ;
                  BQ_AUXILIAIRE := TOBBANQUECP.GetValue('BQ_AUXILIAIRE') ;
                  BQ_DEVISE := TOBBANQUECP.GetValue('BQ_DEVISE') ;

                  //* Contrôle de l'établissement bancaire doit etre identique sur tout les RIB du GENERAL

                  if (indexTiers = 0) and (savBQ_BANQUE = '')  then
                  begin
                       savBQ_BANQUE := BQ_BANQUE ;
                       savBQ_CODE   := BQ_CODE ;
                  end
                  else if (savBQ_BANQUE <> savBQ_BANQUE) then
                          ChargerMessageErreur(theTOBERREURList, 3, [theGENERAL, savBQ_CODE, BQ_CODE]) ;

                end ;
              end ;
              FreeAndNil(TOBBANQUECPLIST) ;
            end
            else ChargerMessageErreur(theTOBERREURList, 2, [theGENERAL, T_AUXILIAIRE]) ;
            Ferme(QBANQUECP);
          end ;
        end ;
      end ;
      FreeAndNil(TOBTIERSLIST) ;
    end ;
    Ferme(QTIERS);
    FreeAndNil(couple) ; 
  end ;

  result := (theTOBERREURList.Detail.Count = 0) ;
END ;
{------------------------------------------------------------------------------}
Function IsCoherenceComptesBancairesSurListeCollectifs(theTOBERREURList : TOB) : Boolean ;
var
  indexGeneraux     : Integer ;
  QRY               : TQuery ;
  SQL               : String ;
  TOBGENERAUXLIST   : TOB ;
  TOBGENERAUX       : TOB ;
  G_GENERAL         : String ;
  bFlagFg           : Boolean ;
BEGIN
  result := True ;
  SQL := 'SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE= "BQE" AND G_COLLECTIF="X"' ;
  QRY := OpenSQL(SQL, True) ;
  IF Not QRY.EOF then
  begin
    TOBGENERAUXLIST := TOB.Create('', Nil, -1) ;
    TOBGENERAUXLIST.LoadDetailDB('GENERAUX', '', '', QRY, True) ;
    for indexGeneraux := 0 to TOBGENERAUXLIST.Detail.Count -1 do
    begin
      TOBGENERAUX := TOBGENERAUXLIST.Detail[indexGeneraux] ;
      if (TOBGENERAUX <> Nil) then
      begin
        G_GENERAL := TOBGENERAUX.GetValue('G_GENERAL') ;
        bFlagFg   := IsCoherenceComptesBancairesPourGeneral(theTOBERREURList, G_GENERAL) ;
        if result  and Not bFlagFg then Result := bFlagFg ;
      end ;
    end ;
    FreeAndNil(TOBGENERAUXLIST) ;
  end ;
  Ferme(QRY) ;
END ;

{=============================================================================}
Function IsFautGuillemet(theChamp : String) : Boolean ;
var
  typ     : string ;
begin
  result := False ;

  Typ := ChampToType_(theChamp);

  if (Typ <> 'INTEGER')
   and (Typ <> 'SMALLINT')
   and (Typ <> 'DOUBLE')
   and (Typ <> 'RATE')
   and (Typ <> 'EXTENDED') then result := True ;
end ;

{------------------------------------------------------------------------------}
function PresenceInTable(Fichier: string ; ChampList : array of string; Valeurs: array of hstring): Boolean;
var
  SQL     : String ;
  Valeur  : String ;
  index   : Integer ;
begin
  SQL := 'SELECT (1) FROM ' + VireDbo(Fichier) + ' WHERE ' ;

  for index :=Low(ChampList) to high(ChampList) do
  begin
    if IsFautGuillemet(ChampList[index]) then Valeur := '"' + Valeurs[index]  + '"'
                                        else Valeur :=  Valeurs[index]  ;

    SQL := SQL + ' ' +  ChampList[index] + ' = ' + Valeur + ' ' ;
    if (index <> High(ChampList)) then   SQL := SQL + ' AND ' ;
  end ;

  Result := ExisteSQL(SQL);
end;

{------------------------------------------------------------------------------}
function IsCompteBancaireAuxiliairePointable(theGENERAL : string ; TOBGENERAUX : TOB = Nil): Boolean;
var
  WTOBLIST      : TOB ;
  WTOB          : TOB ;
  SQL           : String ;
  QGENERAUX     : TQuery ;
  G_NATUREGENE  : String ;
  G_GENERAL     : String ;
  G_POINTABLE   : String3 ;
  G_COLLECTIF   : String3 ;
BEGIN
  result    := False ;
  WTOBLIST  := Nil ;
  WTOB      := Nil ;

  try
    if Not IsGestionCompteBancaireAuxiliarisee() then Exit ;

    //* Si on lui passe la table des generaux, on controle si compte de banque
    //* sinon on va rechercher dans la base de données.

    if (TOBGENERAUX <> Nil) then
    begin
      if (TOBGENERAUX.Detail.Count <> 0) then
        WTOB := TOBGENERAUX.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [theGENERAL, 'BQE'], True)
      else
        if (TOBGENERAUX.GetValue('G_GENERAL') = theGENERAL) then WTOB := TOBGENERAUX ;
    end ;

    if (WTOB = Nil) then
    begin
      SQL := 'SELECT G_GENERAL, G_NATUREGENE, G_COLLECTIF, G_POINTABLE FROM GENERAUX WHERE G_GENERAL = "' + theGENERAL + '" AND G_NATUREGENE = "BQE"' ;
      QGENERAUX := OpenSQL(SQL, True) ;
      If Not QGENERAUX.EOF then
      begin
        WTOBLIST := TOB.Create('', Nil, -1);
        WTOBLIST.LoadDetailDB('GENERAUX', '', '', QGENERAUX, True) ;
        WTOB := WTOBLIST.Detail[WTOBLIST.Detail.Count -1] ;
      end ;
      Ferme(QGENERAUX) ;
    end ;

    if (WTOB = Nil) then Exit ;

    G_GENERAL     := WTOB.GetValue('G_GENERAL') ;
    G_NATUREGENE  := WTOB.GetValue('G_NATUREGENE') ;
    G_COLLECTIF   := WTOB.GetValue('G_COLLECTIF') ;
    G_POINTABLE   := WTOB.GetValue('G_COLLECTIF') ;

    if (G_GENERAL <> theGENERAL) or (G_NATUREGENE <> 'BQE') or (G_COLLECTIF <> 'X') then Exit ;
    if (G_POINTABLE <> 'X') then Exit ;

    result    := True ;
  finally
    if (WTOBLIST <> Nil) then FreeAndNil(WTOBLIST)  ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function RechercherCpteCollectifPourAuxiliaireBanque(theAUXILIAIRE : string): String ;
var
  SQL : String ;
  QRY : TQuery ;
BEGIN
  result        := '' ;

  if Not IsGestionCompteBancaireAuxiliarisee() then Exit ;

  SQL := 'SELECT T_COLLECTIF FROM TIERS ' ;
  SQL := SQL + ' LEFT JOIN GENERAUX ON G_GENERAL = T_COLLECTIF ' ;
  SQL := SQL + ' WHERE T_AUXILIAIRE = "' + theAUXILIAIRE + '" AND T_NATUREAUXI = "DIV" ';
  SQL := SQL + ' AND G_NATUREGENE = "BQE" ' ;

  QRY := OpenSQL(SQL, True) ;
  IF Not QRY.EOF then result := QRY.Fields[0].AsString ;
  Ferme(QRY) ;
END ;

{------------------------------------------------------------------------------}
Function IsCoherenceCompteGeneralCodeBanque(theGENERAL : String ; theCODEBANQUE : String) : Boolean ;
var
  SQL           : String ;
  QRY           : TQuery ;
  TBANQUECPLIST : TOB ;
  TBANQUECP     : TOB ;
  index         : Integer ;
  BQ_ETABBQ     : String ;
BEGIN
  result := True ;
  IF (theCODEBANQUE = '') then Exit ;
  
  SQL := 'SELECT * FROM BANQUECP WHERE BQ_GENERAL = "' +theGENERAL + '" AND BQ_NODOSSIER="'+ V_PGI.NoDossier + '"' ;
  QRY := OpenSQL(SQL, True) ;
  If Not QRY.EOF then
  begin
    TBANQUECPLIST := TOB.Create('', Nil, -1) ;
    TBANQUECPLIST.LoadDetailDB('', '', '', QRY, True) ;
    For index := 0 to TBANQUECPLIST.Detail.Count -1 do
    begin
      TBANQUECP := TBANQUECPLIST.Detail[index] ;
      if (TBANQUECP <> Nil) then
      begin
        BQ_ETABBQ := TBANQUECP.GetValue('BQ_BANQUE') ;
        if (BQ_ETABBQ <> theCODEBANQUE) then begin result := False ; break ; end ;
      end ;
    end ;
    FreeAndNil(TBANQUECPLIST) ;
  end ;
  Ferme(QRY) ;
END ;
{------------------------------------------------------------------------------}
Function RechercherCodeBanquePourGeneral(theGENERAL : string): String ;
var
  SQL : String ;
  QRY : TQuery ;
BEGIN
  result  := '' ;

  if Not IsCompteBancaireAuxiliarise(theGENERAL) then Exit ;

  SQL := 'SELECT BQ_BANQUE FROM BANQUECP '
      + ' WHERE BQ_GENERAL = "' + theGENERAL + '"'
      + ' AND BQ_NODOSSIER = "'+ V_PGI.NoDossier + '"'  ;

  QRY := OpenSQL(SQL, True) ;
  IF Not QRY.EOF then result := QRY.Fields[0].AsString ;
  Ferme(QRY) ;
END ;
{------------------------------------------------------------------------------}
Procedure RechercherContrepartiePourUnJournal(theJNL : string ; var theGENERAL, theAUXILIAIRE : String) ;
var
  SQL : String ;
  QRY : TQuery ;
BEGIN
  theGENERAL    := '' ;
  theAUXILIAIRE  := '' ;

  if Not IsCompteBancaireAuxiliarise(theGENERAL) then Exit ;

  SQL := 'SELECT J_CONTREPARTIE, J_CONTREPARTIEAUX FROM JOURNAL '
      + ' WHERE J_JOURNAL = "' + theJNL + '"' ;

  QRY := OpenSQL(SQL, True) ;
  IF Not QRY.EOF then
  BEGIN
     theGENERAL := QRY.Fields[0].AsString ;
     theAUXILIAIRE := QRY.Fields[0].AsString ;
  END ;
  
  Ferme(QRY) ;
END ;

end.
