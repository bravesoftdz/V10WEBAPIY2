{-------------------------------------------------------------------------------------
  Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
08.01.001.009  28/01/07  JP  Création de l'unité : Contient différentes fonctions et variables
                             utiles au nouveau pointage
08.01.001.009  14/02/07  JP  Ajout de la classe ancêtre aux deux saisies simplifiées : TofAncetreEcr
08.10.001.010  20/09/07  JP  FQ 21349 : annulationn du pointage de mouvements issus d'autres relevés
                             mais pointés lors de la session que l'on vient de supprimer
08.10.001.010  20/09/07  JP  FQ 21333 : Fenêtre de choix d'une date comptable
08.10.002.001  22/10/07  JP  FQ 21706 : Fonction de conversion d'un fichier sans retour de chariot
08.10.005.001  14/11/07  JP  Gestion des comptes pointables qui ne sont pas bancaires
--------------------------------------------------------------------------------------}
unit uInitBanque;

interface

uses
  {$IFNDEF EAGLCLIENT}
  uDbxDataSet,
  {$ENDIF EAGLCLIENT}
  SysUtils,
  hMsgBox,
  UTob,
  HCtrls,
  Hent1,
  Classes,
  cbptrace;



function InitialiserBanques() : Boolean;

function IsDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET  : String ; var BQ_AGENCE, BQ_DOMICILIATION : String) : Boolean ;  overload;
function IsDomiciliationBancaire(TRA_AGENCE : String ; TOBAGENCELIST : TOB) : Boolean ;  overload;

function IsEtablissementBancaire(PQ_BANQUE : String ; TOBBANQUELIST : TOB) : Boolean ; overload;
function IsEtablissementBancaire(BQ_ETABBQ : String ; var BQ_BANQUE : String) : Boolean ; overload;


implementation

{
********************************************************************************
* Traitement d'initialisation des banques.
*
*   On va lire les banques à RIBs et
*   on analyse le code banque par rapport à l'établissement bancaire
*      Si l'établissement bancaire n'est pas trouvé on crée un établissement bancaire
*   on fait la meme chose avec le code guichet et la domiciliation
*
********************************************************************************
}
{------------------------------------------------------------------------------}
function CodeSuivant3 ( CL : String ) : String ;
Var i : integer ;
BEGIN
  CL:=uppercase(Trim(CL)) ;
  if Length(CL)<3 then CL:='AAA' else
  BEGIN
    i:=3 ;
    While CL[i]='Z' do
    BEGIN
      CL[i]:='A' ;
      Dec(i) ;
    END ;
    CL[i]:=Succ(CL[i]) ;
  END ;

  CodeSuivant3 := Copy(CL,1,3) ;
END ;

{------------------------------------------------------------------------------}
Function IncString3(St : String) : String ;
var i : Integer ;
BEGIN
  If IsNumeric(St) Then
  BEGIN
    i:=StrToint(St) ; Inc(i) ;
    If (i <= 999) Then Result := FormatFloat('000',i)
                  Else Result := CodeSuivant3('') ;
  END
  Else
    Result := CodeSuivant3(St) ;
END ;

{------------------------------------------------------------------------------}
Function GetNextCode(boTableAgence : Boolean ; CodeEnCours : String = '') : String ;
Var
  SQL : String ;
  QRY : TQuery ;
  ST  : String ;
BEGIN
  If (CodeEnCours = '') Then
  BEGIN
    if boTableAgence then SQL := 'SELECT MAX(TRA_AGENCE) FROM AGENCE'
                     else SQL := 'SELECT MAX(PQ_BANQUE)  FROM BANQUES'   ;

    QRY := OpenSql(SQL, true) ;
    ST  := QRY .Fields[0].AsString ;
    If (ST = '') Then ST:='000' ;
    Ferme(QRY) ;
  END
  Else ST := CodeEnCours ;

  Result := IncString3(ST) ;
END ;

{------------------------------------------------------------------------------}
function IsEtablissementBancaire(PQ_BANQUE : String; TOBBANQUELIST : TOB) : Boolean ;
var
  SQL : String ;
  WTOB : TOB ;
begin
  WTOB := TOBBANQUELIST.FindFirst(['PQ_BANQUE'], [PQ_BANQUE], True) ;
  if Assigned(WTOB) then
  begin
    result := True ;
    Exit ;
  end ;

  SQL := 'SELECT PQ_BANQUE FROM BANQUES WHERE PQ_BANQUE = "' + PQ_BANQUE + '"' ;
  result := ExisteSQL(SQL) ;
end;
{------------------------------------------------------------------------------}
function IsEtablissementBancaire(BQ_ETABBQ : String ; var BQ_BANQUE : String) : Boolean ;
var
  SQL              : String ;
  QRY              : TQuery ;
  PQ_BANQUE        : String ;
begin
  Result := False ;
  try
    PQ_BANQUE := '' ;

    SQL := 'SELECT PQ_BANQUE FROM BANQUES WHERE PQ_ETABBQ = "' + BQ_ETABBQ + '"' ;
    if (BQ_BANQUE <> '') then SQL := SQL + ' AND  PQ_BANQUE = "' + BQ_BANQUE + '"' ;
    QRY := OpenSQL(SQL, True) ;
    if not QRY.EOF then
    begin
      PQ_BANQUE := QRY.Fields[0].AsString ;

      //* Si on a trouvé un autre etablissement bancaire, on le remplace par celui
      //* de la table BANQUES

      if (PQ_BANQUE <> BQ_BANQUE) then  BQ_BANQUE := PQ_BANQUE ;

      Result := True ;
   end ;
    Ferme(QRY) ;
  except
    on E : Exception do
    begin
      trace.traceError('ERROR','Execption (Initialisation des établissements et agences bancaires) Procedure IsEtablissementBancaire() Message : '+ E.Message) ;
    end;
  end ;
end;

{------------------------------------------------------------------------------}
function CreateEtablissementBancaire(BQ_ETABBQ, PQ_BANQUE : String ; TOBBANQUELIST : TOB) : String ;
var
  TOBBANQUES : TOB ;
  WTOB : TOB ;
begin
  Result := '' ;

  //* Si le code banque est présent dans la liste on retourne le code etablissement bancaire

  WTOB := TOBBANQUELIST.FindFirst(['PQ_ETABBQ'], [BQ_ETABBQ], True) ;
  if Assigned(WTOB) then
  begin
    result := WTOB.GetValue('PQ_BANQUE') ;
    Exit ;
  end ;

  //* Calculer la clé, si non trouvé, on prend ce code sinon on va recalculer le code
  //* jusqu'à ce que le code n'existe pas ni dans la tob, ni dans la tablme

  if (PQ_BANQUE = '') then PQ_BANQUE := GetNextCode(False, '') ;
  while True do
  begin
    if Not IsEtablissementBancaire(PQ_BANQUE, TOBBANQUELIST) then Break ;
    PQ_BANQUE := GetNextCode(False, PQ_BANQUE) ;
  end ;

  //* Ajout d'une banque dans la liste des banques

  TOBBANQUES := TOB.Create('BANQUES', TOBBANQUELIST, -1);
  TOBBANQUES.InitValeurs();
  TOBBANQUES.PutValue('PQ_BANQUE',  PQ_BANQUE );
  TOBBANQUES.PutValue('PQ_LIBELLE', BQ_ETABBQ + ' (A COMPLETER)');
  TOBBANQUES.PutValue('PQ_ABREGE',  BQ_ETABBQ );
  TOBBANQUES.PutValue('PQ_ETABBQ',  BQ_ETABBQ );

  result := PQ_BANQUE ;
end;

{------------------------------------------------------------------------------}
function CreateDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET, TRA_AGENCE, BQ_DOMICILIATION : String ; TOBAGENCELIST : TOB) : String ;
var
  WTOB      : TOB ;
  TOBAGENCE : TOB ;
begin
  Result := '' ;

  //* Si banque+guichet présents on recupère le code agence

  WTOB := TOBAGENCELIST.FindFirst(['TRA_ETABBQ', 'TRA_GUICHET'], [BQ_ETABBQ, BQ_GUICHET], True) ;
  if Assigned(WTOB) then
  begin
    result := WTOB.GetValue('TRA_AGENCE') ;
    Exit ;
  end ;

 //* Calculer la clé

  if (TRA_AGENCE = '') then TRA_AGENCE := GetNextCode(True, '') ;
  while True do
  begin
    if not IsDomiciliationBancaire(TRA_AGENCE, TOBAGENCELIST) then Break  ;
    TRA_AGENCE := GetNextCode(True, TRA_AGENCE) ;
  end ;


  TOBAGENCE := TOB.Create('AGENCE', TOBAGENCELIST, -1);
  TOBAGENCE.InitValeurs();

  TOBAGENCE.PutValue('TRA_AGENCE',  TRA_AGENCE);
  TOBAGENCE.PutValue('TRA_GUICHET',  BQ_GUICHET);

  IF (BQ_DOMICILIATION = '') then
  begin
    TOBAGENCE.PutValue('TRA_LIBELLE',  BQ_GUICHET + ' (A COMPLETER)');
    TOBAGENCE.PutValue('TRA_ABREGE',  BQ_GUICHET);
  end
  else
  begin
    TOBAGENCE.PutValue('TRA_LIBELLE', '(A COMPLETER) ' +  BQ_DOMICILIATION);
    TOBAGENCE.PutValue('TRA_ABREGE',   '(A COMPLETER)');
  end;

  TOBAGENCE.PutValue('TRA_ETABBQ',  BQ_ETABBQ);
  TOBAGENCE.PutValue('TRA_BANQUE',  BQ_BANQUE);

  Result := TRA_AGENCE ;
end;
{------------------------------------------------------------------------------}
function IsDomiciliationBancaire(TRA_AGENCE : String ; TOBAGENCELIST : TOB) : Boolean ;
var
  SQL : String ;
  WTOB : TOB ;
begin
  WTOB := TOBAGENCELIST.FindFirst(['TRA_AGENCE'], [TRA_AGENCE], False) ;
  if Assigned(WTOB) then
  begin
    Result := True ;
    Exit ;
  end ;

  SQL := 'SELECT TRA_AGENCE FROM AGENCE WHERE TRA_AGENCE = "' + TRA_AGENCE + '"' ;
  result := ExisteSQL(SQL) ;
end;
{------------------------------------------------------------------------------
En entrée :
-  le code établissement bancaire est toujours renseigné.
- La domiciliation est peut être renseignée.

Si aucune agence pour le code banque, code guichet correspondant à la code etablissement bancaire,
on crée toutes les agences de cet etablissement

Si des agences sont trouvée dans la table AGENCE
Si la domiciliation est renseignée dans banquecp ou pas mais différente de l'enregistrement, on ecrase la domiciliation
par l'info de la table. et on ne crée pas la domiciliation bancaire

Si aucune domiciliation n'est trouvée, on crée la domiciliation


}
function IsDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET  : String ; var BQ_AGENCE, BQ_DOMICILIATION : String) : Boolean ;
var
  SQL         : String ;
  QRY         : TQuery ;
  TRA_AGENCE  : string ;
  TRA_BANQUE  : String ;
  TRA_LIBELLE : String ;
  wTOBList    : TOB ;
  wTOB        : TOB ;
  index       : Integer;
Begin
  Result := False ; //* Création d'une domicialiation bancaire
  wTOBList := Nil ;
  try
    try
      TRA_AGENCE := '' ;
      TRA_BANQUE := '' ;

      SQL := 'SELECT TRA_AGENCE, TRA_BANQUE, TRA_LIBELLE FROM AGENCE '
          + ' WHERE TRA_ETABBQ = "' + BQ_ETABBQ + '"'
          + ' AND '
          + ' TRA_GUICHET = "' + BQ_GUICHET + '"' ;
      if (BQ_BANQUE <> '') then SQL := SQL + ' AND TRA_BANQUE =  "' + BQ_BANQUE + '" ' ;
      if (BQ_AGENCE <> '') then SQL := SQL + ' AND TRA_AGENCE =  "' + BQ_AGENCE + '" ' ;

      QRY := OpenSQL(SQL, True) ;
      if not QRY.EOF then
      begin
        wTOBList := TOB.Create('', nil, -1);
        wTOBList.LoadDetailDB('', '', '', QRY, true) ;

        for index := 0 to wTOBList.Detail.Count - 1 do
        begin
          wTOB := wTOBList.Detail[index] ;
          if Assigned(wTOB) then
          begin
            TRA_AGENCE  := wTOB.GetValue('TRA_AGENCE') ;
            TRA_BANQUE  := wTOB.GetValue('TRA_BANQUE') ;
            TRA_LIBELLE := wTOB.GetValue('TRA_LIBELLE') ;

            if (BQ_BANQUE <> '') and (BQ_BANQUE = TRA_BANQUE) then
            begin
              if (BQ_AGENCE <> TRA_AGENCE) then
              begin
                BQ_AGENCE         := TRA_AGENCE ;
                BQ_DOMICILIATION  := TRA_LIBELLE ;
                result            := True ; // Pas de création de l'agence
                Ferme(QRY);
                Exit ;
              end
              else
              begin
                BQ_DOMICILIATION  := TRA_LIBELLE ;
                result            := True ; // Pas de création de l'agence
              end;
            end;
          end;
        end;
      end ;
      Ferme(QRY) ;
    except
      on E : Exception do
      begin
        trace.traceError('ERROR', 'Execption (Initialisation des établissements et agences bancaires) Procedure IsDomicialiationBancaire(). Message : '+ E.Message) ;
    end;
  end ;
  finally
    if Assigned(wTOBList) then FreeAndNil(wTOBList);
  end ;
end;

{------------------------------------------------------------------------------}
function InitialiserBanques() : Boolean;
var
  SQL              : String ;
  QRY              : TQuery ;

  TOBBANQUECPLIST  : TOB ;
  TOBBANQUECP      : TOB ;

  TOBBANQUELIST    : TOB ;
  TOBAGENCELIST    : TOB ;
  wTOB             : TOB ;
  wTOBT            : TOB ;

  PQ_ETABBQ        : String ;
  TRA_ETABBQ       : String ;
  TRA_GUICHET      : String ;

  Index            : Integer ;
  BQ_ETABBQ        : String ;
  BQ_GUICHET       : String ;
  BQ_BANQUE        : String ;
  BQ_AGENCE        : String ;
  BQ_DOMICILIATION : String ;
  BQ_CODE          : String ;
  boMajBanqueCP    : Boolean ;
begin
//  trace.traceError('ERROR', '(Initialisation des établissements et agences bancaires) Procedure InitialiserBanques(). Message : DEBUT DU TRAITEMENT') ;

  result          := False ;
  boMajBanqueCP   := False ;
  TOBBANQUECPLIST := Nil ;
  TOBBANQUECP     := Nil ;

  TOBBANQUELIST   := Nil ;

  TOBAGENCELIST   := Nil ;

  try
    try
      //* FQ 10137 Fin

      SQL := 'SELECT * FROM BANQUES '  ;
      TOBBANQUELIST  := TOB.Create('', nil, -1) ;
      QRY := OpenSQL(SQL, True) ;
      if not QRY.Eof then TOBBANQUELIST.LoadDetailDB('BANQUES', '', '', QRY, True) ;
      Ferme(QRY);

      SQL := 'SELECT * FROM AGENCE '  ;
      TOBAGENCELIST  := TOB.Create('', nil, -1) ;
      QRY := OpenSQL(SQL, True) ;
      if not QRY.Eof then TOBAGENCELIST.LoadDetailDB('AGENCE', '', '', QRY, True) ;
      Ferme(QRY);

      //* FQ 10137 Fin

      //* Je ne prends que les comptes bancaire francais. c'est à dire ceux qui ont un
      //* code banque et un code guichet

      SQL := 'SELECT * FROM BANQUECP WHERE (BQ_ETABBQ <> "") OR (BQ_GUICHET <> "") '
          +  ' ORDER BY BQ_ETABBQ, BQ_GUICHET ' ;


      QRY := OpenSQL(SQL, True) ;
      if not QRY.EOF then
      begin
        TOBBANQUECPLIST  := TOB.Create('', nil, -1) ;
        TOBBANQUECPLIST.LoadDetailDB('BANQUECP', '', '', QRY, True) ;


        for index := 0 to TOBBANQUECPLIST.Detail.Count - 1 do
        begin
          TOBBANQUECP  := TOBBANQUECPLIST.Detail[index] ;
          if Assigned(TOBBANQUECP) then
          begin
            BQ_CODE         := TOBBANQUECP.GetValue('BQ_CODE') ;
            BQ_ETABBQ         := TOBBANQUECP.GetValue('BQ_ETABBQ') ;
            BQ_GUICHET        := TOBBANQUECP.GetValue('BQ_GUICHET') ;
            BQ_BANQUE         := TOBBANQUECP.GetValue('BQ_BANQUE') ;
            BQ_AGENCE         := TOBBANQUECP.GetValue('BQ_AGENCE') ;
            BQ_DOMICILIATION  := TOBBANQUECP.GetValue('BQ_DOMICILIATION') ;

            //* ----------------------------------------------------------------
            //* Gestion de l'établissement bancaire

            //* FQ 10137
            //* On regarde si l'établissement bancaire existe dans la table
            //* Si on regarde si le code banque n'est pas renseigné. Si oui on met l'établissement
            //* de BANQUES


            if (BQ_ETABBQ <> '') then
            begin
              wTOBT := TOBBANQUELIST.FindFirst(['PQ_BANQUE'], [BQ_BANQUE], True) ;
              if Assigned(wTOBT) then
              begin
                PQ_ETABBQ := wTOBT.GetValue('PQ_ETABBQ') ;
                if (PQ_ETABBQ = '') then
                begin
                  wTOBT.PutValue('PQ_ETABBQ', BQ_ETABBQ);
                  if not boMajBanqueCP then boMajBanqueCP := True ;
                end
                else if (PQ_ETABBQ <> BQ_ETABBQ) then
                begin
                  BQ_BANQUE := CreateEtablissementBancaire(BQ_ETABBQ, BQ_BANQUE, TOBBANQUELIST) ;
                  TOBBANQUECP.PutValue('BQ_BANQUE', BQ_BANQUE);
                  if not boMajBanqueCP then boMajBanqueCP := True ;
                end;
              end
              else
              begin
                wTOBT := TOBBANQUELIST.FindFirst(['PQ_ETABBQ'], [BQ_ETABBQ], True) ;
                if Not Assigned(wTOBT) then
                begin
                  if Not IsEtablissementBancaire(BQ_ETABBQ, BQ_BANQUE) then
                  begin
                    BQ_BANQUE := CreateEtablissementBancaire(BQ_ETABBQ, BQ_BANQUE, TOBBANQUELIST) ;
                    TOBBANQUECP.PutValue('BQ_BANQUE', BQ_BANQUE);
                    if not boMajBanqueCP then boMajBanqueCP := True ;
                  end
                  else
                  begin
                    TOBBANQUECP.PutValue('BQ_BANQUE', BQ_BANQUE);
                    if not boMajBanqueCP then boMajBanqueCP := True ;
                  end;
                end
                else
                begin
                  BQ_BANQUE := wTOBT.GetValue('PQ_BANQUE') ;
                  TOBBANQUECP.PutValue('BQ_BANQUE', BQ_BANQUE);
                  if not boMajBanqueCP then boMajBanqueCP := True ;
                end;
              end ;
            end ;

            //* ----------------------------------------------------------------
            //* Gestion de la domiciliation


            //* FQ 10137
            //* On regarde si l'agence bancaire existe dans la table
            //* Si on regarde si le code agence n'est pas renseigné. Si oui on met l'agence
            //* de AGENCE et on transfere la banque dans la TOB a mettre à jour.

            if (BQ_GUICHET <> '') then
            begin
              wTOBT := TOBAGENCELIST.FindFirst(['TRA_AGENCE'], [BQ_AGENCE], True) ;
              if Assigned(wTOBT) then
              begin
                TRA_GUICHET := wTOBT.GetValue('TRA_GUICHET') ;
                TRA_ETABBQ  := wTOBT.GetValue('TRA_ETABBQ') ;

               //* Si l'établissement bancaire est n'est pas renseigné

                if (TRA_ETABBQ = '') then
                begin
                  wTOBT.PutValue('TRA_ETABBQ', BQ_ETABBQ);

                  if (TRA_GUICHET = '') then  //* On prend l'agence
                  begin

                    wTOBT.PutValue('TRA_GUICHET', BQ_GUICHET);
                    if not boMajBanqueCP then boMajBanqueCP := True ;

                  end
                  else if (TRA_GUICHET <> BQ_GUICHET) then // Si guichet différents on crée une nouvelle agence
                  begin
                    BQ_AGENCE := CreateDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET, BQ_AGENCE, BQ_DOMICILIATION, TOBAGENCELIST) ;
                    if not boMajBanqueCP then boMajBanqueCP := True ;
                    TOBBANQUECP.PutValue('BQ_AGENCE', BQ_AGENCE);

                  end;

                end
                else //* Cas ou l'établissement est renseigné
                begin

                  if (TRA_ETABBQ = BQ_ETABBQ) then // Si les établissements bancaires sont identiques
                  begin
                    if (TRA_GUICHET = '') then   //* le guichet est renseigné
                    begin

                      wTOBT.PutValue('TRA_GUICHET', BQ_GUICHET);
                      if not boMajBanqueCP then boMajBanqueCP := True ;

                    end
                    else
                    begin

                      if (TRA_GUICHET <> BQ_GUICHET) then  // Si guichet différents on crée une nouvelle agence
                      begin

                        BQ_AGENCE := CreateDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET, BQ_AGENCE, BQ_DOMICILIATION, TOBAGENCELIST) ;
                        if not boMajBanqueCP then boMajBanqueCP := True ;
                        TOBBANQUECP.PutValue('BQ_AGENCE', BQ_AGENCE);

                      end ;
                    end ;
                  end
                  else // Si les établissements bancaires sont différents, on crée une nouvelle agence
                  begin

                    BQ_AGENCE := CreateDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET, BQ_AGENCE, BQ_DOMICILIATION, TOBAGENCELIST) ;
                    if not boMajBanqueCP then boMajBanqueCP := True ;
                    TOBBANQUECP.PutValue('BQ_AGENCE', BQ_AGENCE);

                  end ;

                end  ;

              //* On alimente la domiciliation de banquecp si non renseignée

                if (BQ_DOMICILIATION = '') then
                begin
                  wTOB := TOBAGENCELIST.FindFirst(['TRA_AGENCE'],[BQ_AGENCE], True) ;
                  if Assigned(wTOB) then
                  begin
                    BQ_DOMICILIATION := wTOB.GetValue('TRA_LIBELLE') ;
                    TOBBANQUECP.PutValue('BQ_DOMICILIATION', BQ_DOMICILIATION);
                  end ;
                end ;

              end
              else
              begin
                wTOBT :=  TOBAGENCELIST.FindFirst(['TRA_ETABBQ', 'TRA_GUICHET'], [BQ_ETABBQ, BQ_GUICHET], True) ;
                if Not Assigned(wTOBT)  then
                begin
                  if Not IsDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET, BQ_AGENCE, BQ_DOMICILIATION) then
                  begin
                    BQ_AGENCE := CreateDomiciliationBancaire(BQ_ETABBQ, BQ_BANQUE, BQ_GUICHET, BQ_AGENCE, BQ_DOMICILIATION, TOBAGENCELIST) ;
                    if not boMajBanqueCP then boMajBanqueCP := True ;
                    TOBBANQUECP.PutValue('BQ_AGENCE', BQ_AGENCE);

                    if (BQ_DOMICILIATION = '') then
                    begin
                      wTOB := TOBAGENCELIST.FindFirst(['TRA_AGENCE'],[BQ_AGENCE], True) ;
                      if Assigned(wTOB) then
                      begin
                        BQ_DOMICILIATION := wTOB.GetValue('TRA_LIBELLE') ;
                        TOBBANQUECP.PutValue('BQ_DOMICILIATION', BQ_DOMICILIATION);
                      end ;
                    end;
                  end
                  else
                  begin
                     if not boMajBanqueCP then boMajBanqueCP := True ;
                     TOBBANQUECP.PutValue('BQ_AGENCE', BQ_AGENCE);

                    if (BQ_DOMICILIATION = '') then
                    begin
                      wTOB := TOBAGENCELIST.FindFirst(['TRA_AGENCE'],[BQ_AGENCE], True) ;
                      if Assigned(wTOB) then
                      begin
                        BQ_DOMICILIATION := wTOB.GetValue('TRA_LIBELLE') ;
                        TOBBANQUECP.PutValue('BQ_DOMICILIATION', BQ_DOMICILIATION);
                      end ;
                    end;
                  end ;
                end
                else
                begin
                  BQ_AGENCE :=  wTOBT.GetValue('TRA_AGENCE') ; // Agence trouvée
                  TOBBANQUECP.PutValue('BQ_AGENCE', BQ_AGENCE);

                  if (BQ_DOMICILIATION = '') then
                  begin
                    wTOB := TOBAGENCELIST.FindFirst(['TRA_AGENCE'],[BQ_AGENCE], True) ;
                    if Assigned(wTOB) then
                    begin
                      BQ_DOMICILIATION := wTOB.GetValue('TRA_LIBELLE') ;
                      TOBBANQUECP.PutValue('BQ_DOMICILIATION', BQ_DOMICILIATION);
                    end ;
                  end;
                  if not boMajBanqueCP then boMajBanqueCP := True ;
                end;
              end;
            end;
          end;
        end;
      end;
      Ferme(QRY)  ;

      if boMajBanqueCP then
      begin
        if (TOBBANQUELIST.Detail.Count <> 0) then TOBBANQUELIST.InsertOrUpdateDB();
        if (TOBAGENCELIST.Detail.Count <> 0) then TOBAGENCELIST.InsertOrUpdateDB();
        TOBBANQUECPLIST.UpdateDB();
      end;

      result := True ;
    trace.traceError('ERROR', '(Initialisation des établissements et agences bancaires) Procedure InitialiserBanques(). Message : FIN DU TRAITEMENT') ;
    except

      on E : Exception do
      begin
        trace.traceError('ERROR','Execption (Initialisation des établissements et agences bancaires) Procedure InitialiserBanques(). Message : '+ E.Message) ;
      end ;

    end ;
  finally
    if Assigned(TOBAGENCELIST) then FreeAndNil(TOBAGENCELIST);
    if Assigned(TOBBANQUELIST) then FreeAndNil(TOBBANQUELIST);
    if Assigned(TOBBANQUECP) then FreeAndNil(TOBBANQUECP);
    if Assigned(TOBBANQUECPLIST) then FreeAndNil(TOBBANQUECPLIST);
  end ;
end;

initialization
End.

