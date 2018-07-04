{
********************************************************************************

Objet ........ : Gestion de la table souche pointage

Attention .... : Cette gestion est valable à partir de la V9


--------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------
               16/04/08  3P8  Création de l'unité
               17/04/08  3P8  Suppression de la méthode loadData()
                              Déportation du code vers la méthode Create()
               21/04/08  3P8  Gestion des contrôles
                              Gestion de la création et de la suppression
                              Gestion des setters
               23/04/08  3P8  Gestion de la date d'exercice
               24/04/08  3P8  Gestion de la liste des banqueSh
********************************************************************************
}
Unit CPOBJBANQUESH ;

Interface

Uses StdCtrls,
{$IFNDEF EAGLCLIENT}
     uDbxDataSet,
{$ENDIF}
     Controls,
     Classes,
     uTob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox ;


{------------------------------------------------------------------------------}
 type
     TActionTableBq = (tBqRIEN, tBqCREAT, tBqMODIF, tBqDELETE);
     TActionECRAN = (tEcRIEN, tEcCREAT, tEcMODIF, tEcDELETE);

 type
     TBanqueSh = Class (TObject)
     private
        compteGeneral     : String ;
        compteAuxiliaire  : String ;

        Action       : TActionTableBq ;
        ActionEcran  : TActionECRAN ;

        EcraseLastNum : Boolean ;

        refPointage   : String ;

        TOBBANQUESH   : TOB ;
        TOBEXERCICE   : TOB ;
        TOBDATERELEVE : TOB ;

        ActionDefaire        : TActionTableBq ;
        ActionEcranDefaire   : TActionECRAN ;
        EcraseLastNumDefaire : Boolean ;

        TOBDEFAIRE   : TOB ;
        TOBEXDEFAIRE : TOB ;

        Procedure ChargerReferenceReleve(theDateReleve : TdateTime) ;

        Procedure UpdateDataBase() ;
        Procedure DeleteDataBase() ;
        Procedure SetCurrentNumero(theCurrentNumero : Integer)   ;

        Procedure InitReferencePointage(theDateReleve : TDateTime);

        Function  IsCode()  : Boolean ;
        Function  IsCodeExercice()  : Boolean ;
        Function  IsDomiciliation(theDomiciliation : String = '')  : Boolean ;
        Function  IsConstante()  : Boolean ;
        Function  IsTaille()  : Boolean ;
        Function  IsBorneSoucheOk()  : Boolean ;


        procedure readCompteGeneralAuxiliaire() ;
     published
        Constructor Create              () ;

        Destructor  Destroy             () ; override ;

        Function  Defaire               () : Boolean;
        Function  DeleteBD              () : Boolean;

        Function  GetAction             () : TActionTableBq;
        Function  GetActionEcran        () : TActionECRAN ;
        Function  GetBanque             () : String ;
        Function  GetCleRib             () : String ;
        Function  GetCode               () : String ;
        Function  GetCompte             () : String ;
        Function  GetContanteBancaire   () : String ;
        Function  GetCurrentNumero      () : Integer ;
        Function  GetCodeExercice       () : String ;
        Function  GetDateDebutExercice  () : TDateTime ;
        Function  GetDateExercice       () : TDateTime ;
        Function  GetDateFinExercice    () : TDateTime ;
        Function  GetDernierNumero      () : Integer ;
        Function  GetDomiciliation      () : String ;
        function  GetEcraseLastNum      () : Boolean ;
        Function  GetGuichet            () : String ;
        Function  GetPositionnementAnnee() : Integer ;
        Function  GetPremierNumero      () : Integer ;
        Function  GetReferencePointage  () : String ;
        Function  GetTaille             () : Integer ;

        function  GetTOBBANQUESH(): TOB ;

        Function  GetNextReferencePointage(theBUpdateBDFg : Boolean ; theDateReleve : TdateTime) : String ;

        Function  IsAllControler (var theNumChampError : Integer)  : Boolean ;
        Function  IsControler    (var theNumChampError : Integer ; theBoCode,  theBoExercice, boParametrageRef : Boolean)  : Boolean ;

        Procedure SetAction              (theAction : TActionTableBq) ;
        Procedure SetActionEcran         (theActionEcran : TActionECRAN)  ;
        Procedure SetCode                (theCode : String)  ;
        Procedure SetContanteBancaire    (theConstante : String) ;
        Procedure SetCodeExercice        (theCodeExercice : String)  ;
        Procedure SetDateExercice        (theDateExercice  : TdateTime) ;
        Procedure SetDernierNumero       (theLastSouche : Integer) ;
        Procedure SetDomiciliation       (theDomiciliation : String)  ;
        Procedure SetEcraseLastNum       (theEcraseLastNum : Boolean);
        Procedure SetPositionnementAnnee (theLgAnnee : Integer) ;
        Procedure SetPremierNumero       (theFirstSouche : Integer) ;
        Procedure SetTaille              (theTaille  : Integer) ;

        Function  UpdateBD() : Boolean ;
     public
        procedure InitialiserCreation    (theCode, theDomiciliation : String ; theCodeExercice : String = '') ; overload ;
        procedure InitialiserCreation    (theCode, theDomiciliation : String ; theDateExercice : TDateTime)  ; overload ;

        Function  LoadData(theCode, theCodeExercice : String) : Boolean ; overload ;
        Function  LoadData(theTOBBANQUESH : TOB) : Boolean  ;  overload ;
     end ;

{------------------------------------------------------------------------------}
 type
     TBanqueShList = Class (TObject)
     private
        bBqeAuxFg   : Boolean ;

        Liste       : TList ;
        TOBBANQUECP : TOB ;
     published

        Constructor Create          () ;
        Destructor  Destroy         () ; override ;

        Function LoadDataByCompte(theCptGENERAL, theCptAUXILIAIRE : String) : Boolean  ;
        Function LoadData(theCode : String;  theBANQUECP : TOB = Nil) : Boolean ;

        Function  Add               (theBanqueSh : TBanqueSh) : Boolean ;
        Function  Count             : Integer  ;
        Function  Detail            (theOccurenceNumber : Integer) : TBanqueSh ;
        Function  DeleteBD          () : Boolean;
        Function  GetBanque         () : String ;
        Function  GetCleRib         () : String ;
        Function  GetCode           () : String ;
        Function  GetCompte         () : String ;
        Function  GetGuichet        () : String ;
        Function  GetLibelle        () : String ;
        Function  GetCodeIBAN       () : String ;
        Function  GetGeneral        () : String ;  {FP 25/01/2010 FQ26918}
        Function  GetAuxiliaire     () : String ;  {FP 25/01/2010 FQ26918}
        Function  IsAllControler    (var theIndexDetail : Integer ; var theNumChampError : Integer) : Boolean ;
        Function  UpdateAndDeleteBD () : Boolean ;
        Function  UpdateBD          () : Boolean ;
     public

        Function  GetBanqueSh(theDate : TDateTime) : TBanqueSh   ; overload ;
        Function  GetBanqueSh(theCode, theCodeExercice : String) : TBanqueSh  ; overload ;
        Function  GetBanqueSh(theCode : String ; theDate : TDateTime) : TBanqueSh  ; overload ;

 end ;


{------------------------------------------------------------------------------}

function ChargerParametrageParDefaut() : TBanqueSh  ;

function GenererToutesSouchesBancairesPourExerciceSuivant() : Boolean  ;

function LancerInitialiserEnSerieDesSouchesBancaires( boPGIMajVer : Boolean = True) : Boolean  ; overload ;
function LancerInitialiserEnSerieDesSouchesBancaires(theBQCODE : String) : Boolean  ; overload ;

function InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; theTBanqueSH : TBanqueSh) : Boolean  ; overload ;
function InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; boPGIMajVer : Boolean ; theExerciceList : TStringList) : Boolean  ; overload ;

function IsExisteSoucheBancaire(theCode : String ; theExerciceCd :String = '') : Boolean  ;

function ModifierEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; theTBanqueSH : TBanqueSh) : Boolean  ;

function SupprimerEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; theExerciceCd : String) : Boolean  ;

Implementation

uses DB,
    Ent1 ,
    UtilCBP,
    HStatus,
    paramsoc,
    CPOBJITRTPARAM,
    CPProcBanqueAuxiliaire ;

{---------------------------------------------------------------------------------------
 Controle s'il existe une souche bancaire pour la domiciliation
}
function IsExisteSoucheBancaire(theCode : String ; theExerciceCd :String = '') : Boolean  ;
var
  SQL : String ;
BEGIN

  SQL := 'SELECT BSH_CODE FROM BANQUESH WHERE '
      +  ' BSH_CODE = "' + theCode + '"'  ;
  if (theExerciceCd <> '')    then SQL := SQL + ' AND BSH_EXERCICE = "' + theExerciceCd + '"' ;

  if ExisteSQL(SQL) then result := True else result := False ;
END ;

{---------------------------------------------------------------------------------------
* Rechercher les exercices ouverts.
}
function RechercherExercicesOuverts() : TStringList  ;
var
  TOBEXERCICE     : TOB ;
  TOBEXERCICEList : TOB ;
  QRY             : TQuery ;
  SQL             : String ;
  stExercice      : string ;
  index           : Integer ;
begin
  result := TStringList.Create() ;
  TOBEXERCICEList := Nil ;

  try
    SQL := 'SELECT EX_EXERCICE FROM EXERCICE '   //*FQ28058 on supprime le controle des exercices ouvers ou clos.
        +  ' ORDER BY EX_EXERCICE ' ;
    QRY   := OpenSQL(SQL, True) ;
    If not QRY.EOF then
    begin
      TOBEXERCICEList := TOB.Create('', Nil, -1);
      TOBEXERCICEList.LoadDetailDB('', '', '', QRY, True) ;

     for index := 0 to TOBEXERCICEList.Detail.Count - 1 do
      begin
        TOBEXERCICE := TOBEXERCICEList.Detail[index] ;
        if Assigned(TOBEXERCICE) then
        begin
          stExercice := TOBEXERCICE.GetValue('EX_EXERCICE') ;
          if (Result.IndexOf(stExercice) = -1 )  then Result.Add(stExercice) ;
        end;
      end ;
    end ;
    Ferme(QRY) ;

  finally
    if Assigned(TOBEXERCICEList) then FreeAndNil(TOBEXERCICEList) ;
  end ;
end;
{---------------------------------------------------------------------------------------
* Constitue une chaine SQL des exercices
}
function ConstituerListeExercice(theExerciceList : TStringList) : String  ;
var
  index       : Integer ;
  StExercice  : string ;
begin
   result := '' ; 
   if (theExerciceList.Count <> 0) then
    begin
      for index := 0 to theExerciceList.Count - 1 do
      begin
        StExercice := theExerciceList[index] ;
        if (StExercice <> '') then
        begin
          result := result + ' (BSH_EXERCICE = "' +  stExercice + '") ' ;
          if index < theExerciceList.Count - 1  then  result := result + ' OR ' ;
        end;
      end ;
    end ;
end;
{---------------------------------------------------------------------------------------
* Création des la souche bancaire en fonction d'un paramétrage saisi
*
* Les souches bancaires ne sont créés que si celles-ci ne sont pas trouvées
}
function LancerInitialiserEnSerieDesSouchesBancaires(boPGIMajVer : Boolean = True) : Boolean  ;
var
  SQL             : String ;
  SQLs            : String ;
  TOBBANQUECPList : TOB ;
  QRY             : TQuery ;
  exerciceList    : TStringList ;
  SQLExercice     : String ;
begin
  TOBBANQUECPList := Nil ;
  exerciceList    := Nil ;

  try

    //* Modification à cause du PGIMajVer : Recherche des dexercices ouverts à partir
    //* de la table des exercices car la VH n'est pas utilisable dans le pgimajver.
    //* Dans le traitement en cours, on prend les valeurs dans la VH

    SQLExercice := '' ;
    if boPGIMajVer then
    begin
      exerciceList := RechercherExercicesOuverts() ;
      SQLExercice  := ConstituerListeExercice(exerciceList) ;
    end
    else
    begin
      if (VH^.EnCours.Code <> '')
        or ( VH^.Suivant.Code <> '') then
      begin
        if (VH^.EnCours.Code <>'') then
          SQLExercice := SQLExercice + ' (BSH_EXERCICE = "' + VH^.EnCours.Code + '") ';

        if (VH^.EnCours.Code <> '')
            and ( VH^.Suivant.Code <> '') then
            SQLExercice := SQLExercice + ' OR ' ;

        if (VH^.Suivant.Code <>'') then
            SQLExercice := SQLExercice + ' (BSH_EXERCICE = "' + VH^.Suivant.Code + '") ';
      end;
    end ;

    if (SQLExercice <> '') then
    begin
      SQLs  := 'SELECT BSH_CODE FROM BANQUESH  WHERE ' + SQLExercice ;

      SQL   := 'SELECT * FROM BANQUECP '
            + ' LEFT JOIN GENERAUX ON BQ_GENERAL = G_GENERAL '
            + ' WHERE  (G_NATUREGENE = "BQE" ) AND NOT BQ_CODE IN (' + SQLs + ')' ;

      QRY   := OpenSQL(SQL, True) ;
      If not QRY.EOF then
      begin
        TOBBANQUECPList := TOB.Create('', Nil, -1);
        TOBBANQUECPList.LoadDetailDB('', '', '', QRY, True) ;
        InitialiserEnSerieDesSouchesBancaires(TOBBANQUECPList, boPGIMajVer, exerciceList) ;
      end ;
      Ferme(QRY) ;
    end ;
    result := True ;
    
  finally
    if Assigned(TOBBANQUECPList) then FreeAndNil(TOBBANQUECPList) ;
    if Assigned(exerciceList)    then FreeAndNil(exerciceList) ;
  end ;
end ;     
{
--------------------------------------------------------------------------------
}
function LancerInitialiserEnSerieDesSouchesBancaires(theBQCODE : String) : Boolean  ;
var
  SQL             : String ;
  SQLs            : String ;
  TOBBANQUECPList : TOB ;
  QRY             : TQuery ;
  SQLExercice     : String ;
begin
  TOBBANQUECPList := Nil ;
  try

    if (VH^.EnCours.Code <>'') then
        SQLExercice := SQLExercice + ' (BSH_EXERCICE = "' + VH^.EnCours.Code + '") ';

    if (VH^.EnCours.Code <> '')
        and ( VH^.Suivant.Code <> '') then
        SQLExercice := SQLExercice + ' OR ' ;

    if (VH^.Suivant.Code <>'') then
        SQLExercice := SQLExercice + ' (BSH_EXERCICE = "' + VH^.Suivant.Code + '") ';

    if (SQLExercice <> '') then
    begin
      SQLs  := 'SELECT BSH_CODE FROM BANQUESH WHERE '
            + ' BSH_CODE = "' + theBQCODE + '" '
            + ' AND '
           + SQLExercice  ;

      SQL   := 'SELECT * FROM BANQUECP '
            + ' LEFT JOIN GENERAUX ON BQ_GENERAL = G_GENERAL '
            + ' WHERE  (G_NATUREGENE = "BQE" ) AND NOT BQ_CODE IN (' + SQLs + ')' ;

      QRY   := OpenSQL(SQL, True) ;
      If not QRY.EOF then
      begin
        TOBBANQUECPList := TOB.Create('', Nil, -1);
        TOBBANQUECPList.LoadDetailDB('', '', '', QRY, True) ;
        InitialiserEnSerieDesSouchesBancaires(TOBBANQUECPList, False, Nil) ;
      end ;
      Ferme(QRY) ;
    end ;
    result := True ;

  finally
    if Assigned(TOBBANQUECPList) then FreeAndNil(TOBBANQUECPList) ;
  end ;
end ;
{
--------------------------------------------------------------------------------
}
function InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; boPGIMajVer  : Boolean ; theExerciceList : TStringList) : Boolean  ;
var
  BanqueSH        : TBanqueSh ;
  ENTETE          : String ;
  MSG             : String ;
  index           : Integer ;
  stExercice      : String ;
begin

  //* Constituer le paramétrage par défaut

  BanqueSH := ChargerParametrageParDefaut() ;

  if Not boPGIMajVer then   // Façon "Comptabilité"
  begin
    BanqueSh.SetCodeExercice(VH^.EnCours.Code);

    //* Exercice en cours

    if Not InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList, BanqueSh) then
    begin
      MSG := TraduireMemoire('Le traitement de création des souches bancaires sur l''exercice en cours s''est terminé en erreur.')
            + #13#13
            + TraduireMemoire('Le pointage avancé sera quand même activé.')
            + #13#13
            + TraduireMemoire('Mais, vous devez aller dans le menu "Structure/Paramètres - Banque - souches bancaires" pour contrôler la présence des souches sur tous les comptes bancaires.') ;
      ENTETE :=  TraduireMemoire('Attention')  ;
      PGIError(MSG, ENTETE);
    end ;

    //* Exercice suivant

    BanqueSh.SetCodeExercice(VH^.Suivant.Code);
    if Not InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList, BanqueSh) then
    begin
      MSG := TraduireMemoire('Le traitement de création des souches bancaires sur l''exercice suivant s''est terminé en erreur.')
          + #13#13
          + TraduireMemoire('Le pointage avancé sera quand même activé.')
          + #13#13
          + TraduireMemoire('Mais, vous devez aller dans le menu "Structure/Paramètres - Banque - souches bancaires" pour contrôler la présence des souches sur tous les comptes bancaires.') ;
      ENTETE :=  TraduireMemoire('Attention')  ;
      PGIError(MSG, ENTETE);
    end ;
  end
  else  //* Façon "pgiMajver"
  begin

    if Assigned(theExerciceList) then
    begin
      for index := 0 to theExerciceList.Count - 1 do
      begin
        stExercice := theExerciceList[index] ;
        if (stExercice <> '') then
        begin
          BanqueSh.SetCodeExercice(stExercice);
          InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList, BanqueSh) ;
        end;
      end;
    end;

  end;

  if Assigned(BanqueSH) then FreeAndNil(BanqueSH) ;

  result := True ;
end ;
{
--------------------------------------------------------------------------------
}
function InitialiserEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; theTBanqueSH : TBanqueSh) : Boolean  ;
var
  TOBBANQUECP     : TOB ;
  indexCp         : Integer ;
  BQ_CODE         : String ;
  BQ_CODEIBAN     : String ;
  BQ_LIBELLE      : String ;
  EX_EXERCICE     : String ;
  wBanqueSh       : TBanqueSh ;
begin
  result    := False ;
  wBanqueSh := Nil ;
  try

    INITMOVE(theTOBBANQUECPList.Detail.Count, '');
    for indexCp := 0 to theTOBBANQUECPList.Detail.Count - 1 do
    begin
      TOBBANQUECP := theTOBBANQUECPList.Detail[indexCp] ;
      if Assigned(TOBBANQUECP) then
      begin
        MOVECUR(False);
        BQ_CODE     := TOBBANQUECP.GetValue('BQ_CODE') ;
        BQ_CODEIBAN := TOBBANQUECP.GetValue('BQ_CODEIBAN') ;
        BQ_LIBELLE  := TOBBANQUECP.GetValue('BQ_LIBELLE') ;

        EX_EXERCICE := theTBanqueSH.GetCodeExercice() ; ;
        if (EX_EXERCICE <> '') then
        begin
          if Not IsExisteSoucheBancaire(BQ_CODE, EX_EXERCICE) then
          begin
            wBanqueSh := TBanqueSh.Create() ;
            wBanqueSh.SetAction(tBqCREAT);
            wBanqueSh.SetCode(BQ_CODE);
            wBanqueSh.SetDomiciliation(BQ_CODEIBAN);
            wBanqueSh.SetCodeExercice(EX_EXERCICE);
            wBanqueSh.SetContanteBancaire(theTBanqueSH.GetContanteBancaire());
            wBanqueSh.SetPositionnementAnnee(theTBanqueSH.GetPositionnementAnnee());
            wBanqueSh.SetTaille(theTBanqueSH.GetTaille());
            wBanqueSh.SetPremierNumero(theTBanqueSH.GetPremierNumero());
            wBanqueSh.SetDernierNumero(theTBanqueSH.GetDernierNumero());
            If Not wBanqueSh.UpdateBD() then Exit ;
            FreeAndNil(wBanqueSh);
          end
        end ;
      end ;
    end ;

    result := True ;
    FINIMOVE() ;
  finally
    if Assigned(wBanqueSh) then FreeAndNil(wBanqueSh) ;
  end ;
end ;
{---------------------------------------------------------------------------------------
* Suppression des la souche bancaire en fonction d'un paramétrage saisi
*
* Les souches bancaires sont supprimées uniquement si elles sont supprimées.
}

function SupprimerEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; theExerciceCd : String) : Boolean  ;
var
  TOBBANQUECP     : TOB ;
  indexCp         : Integer ;
  BQ_CODE         : String ;
  wBanqueSh       : TBanqueSh ;
begin
  result    := False ;
  wBanqueSh := Nil ;
  try

    INITMOVE(theTOBBANQUECPList.Detail.Count, '');
    for indexCp := 0 to theTOBBANQUECPList.Detail.Count - 1 do
    begin
      TOBBANQUECP := theTOBBANQUECPList.Detail[indexCp] ;
      if Assigned(TOBBANQUECP) then
      begin
        MOVECUR(False);
        BQ_CODE     := TOBBANQUECP.GetValue('BQ_CODE') ;
        if (theExerciceCd <> '') then
        begin
          if IsExisteSoucheBancaire(BQ_CODE, theExerciceCd) then
          begin
            wBanqueSh := TBanqueSh.Create() ;
            wBanqueSh.LoadData(BQ_CODE, theExerciceCd) ;
            wBanqueSh.SetAction(tBqDELETE);
            If Not wBanqueSh.DeleteBD() then Exit ;
            FreeAndNil(wBanqueSh);
          end
        end ;
      end ;
    end ;

    result := True ;
    FINIMOVE() ;
  finally
    if Assigned(wBanqueSh) then FreeAndNil(wBanqueSh) ;
  end ;
end ;


{---------------------------------------------------------------------------------------
* Modification des la souche bancaire en fonction d'un paramétrage saisi
*
* Les souches bancaires ne sont modifiées que si celles-ci sont trouvées
}

function ModifierEnSerieDesSouchesBancaires(theTOBBANQUECPList : TOB ; theTBanqueSH : TBanqueSh) : Boolean  ;
var
  TOBBANQUECP     : TOB ;
  indexCp         : Integer ;
  BQ_CODE         : String ;
  BQ_CODEIBAN     : String ;
  BQ_LIBELLE      : String ;
  EX_EXERCICE     : String ;
  wBanqueSh       : TBanqueSh ;
begin
  result := False ;
  wBanqueSh   := Nil ;
  try

    INITMOVE(theTOBBANQUECPList.Detail.Count, '');
    for indexCp := 0 to theTOBBANQUECPList.Detail.Count - 1 do
    begin
      TOBBANQUECP := theTOBBANQUECPList.Detail[indexCp] ;
      if Assigned(TOBBANQUECP) then
      begin
        MOVECUR(False);
        BQ_CODE     := TOBBANQUECP.GetValue('BQ_CODE') ;
        BQ_CODEIBAN := TOBBANQUECP.GetValue('BQ_CODEIBAN') ;
        BQ_LIBELLE  := TOBBANQUECP.GetValue('BQ_LIBELLE') ;

        EX_EXERCICE := theTBanqueSH.GetCodeExercice() ; ;
        if (EX_EXERCICE <> '') then
        begin
          if IsExisteSoucheBancaire(BQ_CODE, EX_EXERCICE) then
          begin
            wBanqueSh := TBanqueSh.Create() ;
            wBanqueSh.SetAction(tBqMODIF);
            wBanqueSh.SetEcraseLastNum(True);
            wBanqueSh.SetCode(BQ_CODE);
            wBanqueSh.SetDomiciliation(BQ_CODEIBAN);
            wBanqueSh.SetCodeExercice(EX_EXERCICE);
            wBanqueSh.SetContanteBancaire(theTBanqueSH.GetContanteBancaire());
            wBanqueSh.SetPositionnementAnnee(theTBanqueSH.GetPositionnementAnnee());
            wBanqueSh.SetTaille(theTBanqueSH.GetTaille());
            wBanqueSh.SetPremierNumero(theTBanqueSH.GetPremierNumero());
            wBanqueSh.SetDernierNumero(theTBanqueSH.GetDernierNumero());
            If Not wBanqueSh.UpdateBD() then Exit ;
            FreeAndNil(wBanqueSh);
          end
        end ;
      end ;
    end ;

    result := True ;
    FINIMOVE() ;
  finally
    if Assigned(wBanqueSh) then FreeAndNil(wBanqueSh) ;
  end ;
end ;



{---------------------------------------------------------------------------------------
* Retourne le paramétrage par défaut d'une souche bancaire
*
}
function ChargerParametrageParDefaut() : TBanqueSh  ;
var
  SQLITRTPARAM    : String ;
  QRYTRTPARAM     : TQuery ;
  iTRTPARAM       : TITRTParam ;
  TOBITRTPARAM    : TOB ;
  TOBDATA         : TOB ;
  constante       : String ;
  positionDate    : Integer ;
  Longueur        : Integer ;
  taille          : Integer ;
  dernierNumero   : INteger ;
begin
  result := TBanqueSh.Create() ;
  result.SetAction(tBqCREAT);

  SQLITRTPARAM := 'SELECT * FROM ITRTPARAM WHERE ITT_TYPPAR = "CPB" AND ITT_DEFAUT ="X"' ;
  QRYTRTPARAM := OpenSQL(SQLITRTPARAM, True) ;

  if Not QRYTRTPARAM.EOF then
  begin
    TOBITRTPARAM := TOB.Create('ITRTPARAM', nil, -1) ;
    TOBITRTPARAM.LoadDetailDB('ITRTPARAM', '', '', QRYTRTPARAM, True) ;

    iTRTPARAM := TITRTParam.Create() ;
    iTRTPARAM.LoadData(TOBITRTPARAM.Detail[TOBITRTPARAM.Detail.Count -1]) ;
    TOBDATA := iTRTPARAM.GetData() ;

    result.LoadData(TOBDATA) ;
    result.SetCodeExercice(VH.EnCours.code); //* Pour la beauté du geste....
    result.SetCurrentNumero(result.GetPremierNumero());

    If Assigned(TOBITRTPARAM) then FreeAndNil(TOBITRTPARAM) ;
    If Assigned(iTRTPARAM) then FreeAndNil(iTRTPARAM) ;
  end
  else
  begin
    Longueur      :=   GetParamSocSecur('SO_LGCPTEGEN', 6)  ;

    if (Longueur > 8) then Longueur := 8 ;
    constante := '$C' + IntToStr(longueur) + '$';

    positionDate  := 6 ;

    taille := 17 - (Longueur + positionDate + 2) ;

    if (taille <> 0) then dernierNumero := StrToInt(StringOfChar('9', taille))
                     else dernierNumero := 0;

    result.SetContanteBancaire(constante);
    result.SetCodeExercice(VH.EnCours.code);
    result.SetPositionnementAnnee(positionDate);
    result.SetTaille(taille);
    result.SetPremierNumero(0);
    result.SetDernierNumero(dernierNumero);
    result.SetCurrentNumero(result.GetPremierNumero());
  end ;
  Ferme(QRYTRTPARAM) ;

end ;

{---------------------------------------------------------------------------------------
* Génération des souches manquantes sur un exercice avec la possibilité de reprendre le
* paramétrage de l'année précédente
*
}
function GenererToutesSouchesBancairesPourExerciceSuivant() : Boolean  ;
var
  QRYBANQUECP     : TQuery ;
  TOBBANQUECPList : TOB ;
  TOBBANQUECP     : TOB ;
  indexCp         : Integer ;
  BQ_CODE         : String ;
  BQ_CODEIBAN     : String ;
  BQ_LIBELLE      : String ;

  SQL             : String ;

  EX_EXERCICE     : String ;
  EX_EXERCICEPRC  : String ;

  wBanqueShPrc    : TBanqueSh ;
  wBanqueSh       : TBanqueSh ;
  defBanqueSh     : TBanqueSh ;

BEGIN
  result        := False ;
  wBanqueSh     := Nil ;
  wBanqueShPrc  := Nil ;

 //* Exercice 

  EX_EXERCICE    := VH^.Suivant.Code ;
  EX_EXERCICEPRC := VH^.EnCours.Code ;

  //* Récupération du paramétrage par défaut -> Sinon

  defBanqueSh := ChargerParametrageParDefaut() ;

  //* Boucle sur les comptes bancaires avec création des souches manquantes

  TOBBANQUECPList := TOB.Create('', Nil, -1) ;
  try
    SQL := 'SELECT * FROM BANQUECP ORDER BY BQ_CODE ' ;
    QRYBANQUECP := OpenSQL(SQL, True) ;
    if Not QRYBANQUECP.EOF then
    begin
      TOBBANQUECPList.LoadDetailDB('BANQUECP', '', '', QRYBANQUECP, True) ;

      INITMOVE(TOBBANQUECPList.Detail.Count, '');
      for indexCp := 0 to TOBBANQUECPList.Detail.Count - 1 do
      begin
        TOBBANQUECP := TOBBANQUECPList.Detail[indexCp] ;
        if (TOBBANQUECP <> Nil) then
        begin
          MOVECUR(False);
          BQ_CODE     := TOBBANQUECP.GetValue('BQ_CODE') ;
          BQ_CODEIBAN := TOBBANQUECP.GetValue('BQ_CODEIBAN') ;
          BQ_LIBELLE  := TOBBANQUECP.GetValue('BQ_LIBELLE') ;

          if Not IsExisteSoucheBancaire(BQ_CODE, EX_EXERCICE) then
          begin
            if (EX_EXERCICEPRC = '')
              or Not IsExisteSoucheBancaire(BQ_CODE, EX_EXERCICEPRC) then
            begin                                         //* Souche précédente n'existe pas ou premier exercice
              wBanqueSh := TBanqueSh.Create() ;
              wBanqueSh.SetEcraseLastNum(True);
              wBanqueSh.SetAction(tBqCREAT);
              wBanqueSh.SetCode(BQ_CODE);
              wBanqueSh.SetDomiciliation(BQ_CODEIBAN);
              wBanqueSh.SetCodeExercice(EX_EXERCICE);
              wBanqueSh.SetContanteBancaire(defBanqueSh.GetContanteBancaire()) ;
              wBanqueSh.SetPositionnementAnnee(defBanqueSh.GetPositionnementAnnee());          
              wBanqueSh.SetTaille(defBanqueSh.GetTaille());
              wBanqueSh.SetPremierNumero(defBanqueSh.GetPremierNumero());
              wBanqueSh.SetDernierNumero(defBanqueSh.GetDernierNumero());

              If Not wBanqueSh.UpdateBD() then
              begin
                Ferme(QRYBANQUECP) ;
                Exit ;
             end ;

              if assigned(wBanqueSh) then FreeAndNil(wBanqueSh);
            end
            else
            begin                                       //* Souche précédent existe
              wBanqueShPrc := TBanqueSh.Create() ;
              wBanqueShPrc.LoadData(BQ_CODE, EX_EXERCICEPRC) ;

              wBanqueSh := TBanqueSh.Create() ;
              wBanqueSh.SetEcraseLastNum(True);
              wBanqueSh.SetAction(tBqCREAT);
              wBanqueSh.SetCode(BQ_CODE);
              wBanqueSh.SetDomiciliation(BQ_CODEIBAN);
              wBanqueSh.SetCodeExercice(EX_EXERCICE);
              wBanqueSh.SetContanteBancaire(wBanqueShPrc.GetContanteBancaire());
              wBanqueSh.SetPositionnementAnnee(wBanqueShPrc.GetPositionnementAnnee());
              wBanqueSh.SetTaille(wBanqueShPrc.GetTaille());
              wBanqueSh.SetPremierNumero(wBanqueShPrc.GetPremierNumero());
              wBanqueSh.SetDernierNumero(wBanqueShPrc.GetDernierNumero());

              If Not wBanqueSh.UpdateBD() then
              begin
                Ferme(QRYBANQUECP) ;
                Exit ;
              end ;

              if assigned(wBanqueSh) then FreeAndNil(wBanqueSh);
              if assigned(wBanqueShPrc) then FreeAndNil(wBanqueShPrc);
            end ;
          end ;
        end ;
      end ;
    end ;
    result := True ;
    FINIMOVE() ;
    Ferme(QRYBANQUECP) ;

  finally
    if assigned(TOBBANQUECPList)  then FreeAndNil(TOBBANQUECPList);
    if assigned(defBanqueSh)      then FreeAndNil(defBanqueSh);
    if assigned(wBanqueSh)        then FreeAndNil(wBanqueSh);
   if assigned(wBanqueShPrc)      then FreeAndNil(wBanqueShPrc);
  end ;
END ;

{
*******************************************************************************
* objeT   TBanqueShList
*******************************************************************************
}
{------------------------------------------------------------------------------}
Constructor TBanqueShList.Create() ;
BEGIN

  inherited Create;

  bBqeAuxFg   := False ;
  Liste       :=  TList.Create() ;
  TOBBANQUECP := TOB.Create('BANQUECP', Nil, -1);
END ;

{------------------------------------------------------------------------------}
Destructor TBanqueShList.Destroy() ;
var
  index     : Integer ;
  banqueSh  : TBanqueSh ;
BEGIN
if Liste.Count <> 0 then
  begin
  for index := Liste.Count -1 downto 0 do
    begin
    banqueSh := TBanqueSh(Liste[index]) ;
    if (banqueSh <> Nil) then  FreeAndNil(banqueSh) ;
    end ;
  end ;

 if (Liste <> Nil) then FreeAndNil(Liste) ;
 if (TOBBANQUECP <> Nil) then FreeAndNil(TOBBANQUECP) ;
 Inherited Destroy ;
END ;

{------------------------------------------------------------------------------}

Function TBanqueShList.LoadDataByCompte(theCptGENERAL, theCptAUXILIAIRE : String) : Boolean  ;
var
  MSG      : String ;
  ENTETE   : String ;
  SQL      : String ;
  QRY      : TQuery ;
  WTOBLIST : TOB ;
  wTOB     : TOB ;
  index    : Integer ;
  banqueSh : TBanqueSh ;
  BQ_CODE     : String ;
BEGIN
  result := False ;

  if (theCptGENERAL = '') and (theCptAUXILIAIRE = '') then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Vous devez renseigner le compte général voire le compte auxiliaire si le compte général est collectif.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;                           

  bBqeAuxFg := IsCompteBancaireAuxiliarise(theCptGENERAL) ;

  if bBqeAuxFg  and (theCptAUXILIAIRE = '') then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Vous devez renseigner le compte auxiliaire si le compte général est collectif.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;


  if bBqeAuxFg and Not IsTiersBanquePourGeneral(theCptGENERAL, theCptAUXILIAIRE) then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Le compte auxiliaire (') + theCptAUXILIAIRE
           +  TraduireMemoire(') n''est pas associé au compte général (') + theCptGENERAL  + ').'  ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  SQL := 'SELECT * FROM BANQUECP WHERE BQ_GENERAL = "' + theCptGENERAL + '"'
      + ' AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"' ;
  if bBqeAuxFg then SQL := SQL +  ' AND BQ_AUXILIAIRE = "' + theCptAUXILIAIRE + '"' ;

  QRY := OpenSQL(SQL, True) ;
  if QRY.EOF then
  begin
    Ferme(QRY) ;

    ENTETE  := TraduireMemoire('Attention') ;
    MSG     := TraduireMemoire('Aucune souche bancaire n''a été trouvée pour le compte ') + theCptGENERAL ;
    if bBqeAuxFg then  MSG := MSG  + ' - ' + theCptAUXILIAIRE ;
    MSG := MSG + ').'  ;
    PGIError(MSG, ENTETE);

    Exit ;
  end ;

  if (WTOBLIST <> Nil) then FreeAndNil(WTOBLIST);
  WTOBLIST := TOB.Create('', Nil, -1) ;
  WTOBLIST.LoadDetailDB('', '', '', QRY, False) ;
  TOBBANQUECP.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, False);
  FreeAndNil(WTOBLIST);
  Ferme(QRY) ;

  BQ_CODE     := TOBBANQUECP.GetValue('BQ_CODE') ;

  SQL := 'SELECT * FROM BANQUESH ' ;
  SQL := SQL + ' WHERE BSH_CODE = "' + BQ_CODE + '" ' ;

  QRY := OpenSQL(SQL, True) ;
  if QRY.EOF then
  begin
    Ferme(QRY) ;

    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Pas de souche bancaire pour le compte (') + theCptGENERAL ;
    if bBqeAuxFg then  MSG := MSG  + ' - ' + theCptAUXILIAIRE ;
    MSG := MSG + ').'  ;
    PGIError(MSG, ENTETE);

    Exit ;
  end ;

  WTOBLIST := TOB.Create('', Nil, -1) ;
  WTOBLIST.LoadDetailDB('BANQUESH', '', '', QRY, False) ;
  Ferme(QRY) ;

  for index := 0 to WTOBLIST.Detail.Count -1 do
  begin
    wTOB := WTOBLIST.Detail[index] ;
    if (wTOB <> Nil) then
    begin
      banqueSh := TBanqueSh.Create() ;
      banqueSh.LoadData(wTOB) ;
      Liste.Add(banqueSh) ;
    end ;
  end ;
  FreeAndNil(WTOBLIST);

  result := True ;
END ;
{------------------------------------------------------------------------------}
Function TBanqueShList.LoadData(theCode : String;
                                theBANQUECP : TOB = Nil) : Boolean  ;
var
  MSG        : String ;
  ENTETE     : String ;
  SQL        : String ;
  QRY        : TQuery ;
  WTOBLIST   : TOB ;
  wTOB       : TOB ;
  index      : Integer ;
  banqueSh   : TBanqueSh ;
  BQ_GENERAL : String ;
BEGIN
  result := False ;
  //*  Recherche du paramétrage de BANQUECP.

  if (theBANQUECP <> Nil) then
  begin
    if (theBANQUECP.Detail.Count <> 0) then
    begin
      MSG      := TraduireMemoire('Attention') ;
      ENTETE   := TraduireMemoire('Impossible de charger les informations du compte bancaire (BANQUECP).') ;
      PGIError(MSG, ENTETE);
      Exit ;
    end ;
    TOBBANQUECP.Dupliquer(theBANQUECP, True, True, False);
  end
  else
  begin
    SQL := 'SELECT * FROM BANQUECP WHERE BQ_CODE = "' + theCode + '"'  ;
    QRY := OpenSQL(SQL, True) ;
    if Not QRY.EOF then
    begin
      WTOBLIST := TOB.Create('', Nil, -1) ;
      WTOBLIST.LoadDetailDB('', '', '', QRY, False) ;
      TOBBANQUECP.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, False);
      FreeAndNil(WTOBLIST);
    end ;
    Ferme(QRY) ;
  end ;

  BQ_GENERAL := TOBBANQUECP.GetValue('BQ_GENERAL') ;
  bBqeAuxFg := IsCompteBancaireAuxiliarise(BQ_GENERAL) ;

  SQL := 'SELECT * FROM BANQUESH ' ;
  SQL := SQL + 'LEFT JOIN EXERCICE ON (EX_EXERCICE = BSH_EXERCICE) ' ;     //* FQ 28058 - On supprime le controle sur l'exercice ouvert ou clos
  SQL := SQL + ' WHERE BSH_CODE = "' + theCode + '" ' ;

  QRY := OpenSQL(SQL, True) ;
  if Not QRY.EOF then
  begin
    WTOBLIST := TOB.Create('', Nil, -1) ;
    WTOBLIST.LoadDetailDB('BANQUESH', '', '', QRY, False) ;
    for index := 0 to WTOBLIST.Detail.Count -1 do
    begin
      wTOB := WTOBLIST.Detail[index] ;
      if (wTOB <> Nil) then
      begin
        banqueSh := TBanqueSh.Create() ;
        banqueSh.LoadData(wTOB) ;
        Liste.Add(banqueSh) ;
      end ;
    end ;
    FreeAndNil(WTOBLIST);
  end ;
  result := True ;
  Ferme(QRY) ;
END ;

{------------------------------------------------------------------------------
* Ajoute dans la liste une souche bancaire
}
Function TBanqueShList.Add(theBanqueSh : TBanqueSh) : Boolean ;
var
  BSH_CODE          : String ;
  BSH_CODEEXERCICE  : String ;
  ENTETE            : String ;
  MSG               : String ;
  wwBanqueSh        : TBanqueSh ;
  ttCount           : Integer ;
BEGIN
  result := False ;
  ENTETE  := TraduireMemoire('Attention') ;

  BSH_CODE          := theBanqueSh.GetCode() ;
  BSH_CODEEXERCICE  := theBanqueSh.GetCodeExercice() ;

  if (BSH_CODE = '') then
  begin
    MSG := TraduireMemoire('Le code bancaire (BSH_CODE) doit être renseigné.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  if (BSH_CODEEXERCICE = '') then
  begin
    MSG := TraduireMemoire('Le code exercice (BSH_CODEEXERCICE) doit être renseigné.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  wwBanqueSh := GetBanqueSh(BSH_CODE, BSH_CODEEXERCICE) ;
  if (wwBanqueSh <> Nil) then
  begin
    if (wwBanqueSh.GetAction() <> tBqDELETE) then
    begin
      MSG := TraduireMemoire('Impossible d''ajouter cette souche. Une souche est déjà présente dans la liste des souches bancaires.') ;
      PGIError(MSG, ENTETE);
      Exit ;
    end ;
  end ;

  //* La première fois, le Liste.Count est à 0. Donc le premier objet sera sur l'occurence 0
  //* Les fois suivantes, le count est 1, 2, 3, 4, Etc.
  //* Donc si on rajoute l'occurence de l'objet sera égal au count avant l'ajout

  ttCount := Liste.Count ;
  liste.Insert(ttCount, theBanqueSh) ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.Detail(theOccurenceNumber : Integer) : TBanqueSh  ;
BEGIN
  result := Nil ;
  if (theOccurenceNumber >  Liste.Count - 1) then Exit ;

  result := TBanqueSh(Liste.Items[theOccurenceNumber]) ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.Count() : Integer  ;
BEGIN
  result := Liste.Count ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueShList.GetBanque() : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_ETABBQ');
END ;

{------------------------------------------------------------------------------}
Function  TBanqueShList.GetCleRib()  : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_CLERIB');
END ;

{------------------------------------------------------------------------------}
Function  TBanqueShList.GetCode()  : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_CODE');
END ;

{------------------------------------------------------------------------------}
Function  TBanqueShList.GetCompte()  : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_NUMEROCOMPTE');
END ;

{------------------------------------------------------------------------------}
Function  TBanqueShList.GetGuichet()  : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_GUICHET');
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.GetLibelle()  : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_LIBELLE');
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.GetCodeIBAN() : String ;
BEGIN
  result := TOBBANQUECP.GetValue('BQ_CODEIBAN');
END ;

function TBanqueShList.GetAuxiliaire: String;
begin
  result := TOBBANQUECP.GetValue('BQ_AUXILIAIRE');  {FP 25/01/2010 FQ26918}
end;

function TBanqueShList.GetGeneral: String;
begin
  result := TOBBANQUECP.GetValue('BQ_GENERAL');  {FP 25/01/2010 FQ26918}
end;


{------------------------------------------------------------------------------}
Function TBanqueShList.GetBanqueSh(theCode, theCodeExercice : String) : TBanqueSh  ;
var
  index             : Integer  ;
  banqueSh          : TBanqueSh ;
  BSH_CODE          : String ;
  BSH_DATEXE        : String ;
BEGIN
  result := Nil ;
  for index := 0 to liste.Count - 1 do
  begin
    banqueSh := Detail(index) ;

    BSH_CODE          := banqueSh.Getcode() ;
    BSH_DATEXE        := banqueSh.GetCodeExercice() ;
    if (BSH_CODE  = theCode)
      and (BSH_DATEXE = theCodeExercice) then
    begin
      Result := banqueSh ;
      Exit ;
    end ;

  end ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.GetBanqueSh(theCode : String ; theDate : TDateTime) : TBanqueSh  ;
var
  index             : Integer  ;
  banqueSh          : TBanqueSh ;
  BSH_CODE          : String ;
  BSH_DATEXED       : TDateTime ;
  BSH_DATEXEF       : TDateTime ;
BEGIN
  result := Nil ;
  for index := 0 to liste.Count - 1 do
  begin
    banqueSh := Detail(index) ;

    BSH_CODE          := banqueSh.Getcode() ;
    BSH_DATEXED       := banqueSh.GetDateDebutExercice() ;
    BSH_DATEXEF       := banqueSh.GetDateFinExercice() ;
    if (BSH_CODE  = theCode)
      and ((theDate >= BSH_DATEXED) and (theDate <= BSH_DATEXEF)) then
    begin
      Result := banqueSh ;
      Exit ;
    end ;

  end ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.GetBanqueSh(theDate : TDateTime) : TBanqueSh  ;
var
  index             : Integer  ;
  banqueSh          : TBanqueSh ;
  BSH_DATEXED       : TDateTime ;
  BSH_DATEXEF       : TDateTime ;
BEGIN
  result := Nil ;
  for index := 0 to liste.Count - 1 do
  begin
    banqueSh := Detail(index) ;

    BSH_DATEXED       := banqueSh.GetDateDebutExercice() ;
    BSH_DATEXEF       := banqueSh.GetDateFinExercice() ;
    if ((theDate >= BSH_DATEXED) and (theDate <= BSH_DATEXEF)) then
    begin
      Result := banqueSh ;
      Exit ;
    end ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.UpdateBD() : Boolean ;
var
  index     : Integer ;
  banqueSh  : TBanqueSh ;
BEGIN
  result := False ;
  for index := 0 to Liste.Count - 1  do
  begin
    banqueSh := Detail(index) ;
    if (banqueSh <> Nil) then
    begin
      if Not banqueSh.UpdateBD() then Exit ;
    end ;
  end ;
  result := True ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.UpdateAndDeleteBD() : Boolean ;
var
  index     : Integer ;
  banqueSh  : TBanqueSh ;
BEGIN
  result := False ;
  for index := 0 to Liste.Count - 1  do
  begin
    banqueSh := Detail(index) ;
    if (banqueSh <> Nil) then
    begin
      case banqueSh.getAction() of
        tBqCREAT, tBqMODIF : if not  banqueSh.UpdateBD() then Exit ;
        tBqDELETE : if not banqueSh.DeleteBD() then Exit  ;
      end ;
    end ;
  end ;
  result := True ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueShList.DeleteBD(): Boolean;
var
  index     : Integer ;
  banqueSh  : TBanqueSh ;
BEGIN
  result := False ;
  for index := 0 to Liste.Count - 1  do
  begin
    banqueSh := Detail(index) ;
    if (banqueSh <> Nil) then
    begin
      if Not banqueSh.DeleteBD() then Exit ;
    end ;
  end ;
  result := True ;
END ;


{------------------------------------------------------------------------------}
Function TBanqueShList.IsAllControler(var theIndexDetail : Integer ; var theNumChampError : Integer) : Boolean ;
var
  index     : Integer ;
  banqueSh  : TBanqueSh ;
BEGIN
  result := False ;

  theIndexDetail := 0 ;
  for index := 0 to Liste.Count - 1  do
  begin
    banqueSh := Detail(index) ;
    if (banqueSh <> Nil) then
    begin
      if Not banqueSh.IsAllControler(theNumChampError) then
      begin
        theIndexDetail := Index ;
        Exit ;
      end ;
    end ;
  end ;
  result := True ;
END ;

{
*******************************************************************************
* objeT   TBanqueSh
*******************************************************************************
}

{------------------------------------------------------------------------------}
Constructor TBanqueSh.Create() ;
BEGIN

  inherited Create;
  compteGeneral     := '' ;
  compteAuxiliaire  := '' ;

  TOBBANQUESH  := TOB.Create('BANQUESH', Nil, -1) ;
  TOBEXERCICE  := TOB.Create('EXERCICE', Nil, -1) ;
  TOBDEFAIRE   := TOB.Create('BANQUESH', Nil, -1) ;
  TOBEXDEFAIRE := TOB.Create('EXERCICE', Nil, -1) ;
  TOBDATERELEVE := TOB.Create('TOBDATERELEVE', Nil, -1) ;

  TOBBANQUESH.InitValeurs();
  TOBDEFAIRE.InitValeurs();

  refPointage := '' ;

  EcraseLastNum := False ;
  
  action      := tBqRIEN ;
  ActionEcran := tEcRIEN ;

  ActionDefaire  := tBqRIEN  ;
  ActionEcranDefaire := tEcRIEN ;
  EcraseLastNumDefaire := False ;

END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.LoadData(theTOBBANQUESH : TOB) : Boolean  ;
var
  ENTETE  : String ;
  MSG     : String ;
  SQL     : String ;
  QRYEXERCICE : TQuery ;
  WTOBLIST  : TOB ;
BEGIN
  result := False ;
  ENTETE  := TraduireMemoire('Attention') ;

  if (theTOBBANQUESH.detail.Count <> 0) then
  begin
     MSG     := TraduireMemoire('Impossible de charger les données sur la souche bancaire. Il y a trop d''informations.') ;
     PGIError(MSG, ENTETE);
     Exit ;
  end ;

  TOBBANQUESH.ClearDetail() ;
  TOBBANQUESH.Dupliquer(theTOBBANQUESH, True, True, False);

  TOBDEFAIRE.ClearDetail() ;
  TOBDEFAIRE.Dupliquer(theTOBBANQUESH, True, True, False);
  SetAction(tBqMODIF);

  SQL := 'SELECT * FROM EXERCICE WHERE  EX_EXERCICE = "' + GetCodeExercice()  + '"'  ;
  QRYEXERCICE := OpenSQL(SQL, True) ;
  if Not QRYEXERCICE.EOF then
  begin
    WTOBLIST := TOB.Create('', Nil, -1) ;
    WTOBLIST.LoadDetailDB('EXERCICE', '', '', QRYEXERCICE, True) ;

    TOBEXERCICE.ClearDetail() ;
    TOBEXERCICE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);

    TOBEXDEFAIRE.ClearDetail() ;
    TOBEXDEFAIRE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);
    FreeAndNil(WTOBLIST) ;
  end ;
  Ferme(QRYEXERCICE) ;

 readCompteGeneralAuxiliaire();
 InitReferencePointage(Date()) ;
 result := True ;
END ;
{------------------------------------------------------------------------------}
Function TBanqueSh.LoadData(theCode, theCodeExercice : String) : Boolean  ;
var
  SQL       : String ;
  ENTETE    : String ;
  MSG       : String ;
  QRY       : TQuery ;
  QRYEXERCICE : TQuery ;
  WTOBLIST  : TOB ;
BEGIN
  result := False ;

  SQL := 'SELECT * FROM BANQUESH WHERE '
        + ' BSH_CODE              = "' + theCode          + '"'
        + ' AND BSH_EXERCICE = "'      + theCodeExercice  + '"' ;

  QRY := OpenSQL(SQL, True) ;
  if Not QRY.EOF then
  begin
    Action := tBqMODIF ;
    WTOBLIST := TOB.Create('', Nil, -1) ;

    WTOBLIST.LoadDetailDB('BANQUESH', '', '', QRY, True) ;

    TOBBANQUESH.ClearDetail() ;
    TOBBANQUESH.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);

    TOBDEFAIRE.ClearDetail() ;
    TOBDEFAIRE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);
    WTOBLIST.ClearDetail() ;

    SQL := 'SELECT * FROM EXERCICE WHERE  EX_EXERCICE = "' + GetCodeExercice()  + '"'  ;
    QRYEXERCICE := OpenSQL(SQL, True) ;
    WTOBLIST.LoadDetailDB('EXERCICE', '', '', QRYEXERCICE, True) ;

    TOBEXERCICE.ClearDetail() ;
    TOBEXERCICE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);

    TOBEXDEFAIRE.ClearDetail() ;
    TOBEXDEFAIRE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);

    Ferme(QRYEXERCICE) ;
    FreeAndNil(WTOBLIST) ;
    readCompteGeneralAuxiliaire();
    InitReferencePointage(Date()) ;
    result := True ;
  end
  else
  begin
    ENTETE  := TraduireMemoire('Attention') ;
    MSG     := TraduireMemoire('La souche bancaire n''existe pas pour le code ')
            +  theCode
            +  TraduireMemoire(' et le code exercice comptable "')
            +  theCodeExercice
            +  '".';
    PGIError(MSG, ENTETE);
  end ;
  Ferme(QRY) ;
END ;

{------------------------------------------------------------------------------}
procedure TBanqueSh.readCompteGeneralAuxiliaire() ;
var
  SQL : String ;
  QRY : TQuery ;
begin
  if (GetCode <> '') then
  begin
    SQL := 'SELECT BQ_GENERAL, BQ_AUXILIAIRE FROM BANQUECP WHERE '
        + ' BQ_CODE = "' +  GetCode() +'"'
        + ' AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '" ' ;
    QRY := OpenSQL(SQL, True) ;
    If Not QRY.EOF then
    begin
      compteGeneral     := QRY.Fields[0].AsString ;
      compteAuxiliaire  := QRY.Fields[1].AsString ;
    end ;
    Ferme(QRY) ;
  end
  else
  begin
    compteGeneral     := '' ;
    compteAuxiliaire  := '' ;
  end ;
end ;

{------------------------------------------------------------------------------}
procedure TBanqueSh.InitialiserCreation(theCode, theDomiciliation : String ; theCodeExercice : String = '') ;
var
  SQL : String ;
  QRY : TQuery ;
  WTOBLIST : TOB ;
BEGIN
  SetAction(tBqCREAT);
  SetCode(theCode);
  SetDomiciliation(theDomiciliation);
  SetDateExercice(iDate1900);
  SetCodeExercice('');

  if (theCodeExercice <> '') then
  begin
    SQL := 'SELECT * FROM EXERCICE WHERE EX_EXERCICE = "' + theCodeExercice + '" ' ;
    QRY := OpenSQL(SQL, True) ;
    if Not QRY.EOF then
    begin
      WTOBLIST := TOB.Create('', Nil, -1) ;
      TOBEXERCICE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, False);
      SetCodeExercice(theCodeExercice);
      SetDateExercice(TOBEXERCICE.GetValue('EX_DATEFIN'));
      FreeAndNil(WTOBLIST);
    end ;
    Ferme(QRY);
  End ;

  SetContanteBancaire('') ;
  SetPositionnementAnnee(0);
  SetTaille(8);
  SetPremierNumero(0);
  SetDernierNumero(99999999);
  SetCurrentNumero(0);

  //* On doit initialiser la TOBDEBFAIRE

  IF (TOBDEFAIRE <> Nil) then
  begin
    TOBDEFAIRE.PutValue('BSH_CODE',           theCode) ;
    TOBDEFAIRE.PutValue('BSH_DOMICILIATION',  theDomiciliation) ;
    TOBDEFAIRE.PutValue('BSH_DATEXE',         iDate1900) ;
    TOBDEFAIRE.PutValue('BSH_EXERCICE',       '') ;
    TOBDEFAIRE.PutValue('BSH_CONSTBQ',        '') ;
    TOBDEFAIRE.PutValue('BSH_POSANNEE',       0) ;
    TOBDEFAIRE.PutValue('BSH_LGNUMERO',       8) ;
    TOBDEFAIRE.PutValue('BSH_NUMERODEBUT',    0) ;
    TOBDEFAIRE.PutValue('BSH_NUMEROFIN',      99999999) ;
    TOBDEFAIRE.PutValue('BSH_LASTNUM',        0) ;

    ActionDefaire         := GetAction() ;
    ActionEcran           := GetActionEcran() ;
    EcraseLastNumDefaire  := EcraseLastNum ;
  end ;
END ;

{------------------------------------------------------------------------------}
procedure TBanqueSh.InitialiserCreation(theCode, theDomiciliation : String ; theDateExercice : TDateTime)  ;
var
  SQL : String ;
  QRY : TQuery ;
  WTOBLIST : TOB ;
BEGIN
  SetAction(tBqCREAT);
  SetCode(theCode);
  SetDomiciliation(theDomiciliation);
  SetDateExercice(iDate1900);
  SetCodeExercice('');

  if (theDateExercice <> iDate1900)
    and (theDateExercice <> iDate2099) then
  begin
    SQL := 'SELECT * FROM EXERCICE WHERE EX_DATEFIN = "' + usdatetime_(theDateExercice)+ '" ' ;
    QRY := OpenSQL(SQL, True) ;
    if Not QRY.EOF then
    begin
      WTOBLIST := TOB.Create('', Nil, -1) ;
      TOBEXERCICE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, False);
      SetDateExercice(theDateExercice);
      SetCodeExercice(TOBEXERCICE.GetValue('EX_EXERCICE'));
      FreeAndNil(WTOBLIST);
    end ;
    Ferme(QRY);
  End ;

  SetContanteBancaire('') ;
  SetPositionnementAnnee(0);
  SetTaille(8);
  SetPremierNumero(0);
  SetDernierNumero(99999999);
  SetCurrentNumero(0);

  //* On doit initialiser la TOBDEBFAIRE

  IF (TOBDEFAIRE <> Nil) then
  begin
    TOBDEFAIRE.PutValue('BSH_CODE',           theCode) ;
    TOBDEFAIRE.PutValue('BSH_DOMICILIATION',  theDomiciliation) ;
    TOBDEFAIRE.PutValue('BSH_DATEXE',         iDate1900) ;
    TOBDEFAIRE.PutValue('BSH_EXERCICE',       '') ;
    TOBDEFAIRE.PutValue('BSH_CONSTBQ',        '') ;
    TOBDEFAIRE.PutValue('BSH_POSANNEE',       0) ;
    TOBDEFAIRE.PutValue('BSH_LGNUMERO',       8) ;
    TOBDEFAIRE.PutValue('BSH_NUMERODEBUT',    0) ;
    TOBDEFAIRE.PutValue('BSH_NUMEROFIN',      99999999) ;
    TOBDEFAIRE.PutValue('BSH_LASTNUM',        0) ;

    ActionDefaire         := GetAction() ;
    ActionEcran           := GetActionEcran() ;
    EcraseLastNumDefaire  := EcraseLastNum ;
  end ;
  readCompteGeneralAuxiliaire()
END ;

{------------------------------------------------------------------------------}
Destructor TBanqueSh.Destroy() ;
BEGIN
  if (TOBBANQUESH <> Nil)  then   FreeAndNil(TOBBANQUESH) ;
  if (TOBDEFAIRE  <> Nil)  then   FreeAndNil(TOBDEFAIRE) ;
  if (TOBEXERCICE <> Nil)  then   FreeAndNil(TOBEXERCICE) ;
  if (TOBEXDEFAIRE <> Nil) then   FreeAndNil(TOBEXDEFAIRE) ;
  if (TOBDATERELEVE <> Nil) then  FreeAndNil(TOBDATERELEVE) ;
  Inherited Destroy ;
END ;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.UpdateDataBase ;
var
  sql     : String ;
  wCode   : String ;
  ENTETE  : String ;
  MSG     : String ;
  wdatexe : String ;
BEGIN
  CASE Action OF
    tBqRIEN :
      begin
        ENTETE  :=  TraduireMemoire('Erreur de petit programmeur') ;
        MSG     := TraduireMemoire('Impossible de faire une action dans la base de données.')
                +  #13#13
                +  TraduireMemoire('Le paramètre "Action" n''est pas en mode "taModif" ou "taCreat".') ;
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    else
    begin
      wCode   := GetCode() ;
      wdatexe := GetCodeExercice() ;

      sql := 'SELECT BSH_CODE FROM BANQUESH WHERE '
            + ' BSH_CODE              = "' + wCode    + '"'
            + ' AND BSH_EXERCICE      = "' + wdatexe  + '"' ;
      Case Action of
        tBqMODIF :
          begin
            if Not ExisteSQL(sql) then
            begin
              ENTETE :=  TraduireMemoire('Attention') ;
              MSG    := TraduireMemoire('Le compte bancaire "')
                     + wCode
                     +  TraduireMemoire('" n''existe pas pour cet exercice "')
                     + wdatexe
                     + '".' ;
              PGIError(MSG, ENTETE) ;
              V_PGI.IoError := oeUnknown ;
              Exit ;
            end ;
            TOBBANQUESH.UpdateDB(False) ;
          end ;
        tBqCREAT :
          begin
            if ExisteSQL(sql) then
            begin
              ENTETE :=  TraduireMemoire('Attention') ;
              MSG    := TraduireMemoire('Le compte bancaire "') + wCode
                     +  TraduireMemoire('" existe pour cet exercice "') + wdatexe
                     + '".' ;
              PGIError(MSG, ENTETE) ;
              V_PGI.IoError := oeUnknown ;
              Exit ;
            end ;
            TOBBANQUESH.InsertDB(Nil) ;
          end ;
      End ;
    end ;
  END ;
END ;
{------------------------------------------------------------------------------}
Function TBanqueSh.Defaire(): Boolean;
BEGIN

  TOBBANQUESH.ClearDetail() ;
  TOBBANQUESH.Dupliquer(TOBDEFAIRE, True, True, True);

  TOBEXERCICE.ClearDetail() ;
  TOBEXERCICE.Dupliquer(TOBEXDEFAIRE, True, True, True);

  Action        := ActionDefaire        ;
  ActionEcran   := ActionEcranDefaire   ;
  EcraseLastNum := EcraseLastNumDefaire ;


  result := True ;
END ;
{------------------------------------------------------------------------------}
function TBanqueSh.UpdateBD(): Boolean;
var
  IOERR   : TIOErr ;
  ENTETE  : String ;
  MSG     : String ;
BEGIN
  result := False ;


  IOERR := Transactions(UpdateDataBase, 0) ;
  if  (IOERR <> oeOK) And (IOERR <> oeSaisie) Then
  begin
    ENTETE  := TraduireMemoire('Erreur technique') ;
    MSG     := TraduireMemoire('Impossible de mettre à jour la table BANQUESH.') ;
    PGIInfo(MSG, ENTETE) ;
    Exit ;
  end ;
  result := True ;
END ;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.DeleteDataBase() ;
var
  sql     : String ;
  wCode   : String ;
  wdatexe : String ;
  ENTETE  : String ;
  MSG     : String ;
BEGIN
  ENTETE :=  TraduireMemoire('Attention') ;
  CASE Action OF
    tBqRIEN :
      begin
        ENTETE  :=  TraduireMemoire('Erreur de petit programmeur') ;
        MSG     := TraduireMemoire('Impossible de faire une action dans la base de données.')
                +  #13#13
                +  TraduireMemoire('Le paramètre "Action" n''est pas en mode "taModif" ou "taCreat".') ;
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    tBqCREAT :
      begin
        MSG    :=  TraduireMemoire('Impossible de supprimer une souche bancaire qui n''existe pas dans la base de données.');
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    tBqMODIF :
      begin
        MSG    :=  TraduireMemoire('Impossible de supprimer une souche bancaire dans le mode modification.');
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    tBqDELETE :
      begin
        wCode   := GetCode() ;
        wdatexe := GetCodeExercice() ;

        sql := 'SELECT BSH_CODE FROM BANQUESH WHERE '
            + ' BSH_CODE              = "' + wCode    + '"'
            + ' AND BSH_EXERCICE      = "' + wdatexe  + '"' ;
        if Not ExisteSQL(sql) then
        begin
          MSG    := TraduireMemoire('Le compte bancaire "')
                 + wCode
                 +  TraduireMemoire('" n''existe pas pour cet exercice "')
                 + wdatexe
                 + '".' ;
          PGIError(MSG, ENTETE) ;
          V_PGI.IoError := oeUnknown ;
          Exit ;
        end ;
        TOBBANQUESH.DeleteDB(False) ;
      end ;
  END ;
END ;

{------------------------------------------------------------------------------}
function TBanqueSh.DeleteBD(): Boolean;
var
  IOERR   : TIOErr ;
  ENTETE  : String ;
  MSG     : String ;
  wAction : TActionTableBq ;
BEGIN
  result := True ;
  wAction := GetAction() ;
  SetAction(tBqDELETE);
  ENTETE :=  TraduireMemoire('Erreur technique') ;

  IOERR := Transactions(DeleteDataBase, 0) ;
  if  (IOERR <> oeOK) And (IOERR <> oeSaisie) Then
  begin
    MSG := TraduireMemoire('Impossible de supprimer la souche de la table BANQUESH.') ;
    PGIError(MSG, ENTETE) ;
    SetAction(wAction);
    result := False ;
    Exit ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetDateDebutExercice() : TDateTime ;
begin
  result := TOBEXERCICE.GetValue('EX_DATEDEBUT') ;
end;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetDateFinExercice()   : TDateTime ;
begin
  result := TOBEXERCICE.GetValue('EX_DATEFIN') ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetPositionnementAnnee(theLgAnnee : Integer) ;
begin
  TOBBANQUESH.PutValue('BSH_POSANNEE', theLgAnnee) ;
end;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetPositionnementAnnee() : Integer ;
begin
  result := TOBBANQUESH.GetValue('BSH_POSANNEE') ;
end;

{------------------------------------------------------------------------------}
function TBanqueSh.GetAction(): TActionTableBq;
begin
  result := Action ;
end;

{------------------------------------------------------------------------------}
function TBanqueSh.GetActionEcran(): TActionECRAN ;
begin
  result := ActionEcran ;
end;

{------------------------------------------------------------------------------}
function TBanqueSh.GetEcraseLastNum(): Boolean ;
begin
  result := EcraseLastNum ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetEcraseLastNum(theEcraseLastNum : Boolean);
begin
  EcraseLastNum := theEcraseLastNum ;
end;

{------------------------------------------------------------------------------}
function TBanqueSh.GetCode(): String;
begin
  result := TOBBANQUESH.GetValue('BSH_CODE') ;
end;

{------------------------------------------------------------------------------}
function TBanqueSh.GetDomiciliation() : String;
begin
  result := TOBBANQUESH.GetValue('BSH_DOMICILIATION') ;
end;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetBanque() : String ;
var
  value : String ;
BEGIN
  result := '' ;
  value := GetDomiciliation() ;
  if (value = '') then Exit ;
  result := Copy(GetDomiciliation(), 5, 5) ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetGuichet() : String ;
var
  value : String ;
BEGIN
  result := '' ;
  value := GetDomiciliation() ;
  if (value = '') then Exit ;
  result := Copy(GetDomiciliation(), 10, 5) ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetCompte() : String ;
var
  value : String ;
BEGIN
  result := '' ;
  value := GetDomiciliation() ;
  if (value = '') then Exit ;
  result := Copy(GetDomiciliation(), 15, 11) ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetCleRib() : String ;
var
  value : String ;
BEGIN
  result := '' ;
  value := GetDomiciliation() ;
  if (value = '') then Exit ;
  result := Copy(GetDomiciliation(), 26, 2) ;
END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetCodeExercice() : String ;
begin
  result := TOBBANQUESH.GetValue('BSH_EXERCICE') ;
end;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetDateExercice() : TDateTime ;
begin
  result := TOBBANQUESH.GetValue('BSH_DATEXE') ;
end;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetTaille() : Integer ;
begin
  result := TOBBANQUESH.GetValue('BSH_LGNUMERO') ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetTaille(theTaille  : Integer) ;
var
  MSG     : String ;
  ENTETE  : String ;
BEGIN
  if (theTaille > 9) then
  begin
    MSG     := TraduireMemoire('La longueur du compteur ne peut pas dépasser 9 caractères.') ;
    ENTETE  := TraduireMemoire('Attention') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  TOBBANQUESH.PutValue('BSH_LGNUMERO', theTaille);
END ;

{------------------------------------------------------------------------------}
Function TBanqueSh.GetContanteBancaire() : String ;
begin
  result := TOBBANQUESH.GetValue('BSH_CONSTBQ') ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetContanteBancaire (theConstante : String) ;
BEGIN
  TOBBANQUESH.PutValue('BSH_CONSTBQ', theConstante);
END ;

{------------------------------------------------------------------------------}
function TBanqueSh.GetPremierNumero() : Integer ;
begin
  result := TOBBANQUESH.GetValue('BSH_NUMERODEBUT') ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetPremierNumero(theFirstSouche : Integer) ;
var
  currentNb : Integer ;
  ENTETE    : String ;
  MSG       : String ;
  firstSouche : Integer ;
BEGIN
 CASE Action OF
    tBqCREAT :
      begin
        TOBBANQUESH.PutValue('BSH_NUMERODEBUT',   theFirstSouche);
        SetCurrentNumero(theFirstSouche);
      end ;
    tBqMODIF :
      begin
        firstSouche := TOBBANQUESH.GetValue('BSH_NUMERODEBUT') ;
        currentNb := GetCurrentNumero();

        //* Si souche utilsée, on controle si le numéro de début est toujours inférieur au dernier numéro de souche utilsé.

        if Not GetEcraseLastNum() then
        begin
          if (currentNb <> 0) then
          begin
            if (firstSouche <> 0) then
            begin
              if (theFirstSouche <> firstSouche) then
              begin
                if (theFirstSouche > currentNb) then
                begin
                  ENTETE := TraduireMemoire('Attention') ;
                  MSG    := TraduireMemoire('Le numéro de début de la souche (')
                          + IntToStr(theFirstSouche) + ') '
                          + TraduireMemoire('est supérieur au dernier numéro de souche utilisé.')
                          + #13#13
                          + TraduireMemoire('Que voulez-vous faire ?')
                          + #13#13
                          + TraduireMemoire('OUI - Pour écraser le dernier numéro de souche utilisé.')
                          + #13
                          + TraduireMemoire('NON - Pour ressaisir le numéro de début de souche.') ;
                  if (PGIAsk(MSG, ENTETE) = mrNo) then  Exit ;
                end ;
              end ;
            end ;
          end ;
        end ;

        TOBBANQUESH.PutValue('BSH_NUMERODEBUT', theFirstSouche);

        if GetEcraseLastNum() then
                TOBBANQUESH.PutValue('BSH_LASTNUM', theFirstSouche);
      end ;
  END ;
END ;

{------------------------------------------------------------------------------}
function TBanqueSh.GetDernierNumero() : Integer ;
begin
  result := TOBBANQUESH.GetValue('BSH_NUMEROFIN') ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetDernierNumero(theLastSouche : Integer) ;
BEGIN
  TOBBANQUESH.PutValue('BSH_NUMEROFIN', theLastSouche);
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.GetCurrentNumero() : Integer ;
begin
  result := TOBBANQUESH.GetValue('BSH_LASTNUM') ;
end;


{------------------------------------------------------------------------------}
Procedure  TBanqueSh.SetCurrentNumero(theCurrentNumero : Integer) ;
begin
  TOBBANQUESH.PutValue('BSH_LASTNUM', theCurrentNumero) ;
end;


{------------------------------------------------------------------------------}
Procedure TBanqueSh.ChargerReferenceReleve(theDateReleve : TdateTime) ;
var
  SQL             : String ;
  QRY             : TQuery ;
  EE_REFPOINTAGE  : String ;
  constante       : String ;
  longueur        : Integer ;
  posDate         : Integer ;
  restant         : String ;
  reste           : Integer ;
  TRELEVE         : TOB ;
begin
  SQl  := 'SELECT MAX(EE_REFPOINTAGE) FROM EEXBQ WHERE '
        + ' EE_DATEPOINTAGE = "'    + usDateTime_(theDateReleve)  + '"'
        + ' AND EE_GENERAL = "'     + compteGeneral               + '"'
        + ' AND EE_AUXILIAIRE = "'  + compteAuxiliaire            + '"' ;
  QRY := OpenSQL(SQL, TRue) ;
  if Not QRY.EOF then
  begin
    EE_REFPOINTAGE := QRY.Fields[0].ASString ;
    if (EE_REFPOINTAGE <> '') then
    begin
      constante := GetContanteBancaire() ;
      longueur  := StrToInt(Copy(constante, 3, 1)) ;
      posDate   := GetPositionnementAnnee() ;
      restant   := Copy(EE_REFPOINTAGE, longueur + 2, 17 - (longueur + 1)) ;

      if (posDate <> 0) then
      begin
        reste := Length(restant) ;
        restant := Copy(restant, posDate + 1, reste - (posDate + 1)) ;
      end ;

      if IsNumeric(restant) then
      begin
        TRELEVE := TOB.Create('***', TOBDATERELEVE, -1) ;
        TRELEVE.AddChampSupValeur('DATER', theDateReleve, False);
        TRELEVE.AddChampSupValeur('COMPTEUR', StrToInt(restant), False);
      end ;
    end ;
  end ;
  Ferme(QRY) ;
end ;

{------------------------------------------------------------------------------
La Référence de pointage possede le format suivant :

- Une constante ou le compte général banque
- Une Année sur deux, quatre caractères ou la date du rélevé au format jjmmaa
- Compteur (d'une taille plus ou moins longue en fonction des deux autres paramètres)

}
function TBanqueSh.GetNextReferencePointage(theBUpdateBDFg: Boolean ; theDateReleve : TdateTime) : String ;
var
  compteur  : Integer ;
  ENTETE    : String ;
  MSG       : String ;
  frm       : String ;
  position  : Integer ;
  sAnnee    : String ;
  annee     : Word ;
  mois      : Word ;
  jour      : Word ;
  lg        : Integer ;
  constante : String ;
  boForCpte : Boolean ;
  longueur  : Integer ;
  param1    : String ;
  param2    : String ;
  param3    : String ;
  TRELEVE   : TOB ;
begin
  ENTETE   := TraduireMemoire('Attention') ;

  param1   := '' ;
  param2   := '' ;
  param3   := '' ;

  constante :=  GetContanteBancaire() ;

  //* Recherche le code année 0/2/4/6

  position := GetPositionnementAnnee() ;

  //* Paramétrage sur le compte  Oui/Non

  boForCpte := (Pos('$C', constante) <> 0) ;

  if Not boForCpte then
          param1 :=  constante
  else
  begin
    longueur  := StrToInt(Copy(constante, 3, 1)) ;
    param1     := Copy(compteGeneral, 1, longueur) ;
  end ;

  //* Détermination du n° de compteur

  lg        :=  GetTaille() ;
  frm       := StringOfChar('0', lg) ;

  //* FQ 27227    Si autre chose que la date de la référence de pointage le compteur est annuel

  if Not boForCpte
    or (position <> 6) then
        compteur  := GetCurrentNumero()
  else
  begin
    //* Comme le compteur est par jour, nous devons traiter différemment :
    //* On doit aller chercher


    TRELEVE := TOBDATERELEVE.FindFirst(['DATER'], [theDateReleve], False) ;
    if Not Assigned(TRELEVE) then
    begin
      //* Recup dans la base de donnée

      ChargerReferenceReleve(theDateReleve) ;

      TRELEVE := TOBDATERELEVE.FindFirst(['DATER'], [theDateReleve], False) ;
      if Not Assigned(TRELEVE) then
      begin
        TRELEVE := TOB.Create('***', TOBDATERELEVE, -1) ;
        TRELEVE.AddChampSupValeur('DATER', theDateReleve, False);
        TRELEVE.AddChampSupValeur('COMPTEUR', 0, False);
      end ;
    end ;
    compteur := TRELEVE.GetValue('COMPTEUR') ;
  end ;
  Inc(compteur) ;

  result := '' ;

  //* Pas de contrôle sur le n° de compteur  quand on est sur la date du relevé.

  if (compteur < GetPremierNumero())
    or (compteur > GetDernierNumero()) then
  begin
    MSG := TraduireMemoire('La référence de pointage est en dehors de la plage autorisée : ')
         +  intToStr(GetPremierNumero())
         +  ' - '
         + intToStr(GetDernierNumero())
         +  ' - '
         + TraduireMemoire('Compteur de la référence de pointage : ') + intToStr(compteur);
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  SetCurrentNumero(compteur) ;
  param3 := FormatFloat(frm, compteur)  ;

  //* Mise à jour du numéro de compteur

  if theBUpdateBDFg and Not boForCpte then
  begin
   if Not UpdateBD() then Exit ;
  end ;

  case position of
    0 :
        begin
          result := param1 + ' ' + param3  ;
          refPointage :=  Result ;
        end ;
    2 :
        begin
          DecodeDate(GetDateFinExercice(), annee, mois, jour);
          sAnnee := IntToStr(annee) ;
          sAnnee := Copy(sAnnee, 3, 2) ;
          result := param1 + ' ' +  sAnnee + ' ' + param3  ;
          refPointage :=  Result ;
        end ;
    4 :
        begin
          DecodeDate(GetDateFinExercice(), annee, mois, jour);
          result := param1 + ' ' +  IntToStr(annee) + ' ' + param3  ;
          refPointage :=  Result ;
        end ;
    6 :
        begin
          DecodeDate(theDateReleve, annee, mois, jour);
          sAnnee := IntToStr(annee) ;
          Annee := StrToInt(Copy(sAnnee, 3, 2)) ;
          result := param1 + ' ' + FormatFloat('00', jour)
                                 + FormatFloat('00', mois)
                                 + FormatFloat('00', Annee) + ' ' + param3  ;
          refPointage :=  Result ;
        end ;
  end ;
end;

{-------------------------------------------------------------------------------}
Procedure TBanqueSh.InitReferencePointage(theDateReleve : TDateTime);
var
  compteur  : Integer ;
  frm       : String ;
  position  : Integer ;
  sAnnee    : String ;
  annee     : Word ;
  mois      : Word ;
  jour      : Word ;
  lg        : Integer ;
  constante : String ;
  boForCpte : Boolean ;
  longueur  : Integer ;
  param1    : String ;
  param2    : String ;
  param3    : String ;
begin

  param1   := '' ;
  param2   := '' ;
  param3   := '' ;

  constante :=  GetContanteBancaire() ;

  //* Paramétrage sur le compte

  boForCpte := (Pos('$C', constante) <> 0) ;

  if Not boForCpte then param1 :=  constante
  else
  begin
    longueur  := StrToInt(Copy(constante, 3, 1)) ;
    param1     := Copy(compteGeneral, 1, longueur) ;
  end ;

  lg        :=  GetTaille() ;
  frm       := StringOfChar('0', lg) ;
  compteur  := GetCurrentNumero() ;

  param3 := FormatFloat(frm, compteur)  ;

 //* Recherche le code année 0/2/4/6

  position := GetPositionnementAnnee() ;
  case position of
    0 :
        begin
          refPointage := param1 + ' ' + param3  ;
        end ;
    2 :
        begin
          DecodeDate(GetDateFinExercice(), annee, mois, jour);
          sAnnee := IntToStr(annee) ;
          sAnnee := Copy(sAnnee, 3, 2) ;
          refPointage := param1 + ' ' +  sAnnee + ' ' + param3  ;
        end ;
    4 :
        begin
          DecodeDate(GetDateFinExercice(), annee, mois, jour);
          refPointage := param1 + ' ' +  IntToStr(annee) + ' ' + param3  ;
        end ;
    6 :
        begin
          DecodeDate(theDateReleve, annee, mois, jour);
          refPointage := param1 + ' ' + FormatFloat('00', jour)
                                      + FormatFloat('00', mois)
                                      + FormatFloat('00', annee) + ' ' + param3  ;
        end ;
  end ;
end;


{------------------------------------------------------------------------------}
function TBanqueSh.GetReferencePointage(): String ;
begin
  result := refPointage ;
end;


{------------------------------------------------------------------------------}
function TBanqueSh.GetTOBBANQUESH(): TOB ;
begin
  Result := TOB.Create('', Nil, -1) ;
  result.Dupliquer(TOBBANQUESH, True, True, True) ;
end;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetAction(theAction : TActionTableBq)  ;
BEGIN
  action := theAction ;
END ;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetActionEcran(theActionEcran : TActionECRAN)  ;
BEGIN
  actionEcran := theActionEcran ;
END ;
{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetCode(theCode : String)  ;
BEGIN
  TOBBANQUESH.PutValue('BSH_CODE', theCode);
END ;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetDomiciliation(theDomiciliation : String)  ;
BEGIN
  TOBBANQUESH.PutValue('BSH_DOMICILIATION', theDomiciliation);
END ;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetCodeExercice(theCodeExercice : String)  ;
var
  SQL : String ;
  QRY : TQuery ;
  WTOBLIST : TOB ;
BEGIN
  TOBBANQUESH.PutValue('BSH_EXERCICE', theCodeExercice);
  if (theCodeExercice <> '') then
  begin
    SQL := 'SELECT * FROM EXERCICE WHERE EX_EXERCICE = "' + theCodeExercice + '"'  ;
    QRY := OpenSQL(SQL, True) ;
    if Not QRY.EOF then
    begin
      WTOBLIST := TOB.Create('', Nil, -1) ;
      WTOBLIST.LoadDetailDB('EXERCICE', '', '', QRY, True) ;
      TOBEXERCICE.ClearDetail() ;
      TOBEXERCICE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);
      FreeAndNil(WTOBLIST) ;
      TOBBANQUESH.PutValue('BSH_DATEXE',TOBEXERCICE.GetValue('EX_DATEFIN')) ;
    end
    else TOBEXERCICE.ClearDetail() ;
    Ferme(QRY) ;
  end
  else TOBEXERCICE.ClearDetail() ;
END ;

{------------------------------------------------------------------------------}
Procedure TBanqueSh.SetDateExercice(theDateExercice  : TdateTime) ;
var
  SQL : String ;
  QRY : TQuery ;
  WTOBLIST : TOB ;
BEGIN
  TOBBANQUESH.PutValue('BSH_DATEXE', theDateExercice);
  if (theDateExercice <> iDate1900)
    and (theDateExercice <> iDate2099) then
  begin
    SQL := 'SELECT * FROM EXERCICE WHERE EX_DATEFIN = "' + usdatetime_(theDateExercice) + '"'  ;
    QRY := OpenSQL(SQL, True) ;
    if Not QRY.EOF then
    begin
      WTOBLIST := TOB.Create('', Nil, -1) ;
      WTOBLIST.LoadDetailDB('EXERCICE', '', '', QRY, True) ;
      TOBEXERCICE.ClearDetail() ;
      TOBEXERCICE.Dupliquer(WTOBLIST.Detail[WTOBLIST.Detail.Count -1], True, True, True);
      FreeAndNil(WTOBLIST) ;
      TOBBANQUESH.PutValue('BSH_EXERCICE',TOBEXERCICE.GetValue('EX_EXERCICE')) ;
    end
    else TOBEXERCICE.ClearDetail() ;
    Ferme(QRY) ;
  end
  else TOBEXERCICE.ClearDetail() ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsCode()  : Boolean ;
var
  ENTETE    : String ;
  MSG       : String ;
  code      : String ;
  SQL       : String ;
BEGIN
  result  := False ;

  code := GetCode() ;
  if (code = '') then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Le code banque doit être renseigné.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  SQL := 'SELECT * FROM BANQUECP WHERE BQ_CODE = "' + code + '"'
      +  ' AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"'  ;
  if Not ExisteSQL(SQL) then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Le code banque n''est pas correct.') ;
    PGIError(MSG, ENTETE);
    Exit;
  end ;

  result := True ;
END ;


{------------------------------------------------------------------------------}
Function  TBanqueSh.IsCodeExercice()  : Boolean ;
var
  ENTETE    : String ;
  MSG       : String ;
  code      : String ;
  SQL       : String ;
BEGIN
  result  := False ;

  code := GetCodeExercice() ;
  if (code = '') then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Le code exercice doit être renseigné.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  SQL := 'SELECT * FROM EXERCICE WHERE EX_EXERCICE = "' + code + '"' ;
  if Not ExisteSQL(SQL) then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('Le code exercice n''est pas correct.') ;
    PGIError(MSG, ENTETE);
    Exit;
  end ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsDomiciliation(theDomiciliation : String = '')  : Boolean ;
var
  ENTETE    : String ;
  MSG       : String ;
  data      : String ;
BEGIN
  result := False ;
  data    := theDomiciliation ;
  if (data = '') then data := GetDomiciliation() ;
  if (data = '') then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('La domiciliation doit être renseignée.')  ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsConstante()  : Boolean ;
var
  ENTETE    : String ;
  MSG       : String ;
BEGIN
  result := False ;

  if (GetContanteBancaire() = '') then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('La constante bancaire doit être renseignée.')  ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsTaille()  : Boolean ;
var
  ENTETE : String ;
  MSG    : String ;
BEGIN
  result := False ;

  if (GetTaille() > 9) then
  begin
    MSG     := TraduireMemoire('La longueur du compteur ne peut pas dépasser 9 caractères.') ;
    ENTETE  := TraduireMemoire('Attention') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsBorneSoucheOk()  : Boolean ;
var
  ENTETE    : String ;
  MSG       : String ;
BEGIN
  result := False ;
  if (GetPremierNumero() > GetDernierNumero()) then
  begin
    ENTETE := TraduireMemoire('Attention') ;
    MSG    := TraduireMemoire('La borne souche bancaire est invalide.')  ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsAllControler(var theNumChampError : Integer)  : Boolean ;
BEGIN
  result := IsControler(theNumChampError, True,  True, True) ;
END ;

{------------------------------------------------------------------------------}
Function  TBanqueSh.IsControler(var theNumChampError : Integer ; theBoCode,  theBoExercice, boParametrageRef : Boolean)  : Boolean ;
begin
  result := False ;
  theNumChampError := 0 ;

  if (Action in [tBqRIEN, tBqCREAT, tBqMODIF]) then
  begin
    if theBoCode then
    begin
      if Not IsCode() then
      begin
        theNumChampError := 1 ;
        Exit ;
      end ;

      if Not IsDomiciliation()  then
      begin
        theNumChampError := 2 ;
        Exit ;
      end ;
    end ;

    if theBoExercice then
    begin
      if Not IsCodeExercice() then
      begin
        theNumChampError := 3 ;
        Exit ;
      end ;
    end ;

    if boParametrageRef then
    begin
      if Not IsConstante () then
      begin
        theNumChampError := 4 ;
        Exit ;
      end ;

      if Not IsTaille () then
      begin
        theNumChampError := 5 ;
        Exit ;
      end ;
      if Not IsBorneSoucheOk()  then
      begin
        theNumChampError := 6 ;
        Exit ;
      end ;
    end ;
  end ;

  result := True ;
end ;

end.
