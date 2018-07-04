{***********UNITE**************************************************************
Auteur  ...... : Sylvie DE ALMEIDA
Créé le ...... : 03/2007
Modifié le ... :
Description .. : Fonctions et procédures pour gestion de la taxe internationale
Mots clefs ... :
*******************************************************************************}
unit YYInternationalTax;

interface
uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    uTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DB,
    {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$ENDIF}
    utilPGI,
    EntPGI,
    HCtrls,
    HEnt1, 
    utobdebug;
const
T1 =  'CCT_CODEMODELE';
T2 =  'CCT_CODECAT';
T3 =  'CCT_CODETYP';
T4 =  'CCT_CODEREG';
T5 =  'CCT_PAYS';
T6 =  'CCT_PERDEB';
T7 =  'CCT_PERFIN';
T8 =  'CCT_ASSDEB';
T9 =  'CCT_ASSFIN';


type

 TInternationalTax = class
 public
    FTOBPARAMETRESTAXES  : tMemoryTob;
    Constructor Create;
    Destructor  Destroy; override;
    Function    GetInternationalTax ( TobQuestion : TOB ; bAll : boolean) : Tob;
    Function    GetSqlInternatTaxe: string;
    Function    strGetPaysEtablissement (strEtablissement : string) : string;
    Function    strGetPaysEtabModele (strModele : string) : string;
    Function    strGetTypeQualifiant (strType : string) : string;
    Function    iChercheTobParametrestaxes (strModele,strCategorie,strType,strRegime : string; strDatePiece : TDatetime; dMontant : double; strCriLocalisation, strCP, strVille, strRegion, strPays : string; bCherchePaysVide : boolean):integer;
    Function    TobChercheEnReduisantUnElement (strModele,strCategorie,strType,strRegime : string; strDatePiece : TDatetime; dMontant : double; strCriLocalisation, strCP, strVille, strRegion, strPays, strTaxesIncluses, strFlux, strElementEnleve : string): TOB;
  //Function    bChercheTobParametrestaxes (strModele,strCategorie,strType,strRegime : string; strDatePiece : TDatetime; dMontant : double; strCriLocalisation, strCP, strVille, strRegion, strPays : string):boolean;
    Function    RechercheInfosDansBase ( strModele, strCategorie, strType, strRegime : string) : boolean;
    Function    tAlimenteLigneReponse ( TobQuestionDetail : TOB; strCriLocalisation, strTaxesIncluses : string; bCherchePaysVide : boolean ) : Tob;
    Function    GetNomChampParam(Num : integer) : string;
    Procedure   SetParamInTobquestion(TobAlimQuestion : TOB; NumChamp : integer; stValeur : variant);
    Procedure   tCreeLigneQuestion (TobAlimQuestion : Tob; strModele, strCategorie, strType, strRegime, strEtablissement, strFlux, strCP, strVille, strRegion, strPays  : string ; tDatepiece : TDatetime; iQuantiteLigne : integer; dMontantLigne : double );
    Function    tstrGetLibelleTaxesModele(strModele : string) : String;
    Function    tstrGetLibelleTypesTaxesModele(strModele, strCategorie : string) : String;
    Function    tstrGetLibelleRegimesTaxesModele(strModele, strCategorie : string) : String;
    Function    tstrGetCodeTypesTaxesModele(strModele, strCategorie : string) : String;
    Function    tstrGetCodeRegimesTaxesModele(strModele, strCategorie : string) : String;
    Function    bChercheSiRien (strModele,strCategorie,strType,strRegime : string) : boolean;
    Function    TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays : string; strDatePiece : TDateTime; dMontant : double; bAvecMontant : boolean) : TDateTime;
    Function    iChercheInfosExoneration (strModele,strCategorie,strType,strRegime : string; strDatePiece : TDatetime; dMontant : double; strCriLocalisation, strCP, strVille, strRegion, strPays : string):integer;
 end;

implementation

constructor TInternationalTax.Create;
begin
 FTOBPARAMETRESTAXES := tMemoryTob.Create('InternationalTax',GetSqlInternatTaxe);
end;

function TInternationalTax.GetSqlInternatTaxe: string;
begin
  Result := 'SELECT YDETAILTAXES.* FROM YDETAILTAXES INNER JOIN YMODELETAXE '
                                       + ' ON YDETAILTAXES.CCT_CODEMODELE = YMODELETAXE.YMT_CODE '
                                       + ' WHERE YMT_PRINCIPAL = "X" '
                                       + ' ORDER BY '
                                       + ' YDETAILTAXES.CCT_CODECAT, '
                                       + ' YDETAILTAXES.CCT_CODETYP, '
                                       + ' YDETAILTAXES.CCT_CODEREG, '
                                       + ' YDETAILTAXES.CCT_PERDEB, '
                                       + ' YDETAILTAXES.CCT_PERFIN, '
                                       + ' YDETAILTAXES.CCT_ASSDEB, '
                                       + ' YDETAILTAXES.CCT_ASSFIN, '
                                       + ' YDETAILTAXES.CCT_CODEPOSTAL, '
                                       + ' YDETAILTAXES.CCT_VILLE, '
                                       + ' YDETAILTAXES.CCT_REGION, '
                                       + ' YDETAILTAXES.CCT_PAYS'
end;

destructor TInternationalTax.Destroy;
begin
 if assigned(FTOBPARAMETRESTAXES) then FTOBPARAMETRESTAXES.Free;
 FTOBPARAMETRESTAXES := nil;
 inherited;
end;

function TInternationalTax.GetInternationalTax ( TobQuestion : TOB ; bAll : boolean) : Tob;
var
    TobResult, TobResult2, TStockageM : TOB ;
    strModele, strCategorie, strType, strRegime : string;
    strEtablissement, strFlux  : string;
    strDatePiece : TDateTime;
    i, k, n: integer;
    bRien : boolean;
    strSQL, strTaxesIncluses, strCriLocalisation : string;
    OQuery : TQuery;
    strCP, strVille, strRegion, strPays : string;
    iQuantiteLigne, iNombre, iNombre2 : integer;
    dMontantLigne : double;
    bEstTrouve : Boolean;
begin


  // Préparation de la tob retournée
  TOBResult := TOB.Create('ResultInternationalTax', nil, -1) ;
  TOBResult.ClearDetail;
  for n:= 0 to TobQuestion.Detail.count -1 do
    TOBResult2 := TOB.Create('ResultInternationalTax2', TOBResult, -1) ;

  k := 0;
  for i := 0 to TobQuestion.Detail.Count - 1 do
    begin
      bEstTrouve := false;
      iNombre2 := 0;
      //Récupération de la ligne
      strModele := TobQuestion.detail[i].GetValue('MEM_MODELE');
      strCategorie := TobQuestion.detail[i].GetValue('MEM_CATEGORIE');
      strType := TobQuestion.detail[i].GetValue('MEM_TYPE');
      strRegime := TobQuestion.detail[i].GetValue('MEM_REGIME');
      strEtablissement := TobQuestion.detail[i].GetValue('MEM_ETABLISSEMENT');
      strFlux := TobQuestion.detail[i].GetValue('MEM_FLUX');
      strCP := TobQuestion.detail[i].GetValue('MEM_CPLIV');
      strVille := TobQuestion.detail[i].GetValue('MEM_VILLELIV');
      strRegion := TobQuestion.detail[i].GetValue('MEM_REGIONLIV');
      strPays :=  TobQuestion.detail[i].GetValue('MEM_PAYSLIV');
      strDatePiece := TobQuestion.detail[i].GetValue('MEM_DATEPIECE');
      iQuantiteLigne := TobQuestion.detail[i].GetValue('MEM_QUANTITELIGNE');
      dMontantLigne := TobQuestion.detail[i].GetValue('MEM_MONTANTLIGNE');

      //Si le pays n'est pas passé à la fonction on récupère le pays de l'établissement
      //if (strPays = '') then
      //  strPays := strGetPaysEtablissement (strEtablissement);

      //Recherche taxes incluses et niveau de recherche concernant la localisation
      strSQL := 'SELECT YMODELETAXE.YMT_CAT' + copy(strCategorie,3,1) + 'DEP, YMODELETAXE.YMT_CRI' + copy(strCategorie,3,1) + 'LOC FROM YMODELETAXE WHERE ';
      strSQL := strSQL + ' YMT_CODE = "'+ strModele +'"';
      OQuery := OpenSql (strSQL, TRUE);
      try
      While Not OQuery.Eof do
      BEGIN
        if OQuery.Fields[0].AsString <> '<<Aucun>>' then
          strTaxesIncluses := OQuery.Fields[0].AsString
        else
           strTaxesIncluses := '';
        strCriLocalisation := OQuery.Fields[1].AsString;
        OQuery.Next ;
      END ;
      finally
        Ferme(OQuery) ;
      end;

      //if bAll then //toutes les informations sont demandées par l'application demandeuse
      //  begin
        // Premier test : je regarde si les informations importantes sont présentes dans la TOB mémoire et la base de données
        // si OUI -> on continue
        // si NON -> on envoie RIE dans MEM_REPONSE
        bRien := bChercheSiRien (strModele,strCategorie,strType,strRegime);  //TOB MEMOIRE
        if bRien then  // infos non trouvées dans tob mémoire
        begin
          bRien := RechercheInfosDansBase (strModele, strCategorie, strType, strRegime ); // Recherche BD
          if bRien then // infos non trouvées ni dans tob mémoire ni en base -> on envoie RIE comme réponse
          begin
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
            TOBResult.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
            TOBResult.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
            TOBResult.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', '' , false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_MTMINI', '' , false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_FORMULE', '' , false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', '' , false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_EXO', '' , false);
            TOBResult.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'RIE' , false);
            k:= k + 1;
          end;

        end;

        //Deuxième étape : on commence par voir si tout est OK (tous les éléments sont trouvés)
        //Est considéré OK un seul enregistrement trouvé avec clé qui correspond à la recherche,
        //période trouvée, assiette trouvée, critère de localisation trouvé avec une particularité sur le pays
        //pays est OK si pays trouvé ou au minimum si pays à vide trouvé

        //Si + de 1 enr. OK trouvé -> réponse est PAR
        iNombre := 0;
        if not bRien then
        begin
          iNombre := iChercheTobParametrestaxes (strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays, false );
          if iNombre >= 1 then // un enr OK trouvé : réponse OK et plus de 1 enr trouvés : réponse PAR
          begin
            TobResult.detail[k] := tAlimenteLigneReponse (TobQuestion.detail[i], strCriLocalisation, strTaxesIncluses, false);
            k := k + 1;
          end;
        end;

        //Troisième étape : cas d'exonération tout est trouvé mais pays non trouvé
        if (iNombre = 0) and (not bRien) then
        begin
          iNombre := iChercheTobParametrestaxes (strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays, true );
          if iNombre = 0 then // pas trouvé -> regarde si exo pur
          begin
            iNombre2 := iChercheInfosExoneration (strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays );
            //if iNombre2 = 0 then
            //  iNombre2 := iChercheInfosExoneration (strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays, true );
            if iNombre2 = 1 then //Vrai cas d'exonération
            begin
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
              TOBResult.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
              TOBResult.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
              TOBResult.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', '' , false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_MTMINI', '' , false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_FORMULE', '' , false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', '' , false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_EXO', 'O' , false);
              TOBResult.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
              k:= k + 1;
            end;
          end;
          if iNombre >= 1 then // un enr OK trouvé : réponse OK et plus de 1 enr trouvés : réponse PAR
          begin
            TobResult.detail[k] := tAlimenteLigneReponse (TobQuestion.detail[i], strCriLocalisation, strTaxesIncluses, true);
            k := k + 1;
          end;
        end;

        //Quatrième étape : cas PAR ou NOK (NOK = 1 enr. trouvé en procédant au filtre ascendant donc réponse
        //non fiable à traiter par l'application demandeuse)
        if (iNombre = 0) and (iNombre2 <> 1 )and (not bRien) then
        begin

        //Tob de stockage mère et fille
        TStockageM := TOB.Create('TobTravailM', nil, -1) ;

        //Filtre sans pays
        TStockageM := TobChercheEnReduisantUnElement(strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays, strTaxesIncluses, strFlux, 'PAYS' );
        if TStockageM.Detail.Count >= 1 then // si 1 enr réponse NOK et plusieurs enr. PAR
        begin
            TobResult.detail[k] := TStockageM.Detail[TStockageM.Detail.Count - 1];
            bEstTrouve := True;
            k := k + 1;
        end
        else
          begin
          TStockageM := TobChercheEnReduisantUnElement(strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays, strTaxesIncluses, strFlux, 'ASSIETTE' );
          if TStockageM.Detail.Count >= 1 then // si 1 enr réponse NOK et plusieurs enr. PAR
          begin
            TobResult.detail[k] := TStockageM.Detail[TStockageM.Detail.Count - 1];
            bEstTrouve := True;
            k := k + 1;
          end
          else
            begin
              TStockageM := TobChercheEnReduisantUnElement(strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays, strTaxesIncluses,strFlux, 'PERIODE' );
              if TStockageM.Detail.Count >= 1 then // si 1 enr réponse NOK et plusieurs enr. PAR
              begin
                TobResult.detail[k] := TStockageM.Detail[TStockageM.Detail.Count - 1];
                bEstTrouve := True;
                k := k + 1;
              end
            end
          end
        end;

        if (iNombre = 0) and (iNombre2 <> 1 )and (not bRien) and (not bEstTrouve) then
        //Recherche infructueuse -> réponse RIE
        begin
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_MTMINI', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_FORMULE', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_EXO', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'RIE' , false);
          k:= k + 1;
        end;

       { if not bRien then  // première étape OK - à partir de là, la réponse est soit OK, soit EXO soit PAR
        begin
        //Recherche si la ligne réponse se trouve dans la table des paramétrages taxes
        bTrouveEnr := bChercheTobParametrestaxes (strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays );

        //Si non trouvé dans la tob des paramètres taxes, je cherche si l'élément est dans la base
        if not bTrouveEnr then
          bTrouveEnr2 := RechercheInfosDansBase (strModele, strCategorie, strType, strRegime );

        if bTrouveEnr2 then
          bTrouveEnr3 := bChercheTobParametrestaxes (strModele, strCategorie, strType, strRegime, strDatePiece, dMontantLigne, strCriLocalisation, strCP, strVille, strRegion, strPays );

        //Si non trouvé dans TOB des paramètres taxes ET en base de données : EXONERATION
        if not bTrouveEnr and not bTrouveEnr3 then
        begin
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_MTMINI', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_FORMULE', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_EXO', 'O' , false);
          k:= k + 1;
        end
        else
        begin
          //Ligne trouvée (une seule) : on alimente la ligne résultat
          TobResult.detail[k] := tAlimenteLigneReponse (TobQuestion.detail[i], strCriLocalisation, strTaxesIncluses);
          k := k + 1;
        end }
      {end
      else
      begin
          //ici bAll = false on n'envoie que les taxes incluses
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
          TOBResult.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_MTMINI', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_FORMULE', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_EXO', '' , false);
          TOBResult.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
          k:= k + 1;
      end; }

    result := TobResult ; // retour résultat(s)
end;
end;

function TInternationalTax.RechercheInfosDansBase ( strModele, strCategorie, strType, strRegime : string ) : boolean;
var
strSQL : string;
OQuery : TQuery;
begin
    //fonction qui permet d'alimenter la TOB des paramètres taxes via BD
    result := false; //initialisation à "faux" de la fonction

    // Recherche dans base de données et alimentation des enregistrements dans la TOB des paramétrages

    strSQL := 'SELECT CCT_CODEMODELE, ';  //0
	  strSQL := strSQL + 'CCT_CODECAT, '; //1
	  strSQL := strSQL + 'CCT_CODETYP, '; //2
	  strSQL := strSQL + 'CCT_CODEREG, '; //3
	  strSQL := strSQL + 'CCT_CODEPOSTAL, '; //4
	  strSQL := strSQL + 'CCT_VILLE,'; //5
	  strSQL := strSQL + 'CCT_REGION, '; //6
	  strSQL := strSQL + 'CCT_PAYS,'; //7
	  strSQL := strSQL + 'CCT_PERDEB, ';  //8
	  strSQL := strSQL + 'CCT_PERFIN, ';  //9
	  strSQL := strSQL + 'CCT_ASSDEB, ';  //10
	  strSQL := strSQL + 'CCT_ASSFIN, ';  //11
	  strSQL := strSQL + 'CCT_MTMINI, ';  //12
	  strSQL := strSQL + 'CCT_TXMTACHAT, ';//13
	  strSQL := strSQL + 'CCT_TXMTVENTE, ';//14
	  strSQL := strSQL + 'CCT_CPTACHAT, '; //15
	  strSQL := strSQL + 'CCT_CPTVENTE, '; //16
	  strSQL := strSQL + 'CCT_ENCACHAT, '; //17
	  strSQL := strSQL + 'CCT_ENCVENTE, '; //18
	  strSQL := strSQL + 'CCT_FORMULE FROM YDETAILTAXES WHERE ';
    strSQL := strSQL + 'CCT_CODEMODELE = "'+strModele+'"';
    strSQL := strSQL + ' AND CCT_CODECAT = "'+strCategorie+'"';
    strSQL := strSQL + ' AND CCT_CODETYP = "'+strType+'"';
    strSQL := strSQL + ' AND CCT_CODEREG = "'+strRegime+'"';
    strSQL := strSQL + ' ORDER BY ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_CODECAT, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_CODETYP, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_CODEREG, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_PERDEB, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_PERFIN, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_ASSDEB, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_ASSFIN, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_CODEPOSTAL, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_VILLE, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_REGION, ';
    strSQL := strSQL + ' YDETAILTAXES.CCT_PAYS';

    OQuery := OpenSql (strSQL, TRUE);
    try
    if OQuery.Eof then
      result := true
    else
      FTOBPARAMETRESTAXES.LoadDetailDBFromSQL( 'InternationalTax', strSql, True ) ;
    finally
      Ferme(OQuery);
    end;
end;

{function TInternationalTax.bChercheTobParametrestaxes (strModele,strCategorie,strType,strRegime : string; strDatePiece : TDatetime; dMontant : double; strCriLocalisation, strCP, strVille, strRegion, strPays : string):boolean;
var
j, iTrouve : integer;
TFind : Tob;
strDateDeb, strDateFin : tDatetime;
dAssietteDeb, dAssietteFin : double;
sCP, sVille, sRegion, sPays : string;

begin

result := false; //initialisation à "faux" de la fonction

TFind := nil;

iTrouve := 0; // initialisation nombre enr. trouvé

for j := 0 to FTOBPARAMETRESTAXES.Detail.count - 1 do
begin
  //Recherche enregistrement dans tob de paramétrage
  TFind := FTOBPARAMETRESTAXES.FindFirst(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  While TFind<>Nil do
  BEGIN
    strDateDeb := TFind.GetValue('CCT_PERDEB') ;
    strDateFin := TFind.GetValue('CCT_PERFIN') ;
    dAssietteDeb := TFind.GetValue('CCT_ASSDEB') ;
    dAssietteFin := TFind.GetValue('CCT_ASSFIN') ;
    sCP :=  TFind.GetValue('CCT_CODEPOSTAL') ;
    sVille :=  TFind.GetValue('CCT_VILLE') ;
    sRegion := TFind.GetValue('CCT_REGION') ;
    sPays :=  TFind.GetValue('CCT_PAYS') ;

    //if (sPays = '') then // on prend le pays de l'établissement sur lequel est l'établissement
    //  sPays := strGetPaysEtabModele (strModele);

    if (strDatePiece >= strDateDeb) and (strDatePiece <= strDateFin)
    and (dMontant >= dAssietteDeb) and (dMontant <= dAssietteFin) then
    begin
      //Recherche concernant le lieu de livraison
      if (strCriLocalisation = 'PAS') and ((strPays = sPays)) then // pas de recherche prévue sur lieu livraison au niveau catégorie
      begin
        result := true;
        iTrouve := iTrouve + 1;
        //exit;
      end
      else
      begin
        if (strCriLocalisation = 'CPO') and (strCP <> '') then
        begin
          if (sCP = strCP) or (sVille = strVille) or (sRegion = strRegion) or (sPays = strPays) then
          begin
            result := true;
            iTrouve := iTrouve + 1;
            //exit;
          end
        end
        else
        begin
          if ((strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO'))  and (strVille <> '') then
          begin
            if (sVille = strVille)  or (sRegion = strRegion) or (sPays = strPays) then
            begin
              result := true;
              iTrouve := iTrouve + 1;
              //exit;
            end
          end
          else
          begin
            if ((strCriLocalisation = 'REG') or (strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO')) and (strRegion <> '') then
            begin
              if (sRegion = strRegion)or (sPays = strPays) then
              begin
                result := true;
                iTrouve := iTrouve + 1;
                //exit;
              end
            end
            else
            if ((strCriLocalisation = 'REG') or (strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO')) and (strPays <> '') then
              begin
                if (sPays = strPays) then
                begin
                  result := true;
                  iTrouve := iTrouve + 1;
                end
              end
            end
        end
      end;
    end;
    TFind:=FTOBPARAMETRESTAXES.FindNext(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
   END ;
   if iTrouve > 1 then // Plusieurs enr trouvés prob. PAR
    result := false;
END ;
end; }

function TInternationalTax.tAlimenteLigneReponse ( TobQuestionDetail : TOB; strCriLocalisation, strTaxesIncluses : string ; bCherchePaysVide : boolean) : Tob;
var
TobLigneReponse0, TobLigneReponse: Tob;
strFormule : string;
strModele, strCategorie, strType, strRegime, strFlux, strQualifiant : string;
TFind, TCodeP, TCodePBis,  TVille, TVilleD, TRegion,  TRegionD, TPays, TPaysD: Tob;
strDateDeb, strDateFin, strDatePiece : tDatetime;
dAssietteDeb, dAssietteFin : double;
sCP, sVille, sRegion, sPays, strEtablissement: string;
strCP, strVille, strRegion, strPays : string;
iQuantite : integer;
dMontant, dTxMt, dMini : double;
iTrouve : integer;
TDateMaxDateDeb : TDateTime;
begin
//1 ligne réponse pour 1 ligne question

//On alimente ici une ligne OK si un seul enr trouvé sinon PAR

TobLigneReponse := TOB.Create('AlimenteReponseInternationalTaxe', nil, -1);
TobLigneReponse0 := TOB.Create('AlimenteReponseInternationalTaxe', TobLigneReponse, -1);

TPays := Tob.Create('Tob Pays', nil, -1);
TPaysD := Tob.Create('Tob PaysBis', TPays, -1);

TCodeP := Tob.Create('Tob CP', nil, -1);
TCodePBis := Tob.Create('Tob CP Bis', TCodeP, -1);

TVille := Tob.Create('Tob Ville', nil, -1);
TVilleD := Tob.Create('Tob VilleBis', TVille, -1);

TRegion := Tob.Create('Tob Reg', nil, -1);
TRegionD := Tob.Create('Tob RegBis', TRegion, -1);

try
begin
//Initialisation
TFind := nil;

//Récupération des valeurs de la ligne question
strModele := TobQuestionDetail.GetValue('MEM_MODELE');
strCategorie:= TobQuestionDetail.GetValue('MEM_CATEGORIE');
strType := TobQuestionDetail.GetValue('MEM_TYPE');
strRegime := TobQuestionDetail.GetValue('MEM_REGIME');
strFlux := TobQuestionDetail.GetValue('MEM_FLUX');
strEtablissement := TobQuestionDetail.GetValue('MEM_ETABLISSEMENT');
strCP := TobQuestionDetail.GetValue('MEM_CPLIV');
strVille  := TobQuestionDetail.GetValue('MEM_VILLELIV');
strRegion := TobQuestionDetail.GetValue('MEM_REGIONLIV');
strPays  := TobQuestionDetail.GetValue('MEM_PAYSLIV');
strDatePiece := TobQuestionDetail.GetValue('MEM_DATEPIECE');
iQuantite :=  TobQuestionDetail.GetValue('MEM_QUANTITELIGNE');
dMontant :=  TobQuestionDetail.GetValue('MEM_MONTANTLIGNE');

strQualifiant := strGetTypeQualifiant (strType);

TDateMaxDateDeb := TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays, strDatePiece, dMontant, true);

iTrouve := 0;
  //Recherche enregistrement dans tob de paramétrage
  TFind := FTOBPARAMETRESTAXES.FindFirst(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  While TFind<>Nil do
  BEGIN
    strDateDeb := TFind.GetValue('CCT_PERDEB') ;
    strDateFin := TFind.GetValue('CCT_PERFIN') ;
    dAssietteDeb := TFind.GetValue('CCT_ASSDEB') ;
    dAssietteFin := TFind.GetValue('CCT_ASSFIN') ;
    sCP :=  TFind.GetValue('CCT_CODEPOSTAL') ;
    sVille :=  TFind.GetValue('CCT_VILLE') ;
    sRegion := TFind.GetValue('CCT_REGION') ;
    sPays :=  TFind.GetValue('CCT_PAYS') ;

    //if (sPays = '') then // on prend le pays de l'établissement sur lequel est l'établissement
    //  sPays := strGetPaysEtabModele (strModele);

    //if bCherchePaysVide then // pour rechercher pays vide
    //  strPays := '';

    if strFlux = 'ACH' then  // flux de type achat  ACH
      dTxMt := TFind.GetValue('CCT_TXMTACHAT')
    else                   //flux de type vente  VEN
      dTxMt := TFind.GetValue('CCT_TXMTVENTE');
    dMini := TFind.GetValue('CCT_MTMINI');
    strFormule := TFind.GetValue('CCT_FORMULE');

    //if (strPays = '') then  // si pays non passé en param on prend le pays de l'établissement
    //  strPays := strGetPaysEtablissement (strEtablissement);


    if (strDatePiece >= strDateDeb) and (strDatePiece <= strDateFin)
    and (dMontant >= dAssietteDeb) and (dMontant <= dAssietteFin) then
    begin
      //On récupère le max date début concernée
      //TDateMaxDateDeb := TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays, strDatePiece, dMontant);
      if (strDateDeb = TDateMaxDateDeb) or (bCherchePaysVide) then
      begin
      //Recherche concernant le lieu de livraison
      if (strCriLocalisation = 'PAS') and (strPays = sPays) then // pas de recherche prévue sur lieu livraison au niveau catégorie
      begin
        TPays.AddChampSupValeur( 'MEM_MODELE', strModele , false);
        TPays.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
        TPays.AddChampSupValeur( 'MEM_TYPE',  strType, false);
        TPays.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
        TPays.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
        TPays.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
        TPays.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
        TPays.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
        TPays.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
        TPays.AddChampSupValeur( 'MEM_EXO', 'N' , false);
        TPays.AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
        iTrouve := iTrouve + 1;
        if iTrouve > 1 then // plus de 1 enr. trouvé
        begin
          TPays.AddChampSupValeur( 'MEM_MODELE', strModele , false);
          TPays.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
          TPays.AddChampSupValeur( 'MEM_TYPE',  strType, false);
          TPays.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
          TPays.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
          TPays.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
          TPays.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
          TPays.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
          TPays.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
          TPays.AddChampSupValeur( 'MEM_EXO', 'N' , false);
          TPays.AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
        end
      end
      else
      begin
        if (strCriLocalisation = 'CPO') and (strCP <> '') and (sCP <> '') then
        begin
          if (sCP = strCP) then
          begin
            TCodeP.AddChampSupValeur( 'MEM_MODELE', strModele , false);
            TCodeP.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
            TCodeP.AddChampSupValeur( 'MEM_TYPE',  strType, false);
            TCodeP.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
            TCodeP.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
            TCodeP.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
            TCodeP.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
            TCodeP.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
            TCodeP.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
            TCodeP.AddChampSupValeur( 'MEM_EXO', 'N' , false);
            TCodeP.AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
            iTrouve := iTrouve + 1;
            if iTrouve > 1 then // plus de 1 enr. trouvé
            begin
              TCodeP.AddChampSupValeur( 'MEM_MODELE', strModele , false);
              TCodeP.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
              TCodeP.AddChampSupValeur( 'MEM_TYPE',  strType, false);
              TCodeP.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
              TCodeP.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
              TCodeP.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
              TCodeP.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
              TCodeP.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
              TCodeP.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
              TCodeP.AddChampSupValeur( 'MEM_EXO', 'N' , false);
              TCodeP.AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
            end
          end
        end
        else
        begin
          if ((strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO'))  and (strVille <> '') and (sVille <> '') then
          begin
            if (sVille = strVille) then
            begin
              TVille.AddChampSupValeur( 'MEM_MODELE', strModele , false);
              TVille.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
              TVille.AddChampSupValeur( 'MEM_TYPE',  strType, false);
              TVille.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
              TVille.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
              TVille.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
              TVille.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
              TVille.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
              TVille.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
              TVille.AddChampSupValeur( 'MEM_EXO', 'N' , false);
              TVille.AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
              iTrouve := iTrouve + 1;
              if iTrouve > 1  then // plus de 1 enr. trouvé
              begin
                TVille.AddChampSupValeur( 'MEM_MODELE', strModele , false);
                TVille.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
                TVille.AddChampSupValeur( 'MEM_TYPE',  strType, false);
                TVille.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
                TVille.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
                TVille.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
                TVille.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
                TVille.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
                TVille.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
                TVille.AddChampSupValeur( 'MEM_EXO', 'N' , false);
                TVille.AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
              end
            end
          end
          else
          begin
            if ((strCriLocalisation = 'REG') or (strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO')) and (strRegion <> '') and (sRegion <> '') then
            begin
              if (sRegion = strRegion) then
              begin
                TRegion.AddChampSupValeur( 'MEM_MODELE', strModele , false);
                TRegion.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
                TRegion.AddChampSupValeur( 'MEM_TYPE',  strType, false);
                TRegion.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
                TRegion.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
                TRegion.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
                TRegion.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
                TRegion.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
                TRegion.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
                TRegion.AddChampSupValeur( 'MEM_EXO', 'N' , false);
                TRegion.AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
                iTrouve := iTrouve + 1;
                if iTrouve > 1  then // plus de 1 enr. trouvé
                begin
                  TRegion.AddChampSupValeur( 'MEM_MODELE', strModele , false);
                  TRegion.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
                  TRegion.AddChampSupValeur( 'MEM_TYPE',  strType, false);
                  TRegion.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
                  TRegion.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
                  TRegion.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
                  TRegion.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
                  TRegion.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
                  TRegion.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
                  TRegion.AddChampSupValeur( 'MEM_EXO', 'N' , false);
                  TRegion.AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
                end
              end
            end
            else
            if ((strCriLocalisation = 'REG') or (strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO')) and (strPays <> '') then
              begin
                if (sPays = strPays) then
                begin
                  TPays.AddChampSupValeur( 'MEM_MODELE', strModele , false);
                  TPays.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
                  TPays.AddChampSupValeur( 'MEM_TYPE',  strType, false);
                  TPays.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
                  TPays.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
                  TPays.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
                  TPays.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
                  TPays.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
                  TPays.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
                  TPays.AddChampSupValeur( 'MEM_EXO', 'N' , false);
                  TPays.AddChampSupValeur( 'MEM_REPONSE', 'OK' , false);
                  iTrouve := iTrouve + 1;
                  if iTrouve > 1  then // plus de 1 enr. trouvé
                  begin
                    TPays.AddChampSupValeur( 'MEM_MODELE', strModele , false);
                    TPays.AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
                    TPays.AddChampSupValeur( 'MEM_TYPE',  strType, false);
                    TPays.AddChampSupValeur( 'MEM_REGIME', strRegime, false);
                    TPays.AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
                    TPays.AddChampSupValeur( 'MEM_MTMINI', dMini , false);
                    TPays.AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
                    TPays.AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
                    TPays.AddChampSupValeur( 'MEM_QUALIFIANT', strQualifiant , false);
                    TPays.AddChampSupValeur( 'MEM_EXO', 'N' , false);
                    TPays.AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
                  end;

                end
              end
            end
        end
      end;
    end;
    end;
    TFind:=FTOBPARAMETRESTAXES.FindNext(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  end;
end;
//On alimente la ligne réponse trouvée précédemment : on prend en compte l'élément le plus fin
if TCodeP.Detail.count - 1 > 0 then
  Result := TCodeP
else
  if TVille.Detail.count - 1 > 0 then
    Result := TVille
  else
    if TRegion.Detail.count - 1 > 0 then
      Result := TRegion
    else
      Result := TPays;
finally
  FreeAndNil (TobLigneReponse0);
  FreeAndNil (TPaysD);
  FreeAndNil (TCodePBis);
  FreeAndNil (TVilleD);
  FreeAndNil (TRegionD);
end;
end;

function TInternationalTax.GetNomChampParam(Num : integer) : string;
begin
  if Num = 0 then Result := ''
  else if Num = 1 then Result := 'MEM_MODELE'
  else if Num = 2 then Result := 'MEM_CATEGORIE'
  else if Num = 3 then Result := 'MEM_TYPE'
  else if Num = 4 then Result := 'MEM_REGIME'
  else if Num = 5 then Result := 'MEM_ETABLISSEMENT'
  else if Num = 6 then Result := 'MEM_FLUX'
  else if Num = 7 then Result := 'MEM_CPLIV'
  else if Num = 8 then Result := 'MEM_VILLELIV'
  else if Num = 9 then Result := 'MEM_REGIONLIV'
  else if Num = 10 then Result := 'MEM_PAYSLIV'
  else if Num = 11 then Result := 'MEM_DATEPIECE'
  else if Num = 12 then Result := 'MEM_QUANTITELIGNE'
  else if Num = 13 then Result := 'MEM_MONTANTLIGNE';
end;

procedure TInternationalTax.SetParamInTobquestion(TobAlimQuestion : TOB; NumChamp : integer; stValeur : variant);
begin
  TobAlimQuestion.PutValue(GetNomChampParam(NumChamp), stValeur);
end;

procedure  TInternationalTax.tCreeLigneQuestion (TobAlimQuestion : Tob; strModele, strCategorie, strType, strRegime, strEtablissement, strFlux, strCP, strVille, strRegion, strPays  : string ; tDatepiece : TDatetime; iQuantiteLigne : integer; dMontantLigne : double );
var Num : integer;

  procedure AjouteChamp;
  begin
    if not TobAlimQuestion.FieldExists(GetNomChampParam(Num)) then
      TobAlimQuestion.AddChampSup(GetNomChampParam(Num), False);
  end;

begin
  Num := 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, StrModele);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strCategorie);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strType);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strRegime);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strEtablissement);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strFlux);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strCP);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strVille);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strRegion);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, strPays);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, tDatepiece);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, iQuantiteLigne);
  Num := Num + 1; AjouteChamp; SetParamInTobquestion(TobAlimQuestion, Num, dMontantLigne);
end;

function TInternationalTax.strGetPaysEtablissement(strEtablissement: string): string;
var strSql : string;
OQuery : TQuery;
begin
  try
    if strEtablissement = '' then exit;
    strSQL := 'SELECT ET_PAYS FROM ETABLISS WHERE ET_ETABLISSEMENT = "'+ strEtablissement+ '"';
    OQuery := OpenSql (strSQL, TRUE);
    try
    if not OQuery.Eof then
      result := OQuery.Fields[0].AsString;
    finally
      Ferme (OQuery);
    end
  except
    result := '';
  end;
end;

function TInternationalTax.strGetPaysEtabModele(strModele : string): string;
var strSql : string;
OQuery : TQuery;
begin
  try
    if strModele = '' then exit;
    strSQL := 'SELECT ET_PAYS FROM ETABLISS WHERE ET_CODE_MODELETAXE = "'+ strModele + '"';
    OQuery := OpenSql (strSQL, TRUE);
    try
    if not OQuery.Eof then
      result := OQuery.Fields[0].AsString;
    finally
      Ferme (OQuery);
    end
  except
    result := '';
  end;
end;

function TInternationalTax.strGetTypeQualifiant(strType: string): string;
var strSql : string;
OQuery : TQuery;
begin
  try
    if strType = '' then exit;
    strSQL := 'SELECT TYP_QUALIFIANT FROM YTYPETAUX WHERE TYP_CODE = "'+ strType + '"';
    OQuery := OpenSql (strSQL, TRUE);
    try
    if not OQuery.Eof then
      result := OQuery.Fields[0].AsString;
    finally
      Ferme (OQuery);
    end
  except
    result := '';
  end;
end;

function TInternationalTax.tstrGetLibelleTaxesModele(strModele: string): string;
var strCat : string ;
  strSQL, St : string;
  OQuery : TQuery ;
begin
  //Fontion permettant de récupérer les libellés des catégories d'un modèle
  strCat  := '' ;

  result := strCat;
  strSQL := 'SELECT "TX1", YMODELETAXE.YMT_LIBCAT1 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT1 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT1 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
  strSql := strSQL + 'AND "TX1" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX1" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX2", YMODELETAXE.YMT_LIBCAT2 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT2 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT2 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
  strSql := strSQL + 'AND "TX2" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX2" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX3", YMODELETAXE.YMT_LIBCAT3 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT3 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT3 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
  strSql := strSQL + 'AND "TX3" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX3" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX4", YMODELETAXE.YMT_LIBCAT4 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT4 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT4 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
  strSql := strSQL + 'AND "TX4" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX4" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';
  strSQL := strSQL + ' UNION ';
  strSQL := strSQL + 'SELECT "TX5", YMODELETAXE.YMT_LIBCAT5 FROM YMODELETAXE WHERE YMODELETAXE.YMT_LIBCAT5 IS NOT NULL AND YMODELETAXE.YMT_LIBCAT5 <> '' AND YMODELETAXE.YMT_CODE = "'+strModele+'"';
  strSql := strSQL + 'AND "TX5" IN (SELECT YCY_CODECAT FROM YMODELECATTYP WHERE YCY_TYPGERE = "X") AND "TX5" IN (SELECT YCR_CODECAT FROM YMODELECATREG WHERE YCR_REGGERE = "X")';

  OQuery := OpenSQL( strSql , false);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[1].AsString;
    strCat := strCat + St + ';';
    OQuery.next;
  end;
  result := strCat;
  finally
    Ferme(OQuery);
  end;
end;

function TInternationalTax.bChercheSiRien(strModele, strCategorie, strType, strRegime: string): boolean;
var
  TFind : Tob;

begin
//Recherche dans TOB si infos importantes existantes

  result := true; //initialisation à "vrai" de la fonction

  TFind := nil;

  //Recherche enregistrement dans tob de paramétrage
  TFind := FTOBPARAMETRESTAXES.FindFirst(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  if TFind <> nil then
    result := false;

end;

function TInternationalTax.iChercheTobParametrestaxes(strModele, strCategorie, strType, strRegime: string; strDatePiece: TDatetime; dMontant: double; strCriLocalisation, strCP, strVille, strRegion, strPays: string; bCherchePaysVide : boolean): integer;
var
iTrouve : integer;
TFind : Tob;
strDateDeb, strDateFin : tDatetime;
dAssietteDeb, dAssietteFin : double;
sCP, sVille, sRegion, sPays : string;
TDateMaxDateDeb : tDateTime;
begin

  result := 0; //initialisation à 0

  TFind := nil;

  iTrouve := 0; // initialisation nombre enr. trouvé

  TDateMaxDateDeb := TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays, strDatePiece, dMontant, true);

  //Recherche enregistrement dans tob de paramétrage
  TFind := FTOBPARAMETRESTAXES.FindFirst(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  While TFind<>Nil do
  BEGIN
    strDateDeb := TFind.GetValue('CCT_PERDEB') ;
    strDateFin := TFind.GetValue('CCT_PERFIN') ;
    dAssietteDeb := TFind.GetValue('CCT_ASSDEB') ;
    dAssietteFin := TFind.GetValue('CCT_ASSFIN') ;
    sCP :=  TFind.GetValue('CCT_CODEPOSTAL') ;
    sVille :=  TFind.GetValue('CCT_VILLE') ;
    sRegion := TFind.GetValue('CCT_REGION') ;
    sPays :=  TFind.GetValue('CCT_PAYS') ;

    //if bCherchePaysVide then
    //  strPays := '';

    if (strDatePiece >= strDateDeb) and (strDatePiece <= strDateFin)
    and (dMontant >= dAssietteDeb) and (dMontant <= dAssietteFin) then
    begin
      //On récupère le max date début concernée
      //TDateMaxDateDeb := TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays, strDatePiece, dMontant);
      if (strDateDeb = TDateMaxDateDeb) or (bCherchePaysVide) then
      begin
      //Recherche concernant le lieu de livraison
      if (strCriLocalisation = 'PAS') and ((strPays = sPays)) then // pas de recherche prévue sur lieu livraison au niveau catégorie
       begin
        iTrouve := iTrouve + 1;
      end
      else
      begin
        if (strCriLocalisation = 'CPO') and (strCP <> '') then
        begin
          if (sCP = strCP) or (sVille = strVille) or (sRegion = strRegion) or (sPays = strPays) then
          begin
            iTrouve := iTrouve + 1;
          end
        end
        else
        begin
          if ((strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO'))  and (strVille <> '') then
          begin
            if (sVille = strVille)  or (sRegion = strRegion) or (sPays = strPays) then
            begin
              iTrouve := iTrouve + 1;
            end
          end
          else
          begin
            if ((strCriLocalisation = 'REG') or (strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO')) and (strRegion <> '') then
            begin
              if (sRegion = strRegion)or (sPays = strPays) then
              begin
                iTrouve := iTrouve + 1;
              end
            end
            else
            if ((strCriLocalisation = 'REG') or (strCriLocalisation = 'VIL') or (strCriLocalisation = 'CPO')) and (strPays <> '') then
              begin
                if (sPays = strPays) then
                begin
                  iTrouve := iTrouve + 1;
                end
              end
            end
        end
      end;
    end;
    end;
    TFind:=FTOBPARAMETRESTAXES.FindNext(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
   END ;
   result := iTrouve;
end;

function TInternationalTax.TobChercheEnReduisantUnElement(strModele, strCategorie, strType, strRegime: string; strDatePiece: TDatetime; dMontant: double; strCriLocalisation, strCP, strVille, strRegion, strPays, strTaxesIncluses, strFlux, strElementEnleve: string): TOB;
var
iTrouve, k : integer;
strFormule : string;
TFind, TStockageM, TStockageF : Tob;
strDateDeb, strDateFin : tDatetime;
dAssietteDeb, dAssietteFin : double;
sCP, sVille, sRegion, sPays : string;
dTxMt, dMini : double;
TDateMaxDateDeb : TDateTime;
begin


  TFind := nil;

  iTrouve := 0; // initialisation nombre enr. trouvés

  //Tob de stockage mère et fille
  TStockageM := TOB.Create('TobTravailM', nil, -1) ;
  //TStockageF := TOB.Create('TobTravailF', TStockageM, -1) ;
  k := 0;

  if (strElementEnleve = 'PAYS') or (strElementEnleve = 'ASSIETTE') then
    TDateMaxDateDeb := TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays, strDatePiece, dMontant, false);

  //Recherche enregistrement dans tob de paramétrage
  TFind := FTOBPARAMETRESTAXES.FindFirst(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  While TFind<>Nil do
  BEGIN
    strDateDeb := TFind.GetValue('CCT_PERDEB') ;
    strDateFin := TFind.GetValue('CCT_PERFIN') ;
    dAssietteDeb := TFind.GetValue('CCT_ASSDEB') ;
    dAssietteFin := TFind.GetValue('CCT_ASSFIN') ;
    sCP :=  TFind.GetValue('CCT_CODEPOSTAL') ;
    sVille :=  TFind.GetValue('CCT_VILLE') ;
    sRegion := TFind.GetValue('CCT_REGION') ;
    sPays :=  TFind.GetValue('CCT_PAYS') ;

    if strFlux = 'ACH' then  // flux de type achat  ACH
      dTxMt := TFind.GetValue('CCT_TXMTACHAT')
    else                   //flux de type vente  VEN
      dTxMt := TFind.GetValue('CCT_TXMTVENTE');
    dMini := TFind.GetValue('CCT_MTMINI');
    strFormule := TFind.GetValue('CCT_FORMULE');


    //si le pays est écarté
    if (strDatePiece >= strDateDeb) and (strDatePiece <= strDateFin)
    and (dMontant >= dAssietteDeb) and (dMontant <= dAssietteFin)
    and (strDateDeb = TDateMaxDateDeb)
    and (strElementEnleve = 'PAYS') then
    begin
      TStockageF := TOB.Create('TobTravailF', TStockageM, -1) ;
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_MTMINI', dMini , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', strGetTypeQualifiant (strType) , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_EXO', 'N' , false);
      if k = 0 then
        TStockageM.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'NOK' , false)
      else
        TStockageM.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
      k:= k + 1;
    end;

    //si l'assiette est écartée
    if (strDatePiece >= strDateDeb) and (strDatePiece <= strDateFin)
    and (strDateDeb = TDateMaxDateDeb)
    and (strElementEnleve = 'ASSIETTE') then
    begin
      TStockageF := TOB.Create('TobTravailF', TStockageM, -1) ;
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_MTMINI', dMini , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', strGetTypeQualifiant (strType) , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_EXO', 'N' , false);
      if k = 0 then
        TStockageM.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'NOK' , false)
      else
        TStockageM.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
      k:= k + 1;
    end;

    //si la période est écartée
    if (strElementEnleve = 'PERIODE') then
    begin
      TStockageF := TOB.Create('TobTravailF', TStockageM, -1) ;
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_MODELE', strModele , false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_CATEGORIE', strCategorie, false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_TYPE',  strType, false);
      TStockageM.Detail[K].AddChampSupValeur( 'MEM_REGIME', strRegime, false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_TXMTACHAT', dTxMt , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_MTMINI', dMini , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_TAXESINCLUSES', strTaxesIncluses , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_FORMULE', strFormule , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_QUALIFIANT', strGetTypeQualifiant (strType) , false);
      TStockageM.Detail[k].AddChampSupValeur( 'MEM_EXO', 'N' , false);
      if k = 0 then
        TStockageM.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'NOK' , false)
      else
        TStockageM.Detail[k].AddChampSupValeur( 'MEM_REPONSE', 'PAR' , false);
      k:= k + 1;
   end;


    TFind:=FTOBPARAMETRESTAXES.FindNext(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
   END ;
   result := TStockageM;
end;

function TInternationalTax.TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays: string; strDatePiece : TDateTime; dMontant : double; bAvecMontant : boolean) : TDateTime;
var
   strSQL  : string;
   OQuery : TQuery ;
   strMontant : string;
begin
   result := iDate1900 ;
   strMontant := FloatToStr (dMontant);
   strSQL := 'SELECT TOP 1 CCT_PERDEB FROM YDETAILTAXES WHERE ';
   strSQL := strSQL + 'CCT_CODEMODELE = "'+strModele+'"';
   strSQL := strSQL + ' AND CCT_CODECAT = "'+strCategorie+'"';
   strSQL := strSQL + ' AND CCT_CODETYP = "'+strType+'"';
   strSQL := strSQL + ' AND CCT_CODEREG = "'+strRegime+'"';
   strSQL := strSQL + ' AND CCT_PERDEB >= "'+ USDateTime(strDatePiece)+ '"';
   strSQL := strSQL + ' AND CCT_PAYS = "' + strPays + '"';
   if bAvecMontant then
   begin
    strSQL := strSQL + ' AND CCT_ASSDEB <= '+  strMontant;
    strSQL := strSQL + ' AND CCT_ASSFIN >= '+  strMontant;
   end;
   strSQL := strSQL + ' ORDER BY CCT_PERDEB DESC';


   OQuery := OpenSql (strSQL, TRUE);
   try
   if not OQuery.Eof then
      result := OQuery.Fields[0].AsDateTime;
   finally
    Ferme(OQuery);
   end;

end;

function TInternationalTax.iChercheInfosExoneration(strModele, strCategorie, strType, strRegime: string; strDatePiece: TDatetime;
         dMontant: double; strCriLocalisation, strCP, strVille, strRegion,strPays: string): integer;
//Ici assiette et période non prise en compte
var
iTrouve : integer;
TFind : Tob;
sCP, sVille, sRegion, sPays : string;
TDateMaxDateDeb : TDateTime;
TDateDeb, TDateFin : tDatetime;
dAssietteDeb, dAssietteFin : double;
begin

  result := 0;

  TFind := nil;

  iTrouve := 0;

  TDateMaxDateDeb := TDateChercheFourchette (strModele, strCategorie, strType, strRegime, strPays, strDatePiece, dMontant, true);

  TFind := FTOBPARAMETRESTAXES.FindFirst(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
  While TFind<>Nil do
  BEGIN
    TDateDeb := TFind.GetValue('CCT_PERDEB') ;
    TDateFin := TFind.GetValue('CCT_PERFIN') ;
    dAssietteDeb := TFind.GetValue('CCT_ASSDEB') ;
    dAssietteFin := TFind.GetValue('CCT_ASSFIN') ;
    sCP :=  TFind.GetValue('CCT_CODEPOSTAL') ;
    sVille :=  TFind.GetValue('CCT_VILLE') ;
    sRegion := TFind.GetValue('CCT_REGION') ;
    sPays :=  TFind.GetValue('CCT_PAYS') ;

    if (strDatePiece >= TDateDeb) and (strDatePiece <= TDateFin)
    and (dMontant >= dAssietteDeb) and (dMontant <= dAssietteFin) then
    begin
      iTrouve := 1;
      if (TDateDeb = TDateMaxDateDeb) and (strPays = sPays) then
      begin
        iTrouve := 2;
      end;
    end;
    TFind:=FTOBPARAMETRESTAXES.FindNext(['CCT_CODEMODELE', 'CCT_CODECAT', 'CCT_CODETYP', 'CCT_CODEREG'],[strModele, strCategorie, strType, strRegime], True);
   END ;
   result := iTrouve;
end;

function TInternationalTax.tstrGetLibelleTypesTaxesModele(strModele, strCategorie: string): String;
var
   strType, strSQL, St : string;
   OQuery : TQuery ;
begin
  strType  := '' ;

  result := strType;

  strSQL := 'SELECT YMODELECATTYP.YCY_CODETYP, YTYPETAUX.TYP_LIBELLE FROM YMODELECATTYP INNER JOIN YTYPETAUX ON YMODELECATTYP.YCY_CODETYP = YTYPETAUX.TYP_CODE WHERE YMODELECATTYP.YCY_TYPGERE = "X" AND YMODELECATTYP.YCY_CODECAT  = "'+strCategorie+'" AND YMODELECATTYP.YCY_CODEMODELE = "'+strModele+'"';

  OQuery := OpenSQL( strSql , false);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[1].AsString;
    strType := strType + St + ';';
    OQuery.next;
  end;
  result := strType;
  finally
    Ferme(OQuery);
  end;

end;

function TInternationalTax.tstrGetLibelleRegimesTaxesModele(strModele, strCategorie: string): String;
var
   strRegime, strSQL, St : string;
   OQuery : TQuery ;
begin
  strRegime  := '' ;

  result := strRegime;

  strSQL := 'SELECT YMODELECATREG.YCR_CODEREG, CHOIXCOD.CC_LIBELLE FROM YMODELECATREG INNER JOIN CHOIXCOD ON YMODELECATREG.YCR_CODEREG = CHOIXCOD.CC_CODE WHERE YMODELECATREG.YCR_REGGERE = "X" AND CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_CODECAT  = "'+strCategorie+'" AND YMODELECATREG.YCR_CODEMODELE = "'+strModele+'"';

  OQuery := OpenSQL( strSql , false);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[1].AsString;
    strRegime := strRegime + St + ';';
    OQuery.next;
  end;
  result := strRegime;
  finally
    Ferme(OQuery);
  end;

end;


function TInternationalTax.tstrGetCodeTypesTaxesModele(strModele, strCategorie: string): String;
var
   strType, strSQL, St : string;
   OQuery : TQuery ;
begin
  strType  := '' ;

  result := strType;

  strSQL := 'SELECT YMODELECATTYP.YCY_CODETYP, YTYPETAUX.TYP_LIBELLE FROM YMODELECATTYP INNER JOIN YTYPETAUX ON YMODELECATTYP.YCY_CODETYP = YTYPETAUX.TYP_CODE WHERE YMODELECATTYP.YCY_TYPGERE = "X" AND YMODELECATTYP.YCY_CODECAT  = "'+strCategorie+'" AND YMODELECATTYP.YCY_CODEMODELE = "'+strModele+'"';

  OQuery := OpenSQL( strSql , false);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    strType := strType + St + ';';
    OQuery.next;
  end;
  result := strType;
  finally
    Ferme(OQuery);
  end;

end;

function TInternationalTax.tstrGetCodeRegimesTaxesModele(strModele, strCategorie: string): String;
var
   strRegime, strSQL, St : string;
   OQuery : TQuery ;
begin
  strRegime  := '' ;

  result := strRegime;

  strSQL := 'SELECT YMODELECATREG.YCR_CODEREG, CHOIXCOD.CC_LIBELLE FROM YMODELECATREG INNER JOIN CHOIXCOD ON YMODELECATREG.YCR_CODEREG = CHOIXCOD.CC_CODE WHERE YMODELECATREG.YCR_REGGERE = "X" AND CHOIXCOD.CC_TYPE = "RTV" AND YMODELECATREG.YCR_CODECAT  = "'+strCategorie+'" AND YMODELECATREG.YCR_CODEMODELE = "'+strModele+'"';

  OQuery := OpenSQL( strSql , false);
  try
  While not OQuery.Eof do
  begin
    St := OQuery.Fields[0].AsString;
    strRegime := strRegime + St + ';';
    OQuery.next;
  end;
  result := strRegime;
  finally
    Ferme(OQuery);
  end;

end;

end.
