{
********************************************************************************

Objet ........ : Cet objet gère la table ITRTPARAM.

                  Cet Objet permet de créer, modifier, lire, supprimer un paramétrage
                  quelconque par utilisateur, (voir groupe utilisateurs Pour l'instant non géré)

                  Actuellement, Cet objet a été utilisé pour le pointage automatique pour
                  mémoriser les règles d'accorchage    (CPA)    (Uniquement)

--------------------------------------------------------------------------------
    Version   | Date   | Qui |   Commentaires
--------------------------------------------------------------------------------
               05/05/08  3P8  Création de l'unité
               03/12/09  3p8  ReFactoring de l'unité CPOBJREGLEACCROCHAGE en
                              ITRTPARAM
               07/12/09  3p8  Changement de méthode pour acceder au paramètre par défaut
********************************************************************************
 Valeur du champ  ITT_TYPPAR
 ---------------------------
    CPA : Mémorisation des règles d'accrochage
    CPB : Memorisation du paramétrage des souches bancaires 
    CPM : Pointage Manuel
    CPT : Pointage Manuel Tréso
********************************************************************************
}
Unit CPOBJITRTPARAM ;

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
     TActionTable = (ttaRIEN, ttaCREAT, ttaMODIF, ttaDELETE);

{------------------------------------------------------------------------------}
 type
     TITRTParam = Class (TObject)
     private
        Action            : TActionTable ;
        bLibKeyFg         : Boolean ;

        TOBITRTPARAM      : TOB ;
        TOBDEFAIRE        : TOB ;
        TDATAList         : TOB ;
        TDATADEFAIREList  : TOB ;


        Procedure UpdateDataBase() ;
        Procedure DeleteDataBase() ;

        Function  IsAllControler()  : Boolean ;


        //* Transforme au format XML les données des RIB pour les règles d'accrochage

        Procedure  Rempli_CPA_ITT_DATA() ;

       //* Constitue la TOBDATAList des RIB pour les règles d'accrochage

        Procedure  Rempli_CPA_TDATAList(bDefaireFg : Boolean = False) ;

     published
        Constructor Create     () ;

        Destructor  Destroy    () ; override ;

        Function  AffecterCle () : Boolean;

        Function  Defaire      () : Boolean;
        Function  DeleteBD     () : Boolean;

        function LoadData(theTOBITRTPARAM: TOB): Boolean;

        Function  Exists       () : Boolean ;

        Function  UpdateBD     () : Boolean ;

        Function  GetAction    () : TActionTable;
        Procedure SetAction    (theAction : TActionTable) ;

        Function  GetKey()                : String ;
        Procedure SetLikKey(theBLibKeyFg  : Boolean) ;

        Function  GetType()               : String ;
        Procedure SetType(theType         : String) ;

        Function  GetUser()               : String ;
        Procedure SetUser(theUser         : String) ;

        Function  GetUserGroupe()         : String ;
        Procedure SetUserGroupe(theGroupe : String) ;

        Function  GetCle()                : Integer ;
        Procedure SetCle(theCle           : Integer) ;

        Function  GetVersion()            : Integer ;
        Procedure SetVersion(theVersion   : Integer) ;

        Function  GetLibelle()            : String ;
        Procedure SetLibelle(theLibelle   : String) ;

        Function  GetDefaut()             : Boolean ;
        Procedure SetDefaut(theDefaultFg  : Boolean) ;

        Function  GetRegleAccrochage(theTOBList : TOB) : Boolean;
        Function  SetRegleAccrochage(theTOBList : TOB) : Boolean;

        {b FP 14/12/2009 Paramétrage de la grille - FD 516}
        Function  GetData()      : TOB;
     public
        Procedure SetData(TOBData: TOB) ; overload ;
        Procedure SetData(theXMLChaine : String); overload ;
        {e FP 14/12/2009}
     end ;
{------------------------------------------------------------------------------}
 type
     TITRTParamList = Class (TObject)
     private
        Liste               : TStringList ;

        EnCours             : TITRTParam ;
        bSearchByLibelleFg  : Boolean ;

        Function ConstituerCle(theITrtParam : TITRTParam) : String  ;

        Function IsExisteParDefaut() : Boolean ;

     public
        Function  GetParDefaut()                          : TITRTParam ; overload ;
        Function  GetParDefaut(theQuelType : Integer)     : TITRTParam ; overload ;
     published

        Constructor Create          () ;
        Destructor  Destroy         () ; override ;

        Function LoadData(theType : String ; theSearchOnLibelleFg : Boolean = False) : Boolean ;

        Function  Add               (theITrtParam : TITRTParam) : Boolean ;
        Function  remove            (theITrtParam : TITRTParam) : Boolean ;
        Function  Count             : Integer  ;
        Function  Detail            (theOccurenceNumber : Integer) : TITRTParam ;
        Function  DeleteBD          () : Boolean;

        Function  GetEnCours()                            : TITRTParam ;
        Function  Selectionner(theSearchKey : String)     : TITRTParam ;

        Function  UpdateAndDeleteBD () : Boolean ;
        Function  UpdateBD          () : Boolean ;
 end ;

function IsExisteListeParametrable(theParameterType : String) : Boolean  ;

{------------------------------------------------------------------------------}
function CreerListeParametrablePourPointageManuel(theParameterType : String) : Boolean  ;

Implementation

uses DB,
    HStatus ;

var
  MSG    : String ;
  ENTETE : String ;


{------------------------------------------------------------------------------}
Function  FormatTag(theTagName : String ; theValue : String) : String ;
BEGIN
 result := '<$' + theTagName +'$>' + theValue+ '</$' + theTagName + '$>' ;
END ;

{------------------------------------------------------------------------------}
Function RechercherParameterXML(theTagName : String ; theParamList : String) : String ;
var tagDebut, tagFin : String ;
    posDebut, posFin : Integer ;
    lengthDebut, longueur, Demarre : Integer ;
BEGIN
  tagDebut := '<$' + theTagName + '$>' ;
  tagFin   := '</$' + theTagName + '$>' ;

  lengthDebut := Length(tagDebut) ;
  posDebut := Pos(tagDebut, theParamList) ;
  posFin := Pos(tagFin, theParamList) ;

  Demarre := posDebut + lengthDebut ;

  longueur := posFin - Demarre;
  result := Copy(theParamList, Demarre, longueur) ;
END ;

{------------------------------------------------------------------------------}
Function RechercherParameterXMLSimple(theTagName : String ; theParamList : String) : String ;
var tagDebut, tagFin : String ;
    posDebut, posFin : Integer ;
    lengthDebut, longueur, Demarre : Integer ;
BEGIN
  tagDebut := '<' + theTagName + '>' ;
  tagFin   := '</' + theTagName + '>' ;

  lengthDebut := Length(tagDebut) ;
  posDebut := Pos(tagDebut, theParamList) ;
  posFin := Pos(tagFin, theParamList) ;

  Demarre := posDebut + lengthDebut ;

  longueur := posFin - Demarre;
  result := Copy(theParamList, Demarre, longueur) ;
END ;

{------------------------------------------------------------------------------}
function IsExisteListeParametrable(theParameterType : String) : Boolean  ;
var
  SQL : String ;
BEGIN

  SQL := 'SELECT ITT_CLE FROM ITRTPARAM WHERE '
      +  ' ITT_CLE = 1 AND ITT_DEFAUT = "X" AND ITT_TYPPAR = "' + theParameterType + '"'  ;

  if ExisteSQL(SQL) then result := True else result := False ;
END ;

{------------------------------------------------------------------------------}
Procedure CreerListe(theTOBDATAList : TOB ; theXmlParam : String)  ;
var
  SFIELDNAME1   : String ;
  SFIELDNAME2   : String ;
  STITRE        : String ;
  SALIAS        : String ;
  GFFUNCTION    : Integer ;
  TITLEALIGN    : Integer ;
  TEXTALIGN     : Integer;
  CLBACKCOLOR   : Integer;
  BFIELDMANUAL  : String ;
  DTDISPLAYTYPE : Integer ;
  IWIDTH        : Integer;
  ILEFT         : Integer;
  BREQUIRED     : String ;
  BSEPMILLIER   : String ;
  BVISIBLE      : String ;
  BEMPTYIFNULL  : String ;
  INBDEC        : Integer ;
  BDESCRIPTION  : String ;
  wTOB          : TOB ;
begin
  SFIELDNAME1   := RechercherParameterXMLSimple('SFIELDNAME1', theXmlParam) ;
  SFIELDNAME2   := RechercherParameterXMLSimple('SFIELDNAME2', theXmlParam) ;
  STITRE        := RechercherParameterXMLSimple('STITRE', theXmlParam) ;
  SALIAS        := RechercherParameterXMLSimple('SALIAS', theXmlParam) ;
  GFFUNCTION    := StrToInt(RechercherParameterXMLSimple('GFFUNCTION', theXmlParam)) ;
  TITLEALIGN    := StrToInt(RechercherParameterXMLSimple('TITLEALIGN', theXmlParam)) ;
  TEXTALIGN     := StrToInt(RechercherParameterXMLSimple('TEXTALIGN', theXmlParam));
  CLBACKCOLOR   := StrToInt(RechercherParameterXMLSimple('CLBACKCOLOR', theXmlParam));
  BFIELDMANUAL  := RechercherParameterXMLSimple('BFIELDMANUAL', theXmlParam) ;
  DTDISPLAYTYPE := StrToInt(RechercherParameterXMLSimple('DTDISPLAYTYPE', theXmlParam)) ;
  IWIDTH        := StrToInt(RechercherParameterXMLSimple('IWIDTH', theXmlParam));
  ILEFT         := StrToInt(RechercherParameterXMLSimple('ILEFT', theXmlParam));
  BREQUIRED     := RechercherParameterXMLSimple('BREQUIRED', theXmlParam) ;
  BSEPMILLIER   := RechercherParameterXMLSimple('BSEPMILLIER', theXmlParam) ;
  BVISIBLE      := RechercherParameterXMLSimple('BVISIBLE', theXmlParam) ;
  BEMPTYIFNULL  := RechercherParameterXMLSimple('BEMPTYIFNULL', theXmlParam) ;
  INBDEC        := StrToInt(RechercherParameterXMLSimple('INBDEC', theXmlParam)) ;
  BDESCRIPTION  := RechercherParameterXMLSimple('BDESCRIPTION', theXmlParam) ;

  wTOB := TOB.Create('$COL', theTOBDATAList, -1) ;
  wTOB.addChampSupValeur('SFIELDNAME1',   SFIELDNAME1,    False) ;
  wTOB.addChampSupValeur('SFIELDNAME2',   SFIELDNAME2,    False) ;
  wTOB.addChampSupValeur('STITRE',        STITRE,         False) ;
  wTOB.addChampSupValeur('SALIAS',        SALIAS,         False) ;
  wTOB.addChampSupValeur('GFFUNCTION',    GFFUNCTION,     False) ;
  wTOB.addChampSupValeur('TITLEALIGN',    TITLEALIGN,     False) ;
  wTOB.addChampSupValeur('TEXTALIGN',     TEXTALIGN,      False) ;
  wTOB.addChampSupValeur('CLBACKCOLOR',   CLBACKCOLOR,    False) ;
  wTOB.addChampSupValeur('BFIELDMANUAL',  BFIELDMANUAL,   False) ;
  wTOB.addChampSupValeur('DTDISPLAYTYPE', DTDISPLAYTYPE,  False) ;
  wTOB.addChampSupValeur('IWIDTH',        IWIDTH,         False) ;
  wTOB.addChampSupValeur('ILEFT',         ILEFT,          False) ;
  wTOB.addChampSupValeur('BREQUIRED',     BREQUIRED,      False) ;
  wTOB.addChampSupValeur('BSEPMILLIER',   BSEPMILLIER,    False) ;
  wTOB.addChampSupValeur('BVISIBLE',      BVISIBLE,       False) ;
  wTOB.addChampSupValeur('BEMPTYIFNULL',  BEMPTYIFNULL,   False) ;
  wTOB.addChampSupValeur('INBDEC',        INBDEC,         False) ;
  wTOB.addChampSupValeur('BDESCRIPTION',  BDESCRIPTION,   False) ;
end ;

{------------------------------------------------------------------------------}
function ConstruireListePointageManuelStandard() : TOB  ;
var
  XmlParam : String ;
BEGIN
  result := TOB.Create('$COLS', Nil, -1) ;
  result.AddChampSupValeur('TABLE1', 'ECRITURE', False);
  result.AddChampSupValeur('TABLE2', 'EEXBQLIG', False);

  XmlParam :=
          '<$COL><SFIELDNAME1>E_DATECOMPTABLE</SFIELDNAME1><SFIELDNAME2>CEL_DATEOPERATION</SFIELDNAME2>'
        + '<STITRE>Date Comptable</STITRE><SALIAS>CLE_DATECOMPTABLE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION><TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>2</TEXTALIGN><CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL><DTDISPLAYTYPE>1</DTDISPLAYTYPE><IWIDTH>70</IWIDTH>'
        + '<ILEFT>0</ILEFT><BREQUIRED>X</BREQUIRED><BSEPMILLIER>-</BSEPMILLIER><BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL><INBDEC>0</INBDEC><BDESCRIPTION>-</BDESCRIPTION></$COL>';

  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL><SFIELDNAME1>E_NUMEROPIECE</SFIELDNAME1><SFIELDNAME2>CEL_NUMRELEVE</SFIELDNAME2>'
        + '<STITRE>Pièce</STITRE><SALIAS>CLE_NUMEROPIECE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION><TITLEALIGN>0</TITLEALIGN><TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR><BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE><IWIDTH>80</IWIDTH>'
        + '<ILEFT>0</ILEFT><BREQUIRED>X</BREQUIRED><BSEPMILLIER>-</BSEPMILLIER><BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL><INBDEC>0</INBDEC><BDESCRIPTION>-</BDESCRIPTION></$COL>' ;

  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_NUMLIGNE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_NUMLIGNE</SFIELDNAME2>'
        + '<STITRE>Ligne</STITRE>'
        + '<SALIAS>CLE_NUMLIGNE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>70</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + ' <BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;


  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_LIBELLE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_LIBELLE</SFIELDNAME2>'
        + '<STITRE>Libellé</STITRE>'
        + '<SALIAS>C1</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>200</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>-</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>'  ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_REFINTERNE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_LIBELLE1</SFIELDNAME2>'
        + '<STITRE>Référence interne</STITRE>'
        + '<SALIAS>C2</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>150</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>-</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_MODEPAIE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_CODEAFB</SFIELDNAME2>'
        + '<STITRE>Cib</STITRE>'
        + '<SALIAS>CLE_CIB</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>2</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>50</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_REFPOINTAGE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_REFPOINTAGE</SFIELDNAME2>'
        + '<STITRE>Réf. Pointage</STITRE>'
        + '<SALIAS>CLE_REFPOINTAGE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>100</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_DEVISE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_DEVISE</SFIELDNAME2>'
        + '<STITRE>Devise</STITRE>'
        + '<SALIAS>CLE_DEVISE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>2</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>40</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_DEBIT</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_DEBITEURO</SFIELDNAME2>'
        + '<STITRE>Debit</STITRE>'
        + '<SALIAS>CLE_DEBIT</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>100</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>X</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>2</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>'  ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_CREDIT</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_CREDITEURO</SFIELDNAME2>'
        + '<STITRE>Crédit</STITRE>'
        + '<SALIAS>CLE_CREDIT</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>100</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>X</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>2</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>'  ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>E_NATURETRESO</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_CODEPOINTAGE</SFIELDNAME2>'
        + '<STITRE>Code</STITRE>'
        + '<SALIAS>CLE_POINTE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>40</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

END ;

{------------------------------------------------------------------------------}
function ConstruireListePointageManuelTresorerie() : TOB  ;
var
  XmlParam : String ;
BEGIN
  result := TOB.Create('$COLS', Nil, -1) ;
  result.AddChampSupValeur('TABLE1', 'TRECRITURE', False);
  result.AddChampSupValeur('TABLE2', 'EEXBQLIG', False);

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_DATECOMPTABLE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_DATEOPERATION</SFIELDNAME2>'
        + '<STITRE>Date Compt.</STITRE>'
        + '<SALIAS>CLE_DATECOMPTABLE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>2</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>1</DTDISPLAYTYPE>'
        + '<IWIDTH>70</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_NUMEROPIECE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_NUMRELEVE</SFIELDNAME2>'
        + '<STITRE>Pièce</STITRE>'
        + '<SALIAS>CLE_NUMEROPIECE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>80</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>';
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_NUMLIGNE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_NUMLIGNE</SFIELDNAME2>'
        + '<STITRE>Ligne</STITRE>'
        + '<SALIAS>CLE_NUMLIGNE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>70</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>';
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_LIBELLE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_LIBELLE</SFIELDNAME2>'
        + '<STITRE>Libellé</STITRE>'
        + '<SALIAS>C1</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>200</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>-</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>';
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_REFINTERNE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_LIBELLE1</SFIELDNAME2>'
        + '<STITRE>Référence interne</STITRE>'
        + '<SALIAS>C2</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>150</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>-</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_CODECIB</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_CODEAFB</SFIELDNAME2>'
        + '<STITRE>Cib</STITRE>'
        + '<SALIAS>CLE_CIB</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>2</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>50</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_REFPOINTAGE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_REFPOINTAGE</SFIELDNAME2>'
        + '<STITRE>Réf. Pointage</STITRE>'
        + '<SALIAS>CLE_REFPOINTAGE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>100</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_DEVISE</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_DEVISE</SFIELDNAME2>'
        + '<STITRE>Devise</STITRE>'
        + '<SALIAS>CLE_DEVISE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>2</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>0</DTDISPLAYTYPE>'
        + '<IWIDTH>40</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>IIF(TE_MONTANT &lt; 0, (TE_MONTANT * (-1)), 0)</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_DEBITEURO</SFIELDNAME2>'
        + '<STITRE>Debit</STITRE>'
        + '<SALIAS>CLE_DEBIT</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>100</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>X</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>2</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>';
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>IIF(TE_MONTANT &gt; 0, TE_MONTANT, 0)</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_CREDITEURO</SFIELDNAME2>'
        + '<STITRE>Crédit</STITRE>'
        + '<SALIAS>CLE_CREDIT</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>1</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>100</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>X</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>2</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>' ;
  CreerListe(result, XmlParam) ;

  XmlParam :=
          '<$COL>'
        + '<SFIELDNAME1>TE_CODERAPPRO</SFIELDNAME1>'
        + '<SFIELDNAME2>CEL_CODEPOINTAGE</SFIELDNAME2>'
        + '<STITRE>Code</STITRE>'
        + '<SALIAS>CLE_POINTE</SALIAS>'
        + '<GFFUNCTION>0</GFFUNCTION>'
        + '<TITLEALIGN>0</TITLEALIGN>'
        + '<TEXTALIGN>0</TEXTALIGN>'
        + '<CLBACKCOLOR>0</CLBACKCOLOR>'
        + '<BFIELDMANUAL>-</BFIELDMANUAL>'
        + '<DTDISPLAYTYPE>2</DTDISPLAYTYPE>'
        + '<IWIDTH>40</IWIDTH>'
        + '<ILEFT>0</ILEFT>'
        + '<BREQUIRED>X</BREQUIRED>'
        + '<BSEPMILLIER>-</BSEPMILLIER>'
        + '<BVISIBLE>X</BVISIBLE>'
        + '<BEMPTYIFNULL>X</BEMPTYIFNULL>'
        + '<INBDEC>0</INBDEC>'
        + '<BDESCRIPTION>-</BDESCRIPTION>'
        + '</$COL>';
  CreerListe(result, XmlParam) ;

END ;

{------------------------------------------------------------------------------}
function CreerListeParametrablePourPointageManuel(theParameterType : String) : Boolean  ;
var
  iTrtParam : TITRTParam ;
  wTOBDATALIST  : TOB ;
BEGIN
  result := False ;
  wTOBDATALIST := Nil ;
  try
  if (theParameterType = 'CPM')
    or (theParameterType = 'CPT') then
  begin
    if Not IsExisteListeParametrable(theParameterType) then
    begin
      iTrtParam := TITRTParam.Create() ;
      iTrtParam.SetAction(ttaCREAT);
      iTrtParam.SetType(theParameterType);
      iTrtParam.SetUser('');
      iTrtParam.SetUserGroupe('');
      iTrtParam.SetCle(1);
      iTrtParam.SetVersion(0);
      iTrtParam.SetLibelle('Défaut');
      iTrtParam.SetDefaut(True);
      if (theParameterType = 'CPM') then
          wTOBDATALIST := ConstruireListePointageManuelStandard()
      Else
          wTOBDATALIST := ConstruireListePointageManuelTresorerie() ;
      iTrtParam.SetData(wTOBDATALIST);
      If Not iTrtParam.UpdateBD() then Exit ;
      result := True ;
    end
    else result := True ;
  end ;
  finally
    if assigned(wTOBDATALIST) then FreeAndNil(wTOBDATALIST) ;
  end ;
END ;
{
------------------------------------------------------------------------------
* Controle le type de paramètre autorisé dans cet objet ITRTPARAM
------------------------------------------------------------------------------
}

Function IsTypeAutorise(theType : String) : Boolean ;
var
  ENTETE : String ;
  MSG    : String ;
BEGIN
  result := False ;

  if (theType <> 'CPA')
    And (theType <> 'CPB')
    and (theType <> 'CPM')
    and (theType <> 'CPT')        {FP 14/12/2009 Paramétrage de la grille - FD 516}
    and (theType <> 'CPD') then   {FP 14/09/2010 FQ27742 Infos pour la traduction du libellé de la devise}
  begin
    ENTETE  := TraduireMemoire('Erreur technique') ;
    MSG     := TraduireMemoire('Le type de données n''est pas renseigné ou est incorrect') + ' :  "' +  theType + '".'  ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  result := True ;
END ;


{
*******************************************************************************
* objeT   TITRTParamList
*******************************************************************************
}
{------------------------------------------------------------------------------}
Function TITRTParamList.IsExisteParDefaut() : Boolean ;
var
  index       : Integer ;
  wAccro      : TITRTParam ;
  ITT_LIBELLE : String ;
  position    : Integer ;
BEGIN
  result := False ;

  for index := 0 to Liste.Count -1 do
  begin
    wAccro := TITRTParam(Liste.Objects[index]) ;
    if (wAccro <> Nil) then
    begin
      ITT_LIBELLE := wAccro.GetLibelle() ;
      position := Pos('*DEFAUT*', ITT_LIBELLE) ;
      if (position <> 0) then
      begin
        result := True ;
        Exit ;
      end ;
    end ;
  end ;
END ;

{------------------------------------------------------------------------------}
Constructor TITRTParamList.Create() ;
BEGIN

  inherited Create;
  Liste :=  TStringList.Create() ;
  Liste.Sorted := true ;
  bSearchByLibelleFg := False ;
  EnCours := Nil ;
END ;

{------------------------------------------------------------------------------}
Destructor TITRTParamList.Destroy() ;
BEGIN

 if (EnCours <> Nil) then FreeAndNil(EnCours) ;
 if (Liste <> Nil) then FreeAndNil(Liste) ;

 Inherited Destroy ;
END ;

{------------------------------------------------------------------------------}
Function TITRTParamList.LoadData(theType : String ; theSearchOnLibelleFg : Boolean = False) : Boolean ;
var
  SQL         : String ;
  QRY         : TQuery ;
  WTOBLIST    : TOB ;
  WTOB        : TOB ;
  Index       : Integer ;
  iTrtParam   : TITRTParam ;
  cle         : String ;
BEGIN
  result := False ;
  bSearchByLibelleFg := theSearchOnLibelleFg ;

  if Not IsTypeAutorise(theType) then Exit ;

  SQL := 'SELECT * FROM ITRTPARAM WHERE '
      +  ' ITT_TYPPAR = "' + theType + '" '
      +  ' AND '
      +  ' (ITT_USER="" OR ITT_USER="' + V_PGI.User + '" OR ITT_USER IS NULL'
      +  ' OR ITT_USERGRP = "" '
      +  ' OR ITT_USERGRP IS NULL '
      +  ' OR ITT_USERGRP = "' + FormatFloat('000', V_PGI.UserGrp) + '" )'

       ; {FP 18/12/2009 Pour ORACLE}

  QRY := OpenSQL(SQL, True) ;
  if Not QRY.EOF then
  begin
    WTOBLIST := TOB.Create('', Nil, -1) ;
    WTOBLIST.LoadDetailDB('ITRTPARAM', '', '', QRY, False) ;
    for index := 0 to WTOBLIST.Detail.Count -1 do
    begin
      wTOB := WTOBLIST.Detail[index] ;
      if (wTOB <> Nil) then
      begin
        iTrtParam := TITRTParam.Create() ;
        iTrtParam.LoadData(wTOB) ;
        iTrtParam.SetLikKey(bSearchByLibelleFg);
        cle := ConstituerCle(iTrtParam) ;
        Liste.AddObject(cle, iTrtParam) ;
      end ;
    end ;
    FreeAndNil(WTOBLIST);
  end ;
  Ferme(QRY) ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.Add(theITrtParam : TITRTParam) : Boolean ;
var
  cle         : String ;
  ITT_TYPPAR    : String ;
  ITT_LIBELLE : String ;
  ENTETE      : String ;
  MSG         : String ;
  position    : Integer ;
BEGIN
  result := False ;

  //* Controler le type de données

  ITT_TYPPAR    := theITrtParam.GetType() ;
  if Not IsTypeAutorise(ITT_TYPPAR) then Exit ;

  //* Controler si un autre item est déjà valeur par défaut

  ITT_LIBELLE := theITrtParam.GetLibelle() ;
  position := Pos('DEFAUT :', ITT_LIBELLE) ;
  if (position <> 0)
    and IsExisteParDefaut() then
  begin
    ENTETE  := TraduireMemoire('Attention') ;
    MSG     := TraduireMemoire('Vous ne pouvez pas avoir plusieurs paramétrages par défaut. Création impossible.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  theITrtParam.SetLikKey(bSearchByLibelleFg);

  cle := ConstituerCle(theITrtParam) ;
  Liste.AddObject(cle, theITrtParam) ;

  EnCours := theITrtParam ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function TITRTParamList.remove(theITrtParam : TITRTParam) : Boolean ;
var
  cle : String ;
  ITT_TYPPAR  : String ;
  Index     : Integer ;
BEGIN
  result := False ;

  ITT_TYPPAR    := theITrtParam.GetType() ;

  if Not IsTypeAutorise(ITT_TYPPAR) then Exit ;

  cle := ConstituerCle(theITrtParam) ;
  index := liste.IndexOf(cle) ;
  if (index <> -1) then Liste.Delete(index) ;

  result := True ;
END ;

{
------------------------------------------------------------------------------
* theQuelType : 0 = Tous, 1 = User uniquement,  2=Sans User
------------------------------------------------------------------------------
}
Function  TITRTParamList.GetParDefaut(theQuelType : Integer) : TITRTParam ;
var
  index   : Integer ;
  wAccro  : TITRTParam ;
  user    : String ;
begin
  result := Nil ;
   user   := V_PGI.User ;

  for index := 0 to liste.Count - 1 do
  begin;
    wAccro := TITRTParam(Liste.Objects[index]) ;
    if (wAccro <> Nil) then
    begin
      if  wAccro.getDefaut() then
      begin
        CASE theQuelType of
          0 :
            begin
              result := wAccro ;
              EnCours := wAccro ;
              Exit ;
            end ;
          1 :
            begin
              if (wAccro.GetUser() = user) then
              begin
                result := wAccro ;
                EnCours := wAccro ;
                Exit ;
              end ;
            end ;
          2 :
            begin
              if (wAccro.GetUser() = '') then
              begin
                result := wAccro ;
                EnCours := wAccro ;
                Exit ;
              end ;
            end ;
        END ;
      end ;
    end ;
  end ;
end ;
{
------------------------------------------------------------------------------
* Dans ce cas, on recherche par rapport à l'utilateur et ensuite par le public
}
Function  TITRTParamList.GetParDefaut() : TITRTParam ;
var
  index   : Integer ;
  wAccro  : TITRTParam ;
  user    : String ;
BEGIN
  result := Nil ;
  user   := V_PGI.User ;

//* Recherche dans le domaine Privé

  for index := 0 to liste.Count - 1 do
  begin;
    wAccro := TITRTParam(Liste.Objects[index]) ;
    if (wAccro <> Nil) then
    begin
      if (wAccro.GetUser() = user) and wAccro.getDefaut() then
      begin
        result := wAccro ;
        EnCours := wAccro ;
        Exit ;
      end ;
    end ;
  end ;

//* Recherche dans le domaine public

  for index := 0 to liste.Count - 1 do
  begin;
    wAccro := TITRTParam(Liste.Objects[index]) ;
    if (wAccro <> Nil) then
    begin
      if (wAccro.GetUser() = '') and wAccro.getDefaut() then
      begin
        result := wAccro ;
        EnCours := wAccro ;
        Exit ;
      end ;
    end ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.GetEnCours() : TITRTParam ;
BEGIN
  result := Nil ;
  if (EnCours <> Nil) then result := EnCours ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.Selectionner(theSearchKey : String) : TITRTParam ;
var
  index : Integer ;
BEGIN
  result := Nil ;
  if (theSearchKey = '') then exit ;

  index := Liste.IndexOf(theSearchKey) ;
  if (index <> -1) then
  begin
    result := TITRTParam(Liste.Objects[index]) ;
    EnCours := result ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.Count : Integer  ;
BEGIN
  result := Liste.Count ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.Detail (theOccurenceNumber : Integer) : TITRTParam ;
BEGIN
  result := Nil ;
  if (theOccurenceNumber >  Liste.Count - 1) then Exit ;

  result := TITRTParam(Liste.Objects[theOccurenceNumber]) ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.DeleteBD() : Boolean;
var
  index     : Integer ;
  accochage : TITRTParam ;
BEGIN
  result := False ;
  for index := 0 to Liste.Count - 1  do
  begin
    accochage := Detail(index) ;
    if (accochage <> Nil) then
    begin
      if Not accochage.DeleteBD() then Exit ;
    end ;
  end ;

  result := True ;
END ;

{------------------------------------------------------------------------------}
Function  TITRTParamList.UpdateAndDeleteBD () : Boolean ;
var
  index     : Integer ;
  accochage : TITRTParam ;
BEGIN
  result := False ;
  for index := 0 to Liste.Count - 1  do
  begin
    accochage := Detail(index) ;
    if (accochage <> Nil) then
    begin
      case accochage.getAction() of
        ttaCREAT, ttaMODIF : if not  accochage.UpdateBD() then Exit ;
        ttaDELETE : if not accochage.DeleteBD() then Exit  ;
      end ;
    end ;
  end ;
  result := True ;
END ;
{------------------------------------------------------------------------------}
Function  TITRTParamList.UpdateBD() : Boolean ;
var
  index     : Integer ;
  accochage : TITRTParam ;
BEGIN
  result := False ;
  for index := 0 to Liste.Count - 1  do
  begin
    accochage := Detail(index) ;
    if (accochage <> Nil) then
    begin
      if Not accochage.UpdateBD() then Exit ;
    end ;
  end ;

  result := True ;
END ;


{------------------------------------------------------------------------------}
Function TITRTParamList.ConstituerCle(theITrtParam : TITRTParam) : String  ;
var
  ITT_TYPPAR    : String ;
  ITT_USERGRP : String ;
  ITT_USER    : String ;
  ITT_CLE     : Integer ;
  ITT_VERSION : Integer ;
  ITT_LIBELLE : String  ;
BEGIN
  if bSearchByLibelleFg then
  begin
    ITT_LIBELLE := theITrtParam.GetLibelle() ;
    ITT_CLE     := theITrtParam.GetCle() ;
    Result      := '<LIB>' +  ITT_LIBELLE       + '</LIB>'
                +  '<CLE>' +  IntToStr(ITT_CLE) + '</CLE>'  ;
  end
  else
  begin
    ITT_TYPPAR    := theITrtParam.GetType() ;
    ITT_USERGRP := theITrtParam.GetUserGroupe() ;
    ITT_USER    := theITrtParam.GetUser() ;
    ITT_CLE     := theITrtParam.GetCle() ;
    ITT_VERSION := theITrtParam.GetVersion ;
    Result := ITT_TYPPAR + ITT_USERGRP + ITT_USER + IntToStr(ITT_CLE) + IntToStr(ITT_VERSION) ;
  end ;
END ;

{
*******************************************************************************
* objeT   TITRTParam
*******************************************************************************
}

{------------------------------------------------------------------------------}
Constructor TITRTParam.Create() ;
BEGIN

  inherited Create;

  TOBITRTPARAM      := TOB.Create('ITRTPARAM', Nil, -1) ;
  TOBDEFAIRE        := TOB.Create('', Nil, -1) ;
  TDATAList         := TOB.Create('', Nil, -1) ;
  TDATADEFAIREList  := TOB.Create('', Nil, -1) ;

  TOBITRTPARAM.InitValeurs();
  TOBDEFAIRE.InitValeurs();

  action      :=  ttaRIEN ;
  bLibKeyFg   := False ;
END ;

{------------------------------------------------------------------------------}
Function TITRTParam.LoadData(theTOBITRTPARAM : TOB) : Boolean  ;
BEGIN
  result := False ;

  if (theTOBITRTPARAM.detail.Count <> 0) then
  begin
    ENTETE  := TraduireMemoire('Erreur technique') ;
    MSG     := TraduireMemoire('Impossible de charger les données. Il y a trop d''informations.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  if (theTOBITRTPARAM.NomTable = '') then
  begin
    ENTETE  := TraduireMemoire('Erreur technique') ;
    MSG     := TraduireMemoire('Le nom de la table n''est pas renseigné.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  if (theTOBITRTPARAM.NomTable <> 'ITRTPARAM') then
  begin
    ENTETE  := TraduireMemoire('Erreur technique') ;
    MSG     := TraduireMemoire('Le TOB passée en paramètre ne contient pas le type "ITRTPARAM" de données attendues.') ;
    PGIError(MSG, ENTETE);
    Exit ;
  end ;

  TOBITRTPARAM.ClearDetail() ;
  TOBITRTPARAM.Dupliquer(theTOBITRTPARAM, True, True, False);

  if (GetType = 'CPA') then Rempli_CPA_TDATAList(False) ;

  TOBDEFAIRE.ClearDetail() ;
  TOBDEFAIRE.Dupliquer(theTOBITRTPARAM, True, True, False);

  if (GetType = 'CPA') then Rempli_CPA_TDATAList(True) ;

  SetAction(ttaMODIF);

  result := True ;
END ;



{------------------------------------------------------------------------------}
Destructor TITRTParam.Destroy() ;
BEGIN
  if (TOBITRTPARAM <> Nil)     then FreeAndNil(TOBITRTPARAM) ;
  if (TOBDEFAIRE <> Nil)       then FreeAndNil(TOBDEFAIRE) ;
  if (TDATAList <> Nil)        then FreeAndNil(TDATAList) ;
  if (TDATADEFAIREList <> Nil) then FreeAndNil(TDATADEFAIREList) ;
  Inherited Destroy ;
END ;

{------------------------------------------------------------------------------}
Procedure TITRTParam.UpdateDataBase ;
var
  ITT_CLE     : Integer ;
BEGIN
  CASE Action OF
    ttaRIEN :
      begin
        ENTETE  :=  TraduireMemoire('Erreur technique') ;
        MSG     := TraduireMemoire('Impossible de faire une action dans la base de données.')
                +  #13#13
                +  TraduireMemoire('Le paramètre "Action" n''est pas en mode "taModif" ou "taCreat".') ;
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    else
    begin
      ITT_CLE     := GetCle() ; ;
      Case Action of
        ttaMODIF :
          begin
            if Not Exists() then
            begin
              ENTETE :=  TraduireMemoire('Attention') ;

              if (GetType() = 'CPA') then
                    MSG := TraduireMemoire('La règle d''accrochage "')
              else
                    MSG := TraduireMemoire('L''identifiant "')  ;

              MSG   := MSG
                    + IntToStr(ITT_CLE) +'" - "'
                    + GetLibelle()
                    + TraduireMemoire('" n''existe pas.') ;

              PGIError(MSG, ENTETE) ;
              V_PGI.IoError := oeUnknown ;
              Exit ;
            end ;
            TOBITRTPARAM.UpdateDB(False) ;
          end ;
        ttaCREAT :
          begin
            if Exists() then
            begin
              ENTETE :=  TraduireMemoire('Attention') ;

              if (GetType() = 'CPA') then
                    MSG := TraduireMemoire('La règle d''accrochage "')
              else
                    MSG := TraduireMemoire('L''identifiant "')  ;

              MSG   := MSG
                    + IntToStr(ITT_CLE) +'" - "'
                    + GetLibelle()
                    + TraduireMemoire('" existe déjà.') ;

              PGIError(MSG, ENTETE) ;
              V_PGI.IoError := oeUnknown ;
              Exit ;
            end ;
            TOBITRTPARAM.InsertDB(Nil) ;
          end ;
      End ;
    end ;
  END ;
END ;
{------------------------------------------------------------------------------}
Function TITRTParam.Defaire(): Boolean;
BEGIN
  TOBITRTPARAM.ClearDetail() ;
  TOBITRTPARAM.Dupliquer(TOBDEFAIRE, True, True, True);

  TDATAList.ClearDetail() ;
  TDATAList.Dupliquer(TDATADEFAIREList, True, True, True);

  result := True ;
END ;
{------------------------------------------------------------------------------}
function TITRTParam.UpdateBD(): Boolean;
var
  IOERR   : TIOErr ;
BEGIN
  result := False ;

  if Not IsAllControler() then Exit ;

  ENTETE :=  TraduireMemoire('Erreur technique') ;

  IOERR := Transactions(UpdateDataBase, 0) ;
  if  (IOERR <> oeOK) And (IOERR <> oeSaisie) Then
  begin
    MSG := TraduireMemoire('Impossible de mettre à jour la table ITRTPARAM.') ;
    PGIInfo(MSG, ENTETE) ;
    Exit ;
  end ;
  result := True ;
END ;

{------------------------------------------------------------------------------}
Procedure TITRTParam.DeleteDataBase() ;
var
  ITT_CLE     : Integer ;
BEGIN
  ENTETE :=  TraduireMemoire('Attention') ;
  CASE Action OF
    ttaRIEN :
      begin
        ENTETE  :=  TraduireMemoire('Erreur de petit programmeur') ;
        MSG     := TraduireMemoire('Impossible de faire une action dans la base de données.')
                +  #13#13
                +  TraduireMemoire('Le paramètre "Action" n''est pas en mode "taModif" ou "taCreat".') ;
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    ttaCREAT :
      begin
        MSG    :=  TraduireMemoire('Impossible de supprimer une clé qui n''existe pas dans la base de données.');
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    ttaMODIF :
      begin
        MSG    :=  TraduireMemoire('Impossible de supprimer une clé dans le mode modification.');
        PGIError(MSG, ENTETE) ;
        V_PGI.IoError := oeUnknown ;
        Exit ;
      end ;
    ttaDELETE :
      begin
        ITT_CLE     := GetCle() ; ;
        if Not Exists() then
        begin
              if (GetType() = 'CPA') then
                    MSG := TraduireMemoire('La règle d''accrochage "')
              else
                    MSG := TraduireMemoire('L''identifiant "')  ;

              MSG   := MSG
                    + IntToStr(ITT_CLE) +'" - "'
                    + GetLibelle()
                 + TraduireMemoire('" n''existe pas.') ;
          PGIError(MSG, ENTETE) ;
          V_PGI.IoError := oeUnknown ;
          Exit ;
        end ;
        TOBITRTPARAM.DeleteDB(False) ;
      end ;
  END ;
END ;

{------------------------------------------------------------------------------}
function TITRTParam.DeleteBD(): Boolean;
var
  IOERR   : TIOErr ;
  wAction : TActionTable ;
BEGIN
  result := False ;

  wAction := GetAction() ;
  SetAction(ttaDELETE);

  IOERR := Transactions(DeleteDataBase, 0) ;
  if  (IOERR <> oeOK) And (IOERR <> oeSaisie) Then
  begin
    ENTETE :=  TraduireMemoire('Erreur technique') ;
    MSG := TraduireMemoire('Impossible de supprimer la souche de la table ITRTPARAM.') ;
    PGIError(MSG, ENTETE) ;
    SetAction(wAction);
    Exit ;
  end ;

  result  := True ;
END ;
{------------------------------------------------------------------------------}
Function TITRTParam.GetKey() : String ;
var
  ITT_TYPPAR  : String ;
  ITT_USERGRP : String ;
  ITT_USER    : String ;
  ITT_LIBELLE : String ;
  ITT_CLE     : Integer ;
  ITT_VERSION : Integer ;
begin
  if bLibKeyFg then
  begin
    ITT_LIBELLE := GetLibelle() ;
    ITT_CLE     := GetCle() ;
    Result      := '<LIB>' +  ITT_LIBELLE       + '</LIB>'
                +  '<CLE>' +  IntToStr(ITT_CLE) + '</CLE>'  ;
  end
  else
  begin
    ITT_TYPPAR  := GetType() ;
    ITT_USERGRP := GetUserGroupe() ;
    ITT_USER    := GetUser() ;
    ITT_CLE     := GetCle() ;
    ITT_VERSION := GetVersion() ;
    Result := ITT_TYPPAR + ITT_USERGRP + ITT_USER + IntToStr(ITT_CLE) + IntToStr(ITT_VERSION) ;
  end ;
end ;
{------------------------------------------------------------------------------}
function TITRTParam.GetAction(): TActionTable;
begin
  result := Action ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetAction(theAction : TActionTable)  ;
BEGIN
  action := theAction ;
END ;

{------------------------------------------------------------------------------}
function TITRTParam.GetCle: Integer;
begin
  result := TOBITRTPARAM.GetValue('ITT_CLE') ;
end;

{------------------------------------------------------------------------------}
function TITRTParam.GetDefaut: Boolean ;
begin
  result := StrToBool(TOBITRTPARAM.GetValue('ITT_DEFAUT')) ;
end;

{------------------------------------------------------------------------------}
function TITRTParam.GetLibelle: String;
begin
  result := TOBITRTPARAM.GetValue('ITT_LIBELLE') ;
end;

{------------------------------------------------------------------------------}
function TITRTParam.GetType: String;
begin
  result := TOBITRTPARAM.GetValue('ITT_TYPPAR') ;
end;

{------------------------------------------------------------------------------}
function TITRTParam.GetUser: String;
begin
  result := TOBITRTPARAM.GetValue('ITT_USER') ;
end;

{------------------------------------------------------------------------------}
Function  TITRTParam.GetUserGroupe()  : String ;
begin
  result := TOBITRTPARAM.GetValue('ITT_USERGRP') ;
end;

{------------------------------------------------------------------------------}
Function  TITRTParam.GetVersion()     : Integer ;
begin
  result := TOBITRTPARAM.GetValue('ITT_VERSION') ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetType(theType         : String) ;
begin

  //* Controle du type de paramètre géré dans le 

  if Not IsTypeAutorise(theType) then Exit ;

  TOBITRTPARAM.PutValue('ITT_TYPPAR',theType) ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetUser(theUser         : String) ;
begin
  TOBITRTPARAM.PutValue('ITT_USER',theUser) ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetUserGroupe(theGroupe : String) ;
begin
  TOBITRTPARAM.PutValue('ITT_USERGRP',theGroupe) ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetCle(theCle           : Integer) ;
begin
  TOBITRTPARAM.PutValue('ITT_CLE', theCle) ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetVersion(theVersion   : Integer) ;
begin
  TOBITRTPARAM.PutValue('ITT_VERSION', theVersion) ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetLibelle(theLibelle   : String) ;
begin
  TOBITRTPARAM.PutValue('ITT_LIBELLE', theLibelle) ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetDefaut(theDefaultFg  : Boolean) ;
begin
  TOBITRTPARAM.PutValue('ITT_DEFAUT', BoolToStr(theDefaultFg)) ;
end;


{------------------------------------------------------------------------------}
Function  TITRTParam.AffecterCle() : Boolean;
var
  SQL     : String ;
  QRY     : TQuery ;
  ITT_CLE : Integer ;
BEGIN
  result := True ;

  ITT_CLE := 0 ;
  SQL := 'SELECT MAX(ITT_CLE) AS ITT_CLE FROM ITRTPARAM WHERE ITT_TYPPAR = "' + GetType()  + '"' ;
  QRY := OpenSQL(SQL, True) ;
  If Not QRY.EOF then ITT_CLE :=  QRY.Fields[0].AsInteger ;
  Inc(ITT_CLE) ;
  SetCle(ITT_CLE);
END ;

{------------------------------------------------------------------------------}
Function TITRTParam.IsAllControler()  : Boolean ;
BEGIN
  result := True ;
END ;

{------------------------------------------------------------------------------
* Analyse les données du champ ITT_DATA pour les mettre dans la TOBDATA.
}
Procedure  TITRTParam.Rempli_CPA_TDATAList(bDefaireFg : Boolean = False) ;
var
  TDATA     : TOB ;

  ITT_DATA  : String ;

  wITT_DATA : String ;
  _ENREG    : String ;
  _RIB      : String ;
  position  : Integer ;
  longueur  : Integer ;
  _longueur : Integer ;
BEGIN
  ITT_DATA := TOBITRTPARAM.GetValue('ITT_DATA') ;

  while (ITT_DATA <> '') do
  begin
    _ENREG := RechercherParameterXML('ENREG', ITT_DATA) ;
    if  (_ENREG <> '') then
    begin
      _RIB := RechercherParameterXML('RIB', _ENREG) ;
      if (_RIB <> 'RIB') then
      begin
        if Not bDefaireFg then TDATA := TOB.Create('', TDATAList, -1)
                          else TDATA := TOB.Create('', TDATADEFAIREList, -1);
        TDATA.AddChampSupValeur('RIB',              RechercherParameterXML('RIB', _ENREG)) ;
        TDATA.AddChampSupValeur('BANQUE',           RechercherParameterXML('BANQUE', _ENREG)) ;
        TDATA.AddChampSupValeur('GUICHET',          RechercherParameterXML('GUICHET', _ENREG)) ;
        TDATA.AddChampSupValeur('COMPTE',           RechercherParameterXML('COMPTE', _ENREG)) ;
        TDATA.AddChampSupValeur('CLE',              RechercherParameterXML('CLE', _ENREG)) ;
        TDATA.AddChampSupValeur('MONTANT',          RechercherParameterXML('MONTANT', _ENREG)) ;
        TDATA.AddChampSupValeur('SENS',             RechercherParameterXML('SENS', _ENREG)) ;
        TDATA.AddChampSupValeur('COMPTECOMPTABLE',  RechercherParameterXML('COMPTECOMPTABLE', _ENREG)) ;
        TDATA.AddChampSupValeur('CAPPLIQUER',       RechercherParameterXML('CAPPLIQUER', _ENREG)) ;
        TDATA.AddChampSupValeur('CDATE',            RechercherParameterXML('CDATE', _ENREG)) ;
        TDATA.AddChampSupValeur('TDATE',            RechercherParameterXML('TDATE', _ENREG)) ;
        TDATA.AddChampSupValeur('CMODEREG',         RechercherParameterXML('CMODEREG', _ENREG)) ;
        TDATA.AddChampSupValeur('CKLIBCIB',         RechercherParameterXML('CKLIBCIB', _ENREG)) ;
        TDATA.AddChampSupValeur('CCOMBINATOIRE',    RechercherParameterXML('CCOMBINATOIRE', _ENREG)) ;
        TDATA.AddChampSupValeur('TNIVEAU',          RechercherParameterXML('TNIVEAU', _ENREG)) ;
        TDATA.AddChampSupValeur('CNUMCHEQUE',       RechercherParameterXML('CNUMCHEQUE', _ENREG)) ;
        TDATA.AddChampSupValeur('TNUMCHEQUE',       RechercherParameterXML('TNUMCHEQUE', _ENREG)) ;
        TDATA.AddChampSupValeur('TPOSCHEQUE',       RechercherParameterXML('TPOSCHEQUE', _ENREG)) ;
        TDATA.AddChampSupValeur('COMBOREF',         RechercherParameterXML('COMBOREF', _ENREG)) ;
        TDATA.AddChampSupValeur('CBPARAM2',         RechercherParameterXML('CBPARAM2', _ENREG)) ;
        TDATA.AddChampSupValeur('CBCASSE',          RechercherParameterXML('CBCASSE', _ENREG)) ;
        TDATA.AddChampSupValeur('COMBOREFBQ',       RechercherParameterXML('COMBOREFBQ', _ENREG)) ;
        TDATA.AddChampSupValeur('CBOKPOSBQ',        RechercherParameterXML('CBOKPOSBQ', _ENREG)) ;
        TDATA.AddChampSupValeur('POSBQ1',           RechercherParameterXML('POSBQ1', _ENREG)) ;
        TDATA.AddChampSupValeur('POSBQ2',           RechercherParameterXML('POSBQ2', _ENREG)) ;
        TDATA.AddChampSupValeur('COMBOREFECR',      RechercherParameterXML('COMBOREFECR', _ENREG)) ;
        TDATA.AddChampSupValeur('CBOKPOSECR',       RechercherParameterXML('CBOKPOSECR', _ENREG)) ;
        TDATA.AddChampSupValeur('POSECR1',          RechercherParameterXML('POSECR1', _ENREG)) ;
        TDATA.AddChampSupValeur('POSECR2',          RechercherParameterXML('POSECR2', _ENREG)) ;
      end ;

        //* Je vire les données traitées du champ ITT_DATA ;

      position := Pos('</$ENREG$>', ITT_DATA) ;
      if (position <> 0) then
      begin
        longueur  := Length(ITT_DATA) ;
        _longueur := Length(_ENREG) ;
        if (longueur <> _longueur) then
        begin
          longueur  := (longueur - _longueur) + 2 ;
          Position  := position + length('</$ENREG$>');
          wITT_DATA := Trim(Copy(ITT_DATA, position, longueur)) ;
          ITT_DATA  :=  wITT_DATA ;
        end ;
      end ;

    end ;
  end ;
END ;

{------------------------------------------------------------------------------}
Function TITRTParam.GetRegleAccrochage(theTOBList : TOB) : Boolean;
var
  index : Integer ;
  wTOB  : TOB ;
  TDATA : TOB ;
  RIB   : String ;
BEGIN
  result := False ;
  INITMOVE(theTOBList.Detail.Count, '');
  for index := 0 to theTOBList.Detail.Count -1 do
  begin
    WTOB := theTOBList.Detail[index] ;
    if (WTOB <> Nil) then
    begin
      MOVECUR(False);
      RIB := WTOB.GetValue('RIB') ;
      TDATA := TDATAList.FindFirst(['RIB'], [RIB], True) ;
      if (TDATA <> Nil) then
      begin
        WTOB.PutValue('RIB',              TDATA.getValue('RIB')) ;
        WTOB.PutValue('BANQUE',           TDATA.getValue('BANQUE')) ;
        WTOB.PutValue('GUICHET',          TDATA.getValue('GUICHET')) ;
        WTOB.PutValue('COMPTE',           TDATA.getValue('COMPTE')) ;
        WTOB.PutValue('CLE',              TDATA.getValue('CLE')) ;
        WTOB.PutValue('MONTANT',          TDATA.getValue('MONTANT')) ;
        WTOB.PutValue('SENS',             TDATA.getValue('SENS')) ;
        WTOB.PutValue('COMPTECOMPTABLE',  TDATA.getValue('COMPTECOMPTABLE')) ;
        WTOB.PutValue('CAPPLIQUER',       TDATA.getValue('CAPPLIQUER')) ;
        WTOB.PutValue('CDATE',            TDATA.getValue('CDATE')) ;
        WTOB.PutValue('TDATE',            TDATA.getValue('TDATE')) ;
        WTOB.PutValue('CMODEREG',         TDATA.getValue('CMODEREG')) ;
        WTOB.PutValue('CKLIBCIB',         TDATA.getValue('CKLIBCIB')) ;
        WTOB.PutValue('CCOMBINATOIRE',    TDATA.getValue('CCOMBINATOIRE')) ;
        WTOB.PutValue('TNIVEAU',          TDATA.getValue('TNIVEAU')) ;
        WTOB.PutValue('CNUMCHEQUE',       TDATA.getValue('CNUMCHEQUE')) ;
        WTOB.PutValue('TNUMCHEQUE',       TDATA.getValue('TNUMCHEQUE')) ;
        WTOB.PutValue('TPOSCHEQUE',       TDATA.getValue('TPOSCHEQUE')) ;
        WTOB.PutValue('COMBOREF',         TDATA.getValue('COMBOREF')) ;
        WTOB.PutValue('CBPARAM2',         TDATA.getValue('CBPARAM2')) ;
        WTOB.PutValue('CBCASSE',          TDATA.getValue('CBCASSE')) ;
        WTOB.PutValue('COMBOREFBQ',       TDATA.getValue('COMBOREFBQ')) ;
        WTOB.PutValue('CBOKPOSBQ',        TDATA.getValue('CBOKPOSBQ')) ;
        WTOB.PutValue('POSBQ1',           TDATA.getValue('POSBQ1')) ;
        WTOB.PutValue('POSBQ2',           TDATA.getValue('POSBQ2')) ;
        WTOB.PutValue('COMBOREFECR',      TDATA.getValue('COMBOREFECR')) ;
        WTOB.PutValue('CBOKPOSECR',       TDATA.getValue('CBOKPOSECR')) ;
        WTOB.PutValue('POSECR1',          TDATA.getValue('POSECR1')) ;
        WTOB.PutValue('POSECR2',          TDATA.getValue('POSECR2')) ;

        result := True ;
      end ;
    end ;
  end ;
  FINIMOVE() ;
END ;


{
------------------------------------------------------------------------------
* Format une chaine XML pour les règles d'accrochage
------------------------------------------------------------------------------
}
Procedure  TITRTParam.Rempli_CPA_ITT_DATA() ;
var
  index     : Integer ;
  TDATA     : TOB ;
  ITT_DATA  : String ;
  _RIB      : String ;
  _Enreg    : String ;
  boCbParam2 : Boolean ;
BEGIN
  ITT_DATA := '' ;
  TDATAList.Detail.Sort('RIB');
  for index := 0 to TDATAList.Detail.Count -1 do
  begin
    TDATA := TDATAList.Detail[index] ;
    if (TDATA <> Nil) then
    begin
      _RIB :=  TDATA.getValue('RIB') ;
      if (_RIB <> 'RIB') then
      begin
        boCbParam2 := Hent1.StrToBool(TDATA.getValue('CBPARAM2')) ;

        _enreg := '' ;
        _enreg := _enreg + FormatTag('RIB',             TDATA.getValue('RIB')) ;
        _enreg := _enreg + FormatTag('BANQUE',          TDATA.getValue('BANQUE')) ;
        _enreg := _enreg + FormatTag('GUICHET',         TDATA.getValue('GUICHET')) ;
        _enreg := _enreg + FormatTag('COMPTE',          TDATA.getValue('COMPTE')) ;
        _enreg := _enreg + FormatTag('CLE',             TDATA.getValue('CLE')) ;
        _enreg := _enreg + FormatTag('MONTANT',         TDATA.getValue('MONTANT')) ;
        _enreg := _enreg + FormatTag('SENS',            TDATA.getValue('SENS')) ;
        _enreg := _enreg + FormatTag('COMPTECOMPTABLE', TDATA.getValue('COMPTECOMPTABLE')) ;
        _enreg := _enreg + FormatTag('CAPPLIQUER',      TDATA.getValue('CAPPLIQUER')) ;
        _enreg := _enreg + FormatTag('CDATE',           TDATA.getValue('CDATE')) ;
        _enreg := _enreg + FormatTag('TDATE',           TDATA.getValue('TDATE')) ;
        _enreg := _enreg + FormatTag('CMODEREG',        TDATA.getValue('CMODEREG')) ;
        _enreg := _enreg + FormatTag('CKLIBCIB',        TDATA.getValue('CKLIBCIB')) ;
        _enreg := _enreg + FormatTag('CCOMBINATOIRE',   TDATA.getValue('CCOMBINATOIRE')) ;
        _enreg := _enreg + FormatTag('TNIVEAU',         TDATA.getValue('TNIVEAU')) ;
        _enreg := _enreg + FormatTag('CNUMCHEQUE',      TDATA.getValue('CNUMCHEQUE')) ;
        _enreg := _enreg + FormatTag('TNUMCHEQUE',      TDATA.getValue('TNUMCHEQUE')) ;
        _enreg := _enreg + FormatTag('TPOSCHEQUE',      TDATA.getValue('TPOSCHEQUE')) ;
        _enreg := _enreg + FormatTag('COMBOREF',        TDATA.getValue('COMBOREF')) ;
        _enreg := _enreg + FormatTag('CBPARAM2',        TDATA.getValue('CBPARAM2')) ;
        if boCbParam2 then
        begin
          _enreg := _enreg + FormatTag('CBCASSE',       TDATA.getValue('CBCASSE')) ;
          _enreg := _enreg + FormatTag('COMBOREFBQ',    TDATA.getValue('COMBOREFBQ')) ;
          _enreg := _enreg + FormatTag('CBOKPOSBQ',     TDATA.getValue('CBOKPOSBQ')) ;
          _enreg := _enreg + FormatTag('POSBQ1',        TDATA.getValue('POSBQ1')) ;
          _enreg := _enreg + FormatTag('POSBQ2',        TDATA.getValue('POSBQ2')) ;
          _enreg := _enreg + FormatTag('COMBOREFECR',   TDATA.getValue('COMBOREFECR')) ;
          _enreg := _enreg + FormatTag('CBOKPOSECR',    TDATA.getValue('CBOKPOSECR')) ;
          _enreg := _enreg + FormatTag('POSECR1',       TDATA.getValue('POSECR1')) ;
          _enreg := _enreg + FormatTag('POSECR2',       TDATA.getValue('POSECR2')) ;
        end
        else
        begin
          _enreg := _enreg + FormatTag('CBCASSE',     '-') ;
          _enreg := _enreg + FormatTag('COMBOREFBQ',  '') ;
          _enreg := _enreg + FormatTag('CBOKPOSBQ',   '-') ;
          _enreg := _enreg + FormatTag('POSBQ1',      '0') ;
          _enreg := _enreg + FormatTag('POSBQ2',      '0') ;
          _enreg := _enreg + FormatTag('COMBOREFECR', '') ;
          _enreg := _enreg + FormatTag('CBOKPOSECR',  '-') ;
          _enreg := _enreg + FormatTag('POSECR1',     '0') ;
          _enreg := _enreg + FormatTag('POSECR2',     '0') ;
        end ;

        ITT_DATA := ITT_DATA + FormatTag('ENREG', _enreg) ;
      end ;
    end ;
  end ;
  TOBITRTPARAM.PutValue('ITT_DATA', ITT_DATA) ;
END ;

{
------------------------------------------------------------------------------
* Met à jour la TOBDATA de ITRTPARAM pour les règles d'accrochage
------------------------------------------------------------------------------
}
Function TITRTParam.SetRegleAccrochage(theTOBList : TOB) : Boolean;
var
  index : Integer ;
  wTOB  : TOB ;
  TDATA : TOB ;
  RIB   : String ;
BEGIN
  result := False ;

  for index := 0 to theTOBList.Detail.Count -1 do
  begin
    WTOB := theTOBList.Detail[index] ;
    if (WTOB <> Nil) then
    begin
      RIB := WTOB.GetValue('RIB') ;
      if (RIB <> 'RIB') then
      begin
      TDATA := TDATAList.FindFirst(['RIB'], [RIB], True) ;
      if (TDATA <> Nil) then
      begin
          TDATA.PutValue('RIB',              WTOB.getValue('RIB')) ;
          TDATA.PutValue('BANQUE',           WTOB.getValue('BANQUE')) ;
          TDATA.PutValue('GUICHET',          WTOB.getValue('GUICHET')) ;
          TDATA.PutValue('COMPTE',           WTOB.getValue('COMPTE')) ;
          TDATA.PutValue('CLE',              WTOB.getValue('CLE')) ;
          TDATA.PutValue('MONTANT',          WTOB.getValue('MONTANT')) ;
          TDATA.PutValue('SENS',             WTOB.getValue('SENS')) ;
          TDATA.PutValue('COMPTECOMPTABLE',  WTOB.getValue('COMPTECOMPTABLE')) ;
          TDATA.PutValue('CAPPLIQUER',       WTOB.getValue('CAPPLIQUER')) ;
          TDATA.PutValue('CDATE',            WTOB.getValue('CDATE')) ;
          TDATA.PutValue('TDATE',            WTOB.getValue('TDATE')) ;
          TDATA.PutValue('CMODEREG',         WTOB.getValue('CMODEREG')) ;
          TDATA.PutValue('CKLIBCIB',         WTOB.getValue('CKLIBCIB')) ;
          TDATA.PutValue('CCOMBINATOIRE',    WTOB.getValue('CCOMBINATOIRE')) ;
          TDATA.PutValue('TNIVEAU',          WTOB.getValue('TNIVEAU')) ;
          TDATA.PutValue('CNUMCHEQUE',       WTOB.getValue('CNUMCHEQUE')) ;
          TDATA.PutValue('TNUMCHEQUE',       WTOB.getValue('TNUMCHEQUE')) ;
          TDATA.PutValue('TPOSCHEQUE',       WTOB.getValue('TPOSCHEQUE')) ;
          TDATA.PutValue('COMBOREF',         WTOB.getValue('COMBOREF')) ;
          TDATA.PutValue('CBPARAM2',         WTOB.getValue('CBPARAM2')) ;
          TDATA.PutValue('CBCASSE',          WTOB.getValue('CBCASSE')) ;
          TDATA.PutValue('COMBOREFBQ',       WTOB.getValue('COMBOREFBQ')) ;
          TDATA.PutValue('CBOKPOSBQ',        WTOB.getValue('CBOKPOSBQ')) ;
          TDATA.PutValue('POSBQ1',           WTOB.getValue('POSBQ1')) ;
          TDATA.PutValue('POSBQ2',           WTOB.getValue('POSBQ2')) ;
          TDATA.PutValue('COMBOREFECR',      WTOB.getValue('COMBOREFECR')) ;
          TDATA.PutValue('CBOKPOSECR',       WTOB.getValue('CBOKPOSECR')) ;
          TDATA.PutValue('POSECR1',          WTOB.getValue('POSECR1')) ;
          TDATA.PutValue('POSECR2',          WTOB.getValue('POSECR2')) ;
          result := True ;
        end
        else
        begin
          TDATA := TOB.Create('', TDATAList, -1) ;

          TDATA.AddChampSupValeur('RIB',              WTOB.getValue('RIB')) ;
          TDATA.AddChampSupValeur('BANQUE',           WTOB.getValue('BANQUE')) ;
          TDATA.AddChampSupValeur('GUICHET',          WTOB.getValue('GUICHET')) ;
          TDATA.AddChampSupValeur('COMPTE',           WTOB.getValue('COMPTE')) ;
          TDATA.AddChampSupValeur('CLE',              WTOB.getValue('CLE')) ;
          TDATA.AddChampSupValeur('MONTANT',          WTOB.getValue('MONTANT')) ;
          TDATA.AddChampSupValeur('SENS',             WTOB.getValue('SENS')) ;
          TDATA.AddChampSupValeur('COMPTECOMPTABLE',  WTOB.getValue('COMPTECOMPTABLE')) ;
          TDATA.AddChampSupValeur('CAPPLIQUER',       WTOB.getValue('CAPPLIQUER')) ;
          TDATA.AddChampSupValeur('CDATE',            WTOB.getValue('CDATE')) ;
          TDATA.AddChampSupValeur('TDATE',            WTOB.getValue('TDATE')) ;
          TDATA.AddChampSupValeur('CMODEREG',         WTOB.getValue('CMODEREG')) ;
          TDATA.AddChampSupValeur('CKLIBCIB',         WTOB.getValue('CKLIBCIB')) ;
          TDATA.AddChampSupValeur('CCOMBINATOIRE',    WTOB.getValue('CCOMBINATOIRE')) ;
          TDATA.AddChampSupValeur('TNIVEAU',          WTOB.getValue('TNIVEAU')) ;
          TDATA.AddChampSupValeur('CNUMCHEQUE',       WTOB.getValue('CNUMCHEQUE')) ;
          TDATA.AddChampSupValeur('TNUMCHEQUE',       WTOB.getValue('TNUMCHEQUE')) ;
          TDATA.AddChampSupValeur('TPOSCHEQUE',       WTOB.getValue('TPOSCHEQUE')) ;
          TDATA.AddChampSupValeur('COMBOREF',         WTOB.getValue('COMBOREF')) ;
          TDATA.AddChampSupValeur('CBPARAM2',         WTOB.getValue('CBPARAM2')) ;
          TDATA.AddChampSupValeur('CBCASSE',          WTOB.getValue('CBCASSE')) ;
          TDATA.AddChampSupValeur('COMBOREFBQ',       WTOB.getValue('COMBOREFBQ')) ;
          TDATA.AddChampSupValeur('CBOKPOSBQ',        WTOB.getValue('CBOKPOSBQ')) ;
          TDATA.AddChampSupValeur('POSBQ1',           WTOB.getValue('POSBQ1')) ;
          TDATA.AddChampSupValeur('POSBQ2',           WTOB.getValue('POSBQ2')) ;
          TDATA.AddChampSupValeur('COMBOREFECR',      WTOB.getValue('COMBOREFECR')) ;
          TDATA.AddChampSupValeur('CBOKPOSECR',       WTOB.getValue('CBOKPOSECR')) ;
          TDATA.AddChampSupValeur('POSECR1',          WTOB.getValue('POSECR1')) ;
          TDATA.AddChampSupValeur('POSECR2',          WTOB.getValue('POSECR2')) ;
          result := True ;
        end ;
      end ;
    end ;
  end ;
  Rempli_CPA_ITT_DATA() ;
END ;

Function TITRTParam.GetData(): TOB;
var
  lStream:   TStringStream;
  lBoEntete: Boolean;
  lStDecode: String;
begin
  Result := Nil ;
  if ( TOBITRTPARAM.GetString('ITT_DATA') <> '') then
  begin
    Result := TOB.Create('', nil, -1);
    lStream := TStringStream.Create( TOBITRTPARAM.GetString('ITT_DATA') ) ;
    try
    Result.LoadFromXMLStream(lStream, lBoEntete, lStDecode );
    finally
      FreeAndNil( lStream ) ;
    end;
  end ;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetData(TOBData: TOB);
var
  lStream:  TStringStream;
begin
  lStream := TStringStream.Create( '' ) ;
  try
    TOBData.SaveToXmlStream( lStream, False, True );
    TOBITRTPARAM.SetString('ITT_DATA', lStream.DataString );
  finally
    lStream.Free;
    end;
end;

{------------------------------------------------------------------------------}
Procedure TITRTParam.SetData(theXMLChaine : String) ;
begin
  TOBITRTPARAM.SetString('ITT_DATA', theXMLChaine);
end ;

{------------------------------------------------------------------------------}
procedure TITRTParam.SetLikKey(theBLibKeyFg: Boolean);
begin
 bLibKeyFg := theBLibKeyFg ;
end;

function TITRTParam.Exists: Boolean;
var
  sql         : String ;
  ITT_USERGRP : String ;
  ITT_TYPPAR    : String ;
  ITT_CLE     : Integer ;
  ITT_VERSION : Integer ;
  ITT_USER    : String ;
begin
  ITT_USERGRP := GetUserGroupe() ; // Non traitée
  ITT_TYPPAR    := GetType() ;
  ITT_CLE     := GetCle() ; ;
  ITT_USER    := GetUser() ; ;
  ITT_VERSION := GetVersion() ; ;

  sql   := 'SELECT ITT_CLE FROM ITRTPARAM WHERE '
        + ' ITT_TYPPAR = "'       + ITT_TYPPAR         + '"'
        + ' AND ITT_USERGRP = "'  + ITT_USERGRP        + '"'
        + ' AND ITT_USER = "'     + ITT_USER           + '"'
        + ' AND ITT_CLE = '       + IntToStr(ITT_CLE)
        + ' AND ITT_VERSION = '   + IntToStr(ITT_VERSION)  ;
  Result := ExisteSQL(sql);
end;

end.
