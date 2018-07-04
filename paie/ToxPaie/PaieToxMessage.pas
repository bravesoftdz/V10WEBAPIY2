unit PaieToxMessage;

interface

uses DBTables,
     Hctrls,
     SysUtils,
     uTob ;

/////////////////////////////////////////////////////////
// Liste des fonctions permettant d'afficher un message
////////////////////////////////////////////////////////
function PaieMessageTOBTablePaie (TPST : TOB; MessErreur : integer; NomTable : String) : string ;
function PaieMessageTOBChoixcod     (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBChoixExt     (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBEtabliss     (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBLiensOLE     (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBModeles      (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBModeData     (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBPays         (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBTradDico     (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBTradTablette (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBTraduc       (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBUserGrp      (TPST : TOB; MessErreur : integer) : string ;
function PaieMessageTOBUtilisat     (TPST : TOB; MessErreur : integer) : string ;
function ToxEditeMessage        (NomTable : string; ToxInfo : TOB; MessErreur : integer) : string;

//////////////////////////////////////////////////
// Codes Erreurs
//////////////////////////////////////////////////
const

      EnregNotTrt      = -100 ;     // Enregistrement non traité
	    IntegOK          = 0  ;     	// Enregistrement OK
 	    TableNonReconnu  = 1  ;       // Table non reconnue et non gérée par la TOX

      ErrPays          = 38 ;				// Pays inexistant
      ErrLangue        = 41 ;       // Langue inexistante
      ErrEtablissement = 31 ;				// Etablissement inexistant
      ErrDico          = 81 ;       // Dictionnaire inexistant
      ErrCodeTablette  = 82 ;       // Code de tablette inexistant
      ErrForme         = 83 ;       // Forme inexistante
      ErrInsertDB            = 499 ;      // Erreur lors la mise en jour DB (InsertDB)
      ErrUpdateDB            = 500 ;      // Erreur lors la mise en jour DB (UpdateDB)

    	EnregNotTrtErr         = 1000 ;     // Enregistrement non traité car erreur dans traitement TOB tarnsactionnelle



implementation

Uses PaieToxIntegre ;

////////////////////////////////////////////////////////
//  Message TOB Table de la paie
////////////////////////////////////////////////////////
function  PaieMessageTOBTablePaie (TPST : TOB; MessErreur : integer; NomTable : String) : string ;
var T1      : TOB ;
    Cle,CC  : String ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur table '+ NomTable + ' n°' + IntToStr (MessErreur);
  end else
  begin
    T1 := PaieToxRendTablePaie (NomTable) ;
		result := 'L''enregistrement de la table ' + NomTable + ' a été traité : ';
    if Assigned (T1) then
      begin
        Cle := T1.GetValue ('DT_CLE1');
        Cle := StringReplace(Cle, ',', ';', [rfReplaceAll]) ;
        while Cle <> '' do
        begin
          CC := ReadTokenSt (Cle);
          result := result + TPST.GetValue (CC) ;
        end;
      end;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB Table de la TOX
////////////////////////////////////////////////////////
function  PaieMessageTOBTOX (TPST : TOB; MessErreur : integer; NomTable : String) : string ;
var T1      : TOB ;
    Cle,CC  : String ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur table '+ NomTable + ' n°' + IntToStr (MessErreur);
  end else
  begin
    T1 := PaieToxRendTablePaie (NomTable) ;
		result := 'L''enregistrement de la table ' + NomTable + ' a été traité : ';
    if Assigned (T1) then
      begin
        Cle := T1.GetValue ('DT_CLE1');
        Cle := StringReplace(Cle, ',', ';', [rfReplaceAll]) ;
        while Cle <> '' do
        begin
          CC := ReadTokenSt (Cle);
          result := result + TPST.GetValue (CC) ;
        end;
      end;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CHOIXCOD
////////////////////////////////////////////////////////
function  PaieMessageTOBChoixcod (TPST : TOB; MessErreur : integer) : string ;
var CCType       : string ;
    CCCode       : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur ChoixCod n°' + IntToStr (MessErreur);
  end else
  begin
  	CCType := TPST.GetValue ('CC_TYPE');
  	CCCode := TPST.GetValue ('CC_CODE');
		result := 'L''enregistrement de la table CHOIXCOD de type ' + CCType + ' et de code '+ CCCode;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB CHOIXEXT
////////////////////////////////////////////////////////
function  PaieMessageTOBChoixExt (TPST : TOB; MessErreur : integer) : string ;
var YXType       : string ;
    YXCode       : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur ChoixExt n°' + IntToStr (MessErreur);
  end else
  begin
  	YXType := TPST.GetValue ('YX_TYPE');
  	YXCode := TPST.GetValue ('YX_CODE');
		result := 'L''enregistrement de la table CHOIXEXT de type ' + YXType + ' et de code '+ YXCode;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB ETABLISS
////////////////////////////////////////////////////////
function  PaieMessageTOBEtabliss (TPST : TOB; MessErreur : integer) : string ;
var CodeEta  : string ;
		Libelle  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
    ///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
  	if MessErreur = ErrPays then
    begin
	    Probleme := TPST.GetValue ('ET_PAYS');
      result   := 'Le pays n''existe pas : "' + Probleme + '"';
      exit;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('ET_LANGUE');
      result   := 'La langue n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result   := 'Erreur Etablissement n°' + IntToStr (MessErreur);
  end else
  begin
  	CodeEta := TPST.GetValue ('ET_ETABLISSEMENT');
  	Libelle := TPST.GetValue ('ET_LIBELLE');
  	result  := 'La fiche de l''établissement "' + CodeEta + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB LIENSOLE
////////////////////////////////////////////////////////
function  PaieMessageTOBLiensOLE (TPST : TOB; MessErreur : integer) : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur LIENSOLE n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Le lien OLE ' + IntToStr(TPST.GetValue ('LO_RANGBLOB'));
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB MODELES
////////////////////////////////////////////////////////
function  PaieMessageTOBModeles (TPST : TOB; MessErreur : integer) : string ;
var Probleme : string ;
    Libelle  : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('MO_LANGUE');
      result   := 'La langue n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur MODELES n°' + IntToStr (MessErreur);
  end else
  begin
  	Libelle := TPST.GetValue ('MO_LIBELLE');
  	result  := 'Le modèle de document "' + Libelle + '"' ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB MODEDATA
////////////////////////////////////////////////////////
function  PaieMessageTOBModeData (TPST : TOB; MessErreur : integer) : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result := 'L''enregistrement a été intégré';
      exit ;
    end;
    result := 'Erreur MODEDATA n°' + IntToStr (MessErreur);
  end else
  begin
  	result := 'Le contenu du modèle de document' ;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB PAYS
////////////////////////////////////////////////////////
function  PaieMessageTOBPays (TPST : TOB; MessErreur : integer) : string ;
var CodePays : string ;
    Libelle  : string ;
    Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('PY_LANGUE');
      result   := 'La langue n''existe pas : "' + Probleme + '"';
      exit;
    end;
    result := 'Erreur PAYS n°' + IntToStr (MessErreur);
  end else
  begin
  	CodePays := TPST.GetValue ('PY_PAYS');
  	Libelle  := TPST.GetValue ('PY_LIBELLE');
  	result   := 'Le pays "' + CodePays + '" ' + Libelle;
  end;
end;


//
// Message TOB TRADDICO
//
function PaieMessageTOBTradDico     (TPST : TOB; MessErreur : integer) : string ;
var CodeDico,Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    if MessErreur = ErrDico then
    begin
	    Probleme := TPST.GetValue ('DX_FRA');
      result   := 'Le libellé n''existe pas : "' + Probleme + '"';
      exit;
    end
  end else
  begin
  	CodeDico := TPST.GetValue ('DX_FRA');
  	result   := 'Le dictionnaire "' + CodeDico + '" ' ;
  end;
end ;

//
// Message TRADTABLETTE
//
function PaieMessageTOBTradTablette (TPST : TOB; MessErreur : integer) : string ;
var sLangue, sTypeTablette, sCodeTablette ,Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end ;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('YTT_LANGUE');
      result   := 'La langue "'+ Probleme +'" n''existe pas.' ;
      exit;
    end ;
    if MessErreur = ErrCodeTablette then
    begin
	    Probleme := TPST.GetValue ('YTT_CODE');
      result   := 'La tablette de code "'+ Probleme + '" n''existe pas.' ;
      exit;
    end
  end else
  begin
  	sLangue       := TPST.GetValue ('YTT_LANGUE');
    sTypeTablette := TPST.GetValue ('YTT_TYPE') ;
    sCodeTablette := TPST.GetValue ('YTT_CODE') ;
  	result        := 'La tablette de type "'+sTypeTablette+'" de code "'+sCodeTablette+'" en langue "'+sLangue+'"' ;
  end;
end ;

//
// Message TRADUC
//
function PaieMessageTOBTraduc (TPST : TOB; MessErreur : integer) : string ;
var sLangue, sForme, Probleme : string ;
begin
	if MessErreur >= 0 then
  begin
  	//
    // Gestion des messages d'erreur
    //
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end ;
    if MessErreur = ErrLangue then
    begin
	    Probleme := TPST.GetValue ('TR_LANGUE');
      result   := 'La langue "'+ Probleme + '" n''existe pas.' ;
      exit;
    end ;
    if MessErreur = ErrForme then
    begin
	    Probleme := TPST.GetValue ('TR_FORME');
      result   := 'La fiche "'+ Probleme +'" n''existe pas.' ;
      exit;
    end
  end else
  begin
  	sLangue := TPST.GetValue ('TR_LANGUE');
    sForme  := TPST.GetValue ('TR_FORME') ;
  	result  := 'La fiche "'+sForme+'" en "'+sLangue+'"' ;
  end;
end ;

////////////////////////////////////////////////////////
//  Message TOB USERGRP
////////////////////////////////////////////////////////
function  PaieMessageTOBUserGrp (TPST : TOB; MessErreur : integer) : string ;
var	Groupe 	: string ;
		Libelle : string ;

begin
	if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur USERGRP n°' + IntToStr (MessErreur);
  end else
  begin
  	Groupe  := TPST.GetValue ('UG_GROUPE');
  	Libelle := TPST.GetValue ('UG_LIBELLE');
  	result  := 'Le groupe d''utilisateur "' + Groupe + '" ' + Libelle;
  end;
end;

////////////////////////////////////////////////////////
//  Message TOB UTILISAT
////////////////////////////////////////////////////////
function  PaieMessageTOBUtilisat (TPST : TOB; MessErreur : integer) : string ;
var	Utilisateur	: string ;
	Libelle     : string ;

begin
  if MessErreur >= 0 then
  begin
  	///////////////////////////////////////////////////////////////////////////
    // Gestion des messages d'erreur
    ///////////////////////////////////////////////////////////////////////////
    if MessErreur = IntegOK then
    begin
      result   := 'L''enregistrement a été intégré';
      exit ;
    end;
    result   := 'Erreur UTILISAT n°' + IntToStr (MessErreur);
  end else
  begin
  	Utilisateur := TPST.GetValue ('US_UTILISATEUR');
  	Libelle     := TPST.GetValue ('US_LIBELLE');
  	result      := 'L''utilisateur "' + Utilisateur + '" ' + Libelle;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Descriptif des informations intégrées
///////////////////////////////////////////////////////////////////////////////////////
function ToxEditeMessage (NomTable : string; ToxInfo : TOB; MessErreur : integer) : string;
var MessageInfo,SQL : string;
    Q               : TQuery ;
begin
  if MessErreur = EnregNotTrt then
  begin
    MessageInfo := 'L''enregistrement n''a pas encore été traité' ;
    exit ;
  end;
  // Identification de la table par rapport à une table de la paie mais non DP+STD
  SQL := 'SELECT DT_NOMTABLE FROM DETABLES WHERE DT_DOMAINE = "P" '+
         ' AND DT_NOMTABLE="' + NomTable + '"';
  Q:=OPENSQL(SQL,TRUE) ;
  if NOT Q.EOF then
  begin
    FERME (Q) ;
    RESULT := PaieMessageTOBTablePaie (ToxInfo, MessErreur, NomTable) ;
    exit ;
  end;
  Ferme (Q) ;

  if MessErreur = EnregNotTrtErr then MessageInfo := 'Une erreur a été détectée dans l''unité transactionnelle'
  else if (NomTable = 'CHOIXCOD')	  then MessageInfo := PaieMessageTOBChoixCod     (ToxInfo, MessErreur)
  else if (NomTable = 'CHOIXEXT')	  then MessageInfo := PaieMessageTOBChoixExt     (ToxInfo, MessErreur)
  else if (NomTable = 'ETABLISS')	  then MessageInfo := PaieMessageTOBEtabliss     (ToxInfo, MessErreur)
  else if (NomTable = 'LIENSOLE')	  then MessageInfo := PaieMessageTOBLiensOLE     (ToxInfo, MessErreur)
  else if (NomTable = 'MODELES')	  then MessageInfo := PaieMessageTOBModeles      (ToxInfo, MessErreur)
  else if (NomTable = 'MODEDATA')	  then MessageInfo := PaieMessageTOBModeData     (ToxInfo, MessErreur)
  else if (NomTable = 'PAYS')	      then MessageInfo := PaieMessageTOBPays         (ToxInfo, MessErreur)
  else if (NomTable = 'TRADDICO')	  then MessageInfo := PaieMessageTOBTradDico     (ToxInfo, MessErreur)
  else if (NomTable = 'TRADTABLETTE') then MessageInfo := PaieMessageTOBTradTablette (ToxInfo, MessErreur)
  else if (NomTable = 'TRADUC')	      then MessageInfo := PaieMessageTOBTraduc       (ToxInfo, MessErreur)
  else if (NomTable = 'USERGRP')   	  then MessageInfo := PaieMessageTOBUserGrp      (ToxInfo, MessErreur)
  else if (NomTable = 'UTILISAT')  	  then MessageInfo := PaieMessageTOBUtilisat     (ToxInfo, MessErreur)
  else if (Copy (NomTable,1,4) = 'STOX') then MessageInfo := PaieMessageTOBTOX     (ToxInfo, MessErreur, NomTable)
  else MessageInfo:='L''intégration de la table '+ Nomtable + ' n''est pas gérée' ;

  result:=MessageInfo;
end;

end.

