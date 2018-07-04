unit UtilSaisieConso;

interface

Uses HEnt1, UTOB, Ent1, LookUp,
     SysUtils, UtilPGI, AGLInit,HCtrls,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
{$ENDIF}
{$IFDEF BTP}
	   CalcOLEGenericBTP,
{$ENDIF}
     AppelsUtil,
     UtilConso,
     forms,
     EntGC, Classes, HMsgBox,
     ParamSoc,UPlannifchUtil,factUtil,FactTOB,uEntCommun,UtilTOBPiece;

type TGrilleModeSaisie = (TgsMO, TgsFRS, TgsRES, TgsFOURN,TgsMoext,TgsRecettes,TgsNone);
     TGncERROR = ( GncOk,GncExists,GncAbort);
     TUModeValo = (TmvAll,TmvPa,TmvPR,TmvPV);
     Tmodevalo = set of TUModeValo;
procedure AddLesSupLignesConso (TOBL : TOB);
function GetFonction (Code : string ; ToutPermis : boolean=true) : string;
function GetMainDoeuvre (Code : string) : string;
function OKTypeHeure (TOBConso: TOB;Thevaleur:string;TOBTypeHeure : TOB) : boolean;
function OkChantier (Chantier : string) : boolean;
function OkMATERIAUX (MATERIAUX : string) : boolean;
function OKRessource (TOBConso : TOB;MAINDOEUVRE,LeType : string; Var PrestationDefaut : string;WithMajPrix : boolean; fonction : string='';  WithRecupPrestation : boolean = true) : single;
function OKRessourceMAT (TOBCOnso : TOB; MAINDOEUVRE: string;WithMajPrix : boolean;var PrestationDefaut : string; fonction : string='') : boolean;
function OKRessourceExterne (TOBCOnso : TOB; MAINDOEUVRE: string; WithMajPrix: boolean;var PrestationDefaut : string; fonction : string='') : boolean;
function IsArticleKnown (Code: string; TOBL:TOB; FromArt : boolean; AvecChargement : Boolean=True) : boolean;
function IsPrestationKnown (Code : string; TOBL : TOB; FromArt : boolean; Mode : TGrilleModeSaisie; AvecChargement : Boolean=True) : boolean;
function IsFraisKnown (Code : string; TOBL:TOB; fromArt : boolean; AvecChargement : Boolean=True) : boolean;
procedure SetInfoArticle (TOBL,TOBRESSOURCE : TOB; Article : string; var PrestationDefaut : string;WithMajPrixRessource : boolean=True);
procedure SetInfoRessource (TOBL,TOBres : TOB; var WithMajPrix : boolean; WithRecupPrestation : boolean; var PrestationDefaut : string);
procedure CalculeAndAfficheLaLigne (Grille : THGrid ;lesChamps : string; Ligne : integer; TOBS : TOB);
function IsExisteAffaire (Code : string; Var Domaine,LibAffaire : String) : boolean;
function IsExistePhaseAffaire (Code,Phase : string) : boolean;
function GetChantier (Code : string; ChangeStatut:Boolean=false) : string;
function GetLibelleMateriaux (Code : string) : string;
function GetTypeRessource (CodeRessource : string) : string;
procedure calculeLaLigne (TOBL : TOB; coefPaPr : double = 0; CoefPrPv : double= 0);
procedure AfficheLaLigne (TOBL : TOB; Grille : THgrid; LesChamps : string; Arow : integer);
function LigneFromPieces (TOBL : TOB) : boolean;
function LigneDejaValide (TOBL : TOB) : boolean;
// Fonction permettant de récupérer un identifiant unique de consommation
function GetNumUniqueConso (var TheResult : double) : TGncERROR;
//
procedure RecupCoefs(TOBLigne: TOB; venteAchat : string; var CoefpaPr, CoefPrPv: double);
procedure DefiniInfoArticle (TOBL : TOb; Article : string ; RecupValorisation : boolean); overload;
procedure DefiniInfoArticle (TOBL,TOBA : TOb; RecupValorisation : boolean); overload;
procedure SetValoFromRessource (TOBL: TOB;TOBRes : TOB;ModeValo : Tmodevalo=[TmvAll];Coefmajoration : double=1);
procedure SetvaloArticle (TOBL,TOBA : TOB; ModeValo : TModevalo=[TmvAll]);
function GetNatureMouv (TypeArticle,TypeRessource : string) : string;
procedure GetPVPrestation (TOBL : TOB);
procedure GetValoPrestation(TOBL : TOB);
function GetprixFromBordereau (CodeTiers,CodeArticle,CodeAffaire : string; var Prixvente : double) : boolean;
function getTiersAffaire(CodeAffaire : string) : string;
function GetFamilleTaxe(Article : string) : string;
function ConstitueDateDebutTravaux (TOBL : TOB) : TDateTime;

//FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
Procedure AppliqueCoeff(TOBL : TOB);

implementation
uses Utilsuggestion,factdomaines,DateUtils,UDateUtils,UspecifPOC;

function GetFonction (Code : string ; ToutPermis : boolean=true) : string;
var QQ : TQuery;
    Sql : String;
begin
  if ToutPermis then result := TraduireMemoire('Toutes Les fonctions')
                else result := '';
  Sql := 'SELECT AFO_LIBELLE FROM FONCTION WHERE AFO_FONCTION="'+
          Code+'"';
  QQ := OpenSql (SQL,true);
  if not QQ.eof then
  begin
    result := QQ.findfield('AFO_LIBELLE').asString ;
  end;
  ferme (QQ);
end;

function GetMainDoeuvre (Code : string) : string;
var QQ : TQuery;
    Sql : String;
begin
  result := '';
  //FV1 : 26/08/2013 - FS#564 - VERRE & METAL - Saisie de consommations / ressource de type matériel fermée
  //Sql := 'SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+Code+'"';
  Sql := 'SELECT ARS_LIBELLE, ARS_FERME FROM RESSOURCE WHERE ARS_RESSOURCE="'+Code+'"';
  QQ := OpenSql (SQL,true);
  if not QQ.eof then
  begin
    if QQ.FindField('ARS_FERME').AsString <> 'X' then
      result := QQ.findfield('ARS_LIBELLE').asString
    else
      PGIError('Saisie impossible cette ressource est fermée !','Ressources')
  end
  Else
  begin
    PGIError('Saisie impossible cette ressource est Inexistante !','Ressources')
  end;

  ferme (QQ);
end;

function OkChantier (Chantier : string) : boolean;
var QQ    : Tquery;
    Etat  : String;
begin
  result := false;
  QQ := OpenSql ('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Chantier+'"',true);
  if not QQ.eof then result  := true;
  ferme (QQ);
end;

function OkMATERIAUX (MATERIAUX : string) : boolean;
var QQ : Tquery;
begin
  result := true;

  if MATERIAUX = '' then exit;

  result := false;
  QQ := OpenSql ('SELECT GA_LIBELLE, GA_FERME, GA_DATESUPPRESSION FROM ARTICLE WHERE GA_ARTICLE="'+MATERIAUX+'"',true);

  if not QQ.eof then
  begin
    if QQ.FindField('GA_FERME').AsString <> 'X' then
      result := True
    else
    begin
      PGIError('Saisie impossible cet article est fermé !','Matériaux');
      Result := False;
    end;
  end
  Else
  begin
    PGIError('Saisie impossible cet article est Inexistant !','Matériaux');
    Result :=  false;
  end;
  ferme (QQ);
end;

function OKTypeHeure (TOBConso: TOB;Thevaleur:string;TOBTypeHeure : TOB) : boolean;
var QQ : Tquery;
    SQl : String;
    Coef,Coeffg,CoefMarg : Double;
    LaValeur : string;
    LaTOb : TOB;
begin

	result := false;
  coef := 1;

  if TOBCOnso.getDouble('BCO_DPA') <> 0 then
    CoeffG := TOBCOnso.getDouble('BCO_DPR')/TOBCOnso.getDouble('BCO_DPA')
  else
    CoeffG := 1;

  if TOBCOnso.getDouble('BCO_DPR') <> 0 then
    CoefMarg := TOBCOnso.getDouble('BCO_PUHT')/TOBCOnso.getDouble('BCO_DPR')
  else
    CoefMarg := 1;

  //if TheValeur = '' then TheValeur := 'NOR';
  if (TobTypeHeure <> nil) and (TheValeur <> '') then LATOB := TOBTypeHeure.findFirst(['CC_CODE'],[TheValeur],true);

  if (LaTOB = nil) and (TheValeur <> '') then
  begin
    Sql := 'SELECT CC_LIBELLE,CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="ATH" AND CC_CODE="'+TheValeur+'"';
    QQ := OpenSql (SQl,true);
  	if not QQ.eof then
    begin
    	LaTOB := TOB.Create ('CHOIXCOD',TOBTypeheure,-1);
    	LaTOB.selectdb ('',QQ);
      result := true;
    end;
  	ferme (QQ);
  end else result := true;


  if (result) and (TheValeur <> '') then
  begin
    TOBCONSO.PutValue('LIBELLETYPE',LaTOB.GetValue('CC_LIBELLE'));
    LaValeur :=  LaTOB.GetValue ('CC_ABREGE');
    if (LaValeur <> '') and (IsNumeric (LaValeur)) then
    begin
      Coef := 1 + VALEUR(laValeur)/100;
    end;
  end;

  if (result) (*and (Coef > 0) modif BRL 18/05/05*) then
  begin
    IF TOBCONSO.GetValue ('BCO_TYPEHEURE') <> TheValeur Then TOBCONSO.PutValue ('MODIF', 'X');
  	TOBCONSO.putvalue('BCO_TYPEHEURE',TheValeur);
    if TheValeur <> '' then
       TOBCONSO.PutValue('LIBELLETYPE',LaTOB.GetValue('CC_LIBELLE'))
    else
       TOBCONSO.PutValue('LIBELLETYPE','');

	  TOBConso.PutValue('BCO_DPA',Arrondi(TOBConso.GetValue('_PA_INIT')*Coef,V_PGI.okdecP));
 	  TOBConso.PutValue('BCO_DPR',Arrondi(TOBConso.GetValue('BCO_DPA')*Coeffg,V_PGI.okdecP));
 	  TOBConso.PutValue('BCO_PUHT',Arrondi(TOBConso.GetValue('BCO_DPR')*CoefMarg,V_PGI.okdecP));
    TOBConso.SetDouble('COEFMARGE',CoefMarg);
	  calculeLaLigne (TOBCONSO);
  end;
end;

function OKRessource (TOBConso : TOB;MAINDOEUVRE,LeType : string; Var PrestationDefaut : string;WithMajPrix : boolean; fonction : string='';  WithRecupPrestation : boolean = true) : single;
var QQ : Tquery;
    SQl : String;
    TOBRes : TOB;
begin

	TOBRES := TOB.Create ('RESSOURCE',nil,-1);

  result := 0;

  if MAINDOEUVRE = '' then exit;

  Sql := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+ MAINDOEUVRE +'"';

  if (LeType <> '""') and (LeType<>'') then Sql := Sql + ' AND ARS_TYPERESSOURCE IN ('+LeType+')';

  if fonction <> '' then Sql := Sql + ' AND ARS_FONCTION1="'+fonction+'"';

  QQ := OpenSql (SQL,true);

  if not QQ.eof then
  begin
    result := 1;
    if QQ.FindField('ARS_FERME').AsString = 'X' then
      result := 2
    else
    begin
      if (TOBCOnso <> nil) then
      begin
        TOBREs.SelectDB ('',QQ,true);
        TOBConso.putValue('BCO_RESSOURCE',trim(MAINDOEUVRE));
        SetInfoRessource (TOBConso,TOBres,WithMajPrix,WithRecupPrestation,PrestationDefaut);
      end
      else
        PrestationDefaut:=QQ.findField ('ARS_ARTICLE').asString;
    end;
  end;

  TOBRES.free;

  ferme (QQ);

end;

function OKRessourceMAT (TOBCOnso : TOB; MAINDOEUVRE: string;WithMajPrix : boolean;var PrestationDefaut : string; fonction : string='') : boolean;
var QQ      : Tquery;
    SQl     : String;
    TOBREs  : TOB;
    Stwhere : String;
begin
	TOBREs := TOB.Create ('RESSOURCE',nil,-1);

  result := false;

  if MAINDOEUVRE = '' then exit;

  Sql := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+MAINDOEUVRE+'"';
  Sql := Sql + ' AND (ARS_TYPERESSOURCE="MAT" OR ARS_TYPERESSOURCE="OUT")';

  if fonction <> '' then
    Sql := Sql + ' AND ARS_FONCTION1="'+fonction+'"';

  QQ := OpenSql (SQL,true);

  if Not QQ.eof then
  begin
    if QQ.FindField('ARS_FERME').AsString <> 'X' then
      result := True
    else
    begin
      PGIError('Saisie impossible cette Ressource est fermée !','Ressources');
    end;
  end
  Else
  begin
    PGIError('Saisie impossible cette Ressource est Inexistante !','Ressources');
  end;

  if result then
  begin
  	if (TOBCOnso <> nil) then
    begin
      TOBREs.SelectDB ('',QQ,true);
      TOBConso.putValue('BCO_RESSOURCE',trim(MAINDOEUVRE));
      if TOBres.GetValue('ARS_ARTICLE')<> '' then
         begin
         TOBConso.putValue('BCO_ARTICLE',TOBres.GetValue('ARS_ARTICLE'));
         TOBConso.putValue('BCO_CODEARTICLE',Copy(TOBres.GetValue('ARS_ARTICLE'),1,18));
         SetInfoArticle (TOBConso,TOBres,TOBConso.GetValue('BCO_ARTICLE'),PrestationDefaut,WithMajPrix);
         end
      else
         SetInfoRessource (TOBConso,TOBres,WithMajPrix,True,PrestationDefaut);
    end;
  end;

  TOBRES.free;

  ferme (QQ);
end;

function OKRessourceExterne (TOBCOnso : TOB; MAINDOEUVRE: string; WithMajPrix: boolean;var PrestationDefaut : string; fonction : string='') : boolean;
var QQ : Tquery;
    SQl : String;
    TOBREs : TOB;
begin
	TOBREs := TOB.Create ('RESSOURCE',nil,-1);
  result := false;
  if MAINDOEUVRE = '' then exit;
  Sql := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+MAINDOEUVRE+'"';
  Sql := Sql + ' AND (ARS_TYPERESSOURCE="INT" OR ARS_TYPERESSOURCE="ST" OR ARS_TYPERESSOURCE="LOC" OR ARS_TYPERESSOURCE="AUT")';
  if fonction <> '' then
    Sql := Sql + ' AND ARS_FONCTION1="'+fonction+'"';
  QQ := OpenSql (SQL,true);

  if not QQ.eof then
  begin
    if QQ.FindField('ARS_FERME').AsString <> 'X' then
    Begin
      //FV1 - 20/10/2017- FS#2647 - EMC - nom ressource incorrecte sur onglet prestations externes
      TOBREs.SelectDB ('',QQ,true);
      result := True;
    end
    else
    begin
      PGIError('Saisie impossible cet article est fermé !','Matériaux');
      Result := False;
    end;
  end
  Else
  begin
    PGIError('Saisie impossible cet article est Inexistant !','Matériaux');
    Result :=  false;
  end;
  ferme (QQ);

  if result then
  begin
  	if (TOBCOnso <> nil) then
    begin
      //FV1 - 20/10/2017- FS#2647 - EMC - nom ressource incorrecte sur onglet prestations externes
      //TOBREs.SelectDB ('',QQ,true);
      TOBConso.putValue('BCO_RESSOURCE',trim(MAINDOEUVRE));
      SetInfoRessource (TOBConso,TOBres,WithMajPrix,TRue,PrestationDefaut);
    end;
  end;
  TOBRES.free;
  ferme (QQ);
end;

function ControleVideSuivantMode (TOBL : TOB) : boolean;
begin
  result := false;
  if (TOBL.getValue('BCO_AFFAIRE') = '') or (TOBL.GetValue('BCO_PHASETRA')='') then exit;
  if TOBL.GetValue('BCO_NATUREMOUV')= 'MO' then
  begin
    if (TOBL.getValue('BCO_RESSOURCE') = '') then exit;
  end else if TOBL.GetValue('BCO_NATUREMOUV')= 'FRS' then
  begin
    if (TOBL.getValue('BCO_ARTICLE') = '') then exit;
  end else if TOBL.GetValue('BCO_NATUREMOUV')= 'RES' then
  begin
    if (TOBL.getValue('BCO_RESSOURCE') = '') then exit;
  end else if TOBL.GetValue('BCO_NATUREMOUV')= 'EXT' then
  begin
    if (TOBL.getValue('BCO_RESSOURCE') = '') then exit;
  end else if TOBL.GetValue('BCO_NATUREMOUV')= 'FOU' then
  begin
    if (TOBL.getValue('BCO_ARTICLE') = '') then exit;
  end;
  result := true;
end;

function IsArticleKnown (Code: string; TOBL:TOB; FromArt : boolean; AvecChargement : Boolean=True) : boolean;
var QQ 					: Tquery;
    Requete			: String;
    Article     : String;
    PrixVente 	: Double;
    PrixRevient	: Double;
    CoefMarge		: Double;
Begin

  result := false;

  if Code = '' then exit;

  Requete := 'SELECT GA_STATUTART,GA_CODEARTICLE,GA_ARTICLE,' +
             'GA_LIBELLE, GA_PVHT, GA_DPR,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3 ' +
             'FROM ARTICLE WHERE ';

  if FromArt then
    QQ := OpenSql(Requete + 'GA_ARTICLE="'+Code+'" AND GA_TYPEARTICLE IN ("ARP","MAR")',true)
  else
    QQ := OpenSql(Requete + 'GA_CODEARTICLE="'+Code+'" AND GA_TYPEARTICLE IN ("ARP","MAR") AND GA_STATUTART IN ("GEN","UNI")',true);

	// Article dimensionné
  if (not QQ.eof) and (not fromart)then
	   begin
     if (QQ.findField ('GA_STATUTART').asString = 'GEN') then
        Article := SelectUneDimension (QQ.findField ('GA_ARTICLE').asString)
     else
	      Article := QQ.findField ('GA_ARTICLE').asString;
     ferme (QQ);
     QQ := OpenSql (Requete + 'GA_ARTICLE="'+ Article +'" AND GA_TYPEARTICLE IN ("ARP","MAR")',true);
     end;

  if QQ.eof then
     Begin
     ferme (QQ);
     exit;
     end;

  if AvecChargement then
  begin
		PrixVente 	:= QQ.findField('GA_PVHT').AsFloat;
		PrixRevient := QQ.findField('GA_DPR').AsFloat;
  	if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;

  	TOBL.PutValue('BCO_ARTICLE',QQ.findField('GA_ARTICLE').asString);
  	TOBL.PutValue('BCO_CODEARTICLE',QQ.findField('GA_CODEARTICLE').asString);
		TOBL.PUTVALUE('BCO_LIBELLE',QQ.findField('GA_LIBELLE').AsString);
  	TOBL.PutValue('COEFMARGE', CoefMarge);
    TOBL.PUTVALUE('BCO_FAMILLENIV1',QQ.findField('GA_FAMILLENIV1').AsString);
    TOBL.PUTVALUE('BCO_FAMILLENIV2',QQ.findField('GA_FAMILLENIV2').AsString);
    TOBL.PUTVALUE('BCO_FAMILLENIV3',QQ.findField('GA_FAMILLENIV3').AsString);
  end;

  result := true;

  ferme (QQ);

end;

function IsPrestationKnown (Code : string; TOBL : TOB; FromArt : boolean; Mode : TGrilleModeSaisie; AvecChargement : Boolean=True) : boolean;
var QQ 					: Tquery;
    requete 		: String;
    PrixVente 	: Double;
    PrixRevient	: Double;
    CoefMarge		: Double;
begin

  result := false;

  if Code = '' then exit;

  requete := 'SELECT GA_CODEARTICLE,GA_ARTICLE,GA_LIBELLE,' +
             'GA_PVHT, GA_DPR,BNP_TYPERESSOURCE,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3 FROM ARTICLE '+
             'LEFT JOIN NATUREPREST ON GA_NATUREPRES=BNP_NATUREPRES '+
             'WHERE ';

  if FromArt then
     requete := requete + 'GA_ARTICLE="'+Code+'" AND GA_TYPEARTICLE="PRE"'
  else
     requete := requete + 'GA_CODEARTICLE="'+Code+'" AND GA_TYPEARTICLE="PRE"';

  if Mode = TgsMo then
     requete := requete + ' AND BNP_TYPERESSOURCE IN ("SAL","INT")'
  else if Mode = TgsRes then
  	 requete := requete + ' AND ((BNP_TYPERESSOURCE="MAT") OR (BNP_TYPERESSOURCE="OUT"))'
  else if Mode = TgsMOEXt then
  	 begin
     requete := requete + ' AND ((BNP_TYPERESSOURCE="INT") OR (BNP_TYPERESSOURCE="ST")';
     requete := requete + ' OR (BNP_TYPERESSOURCE="LOC") OR (BNP_TYPERESSOURCE="AUT"))';
	   end;

  QQ := OpenSql (requete,true);

  if QQ.eof then
     Begin
     ferme (QQ);
     exit;
     end;

  if AvecChargement then
  begin
		PrixVente 	:= QQ.findField('GA_PVHT').AsFloat;
		PrixRevient := QQ.findField('GA_DPR').AsFloat;
  	if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;

  	TOBL.PutValue('BCO_ARTICLE',QQ.findField('GA_ARTICLE').asString);
  	TOBL.PutValue('BCO_CODEARTICLE',QQ.findField('GA_CODEARTICLE').asString);
  	TOBL.PUTVALUE('BCO_LIBELLE',QQ.findField('GA_LIBELLE').AsString);
  	TOBL.PutValue('COEFMARGE', CoefMarge);
    TOBL.PUTVALUE('BCO_FAMILLENIV1',QQ.findField('GA_FAMILLENIV1').AsString);
    TOBL.PUTVALUE('BCO_FAMILLENIV2',QQ.findField('GA_FAMILLENIV2').AsString);
    TOBL.PUTVALUE('BCO_FAMILLENIV3',QQ.findField('GA_FAMILLENIV3').AsString);
	end;
  result := true;

  ferme (QQ);

end;

function IsFraisKnown (Code : string; TOBL : TOB; fromArt : boolean; AvecChargement : Boolean=True) : boolean;
var QQ 					: Tquery;
    Requete			: String;
	  PrixVente 	: Double;
    PrixRevient	: Double;
    CoefMarge		: Double;
begin

  result := false;

  if Code = '' then exit;

  Requete := 'SELECT GA_STATUTART,GA_CODEARTICLE,GA_ARTICLE,' +
             'GA_LIBELLE, GA_PVHT, GA_DPR,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3 ' +
             'FROM ARTICLE WHERE ';

  if FromArt then
  	 QQ := OpenSql (Requete + 'GA_ARTICLE="'+Code+'" AND GA_TYPEARTICLE="FRA"',true)
  else
     QQ := OpenSql (Requete + 'GA_CODEARTICLE="'+Code+'" AND GA_TYPEARTICLE="FRA"',true);

  if QQ.eof then      Begin
     ferme (QQ);
     exit;
     end;

  if AvecChargement then
  begin
		PrixVente 	:= QQ.findField('GA_PVHT').AsFloat;
		PrixRevient := QQ.findField('GA_DPR').AsFloat;
  	if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;

    TOBL.PutValue('BCO_ARTICLE',QQ.findField('GA_ARTICLE').asString);
  	TOBL.PutValue('BCO_CODEARTICLE',QQ.findField('GA_CODEARTICLE').asString);
  	TOBL.PUTVALUE('BCO_LIBELLE',QQ.findField('GA_LIBELLE').AsString);
  	TOBL.PutValue('COEFMARGE', 0);
    TOBL.PUTVALUE('BCO_FAMILLENIV1',QQ.findField('GA_FAMILLENIV1').AsString);
    TOBL.PUTVALUE('BCO_FAMILLENIV2',QQ.findField('GA_FAMILLENIV2').AsString);
    TOBL.PUTVALUE('BCO_FAMILLENIV3',QQ.findField('GA_FAMILLENIV3').AsString);
	end;

  result := true;

  ferme (QQ);

end;

procedure AlimenteValeurInit (TOBL : TOB);
begin
  TOBL.PutValue ('_PA_INIT',TOBL.GetValue('BCO_DPA'));
  TOBL.PutValue ('_PR_INIT',TOBL.GetValue('BCO_DPR'));
  TOBL.PutValue ('_PV_INIT',TOBL.GetValue('BCO_PUHT'));
end;

function TypePrestationAutorise(Ressource : string) : boolean;
var QQ : Tquery;
begin

end;

procedure DefiniInfoArticle (TOBL : TOb; Article : string ; RecupValorisation : boolean);
var QQ 					: Tquery;
		TOBA 				: TOB;
    PrixVente 	: Double;
    PrixRevient	: Double;
    CoefMarge		: Double;
begin

	TOBA := TOB.Create ('ARTICLE',nil,-1);

  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true);

  if (not QQ.eof) then
  	 begin
  	 TOBA.SelectDB ('',QQ);
     DefiniInfoArticle(TOBL,TOBA,Recupvalorisation);
     //chargement des valeur numéraire
     PrixVente 	:= QQ.findField('GA_PVHT').AsFloat;
		 PrixRevient := QQ.findField('GA_DPR').AsFloat;
	   if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;
     //Chargement de la tobconso
     TOBL.PutValue('BCO_ARTICLE',QQ.findField('GA_ARTICLE').asString);
     TOBL.PutValue('BCO_CODEARTICLE',QQ.findField('GA_CODEARTICLE').asString);
     TOBL.PUTVALUE('BCO_LIBELLE',QQ.findField('GA_LIBELLE').AsString);
     TOBL.PUTVALUE('BCO_FAMILLETAXE1',QQ.findField('GA_FAMILLETAXE1').AsString);
     TOBL.PUTVALUE('PRIXPASMODIF',QQ.findField('GA_PRIXPASMODIF').AsString);
     TOBL.PUTVALUE('BCO_FAMILLENIV1',QQ.findField('GA_FAMILLENIV1').AsString);
     TOBL.PUTVALUE('BCO_FAMILLENIV2',QQ.findField('GA_FAMILLENIV2').AsString);
     TOBL.PUTVALUE('BCO_FAMILLENIV3',QQ.findField('GA_FAMILLENIV3').AsString);
     //
     TOBL.PutValue('COEFMARGE', CoefMarge);
(*
  	if RecupValorisation then TOBL.putValue('BCO_QUALIFQTEMOUV',TOBA.GetValue('GA_QUALIFUNITEVTE'));
    TOBL.putValue('BCO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
    if RecupValorisation then
    begin
    	ReactualisePrPv (TOBA);
      TOBL.putValue('BCO_DPA',TOBA.GetValue('GA_PAHT') );
      TOBL.putValue('BCO_DPR',TOBA.GetValue('GA_DPR') );
      TOBL.putValue('BCO_PUHT',TOBA.GetValue('GA_PVHT'));
    end;
  	if RecupValorisation then AlimenteValeurInit (TOBL);
*)
  end;

  ferme (QQ);
  TOBA.free;
end;

procedure SetvaloArticle (TOBL,TOBA : TOB; ModeValo : TModevalo=[TmvAll]);
var fournisseur,Article : string;
		Ua : string;
    CoefuaUs : double;
begin
	Fournisseur := TOBA.GetValue('GA_FOURNPRINC');
  Article := TOBA.GetValue('GA_ARTICLE');
  RecupTarifAch (Fournisseur,Article,Ua,CoefUaUs,TurVente,true,True,TOBA);
  ReactualisePrPv (TOBA);
  if (TmvAll in ModeValo) or (TmvPa in ModeValo) then TOBL.putValue('BCO_DPA',TOBA.GetValue('GA_PAHT') );
  if (TmvAll in ModeValo) or (TmvPr in ModeValo) then TOBL.putValue('BCO_DPR',TOBA.GetValue('GA_DPR') );
  if (TmvAll in ModeValo) or (TmvPv in ModeValo) then TOBL.putValue('BCO_PUHT',TOBA.GetValue('GA_PVHT'));
end;

procedure DefiniInfoArticle (TOBL,TOBA : TOb; RecupValorisation : boolean);
var QQ : Tquery;
begin
  if RecupValorisation then TOBL.putValue('BCO_QUALIFQTEMOUV',TOBA.GetValue('GA_QUALIFUNITEVTE'));
  TOBL.putValue('BCO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
  if RecupValorisation then
  begin
  	SetvaloArticle (TOBL,TOBA);
  end;
  if RecupValorisation then AlimenteValeurInit (TOBL);
end;

procedure SetInfoArticle (TOBL,TOBRESSOURCE : TOB; Article : string; var PrestationDefaut : string;WithMajPrixRessource : boolean=True);
var QQ : Tquery;
		RecupValorisation : boolean;
    Ressource : string;
begin
	RecupValorisation := true;
  if Article = '' then exit;
  Ressource := TOBL.GetValue('BCO_RESSOURCE');
  if (Ressource <> '') and (TOBRessource <> nil) then
  begin
  	SetInfoRessource (TOBL,TOBREssource,WithMajPrixRessource,false,PrestationDefaut);
  	RecupValorisation := not WithMajPrixRessource; // si ressource renseignée alors pas de récup de prix
  end;
  DefiniInfoArticle (TOBL,Article,RecupValorisation);
end;

procedure SetValoFromRessource (TOBL: TOB;TOBRes : TOB;ModeValo : Tmodevalo=[TmvAll];Coefmajoration : double=1);
var PrixVente : Double;
    CoefMarge	: Double;
    PVCalcul	: Double;
    PxRevient : Double;
begin

	if TOBres = nil then exit;

  PrixVente := 0;
  CoefMarge := 0;
  PVCalcul  := 0;
  PxRevient := 0;

  if (TmvAll in ModeValo) or (TmvPa in ModeValo) then
     TOBL.putValue('BCO_DPA',TOBREs.GetValue('ARS_TAUXUNIT')*CoefMajoration);

  if (TmvAll in ModeValo) or (TmvPR in ModeValo) then
  begin
     TOBL.putValue('BCO_DPR',TOBREs.GetValue('ARS_TAUXREVIENTUN'));
     if TOBL.GetValue('BCO_DPR') = 0 then TOBL.putValue('BCO_DPR',TOBL.GetValue('BCO_DPA'));
     TOBL.putValue('BCO_DPR',TOBL.GetValue('BCO_DPR')*CoefMajoration);
  end;

  //Chargement des valorisations
  PVCalcul  := TOBRES.GEtValue('ARS_PVHTCALCUL');
  PrixVente := TOBRES.GetValue('ARS_PVHT');
  CoefMarge := TOBL.GEtDouble('COEFMARGE');
  PxRevient := TOBRES.GetValue('ARS_TAUXREVIENTUN');

  if (TmvAll in ModeValo) or (TmvPV in ModeValo) then
     begin
     if (PVCalcul <> 0) and (TOBRES.GEtValue('ARS_CALCULPV') = 'X') then
        TOBL.putValue('BCO_PUHT',PVCalcul)
     else
        if PrixVente <> 0 then
           TOBL.putValue('BCO_PUHT',PrixVente)
        else
           //calcul du prix de revient à partir du coef de marge de la prestation
        	 TOBL.putValue('BCO_PUHT', CoefMarge * PxRevient);
		end;
end;

procedure GetValoPrestation(TOBL : TOB);
var QQ : Tquery;
    Prixvente,PrixRevient,CoefMarge,PA : double;
    FromBordereau : Boolean;
begin

  FromBordereau := false;

	// Recherche préalable d'entente de prix avec le client (Bordereau de prix)
  if GetprixFromBordereau (getTiersAffaire(TOBL.getValue('BCO_AFFAIRE')),
  													TOBL.getValue('BCO_ARTICLE'),
                            TOBL.getValue('BCO_AFFAIRE'),
                            prixVente) then
  begin
    PrixRevient := TOBL.GetValue('BCO_DPR');
    if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;
    //Chargement de la tobconso
    TOBL.PutValue('BCO_PUHT',prixvente);
    TOBL.PutValue('COEFMARGE', CoefMarge);
  	FromBordereau := True;
  end;
	//
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.getValue('BCO_ARTICLE')+'"',true);
  if (not QQ.eof) then
  begin
    //chargement des valeur numéraire
    if not FromBordereau then PrixVente := QQ.findField('GA_PVHT').AsFloat;
    PA := QQ.findField('GA_PAHT').AsFloat;
    PrixRevient := QQ.findField('GA_DPR').AsFloat;
    if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;
    //Chargement de la tobconso
    TOBL.PutValue('BCO_PUHT',prixvente);
    TOBL.PutValue('BCO_DPA',PA);
    TOBL.PutValue('BCO_DPR',PrixRevient);
    TOBL.PutValue('COEFMARGE', CoefMarge);
  end;
  ferme (QQ);
end;



procedure GetPVPrestation (TOBL : TOB);
var QQ : Tquery;
    Prixvente,PrixRevient,CoefMarge : double;
    TOBA : TOB;
begin
  //
	TOBA := TOB.Create ('ARTICLE',nil,-1);
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.getValue('BCO_ARTICLE')+'"',true);
  if (not QQ.eof) then
  begin
		TOBA.selectDb ('',QQ);
  end;
  ferme (QQ);
  //
	// Recherche préalable d'entente de prix avec le client (Bordereau de prix)
  if GetprixFromBordereau (getTiersAffaire(TOBL.getValue('BCO_AFFAIRE')),
  													TOBL.getValue('BCO_ARTICLE'),
                            TOBL.getValue('BCO_AFFAIRE'),
                            prixVente) then
  begin
    PrixRevient := TOBL.GetValue('BCO_DPR');
    if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;
    //Chargement de la tobconso
    TOBL.PutValue('BCO_PUHT',prixvente);
    TOBL.PutValue('COEFMARGE', CoefMarge);
    exit;
  end;
	//
  if TOBL.GetValue('BCO_DPA')=0 then
  begin
    TOBL.PutValue('BCO_DPA',TOBA.GetDouble ('GA_PAHT'));
    TOBL.PutValue('BCO_DPR',TOBA.GetDouble ('GA_DPR'));
  end;
  //chargement des valeur numéraire
  PrixRevient := TOBL.GetValue('BCO_DPR');
  PrixVente 	:= TOBA.GetDouble ('GA_PVHT');
  if PrixRevient <> 0 then CoefMarge 	:=  PrixVente/PrixRevient;
  if Prixvente <> 0 then
  begin
    //Chargement de la tobconso
    TOBL.PutValue('BCO_PUHT',prixvente);
    TOBL.PutValue('COEFMARGE', CoefMarge);
  end;
end;

procedure SetInfoRessource (TOBL,TOBres : TOB; var WithMajPrix : boolean; WithRecupPrestation : boolean; var PrestationDefaut : string);
begin

  TOBL.putValue('LIBELLEMO',TOBRES.GetValue('ARS_LIBELLE'));
  if (TOBL.GetValue('BCO_NATUREMOUV') <> 'FRS') then TOBL.putValue('BCO_QUALIFQTEMOUV',TOBREs.GetValue('ARS_UNITETEMPS'));

  // On force dans le cas où MO interimaire ou matériels
  if ((TOBRes.GetValue('ARS_TYPERESSOURCE') = 'INT') and
     (TOBL.GetValue('BCO_NATUREMOUV') <> 'FRS')) or
     (TOBRes.GetValue('ARS_TYPERESSOURCE') = 'OUT') or
     (TOBRes.GetValue('ARS_TYPERESSOURCE') = 'MAT') Then
     WithMajPrix := true;

 // Correction FQ 11968 - evitera que la valorisation du document ne soit cassé par la saisie de la ressource
 if TOBL.GetValue('BCO_NATUREPIECEG')<> '' then exit;
 // --

 if (WithRecupPrestation) and (TOBRES.GetValue('ARS_ARTICLE')<> '') then
     begin
  	 TOBL.PutValue('BCO_ARTICLE',TOBRES.GetValue('ARS_ARTICLE'));
  	 TOBL.PutValue('BCO_CODEARTICLE',TRIM(Copy(TOBRES.GetValue('ARS_ARTICLE'),1,18)));
     PrestationDefaut := TOBRES.GetValue('ARS_ARTICLE');
     DefiniInfoArticle (TOBL,string(TOBRES.GetValue('ARS_ARTICLE')),not WithMajPrix);
     end;

//MODIF BRL 17/11/04
//if (WithMajPrix) and (TOBRES.GetValue('ARS_ARTICLE')<> '')then
	if (WithMajPrix) then
  begin
  	SetValoFromRessource (TOBL,TOBRes);
  end;

  if (copy(TOBL.getValue('BCO_AFFAIRE'),1,1)='A') and (getParamSocSecur('SO_BTVALOSCONSO',false)) then
  begin
  	if TOBL.GetValue('BCO_ARTICLE')<> '' then
    begin
    	GetValoPrestation (TOBL);
    end;
  end else if (copy(TOBL.getValue('BCO_AFFAIRE'),1,1)='W') and (getParamSocSecur('SO_BTVALOAPPELS',false)) then
  begin
  	if TOBL.GetValue('BCO_ARTICLE')<> '' then
    begin
    	GetPVPrestation (TOBL);
    end;
  end;
  AlimenteValeurInit (TOBL);

end;

procedure CalculeLaLigne (TOBL : TOB; coefPaPr : double = 0; CoefPrPv : double= 0);
var Qte : double;
begin

  Qte := TOBL.getValue('BCO_QUANTITE');

  if CoefpaPr <> 0 then
  begin
    TOBL.PutValue('BCO_DPR',Arrondi(TOBL.GetValue('BCO_DPA') * coefPaPr,V_PGI.okDecP))
  end;
  if CoefprPV <> 0 then
  begin
    TOBL.PutValue('BCO_PUHT',Arrondi(TOBL.GetValue('BCO_DPR') * coefPrPv,V_PGI.okDecP))
  end;

  TOBL.putValue ('BCO_MONTANTACH', Arrondi(Qte * TOBL.GetValue('BCO_DPA'),V_PGI.okDecV));
  TOBL.putValue ('BCO_MONTANTPR',  Arrondi (Qte * TOBL.GetValue('BCO_DPR'),V_PGI.okDecV));
  TOBL.putValue ('BCO_MONTANTHT', Arrondi(Qte * TOBL.GetValue('BCO_PUHT'),V_PGI.okDecV));
end;

procedure AfficheLaLigne (TOBL : TOB; Grille : THgrid; LesChamps : string; Arow : integer);
begin
  TOBL.PutLigneGrid (Grille,Arow,false,false,lesChamps);
end;


procedure CalculeAndAfficheLaLigne (Grille : THGrid ;lesChamps : string; Ligne : integer; TOBS : TOB);
var TOBL : TOB;
begin
  TOBL := TOBS.detail[ligne-1];
  calculeLaLigne (TOBL);
  AfficheLaLigne (TOBL,Grille,LesChamps,Ligne);
end;

function IsExisteAffaire (Code : string; Var Domaine, LibAffaire : String) : boolean;
var QQ : Tquery;
begin
  result := false;
  if Code = '' then exit;
  QQ := OpenSql ('SELECT AFF_LIBELLE, AFF_DOMAINE FROM AFFAIRE WHERE AFF_AFFAIRE="'+Code+'"',true);
  if not QQ.eof then
  begin
    Domaine := QQ.findfield('AFF_DOMAINE').asString;
    LibAffaire  := QQ.findfield('AFF_LIBELLE').asString;
    result := true;
  end;
  ferme (QQ);
end;

function IsExistePhaseAffaire (Code,Phase : string) : boolean;
var QQ : Tquery;
begin
  result := false;
  if (Code = '') then exit;
  if (Phase = '') then BEGIN result := true; Exit; END;
  QQ := OpenSql ('SELECT BPC_LIBELLE FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+Code+'" AND BPC_PHASETRA="'+Phase+'"',true);
  if not QQ.eof then result := true;
  ferme (QQ);
end;

function GetChantier (Code : string; ChangeStatut:Boolean=false) : string;
var StChamps 		: String;
	 X					: Integer;
    tmp				: String;
    StArgument	   : string;
    Part0			: string;
    Part1			: string;
    Part2			: String;
    Part3			: String;
    Avenant 		: string;
    Tiers 			: String;
    Affaire 		: String;
begin

  Tmp:=(Trim(ReadTokenSt(Code)));
  Part0 := tmp ;

  Tmp := (Trim(ReadTokenSt(Code)));
  Tiers := Tmp;

  Tmp := (Trim(ReadTokenSt(Code)));
  Affaire := Tmp;

  {$IFDEF BTP}
  //FV1 : 26/11/2013 - FS#784 - BAGE : en saisie consos, gérer la saisie partielle du code chantier dans la grille.
  BTPCodeAffaireDecoupe (Affaire,Part0,Part1,Part2,Part3,Avenant,taModif,false);
  {$ELSE}
  CodeAffaireDecoupe (Affaire,Part0,Part1,Part2,Part3,Avenant,taModif,false);
  {$ENDIF}

  if Part0 = 'W' then
  Begin
    Stchamps := 'AFF_STATUTAFFAIRE=APP';
    StChamps := 'AFF_ETATAFFAIRE=AFF';
    StChamps := StChamps + ';AFF_RESPONSABLE='+ Tiers;
    //
    StArgument := 'STATUT=APP';
  end
  Else if Part0 = 'I' then
  begin
    Stchamps := ';AFF_STATUTAFFAIRE=INT';
    StChamps := StChamps + 'AFF_AFFAIRE0=' + Part0;
    //
    StArgument := 'STATUT:INT';
  end
  Else if Part0 = 'A' then
  begin
    Stchamps := 'AFF_STATUTAFFAIRE=AFF;';
    //
    StArgument := 'STATUT:AFF';
  end
  else
  begin
   Stchamps := 'AFF_STATUTAFFAIRE=AFF';
   //
   StArgument := 'STATUT:AFF';
  end;

  //FV1 : 26/11/2013 - FS#784 - BAGE : en saisie consos, gérer la saisie partielle du code chantier dans la grille.
  StChamps := StChamps + ';AFF_AFFAIRE0='  + Part0;
  StChamps := StChamps + ';AFF_AFFAIRE1=' + Part1;
  StChamps := StChamps + ';AFF_AFFAIRE2=' + Part2;
  StChamps := StChamps + ';AFF_AFFAIRE3=' + Part3;
  StChamps := StChamps + ';AFF_AVENANT='  + Avenant;
  //
  stArgument := Stargument + ';TOUS=OUI';
  StArgument := Stargument +';ACTION=CONSULTATION' ;//mcd 25/09/03 pour acceptation proposision (pas créat sur mul rech misison origine)
  if not ChangeStatut then Stargument := Stargument + ';NOCHANGESTATUT';
  //Stargument := Stargument + ';NOCHANGETIERS';

  //if Code <> '' then
  //begin
	//  StChamps := StChamps + ';AFF_AFFAIRE1=' + Part1;
  //   StChamps := StChamps + ';AFF_AFFAIRE2=' + Part2;
  //   StChamps := StChamps + ';AFF_AFFAIRE3=' + Part3;
  //   if (Avenant <> '') then
  //   	  if Avenant <> '00' then
  //         StChamps := StChamps + ';AFF_AVENANT=' + Avenant;
	//end;

  Stargument := Stargument + ';ACTION=RECH';

  tmp := AGLLanceFiche('BTP', 'BTAFFAIRE_MUL', StChamps, '', StArgument);

  if tmp <> '' then
     begin
     result := ReadTokenSt(tmp);
     end;

end;

function GetLibelleMateriaux (Code : string) : string;
var QQ : TQuery;
begin
  result := '';
  if Code= '' then exit;
  QQ := OpenSql ('SELECT GA_LIBELLE FROM ARTICLE WHERE GA_ARTICLE="'+Code+'"',true);
  if not QQ.eof then result := QQ.findField('GA_LIBELLE').asstring;
  ferme (QQ);
end;

procedure AddLesSupLignesConso (TOBL : TOB);
var year,Month,Day : word;
begin
  DecodeDate (TOBL.GetValue('BCO_DATEMOUV'),year,Month,Day);

  TOBL.AddChampSupValeur ('NEW','-');
  TOBL.AddChampSupValeur ('MODIF','-');
  TOBL.AddChampSupValeur ('TYPEMODIF','-');
  TOBL.AddChampSupValeur ('SELECT','');
  TOBL.AddChampSupValeur ('JOUR', InttoStr(Day));
  TOBL.addChampsupvaleur ('PRIXRECUP','-');
  TOBL.AddChampSupValeur ('_PA_INIT',0);
  TOBL.AddChampSupValeur ('_PR_INIT',0);
  TOBL.AddChampSupValeur ('_PV_INIT',0);
  TOBL.AddChampSupValeur ('LIBELLETYPE','');
  TOBL.AddChampSupValeur ('LIBELLEMO','');
  TOBL.AddChampSupValeur ('LIBELLENATURE','');
  TOBL.AddChampSupValeur ('LIBELLEARTICLE','');
  TOBL.AddChampSupValeur ('LIBELLECHANTIER','');
  TOBL.AddChampSupValeur ('LIBELLEPHASE','');
  TOBL.AddChampSupValeur ('COEFMARGE',0);
  TOBL.AddChampSupValeur ('COEFFG',0);
  TOBL.AddChampSupValeur ('DOMAINE', '');
  TOBL.AddChampSupValeur ('PRIXPASMODIF', '-');
  TOBL.AddChampSupValeur ('ETATAFFAIRE', '');
  TOBL.AddChampSupValeur ('TIERS', '');
  TOBL.AddChampSupValeur ('LIBTIERS', '');
  TOBL.AddChampSupValeur ('CHANGEETAT', '-');

end;

function GetTypeRessource (CodeRessource : string) : string;
var QQ : Tquery;
    SQl : String;
begin
  result := '';
  Sql := 'SELECT ARS_TYPERESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE="'+CodeRessource+'"';
  QQ := OpenSql (SQL,true);
  if not QQ.eof then result := Trim(QQ.findField('ARS_TYPERESSOURCE').AsString);
  ferme (QQ);
end;

function GetNumUniqueConso (var TheResult : double) : TGncERROR;

  function GetNumParamSoc (var Numero : string) : TGncError;
  var QQ : TQuery;
  begin
    result := GncOk;
    QQ := OPENSQL ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_BTCPTLIGCONSO"',True);
    TRY
      if QQ.eof Then
      begin
        result := GncAbort; // le compteur n'existe pas abandon
        Exit;
      end;
      Numero := QQ.findField('SOC_DATA').AsString ;
    FINALLY
      ferme(QQ);
    END;
  end;

var NumUnique : double;
    PrecNum : string;
    CptNombre : integer;
    OkNumero : boolean;
begin
  OkNumero := false;
  CptNombre := 1;
  Repeat
    result := GetNumParamSoc (precNum);
    if result = GncAbort then break;
    NumUnique := Valeur (PrecNum)+1;
    if ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="'+FloatToStr (NumUnique)+'" WHERE SOC_NOM="SO_BTCPTLIGCONSO" AND SOC_DATA="'+PrecNum+'"') > 0 then
    begin
      OkNumero := true;
      break;
    end;
    if CptNombre > 1000 then
    begin
      result := GncAbort;               
      break;
    end;
    inc(CptNombre);
  until OKNumero;

  if Result = GncOk then
  begin
    TheResult := NumUnique;
  end;

end;

function LigneFromPieces (TOBL : TOB) : boolean;
begin
  // Modif BRL 23/03/07 : on ne bloque pas l'accès pour les matériels
  // en cas de consommations issues de livraisons sur chantiers
	if TOBL.getValue('BCO_NATUREMOUV') = 'RES' then
  	result:=False
  else
  	result := (Trim(TOBL.getValue('BCO_NATUREPIECEG'))<> '');
end;

function LigneDejaValide (TOBL : TOB) : boolean;
begin
  result := (TOBL.getValue('BCO_VALIDE')= 'X');
end;

procedure GetCoefFromArt (TOBLIgne : TOB;var CoefPaPr,CoefPrPv : double);
var Select : string;
    Q : Tquery;
    prefixe,Domaine : string;
    pasTouche : boolean;
begin
	if ToBLigne.NomTable = 'LIGNE' then prefixe := 'GL'
  else if TOBLIGNE.nomtable = 'CONSOMMATIONS' then prefixe := 'BCO';
  Domaine := TOBLIGNE.GetValue(prefixe+'_DOMAINE');
  Select := 'SELECT GA_PRIXPASMODIF,GA_COEFFG,GA_COEFCALCHT FROM ARTICLE WHERE GA_ARTICLE="'+TOBLigne.GetValue(prefixe+'_ARTICLE')+'"';
  Q := OpenSql (Select,true);
  if not Q.eof then
  begin
    if Q.findfield('GA_COEFFG').AsFloat <> 0 then CoefPaPR := Q.findfield('GA_COEFFG').AsFloat;
    if Q.findfield('GA_COEFCALCHT').AsFloat <> 0 then CoefPRPV := Q.findfield('GA_COEFCALCHT').AsFloat;
    if (Q.findField('GA_PRIXPASMODIF').Asstring<>'X') then PasTouche := false else PasTouche := true;
  end;
  ferme (Q);
  if not PasTouche Then
  begin
    GetCoefDomaine (Domaine,CoefPaPR,CoefPRPV);
    if VH_GC.BTCODESPECIF = '001' then
    begin
      CalculeDonneelignePOC (TOBLigne,CoefPaPR,CoefPRPV);
  end;
  end;
end;

procedure RecupCoefs (TOBLigne : TOB; VenteAchat : string; var CoefpaPr : double; var CoefPrPv : double);
var CD : r_cledoc;
		Select : string;
    Q : Tquery;
begin
  if TOBLigne.getValue('GL_PIECEORIGINE') <> '' then
  begin
    // on récupere les coefs de la ligne de document d'origine
    DecodeRefPiece (TOBLigne.getValue('GL_PIECEORIGINE'),CD);
  	if (VenteAchat = 'ACH') and (GetInfoParPiece (CD.Naturepiece,'GPP_VENTEACHAT')='ACH') then
    begin
    	GetCoefFromArt (TOBLIgne,CoefPaPr,CoefPrPv);;
    end else
    begin
    	Select := 'SELECT GL_DPR,GL_DPA,GL_PUHT FROM LIGNE WHERE '+WherePiece(CD,TtdLigne,true,true);
    	Q := OpenSql (Select,true);
   	  if not Q.eof then
   	  begin
      	if Q.findfield('GL_DPA').AsFloat > 0 then CoefPaPR := Q.findfield('GL_DPR').AsFloat/Q.findfield('GL_DPA').AsFloat;
      	if Q.findfield('GL_DPR').AsFloat > 0 then CoefPRPV := Q.findfield('GL_PUHT').AsFloat/Q.findfield('GL_DPR').AsFloat;
      end;
      ferme (Q);
    end;
  end else
  if (TOBLIgne.getValue('GL_PIECEPRECEDENTE') <> '') then // and (venteAchat = 'VEN') then
  begin
    DecodeRefPiece (TOBLigne.getValue('GL_PIECEPRECEDENTE'),CD);
  	if (VenteAchat = 'ACH') and (GetInfoParPiece (CD.Naturepiece,'GPP_VENTEACHAT')='ACH') then
    begin
    	GetCoefFromArt (TOBLIgne,CoefPaPr,CoefPrPv);
    end else
    begin
      Select := 'SELECT GL_DPR,GL_DPA,GL_PUHT FROM LIGNE WHERE '+WherePiece(CD,TtdLigne,true,true);
      Q := OpenSql (Select,true);
      if not Q.eof then
      begin
        if Q.findfield('GL_DPA').AsFloat > 0 then CoefPaPR := Q.findfield('GL_DPR').AsFloat/Q.findfield('GL_DPA').AsFloat;
        if Q.findfield('GL_DPR').AsFloat > 0 then CoefPRPV := Q.findfield('GL_PUHT').AsFloat/Q.findfield('GL_DPR').AsFloat;
      end;
      ferme (Q);
    end;
  end else
  begin
  	GetCoefFromArt (TOBLIgne,CoefPaPr,CoefPrPv);;
  end;
end;

function GetNatureMouv (TypeArticle,TypeRessource : string) : string;
begin
    if TypeArticle = 'PRE' then
    begin
      if TypeRessource = 'SAL' then
        result :=  'MO'
      else if  TypeRessource = 'ST' then
        result := 'EXT'
      else if  TypeRessource = 'INT' then
        Result := 'EXT'
      else if  TypeRessource = 'AUT' then
        result := 'EXT'
      else if  TypeRessource = 'LOC' then
        result := 'EXT'
      else if  TypeRessource = 'MAT' then
        result := 'RES'
      else if  TypeRessource = 'OUT' then
        result := 'RES'
    end else if (TypeArticle = 'MAR') or (TypeArticle = 'ARP') then
    Begin
      result := 'FOU'
    end else if TypeArticle = 'FRA' then
    Begin
      result :='FRS';
    end;
end;

function GetprixFromBordereau (CodeTiers,CodeArticle,CodeAffaire : string; var Prixvente : double) : boolean;
var Sql, TiersFacture : String;
		QQ : TQuery;
    fDateNow : TdateTime;
    year,Month,Day : word;
begin
  DecodeDate(Now,year,Month,Day);
  fDateNow := EncodeDate(Year,Month,day);
	result := false;
  prixvente := 0;

  // En contrat-interventions, recherche préalable d'un bordereau pour le client facturé s'il y en a un
  if (Copy(CodeAffaire,1,1)='I') or (Copy(CodeAffaire,1,1)='W') then
  begin
    Sql := 'SELECT T2.T_TIERS FROM TIERS AS T2, TIERS AS T1 WHERE T1.T_TIERS="'+CodeTiers+'"'+' AND T2.T_AUXILIAIRE=T1.T_FACTURE';
    QQ := OpenSql(Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      TiersFacture := QQ.Fields[0].AsString;
      Ferme(QQ);
  	  Sql := 'SELECT BDE_AFFAIRE,BDE_DATEFIN,BDE_DATEDEPART,GL_PUHT FROM LIGNE '+
  		  	   'LEFT JOIN BDETETUDE ON BDE_NATUREPIECEG=GL_NATUREPIECEG AND '+
  			     'BDE_SOUCHE=GL_SOUCHE AND BDE_NUMERO=GL_NUMERO AND BDE_INDICEG=GL_INDICEG '+
             'WHERE '+
             'BDE_NATUREAUXI="CLI" AND BDE_CLIENT="'+TiersFacture+'" AND BDE_AFFAIRE="" AND ';
      Sql := Sql + 'GL_NATUREPIECEG="BBO" AND '+
       			 'BDE_DATEFIN > "'+USDATETIME(fdateNow)+'" AND '+
             'GL_ARTICLE="'+CodeArticle+'" ORDER BY BDE_AFFAIRE DESC';
      QQ := OpenSql(Sql,true,-1,'',true);
      if not QQ.eof then
      begin
  	    result := true;
        QQ.first;
        prixvente := QQ.FindField('GL_PUHT').AsFloat;
      end;
      ferme (QQ);
    end else
      ferme(QQ);
  end;

  if result = True then Exit; // On a trouvé un bordereau pour le client facturé donc on sort.

	Sql := 'SELECT BDE_AFFAIRE,BDE_DATEFIN,BDE_DATEDEPART,GL_PUHT FROM LIGNE '+
  			 'LEFT JOIN BDETETUDE ON BDE_NATUREPIECEG=GL_NATUREPIECEG AND '+
  			 'BDE_SOUCHE=GL_SOUCHE AND BDE_NUMERO=GL_NUMERO AND BDE_INDICEG=GL_INDICEG '+
         'WHERE '+
				 'BDE_NATUREAUXI="CLI" AND BDE_CLIENT="'+CodeTiers+'" AND ';
  if Copy(CodeAffaire,1,1)='A' then
  begin
  	Sql := Sql + '(BDE_AFFAIRE="'+CodeAffaire+'" OR BDE_AFFAIRE="") AND ';
  end else
  begin
  	Sql := Sql +'BDE_AFFAIRE="" AND ';
  end;
  Sql := Sql + 'GL_NATUREPIECEG="BBO" AND '+
         			 'BDE_DATEFIN > "'+USDATETIME(fdateNow)+'" AND '+
               'GL_ARTICLE="'+CodeArticle+'" ORDER BY BDE_AFFAIRE DESC';
  QQ := OpenSql(Sql,true,-1,'',true);
  if not QQ.eof then
  begin
  	result := true;
    QQ.first;
    prixvente := QQ.FindField('GL_PUHT').AsFloat;
  end;
  ferme (QQ);
end;

function getTiersAffaire(CodeAffaire : string) : string;
var Sql : string;
		QQ : Tquery;
begin
	result := '';
	Sql := 'SELECT AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAffaire+'"';
  QQ := OpenSql (Sql,true,1,'',True);
  if not QQ.eof then
  begin
  	Result := QQ.findField('AFF_TIERS').AsString;
  end;
  Ferme (QQ);
end;

//FV1 : 28/08/2013- FS#632 - POUCHAIN : Ajout de coef par domaine d'activité pour les interventions
Procedure AppliqueCoeff(TOBL : TOB);
var CoefFG     : Double;
    CoefMarge  : Double;
    DPA         : Double;
    DPR         : Double;
    Domaine     : String;
  	ValoFromPresta : Boolean;
begin
  Domaine := TOBL.GetString('DOMAINE');
  ValoFromPresta :=  GetParamSocSecur('SO_BTVALOAPPELS',True);

  if Domaine = '' then exit;

  if (ValoFromPresta) and (TOBL.GetString('PRIXPASMODIF')='X') then exit;
  if (not ValoFromPresta) and (TOBL.GetString('PRIXPASMODIF')='X') AND
  	 (Pos(TOBL.GetString('BCO_NATUREMOUV'),'MO;RES')<=0)  then EXIT;

  CoefFG := 0;
  CoefMarge := 0;
  GetCoefDomaine(Domaine,CoefFG,CoefMarge, 'W');
  if (CoefFG = 0) and (CoefMarge = 0) then Exit;

  DPA := TOBL.GetDouble('BCO_DPA');

  TOBL.SetDouble('COEFFG', CoefFG);
  TOBL.PutValue('BCO_DPR',Arrondi(DPA*CoefFG,V_PGI.okdecP));

  DPR := TOBL.GetDouble('BCO_DPR');
  TOBL.SetDouble('COEFMARGE', CoefMarge);
  TOBL.PutValue('BCO_PUHT',Arrondi(DPR*CoefMarge,V_PGI.okdecP));

end;

function GetFamilleTaxe(Article : string) : string;
var QQ : TQuery;
begin
  Result := 'TN';
  if Article ='' then Exit;
  QQ := OpenSQL('SELECT GA_FAMILLETAXE1 FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',True,1,'',true);
  if not QQ.eof then
  begin
    Result := QQ.fields[0].AsString;
  end;
  ferme (QQ);
end;


function ConstitueDateDebutTravaux (TOBL : TOB) : TDateTime;
var year,Month,day,hh,mm,ss,ms,yy1,mm1,dd1 : Word;
begin
  DecodeDate (TOBL.GetValue('BCO_DATEMOUV'),year,month,day);
  DecodeDateTime(TOBL.GetValue('BCO_DATETRAVAUX'),yy1,mm1,dd1,hh,mm,ss,ms);
  if hh = 0 then DecodeDateTime(GetDebutMatinee,yy1,mm1,dd1,hh,mm,ss,ms);
  result := EncodeDateTime( year,Month,day,hh,mm,ss,0);
end;


end.
