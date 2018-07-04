unit TZ;

interface

uses
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     HCtrls, HEnt1,                     { +- Bibli  -+ }
     UTob,
     SaisUtil ;                             { +- Mac4   -+ }

type
  TZF = class(TOB)
    private
    procedure   ChampSupEcr ;
    procedure   ChampSupBal ;
    procedure   ChampSupRbp ;
    procedure   DefautDBVal ;
    procedure   DefautEcr ;
    procedure   DefautAna ;
    procedure   DefautBal ;
    procedure   DefautRbp ;
    procedure   DefautSupVal ;
    procedure   DefautSupEcr ;
    procedure   DefautSupAna ;
  public
    constructor Create(LeNomTable : string ; LeParent : TOB ; IndiceFils : Integer) ; override ;
    constructor CreateDB(LeNomTable : string ; LeParent : TOB ; IndiceFils : Integer ; Q : TQuery) ;
    procedure   HistoMontants;
    function    GetCotation(DateEcr : TDateTime; Taux : Double; Dev : String3) : Double ;
    procedure   SetCotation ;
    procedure   SetMPACC ;
    procedure   SetMontants(XD, XC : Double ; Dev : RDEVISE ; Force : Boolean) ;
    procedure   ChampSupAna ;
  end ;

implementation

uses
  SysUtils,
  ULibEcriture,
  Ent1 
   ;       { +- Delphi -+ }

constructor TZF.Create(LeNomTable : String ; LeParent : TOB ; IndiceFils : Integer) ;
begin
inherited Create(LeNomTable, LeParent, IndiceFils) ;
if NomTable='ECRITURE' then ChampSupEcr ;
if NomTable='ANALYTIQ' then ChampSupAna ;
if NomTable='HISTOBAL' then ChampSupBal ;
if NomTable='RBLIGECR' then ChampSupRbp ;
DefautDBVal ;
DefautSupVal ;
end ;

constructor TZF.CreateDB(LeNomTable : String ; LeParent : TOB ; IndiceFils : Integer ; Q : TQuery) ;
begin
inherited CreateDB(LeNomTable, LeParent, IndiceFils, Q) ;
if NomTable='ECRITURE' then ChampSupEcr ;
if NomTable='ANALYTIQ' then ChampSupAna ;
if NomTable='HISTOBAL' then ChampSupBal ;
if NomTable='RBLIGECR' then ChampSupRbp ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 18/05/2004
Modifié le ... : 11/09/2006
Description .. : - LG - 18/05/2004 - modif pour acceleration en saisie ajout 
Suite ........ : de acc et cptcegid
Suite ........ : -FB 18771 - LG - 11/09/2006 - ajout du champs 
Suite ........ : NEWVENTIL. identifie 
Suite ........ : une nouvelle ligne d'analytique
Mots clefs ... : 
*****************************************************************}
procedure TZF.ChampSupEcr ;
begin
AddChampSup('OLDDEBIT',        FALSE) ; AddChampSup('OLDCREDIT',         FALSE) ;
AddChampSup('OLDGENERAL',      FALSE) ; AddChampSup('OLDAUXILIAIRE',     FALSE) ;
AddChampSup('GUIDE',           FALSE) ; AddChampSup('BADROW',            FALSE) ;
//AddChampSup('INITGUIDE',       FALSE) ;
//AddChampSup('SPEEDROW',        FALSE) ; AddChampSup('INSROW',            FALSE) ;
AddChampSup('EG_NUMLIGNE',     FALSE) ;
AddChampSup('EG_DEVISE',       FALSE) ;
AddChampSup('EG_GENERAL',      FALSE) ; AddChampSup('EG_AUXILIAIRE',     FALSE) ;
AddChampSup('EG_REFINTERNE',   FALSE) ; AddChampSup('EG_LIBELLE',        FALSE) ;
AddChampSup('EG_DEBITDEV',     FALSE) ; AddChampSup('EG_CREDITDEV',      FALSE) ;
AddChampSup('EG_MODEPAIE',     FALSE) ; AddChampSup('EG_DATEECHEANCE',   FALSE) ;
AddChampSup('EG_REFEXTERNE',   FALSE) ; AddChampSup('EG_DATEREFEXTERNE', FALSE) ;
AddChampSup('EG_REFLIBRE',     FALSE) ; AddChampSup('EG_AFFAIRE',        FALSE) ;
AddChampSup('EG_QTE1',         FALSE) ; AddChampSup('EG_QTE2',           FALSE) ;
AddChampSup('EG_QUALIFQTE1',   FALSE) ; AddChampSup('EG_QUALIFQTE2',     FALSE) ;
AddChampSup('AG_AXE',          FALSE) ; AddChampSup('AG_SECTION',        FALSE) ;
AddChampSup('AG_POURCENTAGE',  FALSE) ;
AddChampSup('AG_POURCENTQTE1', FALSE) ; AddChampSup('AG_POURCENTQTE2',   FALSE) ;
// (PFU : ?) A UTILISER
AddChampSup('EG_ARRET',        FALSE) ; AddChampSup('AG_ARRET',          FALSE) ;
AddChampSup('COMPS',           FALSE) ; // saisie complemetaire ? equivalent Au champs CompS du TOBM
AddChampSup('ACC',             FALSE) ; // true si une saisie accelera a été faite
AddChampSupValeur('CPTCEGID'   , '')  ; AddChampSupValeur('YTCTIERS', '') ;
AddChampSupValeur('COMPTAG'    , '-') ; AddChampSup('PCOMPL',false)   ;
AddChampSupValeur('LETTAG', '-' )     ; AddChampSup('NEWVENTIL',false)   ;
end ;

procedure TZF.ChampSupAna ;
begin
AddChampSup('OLDDEBIT',        FALSE) ; AddChampSup('OLDCREDIT',         FALSE) ;
AddChampSup('BADROW',          FALSE) ;
end ;

procedure TZF.ChampSupBal ;
begin
AddChampSup('BADROW',          FALSE) ; AddChampSup('TYPEGENE',          FALSE) ;
AddChampSup('LIBELLE',         FALSE) ; AddChampSup('LETTRABLE',         FALSE) ;
AddChampSup('MODEREGL',        FALSE) ;
end ;

procedure TZF.ChampSupRbp ;
begin
AddChampSup('BADROW',          FALSE) ; AddChampSup('GUIDE',             FALSE) ;
end ;

procedure TZF.DefautDBVal ;
begin
if NomTable='ECRITURE' then DefautEcr ;
if NomTable='ANALYTIQ' then DefautAna ;
if NomTable='HISTOBAL' then DefautBal ;
if NomTable='RBLIGECR' then DefautRbp ;
end ;

procedure TZF.DefautEcr ;
begin
CPutDefautEcr(self) ;
PutValue('COMPS','-') ;
end ;

procedure TZF.DefautAna ;
begin
PutValue('Y_EXERCICE',      VH^.Entree.Code) ;
PutValue('Y_NATUREPIECE',   'OD') ;
PutValue('Y_QUALIFPIECE',   'N') ;
PutValue('Y_ETAT',          '0000000000') ;
PutValue('Y_UTILISATEUR',   V_PGI.User) ;
PutValue('Y_DATECREATION',  Date) ;
PutValue('Y_DATEMODIF',     NowH) ;
PutValue('Y_SOCIETE',       V_PGI.CodeSociete) ;
PutValue('Y_ETABLISSEMENT', VH^.EtablisDefaut) ;
PutValue('Y_DEVISE',        V_PGI.DevisePivot) ;
PutValue('Y_TAUXDEV',       1) ;
PutValue('Y_DATETAUXDEV',   V_PGI.DateEntree) ;
PutValue('Y_QUALIFQTE1',    '...') ;
PutValue('Y_QUALIFQTE2',    '...') ;
PutValue('Y_QUALIFECRQTE1', '...') ;
PutValue('Y_QUALIFECRQTE2', '...') ;
PutValue('Y_ECRANOUVEAU',   'N') ;
PutValue('Y_CREERPAR',      'SAI') ;
PutValue('Y_EXPORTE',       '---') ;
PutValue('Y_CONFIDENTIEL',  '0') ;
end ;

procedure TZF.DefautBal ;
begin
PutValue('HB_SOCIETE',       V_PGI.CodeSociete) ;
PutValue('HB_ETABLISSEMENT', VH^.EtablisDefaut) ;
end ;

procedure TZF.DefautRbp ;
begin
PutValue('CRB_SOCIETE',      V_PGI.CodeSociete) ;
end ;

procedure TZF.DefautSupVal ;
begin
if NomTable='ECRITURE' then DefautSupEcr ;
if NomTable='ANALYTIQ' then DefautSupAna ;
end ;

procedure TZF.DefautSupEcr ;
begin
PutValue('OLDDEBIT', 0) ;   PutValue('OLDCREDIT', 0) ;
//PutValue('SPEEDROW', '-') ;
//PutValue('INSROW',   '-') ;
end ;

procedure TZF.DefautSupAna ;
begin
PutValue('OLDDEBIT', 0) ;   PutValue('OLDCREDIT', 0) ;
end ;


procedure TZF.HistoMontants;
var P : string ;
begin
P:=TableToPrefixe(NomTable) ;
PutValue('OLDDEBIT',  GetValue(P+'_DEBIT')) ;
PutValue('OLDCREDIT', GetValue(P+'_CREDIT'))  ;
end ;

function TZF.GetCotation(DateEcr : TDateTime; Taux : Double; Dev : String3) : Double ;
var Cote : Double ;
begin
if DateEcr<V_PGI.DateDebutEuro then Cote:=Taux else
   begin
   if Dev=V_PGI.DevisePivot then Cote:=1.0 else
     if V_PGI.TauxEuro<>0 then Cote:=V_PGI.TauxEuro/Taux else Cote:=1 ;
   end ;
Result:=Cote ;
end ;

procedure TZF.SetCotation ;
begin
if NomTable<>'ECRITURE' then Exit ;
CSetCotation(Self) ;
end ;

procedure TZF.SetMPACC ;
begin
if NomTable<>'ECRITURE' then Exit ;
PutValue('E_CODEACCEPT', MPTOACC(GetValue('E_MODEPAIE'))) ;
end ;

procedure TZF.SetMontants(XD, XC : double ; Dev : RDEVISE ; Force : boolean) ;
begin
CSetMontants(self,XD,XC,DEv,Force) ;
end ;

end.
