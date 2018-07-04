unit UtilSuprAuxi;

interface

uses
{$IFDEF EAGLCLIENT}
    UTob,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    Hctrls,ParamSoc,Hent1, Ent1;

Function SupAEstDansPiece ( St : String ) : Boolean ;
Function SupAEstDansActivite ( St : String ) : Boolean ;
Function SupAEstDansRessource ( St : String ) : Boolean ;
Function SupAEstDansActions ( St : String ) : Boolean ;
Function SupAEstDansPersp ( St : String ) : Boolean ;
Function SupAEstDansAffaire ( St : String ) : Boolean ;
Function SupAEstDansCata ( St : String ) : Boolean ;
Function SupAEstDansPaie ( St : String ) : Boolean ;
Function SupAEstDansMvtPaie ( St : String ) : Boolean ;
Function SupAEstEcrGuide(St : String) : Boolean ;
Function SupAEstDansSociete(St : String) : Boolean ;
Function SupAEstDansSection(St : String) : Boolean ;
Function SupAEstDansUtilisat(St : String) : Boolean ;
Function SupAEstCpteCorresp(St : String) : Boolean ;
Function SupAEstUnPayeur ( St : String ) : Boolean ;
Function SupAEstMouvemente(St : String) : Boolean  ;
Function SupAEstDansProjets ( St : String ) : Boolean ;
Function SupAEstDansGED ( St : String ) : Boolean ;

implementation

Function SupAEstDansPiece ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT GP_TIERS FROM PIECE WHERE GP_TIERS="'+St+'" OR GP_TIERSLIVRE="'+St+'" OR GP_TIERSFACTURE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansActivite ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT ACT_TIERS FROM ACTIVITE WHERE ACT_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansRessource ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT ARS_AUXILIAIRE FROM RESSOURCE WHERE ARS_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansActions ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT RAC_AUXILIAIRE FROM ACTIONS WHERE RAC_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansPersp ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT RPE_AUXILIAIRE FROM PERSPECTIVES WHERE RPE_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansAffaire ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT AFF_TIERS FROM AFFAIRE WHERE AFF_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansCata ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT GCA_TIERS FROM CATALOGU WHERE GCA_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansPaie ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT PSA_AUXILIAIRE FROM SALARIES WHERE PSA_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansMvtPaie ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT PPU_SALARIE FROM PAIEENCOURS LEFT JOIN SALARIES ON PPU_SALARIE=PSA_SALARIE WHERE PSA_AUXILIAIRE="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstEcrGuide(St : String) : Boolean ;
BEGIN Result:=Presence('ECRGUI','EG_AUXILIAIRE',St) ; END ;

Function SupAEstDansSociete(St : String) : Boolean ;
Var StLoc : String ;
{$IFDEF SPEC302}
    Q : TQuery ;
{$ENDIF}
BEGIN
{$IFDEF SPEC302}
Q:=OpenSql('Select SO_CLIATTEND,SO_FOUATTEND From SOCIETE Where SO_CLIATTEND="'+St+'" OR SO_FOUATTEND="'+St+'"',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
{$ELSE}
Result:=True ;
StLoc:=GetParamSoc('SO_CLIATTEND') ; if StLoc=St then Exit ;
StLoc:=GetParamSoc('SO_FOUATTEND') ; if StLoc=St then Exit ;
if St=VH^.TiersDefSal then Exit ;
if St=VH^.TiersDefDiv then Exit ;
if St=VH^.TiersDefCli then Exit ;
if St=VH^.TiersDefFou then Exit ;
Result:=False ;
{$ENDIF}
END ;

Function SupAEstDansSection(St : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select S_MAITREOEUVRE,S_CHANTIER From SECTION Where S_MAITREOEUVRE="'+St+'" OR S_CHANTIER="'+St+'"',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function SupAEstDansUtilisat(St : String) : Boolean ;
BEGIN Result:=Presence('UTILISAT','US_AUXILIAIRE',St) ; END ;

Function  SupAEstCpteCorresp(St : String) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('Select CR_CORRESP From CORRESP Where (CR_TYPE="AU1" OR CR_TYPE="AU2" )AND CR_CORRESP="'+St+'" ',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function  SupAEstUnPayeur ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSql('SELECT T_PAYEUR FROM TIERS WHERE T_PAYEUR="'+St+'"',True) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function SupAEstMouvemente(St : String) : Boolean  ;
BEGIN
Result:=ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+St+'" '+
                  'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="'+St+'" )') ;
END ;

Function SupAEstDansProjets ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT RPJ_TIERS FROM PROJETS WHERE RPJ_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;

Function SupAEstDansGED ( St : String ) : Boolean ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT RTD_TIERS FROM RTDOCUMENT WHERE RTD_TIERS="'+St+'"',True) ;
Result:=Not Q.EOF ;
Ferme(Q) ;
END ;
end.
 