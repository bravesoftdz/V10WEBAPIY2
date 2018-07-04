unit FactUtilBTP;

interface

uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  {$IFDEF NOMADE}
  {$IFDEF EAGLCLIENT}
  UToxClasses,
  {$ENDIF}
  {$ENDIF}
  SysUtils, Dialogs, Utiltarif, TarifUtil, UtilPGI, UtilGC, AGLInit, EntGC, SaisUtil, Windows,
  Forms, Classes, AGLInitGC, UtilArticle, UtilDimArticle, HDimension, HMsgBox,
  ParamSoc, HPanel,UentCommun;

procedure AffectePrixValoBTP (TOBL,TOBOuvrage : TOB);

implementation

uses StockUtil,
	 FactComm,
     FactSpec,
     Facture,
     TiersUtil,
     FactCpta,
     NomenUtil,
     FactNomen,
     DepotUtil //MODIF AC

  {$IFDEF AFFAIRE}
  , AffaireUtil
  {$ENDIF}

  {$IFDEF BTP}
  , CalcOleGenericBTP
  {$ENDIF}

  , FactAcompte
  , FactOuvrage
  , UtilSuggestion
  , FactUtil
  {$IFDEF GPAO}
  , wTarifs
  {$ENDIF}
  ;

procedure AffectePrixValoBTP (TOBL,TOBOuvrage : TOB);
var DEv : Rdevise;
    TOBOUV : TOB;
    IndiceNomen: integer;
    valeurs: T_Valeurs;

begin
  DEV.Code := TOBL.GetValue('GL_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBL.GetValue('GL_TAUXDEV');
  // recupération des prix d'achat en procenance des details d'ouvrages
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;
  TOBOuv := TOBOuvrage.Detail[IndiceNomen - 1];
  AffecteValoFromOuv(TOBOUV, DEV, valeurs);
  TOBL.Putvalue('GL_DPA', valeurs[0]);
  TOBL.Putvalue('GL_DPR', valeurs[1]);
  TOBL.Putvalue('GL_PMAP', valeurs[6]);
  TOBL.Putvalue('GL_PMRP', valeurs[7]);
end;

end.
