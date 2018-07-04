unit FactComm;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls, CalcOleGescom,
  {$IFDEF BTP}
  factCommBtp,
  forms,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  HPdfPrev, UtileAGL, Maineagl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB, EdtEtat, EdtDoc,
  EdtREtat, EdtRDoc,
  {$ENDIF}
  SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit, FactUtil, FactTOB, FactArticle,
  FactLotSerie,
  Math, StockUtil, EntGC, Classes, HMsgBox, FactNomen, TiersUtil, FactTiers, FactPiece,
  {$IFDEF CHR}HRReglement, {$ENDIF}
  {$IFDEF GRC}UtilRT,{$ENDIF}
  uRecupSQLModele,
	BTFactImprTOB,
  Echeance, UtilGC, UtilArticle, ParamSoc,uEntCommun,UtilTOBPiece;

Type T_UpDateReliquat = (urCreat, urModif, urSuppr); { NEWPIECE }

  // Modif BTP
function RecupDebutPourcent(TOBpiece: TOB; Indice, niveau: Integer): Integer;
function RecupDebutParagraph(TOBpiece: TOB; Indice, niveau: Integer; ParagInclus: boolean = false): Integer;
function RecupFinParagraph(TOBpiece: TOB; Indice, niveau: Integer; ParagInclus: boolean = false): Integer;
// Modif BTP
procedure InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire: TOB; ARow: integer; AvecCalcul: boolean = false);
// Modif BTP
//procedure CreerTOBLignes(GS: THGrid; TOBPiece, TOBTiers, TOBAffaire: TOB; ARow: integer; AvecCalcul: boolean = false);
procedure CreerTOBLignes(GS: THGrid; TOBPiece, TOBTiers, TOBAffaire: TOB; ARow: integer; AvecCalcul: boolean = false; TheDestinationLiv : TObject =nil );
// ---
function GetTOBPrec(TOBL, TOBPiece_O: TOB; AvecCommentDim: Boolean = False): TOB;
// Modif BTP
function DetruitAncien(TOBPiece_O, TOBBases_O, TOBEches_O, TOBN_O, TOBLOT_O, TOBAcomptes_O, TOBPorcs_O, TOBSerie_O, TOBOuvrage_O, TOBLienOle_O, TOBPieceRG_O, TOBBasesRg_O, TobLigneTarif_O: TOB; SuppTOX: Boolean = False; OkSuppLigneFac : boolean=true): boolean; overload;
function DetruitAncien(Cledoc : R_cledoc; DateModif :TdateTime ): boolean; overload;
// ------
{ NEWPIECE }
function UpdateAncien(TOBPiece_O, TOBPiece, TOBAcomptes_O, TobTiers: TOB; ViaSQL: Boolean = False; NewNum: Integer = 0): boolean;
procedure ValideLesArticles(TOBPiece, TOBArticles: TOB);
procedure ValideLesCatalogues(TOBPiece, TOBCatalogu: TOB);
procedure ValideLeTiers(TOBPiece, TOBTiers: TOB);
procedure InvalideModifTiersPiece(TOBTiers: TOB);
function DepotGererSurSite(Depot: string): boolean;
//procedure ValideLesLignes(TOBPiece, TOBArticles, TOBCatalogu, TOBNomenclature, TOBOuvrage, TOBPieceRG, TOBBasesRG: TOB; Reliquat, Rafale: boolean);
procedure ValideLesLignes(TOBPiece, TOBArticles, TOBCatalogu, TOBNomenclature, TOBOuvrages, TOBPieceRG, TOBBasesRG: TOB; Reliquat, Rafale: boolean; Ajout : boolean = false;GenereAuto:boolean=false;TOBOldPiece : TOB=nil; AutoriseDecoup : boolean=false);

procedure ValideLesPorcs(TOBPiece, TOBPorcs: TOB);
procedure ValideLaCotation(TOBPiece, TOBBases, TOBEches: TOB);
procedure ValideLaPeriode(TOBPiece: TOB);

procedure InverseStock (TOBPiece_O, TOBArticles, TOBCatalogu, TOBOuvrage_O : TOB; ControleStock : boolean=true) ; { NEWPIECE }
procedure InverseStockTransfo(TOBPiece_O, TobPiece, TOBArticles, TOBCatalogu, TOBOuvrage_O : TOB; Litige: Boolean = False); { NEWPIECE }

function CreerTOBPieceVide(var CleDoc: R_CleDoc; CodeTiers, CodeAffaire, Etablissement, Domaine: string; EnHT, SaisieContre: Boolean; NumPiece: integer = -1):TOB;
function CreerPieceVide(var CleDoc: R_CleDoc; CodeTiers, CodeAffaire, Etablissement, Domaine: string; EnHT, SaisieContre: Boolean): boolean;
// Impression document
Procedure ModificationModele(CleDoc:R_CleDoc;ImpDoc,OkEtat,Bapercu,OkEtiq,BApercuEtiq: Boolean; Modele,ModeleEtiq:String;NbCopies:Integer;TOBPiece:TOB;TL:TList;TT:TStrings;DGD : boolean=false; ImpressionViaTOB : TImprPieceViaTOB = nil) ;

(* OPTIMIZATION *)
procedure ImpressionPieceEtiq(CleDoc: R_CleDoc; ImpDoc, OkEtat, Bapercu, OkEtiq, BApercuEtiq: Boolean; Modele, ModeleEtiq: string; NbCopies: Integer; TOBPiece:
  TOB; TL: TList; TT: TStrings; DGD: boolean = false; ImpressionViaTOB : TImprPieceViaTOB = nil);
procedure ImprimerLaPiece(TOBPiece, TOBTiers: TOB; CleDoc: R_CleDoc; BImpClick: Boolean = False; DGD : boolean = false; TOBAffaire : TOB=nil;ImpressionViaTOB : TImprPieceViaTOB = nil);
(* ------------ *)
//
procedure DeflagEdit(TOBPiece: TOB);
procedure EtudieColsGrid(GS: THgrid);
function ChoixCommercial(TypeCom, ZoneCom: string): string;
function AdapteCond(TOBL, TOBConds: TOB; var Qte: Double): boolean;
procedure MontreNumero(TOBPiece: TOB);
//function EncodeRefPiece(TOBPL: TOB; NewNum: Integer = 0): string; { NEWPIECE }
function EncodeRefPiece(TOBPL: TOB; NewNum: Integer = 0; WithNumLigne : boolean=true): string; overload;
function EncodeRefPiece (Cledoc : r_cledoc) : string; overload;
procedure DecodeRefPiece(St: string; var CleDoc: R_CleDoc); { NEWPIECE }
procedure AjouteRepres(CodeComm, TypeCom: string; TOBComms: TOB);
procedure CommVersLigne(TOBPiece, TOBArticles, TOBComms: TOB; ARow: integer; Totale: boolean);
function ExistePiece(CD: R_CleDoc): Boolean;
procedure ShowPiecePasHisto(CD: R_CleDoc; Prec: Boolean);
// Modif BTP
function GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs: TOB; Action: TActionFiche; DEV: RDevise; Ouv: boolean): Boolean;
// -----
{$IFDEF CHR}
function GereReglementsChr(TOBPiece, TOBTiers, TOBEches, TOBAcomptes: TOB; Action: TActionFiche; DEV: RDevise; Ouv: boolean; Famreg: string; TypeImpFac:
  integer;ModificationTicket:boolean = False): Boolean;
{$ENDIF}
procedure ChangeTauxAcomptes(TOBPiece, TOBAcomptes: TOB; DEV: RDEVISE);
procedure ChangeTauxEches(TOBPiece, TOBEches: TOB; DEV: RDEVISE);
// modif BTP
procedure ProrateEchesGC(TOBPiece, TOBEches, TOBAcomptes, TOBPieceRG, TOBPieceTrait, TOBPOrcs: TOB; DEV: RDEVISE);
// --
function PieceEncoreVivante(TOBPiece: TOB): Boolean;
function GCPieceCorrecte(TOBPiece, TOBArticles, TOBCatalogu: TOB;Ecran : Tform;ModeBordereau : boolean = false; SaisieCM : Boolean=false) : integer;
function PositionneExige(TOBTiers: TOB): string;
procedure ClearAffaire(TOBP: TOB);
procedure DecoupeCodeAffsurTOB(TOBP: TOB);
// Modif BTP
function GCValideReglementOblig(TOBPiece, TobTiers, TOBPIeceRG: tob; DEV: RDEVISE): Integer;
// --
// Modif BTP
procedure ValideLesLiensOle(TOBPiece, TOBPiece_O, TOBLiens: TOB);
// -----

{ Nouvelle Gestion Reliquat - Divers }
function FindCleDocInTob(CleDoc: R_CleDoc; TheTob: Tob; MultiNiveau: Boolean = True): Tob;
function FindCleDocInTobPiece(CleDoc: R_CleDoc; TobPiece: Tob): Tob; { NEWPIECE }
function FindTobLigInOldTobPiece(TobL, TobPiece_O: Tob): Tob; { NEWPIECE }
function FindOldTobLigInTobPiece(TobL_O, TobPiece: Tob): Tob;
function EstLigneArticle(TobL: Tob): Boolean; { NEWPIECE }
function CleDoc2TobLigne(CleDocL: R_CleDoc; TobLigne: Tob): Boolean; { NEWPIECE }
function CompareCleDoc(Source, Cible: R_Cledoc): Boolean;

{ Nouvelle Gestion Reliquat }
function EstLigneSoldee(TobL: Tob; Fromvalidation : boolean=false): Boolean;
function ToutesLignesSoldees(TobPiece_O: Tob): Boolean; { NEWPIECE }

function UpdateResteALivrer(TobPiece, TobPiece_O, TobArticles, TobCatalogu, TOBOuvrage_O: Tob; Action: T_UpDateReliquat): Boolean; { NEWPIECE }
function UpdateEtatPiecePrecedente(CleDocLignePrecedente: R_CleDoc): Boolean; { NEWPIECE }
function GetTobsSuiv(TobLigne: Tob; TobLignesSuiv: Tob): Boolean;

// Ne pas supprimer, merci...
//function  GetTobsSuiv(TobLigne: Tob; TobLignesSuiv: Tob): Boolean; { NEWPIECE }
//procedure UpdateQuantiteReliquat(TobPiece, TobPiece_O: Tob); { NEWPIECE }

function VerifQuantitePieceSuivante(TobL: Tob; Quantite: double): Boolean; overload; { NEWPIECE }
function VerifQuantitePieceSuivante(TobPiece: Tob): Integer; overload; { NEWPIECE }
function WithPieceSuivante(TobLigne: Tob): Boolean; { NEWPIECE }
function WithPiecePrecedente(TobL: Tob): Boolean; { NEWPIECE }
procedure CompareTobPiece(TobPiece, TobPiece_O, TobDifference: Tob; OnlyWithPiecePrecedente: Boolean); { NEWPIECE }
procedure PutQteReliquat(TobL: Tob; OldQte: Double); { NEWPIECE }
procedure PutMTReliquat (TobL: Tob; OldMT : Double); { NEWPIECE }
function CanModifyLigne(TOBL: TOB; TransfoPiece: Boolean; TobPiece_O: TOB;Var WarningSupressionRecep : boolean; gestionreliquat: boolean=true;Action : TactionFiche=TaCreat): Boolean;
function CanDeletePiece(TobPiece, TobPiece_O: Tob; TransfoPiece: Boolean; var NumL: Integer;Action : TActionFiche = TaCreat): Boolean;
function  CanModifyPiece(TobPiece: Tob): Boolean; { NEWPIECE }
procedure RazResteALivrer(TobPiece: Tob); { NEWPIECE }
procedure RazResteALivrerAndKillPiece(TobPiece_O: Tob); { NEWPIECE }
function  GetMaxIndice(CleDoc: R_CleDoc): Integer; { NEWPIECE }
function CalculResteALivrer(TOBL: Tob; AddTobPiece: Tob = nil): Double;
procedure GereLesLignesSoldees(TobPiece, TobPiece_O, TobArticles, TobCatalogu, TOBOuvrage_O: Tob);
procedure AjouteEtatLignePrecedente(TOBPiece, TOBPiece_O: TOB; TransfoPiece: Boolean; Action: TActionFiche);
function GetTOBPrecBySQL(TOBL: TOB; Select : String = '*'): TOB;
{ Gestion de la table LIGNECOMPL }
function  MakeSelectLigne: String;
procedure ValideLesLignesCompl(TobPiece, TobPiece_O: Tob);
function  ExisteLigneCompl(CleDocLigne: R_CleDoc): Boolean;
Function OKRepresentant (CodeCom,TypeCom : string) : boolean;
function GetPrefixeTable (TOBL : TOB) : string;
function IsgestionAvanc (TOBpiece : TOB) : boolean;
function ISFactureCloture (Cledoc : r_cledoc; var DateFin : String ) : boolean;
procedure PieceToEcheGC(TOBPiece, TOBH: TOB);
procedure ReinitMode(Mode :T_ModeRegl; Niveau : integer);
// Nouveau regroupement de pièce
function UpdateTouslesAncien (TOBListPiece, TOBPiece, TOBAcomptes_O, TobTiers: TOB; ViaSQL: Boolean = False; NewNum: Integer = 0): boolean;
procedure InverseStockTransfoMultiple (TOBlistPieces,TobPiece, TOBArticles, TOBCatalogu, TOBOuvrage_O: TOB; Litige: Boolean = False);
procedure RazResteALivrerAndKillLesPieces(TOBListPieces : TOB);
procedure LibelleCommercial(CodeRep: string; var lib1, lib2: string; IsRessource : Boolean=False);

function EncodeRefPieceSup(TOBPL: TOB): string; { NEWPIECE }
procedure ValideLesEvtsMateriels(TOBPiece,TOBEvtsmat : TOB);
procedure LoadlesEvtsmat (TOBpiece,TOBEvtsmat : TOB);
procedure RecupEvtsMateriels (TOBL,TOBEvtsmat : TOB);
procedure SetInfosFromEvts  (TOBL,TOBEvtsmat : TOB);
procedure EnleveEvtsMateriels (TOBL,TOBEvtsMat : TOB);
procedure DeleteEvtsMateriels (TOBL,TOBEvtsMat : TOB);
function GetTauxTaxe (RegimeTaxe,FamilleTaxe : string; Achat : boolean=false) : double;
function ExistRefBsv (TOBpiece : TOB) : boolean;
function GetIDBSVDOC (TOBpiece : TOB) : string;


type T_ValideNumPieceVide = class
    NaturePiece    : string;
    Souche    : string;
    DatePiece : TdateTime;
    NewNum    : Integer;
    procedure ValideNumPieceVide;
  end;

var RelancerEtat: Boolean;

implementation


Uses
  	 FactVariante,
     FactCpta,
     FactRg,
     UTOFMBOValideModele,
     FactSpec,
{$IFDEF BTP}
		 FactEmplacementLivr,
     AppelsUtil,
	   CalcOleGenericBTP,
     BTPUtil,
     FactOuvrage,
     FactRgpBesoin,
{$ENDIF}
{$IFDEF FOS5}
     TickUtilFO,
{$ENDIF}
{$IFDEF GPAO}
     uordresttype,
{$ENDIF GPAO}
{$IFDEF GPAOLIGHT}
     wOrdreCmp,
{$ENDIF GPAOLIGHT}
  	 FactCalc,
     FactTarifs,
     BTStructChampSup,
     factgrp,
     FactAdresse,
     FactAdresseBTP,
     AffaireUtil,
     UtilsRapport,
     Facture,
     UCotraitance,
     DateUtils,
     factRetenues
   ,CbpMCD
    ,CbpEnumerator
    ,UTransferts
     ;

function GetTauxTaxe (RegimeTaxe,FamilleTaxe : string; Achat : boolean=false) : double;
var TOBT : TOB;
begin
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],['TX1',RegimeTaxe,FamilleTaxe],False) ;
  if TOBT<>Nil then
  BEGIN
    Result := 0;
    if Achat then
    begin
      result:=TOBT.GetValue('TV_TAUXACH')
    end else
    begin
      result:=TOBT.GetValue('TV_TAUXVTE')
    end;
  END;
end;


function ISFactureCloture (Cledoc : r_cledoc; var DateFin : String ) : boolean;
var QQ : TQuery;
    PiecePrec : string;
    fCledoc : r_cledoc;
begin
  result := false;
  PiecePrec := '';
  QQ := OpenSql ('SELECT ##TOP 1## GL_PIECEPRECEDENTE FROM LIGNE WHERE '+WherePiece(Cledoc,ttdLigne,false)+' AND GL_ARTICLE <> "" AND GL_PIECEPRECEDENTE <> ""',true,1,'',true);
  if not QQ.eof then
  begin
    PiecePrec := QQ.FindField('GL_PIECEPRECEDENTE').AsString;
  end;
  ferme (QQ);
  if PiecePrec = '' then exit;
  // --
  DecodeRefPiece(PiecePrec,fCledoc);
  QQ := OpenSql ('select AFF_ETATAFFAIRE, AFF_DATEFIN FROM AFFAIRE WHERE AFF_AFFAIRE = (SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+WherePiece(fCledoc,ttdPiece,false)+')',true,1,'',true);
  if not QQ.eof then
  begin
    result := (QQ.findField('AFF_ETATAFFAIRE').asString='TER');
    DateFin := QQ.findField('AFF_DATEFIN').asString;
  end;
  ferme (QQ);
end;


procedure ClearAffaire(TOBP: TOB);
begin
  if TOBP = nil then Exit;
  if TOBP.NomTable = 'PIECE' then
  begin
    TOBP.PutValue('GP_AFFAIRE', '');
    TOBP.PutValue('GP_AFFAIRE1', '');
    TOBP.PutValue('GP_AFFAIRE2', '');
    TOBP.PutValue('GP_AFFAIRE3', '');
    TOBP.PutValue('GP_AVENANT', '');
  end else if TOBP.NomTable = 'LIGNE' then
  begin
    TOBP.PutValue('GL_AFFAIRE', '');
    TOBP.PutValue('GL_AFFAIRE1', '');
    TOBP.PutValue('GL_AFFAIRE2', '');
    TOBP.PutValue('GL_AFFAIRE3', '');
    TOBP.PutValue('GL_AVENANT', '');
  end;
end;

procedure DecoupeCodeAffsurTOB(TOBP: TOB);
var Prefixe, CodeAff: string;
    {$IFDEF AFFAIRE}
    Aff0, Aff1, Aff2, Aff3, Aff4: string;
    {$ENDIF}
begin
  Prefixe := '';
  if TOBP = nil then Exit;
  if TOBP.NomTable = 'PIECE' then Prefixe := 'GP_' else
    if TOBP.NomTable = 'LIGNE' then Prefixe := 'GL_';
  if Prefixe = '' then Exit;
  CodeAff := TOBP.GetValue(Prefixe + 'AFFAIRE');

  {$IFDEF BTP}
  BTPCodeAffaireDecoupe(CodeAff, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
  {$ELSE}
  CodeAffaireDecoupe(CodeAff, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
  {$ENDIF}
  TOBP.PutValue(Prefixe + 'AFFAIRE1', Aff1);
  TOBP.PutValue(Prefixe + 'AFFAIRE2', Aff2);
  TOBP.PutValue(Prefixe + 'AFFAIRE3', Aff3);
  TOBP.PutValue(Prefixe + 'AVENANT', Aff4);

end;

procedure InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire: TOB; ARow: integer; AvecCalcul: boolean = false);
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  InitLigneVide(TOBPiece, TOBL, TOBTiers, TOBAffaire, ARow, 1);
  //Init: pas en contremarque
  TOBL.PutValue('GL_ENCONTREMARQUE', '-');
  if AvecCalcul then
  begin
    TOBPiece.PutValue('GP_RECALCULER', 'X');
    TOBL.PutValue('GL_RECALCULER', 'X');
  end;
{$IFDEF BTP}
//	TOBL.PutValue('GL_BLOQUETARIF',TOBPiece.GetValue('_BLOQUETARIF'));
{$ENDIF}
  if TOBL.FieldExists ('BLF_MTMARCHE') then
  begin
    TOBL.putvalue('BLF_MTMARCHE',0);
    TOBL.putvalue('BLF_MTDEJAFACT',0);
    TOBL.putvalue('BLF_MTCUMULEFACT', 0);
    TOBL.putvalue('BLF_MTSITUATION', 0);
    TOBL.putvalue('BLF_POURCENTAVANC',0);
  end;
  // init domaine d'activité
  if ARow > 1 then TOBL.PutValue('GL_DOMAINE', TOBPiece.Detail[ARow - 2].GetValue('GL_DOMAINE'));
  if TOBPiece.getValue('GP_DOMAINE')<>'' then TOBL.PutValue('GL_DOMAINE', TOBPiece.GetValue('GP_DOMAINE'));
end;

procedure CreerTOBLignes(GS: THGrid; TOBPiece, TOBTiers, TOBAffaire: TOB; ARow: integer; AvecCalcul: boolean = false; TheDestinationLiv : TObject=nil );
var NumL: integer;
{$IFDEF BTP}
	TOBL : TOB;
{$ENDIF}
begin
  while TOBPiece.Detail.Count < ARow do
  begin
    TOBL := NewTOBLigne(TOBPiece, 0);
    NumL := TOBPiece.Detail.Count;
    GS.Cells[SG_NL, NumL] := IntToStr(NumL);
    InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, NumL, AvecCalcul);
{$IFDEF BTP}
		if (TheDestinationLiv <> nil) and (TheDestinationLiv is TDestinationLivraison) then TdestinationLivraison(TheDestinationLiv).SetDestLivrDefaut (TOBL);
{$ENDIF}
  end;
end;

function GetTOBPrec(TOBL, TOBPiece_O: TOB; AvecCommentDim: Boolean = False): TOB; { NEWPIECE }
var StPrec, RefUnique, RefCata, TypeL, TypeD: string;
  CD: R_CleDoc;
//    NumL : integer ; { NEWPIECE }
begin
  Result := nil;
  RefCata := '';
  if ((TOBL = nil) or (TOBPiece_O = nil)) then Exit;
  TypeL := TOBL.GetValue('GL_TYPELIGNE');
  // VARIANTE
  if IsVariante(TOBL) then exit;
  if ((not AvecCommentDim) or (not IsCommentaire(TOBL))) then
  begin
    if TOBL.GetValue('GL_TYPEREF') = 'CAT' then RefCata := TOBL.GetValue('GL_REFCATALOGUE');
    RefUnique := TOBL.GetValue('GL_ARTICLE');
    if (RefUnique = '') and (RefCata = '') then Exit;
  end else
  begin
    TypeD := TOBL.GetValue('GL_TYPEDIM');
    if TypeD <> 'GEN' then Exit;
  end;
  StPrec := TOBL.GetValue('GL_PIECEPRECEDENTE');
  if StPrec = '' then Exit;
  DecodeRefPiece(StPrec, CD); // NumL:=CD.NumLigne ; { NEWPIECE }
  Result := FindCleDocInTobPiece(CD, TobPiece_O);
  //if ((NumL<1) or (NumL>TOBPiece_O.Detail.Count)) then Exit ; { NEWPIECE }
  //Result:=TOBPiece_O.Detail[NumL-1] ; { NEWPIECE }
end;

function DetruitAncien(TOBPiece_O, TOBBases_O, TOBEches_O, TOBN_O, TOBLOT_O, TOBAcomptes_O, TOBPorcs_O, TOBSerie_O, TOBOuvrage_O, TOBLienOle_O, TOBPieceRG_O, TOBBasesRg_O, TobLigneTarif_O: TOB; SuppTOX: Boolean = False; OkSuppLigneFac : Boolean=true) : boolean; overload;
var CleDoc: R_CleDoc;
  Nb: integer;
  RefA: string;
  OldD: TDateTime;
begin
  Result := False;
  CleDoc := TOB2CleDoc(TOBPiece_O);
  OldD := TOBPiece_O.GetValue('GP_DATEMODIF');
  TRY
    Nb := ExecuteSQL('DELETE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False) + ' AND GP_DATEMODIF="' + USTime(OldD) + '"');
  if Nb <= 0 then
  begin
    V_PGI.IoError := oeSaisie;
    Exit;
  end;
  if TOBPiece_O <> nil then if TOBPiece_O.Detail.Count > 0 then
    begin
      Nb := ExecuteSQL('DELETE FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False));
      if Nb <= 0 then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
      Nb := ExecuteSQL('DELETE FROM LIGNECOMPL WHERE ' + WherePiece(CleDoc, ttdLigneCompl, False));
    end;
  Nb := ExecuteSQL('DELETE FROM PIECETRAIT WHERE ' + WherePiece(CleDoc, ttdPieceTrait, False));
  Nb := ExecuteSQL('DELETE FROM PIECEINTERV WHERE ' + WherePiece(CleDoc, ttdPieceInterv, False));
  Nb := ExecuteSQL('DELETE FROM PIEDBASE WHERE ' + WherePiece(CleDoc, ttdPiedBase, False));
  Nb := ExecuteSQL('DELETE FROM TIMBRESPIECE WHERE ' + WherePiece(CleDoc, ttdTimbres, False));
    Nb := ExecuteSQL('DELETE FROM PIEDCOLLECTIF WHERE ' + WherePiece(CleDoc, ttdVteColl, False));
  Nb := ExecuteSQL('DELETE FROM LIGNEBASE WHERE ' + WherePiece(CleDoc, ttdLigneBase, False));
  if TOBEches_O <> nil then if TOBEches_O.Detail.Count > 0 then
    begin
      Nb := ExecuteSQL('DELETE FROM PIEDECHE WHERE ' + WherePiece(CleDoc, ttdEche, False));
      if Nb <= 0 then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBN_O <> nil then if TOBN_O.Detail.Count > 0 then
    begin
      Result := TOBN_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBLOT_O <> nil then if TOBLOT_O.Detail.Count > 0 then
    begin
      Result := TOBLOT_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBSerie_O <> nil then if TOBSerie_O.Detail.Count > 0 then
    begin
      Result := TOBSerie_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBAcomptes_O <> nil then if TOBAcomptes_O.Detail.Count > 0 then
    begin
      Result := TOBAcomptes_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBPorcs_O <> nil then if TOBPorcs_O.Detail.Count > 0 then
    begin
      Result := TOBPorcs_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;

  if TOBLigneTarif_O <> nil then if TOBLigneTarif_O.Detail.Count > 0 then
  begin
    Result := TOBLigneTarif_O.DeleteDB;
    if not Result then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
  end;

  // Modif BTP
  if TOBLIENOLE_O <> nil then if TOBLIENOLE_O.Detail.Count > 0 then
    begin
      Result := TOBLIENOLE_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
  if TOBOuvrage_O <> nil then if TOBOuvrage_O.Detail.Count > 0 then
    begin
  (*
      Result := TOBOuvrage_O.DeleteDB;
      if not Result then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
  *)
      Nb := ExecuteSQL('DELETE FROM LIGNEOUV WHERE ' + WherePiece(CleDoc, ttdOuvrage, False));
      if Nb <= 0 then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;


  //Suppression des demandes de prix
  ExecuteSQL('DELETE FROM PIECEDEMPRIX WHERE ' + WherePiece(CleDoc, ttdPieceDemPrix , False));
  ExecuteSQL('DELETE FROM DETAILDEMPRIX WHERE ' + WherePiece(CleDoc, TtdDetailDemPrix , False));
  ExecuteSQL('DELETE FROM ARTICLEDEMPRIX WHERE ' + WherePiece(CleDoc, TTdArticleDemPrix , False));
  ExecuteSQL('DELETE FROM FOURLIGDEMPRIX WHERE ' + WherePiece(CleDoc, TtdFournDemprix , False));
  //

  ExecuteSQL('DELETE FROM LIGNEOUVPLAT WHERE ' + WherePiece(CleDoc, ttdOuvrageP , False));
  //
  ExecuteSQL('DELETE FROM BLIGNEMETRE WHERE ' + WherePiece(CleDoc, ttdLigneMetre , False));
  //
    if (Pos (cledoc.NaturePiece,'FBT;DAC;FBP;BAC') > 0) and (OkSuppLigneFac) then
  begin
    ExecuteSQL('DELETE FROM LIGNEFAC WHERE ' + WherePiece(CleDoc, ttdLigneFac , False));
  end;
  //
  Nb := ExecuteSQL('DELETE FROM PIECERG WHERE ' + WherePiece(CleDoc, ttdRetenuG, False));
  Nb := ExecuteSQL('DELETE FROM PIEDBASERG WHERE ' + WherePiece(CleDoc, ttdBaseRG, False));
  //
  ExecuteSQL('DELETE FROM BTPIECEMILIEME WHERE ' + WherePiece(CleDoc, TTdRepartmill , False));
  // --------
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    ExecuteSQL('DELETE FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc, ttdPieceAdr, False));
  end;
  //Nb := ExecuteSQL('DELETE FROM BREVISIONS WHERE ' + WherePiece(CleDoc, ttdRevision, False));

  RefA := EncodeRefPresqueCPGescom(TOBPiece_O);
  ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');
  if isComptaStock(TOBPiece_O.GetValue('GP_NATUREPIECEG')) then ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GS" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');
  RefA := EncodeRefpieceSup(TOBPiece_O);
  ExecuteSQL('UPDATE BTEVENTMAT SET BEM_CODEETAT="ARE",BEM_REFPIECE="",BEM_LIBREALISE="" WHERE BEM_REFPIECE LIKE "'+RefA+'"');
  DeleteLesTOBTRF (cledoc);
  Result := True;
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Suppression ancienne pièce');
    end;
  END;
end;

function DetruitAncien(Cledoc : r_cledoc; DateModif : TdateTime): boolean; overload;
var
  Nb: integer;
  RefA: string;
  OldD: TDateTime;
begin
  Result := False;
  OldD := DateModif;
  if OldD <> iDate1900 then
  begin
  	Nb := ExecuteSQL('DELETE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False) + ' AND GP_DATEMODIF="' + USTime(OldD) + '"');
  end else
  begin
  	Nb := ExecuteSQL('DELETE FROM PIECE WHERE ' + WherePiece(CleDoc, ttdPiece, False) );
  end;
  if Nb <= 0 then
  begin
    V_PGI.IoError := oeSaisie;
    Exit;
  end;
  Nb := ExecuteSQL('DELETE FROM LIGNE       WHERE ' + WherePiece(CleDoc, ttdLigne, False));
  Nb := ExecuteSQL('DELETE FROM LIGNECOMPL  WHERE ' + WherePiece(CleDoc, ttdLigneCompl, False));
  Nb := ExecuteSQL('DELETE FROM PIEDBASE    WHERE ' + WherePiece(CleDoc, ttdPiedBase, False));
  Nb := ExecuteSQL('DELETE FROM LIGNEBASE   WHERE ' + WherePiece(CleDoc, ttdLigneBase, False));
  Nb := ExecuteSQL('DELETE FROM PIEDECHE    WHERE ' + WherePiece(CleDoc, ttdEche, False));
  Nb := ExecuteSQL('DELETE FROM LIGNENOMEN  WHERE ' + WherePiece(CleDoc, ttdNomen, False));
  Nb := ExecuteSQL('DELETE FROM ACOMPTES    WHERE ' + WherePiece(CleDoc, ttdAcompte, False));
  Nb := ExecuteSQL('DELETE FROM PIEDPORT    WHERE ' + WherePiece(CleDoc, ttdPorc, False));
  //Nb := ExecuteSQL('DELETE FROM PORT        WHERE ' + WherePiece(CleDoc, ttdPorc, False));
  Nb := ExecuteSQL('DELETE FROM LIENSOLE    WHERE ' + WherePiece(CleDoc, ttdLienOle , False));
  Nb := ExecuteSQL('DELETE FROM LIGNEOUV    WHERE ' + WherePiece(CleDoc, ttdOuvrage, False));

  //Suppression des demandes de prix
  ExecuteSQL('DELETE FROM PIECEDEMPRIX WHERE ' + WherePiece(CleDoc, ttdPieceDemPrix , False));
  ExecuteSQL('DELETE FROM DETAILDEMPRIX WHERE ' + WherePiece(CleDoc, TtdDetailDemPrix , False));
  ExecuteSQL('DELETE FROM ARTICLEDEMPRIX WHERE ' + WherePiece(CleDoc, TTdArticleDemPrix , False));
  ExecuteSQL('DELETE FROM FOURLIGDEMPRIX WHERE ' + WherePiece(CleDoc, TtdFournDemprix , False));
  //

  ExecuteSQL('DELETE FROM LIGNEOUVPLAT WHERE ' + WherePiece(CleDoc, ttdOuvrageP , False));
  //
  ExecuteSQL('DELETE FROM BLIGNEMETRE WHERE ' + WherePiece(CleDoc, ttdLigneMetre , False));
  ExecuteSQL('DELETE FROM LIGNEFAC WHERE ' + WherePiece(CleDoc, ttdLigneFac , False));
  Nb := ExecuteSQL('DELETE FROM PIECERG WHERE ' + WherePiece(CleDoc, ttdRetenuG, False));
  Nb := ExecuteSQL('DELETE FROM PIEDBASERG WHERE ' + WherePiece(CleDoc, ttdBaseRG, False));
  Nb := ExecuteSQL('DELETE FROM PIECEINTERV WHERE ' + WherePiece(CleDoc, ttdPieceInterv, False));
  //
  ExecuteSQL('DELETE FROM BTPIECEMILIEME WHERE ' + WherePiece(CleDoc, TTdRepartmill , False));
  // --------
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    ExecuteSQL('DELETE FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc, ttdPieceAdr, False));
  end;
  RefA := EncodeRefPresqueCPGescom(cledoc);
  ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');
  if isComptaStock(Cledoc.NaturePiece) then ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GS" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');
  Result := True;
end;

function ToutesLignesSoldees(TOBPiece_O: Tob): Boolean; { NEWPIECE }
var
  i: Integer;
  ArtOk : Boolean;
begin
  Result := True;
  ArtOk := False;
  i := -1;
  while (i < TOBPiece_O.Detail.Count - 1) and Result do
  begin
    Inc(i);
    if EstLigneArticle(TOBPiece_O.Detail[i]) then
      begin
      ArtOk := True;
      Result := EstLigneSoldee(TOBPiece_O.Detail[i],true);
      end;
  end;
  if ArtOk = False then Result := False;
end;

{ NEWPIECE }
function UpdateTouslesAncien (TOBListPiece, TOBPiece, TOBAcomptes_O, TobTiers: TOB; ViaSQL: Boolean = False; NewNum: Integer = 0): boolean;
var II : Integer;
begin
  Result := True;
  if  (GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
  begin
    result := true;
    exit;
  end;
	For II := 0 to TOBListPiece.Detail.count -1 do
  begin
    if not UpdateAncien (TOBListPiece.detail[II],TOBPiece,TOBAcomptes_O,TobTiers,ViaSQL,NewNum) then
    begin
      result := false;
    	break;
    end;
  end;

end;

function UpdateAncien(TOBPiece_O, TOBPiece, TOBAcomptes_O, TobTiers: TOB; ViaSQL: Boolean = False; NewNum: Integer = 0): boolean;
var OldNat, ActF, RefSuiv, StViv: string;
  Nb: integer;
  CleDoc: R_CLEDOC;
  ExisteModif: Boolean;
begin
  {$IFDEF AFFAIRE}
  // gm 16/10/03
  // permet de ne pas  modifier la piece d'origine qd on transforme une facture en avoir.
  // Utiliser en GA dans la génération automatique d'avoir à partir de factures
  // sera généraliser si besoin après la V5.0
  if  (GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
  begin
    result := true;
    exit;
  end;
  {$ENDIF}
  ExisteModif := TOBPiece_O.Modifie;
  RefSuiv := EncodeRefPiece(TOBPiece, NewNum);
  StViv := 'X';
  TOBPiece_O.PutValue('GP_DEVENIRPIECE', RefSuiv);
  TOBPiece_O.PutValue('GP_REFCOMPTABLE', '');
  OldNat := TOBPiece_O.GetValue('GP_NATUREPIECEG');
  ActF := GetInfoParPiece(OldNat, 'GPP_ACTIONFINI');
  if ActF = 'TRA' then
  begin
    PutValueDetail(TOBPiece_O, 'GP_VIVANTE', '-');
    StViv := '-';
  end;
  BeforeUpdateAncien(TOBPiece_O, TOBPiece, TOBAcomptes_O, TobTiers);
  if ((ViaSQL) and (not ExisteModif)) then
  begin
    CleDoc := TOB2CleDoc(TOBPiece_O);
    Result := True;
    Nb := ExecuteSQL('UPDATE PIECE SET GP_DEVENIRPIECE="' + RefSuiv + '", GP_REFCOMPTABLE="", GP_VIVANTE="' + StViv + '" WHERE ' + WherePiece(CleDoc, ttdPiece,
      False));
    if Nb <> 1 then Result := False else if StViv = '-' then
    begin
      Nb := ExecuteSQL('UPDATE LIGNE SET GL_VIVANTE="-" WHERE ' + WherePiece(CleDoc, ttdLigne, False));
      if ((ctxAffaire in v_PGI.PGIContexte) or (VH_GC.GASeria)) and ((TOBPiece_O.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatAffaire) or
        (TOBPiece_O.GetValue('GP_NATUREPIECEG') = VH_GC.AFNatProposition)) then
      begin
        if ((Nb < 0) or (Nb > TOBPiece_O.Detail.Count)) then Result := False;
      end
      else
        if ((Nb <= 0) or (Nb > TOBPiece_O.Detail.Count)) then Result := False;
    end;
  end else
  begin
 // Mise à jour de la pièce d'origine effectuée plus tard
    Result := True;
  end;
  if not Result then BEGIN MessageValid := 'Erreur Lors de la mise à jour du document précédent'; V_PGI.IoError := oeSaisie; END;
  if ((TOBAcomptes_O <> nil) and (Result)) then
  begin
    Result := TOBAcomptes_O.DeleteDB;
    if not Result then BEGIN MessageValid := 'Erreur Lors de la suppression des acomptes précédents';V_PGI.IoError := oeSaisie; END;
  end;
end;

procedure ValideLesLiensOle(TOBPiece, TOBPiece_O, TOBLiens: TOB);
var indice: integer;
  T: TOB;
  numpiece: string;
  okok : boolean;
begin
  Numpiece := TOBPiece.GetValue('GP_NATUREPIECEG') + ':' + TOBPiece.GetValue('GP_SOUCHE') + ':' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ':' +
    IntToStr(TOBPiece.GetValue('GP_INDICEG'));
  for indice := 0 to TOBLiens.detail.count - 1 do
  begin
    T := TOBLIens.detail[Indice];
    T.PutValue('LO_IDENTIFIANT', NumPiece);
  end;
  okok := false;
  TRY
    okok := TOBLiens.InsertOrUpdateDB(False);
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur écriture / LIENOLE');
    end;
  END;
  if not  okok then
  BEGIN
    V_PGI.IoError := oeUnknown;
  END;
end;

procedure ValideLesArticles(TOBPiece, TOBArticles: TOB);
var NaturePieceG: string;
begin
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
//  if GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO') = '' then exit;
  if ValideArticleEtStock(TOBArticles,Naturepieceg) then
  begin
    if GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO') <> '' then MajCatalogueSaisie(TOBArticles);
  end
  else
  begin
    V_PGI.IoError := oeUnknown;
  end;
end;

procedure ValideLesCatalogues(TOBPiece, TOBCatalogu: TOB);
var TOBDCM, TOBCATA, TOBTMP: TOB;
  Q: TQuery;
  i, j, MaxRg: integer;
  ToDel: boolean;
  Qte: double;
  stSqlDel: string;
begin
  Q := OpenSQL('SELECT MAX(GQC_RANG) FROM DISPOCONTREM', True,-1, '', True);
  if Q.EOF then MaxRg := 1 else MaxRg := Q.Fields[0].AsInteger + 1;
  Ferme(Q);
  stSqlDel := '';
  for j := 0 to TOBCatalogu.Detail.Count - 1 do
  begin
    TOBCATA := TOBCatalogu.Detail[j];
    for i := TOBCATA.Detail.Count - 1 downto 0 do
    begin
      TOBDCM := TOBCATA.Detail[i];
      if TOBDCM.GetValue('GQC_RANG') <= 0 then
      begin
        TOBDCM.PutValue('GQC_RANG', MaxRg);
        MaxRg := MaxRg + 1;
      end;
      //Suppression des lignes avec Phy=0 ResCli=0 et ResFou=0
      //Passer les lignes à SansClient si Phy<>0 et ResCli=0 et ResFou=0
      ToDel := False;
      if (TOBDCM.GetValue('GQC_CLIENT') = '') then
      begin
        if (TOBDCM.GetValue('GQC_PHYSIQUE') = 0) and (TOBDCM.GetValue('GQC_RESERVEFOU') = 0) then
          ToDEL := True;
      end else
      begin
        Qte := TOBDCM.GetValue('GQC_PHYSIQUE') - (TOBDCM.GetValue('GQC_RESERVECLI') + TOBDCM.GetValue('GQC_PREPACLI'));
        if Qte > 0 then
        begin
          TOBTMP := TOBCata.FindFirst(['GQC_REFERENCE', 'GQC_FOURNISSEUR', 'GQC_CLIENT', 'GQC_DEPOT'],
            [TOBDCM.GetValue('GQC_REFERENCE'), TOBDCM.GetValue('GQC_FOURNISSEUR'),
            '', TOBDCM.GetValue('GQC_DEPOT')], False);
          if TOBTMP <> nil then
          begin
            TOBTMP.PutValue('GQC_PHYSIQUE', TOBTMP.GetValue('GQC_PHYSIQUE') + Qte);
          end else
          begin
            TOBTMP := TOB.Create('', TOBCATA, -1);
            TOBTMP.Dupliquer(TOBDCM, True, True);
            TOBTMP.PutValue('GQC_RESERVECLI', 0);
            TOBTMP.PutValue('GQC_RESERVEFOU', 0);
            TOBTMP.PutValue('GQC_PREPACLI', 0);
            TOBTMP.PutValue('GQC_PHYSIQUE', Qte);
            TOBTMP.PutValue('GQC_CLIENT', '');
            TOBTMP.PutValue('GQC_RANG', MaxRg);
            MaxRg := MaxRg + 1;
          end;
          TOBDCM.PutValue('GQC_PHYSIQUE', TOBDCM.GetValue('GQC_PHYSIQUE') - Qte);
        end;
        if (TOBDCM.GetValue('GQC_PHYSIQUE') = 0) and (TOBDCM.GetValue('GQC_RESERVEFOU') = 0) and
          (TOBDCM.GetValue('GQC_RESERVECLI') = 0) and (TOBDCM.GetValue('GQC_PREPACLI') = 0) then
          ToDEL := True;
      end;
      if (TOBDCM.GetValue('GQC_PHYSIQUE') < 0) then
      begin
        V_PGI.IoError := oeStock;
        exit;
      end;

      if ToDel then
      begin
        if stSqlDel <> '' then stSqlDel := stSqlDel + ' OR ';
        stSqlDel := stSqlDel + 'GQC_RANG=' + IntToStr(TOBDCM.GetValue('GQC_RANG'));
        TOBDCM.free;
      end;
    end;
  end;

  if stSqlDel <> '' then
  begin
    ExecuteSql('DELETE FROM DISPOCONTREM WHERE ' + stSqlDel);
  end;

  if not TOBCatalogu.InsertOrUpdateDB(False) then V_PGI.IoError := oeUnknown;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 16/10/2003
Description .. : Compare l'ancien et le nouveau GL_QTERESTE
Suite ........ : Attention ! On considère que la mise à jour du
Suite ........ : GL_QTERESTE n'a pas encore été faite....
*****************************************************************}
{*
function GetReliquatQteReste(TobLigPiecePrecedente, NewTobLigPiece, OldTobLigPiece: Tob): Double;
var
  OldQteReste, Qte, NewQteReste,Reliquat: Double;
  OldMtReste  : Double;
  MtReste     : Double;
  MtReliquat  : Double;
  OK_Reliquat : Boolean;
  NewMtReste  : Double;
begin

  Result  := 0;
  MtReste := 0;
  Qte     := 0;
  //
  NewQteReste := 0;
  NewMtReste  := 0;

  if Assigned(TobLigPiecePrecedente) and Assigned(NewTobLigPiece) then
  begin
{V500_005
//    OldQteReste := Max(0.0, Double(TobLigPiecePrecedente.G('GL_QTERESTE')));
		oldQteReste := TobLigPiecePrecedente.G('GL_QTERESTE');
    oldMtReste  := TobLigPiecePrecedente.G('GL_MTRESTE');
    OK_Reliquat := CtrlOkReliquat(TobLigPiecePrecedente, 'GL');
    if not NewTobLigPiece.FieldExists('SOLDERELIQUAT') then
    begin
      { Calcul du nouveau reste à livrer de la ligne d'origine
      if OldTobLigPiece = nil then
      begin
        Qte       := TobLigPiecePrecedente.G('GL_QTERESTE') - (NewTobLigPiece.G('GL_QTESTOCK'));
        Reliquat  := abs (TobLigPiecePrecedente.G('GL_QTERESTE')) - abs((NewTobLigPiece.G('GL_QTESTOCK')));
        if OK_Reliquat then
        begin
          MtReste   := TobLigPiecePrecedente.G('GL_MTRESTE')  - (NewTobLigPiece.G('GL_MONTANTHTDEV'));
          MtReliquat:= abs (TobLigPiecePrecedente.G('GL_MTRESTE')) - abs((NewTobLigPiece.G('GL_MONTANTHTDEV')));
        end;
      end
      else { Cas de la modification d'une ligne dans une pièce transformée (Modif. d'une ligne de livraison)
    	begin
        Qte       := TobLigPiecePrecedente.G('GL_QTERESTE') - OldTobLigPiece.G('GL_QTESTOCK') + NewTobLigPiece.G('GL_QTESTOCK');
        Reliquat  := Abs(TobLigPiecePrecedente.G('GL_QTERESTE')) - Abs(OldTobLigPiece.G('GL_QTESTOCK')) + Abs(NewTobLigPiece.G('GL_QTESTOCK'));
        if OK_Reliquat then
        begin
          MtReste   := TobLigPiecePrecedente.G('GL_MTRESTE')- OldTobLigPiece.G('GL_MONTANTHTDEV') + NewTobLigPiece.G('GL_MONTANTHTDEV');
          MtReliquat:= abs (TobLigPiecePrecedente.G('GL_MTRESTE')) - Abs(OldTobLigPiece.G('GL_MONTANTHTDEV')) + Abs(NewTobLigPiece.G('GL_MONTANTHTDEV'));
        end;
      end;
			if reliquat < 0 then qte := 0; //NewQteReste := Max(0 , Qte);
      NewQteReste := Qte;
      if OK_Reliquat then
      begin
        if MtReliquat < 0 then MtReste := 0;
        NewMtReste := MtReste;
      end;
    end;
    //
    NewTobLigPiece.PutValue('GL_MTRESTE', OldMtReste - NewMtReste);
    //
    Result := OldQteReste - NewQteReste;
    //
  end;

end;
*}

{ NEWPIECE }
{
  On transforme une commande en livraison
  Si pas transfo
    On annule complètement ce que la commande avait fait
  Si transfo
    Au niveau de Dispo on défait ce que la piece CC avait fait pour la quantité transformée uniquement
    Une ligne de commande de 10 / Une ligne de livraison de 3 :
      il faut décrémenter le compteur RC de 3 sur le dépot de la ligne de commande
}
procedure _InverseStock(TOBPiece_O, TOBPiece, TOBArticles, TOBCatalogu, TOBOuvrage_O: TOB; Transfo: Boolean; Litige: Boolean = False; ControleStock : boolean=true);
var
  i, IndiceNomen, iArt, iTypeRef: integer;
  TheTobPiece, TOBL, TOBA, TOBOuvrage, TOBCata: TOB;
  NewTobLig   : TOB;
  OldQteStock : Double;
  OldMtStock  : Double;
  NaturePieceG, ColPlus, ColMoins, StPrixValo, RefUnique: string;
  RatioVA, RatioVente: Double;
  DepotRecepteurGererSurSite: Boolean;
  OldCleDoc       : R_CleDoc;
  Maj             : Boolean;
  Centralisateur  : boolean;
  Ok_ReliquatMt   : Boolean;
begin

  NaturePieceG := TobPiece_O.GetValue('GP_NATUREPIECEG');

  TheTobPiece := TobPiece_O;

  DepotRecepteurGererSurSite := True;

  ColMoins := GetInfoParPiece(NaturePieceG, 'GPP_QTEMOINS');
  ColPlus := GetInfoParPiece(NaturePieceG, 'GPP_QTEPLUS');
  StPrixValo := GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO');

  for i := 0 to TheTobPiece.Detail.Count - 1 do
  begin
    TobL := TheTobPiece.Detail[i];
    if TobL = nil then
    begin
      V_PGI.IoError := oeUnknown;
      Break;
    end;
    // --- GUINIER ---
    Ok_ReliquatMt := CtrlOKReliquat(TOBL, 'GL');
    //
    Centralisateur := IsCentralisateurBesoin (TOBL); // si c'est le cas on ne fera pas les mvt de stock
    // Gestion des pieces d'avoir
    // MODIF_DBR_DEBUT // Pas top de le faire avant tout (dans validelapiece), et de le défaire ici
    if (GetInfoParPiece(TOBPiece_O.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
    begin
      InverseX(TOBL,'GL_QTESTOCK') ;
      InverseX(TOBL,'GL_QTEFACT') ;
      InverseX(TOBL,'GL_QTERESTE') ;
      InverseX(TOBL,'GL_MTRESTE') ;  //--- GUINIER---
    end;
    // -----
    NewTobLig := nil;
    OldQteStock := TobL.G('GL_QTESTOCK');
    if Ok_ReliquatMt then OldMtStock := TobL.G('GL_MONTANTHTDEV');
//    Maj := True;
		Maj := not Centralisateur;
    if i = 0 then
    begin
      iArt := Tobl.GetNumChamp('GL_ARTICLE');
      iTyperef := Tobl.GetNumChamp('GL_TYPEREF');
    end;
    TOBA := nil;
    TOBCATA := nil;
    if Tobl.GetValeur(iTypeRef) = 'CAT' then
    begin
      TOBCata := FindTOBCataRow(TheTobPiece, TobCatalogu, i + 1);
      if TOBCata <> nil then
        TOBCata.PutValue('UTILISE', 'X');
    end
    else
    begin
      RefUnique := TOBL.GetValeur(iArt);
      if RefUnique = '' then Continue;
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
      ////
      if (TOBA = nil) then
      begin
        AjouteUnArticle ( TOBL,TOBArticles,TOBCaTalogu ) ;
      	TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [RefUnique], False);
      end;
    	/////
      if TOBA <> nil then
        TOBA.PutValue('UTILISE', 'X');
      if (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X') then // MR
      begin
        TOBCata := FindTOBCataRow(TheTobPiece, TobCatalogu, i + 1);
        if TOBCata <> nil then
          TOBCata.PutValue('UTILISE', 'X');
      end;
    end;
    { Nomenclatures }
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    // Correction LS
    if (IndiceNomen > 0) and (not isOuvrage(TOBL)) then BEGIN IndiceNomen := 0; TOBL.putValue('GL_INDICENOMEN',0); END;
    { Modif BTP }
    if (IndiceNomen > 0) and
       (TobOuvrage_O.detail.count > 0) and
       (IndiceNomen <= TobOuvrage_O.detail.count) and
       (Pos(TOBL.GetValue('GL_TYPEARTICLE'),'OUV;ARP')>0) then
    begin
      TOBOuvrage := TOBOuvrage_O.Detail[IndiceNomen - 1]
    end else
    begin
      TOBOuvrage := nil;
    end;
    if (TOBA <> nil) or (TOBCATA <> nil) then
    begin
      RatioVA := GetRatio(TOBL, nil, trsStock);
      RatioVente := GetRatio(TOBL, nil, trsVente);
      if Transfo and (not Litige) then
      begin
        if not TobPiece_O.FieldExists('GENAUTO') then
          NewTobLig := FindOldTobLigInTobPiece(TobL, TobPiece)
        else
        begin
          OldCleDoc := Tob2CleDoc(TobL);
          OldCleDoc.NumLigne := TobL.GetValue('GL_NUMLIGNE');
          OldCleDoc.NumOrdre := TobL.GetValue('GL_NUMORDRE');
          NewTobLig := FindCleDocInTob(OldCleDoc, TobPiece, True);
        end;
        if Assigned(NewTobLig) then
        begin
          // Gestion des pièces d'avoir
          if (GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
          begin
            InverseX(NewTobLig,'GL_QTESTOCK') ;
            InverseX(NewTobLig,'GL_QTEFACT')  ;
            InverseX(NewTobLig,'GL_QTERESTE') ;
            InverseX(NewTobLig,'GL_MTRESTE')  ; //---GUINIER---
          end;
          // -- gestion du solde de reliquat
          if (NewTobLig.FieldExists('SOLDERELIQUAT')) and (NewTobLig.getString('SOLDERELIQUAT')='X') then
          begin
            Tobl.P('GL_QTESTOCK',  TOBL.GetDouble('GL_QTERESTE'));
            Tobl.AddChampSupValeur('MTSTOCK',TOBL.GetDouble('GL_MTRESTE'));
          end
          else
          begin
            if NewTobLig.GetValue('GL_QTEFACT') > TOBL.GetDouble('GL_QTERESTE') then
              Tobl.P('GL_QTESTOCK',  TOBL.GetDouble('GL_QTERESTE')) // on ne peux pas restituer plus que la valeur de depart
            else
              Tobl.P('GL_QTESTOCK',  NewTobLig.GetValue('GL_QTEFACT'));
            // --- GUINIER ---
            if Ok_ReliquatMt then
            begin
              if NewTobLig.GetValue('GL_MONTANTHTDEV') > TOBL.GetDouble('GL_MTRESTE') then
                Tobl.P('MTSTOCK',  TOBL.GetDouble('GL_MTRESTE')) // on ne peux pas restituer plus que la valeur de depart
              else
                Tobl.P('MTSTOCK',  TOBL.GetDouble('GL_MONTANTHTDEV')); // on ne peux pas restituer plus que la valeur de depart
            end;
          end;
          // Gestion des pièces d'avoir
          if (GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
          begin
            InverseX(NewTobLig,'GL_QTESTOCK') ;
            InverseX(NewTobLig,'GL_QTEFACT') ;
            InverseX(NewTobLig,'GL_QTERESTE') ;
            // --- GUINIER ---
            InverseX(NewTobLig,'GL_MTRESTE') ;
          end;
        end
        else
          { Ligne supprimée dans la nouvelle pièce }
          Maj := False;
      end;

      // Gestion des pièces d'avoir
      if (GetInfoParPiece(TOBPiece_O.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
      begin
        InverseX(TOBL,'GL_QTESTOCK') ;
        InverseX(TOBL,'GL_QTEFACT') ;
        InverseX(TOBL,'GL_QTERESTE') ;
        // --- GUINIER ---
        InverseX(TOBL,'GL_MTRESTE') ;
      end;

      if Maj then
      begin
        if NaturePieceG = 'TRE' then
        begin
          if DepotRecepteurGererSurSite then
            MajQteStock(TobL, TOBArticles, TOBOuvrage, ColPlus, ColMoins, False, RatioVA,'',ControleStock);
        end
        else
        begin
          if TOBL.GetValue('GL_TYPEREF') <> 'CAT' then // MR
          begin
            if Transfo and (not Litige) then
            begin
              MajQteStock(TOBL, TOBArticles, TOBOuvrage, ColPlus, ColMoins, False, RatioVA, '',ControleStock);
            if TOBL.GetValue('GL_ENCONTREMARQUE') = 'X' then
                MajQteStockContreM(NewTobLig, TOBCatalogu, ColPlus, ColMoins, False, RatioVA);
            end else
            begin
              MajQteStock(TOBL, TOBArticles, TOBOuvrage, ColPlus, ColMoins, False, RatioVA, '',ControleStock);
              if TOBL.GetValue('GL_ENCONTREMARQUE') = 'X' then
              MajQteStockContreM(TOBL, TOBCatalogu, ColPlus, ColMoins, False, RatioVA);
            end;
          end
          else
            MajQteStockContreM(TOBL, TOBCatalogu, ColPlus, ColMoins, False, RatioVA);
        end;

        MajPrixValo(TOBL, TOBA, StPrixValo, False, RatioVA, RatioVente);
        if (TOBCATA <> nil) then
          MajPrixValoContreM(TOBL, TOBCata, StPrixValo, True, RatioVA, RatioVente);

        if Transfo and Assigned(NewTobLig) and (not Litige) then
        begin
      	// Gestion des pièces d'avoir
          if (GetInfoParPiece(TOBPiece_O.GetValue('GP_NATUREPIECEG'), 'GPP_ESTAVOIR') = 'X') then
          begin
            Tobl.P('GL_QTESTOCK', OldQteStock * -1);
            Tobl.P('MTSTOCK', OldQteStock * -1);
          end
          else
          begin
            Tobl.P('GL_QTESTOCK', OldQteStock);
            Tobl.P('MTSTOCK', OldMtStock);
          end;
        end;
      end;
    end
    else
    begin
      MessageValid := 'Erreur sur Article/ouvrage '+TOBL.GetValue('GL_CODEARTICLE');
      PGiError (messageValid);
      V_PGI.IoError := oeSaisie;
      Break;
    end;
  end;
end;

{ NEWPIECE }

procedure InverseStock(TOBPiece_O, TOBArticles, TOBCatalogu, TOBOuvrage_O: TOB; ControleStock : boolean=true);
begin
  _InverseStock(TOBPiece_O, nil, TOBArticles, TOBCatalogu, TOBOuvrage_O, False,False,ControleStock);
end;

{ NEWPIECE }

procedure InverseStockTransfoMultiple (TOBlistPieces,TobPiece, TOBArticles, TOBCatalogu, TOBOuvrage_O: TOB; Litige: Boolean = False);
var II : Integer;
begin
	for II := 0 to TOBlistPieces.detail.count -1 do
  begin
  	_InverseStock(TOBlistPieces.detail[II], TobPiece, TOBArticles, TobCatalogu, TOBOuvrage_O, True, Litige);
  end;
end;

procedure InverseStockTransfo(TOBPiece_O, TobPiece, TOBArticles, TOBCatalogu, TOBOuvrage_O: TOB; Litige: Boolean = False);
begin
  _InverseStock(TOBPiece_O, TobPiece, TOBArticles, TobCatalogu, TOBOuvrage_O, True, Litige);
end;

procedure ValideLaPeriode(TOBPiece: TOB);
var i, iper, isem: integer;
  TOBL: TOB;
  DD: TDateTime;
  Per, Sem: integer;
begin
  DD := TOBPiece.GetValue('GP_DATEPIECE');
  Per := GetPeriode(DD);
  Sem := NumSemaine(DD);
  TOBPiece.PutValue('GP_PERIODE', Per);
  TOBPiece.PutValue('GP_SEMAINE', Sem);
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i = 0 then
    begin
      iPer := TOBL.GetNumChamp('GL_PERIODE');
      iSem := TOBL.GetNumChamp('GL_SEMAINE');
    end;
    TOBL.PutValeur(iPer, Per);
    TOBL.PutValeur(iSem, Sem);
  end;
end;

procedure ValideLaCotation(TOBPiece, TOBBases, TOBEches: TOB);
var i: integer;
begin
  AttribCotation(TOBPiece);
  for i := 0 to TOBBases.Detail.Count - 1 do AttribCotation(TOBBases.Detail[i]);
  for i := 0 to TOBEches.Detail.Count - 1 do AttribCotation(TOBEches.Detail[i]);
end;

procedure ValideLesPorcs(TOBPiece, TOBPorcs: TOB);
var i, Numero, Indice: integer;
  Nature, Souche: string;
  TOBP: TOB;
  okok  : boolean;
begin
  if TOBPorcs = nil then Exit;
  if TOBPorcs.Detail.Count <= 0 then Exit;
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  for i := 0 to TOBPorcs.Detail.Count - 1 do
  begin
    TOBP := TOBPorcs.Detail[i];
    TOBP.PutValue('GPT_NATUREPIECEG', Nature);
    TOBP.PutValue('GPT_SOUCHE', Souche);
    TOBP.PutValue('GPT_NUMERO', Numero);
    TOBP.PutValue('GPT_INDICEG', Indice);
    TOBP.PutValue('GPT_NUMPORT', i + 1);
  end;
  okok := false;
  TRY
    okok := TOBPorcs.InsertDB(nil);
  EXCEPT
    on E: Exception do
  begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour des Ports/frais');
    end;
  END;
  if not okok then
  begin
    V_PGI.IoError := oeUnknown;
  end;
end;

function DepotGererSurSite(Depot: string): boolean;
var Q: TQuery;
begin
  result := False;
  Q := OpenSQL('Select GDE_SURSITE from DEPOTS WHERE GDE_DEPOT="' + Depot + '"', True,-1, '', True);
  if not Q.EOF then result := Q.FindField('GDE_SURSITE').ASString = 'X';
  Ferme(Q);
end;

procedure RestitueQteOrigine (TOBPiecePrec,TOBL,TOBARTICLES : TOB);
var Cledoc : R_Cledoc;
    TheLignePrec : TOB;
    ColPlus,ColMoins : string;
    RatioVa : Double;
begin

  If TOBPiecePrec = nil then exit;

  DecodeRefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),cledoc);
  TheLignePrec := TOBPiecePrec.findFirst (['GL_NUMORDRE'],[cledoc.NumOrdre],True);

  if TheLignePrec <> nil then
  begin
    // --- GUINIER ---
    TheLignePrec.PutValue('GL_QTERESTE', TheLignePrec.GetValue('GL_QTERESTE') + TOBL.GEtValue('GL_QTESTOCK'));
    if TheLignePrec.GetValue('GL_QTERESTE') > TheLignePrec.GetValue('GL_QTEFACT') then TheLignePrec.PutValue('GL_QTERESTE', TheLignePrec.GetValue('GL_QTEFACT'));
    if CtrlOkReliquat(TheLignePrec, 'GL') then TheLignePrec.PutValue('GL_MTRESTE', TheLignePrec.GetValue('GL_MONTANTHTDEV'));
    //
    if (TheLignePrec <> nil) and (TheLignePrec.GetValue('GL_TENUESTOCK')='X') then
    begin
      ColMoins := GetInfoParPiece(cledoc.Naturepiece, 'GPP_QTEMOINS');
      ColPlus := GetInfoParPiece(Cledoc.Naturepiece, 'GPP_QTEPLUS');
      RatioVA := GetRatio(TheLignePrec, nil, trsStock);
      MajQteStock (TheLignePrec,TOBArticles,nil,ColPlus,ColMoins,true,RatioVa);
    end;
  end;

end;

procedure ValideLesLignes(TOBPiece, TOBArticles, TOBCatalogu, TOBNomenclature, TOBOuvrages, TOBPieceRG, TOBBasesRG: TOB; Reliquat, Rafale: boolean; Ajout : boolean = false;GenereAuto:boolean=false;TOBOldPiece : TOB=nil; AutoriseDecoup : boolean=false);
var
  i: integer;
  TOBL, TOBA, TOBCATA, TOBNomen,TOBOuvrage: TOB;
  NaturePieceG, ActF, QualifMvt, DepotEmetteur: string;
  RatioVA, RatioVente: Double;
  IndiceNomen: integer;
  ColPlus, ColMoins, StPrixValo: string;
  DepotRecepteurGererSurSite, RetourMajStock: Boolean;
  IndiceOuv: Integer;
  {$IFDEF BTP}
  RGNonImprimable: string;
  ReliquatMt : Boolean;
  {$ENDIF}
  FerrorSD : boolean;
begin
  {$IFDEF BTP}
  if not (RGMultiple(TOBpiece)) then
    RGNonImprimable := 'X'
  else
    RGNonImprimable := '-';
  {$ENDIF}
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  DepotRecepteurGererSurSite := True;
  QualifMvt := GetInfoParPiece(NaturePieceG, 'GPP_QUALIFMVT');
  ActF := GetInfoParPiece(NaturePieceG, 'GPP_ACTIONFINI');
  ColPlus := GetInfoParPiece(NaturePieceG, 'GPP_QTEPLUS');
  ColMoins := GetInfoParPiece(NaturePieceG, 'GPP_QTEMOINS');
  StPrixValo := GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO');
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    FerrorSD  := false;
    ReliquatMt := False;
    TOBL := TOBPiece.Detail[i];

    if TOBL.GetString('GL_TYPELIGNE')='ART' then
    begin
      //
      //--- GUINIER : Gestion des Reliquats sur Montant pour Prestation
      //
      TOBA := FindTOBArtRow(TobPiece, TobArticles, i + 1);
      if TOBA <> nil then if TOBA.GetValue('GA_TYPEARTICLE') = 'PRE' then ReliquatMt := TOBA.GetBoolean('GA_RELIQUATMT');
    end;

    if ((TOBA <> nil) and (not Reliquat)) then TOBA.PutValue('UTILISE', 'X');

    TOBL.PutValue('GL_QUALIFMVT', QualifMvt);

    if TOBL.GetValue('GL_DEPOT') = ''     then TOBL.PutValue('GL_DEPOT', TOBPiece.GetValue('GP_DEPOT'));

    TOBL.PutValue('GL_TIERSFACTURE', TOBPiece.GetValue('GP_TIERSFACTURE'));
    TOBL.PutValue('GL_TIERSPAYEUR', TOBPiece.GetValue('GP_TIERSPAYEUR'));
    TOBL.PutValue('GL_TIERSLIVRE', TOBPiece.GetValue('GP_TIERSLIVRE'));

    if (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X') then
    begin
      TobCata := FindTOBCataRow(TobPiece, TobCatalogu, i + 1);
      if ((TOBCATA <> nil) and (not Reliquat)) then
        TOBCATA.PutValue('UTILISE', 'X');
    end;

    { Nomenclatures }
    IndiceOuv := 0;
    IndiceNOmen := 0;
    // Modif BTP
    if IsOuvrage (TOBL) then
    begin
      IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
      if (IndiceOuv = 0) and (NaturepieceOKPourOuvrage(TOBL)) and (not AutoriseDecoup) then
      begin
        if (V_PGI.SAV) or (V_PGI.Superviseur) then
        begin
          FerrorSD := true;
          TOBL.SetInteger('GL_INDICENOMEN',0)
        end else
        begin
          MessageValid := 'Erreur sur ouvrage '+TOBL.GetValue('GL_CODEARTICLE')+ '(Indice Nul)';
          PgiError(MessageValid);
          //
          V_PGI.Ioerror := OeUnknown;
          exit;
        end;
      end;
    end else if IsNomenclature(TOBL) then
    begin
    	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN')
    end else
    begin
    	TOBL.PutValue('GL_INDICENOMEN',0);
    end;
    // fin Modif BTP
    TobNomen := nil;
    if (IndiceOuv > 0) and (TOBOuvrages <> nil) then
    begin
      if (not FerrorSD) then
      begin
        if (not ReAffecteLigOuv(IndiceOuv, TobL, TOBOuvrages)) then
        begin
          MessageValid := 'Erreur de réaffectation / ouvrage '+ TOBL.GetValue('GL_CODEARTICLE');
          PgiError(MessageValid);
          V_PGI.Ioerror := OeUnknown;
          exit;
        end;
        TOBOuvrage := TOBOuvrages.Detail[IndiceOuv - 1];
      end;
    end else
    begin
      if (IndiceNomen > 0) and (TOBNomenclature <> nil) then
      begin
        TOBNomen := TOBNomenclature.Detail[IndiceNomen - 1];
        ReAffecteLigNomen(IndiceNomen, TOBL, TOBNomenclature);
      end;
    end;
    { Dispo }
    RatioVA := GetRatio(TOBL, nil, trsStock);
    RatioVente := GetRatio(TOBL, nil, trsVente);
    if (not Ajout) or ((Ajout) and (IsNewLigne(TOBL))) then
    begin
      // Transfert Inter Dépôts : MAJ du stock en tenant compte du fait que le dépôt est géré (ou n'est pas géré) sur site.
      if NaturePieceG = 'TRE' then
      begin
            DepotEmetteur := TOBPiece.GetValue('GP_DEPOTDEST');
            TOBPiece.PutValue('GP_DEPOTDEST', '');
          RetourMajStock := MajQteStock(TOBL, TOBArticles, TOBOuvrage, ColPlus, ColMoins, True, RatioVA, DepotEmetteur);
      end
      else if QualifMvt = 'AMV' then
      begin
        // Le stock est déjà décompté sur le document initial
        RetourMajStock := True;
      end else
      begin
        if TOBL.GetValue('GL_TYPEREF') <> 'CAT' then
        begin
          RetourMajStock := MajQteStock(TOBL, TOBArticles, TOBouvrage, ColPlus, ColMoins, True, RatioVA);
          if (RetourMajStock) and (TOBL.GetValue('GL_ENCONTREMARQUE') = 'X') then
            RetourMajStock := MajQteStockContreM(TOBL, TOBCatalogu, ColPlus, ColMoins, True, RatioVA);
        end
        else
          RetourMajStock := MajQteStockContreM(TOBL, TOBCatalogu, ColPlus, ColMoins, True, RatioVA);
      end;
      if not RetourMajStock then
      begin
        if GenereAuto then
        begin
        	// Dans le cas ou le stock n'est pas suffisant
          // On reinit simplement
          if NaturePieceG <> 'LBT' then
          begin
            TOBL.PutValue('GL_QTERELIQUAT',TOBL.GetValue('GL_QTERELIQUAT') + TOBL.GetValue('GL_QTEFACT'));
            TOBL.PutValue('GL_QTEFACT', 0);
            TOBL.PutValue('GL_QTESTOCK',0);
            TOBL.PutValue('GL_QTERESTE',0);
            // --- GUINIER ---
            if ReliquatMt then
            begin
              TOBL.PutValue('GL_MTRELIQUAT',TOBL.GetValue('GL_MTRELIQUAT') + TOBL.GetValue('GL_MONTANTHTDEV'));
              TOBL.PutValue('GL_MTRESTE',0);
            end;
            TOBL.putValue('GL_RECALCULER','X');
            TOBPiece.putValue('GP_RECALCULER','X');
          end;
        end else
        begin
           MessageValid := 'Erreur de réaffectation stock / ouvrage '+TOBL.GetValue('GL_CODEARTICLE');
          PgiError(MessageValid);
           V_PGI.IoError := oeStock;
           Break;
        end;
      end;
      // MODIF BTP
      if TOBL.fieldExists('SOLDERELIQUAT') and (TOBL.GetValue('SOLDERELIQUAT')='X') then
      begin
      	TOBL.PutValue('GL_QTERELIQUAT',0);
        if ReliquatMt then TOBL.PutValue('GL_MTRELIQUAT',0);
      end;
      // --
      if (NaturePieceG<>'INV') then
      	MajPrixValo(TOBL, TOBA, StPrixValo, True, RatioVA, RatioVente);
      if (TOBCATA <> nil) then
        MajPrixValoContreM(TOBL, TOBCata, StPrixValo, True, RatioVA, RatioVente);
      if not Reliquat then
      begin
        if ActF = 'ENR' then
          TOBL.PutValue('GL_VIVANTE', '-');
      end;
      if not ReliquatMt then TOBL.PutValue('GL_MTRESTE',0);
      (*
      //--- GUINIER : Gestion des Reliquats sur Montant pour Prestation
      if not ReliquatMt then
        TOBL.PutValue('GL_MTRESTE',0)
      else
        TOBL.PutValue('GL_MTRESTE', TOBL.GEtVALUE('GL_MONTANTHTDEV'));
      *)
    end;
    {$IFDEF BTP}
    if TOBL.getValue('INDICERG') <> 0 then
    begin
      ReIndiceLigneRg(TOBPieceRG, TOBL.GetValue('GL_NUMORDRE'), TOBL.GetValue('INDICERG'));
      ReIndiceLigneBasesRg(TOBBasesRG, TOBL.GetValue('GL_NUMORDRE'), TOBL.GetValue('INDICERG'));
      TOBL.PutValue('GL_NONIMPRIMABLE', RGNonImprimable);
    end;
    {$ENDIF}

    {$IFDEF GPAOLIGHT} { Init de champs dans LIGNECOMPL}
      if (NaturePieceG = 'CC') and wEstGereALaCommande(TOBL.G('GL_ARTICLE')) then
      begin
        TobL.PutValue('GLC_QACCSAIS', TobL.G('GL_QTEFACT'));
        TobL.PutValue('GLC_DATEACC' , TobL.G('GL_DATELIVRAISON'));
      end;
    {$ENDIF GPAOLIGHT}
    // FQ 11961 --> protection des coefs de marge, fg, fc , FR pour la validation
    ProtectionCoef  (TOBL);
  end;

  {$IFDEF GPAO}
  if Rafale then
  begin
    if (V_PGI.IoError = oeOK) and (GetInfoParpiece(NaturePieceG, 'GPP_PILOTEORDRE') = 'X') then
      PiloteRecepOrdreST(TobPiece);
  end;
  {$ENDIF GPAO}
  if ((not Reliquat) and (not Rafale)) then
  begin
    if TOBArticles <> nil then
    begin
      for i := TOBArticles.Detail.Count - 1 downto 0 do
        if ((TOBArticles.Detail[i].GetValue('UTILISE') <> 'X') and (TOBArticles.Detail[i].GetValue('_FROMOUVRAGE') <> 'X')) or
           (TOBArticles.Detail[i].GetValue('SUPPRIME')='X') then
          TOBArticles.Detail[i].Free;
    end;
    if TOBCatalogu <> nil then
    begin
      for i := TOBCatalogu.Detail.Count - 1 downto 0 do
        if TOBCatalogu.Detail[i].GetValue('UTILISE') <> 'X' then
          TOBCatalogu.Detail[i].Free;
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean Louis Decosse
Créé le ...... : 19/01/2000
Modifié le ... :   /  /
Description .. : Création d'une pièce vide
Mots clefs ... : PIECE;CREATION;
*****************************************************************}

function CreerPieceVide(var CleDoc: R_CleDoc; CodeTiers, CodeAffaire, Etablissement, Domaine: string; EnHT, SaisieContre: Boolean): boolean;
var TOBPiece, TOBTiers,TOBAffaire, TOBAdresses: TOB;
	i,Numf,numl : integer;
  DEV: RDevise;
  {$IFDEF AFFAIRE}
  Aff0, Aff1, Aff2, Aff3, Aff4: string;
  {$ENDIF}
  ValNumP: T_ValideNumPieceVide;
  io: TIoErr;
  TOBADr : TOB;
  QQ : TQuery;
begin
  Result := true;
  // Initialisations
  if (Etablissement = '') then Etablissement := VH^.EtablisDefaut;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, Etablissement, Domaine);
  // CleDoc.NumeroPiece:=GetNumSoucheG(CleDoc.Souche) ;   // IncNumSoucheG(CleDoc.Souche) ;
  if CleDoc.NumeroPiece = 0 then
  begin
    ValNumP := T_ValideNumPieceVide.Create;
    ValNumP.NaturePiece := CleDoc.NaturePiece;
    ValNumP.Souche    := CleDoc.Souche;
    ValNumP.DatePiece := CleDoc.DatePiece;
    io := Transactions(ValNumP.ValideNumPieceVide, 5);
    if io = oeOk then CleDoc.NumeroPiece := ValNumP.NewNum else CleDoc.NumeroPiece := 0;
    ValNumP.Free;
  end;

  if CleDoc.NumeroPiece <= 0 then
  begin
  	result := false;
    PGIBox('Souche associée à la nature de pièce non valide', 'Création de pièce');
    Exit;
  end;

  CleDoc.Indice := 0;
  TOBPiece := TOB.Create('PIECE', nil, -1);
  TOBAffaire := TOB.Create ('AFFAIRE',nil,-1);
  QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeAffaire+'"',true,1,'',true);
  TOBAffaire.SelectDB('',QQ);
  ferme (QQ);
  TOBAdresses := TOB.Create ('les adresses', nil,-1);
  if GetParamSoc('SO_GCPIECEADRESSE') then
  begin
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
    TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
  end else
  begin
    TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
    TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
  end;
  // Modif BTP
  AddLesSupEntete(TOBPiece);
  // ---
  InitTOBPiece(TOBPiece);
  // Divers
  TOBPiece.PutValue('GP_ETABLISSEMENT', Etablissement);
  TOBPiece.PutValue('GP_DEPOT', VH_GC.GCDepotDefaut);
  TOBPiece.PutValue('GP_REGIMETAXE', VH^.RegimeDefaut);
  TOBPiece.PutValue('GP_VENTEACHAT', GetInfoParPiece(CleDoc.NaturePiece, 'GPP_VENTEACHAT'));
  if EnHT then TOBPiece.PutValue('GP_FACTUREHT', 'X')
  else TOBPiece.PutValue('GP_FACTUREHT', '-');
  if SaisieContre then TOBPiece.PutValue('GP_SAISIECONTRE', 'X')
  else TOBPiece.PutValue('GP_SAISIECONTRE', '-');
  // Tiers
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  RemplirTOBTiers(TOBTiers, CodeTiers, CleDoc.NaturePiece, False);
  TOBPiece.PutValue('GP_TIERS', CodeTiers);
  TiersVersPiece(TOBTiers, TOBPiece);
  // Représentant du tiers
  if TobPiece.GetValue('GP_REPRESENTANT') = '' then
  begin
    if TOBTiers.FieldExists('T_REPRESENTANT') then
      TobPiece.PutValue('GP_REPRESENTANT', TOBTiers.GetValue('T_REPRESENTANT'));
  end;

  // Entête
  TOBPiece.PutValue('GP_NATUREPIECEG', CleDoc.NaturePiece);
  TOBPiece.PutValue('GP_SOUCHE', CleDoc.Souche);
  TOBPiece.PutValue('GP_NUMERO', CleDoc.NumeroPiece);
  TOBPiece.PutValue('GP_DATEPIECE', CleDoc.DatePiece);
  TOBPiece.PutValue('GP_AFFAIRE', CodeAffaire);

  if CodeAffaire <> '' then
  begin
	{$IFDEF BTP}
    BTPCodeAffaireDecoupe(CodeAffaire, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
    {$ELSE}
    CodeAffaireDecoupe(CodeAffaire, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
    {$ENDIF}
    TOBPiece.PutValue('GP_AFFAIRE1', Aff1);
    TOBPiece.PutValue('GP_AFFAIRE2', Aff2);
    TOBPiece.PutValue('GP_AFFAIRE3', Aff3);
    TOBPiece.PutValue('GP_AVENANT', Aff4);
    // mcd 19/03/02 dans le cas d'une pièce sur affaire, il faut renseigner le champ TAVENCAISSEMENT
    TobPiece.Putvalue('GP_TVAENCAISSEMENT', PositionneExige(TOBTiers));
    if Cledoc.Naturepiece = 'DAP' then TOBPiece.PutValue('GP_AFFAIREDEVIS', CodeAffaire);
  end;

  // Devise
  DEV.Code := TOBTiers.GetValue('T_DEVISE');
  if (DEV.Code = '') then DEV.Code := V_PGI.DevisePivot;
  GetInfosDevise(DEV);
  DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);
  TOBPiece.PutValue('GP_DEVISE', DEV.Code);
  TOBPiece.PutValue('GP_TAUXDEV', DEV.Taux);
  TOBPiece.PutValue('GP_DATETAUXDEV', CleDoc.DatePiece);
  // Pièce associée à une affaire
  if (GetInfoParPiece(CleDoc.NaturePiece, 'GPP_PREVUAFFAIRE') = 'X') then
  begin
    TOBPiece.PutValue('GP_MAJLIBRETIERS', 'XXXXXXXXXXX');
    if VH_GC.GCIfDefCEGID then
       TobPiece.PutValue('GP_DOMAINE', '001'); // En attente gestion compléte du dommaine sur affaire
  end;

  if (TOBPiece.getValue('GP_NATUREPIECEG')='DAP') Then
  begin
    // préparation des pieces adresses
    TiersVersAdresses(TOBTiers, TOBAdresses, TOBPiece);
    AffaireVersAdresses(TOBAffaire,TOBAdresses,TOBPiece); //mcd 29/09/03 déplacé pour résoudre cas clt fact # ds affaire et adresse fact
    LivAffaireVersAdresses (TOBAffaire,TOBAdresses,TOBPiece);
    // --
    if GetParamSoc('SO_GCPIECEADRESSE') then
    begin
      for i := 0 to TOBAdresses.Detail.Count - 1 do
      begin
        TOBAdr := TOBAdresses.Detail[i];
        TOBAdr.PutValue('GPA_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
        TOBAdr.PutValue('GPA_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
        TOBAdr.PutValue('GPA_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
        TOBAdr.PutValue('GPA_INDICEG', TOBPiece.GetValue('GP_INDICEG'));
        TOBAdr.PutValue('GPA_NUMLIGNE', 0);
        if i = 0 then TOBAdr.PutValue('GPA_TYPEPIECEADR', '001') else TOBAdr.PutValue('GPA_TYPEPIECEADR', '002');
      end;
    end else
    begin
      AddReferenceAdd(TOBPiece, TOBAdresses);
      // Livraison
      NumL := CreerAdresse(TOBAdresses.Detail[0]);
      TOBPiece.PutValue('GP_NUMADRESSELIVR', NumL);
      NumF := CreerAdresse(TOBAdresses.Detail[1]);
      TOBPiece.PutValue('GP_NUMADRESSEFACT', NumF);
    end;
    result := TOBAdresses.InsertDB(nil);
  end;
  // Enregistrement
  if result then Result := TOBPiece.InsertDB(nil);
  TobPiece.free;
  TOBADresses.Free;
  TOBAffaire.free;
  TobTiers.free;
end;

procedure T_ValideNumPieceVide.ValideNumPieceVide;
begin
  NewNum := 0;
  if Souche = '' then Exit;
  NewNum := SetNumberAttribution(NaturePiece, Souche,DatePiece, 1);
end;

function CreerTOBPieceVide(var CleDoc: R_CleDoc; CodeTiers, CodeAffaire, Etablissement, Domaine: string; EnHT, SaisieContre: Boolean; NumPiece: integer = -1):TOB;
var TOBPiece, TOBTiers: TOB;
  DEV: RDevise;
  {$IFDEF AFFAIRE}
  Aff0, Aff1, Aff2, Aff3, Aff4: string;
  {$ENDIF}
begin
  Result := nil;
  // Initialisations
  if (Etablissement = '') then Etablissement := VH^.EtablisDefaut;
  CleDoc.Souche := GetSoucheG(CleDoc.NaturePiece, Etablissement, Domaine);
  if NumPiece <> -1 then CleDoc.NumeroPiece := Numpiece
                    else CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche,CleDoc.DatePiece);
  if CleDoc.NumeroPiece <= 0 then
  begin
    PGIBox('Souche associée à la nature de pièce non valide', 'Création de pièce');
    Exit;
  end;
  //IncNumSoucheG(CleDoc.Souche) ;
  CleDoc.Indice := 0;
  TOBPiece := TOB.Create('PIECE', nil, -1);
  // Modif BTP
  AddLesSupEntete(TOBPiece);
  // ---
  InitTOBPiece(TOBPiece);

  //FV1 : 14/10/2015 - Ajout du paramétrage société Force PA
  TOBPiece.PutValue('GP_PIECEENPA', GetInfoParPiece(CleDoc.NaturePiece, 'GPP_FORCEENPA'));    

  // Divers
  TOBPiece.PutValue('GP_ETABLISSEMENT', Etablissement);
  TOBPiece.PutValue('GP_DEPOT', VH_GC.GCDepotDefaut);
  TOBPiece.PutValue('GP_REGIMETAXE', VH^.RegimeDefaut);
  TOBPiece.PutValue('GP_VENTEACHAT', GetInfoParPiece(CleDoc.NaturePiece, 'GPP_VENTEACHAT'));
  if EnHT then TOBPiece.PutValue('GP_FACTUREHT', 'X')
  else TOBPiece.PutValue('GP_FACTUREHT', '-');
  if SaisieContre then TOBPiece.PutValue('GP_SAISIECONTRE', 'X')
  else TOBPiece.PutValue('GP_SAISIECONTRE', '-');
  // Tiers
  TOBTiers := TOB.Create('TIERS', nil, -1);
  TOBTiers.AddChampSup('RIB', False);
  RemplirTOBTiers(TOBTiers, CodeTiers, CleDoc.NaturePiece, False);
  TOBPiece.PutValue('GP_TIERS', CodeTiers);
  TiersVersPiece(TOBTiers, TOBPiece);
  {$IFDEF GPAO}
  // Représentant du tiers
  if TobPiece.GetValue('GP_REPRESENTANT') = '' then
  begin
    if TOBTiers.FieldExists('T_REPRESENTANT') then
      TobPiece.PutValue('GP_REPRESENTANT', TOBTiers.GetValue('T_REPRESENTANT'));
  end;
  {$ENDIF GPAO}
  // Entête
  TobPiece.Putvalue('GP_TVAENCAISSEMENT', PositionneExige(TOBTiers));

  TOBPiece.PutValue('GP_NATUREPIECEG', CleDoc.NaturePiece);
  TOBPiece.PutValue('GP_SOUCHE', CleDoc.Souche);
  TOBPiece.PutValue('GP_NUMERO', CleDoc.NumeroPiece);
  TOBPiece.PutValue('GP_DATEPIECE', CleDoc.DatePiece);
  TOBPiece.PutValue('GP_AFFAIRE', CodeAffaire);

  if CodeAffaire <> '' then
  begin
    {$IFDEF BTP}
    //if Copy(CodeAffaire,1,1) = 'W' then
    //	 CodeAppelDecoupe(CodeAffaire, Aff0, Aff1, Aff2, Aff3, Aff4)
    //else
    BTPCodeAffaireDecoupe(CodeAffaire, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
    {$ELSE}
    CodeAffaireDecoupe(CodeAffaire, Aff0, Aff1, Aff2, Aff3, Aff4, taCreat, false);
	{$ENDIF}
    TOBPiece.PutValue('GP_AFFAIRE1', Aff1);
    TOBPiece.PutValue('GP_AFFAIRE2', Aff2);
    TOBPiece.PutValue('GP_AFFAIRE3', Aff3);
    TOBPiece.PutValue('GP_AVENANT', Aff4);
    PutValueDetail(TOBPiece,'GP_ETABLISSEMENT', GetChampsAffaire (CodeAffaire,'AFF_ETABLISSEMENT'));
    if TOBPiece.GetValue('GP_ETABLISSEMENT') = '' then PutValueDetail(TOBPiece,'GP_ETABLISSEMENT', VH^.EtablisDefaut);
  end;

  // Devise
  DEV.Code := TOBTiers.GetValue('T_DEVISE');
  if (DEV.Code = '') then DEV.Code := V_PGI.DevisePivot;
  GetInfosDevise(DEV);
  DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);
  TOBPiece.PutValue('GP_DEVISE', DEV.Code);
  TOBPiece.PutValue('GP_TAUXDEV', DEV.Taux);
  TOBPiece.PutValue('GP_COTATION', DEV.Cotation);
  TOBPiece.PutValue('GP_DATETAUXDEV', CleDoc.DatePiece);
  // Pièce associée à une affaire
  if (GetInfoParPiece(CleDoc.NaturePiece, 'GPP_PREVUAFFAIRE') = 'X') then
  begin
    TOBPiece.PutValue('GP_MAJLIBRETIERS', 'XXXXXXXXXXX');
  end;
  // Enregistrement
  TOBTiers.Free;
  Result := TOBPiece;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 23/01/2002
Modifié le ... :   /  /
Description .. : Modification des parmètres d'edition si
Suite ........ : GPP_VALMODELE='X'
Mots clefs ... :
*****************************************************************}

Procedure ModificationModele(CleDoc:R_CleDoc;ImpDoc,OkEtat,Bapercu,OkEtiq,BApercuEtiq: Boolean; Modele,ModeleEtiq:String;NbCopies:Integer;TOBPiece:TOB;TL:TList;TT:TStrings;DGD : boolean=false; ImpressionViaTOB : TImprPieceViaTOB = nil) ;
var Apercu,ApercuEtiq,ModeleEtat,NouvelleValeur,ImpEtiq,IsEtat:String ;
begin
if BApercu then Apercu:='X' else Apercu:='-' ;
if BApercuEtiq then ApercuEtiq:='X' else ApercuEtiq:='-' ;
if OkEtiq then ImpEtiq:='X' else ImpEtiq:='-' ;
ModeleEtat:=Modele ;
NouvelleValeur:=AGLLanceFiche('MBO','VALIDMODELE','','','MODELEETAT='+ModeleEtat+';'
+';MODELEETIQ='+ModeleEtiq+';APERCU='+Apercu+';APERCUETIQ='+ApercuEtiq+';IMPETIQ='+ImpEtiq+';NBCOPIE='+IntToStr(NbCopies)+'') ;
if NouvelleValeur <>'' then
  begin
  IsEtat:=ReadTokenSt(NouvelleValeur) ;
  if IsEtat='X' then OkEtat:=True else OkEtat:=False ;
  Modele:=ReadTokenSt(NouvelleValeur) ;
  Apercu:=ReadTokenSt(NouvelleValeur) ;
  if Apercu='X' then Bapercu:=True else BApercu:=False ;
  NBCopies:=StrToInt(ReadTokenST(NouvelleValeur)) ;
  // Etiquettes
  ModeleEtiq:=ReadTokenSt(NouvelleValeur) ;
  ApercuEtiq:=ReadTokenSt(NouvelleValeur) ;
  if ApercuEtiq='X' then BapercuEtiq:=True else BApercuEtiq:=False ;
  ImpEtiq:=ReadTokenSt(NouvelleValeur) ;
  if ImpEtiq='X' then OkEtiq:=True else OkEtiq:=False ;
  end else ImpDoc:=False ;
ImpressionPieceEtiq(CleDoc,ImpDoc,OkEtat,BApercu,OkEtiq,BApercuEtiq,Modele,ModeleEtiq,NbCopies,TOBPiece,TL,TT,DGD,ImpressionViaTOB) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 19/01/2000
Modifié le ... : 23/01/2002
Description .. : Modif AC: Récupération des paramètres d'impression
Mots clefs ... : SAISIE;PIECE;IMPRIMER;
*****************************************************************}

procedure ImprimerLaPiece(TOBPiece, TOBTiers: TOB; CleDoc: R_CleDoc; BImpClick: Boolean = False; DGD : boolean = false; TOBAffaire : TOB=nil;ImpressionViaTOB : TImprPieceViaTOB = nil);
var TL: TList;
  TT: TStrings;
  TOBT: TOB;
  Modele, NaturePiece, Etab,Domaine, ModeleEtiq: string;
  NbCopies: integer;
  OkEtat, ImpDoc, OkImp, BApercu, OkEtiq, BApercuEtiq: boolean;
  {$IFDEF BTP}
  TypeFac: string;
  {$ENDIF}
begin
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  OkEtat := False;
  Etab := TOBPiece.GetValue('GP_ETABLISSEMENT');
  Domaine := TOBPiece.GetValue('GP_DOMAINE');
  if GetInfoParPiece(NaturePiece, 'GPP_ACTIONFINI') = 'IMP' then PutValueDetail(TOBPiece, 'GP_VIVANTE', '-');
  Modele := '';
  NbCopies := GetInfoParPiece(NaturePiece, 'GPP_NBEXEMPLAIRE');
  if ((NbCopies < 0) or (NbCopies > 99)) then NbCopies := 0;
  BApercu := GetInfoParPiece(NaturePiece, 'GPP_APERCUAVIMP') = 'X';
  V_PGI.NoPrintDialog := not BApercu;

  if ctxMode in V_PGI.PGIContexte then ImpDoc := (GetInfoParPiece(NaturePiece, 'GPP_IMPIMMEDIATE') = 'X') or BImpClick
  else ImpDoc := True;

  {Recherche sur profils affaire}
  {$IFDEF AFFAIRE}
  if (TOBAffaire <> nil) and GetParamsoc('SO_AFPARPIECEAFF') then
  begin
    TOBT := TOBAffaire.FindFirst(['API_NATUREPIECEG'], [CleDoc.NaturePiece], False);
    if TOBT <> nil then
    begin
      Modele := TOBT.GetValue('API_IMPETAT');
      if Modele <> '' then OkEtat := True else Modele := TOBT.GetValue('API_IMPMODELE');
      NbCopies := TOBT.GetValue('API_NBEXEMPLAIRE');
      if ((NbCopies < 0) or (NbCopies > 99)) then NbCopies := 0;
    end;
  end;
  {$ENDIF}

  {Recherche sur profils tiers}
  if Modele = '' then
  begin
    TOBT := TOBTiers.FindFirst(['GTP_NATUREPIECEG'], [CleDoc.NaturePiece], False);
    if TOBT <> nil then
    begin
      Modele := TOBT.GetValue('GTP_IMPETAT');
      if Modele <> '' then OkEtat := True else Modele := TOBT.GetValue('GTP_IMPMODELE');
      NbCopies := TOBT.GetValue('GTP_NBEXEMPLAIRE');
      if ((NbCopies < 0) or (NbCopies > 99)) then NbCopies := 0;
    end;
  end;
  {Recherche sur "les" PARPIECE}
  if Modele = '' then
  begin
    {Etat Nature/Domaine}
    Modele := GetInfoParPieceDomaine(NaturePiece, Domaine, 'GDP_IMPETAT');
    if Modele <> '' then OkEtat := True else
    begin
      {Document Nature/Domaine}
      Modele := GetInfoParPieceDomaine(NaturePiece, Domaine, 'GDP_IMPMODELE');
      if Modele = '' then
      begin
        {Etat Nature/Etablissement}
        Modele := GetInfoParPieceCompl(NaturePiece, Etab, 'GPC_IMPETAT');
        if Modele <> '' then OkEtat := True else
        begin
          {Document Nature/Etablissement}
          Modele := GetInfoParPieceCompl(NaturePiece, Etab, 'GPC_IMPMODELE');
          if Modele = '' then
          begin
            {Etat Nature}
				    if (ImpressionViaTOB <> nil) and (ImpressionViaTOB.Usable) then
            	Modele:=GetInfoParPiece(NaturePiece, 'GPP_IMPMODELE')
            else
            begin
              Modele := GetInfoParPiece(NaturePiece, 'GPP_IMPETAT');
            end;
            OkEtat := True;
          end;
        end;
      end;
    end;
  end;
  // Recherche sur Paramsoc
  if Modele = '' then Modele := VH_GC.GCImpModeleDefaut;
  // --- Modif BTP
  // Recherche spéciale pour le modèle des situations de travaux (BTP)
  {$IFDEF BTP}
  TypeFac := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));

  if (Pos(CleDoc.NaturePiece ,'DAC;FBT;FBP;BAC')>0) and
    ((TypeFac = 'AVA') or (TypeFac = 'DAC')) then
  begin
    // Récupérer le modèle des situations dans les paramètres
    if (ImpressionViaTOB= nil) or (not ImpressionViaTOB.Usable) then
    begin
      Modele := GetParamsoc('SO_BTETATSIT');
    end else
    begin
    	if DGD then
      begin
      	Modele := GetParamsoc('SO_BTETATDGDDIR');
      end else
      begin
      	if Pos(Cledoc.NaturePiece,'DAC;FBT') > 0 then Modele := GetParamsoc('SO_BTETATSITDIR')
                                                 else Modele := GetParamsoc('SO_BTETATFBP');
      end;
    end;
    OkEtat := True;
  end;
  {$ENDIF}
  // ----------------

  // Traitement
  TL := TList.Create;
  TT := TStringList.Create;
  OkImp := True;
  {$IFDEF FOS5}
  OkImp := not (FOImprimeTicket(TOBPiece, CleDoc));
  {$ENDIF}
  // Recherche du modèle d'étiquette
  if ModeleEtiq = '' then
  begin
    {Etat Nature}
    OkEtiq := GetInfoParPiece(NaturePiece, 'GPP_IMPETIQ') = 'X';
    ModeleEtiq := GetInfoParPiece(NaturePiece, 'GPP_ETATETIQ');
    BApercuEtiq := GetInfoParPiece(NaturePiece, 'GPP_APERCUAVETIQ') = 'X';
  end;
  if NbCopies <= 0 then NbCopies := 1;
  if OkImp then
  begin
    if GetInfoParPiece(NaturePiece, 'GPP_VALMODELE') = 'X' then
      ModificationModele(CleDoc, ImpDoc, OkEtat, BApercu, OkEtiq, BApercuEtiq, Modele,ModeleEtiq, NbCopies, TOBPiece, TL, TT,DGD, ImpressionViaTOB)
    else
    begin
    	ImpressionPieceEtiq(CleDoc, ImpDoc, OkEtat, BApercu, OkEtiq, BApercuEtiq, Modele, ModeleEtiq, NbCopies, TOBPiece, TL, TT,DGD, ImpressionViaTOB);
    end;
  end;
  TT.Free;
  TL.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 23/01/2002
Modifié le ... : 23/01/2002
Description .. : Impression d'un document et Etiquettes
Mots clefs ... : PIECE;ETIQUETTE;IMPRIMER;
*****************************************************************}

procedure ImpressionPieceEtiq(CleDoc: R_CleDoc; ImpDoc, OkEtat, Bapercu, OkEtiq, BApercuEtiq: Boolean; Modele, ModeleEtiq: string; NbCopies: Integer; TOBPiece:
  TOB; TL: TList; TT: TStrings; DGD: boolean = false; ImpressionViaTOB : TImprPieceViaTOB = nil);
var TOBB, TOBE, TobEtiq: TOB;
  Souche, Numero, indice, NaturePiece, RegimePrix, CodeEtat, stWhere, LibModele, SQL, SQLOrder: string;
  i_ind1, nbex, i_ind2: Integer;
  CodeArticleGen,TypeFacturation: string;
begin
  (*
  if (Not Assigned(ImpressionViaTOB)) or (ImpressionViaTOB=nil) then
  begin
    pgibox('Impression Directe non autorisée pour ce document', 'Erreur d''impression');
    exit;
  end;
  *)
  // Récupération de la requette
  if (ImpressionViaTOB = nil) or (not ImpressionViaTOB.Usable) then
  begin
    {*
    if OkEtat then
    begin
      SQL := RecupSQLModele('E', 'GPJ', Modele, '', '', '', ' WHERE ' + WherePiece(CleDoc, ttdLigne, False));
      SQL := SQL + ' and GL_NONIMPRIMABLE<>"X"';
      if (pos('ORDER BY', uppercase(SQL)) <= 0) then SQL := SQL + ' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE';
    end else
    begin
      SQL := RecupSQLModele('L', 'GPI', Modele, '', '', '', ' WHERE ' + WherePiece(CleDoc, ttdPiece, False));
      if (pos('ORDER BY', uppercase(SQL)) <= 0) then SQL := SQL + ' ORDER BY GL_NUMLIGNE';
    end;
    *}
    if OkEtat then
    begin
      SQLOrder := '';
      SQL := RecupSQLEtat('E', 'GPJ', Modele, '', '', '', ' WHERE ' + WherePiece(CleDoc, ttdLigne, False), SQLOrder);
      SQL := SQL + ' and GL_NONIMPRIMABLE<>"X"';
      if (SQLOrder = '') then SQL := SQL + ' order by GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_ARTICLE'
      else SQL := SQL + ' ' + SQLOrder;
    end else
    begin
      SQLOrder := '';
      SQL := RecupSQLEtat('L', 'GPI', Modele, '', '', '', ' WHERE ' + WherePiece(CleDoc, ttdPiece, False), SQLOrder);
      if (SQLOrder = '') then SQL := SQL + ' ORDER BY GL_NUMLIGNE'
      else SQL := SQL + ' ' + SQLOrder;
    end;
  end else
  begin
    //
  	ImpressionViaTOB.PreparationEtat;
    TheImpressionViaTOB := ImpressionViaTOB;
  end;
  TT.Add(SQL);
  TL.Add(TT);
  // Edition du document
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  if ImpDoc and (NbCopies <> 0) then
  begin
    V_PGI.DefaultDocCopies := NbCopies;
{$IFDEF BTP}
    if OkEtat then
    begin
      TheTob := TobPiece;
      (* OPTIMIZATION *)
      TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
      ImprimePieceBTP(CleDoc, Bapercu, Modele, Trim(SQL), Typefacturation, DGD, ImpressionViaTOB);
      (* ---- *)
      TheTob := nil;
    end;
{$ELSE}
    if OkEtat then
    begin
      if LanceEtat('E', 'GPJ', Modele, BApercu, False, False, nil, Trim(SQL), '', False) <= 0 then V_PGI.IoError := oeUnknown;
    end
    else
    begin
      if LanceDocument('L', 'GPI', Modele, TL, nil, BApercu, False) <= 0 then V_PGI.IoError := oeUnknown;
    end;
{$ENDIF}
    Souche := TOBPiece.GetValue('GP_SOUCHE');
    Numero := TOBPiece.GetValue('GP_NUMERO');
    Indice := TOBPiece.GetValue('GP_INDICEG');
    TOBB := VH_GC.TOBGCB.FindFirst(['GPB_NATUREPIECEG', 'GPB_SOUCHE', 'GPB_NUMERO', 'GPB_INDICEG'], [NaturePiece, Souche, Numero, Indice], True);
    if TOBB <> nil then TOBB.Parent.Free;
    TOBE := VH_GC.TOBGCE.FindFirst(['GPE_NATUREPIECEG', 'GPE_SOUCHE', 'GPE_NUMERO', 'GPE_INDICEG'], [NaturePiece, Souche, Numero, Indice], True);
    if TOBE <> nil then TOBE.Parent.Free;
  end
  else DeflagEdit(TOBPiece);
  if (ImpressionViaTOB <> nil) and (ImpressionViaTOB.Usable) then
  begin
  	ImpressionViaTOB.ClearInternal;
    TheImpressionViaTOB := nil;
  end;
  //////////////////////////
  // Edition des Etiquettes
  // Si pas de modéle renseigné je n'édite pas
  //if ModeleEtiq='' then OkETiq:=False else OkETiq:=True ;
  //if GetInfoParPiece(NaturePiece,'GPP_IMPETIQ')='X' or
  if OkEtiq and (ModeleEtiq <> '') then
  begin
    // Suppression des enregistrements de cet utilisateur
    ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "' + V_PGI.USer + '"');

    TobEtiq := TOB.Create('GCTMPETQ', nil, -1);
    if ctxMode in V_PGI.PGIContexte then
      RegimePrix := 'TTC'
    else RegimePrix := 'HT';
    // Code de l'état différent suivant le context
    if ctxMode in V_PGI.PGIContexte then
    begin
      CodeEtat := 'GED';
      LibModele := RechDom('TTMODELETIQART', ModeleEtiq, FALSE);
      EditMonarchSiEtat(LibModele);
    end
    else CodeEtat := 'GEA';
    // Si impression avec modele etiquettes factorisées
    if EtatMonarchFactorise(LibModele) then
    begin
      for i_ind1 := 0 to TOBPiece.detail.count - 1 do
      begin
        if ((TOBPiece.Detail[i_ind1].getvalue('GL_ARTICLE') <> '') and (TOBPiece.Detail[i_ind1].getvalue('GL_QTEFACT') <> 0)) then
        begin
          TOBE := TOB.Create('GCTMPETQ', TobEtiq, -1);
          TOBE.PutValue('GZD_UTILISATEUR', V_PGI.USer);
          TOBE.PutValue('GZD_COMPTEUR', i_ind2);
          TOBE.PutValue('GZD_ARTICLE', TOBPiece.Detail[i_ind1].getvalue('GL_ARTICLE'));
          TOBE.PutValue('GZD_CODEARTICLE', TOBPiece.Detail[i_ind1].getvalue('GL_CODEARTICLE'));
          TOBE.PutValue('GZD_DEPOT', TOBPiece.Detail[i_ind1].getvalue('GL_ETABLISSEMENT'));
          TOBE.PutValue('GZD_REGIMEPRIX', RegimePrix);
          TOBE.PutValue('GZD_NUMERO', TOBPiece.Detail[i_ind1].getvalue('GL_NUMERO'));
          TOBE.PutValue('GZD_SOUCHE', TOBPiece.Detail[i_ind1].getvalue('GL_SOUCHE'));
          TOBE.PutValue('GZD_INDICEG', TOBPiece.Detail[i_ind1].getvalue('GL_INDICEG'));
          TOBE.PutValue('GZD_NUMLIGNE', TOBPiece.Detail[i_ind1].getvalue('GL_NUMLIGNE'));
          TOBE.PutValue('GZD_NBETIQDIM', TOBPiece.Detail[i_ind1].getvalue('GL_QTEFACT'));
          // Appel de la fonction qui transforme le codearticle en article générique (avec le X à la fin)
          CodeArticleGen := GCDimToGen(TOBPiece.Detail[i_ind1].getvalue('GL_CODEARTICLE'));
          TOBE.PutValue('GZD_CODEARTICLEGEN', CodeArticleGen);
        end;
      end;
    end
    else
    begin
      // Remplissage de la table temporaire
      for i_ind1 := 0 to TOBPiece.detail.count - 1 do
      begin
        if (TOBPiece.Detail[i_ind1].getvalue('GL_ARTICLE') <> '') then
        begin
          nbex := TOBPiece.Detail[i_ind1].getvalue('GL_QTEFACT');
          for i_ind2 := 0 to nbex - 1 do
          begin
            TOBE := TOB.Create('GCTMPETQ', TobEtiq, -1);
            TOBE.PutValue('GZD_UTILISATEUR', V_PGI.USer);
            TOBE.PutValue('GZD_COMPTEUR', i_ind2);
            TOBE.PutValue('GZD_ARTICLE', TOBPiece.Detail[i_ind1].getvalue('GL_ARTICLE'));
            TOBE.PutValue('GZD_CODEARTICLE', TOBPiece.Detail[i_ind1].getvalue('GL_CODEARTICLE'));
            TOBE.PutValue('GZD_DEPOT', TOBPiece.Detail[i_ind1].getvalue('GL_ETABLISSEMENT'));
            TOBE.PutValue('GZD_REGIMEPRIX', RegimePrix);
            TOBE.PutValue('GZD_NUMERO', TOBPiece.Detail[i_ind1].getvalue('GL_NUMERO'));
            TOBE.PutValue('GZD_SOUCHE', TOBPiece.Detail[i_ind1].getvalue('GL_SOUCHE'));
            TOBE.PutValue('GZD_INDICEG', TOBPiece.Detail[i_ind1].getvalue('GL_INDICEG'));
            TOBE.PutValue('GZD_NUMLIGNE', TOBPiece.Detail[i_ind1].getvalue('GL_NUMLIGNE'));
            // Appel de la fonction qui transforme le codearticle en article générique (avec le X à la fin)
            CodeArticleGen := GCDimToGen(TOBPiece.Detail[i_ind1].getvalue('GL_CODEARTICLE'));
            TOBE.PutValue('GZD_CODEARTICLEGEN', CodeArticleGen);
          end;
        end;
      end;
    end;
    if TobEtiq.Detail.Count > 0 then TobEtiq.InsertDB(nil)
    else OkEtat := False ;
    TobEtiq.Free;
    stWhere := ' AND GZD_UTILISATEUR="' + V_PGI.USer + '"';
    if OkEtat then
    begin
      if LanceEtat('E', CodeEtat, ModeleEtiq, BApercuEtiq, False, False, nil, stWhere, '', False) <= 0 then V_PGI.IoError := oeUnknown;
    end;
    if ctxMode in V_PGI.PGIContexte then EditMonarch('');
    // Suppression des enregistrements
    ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "' + V_PGI.USer + '"');
  end;
end;

procedure DeflagEdit(TOBPiece: TOB);
var CleDoc: R_CleDoc;
begin
  if TOBPiece = nil then Exit;
  if TOBPiece.GetValue('GP_EDITEE') <> 'X' then Exit;
  try
    CleDoc := TOB2CleDoc(TOBPiece);
    ExecuteSQL('UPDATE PIECE SET GP_EDITEE="-" WHERE ' + WherePiece(CleDoc, ttdPiece, False));
  except
  end;
end;

procedure ValideLeTiers(TOBPiece, TOBTiers: TOB);
var SQL: string;
  Nb: integer;
  TOBTT : TOB;
  QQ : TQuery;
begin
  TOBTT := nil;
  if TOBTiers = nil then Exit;
  if TOBPiece = nil then Exit;
  TRY
  //{$IFDEF MODE}
  // Indicateur de mise à jour du tiers selon la nature de pièce
  if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_MAJINFOTIERS') = '-' then Exit;
    if TOBPiece.getString('GP_VENTEACHAT')='ACH' then
    begin
      TOBTT := TOB.Create ('TIERS',nil,-1);
      SQl :='SELECT * FROM TIERS WHERE T_TIERS="'+TOBpiece.GetString('GP_TIERS')+'" AND T_NATUREAUXI="FOU"';
      QQ := OpenSQL(Sql,True,1,'',true);
      if not QQ.eof then
      begin
        TOBTT.SelectDB('',QQ);
      end else
      begin
        Exit;
      end;
      ferme (QQ);
    end else
    begin
      TOBTT := TOBTiers;
    end;
    //{$ENDIF}
  {$IFDEF GRC}
  if ((ctxGRC in V_PGI.PGIContexte) and (V_PGI.IoError = oeOK)) then
      MajProspectClient(TOBPiece, TOBTT);
  {$ENDIF}

  SQL := 'UPDATE TIERS SET T_DATEDERNPIECE="' + UsDateTime(TOBPiece.GetValue('GP_DATEPIECE')) + '", '
      + 'T_NUMDERNPIECE=' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ','
      + 'T_TOTDERNPIECE=' + StrfPoint(TOBPiece.GetValue('GP_TOTALHT')) + ','
    + 'T_DATEPROCLI="' + UsDateTime(TOBTiers.GetValue('T_DATEPROCLI')) + '" '
      + 'WHERE T_TIERS="' + TOBTT.GetValue('T_TIERS') + '" AND ' 
      + 'T_NATUREAUXI="' + TOBTT.GetValue('T_NATUREAUXI') + '" AND '
      + 'T_FERME = "-"';
  TRY
  Nb := ExecuteSQL(SQL);
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur ecriture tiers');
    end;
  END;
  if Nb <> 1 then
  BEGIN
    V_PGI.IoError := oeTiers;
  END;
  FINALLY
    if TOBPiece.getString('GP_VENTEACHAT')='ACH' then
    begin
      if TOBTT <>nil Then TOBTT.free;
    end;
  END;
  // if Not TOBTiers.UpdateDB then V_PGI.IoError:=oeTiers ;
end;

procedure InvalideModifTiersPiece(TOBTiers: TOB);
var i: integer;
begin
  if TOBTiers = nil then Exit;
  for i := 0 to TOBTiers.Detail.Count - 1 do TOBTiers.Detail[i].SetAllModifie(False);
end;

procedure EtudieColsGrid(GS: THGrid);

  function GetPrefixe(Nomcol : string) : string;
  begin
    result := Copy(NomCol,1,Pos('_',NomCol)-1);
  end;

var NomCol, LesCols: string;
  icol, ichamp, iTableLigne: integer;
  prefixe : string;
  Mcd : IMCDServiceCOM;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  LesCols := GS.Titres[0];
  icol := 0;
  repeat
    NomCol := uppercase(Trim(ReadTokenSt(LesCols)));
    if NomCol <> '' then
    begin
      iTableLigne := PrefixeToNum(GetPrefixe(Nomcol));
      ichamp := ChampToNum(NomCol);
      if ichamp >= 0 then
      begin
        if (iTableLigne>0) and (iChamp>0) and (Pos('X', (Mcd.GetField(NomCol) as IFieldCOM ).Control ) > 0) then
        begin
(*
          // "X" interdit la saisie sauf si aussi "Y" pour FrontOffice
          if ((Pos('Y', (Mcd.GetField(NomCol) as IFieldCOM ).Control) > 0) and ((ctxFO in V_PGI.PGIContexte) or (ctxMode in V_PGI.PGIContexte)))
            then else GS.ColLengths[icol] := -1;
*)
        end;

        if NomCol = 'GL_NUMLIGNE'       then SG_NL          := icol else
        if NomCol = 'GL_ENCONTREMARQUE' then SG_ContreM     := icol else
        if NomCol = 'GL_REFARTSAISIE'   then SG_RefArt      := icol else
        if NomCol = 'GL_REFCATALOGUE'   then SG_CATALOGUE   := icol else
        if NomCol = 'GL_REFARTTIERS'    then SG_REFTiers    := icol else
        if NomCol = 'GLC_GETCEDETAIL'   then SG_DETAILBORD  := icol else
        if NomCol = 'GL_TPSUNITAIRE'    then SG_TEMPS       := icol else
        if NomCol = 'GL_HEURE'          then SG_TEMPS       := icol else
        if NomCol = 'GL_TOTALHEURE'     then SG_TEMPSTOT    := icol else
        if NomCol = 'GL_QUALIFHEURE'    then SG_QUALIFTEMPS := icol else
        if NomCol = 'GL_LIBELLE'        then SG_Lib         := icol else
        if NomCol = 'GL_QTEFACT'        then SG_QF          := icol else
        if NomCol = 'GL_QTESTOCK'       then SG_QS          := icol else
        if NomCol = 'GL_PUHTDEV'        then SG_Px          := icol else
        if NomCol = 'GL_PUTTCDEV'       then SG_Px          := icol else
        if NomCol = 'GL_REMISELIGNE'    then SG_Rem         := icol else
        if NomCol = 'GL_REPRESENTANT'   then SG_Rep         := icol else
        if NomCol = 'GL_DEPOT'          then SG_Dep         := icol else
        if NomCol = 'GL_VALEURREMDEV'   then SG_RV          := icol else
        if NomCol = 'GL_TOTREMLIGNEDEV' then SG_RL          := icol else
        if NomCol = 'GL_MONTANTHTDEV'   then SG_Montant     := icol else //AC
        if NomCol = 'GL_MONTANTTTCDEV'  then SG_Montant     := icol else //AC
        if NomCol = 'GL_DATELIVRAISON'  then SG_DateLiv     := icol else //AC
        if NomCol = 'GL_TOTALHTDEV'     then SG_Total       := icol else //AC
        if NomCol = 'GL_TOTALTTCDEV'    then SG_Total       := icol else //AC
        if NomCol = 'GL_QTERELIQUAT'    then SG_QR          := icol else //AC
        if NomCol = 'GLC_CIRCUIT'       then SG_CIRCUIT     := icol else
        if NomCol = 'GL_MOTIFMVT'       then SG_Motif       := icol else
        if NomCol = 'GL_AFFAIRE'        then SG_Aff         := icol else
        if NomCol = 'GL_QTERESTE'       then SG_QReste      := icol else { NEWPIECE }
        if NomCol = 'GL_MTRESTE'        then SG_MTReste     := icol else { GUINIER }
        // Modif BTP
        if NomCol = 'GL_QTEPREVAVANC'   then SG_QA          := icol else //MODIFBTP 190301
        if NomCol = 'GL_POURCENTAVANC'  then SG_Pct         := icol else //MODIFBTP 190301
        if NomCol = 'GL_QUALIFQTEVTE'   then SG_Unit        := icol else //MODIFBTP 250401
        if NomCol = 'GL_QUALIFQTEACH'   then SG_Unit        := icol else //MODIFBTP 250401
        if NomCol = 'GL_QUALIFQTESTO'   then SG_Unit        := icol else //MODIFBTP 250401
        if NomCol = 'GL_TYPEARTICLE'    then SG_TypA        := icol else
        if NomCol = 'GL_TOTPREVAVANC'   then SG_montant     := icol else //MODIFBTP 250401
        if NomCol = 'GL_DPA'            then SG_pxAch       := icol else
        if NomCol = 'GL_MONTANTPA'      then SG_MontantAch  := icol;
        // ----------
        // DEBUT Modif MODE 31/07/2002
        if NomCol = 'GL_TYPEREMISE'     then SG_TypeRem     := icol else
        if NomCol = 'GL_PUHTNETDEV'     then SG_PxNet       := icol else
        if NomCol = 'GL_PUTTCNETDEV'    then SG_PxNet       := icol else
        // FIN Modif MODE 31/07/2002
        if NomCol = 'GL_FAMILLETAXE1'   then SG_CODTVA1     := icol else
        if NomCol = 'BLF_QTEDEJAFACT'   then SG_DEJAFACT    := icol else
        if NomCol = 'BLF_QTEMARCHE'     then SG_QTEPREVUE   := icol else
        if NomCol = 'BLF_QTESITUATION'  then SG_QTESITUATION:= icol else
        if NomCol = 'BLF_POURCENTAVANC' then SG_POURCENTAVANC := icol else
        if NomCol = 'BLF_MTSITUATION'   then SG_MontantSit  := icol else
        if NomCol = 'BLF_MTMARCHE'      then SG_MontantMarche := icol else
        if NomCol = 'BLF_MTDEJAFACT'    then SG_MTDEJAFACT  := icol else
        if NomCol = 'GLC_NATURETRAVAIL' then SG_NATURETRAVAIL := icol else
        if NomCol = 'GL_FOURNISSEUR'    then SG_FOURNISSEUR := icol else
        if NomCol = 'GLC_VOIRDETAIL'    then SG_VOIRDETAIL  := icol else
        if NomCol = 'BLF_MTPRODUCTION'  then SG_MTPRODUCTION:= icol else
        if NomCol = 'GLC_CODEMATERIEL'  then SG_MATERIEL    := icol else
        if NomCol = 'GLC_BTETAT'        then SG_TYPEACTION := icol else
        if NomCol = 'GLC_QTEAPLANIF'    then SG_QTEAPLANIF := icol else
        if NomCol = 'GL_COEFMARG'       then SG_COEFMARG := icol else
        if NomCol = 'GLC_NUMEROTATION'  then SG_NUMAUTO := icol else
        if NomCol = 'GL_QTESAIS'        then SG_QTESAIS := icol else
        if NomCol = 'GL_RENDEMENT'      then SG_RENDEMENT := icol else
        if NomCol = 'GL_PERTE'          then SG_PERTE := icol else
        if NomCol = 'GL_QUALIFHEURE'    then SG_QUALIFHEURE := icol else
        if NomCol = 'GLC_QTECOND'       then SG_QTECOND := icol else
        if NomCol = 'GLC_COEFCOND'       then SG_COEFCOND := icol else
        if NomCol = 'GL_CODECOND'       then SG_CODECOND := icol else
        if NomCol = 'GL_CODEMARCHE'    then SG_CODEMARCHE := icol else
        ;
      end else
      begin
        if NomCol = 'POURCENTMARG'      then SG_POURMARG := icol;
        if NomCol = 'POURCENTMARQ'      then SG_POURMARQ := icol;
        if NomCol = 'TOTALTS'           then SG_TOTALTS := icol;
      end;
    end;
    Inc(icol);
  until ((LesCols = '') or (NomCol = ''));
end;

function ChoixCommercial(TypeCom, ZoneCom: string): string;
var Q: TQuery;
		st : string;
begin
	st := 'SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE ';
  if TypeCom <> '' then st := st + 'GCL_TYPECOMMERCIAL="' + TypeCom + '" AND ';
  st := st + 'GCL_ZONECOM="' + ZoneCom + '"';
  Result := '';
//  Q := OpenSQL('SELECT GCL_COMMERCIAL FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="' + TypeCom + '" AND GCL_ZONECOM="' + ZoneCom + '"', True);
  Q := OpenSQL(st, True,-1, '', True);
  if not Q.EOF then Result := Q.Fields[0].AsString;
  Ferme(Q);
end;

function AdapteCond(TOBL, TOBConds: TOB; var Qte: Double): boolean;
var 
  QteCond,OneQTe : Double;
begin
  Result := true;
  if TOBL.GetDouble('GLC_COEFCOND') = 0 Then Exit;
  QteCond := round(Qte / TOBL.GetDouble('GLC_COEFCOND'));
  if ARRONDI(QteCond*TOBL.GetDouble('GLC_COEFCOND'),0) < Qte then QteCond := QteCond +1;
  TOBL.SetDouble('GLC_QTECOND',QteCond);
  Qte := ARRONDI(TOBL.GetDouble('GLC_QTECOND')*TOBL.GetDouble('GLC_COEFCOND'),V_PGI.okdecQ);
end;

procedure DecodeRefPiece(St: string; var CleDoc: R_CleDoc); { NEWPIECE }
var
  StC, StL: string;
begin
  if St = '' then  exit;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StC := St;
  CleDoc.DatePiece := EvalueDate(ReadTokenSt(StC));
  CleDoc.NaturePiece := ReadTokenSt(StC);
  CleDoc.Souche := ReadTokenSt(StC);
  CleDoc.NumeroPiece := StrToInt(ReadTokenSt(StC));
  CleDoc.Indice := StrToInt(ReadTokenSt(StC));
  StL := ReadTokenSt(StC);
  if StL <> '' then
  begin
    CleDoc.NumLigne := StrToInt(StL);
    CleDoc.NumOrdre := StrToInt(StL); { NEWPIECE }
  end;
end;

function EncodeRefPiece (Cledoc : r_cledoc) : string;
var std : string;
begin
  StD := FormatDateTime('ddmmyyyy', cledoc.DatePiece);
  Result := StD + ';' + cledoc.NaturePiece  + ';' + Cledoc.Souche  + ';' +
  					IntToStr(cledoc.NumeroPiece) + ';' + IntToStr(cledoc.NumOrdre) + ';;';
end;

function EncodeRefPiece(TOBPL: TOB; NewNum: Integer = 0; WithNumLigne : boolean=true): string; { NEWPIECE }
var
  DD: TDateTime;
  StD: string;
begin
  Result := '';
  if TOBPL = nil then Exit;
  if TOBPL.NomTable = 'PIECE' then
  begin
    DD := TOBPL.GetValue('GP_DATEPIECE');
    StD := FormatDateTime('ddmmyyyy', DD);
    if NewNum = 0 then
      Result := StD + ';' + TOBPL.GetValue('GP_NATUREPIECEG') + ';' + TOBPL.GetValue('GP_SOUCHE') + ';'
        + IntToStr(TOBPL.GetValue('GP_NUMERO')) + ';' + IntToStr(TOBPL.GetValue('GP_INDICEG')) + ';;'
    else
      Result := StD + ';' + TOBPL.GetValue('GP_NATUREPIECEG') + ';' + TOBPL.GetValue('GP_SOUCHE') + ';'
        + IntToStr(NewNum) + ';' + IntToStr(TOBPL.GetValue('GP_INDICEG')) + ';;';
  end
  else if TOBPL.NomTable = 'LIGNE' then
  begin
    DD := TOBPL.GetValue('GL_DATEPIECE');
    StD := FormatDateTime('ddmmyyyy', DD);
    Result := StD + ';' + TOBPL.GetValue('GL_NATUREPIECEG') + ';' + TOBPL.GetValue('GL_SOUCHE') + ';'
      + IntToStr(TOBPL.GetValue('GL_NUMERO')) + ';' + IntToStr(TOBPL.GetValue('GL_INDICEG')) + ';';
      //               +IntToStr(TOBPL.GetValue('GL_NUMLIGNE'))+';' ;
//    + IntToStr(TOBPL.GetValue('GL_NUMORDRE')) + ';'; { NEWPIECE }
    if WithNumLigne then
    begin
      result := result + IntToStr(TOBPL.GetValue('GL_NUMORDRE')) + ';'; { NEWPIECE }
    end else
    begin
      result := result + ';';
    end;
  end else if TOBPL.NomTable = 'PIECERG' then
  begin
    DD := TOBPL.GetValue('PRG_DATEPIECE');
    StD := FormatDateTime('ddmmyyyy', DD);
    Result := StD + ';' + TOBPL.GetValue('PRG_NATUREPIECEG') + ';' + TOBPL.GetValue('PRG_SOUCHE') + ';'
      + IntToStr(TOBPL.GetValue('PRG_NUMERO')) + ';' + IntToStr(TOBPL.GetValue('PRG_INDICEG')) + ';';
    if WithNumLigne then
    begin
      result := result + IntToStr(TOBPL.GetValue('PRG_NUMLIGNE')) + ';'; { NEWPIECE }
    end else
    begin
      result := result + ';';
    end;
  end else if TOBPL.NomTable = 'ACOMPTES' then
  begin
  	if TOBPL.GetValue('GAC_NUMERO')  > 0 then
    begin
    Result :=  '01011900;' + TOBPL.GetValue('GAC_NATUREPIECEG') + ';' + TOBPL.GetValue('GAC_SOUCHE') + ';'
      + IntToStr(TOBPL.GetValue('GAC_NUMERO')) + ';' + IntToStr(TOBPL.GetValue('GAC_INDICEG')) + ';';
    end;
  end;
end;


function EncodeRefPieceSup(TOBPL: TOB): string; { NEWPIECE }
begin
  Result := '';
  if TOBPL = nil then Exit;
  if TOBPL.NomTable = 'PIECE' then
  begin
    Result := '%;' + TOBPL.GetValue('GP_NATUREPIECEG') + ';' + TOBPL.GetValue('GP_SOUCHE') + ';'
      + IntToStr(TOBPL.GetValue('GP_NUMERO')) + ';' + IntToStr(TOBPL.GetValue('GP_INDICEG')) + ';%'
  end
  else if TOBPL.NomTable = 'LIGNE' then
  begin
    Result := '%;' + TOBPL.GetValue('GL_NATUREPIECEG') + ';' + TOBPL.GetValue('GL_SOUCHE') + ';'
      + IntToStr(TOBPL.GetValue('GL_NUMERO')) + ';' + IntToStr(TOBPL.GetValue('GL_INDICEG')) + ';%';
  end;
end;

procedure MontreNumero(TOBPiece: TOB);
var St: string;
begin
  if not GetParamSocSecur('SO_GCMONTRENUMERO', False) then Exit;
  St := RechDom('GCNATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'), False)
    + ' N° ' + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' du ' + DateToStr(TOBPiece.GetValue('GP_DATEPIECE'));
  HShowMessage('0;Dernière pièce créée;' + St + ';E;O;O;O;', '', '');
end;

procedure AffecteComm(TOBC, TOBL: TOB; Totale: boolean);
var VC: Boolean;
  LaComm: double;
  LeTyp: string;
begin
  if TOBC.NomTable = 'COMMISSION' then
  begin
    LeTyp := TOBC.GetValue('GCM_TYPECOM');
    LaComm := TOBC.GetValue('GCM_COMMISSION');
    VC := (TOBC.GetValue('GCM_NATUREPIECEG') = TOBL.GetValue('GL_NATUREPIECEG'));
  end else
  begin
    LeTyp := TOBC.GetValue('GCL_TYPECOM');
    LaComm := TOBC.GetValue('GCL_COMMISSION');
    if TOBC.GetValue('GCL_NATUREPIECEG') = 'ZZ1' then
{$IFDEF BTP}
      VC := ((TOBL.GetValue('GL_NATUREPIECEG') = 'FAC') or (TOBL.GetValue('GL_NATUREPIECEG') = 'AVS')
        or (TOBL.GetValue('GL_NATUREPIECEG') = 'AVC') or (TOBL.GetValue('GL_NATUREPIECEG') = 'FBT')
        or (TOBL.GetValue('GL_NATUREPIECEG') = 'FBP') or (TOBL.GetValue('GL_NATUREPIECEG') = 'BAC')
        or (TOBL.GetValue('GL_NATUREPIECEG') = 'ABT') or (TOBL.GetValue('GL_NATUREPIECEG') = 'ABP'))
{$ELSE}
      VC := ((TOBL.GetValue('GL_NATUREPIECEG') = 'FAC') or (TOBL.GetValue('GL_NATUREPIECEG') = 'AVS')
        or (TOBL.GetValue('GL_NATUREPIECEG') = 'AVC'))
{$ENDIF}
    else VC := (TOBC.GetValue('GCL_NATUREPIECEG') = TOBL.GetValue('GL_NATUREPIECEG'));
  end;
  if Totale then
  begin
    TOBL.PutValue('GL_TYPECOM', LeTyp);
    TOBL.PutValue('GL_COMMISSIONR', LaComm);
  end;
  if VC then TOBL.PutValue('GL_VALIDECOM', 'VAL') else TOBL.PutValue('GL_VALIDECOM', 'AFF');
end;

procedure CommVersLigne(TOBPiece, TOBArticles, TOBComms: TOB; ARow: integer; Totale: boolean);
var TOBL, TOBA, TOBC, TOBEx: TOB;
  Repres, RefU, FN1, FN2, FN3: string;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  if ((Totale) and (TOBL.GetValue('GL_VALIDECOM') <> 'NON')) then Exit;
  TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  if TOBA = nil then
  begin
    InitialiseComm(TOBL, True);
    Exit;
  end;
  if TOBA.GetValue('GA_COMMISSIONNABLE') <> 'X' then
  begin
    InitialiseComm(TOBL, True);
    Exit;
  end;
  Repres := TOBL.GetValue('GL_REPRESENTANT');
  if Repres = '' then
  begin
    InitialiseComm(TOBL, True);
    Exit;
  end;
  RefU := TOBL.GetValue('GL_ARTICLE');
  if RefU = '' then
  begin
    InitialiseComm(TOBL, True);
    Exit;
  end;
  TOBC := TOBComms.FindFirst(['GCL_COMMERCIAL'], [Repres], False);
  if TOBC = nil then
  begin
    InitialiseComm(TOBL, True);
    Exit;
  end;
  if TOBC.Detail.Count > 0 then
  begin
    TOBEx := TOBC.FindFirst(['GCM_ARTICLE'], [RefU], False);
    if TOBEx <> nil then AffecteComm(TOBEx, TOBL, Totale) else
    begin
      FN1 := TOBA.GetValue('GA_FAMILLENIV1');
      FN2 := TOBA.GetValue('GA_FAMILLENIV2');
      FN3 := TOBA.GetValue('GA_FAMILLENIV3');
      if ((FN1 <> '') and (FN2 <> '') and (FN3 <> '')) then
        TOBEx := TOBC.FindFirst(['GCM_FAMILLENIV1', 'GCM_FAMILLENIV2', 'GCM_FAMILLENIV3'], [FN1, FN2, FN3], False)
      else TOBEx := nil;
      if TOBEx <> nil then AffecteComm(TOBEx, TOBL, Totale) else
      begin
        if ((FN1 <> '') and (FN2 <> '')) then TOBEx := TOBC.FindFirst(['GCM_FAMILLENIV1', 'GCM_FAMILLENIV2'], [FN1, FN2], False)
        else TOBEx := nil;
        if TOBEx <> nil then AffecteComm(TOBEx, TOBL, Totale) else
        begin
          if FN1 <> '' then TOBEx := TOBC.FindFirst(['GCM_FAMILLENIV1'], [FN1], False) else TOBEx := nil;
          if TOBEx <> nil then AffecteComm(TOBEx, TOBL, Totale) else AffecteComm(TOBC, TOBL, Totale);
        end;
      end;
    end;
  end else
  begin
    AffecteComm(TOBC, TOBL, Totale);
  end;
end;

Function OKRepresentant (CodeCom,TypeCom : string) : boolean;
var st : string;
begin
	result := true;
  st := 'SELECT * FROM COMMERCIAL WHERE ';
  if TypeCom <> '' then st := st + 'GCL_TYPECOMMERCIAL="' + TypeCom + '" AND ';
  st := st + 'GCL_COMMERCIAL="' + CodeCom + '"';
  if not existeSQL(st) then result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : F.Vautrain
Créé le ...... : 03/03/2014
Modifié le ... :   /  /
Description .. : Recherche du nom du commercial
Mots clefs ... : COMMERCIAL
*****************************************************************}
procedure LibelleCommercial(CodeRep: string; var lib1, lib2: string; IsRessource : Boolean=False);
var qq  : TQuery;
begin

  if IsRessource then
    lib1  := 'SELECT ARS_LIBELLE,ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_RESSOURCE="' + CodeRep + '"'
  Else
    lib1  := 'SELECT GCL_LIBELLE,GCL_PRENOM FROM COMMERCIAL WHERE GCL_COMMERCIAL="' + CodeRep + '"';

  QQ := OpenSQL(lib1, false);
  if not QQ.EOF then
  begin
    if IsRessource then
    begin
    lib1 := QQ.findField('ARS_LIBELLE').asString;
    lib2 := QQ.findField('ARS_LIBELLE2').asString;
    end Else
    begin
      lib1 := QQ.findField('GCL_LIBELLE').asString;
      lib2 := QQ.findField('GCL_PRENOM').asString;
    end;
  end
  else
  begin
    lib1 := '';
    lib2 := '';
  end;
  ferme(QQ);
end;

procedure AjouteRepres(CodeComm, TypeCom: string; TOBComms: TOB);
var TOBC: TOB;
  QQ: TQuery;
{$IFDEF BTP}
	st : string;
{$ENDIF}
begin
  if CodeComm = '' then Exit;
  TOBC := TOBComms.FindFirst(['GCL_COMMERCIAL'], [CodeComm], False);
  if TOBC <> nil then Exit;
{$IFNDEF BTP}
  if TypeCom = '' then if ctxFO in V_PGI.PGIContexte then TypeCom := 'VEN' else TypeCom := 'REP';
{$ENDIF}
	st := 'SELECT * FROM COMMERCIAL WHERE ';
  if TypeCom <> '' then st := st + 'GCL_TYPECOMMERCIAL="' + TypeCom + '" AND ';
  st := st + 'GCL_COMMERCIAL="' + CodeComm + '"';
//  QQ := OpenSQL('SELECT * FROM COMMERCIAL WHERE GCL_TYPECOMMERCIAL="' + TypeCom + '" AND GCL_COMMERCIAL="' + CodeComm + '"', True);
  QQ := OpenSQL(st, True,-1, '', True);
  if not QQ.EOF then TOBC := TOB.CreateDB('COMMERCIAL', TOBComms, -1, QQ);
  Ferme(QQ);
  if TOBC <> nil then
  begin
    QQ := OpenSQL('SELECT * FROM COMMISSION WHERE GCM_COMMERCIAL="' + CodeComm + '"', True,-1, '', True);
    TOBC.LoadDetailDB('COMMISSION', '', '', QQ, False, True);
    Ferme(QQ);
  end;
end;

function ExistePiece(CD: R_CleDoc): Boolean;
var QQ: TQuery;
begin
	(*
  QQ := OpenSQL('SELECT * FROM PIECE WHERE ' + WherePiece(CD, ttdPiece, False), True,-1, '', True);
  Result := not QQ.EOF;
  Ferme(QQ);
  *)
  result := ExisteSql ('SELECT 1 FROM PIECE WHERE '+WherePiece(CD, ttdPiece, False));
end;

procedure ShowPiecePasHisto(CD: R_CleDoc; Prec: Boolean);
var St, Titre: string;
begin
  St := RechDom('GCNATUREPIECEG', CD.NaturePiece, False) + ' N° ' + IntToStr(CD.NumeroPiece);
  if CD.Indice <> 0 then St := St + ' ' + IntToStr(CD.Indice);
  St := St + ' du ' + DateToStr(CD.DatePiece);
  if Prec then Titre := 'La pièce liée n''est plus historisée'
  else Titre := 'La pièce d''origine n''est plus historisée';
  HShowMessage('0;' + Titre + ';' + St + ';E;O;O;O;', '', '');
end;

procedure PieceToEcheGC(TOBPiece, TOBH: TOB);
begin
  TOBH.PutValue('GPE_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
  TOBH.PutValue('GPE_DATEPIECE', TOBPiece.GetValue('GP_DATEPIECE'));
  TOBH.PutValue('GPE_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
  TOBH.PutValue('GPE_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
  TOBH.PutValue('GPE_INDICEG', TOBPiece.GetValue('GP_INDICEG'));
  TOBH.PutValue('GPE_DEVISE', TOBPiece.GetValue('GP_DEVISE'));
  TOBH.PutValue('GPE_DEVISEESP', TOBPiece.GetValue('GP_DEVISE'));
  TOBH.PutValue('GPE_TAUXDEV', TOBPiece.GetValue('GP_TAUXDEV'));
  TOBH.PutValue('GPE_COTATION', TOBPiece.GetValue('GP_COTATION'));
  TOBH.PutValue('GPE_DATETAUXDEV', TOBPiece.GetValue('GP_DATETAUXDEV'));
  TOBH.PutValue('GPE_SAISIECONTRE', TOBPiece.GetValue('GP_SAISIECONTRE'));
  TOBH.PutValue('GPE_TIERS', TOBPiece.GetValue('GP_TIERS'));
  TOBH.PutValue('GPE_CAISSE', TOBPiece.GetValue('GP_CAISSE'));
  TOBH.PutValue('GPE_NUMZCAISSE', TOBPiece.GetValue('GP_NUMZCAISSE'));
end;

procedure ModeToEchesGC(Mode: T_ModeRegl; TOBPiece, TOBEches,TOBpieceTrait,TOBAcomptes : TOB; NbLigAcpte : integer);
var TOBH: TOB;
  i: integer;
  MtDed, MtMaj : double;
begin
  TOBEches.ClearDetail;
  //JT - eQualité n° 10657
  //Test s'il faut créer l'échéance (la ligne peut être =0 et on a pas saisi
  //de règlement/acomptes, de remise ou d'escompte)
  i := 0;
  if TOBPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then
//    MtMaj := Mode.TabEche[i].MontantD
    MtMaj := TOBPiece.GetValue('GP_TOTALTTCDEV')
    else
//    MtMaj := Mode.TabEche[i].MontantP;
    MtMaj := TOBPiece.GetValue('GP_TOTALTTC');
  MtDed := MtMaj + TOBPiece.GetValue('GP_ESCOMPTE') + TOBPiece.GetValue('GP_REMISEPIED')
        + TOBPiece.Somme('GL_REMISELIGNE',['!GL_SOUCHE'],[''],True);
  if (NbLigAcpte = 0) and (MtDed = 0) then exit;
  //Fin JT - eQualité n° 10657
  for i := 1 to Mode.NbEche do
  begin
    TOBH := TOB.Create('PIEDECHE', TOBEches, -1);
    TOBH.AddChampSupValeur('LIBELLE','');
    PieceToEcheGC(TOBPiece, TOBH);
    TOBH.PutValue('GPE_NUMECHE', i);
    TOBH.PutValue('GPE_ACOMPTE', '-');
    TOBH.PutValue('GPE_MODEPAIE', Mode.TabEche[i].ModePaie);
    TOBH.PutValue('GPE_DATEECHE', Mode.TabEche[i].DateEche);
    TOBH.PutValue('GPE_MONTANTECHE', Mode.TabEche[i].MontantP);
    {$IFDEF CHR}
    TOBH.PutValue('GPE_NOFOLIO', Mode.TabEche[i].Folio);
    {$ENDIF}
    if TOBPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then
    begin
      TOBH.PutValue('GPE_MONTANTDEV', Mode.TabEche[i].MontantD);
    end else
    begin
// MontantE dans TabEche n'existe plus
//      if TOBPiece.GetValue('GP_SAISIECONTRE') = 'X' then TOBH.PutValue('GPE_MONTANTDEV', Mode.TabEche[i].MontantE)
//      else TOBH.PutValue('GPE_MONTANTDEV', Mode.TabEche[i].MontantP);
      TOBH.PutValue('GPE_MONTANTDEV', Mode.TabEche[i].MontantP);
    end;
    TOBH.PutValue('GPE_FOURNISSEUR', Mode.tabEche[i].Pourqui);
    TOBH.SetString('LIBELLE',Mode.tabEche[i].Pourquilib);
  end;
end;


function InitModeGC(TOBPiece, TOBTiers, TOBAcomptes, TOBPieceRG, TOBpieceTrait,TOBPorcs : TOB; DEV: RDEVISE): T_ModeRegl;
var Mode: T_ModeRegl;
  XP, XD : Double;
  DD: TDateTime;
  // Modif BTP
  RGP, RGD, RP,RD : double;
  MtAprendreEncompte : double;
  NumCaution : string;
  // ---
begin
  Rp := 0; RD := 0;
  GetSommeAcomptes(TOBAcomptes, XP, XD,'',true);
  // Modif BTP
  GetRg(TOBPieceRG,False,True,RGP,RGD,numcaution);
  // -- GESTION DES AVANCES ET RETENUES SUR COLECTIFS TIERS-
  GetRetenuesCollectif (TOBPorcs ,RP,RD,'TTC');
  //
  MtAprendreEncompte := Arrondi(TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - RGD - RD,DEV.decimale);
  FillChar(Mode, Sizeof(Mode), #0);
  Mode.Action := taModif;
  Mode.ModeInitial := TOBPiece.GetValue('GP_MODEREGLE');
  Mode.ModeFinal := TOBPiece.GetValue('GP_MODEREGLE');
  Mode.CodeDevise := TOBPiece.GetValue('GP_DEVISE');
  Mode.TauxDevise := TOBPiece.GetValue('GP_TAUXDEV');
  Mode.Quotite := DEV.Quotite;
  Mode.Decimale := DEV.Decimale;
  Mode.Symbole := DEV.Symbole;
  Mode.DateFact := TOBPiece.GetValue('GP_DATEPIECE');
  Mode.DateFactExt := Mode.DateFact;
  DD := TOBPiece.GetValue('GP_DATELIVRAISON');
  if DD <= 10 then DD := Mode.DateFact;
  Mode.DateBL := DD;
  Mode.Aux := TOBTiers.GetValue('T_AUXILIAIRE');
  Mode.JourPaiement1 := TOBTiers.GetValue('T_JOURPAIEMENT1');
  Mode.JourPaiement2 := TOBTiers.GetValue('T_JOURPAIEMENT2');
//  Mode.ModeOppose := (TOBPiece.GetValue('GP_SAISIECONTRE') = 'X');
  // Modif BTP
  (*
  Mode.TotalAPayerP := TOBPiece.GetValue('GP_TOTALTTC') - XP - RGP;
  if DEV.Code <> V_PGI.DevisePivot then Mode.TotalAPayerD := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD - RGD
  else Mode.TotalAPayerD := Mode.TotalAPayerP;
  *)
  Mode.TotalAPayerD := MtAprendreEncompte;
  Mode.TotalAPayerP := DEVISETOPIVOTEx (MtAprendreEncompte,DEV.Taux,DEV.Quotite,V_PGI.okdecV);
  Result := Mode;
end;
{$IFDEF CHR}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Anne FOURCADE
Créé le ...... : 01/02/2002
Modifié le ... :   /  /
Description .. : Initialisation du tableau de saisie des reglmt.
Mots clefs ... : INITIALISATION;REGLEMENT
*****************************************************************}

function InitModeChr(TOBPiece, TOBTiers, TOBEches: TOB; DEV: RDEVISE): T_ModeRegl;
var Mode: T_ModeRegl;
  XP, XD : Double;
  DD: TDateTime;
  i, folio: integer;
  TOBL: TOB;
begin
  GetSommeReglements(TOBEches, XP, XD);
  FillChar(Mode, Sizeof(Mode), #0);
  Mode.Action := taModif;
  Mode.ModeInitial := TOBPiece.GetValue('GP_MODEREGLE');
  Mode.ModeFinal := TOBPiece.GetValue('GP_MODEREGLE');
  Mode.CodeDevise := TOBPiece.GetValue('GP_DEVISE');
  Mode.TauxDevise := TOBPiece.GetValue('GP_TAUXDEV');
  Mode.Quotite := DEV.Quotite;
  Mode.Decimale := DEV.Decimale;
  Mode.Symbole := DEV.Symbole;
  Mode.DateFact := TOBPiece.GetValue('GP_DATEPIECE');
  Mode.DateFactExt := Mode.DateFact;
  DD := TOBPiece.GetValue('GP_DATELIVRAISON');
  if DD <= 10 then DD := Mode.DateFact;
  Mode.DateBL := DD;
  Mode.Aux := TOBTiers.GetValue('T_AUXILIAIRE');
  Mode.TotalAPayerP := TOBPiece.GetValue('GP_TOTALTTC') - XP;
  if DEV.Code <> V_PGI.DevisePivot then Mode.TotalAPayerD := TOBPiece.GetValue('GP_TOTALTTCDEV') - XD
  else Mode.TotalAPayerD := Mode.TotalAPayerP;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, i + 1);
    Folio := TOBL.GetValue('GL_NOFOLIO');
    // Ne totaliser que les articles
    if (Folio <> 0) AND (TOBL.GetValue('GL_TYPELIGNE')='ART') then
    begin
      Mode.TabFolioP[Folio] := Mode.TabFolioP[Folio] + TOBL.GetValue('GL_MONTANTTTC');
      if DEV.Code <> V_PGI.DevisePivot then Mode.TabFolioD[Folio] := Mode.TabFolioD[Folio] + TOBL.GetValue('GL_MONTANTTTCDEV')
      else Mode.TabFolioD[Folio] := Mode.TabFolioP[Folio];
    end;
  end;
  if TOBPiece.GetValue('GP_REMISEPIED') <> 0 then
  begin
    Mode.TabFolioP[1] := Mode.TabFolioP[1] - TOBPiece.GetValue('GP_TOTREMISETTC');
    Mode.TabFolioD[1] := Mode.TabFolioD[1] - TOBPiece.GetValue('GP_TOTREMISETTCDEV');
  end;
  Result := Mode;
end;
{$ENDIF}

function FabriqueEchesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBpieceTrait,TOBPorcs : TOB; DEV: RDEVISE): T_ModeRegl;
var Mode: T_ModeRegl;
    nbacpt : Integer;
begin
  Mode := InitModeGC(TOBPiece, TOBTiers, TOBAcomptes, TOBPieceRG,TOBpieceTrait,TOBPorcs, DEV);
  CalculModeRegle(Mode, True); // pour l'entreprise
  if (TOBPieceTrait <> nil) and (TOBpieceTrait.detail.count > 0) and (ExisteReglDirect(TOBpieceTrait)) then
  begin
		// pour les sous traitants en paiement direct
		AjoutePaiementDirectMode (Mode,TOBPiece,TOBTiers,TOBPieceTrait,TOBPieceRG,TOBAcomptes,DEV);
    TOBEches.clearDetail;
  end;
  if (TOBAcomptes = NIL)  then nbacpt := 0 else nbacpt := TOBAcomptes.Detail.count;
  ModeToEchesGC(Mode, TOBPiece, TOBEches,TOBpieceTrait, TOBAcomptes,nbacpt);
  Result := Mode;
end;

procedure ReinitMode(Mode :T_ModeRegl; Niveau : integer);
var Indice : Integer;
begin
	for Indice := Niveau to MaxEche do
  begin
    Mode.TabEche[indice].ModePaie := '';
    Mode.TabEche[indice].pourQui := '';
    Mode.TabEche[indice].pourQuiLib := '';
    Mode.TabEche[indice].MontantD := 0;
    Mode.TabEche[indice].MontantP := 0;
  end;
end;

procedure EchesGCToMode(var Mode: T_ModeRegl; TOBPiece, TOBEches: TOB);
var i: integer;
  TOBH: TOB;
  Fournisseur : string;
begin
  Mode.NbEche := TOBEches.Detail.Count;
  FillChar(Mode.TabEche, Sizeof(Mode.TabEche), #0);
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBH := TOBEches.Detail[i];
    Mode.TabEche[i + 1].ModePaie := TOBH.GetValue('GPE_MODEPAIE');
    Mode.TabEche[i + 1].pourQui  := TOBH.GetValue('GPE_FOURNISSEUR');
    if Mode.TabEche[i + 1].pourQui = '' then
    begin
			Mode.TabEche[i + 1].pourQuiLib := GetParamSocSecur('SO_LIBELLE','Notre Société');
    end else
    begin
    	Mode.TabEche[i + 1].pourQuiLib := TOBH.GetValue('LIBELLE');
    end;
    Mode.TabEche[i + 1].DateEche := TOBH.GetValue('GPE_DATEECHE');
    Mode.TabEche[i + 1].MontantP := TOBH.GetValue('GPE_MONTANTECHE');
    // DEBUT AJOUT CHR
    {$IFDEF CHR}
    Mode.TabEche[i + 1].Folio := TOBH.GetValue('GPE_NOFOLIO');
    {$ENDIF}
    // FIN AJOUT CHR
    if TOBPiece.GetValue('GP_DEVISE') <> V_PGI.DevisePivot then
    begin
      Mode.TabEche[i + 1].MontantD := TOBH.GetValue('GPE_MONTANTDEV');
    end else
    begin
      Mode.TabEche[i + 1].MontantD := Mode.TabEche[i + 1].MontantP;
    end;
  end;
end;

procedure ProrateEchesGC(TOBPiece, TOBEches, TOBAcomptes, TOBPieceRG, TOBPieceTrait, TOBPorcs: TOB; DEV: RDEVISE);
var TotD, TotP, Coef : Double;
  XD, XP, AccP, AccD : Double;
  i: integer;
  DatePiece: TDateTime;
  TOBH: TOB;
  // Modif BTP
  RGP, RGD, RP,RD: double;
  numcaution : string;
  // --
begin
  if TOBEches.Detail.Count <= 0 then Exit;
  TotD := 0;
  TotP := 0;
  GetSommeAcomptes(TOBAcomptes, AccP, AccD);
  // Modif BTP
  GetRg(TOBPieceRG,False,True,RGP,RGD,numcaution);
  // ---
  GetRetenuesCollectif(TOBPorcs, RP,RD,'TTC'); 
  DatePiece := TOBPiece.GetValue('GP_DATEPIECE');
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBH := TOBEches.Detail[i];
    TotD := TotD + TOBH.GetValue('GPE_MONTANTDEV');
    TotP := TotP + TOBH.GetValue('GPE_MONTANTECHE');
    TOBH.PutValue('GPE_DATEPIECE', DatePiece);
  end;
  if TotD = 0 then Exit;
  // Modif BTP
  Coef := (TOBPiece.GetValue('GP_TOTALTTCDEV') - AccD - RGD - RD) / TotD;
  if Arrondi(Coef - 1, 9) = 0 then Exit;
  TotD := 0;
  TotP := 0;
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBH := TOBEches.Detail[i];
    if i < TOBEches.Detail.Count - 1 then
    begin
      XD := TOBH.GetValue('GPE_MONTANTDEV');
      XD := Arrondi(XD * Coef, DEV.Decimale);
      TOBH.PutValue('GPE_MONTANTDEV', XD);
      XP := TOBH.GetValue('GPE_MONTANTECHE');
      XP := Arrondi(XP * Coef, V_PGI.OkDecV);
      TOBH.PutValue('GPE_MONTANTECHE', XP);
      TotD := TotD + XD;
      TotP := TotP + XP;
    end else
    begin
      XD := Arrondi(TOBPiece.GetValue('GP_TOTALTTCDEV') - AccD - TotD - RGD - RD, DEV.Decimale);
      TOBH.PutValue('GPE_MONTANTDEV', XD);
      XP := Arrondi(TOBPiece.GetValue('GP_TOTALTTC') - AccP - TotP - RGP - RP, V_PGI.OkDecV);
      TOBH.PutValue('GPE_MONTANTECHE', XP);
    end;
  end;
end;

function RenseigneEchesGC (TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPIECERG, TOBpieceTrait,TOBPorcs: TOB;
                           DEV: RDEVISE; Action : TActionFiche): T_ModeRegl;
var Mode: T_ModeRegl;
begin
 	Mode := InitModeGC(TOBPiece, TOBTiers, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs, DEV);
  if Action <> taConsult then
  begin
  	ProrateEchesGC(TOBPiece, TOBEches, TOBAcomptes, TOBPieceRG, TOBPieceTrait,TOBPorcs, DEV);
  end;
 	EchesGCToMode(Mode, TOBPiece, TOBEches);
  Result := Mode;
end;

function GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs: TOB; Action: TActionFiche; DEV: RDevise; Ouv: boolean): Boolean;
var Mode: T_ModeRegl;
  ModeR: string;
  Okok: Boolean;
begin
  Result := True;
  AcomptesVersPiece(TOBAcomptes, TOBPiece);
  if (TOBpieceTrait <> nil) and (TOBpieceTrait.detail.count > 0) and (ExisteReglDirect(TOBpieceTrait)) then
  begin
		Mode := FabriqueEchesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBpieceTrait,TOBPorcs,  DEV)
  end else
  begin
    if TOBEches.Detail.Count <= 0 then Mode := FabriqueEchesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBpieceTrait,TOBPorcs,  DEV)
                                	else Mode := RenseigneEchesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBpieceTrait,TOBPorcs, DEV, Action);
  end;
  if Ouv then
  begin
    if Action = taConsult then Mode.Action := taConsult;
    if Mode.TotalAPayerD = 0 then Okok := True else Okok := SaisirEcheance(Mode);
    if Okok then ModeToEchesGC(Mode, TOBPiece, TOBEches,TOBpieceTrait,TOBAcomptes, TOBAcomptes.Detail.count) else Result := False;
  end;
  if Result then
  begin
    ModeR := Mode.ModeFinal;
    if ModeR = '' then ModeR := Mode.ModeInitial;
    if ModeR <> '' then TOBPiece.PutValue('GP_MODEREGLE', ModeR);
  end;
end;

{$IFDEF CHR}

function GereReglementsChr(TOBPiece, TOBTiers, TOBEches, TOBAcomptes: TOB; Action: TActionFiche; DEV: RDevise; Ouv: boolean; Famreg: string; TypeImpFac:
  integer;ModificationTicket:boolean = False): Boolean;
var Mode: T_ModeRegl;
  ModeR, CodeCli: string;
  Okok: Boolean;
  Q: TQuery;
  i: integer;
  TOBL: TOB;
begin
  Okok := False;
  Result := True;
  EcheVersPiece(TOBEches, TOBPiece);
  if TOBEches.Detail.Count <= 0 then
  begin
    Mode := InitModeChr(TOBPiece, TOBTiers, TOBEches, DEV);
    Mode.Action := taCreat;
  end
  else Mode := RenseigneEchesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, nil {AnNE A VOIR TOBPieceRG}, DEV, Action);
  if Ouv then
  begin
    if Action = taConsult then Mode.Action := taConsult;
    CodeCli := TOBPiece.GetValue('GP_TIERS');
    Q := OpenSQL('SELECT T_TIERS, T_MODEREGLE FROM TIERS Where T_TIERS = "' + CodeCli + '"', TRUE,-1, '', True);
    Mode.ModeInitial := Q.FindField('T_MODEREGLE').AsString;
    Mode.ModeFinal := Q.FindField('T_MODEREGLE').AsString;
    if (Mode.Action = taCreat) then
    begin
      if (Mode.TabFolioP[2] = 0) and (Mode.TabFolioP[3] = 0)
        and (Mode.TabFolioP[4] = 0) and (Mode.TabFolioP[5] = 0) then
      begin
        Mode.TabEche[1].MontantP := Mode.TotalAPayerP;
        Mode.TabEche[1].MontantD := Mode.TotalAPayerD;
      end else
      begin
        Mode.TabEche[1].MontantP := Mode.TabFolioP[1];
        Mode.TabEche[1].MontantD := Mode.TabFolioD[1];
      end;
      Mode.TabEche[1].ModePaie := Q.FindField('T_MODEREGLE').AsString;
      Mode.TabEche[1].DateEche := Mode.DateFact;
    end;
    Ferme(Q);
    Okok := SaisirReglement(Mode, Famreg, TypeImpFac, ModificationTicket);
    if (okok) then
    begin
      if (Mode.TabFolioP[1] = Mode.TotalAPayerP) and
        (Mode.TabFolioP[2] = 0) and (Mode.TabFolioP[3] = 0) and
        (Mode.TabFolioP[4] = 0) and (Mode.TabFolioP[5] = 0) then
      begin
        for i := 0 to TOBPiece.Detail.Count - 1 do
        begin
          TOBL := TOBPiece.Detail[i];
          TOBL.PutValue('GL_NOFOLIO', '1');
        end;
      end;
      {Modif GF 141003 Qualite N°10159}
      //ModeToEchesGC(Mode, TOBPiece, TOBEches, TOBAcomptes.Detail.count);
      ModeToEchesGC(Mode, TOBPiece, TOBEches,1);
      EcheVersPiece(TOBEches, TOBPiece);
    end
    else Result := False;
  end;
  if Result then
  begin
    ModeR := Mode.ModeFinal;
    if ModeR = '' then ModeR := Mode.ModeInitial;
    if ModeR <> '' then TOBPiece.PutValue('GP_MODEREGLE', ModeR);
  end;
end;
{$ENDIF}

procedure ChangeTauxAcomptes(TOBPiece, TOBAcomptes: TOB; DEV: RDEVISE);
var TauxPiece: Double;
  i: integer;
  TOBH: TOB;
  XP,XD: Double;
begin
  TauxPiece := DEV.Taux;
  for i := 0 to TOBAcomptes.Detail.Count - 1 do
  begin
    TOBH := TOBAcomptes.Detail[i];
    XD := TOBH.GetValue('GAC_MONTANTDEV');
    XP := DeviseToEuro(XD, TauxPiece, DEV.Quotite);
    TOBH.PutValue('GAC_MONTANT', XP);
  end;
  AcomptesVersPiece(TOBAcomptes, TOBPiece);
end;

procedure ChangeTauxEches(TOBPiece, TOBEches: TOB; DEV: RDEVISE);
var APayerP : double;
  i: integer;
  TOBH: TOB;
  XP, XD, TotP: double;
begin
  {Echéances}
  APayerP := Arrondi(TOBPiece.GetValue('GP_TOTALTTC') - TOBPiece.GetValue('GP_ACOMPTE'), V_PGI.OkDecV);
  TotP := 0;
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBH := TOBEches.Detail[i];
    XD := TOBH.GetValue('GPE_MONTANTDEV');
    if i < TOBEches.Detail.Count - 1 then
    begin
      XP := DeviseToEuro(XD, DEV.Taux, DEV.Quotite);
      TotP := Arrondi(TotP + XP, V_PGI.OkDecV);
    end else
    begin
      XP := Arrondi(APayerP - TotP, V_PGI.OkDecV);
    end;
    TOBH.PutValue('GPE_MONTANTECHE', XP);
  end;
end;

function PieceEncoreVivante(TOBPiece: TOB): Boolean;
var StSQL: string;
begin
  Result := True;
  if TobPiece = nil then exit;
  if TOBPiece.Detail.Count = 0 then exit;
  StSQL := 'Select GP_NATUREPIECEG FROM PIECE WHERE GP_NATUREPIECEG="';
  StSQL := StSQL + TOBPiece.GetValue('GP_NATUREPIECEG') + '" AND GP_SOUCHE="';
  StSQL := StSQL + TOBPiece.GetValue('GP_SOUCHE') + '" AND GP_NUMERO=';
  StSQL := StSQL + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ' AND GP_INDICEG=';
  StSQL := StSQL + IntToStr(TOBPiece.GetValue('GP_INDICEG')) + ' AND GP_VIVANTE="X"';
  if not ExisteSql(StSQL) then Result := False;
end;

function GCPieceCorrecte(TOBPiece, TOBArticles, TOBCatalogu: TOB;Ecran : Tform;ModeBordereau : boolean = false; SaisieCM : Boolean=false): integer;

  function ChampsCMOk (TOBl : TOB) : boolean;
  begin
    Result := True;
    //
    if TobL.GetValue('GL_ARTICLE') = '' then exit;
    if TobL.GetString('GLC_CODEMATERIEL')='' then
    begin
      Result := False;
      exit;
    end;
    if TobL.GetString('GLC_BTETAT')='' then
    begin
      Result := False;
      exit;
    end;
    (*
    if TobL.GetDouble('GLC_QTEAPLANIF')= 0  then
    begin
      Result := False;
      exit;
    end;
    *)
  end;


const erPieceVide = 1;
  erPasArticle = 2;
  erPasMvt = 3;
  erDateEuro = 4;
  erAcompteSup = 5;
  erDevise = 6;
  erCircuit = 9;
  erPxZero = 10;
  erPxokCM = 11;

var i, iArt, iMt, iCirc: integer;
		iartcli : integer;
  TOBL: TOB;
  OkArt     : Boolean;
  OkMtant   : boolean;
  OkDate    : Boolean;
  OkNatAffaire  : Boolean;
  OkCircuit : boolean;
  MtAcc     : Double;
  CodeD     : String;
  StDev     : String;
  NaturePieceG  : string;
  //FV1 - 29/10/2012 : prioritaire pouchain contrôle sur ligne à zéro
//  Rapport   : TGestionRapport;
  OKPxZero  : Boolean;
  PxAchat   : Double;
  PxRevient : Double;
  PxVente   : Double;
  MsgErreur : String;
  XX : TFFacture;
  okrapport : Boolean;
begin
  okrapport := (Ecran.Name <> 'GCMOUVSTKEX');
  XX := TFfacture(ecran);
  NaturePieceG := TobPiece.GetValue('GP_NATUREPIECEG');

  {$IFDEF BTP}
  OkNatAffaire := ((NaturePieceG = VH_GC.AFNatAffaire) or (NaturePieceG = VH_GC.AFNatProposition) or (NaturePieceG = 'AFF'));
  {$ELSE}
  OkNatAffaire := ((NaturePieceG = VH_GC.AFNatAffaire) or (NaturePieceG = VH_GC.AFNatProposition));
  {$ENDIF}

  {$IFDEF CHR}
  Result := 0;
  {$ELSE}
  if (ctxAffaire in V_PGI.PGIContexte) and (OkNatAffaire) then
  begin
    Result := 0;
  end else
  begin
    if TOBPiece.Detail.Count <= 0 then
    begin
      Result := erPieceVide;
      Exit;
    end;
    if (TOBArticles.Detail.Count <= 0) and ((TOBCatalogu = nil) or (TOBCatalogu.Detail.Count <= 0)) and (not ModeBordereau) then
    begin
      Result := erPasArticle;
      Exit;
    end;
  end;
  {$ENDIF}

  if TOBPiece.GetValue('GP_DATEPIECE') < V_PGI.DateDebutEuro then
  begin
    Result := erDateEuro;
    Exit;
  end;

  CodeD := TOBPiece.GetValue('GP_DEVISE');
  if CodeD = '' then StDev := '' else StDev := RechDom('TTDEVISETOUTES', CodeD, False);

  if ((StDev = '') or (StDev = 'Error')) then
  begin
    Result := erDevise;
    Exit;
  end;

  MtAcc := TOBPiece.GetValue('GP_ACOMPTEDEV');

  // DEBUT MODIF CHR
  {$IFNDEF CHR}
  if ((MtAcc > 0) and (Arrondi(TOBPiece.GetValue('GP_TOTALTTCDEV') - MtAcc, 6) < 0)) then
  begin
    Result := erAcompteSup;
    Exit;
  end;
  {$ENDIF}
  //FIN MODIF CHR

  OkArt := False;
  OkMtant := False;
  OkDate := False;
  OkCircuit := true;
  OKPxZero := false;
  //
  if okrapport then
  begin
    XX.RapportPxZero.InitRapport;
    XX.RapportPxZero.Visible := false;
    XX.RapportPxZero.PosLeft := Round(ecran.Width / 1.5);
    XX.RapportPxZero.PosTop  := Round(ecran.Top);
  end;
  //
  //création de l'objet rapport
  if TOBPiece.GetValue('GP_VENTEACHAT') = 'VEN' then
  begin
    OKPxZero := GetParamSocSecur('SO_PXZERODOCVTE', false);
  end;

  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i = 0 then
    begin
      if TOBL.GetValue('GL_TYPEREF') = 'CAT' then iArt := TOBL.GetNumChamp('GL_REFCATALOGUE')
      else iArt := TOBL.GetNumChamp('GL_ARTICLE');
      iArtCli := TOBL.GetNumChamp('GL_REFARTTIERS');
      iMt := TOBL.GetNumChamp('GL_MONTANTHTDEV');
    end;
    //
    if (SaisieCM) and (not ChampsCMOk(TOBL)) then
    BEGIN
        Result := erPxokCM;
        Exit;
    end;
    //
    //Vérification si Prix d'achat, revient et vente à zéro et chargement
    //dans rapport pour affichage
    if (okrapport) and (OkPxZero) and (TOBL.GetValue('GL_TYPELIGNE')='ART') then
    begin
      MsgErreur := '';
      PxAchat   := TOBL.GetDouble('GL_DPA');
      PxRevient := TOBL.GetDouble('GL_DPR');
      PxVente   := TOBL.GetDouble('GL_PUHTDEV');
      if (PxAchat = 0) and (PxRevient=0) and (PxVente=0) then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : aucun prix de renseigné'
      Else if (PxAchat = 0) and (PxRevient=0) then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : Prix d''Achat et Prix de Revient à Zéro'
      Else if (PxAchat = 0) and (PxVente=0) then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : Prix d''Achat et Prix de Vente à Zéro'
      Else if (PxVente = 0) and (PxRevient=0) then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : Prix de Vente et Prix de Revient à Zéro'
      Else if PxAchat = 0 then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : Prix d''Achat à Zéro'
      else if PxRevient = 0 then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : Prix de revient à Zéro'
      else if PxVente = 0 then
        MsgErreur := 'Ligne :' + IntToStr(TOBL.GetValue('GL_NUMLIGNE')) + ' : Prix de Vente à Zéro';
      //Gestion de l'entête de rapport
      if MsgErreur <> '' then XX.RapportPxZero.SauveLigMemo(MsgErreur);
    end;

    if TOBL.GetValeur(iArt) <> '' then OkArt := True;
    if (ModeBordereau) and (TOBL.GetValeur(iart)='') and (TOBL.getValeur(iArtCli)<>'') then okArt := true;
    if TOBL.GetValeur(iMt) <> 0 then OkMtant := True;
    iCirc := TOBL.GetNumChamp('GLC_CIRCUIT');
    if SG_CIRCUIT <> -1 then
    begin
      if (TOBL.GetValeur(iCirc) <> '') and (TOBL.GetValeur(iCirc) <> #0) and (TOBL.GetValeur(iCirc) <> Null) and (OkCircuit) then
        OkCircuit := true
      else
        OkCircuit := false;
    end
    else
      OkCircuit := true;
  end;

  if not OkArt then
  begin
    {$IFDEF CHR}
    Result := 0;
    {$ELSE}
    if (ctxAffaire in V_PGI.PGIContexte) and (OkNatAffaire) then
    begin
      Result := 0
    end else
    begin
      if GetInfoParPiece(NaturePieceG, 'GPP_GESTIONGRATUIT') <> 'X'
        then
        begin
        Result := erPasMvt;
        Exit;
      end;
    end;
    {$ENDIF}
  end;

  if not OkCircuit then
  begin
    Result := erCircuit;
    Exit;
  end;

  if not VH_GC.GCIfDefCEGID then
    if not OkMtant and not ((NaturePieceG = 'EEX') or (NaturePieceG = 'SEX')) then
    begin
      if (ctxAffaire in V_PGI.PGIContexte) and (OkNatAffaire) then
      begin
        Result := 0;
      end else
      begin
{$IFNDEF BTP}
        if GetInfoParPiece(NaturePieceG, 'GPP_GESTIONGRATUIT') <> 'X'
          then
          begin
          Result := erPasMvt;
          Exit;
        end;
{$ENDIF}
      end;
    end;
  Result := 0;

  //Fin de traitement
  if (okrapport) and (OkPxZero) then
  begin
    if XX.RapportPxZero.Memo.Lines.Count > 0 then
    begin
      XX.RapportPxZero.AfficheRapport;    //affichage du rapport d'intégration
      if PGIAsk('Attention des prix sont à zéro, voulez-vous valider la saisie ?','') <> mrYes then
      Begin
         Result := erPxZero;
      end
      else
      begin
         XX.RapportPxZero.Visible := false;
      end;
    end;
  end;

end;

function PositionneExige(TOBTiers: TOB): string;
var NatAux, ExiTva: string;
begin
  Result := 'TD';
  if not VH^.OuiTvaEnc then Exit;
  NatAux := TOBTiers.GetValue('T_NATUREAUXI');
  if ((NatAux = 'CLI') or (NatAux = 'AUD')) then
    ExiTva := VH^.TvaEncSociete
    else
    ExiTva := TOBTiers.GetValue('T_TVAENCAISSEMENT');
  if ExiTva = '' then
    ExiTva := 'TD';
  Result := ExiTva;
end;

function GCValideReglementOblig(TOBPiece, TobTiers, TOBPIECERG: tob; DEV: RDEVISE): Integer;
var MtNet: Double;
  St: string;
  // Modif BTP
  RGP, RGE, RGD: double;
  Numcaution : string;
begin
  Result := 2;
  // Modif BTP
//  GetCumulRG(TOBPIeceRG, RGP, RGD);
  GetRg(TOBPieceRG,False,True,RGP,RGD,numcaution);
  // ---
  MtNet := Arrondi(TOBPiece.GetValue('GP_TOTALTTCDEV') - TOBPiece.GetValue('GP_ACOMPTEDEV') - RGD, DEV.Decimale);
  if (MtNet <> 0) then
  begin
    St := 'Il y a une différence de ' + strf00(MtNet, DEV.Decimale) + ' ' + DEV.symbole + ' entre le montant de la pièce et les acomptes/réglements affectés ';
    TheTob := TobTiers;
    result := strToInt(AglLanceFiche('GC', 'GCVALIDPIECE', '', '', St));
  end;
end;

// Modif BTP

function RecupDebutPourcent(TOBpiece: TOB; Indice, niveau: Integer): Integer;
var
  TOBL: TOB;
  Ind, ITYPL, ITypA, Iimbric: Integer;
begin
  result := 0;
  if Indice <= 0 then exit;
  ITYPL := TOBPiece.detail[Indice].GetNumChamp('GL_TYPELIGNE');
  ITypA := TOBPiece.detail[Indice].GetNumChamp('GL_TYPEARTICLE');
  Iimbric := TOBPiece.detail[Indice].GetNumChamp('GL_NIVEAUIMBRIC');
  for ind := Indice downto 0 do
  begin
    TOBL := TOBPiece.detail[Ind];
    // si niveau > 0 et DP+niveau trouve ou si niveau=0 et POU ou EPO trouve
    // VARIANTE
    (* if (niveau > 0) and (TOBL.GetValeur(ITYPL) = 'DP' + inttostr(niveau)) then *)
    if (niveau > 0) and (IsDebutParagraphe(TOBL,niveau)) then
    begin
      // On prend en compte jusqu'au pourcentage précédent
      result := Ind + 1;
      break;
    end;
    if ((niveau = 0) and (((TOBL.GetValeur(ITypA) = 'POU') and (TOBL.GetValeur(Iimbric) = 0)) or (TOBL.GetValeur(ITYPA) = 'EPO'))) then
    begin
      // On prend en compte jusqu'au pourcentage précédent
      result := Ind + 1;
      break;
    end;
  end;
end;

function RecupDebutParagraph(TOBpiece: TOB; Indice, niveau: Integer; ParagInclus: boolean = false): Integer;
var
  TOBL: TOB;
  Ind, ITYPL, ITypA, Iimbric: Integer;
begin
  result := 0;
  if Indice < 0 then exit;
  ITYPL := TOBPiece.detail[Indice].GetNumChamp('GL_TYPELIGNE');
  ITypA := TOBPiece.detail[Indice].GetNumChamp('GL_TYPEARTICLE');
  Iimbric := TOBPiece.detail[Indice].GetNumChamp('GL_NIVEAUIMBRIC');
  for ind := Indice downto 0 do
  begin
    TOBL := TOBPiece.detail[Ind];
    // MODIF VARIANTE
    if IsDebutParagraphe (TOBL,niveau) then
    (* if (TOBL.GetValeur(ITYPL) = 'DP' + inttostr(niveau)) then *)
    begin
      if ParagInclus then result := ind else result := Ind + 1;
      break;
    end;
  end;
end;

function RecupFinParagraph(TOBpiece: TOB; Indice, niveau: Integer; ParagInclus: boolean = false): Integer;
var
  TOBL: TOB;
  Ind, ITYPL, ITypA, Iimbric: Integer;
begin
  result := 0;
  if Indice < 0 then exit;
  ITYPL := TOBPiece.detail[Indice].GetNumChamp('GL_TYPELIGNE');
  ITypA := TOBPiece.detail[Indice].GetNumChamp('GL_TYPEARTICLE');
  Iimbric := TOBPiece.detail[Indice].GetNumChamp('GL_NIVEAUIMBRIC');
  for ind := Indice to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPiece.detail[Ind];
    // MODIF VARIANTE
    if IsFinParagraphe(TOBL,niveau) then
    (* if (TOBL.GetValeur(ITYPL) = 'TP' + inttostr(niveau)) then *)
    begin
      if ParagInclus then result := ind else result := Ind - 1;
      break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 04/06/2003
Description .. : Recherche d'un R_CleDoc dans un ensemble de TOB
*****************************************************************}

function FindCleDocInTob(CleDoc: R_CleDoc; TheTob: Tob; MultiNiveau: Boolean = True): Tob;
begin
  Result := nil;
  if Assigned(TheTob) then
  begin
    Result := TheTob.FindFirst(['GL_NATUREPIECEG',  'GL_SOUCHE'
      , 'GL_NUMERO', 'GL_INDICEG', 'GL_NUMORDRE']
        , [CleDoc.NaturePiece, CleDoc.Souche
      , CleDoc.NumeroPiece, CleDoc.Indice, CleDoc.NumOrdre]
        , MultiNiveau);
  end;
end;

function FindCleDocInTobPiece(CleDoc: R_CleDoc; TobPiece: Tob): Tob;
begin
  Result := FindCleDocInTob(CleDoc, TobPiece, True);
end;

function FindTobLigInOldTobPiece(TobL, TobPiece_O: Tob): Tob;
var
  CleDoc: R_CleDoc;
begin
  if Assigned(TobL) and Assigned(TobPiece_O) then
  begin
    CleDoc := TOB2CleDoc(TobL);
    CleDoc.NumOrdre := TobL.GetValue('GL_NUMORDRE');
    Result := FindCleDocInTobPiece(CleDoc, TobPiece_O);
  end
  else
    Result := nil;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 29/09/2003
Description .. : Retrouve une TobLig dans une TobPiece à partir de sa clé
*****************************************************************}
function FindOldTobLigInTobPiece(TobL_O, TobPiece: Tob): Tob;
var
  PiecePrec: String;
  i: Integer;
begin
  Result := nil;
  if Assigned(TobL_O) and Assigned(TobPiece) then
  begin
    PiecePrec := EncodeRefPiece(Tobl_O);
    i := -1;
    while (i < TobPiece.Detail.Count - 1) and (Result = nil) do
    begin
      Inc(i);
      if TobPiece.Detail[i].G('GL_PIECEPRECEDENTE') = PiecePrec then
        Result := TobPiece.Detail[i];
    end;
  end;
end;

procedure CompareTobPiece(TobPiece, TobPiece_O, TobDifference: Tob; OnlyWithPiecePrecedente: Boolean);
var
  i: Integer;
  TobLigneNew, TobLigneOld: Tob;
  TobModifFille: Tob;
  CleDoc: R_CLeDoc;
  OkCleDoc: Boolean;

  function CreateTobDifference(CleDoc: R_CleDoc): Tob;
  begin
    Result := Tob.Create('LIGNE', TobDifference, -1);
    Result.PutValue('GL_NATUREPIECEG', CleDoc.NaturePiece);
    Result.PutValue('GL_DATEPIECE', CleDoc.DatePiece);
    Result.PutValue('GL_SOUCHE', CleDoc.Souche);
    Result.PutValue('GL_NUMERO', CleDoc.NumeroPiece);
    Result.PutValue('GL_INDICEG', CleDoc.Indice);
    Result.PutValue('GL_NUMLIGNE', CleDoc.NumLigne);
    Result.PutValue('GL_NUMORDRE', CleDoc.NumOrdre);
    Result.AddChampSupValeur('OLD_EXISTE', '-');
    Result.AddChampSupValeur('OLD_QTEFACT', 0);
    Result.AddChampSupValeur('OLD_QTESTOCK', 0);
    Result.AddChampSupValeur('NEW_EXISTE', '-');
    Result.AddChampSupValeur('NEW_QTEFACT', 0);
    Result.AddChampSupValeur('NEW_QTESTOCK', 0);
    Result.AddChampSupValeur('OLD_SOLDERELIQUAT', '-');
    Result.AddChampSupValeur('NEW_SOLDERELIQUAT', '-');
    Result.AddChampSupValeur('OLD_SOLDERLIGNE', '-');
    Result.AddChampSupValeur('NEW_SOLDERLIGNE', '-');
    Result.AddChampSupValeur('OLD_MTHTDEV', 0);
    Result.AddChampSupValeur('NEW_MTHTDEV', 0);
  end;

begin

  if Assigned(TobDifference) then
  begin
    if Assigned(TobPiece_O) then
    begin
      { Parcourt les lignes de l'ancienne TobPiece }
      for i := 0 to TobPiece_O.Detail.count - 1 do
      begin
        if EstLigneArticle(TobPiece_O.Detail[i]) then
        begin
          TobLigneOld := TobPiece_O.Detail[i];
          { Récupère la clé de la ligne d'origine }
          if OnlyWithPiecePrecedente then { Seulement les lignes issues d'une pièce }
            OkCleDoc := TobLigneOld.GetValue('GL_PIECEPRECEDENTE') <> ''
          else
            OkCleDoc := True;
          if OkCleDoc then
          begin
            CleDoc := TOB2CleDoc(TobLigneOld);
            CleDoc.NumLigne := TobLigneOld.GetValue('GL_NUMLIGNE');
            CleDoc.NumOrdre := TobLigneOld.GetValue('GL_NUMORDRE');
            TobModifFille := FindCleDocInTobPiece(CleDoc, TobDifference);
            if not Assigned(TobModifFille) then
              TobModifFille := CreateTobDifference(CleDoc);
            TobModifFille.PutValue('GL_TYPELIGNE', TobLigneOld.GetValue('GL_TYPELIGNE'));
            TobModifFille.PutValue('GL_TYPEARTICLE', TobLigneOld.GetValue('GL_TYPEARTICLE'));
            TobModifFille.PutValue('GL_ARTICLE', TobLigneOld.GetValue('GL_ARTICLE'));
            TobModifFille.PutValue('GL_CODEARTICLE', TobLigneOld.GetValue('GL_CODEARTICLE'));
            TobModifFille.PutValue('OLD_EXISTE', 'X');
            TobModifFille.PutValue('OLD_QTEFACT',  TobModifFille.GetValue('OLD_QTEFACT') + TobLigneOld.GetValue('GL_QTEFACT'));
            TobModifFille.PutValue('OLD_QTESTOCK', TobModifFille.GetValue('OLD_QTESTOCK') + TobLigneOld.GetValue('GL_QTESTOCK'));
            // --- GUINIER ---
            TobModifFille.PutValue('OLD_MTHTDEV',   TobModifFille.GetValue('OLD_MTHTDEV') + TobLigneOld.GetValue('GL_MONTANTHTDEV'));
          end;
        end;
      end;
    end;
    if Assigned(TobPiece) then
    begin
      { Parcourt les lignes de la nouvelle TobPiece }
      for i := 0 to TobPiece.Detail.count - 1 do
      begin
        if EstLigneArticle(TobPiece.Detail[i]) then
        begin
          TobLigneNew := TobPiece.Detail[i];
          if OnlyWithPiecePrecedente then
            OkCleDoc := TobLigneNew.GetValue('GL_PIECEPRECEDENTE') <> ''
          else
            OkCleDoc := True;
          if OkCleDoc then
          begin
            CleDoc := TOB2CleDoc(TobLigneNew);
            CleDoc.NumLigne := TobLigneNew.GetValue('GL_NUMLIGNE');
            CleDoc.NumOrdre := TobLigneNew.GetValue('GL_NUMORDRE');
            TobModifFille := FindCleDocInTobPiece(CleDoc, TobDifference);
            if not Assigned(TobModifFille) then
              TobModifFille := CreateTobDifference(CleDoc);
            TobModifFille.PutValue('NEW_EXISTE', 'X');
            TobModifFille.PutValue('NEW_QTEFACT',  TobModifFille.GetValue('NEW_QTEFACT') + TobLigneNew.GetValue('GL_QTEFACT'));
            TobModifFille.PutValue('NEW_QTESTOCK', TobModifFille.GetValue('NEW_QTESTOCK') + TobLigneNew.GetValue('GL_QTESTOCK'));
            TobModifFille.PutValue('NEW_MTHTDEV',  TobModifFille.GetValue('NEW_MTHTDEV') + TobLigneNew.GetValue('GL_MONTANTHTDEV'));

            TobModifFille.PutValue('NEW_SOLDERELIQUAT', TobLigneNew.GetValue('GL_SOLDERELIQUAT'));
            if TobLigneNew.FieldExists('SOLDERLIGNE') then
              TobModifFille.PutValue('NEW_SOLDERLIGNE', 'X');
          end;
        end;
      end;
    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 10/03/2003
Description .. : Récupère les lignes 'suivantes' d'une ligne
Description .. : Par exemple récupère les lignes de BL venant d'une commande
*****************************************************************}


function GetTobsSuiv(TobLigne: Tob; TobLignesSuiv: Tob): Boolean;
var
  Sql, PiecePrecedente: String;
  Q: TQuery;
  i: Integer;
begin
  Result := False;
  if Assigned(TobLigne) and Assigned(TobLignesSuiv) then
  begin
    PiecePrecedente := EncodeRefPiece(TobLigne);
    Sql := 'SELECT * FROM LIGNE WHERE GL_PIECEPRECEDENTE="' + PiecePrecedente + '" ORDER BY GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMLIGNE';
    Q := OpenSQL(Sql, True,-1, '', True);
    try
      TobLignesSuiv.LoadDetailDB('LIGNE', '', '', Q, False);
    finally
      Ferme(Q);
    end;
    { On ne récupère que les lignes articles }
    for i := TobLignesSuiv.Detail.Count -1 downto 0 do
    begin
      if not EstLigneArticle(TobLignesSuiv.Detail[i]) then
        TobLignesSuiv.Detail[i].Free;
    end;
    Result := (TobLignesSuiv.Detail.Count > 0);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 10/03/2003
Modifiée ..... : 29/04/2003
Description .. : Vérifie qu'une quantité est bien inférieure au
Suite ........ : total des quantités livrées pour une ligne de pièce
*****************************************************************}

function VerifQuantitePieceSuivante(TobL: Tob; Quantite: Double): Boolean; overload;
var DejaLivre: Double;
begin
  if Assigned(Tobl) and WithPieceSuivante(TobL) then // le 26/06/2003 Thierry Petetin
  begin
    DejaLivre := TobL.GetValue('GL_QTESTOCK') - TobL.GetValue('GL_QTERESTE');
    if CtrlOkReliquat(TobL, 'GL') then DejaLivre := TobL.GetValue('GL_MONTANTHTDEV') - TobL.GetValue('GL_MTRESTE');
    //
    Result := (Quantite >= DejaLivre);
  end
  else
    Result := True;
end;


function VerifQuantitePieceSuivante(TobPiece: Tob): Integer; overload;
var Champ   : String;
    i       : Integer;
    TobLigne: Tob;
begin
  Result := 0;
  if Assigned(TobPiece) then
  begin
    i := -1;
    while (i < TobPiece.Detail.Count - 1) and (Result = 0) do
    begin
      Inc(i);
      TobLigne := TobPiece.Detail[i];
      // --- GUINIER ---
      if CtrlOkReliquat(TobLigne, 'GL') then Champ := 'GL_MONTANTHTDEV' else Champ := 'GL_QTEFACT';
      if not VerifQuantitePieceSuivante(TobLigne, TobLigne.GetValue(Champ)) then Result := i;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/03/2003
Description .. : Contrôle si des lignes suivantes existent pour une ligne
*****************************************************************}

function WithPieceSuivante(TobLigne: Tob): Boolean; { NEWPIECE }
begin
  if Assigned(TobLigne) then
  begin
    // --- GUINIER ---
    if CtrlOkReliquat(TobLigne, 'GL') then
      Result := (TobLigne.GetValue('GL_MONTANTHTDEV') <> TobLigne.GetValue('GL_MTRESTE'))
    else
      Result := (TobLigne.GetValue('GL_QTESTOCK') <> TobLigne.GetValue('GL_QTERESTE'));
  end
  else
    Result := False;
end;

function IsExistAchatFrom (TOBpiece : TOB) : Boolean;
var StChaine : string;
		Req : string;
begin
  StChaine := '%;' + TOBPiece.GetValue('GP_NATUREPIECEG') + ';' + TOBPiece.GetValue('GP_SOUCHE') + ';'
    + IntToStr(TOBPiece.GetValue('GP_NUMERO')) + ';' + IntToStr(TOBPIECE.GetValue('GP_INDICEG')) + ';%';
  Req := 'SELECT ##TOP 1## BAD_REFGESCOM FROM DECISIONACHLIG LEFT JOIN DECISIONACH ON BAE_NUMERO=BAD_NUMERO WHERE BAD_REFGESCOM LIKE "'+StChaine+'" AND BAE_VALIDE <> "X"';
  Result := ExisteSQL(Req);
  if Result then Exit;
  Req := 'SELECT ##TOP 1## GL_PIECEORIGINE FROM LIGNE WHERE GL_PIECEORIGINE LIKE "'+StChaine+'"';
  Result := ExisteSQL(Req);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 06/03/2003
Description .. : Lors de la modification d'une pièce gérée en reliquat,
Suite ........ : remise à jour des quantités restant à livrer dans les lignes d'origine
*****************************************************************}

function UpdateResteALivrer(TobPiece, TobPiece_O, TobArticles, TobCatalogu, TOBOuvrage_O: Tob; Action: T_UpDateReliquat): Boolean; { NEWPIECE }
var
  i, Nb: Integer;
  IndiceNomen: Integer;
  TobModifMere, TobModifFille: Tob;
  TobPiecePrecedente: Tob;
  TobLigne, TobLignePiecePrecedente,TOBOuvrage_P: Tob;
  TobOuvrage: Tob;
  Sql, Signe,Article: string;
  CleDoc, CleDocPrecedente, OldCleDocPrecedente: R_CLeDoc;
  NaturePieceG, ColMoins, ColPlus, StPrixValo: string;
  Diff, RatioVA: Double;
  DiffMt        : Double;
  DiffDispo,  OldQteReste, NewQteReste : Double;
  DiffMtDispo,OldMtReste, NewMtReste              : Double;
  CoeffDispo: Integer;
  LigneSuppr, OkStock: Boolean;
  Ok_ReliquatMt,SolderReliquat, DesolderReliquat: Boolean;
  NewReste, QteTransfo, Delta,DeltaMt, Temp: Double;

begin
  Result := True;

  FillChar(OldCleDocPrecedente,Sizeof(OldCleDocPrecedente),#0);

  TobModifMere := Tob.Create('_MODIF_', nil, -1);

  try
    { Comparaison entre l'ancienne et la nouvelle piece }
    CompareTobPiece(TobPiece, TobPiece_O, TobModifMere, True);
    { Mise à jour des lignes origines }
    i := -1;
    while (i < TobModifMere.detail.count - 1) and (V_PGI.IoError = oeOk) do
    begin
      Inc(i);
      TobModifFille := TobModifMere.Detail[i];
      Diff        := 0;
      LigneSuppr  := False;
      CoeffDispo  := 1;
      DiffMt      := 0;
      Article := TobModifFille.getValue('GL_CODEARTICLE');
      Temp := TobModifFille.getValue('OLD_QTEFACT');

      Ok_ReliquatMt := CtrlOkReliquat(TobModifFille, 'GL');

      SolderReliquat    := False;
      DesolderReliquat  := False;
      case Action of
        urCreat: { Création d'une piece }
          begin
            if (TobModifFille.GetValue('NEW_EXISTE') = 'X') then
            begin
              Diff := -1 * TobModifFille.GetValue('NEW_QTEFACT');
              DiffMt := -1 * TobModifFille.GetValue('NEW_MTHTDEV');
            end;
          end;
        urModif: { Modification d'une piece }
          begin
            if (TobModifFille.GetValue('OLD_EXISTE') = 'X') and (TobModifFille.GetValue('NEW_EXISTE') = 'X') then
            begin
              if Ok_ReliquatMt then
              begin
                if (TobModifFille.GetValue('OLD_MTHTDEV') <> TobModifFille.GetValue('NEW_MTHTDEV')) then
                begin
                  DiffMT := TobModifFille.GetValue('OLD_MTHTDEV') - TobModifFille.GetValue('NEW_MTHTDEV');
                  CoeffDispo := -1;
                end
                else if (TobModifFille.GetValue('OLD_SOLDERELIQUAT') <> TobModifFille.GetValue('NEW_SOLDERELIQUAT')) then
                begin
                  if TobModifFille.GetValue('NEW_SOLDERELIQUAT') = 'X' then
                    SolderReliquat := True { On vient de solder une ligne dans un BL }
                  else if TobModifFille.GetValue('NEW_SOLDERELIQUAT') = '-' then
                    DesolderReliquat := True;
                end;
              end else
              begin
                if (TobModifFille.GetValue('OLD_QTEFACT') <> TobModifFille.GetValue('NEW_QTEFACT')) then
                begin
                  Diff := TobModifFille.GetValue('OLD_QTEFACT') - TobModifFille.GetValue('NEW_QTEFACT');
                  CoeffDispo := -1;
                end
                else if (TobModifFille.GetValue('OLD_SOLDERELIQUAT') <> TobModifFille.GetValue('NEW_SOLDERELIQUAT')) then
                begin
                  if TobModifFille.GetValue('NEW_SOLDERELIQUAT') = 'X' then
                    SolderReliquat := True { On vient de solder une ligne dans un BL }
                  else if TobModifFille.GetValue('NEW_SOLDERELIQUAT') = '-' then
                    DesolderReliquat := True;
                end;
              end;
            end
            else if (TobModifFille.GetValue('OLD_EXISTE') = 'X') and (TobModifFille.GetValue('NEW_EXISTE') = '-') then
            begin
              if Ok_ReliquatMt then
              begin
               { Modification d'une pièce avec une suppression de ligne }
               DiffMT := TobModifFille.GetValue('OLD_MTHTDEV');
               LigneSuppr := True;
               CoeffDispo := -1;
              end else
              begin
               { Modification d'une pièce avec une suppression de ligne }
               Diff := TobModifFille.GetValue('OLD_QTEFACT');
               LigneSuppr := True;
               CoeffDispo := -1;
              end;
            end;
          end;
        urSuppr: { Suppression de piece }
          begin
            if Ok_ReliquatMt then
            begin
              if (TobModifFille.GetValue('OLD_EXISTE') = 'X') and (TobModifFille.GetValue('NEW_EXISTE') = 'X') then
              begin
                DiffMT := TobModifFille.GetValue('OLD_MTHTDEV');
                LigneSuppr := True;
                if DiffMT < 0 then DiffMT := Abs(DiffMT);
                CoeffDispo := -1;
              end;
            end else
            begin
              if (TobModifFille.GetValue('OLD_EXISTE') = 'X') and (TobModifFille.GetValue('NEW_EXISTE') = 'X') then
              begin
                Diff := TobModifFille.GetValue('OLD_QTEFACT');
                LigneSuppr := True;
                if Diff < 0 then Diff := Abs(Diff);
                CoeffDispo := -1;
              end;
            end;
          end;
      end;
      { Les quantités ont été modifiées }
      if (Diff <> 0) or (DiffMT <> 0) or SolderReliquat or DesolderReliquat then
      begin
        { Recherche de la ligne de pièce }
        CleDoc := TOB2CleDoc(TobModifFille);
        CleDoc.NumLigne := TobModifFille.GetValue('GL_NUMLIGNE');
        CleDoc.NumOrdre := TobModifFille.GetValue('GL_NUMORDRE');
        if not LigneSuppr then
          TobLigne := FindCleDocInTobPiece(CleDoc, TobPiece)
        else
          TobLigne := FindCleDocInTobPiece(CleDoc, TobPiece_O);
        if Assigned(TobLigne) then
        begin
          { Requete de mise à jour du reste à livrer }
          DecodeRefPiece(TobLigne.GetValue('GL_PIECEPRECEDENTE'), CleDocPrecedente);
            {
              iQualite 10999
              Avant la nouvelle gestion des reliquats, lors d'une modification de pièce, le GL_PIECEPRECEDENTE
              de la nouvelle pièce pointait vers l'ancienne pièce (bien qu'il n'y est pas eu de transformation)
              Ce test empèche donc la mise à jour de la pièce précédente dans ce cas.
            }
            // Ajout LS Suite amise ajour intempstive des previsions suite a modif des besoins...
            // mais c'est quoi que je prends comme produits dejà ??
            // if CleDocPrecedente.NaturePiece = 'PBT' then continue;
            // --
            if (((CleDoc.NaturePiece <> CleDocPrecedente.NaturePiece)
                or (CleDoc.Souche <> CleDocPrecedente.Souche)
                or (CleDoc.NumeroPiece <> CleDocPrecedente.NumeroPiece))) and
                (GetInfoParPiece (Cledoc.NaturePiece,'GPP_RELIQUAT')='X') then
            begin
              { Mise à jour du Delta pour la pièce précédente dans le dispo }
              { Récupère la pièce précédente et ses lignes pour mise à jour du Dispo }
              TobPiecePrecedente := Tob.Create('PIECE', nil, -1);
              TOBOuvrage_P := TOB.Create ('LES OUV',nil,-1);
              try
                { Charge la pièce précédente avant la mise à jour }
                ChargelaPieceEtUneLigne (CleDocPrecedente, TobPiecePrecedente,TOBOuvrage_P);
//                LoadPieceLignes(CleDocPrecedente, TobPiecePrecedente,true,True);
                { Récupère la ligne précédente dans la pièce précédente }
//                TobLignePiecePrecedente := FindCleDocInTobPiece(CleDocPrecedente, TobPiecePrecedente);
								if TobPiecePrecedente.Detail.Count > 0 then
                	TobLignePiecePrecedente := TobPiecePrecedente.detail[0]
                else
                	TobLignePiecePrecedente := nil;
                {GPAO_TP_RELIQUAT_SOLDESUPPRLIG-Début}
                if SolderReliquat or DeSolderReliquat then
                begin
                  Diff    := 0;
                  DiffMt  := 0;
                  if Assigned(TobLignePiecePrecedente) then
                  begin
                    if SolderReliquat then
                    begin
                      // --- GUINIER ---
                      if Ok_ReliquatMt then
                      Begin
                        if TobLignePiecePrecedente.GetValue('GL_MTRESTE') > 0 then DiffMt := -1 * TobLignePiecePrecedente.GetValue('GL_MTRESTE');
                      end else
                      begin
                        if TobLignePiecePrecedente.GetValue('GL_QTERESTE') > 0 then Diff := -1 * TobLignePiecePrecedente.GetValue('GL_QTERESTE');
                      end;
                    end
                    else { Désolde le reliquat : Il faut remettre à jour le GL_QTERESTE }
                    begin
                      { Total des quantités déjà transformée + Piece actuelle qui vient d'être supprimée }
                      QteTransfo    := CalculResteALivrer(TobLignePiecePrecedente, TobPiece);
                      NewReste      := TobLignePiecePrecedente.GetValue('GL_QTESTOCK') - QteTransfo;
                      Diff          := NewReste - TobLignePiecePrecedente.GetValue('GL_QTERESTE');
                      // --- GUINIER ---
                      if Ok_ReliquatMt then
                      Begin
                        QteTransfo  := CalculResteALivrer(TobLignePiecePrecedente, TobPiece);
                        NewReste    := TobLignePiecePrecedente.GetValue('GL_MONTANTHTDEV') - QteTransfo;
                        DiffMt        := NewReste - TobLignePiecePrecedente.GetValue('GL_MTRESTE');
                      end;
                    end;
                  end;
                end;
                {GPAO_TP_RELIQUAT_SOLDESUPPRLIG-Fin}
                //
                if TobLignePiecePrecedente = nil then exit;
                // --- GUINIER ---

                //

                if Ok_ReliquatMt then
                begin
                  DeltaMT := TobLignePiecePrecedente.GetValue('GL_MTRESTE') + DiffMt;
                  if DiffMt >= 0 then
                  begin
                    Signe := '+';
                    // --- GUINIER ---
                    if DeltaMt > TobLignePiecePrecedente.GetValue('GL_MONTANTHTDEV')  then DiffMt := TobLignePiecePrecedente.GetValue('GL_MONTANTHTDEV') - TobLignePiecePrecedente.GetValue('GL_MTRESTE');
                  end
                  else
                  begin
                    Signe := '-';
                    if DeltaMt < 0  then	Diffmt := TobLignePiecePrecedente.GetValue('GL_MTRESTE') * (-1);
                    If DiffMt = 0   then  Signe := '-';
                    if diffmt < 0   then  diffmt := ABS(diffmt); Signe := '-';
                  end;
                end else
                begin
                  Delta := TobLignePiecePrecedente.GetValue('GL_QTERESTE') + Diff;
                  if Diff >= 0 then
                  begin
                    Signe := '+';
                    if Delta > TobLignePiecePrecedente.GetValue('GL_QTEFACT')         then Diff := TobLignePiecePrecedente.GetValue('GL_QTEFACT') - TobLignePiecePrecedente.GetValue('GL_QTERESTE');
                  end
                  else
                  begin
                    Signe := '-';
                    if Delta < 0 then	Diff := TobLignePiecePrecedente.GetValue('GL_QTERESTE') * (-1);
                    If Diff = 0 then Signe := '-';
                    if diff < 0 then diff := ABS(diff); Signe := '-';
                  end;
                end;

                if (diff <> 0) Or (DiffMt <> 0) then
                begin
                  { Mise à jour du reste à livrer dans la ligne précédente }
                  // --- GUINIER ---
                  Sql := 'UPDATE LIGNE SET ';
                  if Ok_ReliquatMt then
                  begin
                    Sql := Sql + 'GL_MTRESTE = GL_MTRESTE ' + Signe + ' ' + StrFPoint(DiffMt)+', GL_QTERESTE=GL_QTEFACT';
                  end else
                  begin
                    Sql := Sql + 'GL_QTERESTE = GL_QTERESTE ' + Signe + ' ' + StrFPoint(Diff);
                  end;
                  //
                  Sql := SQL + ',GL_VIVANTE="X" WHERE ' + WherePiece(CleDocPrecedente, ttdLigne, True, True);
                  Result := (ExecuteSQL(Sql) = 1);
                  if Result then
                  begin
                    { MODE DCA 27/08/2003 - Modif/Suppr - Mise à jour de GP_DATEMODIF de la pièce précédente }
                    if ((OldCleDocPrecedente.NaturePiece <> CleDocPrecedente.NaturePiece) or
                        (OldCleDocPrecedente.Souche <> CleDocPrecedente.Souche) or
                        (OldCleDocPrecedente.NumeroPiece <> CleDocPrecedente.NumeroPiece) or
                        (OldCleDocPrecedente.Indice <> CleDocPrecedente.Indice)) then
                    begin
                      { Mise à jour de GP_DATEMODIF de la pièce précédente }
                      Sql := 'UPDATE PIECE SET GP_DATEMODIF="' + USTime(NowH) + '",GP_VIVANTE="X" WHERE ' + WherePiece(CleDocPrecedente, ttdPiece, False);
                      Result := (ExecuteSQL(Sql) = 1);
                      OldCleDocPrecedente.NaturePiece := CleDocPrecedente.NaturePiece;
                      OldCleDocPrecedente.Souche := CleDocPrecedente.Souche;
                      OldCleDocPrecedente.NumeroPiece := CleDocPrecedente.NumeroPiece;
                      OldCleDocPrecedente.Indice := CleDocPrecedente.Indice;
                    end;
                    if Result then
                    begin
                      OkStock := True;
                      if Assigned(TobLignePiecePrecedente) then
                      begin
                        NaturePieceG := TobPiecePrecedente.GetValue('GP_NATUREPIECEG');
                        ColMoins := GetInfoParPiece(NaturePieceG, 'GPP_QTEMOINS');
                        ColPlus := GetInfoParPiece(NaturePieceG, 'GPP_QTEPLUS');
                        StPrixValo := GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO');
                        { Mise à jour de la quantité }
//                        OldQteReste := Max(0, Double(TobLignePiecePrecedente.G('GL_QTERESTE')));
//                        NewQteReste := Max(0, Double(TobLignePiecePrecedente.G('GL_QTERESTE')) + Diff);
                        // --- GUINIER ---
                        if Ok_ReliquatMt then
                        begin
                          OldMtReste := TobLignePiecePrecedente.G('GL_MTRESTE');
                          if Signe = '+' then
                            NewMtReste := TobLignePiecePrecedente.G('GL_MTRESTE')+diffMt
                          else
                            NewMtReste := TobLignePiecePrecedente.G('GL_MTRESTE')-diffMt;
                          DiffMtDispo := OldMtReste - NewMtReste;
                          TobLignePiecePrecedente.PutValue('GL_MONTANTHTDEV', DiffMtDispo);
                          if TobLignePiecePrecedente.GetDouble('GL_MONTANTHTDEV')=0 then TobLignePiecePrecedente.SetDouble('GL_QTERESTE',0)
                                                                                    else TobLignePiecePrecedente.SetDouble('GL_QTERESTE',TobLignePiecePrecedente.GetDouble('GL_QTEFACT'));
                        end else
                        begin
                          OldQteReste := TobLignePiecePrecedente.G('GL_QTERESTE');
                          if Signe = '+' then
                            NewQteReste := TobLignePiecePrecedente.G('GL_QTERESTE')+diff
                          else
                            NewQteReste := TobLignePiecePrecedente.G('GL_QTERESTE')-diff;
                          DiffDispo := OldQteReste - NewQteReste;
                        end;
                        if DiffDispo <> 0 then
                        begin
                          DiffDispo   := DiffDispo * CoeffDispo;
                          TobLignePiecePrecedente.PutValue('GL_QTESTOCK', DiffDispo);
                          TobLignePiecePrecedente.PutValue('GL_QTEFACT',  DiffDispo);
                          RatioVA     := GetRatio(TobLignePiecePrecedente, nil, trsStock);
                          IndiceNomen := TobLignePiecePrecedente.GetValue('GL_INDICENOMEN');
                          if (IndiceNomen > 0) and Assigned(TOBOuvrage_O) and (TOBOuvrage_O.detail.count >0)  then
                            TobOuvrage := TOBOuvrage_O.Detail[IndiceNomen - 1]
                          else
                            TobOuvrage := nil;
                          { Ajout de la TobNomen dans les paramètres }
                          OkStock := MajQteStock(TobLignePiecePrecedente, TobArticles, TOBouvrage, ColPlus, ColMoins, True, RatioVA);
                        end;
                      end;
                      { On vient de mettre à jour le reste à livrer, mise à jour de l'état de la pièce  d'origine }
                      if OkStock then
                        UpDateEtatPiecePrecedente(CleDocPrecedente)
                      else
                      begin
                        PgiError ('Erreur Mise à jour des stocks');
                        V_PGI.IoError := oeStock;
                      end;
                    end
                    else
                    begin
                      if V_PGI.Sav then
                        PGIError(TraduireMemoire('La requête') + ' : ' + Sql + ' ' + TraduireMemoire('a échoué')
                            , TraduireMemoire('Mise à jour de la date de modification dans la pièce origine'));
                        V_PGI.IoError := oeUnknown;
                    end;
                  end else
                  begin
                    if V_PGI.Sav then
                      PGIError(TraduireMemoire('La requête') + ' : ' + Sql + ' ' + TraduireMemoire('a échoué')
                        , TraduireMemoire('Mise à jour des quantités dans la pièce origine'));
                    V_PGI.IoError := oeUnknown;
                  end;
                end;
              finally
                TobPiecePrecedente.Free;
                TOBOuvrage_P.Free;
              end;
            end
            else if (((CleDoc.NaturePiece <> CleDocPrecedente.NaturePiece)
                   or (CleDoc.Souche <> CleDocPrecedente.Souche)
                   or (CleDoc.NumeroPiece <> CleDocPrecedente.NumeroPiece))) and
                      (GetInfoParPiece (Cledoc.NaturePiece,'GPP_RELIQUAT')='-') then
            begin
              // --- GUINIER ---
              Sql := 'UPDATE LIGNE SET GL_VIVANTE="X",GL_QTERESTE=GL_QTEFACT ';
              If Ok_ReliquatMt then Sql := Sql + ',GL_MTRESTE=GL_MONTANTHTDEV ';
              Sql := Sql + ' WHERE ' + WherePiece(CleDocPrecedente, ttdLigne, True, True);
              Result := (ExecuteSQL(Sql) = 1);
              if Result then
              begin
                { Mise à jour de la pièce }
                Sql := 'UPDATE PIECE SET GP_VIVANTE="X" WHERE ' + WherePiece(CleDocPrecedente, ttdPiece, False);
                Result := ExecuteSQL(Sql) = 1;
              end;
            end;
        end;
      end;
    end;
  finally
    TobModifMere.Free;
  end;
end;

(*
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 13/03/2003
Description .. : Mise à jour des quantités en reliquat dans les lignes suivantes de la pièce
*****************************************************************}

    Le champ GL_QTERELIQUAT est correct tant que l'on ne modifie pas la quantité d'origine

    Il est toutefois trop 'couteux' de le tenir à jour en permanence
    (nécessite de remettre à jour toutes les pièces suivantes à chaque modification de la quantité originelle)

    Cette fonction est désactivée, mais elle est fonctionnelle donc au cas ou...

procedure UpdateQuantiteReliquat(TobPiece, TobPiece_O: Tob); { NEWPIECE }
var
  i,j: Integer;
  TobDifference, TobModifFille: Tob;
  TobLigne, TobLignesSuiv: Tob;
  Diff, ALivrer, DejaLivre: Double;
  CleDoc: R_CLeDoc;
begin
  TobDifference := Tob.Create('_MODIF_', nil, -1);
  try
    { Compare les quantités dans l'ancienne et la nouvelle TobPiece }
    CompareTobPiece(TobPiece, TobPiece_O, TobDifference, False);
    { Mise à jour des lignes }
    i := -1;
    while (i < TobDifference.Detail.count - 1) and (V_PGI.IoError = oeOk) do
    begin
      Inc(i);
      TobModifFille := TobDifference.Detail[i];
      if (TobModifFille.GetValue('OLD_EXISTE') = 'X') and (TobModifFille.GetValue('NEW_EXISTE') = 'X') then
      begin
        if (TobModifFille.GetValue('OLD_QTEFACT') <> TobModifFille.GetValue('NEW_QTEFACT')) then
        begin
          Diff := TobModifFille.GetValue('OLD_QTEFACT') - TobModifFille.GetValue('NEW_QTEFACT');
          { La quantité a été modifiée }
          if Diff <> 0 then
          begin
            { Récupère la TobLigne dans TobPiece }
            CleDoc := TOB2CleDoc(TobModifFille);
            CleDoc.NumLigne := TobModifFille.GetValue('GL_NUMLIGNE');
            TobLigne := FindCleDocInTobPiece(CleDoc, TobPiece);
            if Assigned(TobLigne) then
            begin
              { Récupère les lignes suivantes }
              TobLignesSuiv := Tob.Create('LIGNE', nil, -1);
              try
                GetTobsSuiv(TobLigne, TobLignesSuiv);
                DejaLivre := TobModifFille.GetValue('NEW_QTEFACT');
                for j := 0 to TobLignesSuiv.Detail.Count - 1 do
                begin
                  if EstLigneArticle(TobLignesSuiv.Detail[j]) then
                  begin
                    TobLignesSuiv.SetAllModifie(False);
                    DejaLivre := DejaLivre - TobLignesSuiv.Detail[j].GetValue('GL_QTEFACT');
                    TobLignesSuiv.Detail[j].PutValue('GL_QTERELIQUAT', DejaLivre);
                    TobLignesSuiv.Detail[j].UpdateDB(False);
                  end;
                end;
              finally
                TobLignesSuiv.Free;
              end;
            end;
          end;
        end;
      end
      else
      begin
      end;
    end;
  finally
    TobDifference.free;
  end;
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/03/2003
Description .. : Recalcul le GL_QTERELIQUAT d'une ligne de pièce
*****************************************************************}

procedure PutQteReliquat(TobL: Tob; OldQte: Double); { NEWPIECE }
Var Ok_ReliquatMt : Boolean;
begin

  if Assigned(TobL) then
  begin
    if (TobL.GetValue('GL_PIECEPRECEDENTE') <> '') and IsPieceGerableReliquat(TOBL) then
    begin
      if (TobL.GetString('GL_NATUREPIECEG')='BFA') then
      begin
      	TobL.PutValue('GL_QTERELIQUAT', 0);
        Exit;
      end;
      if Abs(TobL.GetValue('GL_QTERELIQUAT')) + Abs(OldQte) - Abs(TobL.GetValue('GL_QTESTOCK')) <= 0 then
        TobL.PutValue('GL_QTERELIQUAT', 0)
      else
        TobL.PutValue('GL_QTERELIQUAT', TobL.GetValue('GL_QTERELIQUAT') + OldQte - TobL.GetValue('GL_QTESTOCK'));
    end;
  end;

end;

procedure PutMTReliquat(TobL: Tob; OldMT: Double); { NEWPIECE }
Var Ok_ReliquatMt : Boolean;
begin

  if Assigned(TobL) then
  begin
    if (TobL.GetValue('GL_PIECEPRECEDENTE') <> '') and IsPieceGerableReliquat(TOBL) then
    begin
      if (TobL.GetString('GL_NATUREPIECEG')= 'BFA') then
      begin
        TobL.PutValue('GL_MTRELIQUAT',  0);
        Exit;
      end;
      if Abs(TobL.GetValue('GL_MTRELIQUAT')) + Abs(OldMT) - Abs(TobL.GetValue('GL_MTRESTE')) <= 0 then
        TobL.PutValue('GL_MTRELIQUAT', 0)
      else
        TobL.PutValue('GL_MTRELIQUAT', TobL.GetValue('GL_MTRELIQUAT') + OldMT - TobL.GetValue('GL_MTRESTE'));
    end;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 22/05/2003
Description .. : Lors de la transformation d'une pièce sans gestion des reliquats,
Suite ........ : mets les reste à livrer à 0 dans la pièce d'origine et 'tue' la pièce d'origine
*****************************************************************}

procedure RazResteALivrerAndKillLesPieces(TOBListPieces : TOB);
var II : Integer;
begin
	for II := 0 to TOBListPieces.detail.Count-1  do
  begin
    RazResteALivrerAndKillPiece(TOBListPieces.detail[II]);
  end;
end;

procedure RazResteALivrerAndKillPiece(TobPiece_O: Tob);
var
  i: Integer;
begin
  if Assigned(TobPiece_O) then
  begin
    RazResteALivrer(TobPiece_O);
    for i := 0 to TobPiece_O.Detail.Count - 1 do
      TobPiece_O.Detail[i].PutValue('GL_VIVANTE', '-');
    TobPiece_O.PutValue('GP_VIVANTE', '-');
    TobPiece_O.UpdateDB(False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/03/2003
Description .. : Retourne Vrai si la ligne contient un article
*****************************************************************}

function EstLigneArticle(TobL: Tob): Boolean; { NEWPIECE }
begin
  Result := False;
  if Assigned(TobL) then
  begin
    // VARIANTE
    (* if TobL.GetValue('GL_TYPELIGNE') <> 'COM' then *)
    if not IsCommentaire(TOBL) then
    begin
      if TobL.GetValue('GL_TYPEREF') = 'CAT' then
        Result := True
      else if TobL.GetValue('GL_TYPEREF') = 'ART' then
        Result := True
      else
        Result := (TobL.GetValue('GL_REFCATALOGUE') <> '') or (TobL.GetValue('GL_ARTICLE') <> '');
    end
    else
      Result := (TobL.GetValue('GL_TYPEDIM') = 'GEN');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/03/2003
Description .. : Crée et charge une TobLigne à partir d'une CleDoc
*****************************************************************}

function CleDoc2TobLigne(CleDocL: R_CleDoc; TobLigne: Tob): Boolean; { NEWPIECE }
var
  Q: TQuery;
begin
  Result := Assigned(TobLigne);
  if Result then
  begin
    Q := OpenSQL('SELECT * FROM LIGNE WHERE ' + WherePiece(CleDocL, ttdLigne, True, True), True,-1, '', True);
    try
      Result := (not Q.Eof);
      if Result then
        TobLigne.SelectDB('', Q);
    finally
      Ferme(Q);
    end;
  end;
end;

function ExisteLignePrec(TobL: TOB): Boolean; { NEWPIECE }
var
  StPrec: string;
  CleL: R_CleDoc;
begin
  Result := False;
  if Assigned(TobL) and EstLigneArticle(TobL) then
  begin
    StPrec := TobL.GetValue('GL_PIECEPRECEDENTE');
    if StPrec <> '' then
    begin
      DecodeRefPiece(StPrec, CleL);
      Result := ExisteSQL('SELECT GL_NATUREPIECEG FROM LIGNE WHERE ' + WherePiece(CleL, ttdLigne, True));
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 27/03/2003
Description .. : Compare deux Record CleDoc et retourne Vrai si ils sont
Suite ........ : identiques
*****************************************************************}

function CompareCleDoc(Source, Cible: R_Cledoc): Boolean;
begin
  Result := (Source.NaturePiece = Cible.NaturePiece)
    and (Source.Souche = Cible.Souche)
    and (Source.NumeroPiece = Cible.NumeroPiece)
    and (Source.Indice = Cible.Indice);
end;

function CanModifyLigne(TOBL: TOB; TransfoPiece: Boolean; TobPiece_O: TOB;Var WarningSupressionRecep : boolean; gestionreliquat : boolean=true;Action : TactionFiche=TaCreat): Boolean;
(*
var
  TOBL_O: Tob;
{$IFDEF BTP}
  PieceAchat : string;
{$ENDIF}
*)
begin
  Result :=  True;
(*
  if Assigned(TOBL) then
  begin
    if EstLigneArticle(TOBL) then
    begin
{$IFDEF BTP}
			PieceAchat := GetPieceAchat (false,false,false,true);
      if (Action = TaModif) and (Pos(TobL.getValue('GL_NATUREPIECEG'),PieceAchat)>0) and (TOBL.GetValue('BCO_TRAITEVENTE') = 'X' ) then
      begin
      	WarningSupressionRecep := true;
      end;
{$ENDIF}
      { Dans tous les cas si la ligne est soldée on ne peut plus toucher à rien }
      if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then
      begin
        Result := not EstLigneSoldee(TOBL)
      end else
      begin
        { Vérifie si la ligne précédente n'est pas soldée }
        TOBL_O := GetTOBPrec(TOBL, TOBPiece_O);
        if TransfoPiece then { La piece précédente est dans TOBPIECE_O }
        begin
          Result := not EstLigneSoldee(TOBL_O);
        end
        else
        begin

{$IFDEF BTP}
					if (GetInfoParPiece (TOBL_O.GetValue('GL_NATUREPIECEG'),'GPP_RELIQUAT')='X') then exit;

					if TOBL.GetValue('GL_NATUREPIECEG')='PBT' Then exit; // pas de controle sur les previsions
          //
          if (TobL.FieldExists ('LPREC_QTERESTE')) and (Vartype (TobL.GetValue ('LPREC_QTERESTE')) = varstring) then
          begin
            TOBL.putvalue('LPREC_QTERESTE',valeur(TobL.GetValue ('LPREC_QTERESTE')));
          end;
{$ENDIF}
          if TOBL.FieldExists('LPREC_SOLDERELIQUAT') then
            Result := ((TOBL.GetValue('LPREC_SOLDERELIQUAT') = '-') and (TobL.GetValue ('LPREC_QTERESTE') <> 0));
        end;
      end;
    end;
  end;
*)
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/03/2003
Description .. : Contrôle si on peut supprimer une pièce
*****************************************************************}

function CanDeletePiece(TobPiece, TobPiece_O: Tob; TransfoPiece: Boolean; var NumL: Integer;Action : TActionFiche = TaCreat): Boolean;
var WarningSuppression : boolean;
		NaturePiece : string;
//		GestionReliquat : Boolean;
begin
	NaturePiece := TobPiece.GetValue('GP_NATUREPIECEG');
	WarningSuppression := false;
  Result := CanModifyPiece(TobPiece);
//  GestionReliquat := (GetInfoParPiece (TOBPiece_O.GetValue('GP_NATUREPIECEG'),'GPP_RELIQUAT')='X');
  if (Result) and (NaturePiece='CBT') then
  begin
    // FS#878 - TEST BRL : il est possible de supprimer un besoin de chantier généré en décisionnel.
    Result := not IsExistAchatFrom(TOBPiece);
    if Result then ;
  end;
  if Result then
  begin
    NumL := -1;
    if Assigned(TobPiece) then
    begin
      while (NumL < TobPiece.Detail.Count - 1) and Result do
      begin
        Inc(NumL);
        if EstLigneArticle(TobPiece.Detail[NumL]) then
        begin
        	if (TobPiece.Detail[NumL].FieldExists ('MODIFIABLE')) and (TobPiece.Detail[NumL].GetValue('MODIFIABLE')='-') then result := false;
          if result then Result := not WithPieceSuivante(TobPiece.Detail[NumL]);
          (*
          if Result then
            Result := CanModifyLigne(TobPiece.Detail[NumL], TransfoPiece, TobPiece_O,WarningSuppression,GestionReliquat,Action);
          *)
        end;
      end;
    end;
{$IFDEF BTP}
    if WarningSuppression then
    begin
      PGIInfo (TraduireMemoire ('Attention : Certains articles ont déjà été livrés.#13#10Les consommations ne seront pas mises à jour.#13#10Veuillez mettre à jour les livraisons de chantiers'),'Suppression de pièce');
    end;
{$ENDIF}

  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/03/2003
Modifiée le .. : 26/09/2003
Description .. : Contrôle si on peut modifier une pièce
                 Ligne par ligne
                 Si la ligne est issue d'une pièce précédente, il faut que celle-ci
                 pointe vers le dernier indice existant, pour ne pas risquer de réactiver
                 des pièces d'indice n-1, alors que la piece d'indice n est elle même active.
*****************************************************************}
function CanModifyPiece(TobPiece: Tob): Boolean;

  function IsDevisFacture (Zcledoc : r_Cledoc) : boolean;
  begin
    result := ExisteSql ('SELECT 1 FROM LIGNE WHERE '+
                        WherePiece( Zcledoc,Ttdligne,false) + ' AND '+
                        '(GL_QTEPREVAVANC<> 0 OR GL_QTESIT<>0) AND GL_TYPELIGNE="ART"');
  end;
var
  CleDoc, PrevCleDoc, CurrPieceCleDoc,Zcledoc: R_CleDoc;
  i, MaxIndice: Integer;
begin
  MaxIndice := -1;
  //
  if (TOBPiece.getValue('GP_ETATVISA')='VIS') then BEGIN result := false; exit; END;
  //
  FillChar(PrevCleDoc, Sizeof(PrevCleDoc), #0);
  FillChar(CurrPieceCleDoc, Sizeof(CurrPieceCleDoc), #0);
  { CleDoc de la pièce à traiter }
  CurrPieceCleDoc := TOB2CleDoc(TobPiece);
{$IFDEF BTP}
	result := VerifEcriturePieceModifiable (TOBPiece);
  if not result then exit;
  if TOBPiece.getString('GP_NATUREPIECEG')='B00' then
  begin
    if TOBpiece.getString('GP_ATTACHEMENT')='' then exit;
    DecodeRefPiece(TOBpiece.getString('GP_ATTACHEMENT'),Zcledoc);
    result := not isDevisfacture (Zcledoc);
  end;
  if not result then exit;
{$ENDIF}
  i := -1;
  Result := True;
  while (i < TobPiece.Detail.Count - 1) and Result do
  begin
    Inc(i);
    if EstLigneArticle(TobPiece.Detail[i]) and (TobPiece.Detail[i].GetValue('GL_PIECEPRECEDENTE') <> '') then
    begin
      DecodeRefPiece(TobPiece.Detail[i].GetValue('GL_PIECEPRECEDENTE'), CleDoc);
      {
        Si même Nature, Souche et N°, la modification est autorisée
        Cas de la modification d'une commande (sans transformation).
        Dans la commande d'indice n+1, les GL_PIECEPRECEDENTE pointent vers la piece d'indice - 1
      }
      if (CleDoc.NaturePiece <> CurrPieceCleDoc.NaturePiece) or (CleDoc.Souche <> CurrPieceCleDoc.Souche)
                                                        or (CleDoc.NumeroPiece <> CurrPieceCleDoc.NumeroPiece) then
      begin
        { optimisation pour ne pas relancer systématiquement la recherche du dernière indice }
        if (CleDoc.NaturePiece <> PrevCleDoc.NaturePiece) or (CleDoc.Souche <> PrevCleDoc.Souche)
                                                          or (CleDoc.NumeroPiece <> PrevCleDoc.NumeroPiece) then
           MaxIndice := GetMaxIndice(CleDoc);

        Result := (CleDoc.Indice = MaxIndice);
      end;
      PrevCleDoc := CleDoc;
    end;
  end;
end;

function GetMaxIndice(CleDoc: R_CleDoc): Integer;
var
  Q: TQuery;
begin
  Q := OpenSQL('SELECT MAX(GL_INDICEG) FROM LIGNE WHERE GL_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND '
                                                     + 'GL_SOUCHE="' + CleDoc.Souche + '" AND '
                                                     + 'GL_NUMERO=' + IntToStr(CleDoc.NumeroPiece), True,-1, '', True);
  if Q.Eof then
    Result := -1
  else
    Result := Q.Fields[0].AsInteger;
  Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 27/03/2003
Description .. : Après avoir mis à jour le GL_QTERESTE sur CleDocLigne,
Suite ........ : remise à jour de GL_VIVANTE et GP_VIVANTE pour la
Suite ........ : pièce précédente
*****************************************************************}

function UpDateEtatPiecePrecedente(CleDocLignePrecedente: R_CleDoc): Boolean;
var
  TobLignePrecedente, TobPiecePrecedente: Tob;
  NewEtatLigne: string;
  i: Integer;
  AutresLignesMortes, AutresLignesVivantes: Boolean;
  Sql: string;

  function LigneIdentique(Tobl: Tob): Boolean;
  var
    CurrentCleDoc: R_CleDoc;
  begin
    Result := False;
    if Assigned(Tobl) then
    begin
      CurrentCleDoc := TOB2CleDoc(Tobl);
      CurrentCleDoc.NumOrdre := TobL.GetValue('GL_NUMORDRE');
      Result := CompareCleDoc(CleDocLignePrecedente, CurrentCleDoc);
      if Result then
        Result := (CleDocLignePrecedente.NumOrdre = CurrentCleDoc.NumOrdre);
    end;
  end;

begin
  Result := True;
  Sql := '';
  TobPiecePrecedente := TOB.Create('PIECE', nil, -1);
  TobLignePrecedente := TOB.Create('LIGNE', nil, -1);
  InitTobLigne(TobLignePrecedente);
  try
    { Récupère la TobLigne correspondante }
    if CleDoc2TobLigne(CleDocLignePrecedente, TobLignePrecedente) then
    begin
      { Nouvel état de la ligne dont on vient d'updater le GL_QTERESTE ( >= 0 (Jamais négatif) }
      if not EstLigneSoldee(TobLignePrecedente) then
        NewEtatLigne := 'X'
      else
        NewEtatLigne := '-';
      if NewEtatLigne <> TobLignePrecedente.GetValue('GL_VIVANTE') then { L'état de la ligne change }
      begin
        { Charge la pièce d'origine }
        LoadPieceLignes(CleDocLignePrecedente, TobPiecePrecedente);
        { }
        if NewEtatLigne = '-' then { La ligne meurt, mais ne se rend pas ! }
        begin
          { Vérifie si toutes les autres lignes de la pièce sont dans le même état }
          i := -1;
          AutresLignesMortes := True;
          while (i < TobPiecePrecedente.Detail.Count - 1) and AutresLignesMortes do
          begin
            Inc(i);
            if (not LigneIdentique(TobPiecePrecedente.Detail[i])) then
              AutresLignesMortes := (TobPiecePrecedente.Detail[i].GetValue('GL_VIVANTE') = '-');
          end;
          if AutresLignesMortes then
          begin
            { Mise à jour de la ligne }
            Sql := 'UPDATE LIGNE SET GL_VIVANTE="-" WHERE ' + WherePiece(CleDocLignePrecedente, ttdLigne, True, True);
            Result := (ExecuteSQL(Sql) = 1);
            if Result then
            begin
              { Mise à jour de la pièce }
              Sql := 'UPDATE PIECE SET GP_VIVANTE="-" WHERE ' + WherePiece(CleDocLignePrecedente, ttdPiece, False);
              Result := ExecuteSQL(Sql) = 1;
            end;
          end;
        end
        else { La ligne renait de ses cendres }
        begin
          { Vérifie si toutes les lignes sont vivantes }
          i := -1;
          AutresLignesVivantes := True;
          while (i < TobPiecePrecedente.Detail.Count - 1) and AutresLignesVivantes do
          begin
            Inc(i);
            if not LigneIdentique(TobPiecePrecedente.Detail[i]) then
              AutresLignesVivantes := (TobPiecePrecedente.Detail[i].GetValue('GL_VIVANTE') = 'X');
          end;
          {
              Mise à jour des lignes :
              . Si toutes les lignes sont vivantes on update uniquement la ligne qui vient de changer d'état
              . Si toutes les lignes ne sont pas vivantes, on update toutes les lignes de la pièce
          }
          Sql := 'UPDATE LIGNE SET GL_VIVANTE="X" WHERE ' + WherePiece(CleDocLignePrecedente, ttdLigne, AutresLignesVivantes, True);
          Result := (ExecuteSQL(Sql) > 0);
          if Result then
          begin
            { Mise à jour de la pièce }
            Sql := 'UPDATE PIECE SET GP_VIVANTE="X" WHERE ' + WherePiece(CleDocLignePrecedente, ttdPiece, False);
            Result := (ExecuteSQL(Sql) = 1);
          end;
        end;
      end;
    end;
  finally
    TobPiecePrecedente.Free;
    TobLignePrecedente.Free;
  end;
  if not Result and V_PGI.Sav then
    PGIError(TraduireMemoire('La requête') + ' : ' + Sql + ' ' + TraduireMemoire('a échoué')
      , TraduireMemoire('Mise à jour de l''état de la pièce origine'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 01/04/2003
Description .. : Remise à 0 du reste à livrer dans la pièce d'origine
*****************************************************************}

procedure RazResteALivrer(TobPiece: Tob); { NEWPIECE }
var
  i: Integer;
begin

  if Assigned(TobPiece) then
  begin
    for i := 0 to TobPiece.Detail.Count - 1 do
    Begin
      TobPiece.Detail[i].PutValue('GL_QTERESTE', 0);
      // --- GUINIER ---
      TobPiece.Detail[i].PutValue('GL_MTRESTE', 0);
    end;
  end;

end;

function EstLigneSoldee(TobL: Tob; Fromvalidation : boolean=false): Boolean;
var
  TobPiece, TobDim: Tob;
  i, iTypeDim, iCodeArticle, iQteReste, iMtReste : Integer;
  CodeArticleGen: String;
  TouteLigneSoldee: Boolean;

  function _EstLigneSoldee(TobL: Tob; Fromvalidation : boolean): Boolean;
  begin
    // --- GUINIER ---
    if CtrlOkReliquat(TobL, 'GL') then
    begin
      if Fromvalidation then
        Result := (TobL.GetValue('GL_MTRESTE') = 0) or (TobL.GetValue('GL_SOLDERELIQUAT') = 'X')
      else
        Result := ((TobL.GetValue('GL_MTRESTE') = 0) and (TOBL.GetValue('GL_MTRELIQUAT') <> 0)) or (TobL.GetValue('GL_SOLDERELIQUAT') = 'X');
    end
    else
    begin
      if Fromvalidation then
        Result := (TobL.GetValue('GL_QTERESTE') = 0) or (TobL.GetValue('GL_SOLDERELIQUAT') = 'X')
      else
        Result := ((TobL.GetValue('GL_QTERESTE') = 0) and (TOBL.GetValue('GL_QTEFACT') <> 0)) or (TobL.GetValue('GL_SOLDERELIQUAT') = 'X');
    end;
  end;

begin
  Result := False;
  {$IFDEF CHR} // pas de gestion de ligne barrée actuellement pour CHR
  exit;
  {$ENDIF}
  if Assigned(TobL) then
  begin
    if Tobl.GetValue('GL_TYPEDIM') <> 'GEN' then
      Result := _EstLigneSoldee(TobL,FromValidation)   { Ligne Article ou DIM }
    else
    begin
      TobPiece := TobL.Parent;
      if Assigned(TobPiece) then
      begin
        Result := True;
        i := TobL.GetIndex + 1;
        if i <= TobPiece.Detail.Count - 1 then
        begin
          TobDim := TobPiece.Detail[i];
          Result := True;
          CodeArticleGen := TobDim.GetValue('GL_CODEARTICLE');
          iTypeDim := TobDim.GetNumChamp('GL_TYPEDIM');
          iCodeArticle := TobDim.GetNumChamp('GL_CODEARTICLE');
          iQteReste := TobDim.GetNumChamp('GL_QTERESTE');
          // --- GUINIIER ---
          if CtrlOkReliquat(TobDim, 'GL') then iMtReste := TobDim.GetNumChamp('GL_MTRESTE');
          while (TobDim.GetValeur(iTypeDim) = 'DIM') and (TobDim.GetValeur(iCodeArticle) = CodeArticleGen) and Result do
          begin
            Result := _EstLigneSoldee(TobDim,FromValidation);
            if Result then
            begin
              Inc(i);
              if i > TOBPiece.Detail.Count-1 then
                Break
              else
                TobDim := TobPiece.Detail[i];
            end;
          end;
        end;
      end;
    end;
  end;

end;

function WithPiecePrecedente(TobL: Tob): Boolean; { NEWPIECE }
begin
  if Assigned(TobL) then
    Result := (TobL.GetValue('GL_PIECEPRECEDENTE') <> '')
  else
    Result := False;
end;

function CalculResteALivrer(TOBL: Tob; AddTobPiece: Tob = nil): Double;
var
  i: Integer;
  TOBLSuiv: Tob;
  PiecePrecedente: String;
begin
  Result := 0;
  if Assigned(TOBL) then
  begin
    TOBLSuiv := Tob.Create('_LIGNESUIV_', nil, -1);
    try
      if GetTobsSuiv(TOBL, TOBLSuiv) then
      begin
        for i := 0 to TOBLSuiv.Detail.Count - 1 do
        begin
          if EstLigneArticle(TOBLSuiv.Detail[i]) then
          begin
            // --- GUINIER ---
            if CtrlOkReliquat(TOBLSuiv.Detail[i], 'GL') Then
              Result := Result + TOBLSuiv.Detail[i].GetValue('GL_MONTANTHTDEV')
            else
              Result := Result + TOBLSuiv.Detail[i].GetValue('GL_QTESTOCK');
          end;
        end;
      end;
    finally
      TOBLSuiv.Free;
    end;
    if Assigned(AddTobPiece) then
    begin
      PiecePrecedente := EncodeRefPiece(TOBL);
      for i := 0 to AddTobPiece.Detail.Count - 1 do
      begin
        if EstLigneArticle(AddTobPiece.Detail[i]) then
        begin
          if PiecePrecedente = AddTobPiece.Detail[i].GetValue('GL_PIECEPRECEDENTE') then
          begin
            // --- GUINIER ---
            if CtrlOkReliquat(TOBLSuiv.Detail[i], 'GL') Then
              Result := Result + AddTobPiece.Detail[i].GetValue('GL_MONTANTHTDEV')
            else
              Result := Result + AddTobPiece.Detail[i].GetValue('GL_QTESTOCK');
          end;
        end;
      end;
    end;
  end;

end;

// --------------------

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 18/07/2003
Modifié le ... : 18/07/2003
Description .. : Construit le SELECT de la requete de récupération des
Suite ........ : lignes et des compléments de ligne
Mots clefs ... : LIGNE;LIGNECOMPL
*****************************************************************}
function MakeSelectLigne: String;
begin
  Result := 'SELECT LIGNE.*,LIGNECOMPL.* '
          +  'FROM LIGNE '
          +  'LEFT JOIN LIGNECOMPL '
          +  'ON (GL_NATUREPIECEG = GLC_NATUREPIECEG and GL_SOUCHE = GLC_SOUCHE and GL_NUMERO = GLC_NUMERO '
                                                  + 'and GL_INDICEG = GLC_INDICEG and GL_NUMORDRE = GLC_NUMORDRE)'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 18/07/2003
Description .. : Copie les champs clés d'une TobLigne dans les champs
Suite ........ : clés d'une TobLigneCompl
Mots clefs ... : LIGNE;LIGNECOMPL
*****************************************************************}
Procedure CleLigneToCleLigneCompl(TobL, TobLigneCompl: Tob);
begin
  if Assigned(Tobl) and Assigned(TobLigneCompl) then
  begin
    TobLigneCompl.PutValue('GLC_NATUREPIECEG', TobL.GetValue('GL_NATUREPIECEG'));
    TobLigneCompl.PutValue('GLC_SOUCHE', TobL.GetValue('GL_SOUCHE'));
    TobLigneCompl.PutValue('GLC_NUMERO', TobL.GetValue('GL_NUMERO'));
    TobLigneCompl.PutValue('GLC_INDICEG', TobL.GetValue('GL_INDICEG'));
    TobLigneCompl.PutValue('GLC_NUMORDRE', TobL.GetValue('GL_NUMORDRE'));
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 18/07/2003
Description .. : Contrôle si un nom de champ appartient à la table
Suite ........ : LIGNECOMPL
*****************************************************************}
function EstChampDeLigneCompl(CodeChamp: String): boolean; 
begin
  if ExtractPrefixe(CodeChamp) = 'GLC' then      { Table LIGNECOMPL }
    Result := ChampToNum(CodeChamp) <> -1        { Champ existant }
  else
    Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 18/07/2003
Description .. : Test si un nom de champ fait partie de la clef1 de la table
*****************************************************************}
function ChampDeLaCle1(NomChamp: String): Boolean; 
var
	iTable: Integer;
  Clef1: String;
	Mcd : IMCDServiceCOM;
	Table     : ITableCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Result := False;
  iTable := TableToNum(PrefixeToTable(ExtractPrefixe(NomChamp)));
  if iTable > 0 then
  begin
    Clef1 := Mcd.TableToCle1(Mcd.GetTable(PrefixeToTable(ExtractPrefixe(NomChamp))).name);
    NomChamp := Trim(UpperCase(Nomchamp));
    Result := Pos(NomChamp, Clef1) <> 0;
  end;
end;

function EstVide(NomChamp: String; TheTob: Tob): Boolean; 
var
  s: String;
begin
  Result := False;
  if Assigned(TheTob) then
  begin
    { Valeur nulle dans la Tob }
    Result := (TheTob.G(NomChamp) <> null);
    if Result then
    begin
      s := TypeDeChamp(NomChamp);
      if s = 'C' then
        Result := TheTob.GetValue(NomChamp) = ''
      else if (s = 'I') or (s = 'N') then
        Result := TheTob.GetValue(NomChamp) = 0
      else if (s = 'B')  then
        Result := TheTob.GetValue(NomChamp) = '-'
      else
        Result := False;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 21/07/2003
Description .. : Contrôle si il faut insérer une TobLigne dans LIGNECOMPL
Mots clefs ... : LIGNE;LIGNECOMPL
*****************************************************************}
function GLCNewLigne(TobL: Tob): Boolean; 
var
  iChampSup: Integer;
  NomCHamp: String;
begin
  Result := False;
  if Assigned(TobL) then
  begin
    { Boucle sur les champs sups de la pièce }
    Result := False; iChampSup := 999;
    while (iChampSup < 1000 + TobL.ChampsSup.Count - 1) and not Result do
    begin
      Inc(iChampSup);
      NomChamp := TobL.GetNomChamp(iChampSup);
      { Champ GLC_xxx ne faisant pas partie de la clef1 }
      Result := EstChampDeLigneCompl(NomChamp) and (not ChampDeLaCle1(NomChamp)) and (not EstVide(NomChamp, Tobl));
    end;
  end;
end;

function GLCModifLigne(TobLNew, TobLOld: Tob): Boolean; 
var
  iChampSup: Integer;
  NomCHamp: String;
begin
  Result := False;
  if Assigned(TobLNew) and Assigned(ToblOld) then
  begin
    { Boucle sur les champs sups de la ligne de pièce }
    Result := False; iChampSup := 999;
    while (iChampSup < 1000 + TobLNew.ChampsSup.Count - 1) and not Result do
    begin
      Inc(iChampSup);
      NomChamp := TobLNew.GetNomChamp(iChampSup);
      { Champ GLC_xxx ne faisant pas partie de la clef1 }
      if EstChampDeLigneCompl(NomChamp) and (not ChampDeLaCle1(NomChamp)) and (not EstVide(NomChamp, TobLNew)) then
        Result := ToblNew.GetValue(NomChamp) <> ToblOld.GetValue(NomChamp);
    end;
  end;
end;

procedure GLCAddTob(TobL, TobLignesCompl: Tob);
  function ClefVide (TOBL : TOB) : boolean;
  begin
    result := (TOBL.getValue('GLC_NATUREPIECEG')='') and (TOBL.getValue('GLC_NUMERO')=0);
  end;
var
  TobLigneCompl: Tob;
  iChamp: Integer;
  NomCHamp: String;
begin
  if Assigned(Tobl) and Assigned(TobLignesCompl) then
  begin
    TobLigneCompl := Tob.Create('LIGNECOMPL', TobLignesCompl, -1);
    { Copie les valeurs de la clé }
    CleLigneToCleLigneCompl(TobL, TobLigneCompl);
    for iChamp := 1000 to 1000 + TobL.ChampsSup.Count - 1 do
    begin
      NomChamp := TobL.GetNomChamp(iChamp);
      { Champ GLC_xxx ne faisant pas partie de la clef1 }
      if EstChampDeLigneCompl(NomChamp) and (not ChampDeLaCle1(NomChamp)) then
      begin
      	if (not VarIsNull(TOBL.GetValue(Nomchamp))) And (VarAsType(TOBL.getValue(nomchamp), varString) <> #0) then
        begin
        	TRY
        	TobLigneCompl.PutValue(NomChamp, Tobl.GetValue(NomChamp));
          EXCEPT
          	raise exception.create ('Le champ '+Nomchamp+' n''est pas de type attendu');
          END;
        end;
      end;
    end;
    TobLigneCompl.SetAllModifie(True);
  end;
  if ClefVide(TOBLigneCompl) then TobLigneCompl.free;
end;

{$IFDEF STOPAUC....RIE}
procedure ValideLesLignesCompl(TobPiece, TobPiece_O: Tob);
var
  iLigne: Integer;
  TobLignesComplInsertUpdate, TobLignesComplDelete, TobLNew, ToblOld, TobModifMere, TobModifFille: Tob;
  CleDocCompl: R_CleDoc;

  function GetLigne(TobMod, Where: Tob): Tob;
  var
    CleDoc: R_CleDoc;
  begin
    Result := nil;
    if Assigned(Where) then
    begin
      CleDoc := TOB2CleDoc(TobMod);
      CleDoc.NumLigne := TobMod.GetValue('GL_NUMLIGNE');
      CleDoc.NumOrdre := TobMod.GetValue('GL_NUMORDRE');
      Result := FindCleDocInTobPiece(CleDoc, Where);
    end;
  end;

begin
  if Assigned(TobPiece) then
  begin
    TobLignesComplInsertUpdate := Tob.Create('_LIGNECOMPL_', nil, -1);
    TobLignesComplDelete := Tob.Create('_LIGNECOMPL_', nil, -1);
    try
      TobModifMere := Tob.Create('_MODIF_', nil, -1);
      try
        { Comparaison entre les lignes de l'ancienne et de la nouvelle piece }
        CompareTobPiece(TobPiece, TobPiece_O, TobModifMere, false);
        { Vu que l'on supprime la pièce initiale (DelOrUpdateAncien), il faut recrée les lignes dans LIGNECOMPL }
        iLigne := -1;
        while (iLigne < TobModifMere.Detail.Count - 1) do
        begin
          Inc(iLigne);
          TobModifFille := TobModifMere.Detail[iLigne];
          if TobModifFille.G('OLD_EXISTE') = 'X' then
          begin
            TobLOld := GetLigne(TobModifFille, TobPiece_O);
            if Assigned(TobLOld) then
            begin
              CleDocCompl := TOB2CleDoc(ToblOld); CleDocCompl.NumOrdre := TobLOld.G('GL_NUMORDRE');
              if not ExisteLigneCompl(CleDocCompl) then
                GLCAddTob(TobLOld, TobLignesComplInsertUpdate);
            end;
          end;
        end;
        if TobLignesComplInsertUpdate.Detail.Count > 0 then
        begin
          TobLignesComplInsertUpdate.InsertOrUpdateDB;
          TobLignesComplInsertUpdate.ClearDetail;
        end;
        { Compare l'ancienne et la nouvelle piece }
        iLigne := -1;
        while (iLigne < TobModifMere.Detail.Count - 1) do
        begin
          Inc(iLigne);
          TobModifFille := TobModifMere.Detail[iLigne];
          if TobModifFille.G('OLD_EXISTE') = 'X' then
          begin
            { Vu que l'on supprime la pièce initiale (DelOrUpdateAncien), il faut recrée les lignes dans LIGNECOMPL }
            TobLOld := GetLigne(TobModifFille, TobPiece_O);
            if Assigned(TobLOld) then
            begin
              CleDocCompl := TOB2CleDoc(ToblOld); CleDocCompl.NumOrdre := TobLOld.G('GL_NUMORDRE');
              if not ExisteLigneCompl(CleDocCompl) then
              begin

              end;
            end;
          end;
          if (TobModifFille.G('NEW_EXISTE') = 'X') and (TobModifFille.G('OLD_EXISTE') = '-') then
          begin
            { Nouvelle ligne }
            TobLNew := GetLigne(TobModifFille, TobPiece);
            { Faut-il l'insérer dans la table LIGNECOMPL }
            if GLCNewLigne(TobLNew) then
              GLCAddTob(TobLNew, TobLignesComplInsertUpdate);
          end
          else if (TobModifFille.G('NEW_EXISTE') = 'X') and (TobModifFille.G('OLD_EXISTE') = 'X') then
          begin
            { Ligne modifiée }
            TobLNew := GetLigne(TobModifFille, TobPiece);
            TobLOld := GetLigne(TobModifFille, TobPiece_O);
            { Des champs Sups de LIGNECOMPL ont il été modifié ? }
            if GLCModifLigne(TobLNew, TobLOld) then
              GLCAddTob(TobLNew, TobLignesComplInsertUpdate);
          end
          else if (TobModifFille.G('NEW_EXISTE') = '-') and (TobModifFille.G('OLD_EXISTE') = 'X') then
          begin
            { Ligne supprimée }
            TobLOld := GetLigne(TobModifFille, TobPiece_O);
            { Mise à jour de la tob des lignes supprimées }
            GLCAddTob(TobLOld, TobLignesComplDelete);
          end;
        end;
      finally
        TobModifMere.Free;
      end;
      { Ligne nouvelle et modifiée }
      if TobLignesComplInsertUpdate.Detail.Count > 0 then
        TobLignesComplInsertUpdate.InsertOrUpdateDB;
      { Ligne supprimée }
      if TobLignesComplDelete.Detail.Count > 0 then
        TobLignesComplDelete.DeleteDB;
    finally
      TobLignesComplInsertUpdate.Free;
      TobLignesComplDelete.Free;
    end;
  end;
end;
{$ENDIF}

procedure ValideLesLignesCompl(TobPiece, TobPiece_O: Tob);
var
  indice: Integer;
  TobLignesCompl,TOBL: Tob;
  okok : boolean;
begin
  if Assigned(TobPiece) then
  begin
    TobLignesCompl := Tob.Create('LES LIGNECOMPL', nil, -1);
    TRY
      for Indice := 0 to TOBPiece.detail.count -1 do
      begin
      	TOBL := TOBPIece.detail[Indice];
        GLCAddTob(TobL, TobLignesCompl);
      end;
      okok := false;
      TRY
        okok := TobLignesCompl.InsertDB (nil,true);
      EXCEPT
        on E: Exception do
      begin
          PgiError('Erreur SQL : ' + E.Message, 'Erreur écriture LIGNECOMPL');
        end;
      END;
      if not okok then
      begin
        V_PGI.Ioerror := OeUnknown;
      end;
    finally
      TobLignesCompl.Free;
    end;
  end;
end;


procedure UpdTobEvtMat(TobL, TobEVTMAT: Tob);
var
  TT: Tob;
  NumEvt : Integer;
  NbMinutes : integer;
  DateD,DateF : TDateTime;
begin
  if Assigned(Tobl) and Assigned(TobEVTMAT) then
  begin
    NbMinutes := ceil(TobL.getDouble('GLC_QTEAPLANIF')*60);
    DateD := StrToDate(DateToStr(TOBL.GetDateTime('GL_DATEPIECE')));
    DateF := IncMinute(DateD,NbMinutes); DateF := StrToDate(DateToStr(DateF));
    TT := TobEVTMAT.FindFirst(['BEM_IDEVENTMAT'],[TobL.GetInteger('GLC_IDEVENTMAT')],true);
    if TT <> NIL then
    begin
      TT.SetString('BEM_CODEMATERIEL',TobL.GetString('GLC_CODEMATERIEL'));
//      TT.SetDateTime('BEM_DATEDEB',DateD);
//      TT.SetDateTime('BEM_DATEFIN',DateF);
      TT.SetString('BEM_BTETAT',TobL.GetString('GLC_BTETAT'));
      if TobL.GetString('GL_AFFAIRE') <> '' Then
        TT.SetString('BEM_AFFAIRE',TobL.GetString('GL_AFFAIRE'));
      TT.SetDouble('BEM_PA',TobL.GetDouble('GL_PUHTNET'));
      TT.SetDouble('BEM_PR',TobL.GetDouble('GL_PUHTNET'));
      TT.SetDouble('BEM_PV',TobL.GetDouble('GL_PUHTNET'));
      TT.SetString('BEM_DEVISE',V_PGI.DevisePivot);
      TT.SetDouble('BEM_NBHEURE',TOBL.GetDouble('GLC_QTEAPLANIF'));
      if TT.getString('SUPPR')<>'X' then
      begin
        TT.SetString('BEM_CODEETAT','REA');     // on saisie des factures ..donc..
        TT.SetString('BEM_REFPIECE',EncodeRefPiece(TobL));
        TT.SetString('BEM_LIBREALISE',TobL.GetString('GL_LIBELLE'));
      end else
      begin
        TT.SetString('BEM_CODEETAT','ARE');
        TT.SetString('BEM_REFPIECE','');
        TT.SetString('BEM_LIBREALISE','');
      end;
      TT.SetString('UPDATED','X');
      TT.SetAllModifie(True);
    end;
  end;
end;

(*
procedure AddTobEvtMat(TobL, TobEVTMAT: Tob);
var
  TT: Tob;
  NumEvt : Integer;
  NbMinutes : integer;
  DateD,DateF : TDateTime;
begin
  if Assigned(Tobl) and Assigned(TobEVTMAT) then
  begin
    if not GetNumCompteur('BEM',iDate1900, NumEvt) then  V_PGI.IOError := oeUnknown;
    NbMinutes := ceil(TobL.getDouble('GLC_QTEAPLANIF')*60);
    DateD := StrToDate(DateToStr(TOBL.GetDateTime('GL_DATEPIECE')));
    DateF := IncMinute(DateD,NbMinutes); DateF := StrToDate(DateToStr(DateF));
    TT := Tob.Create('BTEVENTMAT', TobEVTMAT, -1);
    TT.AddChampSupValeur ('UPDATED','X');
    TT.SetInteger('BEM_IDEVENTMAT',NumEvt);
    TT.SetString('BEM_REFPIECE',EncodeRefPiece(TobL));
    TT.SetString('BEM_CODEMATERIEL',TobL.GetString('GLC_CODEMATERIEL'));
    TT.SetDateTime('BEM_DATEDEB',DateD);
    TT.SetDateTime('BEM_DATEFIN',DateF);
    TT.SetString('BEM_BTETAT',TobL.GetString('GLC_BTETAT'));
    TT.SetString('BEM_LIBACTION',TobL.GetString('GL_LIBELLE'));
    if TobL.GetString('GL_AFFAIRE') <> '' Then
      TT.SetString('BEM_AFFAIRE',TobL.GetString('GL_AFFAIRE'));
    TT.SetDouble('BEM_PA',TobL.GetDouble('GL_PUHTNET'));
    TT.SetDouble('BEM_PR',TobL.GetDouble('GL_PUHTNET'));
    TT.SetDouble('BEM_PV',TobL.GetDouble('GL_PUHTNET'));
    TT.SetString('BEM_DEVISE',V_PGI.DevisePivot);
    TT.SetDouble('BEM_NBHEURE',TOBL.GetDouble('GLC_QTEAPLANIF'));
    TT.SetString('BEM_RESSOURCE','');
    TT.SetString('BEM_CODEETAT','REA');     // on saisie des factures ..donc..
    TT.SetString('BEM_LIBREALISE',TobL.GetString('GL_LIBELLE'));
    // ---
    TOBL.SetInteger('GLC_IDEVENTMAT',NumEvt);
    TT.SetAllModifie(True);
  end;
end;
*)

procedure LoadlesEvtsmat (TOBpiece,TOBEvtsmat : TOB);
var SQl,RefA : string;
    QQ : TQuery;
begin
  RefA := EncodeRefpieceSup(TOBPiece);
  SQl := 'SELECT *,"X" AS UPDATED,"-" AS SUPPR FROM BTEVENTMAT WHERE BEM_REFPIECE LIKE "'+Refa+'"';
  QQ := OpenSQL(Sql,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBEvtsmat.LoadDetailDB('BTEVENTMAT','','',QQ,false);
  end;
  ferme (QQ);
end;

procedure EnleveEvtsMateriels (TOBL,TOBEvtsMat : TOB);
var TT : TOB;
begin
  TT := TOBEvtsmat.FindFirst(['BEM_IDEVENTMAT'],[TOBL.GetInteger('GLC_IDEVENTMAT')],true);
  if TT <> nil then
  begin
    TT.SetString('UPDATED','X');
    TT.SetString('SUPPR','X');
  end;
end;


procedure SetInfosFromEvts  (TOBL,TOBEvtsmat : TOB);
var TT : TOB;
begin
  TT := TOBEvtsmat.FindFirst(['BEM_IDEVENTMAT'],[TOBL.GetInteger('GLC_IDEVENTMAT')],true);
  if TT <> nil then
  begin
    TOBL.SetDouble('GLC_QTEAPLANIF',TT.GetDouble('BEM_NBHEURE'));
    TOBL.SetString('GL_AFFAIRE',TT.GetString('BEM_AFFAIRE'));
    TOBL.SetString('EVTMATPRESENT','X');
    DecoupeCodeAffsurTOB (TOBL);
  end;
end;

procedure RecupEvtsMateriels (TOBL,TOBEvtsmat : TOB);
var QQ : TQuery;
    Sql : string;

begin
  SQl := 'SELECT *,"-" AS UPDATED,"-" AS SUPPR FROM BTEVENTMAT WHERE BEM_IDEVENTMAT='+TOBL.GetString('GLC_IDEVENTMAT');
  QQ := OpenSQL(Sql,True,-1,'',true);
  if not QQ.eof then
  begin
    TOBEvtsmat.LoadDetailDB('BTEVENTMAT','','',QQ,true);
  end;
  ferme (QQ);
end;

procedure DeleteEvtsMateriels (TOBL,TOBEvtsMat : TOB);
var TT : TOB;
begin
  TT := TOBEvtsmat.FindFirst(['BEM_IDEVENTMAT'],[TOBL.GetInteger('GLC_IDEVENTMAT')],true);
  if TT <> nil then
  begin
    TT.SetString('UPDATED','X');
    TT.SetString('SUPPR','X');
  end;
end;

procedure ValideLesEvtsMateriels(TOBPiece,TOBEvtsmat : TOB);
var
  indice: Integer;
  TOBL,TobNEVTMAT: Tob;
  okok : boolean;
begin
  if Assigned(TobPiece) then
  begin
    TobNEVTMAT := TOB.Create ('LES EVENTS',nil,-1);
    for Indice := 0 to TOBPiece.detail.count -1 do
    begin
      TOBL := TOBPIece.detail[Indice];
      if TOBL.GetString('GLC_CODEMATERIEL')='' then Continue;
      (*
      if (TOBL.GetInteger('GLC_IDEVENTMAT')=0) or (
          TOBL.GetString('EVTMATPRESENT')='-')then
          AddTobEvtMat(TobL,TobNEVTMAT)
                                              else UpdTobEvtMat(TOBL,TOBEvtsmat);
      *)
      if (TOBL.GetInteger('GLC_IDEVENTMAT')<>0) {and (TOBL.GetString('EVTMATPRESENT')<>'-')} then
        UpdTobEvtMat(TOBL,TOBEvtsmat);
    end;
    // nettoyage
    indice := 0;
    if TOBEvtsmat.detail.count > 0 then
    begin
      repeat
        if TOBEvtsmat.detail[indice].GetString('UPDATED')<>'X' then
        begin
          TOBEvtsmat.detail[indice].free;
        end else inc(indice);
      until indice >= TOBEvtsmat.Detail.Count;
    end;
    // --
    (*
    if not TobNEVTMAT.InsertDb (nil,true) then
    begin
      MessageValid := 'Erreur écriture BTEVENTMAT (INS)';
      V_PGI.Ioerror := OeUnknown;
      Exit;
    end;
    *)
    okok := false;
    TRY
      okok := TOBEvtsmat.UpdateDB (true);
    EXCEPT
      on E: Exception do
    begin
        PgiError('Erreur SQL : ' + E.Message, 'Ecriture Evts matériel');
      end;
    END;
    if not okok then
    begin
      V_PGI.Ioerror := OeUnknown;
      Exit;
    end;
    TobNEVTMAT.free;

  end;

end;

function ExisteLigneCompl(CleDocLigne: R_CleDoc): Boolean;
begin
  Result := ExisteSQL('SELECT GLC_NATUREPIECEG FROM LIGNECOMPL WHERE ' + WherePiece(CleDocLigne, ttdLigneCompl, True));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 06/11/2003
Description .. : Lors de la validation d'une pièce en modification gère le
Suite ........ : solde ou le dé-solde des lignes.
*****************************************************************}
procedure GereLesLignesSoldees(TobPiece, TobPiece_O, TobArticles, TobCatalogu, TOBOuvrage_O: Tob);
var
  i, IndiceNomen: Integer;
  TOBL, TOBL_O, TobOuvrage: Tob;
  Find, Inverse: Boolean;
  NewEtat, NaturePieceG, ColPlus, ColMoins, StPrixValo: string;
  OldQteStock, OldQteFact, Qte, RatioVA: Double;
  Ok_ReliquatMT : Boolean;
  MtReste       : Double;
  OldMontant    : Double;
begin
  Find := False;
  for i := 0 to TobPiece.Detail.Count - 1 do
  begin
    TOBL := TobPiece.Detail[i];
    TOBL_O := FindTobLigInOldTobPiece(TOBL, TobPiece_O);
    // --- GUINIER ---
    Ok_ReliquatMT := CtrlOkReliquat(TOBL, 'GL');
    if Assigned(TOBL_O) and (TOBL.GetValue('GL_SOLDERELIQUAT') <> TOBL_O.GetValue('GL_SOLDERELIQUAT')) then
    begin
      Find := True;
      // --- GUINIER ---
      Qte := TOBL.GetValue('GL_QTERESTE');
      if Ok_ReliquatMT then MtReste := TOBL.GetValue('GL_QTERESTE');
      //
      if TOBL.GetValue('GL_SOLDERELIQUAT') = 'X' then
      begin
        { On vient de solder la ligne }
        Inverse := False;
        Qte := TOBL.GetValue('GL_QTERELIQUAT');
        if Ok_ReliquatMT then MtReste := TOBL.GetValue('GL_MTRELIQUAT');
      end
      else
      begin
        { On vient de dé-solder la ligne }
        Inverse := True;
        // --- GUINIER ---
        Qte := TOBL.GetValue('GL_QTERESTE');
        if Ok_ReliquatMT then MtReste := TOBL.GetValue('GL_MTRESTE');
      end;
      NaturePieceG  := TobPiece.GetValue('GP_NATUREPIECEG');
      ColMoins      := GetInfoParPiece(NaturePieceG, 'GPP_QTEMOINS');
      ColPlus       := GetInfoParPiece(NaturePieceG, 'GPP_QTEPLUS');
      StPrixValo    := GetInfoParPiece(NaturePieceG, 'GPP_MAJPRIXVALO');
      { Sauve la quantité }
      OldQteStock := TOBL.GetValue('GL_QTESTOCK');
      OldQteFact  := TOBL.GetValue('GL_QTEFACT');
      OldMontant  := TOBL.GetValue('GL_MONTANTHTDEV');
      try
        TOBL.PutValue('GL_QTEFACT',  Qte);
        TOBL.PutValue('GL_QTESTOCK', Qte);
        // --- GUINIER ---
        if Ok_ReliquatMT then TOBL.PutValue('GL_MONTANTHTDEV', MtReste);
        //
        RatioVA := GetRatio(TOBL, nil, trsStock);
        IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
        if (IndiceNomen > 0) and Assigned(TOBOuvrage_O) and (TOBOuvrage_O.detail.count >0)  then
          TOBOuvrage := TOBOuvrage_O.Detail[IndiceNomen - 1]
        else
          TOBOuvrage := nil;
        { Ajout de la TobNomen dans les paramètres }
        MajQteStock(TOBL, TobArticles, TOBOuvrage, ColPlus, ColMoins, Inverse, RatioVA);
      finally
        // --- GUINIER ---
        TOBL.PutValue('GL_QTESTOCK', OldQteStock);
        TOBL.PutValue('GL_QTEFACT',  OldQteFact);
        if Ok_ReliquatMT then TOBL.PutValue('GL_MONTANTHTDEV', OldMontant);
      end;
    end;
  end;

  { Mise à jour de l'état de la pièce }
  if Find then
  begin
    NewEtat := '';
    if ToutesLignesSoldees(TobPiece) then
      NewEtat := '-'
    else
      NewEtat := 'X';
    TobPiece.PutValue('GP_VIVANTE', NewEtat);
    { Etat des lignes = Etat de l'entête }
    for i := 0 to TobPiece_O.Detail.Count - 1 do
      TobPiece.detail[i].PutValue('GL_VIVANTE', NewEtat);
  end;

end;
{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 06/11/2003
Description .. : Lors du chargement d'une pièce en modification, ajoute les
Suite ........ : champs SOLDERELIQUAT et QTERELIQUAT de la pièce
Suite ........ : précédente dans la pièce courante, pour éviter
Suite ........ : des requetes durant la saisie d'uen pièce
*****************************************************************}
procedure AjouteEtatLignePrecedente(TOBPiece, TOBPiece_O: TOB; TransfoPiece: Boolean; Action: TActionFiche);
var
  TOBL, TOBL_O: TOB;
  i: Integer;
begin
  if Action = taCreat then
    Exit;
  for i := 0 to TobPiece.Detail.Count - 1 do
  begin
    TOBL := TobPiece.Detail[i];
    if EstLigneArticle(TOBL) and (TOBL.GetValue('GL_PIECEPRECEDENTE') <> '') then
    begin
      TOBL_O := nil;
      if TransfoPiece then
        TOBL_O := GetTOBPrec(TOBL, TOBPiece_O)
      else
        { Il faut envoyer une requete pour récupérer la ligne précédente }
        TOBL_O := GetTOBPrecBySQL(TOBL);
        // DBR 17112003 TOBL_O := GetTOBPrecBySQL(TOBL, 'GL_SOLDERELIQUAT,GL_QTERELIQUAT');
        TOBL_O := GetTOBPrecBySQL(TOBL, 'GL_SOLDERELIQUAT, GL_QTERELIQUAT, GL_QTERESTE, GL_MTRESTE, GL_MTRELIQUAT');
      if Assigned(TOBL_O) then
      begin
        TOBL.AddChampSupValeur('LPREC_SOLDERELIQUAT', TOBL_O.GetValue('GL_SOLDERELIQUAT'));
        TOBL.AddChampSupValeur('LPREC_QTEERELIQUAT',  TOBL_O.GetValue('GL_QTERELIQUAT'));
        TOBL.AddChampSupValeur('LPREC_QTERESTE',      TOBL_O.GetValue('GL_QTERESTE')); // DBR 17112003
        // --- GUINIER ---
        TOBL.AddChampSupValeur('LPREC_MTRELIQUAT',    TOBL_O.GetValue('GL_MTRELIQUAT'));
        TOBL.AddChampSupValeur('LPREC_MTRESTE',       TOBL_O.GetValue('GL_MTRESTE'));
      end;
    end;
  end;
end;

function GetTOBPrecBySQL(TOBL: TOB; Select : String = '*'): TOB;
var
  CleDoc: R_CLeDoc;
  Sql: String;
  Q: TQuery;
begin
  Result := Nil;
  if (TOBL.GetValue('GL_PIECEPRECEDENTE') <> '') then
  begin
    FillChar(CleDoc, Sizeof(CleDoc), #0);
    DecodeRefPiece(TOBL.GetValue('GL_PIECEPRECEDENTE'), CleDoc);
    Result := TOB.Create('LIGNE', nil, -1);
    Sql := 'SELECT ' + Select + ' FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, True, True);
    Q := OpenSQL(Sql, True,-1, '', True);
    try
      if not Q.Eof then
        Result.SelectDB('', Q)
      else
      begin
        Result.Free;
        Result := nil;
      end;
    finally
      Ferme(Q);
    end;
  end;
end;

function GetPrefixeTable (TOBL : TOB) : string;
begin
  result := '';
  if TOBL = nil then exit;
  if TOBL.nomtable = 'PIECE' then result := 'GP' else
  if TOBL.nomtable = 'LIGNE' then result := 'GL' else
  if TOBL.nomtable = 'LIGNEOUVPLAT' then result := 'BOP' else
  if TOBL.nomtable = 'LIGNEOUV' then result := 'BLO';
end;

function IsgestionAvanc (TOBpiece : TOB) : boolean;
begin
	result := (Pos(TOBpiece.getValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) and
  				  (Pos(RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')),'AVA;DAC')>0);
end;

function ExistRefBsv (TOBpiece : TOB) : boolean;
var RefGEscom : string;
begin
  if TOBpiece.getString('GP_BSVREF')<>'' then
  begin
    result := true
  end else
  begin
    RefGescom := EncodeRefPresqueCPGescom(TOBpiece);
    TRY
      result := ExisteSql('SELECT 1 FROM BASTENT WHERE BM4_REFGESCOM="'+RefGescom+'"');
    EXCEPT
      Result := false;
    END;
  end;
end;

function GetIDBSVDOC (TOBpiece : TOB) : string;
var RefGEscom : string;
    QQ : Tquery;
begin
  if TOBpiece.getString('GP_BSVREF')<>'' then
  begin
    result := TOBpiece.getString('GP_BSVREF')
  end else
  begin
    RefGescom := EncodeRefPresqueCPGescom(TOBpiece);
    TRY
      QQ := openSql('SELECT BM4_IDZEDOC FROM BASTENT WHERE BM4_REFGESCOM="'+RefGescom+'"',True,1,'',true);
      if not QQ.eof then
      begin
        Result := IntToStr(QQ.Fields[0].asInteger);
      end;
      ferme (QQ);
    EXCEPT
      Result := '';
    END;
  end;
end;

end.
