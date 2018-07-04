unit uEntCommun;

interface

uses
  HCtrls,
  Classes,
  sysutils,
  {$IFNDEF EAGLCLIENT}
       db,
       {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
       mul,
  {$else}
       eMul,
  {$ENDIF}
       UTob, 
       HEnt1;

Const CodeISOFR = 'FR' ; //XMG 14/07/03
      CodeISOES = 'ES' ; //XMG 14/07/03
      CodeISOBE = 'BE';  //SDA le 27/12/2007 - Belgique
      MaxSousPlan = 6 ;
      MaxAxe  = 5 ;

      ZONETOTAL = 'TOTALMOPA;'+
                  'TOTALMOPR;'+
                  'TOTALMOPV;'+
                  'TOTALFOUPA;'+
                  'TOTALINTPA;'+
                  'TOTALLOCPA;'+
                  'TOTALMATPA;'+
                  'TOTALOUTPA;'+
                  'TOTALSTPA;'+
                  'TOTALAUTPA;'+
                  'TOTALFOUPR;'+
                  'TOTALINTPR;'+
                  'TOTALLOCPR;'+
                  'TOTALMATPR;'+
                  'TOTALOUTPR;'+
                  'TOTALSTPR;'+
                  'TOTALAUTPR;'+
                  'TOTALFOUPV;'+
                  'TOTALINTPV;'+
                  'TOTALLOCPV;'+
                  'TOTALMATPV;'+
                  'TOTALOUTPV;'+
                  'TOTALSTPV;'+
                  'TOTALAUTPV';


      LISTCHPS = 'GL_NUMLIGNE|N;'+
                 'GL_TYPEARTICLE|N;'+
                 'GL_REFCATALOGUE|N;'+
                 'GL_REFARTSAISIE|O;'+
                 'GL_REFARTTIERS|O;'+
                 'GL_QTESAIS;'+
                 'GL_RENDEMENT;'+
                 'GL_PERTE;'+
                 'GL_LIBELLE;'+
                 'GL_QTEFACT;'+
                 'GL_QTESIT;'+
                 'GL_QTESTOCK|N;'+
                 'GL_QTERESTE|N;'+
                 'GL_QTERELIQUAT|N;'+
                 'GL_QTEPREVAVANC|N;'+
                 'GL_POURCENTAVANC|O;'+
                 'GL_TOTPREVAVANC|N;'+
                 'GL_TOTPREVDEVAVAN|N;'+
                 'GLC_VOIRDETAIL|N;'+
                 'GL_COEFMARG|O;'+
                 'GL_QUALIFQTEVTE|O;'+
                 'GL_QUALIFQTEACH|O;'+
                 'GLC_CODEMATERIEL|O;'+
                 'GLC_BTETAT|N;'+
                 'GL_FAMILLETAXE1;'+
                 'BLF_MTPRODUCTION;'+
                 'GL_AFFAIRE;'+
                 'GL_MTRESTE|N;'+
                 'GL_MONTANTPA|N;'+
                 'GL_DPA;'+
                 'GL_DPR|N;'+
                 'GL_PUHTDEV;'+
                 'GL_PUHTNETDEV;'+
                 'GL_PUTTCDEV;'+
                 'GL_REMISELIGNE;'+
                 'GL_DATELIVRAISON;'+
                 'GL_ETATSOLDE|N;'+
                 'GL_TPSUNITAIRE|N;'+
                 'GL_TOTALHEURE|N;'+
                 'GL_HEURE|N;'+
                 'GL_QUALIFHEURE;'+
                 'GL_REPRESENTANT;'+
                 'GL_DEPOT;'+
                 'GL_VALEURREMDEV|N;'+
                 'GL_MONTANTTTCDEV|N;'+
                 'GL_TOTALHTDEV|N;'+
                 'GL_TOTALTTCDEV|N;'+
                 'GL_MONTANTHTDEV|N;'+
                 'GL_CODEMARCHE|O;'+
                 'GLC_NUMEROTATION;'+
                 'GLC_GETCEDETAIL|N;'+
                 'GL_FOURNISSEUR|N;'+
                 'GL_CODECOND|O;'+
                 'GLC_NATURETRAVAIL|N;'+
                 'BLF_QTEMARCHE|N;'+
                 'BLF_QTEDEJAFACT|N;'+
                 'BLF_QTESITUATION;'+
                 'BLF_MTMARCHE|N;'+
                 'BLF_MTDEJAFACT|N;'+
                 'BLF_POURCENTAVANC;'+
                 'BLF_MTSITUATION;'+
                 'POURCENTMARG|O;'+
                 'POURCENTMARQ|O;'+
                 'GLC_MTMOPA|N;'+
                 'GLC_MTFOUPA|N;'+
                 'GLC_MTINTPA|N;'+
                 'GLC_MTLOCPA|N;'+
                 'GLC_MTMATPA|N;'+
                 'GLC_MTOUTPA|N;'+
                 'GLC_MTSTPA|N;'+
                 'GLC_MTAUTPA|N;'+
                 'GLC_MTMOPR|N;'+
                 'GLC_MTFOUPR|N;'+
                 'GLC_MTINTPR|N;'+
                 'GLC_MTLOCPR|N;'+
                 'GLC_MTMATPR|N;'+
                 'GLC_MTOUTPR|N;'+
                 'GLC_MTSTPR|N;'+
                 'GLC_MTAUTPR|N;'+
                 'GLC_MTMOPV|N;'+
                 'GLC_MTFOUPV|N;'+
                 'GLC_MTINTPV|N;'+
                 'GLC_MTLOCPV|N;'+
                 'GLC_MTMATPV|N;'+
                 'GLC_MTOUTPV|N;'+
                 'GLC_MTSTPV|N;'+
                 'GLC_MTAUTPV|N;'+
                 'GLC_QTECOND|O;'+
                 'GLC_COEFCOND|O;'+
                 'TOTALMOPA|N;'+
                 'TOTALMOPR|N;'+
                 'TOTALMOPV|N;'+
                 'TOTALFOUPA|N;'+
                 'TOTALINTPA|N;'+
                 'TOTALLOCPA|N;'+
                 'TOTALMATPA|N;'+
                 'TOTALOUTPA|N;'+                                
                 'TOTALSTPA|N;'+
                 'TOTALAUTPA|N;'+
                 'TOTALFOUPR|N;'+
                 'TOTALINTPR|N;'+
                 'TOTALLOCPR|N;'+
                 'TOTALMATPR|N;'+
                 'TOTALOUTPR|N;'+
                 'TOTALSTPR|N;'+
                 'TOTALAUTPR|N;'+
                 'TOTALFOUPV|N;'+
                 'TOTALINTPV|N;'+
                 'TOTALLOCPV|N;'+
                 'TOTALMATPV|N;'+
                 'TOTALOUTPV|N;'+
                 'TOTALSTPV|N;'+
                 'TOTALAUTPV|N';



const XmodeCreat = 1;
			XmodeMaj = 2;
			XmodeRemplace = 3;
Type
	TSousPlan = Record
              Code,Lib : String ;
              Debut,Longueur : Integer ;
              ListeSP : HTStringList ;
              End ;

  R_CLEDOC = RECORD
                NaturePiece : String3;
                Souche      : String3 ;
                DatePiece   : TDateTime ;
                NumeroPiece : Integer;
                NumLigne    : Integer;
                NumOrdre    : Integer ;
                Indice      : Integer;
                NoPersp     : integer ;
                UniqueBlo : integer;
             END ;
  {ne pas changer l'ordre des 5 premiers (axes)}
  TFichierBase = (fbAxe1,fbAxe2,fbAxe3,fbAxe4,fbAxe5,fbGene,fbAux,fbJal,fbBudGen,fbBudJal,
                     fbImmo,fbSect,fbCorresp,fbNatCpt,
                     fbBudSec1,fbBudSec2,fbBudSec3,fbBudSec4,fbBudSec5,
                     fbAxe1SP1,fbAxe1SP2,fbAxe1SP3,fbAxe1SP4,fbAxe1SP5,fbAxe1SP6,
                     fbAxe2SP1,fbAxe2SP2,fbAxe2SP3,fbAxe2SP4,fbAxe2SP5,fbAxe2SP6,
                     fbAxe3SP1,fbAxe3SP2,fbAxe3SP3,fbAxe3SP4,fbAxe3SP5,fbAxe3SP6,
                     fbAxe4SP1,fbAxe4SP2,fbAxe4SP3,fbAxe4SP4,fbAxe4SP5,fbAxe4SP6,
                     fbAxe5SP1,fbAxe5SP2,fbAxe5SP3,fbAxe5SP4,fbAxe5SP5,fbAxe5SP6,
                     fbRubrique,fbNone
                     ) ;
	TSousPlanAxe = Array[fbAxe1..fbAxe5,1..MaxSousPlan] Of TSousPlan ;

	TsetFichierBase = Set of TFichierBase ;

  TClesDocs = Array of R_CleDoc;
  T_TableDoc = (ttdPiece, ttdLigne, ttdLigneTarif, ttdPiedBase, ttdEche,
  						 ttdNomen, ttdLot, ttdAcompte, ttdPorc, ttdOuvrage, ttdLienOle,
               ttdPieceAdr, ttdSerie, ttdretenuG, ttdBaseRg, ttdSit, ttdParDoc,
               ttdAdresse, ttdRevision, ttdVariable, ttdLigneCompl, TtdLignePhase,
               TTdRepartmill,ttdLigneBase,ttdOuvrageP,ttdLignefac,ttdVarDoc,ttdLigneMetre,
               ttdPieceTrait,ttdTimbres,ttdLigneOUVP,ttdPieceInterv,ttdPieceDemPrix,TTdArticleDemPrix,
               TtdDetailDemPrix,TtdFournDemprix,ttdVteColl,ttdReajAnal,ttdEntRAnal,ttdTSPOC,ttdTSDetPOC,ttdTRFEntPOC,ttdTRFDetPOC);

  TCaligment = (TalLeft,TalCenter,TalRight);
  TListeTypeDat =(TCNomChps,TClibChps,TClargChps,TCAligmentChps,TCNumChps);
  //
  T_Valeurs = array [0..46] of double;
  //
procedure GetInfoChamps (NomChamps : string;var TypeInfo : string;  var Libelle : string; var Taille : integer);
function IsGraphique(TheTablette : string) : boolean;
function FindTablette (TheSuffixe : string) : string;
function RemplaceColonne( LaListe : string; ColaTrouve, ColRemplace: string) : string;
procedure StockeMontantTypeSurLigne (TOBLIG : TOB);
procedure CumuleSuivantType (TOBP,TOBL : TOB);
procedure CumuleTypeSurPiece (TOBP,TOBL : TOB);
procedure  RecalculeTotauxTypes (TOBP  :TOB);
procedure RecupSousDetailSurLigne (TOBLIG,TOBOUV : TOB);
function GetChampsTot (Nomchamps : string) : string;
procedure ZeroMtLigne (TOBP : TOB);

implementation
uses FactComm,FactVariante;


function GetChampsTot (Nomchamps : string) : string;
begin
  result := '';
  if Copy (Nomchamps,1,6)= 'GLC_MT' then
  begin
    Result := 'GP_TOTAL'+copy(Nomchamps,7,length(Nomchamps)-6);
  end else if (Pos(NomChamps,ZONETOTAL)>0) then
  begin
    Result := 'GP_TOTAL'+copy(Nomchamps,6,length(Nomchamps)-5);
  end;
end;

procedure CumuleTypeSurPiece (TOBP,TOBL : TOB);
var Qte,X : double;
begin
  if IsVariante(TOBL) then exit;
  Qte := TOBL.GetDouble('GL_QTEFACT');
  X:=TOBP.Getdouble('GP_TOTALMOPA')+ Arrondi(TOBL.Getdouble('GLC_MTMOPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALMOPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALMOPR')+Arrondi(TOBL.Getdouble('GLC_MTMOPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALMOPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALMOPV')+Arrondi(TOBL.Getdouble('GLC_MTMOPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALMOPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALFOUPA')+Arrondi(TOBL.Getdouble('GLC_MTFOUPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALFOUPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALFOUPR')+Arrondi(TOBL.Getdouble('GLC_MTFOUPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALFOUPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALFOUPV')+Arrondi(TOBL.Getdouble('GLC_MTFOUPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALFOUPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALINTPA')+Arrondi(TOBL.Getdouble('GLC_MTINTPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALINTPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALINTPR')+Arrondi(TOBL.Getdouble('GLC_MTINTPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALINTPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALINTPV')+Arrondi(TOBL.Getdouble('GLC_MTINTPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALINTPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALLOCPA')+Arrondi(TOBL.Getdouble('GLC_MTLOCPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALLOCPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALLOCPR')+Arrondi(TOBL.Getdouble('GLC_MTLOCPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALLOCPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALLOCPV')+Arrondi(TOBL.Getdouble('GLC_MTLOCPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALLOCPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALMATPA')+Arrondi(TOBL.Getdouble('GLC_MTMATPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALMATPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALMATPR')+Arrondi(TOBL.Getdouble('GLC_MTMATPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALMATPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALMATPV')+Arrondi(TOBL.Getdouble('GLC_MTMATPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALMATPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALOUTPA')+Arrondi(TOBL.Getdouble('GLC_MTOUTPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALOUTPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALOUTPR')+Arrondi(TOBL.Getdouble('GLC_MTOUTPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALOUTPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALOUTPV')+Arrondi(TOBL.Getdouble('GLC_MTOUTPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALOUTPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALSTPA')+Arrondi(TOBL.Getdouble('GLC_MTSTPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALSTPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALSTPR')+Arrondi(TOBL.Getdouble('GLC_MTSTPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALSTPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALSTPV')+Arrondi(TOBL.Getdouble('GLC_MTSTPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALSTPV',X) ;
  X:=TOBP.Getdouble('GP_TOTALAUTPA')+Arrondi(TOBL.Getdouble('GLC_MTAUTPA')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALAUTPA',X) ;
  X:=TOBP.Getdouble('GP_TOTALAUTPR')+Arrondi(TOBL.Getdouble('GLC_MTAUTPR')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALAUTPR',X) ;
  X:=TOBP.Getdouble('GP_TOTALAUTPV')+Arrondi(TOBL.Getdouble('GLC_MTAUTPV')*Qte,V_PGI.okdecV)   ; TOBP.PutValue('GP_TOTALAUTPV',X) ;
end;

procedure ZeroMtLigne (TOBP : TOB);
begin
  TOBP.SetDouble('GLC_MTMOPA', 0);
  TOBP.SetDouble('GLC_MTMOPR', 0);
  TOBP.SetDouble('GLC_MTMOPV', 0);
  TOBP.SetDouble('GLC_MTFOUPA', 0);
  TOBP.SetDouble('GLC_MTFOUPR', 0);
  TOBP.SetDouble('GLC_MTFOUPV', 0);
  TOBP.SetDouble('GLC_MTINTPA', 0);
  TOBP.SetDouble('GLC_MTINTPR', 0);
  TOBP.SetDouble('GLC_MTINTPV', 0);
  TOBP.SetDouble('GLC_MTLOCPA', 0);
  TOBP.SetDouble('GLC_MTLOCPR', 0);
  TOBP.SetDouble('GLC_MTLOCPV', 0);
  TOBP.SetDouble('GLC_MTMATPA', 0);
  TOBP.SetDouble('GLC_MTMATPR', 0);
  TOBP.SetDouble('GLC_MTMATPV', 0);
  TOBP.SetDouble('GLC_MTOUTPA', 0);
  TOBP.SetDouble('GLC_MTOUTPR', 0);
  TOBP.SetDouble('GLC_MTOUTPV', 0);
  TOBP.SetDouble('GLC_MTSTPA', 0);
  TOBP.SetDouble('GLC_MTSTPR', 0);
  TOBP.SetDouble('GLC_MTSTPV', 0);
  TOBP.SetDouble('GLC_MTAUTPA', 0);
  TOBP.SetDouble('GLC_MTAUTPR', 0);
  TOBP.SetDouble('GLC_MTAUTPV', 0);
  //
  TOBP.SetDouble('TOTALMOPA', 0);
  TOBP.SetDouble('TOTALMOPR', 0);
  TOBP.SetDouble('TOTALMOPV', 0);
  TOBP.SetDouble('TOTALFOUPA', 0);
  TOBP.SetDouble('TOTALFOUPR', 0);
  TOBP.SetDouble('TOTALFOUPV', 0);
  TOBP.SetDouble('TOTALINTPA', 0);
  TOBP.SetDouble('TOTALINTPR', 0);
  TOBP.SetDouble('TOTALINTPV', 0);
  TOBP.SetDouble('TOTALLOCPA', 0);
  TOBP.SetDouble('TOTALLOCPR', 0);
  TOBP.SetDouble('TOTALLOCPV', 0);
  TOBP.SetDouble('TOTALMATPA', 0);
  TOBP.SetDouble('TOTALMATPR', 0);
  TOBP.SetDouble('TOTALMATPV', 0);
  TOBP.SetDouble('TOTALOUTPA', 0);
  TOBP.SetDouble('TOTALOUTPR', 0);
  TOBP.SetDouble('TOTALOUTPV', 0);
  TOBP.SetDouble('TOTALSTPA', 0);
  TOBP.SetDouble('TOTALSTPR', 0);
  TOBP.SetDouble('TOTALSTPV', 0);
  TOBP.SetDouble('TOTALAUTPA', 0);
  TOBP.SetDouble('TOTALAUTPR', 0);
  TOBP.SetDouble('TOTALAUTPV', 0);
end;

procedure CumuleSuivantType (TOBP,TOBL : TOB);
var prefixeP,PrefixeL,PrefixeLN : string;
begin

  prefixeP := GetPrefixeTable (TOBP);   if PrefixeP = 'GL' then PrefixeP := 'GLC'; 
  prefixeL := GetPrefixeTable (TOBL);   if PrefixeLN = 'GL' then PrefixeLN := 'GLC';
//
  if TOBL.GetString('GL_TYPELIGNE')<>'ART' then exit;
  (*
  TOBP.SetDouble(PrefixeP+'_MTMOPA', TOBP.GetDouble(PrefixeP+'_MTMOPA') + TOBL.GetDouble(PrefixeL+'_MTMOPA'));
  TOBP.SetDouble(PrefixeP+'_MTMOPR', TOBP.GetDouble(PrefixeP+'_MTMOPR') + TOBL.GetDouble(PrefixeL+'_MTMOPR'));
  TOBP.SetDouble(PrefixeP+'_MTMOPV', TOBP.GetDouble(PrefixeP+'_MTMOPV') + TOBL.GetDouble(PrefixeL+'_MTMOPV'));
  TOBP.SetDouble(PrefixeP+'_MTFOUPA', TOBP.GetDouble(PrefixeP+'_MTFOUPA') + TOBL.GetDouble(PrefixeL+'_MTFOUPA'));
  TOBP.SetDouble(PrefixeP+'_MTFOUPR', TOBP.GetDouble(PrefixeP+'_MTFOUPR') + TOBL.GetDouble(PrefixeL+'_MTFOUPR'));
  TOBP.SetDouble(PrefixeP+'_MTFOUPV', TOBP.GetDouble(PrefixeP+'_MTFOUPV') + TOBL.GetDouble(PrefixeL+'_MTFOUPV'));
  TOBP.SetDouble(PrefixeP+'_MTINTPA', TOBP.GetDouble(PrefixeP+'_MTINTPA') + TOBL.GetDouble(PrefixeL+'_MTINTPA'));
  TOBP.SetDouble(PrefixeP+'_MTINTPR', TOBP.GetDouble(PrefixeP+'_MTINTPR') + TOBL.GetDouble(PrefixeL+'_MTINTPR'));
  TOBP.SetDouble(PrefixeP+'_MTINTPV', TOBP.GetDouble(PrefixeP+'_MTINTPV') + TOBL.GetDouble(PrefixeL+'_MTINTPV'));
  TOBP.SetDouble(PrefixeP+'_MTLOCPA', TOBP.GetDouble(PrefixeP+'_MTLOCPA') + TOBL.GetDouble(PrefixeL+'_MTLOCPA'));
  TOBP.SetDouble(PrefixeP+'_MTLOCPR', TOBP.GetDouble(PrefixeP+'_MTLOCPR') + TOBL.GetDouble(PrefixeL+'_MTLOCPR'));
  TOBP.SetDouble(PrefixeP+'_MTLOCPV', TOBP.GetDouble(PrefixeP+'_MTLOCPV') + TOBL.GetDouble(PrefixeL+'_MTLOCPV'));
  TOBP.SetDouble(PrefixeP+'_MTMATPA', TOBP.GetDouble(PrefixeP+'_MTMATPA') + TOBL.GetDouble(PrefixeL+'_MTMATPA'));
  TOBP.SetDouble(PrefixeP+'_MTMATPR', TOBP.GetDouble(PrefixeP+'_MTMATPR') + TOBL.GetDouble(PrefixeL+'_MTMATPR'));
  TOBP.SetDouble(PrefixeP+'_MTMATPV', TOBP.GetDouble(PrefixeP+'_MTMATPV') + TOBL.GetDouble(PrefixeL+'_MTMATPV'));
  TOBP.SetDouble(PrefixeP+'_MTOUTPA', TOBP.GetDouble(PrefixeP+'_MTOUTPA') + TOBL.GetDouble(PrefixeL+'_MTOUTPA'));
  TOBP.SetDouble(PrefixeP+'_MTOUTPR', TOBP.GetDouble(PrefixeP+'_MTOUTPR') + TOBL.GetDouble(PrefixeL+'_MTOUTPR'));
  TOBP.SetDouble(PrefixeP+'_MTOUTPV', TOBP.GetDouble(PrefixeP+'_MTOUTPV') + TOBL.GetDouble(PrefixeL+'_MTOUTPV'));
  TOBP.SetDouble(PrefixeP+'_MTSTPA', TOBP.GetDouble(PrefixeP+'_MTSTPA') + TOBL.GetDouble(PrefixeL+'_MTSTPA'));
  TOBP.SetDouble(PrefixeP+'_MTSTPR', TOBP.GetDouble(PrefixeP+'_MTSTPR') + TOBL.GetDouble(PrefixeL+'_MTSTPR'));
  TOBP.SetDouble(PrefixeP+'_MTSTPV', TOBP.GetDouble(PrefixeP+'_MTSTPV') + TOBL.GetDouble(PrefixeL+'_MTSTPV'));
  TOBP.SetDouble(PrefixeP+'_MTAUTPA', TOBP.GetDouble(PrefixeP+'_MTAUTPA') + TOBL.GetDouble(PrefixeL+'_MTAUTPA'));
  TOBP.SetDouble(PrefixeP+'_MTAUTPR', TOBP.GetDouble(PrefixeP+'_MTAUTPR') + TOBL.GetDouble(PrefixeL+'_MTAUTPR'));
  TOBP.SetDouble(PrefixeP+'_MTAUTPV', TOBP.GetDouble(PrefixeP+'_MTAUTPV') + TOBL.GetDouble(PrefixeL+'_MTAUTPV'));
  *)
  //
  TOBP.SetDouble('TOTALMOPA', TOBP.GetDouble('TOTALMOPA')+TOBL.GetDouble('TOTALMOPA'));
  TOBP.SetDouble('TOTALMOPR', TOBP.GetDouble('TOTALMOPR')+TOBL.GetDouble('TOTALMOPR'));
  TOBP.SetDouble('TOTALMOPV', TOBP.GetDouble('TOTALMOPV')+TOBL.GetDouble('TOTALMOPV'));
  TOBP.SetDouble('TOTALFOUPA', TOBP.GetDouble('TOTALFOUPA')+TOBL.GetDouble('TOTALFOUPA'));
  TOBP.SetDouble('TOTALFOUPR', TOBP.GetDouble('TOTALFOUPR')+TOBL.GetDouble('TOTALFOUPR'));
  TOBP.SetDouble('TOTALFOUPV', TOBP.GetDouble('TOTALFOUPV')+TOBL.GetDouble('TOTALFOUPV'));
  TOBP.SetDouble('TOTALINTPA', TOBP.GetDouble('TOTALINTPA')+TOBL.GetDouble('TOTALINTPA'));
  TOBP.SetDouble('TOTALINTPR', TOBP.GetDouble('TOTALINTPR')+TOBL.GetDouble('TOTALINTPR'));
  TOBP.SetDouble('TOTALINTPV', TOBP.GetDouble('TOTALINTPV')+TOBL.GetDouble('TOTALINTPV'));
  TOBP.SetDouble('TOTALLOCPA', TOBP.GetDouble('TOTALLOCPA')+TOBL.GetDouble('TOTALLOCPA'));
  TOBP.SetDouble('TOTALLOCPR', TOBP.GetDouble('TOTALLOCPR')+TOBL.GetDouble('TOTALLOCPR'));
  TOBP.SetDouble('TOTALLOCPV', TOBP.GetDouble('TOTALLOCPV')+TOBL.GetDouble('TOTALLOCPV'));
  TOBP.SetDouble('TOTALMATPA', TOBP.GetDouble('TOTALMATPA')+TOBL.GetDouble('TOTALMATPA'));
  TOBP.SetDouble('TOTALMATPR', TOBP.GetDouble('TOTALMATPR')+TOBL.GetDouble('TOTALMATPR'));
  TOBP.SetDouble('TOTALMATPV', TOBP.GetDouble('TOTALMATPV')+TOBL.GetDouble('TOTALMATPV'));
  TOBP.SetDouble('TOTALOUTPA', TOBP.GetDouble('TOTALOUTPA')+TOBL.GetDouble('TOTALOUTPA'));
  TOBP.SetDouble('TOTALOUTPR', TOBP.GetDouble('TOTALOUTPR')+TOBL.GetDouble('TOTALOUTPR'));
  TOBP.SetDouble('TOTALOUTPV', TOBP.GetDouble('TOTALOUTPV')+TOBL.GetDouble('TOTALOUTPV'));
  TOBP.SetDouble('TOTALSTPA', TOBP.GetDouble('TOTALSTPA')+TOBL.GetDouble('TOTALSTPA'));
  TOBP.SetDouble('TOTALSTPR', TOBP.GetDouble('TOTALSTPR')+TOBL.GetDouble('TOTALSTPR'));
  TOBP.SetDouble('TOTALSTPV', TOBP.GetDouble('TOTALSTPV')+TOBL.GetDouble('TOTALSTPV'));
  TOBP.SetDouble('TOTALAUTPA', TOBP.GetDouble('TOTALAUTPA')+TOBL.GetDouble('TOTALAUTPA'));
  TOBP.SetDouble('TOTALAUTPR', TOBP.GetDouble('TOTALAUTPR')+TOBL.GetDouble('TOTALAUTPR'));
  TOBP.SetDouble('TOTALAUTPV', TOBP.GetDouble('TOTALAUTPV')+TOBL.GetDouble('TOTALAUTPV'));
  //
  TOBP.SetDouble(PrefixeP+'_MTMOPA', TOBP.GetDouble('TOTALMOPA'));
  TOBP.SetDouble(PrefixeP+'_MTMOPR', TOBP.GetDouble('TOTALMOPR'));
  TOBP.SetDouble(PrefixeP+'_MTMOPV', TOBP.GetDouble('TOTALMOPV'));
  TOBP.SetDouble(PrefixeP+'_MTFOUPA', TOBP.GetDouble('TOTALFOUPA'));
  TOBP.SetDouble(PrefixeP+'_MTFOUPR', TOBP.GetDouble('TOTALFOUPR'));
  TOBP.SetDouble(PrefixeP+'_MTFOUPV', TOBP.GetDouble('TOTALFOUPV'));
  TOBP.SetDouble(PrefixeP+'_MTINTPA', TOBP.GetDouble('TOTALINTPA'));
  TOBP.SetDouble(PrefixeP+'_MTINTPR', TOBP.GetDouble('TOTALINTPR'));
  TOBP.SetDouble(PrefixeP+'_MTINTPV', TOBP.GetDouble('TOTALINTPV'));
  TOBP.SetDouble(PrefixeP+'_MTLOCPA', TOBP.GetDouble('TOTALLOCPA'));
  TOBP.SetDouble(PrefixeP+'_MTLOCPR', TOBP.GetDouble('TOTALLOCPR'));
  TOBP.SetDouble(PrefixeP+'_MTLOCPV', TOBP.GetDouble('TOTALLOCPV'));
  TOBP.SetDouble(PrefixeP+'_MTMATPA', TOBP.GetDouble('TOTALMATPA'));
  TOBP.SetDouble(PrefixeP+'_MTMATPR', TOBP.GetDouble('TOTALMATPR'));
  TOBP.SetDouble(PrefixeP+'_MTMATPV', TOBP.GetDouble('TOTALMATPV'));
  TOBP.SetDouble(PrefixeP+'_MTOUTPA', TOBP.GetDouble('TOTALOUTPA'));
  TOBP.SetDouble(PrefixeP+'_MTOUTPR', TOBP.GetDouble('TOTALOUTPR'));
  TOBP.SetDouble(PrefixeP+'_MTOUTPV', TOBP.GetDouble('TOTALOUTPV'));
  TOBP.SetDouble(PrefixeP+'_MTSTPA', TOBP.GetDouble('TOTALSTPA'));
  TOBP.SetDouble(PrefixeP+'_MTSTPR', TOBP.GetDouble('TOTALSTPR'));
  TOBP.SetDouble(PrefixeP+'_MTSTPV', TOBP.GetDouble('TOTALSTPV'));
  TOBP.SetDouble(PrefixeP+'_MTAUTPA', TOBP.GetDouble('TOTALAUTPA'));
  TOBP.SetDouble(PrefixeP+'_MTAUTPR', TOBP.GetDouble('TOTALAUTPR'));
  TOBP.SetDouble(PrefixeP+'_MTAUTPV', TOBP.GetDouble('TOTALAUTPV'));
end;

procedure  RecalculeTotauxTypes (TOBP  :TOB);
var prefixeP : string;
    Qte : double;
begin
  prefixeP := GetPrefixeTable (TOBP);
  Qte := TOBP.GetDouble(prefixeP+'_QTEFACT');
  if PrefixeP = 'GL' then PrefixeP := 'GLC';
  TOBP.SetDouble('TOTALMOPA', TOBP.GetDouble(PrefixeP+'_MTMOPA') * Qte);
  TOBP.SetDouble('TOTALMOPR', TOBP.GetDouble(PrefixeP+'_MTMOPR') * Qte);
  TOBP.SetDouble('TOTALMOPV', TOBP.GetDouble(PrefixeP+'_MTMOPV') * Qte);
  TOBP.SetDouble('TOTALFOUPA', TOBP.GetDouble(PrefixeP+'_MTFOUPA') * Qte);
  TOBP.SetDouble('TOTALFOUPR', TOBP.GetDouble(PrefixeP+'_MTFOUPR') * Qte);
  TOBP.SetDouble('TOTALFOUPV', TOBP.GetDouble(PrefixeP+'_MTFOUPV') * Qte);
  TOBP.SetDouble('TOTALINTPA', TOBP.GetDouble(PrefixeP+'_MTINTPA') * Qte);
  TOBP.SetDouble('TOTALINTPR', TOBP.GetDouble(PrefixeP+'_MTINTPR') * Qte);
  TOBP.SetDouble('TOTALINTPV', TOBP.GetDouble(PrefixeP+'_MTINTPV') * Qte);
  TOBP.SetDouble('TOTALLOCPA', TOBP.GetDouble(PrefixeP+'_MTLOCPA') * Qte);
  TOBP.SetDouble('TOTALLOCPR', TOBP.GetDouble(PrefixeP+'_MTLOCPR') * Qte);
  TOBP.SetDouble('TOTALLOCPV', TOBP.GetDouble(PrefixeP+'_MTLOCPV') * Qte);
  TOBP.SetDouble('TOTALMATPA', TOBP.GetDouble(PrefixeP+'_MTMATPA') * Qte);
  TOBP.SetDouble('TOTALMATPR', TOBP.GetDouble(PrefixeP+'_MTMATPR') * Qte);
  TOBP.SetDouble('TOTALMATPV', TOBP.GetDouble(PrefixeP+'_MTMATPV') * Qte);
  TOBP.SetDouble('TOTALOUTPA', TOBP.GetDouble(PrefixeP+'_MTOUTPA') * Qte);
  TOBP.SetDouble('TOTALOUTPR', TOBP.GetDouble(PrefixeP+'_MTOUTPR') * Qte);
  TOBP.SetDouble('TOTALOUTPV', TOBP.GetDouble(PrefixeP+'_MTOUTPV') * Qte);
  TOBP.SetDouble('TOTALSTPA', TOBP.GetDouble(PrefixeP+'_MTSTPA') * Qte);
  TOBP.SetDouble('TOTALSTPR', TOBP.GetDouble(PrefixeP+'_MTSTPR') * Qte);
  TOBP.SetDouble('TOTALSTPV', TOBP.GetDouble(PrefixeP+'_MTSTPV') * Qte);
  TOBP.SetDouble('TOTALAUTPA', TOBP.GetDouble(PrefixeP+'_MTAUTPA') * Qte);
  TOBP.SetDouble('TOTALAUTPR', TOBP.GetDouble(PrefixeP+'_MTAUTPR') * Qte);
  TOBP.SetDouble('TOTALAUTPV', TOBP.GetDouble(PrefixeP+'_MTAUTPV') * Qte);
end;

procedure RecupSousDetailSurLigne (TOBLIG,TOBOUV : TOB);
var Prefixe,NPrefixe,PrefixeO : string;
begin
  Prefixe := GetPrefixeTable(TOBLIG);
  PrefixeO := GetPrefixeTable(TOBOUV);
  if Prefixe = 'GL' then NPrefixe := 'GLC' else Nprefixe := Prefixe;
  //
  TOBLIG.SetDouble(NPrefixe+'_MTMOPA',TOBOUV.GetDouble(PrefixeO+'_MTMOPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTMOPR',TOBOUV.GetDouble(PrefixeO+'_MTMOPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTMOPV',TOBOUV.GetDouble(PrefixeO+'_MTMOPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTINTPA',TOBOUV.GetDouble(PrefixeO+'_MTINTPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTINTPR',TOBOUV.GetDouble(PrefixeO+'_MTINTPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTINTPV',TOBOUV.GetDouble(PrefixeO+'_MTINTPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTMATPA',TOBOUV.GetDouble(PrefixeO+'_MTMATPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTMATPR',TOBOUV.GetDouble(PrefixeO+'_MTMATPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTMATPV',TOBOUV.GetDouble(PrefixeO+'_MTMATPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTLOCPA',TOBOUV.GetDouble(PrefixeO+'_MTLOCPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTLOCPR',TOBOUV.GetDouble(PrefixeO+'_MTLOCPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTLOCPV',TOBOUV.GetDouble(PrefixeO+'_MTLOCPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTOUTPA',TOBOUV.GetDouble(PrefixeO+'_MTOUTPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTOUTPR',TOBOUV.GetDouble(PrefixeO+'_MTOUTPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTOUTPV',TOBOUV.GetDouble(PrefixeO+'_MTOUTPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTSTPA',TOBOUV.GetDouble(PrefixeO+'_MTSTPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTSTPR',TOBOUV.GetDouble(PrefixeO+'_MTSTPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTSTPV',TOBOUV.GetDouble(PrefixeO+'_MTSTPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTAUTPA',TOBOUV.GetDouble(PrefixeO+'_MTAUTPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTAUTPR',TOBOUV.GetDouble(PrefixeO+'_MTAUTPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTAUTPV',TOBOUV.GetDouble(PrefixeO+'_MTAUTPV'));
  TOBLIG.SetDouble(NPrefixe+'_MTFOUPA',TOBOUV.GetDouble(PrefixeO+'_MTFOUPA'));
  TOBLIG.SetDouble(NPrefixe+'_MTFOUPR',TOBOUV.GetDouble(PrefixeO+'_MTFOUPR'));
  TOBLIG.SetDouble(NPrefixe+'_MTFOUPV',TOBOUV.GetDouble(PrefixeO+'_MTFOUPV'));
  RecalculeTotauxTypes (TOBLIG);
end;


procedure StockeMontantTypeSurLigne (TOBLIG : TOB);
var Prefixe,NPrefixe : string;
begin
  Prefixe := GetPrefixeTable(TOBLIG);
  if Prefixe = 'GL' then NPrefixe := 'GLC' else if Prefixe='BOP' then NPrefixe := 'GLC' else Nprefixe := Prefixe;
  if (Prefixe = 'BLO') and (TOBLIG.detail.count > 0) then exit;
  //
  if TOBLIG.getString('BNP_TYPERESSOURCE')='SAL' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTMOPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTMOPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTMOPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else if TOBLIG.getString('BNP_TYPERESSOURCE')='INT' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTINTPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTINTPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTINTPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else if TOBLIG.getString('BNP_TYPERESSOURCE')='MAT' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTMATPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTMATPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTMATPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else if TOBLIG.getString('BNP_TYPERESSOURCE')='LOC' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTLOCPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTLOCPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTLOCPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else if TOBLIG.getString('BNP_TYPERESSOURCE')='OUT' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTOUTPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTOUTPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTOUTPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else if TOBLIG.getString('BNP_TYPERESSOURCE')='ST' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTSTPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTSTPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTSTPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else if TOBLIG.getString('BNP_TYPERESSOURCE')='AUT' then
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTAUTPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTAUTPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTAUTPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end else
  begin
    TOBLIG.SetDouble(NPrefixe+'_MTFOUPA',TOBLIG.GetDouble(prefixe+'_DPA'));
    TOBLIG.SetDouble(NPrefixe+'_MTFOUPR',TOBLIG.GetDouble(prefixe+'_DPR'));
    TOBLIG.SetDouble(NPrefixe+'_MTFOUPV',TOBLIG.GetDouble(prefixe+'_PUHT'));
  end;
  RecalculeTotauxTypes (TOBLIG);
end;

function RemplaceColonne( LaListe : string; ColaTrouve, ColRemplace: string) : string;
var II : integer;
begin
  II := Pos(ColaTrouve,LaListe);
  if II >=0  then
  begin
    II := Pos(ColaTrouve,ColRemplace);
    if II < 0 then   // on remplace uniquement si elle n'est pas déjà présente dans la liste
    begin
      LaListe := FindEtReplace(LaListe, ColaTrouve, ColRemplace, False);
    end;
  end;
end;

function FindTablette (TheSuffixe : string) : string;
var QQ: TQuery;
begin
  result := '';
  if TheSuffixe = '' then exit;
  QQ := OPenSql ('SELECT DO_COMBO FROM DECOMBOS WHERE DO_NOMCHAMP LIKE "%'+TheSuffixe+'%"',true,1,'',true);
  if not QQ.eof then
  begin
    result := QQ.fields[0].AsString;
  end;
  ferme (QQ);
end;

function IsGraphique(TheTablette : string) : boolean;
var QQ: TQuery;
begin
  result := false;
  if TheTablette = '' then exit;
  QQ := OPenSql ('SELECT DO_NOMCHAMP FROM DECOMBOS WHERE DO_COMBO="'+TheTablette+'"',true,1,'',true);
  if not QQ.eof then
  begin
    result := (pos ('GRAPHIQUE',QQ.fields[0].AsString) >0);
  end;
  ferme (QQ);
end;

procedure GetInfoChamps (NomChamps : string;var TypeInfo : string;  var Libelle : string; var Taille : integer);
var QQ : TQuery;
    NBC : integer;
begin
  if NomChamps = 'NUMTRANSFERT' then
  begin
    Libelle:= 'N° transfert';
    TypeInfo := 'VARCHAR(6)';
    Taille := 6;
  end else if NomChamps = 'MTTRANSFERT' then
  begin
    Libelle:= 'Mt Transfert';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'SUMTOTALTS' then
  begin
    Libelle:= 'Total TS';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'POURCENTMARG' then
  begin
    Libelle:= '% Marge';
    TypeInfo := 'RATE';
    Taille := 6;
  end else if NomChamps = 'POURCENTMARQ' then
  begin
    Libelle:= '% Marque';
    TypeInfo := 'RATE';
    Taille := 12;
  end else if NomChamps = 'TOTALMOPA' then
  begin
    Libelle:= 'Tot Mo PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALMOPR' then
  begin
    Libelle:= 'Tot Mo PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALMOPV' then
  begin
    Libelle:= 'Tot Mo PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALFOUPA' then
  begin
    Libelle:= 'Tot Fourn. PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALFOUPR' then
  begin
    Libelle:= 'Tot Fourn. PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALFOUPV' then
  begin
    Libelle:= 'Tot Fourn. PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALINTPA' then
  begin
    Libelle:= 'Tot Interim. PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALINTPR' then
  begin
    Libelle:= 'Tot Interim. PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALINTPV' then
  begin
    Libelle:= 'Tot Interim. PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALLOCPA' then
  begin
    Libelle:= 'Tot Location PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALLOCPR' then
  begin
    Libelle:= 'Tot Location PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALLOCPV' then
  begin
    Libelle:= 'Tot Location PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALMATPA' then
  begin
    Libelle:= 'Tot Matériel PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALMATPR' then
  begin
    Libelle:= 'Tot Matériel PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALMATPV' then
  begin
    Libelle:= 'Tot Matériel PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALOUTPA' then
  begin
    Libelle:= 'Tot Outil. PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALOUTPR' then
  begin
    Libelle:= 'Tot Outil. PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALOUTPV' then
  begin
    Libelle:= 'Tot Outil. PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALSTPA' then
  begin
    Libelle:= 'Tot S/T PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALSTPR' then
  begin
    Libelle:= 'Tot S/T PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALSTPV' then
  begin
    Libelle:= 'Tot S/T PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALAUTPA' then
  begin
    Libelle:= 'Tot Autres PA';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALAUTPR' then
  begin
    Libelle:= 'Tot Autres PR';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else if NomChamps = 'TOTALAUTPV' then
  begin
    Libelle:= 'Tot Autres PV';
    TypeInfo := 'DOUBLE';
    Taille := 12;
  end else
  begin
    Libelle := '?????';
    Taille  := 20;
    QQ := OpenSql ('SELECT DH_LIBELLE,DH_TYPECHAMP FROM DECHAMPS WHERE DH_NOMCHAMP="'+NomChamps+'"',true,1,'',true);
    if not QQ.eof then
    begin
      Libelle := QQ.fields[0].AsString;
      TypeInfo := QQ.Fields[1].AsString;
      if TypeInfo = 'INTEGER' then Taille := 5
      else if TypeInfo = 'DATE' then taille := 10
      else if TypeInfo = 'RATE' then Taille := 8
      else if TypeInfo = 'BOOLEAN' then Taille := 1
      else if TypeInfo = 'COMBO' then Taille := 35
      else if copy(TypeInfo,1,7) = 'VARCHAR' then
      begin
        NBC := Length(TypeInfo) - 10;
        Taille := StrToInt(copy(TypeInfo,9,NBC));
      end else if (TypeInfo = 'DOUBLE') or (TypeInfo = 'EXTENDED') then Taille := 12;
    end;
    ferme (QQ);
  end;
end;


end.















