unit EtudePiece;


interface

uses {$IFDEF VER150} variants,{$ENDIF} Classes,UTOB,Hctrls,UtilXlsBTP,HMsgBox,EtudesExt,HEnt1,SysUtils,ed_tools,Ent1,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,UserChg,AglIsoflex,
{$ENDIF}
{$IFDEF BTP}
 FactOuvrage,BTPUtil, CalcOLEGenericBTP,
{$ENDIF}
 FactComm,FactCpta,
 FactTOB,FactTiers,FactAdresse,FactLotSerie,FactPiece,
 ParamSoc,EntGC,FactUtil,SaisUtil,FactNomen,FactCalc,
 AffaireUtil,FactAdresseBTP,UtilArticle,UtilTOBpiece,uEntCommun,HRichOLE,forms, UtilsMetresXLS;

type TPlageExcel = class
     feuille        : string;
     mini           : integer;
     maxi           : integer;

end;
type TColonnes = class
     colonne : integer;
     transfert : boolean;
end;
type TPieceEtude = class
     private
     fform          : Tform;
     fcledoc        : R_CleDoc;
     fTiers         : string;
     fCodeAffaire   : string;
     fEtablissement : String;
     fDomaine       : String;
     FDocExcel      : String;
     fCurSheet      : String;
     fEnHT          : Boolean;
     fSaisieContre  : Boolean;
     SaContexte     : TModeAff;
     docOpen        : boolean;
     DEV            : Rdevise;
     fWinExcel      : OleVariant;
     WorkBook       : Variant;
     WorkSheet      : Variant;
     IndiceOuv      : integer;

     fColCode       : Tcolonnes;
     fColDes        : Tcolonnes;
     fColQte        : Tcolonnes;
     fColUnite      : Tcolonnes;
     fColPu         : Tcolonnes;
     fColMontant    : Tcolonnes;
     fColArticle    : Tcolonnes;

     fPlages        : Tlist;
     FModifBordereau: Boolean;

     fcodePrest     : string; // il servira pour les etudes secondaires
     fNatureTravail : string;
     fFournisseur   : string;
     //
     TOBOuvragesP : TOB;

     TOBPieceTrait : TOB;

     TOBBases       : TOB;
     TOBBasesL      : TOB;
     TOBEches       : TOB;
     TOBPorcs       : TOB;
     TOBTiers       : TOB;
     TOBArticles    : TOB;
     TOBConds       : TOB;
     TOBTarif       : TOB;
     TOBComms       : TOB;
     TOBCatalogu    : TOB;
     TOBAdresses    : TOB;
     TOBCXV         : TOB;
     TOBNomen       : TOB;
     TOBDim         : TOB;
     TOBLOT         : TOB;
     TOBSerie       : TOB;
     TOBSerRel      : TOB;
     TOBAcomptes    : TOB;
     TOBDispoContreM: TOB;
     TOBAffaire     : TOB;
     TOBCPTA        : TOB;
     TOBANAS        : TOB;
     TOBANAP        : TOB;
     TOBGSA         : TOB;
     TOBOuvrage     : TOB;
     TOBLIENOLE     : TOB;
     TOBPieceRG     : TOB;
     TobBasesRG     : TOB;
     (*TOBLignesRg,*)
     TOBLienXls     : TOB;
     TOBPIECE_O     : TOB;
     TOBSSTRAIT     : TOB;
     TOBPIECE       : TOB;

     WithSousDetail : boolean;

     TheMetreDoc    : TMetreArt;

     procedure alloueTob;
     procedure DetruitTOB;
     procedure ChargeTobs;
     function  CreeDoc : R_Cledoc;
     procedure ouvreDocPGI;

     procedure FermeDocExcelAssocie;
//     Procedure SauveDocPgi;
     procedure definiPlages(TobStrDoc: TOB);
     procedure IndiquePlages(TOBL:TOB);
     procedure DefiniColonnes(TOBStrDoc: TOB);
     procedure MajTobs;
     procedure SupprPrecedent;
     procedure QualifiquationColsXls;
     //procedure FormatColxls(Colonne: Integer; Format: string);
     procedure FermeDoc;
     procedure ReinitPiece;
     procedure DefiniPlagesWorkbook;
     procedure IncidenceTiers (TOBPiece,TOBTiers,TOBAdresses : TOB);
     function NewPieceVide(var CleDoc: R_CleDoc; CodeTiers, CodeAffaire, Etablissement, Domaine: String; EnHT, SaisieContre: Boolean): TOB;
     procedure ChargeAffaireEtude;
     function findTobArt (ArticleDivers : string) : TOB;
     procedure ajusteDocument;
     procedure SetModeBordereau (Value : boolean);
    	procedure TraitePrixLigneCourante(TOBL: TOB; Pu: double);
     // Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
     procedure TraitementArticle(TOBL: TOB; CodeArticle, Qte : string; LignePgi : integer);
     public

     destructor Destroy; override;
     constructor Create (form : Tform);
     procedure OuvreDocExcelAssocie;

     property  NomFicXls : string read FDocExcel write FdocExcel;
     property  CleDoc : R_Cledoc read FCledoc write Fcledoc;
     property  CodePrestation : string read fcodeprest write fcodeprest;
     property  ModifBord : Boolean      Read FModifBordereau  Write FModifBordereau;

     property  Bordereau : boolean write SetModeBordereau ;
     property  fournisseur : String  read fFournisseur write fFournisseur;
     property  NatureTravail : String read fNatureTravail  write fNatureTravail ;
     procedure DefiniStructureXls (StructureExcel : string);
     procedure NewDoc ( var CleDoc : R_CleDoc ; CodeTiers,CodeAffaire,Etablissement,Domaine : String ; EnHT,SaisieContre : Boolean );
     procedure printPGI ;
     procedure printXls ;
     Procedure VersBordereau ;
     Procedure VersExcel;
     Procedure VersPgi ;
     procedure ChargeClefAffairePiece;
     function IsDocumentvalorise : Boolean;
     procedure NettoieDoc (TobPiece : TOB);

end;

implementation
uses factureBTP,Factspec,UtilPGI,HTB97;

{ TPieceEtude }

procedure TPieceEtude.alloueTob;
begin
TOBPiece:=TOB.Create('PIECE',Nil,-1)     ;
AddLesSupEntete (TOBPiece);
TOBPiece_O := TOB.create ('PIECEORIG',nil,-1);
// ---
TOBBases:=TOB.Create('BASES',Nil,-1)     ;
TOBBasesL:=TOB.Create('LES BASES LIGNE',Nil,-1)     ;
TOBEches:=TOB.Create('LES ECHEANCES',Nil,-1) ;
TOBPorcs:=TOB.Create('PORCS',Nil,-1)     ;
// Fiches
TOBTiers:=TOB.Create('TIERS',Nil,-1) ;
TOBTiers.AddChampSup('RIB',False) ;

TOBArticles:=TOB.Create('ARTICLES',Nil,-1) ;
TOBConds:=TOB.Create('CONDS',Nil,-1) ;
TOBTarif:=TOB.Create('TARIF',Nil,-1) ;
TOBComms:=TOB.Create('COMMERCIAUX',Nil,-1) ;
TOBCatalogu:=TOB.Create('LECATALOGUE',Nil,-1) ;
// Adresses
TOBAdresses:=TOB.Create('LESADRESSES',Nil,-1) ;
if GetParamSoc('SO_GCPIECEADRESSE') then
   BEGIN
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Livraison}
   TOB.Create('PIECEADRESSE',TOBAdresses,-1) ; {Facturation}
   END else
   BEGIN
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Livraison}
   TOB.Create('ADRESSES',TOBAdresses,-1) ; {Facturation}
   END ;
// Divers
TOBCXV:=TOB.Create('',Nil,-1) ;
TOBNomen:=TOB.Create('NOMENCLATURES',Nil,-1) ;
TOBDim:=TOB.Create('',Nil,-1) ;
TOBLOT:=TOB.Create('',Nil,-1) ;
TOBSerie:=TOB.Create('',Nil,-1) ;
TOBSerRel:=TOB.Create('',Nil,-1) ;
TOBAcomptes:=TOB.Create('',Nil,-1) ;
TOBDispoContreM:=TOB.Create('',Nil,-1);
{$IFDEF CHR}
TOBHRDossier:=TOB.Create('HRDOSSIER',Nil,-1);
TOBHrDossier.AddChampSup ('HDR_HRDOSSIER', False);
TOBHrDossier.AddChampSup ('HDR_DOSRES', False);
TOBHrDossier.AddChampSup ('HDR_DATEARRIVEE', False);
TOBHrDossier.AddChampSup ('HDR_DATEDEPART', False);
TOBHrDossier.AddChampSup ('HDR_TIERS', False);
TOBHrDossier.AddChampSup ('HDR_NBPERSONNE1', False);
TOBHrDossier.AddChampSup ('HDR_RESSOURCE', False);
{$ENDIF}
// Affaires
TOBAffaire:=TOB.Create('AFFAIRE',Nil,-1) ;
// Comptabilité
TOBCPTA:=CreerTOBCpta ;
TOBANAP:=TOB.Create('',Nil,-1) ;
TOBANAS:=TOB.Create('',Nil,-1) ;
//Saisie Code Barres
TOBGSA:=TOB.Create('',Nil,-1);
// Ouvrages
TOBOuvrage:=TOB.Create ('OUVRAGES',nil,-1);
// Ouvrages^plat
TOBOuvragesP:=TOB.Create ('OUVRAGES PLAT',nil,-1);
// textes debut et fin
TOBLIENOLE:=TOB.Create ('LIENS',nil,-1);
// retenues de garantie
TOBPieceRG:=TOB.create('PIECERRET',Nil,-1);
// Bases de tva sur RG
TOBBasesRG:=TOB.create('BASESRG',Nil,-1);
TOBLienXls := TOB.Create ('LIENSPGIEXCEL',nil,-1);
TOBPieceTrait := TOB.Create ('LES PIECES TRAITS',nil,-1);
TOBSSTrait := TOB.Create ('LES SOUS TRAITS',nil,-1);
end;

procedure TPieceEtude.DetruitTOB;
begin
TOBPiece.free; TOBPiece_O.free;
TOBBases.free;
TOBBasesL.free;
TOBEches.free;
TOBPorcs.free;
// Fiches
TOBTiers.free;
TOBArticles.free;
TOBConds.free;
TOBTarif.free;
TOBComms.free;
TOBCatalogu.free;
// Adresses
TOBAdresses.free;
// Divers
TOBCXV.free;
TOBNomen.free;
TOBDim.free;
TOBLOT.free;
TOBSerie.free;
TOBSerRel.free;
TOBAcomptes.free;
TOBDispoContreM.free;
{$IFDEF CHR}
TOBHRDossier.free;
{$ENDIF}
// Affaires
TOBAffaire.free;
// Comptabilité
TOBCPTA.free;
TOBANAP.free;
TOBANAS.free;
//Saisie Code Barres
TOBGSA.free;
// Ouvrages
TOBOuvrage.free;
// Ouvrages plat
TOBOuvragesP.free;
// textes debut et fin
TOBLIENOLE.free;
// retenues de garantie
TOBPieceRG.free;
// Bases de tva sur RG
TOBBasesRG.free;
// Lien PGI-Excel
TOBLienXls.free;
TOBPieceTrait.free;
TOBSSTrait.free;
end;

constructor TPieceEtude.Create (form : Tform);
begin
//
fform := form;
WithSousDetail := true; // par defaut;
docOpen        := false;
IndiceOuv      := 1;
fPlages        := Tlist.create;
fColCode       := TColonnes.create;
fColDes        := TColonnes.create;
fColQte        := TColonnes.create;
fColUnite      := TColonnes.create;
fColPu         := TColonnes.create;
fColMontant    := TColonnes.create;
fColArticle    := TColonnes.create;

fFournisseur   := '';
FModifBordereau:= False;

end;

destructor TPieceEtude.Destroy;
var Plage : TPlageExcel;
    Indice : integer;
begin
if DocOpen then FermeDoc;
for Indice := 0 to fPlages.Count -1 do
    begin
    Plage := fPlages.items[Indice];
    Plage.Free;
    end;
fplages.free;
fplages := nil;
fColCode.free;
fColDes.free;
fColQte.free;
fColUnite.free;
fColPu.free;
fColMontant.free;
fColArticle.Free;
inherited;
end;

function TPieceEtude.creedoc: R_Cledoc;
var LocCledoc : R_Cledoc;
begin
loccledoc := fcledoc;
TOBPiece := NewPieceVide (loccledoc,ftiers,fcodeaffaire,fetablissement,fdomaine,fenHt,fsaisiecontre);
if LocCleDoc.NumeroPiece = 0 then exit;
fcledoc := loccledoc;
result := fcledoc;
end;

procedure TPieceEtude.NewDoc(var CleDoc: R_CleDoc; CodeTiers, CodeAffaire, Etablissement, Domaine: String; EnHT, SaisieContre: Boolean);
var DomaineRetenue : string;
begin
	DomaineRetenue := GetParamSocSecur('SO_BTDOMAINEDEFAUT','');
  //
  fcledoc := cledoc;
  fCodeAffaire := CodeAffaire;
  fTiers := CodeTiers;
  fEtablissement := Etablissement;
  fdomaine := domaine;
  FEnht := EnHt;
  fSaisieContre := SaisieContre;
  alloueTob;
  fcledoc := creedoc;
  if TOBPiece <> nil then
  begin
     RemplirTOBTiers (TOBTiers,TOBPiece.getValue('GP_TIERS'),fcledoc.NaturePiece,false);
     IncidenceTiers (TOBPiece,TOBTiers,TOBAdresses);
  end;
  if (TobAffaire <> nil) and (fCodeAffaire<>'') then
  begin
  	ChargeClefAffairePiece;
  	ChargeAffaireEtude;
  	AffaireVersPiece (TobPiece,TOBaffaire);
		LivAffaireVersAdresses (TobAffaire,TOBAdresses,TOBPiece);
    if(DomaineRetenue = '') and (TOBAffaire.GetValue('AFF_DOMAINE')<>'') then
    begin
      DomaineRetenue := TOBAffaire.GetValue('AFF_DOMAINE');
    end;
  end;
  if DomaineRetenue <> '' then
  begin
    TOBPiece.SetString('GP_DOMAINE',DomaineRetenue);
  end;
  if TRANSACTIONS (MajTobs,1) <> oeok then PGIBox (TraduireMemoire('Erreur dans l''écriture de la pièce'),'Liaison EXCEL');
  FermeDoc;
  cledoc := fcledoc;
end;

procedure TPieceEtude.ouvreDocPGI;
var GestionNotesDoc : string;
begin
if DocOpen then exit;
if GetParamSoc('SO_BTDOCAVECTEXTE') then
   begin
   SaContexte := saContexte + [tModeBlocNotes];
   end;
GestionNotesDoc :=GetInfoParPiece(fcledoc.naturePiece,'GPP_TYPEPRESDOC');
if GestionNotesDoc = 'AUC' then
   begin
   if (tModeBlocNotes in SaContexte) then SaContexte := SaContexte - [TModeBlocNotes];
   end;
if GestionNotesDoc = 'TOU' then
   begin
   if not (tModeBlocnotes in SaContexte) then SaContexte := SaContexte + [tModeBlocNotes];
   end;
alloueTob;
chargeTobs;
{On en profite pour positionner les variables locales}
fTiers := TOBPiece.GetValue('GP_TIERS');
fCodeAffaire := TOBPiece.GetValue('GP_AFFAIRE');
fEtablissement:= TOBPIece.GEtValue('GP_ETABLISSEMENT');
fDomaine := TOBPiece.GEtVAlue('GP_DOMAINE');
fEnHT:= (TOBPIece.getvalue('GP_FACTUREHT')='X');
fSaisieContre := (TOBPIece.getvalue('GP_SAISIECONTRE')='X');
{et voila ..on est pret}
if FdocExcel <> '' then OuvreDocExcelAssocie;
DocOpen := true;
end;

procedure TPieceEtude.FermeDoc;
begin
if DocOpen then
   begin
   DetruitTOB;
   DocOpen := false;
   end;
if FdocExcel <> '' then FermeDocExcelAssocie;
end;

procedure TPieceEtude.ChargeTobs;
var Q: Tquery;
begin
(*
// Lecture Piece
Q:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(fCleDoc,ttdPiece,False),True) ;
TOBPiece.SelectDB('',Q) ;
Ferme(Q) ;
TOBPiece_O.Dupliquer (TOBPiece,true,true);
// Lecture Lignes
Q:=OpenSQL('SELECT LIGNE.*,LIGNECOMPL.* FROM LIGNE '+
					 'LEFT JOIN LIGNECOMPL ON (GL_NATUREPIECEG = GLC_NATUREPIECEG and GL_SOUCHE = GLC_SOUCHE AND GL_NUMERO = GLC_NUMERO ' +
           'AND GL_INDICEG = GLC_INDICEG and GL_NUMORDRE = GLC_NUMORDRE) '+
           'WHERE '+ WherePiece(fCleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
TOBPiece.LoadDetailDB('LIGNE','','',Q,False,True) ;
Ferme(Q) ;
PieceAjouteSousDetail(TOBPiece);
*)
  LoadPieceLignes(fCleDoc, TobPiece);
  PieceAjouteSousDetail(TOBPiece,true,false,true);

// Lecture bases
Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(fCleDoc,ttdPiedBase,False),True,-1, '', True) ;
TOBBases.LoadDetailDB('PIEDBASE','','',Q,False) ;
Ferme(Q) ;
// Lecture bases Lignes
Q:=OpenSQL('SELECT * FROM LIGNEBASE WHERE '+WherePiece(fCleDoc,ttdLigneBase,False),True,-1, '', True) ;
TOBBasesL.LoadDetailDB('LIGNEBASE','','',Q,False) ;
Ferme(Q) ;
// Lecture Echéances
Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(fCleDoc,ttdEche,False),True,-1, '', True) ;
TOBEches.LoadDetailDB('PIEDECHE','','',Q,False) ;
Ferme(Q) ;
// Lecture Ports
Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(fCleDoc,ttdPorc,False),True,-1, '', True) ;
TOBPorcs.LoadDetailDB('PIEDPORT','','',Q,False) ;
Ferme(Q) ;
// lecture table tiers
RemplirTOBTiers (TOBTiers,TOBPiece.getValue('GP_TIERS'),fcledoc.NaturePiece,false);
// Lecture Adresses
LoadLesAdresses(TOBPiece,TOBAdresses) ;
// Lecture Nomenclatures
LoadLesNomen(TOBPiece,TOBNomen,TOBArticles,fCleDoc) ;
// Lecture Lot
LoadLesLots(TOBPiece,TOBLOT,fCleDoc) ;
{$IFNDEF CCS3}
// Lecture Serie
LoadLesSeries(TOBPiece,TOBSerie,fCleDoc) ;
{$ENDIF}
// Lecture ACompte
LoadLesAcomptes(TOBPiece,TOBAcomptes,fCleDoc) ;
// Lecture Analytiques
LoadLesAna(TOBPiece,TOBAnaP,TOBAnaS) ;
// Modif BTP
// chargement textes debut et fin
LoadLesBlocNotes (SaContexte,TOBLienOle,fCledoc);
// Chargement des retenues de garantie et Tva associe
LoadLesRetenues (TOBPiece,TOBPieceRG,TOBBasesRG);
// Lecture Ouvrages
{$IFDEF BTP}
LoadLesOuvrages (TOBPiece,TOBOuvrage,TOBArticles,fCledoc,IndiceOuv);
{$ENDIF}
// chargement des details ouvrage ou nomenclatures
DEV.Code:=TOBPIECE.GetValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
DEV.Taux:=TOBPiece.GetValue('GP_TAUXDEV');
if WithSousDetail then LoadLesLibDetail(TOBPiece,TOBNomen,TobOuvrage,TOBTiers,TOBAffaire,nil,DEV,False, TheMetreDoc);
{$IFDEF CHR}
//LoadLesTobChr; {pas possible pour l'instant ..voir avec CHR si besoin est }
{$ENDIF}

Q:=OpenSQL('SELECT * FROM BLIENEXCEL WHERE BLX_PIECEASSOCIEE="'+CodeLaPiece(fcledoc)+'"',True,-1, '', True) ;
TOBLienXls.LoadDetailDB('BLIENEXCEL','','',Q,False) ;
Ferme(Q) ;
//
// Lecture Affaire
Q:=OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBPIECE.GetValue('GP_AFFAIRE')+'"',True,-1, '', True) ;
TOBAffaire.SelectDB('',Q) ;
Ferme(Q) ;
end;

procedure TPieceEtude.FermeDocExcelAssocie;
begin
if not VarIsEmpty(fWinExcel) then fWinExcel.Quit;
fWinExcel := unassigned;
end;

procedure TPieceEtude.OuvreDocExcelAssocie;
var WinNew : boolean;
begin
if fdocExcel = '' then BEGIN PgiBox (TraduireMemoire('Nom de fichier non défini'),'Liaison EXCEL'); exit; END;
if not OpenExcel (true,fWinExcel,WinNew) then
   begin
   PGIBox (TraduireMemoire('Liaison Excel impossible'),'Liaison EXCEL');
   FDocExcel := '';
   exit;
   end;
WorkBook := OpenWorkBook (FDocExcel ,fWinExcel);
If Workbook.MultiUserEditing Then
   begin
   Workbook.AutoUpdateFrequency:=0;
   end;
end;

procedure TPieceEtude.VersExcel;
var
    Indice : integer;
    TOBL,TOBLXLS : TOB;
    LigneXls,IndOrdre : integer;
    ValeurXls,CodPiece,Donnee : string;
begin
   if fcledoc.NumeroPiece = 0 then BEGIN PGIBox (TraduireMemoire('Document PGI non défini'),'Liaison EXCEL'); exit; END;
   TRY
   WithSousDetail := false;
   ouvreDocPGI;
   CodPiece := EncodeRefPiece (TOBPiece);
// génération a travers la structure du fichier Excel
   for Indice := 0 to TobPiece.detail.Count -1 do
       begin
       TOBL := TOBPiece.detail[Indice];
       IndOrdre := TOBL.getValue('GL_NUMORDRE');
       TOBLXLS := TOBLienXls.findfirst(['BLX_PIECEASSOCIEE','BLX_NUMLIGNE'],[CodPiece,IndOrdre],true);
       if TOBLXLS <> nil then
           begin
           fcursheet := TOBLXLS.GetValue('BLX_FEUILLEXLS');
           WorkSheet := SelectSheet (WorkBook,fcurSheet);
           QualifiquationColsXls;
           LigneXls := TOBLXLS.GetValue('BLX_NUMLIGXLS');
           if (fColQte.colonne > 0) and (fcolQte.transfert ) then
           begin
             ValeurXls := GetExcelText (WorkSheet,LigneXls,fColQte.colonne);
             if ValeurXls <> '' then
             begin
              if pos(',',ValeurXls) > 0 then StringReplace (ValeurXls,',','.',[rfReplaceAll]);
//                 if Valeur(ValeurXls) <> 0 then ExcelValue (WorkSheet,LigneXls-1,fcolQte.colonne-1,valeur(ValeurXls));
             end;
           end else
           begin
             ValeurXls := '1';
           end;
           if (fColPu.colonne > 0) and (fcolpu.transfert ) and (TOBL.GetValue('GL_PUHTDEV') <> 0) then
              begin
                if ValeurXls <> '' then ExcelValue (WorkSheet,LigneXls-1,fcolpu.colonne-1,TOBL.GetValue('GL_PUHTDEV'));
              end;
           if (fColMontant.colonne > 0) and (fColMontant.transfert ) then
              begin
              if (TOBL.GetValue('GL_MONTANTHTDEV') <> 0) then
                 begin
                  if ValeurXls <> '' then ExcelValue (WorkSheet,LigneXls-1,fcolMontant.colonne-1,TOBL.GetValue('GL_MONTANTHTDEV'));
                 end;
              Donnee := GetExcelText(worksheet,LigneXls,fcolMontant.colonne);
              if pos('[GP_MONTANTHTDEV]',Donnee) > 0 then
                 begin
                 if TobPiece.GetValue('GP_TOTALHTDEV') <> 0 then ExcelValue (WorkSheet,LigneXls-1,fcolMontant.colonne-1,TOBPiece.GetValue('GP_TOTALHTDEV'));
                 end;
              if pos('[GP_MONTANTTTCDEV]',Donnee) > 0 then
                 begin
                 if TobPiece.GetValue('GP_TOTALTTCDEV') <> 0 then ExcelValue (WorkSheet,LigneXls-1,fcolMontant.colonne-1,TOBPiece.GetValue('GP_TOTALTTCDEV'));
                 end;
              if pos('[GP_MONTANTTVA]',Donnee) > 0 then
                 begin
                 if (TobPiece.GetValue('GP_TOTALTTCDEV')-TobPiece.GetValue('GP_TOTALHTDEV')) <> 0 then ExcelValue (WorkSheet,LigneXls-1,fcolMontant.colonne-1,(TOBPiece.GetValue('GP_TOTALTTCDEV')-TOBPiece.GetValue('GP_TOTALHTDEV')));
                 end;
              end;
            if (fColArticle.colonne > 0) and (fColArticle.transfert ) and ((TOBL.GetValue('GL_CODEARTICLE') <> '') OR (TOBL.GetValue('GL_ARTICLE') <> '')) then
            begin
              if ValeurXls <> '' then ExcelValue (WorkSheet,LigneXls-1,fcolArticle.colonne-1,TOBL.GetValue('GL_CODEARTICLE'));
           end;
       end;
       end;
   WorkBook.Save; // sauvegarde des données Excel
   // Mise a jour du document PGI Termine
   FINALLY
   FermeDoc;
   END;
end;

procedure TPieceEtude.IndiquePlages (TOBL : TOB);
var Plage,PlageCmp : TPlageExcel;
    chaine ,schaine : string;
    mini,maxi,Indice,position : integer;
begin
chaine := TOBL.GetValue ('VALEUR');
if length (chaine) = 0 then exit;
repeat
      schaine := ReadTokenSt(Chaine);
      if schaine <> '' then
         begin
         Plage := TPlageExcel.create;
         determineMiniMaxi (schaine,mini,maxi);
         Plage.mini := mini;
         Plage.maxi := maxi;
         Plage.feuille  := TOBL.GetValue('FEUILLE');
         position := -1;
         if fplages.Count > 0 then
            begin
            for Indice := 0 to fplages.Count -1 do
                begin
                plagecmp := fplages.items[Indice];
                if plage.feuille < PlageCmp.feuille then BEGIN Position := Indice; break; END;
                if (plage.feuille=PlageCmp.feuille) and (Plage.mini < PlageCmp.mini ) then BEGIN Position :=indice; break;END;
                end;
            end;
         if position = -1 then fPlages.Add (Plage)
                          else fplages.Insert (position,plage);
         end;
until schaine = '';
end;

procedure TPieceEtude.definiPlages (TobStrDoc: TOB);
var TOBL : TOB;
begin
if TOBStrDoc.detail.count = 0 then exit;
TOBL := TOBStrDoc.findfirst(['NATURE'],['000'],false);
if TOBL = nil then exit;
repeat
      IndiquePlages (TOBL);
      TOBL := TOBStrDoc.findnext(['NATURE'],['000'],false);
until TOBL = nil;
end;

procedure TPieceEtude.DefiniColonnes (TOBStrDoc:TOB);
var Indice : integer;
    TOBL : TOB;
    Colonne : string;
begin
{Initialisation}
fColCode.colonne := -1;
fColDes.colonne  := -1;
fColQte.colonne := -1;
fColUnite.colonne := -1;
fColPu.colonne := -1;
fColMontant.colonne := -1;
fColArticle.colonne := -1;
for Indice := 0 to TOBStrDoc.detail.count -1 do
    begin
    TOBL := TOBStrDoc.detail[Indice];
    Colonne := TOBL.GetValue ('VALEUR');
    if TOBL.GetValue('NATURE')='001' then  {Colonne N° Prix}
       begin
       if Colonne <> '' then
          begin
          fcolcode.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcolcode.transfert := (TOBL.GetValue('OBLIG')='X');
          fColCode.colonne := GetIndiceColonne (Colonne);
          end;
       end;
    if TOBL.GetValue('NATURE')='002' then  {Colonne Désignation}
       begin
       if Colonne <> '' then
          begin
          fcoldes.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcoldes.transfert := (TOBL.GetValue('OBLIG')='X');
          fColDes.colonne := GetIndiceColonne (Colonne);
          end;
       end;
    if TOBL.GetValue('NATURE')='003' then  {Colonne Quantité}
       begin
       if Colonne <> '' then
          begin
          fcolQte.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcolQte.transfert := (TOBL.GetValue('OBLIG')='X');
          fColQte.colonne := GetIndiceColonne (Colonne);
          end;
       end;
    if TOBL.GetValue('NATURE')='004' then  {Colonne Unité}
       begin
       if Colonne <> '' then
          begin
          fcolUnite.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcolUnite.transfert := (TOBL.GetValue('OBLIG')='X');
          fColUnite.colonne := GetIndiceColonne (Colonne);
          end;
       end;
    if TOBL.GetValue('NATURE')='005' then  {Colonne Prix Unitaire}
       begin
       if Colonne <> '' then
          begin
          fcolPU.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcolPu.transfert := (TOBL.GetValue('OBLIG')='X');
          fColPu.colonne := GetIndiceColonne (Colonne);
          end;
       end;
    if TOBL.GetValue('NATURE')='006' then  {Colonne Montant}
       begin
       if Colonne <> '' then
          begin
          fcolMontant.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcolMontant.transfert := (TOBL.GetValue('OBLIG')='X');
          fColMontant.colonne := GetIndiceColonne (Colonne);
          end;
       end;
    if TOBL.GetValue('NATURE')='007' then  {Colonne Article}
       begin
       if Colonne <> '' then
          begin
          fcolArticle.transfert := true;
          if TOBL.fieldExists('OBLIG') then fcolArticle.transfert := (TOBL.GetValue('OBLIG')='X');
          fcolArticle.colonne := GetIndiceColonne (Colonne);
    end;
       end;
    end;
end;

procedure TPieceEtude.SupprPrecedent;
begin
{$IFDEF BTP}
BTPSupprimePiece (fcledoc,True);
BTPSupprimePieceFrais (fcledoc);
{$ENDIF}
end;

{*
procedure TPieceEtude.FormatColxls(Colonne: Integer; Format: string) ;
var
  S: String ;
  VColl: OleVariant ;
begin
  if Not varIsEmpty(fCurSheet) then
    begin
    S:=GetCelluleName(Colonne,-1);
    VColl:=WorkSheet.Range[S+':'+S];
    if Not VarIsEmpty(VColl) then VColl.NumberFormat:=Format ;
    end ;
end;
*}

procedure TPieceEtude.QualifiquationColsXls;
begin
(*
if fColQte.colonne > 0 then
   begin
   FormatColxls (fcolQte.colonne,'# ##0,00');
   ExcelAlignCol (WorkSheet,fcolqte.colonne-1,taRightJustify);
   end;
*)
{
if FcolPU.colonne > 0 then
   begin
   FormatColxls (fcolpu.colonne,'# ##0,00');
   ExcelAlignCol (WorkSheet,fcolpu.colonne-1,taRightJustify);
   end;
if FcolMontant.colonne > 0 then
   begin
   FormatColxls (fcolMontant.colonne,'# ##0,00');
   ExcelAlignCol (WorkSheet,fcolMontant.colonne-1,taRightJustify);
   end;
}
end;

procedure TPieceEtude.VersPgi;
var TOBL          : TOB;
    TOBLXLS       : TOB;
    LignePgi      : integer;
    LigneXls      : Integer;
    vide          : integer;
    Indice        : integer;
    Plage         : TPlageExcel;
    NumPrix       : String;
    designation   : String;
    qte           : String;
    unite         : String;
    pu            : String;
    montant       : string;
    Libtmp        : string;
    CodeArticle   : string;
    DocValorise   : boolean;
    LigVide       : Boolean;
    TheWindow     : TToolWindow97;
    Descriptif    : THricheditOle;
begin

  TheWindow := TToolWindow97.create (application);
  TheWindow.Parent := fform;
  TheWindow.visible := false;

  Descriptif := THRichEditOLE.Create(TheWindow);
  Descriptif.Parent := TheWindow;
  descriptif.Visible := false;
  AppliqueFontDefaut (Descriptif);

  DocValorise := not (tModeSaisieBordereau in SaContexte);
  if fcledoc.NumeroPiece = 0 then
  BEGIN
    PGIBox (TraduireMemoire('Document PGI non défini'),'Liaison EXCEL');
    exit;
  END;

  ouvreDocPGI;

  //FV 18-06-2012 gestion modification intégration document XLS
  if not FModifBordereau then ReinitPiece;

  // Reinit Avant traitement
  if fplages.count = 0 then DefiniPlagesWorkbook;
  LignePgi := 1; {Emulation du row d'une grille}

  // génération a travers la structure du fichier Excel
  for Indice := 0 to fPlages.Count -1 do
  begin
    Plage := fPLages.items[Indice];
    fCurSheet := Plage.feuille;
    TRY
      WorkSheet := SelectSheet (WorkBook,fcurSheet);
      QualifiquationColsXls;
      vide := 0;
      for LigneXls := Plage.mini to Plage.maxi do
      begin
        LigVide := false;
        numprix := '';
        designation := '';
        qte := '';
        unite := '';
        pu := '';
        montant := '';
        if vide > 20 then break; // au dela de 16 ligne vide on cosidere que + rien a traiter
        if (fColCode.colonne > 0) and (fcolcode.transfert) then
        begin
          numPrix := GetExcelFormated (WorkSheet,LigneXls,fColCode.colonne);
        end;
        if (fColdes.colonne > 0) and (fcoldes.transfert ) then
        begin
          designation := trimright(trimleft(GetExcelText (WorkSheet,LigneXls,fColDes.colonne)));
        end;
        if (fColQte.colonne  > 0) and (fcolqte.transfert ) then
        begin
          qte := GetExcelText (WorkSheet,LigneXls,fColQte.colonne);
          if pos(',',Qte) > 0 then StringReplace (Qte,',','.',[rfReplaceAll]);
          if Valeur(Qte) = 0 then Qte := '';
        end;
        if (fColUnite.colonne  > 0) and (FcolUnite.transfert ) then
        begin
          unite := GetExcelText (WorkSheet,LigneXls,fColUnite.colonne);
        end;
        if (fColPu.colonne  > 0 ) and (fcolpu.transfert ) then
        begin
          pu := GetExcelText (WorkSheet,LigneXls,fColPu.colonne);
          if pos(',',Pu) > 0 then StringReplace (Pu,',','.',[rfReplaceAll]);
          if Valeur(Pu) = 0 then Pu := '';
        end;
        if (fColMontant.colonne > 0) and (fcolMontant.transfert ) then
        begin
          montant := GetExcelText (WorkSheet,LigneXls,fColMontant.colonne);
          if pos(',',montant) > 0 then StringReplace (montant,',','.',[rfReplaceAll]);
          if Valeur(montant) = 0 then montant := '';
        end;
        if (fColArticle.colonne  > 0 ) and (fColArticle.transfert ) then
        begin
          CodeArticle := GetExcelText (WorkSheet,LigneXls,fColArticle.colonne);
        end;
        if (numprix='') and (designation='') and (qte='') and (unite='') and (pu='') and (montant='') then
        begin
          LigVide := true;
          inc(vide);
        end else
        begin
          vide := 0;
        end;
        //Recherche de la ligne dans la tob piece correspondante à la ligne en cours de traitement
        //si n'existe pas création de cette ligne et traitement normal
        if FModifBordereau then
        begin
         TOBL := TOBpiece.FindFirst(['GL_REFARTTIERS','GL_NUMLIG'], [NumPrix, LignePGI], false);
         if TOBL = nil then
         begin
           TOBL := Newtobligne (TOBPiece,LignePgi);
           InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,lignePgi,false);
         end;
        end
        else
        begin
         TOBL := Newtobligne (TOBPiece,LignePgi);
         InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,lignePgi,false);
        end;

        TOBLXLS := TOB.Create ('BLIENEXCEL',TOBLienXls,-1);
        TOBLXLS.putvalue('BLX_PIECEASSOCIEE',EncodeRefPiece (TOBPiece));
        TOBLXLS.putvalue('BLX_NUMLIGNE',TOBL.getValue('GL_NUMORDRE'));
        TOBLXLS.putvalue('BLX_NUMLIGXLS',LigneXls);
        TOBLXLS.PutValue('BLX_FEUILLEXLS',plage.feuille);
        if ligvide then
        begin
          inc(lignePgi);
          continue;
        end;
        if (NumPrix = '') and (Designation <> '') and ((Qte<>'') or (Pu <> '') or (Unite <> '')) then
        begin
          NumPrix := Copy(Designation,1,18);
          if pos(#$A,NumPrix) > 0 then Numprix := copy (NumPrix,1,pos(#$A,numPrix)-1);
        end;
        //FV1 => gestion avant calcul du cotraitant.
        if (ffournisseur <> '') then
        begin
          if (fNatureTravail <> '') then TOBL.AddChampSupValeur('GLC_NATURETRAVAIL',fNatureTravail);
          TOBL.putValue('GL_FOURNISSEUR',ffournisseur);
          TOBL.AddChampSupValeur('GLC_NONAPPLICFC','X');
          TOBL.AddChampSupValeur('GLC_NONAPPLICFG','X');
          TOBL.AddChampSupValeur('GLC_NONAPPLICFRAIS','X');
          TOBL.putValue('GL_COEFFC',0.0);
          TOBL.putValue('GL_COEFFR',0.0);
          TOBL.putValue('GL_COEFFG',0.0);
          TOBL.putValue('GL_COEFMARG',1.0);
        end;
        //
        if (NumPrix<>'') then
        begin
          if tModeSaisieBordereau in SaContexte then
          begin
            TOBL.putValue('GL_REFARTTIERS',NumPrix);
          end else
          begin
            TOBL.putValue('GL_REFARTSAISIE',NumPrix);
          end;
          TOBL.SetString('GLC_NUMEROTATION',NumPrix);
          TOBL.SetBoolean('GLC_NUMFORCED',true);
        end;
        if (Qte<>'') or ((Qte='') and ((Pu <> '')  or (Unite <> ''))) then
        begin
          if ((Qte='') and ((Pu <> '')or (Unite <> ''))) then
          begin
            if pos(',',Qte) > 0 then StringReplace (Qte,',','.',[rfReplaceAll]);
            if Valeur(Qte) <> 0 then Qte := '1' else Qte := '0';
          end;
          if Qte <> '' then
          begin
//          if Valeur(Qte) = 0 then Qte := '1'; // cas ou une désignation est présente dans la quantité
            TOBL.putValue('GL_QTEFACT',   valeur(Qte));
            TOBL.putValue('GL_QTERESTE',  valeur(Qte));
            TOBL.putValue('GL_QTESTOCK',  valeur(Qte));
            TOBL.putValue('GL_QTERELIQUAT', 0);
            // --- GUINIER ---
            If CtrlOkReliquat(TOBL, 'GL') then
            begin
              TOBL.putValue('GL_MTRESTE',  (Valeur(Qte) * Valeur(PU)));
              TOBL.putValue('GL_MTRELIQUAT',0);
            end;
                  end;
          //              if Valeur(Qte) <> 0 then ExcelValue (WorkSheet,LigneXls-1,fcolQte.colonne-1,valeur(Qte));
        end;
        //
        // Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
        TraitementArticle(TOBL, CodeArticle, Qte, LignePgi);
        //
        // Modified by f.vautrain 09/11/2017 09:58:00 - FS#2777 - TREUIL - Gestion des libellés/descriptifs dans l'intégration des appels d'offre
        if (designation <> '') then
        begin
          libtmp := StringReplace(designation,#$D#$A,'',[rfReplaceAll]);
          libtmp := StringReplace(libTmp,#$A,'',[rfReplaceAll]);
          TOBL.SetString('GL_LIBELLE',copy(libtmp,1,70));
          //Manque la gestion du paramètre société... Bordel de merde !!!!!
          //FV1 : 15/02/2018 - FS#2908 - MOUTHON, SCETEC : Impossible de créer des paragraphes sur un devis venant d'un appel d'offre
          If getParamSocSecur('SO_BTDESCEQUALLIB', False) then
          Begin
            StringToRich(Descriptif, designation);
            TOBL.PutValue('GL_BLOCNOTE', ExRichToString(Descriptif));
          end;
        end;
        //
        if (unite<>'') then
        begin
          TOBL.putValue('GL_QUALIFQTEVTE',Unite);
        end;
        if (Pu<>'') then
        begin
          if pos(',',pu) > 0 then StringReplace (pu,',','.',[rfReplaceAll]);
          if Valeur(Pu) <> 0 then
          begin
            TraitePrixLigneCourante (TOBL,valeur(Pu));
            ExcelValue (WorkSheet,LigneXls-1,fcolPu.colonne-1,valeur(pu));
            DocValorise := True;
          end;
        end;
        if (Montant <> '' ) then
        begin
          if pos(',',Montant) > 0 then StringReplace (Montant,',','.',[rfReplaceAll]);
          TOBL.putValue('GL_MONTANTHTDEV',Valeur(Montant));
          if Valeur(Montant) <> 0 then
          begin
            ExcelValue (WorkSheet,LigneXls-1,fcolMontant.colonne-1,valeur(Montant));
          end;
        end;
        // pour les etudes secondaires
        if fcodePrest <> '' then
        begin
          if (TOBL.GetValue('GL_QTEFACT')<>0) and
             ((TOBL.GetValue('GL_PUHTDEV')<>0) or (TOBL.GetValue('GL_PUTTCDEV')<>0)) then
          begin
            DocValorise := True;
            // + tard mettre en place le code prestation dans le code article ainsi que le code fournisseur (fcodeprest)
          end;
        end;
        Inc(LignePgi);
      end;
    EXCEPT
    END;
  end;
  NettoieDoc (TobPiece);
  NumeroteLignesGC (nil,TOBPiece);
  if DocValorise then
  begin
    PutValueDetail (TOBPiece,'GP_RECALCULER','X');
    CalculeMontantsDoc (TOBPiece,TOBOuvrage,true,true);
    CalculFacture (TOBAFFAire,TOBPiece,TOBPieceTrait,TOBSSTRAIT,TOBOuvrage,TOBOUvragesP,TOBBases,TOBBasesL,TOBTiers,TOBArticles,TOBPorcs,TOBPieceRG,TOBBasesRG,nil,DEV);
  end;

  ajusteDocument;

  WorkBook.save;

  // Mise a jour du document PGI Termine
  if TRANSACTIONS (MajTobs,1) <> oeok then PGIBox (TraduireMemoire('Erreur dans l''écriture de la pièce'),'Liaison EXCEL');

  FermeDoc;

  Descriptif.Free;
  TheWindow.free;

end;

// Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
procedure TPieceEtude.TraitementArticle(TOBL : TOB; CodeArticle, Qte : string;LignePgi : integer);
Var ArticleDivers : string;
    TOBA          : TOB;
begin

  // ARTICLEDIV
  if CodeArticle <> '' then
    ArticleDivers := CodeArticle
  else
  begin
    if fFournisseur <> '' then
      ArticleDivers := GetParamSoc ('SO_BTARTCOTRAIT')
    else
      ArticleDivers := GetParamSoc ('SO_BTARTICLEDIV');
  end;

  //FV1 - 22/02/2018FS#2895 - DUHALDE - Blocage de l'application lors du traitement d'intégration Appel d'Offre
  if copy(ArticleDivers,0,15) = '' then
  begin
    PGIError('Pas d''article divers en paramètres société', 'Erreur de Paramétrage');
    Exit;
  end;

  ArticleDivers := CodeArticleUnique(ArticleDivers,'','','','','');

  // ARTICLEDIV
  // Correction LS du 20/02/07 pour eviter de coller l'article par defaut dans le cas de bordereau de prix
  if (ArticleDivers <> '') and (valeur(Qte) <> 0) and (not (tModeSaisieBordereau in SaContexte)) then
  begin
    TOBA := findTobArt(ArticleDivers);
    if TOBA <> nil then
    begin
      TOBL.PutValue('GL_ARTICLE',ArticleDivers);
      TOBL.PutValue('GL_TYPELIGNE','ART');
      TOBL.PutValue('GL_TYPEARTICLE',TOBA.GetValue('GA_TYPEARTICLE'));
      TOBL.PutValue('GL_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
      if tModeSaisieBordereau in SaContexte then
      begin
        TOBL.putValue('GL_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE'));
      end
      else
      begin
        if TOBL.GetValue('GL_REFARTSAISIE') = '' then
        begin
          TOBL.PutValue('GL_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE'));
        end;
      end;
      TOBL.PutValue('GL_REFARTBARRE', TOBA.GetValue('GA_CODEBARRE'));
      TOBL.PutValue('GL_PRIXPOURQTE', TOBA.GetValue('GA_PRIXPOURQTE'));
      TOBL.PutValue('GL_TYPEREF', 'ART');
      TOBL.PutValue('GL_RECALCULER','X');
      ArticleVersLigne (TOBPiece,TOBA,TobConds,TOBL,TOBTiers);
      TraiteLesOuvrages(nil,TOBPiece, TOBArticles, TOBOuvrage,nil,nil, LignePGI, False, DEV, true);
    end
    else
    begin
      TraitementArticle(TOBL, '', Qte, LignePgi);
    end;
  end;

end;
{
procedure TPieceEtude.SauveDocPgi;
begin
if TRANSACTIONS (MajTobs,1) <> oeok then PGIBox (TraduireMemoire('Erreur dans l''écriture de la pièce'),'Liaison EXCEL');
end;
}
procedure TPieceEtude.DefiniStructureXls(StructureExcel: string);
var EtudeExt : TEtudesExt;
    TOBStrDoc: TOB;
begin
EtudeExt := TEtudesExt.Create;
TOBStrDoc := TOB.create ('BSTRDOC',nil,-1);
TOBStrDoc.putvalue('BSD_DESCRIPT',StructureExcel);
EtudeExt.createStruct (TOBSTRDOC,nil,-1);
EtudeExt.RempliStruct (TOBStrDoc);
definiPlages (TobStrDoc);
DefiniColonnes (TOBStrDoc);
TOBStrDoc.free;
EtudeExt.free;
end;

procedure TPieceEtude.MajTobs;
begin
V_PGI.ioError := OeOk;
SupprPrecedent; // nettoyage ancien document
TOBLienXls.SetAllModifie (true);

  if FModifBordereau then
    TOBLienXls.InsertOrUpdateDB
  else
    TOBLienXls.InsertDB(nil,true);

// --
TOBAdresses.setAllmodifie (true);
if V_PGI.IoError=oeOk then ValideLesAdresses (ToBPiece,TOBPiece_O,TOBadresses);
TOBOUVrage.SetAllModifie(true);
if V_PGI.IoError = oeOk then ValideLesOuv(TOBOuvrage, TOBPiece);
//if V_PGI.IoError = oeOk then ValideLesOuvPlat (TOBOuvragesP, TOBPiece);
if V_PGI.IoError = oeOk then ValideLesBases(TOBPiece,TobBases,TOBBasesL);
// Correction OPTIMISATION
if V_PGI.IoError = oeOk then ValideLesArticlesFromOuv(TOBarticles,TOBOuvrage);
// --
if V_PGI.IOError = OeOk then ValideLesLignesCompl (TOBpiece,TOBPiece_O);
//
TOBPiece.SetAllModifie (true);
if V_PGI.IoError=oeOk then if not TOBPiece.InsertDBByNivel (true) then V_PGI.IoError := OeUnknown;
//
//if V_PGI.IoError = oeOk then
//begin
//  if not TOBOuvragesP.InsertDBByNivel(false) then V_PGI.IoError := oeUnknown;
//end;
if V_PGI.IoError = oeOk then
begin
  TOBBases.SetAllModifie(true);
  if not TOBBases.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;
if V_PGI.IoError = oeOk then
begin
  TOBBasesL.SetAllModifie(true);
  if not TOBBasesL.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;
if V_PGI.IoError = oeOk then
begin
  TOBEches.SetAllModifie(true);
  if not TOBEches.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;
if V_PGI.IoError = oeOk then
begin
  TOBAnaP.SetAllModifie(true);
  if not TOBAnaP.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;
if V_PGI.IoError = oeOk then
begin
  TOBAnaS.SetAllModifie(true);
  if not TOBAnaS.InsertDB(nil) then V_PGI.IoError := oeUnknown;
end;
//
if V_PGI.IoError=oeOk then MajSousAffaire(TobPiece,nil,'00',taCreat,false,false);
end;

procedure TPieceEtude.VersBordereau ;
var
    Indice,Posdollar,vide : integer;
    TOBL : TOB;
    LigneXls : integer;
    NumPrix,refNumPrix,designation,Prix,quantite,montant : string;
    Plage : TPlageExcel;
    IndiceRef,IndiceDoc : integer;
begin
   TOBL:=Nil;
   if fcledoc.NumeroPiece = 0 then BEGIN PGIBox (TraduireMemoire('Document PGI non défini'),'Liaison EXCEL'); exit; END;
   ouvreDocPGI;
   if TOBPiece.detail.count = 0 then exit;
   if fplages.count = 0 then DefiniPlagesWorkbook;
   if varIsEmpty(WorkBook) then BEGIN PGIBox (TraduireMemoire('Le classeur n''est pas défini'),'Liaison EXCEL'); exit; END;
// génération a travers la structure du fichier Excel
   for Indice := 0 to fplages.count -1 do
       begin
       Plage := fPLages.items[Indice];
       fCurSheet := Plage.feuille;
       WorkSheet := SelectSheet (WorkBook,fcurSheet);
       QualifiquationColsXls;
       Vide := 0;
       //
       for LigneXls := Plage.mini to Plage.maxi do
           begin
           if Vide > 20 then break; {Au dela de 20 ligne vide consecutive on arrete le traitement sur la feuille}
           numprix := '';
           designation := '';
           Prix := '';
           quantite := '';
           montant := '';

           NumPrix := GetExcelText (WorkSheet,LigneXls,fColCode.colonne);
           if (fColdes.colonne > 0) and (fcoldes.transfert ) then
              begin
              designation := GetExcelText (WorkSheet,LigneXls,fColdes.colonne);
              end;
           if (fColQte.colonne > 0) and (fcolQte.transfert ) then
              begin
              quantite := GetExcelText (WorkSheet,LigneXls,fColqte.colonne);
              end;
           if (fColPu.colonne > 0) and (fcolpu.transfert ) then
              begin
              prix := GetExcelText (WorkSheet,LigneXls,fColpu.colonne);
              end;
           if (fColMontant.colonne > 0) and (fColMontant.transfert ) then
              begin
              Montant := GetExcelText (WorkSheet,LigneXls,fColMontant.colonne);
              end;
           if (Numprix = '') and (designation = '') and (Prix = '') then BEGIN inc(Vide);continue; END;
           vide := 0;
           if NumPrix <> '' then
              begin
              IndiceRef := -1;
           // traitement du prix dans le libelle
              for Indicedoc := 0 to TOBPiece.detail.count -1 do
                  begin
                  TOBL := TOBPiece.detail[IndiceDOc];
                  if TOBL.getValue('GL_REFARTSAISIE') = NumPrix then BEGIN IndiceRef:= IndiceDoc; break; END;
                  end;
              TOBL := nil;
              if indiceref <> -1 then
                 begin
                 RefNumPrix := Numprix;
                 TOBL := TOBPiece.detail[Indiceref];
                 if TOBL.GetValue('GL_PUHTDEV') = 0 then // pas de prix indique ..on va voir les lignes suivantes
                    begin
                    for Indicedoc := IndiceRef+1 to TOBPiece.detail.count -1 do
                        begin
                        TOBL := TOBPiece.detail[IndiceDOc];
                        if (TOBL.GetValue('GL_REFARTSAISIE') <> '') and (TOBL.GetValue('GL_REFARTSAISIE')<>RefNumPrix) then
                           begin
                           TOBL := nil;
                           break;
                           end;
                        if (TOBL.getValue('GL_REFARTSAISIE') = '') and (TOBL.GetValue('GL_PUHTDEV') <> 0) then
                           BEGIN
                           break;
                           END;
                        end;
                    end;
                 end;
              end;

           if (TOBL <> Nil) and (pos ('[GL_PUHTDEV]',designation) > 0) then
              begin
              posdollar := pos('[GL_PUHTDEV]',designation);
              designation := copy (designation,1,posdollar-1)+chiffrelettre (TOBL.GetValue('GL_PUHTDEV'),'EURO;CENT');
              ExcelText (WorkSheet,LigneXls-1,fcoldes.colonne-1,designation);
              end;
           if (TOBL <> nil)  and (pos ('[GL_PUHTDEV]',Prix) > 0) then
              begin
              ExcelValue (WorkSheet,LigneXls-1,fcolpu.colonne-1,TOBL.GetValue('GL_PUHTDEV'));
              end;
           end;
       //
       end;
   WorkBook.Save; // sauvegarde des données Excel
   FermeDoc;
end;

procedure TPieceEtude.printXls;
begin
OuvreDocExcelAssocie;
if varIsEmpty(WorkBook) then exit;
TRY
   workbook.printout;
FINALLY
   FermeDocExcelAssocie;
END;
end;

procedure TPieceEtude.printPGI;
begin
if fcledoc.NumeroPiece = 0 then BEGIN PGIBox (TraduireMemoire('Document PGI non défini'),'Liaison EXCEL'); exit; END;
ouvreDocPGI;
if TOBPiece.detail.count = 0 then exit;
TRY
   ImprimerLaPiece (TOBPiece,TOBTiers,fcledoc);
FINALLY
   FermeDoc;
END;
end;

procedure TPieceEtude.DefiniPlagesWorkbook;
var Indice: integer;
    Plage : TPlageExcel;
begin
if varIsEmpty(WorkBook) then BEGIN PGIBox (TraduireMemoire('Le classeur n''est pas défini'),'Liaison EXCEL'); exit; END;
for Indice := 1 to WorkBook.sheets.count do
    begin
    Plage := TPlageExcel.create;
    Plage.feuille := workbook.sheets.item[Indice].name;
    Plage.mini := 1;
    Plage.maxi := WorkBook.sheets.item[Indice].rows.count;
    fPlages.Add (Plage); // on indique la plage dans la liste
    end;
end;

procedure TPieceEtude.IncidenceTiers(TOBPiece, TOBTiers,TOBAdresses: TOB);
var OkE,OkOut : boolean;
    CodeD : string;
    DEV : Rdevise;
begin
TOBPIece.PutValue('GP_ESCOMPTE',TOBTiers.GetValue('T_ESCOMPTE')) ;
TOBPiece.PutValue('GP_REMISEPIED',TOBTiers.GetValue('T_REMISE')) ;
TOBPiece.putValue('GP_MODEREGLE',TOBTiers.GetValue('T_MODEREGLE')) ;
TOBPIECE.putValue('GP_REGIMETAXE',TOBTiers.GetValue('T_REGIMETVA')) ;
if TOBTiers.FieldExists('T_REPRESENTANT') then TOBPIece.putValue('GP_REPRESENTANT',TOBTiers.GetValue('T_REPRESENTANT'));
TOBPIece.putValue('GP_FACTUREHT',TOBTiers.GetValue('T_FACTUREHT')) ;
TOBPiece.PutValue('GP_DEVISE',TOBTiers.GetValue('T_DEVISE')) ;
TiersVersPiece(TOBTiers,TOBPiece) ;
TiersVersAdresses(TOBTiers,TOBAdresses,TOBPiece) ;
if Not GetParamSoc('SO_GCPIECEADRESSE') then AffaireVersPiece (TOBPIECE,TOBAFFAIRE);
if ForceEuroGC then OkE:=True else OkE:=(TOBTiers.GetValue('T_EURODEFAUT')='X') ;
CodeD:=TOBPIece.GetValue('GP_DEVISE');
OkOut:=((CodeD<>'') and (CodeD<>V_PGI.DevisePivot) and (CodeD<>V_PGI.DeviseFongible) and (Not EstMonnaieIn(CodeD))) ;
if OkOut then
   if OkE then
   BEGIN
   if (not VH^.TenueEuro) then TOBPiece.putValue('GP_SAISIECONTRE','X');
   TOBPIECE.PutValue('GP_DEVISE',V_PGI.DevisePivot );
   END else if CodeD=V_PGI.DevisePivot then
       BEGIN
       if VH^.TenueEuro then
          BEGIN
          TOBPiece.putValue('GP_SAISIECONTRE','X');
          END;
       END ;
DEV.Code:=TOBTiers.GetValue('T_DEVISE');
if (DEV.Code = '') then DEV.Code:=V_PGI.DevisePivot ;
GetInfosDevise(DEV) ;
DEV.Taux:= GetTaux(DEV.Code,DEV.DateTaux,CleDoc.DatePiece) ;
TOBPiece.PutValue('GP_DEVISE',DEV.Code) ;
TOBPiece.PutValue('GP_TAUXDEV',DEV.Taux) ;
TOBPiece.PutValue('GP_DATETAUXDEV',CleDoc.DatePiece) ;
end;

procedure TPieceEtude.ReinitPiece;
begin
TOBPiece.clearDetail;  ZeroFacture (TOBpiece);
TOBBases.ClearDetail;
TOBBasesL.ClearDetail;
TOBEches.ClearDetail;
TOBPorcs.ClearDetail;
TOBArticles.ClearDetail;
TOBConds.ClearDetail;
TOBTarif.ClearDetail;
TOBComms.ClearDetail;
TOBCatalogu.ClearDetail;
TOBCXV.ClearDetail;
TOBNomen.ClearDetail;
TOBDim.ClearDetail;
TOBLOT.ClearDetail;
TOBSerie.ClearDetail;
TOBSerRel.ClearDetail;
TOBAcomptes.ClearDetail;
TOBDispoContreM.ClearDetail;
// Comptabilité
TOBCPTA.ClearDetail;
TOBANAP.ClearDetail;
TOBANAS.ClearDetail;
//Saisie Code Barres
TOBGSA.ClearDetail;
// Ouvrages
TOBOuvrage.ClearDetail;
// Ouvrages
TOBOuvragesP.ClearDetail;
// textes debut et fin
TOBLIENOLE.ClearDetail;
// retenues de garantie
TOBPieceRG.ClearDetail;
// Bases de tva sur RG
TOBBasesRG.ClearDetail;
TOBLienXls.ClearDetail;
end;


Function TPieceEtude.NewPieceVide ( Var CleDoc : R_CleDoc ; CodeTiers,CodeAffaire,Etablissement,Domaine : String; EnHT,SaisieContre : Boolean ) : TOB ;
Var TOBTiers : TOB ;
    DEV               : RDevise ;
//    Aff0,Aff1,Aff2,Aff3,Aff4 : string;
    ValNumP : T_ValideNumPieceVide;
    io   : TIoErr ;
BEGIN
Result:=nil ;
// Initialisations
if (Etablissement='') then Etablissement :=VH^.EtablisDefaut;
CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,Etablissement,Domaine) ;

if CleDoc.NumeroPiece = 0 then
   begin
   ValNumP := T_ValideNumPieceVide.Create;
   ValNumP.Souche := CleDoc.Souche;
   io:=Transactions(ValNumP.ValideNumPieceVide,5) ;
   if io=oeOk then CleDoc.NumeroPiece := ValNumP.NewNum else CleDoc.NumeroPiece := 0;
   ValNumP.Free;
   end;

if CleDoc.NumeroPiece<=0 then
    BEGIN
    PGIBox (TraduireMemoire('Souche associée à la nature de pièce non valide'),'Création de pièce');
    Exit ;
    END;
CleDoc.Indice:=0 ;
InitTOBPiece(TOBPiece) ;
// Divers
TOBPiece.PutValue('GP_ETABLISSEMENT',Etablissement) ;
TOBPiece.PutValue('GP_DEPOT',VH_GC.GCDepotDefaut) ;
TOBPiece.PutValue('GP_REGIMETAXE',VH^.RegimeDefaut) ;
TOBPiece.PutValue('GP_VENTEACHAT',GetInfoParPiece(CleDoc.NaturePiece,'GPP_VENTEACHAT'));
if EnHT then TOBPiece.PutValue('GP_FACTUREHT','X')
        else TOBPiece.PutValue('GP_FACTUREHT','-');
if SaisieContre then TOBPiece.PutValue('GP_SAISIECONTRE','X')
                else TOBPiece.PutValue('GP_SAISIECONTRE','-');
// Tiers
TOBTiers:=TOB.Create('TIERS',Nil,-1) ; TOBTiers.AddChampSup('RIB',False) ;
RemplirTOBTiers (TOBTiers,CodeTiers,CleDoc.NaturePiece,False) ;
TOBPiece.PutValue('GP_TIERS',CodeTiers) ;
TiersVersPiece(TOBTiers,TOBPiece) ;
// Entête
TOBPiece.PutValue('GP_NATUREPIECEG',CleDoc.NaturePiece) ;
TOBPiece.PutValue('GP_SOUCHE',CleDoc.Souche) ;
TOBPiece.PutValue('GP_NUMERO',CleDoc.NumeroPiece) ;
TOBPiece.PutValue('GP_DATEPIECE',CleDoc.DatePiece) ;
TOBPiece.PutValue('GP_AFFAIRE',CodeAffaire) ;
{$IFDEF BTP}
{$ELSE IF AFFAIRE}
if CodeAffaire <> '' then
    BEGIN
    CodeAffaireDecoupe(CodeAffaire,Aff0,Aff1,Aff2,Aff3,Aff4, taCreat, false);
    TOBPiece.PutValue('GP_AFFAIRE1',Aff1) ; TOBPiece.PutValue('GP_AFFAIRE2',Aff2) ;
    TOBPiece.PutValue('GP_AFFAIRE3',Aff3) ; TOBPiece.PutValue('GP_AVENANT',Aff4) ;
    END;
{$ENDIF}
// Devise
DEV.Code:=TOBTiers.GetValue('T_DEVISE');
if (DEV.Code = '') then DEV.Code:=V_PGI.DevisePivot ;
GetInfosDevise(DEV) ;
DEV.Taux:= GetTaux(DEV.Code,DEV.DateTaux,CleDoc.DatePiece) ;
TOBPiece.PutValue('GP_DEVISE',DEV.Code) ;
TOBPiece.PutValue('GP_TAUXDEV',DEV.Taux) ;
TOBPiece.PutValue('GP_DATETAUXDEV',CleDoc.DatePiece) ;
// Pièce associée à une affaire
if (GetInfoParPiece(CleDoc.NaturePiece,'GPP_PREVUAFFAIRE')='X') Then
   begin
   TOBPiece.PutValue('GP_MAJLIBRETIERS','XXXXXXXXXXX') ;
   end;
// Enregistrement
if TOBTiers <> nil then TOBTiers.free; // ca peux servir !!
Result:=TOBPiece;
END ;


function TPieceEtude.findTobArt (ArticleDivers : string) : TOB;
var TOBA : TOB;
    QQ : Tquery;
begin
  result := nil;
  TOBA := TOBArticles.FindFirst (['GA_ARTICLE'],[ArticleDivers],true);
  if TOBA = nil then
  begin
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+ArticleDivers+'"',true,-1, '', True);
    if not QQ.eof then
    begin
      TOBA := TOB.create ('ARTICLE',TOBArticles,-1);
      TOBA.SelectDB ('',QQ);
      InitChampsSupArticle (TOBA);
      result := TOBA;
    end;
    ferme (QQ);
  end else result := TOBA;
end;

procedure TPieceEtude.ajusteDocument;
var indice : integer;
    TOBL : TOB;
    IndiceOuv : integer;
begin
  {$IFDEF BTP}
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBpiece.detail[indice];
    IndiceOuv := TOBL.GetValue('GL_INDICENOMEN');
    if (IndiceOuv > 0) and (IsOuvrage(TOBL)) then ReAffecteLigOuv(IndiceOuv, TobL, TobOuvrage);
    // protection des coefs de marge, fg, fc , FR pour la validation : cf pb plumard rzzzz rrzzz Avril 2010
    ProtectionCoef  (TOBL);
  end;
  {$ENDIF}
end;

procedure TPieceEtude.ChargeAffaireEtude;
var QQ : TQuery;
		req : string;
begin
	Req := 'SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE = "'+fCodeAffaire+'"';
	QQ := OpenSQL (Req,True,-1, '', True);
  if not QQ.eof then TobAffaire.SelectDB ('',QQ);
  ferme (QQ);
end;

procedure TPieceEtude.ChargeClefAffairePiece;
var part0,part1,part2,part3,avenant : string;
begin
	BTPCodeAffaireDecoupe (fCodeAffaire ,Part0,Part1,Part2,Part3,Avenant,tamodif,false);
  TOBPiece.putValue('GP_AFFAIRE',fcodeaffaire);
  TOBPiece.putValue('GP_AFFAIRE1',Part1);
  TOBPiece.putValue('GP_AFFAIRE2',Part2);
  TOBPiece.putValue('GP_AFFAIRE3',Part3);
  TOBPiece.putValue('GP_AVENANT',Avenant);
end;

procedure TPieceEtude.SetModeBordereau (value : boolean);
begin
	if Value then SaContexte := SaContexte+[tModeSaisieBordereau] else SaContexte := SaContexte-[tModeSaisieBordereau]
end;

procedure TPieceEtude.TraitePrixLigneCourante (TOBL : TOB;Pu : double);
var TypeArticle : string;
		EnHt : boolean;
begin
//
	TypeArticle := TOBL.getValue('GL_TYPEARTICLE');
  EnHt := (TOBPiece.getValue('GP_FACTUREHT')='X');
  if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
  begin
  	TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL,TOBOuvrage, EnHt, Pu,0,DEV,false,False,True,GetParamSocSecur('SO_BTAPPOFFPAEQPV',false));
		ReinitCoefMarg (TOBL,TOBOuvrage);
    if tModeSaisieBordereau in SaContexte then TOBL.PutValue('GL_REMISABLELIGNE','X');
  end;
//
  if EnHt then
  begin
    TOBL.PutValue('GL_PUHTDEV',Pu);
    TOBL.putValue('GL_PUHT',DEVISETOPIVOTEx (Pu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
    TOBL.PutValue('GL_PUHTNETDEV',Pu);
    TOBL.putValue('GL_PUHTNET',DEVISETOPIVOTEx (Pu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
  end else
  begin
    TOBL.PutValue('GL_PUTTCDEV',Pu);
    TOBL.putValue('GL_PUTTC',DEVISETOPIVOTEx (Pu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
    TOBL.PutValue('GL_PUTTCNETDEV',Pu);
    TOBL.putValue('GL_PUTTCNET',DEVISETOPIVOTEx (Pu,DEV.Taux,DEV.Quotite,V_PGI.okdecP));
  end;
  if GetParamSocSecur('SO_BTAPPOFFPAEQPV',false) then
  begin
    TOBL.PutValue('GL_DPA',Pu);
    TOBL.SetDouble('GL_COEFFG',0);
    TOBL.SetDouble('GL_COEFMARG',1);
  end;
//
	TOBL.putValue('GL_BLOQUETARIF','X');
end;

function TPieceEtude.IsDocumentvalorise: Boolean;
var Indice        : integer;
    LigneXls      : Integer;
    vide          : integer;
    Plage         : TPlageExcel;
    qte           : String;
    pu            : string;
    Descriptif    : THricheditOle;
begin

  result := False;

  Descriptif := THRichEditOLE.Create(application);
  Descriptif.Parent := Application.MainForm;
  descriptif.Visible := false;
  AppliqueFontDefaut (Descriptif);

  {*if fcledoc.NumeroPiece = 0 then
  BEGIN
    PGIBox (TraduireMemoire('Document PGI non défini'), 'Liaison EXCEL');
    exit;
  END;*}

  ouvreDocPGI;

  ReinitPiece;

// Reinit Avant traitement
  if fplages.count = 0 then DefiniPlagesWorkbook;

// génération a travers la structure du fichier Excel
  for Indice := 0 to fPlages.Count -1 do
  begin
    Plage := fPLages.items[Indice];
    fCurSheet := Plage.feuille;
    WorkSheet := SelectSheet (WorkBook,fcurSheet);
    QualifiquationColsXls;
    vide := 0;
    for LigneXls := Plage.mini to Plage.maxi do
    begin
      qte := '';
      pu := '';
      if vide > 20 then break; // au dela de 16 ligne vide on cosidere que + rien a traiter
      if (fColQte.colonne  > 0) and (fcolqte.transfert ) then
        qte := GetExcelText (WorkSheet,LigneXls,fColQte.colonne);
      if (fColPu.colonne  > 0 ) and (fcolpu.transfert ) then
        pu := GetExcelText (WorkSheet,LigneXls,fColPu.colonne);
      if Valeur(Qte) = 0 then Qte := '';
      if (qte='') and (pu='') then
      begin
        inc(vide);
        continue;
      end;
      if (qte <> '') and (pu = '') then
      begin
        result := False;
        Break;
      end;
      Result := True;
    end;
  end;

  FermeDoc;
  Descriptif.Free;

end;


procedure TPieceEtude.NettoieDoc(TobPiece: TOB);
var Indice : integer;
		Okok : boolean;
begin
  okok := false;
  repeat
    if (TOBpiece.detail[Tobpiece.detail.count -1].getString('GL_TYPELIGNE')='COM') and
    	 (TOBpiece.detail[Tobpiece.detail.count -1].getString('GL_LIBELLE')='') then
    begin
			TOBpiece.detail[Tobpiece.detail.count -1].free;
    end else break;
  until okok;
end;

end.
