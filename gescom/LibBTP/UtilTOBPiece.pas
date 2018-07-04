unit UtilTOBPiece;

interface

uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  uhttp,paramsoc,
  utob,uEntCommun,
  UtilsTOB;
type
  Tdocument = class (Tobject)
  	cledoc : r_cledoc;
  end;

function CleDocToString(CleDoc: R_CleDoc): string;
procedure DecodeRefPiece(St: string; var CleDoc: R_CleDoc); { NEWPIECE }
function IsOuvrageOkInPiece (naturepiece : string) : Boolean;
function  IsTransfert(NaturePiece : String): Boolean;
procedure LoadLignes(CleDocLigne: R_CleDoc; TobPiece: Tob; WithLigneCompl: Boolean = True; QueLaligne : boolean=false; WithLigneFac:boolean=false; FromExcel : boolean=false);
function LoadPiece (Cledoc : r_cledoc; TOBPiece : TOB) : boolean;
function ChargelaPieceEtUneLigne (CleDocLigne: R_CleDoc; TobPiece,TOBOuvrage: Tob) : boolean;
procedure LoadPieceLignes(CleDocLigne: R_CleDoc; TobPiece: Tob; WithLigneCompl: Boolean = True;QueLaLigne : boolean=false; FromExcel : boolean=false; duplication : Boolean=false);
function MakeSelectLigneBtp (WithLigneCOmpl : boolean;WithEntetePiece : boolean = false;WithLigneFac : boolean=false;FromExcel : boolean = false;Document : Tdocument = nil) : Widestring;
function MakeSelectLigneOuvBtp (WithLigneFac : boolean=false; NaturepieceG : string=''): string;
procedure StringToCleDoc(StA: string; var CleDoc: R_CleDoc);
function  WherePiece(CleDoc: R_CleDoc; ttd: T_TableDoc; Totale: boolean; WithNumOrdre: Boolean = False): string; { NEWPIECE }
procedure ConstitueLesLignes (TOBListPieces,TOBPiece,TOBArticles,TOBAFormule,TOBConds  : TOB; CledocAffaire,Cledoc : R_cledoc; VenteAchat : string);
function isExerciceClo (DateDoc : TDateTime) : boolean;
function ExisteSituationProv (Avancement : Boolean; NaturePieceg,Souche: string; Numero,IndiceG : integer;
                              var NatS,SoucheS : string; var NumS,indiceS : integer) : boolean;
procedure LoadLesVTECOLLECTIF (CleDoc : r_cledoc ;TOBVTECOLLECTIF : TOB);


implementation
uses factcomm,FactTOB,FactUtil,BTPUtil,EntGC;


function ExisteSituationProv (Avancement : Boolean; NaturePieceg,Souche: string; Numero,IndiceG : integer;
                              var NatS,SoucheS : string; var NumS,indiceS : integer) : boolean;
var QQ : TQuery;
    SQl,RefDossier,SqlDoc : String;
begin
  NatS := ''; SoucheS := '';
  NumS := 0; IndiceS := 0;
  Result := false;
  // on recup le devis principal de facturation s'il existe bien sur un dossier de facturation
  SQl := 'SELECT M1.BMF_NATUREPIECEG,M1.BMF_SOUCHE,M1.BMF_NUMERO,M1.BMF_INDICEG FROM BTMEMOFACTURE M1 WHERE M1.BMF_DEVISPRINC = '+
         '(SELECT M2.BMF_DEVISPRINC FROM BTMEMOFACTURE M2 WHERE '+
         'M2.BMF_NATUREPIECEG="'+NaturePieceg+'" AND '+
         'M2.BMF_SOUCHE="'+Souche+'" AND '+
         'M2.BMF_NUMERO='+InttOstr(Numero)+' AND '+
         'M2.BMF_INDICEG='+InttOstr(IndiceG)+')';

  QQ := OpenSql (SQl,true);
  if not QQ.eof then
  begin
    NatS := QQ.Fields[0].AsString;
    SoucheS := QQ.Fields[1].AsString;
    NumS := QQ.Fields[2].AsInteger;
    IndiceS := QQ.Fields[3].AsInteger;
  end;
  ferme (QQ);
  if (NatS = '') then Exit;
  SQlDoc := 'SELECT GP_AFFAIREDEVIS FROM PIECE WHERE '+
            'GP_NATUREPIECEG="'+NatS+'" AND '+
            'GP_SOUCHE="'+SoucheS+'" AND '+
            'GP_NUMERO='+IntToStr(NumS)+' AND '+
            'GP_INDICEG='+InttoStr(IndiceS);
  //
  QQ := OpenSQL('SELECT 1 FROM PIECE  WHERE GP_NATUREPIECEG="FBP" AND '+
                'GP_AFFAIREDEVIS=('+SQlDoc+') AND '+
                'GP_VIVANTE="X"',true,1,'',true);
  Result := (not QQ.Eof);
  Ferme(QQ);
end;


function isExerciceClo (DateDoc : TDateTime) : boolean;
var Sql : string;
    QQ : TQuery;
begin
  Result := false;
  Sql := 'SELECT EX_ETATCPTA FROM EXERCICE WHERE '+
         'EX_DATEDEBUT <="'+USDAteTime(DateDoc)+ '" AND '+
         'EX_DATEFIN >="'+USDAteTime(DateDoc)+ '"';
  QQ := OpenSql (Sql,True,1,'',true);
  if not QQ.eof then
  begin
    Result := (QQ.fields[0].AsString ='CDE');
  end;
end;


function LoadPiece (Cledoc : r_cledoc; TOBPiece : TOB) : boolean;
var
  Q: TQuery;
  Req : string;
begin
  { Chargement entête de la pièce }
  Req := 'SELECT PIECE.*';
  req := req + ',AFF_GENERAUTO,AFF_OKSIZERO,AFF_ETATAFFAIRE AS ETATDOC,'+
  						 '(SELECT COUNT(GLC_NATURETRAVAIL) from LIGNECOMPL '+
               'WHERE GLC_NATUREPIECEG=GP_NATUREPIECEG AND GLC_SOUCHE=GP_SOUCHE AND GLC_NUMERO=GP_NUMERO '+
               'AND GLC_NATURETRAVAIL IN ("002","011","012")) AS SSTRAITE, '+
               '(SELECT BST_NUMEROSIT FROM BSITUATIONS '+
                 'WHERE '+
                 'BST_NATUREPIECE=GP_NATUREPIECEG AND '+
                 'BST_SOUCHE=GP_SOUCHE AND '+
                 'BST_NUMEROFAC=GP_NUMERO '+
               ') AS NUMEROSIT,'+
               '(SELECT MAX(BT2_UNIQUE) FROM BTRFENTETE '+
               ' WHERE '+
               ' BT2_NATUREPIECEG=GP_NATUREPIECEG AND '+
               ' BT2_SOUCHE=GP_SOUCHE AND '+
               ' BT2_NUMERO=GP_NUMERO AND '+
               ' BT2_INDICEG=GP_INDICEG '+
               ') AS MAXUNIQUE '+
  						 'FROM PIECE '+
               'LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS ';
  Req := Req + 'WHERE ' + WherePiece(Cledoc, ttdPiece, False);
  Q := OpenSQL(Req, True,-1, '', True);
  result := not Q.eof;
  TobPiece.SelectDB('', Q);
  Ferme(Q);
end;

function ChargelaPieceEtUneLigne (CleDocLigne: R_CleDoc; TobPiece,TOBOuvrage: Tob) : boolean;


  Procedure DefinirLesOuv (TOBGroupeNomen,TOBNomen : TOB ; LaLig,idep : integer) ;
  Var i,Lig : integer ;
  TOBLN,TOBPere,TOBLoc,TOBArt,TOBRef : TOB ;
  RefArticle,TypeArticle : String ;
  LigPere1,LigPere2,LigPere3,LigPere4 : integer;
  LigneN1,LigneN2,LigneN3,LigneN4,LigneN5: integer;
  BEGIN
    for i:=idep to TOBNomen.Detail.Count-1 do
    BEGIN
      TOBLN:=TOBNomen.Detail[i] ;
      Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
      if lig <> LaLig then break;
      LigPere1 := 0; LigneN1 := TOBLN.GetValue('BLO_N1');
      LigPere2 := 0; LigneN2 := TOBLN.GetValue('BLO_N2');
      LigPere3 := 0; LigneN3 := TOBLN.GetValue('BLO_N3');
      LigPere4 := 0; LigneN4 := TOBLN.GetValue('BLO_N4');
      LigneN5 := TOBLN.GetValue('BLO_N5');
      if LigneN5 > 0 then
      begin
        // recherche du pere au niveau 4
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
      end else if TOBLN.GetValue('BLO_N4') > 0 then
      begin
        // recherche du pere au niveau 3
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
      end else if TOBLN.GetValue('BLO_N3') > 0 then
      begin
        // recherche du pere au niveau 2
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
      end else if TOBLN.GetValue('BLO_N2') > 0 then
      begin
        // recherche du pere au niveau 1
        TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,0,0,0,0],True) ;
      end else TOBPere:=TOBGroupeNomen;
      if TOBPere<>Nil then
      BEGIN
        TOBLoc:=TOB.Create('LIGNEOUV',TOBPere,-1) ;
        TOBLoc.Dupliquer(TOBLN,False,True) ;
      END ;

    END ;
  END ;

var
  Q: TQuery;
  oldL,I,Lig,IndiceNomen : Integer;
  Req : string;
  TOBNomen,TOBL,TOBLN,TOBGroupeNomen : TOB;
begin
  TOBNomen := TOB.Create('LES OUV',nil,-1);
  { Chargement entête de la pièce }
  Req := 'SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_DATEMODIF,GP_VIVANTE ' +
  			 'FROM PIECE '+
  			 'WHERE ' + WherePiece(CleDocLigne, ttdPiece, False);
  Q := OpenSQL(Req, True,1, '', True);
  result := not Q.eof;
  if not Q.Eof then TobPiece.SelectDB('', Q);
  Ferme(Q);
  if not result then Exit;
  //
  Req := 'SELECT GL_NATUREPIECEG,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMORDRE,GL_QTEFACT,GL_NUMLIGNE,'        +
  			 'GL_QTERESTE,GL_QTERELIQUAT,GL_QTESTOCK,GL_DEPOT,GL_TENUESTOCK,GL_ARTICLE,GL_CODEARTICLE,'         +
         'GL_COEFCONVQTE,GL_COEFCONVQTEVTE,GL_QUALIFQTEACH,GL_QUALIFQTESTO,GL_QUALIFQTEVTE,GL_INDICENOMEN,' +
         'GL_TYPEREF,GL_REFCATALOGUE,GL_TYPEDIM,GL_TYPELIGNE,GL_TYPEARTICLE,GL_DPA,GL_PMAP,GL_DPR,GL_PMRP,' +
         'GL_QTESTOCK AS LIVREORIGINE,GL_PIECEPRECEDENTE,GL_PIECEORIGINE,GL_IDENTIFIANTWOL,GL_AFFAIRE, '    +
         'GL_MTRESTE,GL_MTRELIQUAT,GL_MONTANTHTDEV ' +
         'FROM LIGNE WHERE '+WherePiece (CleDocLigne,ttdLigne ,true,True);
  Q := OpenSQL(Req,True,-1,'',true);
  if not Q.eof then
  begin
		TobPiece.LoadDetailDB('LIGNE','','',Q,false);
  end;
  Ferme(Q);

  if not NaturepieceOKPourOuvrage (TOBPiece)  then exit;
  if GetInfoParPiece (TOBPiece.GetString('GP_NATUREPIECEG'),'GPP_STOCKSSDETAIL')<>'X' then Exit;
  //
  TOBL := TOBPiece.detail[0]; // la ligne
  //
  Req :='SELECT * FROM LIGNEOUV '+
        'WHERE '+WherePiece(CleDocLigne,ttdOuvrage,false)+' AND '+
        'BLO_NUMLIGNE='+IntToStr(TOBL.getInteger('GL_NUMLIGNE'))+
        ' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5';

  Q:=OpenSQL(req,True,-1, '', True) ;
  TOBNomen.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
  Ferme(Q) ;
  OldL := -1;
  for i:=0 to TOBNomen.Detail.Count-1 do
  BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if OldL<>Lig then
    BEGIN
      TOBGroupeNomen:=TOB.Create('',TOBOuvrage,-1) ;
      IndiceNomen := TobOuvrage.detail.count;
      TOBL.PutValue('GL_INDICENOMEN',IndiceNomen);
      DefinirLesOuv(TOBGroupeNomen,TOBNomen,Lig,i);
    END ;
    OldL:=Lig ;
  END ;
  TOBNomen.Free ;
end;

procedure ConstitueLesLignes (TOBListPieces,TOBPiece,TOBArticles,TOBAFormule,TOBConds  : TOB; CledocAffaire,Cledoc : R_cledoc; VenteAchat : string);

	function AddLigneComment (TOBPiece,TOBPREF : TOB) : TOB ;
  var Newl : TOB;
  		RefP : string;
  begin
    NewL:=NewTOBLigne(TOBPiece,-1) ;
    NewL.ClearDetail; // ecite la démultiplication des filles
    NewL.PutValue('GL_NUMORDRE', 0) ;
    //
    RefP:=RechDom('GCNATUREPIECEG',TOBPREF.GetValue('GP_NATUREPIECEG'),False)
       +' N° '+IntToStr(TOBPREF.GetValue('GP_NUMERO'))
       +' du '+DateToStr(TOBPREF.GetValue('GP_DATEPIECE'))
       +'  '+TOBPREF.GetValue('GP_REFINTERNE') ;
    //
    RefP:=Copy(RefP,1,70) ;
    NewL.PutValue('GL_LIBELLE',RefP)    ; NewL.PutValue('GL_TYPELIGNE','COM') ;
    NewL.PutValue('GL_TYPEDIM','NOR')   ; NewL.PutValue('GL_CODEARTICLE','') ;
    NewL.PutValue('GL_ARTICLE','')      ; NewL.PutValue('GL_QTEFACT',0) ;
    NewL.PutValue('GL_QTESTOCK',0)      ; NewL.PutValue('GL_PUHTDEV',0) ;
    NewL.PutValue('GL_QTERESTE',0)      ; { NEWPIECE }
    NewL.PutValue('GL_MTRESTE', 0)      ; { NEWPIECE }
    NewL.PutValue('GL_PUTTCDEV',0)      ; NewL.PutValue('GL_TYPEARTICLE','') ;
    NewL.PutValue('GL_PUHT',0)          ; NewL.PutValue('GL_PUHTNET',0) ;
    NewL.PutValue('GL_PUTTC',0)         ; NewL.PutValue('GL_PUTTCNET',0) ;
    NewL.PutValue('GL_PUHTBASE',0)      ; NewL.PutValue('GL_FAMILLETAXE1','') ;
    NewL.PutValue('GL_TYPENOMENC','')   ; NewL.PutValue('GL_QUALIFMVT','') ;
    NewL.PutValue('GL_REFARTSAISIE','') ; NewL.PutValue('GL_REFARTBARRE','') ;
    NewL.PutValue('GL_REFCATALOGUE','') ; NewL.PutValue('GL_TYPEREF','') ;
    NewL.PutValue('GL_REFARTTIERS','')  ;
    {Modif AC 4/07/03 Pas de GL_CODESDIM sur les lignes commentaire}
    NewL.PutValue('GL_CODESDIM','')  ;
    {Fin Modif AC}
    {Modif JLD 20/06/2002}
    NewL.PutValue('GL_ESCOMPTE',TOBPiece.GetValue('GP_ESCOMPTE')) ;
    NewL.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
    {Fin modif}

    //JS 17/06/03
    NewL.PutValue('GL_INDICESERIE',0) ; NewL.PutValue('GL_INDICELOT',0) ;
    NewL.PutValue('GL_REMISELIGNE',0) ;
    // Modif BTP
    NewL.PutValue('GL_TYPEARTICLE','EPO');
    NewL.PutValue('GL_PUHTNETDEV',0)    ; NewL.PutValue('GL_PUTTCNETDEV',0) ;
    NewL.PutValue('GL_BLOCNOTE','')    ; NewL.PutValue('GL_QUALIFQTEVTE','') ;
    NewL.PutValue('GL_INDICENOMEN',0) ;
    NewL.PutValue('GL_VIVANTE','X')  ;
    // ---
    ZeroLigne(NewL) ;
    Result := Newl;
  end;

  procedure PositionneCommentPourDocument(TOBLD,TOBLL : TOB);
  begin
    TOBLD.setString('GL_TIERS',TOBLL.GetString('GL_TIERS'));
    TOBLD.setString('GL_TIERSLIVRE',TOBLL.GetString('GL_TIERSLIVRE'));
    TOBLD.setString('GL_TIERSFACTURE',TOBLL.GetString('GL_TIERSFACTURE'));
    TOBLD.setString('GL_TIERSPAYEUR',TOBLL.GetString('GL_TIERSPAYEUR'));
    TOBLD.setString('GL_FACTUREHT',TOBLL.GetString('GL_FACTUREHT'));
    TOBLD.setString('GL_TVAENCAISSEMENT',TOBLL.GetString('GL_TVAENCAISSEMENT'));
    TOBLD.setString('GL_DATELIVRAISON',TOBLL.GetString('GL_DATELIVRAISON'));
    TOBLD.setString('GL_REGIMETAXE',TOBLL.GetString('GL_REGIMETAXE'));
    TOBLD.setString('GL_ETABLISSEMENT',TOBLL.GetString('GL_ETABLISSEMENT'));
  end;

  procedure AddDetailPiece (TOBPiece,TOBPP : TOB);
  var II : Integer;
      cledoc : R_CLEDOC;
      TOBLL,NewL,TOBLD : TOB;
      RefPiece : string;
  begin
    Cledoc := TOB2Cledoc (TOBPP);
    LoadLignes(CleDoc, TobPP);
    TOBLD := AddLigneComment (TOBPiece,TOBPP);
    II := 0;
    for II := 0 to TOBPP.detail.count -1 do
    begin
      TOBLL := TOBPP.detail[II];
      if II = 0 then PositionneCommentPourDocument(TOBLD,TOBLL);
      NewL:=NewTOBLigne(TOBPiece,-1) ;
      NewL.Dupliquer(TOBLL,false,true);
      NewL.ClearDetail; // evite la démultiplication des filles
    end;
  end;

  procedure  ChargeArticles(TOBPP,TOBArticles,TOBAFormule,TOBConds  : TOB; CledocAffaire,Cledoc : R_cledoc; VenteAchat : string);
  var TOBARTICLE : TOB;
  		II : Integer;
  begin
		TOBARTICLE := TOB.Create ('LES ART PIECES',nil,-1);
    TRY
      RecupTOBArticle (TOBPiece,TOBARTICLE,TOBAFormule,TOBConds,CledocAffaire,Cledoc,VenteAchat);
      II := 0;
      repeat
        if TOBGetInto (TOBArticles,['GA_ARTICLE'],[TOBARTICLE.detail[II].GetString('GA_ARTICLE')]) = nil then
        begin
          // l'article n'exsite pas dans la tob --> on l'insere
          TOBInsertInto(TOBArticles,TOBARTICLE.detail[II],['GA_ARTICLE']);
        end else inc(ii);
      until (TOBARTICLE.detail.Count=0) or (ii >= TOBARTICLE.detail.Count);
    FINALLY
    	TOBARTICLE.Free;
    END;
  end;

var II : Integer;
		TOBPP : TOB;
    CledocC : R_cledoc;
begin
  for II := 0 to TOBListPieces.detail.count -1 do
  begin
    TOBPP := TOBListPieces.detail[II];
    CledocC := TOB2Cledoc (TOBPP);
    if II = 0 then
    begin
  		LoadPiece (CledocC, TOBPiece);
    end;
    //
    ChargeArticles(TOBPP,TOBArticles,TOBAFormule,TOBConds,CledocAffaire,CledocC ,VenteAchat);
    AddDetailPiece (TOBPiece,TOBPP);
  end;
end;


procedure LoadPieceLignes(CleDocLigne: R_CleDoc; TobPiece: Tob; WithLigneCompl: Boolean = True;QueLaLigne : boolean=false; FromExcel : boolean=false; duplication : Boolean=false);
var WithLigneFac : boolean;
begin
  { Chargement entête de la pièce }
  LoadPiece (CledocLigne, TOBPiece);
  WithLigneFac := (Pos(TOBpiece.GetValue('GP_NATUREPIECEG'),'FBT;DAC;FBP;BAC')>0) and (not duplication);
  LoadLignes(CleDocLigne, TobPiece, WithLigneCompl,QueLaLigne,WithLigneFac,FromExcel);
end;

procedure LoadLesVTECOLLECTIF (CleDoc : r_cledoc ;TOBVTECOLLECTIF : TOB);
var QQ: TQuery;
begin
  QQ := OpenSql ('SELECT * FROM PIEDCOLLECTIF WHERE '+WherePiece (CleDoc,ttdVteColl,true),True,-1,'',true);
  if not QQ.eof then TOBVTECOLLECTIF.LoadDetailDB('PIEDCOLLECTIF','','',QQ,false);
  ferme (QQ); 
end;


procedure LoadLignes(CleDocLigne: R_CleDoc; TobPiece: Tob; WithLigneCompl: Boolean = True; QueLaligne : boolean=false; WithLigneFac:boolean=false; FromExcel : boolean=false);
var
  Q: TQuery;
  Sql: String;
  Document : Tdocument;
begin
  Document := Tdocument.Create;
  Sql := 'SELECT * FROM LIGNE';
{$IFDEF GPAOLIGHT}
  if WithLigneCompl then
    Sql := MakeSelectLigne;
{$ENDIF GPAOLIGHT}
  Document.cledoc := CledocLigne;
	if FromExcel then
  begin
		Sql := MakeSelectLigneBtp (WithLigneCompl,false,WithLigneFac,true,Document);
  end else
  begin
		Sql := MakeSelectLigneBtp (WithLigneCompl,false,WithLigneFac,False,Document);
  end;
  Sql := Sql + ' WHERE ' + WherePiece(CleDocLigne, ttdLigne,QueLaLigne,WithLigneCompl) + ' ORDER BY GL_NUMLIGNE';
  Q := OpenSQL(Sql, True,-1, '', True);
  if not Q.eof then TobPiece.LoadDetailDB('LIGNE', '', '', Q, False, True);
  Ferme(Q);
  Document.free;
end;

function MakeSelectLigneOuvBtp (WithLigneFac : boolean=false; NaturepieceG : string=''): string;
begin
  result :='SELECT O.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,';
  if WithLigneFac then
  begin
    result := result + 'LIGNEFAC.*,BLF_MTSITUATION AS OLD_MTSITUATION, BLF_QTESITUATION AS OLD_QTESITUATION,';
  end else
  begin
    result := result +'LIGNEFAC.BLF_QTEDEJAFACT,';
  end;
  result := result + '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=BLO_FOURNISSEUR) AS LIBELLEFOU,';
  result := result + 'IIF(BLO_COEFMARG=0,0,(BLO_COEFMARG -1) * 100) AS POURCENTMARG,';
  result := result + 'IIF(BLO_PUHT=0,0,((BLO_PUHT-BLO_DPR)/BLO_PUHT)*100) AS POURCENTMARQ,(SELECT GA_RELIQUATMT FROM ARTICLE WHERE GA_ARTICLE=BLO_ARTICLE) AS GESTRELIQUAT,';
  result := result + 'IIF(BLO_PUHT=0,0,((BLO_PUHT-BLO_DPR)/BLO_PUHT)*100) AS POURCENTMARQ, ';
  result := result + '(BLO_QTEFACT * BLO_MTMOPA) AS TOTALMOPA,';
  result := result + '(BLO_QTEFACT * BLO_MTMOPR) AS TOTALMOPR,';
  result := result + '(BLO_QTEFACT * BLO_MTMOPV) AS TOTALMOPV,';
  result := result + '(BLO_QTEFACT * BLO_MTFOUPA) AS TOTALFOUPA,';
  result := result + '(BLO_QTEFACT * BLO_MTFOUPR) AS TOTALFOUPR,';
  result := result + '(BLO_QTEFACT * BLO_MTFOUPV) AS TOTALFOUPV,';
  result := result + '(BLO_QTEFACT * BLO_MTINTPA) AS TOTALINTPA,';
  result := result + '(BLO_QTEFACT * BLO_MTINTPR) AS TOTALINTPR,';
  result := result + '(BLO_QTEFACT * BLO_MTINTPV) AS TOTALINTPV,';
  result := result + '(BLO_QTEFACT * BLO_MTLOCPA) AS TOTALLOCPA,';
  result := result + '(BLO_QTEFACT * BLO_MTLOCPR) AS TOTALLOCPR,';
  result := result + '(BLO_QTEFACT * BLO_MTLOCPV) AS TOTALLOCPV,';
  result := result + '(BLO_QTEFACT * BLO_MTMATPA) AS TOTALMATPA,';
  result := result + '(BLO_QTEFACT * BLO_MTMATPR) AS TOTALMATPR,';
  result := result + '(BLO_QTEFACT * BLO_MTMATPV) AS TOTALMATPV,';
  result := result + '(BLO_QTEFACT * BLO_MTOUTPA) AS TOTALOUTPA,';
  result := result + '(BLO_QTEFACT * BLO_MTOUTPR) AS TOTALOUTPR,';
  result := result + '(BLO_QTEFACT * BLO_MTOUTPV) AS TOTALOUTPV,';
  result := result + '(BLO_QTEFACT * BLO_MTSTPA)  AS TOTALSTPA,';
  result := result + '(BLO_QTEFACT * BLO_MTSTPR)  AS TOTALSTPR,';
  result := result + '(BLO_QTEFACT * BLO_MTSTPV)  AS TOTALSTPV,';
  result := result + '(BLO_QTEFACT * BLO_MTAUTPA) AS TOTALAUTPA,';
  result := result + '(BLO_QTEFACT * BLO_MTAUTPR) AS TOTALAUTPR,';
  result := result + '(BLO_QTEFACT * BLO_MTAUTPV) AS TOTALAUTPV ';
  if (VH_GC.BTCODESPECIF = '001') and (NaturepieceG ='BCE') then
  begin
    Result := Result +','+
                      '('+
                      'SELECT DISTINCT IIF(BT3_TYPELIGNETRF="000" AND BT3_CONTREP="-",0,BT3_UNIQUE) '+
                      'FROM BTRFDETAIL '+
                      'WHERE '+
                      'BT3_NATUREPIECEG=BLO_NATUREPIECEG AND '+
                      'BT3_SOUCHE=BLO_SOUCHE AND '+
                      'BT3_NUMERO=BLO_NUMERO AND '+
                      'BT3_INDICEG=BLO_INDICEG AND '+
                      'BT3_UNIQUEBLO=BLO_UNIQUEBLO AND '+
                      'BT3_TYPELIGNETRF="001" '+
                      ') AS NUMTRANSFERT, ';
    Result := result +
                      '('+
                      'SELECT DISTINCT IIF(BT3_TYPELIGNETRF="000" AND BT3_CONTREP="-","",BT3_CONTREP) '+
                      'FROM BTRFDETAIL '+
                      'WHERE '+
                      'BT3_NATUREPIECEG=BLO_NATUREPIECEG AND '+
                      'BT3_SOUCHE=BLO_SOUCHE AND '+
                      'BT3_NUMERO=BLO_NUMERO AND '+
                      'BT3_INDICEG=BLO_INDICEG AND '+
                      'BT3_UNIQUEBLO=BLO_UNIQUEBLO AND '+
                      'BT3_TYPELIGNETRF="001"'+
                      ') AS TYPETRANSFERT, ';
    Result := result +
                      '('+
                      'SELECT SUM(BT3_MTTRANSFERT) '+
                      'FROM BTRFDETAIL '+
                      'WHERE '+
                      'BT3_NATUREPIECEG=BLO_NATUREPIECEG AND '+
                      'BT3_SOUCHE=BLO_SOUCHE AND '+
                      'BT3_NUMERO=BLO_NUMERO AND '+
                      'BT3_INDICEG=BLO_INDICEG AND '+
                      'BT3_UNIQUEBLO=BLO_UNIQUEBLO'+
                      ') AS MTTRANSFERT ';
  end else
  begin
    Result := Result +',0 AS NUMTRANSFERT,"" AS TYPETRANSFERT,0 AS MTTRANSFERT ';
  end;
  result := result + 'FROM LIGNEOUV O '+
            'LEFT JOIN NATUREPREST N ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=O.BLO_ARTICLE) ';
  if WithLigneFac then
  begin
  	result := result + 'LEFT JOIN LIGNEFAC ON BLF_NATUREPIECEG=BLO_NATUREPIECEG AND BLF_SOUCHE=BLO_SOUCHE AND BLF_NUMERO=BLO_NUMERO '+
    									 'AND BLF_INDICEG=BLO_INDICEG AND BLF_UNIQUEBLO=BLO_UNIQUEBLO AND BLF_UNIQUEBLO <> 0 ';
  end else
  begin
  	result := result + 'LEFT JOIN LIGNEFAC ON BLF_NATUREPIECEG=BLO_NATUREPIECEG AND BLF_SOUCHE=BLO_SOUCHE AND BLF_NUMERO=BLO_NUMERO '+
    									 'AND BLF_INDICEG=BLO_INDICEG AND BLF_UNIQUEBLO=BLO_UNIQUEBLO AND BLF_UNIQUEBLO <> 0 ';
  end;
end;

function MakeSelectLigneBtp (WithLigneCOmpl : boolean;WithEntetePiece : boolean = false;WithLigneFac : boolean=false;FromExcel : boolean = false;Document : Tdocument = nil) : Widestring;
var stdocument,NaturepieceG : string;
begin
  stDocument := '';
  if (fromExcel) and (Document  <> nil) then
  begin
		stDocument := EncodeRefPiece (Document.cledoc);
  end;
  if Document <> nil then NaturepieceG := Document.cledoc.NaturePiece;
  Result := 'SELECT LIGNE.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,';
  if WithLigneCompl then Result := Result + 'LIGNECOMPL.*,';
  if WithLigneFac then result := result + 'LIGNEFAC.*,BLF_MTSITUATION AS OLD_MTSITUATION, BLF_QTESITUATION AS OLD_QTESITUATION,';
  Result := result + 'GL_MONTANTHTDEV AS OLD_MTHTDEV,"X" AS CTRLLIGOK, ';
  if (fromExcel) and (Document  <> nil) then
  begin
    result := result + 'BLIENEXCEL.BLX_NUMLIGXLS,';
  end;
  if WithEntetePiece then result := Result + 'GP_REFINTERNE,GP_NUMADRESSEFACT,GP_NUMADRESSELIVR,';
  result:= result +	 'P.BLP_PHASETRA,P.BLP_NUMMOUV,BCO_LIENTRANSFORME,BCO_INDICE,'
          				+  'BCO_TRANSFORME,BCO_QTEVENTE,BCO_QUANTITE,BCO_LIENVENTE,BCO_TRAITEVENTE,'
          				+  'BCO_QTEINIT,BCO_LIENRETOUR,(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GL_FOURNISSEUR) AS LIBELLEFOU,';
  result := result + 'IIF(GL_COEFMARG=0,0,(GL_COEFMARG -1) * 100) AS POURCENTMARG, ';
  result := result + 'IIF(GL_PUHT=0,0,((GL_PUHT-GL_DPR)/GL_PUHT)*100) AS POURCENTMARQ,';
  result := result + 'IIF(GL_GESTIONPTC="X", GL_MONTANTHTDEV , 0 ) AS MONTANTPTC, (SELECT GA_RELIQUATMT FROM ARTICLE WHERE GA_ARTICLE=GL_ARTICLE) AS GESTRELIQUAT ';
  if (VH_GC.BTCODESPECIF = '001') and (NaturepieceG ='BCE') then
  begin
    Result := Result +','+
                      '('+
                      'SELECT DISTINCT IIF(BT3_TYPELIGNETRF="000" AND BT3_CONTREP="-",0,BT3_UNIQUE) '+
                      'FROM BTRFDETAIL '+
                      'WHERE '+
                      'BT3_NATUREPIECEG=GL_NATUREPIECEG AND '+
                      'BT3_SOUCHE=GL_SOUCHE AND '+
                      'BT3_NUMERO=GL_NUMERO AND '+
                      'BT3_INDICEG=GL_INDICEG AND '+
                      'BT3_NUMORDRE=GL_NUMORDRE AND '+
                      'BT3_TYPELIGNETRF="001" '+
                      ') AS NUMTRANSFERT, ';
    Result := result +
                      '('+
                      'SELECT DISTINCT IIF(BT3_TYPELIGNETRF="000" AND BT3_CONTREP="-","",BT3_CONTREP) '+
                      'FROM BTRFDETAIL '+
                      'WHERE '+
                      'BT3_NATUREPIECEG=GL_NATUREPIECEG AND '+
                      'BT3_SOUCHE=GL_SOUCHE AND '+
                      'BT3_NUMERO=GL_NUMERO AND '+
                      'BT3_INDICEG=GL_INDICEG AND '+
                      'BT3_NUMORDRE=GL_NUMORDRE AND '+
                      'BT3_TYPELIGNETRF="001"'+
                      ') AS TYPETRANSFERT, ';
    Result := result +
                      '('+
                      'SELECT SUM(BT3_MTTRANSFERT) '+
                      'FROM BTRFDETAIL '+
                      'WHERE '+
                      'BT3_NATUREPIECEG=GL_NATUREPIECEG AND '+
                      'BT3_SOUCHE=GL_SOUCHE AND '+
                      'BT3_NUMERO=GL_NUMERO AND '+
                      'BT3_INDICEG=GL_INDICEG AND '+
                      'BT3_NUMORDRE=GL_NUMORDRE '+
                      ') AS MTTRANSFERT, ';
    Result := Result +
                      '('+
                'SELECT SUM(BLE_MONTANT) '+
                'FROM BLIGNEETS '+
                'WHERE '+
                'BLE_NATUREPIECEG=GL_NATUREPIECEG AND BLE_SOUCHE=GL_SOUCHE AND BLE_NUMERO=GL_NUMERO AND BLE_INDICEG=GL_INDICEG AND BLE_NUMORDRE=GL_NUMORDRE'+
              ') AS SUMTOTALTS ';
  end else
  begin
    Result := Result +',0 AS NUMTRANSFERT,"" AS TYPETRANSFERT,0 AS MTTRANSFERT, 0 AS SUMTOTALTS ';
  end;
  if WithLigneCompl Then
  begin
    result := result + ',IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTMOPA,(GL_QTEFACT * GLC_MTMOPA)) AS TOTALMOPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTMOPR,(GL_QTEFACT * GLC_MTMOPR)) AS TOTALMOPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTMOPV,(GL_QTEFACT * GLC_MTMOPV)) AS TOTALMOPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTFOUPA,(GL_QTEFACT * GLC_MTFOUPA)) AS TOTALFOUPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTFOUPR,(GL_QTEFACT * GLC_MTFOUPR)) AS TOTALFOUPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTFOUPV,(GL_QTEFACT * GLC_MTFOUPV)) AS TOTALFOUPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTINTPA,(GL_QTEFACT * GLC_MTINTPA)) AS TOTALINTPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTINTPR,(GL_QTEFACT * GLC_MTINTPR)) AS TOTALINTPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTINTPV,(GL_QTEFACT * GLC_MTINTPV)) AS TOTALINTPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTLOCPA,(GL_QTEFACT * GLC_MTLOCPA)) AS TOTALLOCPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTLOCPR,(GL_QTEFACT * GLC_MTLOCPR)) AS TOTALLOCPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTLOCPV,(GL_QTEFACT * GLC_MTLOCPV)) AS TOTALLOCPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTMATPA,(GL_QTEFACT * GLC_MTMATPA)) AS TOTALMATPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTMATPR,(GL_QTEFACT * GLC_MTMATPR)) AS TOTALMATPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTMATPV,(GL_QTEFACT * GLC_MTMATPV)) AS TOTALMATPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTOUTPA,(GL_QTEFACT * GLC_MTOUTPA)) AS TOTALOUTPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTOUTPR,(GL_QTEFACT * GLC_MTOUTPR)) AS TOTALOUTPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTOUTPV,(GL_QTEFACT * GLC_MTOUTPV)) AS TOTALOUTPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTSTPA,(GL_QTEFACT * GLC_MTSTPA))  AS TOTALSTPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTSTPR,(GL_QTEFACT * GLC_MTSTPR))  AS TOTALSTPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTSTPV,(GL_QTEFACT * GLC_MTSTPV))  AS TOTALSTPV,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTAUTPA,(GL_QTEFACT * GLC_MTAUTPA)) AS TOTALAUTPA,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTAUTPR,(GL_QTEFACT * GLC_MTAUTPR)) AS TOTALAUTPR,';
    result := result + 'IIF(SUBSTRING(GL_TYPELIGNE,1,2)="TP",GLC_MTAUTPV,(GL_QTEFACT * GLC_MTAUTPV)) AS TOTALAUTPV ';
  end;
  result := result + 'FROM LIGNE '
                  +  'LEFT JOIN LIGNEPHASES P '
                  +  'ON (P.BLP_NATUREPIECEG = GL_NATUREPIECEG and P.BLP_SOUCHE=GL_SOUCHE AND P.BLP_NUMERO = GL_NUMERO '
                  +  'and P.BLP_INDICEG = GL_INDICEG and P.BLP_NUMORDRE = GL_NUMORDRE) ';
  result := result + 'LEFT JOIN NATUREPREST N ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=LIGNE.GL_ARTICLE) ';
  if WithLigneCompl Then
  begin
  	result := result + 'LEFT JOIN LIGNECOMPL '
          					 + 'ON (GL_NATUREPIECEG = GLC_NATUREPIECEG and GL_SOUCHE = GLC_SOUCHE AND GL_NUMERO = GLC_NUMERO '
                     + 'AND GL_INDICEG = GLC_INDICEG and GL_NUMORDRE = GLC_NUMORDRE)'

  end;
  if WithLigneFac then
  begin
    result := result + 'LEFT JOIN LIGNEFAC '
                     + 'ON (GL_NATUREPIECEG = BLF_NATUREPIECEG and GL_SOUCHE = BLF_SOUCHE AND GL_NUMERO = BLF_NUMERO '
                     + 'AND GL_INDICEG = BLF_INDICEG and GL_NUMORDRE = BLF_NUMORDRE AND BLF_UNIQUEBLO=0)';
  end;
  if (fromExcel) and (Document  <> nil) then
  begin
    result := result + 'LEFT JOIN BLIENEXCEL ON BLX_PIECEASSOCIEE="'+stDocument+'" AND BLX_NUMLIGNE=GL_NUMORDRE ';
  end;
  if WithEntetePiece then
  begin
  	result := Result + ' LEFT JOIN PIECE ON (GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_NUMERO=GL_NUMERO'
    								 + ' AND GP_SOUCHE=GL_SOUCHE AND GP_INDICEG=GL_INDICEG)';
  end;
  result := result + ' LEFT OUTER JOIN CONSOMMATIONS '
                   + 'ON (BCO_NUMMOUV = BLP_NUMMOUV AND BCO_INDICE=0 AND BCO_NATUREPIECEG = BLP_NATUREPIECEG and BCO_SOUCHE = BLP_SOUCHE AND BCO_NUMERO = BLP_NUMERO AND BCO_INDICEG = BLP_INDICEG) ';
end;

function CleDocToString(CleDoc: R_CleDoc): string;
begin
  result := FormatDateTime('ddmmyyyy', CleDoc.DatePiece) + ';';
  result := result + Cledoc.NaturePiece + ';';
  result := result + Cledoc.Souche + ';';
  result := result + inttostr(Cledoc.NumeroPiece) + ';';
  result := result + inttostr(Cledoc.Indice) + ';';
  if cledoc.NumLigne > 0 then result := result + inttostr(Cledoc.NumLigne) + ';'
  else Result := result + ';';
end;
Function Evaluedate ( St : String ) : TDateTime ;
Var dd,mm,yy : Word ;
BEGIN
  Result:=0 ; if St='' then Exit ;
  dd:=StrToInt(Copy(St,1,2)) ; mm:=StrToInt(Copy(St,3,2)) ; yy:=StrToInt(Copy(St,5,4)) ;
  Result:=Encodedate(yy,mm,dd) ;
END ;

procedure DecodeRefPiece(St: string; var CleDoc: R_CleDoc); { NEWPIECE }
var
  StC, StL: string;
begin

  FillChar(CleDoc, Sizeof(CleDoc), #0);

  //FV1 : 07/02/2017 - FS#2383 - CLOSSUR - à l'édition d'une facture message bloquant " n'est pas une valeur entière correcte
  if st = '' then exit;

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

procedure StringToCleDoc(StA: string; var CleDoc: R_CleDoc);
var StC: string;
    Dattempo : String;
    IposBlanc : integer;
begin
  StC := uppercase(Trim(ReadTokenSt(StA)));
  if StC = '' then Exit;
  CleDoc.NaturePiece := StC;
  StC := uppercase(Trim(ReadTokenSt(StA)));
  if StC = '' then CleDoc.DatePiece := 0 else
     Begin
     IpoSBlanc := pos(' ', stc);
     if IPosBlanc = 0 then IposBlanc := Length (stc);
     DatTempo := copy(stc, 1,IposBlanc);
     CleDoc.DatePiece := StrToDate(DatTempo);
     end;
  StC := uppercase(Trim(ReadTokenSt(StA)));
  if StC = '' then Exit else CleDoc.Souche := StC;
  StC := uppercase(Trim(ReadTokenSt(StA)));
  if StC = '' then Exit else CleDoc.NumeroPiece := StrtoInt(StC);
  StC := uppercase(Trim(ReadTokenSt(StA)));
  if Stc = '' then Exit else CleDoc.Indice := StrToInt(StC);
end;


function WherePiece(CleDoc: R_CleDoc; ttd: T_TableDoc; Totale: boolean; WithNumOrdre: Boolean = False): string; { NEWPIECE }
var St: string;
  Numpiece: string;
begin
  St := '';
  case ttd of
    ttdTRFEntPOC :
    begin
      St := ' BT2_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BT2_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BT2_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BT2_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    end;
    ttdTRFDetPOC :
    begin
      St := ' BT3_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BT3_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BT3_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BT3_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    end;
    ttdEntRAnal :
    begin
      St := ' BER_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BER_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BER_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BER_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    end;
    ttdReajAnal :
    begin
      St := ' BLR_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLR_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLR_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLR_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    end;
    ttdVteColl :
    begin
      St := ' BPB_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BPB_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BPB_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BPB_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    end;
	  ttdLigneMetre :
    begin
      St := ' BLM_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLM_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLM_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLM_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    end;
    ttdPiece:
    Begin
      St := ' GP_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GP_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND GP_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    End;
    ttdPieceTrait:
    Begin
      St := ' BPE_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BPE_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BPE_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BPE_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    End;
    ttdPiecedemPrix :
    Begin
      St := ' BPP_NATUREPIECEG="'+CleDoc.NaturePiece + '" AND BPP_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BPP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BPP_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    End;
    ttdArticledemPrix :
    Begin
      St := ' BDP_NATUREPIECEG="'+CleDoc.NaturePiece + '" AND BDP_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BDP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BDP_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    End;
    ttdDetaildemPrix :
    Begin
      St := ' BD0_NATUREPIECEG="'+CleDoc.NaturePiece + '" AND BD0_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BD0_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BD0_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    End;
    ttdFourndemPrix :
    Begin
      St := ' BD1_NATUREPIECEG="'+CleDoc.NaturePiece + '" AND BD1_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BD1_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BD1_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
    End;
    ttdLigne:
      begin
        St := ' GL_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GL_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GL_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND GL_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then
        begin
          if not WithNumOrdre then
            St := St + ' AND GL_NUMLIGNE=' + IntToStr(CleDoc.NumLigne) + ' '
          else
            St := St + ' AND GL_NUMORDRE=' + IntToStr(CleDoc.NumOrdre) + ' ';
        end;
      end;
      
    ttdTimbres:
      begin
        St := ' BT0_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BT0_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BT0_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BT0_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;

    ttdLigneTarif:
      begin
        St := ' GLT_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GLT_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GLT_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND GLT_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then St := St + ' AND GLT_NUMLIGNE=' + IntToStr(CleDoc.NumLigne) + ' ';
      end;

    ttdPieceAdr:
      begin
        St := ' GPA_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GPA_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GPA_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND GPA_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then St := St + ' AND GPA_NUMLIGNE=' + IntToStr(CleDoc.NumLigne) + ' ';
      end;
    ttdPieceInterv :
      begin
        St := ' BPI_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BPI_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BPI_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND BPI_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdPiedBase:
      begin
        St := ' GPB_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GPB_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GPB_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND GPB_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdLigneBase:
      begin
        St := ' BLB_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLB_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLB_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLB_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdLignefac:
      begin
        St := ' BLF_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLF_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLF_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLF_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then
        begin
          St := St + ' AND BLF_NUMORDRE=' + IntToStr(CleDoc.Numordre) + ' AND BLF_UNIQUEBLO=0 ';
        end;
      end;
    ttdEche:
      begin
        St := ' GPE_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GPE_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GPE_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND GPE_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdNomen:
      begin
        St := ' GLN_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GLN_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GLN_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND GLN_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdLot:
      begin
        St := ' GLL_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GLL_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GLL_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND GLL_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdAcompte:
      begin
        St := ' GAC_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GAC_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GAC_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND GAC_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdPorc:
      begin
        St := ' GPT_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GPT_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GPT_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND GPT_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdSerie:
      begin
        St := ' GLS_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GLS_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GLS_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND GLS_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdOuvrage:
      begin
        St := ' BLO_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLO_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLO_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND BLO_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdOuvrageP:
      begin
        St := ' BOP_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BOP_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BOP_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND BOP_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdLienOle:
      begin
        Numpiece := cledoc.NaturePiece + ':' + cledoc.Souche + ':' + IntToStr(cledoc.NumeroPiece) + ':' + IntToStr(cledoc.indice);
        St := 'LO_TABLEBLOB="GP" AND LO_IDENTIFIANT="' + NumPiece + '"';
      end;
    ttdretenuG:
      begin
        St := ' PRG_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND PRG_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND PRG_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' AND PRG_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdBaseRG:
      begin
        St := ' PBR_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND PBR_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND PBR_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND PBR_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdSit:
      begin
        St := ' BST_NATUREPIECE="' + CleDoc.NaturePiece + '" AND BST_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BST_NUMEROFAC=' + IntToStr(CleDoc.NumeroPiece) + ' ';
      end;
    ttdParDoc:
      begin
        St := ' BPD_NATUREPIECE="' + CleDoc.NaturePiece + '" AND BPD_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BPD_NUMPIECE=' + IntToStr(CleDoc.NumeroPiece) + ' ';
      end;
    ttdAdresse:
      begin
        St := ' ADR_REFCODE="' + CleDoc.NaturePiece + ';' + CleDoc.Souche + ';' + IntToStr(CleDoc.NumeroPiece) + ';"';
      end;
    ttdRevision:
      begin
        (*
        St := ' BRV_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BRV_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BRV_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BRV_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        *)
        St := ' AFR_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND AFR_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND AFR_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND AFR_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
    ttdVariable:
      begin //Affaire-ONYX
        St := ' AVV_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND AVV_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND AVV_NUMERO=' + IntToStr(CleDoc.NumeroPiece) + ' ';
      end;
    ttdLigneCompl: 
      begin
        St := ' GLC_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND GLC_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND GLC_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND GLC_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then
          St := St + ' AND GLC_NUMORDRE=' + IntToStr(CleDoc.NumOrdre) + ' ';
      end;
    ttdLignePhase:
      begin
        St := ' BLP_NATUREPIECEG="' + CleDoc.NaturePiece + '" '
          + ' AND BLP_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLP_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then
          St := St + ' AND BLP_NUMORDRE=' + IntToStr(CleDoc.NumOrdre) + ' ';
      end;
		TTdRepartmill :
      begin
        St := ' BPM_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BPM_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BPM_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BPM_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
		TTdVarDoc :
      begin
        St := ' BVD_NATUREPIECE="' + CleDoc.NaturePiece + '" AND BVD_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BVD_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BVD_INDICE=' + IntToStr(CleDoc.Indice) + ' ';
      end;
	  ttdLigneOUVP :
      begin
        St := ' BOP_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BOP_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BOP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BOP_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
        if Totale then
          St := St + ' AND BOP_NUMORDRE=' + IntToStr(CleDoc.NumOrdre) + ' ';
      end;
	  ttdTSPOC :
      begin
        St := ' BLE_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLE_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLE_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLE_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
	  ttdTSDetPOC :
      begin
        St := ' BLT_NATUREPIECEG="' + CleDoc.NaturePiece + '" AND BLT_SOUCHE="' + CleDoc.Souche + '" '
          + ' AND BLT_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
          + ' AND BLT_INDICEG=' + IntToStr(CleDoc.Indice) + ' ';
      end;
  end;
  Result := St;
end;



function IsOuvrageOkInPiece (naturepiece : string) : Boolean;
begin
	result := false;
  if (Naturepiece = 'DBT') or (Naturepiece = 'FBT') or
     (Naturepiece = 'AFF') or (Naturepiece = 'FAC') or
     (Naturepiece = 'ETU') or (Naturepiece = 'DAC') or
     (Naturepiece = 'BCE') or (Naturepiece = 'AVC') or
     (Naturepiece = GetParamSoc('SO_BTNATBORDEREAUX')) or
     (Naturepiece = 'ABT') or (Naturepiece = 'FRC') or
     (Naturepiece = 'ABP') or
     (Naturepiece = 'FBP') or (Naturepiece = 'BAC') or
     (Naturepiece = 'DAP') or (Naturepiece = 'FPR') then Result := true;
end;

function IsTransfert(NaturePiece : String): Boolean;
begin
  if (NaturePiece = 'TEM') or (NaturePiece = 'TRV') or (NaturePiece = 'TRE') then
    Result := True
  else Result := False;
end;

end.
