unit FactNomen ;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}  DB, Fe_Main,
{$IFDEF V530} EdtDOC,{$ELSE} EdtRdoc, {$ENDIF}
{$ENDIF}
     SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit,
     FactUtil, FactArticle, FactTOB,  
     Math, StockUtil, EntGC, Classes, HMsgBox, Choix,uEntCommun,UtilTOBPiece;

Function  ChoixNomenclature ( RefUnique,Depot : String ; TOBArticles : TOB ; LePremier : boolean ) : TOB ;
Procedure NomenligVersLignomen ( TOBL,TOBNL,TOBLN : TOB ; Niv,OrdreCompo : integer ) ;
Procedure LoadLesNomen ( TOBPiece,TOBNomenclature,TOBArticles : TOB ; CleDoc : R_CleDoc ) ;
Procedure ReAffecteLigNomen (IndiceNomen : integer ; TOBLig,TOBNomenclature : TOB ) ;
Procedure NomenAPlat ( TOBN,TOBP : TOB ; Qte : double ) ;
Procedure RenseigneValoNomen ( TOBL,TOBN : TOB ) ;
Procedure CreerLesTOBNomen (TOBGroupeNomen,TOBNomen,TOBArticles : TOB ; LaLig,MaxNiv,idep : integer; stDepot : string = '' ) ; // DBR : Depot unique chargé
Procedure ValideLesNomen ( TOBNomenclature : TOB ) ;
Procedure TraiteLesNomenclatures ( TOBPiece,TOBArticles,TOBNomenclature : TOB ; ARow : integer ; LePremier : boolean ) ;
procedure LoadLesLibDetailNomen (TOBPiece,TOBNomenclature,TOBTiers,TOBAffaire:TOB;DEV:RDevise);
procedure LoadLesLibDetailNomLig (TOBPIECE,TOBnomenc,TOBTiers,TOBAffaire,TOBL : TOB;var Indice: integer; DEV:RDevise);
function IsNomenClature (TOBL : TOB) : boolean;

implementation

Uses FactComm,LigNomen ; 


function IsNomenclature (TOBL : TOB) : boolean;
begin
	result := ((TOBL.GetValue('GL_TYPEARTICLE')='NOM') and (TOBL.GetValue('GL_TYPENOMENC')='ASS'));
end;

Procedure ValideLesNomen ( TOBNomenclature : TOB ) ;
Var i : integer ;
BEGIN
for i:=TOBNomenclature.Detail.Count-1 downto 0 do
    if TOBNomenclature.Detail[i].GetValue('UTILISE')<>'X' then TOBNomenclature.Detail[i].Free ;
if Not TOBNomenclature.InsertDB(Nil) then V_PGI.IoError:=oeUnknown ;
END ;

Procedure RenseigneValoNomen ( TOBL,TOBN : TOB ) ;
Var TOBPlat,TOBPD : TOB ;
    i       : integer ;
    Qte,DPA,DPR,PMAP,PMRP : Double ;
BEGIN
TOBPlat:=TOB.Create('',Nil,-1) ;
NomenAPlat(TOBN,TOBPlat,1) ;
DPA:=0 ; DPR:=0 ; PMAP:=0 ; PMRP:=0 ;
for i:=0 to TOBPlat.Detail.Count-1 do
    BEGIN
    TOBPD:=TOBPlat.Detail[i] ;
    Qte:=TOBPD.GetValue('GLN_QTE') ;
    DPA:=DPA+Qte*TOBPD.GetValue('GLN_DPA')    ; DPR:=DPR+Qte*TOBPD.GetValue('GLN_DPR') ;
    PMAP:=PMAP+Qte*TOBPD.GetValue('GLN_PMAP') ; PMRP:=PMRP+Qte*TOBPD.GetValue('GLN_PMRP') ;
    END ;
TOBL.PutValue('GL_DPA',DPA)   ; TOBL.PutValue('GL_DPR',DPR) ;
TOBL.PutValue('GL_PMAP',PMAP) ; TOBL.PutValue('GL_PMRP',PMRP) ;
if TOBPlat<>Nil then TOBPlat.Free ;
END ;

Procedure NomenAPlat ( TOBN,TOBP : TOB ; Qte : double ) ;
Var i : integer ;
    TOBL,TOBPD : TOB ;
    RefArt    : String ;
    QteN,X : Double ;
BEGIN
for i:=0 to TOBN.Detail.Count-1 do
    BEGIN
    TOBL:=TOBN.Detail[i] ;
    QteN:=TOBL.GetValue('GLN_QTE') ;
    if TOBL.Detail.Count>0 then
       BEGIN
       NomenAPlat(TOBL,TOBP,Qte*QteN) ;
       END else
       BEGIN
       RefArt:=TOBL.GetValue('GLN_ARTICLE') ;
       TOBPD:=TOBP.FindFirst(['GLN_ARTICLE'],[RefArt],False) ;
       if TOBPD<>Nil then
          BEGIN
          X:=TOBPD.GetValue('GLN_QTE')  ; X:=X+Qte*QteN ; TOBPD.PutValue('GLN_QTE',X) ;
          END else
          BEGIN
          TOBPD:=TOB.Create('LIGNENOMEN',TOBP,-1) ;
          TOBPD.Dupliquer(TOBL,False,True) ;
          TOBPD.PutValue('GLN_QTE',Qte*QteN) ;
          END ;
       END ;
    END ;
END ;

Procedure MajLigNomen (TOBN,TOBLig : TOB ) ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBN.Detail.Count-1 do
    BEGIN
    TOBL:=TOBN.Detail[i] ;
    TOBL.PutValue('GLN_NATUREPIECEG',TOBLig.GetValue('GL_NATUREPIECEG')) ;
    TOBL.PutValue('GLN_SOUCHE',TOBLig.GetValue('GL_SOUCHE')) ;
    TOBL.PutValue('GLN_NUMERO',TOBLig.GetValue('GL_NUMERO')) ;
    TOBL.PutValue('GLN_INDICEG',TOBLig.GetValue('GL_INDICEG')) ;
    TOBL.PutValue('GLN_NUMLIGNE',TOBLig.GetValue('GL_NUMLIGNE')) ;
    MajLigNomen(TOBL,TOBLig) ;
    END ;
END ;

Procedure ReAffecteLigNomen (IndiceNomen : integer ; TOBLig,TOBNomenclature : TOB ) ;
Var TOBN : TOB ;
BEGIN
if IndiceNomen<=0 then Exit ;
if TOBNomenclature=Nil then Exit ;
if TOBNomenclature.Detail.Count-1<IndiceNomen-1 then Exit ;
TOBN:=TOBNomenclature.Detail[IndiceNomen-1] ; TOBN.PutValue('UTILISE','X') ;
MajLigNomen(TOBN,TOBLig) ;
END ;

Procedure CreerLesTOBNomen (TOBGroupeNomen,TOBNomen,TOBArticles : TOB ; LaLig,MaxNiv,idep : integer; stDepot : string = '' ) ; // DBR : Depot unique chargé
Var LeNiv,Niv,i,Lig : integer ;
    TOBLN,TOBPere,TOBLoc,TOBArt : TOB ;
    RefArticle           : String ;
BEGIN
for LeNiv:=1 to MaxNiv do
    BEGIN
    for i:=idep to TOBNomen.Detail.Count-1 do
        BEGIN
        TOBLN:=TOBNomen.Detail[i] ;
        Niv:=TOBLN.GetValue('GLN_NIVEAU') ; Lig:=TOBLN.GetValue('GLN_NUMLIGNE') ;
        if Lig<>LaLig then Continue ;
        if Niv=LeNiv then
           BEGIN
           if Niv=1 then TOBPere:=TOBGroupeNomen else
              BEGIN
              TOBPere:=TOBGroupeNomen.FindFirst(['GLN_NUMLIGNE','GLN_NIVEAU','GLN_ARTICLE','GLN_NUMORDRE'],[Lig,Niv-1,TOBLN.GetValue('GLN_COMPOSE'),TOBLN.GetValue('GLN_ORDRECOMPO')],True) ;
              END ;
           if TOBPere<>Nil then
              BEGIN
              TOBLoc:=TOB.Create('LIGNENOMEN',TOBPere,-1) ;
              TOBLoc.Dupliquer(TOBLN,False,True) ;
              RefArticle:=TOBLoc.GetValue('GLN_ARTICLE') ; TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefArticle],False) ;
              if TOBArt=Nil then
                 BEGIN
                 TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.SelectDB('"'+RefArticle+'"',Nil) ;
                 LoadTOBDispo(TOBArt,False, '"' + stDepot + '"'); // DBR : Depot unique à charger
                 END ;
              END ;
           END ;
        END ;
    END ;
END ;

Procedure LoadLesNomen ( TOBPiece,TOBNomenclature,TOBArticles : TOB ; CleDoc : R_CleDoc ) ;
Var i,OldL,Lig,MaxNiv,k,Niv,IndiceNomen : integer ;
    TOBL,TOBLN,TOBNomen,TOBGroupeNomen,TOBLoc : TOB ;
    OkN  : boolean ;
    stDepot : String ; // DBR : Depot unique chargé
    Q    : TQuery ;
BEGIN
OkN:=False ; OldL:=-1 ; IndiceNomen:=1 ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ; TOBL.PutValue('GL_INDICENOMEN',0) ;
    if ((TOBL.GetValue('GL_TYPEARTICLE')='NOM') and (TOBL.GetValue('GL_TYPENOMENC')='ASS')) then OkN:=True ;
    END ;
if Not OkN then Exit ;
TOBNomen:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM LIGNENOMEN WHERE '+WherePiece(CleDoc,ttdNomen,False)+' ORDER BY GLN_NUMLIGNE, GLN_NIVEAU, GLN_NUMORDRE',True,-1, '', True) ;
TOBNomen.LoadDetailDB('LIGNENOMEN','','',Q,True,False) ;
Ferme(Q) ;
for i:=0 to TOBNomen.Detail.Count-1 do
    BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('GLN_NUMLIGNE') ;
    if OldL<>Lig then
       BEGIN
       TOBGroupeNomen:=TOB.Create('',TOBNomenclature,-1) ; MaxNiv:=-1 ;
       TOBGroupeNomen.AddChampSup('UTILISE',False) ; TOBGroupeNomen.PutValue('UTILISE','-') ;
       for k:=i to TOBNomen.Detail.Count-1 do
           BEGIN
           TOBLoc:=TOBNomen.Detail[k] ;
           if TOBLoc.GetValue('GLN_NUMLIGNE')<>Lig then Break ;
           Niv:=TOBLoc.GetValue('GLN_NIVEAU') ; if Niv>MaxNiv then MaxNiv:=Niv ;
           END ;
       TobL := TobPiece.FindFirst(['GL_NUMLIGNE'], [Lig], False);  // DBR : Depot unique chargé
       if TobL = nil then stDepot := '' else stDepot := TobL.GetValue ('GL_DEPOT');
       CreerLesTOBNomen(TOBGroupeNomen,TOBNomen,TOBArticles,Lig,MaxNiv,i, stDepot);
       END ;
    OldL:=Lig ;
    END ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    if ((TOBL.GetValue('GL_TYPEARTICLE')='NOM') and (TOBL.GetValue('GL_TYPENOMENC')='ASS')) then
       BEGIN
       TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ; Inc(IndiceNomen) ;
       END ;
    END ;
TOBNomen.Free ;
END ;

Procedure NomenligVersLignomen ( TOBL,TOBNL,TOBLN : TOB ; Niv,OrdreCompo : integer ) ;
Var i : integer ;
    TOBNLD,TOBLND : TOB ;
BEGIN
for i:=0 to TOBNL.Detail.Count-1 do
    BEGIN
    TOBNLD:=TOBNL.Detail[i] ;
    TOBLND:=TOB.Create('LIGNENOMEN',TOBLN,-1) ;
    TOBLND.PutValue('GLN_ARTICLE',TOBNLD.GetValue('GNL_ARTICLE')) ;
    TOBLND.PutValue('GLN_CODEARTICLE',TOBNLD.GetValue('GNL_CODEARTICLE')) ;
    TOBLND.PutValue('GLN_NUMORDRE',TOBNLD.GetValue('GNL_NUMLIGNE')) ;
    TOBLND.PutValue('GLN_LIBELLE',TOBNLD.GetValue('GNL_LIBELLE')) ;
    TOBLND.PutValue('GLN_QTE',TOBNLD.GetValue('GNL_QTE')) ;
    TOBLND.PutValue('GLN_COMPOSE',TOBNLD.GetValue('COMPOSE')) ;
    TOBLND.PutValue('GLN_TENUESTOCK',TOBNLD.GetValue('TENUESTOCK')) ;
    TOBLND.PutValue('GLN_QUALIFQTESTO',TOBNLD.GetValue('QUALIFQTESTO')) ;
    TOBLND.PutValue('GLN_QUALIFQTEACH',TOBNLD.GetValue('QUALIFQTEACH')) ;
    TOBLND.PutValue('GLN_QUALIFQTEVTE',TOBNLD.GetValue('QUALIFQTEVTE')) ;
    TOBLND.PutValue('GLN_DPA',TOBNLD.GetValue('DPA'))   ;
    TOBLND.PutValue('GLN_DPR',TOBNLD.GetValue('DPR'))   ;
    TOBLND.PutValue('GLN_PMAP',TOBNLD.GetValue('PMAP')) ;
    TOBLND.PutValue('GLN_PMRP',TOBNLD.GetValue('PMRP')) ;
    TOBLND.PutValue('GLN_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
    TOBLND.PutValue('GLN_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
    TOBLND.PutValue('GLN_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
    TOBLND.PutValue('GLN_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
    TOBLND.PutValue('GLN_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
    TOBLND.PutValue('GLN_NIVEAU',Niv) ;
    TOBLND.PutValue('GLN_ORDRECOMPO',OrdreCompo) ;
    NomenligVersLigNomen(TOBL,TOBNLD,TOBLND,Niv+1,TOBLND.GetValue('GLN_NUMORDRE')) ;
    END ;
END ;

Procedure DerouleNomenclature ( TOBN : TOB ; CodeNomen,RefUnique,Depot : String ; TOBArticles : TOB ) ;
Var Q : TQuery ;
    TOBL,TOBArt : TOB ;
    i    : integer ;
    CodeN,RefArticle  : String ;
BEGIN
// Nomenclatures
Q:=OpenSQL('SELECT * FROM NOMENENT WHERE GNE_NOMENCLATURE="'+CodeNomen+'"',True,-1, '', True);
TOBN.SelectDB('', Q);
Ferme(Q) ;
Q:=OpenSQL('SELECT * FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+CodeNomen+'" ORDER BY GNL_NUMLIGNE',True,-1, '', True) ;
TOBN.LoadDetailDB('NOMENLIG','','',Q,False,False) ;
Ferme(Q) ;
for i:=0 to TOBN.Detail.Count-1 do
    BEGIN
    TOBL:=TOBN.Detail[i] ;
    TOBL.AddChampSup('COMPOSE',False) ; TOBL.AddChampSup('TENUESTOCK',False) ;
    TOBL.AddChampSup('QUALIFQTEACH',False) ; TOBL.AddChampSup('QUALIFQTEVTE',False) ; TOBL.AddChampSup('QUALIFQTESTO',False) ;
    TOBL.AddChampSup('DPA',False) ; TOBL.AddChampSup('DPR',False) ; TOBL.AddChampSup('PMAP',False) ; TOBL.AddChampSup('PMRP',False) ;
    TOBL.PutValue('COMPOSE',RefUnique) ;
    // Articles
    RefArticle:=TOBL.GetValue('GNL_ARTICLE') ; TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefArticle],False) ;
    if TOBArt=Nil then
       BEGIN
       TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.SelectDB('"'+RefArticle+'"',Nil) ;
       LoadTOBDispo(TOBArt,False, '"' + Depot + '"') ; // DBR : Depot unique chargé
       END ;
    TOBL.PutValue('TENUESTOCK',TOBArt.GetValue('GA_TENUESTOCK')) ;
    TOBL.PutValue('QUALIFQTESTO',TOBArt.GetValue('GA_QUALIFUNITESTO')) ;
    TOBL.PutValue('QUALIFQTEVTE',TOBArt.GetValue('GA_QUALIFUNITEVTE')) ;
    AffecteValoNomen(TOBL,TOBArt,Depot) ;
    // Nomenclatures
    CodeN:=TOBL.GetValue('GNL_SOUSNOMEN') ;
    if CodeN<>'' then DerouleNomenclature(TOBL,CodeN,TOBL.GetValue('GNL_ARTICLE'),Depot,TOBArticles) ;
    END ;
END ;

Procedure TraiteLesNomenclatures ( TOBPiece,TOBArticles,TOBNomenclature : TOB ; ARow : integer ; LePremier : boolean ) ;
Var RefUnique,TypeArt,TypeNomenc,Depot : String ;
    TOBL,TOBNomen,TOBLN,TOBTMP : TOB ;
    IndiceNomen : integer ;
BEGIN
RefUnique:=GetCodeArtUnique(TOBPiece,ARow) ; if RefUnique='' then Exit ;
TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
TypeArt:=TOBL.GetValue('GL_TYPEARTICLE') ; if TypeArt<>'NOM' then Exit ;
TypeNomenc:=TOBL.GetValue('GL_TYPENOMENC') ; if TypeNomenc<>'ASS' then Exit ;
Depot:=TOBL.GetValue('GL_DEPOT') ; if Depot='' then Exit ;
TOBNomen:=ChoixNomenclature(RefUnique,Depot,TOBArticles,False) ;
{$IFDEF GPAOLIGHT}
if (VH_GC.OASeria) and (TOBL.GetValue('GL_TENUESTOCK') = 'X') then
begin
  TOBL.PutValue('GLC_IDENTIFIANTWNT', TobNomen.GetValue('GNE_IDENTIFIANTWNT'));
  TOBNomen.Free ;
  Exit;
end;
{$ENDIF GPAOLIGHT}
TOBLN:=TOB.Create('',TOBNomenclature,-1) ;
TOBLN.AddChampSup('UTILISE',False) ; TOBLN.PutValue('UTILISE','-') ;
if ((TOBNomen<>Nil) and (TOBNomen.Detail.Count>0)) then
   BEGIN
   NomenligVersLignomen(TOBL,TOBNomen,TOBLN,1,0) ;
   END else
   BEGIN
   TOBTMP:=TOB.Create('LIGNENOMEN',TOBLN,-1) ;
   TOBTMP.PutValue('GLN_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
   TOBTMP.PutValue('GLN_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
   TOBTMP.PutValue('GLN_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
   TOBTMP.PutValue('GLN_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
   TOBTMP.PutValue('GLN_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
   TOBTMP.PutValue('GLN_COMPOSE',RefUnique) ;
   Entree_LigneNomen(TOBLN,TobArticles,False,1,0, taModif) ;
   END ;
RenseigneValoNomen(TOBL,TOBLN) ;
TOBNomen.Free ;
IndiceNomen:=TOBNomenclature.Detail.Count ; TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ;
END ;

Function  ChoixNomenclature ( RefUnique,Depot : String ; TOBArticles : TOB ; LePremier : boolean ) : TOB ;
Var TOBN : TOB ;
    Q    : TQuery ;
    TT   : TStrings ;
    CodeNomen,St : String ;
BEGIN
Result:=Nil ; CodeNomen:='' ;
Q:=OpenSQL('SELECT GNE_NOMENCLATURE, GNE_LIBELLE FROM NOMENENT WHERE GNE_ARTICLE="'+RefUnique+'"',True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
   TT:=TStringList.Create ;
   While Not Q.EOF do
      BEGIN
      TT.Add(Q.Fields[0].AsString+';'+Q.Fields[1].AsString) ;
      if LePremier then Break else Q.Next ;
      END ;
   if TT.Count>1 then
      BEGIN
      CodeNomen:=Choisir('Choix d''une nomenclature','NOMENENT','GNE_LIBELLE','GNE_NOMENCLATURE','GNE_ARTICLE="'+RefUnique+'"','') ;
      END else
      BEGIN
      St:=TT[0] ; CodeNomen:=ReadTokenSt(St) ;
      END ;
   TT.Clear ; TT.Free ;
   END ;
Ferme(Q) ;
if CodeNomen='' then Exit ;
TOBN:=TOB.Create('',Nil,-1) ;
DerouleNomenclature(TOBN,CodeNomen,RefUnique,Depot,TOBArticles) ;
Result:=TOBN ;
END ;


procedure LoadLesLibDetailNomLig (TOBPIECE,TOBnomenc,TOBTiers,TOBAffaire,TOBL : TOB;var Indice: integer; DEV:RDevise);
var IndOuv,IndiceNomen,IndiceLig,TypePresent,NiveauImbric : integer;
    TOBL1,TOBNom,TOBLoc : TOB;
//    QteDuDetail : double;
BEGIN
if (TOBL.GetValue('GL_TYPENOMENC') = 'ASS') and (TOBL.GetValue('GL_TYPEPRESENT')>0) then
    begin
    TypePresent := TOBL.GetValue('GL_TYPEPRESENT');
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    NiveauImbric := TOBL.Getvalue('GL_NIVEAUIMBRIC');
    if IndiceNomen = 0 then
       begin
       inc (indice);
       exit;
       end;
    TOBLoc := TOBnomenc.detail[IndiceNomen -1];
    if TOBLoc = nil then
       begin
       // Surement une ligne nomenclature
       inc (indice);
       exit;
       end;
    IndiceLig := Indice+1;
    for IndOuv := 0 to TOBLoc.detail.count -1 do
       begin
       TOBNom := TOBLOC.detail[Indouv];
       TOBL1:=NewTOBLigne(TOBPiece,IndiceLig+indOuv+1);
{       TOBL1:=TOB.Create('LIGNE',TOBPiece,IndiceLig+indOuv);
       TOB.Create('',TOBL1,-1) ;
       AddLesSupLigne(TOBL1,False) ;
       InitLesSupLigne(TOBL1) ;  }
       InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,IndiceLig+indOuv+1) ;
       TOBL1.PutValue('GL_TYPELIGNE','COM') ;
       TOBL1.PutValue('GL_INDICENOMEN',IndiceNomen);
       TOBL1.PutValue('GL_NIVEAUIMBRIC',NiveauImbric);
{
       if EnHt then Montant := arrondi(TOBOUV.GetValue('BLO_PUHTDEV')* TOBL.getvalue('GL_QTEFACT') * (TOBOUV.GetValue('BLO_QTEFACT')/ (QteDuPv * QteDuDetail)),DEV.Decimale )
       else Montant := arrondi(TOBOUV.GetValue('BLO_PUTTCDEV')* TOBL.getvalue('GL_QTEFACT') * (TOBOUV.GetValue('BLO_QTEFACT')/ (QteDuPv * QteDuDetail)),DEV.Decimale );
}
       if (TypePresent and 1) = 1  then
          BEGIN
          TOBL1.PutValue('GL_REFARTSAISIE',TobNom.GetValue('GLN_CODEARTICLE'));
          END;
       TOBL1.PutValue('GL_REFARTTIERS',TobNom.GetValue('GLN_ARTICLE'));
       if (TypePresent and 2) = 2  then TOBL1.PutValue('GL_LIBELLE',TobNom.GetValue('GLN_LIBELLE'));
       if (typepresent and 4) = 4  then TOBL1.PutValue('GL_QTEFACT',TobNom.GetValue('GLN_QTEFACT'));
       if (typepresent and 8) = 8  then TOBL1.PutValue('GL_QUALIFQTEVTE',TobNom.GetValue('GLN_QUALIFQTEVTE'));
       end;
    indice := Indice + TOBLoc.detail.count+1;
    end else
    begin
    inc (indice);
    end;
END;

procedure LoadLesLibDetailnomen (TOBPIECE,TOBNomenclature,TOBTiers,TOBAffaire : TOB;DEV:RDevise );
var indice : integer;
    TOBL : TOB;
begin
Indice := 0;
if TOBPiece.detail.count = 0 then exit; // mcd 27/02/02 cas saisie ligne sur affaire dans GI/GA
repeat
    TOBL := TOBPiece.detail[indice];
    if TOBL <> nil then LoadLesLibDetailNomLig ( TOBPIECE,TOBNomenclature,TOBTiers,TOBAffaire,TOBL ,Indice,DEV)
                   else inc(Indice);
until indice > TOBPiece.detail.count -1;
end;

end.
