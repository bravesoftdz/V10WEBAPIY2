unit UtilGrp ;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, SysUtils, SaisUtil, UtilPGI, AGLInit,
     FactUtil, UtilSais, EntGC, Classes, SaisComm, FactComm, FactNomen,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     FactAffaire, FactPiece, FactCpta, FactCalc, ParamSoc, UtilArticle,
{$IFDEF BTP}
     BTPUtil,UplannifChUtil,
{$ENDIF}
     FactArticle, DepotUtil, FactTOB,
     FactContreM , StockUtil,uEntCommun,UtilTOBPiece ;

Procedure UG_ValideAnals(TOBGenere,TOBAnaP,TOBAnaS : TOB ) ;
Procedure UG_ChargeLesPorcs ( TOBPiece,TOBPorcs : TOB ) ;
Procedure UG_SommeLesPorcs ( TOBPL,TOBP : TOB ; Manuelle : boolean ) ;
Procedure UG_PieceVersLigne ( TOBPiece,TOBL : TOB; AvecRemEsc : boolean=true ) ;
Procedure UG_PieceVersLigneTarif(TobGenere, TobLigneTarif, TobL: TOB; NewNumLigne: Integer);
Procedure UG_ChargeLesAcomptes ( TOBPiece,TOBAcomptes : TOB ) ;
Procedure UG_LoadAnaPiece ( TOBPiece,TOBAnaP,TOBAnaS : TOB ) ;
Procedure UG_LoadAnaLigne ( TOBPiece,TOBL : TOB ) ;
Procedure UG_AjouteLesArticles ( TOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu : TOB ; Manuelle : boolean ) ;
Procedure UG_MajAnalPiece ( TOBAnaP,TOBAnaS : TOB ; RefA : String ) ;
Procedure UG_MajAnalLigne ( TOBL : TOB ; RefA : String ) ;
Procedure UG_MajLesComms ( TOBPiece,TOBArticles,TOBComms : TOB ) ;
Procedure UG_AjouteLesRepres ( TOBPiece,TOBComms : TOB ) ;
Procedure UG_ConstruireTobArt(TOBPiece,TOBArticles:Tob) ;
{$IFDEF BTP}
procedure UG_AjouteLesChampsBtp (TOBPiece : TOB);
{$ENDIF}
Procedure UG_RecupLesAcomptes (TOBPiece,TOBAcomptes,TOBAcomptesGen,TOBAcomptesGen_O : TOB);
Procedure UG_RegroupeLesAcomptes(TOBAcomptes : TOB); // JT eQualité 11013
//
procedure UG_ChargePiecesPrecedente (TOBPiece,TOBGenere : TOB;GestionReliquat : boolean);
procedure UG_ReajustePiecesPrecedente (TOBGenere,TOBPiece,TOBArticles: TOB) ;
//
procedure UG_CreeLesTob (var TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                      TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TOBLiaison_o,TobLigneTarif_O: TOB);
procedure UG_DropLesTobs (TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                      TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLigneTarif_O : TOB);
procedure UG_VideLesTobs (TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                       TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLigneTarif_O : TOB);
procedure UG_ChargeLesTobs (TOBPiece,TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                       TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TOBLiaison_O,TOBArticles,TobLigneTarif_O : TOB);
procedure UG_DetruitLaPiece (TOBPiece: TOB) ; { NEWPIECE }
Procedure UG_ChargeLesRG (TOBPiece,TOBPieceRG, TOBBasesRG : TOB);
Procedure UG_LoadOuvrageLigne ( TOBL , TOBOuvrage, TOBArticles : TOB ; var IndiceNomen : integer ; RepriseAffaireRef : boolean) ;
//
function UG_IsAnalMultiple (TOBSource : TOB) : boolean;
Procedure UG_LoadAnaLigneOuPas ( TOBPiece,TOBL : TOB ) ;
//
procedure UG_RegenAnalytique ( TOBPiece,TOBArticles,TOBAFFAIRE,TOBL,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu : TOB ; Manuelle : boolean ) ;

Const NbArtParRequete=200;

implementation

uses
{$IFDEF BTP}
  factOuvrage, FactVariante,
{$ENDIF}
  FactTarifs,UtilBTPgestChantier;

Procedure UG_ValideAnals(TOBGenere,TOBAnaP,TOBAnaS : TOB ) ;
BEGIN
ValideAnalytiques(TOBGenere,TOBAnaP,TOBAnaS) ;
END ;

Procedure UG_ChargeLesPorcs ( TOBPiece,TOBPorcs : TOB ) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
    i,k : integer ;
    PorcsLoc,TOBP,TOBPL : TOB ;
    FTi : Array[1..5] of String ;
    CodePort,TypePort,ModeGroupPort : String ;
BEGIN
PorcsLoc:=TOB.Create('LESPORCS',Nil,-1) ; CD:=TOB2CleDoc(TOBPiece) ;
if TOBPiece.fieldExists('PIECEORIGINE') then
begin
	DecodeRefPiece (TOBPiece.GetValue('PIECEORIGINE'),CD);
end else
begin
	CD:=TOB2CleDoc(TOBPiece) ;
end;
Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CD,ttdPorc,False),True) ;
if Not Q.EOF then PorcsLoc.LoadDetailDB('PIEDPORT','','',Q,True,False) ;
Ferme(Q) ;
for i:=0 to PorcsLoc.Detail.Count-1 do
    BEGIN
    TOBPL:=PorcsLoc.Detail[i] ;
    CodePort:=TOBPL.GetValue('GPT_CODEPORT') ;
    TypePort:=TOBPL.GetValue('GPT_TYPEPORT') ;
    for k:=1 to 5 do FTi[k]:=TOBPL.GetValue('GPT_FAMILLETAXE'+IntToStr(k)) ;
    TOBP:=TOBPorcs.FindFirst(['GPT_CODEPORT','GPT_TYPEPORT','GPT_FAMILLETAXE1','GPT_FAMILLETAXE2','GPT_FAMILLETAXE3','GPT_FAMILLETAXE4','GPT_FAMILLETAXE5'],
                             [CodePort,TypePort,FTi[1],FTi[2],FTi[3],FTi[4],FTi[5]],True) ;
    if TOBP<>Nil then
       BEGIN
       ModeGroupPort:=GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_MODEGROUPEPORT') ;
       UG_SommeLesPorcs(TOBPL,TOBP,(ModeGroupPort='FUS')) ;
       END else
       BEGIN
       TOBP:=TOB.Create('PIEDPORT',TOBPorcs,-1);
       TOBP.Dupliquer(PorcsLoc.Detail[i],True,True);
       END;
    END ;
PorcsLoc.Free ;
END ;

Procedure UG_SommeLesPorcs ( TOBPL,TOBP : TOB ; Manuelle : boolean ) ;
Var X : Double ;
BEGIN
X:=TOBP.GetValue('GPT_BASEHT')      ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASEHT'),9)      ; TOBP.PutValue('GPT_BASEHT',X) ;
X:=TOBP.GetValue('GPT_BASEHTDEV')   ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASEHTDEV'),9)   ; TOBP.PutValue('GPT_BASEHTDEV',X) ;
X:=TOBP.GetValue('GPT_BASETTC')     ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASETTC'),9)     ; TOBP.PutValue('GPT_BASETTC',X) ;
X:=TOBP.GetValue('GPT_BASETTCDEV')  ; X:=Arrondi(X+TOBPL.GetValue('GPT_BASETTCDEV'),9)  ; TOBP.PutValue('GPT_BASETTCDEV',X) ;
X:=TOBP.GetValue('GPT_TOTALHT')     ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALHT'),9)     ; TOBP.PutValue('GPT_TOTALHT',X) ;
X:=TOBP.GetValue('GPT_TOTALHTDEV')  ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALHTDEV'),9)  ; TOBP.PutValue('GPT_TOTALHTDEV',X) ;
X:=TOBP.GetValue('GPT_TOTALTTC')    ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALTTC'),9)    ; TOBP.PutValue('GPT_TOTALTTC',X) ;
X:=TOBP.GetValue('GPT_TOTALTTCDEV') ; X:=Arrondi(X+TOBPL.GetValue('GPT_TOTALTTCDEV'),9) ; TOBP.PutValue('GPT_TOTALTTCDEV',X) ;
if Not Manuelle then
   BEGIN
   if TOBPL.GetValue('GPT_TYPEPORT')='MT' then
      TOBP.PutValue('GPT_MONTANTMINI',TOBP.GetValue('GPT_TOTALHTDEV')) ;
   END ;
END ;

Procedure UG_PieceVersLigne ( TOBPiece,TOBL : TOB ; AvecRemEsc : boolean=true) ;
var TypePresent : integer;
BEGIN
//TypePresent := 0;
TOBL.PutValue('GL_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
TOBL.PutValue('GL_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
TOBL.PutValue('GL_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
TOBL.PutValue('GL_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
TOBL.PutValue('GL_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
TOBL.PutValue('GL_VIVANTE',TOBPiece.GetValue('GP_VIVANTE')) ;
TOBL.PutValue('GL_DATEMODIF',TOBPiece.GetValue('GP_DATEMODIF')) ;
TOBL.PutValue('GL_DATECREATION',TOBPiece.GetValue('GP_DATECREATION')) ;
TOBL.PutValue('GL_ETABLISSEMENT',TOBPiece.GetValue('GP_ETABLISSEMENT')) ;
if AvecRemEsc then
begin
  TOBL.PutValue('GL_REMISEPIED',TOBPiece.GetValue('GP_REMISEPIED')) ;
  TOBL.PutValue('GL_ESCOMPTE',TOBPiece.GetValue('GP_ESCOMPTE')) ;
end;
TOBL.PutValue('GL_CREERPAR',TOBPiece.GetValue('GP_CREEPAR')) ;
// MODIF LS
if TOBPiece.GetValue('GP_TVAENCAISSEMENT')='PE' then TOBL.PutValue('GL_TVAENCAISSEMENT','X')
																								else TOBL.PutValue('GL_TVAENCAISSEMENT','-');
TOBL.PutValue('GL_COTATION',TOBPiece.GetValue('GP_COTATION')) ;
// Modif Btp
if ((TOBL.GetValue('GL_TYPENOMENC') = 'OUV') or (TOBL.GetValue('GL_TYPENOMENC') = 'NOM')) and
   (GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_RECUPPRE') = 'SUI') then
begin
  TypePresent := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_TYPEPRESENT');
  TOBL.PutValue('GL_TYPEPRESENT',TypePresent) ;
end;
// --------------
END ;

Procedure UG_PieceVersLigneTarif(TobGenere, TobLigneTarif, TobL: TOB; NewNumLigne: Integer);
var                                                                                                   
   i: Integer;                                                                                        
   TobLTarif: Tob;                                                                                    
begin                                                                                                 
   TobLTarif := GetLigneTarif(TobLigneTarif, TobL);                                                   
   if TobLTarif <> nil then                                                                           
   begin                                                                                              
      for i := 0 to TobLTarif.Detail.Count - 1 do                                                     
      begin                                                                                           
         ToblTarif.Detail[i].PutValue('GLT_NATUREPIECEG', TobGenere.GetValue('GP_NATUREPIECEG'));     
         ToblTarif.Detail[i].PutValue('GLT_SOUCHE'      , TobGenere.GetValue('GP_SOUCHE'));           
         ToblTarif.Detail[i].PutValue('GLT_NUMERO'      , TobGenere.GetValue('GP_NUMERO'));           
         ToblTarif.Detail[i].PutValue('GLT_INDICEG'     , TobGenere.GetValue('GP_INDICEG'));          
         ToblTarif.Detail[i].PutValue('GLT_NUMLIGNE'    , NewNumLigne);                               
      end;                                                                                            
      ToblTarif.PutValue('GL_NATUREPIECEG', TobGenere.GetValue('GP_NATUREPIECEG'));                   
      ToblTarif.PutValue('GL_SOUCHE'      , TobGenere.GetValue('GP_SOUCHE'));                         
      ToblTarif.PutValue('GL_NUMERO'      , TobGenere.GetValue('GP_NUMERO'));                         
      ToblTarif.PutValue('GL_INDICEG'     , TobGenere.GetValue('GP_INDICEG'));                        
      ToblTarif.PutValue('GL_NUMLIGNE'    , NewNumLigne);                                             
   end;                                                                                               
end;

Procedure UG_RecupLesAcomptes (TOBPiece,TOBAcomptes,TOBAcomptesGen,TOBAcomptesGen_O : TOB);
var Indice : Integer;
    TOBACC,TOBNACC : TOB;
    Q : TQuery ;
    CD : R_CleDoc ;
begin
TOBACC := TOBAcomptesGen.findfirst(['NATURE','NUMERO'],[TOBPiece.GetValue('GP_NATUREPIECEG'),TOBPiece.Getvalue('GP_NUMERO')],false);
if TOBACC <> nil then
   BEGIN
   for Indice := 0 to TOBACC.detail.count -1 do
       begin
       TOBNACC := TOB.Create ('ACOMPTES',TOBAcomptes,-1);
       TOBNACC.dupliquer (TOBACC.detail[Indice],true,true);
       end;
   END;
CD:=TOB2CleDoc(TOBPiece) ;
Q:=OpenSQL('SELECT * FROM ACOMPTES WHERE '+WherePiece(CD,ttdAcompte,False),True) ;
TOBAcomptesGEN_O.LoadDetailDB('ACOMPTES','','',Q,true) ;
Ferme(Q) ;
end;

Procedure UG_ChargeLesAcomptes ( TOBPiece,TOBAcomptes : TOB ) ;
Var Q : TQuery ;
    CD : R_CleDoc ;
BEGIN
CD:=TOB2CleDoc(TOBPiece) ;
Q:=OpenSQL('SELECT * FROM ACOMPTES WHERE '+WherePiece(CD,ttdAcompte,False),True) ;
if Not Q.EOF then TOBAcomptes.LoadDetailDB('ACOMPTES','','',Q,True,False) ;
Ferme(Q) ;
END ;

// JT eQualité 11013
Procedure UG_RegroupeLesAcomptes(TOBAcomptes : TOB);
var TobAcptTmp, TobTmp, TobTmp1 : TOB;
    Cpt : integer;
begin
// Regrpe acompte si journal et n° écriture identique
  if TOBAComptes = nil then exit;
  if TOBAcomptes.Detail.count <= 1 then exit;
  TobAcptTmp := TOB.Create('',nil,-1);
  TobAcptTmp.Dupliquer(TOBAcomptes,true,true,true);
  TobAcomptes.ClearDetail;
  for Cpt := 0 to TobAcptTmp.Detail.count -1 do
  begin
    TobTmp := TobAcptTmp.Detail[Cpt];
    TobTmp1 := TobAcomptes.FindFirst(['GAC_JALECR','GAC_NUMECR'],
                                     [TobTmp.GetValue('GAC_JALECR'),TobTmp.GetValue('GAC_NUMECR')],True);
    if TobTmp1 = nil then
    begin
      TobTmp1 := TOB.Create('ACOMPTES',TOBAcomptes,-1);
      TobTmp1.Dupliquer(TobTmp,true,true,true);
    end else
    begin
      TobTmp1.PutValue('GAC_MONTANT',TobTmp1.GetValue('GAC_MONTANT')+TobTmp.GetValue('GAC_MONTANT'));
      TobTmp1.PutValue('GAC_MONTANTDEV',TobTmp1.GetValue('GAC_MONTANTDEV')+TobTmp.GetValue('GAC_MONTANTDEV'));
    end;
  end;
  if TobAcptTmp <> nil then
    FreeAndNil(TobAcptTmp);
end;

Procedure UG_LoadAnaPiece ( TOBPiece,TOBAnaP,TOBAnaS : TOB ) ;
Var Q : TQuery ;
    RefA : String ;
    TOBA : TOB ;
    i    : integer ;
BEGIN
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE (YVA_TABLEANA="GL" OR YVA_TABLEANA="GS") AND YVA_IDENTIFIANT="'+RefA+'" AND YVA_IDENTLIGNE="000"',True) ;
if Not Q.EOF then
   BEGIN
   TOBAnaP.LoadDetailDB('VENTANA','','',Q,False) ;
   for i:=TOBAnaP.Detail.Count-1 downto 0 do
       BEGIN
       TOBA:=TOBAnaP.Detail[i] ;
       if TOBA.GetValue('YVA_TABLEANA')='GS' then TOBA.ChangeParent(TOBAnaS,0) ;
       END ;
   END ;
Ferme(Q) ;
END ;

Procedure UG_ChargeLesRG (TOBPiece,TOBPieceRG, TOBBasesRG : TOB);
var CD : r_cledoc;
begin
   CD:=TOB2CleDoc(TOBPiece) ;
	 LoadLesRetenues (TOBPiece,TOBPieceRG,TOBBasesRG);
end;

Procedure UG_LoadAnaLigne ( TOBPiece,TOBL : TOB ) ;
Var Q : TQuery ;
    RefA : String ;
    NumL : integer ;
    TOBAL : TOB ;
BEGIN
if TOBL=Nil then Exit ;
if TOBL.Detail.Count<=0 then TOBAL:=TOB.Create('',TOBL,-1) else TOBAL:=TOBL.Detail[0] ;
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
NumL:=TOBL.GetValue('GL_NUMLIGNE') ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'" AND YVA_IDENTLIGNE="'+FormatFloat('000',NumL)+'" ORDER BY YVA_AXE, YVA_NUMVENTIL',True) ;
if Not Q.EOF then TOBAL.LoadDetailDB('VENTANA','','',Q,False) ;
Ferme(Q) ;
END ;

Procedure UG_AjouteLesRepres ( TOBPiece,TOBComms : TOB ) ;
Var i : integer ;
    TOBL : TOB ;
    TypeCom : String ;
BEGIN
TypeCom:=GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_TYPECOMMERCIAL') ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    AjouteRepres(TOBL.GetValue('GL_REPRESENTANT'),TypeCom,TOBComms) ;
    END ;
END ;

Procedure UG_AjouteUnArticle ( TOBL,TOBArticles,TOBCaTalogu : TOB ; Manuelle : boolean ) ;
Var TOBCata : TOB ;
    RefCata,RefFour : String ;
BEGIN
if TOBL.GetValue('GL_ENCONTREMARQUE')='X' then
   BEGIN
   RefCata:=TOBL.GetValue('GL_REFCATALOGUE') ;
   RefFour:=GetCodeFourDCM(TOBL) ;
   TOBCata:=TOBCatalogu.FindFirst(['GCA_REFERENCE','GCA_TIERS'],[RefCata,RefFour],False) ;
   if TOBCata=Nil then TOBCata:=InitTOBCata(TOBCatalogu,RefCata,RefFour) ;
   LoadTOBDispoContreM(TOBL, TOBCata, true) ;
   END ;
END ;

function FindStInTAB(St:String;Tab:Array of string):boolean;
var i : integer;
begin
result:=false;
for i:=Low(Tab) to High(Tab) do
  if Pos(St,Tab[i])<>0 then begin result:=true; break; end;
end;

Procedure UG_ConstruireTobArt(TOBPiece,TOBArticles:Tob) ;
var i,NbArt,CountStArt,NbRequete,y : integer;
    StSelect,StSelectDepot,StArticle,StCodeArticle,ListeDepot,RefGen,Statut,TypeArt : string;
    TabWhere,TabWhereDepot : array of String;
    TOBDispo, TOBArt, TobDispoArt, TobDispoArtDeja, TobArtLoc : TOB;
    QArticle,QDepot : TQuery ;
    StSelectA : string;
begin
    TobArtLoc := Tob.Create ('ARTICLE', nil, -1);

    StSelect :='SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    					 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    					 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE';
    StSelectDepot := 'SELECT * FROM DISPO' ;
    CountStArt:=0; NbRequete:=0; ListeDepot:='';
    SetLength(TabWhere, 1); SetLength(TabWhereDepot, 1);
    ListeDepot:='"'+TOBPiece.GetValue('GP_DEPOT')+'"';
    For i:=0 to TOBPiece.Detail.Count-1 do
    begin
    	if TOBPiece.Detail[i].GetValue('GL_TYPELIGNE')<>'ART' then continue; 
        StArticle:=TOBPiece.Detail[i].GetValue('GL_ARTICLE') ;
        StCodeArticle:=TOBPiece.Detail[i].GetValue('GL_CODEARTICLE') ;
        RefGen:=TOBPiece.Detail[i].GetValue('GL_CODESDIM') ;
        Statut:=TOBPiece.Detail[i].GetValue('GL_TYPEDIM') ;
        TypeArt:=TOBPiece.Detail[i].GetValue('GL_TYPEARTICLE') ;
        if Pos(TobPiece.Detail[i].GetValue ('GL_DEPOT'), ListeDepot) = 0 then
        begin
            ListeDepot := ListeDepot + ', "' + TobPiece.Detail[i].GetValue ('GL_DEPOT') + '"';
        end;
        if CountStArt>=NbArtParRequete then
        begin
            NbRequete:=NbRequete+1;
            SetLength(TabWhere, NbRequete+1); SetLength(TabWhereDepot, NbRequete+1);
            CountStArt:=0;
        end;
        if Not FindStInTAB(StArticle,TabWhere) then
        begin
            if (Statut='GEN') or (Statut='DIM') or (Statut='NOR')  then
            begin
                if (StArticle='') and (RefGen<>'') then RefGen:=CodeArticleUnique2(RefGen,'') ;
                if StArticle<>'' then RefGen:=StArticle ;
                if TabWhere[NbRequete]='' then TabWhere[NbRequete] := '"'+RefGen+'"'
                else TabWhere[NbRequete] := TabWhere[NbRequete]+',"'+RefGen+'"';
            end ;
            if ((Statut='DIM') or (Statut='NOR')) and ((TypeArt<>'CTR') and (TypeArt<>'PRE') and (TypeArt<>'FRA')) then
            begin
                if TabWhereDepot[NbRequete]='' then TabWhereDepot[NbRequete] := '"'+StArticle+'"'
                else TabWhereDepot[NbRequete] := TabWhereDepot[NbRequete]+',"'+StArticle+'"';
            end;
            CountStArt:=CountStArt+1;
        end ;
    end;

    if TabWhere[0]<>'' then
    begin
        TOBDispo := TOB.CREATE ('Les Dispos', nil, -1);
        for y:=Low(TabWhere) to High(TabWhere) do
        begin
            if (TabWhere[y]<>'') and (TabWhere[y]<>'""') then
            begin
            		StSelectA := StSelect + ' WHERE GA_ARTICLE IN ('+TabWhere[y]+')';
                QArticle:=OpenSQL(StSelectA,True);
                if Not QArticle.EOF then TobArtLoc.LoadDetailDB('ARTICLE','','',QArticle,True);
                Ferme(QArticle);
            end;
            if TabWhereDepot[y]<>'' then
            begin
                QDepot:=OpenSQL(StSelectDepot+' WHERE GQ_ARTICLE IN ('+TabWhereDepot[y]+') AND GQ_DEPOT IN ('+ListeDepot+') AND GQ_CLOTURE="-"',True);
                if Not QDepot.EOF then TOBDispo.LoadDetailDB('DISPO','','',QDepot,True);
                Ferme(QDepot);
            end;
        end;

        { XP }
        if TobArtLoc.Detail.Count>0 then
        begin
            TobArtLoc.detail[0].AddChampSupValeur('UTILISE','-',True);
            TobArtLoc.detail[0].AddChampSupValeur('REFARTSAISIE','',True);
            TobArtLoc.detail[0].AddChampSupValeur('REFARTBARRE','',True);
            TobArtLoc.detail[0].AddChampSupValeur('REFARTTIERS','',True);
        end ;
        //Affecte les stocks aux articles sélectionnés
        for NbArt := 0 to TobArtLoc.Detail.Count - 1 do // si dans les pièces regroupées on a deux fois le meme article
        begin
            if TobArticles.FindFirst (['GA_ARTICLE'], [TobArtLoc.Detail[NbArt].GetValue ('GA_ARTICLE')], False) = nil then
            begin
                Tob.Create ('ARTICLE', TobArticles, -1).Dupliquer (TobArtLoc.Detail[NbArt], True, True);
            end;
        end;
        for NbArt:=0 to TOBArticles.detail.Count-1 do
        begin
            TOBArt:=TOBArticles.detail[NbArt];
            InitChampsSupArticle (TOBART);
            TobDispoArt:=TOBDispo.FindFirst(['GQ_ARTICLE'],[TOBArt.GetValue('GA_ARTICLE')],False) ;
            while TobDispoArt<>nil do
            begin
                TobDispoArtDeja := TobArt.FindFirst(['GQ_DEPOT'],[TobDispoArt.GetValue('GQ_DEPOT')],False);
                if TobDispoArtDeja = nil then // si le dispo n'existe pas déja sous l'article
                begin
                    DispoChampSupp (TobDispoArt);
                    TobDispoArt.Changeparent(TOBArt,-1);
                end;
                TobDispoArt:=TOBDispo.FindNext(['GQ_ARTICLE'],[TOBArt.GetValue('GA_ARTICLE')],False) ;
            end;
        end;
        TOBDispo.Free;
    end;
    TobArtLoc.Free;
end;

procedure UG_RegenAnalytique ( TOBPiece,TOBArticles,TOBAFFAIRE,TOBL,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu : TOB ; Manuelle : boolean ) ;
Var TOBArt,TOBC,TOBCata,TOBAffaires : TOB ;
    i    : integer ;
    RefU : String ;
    QQ : TQuery;
begin
  if TOBL.detail.count > 0 then TOBL.detail[0].ClearDetail;
  if TOBL.detail.count > 1 then TOBL.detail[1].ClearDetail;
  If TOBL.GetValue('GL_TYPEREF')='CAT' then TOBCata:=FindTOBCataRow(TOBPiece,TOBCataLogu,i+1)
                                       else TOBArt:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ;
  if TOBArt<>nil then
  begin
    TOBC:=ChargeAjouteCompta(TOBCpta,TOBPiece,TOBL,TOBArt,TOBTiers,TOBCata,TOBAFFAIRE,True) ;
    if TOBC <> nil then PreVentileLigne(TOBC,TOBAnaP,TOBAnaS,TOBL) ;
  end ;
end;

Procedure UG_AjouteLesArticles ( TOBPiece,TOBArticles,TOBCpta,TOBTiers,TOBAnaP,TOBAnaS,TOBCatalogu : TOB ; Manuelle : boolean ) ;
Var TOBL,TOBArt,TOBC,TOBCata,TOBAffaires,TOBAffaireL : TOB ;
    i    : integer ;
    RefU : String ;
    QQ : TQuery;
BEGIN
TOBAffaires := TOB.Create ('LES AFFAIRES',nil,-1);
UG_ConstruireTobArt(TOBPiece,TOBArticles) ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    TOBArt:=Nil; TOBCata:=Nil;
    RefU:=TOBL.GetValue('GL_ARTICLE') ;
    if (TOBL.GetValue('GL_TYPEREF')='CAT') OR (RefU<>'') OR (TOBL.GetValue('GL_TYPEDIM')='GEN') then
       BEGIN
       if TOBL.GetValue ('GL_AFFAIRE') <> '' then
       begin
       	TOBAffaireL := TOBaffaires.findFirst(['AFF_AFFAIRE'],[TOBL.GetValue('GL_AFFAIRE')],true);
        if TOBAffaireL = nil then
        begin
          QQ := OpenSql ('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="'+TOBL.GEtVALue('GL_AFFAIRE')+'"',true);
          if not QQ.eof then
          begin
        		TOBaffaireL := TOB.create('AFFAIRE',TOBAffaires,-1);
            TOBAffaireL.SelectDB ('',QQ);
          end;
          ferme (QQ);
        end;
       end;
       UG_AjouteUnArticle(TOBL,TOBArticles,TOBCataLogu,Manuelle) ;
       If TOBL.GetValue('GL_TYPEREF')='CAT' then TOBCata:=FindTOBCataRow(TOBPiece,TOBCataLogu,i+1)
                                            else TOBArt:=FindTOBArtRow(TOBPiece,TOBArticles,i+1) ;
       if TOBArt<>nil then
         begin
         TOBC:=ChargeAjouteCompta(TOBCpta,TOBPiece,TOBL,TOBArt,TOBTiers,TOBCata,TOBaffaireL,True) ;
         if TOBC <> nil then PreVentileLigne(TOBC,TOBAnaP,TOBAnaS,TOBL) ;
         end ;
       END ;
    END ;
TOBAffaires.free;
END ;

Procedure UG_MajAnalPiece ( TOBAnaP,TOBAnaS : TOB ; RefA : String ) ;
Var i : integer ;
    TOBV : TOB ;
BEGIN
if TOBAnaP=Nil then for i:=0 to TOBAnaP.Detail.Count-1 do
    BEGIN
    TOBV:=TOBAnaP.Detail[i] ;
    if TOBV.NomTable='VENTANA' then TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
    END ;
if TOBAnaS=Nil then for i:=0 to TOBAnaS.Detail.Count-1 do
    BEGIN
    TOBV:=TOBAnaS.Detail[i] ;
    if TOBV.NomTable='VENTANA' then TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
    END ;
END ;

Procedure UG_MajAnalLigne ( TOBL : TOB ; RefA : String ) ;
Var i,NumL,k : integer ;
    TOBV,TOBAL : TOB ;
BEGIN
NumL:=TOBL.GetValue('GL_NUMLIGNE') ;
for i:=0 to TOBL.Detail.Count-1 do
    BEGIN
    TOBAL:=TOBL.Detail[i] ;
    for k:=0 to TOBAL.Detail.Count-1 do
        BEGIN
        TOBV:=TOBAL.Detail[k] ;
        if TOBV.NomTable='VENTANA' then
           BEGIN
           TOBV.PutValue('YVA_IDENTIFIANT',RefA) ;
           TOBV.PutValue('YVA_IDENTLIGNE',FormatFloat('000',NumL)) ;
           END ;
       END ;
    END ;
END ;

Procedure UG_MajLesComms ( TOBPiece,TOBArticles,TOBComms : TOB ) ;
Var i : integer ;
    EtatVC : String ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    EtatVC:=TOBL.GetValue('GL_VALIDECOM') ;
    if EtatVC='VAL' then TOBL.PutValue('GL_VALIDECOM','AFF') else
     if EtatVC='AFF' then CommVersLigne(TOBPiece,TOBArticles,TOBComms,i+1,False) ;
    END ;
END ;

{$IFDEF BTP}
procedure UG_AjouteLesChampsBtp (TOBPiece : TOB);
//var TOBL : TOB;
//    Indice : integer;
begin
	AddLesSupEntete (TOBPiece);
{
for Indice:=0 to TOBPiece.detail.count -1 do
    begin
    TOBL := TOBPiece.detail[Indice];
    TOBL.AddChampSupValeur ('AFF_GENERAUTO',RenvoieTypeFact(TOBL.GetValue('GP_AFFAIREDEVIS')) ,false);
    end;
}
end;

{$ENDIF}


procedure UG_ChargePiecesPrecedente (TOBPiece,TOBGenere : TOB;GestionReliquat : boolean);
var Indice : integer;
    TOBL,TOBLP : TOB;
    cledoc : R_Cledoc;
    Q : TQuery;
//    Reste : double;
begin
  for Indice := 0 to TOBGenere.detail.count -1 do
  begin
    TOBL := TOBGenere.detail[Indice];
    if TOBL.GetValue('GL_PIECEPRECEDENTE') = '' then continue;
    if TOBL.GetValue('GL_TYPELIGNE') <> 'ART' then continue;
    DecoderefPiece (TOBL.GetValue('GL_PIECEPRECEDENTE'),Cledoc);
    Q := OpenSQL('SELECT * FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, true,True), True);
    if not Q.eof then
    begin
      TOBLP := TOB.Create ('LIGNE',TOBPiece,-1);
      TOBLP.SelectDB ('',Q);
(*
      if GestionReliquat then
      begin
        Reste := Arrondi(TobLP.GetValue('GL_QTERESTE') - TobL.GetValue('GL_QTESTOCK'), 6);
        if Reste <= 0 then
        BEGIN
          Reste := 0;  { Le reste à livrer n'est jamais négatif }
//          TOBLP.PutValue('GL_VIVANTE','-');
        END;
       // Mise à jour du reste à livrer
        TobLP.PutValue('GL_QTERESTE', Reste) ;
      end else
      begin
        TobLP.PutValue('GL_QTERESTE', 0) ;
//        TOBLP.PutValue('GL_VIVANTE','-');
      end;
*)
    end;
    ferme (Q);
  end;
end;


procedure UG_CreeLesTob (var TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                      TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TOBLiaison_o,TobLigneTarif_O: TOB);
begin
  // Series
  TOBSerie_O := TOB.Create ('LES SERIES',nil,-1);
  // Lecture bases
  TOBBases_O:=TOB.Create('BASES',Nil,-1) ;
  // Lecture Echéances
  TOBEches_O:=TOB.Create('Les ECHEANCES',Nil,-1) ;
  // Lecture Acomptes
  TOBAcomptes_O:=TOB.Create('',Nil,-1) ;
  // Lecture Ports
  TOBPorcs_O:=TOB.Create('',Nil,-1) ;
  // Lecture nomenclatures
  TOBN_O:=TOB.Create('NOMENCLATURES',Nil,-1) ;
  // Lecture Ouvrage
  TOBOuvrage_O:=TOB.Create('OUVRAGES',Nil,-1) ;
  // Lecture tEXTES DEBUT ET fin
  TOBLIENOLE_O:=TOB.Create('OLE',Nil,-1) ;
  // Retenues de garantie
  TOBPIECERG_O:=TOB.create('RETENUES',nil,-1);
  // Base de retenues de garantie
  TOBBasesRg_O:=TOB.create('BASESRG',nil,-1);
{$IFDEF AFFAIRE}
  //Affaire-ONYX
  TobVariable_O := TOB.Create('The Variables', nil, -1);
  //Affaire-ONYX
  TobRevision_O := TOB.Create('The REVISIONS', nil, -1);
{$ENDIF}
{$IFDEF BTP}
  Tobliaison_O := TOB.Create('The LIAISONS', nil, -1);
{$ENDIF}
  TobLigneTarif_O := TOB.Create('The TARIFS', nil, -1);
end;

procedure UG_DropLesTobs (TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                      TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLigneTarif_O : TOB);
begin
  // Series
  TOBSerie_O.free;
  // Lecture bases
  TOBBases_O.free;
  // Lecture Echéances
  TOBEches_O.free;
  // Lecture Acomptes
  TOBAcomptes_O.free;
  // Lecture Ports
  TOBPorcs_O.free;
  // Lecture nomenclatures
  TOBN_O.free;
  // Lecture Ouvrage
  TOBOuvrage_O.free;
  // Lecture tEXTES DEBUT ET fin
  TOBLIENOLE_O.free;
  // Retenues de garantie
  TOBPIECERG_O.free;
  // Base de retenues de garantie
  TOBBasesRg_O.free;
  //Affaire-ONYX
  TobVariable_O.free;
  //Affaire-ONYX
  TobRevision_O.free;
  //
  TobLigneTarif_O.free;

end;

procedure UG_VideLesTobs (TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                       TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLigneTarif_O : TOB);
begin
  // Series
  TOBSerie_O.ClearDetail;
  // Lecture bases
  TOBBases_O.ClearDetail;
  // Lecture Echéances
  TOBEches_O.ClearDetail;
  // Lecture Acomptes
  TOBAcomptes_O.ClearDetail;
  // Lecture Ports
  TOBPorcs_O.ClearDetail;
  // Lecture nomenclatures
  TOBN_O.ClearDetail;
  // Lecture Ouvrage
  TOBOuvrage_O.ClearDetail;
  // Lecture tEXTES DEBUT ET fin
  TOBLIENOLE_O.ClearDetail;
  // Retenues de garantie
  TOBPIECERG_O.ClearDetail;
  // Base de retenues de garantie
  TOBBasesRg_O.ClearDetail;
  //Affaire-ONYX
  TobVariable_O.ClearDetail;
  //Affaire-ONYX
  TobRevision_O.ClearDetail;
  //
  TobLigneTarif_O.clearDetail;
end;

procedure UG_ChargeLesTobs (TOBPiece,TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                       TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TOBLiaison_O,TOBArticles,
                       TobLigneTarif_O : TOB);
var cledoc : R_Cledoc;
    Q : Tquery;
    indiceOuv : integer;
begin
  CleDoc:=TOB2CleDoc(TOBPiece) ;
  // SERIE
  Q:=OpenSQL('SELECT * FROM LIGNESERIE WHERE '+WherePiece(CleDoc,ttdSerie,False),true) ;
  TOBSerie_O.LoadDetailDB('LIGNESERIE','','',Q,False) ;
  Ferme(Q) ;
  // Lecture bases
  Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDoc,ttdPiedBase,False),True) ;
  TOBBases_O.LoadDetailDB('PIEDBASE','','',Q,False) ;
  Ferme(Q) ;
  // Lecture Echéances
  Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CleDoc,ttdEche,False),True) ;
  TOBEches_O.LoadDetailDB('PIEDECHE','','',Q,False) ;
  Ferme(Q) ;
  // Lecture Acomptes
  Q:=OpenSQL('SELECT * FROM ACOMPTES WHERE '+WherePiece(CleDoc,ttdAcompte,False),True) ;
  TOBAcomptes_O.LoadDetailDB('ACOMPTES','','',Q,False) ;
  Ferme(Q) ;
  // Lecture Ports
  Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+WherePiece(CleDoc,ttdPorc,False),True) ;
  TOBPorcs_O.LoadDetailDB('PIEDPORT','','',Q,False) ;
  Ferme(Q) ;
  // Lecture nomenclatures
  LoadLesNomen(TOBPiece,TOBN_O,TOBArticles,CleDoc) ;
  // Modif BTP
  // Lecture Ouvrage
  {$IFDEF BTP}
  LoadLesOuvrages(TOBPiece,TOBOuvrage_O,TOBArticles,CleDoc,IndiceOuv) ;
  {$ENDIF}
  // Lecture tEXTES DEBUT ET fin
  Q:=OpenSQL('SELECT * FROM LIENSOLE WHERE '+WherePiece(CleDoc,ttdLienOle,False),True) ;
  TOBLIENOLE_O.LoadDetailDB('LIENSOLE','','',Q,False) ;
  Ferme(Q) ;
  Q:=OpenSQL('SELECT * FROM PIECERG WHERE '+WherePiece(CleDoc,ttdretenuG,False),True) ;
  TOBPieceRg_O.selectDB ('',Q) ;
  Ferme(Q) ;
  Q:=OpenSQL('SELECT * FROM PIEDBASERG WHERE '+WherePiece(CleDoc,ttdBaseRG,False),True) ;
  TOBBasesRg_O.selectDB ('',Q) ;
  Ferme(Q) ;
{$IFDEF BTP}
  // liens inter-document
  Q := OpenSQL('SELECT * FROM LIENDEVCHA WHERE BDA_REFD="'+ EncodeLienDevCHA(TOBPIECE)+'"',true);
  TOBliaison_O.LoaddetailDB ('LIENDEVCHA','','',Q,false) ;
  Ferme(Q) ;
{$ENDIF}
  //
{$IFDEF AFFAIRE}
  //Affaire-ONYX
  Q := OpenSQL('SELECT * FROM AFORMULEVARQTE WHERE ' + WherePiece(CleDoc, ttdVariable, False), True);
  TobVariable_O.LoadDetailDB('AFORMULEVARQTE', '', '', Q, False);
  Ferme(Q);
  //Affaire-ONYX
  Q := OpenSQL('SELECT * FROM AFREVISION WHERE ' + WherePiece(CleDoc, ttdRevision, False), True);
  TobRevision_O.LoadDetailDB('AFREVISION', '', '', Q, False);
  Ferme(Q);
{$ENDIF}
  { Chargement des détails tarif par ligne de pièce }
  LoadLesLignesTarifs(TobLigneTarif_O, CleDoc);
end;

procedure UG_ReajustePiecesPrecedente (TOBGenere,TOBPiece,TOBArticles: TOB) ;

  function ISAliveLine (TOBpiece : TOB) : boolean;
  var Requete : String;
      QQ : TQuery;
      cledoc : R_cledoc;
  begin
    result := false;
    cledoc := TOB2CleDoc (TOBPiece);
    Requete := 'SELECT GL_NUMERO FROM LIGNE WHERE '+WherePiece (Cledoc,ttdLigne,false) +
//               ' AND GL_VIVANTE="X" AND GL_TYPELIGNE="ART"';
               ' AND GL_QTERESTE <> 0 AND GL_TYPELIGNE="ART"';
    QQ := OpenSql (requete,true);
    if not QQ.eof then result := true;
    ferme (QQ);
  end;

  function ISPiecePreparable (TOBgenere,TOBLP : TOB) : boolean;
  begin
  	// si la piece generée est bien une prevision de chantier
    // et que la piece origine est soit un devis soit une contre etude
		result := (
    					(TOBGenere.GetValue ('GP_NATUREPIECEG')=GetParamSoc('SO_BTNATCHANTIER')) and
              	((TOBLP.GetValue('GP_NATUREPIECEG')=GetparamSoc('SO_AFNATAFFAIRE')) or
              	(TOBLP.GetValue('GP_NATUREPIECEG')='BCE')))
              or
              ((TOBgenere.GetValue ('GP_NATUREPIECEG')='BCE') and
             (TOBLP.GetValue('GP_NATUREPIECEG')=GetparamSoc('SO_AFNATAFFAIRE')));
  end;


var Indice,nb : integer;
    TOBPieceOrig,TOBL,TOBLP : TOB;
    // gestion Piece
    TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O : TOB;
    TobVariable_O,TobRevision_O,TOBLiaison_O,TobLigneTarif_O: TOB;
    OldEcr,OldStk   : RMVT ;
    NowFutur : TDateTime ;
    cledoc : R_CLEDOC;
begin
  TOBPieceOrig := TOB.Create ('LES PIECE ORIG', nil,-1);
  TRY
    if TOBPiece.detail.count > 0 then // traitement sur piece géré en reliquat
    begin
      for Indice := 0 To TOBPiece.detail.count -1 do
      begin
        TOBL := TOBPiece.detail[Indice];

        // Verifie que l'on essaye pas de mettre a jour des infos sur la piece courante
        if (TOBL.GetValue('GL_NATUREPIECEG') = TOBGenere.GetValue('GP_NATUREPIECEG')) and
           (TOBL.GetValue('GL_SOUCHE') = TOBGenere.GetValue('GP_SOUCHE')) and
           (TOBL.GetValue('GL_NUMERO') = TOBGenere.GetValue('GP_NUMERO')) and
           (TOBL.GetValue('GL_INDICEG') = TOBGenere.GetValue('GP_INDICEG'))  then continue;

        TOBLP := TOBPieceOrig.findFirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                        [TOBL.GetValue('GL_NATUREPIECEG'),
                                         TOBL.GetValue('GL_SOUCHE'),
                                         TOBL.GetValue('GL_NUMERO'),
                                         TOBL.GetValue('GL_INDICEG')],true);
        if TOBLP = nil then
        begin
          TOBLP := TOB.Create ('PIECE',TOBPieceOrig,-1);
          TOBLP.PutValue('GP_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG'));
          TOBLP.PutValue('GP_SOUCHE',TOBL.GetValue('GL_SOUCHE'));
          TOBLP.PutValue('GP_NUMERO',TOBL.GetValue('GL_NUMERO'));
          TOBLP.PutValue('GP_INDICEG',TOBL.GetValue('GL_INDICEG'));
          TOBLP.LoadDB (true);
        end;
        if not TOBL.UpdateDB then
        begin
          V_PGI.IoError := oeUnknown;
          exit;
        end;
      end;
      UG_CreeLesTob (TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                  TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TOBLiaison_o,TobLigneTarif_O);
      // mise à jour des pièces précédentes

			NowFutur:=NowH ;

      for Indice := 0 to TOBPieceOrig.detail.count -1 do
      begin
        TOBL := TOBPieceOrig.detail[Indice];
        if not ISAliveLine (TOBL) then
        begin
          TOBL.PutValue('GP_VIVANTE','-');
//
          if V_PGI.IoError=oeOk then DetruitCompta(TOBL,NowFutur,OldEcr,OldStk) ; // Est-ce bien nécéssaire ??
//
          if not HistoPiece(TOBL) then { NEWPIECE }
          begin
            UG_VideLesTobs (TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                          TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLigneTarif_O);
            UG_ChargeLesTobs (TOBL,TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                            TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLiaison_o,TOBArticles,TobLigneTarif_O);
{$IFDEF AFFAIRE}
            if not DetruitAffaireComplement(TobVariable_O, TobRevision_O) then
            begin
              V_PGI.IoError := oeSaisie;
              exit;
            end;
{$ENDIF}
            if not DetruitAncien(TOBL,TOBBases_O,TOBEches_O,TOBN_O,nil,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobLigneTarif_O) then
            begin
              V_PGI.IoError := oeSaisie;
              exit;
            end;
          end else
          begin
            TOBL.SetDateModif(NowFutur) ; TOBL.CleWithDateModif:=True ;
            if not TOBL.UpdateDB then begin V_PGI.IoError := oeUnknown; Exit; END;
          end;
 //
        end;
      end;
      UG_DropLesTobs(TOBBases_O,TOBEches_O,TOBN_O,TOBAcomptes_O,TOBPorcs_O,TOBSerie_O,TOBOuvrage_O,
                  TOBLienOle_O,TOBPieceRG_O,TOBBasesRg_O,TobVariable_O,TobRevision_O,TobLigneTarif_O);
    end;

    TOBPieceOrig.ClearDetail;

    if TOBGenere.detail.count > 0 then // traitement sur affaire de la piece d'origine
    begin
      for Indice := 0 To TOBGenere.detail.count -1 do
      begin
        TOBL := TOBGenere.detail[Indice];
        if (not (TOBL.FieldExists('PIECEORIGINE'))) or (Trim(TOBL.getValue('PIECEORIGINE')) = '') then
          continue;

        DecoderefPiece (TOBL.GetValue('PIECEORIGINE'),cledoc);
        TOBLP := TOBPieceOrig.findFirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                        [ cledoc.naturepiece  ,
                                          cledoc.souche,
                                          inttostr(cledoc.NumeroPiece),
                                         inttostr(cledoc.Indice)],true);
        if TOBLP = nil then
        begin
          TOBLP := TOB.Create ('PIECE',TOBPieceOrig,-1);
          TOBLP.PutValue('GP_NATUREPIECEG',cledoc.naturepiece);
          TOBLP.PutValue('GP_SOUCHE',cledoc.souche);
          TOBLP.PutValue('GP_NUMERO',cledoc.NumeroPiece);
          TOBLP.PutValue('GP_INDICEG',cledoc.Indice);
          TOBLP.LoadDB (true);
//          if (TOBGenere.GetValue ('GP_NATUREPIECEG')=GetParamSoc('SO_BTNATCHANTIER')) and
//              (TOBLP.GetValue('GP_NATUREPIECEG')=GetparamSoc('SO_AFNATAFFAIRE')) then
          if ISPiecePreparable (TOBgenere,TOBLP) then
          begin
            nb := ExecuteSQL ('UPDATE AFFAIRE SET AFF_PREPARE="X" WHERE AFF_AFFAIRE="'+TOBLP.GetValue('GP_AFFAIREDEVIS')+'"');
            if nb < 0 then begin V_PGI.IoError := oeUnknown; Exit; END;
          end;
        end;
      end;
    end;

  FINALLY
    TOBPieceOrig.free;
  END;
end;


procedure UG_DetruitLaPiece (TOBPiece: TOB) ; { NEWPIECE }
var nb :integer;
    cledoc : R_cledoc;
    oldecr,OldStk : RMVT;
    refA : string;
begin
  cledoc := TOB2CleDoc (TOBPiece);

  nb := ExecuteSQL ('DELETE FROM PIECE WHERE '+WherePiece(cledoc,ttdPiece,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM LIGNE WHERE '+WherePiece(cledoc,ttdLigne,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM LIGNECOMPL WHERE '+WherePiece(cledoc,ttdLigneCompl,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM PIEDECHE WHERE '+WherePiece(cledoc,ttdEche,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM PIEDBASE WHERE '+WherePiece(cledoc,ttdPiedBase,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM PIEDPORT WHERE '+WherePiece(cledoc,ttdPorc,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM LIGNENOMEN WHERE '+WherePiece(cledoc,ttdNomen,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM LIGNEOUV WHERE '+WherePiece(cledoc,ttdOuvrage,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM LIGNEFAC WHERE '+WherePiece(cledoc,ttdLignefac,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM LIGNESERIE WHERE '+WherePiece(cledoc,ttdSerie,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM PIECERG WHERE '+WherePiece(cledoc,ttdretenuG,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM PIEDBASERG WHERE '+WherePiece(cledoc,ttdBaseRG,false));
  if nb < 0 then V_PGI.IOerror := oePiece;

  nb := ExecuteSQL ('DELETE FROM PIECETRAIT WHERE '+WherePiece(cledoc,ttdPieceTrait,false));
  nb := ExecuteSQL ('DELETE FROM PIECEINTERV WHERE '+WherePiece(cledoc,ttdPieceInterv,false));
  //
{$IFDEF AFFAIRE}
  //Affaire-ONYX
  nb := ExecuteSQL ('DELETE FROM AFORMULEVARQTE WHERE '+WherePiece(cledoc,ttdVariable,false));
  if nb < 0 then V_PGI.IOerror := oePiece;
  //Affaire-ONYX
  nb := ExecuteSQL ('DELETE FROM AFREVISION WHERE '+WherePiece(cledoc,ttdRevision,false));
  if nb < 0 then V_PGI.IOerror := oePiece;
{$ENDIF}
//
  RefA := EncodeRefPresqueCPGescom(TOBPiece);

  ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');
  if isComptaStock(Cledoc.naturepiece) then ExecuteSQL('DELETE FROM VENTANA WHERE YVA_TABLEANA="GS" AND YVA_IDENTIFIANT LIKE "' + RefA + '"');

  if V_PGI.IoError = oeOk then DetruitCompta(TOBPiece, TOBpiece.getValue('GP_DATEMODIF'), OldEcr, OldStk);
end;


Procedure UG_LoadOuvrageLigne ( TOBL , TOBOuvrage, TOBArticles : TOB ; var IndiceNomen : integer ; RepriseAffaireRef : boolean) ;
Var i,OldL,Lig,k : integer ;
    TOBLN,TOBNomen,TOBGroupeNomen,TOBLoc : TOB ;
    Q    : TQuery ;
    CleDoc : R_CleDoc ;
    TypeArticle : string;
BEGIN
{$IFDEF BTP}
  if not NaturepieceOKPourOuvrage (TOBL) then exit;
{$ENDIF}
  OldL:=-1 ;
  TypeArticle := TOBL.GetValue('GL_TYPEARTICLE');
  if (TypeArticle<>'OUV') and (TypeArticle<>'ARP') then Exit ;
  //IndiceOuv:=1 ;
  TOBNomen:=TOB.Create('',Nil,-1) ; CleDoc:=TOB2CleDoc(TOBL) ;
  Q:=OpenSQL('SELECT * FROM LIGNEOUV WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+
  ' AND BLO_NUMLIGNE='+IntToStr(TOBL.GetValue('GL_NUMLIGNE'))+ ' ORDER BY BLO_N1,BLO_N2,BLO_N3,BLO_N4,BLO_N5',True) ;
  TOBNomen.LoadDetailDB('LIGNEOUV','','',Q,True,False) ;
  Ferme(Q) ;
{$IFDEF BTP}
  // VARIANTE
  SupprimeLesVariantes (TOBNomen);
{$ENDIF}
  // --
  if TobNomen.detail.count > 0 then
  begin
    i := 0;
    repeat
  //  for i:=0 to TOBNomen.Detail.Count-1 do
  //  BEGIN
      TOBLN:=TOBNomen.Detail[i] ;
      if RepriseAffaireRef then    // niveau 1
      begin
        TOBLN.PutValue('BLO_AFFAIRE', TOBL.GetValue('GL_AFFAIRE')) ;
        TOBLN.PutValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
        TOBLN.PutValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
        TOBLN.PutValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
        TOBLN.PutValue('BLO_AVENANT', TOBL.GetValue('GL_AVENANT')) ;
      end;
      Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
      if OldL<>Lig then
      BEGIN
        TOBGroupeNomen:=TOB.Create('',TOBOuvrage,-1) ;
        TOBGroupeNomen.AddChampSup('ANCPV',False) ; TOBGroupeNomen.PutValue('ANCPV',0) ;
        TOBGroupeNomen.AddChampSup('UTILISE',False) ; TOBGroupeNomen.PutValue('UTILISE','-') ;
        //       IndiceNomen := TOBOuvrage.detail.count;
        //       TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ;
        for k:=i to TOBNomen.Detail.Count-1 do
        BEGIN
          TOBLoc:=TOBNomen.Detail[k] ;
          if TOBLoc.GetValue('BLO_NUMLIGNE')<>Lig then Break ;
          if RepriseAffaireRef then
          begin
            TOBLOC.PutValue('BLO_AFFAIRE', TOBL.GetValue('GL_AFFAIRE')) ;
            TOBLOC.PutValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
            TOBLOC.PutValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
            TOBLOC.PutValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
            TOBLOC.PutValue('BLO_AVENANT', TOBL.GetValue('GL_AVENANT')) ;
          end;
        END ;
  {$IFDEF BTP}
        CreerLesTOBOuv(TOBL,TOBGroupeNomen,TOBNomen,TOBArticles,Lig,i) ;
  {$ENDIF}
      END ;
      OldL:=Lig ;
  //  END ;
    until i>= TOBNomen.Detail.Count;
    TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ; Inc(IndiceNomen) ;
  end else
  begin
  	TOBL.PutValue('GL_INDICENOMEN',0) ;
  end;
  TOBNomen.Free ;
END ;

function UG_IsAnalMultiple (TOBSource : TOB) : boolean;
var TOBA,TOBAP,TOBAL : TOB;
		Indice,II : integer;
    refA,Axe : string;
    QQ : TQuery;
    LastSection : string;
begin
	result := false;
  lastSection := '';
  TOBA := TOB.Create ('LES ANA',nil,-1);
  TOBAP := TOB.Create ('LES ANA / PIECE',nil,-1);
  TRY
  	for Indice := 0 to TOBSource.detail.count -1 do
    begin
    	TOBAP.ClearDetail;
      RefA := EncodeRefCPGescom(TOBSource.detail[Indice]);
      QQ := OpenSQl('SELECT YVA_AXE,YVA_IDENTLIGNE,YVA_SECTION FROM VENTANA WHERE '+
      							'YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="' + RefA + '" ORDER BY YVA_AXE,YVA_IDENTLIGNE',
      							true,-1,'',true);
      if not QQ.eof then
      begin
      	TOBAP.LoadDetailDB('VENTANA','','',QQ,false);
        II := 0;
        repeat
          Axe := TOBAP.detail[II].getValue('YVA_AXE');
          TOBAL := TOBA.findFirst(['AXE'],[Axe],true);
          if TOBAL = nil then
          begin
          	TOBAL := TOB.Create ('UN AXE',TOBA,-1);
            TOBAL.AddChampSupValeur('AXE',Axe);
          end;
          TOBAP.detail[II].ChangeParent(TOBAL,-1);
        until II >= TOBAP.detail.count;
      end else
      begin
      	result := true;
      end;
      ferme (QQ);
    	if result then break;
    end;
    // Parcours des lignes suivant les axes
    if not result then
    begin
      for Indice := 0 to TOBA.detail.count -1 do
      begin
        lastSection := '';
        for II := 0 to TOBA.detail[Indice].detail.count -1 do
        begin
          if II = 0 then
          begin
            lastSection := TOBA.detail[Indice].detail[II].getValue('YVA_SECTION');
          end else if LastSection <> TOBA.detail[Indice].detail[II].getValue('YVA_SECTION') then
          begin
            result := true;
            break;
          end;
        end;
        if result then break;
      end;
    end;
  FINALLY
  	FreeAndNil(TOBA);
  	FreeAndNil(TOBAP);
  END;
end;

Procedure UG_LoadAnaLigneOuPas ( TOBPiece,TOBL : TOB ) ;
Var Q : TQuery ;
    RefA : String ;
    NumL : integer ;
    TOBAL : TOB ;
BEGIN
if TOBL=Nil then Exit ;
if TOBL.Detail.Count<=0 then TOBAL:=TOB.Create('',TOBL,-1) else TOBAL:=TOBL.Detail[0] ;
RefA:=EncodeRefCPGescom(TOBPiece) ; if RefA='' then Exit ;
NumL:=TOBL.GetValue('GL_NUMLIGNE') ;
Q:=OpenSQL('SELECT * FROM VENTANA WHERE YVA_TABLEANA="GL" AND YVA_IDENTIFIANT="'+RefA+'" '+
					 'AND ((YVA_IDENTLIGNE="'+FormatFloat('000',NumL)+'") OR (YVA_IDENTLIGNE="000")) ORDER BY YVA_AXE, YVA_NUMVENTIL',True) ;
if Not Q.EOF then TOBAL.LoadDetailDB('VENTANA','','',Q,False) ;
Ferme(Q) ;
END ;

end.
