unit ModifTableLigOuv;

interface
Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtDoc, DB, Fe_Main,
{$ENDIF}
{$IFDEF BTP}
  BTPUtil,
{$ENDIF}
     SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit, FactUtil,factcalc,factcomm,factcpta,
     Math, StockUtil, EntGC, Classes, HMsgBox, Choix,NomenUtil,LigNomen,paramsoc,TarifUtil,utofbtanaldev,
     UtilArticle,
     FactTOB,uEntCommun ,UtilTOBPiece;

procedure ReinitAncOuvrage;

implementation

Procedure CreerLesNewTOBOuv (TOBpiece,TOBGroupeNomen,TOBNomen: TOB ; LaLig,MaxNiv,idep : integer ) ;
Var LeNiv,Niv,i,Lig : integer ;
    TOBLN,TOBPere,TOBLoc : TOB ;
    LigNiveau1,LigNiveau2,LigNiveau3,LigNiveau4,LigNIveau5 : integer;
    Pniveau : ^Integer;
BEGIN
Pniveau := Nil;
LigNiveau1 := 0;
LigNiveau2 := 0;
LigNiveau3 := 0;
LigNiveau4 := 0;
LigNiveau5 := 0;
for LeNiv:=1 to MaxNiv do
    BEGIN
    for i:=idep to TOBNomen.Detail.Count-1 do
        BEGIN
        TOBLN:=TOBNomen.Detail[i] ;
        Niv:=TOBLN.GetValue('BLO_NIVEAU') ; Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
        if Lig<>LaLig then Continue ;
        if Niv=LeNiv then
           BEGIN
           if Niv = 1 then Pniveau := @LigNiveau1
              else if Niv=2 then PNiveau := @LigNiveau2
                   else if Niv=3 then PNiveau := @LigNiveau3
                        else if Niv=4 then PNiveau := @LigNiveau4
                             else if Niv=5 then PNiveau := @LigNiveau5;
           if Niv=1 then
              BEGIN
              TOBPere:=TOBGroupeNomen;
              LigNiveau1 := TOBPere.detail.count;
              end else
              BEGIN
              TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_NIVEAU','BLO_ARTICLE','BLO_NUMORDRE'],[Lig,Niv-1,TOBLN.GetValue('BLO_COMPOSE'),TOBLN.GetValue('BLO_ORDRECOMPO')],True) ;
              LigNiveau1 := TOBPere.Getvalue ('BLO_N1');
              LigNiveau2 := TOBPere.Getvalue ('BLO_N2');
              LigNiveau3 := TOBPere.Getvalue ('BLO_N3');
              LigNiveau4 := TOBPere.Getvalue ('BLO_N4');
              LigNiveau5 := TOBPere.Getvalue ('BLO_N5');
              END ;
           if TOBPere<>Nil then
              BEGIN
              TOBLoc:=TOB.Create('LIGNEOUV',TOBPere,-1) ;
              Pniveau^ := TOBPere.detail.count;
              TOBLoc.Dupliquer(TOBLN,False,True) ;
              TOBLOC.putvalue ('BLO_N1',Ligniveau1);
              TOBLOC.putvalue ('BLO_N2',Ligniveau2);
              TOBLOC.putvalue ('BLO_N3',Ligniveau3);
              TOBLOC.putvalue ('BLO_N4',Ligniveau4);
              TOBLOC.putvalue ('BLO_N5',Ligniveau5);
              //
              TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
              TOBLOC.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
              //
              inc (Pniveau^);
              END ;
           END ;
        END ;
    END ;
END ;

Procedure LoadLesAncOuvrages ( TOBPiece,TOBOuvrage: TOB ; CleDoc : R_CleDoc );
Var i,OldL,Lig,MaxNiv,k,Niv: integer ;
    TOBL,TOBLN,TOBNomen,TOBGroupeNomen,TOBLoc : TOB ;
    OkN  : boolean ;
    Q    : TQuery ;
BEGIN
OkN:=False ; OldL:=-1 ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL := TOBPiece.detail[I];
    if (copy(TOBL.GetValue('GL_TYPEARTICLE'),1,2)='OU') then OkN:=True ;
    END ;
if Not OkN then Exit ;

TOBNomen:=TOB.Create('',Nil,-1) ;
Q:=OpenSQL('SELECT * FROM LIGNEOUV WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+' ORDER BY BLO_NUMLIGNE, BLO_NIVEAU, BLO_NUMORDRE,BLO_ORDRECOMPO',True,-1, '', True) ;
TOBNomen.LoadDetailDB('LIGNEOUV','','',Q,True,False) ;
Ferme(Q) ;
for i:=0 to TOBNomen.Detail.Count-1 do
    begin
    TOBLN:=TOBNomen.Detail[i] ;
    TOBLN.PutValue ('BLO_N1',0);
    TOBLN.PutValue ('BLO_N2',0);
    TOBLN.PutValue ('BLO_N3',0);
    TOBLN.PutValue ('BLO_N4',0);
    TOBLN.PutValue ('BLO_N5',0);
    end;
for i:=0 to TOBNomen.Detail.Count-1 do
    BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if OldL<>Lig then
       BEGIN
       TOBGroupeNomen:=TOB.Create('',TOBOuvrage,-1) ; MaxNiv:=-1 ;
       TOBGroupeNomen.AddChampSup('UTILISE',False) ;
       TOBGroupeNomen.PutValue('UTILISE','-') ;
       for k:=i to TOBNomen.Detail.Count-1 do
           BEGIN
           TOBLoc:=TOBNomen.Detail[k] ;
           if TOBLoc.GetValue('BLO_NUMLIGNE')<>Lig then Break ;
           Niv:=TOBLoc.GetValue('BLO_NIVEAU') ; if Niv>MaxNiv then MaxNiv:=Niv ;
           END ;
       CreerLesNewTOBOuv(TOBpiece,TOBGroupeNomen,TOBNomen,Lig,MaxNiv,i) ;
       END ;
    OldL:=Lig ;
    END ;
TOBNomen.Free ;
END ;

procedure ReinitAncOuvrage;
var TOBPieces,TOBPiece,TOBOuvrage : TOB;
    Indice : integer;
    Q : TQuery;
    cledoc : R_CLEDOC;
begin
TOBPieces := TOB.Create ('LES PIECES',nil,-1);
TOBOuvrage := TOB.Create ('LES OUVRAGES',nil,-1);
Q := OpenSql ('SELECT * FROM PIECE',true,-1, '', True);
TOBPIeces.loadDetailDb ('LES PIECES','','',Q,false);
ferme (Q);
TRY
for Indice := 0 to TOBPieces.detail.count -1 do
    begin
    TOBPiece := TOBPieces.detail[Indice];
    CleDoc.NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
    CleDoc.DatePiece := TOBPiece.GetValue('GP_DATEPIECE');
    CleDoc.Souche := TOBPiece.GetValue('GP_SOUCHE');
    CleDoc.NumeroPiece := TOBPiece.GetVAlue('GP_NUMERO');
    cledoc.indice := TOBPiece.GetVAlue('GP_INDICEG');
    Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True,-1, '', True) ;
    TOBPiece.LoadDetailDB('LIGNE','','',Q,False,True) ;
    Ferme(Q) ;
    LoadLesAncOuvrages (TOBPiece,TOBOuvrage,cledoc);
    if TOBOuvrage.detail.count > 0 then
       BEGIN
       ExecuteSQL ('DELETE FROM LIGNEOUV WHERE BLO_NATUREPIECEG="'+cledoc.NaturePiece
               +'" AND BLO_SOUCHE="'+cledoc.Souche +'" AND BLO_NUMERO='+inttostr(cledoc.NumeroPiece));
       TOBOuvrage.InsertDB (nil,true);
       END;
    TOBpiece.UpdateDB(false);
    TOBPiece.clearDetail;
    TOBOuvrage.clearDetail;
    end;
FINALLY
TOBPieces.free;
TOBouvrage.free;
end;
end;

end.
