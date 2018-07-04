unit UtilStock;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,HStatus,Hdimension,UTOB,UTOM,AGLInit,
{$IFDEF EAGLCLIENT}
      emul,Maineagl,
{$ELSE}
      dbTables,DBGrids,db,Fe_Main,
{$ENDIF}
      LookUp,EntGC,ParamSoc ;

function RecalculStockDate(Article,DateStock:string; Depot:string='???'; Depart:integer=1; MvtJinclus:boolean=False;  TobRetour:tob=nil ) : double ;

implementation

Const
    // Paramètre Départ : recalcul depuis clôture ou depuis stock actuel
    STOCK_CLOTURE     = 1 ;
    STOCK_ACTUEL      = 2 ;


// NaturePiece affecte-t-elle la quantité TypeQte ?
function GetPlusMoins(TypeQte,QtePlus,QteMoins : string) : integer ;
begin
if Pos(TypeQte,QtePlus)>0 then result:=1
else if Pos(TypeQte,QteMoins)>0 then result:=-1
else result:=0 ;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 21/01/2002
Modifié le ... :   /  /
Description .. : Recalcul du stock à une date donnée
Mots clefs ... : ARTICLE;STOCK;RECALCUL
*****************************************************************}
function RecalculStockDate(Article,DateStock:string; Depot:string='???'; Depart:integer=1; MvtJinclus:boolean=False;  TobRetour:tob=nil ) : double ;
var  QQ,QDispo,QLigne : TQuery ;
     LigneATraiter,LigneVivante : boolean ;
     NbLig,Sens,iTob : integer ;
     StatutArt,CodeArticle,stWhereArt,stDateMax,stSql,stNature,stMvtJinclus : string ;
     stQtePlus,stQteMoins : string ;
     TobDispo,TobD : TOB ;
begin
{
Paramètres :
-	Article : article en taille unique, dimensionné ou générique.
-	DateStock : date à laquelle on souhaite connaître le stock.
-	Dépôt : (facultatif) dépôt du stock interrogé.
-	Départ : (facultatif) précise si le recalcul commence à la première clôture postérieure
            ou si l'on repart du stock actuel. Par défaut, recalcul depuis clôture.
-	TobRetour : (facultatif) si une tob est renseignée, tous les compteurs (quantités, pmap et prmp)
            seront recalculés et renseignés dans cette tob. Par défaut, pas de Tob.
Retour :
-	La quantité en stock,
-	Tous les compteurs ainsi que les pmap et prmp de la fiche stock via le paramètre TobRetour s'il est défini.
}
Result:=-999 ;

// Lecture statut + code article
QQ:=OpenSQL('select GA_STATUTART,GA_CODEARTICLE from ARTICLE where GA_ARTICLE="'+ CodeArticle +'"',True);
if QQ.EOF then BEGIN Ferme(QQ) ; exit END ;
StatutArt:=QQ.findField('GA_STATUTART').asString ;
CodeArticle:=QQ.findField('GA_CODEARTICLE').asString ;
Ferme(QQ) ;

if Depot='???' then Depot:=GetParamSoc('SO_GCDEPOTDEFAUT') ;
if StatutArt<>'GEN' then stWhereArt:=' GQ_ARTICLE="'+Article+'"'
else stWhereArt:=' GQ_ARTICLE in (select GA_ARTICLE from ARTICLE where GA_STATUTART<>"GEN" and GA_CODEARTICLE="'+CodeArticle+'")' ;


// Recherche première clôture suivant la date d'arrêté de stock
QDispo:=nil ;
if Depart=STOCK_CLOTURE then
    BEGIN
    stSql:='select GQ_ARTICLE,GQ_DATECLOTURE,GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,' +
           'GQ_PREPACLI,GQ_LIVRECLIENT,GQ_LIVREFOU,GQ_TRANSFERT,GQ_CUMULSORTIES,' +
           'GQ_CUMULENTREES,GQ_VENTEFFO,GQ_ECARTINV,GQ_PMAP,GQ_PMRP ' +
           'from DISPO ' +
           'where'+stWhereArt+' and GQ_DEPOT="'+Depot+'" and GQ_CLOTURE="X" ' +
           'and GQ_DATECLOTURE=(select min(GQ_DATECLOTURE) from DISPO where'+stWhereArt +
                               ' and GQ_DEPOT="'+Depot+'" and GQ_CLOTURE="X"' +
                               ' and GQ_DATECLOTURE>="'+USDateTime(StrToDate(DateStock))+'")' ;
    QDispo:=OpenSQL(stSql,True) ;
    // Si pas de clôture trouvée, on repart du stock actuel.
    if QDispo.EOF then BEGIN Depart:=STOCK_ACTUEL ; Ferme(QDispo) END
    else stDateMax:=QDispo.findField('GQ_DATECLOTURE').AsString ;
    END ;
if Depart=STOCK_ACTUEL then
    BEGIN
    stSql:='select GQ_ARTICLE,GQ_DATECLOTURE,GQ_PHYSIQUE,GQ_RESERVECLI,GQ_RESERVEFOU,' +
           'GQ_PREPACLI,GQ_LIVRECLIENT,GQ_LIVREFOU,GQ_TRANSFERT,GQ_QTE1,GQ_QTE2,' +
           'GQ_QTE3,GQ_VENTEFFO,GQ_ENTREESORTIES,GQ_ECARTINV,GQ_PMAP,GQ_PMRP ' +
           'from DISPO ' +
           'where'+stWhereArt+' and GQ_DEPOT="'+Depot+'" and GQ_CLOTURE="-"' ;
    QDispo:=OpenSQL(stSql,True) ;
    stDateMax:=DateToStr(V_PGI.DateEntree) ;
    END ;
if QDispo=nil then exit
else if QDispo.EOF then BEGIN Ferme(QDispo) ; exit END ;
TobDispo:=TOB.Create('DISPO',Nil,-1) ;
TobDispo.LoadDetailDB('DISPO','','',QDispo,False) ;
Ferme(QDispo) ;

// Recherche nombre de lignes traitées -> pour la fenêtre de patience
if MvtJinclus then stMvtJinclus:='>=' else stMvtJinclus:='>' ;
stSql:='select count(*) as NBLIG ' +
       'from LIGNE left join PARPIECE on GL_NATUREPIECEG=GPP_NATUREPIECEG ' +
       'where (GPP_QTEMOINS<>"" or GPP_QTEPLUS<>"")' +
       ' and'+stWhereArt +
       ' and GL_TENUESTOCK="X"' +
       ' and GL_QUALIFMVT<>"ANN"' +
       ' and GL_DEPOT="'+Depot+'"' +
       ' and GL_DATEPIECE'+stMvtJinclus+'"'+USDateTime(StrToDate(DateStock))+'"' +
       ' and GL_DATEPIECE<="'+USDateTime(StrToDate(stDateMax))+'"' ;
QQ:=OpenSQL(stSql,True) ;
if not QQ.Eof then NbLig:=QQ.FindField('NBLIG').AsInteger else NbLig:=0 ;
Ferme(QQ) ;

stSql:='select GL_ARTICLE,GL_QUALIFQTESTO,GL_QUALIFQTEACH,GL_QUALIFQTEVTE,' +
       'GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_PIECEPRECEDENTE,' +
       'GL_QTEFACT,GL_QTERESTE,GL_VIVANTE,GL_DPA,GL_DPR,GPP_VENTEACHAT,GPP_QTEPLUS,GPP_QTEMOINS' +
       'from LIGNE left join PARPIECE on GL_NATUREPIECEG=GPP_NATUREPIECEG ' +
       'left join PIECE on GP_NATUREPIECEG=GL_NATUREPIECEG and GP_SOUCHE=GL_SOUCHE ' +
       'and GP_NUMERO=GL_NUMERO and GP_INDICEG=GL_INDICEG ' +
       'where (GPP_QTEMOINS<>"" or GPP_QTEPLUS<>"")' +
       ' and'+stWhereArt +
       ' and GL_TENUESTOCK="X"' +
       ' and GL_QUALIFMVT<>"ANN"' +
       ' and GL_DEPOT="'+Depot+'"' +
       ' and GL_DATEPIECE'+stMvtJinclus+'"'+USDateTime(StrToDate(DateStock))+'"' +
       ' and GL_DATEPIECE<="'+USDateTime(StrToDate(stDateMax))+'"' +
       ' order by GL_DATEPIECE desc,GP_HEURECREATION desc' ;
QLigne:=OpenSQL(stSql,False) ;

InitMove(NbLig,'') ;
QLigne.FindFirst ;
while not QLigne.Eof do
    BEGIN
    // Recherche TobDispo de l'article
    TobD:=TobDispo.FindFirst(['GQ_ARTICLE'],[QLigne.FindField('GL_ARTICLE').AsString],False)  ;
    LigneVivante:=QLigne.FindField('GL_VIVANTE').AsString='X' ;
    stNature:=QLigne.FindField('GL_NATUREPIECEG').AsString ;

    // Traitements des seules lignes VIVANTES pour CdeFou,LivreFou,CdeCli,PrepaCli,LivreCli ;
    //             de toutes les lignes pour les autres natures de pièce.
    LigneATraiter:=(TobD<>nil) ;
    if (not LigneVivante) and (pos(stNature,'CF;BLF;CC;PRE;BLC')>0) then LigneATraiter:=False ;
    if LigneATraiter then
        BEGIN
        stQtePlus:=QLigne.FindField('GPP_PLUS').AsString ;
        stQteMoins:=QLigne.FindField('GPP_MOINS').AsString ;
{
        // Maj GQ_PMAP et GQ_PMRP si le mvt a incrémenté le stock : A FAIRE ...
        --> GPP_MAJPRIXVALO
}

        Sens:=GetPlusMoins('ESE',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_ENTREESORTIES',
                TobD.GetValue('GQ_ENTREESORTIES')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('INV',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_ECARTINV',
                TobD.GetValue('GQ_ECARTINV')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('LB1',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_QTE1',
                TobD.GetValue('GQ_QTE1')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('LB2',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_QTE2',
                TobD.GetValue('GQ_QTE2')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('LB3',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_QTE3',
                TobD.GetValue('GQ_QTE3')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('LC',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_LIVRECLIENT',
                TobD.GetValue('GQ_LIVRECLIENT')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('LF',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_LIVREFOU',
                TobD.GetValue('GQ_LIVREFOU')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('PHY',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_PHYSIQUE',
                TobD.GetValue('GQ_PHYSIQUE')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('PRE',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_PREPACLI',
                TobD.GetValue('GQ_PREPACLI')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('RC',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_RESERVECLI',
                TobD.GetValue('GQ_RESERVECLI')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('RF',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_RESERVEFOU',
                TobD.GetValue('GQ_RESERVEFOU')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('TRA',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_TRANSFERT',
                TobD.GetValue('GQ_TRANSFERT')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        Sens:=GetPlusMoins('VFO',stQtePlus,stQteMoins) ;
        if (Sens<>0) then
            TobD.PutValue('GQ_VENTEFFO',
                TobD.GetValue('GQ_VENTEFFO')-(QLigne.FindField('GL_QTEFACT').AsInteger*Sens)) ;

        END ;

    MoveCur(False) ;
    QLigne.next ;
    END ;
FiniMove ;
Ferme(QLigne) ;

// Maj du code retour
if TobRetour<>nil then TobRetour.Dupliquer(TobDispo,True,True,True) ;
if StatutArt='GEN' then
    BEGIN
    Result:=0.0 ;
    for iTob:=0 to TobDispo.Detail.Count-1
        do Result:=Result+TobDispo.Detail[iTob].GetValue('GQ_PHYSIQUE') ;
    END
    else Result:=TobDispo.GetValue('GQ_PHYSIQUE') ;
TobDispo.Free ;

end;

end.
