unit AfUtilArticle;


interface
Uses
{$IFDEF EAGLCLIENT}
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      Hent1,HCtrls,utob,UtilArticle,EntGC
     ;
  // fct pour condition plus sur article
Function PlusTypeArticle (bPreFraMarOnly:boolean=FALSE):string;
Function PlusTypeArticleText : string;
Function TrouverArticleSQL_GI (AcbRefUnique:boolean; RefSais : String ; TOBArt : TOB; AcsTypeArticle:string ) : T_RechArt ;


implementation



  {***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 05/03/2002
Modifié le ... :   /  /
Description .. : permet de renseigner la condition plus sur la tablette des
Suite ........ : typearticle en fct de natures gérées dans l'appli (Btp, Affaire,
Suite ........ : Gi ...)
Mots clefs ... : TYPEARTICLE;CONDITIONPLUS
*****************************************************************}
// $$$JP 02/06/03: pouvoir ne lister que PRE, FRA et MAR
//Function PlusTypeArticle : string  ;
function PlusTypeArticle (bPreFraMarOnly : boolean = FALSE) : string;
begin
{$ifndef BTP}
  // PRE,MAR;FRA
  //  result := 'AND (CO_CODE="PRE" OR CO_CODE="MAR" OR CO_CODE="FRA" OR CO_CODE="POU"';
  Result := ' AND (CO_CODE="PRE" OR CO_CODE="MAR" OR CO_CODE="FRA"';
  if bPreFraMarOnly = FALSE then
    begin
      // + POU
      Result := Result + ' OR CO_CODE="POU"';

//      if (ctxTempo in V_PGI.PGIContexte) or (ctxGCAff in V_PGI.PGIContexte) then
//        Result := Result + ' OR CO_CODE="CTR"';
    end;

  // + CTR pour Affaire
  // PL le 23/10/03 : Si on veut le type contrat en GA ou Gescom, on veut le type contrat !
  if (ctxTempo in V_PGI.PGIContexte) or (ctxGCAff in V_PGI.PGIContexte) then
    Result := Result + ' OR CO_CODE="CTR"';

  Result := Result + ')';
{$endif}
end;
// $$$JP fin

function PlusTypeArticleText : string  ;
begin
// pour renseigner la zone texte d'un MulitValcombobox
{$ifndef BTP}
     // PRE,MAR;FRA et POU
  result := 'FRA;MAR;POU;PRE';
  // + CTR pour Affaire
  //mcd 24/04/02 if ctxTempo in V_PGI.PGIContexte then result:=result + ';CTR';
  if not(ctxScot in V_PGI.PGIContexte) then result:=result + ';CTR';
{$endif}
end;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Fonction de remplissage d'une TOB Article suivant le code
Suite ........ : article passé en entrée, le type de code : unique ou pas
Suite ........ : (ARTICLE ou CODEARTICLE), et le type d'article s'il est
Suite ........ : précisé
Mots clefs ... : ARTICLE; TOB; RECHERCHE ; GIGA
*****************************************************************}
Function TrouverArticleSQL_GI (AcbRefUnique:boolean; RefSais : String ; TOBArt : TOB; AcsTypeArticle:string ) : T_RechArt ;
Var Q          : TQuery ;
    sReq       : string;
BEGIN
Result:=traAucun ;
if ((RefSais='') or (TOBArt=Nil)) then Exit ;
Q := nil;
try
// Recherche via code unique ou générique
if AcbRefUnique then
  // SELECT * : on ne ramène qu'un article et c'est une fonction générique dont on ne peut présumer des
  // éléments absoluments nécessaires en retour
    sReq := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+RefSais+'"'
else
    sReq := 'SELECT * FROM ARTICLE WHERE GA_CODEARTICLE="'+RefSais+'"';
if (AcsTypeArticle<>'') then sReq := sReq + ' AND GA_TYPEARTICLE="'+AcsTypeArticle+'"';
Q:=OpenSQL(sReq, True,-1,'',true) ;
if Not Q.EOF then
   BEGIN
   TOBArt.SelectDB('',Q) ;
   Result:=traOk ;
   END ;
finally
Ferme(Q) ;
end;
END ;

end.

