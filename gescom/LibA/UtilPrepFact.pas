///////////////////////////////////////////////
// fonctions communes preparation de facture //
///////////////////////////////////////////////
unit UtilPrepFact;

interface

uses SysUtils, Hent1, HCtrls, AGLInit,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}

  UTOB, SAISUTIL, entGC, ParamSoc, FactComm,factarticle,
  FactUtil, UtilPGI,FactPiece,Factadresse,FactTob, Formule,UtilArticle,uEntCommun,UtilTOBPiece;

  // uniques
  procedure RemplirTOBTiersBis(pTobTiers: Tob; CodeTiers: string; pCleDoc: R_CleDoc);
  procedure RenumLesLignes(pTobPiece: TOB);
  procedure RecalculeSousTotaux(pTobPiece: Tob);
  procedure GestionDuVisa(pTobPiece: Tob);
  procedure MajTobRev(pTobRevision, pTobLig : TOB; pInNum: Integer; pBoOrigine : Boolean);
  Procedure MajVariable(pTobNewPiece, pTobVarQte_LIG : Tob);
  procedure LoadTobVarQte(pTob, pTobAffaire : TOB; pStOri : String; CleDocOrigine : R_CleDoc);
  procedure AlimTobVarQte(TobVarQte_LIG, TobLig, TobAct, TobVarQte: TOB; stOri : string; var pInCpt : Integer);
  function  RechVariable(TobVarQte,TobLig,TobAct:TOB ; stOri : String):TOB;
  procedure FinalisationTobVar(TobVarQte_LIG, TobVarDet, Toblig : TOB; var pInCpt : Integer);
  procedure MajBasePort(pTobPiece, pTobPorts: Tob);
  Procedure GestionChampSuppLigne(Toblignes : TOB;zaff,zdat,zcli,AffRef : string);
  
  // en double
  procedure CreationLignePiece(pTobPiece, pTobLigne, pTobAffaire, pTobArticles,
                               pTobTiers: TOB; pRdQte, pRdPuV: double; pStLib,
                               pStOri, pStAffRef: string; pCleDoc: R_CleDoc);

procedure ValideLaPiece(pTOBPiece, pTOBEches, pTobTiers, pTobPorts, pTobBases,
                        pTOBBasesL,pTOBNomenclature, pTobAdresses: Tob;
                        pInNewNum: Integer; pCleDoc: R_CleDoc;pDev: RDEVISE);

  function RechNumeroPiece(pStNat: string): Integer;

  function TestSiCumulArticle(pTobLig: TOB; pBoCumul: Boolean;
    pInPres, pInFour, pInFrais, pInCtr: Integer): Boolean;


implementation

{***********A.G.L.***********************************************
Auteur  ...... : GM - CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : déplacé de AFPrepFact
Mots clefs ... :
*****************************************************************}
procedure RemplirTOBTiersBis(pTobTiers: Tob; CodeTiers: string; pCleDoc: R_CleDoc);
var
  Q: TQuery;
  isFerme, Auto: boolean;
  TOBT: TOB;
  Req, zchp: string;

begin
  if CodeTiers = '' then
  begin
    pTobTiers.ClearDetail;
    pTobTiers.InitValeurs(False);
  end
  else if CodeTiers <> pTobTiers.GetValue('T_TIERS') then
  begin

    if not VH_GC.GCIfDefCEGID then
    //***Création des pieces, il faut tout prendre **** ?? gm à vérifier
    // si ca marche pour CEGID, pourquoi ca ne marcherait pas pour les autres ?
       Q := OpenSQL('SELECT * FROM TIERS WHERE T_TIERS="' + CodeTiers + '"', True)
    else
    begin
    // ?? mcd 21/08/03 manque T_tarifTiers si tarif sur catégorie client ??
    zchp := 'T_TIERS,T_AUXILIAIRE,T_FACTURE,T_PAYEUR,T_DEVISE,T_LIBELLE,T_PRENOM';
    zchp := zchp + ',T_JURIDIQUE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE3,T_CODEPOSTAL';
    zchp := zchp + ',T_VILLE,T_PAYS,T_TELEPHONE,T_EURODEFAUT';
    Req := 'SELECT ' + zchp + '  FROM TIERS WHERE T_TIERS="' + CodeTiers + '"';
    Q := OpenSQL(Req, True);
    end;

    pTobTiers.SelectDB('', Q);
    Ferme(Q);

    //***Création des pieces, il faut tout prendre ****
    Q := OpenSQL('SELECT * FROM TIERSPIECE WHERE GTP_TIERS="' + CodeTiers + '" AND GTP_NATUREPIECEG="' + pCleDoc.NaturePiece + '"', True);
    pTOBTiers.LoadDetailDB('TIERSPIECE', '', '', Q, False);
    Ferme(Q);
  end;

  isFerme := (pTobTiers.GetValue('T_FERME') = 'X');
  Auto := True;
  if isFerme then
  begin
    TOBT := pTobTiers.FindFirst(['GTP_NATUREPIECEG'], [pCleDoc.NaturePiece], False);
    if TOBT <> nil then
      Auto := (TOBT.GetValue('GTP_SUSPENSION') = 'X')
    else
      Auto := False;
  end;

  if not Auto then
  begin
    pTOBTiers.ClearDetail;
    pTOBTiers.InitValeurs(False);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GM
Créé le ...... :
Modifié le ... :
Description .. :  Alimentation de la ligne par rapport aus données de ARTICLE, etc...
                  zori :  A (acompte,ligne affaire)

                          T (activité)

                          C (commentaire)

                  Ajout d'un champ pStAFFREF pour garder l'affaire principale ou de référence

                  en tête lors des tris

                  Ajout champ sup , DATACT  date activité pour tri

Mots clefs ... :
*****************************************************************}
procedure CreationLignePiece(pTobPiece, pTobLigne, pTobAffaire, pTobArticles, pTobTiers: TOB; pRdQte, pRdPuV: double; pStLib, pStOri, pStAffRef: string;
  pCleDoc: R_CleDoc);
var
  vTobA: Tob;
  vTobArt: Tob;
  refunique: string;
  vStAff: string;
  vStDat: string;
  ind: integer;
  QQ: Tquery;
  vRdCalc: double;
  vDEV: RDEVISE;

begin

  vDEV.Code := pTOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(vDEV);
  //mcd 11/04/2003 pb si tarification et forfait.. le taux est à 0 et calcul faux
  vDEV.Taux := GetTaux(vDEV.Code, vDEV.DateTaux, pCleDoc.DatePiece);

  ind := pTOBPiece.Detail.Count;
  RefUnique := pTobligne.GetValue('GL_ARTICLE');
  if (RefUnique <> '') then
  begin
    vTOBArt := FindTOBArtSais(pTOBArticles, RefUnique);
    if vTOBArt = nil then
    begin
      //***Création des pieces, il faut tout prendre ****
      QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                     'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                     'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefUnique+'"',true);
      if not QQ.EOF then
      begin
        vTOBArt := CreerTOBArt(pTOBArticles);
        vTOBArt.SelectDB('', QQ);
        InitChampsSupArticle (vTOBArt);
        vTOBArt.addchampsup('GCA_LIBELLE', false);
      end;
      Ferme(QQ);
    end;

    vTobA := FindTOBArtRow(pTOBPiece, pTOBArticles, ind);
    vTobA.putvalue('GCA_LIBELLE', pStlib);
    vTobA.putvalue('GA_PVHT', pRdPuV); // pour reprendre notre prix et pas celui de la fiche article
    vTobA.putvalue('GA_PCB', 0); // sinon on reprend systématiquement GA_PCB et non notre Quantité
  end
  else
    vTobA := nil;

  InitLigneVide(pTobPiece, pTobLigne, pTobTiers, pTobAffaire, ind, pRdQte);
  pTobLigne.putvalue('GL_REFARTSAISIE', ''); // mcd 26/02/2003 pour OK sur fct ArticleVrs lignes
  ArticleVersLigne(pTOBPiece, vTobA, nil {TOBConds}, pTOBLigne, pTobTiers);
  pTobLigne.PutValue('GL_REFARTSAISIE', pTobLigne.getvalue('GL_CODEARTICLE')); // mcd 26/02/03
  if (pStOri = 'A') then
  begin
    pTobLigne.PutValue('GL_PUHTDEV', pRdPuV);
    pTobLigne.PutValue('GL_PUHT', DeviseToPivotEx(pRdPuV, vDEV.Taux, vDEV.Quotite,V_PGI.okdecP));
  end
  else
  begin
    pTOBLigne.PutValue('GL_PUHT', pRdPuV);
    pTobLigne.PutValue('GL_PUHTDEV', PivotToDevise(pRdPuV, vDEV.Taux, vDEV.Quotite, V_PGI.OkdecP));
  end;

  vRdCalc := Arrondi(pTobligne.GetValue('GL_QTEFACT') * pTobligne.GetValue('GL_PUHTDEV'), V_PGI.OkDecV);
  pTobligne.PutValue('GL_TOTALHTDEV', vRdCalc);

  pTobLigne.PutValue('GL_AFFAIRE', pTOBAffaire.GetValue('AFF_AFFAIRE'));
  pTobLigne.PutValue('GL_AFFAIRE1', pTOBAffaire.GetValue('AFF_AFFAIRE1'));
  pTobLigne.PutValue('GL_AFFAIRE2', pTOBAffaire.GetValue('AFF_AFFAIRE2'));
  pTobLigne.PutValue('GL_AFFAIRE3', pTOBAffaire.GetValue('AFF_AFFAIRE3'));
  pTobLigne.PutValue('GL_AVENANT', pTOBAffaire.GetValue('AFF_AVENANT'));
  pTobLigne.PutValue('GL_REPRESENTANT', pTOBAffaire.GetValue('LEREP'));

  vStAff := pTobligne.getValue('GL_AFFAIRE');
  vStDat := DatetoStr(idate2099);

  // gm    18/09/2001  , grosse bidouille car certains champ supp ne sont pas créés systématiquement
  //  donc mes champs ci dessous ne sont pas créé au même endroit, et le tri plante
  if not pTOBLigne.FieldExists('SUPPRIME') then
    pTobLigne.AddChampSup('SUPPRIME', False);

  pTobLigne.PutValue('SUPPRIME', '-');
  pTobLigne.AddChampSup('GLLAFFREF', false);
  pTobLigne.AddChampSup('GLL_DATEACT', false);
  pTobLigne.AddChampSup('GLLTIERSORI', false); // tiers d'origine

  if (vStAff = pStAffRef) then
    pTobLigne.PutValue('GLLAFFREF', '0')
  else
    pTobLigne.PutValue('GLLAFFREF', '1');

  pTobLigne.PutValue('GLL_DATEACT', vStDat);
  pTobLigne.PutValue('GLLTIERSORI', pTobAffaire.GetValue('AFF_TIERS'));

end;

{***********A.G.L.***********************************************
Auteur  ...... : GM - CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : déplacé de AFPrepFact
                 réinitialisé car sinon la numérotation repart du
                 max du numordre de la piece d'origine
Mots clefs ... :
*****************************************************************}
procedure RenumLesLignes(pTobPiece: TOB);
var
  wj: Integer;
begin
  pTobPiece.putvalue('GP_CODEORDRE', 0);
  for wj := 0 to pTOBPiece.Detail.Count - 1 do
    pTOBPiece.Detail[wj].putvalue('GL_NUMORDRE', 0);
  NumeroteLignesGC(nil, pTobPiece);
end;

{***********A.G.L.***********************************************
Auteur  ...... : GM - CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : déplacé de AFPrepFact
Mots clefs ... :
*****************************************************************}
procedure MajBasePort(pTobPiece, pTobPorts: Tob);
var
  wi, wz: Integer;
  base: Double;
  vtobDet: TOB;
  vtobdet1: TOB;

begin
  Base := 0;
  for wi := 0 to pTobPorts.detail.count - 1 do
  begin
    vTobDet := pTobports.detail[wi];
    if (vTobDet.GetValue('GPT_TYPEPORT') <> 'MT') then
    begin
      for wz := 0 to pTobPiece.detail.count - 1 do
      begin
        vTobDet1 := pTobPiece.detail[wz];
        if (vTobDet1.getvalue('GL_TYPELIGNE') = 'ART') and (vTobDet1.getvalue('GL_TYPEARTICLE') <> 'POU') then
          base := base + vTobDet1.getvalue('GL_TOTALHTDEV');

        // prix par 100 par encore appliqué
        if (vTobDet1.getvalue('GL_TYPELIGNE') = 'ART') and (vTobDet1.getvalue('GL_TYPEARTICLE') = 'POU') then
          base := base + (vTobDet1.getvalue('GL_TOTALHTDEV') / 100);
      end;
      vTOBdet.PutValue('GPT_BASEHTDEV', base);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GM
Créé le ...... :
Modifié le ... :   /  /
Description .. : Gestion des champs supplémantaire de la Tob Lignes (pour Tri)
Mots clefs ... :
*****************************************************************}
Procedure GestionChampSuppLigne(Toblignes : TOB;zaff,zdat,zcli,AffRef : string);
BEGIN
 
  // gm    18/09/2001  , grosse bidouille car certains champ supp ne sont pas créés systématiquement
  //  donc mes champs ci dessous ne sont pas créé au même endroit, et le tri plante
  if Not TOBLignes.FieldExists('SUPPRIME') then TOBLignes.AddChampSup('SUPPRIME',False) ;
  TOBLignes.PutValue('SUPPRIME','-') ;

	TobLignes.AddChampSup('GLLAFFREF',false);
  TobLignes.AddChampSup('GLL_DATEACT',false);
  TobLignes.AddChampSup('GLLTIERSORI',false);      // tiers d'origine

  // ONYX pour appel fonction Révision Prix
  TobLignes.AddChampSupValeur('GLLAFFAIRE1','',false);      // clé pour pointer sur tob Revision
  TobLignes.AddChampSupValeur('GLLFORCODE1','',false);
  TobLignes.AddChampSupValeur('GLLNUMEROLIGNE1','',false);
  TobLignes.AddChampSupValeur('GLLAFFAIRE2','',false);      // clé pour pointer sur tob Revision 2
  TobLignes.AddChampSupValeur('GLLFORCODE2','',false);
  TobLignes.AddChampSupValeur('GLLNUMEROLIGNE2','',false);

  // ONYX pour appel fonction Révision Prix
  TobLignes.AddChampSupValeur('GLLLIENVAR',0,false);      // clé pour pointer sur tob Variables

  if (zaff = AffRef) then
    TobLignes.PutValue('GLLAFFREF','0')
  else
    TobLignes.PutValue('GLLAFFREF','1');

  TobLignes.PutValue('GLL_DATEACT',zdat);
  TobLignes.PutValue('GLLTIERSORI',zcli);
END;

{***********A.G.L.***********************************************
Auteur  ...... : GM - CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : déplacé de AFPrepFact
Mots clefs ... :
*****************************************************************}

procedure RecalculeSousTotaux(pTobPiece: Tob);
var
  i: integer;
  vTOBL: TOB;

begin
  for i := 0 to pTOBPiece.Detail.Count - 1 do
  begin
    vTOBL := pTOBPiece.Detail[i];
    if vTOBL.GetValue('GL_TYPELIGNE') = 'TOT' then
      SommerLignes(pTOBPiece, i + 1, 0);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GM - CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : déplacé de AFPrepFact
Mots clefs ... :
*****************************************************************}

procedure ValideLaPiece(pTOBPiece, pTOBEches, pTobTiers, pTobPorts, pTobBases,
                        pTOBBasesL,pTOBNomenclature, pTobAdresses: Tob;
                        pInNewNum: Integer; pCleDoc: R_CleDoc;pDev: RDEVISE);
var
  vStNatPieceG: string;
  i, zz: integer;
  vTob: Tob;
  vTobb: Tob;

begin

  GestionDuVisa(pTobPiece);
  pTOBPiece.SetAllModifie(True);
  pTOBEches.SetAllModifie(True);
  pTOBBases.SetAllModifie(True);

  vStNatPieceG := pTOBPiece.GetValue('GP_NATUREPIECEG');

  if not SetDefinitiveNumber(pTOBPiece, pTOBBases, pTOBBasesL,pTOBEches, pTOBNomenclature, nil, nil, nil, nil, pInNewNum) then
    V_PGI.IoError := oePointage;

  if GetInfoParPiece(vStNatPieceG, 'GPP_ACTIONFINI') = 'ENR' then
    pTOBPiece.PutValue('GP_VIVANTE', '-');

  pCleDoc.NumeroPiece := pTOBPiece.GetValue('GP_NUMERO');

  for i := 0 to pTOBBases.Detail.Count - 1 do
  begin
    vTOBB := pTOBBases.Detail[i];
    MajFromCleDoc(vTOBB, pCleDoc);
  end;

  for i := 0 to pTOBEches.Detail.Count - 1 do
  begin
    vTOBB := pTOBEches.Detail[i];
    MajFromCleDoc(vTOBB, pCleDoc);
  end;

  // remettre le montant à 1
  GereEcheancesGC(pTOBPiece, pTobTiers, pTOBEches, nil, nil, nil,nil,taCreat, pDEV, False);

  // MAJ des prix valo et des stocks , trt nomenclature
  // ValideLesLignes(TobPiece_OK, TOBArticles,TOBNomenclature,False) ;

  ValideLesAdresses(pTobPiece, pTobPiece, pTobAdresses);
  //gerelesreliquats
  //ValideLesArticles   //maj TobArticle + Catalogue
  //ValideLaCompta

  // maj de creerpar en attendant la modif de JLD dans la fonction PieceVersLigne
  for zz := 0 to pTobpiece.Detail.count - 1 do
  begin
    vTob := pTobpiece.Detail[zz];
    vTob.putvalue('GL_CREERPAR', 'GEN');
  end;

  pTOBPiece.InsertDBByNivel(False);
  pTOBBases.InsertDB(nil);
  pTOBEches.InsertDB(nil);
  ValideLesPorcs(pTOBPiece, pTobPorts);
                            
  // TOBAnaP.InsertDB(Nil) ;
  // ValidelesNomen ;}
end;

procedure GestionDuVisa(pTobPiece: Tob);
var
  vBoNeedvisa: boolean;
  vRdMntVisa: double;
  vStNat: string;

begin
  vStNat := pTobPiece.getvalue('GP_NATUREPIECEG');
  vBoNeedVisa := GetInfoParPiece(vStNat, 'GPP_VISA') = 'X';
  vRdMntVisa := GetInfoParPiece(vStNat, 'GPP_MONTANTVISA');

  if (vBoNeedVisa) and (Abs(pTOBPiece.GetValue('GP_TOTALHT')) >= vRdMntVisa) then
    pTOBPiece.PutValue('GP_ETATVISA', 'ATT')
  else
    pTOBPiece.PutValue('GP_ETATVISA', 'NON');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 26/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function RechNumeroPiece(pStNat: string): Integer;
var
  vInNewNum: Integer;
  vStSouche: string;

  function TestSetNumberAttribution(Souche: string): integer;
  var
    SoucheG: String3;
    QQ: TQuery;
    NumDef: Longint;
    Nb, ii: integer;
    Okok: boolean;

  begin
    Result := 0;
    SoucheG := Souche;
    //okok := True;
    if SoucheG = '' then Exit;
    ii := 0;

    repeat
      NumDef := -1;
      inc(ii);
      QQ := OpenSQL('SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '"', True);
      if not QQ.EOF then NumDef := QQ.Fields[0].AsInteger;
      Ferme(QQ);
      if NumDef <= 0 then exit;
      Nb := ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART=' + IntToStr(NumDef + 1) + ' WHERE SH_TYPE="GES" AND SH_SOUCHE="' + SoucheG + '" AND SH_NUMDEPART=' +
        IntToStr(NumDef));
      Okok := (Nb = 1);
    until ((Okok) or (ii > 20));

    if not Okok then
    begin
      V_PGI.IoError := oeUnknown;
      Exit;
    end;
    Result := NumDef;
  end;

begin

  vStSouche := GetSoucheG(pStNat, '', ''); // attention, pour gérer une souche par établiessement
  // (demande mcd), il faudrait accéder à l'affaire pour

                // connaitre l'etablisseùment et le domaine.

                // attention, si on veut générer un avoir, il faudrait le savoir

                // ici ????

  vInNewNum := TestSetNumberAttribution(vStSouche);
  if vInNewNum > 0 then
    result := vInNewNum
  else
    result := -1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 27/06/2003
Modifié le ... : 27/06/2003
Description .. : contrôle du mode de cumul.
Suite ........ : Si on a un cumul sur code artilce, on perd la notion de
Suite ........ : variable
Mots clefs ... : GIGA;FACTURATION;VARIABLE
*****************************************************************}

function TestSiCumulArticle(pTobLig: TOB; pBoCumul: Boolean;
  pInPres, pInFour, pInFrais, pInCtr: Integer): Boolean;
begin
  result := true;
  if pBoCumul then // profil ressource par prestation
    result := false
  else
    if (pTobLig.GetValue('GL_TYPEARTICLE') = 'PRE') and (pInPres = 1) then
    result := false
  else
    if (pTobLig.GetValue('GL_TYPEARTICLE') = 'MAR') and (pInFour = 1) then
    result := false
  else
    if (pTobLig.GetValue('GL_TYPEARTICLE') = 'FRA') and (pInFrais = 1) then
    result := false
  else
    if (pTobLig.GetValue('GL_TYPEARTICLE') = 'CTR') and (pInCtr = 1) then
    result := false;

  if not (result) then pTobLig.PutValue('GL_FORMULEVAR', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 30/06/2003
Modifié le ... :   /  /
Description .. : dans le cas des factures de regul, on peut avoir une formule
                 et une revision, mais que la ligne ne soit ps crée mais
                 reprise de l'ancienne facture. les zones GLL_AFFAIRE... ne sont
                 pas remplies dans ce cas

                 pBoOrigine : on remplie les données sur la ligne d'origine avant
                              avant qu'on perde les données de la facture d'origine
Mots clefs ... :
*****************************************************************}
procedure MajTobRev(pTobRevision, pTobLig : TOB; pInNum: Integer; pBoOrigine : Boolean);
var
  vTobRev: TOB;
  vInNumLig: Integer;
  vStAff: string;
  vStForcode: string;

begin
  if not pTOBLIG.FieldExists ('GLLAFFAIRE'+ IntToStr(pInNum))  then exit;
  vStAff := pTobLig.GetValue('GLLAFFAIRE' + IntToStr(pInNum));
  if (vStAff <> '') then //and (vStAff <> #0) then
  begin
    vStForcode := pTobLig.GetValue('GLLFORCODE' + IntToStr(pInNum));
    vInNumLig := pTobLig.GetValue('GLLNUMEROLIGNE' + IntToStr(pInNum));
    vTobRev := pTobRevision.FindFirst(['AFR_AFFAIRE', 'AFR_FORCODE', 'AFR_NUMEROLIGNE'], [vStAff, vStForcode, vInNumLig], True);
    if (vTobRev <> nil) then
    begin
      if pBoOrigine then
      begin
        if pTobLig.getvalue('GL_NATUREPIECEG') = 'FPR' then
          vTobRev.PutValue('AFR_ONATUREPIECEG', 'FAC')
        else
          vTobRev.PutValue('AFR_ONATUREPIECEG', 'FPR');
        vTobRev.PutValue('AFR_OSOUCHE', pTobLig.getvalue('GL_SOUCHE'));
        vTobRev.PutValue('AFR_ONUMERO', pTobLig.getvalue('GL_NUMERO'));
        vTobRev.PutValue('AFR_OINDICEG', 0);
        vTobRev.PutValue('AFR_ONUMORDRE', pTobLig.getvalue('GL_NUMORDRE'));
      end
      else                                           
      begin
        vTobRev.PutValue('AFR_NATUREPIECEG', pTobLig.getvalue('GL_NATUREPIECEG'));
        vTobRev.PutValue('AFR_SOUCHE', pTobLig.getvalue('GL_SOUCHE'));
        vTobRev.PutValue('AFR_NUMERO', pTobLig.getvalue('GL_NUMERO'));
        vTobRev.PutValue('AFR_INDICEG', 0);
        vTobRev.PutValue('AFR_NUMORDRE', pTobLig.getvalue('GL_NUMORDRE'));
      end;
    end;
  end;
end;

procedure G_ChargeLesVariables(TOBPiece, TOBVariable: TOB); //Affaire-ONYX
var Q: TQuery;
  CD: R_CleDoc;
begin
  CD := TOB2CleDoc(TOBPiece);
  Q := OpenSQL('SELECT * FROM AFORMULEVARQTE WHERE ' + WherePiece(CD, ttdVariable, False), True);
  if not Q.EOF then TOBVariable.LoadDetailDB('AFORMULEVARQTE', '', '', Q, True, True);
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : merieux
Créé le ...... : 16/05/2003
Modifié le ... :   /  /
Description .. : Alimentation dans une tob des éléments variables liés aux
Suite ........ : lignes d'affaire
Mots clefs ... : GIGA;VARIABLES;ONYX
*****************************************************************}
procedure LoadTobVarQte(pTob, pTobAffaire : TOB; pStOri : String; CleDocOrigine : R_CleDoc);
Var
  vQr     : TQuery;
  vSt     : String;
  vStAff  : String;
  vStNat  : String;

begin

  vStAff := pTobAffaire.getvalue('AFF_AFFAIRE');
  vStNat := CleDocOrigine.naturepiece;

  if (pStOri = 'LIG') then
    vSt := 'select * from AFORMULEVARQTE where AVV_AFFAIRE="'+ vStAff + '"AND AVV_ORIGVAR="'+ vStNat +'"'
  else
    vSt := 'select * from AFORMULEVARQTE where AVV_AFFAIRE="'+ vStAff + '"AND AVV_ORIGVAR="'+ pStOri + '"';

  vQr := nil;
  Try
    vQr := OpenSQL(vSt, True);
    pTob.LoadDetailDB('AFORMULEVARQTE','','', vQr, False) ;  // ajout dans la tob
    if (pTob.detail.count <> 0) then
      pTob.detail[0].AddChampSupvaleur('AVVLIENVAR', 0, True);
  Finally
    Ferme(vQr);
  end;
end;

// alim de la tob des qte variable finale(_LIG) à partir de la tob d'origine (_AFF)+ lien avec la ligne
procedure AlimTobVarQte(TobVarQte_LIG, TobLig, TobAct, TobVarQte: TOB; stOri : string; var pInCpt : Integer);
Var
  TobVarDet     : TOB;

begin

  // détail de variable non stocké si cumul sur code article
  if not(TestSiCumulArticle(TobLig, False, 0, 0, 0, 0)) then exit;
  {
  if (stOri = 'LIG') then
  begin
    stAff     := Toblig.getvalue('GL_AFFAIRE');
    iNumOrdre := Toblig.getvalue('GL_NUMORDRE');
    Result := TobVarQte.FindFirst(['AVV_AFFAIRE','AVV_NUMORDRE'],[StAff,iNumOrdre],true);
  end
  else
  begin
    stAff         := TobAct.getvalue('ACT_AFFAIRE');
    iNumligUnique := TobAct.getvalue('ACT_NUMLIGNEUNIQUE');
    TobVarDet     := TobVarQte.FindFirst(['AVV_AFFAIRE','AVV_NUMLIGNEUNIQUE'],[StAff,iNumligUnique],true);
  end;
  }
  TobVarDet := RechVariable(TobVarQte,TobLig,TobAct,stOri);

  FinalisationTobVar(TobVarQte_LIG, TobVarDet,Toblig, pInCpt);

end;

function RechVariable(TobVarQte,TobLig,TobAct:TOB ; stOri : String):TOB;
Var
  stAff         : String;
  iNumOrdre     : Integer;
  iNumLigUnique : Integer;
begin
  if (stOri = 'LIG') then
  begin
    stAff     := Toblig.getvalue('GL_AFFAIRE');
    iNumOrdre := Toblig.getvalue('GL_NUMORDRE');
    Result    := TobVarQte.FindFirst(['AVV_AFFAIRE','AVV_NUMORDRE'],[StAff,iNumOrdre],true);
  end
  else
  begin
    stAff         := TobAct.getvalue('ACT_AFFAIRE');
    iNumligUnique := TobAct.getvalue('ACT_NUMLIGNEUNIQUE');
    Result        := TobVarQte.FindFirst(['AVV_AFFAIRE','AVV_NUMLIGNEUNIQUE'],[StAff,iNumligUnique],true);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 24/06/2003
Modifié le ... :   /  /    
Description .. : Gestion d'un compteur interne entre tob
Mots clefs ... : GIGA;FACTURATION;VARIABLE
*****************************************************************}
procedure FinalisationTobVar(TobVarQte_LIG, TobVarDet, Toblig : TOB; var pInCpt : Integer);
begin
  if (TobVarDet <> Nil) then
  begin
    pInCpt := pInCpt + 1;
    // alim du compteur afin de faire le lien entre la TobLigne et la TobQteVar
    Toblig.PutValue('GLLLIENVAR', pInCpt);
    TobVarDet.PutValue('AVVLIENVAR', pInCpt);
    if (toblig.GetValue('GL_FORMULEVAR')='') then
      TobLig.PutValue('GL_FORMULEVAR',TobVarDet.getvalue('AVV_FORMULEVAR'));

    TobVarDet.ChangeParent(TobVarQte_LIG,-1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GM - CB
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : GIGA;FACTURE;VARIABLES
*****************************************************************}
Procedure MajVariable(pTobNewPiece, pTobVarQte_LIG : Tob);
Var
  TobLig  : Tob;
  TobVar  : TOB;
  iInd    : Integer;
  wj      : Integer;

Begin
       
  for iInd :=0 to pTobNewPiece.Detail.count-1 do
  Begin
    TobLig := pTobNewPiece.Detail[iInd];
    if (toblig.GetValue('GL_TYPELIGNE') = 'ART') and (toblig.GetValue('GL_TYPEARTICLE') <> 'POU') then
    begin
      if (TestSiCumulArticle(TobLig, False, 0, 0, 0, 0)) then
      begin
        TobVar := pTobVarQte_LIG.FindFirst(['AVVLIENVAR'],[TobLig.GetValue('GLLLIENVAR')],true);
        if (TobVar <> NIL) then
        begin
          TobVar.PutValue('AVV_NATUREPIECEG',TobLig.getvalue('GL_NATUREPIECEG'));
          TobVar.PutValue('AVV_ORIGVAR',TobLig.getvalue('GL_NATUREPIECEG'));
          TobVar.PutValue('AVV_SOUCHE',TobLig.getvalue('GL_SOUCHE'));
          TobVar.PutValue('AVV_NUMERO',TobLig.getvalue('GL_NUMERO'));
          TobVar.PutValue('AVV_NUMORDRE',TobLig.getvalue('GL_NUMORDRE'));
        end;
      end;
    end;
  end;

  wj := 0; Iind := 0;
  while  (wj < pTobVarQte_LIG.Detail.count) do
  begin
    TobVar := pTobVarQte_LIG.Detail[wj];
    if tobvar.getValue('AVV_NUMERO') = 0 then
      TobVar.free
    else
      Inc(wj);
    TobVar.PutValue('AVV_RANGVAR',iInd+1);
    Inc(Iind);
  end;
      
  if (pTobVarQte_LIG.detail.count <> 0) then
    pTobVarQte_LIG.InsertDB(Nil);
End;


end.
