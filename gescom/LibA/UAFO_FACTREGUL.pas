unit UAFO_FACTREGUL;

interface

  uses UTob, UTOF, hctrls, sysutils, entGC, HEnt1,

  {$IfDEF EAGLCLIENT}
    eFiche, MaineAGL,
  {$ELSE}
    Fiche,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$EndIf}

  Dicobtp, UAFO_REVPRIXCALCULCOEF, AFPrepFact,
  FactAffaire, AffaireUtil, FactCalc,
  FactUtil, FactComm, SAISUTIL, uTofAfPiece,
  UAFO_REVPRIXLECTURECOEF, paramsoc, {voirtob,}
  utilRevision, utilPrepFact,FactTob,FactAdresse,
  FactTiers,uEntCommun,UtilTOBPiece;
                                    
type                  
  TFACTREGUL = class

    private
      fTobPiece       : Tob;
      fTobLigne       : Tob;
      fTobLigneRegul  : Tob;
      fTobCoef        : Tob;
      fTobLien        : Tob;
      fTobRevision    : Tob;
      fTobArticles    : Tob;
      fTOBVarQte_LIG  : Tob;

      fStComment      : String;
       
      fInNbFac    : Integer;
      fInFacDeb   : Integer;
      fInFacFin   : integer;

      function LoadFactures(pStWhere : String) : R_CleDoc;
      procedure LoadCoef;
      function CreatePiece(pTobMaPiece : Tob; pDtRegul : TDateTime) : Tob;
      procedure CalculCoefRegul;
      function CreateFactures(pDtRegul : TDateTime) : Boolean;
      procedure FreeFactures;

      function ValiderFacture(pTobNewPiece, pTobVarQte_AFF, pTobAffaire, pTobArticles, pTobAdresses, pTobTiers : Tob;
                              pCleDoc : R_CleDoc; pDev : RDEVISE; pDtRegul : TDateTime;
                              var pInNumRev : Integer) : Boolean;

      procedure AddLigneFacture(pTobNewPiece, pTobLigne, pTobVarQte_AFF, pTobAffaire,
                                pTobTiers : Tob; pCleDoc : R_CleDoc;
                                var pInNumRev, pInCptVariables : Integer;
                                pDtRegul : TDateTime);

      procedure CreateLigneRegul(pTobNewPiece, pTobAffaire, pTobArticles, pTobTiers,
                                 pTobLigne : Tob; pCleDoc : R_CleDoc;
                                 var pInNumRev : Integer; pDtRegul : TDateTime);

      //procedure AddTobLien(pTobOrig, pTobDest : Tob);
      procedure LoadTob(pCleDocOrigine : R_CleDoc; pTobBases,pTOBBasesL, pTobPorts : Tob);
      procedure LoadTOBAffaire(var pTobAffaire : Tob; pStCodeAffaire, pStRep : String);
      procedure LoadTobAdresses(var pTobAdresses: Tob; pTobPiece, pTobTiers, pTobAffaire : Tob);
      Procedure ReplaceSubStr(Var pSt : String; pStSup, pStRep : string);
      function GestionRevisionRegul(pTobNewPiece : Tob; pDtEcheance : TDateTime; var pInNumRev : Integer) : boolean;
      Procedure MajRevision(pTobNewPiece : Tob);
      procedure CreationLigneRevision(pTobNewPiece, pTobNewLigne, pTobLigne : Tob;
                                      var pInNumRev : integer);
      procedure LigneComAff(TOBAffaire, TOBPiece : Tob; Comment : string);

    public
      function RegulFactures(pStWhere : String; pDtRegul : TDateTime; pStComment : String) : Boolean;
      
  end;
const
	TexteMsg: array[1..3] of string 	= (
     {1}        'Aucune facture de régularisation généré.',
     {2}        '%d Facture(s) de régularisation générée(s) : n° %d à %d  de %s à %s.',
     {3}        'Erreur lors de la génération des factures de régularisation.');
 
implementation

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/06/2003
Modifié le ... :   /  /
Description .. : Regularisation des pièces sélectionnées
                 On fait une facture de régul par client
                 Tob triée par client et par affaire
Mots clefs ... :
*****************************************************************}
function TFACTREGUL.RegulFactures(pStWhere : String; pDtRegul : TDateTime; pStComment : String) : Boolean;
var
  vCleDoc     : R_CleDoc;
  vStTexte    : String;
  vStDateDeb  : String;
  vStDateFin  : String;
  vStParam    : String;

begin
 
  vStDateDeb := FormatDateTime('dd/mm/yyyy ttttt', now);

  result := True;

  fTobLien      := Tob.create('tob lien', nil, -1);
  fTobArticles  := Tob.create('ARTICLE', nil, -1);
  fTobCoef      := Tob.create('AFREVISION', nil, -1);
  fTobRevision  := TOB.Create('les revisions',Nil,-1) ;

  fStComment := pStComment;

  if GetParamSoc('SO_AFVARIABLES') then
    fTOBVarQte_LIG := TOB.Create('Les Qte variable de la piece finale',Nil,-1) ;

  try
    fInNbFac := 0;
    vCleDoc := LoadFactures(pStWhere);
    BeginTrans;
    Try
      CalculCoefRegul;
      CommitTrans;
    Except
      result := False;
      Rollback;
    End;

    if result then
    begin
      LoadCoef;

      BeginTrans;
      Try
        if not CreateFactures(pDtRegul) then
        begin
          PGIInfoAF(texteMsg[1], '');
          result := False;
        end

        else if (fInFacDeb <> 0) then
        begin
          vStDateFin := FormatDateTime('dd/mm/yyyy ttttt', now);
          vStTexte := Format(TexteMsg[2], [fInNbFac, fInFacDeb, fInFacFin, vStDateDeb, vStDateFin]);
          PGIInfoAf(vStTexte,'');
        end;
        CommitTrans;
      Except
        Rollback;
        PGIBoxAF(TexteMsg[3],'');
      End;
    end;

    // enchainer sur la validation de facture
    if result then
    Begin
      vStparam := 'TABLE:GP;NATURE:FPR;NOCHANGE_NATUREPIECE'+';NUMDEB:'+IntToStr(fInFacDeb)+';NUMFIN:'+IntToStr(fInFacFin)+';';
      AFLanceFiche_Mul_Piece_Provisoire(vStParam);
    End
  finally
    FreeFactures;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : On selectionne toutes les lignes de factures regularisables
Mots clefs ... :
*****************************************************************}
function TFACTREGUL.LoadFactures(pStWhere : String) : R_CleDoc;
var
  vSt           : String;
  vQr           : TQuery;

begin

  fTobPiece := Tob.create('PIECE', nil, -1);
  fTobLigne := Tob.create('Mes lignes', nil, -1);
  fTobLigneRegul := Tob.create('Mes lignes regularisees', nil, -1);

  AddLesSupEntete(fTobPiece);

  vQr := nil;
  Try
    // seules des factures sélectionnées sont regularisées
    vSt := 'SELECT * FROM PIECE ';
    vSt := vSt + ' WHERE ' + pStWhere;
    vSt := vSt + ' AND GP_NATUREPIECEG = "FAC"';
    vSt := vSt + ' ORDER BY GP_AFFAIRE ';

    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
      fTobPiece.LoadDetailDB('PIECE', '', '', vQr, False) ;
  Finally
    if vQR <> nil then ferme(vQR);
  End;

  vQr := nil;
  vSt := 'SELECT LIGNE.*, ';
  vSt := vSt + 'A1.AFR_AFFAIRE AS A1_AFR_AFFAIRE, ';
  vSt := vSt + 'A1.AFR_FORCODE AS A1_AFR_FORCODE, ';
  vSt := vSt + 'A1.AFR_NUMERO AS A1_AFR_NUMERO,';
  vSt := vSt + 'A1.AFR_DATECALCCOEF AS A1_AFR_DATECALCCOEF,';
  vSt := vSt + 'A1.AFR_COEFREGUL AS A1_AFR_COEFREGUL,';
  vSt := vSt + 'A1.AFR_NATUREPIECEG AS A1_AFR_NATUREPIECEG,';
  vSt := vSt + 'A1.AFR_SOUCHE AS A1_AFR_SOUCHE,';
  vSt := vSt + 'A1.AFR_INDICEG AS A1_AFR_INDICEG,';
  vSt := vSt + 'A1.AFR_NUMORDRE AS A1_AFR_NUMORDRE,';
  vSt := vSt + 'A1.AFR_NUMFORMULE AS A1_AFR_NUMFORMULE,';
  vSt := vSt + 'A1.AFR_CREATEUR AS A1_AFR_CREATEUR,';
  vSt := vSt + 'A1.AFR_DATECREATION AS A1_AFR_DATECREATION,';
  vSt := vSt + 'A1.AFR_DATEMODIF AS A1_AFR_DATEMODIF,';
  vSt := vSt + 'A1.AFR_COEFCALC AS A1_AFR_COEFCALC,';
  vSt := vSt + 'A1.AFR_COEFAPPLIQUE AS A1_AFR_COEFAPPLIQUE,';
  vSt := vSt + 'A1.AFR_DERNCOEFCALC AS A1_AFR_DERNCOEFCALC,';
  vSt := vSt + 'A1.AFR_DERNIERCOEF AS A1_AFR_DERNIERCOEF,';
  vSt := vSt + 'A1.AFR_APPLIQUECOEF AS A1_AFR_APPLIQUECOEF,';
  vSt := vSt + 'A1.AFR_OKCOEFAPPLIQUE AS A1_AFR_OKCOEFAPPLIQUE,';
  vSt := vSt + 'A1.AFR_VALINDICE1 AS A1_AFR_VALINDICE1,';
  vSt := vSt + 'A1.AFR_VALINDICE2 AS A1_AFR_VALINDICE2,';
  vSt := vSt + 'A1.AFR_VALINDICE3 AS A1_AFR_VALINDICE3,';
  vSt := vSt + 'A1.AFR_VALINDICE4 AS A1_AFR_VALINDICE4,';
  vSt := vSt + 'A1.AFR_VALINDICE5 AS A1_AFR_VALINDICE5,';
  vSt := vSt + 'A1.AFR_VALINDICE6 AS A1_AFR_VALINDICE6,';
  vSt := vSt + 'A1.AFR_VALINDICE7 AS A1_AFR_VALINDICE7,';
  vSt := vSt + 'A1.AFR_VALINDICE8 AS A1_AFR_VALINDICE8,';
  vSt := vSt + 'A1.AFR_VALINDICE9 AS A1_AFR_VALINDICE9,';
  vSt := vSt + 'A1.AFR_VALINDICE10 AS A1_AFR_VALINDICE10,';
  vSt := vSt + 'A1.AFR_ONATUREPIECEG AS A1_AFR_ONATUREPIECEG,';
  vSt := vSt + 'A1.AFR_OSOUCHE AS A1_AFR_OSOUCHE,';
  vSt := vSt + 'A1.AFR_OINDICEG AS A1_AFR_OINDICEG,';
  vSt := vSt + 'A1.AFR_ONUMORDRE AS A1_AFR_ONUMORDRE,';
  vSt := vSt + 'A1.AFR_ONUMERO AS A1_AFR_ONUMERO,';

  vSt := vSt + 'A2.AFR_AFFAIRE AS A2_AFR_AFFAIRE, ';
  vSt := vSt + 'A2.AFR_FORCODE AS A2_AFR_FORCODE, ';
  vSt := vSt + 'A2.AFR_NUMERO AS A2_AFR_NUMERO,';
  vSt := vSt + 'A2.AFR_DATECALCCOEF AS A2_AFR_DATECALCCOEF,';
  vSt := vSt + 'A2.AFR_COEFREGUL AS A2_AFR_COEFREGUL,';
  vSt := vSt + 'A2.AFR_NATUREPIECEG AS A2_AFR_NATUREPIECEG,';
  vSt := vSt + 'A2.AFR_SOUCHE AS A2_AFR_SOUCHE,';
  vSt := vSt + 'A2.AFR_INDICEG AS A2_AFR_INDICEG,';
  vSt := vSt + 'A2.AFR_NUMORDRE AS A2_AFR_NUMORDRE,';
  vSt := vSt + 'A2.AFR_NUMFORMULE AS A2_AFR_NUMFORMULE,';
  vSt := vSt + 'A2.AFR_CREATEUR AS A2_AFR_CREATEUR,';
  vSt := vSt + 'A2.AFR_DATECREATION AS A2_AFR_DATECREATION,';
  vSt := vSt + 'A2.AFR_DATEMODIF AS A2_AFR_DATEMODIF,';
  vSt := vSt + 'A2.AFR_COEFCALC AS A2_AFR_COEFCALC,';
  vSt := vSt + 'A2.AFR_COEFAPPLIQUE AS A2_AFR_COEFAPPLIQUE,';
  vSt := vSt + 'A2.AFR_DERNCOEFCALC AS A2_AFR_DERNCOEFCALC,';
  vSt := vSt + 'A2.AFR_DERNIERCOEF AS A2_AFR_DERNIERCOEF,';
  vSt := vSt + 'A2.AFR_APPLIQUECOEF AS A2_AFR_APPLIQUECOEF,';
  vSt := vSt + 'A2.AFR_OKCOEFAPPLIQUE AS A2_AFR_OKCOEFAPPLIQUE,';
  vSt := vSt + 'A2.AFR_VALINDICE1 AS A2_AFR_VALINDICE1,';
  vSt := vSt + 'A2.AFR_VALINDICE2 AS A2_AFR_VALINDICE2,';
  vSt := vSt + 'A2.AFR_VALINDICE3 AS A2_AFR_VALINDICE3,';
  vSt := vSt + 'A2.AFR_VALINDICE4 AS A2_AFR_VALINDICE4,';
  vSt := vSt + 'A2.AFR_VALINDICE5 AS A2_AFR_VALINDICE5,';
  vSt := vSt + 'A2.AFR_VALINDICE6 AS A2_AFR_VALINDICE6,';
  vSt := vSt + 'A2.AFR_VALINDICE7 AS A2_AFR_VALINDICE7,';
  vSt := vSt + 'A2.AFR_VALINDICE8 AS A2_AFR_VALINDICE8,';
  vSt := vSt + 'A2.AFR_VALINDICE9 AS A2_AFR_VALINDICE9,';
  vSt := vSt + 'A2.AFR_VALINDICE10 AS A2_AFR_VALINDICE10,';
  vSt := vSt + 'A2.AFR_ONATUREPIECEG AS A2_AFR_ONATUREPIECEG,';
  vSt := vSt + 'A2.AFR_OSOUCHE AS A2_AFR_OSOUCHE,';
  vSt := vSt + 'A2.AFR_OINDICEG AS A2_AFR_OINDICEG,';
  vSt := vSt + 'A2.AFR_ONUMORDRE AS A2_AFR_ONUMORDRE,';
  vSt := vSt + 'A2.AFR_ONUMERO AS A2_AFR_ONUMERO ';

  vSt := vSt + 'FROM AFREVISION A1, LIGNE LEFT OUTER JOIN AFREVISION A2 ';
  vSt := vSt + ' ON (A2.AFR_AFFAIRE = GL_AFFAIRE';
  vSt := vSt + ' AND A2.AFR_FORCODE = GL_FORCODE2';
  vSt := vSt + ' AND A2.AFR_NUMERO = GL_NUMERO';
  vSt := vSt + ' AND A2.AFR_NUMORDRE = GL_NUMORDRE)';

  if pStWhere <> '' then
  begin
    vSt := vSt + ' WHERE ' + pStWhere;
    ReplaceSubStr(vSt, 'GP_', 'GL_');
    vSt := vSt + ' AND A1.AFR_AFFAIRE = GL_AFFAIRE';
  end
  else
    vSt := vSt + ' WHERE A1.AFR_AFFAIRE = GL_AFFAIRE';

  vSt := vSt + ' AND A1.AFR_FORCODE = GL_FORCODE1';
  vSt := vSt + ' AND A1.AFR_NUMERO = GL_NUMERO';
  vSt := vSt + ' AND A1.AFR_NUMORDRE = GL_NUMORDRE';
  vSt := vSt + ' AND A1.AFR_NATUREPIECEG = GL_NATUREPIECEG';
  vSt := vSt + ' AND A1.AFR_SOUCHE = GL_SOUCHE';
  vSt := vSt + ' AND A1.AFR_INDICEG = GL_INDICEG';
                                                   
  vSt := vSt + ' ORDER BY GL_TIERS, A1.AFR_ONUMERO, A1.AFR_ONUMORDRE ';

  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
      fTobLigne.LoadDetailDB('Mes lignes', '', '', vQr, False) ;

  Finally
    if vQR <> nil then ferme(vQR);
  End;

//  vSt := 'SELECT GL_NATUREPIECEG, GL_SOUCHE, GL_NUMERO, ';
//  vSt := vSt + ' GL_INDICEG, GL_NUMORDRE, GL_REGULARISE ';

  // obligé de faire un select * pour avoir le droit de faire un updateDb sur cette table
  vSt := 'SELECT * FROM LIGNE';
  if pStWhere <> '' then
  begin
    vSt := vSt + ' WHERE ' + pStWhere;
    ReplaceSubStr(vSt, 'GP_', 'GL_');
  end;

  vQr := nil;
  Try
    vQR := OpenSql(vSt, True,-1,'',true);
    if Not vQR.Eof then
      fTobLigneRegul.LoadDetailDB('LIGNE', '', '', vQr, False);
  Finally
    if vQR <> nil then ferme(vQR);
  End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/06/2003
Modifié le ... :
Description .. : création des factures de regul
Mots clefs ... :
*****************************************************************}
function TFACTREGUL.CreateFactures(pDtRegul : TDateTime) : Boolean;
var
  i               : Integer;
  vInAFFAIRE      : Integer;
  vInTIERS        : Integer;

  vInNUMERO       : Integer;
  vInNUMORDRE     : Integer;
  vInNATUREPIECEG : Integer;
  vInSOUCHE       : Integer;
  vInINDICEG      : Integer;

  vStAFFAIRE      : String;
  vStTIERS        : String;
  vStNUMERO       : String;
  vStRep          : String;

  vTobUneLigne    : Tob;
  vTobPieceOrig   : Tob;
  vTobNewPiece    : Tob;

  vCleDocOrigine  : R_CleDoc;
//  vInNbPiece      : integer;
  vTobAffaire     : Tob;
  vTobAdresses    : Tob;
  vTOBTiers       : Tob;
  vDev            : RDEVISE;
  vInNumRev       : Integer;
  vTOBVarQte_LIG  : Tob;
  vTobVarQte_AFF  : Tob;
  vInCptVariables : Integer;

begin

  vInCptVariables := 0;
  vStRep := '';
  result := True;
  vTobNewPiece := nil;
  vTobTiers    := nil;

  vTOBVarQte_LIG := TOB.Create('Les Qte variable de la piece finale',Nil,-1) ;
  vTobVarQte_AFF := TOB.Create('Les qtés variables de affaire',Nil,-1) ;

  // on regroupe sur la nouvelle facture, par ligne d'affaire d'origine
  // et par client
  if fTobLigne.detail.count > 0 then
  begin
    i := 0;
    vStAFFAIRE := '';

    vInAFFAIRE      := fTobLigne.detail[0].GetNumChamp('GL_AFFAIRE');
    vInTIERS        := fTobLigne.detail[0].GetNumChamp('GL_TIERS');
    vInNATUREPIECEG := fTobLigne.detail[0].GetNumChamp('GL_NATUREPIECEG');
    vInNUMERO       := fTobLigne.detail[0].GetNumChamp('GL_NUMERO');
    vInSOUCHE       := fTobLigne.detail[0].GetNumChamp('GL_SOUCHE');
    vInINDICEG      := fTobLigne.detail[0].GetNumChamp('GL_INDICEG');
    vInNUMORDRE     := fTobLigne.detail[0].GetNumChamp('GL_NUMORDRE');

    // pour toutes les lignes d'affaires, on crée la ligne en négatif
    // puis on créé la ligne de regul
    while i < fTobLigne.detail.count do
    begin
      // si on n'est pas sur le même client, on change de facture
      // sinon, on passe a la ligne d'affaire suivante
      if (i = 0) or (fTobLigne.detail[i-1].GetValeur(vInTIERS) <> fTobLigne.detail[i].GetValeur(vInTIERS)) then
      begin

        vInNumRev := MaxRevision(fTobLigne.detail[i].GetValeur(vInAFFAIRE));

        // valide la facture precedente
        if (i <> 0) then ValiderFacture(vTobNewPiece, vTobVarQte_AFF, vTobAffaire, fTobArticles, vTobAdresses, vTobTiers, vCleDocOrigine, vDev, pDtRegul, vInNumRev);

        if assigned(vTobTiers) then vTobTiers.Free;
        if assigned(vTOBAdresses) then vTOBAdresses.Free;

        vTobTiers := TOB.Create('TIERS',Nil,-1);
        vTOBAdresses:=TOB.Create('LesADRESSES',Nil,-1) ;
        if GetParamSoc('SO_GCPIECEADRESSE') then
        Begin
          TOB.Create('PIECEADRESSE', vTOBAdresses,-1) ; {Livraison}
          TOB.Create('PIECEADRESSE', vTOBAdresses,-1) ; {Facturation}
        End
        else
        begin
          TOB.Create('ADRESSES', vTOBAdresses,-1) ; {Livraison}
          TOB.Create('ADRESSES', vTOBAdresses,-1) ; {Facturation}
        end;

        // creer la nouvelle facture
        vStAFFAIRE  := fTobLigne.detail[i].GetValeur(vInAFFAIRE);
        vStNUMERO   := fTobLigne.detail[i].GetValeur(vInNUMERO);
        vStTIERS    := fTobLigne.detail[i].GetValeur(vInTIERS);

        vTobPieceOrig := fTobPiece.FindFirst(['GP_AFFAIRE', 'GP_NATUREPIECEG', 'GP_NUMERO'],
                                             [vStAFFAIRE, 'FAC', vStNUMERO], true);
                                                       
        vTobNewPiece := CreatePiece(vTobPieceOrig, pDtRegul);

        SelectPieceAffaireBis(vStAFFAIRE, 'AFF', vCleDocOrigine, vStTIERS, vStRep);

        //Devise
        vDev.Code := vTobNewPiece.GetValue('GP_DEVISE');
        GetInfosDevise(vDev);
        vDev.Taux := GetTaux(vDev.Code, vDev.DateTaux, vCleDocOrigine.DatePiece);

        //Tiers
        if assigned(vTobTiers) then vTobTiers.Free;
        vTOBTiers := TOB.Create('TIERS',Nil,-1);
        vTOBTiers.AddChampSup('RIB',False);
        //RemplirTOBTiersBis(vTobTiers, vStTIERS, vCleDocOrigine);
        RemplirTOBTiers(vTobTiers, vStTIERS, vCleDocOrigine.NaturePiece, False);
                        
        LoadTOBAffaire(vTobAffaire, vStAFFAIRE, vStRep);
        LoadTobAdresses(vTobAdresses, vTobNewPiece, vTobTiers, vTobAffaire);
 
        // commentaire de facture de régul
        if fStComment <> '' then
          LigneComAff(vTOBAffaire, vTobNewPiece, fStComment);
      end;
 
      CreateLigneRegul(vTobNewPiece, vTobAffaire, fTobArticles, vTobTiers, fTobLigne.detail[i], vCleDocOrigine, vInNumRev, pDtRegul);
      AddLigneFacture(vTobNewPiece, fTobLigne.detail[i], vTobVarQte_AFF, vTobAffaire, vTobTiers, vCleDocOrigine, vInNumRev, vInCptVariables, pDtRegul);

      // marquer la ligne régularisée dans la facture d'origine
      vTobUneLigne := fTobLigneRegul.FindFirst(['GL_NATUREPIECEG', 'GL_SOUCHE', 'GL_NUMERO',
                                           'GL_INDICEG', 'GL_NUMORDRE'],
                                           [fTobLigne.detail[i].GetValeur(vInNATUREPIECEG),
                                            fTobLigne.detail[i].GetValeur(vInSOUCHE),
                                            fTobLigne.detail[i].GetValeur(vInNUMERO),
                                            fTobLigne.detail[i].GetValeur(vInINDICEG),
                                            fTobLigne.detail[i].GetValeur(vInNUMORDRE)], True);

      if assigned(vTobUneLigne) then
        vTobUneLigne.PutValue('GL_REGULARISE','X');

      i := i + 1;
    end;

    // valider la derniere facture
    if assigned(vTobNewPiece) then
      ValiderFacture(vTobNewPiece, vTobVarQte_AFF, vTobAffaire, fTobArticles, vTobAdresses, vTobTiers, vCleDocOrigine, vDev, pDtRegul, vInNumRev);
                                        
    // valider les lignes regularisées
    fTobLigneRegul.UpdateDB(False);

    if assigned(vTobAffaire) then vTobAffaire.Free;
    if assigned(vTobAdresses) then vTobAdresses.Free;
    if assigned(vTobTiers) then vTobTiers.Free;
    if assigned(vTobNewPiece) then vTobNewPiece.Free;
    if assigned(vTOBVarQte_LIG) then vTOBVarQte_LIG.Free;
    if assigned(vTobVarQte_AFF) then vTobVarQte_AFF.Free;
  end
  else
    result := False;
end;

function TFACTREGUL.CreatePiece(pTobMaPiece : Tob; pDtRegul : TDateTime) : Tob;
begin
  result := Tob.create('PIECE', nil, -1);

  result.Dupliquer(pTobMaPiece, true, true, false);
  AddLesSupEntete(result);

  result.PutValue('GP_NATUREPIECEG', 'FPR');
  result.PutValue('GP_DATEPIECE', pDtRegul);

  InitTOBPiece(result);

  result.PutValue('GP_CREEPAR','REG');
  result.PutValue('GP_VENTEACHAT', GetInfoParPiece('FPR', 'GPP_VENTEACHAT'));
  result.PutValue('GP_EDITEE','-');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : calcul des coefficients de regul pour toutes les formules
                 de toutes les affaires concernées
Mots clefs ... :
*****************************************************************}
procedure TFACTREGUL.CalculCoefRegul;
var
  vCalculRegul : TCALCULCOEF;
  i            : Integer;
begin

  for i := 0 to fTobLigne.detail.count -1 do
  begin
    // on calcule un coef si on change d'affaire, ou de date
    // CalculCoefRegularisation calcule le coef de toutes les formules
    // de l'affaire a cette date
    if (i = 0) or
       (fTobLigne.detail[i].GetValue('GL_DATEPIECE') <> fTobLigne.detail[i-1].GetValue('GL_DATEPIECE')) or
       (fTobLigne.detail[i].GetValue('GL_AFFAIRE') <> fTobLigne.detail[i-1].GetValue('GL_AFFAIRE')) then
    begin
      vCalculRegul := TCALCULCOEF.create;
      try
        vCalculRegul.CalculCoefRegularisation(fTobLigne.detail[i].Getvalue('GL_AFFAIRE'), fTobLigne.detail[i].GetValue('GL_DATEPIECE'));
      finally
        vCalculRegul.Free;
      end;
    end;
  end;
end;

function TFACTREGUL.ValiderFacture(pTobNewPiece, pTobVarQte_AFF, pTobAffaire,
                                   pTobArticles, pTobAdresses, pTobTiers : Tob;
                                   pCleDoc : R_CleDoc; pDev : RDEVISE;
                                   pDtRegul : TDateTime; var pInNumRev : Integer) : Boolean;
var
  vBoArt      : Boolean;
  i           : Integer;
  vTOBBases   : Tob;
  vTOBBasesL  : Tob;
  vTobPorts   : Tob;
  vTobEches   : Tob;
  vTobNomen   : Tob;
  vInNewNum   : Integer;

begin
  result := True;

  vTOBBases := TOB.Create('Les Bases',Nil,-1);
  vTOBBasesL := TOB.Create('Les Bases LIGNES',Nil,-1);
  vTOBPorts := TOB.Create('Les Ports',Nil,-1);
  vTOBEches := TOB.Create('Les Eches',Nil,-1);
  vTOBNomen := TOB.Create('Les Nomen',Nil,-1);

  LoadTob(pCleDoc, vTOBBases, vTOBBasesL, vTobPorts);

  // Elements variables
  if GetParamSoc('SO_AFVARIABLES') then
    LoadTobVarQte(pTobVarQte_AFF, pTobAffaire, 'LIG', pCleDoc);

  // Revision
  GestionRevisionRegul(pTobNewPiece, pDtRegul, pInNumRev);
  Try
    // renum des lignes de Pièces
    RenumLesLignes(pTobNewPiece);

    // validation fature avec fonction JLD
    if pTobNewPiece.detail.count <> 0 then
    Begin
      ValideLaPeriode(pTobNewPiece);
      MajBasePort(pTobNewPiece, vTobPorts);
      CalculFacture(pTobAffaire,pTobNewPiece, nil,nil,nil,nil,vTOBBases, vTOBBasesL, pTOBTiers, pTOBArticles, vTobPorts, Nil, Nil,nil, pDev);
      RecalculeSousTotaux(pTobNewPiece);
    End;

    // Controle de l'existence de lignes article pour ne pas générer de factures vide
    vBoArt := false;
    for i := 0 to pTobNewPiece.Detail.count-1 do
    Begin
      if (pTobNewPiece.Detail[i].getvalue('GL_TYPELIGNE') = 'ART') then
      Begin
        vBoArt := true;
        break;
      End;
    End;

    if not vBoArt then result := false;

    if result and (pTobNewPiece.detail.count <> 0) then
    Begin
      vInNewNum := RechNumeroPiece('FPR');
      if vInNewNum <> -1 then
        ValideLaPiece(pTobNewPiece, vTOBEches, pTobTiers, vTobPorts, vTOBBases, vTOBBasesL,vTOBNomen, pTobAdresses, vInNewNum, pCleDoc, pDev);
      if (fInNbFac = 0) then
      Begin
        fInFacDeb := pTobNewPiece.getValue('GP_NUMERO');
        fInFacFin := pTobNewPiece.getValue('GP_NUMERO');
      End
      else
        fInFacFin := pTobNewPiece.getValue('GP_NUMERO');
      fInNbFac := fInNbFac + 1;
 
      MajRevision(pTobNewPiece);
      if GetParamSoc('SO_AFVARIABLES') then MajVariable(pTobNewPiece, fTOBVarQte_LIG);
    End;

  finally
    fTobPiece.Free;
    vTobBases.Free;
    vTobEches.Free;
    vTobNomen.Free;
  end;
end;

procedure TFACTREGUL.LoadTob(pCleDocOrigine : R_CleDoc; pTobBases, pTOBBasesL,pTobPorts : Tob);
var
  Q : TQuery;

begin
  // lecture bases
  Q := OpenSQL('SELECT * FROM PIEDBASE WHERE '+ WherePiece(pCleDocOrigine, ttdPiedBase, False), True,-1,'',true);
  pTOBBases.LoadDetailDB('PIEDBASE','','',Q,False) ;
  Ferme(Q);
  // lecture bases lignes
  Q := OpenSQL('SELECT * FROM LIGNEBASE WHERE '+ WherePiece(pCleDocOrigine, ttdLigneBase, False), True,-1,'',true);
  pTOBBasesL.LoadDetailDB('LIGNEBASE','','',Q,False) ;
  Ferme(Q);
  // Lecture Ports   ***Création des pieces, il faut tout prendre ****
  Q:=OpenSQL('SELECT * FROM PIEDPORT WHERE '+ WherePiece(pCleDocOrigine, ttdPorc, False), True,-1,'',true);
  pTOBPorts.LoadDetailDB('PIEDPORT','','',Q,False) ;
  Ferme(Q);
end;

procedure TFACTREGUL.FreeFactures;
begin
  fTobLigne.Free;
  fTobCoef.Free;
  fTobLien.Free;
  fTobArticles.Free;
  fTOBVarQte_LIG.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. : soustraction de la ligne deja facturée
                 et creation de la ligne dans afrevision
Mots clefs ... :
*****************************************************************}
procedure TFACTREGUL.AddLigneFacture(pTobNewPiece, pTobLigne, pTobVarQte_AFF,
                                     pTobAffaire, pTobTiers : Tob; pCleDoc : R_CleDoc;
                                     var pInNumRev, pInCptVariables : Integer;
                                     pDtRegul : TDateTime);
var
  vTobNewLigne  : Tob;
  vStLib        : String;
  vRdPuV        : Double;
  vRdQte        : Double;

begin
  vTobNewLigne := NewTobLigne(pTobNewPiece, 0);

  vTobNewLigne.PutValue('GL_DATEPIECE', pDtRegul);
  vTobNewLigne.AddChampSupValeur('GLLAFFAIRE1','',false);      // clé pour pointer sur tob Revision
  vTobNewLigne.AddChampSupValeur('GLLFORCODE1','',false);
  vTobNewLigne.AddChampSupValeur('GLLNUMEROLIGNE1','',false);
  vTobNewLigne.AddChampSupValeur('GLLAFFAIRE2','',false);      // clé pour pointer sur tob Revision 2
  vTobNewLigne.AddChampSupValeur('GLLFORCODE2','',false);
  vTobNewLigne.AddChampSupValeur('GLLNUMEROLIGNE2','',false);

  vTobNewLigne.AddChampSupValeur('GLLDATEPIECE','',false);
  vTobNewLigne.putValue('GLLDATEPIECE', idate1900);        
                                                  
  TOBCopyFieldValues(pTobLigne, vTobNewLigne, ['GL_ARTICLE','GL_CODEARTICLE',
  'GL_TYPELIGNE','GL_TYPEARTICLE','GL_REFARTSAISIE','GL_RESSOURCE','GL_REMISELIGNE',
  'GL_NONIMPRIMABLE','GL_AFFAIRE','GL_AFFAIRE1','GL_AFFAIRE2','GL_AFFAIRE3','GL_AVENANT',
  'GL_PRIXPOURQTE','GL_FORMULEVAR','GL_GENERAUTO','GL_FACTURABLE','GL_FORCODE1','GL_FORCODE2'
  ,'GL_NUMORDRE', 'GL_QTESTOCK', 'GL_QTETARIF']);

  //articles
  if (pTobLigne.GetValue('GL_TYPELIGNE') = 'ART') then
  Begin
    vStLib := pTobLigne.GetValue('GL_LIBELLE');
    vRdPuV := pTobLigne.GetValue('GL_PUHTDEV');

    // quantité negative
    vRdQte := - pTobLigne.GetValue('GL_QTEFACT');

    CreationLignePiece(pTobNewPiece, vTobNewLigne, pTobAffaire, fTobArticles, pTobTiers,
                       vRdQte, vRdPuV, vStLib, 'A', pTobLigne.GetValue('GL_AFFAIRE'), pCleDoc);

    // reprise des infos personnalisées dans la ligne
    TOBCopyFieldValues(pTobLigne, vTobNewLigne, ['GL_LIBREART1','GL_LIBREART2','GL_LIBREART3','GL_LIBREART4','GL_LIBREART5',
    'GL_LIBREART6','GL_LIBREART7','GL_LIBREART8','GL_LIBREART9','GL_LIBREARTA',
    'GL_FAMILLENIV1','GL_FAMILLENIV2','GL_FAMILLENIV3','GL_DPR','GL_LIBCOMPL']);

    if GetParamSoc('SO_AFVARIABLES') then
      AlimTobVarQte(fTobVarQte_LIG, vTobNewLigne, Nil, pTobVarQte_AFF, 'LIG', pInCptVariables);

  End;

  CreationLigneRevision(pTobNewPiece, vTobNewLigne, pTobLigne, pInNumRev);
end;


procedure TFACTREGUL.CreationLigneRevision(pTobNewPiece, pTobNewLigne, pTobLigne : Tob; var pInNumRev : integer);
var
  vTobCoef      : Tob;
  vTobNewCoef   : Tob;

begin
  // creation de la ligne dans afRevision
  vTobCoef := fTobCoef.FindFirst(['AFR_AFFAIRE',
                                 'AFR_NATUREPIECEG',
                                 'AFR_NUMERO',
                                 'AFR_SOUCHE',
                                 'AFR_INDICEG',
                                 'AFR_NUMORDRE'],
                                 [pTobLigne.getvalue('GL_AFFAIRE'),
                                  pTobLigne.getvalue('GL_NATUREPIECEG'),
                                  pTobLigne.getvalue('GL_NUMERO'),
                                  pTobLigne.getvalue('GL_SOUCHE'),
                                  pTobLigne.getvalue('GL_INDICEG'),
                                  pTobLigne.getvalue('GL_NUMORDRE')], True);

  if vTobCoef <> nil then
  begin

    vTobNewCoef := Tob.create('AFR_REVISION', fTobRevision, -1);
    vTobNewCoef.Dupliquer(vTobCoef, false, true);
    vTobNewCoef.Putvalue('AFR_AFFAIRE', pTobNewLigne.GetValue('GL_AFFAIRE'));
    vTobNewCoef.Putvalue('AFR_NATUREPIECEG', pTobNewLigne.GetValue('GL_NATUREPIECEG'));
    vTobNewCoef.Putvalue('AFR_NUMERO', pTobNewLigne.GetValue('GL_NUMERO'));
    vTobNewCoef.Putvalue('AFR_SOUCHE', pTobNewLigne.GetValue('GL_SOUCHE'));
    vTobNewCoef.Putvalue('AFR_INDICEG', pTobNewLigne.GetValue('GL_INDICEG'));
    vTobNewCoef.Putvalue('AFR_NUMORDRE', pTobNewLigne.GetValue('GL_NUMORDRE'));
    vTobNewCoef.Putvalue('AFR_NUMEROLIGNE', pInNumRev);
                                               
    vTobNewCoef.Putvalue('AFR_ONATUREPIECEG', 'FAC');
    vTobNewCoef.Putvalue('AFR_ONUMERO', pTobLigne.GetValue('GL_NUMERO'));
    vTobNewCoef.Putvalue('AFR_OSOUCHE', pTobLigne.GetValue('GL_SOUCHE'));
    vTobNewCoef.Putvalue('AFR_OINDICEG', pTobLigne.GetValue('GL_INDICEG'));
    vTobNewCoef.Putvalue('AFR_ONUMORDRE', pTobLigne.GetValue('GL_NUMORDRE'));

    if vTobCoef.GetValue('AFR_NUMFORMULE') = 1 then
    begin
      pTobNewLigne.PutValue('GLLAFFAIRE1', vTobNewCoef.GetValue('AFR_AFFAIRE'));
      pTobNewLigne.PutValue('GLLFORCODE1', vTobNewCoef.GetValue('AFR_FORCODE'));
      pTobNewLigne.PutValue('GLLNUMEROLIGNE1', pInNumRev);
      pInNumRev := pInNumRev + 1;
    end
    else
    begin
      pTobNewLigne.PutValue('GLLAFFAIRE2', vTobNewCoef.GetValue('AFR_AFFAIRE'));
      pTobNewLigne.PutValue('GLLFORCODE2', vTobNewCoef.GetValue('AFR_FORCODE'));
      pTobNewLigne.PutValue('GLLNUMEROLIGNE2', pInNumRev);
      pInNumRev := pInNumRev + 1;
    end;
  end;
end;

procedure TFACTREGUL.CreateLigneRegul(pTobNewPiece, pTobAffaire, pTobArticles,
                                      pTobTiers, pTobLigne : Tob; pCleDoc : R_CleDoc;
                                      var pInNumRev : Integer; pDtRegul : TDateTime);
var
  vTobNewLigne  : Tob;
  vStLib        : String;
  vRdPuV        : Double;
  vRdQte        : Double;

begin

  vTobNewLigne := NewTobLigne(pTobNewPiece, 0);

  vTobNewLigne.PutValue('GL_DATEPIECE', pDtRegul);
  vTobNewLigne.AddChampSupValeur('GLLAFFAIRE1','',false);      // clé pour pointer sur tob Revision
  vTobNewLigne.AddChampSupValeur('GLLFORCODE1','',false);
  vTobNewLigne.AddChampSupValeur('GLLNUMEROLIGNE1','',false);
  vTobNewLigne.AddChampSupValeur('GLLAFFAIRE2','',false);      // clé pour pointer sur tob Revision 2
  vTobNewLigne.AddChampSupValeur('GLLFORCODE2','',false);
  vTobNewLigne.AddChampSupValeur('GLLNUMEROLIGNE2','',false);

  // sauvegarde la date de la facture
  vTobNewLigne.AddChampSupValeur('GLLDATEPIECE','',false);
  vTobNewLigne.putValue('GLLDATEPIECE', pTobLigne.Getvalue('GL_DATEPIECE'));

  TOBCopyFieldValues(pTobLigne, vTobNewLigne, ['GL_ARTICLE','GL_CODEARTICLE',
  'GL_TYPELIGNE','GL_TYPEARTICLE','GL_REFARTSAISIE','GL_RESSOURCE','GL_REMISELIGNE',
  'GL_NONIMPRIMABLE','GL_AFFAIRE','GL_AFFAIRE1','GL_AFFAIRE2','GL_AFFAIRE3','GL_AVENANT',
  'GL_PRIXPOURQTE','GL_FORMULEVAR','GL_GENERAUTO','GL_FACTURABLE','GL_FORCODE1','GL_FORCODE2'
  ,'GL_NUMORDRE', 'GL_QTESTOCK', 'GL_QTETARIF']);

  //articles
  if (pTobLigne.GetValue('GL_TYPELIGNE') = 'ART') then
  Begin
    vStLib := pTobLigne.GetValue('GL_LIBELLE');
    vRdPuV := pTobLigne.GetValue('GL_PUHTDEV');
    vRdQte := pTobLigne.GetValue('GL_QTEFACT');

    CreationLignePiece(pTobNewPiece, vTobNewLigne, pTobAffaire, pTobArticles, pTobTiers, vRdQte, vRdPuV, vStLib, 'A', '', pCleDoc);

    // reprise des infos personnalisées dans la ligne
    TOBCopyFieldValues(pTobLigne, vTobNewLigne, ['GL_LIBREART1','GL_LIBREART2','GL_LIBREART3','GL_LIBREART4','GL_LIBREART5',
    'GL_LIBREART6','GL_LIBREART7','GL_LIBREART8','GL_LIBREART9','GL_LIBREARTA',
    'GL_FAMILLENIV1','GL_FAMILLENIV2','GL_FAMILLENIV3','GL_DPR','GL_LIBCOMPL']);
  End;

  // la revision est créer dans LectureCoefFactureRegul
//  CreationLigneRevision(pTobNewPiece, vTobNewLigne, pTobLigne, pInNumRev);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 30/08/2003
Modifié le ... :   /  /
Description .. : chargement des revisions pour les lignes recopiées
                 dans les factures de regul donc pas des coef de regul
Mots clefs ... :
*****************************************************************}
procedure TFACTREGUL.LoadCoef;
var    
  S : String;
  Q : TQuery;

begin
  S := 'SELECT * FROM AFREVISION ';
  S := S + ' WHERE AFR_OKCOEFAPPLIQUE = "X"';
  S := S + ' AND AFR_COEFREGUL = "-"';
  Q := nil;

  try
    Q := OpenSQL(S,True,-1,'',true);
    fTobCoef.LoadDetailDB('AFREVISION', '', '', Q, False) ;
  finally
    Ferme(Q);
  end;
end;

{procedure TFACTREGUL.AddTobLien(pTobOrig, pTobDest : Tob);
var
 vTobLien : Tob;

begin
  vTobLien := tob.create('tob lien', fTobLien, -1);
  vTobLien.AddChampSupValeur('AFFAIRE', pTobDest.Getvalue('GL_AFFAIRE'));
  vTobLien.AddChampSupValeur('AFFAIRE_FACTORIG', pTobOrig.Getvalue('GL_AFFAIRE'));
  vTobLien.AddChampSupValeur('NATUREPIECEG', pTobDest.Getvalue('GL_NATUREPIECEG'));
  vTobLien.AddChampSupValeur('NATUREPIECEG_FACTORIG', pTobOrig.Getvalue('GL_NATUREPIECEG'));
  vTobLien.AddChampSupValeur('NUMERO', pTobDest.Getvalue('GL_NUMERO'));
  vTobLien.AddChampSupValeur('NUMERO_FACTORIG', pTobOrig.Getvalue('GL_NUMERO'));
  vTobLien.AddChampSupValeur('NUMORDRE', pTobDest.Getvalue('GL_NUMORDRE'));
  vTobLien.AddChampSupValeur('NUMORDRE_FACTORIG', pTobOrig.Getvalue('GL_NUMORDRE'));
end;
}

procedure TFACTREGUL.LoadTOBAffaire(var pTobAffaire : Tob; pStCodeAffaire, pStRep : String);
Var
  Q : TQuery;
	S : string;

begin

  if Assigned(pTobAffaire) then pTobAffaire.Free;
  pTobAffaire := Tob.create('AFFAIRE', nil, -1);                               

  S := 'AFF_AFFAIRE,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_DEVISE';
  S := S + ',AFF_PROFILGENER,AFF_AFFAIREHT,AFF_SAISIECONTRE,AFF_APPORTEUR,AFF_ETABLISSEMENT';
  S := S + ',AFF_DATEDEBGENER,AFF_DATEFIN,AFF_DATERESIL,AFF_METHECHEANCE,AFF_PERIODICITE';
  S := S + ',AFF_INTERVALGENER,AFF_GENERAUTO,AFF_LIBELLE,AFF_COMPTAAFFAIRE,AFF_REFEXTERNE';
  S := S + ',AFF_LIBREAFF1,AFF_LIBREAFF2,AFF_LIBREAFF3,AFF_RESPONSABLE,AFF_DATEFINGENER';
  S := S + ',AFF_REGSURCAF,AFF_TIERS,AFF_PRINCIPALE,AFF_FORCODE1,AFF_FORCODE2';

  if pStCodeAffaire = '' then
    pTobAffaire.InitValeurs(False)

  else if pStCodeAffaire <> pTOBAffaire.GetValue('AFF_AFFAIRE') then
  begin
    Q := nil;
    Try
      Q := OpenSQL('SELECT '+ S +' FROM AFFAIRE WHERE AFF_AFFAIRE="'+ pStCodeAffaire+'"',True,-1,'',true) ;
      If (Not Q.EOF) then
        pTobAffaire.SelectDB('',Q);
      pTobAffaire.AddChampSupValeur('LEREP', pStRep, false);
    finally
      Ferme(Q);
    end;
  end;
end;

procedure TFACTREGUL.LoadTobAdresses(var pTobAdresses: Tob; pTobPiece, pTobTiers, pTobAffaire : Tob);
begin
  TiersVersAdresses(pTobTiers, pTobAdresses, pTobPiece, pTOBPiece.GetValue('GP_TIERSFACTURE'));
  AffaireVersAdresses(pTobAffaire, pTobAdresses, pTobPiece);
end;


{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : remplace une sous chaine par une autre dans une chaine
Mots clefs ... :
*****************************************************************}
Procedure TFACTREGUL.ReplaceSubStr(Var pSt : String; pStSup, pStRep : string);
Var
  lInPos, lInLen : Integer;
begin
  lInLen := Length(pStSup);
  While true do begin
    lInPos := Pos(pStSup, pSt);
    if (lInPos > 0) then
      begin
        System.Delete(pSt, lInPos, lInLen);
        Insert(pStRep, pSt, lInPos);
      end
    else
      Exit;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 02/07/2003
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
function TFACTREGUL.GestionRevisionRegul(pTobNewPiece : Tob; pDtEcheance : TDateTime; var pInNumRev : Integer) : boolean;
Var
  i : Integer;
  vTobLig : TOB;

Begin

  for i := 0 to pTobNewPiece.Detail.count-1 do
  Begin
    vTobLig :=  pTobNewPiece.Detail[i];
    if (vToblig.GetValue('GL_TYPELIGNE') = 'ART') and (vTobLig.GetValue('GL_TYPEARTICLE') <> 'POU') then
    begin
    // pas de révision si on gére un cumul sur code
    // car dans ce cas on a perdu le détail des lignes artilces, et le prix correspond au montant cumulé
      if not(TestSiCumulArticle(vTobLig, False, 0, 0, 0, 0)) then
      begin
         vTobLig.PutValue('GL_FORCODE1','');
         vTobLig.PutValue('GL_FORCODE2','');
      end;
    end;
  end;
  result := LectureCoefFactureRegul(pTobNewPiece, fTobRevision ,pDtEcheance, pInNumRev, True);
  if (fTobRevision.detail.count <> 0) then fTobRevision.InsertDB(Nil);
End;


{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 02/07/2003
Modifié le ... :
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TFACTREGUL.MajRevision(pTobNewPiece : Tob);
Var
  vTobLig : TOB;
  i       : Integer;
Begin
  for i :=0  to pTobNewPiece.Detail.count-1 do
  Begin
    vTobLig := pTobNewPiece.Detail[i];
    MajTobRev(fTobRevision, vTobLig,1, False);
    MajTobRev(fTobRevision, vTobLig,2, False);
  end;
  if (fTobRevision.detail.count <> 0) then fTobRevision.InsertDB(Nil);
End;                  

procedure TFACTREGUL.LigneComAff(TOBAffaire, TOBPiece : Tob; Comment : string);
Var
  TobLignes : TOB;
  rLib      : string;
  zaff,zdat : string;
  xemp,xdeb : string;
  xfin      : string;
             
Begin

  TobLignes := NewTOBLigne( TOBPiece,0);
  TobLignes.PutValue('GL_ARTICLE','');
  // pour identifier type de commentaire par rapport au profil
  TobLignes.PutValue('GL_TYPELIGNE','COM');
  // formattage du libellé paramètré
  xdeb := '01/01/1900'; xfin := '31/12/2099'; xemp := '';
  rLib := '{'+Comment+'}';
  rLib := trim(copy(rLib,1,70)); // gm 19/11/02
  TobLignes.PutValue('GL_LIBELLE',rLib);

  PieceVersLigne ( TOBPiece,TobLignes ) ;

  TobLignes.PutValue('GL_AFFAIRE',TOBAffaire.GetValue('AFF_AFFAIRE'));
  TobLignes.PutValue('GL_AFFAIRE1',TOBAffaire.GetValue('AFF_AFFAIRE1'));
  TobLignes.PutValue('GL_AFFAIRE2',TOBAffaire.GetValue('AFF_AFFAIRE2'));
  TobLignes.PutValue('GL_AFFAIRE3',TOBAffaire.GetValue('AFF_AFFAIRE3'));
  TobLignes.PutValue('GL_AVENANT',TOBAffaire.GetValue('AFF_AVENANT'));
  TobLignes.PutValue('GL_REPRESENTANT',TOBAffaire.GetValue('LEREP'));

  zaff := Toblignes.getValue('GL_AFFAIRE');
  zdat := DatetoStr(idate2099);
  GestionChampSuppLigne(Toblignes,zaff,zdat,TobAffaire.GetValue('AFF_TIERS'),zaff);
End;

end.
