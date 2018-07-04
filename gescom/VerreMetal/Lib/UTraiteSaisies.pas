unit UTraiteSaisies;

interface

uses sysutils,classes,windows,messages,hmsgbox,
     HEnt1,Ent1,EntGC,UtilPGI,UTOB,Hctrls,ED_TOOLS, paramsoc,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
     Forms,
		 variants,
     Stdctrls,
     CBPPath,
     AglInit,
     UserChg,
     MajTable,
     Dialogs,
     Inifiles,
     CalcOleGenericBTP,
     AffaireUtil,
     UtilArticle,
     UtilxlsBTP,
     UtilSaisieConso,
     heureUtil,
     UtilFichiers,
     gerepiece;

Var fWorkBook : Variant;
    TOBConso,TOBressources,TOBARTicles : TOB;
    TOBPiece,TOBLignes : TOB;
    LogFile : textfile;
    NbTraiteOk : integer;
    NbTraiteErr : integer;
    LafeuilleCourante : String;
    Nomlog : String;

Function LanceRecupChantiers : Boolean;

Type
  TTypeErreurControle  = (TcdErrNumConso,TcdErrAffaire,TcdErrPhase,TcdErrNatureMouv,TcdErrRessource,TcdErrCodeArticle,TcdErrQte,TcdErrTypeHeure);

implementation

procedure EnregistreMvt(TOBD: TOB);
var messageErr  : string;
begin
  if TOBD.getValue('BCO_NATUREMOUV') = 'FAN' then
  begin
    if V_PGI.ioError = oeOK then
      MessageErr := Format ('Enregistrement réussi pour le chantier %s, nature %s et fournisseur %s',
                            [TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_NATUREMOUV'),TOBD.getValue('BCO_LIBELLE')])
    else
      MessageErr := Format ('Erreur enregistrement pour le chantier %s, nature %s et fournisseur %s',
                            [TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_NATUREMOUV'),TOBD.getValue('BCO_LIBELLE')]);
  end else if TOBD.getValue('BCO_NATUREMOUV') = 'FOU' then
  begin

    if V_PGI.ioError = oeOK then
      MessageErr := Format ('Enregistrement réussi pour le chantier %s, nature %s et article %s',
                            [TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_NATUREMOUV'),TOBD.getValue('BCO_CODEARTICLE')])
    else
      MessageErr := Format ('Erreur enregistrement pour le chantier %s, nature %s et article %s',
                            [TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_NATUREMOUV'),TOBD.getValue('BCO_CODEARTICLE')]);
  end else
  begin
    if V_PGI.ioError = oeOK then
      MessageErr := Format ('Enregistrement réussi pour le chantier %s, nature %s et ressource %s',
                            [TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_NATUREMOUV'),TOBD.getValue('BCO_RESSOURCE')])
    else
      MessageErr := Format ('Erreur enregistrement pour le chantier %s, nature %s et ressource %s',
                            [TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_NATUREMOUV'),TOBD.getValue('BCO_RESSOURCE')]);
  end;

  Writeln(Logfile,messageErr);
end;

procedure EnregistreConso;
var Indice : integer;
		TOBD : TOB;
begin
//
	for Indice := 0 to TOBConso.detail.count -1 do
  begin
  	TOBD := TOBConso.detail[Indice];
  	if TOBD.GetValue('ERREUR') = -1 then
    begin
    	if not TOBD.InsertDB (nil) then
      begin
        V_PGI.ioError := oeUnknown;
  		  inc(NbTraiteErr);
      end else
        V_PGI.ioError := oeOK;

      EnregistreMvt(TOBD);
  		inc(NbTraiteOk);
    end;
  end;
end;

procedure EnregistreFinTrait;
var MessageErr : string;
begin
  MessageErr := Format ('FIN DE TRAITEMENT : %d lignes enregistrées, %d erreurs',[NbTraiteOk, NbTraiteErr]);
  Writeln(Logfile,messageErr);

  PGIBox('Traitement terminé',Application.Title);
end;

function AddArticle(TOBD: TOB): TOB;
var QQ : TQuery;
begin
	result := nil;
	if TOBD.GetValue('BCO_ARTICLE')='' then exit;
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBD.GetValue('BCO_ARTICLE')+'"',true,-1, '', True);
  if not QQ.eof then
  begin
  	result := TOB.Create ('ARTICLE',TOBArticles,-1);
  	result.selectdb('',QQ);
  end;
  ferme (QQ);
end;

function AddRessource(TOBD: TOB): TOB;
var QQ : TQuery;
begin
  Result := Nil;
  QQ := OpenSql ('SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+TOBD.GetValue('BCO_RESSOURCE')+'"',true,-1, '', True);
  if not QQ.eof then
  begin
  	result := TOB.Create ('RESSOURCE',TOBressources,-1);
  	result.selectdb('',QQ);
  end;
  ferme (QQ);
end;

procedure SetErreur(TOBD: TOB);
var messageErr : string;
		TheErreur : string;
begin
  if TOBD.GetVAlue('ERREUR') = TcdErrNumConso then
  BEGIN
  	TheErreur := 'Problème d''affectation du N° de consommation';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrAffaire then
  BEGIN
  	TheErreur := 'Ce chantier ('+TOBD.GetValue('BCO_AFFAIRE')+') n''existe pas';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrPhase then
  BEGIN
  	TheErreur := 'Cette phase ('+TOBD.GetValue('BCO_PHASETRA')+') sur le chantier ('+TOBD.GetValue('BCO_AFFAIRE')+') n''existe pas';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrNatureMouv then
  BEGIN
  	TheErreur := 'La nature de mouvement ('+TOBD.GetValue('BCO_NATUREMOUV')+')  n''existe pas';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrRessource then
  BEGIN
  	TheErreur := 'La ressource ('+TOBD.GetValue('BCO_RESSOURCE')+')  n''existe pas';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrCodeArticle then
  BEGIN
  	TheErreur := 'La prestation ('+TOBD.GetValue('BCO_CODEARTICLE')+')  n''existe pas';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrQte then
  BEGIN
  	TheErreur := 'La quantité n''est pas renseignée';
	  inc(NbTraiteErr);
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrTypeHeure then
  BEGIN
  	TheErreur := 'Le Type d''heure ('+TOBD.GetValue('BCO_TYPEHEURE')+')  n''existe pas';
	  inc(NbTraiteErr);
  END;
  MessageErr := Format ('Anomalie sur %s ligne %s : %s',[LaFeuilleCourante,TOBD.getValue('NOLIGNE'),TheErreur]);
  Writeln(Logfile,messageErr);
end;

procedure FinaliseConsommation(TOBD: TOB);
var Part0				: string;
    Part1				: string;
    Part2				: String;
    Part3				: String;
    Avenant 		: string;
    Code        : string;
    TypRes			: string;
    NumConso    : double;
    Annee     : word;
    Mois      : word;
    Jour      : word;
    Semaine   : integer;
    DateMouv  : TDateTime;
    TOBres,TOBA  : TOB;
begin
  if (GetNumUniqueConso (NumConso) <> Gncabort) then
  begin
  	if TOBD.GetValue('BCO_CODEARTICLE') <> '' then TOBD.putValue('BCO_ARTICLE',CodeArticleUnique (TOBD.GEtValue('BCO_CODEARTICLE'),'','','','',''))
    else  TOBD.putValue('BCO_ARTICLE','');

    TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBD.getValue('BCO_ARTICLE')],true);
    if TOBA = nil then TOBA := AddArticle(TOBD);

  	TOBD.putValue('BCO_NUMMOUV',NumConso);
    Code := TOBD.GetValue('BCO_AFFAIRE');
    BTPCodeAffaireDecoupe (Code,Part0,Part1,Part2,Part3,Avenant,taModif,false);
    TOBD.putValue('BCO_AFFAIRE0',Part0);
    TOBD.putValue('BCO_AFFAIRE1',Part1);
    TOBD.putValue('BCO_AFFAIRE2',Part2);
    TOBD.putValue('BCO_AFFAIRE3',Part3);
    TOBD.putValue('BCO_QTEFACTUREE',TOBD.GetValue('BCO_QUANTITE'));
    TOBD.putValue('BCO_QTEINIT',TOBD.GetValue('BCO_QUANTITE'));
//    if TOBA <> nil then TOBD.putValue('BCO_QUALIFQTEMOUV',TOBA.GetValue('GA_QUALIFUNITEVTE'));

    // Découpage de la date
    DateMouv:=StrToDate(TOBD.GetValue('BCO_DATEMOUV'));
    DecodeDate(DateMouv, Annee, Mois, Jour);
    // Recherche du Numéro de semaine
    Semaine := NumSemaine(DateMouv);

    TOBD.PutValue('BCO_MOIS', Mois);
    TOBD.PutValue('BCO_FACTURABLE', 'N');
    TOBD.PutValue('BCO_VALIDE', '-');
    TOBD.PutValue('BCO_SEMAINE', Semaine);

    calculeLaLigne (TOBD);

    // Mise à jour libellé :
    if (TOBD.GetValue('BCO_NATUREMOUV')<>'RAN') then TOBD.putValue('BCO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));

  end else
  begin
  	TOBD.putValue('ERREUR',TcdErrNumConso);
    SetErreur (TOBD);
    exit;
  end;
end;

procedure ChargeTobConsos(DateSaisie : TDateTime ; CodeAffaire, CodeArticle, NatureMouv : string ; Montant: Double ; Ligne : integer);
var TOBD : TOB;
begin
	TOBD := TOB.Create ('CONSOMMATIONS',TOBConso,-1);
  TOBD.InitValeurs;
  AddLesSupLignesConso (TOBD);
  TOBD.AddChampSupValeur ('ERREUR',-1);
	TOBD.addChampSupValeur('NOLIGNE',inttostr(Ligne));
  TOBD.putValue('BCO_AFFAIRE',CodeAffaire);
  TOBD.putValue('BCO_PHASETRA','');
  TOBD.putValue('BCO_DATEMOUV',DateSaisie);
  TOBD.putValue('BCO_NATUREMOUV',NatureMouv);
  if NatureMouv = 'RAN' then
  begin
    TOBD.putValue('BCO_RESSOURCE','');
    TOBD.putValue('BCO_CODEARTICLE','');
    TOBD.putValue('BCO_LIBELLE','Reprise facturation LSE');
    TOBD.putValue('BCO_QUANTITE',-1);
    TOBD.putValue('BCO_DPA',Montant);
    TOBD.putValue('BCO_DPR',Montant);
    TOBD.putValue('BCO_PUHT',Montant);
  end else
  begin
    TOBD.putValue('BCO_RESSOURCE','');
    TOBD.putValue('BCO_CODEARTICLE',CodeArticle);
    TOBD.putValue('BCO_QUANTITE',1);
    TOBD.putValue('BCO_DPA',Montant);
    TOBD.putValue('BCO_DPR',Montant);
    TOBD.putValue('BCO_PUHT',0);
  end;
  TOBD.putValue('BCO_QUALIFQTEMOUV','');
  TOBD.putValue('BCO_TYPEHEURE','');
//  ControleDonnee (TOBD);
  if TOBD.GetValue('ERREUR')=-1 then
  begin
  	FinaliseConsommation(TOBD);
  end;
end;

Function TraiteBudgetActualise : Boolean;
Var i, lig, col, nbcol : integer;
    Chaine, CodeAffaire, Article, Quantite, Unite : String;
    Qte, Montant : Double;
begin
//      PgiInfo ('Traitement Budget Actualise','');

  Result := True;
  LaFeuilleCourante := 'Feuille1';
  For i:=5 to 200 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'D'+intToStr(i));
    if Chaine = '' then Exit;

    CodeAffaire := 'A'+Copy(Chaine,5,6)+'000     00';

    // Création pour entete piece
    TOBPiece := Tob.Create ('', nil, -1);

    TobPiece.AddChampSupValeur ('NATUREPIECEG', 'PBT');
    TobPiece.AddChampSupValeur ('AFFAIRE', CodeAffaire);
    TobPiece.AddChampSupValeur ('TIERS', getTiersAffaire(CodeAffaire));
    TobPiece.AddChampSupValeur ('ETABLISSEMENT', GetParamSocSecur('SO_ETABLISDEFAUT','001'));
    TobPiece.AddChampSupValeur ('DOMAINE', '');
    TobPiece.AddChampSupValeur ('DATEPIECE', V_PGI.DateEntree);
    TobPiece.AddChampSupValeur ('REFINTERNE', 'Reprise LSE');

    // Création pour maj lignes de pièce : total production
    Article := 'REPLOTCA';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'I'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', Montant);
    TobLignes.AddChampSupValeur ('PUHT', Montant);
    TobLignes.AddChampSupValeur ('DPA', 0.0);
    TobLignes.AddChampSupValeur ('DPR', 0.0);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé MO
    Article := 'PREPLOTMO';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'J'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé ST
    Article := 'PREPLOTST';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'K'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé FOURNITURES
    Article := 'REPLOTFO';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'L'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé Frais de Chantiers
    Article := 'PREPLOTFC';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'M'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    if (TOBPiece <> nil) and (TOBPiece.Detail.count > 0) then
    begin
      CreatePieceFromTob (TobPiece);
    end;

    TOBPiece.free;
    TOBPiece := nil;
  end;
end;

Function TraiteBudgetAdate : Boolean;
Var i, lig, col, nbcol : integer;
    Chaine, CodeAffaire, Article, Quantite, Unite : String;
    Qte, Montant : Double;
begin
//      PgiInfo ('Traitement Budget A Date','');

  Result := True;
  LaFeuilleCourante := 'Feuille1';
  For i:=5 to 200 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'D'+intToStr(i));
    if Chaine = '' then Exit;

    CodeAffaire := 'A'+Copy(Chaine,5,6)+'000     00';

    // Création pour entete piece
    TOBPiece := Tob.Create ('', nil, -1);

    TobPiece.AddChampSupValeur ('NATUREPIECEG', 'BCE');
    TobPiece.AddChampSupValeur ('AFFAIRE', CodeAffaire);
    TobPiece.AddChampSupValeur ('TIERS', getTiersAffaire(CodeAffaire));
    TobPiece.AddChampSupValeur ('ETABLISSEMENT', GetParamSocSecur('SO_ETABLISDEFAUT','001'));
    TobPiece.AddChampSupValeur ('DOMAINE', '');
    TobPiece.AddChampSupValeur ('DATEPIECE', V_PGI.DateEntree);
    TobPiece.AddChampSupValeur ('REFINTERNE', 'Reprise LSE');

    // Création pour maj lignes de pièce : Déboursé MO
    Article := 'PREPLOTMO';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'O'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé ST
    Article := 'PREPLOTST';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'P'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé FOURNITURES
    Article := 'REPLOTFO';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'Q'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    // Création pour maj lignes de pièce : Déboursé Frais de Chantiers
    Article := 'PREPLOTFC';
    Qte := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'R'+intToStr(i));
    Montant := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', V_PGI.DateEntree);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', 0.0);
    TobLignes.AddChampSupValeur ('PUHT', 0.0);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    if (TOBPiece <> nil) and (TOBPiece.Detail.count > 0) then
    begin
      CreatePieceFromTob (TobPiece);
    end;

    TOBPiece.free;
    TOBPiece := nil;
  end;
end;

Function TraiteCumuleAdate : Boolean;
Var i, lig, col, nbcol : integer;
    Chaine, CodeAffaire : String;
    Montant : Double;
begin
//      PgiInfo ('Traitement Cumule A date','');
  Result := True;
  LaFeuilleCourante := 'Feuille1';
  For i:=5 to 200 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'D'+intToStr(i));
    if Chaine = '' then Exit;

    CodeAffaire := 'A'+Copy(Chaine,5,6)+'000     00';

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'X'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, '', 'RAN', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'Y'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'PREPLOTMO', 'MO', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'Z'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'PREPLOTST', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'AA'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REPLOTFO', 'FOU', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'AB'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'PREPLOTFC', 'EXT', Montant, i);
    end;
end;

Function TraiteLesDonnees (filename : String) : Boolean;
begin
  Result := True;
  LaFeuilleCourante := '';

  // tester existence affaire
  Application.BringToFront;
  Application.ShowMainForm:=True;

  NomLog := ChangeFileExt (filename,'.log');
  AssignFile (LogFile,NomLog);
  rewrite(Logfile);

  TOBConso := TOB.Create ('LES CONSO',nil,-1);
  TOBressources := TOB.Create ('LES RESSOURCES',nil,-1);
	TOBArticles := TOB.Create ('LES ARTICLES',nil,-1);

  TOBPiece := nil;

  NbTraiteOk := 0;

  TraiteBudgetActualise;
  TraiteBudgetAdate;
  TraiteCumuleAdate;

  if TOBConso.detail.count > 0 then
  begin
    EnregistreConso;
    TOBConso.ClearDetail;
  end;

	EnregistreFinTrait;

  TOBConso.free;
  TOBressources.free;
  TOBArticles.free;

  CloseFile(logfile);

end;

Function LanceRecupChantiers : Boolean;
var filename, LocalRepert, filenamecourt, filenamearchive : string;
    IniPGIFile : Tinifile;
    ODExcelFile: TOpenDialog;
    fWinExcel: OleVariant;
    fnewInst : boolean;
    Affaire, CodeAffaire : String;
    DateSaisie : TDateTime;
begin
  Result := True;

  InitMoveProgressForm(Application.MainForm,Application.Title, 'Traitement en cours, veuillez patienter...', 200, False, true) ;

  ODExcelFile := TOpenDialog.create(nil) ;
  ODExcelFile.Filter := 'fichiers Excel (*xls)|*.xls';
  if ODExcelFile.execute then
  begin
    Filename := ODexcelFile.FileName;
  end;
  ODExcelFile.free;

  if filename <> '' then
  begin
    if not OpenExcel(true,fWinExcel,fNewInst) then
    begin
      PgiError ('Excel n''est pas installé sur ce poste','');
      Result:=False;
      FiniMoveProgressForm ;
      exit;
    end;
    fWorkBook := OpenWorkBook(filename ,fWinExcel);
    TRY
//      PgiInfo ('Traitement du document Excel : '+filename,'');
      Result:=TraiteLesDonnees (filename);
    EXCEPT
      PgiInfo ('Traitement impossible du document Excel','');
      Result:=False;
    END;

    ExcelClose (fwinExcel);
  end;

  FiniMoveProgressForm ;

end ;

end.

