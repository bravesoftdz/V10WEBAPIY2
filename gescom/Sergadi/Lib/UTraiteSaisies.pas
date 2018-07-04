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

Function TraiteLaSaisie : Boolean;
Function TraiteLesFeuilles (filename, CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Function TraiteLesRessourcesHumaines (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Function TraiteLesMaterielsSergadi (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Function TraiteLesMaterielsLocation (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Function TraiteLesMateriaux (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Function TraiteLesPiecesStock (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Function TraiteLesPiecesHorsStock (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
function IsExisteArticle (TOBD : TOB) : boolean;
function OKTypeHeure(Typeheure : string) : boolean;
procedure LanceSaisieSERGADI (Param : string) ;

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

procedure ControleDonnee(TOBD: TOB);
var PrestationDefaut : string;
begin
	PrestationDefaut := '';
  if (TOBD.GetValue('BCO_AFFAIRE') = '') or (not OkChantier(TOBD.GetValue('BCO_AFFAIRE'))) then
  BEGIN
  	TOBD.putValue('ERREUR',TcdErrAffaire);
    SetErreur (TOBD);
    Exit;
  END;
  if (TOBD.GetValue('BCO_PHASETRA')<>'') and (not IsExistePhaseAffaire (TOBD.GetValue('BCO_AFFAIRE'),TOBD.GetValue('BCO_PHASETRA'))) then
  begin
  	TOBD.putValue('ERREUR',TcdErrPhase);
    SetErreur (TOBD);
    Exit;
  end;
  if (TOBD.GetValue('BCO_NATUREMOUV') <> 'MO') and
     (TOBD.GetValue('BCO_NATUREMOUV') <> 'EXT') and
     (TOBD.GetValue('BCO_NATUREMOUV') <> 'RES') and
     (TOBD.GetValue('BCO_NATUREMOUV') <> 'FOU') and
     (TOBD.GetValue('BCO_NATUREMOUV') <> 'FAN') then
  begin
  	TOBD.putValue('ERREUR',TcdErrNatureMouv);
    SetErreur (TOBD);
    Exit;
  end;
  if (TOBD.GetValue('BCO_NATUREMOUV') <> 'FOU') and (TOBD.GetValue('BCO_NATUREMOUV') <> 'FAN') then
  begin
    if TOBD.GetValue('BCO_RESSOURCE') = '' then
    begin
      TOBD.putValue('ERREUR',TcdErrRessource);
      SetErreur (TOBD);
      Exit;
    end;
    if not OKRessource (nil,TOBD.GetValue('BCO_RESSOURCE'),'',PrestationDefaut,false) then
    begin
      TOBD.putValue('ERREUR',TcdErrRessource);
      SetErreur (TOBD);
      Exit;
    end;
  end;
  if (TOBD.GetValue('BCO_CODEARTICLE')<>'') and (not IsExisteArticle(TOBD)) then
  begin
  	TOBD.putValue('ERREUR',TcdErrCodeArticle);
    SetErreur (TOBD);
    Exit;
  end;
  // si le code prestation n'est pas renseigné, on récupère celui-ci de la ressource
  if (TOBD.GetValue('BCO_CODEARTICLE')='') and (TOBD.GetValue('BCO_NATUREMOUV')<>'FAN') then
  begin
  	if PrestationDefaut<>'' then
      TOBD.PutValue('BCO_CODEARTICLE', PrestationDefaut)
    else begin
    	TOBD.putValue('ERREUR',TcdErrCodeArticle);
    	SetErreur (TOBD);
    	Exit;
    end
  end;
  if TOBD.GetValue('BCO_QUANTITE') = 0 then
  begin
  	TOBD.putValue('ERREUR',TcdErrQte);
    SetErreur (TOBD);
    Exit;
  end;
  if (TOBD.GetValue('BCO_TYPEHEURE')<>'') and ( not OKTypeHeure(TOBD.GetValue('BCO_TYPEHEURE'))) then
  begin
  	TOBD.putValue('ERREUR',TcdErrTypeHeure);
    SetErreur (TOBD);
    Exit;
  end;
end;

function IsExisteArticle (TOBD : TOB) : boolean;
var Q : Tquery;
begin
  Q := OpenSql ('SELECT GA_ARTICLE FROM ARTICLE  WHERE GA_CODEARTICLE="'+
                 TOBD.GetValue('BCO_CODEARTICLE') + '" AND GA_STATUTART <> "DIM" ',true,-1, '', True);
  result := not Q.eof;
  if result then TOBD.putValue('BCO_ARTICLE',Q.findfield('GA_ARTICLE').asString);
  ferme (Q);
end;

function OKTypeHeure(Typeheure : string) : boolean;
var Sql : string;
	  QQ : TQuery;
begin
	result := true;
  if (Typeheure = '' ) then exit;
  Sql := 'SELECT CC_LIBELLE,CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="ATH" AND CC_CODE="'+Typeheure+'"';
  QQ := OpenSql (SQl,true,-1, '', True);
  result :=  not QQ.eof;
  ferme (QQ);
end;

procedure PrepaLivraisonChantier(DateSaisie : TDateTime ; CodeAffaire, Article, Unite : string ; Quantite: Double);
begin
  // Création pour maj stock lignes de pièce
  TobLignes := Tob.Create ('', TobPiece, -1);

  TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
  TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
  TobLignes.AddChampSupValeur ('QTEFACT', Quantite);
  TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
  TobLignes.AddChampSupValeur ('DATELIVRAISON', DateSaisie);
  TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
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
    TOBres := TOBressources.findFirst(['ARS_RESSOURCE'],[TOBD.getValue('BCO_RESSOURCE')],true);
    if TOBREs = nil then TOBres := AddRessource(TOBD);

    // Si le code article n'est pas renseigné, on récupère la prestation par défaut de la ressource
    if TOBD.GetValue('BCO_CODEARTICLE') = '' then
    begin
      if TOBREs <> nil then TOBD.PutValue('BCO_ARTICLE', TOBres.GetValue('ARS_ARTICLE'));
    end else
    	TOBD.putValue('BCO_ARTICLE',CodeArticleUnique (TOBD.GEtValue('BCO_CODEARTICLE'),'','','','',''));

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
    if TOBA <> nil then TOBD.putValue('BCO_QUALIFQTEMOUV',TOBA.GetValue('GA_QUALIFUNITEVTE'));

    // Découpage de la date
    DateMouv:=StrToDate(TOBD.GetValue('BCO_DATEMOUV'));
    DecodeDate(DateMouv, Annee, Mois, Jour);
    // Recherche du Numéro de semaine
    Semaine := NumSemaine(DateMouv);

    TOBD.PutValue('BCO_MOIS', Mois);
    TOBD.PutValue('BCO_FACTURABLE', 'N');
    TOBD.PutValue('BCO_VALIDE', '-');
    TOBD.PutValue('BCO_SEMAINE', Semaine);

    // récupération du type de resssource
		TypRes:=GetTypeRessource (TOBD.GetValue('BCO_RESSOURCE'));

    if (TOBD.getValue('BCO_DPA')=0) then
    begin
    	if (TOBD.GetValue('BCO_NATUREMOUV')='MO') or
      	 (TOBD.GetValue('BCO_NATUREMOUV')='RES') or
         ((TOBD.GetValue('BCO_NATUREMOUV')='EXT') and ((TypRes='INT') or (TypRes='ST'))) then
      begin
         SetValoFromRessource (TOBD,TOBREs,[TmvPa]);
      end else
      begin
      	if TOBD.GetValue('BCO_ARTICLE')<> '' then SetvaloArticle (TOBD,TOBA,[TmvPa]);
      end;
    end;
    if (TOBD.GetValue('BCO_DPR')=0) then
    begin
    	if (TOBD.GetValue('BCO_NATUREMOUV')='MO') or
      	 (TOBD.GetValue('BCO_NATUREMOUV')='RES') or
         ((TOBD.GetValue('BCO_NATUREMOUV')='EXT') and ((TypRes='INT') or (TypRes='ST'))) then
      begin
         SetValoFromRessource (TOBD,TOBREs,[TmvPR]);
      end else
      begin
      	if TOBD.GetValue('BCO_ARTICLE')<> '' then SetvaloArticle (TOBD,TOBA,[TmvPr]);
      end;
    end;
    if (TOBD.GetValue('BCO_PUHT')=0) then
    begin
    	if (TOBD.GetValue('BCO_NATUREMOUV')='MO') or
      	 (TOBD.GetValue('BCO_NATUREMOUV')='RES') or
         ((TOBD.GetValue('BCO_NATUREMOUV')='EXT') and ((TypRes='INT') or (TypRes='ST'))) then
      begin
         SetValoFromRessource (TOBD,TOBREs,[TmvPV]);
      end else
      begin
      	if TOBD.GetValue('BCO_ARTICLE')<> '' then SetvaloArticle (TOBD,TOBA,[TmvPv]);
      end;
    end;
    calculeLaLigne (TOBD);

    // Mise à jour libellé :
    if (TOBD.GetValue('BCO_NATUREMOUV')<>'FAN') then
      TOBD.putValue('BCO_LIBELLE','@ '+TOBA.GetValue('GA_LIBELLE'))


  end else
  begin
  	TOBD.putValue('ERREUR',TcdErrNumConso);
    SetErreur (TOBD);
    exit;
  end;
end;

procedure ChargeTobConsos(DateSaisie : TDateTime ; CodeAffaire, Ressource, NatureMouv : string ; Quantite: Double ; Ligne : integer);
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
  if NatureMouv = 'FAN' then
  begin
    TOBD.putValue('BCO_RESSOURCE','');
    TOBD.putValue('BCO_CODEARTICLE','');
    TOBD.putValue('BCO_LIBELLE','@ '+Ressource);
    TOBD.putValue('BCO_QUANTITE',1);
    TOBD.putValue('BCO_DPA',Quantite);
    TOBD.putValue('BCO_DPR',Quantite);
    TOBD.putValue('BCO_PUHT',Quantite);
  end else
  begin
    if NatureMouv = 'FOU' then
    begin
      TOBD.putValue('BCO_RESSOURCE','');
      TOBD.putValue('BCO_CODEARTICLE',Ressource);
    end else
    begin
      TOBD.putValue('BCO_RESSOURCE',Ressource);
      TOBD.putValue('BCO_CODEARTICLE','');
    end;
    TOBD.putValue('BCO_QUANTITE',Quantite);
    TOBD.putValue('BCO_DPA',0);
    TOBD.putValue('BCO_DPR',0);
    TOBD.putValue('BCO_PUHT',0);
  end;
  TOBD.putValue('BCO_QUALIFQTEMOUV','');
  TOBD.putValue('BCO_TYPEHEURE','');
  ControleDonnee (TOBD);
  if TOBD.GetValue('ERREUR')=-1 then
  begin
  	FinaliseConsommation(TOBD);
  end;
end;

Function TraiteLesRessourcesHumaines (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Var i, lig, col, nbcol : integer;
    Ressource, NbHeures : String;
    nbh : Double;
begin
  Result := True;
  LaFeuilleCourante := 'Ressources humaines';
  For i:=13 to 24 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Ressource := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Ressource <> '' then
    begin
      ExcelFindRange (fWorkBook, LaFeuilleCourante, 'G'+intToStr(i),lig,col,nbcol);
      NbHeures := GetExcelFormated (fWorkBook.activeSheet,lig,col);
      if NbHeures <> '' then
      begin
        NbHeures := StringReplace (NbHeures,':',',',[rfReplaceAll]);
        Nbh := StrToFloat(NbHeures);
        Nbh := HeureBase60To100 (Nbh);
      end else
      begin
        Nbh := 0;
      end;
      // Traitement de la ligne
      ChargeTOBConsos(DateSaisie, CodeAffaire, Ressource, 'MO', Nbh, i);
    end;
  end;
end;

Function TraiteLesMaterielsSergadi (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Var i, lig, col, nbcol : integer;
    Ressource, NbHeures : String;
    nbh : Double;
begin
  Result := True;
  LaFeuilleCourante := 'Matériel SERGADI';
  For i:=13 to 32 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Ressource := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Ressource <> '' then
    begin
      ExcelFindRange (fWorkBook, LaFeuilleCourante, 'D'+intToStr(i),lig,col,nbcol);
      NbHeures := GetExcelFormated (fWorkBook.activeSheet,lig,col);
      if NbHeures <> '' then
      begin
        Nbh := StrToFloat(NbHeures);
      end else
      begin
        Nbh := 0;
      end;
      // Traitement de la ligne
      ChargeTOBConsos(DateSaisie, CodeAffaire, Ressource, 'RES', Nbh, i);
    end;
  end;
end;

Function TraiteLesMaterielsLocation (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Var i, lig, col, nbcol : integer;
    Ressource, NbHeures : String;
    nbh : Double;
begin
  Result := True;
  LaFeuilleCourante := 'Matériel Location';
  For i:=13 to 27 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Ressource := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Ressource <> '' then
    begin
      ExcelFindRange (fWorkBook, LaFeuilleCourante, 'D'+intToStr(i),lig,col,nbcol);
      NbHeures := GetExcelFormated (fWorkBook.activeSheet,lig,col);
      if NbHeures <> '' then
      begin
        Nbh := StrToFloat(NbHeures);
      end else
      begin
        Nbh := 0;
      end;
      // Traitement de la ligne
      ChargeTOBConsos(DateSaisie, CodeAffaire, Ressource, 'EXT', Nbh, i);
    end;
  end;
end;

Function TraiteLesMateriaux (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Var i, lig, col, nbcol : integer;
    Article, Quantite : String;
    Qte : Double;
begin
  Result := True;
  LaFeuilleCourante := 'Matériaux';
  For i:=13 to 24 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Article := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Article <> '' then
    begin
      ExcelFindRange (fWorkBook, LaFeuilleCourante, 'C'+intToStr(i),lig,col,nbcol);
      Quantite := GetExcelFormated (fWorkBook.activeSheet,lig,col);
      if Quantite <> '' then
      begin
        Qte := StrToFloat(Quantite);
      end else
      begin
        Qte := 0;
      end;
      // Traitement de la ligne
      ChargeTOBConsos(DateSaisie, CodeAffaire, Article, 'FOU', Qte, i);
    end;
  end;
end;

Function TraiteLesPiecesStock (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Var i, lig, col, nbcol : integer;
    Article, Quantite, Unite : String;
    Qte : Double;
begin
  Result := True;
  LaFeuilleCourante := 'Pièces stock';
  For i:=13 to 62 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Article := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Article <> '' then
    begin
      if TOBPiece = nil then
      begin
        // Création pour maj stock (entete piece)
        TOBPiece := Tob.Create ('', nil, -1);

        TobPiece.AddChampSupValeur ('NATUREPIECEG', 'LBT');
        TobPiece.AddChampSupValeur ('AFFAIRE', CodeAffaire);
        TobPiece.AddChampSupValeur ('TIERS', getTiersAffaire(CodeAffaire));
        TobPiece.AddChampSupValeur ('ETABLISSEMENT', GetParamSocSecur('SO_ETABLISDEFAUT','001'));
        TobPiece.AddChampSupValeur ('DOMAINE', '');
        TobPiece.AddChampSupValeur ('DATEPIECE', DateSaisie);
        TobPiece.AddChampSupValeur ('REFINTERNE', 'Intégration EXCEL');
      end;
      ExcelFindRange (fWorkBook, LaFeuilleCourante, 'G'+intToStr(i),lig,col,nbcol);
      Quantite := GetExcelFormated (fWorkBook.activeSheet,lig,col);
      if Quantite <> '' then
      begin
        Qte := StrToFloat(Quantite);
      end else
      begin
        Qte := 0;
      end;
      Unite := ExcelGetValue(fWorkBook, LaFeuilleCourante,'G'+intToStr(i));

      // Création pour maj stock lignes de pièce
      TobLignes := Tob.Create ('', TobPiece, -1);
      //
      TobLignes.AddChampSupValeur('TYPELIGNE', 'ART');
      //
      TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
      TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
      TobLignes.AddChampSupValeur ('LIBELLE', '@ '+ GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
      TobLignes.AddChampSupValeur ('QTEFACT', Qte);
      TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
      TobLignes.AddChampSupValeur ('DATELIVRAISON', DateSaisie);
      TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));

    end;
  end;
  if (TOBPiece <> nil) and (TOBPiece.Detail.count > 0) then
  begin
     CreatePieceFromTob (TobPiece);
  end;
  TOBPiece.free;
  TOBPiece := nil;
end;


Function TraiteLesPiecesHorsStock (CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
Var i, lig, col, nbcol : integer;
    Fournisseur, Montant : String;
    Mt : Double;
begin
  Result := True;
  LaFeuilleCourante := 'Pièces hors-stock';
  For i:=13 to 22 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Fournisseur := ExcelGetValue(fWorkBook, LaFeuilleCourante,'B'+intToStr(i));
    if Fournisseur <> '' then
    begin
      ExcelFindRange (fWorkBook, LaFeuilleCourante, 'E'+intToStr(i),lig,col,nbcol);
//      Montant := GetExcelFormated (fWorkBook.activeSheet,lig,col);
      Montant := EXCELGETVALUE (fWorkBook,'Pièces hors-stock','E'+intToStr(i));
      if Montant <> '' then
      begin
        Mt := StrToFloat(Montant);
      end else
      begin
        Mt := 0;
      end;
      // Traitement de la ligne
      ChargeTOBConsos(DateSaisie, CodeAffaire, Fournisseur, 'FAN', Mt, i);
    end;
  end;
end;

Function TraiteLesFeuilles (filename, CodeAffaire : String ; DateSaisie : TDateTime) : Boolean;
begin
  Result := True;
  LaFeuilleCourante := '';

  // tester existence affaire
  if Not ExisteAffaire(CodeAffaire,'') then
  begin
    PgiInfo ('Traitement impossible, l''affaire ' + CodeAffaire + ' saisie dans le document n''existe pas');
    Result := False;
    Exit;
  end;

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

  TraiteLesRessourcesHumaines (CodeAffaire, DateSaisie);
  TraiteLesMaterielsSergadi (CodeAffaire, DateSaisie);
  TraiteLesMaterielsLocation (CodeAffaire, DateSaisie);
  TraiteLesMateriaux (CodeAffaire, DateSaisie);
  TraiteLesPiecesStock (CodeAffaire, DateSaisie);
  TraiteLesPiecesHorsStock (CodeAffaire, DateSaisie);

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

Function TraiteLaSaisie : Boolean;
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
  ODExcelFile.Filter := 'fichiers Excel (*xlsm)|*.xlsm';
  Localrepert := TCbpPath.GetcegiduserDocument; // repertoire Utilisateur
  IniPgiFile := TiniFile.create (IncludeTrailingBackslash (LocalRepert)+'ACCESPGISERGADI.INI');
  ODExcelFile.InitialDir := IniPgiFile.ReadString('SERGADI','RepServer','')+'\..';
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
      Affaire := ExcelGetValue(fWorkBook, 'choix affaire','AFFAIRE');
      CodeAffaire := 'A'+Copy(Affaire,1,3)+Copy(Affaire,5,3)+Copy(Affaire,9,6)+'  00';
      DateSaisie := STrToDate(ExcelGetValue(fWorkBook, 'choix affaire','DATESAISIE'));

      if CodeAffaire <> '' then
        Result:=TraiteLesFeuilles (filename, CodeAffaire, DateSaisie)
      else
      begin
        Result:=False;
        PgiInfo ('Affaire introuvable dans le fichier choisi');
      end;

    EXCEPT
      PgiInfo ('Ouverture impossible du document Excel','');
      Result:=False;
    END;

    ExcelClose (fwinExcel);
  end;

  if Result = True then
  begin
    // Déplacement du fichier EXCEL et du fichier LOG dans le répertoire ARCHIVES de l'utilisateur
    CreateDir(IniPgiFile.ReadString('SERGADI','RepServer','') + '\ARCHIVES\');

    filenamecourt := RecupNomFic(filename);
    filenamearchive := IniPgiFile.ReadString('SERGADI','RepServer','')  + '\' + '\ARCHIVES\' + filenamecourt;
    RenameFichier(filename, filenamearchive);

    if nomlog <> '' then
    begin
      filenamecourt := RecupNomFic(nomlog);
      filenamearchive := IniPgiFile.ReadString('SERGADI','RepServer','')  + '\' + '\ARCHIVES\' + filenamecourt;
      RenameFichier(nomlog, filenamearchive);

      //Affichage fichier LOG
      FileExecOrTop('notepad.exe '+filenamearchive, False, True, False, '');
    end;
  end;

  FiniMoveProgressForm ;

end ;

procedure LanceSaisieSERGADI (Param : string) ;
var filename, LocalRepert : string;
    ODExcelFile: TOpenDialog;
    IniPGIFile : Tinifile;
begin
  if Param = 'VAL' then
	begin
    TraiteLaSaisie;
  end else if Param = 'CRE' then
	begin
    Localrepert := TCbpPath.GetcegiduserDocument; // repertoire Utilisateur
    IniPgiFile := TiniFile.create (IncludeTrailingBackslash (LocalRepert)+'ACCESPGISERGADI.INI');

    CreateDir(IniPgiFile.ReadString('SERGADI','RepServer',''));

    filename := '\\'+IniPgiFile.ReadString('SERGADI','Server','')+'\PGI01$\MODELE_SERGADI.xlsm';
    //filename := 'C:\PGI01\MODELE_SERGADI.xlsm';
	end else
  begin
    ODExcelFile := TOpenDialog.create(Application) ;
    ODExcelFile.Filter := 'fichiers Excel (*xlsm)|*.xlsm';
    Localrepert := TCbpPath.GetcegiduserDocument; // repertoire Utilisateur
    IniPgiFile := TiniFile.create (IncludeTrailingBackslash (LocalRepert)+'ACCESPGISERGADI.INI');
    ODExcelFile.InitialDir := IniPgiFile.ReadString('SERGADI','RepServer','');
    if ODExcelFile.execute then
    begin
      Filename := ODexcelFile.FileName;
    end;
    ODExcelFile.free;
  end;
  if filename <> '' then FileExecAndWait (GetExcelPath+'EXCEL.exe "'+filename+'"');
end;

end.

