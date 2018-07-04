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
     BtpUtil,
     gerepiece;

Var fWorkBook : Variant;
    TOBConso,TOBressources,TOBARTicles : TOB;
    TOBPiece,TOBLignes : TOB;
    LogFile : textfile;
    NbTraiteOk : integer;
    NbTraiteErr : integer;
    LafeuilleCourante : String;
    Nomlog : String;

Function LanceRecupDocument (NaturePiece : String) : Boolean;

Type
  TTypeErreurControle  = (TcdErrNumConso,TcdErrAffaire,TcdErrPhase,TcdErrNatureMouv,TcdErrRessource,TcdErrCodeArticle,TcdErrQte,TcdErrTypeHeure);

implementation
uses FactCommBtp,FactComm;

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
    NumConso    : double;
    Annee     : word;
    Mois      : word;
    Jour      : word;
    Semaine   : integer;
    DateMouv  : TDateTime;
    TOBA  : TOB;
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
    if (TOBD.GetValue('BCO_NATUREMOUV')<>'FAN') and (TOBA <> Nil) then TOBD.putValue('BCO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));

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
  If Montant = 0.0 then Exit;

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
    TOBD.putValue('BCO_LIBELLE','Déjà facturé - Reprise LSE');
    TOBD.putValue('BCO_QUANTITE',-1);
    TOBD.putValue('BCO_DPA',Montant);
    TOBD.putValue('BCO_DPR',Montant);
    TOBD.putValue('BCO_PUHT',Montant);
  end else if NatureMouv = 'FAN' then
  begin
    TOBD.putValue('BCO_RESSOURCE','');
    TOBD.putValue('BCO_CODEARTICLE','');
    TOBD.putValue('BCO_LIBELLE','Prorata - Reprise LSE');
    TOBD.putValue('BCO_QUANTITE',1);
    TOBD.putValue('BCO_DPA',Montant);
    TOBD.putValue('BCO_DPR',Montant);
    TOBD.putValue('BCO_PUHT',Montant);
  end else
  begin
    if Naturemouv <> 'FOU' then
      TOBD.putValue('BCO_RESSOURCE',CodeArticle)
    else
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

function BeforeGenDevis (var RC: TGerePieceRecordTobs) : boolean;

  function GetFamilleTaxe  (Typetaxe: string; tauxTva : double) : string;
  var QQ : Tquery;
      Str,Autoliquid : string;
  begin
    Autoliquid := getparamSocSecur('SO_CODETVALIQUIDST','');
    result := 'TN';
    Str := 'SELECT TV_CODETAUX FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_REGIME="FRA" AND TV_TAUXVTE='+StrfPoint(TauxTva);
    QQ := OpenSql(Str,true,-1,'',true);
    if not QQ.eof then
    begin
      QQ.first;
      while not QQ.eof do
      begin
        if (QQ.fields[0].AsString <> '') and (QQ.fields[0].AsString <> Autoliquid) then
        begin
          result := QQ.fields[0].AsString;
          break;
        end;
        QQ.next;
      end;
    end;
    ferme (QQ);
  end;

  procedure AddChampssupRatio (TOBLR : TOB);
  begin
    TOBLR.AddChampSupValeur('BASE',0);
    TOBLR.AddChampSupValeur('TAUX',0);
    TOBLR.AddChampSupValeur('MONTANT',0);
    TOBLR.AddChampSupValeur('CATEGORIE','');
    TOBLR.AddChampSupValeur('BASERG',0);
    TOBLR.AddChampSupValeur('TVARG',0);
    TOBLR.AddChampSupValeur('TTCRG',0);
  end;

var TOBLR : TOB;
    ratio : double;
    MtTvaGlob : double;
    MtHTratio,MaxHtRatio,SumMt,Diff : double;
    II,IND : integer;
begin
  result := false;
  RC.TOBrepartTva.clearDetail;
  //
  RC.TOBrepartTva.SetDouble ('BASETVATOT',0);
  RC.TOBrepartTva.SetDouble('TVATOT',0);
  RC.TOBrepartTva.SetDouble('TTCTOT',0);
  RC.TOBrepartTva.SetDouble('BASERGTOT',0);
  RC.TOBrepartTva.SetDouble('TVARGTOT',0);
  RC.TOBrepartTva.SetDouble('TTCRGTOT',0);

  MtTvaGlob := RC.TOBData.getDouble('PRECTVA20') + RC.TOBData.getDouble('PRECTVA196') + RC.TOBData.getDouble('PRECTVA10')+RC.TOBData.getDouble('PRECTVA7');
  if (MTTvaGlob = 0) then
  begin
    TOBLR := TOB.Create ('UNE REPART TVA',RC.TOBrepartTva,-1);
    AddChampssupRatio (TOBLR);
    if (RC.TobData.GetString('AUTOLIQUID')='X') then
    begin
      RC.TobData.SetDouble('TAUXTVAGLOB',0);
      TOBLR.SetString('CATEGORIE',getparamSocSecur('SO_CODETVALIQUIDST',''));
    end else
    begin
      if RC.TobData.GetDouble('TAUXTVAGLOB')=0 then RC.TobData.SetDouble('TAUXTVAGLOB',20); // protection
      TOBLR.SetString('CATEGORIE',GetFamilleTaxe('TX1',RC.TobData.GetDouble('TAUXTVAGLOB')));
    end;
    TOBLR.SetDouble('BASE',RC.TobData.GetDouble('DEJAFACTURE'));
    TOBLR.SetDouble('TAUX',RC.TobData.GetDouble('TAUXTVAGLOB'));
    TOBLR.SetDouble('MONTANT',Arrondi(RC.TobData.GetDouble('DEJAFACTURE')*RC.TobData.GetDouble('TAUXTVAGLOB')/100,V_PGI.okdecV));
    TOBLR.SetDouble('BASERG',Arrondi(TOBLR.GetDouble('BASE')*0.05,V_PGI.okdecV));
    TOBLR.SetDouble('TVARG',Arrondi(TOBLR.GetDouble('BASERG')*TOBLR.GetDouble('TAUX')/100,V_PGI.okdecV));
    TOBLR.SetDouble('TTCRG',Arrondi(TOBLR.GetDouble('BASERG')+ TOBLR.GetDouble('TVARG'),V_PGI.okdecV));
  end else
  begin
    if (RC.TOBData.GetDouble ('PRECTVA20') <> 0)  then
    begin
      TOBLR := TOB.Create ('UNE REPART TVA',RC.TOBrepartTva,-1);
      AddChampssupRatio (TOBLR);
      TOBLR.SetDouble('TAUX',20);
      TOBLR.SetDouble('MONTANT',RC.TOBData.getDouble('PRECTVA20'));
      TOBLR.SetDouble('BASE',Arrondi(TOBLR.GetDouble('MONTANT')/0.2,V_PGI.okdecV));
      TOBLR.SetString('CATEGORIE',GetFamilleTaxe('TX1',20));
      TOBLR.SetDouble('BASERG',Arrondi(TOBLR.GetDouble('BASE')*0.05,V_PGI.okdecV));
      TOBLR.SetDouble('TVARG',Arrondi(TOBLR.GetDouble('BASERG')*TOBLR.GetDouble('TAUX')/100,V_PGI.okdecV));
      TOBLR.SetDouble('TTCRG',Arrondi(TOBLR.GetDouble('BASERG')+ TOBLR.GetDouble('TVARG'),V_PGI.okdecV));
    end;
    if RC.TOBData.GetDouble ('PRECTVA196') <> 0 then
    begin
      TOBLR := TOB.Create ('UNE REPART TVA',RC.TOBrepartTva,-1);
      AddChampssupRatio (TOBLR);
      TOBLR.SetDouble('TAUX',19.6);
      TOBLR.SetDouble('MONTANT',RC.TOBData.getDouble('PRECTVA196'));
      TOBLR.SetDouble('BASE',Arrondi(TOBLR.GetDouble('MONTANT')/0.196,V_PGI.okdecV));
      TOBLR.SetString('CATEGORIE',GetFamilleTaxe('TX1',19.6));
      TOBLR.SetDouble('BASERG',Arrondi(TOBLR.GetDouble('BASE')*0.05,V_PGI.okdecV));
      TOBLR.SetDouble('TVARG',Arrondi(TOBLR.GetDouble('BASERG')*TOBLR.GetDouble('TAUX')/100,V_PGI.okdecV));
      TOBLR.SetDouble('TTCRG',Arrondi(TOBLR.GetDouble('BASERG')+ TOBLR.GetDouble('TVARG'),V_PGI.okdecV));
    end;
    if RC.TOBData.GetDouble ('PRECTVA10') <> 0 then
    begin
      TOBLR := TOB.Create ('UNE REPART TVA',RC.TOBrepartTva,-1);
      AddChampssupRatio (TOBLR);
      TOBLR.SetDouble('TAUX',10);
      TOBLR.SetDouble('MONTANT',RC.TOBData.getDouble('PRECTVA10'));
      TOBLR.SetDouble('BASE',Arrondi(TOBLR.GetDouble('MONTANT')/0.1,V_PGI.okdecV));
      TOBLR.SetString('CATEGORIE',GetFamilleTaxe('TX1',10));
      TOBLR.SetDouble('BASERG',Arrondi(TOBLR.GetDouble('BASE')*0.05,V_PGI.okdecV));
      TOBLR.SetDouble('TVARG',Arrondi(TOBLR.GetDouble('BASERG')*TOBLR.GetDouble('TAUX')/100,V_PGI.okdecV));
      TOBLR.SetDouble('TTCRG',Arrondi(TOBLR.GetDouble('BASERG')+ TOBLR.GetDouble('TVARG'),V_PGI.okdecV));
    end;
    if RC.TOBData.GetDouble ('PRECTVA7') <> 0 then
    begin
      TOBLR := TOB.Create ('UNE REPART TVA',RC.TOBrepartTva,-1);
      AddChampssupRatio (TOBLR);
      TOBLR.SetDouble('TAUX',7);
      TOBLR.SetDouble('MONTANT',RC.TOBData.getDouble('PRECTVA7'));
      TOBLR.SetDouble('BASE',Arrondi(TOBLR.GetDouble('MONTANT')/0.07,V_PGI.okdecV));
      TOBLR.SetString('CATEGORIE',GetFamilleTaxe('TX1',7));
      TOBLR.SetDouble('BASERG',Arrondi(TOBLR.GetDouble('BASE')*0.05,V_PGI.okdecV));
      TOBLR.SetDouble('TVARG',Arrondi(TOBLR.GetDouble('BASERG')*TOBLR.GetDouble('TAUX')/100,V_PGI.okdecV));
      TOBLR.SetDouble('TTCRG',Arrondi(TOBLR.GetDouble('BASERG')+ TOBLR.GetDouble('TVARG'),V_PGI.okdecV));
    end;
  end;
  MaxHtRatio := 0;
  SumMt := 0;
  IND := -1;

  for II := 0 to RC.TOBrepartTva.detail.count -1 do
  begin
    TOBLR :=  RC.TOBrepartTva.detail[II];
    if abs(TOBLR.GetDouble('BASE'))> abs(Maxhtratio) then
    begin
      IND := II;
      MaxHTRatio := TOBLR.GetDouble('BASE');
    end;
    SumMt := Summt + TOBLR.GetDouble('BASE');
  end;

  if SumMt <> RC.TobData.GetDouble('DEJAFACTURE') then  // reajustement du montant HT par rapport au deja facture
  begin
    TOBLR :=  RC.TOBrepartTva.detail[IND];
    TOBLR.SetDouble('BASE',TOBLR.GetDouble('BASE')+ (RC.TobData.GetDouble('DEJAFACTURE')-Summt));
  end;

  for II := 0 to RC.TOBrepartTva.detail.count -1 do
  begin
    TOBLR :=  RC.TOBrepartTva.detail[II];
    //
    RC.TOBrepartTva.SetDouble('BASETVATOT',RC.TOBrepartTva.GetDouble('BASETVATOT')+ TOBLR.GetDouble('BASE'));
    RC.TOBrepartTva.SetDouble('TVATOT',RC.TOBrepartTva.GetDouble('TVATOT')+ TOBLR.GetDouble('MONTANT'));
    RC.TOBrepartTva.SetDouble('TTCTOT',RC.TOBrepartTva.GetDouble('BASETVATOT')+RC.TOBrepartTva.GetDouble('TVATOT'));
    //
    RC.TOBrepartTva.SetDouble('BASERGTOT',RC.TOBrepartTva.GetDouble('BASERGTOT')+ TOBLR.GetDouble('BASERG'));
    RC.TOBrepartTva.SetDouble('TVARGTOT',RC.TOBrepartTva.GetDouble('TVARGTOT')+ TOBLR.GetDouble('TVARG'));
    RC.TOBrepartTva.SetDouble('TTCRGTOT',RC.TOBrepartTva.GetDouble('TTCRGTOT')+TOBLR.GetDouble('TTCRG'));
  end;
  //
  SumMt := RC.TOBData.GetDouble('VALEURCAUTION')+ RC.TOBData.GetDouble('CUMULRG');
  // regulation des RG
  if RC.TOBrepartTva.GetDouble('TTCRGTOT') <> SumMt then
  begin
    TOBLR :=  RC.TOBrepartTva.detail[IND];
    Diff := SumMt-RC.TOBrepartTva.GetDouble('TTCRGTOT');
    RC.TOBrepartTva.SetDouble('TTCRGTOT',RC.TOBrepartTva.GetDouble('TTCRGTOT')+ diff);
    TOBLR.SetDouble('TTCRG',TOBLR.GetDouble('TTCRG')+ diff);
    // Et retour arriere on reclacule une base depuis le montant -- cas de figure ou la rg est arrivé apres les premierzes facturations
    TOBLR.setDouble('BASERG', Arrondi(TOBLR.GetDouble('TTCRG') / (1+(TOBLR.GetDouble('TAUX')/100)),V_PGI.okdecV));
    TOBLR.setDouble('TVARG', Arrondi(TOBLR.GetDouble('TTCRG') - TOBLR.GetDouble('BASERG'),V_PGI.okdecV));
    //
    RC.TOBrepartTva.SetDouble('BASERGTOT',0);
    RC.TOBrepartTva.SetDouble('TVARGTOT',0);
    RC.TOBrepartTva.SetDouble('TTCRGTOT',0);
    for II := 0 to RC.TOBrepartTva.detail.count -1 do
    begin
      TOBLR :=  RC.TOBrepartTva.detail[II];
      //
      RC.TOBrepartTva.SetDouble('BASERGTOT',RC.TOBrepartTva.GetDouble('BASERGTOT')+ TOBLR.GetDouble('BASERG'));
      RC.TOBrepartTva.SetDouble('TVARGTOT',RC.TOBrepartTva.GetDouble('TVARGTOT')+ TOBLR.GetDouble('TVARG'));
      RC.TOBrepartTva.SetDouble('TTCRGTOT',RC.TOBrepartTva.GetDouble('TTCRGTOT')+TOBLR.GetDouble('TTCRG'));
    end;
  end;

  result := true;
end;

function BeforevalideLeDevis (var RC: TGerePieceRecordTobs) : boolean;
var TOBSS,TOBLF,TOBL,TOBPRG,TOBLRG,TOBLR : TOB;
    II : integer;
begin
  result := false;
// ici on va préparer l'enregistrement de BSITUATIONS, LIGNEFAC et des PIEDBASRG,PIECERG pour la situation cumulative précédente
  if RC.TobData.GetInteger('NUMSITUATION') = 0 then RC.TobData.SetInteger('NUMSITUATION',1); // protection au cas ou...

  TOBL := RC.TobPiece.findfirst(['GL_TYPELIGNE'],['ART'],True);
  if TOBL = nil then exit;

  TOBLRG := RC.TobPiece.findfirst(['GL_TYPELIGNE'],['RG'],True);

  // Bsituation de la situation
  TOBSS := TOB.Create('BSITUATIONS',RC.TOBSITUATIONS,-1);
  TOBSS.SetString('BST_NATUREPIECE','FBT');
  TOBSS.SetString('BST_SOUCHE','XXX');
  TOBSS.SetString('BST_NUMEROFAC',RC.TobData.getString('NUMFACSIT'));
  TOBSS.SetString('BST_DATESIT',RC.TobData.getString('DATEDERSIT'));
  TOBSS.SetString('BST_AFFAIRE',RC.TobPiece.GetString('GP_AFFAIRE'));
  TOBSS.SetDouble('BST_MONTANTHT',RC.TobData.GetDouble('DEJAFACTURE'));
  TOBSS.SetDouble('BST_MONTANTTTC',TOBSS.getDouble('BST_MONTANTHT')+RC.TOBrepartTva.GetDouble('TVATOT'));
  TOBSS.SetDouble('BST_MONTANTTVA',RC.TOBrepartTva.GetDouble('TVATOT'));
  TOBSS.SetString('BST_VIVANTE','X');
  TOBSS.SetInteger('BST_INDICESIT',0);
  TOBSS.SetInteger('BST_NUMEROSIT',RC.TobData.GetInteger('NUMSITUATION'));

  // le lignefac de la situation
  TOBLF := TOB.Create ('LIGNEFAC',RC.TOBLigneFac,-1);
  TOBLF.SetString('BLF_NATUREPIECEG','FBT');
  TOBLF.SetString('BLF_SOUCHE','XXX');
  TOBLF.SetString('BLF_NUMERO',RC.TobData.getString('NUMFACSIT'));
  TOBLF.SetInteger('BLF_INDICEG',0);
  TOBLF.SetInteger('BLF_NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
  TOBLF.SetDouble('BLF_MTMARCHE',RC.TOBData.GetDouble('MONTANT'));
  TOBLF.SetDouble('BLF_MTCUMULEFACT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_MTSITUATION',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_QTEMARCHE',RC.TOBData.GetDouble('MONTANT'));
  TOBLF.SetDouble('BLF_QTECUMULEFACT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_QTESITUATION',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_TOTALTTCDEV',RC.TOBrepartTva.GetDouble('TTCTOT'));
  TOBLF.SetDouble('BLF_MTPRODUCTION',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_POURCENTAVANC',Arrondi(RC.TOBData.GetDouble('DEJAFACTURE')/RC.TOBData.GetDouble('MONTANT')*100,2));

  for II := 0 to RC.TOBrepartTva.detail.Count -1 do
  begin
    TOBLR := RC.TOBrepartTva.detail[II];
    TOBLF := TOB.Create ('PIEDBASE',RC.TobPiedBaseP,-1);
    TOBLF.SetString('GPB_NATUREPIECEG','FBT');
    TOBLF.SetString('GPB_SOUCHE','XXX');
    TOBLF.SetString('GPB_NUMERO',RC.TobData.getString('NUMFACSIT'));
    TOBLF.SetInteger('GPB_INDICEG',0);
    TOBLF.SetString('GPB_CATEGORIETAXE','TX1');
    TOBLF.SetString('GPB_FAMILLETAXE',TOBLR.GetString('CATEGORIE'));
    TOBLF.SetDouble('GPB_TAUXTAXE',TOBLR.GetDouble('TAUX'));
    TOBLF.SetDouble('GPB_BASETAXE',TOBLR.GetDouble('BASE'));
    TOBLF.SetDouble('GPB_VALEURTAXE',TOBLR.GetDouble('MONTANT'));
    TOBLF.SetDouble('GPB_BASEDEV',TOBLF.getDouble('GPB_BASETAXE'));
    TOBLF.SetDouble('GPB_VALEURDEV',TOBLF.GetDouble('GPB_VALEURTAXE'));
    TOBLF.SetDouble('GPB_TAUXDEV',1);
    TOBLF.SetDouble('GPB_COTATION',1);
    TOBLF.SetString('GPB_DATETAUXDEV',RC.TobData.getString('DATEDERSIT'));
    TOBLF.SetString('GPB_DEVISE','EUR');
    TOBLF.SetString('GPB_SAISIECONTRE','-');
  end;
  if TOBLRG <> nil then
  begin
    // calcul de la tva associée a la RG sur chaque base
    for II := 0 to RC.TOBrepartTva.detail.Count -1 do
    begin
      TOBLR := RC.TOBrepartTva.detail[II];
      // le baseRG de la situation
      TOBPRG := TOB.Create('PIEDBASERG',RC.TObbasesRGP,-1);
      TOBPRG.SetString('PBR_NATUREPIECEG','FBT');
      TOBPRG.SetString('PBR_SOUCHE','XXX');
      TOBPRG.SetString('PBR_DATEPIECE',RC.TobData.getString('DATEDERSIT'));
      TOBPRG.SetInteger('PBR_NUMERO',RC.TobData.GetInteger('NUMFACSIT'));
      TOBPRG.SetInteger('PBR_INDICEG',0);
      TOBPRG.SetInteger('PBR_NUMLIGNE',TOBLRG.GetInteger('GL_NUMORDRE'));
      TOBPRG.SetString('PBR_CATEGORIETAXE','TX1');
      TOBPRG.SetString('PBR_FAMILLETAXE',TOBLR.GetString('CATEGORIE'));
      TOBPRG.SetDouble('PBR_TAUXTAXE',TOBLR.GetDouble('TAUX'));
      TOBPRG.SetDouble('PBR_BASETAXE',TOBLR.GetDouble('BASERG'));
      TOBPRG.SetDouble('PBR_BASEDEV',TOBPRG.GetDouble('PBR_BASETAXE'));
      TOBPRG.SetDouble('PBR_VALEURTAXE',TOBLR.GetDouble('BASERG')+TOBLR.GetDouble('TVARG'));
      TOBPRG.SetDouble('PBR_VALEURDEV',TOBPRG.GetDouble('PBR_VALEURTAXE'));
      TOBPRG.SetDouble('PBR_TAUXDEV',1);
      TOBPRG.SetString('PBR_DATETAUXDEV',RC.TobData.getString('DATEDERSIT'));
      TOBPRG.SetString('PBR_DEVISE','EUR');
    end;
    // le piecerg de la situation
    TOBPRG := TOB.Create('PIECERG',RC.TOBPieceRGP,-1);
    TOBPRG.SetString('PRG_NATUREPIECEG','FBT');
    TOBPRG.SetString('PRG_SOUCHE','XXX');
    TOBPRG.SetString('PRG_DATEPIECE',RC.TobData.getString('DATEDERSIT'));
    TOBPRG.SetInteger('PRG_NUMERO',RC.TobData.GetInteger('NUMFACSIT'));
    TOBPRG.SetInteger('PRG_INDICEG',0);
    TOBPRG.SetInteger('PRG_NUMLIGNE',TOBLRG.GetInteger('GL_NUMORDRE'));
    TOBPRG.SetString('PRG_TYPERG','TTC');
    TOBPRG.SetDouble('PRG_TAUXRG',5);
    TOBPRG.SetDouble('PRG_MTHTRG',TOBLRG.GetDouble('MONTANT'));
    TOBPRG.SetDouble('PRG_MTHTRGDEV',TOBPRG.GetDouble('PRG_MTHTRG'));
    TOBPRG.SetDouble('PRG_MTTTCRG',RC.TOBrepartTva.GetDouble('TTCRGTOT'));
    TOBPRG.SetDouble('PRG_MTTTCRGDEV',TOBPRG.GetDouble('PRG_MTTTCRG'));
    TOBPRG.SetString('PRG_NUMCAUTION',RC.TOBData.GetString('NUMCAUTION'));
    TOBPRG.SetString('PRG_BANQUECP',RC.TOBData.GetString('BANQUECP'));
    TOBPRG.SetDouble('PRG_CAUTIONMT',RC.TOBData.GetDouble('VALEURCAUTION'));
    TOBPRG.SetDouble('PRG_CAUTIONMTDEV',RC.TOBData.GetDouble('VALEURCAUTION'));
    TOBPRG.SetString('PRG_APPLICABLE','X');

  end;
  result := true;
end;

function AfterValideLeDevis (var RC: TGerePieceRecordTobs) : boolean;

  function CalculeHtFromTTC (Montant : double; CodeTaxe : string) : double;
  var QQ : Tquery;
      Taux : double;
  begin
    Taux := 1;
    QQ := OpenSql('SELECT TV_TAUXVTE FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="'+CodeTaxe+'" AND TV_REGIME="FRA"',true,1,'',true);
    if not QQ.eof then
    begin
      Taux := QQ.fields[0].asFloat / 100;
    end;
    ferme (QQ);
    result := Arrondi(Montant / (1+Taux),V_PGI.okdecV);
  end;

var TOBSS,TOBL,TOBLF,TOBMF,TOBLRG,TOBLL,TOBPRG,TOBPP : TOB;
    RefDevis : string;
    QQ : TQuery;
begin
  result := false;
  TOBL := RC.TobPiece.findfirst(['GL_TYPELIGNE'],['ART'],True);
  if TOBL = nil then exit;
  //
  TOBLRG := RC.TobPiece.findfirst(['GL_TYPELIGNE'],['RG'],True);
  // le lignefac du devis
  TOBLF := TOB.Create ('LIGNEFAC',RC.TOBLigneFac,-1);
  TOBLF.SetString('BLF_NATUREPIECEG',RC.TOBpiece.getString('GP_NATUREPIECEG'));
  TOBLF.SetString('BLF_SOUCHE',RC.TOBpiece.getString('GP_SOUCHE'));
  TOBLF.SetInteger('BLF_NUMERO',RC.TOBpiece.getInteger('GP_NUMERO'));
  TOBLF.SetInteger('BLF_INDICEG',0);
  TOBLF.SetInteger('BLF_NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
  TOBLF.SetDouble('BLF_MTMARCHE',RC.TOBData.GetDouble('MONTANT'));
  TOBLF.SetDouble('BLF_MTDEJAFACT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_MTCUMULEFACT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_QTEMARCHE',RC.TOBData.GetDouble('MONTANT'));
  TOBLF.SetDouble('BLF_QTEDEJAFACT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_QTECUMULEFACT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBLF.SetDouble('BLF_TOTALTTCDEV',Arrondi(RC.TOBData.GetDouble('DEJAFACTURE')*1.2,V_PGI.okdecV));
  TOBLF.SetDouble('BLF_POURCENTAVANC',Arrondi(RC.TOBData.GetDouble('DEJAFACTURE')/RC.TOBData.GetDouble('MONTANT')*100,2));
  //
  TOBL.SetDouble('GL_QTESIT',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBL.SetDouble('GL_QTEPREVAVANC',RC.TOBData.GetDouble('DEJAFACTURE'));
  TOBL.SetDouble('GL_POURCENTAVANC',TOBLF.GetDouble('BLF_POURCENTAVANC'));
  TOBL.UpdateDB;

  RefDevis  := EncoderefSel(TOBl);
  TOBMF := TOB.Create ('BTMEMOFACTURE',nil,-1);
  TOBMF.SetString ('BMF_DEVISPRINC',RefDevis);
  TOBMF.SetInteger ('BMF_NUMORDRE',1);
  TOBMF.SetString ('BMF_DEVIS',RefDevis);
  TOBMF.SetString ('BMF_AFFAIRE',RC.TobPiece.GetString('GP_AFFAIRE'));
  TOBMF.SetString ('BMF_AFFAIRE1',RC.TobPiece.GetString('GP_AFFAIRE1'));
  TOBMF.SetString ('BMF_AFFAIRE2',RC.TobPiece.GetString('GP_AFFAIRE2'));
  TOBMF.SetString ('BMF_AFFAIRE3',RC.TobPiece.GetString('GP_AFFAIRE3'));
  TOBMF.SetString ('BMF_NATUREPIECEG',TOBL.GetString('GL_NATUREPIECEG'));
  TOBMF.SetString ('BMF_SOUCHE',TOBL.GetString('GL_SOUCHE'));
  TOBMF.SetInteger ('BMF_NUMERO',TOBL.GetInteger('GL_NUMERO'));
  TOBMF.SetInteger ('BMF_INDICEG',TOBL.GetInteger('GL_INDICEG'));
  TOBMF.SetAllModifie(true);
  TOBMF.InsertDB(nil);
  TOBMF.free;

  if TOBLRG <> nil then
  begin
    TOBLL := TOB.Create('LIGNE',nil,-1);
    TOBLL.SetString('GL_NATUREPIECEG','FBT');
    TOBLL.SetString('GL_SOUCHE','XXX');
    TOBLL.SetString('GL_DATEPIECE',RC.TobData.getString('DATEDERSIT'));
    TOBLL.SetInteger('GL_NUMERO',RC.TobData.GetInteger('NUMFACSIT'));
    TOBLL.SetInteger('GL_INDICEG',0);
    TOBLL.SetInteger('GL_NUMORDRE',TOBLRG.GetInteger('GL_NUMORDRE'));
    TOBLL.SetInteger('GL_NUMLIGNE',TOBLRG.GetInteger('GL_NUMORDRE'));
    TOBLL.SetString('GL_PIECEPRECEDENTE',EncodeRefPiece(TOBLRG));
    TOBLL.SetString('GL_PIECEORIGINE',EncodeRefPiece(TOBLRG));
    TOBLL.SetAllModifie(true); 
    TOBLL.InsertDB(nil);
    TOBLL.Free; 
  end;
  
  if (RC.TobData.getString('CODEGBF')<>'') and (RC.TobData.getdouble('CUMULRBF')<>0) then
  begin
    TOBPP := TOB.Create('PORT',nil,-1);
    TOBLL := TOB.Create('PIEDPORT',nil,-1);
    QQ := OpenSql ('SELECT * FROM PORT WHERE GPO_CODEPORT="'+ RC.TobData.GetString('CODEGBF')+'"',true,1,'',true);
    if not QQ.eof then
    begin
      TOBPP.selectdb('',QQ);
      TOBLL.SetString('GPT_NATUREPIECEG','FBT');
      TOBLL.SetString('GPT_SOUCHE','XXX');
      TOBLL.SetInteger('GPT_NUMERO',RC.TobData.GetInteger('NUMFACSIT'));
      TOBLL.SetInteger('GPT_INDICEG',0);
      TOBLL.SetInteger('GPT_NUMPORT',1);
      TOBLL.SetString('GPT_CODEPORT',RC.TobData.GetString('CODEGBF'));
      TOBLL.SetString('GPT_LIBELLE',TOBPP.GetString('GPO_LIBELLE'));
      TOBLL.SetString('GPT_TYPEPORT',TOBPP.GetString('GPO_TYPEPORT'));
      TOBLL.PutValue('GPT_FAMILLETAXE1', TOBPP.GetValue('GPO_FAMILLETAXE1'));
      TOBLL.PutValue('GPT_COMPTAARTICLE', TOBPP.GetValue('GPO_COMPTAARTICLE'));
      TOBLL.PutValue('GPT_TOTALTTCDEV', RC.TobData.getdouble('CUMULRBF'));
      TOBLL.PutValue('GPT_TOTALHTDEV', CalculeHtFromTTC(TOBLL.GetValue('GPT_TOTALTTCDEV'),TOBLL.GetValue('GPT_FAMILLETAXE1')));
      TOBLL.PutValue('GPT_TOTALTAXE1', TOBLL.GetValue('GPT_TOTALTTCDEV')-TOBLL.GetValue('GPT_TOTALHTDEV'));
      TOBLL.SetString('GPT_RETENUEDIVERSE','X');
      TOBLL.SetAllModifie(true);
      TOBLL.InsertDB(nil);
    end;
    ferme (QQ);
    TOBLL.Free;
    TOBPP.Free;

  end;


  TOBSS := RC.TOBSITUATIONS.detail[0];
  if TOBSS <> nil then
  begin
    TOBSS.SetString('BST_SSAFFAIRE',RC.TobPiece.GetString('GP_AFFAIREDEVIS'));
    TOBSS.SetAllModifie(true);
    TOBSS.InsertDB(nil);
  end;
  
  if RC.TOBLigneFac.detail.count > 0 then
  begin
    RC.TOBLigneFac.SetAllModifie(true);
    RC.TOBLigneFac.InsertDB(nil);
  end;

  if RC.TobPiedBaseP.detail.count > 0 then
  begin
    RC.TobPiedBaseP.SetAllModifie(true);
    RC.TobPiedBaseP.InsertDB(nil);
  end;

  if RC.TOBPieceRGP.detail.count > 0 then
  begin
    TOBPRG := RC.TOBPieceRGP.detail[0];
    TOBPRG.SetInteger('PRG_NUMLIGNE',TOBLRG.GetInteger('GL_NUMORDRE'));
    RC.TOBPieceRGP.SetAllModifie(true);
    RC.TOBPieceRGP.InsertDB(nil);
  end;

  if RC.TObbasesRGP.detail.count > 0 then
  begin
    RC.TObbasesRGP.SetAllModifie(true);
    RC.TObbasesRGP.InsertDB(nil);
  end;
  result := true;

end;


Function TraiteLesDevis : Boolean;
Var i, lig, col, nbcol : integer;
    Chaine, CodeAffaire, Article,  Unite, Datepiece : String;
    Qte, Montant : Double;
begin
//      PgiInfo ('Traitement Devis','');

  Result := True;
  LaFeuilleCourante := 'Feuil1';
  For i:=2 to 1000 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Chaine = '' then Exit;

    CodeAffaire := 'A'+Copy(Chaine,1,9)+'     00';

    // Création pour entete piece
    TOBPiece := Tob.Create ('', nil, -1);

    TobPiece.AddChampSupValeur ('NATUREPIECEG', 'DBT');
    TobPiece.AddChampSupValeur ('AFFAIRE', CodeAffaire);
    TobPiece.AddChampSupValeur ('TIERS', getTiersAffaire(CodeAffaire));
    TobPiece.AddChampSupValeur ('ETABLISSEMENT', GetParamSocSecur('SO_ETABLISDEFAUT','001'));
    TobPiece.AddChampSupValeur ('DOMAINE', '');
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'B'+intToStr(i));
    if (Chaine <> '') and (IsvalidDate(Chaine)) then TobPiece.AddChampSupValeur ('DATEPIECE', StrToDate(Chaine))
                                                else TobPiece.AddChampSupValeur ('DATEPIECE', StrToDate('01/01/2015'));
    DatePiece := TobPiece.Getvalue ('DATEPIECE');
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'D'+intToStr(i));
    TobPiece.AddChampSupValeur ('REFINTERNE', chaine + ' - Reprise LSE');
    TobPiece.AddChampSupValeur ('ETATAFFAIRE', 'ACP');

    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'E'+intToStr(i));
    Montant := valeur(Chaine);
    TobPiece.AddChampSupValeur ('MONTANT', valeur(Chaine));
    if Montant = 0.0 then Continue;

    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'F'+intToStr(i));
    TobPiece.AddChampSupValeur ('DEJAFACTURE', valeur(Chaine));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'H'+intToStr(i));
    TobPiece.AddChampSupValeur ('CUMULRG', valeur(Chaine));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'I'+intToStr(i));
    TobPiece.AddChampSupValeur ('CUMULRBF', valeur(Chaine));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'J'+intToStr(i)); if chaine = '' then chaine := '0';
    TobPiece.AddChampSupValeur ('VALEURCAUTION', valeur(Chaine)); if chaine = '' then chaine := '0';
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'K'+intToStr(i));
    TobPiece.AddChampSupValeur ('NUMCAUTION', Chaine);
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'L'+intToStr(i)); if chaine = '' then chaine := '0';
    TobPiece.AddChampSupValeur ('NUMSITUATION', StrToInt(Chaine));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'M'+intToStr(i));
    if (Chaine <> '') and (IsvalidDate(Chaine)) then TobPiece.AddChampSupValeur ('DATEDERSIT', StrToDate(Chaine))
                                                else TobPiece.AddChampSupValeur ('DATEDERSIT', StrToDate('31/12/2015'));
    TOBPiece.AddChampSupValeur ('NUMFACSIT', Format('2015%.04d',[I-1]));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'N'+intToStr(i));
    TobPiece.AddChampSupValeur ('BANQUECP', Chaine);
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'O'+intToStr(i));
    TobPiece.AddChampSupValeur ('CODEGBF', Chaine); // ocde garantie bonne fin
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'P'+intToStr(i)); if chaine = '' then chaine := '0';
    TobPiece.AddChampSupValeur ('PRECTVA20', arrondi(valeur(Chaine),V_PGI.okdecV)); // mont de Tva 20 % deja facturé
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'Q'+intToStr(i)); if chaine = '' then chaine := '0';
    TobPiece.AddChampSupValeur ('PRECTVA196', Arrondi(valeur(Chaine),V_PGI.okdecV)); // mont de Tva 19.6 % deja facturé
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'R'+intToStr(i)); if chaine = '' then chaine := '0';
    TobPiece.AddChampSupValeur ('PRECTVA10', Arrondi(valeur(Chaine),V_PGI.okdecV)); // mont de Tva 10 % deja facturé
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'S'+intToStr(i)); if chaine = '' then chaine := '0';
    TobPiece.AddChampSupValeur ('PRECTVA7', Arrondi(valeur(Chaine),V_PGI.okdecV)); // mont de Tva 7 % deja facturé
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'T'+intToStr(i)); if chaine = '' then chaine := '0';  
    TobPiece.AddChampSupValeur ('TAUXTVAGLOB', Arrondi(valeur(Chaine)*100,0)); // Taux TVA GLOBALE
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'U'+intToStr(i));
    if Chaine = 'X' then TobPiece.AddChampSupValeur ('AUTOLIQUID', 'X')
                    else TobPiece.AddChampSupValeur ('AUTOLIQUID', '-');
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'V'+intToStr(i));
    TobPiece.AddChampSupValeur ('TIERSFAC', Chaine); // tiers a facturer

    // Création pour maj lignes de pièce : total devis
    Article := 'REP_DEVIS';
    Qte := 1.0;
    Unite := '';

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Montant);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DATELIVRAISON', DatePiece);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', Qte);
    TobLignes.AddChampSupValeur ('PUHT', Qte);
    TobLignes.AddChampSupValeur ('DPA', 0.0);
    TobLignes.AddChampSupValeur ('DPR', 0.0);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');

    if (TOBPiece.GetDouble('CUMULRG')<> 0) or (TOBPiece.GetDouble('VALEURCAUTION')<> 0)  then
    begin
      TobPiece.AddChampSupValeur ('BASERG', Montant);
      TobLignes := Tob.Create ('', TobPiece, -1);
      TobLignes.AddChampSupValeur ('TYPELIGNE', 'RG');
      TobLignes.AddChampSupValeur ('CODEARTICLE', '');
      TobLignes.AddChampSupValeur ('ARTICLE', '');
      TobLignes.AddChampSupValeur ('LIBELLE', 'Retenue de garantie TTC 5%');
      TobLignes.AddChampSupValeur ('QTEFACT', 1);
      TobLignes.AddChampSupValeur ('QUALIFQTE', '');
      TobLignes.AddChampSupValeur ('DATELIVRAISON', '01/01/1900');
      TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
      TobLignes.AddChampSupValeur ('PUHTDEV', Arrondi(Montant * 0.05,V_PGI.okdecV));
      TobLignes.AddChampSupValeur ('PUHT', Arrondi(Montant * 0.05,V_PGI.okdecV));
      TobLignes.AddChampSupValeur ('DPA', 0.0);
      TobLignes.AddChampSupValeur ('DPR', 0.0);
      TobLignes.AddChampSupValeur ('AVECPRIX', 'X');
    end;

    if (TOBPiece <> nil) and (TOBPiece.Detail.count > 0) then
    begin
      CreatePieceFromTob (TobPiece,BeforevalideLeDevis,AfterValideLeDevis,BeforeGenDevis,'ACP');
    end;

    TOBPiece.free;
    TOBPiece := nil;
  end;
end;

Function TraiteLesCF : Boolean;
Var i : integer;
    Chaine, CodeAffaire, Article, Unite, Datepiece, DateLivraison : String;
    Qte, Montant : Double;
begin
//      PgiInfo ('Traitement CF','');

  Result := True;
  LaFeuilleCourante := 'Feuil1';
  For i:=2 to 1000 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Chaine = '' then Exit;

    CodeAffaire := 'A'+Copy(Chaine,1,9)+'     00';

    // Création pour entete piece
    TOBPiece := Tob.Create ('', nil, -1);

    TobPiece.AddChampSupValeur ('NATUREPIECEG', 'CF');
    TobPiece.AddChampSupValeur ('AFFAIRE', CodeAffaire);
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'E'+intToStr(i));
    TobPiece.AddChampSupValeur ('TIERS', chaine);
    TobPiece.AddChampSupValeur ('ETABLISSEMENT', GetParamSocSecur('SO_ETABLISDEFAUT','001'));
    TobPiece.AddChampSupValeur ('DOMAINE', '');
    DatePiece := ExcelGetValue(fWorkBook, LaFeuilleCourante,'B'+intToStr(i));
    TobPiece.AddChampSupValeur ('DATEPIECE', StrToDate(DatePiece));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'C'+intToStr(i));
    TobPiece.AddChampSupValeur ('REFINTERNE', chaine + ' - Reprise LSE');
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'H'+intToStr(i));
    TobPiece.AddChampSupValeur ('REFEXTERNE', chaine);
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'J'+intToStr(i));
    TobPiece.AddChampSupValeur ('RESSOURCE', chaine);
    DateLivraison := ExcelGetValue(fWorkBook, LaFeuilleCourante,'L'+intToStr(i));
    if DateLivraison = '' then DateLivraison :=  DatePiece;
    TobPiece.AddChampSupValeur ('DATELIVRAISON', StrToDate(DateLivraison));
    chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'K'+intToStr(i));
    if UPPERCASE(Chaine) = 'X' then TobPiece.AddChampSupValeur ('AUTOLIQUID', 'X')
                    else TobPiece.AddChampSupValeur ('AUTOLIQUID', '-');

    // Création pour maj lignes de pièce : total devis
    Article := 'REP_' + ExcelGetValue(fWorkBook, LaFeuilleCourante,'G'+intToStr(i));;
    Article := UpperCase( Article);
    Montant := 1.0;
    Unite := '';
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'F'+intToStr(i));
    Qte := StrToFloat(Chaine);

    TobLignes := Tob.Create ('', TobPiece, -1);

    TobLignes.AddChampSupValeur ('TYPELIGNE', 'ART');
    TobLignes.AddChampSupValeur ('CODEARTICLE', Article);
    TobLignes.AddChampSupValeur ('ARTICLE', CodeArticleUnique (Article,'','','','',''));
    TobLignes.AddChampSupValeur ('LIBELLE', GetLibelleMateriaux(TOBLignes.GetValue('ARTICLE')));
    TobLignes.AddChampSupValeur ('QTEFACT', Qte);
    TobLignes.AddChampSupValeur ('QUALIFQTE', Unite);
    TobLignes.AddChampSupValeur ('DEPOT', GetParamsocsecur('SO_GCDEPOTDEFAUT','001'));
    TobLignes.AddChampSupValeur ('PUHTDEV', Montant);
    TobLignes.AddChampSupValeur ('PUHT', Montant);
    TobLignes.AddChampSupValeur ('DPA', Montant);
    TobLignes.AddChampSupValeur ('DPR', Montant);
    TobLignes.AddChampSupValeur ('AVECPRIX', 'X');
    TobLignes.AddChampSupValeur ('DATELIVRAISON', StrToDate(DateLivraison));

    if (TOBPiece <> nil) and (TOBPiece.Detail.count > 0) then
    begin
      CreatePieceFromTob (TobPiece);
    end;

    TOBPiece.free;
    TOBPiece := nil;
  end;
end;

Function TraiteLesConsos : Boolean;
Var i : integer;
    Chaine, CodeAffaire : String;
    Montant : Double;
begin
//      PgiInfo ('Traitement Cumule A date','');
  Result := True;
  LaFeuilleCourante := 'Feuil1';
  For i:=3 to 1000 Do
  begin
    MoveCurProgressForm('Traitement '+LafeuilleCourante+' en cours ...');

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'A'+intToStr(i));
    if Chaine = '' then Exit;

    CodeAffaire := 'A'+Copy(Chaine,1,9)+'     00';

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'B'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_BEENT', 'MO', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'C'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_BEINT', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'D'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_BEST', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'E'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_MOOUV', 'MO', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'F'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_MOTEC', 'MO', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'G'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_MOINT', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'H'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_MOST', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'I'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_SSTAC', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'J'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_ACHFM', 'FOU', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'K'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, 'REP_FRDIV', 'EXT', Montant, i);

    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'L'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, '', 'FAN', Montant, i);
    //
    Chaine := ExcelGetValue(fWorkBook, LaFeuilleCourante,'N'+intToStr(i));
    Montant := StrToFloat(Chaine);
    // Traitement de la ligne
    ChargeTOBConsos(V_PGI.DateEntree, CodeAffaire, '', 'RAN', Montant, i);

    end;
end;

Function TraiteLesDonnees (filename, Typ : String) : Boolean;
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

  if Typ = 'DBT' then
    TraiteLesDevis
  else if Typ = 'CF' then
    TraiteLesCF
  else if Typ = 'CONSOS' then
    TraiteLesConsos;

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

Function LanceRecupDocument (NaturePiece : String) : Boolean;
var filename : string;
    ODExcelFile: TOpenDialog;
    fWinExcel: OleVariant;
    fnewInst : boolean;
begin
  Result := True;

  InitMoveProgressForm(Application.MainForm,Application.Title, 'Traitement en cours, veuillez patienter...', 200, False, true) ;

  ODExcelFile := TOpenDialog.create(nil) ;
  ODExcelFile.Filter := 'fichiers Excel (*xls;*xlsx)|*.xls;*.xlsx';
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
      Result:=TraiteLesDonnees (filename, NaturePiece);
    EXCEPT
      PgiInfo ('Traitement impossible du document Excel','');
      Result:=False;
    END;

    ExcelClose (fwinExcel);
  end;

  FiniMoveProgressForm ;

end ;

end.

