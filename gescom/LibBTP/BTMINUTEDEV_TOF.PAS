{***********UNITE*************************************************
Auteur  ...... :  LS
Créé le ...... : 11/02/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTMINUTEDEV ()
Mots clefs ... : TOF;BTMINUTEDEV
*****************************************************************}
Unit BTMINUTEDEV_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul, MainEagl,
{$ENDIF}
     uTob,
     forms,
     Aglinit,
     M3FP,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     SaisUtil,
     stat,
     UTOF ;

Type
  TOF_BTMINUTEDEV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	//
  end ;


procedure LanceMinuteDevis(TOBPiece,TOBTiers,TOBOuvrages,TOBOuvragesP,TOBARTicles : TOB; DEV : RDEVISE);

Implementation
uses FactureBTP,FactUtil,factOuvrage,ParamSoc,Factcalc;
var TOBMinute,TOBOUVPP : TOB;

procedure AddlesChamps (TOBDM :TOB);
begin
	TOBDM.AddChampSupValeur('BMN_LIGNE',0);
  TOBDM.AddChampSupValeur('BMN_NUMEROTATION',''); //Nouvelle Numérotation de Ligne
	TOBDM.AddChampSupValeur('BMN_PARAG','');
	TOBDM.AddChampSupValeur('BMN_SPARAG','');
  TOBDM.AddChampSupValeur('BMN_SPARAG3','');
  TOBDM.AddChampSupValeur('BMN_SPARAG4','');
  TOBDM.AddChampSupValeur('BMN_SPARAG5','');
  TOBDM.AddChampSupValeur('BMN_SPARAG6','');
  TOBDM.AddChampSupValeur('BMN_SPARAG7','');
  TOBDM.AddChampSupValeur('BMN_SPARAG8','');
  TOBDM.AddChampSupValeur('BMN_SPARAG9','');
	TOBDM.AddChampSupValeur('BMN_CATEGORIE',''); // Catégorie : type de ressource
	TOBDM.AddChampSupValeur('BMN_NATURE',''); // Nature de la ligne Fourniture,Salarie,sous traitant, etc..
	TOBDM.AddChampSupValeur('BMN_FAMILLE1',''); // Famille de niveau 1
	TOBDM.AddChampSupValeur('BMN_FAMILLE2',''); // Famille de niveau 2
	TOBDM.AddChampSupValeur('BMN_FAMILLE3',''); // Famille de niveau 3
	TOBDM.AddChampSupValeur('BMN_OUVRAGE',''); // Code ouvrage
	TOBDM.AddChampSupValeur('BMN_UNITEOUV',''); // Unité ouvrage
	TOBDM.AddChampSupValeur('BMN_QTEOUV',''); // Qté ouvrage
	TOBDM.AddChampSupValeur('BMN_ARTICLE',''); // Code article
	TOBDM.AddChampSupValeur('BMN_QTEACH',0.0); // Qte marchandise
	TOBDM.AddChampSupValeur('BMN_UNITEACH',''); // Unite marchandise
	TOBDM.AddChampSupValeur('BMN_PAACH',0.0); // Pa marchandise
	TOBDM.AddChampSupValeur('BMN_MTACH',0.0); // Pa marchandise
	TOBDM.AddChampSupValeur('BMN_QTEHRS',0.0); // Qte MO
	TOBDM.AddChampSupValeur('BMN_UNITEHRS',''); // Unite MO
	TOBDM.AddChampSupValeur('BMN_PAHRS',0.0); // PA MO
	TOBDM.AddChampSupValeur('BMN_MTHRS',0.0); // Mt MO
	TOBDM.AddChampSupValeur('BMN_QTE',0.0); // Qte Unitaire
	TOBDM.AddChampSupValeur('BMN_PUV',0.0); // Pu Vente
	TOBDM.AddChampSupValeur('BMN_MTPUV',0.0); // Montant PV
	TOBDM.AddChampSupValeur('BMN_PR',0.0); // Pu revient
	TOBDM.AddChampSupValeur('BMN_MTPR',0.0); // Montant Revient

end;

procedure  AddLignePrestation (TOBMinute,TOBOUV,TOBL,TOBArticles : TOB; NiveauImbric : integer; LibP1,LibP2,LibP3,LibP4,LibP5,LibP6,LibP7,LibP8,LibP9 : string);
var prefixe : string;
		TOBArt  : TOB;
    TOBDM   : TOB;
    NewArt  : boolean;
    QQ      : TQuery;
    QtHeure : Double;
begin

	if TOBL.NomTable = 'LIGNE' then
  	preFixe := 'GL'
  else
  	preFixe := 'BOP';

  //FV1 : Si Qté d'heure à zéro sur une prestation la ligne ne doit pas être prise en compte dans la génération de la minute
  QtHeure := TOBL.GetValue(prefixe+'_QTEFACT');
  if QtHeure = 0 then exit;

  TOBDM := TOB.Create ('LIGNE MINUTE',TOBMinute,-1);

  AddlesChamps(TOBDM);

  newArt := false;

  TOBART := TOBArticles.findFirst(['GA_ARTICLE'], [TOBL.GetValue(prefixe + '_ARTICLE')],true);

  if TOBART = Nil then
  begin
    TOBART := TOB.Create('ARTICLE',nil,-1);
    QQ := OpenSQl('SELECT GA_QUALIFUNITEVTE, GA_NATUREPRES,GA_FAMILLENIV1,GA_FAMILLENIV2,'+
    							'GA_FAMILLENIV3,BNP_LIBELLE,BNP_TYPERESSOURCE FROM ARTICLE '+
									'LEFT OUTER JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES '+
    						  'WHERE GA_ARTICLE="'+ TOBL.GetValue(prefixe+'_ARTICLE')+'"',true,1,'',true);
    if not QQ.eof then TOBART.SelectDB('',QQ);
    ferme(QQ);
    newArt := true;
  end;

  if NiveauImbric > 8 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
    TOBDM.PutValue('BMN_SPARAG7',LIBP7);
    TOBDM.PutValue('BMN_SPARAG8',LIBP8);
    TOBDM.PutValue('BMN_SPARAG9',LIBP9);
  end else
  if NiveauImbric > 7 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
    TOBDM.PutValue('BMN_SPARAG7',LIBP7);
    TOBDM.PutValue('BMN_SPARAG8',LIBP8);
  end else
  if NiveauImbric > 6 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
    TOBDM.PutValue('BMN_SPARAG7',LIBP7);
  end else
  if NiveauImbric > 5 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
  end else
  if NiveauImbric > 4 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
  end else
  if NiveauImbric > 3 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
  end else
  if NiveauImbric > 2 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
  end else
  if NiveauImbric > 1 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',LIBP2);
  end
  else if NiveauImbric > 0 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG','   0 : Hors paragraphe');
  end
  else
  begin
    TOBDM.PutValue('BMN_PARAG', '  0 : Hors paragraphe');
    TOBDM.PutValue('BMN_SPARAG','   0 : Hors paragraphe');
  end;

	TOBDM.PutValue('BMN_CATEGORIE', TOBArt.getValue('BNP_TYPERESSOURCE'));
	TOBDM.PutValue('BMN_NATURE',    TOBArt.getValue('GA_NATUREPRES')+' : '+TOBArt.getValue('BNP_LIBELLE'));

	TOBDM.PutValue('BMN_FAMILLE1',rechdom('GCFAMILLENIV1',TOBArt.getValue('GA_FAMILLENIV1'),false)); // Famille de niveau 1
	TOBDM.PutValue('BMN_FAMILLE2',rechdom('GCFAMILLENIV2',TOBArt.getValue('GA_FAMILLENIV2'),false)); // Famille de niveau 2
	TOBDM.PutValue('BMN_FAMILLE3',rechdom('GCFAMILLENIV3',TOBArt.getValue('GA_FAMILLENIV3'),false)); // Famille de niveau 3

  if TOBOUV <> nil then
  begin
		TOBDM.PutValue('BMN_OUVRAGE',TOBOUV.GetValue('GL_CODEARTICLE')+' - '+TOBOUV.GetValue('GL_LIBELLE')+
                                                 ' : '+FloatToStr(TOBOUV.GetValue('GL_QTEFACT'))+
                                                 ' '+TOBOUV.GetValue('GL_QUALIFQTEVTE')); // Code ouvrage + libelle + qté + unité
	  TOBDM.PutValue('BMN_QTEOUV',TOBOUV.GetValue('GL_QTEFACT')); // Qte ouvrage
	  TOBDM.PutValue('BMN_UNITEOUV',TOBOUV.GetValue('GL_QUALIFQTEVTE')); // Unité ouvrage
    if TOBOUV.GetValue('GL_QTEFACT') <> 0 then
      TOBDM.PutValue('BMN_QTE',(TOBL.GetValue(prefixe+'_QTEFACT')/TOBOUV.GetValue('GL_QTEFACT'))); // Qte  Unitaire
    //FV1 : 28/06/2016 - Nouvelle numérotation des lignes...
  	TOBDM.PutValue('BMN_NUMEROTATION', TOBOUV.GetValue('GLC_NUMEROTATION'));
  end
  else
  begin
  	TOBDM.PutValue('BMN_OUVRAGE','Hors ouvrage');
	  TOBDM.PutValue('BMN_QTE',TOBL.GetValue(prefixe+'_QTEFACT')); // Qte  Unitaire
    //FV1 : 28/06/2016 - Nouvelle numérotation des lignes...
  	TOBDM.PutValue('BMN_NUMEROTATION', TOBL.GetValue('GLC_NUMEROTATION'));
  end;

	TOBDM.PutValue('BMN_ARTICLE',TOBL.GetValue(prefixe+'_CODEARTICLE')+' : '+TOBL.GetValue(prefixe+'_LIBELLE')); // Code article

  if pos(TOBArt.getValue('BNP_TYPERESSOURCE'), GetparamSocSecur ('SO_BTMOINTERNE','INT;SAL'))>0  then
  begin
    // modif brl 30/03/2011 :
    // en attendant l'ajout du champ BOP_QUALIFQTEVTE manquant dans la table LIGNEOUVPLAT
    // on prend l'unité de vente dans la table ARTICLE
    if prefixe = 'BOP' then
      TOBDM.PutValue('BMN_UNITEHRS',TOBArt.GetValue('GA_QUALIFUNITEVTE'))
    else // Unite MO
      TOBDM.PutValue('BMN_UNITEHRS',TOBL.GetValue(prefixe+'_QUALIFQTEVTE')); // Unite MO

    TOBDM.PutValue('BMN_QTEHRS',TOBL.GetValue(prefixe+'_QTEFACT'));          // Qte MO
    TOBDM.PutValue('BMN_PAHRS',TOBL.GetValue(prefixe+'_DPA'));               // PA MO
    TOBDM.PutValue('BMN_MTHRS',TOBL.GetValue(prefixe+'_MONTANTPA'));         // Mt MO
    {*
    TOBDM.PutValue('BMN_QTEHRS',0.0);  // Qte MO
  	TOBDM.PutValue('BMN_UNITEHRS',''); // Unite MO
	  TOBDM.PutValue('BMN_PAHRS',0.0);   // PA MO
	  TOBDM.PutValue('BMN_MTHRS',0.0);   // Mt MO
    *}
  end else
  begin
    TOBDM.PutValue('BMN_QTEACH',TOBL.GetValue(prefixe+'_QTEFACT')); // Qte marchandise

    // modif brl 30/03/2011 :
    // en attendant l'ajout du champ BOP_QUALIFQTEVTE manquant dans la table LIGNEOUVPLAT
    // on prend l'unité de vente dans la table ARTICLE
    if prefixe = 'BOP' then
      TOBDM.PutValue('BMN_UNITEACH',TOBArt.GetValue('GA_QUALIFUNITEVTE'))  // Unite marchandise
    else
      TOBDM.PutValue('BMN_UNITEACH',TOBL.GetValue(prefixe+'_QUALIFQTEVTE')); // Unite Ligne

    TOBDM.PutValue('BMN_PAACH',TOBL.GetValue(prefixe+'_DPA')); // Pa marchandise
    TOBDM.PutValue('BMN_MTACH',TOBL.GetValue(prefixe+'_MONTANTPA')); // Pa marchandise
  end;

	TOBDM.PutValue('BMN_PUV',TOBL.GetValue(prefixe+'_PUHTDEV')); // Pu Vente
	TOBDM.PutValue('BMN_MTPUV',TOBL.GetValue(prefixe+'_MONTANTHTDEV')); // Montant PV
	TOBDM.PutValue('BMN_PR',TOBL.GetValue(prefixe+'_DPR')); // Pu Revient
	TOBDM.PutValue('BMN_MTPR',TOBL.GetValue(prefixe+'_MONTANTPR')); // Montant revient

  if TOBOUV <> nil then
  begin
		TOBDM.PutValue('BMN_LIGNE',TOBL.GetValue(prefixe+'_NUMLIGNE')); // Numéro Ligne
  end
  else
  begin
		TOBDM.PutValue('BMN_LIGNE',TOBL.GetValue('GL_NUMLIGNE')); // Numéro Ligne
  end;

  if NewArt then TOBArt.free;

end;

procedure  AddLigneMarchandise (TOBMinute,TOBOUV,TOBL,TOBArticles : TOB; NiveauImbric : integer; LibP1,LibP2,LibP3,LibP4, LibP5,LibP6,LibP7,Libp8, Libp9 : string; WithMemoOuv : boolean=true);
var prefixe : string;
		TOBArt  : TOB;
    TOBDM   : TOB;
    NewArt  : boolean;
    StSQl   : String;
    QQ      : TQuery;
begin
	if TOBL.NomTable = 'LIGNE' then
  begin
  	preFixe := 'GL';
  end else
  begin
  	preFixe := 'BOP';
  end;
  //
  TOBDM := TOB.Create ('LIGNE MINUTE',TOBMinute,-1);
  AddlesChamps(TOBDM);
  newArt := false;
  TOBART := TOBArticles.findFirst(['GA_ARTICLE'], [TOBL.GetValue(prefixe+'_ARTICLE')],true);

  if TOBART = Nil then
  begin
  	TOBART := TOB.Create('ARTICLE',nil,-1);
    QQ := OpenSQl('SELECT GA_QUALIFUNITEVTE, GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3 FROM ARTICLE WHERE GA_ARTICLE="'+
                 TOBL.GetValue(prefixe+'_ARTICLE')+'"',true,1,'',true);
    if not QQ.eof then
    begin
    	TOBART.SelectDB('',QQ);
    end;
    newArt := true;
  end;

  if NiveauImbric > 8 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
    TOBDM.PutValue('BMN_SPARAG7',LIBP7);
    TOBDM.PutValue('BMN_SPARAG8',LIBP8);
    TOBDM.PutValue('BMN_SPARAG9',LIBP9);
  end else
  if NiveauImbric > 7 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
    TOBDM.PutValue('BMN_SPARAG7',LIBP7);
    TOBDM.PutValue('BMN_SPARAG8',LIBP8);
  end else
  if NiveauImbric > 6 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
    TOBDM.PutValue('BMN_SPARAG7',LIBP7);
  end else
  if NiveauImbric > 5 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
    TOBDM.PutValue('BMN_SPARAG6',LIBP6);
  end else
  if NiveauImbric > 4 then
  begin
    TOBDM.PutValue('BMN_PARAG',  LibP1);
    TOBDM.PutValue('BMN_SPARAG', LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
    TOBDM.PutValue('BMN_SPARAG5',LIBP5);
  end else
  if NiveauImbric > 3 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
    TOBDM.PutValue('BMN_SPARAG4',LIBP4);
  end else
if NiveauImbric > 2 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',LIBP2);
    TOBDM.PutValue('BMN_SPARAG3',LIBP3);
  end else if NiveauImbric > 1 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',LIBP2);
  end else if NiveauImbric > 0 then
  begin
    TOBDM.PutValue('BMN_PARAG',LibP1);
    TOBDM.PutValue('BMN_SPARAG',' 0 : Hors paragraphe');
  end else
  begin
    TOBDM.PutValue('BMN_PARAG','  0 : Hors paragraphe');
    TOBDM.PutValue('BMN_SPARAG',' 0 : Hors paragraphe');
  end;

	TOBDM.PutValue('BMN_CATEGORIE','MAR');
	TOBDM.PutValue('BMN_NATURE','Fournitures');
	TOBDM.PutValue('BMN_FAMILLE1',rechdom('GCFAMILLENIV1',TOBArt.getValue('GA_FAMILLENIV1'),false)); // Famille de niveau 1
	TOBDM.PutValue('BMN_FAMILLE2',rechdom('GCFAMILLENIV2',TOBArt.getValue('GA_FAMILLENIV2'),false)); // Famille de niveau 2
	TOBDM.PutValue('BMN_FAMILLE3',rechdom('GCFAMILLENIV3',TOBArt.getValue('GA_FAMILLENIV3'),false)); // Famille de niveau 3
  if (TOBOUV <> nil) and (WithMemoOuv)  then
  begin
		TOBDM.PutValue('BMN_OUVRAGE',TOBOUV.GetValue('GL_CODEARTICLE')+' - '+TOBOUV.GetValue('GL_LIBELLE')+
                                                 ' : '+FloatToStr(TOBOUV.GetValue('GL_QTEFACT'))+
                                                 ' '+TOBOUV.GetValue('GL_QUALIFQTEVTE')); // Code ouvrage + libelle + qté + unité
	  TOBDM.PutValue('BMN_QTEOUV',TOBOUV.GetValue('GL_QTEFACT')); // Qte ouvrage
	  TOBDM.PutValue('BMN_UNITEOUV',TOBOUV.GetValue('GL_QUALIFQTEVTE')); // Unité ouvrage
    if TOBOUV.GetValue('GL_QTEFACT') <> 0 then TOBDM.PutValue('BMN_QTE',(TOBL.GetValue(prefixe+'_QTEFACT')/TOBOUV.GetValue('GL_QTEFACT'))); // Qte Unitaire
    //FV1 : 28/06/2016 - Nouvelle numérotation des lignes...
	  TOBDM.PutValue('BMN_NUMEROTATION', TOBOUV.GetValue('GLC_NUMEROTATION'));
  end
  else
  begin
  	TOBDM.PutValue('BMN_OUVRAGE','Hors ouvrage');
	  TOBDM.PutValue('BMN_QTE',TOBL.GetValue(prefixe+'_QTEFACT')); // Qte Unitaire
    //FV1 : 28/06/2016 - Nouvelle numérotation des lignes...
	  TOBDM.PutValue('BMN_NUMEROTATION', TOBL.GetValue('GLC_NUMEROTATION'));
  end;

	//TOBDM.PutValue('BMN_ARTICLE',TOBL.GetValue(prefixe+'_CODEARTICLE')+' : '+TOBL.GetValue(prefixe+'_LIBELLE')); // Code article
  //FV1 : 16/07/2013 => Récupération du libellé de la saisie si Article en prix posé, on remonte
  //d'un niveau de façon à récupérer le libellé saisit dans le document.
	if TOBL.GetString('BOP_TYPEARTICLE') = 'ARP' then
  begin
    StSQl := 'SELECT BLO_LIBELLE FROM LIGNEOUV WHERE BLO_NATUREPIECEG="' + TOBL.GetString('BOP_NATUREPIECEG') + '"' +
             ' AND BLO_SOUCHE="'   + TOBL.GetString('BOP_SOUCHE')   + '"'     +
             ' AND BLO_NUMERO= '   + IntToStr(TOBL.GetValue('BOP_NUMERO'))    +
             ' AND BLO_INDICEG= '  + IntToStr(TOBL.GetInteger('BOP_INDICEG')) +
             ' AND BLO_NUMLIGNE= ' + IntToStr(TOBL.GetInteger('BOP_NUMLIGNE'))+
             ' AND BLO_N1= ' + IntToStr(TOBL.GetInteger('BOP_N1'));
    if TOBL.GetInteger('BOP_N5')<> 0 then
    begin
      StSQL := StSQL + ' AND BLO_N2= ' + IntToStr(TOBL.GetInteger('BOP_N2')) + ' AND BLO_N3= ' + IntToStr(TOBL.GetInteger('BOP_N3')) +
               ' AND BLO_N4= ' + IntToStr(TOBL.GetInteger('BOP_N4'));
    end
    else if TOBL.GetInteger('BOP_N4')<> 0 then
    begin
      StSQL := StSQL + ' AND BLO_N2= ' + IntToStr(TOBL.GetInteger('BOP_N2')) + ' AND BLO_N3= ' + IntToStr(TOBL.GetInteger('BOP_N3'));
    end
    else if TOBL.GetInteger('BOP_N3')<> 0 then
    begin
      StSQL := StSQL + ' AND BLO_N2= ' + IntToStr(TOBL.GetInteger('BOP_N2'));
    end;
    QQ := OpenSQL(StSQL,true);
    if QQ.eof then
      TOBDM.PutValue('BMN_ARTICLE',TOBL.GetValue(prefixe+'_CODEARTICLE')+' : ' + TOBL.GetValue(Prefixe + '_LIBELLE'))  // Code article
    else
      TOBDM.PutValue('BMN_ARTICLE',TOBL.GetValue(prefixe+'_CODEARTICLE')+' : ' + QQ.FindField('BLO_LIBELLE').AsString); // Code article
  end
  else
      TOBDM.PutValue('BMN_ARTICLE',TOBL.GetValue(prefixe+'_CODEARTICLE')+' : ' + TOBL.GetValue(prefixe+'_LIBELLE')); // Code article

	TOBDM.PutValue('BMN_QTEACH',TOBL.GetValue(prefixe+'_QTEFACT')); // Qte marchandise
  // modif brl 30/03/2011 :
  // en attendant l'ajout du champ BOP_QUALIFQTEVTE manquant dans la table LIGNEOUVPLAT
  // on prend l'unité de vente dans la table ARTICLE
  if prefixe = 'BOP' then TOBDM.PutValue('BMN_UNITEACH',TOBArt.GetValue('GA_QUALIFUNITEVTE'))else // Unite marchandise
                          TOBDM.PutValue('BMN_UNITEACH',TOBL.GetValue(prefixe+'_QUALIFQTEVTE')); // Unite marchandise
	TOBDM.PutValue('BMN_PAACH',TOBL.GetValue(prefixe+'_DPA')); // Pa marchandise
	TOBDM.PutValue('BMN_MTACH',TOBL.GetValue(prefixe+'_MONTANTPA')); // Pa marchandise
	TOBDM.PutValue('BMN_QTEHRS',0.0); // Qte MO
	TOBDM.PutValue('BMN_UNITEHRS',''); // Unite MO
	TOBDM.PutValue('BMN_PAHRS',0.0); // PA MO
	TOBDM.PutValue('BMN_MTHRS',0.0); // Mt MO
	TOBDM.PutValue('BMN_PUV',TOBL.GetValue(prefixe+'_PUHTDEV')); // Pu Vente
	TOBDM.PutValue('BMN_MTPUV',TOBL.GetValue(prefixe+'_MONTANTHTDEV')); // Montant PV
	TOBDM.PutValue('BMN_PR',TOBL.GetValue(prefixe+'_DPR')); // Pu Revient
	TOBDM.PutValue('BMN_MTPR',TOBL.GetValue(prefixe+'_MONTANTPR')); // Montant revient
  if (TOBOUV <> nil) and (WithMemoOuv)  then
  begin
		TOBDM.PutValue('BMN_LIGNE',TOBL.GetValue(prefixe+'_NUMLIGNE')); // Numéro Ligne
  end else
  begin
		TOBDM.PutValue('BMN_LIGNE',TOBL.GetValue('GL_NUMLIGNE')); // Numéro Ligne
  end;

  if NewArt then TOBArt.free;

end;

procedure AddOuvragePlat (TOBminute,TOBPiece,TOBTiers,TOBL,TOBOuvrages,TOBOuvragesP,TOBArticles: TOB;Niv: integer; LibP1,LibP2, LibP3,LibP4,LibP5,LibP6,LibP7,LibP8,Libp9 : string; DEV : Rdevise; WithMemoOuv : Boolean=true);
var TOBPlat,TOBOL : TOB;
		Indice,ITypeArt : integer;
begin
  TOBPLat := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('GL_NUMORDRE'));
  GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers,DEV,False);
  for Indice := 0 to TOBPLat.detail.count -1 do
  begin
    TOBOL := TOBPlat.detail[Indice];
    if TOBPiece.GetValue('GP_FACTUREHT')='X' then CalculeLigneHT (TOBOL,nil,TOBPiece,DEV,DEV.decimale)
                                             else CalculeLigneTTC (TOBOL,nil,TOBPiece,DEV,DEV.decimale);
  	if Indice = 0 then
    begin
    	ITypeArt := TOBOL.GetNumChamp('BOP_TYPEARTICLE');
    end;
    if TOBOL.GetValeur(ITypeArt)='PRE' then
    begin
      AddLignePrestation (TOBMinute,TOBL,TOBOL,TOBArticles,Niv,LibP1,LibP2, LibP3, LibP4, LibP5, LibP6, LibP7, LibP8, LibP9);
    end else
    begin
      AddLigneMarchandise (TOBMinute,TOBL,TOBOL,TOBArticles,Niv,LibP1,LibP2, LibP3, LibP4, LibP5, LibP6, LibP7, LibP8, LibP9,WithMemoOuv);
    end;
  end;
end;

function IsOnlyMarchandise (TOBL,TOBOuvrages : TOB) : Boolean;
var IndiceNomen : Integer;
		TOBOUV : TOB;
begin
  Result := false;
  if TOBOuvrages = nil then Exit;
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;
  TOBOUV := TOBOuvrages.detail[IndiceNomen -1];
  if TOBOUV.detail[0].detail.Count > 1 then Exit;
  Result := True;
end;


procedure ConstitueLaMinute (TOBMinute,TOBPiece,TOBTiers,TOBOUvrages,TOBOuvragesP,TOBArticles : TOB; DEV : Rdevise);

  function ConstitueNumPararag ( NumP1: integer=0;Nump2: integer=0; Nump3: integer=0; Nump4: integer=0; Nump5: integer=0; Nump6: integer=0; Nump7: integer=0; Nump8: integer=0; Nump9: integer=0 ) : string;
  begin
    Result := Format('%3d',[Nump1]);
    if Nump2 > 0 then Result := result + '.'+Format('%3d',[Nump2]);
    if Nump3 > 0 then Result := result + '.'+Format('%3d',[Nump3]);
    if Nump4 > 0 then Result := result + '.'+Format('%3d',[Nump4]);
    if Nump5 > 0 then Result := result + '.'+Format('%3d',[Nump5]);
    if Nump6 > 0 then Result := result + '.'+Format('%3d',[Nump6]);
    if Nump7 > 0 then Result := result + '.'+Format('%3d',[Nump7]);
    if Nump8 > 0 then Result := result + '.'+Format('%3d',[Nump8]);
    if Nump9 > 0 then Result := result + '.'+Format('%3d',[Nump9]);
  end;

var indice : integer;
		TOBL : TOB;
    I1,I2,I3,I4,I5,I6,I7,I8,I9, Niv : integer;
    LibP1,libP2, LibP3,Libp4,LibP5,LibP6,LibP7,LibP8,LibP9 : string;
    ITypeLigne,ITypeArt,ILibelle,IImbric : integer;
begin

	I1 := 0;
	I2 := 0;
  I3 := 0;
  I4 := 0;
  I5 := 0;
  I6 := 0;
  I7 := 0;
  I8 := 0;
  I9 := 0;

  ITypeLigne  := -1;
  ITypeArt    := -1;
  ILibelle    := -1;

	for Indice := 0 to TOBpiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
  	if Indice = 0 then
    begin
    	ITypeLigne  := TOBL.GetNumChamp('GL_TYPELIGNE');
    	ITypeArt    := TOBL.GetNumChamp('GL_TYPEARTICLE');
    	ILibelle    := TOBL.GetNumChamp('GL_LIBELLE');
    	IImbric     := TOBL.GetNumChamp('GL_NIVEAUIMBRIC');
    end;
    //
    Niv := TOBL.GetValeur(IImbric);
    //
    if (TOBL.GetValeur(ITypeLigne)<>'ART') AND (Pos(copy(TOBL.GetValeur(ITypeLigne),1,2),'DP;DV;TP;TV')=0) then continue;

    if Pos(copy(TOBL.GetValeur(ITypeLigne),1,2),'DP;DV')>0 then // Début de paragraphe
    begin
    	if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))>9 then continue;
      //
    	if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=9 then
      begin
      	Inc(I9); LibP9 := ConstitueNumPararag (I1,I2,I3,I4,I5,I6,I7,I8,I9) +' : ' +TOBL.GetValeur(ILibelle);
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=8 then
      begin
      	Inc(I8); LibP8 := ConstitueNumPararag (I1,I2,I3,I4,I5,I6,I7,I8) +' : ' +TOBL.GetValeur(ILibelle);
        I9 := 0; LibP9 := '';
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=7 then
      begin
      	Inc(I7); LibP7 := ConstitueNumPararag (I1,I2,I3,I4,I5,I6,I7) +' : ' +TOBL.GetValeur(ILibelle);
        I8 := 0; LibP8 := '';
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=6 then
      begin
      	Inc(I6); LibP6 := ConstitueNumPararag (I1,I2,I3,I4,I5,I6) +' : ' +TOBL.GetValeur(ILibelle);
        I7 := 0; LibP7 := '';
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=5 then
      begin
      	Inc(I5); LibP5 := ConstitueNumPararag (I1,I2,I3,I4,I5) +' : ' +TOBL.GetValeur(ILibelle);
        I6 := 0; LibP6 := '';
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=4 then
      begin
      	Inc(I4); LibP4 := ConstitueNumPararag (I1,I2,I3,I4) +' : ' +TOBL.GetValeur(ILibelle);
        I5 := 0; LibP5 := '';
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=3 then
      begin
      	Inc(I3); LibP3 := ConstitueNumPararag (I1,I2,I3) +' : ' +TOBL.GetValeur(ILibelle);
        I4 := 0; LibP4 := '';
      end
    	else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=2 then
      begin
      	Inc(I2); LibP2 := ConstitueNumPararag (I1,I2) +' : ' +TOBL.GetValeur(ILibelle);
        I3 := 0; LibP3 := '';
      end
      else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=1 then
      begin
      	Inc(I1); LibP1 := ConstitueNumPararag (I1) +' : ' +TOBL.GetValeur(ILibelle);
        I2 := 0; LibP2 := '';
      end;
    end
    else if Pos(copy(TOBL.GetValeur(ITypeLigne),1,2),'TP;TV')>0 then // Fin de paragraphe
    begin
    	if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))>9 then continue;
      //
      if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=9 then
      begin
      	LibP9:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=8 then
      begin
      	LibP8:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=7 then
      begin
      	LibP7:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=6 then
      begin
      	LibP6:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=5 then
      begin
      	LibP5:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=4then
      begin
      	LibP4:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=3then
      begin
      	LibP3:= '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=2 then
      begin
      	LibP2 := '';
      end else if StrToInt(copy(TOBL.GetValeur(ITypeLigne),3,1))=1 then
      begin
      	LibP1 := '';
      end;
    end
    else if  Pos(TOBL.GetValeur(ITypeArt),'OUV')>0 then // ouvrage ou article en prix posés
    begin
      AddOuvragePlat (TOBminute,TOBPiece,TOBTiers,TOBL,TOBOuvrages,TOBOuvragesP,TOBArticles,Niv,LibP1,LibP2, LibP3,LibP4,LibP5,LibP6,LibP7,LibP8,Libp9,DEV);
    end
    else if  Pos(TOBL.GetValeur(ITypeArt),'ARP')>0 then // ouvrage ou article en prix posés
    begin
      if IsOnlyMarchandise (TOBL,TOBOuvrages) then
      begin
        AddOuvragePlat (TOBminute,TOBPiece,TOBTiers,TOBL,TOBOuvrages,TOBOuvragesP,TOBArticles,Niv,LibP1,LibP2,LibP3,Libp4,LibP5,LibP6,LibP7,LibP8,LibP9,DEV,false);
      end else
      begin
        AddOuvragePlat (TOBminute,TOBPiece,TOBTiers,TOBL,TOBOuvrages,TOBOuvragesP,TOBArticles,Niv,LibP1,LibP2,LibP3,Libp4,LibP5,LibP6,LibP7,LibP8,LibP9,DEV);
      end;
    end
    else if TOBL.GetValeur(ITypeArt)='PRE' then
    begin
      AddLignePrestation (TOBMinute,nil,TOBL,TOBArticles,Niv,LibP1,LibP2,LibP3,LibP4,LibP5,LibP6,LibP7,LibP8,LibP9);
    end
    else
    begin
      AddLigneMarchandise (TOBMinute,nil,TOBL,TOBArticles,Niv,LibP1,LibP2,LibP3,LibP4,LibP5,LibP6,LibP7,LibP8,LibP9);
    end;
  end;
end;

procedure LanceMinuteDevis(TOBPiece,TOBTiers,TOBOuvrages,TOBOuvragesP,TOBARTicles : TOB; DEV : RDEVISE);
var Titre : string;
begin
 	TOBMinute := TOB.Create ('LA MINUTE',nil,-1);
  TOBOUVPP := TOB.Create ('LES OV PLATS',nil,-1);
  TRY
  	ConstitueLaMinute (TOBMinute,TOBPiece,TOBTiers,TOBOUvrages,TOBOuvPP,TOBArticles,DEV);
    titre := 'LIBELLE=Minute '+
    				 rechdom('GCNATUREPIECEG',TOBpiece.getValue('GP_NATUREPIECEG'),false)+
             ' N° '+Inttostr(TOBPiece.getValue('GP_NUMERO'))+ ' - Client : '+ TOBPiece.getValue('GP_TIERS')+' '+TOBTiers.getValue('T_LIBELLE') ;
    V_PGI.AutoSearch := true;
    AGLLanceFiche('BTP','BTMINUTEDEVIS','','',Titre);
    V_PGI.AutoSearch := false;
  FINALLY
  	TOBMinute.free;
    TOBOUVPP.free;
  END;
end;

procedure TOF_BTMINUTEDEV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTMINUTEDEV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTMINUTEDEV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTMINUTEDEV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTMINUTEDEV.OnArgument (S : String ) ;
var St1 : String;
begin
  Inherited ;
	st1 := Copy(S, 0, Pos('=',S) - 1);
  if st1 = 'LIBELLE' then
  begin
    st1 := Copy(S, Pos('=',S) + 1, 255);
    ecran.Caption := st1;
  end;
  TFStat (Ecran).LaTOB := TOBminute;

end ;

procedure TOF_BTMINUTEDEV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTMINUTEDEV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMINUTEDEV.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_BTMINUTEDEV ] ) ; 
end.
