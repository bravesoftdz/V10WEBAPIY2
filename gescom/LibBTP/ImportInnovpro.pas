{***********UNITE*************************************************
Auteur  ...... : Santucci Lionel
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : IMPORTINNOVPRO ()
Mots clefs ... : TOF;IMPORTINNOVPRO
*****************************************************************}
Unit ImportInnovpro ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     uTob,
     maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Dialogs,
     HTB97,
     AglInit,
     EntGc,
     UTOB,
     UTOF,
     UtilArticle,
     CalcOleGenericBTP,
     M3FP,
     BTPUtil,
     UtilSaisieconso ;

Type
  TTypeInfosEnreg = (TtieChantier,TtiePhase,TtieDateMouv, TTieNatureMouv,TTieRessource,TTiePrestation,TTieDesignation,TTieQTE,TTieUnite,TTiePUA,TTiePUR,TTieTypeHeure);
  TTypeErreurControle  = (TcdErrNumConso,TcdErrAffaire,TcdErrPhase,TcdErrNatureMouv,TcdErrRessource,TcdErrCodeArticle,TcdErrQte,TcdErrTypeHeure);
  TOF_IMPORTINNOVPRO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	OpenDlg : TopenDialog;
    Rapport : Tmemo;
    TOBConso,TOBressources,TOBARTicles : TOB;
    LogFile : textfile;
    NbTraiteOk : integer;
    Numintegr : integer;
    DateIntegre : TdateTime;
    procedure FicNameClick ( sender : Tobject);
    function GetInfosEnreg(st_enreg : string; typeInfo : TTypeInfosEnreg) : variant;
    procedure EnregistreConso;
    procedure ControleDonnee (TOBD : TOB);
    procedure FinaliseConsommation (TOBD : TOB);
    function IsExisteArticle(TOBD: TOB): boolean;
    function OKTypeHeure(Typeheure: string): boolean;
    function AddRessource(TOBD : TOB) : TOB;
		function AddArticle (TOBD : TOB) : TOB;
    procedure SetErreur (TOBD : TOB);
    procedure LanceTraitementImport(sender: Tobject);
    procedure EnregistreInfos(st_enreg: string; Ligne: integer);
  	procedure EnregistreFinTrait(NbrLigne : integer);
    procedure EnregistreMvt(TOBD : TOB);

  end ;

  TOF_ANNULIMPORT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    procedure AnnulationImportConso (Numero : integer);
  end ;

procedure ImportConsoInnovPro;
procedure AnnulImportConsoInnovPro;

Implementation

procedure AnnulImportConsoInnovPro;
begin
	AglLancefiche ('BTP','BANNULTIMPINNOVPR','','','ACTION=MODIFICATION');
end;

procedure ImportConsoInnovPro;
begin
	AglLancefiche ('BTP','BTIMPORTINNOVPRO','','','ACTION=MODIFICATION');
end;

procedure TOF_IMPORTINNOVPRO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTINNOVPRO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTINNOVPRO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTINNOVPRO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTINNOVPRO.OnArgument (S : String ) ;
begin
  Inherited ;
  OpenDlg := TOpenDialog.Create(ecran);
  OpenDlg.Filter := 'Fichier InnovPro (*.pro)|*.PRO';
  Thedit(GetCOntrol('FICNAME')).OnElipsisClick  := FicNameClick;
  TToolbarButton97(GetCOntrol('BLANCETRAIT')).OnClick  := LanceTraitementImport;
  Rapport := TMemo(GetControl('RAPPORT'));
  TOBConso := TOB.Create ('LES CONSO',nil,-1);
  TOBressources := TOB.Create ('LES RESSOURCES',nil,-1);
	TOBArticles := TOB.Create ('LES ARTICLES',nil,-1);
  NbTraiteOk := 0;
end ;

procedure TOF_IMPORTINNOVPRO.OnClose ;
begin
  Inherited ;
  OpenDlg.Free;
  TOBConso.free;
  TOBressources.free;
  TOBArticles.free;
end ;

procedure TOF_IMPORTINNOVPRO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTINNOVPRO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_IMPORTINNOVPRO.FicNameClick(sender: Tobject);
begin
  if OpenDlg.Execute then
  begin
  	Thedit(GetCOntrol('FICNAME')).Text  := OpenDlg.FileName;
  end;
end;

procedure TOF_IMPORTINNOVPRO.LanceTraitementImport(sender: Tobject);
var InputFile : textfile;
		nomfic,NomLog,NewNomFic,NewNomLog : string;
    st_enreg : string;
    Ioerror : Tioerr;
    Annee,Mois,jour,Heure,Min,Sec,Msec : word;
    NOLIgne : integer;
begin
  DateIntegre := Now;
  NOLigne := 1;
	DecodeDate(DateIntegre, Annee, Mois, Jour);
	DecodeTime (DateIntegre,heure, min, sec,msec);
  if not GetNumCompteur ('INP',iDate1900,NumIntegr) then exit;
	Nomfic := Thedit(GetCOntrol('FICNAME')).Text;
  NewNomFic := format ('%sCONSO%d-%2.2d%2.2d%4.4d-%2.2d%2.2d.INT',[ExtractFilePath (Nomfic),Numintegr,jour,mois,annee,heure,min]);
//  newNomFic := ExtractFilePath (Nomfic)+'CONSO-'+InttoStr(jour)+'-'+InttoStr(Mois)+
//  						 '-'+InttoStr(Annee)+'-'+InttoStr(Heure)+':'+InttoStr(min)+'.INT';
  NomLog := ChangeFileExt (Nomfic,'.log');
  NewNomLog := ChangeFileExt (newNomfic,'.log');
  if Nomfic = '' then exit;
  Rapport.Clear;
  if FileExists (NomFic) then
  begin
    AssignFile(InputFile,NomFic);
    AssignFile (LogFile,NomLog);
    Reset (InputFile);
    rewrite(Logfile);
    IoError := OeOk;
    TRY
//      readln (InputFile,st_enreg);
      while not EOF(InputFile) do
      Begin
        readln (InputFile,st_enreg);
        EnregistreInfos (st_enreg,NOLigne);
        if TOBConso.detail.count > 1000 then
        begin
          Ioerror := TRANSACTIONS(EnregistreConso,0);
        	TOBConso.ClearDetail;
          if ioerror <> OeOk then
          begin
            break; // on arrete les frais
          end;
        end;
//        readln (InputFile,st_enreg);
        Inc(Noligne);
      end;
      if TOBConso.detail.count > 0 then
      begin
        Ioerror := TRANSACTIONS(EnregistreConso,0);
        TOBConso.ClearDetail;
      end;
      if IoError = OeOK then
      begin
      	EnregistreFinTrait(noligne-1);
      end;
    FINALLY
      CloseFile(InputFile);
      CloseFile(logfile);
    END;
    if (Ioerror = oeok) then
    begin
    	RenameFile (Nomfic,NewNomFic);
    	RenameFile (Nomlog,NewNomlog);
    end;
  end;

end;

procedure TOF_IMPORTINNOVPRO.EnregistreInfos(st_enreg: string ; Ligne : integer);
var TOBD : TOB;
begin
	TOBD := TOB.Create ('CONSOMMATIONS',TOBConso,-1);
  TOBD.InitValeurs;

  AddLesSupLignesConso (TOBD);

  TOBD.AddChampSupValeur ('ERREUR',-1);
	TOBD.addChampSupValeur('NOLIGNE',inttostr(Ligne));
  TOBD.putValue('BCO_AFFAIRE',GetInfosEnreg(st_enreg,TTieChantier));
  TOBD.putValue('BCO_PHASETRA',GetInfosEnreg(st_enreg,TTiePhase));
  TOBD.putValue('BCO_DATEMOUV',GetInfosEnreg(st_enreg,TTieDateMouv));
  TOBD.putValue('BCO_NATUREMOUV',GetInfosEnreg(st_enreg,TTieNatureMouv));
  TOBD.putValue('BCO_RESSOURCE',GetInfosEnreg(st_enreg,TTieRessource));
  TOBD.putValue('BCO_CODEARTICLE',GetInfosEnreg(st_enreg,TTiePrestation));
  TOBD.putValue('BCO_LIBELLE',GetInfosEnreg(st_enreg,TTieDesignation));
  TOBD.putValue('BCO_QUANTITE',GetInfosEnreg(st_enreg,TTieQTE));
  TOBD.putValue('BCO_QUALIFQTEMOUV',GetInfosEnreg(st_enreg,TTieUnite));
  TOBD.putValue('BCO_DPA',GetInfosEnreg(st_enreg,TTiePUA));
  TOBD.putValue('BCO_DPR',GetInfosEnreg(st_enreg,TTiePUR));
  TOBD.putValue('BCO_TYPEHEURE',GetInfosEnreg(st_enreg,TTieTypeHeure));
  ControleDonnee (TOBD);
  if TOBD.GetValue('ERREUR')=-1 then
  begin
  	FinaliseConsommation(TOBD);
  end;
end;


function TOF_IMPORTINNOVPRO.GetInfosEnreg(st_enreg: string;typeInfo: TTypeInfosEnreg): variant;
	function SetDate (LaDate : string) : string;
  var AA,MM,JJ : string;
  begin
  	JJ := Copy(laDate,1,2);
    MM := Copy(LADate,3,2);
    AA := Copy (LADate,5,4);
  	result := format ('%s/%s/%s',[jj,mm,aa]);
  end;
begin
	case TypeInfo of
  	TtieChantier :		begin
                    		Result := Copy(st_enreg,1,17);
									 		end;
  	TtiePhase :				begin
                  			Result := trim(Copy(st_enreg,18,35));
                			end;
  	TtieDateMouv :		begin
                  			Result := StrToDate(SetDate (Copy(st_enreg,53,8)));
                			end;
  	TTieNatureMouv : 	begin
    										Result := trim(Copy(st_enreg,61,3));
                		 	end;
  	TTieRessource : 	begin
    										Result := trim(Copy(st_enreg,64,17));
    								  end;
  	TTiePrestation :  begin
    									  Result := trim(Copy(st_enreg,81,18));
    								  end;
  	TTieDesignation : begin
    										Result := trim(Copy(st_enreg,99,35));
    								  end;
  	TTieQTE : 				begin
    										Result := Valeur(trim(Copy(st_enreg,134,12)));
    								  end;
  	TTieUnite :				begin
    										Result := trim(Copy(st_enreg,146,3));
    								  end;
  	TTiePUA :					begin
    										Result := Valeur(trim(Copy(st_enreg,149,12)));
    								  end;
  	TTiePUR :					begin
    										Result := Valeur(trim(Copy(st_enreg,161,12)));
    								  end;
  	TTieTypeHeure :		begin
    										Result := trim(Copy(st_enreg,173,3));
    								  end;
  end;
end;

procedure TOF_IMPORTINNOVPRO.EnregistreConso;
var Indice : integer;
		TOBD : TOB;
begin
//
	for Indice := 0 to TOBConso.detail.count -1 do
  begin
  	TOBD := TOBConso.detail[Indice];
  	if TOBD.GetValue('ERREUR') = -1 then
    begin
    	if not TOBD.InsertDB (nil) then V_PGI.ioError := oeUnknown;
      EnregistreMvt(TOBD);
  		inc(NbTraiteOk);
    end;
  end;
end;

procedure TOF_IMPORTINNOVPRO.ControleDonnee(TOBD: TOB);
var PrestationDefaut, CodeRessource : string;
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
     (TOBD.GetValue('BCO_NATUREMOUV') <> 'FRS') then
  begin
  	TOBD.putValue('ERREUR',TcdErrNatureMouv);
    SetErreur (TOBD);
    Exit;
  end;
  if TOBD.GetValue('BCO_RESSOURCE') = '' then
  begin
  	TOBD.putValue('ERREUR',TcdErrRessource);
    SetErreur (TOBD);
    Exit;
  end;
  CodeRessource := TOBD.GetValue('BCO_RESSOURCE');
  if OKRessource (nil,CodeRessource,'',PrestationDefaut,false) <> 1 then
  begin
  	TOBD.putValue('ERREUR',TcdErrRessource);
    SetErreur (TOBD);
    Exit;
  end;
  if (TOBD.GetValue('BCO_CODEARTICLE')<>'') and (not IsExisteArticle(TOBD)) then
  begin
  	TOBD.putValue('ERREUR',TcdErrCodeArticle);
    SetErreur (TOBD);
    Exit;
  end;
  // si le code prestation n'est pas renseigné, on récupère celui-ci de la ressource
  if TOBD.GetValue('BCO_CODEARTICLE')='' then
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

procedure TOF_IMPORTINNOVPRO.FinaliseConsommation(TOBD: TOB);
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
    CoefMajoration : double;
    QQ : Tquery;
    Sql : string;
begin
  CoefMajoration := 1;
  if (GetNumUniqueConso (NumConso) <> Gncabort) then
  begin
    TOBres := TOBressources.findFirst(['ARS_RESSOURCE'],[TOBD.getValue('BCO_RESSOURCE')],true);
    if TOBREs = nil then TOBres := AddRessource(TOBD);

    // Si le code article n'est pas renseigné, on récupère la prestation par défaut de la ressource
    if TOBD.GetValue('BCO_CODEARTICLE') = '' then
      TOBD.PutValue('BCO_ARTICLE', TOBres.GetValue('ARS_ARTICLE'))
    else
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
    TOBD.SetString('BCO_FAMILLETAXE1',GetFamilleTaxe(TOBD.GetString('BCO_ARTICLE')));
    TOBD.PutValue('BCO_DATETRAVAUX',ConstitueDateDebutTravaux(TOBD));

    // Découpage de la date
    DateMouv:=StrToDate(TOBD.GetValue('BCO_DATEMOUV'));
    DecodeDate(DateMouv, Annee, Mois, Jour);
    // Recherche du Numéro de semaine
    Semaine := NumSemaine(DateMouv);

    TOBD.PutValue('BCO_MOIS', Mois);
    TOBD.PutValue('BCO_FACTURABLE', 'N');
    TOBD.PutValue('BCO_VALIDE', 'X');
    TOBD.PutValue('BCO_SEMAINE', Semaine);
    if TOBD.GetValue('BCO_TYPEHEURE') <> '' then
    begin
      // Gestion du type d'heure avec coef majorateur
      Sql := 'SELECT CC_ABREGE FROM CHOIXCOD WHERE CC_TYPE="ATH" AND CC_CODE="'+TOBD.GetValue('BCO_TYPEHEURE')+'"';
      QQ := OpenSql (SQl,true,-1,'',true);
      if not QQ.eof then
      begin
        if (QQ.fields[0].AsString <> '') and (IsNumeric (QQ.fields[0].AsString)) then
        begin
          CoefMajoration := 1 + (VALEUR(QQ.fields[0].AsString)/100);
        end;
      end;
      ferme (QQ);
    end;

    // récupération du type de resssource
		TypRes:=GetTypeRessource (TOBD.GetValue('BCO_RESSOURCE'));

    if (TOBD.getValue('BCO_DPA')=0) then
    begin
    	if (TOBD.GetValue('BCO_NATUREMOUV')='MO') or
      	 (TOBD.GetValue('BCO_NATUREMOUV')='RES') or
         ((TOBD.GetValue('BCO_NATUREMOUV')='EXT') and ((TypRes='INT') or (TypRes='ST'))) then
      begin
         SetValoFromRessource (TOBD,TOBREs,[TmvPa],Coefmajoration);
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
         SetValoFromRessource (TOBD,TOBREs,[TmvPR],Coefmajoration);
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
    TOBD.putValue('BCO_INTEGRNUMFIC',Numintegr);
    calculeLaLigne (TOBD);
  end else
  begin
  	TOBD.putValue('ERREUR',TcdErrNumConso);
    SetErreur (TOBD);
    exit;
  end;
end;


function TOF_IMPORTINNOVPRO.IsExisteArticle (TOBD : TOB) : boolean;
var Q : Tquery;
begin
  Q := OpenSql ('SELECT GA_ARTICLE FROM ARTICLE  WHERE GA_CODEARTICLE="'+
                 TOBD.GetValue('BCO_CODEARTICLE') + '" AND GA_STATUTART <> "DIM" ',true,-1, '', True);
  result := not Q.eof;
  if result then TOBD.putValue('BCO_ARTICLE',Q.findfield('GA_ARTICLE').asString);
  ferme (Q);
end;


function TOF_IMPORTINNOVPRO.OKTypeHeure(Typeheure : string) : boolean;
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

function TOF_IMPORTINNOVPRO.AddRessource(TOBD: TOB): TOB;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="'+TOBD.GetValue('BCO_RESSOURCE')+'"',true,-1, '', True);
  if not QQ.eof then
  begin
  	result := TOB.Create ('RESSOURCE',TOBressources,-1);
  	result.selectdb('',QQ);
  end;
  ferme (QQ);
end;

procedure TOF_IMPORTINNOVPRO.SetErreur(TOBD: TOB);
var messageErr : string;
		TheErreur : string;
begin
  if TOBD.GEtVAlue('ERREUR') = TcdErrNumConso then
  BEGIN
  	TheErreur := 'Problème d''affectation du N° de consommation';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrAffaire then
  BEGIN
  	TheErreur := 'Ce chantier ('+TOBD.GetValue('BCO_AFFAIRE')+') n''existe pas';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrPhase then
  BEGIN
  	TheErreur := 'Cette phase ('+TOBD.GetValue('BCO_PHASETRA')+') sur le chantier ('+TOBD.GetValue('BCO_AFFAIRE')+') n''existe pas';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrNatureMouv then
  BEGIN
  	TheErreur := 'La nature de mouvement ('+TOBD.GetValue('BCO_NATUREMOUV')+')  n''existe pas';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrRessource then
  BEGIN
  	TheErreur := 'La ressource ('+TOBD.GetValue('BCO_RESSOURCE')+')  n''existe pas';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrCodeArticle then
  BEGIN
  	TheErreur := 'La prestation ('+TOBD.GetValue('BCO_CODEARTICLE')+')  n''existe pas';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrQte then
  BEGIN
  	TheErreur := 'La quantité n''est pas renseignée';
  END else
  if TOBD.GEtVAlue('ERREUR') = TcdErrTypeHeure then
  BEGIN
  	TheErreur := 'Le Type d''heure ('+TOBD.GetValue('BCO_TYPEHEURE')+')  n''existe pas';
  END;
  MessageErr := Format ('Anomalie ligne %s %s',[TOBD.getValue('NOLIGNE'),TheErreur]);
  Writeln(Logfile,messageErr);
  Rapport.Lines.Add (MessageErr);
end;

procedure TOF_IMPORTINNOVPRO.EnregistreFinTrait(NbrLigne: integer);
var MessageErr : string;
		TOBInfo : TOB;
begin
  MessageErr := Format ('FIN DE TRAITEMENT %d/%d',[NbTraiteOk,NbrLigne]);
  Writeln(Logfile,messageErr);
  Rapport.Lines.Add (MessageErr);
  TOBInfo := TOB.Create ('BENTETEINTEGR',nil,-1);
  TOBInfo.putValue('BEI_INTEGRNUMFIC',Numintegr);
  TOBInfo.putValue('BEI_DATEINTEGR',DateIntegre);
  TOBInfo.InsertDB  (nil);
  TOBInfo.free;
  PGIBox('Traitement terminé',ecran.caption);
end;

function TOF_IMPORTINNOVPRO.AddArticle(TOBD: TOB): TOB;
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

procedure TOF_IMPORTINNOVPRO.EnregistreMvt(TOBD: TOB);
var messageErr  : string;
begin
  MessageErr := Format ('Enregistrement %s Mvt: %s pour le chantier %s et phase %s par la ressource %s et la prestation %s',
  											[TOBD.getValue('NOLIGNE'),TOBD.getValue('BCO_NATUREMOUV'),
                         TOBD.getValue('BCO_AFFAIRE'),TOBD.getValue('BCO_PHASETRA'),
                         TOBD.getValue('BCO_RESSOURCE'),TOBD.getValue('BCO_CODEARTICLE')]);
  Writeln(Logfile,messageErr);
  Rapport.Lines.Add (MessageErr);
end;

{ TOF_ANNULIMPORT }

procedure TOF_ANNULIMPORT.AnnulationImportConso(Numero: integer);
begin
	if PgiAsk ('Etes-vous sûr(e) de vouloir annuler les consommations apportées par cet import ?',ecran.caption) = MrYes then
  begin
    BEGINTRANS;
    TRY
      ExecuteSQL ('DELETE FROM CONSOMMATIONS WHERE BCO_INTEGRNUMFIC='+InttoStr(Numero));
      ExecuteSql ('DELETE FROM BENTETEINTEGR WHERE BEI_INTEGRNUMFIC='+InttoStr(Numero));
      COMMITTRANS;
      PGIInfo ('Import Annulé',ecran.caption);
    EXCEPT
      ROLLBACK;
    END;
  end;
end;

procedure TOF_ANNULIMPORT.OnArgument(S: String);
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnCancel;
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnClose;
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnDelete;
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnDisplay;
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnLoad;
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnNew;
begin
  inherited;

end;

procedure TOF_ANNULIMPORT.OnUpdate;
begin
  inherited;

end;

(* --- *)

procedure AGLAnnulationImportInnovPro (parms: array of variant; nb: integer) ;
var
  F: TForm;
  OM: TOF;
begin
  F := TForm (Longint (Parms [0] ) ) ;
  if (F is TFMul ) then
    OM := TFMul (F) .LaTOF
  else
    exit;
  if (OM is TOF_ANNULIMPORT) then
    TOF_ANNULIMPORT (OM).AnnulationImportConso (Parms [1])
  else
    exit;
end;

Initialization
  registerclasses ( [ TOF_IMPORTINNOVPRO,TOF_ANNULIMPORT ] ) ;
  RegisterAglProc ('AnnulationImportInnovPro', True, 1, AGLAnnulationImportInnovPro) ;

end.

