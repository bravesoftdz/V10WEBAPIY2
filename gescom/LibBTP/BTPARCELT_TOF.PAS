{***********UNITE*************************************************
Auteur  ...... :  LS
Créé le ...... : 14/02/2010
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPARCELT ()
Mots clefs ... : TOF;BTPARCELT
*****************************************************************}
Unit BTPARCELT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,MainEagl,UtilEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     HTB97,
     utofAfBaseCodeAffaire,
     utilArticle,
     AGLInitGc,
     UTOF,
     affaireutil,
     UserialiseLSE,
     UtilParc;

Type
  TOF_BTPARCELT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	fTOBParcElt,TOBART : TOb;
    fCodeInterneP,fCodeArticleP : string;
    fniv1,fniv2,fniv3,fniv4,fniv5 : string;
    fArticle : string;
    faction : TActionFiche;
    fNiveauEcr,fNiveauTrt : integer;

    //
    BP1_CODEARTICLE,BP1_ARTICLE,LIBELLEPARC,BP1_CODEINTERNE,BP1_CODEDEBLOC,BP1_SERIE,BP1_REFCLIENT : Thedit;
    BP1_ETATPARC : THValComboBox;
    BP1_QTE : THSpinEdit;
    BCP_NBMOISGARANTIE : ThNumEdit;
    BP1_DATEACHAT,BP1_DATEFINGAR,DATEFINCONTRAT,DATERESILIATION,BP1_DATEFINSERIA : THEdit;
    BP1_AFFAIRE,BP1_AFFAIRE1,BP1_AFFAIRE2,BP1_AFFAIRE3,AFF_AFFAIRE0,AVENANT : Thedit;
    BP1_REFFABRICANT,BP1_TYPEARTICLE,BP1_TIERS : Thedit;
    BP1_ISCONTRAT : THCheckbox;
    BP1_CODEVERSION,BCP_ISVERSION,BCP_QTEZERO : Thedit;
    GB_CONTRAT : TgroupBox;
    LVERSION : ThLabel;
    BSELECTAFF1,BDETAILNOMEN : TToolbarButton97;
    LIBCONTRAT,LDATEFINCONTRAT,RESILIE : Thlabel;
    //
    procedure GetComponents;
    procedure SetNewEltParc;
    procedure DecodeArgument (S : string);
    procedure GetEltParc;
    function ContitueWhere: string;
    procedure SetEvents (Actif : boolean=true);
    procedure ISContratClick (Sender : Tobject);
    procedure AffaireChange (Sender : Tobject);
    procedure CodeArticleRech (Sendre : TOBject);
    procedure ArticleExit (Sender : Tobject);
    procedure SetModif;
    procedure SetConsult;
    procedure QteChange (Sender : Tobject);
    procedure DateAchExit (Sener : Tobject);
    procedure SetInfoArticle;
    procedure ShowInfos;
    procedure CODEVERSIONExit (Sender : TObject);
    procedure CODEINTERNEExit (Sender : TObject);
    procedure CalculeFinGarantie;
    procedure rechContrat (Sender : Tobject);
    procedure LectureContrat(Contrat: string);
    procedure AfficheContrat(TobContrat: tob);
    procedure SetNumeroClef;
    function ConstitueWherenew(Niveau: integer): string;
    procedure SetChampsTOB(Niveau: integer);
    procedure SetNiveau;
    function EcritureParc(OnTOB: TOB): boolean;
    procedure incrementeNiveau(OneTOB : TOB);
    procedure infosContrat;
    procedure VoirDetail(Sender: Tobject);
    function ConstitueArg(ActionSuiv: TActionFiche): string;
    procedure CompteNiveau (Article : string; var Niveau : integer);
    procedure RechcodeVersion (Sender : Tobject);
    procedure  ConstitueNomenParc (TOBPere : TOB);
    procedure DateSeriaChange (Sender : Tobject);
  end ;

Implementation

uses TntStdCtrls, DateUtils;

procedure TOF_BTPARCELT.OnNew ;
begin
//  BP1_DATEFINSERIA.Text := '31/12/2099';
  Inherited ;
end ;

procedure TOF_BTPARCELT.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCELT.OnUpdate ;
var Indice : integer;
begin
  Inherited ;
  SetEvents (false);  // désactivation des événements
  if BP1_CODEARTICLE.text = '' then
  begin
  	PgiInfo ('Veuillez renseigner le code élément de parc');
    TFVierge(Ecran).ModalResult := 0;
    BP1_DATEFINGAR.setfocus;
  	SetEvents;
    Exit;
  end;
  if StrToDate(BP1_DATEFINGAR.text) < StrToDAte(BP1_DATEACHAT.text) then
  begin
  	PgiInfo ('La Date de fin de garantie ne peut pas être postérieure à la date d''achat');
    TFVierge(Ecran).ModalResult := 0;
    BP1_DATEFINGAR.setfocus;
  	SetEvents;
    Exit;
  end;
  if BP1_ISCONTRAT.checked then
  begin
  	if BP1_AFFAIRE.Text = '' then
    begin
      PgiInfo ('Veuillez renseigner le numéro de contrat de maintenance');
      TFVierge(Ecran).ModalResult := 0;
      BP1_AFFAIRE1.setfocus;
  		SetEvents;
      Exit;
    end;
  end;
  // La tout est ok -- > on enregistre
  fTOBParcElt.GetEcran(ecran,nil);
  if faction = taCreat then
  begin
  	SetNumeroClef;
    fTOBParcElt.SetAllModifie(true);
  end;
  BEGINTRANS;
  TRY
    if not EcritureParc(fTOBParcElt) then
    begin
      PGIError('Erreur durant l''écriture. Opération abandonnée');
      exit;
    end;
    if faction = taCreat then
    begin
      if fTOBParcElt.getValue('BP1_TYPEARTICLE')='PA2' then
      begin
        // c'est une nomenclature de parc --> on récupère son détail et on l'inclus
        ConstitueNomenParc (fTOBParcElt);
        fTOBParcElt.SetAllModifie(true); 
        For Indice := 0 to fTOBParcElt.detail.count -1 do
        begin
        	if not EcritureParc(fTOBParcElt.detail[Indice]) then
    			begin
      			PGIError('Erreur durant l''écriture des composants. Opération abandonnée');
      			exit;
          end;
        end;
      end;
    end;
  	COMMITTRANS;
  EXCEPT
  	ROLLBACK;
  END;
end ;

procedure TOF_BTPARCELT.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_BTPARCELT.SetNiveau ;
begin
	if Valeur(fNiv5)> 0 then
  begin
  	fNiveauTrt := 5;
	end else if valeur(fniv4) > 0 then
  begin
  	fNiveauEcr := 5;
  	fNiveauTrt := 4;
  end else if valeur(fniv3) > 0 then
  begin
  	fNiveauEcr := 4;
  	fNiveauTrt := 3;
  end else if valeur(fniv2) > 0 then
  begin
  	fNiveauEcr := 3;
  	fNiveauTrt := 2;
  end else if valeur(fniv1) > 0 then
  begin
  	fNiveauEcr := 2;
  	fNiveauTrt := 1;
  end else
  begin
  	fNiveauEcr := 1;
  	fNiveauTrt := 0;
  end;
end;

procedure TOF_BTPARCELT.OnArgument (S : String ) ;
begin
  Inherited ;
  //

  faction := taCreat;
	fTOBParcElt := TOB.Create ('PARCTIERS',nil,-1);
  TOBART := TOB.Create ('ARTICLE',nil,-1);
  GetComponents;
  AFF_AFFAIRE0.text := 'I';
	RESILIE.Caption  := '';
  LIBELLEPARC.Text := '';
  BP1_ISCONTRAT.checked := false;
  BP1_CODEARTICLE.ElipsisButton := true;
  LVERSION.Visible := false;
  BP1_CODEVERSION.visible := false;
  BCP_ISVERSION.text := '-';
  BCP_QTEZERO.text := '-';
  BCP_NBMOISGARANTIE.value := 0;
  //formatage du code contrat
  ChargeCleAffaire(THEDIT(GetControl('AFF_AFFAIRE0')),THEDIT(GetControl('BP1_AFFAIRE1')),
  									THEDIT(GetControl('BP1_AFFAIRE2')), THEDIT(GetControl('BP1_AFFAIRE3')),
                    THEDIT(GetControl('AVENANT')), BSELECTAFF1 , fAction, BP1_AFFAIRE.text, False);

  DecodeArgument (S);
  SetNiveau;
  if faction <> TaCreat then
  begin
    GetEltParc;

    if Faction = taModif then SetModif else SetConsult;
  end else
  begin
  	SetNewEltParc;
  end;
  fTOBParcElt.putEcran(ecran,nil);
  // pour pallier a un grooos bug de CBP --
  //	le putecran ne gère pas les type de champs autres que thedit, THnumedit 
  BP1_QTE.Value := fTOBParcElt.GetValue('BP1_QTE');
  // --
	infosContrat;
  ShowInfos;
  SetEvents;
end ;

procedure TOF_BTPARCELT.OnClose ;
begin
  //
	FreeAndNil(fTOBParcElt);
  FreeAndNil(TOBART);
  Inherited ;
end ;

procedure TOF_BTPARCELT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCELT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPARCELT.DecodeArgument(S: string);
var Theparams : string;
		TheParam,TheChamps,TheValue : string;
begin
  TheParams := S;
  fNiv1:='0';
  fNiv2:='0';
  fNiv3:='0';
  fNiv4:='0';
  fNiv5:='0';
  repeat
  	TheParam := READTOKENST(TheParams);
    if TheParam = '' then break;
    theChamps := READTOKENPipe (TheParam,'=');
    thevalue := TheParam;
    if TheChamps = 'TIERS' then
    begin
    	BP1_TIERS.text := Thevalue;
    end else if TheChamps = 'ARTICLE' then
    begin
    	fArticle := TheValue;
    end else if TheChamps = 'CODEINTERNEP' then
    begin
    	fcodeinterneP := TheValue;
    end else if TheChamps = 'CODEARTICLEP' then
    begin
    	fCodeArticleP := TheValue;
    end else if TheChamps = 'N1' then
    begin
    	fniv1 := TheValue;
    end else if TheChamps = 'N2' then
    begin
    	fniv2 := TheValue;
    end else if TheChamps = 'N3' then
    begin
    	fniv3 := TheValue;
    end else if theChamps = 'N4' then
    begin
    	fniv4 := TheValue;
    end else if theChamps = 'N5' then
    begin
    	fniv5 := TheValue;
    end else if TheChamps = 'ACTION' then
    begin
    	if theValue = 'CONSULTATION' then fAction := taConsult else IF TheValue='MODIFICATION' then fAction := Tamodif;
    end;
  until TheParams = '';
end;

procedure TOF_BTPARCELT.GetComponents;
begin
  BP1_CODEARTICLE := THEdit (GetControl('BP1_CODEARTICLE'));
  BP1_ARTICLE := THEdit (GetControl('BP1_ARTICLE'));
  LIBELLEPARC := ThEdit(GetControl('LIBELLEPARC'));
  BP1_CODEINTERNE := ThEdit(GetControl('BP1_CODEINTERNE'));
  BP1_CODEDEBLOC := ThEdit(GetControl('BP1_CODEDEBLOC'));
  BP1_SERIE := ThEdit(GetControl('BP1_SERIE'));
  BP1_REFCLIENT := ThEdit(GetControl('BP1_REFCLIENT'));
  BP1_REFFABRICANT := ThEdit(GetControl('BP1_REFFABRICANT'));
  BP1_ETATPARC := ThValComboBox(GetControl('BP1_ETATPARC'));
  BP1_QTE := THSpinEdit (getControl('BP1_QTE'));
  BP1_DATEACHAT := ThEdit(GetControl('BP1_DATEACHAT'));
  BP1_DATEFINGAR := ThEdit(GetControl('BP1_DATEFINGAR'));
  DATEFINCONTRAT := ThEdit(GetControl('DATEFINCONTRAT'));
  DATERESILIATION := ThEdit(GetControl('DATERESILIATION'));
  BP1_AFFAIRE := ThEdit(GetControl('BP1_AFFAIRE'));
  BP1_AFFAIRE1 := ThEdit(GetControl('BP1_AFFAIRE1'));
  BP1_AFFAIRE2  := ThEdit(GetControl('BP1_AFFAIRE2'));
  BP1_AFFAIRE3 := ThEdit(GetControl('BP1_AFFAIRE3'));
  AFF_AFFAIRE0 := ThEdit(GetControl('AFF_AFFAIRE0'));
  BP1_ISCONTRAT := THCheckbox(getControl('BP1_ISCONTRAT'));
  GB_CONTRAT := TGroupBox (getControl('GB_CONTRAT'));
  BP1_CODEVERSION := ThEdit(getControl('BP1_CODEVERSION'));
  BCP_NBMOISGARANTIE := THNumEdit(getcontrol('BCP_NBMOISGARANTIE'));
  BCP_ISVERSION := Thedit(GetControl('BCP_ISVERSION'));
  LVERSION := THlabel (GetControl('LVERSION'));
  BSELECTAFF1 := TToolbarButton97 (getControl('BSELECTAFF1'));
  BP1_TYPEARTICLE := THEdit(getControl('BP1_TYPEARTICLE'));
  BP1_TIERS := Thedit(getCOntrol('BP1_TIERS'));
  LIBCONTRAT := Thlabel(getControl('LIBCONTRAT'));
  LDATEFINCONTRAT := Thlabel (GetControl('LDATEFINCONTRAT'));
  RESILIE := Thlabel(getCOntrol('RESILIE'));
  AVENANT := THEdit(getCOntrol('AVENANT'));
  BDETAILNOMEN := TToolBarButton97(GetControl('BDETAILNOMEN'));
  BCP_QTEZERO := THEdit(getControl('BCP_QTEZERO'));
  BP1_DATEFINSERIA := ThEdit(GetControl('BP1_DATEFINSERIA'));
end;

function TOF_BTPARCELT.ContitueWhere : string;
begin
	result := 'P.BP1_TIERS="'+BP1_TIERS.text+'" AND P.BP1_N1='+fniv1+
  				  ' AND P.BP1_N2='+fniv2+' '+
  					'AND P.BP1_N3='+fniv3+
            ' AND P.BP1_N4='+fniv4+' AND P.BP1_N5='+fniv5;
end;

procedure TOF_BTPARCELT.GetEltParc;
var req : string;
		QQ : TQuery;
begin
	Req := 'SELECT P.*,A.GA_LIBELLE AS LIBELLEPARC,AC.*,'+
  			 '(CASE WHEN (SELECT AFF_DATEFIN FROM AFFAIRE WHERE AFF_AFFAIRE=BP1_AFFAIRE) IS NULL '+
         'THEN "1900/01/01" ELSE (SELECT AFF_DATEFIN FROM AFFAIRE WHERE AFF_AFFAIRE=BP1_AFFAIRE) END) '+
         'AS DATEFINCONTRAT,'+
  			 '(CASE WHEN (SELECT AFF_DATERESIL FROM AFFAIRE WHERE AFF_AFFAIRE=BP1_AFFAIRE) IS NULL '+
         'THEN "2099/12/31" ELSE (SELECT AFF_DATERESIL FROM AFFAIRE WHERE AFF_AFFAIRE=BP1_AFFAIRE) END) '+
         'AS DATERESILIATION '+
         'FROM PARCTIERS P '+
         'LEFT JOIN ARTICLE A ON A.GA_ARTICLE=BP1_ARTICLE '+
         'LEFT JOIN ARTICLECOMPPARC AC ON BCP_ARTICLE=BP1_ARTICLE '+
         'WHERE '+ContitueWhere;
  QQ := OpenSql (REq,true,1,'',true);
  fTOBParcElt.SelectDB('',QQ);
  ferme (QQ);
	LectureContrat (fTOBParcElt.GetValue('BP1_AFFAIRE'));
end;

procedure TOF_BTPARCELT.SetNewEltParc;
begin
  fTOBParcElt.putValue('BP1_TIERS',BP1_TIERS.text);
  fTOBParcElt.putValue('BP1_ETATPARC','ES');
  fTOBParcElt.putValue('BP1_DATEACHAT',DateToStr(now));
  fTOBParcElt.putValue('BP1_DATEFINGAR',DateToStr(Now));
  fTOBParcElt.putValue('BP1_DATEFINSERIA',DateToStr(Idate2099));
  DATEFINCONTRAT.text := '01/01/1900';
  DATERESILIATION.text := '31/12/2099';
//	BP1_DATEFINSERIA.Text := '31/12/2099';
  //
end;


procedure TOF_BTPARCELT.SetEvents (Actif : boolean=true);
begin
	if Actif then
  begin
    BP1_CODEARTICLE.OnExit := ArticleExit;
    BP1_CODEARTICLE.OnElipsisClick := CodeArticleRech;
    BP1_ISCONTRAT.OnClick := ISContratClick;
    BP1_AFFAIRE1.OnExit := AffaireChange;
    BP1_AFFAIRE2.OnExit := AffaireChange;
    BP1_AFFAIRE3.OnExit := AffaireChange;
    BP1_QTE.OnChange := QteChange; // Specif pour nous LSE
    BP1_DATEACHAT.OnExit := DateAchExit;
    BP1_CODEVERSION.OnExit := CODEVERSIONExit;
    BP1_CODEVERSION.OnElipsisClick := RechcodeVersion;
    BP1_CODEINTERNE.OnExit := CODEINTERNEExit;
    BSELECTAFF1.onClick := rechContrat;
    BDETAILNOMEN.onclick := VoirDetail;
    BP1_DATEFINSERIA.OnChange := DateSeriaChange;
  end else
  begin
    BP1_CODEARTICLE.OnExit := nil;
    BP1_CODEARTICLE.OnElipsisClick := nil;
    BP1_ISCONTRAT.OnClick := nil;
    BP1_AFFAIRE1.OnExit := nil;
    BP1_AFFAIRE2.OnExit := nil;
    BP1_AFFAIRE3.OnExit := nil;
    BP1_QTE.OnChange := nil; // Specif pour nous LSE
    BP1_DATEACHAT.OnExit := nil;
    BP1_CODEVERSION.OnExit := nil;
    BP1_CODEINTERNE.OnExit := nil;
    BSELECTAFF1.onClick := nil;
    BDETAILNOMEN.onclick := nil;
    BP1_DATEFINSERIA.OnChange := nil;
  end;
end;

procedure TOF_BTPARCELT.ISContratClick(Sender: Tobject);
begin
  BSELECTAFF1.enabled := BP1_ISCONTRAT.Checked;
  BP1_AFFAIRE1.enabled := BP1_ISCONTRAT.Checked;
  BP1_AFFAIRE2.enabled := BP1_ISCONTRAT.Checked;
  BP1_AFFAIRE3.enabled := BP1_ISCONTRAT.Checked;
  DATEFINCONTRAT.visible := BP1_ISCONTRAT.Checked;
  if (not BP1_ISCONTRAT.checked) then
  begin
  	BP1_AFFAIRE.text := '';
  	BP1_AFFAIRE1.text := '';
  	BP1_AFFAIRE2.text := '';
  	BP1_AFFAIRE3.text := '';
    AVENANT.Text := '';
  end;
end;

procedure TOF_BTPARCELT.AffaireChange(Sender: Tobject);
begin

end;

procedure TOF_BTPARCELT.CodeArticleRech(Sendre: TOBject);
var stWhere,stFiche : string;
		ART : THCritMaskEdit;

begin
	ART := THCritMaskEdit.Create (ecran); ART.Visible := false;
  ART.Text := '';
  //
  stFiche := 'BTARTPARC_RECH';
	stWhere := GetTypeArticleParc;
  StWhere := 'GA_CODEARTICLE=' + Trim (Copy (ART.Text, 1, 18))+';XX_WHERE=AND '+stWhere;
  DispatchRecherche (ART, 1, '',stWhere, stFiche);
	if ART.Text <> '' then
  	BP1_CODEARTICLE.Text := Trim(copy(ART.Text, 0, Length(ART.Text) - 1));
  ART.free;
end;

procedure TOF_BTPARCELT.ArticleExit(Sender: Tobject);
var QQ : Tquery;
		MaxNiveau : integer;
begin
	QQ := openSql ('SELECT * FROM ARTICLE '+
  							'LEFT JOIN ARTICLECOMPPARC ON BCP_ARTICLE=GA_ARTICLE '+
                'WHERE GA_CODEARTICLE="'+BP1_CODEARTICLE.text+'"',true,1,'',true);
	if not QQ.eof then
  begin
  	TOBART.SelectDB('',QQ);
  	ferme (QQ);
    if (TOBARt.GetValue('GA_TYPEARTICLE')='PA2') then
    begin
    	if (fNiveauTrt = 5) then
      begin
        PgiInfo ('Le nombre maximum de niveau (5) est atteind.#13#10 Vous ne pouvez pas ajouter de nomenclature');
        BP1_CODEARTICLE.Text := '';
        BP1_CODEARTICLE.SetFocus;
        exit;
      end else
      begin
      	MaxNiveau := 0;
      	CompteNiveau (TOBArt.getValue('GA_CODEARTICLE'),MaxNiveau);
        if fNiveauEcr + MaxNiveau > 5 then
        begin
          PgiInfo ('Le nombre maximum de niveau (5) est atteind.#13#10 Vous ne pouvez pas ajouter de nomenclature');
          BP1_CODEARTICLE.Text := '';
          BP1_CODEARTICLE.SetFocus;
          exit;
        end;
      end;
    end;
    SetInfoArticle;
  end else
  begin
  	ferme (QQ);
    PgiError ('Cet Article n''existe pas');
    SetNewEltParc;
    BP1_CODEARTICLE.SetFocus;
  end;
end;

procedure TOF_BTPARCELT.SetModif;
begin
	BP1_CODEARTICLE.Enabled := false;
end;

procedure TOF_BTPARCELT.SetConsult;
var Indice : integer;
begin
  GB_CONTRAT.Enabled := false;
  for Indice := 0 to ecran.ComponentCount -1 do
  begin
  	if (ecran.Components [Indice] is Thedit) or
    	 (ecran.Components [Indice] is ThValComboBox) or
			 (ecran.Components [Indice] is ThSpinEdit) then
    begin
  		TWinControl (ecran.Components [Indice]).Enabled := false;
    end;
  end;
  TToolBarButton97(getControl('Bvalider')).visible := false;
end;

procedure TOF_BTPARCELT.QteChange(Sender: Tobject);
var retour : string;
begin
  retour := GetSeriaLSE(fCodeInterneP,BP1_CODEINTERNE.text,BP1_CODEVERSION.text,BP1_QTE.value,strToDate(BP1_DATEFINSERIA.text));
  if retour <> '' then BP1_CODEDEBLOC.Text := retour;
end;

procedure TOF_BTPARCELT.DateAchExit(Sener: Tobject);
begin
  CalculeFinGarantie;
end;

procedure TOF_BTPARCELT.SetInfoArticle;
begin
  LIBELLEPARC.Text := TOBART.GetValue('GA_LIBELLE');
  BP1_CODEINTERNE.Text := TOBART.GetValue('BCP_CODEINTERNE');
  BP1_TYPEARTICLE.Text := TOBART.GetValue('GA_TYPEARTICLE');
  BP1_ARTICLE.Text := TOBART.GetValue('GA_ARTICLE');
  BCP_ISVERSION.Text := TOBART.GetValue('BCP_ISVERSION');
  BCP_QTEZERO.Text := TOBART.GetValue('BCP_QTEZERO');
  BCP_NBMOISGARANTIE.value := TOBART.GetValue('BCP_NBMOISGARANTIE');
  BP1_REFFABRICANT.text := TOBART.GetValue('BCP_REFFABRICANT');
  CalculeFinGarantie;
  ShowInfos;
end;

procedure TOF_BTPARCELT.ShowInfos;
begin
  if BCP_ISVERSION.Text = 'X' then
  begin
  	LVERSION.visible := true;
    BP1_CODEVERSION.visible := true;
  end else
  begin
  	LVERSION.visible := false;
    BP1_CODEVERSION.visible := false;
  end;
  if BCP_QTEZERO.text = 'X' then
  begin
  	BP1_QTE.Value := 0;
  	BP1_QTE.visible := false;
  end else
  begin
  	BP1_QTE.visible := true;
  end;
  BSELECTAFF1.enabled := BP1_ISCONTRAT.Checked;
  BP1_AFFAIRE1.enabled :=BP1_ISCONTRAT.Checked;
  BP1_AFFAIRE2.enabled := BP1_ISCONTRAT.Checked;
  BP1_AFFAIRE3.enabled := BP1_ISCONTRAT.Checked;
  DATEFINCONTRAT.visible := BP1_ISCONTRAT.Checked;
  BDETAILNOMEN.visible := (BP1_TYPEARTICLE.text='PA2') and (faction <> TaCreat);
end;

procedure TOF_BTPARCELT.CODEVERSIONExit(Sender: TObject);
var retour : string;
begin
  retour := GetSeriaLSE(fCodeInterneP,BP1_CODEINTERNE.text,BP1_CODEVERSION.text,BP1_QTE.value,StrToDate(BP1_DATEFINSERIA.text));
  if retour <> '' then BP1_CODEDEBLOC.Text := retour;
end;

procedure TOF_BTPARCELT.CalculeFinGarantie;
begin
	BP1_DATEFINGAR.text := DateToStr(PlusDate(StrToDate(BP1_DATEACHAT.text),StrToInt(floattoStr(BCP_NBMOISGARANTIE.value)),'M' ));
end;

procedure TOF_BTPARCELT.CODEINTERNEExit(Sender: TObject);
var retour : string;
begin
  retour := GetSeriaLSE(fCodeInterneP,BP1_CODEINTERNE.text,BP1_CODEVERSION.text,BP1_QTE.value,StrToDate(BP1_DATEFINSERIA.text));
  if retour <> '' then BP1_CODEDEBLOC.Text := retour;
end;

procedure TOF_BTPARCELT.LectureContrat (Contrat : string);
Var TobContrat : TOB;
    Req,CodeContrat        : String;
Begin
  CodeContrat := READTOKENST(Contrat);
  if CodeContrat = '' then exit;
  Req := '';

  Req := 'SELECT * FROM AFFAIRE ';
  Req := Req + 'WHERE AFF_AFFAIRE ="' + CodeContrat + '"';

  Req := Req + 'AND AFF_TIERS ="' + BP1_TIERS.text + '"';

  TobContrat := Tob.Create('LesContrats',Nil, -1);
  TobContrat.LoadDetailDBFromSQL('AFFAIRE',req,false);

  if TobContrat.Detail.Count = 1 then AfficheContrat(TobContrat);

  TobContrat.free;

End;

procedure TOF_BTPARCELT.infosContrat;
begin
  if StrToDate(DATERESILIATION.Text)  <> idate2099 then
  begin
    if StrToDate(DATERESILIATION.Text) < date then
    Begin
			RESILIE.caption := 'Contrat résilié depuis le'+ DATERESILIATION.Text;
    End;
  end;

  if StrToDate(DATEFINCONTRAT.Text)  <> iDate1900 then
  begin
    if StrToDate(DATEFINCONTRAT.text) < date then
    Begin
       RESILIE.caption := 'Contrat arrivé à terme le '+ DATEFINCONTRAT.text;
    end;
  end;
end;

Procedure TOF_BTPARCELT.AfficheContrat(TobContrat : tob);
Var A1      : String;
    A2      : String;
    A3      : String;
Begin


  A1 := TobContrat.Detail[0].GetValue('AFF_AFFAIRE1');
  A2 := TobContrat.Detail[0].GetValue('AFF_AFFAIRE2');
  A3 := TobContrat.Detail[0].GetValue('AFF_AFFAIRE3');

  BP1_AFFAIRE.text := TobContrat.Detail[0].GetValue('AFF_AFFAIRE');
  AFF_AFFAIRE0.text := TobContrat.Detail[0].GetValue('AFF_AFFAIRE0');
  BP1_AFFAIRE1.text := A1;
  BP1_AFFAIRE2.text:=A2;
  BP1_AFFAIRE3.text:= A3;
  AVENANT.Text := TobContrat.Detail[0].GetValue('AFF_AVENANT');

  DATEFINCONTRAT.text := TobContrat.Detail[0].GetValue('AFF_DATEFIN');
  DATERESILIATION.text := TobContrat.Detail[0].GetValue('AFF_DATERESIL');


End;

procedure TOF_BTPARCELT.rechContrat(Sender: Tobject);
var Contrat : string;
begin
   Contrat := AglLanceFiche('BTP','BTAFFAIRE_MUL','AFF_AFFAIRE0=I','','AFF_TIERS=' +BP1_TIERS.Text + ';ETAT=ENC;STATUT=INT;ACTION=RECH');   // Affaires
   if Contrat <> '' then Lecturecontrat (contrat);
end;

function TOF_BTPARCELT.ConstitueWherenew (Niveau : integer) : string;
begin
	result := 'BP1_TIERS="'+fTOBParcElt.getValue('BP1_TIERS')+'"';
	if Niveau = 5 then
  begin
  	result := result + ' AND BP1_N1='+fniv1+
    									 ' AND BP1_N2='+fniv2+
    									 ' AND BP1_N3='+fniv3+
    									 ' AND BP1_N4='+fniv4;
  end else if Niveau = 4 then
  begin
  	result := result + ' AND BP1_N1='+fniv1+
    									 ' AND BP1_N2='+fniv2+
    									 ' AND BP1_N3='+fniv3;
  end else if Niveau = 3 then
  begin
  	result := result + ' AND BP1_N1='+fniv1+
    									 ' AND BP1_N2='+fniv2;
  end else if Niveau = 2 then
  begin
  	result := result + ' AND BP1_N1='+fniv1;
  end;
end;

procedure TOF_BTPARCELT.SetChampsTOB(Niveau : integer);
begin
	if Niveau = 5 then
  begin
		fTOBParcElt.putValue('BP1_N1',fniv1);
		fTOBParcElt.putValue('BP1_N2',fniv2);
		fTOBParcElt.putValue('BP1_N3',fniv3);
		fTOBParcElt.putValue('BP1_N4',fniv4);
  end else if Niveau = 4 then
  begin
		fTOBParcElt.putValue('BP1_N1',fniv1);
		fTOBParcElt.putValue('BP1_N2',fniv2);
		fTOBParcElt.putValue('BP1_N3',fniv3);
  end else if Niveau = 3 then
  begin
		fTOBParcElt.putValue('BP1_N1',fniv1);
		fTOBParcElt.putValue('BP1_N2',fniv2);
  end else if Niveau = 2 then
  begin
		fTOBParcElt.putValue('BP1_N1',fniv1);
  end;
end;

procedure TOF_BTPARCELT.SetNumeroClef;
var req : string;
		IndiceNiv : integer;
		ChampsNiv : string;
    QQ : TQuery;
begin
  SetChampsTOB(fNiveauEcr);
  ChampsNiv := 'BP1_N'+Inttostr(fNiveauEcr);
  IndiceNiv := 1;
  Req := 'SELECT MAX('+ChampsNiv+') AS MAXI FROM PARCTIERS WHERE '+ConstitueWherenew (fNiveauEcr);
  QQ := OpenSql (req,true,1,'',true);
  if not QQ.eof then
  begin
  	IndiceNiv := QQ.findField('MAXI').AsInteger+1;
  end;
  ferme (QQ);
  fTOBParcElt.putValue(ChampsNiv,indiceNiv);
end;


function TOF_BTPARCELT.EcritureParc (OnTOB : TOB) : boolean;
var Nbredo : integer;
begin
	result := false;
  Nbredo := 1;
  repeat
    if faction = taCreat then
    begin
    	result := OnTOB.InsertDB(nil,false);
      if not result then incrementeNiveau(OnTOB);
    end else
    begin
    	result := ONTOB.UpdateDB(false);
    end;
    if not result then BEGIN {evite de saturer le serv SQL} sleep(500); inc(nbredo); END;
  until (Nbredo > 100) or (result);
end;

procedure TOF_BTPARCELT.incrementeNiveau(OneTOB: TOB);
var ChampsNiv : string;
begin
  ChampsNiv := 'BP1_N'+Inttostr(fNiveauEcr);
  OneTOB.putValue(ChampsNiv,OneTOB.getValue(ChampsNiv)+1);
end;


function TOF_BTPARCELT.ConstitueArg(ActionSuiv : TActionFiche): string;
begin
  result := 'TIERS='+BP1_TIERS.Text+';CODEARTICLEP='+BP1_CODEARTICLE.Text+';CODEINTERNEP='+BP1_CODEINTERNE.text;
  if fNiveauTrt = 5 then
  begin
  result := result +
            ';N5='+fniv5+
            ';N4='+fniv4+
            ';N3='+fniv3+
            ';N2='+fniv2+
            ';N1='+fniv1;
  end else if fNiveauTrt = 4 then
  begin
    result := result +
              ';N4='+fniv4+
              ';N3='+fniv3+
              ';N2='+fniv2+
              ';N1='+fniv1;
  end else if fNiveauTrt = 3 then
  begin
    result := result +
              ';N3='+fniv3+
              ';N2='+fniv2+
              ';N1='+fniv1;
  end else if fNiveauTrt = 2 then
  begin
    result := result +
              ';N2='+fniv2+
              ';N1='+fniv1;
  end else if fNiveauTrt = 1 then
  begin
    result := result +
              ';N1='+fniv1;
  end;
  if ActionSuiv=TaConsult then
  	result := result +';ACTION=CONSULTATION'
  else if ActionSuiv = Tamodif then
  	result := result +';ACTION=MODIFICATION'
  else result := result +';ACTION=CREATION';
end;

procedure TOF_BTPARCELT.VoirDetail (Sender : Tobject);
var starg : string;
begin
  StArg := ConstitueArg (faction);
  AGLLanceFiche ('BTP','BTPARCTIER','','',stArg + ';MONOFICHE');
end;

procedure TOF_BTPARCELT.CompteNiveau(Article: string; var Niveau : integer);
var QQ : TQuery;
		TOBNomen : TOB;
    indice : integer;
    nivSuiv : integer;
begin
	TOBNomen := TOB.Create ('NOMENLIG',nil,-1);
  QQ := OpenSql ('SELECT GNL_NUMLIGNE,GNL_CODEARTICLE,GNL_QTE,GA_TYPEARTICLE FROM NOMENLIG '+
                 'LEFT JOIN ARTICLE ON GA_CODEARTICLE=GNL_CODEARTICLE '+
                 'WHERE GNL_NOMENCLATURE="'+Article+'"',true,-1,'',true);
  if not QQ.eof then
  begin
    TOBNomen.LoadDetailDB('NOMENLIG','','',QQ,false);
  end;
  ferme (QQ);
  if TOBNomen.detail.count > 0 then
  begin
  	inc(niveau);
    for Indice := 0 to TOBNomen.detail.count - 1 do
    begin
      if TOBNomen.detail[Indice].getValue('GA_TYPEARTICLE')='PA2' then
      begin
        nivsuiv := niveau;
        CompteNiveau (TOBNomen.detail[Indice].getValue('GA_CODEARTICLE'),nivsuiv);
        if nivsuiv > niveau then niveau := nivsuiv;
      end;
    end;
  end;
  TOBNomen.free;
end;

procedure TOF_BTPARCELT.RechcodeVersion(Sender: Tobject);
var retour : string;
begin
  retour := AglLAnceFiche('BTP','BPARCVERSIONS_MUL','CRET=SELECTION;BVP_CODEARTICLE='+BP1_CODEARTICLE.text,'','ACTION=MODIFICATION');
  if retour <> '' then BP1_CODEVERSION.Text := retour;
end;

procedure TOF_BTPARCELT.ConstitueNomenParc(TOBpere : TOB);
var CodeArticle,CodeDetail : string;
		TOBDetail,TOBNomen,TOBArticle : TOB;
    QQ : TQuery;
    Indice : integer;
    Qte : double;
begin
	CodeArticle := TOBpere.getValue('BP1_CODEARTICLE');
  TOBNomen := TOB.Create ('NOMENLIG',nil,-1);
  TOBArticle := TOB.Create ('ARTICLE',nil,-1);
  TRY
    QQ := OpenSql ('SELECT GNL_NUMLIGNE,GNL_CODEARTICLE,GNL_QTE,GA_TYPEARTICLE FROM NOMENLIG '+
                   'LEFT JOIN ARTICLE ON GA_CODEARTICLE=GNL_CODEARTICLE '+
                   'WHERE GNL_NOMENCLATURE="'+CodeArticle+'"',true,-1,'',true);
    if not QQ.eof then
    begin
      TOBNomen.LoadDetailDB('NOMENLIG','','',QQ,false);
    end;
    ferme (QQ);
    for Indice := 0 to TOBNomen.detail.count -1 do
    begin
      TOBArticle.InitValeurs;
      CodeDetail := TOBNomen.detail[Indice].getValue('GNL_CODEARTICLE');
      Qte := TOBNomen.detail[Indice].getValue('GNL_QTE');
      QQ := openSql ('SELECT * FROM ARTICLE '+
                  'LEFT JOIN ARTICLECOMPPARC ON BCP_ARTICLE=GA_ARTICLE '+
                  'WHERE GA_CODEARTICLE="'+CodeDetail+'"',true,1,'',true);
      TOBArticle.selectDb ('ARTICLE',QQ);
      ferme (QQ);
      TOBDetail := TOB.Create ('PARCTIERS',TOBpere,-1);
      ConstitueEltParcFromPere (TOBpere,TOBDetail,TOBArticle,Qte,Indice);
      if TOBNomen.detail[Indice].getValue('GA_TYPEARTICLE')='PA2' then
      begin
        ConstitueNomenParc (TOBDetail);
      end;
    end;
  FINALLY
    freeAndNil (TOBArticle);
    FreeAndNil(TOBNomen);
  END;
end;

procedure TOF_BTPARCELT.DateSeriaChange(Sender: Tobject);
var retour : string;
		yy,dd,mm : Word;
begin
  TRY
  	DecodeDate(StrToDate(BP1_DATEFINSERIA.text),yy,mm,dd);
  	if not IsValidDate (yy,mm,dd) then Exit;
  except
    BP1_CODEDEBLOC.Text := '';
    Exit;
  end;
  retour := GetSeriaLSE(fCodeInterneP,BP1_CODEINTERNE.text,BP1_CODEVERSION.text,BP1_QTE.value,StrToDate(BP1_DATEFINSERIA.text));
  if retour <> '' then BP1_CODEDEBLOC.Text := retour;
end;

Initialization
  registerclasses ( [ TOF_BTPARCELT ] ) ;
end.
