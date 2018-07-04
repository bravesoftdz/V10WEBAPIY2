{***********UNITE*************************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 28/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULAFFECH_TOF ()
				 Génération automatique des actions ou reconduction
                 de contrats.
Mots clefs ... : TOF;MULAFFECH_TOF
*****************************************************************}
Unit MULAFFECH_TOF ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOB,
     Ent1,
     EntGc,
     AglInit,
     EntRT,
     SaisUtil,
     AffEcheanceUtil,
     AffaireUtil,
     uTofAfBaseCodeAffaire,
     ParamSoc,
     UTOF;


Type

    TOF_MULAFFECH_TOF = Class (TOF_AFBASECODEAFFAIRE)
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

Private
    BOuvrir     : TToolbarButton97;
    PComplement : TTabsheet;
    MOIS        : THvalComboBox;
    ANNEE       : THEdit;
    CHKDateFact : TCheckBox;
    Titre : string;

    Procedure SetEvent;
    Procedure BOuvrirClick (Sender :TObject);
    procedure MAJReconduction;
    procedure CreationAutoAction(TOBA: TOB);
    procedure GenerAutoReconduction(TOBA: TOB);
    procedure ControleChamp(Champ, Valeur: String);
    procedure MoisAnneeChange(Sender: Tobject);
end;

Implementation
uses dateutils,TiersUtil,UtilTaches;

procedure TOF_MULAFFECH_TOF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.OnArgument (S : String ) ;
Var Critere : String;
	ChampMul : String;
    ValMul : String;
    X : integer;
    CC      : THValComboBox;
begin
	fMulDeTraitement := true;

  Inherited ;
//  FTableName := 'AFFAIRE';
  titre := 'Génération des Reconductions de contrat';

  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
     begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        end
     else
        ChampMul := Critere;
     ControleChamp(ChampMul, ValMul);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
     end;

  //Définition des bouton de la grille
  BOuvrir := TToolbarButton97(GetControl('BOuvrir1'));

  PComplement   := TTabsheet(GetControl('PComplement'));
  PComplement.TabVisible := false;
  SetControlProperty ('AFF_ETATAFFAIRE','Plus', ' AND (CC_LIBRE="BTP" AND CC_CODE<>"TER")');
  MOIS          := THvalComboBox(GetControl('MOIS'));
  MOIS.OnChange := MoisAnneeChange;
  ANNEE         := THedit(GetControl('ANNEE'));
  ANNEE.OnExit  := MoisAnneeChange;

  CHKDateFact   := TCheckBox(GetControl('CHKDATEDEBFAC'));
  ChkDateFact.Checked := True;

  THEdit(GetCOntrol('XX_WHERE')).Text := '';

  SetEvent;
  
  // gestion Etablissement (BTP)
	CC:=THValComboBox(GetControl('BTBETABLISSEMENT')) ;
	if CC<>Nil then
  begin
  	PositionneEtabUser(CC) ;
    if not VH^.EtablisCpta then
    begin
    	if THLabel(GetControl('TBTB_ETABLISSEMENT')) <> nil then THLabel(GetControl('TBTB_ETABLISSEMENT')).Visible := false;
			CC.visible := false;
    end;
	end;

  // gestion Domaine (BTP)
	CC:=THValComboBox(GetControl('AFF_DOMAINE')) ;
	if CC<>Nil then PositionneDomaineUser(CC) ;

  // Pour raffraichir le mul avec l'année passée en paramètre
  MoisAnneeChange(Self);

end ;

Procedure TOF_MULAFFECH_TOF.ControleChamp(Champ : String;Valeur : String);
Begin

  SetControlText('XX_WHERE', '');

  If Champ = 'ETAT' then
  	THValComboBox(GetControl('AFF_ETATAFFAIRE')).Value := Valeur
  Else if Champ = 'ANNEE' then
    THEdit(GetControl('ANNEE')).Text := Valeur
  Else if Champ = 'STATUT' then
     Begin
    if Valeur = 'APP' then
        Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then
      begin
        SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
        SetControlText('AFFAIRE0', 'W');
      end;
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel');
          SetControlText('TAFF_AFFAIRE', 'Appel');
        end
     Else if valeur = 'INT' then
          Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then
      begin
        SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
        SetControlText('AFFAIRE0', 'I');
      end;
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Contrat');
          SetControlText('TAFF_AFFAIRE', 'Contrat');
          end
     Else if valeur = 'AFF' then
          Begin
          if assigned(GetControl('AFF_AFFAIRE0'))  then
            SetControlText('AFF_AFFAIRE0', 'A')
          else if assigned(GetControl('AFFAIRE0')) then
            SetControlText('AFFAIRE0', 'A');
          SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "")');
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Chantier');
          SetControlText('TAFF_AFFAIRE', 'Chantier');
          end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0')) then SetControlText('XX_WHERE', ' AND AFF_AFFAIRE0 IN ("W","A")');
      if assigned(GetControl('AFFAIRE0'))     then SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
    end
     Else if valeur = 'PRO' then
          Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then
      begin
        SetControlText('XX_WHERE', ' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("P")');
        SetControlText('AFFAIRE0', 'P');
      end;
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Appel d''offre');
          SetControlText('TAFF_AFFAIRE', 'Appel d''offre');
          end
     else
          Begin
          SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
          SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Affaire');
          SetControlText('TAF_AFFAIRE', 'Affaire');
          end
  end;


end;

procedure TOF_MULAFFECH_TOF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_MULAFFECH_TOF.BOuvrirClick(Sender: TObject);
begin

  try
  	TRANSACTIONS (MAJReconduction,1);
  finally
  	ttoolbarbutton97(getcontrol('BCherche')).Click;
  end;

end;

procedure TOF_MULAFFECH_TOF.MAJReconduction;
var Q  : TQuery;
    QQ : TQuery;
	i  : integer;
    st : string;
    Codeaffaire : String;
    TOBaffaire,TOBA : TOB;
Begin

  St := 'Confirmez vous la reconduction ?';

  if (TFMul(ecran).Fliste.nbSelected=0) and (not TFMul(Ecran).Fliste.AllSelected)  then exit;

  If (PGIAsk(st,titre)<> mrYes) then exit;

  SourisSablier;

  Q:=TFMul(Ecran).Q;

  // on crée une TOB de toutes les affaires sélectionnées
  TobAffaire := TOB.Create ('Liste des Affaires',NIL, -1);

  try
	if TFMul(ecran).Fliste.AllSelected then
	BEGIN
	  Q.First;
	  while Not Q.EOF do
	  BEGIN
        CodeAffaire := Q.FindField('AFF_AFFAIRE').AsString;
        st := 'SELECT AFF.*,BRE_TYPEACTION, BRE_GENERE, BRE_NBMOIS FROM AFFAIRE AFF LEFT JOIN BRECONDUCTION ON BRE_CODE=AFF.AFF_RECONDUCTION WHERE AFF.AFF_AFFAIRE ="'+ CodeAffaire +'"';
        QQ:=OpenSQL(st, TRUE,-1, '', True);
        if not QQ.eof then
        begin
          TOBA := TOB.Create ('AFFAIRE',TOBaffaire,-1);
          TOBA.selectDb ('',QQ);
          if (not VarIsNull (TOBA.GetValue('BRE_TYPEACTION'))) and (TOBA.GetValue('BRE_TYPEACTION') <> '') then
             CreationAutoAction(TOBA)
          else
             GenerAutoReconduction(TOBA);
        end;
        ferme(QQ);
    	Q.NEXT;
	  END;
  	  TFMul(ecran).Fliste.AllSelected:=False;
    END
    else
	begin
	  for i:=0 to TFMul(ecran).Fliste.nbSelected-1 do
	  begin
	      TFMul(Ecran).Fliste.GotoLeBookmark(i);
        CodeAffaire := TFMul(ecran).Fliste.datasource.dataset.FindField('AFF_AFFAIRE').AsString;
        QQ:=OpenSQL('SELECT AFF.*,BRE_TYPEACTION, BRE_GENERE, BRE_NBMOIS FROM AFFAIRE AFF LEFT JOIN BRECONDUCTION ON BRE_CODE=AFF.AFF_RECONDUCTION WHERE AFF.AFF_AFFAIRE ="'+ CodeAffaire +'"', TRUE,-1, '', True);
        if not QQ.eof then
        begin
          TOBA := TOB.Create ('AFFAIRE',TOBaffaire,-1);
          TOBA.selectDb ('',QQ);
          if (not VarIsNull (TOBA.GetValue('BRE_TYPEACTION'))) and (TOBA.GetValue('BRE_TYPEACTION') <> '') then
             CreationAutoAction(TOBA)
          else
             GenerAutoReconduction(TOBA);
        end;
        ferme(QQ);
	  end;
    end;
  Finally
  end;

  TobA.free;

  //Lancement automatique de la génération du Publipostage
  //LancePublipostage('Reconduction',stMaquette,stDocument,TOBM,Nil,True)

  SourisNormale;

end;

Procedure TOF_MULAFFECH_TOF.CreationAutoAction(TOBA:TOB);
var TOBAction : TOB;
    NbAct     : Integer;
    StSQL     : String;
    Codetiers : String;
    Req       : TQuery;
Begin

  if TOBA = nil then Exit;

  CodeTiers := TiersAuxiliaire(TOBA.GetValue('AFF_TIERS'), False);

//Creation automatique d'un action
  TOBAction := TOB.Create ('ACTIONS',NIL, -1);
  StSQL := 'SELECT MAX(RAC_NUMACTION) FROM ACTIONS WHERE RAC_AUXILIAIRE = "'+ CodeTiers + '"';
  Req := OpenSQL(StSQL, true,-1, '', True);
  if Req.eof then
  	TOBAction.PutValue('RAC_NUMACTION', 1)
  else
  begin
	   NbAct := Req.Fields[0].AsInteger;
     TOBAction.PutValue('RAC_NUMACTION', NbAct+1);
  end;
  ferme(Req);

  TOBAction.PutValue('RAC_LIBELLE', 'Demande de reconduction');
  TOBAction.putValue('RAC_AUXILIAIRE', CodeTiers);
  TOBAction.putValue('RAC_TIERS', TOBA.GetValue('AFF_TIERS'));
  TOBAction.putValue('RAC_TYPEACTION', TOBA.GetValue('BRE_TYPEACTION'));
  TOBAction.putValue('RAC_PRODUITPGI', 'GRC');

  if TOBA.GetValue('AFF_RESPONSABLE')<> '' then
  	TOBAction.putValue('RAC_INTERVENANT', TOBA.GetValue('AFF_RESPONSABLE'))
  else
    TOBAction.putValue('RAC_INTERVENANT',VH_RT.RTResponsable);

  TOBAction.putValue('RAC_INTERVINT',TOBA.getvalue('RAC_INTERVENANT'));
  TOBAction.putValue('RAC_OPERATION','');
  TOBAction.putValue('RAC_NUMCHAINAGE', 0);
  TOBAction.putValue('RAC_NUMACTGEN', 0);
  TOBAction.putValue('RAC_PROJET', '');

  TOBAction.putValue('RAC_DATEACTION', V_PGI.DateEntree);
  TOBAction.putValue('RAC_HEUREACTION', Time);
  TOBAction.putValue('RAC_DUREEACT', 0);
  TOBAction.putValue('RAC_DUREEACTION', 0);

  TOBAction.putValue('RAC_DATEECHEANCE', iDate2099);
  TOBAction.putValue('RAC_ETATACTION', 'PRE');

  TOBAction.putValue('RAC_COUTACTION', 0);

  TOBAction.putValue('RAC_UTILISATEUR', V_PGI.user);
  TOBAction.putValue('RAC_CREATEUR', V_PGI.user);

  TOBAction.putValue('RAC_DATECREATION', V_PGI.DateEntree);
  TOBAction.putValue('RAC_DATEMODIF', V_PGI.DateEntree);

  TOBAction.putValue('RAC_GESTRAPPEL', '-');
  TOBAction.putValue('RAC_CHAINAGE', '---');
  TOBAction.putValue('RAC_MAILENVOYE', '-');
  TOBAction.putValue('RAC_MAILAUTO', '-');

  TOBACTION.PutValue('RAC_AFFAIRE', TOBA.GetValue('AFF_AFFAIRE'));
  TOBACTION.PutValue('RAC_AFFAIRE0', TOBA.GetValue('AFF_AFFAIRE0'));
  TOBACTION.PutValue('RAC_AFFAIRE1', TOBA.GetValue('AFF_AFFAIRE1'));
  TOBACTION.PutValue('RAC_AFFAIRE2', TOBA.GetValue('AFF_AFFAIRE2'));
  TOBACTION.PutValue('RAC_AFFAIRE3', TOBA.GetValue('AFF_AFFAIRE3'));
  TOBACTION.PutValue('RAC_AVENANT', TOBA.GetValue('AFF_AVENANT'));

  if not TOBAction.InsertDB(nil) then V_PGI.IoError := oeUnknown;

  //Mise a jour de la tob de génération des Mailings

  //Maj du code etat de l'affaire-Contrat
  TOBA.PutValue('AFF_ETATAFFAIRE', 'ATT');
  if not TOBA.UpdateDB then V_PGI.IoError := oeUnknown;

  TOBAction.Free;

end;

Procedure TOF_MULAFFECH_TOF.GenerAutoReconduction(TOBA:TOB);
Var TypeModif	:	T_TypeModifAff; // T_TypeModifAff = (tmaMntEch,tmaPourcEch,tmaDate,tmaDFacLiq) ;
		//Res				: TIOErr;
    DEV				: RDEVISE;
//
	  //Year			:	Word;
	  //Month			:	Word;
  	//Day 			: Word;
    yearF     : Word;
    MonthF    : Word;
    DayF      : word;
//
    TypeCalcul    :	String;
    //Choix			    :	string;
    ModeMontantDoc: String;
    MethodeCalcul : string;
//
		//NbIt			    :	Integer;
    NbMois        : integer;
		zmont			    :	double;
    //
		bMtPeriodique : Boolean;
    bPourcentage  : Boolean;
    //bRecalculOK   : Boolean;
    //
    DateFinFac    : TDateTime;
		//DateDebFac    : TDateTime;
	  DateDebut	    :	TDAteTime;
    newDateFin    : TDateTime;
    DateDebGen    : TdateTime;
    DateCal       : TdateTime;
		DateLiquid    :	TDateTime;
    DateResil	    :	TDateTime;
//  DateFin		    :	TDateTime;
//  DateDebCal    : TDateTime;
//
    StSQl         : String;
    QQ            : TQuery;
    TobTaches     : Tob;
    TobTache      : Tob;
    TobTachesD    : Tob;
    fTOBDet       : TOb;
    i             : Integer;
    NoTacheDup    : Integer;
//
Begin

  IF GetControlText('ANNEE')= '' Then exit;

	If TOBA.GETVALUE('BRE_GENERE')='-' then exit;

  bMtPeriodique := true;
  bPourcentage := ((TOBA.GetValue ('AFF_GENERAUTO') = 'POU') or (TOBA.GetValue('AFF_GENERAUTO') = 'POT') ) ;

  //Maj des date de contrat en fonction de la Tacite Reconduction
  TypeCalcul := TOBA.GetValue('AFF_PERIODICITE');
  TypeModif := TypeModif + [tmaDate];

  if ((TOBA.GetValue('AFF_GENERAUTO') = 'POU') or (TOBA.GetValue('AFF_GENERAUTO') = 'POT') ) then
    zmont := double (TOBA.GetValue('AFF_POURCENTAGE') )
  else
    zmont := double (TOBA.GetValue('AFF_MONTANTECHEDEV') ) ;
  //
  newDateFin := IncMonth(StrToDate(DateToStr(TOBA.GETVALUE('AFF_DATEFIN')+1)), TOBA.GetVALUE('BRE_NBMOIS'))-1;
  DateDebGen := TOBA.GETVALUE('AFF_DATEDEBGENER');
  //
  if chkdatefact.Checked then
    DateDebut := IncMonth(StrToDate(DateToStr(DateDebGen+1)), TOBA.GetVALUE('BRE_NBMOIS'))-1
  else
    DateDebut := TOBA.GETVALUE('AFF_DATEDEBGENER');
  //
  DateFinFac := IncMonth(StrToDate(DateToStr(TOBA.GETVALUE('AFF_DATEFINGENER')+1)), TOBA.GetVALUE('BRE_NBMOIS'))-1;
  //
	DeCodeDate (DateFinFac,yearF,MonthF,DayF);
	DateFinFac := EncodeDateTime ( YearF,MonthF,DayF,23,59,59,0);
  //
  DateLiquid := TOBA.GETVALUE('AFF_DATEFACTLIQUID');
  DateResil  := TOBA.GetValue('AFF_DATERESIL');

  ModeMontantDoc := TOBA.getValue('AFF_DETECHEANCE');
	MethodeCalcul := TOBA.getValue('AFF_METHECHEANCE');

  UtilCalculEcheances (TOBA.GetValue('AFF_TOTALHTGLODEV'),
                       TOBA.GetValue('AFF_PROFILGENER'),
                       TypeCalcul,
                       TOBA.GetValue('AFF_AFFAIRE'),
                       TOBA.GetValue('AFF_REPRISEACTIV'),
                       TOBA.GetValue('AFF_TIERS'),
                       TOBA.GetValue('AFF_GENERAUTO'),
                       TypeModif,
                       Integer (TOBA.GetValue('AFF_INTERVALGENER') ),
                       zmont,
                       DateDebut,
                       DateFinFac,
                       DateLiquid,
                       DateDebut,
                       idate2099,
                       DateResil,
			           			 DEV,
                       false,
                       bMtPeriodique,
                       bPourcentage,
                       TOBA.GetValue('AFF_MULTIECHE'),ModeMontantDoc,MethodeCalcul,DateDebut,DateFinFac,TOBA.getValue('AFF_TERMEECHEANCE') ) ;

  TOBA.PutValue('AFF_DATEFIN', NewDateFin);

  if CHKDateFact.checked then TOBA.PutValue('AFF_DATEDEBGENER', DateDebut);

	DeCodeDate (DateFinFac,yearF,MonthF,DayF);
	DateFinFac := EncodeDate ( YearF,MonthF,DayF);

  TOBA.PutValue('AFF_DATEFINGENER', DateFinFac);

  //
  NbMois:= StrToInt(GetParamSoc ('SO_AFALIMCLOTURE'));
  DateCal := IncMonth(NewDateFin, NbMois);
  TOBA.PutValue('AFF_DATECLOTTECH', DateCal);
  NbMois:= StrToInt(GetParamSoc ('SO_AFCALCULFIN'));
  DateCal := IncMonth(NewDateFin, NbMois);
  TOBA.PutValue('AFF_DATELIMITE', DateCal);
  //
  if not TOBA.UpdateDB then V_PGI.IoError := oeUnknown;
  //
  //
  //
  TobTachesD:= Tob.Create('TACHES DUPLIQUEES',nil, -1);
  TobTaches := Tob.Create('Les TACHES',nil, -1);
  StSQl := 'SELECT * FROM TACHE WHERE ATA_AFFAIRE="' + TOBA.GetString('AFF_AFFAIRE') + '" ';
  StSQl := StSQl + '  AND(YEAR(ATA_DATEFINPERIOD)=' + GetControlText('ANNEE') + ') ';
  QQ := OpenSQL(StSQl, False);
  if not QQ.Eof then
  begin
    TobTaches.LoadDetailDB('TACHE', '','',QQ, False);
    For i := 0 to TobTaches.Detail.count -1 do
    begin
      //chargement de l'enregistrement initial
      NoTacheDup := TobTaches.detail[TobTaches.detail.count-1].Getvalue('ATA_NUMEROTACHE');
      NoTacheDup := NoTacheDup + 1;

      //if ClientFerme then PGIBoxAF(TexteMessage[31], ecran.Caption); //Attention, ce client est fermé. Vous ne devriez pas créer de tâches !
      //controle si la tache à dupliquer n'est pas déjà générée

      fTobDet := TOB.Create('TACHE', TobTachesD, -1);
      fTobDet.Dupliquer(TobTaches.detail[i], false, true);

    	DeCodeDate (fTobDet.GetDateTime('ATA_DATEDEBPERIOD'),yearF,MonthF,DayF);
      YearF      := StrToInt(GetControlText('ANNEE'));
      DateDebut  := EncodeDateTime ( YearF,MonthF,DayF,0,0,0,0);
      DeCodeDate (fTobDet.GetDateTime('ATA_DATEFINPERIOD'),yearF,MonthF,DayF);
      if YearF = StrToInt(GetControlText('ANNEE')) then
        YearF := YearF + 1
      else if YearF < StrToInt(GetControlText('ANNEE')) then
        YearF := StrToInt(GetControlText('ANNEE'));
    	DateFinFac := EncodeDateTime ( YearF,MonthF,DayF,0,0,0,0);

      fTobDet.PutValue('ATA_DATEDEBPERIOD', DateToStr(DateDebut));
      fTobDet.PutValue('ATA_DATEFINPERIOD', DateToStr(DateFinFac));
      FtobDet.PutValue('ATA_NUMEROTACHE', NoTacheDup);
      FtobDet.PutValue('ATA_TERMINE', 'X');
      if not FtobDet.InsertOrUpdateDB(false) then V_PGI.IoError := oeUnknown;
      GenereAppel(fTOBDet);
      FreeandNil(fTobDet);
    end;
  end;

  FreeandNil(TobTaches);
  FreeandNil(TobTachesD);


end;

procedure TOF_MULAFFECH_TOF.SetEvent;
begin

  BOuvrir.OnClick := BOuvrirClick;

end;

procedure TOF_MULAFFECH_TOF.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,  Tiers_: THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));  
end;

procedure TOF_MULAFFECH_TOF.MoisAnneeChange(Sender: Tobject);
var stext : string;
		//Year,Month,Jour : word;
begin
  if MOIS.Value = '' then  stext := ''
  else if MOIS.Value = '01' then stext := ' AND (MONTH(AFF_DATEFIN)=1)'
  else if MOIS.Value = '02' then stext := ' AND (MONTH(AFF_DATEFIN)=2)'
  else if MOIS.Value = '03' then stext := ' AND (MONTH(AFF_DATEFIN)=3)'
  else if MOIS.Value = '04' then stext := ' AND (MONTH(AFF_DATEFIN)=4)'
  else if MOIS.Value = '05' then stext := ' AND (MONTH(AFF_DATEFIN)=5)'
  else if MOIS.Value = '06' then stext := ' AND (MONTH(AFF_DATEFIN)=6)'
  else if MOIS.Value = '07' then stext := ' AND (MONTH(AFF_DATEFIN)=7)'
  else if MOIS.Value = '08' then stext := ' AND (MONTH(AFF_DATEFIN)=8)'
  else if MOIS.Value = '09' then stext := ' AND (MONTH(AFF_DATEFIN)=9)'
  else if MOIS.Value = '10' then stext := ' AND (MONTH(AFF_DATEFIN)=10)'
  else if MOIS.Value = '11' then stext := ' AND (MONTH(AFF_DATEFIN)=11)'
  else if MOIS.Value = '12' then stext := ' AND (MONTH(AFF_DATEFIN)=12)';
(*
  if ANNEE.text <> '' then
  begin
    DecodeDate (StrToDate(ANNEE.text),Year,month,Jour);
    stext := sText+' AND (YEAR(AFF_DATEFIN)='+IntTOStr(Year)+')';
  end;
*)
  if ANNEE.text <> '' then
  begin
    stext := sText+' AND (YEAR(AFF_DATEFIN)='+ANNEE.text+')';
  end;
  THEdit(GetCOntrol('XX_WHERE')).Text := stext;

end;

Initialization
  registerclasses ( [ TOF_MULAFFECH_TOF ] ) ;
end.

