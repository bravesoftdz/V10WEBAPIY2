unit UTofAfPrepFact;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, windows,
      HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,UTob,Grids,EntGC,
      utilarticle,Dicobtp,AFPrepFact,Facture,Hstatus,TiersUtil,Ent1, CBPPath,
{$IFDEF EAGLCLIENT}
      MaineAGL, eMul,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,Mul,EdtEtat,HDB,
{$ENDIF}
	  SaisUtil,
      UtofAfFactDecideur,  // a laisser car tof appele depuis script , le source doit être dans le projet
      Menus,UtofAftableaubord, UtilGa,
      utofAfBaseCodeAffaire,M3FP,TraducAffaire,UtilMulTrt,UtofAfPiece,paramsoc
      ,factaffaire, Shellapi, HTB97,Ed_tools,uEntCommun;

      Function CtrlDate(Zdat : string) : integer;

Type
     TOF_AFPREPFACT = Class (TOF_AFBASECODEAFFAIRE)
     public
     topprem : boolean;
        procedure OnLoad ; override ;
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
        Procedure MenuPopOnClick(Sender:TobJect);


     procedure Generation(stpar : string);
     function DemandeDate : TDatetime;
     function TrtGeneration (DateFacture : TdateTime): boolean;
     Procedure TrtRegroupeAffaire(TobMAff ,TobRegAff : Tob;topcaf : boolean; var stdate : string);
     Procedure TrtRechercheCaf(TobCafAff : Tob);
     Procedure ExtractAffaireRegSurCaf(TobMaff,TobCAFAff : Tob);

     procedure TestEvent (Sender: TObject);
     procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;

     private
       procedure BEffaceFicLogOnClick(Sender: TObject);
       procedure BFicLogOnClick(Sender: TObject);
		   procedure Chargement_Preparation_Appel;
	     procedure ControleChamp(Champ, Valeur: String);
       function NomFichierLog : string;
    END;

     const
	// libellés des messages de la TOF  AFPREPFACT
	TexteMsgAffaire: array[1..8] of string 	= (
          {1}        'Aucune facture n''a été générée'
          {2}        ,'Code Prestation invalide'
          {3}        ,'Date Invalide'
          {4}        ,'Dates Incompatibles'
          {5}        ,'Aucune donnée sélectionnée'
          {6}        ,'Génération en cours'
          {7}        ,'Vérifiez vos dates de génération'
          {8}        ,'Attention : vos échéances n''ont pas été regénérées'
                     );           // gm 22/08/02

  var titre					: string;
  		typ_gen				: String;
      typ_ech 			: string;
		  StatutAffaire : String;
      fac_deb				: integer;
      fac_fin				: Integer;
      nb_fac 				: integer;

Procedure AFLanceFiche_Mul_PrepFac;
Procedure AFLanceFiche_Mul_PrepFacApp;

      implementation

{ TOF_AFPREPFACT}


procedure TOF_AFPREPFACT.OnLoad;
 Var T : TTabSheet ;
 st : string;
Begin
  Inherited;
   
  // Traduction
  if (topprem) then
  Begin
       SetActiveTabSheet('P_PARAMETRE');
       topprem := false;
  End;

  T:=TTabSheet(GetControl('PCRITERE')) ;
  st := T.caption;
  SetControltext('PCRITERE',TraduitGA(st));
  titre:=ecran.caption;

  if GetParamSoc('SO_AFREVISIONPRIX') and  (THEdit(GetControl('FICHIERLOG')) <> nil) then
  begin
    if GetParamSoc('SO_AFREVPATH') = '' then
      SetControlText('FICHIERLOG', ExtractFileDrive(Application.ExeName) + '\RevisionPrix.log')
    else if copy(GetParamSoc('SO_AFREVPATH'), length(GetParamSoc('SO_AFREVPATH')) -1, 1) = '\' then
      SetControlText('FICHIERLOG', GetParamSoc('SO_AFREVPATH') + 'RevisionPrix.log')
    else
      SetControlText('FICHIERLOG', GetParamSoc('SO_AFREVPATH') + '\RevisionPrix.log');

     TToolBarButton97 (GetControl('BFICLOG')).OnClick       := BFicLogOnClick;
     TToolBarButton97 (GetControl('BEFFACEFICLOG')).OnClick := BEffaceFicLogOnClick;
  end;
End;

procedure TOF_AFPREPFACT.OnArgument(stArgument : String );
Var CC1			: THEdit;
		CC2			: THEdit;
    CC3			: THEdit;
    CC4 		: THEdit;
    CC      : THValComboBox;
    Menu 		: TPopUpMenu;
    Critere	: string;
    Valeur  : String;
    Champ   : String;
    X       : integer;
Begin
fMulDeTraitement  := true;
Inherited;
	 fTableName := 'AFFAIRE';
  topprem := true;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(stArgument)));

  while (Critere <> '') do
      begin
      if Critere <> '' then
         begin
         X := pos (':', Critere) ;
         if x = 0 then
            X := pos ('=', Critere) ;
         if x <> 0 then
            begin
            Champ := copy (Critere, 1, X - 1) ;
            Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	  ControleChamp(champ, valeur);
				    end
         end;
      Critere := (Trim(ReadTokenSt(stArgument)));
      end;



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
      

	if StatutAffaire = 'INT' then
  	 begin
     SetcontrolVisible('AFA_GENERAUTO', False);
		 SetcontrolVisible('TAFA_GENERAUTO', False);
     SetcontrolVisible('AFF_ETATAFFAIRE', False);
     SetcontrolVisible('TAFF_ETATAFFAIRE', False);
     SetControlTExt('AFF_ETATAFFAIRE','ENC;ACP');
     setcontroltext('PCRITERE', 'Critéres Contrat');
     setcontroltext('TAFF_AFFAIRE', 'Contrat');
     setcontroltext('PSTATAFF', 'Stat Contrat');
     end;

  if VH_GC.GCIfDefCEGID then  // gm 22/08/02
  begin
    SetControlText('AFA_DATEECHE',DateToStr(Nowh));
    SetControlText('AFA_DATEECHE_',DateToStr(Nowh));
  end;

  If StatutAffaire = 'APP' then
     Begin
     Chargement_Preparation_Appel;
     exit;
     end;

  // ruse pour ne pas passer par le script
  CC1 := THEDIT(GetControl('AFA_DATEECHE'));
  CC1.OnExit:=TestEvent;

  CC2 := THEDIT(GetControl('AFA_DATEECHE_'));
  CC2.OnExit:=TestEvent;

  CC3 := THEDIT(GetControl('ZDATEACT'));
  CC3.OnExit:=TestEvent;

  CC4 := THEDIT(GetControl('ZDATEACT_'));
  CC4.OnExit:=TestEvent;

  {$IFDEF BTP}
  SetControlVisible('TZDATEACT_D',false);
  SetControlVisible('TZ_AU',false);
  SetControlVisible('BTDATE',false);
  SetControlVisible('ZDATEACT',false);
  SetControlVisible('ZDATEACT_',false);
  {$ENDIF}

  If GetControl('Afa_GenerAuto') <> Nil then
    begin
    {$IFDEF BTP}
    // si modif voir Onargument de Utofetatsaffaire
    SetControlProperty('AFA_GENERAUTO','Plus','GA" AND CO_CODE="CON');
    SetControlTExt('AFA_GENERAUTO','CON');
    {$ELSE}
    // si modif voir Onargument de Utofetatsaffaire
    // ohligation de remplir texte pour ne pas prendre MAN si TOUT
    SetControlProperty('AFA_GENERAUTO','Plus','GA');
    SetControlTExt('AFA_GENERAUTO','CON;FOR;MAN;POT;POU;ACT');
    {$ENDIF}
    if ctxScot in V_PGI.PGIContexte then
    begin
     SetControlProperty('AFA_GENERAUTO','Plus','GA" AND CO_CODE<>"MAN" AND CO_CODE<>"CON');
     SetControlTExt('AFA_GENERAUTO','FOR;POT;POU;ACT');
     SetControlVisible('TAFF_TERMEECHEANCE',false);
     SetControlVisible('AFF_TERMEECHEANCE',false);
    end;
  end;

  if VH_GC.GCIfDefCEGID then  // gm 26/08/03 Cegid facture les contrats même des clients ROUGE
    // on ne peut pas mettre une multi-combo car ETATRISQUE est parfois à ''
    if GetControl('T_ETATRISQUE') <> Nil then
    begin
      SetControlTExt('T_ETATRISQUE','');
      SetControlProperty('T_ETATRISQUE','Operateur',Egal);
    end;

  SetControlChecked ('AFA_ECHEFACT',false);
  SetControlVisible('BSIMUL',false);  // 17/12/01 ne marche pas on supprime pour l'instant

  if not(ctxBTP in V_PGI.PGIContexte) then
     SetControlText ('AFF_STATUTAFFAIRE','AFF');

  if (ctxScot in V_PGI.PGIContexte) then begin
              // mcd 10/07/02
     SetControlVisible('AFF_METHECHEANCE',False);
     SetControlVisible('TAFF_METHECHEANCE',False);
     SetControlVisible('T_ZONECOM',False);
     SetControlVisible('TT_ZONECOM',False);
     end;
    // mcd 07/02/03
  If Not (GetParamSoc('SO_AFGERELIQUIDE')) then
    begin
    SetControlVisible('AFA_LIQUIDATIVE',False);
    SetControlChecked ('AFA_LIQUIDATIVE',False);
    end;
     // PL le 23/01/02 pour bogue sur affichage des préparations de factures
     // A enlever quand corrigé sur le form
  SetControlProperty('AFA_APPRECIEE','AllowGrayed',true);
  SetControlProperty('AFA_APPRECIEE','State',CbGrayed);
  SetControlVisible('AFA_APPRECIEE',false);
      // Fin PL le 23/01/02
  Menu:=TPopUpMenu(GetControl('POPREFINT'));
  if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'ARI','');

  if (not GetParamSoc('SO_AFREVISIONPRIX')) and  (THEdit(GetControl('FICHIERLOG')) <> nil) then
  begin
    SetControlVisible('TFICHIERLOG', False);
    SetControlVisible('BEFFACEFICLOG', False);
    SetControlVisible('FICHIERLOG', False);
    SetControlVisible('BFICLOG', False);
  end;

  // PL le 26/09/03 : on ne veut que les appréciations visées
  SetControlProperty ('AFA_ETATVISA', 'Value', 'VIS');

end;

procedure TOF_AFPREPFACT.OnUpdate;
Begin
Inherited;
{$IFDEF EAGLCLIENT}
TraduitAFLibGridSt(TFMul(Ecran).FListe);
{$ELSE}
TraduitAFLibGridDB(TFMul(Ecran).FListe);
{$ENDIF}
End;

Procedure TOF_AFPREPFACT.MenuPopOnCLick(Sender:Tobject);
Var Formule : THEdit;
BEGIN
Formule := THEdit(GetControl ('ZZREFINT'));
if Formule = Nil then Exit;
Formule.SelText:='['+TmenuItem(Sender).name+']';
end;


procedure TOF_AFPREPFACT.TestEvent (Sender: TObject);
var st,msg: string;
    zdeb,zfin : Tdatetime;

begin

    if (Thedit(Sender).Name='AFA_DATEECHE') then //and  Thedit(Sender).Modified then
      begin
      Thedit(Sender).Modified :=False ;
      st := Thedit(Sender).text;
      if not(IsValidDate(St)) then
        Begin
        PGIBoxAf(TexteMsgAffaire[3],'Préparation des factures');
        SetFocusControl('AFA_DATEECHE');
        exit;
        End
      else
        Begin
        zdeb :=StrToDate(St);
        if (zdeb = idate1900) then
            Begin
            PGIBoxAf(TexteMsgAffaire[7],Titre);
            SetFocusControl('AFA_DATEECHE');
            exit;
            End;
        if (zdeb > GetParamSoc('SO_AFDATEFINGENER')) then  // GM 22/08/02
        		Begin
            msg := format(TexteMsgAffaire[8]+' aprés le %s',[DateTostr(GetParamSoc('SO_AFDATEFINGENER'))]);
            PGIBoxAf(msg,Titre);
            SetFocusControl('AFA_DATEECHE');
            exit;
            End;
        End;
      end;

    if (Thedit(Sender).Name='AFA_DATEECHE_') then //and  Thedit(Sender).Modified then
      begin
      Thedit(Sender).Modified :=False ;
      st := Thedit(Sender).text;
      if not(IsValidDate(St)) then
        Begin
        PGIBoxAf(TexteMsgAffaire[3],Titre);
        SetFocusControl('AFA_DATEECHE_');
        exit;
        End
      else
        Begin
        zfin:=StrToDate(St);
        St := Thedit(GetControl('AFA_DATEECHE')).text;
        zdeb :=StrToDate(St);
        if (zfin < zdeb) Then
            Begin
            PGIBoxAf(TexteMsgAffaire[4],Titre);
            SetFocusControl('AFA_DATEECHE');
            exit;
            End
        else
          if (zfin = idate1900) then
                Begin
                PGIBoxAf(TexteMsgAffaire[7],Titre);
                SetFocusControl('AFA_DATEECHE_');
                exit;
                End
         End;
      end;

      if (Thedit(Sender).Name='ZDATEACT') then //and  Thedit(Sender).Modified then
      begin
      Thedit(Sender).Modified :=False ;
      st := Thedit(Sender).text;
      if not(IsValidDate(St)) then
        Begin
        PGIBoxAf(TexteMsgAffaire[3],Titre);
        SetFocusControl('ZDATEACT');
        exit;
        End;
      End;


      if (Thedit(Sender).Name='ZDATEACT_') then //and  Thedit(Sender).Modified then
      begin
      Thedit(Sender).Modified :=False ;
      st := Thedit(Sender).text;
      if not(IsValidDate(St)) then
        Begin
        PGIBoxAf(TexteMsgAffaire[3],Titre);
        SetFocusControl('ZDATEACT_');
        exit;
        End
      else
        Begin
        zfin:=StrToDate(St);
        St := Thedit(GetControl('ZDATEACT')).text;
        zdeb :=StrToDate(St);
        if (zfin < zdeb) Then
            Begin
            PGIBoxAf(TexteMsgAffaire[4],Titre);
            SetFocusControl('ZDATEACT');
            exit;
            End
        else
          if (zfin = idate1900) then
                Begin
                PGIBoxAf(TexteMsgAffaire[7],Titre);
                SetFocusControl('ZDATEACT_');
                exit;
                End
         End;
      end;

end;

// Fonction de recupération des affaires sélectionnées par le mul.

procedure TOF_AFPREPFACT.Generation(stpar : string);
var
     Ret : boolean;
     St,Sttexte,nature : String;
     zpar,zlib,deb,fin  : String;
     ThedateFac : Tdatetime;

begin

//if not VH_GC.GCIfDefCEGID then
// if Blocage (['nrGener'],True,'nrGener') then  exit;

  nb_fac := 0;
//  F:=TFMul(Ecran);

  typ_gen := copy(stpar,1,1);
  typ_ech := copy(stpar,3,3);

  If (typ_gen = 'G') then
      if StatutAffaire = 'APP' then
 	       St := 'Confirmez vous la génération de ces Appels'
      Else if StatutAffaire = 'INT' then
 	       St := 'Confirmez vous la génération de ces Echéances'
      Else
	       St := 'Confirmez vous la génération de ces Affaires'
  Else
      if StatutAffaire = 'APP' then
 	       St := 'Confirmez vous la SIMULATION de préparation de ces  Appels'
      Else if StatutAffaire = 'INT' then
 	       St := 'Confirmez vous la SIMULATION de préparation de ces  Contrats'
      Else
         St := 'Confirmez vous la SIMULATION de préparation de ces  Affaires';

  If (PGIAskAf(st,titre)<> mrYes) then exit;

  if GetParamSoc('SO_FACTPROV') = false then
	   Begin
     TheDateFac := DemandeDate;
     if TheDateFac= idate1900 then exit;
     end
  else
     TheDateFac := iDate1900;

  deb := FormatDateTime('dd/mm/yyyy ttttt',NowH);
  Ret := TrtGeneration (TheDateFac);
  fin := FormatDateTime('dd/mm/yyyy ttttt',NowH);

//if not VH_GC.GCIfDefCEGID then
//  Bloqueur ('nrGener',False);

  if (not Ret) then
	   Begin
     StTexte := 'Le traitement ne s''est pas terminé correctement';
     PGIInfoAf(StTexte,Titre);
     End;

  if (typ_gen = 'G') then
     zlib := 'Générée(s)'
  else
     zlib := 'Simulée(s)';

  if (fac_deb <> 0) then
     Begin
     if VH_GC.GCIfDefCEGID then
        StTexte := Format('%d Facture(s) %s : n° %d à %d  de %s à %s',[nb_fac,zlib,fac_deb,fac_fin,deb,fin])
     else
        StTexte := Format('%d Facture(s) %s : n° %d à %d ',[nb_fac,zlib,fac_deb,fac_fin]);
     PGIInfoAf(StTexte,Titre);
     if (typ_gen = 'G') then
	      if GetParamSoc('SO_FACTPROV') = False then
           nature := 'FAC'
        else
	         nature := 'FPR'
     else
        nature := 'FSI';
     if (ret) then
        Begin
        zpar := 'TABLE:GP;NATURE:'+nature+';NOCHANGE_NATUREPIECE'+';NUMDEB:'+IntToStr(fac_deb)+';NUMFIN:'+IntToStr(fac_fin)+';';
     	  if GetInfoParPiece(nature, 'GPP_VISA')='X' then
           Exit
        else
      	   if GetParamSoc('SO_FACTPROV') = True then
              AFLanceFiche_Mul_Piece_Provisoire(zpar);
        End
     End
  Else
     PgiInfoAf(TexteMsgAffaire[1],titre);

END;

//Fonction de génération automatique des factures si on passe pas
// par les factures provisoires
function TOF_AFPREPFACT.DemandeDate : Tdatetime;
var  St,zdate 		: string;
Begin

result := idate1900;

Repeat
  zdate :=AFLanceFiche_Date_ValidPiecePro('ZZDATE:'+zdate);
until (zdate = '0') or  ( ctrldate(zdate) = 0  );

if (zdate = '0') then
 begin
     PGIInfo('Traitement abandonné',titre);
     exit;
 end;


St:= 'Confirmez vous la validation de ces Factures au ' + zdate;

If (PGIAsk(st,titre)<> mrYes) then
 begin
	exit;
 end
 else
   result := StrToDate(zdate);
End;

// Fonction de recupération des affaires sélectionnées par le mul.
function TOF_AFPREPFACT.TrtGeneration (DateFacture : TdateTime): boolean;
var  	wi,nb_aff,nbmove : Integer;
			sttexte,stdate,Tmp: string;
      TobMAff: TOB; // Toutes les affaires sélectionnées
      TobCAFAff : TOB; // Toutes les affaires sélectionnées  gm 19/09/02
     	TobRegAff ,TobProfil,TobArticles,TobFormuleVar : TOB; // L'affaire ou les affaires regroupées pour 1 facture
      GenFac : R_GENFAC;
      tb_datech : array[0..99] of string;
     	AvecTrans ,TopCaf: boolean;
 //    	QQ :Tquery;
      CodeClient: String;
      LaDatefac : Tdatetime;

begin
  //result := true;
  nb_fac := 0;
  fac_deb := 0;
  fac_fin := 0;
  fillchar(tb_datech,sizeof(tb_datech),#0);

  InitToutGenfac(Genfac);

  Genfac.zgen :=  typ_gen;
  Genfac.ztypech :=  typ_ech;
  //chargement de tous les profils, afin de ne pas avoir à les relire a chaque affaire
  // on passera la tobprofil en paramètre

  TobProfil := Tob.Create('les profils',NIL,-1);
  result := ChargeTobProfil(TobProfil);

  if GetParamSoc('SO_AFVARIABLES') then
  begin
  TobFormuleVar := Tob.Create('les formules de variable',NIL,-1);
  LoadDescriptifAFFormuleVar(TOBFormuleVar);
  end
  else TobFormuleVar := Nil;

  if not(result) then
  begin
    StTexte := 'Il y a un code article inconnu dans vos profils';
    PGIInfoAf(StTexte,Titre);
    exit;
  end;

  // Controle de la date de génération des échéances  GM 22/08/02
  if (GetParamSoc('SO_AFDATEFINGENER') <> idate2099) then
  Begin
	if  (nowh >= PlusDate (GetParamSoc('SO_AFDATEFINGENER'),-3,'M'))  then
  begin
    StTexte := 'Ne pas oublier de regénérer vos échéances  pour la prochaine facturation #13#10 ';
    StTexte := StTexte + Format('Elles sont générées jusqu''au %s',[DateTostr(GetParamSoc('SO_AFDATEFINGENER'))]);
    PGIInfoAf(StTexte,Titre);
	end;
  end;

  TOBArticles:=TOB.Create('Les Articles',Nil,-1) ;

  // Gestion d'une Tob pour stocker toutes les affaires à traiter
  // Création de cette tob à partir de la multi-sélection  du multi-critère
  // Ce query est basé sur la vue AFACTIERSAFFAIRE

  TobMAff := Tob.Create('les affaires',NIL,-1);

  // PA  29/08/2001 - Fonction de traitement des enreg du mul externaliséele
  TraiteEnregMulListe (TFMul(Ecran), 'AFF_AFFAIRE','AFACTTIERSAFFAIRE', TobMAff, True);

  if not VH_GC.GCIfDefCEGID then  // gm A vérifier pour Cegid
  // mcd ajout 21/10/02 on regarde si le tiers est en risque client ==> on ne traite pas
  if (GetInfoParPiece(GenFac.Znat,'GPP_ENCOURS')='X')  and (GetInfoParPiece(GenFac.Znat,'GPP_ESTAVOIR')<>'X')then  //mcd 20/11/02 idem saisie piece
   begin
   for wi:=0 to TobMAff.Detail.count-1 do
      begin
      CodeClient:=TobMaff.detail[wi].Getvalue('T_TIERS');
      if GetEtatRisqueClient (CodeClient)='R' then TobMaff.detail[wi].free;
      end;
   end;

  //gm 19/09/02
  TobMAff.detail[0].AddChampSup('LECAF',true);

  // extraction des affaires regroupées sur CAF  dans la tob TobCafAff
  TobCAFAff := Tob.Create('les affaires sur CAF',NIL,-1);

  ExtractAffaireRegSurCaf(TobMaff,TobCAFAff);

  if StatutAffaire = 'APP' then
	 Begin
   GenFac.zdatedeb := GetControlText('AFA_DATEECHE');
   GenFac.zdatefin := GetControlText('AFA_DATEECHE_');
   GenFac.zdateact_d := GetControlText('ZDATEACT');
   GenFac.zdateact_f := GetControlText('ZDATEACT_');
   GenFac.zcom1 := GetControlText('ZZCOM1');
   GenFac.zcom2 := GetControlText('ZZCOM2');
   GenFac.zRefInt := GetControlText('ZZREFINT');
   GenFac.zcumaff := false;
   nbmove :=  TobMaff.Detail.count +   TobCAFAff.Detail.count;
	 end
  else if StatutAffaire = 'INT' then
	 Begin
   GenFac.zdatedeb := GetControlText('AFA_DATEECHE');
   GenFac.zdatefin := GetControlText('AFA_DATEECHE_');
   GenFac.zdateact_d := GetControlText('ZDATEACT');
   GenFac.zdateact_f := GetControlText('ZDATEACT_');
   GenFac.zcom1 := GetControlText('ZZCOM1');
   GenFac.zcom2 := GetControlText('ZZCOM2');
   GenFac.zRefInt := GetControlText('ZZREFINT');
   GenFac.zcumaff := false;
   nbmove :=  TobMaff.Detail.count +   TobCAFAff.Detail.count;
	 end
  else
	 begin
	 GenFac.zcumaff := false;
	 nbmove :=  TobMaff.Detail.count;
   end;

  InitMoveProgressForm (Ecran,'Préparation des factures','', nbmove, false, true) ;

  // traitement des affaires (SANS regroupement sur CAF)
  nb_aff := 0;
  while (TobMAff.Detail.count > 0)  do
    Begin
    MoveCurProgressForm('');
    // Gestion d'une Tob pour stocker l'(les) affaires constituant 1 fact
    // transfert de la tob fille TobFaff depuis TobMAff vers TobRegAff
    // La 1ère affaire sera l'affaire principale
    TobRegAff := Tob.Create('Affaire Select',NIL,-1);

    // Regroupement de affaires sur code regroupement ou Affaire/sous-affaire
    TopCaf := false;
    TrtRegroupeaffaire(TobMAff,TobRegAff,TopCaf,stdate);  // A ce niveau on a dans tobregaff les différentes affaires regroupées
    nb_aff := TobRegAff.Detail.count;

    // appel de la génération
    AvecTrans := true;
    if THEdit(GetControl('FICHIERLOG')) <> nil  then
       tmp :=  GetControlText('FICHIERLOG')
    else
       tmp:='';

    AFPrepFact_init(TobRegAff,TobProfil,TobFormuleVar,TobArticles,GenFac,nb_aff,stdate,avectrans,nb_fac,fac_deb,fac_fin,tmp,DateFacture);

    TobRegAff.free;
  End; // fin while

  // traitement des affaires (AVEC regroupement sur CAF)
  TrtRechercheCaf(TobCafAff);

  while (TobCafAff.Detail.count > 0)  do
    Begin
    MoveCurProgressForm('');
    // Gestion d'une Tob pour stocker l'(les) affaires constituant 1 fact
    // transfert de la tob fille TobFaff depuis TobMAff vers TobRegAff
    // La 1ère affaire sera l'affaire principale
    TobRegAff := Tob.Create('Affaire Select',NIL,-1);

    // Regroupement des affaires sur le CAF
    TopCaf := true;
    TrtRegroupeaffaire(TobCafAff,TobRegAff,topCaf,Stdate);
  //   nb_aff := TobRegAff.Detail.count;

    // appel de la génération
    AvecTrans := true;
    if THEdit(GetControl('FICHIERLOG')) <> nil  then
       tmp :=  GetControlText('FICHIERLOG')
    else
       tmp:='';
    AFPrepFact_init(TobRegAff,TobProfil,TobFormuleVar,TobArticles,GenFac,nb_aff,stdate,avectrans,nb_fac,fac_deb,fac_fin,Tmp,DateFacture);
    TobRegAff.free;
  End;  // fin while


  TobMaff.free;
  TobCafAff.free;
  TOBArticles.Free ;
  TobProfil.free;

  if  GetParamSoc('SO_AFVARIABLES') then  TobFormuleVar.free;

  FiniMoveProgressForm;

END;
{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 01/10/2002
Modifié le ... :   /  /
Description .. : Extraction des affaires regroupées sur CAF
Mots clefs ... : GIGA;FACTURATION
*****************************************************************}
Procedure TOF_AFPREPFACT.ExtractAffaireRegSurCaf(TobMaff,TobCAFAff : TOB);
Var wj : Integer;
		Tobdet : Tob;
BEGIN
	wj := 0;
  while (wj < TobMAff.Detail.Count) do
  Begin
    TobDet := TobMAff.Detail[wj];
    if (Tobdet.getValue('AFF_REGSURCAF') = 'X') then
      Tobdet.ChangeParent(TobCafAff,-1)
    else
      inc(wj);
  end;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Merieux
Créé le ...... : 19/09/2002
Modifié le ... :   /  /    
Description .. : Regourpement des affaire sur  divers critères :
Suite ........ : * Mode de regroupement 
Suite ........ : * Afafire/sous affaire
Mots clefs ... : GIGA;FACTURATION;REGROUPEMENT
*****************************************************************}
Procedure TOF_AFPREPFACT.TrtRegroupeAffaire(TobMAff ,TobRegAff : Tob;TopCaf:Boolean;var stdate : string);
Var TobFaff : Tob;
		StAff,StTiers,StCaf,StReg,StAffRef,StIsAffRef,sttiersfact : String;
    zord,zcas,wj: Integer;
BEGIN
    TobFAff := TobMAff.Detail[0];
    StAff := TobFaff.getValue('AFF_AFFAIRE');
    StAffRef := TobFaff.getValue('AFF_AFFAIREREF');
    StIsAffRef := TobFaff.getValue('AFF_ISAFFAIREREF');
    StTiers := TobFaff.getValue('T_TIERS');
    sttiersfact := TobFaff.getValue('AFF_FACTURE');
    StReg :=  TobFaff.getValue('AFF_REGROUPEFACT');
    StDate := TobFaff.getValue('AFA_DATEECHE');
    StCAF := TobFaff.getValue('LECAF');

    // recup 1ere affaire  de la tob
    TobFAff.ChangeParent(TobRegAff,-1);

    // détermination des différents cas :
    //  0 : Affaire indépendante
    //  1 : Affaire étant rattaché à une affaire de référence
    //  2 : Affaire de référence
    //  3 : Affaire ayant un code regroupement
    //  4 : regroupement sur CAF
    zcas := 0;
    if (StAffRef <> Staff)  then
       zcas := 1
    else
        Begin
        if (StIsAffRef = 'X') then
           zcas := 2
        else
            Begin
            if ((StReg <> '') and (StReg <> 'AUC')) Then  zcas := 3;
            End;
        End;

    if (TopCaf) then zcas := 4;

    // Recherche des affaires regroupées  (même client, même échéance,même code reg.)
    if (zcas <> 0) Then
    Begin
      wj := 0;
      while wj < TobMAff.Detail.count do
      Begin
        TobFAff := TobMAff.Detail[wj];
        case zcas of
        //  1 : Affaire étant rattaché à une affaire de référence
        1 : Begin
            if ((sttiersfact = TobFaff.getValue('AFF_FACTURE')) and
                (StDate = datetostr(TobFaff.getValue('AFA_DATEECHE'))) and
                ((StAffRef = TobFaff.getValue('AFF_AFFAIRE')) or(StAffRef = TobFaff.getValue('AFF_AFFAIREREF')) ))  then
            Begin
              if (TobFaff.getValue('AFF_ISAFFAIREREF') = 'X') Then zord := 0 else zord:= -1;
              TobFAff.ChangeParent(TobRegAff,zord);
            End
            else
                inc(wj);
            End;

        // 2 : Affaire de référence
        2 : Begin
            if ((sttiersfact = TobFaff.getValue('AFF_FACTURE')) and
                (StDate = datetostr(TobFaff.getValue('AFA_DATEECHE'))) and
                (StAff = TobFaff.getValue('AFF_AFFAIREREF')) )  then
            Begin
              zord:= -1;
              TobFAff.ChangeParent(TobRegAff,zord);
            End
            else
                inc(wj);
            End;
        //  3 : Affaire ayant un code regroupement
        3 : Begin
            if ((sttiersfact = TobFaff.getValue('AFF_FACTURE')) and
                (StDate = datetostr(TobFaff.getValue('AFA_DATEECHE'))) and
                (StReg = TobFaff.getValue('AFF_REGROUPEFACT')) )  then
            Begin
              if (TobFaff.getValue('AFF_PRINCIPALE') = 'X') Then zord := 0 else zord:= -1;
              TobFAff.ChangeParent(TobRegAff,zord);
            End
            else
                inc(wj);
            End;
        //  4 : Affaire regroupée sur CAF
        4 : Begin
            if ((StCAF = TobFaff.getValue('LECAF')) and
                (StDate = datetostr(TobFaff.getValue('AFA_DATEECHE'))) and
                (StReg = TobFaff.getValue('AFF_REGROUPEFACT')) )  then
            Begin
              if (TobFaff.getValue('AFF_PRINCIPALE') = 'X') Then zord := 0 else zord:= -1;
              TobFAff.ChangeParent(TobRegAff,zord);
            End
            else
                inc(wj);
            End
        End; // fin case
      End; //fin while
    End;
END;


{***********A.G.L.***********************************************
Auteur  ...... : G Merieux
Créé le ...... : 19/09/2002
Modifié le ... :   /  /
Description .. : Traitement de la tob pour regroupement sur le CAF
Mots clefs ... : GIGA;FACTURATION
*****************************************************************}
Procedure TOF_AFPREPFACT.TrtRechercheCaf(TobCafAff : Tob);
Var Tobdet : TOB;
		wj,nbpiece : integer;
    Sttiers , StRep: string;
    CleDocTemp : R_CleDoc;
BEGIN
// récupération du Client à facturer de l'affaire

wj := 0;
while (wj < TobCafAff.Detail.Count) do
Begin
		TobDet := TobCafAff.detail[wj];
    NbPiece := SelectPieceAffaireBis(TobDet.Getvalue('AFF_AFFAIRE'), 'AFF', CleDocTemp ,Sttiers,StRep);
    if (NbPiece =1) then
    	Begin
    	TobDet.putValue('LECAF',Sttiers);
      inc(wj);
    	End
    else
    	TobDet.free;
End;

// Tri de la tob sur LECAF
if (TobCafAff.Detail.Count > 0) then
	TObCafAff.detail.sort('LECAF;AFF_REGROUPEFACT;T_TIERS;AFF_AFFAIRE');

END;


procedure TOF_AFPREPFACT.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
end;

procedure TOF_AFPREPFACT.BEffaceFicLogOnClick(Sender: TObject);
var
  stDocWord : string ;
begin
  {$IFNDEF EAGLCLIENT}
  if THEdit(GetControl('FICHIERLOG')) <> nil  then   stDocWord := THEdit(GetControl('FICHIERLOG')).Text
     else stDocWord:='';
  if (stDocWord='') then exit;
  if (PGIAskAF ('Confirmez-vous la suppression du fichier ' + stDocWord + ' ?', Ecran.Caption)=mrYes) then
    if not DeleteFile(pchar(stDocWord)) then
       PGIInfoAF('Le fichier ne peut-être supprimé ou n''existe pas.', TFMUL(Ecran).caption);
  {$ENDIF}
end;

procedure TOF_AFPREPFACT.BFicLogOnClick(Sender: TObject);
var
  stDocWord : string;
begin
{$IFNDEF EAGLCLIENT}
  if THEdit(GetControl('FICHIERLOG')) <> nil  then   stDocWord := THEdit(GetControl('FICHIERLOG')).Text
     else stDocWord:='';
  if (stDocWord='') then exit;
  if Not FileExists(stDocWord) then
  begin
    PGIInfoAF('Le fichier ' + stDocWord +' n''existe pas', TFMUL(Ecran).caption);
    exit;
  end;

  ShellExecute (0, PCHAR('open'),PChar(stDocWord), Nil,Nil,SW_RESTORE);
{$ENDIF}
end;

function TOF_AFPREPFACT.NomFichierLog : String;
var
  NomLog, NomFicLog : String;
  NomFic1,NomFic2   : String;
  StPath2,StPath1   : String;

begin

  Result    := '';
  NomLog    := '';
  NomFicLog := '';
  StPath2   :='';
  StPath1   :='';
  NomFic1   :='';
  NomFic2   :='';

  if THEdit(GetControl('FICHIERLOG')) <> nil  then   NomLog := THEdit(GetControl('FICHIERLOG')).Text
     else NomLog:='';
  if (NomLog<>'') then
  begin
    StPath1:=ExtractFilePath(NomLog);
    NomFic1:=ChangeFileExt(ExtractFileName(NomLog),'.log');
  end;

  if (StPath1<>'') then
  begin
    if (NomFic1<>'') then
      NomFicLog:=StPath1+NomFic1
    else
      NomFicLog:=StPath1+'RevisionPrix.log';
  end;

  if (NomFicLog = '') then
     //NomFicLog := 'c:\RevisionPrix.log';
     NomFicLog := IncludeTrailingBackSlash(TCBPPath.GetCegidUserTempFileName) + 'RevisionPrix.log';

  if THEdit(GetControl('FICHIERLOG')) <> nil  then   THEdit(GetControl('FICHIERLOG')).Text := NomFicLog;
  Result := NomFicLog;
end;

procedure AGLPrepFact_Generation( parms: array of variant; nb: integer );
var  F : TForm;
     LaTof : TOF;
begin
F:=TForm(Longint(Parms[0]));
if (F is TFMul) then Latof:=TFMul(F).Latof else latof := nil;
if (Latof is TOF_AFPREPFACT) then TOF_AFPREPFACT(LaTof).Generation(Parms[1]);
end;

Procedure AFLanceFiche_Mul_PrepFac;
begin
{$IFDEF BTP}
AGLLanceFiche ('BTP','BTPREPFACT_MUL','AFF_AFFAIRE0=I','','STATUT=INT');
{$ELSE}
AGLLanceFiche ('BTP','BTPREPFACT_MUL','','','');
{$ENDIF}
end;

Procedure AFLanceFiche_Mul_PrepFacApp;
begin
AGLLanceFiche ('BTP','BTPREPFACAPP','AFF_AFFAIRE0=W','','STATUT=APP');
end;


Function CtrlDate(Zdat : string) : integer;
begin
result := 0;

if not(IsValidDate(Zdat)) then
  Begin
    result := 1;
  End
else
    Begin
    if (ControleDate(Zdat) > 0) then
         Begin
    	result := 2;
        End
  else   if (StrToDAte(Zdat) <> Idate1900) and (ctxscot in V_PGI.PGIContexte)
    and (StrToDAte(Zdat) < GetParamSoc('So_AFDATEDEBUTACT')) then result:=3;  //mcd 09/08/02 bloque si < arrété période
  End;
end;

//Controle des valeurs passées dans le OnArgument
Procedure TOF_AFPREPFACT.ControleChamp(Champ : String;Valeur : String);
Begin

  //Chargement du code affaire
  if Champ ='STATUT' then StatutAffaire :=valeur;

end;

//Chargement de l'écran en fonction du code Statut de l'affaire si APP = APPELS
Procedure TOF_AFPREPFACT.Chargement_Preparation_Appel;
var Menu 		: TPopUpMenu;
Begin

  SetcontrolVisible('AFA_GENERAUTO', False);
	SetcontrolVisible('TAFA_GENERAUTO', False);
  SetcontrolVisible('AFF_ETATAFFAIRE', False);
  SetcontrolVisible('TAFF_ETATAFFAIRE', False);
  setcontroltext('PCRITERE', 'Critéres Appel');
  setcontroltext('TAFF_AFFAIRE', 'Appel');
  setcontroltext('PSTATAFF', 'Stat Appel');

  SetControlProperty('AFA_GENERAUTO','Plus','GA" AND CO_CODE="DIR');
  SetControlText('AFA_GENERAUTO','DIR');
  SetControlText('AFF_ETATAFFAIRE','ACA');
  SetControlText ('AFF_STATUTAFFAIRE','APP');

  SetControlChecked ('AFA_ECHEFACT',false);
  SetControlVisible('BSIMUL',false);

  //déjà dans la fonction principale a voir pour intégrer en une fois...

  SetControlProperty('AFA_APPRECIEE','AllowGrayed',true);
  SetControlProperty('AFA_APPRECIEE','State',CbGrayed);
  SetControlVisible('AFA_APPRECIEE',false);

  // Fin PL le 23/01/02
  Menu:=TPopUpMenu(GetControl('POPREFINT'));
  if Menu <>  Nil then ChargePopUp(Menu,MenuPopOnClick,'ARI','');

  if (not GetParamSoc('SO_AFREVISIONPRIX')) and  (THEdit(GetControl('FICHIERLOG')) <> nil) then
  begin
    SetControlVisible('TFICHIERLOG', False);
    SetControlVisible('BEFFACEFICLOG', False);
    SetControlVisible('FICHIERLOG', False);
    SetControlVisible('BFICLOG', False);
  end;

  // PL le 26/09/03 : on ne veut que les appréciations visées
  SetControlProperty ('AFA_ETATVISA', 'Value', 'VIS');

end;

Initialization
registerclasses([TOF_AFPREPFACT]);
RegisterAglProc( 'PrepFact_Generation',True,1,AGLPrepFact_Generation);
end.



