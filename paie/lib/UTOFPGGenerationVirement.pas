{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Génération des virements
Mots clefs ... : PAIE,VIREMENT
*****************************************************************
PT1    : 10/09/2001 SB V547 Fiche de bug n°288 - Edition du fichier de virement
                            Position de code
PT2    : 12/09/2001 SB V547 Remplacement Q.isEmpty par Q.eof

PT3    : 19/09/2001 SB V547 Emission du rib salarie pouvant être erronée
                            Condition sur jointure et non sur WHERE
                            Fiche de bug n°310
PT4    : 28/11/2001 SB V563 Suit modif. AGL on ne redirige plus le bcherche
PT5    : 13/12/2001 SB V563 Fiche de bug n°408 Erreur Sql Oracle Join non
                            nécessaire
PT6    : 29/01/2002 SB V571 Edition sur banque à emmettre différent
PT7    : 13/03/2002 SB V571 Fiche de bug n°10015 Payé le <> idate1900
PT8    : 14/03/2002 SB V571 Fiche de com n°011205M Edition sur banque à emmettre
                            différent
                            Ajout champ PVI_RIBSALAIREEMIS
PT9    : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                            renseigné en Mono
PT10   : 13/05/2002 PH V575 Fiche Q n°10099 suppression : pour traitement
                            internet
PT11   : 04/06/2002 PH V582 Gestion historique des évènements
PT12   : 17/09/2002 SB V585 FQ n°10136 Edition post génération erronné du fait
                            modif PT1
PT13-1 : 03/03/2003 SB V585 SB FQ n°10537 Montant de la remise érronée..
PT13-2 : 03/03/2003 SB V585 FQ n°10537 Ajout champs en cours de génération pour
                            traitement édition..
PT14   : 07/04/2003 SB V_42 FQ n°10583 Erreur sql si ajout champ table salaries
PT15   : 05/08/2004 JL V_50 FQ n°11476 Message : renseigné au lieu de renseigner
PT16   : 08/02/2005 SB V_60 FQ 11826 Ajout champ PVI_ECHEANCE
PT17   : 14/02/2005 SB V_60 FQ 11964 Anomalie banque à émettre
PT18   : 20/06/2005 PH V_60 FQ 12305 Génération du fichier de virements avec ou
                            sans séparateur
PT19   : 07/07/2005 PH V_60 FQ 12417 Prises en compte uniquement des paies
                            règlées par virement
PT20   : 09/08/2005 PH V_60 FQ 12422 Modifs du PT4 Chargement devient un Integer
PT21   : 14/09/2005 PH V_60 FQ 11813 Controle saisie banque à émettre
PT22   : 27/09/2005 SB V_60 FQ 12460 Ajout sélection salariés sortis
PT23   : 14/04/2006 SB V_65 FQ 12152 Ajout rupture etablissement
PT24   : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT25   : 22/03/2007 FC V_72 FQ 13477 Modification du ORDER BY : établissement
                            puis les dates début et fin
PT26   : 12/12/2007 FL V_81 TESSI - Génération multidossier : ajout de la TOF
                            PGVIREMENTDOS + correction bugs mémoire
PT27   : 03/01/2008 FL V_81 FQ 15049 - Ajout de la combo types de bulletins
                            ("Voir")
}
unit UTOFPGGenerationVirement;
interface
uses StdCtrls, Controls, Classes,sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,  Mul,
  EdtREtat,Fe_Main,
  {$ELSE}
  UtileAGL, eMul, maineagl,
  {$ENDIF}
  ParamSoc,   Dialogs, Windows,  HQry,
  HTB97,
  PgOutilsTreso, Commun, Constantes,PGAnomaliesTraitement,MenuDispPS5;

type
  TOF_PGVIREMENTS = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override;
    procedure OnClose; override;
  private
    RecupTobVir: TOB;
    PeriodeChange, OkGeneration: Boolean;
    Chargement : Integer ; //PT4 Chargement PT20
    Separat : Boolean ; // PT18 Avecou sans séparateur
    //MontRemise : Double;PT13-1
    DatePerEncours: TDateTime;
    GblSalSortis : String; { PT22 }
    procedure ChangePeriode(Sender: TObject);
    procedure AffectRib(Sender: TObject);
    procedure ReglementPaie(Sender: TObject);
    procedure AffectDatePayeLe(PerEnCours: TdateTime);
    procedure LanceEtatVir(ReglePaie: Boolean);
    function GenereFichier(Rib, RibMaj: string): Boolean;
    procedure DetruitFichier(NomFic: string);
    procedure CalcRemise(Sender: TObject);
    function RecupCritereVir(Avec: boolean): string;
    //   procedure LanceCritere(Sender: TObject);
    procedure ExitEdit(Sender: TObject);
    procedure OnChangeTypeBul (Sender: TObject); //PT27
  end;

//PT26 - Début
type
  TOF_PGVIREMENTDOS = class(TOF)
    procedure OnArgument(stArgument: string);    override;
    procedure OnLoad;                            override;
    procedure OnClose;                           override;
  private
    Grille          : THGrid;
    ListeAnomalies  : TAnomalies;
    ProcessInfo     : TProcessInformation;       // Pour l'application lancée une seconde fois
    Procedure ChangeRegroupement    (Sender : TObject);
    procedure LanceGeneration       (Sender : TObject);
    Procedure AfficheListeDossiers;
    procedure SelectAll(Sender: TObject);
  end;
 //PT26 - Fin


implementation
uses EntPaie, PGoutils, PgOutils2, PGEditOutils2, P5Def,
     PGGenerationVirement, PGOperation, HPanel, Vierge, Forms, shellapi, ed_tools ;

type
  Vir03 = record // ENREGISTREMENT DE TYPE ENTETE VIREMENT
    CodeEnr: string[2]; // Code enregistrement
    CodeOP: string[2]; // Code Operation
    ZR1: string[8]; // Zone reservee
    NumEmet: string[6]; // numéro emetteur
    ZR2: string[7]; // Zone reservee
    PayeLe: string[5]; // Date de paiement jjmma
    RaisonS: string[24]; // BQ_LIBELLE  raison sociale emetteur
    RefRem: string[7]; // reference de la remise
    ZR3: string[19]; // Zone reservee
    Monnaie: string[1]; // monnaie de la remise
    ZR4: string[5]; // Zone reservee
    CodeGui: string[5]; // BQ_GUICHET code quichet DO
    NumCpte: string[11]; // BQ_NUMEROCOMPTE numero de cpte DO
    ZR5: string[2]; // Zone reservee
    Identif: string[14]; // Identifiant DO
    ZR6: string[31]; // Zone reservée
    Banque: string[5]; // BQ_ETABBQ code banque du DO
    ZR7: string[6]; // Zone reservee
  end;
type
  Vir06 = record // ENREGISTREMENT DE TYPE SALARIE
    CodeEnr: string[2]; // Code enregistrement
    CodeOP: string[2]; // Code Operation
    ZR1: string[8]; // Zone reservee
    NumEmet: string[6]; // numéro emetteur
    RefInter: string[12]; // Reference interne
    NomDest: string[24]; // Nom du destinataire Nom Salarie
    Domicil: string[20]; // Domiciliation salarie R_DOMICILIATION
    NatEco: string[1]; // Nature Eco pour N.R
    Pays: string[3]; // Pays Eco pour N.R
    BalPay: string[8]; // Balance des paiements
    CodeGui: string[5]; // R_GUICHET code quichet salarie
    NumCpte: string[11]; // R_NUMEROCOMPTE numero de cpte salarie
    Montant: string[16]; // Montant net à payer PVI_MONTANT
    Libelle: string[29]; // Libellé := SALARIE : Numero du salarie
    CodeRej: string[2]; // Code rejet
    Banque: string[5]; // R_ETABBQ code banque
    ZR2: string[6]; // Zone reservee
  end;
type
  Vir08 = record // ENREGISTREMENT DE TYPE TOTAL DE LA REMISE
    CodeEnr: string[2]; // Code enregistrement
    CodeOP: string[2]; // Code Operation
    ZR1: string[8]; // Zone reservee
    NumEmet: string[6]; // numéro emetteur
    ZR2: string[84]; // Zone reservée
    Montant: string[16]; // Montant total de la remise
    ZR3: string[42]; // Zone reservee
  end;


procedure TOF_PGVIREMENTS.OnArgument(stArgument: string);
var
  Ok: Boolean;
  CbxCodOp, CbxRibSal, cbMois, cbAnnee: THValComboBox;
  Min, Max, Mois, annee, Exer: string;
  Btn: TToolBarButton97;
  {$IFNDEF EAGLCLIENT}
  L: THDBGrid;
  {$ELSE}
  L: THGrid;
  {$ENDIF}
  Defaut: THEdit;
  DebExer, FinExer: TDateTime;
  Check: TCheckBox;
begin
  inherited;
  Chargement := 0; //PT4 PT20
  PGSuppression := 'VIR';
  ExecuteSQL('DELETE FROM VIREMENTS');
  //Evenements OnChange
  {$IFNDEF EAGLCLIENT}
  L := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  L := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if L <> nil then L.OnFlipSelection := CalcRemise;
  //Btn:=TToolBarButton97(GetControl('BCherche'));  //PT4
  //if (Btn<>nil) then Btn.OnClick:=LanceCritere;  //PT4
  Btn := TToolBarButton97(GetControl('BOUVRIR'));
  if (Btn <> nil) then Btn.onclick := ReglementPaie;
  CbxRibSal := THValComboBox(GetControl('PVI_RIBSALAIRE'));
  if CbxRibSal <> nil then CbxRibSal.OnChange := AffectRib;
  
  If Assigned(GetControl('CBXBULLCOMPL')) Then THValComboBox(GetControl('CBXBULLCOMPL')).OnChange := OnChangeTypeBul; //PT27

  // Affectation des valeurs par défaut;
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier; //PT9 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  Check := TCheckBox(GetControl('CKEURO'));
  if Check <> nil then Check.Checked := VH_Paie.PGTenueEuro;
  RecupMinMaxTablette('PG', 'SALARIES', 'PSA_SALARIE', Min, Max);
  Defaut := ThEdit(getcontrol('PVI_SALARIE'));
  if Defaut <> nil then
  begin
    Defaut.text := Min;
    Defaut.OnExit := ExitEdit;
  end;
  Defaut := ThEdit(getcontrol('PVI_SALARIE_'));
  if Defaut <> nil then
  begin
    Defaut.text := Max;
    Defaut.OnExit := ExitEdit;
  end;
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  Defaut := ThEdit(getcontrol('PVI_ETABLISSEMENT'));
  if Defaut <> nil then Defaut.text := Min;
  Defaut := ThEdit(getcontrol('PVI_ETABLISSEMENT_'));
  if Defaut <> nil then Defaut.text := Max;
  CbxCodOp := THValComboBox(GetControl('VCBXCODEOP'));
  if CbxCodOp <> nil then CbxCodOp.value := '02';
  SetControlProperty('VCBXSUPPORT', 'value', 'DIS');
  SetControlChecked('PVI_TOPREGLE', False);

  ok := RendExerSocialEnCours(Mois, annee, Exer, DebExer, FinExer);
  if Ok = True then
  begin
    DatePerEncours := EncodeDate(StrToInt(Annee), StrToInt(Mois), 1);
    AffectDatePayeLe(DatePerEncours);
    cbMois := THValComboBox(GetControl('CBMOIS'));
    if cbMois <> nil then
    begin
      if Length(Mois) = 1 then Mois := '0' + Mois;
      cbMois.value := Mois;
      CbMois.onchange := ChangePeriode;
    end;
    cbAnnee := THValComboBox(GetControl('CBANNEE'));
    if cbAnnee <> nil then
    begin
      CbAnnee.value := Exer;
      CbAnnee.onchange := ChangePeriode;
    end;
  end;

  SetControltext ('CBXSALSORTIS', 'SAL');  { PT22 }
  GblSalSortis := 'SAL';                   { PT22 }

  //PT24
  SetPlusBanqueCP (GetControl ('PVI_RIBSALAIRE'));
  SetPlusBanqueCP (GetControl ('VCBXRIBSAL'));
  //FIN PT24

  //PT26 - Début
  If VH_PAIE.ModeFiche Then
  Begin
    VH_PAIE.FicheOuverte    := True;

    // On ajuste la fenêtre au mieux
    TFMul(Ecran).AutoSize   := False;
    TFMul(Ecran).Scaled     := False;
    If stArgument <> '' Then
    Begin
        TFMul(Ecran).Top        := StrToInt(ReadTokenSt(stArgument));
        TFMul(Ecran).Left       := StrToInt(ReadTokenSt(stArgument));
        TFMul(Ecran).Width      := StrToInt(ReadTokenSt(stArgument));
        TFMul(Ecran).Height     := StrToInt(ReadTokenSt(stArgument));
    End;

    // Changement du titre
    TFMul(Ecran).Caption    := 'Génération des virements du dossier '+ V_PGI.CurrentAlias;
    UpdateCaption(TFMul(Ecran));
  end;
  //PT26 - Fin
end;


procedure TOF_PGVIREMENTS.OnLoad;
var
  StPlus, StWhere, DateSes, RibSalaire, Where, StWhereDate, StWhereBul: string; //PT27
  cbMois, cbAnnee, CbxRibSal: THValComboBox;
  Mois: WORD;
  Annee: string;
  DateDSes,DateFSes : TDateTime;
  Remise: THNumEdit;
  QRech: TQuery;
  Pages: TPageControl;
begin
  inherited;

    //PT27
    DateDSes := 0;
    DateFSes := 0;

	//DEB PT4 PT20
	if Chargement < 2  then
	begin
		Chargement := Chargement + 1;
		//PT26 - Début
		// Présélection si 1 seul élément dans la combo
		If THValComboBox(GetControl('PVI_RIBSALAIRE')).Items.Count = 1 Then SetControlText('PVI_RIBSALAIRE',THValComboBox(GetControl('PVI_RIBSALAIRE')).Values[0]);
		If THValComboBox(GetControl('VCBXRIBSAL')).Items.Count = 1 Then SetControlText('VCBXRIBSAL',THValComboBox(GetControl('VCBXRIBSAL')).Values[0]);
		//PT26 - Fin
		exit;
	end;  // FIN PT20
	
	if GetControlText('PVI_RIBSALAIRE') = '' then
	begin
		PgiBox('Vous devez renseigner la banque du donneur d''ordre!', 'Virement des salaires');
		exit;
	end;
	//FIN PT4
	
	if GetControltext ('CBXSALSORTIS') <> GblSalSortis then   { DEB PT22 }
	Begin
		PeriodeChange := True;
		GblSalSortis := GetControltext ('CBXSALSORTIS');
	End;                                                    { FIN PT22 }
	
	Remise := THNumEdit(GetControl('REMISE'));
	Pages := TPageControl(GetControl('Pages'));
	//ReChargement de la table si modif periode ou compte bancaire DO
	if PeriodeChange = True then
	begin
		mois := 1;
		cbMois := THValComboBox(GetControl('CBMOIS'));
		cbAnnee := THValComboBox(GetControl('CBANNEE'));
		CbxRibSal := THValComboBox(GetControl('PVI_RIBSALAIRE'));
		if cbMois.Value <> '' then Mois := Trunc(StrToInt(cbMois.Value));
		if (cbMois.Value <> '') and (cbAnnee.Value <> '') then // cas où mois renseigné alors année obligatoire
		begin
			ControlMoisAnneeExer(cbMois.value, RechDom('PGANNEESOCIALE', cbAnnee.Value, FALSE), Annee);
			DateDSes := EncodeDate(StrToInt(annee), mois, 1);
			DatePerEncours := DateDSes;
			AffectDatePayeLe(DateDSes);
			DateSes := ' AND PPU_DATEFIN>="' + UsDateTime(DateDSes) + '"';
			DateFSes := FINDEMOIS(DateDSes);
			DateSes := DateSes + ' AND PPU_DATEFIN<="' + UsDateTime(DateFSes) + '"';
		end;
	
	
		if CbxRibSal <> nil then
		if CbxRibSal.Value <> '' then
		begin
			StPlus:= PGBanqueCP (True);             //PT24
			RibSalaire := ' AND BQ_GENERAL="' + CbxRibSal.Value + '"'+StPlus;
		end;
	
		if (GetControltext ('CBXSALSORTIS') <> 'SAL') then    { DEB PT22 }
		begin
			if GetControltext ('CBXSALSORTIS') = 'SAN' then // Sans les salariés sortis
			StWhereDate := ' AND (PSA_DATESORTIE >"'+UsDateTime(DateDSes)+'" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE="'+UsDateTime(Idate1900)+'")'
			else // Uniquement les salariés sotis dans la période
			StWhereDate := ' AND (PSA_DATESORTIE >="'+UsDateTime(DateDSes)+'" '+
			'AND PSA_DATESORTIE <="'+UsDateTime(DateFSes)+'")';
		end;   
		{ FIN PT22 }
		
		//PT27 - Début
		if GetControlText('CBXBULLCOMPL') = '001' then StWhereBul := '';
      	if GetControlText('CBXBULLCOMPL') = '002' then StWhereBul := ' AND PPU_BULCOMPL ="-"';
    	if GetControlText('CBXBULLCOMPL') = '003' then StWhereBul := ' AND PPU_BULCOMPL ="X"';
		//PT27 - Fin
	
		if (DateSes <> '') and (RibSalaire <> '') then
		begin
			StWhere := ' ' + DateSes + RibSalaire + StWhereDate + StWhereBul; //critère de sélection  { PT22 } //PT27
			RecupTobVir := RendTobGenVirement(StWhere);
			//PT13-1 MontRemise:=RecupTobVir.Somme('PVI_MONTANT',[''],[''],TRUE);
			//PT13-1 Remise.value:=MontRemise;
		end;
	end;
	
	{ DEB PT13-1 Calcul du montant de la remise en fonction de l'affichage }
	Where := RecupWhereCritere(Pages) + StWhereDate ;           { PT22 }
	{ DEB PT14 Jointure dynamique si utilisation critère salaries }
	if (Pos('PSA_', Where) > 0) then
	Where := 'LEFT JOIN SALARIES ON PVI_SALARIE=PSA_SALARIE ' + Where;
	{ FIN PT14 }
	QRech := Opensql('SELECT SUM(PVI_MONTANT) FROM VIREMENTS ' + Where + '', True);
	if (not QRech.eof) and (Remise <> nil) then
	Remise.value := Qrech.fields[0].asfloat;
	Ferme(QRech);
	{ FIN PT13-1 }

	PeriodeChange := False;
end;


procedure TOF_PGVIREMENTS.ReglementPaie(Sender: TObject);
var
  cbMois, cbAnnee, cbRib, cbRibMaj: THValComboBox;
  Mois, Annee, Rib, RibMaj, LibRibMaj: string;
  DateDeb : TDateTime;
  Btn, bt: TToolBarButton97;
  {$IFNDEF EAGLCLIENT}
  L: THDBGrid;
  {$ELSE}
  L: THGrid;
  {$ENDIF}
  Chk: TCheckBox;
  Where, S1, S2, StWhereBul: string; //PT27
  Trace: TStringList;
begin
  {$IFNDEF EAGLCLIENT}
  L := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  L := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  Btn := TToolBarButton97(GetControl('BCHERCHE'));
  Bt := TToolBarButton97(GetControl('BTSUPPRIMER'));
  cbMois := THValComboBox(GetControl('CBMOIS'));
  if cbMois <> nil then Mois := cbMois.value;
  cbAnnee := THValComboBox(GetControl('CBANNEE'));
  ControlMoisAnneeExer(Mois, RechDom('PGANNEESOCIALE', cbAnnee.Value, FALSE), Annee);
  DateDeb := EncodeDate(StrToInt(Annee), StrToInt(Mois), 1);

  cbRib := THValComboBox(GetControl('PVI_RIBSALAIRE')); //DO initial
  cbRibMaj := THValComboBox(GetControl('VCBXRIBSAL')); //DO Mise à jour
  if cbRib <> nil then Rib := cbRib.value;
  if cbRibMAj <> nil then RibMaj := cbRibMaj.value;
{PT24
  LibRibMaj := RechDom ('TTBANQUECP', RibMaj, False);
}
  LibRibMaj:= RechDom ('TTBANQUECP', RibMaj, False,
                       'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE);
//FIN PT24

  if L <> nil then
  begin
    if (L.NbSelected = 0) then if Btn <> nil then Btn.Click;
    if (L.NbSelected > 0) then
      if MessageDlg('Vous avez selectionné des lignes de paie. Voulez-vous les retirer de la génération ?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
        L.ClearSelected
      else
        if Bt <> nil then Bt.Click; //Suppression des lignes selectionnés
  end;

  if (not TFmul(Ecran).Q.EOF) and (L <> nil) and (Rib <> '') then
  begin //Test si fichier imprimante
    if GetControlText('VCBXSUPPORT') = 'IMP' then
    begin
      LanceEtatVir(False);
      Exit;
    end;
    //Génération du fichier
    ExecuteSql('UPDATE VIREMENTS SET PVI_TOPENCOURS="-"'); //PT13-2 RAZ
    OkGeneration := GenereFichier(Rib, RibMaj);
    if OkGeneration = True then
    begin //  MISE A JOUR de la table PAIEENCOURS
      Where := RecupCritereVir(False);
      
	  //PT27 - Début
  	  if GetControlText('CBXBULLCOMPL') = '001' then StWhereBul := '';
      if GetControlText('CBXBULLCOMPL') = '002' then StWhereBul := ' AND PPU_BULCOMPL ="-"';
      if GetControlText('CBXBULLCOMPL') = '003' then StWhereBul := ' AND PPU_BULCOMPL ="X"';
	  //PT27 - Fin
      
      ExecuteSQL('UPDATE PAIEENCOURS SET PPU_TOPREGLE="X",PPU_RIBSALAIRE="' + RibMaj + '", ' +
        'PPU_BANQUEEMIS="' + LibRibMaj + '",PPU_ECHEANCE="' + USDateTime(StrtoDate(GetControlText('ECHEANCE'))) + '" ' + //PT16
        'WHERE PPU_SALARIE IN (SELECT DISTINCT PVI_SALARIE FROM VIREMENTS ' +
        'WHERE PVI_RIBSALAIRE="' + Rib + '" ' + Where + ' ) ' + //PT5 LEFT JOIN PAIEENCOURS '+    'ON PVI_SALARIE=PPU_SALARIE
        'AND PPU_DATEFIN>="' + USDateTime(DateDeb) + '" AND PPU_DATEFIN<="' + USDateTime(FindeMois(DateDeb)) + '"' + StWhereBul); //PT27
      Where := RecupCritereVir(True); //  MISE A JOUR de la table VIREMENTS
      //DEB PT12 On effectue la maj des virements et on positionne PVI_TOPREGLE à cocher pour pouvoir edité l'état à l'image de la génération
      ExecuteSQL('UPDATE VIREMENTS SET PVI_TOPREGLE="X",PVI_RIBSALAIRE="' + RibMaj + '",' +
        'PVI_RIBSALAIREEMIS="' + RibMaj + '",' + 'PVI_BANQUEEMIS="' + LibRibMaj + '" ,'+
        'PVI_ECHEANCE="' + USDateTime(StrtoDate(GetControlText('ECHEANCE'))) + '" '+Where); //PT8 PT16
      SetControlChecked('PVI_TOPREGLE', True);
      //FIN PT12
      if MessageDlg('Génération du fichier effectuée.Voulez-vous une édition?', mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
      else LanceEtatVir(True); {PT1 On edit avant de mettre à jour}
      // PT11 : 04/06/2002 V582 PH Gestion historique des évènements
      Trace := TStringList.Create;
      S1 := DateToStr(DateDeb);
      S2 := DateToStr(FindeMois(DateDeb));
      Trace.add('Génération de ' + IntToStr(L.NbSelected) + ' virements de paie du ' + S1 + ' au ' + S2);
      CreeJnalEvt('001', '007', 'OK', nil, nil, Trace);
      Trace.free;
      // FIN PT11
      PeriodeChange := True;
      if Btn <> nil then Btn.Click;
      if not TFmul(Ecran).Q.eof then //PT2  Remplacé pour eagl if TFmul(Ecran).Q.IsEmpty then
      begin
        Chk := TCheckBox(GetControl('PVI_TOPREGLE'));
        if Chk <> nil then
          if Chk.State <> cbGrayed then Chk.State := cbGrayed;
        PeriodeChange := True;
        if Btn <> nil then Btn.Click;
      end;
    end
    else
    begin
      // PT11 : 04/06/2002 V582 PH Gestion historique des évènements
      Trace := TStringList.Create;
      S1 := DateToStr(DateDeb);
      S2 := DateToStr(FindeMois(DateDeb));
      Trace.add('Erreur génération des virements de paie du ' + S1 + ' au ' + S2);
      CreeJnalEvt('001', '007', 'ERR', nil, nil, Trace);
      Trace.free;
      // FIN PT11
    end;
  end
  else
    if rib <> '' then
  begin
    if GetControlText('VCBXSUPPORT') = 'IMP' then
      PgiBox('Aucun enregistrement à imprimer!', 'Génération des Virements')
    else PgiBox('Aucun enregistrement à générer!', 'Génération des Virements');
  end;
end;


procedure TOF_PGVIREMENTS.ChangePeriode(Sender: TObject);
var
  cbMois, cbAnnee: THValComboBox;
  Mois: WORD;
  Annee: string;
begin
  PeriodeChange := True;
  Mois := 1;
  cbMois := THValComboBox(GetControl('CBMOIS'));
  cbAnnee := THValComboBox(GetControl('CBANNEE'));
  if cbMois.Value <> '' then Mois := Trunc(StrToInt(cbMois.Value));
  //if cbAnnee.Value<>'' then Annee:=Trunc(StrToInt(cbAnnee.Text));
  if (cbMois.Value <> '') and (cbAnnee.Value <> '') then
  begin
    ControlMoisAnneeExer(cbMois.value, RechDom('PGANNEESOCIALE', cbAnnee.Value, FALSE), Annee);
    DatePerEncours := EncodeDate(StrToInt(Annee), Mois, 1);
    AffectDatePayeLe(DatePerEncours);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/01/2008 / PT27
Modifié le ... :   /  /
Description .. : Sur changement du type de bulletin, on top le PeriodeChange
Suite ........ : à vrai pour forcer le rafraîchissement des lignes de virement.
Mots clefs ... :
*****************************************************************}
procedure TOF_PGVIREMENTS.OnChangeTypeBul (Sender: TObject);
begin
  PeriodeChange := True;
end;


// Affect la plus petite et plus grande date de paiement selon la periode
// selectionnée
procedure TOF_PGVIREMENTS.AffectDatePayeLe(PerEnCours: TdateTime);
var
  Edit, Edit_, Echeance: THEdit;
  Q: TQuery;
begin
  Edit := THEdit(GetControl('PVI_PAYELE'));
  Edit_ := THEdit(GetControl('PVI_PAYELE_'));
  if (edit = nil) and (edit_ = nil) then exit;
  Q := OpenSql('SELECT MIN(PPU_PAYELE) AS MIN_PALELE,MAX(PPU_PAYELE) AS MAX_PALELE ' +
    'FROM PAIEENCOURS WHERE PPU_PGMODEREGLE ="008" ' + // PT19
    'AND PPU_DATEFIN>="' + USDateTime(PerEnCours) + '" ' +
    'AND PPU_DATEFIN<="' + USDateTime(FindeMois(PerEnCours)) + '" ' +
    'AND PPU_PAYELE>"' + USDateTime(idate1900) + '"', True); // PT7 Ajout Cond.
  Edit.Text := Q.FindField('MIN_PALELE').AsString;
  Edit_.Text := Q.FindField('MAX_PALELE').AsString;
  //PORTAGECWAS
  if not IsValidDate(Edit.text) then Edit.text := DateToStr(DebutDeMois(PerEnCours));
  if not IsValidDate(Edit_.text) then Edit_.text := DateToStr(FinDeMois(PerEnCours));
  Ferme(Q);

  Echeance := THEdit(GetControl('ECHEANCE'));
  if (Echeance <> nil) then
    Echeance.text := Edit_.text;
end;

procedure TOF_PGVIREMENTS.OnClose;
begin
  ExecuteSQL('DELETE FROM VIREMENTS');
  If Assigned(RecupTobVir) Then FreeAndNil(RecupTobVir); //PT26
end;

procedure TOF_PGVIREMENTS.LanceEtatVir(ReglePaie: Boolean);
var
  //  Combo,Combo2 : THValComboBox;
  Pages: TPageControl;
  StSql, Temp, ST2: string;
begin
  { DEB PT6 Mise en commentaire
  Combo := THValComboBox(GetControl('PVI_RIBSALAIRE'));
  Combo2 := THValComboBox(GetControl('VCBXRIBSAL'));
  if (Combo<>nil) and (Combo2<>nil) then
    begin
    Temp:=Combo.Name;
    Temp2:=Combo2.Name;
    Combo.Name:='AAA';             
    Combo2.Name:='BBB';
    Combo.Name:=Temp2;
    Combo2.name:=Temp;
    end; FIN PT6 }
  temp := '';
  Pages := TPageControl(GetControl('Pages'));
  Temp := RecupWhereCritere(Pages);
  if Pos('PVI_RIBSALAIRE', Temp) > 0 then //DEB PT8
  begin
    Insert('(', Temp, Pos('PVI_RIBSALAIRE', Temp));
    ST2 := Trim(Copy(Temp, Pos('PVI_RIBSALAIRE', Temp), Length(temp)));
    Insert('OR PVI_RIBSALAIREEMIS="' + GetControlText('VCBXRIBSAL') + '") ', Temp, (Pos('PVI_RIBSALAIRE', temp) + Pos('AND', ST2) - 1));
    //  Insert(Source: string; var S: string; Index: Integer);
  end; //FIN PT8
  { DEB PT13-2 Suite génération on édite que les règlements générés }
  if ReglePaie then
  begin
    if Pos('WHERE', Temp) > 0 then temp := Temp + ' AND PVI_TOPENCOURS="X"'
    else temp := 'WHERE PVI_TOPENCOURS="X"';
  end;
  { FIN PT13-2 }
  StSql := 'SELECT VIREMENTS.*,R_ETABBQ,R_NUMEROCOMPTE,R_GUICHET,R_CLERIB,R_DOMICILIATION ' +
    //'BQ_ETABBQ,BQ_CLERIB,BQ_GUICHET,BQ_NUMEROCOMPTE,BQ_LIBELLE '+    //PT8 Commentaire
  'FROM VIREMENTS ' +
    'LEFT JOIN SALARIES ON PSA_SALARIE=PVI_SALARIE ' + //
  'LEFT JOIN RIB ON R_AUXILIAIRE=PSA_AUXILIAIRE AND R_SALAIRE="X" ' +
    //'LEFT JOIN BANQUECP ON PVI_RIBSALAIREEMIS=BQ_GENERAL '+         //PT8 Commentaire
  '' + Temp + '' +
    'ORDER BY PVI_ETABLISSEMENT,PVI_DATEDEBUT,PVI_DATEFIN,PVI_RIBSALAIRE,PVI_SALARIE';  { PT23 }   //PT25

  LanceEtat('E', 'PRG', 'PGV', True, False, False, pages, StSql, '', False);    

  {DEB PT6 Mise en commentaire
  if (Combo<>nil) and (Combo2<>nil) then
    begin
    Temp:=Combo.Name;
    Temp2:=Combo2.Name;
    Combo.Name:='AAA';
    Combo2.Name:='BBB';
    Combo.Name:=Temp2;
    Combo2.name:=Temp;
    end;  FIN PT6 Mise en commentaire}

end;

function TOF_PGVIREMENTS.GenereFichier(Rib, RibMaj: string): Boolean;
var
  F: TextFile;
  st, StPlus, FileN, BQNEmet, BQCGui, BQNCompte, BQCodeBq, LaDate: string;
  Salarie, Auxi: string;
  Reponse: WORD;
  V6: Vir06;
  V3: Vir03;
  V8: Vir08;
  MontantRemis: Double;
  VCBXOP, VCBXBQ, VCBXSupport: THValComboBox;
  TRIB_Sal, TR: TOB;
  ECH: THEdit;
  Q: TQuery;
  FichierBanq: string;
begin
  result := False;
  // DEB PT21
  if Rib = '' then
  begin
     PgiBox ('Vous n''avez pas renseigné de banque théorique', Ecran.Caption);
     Exit;
  end;
  if RibMaj = '' then
  begin
     PgiBox ('Vous n''avez pas renseigné de banque à émettre', Ecran.Caption);
     Exit;
  end;
  // FIN PT21
  // DEB PT18
  if GetControlText ('CKSEPARAT') = 'X' then Separat := TRUE
  else Separat := FALSE;
  // FIN PT18

  if (VH_Paie.PgSeriaPaie = False) and (V_PGI.Debug = False) then
  begin
    PgiBox('Vous devez sérialiser votre produit pour pouvoir générer un fichier!', 'Sérialisation');
    Exit;
  end;
  MontantRemis := 0;
  VCBXOP := THValComboBox(GetControl('VCBXCODEOP'));
  if VCBXOP = nil then exit;
  VCBXBQ := THValComboBox(GetControl('VCBXRIBSAL'));  { PT17 }
  if VCBXBQ = nil then exit;
  ECH := THEdit(GetControl('ECHEANCE'));
  if ECH = nil then exit;
  VCBXSupport := THValComboBox(GetControl('VCBXSUPPORT'));
  if VCBXSupport = nil then exit
  else if VCBXSupport.value = 'IMP' then exit;
  StPlus:= PGBanqueCP (True); //PT24
  Q := OpenSql('SELECT BQ_REPVIR FROM BANQUECP WHERE BQ_GENERAL="' + RibMaj + '" '+StPlus, True);
  FichierBanq := Q.FindField('BQ_REPVIR').asstring;
  Ferme(Q); //PT26
  LaDate := DateToSTr(NOW);
  if FichierBanq = '' then
  begin
    PGIBox('Vous n''avez pas renseigné de fichier de virement!', 'Génération des virements'); //PT15
    exit;
  end;
  //Recupération du support de virement
  FileN := ConvertiFichierVirement(FichierBanq, VCBXSupport.Value);
  if FileExists(FileN) then
  begin
    reponse := HShowMessage('5;;Voulez-vous supprimer le fichier de virements ' + ExtractFileName(FileN) + ';Q;YN;Y;N', '', '');
    if reponse = 6 then DetruitFichier(FileN)
    else exit;
  end;
  // REcuperation des infos concernant la banque DO
  Q := OpenSQL('SELECT * FROM BANQUECP WHERE BQ_GENERAL="' + VCBXBQ.Value + '"'+StPlus, TRUE);
  if not Q.EOF then
  begin
    BQNEmet := Q.FindField('BQ_NUMEMETVIR').AsString;
    BQCGui := Q.FindField('BQ_GUICHET').AsString;
    BQNCompte := Q.FindField('BQ_NUMEROCOMPTE').AsString;
    BQCodeBq := Q.FindField('BQ_ETABBQ').AsString;
    Ferme(Q);
  end
  else
  begin
    Ferme(Q);
    exit;
  end;
  // Chargement des RIB SALARIES comprenant les rib pour les virements de salaire
  TRIB_Sal := TOB.Create('Les RIB VIREMENTS SALARIES', nil, -1);
  Q := OpenSQL('SELECT R_AUXILIAIRE,R_DOMICILIATION,R_ETABBQ,R_NUMEROCOMPTE,R_GUICHET FROM RIB ' +
    'LEFT JOIN VIREMENTS ON PVI_AUXILIAIRE=R_AUXILIAIRE WHERE R_SALAIRE="X" AND PVI_RIBSALAIRE="' + rib + '"', TRUE); //PT3
  TRIB_Sal.LoadDetailDb('Les RIB salaries', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  AssignFile(F, FileN);
  {$I-}ReWrite(F);
  {$I+}if IoResult <> 0 then
  begin
    PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
    Exit;
  end;
  // Entete du fichier identification du DO
  with V3 do
  begin
    CodeEnr := '03';
    CodeOP := Format_String(VCBXOP.value, 2);
    ZR1 := StringOfChar(' ', sizeof(ZR1)); // Zone reservee
    NumEmet := Format_String(BQNEmet, 6); // numéro emetteur
    ZR2 := StringOfChar(' ', sizeof(ZR2)); // Zone reservee
    PayeLe := Copy(ECH.Text, 1, 2) + Copy(ECH.Text, 4, 2) + Copy(ECH.Text, 10, 1); // Date de paiement jjmma
    RaisonS := Format_String(GetParamSoc('SO_LIBELLE'), 24); // BQ_LIBELLE  raison sociale emetteur
    RefRem := StringOfChar(' ', sizeof(RefRem)); // reference de la remise
    ZR3 := StringOfChar(' ', sizeof(ZR3)); // Zone reservee
    if VH_Paie.PGTenueEuro = True then Monnaie := 'E' // monnaie de la remise
    else if VH_Paie.PGMonnaieTenue = 'FRF' then Monnaie := 'F' else Monnaie := ' ';
    ZR4 := StringOfChar(' ', sizeof(ZR4)); // Zone reservee
    CodeGui := Format_String(BQCGui, 5); // BQ_GUICHET code quichet DO
    NumCpte := Format_String(BQNCompte, 11); // BQ_NUMEROCOMPTE numero de cpte DO
    ZR5 := StringOfChar(' ', sizeof(ZR5)); // Zone reservee
    Identif := StringOfChar(' ', sizeof(Identif)); // Identifiant DO
    ZR6 := StringOfChar(' ', sizeof(ZR6)); // Zone reservee
    Banque := Format_String(BQCodeBq, 11); // BQ_ETABBQ code banque du DO
    ZR7 := StringOfChar(' ', sizeof(ZR7)); // Zone reservee
  end;
  St := V3.CodeEnr + V3.CodeOP + V3.ZR1 + V3.NumEmet + V3.ZR2 + V3.PayeLe + V3.RaisonS + V3.RefRem
    + V3.ZR3 + V3.Monnaie + V3.ZR4 + V3.CodeGui + V3.NumCpte + V3.ZR5 + V3.Identif + V3.ZR6 + V3.Banque + V3.ZR7;
  if Separat then writeln(F, St)
  else write (F, st) ; // PT18 
  // Boucle sur la query = liste des virements à emettre = creation des enregistrements 06
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).Fetchlestous;
  {$ENDIF}
  TFmul(Ecran).Q.First;
  while not TFmul(Ecran).Q.EOF do
  begin
    Salarie := TFmul(Ecran).Q.FindField('PVI_SALARIE').AsString;
    Auxi := TFmul(Ecran).Q.FindField('PVI_AUXILIAIRE').AsString;
    { PT13-2 On top en cours de génération les règlements traités }
    ExecuteSql('UPDATE VIREMENTS SET PVI_TOPENCOURS="X" ' +
      'WHERE PVI_SALARIE="' + salarie + '" ' +
      'AND PVI_ETABLISSEMENT="' + TFmul(Ecran).Q.FindField('PVI_ETABLISSEMENT').AsString + '" ' +
      'AND PVI_DATEDEBUT="' + UsDateTime(TFmul(Ecran).Q.FindField('PVI_DATEDEBUT').AsDateTime) + '" ' +
      'AND PVI_DATEFIN="' + UsDateTime(TFmul(Ecran).Q.FindField('PVI_DATEFIN').AsDateTime) + '"');
    TR := TRIB_Sal.FindFirst(['R_AUXILIAIRE'], [Auxi], FALSE);
    if TR <> nil then
    begin
      with V6 do
      begin
        CodeEnr := '06';
        CodeOP := Format_String(VCBXOP.value, 2);
        ZR1 := StringOfChar(' ', sizeof(ZR1)); // Zone reservee
        NumEmet := Format_String(BQNEmet, 6); // numéro emetteur
        RefInter := StringOfChar(' ', sizeof(RefInter));
        NomDest := Format_String(TFmul(Ecran).Q.FindField('PVI_LIBELLE').AsString, 24); // Nom du destinataire Nom Salarie
        Domicil := Format_String(TR.GetValue('R_DOMICILIATION'), 20); // Domiciliation salarie R_DOMICILIATION
        NatEco := StringOfChar(' ', sizeof(NatEco)); // Nature Eco pour N.R
        Pays := StringOfChar(' ', sizeof(Pays)); // Pays Eco pour N.R
        BalPay := StringOfChar(' ', sizeof(BalPay)); // Balance des paiements
        CodeGui := Format_String(TR.GetValue('R_GUICHET'), 5); // R_GUICHET code quichet salarie
        NumCpte := Format_String(TR.GetValue('R_NUMEROCOMPTE'), 11); // R_NUMEROCOMPTE numero de cpte salarie
        Montant := PGZeroAGauche(FloatToStr(Round(100.0 * TFmul(Ecran).Q.FindField('PVI_MONTANT').AsFloat)), 16); // Montant net à payer PVI_MONTANT
        // PT10 13/05/2002 PH V575 Fiche Q n°10099 : suppression : pour traitement internet
        Libelle := Format_String('Salarie  ' + TFmul(Ecran).Q.FindField('PVI_SALARIE').AsString, 29); // Libellé := SALARIE : Numero du salarie
        CodeRej := StringOfChar(' ', sizeof(CodeRej)); // Code rejet
        Banque := Format_String(TR.GetValue('R_ETABBQ'), 5); // R_GUICHET code quichet salarie
        ZR2 := StringOfChar(' ', sizeof(ZR2)); // Zone reservee
      end; // FIN with
      MontantRemis := MontantRemis + TFmul(Ecran).Q.FindField('PVI_MONTANT').AsFloat;
      St := V6.CodeEnr + V6.CodeOP + V6.ZR1 + V6.NumEmet + V6.RefInter + V6.NomDest + V6.Domicil + V6.NatEco
        + V6.Pays + V6.BalPay + V6.CodeGui + V6.NumCpte + V6.Montant + V6.Libelle + V6.CodeRej + V6.Banque + V6.ZR2;
      if Separat then writeln(F, St)
      else write (F, St); // PT18
    end // FIN if
    else
    begin //  salarie sans rib.
      PGIError('Le salarié ' + Salarie + ' ne possède aucun rib!', 'Abandon du traitement');
      CloseFile(F);
      DetruitFichier(FileN);
      ExecuteSql('UPDATE VIREMENTS SET PVI_TOPENCOURS="-"'); // PT13-2 RAZ
      Exit;
    end;
    TFmul(Ecran).Q.Next;
  end; // FIN WHILE boucle sur la QUERY
  // Enregistrement 08 Total de la remise
  with V8 do
  begin
    CodeEnr := '08';
    CodeOP := Format_String(VCBXOP.value, 2);
    ZR1 := StringOfChar(' ', sizeof(ZR1)); // Zone reservee
    NumEmet := Format_String(BQNEmet, 6); // numéro emetteur
    ZR2 := StringOfChar(' ', sizeof(ZR2)); // Zone reservee
    Montant := PGZeroAGauche(FloatToStr(Round(100.0 * MontantRemis)), 16); // Montant total de la remise
    ZR3 := StringOfChar(' ', sizeof(ZR3)); // Zone reservee
  end;
  St := V8.CodeEnr + V8.CodeOP + V8.ZR1 + V8.NumEmet + V8.ZR2 + V8.Montant + V8.ZR3;
  if Separat then writeln(F, St)
  else write (F, st); // PT18
  CloseFile(F);
  if TRIB_Sal <> nil then TRIB_Sal.Free;
  PGIBox('Génération effectuée sous le fichier : ' + FileN + '', 'Génération des virements');
  result := True;
end;

procedure TOF_PGVIREMENTS.DetruitFichier(NomFic: string);
begin
  DeleteFile(PChar(NomFic));
end;

//Affect le code banque DO de l'initial au mise à jour
//sur le onchange de la banque initial
procedure TOF_PGVIREMENTS.AffectRib(Sender: TObject);
{$IFNDEF EAGLCLIENT}
var
  Vcbx, VcbxRib: ThValComboBox;
{$ENDIF}
begin
  {$IFNDEF EAGLCLIENT}
  Vcbx := ThValComboBox(GetControl('PVI_RIBSALAIRE'));
  VcbxRib := ThValComboBox(GetControl('VCBXRIBSAL'));
  if (Vcbx <> nil) and (VcbxRib <> nil) then
    SetControlText('VCBXRIBSAL', GetControlText('PVI_RIBSALAIRE'));
  // VcbxRib.Value:=Vcbx.Value;
  {$ENDIF}
  PeriodeChange := True;
end;

procedure TOF_PGVIREMENTS.CalcRemise(Sender: TObject);
var
  Remise: THNumEdit;
  Q: THQuery;
  Mont: Double;
  {$IFNDEF EAGLCLIENT}
  L: THDBGrid;
  {$ELSE}
  i: Integer;
  L: THGrid;
  {$ENDIF}
begin
  Mont := 0;
  {$IFNDEF EAGLCLIENT}
  L := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  L := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  Remise := THNumEdit(GetControl('REMISE'));
  Q := THQuery(Ecran.FindComponent('Q'));
  {$IFDEF EAGLCLIENT}
  i := TFMul(Ecran).FListe.Row;
  if i > 1 then
    TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1)
  else Exit;
  {$ENDIF}
  if (Remise <> nil) and (Q <> nil) and (L <> nil) then
    Mont := Q.FindField('PVI_MONTANT').AsFloat;
  if (Mont > 0)
    {$IFNDEF EAGLCLIENT}
  and (L.IsCurrentSelected = True)
    {$ENDIF}
  then
    Remise.Value := Remise.Value - Mont
  else
    Remise.Value := Remise.Value + Mont;
end;


function TOF_PGVIREMENTS.RecupCritereVir(Avec: boolean): string;
var
  St, val: string;
  Chk: TCheckBox;
begin
  result := '';
  if Avec = True then St := 'WHERE PVI_SALARIE<>"00000000000" ' else St := '';
  val := GetControlText('PVI_PAYELE');
  if val <> '' then St := St + 'AND PVI_PAYELE>="' + USDateTime(StrToDate(val)) + '" ';
  val := GetControlText('PVI_PAYELE_');
  if val <> '' then St := St + 'AND PVI_PAYELE<="' + USDateTime(StrToDate(val)) + '" ';
  val := GetControlText('PVI_SALARIE');
  if val <> '' then St := St + 'AND PVI_SALARIE>="' + val + '" ';
  val := GetControlText('PVI_SALARIE_');
  if val <> '' then St := St + 'AND PVI_SALARIE<="' + val + '" ';
  val := GetControlText('PVI_ETABLISSEMENT');
  if val <> '' then St := St + 'AND PVI_ETABLISSEMENT>="' + val + '" ';
  val := GetControlText('PVI_ETABLISSEMENT_');
  if val <> '' then St := St + 'AND PVI_ETABLISSEMENT<="' + val + '" ';
  Chk := TCheckBox(GetControl('PVI_TOPREGLE'));
  if Chk <> nil then
  begin
    if Chk.State = cbGrayed then val := '';
    if Chk.State = cbChecked then val := 'X';
    if Chk.State = cbUnchecked then val := '-';
  end;
  if val = 'X' then St := St + 'AND PVI_TOPREGLE="' + val + '" ';
  if val = '-' then St := St + 'AND (PVI_TOPREGLE="' + val + '" OR PVI_TOPREGLE="") ';
  result := St;
end;
{
 procedure TOF_PGVIREMENTS.LanceCritere(Sender: TObject);
begin
if GetControlText('PVI_RIBSALAIRE')='' then
  Begin
  PgiBox('Vous devez renseigner la banque du donneur d''ordre!','Virement des salaires');
  exit;
  End;
  TFMul(Ecran).ChercheClick;
end;
}
procedure TOF_PGVIREMENTS.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;


//PT26 - Début
{ TOF_PGVIREMENTDOS }

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/12/2007
Modifié le ... :   /  /
Description .. : Ouverture de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_PGVIREMENTDOS.OnArgument(stArgument: string);
Var
  Cb : THValComboBox;
  Btn : TToolbarButton97;
begin
  inherited;
  Cb := THValComboBox(GetControl('MULTIDOSSIER'));
  If Assigned(Cb) Then Cb.OnChange := ChangeRegroupement;

  Grille := THGrid(GetControl('GRLISTEDOSSIERS'));

  Btn := TToolbarButton97(GetControl('bSelectAll'));
  If Btn <> Nil Then Btn.OnClick := SelectAll;

  Btn := TToolbarButton97(GetControl('BValider'));
  if Btn <> nil then Btn.OnClick := LanceGeneration;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/12/2007
Modifié le ... :   /  /
Description .. : Chargement des informations de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_PGVIREMENTDOS.OnLoad;
begin
  inherited;
  TFVierge(Ecran).HMTrad.ResizeGridColumns (Grille);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/12/2007
Modifié le ... :   /  /
Description .. : Validation de l'écran : lance autant d'applications que 
Suite ........ : nécessaires à la génération des virements sur chaque 
Suite ........ : dossier sélectionné.
Mots clefs ... :
*****************************************************************}
procedure TOF_PGVIREMENTDOS.LanceGeneration;
var
    Dossier,DBPrefixe : String;
    NbDos   : Integer;
    ListeBox : TListBox;
    i,NumDos : Integer;
    LibMessage : String;
    Q : TQuery;
    Avant,Apres : Integer;
    MyWidth, MyHeight, MyTop, MyLeft : Integer;
    Taillo : String;
begin
     If (Grille.nbSelected = 0) And (Grille.AllSelected = False) Then
     Begin
        PGIBox(TraduireMemoire('Veuillez sélectionner les sociétés à traiter.'), Ecran.Caption);
        Exit;
     End
     Else
     Begin
        // On bloque tout pour ne pas que l'utilisateur puisse intéragir
        SetControlEnabled('BValider',   False);
        SetControlEnabled('BFerme',     False);
        SetControlEnabled('BSelectAll', False);
        
        ListeBox := TListBox(GetControl('LBANO'));
        If ListeBox <> Nil Then ListeBox.Clear;

        // Création de la liste des anomalies pour suivre le déroulement du traitement
        ListeAnomalies := TAnomalies.Create;
        ListeAnomalies.SetDuplicateSymbols := True;
        ListeAnomalies.SetUnderlinedTitles := True;
        ListeAnomalies.ChangeLibAno(INFO1, TraduireMemoire('Déroulement du traitement :'));
        ListeAnomalies.ChangeLibAno(ERR1,  TraduireMemoire('Dossiers non traités :'));

        NbDos := Grille.NbSelected;

        // Démarrage de l'écran d'attente (6 messages par dossier : 1 à vide, 3 de chargement, 1 de lancement de la fiche, 1 de déconnexion)
        InitMoveProgressForm(Nil, TraduireMemoire('Traitement en cours. Veuillez patienter...'), TraduireMemoire('Génération des virements.'), NbDos*6,  True, True);

        // On lance le traitement sur chaque dossier
        NumDos := 0;
        For i := 1 To Grille.RowCount-1 Do
        Begin
            If Grille.IsSelected(i) Then
                Dossier := Grille.CellValues[1,i]
            Else
                Continue;

            NumDos := NumDos + 1;
            DBPrefixe := Dossier + '.DBO.';

            LibMessage := Format(TraduireMemoire('Traitement du dossier %s / %s : %s'), [IntToStr(NumDos),IntToStr(NbDos),Dossier]);
            ListeAnomalies.Add(INFO1, LibMessage);

            TextsProgressForm((NumDos-1)*6,TraduireMemoire('Traitement en cours. Veuillez patienter...'),LibMessage, '');

            If Not MoveCurProgressForm('') Then
            Begin
                // Clic sur abandon
                If Q <> Nil Then Ferme(Q);
                ListeAnomalies.Add(INFO1, '');
                ListeAnomalies.Add(INFO1, TraduireMemoire('>>> Arrêt du traitement par l''utilisateur <<<'));
                Break;
            End;
            
            // On ne sait pas à l'avance ce que va faire l'utilisateur comme règlement. Le mieux est donc de comparer
            // le nombre d'éléments réglés avant et après le lancement de l'application
            Avant := 0;
            Q := OpenSQL('SELECT COUNT(*) AS CPTREGLE FROM '+DBPrefixe+'PAIEENCOURS WHERE PPU_TOPREGLE="X"', True);
            If Not Q.EOF Then Avant := Q.FindField('CPTREGLE').AsInteger;
            Ferme(Q);

            // Taille et position de la fenêtre à ouvrir
            MyTop    := PositionMul.Y;
            MyLeft   := PositionMul.X;
            MyWidth  := TForm(Ecran).Width;
            MyHeight := TForm(Ecran).Height;
            Taillo := Format('%d;%d;%d;%d', [MyTop, MyLeft, MyWidth, MyHeight]);

            // Lancement de l'exécutable
            LanceApplication (ProcessInfo, Dossier, '/FICHE=VIREMENT_MUL&&'+Taillo, True);

            Apres := 0;
            Q := OpenSQL('SELECT COUNT(*) AS CPTREGLE FROM '+DBPrefixe+'PAIEENCOURS WHERE PPU_TOPREGLE="X"', True);
            If Not Q.EOF Then Apres := Q.FindField('CPTREGLE').AsInteger;
            Ferme(Q);

			// Le dossier a-t-il été traité?            
            If Avant < Apres Then
                ListeAnomalies.Add(INFO1, '   Virements effectués : '+IntToStr(Apres-Avant))
            Else
            Begin
                ListeAnomalies.Add(ERR1, '    '+Dossier);
                ListeAnomalies.Add(INFO1, '   Aucun virement n''a été effectué');
            End;

			ListeBox.Clear;
			ListeAnomalies.PutInList(ListeBox);
        End;

        // Fermeture de l'écran d'attente
        FiniMoveProgressForm;

        ListeAnomalies.Add(INFO1, '');
        ListeAnomalies.Add(INFO1, TraduireMemoire('>> Traitement terminé <<'));

		ListeBox.Clear;	
        ListeAnomalies.PutInList(ListeBox);
        FreeAndNil(ListeAnomalies);

        // On débloque les boutons
        SetControlEnabled('BValider',   True);
        SetControlEnabled('BFerme',     True);
        SetControlEnabled('BSelectAll', True);
     End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 12/12/2007 / PT26
Modifié le ... :   /  /
Description .. : Gestion du changement de regroupement par l'utilisateur
Mots clefs ... :
*****************************************************************}
Procedure TOF_PGVIREMENTDOS.ChangeRegroupement (Sender : TObject);
var i : Integer;
Begin
	// Réinitialisation de la grille
    If Grille.AllSelected Then TToolbarButton97(GetControl('BSelectAll')).Click;
    For i := 1 to Grille.RowCount - 1 do Grille.Rows[i].Clear;
    Grille.RowCount := 2;
    
    // Affichage des dossiers
    If GetControlText('MULTIDOSSIER') <> '' Then
    Begin
        AfficheListeDossiers;
        SetFocusControl('GRLISTEDOSSIERS');
    End;
    
    // Réinitialisation du bouton Sélectionner Tout au cas où
    TToolbarButton97(GetControl('bSelectAll')).Down := False;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 12/12/2007 / PT26
Modifié le ... :   /  /
Description .. : Charge la grille avec la liste des dossiers du regroupement
Mots clefs ... :
*****************************************************************}
Procedure TOF_PGVIREMENTDOS.AfficheListeDossiers;
Begin
	If Not ChargeDossiersDuRegroupement (GetControlText('MULTIDOSSIER'), Grille, 'VIREMENT', '', '') Then //PT1
        Exit;

    // Adaptation de la taille des colonnes
    TFVierge(Ecran).HMTrad.ResizeGridColumns (Grille);
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 12/12/2007 / PT26
Modifié le ... :   /  /
Description .. : Bouton "Sélectionner tout"
Mots clefs ... :
*****************************************************************}
procedure TOF_PGVIREMENTDOS.SelectAll(Sender: TObject);
begin
    If GetControlText('MULTIDOSSIER') <> '' Then 
    	Grille.AllSelected := Not Grille.AllSelected;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/12/2007
Modifié le ... :   /  /
Description .. : Fermeture de l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_PGVIREMENTDOS.OnClose;
begin
    // Si on a ouvert la fiche salarié d'un autre dossier, on bloque la fermeture de la fenêtre mul
    if  (ProcessInfo.hProcess <> 0) Then
        LastError := 1
    Else
        LastError := 0;
end;
//PT26 - Fin

initialization
  registerclasses([TOF_PGVIREMENTS,TOF_PGVIREMENTDOS]);  //PT26

end.

