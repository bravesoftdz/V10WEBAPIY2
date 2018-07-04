{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 07/12/2001
Modifié le ... :   /  /
Description .. : Multi critère gestion des infos complémentaires des salariés
Mots clefs ... : PAIE
*****************************************************************}
{
PT1-1 : 19/03/2002 VG V571 Ajout test si TOrg  <>  nil
PT1-2 : 19/03/2002 VG V571 Gestion de la DADSU BTP : Ajout de champs dans la
                           table DEPORTSAL
PT2   : 22/03/2002 VG V571 Pour le multicritère BTP, il ne faut pas pouvoir
                           changer de liste de salariés (MSA, ...)
PT3   : 27/03/2002 VG V571 Pour le multicritère BTP, l'affectation d'un salarié
                           remplit par défaut la date d'ancienneté dans la
                           profession par la date d'ancienneté du salarié
PT4   : 04/04/2002 JL V571 Modification procédure OnArgument car point d'entré
                           différent pour chaque cas dans le menu.
PT5   : 04/04/2002 JL V571 Modification des requêtes de MAJ pour valeurs champs
                           par défaut
PT6   : 13/09/2002 JL V582 Modification code pour E-AGL
PT7   : 20/03/2002 JL V_42 Modification Maj de la base : suppression des
                           requêtes de Maj remplacés par des tobs car tous les
                           champs n'étaient pas initialisés
                           Ajout de fonction et procédure pour simplifier cette
                           Maj.
PT8   : 13/05/2003 MF V_42 Gestion des tickets restaurant
PT9   : 24/10/2003 JL V_42 Modif traitement par lot CWAS
PT10  : 26/03/2004 JL V_50 Gestion MSA EDI : Nouveaux champs : PSE_MSATECHNIQUE
                           au lieu de PSE_MSAINFOSCOMPL
PT11  : 20/08/2004 VG V_50 FQ N°11483 nécessitant une MAJ de structure.
                           Code écrit et laissé en commentaire en attente
PT12  : 02/06/2005 JL V_60 MSA : message si activité établissement non renseigné
PT13  : 29/06/2005 JL V_60 Gestion responsable dans interimaire
PT14  : 11/10/2005 PH V_60 pour le paramsoc SO_IFDEFCEGID dans le cas de PCL on fait un
                           GetParamsocSecur au lieu de GetParamSoc pour éviter les erreurs
PT15  : 15/03/2006 JL V_650 FQ 11722 Gestion salariés sorties pour ePaie
PT16  : 04/04/2006 JL V_65 gestion UGMSA
PT17  : 05/06/2006 JL V_65 Gestion test date sortie pour intermittent et msa
PT18  : 09/06/2006 JL V_65 FQ 13216 pse_isnumident initialisé avec numéro SS
PT19  : 20/10/2006 JL V_70 Q_Mul n'était plus affecté
PT20  : 08/01/2007 GG V_70 FQ 13590 - MUL liste des salariés : case à cocher
                           Exclure les salariés sortis ne fonctionne pas ; le
                           champ Date d'arrêté n'est jamais accessible
        02/07/2007 GG V_70 FQ 13590 - La vue utilisé pour les salariés affectés au BTP
                           ne contient pas le champs PSA_DATESORTIE 
PT21  : 22/01/2007 FC V_70 FQ 13166 - Pouvoir affecter en masse des salariés à un service
                           lorsque l'on vient du menu Affectation des salariés (GRH)
PT22  : 24/01/2007 FC V_70 Gestion des habilitations : gestion du nom des champs critère établissement
                           suivant d'où l'on vient
PT23  : 19/02/2007 MF V_70 FQ 13586 : Sur liste des non affectés, cacher le bouton tout sélectionner.
PT24  : 19/09/2007 FL V_80 En multidossiers, ne pas afficher les combos établissement
PT25  : 02/04/2008 FL V_80 Adaptation pour les restrictions multidossiers
PT26  : 03/04/2008 FL V_80 Ajout du libellé service
}
unit UTofPG_MulDeportSal;

interface
Uses Controls,Classes,sysutils,ComCtrls,HTB97,UTOF,
{$IFNDEF EAGLCLIENT}
      HDB,Mul,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     emul,eFiche,UtileAGL,
{$ENDIF}
     HMsgBox,HCtrls,HEnt1,Hqry,UTOB,LookUp, ParamSoc, StdCtrls;

Type
  TOF_PGMULDEPORTSAL = Class (TOF)
      procedure OnLoad                   ; override ;
      procedure OnArgument(S : String )  ; override ;
      private
      arg : String;
      {$IFNDEF EAGLCLIENT}
      Liste : THDBGrid;
      {$ELSE}
      Liste : THGrid;
      {$ENDIF}
      Affectation : THValComboBox;
      Baffecter : TToolBarButton97;
      Q_Mul : THQuery ;
      THSal1,THSal2 : THEdit;
      THEtab1,THEtab2 : THEdit; //PT22
      CChoix : THValComboBox;
      LesEtabMsa : String;       //PT12
      CDistribution  : THMultiValComboBox; // PT8
      LDistribution  : THLabel; // PT8
      Q_Service : TQuery; //PT21
      LabelService : THLabel;         //PT21
      ComboService : THEdit;   //PT21 //PT25
      LibelleService : THLabel; //PT26
      Procedure AffecterClick(Sender : TObject);
      Procedure ExitEdit(Sender : TObject);
      Procedure OnClickSalarieSortie(Sender: TObject);
      Procedure ChangeChoix(Sender : TObject);   //15 MARS
      Procedure ExitChoix(Sender : TObject);     //15 MARS
      Function VerifChoix(Choix : String) : Boolean; //15 MARS
      Procedure RespElipsisClick(Sender : TObject);
      //DEBUT PT7
      procedure CreationEnregistrement(TypeEnreg,Salarie : String);
      Function TrouveInfosComplMSA(SalarieMSA : String) : String;
      Function TrouveActiviteMSA(SalarieMSA : String) : String;
      Function TrouveDateAncBTP(SalarieBTP : String) : TDateTime;
      Function TrouveUGMSA(SalarieMSA : String) : String;//PT16
      // FIN PT7
      Procedure AfficherComboService(Sender : TObject); //PT21
      Procedure ChargeListeSalaries (Sender : TObject); //PT25
      end ;

implementation

uses  EntPaie,P5Def,PGoutils2,PgEditOutils2,PGOutilsFormation,PGHierarchie;

{ TOF_PGMULDEPORTSAL }

procedure TOF_PGMULDEPORTSAL.OnLoad;
var
  TOrg : TTabSheet;
  Champ,ListeA,TitreA,ListeNA,TitreNA, StWhere : String;
  DateArret : TDateTime;
  Dossier : String;
begin
	StWhere:= '';         //PT11

	If Arg <> 'RESP' then                                       //15 MARS
	begin
		If VerifChoix(CChoix.Value)=True Then
		begin
			LastError:= 1;
			PgiBox ('Ce type de saisie complémentaire n''est pas géré pour ce dossier',
			'Saisie complémentaire salarié');
		end;
	end;

	TOrg := TTabSheet(GetControl('PORGANISATION'));
	//PT1-1 If Arg <> 'RESP' Then TOrg.TabVisible := False;
	If (Arg <> 'RESP') and (TOrg  <>  nil) Then
	TOrg.TabVisible:= False;

	If Affectation.Value = 'AFFECTE' then
	begin
		THSal1.Name:= 'PSE_SALARIE';
		THSal2.Name:= 'PSE_SALARIE_';
		// d PT8
		{Code de distribution spécifique Tickets restaurant,visible si affecté}
		if (Arg = 'TCK') then
		begin
			CDistribution.Visible:= True;
			LDistribution.Visible:= True;
		end;
		// f PT8
		BAffecter.Visible:= False;
		TFMul(Ecran).BSelectAll.Visible:= False;
		If TFMul(Ecran).BSelectAll.Down=True Then
		begin
			TFMul(Ecran).BSelectAll.Click;
			TFMul(Ecran).BSelectAll.Down:= False;
		end;
	end
	Else
	begin
		if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) And (Arg = 'RESP') then   //PT13 //PT24
		begin
			THSal1.Name:= 'PSI_INTERIMAIRE';
			THSal2.Name:= 'PSI_INTERIMAIRE_';
		end
		else
		begin
			THSal1.Name:= 'PSA_SALARIE';
			THSal2.Name:= 'PSA_SALARIE_';
		end;
		BAffecter.Visible:= True;
		TFMul(Ecran).BSelectAll.Visible:= true; //PT23

		// d PT8
		{Code de distribution spécifique Tickets restaurant, non visible si non affecté}
		if (Arg = 'TCK') then
		begin
			CDistribution.Visible:= False;
			LDistribution.Visible:= False;
			Cdistribution.value:= '';
			TFMul(Ecran).BSelectAll.Visible:= false;       //PT23
			// f PT8
		end;
		//PT23   TFMul(Ecran).BSelectAll.Visible:= true;
	end;

	// AFFECTATION
	ListeNA:= 'PGCOMPLSALARIENA';
	If Arg = 'MSA' Then
	begin
		Champ:= 'PSE_MSA';
		ListeA:= 'PGCOMPLSALARIEMSA';
		TitreA:= 'Liste des salariés affectés à la M.S.A.';
		TitreNA:= 'Liste des salariés non affectés à la M.S.A.';
		if TCheckBox(GetControl('CKSORTIE'))<>nil then //PT17
		Begin
			if (GetControlText('CKSORTIE')='X') and
			(IsValidDate(GetControlText('DATEARRET'))) then
			Begin
				DateArret:= StrtoDate(GetControlText('DATEARRET'));
				StWhere:= ' PSA_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
				'(PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
				' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
				' PSA_DATESORTIE IS NULL) AND'+
				' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'")';
			End;
		End;
		THEtab1.Name:= 'PSA_ETABLISSEMENT';  //PT22
		THEtab2.Name:= 'PSA_ETABLISSEMENT_'; //PT22
	end;

	If Arg = 'IS' Then
	begin
		Champ:= 'PSE_INTERMITTENT';
		ListeA:= 'PGCOMPLSALARIEINT';
		TitreA:= 'Liste des intermittents du spectacle';
		TitreNA:= 'Liste des salariés non intermittents du spectacle';
		if TCheckBox(GetControl('CKSORTIE'))<>nil then //PT17
		Begin
			if (GetControlText('CKSORTIE')='X') and
			(IsValidDate(GetControlText('DATEARRET'))) then
			Begin
				DateArret:= StrtoDate(GetControlText('DATEARRET'));
				StWhere:= ' PSA_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
				'(PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
				' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
				' PSA_DATESORTIE IS NULL) AND'+
				' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'")';
			End;
		End;
		THEtab1.Name:= 'PSA_ETABLISSEMENT';  //PT22
		THEtab2.Name:= 'PSA_ETABLISSEMENT_'; //PT22
	end;

	If Arg = 'RESP' Then
	begin
		Champ:= 'PSE_ORGANIGRAMME';
		if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) And (Arg = 'RESP') then   //PT13 //PT24
		begin
			ListeA:= 'PGDEPORTSALINTERI';
			ListeNA:= 'PGCOMPLSALNAINT';
			//PT15
			if TCheckBox(GetControl('CKSORTIE'))<>nil then
			Begin
				if (GetControlText('CKSORTIE')='X') and
				(IsValidDate(GetControlText('DATEARRET'))) then
				Begin
					DateArret:= StrtoDate(GetControlText('DATEARRET'));
					StWhere:= ' (PSI_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
					' PSI_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
					' PSI_DATESORTIE IS NULL) AND'+
					' PSI_DATEENTREE <="'+UsDateTime(DateArret)+'"';
				End;
			End;
			//FIN PT15
			THEtab1.Name:= 'PSI_ETABLISSEMENT';  //PT22
			THEtab2.Name:= 'PSI_ETABLISSEMENT_'; //PT22
		end
		else
		begin
			ListeA:= 'PGDEPORTSAL';
			//PT15
			if TCheckBox(GetControl('CKSORTIE'))<>nil then
			Begin
				if (GetControlText('CKSORTIE')='X') and
				(IsValidDate(GetControlText('DATEARRET'))) then
				Begin
					DateArret:= StrtoDate(GetControlText('DATEARRET'));
					StWhere:= ' (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
					' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
					' PSA_DATESORTIE IS NULL) AND'+
					' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
				End;
			End;
			//FIN PT15
			THEtab1.Name:= 'PSA_ETABLISSEMENT';  //PT22
			THEtab2.Name:= 'PSA_ETABLISSEMENT_'; //PT22
		end;
		TitreA:= 'Liste des salariés';
		TitreNA:= 'Liste des salariés';

	end;

	//PT1-2
	If Arg = 'BTP' Then
	begin
		Champ:= 'PSE_BTP';
		ListeA:= 'PGCOMPLSALARIEBTP';
		TitreA:= 'Liste des salariés affectés au module B.T.P.';
		TitreNA:= 'Liste des salariés non affectés au module B.T.P.';
		//PT20
		if GetControl('CKSORTIE') <> nil then
		Begin
			if (GetControlText('CKSORTIE')='X') and
			(IsValidDate(GetControlText('DATEARRET'))) then
			Begin
				DateArret:= StrtoDate(GetControlText('DATEARRET'));
				// PT20-2        StWhere:= ' (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
				//                   ' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
				//                   ' PSA_DATESORTIE IS NULL) AND'+
				//                   ' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
				StWhere:= ' PSA_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+
				'(PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
				' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
				' PSA_DATESORTIE IS NULL) AND'+
				' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'")';
			End;
		End;
		//FIN PT20
		THEtab1.Name:= 'PSA_ETABLISSEMENT';  //PT22
		THEtab2.Name:= 'PSA_ETABLISSEMENT_'; //PT22
	end;

	If Arg = 'TCK' Then
	begin
		Champ:= 'PSE_TICKETREST';
		ListeA:= 'PGCOMPLSALARIETCK';
		TitreA:= 'Liste des salariés affectés au module titres restaurant';
		TitreNA:= 'Liste des salariés non affectés au module titres restaurant';

		//PT11
		if TCheckBox(GetControl('CKSORTIE'))<>nil then
		Begin
			if (GetControlText('CKSORTIE')='X') and
			(IsValidDate(GetControlText('DATEARRET'))) then
			Begin
				DateArret:= StrtoDate(GetControlText('DATEARRET'));
				StWhere:= ' (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR'+
				' PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR'+
				' PSA_DATESORTIE IS NULL) AND'+
				' PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
			End;
		End;
		//FIN PT11
		THEtab1.Name:= 'PSA_ETABLISSEMENT';  //PT22
		THEtab2.Name:= 'PSA_ETABLISSEMENT_'; //PT22
	end;
	//FIN PT1-2

	// Choix des listes et des clauses de la requête en fonction tu type de salariés
	// avec les critères définis avant.
	If Affectation.Value = 'AFFECTE' Then
	begin
		{
		TFMul(Ecran).DbListe:= ListeA;
		Q_Mul:= THQuery(Ecran.FindComponent('Q'));
		if Q_Mul <> NIL then
		Q_Mul.Liste:= ListeA;
		}

		TFMul(Ecran).SetDBListe(ListeA);
		TFMul(Ecran).Caption := TitreA;
		UpdateCaption (TFMul(Ecran));
		StWhere:= StWhere+' AND '+Champ+'="X"';
	end
	Else
	begin
		{
		TFMul(Ecran).DbListe:= ListeNA;
		Q_Mul:= THQuery(Ecran.FindComponent('Q'));
		if Q_Mul <> NIL then
		Q_Mul.Liste:= ListeNA;
		}
		TFMul(Ecran).SetDBListe(ListeNa) ;
		TFMul(Ecran).Caption:= TitreNA;
		UpdateCaption (TFMul(Ecran));
		StWhere:= StWhere+' AND ('+Champ+'<>"X" OR PSE_SALARIE IS NULL)';      //PT15
	end;
	
	if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) And (Arg = 'RESP') then   //PT13 //PT24
	begin
		If Arg = 'RESP' Then StWhere := StWhere + ' AND (PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT")';
		//PT25 - Début
		If Not PGDroitMultiForm Then
			Dossier := V_PGI.NoDossier
		Else
			Dossier := GetDossierServiceNum(GetControlText('VSERVICE'));

		If Dossier <> '' Then
		Begin
			If Affectation.Value = 'AFFECTE' Then
				StWhere := StWhere + ' AND PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_NOMSERVICE LIKE "'+Dossier+'%")'
			Else
				StWhere := StWhere + ' AND PSI_NODOSSIER="'+Dossier+'"';
		End
		Else
		Begin
			If Affectation.Value = 'AFFECTE' Then
				StWhere := StWhere + ' AND PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE ' + ServicesDesDossiers + ')'
			Else
				StWhere := StWhere + DossiersAInterroger('DOS','','PSI',False,True);
		End;
		//PT25 - Fin
	end;
	SetControlText('XX_WHERE', StWhere);
end;


procedure TOF_PGMULDEPORTSAL.OnArgument(S : String);
var
  GestionUnique : Boolean;
  Num : Integer;
  Min,Max : String;
  Saisie,LibResp : THLabel;
  THResp : THEdit;
  Check: TCheckBox;
begin
	inherited;
	Arg:= (ReadTokenSt(S));
	If Arg <> 'RESP' Then                           // DEBUT 15 MARS
	begin
		//PT2
		if Arg = 'BTP' then
		begin
			CChoix:= THValComboBox(GetControl('CHOIX'));
			if CChoix  <>  Nil then
			begin
				CChoix.Value:= 'BTP';
				CChoix.Visible:=  False;
			end;
			//PT20
			if (GetControl('CKSORTIE') <> nil) and (GetControl('CKSORTIE') is TCheckBox) then
			(GetControl('CKSORTIE') as TCheckBox).OnClick:= OnClickSalarieSortie;
			//Fin PT20

			// Saisie  :=  THLabel(GetControl('LBLSAISIE')); La zone est rendu invisible dans tous les cas, car points d'entrés différents dans le menu
			// if Saisie  <>  Nil then
			//    Saisie.Visible  :=  False;
		end
		else
		begin
			CChoix:= THValComboBox(GetControl('CHOIX'));
			If CChoix <> Nil Then
			begin
				CChoix.OnChange:= ChangeChoix;
				CChoix.OnExit:= ExitChoix;
			end;
			If Arg = 'MSA' then//VH_Paie.PGMsa=True Then               debut  PT4
			begin
				SetControlText('GRILLE','MSA');
				Arg:= 'MSA';
				CChoix.Value:= 'MSA';
				CChoix.Visible:= False;     // Combo choix non visible car gestion séparée dans le menu
				Check:= TCheckBox(GetControl('CKSORTIE'));  //PT17
				if Check <> nil then Check.OnClick:= OnClickSalarieSortie;
			end
			else
			If Arg = 'IS' Then //VH_Paie.PGIntermittents=True Then
			begin
				SetControlText('GRILLE','IS');
				Arg:= 'IS';
				CChoix.Value:= 'INT';
				CChoix.Visible:= False; // Combo choix non visible car gestion séparée dans le menu
				Check:= TCheckBox(GetControl('CKSORTIE'));  //PT17
				if Check <> nil then Check.OnClick:= OnClickSalarieSortie;
			end
			//       If VH_Paie.PGBTP=True Then
			//          begin
			//          SetControlText('GRILLE','BTP');
			//          Arg := 'BTP';
			//          CChoix.Value := 'BTP';
			//          end;
			else
			// d PT8
			If Arg = 'TCK' Then { Tickets restaurant}
			begin
				SetControlText('GRILLE','TCK');
				Arg:= 'TCK';
				CChoix.Value:= 'TCK';
				CChoix.Visible:= False; // Combo choix non visible car gestion séparée dans le menu
				CDistribution:= THMultiValComboBox(GetControl('PSE_DISTRIBUTION'));
				CDistribution.Visible:= True;
				CDistribution.Enabled:= True;
				CDistribution.Value:= '';
				LDistribution:= THLabel(GetControl('LDISTRIBUTION'));
				LDistribution.Visible:= True;
				LDistribution.Enabled:= True;

				//PT11
				SetControlvisible('CKSORTIE',True);
				SetControlEnabled('CKSORTIE',True);
				SetControlvisible('DATEARRET',True);
				SetControlvisible('TDATEARRET',True);
				SetControlEnabled('DATEARRET',False);
				SetControlEnabled('TDATEARRET',False);
				Check:= TCheckBox(GetControl('CKSORTIE'));
				if Check=nil then
				Begin
					SetControlVisible('DATEARRET',False);
					SetControlVisible('TDATEARRET',False);
				End
				else
				Check.OnClick:= OnClickSalarieSortie;
				//FIN PT11

			end;
			// f PT8
		end;
		//FIN PT2
	end
	Else
	SetControlVisible ('CHOIX',False);  // FIN 15 MARS

	Saisie  :=  THLabel(GetControl('LBLSAISIE'));
	if Saisie  <>  Nil then
	Saisie.Visible:= False;
	SetControlVisible('BINSERT',False);                                       // Fin PT4
	if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) then   //PT13 //PT24
	begin
		for num := 1 to 4 do
		VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),
		GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
		VisibiliteStat(GetControl ('PSI_CODESTAT'),GetControl ('TPSI_CODESTAT'));
	end
	else
	begin
		For Num := 1 to 4 do
		VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),
		GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
		VisibiliteStat(GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT'));
	end;
	Affectation:= THValComboBox(GetControl('VAFFECTATION'));
	If Affectation <> NIL Then
	Affectation.Value:= 'AFFECTE';
	{$IFNDEF EAGLCLIENT}
	Liste:= THDBGrid(GetControl('FLISTE'));
	{$ELSE}
	Liste:= THGrid(GetControl('FLISTE'));
	{$ENDIF}
	Baffecter:= TToolBarButton97(GetControl('BAFFECTER'));
	If BAffecter <> Nil Then
	BAffecter.OnClick:= AffecterClick;

	If Arg <> 'RESP' Then
	begin
		LibResp:= THLabel(GetControl('TPSE_RESPONSABS'));
		If LibResp <> Nil Then
		LibResp.Visible:= False;
		SetControlVisible('PSE_RESPONSABS',False);
	end
	else
	begin
		THResp:= THEdit(GetControl('PSE_RESPONSABS'));
		If THResp <> Nil then
		THResp.OnElipsisClick:= RespElipsisClick;
		//PT15
		SetControlvisible('CKSORTIE',True);
		SetControlEnabled('CKSORTIE',True);
		SetControlvisible('DATEARRET',True);
		SetControlvisible('TDATEARRET',True);
		SetControlEnabled('DATEARRET',False);
		SetControlEnabled('TDATEARRET',False);
		Check:= TCheckBox(GetControl('CKSORTIE'));
		if Check=nil then
		Begin
			SetControlVisible('DATEARRET',False);
			SetControlVisible('TDATEARRET',False);
		End
		else
		Check.OnClick:= OnClickSalarieSortie;
		//FIN PT15
	end;

	THSal1:= THEdit(GetControl('PSE_SALARIE'));
	THSal2:= THEdit(GetControl('PSE_SALARIE_'));
	If THSal1 <> Nil Then
	THSal1.OnExit:= ExitEdit;
	If THSal2 <> Nil Then
	THSal2.OnExit:= ExitEdit;
	GestionUnique:= False;

	//DEB PT22
	If (Arg = 'TCK') or (Arg = 'BTP') or (Arg = 'IS') or (Arg = 'MSA') Then
	begin
		THEtab1:= THEdit(GetControl('PSA_ETABLISSEMENT'));
		THEtab2:= THEdit(GetControl('PSA_ETABLISSEMENT_'));
	end;
	if Arg = 'RESP' Then
	begin
		THEtab1:= THEdit(GetControl('PSI_ETABLISSEMENT'));
		THEtab2:= THEdit(GetControl('PSI_ETABLISSEMENT_'));

		//PT24 - Début
		If PGBundleHierarchie Then
		Begin
			If PGDroitMultiForm Then
			Begin
				THEtab1.Visible := False;
				THEtab2.Visible := False;
				SetControlVisible('LBLETAB', False);
				SetControlVisible('LBLETAB1', False);

				SetControlProperty('VSERVICE', 'PLus', ServicesDesDossiers()); //PT25
			End
			Else
			Begin
				SetControlProperty('VSERVICE', 'PLus', ' AND PGS_NOMSERVICE LIKE "'+V_PGI.NoDossier+'%"'); //PT25
			End;
		End;
	end;
	//FIN PT22

	{PT1-2
	If (VH_Paie.PGRESPONSABLES=True) and (VH_Paie.PGMsa=False) and (VH_Paie.PGIntermittents=False) Then GestionUnique := True;
	If (VH_Paie.PGRESPONSABLES=False) and (VH_Paie.PGMsa=True) and (VH_Paie.PGIntermittents=False) Then GestionUnique := True;
	If (VH_Paie.PGRESPONSABLES=False) and (VH_Paie.PGMsa=False) and (VH_Paie.PGIntermittents=True) Then GestionUnique := True;
	}
	If (VH_Paie.PGRESPONSABLES=True) and (VH_Paie.PGMsa=False) and
	(VH_Paie.PGIntermittents=False) and (VH_Paie.PGBTP=False) and
	(VH_Paie.PGTicketRestau = False) Then // PT8
	GestionUnique := True;

	If (VH_Paie.PGRESPONSABLES=False) and (VH_Paie.PGMsa=True) and
	(VH_Paie.PGIntermittents=False) and (VH_Paie.PGBTP=False) and
	(VH_Paie.PGTicketRestau = False) Then // PT8
	GestionUnique := True;

	If (VH_Paie.PGRESPONSABLES=False) and (VH_Paie.PGMsa=False) and
	(VH_Paie.PGIntermittents=True) and (VH_Paie.PGBTP=False) and
	(VH_Paie.PGTicketRestau =  False) Then // PT8
	GestionUnique := True;

	If (VH_Paie.PGRESPONSABLES=False) and (VH_Paie.PGMsa=False) and
	(VH_Paie.PGIntermittents=False) and (VH_Paie.PGBTP=True) and
	(VH_Paie.PGTicketRestau =  False) Then
	GestionUnique := True;

	// d PT8
	If (VH_Paie.PGRESPONSABLES=False) and (VH_Paie.PGMsa=False) and
	(VH_Paie.PGIntermittents=False) and (VH_Paie.PGBTP=True) and
	(VH_Paie.PGTicketRestau =  True) Then
	GestionUnique := True;
	// f PT8
	//FIN PT1-2

	If GestionUnique=True Then
	TFMul(Ecran).BInsert.Visible:= True;
	if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) then   //PT13 //PT24
	begin
		RecupMinMaxTablette('PG','INTERIMAIRES','PSI_INTERIMAIRE',Min,Max);
		SetControlText('PSE_SALARIE',Min);
		SetControlText('PSE_SALARIE_',Max);
	end
	else
	begin
		RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
		SetControlText('PSE_SALARIE',Min);
		SetControlText('PSE_SALARIE_',Max);
		RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
		//DEB PT22 Gérer d'où l'on vient pour afficher les valeurs par défaut établissement
		If (Arg = 'TCK') or (Arg = 'BTP') or (Arg = 'IS') or (Arg = 'MSA') Then
		begin
			SetControlText('PSA_ETABLISSEMENT',Min);
			SetControlText('PSA_ETABLISSEMENT_',Max);
		end;
		if Arg = 'RESP' Then
		begin
			SetControlText('PSI_ETABLISSEMENT',Min);
			SetControlText('PSI_ETABLISSEMENT_',Max);
		end;
		//FIN PT22
	end;
	LesEtabMSA := '';//PT12
	Q_Mul:= THQuery(Ecran.FindComponent('Q')); //PT19

	// DEB PT21
	// Ces zones ne seront visibles que si Arg = 'RESP' et qu'on a sélectionné les salariés non affectés
	if Arg = 'RESP' then
	begin
		ComboService := THEdit(GetControl('VSERVICE')); //PT25
		LabelService := THLabel(GetControl('TVSERVICE'));
		LibelleService := THLabel(GetControl('TLIBSERVICE')); //PT26
		if ComboService <> nil then
		Begin
			ComboService.Visible := False;
			If PGBundleHierarchie Then ComboService.OnChange := ChargeListeSalaries; //PT25
		End;
		if LabelService <> nil then
		LabelService.Visible := False;
		If Affectation <> Nil Then
		Affectation.OnChange := AfficherComboService;
	end;
	// FIN PT21

end;

Procedure TOF_PGMULDEPORTSAL.AffecterClick(Sender : TObject);
var SalPSE,SalPSA,TitreBox,InfosCompl,MsaAct,ug : String;
    BTPDateAnc : TDateTime;
    Q : TQuery;
    i : Integer;
    TobComplSal : Tob;
begin
        TitreBox := 'Affectation des salariés';
        if (Liste.NbSelected=0) and (TFMul(Ecran).BSelectAll.Down=False) then
        begin
                PGIBox('Aucun élément sélectionné',TitreBox);
                Exit;
        end;
        If Affectation.value = 'AFFECTE' Then
        begin
                PGIBox('Les salariés sélectionnés sont déja affectés',TitreBox);
                Exit;
        end;

        // DEB PT21
        if Arg = 'RESP' then
          If (Affectation.value <> 'AFFECTE') and (ComboService <> nil) and (ComboService.Text = '') Then
          begin
            PGIBox('Vous devez renseigner le service auquel vous voulez affecter les salariés',TitreBox);
            Exit;
          end;

        // Récupérer les infos liées au service (cf UTOMDeportSal)
        if Arg = 'RESP' then
        begin
          Q_Service := OpenSQL('SELECT PGS_SECRETAIREABS,PGS_SECRETAIRENDF,PGS_SECRETAIREVAR,PGS_RESPONSABS,' +
            ' PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_RESPSERVICE,PGS_CODESERVICE' +
            ' FROM SERVICES WHERE PGS_CODESERVICE="' + GetControlText('VSERVICE') + '"', True);
        end;
        // FIN PT21

        // DEBUT PT7
        if ((Liste.nbSelected) > 0) AND (not Liste.AllSelected ) then
        begin
                for i := 0 to Liste.NbSelected-1 do
                begin
                        Liste.GotoLeBOOKMARK(i);        //PT6
                        {$IFDEF EAGLCLIENT}
                        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ; // PT9
                        {$ENDIF}
                        if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) And (GetControlText('GRILLE') = 'RESP') then   //PT13 //PT24
                             SalPSA := TFmul(Ecran).Q.FindField('PSI_INTERIMAIRE').asstring
                        else SalPSA := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
                        SalPSE := TFmul(Ecran).Q.FindField('PSE_SALARIE').asstring;
                        
                        If SalPSE <> '' Then
                        begin
                                Q := OpenSQL('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE="'+SalPSE+'"',True);
                                TobComplSal := Tob.Create('DEPORTSAL',Nil,-1);
                                TobComplSal.LoadDetailDB('DEPORTSAL','','',Q,False);
                                Ferme(Q);
                                
                                If GetControlText('GRILLE') = 'BTP' Then
                                begin
                                        TobComplSal.Detail[0].PutValue('PSE_BTP','X');
                                        BTPDateAnc := TrouveDateAncBTP(SalPSE);
                                        TobComplSal.Detail[0].PutValue('PSE_BTPANCIENNETE',BTPDateAnc);
                                end;
                                
                                If GetControlText('GRILLE') = 'RESP' Then
                                begin
                                  TobComplSal.Detail[0].PutValue('PSE_ORGANIGRAMME','X');
                                  if not Q_Service.eof then
                                  begin
                                    TobComplSal.Detail[0].PutValue('PSE_CODESERVICE',Q_Service.FindField('PGS_CODESERVICE').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTABS',Q_Service.FindField('PGS_SECRETAIREABS').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTVAR',Q_Service.FindField('PGS_SECRETAIREVAR').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTNDF',Q_Service.FindField('PGS_SECRETAIRENDF').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSABS',Q_Service.FindField('PGS_RESPONSABS').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPSERVICE',Q_Service.FindField('PGS_RESPSERVICE').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSNDF',Q_Service.FindField('PGS_RESPONSNDF').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSVAR',Q_Service.FindField('PGS_RESPONSVAR').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSFOR',Q_Service.FindField('PGS_RESPONSFOR').AsString);
                                  end
                                  else
                                  begin
                                    TobComplSal.Detail[0].PutValue('PSE_CODESERVICE',GetControlText('VSERVICE'));
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTABS','');
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTVAR','');
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTNDF','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSABS','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPSERVICE','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSNDF','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSVAR','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSFOR','');
                                  end;
                                end;
                                
                                If GetControlText('GRILLE') = 'TCK' Then TobComplSal.Detail[0].PutValue('PSE_TICKETREST','X');  // PT8
                                If GetControlText('GRILLE') = 'MSA' Then
                                begin
                                        TobComplSal.Detail[0].PutValue('PSE_MSA','X');
                                        MsaAct := TrouveActiviteMSA(SalPSE);
                                        ug := TrouveUGMSA(SalPSE);
                                        TobComplSal.Detail[0].PutValue('PSE_MSAACTIVITE',MsaAct);
                                        TobComplSal.Detail[0].PutValue('PSE_MSATYPEACT','ETB');
                                        TobComplSal.Detail[0].PutValue('PSE_MSAUNITEGES',UG);
                                        TobComplSal.Detail[0].PutValue('PSE_MSATYPUNITEG','ETB');
                                        InfosCompl := TrouveInfosComplMSA(SalPSE);
                                        TobComplSal.Detail[0].PutValue('PSE_MSAINFOSCOMPL',''); //PT10
                                        TobComplSal.Detail[0].PutValue('PSE_MSATECHNIQUE',InfosCompl); //PT10
                                end;
                                
                                If GetControlText('GRILLE') = 'IS' Then
                                begin
                                     TobComplSal.Detail[0].PutValue('PSE_INTERMITTENT','X');
                                     Q := OpenSQL('SELECT PSA_NUMEROSS FROM SALARIES WHERE PSA_SALARIE="'+SalPSE+'"',True); //PT18
                                     If Not Q.Eof then TobComplSal.Detail[0].PutValue('PSE_ISNUMIDENT',Q.FindField('PSA_NUMEROSS').AsString);
                                     Ferme(Q);
                                end;
                                
                                TobComplSal.Detail[0].UpdateDB(False);
                                TobComplSal.free;
                        end
                        Else
                        begin
                                CreationEnregistrement(GetControlText('GRILLE'),SalPSA);
                        end;
                end;
        end;

        {$IFDEF EAGLCLIENT}
        if (TFMul(Ecran).bSelectAll.Down) then
        TFMul(Ecran).Fetchlestous;              //PT9
        {$ENDIF}
        If (Liste.AllSelected = TRUE) then
        begin
                Q_Mul.First;
                while Not Q_Mul.EOF do
                begin
                        if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) And (GetControlText('GRILLE') = 'RESP') then   //PT13 //PT24
                              SalPSA := Q_Mul.FindField('PSI_INTERIMAIRE').asstring
                        else SalPSA := Q_Mul.FindField('PSA_SALARIE').asstring;
                        SalPSE := Q_Mul.FindField('PSE_SALARIE').asstring;
                        If SalPSE <> '' Then
                        begin
                                Q := OpenSQL('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE="'+SalPSE+'"',True);
                                TobComplSal := Tob.Create('DEPORTSAL',Nil,-1);
                                TobComplSal.LoadDetailDB('DEPORTSAL','','',Q,False);
                                Ferme(Q);
                                If GetControlText('GRILLE') = 'BTP' Then
                                begin
                                        TobComplSal.Detail[0].PutValue('PSE_BTP','X');
                                        BTPDateAnc := TrouveDateAncBTP(SalPSE);
                                        TobComplSal.Detail[0].PutValue('PSE_BTPANCIENNETE',BTPDateAnc);
                                end;
                                If GetControlText('GRILLE') = 'RESP' Then
                                //DEB PT21
                                begin
                                  TobComplSal.Detail[0].PutValue('PSE_ORGANIGRAMME','X');
                                  if not Q_Service.eof then
                                  begin
                                    TobComplSal.Detail[0].PutValue('PSE_CODESERVICE',Q_Service.FindField('PGS_CODESERVICE').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTABS',Q_Service.FindField('PGS_SECRETAIREABS').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTVAR',Q_Service.FindField('PGS_SECRETAIREVAR').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTNDF',Q_Service.FindField('PGS_SECRETAIRENDF').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSABS',Q_Service.FindField('PGS_RESPONSABS').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPSERVICE',Q_Service.FindField('PGS_RESPSERVICE').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSNDF',Q_Service.FindField('PGS_RESPONSNDF').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSVAR',Q_Service.FindField('PGS_RESPONSVAR').AsString);
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSFOR',Q_Service.FindField('PGS_RESPONSFOR').AsString);
                                  end
                                  else
                                  begin
                                    TobComplSal.Detail[0].PutValue('PSE_CODESERVICE',GetControlText('VSERVICE'));
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTABS','');
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTVAR','');
                                    TobComplSal.Detail[0].PutValue('PSE_ASSISTNDF','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSABS','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPSERVICE','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSNDF','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSVAR','');
                                    TobComplSal.Detail[0].PutValue('PSE_RESPONSFOR','');
                                  end;
                                end;
                                //FIN PT21

                                // d PT8
                                If GetControlText('GRILLE') = 'TCK' Then
                                {Tickets resstaurant}
                                begin
                                 TobComplSal.Detail[0].PutValue('PSE_TICKETREST','X');
                                 TobComplSal.Detail[0].PutValue('PSE_TYPETYPTK','ETB');
                                end;
// f PT8
                                If GetControlText('GRILLE') = 'MSA' Then
                                begin
                                        TobComplSal.Detail[0].PutValue('PSE_MSA','X');
                                        MsaAct := TrouveActiviteMSA(SalPSE);
                                        ug := TrouveUGMSA(SalPSE);
                                        TobComplSal.Detail[0].PutValue('PSE_MSAACTIVITE',MsaAct);
                                        TobComplSal.Detail[0].PutValue('PSE_MSATYPEACT','ETB');
                                        TobComplSal.Detail[0].PutValue('PSE_MSAUNITEGES',UG);
                                        TobComplSal.Detail[0].PutValue('PSE_MSATYPUNITEG','ETB');
                                        InfosCompl := TrouveInfosComplMSA(SalPSE);
                                        TobComplSal.Detail[0].PutValue('PSE_MSAINFOSCOMPL','');//PT10
                                        TobComplSal.Detail[0].PutValue('PSE_MSATECHNIQUE',InfosCompl);//PT10
                                end;
                                If GetControlText('GRILLE') = 'IS' Then
                                begin
                                     TobComplSal.Detail[0].PutValue('PSE_INTERMITTENT','X');
                                      Q := OpenSQL('SELECT PSA_NUMEROSS FROM SALARIES WHERE PSA_SALARIE="'+SalPSE+'"',True);   //PT18
                                     If Not Q.Eof then TobComplSal.Detail[0].PutValue('PSE_ISNUMIDENT',Q.FindField('PSA_NUMEROSS').AsString);
                                     Ferme(Q);
                                end;
                                TobComplSal.Detail[0].UpdateDB(False);
                                TobComplSal.free;
                        end
                        Else
                        begin
                                CreationEnregistrement(GetControlText('GRILLE'),SalPSA);
                        end;
                        TFmul(Ecran).Q.Next;
                end;
        end;
        //FIN PT7

        // DEB PT21
        if Arg = 'RESP' then
          Ferme(Q_Service);
        // FIN PT21

        TFMul(Ecran).BChercheClick(Nil);
        LesEtabMSA := '';
end;

procedure TOF_PGMULDEPORTSAL.ExitEdit(Sender : TObject);
var edit : thedit;
begin
        edit := THEdit(Sender);
        if edit  <>  nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit,10);
end;

//PT11
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 19/08/2004
Modifié le ... :   /  /
Description .. : Click sur la check-Box "Exclure les salariés sortis"
Mots clefs ... : PAIE;COMPLSALARIE
*****************************************************************}
procedure TOF_PGMULDEPORTSAL.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;
//FIN PT11


Procedure TOF_PGMULDEPORTSAL.ChangeChoix(Sender : TObject);          //15 MARS
begin
        If CChoix.Value = 'RES' Then
        begin
                Arg := 'RESP';
                SetControlText('GRILLE','RESP');
        end;
        If CChoix.Value = 'INT' Then
        begin
                Arg := 'IS';
                SetControlText('GRILLE','IS');
        end;
        If CChoix.Value = 'MSA' Then
        begin
                Arg := 'MSA';
                SetControlText('GRILLE','MSA');
        end;
        If CChoix.Value = 'BTP' Then
        begin
                Arg := 'BTP';
                SetControlText('GRILLE','BTP');
        end;
// d PT8
        If CChoix.Value = 'TCK' Then
        begin
                Arg := 'TCK';
                SetControlText('GRILLE','TCK');
        end;
// f PT8
end;

Procedure TOF_PGMULDEPORTSAL.ExitChoix(Sender : TObject);         //15 MARS
begin
        If THValComboBox(Sender).Value = 'BTP' Then
        If VerifChoix('BTP')=True Then
        begin
                PGIBox('La gestion BTP n''est pas effectuée pour cette société','Saisie complémentaire salarié');
                SetFocusControl('CHOIX');
        end;
// d PT8
        If THValComboBox(Sender).Value = 'TCK' Then
        If VerifChoix('TCK')=True Then
        begin
                PGIBox('La gestion des tickets restaurant n''est pas effectuée pour cette société','Saisie complémentaire salarié');
                SetFocusControl('CHOIX');
        end;
// f PT8
        If THValComboBox(Sender).Value = 'RES' Then
        begin
                If VerifChoix('RES')=True Then
                begin
                        PGIBox('La gestion des responsables n''est pas effectuée pour cette société','Saisie complémentaire salarié');
                        SetFocusControl('CHOIX');
                end;
        end;
        If THValComboBox(Sender).Value = 'MSA' Then
        begin
                If VerifChoix('MSA')=True Then
                begin
                        PGIBox('La gestion de la MSA n''est pas effectuée pour cette société','Saisie complémentaire salarié');
                        SetFocusControl('CHOIX');
                end;
        end;
        If THValComboBox(Sender).Value = 'INT' Then
        begin
                if VerifChoix('INT')=True Then
                begin
                        PGIBox('La gestion intermittents du spectacle n''est pas effectuée pour cette société','Saisie complémentaire salarié');
                        SetFocusControl('CHOIX');
                end;
        end;
end;

Function TOF_PGMULDEPORTSAL.VerifChoix(Choix : String):Boolean;            //15 MARS
begin
        Result := False;
        If Choix = 'RES' Then
        begin
                If VH_Paie.PGRESPONSABLES=False Then result := true;
                exit;
        end;
        If Choix = 'BTP' Then
        begin
                If VH_Paie.PGBTP=False Then result := True;
                exit;
        end;
// d PT8
        If Choix = 'TCK' Then
        begin
                If VH_Paie.PGTicketRestau=False Then result := True;
                exit;
        end;
// f PT8
        If Choix = 'INT' Then
        begin
                if VH_Paie.PGIntermittents=False Then result := True;
                exit;
        end;
        If Choix = 'MSA' Then
        begin
                If VH_Paie.PGMsa=False Then result := True;
                exit;
        end;
end;

Procedure TOF_PGMULDEPORTSAL.RespElipsisClick(Sender : TObject);
var St:String;
begin
         if ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) then //PT24
         begin
              St := 'PSI_INTERIMAIRE IN (SELECT PSE_RESPONSABS FROM DEPORTSAL)';
			//PT25 - Début
              If Not PGDroitMultiForm Then
              	St := St + DossiersAInterroger('DOS',V_PGI.NoDossier,'PSI',False,True)
              Else
                St := St + DossiersAInterroger('DOS','','PSI',False,True);
              LookupList(THEdit(Sender),'Liste des responsables des absences','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',St,'PSI_INTERIMAIRE', True,-1);
			//PT25 - Fin
         end
         else
         begin
              St := 'PSA_SALARIE IN (SELECT PSE_RESPONSABS FROM DEPORTSAL)';
              LookupList(THEdit(Sender),'Liste des responsables des absences','SALARIES','PSA_SALARIE','PSA_LIBELLE,PSI_PRENOM',St,'PSA_SALARIE', True,-1);  //PT25
         end;
end;

// DEBUT PT7

procedure TOF_PGMULDEPORTSAL.CreationEnregistrement(TypeEnreg,Salarie : String);
var T : Tob;
    InfosCompl,MsaAct,ug : String;
    BTPDateAnc : TDateTime;
    Etab    : string; // PT8
    Q : TQuery; // PT8
begin
        T := Tob.Create('DEPORTSAL',Nil,-1);
        T.PutValue('PSE_SALARIE',Salarie);
        // Compléments salarié
        If TypeEnreg = 'RESP' then T.PutValue('PSE_ORGANIGRAMME','X')
        else T.PutValue('PSE_ORGANIGRAMME','-');
        T.PutValue('PSE_RESPONSABS','');
        T.PutValue('PSE_RESPONSNDF','');
        T.PutValue('PSE_RESPONSVAR','');
        T.PutValue('PSE_ASSISTABS','');
        T.PutValue('PSE_ASSISTNDF','');
        T.PutValue('PSE_ASSISTVAR','');
        T.PutValue('PSE_OKVALIDABS','-');
        T.PutValue('PSE_OKVALIDNDF','-');
        T.PutValue('PSE_OKVALIDVAR','-');
        T.PutValue('PSE_VEHICULESOC','-');
        T.PutValue('PSE_ADMINISTABS','-');
        T.PutValue('PSE_CODESERVICE','');
        T.PutValue('PSE_RESPSERVICE','');
        T.PutValue('PSE_RESPONSFOR','');
        T.PutValue('PSE_VISADRH','');
        T.PutValue('PSE_EMAILPROF','');
        T.PutValue('PSE_EMAILENVOYE','-');
        T.PutValue('PSE_EMAILDATE',IDate1900);
        // Intermittent
        If TypeEnreg = 'IS' then T.PutValue('PSE_INTERMITTENT','X')
        Else T.PutValue('PSE_INTERMITTENT','-');
        T.PutValue('PSE_ISNUMIDENT','');
        T.PutValue('PSE_ISCONGSPEC','');
        T.PutValue('PSE_ISRETRAITE','');
        T.PutValue('PSE_ISCATEG','');
        T.PutValue('PSE_ISNUMASSEDIC','');
        If TypeEnreg = 'IS' then //PT18
        begin
           Q := OpenSQL('SELECT PSA_NUMEROSS FROM SALARIES WHERE PSA_SALARIE="'+Salarie+'"',True);   //PT18
           If Not Q.Eof then t.PutValue('PSE_ISNUMIDENT',Q.FindField('PSA_NUMEROSS').AsString);
           Ferme(Q);
        end;
        // MSA
        If TypeEnreg = 'MSA' then
        begin
                T.PutValue('PSE_MSA','X') ;
                MsaAct := TrouveActiviteMSA(Salarie);
                ug := TrouveUGMSA(Salarie);
                T.PutValue('PSE_MSAACTIVITE',MsaAct);
                T.PutValue('PSE_MSATYPEACT','ETB');
                T.PutValue('PSE_MSAUNITEGES',UG);
                T.PutValue('PSE_MSATYPUNITEG','ETB');
                InfosCompl := TrouveInfosComplMSA(Salarie);
                T.PutValue('PSE_MSAINFOSCOMPL','');//PT10
                T.PutValue('PSE_MSATECHNIQUE',InfosCompl);//PT10
        end
        else
        begin
                T.PutValue('PSE_MSA','-');
                T.PutValue('PSE_MSAACTIVITE','');
                T.PutValue('PSE_MSATYPEACT','');
                T.PutValue('PSE_MSAINFOSCOMPL','');
        end;
        T.PutValue('PSE_MSALIEUTRAV','');
        // BTP
        If TypeEnreg = 'BTP' then
        begin
                BTPDateAnc := TrouveDateAncBTP(Salarie);
                T.PutValue('PSE_BTP','X');
                T.PutValue('PSE_BTPANCIENNETE',BTPDateAnc);
        end
        else
        begin
                T.PutValue('PSE_BTP','-');
                T.PutValue('PSE_BTPANCIENNETE',IDate1900);
        end;
        T.PutValue('PSE_BTPADHESION','');
        T.PutValue('PSE_BTPASSEDIC','');
        T.PutValue('PSE_BTPSALMOYEN','');
        T.PutValue('PSE_BTPHORAIRE','');
// d PT8
        {Tickets restaurant}
        T.PutValue('PSE_TICKETREST','-');
        T.PutValue('PSE_TYPETYPTK','');
        T.PutValue('PSE_TYPTICKET','');
        T.PutValue('PSE_DISTRIBUTION','');
        T.PutValue('PSE_PERSONNAL','');
        T.PutValue('PSE_INFOCOMPL','');

        If (TypeEnreg = 'TCK') then
        begin
          T.PutValue('PSE_TYPETYPTK','ETB');
          T.PutValue('PSE_TICKETREST','X');
          {Type de ticket restaurant : gestion idem établissement}
          Q:=OpenSQL('SELECT ETB_ETABLISSEMENT,PSA_SALARIE FROM ETABCOMPL '+
                     'LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT '+
                     'WHERE PSA_SALARIE="'+Salarie+'"',True);
          if not Q.eof then
            Etab := Q.FindField('ETB_ETABLISSEMENT').AsString;
          Ferme(Q);
          Q:=OpenSQL('SELECT ETB_TYPTICKET FROM ETABCOMPL '+
                     'WHERE ETB_ETABLISSEMENT = "'+Etab+'"',True);
          if not Q.eof then
            T.PutValue('PSE_TYPTICKET',Q.FindField('ETB_TYPTICKET').AsString);
          Ferme(Q);

          if GetParamSoc('SO_PGPERSOTICKET') then
          {Utilisation de la zone de personnalisation, par défaut NOm Prénom}
          begin
            Q:=OpenSQL('SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES '+
                       'WHERE PSA_SALARIE = "'+Salarie+'"',True);
            if not Q.eof then
              T.PutValue('PSE_PERSONNAL',
                       Trim(Q.FindField('PSA_LIBELLE').AsString)+' '+
                       Trim(Q.FindField('PSA_PRENOM').AsString));
            Ferme(Q);;
          end
          else
          begin
            T.PutValue('PSE_PERSONNAL','');
          end;
        end;

// f PT8
        // DEB PT21
        If (TypeEnreg = 'RESP') and (not Q_Service.eof) then
        begin
          T.PutValue('PSE_CODESERVICE',Q_Service.FindField('PGS_CODESERVICE').AsString);
          T.PutValue('PSE_ASSISTABS',Q_Service.FindField('PGS_SECRETAIREABS').AsString);
          T.PutValue('PSE_ASSISTVAR',Q_Service.FindField('PGS_SECRETAIREVAR').AsString);
          T.PutValue('PSE_ASSISTNDF',Q_Service.FindField('PGS_SECRETAIRENDF').AsString);
          T.PutValue('PSE_RESPONSABS',Q_Service.FindField('PGS_RESPONSABS').AsString);
          T.PutValue('PSE_RESPSERVICE',Q_Service.FindField('PGS_RESPSERVICE').AsString);
          T.PutValue('PSE_RESPONSNDF',Q_Service.FindField('PGS_RESPONSNDF').AsString);
          T.PutValue('PSE_RESPONSVAR',Q_Service.FindField('PGS_RESPONSVAR').AsString);
          T.PutValue('PSE_RESPONSFOR',Q_Service.FindField('PGS_RESPONSFOR').AsString);
        end;
        // FIN PT21

        T.InsertDB(Nil,False);
        T.Free;
end;

Function TOF_PGMULDEPORTSAL.TrouveInfosComplMSA(SalarieMSA : String) : String;
var Q : TQuery;
    InfosCompl,codeBur : String;
begin
        Q := OpenSQL('SELECT PAT_CODEBUREAU FROM TAUXAT LEFT JOIN SALARIES ON PAT_ORDREAT=PSA_ORDREAT AND'+
        ' PAT_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PSA_SALARIE="'+SalarieMSA+'"',True);
        CodeBur := '';
        If not Q.eof then CodeBur := Q.FindField('PAT_CODEBUREAU').AsString;  // PortageCWAS
        If CodeBur = 'BUR' Then InfosCompl := '-'    //PT10 champ booléen au lieu de VARCHAR
        else InfosCompl := 'X';
        Ferme(Q);
        Result := InfosCompl;
end;

Function TOF_PGMULDEPORTSAL.TrouveActiviteMSA(SalarieMSA : String) : String;
var Q : TQuery;
    UnEtab,Etab,MsaAct,EtabSal : String;
    AfficherMess : Boolean;
begin
        MsaAct := '';
        EtabSal := '';
        Q := OpenSQL('SELECT ETB_MSAACTIVITE,PSA_ETABLISSEMENT FROM ETABCOMPL LEFT JOIN SALARIES '+
        'ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE PSA_SALARIE="'+SalarieMSA+'"',True);
        If Not Q.Eof then
        begin
             MsaAct := Q.findField('ETB_MSAACTIVITE').AsString;
             EtabSal := Q.findField('PSA_ETABLISSEMENT').AsString;
        end;
        Ferme(Q);
        //DEBUT PT12
        If (MsaAct = '') and (EtabSal <> '') then
        begin
             Etab := LesEtabMSA;
             AfficherMess := True;
             unEtab := ReadTokenPipe(Etab,';');
             While UnEtab <> '' do
             begin
                  If EtabSal = UnEtab then AfficherMess := False;
                  unEtab := ReadTokenPipe(Etab,';');
             end;
             If AfficherMess then
             begin
                  If LesEtabMSA <> '' then LesEtabMSA := LesEtabMSA +';'+EtabSal
                  else LesEtabMSA := EtabSal;
                  PGIBox('Vous devez renseigner l''activité de l''établissement '+etabSal + ' '+RechDom('TTETABLISSEMENT',EtabSal,False));
             end;
        end;
        //FIN PT12
        Result := MsaAct;
end;

Function TOF_PGMULDEPORTSAL.TrouveDateAncBTP(SalarieBTP : String) : TDateTime;
var Q : TQuery;
    BTPDateAnc : TDateTime;
begin
        BTPDateAnc := IDate1900;
        Q := OpenSQL('SELECT PSA_DATEANCIENNETE FROM SALARIES WHERE PSA_SALARIE="'+SalarieBTP+'"',True);
        if not Q.eof then BTPDateAnc := Q.FindField('PSA_DATEANCIENNETE').AsDateTime;
        Ferme(Q);
        Result := BTPDateAnc;
end;

// FIN PT7

Function TOF_PGMULDEPORTSAL.TrouveUGMSA(SalarieMSA : String) : String;
var Q : TQuery;
    ug : String;
begin
        Ug := '';
        Q := OpenSQL('SELECT ETB_MSAUNITEGES,PSA_ETABLISSEMENT FROM ETABCOMPL LEFT JOIN SALARIES '+
        'ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE PSA_SALARIE="'+SalarieMSA+'"',True);
        If Not Q.Eof then
        begin
             ug := Q.findField('ETB_MSAUNITEGES').AsString;
        end;
        Ferme(Q);
        Result := ug;
end;

//DEB PT21
procedure TOF_PGMULDEPORTSAL.AfficherComboService(Sender: TObject);
begin
  if Affectation.value <> 'AFFECTE' then
  begin
    ComboService.Visible := True;
    LabelService.Visible := True;
    If LibelleService <> Nil Then LibelleService.Visible := True; //PT26
  end
  else
  begin
    ComboService.Visible := False;
    LabelService.Visible := False;
    If LibelleService <> Nil Then LibelleService.Visible := False; //PT26
  end;
end;
//FIN PT21


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/04/2008 / PT25
Modifié le ... :   /  /
Description .. : Rafraîchit la liste des salariés affectables suite à la sélection d'un service
Mots clefs ... :
*****************************************************************}
Procedure TOF_PGMULDEPORTSAL.ChargeListeSalaries (Sender : TObject);
var Bt : TToolBarButton97;
Begin
	Bt := TToolBarButton97 (GetControl('BCHERCHE'));
	If Bt <> Nil Then BT.Click;
End;


Initialization
registerclasses([TOF_PGMULDEPORTSAL]);
end.

