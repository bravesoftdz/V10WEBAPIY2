{***********UNITE*************************************************
Auteur  ......  :
Créé le ...... : 18/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : SAISIEINSCBUDGET ()
Mots clefs ... : TOF;SAISIEINSCBUDGET
*****************************************************************
PT1  | 24/04/2003 | V_42  | JL | Développement pour compatibilité CWAS
PT2  | 26/01/2007 | V_75  | JL | Supp du tree view
PT3  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    | pour les critères code salarié uniquement
PT4  | 28/09/2007 | V_7   | FL | Emanager / Report / Adaptation cursus + accès assistant
PT5  | 19/10/2007 | V_7   | FL | Emanager / Report / Ajout saisie en masse des salariés
PT6  | 04/12/2007 | V_8   | FL | Emanager / Report / Gestion des sous-niveaux
PT7  | 10/01/2008 | V_802 | FL | FQ 13337 Correction du calcul des frais de repas
PT8  | 02/04/2008 | V_803 | FL | Adaptation pour partage formation
PT9  | 23/04/2008 | V_803 | FL | Inscriptions multiples
PT10 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT11 | 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}
Unit UTofPGSaisieInscBudget;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOB,UTOF,uTableFiltre,SaisieList,
     hTB97,PGOutilsFormation,EntPaie,LookUp,PGOutils;


Type
  TOF_SAISIEINSCBUDGET = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    Private
    TypeSaisie,TypeInsc,MillesimeEC,LeStage,TypeUtilisat,LeMillesime,Utilisateur : String;  //PT5
    TF :  TTableFiltre;
    procedure ChangeFiltreCursus(Sender : TObject);
    procedure SalarieElipsisClick (Sender : TObject);
    //procedure ChangeFiltreHierarchie (Sender : TObject); //PT4
    procedure SaisieEnMasse (Sender : TObject);  //PT5
  end ;

Implementation

Uses UtilPGI,ParamSoc;

procedure TOF_SAISIEINSCBUDGET.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_SAISIEINSCBUDGET.OnLoad ;
{$IFDEF EMANAGER}
var  Q : TQuery; //PT4
{$ENDIF}
begin
  Inherited ;
       //PT4 - Début
       {$IFDEF EMANAGER}
       Q := OpenSQL('SELECT PST_BLOCNOTE FROM STAGE WHERE PST_CODESTAGE="'+LeStage+'" AND PST_MILLESIME="0000"', True);
       If Not Q.EOF Then SetControlText('PST_BLOCNOTE', Q.FindField('PST_BLOCNOTE').AsString);
       Ferme(Q);
       {$ENDIF}
       //PT4 - Fin
end ;

procedure TOF_SAISIEINSCBUDGET.OnArgument (S : String ) ;
var 
    Millesime,SQL : String;
    Num : Integer;
    Edit : THEdit;
    Bt : TToolBarButton97;
    Combo : THValComboBox;
    Action : String;
begin
  Inherited ;
        Utilisateur := V_PGI.UserSalarie;
        TypeSaisie := ReadTokenSt(S);
        TypeInsc := ReadTokenSt(S);
        LeStage := ReadTokenSt(S);
        Millesime := ReadTokenSt(S);
        MillesimeEC := RendMillesimePrevisionnel;
        If Millesime = '0000' then LeMillesime := MillesimeEC
        else LeMillesime := Millesime;
        {$IFNDEF EMANAGER}
        READTOKENPipe(S,'=');
        Action := S;
        {$ELSE}
        ReadTokenSt(S);
        ReadTokenSt(S);
        ReadTokenSt(S);
        Utilisateur := ReadTokenSt(S);
        //PT4 - Début (déplacé)
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+Utilisateur+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        SetControlProperty('PFI_LIBEMPLOIFOR','Plus','AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
                        'WHERE '+AdaptByRespEmanager(TypeUtilisat,'PSE',Utilisateur,False)+')'); //PT2
        MillesimeEC := RendMillesimeEManager;
        //PT4 - Fin
        {$ENDIF}
        If (TypeSaisie = 'SAISIESTAGE') or (TypeSaisie = 'CWASINSCBUDGET')  Then
        begin
                SetControlVisible('BTree',False);
                SetControlVisible('PANTREEVIEW',False);
        end;
        TF  :=  TFSaisieList(Ecran).LeFiltre;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('FORMATION'+IntToStr(Num)),GetControl ('TFORMATION'+IntToStr(Num)));
        end;
        Edit := THEdit(GetControl('PFI_SALARIE'));
        If Edit <> Nil Then Edit.OnelipsisClick := SalarieElipsisClick;
        If (TypeSaisie = 'SAISIEINSC') and (Lestage = '' )Then
        begin
                SetControlVisible('BUDGET',True);
                SetControlVisible('LBUDGET',True);
        end;

        SQL := 'WHERE PFI_CODESTAGE="'+LeStage+'" AND PFI_MILLESIME="'+LeMillesime+'"';
        If (TypeSaisie = 'CWASINSCBUDGET') or (TypeSaisie = 'CONSULTATION') Then
        begin
                If TypeUtilisat = 'R' then SQL := SQL+' AND (PFI_RESPONSFOR="'+Utilisateur+'"'+
                 ' OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))'
                else SQL := SQL + ' AND PFI_MILLESIME=:PST_MILLESIME AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+Utilisateur+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'")))';
        end;
        If (TypeSaisie = 'SAISIEVALID') Or  (TypeSaisie = 'SAISIEINSC') or (TypeSaisie = 'SAISIESTAGE') Then
        begin
                {$IFDEF EMANAGER}
                If TypeUtilisat = 'R' then  SQL := SQL + ' AND (PFI_RESPONSFOR="'+Utilisateur+'"'+
                 ' OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))'
                else SQL := SQL +  ' AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+Utilisateur+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'")))';
                {$ENDIF}
        end;

        //PT10
        // N'affiche pas les salariés confidentiels
        If PGBundleHierarchie Then
        	SQL := SQL + ' AND (PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"))' //PT11
        Else
        	SQL := SQL + ' AND (PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"))'; //PT11
        
        TF.WhereTable := SQL;
       {$IFNDEF EMANAGER}SetControlVisible('BPARAMLISTE',True);{$ENDIF} //PT4
        SetActiveTabSheet('PFORMATION');
        SetControlText('FILTRECURSUS','AUCUN');
       Combo := THValComboBox(GetControl('FILTRECURSUS'));
       If Combo <> Nil then Combo.OnChange := ChangeFiltreCursus;
       SetControlProperty('PFI_NBINSC','MinValue',1);
       SetControlProperty('PFI_NBINSC','MaxValue',10000);
       SetControlEnabled('PFI_DUREESTAGE',False);
       SetControlEnabled('PFI_JOURSTAGE',False);
       {$IFDEF EMANAGER}
       TFSaisieList(Ecran).ParamTreeView := False;
       //PT4
       //SetControltext('CHIERARCHIE','TOUS');
       //SetControlVisible('CHIERARCHIE',True);
       //SetControlVisible('TCHIERARCHIE',True);
       SetControlVisible('FILTRECURSUS',False);
       SetControlVisible('TFILTRECURSUS',False);
       //Combo := THValComboBox(GetControl('CHIERARCHIE'));
       //If Combo <> Nil then Combo.OnChange := ChangeFiltreHierarchie;
       {$ENDIF}

       //PT8
       //{$IFNDEF EMANAGER}
       //If PGBundleInscFormation Then
       //Begin
       //{$ENDIF}
       If (Action = 'CONSULTATION')Then SetControlVisible('BMULTISAL', False);
       Bt := TToolbarButton97(GetControl('BMULTISAL'));
       If Bt <> Nil then Bt.onClick := SaisieEnMasse;  //PT5
       //{$IFNDEF EMANAGER}
       //End;
       //{$ENDIF}
end ;

procedure TOF_SAISIEINSCBUDGET.ChangeFiltreCursus(Sender : TObject);
var SQL : String;
begin
        if (TF.State = DsEdit) or (TF.State = DsInsert) then
        begin
                PGIBox('Changement impossible car vous êtes en modification', Ecran.Caption);
                Exit;
        end;
        SQL := 'WHERE PFI_CODESTAGE="'+LeStage+'" AND PFI_MILLESIME="'+LeMillesime+'"';
        {$IFDEF EMANAGER}
        If TypeUtilisat = 'R' then SQL := SQL + ' AND PFI_RESPONSFOR="'+Utilisateur+'"'
        else SQL := 'WHERE PFI_CODESTAGE=:PST_CODESTAGE AND PFI_MILLESIME=:PST_MILLESIME AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+Utilisateur+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'")))';
        {$ENDIF}
        If GetControlText('FILTRECURSUS') = 'SANS' then SQL := SQL + ' AND PFI_CURSUS=""';
        If GetControlText('FILTRECURSUS') = 'AVEC' then SQL := SQL + ' AND PFI_CURSUS<>""';
        
        //PT10
        // N'affiche pas les salariés confidentiels
        If PGBundleHierarchie Then
        	SQL := SQL + ' AND (PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"))' //PT11
        Else
        	SQL := SQL + ' AND (PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"))'; //PT11
        
        TF.WhereTable:= SQL;
        TF.RefreshLignes;
end;

procedure TOF_SAISIEINSCBUDGET.SalarieElipsisClick(Sender : TObject);
var
  StFrom, StWhere: string;
  Q: TQuery;
  DD: TDateTime;
begin
  Q := OpenSQL('SELECT PFE_DATEDEBUT FROM EXERFORMATION WHERE PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-" ORDER BY PFE_MILLESIME DESC', True);
  if not Q.Eof then DD := Q.FindField('PFE_DATEDEBUT').AsDateTime
  else DD := IDate1900;
  Ferme(Q);
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE>="' + UsDateTime(DD) + '")';
  StFrom := 'SALARIES';
  {$IFDEF EMANAGER}
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  //PT6 - Début
//  If TypeUtilisat = 'R' then StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+Utilisateur+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
//  ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))'
//  else StWhere := StWhere + ' AND PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="' + Utilisateur + '")';
  StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',Utilisateur,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
  //PT6 - Fin
  {$ENDIF}
  StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT3
  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

//PT4
(*
procedure TOF_SAISIEINSCBUDGET.ChangeFiltreHierarchie (Sender : TObject);
var SQL : String;
begin
        if TF.State = DsEdit then
        begin
                PGIBox('Changement impossible car vous êtes en modification', Ecran.Caption);
                Exit;
        end;
        if TF.State = DsInsert then
        begin
                PGIBox('Changement impossible car vous êtes en création', Ecran.Caption);
                Exit;
        end;
        SQL := 'WHERE PFI_CODESTAGE="'+LeStage+'" AND PFI_MILLESIME="'+LeMillesime+'"';
        {$IFDEF EMANAGER}
        If TypeUtilisat = 'R' then
        begin
             If GetControltext('CHIERARCHIE') = 'DIRECT' then SQL := SQL + ' AND PFI_RESPONSFOR="'+Utilisateur+'"'
             else If GetControltext('CHIERARCHIE') = 'INDIRECT' then SQL := SQL + ' AND '+
                 ' PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'"))'
             else SQL := SQL + ' AND (PFI_RESPONSFOR="'+Utilisateur+'"'+
                 ' OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+Utilisateur+'")))';
        end
        else SQL := SQL + ' AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+Utilisateur+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+Utilisateur+'")))';
        {$ENDIF}
        TF.WhereTable:= SQL;
        TF.RefreshLignes;
end;
*)

procedure TOF_SAISIEINSCBUDGET.SaisieEnMasse (Sender : TObject);
var Ret : String;
    Salarie : String;
    t,TobStage,TobInvest,TI : Tob;
    Salaire : Double;
    Q,QAnim : TQuery;
    For1,For2,For3,For4,For5,For6,For7,For8 : String;
    Nature,Lieu : String;
    JourStage,Duree,FraisH,FraisR,FraisT,TotalFrais : Double;
    TauxChargeBudget,NbJours : Double;
    SalaireAnim,NbAnim : Double;
    {$IFDEF EMANAGER}
    Bt : TToolBarButton97;
    {$ENDIF}
    Pfx : String;
begin
	//PT8
	{$IFDEF EMANAGER}
	Ret := AGLLanceFiche('PAY','EM_MULSALARIE','','','SAISIEMASSE');
	{$ELSE}
	If PGBundleHierarchie Or GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) Then //PT9
		Pfx := 'PSI'
    Else
    	Pfx := 'PSA';

   	Ret := AGLLanceFiche('PAY','INTERIMINSC_MUL','','','PREVISIONNEL;PFI_CODESTAGE="'+LeStage+'" AND PFI_MILLESIME="'+LeMillesime+'"');
	{$ENDIF}
	If Ret = '' then exit;

	Duree := 0;
	JourStage := 0;
	NbAnim := 0;

	TF.DisableTom;
	TF.StartUpdate;
	Q := OpenSQL('SELECT PST_JOURSTAGE,PST_LIEUFORM,PST_DUREESTAGE,PST_NBANIM,' +
	'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8,PST_NATUREFORM' +
	' FROM STAGE WHERE PST_CODESTAGE="' + LeStage + '" AND PST_MILLESIME="0000"', True);
	if not Q.eof then
	begin
		JourStage := Q.FindField('PST_JOURSTAGE').AsFloat;
		Duree := Q.FindField('PST_DUREESTAGE').AsFloat;
		Lieu := Q.FindField('PST_LIEUFORM').AsString;
		For1 := Q.FindField('PST_FORMATION1').AsString;
		For2 := Q.FindField('PST_FORMATION2').AsString;
		For3 := Q.FindField('PST_FORMATION3').AsString;
		For4 := Q.FindField('PST_FORMATION4').AsString;
		For5 := Q.FindField('PST_FORMATION5').AsString;
		For6 := Q.FindField('PST_FORMATION6').AsString;
		For7 := Q.FindField('PST_FORMATION7').AsString;
		For8 := Q.FindField('PST_FORMATION8').AsString;
		Nature := Q.FindField('PST_NATUREFORM').AsString;
		NbAnim := Q.FindField('PST_NBANIM').AsFloat;
	end;
	Ferme(Q);

	Q := OpenSQL('SELECT PFE_TAUXBUDGET FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
	if not Q.Eof then TauxChargeBudget := Q.FindField('PFE_TAUXBUDGET').AsFloat
	else TauxChargeBudget := 1;
	Ferme(Q);

	Salarie := ReadTokenPipe(Ret,';');
	While Salarie <> '' do
	begin
		TF.Insert;

		TF.PutValue('PFI_SALARIE',Salarie);
		If PGBundleHierarchie Then
			TF.PutValue('PFI_LIBELLE',Rechdom('PGSALARIEINT',Salarie,False))
		Else
			TF.PutValue('PFI_LIBELLE',Rechdom('PGSALARIE',Salarie,False));
		TF.PutValue('PFI_CODESTAGE',LeStage);
		TF.PutValue('PFI_MILLESIME',MillesimeEC);
		TF.PutValue('PFI_NBINSC',1);

		//PT8
		If PGBundleHierarchie Or GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) Then //PT9
		Begin
			Q := OpenSQL('SELECT INTERIMAIRES.*,PSE_RESPONSFOR ' +
			'FROM INTERIMAIRES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSI_INTERIMAIRE WHERE PSI_INTERIMAIRE="'+Salarie+'"',True);
			Pfx := 'PSI';
		End
		Else
		Begin
			Q := OpenSQL('SELECT PSA_DADSCAT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,' +
			'PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSE_RESPONSFOR ' +
			'FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE PSA_SALARIE="'+Salarie+'"',True);
			Pfx := 'PSA';
		End;
		If Not Q.Eof then
		begin
			TF.PutValue('PFI_ETABLISSEMENT',Q.FindField(Pfx+'_ETABLISSEMENT').AsString);
			TF.PutValue('PFI_LIBEMPLOIFOR',	Q.FindField(Pfx+'_LIBELLEEMPLOI').AsString);
			TF.PutValue('PFI_TRAVAILN1',	Q.FindField(Pfx+'_TRAVAILN1').AsString);
			TF.PutValue('PFI_TRAVAILN2',	Q.FindField(Pfx+'_TRAVAILN2').AsString);
			TF.PutValue('PFI_TRAVAILN3',	Q.FindField(Pfx+'_TRAVAILN3').AsString);
			TF.PutValue('PFI_TRAVAILN4',	Q.FindField(Pfx+'_TRAVAILN4').AsString);
			TF.PutValue('PFI_CODESTAT',		Q.FindField(Pfx+'_CODESTAT').AsString);
			TF.PutValue('PFI_LIBREPCMB1', 	Q.FindField(Pfx+'_LIBREPCMB1').AsString);
			TF.PutValue('PFI_LIBREPCMB2', 	Q.FindField(Pfx+'_LIBREPCMB2').AsString);
			TF.PutValue('PFI_LIBREPCMB3', 	Q.FindField(Pfx+'_LIBREPCMB3').AsString);
			TF.PutValue('PFI_LIBREPCMB4', 	Q.FindField(Pfx+'_LIBREPCMB4').AsString);
			//PT8
			If (Not PGBundleHierarchie) And (Not GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE)) Then  //PT9
				TF.PutValue('PFI_DADSCAT', Q.FindField(Pfx+'_DADSCAT').AsString);
			//PT8
			{$IFDEF EMANAGER}
			TF.PutValue('PFI_RESPONSFOR',Utilisateur);
			{$ELSE}
			TF.PutValue('PFI_RESPONSFOR',Q.FindField('PSE_RESPONSFOR').AsString);
			{$ENDIF}
		end;
		Ferme(Q);

		//PT9
		If PGBundleHierarchie Or GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) Then
		Begin
			Try
				Q := OpenSQL('SELECT PSA_DADSCAT FROM '+GetBase(GetBaseSalarie(Salarie),'SALARIES')+' WHERE PSA_SALARIE="'+Salarie+'"', True);
				If Not Q.EOF Then TF.PutValue('PFI_DADSCAT', Q.FindField('PSA_DADSCAT').AsString);
				Ferme(Q);
				Except
			End;
		End;

		Q := OpenSQL('SELECT MAX (PFI_RANG) AS RANG FROM INSCFORMATION WHERE PFI_ETABLISSEMENT="' + TF.GetValue('PFI_ETABLISSEMENT') + '" AND PFI_MILLESIME="' + MillesimeEC + '"', True);
		if not Q.eof then TF.PutValue('PFI_RANG', Q.FindField('RANG').AsInteger + 1)
		else TF.PutValue('PFI_RANG', 1);
		Ferme(Q);

		//PT8
		if (VH_Paie.PGForValidPrev = False) and (GetField('PFI_TYPEPLANPREV')<>'DIF') then
			TF.PutValue('PFI_ETATINSCFOR', 'VAL')
		Else
			TF.PutValue('PFI_ETATINSCFOR','ATT');

		TF.PutValue('PFI_TYPEPLANPREV','PLF');
		TF.PutValue('PFI_JOURSTAGE', JourStage);
		TF.PutValue('PFI_DUREESTAGE', Duree);
		TF.PutValue('PFI_HTPSTRAV', Duree);
		TF.PutValue('PFI_HTPSNONTRAV', 0);
		TF.PutValue('PFI_LIEUFORM', Lieu);
		TF.PutValue('PFI_FORMATION1', For1);
		TF.PutValue('PFI_FORMATION2', For2);
		TF.PutValue('PFI_FORMATION3', For3);
		TF.PutValue('PFI_FORMATION4', For4);
		TF.PutValue('PFI_FORMATION5', For5);
		TF.PutValue('PFI_FORMATION6', For6);
		TF.PutValue('PFI_FORMATION7', For7);
		TF.PutValue('PFI_FORMATION8', For8);
		TF.PutValue('PFI_NATUREFORM', Nature);
		TF.PutValue('PFI_NIVPRIORITE', '01');
		TF.PutValue('PFI_MOTIFINSCFOR', '2');

		if VH_Paie.PGForValoSalaire = 'VCR' then Salaire := ForTauxHoraireReel(TF.Getvalue('PFI_SALARIE'))
		else Salaire := ForTauxHoraireCategoriel(TF.GetValue('PFI_LIBEMPLOIFOR'), MillesimeEC);
		TF.PutValue('PFI_COUTREELSAL', Salaire * Duree * TauxChargeBudget);
		Q := OpenSQL('SELECT PFF_FRAISHEBERG,PFF_FRAISREPAS,PFF_FRAISTRANSP FROM FORFAITFORM ' +
		'WHERE PFF_MILLESIME="' + MillesimeEC + '" AND PFF_LIEUFORM="' + Lieu + '"' + //PT3
		' AND PFF_ETABLISSEMENT="' + TF.Getvalue('PFI_ETABLISSEMENT') + '"', True);
		FraisH := 0;
		FraisR := 0;
		FraisT := 0;
		if not Q.eof then
		begin
			FraisH := Q.FindField('PFF_FRAISHEBERG').AsFloat;
			FraisR := Q.FindField('PFF_FRAISREPAS').AsFloat;
			FraisT := Q.FindField('PFF_FRAISTRANSP').AsFloat;
		end;
		Ferme(Q);
		NbJours := JourStage;
		FraisH := FraisH * (NbJours - 1);
		if FraisH < 0 then FraisH := 0;
		If FraisH > 0 Then FraisR := FraisR * ((NbJours * 2) - 1) //PT7
		Else FraisR := FraisR * NbJours;
		if FraisR < 0 then FraisR := 0;
		TotalFrais := FraisH + FraisR + FraisT;
		TF.PutValue('PFI_FRAISFORFAIT', TotalFrais);

		//PT8
		If PGBundleInscFormation Then
		Begin
			TF.PutValue('PFI_NODOSSIER', GetNoDossierSalarie(Salarie));
			TF.PutValue('PFI_PREDEFINI', 'DOS');
		End
		Else
		Begin
			TF.PutValue('PFI_NODOSSIER', '000000');
			TF.PutValue('PFI_PREDEFINI', 'STD');
		End;

		TF.Post;
		Salarie := ReadTokenPipe(Ret,';');
	end;
	TF.EndUpdate;
	TF.EnableTom;
	{$IFNDEF EMANAGER}
	MajCoutPrev(LeStage,MillesimeEC,TF);
	{$ENDIF}
        
	if not ExisteSQL('SELECT PST_CODESTAGE FROM STAGE WHERE PST_CODESTAGE="' + LeStage + '" AND PST_MILLESIME="' + MillesimeEC + '"') then
	begin
		QAnim := OpenSQl('SELECT PFE_SALAIREANIM FROM EXERFORMATION WHERE PFE_MILLESIME="' + MillesimeEC + '"', True);
		salaireanim := 0;
		if not QAnim.eof then SalaireAnim := QAnim.FindField('PFE_SALAIREANIM').AsFloat;
		Ferme(QAnim);
		Q := OpenSQL('SELECT * FROM STAGE WHERE ' +
		'PST_CODESTAGE="' + LeStage + '" AND PST_MILLESIME="0000"', True);
		TobStage := Tob.Create('STAGE', nil, -1);
		TobStage.loadDetailDB('STAGE', '', '', Q, False);
		Ferme(Q);
		T := TobStage.FindFirst(['PST_CODESTAGE'], [LeStage], False);
		if T <> nil then
		begin
			T.ChangeParent(TobStage, -1);
			T.PutValue('PST_MILLESIME', MillesimeEC);
			T.PutValue('PST_COUTSALAIR', SalaireAnim * Duree * NbAnim);
			T.InsertOrUpdateDB;
		end;
		TobStage.Free;

		// Duplication investissement pour la formation
		Q := OpenSQL('SELECT * FROM INVESTSESSION ' +
		'WHERE PIS_CODESTAGE="' + LeStage + '" AND PIS_ORDRE=-1 AND PIS_MILLESIME="0000"', True); //DB2
		TobInvest := Tob.Create('INVESTSESSION', nil, -1);
		TobInvest.LoadDetailDB('INVESTSESSION', '', '', Q, False);
		Ferme(Q);
		TI := TobInvest.FindFirst(['PIS_CODESTAGE'], [LeStage], False);
		while TI <> nil do
		begin
			TI.ChangeParent(TobInvest, -1);
			TI.PutValue('PIS_MILLESIME', GetField('PFI_MILLESIME'));
			TI.InsertOrUpdateDB(False);
			TI := TobInvest.FindNext(['PIS_CODESTAGE'], [LeStage], False);
		end;
		TobInvest.Free;
	end;

	{$IFDEF EMANAGER}
	SetControltext('XX_WHERE', 'PST_CODESTAGE="' + lEStage + '" AND PST_MILLESIME="' + MillesimeEC + '"');

    //PT10 Fait par MajCoutsPrev à présent. Donc à ne faire que dans le cas de eManager
	TF.RefreshEntete;
	Bt := TToolBarButton97(GetControl('BCHERCHE'));
	if Bt <> nil then Bt.Click;
	{$ENDIF}
end;

Initialization
  registerclasses ( [ TOF_SAISIEINSCBUDGET ] ) ;
end.

