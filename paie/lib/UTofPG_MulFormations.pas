{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 14/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULFORMATIONS ()
Mots clefs ... : TOF;PGMULFORMATIONS
*****************************************************************
PT1 | 24/04/2003 | V_42  | JL | Développement pour CWAS
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2 | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
    |            |       |    |  pour les critères code salarié uniquement
PT3 | 14/06/2007 | V_720 | JL | Gestion partage formation
PT4 | 08/10/2007 | V_7   | FL | Emanager / Report / Prise en compte des salariés sortis
PT5 | 18/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT6 | 03/04/2008 | V_803 | FL | Adaptation partage formation
PT7 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT8 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + gestion elipsis salarié et responsable
PT9 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
}
Unit UTofPG_MulFormations;

Interface

Uses Controls,Classes,StdCtrls,pgoutils,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Fe_Main,Mul,
{$ELSE}
     MaineAGL,Emul,UTob,
{$ENDIF}
     sysutils,HCtrls,HEnt1,HMsgBox,UTOF,Hqry,ParamDat,LookUp,P5Def,PgoutilsFormation,EntPaie,PGOutils2,Ed_Tools ;

Type
  TOF_PGMULFORMATIONS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    THVMillesime:THValComboBox;
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    {$ENDIF}
    procedure ChangeMillesime(Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure RespElipsisClick(Sender : TObject);
    procedure ExitStage(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
//    {$IFDEF EMANAGER} //PT6
    procedure SalarieElipsisClick(Sender : TObject);
//    {$ENDIF}
    procedure ExitEdit(Sender : TObject);
end ;

Implementation

Uses GalOutil;

procedure TOF_PGMULFORMATIONS.OnLoad ;
var Session : THValComboBox;
    {$IFDEF EMANAGER}
    Whereresp : String;
    MultiNiveau : Boolean;
    {$ENDIF}
    Where : String;
begin
	Inherited ;
	Session := THValComboBox(GetControl('SESSION'));
	If Session <> Nil Then
	begin
		If Session.Value <> '' Then SetControlText('PFO_ORDRE',Session.Value)
		Else SetControlText('PFO_ORDRE','');
	end;
	{$IFDEF EMANAGER}
	If GetControltext('RESPONSFOR') <> '' then WhereResp := 'PSE_RESPONSFOR="'+getControlText('RESPONSFOR')+'" AND '
	else WhereResp := '';
	MultiNiveau := (GetCheckBoxState('CMULTINIVEAU')= CbChecked);
	If typeUtilisat = 'R' then
	begin
		If MultiNiveau then Where := 'PFO_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+'(PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
		' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))))'
		else Where := 'PFO_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE '+WhereResp+'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")';
	end
	else Where := '(PFO_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE (PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
	'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'"))))))';
	SetControlText('XX_WHERE',Where);
	{$ENDIf}

	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PFO_PREDEFINI'),GetControlText('NODOSSIER'),'PFO')); //PT5
	
	//PT9
	// N'affiche pas les salariés confidentiels
	If Where <> '' Then Where := Where + ' AND ';
	If PGBundleHierarchie Then
		Where := Where + 'PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
	Else
		Where := Where + 'PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
	SetControlText('XX_WHERE',Where);
end ;

procedure TOF_PGMULFORMATIONS.OnArgument (S : String ) ;
var {$IFDEF EMANAGER}
    Q : TQuery;
    {$ENDIF}
    THDate : THEdit;
    Resp,Stage,Edit : THEdit;
    Num : Integer;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    Arg : String;
    DD,DF : TDateTime;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        //DEBUT PT1
        THVMillesime := THValComboBox(GetControl('THVMILLESIME'));
        //PT4 - Début
        {$IFDEF EMANAGER}
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" ORDER BY PFE_DATEDEBUT DESC', True);
        If Not Q.EOF Then
        Begin
                SetControlText('PFO_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                SetControlText('PFO_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
        End;
        Ferme(Q);
        THVMillesime.Value := RendMillesimeEManager;
        {$ELSE}
        THVMillesime.Value := RendMillesimeRealise(DD,DF);
        SetControlText('PFO_DATEDEBUT',DateToStr(DD));
        SetControlText('PFO_DATEFIN',DateToStr(DF));
        {$ENDIF}
        //PT4 - Fin
        THVMillesime.OnChange := ChangeMillesime;
        THDate := THEdit(GetControl('PFO_DATEDEBUT'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('PFO_DATEFIN'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        SetControlCaption('LIBSTAGE','');
        SetControlCaption('LIBRESP','');
        SetControlCaption('LIBSALARIE','');
        Stage := THEdit(GetControl('PFO_CODESTAGE'));
        If Stage <> Nil Then Stage.OnExit := ExitStage;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
        end;
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        {$ENDIF}
        if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;
        SetControlVisible('BINSERT',False);
        {$IFDEF EMANAGER}
        If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
        else TypeUtilisat := 'S';
        TFMul(Ecran).Caption := 'Consultation des formations réalisées';
        UpdateCaption(TFMul(Ecran));
        Resp := THedit(Getcontrol('RESPONSFOR'));
        If Resp <> Nil Then Resp.OnElipsisClick := RespElipsisClick;
        //FIN PT1
        {$ELSE}
        If PGBundleHierarchie Then //PT8
        Begin
	        Resp := THedit(Getcontrol('PFO_RESPONSFOR'));
	        If Resp <> Nil Then Resp.OnElipsisClick := RespElipsisClick;
	    End;
	    
        {$ENDIF}
        Edit := THEdit(GetControl('PFO_SALARIE')); //PT6
        If Edit <> Nil then
        Begin
        	{$IFNDEF EMANAGER}If PGBundleHierarchie Then{$ENDIF} Edit.OnElipsisClick := SalarieElipsisClick; //PT8
            Edit.OnExit := ExitEdit;
        End;
        //DEBUT PT3
		If PGBundleCatalogue Then //PT8
		Begin
			If not PGDroitMultiForm then
				SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
			Else If V_PGI.ModePCL='1' Then 
				SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT7
		End;
			
		If PGBundleInscFormation then
		begin
			If not PGDroitMultiForm then
			begin
				SetControlEnabled('NODOSSIER',False);
				SetControlText   ('NODOSSIER',V_PGI.NoDossier);
				SetControlEnabled('PFO_PREDEFINI',False);
				SetControlText   ('PFO_PREDEFINI','');
				//SetControlProperty ('PFO_CODESTAGE', 'Plus', ' AND (PST_PREDEFINI="STD" OR (PST_PREDEFINI="DOS" AND PST_NODOSSIER="'+V_PGI.NoDossier+'"))'); //PT6
				//SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT7 //PT8
			end
        	Else If V_PGI.ModePCL='1' Then 
        	Begin
        		SetControlProperty ('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT7
				//SetControlProperty ('PFO_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT7 //PT8
			End;
		end
		Else
		begin
			SetControlVisible('PFO_PREDEFINI', False);
			SetControlVisible('TPFO_PREDEFINI',False);
			SetControlVisible('NODOSSIER',     False);
			SetControlVisible('TNODOSSIER',    False);
		end;
        //FIN PT3
end ;

procedure TOF_PGMULFORMATIONS.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+THVMillesime.Value+'"',True);
        If not Q.eof Then
        begin
                SetControlText('PFO_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PFO_DATEDEBUT_',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
                SetControlText('PFO_DATEFIN',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PFO_DATEFIN_',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
        end;
        Ferme(Q);
end;

procedure TOF_PGMULFORMATIONS.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULFORMATIONS.RespElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere : String;
    St : String;
    StOrder : String;
{$ENDIF}
begin
        {$IFNDEF EMANAGER}
        //PT6 - Début
        //St := ' SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM FROM SALARIES ';
        //StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES)';
        //LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,'', True,-1);
        ElipsisResponsableMultidos (Sender); 
        //PT6 - Fin
        {$ELSE}
        If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
        else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
        StOrder := 'PSI_LIBELLE';
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
        {$ENDIF}
end;

procedure TOF_PGMULFORMATIONS.ExitStage(Sender : TObject);
Var Plus : String;
begin
        If GetControlText('PFO_CODESTAGE') <> '' Then
        begin
        	Plus :=  'PSS_CODESTAGE="'+GetControlText('PFO_CODESTAGE')+'" AND PSS_MILLESIME="'+THVMillesime.Value+'"'; //PT6
        	If PGBundleInscFormation Then
        	Begin
        		If Not PGDroitMultiForm Then 
        			Plus := Plus + DossiersAInterroger('',V_PGI.NoDossier,'PSS',True,True)  //PT7
        		Else
        			Plus := Plus + DossiersAInterroger('','','PSS',True,True);  //PT7
        	End;
            SetControlProperty('SESSION','Plus', Plus);
            SetControlEnabled('SESSION',True);
        end
        Else SetControlEnabled('SESSION',False);
end;

procedure TOF_PGMULFORMATIONS.GrilleDblClick(Sender : TObject);
var St : String;
    Q_Mul : THQuery ;
    {$IFNDEF EAGLCLIENT}
    //Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
begin
        {$IFNDEF EAGLCLIENT}
        //Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If Q_MUL.FindField('PFO_SALARIE').AsString = '' Then
        begin
                PGIBox('Vous devez choisir un salarié',Ecran.Caption);
                Exit;
        end;
        st  := Q_MUL.FindField('PFO_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PFO_ORDRE').AsInteger)+';'+
        Q_MUL.FindField('PFO_MILLESIME').AsString+';'+Q_Mul.FindField('PFO_SALARIE').AsString;
        AglLanceFiche('PAY','FORMATIONS','',St,'ACTION=CONSULTATION')
end;


procedure TOF_PGMULFORMATIONS.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere,StOrder : String;
{$ENDIF}
begin
     //PT4 - Début
     {$IFDEF EMANAGER}
     StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '") AND '; //PT2
     If (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) Then
     Begin
       If typeUtilisat = 'R' then StWhere := StWhere + 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'
       Else StWhere := StWhere + 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")';
     End
     Else
     Begin
     //{$ENDIF}
       If typeUtilisat = 'R' then StWhere := StWhere + 'PSE_CODESERVICE IN '+
     '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
     ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
     ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
       else StWhere := StWhere + 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     //{$IFDEF EMANAGER}
     End;
     //{$ENDIF}
     //PT4 - Fin
     If GetControlText('RESPONSFOR') <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
     (*{$IFNDEF EMANAGER} //PT4
     If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     {$ENDIF}*)
     StOrder := 'PSA_SALARIE';               
     StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT2
     LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
     //PT6 - Début
     {$ELSE}
    //If PGBundleInscFormation Then //PT8
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
    	//PT6 - Fin
     {$ENDIF}
end;

procedure TOF_PGMULFORMATIONS.ExitEdit(Sender : TObject);
var edit : thedit;
begin
        edit := THEdit(Sender);
        if edit  <>  nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit,10);
end;


Initialization
  registerclasses ( [ TOF_PGMULFORMATIONS ] ) ;
end.


