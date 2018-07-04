{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 10/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MULINSCFOR ()
Mots clefs ... : TOF;MULINSCFOR
*****************************************************************
PT1  | 24/04/2003 | V_42  | JL | Développement pour CWAS + modif lookuplist
PT2  | 13/10/2003 | V_5.0 | JL | Modification validation des inscriptions au previsionnel
PT3  | 12/04/2006 | V_65  | JL | FQ 13035 Millésime consultation DIF = Millésime plan de formation
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT4  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    |  pour les critères code salarié uniquement
PT5  | 11/04/2007 | V_702 | FL | FQ 14047 Forçage de la formation sur le OnLoad et non pas sur
     |            |       |    |  le OnArgument afin d'outrepasser les filtres. Ajout du
     |            |       |    |  mode d'appel "RECUPPREVISIONNEL" utilisé dans la fiche d'inscription
     |            |       |    |  afin de récupérer le prévisionnel d'une formation.
PT6  | 07/01/2008 | V_82  | FL | Correction de bugs mémoire
PT7  | 18/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT8  | 03/04/2008 | V_803 | FL | Adaptation partage formation
PT9  | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT10 | 03/04/2008 | V_803 | FL | Ajout de la coche "Déjà inscrits" + correction bug mémoire
PT11 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + Gestion elipsis salarié et responsable
PT12 | 28/05/2008 | V_804 | FL | Pas d'ajout du AND dans la clause Where si pas de critère précédent
PT13 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT14 | 10/06/2008 | V_804 | FL | Ajout des champs salariés
PT15 | 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}
Unit UTOFPGMul_InscFor;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB,Mul,Fiche,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAGL,emul,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HTB97,UTob,LookUp,PgoutilsFormation,
     ed_tools,EntPaie,Hqry,PGOutils2,HStatus,Menus,PGOutils;

Type
  TOF_MULINSCFOR = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ; //PT6
    private
    Millesime,Arg : String;
    LeMillesimeSession : String; //PT8
{$IFDEF EMANAGER}
    AutoValid : Boolean;
    TypeUtilisat : String;
{$ENDIF}
    Formation : String; //PT5
//    THVMillesime:THMultiValComboBox;

    Q_Mul      : THQuery ;
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
    procedure GrilleDblClick(Sender:TObject);
    procedure CreerInscription(Sender : TObject);
    //{$IFDEF EMANAGER}
    procedure SalarieElipsisClick (Sender : TObject);
    procedure ResponsableELipsisCLick(Sender : TObject);
    //{$ENDIF}
    procedure ExitEdit(Sender : TObject);
    procedure Validation(Sender : Tobject);
//    procedure PlanifierSession(Sender : Tobject);
//    procedure SessionRealise(Sender : Tobject);
    procedure ChargerTobRecup (NumSalarie:String='');  //PT5
//    procedure EnvoyeMailDecision(Rang : Integer;Stage,Millesime,Etablissement : String);
    end ;
    Var
    PGTobRecupPrev : Tob;

Implementation

Uses GalOutil,HDebug,P5Def;

//PT6
procedure TOF_MULINSCFOR.OnClose ;
begin
  Inherited ;
   If (Arg <> 'RECUPPREVISIONNEL') And (Arg <> 'RECUPPRESENCE') Then //PT10
    If Assigned(PGTobRecupPrev) Then FreeAndNil(PGTobRecupPrev);
end ;

procedure TOF_MULINSCFOR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_MULINSCFOR.OnLoad ;
var Where : String;
{$IFDEF EMANAGER}
    MultiNiveau : Boolean;
{$ENDIF}
begin
  Inherited ;
	If Arg = 'DEMANDEDIF' then SetControlText('PFI_TYPEPLANPREV','DIF');
	
	If GetControlText('TYPEINSC')='TOUS' then Where := ''
	else If GetControlText('TYPEINSC')='ACC' then Where := 'PFI_ETATINSCFOR="VAL"'
	else If GetControlText('TYPEINSC')='REF' then Where := 'PFI_ETATINSCFOR="REF"'
	else If GetControlText('TYPEINSC')='REP' then Where := 'PFI_ETATINSCFOR="REP"'
	else If GetControlText('TYPEINSC')='NT'  then Where := 'PFI_ETATINSCFOR="ATT"'
	else If GetControlText('TYPEINSC')='ANN' then Where := 'PFI_ETATINSCFOR="NAN"'; //PT10

	If Arg = 'RECUPPREVISIONNEL' Then
	Begin
	    SetControlText('PFI_CODESTAGE',Formation);  //PT5
	
	    If Where <> '' then Where := Where + ' AND '; //PT12
	    Where := Where + 'PFI_SALARIE<>"" ';
	End;

	{$IFDEF EMANAGER}
	MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
	If typeUtilisat = 'R' then
	begin
	     If Arg = 'VALIDATION' then
	     begin
	        If AutoValid = True then
	        begin
	                If MultiNiveau then
	                begin
	                     If Where <> '' then Where := Where + 'AND (PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	                    ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
	                    else Where := '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	                    ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
	                end
	                else
	                begin
	                     If Where <> '' then Where := Where + 'AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"'
	                    else Where := 'PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"';
	                end;
	        end
	        else
	        begin
	                    If Where <> '' then Where := Where + 'AND (PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	                    ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
	                    else Where := '(PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	                    ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
	        end;
	     end
	     else
	     begin
	        If MultiNiveau then
	        begin
	                If Where <> '' then Where := Where + 'AND (PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	                    ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
	                        else Where := '(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
	                    ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
	        end
	        else
	        begin
	             If Where <> '' then Where := Where + 'AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"'
	             else Where := 'PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"';
	        end;
	     end;
	end
	else
	begin
	        If Where <> '' then Where := Where + ' AND PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
	        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))'
	                else Where := 'PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
	        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
	end;
	where := Where + ' AND PFI_CURSUS=""';
	{$ELSE}
	If GetControlText('DETAILCURSUS') <> 'X' then
	begin
		If ((GetControlText('FILTRECURSUS') = 'AUC') Or (GetControlText('FILTRECURSUS') = 'FORM') Or 
		   (GetControlText('FILTRECURSUS') = 'CURSUS')) And (Where <> '') Then Where := Where + ' AND ';

        If GetControlText('FILTRECURSUS') = 'AUC'    then Where := Where + ' (PFI_CURSUS="" OR PFI_CODESTAGE="--CURSUS--")';
        If GetControlText('FILTRECURSUS') = 'FORM'   then Where := Where + ' PFI_CURSUS=""';
        If GetControlText('FILTRECURSUS') = 'CURSUS' then Where := Where + ' (PFI_CURSUS<>"" AND PFI_CODESTAGE="--CURSUS--")';
	end
	else
	begin
		If ((GetControlText('FILTRECURSUS') = 'AUC') Or (GetControlText('FILTRECURSUS') = 'FORM') Or 
		   (GetControlText('FILTRECURSUS') = 'CURSUS')) And (Where <> '') Then Where := Where + ' AND ';
		   
        If GetControlText('FILTRECURSUS') = 'AUC'    then Where := Where + ' PFI_CODESTAGE<>"--CURSUS--"';
        If GetControlText('FILTRECURSUS') = 'FORM'   then Where := Where + ' PFI_CURSUS=""';
        If GetControlText('FILTRECURSUS') = 'CURSUS' then Where := Where + ' (PFI_CURSUS<>"" AND PFI_CODESTAGE<>"--CURSUS--")';
	end;
	{$ENDIf}
	If GetControlText('REALISATION') = 'OUI' then
	begin
	     If Where <> '' then Where := Where + ' AND PFI_REALISE="X"'
	     else Where := 'PFI_REALISE="X"';
	end
	else
	If GetControlText('REALISATION') = 'NON' then
	begin
	     If Where <> '' then Where := Where + ' AND PFI_REALISE="-"'
	     else Where := 'PFI_REALISE="-"';
	end;
	//PT10 - Début
	If (Arg = 'RECUPPREVISIONNEL') And (GetControlText('CKDEJAINSCRITS') = 'X') Then
	Begin
	    If Where <> '' Then Where := Where + ' AND ';
	    Where := Where + ' PFI_SALARIE NOT IN (SELECT DISTINCT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+Formation+
	             '" AND PFO_MILLESIME="'+LeMillesimeSession+'" AND PFO_ETATINSCFOR<>"REP")';
	End;
	//PT10 - Fin

    //PT13
    // N'affiche pas les salariés confidentiels
    If Where <> '' Then Where := Where + ' AND ';
    If PGBundleHierarchie Then
    	Where := Where + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))' //PT15
    Else
    	Where := Where + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))';  //PT15
	
	SetControlText('XX_WHERE',Where);
	//Setcontroltext('PFI_MILLESIME',THVMillesime.Value);

	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PFI_PREDEFINI'),GetControlText('NODOSSIER'),'PFI')); //PT7
end ;

procedure TOF_MULINSCFOR.OnArgument (S : String ) ;
Var Num:Integer;
    Edit : THEdit;
    BValid,BNew : TToolBarButton97;
//    Formation : String;  //PT5
    DD,DF : TDateTime;
begin
  Inherited ;
  PGTobRecupPrev := Nil;
  Arg := ReadTokenPipe(S,';');
  {$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FLISTE'));
  {$ELSE}
  Liste := THGrid(GetControl('FLISTE'));
  {$ENDIF}
  if Liste = nil then Exit;
  Q_Mul := THQuery(Ecran.FindComponent('Q'));

  If (Arg = 'RECUPPRESENCE') Or (Arg = 'RECUPPREVISIONNEL') then //PT5
  begin
     SetControlVisible('BSelectAll',true);
     SetControlVisible('BVALID',True);
     SetControlVisible('BOuvrir',False);
     SetControlVisible('BInsert',False);
     Ecran.Caption := 'Récupération du prévisionnel';
     Formation := ReadTokenPipe(S,';');
     SetControlText('PFI_CODESTAGE',Formation);
     //PT10 - Début
     LeMillesimeSession := ReadTokenPipe(S,';');
     //PT10 - Fin
     SetControlEnabled('PFI_CODESTAGE',False);
     SetControltext('REALISATION','NON');
     //PT5 - Début
     If Arg = 'RECUPPRESENCE' Then
     Begin
          SetControlProperty('BVALID','Hint','Récupérer dans la feuille de présence');
          SetControlText('PFI_TYPEPLANPREV','DIF');
          SetControlVisible('CKDEJAINSCRITS', True);
     End
     Else
     Begin
          SetControlProperty('BVALID','Hint','Récupérer dans la feuille d''inscription');
          SetControlText('PFI_TYPEPLANPREV','');
		  SetControlVisible('CKDEJAINSCRITS', True); //PT10
     End;
     //PT5 - Fin
  End;


  {$IFDEF EMANAGER}
  If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
        'WHERE PHO_NIVEAUH<=2 AND PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"') then AutoValid := True
  else AutoValid := False;
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  If Arg = 'VALIDATION' then
  begin
       If Not AutoValid then
       begin
            SetControlChecked('CMULTINIVEAU',True);
            SetControlEnabled('CMULTINIVEAU',false);
       end;
  end;
  SetControlVisible('PAvance',False);
  SetControlProperty('PAvance','TabVisible',False);
  If arg ='DEMANDEDIF' then SetControlText('TYPEINSC','NT');
  {$ENDIF}

  {$IFNDEF EMANAGER}
  Ecran.Caption := 'Consultation des inscriptions au prévisionnel';
  If Arg = 'CONSULTATION' then
  begin
     Ecran.Caption := 'Détail des inscriptions';
     SetControlVisible('BInsert',False);
     SetControlVisible('BVALID',False);
     SetControlVisible('BSelectAll',False);
  //   MenuDif := TPopUpMenu(GetControl('MENUDIF'));
  //   If MenuDif <> Nil then
   //  begin
    //      MenuDif.Items[0].OnClick := PlanifierSession;
     //     MenuDif.Items[1].OnClick := SessionRealise;
    // end;
//     SetControlVisible('BREALISE',True);
  end;
  UpdateCaption(Ecran);
  {$ELSE}
//  SetControlEnabled('THVMILLESIME',False);
  If Arg = 'CONSULTATION' then
  begin
     Ecran.Caption := 'Détail des inscriptions';
//     MenuDif := TPopUpMenu(GetControl('TPMAIL'));
//     MenuDif.Items[0].OnClick := PlanifierSession;
//     MenuDif.Items[1].OnClick := SessionRealise;
//     SetControlVisible('BREALISE',True);
       SetControlVisible('BInsert',False);
  end;
  If Arg = 'VALIDATION' then
  begin
        SetControlVisible('BSelectAll',true);
        SetControlVisible('BVALID',True);
        SetControlProperty('BVALID','Hint','Validation des inscriptions');
        Ecran.Caption := 'Validation des inscriptions';
        SetControlVisible('BINsert',False);
  end
  Else
  begin
        SetControlVisible('BVALID',False);
        SetControlVisible('BSelectAll',False);
  end;
  UpdateCaption(Ecran);
{  If arg ='DEMANDEDIF' then
  begin
       SetControlVisible('BSelectAll',true);
       SetControlVisible('BVALID',True);
       SetControlProperty('BVALID','Hint','Validation des inscriptions');
  end;}
  //debut PT1
  SetControlCaption('LIBSAL','');
  SetControlCaption('LIBRESP','');
  {$ENDIF}
  
  //PT8 - Début
  Edit := ThEdit(GetControl('PFI_RESPONSFOR_'));
  If Edit <> Nil then
  begin
  		If PGBundleHierarchie Then Edit.OnElipsisCLick := ResponsableElipsisCLick; //PT11
        Edit.OnExit := Exitedit;
  end;
  Edit := THEdit(GetControl('PFI_SALARIE'));
  If Edit <> Nil then
  begin
  		{$IFNDEF EMANAGER}If PGBundleHierarchie Then{$ENDIF} //PT11
	        Edit.OnElipsisClick := SalarieElipsisClick;
        Edit.OnExit := ExitEdit;
  end;
  //PT8 - Fin
  
  //Fin PT1
  Millesime:='';
  //DEBUT PT3
  If (Arg = 'RECUPPRESENCE') or (arg='DEMANDEDIF') Or (Arg='RECUPPREVISIONNEL') then millesime := RendMillesimeRealise(DD,DF) //PT5
  else Millesime := RendMillesimePrevisionnel;
    //FIN PT3
  If Arg <> 'DEMANDEDIF' then SetControlText('PFI_MILLESIME',Millesime);
  SetControlText('XX_WHERE','');
//  THVMillesime:=THMultiValComboBox(GetControl('THVMILLESIME'));
// THVMillesime.Value:=Millesime;
  SetControlCaption('LIBSTAGE','');
  For Num := 1 to VH_Paie.NBFormationLibre do
  begin
    if Num >8 then Break;
    VisibiliteChampFormation (IntToStr(Num),GetControl ('PFI_FORMATION'+IntToStr(Num)),GetControl ('TPFI_FORMATION'+IntToStr(Num)));
  end;

  //PT14
  {$IFNDEF EMANAGER}
  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
      if Num >4 then Break;
      VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFI_TRAVAILN'+IntToStr(Num)),GetControl ('TPFI_TRAVAILN'+IntToStr(Num)));
  end;
  VisibiliteStat (GetControl ('PFI_CODESTAT'),GetControl ('TPFI_CODESTAT')) ;
  {$ENDIF}

  if Liste <> NIL then Liste.OnDblClick := GrilleDblClick;
  {$IFNDEF EMANAGER}
  If Arg = 'CONSULTATION' then  SetControlVisible('DETAILCURSUS',True);
    SetControlText('FILTRECURSUS','AUC');
  {$ENDIF}
  SetControlProperty('PFI_ETABLISSEMENT','Vide',True);
  SetControlProperty('PFI_ETABLISSEMENT','VideString','<<TOUS>>');
  BValid := TToolBarButton97(GetControl('BVALID'));
  If BValid <> Nil then BValid.OnClick := Validation;
  If Arg = 'DEMANDEDIF' then
  begin
       BNew := TToolBarButton97(GetControl('BInsert'));
       If BNew <> Nil then BNew.OnClick := CreerInscription;
       SetControlText('PFI_TYPEPLANPREV','DIF');
       SetControlEnabled('PFI_TYPEPLANPREV',False);
       ecran.Caption := 'Demandes de Droit Individuel à la Formation';
       UpdateCaption(Ecran);
  end;
  
  	If PGBundleCatalogue Then //PT11
  	Begin
  		If not PGDroitMultiForm then
			SetControlProperty ('PFI_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
       		SetControlProperty ('PFI_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT9
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFI_PREDEFINI',False);
			SetControlText   ('PFI_PREDEFINI','');
			//SetControlProperty ('PFI_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT9 //PT11
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT9
       		//SetControlProperty ('PFI_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT9 //PT11
       	End;
	end
	else
	begin
		SetControlVisible('PFI_PREDEFINI', False);
		SetControlVisible('TPFI_PREDEFINI',False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
end ;

procedure TOF_MULINSCFOR.GrilleDblClick(Sender:TObject);
var St : String;
    Retour : String;
begin
Retour := '';
//PT5 - Début
{ Sur le double clic, on récupère le numéro du salarié en cours et on le passe à la
  méthode ChargeToRecup car une ligne double-cliquée n'est pas une ligne sélectionnée }
If (Arg = 'RECUPPRESENCE') Or (Arg='RECUPPREVISIONNEL') then
begin
     {$IFDEF EAGLCLIENT}
     TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
     {$ENDIF}
     If Q_MUL.FindField('PFI_RANG').AsInteger>0 Then
     begin
          St := Q_MUL.FindField('PFI_SALARIE').AsString;
          ChargerTobRecup (St);
     End;
end
else
//PT5 - Fin
begin
     {$IFDEF EAGLCLIENT}
     TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
     {$ENDIF}
     If Q_MUL.FindField('PFI_RANG').AsInteger<=0 Then
     begin
          PGIBox('Vous devez choisir un salarié',TFMul(Ecran).Caption);
          Exit;
     end;
     St := Q_MUL.FindField('PFI_CODESTAGE').AsString;
     st :=IntToStr(Q_MUL.FindField('PFI_RANG').AsInteger)+';'+Q_MUL.FindField('PFI_CODESTAGE').AsString+';'+
     Q_MUL.FindField('PFI_MILLESIME').AsString+';'+Q_Mul.FindField('PFI_ETABLISSEMENT').AsString;
     If Arg <> 'VALIDATION' then Retour := AglLanceFiche('PAY','INSCFORMATION','',St,'SAISIEDIF;;'+Q_MUL.FindField('PFI_CODESTAGE').AsString+';'+Q_MUL.FindField('PFI_MILLESIME').AsString+';'+Q_MUL.FindField('PFI_SALARIE').AsString)  //PT6
     else Retour := AglLanceFiche('PAY','INSCFORMATION','',St,'VALIDATION;BUDGET;;;;ACTION=MODIFICATION');
end;
     If Retour = 'MODIF' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCHerche);
end;

procedure TOF_MULINSCFOR.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}var StOrder,StWhere : String;{$ENDIF}
begin
	{$IFDEF EMANAGER}
		If typeUtilisat = 'R' then StWhere := 'PSE_CODESERVICE IN '+
			'(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
			' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
			' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
		else StWhere := 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
			'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
		If GetControlText('RESPONSFOR') <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
		If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
		StOrder := 'PSA_SALARIE';
		StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT4
		LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
	{$ELSE}
		//PT8 - Début
		//If PGBundleInscFormation Then //PT11
			ElipsisSalarieMultidos (Sender)
		//Else 
			//Inherited;
		//PT8 - Fin
	{$ENDIF}
end;

procedure TOF_MULINSCFOR.ResponsableElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var StWhere : String;
    St : String;
    StOrder : String;
{$ENDIF}
begin
	{$IFNDEF EMANAGER}
		ElipsisResponsableMultidos (Sender);  //PT8
	{$ELSE}
		If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
			' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
		else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
			'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
		StOrder := 'PSI_LIBELLE';
		LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
	{$ENDIF}
end;


procedure TOF_MULINSCFOR.ExitEdit(Sender : TObject);
var edit : thedit;
begin
        edit := THEdit(Sender);
        if edit  <>  nil then   //AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit,10);
end;

procedure TOF_MULINSCFOR.Validation(Sender : Tobject);
var i,j : Integer;
    TobInscFor : Tob;
    Q : TQuery;
    Rang : Integer;
    Stage,Etab,Millesime : String;
begin
   if Q_Mul = nil then Exit;
  if (Liste.NbSelected = 0) and (TFMul(Ecran).BSelectAll.Down = False) then
  begin
    PGIBox('Aucun élément sélectionné', Ecran.Caption);
    Exit;
  end;
  If (Arg = 'RECUPPRESENCE') Or (Arg = 'RECUPPREVISIONNEL') then //PT5
  begin
     ChargerTobRecup;
     Exit;
  end;
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
       {$IFDEF EAGLCLIENT}
       if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
       {$ENDIF}
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Rang := TFmul(Ecran).Q.FindField('PFI_RANG').AsInteger;
      Stage := TFmul(Ecran).Q.FindField('PFI_CODESTAGE').asstring;
      Millesime := TFmul(Ecran).Q.FindField('PFI_MILLESIME').asstring;
      Etab := TFmul(Ecran).Q.FindField('PFI_ETABLISSEMENT').asstring;
      begin
        Q := OpenSQL('SELECT * FROM INSCFORMATION WHERE '+
        'PFI_ETABLISSEMENT="'+Etab+'" '+
        'AND PFI_RANG='+IntToStr(Rang)+' '+
        'AND PFI_CODESTAGE="'+Stage+'" '+
        'AND PFI_MILLESIME="'+Millesime+'"', True);
        TobInscFor := Tob.Create('INSCFORMATION', nil, -1);
        TobInscFor.LoadDetailDB('INSCFORMATION', '', '', Q, False);
        Ferme(Q);
        for j := 0 to TobInscFor.Detail.Count - 1 do
        begin
          {$IFDEF EMANAGER}
          if TobInscFor.Detail[j].GetValue('PFI_NATUREFORM') = '002' then continue;
          {$ENDIF}
          TobInscFor.Detail[j].putValue('PFI_ETATINSCFOR','VAL');
          TobInscFor.Detail[j].UpdateDB(False);
        end;
        TobInscFor.Free;
      end;
      MoveCurProgressForm(RechDom('PGSTAGEFORM',Stage,False));
    end;
    FiniMoveProgressForm;
  end
  else if liste.AllSelected then
  begin
     {$IFDEF EAGLCLIENT}
     if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
        Rang := TFmul(Ecran).Q.FindField('PFI_RANG').AsInteger;
        Stage := TFmul(Ecran).Q.FindField('PFI_CODESTAGE').asstring;
        Millesime := TFmul(Ecran).Q.FindField('PFI_MILLESIME').asstring;
        Etab := TFmul(Ecran).Q.FindField('PFI_ETABLISSEMENT').asstring;
        Q := OpenSQL('SELECT * FROM INSCFORMATION WHERE '+
        'PFI_ETABLISSEMENT="'+Etab+'" '+
        'AND PFI_RANG='+IntToStr(Rang)+' '+
        'AND PFI_CODESTAGE="'+Stage+'" '+
        'AND PFI_MILLESIME="'+Millesime+'"', True);
        TobInscFor := Tob.Create('INSCFORMATION', nil, -1);
        TobInscFor.LoadDetailDB('INSCFORMATION', '', '', Q, False);
        Ferme(Q);
        for j := 0 to TobInscFor.Detail.Count - 1 do
        begin
          {$IFDEF EMANAGER}
          if TobInscFor.Detail[j].GetValue('PFI_NATUREFORM') = '002' then continue;
          {$ENDIF}
          TobInscFor.Detail[j].PutValue('PFI_ETATINSCFOR','VAL');
          TobInscFor.Detail[j].UpdateDB(False);
        end;
        TobInscFor.Free;
      MoveCurProgressForm(RechDom('PGSTAGEFORM',Stage,False));
      Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  end;
  TFMul(Ecran).bSelectAll.Down := False;
  Liste.AllSelected := False;
  Liste.ClearSelected;
  TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

{procedure TOF_MULINSCFOR.PlanifierSession(Sender : TObject);
var Regroupement,Millesime : String;
begin
     ChargerTobRecup;
     Regroupement := GetControlText('PFI_CODESTAGE');
     Millesime := GetControlText('THVMILLESIME');
     If Regroupement = '' then
     begin
          PGIBox ('Vous devez renseigner la formation',Ecran.Caption);
          Exit;
     end;
     AGLLanceFiche('PAY','MUL_SESSIONSTAGE','','','RECUPPREVPLA;'+Regroupement+';'+Millesime);
     PGTobRecupPrev.Free;
end;        }

{procedure TOF_MULINSCFOR.SessionRealise(Sender : Tobject);
var Regroupement,Millesime : String;
begin
     ChargerTobRecup;
     Regroupement := GetControlText('PFI_CODESTAGE');
     Millesime := GetControlText('THVMILLESIME');
     If Regroupement = '' then
     begin
          PGIBox ('Vous devez renseigner la formation',Ecran.Caption);
          Exit;
     end;
     AGLLanceFiche('PAY','MUL_SESSIONSTAGE','','','RECUPPREVREA;'+Regroupement+';'+Millesime);
     PGTobRecupPrev.Free;
end;     }

procedure TOF_MULINSCFOR.ChargerTobRecup (NumSalarie:String='');  //PT5
var T : Tob;
    Salarie,TypePlan : String;
    i : Integer;
    Q : TQuery;
begin
  if ((Liste.nbSelected) > 0) and (not Liste.AllSelected) then
  begin
    If Assigned(PGTobRecupPrev) Then FreeAndNil (PGTobRecupPrev); //PT6
    PGTobRecupPrev := Tob.Create('RecupInsc',Nil,-1);
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
    begin
       {$IFDEF EAGLCLIENT}
       if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
       {$ENDIF}
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ENDIF}
      Salarie := TFmul(Ecran).Q.FindField('PFI_SALARIE').asstring;
      Q := OpenSQL('SELECT PFI_TYPEPLANPREV FROM INSCFORMATION '+
      'WHERE PFI_RANG='+IntToStr(TFmul(Ecran).Q.FindField('PFI_RANG').AsInteger)+' AND '+
      'PFI_CODESTAGE="'+TFmul(Ecran).Q.FindField('PFI_CODESTAGE').asstring+'" AND '+
      'PFI_MILLESIME="'+TFmul(Ecran).Q.FindField('PFI_MILLESIME').asstring+'" AND '+
      'PFI_ETABLISSEMENT="'+TFmul(Ecran).Q.FindField('PFI_ETABLISSEMENT').asstring+'"',True);
      If Not Q.Eof then TypePlan := Q.FindField('PFI_TYPEPLANPREV').AsString
      else TypePlan := '';
      Ferme(Q);
      T := Tob.Create('UnInsc',PGTobRecupPrev,-1);
      T.AddChampSupValeur('SALARIE',Salarie);
      T.AddChampSupValeur('TYPEPLAN',TypePlan);
      MoveCurProgressForm(RechDom('PGSALARIE',Salarie,False));
    end;
    FiniMoveProgressForm;
  end
  else if liste.AllSelected then
  begin
    If Assigned(PGTobRecupPrev) Then FreeAndNil (PGTobRecupPrev); //PT6
    PGTobRecupPrev := Tob.Create('RecupInsc',Nil,-1);
    {$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) then TFMul(Ecran).Fetchlestous;
    {$ENDIF}
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
    InitMove(TFmul(Ecran).Q.RecordCount, '');
    Q_Mul.First;
    while not Q_Mul.EOF do
    begin
        Salarie := TFmul(Ecran).Q.FindField('PFI_SALARIE').asstring;
        Q := OpenSQL('SELECT PFI_TYPEPLANPREV FROM INSCFORMATION '+
      'WHERE PFI_RANG='+IntToStr(TFmul(Ecran).Q.FindField('PFI_RANG').AsInteger)+' AND '+
      'PFI_CODESTAGE="'+TFmul(Ecran).Q.FindField('PFI_CODESTAGE').asstring+'" AND '+
      'PFI_MILLESIME="'+TFmul(Ecran).Q.FindField('PFI_MILLESIME').asstring+'" AND '+
      'PFI_ETABLISSEMENT="'+TFmul(Ecran).Q.FindField('PFI_ETABLISSEMENT').asstring+'"',True);
      If Not Q.Eof then TypePlan := TFmul(Ecran).Q.FindField('PFI_TYPEPLANPREV').asstring
      else TypePlan := '';
      Ferme(Q);
        T := Tob.Create('UnInsc',PGTobRecupPrev,-1);
        T.AddChampSupValeur('SALARIE',Salarie);
        T.AddChampSupValeur('TYPEPLAN',TypePlan);
        MoveCurProgressForm(RechDom('PGSALARIE',Salarie,False));
        Q_Mul.Next;
    end;
    FiniMoveProgressForm;
  //PT5 - Début
  End
  // Cas où on a double cliqué sur une ligne, on charge la TOB par rapport au n° passé en paramètre
  Else if (NumSalarie <> '') Then
  Begin
    If Assigned(PGTobRecupPrev) Then FreeAndNil (PGTobRecupPrev); //PT6
    PGTobRecupPrev := Tob.Create('RecupInsc',Nil,-1);
    InitMoveProgressForm(nil, 'Début du traitement', 'Veuillez patienter SVP ...', Liste.nbSelected, FALSE, TRUE);
    InitMove(Liste.nbSelected, '');
    Salarie := NumSalarie;
    Q := OpenSQL('SELECT PFI_TYPEPLANPREV FROM INSCFORMATION '+
      'WHERE PFI_RANG='+IntToStr(TFmul(Ecran).Q.FindField('PFI_RANG').AsInteger)+' AND '+
      'PFI_CODESTAGE="'+TFmul(Ecran).Q.FindField('PFI_CODESTAGE').asstring+'" AND '+
      'PFI_MILLESIME="'+TFmul(Ecran).Q.FindField('PFI_MILLESIME').asstring+'" AND '+
      'PFI_ETABLISSEMENT="'+TFmul(Ecran).Q.FindField('PFI_ETABLISSEMENT').asstring+'"',True);
    If Not Q.Eof then TypePlan := TFmul(Ecran).Q.FindField('PFI_TYPEPLANPREV').asstring
    else TypePlan := '';
    Ferme(Q);
    T := Tob.Create('UnInsc',PGTobRecupPrev,-1);
    T.AddChampSupValeur('SALARIE',Salarie);
    T.AddChampSupValeur('TYPEPLAN',TypePlan);
    MoveCurProgressForm(RechDom('PGSALARIE',Salarie,False));
    FiniMoveProgressForm;
  //PT5 - Fin
  end;
  TFMul(Ecran).bSelectAll.Down := False;
  Liste.AllSelected := False;
  Liste.ClearSelected;
  TFMul(Ecran).BAnnulerClick(TFMul(Ecran).BAnnuler);
end;

procedure TOF_MULINSCFOR.CreerInscription(Sender : TObject);
var Retour : String;
begin

     Retour := AGLLanceFiche('PAY','INSCFORMATION','','','GESTIONDIF;DIF;'+GetControlText('PFI_SALARIE')+';;;ACTION=CREATION');
     If Retour = 'MODIF' then TFMul(Ecran).BChercheClick(TFMul(Ecran).BCHerche);
end;

//procedure TOF_MULINSCFOR.EnvoyeMailDecision(Rang : Integer;Stage,Millesime,Etablissement : String);
//begin
//end;

Initialization
  registerclasses ( [ TOF_MULINSCFOR ] ) ;
end.
