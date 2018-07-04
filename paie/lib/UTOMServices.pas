{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 20/03/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SERVICES
Mots clefs ... : TOM;SERVICES
*****************************************************************
PT1 | 01/07/2002 | V585  | JL | Maj des champs assistantes dans la table DEPORTSAL
PT2 | 07/01/2003 | V_42  | JL | Gestion accès formation et éléments variables selon seria
PT3 | 24/09/2003 | V_42  | JL | Gestion Maj responsable formation dans deportsal
PT4 | 26/09/2006 | V_70  | JL | FQ 13520 Suppression des masks pour web access
PT5 | 31/05/2007 | V_72  | JL | FQ 14021 Ajout paramsoc pour gestion intervenants exterieurs
PT6 | 13/02/2008 | V_803 | FL | Utilisation tablette PGSALARIEINT aussi dans le cas du bundle hierarchie + divers aménagements de code
PT7 | 20/02/2008 | V_803 | FL | Mise à jour des données hierarchiques dans les tables reprenant les informations de DEPORTSAL
PT8 | 19/03/2008 | V_803 | FL | Ajout du n° de dossier qui est repris dans le nom du service
PT9 | 02/04/2008 | V_803 | FL | Aménagements du précédent point
}

Unit UTOMServices ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,DBCtrls,HDB,
{$ELSE}
     eFiche,eFichList,UtileAGL,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOM, UTob ,HTB97,EntPaie,PGOutils,PGOutils2,LookUp,Paramsoc;

Type
  TOM_SERVICES = Class (TOM)
    procedure OnLoadRecord               ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord             ; override ;
    procedure OnChangeField (F : TField); override ;
    procedure OnArgument ( S: String )   ; override ;
    Private
    CodeEntre:String;
    AncRespVar, AncRespFor : String; //PT7
    //LibResp,LibRespVar,LibRespAbs,LibRespNdf:THLabel;
    //LibSecr,LibSecrAbs,LibSecrVar,LibSecrNdf:THLabel;
    //LibAdj,LibAdjAbs,LibAdjNdf,LibAdjVar:THLabel;
    Procedure BDupliquerClick(Sender:TObject);
    Procedure Libelles(Champ,Lib:String;CodeNum:Boolean);
    Procedure ClickServiceSup(Sender:TObject);
    Procedure ValeurDefaut;
	Function  BourreChaine ( St : String ; LgCompte : integer ) : string ; //PT8
    procedure ChangeNom (Sender : TObject); //PT8
    procedure RefreshNomService; //PT8
    procedure SalarieElipsisClick (Sender : TObject); //PT8
    end ;

Implementation

Uses PGOutilsFormation, StrUtils, PGHierarchie;

Const MASK_CHAMP = '9999999999';   //PT6
Const TTE_SALINT = 'PGSALARIEINT'; //PT6
Const LG_DOSSIER = 9;              //PT9

procedure TOM_SERVICES.OnLoadRecord;
Var PosSep : Integer;
begin
  Inherited;
  	CodeEntre:=GetField('PGS_CODESERVICE');
  	
	//PT8 - Début
	If PGBundleHierarchie Then
	Begin
		PosSep := Pos (SEP_CHAINE_SERVICE, GetField('PGS_NOMSERVICE')); //PT9
		
		// Eclatement du n° dossier et nom
		If (PosSep > 0) Then //(Length(GetField('PGS_NOMSERVICE'))>LG_DOSSIER) Then
		Begin
			SetControlText('NODOSSIER', LeftStr(GetField('PGS_NOMSERVICE'), PosSep-1));
        	SetControlText('NOMSERVICE', Copy(GetField('PGS_NOMSERVICE'),PosSep+1,Length(GetField('PGS_NOMSERVICE'))));
        End
        Else
        	// Pour la reprise de données existantes sans n° de dossier
        	SetControlText('NOMSERVICE', GetField('PGS_NOMSERVICE'));
	End
	Else
		SetControlText('NOMSERVICE', GetField('PGS_NOMSERVICE'));
	
	
	If PGBundleHierarchie And (Not PGDroitMultiForm) Then
	Begin
		If (DS.State in [dsInsert]) Then SetControlText('NODOSSIER', V_PGI.NoDossier);
		SetControlEnabled('NODOSSIER', False);
	End;
	//PT8 - Fin
	
	//PT7 - Début
	AncRespVar:=GetField('PGS_RESPONSVAR');
	AncRespFor:=GetField('PGS_RESPONSFOR');
	//PT7 - Fin
//NiveauEntre:=Getfield('PGS_HIERARCHIE');
//If GetField('PGS_HIERARCHIE')<>0 Then
//   begin
//   Niveau:=ColleZeroDevant(GetField('PGS_HIERARCHIE'),3);
//   SetControlText('NIVEAU',Niveau);
//   Lab:=THLabel(GetControl('LIBNIVEAU'));
//   If Lab<>Nil Then Lab.Caption:=RechDom('PGNIVEAUSERVICE',Niveau,False);
//   end;
end;

procedure TOM_SERVICES.OnDeleteRecord;
Var Service,ServiceInf,Mes:String;
    TobSer,T:Tob;
    Q:TQuery;
begin
  Inherited;
	Mes:='';
	
	Service := GetField('PGS_CODESERVICE');
	Q:=OpenSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SERVICESUP="'+Service+'" ORDER BY PGS_HIERARCHIE',True);
	TobSer:=TOB.Create('SERVICES',NIL,-1);
	TobSer.LoadDetailDB('SERVICES','','',Q,False);
	Ferme(Q);
	T:=TobSer.FindFirst([''],[''],False);
	While T<>Nil Do
	begin
		ServiceInf := T.GetValue('PGS_CODESERVICE');
		Mes:=Mes+'#13#10 - '+RechDom('PGSERVICE',ServiceInf,False);
		T:=TobSer.FindNext([''],[''],False);
	end;
	If Mes <>'' Then
	begin
		LastError:=1;
		PgiBox('Suppression impossible car le(s) service(s) suivant(s) dépende(nt) de ce service :'+Mes,'Suppression d''un service');
	end;
end;

procedure TOM_SERVICES.OnAfterUpdateRecord;
var RespAbs,RespNdf,RespVar,Service,RespFor:String;
    SecrAbs,SecrVar,SecrNdf:String;
begin
  Inherited;
	RespAbs:=GetField('PGS_RESPONSABS');
	RespNdf:=GetField('PGS_RESPONSNDF');
	RespVar:=GetField('PGS_RESPONSVAR');
	RespFor:=GetField('PGS_RESPONSFOR'); //PT3
	Service:=GetField('PGS_CODESERVICE');
	SecrAbs:=GetField('PGS_SECRETAIREABS');
	SecrNdf:=GetField('PGS_SECRETAIRENDF');
	SecrVar:=GetField('PGS_SECRETAIREVAR');
	
	// Mise à jour de la table contenant les affectations salariés
	ExecuteSQL('UPDATE DEPORTSAL SET PSE_RESPONSABS="'+RespAbs+'", PSE_RESPONSNDF="'+RespNdf+'", PSE_RESPONSFOR="'+RespFor+'"'+
	', PSE_RESPONSVAR="'+RespVar+'", PSE_ASSISTABS="'+SecrAbs+'", PSE_ASSISTNDF="'+SecrNdf+'", PSE_ASSISTVAR="'+SecrVar+'"'+ //PT1
	' WHERE PSE_CODESERVICE="'+Service+'"');
	
	AvertirTable('PGSERVICE');
	
	//PT7 - Début
	// Mise à jour des tables qui contiennent les références des responsables
	If RespVar <> AncRespVar Then MajResponsable(Service, 'VAR', RespVar);
	If RespFor <> AncRespFor Then MajResponsable(Service, 'FOR', RespFor);

    AncRespVar := RespVar;
    AncRespFor := RespFor;
	//PT7 - Fin
end;


procedure TOM_SERVICES.OnArgument ( S: String );
var
   (*Edit,*)Edit1,Edit2,Edit3,Edit4,Edit5,Edit6,Edit7,Edit8,Edit9,
   Edit10,Edit11,Edit12,Edit13,Edit14,Edit15,ServiceSup:{$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF};
   BDupliquer    :TToolBarButton97;
   EditT : THEdit;
begin
  Inherited;
	// DEBUT PT2
	SetControlVisible('PVAR', VH_Paie.PgeAbsences);
	SetControlVisible('PFOR', VH_Paie.PgSeriaFormation);
	// FIN PT2
	
	BDupliquer:=TToolBarButton97(GetControl('BDUPLIQUER'));
	If BDupliquer<>Nil Then BDupliquer.OnClick:=BDupliquerClick;
	
//	Edit	  :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_CODESERVICE'));
	ServiceSup:={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_SERVICESUP'));
	//if (VH_PAIE.PgTypeNumSal='NUM') then //PT6 - Récupération des champs même si matricules alpha
	//begin
	Edit1     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_RESPSERVICE'));
	Edit2     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_SECRETAIRE'));
	Edit3     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_ADJOINTSERVICE'));
	Edit4     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_RESPONSABS'));
	Edit5     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_RESPONSNDF'));
	Edit6     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_RESPONSVAR'));
	Edit7     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_SECRETAIREABS'));
	Edit8     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_SECRETAIRENDF'));
	Edit9     :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_SECRETAIREVAR'));
	Edit10    :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_ADJOINTABS'));
	Edit11    :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_ADJOINTNDF'));
	Edit12    :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_ADJOINTVAR'));
	Edit13    :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_RESPONSFOR'));
	Edit14    :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_SECRETAIREFOR'));
	Edit15    :={$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(GetControl('PGS_ADJOINTFOR'));
		
	{$IFNDEF EAGLCLIENT}
	if (VH_PAIE.PgTypeNumSal='NUM') then
	begin
		if Edit1  <> nil then Edit1.EditMask  := MASK_CHAMP;
		if Edit2  <> nil then Edit2.EditMask  := MASK_CHAMP;
		if Edit3  <> nil then Edit3.EditMask  := MASK_CHAMP;
		if Edit4  <> nil then Edit4.EditMask  := MASK_CHAMP;
		if Edit5  <> nil then Edit5.EditMask  := MASK_CHAMP;
		if Edit6  <> nil then Edit6.EditMask  := MASK_CHAMP;
		if Edit7  <> nil then Edit7.EditMask  := MASK_CHAMP;
		if Edit8  <> nil then Edit8.EditMask  := MASK_CHAMP;
		if Edit9  <> nil then Edit9.EditMask  := MASK_CHAMP;
		if Edit10 <> nil then Edit10.EditMask := MASK_CHAMP;
		if Edit11 <> nil then Edit11.EditMask := MASK_CHAMP;
		if Edit12 <> nil then Edit12.EditMask := MASK_CHAMP;
		if Edit13 <> nil then Edit13.EditMask := MASK_CHAMP;
		if Edit14 <> nil then Edit14.EditMask := MASK_CHAMP;
		if Edit15 <> nil then Edit15.EditMask := MASK_CHAMP;
	End;
	{$ENDIF}
	
	//PT6 - La modif de la tablette doit être faite même si gestion matricules non numérique
	if (GetParamSocSecur('SO_PGINTERVENANTEXT',False)) Or (GetParamSocSecur('SO_IFDEFCEGID',FALSE)) then //PT5
	begin
		if Edit1  <> nil then Edit1.Datatype  := TTE_SALINT;
		if Edit2  <> nil then Edit2.Datatype  := TTE_SALINT;
		if Edit3  <> nil then Edit3.Datatype  := TTE_SALINT;
		if Edit4  <> nil then Edit4.Datatype  := TTE_SALINT;
		if Edit5  <> nil then Edit5.Datatype  := TTE_SALINT;
		if Edit6  <> nil then Edit6.Datatype  := TTE_SALINT;
		if Edit7  <> nil then Edit7.Datatype  := TTE_SALINT;
		if Edit8  <> nil then Edit8.Datatype  := TTE_SALINT;
		if Edit9  <> nil then Edit9.Datatype  := TTE_SALINT;
		if Edit10 <> nil then Edit10.Datatype := TTE_SALINT;
		if Edit11 <> nil then Edit11.Datatype := TTE_SALINT;
		if Edit12 <> nil then Edit12.Datatype := TTE_SALINT;
		if Edit13 <> nil then Edit13.Datatype := TTE_SALINT;
		if Edit14 <> nil then Edit14.Datatype := TTE_SALINT;
		if Edit15 <> nil then Edit15.Datatype := TTE_SALINT;
	end;
	
	//PT8 - Début
	If PGBundleHierarchie Then
	Begin
		if Edit1  <> nil then Edit1.OnElipsisClick  := SalarieElipsisClick;
		if Edit2  <> nil then Edit2.OnElipsisClick  := SalarieElipsisClick;
		if Edit3  <> nil then Edit3.OnElipsisClick  := SalarieElipsisClick;
		if Edit4  <> nil then Edit4.OnElipsisClick  := SalarieElipsisClick;
		if Edit5  <> nil then Edit5.OnElipsisClick  := SalarieElipsisClick;
		if Edit6  <> nil then Edit6.OnElipsisClick  := SalarieElipsisClick;
		if Edit7  <> nil then Edit7.OnElipsisClick  := SalarieElipsisClick;
		if Edit8  <> nil then Edit8.OnElipsisClick  := SalarieElipsisClick;
		if Edit9  <> nil then Edit9.OnElipsisClick  := SalarieElipsisClick;
		if Edit10 <> nil then Edit10.OnElipsisClick := SalarieElipsisClick;
		if Edit11 <> nil then Edit11.OnElipsisClick := SalarieElipsisClick;
		if Edit12 <> nil then Edit12.OnElipsisClick := SalarieElipsisClick;
		if Edit13 <> nil then Edit13.OnElipsisClick := SalarieElipsisClick;
		if Edit14 <> nil then Edit14.OnElipsisClick := SalarieElipsisClick;
		if Edit15 <> nil then Edit15.OnElipsisClick := SalarieElipsisClick;
	End;
	//PT8 - Fin
	
	//PT6 - Mise en commentaires
	(*end;
	{$ELSE}
	// GC_20071004_GM_GC14731_DEBUT Erreur sur nom de champs
	//ServiceSup:=THEdit(GetControl('PSE_SERVICESUP'));
	ServiceSup:=THEdit(GetControl('PGS_SERVICESUP'));
	// GC_20071004_GM_GC14731_FIN
	Edit:=THEdit(GetControl('PGS_CODESERVICE'));
	if (VH_PAIE.PgTypeNumSal='NUM') then
	begin

		Edit1:=THEdit(GetControl('PGS_RESPSERVICE'));
		Edit2:=THEdit(GetControl('PGS_SECRETAIRE'));
		Edit3:=THEdit(GetControl('PGS_ADJOINTSERVICE'));
		Edit4:=THEdit(GetControl('PGS_RESPONSABS'));
		Edit5:=THEdit(GetControl('PGS_RESPONSNDF'));
		Edit6:=THEdit(GetControl('PGS_RESPONSVAR'));
		Edit7:=THEdit(GetControl('PGS_SECRETAIREABS'));
		Edit8:=THEdit(GetControl('PGS_SECRETAIRENDF'));
		Edit9:=THEdit(GetControl('PGS_SECRETAIREVAR'));
		Edit10:=THEdit(GetControl('PGS_ADJOINTABS'));
		Edit11:=THEdit(GetControl('PGS_ADJOINTNDF'));
		Edit12:=THEdit(GetControl('PGS_ADJOINTVAR'));
		Edit13:=THEdit(GetControl('PGS_RESPONSFOR'));
		Edit14:=THEdit(GetControl('PGS_SECRETAIREFOR'));
		Edit15:=THEdit(GetControl('PGS_ADJOINTFOR'));
		{ ******** PT4 Mis en commentaire
		if Edit1<>nil then Edit1.EditMask:='9999999999';
		if Edit2<>nil then Edit2.EditMask:='9999999999';
		if Edit3<>nil then Edit3.EditMask:='9999999999';
		if Edit4<>nil then Edit4.EditMask:='9999999999';
		if Edit5<>nil then Edit5.EditMask:='9999999999';
		if Edit6<>nil then Edit6.EditMask:='9999999999';
		if Edit7<>nil then Edit7.EditMask:='9999999999';
		if Edit8<>nil then Edit8.EditMask:='9999999999';
		if Edit9<>nil then Edit9.EditMask:='9999999999';
		if Edit10<>nil then Edit10.EditMask:='9999999999';
		if Edit11<>nil then Edit11.EditMask:='9999999999';
		if Edit12<>nil then Edit12.EditMask:='9999999999';
		if Edit13<>nil then Edit13.EditMask:='9999999999';
		if Edit14<>nil then Edit14.EditMask:='9999999999';
		if Edit15<>nil then Edit15.EditMask:='9999999999';    }
		if GetParamSocSecur('SO_IFDEFCEGID',FALSE) then
		begin
			if Edit1<>nil then Edit1.Datatype :='PGSALARIEINT';
			if Edit2<>nil then Edit2.Datatype :='PGSALARIEINT';
			if Edit3<>nil then Edit3.Datatype :='PGSALARIEINT';
			if Edit4<>nil then Edit4.Datatype :='PGSALARIEINT';
			if Edit5<>nil then Edit5.Datatype :='PGSALARIEINT';
			if Edit6<>nil then Edit6.Datatype :='PGSALARIEINT';
			if Edit7<>nil then Edit7.Datatype :='PGSALARIEINT';
			if Edit8<>nil then Edit8.Datatype :='PGSALARIEINT';
			if Edit9<>nil then Edit9.Datatype :='PGSALARIEINT';
			if Edit10<>nil then Edit10.Datatype :='PGSALARIEINT';
			if Edit11<>nil then Edit11.Datatype :='PGSALARIEINT';
			if Edit12<>nil then Edit12.Datatype :='PGSALARIEINT';
			if Edit13<>nil then Edit13.Datatype :='PGSALARIEINT';
			if Edit14<>nil then Edit14.Datatype :='PGSALARIEINT';
			if Edit15<>nil then Edit15.Datatype :='PGSALARIEINT';
		end;
	end;
	{$ENDIF}*)
		
//	if Edit <> Nil Then Edit.EditMask  := '999'; //PT9

	If ServiceSup <> Nil Then ServiceSup.OnElipsisClick:=ClickServiceSup;
	
	//PT8 - Début
    EditT := THEdit(GetControl('NOMSERVICE'));
    If EditT <> Nil Then EditT.OnChange  := ChangeNom;

	If PGBundleHierarchie Then
	Begin
        //PT9
        SetControlVisible('PGS_ETABLISSEMENT', False);
        SetControlVisible('TPGS_ETABLISSEMENT', False);

		SetControlVisible('TNODOSSIER', True);
		SetControlVisible('NODOSSIER',  True);
		If Not PGDroitMultiForm Then
		Begin
			SetControlText('NODOSSIER', V_PGI.NoDossier); //PT9
			SetControlEnabled('NODOSSIER', False);
		End
        Else If V_PGI.ModePCL='1' Then 
        	SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT9

		// Réduction de la taille du libellé
//		EditT := THEdit(GetControl('NOMSERVICE'));
//		If EditT <> Nil Then EditT.MaxLength := EditT.MaxLength - LG_DOSSIER;

		EditT := THEdit(GetControl('NODOSSIER'));
		If EditT <> Nil Then EditT.OnChange  := ChangeNom;
	End;
	//PT8 - Fin
end;

procedure TOM_SERVICES.OnChangeField(F: TField);
var CodeService:String;
    CodeNum:Boolean;
    Lab:THLabel;
begin
	if (VH_PAIE.PgTypeNumSal='NUM')  then CodeNum:=True
	Else CodeNum:=False;
	If F.FieldName=('PGS_CODESERVICE') Then
	begin
		CodeService:=GetField('PGS_CODESERVICE'); 
		If CodeService=CodeEntre Then Exit;
		If CodeService='' Then Exit;

		//PT9
        //CodeService := ColleZeroDevant(StrToInt(trim(CodeService)),3);
        CodeService := BourreChaine(CodeService,3);
		
		if CodeService <> GetField('PGS_CODESERVICE') then
            SetField('PGS_CODESERVICE',CodeService);
		
		If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_CODESERVICE="'+GetField('PGS_CODESERVICE')+'"') Then
		begin
			PGIBox('Ce code est déja utilisé, veuillez en saisir un autre',TFFiche(Ecran).Caption);
			SetFocusControl('PGS_CODESERVICE');
		end;
	end;
    //PT8 - Début
	(*If F.FieldName=('PGS_NOMSERVICE') Then
	begin
		Ecran.caption:='Saisie service : ';
		If PGBundleHierarchie Then //PT8
			Ecran.Caption := Ecran.Caption + Copy(GetField('PGS_NOMSERVICE'),9,Length(GetField('PGS_NOMSERVICE')))
		Else
			Ecran.Caption := Ecran.Caption + GetField('PGS_NOMSERVICE');
		TFFiche(Ecran).DisabledMajCaption := True;
	end;*)
	//PT8 - Fin

	If F.FieldName=('PGS_SERVICESUP') Then
	begin
		Lab:=THLabel(GetControl('LIBSERVICESUP'));
		If Lab=Nil Then Exit;
		If GetField('PGS_SERVICESUP')<>'' Then Lab.Caption:=RechDom('PGSERVICE',GetField('PGS_SERVICESUP'),False)
		Else Lab.Caption:='';
	end;

	if F.FieldName=('PGS_RESPSERVICE')    Then Libelles('PGS_RESPSERVICE','LRESP',CodeNum);
	if F.FieldName=('PGS_RESPONSVAR')     Then Libelles('PGS_RESPONSVAR','LRESPVAR',CodeNum);
	if F.FieldName=('PGS_RESPONSABS')     Then Libelles('PGS_RESPONSABS','LRESPABS',CodeNum);
	if F.FieldName=('PGS_RESPONSNDF')     Then Libelles('PGS_RESPONSNDF','LRESPNDF',CodeNum);
	if F.FieldName=('PGS_RESPONSFOR')     Then Libelles('PGS_RESPONSFOR','LRESPFOR',CodeNum);
	if F.FieldName=('PGS_SECRETAIRE')     Then Libelles('PGS_SECRETAIRE','LSECRETAIRE',CodeNum);
	if F.FieldName=('PGS_SECRETAIREABS')  Then Libelles('PGS_SECRETAIREABS','LSECRETAIREABS',CodeNum);
	if F.FieldName=('PGS_SECRETAIRENDF')  Then Libelles('PGS_SECRETAIRENDF','LSECRETAIRENDF',CodeNum);
	if F.FieldName=('PGS_SECRETAIREVAR')  Then Libelles('PGS_SECRETAIREVAR','LSECRETAIREVAR',CodeNum);
	if F.FieldName=('PGS_SECRETAIREFOR')  Then Libelles('PGS_SECRETAIREFOR','LSECRETAIREFOR',CodeNum);
	if F.FieldName=('PGS_ADJOINTSERVICE') Then Libelles('PGS_ADJOINTSERVICE','LADJOINT',CodeNum);
	if F.FieldName=('PGS_ADJOINTABS')     Then Libelles('PGS_ADJOINTABS','LADJOINTABS',CodeNum);
	if F.FieldName=('PGS_ADJOINTNDF')     Then Libelles('PGS_ADJOINTNDF','LADJOINTNDF',CodeNum);
	if F.FieldName=('PGS_ADJOINTVAR')     Then Libelles('PGS_ADJOINTVAR','LADJOINTVAR',CodeNum);
	if F.FieldName=('PGS_ADJOINTFOR')     Then Libelles('PGS_ADJOINTFOR','LADJOINTFOR',CodeNum);
end;

procedure TOM_SERVICES.OnUpdateRecord ;
var Etab,Service : String;
    Mes,SNiveauSup,Titre:String;
    MesAdj,MesResp,MesSecr:String;
begin
  Inherited ;
	Titre:='Saisie des services';
	Etab:=GetControlText('PGS_ETABLISSEMENT');
	Service:=GetControlText('PGS_SERVICESUP');
	SNiveauSup:=GetField('PGS_SERVICESUP');
	Mes:='';
	// If Etab<>EtabSalarie Then Mes:='#13#10 - Le salarié doit être dans le même établissement que le service';
	//SetField('PGS_HIERARCHIE',GetControlText('NIVEAU'));
	//If Service='' Then
	//   begin
	//   If Niveau<>1 Then
	//     begin
	//    Mes:=Mes+'#13#10 - le service doit être de niveau 1 si il n"a pas de service supérieur';
	//   SetField('PGS_HIERARCHIE',1);
	//  SetFocusControl('PGS_HIERARCHIE');
	//  end;
	// end
	//Else
	// begin
	//   Q:=OpenSQL('SELECT PGS_ETABLISSEMENT FROM SERVICES WHERE PGS_CODESERVICE="'+Service+'"',True);
	//   EtabSup:=Q.FindField('PGS_ETABLISSEMENT').AsString;
	//   If EtabSup='' Then Mes:=Mes+' Le service supérieur n''existe pas';
	//   Ferme(Q);
	//   If Etab<>EtabSup Then
	//      begin
	//      QEtabSup:=OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="'+EtabSup+'"',True);
	//      LibEtabSup:=QEtabSup.FindField('ET_LIBELLE').AsString;
	//      Ferme(QEtabSup);
	//      Mes:=Mes+'#13#10 - le service doit être dans le même établissement que le service supérieur qui appartient à '+LibEtabSup;
	//      SetField('PGS_ETABLISSEMENT',EtabSup);
	//      SetFocusControl('PGS_ETABLISSEMENT');
	//      end;
	//   Q1:=OpenSQL('SELECT PGS_HIERARCHIE FROM SERVICES WHERE PGS_CODESERVICE="'+SNiveauSup+'"',True);
	//   NiveauSup:=Q1.FindField('PGS_HIERARCHIE').AsInteger;
	//   Ferme(Q1);
	//   If NiveauSup>=Niveau Then
	//      begin
	//      Mes:=Mes+'#13#10 - le niveau du service doit être supérieur au niveau du service supérieur qui est '+ IntToStr(NiveauSup);
	//      NiveauSup:=NiveauSup+1;
	//      SetField('PGS_HIERARCHIE',NiveauSup);
	//      SetFocusControl('PGS_HIERARCHIE');
	//      end;
	// end;
	If Mes<>'' Then
	begin
		LastError:=1;
		PgiBox('La ou les informations suivantes sont erronnées : '+Mes,Titre);
	end
	Else
	begin
		If GetField('PGS_NOMSERVICE')='' Then
		begin
			Mes:=Mes+'#13#10 - le nom du service';
			SetFocusControl('PGS_NOMSERVCIE');
		end;
		If GetField('PGS_RESPSERVICE')='' Then
		begin
			Mes:=Mes+'#13#10 - le responsable du service';
			SetFocusControl('PGS_RESPSERVICE');
		end;
		//PT9 - Début

        // Cas particulier des utilisateurs ayant accès à plusieurs dossiers mais n'étant pas superviseur
        // Dans ce cas, le n° de dossier est obligatoire
		If PGBundleHierarchie And (Not JaiLeDroitTag(46530)) And (GetControlText('NODOSSIER') = '') Then
		Begin
			Mes:=Mes+'#13#10 - le dossier de rattachement';
			SetFocusControl('NODOSSIER');
		End;

 	    //PT9 - Fin
		If Mes<>'' Then
		begin
			LastError:=1;
			PgiBox('Vous devez renseigner :'+Mes,Titre);
            Exit;
		end;
	end;
	Mes:='';
	If GetField('PGS_RESPSERVICE')<>'' Then
	begin
		If GetField('PGS_RESPONSABS')='' Then MesResp:=MesResp+'#13#10 - Responsable absences';
		If GetField('PGS_RESPONSNDF')='' Then MesResp:=MesResp+'#13#10 - Responsable notes de frais';
		If GetField('PGS_RESPONSVAR')='' Then MesResp:=MesResp+'#13#10 - Responsable éléments variables';
		If GetField('PGS_RESPONSFOR')='' Then MesResp:=MesResp+'#13#10 - Responsable formation';
	end;
	If GetField('PGS_ADJOINTSERVICE')<>'' Then
	begin
		If GetField('PGS_ADJOINTSABS')='' Then MesAdj:=MesAdj+'#13#10 - Adjoint absences';
		If GetField('PGS_ADJOINTNDF')='' Then MesAdj:=MesAdj+'#13#10 - Adjoint notes de frais';
		If GetField('PGS_ADJOINTVAR')='' Then MesAdj:=MesAdj+'#13#10 - Adjoint éléments variables';
		If GetField('PGS_ADJOINTFOR')='' Then MesAdj:=MesAdj+'#13#10 - Adjoint formation';
	end;
	If GetField('PGS_SECRETAIRE')<>'' Then
	begin
		If GetField('PGS_RESPONSABS')='' Then MesSecr:=MesSecr+'#13#10 - Secrétaire absences';
		If GetField('PGS_RESPONSNDF')='' Then MesSecr:=MesSecr+'#13#10 - Secrétaire notes de frais';
		If GetField('PGS_RESPONSVAR')='' Then MesSecr:=MesSecr+'#13#10 - Secrétaire éléments variables';
		If GetField('PGS_RESPONSFOR')='' Then MesSecr:=MesSecr+'#13#10 - Secrétaire formation';
	end;
	If (MesResp<>'') or (MesAdj<>'') or (MesSecr<>'') Then
	begin
		Case PGIAskCancel('Les informations suivantes ne sont pas renseignées :'+MesResp+MesAdj+MesSecr+'#13#10 voulez-vous les renseigner par défaut (reprise des données générales du service) ?',Titre) of
		mrYes : ValeurDefaut;
		mrCancel : exit;
	end;
	
    If Mes = '' Then RefreshNomService; //PT8
end;
end;

Procedure TOM_SERVICES.BDupliquerClick(Sender:TObject);
Var Resp,Adjoint,Secretaire:String;
begin
	ForceUpdate;
	Resp:=GetField('PGS_RESPSERVICE');
	Adjoint:=GetField('PGS_ADJOINTSERVICE');
	Secretaire:=GetField('PGS_SECRETAIRE');
	SetField('PGS_RESPONSABS',Resp);
	SetField('PGS_RESPONSNDF',Resp);
	SetField('PGS_RESPONSVAR',Resp);
	SetField('PGS_RESPONSFOR',Resp);
	SetField('PGS_ADJOINTABS',Adjoint);
	SetField('PGS_ADJOINTNDF',Adjoint);
	SetField('PGS_ADJOINTVAR',Adjoint);
	SetField('PGS_ADJOINTFOR',Adjoint);
	SetField('PGS_SECRETAIREABS',Secretaire);
	SetField('PGS_SECRETAIRENDF',Secretaire);
	SetField('PGS_SECRETAIREVAR',Secretaire);
	SetField('PGS_SECRETAIREFOR',Secretaire);
end;

Procedure TOM_SERVICES.Libelles(Champ,Lib:String;CodeNum:Boolean);
var TLabel:THLabel;
    Resp:String;
begin
	Resp:=Trim(GetField(Champ));
	If (CodeNum=True) and (resp<>'') Then
	begin
		Resp := ColleZeroDevant(StrToInt(trim(Resp)),10);
		If Resp <> (GetField(Champ)) then SetField(Champ,Resp);
	end;
	TLabel:=THLabel(GetControl(Lib));
	If GetField(Champ)<>'' Then
	begin
		if GetParamSocSecur('SO_IFDEFCEGID',False) then
			TLabel.Caption:=RechDom('PGSALARIEINT',GetField(Champ),False)
		else If PGBundleHierarchie Then //PT6
			TLabel.Caption:=RechDom(TTE_SALINT,GetField(Champ),False)
		Else
			TLabel.Caption:=RechDom('PGSALARIE',GetField(Champ),False);
	end
	Else TLabel.Caption:='';
end;

Procedure TOM_SERVICES.ClickServiceSup(Sender:TObject);
var St:String;
    Niveau:Integer;
    Q:TQuery;
begin
	If Sender=Nil Then Exit;
	Q:=OpenSQL('SELECT PHO_NIVEAUH FROM HIERARCHIE WHERE PHO_HIERARCHIE="'+GetField('PGS_HIERARCHIE')+'"',True);
	if NOT Q.EOF then Niveau:=Q.FindField('PHO_NIVEAUH').AsInteger // PORTAGECWAS
	else Niveau := 0;
	Ferme(Q);
	St:=' SELECT PGS_CODESERVICE,PGS_NOMSERVICE,PHO_LIBELLE,PHO_NIVEAUH,PSA_LIBELLE FROM SERVICES'+
	' LEFT JOIN SALARIES ON PSA_SALARIE=PGS_RESPSERVICE LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE'+
	' WHERE PHO_NIVEAUH<'+IntToStr(Niveau)+' ORDER BY PHO_NIVEAUH';   // DB2
	// DEBUT CCMX-CEGID FQ N° 14706
	if Not ExisteSQL(St) then
	begin
		PgiBox('Il n''existe aucun service de niveau supèrieur ',TFFiche(Ecran).Caption);
		Exit;
	end;
	// FIN CCMX-CEGID FQ N° 14706
	{$IFNdef GCGC}  // GC_20071004_GM_GC14731
	{$IFNDEF EAGLCLIENT}
	LookupList(THDBEdit(Sender),'Liste des services','SERVICES','PGS_CODESERVICE','PGS_NOMSERVICE,PHO_NIVEAUH,PHO_LIBELLE,PSA_LIBELLE','','', True,-1,St);
	{$ELSE}
	LookupList(THEdit(Sender),'Liste des services','SERVICES','PGS_CODESERVICE','PGS_NOMSERVICE,PHO_NIVEAUH,PHO_LIBELLE,PSA_LIBELLE','','', True,-1,St);
	{$ENDIF}
	// GC_20071004_GM_GC14731_DEBUT
	{$ELSE}
	St:=' SELECT PGS_CODESERVICE,PGS_NOMSERVICE FROM SERVICES'+
	' LEFT JOIN SALARIES ON PSA_SALARIE=PGS_RESPSERVICE LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE'+
	' WHERE PHO_NIVEAUH<'+IntToStr(Niveau);   // DB2

	{$IFNDEF EAGLCLIENT}
	LookupList(THDBEdit(Sender),'Liste des services','SERVICES','PGS_CODESERVICE','PGS_NOMSERVICE','','PHO_NIVEAUH', True,-1,St,tlLocate);
	{$ELSE}
	LookupList(THEdit(Sender),'Liste des services','SERVICES','PGS_CODESERVICE','PGS_NOMSERVICE','','PHO_NIVEAUH', True,-1,St,tlDefault);
	{$ENDIF}
	{$ENDIF GCGC}
	// GC_20071004_GM_GC14731_FIN
end;

Procedure TOM_SERVICES.ValeurDefaut;
begin
	If GetField('PGS_RESPONSABS')='' Then SetField('PGS_RESPONSABS',GetField('PGS_RESPSERVICE'));
	If GetField('PGS_RESPONSNDF')='' Then SetField('PGS_RESPONSNDF',GetField('PGS_RESPSERVICE'));
	If GetField('PGS_RESPONSVAR')='' Then SetField('PGS_RESPONSVAR',GetField('PGS_RESPSERVICE'));
	If GetField('PGS_RESPONSFOR')='' Then SetField('PGS_RESPONSFOR',GetField('PGS_RESPSERVICE'));

	If GetField('PGS_ADJOINTABS')='' Then SetField('PGS_ADJOINTABS',GetField('PGS_ADJOINTSERVICE'));
	If GetField('PGS_ADJOINTNDF')='' Then SetField('PGS_ADJOINTNDF',GetField('PGS_ADJOINTSERVICE'));
	If GetField('PGS_ADJOINTVAR')='' Then SetField('PGS_ADJOINTVAR',GetField('PGS_ADJOINTSERVICE'));
	If GetField('PGS_ADJOINTFOR')='' Then SetField('PGS_ADJOINTFOR',GetField('PGS_ADJOINTSERVICE'));

	If GetField('PGS_SECRETAIREABS')='' Then SetField('PGS_SECRETAIREABS',GetField('PGS_SECRETAIRE'));
	If GetField('PGS_SECRETAIRENDF')='' Then SetField('PGS_SECRETAIRENDF',GetField('PGS_SECRETAIRE'));
	If GetField('PGS_SECRETAIREVAR')='' Then SetField('PGS_SECRETAIREVAR',GetField('PGS_SECRETAIRE'));
	If GetField('PGS_SECRETAIREFOR')='' Then SetField('PGS_SECRETAIREFOR',GetField('PGS_SECRETAIRE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/03/2008 / PT8
Modifié le ... :   /  /
Description .. : Evènement lancé sur le changement de nom de service
Mots clefs ... :
*****************************************************************}
procedure TOM_SERVICES.ChangeNom (Sender : TObject);
Begin
    Ecran.caption:='Saisie service : ' + GetControlText('NOMSERVICE');
	TFFiche(Ecran).DisabledMajCaption := True;
    RefreshNomService;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/03/2008 / PT8
Modifié le ... :   /  /    
Description .. : Fonction de remplissage d'une chaîne
Mots clefs ... :
*****************************************************************}
Function TOM_SERVICES.BourreChaine ( St : String ; LgCompte : integer ) : string ;
var ll,i : Integer ;
    Bourre  : Char ;
begin
	Bourre:='0';
	Result:=St ; ll:=Length(Result) ;
	If ll<LgCompte then for i:=ll+1 to LgCompte do Result:=Result+Bourre;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/03/2008 / PT8
Modifié le ... :   /  /
Description .. : Rafraîchit le nom du service en concaténant Dossier et libellé
Mots clefs ... :
*****************************************************************}
procedure TOM_SERVICES.RefreshNomService;
Begin
	If PGBundleHierarchie  And (GetControlText('NODOSSIER') <> '') Then
		SetField('PGS_NOMSERVICE', GetControlText('NODOSSIER')+SEP_CHAINE_SERVICE+GetControlText('NOMSERVICE'))
    Else
		SetField('PGS_NOMSERVICE', GetControlText('NOMSERVICE'));
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/03/2008 / PT8
Modifié le ... :   /  /
Description .. : Clic sur Elipsis salarié : retriction au n° de dossier
Mots clefs ... :
*****************************************************************}
procedure TOM_SERVICES.SalarieElipsisClick (Sender : TObject);
var StWhere : String;
Begin
	StWhere := '';
    
	If (Trim(GetControlText('NODOSSIER')) <> '') And (Not PGDroitMultiForm) Then
		StWhere := 'PSI_NODOSSIER="'+Trim(GetControlText('NODOSSIER'))+'"';
    ElipsisSalarieMultiDos (Sender, StWhere);
end;

Initialization
  registerclasses ( [ TOM_SERVICES ] ) ;
end.


