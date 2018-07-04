{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PG_MUL_RESPSERVICES ()
Mots clefs ... : TOF;PG_MUL_RESPSERVICES
*****************************************************************
PT1 : 07/01/2003 JL V_42 Modif requete champ niveauH au lieu de hierarchie
PT2 : 19/09/2007 FL V_80 En multidossiers, ne pas afficher les combos établissement + affiche données table interimaires si besoin
PT3 : 02/04/2008 FL V_80 Adaptation des requêtes pour partage de la hiérarchie
}
Unit UTOFResponsableServices;

Interface

uses     UTOF,
{$IFNDEF EAGLCLIENT}
         {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,Mul,
{$ELSE}
       eMul,UTob,
{$ENDIF}
         Hctrls,HEnt1,Classes,sysutils,PgOutils2,EntPaie,LookUp,StdCtrls;

Type
  TOF_MUL_AFFICHERESP = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Resp:THEdit;
    TypeRech:THValComboBox;
    procedure ExitEdit(Sender: TObject);
    procedure SalElipsisClick(Sender:TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;

Implementation
Uses PGOutilsFormation,ParamSoc,PGHierarchie;

procedure TOF_MUL_AFFICHERESP.OnClickSalarieSortie(Sender: TObject);
begin
	SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
	SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_MUL_AFFICHERESP.OnLoad ;
var Responsable,Recherche,Service,Where,WhereResp:String;
    RespAbs,RespVar,RespNdf,Adj,AdjAbs,AdjVar,AdjNdf,Secr,SecrAbs,SecrVar,SecrNdf:String;
    Q:TQuery;
    NivInf,NivSup:THValComboBox;
    NiveauSup,NiveauInf:Integer;
    DateArret : TDateTime;
    StDateArret : String;
begin
	Inherited;
	NivInf:=THValComboBox(GetControl('NIVEAUINF'));
	NivSup:=THValComboBox(GetControl('NIVEAUSUP'));
	Recherche:=TypeRech.Value;
	Where:='';
	WhereResp:='';
	Responsable:=Resp.Text;
	RespAbs:=GetControlText('RESPONSABS');
	If GetControl('RESPONSNDF') <> Nil Then RespNdf:=GetControlText('RESPONSNDF') Else RespNdf := ''; //PT2
	RespVar:=GetControlText('RESPONSVAR');
	Adj:=GetControlText('ADJOINTSERVICE');
	AdjAbs:=GetControlText('ADJOINTABS');
	AdjNdf:=GetControlText('ADJOINTNDF');
	AdjVar:=GetControlText('ADJOINTVAR');
	Secr:=GetControlText('SECRETAIRE');
	SecrAbs:=GetControlText('SECRETAIREABS');
	SecrNdf:=GetControlText('SECRETAIRENDF');
	SecrVar:=GetControlText('SECRETAIREVAR');
	If Responsable<>'' Then WhereResp:=' WHERE PGS_RESPSERVICE="'+Responsable+'"';
	If RespAbs<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_RESPONSABS="'+RespAbs+'"'
		Else WhereResp:=WhereResp+' AND PGS_RESPONSABS="'+RespAbs+'"';
	end;
	If RespNdf<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_RESPONSNDF="'+RespNdf+'"'
		Else WhereResp:=WhereResp+' AND PGS_RESPONSNDF="'+RespNdf+'"';
	end;
	If RespVar<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_RESPONSVAR="'+RespVar+'"'
		Else WhereResp:=WhereResp+' AND PGS_RESPONSVAR="'+RespVar+'"';
	end;
	If Adj<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_ADJOINTSERVICE="'+Adj+'"'
		Else WhereResp:=WhereResp+' AND PGS_ADJOINTSERVICE="'+Adj+'"';
	end;
	If AdjAbs<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_ADJOINTABS="'+AdjAbs+'"'
		Else WhereResp:=WhereResp+' AND PGS_ADJOINTABS="'+AdjAbs+'"';
	end;
	If AdjNdf<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_ADJOINTNDF="'+AdjNdf+'"'
		Else WhereResp:=WhereResp+' AND PGS_ADJOINTNDF="'+AdjNdf+'"';
	end;
	If AdjVar<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_ADJOINTVAR="'+AdjVar+'"'
		Else WhereResp:=WhereResp+' AND PGS_ADJOINTVAR="'+AdjVar+'"';
	end;
	If Secr<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_SECRETAIRE="'+Secr+'"'
		Else WhereResp:=WhereResp+' AND PGS_SECRETAIRE="'+Secr+'"';
	end;
	If SecrAbs<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_SECRETAIREABS="'+SecrAbs+'"'
		Else WhereResp:=WhereResp+' AND PGS_SECRETAIREABS="'+SecrAbs+'"';
	end;
	If SecrVar<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_SECRETAIREVAR="'+SecrVar+'"'
		Else WhereResp:=WhereResp+' AND PGS_SECRETAIREVAR="'+SecrVar+'"';
	end;
	If SecrNdf<>'' Then
	begin
		If WhereResp='' Then WhereResp:=' WHERE PGS_SECRETAIRENDF="'+SecrNdf+'"'
		Else WhereResp:=WhereResp+' AND PGS_SECRETAIRENDF="'+SecrNdf+'"';
	end;
	If WhereResp<>'' Then
	begin
		NiveauInf:=0;
		NiveauSup:=0;
		Service:='';
		Q:=OpenSQL('SELECT PHO_NIVEAUH FROM HIERARCHIE WHERE PHO_HIERARCHIE="'+NivInf.Value+'"',True);
		if not Q.eof then NiveauInf:=Q.FindField('PHO_NIVEAUH').AsInteger;  // PortageCWAS
		Ferme(Q);
		Q:=OpenSQL('SELECT PHO_NIVEAUH FROM HIERARCHIE WHERE PHO_HIERARCHIE="'+NivSup.Value+'"',True);
		if not Q.eof then NiveauSup:=Q.FindField('PHO_NIVEAUH').AsInteger; // PortageCWAS
		Ferme(Q);
		Q:=OpenSQL('SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE '+WhereResp+' AND PHO_NIVEAUH >="'+IntToStr(NiveauInf)+'"'+
		' AND PHO_NIVEAUH<="'+InttoStr(NiveauSup)+'" order by PHO_HIERARCHIE',True);       //PT1
		if not Q.eof then Service:=Q.FindField('PGS_CODESERVICE').AsString;  // PortageCWAS
		Ferme(Q);
		If Recherche='dir' Then SetControlText('XX_WHERE','PSE_CODESERVICE="'+Service+'"');
		If Recherche='indir' Then
		begin
			Where:='PSE_CODESERVICE IN (SELECT PSO_CODESERVICE FROM SERVICEORDRE WHERE PSO_SERVICESUP="'+Service+'")';
			SetControlText('XX_WHERE',Where);
		end;
		If Recherche='tous' Then
		begin
			where:='PSE_CODESERVICE="'+Service+'" or PSE_CODESERVICE IN (SELECT PSO_CODESERVICE FROM SERVICEORDRE WHERE PSO_SERVICESUP="'+Service+'")';
			SetControlText('XX_WHERE',Where);
		end;
	end
	Else SetControlText('XX_WHERE','');
	Where := GetControlText('XX_WHERE');
	if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET')))then
	Begin
		DateArret:=StrtoDate(GetControlText('DATEARRET'));
		//PT2 - Début
		If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
		Begin
			StDateArret:='(PSI_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSI_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSI_DATESORTIE IS NULL) ';
			StDateArret:=StDateArret + ' AND PSI_DATEENTREE <="'+UsDateTime(DateArret)+'"';
		End
		Else
		Begin
			StDateArret:='(PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
			StDateArret:=StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
		End;
		//PT2 - Fin
		If Where <> '' then Where := WHere + ' AND '+StdateArret
		else Where := StDateArret;
	End;
	SetControlText('XX_WHERE',Where);
end ;

procedure TOF_MUL_AFFICHERESP.ExitEdit(Sender: TObject);
var edit : thedit;
begin
	edit:=THEdit(Sender);
	if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
	if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
	edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_MUL_AFFICHERESP.OnArgument (S : String ) ;
var THDefaut:THEdit;
    Min,Max:String;
    NivInf,NivSup:THValComboBox;
    Q:TQuery;
    Check : TCheckBox;
begin
	Inherited ;
	Q:=OpenSQL('SELECT PHO_HIERARCHIE FROM HIERARCHIE ORDER BY PHO_NIVEAUH',True);
	If Q<>Nil Then
	begin
		Q.First;
		Min:=Q.FindField('PHO_HIERARCHIE').AsString;
		Q.Last;
		Max:=Q.FindField('PHO_HIERARCHIE').AsString;
	end;
	Ferme(Q);
	NivInf:=THValComboBox(GetControl('NIVEAUINF'));
	NivSup:=THValComboBox(GetControl('NIVEAUSUP'));
	If NivInf<>Nil Then NivInf.Value:=Min;
	If NivSup<>Nil Then NivSup.Value:=Max;
	TypeRech:=THValComboBox(GetControl('TYPERECH'));
	TypeRech.Value:='Dir';
	Resp:=THEdit(GetControl('RESPSERVICE'));
	If Resp<>nil then
	begin
		Resp.OnElipsisClick:=SalElipsisClick;
		Resp.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('RESPONSNDF'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('RESPONSABS'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('RESPONSVAR'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('ADJOINTSERVICE'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('ADJOINTVAR'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('ADJOINTNDF'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('ADJOINTABS'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('SECRETAIRE'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('SECRETAIREABS'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('SECRETAIRENDF'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	THDefaut:=THEdit(GetControl('SECRETAIREVAR'));
	If THDefaut<>Nil Then
	begin
		THDefaut.OnElipsisClick:=SalElipsisClick;
		THDefaut.OnExit:=ExitEdit;
	end;
	Check:=TCheckBox(GetControl('CKSORTIE'));
	if Check=nil then
	Begin
		SetControlVisible('DATEARRET',False);
		SetControlVisible('TDATEARRET',False);
	End
	else
	Check.OnClick:=OnClickSalarieSortie;
	//PT2 - Début
	If PGBundleHierarchie And PGDroitMultiForm Then
	Begin
		SetControlVisible('PSA_ETABLISSEMENT', False);
		SetControlVisible('TPSA_ETABLISSEMENT', False);
	End;
	//PT3 - Début
	If PGBundleHierarchie Then
	Begin
		If Not PGDroitMultiForm Then
		Begin
			SetControlText('XX_WHEREPREDEF', 'PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_NOMSERVICE LIKE "'+V_PGI.NoDossier+'%")');
			THValComboBox(GetControl('PSA_ETABLISSEMENT')).Name := 'PSI_ETABLISSEMENT';
		End
		Else If V_PGI.ModePCL='1' Then
			SetControlText('XX_WHEREPREDEF', 'PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE '+ServicesDesDossiers()+')');
	End;
	//PT3 - Fin

	If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
	Begin
		TFMul(Ecran).SetDBListe ('PGDEPORTSALINTERI');
	End;
	//PT2 - Fin
end ;

procedure TOF_MUL_AFFICHERESP.SalElipsisClick(Sender:TObject);
var St:String;
begin             
     //PT2 - Début
     If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
     Begin
          St:='PSI_INTERIMAIRE IN (SELECT PGS_'+THEdit(Sender).Name+' FROM SERVICES)';
		  //PT3 - Début
          If PGBundleHierarchie Then
          	ElipsisSalarieMultidos(Sender, St)
          Else
		  //PT3 - Fin
          	LookupList(THEdit(Sender),'Liste des salariés','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE',St,'PSI_INTERIMAIRE', False,-1);
     End
     Else
     Begin
          St:='PSA_SALARIE IN (SELECT PGS_'+THEdit(Sender).Name+' FROM SERVICES)';
          LookupList(THEdit(Sender),'Liste des salariés','SALARIES','PSA_SALARIE','PSA_LIBELLE',St,'PSA_SALARIE', False,-1);
     End;
     //PT2 - Fin
end;

Initialization
  registerclasses ( [ TOF_MUL_AFFICHERESP ] ) ;
end.
