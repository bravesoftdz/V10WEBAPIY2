{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 20/03/2002
Modifié le ... :   /  /
Description .. : Multi critère gestion des infos complémentaires des salariés
Mots clefs ... : PAIE
*****************************************************************
PT1  | 19/09/2007 | FL | V_80  | En multidossiers, ne pas afficher les combos établissement
PT2  | 20/03/2008 | FL | V_803 | Adaptation accès multi dossiers
PT3  | 02/04/2008 | FL | V_803 | Adaptation pour restriction d'accès
}
unit UTOFPGMul_Service;

interface
Uses StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
     UTOF,
{$IFNDEF EAGLCLIENT}
     DBGrids, HDB,db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     emul,eFiche,UTob,
{$ENDIF}
     HCtrls,EntPaie,HEnt1,PgOutils2;

Type
  TOF_PGMUL_SERVICE = Class (TOF)
      procedure OnLoad                   ; override ;
      procedure OnArgument(S : String )  ; override ;
      private
      procedure ExitEdit(Sender: TObject);
     end ;

implementation

Uses PGOutilsFormation,PgOutils,PGHierarchie;

{ TOF_PGMUL_SERVICE }

procedure TOF_PGMUL_SERVICE.OnLoad;
var Dossier,DosCur : String;
    Chaine  : String;
begin
inherited;
	If GetControlText('CINCLUS')='X' Then
	   begin
	   SetControlText('PGS_SERVICESUP','');
	   end
	Else SetControlText('XX_WHERE','');

	//PT2 - Début
	If PGBundleHierarchie Then
	Begin
			Dossier := GetControlText('NODOSSIER');
			If Dossier <> '' Then
			Begin
	        	// Prise en compte de plusieurs dossiers
	        	If Pos(';', Dossier) = 0 Then
	        		Chaine := 'PGS_NOMSERVICE LIKE "'+Dossier+'%"'
	        	Else
	        	Begin
	        		While Dossier <> '' Do
	        		Begin
	        			DosCur := ReadTokenSt(Dossier);
	        			Chaine := Chaine + 'PGS_NOMSERVICE LIKE "'+DosCur+'%" OR ';
	        		End;
	        		Chaine := Copy (Chaine,1,Length(Chaine)-3);
	        	End;
				SetControlText('XX_WHEREPREDEF', Chaine);
			End
			Else
				SetControlText('XX_WHEREPREDEF', ServicesDesDossiers());
	End
	Else
		SetControlText('XX_WHEREPREDEF', '');
	//PT2 - Fin
end;

procedure TOF_PGMUL_SERVICE.OnArgument(S: String);
var Resp,RespAbs,RespNdf,RespVar:THEdit;
    Min,Max:String;
    Q:TQuery;
    NivInf,NivSup:THValComboBox;
begin
	inherited;
	Resp:=THEdit(GetControl('PGS_RESPSERVICE'));
	If Resp<>Nil Then Resp.OnExit:=ExitEdit;
	RespAbs:=THEdit(GetControl('PGS_RESPONSABS'));
	If RespAbs<>Nil Then RespAbs.OnExit:=ExitEdit;
	RespNdf:=THEdit(GetControl('PGS_RESPONSNDF'));
	If RespNdf<>Nil Then RespNdf.OnExit:=ExitEdit;
	RespVar:=THEdit(GetControl('PGS_RESPONSVAR'));
	If RespVar<>Nil Then RespVar.OnExit:=ExitEdit;
	Q:=OpenSQL('SELECT PHO_HIERARCHIE FROM HIERARCHIE ORDER BY PHO_NIVEAUH',True);
	If not Q.eof Then   // PortageCWAS
	begin
		Q.First;
		Min:=Q.FindField('PHO_HIERARCHIE').AsString;
		Q.Last;
		Max:=Q.FindField('PHO_HIERARCHIE').AsString;
	end;
	Ferme(Q);
	NivInf:=THValComboBox(GetControl('PGS_HIERARCHIE'));
	NivSup:=THValComboBox(GetControl('PGS_HIERARCHIE_'));
	If NivInf<>Nil Then NivInf.Value:=Min;
	If NivSup<>Nil Then NivSup.Value:=Max;

	//PT1 - Début
	If PGBundleHierarchie And PGDroitMultiForm Then
	Begin
		SetControlVisible('TPGS_ETABLISSEMENT', False);
		SetControlVisible('PGS_ETABLISSEMENT', False);
	End;
	//PT1 - Fin

	//PT2 - Début
	If PGBundleHierarchie then
	begin
		If Not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER', False);
			SetControlText   ('NODOSSIER', V_PGI.NoDossier);
			//PT3
			SetControlProperty('PGS_CODESERVICE', 'Plus', 'PGS_NOMSERVICE LIKE "'+V_PGI.NoDossier+'%"'); 
			SetControlProperty('PGS_SERVICESUP',  'Plus', 'PGS_NOMSERVICE LIKE "'+V_PGI.NoDossier+'%"'); 
			SetControlProperty('PGS_RESPSERVICE', 'Plus', DossiersAInterroger('DOS',V_PGI.NoDossier,'PSI',False,True)); 
			SetControlProperty('PGS_RESPONSABS',  'Plus', DossiersAInterroger('DOS',V_PGI.NoDossier,'PSI',False,True)); 
			SetControlProperty('PGS_RESPONSVAR',  'Plus', DossiersAInterroger('DOS',V_PGI.NoDossier,'PSI',False,True)); 
			SetControlProperty('PGS_RESPONSNDF',  'Plus', DossiersAInterroger('DOS',V_PGI.NoDossier,'PSI',False,True)); 
		end
		//PT3
		Else If V_PGI.ModePCL='1' Then 
		Begin
			SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); 
			SetControlProperty('PGS_CODESERVICE', 'Plus', ServicesDesDossiers); 
			SetControlProperty('PGS_SERVICESUP',  'Plus', ServicesDesDossiers); 
			SetControlProperty('PGS_RESPSERVICE', 'Plus', DossiersAInterroger('DOS','','PSI',False,True)); 
			SetControlProperty('PGS_RESPONSABS',  'Plus', DossiersAInterroger('DOS','','PSI',False,True)); 
			SetControlProperty('PGS_RESPONSVAR',  'Plus', DossiersAInterroger('DOS','','PSI',False,True)); 
			SetControlProperty('PGS_RESPONSNDF',  'Plus', DossiersAInterroger('DOS','','PSI',False,True)); 
		End;
	end
	else
	begin
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT2 - Fin
end;

procedure TOF_PGMUL_SERVICE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
	edit:=THEdit(Sender);
	if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

Initialization
registerclasses([TOF_PGMUL_SERVICE]);
end.
