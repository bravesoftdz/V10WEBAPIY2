{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/08/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : ORGANIGRAMME_TV ()
Mots clefs ... : TOF;ORGANIGRAMME_TV
*****************************************************************
PT1   | 13/02/2008 | FL | V_803 | En multidossier, utilisation de la tablette SALARIEINT et de la liste PGINTSERVICE
}
Unit UTOFPGConsultService ;

Interface

uses     UTOF,
{$IFNDEF EAGLCLIENT}
         DBCtrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,
{$ELSE}
       UTOB,eMul,
{$ENDIF}
         Hctrls,ComCtrls,HEnt1,HMsgBox,StdCtrls,Classes,
         sysutils,Vierge,EntPaie,PgOutils,PGoutils2;

Type
  TOF_PGCONSULTSERVICE= Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    service:String;
    procedure ExitEdit(Sender: TObject);
    Procedure ClickElpsisSalarie (Sender: TObject); //PT1
  end ;

Implementation
Uses PGOutilsFormation, PGHierarchie;

procedure TOF_PGCONSULTSERVICE.OnLoad;
begin
  inherited;
	Ecran.Caption := 'Consultation du service '+ RechDom('PGSERVICE',Service,False);
	UpdateCaption(TFmul(Ecran));
end;

procedure TOF_PGCONSULTSERVICE.OnUpdate  ;
begin
  Inherited ;
end;

procedure TOF_PGCONSULTSERVICE.OnArgument (S : String ) ;
var where:String;
    Salarie:THEdit;
    Resp,Adjoint,Secretaire,RespNdf,RespAbs,RespVar,Nom{,Prenom},Matricule:String;
    LResp,LAdjoint,LSecretaire,LRespNdf,LRespVar,LRespAbs:THLabel;
    Q{,QSal}:TQuery;
    a:Integer;
begin
	Service:=ReadTokenSt(S);
	Where:='PSE_CODESERVICE="'+Service+'"';
	SetControlText('XX_WHERE',Where);
	
	Salarie:=THEdit(GetControl('PSA_SALARIE'));
	If Salarie<>nil then Salarie.OnExit:=ExitEdit;
	
	LResp:=THLabel(GetControl('LRESP'));
	LAdjoint:=THLabel(GetControl('LADJOINT'));
	LSecretaire:=THLabel(GetControl('LSECRETAIRE'));
	LRespAbs:=THLabel(GetControl('LRESPABS'));
	LRespNdf:=THLabel(GetControl('LRESPNDF'));
	LRespVar:=THLabel(GetControl('LRESPVAR'));
	
	Q:=OpenSQL('SELECT PGS_RESPSERVICE,PGS_ADJOINTSERVICE,PGS_SECRETAIRE,PGS_RESPONSVAR,PGS_RESPONSNDF,PGS_RESPONSABS FROM SERVICES WHERE PGS_CODESERVICE="'+Service+'"',True);
	Resp:=Q.FindField('PGS_RESPSERVICE').AsString;
	Adjoint:=Q.FindField('PGS_ADJOINTSERVICE').AsString;
	Secretaire:=Q.FindField('PGS_SECRETAIRE').AsString;
	RespAbs:=Q.FindField('PGS_RESPONSABS').AsString;
	RespNdf:=Q.FindField('PGS_RESPONSNDF').AsString;
	RespVar:=Q.FindField('PGS_RESPONSVAR').AsString;
	Ferme(Q);
	
	Nom:='';
	//Prenom:='';
	
	For a:=1 to 6 do
	begin
		If a=1 then Matricule:=Resp;
		If a=2 then Matricule:=Adjoint;
		If a=3 then Matricule:=Secretaire;
		If a=4 then Matricule:=RespAbs;
		If a=5 then Matricule:=RespNdf;
		If a=6 then Matricule:=RespVar;
		
		//PT1 - Début
		(*QSal:=OpenSQL('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="'+Matricule+'"',True);
		// CCMX-CEGID ORGANIGRAMME Suite violation en EAGL
		If not QSal.eof then
		Begin
			Nom:=QSal.FindField('PSA_LIBELLE').AsString;
			Prenom:=QSal.FindField('PSA_PRENOM').AsString;
		End
		else
		Begin
			Nom := '';
			Prenom := '';
		End;
		Ferme(QSal);*)
		If Not PGBundleHierarchie Then
			Nom := RechDom('PGSALARIE', Matricule, False)
		Else
			Nom := RechDom('PGSALARIEINT', Matricule, False);
		//PT1 - Fin
		
		If a=1 Then LResp.Caption:=Nom;//+' '+Prenom; //PT1
		If a=2 Then LAdjoint.Caption:=Nom;//+' '+Prenom; //PT1
		If a=3 Then LSecretaire.Caption:=Nom;//+' '+Prenom; //PT1
		If a=4 Then LRespAbs.Caption:=Nom;//+' '+Prenom; //PT1
		If a=5 Then LRespNdf.Caption:=Nom;//+' '+Prenom; //PT1
		If a=6 Then LRespVar.Caption:=Nom;//+' '+Prenom; //PT1
	end;
	
	//PT1
	If PGBundleHierarchie Then 
	Begin
		TFMul(Ecran).SetDBListe ('PGINTSERVICE'); 
		
		Salarie := THEdit(GetControl('PSA_SALARIE'));
		If Salarie <> Nil Then Salarie.OnElipsisClick := ClickElpsisSalarie;
	End;
end;

procedure TOF_PGCONSULTSERVICE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
	edit:=THEdit(Sender);
	if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
	if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
	edit.text:=AffectDefautCode(edit,10);
end;

//PT1
Procedure TOF_PGCONSULTSERVICE.ClickElpsisSalarie (Sender: TObject);
Var Dossier : String;
    Q       : TQuery;
Begin
	Q := OpenSQL('SELECT PGS_NOMSERVICE FROM SERVICES WHERE PGS_CODESERVICE="'+Service+'"',True);
	If Not Q.EOF Then Dossier := GetDossierService(Q.FindField('PGS_NOMSERVICE').AsString);
	Ferme(Q);
	
	If Dossier <> '' Then Dossier := 'PSI_NODOSSIER="'+Dossier+'"';
	
	ElipsisSalarieMultidos (Sender, Dossier);
End;

Initialization
  registerclasses ( [ TOF_PGCONSULTSERVICE ] ) ;
end.
