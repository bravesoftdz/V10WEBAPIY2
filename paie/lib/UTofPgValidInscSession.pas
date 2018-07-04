{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 07/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGVALIDINSCSESSION ()
Mots clefs ... : TOF;PGVALIDINSCSESSION
*****************************************************************
---- JL 17/10/2006 Modification contrôle des exercices de formations -----
PT1 | 18/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT2 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT3 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
}
Unit UTofPgValidInscSession ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ELSE}
      MaineAgl,
{$ENDIF}
     sysutils,ComCtrls,HCtrls,HEnt1,UTOF,UTOB,uTableFiltre,
     SaisieList,LookUp,PGOutilsFormation;

Type
  TOF_PGVALIDINSCSESSION = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TF: TTableFiltre;
    TypeSaisie,TypeInsc,MillesimeEC:String;
    THVMillesime:THValComboBox;
    procedure ParamTV(Sender:TObject);
    procedure TreeViewdblClick(Sender:TObject);
    procedure StageElipsisClick(Sender:TObject);
    procedure ChangeMillesime(Sender:TObject);
  end ;

Implementation

Uses GalOutil;

procedure TOF_PGVALIDINSCSESSION.TreeViewdblClick(Sender:TObject);
var MyTreenode1:tTreeNode;
    Niveau:Integer;
    St:String;
begin
MyTreenode1:=TF.LeTreeView.Selected;
Niveau:=myTreeNode1.Level;
if Niveau=0 Then
   begin
   St:=TF.TOBFiltre.GetValue('PSS_CODESTAGE')+';'+TF.TOBFiltre.GetValue('PSS_MILLESIME');
   AglLanceFiche('PAY','STAGE','',St,'CONSULTATION');
   end;
if Niveau=1 Then
   begin
   St:=TF.TOBFiltre.GetValue('PSS_CODESTAGE')+';'+IntToStr(TF.TOBFiltre.GetValue('PSS_ORDRE'))+';'+TF.TOBFiltre.GetValue('PSS_MILLESIME');
   AglLanceFiche('PAY','SESSIONSTAGE','',St,'CONSULTATION');
   end;
end;
procedure TOF_PGVALIDINSCSESSION.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGVALIDINSCSESSION.OnLoad ;
begin
  Inherited ;
	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS')); //PT1
end ;

procedure TOF_PGVALIDINSCSESSION.OnArgument (S : String ) ;
var Datedebut,DateFin:TDatetime;
    Edit:THEdit;
begin
	TypeSaisie:=ReadTokenSt(S);
	TypeInsc:=ReadTokenSt(S);
	if (Ecran<>nil) and (Ecran is TFSaisieList ) then
	TF := TFSaisieList(Ecran).LeFiltre;
	
	If TypeSaisie <> 'CWASVALIDINSC' then TF.WhereTable := 'WHERE PFO_CODESTAGE=:PSS_CODESTAGE AND PFO_ORDRE=:PSS_ORDRE AND PFO_MILLESIME=:PSS_MILLESIME'
	else TF.WhereTable := 'WHERE PFO_CODESTAGE=:PSS_CODESTAGE AND PFO_ORDRE=:PSS_ORDRE AND PFO_MILLESIME=:PSS_MILLESIME AND PFO_RESPONSFOR="'+V_PGI.UserSalarie+'"';
	
    //PT3
    // N'affiche pas les salariés confidentiels
    If TF.WhereTable <> '' Then TF.WhereTable := TF.WhereTable + ' AND ';
    If PGBundleHierarchie Then
    	TF.WhereTable := TF.WhereTable + 'PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
    Else
    	TF.WhereTable := TF.WhereTable + 'PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
    
	TF.LeTreeView.OnDblClick:=TreeViewdblClick;
	TFSaisieList(Ecran).BCherche.OnClick:=ParamTV;
	
	MillesimeEC:=RendMillesimeRealise(DateDebut,DateFin);;
	SetControlText('PSS_DATEDEBUT',DateToStr(DateDebut));
	SetControlText('PSS_DATEFIN',DateToStr(DateFin));
	
	Edit:=THEdit(GetControl('PSS_CODESTAGE'));
	If Edit<>Nil Then Edit.OnelipsisClick:=StageElipsisClick;
	
	THVMillesime:=THValComboBox(GetControl('THVMILLESIME'));
	THVMillesime.Value:=MillesimeEC;
	THVMillesime.OnChange:=ChangeMillesime;
	
	//PT1 - Début
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PSS_PREDEFINI',False);
			SetControlText   ('PSS_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT2        
	end
	else
	begin
		SetControlVisible('PSS_PREDEFINI', False);
		SetControlVisible('TPSS_PREDEFINI',False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT1 - Fin
	
	TFSaisieList(Ecran).BCherche.Click;
	
	SetControlVisible('BPARAMLISTE',True);
	SetActiveTabSheet('PSESSION');
end ;

procedure TOF_PGVALIDINSCSESSION.ParamTV(Sender:TObject);
var Previsionnel:THValComboBox;
begin
	If TypeSaisie <> 'CWASVALIDINSC' then
		TF.WhereTable := 'WHERE PFO_CODESTAGE=:PSS_CODESTAGE AND PFO_ORDRE=:PSS_ORDRE AND PFO_MILLESIME=:PSS_MILLESIME'
	else
		TF.WhereTable := 'WHERE PFO_CODESTAGE=:PSS_CODESTAGE AND PFO_ORDRE=:PSS_ORDRE AND PFO_MILLESIME=:PSS_MILLESIME AND PFO_RESPONSFOR="'+V_PGI.UserSalarie+'"';
	//PT3
	// N'affiche pas les salariés confidentiels
	If TF.WhereTable <> '' Then TF.WhereTable := TF.WhereTable + ' AND ';
	If PGBundleHierarchie Then
		TF.WhereTable := TF.WhereTable + ' PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")'
	Else
		TF.WhereTable := TF.WhereTable + ' PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
	
	MillesimeEC:=THVMillesime.Value;
	Previsionnel:=THValComboBox(GetControl('CMILLESIME'));
	If Previsionnel.Value='SB' Then SetControltext('PSS_MILLESIME',MillesimeEC)
	Else If Previsionnel.Value='SNB' Then SetControltext('PSS_MILLESIME','0000')
	Else SetControlText('PSS_MILLESIME','');
	TFSaisieList(Ecran).BChercheClick(Sender);
end ;


procedure TOF_PGVALIDINSCSESSION.StageElipsisClick(Sender:TObject);
var  StWhere,StOrder : String;
begin
StWhere := '(PST_ACTIF="X" AND PST_MILLESIME="'+MillesimeEC+'" ';
If PGBundleInscFormation Then
Begin
	If Not PGDroitMultiForm Then 
		StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True) //PT2
	Else
		StWhere := StWhere + DossiersAInterroger('','','PST',True,True); //PT2
End;
StWhere := StWhere + ') OR (PST_MILLESIME="0000" AND '+
    'PST_CODESTAGE NOT IN (SELECT PST_CODESTAGE FROM STAGE WHERE PST_MILLESIME="'+MillesimeEC+'")';
    
    If PGBundleInscFormation Then
    Begin
    	If Not PGDroitMultiForm Then 
    		StWhere := StWhere + DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True) //PT2
    	Else
    		StWhere := StWhere + DossiersAInterroger('','','PST',True,True); //PT2
    End;
StWHere := StWhere + ')';
    
StOrder := 'PST_MILLESIME,PST_LIBELLE';
LookupList(THEdit(Sender),'Liste des stages','STAGE','PST_CODESTAGE','PST_LIBELLE,PST_MILLESIME',StWhere,StOrder, True,-1);
end;

procedure TOF_PGVALIDINSCSESSION.ChangeMillesime(Sender:TObject);
var Q:TQuery;
begin
Q:=OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+THVMillesime.Value+'"',True);
if not Q.eof then
   begin
   SetControlText('PSS_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
   SetControlText('PSS_DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
   end;
Ferme(Q);
end;

Initialization
  registerclasses ( [ TOF_PGVALIDINSCSESSION ] ) ;
end.
