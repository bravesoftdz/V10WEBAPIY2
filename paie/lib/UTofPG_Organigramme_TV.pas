{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 20/03/2002
Modifié le ... :   /  /
Description .. : Construction de l'organigramme sous forme de TreeView
Mots clefs ... : TOF;ORGANIGRAMME_TV,TreeView
*****************************************************************
PT1 12/09/2002 V582 JL Vérification existence de service avant affichage du TreeView
PT2 22/11/2002 V591 JL Correction requête pour eAGL
PT3 31/05/2007 V_72 JL FQ 14021 Ajout apramsoc pour gestion intervenants exterieurs
PT4 19/09/2007 FL V_80   En multidossier, ne pas afficher les combos établissement
PT5 13/02/2008 FL V_803  En multidossier, utiliser la tablette SALARIEINT plutôt que SALARIE
PT6 20/03/2008 FL V_803  Ajout du n° de dossier (qui est en fait compris dans le nom de service).
PT7 02/04/2008 FL V_803  Adaptation de l'édition pour le partage de la hiérarchie
PT8 02/04/2008 FL V_803  Correctifs suite à la gestion des dossiers pour les services
PT9 03/04/2008 FL V_803  Cache les champs établissement si partage hiérarchie
}

Unit UTofPG_Organigramme_TV ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,FE_Main,
{$ELSE}
     UtileAGL,MaineAgl,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF,Hqry,Vierge,AGLInit,Graphics,HTB97,
     UTob,PgEditOutils,PgEditOutils2,EntPaie,Spin,windows,PGOutils,LookUp,paramsoc ;


Type
  TOF_ORGANIGRAMME_TV = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    BService: TToolbarButton97;
    TVORG : TTreeview;
    MyTreeNode1: TTreeNode;
    CDerouler,CService,CNiveau,Ascendants:TCheckBox;
    EService:THEdit;
    VNivInf,VNivSup:THValComboBox;
    TVTop,TVHaut,BAffTop:Integer;
    procedure AfficheEtab(Sender:TObject);
    procedure AfficheService(Sender:TObject);
    procedure ChoixBooleen(Sender:TObject);
    procedure Derouler(Sender:TObject);
    procedure MAJTable;
    procedure CreationNoeuds(TR,TN:Tob);
    procedure AfficheInfos(TR,TN:Tob);
    Procedure ImpressionTV(Sender:TObject);
    end ;
    
    //PT7
    TOF_EDITORGANIGRAMME = Class(TOF)
    procedure OnArgument (S : String ) ; override ;
	End;

Implementation

Uses PGOutilsFormation,PGHierarchie;

procedure TOF_ORGANIGRAMME_TV.OnUpdate  ;
Var Tob_Racines,Tob_Noeuds:Tob;
    QNoeuds,QRacines,Q:TQuery;
    ChoixNiv,ChoixService:Boolean;
    Choix,NivInf,NivSup:Integer;
    SNivInf,SNivSup:String;
    ServiceRacine,Libelle,Etab1,Etab2,StResp:String;
    StWhere : String;
begin
	Inherited ;
	SetActiveTabSheet('PTREEVIEW');
	Etab1:=GetControlText('ETAB1');
	Etab2:=GetControlText('ETAB2');
	StResp:=',PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_ADJOINTABS,PGS_ADJOINTNDF,PGS_ADJOINTVAR'+
	',PGS_SECRETAIREABS,PGS_SECRETAIREVAR,PGS_SECRETAIRENDF';
	
	// Création des 2 TOB utilisées pour la création du TreeView
	WITH tvorg.items do // Effacement des données du TreeView
	begin
		clear;
	end;
	
	//PT8 - Début
	If PGBundleHierarchie Then
    Begin
        If  GetControlText('NODOSSIER') <> '' Then
    		StWhere := ' AND PGS_NOMSERVICE LIKE "'+GetControlText('NODOSSIER')+'%" '
        Else
    		StWhere := ' AND ('+ ServicesDesDossiers() +')';
    End
	Else
		StWhere := '';
	//PT8 - Fin
	
	// Initialisation des variables pour les choix des services et des niveaux
	ChoixNiv:=False;
	ChoixService:=False;
	ServiceRacine:='';
	If (CService<>Nil) and (CService.Checked=True) Then ChoixService:=True;
	If (CNiveau<>Nil) and (CNiveau.Checked=True) Then  ChoixNiv:=True;
	If EService<>Nil Then ServiceRacine:=EService.Text;
	
	If (VNivInf<>Nil) and (VNivInf.Value<>'') Then
	begin
		SNivInf:=VNivInf.Value;
		Q:=OpenSQL('SELECT PHO_NIVEAUH FROM HIERARCHIE WHERE PHO_HIERARCHIE="'+SNivInf+'"',True);
		NivInf:=0;
		If not Q.eof then NivInf:=Q.FindField('PHO_NIVEAUH').AsInteger;  // PT2
		Ferme(Q);
	end
	Else NivInf:=1;
	
	If (VNivSup<>Nil) and (VNivSup.Value<>'') Then
	begin
		SNivSup:=VNivSup.Value;
		Q:=OpenSQL('SELECT PHO_NIVEAUH FROM HIERARCHIE WHERE PHO_HIERARCHIE="'+SNivSup+'"',True);
		NivSup:=0;
		If not Q.eof Then NivSup:=Q.FindField('PHO_NIVEAUH').AsInteger;   // PT2
		Ferme(Q);
	end
	Else NivSup:=20;
	
	If ChoixNiv=True Then
	begin
		If NivSup<NivInf then
		begin
			PGIBox('Erreur dans le choix des niveaux, Veuillez changer les valeurs.','Visualisation de l''organigramme');
			SetFocusControl('SNIVINF');
			SetActiveTabSheet('PCRITERES');
			Exit;
		end;
	end;
	
	MAJTable;
	
	Choix:=0;
	If ChoixService=True Then
	begin
		If EService.text='' Then begin PgiBox('Vous devez renseigner le service ou décocher la case "choix d''un service"','Visualisation organisramme');
		SetActiveTabSheet('PCRITERES');Exit;end;
		Choix:=1;
	end;

	If ChoixNiv=True Then Choix:=2;
	If (ChoixService=False) and (ChoixNiv=False) Then Choix:=3;

	If PGBundleHierarchie And (Choix=3) Then Choix := 4; //PT6
	
	Case Choix of
		1:
		begin
			If Ascendants.Checked=False Then QRacines:=OpenSQL('SELECT PSO_CODESERVICE,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
			' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE'+StResp+
			' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
			' WHERE PSO_CODESERVICE="'+ServiceRacine+'" AND (PSO_NIVEAUSUP=1 OR PSO_NIVEAUSUP=0) ORDER BY PGS_NOMSERVICE',True)
			Else QRacines:=OpenSQL('SELECT PSO_CODESERVICE,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
			' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE'+StResp+
			' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
			' WHERE PSO_SERVICESUP="-" AND ((PSO_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE WHERE PSO_CODESERVICE="'+ServiceRacine+'"))'+
			' OR (PSO_CODESERVICE="'+ServiceRacine+'")) ORDER BY PGS_NOMSERVICE',True);
			ServiceRacine:='';
			If Not QRacines.Eof then ServiceRacine:=QRacines.FindField('PSO_CODESERVICE').AsString; // PT2
		end;
		2:
		begin
			If NivInf=1 Then QRacines:=OpenSQL('SELECT PSO_CODESERVICE,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
			' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE'+StResp+
			' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
			' LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE'+
			' WHERE PHO_NIVEAUH="'+IntToStr(NivInf)+'" and ((PGS_ETABLISSEMENT>="'+Etab1+'"'+
			' and PGS_ETABLISSEMENT<="'+Etab2+'") or pgs_ETABLISSEMENT="") '+StWhere+'ORDER BY PGS_NOMSERVICE',True)
			Else QRacines:=OpenSQL('SELECT PSO_CODESERVICE,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
			' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE'+StResp+
			' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
			' LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE'+
			' WHERE PSO_NIVEAUSERVICE>='+InttoStr(NivInf)+' AND PSO_NIVEAUSERVICE<='+InttoStr(NivSup)+' and ((PGS_ETABLISSEMENT>="'+Etab1+'"'+
			' and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="") AND PSO_NIVEAUSUP=1 AND PSO_SERVICESUP IN'+
			' (SELECT PSO_CODESERVICE FROM SERVICEORDRE WHERE PSO_NIVEAUSERVICE<'+IntToStr(NivInf)+' )'+StWhere+
			' ORDER BY PGS_NOMSERVICE',True);
		end;
	
		3:QRacines:=OpenSQL('SELECT PSO_CODESERVICE,PSO_NIVEAUSUP,PSO_SERVICESUP,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
		' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE'+StResp+
		' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
		' WHERE PSO_SERVICESUP="-" and ((PGS_ETABLISSEMENT>="'+Etab1+'" and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="") ORDER BY '+
		' PGS_ETABLISSEMENT,PGS_NOMSERVICE',True);
		
		//PT6
		4:QRacines:=OpenSQL('SELECT PSO_CODESERVICE,PSO_NIVEAUSUP,PSO_SERVICESUP,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
		' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE'+StResp+
		' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
		' WHERE (PSO_SERVICESUP="-" OR PSO_SERVICESUP=(SELECT MIN(PSO_SERVICESUP) FROM SERVICEORDRE WHERE '+
		'((PGS_ETABLISSEMENT>="'+Etab1+'" and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="") '+StWhere+')) AND ((PGS_ETABLISSEMENT>="'+Etab1+'" and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="") ORDER BY '+
		' PGS_ETABLISSEMENT,PGS_NOMSERVICE',True);
	End;
	If QRacines.Eof Then begin Ferme(QRacines);Exit;end; //PT1
	TOB_Racines:=TOB.Create('les racines',NIL,-1);
	TOB_Racines.LoadDetailDB('les racines','','',QRacines,False);
	Ferme(QRacines);
	
	Case Choix of
		1:
		begin
			If Ascendants.Checked=False Then QNoeuds:=OpenSQL('SELECT PSO_CODESERVICE,PSO_SERVICESUP,PSO_NIVEAUSUP,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
			' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE,PGS_SERVICESUP'+StResp+
			' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE WHERE PSO_SERVICESUP="'+ServiceRacine+'" ORDER BY PGS_NOMSERVICE',True)
			Else QNoeuds:=OpenSQL('SELECT PSO_CODESERVICE,PSO_SERVICESUP,PSO_NIVEAUSUP,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
			' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE,PGS_SERVICESUP'+StResp+
			' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE WHERE PSO_SERVICESUP="'+ServiceRacine+'" AND'+
			' ((PSO_CODESERVICE IN (SELECT PSO_CODESERVICE FROM SERVICEORDRE WHERE PSO_SERVICESUP="'+EService.Text+'"))'+
			' OR (PSO_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE WHERE PSO_CODESERVICE="'+EService.Text+'" ))'+
			' OR (PSO_CODESERVICE="'+EService.Text+'")) ORDER BY PGS_NOMSERVICE',True);
		end;
		
		2: QNoeuds:=OpenSQL('SELECT PSO_CODESERVICE,PSO_SERVICESUP,PSO_NIVEAUSUP,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
		' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE,PGS_SERVICESUP'+StResp+
		' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
		' WHERE ((PGS_ETABLISSEMENT>="'+Etab1+'" and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="")'+
		' and PSO_NIVEAUSERVICE>='+IntToStr(NivInf)+' and PSO_NIVEAUService<='+IntToStr(NivSup)+''+StWhere+
		' ORDER BY PSO_NIVEAUSERVICE,PGS_NOMSERVICE',True);

		3:QNoeuds:=OpenSQL('SELECT PSO_CODESERVICE,PSO_SERVICESUP,PSO_NIVEAUSUP,PSO_NIVEAUSERVICE,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
		' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE,PGS_SERVICESUP'+StResp+
		' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
		' WHERE ((PGS_ETABLISSEMENT>="'+Etab1+'" and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="")'+
		' ORDER BY PSO_NIVEAUSERVICE,PGS_NOMSERVICE',True);
		
		//PT6
		4:QNoeuds:=OpenSQL('SELECT PSO_CODESERVICE,PSO_SERVICESUP,PSO_NIVEAUSUP,PSO_NIVEAUSERVICE,PGS_NOMSERVICE,PGS_HIERARCHIE,PGS_RESPSERVICE,PGS_ETABLISSEMENT'+
		' ,PGS_ADJOINTSERVICE,PGS_SECRETAIRE,PGS_SERVICESUP'+StResp+
		' FROM SERVICEORDRE LEFT JOIN SERVICES ON PSO_CODESERVICE=PGS_CODESERVICE'+
		' WHERE ((PGS_ETABLISSEMENT>="'+Etab1+'" and PGS_ETABLISSEMENT<="'+Etab2+'") OR PGS_ETABLISSEMENT="")'+StWhere+
		' ORDER BY PSO_NIVEAUSERVICE,PGS_NOMSERVICE',True);
	End;
	Tob_Noeuds:=TOB.Create('les noeuds',NIL,-1);
	Tob_Noeuds.LoadDetailDB('SERVICEORDRE','','',QNoeuds,False);
	Ferme(QNoeuds);
	
	CreationNoeuds(Tob_Racines,Tob_Noeuds);
	AfficheInfos(Tob_Racines,Tob_Noeuds);
	
	If TOB_Racines<>Nil Then TOB_Racines.free;
	If TOB_Noeuds<>Nil Then TOB_Noeuds.free;
	Derouler(CDerouler);
	
	If Ascendants.Checked=True Then
	begin
		With TVORG.Items do
		begin
			MyTreenode1:=TVORG.Items[0];
			While MyTreenode1<>NIL do
			begin
				Libelle:=MyTreeNode1.Text;
				If EService.Text=Trim(ReadTokenPipe(Libelle,'/')) Then
				begin
					MyTreeNode1.Selected:=True;
					MyTreeNode1.MakeVisible;
					break;
				end;
				MyTreenode1:=MytreeNode1.GetNext;
			end;
		end;
	end;
//SetActiveTabSheet('PVISUALISATION');
end;

procedure TOF_ORGANIGRAMME_TV.OnArgument (S : String ) ;
Var Q:TQuery;
    THEtab1,THEtab2 : THEdit;
    LDateMaj:THLabel;
    DateMaj:TDateTime;
    Min,Max: String;
    CMaj:TCheckBox;
    BImprime:TToolBarButton97;
begin
	SetActiveTabSheet('PCRITERES');
	SetControlChecked('CNIVEAU',True);
	SetControlChecked('CRESP',True);
	SetControlChecked('CNOMSER',True);
	BService:=TToolBarButton97(GetControl('BSERVICE'));
	If BService<>NIL Then BService.OnClick:=AfficheService;

	THEtab1:=THEdit(GetControl('ETAB1'));
	THEtab2:=THEdit(GetControl('ETAB2'));

	//PT4 - Début
	If PGBundleHierarchie And PGDroitMultiForm Then
	Begin
		If THEtab1<>NIL Then THEtab1.Visible := False;
		If THEtab2<>NIL Then THEtab2.Visible := False;
		SetControlVisible('TETAB', False);
		SetControlVisible('TETAB1', False);
		SetControlVisible('LIBETAB1', False);
		SetControlVisible('LIBETAB2', False);
		SetControlVisible('CETAB', False);
	End
	Else
	Begin
		If THEtab1<>NIL Then THEtab1.OnChange:=AfficheEtab;
		If THEtab2<>NIL Then THEtab2.OnChange:=AfficheEtab;
		RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
		SetControlText('Etab1',Min);
		SetControlText('Etab2',Max);
	End;
	//PT4 - Fin

	Q:=OpenSQL('SELECT PSO_DATEMAJ FROM SERVICEORDRE',True);
	If (NOT Q.EOF) AND (Q<>Nil) Then
	begin
		DateMAJ:=Q.FindField('PSO_DATEMAJ').AsDateTime;
		LDateMaj:=THLabel(GetControl('LDATEMAJ'));
		LDateMaj.Caption:=DateToStr(DateMaj);
	end;
	Ferme(Q);
	TVORG:= TTreeview(GetControl('TVORG'));
	TVORG.ReadOnly:=True;
	BAffTop:=BService.Top;
	TVTop:=TVOrg.Top;
	TVHaut:=TVOrg.Height;
	EService:=THEdit(GetControl('ESERVICE'));

	Ascendants:=TCheckBox(GetControl('CASCENDANTS'));
	If EService<>Nil Then EService.Enabled:=False;
	If Ascendants<>Nil Then Ascendants.Enabled:=False;
	VNivInf:=THValComboBox(GetControl('SNIVINF'));
	VNivSup:=THValComboBox(GetControl('SNIVSUP'));
	Q:=OpenSQL('SELECT PHO_HIERARCHIE FROM HIERARCHIE ORDER BY PHO_NIVEAUH',True);
	If not Q.eof Then   // PT2
	begin
		Q.First;
		Min:=Q.FindField('PHO_HIERARCHIE').AsString;
		Q.Last;
		Max:=Q.FindField('PHO_HIERARCHIE').AsString;
	end;
	Ferme(Q);
	If VNivInf<>Nil Then
	begin
		VNivInf.Value:=Min;
		VNivInf.Enabled:=False;
	end;
	If VNivSup<>Nil Then
	begin
		VNivSup.Value:=Max;
		VNivSup.Enabled:=False;
	end;
	CService:=TCheckBox(GetControl('CSERVICE'));
	If CService<>Nil Then CService.OnClick:=ChoixBooleen;
	CNiveau:=TCheckBox(GetControl('CNIV'));
	If CNiveau<>Nil Then CNiveau.OnClick:=ChoixBooleen;
	CDerouler:=TCheckBox(GetControl('CDEROULER'));
	If CDerouler<>Nil Then CDerouler.OnClick:=Derouler;
	CMaj:=TCheckBox(GetControl('CMAJ'));
	If CMaj<>Nil Then CMaj.Checked:=True;
	BImprime:=TToolBarButton97(GetControl('BIMPRIMER'));
	If BImprime<>Nil Then BImprime.OnClick:=ImpressionTV;
	
	//PT6 - Début
	If PGBundleHierarchie then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER', False);
			SetControlText   ('NODOSSIER', V_PGI.NoDossier);
		end
		Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT8
	end
	else
	begin
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT6 - DébuFin
end;

procedure TOF_ORGANIGRAMME_TV.AfficheService(Sender:TObject);
var Libelle,Service:String;
begin
	Service:='';
	With TVORG.Items do
	begin
		if TVORG.Selected = nil then begin PgiBox('Aucun service n''a été sélectionné','Affichage d''un service');Exit;end;
		MyTreeNode1 := TVORG.Selected;
		If MyTreeNode1.Selected=True Then Libelle:=MyTreeNode1.Text;
		Service:=Trim(ReadTokenPipe(Libelle,'/')) ;
		AGLLanceFiche ('PAY', 'SERVICE_CONSULT', '','', Service);
	end;
end;

procedure TOF_ORGANIGRAMME_TV.AfficheEtab(Sender:TObject);
var Etab1,Etab2:String;
    THLib1,THLib2:THLabel;
begin
	Etab1:=GetControlText('ETAB1');
	Etab2:=GetControlText('ETAB2');
	THLib1:=THLabel(GetControl('LIBETAB1'));
	THLib2:=THLabel(GetControl('LIBETAB2'));
	THLib1.caption:= RechDom ('TTETABLISSEMENT',Etab1,FALSE);
	THLib2.caption:= RechDom ('TTETABLISSEMENT',Etab2,FALSE);
end;

procedure TOF_ORGANIGRAMME_TV.ChoixBooleen(Sender:TObject);
begin
If Sender=CService Then
   begin
   If CService.Checked=True then
      begin
      CNiveau.Checked:=False;
      EService.Enabled:=True;
      Ascendants.Checked:=False;
      Ascendants.Enabled:=True;
      end
   Else
       begin
       EService.Enabled:=False;
       EService.Text:='';
       Ascendants.Enabled:=False;
       Ascendants.Checked:=False;
       end;
   end;
If Sender=CNiveau Then
   begin
   If CNiveau.Checked=True Then
      begin
      CService.Checked:=False;
      VNivInf.Enabled:=True;
      VNivSup.Enabled:=True;
      end
   Else
       begin
       VNivInf.Enabled:=False;
       VNivSup.Enabled:=False;
       end;
   end;
end;

procedure TOF_ORGANIGRAMME_TV.Derouler(Sender:TObject);
begin
If TCheckBox(Sender).Checked=True Then TVOrg.FullExpand
Else TVOrg.FullCollapse;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/01/2002
Modifié le ... :   /  /
Description .. : Mise à jour de la table organigramme
Mots clefs ... : PAIE;Organigramme
*****************************************************************}
//MAJ de la table ORDRESERVICES que si case MAJ cochée.
procedure TOF_ORGANIGRAMME_TV.MAJTable;
Var NomService,Service,ServiceSup:String;
    QSuperieur,Q,QService:TQuery;
    T,T2,TOB_Service,TOB_Superieur:TOB;
    Tab: array of string;
    y,a,Niveau:integer;
    LDateMaj:THLabel;
    DateMaj:TDateTime;
    CMaj:TCheckBox;//CHOIX DONNEES
    SNiveau:String;
begin
	CMaj:=TCheckBox(GetControl('CMAJ'));

	If CMaj.checked=True Then
	begin
		ExecuteSQL('DELETE FROM SERVICEORDRE'); // la table est vidée avant la création des nouveaux enregistrements.
		SetLength(Tab,20);

		QService:=OpenSQL('SELECT PGS_CODESERVICE,PGS_HIERARCHIE FROM SERVICES ORDER BY PGS_CODESERVICE',True); 
		TOB_Service:=TOB.Create('les services1',NIL,-1);
		TOB_Service.LoadDetailDB('SERVICES','','',QService,False);
		T := TOB_Service.FindFirst ([''],[''],False);
		Ferme (QService);

		QSuperieur:=OpenSQL('SELECT PGS_CODESERVICE,PGS_SERVICESUP FROM SERVICES',True);
		TOB_Superieur:=TOB.Create('Les services2',NIL,-1);
		TOB_Superieur.LoadDetailDB('SERVICES','','',QSuperieur,False);
		Ferme(QSuperieur);

		While T<>NIL do
		begin
			Service:=T.GetValue('PGS_CODESERVICE');
			SNiveau:=T.GetValue('PGS_HIERARCHIE');
			Q:=OpenSQL('SELECT PHO_NIVEAUH FROM HIERARCHIE WHERE PHO_HIERARCHIE="'+SNiveau+'"',True);
			Niveau:=0;
			If Not Q.eof then Niveau:=Q.FindField('PHO_NIVEAUH').AsInteger;  // PT2
			Ferme(Q);
			
			NomService:=RechDom('PGSERVICE',Service,False);
			tab[0]:=Service;
			T2:=TOB_Superieur.FindFirst(['PGS_CODESERVICE'],[Service],False);
			ServiceSup:=T2.GetValue('PGS_SERVICESUP');
			If ServiceSup='' Then y:=0
			Else
			begin
				Tab[1]:=ServiceSup;
				y:=2;
				While TOB_Superieur.FindFirst(['PGS_CODESERVICE'],[ServiceSup],False)<>NIL do
				begin
					If y>19 then break;
					T2:=TOB_Superieur.FindFirst(['PGS_CODESERVICE'],[ServiceSup],False);
					ServiceSup:=T2.GetValue('PGS_SERVICESUP');
					If ServiceSup<>'' Then
					begin
						Tab[y]:=ServiceSup;
						y:=y+1;
					end;
				end;
				y:=y-1;
			end;
			If y<20 Then
			begin
				If y<>0 Then
				begin
					for a:=1 to y do
					ExecuteSQL('INSERT INTO SERVICEORDRE (PSO_CODESERVICE,PSO_NIVEAUSERVICE,PSO_NIVEAUSUP,PSO_SERVICESUP,PSO_DATEMAJ)'+
					' VALUES '+
					' ("'+tab[0]+'","'+IntToStr(Niveau)+'","'+IntToStr(a)+'","'+tab[a]+'","'+UsDateTime(date)+'")');
				end;
				If y=0 Then ExecuteSQL('INSERT INTO SERVICEORDRE (PSO_CODESERVICE,PSO_NIVEAUSERVICE,PSO_NIVEAUSUP,PSO_SERVICESUP,PSO_DATEMAJ)'+
				' VALUES '+
				' ("'+tab[0]+'","'+IntToStr(Niveau)+'","","-","'+UsDateTime(Date)+'")') ;
			end;

			for a:=0 to 19 do tab[a]:='-';
			T := TOB_Service.FindNext ([''],[''], False);
		end;
		TOB_Service.Free;
		TOB_Superieur.Free;
		Q:=OpenSQL('SELECT PSO_DATEMAJ FROM SERVICEORDRE',True);  // Affichage de la nouvelle date
		If not Q.eof Then  // PT2
		begin
			DateMAJ:=Q.FindField('PSO_DATEMAJ').AsDateTime;
			LDateMaj:=THLabel(GetControl('LDATEMAJ'));
			LDateMaj.Caption:=DateToStr(DateMaj);
		end;
		Ferme(Q);
	end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/01/2002
Modifié le ... :   /  /
Description .. : Création des noeuds représentants les services de l'établissement
Mots clefs ... : PAIE;Organigramme,TTreeView
*****************************************************************}
procedure TOF_ORGANIGRAMME_TV.CreationNoeuds(TR,TN:Tob);
Var AffService,Service,ServiceSup,Test:String;
    x:integer;
    T,T2:TOB;
begin
Test:='';
T2 := nil;
T := TR.FindFirst ([''],[''], False);
//*************  Création des racines et des noeuds :
// La racine est crée en premier, puis pour chaque niveau, on prend tous les services qui dépendent directement ouindirectement de la racine
while T<>NIL do
      begin
      Service:=T.GetValue('PSO_CODESERVICE');
      if Test=service then break;
      test:=Service;
      With TVORG.Items do
              begin
              MyTreeNode1 := Add(nil, Service);  // création de la racine
              end;
      For x:=1 to 19 do      // début création des noeuds dépendant de la racine crée
          begin
          T2 := TN.FindFirst (['PSO_SERVICESUP','PSO_NIVEAUSUP'],[Service,x], FALSE);
          while T2<>NIL do
                begin
                AffService:=T2.GetValue('PSO_CODESERVICE');
                ServiceSup:=T2.GetValue('PGS_SERVICESUP');
                With TVORG.Items do
                     begin
                     MyTreeNode1 := TVORG.Items[0];
                     while MyTreenode1.text<>ServiceSup do       // Recherche du noeud correspondant au service supérieur
                          begin
                          MyTreenode1:=MytreeNode1.GetNext;
                          end;
                     AddChild(Mytreenode1, AffService);
                     end;
                T2 := TN.FindNext (['PSO_SERVICESUP','PSO_NIVEAUSUP'],[Service,x], FALSE);
                end;
          end;
      T := TR.FindNext ([''],[''], False);
      end;
If T<>Nil Then T.Free;
If T2<>Nil Then T2.Free;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/01/2002
Modifié le ... :   /  /
Description .. : Affichage des informations sélectionnées par l'utilisateur
Tous les noeuds sont repris.
Mots clefs ... : PAIE;Organigramme,TTreeView
*****************************************************************}
procedure TOF_ORGANIGRAMME_TV.AfficheInfos(TR,TN:Tob);
var CEtab,CNiv,CResp,CAdj,CSecr,CRespAbs,CRespNdf,CRespVar,CAdjAbs,CAdjVar,CAdjNdf,CSecrAbs,CSecrNdf,CSecrvar,CNomSer:TCheckBox;//CHOIX DONNEES
    DonneeAffiche,Adjoint,Secretaire,RespNdf,RespAbs,RespVar,AdjointAbs,AdjointVar,AdjointNdf,SecretaireVar,SecretaireNdf,SecretaireAbs:String;
    Service,Etablissement,Nom,NiveauS,Resp:String;
    T:TOB;

begin
T := nil;
CNomSer:=TCheckBox(GetControl('CNOMSER'));
CEtab:=TCheckBox(GetControl('CETAB'));
CNiv:=TCheckBox(GetControl('CNIVEAU'));
CResp:=TCheckBox(GetControl('CRESP'));
CAdj:=TCheckBox(GetControl('CADJOINT'));
CSecr:=TCheckBox(GetControl('CSECRETAIRE'));
CRespAbs:=TCheckBox(GetControl('CRESPABS'));
CRespNdf:=TCheckBox(GetControl('CRESPNDF'));
CRespVar:=TCheckBox(GetControl('CRESPVAR'));
CAdjAbs:=TCheckBox(GetControl('CADJOINTABS'));
CAdjVar:=TCheckBox(GetControl('CADJOINTVAR'));
CAdjNdf:=TCheckBox(GetControl('CADJOINTNDF'));
CSecrAbs:=TCheckBox(GetControl('CSECRETAIREABS'));
CSecrVar:=TCheckBox(GetControl('CSECRETAIREVAR'));
CSecrNdf:=TCheckBox(GetControl('CSECRETAIRENDF'));
MyTreenode1:=TVORG.Items[0]; // On se place sur le premier noeud
While MyTreenode1<>NIL do
      begin
      If MyTreenode1.Text<>'' Then
         begin
         T := TR.FindFirst (['PSO_CODESERVICE'],[MyTreeNode1.Text], False);
         If T=Nil Then T := TN.FindFirst (['PSO_CODESERVICE'],[MyTreeNode1.Text], False);
         DonneeAffiche:='';
         Service:=MyTreeNode1.Text;
         Nom:=T.GetValue('PGS_NOMSERVICE');
         NiveauS:=T.GetValue('PGS_HIERARCHIE');
         Resp:=T.GetValue('PGS_RESPSERVICE');
         Etablissement:=T.GetValue('PGS_ETABLISSEMENT');
         Adjoint:=T.GetValue('PGS_ADJOINTSERVICE');
         Secretaire:=T.GetValue('PGS_SECRETAIRE');
         RespAbs:=T.GetValue('PGS_RESPONSABS');
         RespNdf:=T.GetValue('PGS_RESPONSNDF');
         RespVar:=T.GetValue('PGS_RESPONSVAR');
         AdjointAbs:=T.GetValue('PGS_ADJOINTABS');
         AdjointNdf:=T.GetValue('PGS_ADJOINTNDF');
         AdjointVar:=T.GetValue('PGS_ADJOINTVAR');
         SecretaireAbs:=T.GetValue('PGS_SECRETAIREABS');
         SecretaireNdf:=T.GetValue('PGS_SECRETAIRENDF');
         SecretaireVar:=T.GetValue('PGS_SECRETAIREVAR');
         DonneeAffiche:=Service;
         If (CNomSer<>Nil) and (CNomSer.checked=True) Then DonneeAffiche:=DonneeAffiche+'/ '+Nom;
         If (CEtab<>Nil) and (CEtab.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Etab: '+RechDom ('TTETABLISSEMENT',Etablissement,FALSE);
         If (Cniv<>Nil) and (Cniv.Checked=True) Then DonneeAffiche:=DonneeAffiche+' / Hiérarchie : '+RechDom('PGHIERARCHIE',NiveauS,False);
         if (GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE)) Or PGBundleHierarchie then   //PT3 //PT5
         begin
         If (CResp<>Nil) and (CResp.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp: '+RechDom ('PGSALARIEINT',Resp,FALSE);
         If (CSecr<>Nil) and (CSecr.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire: '+RechDom ('PGSALARIEINT',Secretaire,FALSE);
         If (CAdj<>Nil) and (CAdj.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint: '+RechDom ('PGSALARIEINT',Adjoint,FALSE);
         If (CRespAbs<>Nil) and (CRespAbs.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp Abs: '+RechDom ('PGSALARIEINT',RespAbs,FALSE);
         If (CAdjAbs<>Nil) and (CAdjAbs.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint Abs: '+RechDom ('PGSALARIEINT',AdjointAbs,FALSE);
         If (CSecrAbs<>Nil) and (CSecrAbs.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire Abs: '+RechDom ('PGSALARIEINT',SecretaireAbs,FALSE);
         If (CRespNdf<>Nil) and (CRespNdf.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp Ndf: '+RechDom ('PGSALARIEINT',RespNdf,FALSE);
         If (CAdjNdf<>Nil) and (CAdjNdf.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint Ndf: '+RechDom ('PGSALARIEINT',AdjointNdf,FALSE);
         If (CSecrNdf<>Nil) and (CSecrNdf.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire Ndf: '+RechDom ('PGSALARIEINT',SecretaireNdf,FALSE);
         If (CRespVar<>Nil) and (CRespVar.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp Var: '+RechDom ('PGSALARIEINT',RespVar,FALSE);
         If (CAdjVar<>Nil) and (CAdjVar.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint Var: '+RechDom ('PGSALARIEINT',AdjointVar,FALSE);
         If (CSecrVar<>Nil) and (CSecrVar.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire Var: '+RechDom ('PGSALARIEINT',SecretaireVar,FALSE);
         end
         else
         begin
              If (CResp<>Nil) and (CResp.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp: '+RechDom ('PGSALARIE',Resp,FALSE);
              If (CSecr<>Nil) and (CSecr.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire: '+RechDom ('PGSALARIE',Secretaire,FALSE);
              If (CAdj<>Nil) and (CAdj.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint: '+RechDom ('PGSALARIE',Adjoint,FALSE);
              If (CRespAbs<>Nil) and (CRespAbs.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp Abs: '+RechDom ('PGSALARIE',RespAbs,FALSE);
              If (CAdjAbs<>Nil) and (CAdjAbs.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint Abs: '+RechDom ('PGSALARIE',AdjointAbs,FALSE);
              If (CSecrAbs<>Nil) and (CSecrAbs.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire Abs: '+RechDom ('PGSALARIE',SecretaireAbs,FALSE);
              If (CRespNdf<>Nil) and (CRespNdf.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp Ndf: '+RechDom ('PGSALARIE',RespNdf,FALSE);
              If (CAdjNdf<>Nil) and (CAdjNdf.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint Ndf: '+RechDom ('PGSALARIE',AdjointNdf,FALSE);
              If (CSecrNdf<>Nil) and (CSecrNdf.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire Ndf: '+RechDom ('PGSALARIE',SecretaireNdf,FALSE);
              If (CRespVar<>Nil) and (CRespVar.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Resp Var: '+RechDom ('PGSALARIE',RespVar,FALSE);
              If (CAdjVar<>Nil) and (CAdjVar.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Adjoint Var: '+RechDom ('PGSALARIE',AdjointVar,FALSE);
              If (CSecrVar<>Nil) and (CSecrVar.checked=True) Then DonneeAffiche:=DonneeAffiche+' / Secrétaire Var: '+RechDom ('PGSALARIE',SecretaireVar,FALSE);
         end;
         MyTreenode1.text:=DonneeAffiche;
         end;
      MyTreenode1:=MytreeNode1.GetNext;
      end;
If T<>Nil Then T.Free;
end;

procedure TOF_ORGANIGRAMME_TV.ImpressionTV(Sender:TObject);
begin
TFVierge(Ecran).Print;
end;

//PT7 - Début
{ TOF_EDITORGANIGRAMME }

procedure TOF_EDITORGANIGRAMME.OnArgument (S : String );
Begin
	If PGBundleHierarchie Then
	Begin
		SetControlVisible('TPGS_ETABLISSEMENT', False); //PT9
		SetControlVisible('PGS_ETABLISSEMENT', False);
		
		If Not PGDroitMultiForm Then
			SetControlText('XX_WHEREPREDEF', 'PGS_NOMSERVICE LIKE "'+V_PGI.NoDossier+'%"')
		Else
		Begin
			SetControlText('XX_WHEREPREDEF', ServicesDesDossiers());
		End;
	End;
End;

//PT7 - Fin

Initialization
  registerclasses ( [ TOF_ORGANIGRAMME_TV, TOF_EDITORGANIGRAMME ] ) ; //PT7
end.
