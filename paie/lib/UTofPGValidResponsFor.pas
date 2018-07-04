{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 14/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGVALIDRESPONSFOR ()
Mots clefs ... : TOF;PGVALIDRESPONSFOR
*****************************************************************
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2 | 19/09/2007 | V_80  | FL | Adaptation des requêtes pour multidossiers
PT3 | 15/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT4 | 02/04/2008 | V_80  | FL | Adaptation des requêtes pour multidossiers et groupes de travail
PT5 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 - Ne pas voir les salariés confidentiels
PT6 | 17/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}
Unit UTofPGValidResponsFor;

Interface

Uses Classes,                    
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     Utob,       
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOF,uTableFiltre,SaisieList,P5Def,EntPaie,LookUp,PGOutilsFormation ;

Type
  TOF_PGVALIDRESPONSFOR = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    TF:TTableFiltre;
    TypeSaisie : String;
    AutoValid : Boolean;
    procedure ClickBCherche(Sender:TObject);
    procedure RespElipsisClick(Sender:TObject);
  end ;

Implementation   
Uses ParamSoc, GalOutil, PGHierarchie;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Chargement des données de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGVALIDRESPONSFOR.OnLoad ;
var i : Integer;
    StWhere,Nomcol : String;
begin
  Inherited ;
     StWhere := '';
     //TF.LaGrid.ColFormats[1] := 'CB=PGSTAGEFORM';
     {$IFNDEF EMANAGER}
     If GetControlText('NIVEAUH') <> '' then StWhere := ' WHERE PGS_HIERARCHIE="'+GetControltext('NIVEAUH')+'"';
     If GetControlText('SERVICE') <> '' then
     begin
          If StWhere <> '' then StWhere := StWhere + ' AND PGS_CODESERVICE="'+GetControltext('SERVICE')+'"'
          else StWhere := ' WHERE PGS_CODESERVICE="'+GetControltext('SERVICE')+'"'
     end;

     //PT2 - Début
     // Si accès multi : Jointure sur la table Intérimaires
     If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
     Begin
          SetControlText('XX_WHERE', 'PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES'+StWhere+')'+
          //PT5
		  // N'affiche pas les salariés confidentiels
		  ' AND (PSI_CONFIDENTIEL IS NULL OR PSI_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")') //PT6
     End
     Else
     Begin
          SetControlText('XX_WHERE','PSA_SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES'+StWhere+')'+
		  //PT5
		  // N'affiche pas les salariés confidentiels
		  ' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")'); //PT6
     End;
     //PT2 - Fin

     {$ELSE}
     If GetControlText('NIVEAUH') <> '' then StWhere := ' AND PGS_HIERARCHIE="'+GetControltext('NIVEAUH')+'"';
     If GetControlText('SERVICE') <> '' then StWhere := StWhere + ' AND PGS_CODESERVICE="'+GetControltext('SERVICE')+'"';
	//PT5
	// N'affiche pas les salariés confidentiels
	StWhere := StWhere + ' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")';  //PT6

     If AutoValid then
     SetControlText('XX_WHERE','(PSA_SALARIE="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'" '+StWhere+'))'+
			//PT5
			// N'affiche pas les salariés confidentiels
			' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"))')  //PT6
     else SetControlText('XX_WHERE','PSA_SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'" '+StWhere+'))'+
			//PT5
			// N'affiche pas les salariés confidentiels
			' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")');  //PT6
     {$ENDIF}
     TFSaisieList(Ecran).BCherche.Click;
     For i := 1 to TF.LaGrid.ColCount - 1 do
     begin
          NomCol := TF.LaGrid.ColNames[i];
          If NomCol = 'NATURE' then
          begin
               TF.ChangeColFormat('NATURE','C');
          end;
     end;
     
     SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSI_PREDEFINI'),GetControlText('NODOSSIER'),'PSI')); //PT3
end ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... :   /  /    
Modifié le ... :   /  /    
Description .. : Ouverture de l'écran
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGVALIDRESPONSFOR.OnArgument (S : String ) ;
var Num:Integer;
    Resp:THEdit;
    MillesimeEC:String;
    NomChamp : String; //PT2
begin
  Inherited ;

  TypeSaisie := ReadTokenPipe(S,';');
  TF := TFSaisieList(Ecran).LeFiltre;
  TF.LaGrid.ColFormats[0] := 'CB=PGSTAGEFORM';
  SetControlvisible('BINSERT',False);
  TFSaisieList(Ecran).LignesMessage := 5;
  If TypeSaisie = 'VALIDRESPCURSUS' then TFSaisieList(Ecran).ZonesMessages := 1
  else TFSaisieList(Ecran).ZonesMessages := 2;
  //BCherche:=TToolBarButton97(GetControl('BCHERCHE'));
  //If BCHerche<>Nil then BCherche.OnClick:=ClickBCherche;
  TFSaisieList(Ecran).BCherche.OnClick := ClickBCherche;

  //PT2 - Début
  If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
  Begin
     NomChamp := 'PSI';
     // Renommage des critères pour s'adapter à la table Intérimaires
     THValComboBox(GetControl('PSA_TRAVAILN1')).Name := 'PSI_TRAVAILN1';
     THValComboBox(GetControl('PSA_TRAVAILN2')).Name := 'PSI_TRAVAILN2';
     THValComboBox(GetControl('PSA_TRAVAILN3')).Name := 'PSI_TRAVAILN3';
     THValComboBox(GetControl('PSA_TRAVAILN4')).Name := 'PSI_TRAVAILN4';
     THValComboBox(GetControl('PSA_CODESTAT')).Name  := 'PSI_CODESTAT';

     TF.TableEntete := 'INTERIMAIRES';
     TF.WhereTable  := 'WHERE PFI_RESPONSFOR=:PSI_INTERIMAIRE';
     TF.LaTreeListe := 'PGTVRESPONSFORINT';
  End
  Else
     NomChamp := 'PSA';

  // Ajout du prédéfini
  If PGBundleInscFormation then
  begin
      If not PGDroitMultiForm then
      begin
          SetControlEnabled('NODOSSIER',False);
          SetControlText   ('NODOSSIER',V_PGI.NoDossier);
          SetControlEnabled('PSI_PREDEFINI',False);
          SetControlText   ('PSI_PREDEFINI','');
      end
      	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT4
  end
  else
  begin
      SetControlVisible('PSI_PREDEFINI', False);
      SetControlVisible('TPSI_PREDEFINI',False);
      SetControlVisible('NODOSSIER',     False);
      SetControlVisible('TNODOSSIER',    False);
  end;
	//PT4 - Début
  
  If PGBundleHierarchie Then
  Begin
		If Not PGDroitMultiForm Then
			SetControlProperty('SERVICE', 'Plus', 'PGS_NOMSERVICE LIKE "'+V_PGI.NoDossier+'%"')
		Else
		Begin
			SetControlProperty('SERVICE', 'Plus', ServicesDesDossiers());
		End;
  End;
	//PT4 - Fin
  //PT2 - Fin

  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
     if Num >4 then Break;
       VisibiliteChampSalarie (IntToStr(Num),GetControl (NomChamp+'_TRAVAILN'+IntToStr(Num)),GetControl ('T'+NomChamp+'_TRAVAILN'+IntToStr(Num))); //PT2
  end;
  VisibiliteStat (GetControl (NomChamp+'_CODESTAT'),GetControl ('T'+NomChamp+'_CODESTAT')) ; //PT2
  SetControlCaption('LIBSERVICE','');

  Resp:=THEdit(GetControl('PSA_SALARIE'));
  If Resp<>Nil Then
  Begin
       If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then //PT2
       Begin
          // Adaptation de la combo de sélection du responsable 
          Resp.Name := 'PSI_INTERIMAIRE';
          Resp.DataType := 'PGINTERIMAIRES';
       End;
       Resp.OnElipsisClick:=RespElipsisClick;
  End;
  SetControlCaption('LIBRESP','');
  MillesimeEC:=RendMillesimePrevisionnel;

  If TypeSaisie = 'VALIDRESPCURSUS' then
     TF.WhereTable:=TF.WhereTable+' AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CURSUS<>"" AND PFI_CODESTAGE="--CURSUS--"'
  else
     TF.WhereTable:=TF.WhereTable+' AND PFI_MILLESIME="'+MillesimeEC+'" AND PFI_CURSUS=""';


  //TF.LaGrid.ColFormats[5]:='CB=TTETABLISSEMENT';
  //TF.LaGrid.ColFormats[6]:='CB=PGLIBEMPLOI';
  //TF.LaGrid.ColFormats[9]:='CB=PGREFUSFORM';
  //TF.LaGrid.ColFormats[11]:='CB=PGREPORTFORM';
  SetControlCaption('PanTitre','Liste des responsables');
  SetControlProperty('PComplement','TabVisible',False);
  SetControlProperty('PComplement','Visible',False);

  {$IFDEF EMANAGER}
  TFSaisieList(Ecran).ParamTreeView := False;
  {$ENDIF}

  If ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
        'WHERE PHO_NIVEAUH<=2 AND PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"') then AutoValid := True
  else AutoValid := False;
end ;

procedure TOF_PGVALIDRESPONSFOR.ClickBCherche(Sender:TObject);
var StWhere:String;
begin
     StWhere := '';
     //TF.LaGrid.ColFormats[1] := 'CB=PGSTAGEFORM';
     {$IFNDEF EMANAGER}
     If GetControlText('NIVEAUH') <> '' then StWhere := ' WHERE PGS_HIERARCHIE="'+GetControltext('NIVEAUH')+'"';
     If GetControlText('NIVEAUH') <> '' then
     begin
        If StWhere <> '' then StWhere := StWhere + ' AND PGS_CODESERVICE="'+GetControltext('SERVICE')+'"'
        else StWhere := ' WHERE PGS_CODESERVICE="'+GetControltext('SERVICE')+'"'
     end;

     If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then //PT2
          SetControlText('XX_WHERE', 'PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES'+StWhere+')'+
			//PT5
			// N'affiche pas les salariés confidentiels
			' AND (PSI_CONFIDENTIEL IS NULL OR PSI_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")')  //PT6
     Else
          SetControlText('XX_WHERE','PSA_SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES'+StWhere+')'+
			//PT5
			// N'affiche pas les salariés confidentiels
			' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")');  //PT6

     {$ELSE}
     If GetControlText('NIVEAUH') <> '' then StWhere := ' AND PGS_HIERARCHIE="'+GetControltext('NIVEAUH')+'"';
     If GetControlText('NIVEAUH') <> '' then StWhere := StWhere + ' AND PGS_CODESERVICE="'+GetControltext('SERVICE')+'"';
     If AutoValid then
     SetControlText('XX_WHERE','(PSA_SALARIE="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'" '+StWhere+'))'+
			//PT5
			// N'affiche pas les salariés confidentiels
			' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"))')  //PT6
     else SetControlText('XX_WHERE','PSA_SALARIE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'" '+StWhere+'))'+
			//PT5
			// N'affiche pas les salariés confidentiels
			' AND (PSA_CONFIDENTIEL IS NULL OR PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'")');  //PT6
     {$ENDIF}
     TFSaisieList(Ecran).BChercheClick(Sender);
end;

procedure TOF_PGVALIDRESPONSFOR.RespElipsisClick(Sender:TObject);
var StWhere : String;
    StOrder : String;
begin
        If AutoValid then
        begin
          //PT2 - Début
          If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
          Begin
               If Not PGDroitMultiForm Then //PT4
               	StWhere := DossiersAInterroger('',V_PGI.NoDossier,'PSI',False,False)
               Else
                 StWhere := DossiersAInterroger('','','PSI',False,False);
               StWhere := StWhere + ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN ';
               StOrder := 'PSI_LIBELLE';
          End
          Else
          Begin
                StWhere := '(PSA_SALARIE="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN ';
                StOrder := 'PSA_LIBELLE';
          End;
          //PT2 - Fin

          StWhere := StWhere + '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                     ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
        end
        else
        begin
          //PT2 - Début
          If ((GetParamSocSecur('SO_PGINTERVENANTEXT',False)) or (PGBundleHierarchie)) Then
          Begin              	
               If Not PGDroitMultiForm Then //PT4
               	StWhere := DossiersAInterroger('',V_PGI.NoDossier,'PSI',False,False)
               Else
                 StWhere := DossiersAInterroger('','','PSI',False,False);
               StWhere := StWhere + ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN ';
               StOrder := 'PSI_LIBELLE';
          End
          Else
          Begin
               StWhere := '(PSA_SALARIE IN ';
               StOrder := 'PSA_LIBELLE';
          End;
          //PT2 - Fin

          StWhere := StWhere + '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
                     ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
        end;

        //PT2 - Début
        If (PGBundleHierarchie) Then
        	ElipsisResponsableMultiDos(Sender, StWhere)
        Else If (GetParamSocSecur('SO_PGINTERVENANTEXT',False)) Then
             LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1)
        Else
             LookupList(THEdit(Sender),'Liste des responsables','SALARIES','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
        //PT2 - Fin
end;


Initialization
  registerclasses ( [ TOF_PGVALIDRESPONSFOR ] ) ;
end.
