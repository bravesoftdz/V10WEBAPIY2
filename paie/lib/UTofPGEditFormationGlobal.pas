{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 30/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITFORGLOBAL ()
Mots clefs ... : TOF;PGEDITFORGLOBAL
*****************************************************************}
{
PT1 | 26/01/2007 | V_80  | FC  | Mise en place du filtage habilitation pour les lookuplist
    |            |       |     | pour les critères code salarié uniquement
PT2 | 10/01/2008 | V_802 | FLO | Report des évolutions appliquées à eManager
PT3 | 03/04/2008 | V_803 | FLO | Adaptation partage formation
PT4 | 03/04/2008 | V_803 | FLO | Gestion des groupes de travail
PT5 | 21/04/2008 | V_804 | FLO | Gestion elipsis salarié et responsable
PT6 | 04/06/2008 | V_804 | FLO | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
PT7 | 04/06/2008 | V_804 | FL | FQ 15458 Prendre en compte les non nominatifs
}
Unit UTofPGEditFormationGlobal;

Interface

Uses StdCtrls,Controls,Classes,PGOutils,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,
{$ELSE}
     eQRS1,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,PGOutilsFormation,EntPaie,PGEditOutils,PgEditOutils2,UTob,LookUp,HQRy ;

Type
  TOF_PGEDITFORGLOBAL = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Arg : String;
    procedure VerifCoche(Sender : TObject);

  end ;

    TOF_PGEDITBILANFOR = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate;override;
    procedure OnClose;override;
    private
    TobEtat : Tob;
    procedure SalarieElipsisClick(Sender: TObject);
	procedure RespElipsisClick(Sender : TObject); //PT3
  end;

Implementation

Uses GalOutil;

procedure TOF_PGEDITFORGLOBAL.OnUpdate ;
var SQL,Where : String;
    Q : TQuery;
    DateDebut,DateFin : TDateTime;
    TypeChampLibre,ChampFormation : String;
begin
  Inherited ;
        TFQRS1(Ecran).CodeEtat := 'PGE';
        SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE FROM EXERFORMATION '+Where+' ORDER BY PFE_MILLESIME';
        TypeChampLibre := '';
        If GetControlText('CRUPT1') = 'X' Then
        begin
                TypeChampLibre := 'PFM';
                ChampFormation := 'PSS_FORMATION1';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE ,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PFM" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PFM" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION1');
                SetControlText('TABLETTE',VH_Paie.FormationLibre1);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT2') = 'X' Then
        begin
                TypeChampLibre := 'PFS';
                ChampFormation := 'PSS_FORMATION2';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PFS" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PFS" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION2');
                SetControlText('TABLETTE',VH_Paie.FormationLibre2);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT3') = 'X' Then
        begin
                TypeChampLibre := 'PCF';
                ChampFormation := 'PSS_FORMATION3';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PCF" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PCF" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION3');
                SetControlText('TABLETTE',VH_Paie.FormationLibre3);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT4') = 'X' Then
        begin
                TypeChampLibre := 'PFC';
                ChampFormation := 'PSS_FORMATION4';
                If Arg <> 'REALISEMENS' Then  SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PFC" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PFC" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION4');
                SetControlText('TABLETTE',VH_Paie.FormationLibre4);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT5') = 'X' Then
        begin
                TypeChampLibre := 'PNF';
                ChampFormation := 'PSS_FORMATION5';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PNF" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PNF" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION5');
                SetControlText('TABLETTE',VH_Paie.FormationLibre5);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT6') = 'X' Then
        begin
                TypeChampLibre := 'PPC';
                ChampFormation := 'PSS_FORMATION6';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PPC" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PPC" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION6');
                SetControlText('TABLETTE',VH_Paie.FormationLibre6);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT7') = 'X' Then
        begin
                TypeChampLibre := 'PFO';
                ChampFormation := 'PSS_FORMATION7';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PFO" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PFO" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION7');
                SetControlText('TABLETTE',VH_Paie.FormationLibre7);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPT8') = 'X' Then
        begin
                TypeChampLibre := 'PFM';
                ChampFormation := 'PSS_FORMATION8';
                If Arg <> 'REALISEMENS' Then SQL := 'SELECT PFE_MILLESIME,PFE_LIBELLE,CC_CODE,(CC_LIBELLE) LIBELLE FROM EXERFORMATION'+
                ' LEFT JOIN CHOIXCOD ON CC_TYPE="PLF" '+Where+' ORDER BY PFE_MILLESIME,CC_CODE'
                Else SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PLF" ORDER BY CC_CODE';
                SetControlText('CHAMP','PST_FORMATION8');
                SetControlText('TABLETTE',VH_Paie.FormationLibre8);
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPTSTAGE') = 'X' Then
        begin
                SQL := 'SELECT PST_CODESTAGE,(PST_LIBELLE) LIBELLE FROM STAGE WHERE PST_MILLESIME="0000" ORDER BY PST_CODESTAGE';
                SetControlText('CHAMP','PST_CODESTAGE');
                SetControlText('TABLETTE','Formation');
                SetControlText('XX_RUPTURE1','PST_CODESTAGE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPTETAB') = 'X' Then
        begin
                SQL := 'SELECT ET_ETABLISSEMENT,(ET_LIBELLE) LIBELLE FROM ETABLISS ORDER BY ET_ETABLISSEMENT';
                SetControlText('CHAMP','PFO_ETABLISSEMENT');
                SetControlText('TABLETTE','Etablissement');
                SetControlText('XX_RUPTURE1','ET_ETABLISSEMENT');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        If GetControlText('CRUPTEMPLOI') = 'X' Then
        begin
                SQL := 'SELECT CC_CODE,(CC_LIBELLE) LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PLE" ORDER BY CC_CODE';
                SetControlText('CHAMP','PFO_LIBEMPLOIFOR');
                SetControlText('TABLETTE','Libellé emploi');
                SetControlText('XX_RUPTURE1','CC_CODE');
                TFQRS1(Ecran).CodeEtat := 'PCG';
        end;
        if Arg = 'REALISEMENS' Then
        begin
                DateDebut := IDate1900;
                DateFin := IDate1900;
                Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="'+GetControltext('PFE_MILLESIME')+'"',True);
                If Not Q.Eof then
                begin
                        DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime;
                        DateFin := Q.FindField('PFE_DATEFIN').AsDatetime;
                end;
                Ferme(Q);
                If GetControlText('CRUPTSTAGE') = 'X' Then
                begin
                        TFQRS1(Ecran).CodeEtat := 'PRF';
                        SQL := 'SELECT DISTINCT PST_CODESTAGE FROM STAGE LEFT JOIN '+
                        'SESSIONSTAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME '+
                        'WHERE PSS_EFFECTUE="X" AND PSS_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'"';
                end
                Else
                begin
                        SQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="'+TypeChampLibre+'" AND CC_CODE IN (SELECT '+ChampFormation+' '+
                        'FROM SESSIONSTAGE WHERE PSS_EFFECTUE="X" AND PSS_DATEDEBUT>="'+UsDateTime(DateDebut)+'" AND PSS_DATEDEBUT<="'+UsDateTime(DateFin)+'")';
                        TFQRS1(Ecran).CodeEtat := 'PRA';
                end;
        end;
        TFQRS1(Ecran).WhereSQL := SQL;
end ;

procedure TOF_PGEDITFORGLOBAL.OnArgument (S : String ) ;
var Num : Integer;
    Min,Max : String;
    Combo : THValComboBox;
    Check : TCheckBox;
begin
  Inherited ;
        Arg := S;
        If Arg = 'REALISEMENS' Then
        begin
                TFQRS1(Ecran).Caption := 'Suivi financier : Coût réel formation';
                SetControlVisible('PFE_MILLESIME_',False);
                SetControlVisible('TPFE_MILLESIME_',False);
                SetControlChecked('CRUPTSTAGE',True);
                Check := TCheckBox(GetControl('CRUPTSTAGE'));
                If Check <> Nil Then Check.OnClick := VerifCoche;
        end
        Else TFQRS1(Ecran).Caption := 'Comparatif global prévisionnel/réalisé';
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        For Num := 1 To 8 do
        begin
                Combo := THValComboBox(GetControl('FORMATION'+IntToStr(Num)));
                If Combo.Visible=True Then SetControlVisible('CRUPT'+IntToStr(Num),True);
                Check := TCheckBox(GetControl('CRUPT'+IntToStr(Num)));
                If Check <> Nil then Check.OnClick := VerifCoche;
        end;
        Check := TCheckBox(GetControl('CRUPTSTAGE'));
        If Check <> Nil then Check.OnClick := VerifCoche;
        RecupMinMaxTablette('PG','EXERFORMATION','PFE_MILLESIME',Min,Max);
        SetControlProperty('PFE_MILLESIME','Value',Min);
        SetControlProperty('PFE_MILLESIME_','Value',Max);
end ;

procedure TOF_PGEDITFORGLOBAL.VerifCoche(Sender : TObject);
var i : Integer;
begin
        If TCheckBox(Sender).Name = 'CRUPTSTAGE' Then
        begin
                If GetCheckBoxState('CRUPTSTAGE')=CbChecked Then
                begin
                        For i := 1 to 8 do
                        begin
                                SetControlChecked('CRUPT'+IntToStr(i),False);
                        end;
                end;
        end
        Else
        begin
                If GetCheckBoxState(TCheckBox(Sender).Name)=CbChecked then
                begin
                        SetControlChecked('CRUPTSTAGE',False);
                        For i := 1 to 8 do
                        begin
                                If TCHeckBox(Sender).Name <> ('CRUPT'+IntToStr(i)) Then SetControlChecked('CRUPT'+IntToStr(i),False);
                        end;
                end;
        end;
end;


procedure TOF_PGEDITBILANFOR.SalarieElipsisClick(Sender: TObject);
var
  StFrom, StWhere: string;
begin
    If PGBundleHierarchie Then //PT3
    	ElipsisSalarieMultidos (Sender)
    Else
    Begin
  		StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  		StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
  		StFrom  := 'SALARIES';
  		LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
  	End;
end;

procedure TOF_PGEDITBILANFOR.OnArgument(Arguments: string);
var Edit : THEdit;
//    Check : TcheckBox;
    DateD,DateF : TDateTime;
begin
  inherited;
	//PT3
	If PGBundleHierarchie Then //PT5
	Begin
		Edit := THEdit(GetControl('SALARIE'));
		if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
	
		Edit := THEdit(GetControl('PFI_RESPONSFOR'));
		if Edit <> nil then Edit.OnElipsisClick := RespElipsisClick;
	End;
	
	TFQRS1(Ecran).Caption := 'Edition du bilan individuel de formation';
	UpdateCaption(Ecran);
	SetControlText('PFI_MILLESIME',RendMillesimePrevisionnel);
	RendMillesimeRealise(DateD,DateF);
	SetControlText('DATEDEB',DateToStr(DateD));
	SetControlText('DATEFIN',DateToStr(DateF));
     
	//PT3 - Début
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PFI_PREDEFINI',False);
			SetControlText   ('PFI_PREDEFINI','');
		end
       	Else If V_PGI.ModePCL='1' Then SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT4
	end
	else
	begin
		SetControlVisible('PFI_PREDEFINI', False);
		SetControlVisible('TPFI_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT3 - Fin  
end;

procedure TOF_PGEDITBILANFOR.OnUpdate;
var Q : TQuery;
    TobRea,TobPrev,T : Tob;
    i : Integer;
    WhereRea,WherePrev : String;
    Pages : TPageControl;
    SDateD,SDateF : String;
    Millesime : String;
    Year, Month, Day, Year2, Month2, Day2 : Word; //PT2
    DD,DF : TDateTime; //PT2
    St,Prefixe : String; //PT2
begin
     inherited;
     
     //PT3
     If PGBundleHierarchie Then
     	Prefixe := 'PSI'
     Else
     	Prefixe := 'PSA';
     
     Millesime := GetControlText('PFI_MILLESIME');
     SDateD := GetcontrolText('DATEDEB');
     SDateF := GetcontrolText('DATEFIN');
     Pages := TPageControl(GetControl('Pages'));
     WherePrev := RecupWhereCritere(Pages);
     WhereRea := 'WHERE PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="'+UsdateTime(StrToDate(GetcontrolText('DATEDEB')))+'" AND PFO_DATEDEBUT<="'+UsdateTime(StrToDate(GetcontrolText('DATEFIN')))+'"';
     if GetControlText('SALARIE') <> '' then
     begin
          If WherePrev <> '' then WherePrev := WherePrev+' AND PFI_SALARIE="'+GetControltext('SALARIE')+'"'
          else WherePrev := 'WHERE PFI_SALARIE="'+GetControltext('SALARIE')+'"';
          Whererea := Whererea+' AND PFO_SALARIE="'+GetControltext('SALARIE')+'"'
     end;
     if GetControlText('LIBEMPLOI') <> '' then
     begin
          If WherePrev <> '' then WherePrev := WherePrev+' AND '+Prefixe+'_LIBELLEEMPLOI="'+GetControltext('LIBEMPLOI')+'"'
          else WherePrev := 'WHERE '+Prefixe+'_LIBELLEEMPLOI="'+GetControltext('LIBEMPLOI')+'"';
          Whererea := Whererea+' AND '+Prefixe+'_LIBELLEEMPLOI="'+GetControltext('LIBEMPLOI')+'"'
     end;
     If WherePrev <> '' then WherePrev := WherePrev+' AND PFI_SALARIE<>""'
     else WherePrev := 'WHERE PFI_SALARIE<>""';
     If GetControlText('PFI_RESPONSFOR') <> '' then
     begin
          WherePrev := WherePrev + ' AND PFI_RESPONSFOR="'+GetControlText('PFI_RESPONSFOR')+'"';
          WhereRea := WhereRea + ' AND PSE_RESPONSFOR="'+GetControlText('PFI_RESPONSFOR')+'"';
     end;
     //PT2 - Début
     //Gestion de la date d'entrée
     If (GetControlText('DATEENTREE') <> '') And (GetControlText('DATEENTREE') <> '  /  /    ')  Then
     Begin
          WherePrev := WherePrev + ' AND PFI';
          WhereRea  := WhereRea  + ' AND PFO';

			If PGBundleHierarchie Then //PT3
			Begin
          		St := '_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_DATEENTREE>="'+UsDateTime(StrToDate(GetControlText('DATEENTREE')))+'"';
          		If (GetControlText('DATEENTREE_') <> '') And (GetControlText('DATEENTREE_') <> '  /  /    ') Then St := St + ' AND PSI_DATEENTREE<="'+UsDateTime(StrToDate(GetControlText('DATEENTREE_')))+'"';
          	End
			Else
			Begin
          		St := '_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_DATEENTREE>="'+UsDateTime(StrToDate(GetControlText('DATEENTREE')))+'"';
          		If (GetControlText('DATEENTREE_') <> '') And (GetControlText('DATEENTREE_') <> '  /  /    ') Then St := St + ' AND PSA_DATEENTREE<="'+UsDateTime(StrToDate(GetControlText('DATEENTREE_')))+'"';
          	End;
          St := St + ')';

          WherePrev := WherePrev  + St;
          WhereRea := WhereRea  + St;
     End;
     //PT2 - Fin
     
     //PT6
	 // N'affiche pas les salariés confidentiels
	 If WherePrev <> '' Then WherePrev := WherePrev + ' AND ';
	 If Whererea <> '' Then Whererea := Whererea + ' AND ';
	 If PGBundleHierarchie Then
	 Begin
	 	WherePrev := WherePrev + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))'; //PT7
	 	Whererea  := Whererea  + 'PFO_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
	 End
	 Else
	 Begin
	 	WherePrev := WherePrev + '(PFI_SALARIE="" OR PFI_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))';  //PT7
	 	Whererea  := Whererea  + 'PFO_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'")';
	 End;
	 
	 
     // Prévisionnel
     If TobEtat<> Nil then FreeAndNil(TobEtat);
     TobEtat := Tob.Create('Edition',Nil,-1);
     
     St := 'SELECT PSE_RESPONSFOR,PFI_SALARIE,'+Prefixe+'_LIBELLE,'+Prefixe+'_PRENOM,PFI_MILLESIME,PFI_CODESTAGE,PFI_DUREESTAGE,PFI_ETATINSCFOR,'+ //PT2
      'PFI_NATUREFORM,PFI_TYPEPLANPREV,'+Prefixe+'_LIBELLEEMPLOI,'+Prefixe+'_DATEENTREE,'+Prefixe+'_DATENAISSANCE,'+Prefixe+'_ETABLISSEMENT,PFI_REALISE FROM INSCFORMATION ';
     If PGBundleHierarchie Then  //PT3
     	St := St + 'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFI_SALARIE'
     Else
     	St := St + 'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE';
     St := St + ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+WherePrev+' AND PFI_CODESTAGE<>"--CURSUS--"';
     If PGBundleHierarchie Then St := St + DossiersAInterroger(GetControlText('PFI_PREDEFINI'),GetControlText('NODOSSIER'),'PSI', False, True); //PT3
     St := St + ' ORDER BY PFI_SALARIE ASC,PFI_MILLESIME DESC';
      
     Q := OpenSQL(St,True);
      TobPrev := Tob.Create('leprevisionnel', nil, -1);
      TobPrev.LoadDetailDB('leprevisionnel', '', '', Q, False);
      Ferme(Q);
      for i := 0 to TobPrev.detail.Count - 1 do
      begin
       T := Tob.Create('filleEdition',TobEtat,-1);
       T.AddChampSupValeur('PFO_SALARIE',TobPrev.Detail[i].GetValue('PFI_SALARIE'));
       T.AddChampSupValeur('PFO_MILLESIME', TobPrev.Detail[i].GetValue('PFI_MILLESIME')); //PT2
       T.AddChampSupValeur('PFO_TYPEPLANPREV','PREV');
       T.AddChampSupValeur('DATEFOR',IDate1900);
       T.AddChampSupValeur('LIBSALARIE',TobPrev.Detail[i].GetValue(Prefixe+'_LIBELLE') + ' '+ TobPrev.Detail[i].GetValue(Prefixe+'_PRENOM'));
       T.AddChampSupValeur('LIBSTAGE',RechDom('PGSTAGEFORMCOMPLET',TobPrev.Detail[i].GetValue('PFI_CODESTAGE'),False));
       T.AddChampSupValeur('LIBELLETYPE','Prévisionnel '+Millesime);
       T.AddChampSupValeur('NBHEURES',TobPrev.Detail[i].GetValue('PFI_DUREESTAGE'));
       T.AddChampSupValeur('NATURE',RechDom('PGNATUREFORM',TobPrev.Detail[i].GetValue('PFI_NATUREFORM'),False));
       T.AddChampSupValeur('LIBEMPLOI',RechDom('PGLIBEMPLOI',TobPrev.Detail[i].GetValue(Prefixe+'_LIBELLEEMPLOI'),False));
       T.AddChampSupValeur('DATEENTREE',TobPrev.Detail[i].GetValue(Prefixe+'_DATEENTREE'));
       T.AddChampSupValeur('DATENAISSANCE',TobPrev.Detail[i].GetValue(Prefixe+'_DATENAISSANCE'));
       T.AddChampSupValeur('ETABLISSEMENT',RechDom('TTETABLISSEMENT',TobPrev.Detail[i].GetValue(Prefixe+'_ETABLISSEMENT'),False));
       T.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',TobPrev.Detail[i].GetValue('PSE_RESPONSFOR'),False));
	   T.AddChampSupValeur('TYPEPLANPREV',TobPrev.Detail[i].GetValue('PFI_TYPEPLANPREV')); //PT2
       T.AddChampSupValeur('ETATINSC',RechDom('PGETATVALIDATION',TobPrev.Detail[i].GetValue('PFI_ETATINSCFOR'),False)); //PT2
       If TobPrev.Detail[i].GetValue('PFI_REALISE') = 'X' then T.AddChampSupValeur('ETAT','Oui')
       else T.AddChampSupValeur('ETAT','Non');
      end;
      TobPrev.Free;
      
      // Réalisé
     St := 'SELECT PFO_DATEDEBUT,PSE_RESPONSFOR,PFO_SALARIE,'+Prefixe+'_LIBELLE,'+Prefixe+'_PRENOM,PFO_LIBEMPLOIFOR,PFO_CODESTAGE,PFO_NBREHEURE,PFO_TYPEPLANPREV,PFO_ETATINSCFOR,'+//PT2
      'PFO_NATUREFORM,PFO_TYPEPLANPREV,'+Prefixe+'_LIBELLEEMPLOI,'+Prefixe+'_DATEENTREE,'+Prefixe+'_DATENAISSANCE,'+Prefixe+'_ETABLISSEMENT FROM FORMATIONS ';
     If PGBundleHierarchie Then  //PT3
     	St := St + 'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFO_SALARIE'
     Else
     	St := St + 'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE';
     St := St + ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE '+WhereRea;
     If PGBundleHierarchie Then St := St + DossiersAInterroger(GetControlText('PFI_PREDEFINI'),GetControlText('NODOSSIER'),'PSI',False,True); //PT3
     St := St + ' ORDER BY PFO_DATEDEBUT DESC';
     
      Q := OpenSQL(St,True); //PT2
      TobRea := Tob.Create('Lerealise', nil, -1);
      TobRea.LoadDetailDB('Lerealise', '', '', Q, False);
      Ferme(Q);
      for i := 0 to TobRea.detail.Count - 1 do
      begin
       T := Tob.Create('filleEdition',TobEtat,-1);
       T.AddChampSupValeur('PFO_SALARIE',TobRea.Detail[i].GetValue('PFO_SALARIE'));
       //PT2 - Début
       DecodeDate(TobRea.Detail[i].GetValue('PFO_DATEDEBUT'), Year, Month, Day);
       T.AddChampSupValeur('PFO_MILLESIME', IntToStr(Year));
       DecodeDate(StrToDate(sDateD), Year2, Month2, Day2);DD := EncodeDate(Year, Month2, Day2);
       DecodeDate(StrToDate(sDateF), Year2, Month2, Day2);DF := EncodeDate(Year, Month2, Day2);
       //PT2 - Fin
       T.AddChampSupValeur('PFO_TYPEPLANPREV','REA');
       T.AddChampSupValeur('DATEFOR',TobRea.Detail[i].GetValue('PFO_DATEDEBUT'));
       T.AddChampSupValeur('LIBSALARIE',TobRea.Detail[i].GetValue(Prefixe+'_LIBELLE') +' '+TobRea.Detail[i].GetValue(Prefixe+'_PRENOM'));
       T.AddChampSupValeur('LIBSTAGE',RechDom('PGSTAGEFORMCOMPLET',TobRea.Detail[i].GetValue('PFO_CODESTAGE'),False));
       T.AddChampSupValeur('LIBELLETYPE','Réalisé du '+DateToStr(DD)+ ' au '+DateToStr(DF)); //PT2
       T.AddChampSupValeur('NBHEURES',TobRea.Detail[i].GetValue('PFO_NBREHEURE'));
       T.AddChampSupValeur('NATURE',RechDom('PGNATUREFORM',TobRea.Detail[i].GetValue('PFO_NATUREFORM'),False));
       T.AddChampSupValeur('LIBEMPLOI',RechDom('PGLIBEMPLOI',TobRea.Detail[i].GetValue(Prefixe+'_LIBELLEEMPLOI'),False));
       T.AddChampSupValeur('DATEENTREE',TobRea.Detail[i].GetValue(Prefixe+'_DATEENTREE'));
       T.AddChampSupValeur('DATENAISSANCE',TobRea.Detail[i].GetValue(Prefixe+'_DATENAISSANCE'));
       T.AddChampSupValeur('ETABLISSEMENT',RechDom('TTETABLISSEMENT',TobRea.Detail[i].GetValue(Prefixe+'_ETABLISSEMENT'),False));
       T.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',TobRea.Detail[i].GetValue('PSE_RESPONSFOR'),False));
       T.AddChampSupValeur('TYPEPLANPREV',TobRea.Detail[i].GetValue('PFO_TYPEPLANPREV')); //PT2
       T.AddChampSupValeur('ETATINSC',TobRea.Detail[i].GetValue('PFO_ETATINSCFOR')); //PT2
       T.AddChampSupValeur('ETATINSC',RechDom('PGETATVALIDATION',TobRea.Detail[i].GetValue('PFO_ETATINSCFOR'),False)); //PT2
       //T.AddChampSupValeur('PFO_TYPEPLANPREV','REA'); //PT2
       //T.AddChampSupValeur('LIBELLETYPE','Réalisé du '+SDateD+ ' au '+SDateF);//PT2
       St := 'SELECT 1 FROM INSCFORMATION ';
       If PGBundleHierarchie Then //PT3
       	St := St + 'LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PFI_SALARIE '
       Else
       	St := St + 'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE ';
       St := St + ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+WherePrev+
             ' AND PFI_CODESTAGE="'+TobRea.Detail[i].GetValue('PFO_CODESTAGE')+'" AND PFI_SALARIE="'+TobRea.Detail[i].GetValue('PFO_SALARIE')+'"';
       If ExisteSQL(St) then 
      	T.AddChampSupValeur('ETAT','Oui')
      else 
      	T.AddChampSupValeur('ETAT','Non');
      end;
      TobRea.Free;
      //TobEtat.Detail.Sort('PFO_SALARIE;PFO_TYPEPLANPREV;DATEFOR');
      TobEtat.Detail.Sort('PFO_SALARIE;PFO_MILLESIME;PFO_TYPEPLANPREV;DATEFOR'); //PT2
      TFQRS1(Ecran).LaTob:= TobEtat;
end;

procedure TOF_PGEDITBILANFOR.OnClose;
begin
     inherited;
     If TobEtat<> Nil then FreeAndNil(TobEtat);
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT3
Modifié le ... :   /  /
Description .. : Clic elipsis Responsable
Mots clefs ... :
*****************************************************************}
procedure TOF_PGEDITBILANFOR.RespElipsisClick(Sender : TObject);
begin
        ElipsisResponsableMultidos (Sender); 
end;

Initialization
  registerclasses ( [ TOF_PGEDITFORGLOBAL,TOF_PGEDITBILANFOR ] ) ;
end.
