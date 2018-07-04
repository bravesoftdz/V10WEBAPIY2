{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/03/2002
Modifié le ... :   /  /
Description .. : unit de gestion du multi critère des animateurs de stages
Suite ........ :
Mots clefs ... : PAIE;FORMATION
*****************************************************************
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT1 | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
    |            |       |    |  pour les critères code salarié uniquement
PT2 | 18/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT3 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT4 | 03/04/2008 | V_803 | FL | Adaptation partage formation
PT5 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + Gestion elipsis salarié et responsable
PT6 | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
}
unit UTofPG_MulSessionAnimat;

interface
uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,pgOutils,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,eMul,
{$ENDIF}
      Grids,HCtrls,HEnt1,vierge,EntPaie,HMsgBox,Hqry,UTOF,UTOB,UTOM,LookUp,
      AGLInit,Paramdat,PGOutilsFormation;
Type
     TOF_PGMulSessionAnimat = Class (TOF)
       private
       Bt         : TToolbarButton97;
       Stage,Millesime,Arg      : String;
       Ordre      : Integer;
       DD,DF      : TDateTime;
       Q_Mul      : THQuery ;
       {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
       procedure CreatSessAnimat (Sender: TObject);
       procedure GrilleDblClick (Sender: TObject);
       procedure DateELipsisClick(Sender : TObject);
       //{$IFDEF EMANAGER} //PT4
       procedure SalarieElipsisClick (Sender : TObject);
       //{$ENDIF}
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnLoad;  override ;
     END ;

implementation

Uses GalOutil;

procedure TOF_PGMulSessionAnimat.CreatSessAnimat(Sender: TObject);
var St : String;
begin
	st := Stage+';'+IntToStr(Ordre)+';'+DateToStr(DD)+';'+DateToStr(DF)+';'+Millesime+
	';'+'ACTION=CREATION';
	AglLanceFiche ('PAY','SESSIONANIMAT', '', '' , st);
	if Bt = NIL then Bt := TToolbarButton97 (GetControl('BCherche'));
	if Bt <> NIL then Bt.click;
end;

procedure TOF_PGMulSessionAnimat.GrilleDblClick(Sender: TObject);
var st : String;
begin
    //If Q_Mul.FindField('PAN_CODESTAGE').AsString='' Then CreatSessAnimat(Sender)
    If Q_Mul.FindField('PAN_CODESTAGE') = Nil Then CreatSessAnimat(Sender)
    Else
    begin
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
      {$ENDIF}
      st := Q_MUL.FindField('PAN_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PAN_ORDRE').AsInteger)+';'
      +IntToStr(Q_MUL.FindField('PAN_RANG').AsInteger)+';'+Q_Mul.FindField('PAN_MILLESIME').AsString;
      //      +DateToStr(Q_MUL.FindField('PAN_DATEDEBUT').AsDateTime)+';'+DateToStr(Q_MUL.FindField('PAN_DATEFIN').AsDateTime);
      AglLanceFiche ('PAY','SESSIONANIMAT', '', St , '');
      if Bt = NIL then Bt := TToolbarButton97 (GetControl('BCherche'));
      if Bt <> NIL then Bt.click;
    end;
end;

procedure TOF_PGMulSessionAnimat.OnArgument(Arguments: String);
Var BTNAgrandir,Bouv  : TToolbarButton97;
    St           : String;
//    Q : TQuery;
    Edit : THEdit;
begin
  inherited ;
	st := Arguments;
	Bt := TToolbarButton97 (GetControl('BCherche'));
	Q_Mul:=THQuery(Ecran.FindComponent('Q'));
	{$IFNDEF EAGLCLIENT}
	Grille:=THDBGrid (GetControl ('Fliste'));
	{$ELSE}
	Grille:=THGrid (GetControl ('Fliste'));
	{$ENDIF}
	if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
	Arg := ReadTokenPipe(St,';');
	If Arg = 'SESSION' then
	begin
		Stage := ReadtokenSt (st);
		if Stage <> '' then
		begin // Appel du multi critère à partir de la session donc on donne les valeurs par defaut
			Ordre := StrToInt (ReadTokenSt (st));
			DD    := StrToDate (ReadTokenSt (st));
			DF    := StrToDate (ReadTokenSt (st));
			Millesime:=ReadTokenSt(St);
			SetControlText ('PAN_CODESTAGE',Stage);
			SetControlText ('XX_WHERE', 'PAN_ORDRE ='+IntToStr (Ordre)+ ' ');
			st := DateToStr (DD);
			SetControlText ('PAN_DATEDEBUT',DateToStr(IDate1900));
			SetControlText ('PAN_DATEDEBUT_',DateToStr(IDate1900));
			SetControlText ('PAN_DATEFIN',DateToStr(IDate1900));
			SetControlText ('PAN_DATEFIN_',DateToStr(IDate1900));
			SetControlText('PAN_ORDRE',IntToStr(Ordre));
			SetControlText('PAN_MILLESIME',Millesime);
			SetControlVisible ('BInsert',TRUE);
			BTNAgrandir := TToolbarButton97 (getcontrol ('BAgrandir'));
			if BTNAgrandir <> NIL then
			begin
				BTNAgrandir.Click ;
				BTNAgrandir.Visible := FALSE;
			end;
		end
		else // tjrs vrai mais pour affecter un contenu à la clause where
		begin
			SetControlText ('XX_WHERE', '');
			SetControlVisible ('BInsert',FALSE);
		end;

	end
	else
	begin
		SetControlText ('XX_WHERE', '');
		SetControlVisible ('BInsert',FALSE);
		RendMillesimeRealise(DD,DF);
		SetControlText('PAN_DATEDEBUT',DateToStr(DD));
		SetControlText('PAN_DATEDEBUT_',DateToStr(DF));
		SetControlText('PAN_DATEFIN',DateToStr(DD));
		SetControlText('PAN_DATEFIN_',DateToStr(DF));
	end;
	BOuv := TToolbarButton97 (getcontrol ('BInsert'));
	if Bouv <> NIL then  BOuv.OnClick := CreatSessAnimat;
	Edit := THEdit(GetControl('PAN_DATEDEBUT'));
	if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
	Edit := THEdit(GetControl('PAN_DATEDEBUT_'));
	if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
	Edit := THEdit(GetControl('PAN_DATEFIN'));
	if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
	Edit := THEdit(GetControl('PAN_DATEFIN_'));
	if Edit <> Nil then Edit.OnElipsisClick := DateElipsisclick;
	//{$IFDEF EMANAGER}
	Edit := THEdit(GetControl('PAN_SALARIE'));
	{$IFNDEF EMANAGER}If PGBundleHierarchie Then{$ENDIF} //PT5
	Begin
		If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	End;
	//{$ENDIF}
	
	If PGBundleCatalogue Then //PT5
	Begin
		If not PGDroitMultiForm then
			SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
		Else If V_PGI.ModePCL='1' Then 
			SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT4
	End;
		
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PAN_PREDEFINI',False);
			SetControlText   ('PAN_PREDEFINI','');
			//SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT4 //PT5
		end
       	Else If V_PGI.ModePCL='1' Then 
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT3
			//SetControlProperty ('PAN_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT4 //PT5
		End;
	end
	else
	begin
		SetControlVisible('PAN_PREDEFINI', False);
		SetControlVisible('TPAN_PREDEFINI', False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
end;

procedure TOF_PGMulSessionAnimat.OnLoad;
var Where : String;
begin
inherited ;
	if not (Ecran is TFMul) then exit;
	{$IFDEF EMANAGER}
    Where := 'PAN_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")';
	{$ENDIF}
	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PAN_PREDEFINI'),GetControlText('NODOSSIER'),'PAN')); //PT2

    //PT6
    // N'affiche pas les salariés confidentiels
    If Where <> '' Then Where := Where + ' AND ';
    If PGBundleHierarchie Then
    	Where := Where + '(PAN_SALARIE="" OR PAN_SALARIE IN (SELECT PSI_INTERIMAIRE FROM INTERIMAIRES WHERE PSI_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))'
    Else
    	Where := Where + '(PAN_SALARIE="" OR PAN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_CONFIDENTIEL <= "'+ RendNivAccesConfidentiel()+'"))';
    SetControlText('XX_WHERE', Where);
end;

procedure TOF_PGMulSessionAnimat.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMulSessionAnimat.SalarieElipsisClick(Sender : TObject);
{$IFDEF EMANAGER}
var St,StWhere,StOrder : String;
{$ENDIF}
begin
		{$IFDEF EMANAGER}
        St := ' SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES'+
        ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE ';
        StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := ' PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
        LookupList(THEdit(Sender),'Liste des stages','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
     //PT4 - Début
     {$ELSE}
    //If PGBundleInscFormation Then //PT5
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
    	//PT4 - Fin
        {$ENDIF}
end;


Initialization
registerclasses([TOF_PGMulSessionAnimat]);
end.
