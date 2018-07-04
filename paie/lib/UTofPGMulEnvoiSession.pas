{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGENVOISESSION_MUL ()
Mots clefs ... : TOF;PGENVOISESSION_MUL
*****************************************************************
PT1  | 25/01/2005 | V_60  | JL | Affichage seulement des sessions inclus 2483
---  | 20/03/2006 |       | JL | Modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2  | 25/04/2008 | V_804 | FL | Adaptation pour partage formation
}
Unit UTofPGMulEnvoiSession ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,DBGrids,HDB,DBCtrls,Mul,
{$ELSE}
     emul,MainEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,HQry,UTOB,ParamDat,PGOutilsFormation;

Type
  TOF_PGENVOISESSION_MUL = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    MillesimeConsult,Arg:String;
    NumeroEnvoi:Integer;
    procedure DateElipsisclick(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
    procedure PreparerEnvoi(Sender : TObject);
    procedure AjouterSession(Sender : TObject);
    procedure EnleverSession(Sender : TObject);
  end ;
var     {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    TFM : TFMul;
    {$ENDIF}
    Q_Mul : THQuery;

Implementation

Uses galOutil;

procedure TOF_PGENVOISESSION_MUL.OnLoad ;
var Where : String;
begin
  Inherited ;
    Where := '';
    
    if GetControlText('VSESSIONS') = 'NE' then Where := 'PSS_NUMENVOI<=0'  //DB2
    Else Where := '';
    If Arg = 'SUPP' then Where := 'PSS_NUMENVOI=' + IntToStr(NumeroEnvoi) + '';  //DB2
    If Arg = 'AJOUT' then Where := 'PSS_NUMENVOI<=0';  //DB2
    If Where <> '' then Where := Where + ' AND PSS_INCLUSDECL="X"' //PT1
    else Where := 'PSS_INCLUSDECL="X"';
    SetControlText('XX_WHERE',Where);
    
    SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS')); //PT2
end ;

procedure TOF_PGENVOISESSION_MUL.OnArgument (S : String ) ;
var
    Date1,date2 : THEdit;
    DateDebut,DateFin : TDateTime;
begin
  Inherited ;
	Arg := ReadTokenPipe(S,';');
	If Arg<>'' then NumeroEnvoi := StrToInt(ReadTokenPipe(S,';'));
	{$IFNDEF EAGLCLIENT}
	Grille := THDBGrid (GetControl ('Fliste'));
	{$ELSE}
	Grille := THGrid (GetControl ('Fliste'));
	TFM := TFMUL(Ecran);
	{$ENDIF}
	Q_Mul := THQuery(Ecran.FindComponent('Q'));
	if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
	if (Arg<>'SUPP') and (arg<>'AJOUT') then TFMul(Ecran).Bouvrir.OnClick := PreparerEnvoi
	Else
	begin
		if Arg = 'AJOUT' then TFMul(Ecran).Bouvrir.OnClick := AjouterSession;
		if Arg = 'SUPP' then TFMul(Ecran).Bouvrir.OnClick := EnleverSession;
	end;
	Date1 := THEdit(GetControl('PSS_DATEDEBUT'));
	If Date1<>Nil then Date1.OnElipsisClick := DateElipsisClick;
	Date2 := THEdit(GetControl('PSS_DATEDEBUT_'));
	If Date2<>Nil then Date2.OnElipsisClick := DateElipsisClick;
	MillesimeConsult := '0000';
	RendMillesimeRealise(DateDebut,DateFin);
	SetControlText('PSS_DATEDEBUT',DatetoStr(DateDebut));
	SetControlText('PSS_DATEDEBUT_',DateToStr(DateFin));
	SetControlText('VSESSIONS','NE');
	If Arg = 'AJOUT' Then Ecran.Caption := 'Ajout de sessions de formation pour l''envoi n° '+IntToStr(NumeroEnvoi);
	If Arg = 'SUPP' Then
	begin
		SetControlVisible('VSESSIONS',False);
		SetControlVisible('TVSESSIONS',False);
		Ecran.Caption := 'Suppression de sessions de formation pour l''envoi n° '+IntToStr(NumeroEnvoi);
	end;
	UpdateCaption(Ecran);
	
	//PT2 - Début
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		Begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PSS_PREDEFINI',False);
			SetControlText   ('PSS_PREDEFINI','');
		End
    	Else If V_PGI.ModePCL='1' Then 
    	Begin
    		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous);
    	End;
	end
	else
	begin
		SetControlVisible('PSS_PREDEFINI', False);
		SetControlVisible('TPSS_PREDEFINI',False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
	//PT2 - Fin
end ;

procedure TOF_PGENVOISESSION_MUL.DateElipsisclick(Sender : TObject);
var key : char;
begin
	key  :=  '*';
	ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGENVOISESSION_MUL.GrilleDblClick(Sender : TObject);
var St : String;
 Q_Mul : THQuery ;
begin
	Q_Mul := THQuery(Ecran.FindComponent('Q'));
	st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+
	Q_MUL.FindField('PSS_MILLESIME').AsString;
	AglLanceFiche ('PAY','SESSIONSTAGE', '', St , 'SAISIE'+';;;'+MillesimeConsult);
end;


procedure TOF_PGENVOISESSION_MUL.PreparerEnvoi(Sender : TObject);
begin
	AGLLanceFiche('PAY','PREP_ENVOIFORM','','',';0;'+GetControlText('PSS_DATEDEBUT') + ';' + GetControlText('PSS_DATEDEBUT_'));
	TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGENVOISESSION_MUL.AjouterSession(Sender : TObject);
begin
	AGLLanceFiche('PAY','PREP_ENVOIFORM','','','AJOUT;'+IntToStr(NumeroEnvoi)+';'+GetControlText('PSS_DATEDEBUT') + ';' + GetControlText('PSS_DATEDEBUT_'));
	TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGENVOISESSION_MUL.EnleverSession(Sender : TObject);
begin
	AGLLanceFiche('PAY','PREP_ENVOIFORM','','','SUPP;'+IntToStr(NumeroEnvoi)+';'+GetControlText('PSS_DATEDEBUT') + ';' + GetControlText('PSS_DATEDEBUT_'));
	TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

Initialization
  registerclasses ( [ TOF_PGENVOISESSION_MUL ] ) ;
end.


