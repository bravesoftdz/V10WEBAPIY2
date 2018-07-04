{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULSSOUSSESSION ()
Mots clefs ... : TOF;PGMULSSOUSSESSION
*****************************************************************}
Unit UTofPGMulSousSession ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,HDB,Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
     MainEAGL, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     hqry,
     UTOF ; 

Type
  TOF_PGMULSSOUSSESSION = Class (TOF)

    procedure OnArgument (S : String ) ; override ;
    private
    Q_Mul : THQuery;
    Stage,NumSessionMaitre,Millesime,Action : String;
    procedure GrilleDblClick(Sender : TObject);
    procedure CreatSousSession(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULSSOUSSESSION.OnArgument (S : String ) ;
var
  {$IFNDEF EAGLCLIENT}
  Grille : THDBGrid;
  {$ELSE}
  Grille : THGrid;
  {$ENDIF}
  Bt : TToolBarButton97;
begin
	Inherited ;

	Stage            := ReadTokenPipe(S,';');
	Millesime        := ReadTokenPipe(S,';');
	NumSessionMaitre := ReadTokenPipe(S,';');
	Action           := ReadTokenPipe(S,';');

	SetControltext('PSS_CODESTAGE', Stage);
	SetControltext('PSS_MILLESIME', Millesime);
	SetControltext('PSS_NUMSESSION',NumSessionMaitre);
	
	Q_Mul := THQuery(Ecran.FindComponent('Q'));
	{$IFNDEF EAGLCLIENT}
	Grille := THDBGrid (GetControl ('Fliste'));
	{$ELSE}
	Grille := THGrid (GetControl ('Fliste'));
	{$ENDIF}
	if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;

	Bt  :=  TToolbarButton97 (GetControl('BInsert'));
	If Bt <> Nil then Bt.OnClick := CreatSousSession;
	
	If Action = 'CONSULTATION' Then
	Begin
		SetControlEnabled('BInsert', False);
	End;
end ;

procedure TOF_PGMULSSOUSSESSION.GrilleDblClick(Sender : TObject);
var st : String;
    Bt : TToolBarButton97;
begin
    If Q_MUL.FindField('PSS_ORDRE').AsInteger = 0 Then Exit;
    
  	st  := Stage+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+Millesime;
  	AglLanceFiche ('PAY','SOUSSESSIONSTAGE', '' , St, 'SOUSSESSION;'+Stage+';'+Millesime+';;ACTION='+Action+';'+NumSessionMaitre);
  	
  	Bt  :=  TToolbarButton97 (GetControl('BCherche'));
  	if Bt  <>  NIL then Bt.click;
end;

procedure TOF_PGMULSSOUSSESSION.CreatSousSession(Sender : TObject);
var St : String;
    Bt : TToolBarButton97;
begin
	st  := Stage+';;'+Millesime;
	AglLanceFiche ('PAY','SOUSSESSIONSTAGE', '', '' ,'SOUSSESSION;'+Stage+';'+Millesime+';;ACTION=CREATION;'+NumSessionMaitre);
	
	Bt  :=  TToolbarButton97 (GetControl('BCherche'));
	if Bt  <>  NIL then Bt.click;
end;


Initialization
  registerclasses ( [ TOF_PGMULSSOUSSESSION ] ) ; 
end.

