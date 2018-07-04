{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 12/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULSCORING ()
Mots clefs ... : TOF;PGMULSCORING
*****************************************************************
PT1  | 24/04/2003 | V_42  | JL | Développement pour CWAS
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    |   pour les critères code salarié uniquement
PT3  | 22/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT4  | 02/04/2008 | V_803 | FL | Adaptation partage formation
PT5  | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue + gestion elipsis salarié et responsable
}
Unit UTofPG_MulScoring;

Interface

Uses StdCtrls,Controls,Classes,pgoutils,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
     MaineAGL,Emul,
{$ENDIF}
     forms,sysutils,UTOB,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,ParamDat,LookUp,P5Def,PgoutilsFormation,EntPaie,Hqry,PGOutils2 ;

Type
  TOF_PGMULSCORING = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    THVMillesime:THValComboBox;
    procedure ChangeMillesime(Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure RespElipsisClick(Sender : TObject);
    procedure ExitStage(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
    procedure SalarieElipsisClick(Sender : TObject);
    procedure ExitEdit(Sender : TObject);
end ;

Implementation

procedure TOF_PGMULSCORING.OnLoad ;
var Session : THValComboBox;
begin
  Inherited ;
        Session := THValComboBox(GetControl('SESSION'));
        If Session <> Nil Then
        begin
                If Session.Value <> '' Then SetControlText('PFO_ORDRE',Session.Value)
                Else SetControlText('PFO_ORDRE','');
        end;
        SetControlText('XX_WHEREPREDEF', DossiersAInterroger('',V_PGI.NoDossier,'PFO',True,False));
end ;

procedure TOF_PGMULSCORING.OnArgument (S : String ) ;
var THDate : THEdit;
    Resp,Stage,Edit : THEdit;
    Num : Integer;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    Arg : String;
    DD,DF : TDateTime;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        //DEBUT PT1
        If Arg = 'CWASCONSULTATION' then
        begin
                SetControlText('PFO_RESPONSFOR',V_PGI.UserSalarie);
                SetControlVisible('PFO_RESPONSFOR',False);
                SetControlVisible('TPFO_RESPONSFOR',False);
        end;
        //FIN PT1
        THVMillesime := THValComboBox(GetControl('THVMILLESIME'));
        THVMillesime.Value := RendMillesimeRealise(DD,DF);
        SetControlText('PFO_DATEDEBUT',DateToStr(DD));
        SetControlText('PFO_DATEFIN',DateToStr(DF));
        THVMillesime.OnChange := ChangeMillesime;
        THDate := THEdit(GetControl('PFO_DATEDEBUT'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('PFO_DATEFIN'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        SetControlCaption('LIBSTAGE','');
        SetControlCaption('LIBRESP','');
        SetControlCaption('LIBSALARIE','');
		If PGBundleHierarchie Then //PT5?
        Begin
	        Resp := THedit(Getcontrol('PFO_RESPONSFOR'));
	        If Resp <> Nil Then Resp.OnElipsisClick := RespElipsisClick;
		End;
	    
        Stage := THEdit(GetControl('PFO_CODESTAGE'));
        If Stage <> Nil Then Stage.OnExit := ExitStage;
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PFO_TRAVAILN'+IntToStr(Num)),GetControl ('TPFO_TRAVAILN'+IntToStr(Num)));
        end;
        VisibiliteStat (GetControl ('PFO_CODESTAT'),GetControl ('TPFO_CODESTAT')) ;
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num >8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PFO_FORMATION'+IntToStr(Num)),GetControl ('TPFO_FORMATION'+IntToStr(Num)));
        end;
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        {$ENDIF}
        if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;
        SetControlVisible('BINSERT',False);
        Edit := THEdit(GetControl('PFO_SALARIE'));
//        {$IFDEF EMANAGER}
		{$IFNDEF EMANAGER}If PGBundleHierarchie Then{$ENDIF} //PT5
        	If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
//        {$ENDIF}
        If Edit <> Nil then Edit.OnExit := ExitEdit;
        
        //PT4
        If PGBundleCatalogue Then  //PT5
        Begin
        	SetControlProperty ('PFO_CODESTAGE', 'Plus', ' AND (PST_PREDEFINI="STD" OR (PST_PREDEFINI="DOS" AND PST_NODOSSIER="'+V_PGI.NoDossier+'"))');
        End;
end ;

procedure TOF_PGMULSCORING.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+THVMillesime.Value+'"',True);
        If not Q.eof Then
        begin
                SetControlText('PFO_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PFO_DATEDEBUT_',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
                SetControlText('PFO_DATEFIN',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PFO_DATEFIN_',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
        end;
        Ferme(Q);
end;

procedure TOF_PGMULSCORING.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULSCORING.RespElipsisClick(Sender : TObject);
//var StWhere : String;
begin
		//PT4
        //StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES)';
        //LookupList(THEdit(Sender),'Liste des responsables de formation','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,'', True,-1);
        ElipsisResponsableMultidos (Sender);
end;

procedure TOF_PGMULSCORING.ExitStage(Sender : TObject);
Var Plus : String;
begin
        If GetControlText('PFO_CODESTAGE') <> '' Then
        begin
        	Plus :=  'PSS_CODESTAGE="'+GetControlText('PFO_CODESTAGE')+'" AND PSS_MILLESIME="'+THVMillesime.Value+'"'; //PT4
        	If PGBundleInscFormation Then
        	Begin
        		If Not PGDroitMultiForm Then 
        			Plus := Plus + DossiersAInterroger('',V_PGI.NoDossier,'PSS',True,True) //PT4
        		Else
        			Plus := Plus + DossiersAInterroger('','','PSS',True,True); //PT4
        	End;
            SetControlProperty('SESSION','Plus', Plus);
            SetControlEnabled('SESSION',True);
        end
        Else SetControlEnabled('SESSION',False);
end;

procedure TOF_PGMULSCORING.GrilleDblClick(Sender : TObject);
var St : String;
    Q_Mul : THQuery ;
    {$IFNDEF EAGLCLIENT}
//    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
begin
        {$IFNDEF EAGLCLIENT}
//        Liste := THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Liste := THGrid(GetControl('FLISTE'));
        {$ENDIF}
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If Q_MUL.FindField('PFO_SALARIE').AsString = '' Then
        begin
                PGIBox('Vous devez choisir un salarié',Ecran.Caption);
                Exit;
        end;
        st  := Q_MUL.FindField('PFO_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PFO_ORDRE').AsInteger)+';'+
        Q_Mul.FindField('PFO_SALARIE').AsString;
        AglLanceFiche ('PAY', 'SCORING', '','',St);
end;

procedure TOF_PGMULSCORING.SalarieElipsisClick(Sender : TObject);
//var St,StWhere,StOrder : String;
begin
     //PT4 - Début
    //If PGBundleInscFormation Then //PT5
    	ElipsisSalarieMultidos (Sender)
    //Else
    	//Inherited;
(*        St := ' SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM FROM SALARIES'+
        ' LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE ';
        StWhere := 'WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
        StOrder := ' PSA_SALARIE';
        StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT2
        LookupList(THEdit(Sender),'Liste des stages','SALARIES','PSA_SALARIE','PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1,St);*)
	//PT4 - Fin
end;

procedure TOF_PGMULSCORING.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;



Initialization
  registerclasses ( [ TOF_PGMULSCORING ] ) ;
end.


