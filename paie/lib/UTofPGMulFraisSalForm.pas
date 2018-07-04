{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 02/09/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULFRAISSALFORM ()
Mots clefs ... : TOF;PGMULFRAISSALFORM
*****************************************************************
PT1  | 03/11/2003 | V_50  | JL | Sasie des frais par salarié au lieu de la saisie par session pour CEGID
PT2  | 01/12/2003 | V_50  | JL | Correction pour affichage de la fiche
PT3  | 26/01/2007 | V_80  | FC | Mise en place du filtage habilitation pour les lookuplist
     |            |       |    |  pour les critères code salarié uniquement
PT4  | 22/05/2007 | V_720 | FL | FQ 11532 Test de la validité de la population de formation avant de pouvoir saisir les frais
PT5  | 19/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT6  | 03/04/2008 | V_803 | FL | Prise en compte d'un paramètre supplémentaire au lancement de l'écran en partage formation
PT7  | 04/06/2008 | V_804 | FL | FQ 15458 Report V7 Ne pas voir les salariés confidentiels
}
Unit UTofPGMulFraisSalForm;

Interface

uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,eMul,
{$ENDIF}
      Grids,HCtrls,HEnt1,vierge,EntPaie,HMsgBox,Hqry,UTOF,UTOB,UTOM,
      AGLInit,ParamDat,LookUp,PGOutils,PgPopulOutils,PgoutilsFormation; //PT4

Type
  TOF_PGMULFRAISSALFORM = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    Q_Mul      : THQuery ;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    LeStage,LeMillesime,LaSession,MillesimeEC,TypeSaisie:String;
    procedure GrilleDblClick(Sender:TObject);
    procedure SalarieElipsisClick(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULFRAISSALFORM.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULFRAISSALFORM.OnArgument (S : String ) ;
var SQL:String;
    Q:TQuery;
    Edit : THEdit;
begin
  Inherited ;
	TypeSaisie:=Trim(ReadTokenPipe(S,';'));
	LeStage:=Trim(ReadTokenPipe(S,';'));
	LaSession:=Trim(ReadTokenPipe(S,';'));
	LeMillesime:=Trim(ReadTokenPipe(S,';'));
	MillesimeEC:=Trim(ReadTokenPipe(S,';'));
	Q_Mul:=THQuery(Ecran.FindComponent('Q'));
	If PGBundleHierarchie Then SetControlText('WHEREPREDEF', ReadTokenPipe(S,';')); //PT6
	
	if TypeSaisie <> 'SAISIESALARIE' then   //DEBUT PT1
	begin
		Q:=OpenSQL('SELECT PSS_DATEDEBUT,PSS_DATEFIN FROM SESSIONSTAGE WHERE PSS_MILLESIME="'+LeMillesime+'"'+
		' AND PSS_ORDRE='+LaSession+' AND PSS_CODESTAGE="'+LeStage+'"',True);  //DB2
		if not Q.eof then
		begin
			SetControlText('DATEDEBUT',Q.FindField('PSS_DATEDEBUT').AsString);
			SetControlText('DATEFIN',Q.FindField('PSS_DATEFIN').AsString);
		end;
		Ferme(Q);
		SetControlProperty('ORDRE','Plus','PSS_CODESTAGE="'+LeStage+'" AND PSS_MILLESIME="'+LeMillesime+'" '+GetControlText('WHEREPREDEF')); //PT6
		SetControlProperty('CODESTAGE','Text',LeStage); //PT6
		SetControlProperty('ORDRE','Value',LaSession);
		SetControlText('MILLESIME',LeMillesime);
		SetControltext('MILLESIMEEC',MillesimeEC);
		If Not PGBundleHierarchie then
		begin
			SQL:='PSA_SALARIE IN (SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+LeStage+'" AND PAN_SALARIE<>"" AND PAN_MILLESIME="'+LeMillesime+'" AND PAN_ORDRE='+LaSession+') '+  //DB2
			'OR PSA_SALARIE IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+' AND PFO_MILLESIME="'+LeMillesime+'")'+ //DB2
		    //PT7
		    // N'affiche pas les salariés confidentiels
		    ' AND PSA_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"';
		end
		else
		begin
			//If Not PGDroitMultiForm then SetControlText('PSI_NODOSSIER',V_PGI.NoDossier);
			SQL:='PSI_INTERIMAIRE IN (SELECT PAN_SALARIE FROM SESSIONANIMAT WHERE PAN_CODESTAGE="'+LeStage+'" AND PAN_SALARIE<>"" AND PAN_MILLESIME="'+LeMillesime+'" AND PAN_ORDRE='+LaSession+') '+  //DB2
			'OR PSI_INTERIMAIRE IN (SELECT PFO_SALARIE FROM FORMATIONS WHERE PFO_CODESTAGE="'+LeStage+'" AND PFO_ORDRE='+LaSession+' AND PFO_MILLESIME="'+LeMillesime+'")'+ //DB2
		    //PT7
		    // N'affiche pas les salariés confidentiels
		    ' AND PSI_CONFIDENTIEL <= "'+RendNivAccesConfidentiel()+'"';
		end;
		SetControlText('XX_WHERE',SQL);
        //PT6 - Début
        if TypeSaisie = 'FRAIS' then
        Begin
            SetControlEnabled('CODESTAGE', False);
            SetControlEnabled('ORDRE',     False);
        End;
        //PT6 - Fin
	end
	else
	begin
		SetControlVisible('DATEDEBUT',False);
		SetControlVisible('DATEFIN',False);
		SetControlVisible('ORDRE',False);
		SetControlVisible('CODESTAGE',False);
		SetControlVisible('MILLESIME', False);
		SetControlVisible('MILLESIMEEC', False);
		SetControlVisible('TDATEDEBUT',False);
		SetControlVisible('TDATEFIN',False);
		SetControlVisible('TPFO_ORDRE',False);
		SetControlVisible('TPFO_CODESTAGE',False);
		SetControlVisible('TMILLESIME', False);
		Edit := THEdit(GetControl('PSA_SALARIE'));
		If Edit <> Nil then Edit.OnElipsisClick := SalarieElipsisClick;
	end;
	//FIN PT1
                
	{$IFNDEF EAGLCLIENT}
	Grille:=THDBGrid (GetControl ('Fliste'));
	{$ELSE}
	Grille:=THGrid (GetControl ('Fliste'));
	{$ENDIF}
	if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
end ;

procedure TOF_PGMULFRAISSALFORM.GrilleDblClick(Sender:TObject);
var St,Champ:String;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
        {$ENDIF}

        If PGBundleHierarchie then Champ := 'PSI_INTERIMAIRE'
        else Champ := 'PSA_SALARIE';

        If Q_MUL.FindField(Champ) = Nil Then Exit; //PT6

        If Q_MUL.FindField(Champ).AsString='' Then
        begin
                PGIBox('Vous devez choisir un salarié',TFMul(Ecran).Caption);
                Exit;
        end;
        //PT4 - Début
        { Si la gestion des plafonds s'effectue par population et que celle propre à la formation
          n'est pas valide, il faut interdire la saisie des inscriptions. }
        If VH_Paie.PGForGestPlafByPop And Not CanUsePopulation(TYP_POPUL_FORM_PREV) Then
              PGIBox('Attention, les populations de formation ne sont pas valides. #13#10'+
                     'La saisie des frais est impossible.',Ecran.Caption)
        Else
        Begin
        //PT4 - Fin
          If TypeSaisie <> 'SAISIESALARIE' then               //PT1
          begin
                St:=Q_MUL.FindField(Champ).AsString+';'+GetControlText('CODESTAGE')+';'+GetControlText('ORDRE')+';'+GetControlText('MILLESIME')+';'+GetControlText('MILLESIMEEC');
                AglLanceFiche('PAY','SAISIEFRAISSALFOR','','','FRAISSESSION;'+St);    //PT2
          end
          else
          begin
                St := Q_Mul.FindField(Champ).AsString;
                AglLanceFiche('PAY','MUL_SESSIONSTAGE', '', '' , 'FRAISSALARIE;;;;'+St );
          end;
        End;
end;

procedure TOF_PGMULFRAISSALFORM.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULFRAISSALFORM.SalarieElipsisClick (Sender : TObject);
var StWhere : String;//PT3
begin
  StWhere := RecupClauseHabilitationLookupList('');  //PT3
  
  If PGBundleHierarchie then //PT6
  	ElipsisSalarieMultiDos(Sender)
  else 
  	LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,'PSA_LIBELLE', True,-1);
end;


Initialization
  registerclasses ( [ TOF_PGMULFRAISSALFORM ] ) ;
end.
