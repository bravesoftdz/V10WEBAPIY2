{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 09/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULSESSIONSTAGE ()
Mots clefs ... : TOF;PGMULSESSIONSTAGE
*****************************************************************
PT1  | 03/11/2003 | V_50  | JL | Sasie des frais par salarié au lieu de la saisie par session pour CEGID
---  | 20/03/2006 |       | JL | modification clé annuaire ----
---  | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2  | 17/05/2007 | V_720 | FL | FQ 11532 Test de la validité de la population de formation avant de pouvoir saisir les frais
PT3  | 14/06/2007 | V_720 | JL | Gestion partage formation
PT4  | 02/01/2008 | V_802 | FL | FQ 14796 - Mod. des requêtes lancées suivant le critère de l'état de session
PT5  | 18/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT6  | 03/04/2008 | V_803 | FL | Adaptation partage formation
PT7  | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT8  | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue
PT9  | 28/05/2008 | V_804 | FL | FQ 14052 Appel de la fiche SESSIONSIMPLE
PT10 | 17/06/2008 | V_804 | FL | FQ 14055 Gestion des sous-sessions
}
Unit UTOFPGMul_SessionStage ;

Interface

uses  StdCtrls,Controls,Classes,Graphics,forms,sysutils,ComCtrls,HTB97,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Mul,Fe_Main,DBGrids,
{$ELSE}
       MaineAgl,eMul,
{$ENDIF}
      Grids,HCtrls,HEnt1,vierge,EntPaie,HMsgBox,Hqry,UTOF,UTOB,UTOM,
      AGLInit,ParamDat,PgoutilsFormation, PGPopulOutils;

Type
  TOF_PGMULSESSIONSTAGE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Bt         : TToolbarButton97;
    TypeSaisie,MillesimeConsult,SalarieFrais,LaNature,LeSalarie : String;
    THVMillesime : THValComboBox;
    MillesimeEnConsultation : Boolean;
    procedure CreatSession(Sender : TObject);
    procedure GrilleDblClick(Sender : TObject);
    procedure ChangeMillesime(Sender : TObject);
    procedure DateElipsisclick(Sender : TObject);
    procedure PasserelleAGEFOS(Sender : TObject);
    procedure LanceCalcSalaire(Sender : TObject);
    procedure ChangementEtat(Sender : TObject);
    procedure TraitementCarriere(Sender : TObject);
  end ;
var TobPrepAGEFOS:Tob;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    TFM : TFMul;
    {$ENDIF}
    Q_Mul : THQuery;


Implementation

Uses galOutil;


procedure TOF_PGMULSESSIONSTAGE.CreatSession(Sender : TObject);
var St : String;
    Stage : String;
begin
        Stage := '';
        If TypeSaisie = 'CARRIERE' then
        begin
                AglLanceFiche ('PAY','SESSIONSTAGE', '', '' ,TypeSaisie+';'+LaNature+';ACTION=CREATION');
        end
        else If typeSaisie = 'SIMPLIFIE' then
        begin
                //AglLanceFiche ('PAY','SAISIEFORMSIMP', '', '' ,'CREATION'); //PT9
                st  := Stage+';'+THVMillesime.Value+';'+THVMillesime.Value+';ACTION=CREATION';
                AglLanceFiche ('PAY','SESSIONSIMPLE', '', '' ,'MUL;'+ st);
        end
        else
        begin
                st  := Stage+';'+THVMillesime.Value+';'+THVMillesime.Value+';ACTION=CREATION';
                AglLanceFiche ('PAY','SESSIONSTAGE', '', '' ,'MUL;'+ st);
        end;
        if Bt = NIL then Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;

procedure TOF_PGMULSESSIONSTAGE.GrilleDblClick(Sender : TObject);
var st          : String;
    Q           : TQuery;
    TypeSession : String;
begin
        {$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
        {$ENDIF}
        If Q_MUL.FindField('PSS_CODESTAGE').AsString = '' Then
        begin
                If TypeSaisie = 'SAISIE' Then CreatSession(TFMul(Ecran).BInsert)
                Else PGIBox('Vous devez choisir une session',TFMul(Ecran).Caption);
                Exit;
        end;
        if TypeSaisie = 'CARRIERE' then
        begin
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche ('PAY','SESSIONSTAGE', '', St , TypeSaisie+';;;'+MillesimeConsult+';ACTION=CONSULTATION')
                Else AglLanceFiche ('PAY','SESSIONSTAGE', '', St , TypeSaisie+';'+LaNature);
                Bt  :=  TToolbarButton97 (GetControl('BCherche'));
                if Bt  <>  NIL then Bt.click;
        end
        else If TypeSaisie = 'SIMPLIFIE' Then
        begin
        	//PT10
        	If Q_MUL.FindField('PSS_PGTYPESESSION') <> Nil Then
        		TypeSession := Q_MUL.FindField('PSS_PGTYPESESSION').AsString
        	Else
        	Begin
        		Q := OpenSQL('SELECT PSS_PGTYPESESSION FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Q_MUL.FindField('PSS_CODESTAGE').AsString+'" AND PSS_ORDRE="'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+'" AND PSS_MILLESIME="'+Q_MUL.FindField('PSS_MILLESIME').AsString+'"', True);
        		If Not Q.EOF Then TypeSession := Q.FindField('PSS_PGTYPESESSION').AsString;
        		Ferme(Q);
        	End;
        	
        	If TypeSession <> 'SOS' Then
        	Begin
                (*st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                AglLanceFiche('PAY','SAISIEFORMSIMP','','','MODIFICATION;'+St);*) //PT9
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche ('PAY','SESSIONSIMPLE', '', St , 'SAISIE;;;'+MillesimeConsult+';ACTION=CONSULTATION')
                Else AglLanceFiche ('PAY','SESSIONSIMPLE', '', St , 'SAISIE;;;'+MillesimeConsult);
            End
            Else
            Begin //PT10
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche ('PAY','SOUSSESSIONSTAGE', '', St , 'SOUSSESSION;;;'+MillesimeConsult+';ACTION=CONSULTATION')
                Else AglLanceFiche ('PAY','SOUSSESSIONSTAGE', '', St , 'SOUSSESSION;;;'+MillesimeConsult);
            End;
            Bt  :=  TToolbarButton97 (GetControl('BCherche'));
            if Bt  <>  NIL then Bt.click;
        end
        else If TypeSaisie = 'SAISIEFORMATION' Then
        begin
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then  AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','SAISIEINSC;'+St+';ACTION=CONSULTATION')
                Else AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','SAISIEINSC;'+St);
                Bt  :=  TToolbarButton97 (GetControl('BCherche'));
                if Bt  <>  NIL then Bt.click;
        end
        else If TypeSaisie = 'PRESENCE' Then
        begin
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche('PAY','PRESENCEFORM','','','SAISIEFP;'+St+';ACTION=CONSULTATION')
                Else AglLanceFiche('PAY','PRESENCEFORM','','','SAISIEFP;'+St+';ACTION=CREATION');
                Bt  :=  TToolbarButton97 (GetControl('BCherche'));
                if Bt  <>  NIL then Bt.click;
        end
        else If TypeSaisie = 'CWASSAISIEFORMATION' Then
        begin
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','CWASSAISIEFORMATION;'+St+';ACTION=CONSULTATION')
                else
                begin
                        Q := OpenSQL('SELECT PSS_CLOTUREINSC FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Q_MUL.FindField('PSS_CODESTAGE').AsString+'" '+
                        'AND PSS_ORDRE='+Q_MUL.FindField('PSS_ORDRE').AsString+' AND PSS_MILLESIME="'+Q_MUL.FindField('PSS_MILLESIME').AsString+'"',True); // DB2
                        If Not Q.Eof then
                        begin
                                If Q.FindField('PSS_CLOTUREINSC').AsString = 'X' then AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','CWASSAISIEFORMATION;'+St+';ACTION=CONSULTATION')
                                Else AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','CWASSAISIEFORMATION;'+St);
                        end
                        else AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','CWASSAISIEFORMATION;'+St);
                        Ferme(Q);
                end;
                Bt  :=  TToolbarButton97 (GetControl('BCherche'));
                if Bt  <>  NIL then Bt.click;

        end
        else if TypeSaisie = 'SAISIE' then
        begin
        	//PT10
        	If Q_MUL.FindField('PSS_PGTYPESESSION') <> Nil Then
        		TypeSession := Q_MUL.FindField('PSS_PGTYPESESSION').AsString
        	Else
        	Begin
        		Q := OpenSQL('SELECT PSS_PGTYPESESSION FROM SESSIONSTAGE WHERE PSS_CODESTAGE="'+Q_MUL.FindField('PSS_CODESTAGE').AsString+'" AND PSS_ORDRE="'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+'" AND PSS_MILLESIME="'+Q_MUL.FindField('PSS_MILLESIME').AsString+'"', True);
        		If Not Q.EOF Then TypeSession := Q.FindField('PSS_PGTYPESESSION').AsString;
        		Ferme(Q);
        	End;
        	
        	If TypeSession <> 'SOS' Then
        	Begin
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche ('PAY','SESSIONSTAGE', '', St , TypeSaisie+';;;'+MillesimeConsult+';ACTION=CONSULTATION')
                Else AglLanceFiche ('PAY','SESSIONSTAGE', '', St , TypeSaisie+';;;'+MillesimeConsult);
            End
            Else
            Begin //PT10
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                if MillesimeEnConsultation=True then AglLanceFiche ('PAY','SOUSSESSIONSTAGE', '', St , 'SOUSSESSION;;;'+MillesimeConsult+';ACTION=CONSULTATION')
                Else AglLanceFiche ('PAY','SOUSSESSIONSTAGE', '', St , 'SOUSSESSION;;;'+MillesimeConsult);
            End;
            Bt  :=  TToolbarButton97 (GetControl('BCherche'));
            if Bt  <>  NIL then Bt.click;
        end
        //PT10
        (*else if TypeSaisie = 'SOUSSESSION' then
        begin
                st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                Q_MUL.FindField('PSS_MILLESIME').AsString;
                AglLanceFiche ('PAY','MULSOUSSESSION', '', '',St);
                Bt  :=  TToolbarButton97 (GetControl('BCherche'));
                if Bt  <>  NIL then Bt.click;
        end*)
        else If (TypeSaisie = 'FRAIS') or (TypeSaisie = 'CONSULTFRAIS') then
        begin
                //PT2 - Début
                { Si la gestion des plafonds s'effectue par population et que celle propre à la formation
                  n'est pas valide, il faut interdire la saisie des inscriptions. }
                If VH_Paie.PGForGestPlafByPop And Not CanUsePopulation(TYP_POPUL_FORM_PREV) Then
                      PGIBox('Attention, les populations de formation ne sont pas valides. #13#10'+
                             'La saisie des frais est impossible.',Ecran.Caption)
                Else
                Begin
                //PT2 - Fin
                    st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                    Q_MUL.FindField('PSS_MILLESIME').AsString+';'+THVMillesime.Value;
                    If TypeSaisie = 'FRAIS' Then
                    begin
                      If PGBundleInscFormation Or PGBundleHierarchie then
                      Begin
					    //PT6
					    If (GetControlText('PSS_PREDEFINI') <> '') Or (GetControlText('NODOSSIER')<>'') Then
							st := st + ';'+DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS',True,True);
                      	AglLanceFiche ('PAY','MULFRAISINTFORM', '', '' , TypeSaisie+';'+St);
                      end
                      else AglLanceFiche ('PAY','MULFRAISSALFORM', '', '' , TypeSaisie+';'+St);
                    end;
                    If TypeSaisie = 'CONSULTFRAIS' Then
                    Begin
						st := st + ';'+GetControlText('NODOSSIER'); //PT6
                        AglLanceFiche ('PAY','FRAISSALFORM', '', '' , TypeSaisie+';'+St);
                    End;
                End;
        end
        else If TypeSaisie = 'FRAISSALARIE' then                              //PT1
        begin
                //PT2 - Début
                { Si la gestion des plafonds s'effectue par population et que celle propre à la formation
                  n'est pas valide, il faut interdire la saisie des inscriptions. }
                If VH_Paie.PGForGestPlafByPop And Not CanUsePopulation(TYP_POPUL_FORM_PREV) Then
                      PGIBox('Attention, les populations de formation ne sont pas valides. #13#10'+
                             'La saisie des frais est impossible.',Ecran.Caption)
                Else
                Begin
                //PT2 - Fin
                    st  := SalarieFrais + ';' + Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+intToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+  //DB2
                    Q_MUL.FindField('PSS_MILLESIME').AsString+ ';' + MillesimeConsult;
                    AglLanceFiche ('PAY','SAISIEFRAISSALFOR', '', '','FRAISSALARIE;'+St);
                End;
        end
        else
        If TypeSaisie = 'RECUPPREVREA' then
        begin
             st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+
             Q_MUL.FindField('PSS_MILLESIME').AsString;
             if MillesimeEnConsultation=True then AglLanceFiche('PAY','PRESENCEFORM','','','RECUPPREVREA;'+St+';ACTION=CONSULTATION')
             Else AglLanceFiche('PAY','PRESENCEFORM','','','RECUPPREVREA;'+St+';ACTION=CREATION');
             Bt  :=  TToolbarButton97 (GetControl('BCherche'));
             if Bt  <>  NIL then Bt.click;
        end
        else
        If TypeSaisie = 'RECUPPREVPLA' then
        begin
             st  := Q_MUL.FindField('PSS_CODESTAGE').AsString+';'+IntToStr(Q_MUL.FindField('PSS_ORDRE').AsInteger)+';'+
             Q_MUL.FindField('PSS_MILLESIME').AsString;
             if MillesimeEnConsultation=True then  AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','RECUPPREVPLA;'+St+';ACTION=CONSULTATION')
             Else AglLanceFiche('PAY','SAISIESTAGEFORMAT','','','RECUPPREVPLA;'+St);
             Bt  :=  TToolbarButton97 (GetControl('BCherche'));
             if Bt  <>  NIL then Bt.click;
        end;
end;

procedure TOF_PGMULSESSIONSTAGE.OnLoad ;
var EtatSession : THValComboBox;
    StWhere : String;
begin
  Inherited ;
        EtatSession := THValComboBox(GetControl('ETATSESSION'));
        StWhere := '';
        IF EtatSession.Value <> '' Then
        begin
                If EtatSession.Value = 'ATT' then StWhere := 'PSS_DATEDEBUT="'+UsdateTime(IDate1900)+'"';
                If EtatSession.Value = 'INC' Then StWhere := 'PSS_CLOTUREINSC="-" AND PSS_EFFECTUE="-" AND PSS_VALIDORG="-"';
                If EtatSession.Value = 'CI' Then StWhere := 'PSS_CLOTUREINSC="X"';// AND PSS_EFFECTUE="-" AND PSS_VALIDORG="-"'; //PT4
                If EtatSession.Value = 'SE' Then StWhere := 'PSS_EFFECTUE="X"'; // AND PSS_VALIDORG="-"'; //PT4
                If EtatSession.Value = 'VO' Then StWhere := 'PSS_VALIDORG="X"';
        end;
        If VH_Paie.PGForGestionOPCA = True then
        begin
                if GetControlText('ETATENVOI') = 'ENV' then
                begin
                If StWhere <> '' then StWhere := StWhere+' AND PSS_NUMENVOI>0'  //DB2
                else StWhere := 'PSS_NUMENVOI>0';  //DB2
                        end;
                if GetControlText('ETATENVOI') = 'NON' then
                begin
                If StWhere <> '' then StWhere := StWhere+' AND PSS_NUMENVOI<=0'  //DB2
                else StWhere := 'PSS_NUMENVOI<=0';  //DB2
                end;
        end;
        If TypeSaisie = 'FRAISSALARIE' then  //PT1
        begin
                If StWHere <> '' then StWhere := StWhere + ' AND (PSS_CODESTAGE IN (SELECT PFO_CODESTAGE FROM FORMATIONS WHERE PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_SALARIE="'+SalarieFrais+'") '+
                                      'OR PSS_CODESTAGE IN (SELECT PAN_CODESTAGE FROM SESSIONANIMAT WHERE PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_SALARIE="'+SalarieFrais+'"))'
                else StWhere := 'PSS_CODESTAGE IN (SELECT PFO_CODESTAGE FROM FORMATIONS WHERE PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_SALARIE="'+SalarieFrais+'") '+
                                      'OR PSS_CODESTAGE IN (SELECT PAN_CODESTAGE FROM SESSIONANIMAT WHERE PAN_CODESTAGE=PSS_CODESTAGE AND PAN_ORDRE=PSS_ORDRE AND PAN_SALARIE="'+SalarieFrais+'")';
        end;
        // PT10
        //If StWhere <> '' then StWHere := StWhere+' AND PSS_PGTYPESESSION<>"SOS"'
        //else StWhere := 'PSS_PGTYPESESSION<>"SOS"';
        SetControlText('XX_WHERE',StWhere);       
        
	SetControlText('XX_WHEREPREDEF', DossiersAInterroger(GetControlText('PSS_PREDEFINI'),GetControlText('NODOSSIER'),'PSS')); //PT5
end ;

procedure TOF_PGMULSESSIONSTAGE.OnArgument (S : String ) ;
var Stage,Millesime : String;
    Bouv,BTraitement : TToolBarButton97;
    EtatSession : THValComboBox;
    Q : TQuery;
    THDate : THEdit;
    Num : Integer;
    DD,DF : TDateTime;
begin
  Inherited ;
        MillesimeConsult := '';
        TypeSaisie := ReadTokenPipe(S,';');
        If TypeSaisie = 'CARRIERE' then
        begin
                LaNature := Trim(ReadTokenPipe(S,';')) ;
                LeSalarie := Trim(ReadTokenPipe(S,';')) ;
        end
        else
        If TypeSaisie = 'RECUPPREV' then
        begin
             Stage :=  Trim(ReadTokenPipe(S,';')) ;
             Millesime := Trim(ReadTokenPipe(S,';'));
             MillesimeConsult := Trim(ReadTokenPipe(S,';'));
        end
        else
        begin
                Stage :=  Trim(ReadTokenPipe(S,';')) ;
                Millesime := Trim(ReadTokenPipe(S,';'));
                MillesimeConsult := Trim(ReadTokenPipe(S,';'));
                SalarieFrais := Trim(ReadTokenPipe(S,';'));
        end;
        BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
        if Bouv  <>  NIL then  BOuv.OnClick  :=  CreatSession;
        If Stage <> '' Then
        begin
                SetControlEnabled('PSS_CODESTAGE',False);
                SetControlText('PSS_CODESTAGE',Stage);
        end;
        MillesimeEnConsultation := True;
        THVMillesime := THValComboBox(GetControl('THVMILLESIME'));
        If MillesimeConsult = '' then
        begin
                Millesimeconsult := RendMillesimeRealise(DD,DF);
                SetControlText('PSS_DATEDEBUT',DateToStr(DD));
                SetControlText('PSS_DATEDEBUT_',DateToStr(DF));
                SetControlText('PSS_DATEFIN',DateToStr(DD));
                SetControlText('PSS_DATEFIN_',DateToStr(DF));
                MillesimeEnConsultation := False;
                THVMillesime.Value := MillesimeConsult;
        end
        Else
        begin
                Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_ACTIF,PFE_CLOTURE FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeConsult+'"',True);
                If Not Q.Eof Then
                begin
                        SetControlText('PSS_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                        SetControlText('PSS_DATEDEBUT_',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
                        SetControlText('PSS_DATEFIN',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDateTime));
                        SetControlText('PSS_DATEFIN_',DateToStr(Q.FindField('PFE_DATEFIN').AsDateTime));
                        If (Q.FindField('PFE_ACTIF').AsString = 'X') and (Q.FindField('PFE_CLOTURE').AsString = '-') then MillesimeEnConsultation := False;
                end;
                THVMillesime.Value := MillesimeConsult;
                THVMillesime.Enabled := False;
                Ferme(Q);
        end;
        if THVMillesime <> Nil then THVMillesime.OnChange := ChangeMillesime;
        If (TypeSaisie = 'FRAIS') or (TypeSaisie = 'CONSULTFRAIS') then
        begin
                Ecran.Caption := 'Saisie des frais : liste des sessions de formation';
                BOuv.Visible := False;
                SetControltext('XX_WHERE','PSS_EFFECTUE="X"');
        end
        Else
        begin
                If Not (ExisteSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_MILLESIME="'+MillesimeConsult+'" AND PFE_ACTIF="X" AND PFE_CLOTURE="-"')) Then
                BOuv.Visible := False
                Else BOuv.Visible := True
        end;
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        {$IFNDEF EAGLCLIENT}
        Grille := THDBGrid (GetControl ('Fliste'));
        {$ELSE}
        Grille := THGrid (GetControl ('Fliste'));
        TFM := TFMUL(Ecran);
        {$ENDIF}
        if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;
        THDate := THEdit(GetControl('PSS_DATEDEBUT'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('PSS_DATEFIN'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('PSS_DATEDEBUT_'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        THDate := THEdit(GetControl('PSS_DATEFIN_'));
        If thdATE <> NIL Then THDate.OnElipsisClick := DateElipsisClick;
        EtatSession := THValComboBox(GetControl('ETATSESSION'));
        If EtatSession <> Nil then EtatSession.OnChange := ChangementEtat;
        If (TypeSaisie = 'SAISIE') or (TypeSaisie = 'SIMPLIFIE') Then SetControlProperty('ETATSESSION','Value','')
        Else SetControlVisible('BINSERT',False);
        If TypeSaisie = 'SAISIEFORMATION' Then
        begin
                SetControlProperty('ETATSESSION','Value','INC');
                Ecran.Caption := 'Saisie des inscriptions aux sessions de formation';
        end;
        If TypeSaisie = 'CWASSAISIEFORMATION' Then
        begin
                SetControlProperty('ETATSESSION','Value','INC');
                Ecran.Caption := 'Saisie des inscriptions aux sessions de formation';
        end;
        If TypeSaisie = 'FRAISSALARIE' then
        begin
                SetControlProperty('ETATSESSION','Value','SE');
                Ecran.Caption := 'Saisie des frais de '+SalarieFrais+' '+RechDom ('PGSALARIE',SalarieFrais,FALSE);
        end;
        If TypeSaisie = 'PRESENCE' Then
        begin
                SetControlProperty('ETATSESSION','Value','SE');
                Ecran.Caption := 'Saisie des feuille de présence : liste des sessions de formation';
        end;
        // Saisie simple / Saisie en grille des frais
        If (TypeSaisie = 'FRAIS') Or (TypeSaisie = 'CONSULTFRAIS') Then  //PT6
        	SetControlProperty('ETATSESSION','Value','SE');  
        
        For Num  :=  1 to VH_Paie.NBFormationLibre do
        begin
                if Num > 8 then Break;
                VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
        end;
        
        If TypeSaisie = 'PASSERELLEAGEFOS' then
        begin
                TFMul(Ecran).BOuvrir.OnClick := PasserelleAGEFOS;
                SetControlProperty ('FListe','Multiselection', True);
                TFMul(Ecran).bSelectAll.Visible := True;
        end;
        If TypeSaisie = 'CALCSALAIRE' then
        begin
                Ecran.Caption := 'Mise à jour des coûts';
                SetControlVisible('BINSERT',False);
                SetControlProperty ('FListe','Multiselection', True);
                TFMul(Ecran).bSelectAll.Visible := True;
                TFMul(Ecran).BOuvrir.OnClick := LanceCalcSalaire;
        end;
        If VH_Paie.PGForGestionOPCA = True then
        begin
                SetControlVisible('ETATENVOI',True);
                SetControlVisible('TETATENVOI',True);
        end;
        If TypeSaisie = 'CARRIERE' then
        begin
                Ecran.Caption := 'Liste des sessions';
                SetControlProperty ('FListe','Multiselection', True);
                TFMul(Ecran).bSelectAll.Visible := True;
                SetControlVisible('BTRAITEMENT',True);
                If LaNature = 'INITIALE' then
                begin
                        SetControlText('PSS_NATUREFORM','004');
                        Bouv.Visible := True;
                        SetControlVisible('THVMILLESIME',False);                   
                        SetControlVisible('LTHVMILLESIME',False);
                end;
                If LaNature = 'EXTERNE' then SetControlText('PSS_NATUREFORM','002');
                If LaNature = 'INTERNE' then SetControlText('PSS_NATUREFORM','001');
                SetControlEnabled('PSS_NATUREFORM',False);
                BTraitement := TToolBarButton97(GetControl('BTRAITEMENT'));
                If BTraitement <> Nil then BTraitement.OnClick := TraitementCarriere;
        end;
        UpdateCaption(TFMul(Ecran));
        
        //DEBUT PT3
		If PGBundleCatalogue Then //PT8
		Begin
			If not PGDroitMultiForm then
				SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True))
			Else If V_PGI.ModePCL='1' Then 
        		SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT7
		End;
		
		If PGBundleInscFormation then
		begin
			If not PGDroitMultiForm then
			Begin
				SetControlEnabled('NODOSSIER',False);
				SetControlText   ('NODOSSIER',V_PGI.NoDossier);
				SetControlEnabled('PSS_PREDEFINI',False);
				SetControlText   ('PSS_PREDEFINI','');
				//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT7 //PT8
			End
        	Else If V_PGI.ModePCL='1' Then 
        	Begin
        		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT6
        		//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT7 //PT8
        	End;
		end
		else
		begin
			SetControlVisible('PSS_PREDEFINI', False);
			SetControlVisible('TPSS_PREDEFINI',False);
			SetControlVisible('NODOSSIER',     False);
			SetControlVisible('TNODOSSIER',    False);
		end;
        //FIN PT3

end ;

procedure TOF_PGMULSESSIONSTAGE.ChangeMillesime(Sender : TObject);
var Q : TQuery;
begin
        Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN,PFE_MILLESIME,PFE_ACTIF,PFE_CLOTURE FROM EXERFORMATION WHERE PFE_MILLESIME="'+THVMillesime.Value+'"',True);
        If Not Q.Eof Then
        begin
                SetControlText('PSS_DATEDEBUT',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PSS_DATEDEBUT_',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
                SetControlText('PSS_DATEFIN',DateToStr(Q.FindField('PFE_DATEDEBUT').AsDatetime));
                SetControlText('PSS_DATEFIN_',DateToStr(Q.FindField('PFE_DATEFIN').AsDatetime));
                If (Q.FindField('PFE_ACTIF').AsString <> 'X') and (Q.FindField('PFE_CLOTURE').AsString = '-') then
                begin
                        SetControlVisible('BINSERT',False);
                        MillesimeEnConsultation := True;
                end
                Else
                begin
                        SetControlVisible('BINSERT',True);
                        MillesimeEnConsultation := False;
                end;
        end;
        Ferme(Q);
end;

procedure TOF_PGMULSESSIONSTAGE.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGMULSESSIONSTAGE.PasserelleAGEFOS(Sender : TObject);
var    {$IFNDEF EAGLCLIENT}
       Grille : THDBGrid;
       {$ELSE}
       Grille : THGrid;
        {$ENDIF}
        Q_Mul : THQuery;
        i,Session : Integer;
        Stage,Millesime : String;
        T : Tob;
begin
        TobPrepAGEFOS := Tob.Create('Les sessions',Nil,-1);
        {$IFNDEF EAGLCLIENT}
        Grille  :=  THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Grille  :=  THGrid(GetControl('FLISTE'));
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If (Grille=nil) then Exit;
        if (Grille.NbSelected=0) and (not Grille.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;
        if ((Grille.nbSelected)>0) AND (not Grille.AllSelected ) then
        begin
                for i := 0 to Grille.NbSelected-1 do
                begin
                        Grille.GotoLeBookmark(i);
                        T := Tob.Create('Une fille session',TobPrepAGEFOS,-1);
                        T.AddChampSup('STAGE',False);
                        T.AddChampSup('SESSION',False);
                        T.AddChampSup('MILLESIME',False);
                        Stage := Q_Mul.FindField('PSS_CODESTAGE').AsString;
                        Session := Q_Mul.FindField('PSS_ORDRE').AsInteger;
                        Millesime := Q_Mul.FindField('PSS_MILLESIME').AsString;
                        T.PutValue('STAGE',Stage);
                        T.PutValue('SESSION',Session);
                        T.PutValue('MILLESIME',Millesime);
                end;
        Grille.ClearSelected;
        end;
        If (Grille.AllSelected=TRUE) then
        begin
                Q_Mul.First;
                while Not Q_Mul.EOF do
                begin
                        T := Tob.Create('Une fille session',TobPrepAGEFOS,-1);
                        T.AddChampSup('STAGE',False);
                        T.AddChampSup('SESSION',False);
                        T.AddChampSup('MILLESIME',False);
                        Stage := Q_Mul.FindField('PSS_CODESTAGE').AsString;
                        Session := Q_Mul.FindField('PSS_ORDRE').AsInteger;
                        Millesime := Q_Mul.FindField('PSS_MILLESIME').AsString;
                        T.PutValue('STAGE',Stage);
                        T.PutValue('SESSION',Session);
                        T.PutValue('MILLESIME',Millesime);
                        Q_Mul.Next;
                end;
                Grille.AllSelected := False;
        end;
        AGLLanceFiche('PAY','PASSERELLEAGEFOS','','','');
        TobPrepAGEFOS.Free;
end;

procedure TOF_PGMULSESSIONSTAGE.LanceCalcSalaire(Sender : TObject);
begin
{        if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;}
        AGLLanceFiche('PAY','CALCSALAIREFORM','','','REALISE');
end;

procedure TOF_PGMULSESSIONSTAGE.ChangementEtat(Sender : TOBject);
var DD,DF : TDateTime;
begin
        If Sender = Nil then Exit;
        If THValComboBox(Sender).Value = 'ATT' then
        begin
                SetControlEnabled('THVMILLESIME',False);
                SetControlEnabled('PSS_DATEDEBUT',False);
                SetControlEnabled('PSS_DATEDEBUT_',False);
                SetControlEnabled('TPSS_DATEDEBUT',False);
                SetControlEnabled('TPSS_DATEDEBUT_',False);
                SetControlEnabled('PSS_DATEFIN',False);
                SetControlEnabled('PSS_DATEFIN_',False);
                SetControlEnabled('TPSS_DATEFIN',False);
                SetControlEnabled('TPSS_DATEFIN_',False);
                SetControlText('PSS_DATEDEBUT',DateToStr(IDate1900));
                SetControlText('PSS_DATEDEBUT_',DateToStr(IDate1900));
                SetControlText('PSS_DATEFIN',DateToStr(IDate1900));
                SetControlText('PSS_DATEFIN_',DateToStr(IDate1900));
        end
        else
        begin
                SetControlEnabled('PSS_DATEDEBUT',True);
                SetControlEnabled('PSS_DATEDEBUT_',True);
                SetControlEnabled('TPSS_DATEDEBUT',True);
                SetControlEnabled('TPSS_DATEDEBUT_',True);
                SetControlEnabled('PSS_DATEFIN',True);
                SetControlEnabled('PSS_DATEFIN_',True);
                SetControlEnabled('TPSS_DATEFIN',True);
                SetControlEnabled('TPSS_DATEFIN_',True);
                If StrToDate(GetControlText('PSS_DATEDEBUT')) = IDate1900 then
                begin
                        RendMillesimeRealise(DD,DF);
                        SetControlText('PSS_DATEDEBUT',DateToStr(DD));
                        SetControlText('PSS_DATEDEBUT_',DateToStr(DF));
                        SetControlText('PSS_DATEFIN',DateToStr(DD));
                        SetControlText('PSS_DATEFIN_',DateToStr(DF));
                        SetControlEnabled('THVMILLESIME',True);
                end;
        end;
        TFMul(Ecran).BChercheClick(TFMul(Ecran).BCherche);
end;

procedure TOF_PGMULSESSIONSTAGE.TraitementCarriere(Sender : TObject);
var    {$IFNDEF EAGLCLIENT}
       Grille : THDBGrid;
       {$ELSE}
       Grille : THGrid;
        {$ENDIF}
        Q_Mul : THQuery;
        Q : TQuery;
        TobFormations,TobSessions,TF: Tob;
        i,j : Integer;
        Session,Stage,Millesime : String;
        Nom,Prenom,Etab,LibEmploi : String;
        TN1,TN2,TN3,TN4,LC1,LC2,LC3,LC4,CS : String;
        Req,Pfx : String;
begin
        {$IFNDEF EAGLCLIENT}
        Grille  :=  THDBGrid(GetControl('FLISTE'));
        {$ELSE}
        Grille  :=  THGrid(GetControl('FLISTE'));
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If (Grille=nil) then Exit;
        if (Grille.NbSelected=0) and (not Grille.AllSelected) then
        begin
                MessageAlerte('Aucun élément sélectionné');
                exit;
        end;
        
        //PT6 - Début
        Req := 'SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBELLE,PSA_ETABLISSEMENT,PSA_LIBELLEEMPLOI,PSA_PRENOM,PSE_RESPONSFOR'+
        ',PSA_LIBREPCMB1,PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4 FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE WHERE PSA_SALARIE="'+LeSalarie+'"';

		Req := AdaptMultiDosForm (Req, False);
        //PT6 - Fin
        
        Q := OpenSQL(Req,True);
        If Not Q.eof then
        begin
                Nom := Q.FindField(Pfx+'_LIBELLE').AsString;
                Prenom := Q.FindField(Pfx+'_PRENOM').AsString;
                Etab := Q.FindField(Pfx+'_ETABLISSEMENT').AsString;
                LibEmploi := Q.FindField(Pfx+'_LIBELLEEMPLOI').AsString;
                TN1 := Q.FindField(Pfx+'_TRAVAILN1').AsString;
                TN2 := Q.FindField(Pfx+'_TRAVAILN2').AsString;
                TN3 := Q.FindField(Pfx+'_TRAVAILN3').AsString;
                TN4 := Q.FindField(Pfx+'_TRAVAILN4').AsString;
                LC1 := Q.FindField(Pfx+'_LIBREPCMB1').AsString;
                LC2 := Q.FindField(Pfx+'_LIBREPCMB2').AsString;
                LC3 := Q.FindField(Pfx+'_LIBREPCMB3').AsString;
                LC4 := Q.FindField(Pfx+'_LIBREPCMB4').AsString;
                CS := Q.FindField(Pfx+'_CODESTAT').AsString;
        end;
        Ferme(Q);

        TobFormations := Tob.Create('Les formations',Nil,-1);
        if ((Grille.nbSelected)>0) AND (not Grille.AllSelected ) then
        begin
                for i := 0 to Grille.NbSelected-1 do
                begin
                        Grille.GotoLeBookmark(i);
                        {$IFDEF EAGLCLIENT}
                        TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ; 
                        {$ENDIF}
                        Stage := Q_mul.FindField('PSS_CODESTAGE').AsString;
                        Session := IntToStr(Q_mul.FindField('PSS_ORDRE').AsInteger);
                        Millesime := q_mul.FindField('PSS_MILLESIME').AsString;
                        Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_DATEDEBUT,PSS_DATEFIN,'+
                        'PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,'+
                        'PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8,PSS_DEBUTDJ,PSS_FINDJ,PSS_ORGCOLLECTSGU FROM SESSIONSTAGE '+
                        'WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE='+Session+' AND PSS_MILLESIME="'+Millesime+'"',True);
                        TobSessions := Tob.Create('Les sessions',Nil,-1);
                        TobSessions.LoadDetailDB('Les sessions','','',Q,false);
                        Ferme(Q);
                        For j := 0 to TobSessions.Detail.Count - 1 do
                        begin
                                TF := Tob.Create('FORMATIONS',TobFormations,-1);
                                TF.PutValue('PFO_SALARIE',LeSalarie);
                                TF.PutValue('PFO_CODESTAGE',TobSessions.Detail[j].GetValue('PSS_CODESTAGE'));
                                TF.PutValue('PFO_ORDRE',TobSessions.Detail[j].GetValue('PSS_ORDRE'));
                                TF.PutValue('PFO_MILLESIME',TobSessions.Detail[j].GetValue('PSS_MILLESIME'));
                                TF.PutValue('PFO_DATEDEBUT',TobSessions.Detail[j].GetValue('PSS_DATEDEBUT'));
                                TF.PutValue('PFO_DATEFIN',TobSessions.Detail[j].GetValue('PSS_DATEFIN'));
                                TF.PutValue('PFO_FORMATION1',TobSessions.Detail[j].GetValue('PSS_FORMATION1'));
                                TF.PutValue('PFO_FORMATION2',TobSessions.Detail[j].GetValue('PSS_FORMATION2'));
                                TF.PutValue('PFO_FORMATION3',TobSessions.Detail[j].GetValue('PSS_FORMATION3'));
                                TF.PutValue('PFO_FORMATION4',TobSessions.Detail[j].GetValue('PSS_FORMATION4'));
                                TF.PutValue('PFO_FORMATION5',TobSessions.Detail[j].GetValue('PSS_FORMATION5'));
                                TF.PutValue('PFO_FORMATION6',TobSessions.Detail[j].GetValue('PSS_FORMATION6'));
                                TF.PutValue('PFO_FORMATION7',TobSessions.Detail[j].GetValue('PSS_FORMATION7'));
                                TF.PutValue('PFO_FORMATION8',TobSessions.Detail[j].GetValue('PSS_FORMATION8'));
                                TF.PutValue('PFO_DEBUTDJ',TobSessions.Detail[j].GetValue('PSS_DEBUTDJ'));
                                TF.PutValue('PFO_FINDJ',TobSessions.Detail[j].GetValue('PSS_FINDJ'));
                                TF.PutValue('PFO_ORGCOLLECTGU',TobSessions.Detail[j].GetValue('PSS_ORGCOLLECTSGU'));
                                TF.PutValue('PFO_NOMSALARIE',Nom);
                                TF.PutValue('PFO_PRENOM',Prenom);
                                TF.PutValue('PFO_ETABLISSEMENT',Etab);
                                TF.PutValue('PFO_LIBEMPLOIFOR',LibEmploi);
                                TF.PutValue('PFO_TRAVAILN1',TN1);
                                TF.PutValue('PFO_TRAVAILN2',TN2);
                                TF.PutValue('PFO_TRAVAILN3',TN3);
                                TF.PutValue('PFO_TRAVAILN4',TN4);
                                TF.PutValue('PFO_LIBREPCMB1',LC1);
                                TF.PutValue('PFO_LIBREPCMB2',LC2);
                                TF.PutValue('PFO_LIBREPCMB3',LC3);
                                TF.PutValue('PFO_LIBREPCMB4',LC4);
                                TF.PutValue('PFO_CODESTAT',CS);
                                TF.PutValue('PFO_EFFECTUE','X');
                                
						        //PT6 - Début
						        If PGBundleInscFormation Then
						        Begin
						        	TF.PutValue('PFO_NODOSSIER', GetNoDossierSalarie(LeSalarie));
						        	TF.PutValue('PFO_PREDEFINI', 'DOS');
                                End
						        Else
						        Begin
						        	TF.PutValue('PFO_NODOSSIER', V_PGI.NoDossier);
						        	TF.PutValue('PFO_PREDEFINI', 'STD');
						        End;
						        //PT6 - Fin
        
                                TF.InsertOrUpdateDB(False);
                        end;
                        TobSessions.Free;
                end;
        Grille.ClearSelected;
        end;
        If (Grille.AllSelected=TRUE) then
        begin
                Q_Mul.First;
                while Not Q_Mul.EOF do
                begin
                        Stage := Q_mul.FindField('PSS_CODESTAGE').AsString;
                        Session := IntToStr(Q_mul.FindField('PSS_ORDRE').AsInteger);
                        Millesime := Q_mul.FindField('PSS_MILLESIME').AsString;
                        Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_DATEDEBUT,PSS_DATEFIN,PSS_NATUREFORM,'+
                        'PSS_FORMATION1,PSS_FORMATION2,PSS_FORMATION3,PSS_FORMATION4,PSS_FORMATION5,'+
                        'PSS_FORMATION6,PSS_FORMATION7,PSS_FORMATION8,PSS_DEBUTDJ,PSS_FINDJ,PSS_ORGCOLLECTSGU FROM SESSIONSTAGE '+
                        'WHERE PSS_CODESTAGE="'+Stage+'" AND PSS_ORDRE='+Session+' AND PSS_MILLESIME="'+Millesime+'"',True);
                        TobSessions := Tob.Create('Les sessions',Nil,-1);
                        TobSessions.LoadDetailDB('Les sessions','','',Q,false);
                        Ferme(Q);
                        For j := 0 to TobSessions.Detail.Count - 1 do
                        begin
                                TF := Tob.Create('FORMATIONS',TobFormations,-1);
                                TF.PutValue('PFO_SALARIE',LeSalarie);
                                TF.PutValue('PFO_CODESTAGE',TobSessions.Detail[j].GetValue('PSS_CODESTAGE'));
                                TF.PutValue('PFO_ORDRE',TobSessions.Detail[j].GetValue('PSS_ORDRE'));
                                TF.PutValue('PFO_MILLESIME',TobSessions.Detail[j].GetValue('PSS_MILLESIME'));
                                TF.PutValue('PFO_DATEDEBUT',TobSessions.Detail[j].GetValue('PSS_DATEDEBUT'));
                                TF.PutValue('PFO_DATEFIN',TobSessions.Detail[j].GetValue('PSS_DATEFIN'));
                                TF.PutValue('PFO_NATUREFORM',TobSessions.Detail[j].GetValue('PSS_NATUREFORM'));
                                TF.PutValue('PFO_FORMATION1',TobSessions.Detail[j].GetValue('PSS_FORMATION1'));
                                TF.PutValue('PFO_FORMATION2',TobSessions.Detail[j].GetValue('PSS_FORMATION2'));
                                TF.PutValue('PFO_FORMATION3',TobSessions.Detail[j].GetValue('PSS_FORMATION3'));
                                TF.PutValue('PFO_FORMATION4',TobSessions.Detail[j].GetValue('PSS_FORMATION4'));
                                TF.PutValue('PFO_FORMATION5',TobSessions.Detail[j].GetValue('PSS_FORMATION5'));
                                TF.PutValue('PFO_FORMATION6',TobSessions.Detail[j].GetValue('PSS_FORMATION6'));
                                TF.PutValue('PFO_FORMATION7',TobSessions.Detail[j].GetValue('PSS_FORMATION7'));
                                TF.PutValue('PFO_FORMATION8',TobSessions.Detail[j].GetValue('PSS_FORMATION8'));
                                TF.PutValue('PFO_DEBUTDJ',TobSessions.Detail[j].GetValue('PSS_DEBUTDJ'));
                                TF.PutValue('PFO_FINDJ',TobSessions.Detail[j].GetValue('PSS_FINDJ'));
                                TF.PutValue('PFO_ORGCOLLECTGU',TobSessions.Detail[j].GetValue('PSS_ORGCOLLECTSGU'));
                                TF.PutValue('PFO_NOMSALARIE',Nom);
                                TF.PutValue('PFO_PRENOM',Prenom);
                                TF.PutValue('PFO_ETABLISSEMENT',Etab);
                                TF.PutValue('PFO_LIBEMPLOIFOR',LibEmploi);
                                TF.PutValue('PFO_TRAVAILN1',TN1);
                                TF.PutValue('PFO_TRAVAILN2',TN2);
                                TF.PutValue('PFO_TRAVAILN3',TN3);
                                TF.PutValue('PFO_TRAVAILN4',TN4);
                                TF.PutValue('PFO_LIBREPCMB1',LC1);
                                TF.PutValue('PFO_LIBREPCMB2',LC2);
                                TF.PutValue('PFO_LIBREPCMB3',LC3);
                                TF.PutValue('PFO_LIBREPCMB4',LC4);
                                TF.PutValue('PFO_CODESTAT',CS);
                                TF.PutValue('PFO_EFFECTUE','X');
                                
						        //PT6 - Début
						        If PGBundleInscFormation Then
						        Begin
						        	TF.PutValue('PFO_NODOSSIER', GetNoDossierSalarie(LeSalarie));
						        	TF.PutValue('PFO_PREDEFINI', 'DOS');
                                End
						        Else
						        Begin
						        	TF.PutValue('PFO_NODOSSIER', V_PGI.NoDossier);
						        	TF.PutValue('PFO_PREDEFINI', 'STD');
						        End;
						        //PT6 - Fin
        
                                TF.InsertOrUpdateDB(False);
                        end;
                        TobSessions.Free;
                        Q_Mul.Next;
                end;
                Grille.AllSelected := False;
        end;
        TobFormations.Free;
end;

Initialization
  registerclasses ( [ TOF_PGMULSESSIONSTAGE ] ) ;
end.

