{***********UNITE*************************************************
Auteur  ...... :  JL
Créé le ...... : 03/12/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGSCORING ()
Mots clefs ... : TOF;PGSCORING
*****************************************************************
PT1  | 25/11/2005 | V_65  | JL | Corrections affichage tablette réponse
PT2  | 15/03/2006 | V_650 | JL | FQ 12973 Edition de la fiche
PT3  | 22/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT4  | 02/04/2008 | V_803 | FL | Adaptation partage formation
PT5  | 09/04/2008 | V_804 | FL | Ne pas afficher l'onglet suite si pas assez de questions de paramétrées
}
Unit UtofPGScoring ;

Interface

Uses Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}EdtREtat,
{$else}
     eMul,MainEAGL, UtileAgl,
{$ENDIF}
     HEnt1,sysutils,HCtrls,UTOF,UTob,extctrls,HTB97,ComCtrls;


Type
  TOF_PGSCORING = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    LeSalarie,LaSession,LeStage : String;
    TableauQuest : array  [1..40] of Integer;
    procedure MiseEnPageTheme;
    procedure AffichageChamps;
    procedure MajScoring(Sender : TObject);
    procedure EditionScoring(Sender : TObject);    //PT2
  end ;

Implementation
Uses PGOutilsFormation;

procedure TOF_PGSCORING.OnArgument (S : String ) ;
var BValid,Bt : TToolBarButton97;
begin
  Inherited;
        LeStage   := readTokenPipe(S,';');
        LaSession := readTokenPipe(S,';');
        LeSalarie := readTokenPipe(S,';');
        BValid := TToolBarButton97(GetControl('BVALIDER'));
        If BValid <> Nil then BValid.OnClick := MajScoring;
        MiseEnPageTheme;
        AffichageChamps;
        Bt := TToolBarButton97(GetControl('BImprimer'));
        If Bt <> Nil then Bt.OnClick := EditionScoring;
end ;

procedure TOF_PGSCORING.MiseEnPageTheme;
var i,NbThemes : Integer;
    Combo : THValComboBox;
    TLabelLib,TLabel : THLabel;
    Cadre : TBevel;
    Report : Boolean;
    Q : TQuery;
    TobQuestions : Tob;
begin
        Report := False;
        NbThemes := 0;
        For i := 1 to 40 do TableauQuest[i] := 0;
        Q := OpenSQL('SELECT POD_NUMSCORING,POD_LIBELLE,POD_THEMESCORE,POD_TABLESCORE FROM SCORINGDEF ORDER BY POD_THEMESCORE',True);
        TobQuestions := Tob.Create('',Nil,-1);
        TobQuestions.LoadDetailDB('SCORINGDEF','','',Q,False);
        Ferme(Q);

        If TobQuestions.Detail.Count < 16 Then TTabSheet(GetControl('PAGE2')).TabVisible := False; //PT5

        For i := 0 to TobQuestions.detail.Count -1 do
        begin
                TableauQuest[i+1] := TobQuestions.Detail[i].GetValue('POD_NUMSCORING');
                SetControlVisible('SCORE'+IntToStr(i+1),True);
                SetControlVisible('TSCORE'+IntToStr(i+1),True);
                SetControlCaption('TSCORE'+IntToStr(i+1),TobQuestions.detail[i].GetValue('POD_LIBELLE'));
                //DEBUT PT1
                If TobQuestions.Detail[i].GetValue('POD_TABLESCORE') = 'SC1' then
                SetControlProperty('SCORE'+IntToStr(i+1),'DataType','PGFSCORING1');
                If TobQuestions.Detail[i].GetValue('POD_TABLESCORE') = 'SC2' then
                SetControlProperty('SCORE'+IntToStr(i+1),'DataType','PGFSCORING2');
                If TobQuestions.Detail[i].GetValue('POD_TABLESCORE') = 'SC3' then
                SetControlProperty('SCORE'+IntToStr(i+1),'DataType','PGFSCORING3');
                If TobQuestions.Detail[i].GetValue('POD_TABLESCORE') = 'SC4' then
                SetControlProperty('SCORE'+IntToStr(i+1),'DataType','PGFSCORING4');
                If TobQuestions.Detail[i].GetValue('POD_TABLESCORE') = 'SC5' then
                SetControlProperty('SCORE'+IntToStr(i+1),'DataType','PGFSCORING5');
                //FIN PT1
                If (i = 0) or (i = 15) then
                begin
                        if i = 0 then NbThemes := 1
                        else nbThemes := 11;
                        SetControlVisible('CADRE'+IntToStr(nbThemes),True);
                        SetControlVisible('LIBELLE'+IntToStr(nbThemes),true);
                        if not Report then SetControlCaption('LIBELLE'+IntToStr(nbThemes),RechDom('PGFTHEMESCORING',TobQuestions.Detail[i].GetValue('POD_THEMESCORE'),False))
                        else SetControlCaption('LIBELLE'+IntToStr(nbThemes),RechDom('PGFTHEMESCORING',TobQuestions.Detail[i].GetValue('POD_THEMESCORE'),False)+' (suite)');
                end
                else
                begin
                        if TobQuestions.detail[i].GetValue('POD_THEMESCORE') <> TobQuestions.detail[i-1].GetValue('POD_THEMESCORE') then
                        begin
                                NbThemes := Nbthemes + 1;
                                SetControlVisible('CADRE'+IntToStr(NbThemes),True);
                                SetControlVisible('LIBELLE'+IntToStr(NbThemes),True);
                                SetControlCaption('LIBELLE'+IntToStr(NbThemes),RECHDOM('PGFTHEMESCORING',TobQuestions.Detail[i].GetValue('POD_THEMESCORE'),False));
                                Combo := THValComboBox(GetControl('SCORE'+IntToStr(i+1)));
                                Cadre := TBevel(GetControl('CADRE'+IntToStr(NbThemes)));
                                Cadre.Top := Combo.Top;
                                TLabel := THLabel(getcontrol('TSCORE'+IntToStr(i + 1)));
                                TLabelLib := THLabel(getcontrol('LIBELLE'+IntToStr(NbThemes)));
                                TLabelLib.Top := TLabel.Top;
                        end;
                end;
                If i = TobQuestions.Detail.Count - 1 then
                begin
                        Combo := THValComboBox(GetControl('SCORE'+IntToStr(i+1)));
                        Cadre := TBevel(GetControl('CADRE'+IntToStr(NbThemes)));
                        Cadre.Height := (Combo.Top -Cadre.Top) + Combo.Height;
                end
                else
                begin
                        if TobQuestions.detail[i].GetValue('POD_THEMESCORE') <> TobQuestions.detail[i+1].GetValue('POD_THEMESCORE') then
                        begin
                                Combo := THValComboBox(GetControl('SCORE'+IntToStr(i+1)));
                                Cadre := TBevel(GetControl('CADRE'+IntToStr(NbThemes)));
                                Cadre.Height := (Combo.Top -Cadre.Top) + Combo.Height;
                        end
                        else
                        begin
                                If i=14 then
                                begin
                                        Report := True;
                                        Combo := THValComboBox(GetControl('SCORE'+IntToStr(i+1)));
                                        Cadre := TBevel(GetControl('CADRE'+IntToStr(NbThemes)));
                                        Cadre.Height := (Combo.Top -Cadre.Top) + Combo.Height;
                                end;
                        end;
                end;
        end;
end;

procedure TOF_PGSCORING.AffichageChamps;
var Q : TQuery;
    TobScoring,TS : Tob;
    i : Integer;
    Req : String; //PT4
begin
	Q := OpenSQL('SELECT * FROM SCORING WHERE PSC_SALARIE="'+LeSalarie+'" AND PSC_CODESTAGE="'+LeStage+'" AND PSC_ORDRE='+LaSession+'',True);
	TobScoring := Tob.Create('LESTAGIAIRE',Nil,-1);
	TobScoring.LoadDetailDB('LeStagiaire','','',Q,False);
	Ferme(Q);

	TS := TobScoring.FindFirst(['PSC_SALARIE','PSC_CODESTAGE','PSC_ORDRE'],[LeSalarie,LeStage,LaSession],False);
	If TS <> Nil then
	begin
		For i := 1 to 40 do
		begin
			If TableauQuest[i] <> 0 then
			SetControlText('SCORE'+IntToStr(i),TobScoring.Detail[0].GetValue('PSC_SCFO'+IntToStr(TableauQuest[i])));
		end;
	end;
	If ToBScoring <> Nil then TobScoring.Free;
	
	Req := 'SELECT PSS_LIBELLE,PST_LIBELLE,PST_LIBELLE1 FROM SESSIONSTAGE LEFT JOIN STAGE ON PST_CODESTAGE=PSS_CODESTAGE AND PST_MILLESIME=PSS_MILLESIME'+
	   	   ' WHERE PSS_ORDRE='+LaSession+' AND PSS_CODESTAGE="'+LeStage+'"';
	If PGBundleInscFormation Then Req := Req + ' AND (PSS_PREDEFINI="STD" OR (PSS_PREDEFINI="DOS" AND PSS_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT4
	
	Q := OpenSQL(Req, True);
	If Not Q.Eof then
	begin
		SetControlText('SESSION',Q.FindField('PSS_LIBELLE').AsString);
		SetControlText('STAGE',Q.FindField('PST_LIBELLE').AsString+' '+Q.FindField('PST_LIBELLE1').AsString);
	end;
	Ferme(Q);
	
	Ecran.Caption := 'Evaluation du salarié : '+LeSalarie+' ';
	//PT4
	If PGBundleHierarchie Then
		Ecran.Caption := Ecran.Caption + RechDom('PGSALARIEINT',LeSalarie,False)
	Else
		Ecran.Caption := Ecran.Caption + RechDom('PGSALARIE',LeSalarie,False);
end;

procedure TOF_PGSCORING.MajScoring(Sender : TObject);
Var TobScoring : Tob;
    i,NumQuestion : Integer;
begin
        TobScoring := Tob.Create('SCORING',Nil,1);
        TobScoring.PutValue('PSC_SALARIE',LeSalarie);
        TobScoring.PutValue('PSC_CODESTAGE',LeStage);
        TobScoring.PutValue('PSC_ORDRE',StrToInt(LaSession));
        For i := 1 to 40 do
        begin
                NumQuestion := TableauQuest[i];
                If NumQuestion <> 0 then
                TobScoring.PutValue('PSC_SCFO'+IntToStr(NumQuestion),GetControlText('SCORE'+IntToStr(i)));
        end;
        TobScoring.InsertOrUpdateDB(False);
        TobScoring.Free;
end;

procedure TOF_PGSCORING.EditionScoring(Sender : TObject);
var TobEtat,T,TobQuest,TQ : Tob;
    i : Integer;
    ThemePrec,Theme,Quest,Tablette,Rep : String;
    NumScoring : Integer;
    Pages : TPageControl;
    Q : TQuery;
    Req : String; //PT4
begin
     Pages := TPageControl(GetControl('PAGES'));
     
     Q := OpenSQL('SELECT POD_NUMSCORING,POD_THEMESCORE,POD_TABLESCORE FROM SCORINGDEF',True);
     TobQuest := Tob.Create('LesQuestions',Nil,-1);
     TobQuest.LoadDetailDB('LesQuestions','','',Q,False);
     Ferme(Q);
     
     TobEtat := Tob.Create('Edition',Nil,-1);
     T := Tob.Create('FilleEdition',TobEtat,-1);
     T.AddChampSupValeur('PSS_CODESTAGE',LeStage);
     T.AddChampSupValeur('PFO_SALARIE',LeSalarie);
     
     Req := 'SELECT PSS_CENTREFORMGU,PSS_NATUREFORM,PSS_DATEDEBUT,PSS_DATEFIN,PSS_LIBELLE FROM SESSIONSTAGE WHERE '+
     		'PSS_CODESTAGE="'+LeStage+'" AND PSS_ORDRE="'+LaSession+'"';
     If PGBundleInscFormation Then Req := Req + ' AND (PSS_PREDEFINI="STD" OR (PSS_PREDEFINI="DOS" AND PSS_NODOSSIER="'+V_PGI.NoDossier+'"))'; //PT4

     Q := OpenSQL(Req, True);
     If Not Q.Eof then
     begin
          T.AddChampSupValeur('PSS_NATUREFORM',Q.findField('PSS_NATUREFORM').AsString);
          T.AddChampSupValeur('PSS_LIBELLE',Q.findField('PSS_LIBELLE').AsString);
          T.AddChampSupValeur('PSS_CENTREFORMGU',Q.findField('PSS_CENTREFORMGU').AsString);
          T.AddChampSupValeur('PSS_DATEDEBUT',Q.findField('PSS_DATEDEBUT').AsDateTime);
          T.AddChampSupValeur('PSS_DATEFIN',Q.findField('PSS_DATEFIN').AsDateTime);
     end
     else
     begin
          T.AddChampSupValeur('PSS_NATUREFORM','');
          T.AddChampSupValeur('PSS_LIBELLE','');
          T.AddChampSupValeur('PSS_CENTREFORMGU','');
          T.AddChampSupValeur('PSS_DATEDEBUT',IDate1900);
          T.AddChampSupValeur('PSS_DATEFIN',IDate1900);
     end;
     ferme(Q);
     ThemePrec := '';
     
     For i := 1 to 30 do
     begin
          NumScoring := TableauQuest[i];
          TQ := TobQuest.FindFirst(['POD_NUMSCORING'],[IntToStr(NumScoring)],False);
          If TQ <> Nil then
          begin
               Theme := TQ.GetValue('POD_THEMESCORE');
               Tablette := TQ.GetValue('POD_TABLESCORE');
               If THLabel(GetControl('TSCORE'+IntToStr(i))) <> nil then Quest := THLabel(GetControl('TSCORE'+IntToStr(i))).Caption
               else Quest := '';
               Rep := GetControlText('SCORE'+IntToStr(i));
               If (ThemePrec = Theme) then Theme := ''
               else
               begin
                    ThemePrec := Theme;
                    Theme := RechDom('PGFTHEMESCORING',Theme,False);
               end;
               If Tablette = 'SC1' then Tablette := 'PGFSCORING1'
               else If Tablette = 'SC2' then Tablette := 'PGFSCORING2'
               else If Tablette = 'SC3' then Tablette := 'PGFSCORING3'
               else If Tablette = 'SC4' then Tablette := 'PGFSCORING4'
               else If Tablette = 'SC5' then Tablette := 'PGFSCORING5';
               T.AddChampSupValeur('THEME'+IntToStr(i),Theme);
               T.AddChampSupValeur('QUEST'+IntToStr(i),Quest);
               T.AddChampSupValeur('REPONSE'+IntToStr(i),RechDom(Tablette,Rep,False));
          end
          else
          begin
               T.AddChampSupValeur('THEME'+IntToStr(i),'');
               T.AddChampSupValeur('QUEST'+IntToStr(i),'');
               T.AddChampSupValeur('REPONSE'+IntToStr(i),'');
          end;
     end;
     TobQuest.Free;
     LanceEtatTOB('E','PFO','POF',TobEtat,True,False,False,Pages,'','',False);
     TobEtat.Free;
end;

Initialization
  registerclasses ( [ TOF_PGSCORING ] ) ;
end.

