{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/09/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TESTPLAN ()
Mots clefs ... : TOF;TESTPLAN                                                                                
*****************************************************************
PT1 | 11/04/2005 | V_60  | JL | FQ 12140 Gestion des états des sessions
--- | 17/10/2006 |       | JL | Modification contrôle des exercices de formations -----
PT2 | 28/09/2007 | V_7   | FL | Emanager / Report / Adaptation cursus + accès assistant
PT3 | 27/10/2007 | V_8   | FL | Emanager / Report / Gestion des favoris + intégration du planning dans la fenêtre
PT4 | 04/12/2007 | V_8   | FL | Emanager / Report / Seul le groupe doit être visible parmi les champs libres formation
PT5 | 18/02/2008 | V_803 | FL | Correction des critères Predefini/NoDossier pour la gestion multidossier
PT6 | 03/04/2008 | V_803 | FL | Gestion des groupes de travail
PT7 | 17/04/2008 | V_804 | FL | Prise en compte du bundle Catalogue
}
Unit UTofPGPlanningFormation ;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     Fe_Main,
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,PGPlanningSessionForm,EntPaie,PGOutilsFormation,HQry,PGEdtEtat,Windows ;

Type
  TOF_PGPLANNINGFORMATION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    Private
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState); //PT2
    //PT3 - Début
    procedure OnExport (Sender : TObject);
    procedure OnDetail (Sender : TObject);
    procedure OnPrint  (Sender : TObject);
    {$IFDEF EMANAGER}
    procedure ClickStagesFavoris (Sender : TObject);
    procedure ClickCursusFavoris (Sender : TObject);
    procedure ChangeCursus (Sender : TObject);
    procedure ChangeStage  (Sender : TObject);
    {$ENDIF}
    //PT3 - Fin
  end ;

Implementation
Uses Hpanel, HTB97,GalOutil;

procedure TOF_PGPLANNINGFORMATION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGPLANNINGFORMATION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGPLANNINGFORMATION.OnUpdate ;
Var NiveauRupt : TNiveauRupture;
    WhereItems,WhereEtats,WheresRess : String;
    Pages : TPageControl;
    WhereEtat,LesEtats,Etat : String;
//    {$IFDEF EMANAGER}
    St : String; //PT2
//    {$ENDIF}
begin
  Inherited ;
{        If GetControlText('RUPTURE1')<>'' then NiveauRupt.NiveauRupt:=1;
        If GetControlText('RUPTURE2')<>'' then NiveauRupt.NiveauRupt:=2;
        If GetControlText('RUPTURE3')<>'' then NiveauRupt.NiveauRupt:=3;
        If GetControlText('RUPTURE4')<>'' then NiveauRupt.NiveauRupt:=4;
        NiveauRupt.ChampsRupt[1]:=GetControlText('RUPTURE1');
        NiveauRupt.ChampsRupt[2]:=GetControlText('RUPTURE2');
        NiveauRupt.ChampsRupt[3]:=GetControlText('RUPTURE3');
        NiveauRupt.ChampsRupt[4]:=GetControlText('RUPTURE4');}
   //     NiveauRupt.CondRupt:=''; NiveauRupt.NiveauRupt:=0;
        Pages := TPageControl(GetControl('PAGES'));
        WhereItems := RecupWhereCritere(Pages);
        If WhereItems <> '' then WheresRess := ConvertPrefixe(WhereItems,'PSS','PST')
        else WheresRess := '';
        If WheresRess <> '' then WheresRess := WheresRess+' AND PST_ACTIF="X"'
        else WheresRess := 'WHERE PST_ACTIF="X"';
    //PT2 - Début
//    {$IFDEF EMANAGER}
    St := GetControlText('CURSUS');
    If (St <> '') And (UpperCase(St) <> '<<TOUS>>') Then
    Begin
        If WheresRess <> '' then WheresRess := WheresRess+' AND ' Else WheresRess := 'WHERE ';
        WheresRess := WheresRess + 'PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS IN (';
        While St <> '' Do WheresRess := WheresRess + '"' + ReadTokenSt(St) + '",';
        System.Delete(WheresRess,Length(WheresRess),1);
        WheresRess := WheresRess + '))';
    End;
    {$IFDEF EMANAGER}
    If GetControlText('CCURSUSFAV') = 'X' Then
    Begin
        If WheresRess <> '' then WheresRess := WheresRess+' AND ' Else WheresRess := 'WHERE ';
        WheresRess := WheresRess + 'PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS IN ('+
        'SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="-C-" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'"))';
    End;
    If GetControlText('CSTAGEFAV') = 'X' Then
    Begin
        If WheresRess <> '' then WheresRess := WheresRess+' AND ' Else WheresRess := 'WHERE ';
        WheresRess := WheresRess + 'PST_CODESTAGE IN (SELECT PCC_CODESTAGE FROM CURSUSSTAGE WHERE PCC_CURSUS="---" AND PCC_LIBELLE="'+V_PGI.UserSalarie+'")';
    End;
    {$ENDIF}
    //PT2 - Fin
     {   If WhereItems <> '' then WhereItems := WhereItems+' AND ((PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                           ' AND PSS_DATEDEBUT<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'") OR '+
                           '(PSS_DATEFIN>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                           ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"))'
        else WhereItems := 'WHERE (PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                           ' AND PSS_DATEDEBUT<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'") OR '+
                           '(PSS_DATEFIN>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                           ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'")';}
        If WhereItems <> '' then WhereItems := WhereItems+' AND PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                           ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"'
        else WhereItems := 'WHERE PSS_DATEDEBUT>="'+UsDateTime(StrToDate(GetControlText('DATEDEBUT')))+'"'+
                           ' AND PSS_DATEFIN<="'+UsDateTime(StrToDate(GetControlText('DATEFIN')))+'"';
        If GetControlText('ETATFORMATION') <> '' then
        begin
             LesEtats := GetControlText('ETATFORMATION');
             WhereEtat := '';
             While LesEtats <> '' do
             begin
                  Etat := ReadTokenPipe(LesEtats,';');
                  If Etat = 'INC' then
                  begin
                       If WhereEtat <> '' then WhereEtat := WhereEtat+' OR PSS_CLOTUREINSC="-"'
                       else WhereEtat := ' AND (PSS_CLOTUREINSC="-"';
                  end;
                  If Etat = 'IC' then
                  begin
                       If WhereEtat <> '' then WhereEtat := WhereEtat+' OR (PSS_CLOTUREINSC="X" AND PSS_EFFECTUE="-")'
                       else WhereEtat := ' AND ((PSS_CLOTUREINSC="X" AND PSS_EFFECTUE="-")';
                  end;
                  If Etat = 'SE' then
                  begin
                       If WhereEtat <> '' then WhereEtat := WhereEtat+' OR (PSS_EFFECTUE="X" AND PSS_VALIDORG="-")'
                       else WhereEtat := ' AND ((PSS_EFFECTUE="X" AND PSS_VALIDORG="-")';
                  end;
                  If Etat = 'VOC' then
                  begin
                       If WhereEtat <> '' then WhereEtat := WhereEtat+' OR PSS_VALIDORG="X"'
                       else WhereEtat := ' AND (PSS_VALIDORG="X"';
                  end;
             end;
             If WhereEtat <> '' then
             begin
                  WhereEtat := WhereEtat + ')';
                  WhereItems := Whereitems + WhereEtat;
             end;                                                                                                                
        end;
    //PT2 - Début
    {$IFDEF EMANAGER}
    WhereItems := WhereItems + ' AND PSS_COMPTERENDU="X"';
    {$ENDIF}
    //PT2 - Fin
        If GetCheckBoxState('CAFFICHE') <> CbChecked then WheresRess := WheresRess+' AND PST_CODESTAGE IN (SELECT PSS_CODESTAGE FROM SESSIONSTAGE '+WhereItems+')';
        WhereEtats := '';
        
    //PT5 - Début
    //If GetControlText('PREDEFINI') <> '' Then
    //Begin
    	WhereItems := WhereItems + DossiersAInterroger(GetControlText('PREDEFINI'),GetControlText('NODOSSIER'),'PSS',True,True);
    	WheresRess := WheresRess + DossiersAInterroger(GetControlText('PREDEFINI'),GetControlText('NODOSSIER'),'PST',True,True);
    //End
    //Else
    //Begin
//		WhereItems := WhereItems + DossiersAInterroger(GetControlText('PREDEFINI'),GetControlText('NODOSSIER'),'PSS') + ')';
		//WheresRess := WheresRess + DossiersAInterroger(GetControlText('PREDEFINI'),GetControlText('NODOSSIER'),'PST') + ')';
    //End;
    //PT5 - Fin
    
    PGPlanningFormation(THPanel(GetControl('PANPLANNING')),StrToDate(GetControlText('DATEDEBUT')),StrToDate(GetControlText('DATEFIN')),WhereEtats,WhereItems,WheresRess,NiveauRupt,False);  //PT2
end ;

procedure TOF_PGPLANNINGFORMATION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGPLANNINGFORMATION.OnArgument (S : String ) ;
var Num : Integer;
    DD,DF : TDateTime;
    Bt : TToolbarButton97;
    {$IFDEF EMANAGER}
    Ck : TCheckbox;
    MCb : THMultiValComboBox;
    Cb : THEdit;
    {$ENDIF}
begin
    Inherited ;
    RendMillesimeRealise(DD,DF);
    SetControlText('DATEDEBUT',DateToStr(DD));
    SetControlText('DATEFIN',DateToStr(DF));
    {$IFNDEF EMANAGER} //PT4    
    For Num  :=  1 to VH_Paie.NBFormationLibre do
    begin
        if Num > 8 then Break;
        VisibiliteChampFormation (IntToStr(Num),GetControl ('PSS_FORMATION'+IntToStr(Num)),GetControl ('TPSS_FORMATION'+IntToStr(Num)));  //PT4
    end;
    {$ENDIF}
    
	If PGBundleCatalogue Then //PT7
	Begin
	 	If Not PGDroitMultiForm Then
	 	Begin
			SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True));
			SetControlProperty ('CURSUS', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True));
		End
		Else If V_PGI.ModePCL='1' Then
		Begin
       		SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6
       		SetControlProperty ('CURSUS', 'Plus', DossiersAInterroger('','','PCU',True,True)); //PT6
		End;
	End;
	
	If PGBundleInscFormation then
	begin
		If not PGDroitMultiForm then
		begin
			SetControlEnabled('NODOSSIER',False);
			SetControlText   ('NODOSSIER',V_PGI.NoDossier);
			SetControlEnabled('PREDEFINI',False);
			SetControlText   ('PREDEFINI','');
			//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PST',True,True)); //PT6 //PT7
			//SetControlProperty ('CURSUS', 'Plus', DossiersAInterroger('',V_PGI.NoDossier,'PCU',True,True)); //PT6 //PT7
		end
       	Else If V_PGI.ModePCL='1' Then
       	Begin
       		SetControlProperty('NODOSSIER', 'Plus', GererCritereGroupeConfTous); //PT6
       		//SetControlProperty ('PSS_CODESTAGE', 'Plus', DossiersAInterroger('','','PST',True,True)); //PT6 //PT7
       		//SetControlProperty ('CURSUS', 'Plus', DossiersAInterroger('','','PCU',True,True)); //PT6 //PT7
       	End;
	end
	else
	begin
		SetControlVisible('PREDEFINI', 	   False);
		SetControlVisible('TPREDEFINI',	   False);
		SetControlVisible('NODOSSIER',     False);
		SetControlVisible('TNODOSSIER',    False);
	end;
    //PT2 - Début
    // Gestion des raccourcis clavier
    Ecran.OnKeyDown := FormKeyDown;
    //PT2 - Fin

    //PT3 - Début
//    {$IFDEF EMANAGER}
    Bt := TToolbarButton97 (GetControl('bAgrandir'));
    If Bt <> Nil Then
    Begin
        BT.Visible := True;
        BT.onClick := OnDetail;
    End;

    Bt := TToolbarButton97 (GetControl('bExport'));
    If Bt <> Nil Then
    Begin
        BT.Visible := True;
        BT.onClick := OnExport;
    End;

    Bt := TToolbarButton97 (GetControl('bImprimer'));
    If Bt <> Nil Then
    Begin
        BT.Visible := True;
        BT.onClick := OnPrint;
    End;

    {$IFDEF EMANAGER}
    Ck := TCheckBox (GetControl('CSTAGEFAV'));
    If Ck <> Nil Then Ck.OnClick := ClickStagesFavoris;

    Ck := TCheckBox (GetControl('CCURSUSFAV'));
    If Ck <> Nil Then Ck.OnClick := ClickCursusFavoris;

    MCb := THMultiValComboBox(GetControl('CURSUS'));
    If MCb <> Nil Then MCb.OnChange := ChangeCursus;

    Cb := THEdit(GetControl('PSS_CODESTAGE'));
    If Cb <> Nil Then CB.OnChange := ChangeStage;
    {$ENDIF}
    //PT3 -Fin
end ;

procedure TOF_PGPLANNINGFORMATION.OnClose ;
begin
  Inherited ;
    PGPlanningFormationDestroy
end ;

procedure TOF_PGPLANNINGFORMATION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGPLANNINGFORMATION.OnCancel () ;
begin
  Inherited ;
end ;

//PT2 - Début
procedure TOF_PGPLANNINGFORMATION.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    case Key of
    VK_F9,VK_F10:OnUpdate;
End;
End;
//PT2 - Fin

//PT3 - Début
//{$IFDEF EMANAGER}
procedure TOF_PGPLANNINGFORMATION.OnDetail(Sender: TObject);
begin
    PGPlanningFormationShowDetail;
end;

procedure TOF_PGPLANNINGFORMATION.OnExport(Sender: TObject);
begin
    PGPlanningFormationExport;
end;

procedure TOF_PGPLANNINGFORMATION.OnPrint(Sender: TObject);
begin
    PGPlanningFormationImprimer;
end;

{$IFDEF EMANAGER}
procedure TOF_PGPLANNINGFORMATION.ClickCursusFavoris(Sender: TObject);
begin
    If TCheckBox(Sender).Checked = True Then
    Begin
        TCheckBox(GetControl('CSTAGEFAV')).Checked := False;
        SetControlText('CURSUS', '');
    End;
end;

procedure TOF_PGPLANNINGFORMATION.ClickStagesFavoris(Sender: TObject);
begin
    If TCheckBox(Sender).Checked = True Then
    Begin
        TCheckBox(GetControl('CCURSUSFAV')).Checked := False;
        SetControlText('PSS_CODESTAGE', '');
    End;
end;

procedure TOF_PGPLANNINGFORMATION.ChangeCursus(Sender: TObject);
begin
    If THMultiValComboBox(Sender).Value <> '' Then SetControlText('CCURSUSFAV', '-');
end;

procedure TOF_PGPLANNINGFORMATION.ChangeStage(Sender: TObject);
begin
    If THEdit(Sender).Text <> '' Then SetControlText('CSTAGEFAV', '-');
end;
{$ENDIF}
//PT3 - Fin

Initialization
  registerclasses ( [ TOF_PGPLANNINGFORMATION ] ) ;
end.

