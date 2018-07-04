{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 27/05/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGTESTEDITFORM ()
Mots clefs ... : TOF;PGTESTEDITFORM
*****************************************************************
PT1 27/09/2006 JL V_750 FQ 13542 Modification des titres affichés
PT2 : 25/09/2009 SJ FQ n°12376 Suppression du LanceEtatTob
}
Unit UTOFPGEdittionsCompetences ;

Interface

Uses StdCtrls,Controls,Classes, UTOB,
{$IFDEF EAGLCLIENT}
     eQRS1,UtilEAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,LookUp,
     PGOutilsFormation,P5Def,PGOutils,PGOutils2,PGEditOutils,HQry,HTB97,ParamDat,PGEdtEtat,ParamSoc, ed_tools, HStatus;

Type
  TOF_PGEDITEMPLOICOMP = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

  TOF_PGEDITSTAGECOMPETENCE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

  TOF_PGEDITCOMPETENCESINTER = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

  TOF_PGEDITINTERCOMPETENCES = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  end ;

  TOF_PGEDITCOMPSALEMPLOI = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ;//PT2
    Procedure OnClose                  ; override ;//PT2
    private
    TobEtat : Tob;//PT2
    procedure EditionComparatif (Sender : TObject);
    procedure DateElipsisClick(Sender : TObject);
    procedure ExitEdit(Sender : TObject);
    procedure SalarieElipsisClick(Sender : TObject);
  end ;

  Type
  TOF_PGEDITCOMPEMPLOI = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Arg : String;
    Function Conditionform(Champ : String) : String;
  end ;


Implementation

Procedure TOF_PGEDITEMPLOICOMP.OnUpdate ;
var Where,WhereTmp,OrderBy : String;
    i : Integer;
    MultiCombo : THMultiValComboBox;
    Condition,StCondition : String;
begin

  Inherited ;
         OrderBy := '';
         // Gestion clause WHERE pour affichage des compétences dans bande détail
        Where := '';
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('COMPETENCE'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_COMPETENCE="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_COMPETENCE="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        For i := 1 to 5 do
        begin
                MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMP'+IntToStr(i)));
                Condition := MultiCombo.Value;
                While Condition <> '' do
                begin
                        StCondition := ReadTokenPipe(Condition,';');
                        If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_TABLELIBRERH'+IntToStr(i)+'="'+StCondition+'"'
                        else WhereTmp := 'AND (PCO_TABLELIBRERH'+IntToStr(i)+'="'+StCondition+'"';
                end;
                If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
                Where := Where + WhereTmp;
                WhereTmp := '';
        end;
        For i := 1 to 3 do
        begin
                MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMPFOR'+IntToStr(i)));
                Condition := MultiCombo.Value;
                While Condition <> '' do
                begin
                        StCondition := ReadTokenPipe(Condition,';');
                        If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR POS_TABLELIBRECR'+IntToStr(i)+'="'+StCondition+'"'
                        else WhereTmp := 'AND (POS_TABLELIBRECR'+IntToStr(i)+'="'+StCondition+'"';
                end;
                If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
                Where := Where + WhereTmp;
                WhereTmp := '';
        end;
        SetControlText('CONDITIONCOMP',Where);
         // Gestion clause WHERE pour affichage des salariés dans bande détail
        Where := '';
        MultiCombo  := THMultiValComboBox(GetControl('ETABLISSEMENT'));
        Condition := MultiCombo.Value;
        WhereTmp := '';
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_ETABLISSEMENT="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_ETABLISSEMENT="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('CODESTAT'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_CODESTAT="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_CODESTAT="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        For i := 1 to VH_Paie.PGNbreStatOrg do
        begin
                MultiCombo  := THMultiValComboBox(GetControl('TRAVAILN'+IntToStr(i)));
                Condition := MultiCombo.Value;
                While Condition <> '' do
                begin
                        StCondition := ReadTokenPipe(Condition,';');
                        If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_TRAVAILN'+IntToStr(i)+'="'+StCondition+'"'
                        else WhereTmp := 'AND (PSA_TRAVAILN'+IntToStr(i)+'="'+StCondition+'"';
                end;
                If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
                Where := Where + WhereTmp;
                WhereTmp := '';
        end;
        SetControlText('CONDITIONSAL',Where);
end;

procedure TOF_PGEDITEMPLOICOMP.OnArgument (S : String ) ;
var Num : Integer;
    Arg : String;
begin
  Inherited ;
  Arg := ReadTokenPipe(S,';');
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  For Num  :=  1 to 5 do
  begin
        VisibiliteChampCompetence (IntToStr(Num),GetControl ('LIBRECOMP'+IntToStr(Num)),GetControl ('TLIBRECOMP'+IntToStr(Num)));
  end;
  For Num  :=  1 to 3 do
  begin
        VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('LIBRECOMPFOR'+IntToStr(Num)),GetControl ('TLIBRECOMPFOR'+IntToStr(Num)));
  end;
  For Num := 1 to VH_Paie.NBFormationLibre do
  begin
        if Num > 8 then Break;
        VisibiliteChampFormation (IntToStr(Num),GetControl ('PST_FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
  end;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('TRAVAILN'+IntToStr(Num)),GetControl ('TTRAVAILN'+IntToStr(Num)));
    end;
    VisibiliteStat(GetControl ('CODESTAT'),GetControl ('TCODESTAT'));
end ;

{TOF_PGEDITSTAGECOMPETENCE}

Procedure TOF_PGEDITSTAGECOMPETENCE.OnUpdate ;
var Where,WhereTmp,OrderBy : String;
    MultiCombo : THMultiValComboBox;
    Condition,StCondition : String;
begin

  Inherited ;
         OrderBy := '';
         // Gestion clause WHERE pour affichage des compétences dans bande détail
        Where := '';
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('COMPETENCE'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_COMPETENCE="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_COMPETENCE="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMP1'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_TABLELIBRERH1="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_TABLELIBRERH1="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMP2'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_TABLELIBRERH2="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_TABLELIBRERH2="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMP3'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_TABLELIBRERH3="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_TABLELIBRERH3="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMP4'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_TABLELIBRERH4="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_TABLELIBRERH4="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMP5'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PCO_TABLELIBRERH5="'+StCondition+'"'
                else WhereTmp := 'AND (PCO_TABLELIBRERH5="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        SetControlText('CONDITIONCOMP',Where);
         // Gestion clause WHERE pour affichage des salariés dans bande détail
        Where := '';
        MultiCombo  := THMultiValComboBox(GetControl('ETABLISSEMENT'));
        Condition := MultiCombo.Value;
        WhereTmp := '';
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_ETABLISSEMENT="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_ETABLISSEMENT="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('CODESTAT'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_CODESTAT="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_CODESTAT="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('TRAVAILN1'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_TRAVAILN1="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_TRAVAILN1="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('TRAVAILN2'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_TRAVAILN2="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_TRAVAILN2="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('TRAVAILN3'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_TRAVAILN3="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_TRAVAILN3="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('TRAVAILN4'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_TRAVAILN4="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_TRAVAILN4="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        SetControlText('CONDITIONSAL',Where);
end;

procedure TOF_PGEDITSTAGECOMPETENCE.OnArgument (S : String ) ;
var Num : Integer;
begin
  Inherited ;
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  For Num  :=  1 to 5 do
        begin
                VisibiliteChampCompetence (IntToStr(Num),GetControl ('LIBRECOMP'+IntToStr(Num)),GetControl ('TTLIBRECOMP'+IntToStr(Num)));
        end;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('TRAVAILN'+IntToStr(Num)),GetControl ('TTRAVAILN'+IntToStr(Num)));
    end;
    VisibiliteStat(GetControl ('CODESTAT'),GetControl ('TCODESTAT'));
end ;

{TOF_PGEDITCOMPETENCESINTER}
procedure TOF_PGEDITCOMPETENCESINTER.OnArgument (S : String ) ;
var Num : Integer;
begin
  Inherited ;
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  SetControlText('PSI_TYPEINTERIM','SAL');
  For Num  :=  1 to 5 do
        begin
                VisibiliteChampCompetence (IntToStr(Num),GetControl ('PCO_TABLELIBRERH'+IntToStr(Num)),GetControl ('TPCO_TABLELIBRERH'+IntToStr(Num)));
        end;
  For Num  :=  1 to 3 do
        begin
                VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('PCH_TABLELIBRECR'+IntToStr(Num)),GetControl ('TPCH_TABLELIBRECR'+IntToStr(Num)));
        end;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
    end;
    VisibiliteStat(GetControl ('PSI_CODESTAT'),GetControl ('TPSI_CODESTAT'));
    SetControlText('PSI_DATEENTREE',DatetoStr(IDate1900));
    SetControlText('PSI_DATEENTREE_',DatetoStr(Date));
    SetControlText('PCH_DATEVALIDATION',DatetoStr(IDate1900));
    SetControlText('PCH_DATEVALIDATION_',DatetoStr(Date));
end ;

{TOF_PGEDITINTERCOMPETENCES}
procedure TOF_PGEDITINTERCOMPETENCES.OnArgument (S : String ) ;
var Num : Integer;
    TLabel : THLabel;
begin
  Inherited ;
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  SetControlText('PSI_TYPEINTERIM','SAL');
  For Num  :=  1 to 5 do
        begin
                VisibiliteChampCompetence (IntToStr(Num),GetControl ('PCO_TABLELIBRERH'+IntToStr(Num)),GetControl ('TPCO_TABLELIBRERH'+IntToStr(Num)));
        end;
  For Num  :=  1 to 3 do
        begin
                VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('PCH_TABLELIBRECR'+IntToStr(Num)),GetControl ('TPCH_TABLELIBRECR'+IntToStr(Num)));
        end;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
        TLabel := THLabel(GetControl('TPCO_TABLELIBRERH'+IntToStr(Num)));
        If TLabel.Visible = True then SetControlText('LIBCO'+IntToStr(Num),TLabel.Caption)
        else SetControlText('LIBCO'+IntToStr(Num),'');
    end;
    VisibiliteStat(GetControl ('PSI_CODESTAT'),GetControl ('TPSI_CODESTAT'));
    SetControlText('PSI_DATEENTREE',DatetoStr(IDate1900));
    SetControlText('PSI_DATEENTREE_',DatetoStr(Date));
    SetControlText('PCH_DATEVALIDATION',DatetoStr(IDate1900));
    SetControlText('PCH_DATEVALIDATION_',DatetoStr(Date));
end ;

{TOF_PGEDITCOMPSALEMPLOI}
procedure TOF_PGEDITCOMPSALEMPLOI.OnArgument (S : String ) ;
var {BEtat : TToolBarButton97;}// PT2 mise en commentaire
    Edit : THEdit;
begin
  Inherited ;
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
{  SetControlText('PSI_TYPEINTERIM','SAL');
  For Num  :=  1 to 5 do
        begin
                VisibiliteChampCompetence (IntToStr(Num),GetControl ('PCO_TABLELIBRERH'+IntToStr(Num)),GetControl ('TPCO_TABLELIBRERH'+IntToStr(Num)));
        end;
  For Num  :=  1 to 3 do
        begin
                VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('PCH_TABLELIBRECR'+IntToStr(Num)),GetControl ('TPCH_TABLELIBRECR'+IntToStr(Num)));
        end;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSI_TRAVAILN'+IntToStr(Num)),GetControl ('TPSI_TRAVAILN'+IntToStr(Num)));
        TLabel := THLabel(GetControl('TPCO_TABLELIBRERH'+IntToStr(Num)));
        If TLabel.Visible = True then SetControlText('LIBCO'+IntToStr(Num),TLabel.Caption)
        else SetControlText('LIBCO'+IntToStr(Num),'');
    end;
    VisibiliteStat(GetControl ('PSI_CODESTAT'),GetControl ('TPSI_CODESTAT'));
    SetControlText('PSI_DATEENTREE',DatetoStr(IDate1900));
    SetControlText('PSI_DATEENTREE_',DatetoStr(Date));
    SetControlText('PCH_DATEVALIDATION',DatetoStr(IDate1900));
    SetControlText('PCH_DATEVALIDATION_',DatetoStr(Date));  }
   { Betat  := TToolBarButton97(GetControl('BVALIDER'));
   If BEtat <> Nil then BEtat.OnClick  := EditionComparatif;}// PT2 mise en commentaire
    Edit := THEdit(GetControl('DATEVALIDATION'));
    If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
    Edit := THEdit(GetControl('DATEVALIDATION'));
    If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
    Edit := THEdit(GetControl('PSI_INTERIMAIRE_'));
    If Edit <> Nil then Edit.OnExit := ExitEdit;
    Edit := THEdit(GetControl('PSI_INTERIMAIRE'));
    If Edit <> Nil then
    begin
        Edit.OnExit := ExitEdit;
        Edit.OnElipsisClick := SalarieElipsisClick;
    end;
    SetControlText('PSI_DATEENTREE',DateToStr(IDate1900));
    SetControlText('PSI_DATEENTREE_',DateToStr(Date));
    SetControlText('DATEVALIDATION',DateToStr(IDate1900));
    SetControlText('DATEVALIDATION_',DateToStr(Date));

end ;

procedure TOF_PGEDITCOMPSALEMPLOI.DateElipsisclick(Sender : TObject);
var key : char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCOMPSALEMPLOI.ExitEdit(Sender: TObject);
var edit : thedit;
begin
        edit:=THEdit(Sender);
        if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
        if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITCOMPSALEMPLOI.SalarieElipsisClick(Sender : TObject);
begin
        LookupList(THEdit(Sender),'Liste des salariés','INTERIMAIRES','PSI_SALARIE','PSI_LIBELLE,PSI_PRENOM','','', True,-1);
end;

procedure TOF_PGEDITCOMPSALEMPLOI.EditionComparatif (Sender : TObject);
var TobSalaries,{TobEtat,PT2 mise en com.}TE,TobCompetSal,TobCompetEmploi : Tob;
    Q : TQuery;
    Pages : TPageControl;
    i,c : Integer;
    Salarie,Emploi,Competence,Where,LeType,Nom,Prenom,Etab : String;
begin
If TobEtat <> Nil Then FreeAndNil(TobEtat); //PT2
        Pages := TPageControl(GetControl('PAGES'));
        Where := RecupWhereCritere(Pages);
        TobEtat := Tob.Create('Edition',Nil,-1);
        Q := OpenSQL('SELECT * FROM INTERIMAIRES '+Where,True);
        TobSalaries := Tob.Create('LesSalaries',Nil,-1);
        TobSalaries.LoadDetailDB('LesSalaries','','',Q,False);
        Ferme(Q);
        InitMoveProgressForm (nil, 'Chargement des données',
                         'Veuillez patienter SVP ...',
                         TobSalaries.Detail.Count, False, True);
        InitMove(TobSalaries.Detail.Count, '');
        For i := 0 to TobSalaries.Detail.Count - 1 do
        begin
                Salarie := TobSalaries.Detail[i].GetValue('PSI_INTERIMAIRE');
                Nom := TobSalaries.detail[i].GetValue('PSI_LIBELLE');
                Prenom := TobSalaries.detail[i].GetValue('PSA_PRENOM');
                Emploi := TobSalaries.Detail[i].GetValue('PSI_LIBELLEEMPLOI');
                LeType := TobSalaries.Detail[i].GetValue('PSI_TYPEINTERIM');
                Etab := TobSalaries.detail[i].GetValue('PSI_ETABLISSEMENT');
                Q := OpenSQL('SELECT * FROM RHCOMPETRESSOURCE WHERE PCH_TYPERESSOURCE="'+LeType+'" AND PCH_SALARIE="'+Salarie+'"',True);
                TobCompetSal := Tob.Create('LesCompetences',Nil,-1);
                TobCompetSal.LoadDetailDB('LesCompetences','','',Q,False);
                Ferme(Q);
                For c := 0 to TobCompetSal.Detail.Count - 1 do
                begin
                        Competence := TobCompetSal.Detail[c].GetValue('PCH_COMPETENCE');
                        TE := TobEtat.FindFirst(['PSA_SALARIE','PCO_COMPETENCE'],[Salarie,Competence],False);
                        If TE <> Nil then
                        begin
                                If TE.GetValue('MAITRISESAL') < TobCompetSal.detail[c].GetValue('PCH_DEGREMAITRISE') then
                                TE.PutValue('MAITRISESAL',TobCompetSal.detail[c].GetValue('PCH_DEGREMAITRISE'));
                        end
                        else
                        begin
                                TE := Tob.Create('UneCompetence',TobEtat,-1);
                                TE.AddChampSupValeur('PSA_SALARIE',Salarie);
                                TE.AddChampSupValeur('PSA_LIBELLEEMPLOI',Emploi);
                                TE.AddChampSupValeur('PCO_COMPETENCE',Competence);
                                TE.AddChampSupValeur('PSA_LIBELLE',Nom);
                                TE.AddChampSupValeur('PSA_PRENOM',Prenom);
                                TE.AddChampSupValeur('PSA_ETABLISSEMENT',Etab);
                                TE.AddChampSupValeur('PCH_DATEVALIDATION',TobCompetSal.detail[c].GetValue('PCH_DATEVALIDATION'));
                                TE.AddChampSupValeur('MAITRISESAL',TobCompetSal.detail[c].GetValue('PCH_DEGREMAITRISE'));
                                TE.AddChampSupValeur('MAITRISEEMP',0);
                                TE.AddChampSupValeur('PCH_TABLELIBRECR1',TobCompetSal.detail[c].GetValue('PCH_TABLELIBRECR1'));
                                TE.AddChampSupValeur('PCH_TABLELIBRECR2',TobCompetSal.detail[c].GetValue('PCH_TABLELIBRECR2'));
                                TE.AddChampSupValeur('PCH_TABLELIBRECR3',TobCompetSal.detail[c].GetValue('PCH_TABLELIBRECR3'));
                        end;
                end;
                TobCompetSal.Free;
                Q := OpenSQL('SELECT * FROM STAGEOBJECTIF WHERE POS_NATOBJSTAGE="EMP" AND POS_CODESTAGE="'+Emploi+'" ',True);
                TobCompetEmploi := Tob.Create('LesCompetencesEmp',Nil,-1);
                TobCompetEmploi.LoadDetailDB('LesCompetencesEmp','','',Q,False);
                Ferme(Q);
                For c := 0 to TobCompetEmploi.Detail.Count - 1 do
                begin
                        Competence := TobCompetEmploi.Detail[c].GetValue('POS_COMPETENCE');
                        TE := TobEtat.FindFirst(['PSA_SALARIE','PCO_COMPETENCE'],[Salarie,Competence],False);
                        If TE <> Nil then
                        begin
                                If TE.GetValue('MAITRISEEMP') < TobCompetEmploi.detail[c].GetValue('POS_DEGREMAITRISE') then
                                TE.PutValue('MAITRISEEMP',TobCompetEmploi.detail[c].GetValue('POS_DEGREMAITRISE'));
                        end
                        else
                        begin
                                TE := Tob.Create('UneCompetence',TobEtat,-1);
                                TE.AddChampSupValeur('PSA_SALARIE',Salarie);
                                TE.AddChampSupValeur('PSA_LIBELLEEMPLOI',Emploi);
                                TE.AddChampSupValeur('PCO_COMPETENCE',Competence);
                                TE.AddChampSupValeur('PSA_LIBELLE',nom);
                                TE.AddChampSupValeur('PSA_PRENOM',Prenom);
                                TE.AddChampSupValeur('PSA_ETABLISSEMENT',Etab);
                                TE.AddChampSupValeur('PCH_DATEVALIDATION',IDate1900);
                                TE.AddChampSupValeur('MAITRISESAL',0);
                                TE.AddChampSupValeur('MAITRISEEMP',TobCompetEmploi.detail[c].GetValue('POS_DEGREMAITRISE'));
                                TE.AddChampSupValeur('PCH_TABLELIBRECR1','');
                                TE.AddChampSupValeur('PCH_TABLELIBRECR2','');
                                TE.AddChampSupValeur('PCH_TABLELIBRECR3','');
                        end;
                end;
                MoveCurProgressForm('Salarie : ' + TobSalaries.Detail[i].GetValue('PSI_LIBELLE'));
                TobCompetEmploi.Free;
        end;
        FiniMoveProgressForm;
        TobSalaries.Free;
//debut PT2
        { LanceEtatTOB('E','PCO','PSE',TobEtat,True,False,False,Pages,'','',False);
          tobEtat.Free;}     //mise en commentaire
  TFQRS1(Ecran).LaTob:= TobEtat;
//fin PT2
end;
//debut PT2
procedure TOF_PGEDITCOMPSALEMPLOI.OnUpdate;
begin
  EditionComparatif(nil);
end;

procedure TOF_PGEDITCOMPSALEMPLOI.OnClose;
begin
  If TobEtat <> Nil Then FreeAndNil(TobEtat);
end;
//fin PT2
{TOF_PGEDITCOMPEMPLOI}

Procedure TOF_PGEDITCOMPEMPLOI.OnUpdate ;
var Where,WhereTmp,OrderBy : String;
    i : Integer;
    MultiCombo : THMultiValComboBox;
    Condition,StCondition : String;
begin

  Inherited ;
         OrderBy := '';
         // Gestion clause WHERE pour affichage des compétences dans bande détail
        Where := '';
        WhereTmp := '';
        If Arg = 'COMPFORM' then
        begin
                WhereTmp := WhereTmp + Conditionform('FORMATION');
                WhereTmp := WhereTmp + Conditionform('NATUREFORM');
                WhereTmp := WhereTmp + Conditionform('CENTREFORMGU');
                WhereTmp := WhereTmp + Conditionform('LIEUFORM');
                For i := 1 to 8 do
                begin
                        WhereTmp := WhereTmp + Conditionform('FORMATION'+IntToStr(i));
                end;
        end
        else
        begin
                MultiCombo  := THMultiValComboBox(GetControl('CODE'));
                Condition := MultiCombo.Value;
                While Condition <> '' do
                begin
                        StCondition := ReadTokenPipe(Condition,';');
                        If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR CC_CODE="'+StCondition+'"'
                        else WhereTmp := 'AND (CC_CODE="'+StCondition+'"';
                end;
                If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        end;
        Where := Where + WhereTmp;
        WhereTmp := '';
        For i := 1 to 3 do
        begin
                MultiCombo  := THMultiValComboBox(GetControl('LIBRECOMPFOR'+IntToStr(i)));
                Condition := MultiCombo.Value;
                While Condition <> '' do
                begin
                        StCondition := ReadTokenPipe(Condition,';');
                        If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR POS_TABLELIBRECR'+IntToStr(i)+'="'+StCondition+'"'
                        else WhereTmp := 'AND (POS_TABLELIBRECR'+IntToStr(i)+'="'+StCondition+'"';
                end;
                If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
                Where := Where + WhereTmp;
                WhereTmp := '';
        end;
        SetControlText('CONDITIONCOMP',Where);
         // Gestion clause WHERE pour affichage des salariés dans bande détail
        Where := '';
        MultiCombo  := THMultiValComboBox(GetControl('ETABLISSEMENT'));
        Condition := MultiCombo.Value;
        WhereTmp := '';
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_ETABLISSEMENT="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_ETABLISSEMENT="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        MultiCombo  := THMultiValComboBox(GetControl('CODESTAT'));
        Condition := MultiCombo.Value;
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_CODESTAT="'+StCondition+'"'
                else WhereTmp := 'AND (PSA_CODESTAT="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Where := Where + WhereTmp;
        WhereTmp := '';
        For i := 1 to VH_Paie.PGNbreStatOrg do
        begin
                MultiCombo  := THMultiValComboBox(GetControl('TRAVAILN'+IntToStr(i)));
                Condition := MultiCombo.Value;
                While Condition <> '' do
                begin
                        StCondition := ReadTokenPipe(Condition,';');
                        If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR PSA_TRAVAILN'+IntToStr(i)+'="'+StCondition+'"'
                        else WhereTmp := 'AND (PSA_TRAVAILN'+IntToStr(i)+'="'+StCondition+'"';
                end;
                If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
                Where := Where + WhereTmp;
                WhereTmp := '';
        end;
        SetControlText('CONDITIONSAL',Where);
end;

Function TOF_PGEDITCOMPEMPLOI.Conditionform(Champ : String) : String;
var MultiCombo  : THMultiValComboBox;
    ChampMcd,Condition,StCondition,WhereTmp : String;
begin
        Result := '';
        MultiCombo  := THMultiValComboBox(GetControl(Champ));
        If MultiCombo = Nil then Exit;
        Condition := MultiCombo.Value;
        ChampMcd := 'PST_' + Champ;
        WhereTmp := '';
        While Condition <> '' do
        begin
                StCondition := ReadTokenPipe(Condition,';');
                If WhereTmp <> '' then WhereTmp := WhereTmp + ' OR '+ChampMcd+'="'+StCondition+'"'
                else WhereTmp := 'AND ('+ChampMcd+'="'+StCondition+'"';
        end;
        If WhereTmp <> '' then WhereTmp := WhereTmp + ')';
        Result := WhereTmp;
end;

procedure TOF_PGEDITCOMPEMPLOI.OnArgument (S : String ) ;
var Num : Integer;
begin
  Inherited ;
  Arg := ReadTokenPipe(S,';');
  SetControlText('DOSSIER', GetParamSoc('SO_LIBELLE'));
  For Num  :=  1 to 5 do
  begin
        VisibiliteChampCompetence (IntToStr(Num),GetControl ('PCO_LIBRECOMP'+IntToStr(Num)),GetControl ('TLIBRECOMP'+IntToStr(Num)));
  end;
  For Num  :=  1 to 3 do
  begin
        VisibiliteChampCompetenceRess (IntToStr(Num),GetControl ('LIBRECOMPFOR'+IntToStr(Num)),GetControl ('TLIBRECOMPFOR'+IntToStr(Num)));
  end;
  For Num := 1 to VH_Paie.NBFormationLibre do
  begin
        if Num > 8 then Break;
        VisibiliteChampFormation (IntToStr(Num),GetControl ('FORMATION'+IntToStr(Num)),GetControl ('TPST_FORMATION'+IntToStr(Num)));
  end;
  For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
        if Num >4 then Break;
        VisibiliteChampSalarie (IntToStr(Num),GetControl ('TRAVAILN'+IntToStr(Num)),GetControl ('TTRAVAILN'+IntToStr(Num)));
    end;
    VisibiliteStat(GetControl ('CODESTAT'),GetControl ('TCODESTAT'));
    // DEBUT PT1
    If ARg = 'COMPEMPLOI' then TFQRS1(Ecran).Caption := 'Edition des emplois par compétence'
    else If ARg = 'COMPFORM' then TFQRS1(Ecran).Caption := 'Edition des formations par compétence';
    UpdateCaption(Ecran);
    //FIN PT1
end ;


Initialization
  registerclasses ( [ TOF_PGEDITEMPLOICOMP,TOF_PGEDITSTAGECOMPETENCE,TOF_PGEDITCOMPETENCESINTER,TOF_PGEDITINTERCOMPETENCES,TOF_PGEDITCOMPSALEMPLOI,TOF_PGEDITCOMPEMPLOI] ) ;
end.

